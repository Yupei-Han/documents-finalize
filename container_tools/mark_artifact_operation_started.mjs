#!/usr/bin/env node

import { createHash, randomUUID } from "node:crypto";
import { createReadStream, existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";
import path from "node:path";
import process from "node:process";

const SCHEMA_VERSION = "1.0";
const SCRIPT_PATH = fileURLToPath(import.meta.url);

function fail(message) {
  console.error(JSON.stringify({ status: "ERROR", error: message }));
  process.exit(2);
}

function absolute(raw, label) {
  if (!raw || !path.isAbsolute(raw)) fail(`${label} must be an absolute path`);
  return path.resolve(raw);
}

function pathKey(value) {
  const resolved = path.resolve(value);
  return process.platform === "win32" ? resolved.toLowerCase() : resolved;
}

function assertDistinct(paths, label) {
  const keys = paths.map(pathKey);
  if (new Set(keys).size !== keys.length) fail(`${label} paths must be distinct`);
}

async function sha256(file) {
  const digest = createHash("sha256");
  await new Promise((resolve, reject) => {
    const stream = createReadStream(file);
    stream.on("data", (chunk) => digest.update(chunk));
    stream.on("end", resolve);
    stream.on("error", reject);
  });
  return digest.digest("hex");
}

async function fingerprint(file, label) {
  const resolved = absolute(file, label);
  if (!existsSync(resolved)) fail(`${label} does not exist: ${resolved}`);
  const item = statSync(resolved);
  if (!item.isFile()) fail(`${label} is not a file: ${resolved}`);
  return {
    path: resolved,
    sha256: await sha256(resolved),
    size_bytes: item.size,
    modified_at_utc: item.mtime.toISOString(),
  };
}

function writeJsonNew(target, payload) {
  const resolved = absolute(target, "JSON output");
  if (path.extname(resolved).toLowerCase() !== ".json") fail(`JSON output must end in .json: ${resolved}`);
  if (existsSync(resolved)) fail(`JSON output already exists: ${resolved}`);
  mkdirSync(path.dirname(resolved), { recursive: true });
  try {
    writeFileSync(resolved, `${JSON.stringify(payload, null, 2)}\n`, { encoding: "utf8", flag: "wx" });
  } catch (error) {
    fail(`unable to create JSON output ${resolved}: ${error.message}`);
  }
  return resolved;
}

function parse(command) {
  try {
    return parseArgs({
      args: process.argv.slice(3),
      strict: true,
      allowPositionals: false,
      options: command === "create"
        ? {
            "operation-kind": { type: "string" },
            "expected-output-count": { type: "string" },
            "output-format": { type: "string" },
            receipt: { type: "string" },
            input: { type: "string", multiple: true, default: [] },
            "planned-output": { type: "string", multiple: true, default: [] },
          }
        : {
            receipt: { type: "string" },
            verification: { type: "string" },
            "require-outputs": { type: "boolean", default: false },
          },
    }).values;
  } catch (error) {
    fail(error.message);
  }
}

async function createReceipt() {
  const values = parse("create");
  const kind = values["operation-kind"];
  if (!new Set(["create", "edit"]).has(kind)) fail("--operation-kind must be create or edit");
  const expected = Number(values["expected-output-count"]);
  if (!Number.isSafeInteger(expected) || expected < 1 || expected > 100) {
    fail("--expected-output-count must be an integer from 1 to 100");
  }
  const format = String(values["output-format"] || "").toLowerCase();
  if (format !== "docx") fail("--output-format currently supports only docx");
  const receipt = absolute(values.receipt, "--receipt");
  const inputs = values.input.map((item, index) => absolute(item, `--input ${index + 1}`));
  const outputs = values["planned-output"].map((item, index) => absolute(item, `--planned-output ${index + 1}`));
  if (kind === "edit" && inputs.length === 0) fail("edit operations require at least one --input");
  if (outputs.length !== expected) fail("--planned-output count must equal --expected-output-count");
  if (path.extname(receipt).toLowerCase() !== ".json") fail("--receipt must end in .json");
  if (existsSync(receipt)) fail(`receipt already exists: ${receipt}`);
  for (const output of outputs) {
    if (path.extname(output).toLowerCase() !== `.${format}`) fail(`planned output has wrong extension: ${output}`);
    if (existsSync(output)) fail(`planned output already exists: ${output}`);
  }
  assertDistinct([receipt, ...inputs, ...outputs], "receipt, input, and planned output");

  const payload = {
    schema_version: SCHEMA_VERSION,
    record_type: "documents_artifact_operation_start",
    status: "STARTED",
    operation_id: randomUUID(),
    created_at_utc: new Date().toISOString(),
    operation_kind: kind,
    expected_output_count: expected,
    output_format: format,
    inputs: await Promise.all(inputs.map((item, index) => fingerprint(item, `input ${index + 1}`))),
    planned_outputs: outputs.map((item) => ({ path: item, must_be_new: true })),
    tool: {
      script: await fingerprint(SCRIPT_PATH, "marker script"),
      node_version: process.version,
    },
  };
  const written = writeJsonNew(receipt, payload);
  console.log(JSON.stringify({ status: "STARTED", receipt: written, operation_id: payload.operation_id }, null, 2));
}

async function verifyReceipt() {
  const values = parse("verify");
  const receipt = absolute(values.receipt, "--receipt");
  if (!existsSync(receipt)) fail(`receipt does not exist: ${receipt}`);
  let payload;
  try {
    payload = JSON.parse(readFileSync(receipt, "utf8"));
  } catch (error) {
    fail(`receipt is not valid JSON: ${error.message}`);
  }
  const issues = [];
  if (payload.schema_version !== SCHEMA_VERSION) issues.push("unsupported schema_version");
  if (payload.record_type !== "documents_artifact_operation_start" || payload.status !== "STARTED") {
    issues.push("unexpected receipt record_type or status");
  }
  const currentScript = await fingerprint(SCRIPT_PATH, "marker script");
  if (payload.tool?.script?.sha256 !== currentScript.sha256) issues.push("marker script hash changed");
  const liveInputs = [];
  for (const [index, recorded] of (payload.inputs || []).entries()) {
    const live = await fingerprint(recorded.path, `recorded input ${index + 1}`);
    liveInputs.push(live);
    if (live.sha256 !== recorded.sha256 || live.size_bytes !== recorded.size_bytes) {
      issues.push(`recorded input ${index + 1} changed after operation start`);
    }
  }
  const liveOutputs = [];
  for (const [index, planned] of (payload.planned_outputs || []).entries()) {
    if (!existsSync(planned.path)) {
      if (values["require-outputs"]) issues.push(`planned output ${index + 1} does not exist`);
      continue;
    }
    liveOutputs.push(await fingerprint(planned.path, `planned output ${index + 1}`));
  }
  if ((payload.planned_outputs || []).length !== payload.expected_output_count) {
    issues.push("recorded planned-output count is inconsistent");
  }
  if (values["require-outputs"] && liveOutputs.length !== payload.expected_output_count) {
    issues.push("live output count does not match expected_output_count");
  }
  if (issues.length) fail(issues.join("; "));

  let verification = null;
  if (values["require-outputs"]) {
    if (!values.verification) fail("--verification is required with --require-outputs");
    const verificationPath = absolute(values.verification, "--verification");
    assertDistinct([receipt, verificationPath, ...liveInputs.map((item) => item.path), ...liveOutputs.map((item) => item.path)], "verification evidence");
    const verificationPayload = {
      schema_version: SCHEMA_VERSION,
      record_type: "documents_artifact_operation_completion",
      status: "VERIFIED",
      verified_at_utc: new Date().toISOString(),
      operation_id: payload.operation_id,
      start_receipt: await fingerprint(receipt, "start receipt"),
      inputs: liveInputs,
      outputs: liveOutputs,
      verifier: {
        script: currentScript,
        node_version: process.version,
      },
    };
    verification = writeJsonNew(verificationPath, verificationPayload);
  }
  console.log(JSON.stringify({ status: "VERIFIED", receipt, verification, outputs: liveOutputs.length }, null, 2));
}

const command = process.argv[2];
if (command === "create") {
  await createReceipt();
} else if (command === "verify") {
  await verifyReceipt();
} else {
  fail("usage: mark_artifact_operation_started.mjs <create|verify> [options]");
}
