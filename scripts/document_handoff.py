#!/usr/bin/env python3
"""Create or verify a typed, hash-bound WORKING handoff for DOCX finalization."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from docx_risk_scan import scan as live_risk_scan
from path_guard import ensure_new_file, same_file


HANDOFF_SCHEMA_VERSION = "1.2"
TASK_BRIEF_SCHEMA_VERSION = "1.0"
ALLOW_CREATE = False


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def fingerprint(path: Path | None) -> dict | None:
    if path is None:
        return None
    resolved = path.resolve(strict=True)
    stat = resolved.stat()
    return {
        "path": str(resolved),
        "sha256": sha256(resolved),
        "size_bytes": stat.st_size,
        "modified_at_utc": datetime.fromtimestamp(stat.st_mtime, timezone.utc).isoformat(),
    }


def ensure_distinct(named_paths: list[tuple[str, Path | None]]) -> None:
    concrete = [(name, path) for name, path in named_paths if path is not None]
    for index, (left_name, left_path) in enumerate(concrete):
        for right_name, right_path in concrete[index + 1 :]:
            if same_file(left_path, right_path):
                raise ValueError(f"{left_name} and {right_name} must be separate files")


def write_json(path: Path, payload: dict) -> None:
    resolved = ensure_new_file(path, suffixes=[".json"])
    with resolved.open("x", encoding="utf-8", newline="\n") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)
        handle.write("\n")


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def emit(payload: dict, *, error: bool = False) -> None:
    print(json.dumps(payload, ensure_ascii=True, indent=2), file=sys.stderr if error else sys.stdout)


def check_docx(path: Path) -> list[str]:
    issues: list[str] = []
    if path.suffix.lower() not in {".docx", ".docm", ".dotx", ".dotm"}:
        return [f"unsupported extension: {path.suffix}"]
    try:
        with zipfile.ZipFile(path) as package:
            names = set(package.namelist())
            bad_member = package.testzip()
            if bad_member:
                issues.append(f"CRC failure: {bad_member}")
            for required in ("[Content_Types].xml", "_rels/.rels", "word/document.xml"):
                if required not in names:
                    issues.append(f"missing required part: {required}")
    except (OSError, zipfile.BadZipFile) as exc:
        issues.append(f"invalid Word package: {exc}")
    return issues


def parse_pages(values: list[str]) -> list[int]:
    pages: set[int] = set()
    for value in values:
        for token in value.split(","):
            token = token.strip()
            if not token:
                continue
            if "-" in token:
                start_text, end_text = token.split("-", 1)
                start, end = int(start_text), int(end_text)
                if start < 1 or end < start:
                    raise ValueError(f"invalid page range: {token}")
                pages.update(range(start, end + 1))
            else:
                page = int(token)
                if page < 1:
                    raise ValueError(f"invalid page: {token}")
                pages.add(page)
    return sorted(pages)


def compare_fingerprint(stored: dict, path: Path, label: str) -> list[str]:
    issues: list[str] = []
    current = fingerprint(path)
    if current["sha256"] != stored.get("sha256"):
        issues.append(f"{label} SHA-256 mismatch")
    if current["size_bytes"] != stored.get("size_bytes"):
        issues.append(f"{label} size mismatch")
    try:
        if not same_file(Path(stored["path"]), path):
            issues.append(f"{label} path does not match the bound file")
    except (KeyError, TypeError, OSError) as exc:
        issues.append(f"{label} stored path is invalid: {exc}")
    return issues


def command_task_brief(args: argparse.Namespace) -> int:
    if not ALLOW_CREATE:
        raise ValueError("this skill may verify task briefs but may not create them")
    output = args.output.resolve()
    inputs = [path.resolve(strict=True) for path in args.input_file]
    if not args.authorized_mode.strip() or not args.scope_summary.strip():
        raise ValueError("authorized mode and scope summary must be non-empty")
    requirements = [item.strip() for item in args.requirement if item.strip()]
    if not requirements:
        raise ValueError("at least one non-empty --requirement is required")
    ensure_distinct([("task-brief output", output), *[(f"input file {index}", path) for index, path in enumerate(inputs, 1)]])
    payload = {
        "schema_version": TASK_BRIEF_SCHEMA_VERSION,
        "record_type": "documents_task_brief",
        "producer_skill": "documents-fast",
        "created_at_utc": utc_now(),
        "authorized_mode": args.authorized_mode.strip(),
        "scope_summary": args.scope_summary.strip(),
        "requirements": requirements,
        "input_files": [fingerprint(path) for path in inputs],
    }
    write_json(output, payload)
    emit({"status": "CREATED", "record_type": payload["record_type"], "output": str(output)})
    return 0


def validate_task_brief(brief_path: Path, expected_authorized_mode: str | None = None) -> tuple[dict, list[str]]:
    brief = brief_path.resolve(strict=True)
    payload = load_json(brief)
    issues: list[str] = []
    if payload.get("schema_version") != TASK_BRIEF_SCHEMA_VERSION:
        issues.append("task brief has unsupported schema_version")
    if payload.get("record_type") != "documents_task_brief":
        issues.append("task brief has unexpected record_type")
    if payload.get("producer_skill") != "documents-fast":
        issues.append("task brief has unexpected producer_skill")
    if not str(payload.get("authorized_mode", "")).strip():
        issues.append("task brief authorized_mode is empty")
    if expected_authorized_mode is not None and payload.get("authorized_mode") != expected_authorized_mode:
        issues.append("task brief authorized_mode does not match the handoff")
    if not str(payload.get("scope_summary", "")).strip():
        issues.append("task brief scope_summary is empty")
    requirements = payload.get("requirements")
    if not isinstance(requirements, list) or not requirements or any(not str(item).strip() for item in requirements):
        issues.append("task brief requirements must contain non-empty entries")
    input_entries = payload.get("input_files")
    if not isinstance(input_entries, list):
        issues.append("task brief input_files must be a list")
        input_entries = []
    for index, entry in enumerate(input_entries, 1):
        try:
            input_path = Path(entry["path"]).resolve(strict=True)
            issues.extend(compare_fingerprint(entry, input_path, f"task brief input {index}"))
        except (OSError, KeyError, TypeError, ValueError) as exc:
            issues.append(f"task brief input {index} is invalid: {exc}")
    return {
        "record": fingerprint(brief),
        "authorized_mode": payload.get("authorized_mode"),
        "scope_summary": payload.get("scope_summary"),
        "requirements": requirements if isinstance(requirements, list) else [],
        "input_files": input_entries,
    }, issues


def validate_risk_report(report_path: Path, document_path: Path) -> tuple[dict, list[str]]:
    issues: list[str] = []
    report_resolved = report_path.resolve(strict=True)
    document_resolved = document_path.resolve(strict=True)
    if same_file(report_resolved, document_resolved):
        issues.append("risk report and document must be separate files")
    report_payload = load_json(report_resolved)
    document_fp = fingerprint(document_resolved)
    live_payload = live_risk_scan(document_resolved)
    if report_payload.get("schema_version") != "1.0":
        issues.append("risk report has unsupported schema_version")
    if report_payload.get("record_type") != "documents_docx_risk_report":
        issues.append("risk report has unexpected record_type")
    if report_payload.get("scanner") != "documents-split/docx_risk_scan.py":
        issues.append("risk report has unexpected scanner identity")
    if report_payload.get("package", {}).get("valid") is not True:
        issues.append("risk report does not record a valid package")
    try:
        if not same_file(Path(report_payload["path"]), document_resolved):
            issues.append("risk report path does not match document")
    except (KeyError, TypeError, OSError) as exc:
        issues.append(f"risk report document path is invalid: {exc}")
    for key in (
        "sha256", "size_bytes", "package", "counts", "feature_flags",
        "field_instruction_fragments", "field_markers_detected",
    ):
        if report_payload.get(key) != live_payload.get(key):
            issues.append(f"risk report does not match live scan for: {key}")
    entry = {
        "report": fingerprint(report_resolved),
        "document": document_fp,
        "package_valid": report_payload.get("package", {}).get("valid"),
        "feature_flags": report_payload.get("feature_flags", {}),
        "counts": report_payload.get("counts", {}),
        "field_markers_detected": report_payload.get("field_markers_detected", []),
    }
    return entry, issues


def verify_stored_risk(entry: dict | None, document_path: Path, label: str) -> list[str]:
    if not entry or not entry.get("report", {}).get("path"):
        return [f"{label} risk-report binding is missing"]
    issues: list[str] = []
    try:
        report_path = Path(entry["report"]["path"]).resolve(strict=True)
        issues.extend(compare_fingerprint(entry["report"], report_path, f"{label} risk-report file"))
        current_entry, binding_issues = validate_risk_report(report_path, document_path)
        issues.extend(f"{label} {item}" for item in binding_issues)
        if current_entry["document"]["sha256"] != entry.get("document", {}).get("sha256"):
            issues.append(f"{label} stored risk document SHA-256 mismatch")
    except (OSError, json.JSONDecodeError, KeyError, TypeError, ValueError) as exc:
        issues.append(f"{label} risk report is invalid: {exc}")
    return issues


def command_create(args: argparse.Namespace) -> int:
    if not ALLOW_CREATE:
        raise ValueError("this skill may verify handoffs but may not create them")
    candidate = args.candidate.resolve(strict=True)
    source = args.source.resolve(strict=True) if args.source else None
    candidate_risk = args.candidate_risk_report.resolve(strict=True)
    source_risk = args.source_risk_report.resolve(strict=True) if args.source_risk_report else None
    task_brief = args.task_brief.resolve(strict=True) if args.task_brief else None
    output = args.output.resolve()
    if source is not None and source_risk is None:
        raise ValueError("--source-risk-report is required when --source is provided")
    if source is None and source_risk is not None:
        raise ValueError("--source-risk-report requires --source")
    if source is None and task_brief is None:
        raise ValueError("a new-document handoff requires --task-brief")
    ensure_distinct([
        ("source", source), ("candidate", candidate), ("source risk report", source_risk),
        ("candidate risk report", candidate_risk), ("task brief", task_brief), ("handoff output", output),
    ])
    package_issues = check_docx(candidate)
    if source is not None:
        package_issues.extend(f"source: {item}" for item in check_docx(source))
    if package_issues:
        raise ValueError("package checks failed: " + "; ".join(package_issues))
    candidate_risk_entry, risk_issues = validate_risk_report(candidate_risk, candidate)
    source_risk_entry = None
    if source is not None and source_risk is not None:
        source_risk_entry, source_risk_issues = validate_risk_report(source_risk, source)
        risk_issues.extend(f"source: {item}" for item in source_risk_issues)
    task_brief_entry = None
    if task_brief is not None:
        task_brief_entry, brief_issues = validate_task_brief(task_brief, args.authorized_mode)
        risk_issues.extend(f"task brief: {item}" for item in brief_issues)
        input_paths = [Path(entry["path"]).resolve(strict=True) for entry in task_brief_entry["input_files"]]
        ensure_distinct([
            ("source", source), ("candidate", candidate), ("source risk report", source_risk),
            ("candidate risk report", candidate_risk), ("task brief", task_brief), ("handoff output", output),
            *[(f"task brief input {index}", path) for index, path in enumerate(input_paths, 1)],
        ])
    if risk_issues:
        raise ValueError("evidence binding failed: " + "; ".join(risk_issues))
    payload = {
        "schema_version": HANDOFF_SCHEMA_VERSION,
        "record_type": "documents_working_handoff",
        "producer_skill": "documents-fast",
        "created_at_utc": utc_now(),
        "artifact_status": "WORKING",
        "release_eligible": False,
        "do_not_claim_final": True,
        "request": {"authorized_mode": args.authorized_mode, "scope_summary": args.summary, "task_brief": task_brief_entry},
        "source": fingerprint(source),
        "candidate": fingerprint(candidate),
        "edit": {"method": args.method},
        "risk": {"source": source_risk_entry, "candidate": candidate_risk_entry},
        "qa": {
            "checks_completed": args.checked,
            "visually_inspected_pages": parse_pages(args.inspected_pages),
            "checks_pending": args.pending,
            "limitations": args.note,
        },
        "unresolved_issues": args.issue,
    }
    write_json(output, payload)
    emit({"status": "CREATED", "output": str(output), "candidate_sha256": payload["candidate"]["sha256"]})
    return 0


def verify_payload(handoff_path: Path, candidate: Path | None, source: Path | None) -> tuple[dict, list[str]]:
    handoff = handoff_path.resolve(strict=True)
    payload = load_json(handoff)
    issues: list[str] = []
    if payload.get("schema_version") != HANDOFF_SCHEMA_VERSION:
        issues.append("unsupported schema_version")
    if payload.get("record_type") != "documents_working_handoff":
        issues.append("unexpected record_type")
    if payload.get("producer_skill") != "documents-fast":
        issues.append("unexpected producer_skill")
    if payload.get("artifact_status") != "WORKING":
        issues.append("handoff is not WORKING")
    if payload.get("release_eligible") is not False or payload.get("do_not_claim_final") is not True:
        issues.append("working-result safety flags are invalid")
    expected_candidate = payload.get("candidate") or {}
    candidate_path = candidate or (Path(expected_candidate["path"]) if expected_candidate.get("path") else None)
    if candidate_path is None:
        issues.append("candidate path is missing")
    else:
        try:
            candidate_path = candidate_path.resolve(strict=True)
            issues.extend(compare_fingerprint(expected_candidate, candidate_path, "candidate"))
            issues.extend(check_docx(candidate_path))
            issues.extend(verify_stored_risk(payload.get("risk", {}).get("candidate"), candidate_path, "candidate"))
        except OSError as exc:
            issues.append(f"candidate unavailable: {exc}")
    expected_source = payload.get("source")
    source_path = source or (Path(expected_source["path"]) if expected_source and expected_source.get("path") else None)
    if expected_source:
        if source_path is None:
            issues.append("source path is missing")
        else:
            try:
                source_path = source_path.resolve(strict=True)
                issues.extend(compare_fingerprint(expected_source, source_path, "source"))
                issues.extend(check_docx(source_path))
                issues.extend(verify_stored_risk(payload.get("risk", {}).get("source"), source_path, "source"))
            except OSError as exc:
                issues.append(f"source unavailable: {exc}")
    elif source is not None:
        issues.append("handoff represents a new document but a source override was supplied")
    task_brief_entry = payload.get("request", {}).get("task_brief")
    if task_brief_entry:
        try:
            brief_path = Path(task_brief_entry["record"]["path"]).resolve(strict=True)
            issues.extend(compare_fingerprint(task_brief_entry["record"], brief_path, "task brief"))
            live_entry, brief_issues = validate_task_brief(brief_path, payload.get("request", {}).get("authorized_mode"))
            issues.extend(brief_issues)
            if live_entry.get("requirements") != task_brief_entry.get("requirements"):
                issues.append("task brief requirements differ from the handoff binding")
        except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
            issues.append(f"task brief is invalid: {exc}")
    elif expected_source is None:
        issues.append("new-document handoff has no typed task brief")
    try:
        ensure_distinct([
            ("handoff", handoff), ("candidate", candidate_path), ("source", source_path),
            ("candidate risk report", Path(payload["risk"]["candidate"]["report"]["path"])),
            ("source risk report", Path(payload["risk"]["source"]["report"]["path"]) if expected_source else None),
            ("task brief", Path(task_brief_entry["record"]["path"]) if task_brief_entry else None),
        ])
    except (KeyError, TypeError, ValueError) as exc:
        issues.append(str(exc))
    return payload, issues


def command_verify(args: argparse.Namespace) -> int:
    try:
        _, issues = verify_payload(args.handoff, args.candidate, args.source)
    except (OSError, json.JSONDecodeError, KeyError, TypeError, ValueError) as exc:
        issues = [f"invalid handoff: {exc}"]
    status = "VALID" if not issues else "INVALID"
    emit({"status": status, "issues": issues}, error=bool(issues))
    return 0 if not issues else 2


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    if ALLOW_CREATE:
        brief = subparsers.add_parser("task-brief", help="Create a typed task brief for net-new document provenance")
        brief.add_argument("--output", type=Path, required=True)
        brief.add_argument("--authorized-mode", required=True)
        brief.add_argument("--scope-summary", required=True)
        brief.add_argument("--requirement", action="append", required=True)
        brief.add_argument("--input-file", type=Path, action="append", default=[])
        brief.set_defaults(handler=command_task_brief)
        create = subparsers.add_parser("create", help="Create a hash-bound WORKING handoff")
        create.add_argument("--source", type=Path, help="Original source; omit for a new document")
        create.add_argument("--source-risk-report", type=Path, help="Required with --source")
        create.add_argument("--candidate", type=Path, required=True)
        create.add_argument("--candidate-risk-report", type=Path, required=True)
        create.add_argument("--task-brief", type=Path, help="Typed task brief; required for a new document")
        create.add_argument("--output", type=Path, required=True)
        create.add_argument("--authorized-mode", required=True)
        create.add_argument("--summary", required=True)
        create.add_argument("--method", required=True)
        create.add_argument("--checked", action="append", required=True)
        create.add_argument("--inspected-pages", action="append", default=[])
        create.add_argument("--pending", action="append", required=True)
        create.add_argument("--issue", action="append", default=[])
        create.add_argument("--note", action="append", default=[])
        create.set_defaults(handler=command_create)
    verify = subparsers.add_parser("verify", help="Verify a WORKING handoff, hashes, packages, risks, and task brief")
    verify.add_argument("handoff", type=Path)
    verify.add_argument("--source", type=Path)
    verify.add_argument("--candidate", type=Path)
    verify.set_defaults(handler=command_verify)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        return args.handler(args)
    except (OSError, ValueError, json.JSONDecodeError, KeyError, TypeError) as exc:
        emit({"status": "ERROR", "error": str(exc)}, error=True)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
