#!/usr/bin/env python3
"""Create and live-validate an isolated DOCX finalization preflight bundle."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from document_handoff import (
    check_docx,
    compare_fingerprint,
    ensure_distinct,
    fingerprint,
    load_json,
    sha256,
    utc_now,
    validate_risk_report,
    write_json,
)
from docx_compare import build_report, validate_report
from docx_risk_scan import scan as risk_scan
from path_guard import ensure_new_directory


SCHEMA_VERSION = "1.0"
RECORD_TYPE = "documents_preflight_bundle"
PRODUCER = "documents-finalize/document_preflight.py"


def _write_payload(path: Path, payload: dict) -> None:
    write_json(path, payload)


def create_bundle(
    source: Path | None,
    candidate: Path,
    final: Path,
    output_dir: Path,
    *,
    new_document: bool,
) -> Path:
    if new_document == (source is not None):
        raise ValueError("choose exactly one provenance mode: --source or --new-document")
    source_path = source.resolve(strict=True) if source else None
    candidate_path = candidate.resolve(strict=True)
    final_path = final.resolve(strict=True)
    ensure_distinct(
        [("source", source_path), ("candidate", candidate_path), ("final", final_path)]
    )
    for label, path in (
        ("source", source_path),
        ("candidate", candidate_path),
        ("final", final_path),
    ):
        if path is None:
            continue
        issues = check_docx(path)
        if issues:
            raise ValueError(f"{label} package failed: {'; '.join(issues)}")
    if sha256(candidate_path) != sha256(final_path):
        raise ValueError(
            "release-only final must be an exact byte-for-byte copy of the working candidate"
        )
    inputs = [candidate_path, final_path]
    if source_path:
        inputs.append(source_path)
    bundle_dir = ensure_new_directory(output_dir, inputs=inputs)
    bundle_dir.mkdir(parents=False, exist_ok=False)

    risk_paths: dict[str, Path | None] = {
        "source": bundle_dir / "source-risk.json" if source_path else None,
        "candidate": bundle_dir / "candidate-risk.json",
        "final": bundle_dir / "final-risk.json",
    }
    for label, document in (
        ("source", source_path),
        ("candidate", candidate_path),
        ("final", final_path),
    ):
        if document is not None:
            _write_payload(risk_paths[label], risk_scan(document))

    comparison_path = bundle_dir / "comparison.json"
    _write_payload(
        comparison_path, build_report(source_path, candidate_path, final_path)
    )

    risk_entries: dict[str, dict | None] = {"source": None}
    for label, document in (
        ("source", source_path),
        ("candidate", candidate_path),
        ("final", final_path),
    ):
        if document is None:
            continue
        entry, issues = validate_risk_report(risk_paths[label], document)
        if issues:
            raise ValueError(f"{label} risk report failed live validation: {'; '.join(issues)}")
        risk_entries[label] = entry
    _, comparison_issues = validate_report(
        comparison_path, source_path, candidate_path, final_path
    )
    if comparison_issues:
        raise ValueError(
            "comparison report failed live validation: " + "; ".join(comparison_issues)
        )

    record_path = bundle_dir / "preflight.json"
    payload = {
        "schema_version": SCHEMA_VERSION,
        "record_type": RECORD_TYPE,
        "producer": PRODUCER,
        "created_at_utc": utc_now(),
        "status": "PASS",
        "provenance_mode": "new-document" if new_document else "source-edit",
        "source": fingerprint(source_path),
        "candidate": fingerprint(candidate_path),
        "final": fingerprint(final_path),
        "risk": risk_entries,
        "comparison": fingerprint(comparison_path),
        "bundle_directory": str(bundle_dir),
    }
    _write_payload(record_path, payload)
    return record_path


def validate_preflight(
    record_path: Path,
    source: Path | None = None,
    candidate: Path | None = None,
    final: Path | None = None,
) -> tuple[dict, dict, list[str]]:
    record = record_path.resolve(strict=True)
    payload = load_json(record)
    issues: list[str] = []
    if payload.get("schema_version") != SCHEMA_VERSION:
        issues.append("preflight has unsupported schema_version")
    if payload.get("record_type") != RECORD_TYPE:
        issues.append("preflight has unexpected record_type")
    if payload.get("producer") != PRODUCER:
        issues.append("preflight has unexpected producer")
    if payload.get("status") != "PASS":
        issues.append("preflight status is not PASS")

    source_entry = payload.get("source")
    try:
        source_path = (
            source.resolve(strict=True)
            if source
            else Path(source_entry["path"]).resolve(strict=True)
            if source_entry
            else None
        )
        candidate_path = (
            candidate.resolve(strict=True)
            if candidate
            else Path(payload["candidate"]["path"]).resolve(strict=True)
        )
        final_path = (
            final.resolve(strict=True)
            if final
            else Path(payload["final"]["path"]).resolve(strict=True)
        )
    except (OSError, KeyError, TypeError, ValueError) as exc:
        return payload, {}, [*issues, f"preflight document binding is invalid: {exc}"]

    if payload.get("provenance_mode") == "new-document" and source_path is not None:
        issues.append("new-document preflight unexpectedly binds a source")
    if payload.get("provenance_mode") == "source-edit" and source_path is None:
        issues.append("source-edit preflight lacks a source")
    for label, entry, path in (
        ("source", source_entry, source_path),
        ("candidate", payload.get("candidate"), candidate_path),
        ("final", payload.get("final"), final_path),
    ):
        if path is None:
            if entry is not None:
                issues.append(f"preflight unexpectedly stores {label}")
            continue
        issues.extend(compare_fingerprint(entry or {}, path, f"preflight {label}"))
        issues.extend(f"{label}: {item}" for item in check_docx(path))
    if sha256(candidate_path) != sha256(final_path):
        issues.append(
            "release-only final is not an exact byte-for-byte copy of the working candidate"
        )
    try:
        ensure_distinct(
            [
                ("preflight", record),
                ("source", source_path),
                ("candidate", candidate_path),
                ("final", final_path),
            ]
        )
    except ValueError as exc:
        issues.append(str(exc))

    risk_paths: dict[str, Path | None] = {"source": None}
    for label, document in (
        ("source", source_path),
        ("candidate", candidate_path),
        ("final", final_path),
    ):
        entry = (payload.get("risk") or {}).get(label)
        if document is None:
            if entry is not None:
                issues.append("preflight unexpectedly stores source risk")
            continue
        try:
            report_path = Path(entry["report"]["path"]).resolve(strict=True)
            risk_paths[label] = report_path
            issues.extend(
                compare_fingerprint(
                    entry["report"], report_path, f"preflight {label} risk report"
                )
            )
            _, found = validate_risk_report(report_path, document)
            issues.extend(f"{label} risk: {item}" for item in found)
        except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
            issues.append(f"{label} risk binding is invalid: {exc}")

    comparison_payload: dict = {}
    comparison_path: Path | None = None
    try:
        comparison_path = Path(payload["comparison"]["path"]).resolve(strict=True)
        issues.extend(
            compare_fingerprint(
                payload["comparison"], comparison_path, "preflight comparison report"
            )
        )
        comparison_payload, found = validate_report(
            comparison_path, source_path, candidate_path, final_path
        )
        issues.extend(f"comparison: {item}" for item in found)
    except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
        issues.append(f"comparison binding is invalid: {exc}")

    try:
        bundle_dir = Path(payload["bundle_directory"]).resolve(strict=True)
        if not bundle_dir.is_dir() or record.parent != bundle_dir:
            issues.append("preflight bundle directory is invalid")
        expected = {record}
        expected.update(path for path in risk_paths.values() if path is not None)
        if comparison_path is not None:
            expected.add(comparison_path)
        actual = {path.resolve() for path in bundle_dir.rglob("*") if path.is_file()}
        if actual != expected:
            issues.append("preflight bundle contains unbound or missing files")
    except (OSError, KeyError, TypeError, ValueError) as exc:
        issues.append(f"preflight bundle inventory is invalid: {exc}")

    context = {
        "record": fingerprint(record),
        "source": source_path,
        "candidate": candidate_path,
        "final": final_path,
        "risk_paths": risk_paths,
        "comparison_path": comparison_path,
        "comparison": comparison_payload,
    }
    return payload, context, issues


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    provenance = parser.add_mutually_exclusive_group(required=True)
    provenance.add_argument("--source", type=Path)
    provenance.add_argument("--new-document", action="store_true")
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--final", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    try:
        record = create_bundle(
            args.source,
            args.candidate,
            args.final,
            args.output_dir,
            new_document=args.new_document,
        )
        print(
            json.dumps(
                {"status": "PASS", "record_type": RECORD_TYPE, "preflight": str(record)},
                ensure_ascii=True,
            )
        )
        return 0
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        print(json.dumps({"status": "ERROR", "error": str(exc)}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
