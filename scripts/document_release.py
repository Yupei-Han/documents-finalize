#!/usr/bin/env python3
"""Build live-verifiable DOCX QA evidence and issue a fail-closed release record."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

from PIL import Image

from document_handoff import (
    compare_fingerprint,
    emit,
    ensure_distinct,
    fingerprint,
    load_json,
    same_file,
    utc_now,
    verify_payload,
    write_json,
)
from document_preflight import validate_preflight
from path_guard import ensure_new_file


SCHEMA_VERSION = "2.0"
OPERATION_SCHEMA_VERSION = "1.0"
DELIVERY_SCHEMA_VERSION = "1.0"
REQUIRED_GATES = (
    "identity_provenance",
    "package_integrity",
    "authorized_scope_content",
    "protected_structures",
    "native_features_policy",
    "application_behavior",
    "visual_qa",
    "output_hygiene",
)
RENDERER_PRODUCERS = {
    "render_docx_word.ps1": "microsoft-word",
    "render_docx_libreoffice.py": "libreoffice",
}
REPAIR_LANGUAGE = re.compile(r"\b(repair|repaired|recover|recovered|corrupt|corrupted)\b", re.I)
PASS_RESULT = re.compile(r"^[^=\n]+\s*=\s*PASS(?:\s*:\s*.+)?$", re.I)
PAGE_RESULT = re.compile(r"^(\d+)=(PASS|FAIL)(?::(.*))?$", re.I)
WORD_REQUIRED_FLAGS = {
    "macros",
    "digital_signatures",
    "document_protection",
    "embedded_objects",
    "fields",
    "zotero_or_csl_citations",
    "endnote_citations",
    "comments",
    "tracked_changes",
    "content_controls",
    "equations",
    "text_boxes",
    "footnotes",
    "endnotes",
    "multiple_sections",
    "alt_chunks",
}


def verify_record_fingerprint(entry: dict, label: str) -> tuple[Path | None, list[str]]:
    issues: list[str] = []
    try:
        path = Path(entry["path"]).resolve(strict=True)
        issues.extend(compare_fingerprint(entry, path, label))
        return path, issues
    except (OSError, KeyError, TypeError, ValueError) as exc:
        return None, [f"{label} is invalid: {exc}"]


def verify_pdf(path: Path) -> None:
    with path.open("rb") as handle:
        if handle.read(5) != b"%PDF-":
            raise ValueError(f"not a PDF file: {path}")


def verify_png(path: Path) -> tuple[int, int]:
    with Image.open(path) as image:
        image.verify()
    with Image.open(path) as image:
        image.load()
        width, height = image.size
    if width < 1 or height < 1:
        raise ValueError(f"invalid PNG dimensions: {path}")
    return width, height


def validate_renderer_manifest(
    manifest_path: Path, document_path: Path
) -> tuple[dict, list[str]]:
    manifest = manifest_path.resolve(strict=True)
    document = document_path.resolve(strict=True)
    payload = load_json(manifest)
    issues: list[str] = []
    if not isinstance(payload, dict):
        raise ValueError("renderer manifest root must be a JSON object")
    if payload.get("schema_version") != SCHEMA_VERSION:
        issues.append("renderer manifest has unsupported schema_version")
    if payload.get("record_type") != "documents_renderer_manifest":
        issues.append("renderer manifest has unexpected record_type")
    producer = payload.get("producer_script")
    expected_renderer = RENDERER_PRODUCERS.get(producer)
    if expected_renderer is None:
        issues.append("renderer manifest producer is unsupported")
    renderer = payload.get("renderer") or {}
    if not isinstance(renderer, dict):
        renderer = {}
        issues.append("renderer entry is invalid")
    if renderer.get("id") != expected_renderer:
        issues.append("renderer identity does not match its producer")
    if not str(renderer.get("version", "")).strip():
        issues.append("renderer version is missing")

    document_entry = payload.get("document") or {}
    issues.extend(compare_fingerprint(document_entry, document, "renderer document"))
    if payload.get("source_sha256_before") != document_entry.get("sha256"):
        issues.append("renderer source hash before rendering is inconsistent")
    if payload.get("source_sha256_after") != document_entry.get("sha256"):
        issues.append("renderer source hash after rendering is inconsistent")

    settings = payload.get("settings") or {}
    if not isinstance(settings, dict):
        settings = {}
        issues.append("renderer settings are invalid")
    if not isinstance(settings.get("dpi"), int) or not 72 <= settings.get("dpi", 0) <= 600:
        issues.append("renderer DPI is invalid")
    markup_mode = settings.get("markup_mode")
    if expected_renderer == "microsoft-word" and markup_mode not in {"include", "hide"}:
        issues.append("Word markup mode is invalid")
    if expected_renderer == "libreoffice" and markup_mode != "application-default":
        issues.append("LibreOffice markup mode is invalid")
    if expected_renderer == "microsoft-word" and settings.get("include_markup") is not (
        markup_mode == "include"
    ):
        issues.append("Word include_markup and markup_mode disagree")
    if expected_renderer == "libreoffice" and settings.get("include_markup") is not None:
        issues.append("LibreOffice include_markup must be null")

    application = payload.get("application_open") or {}
    if not isinstance(application, dict):
        application = {}
        issues.append("application-open evidence is invalid")
    if application.get("status") != "PASS":
        issues.append("application did not record a successful open")
    if application.get("opened_read_only") is not True:
        issues.append("application did not record read-only source opening")
    if application.get("export_status") != "PASS":
        issues.append("application export did not pass")
    if application.get("repair_warning_count") != 0:
        issues.append("application recorded a repair/recovery warning")
    repair_observation = application.get("repair_observation") or {}
    if not isinstance(repair_observation, dict):
        issues.append("application repair observation is invalid")
        repair_observation = {}
    if repair_observation.get("status") != "NO_DIAGNOSTIC_OBSERVED":
        issues.append("application did not record the required no-diagnostic-observed state")
    if repair_observation.get("observed_warning_count") != application.get("repair_warning_count"):
        issues.append("application repair observation count is inconsistent")
    if repair_observation.get("absence_proven") is not False:
        issues.append("application must not claim that silent repair absence was proven")
    if not str(repair_observation.get("capture_scope", "")).strip():
        issues.append("application repair observation lacks a capture scope")
    diagnostics = application.get("diagnostics")
    if not isinstance(diagnostics, list):
        issues.append("application diagnostics must be a list")
        diagnostics = []
    if any(REPAIR_LANGUAGE.search(str(item)) for item in diagnostics):
        issues.append("application diagnostics contain repair/recovery language")

    bound_files: list[tuple[str, Path | None]] = [
        ("renderer manifest", manifest),
        ("document", document),
    ]
    pdf_path, found = verify_record_fingerprint(payload.get("pdf") or {}, "rendered PDF")
    issues.extend(found)
    if pdf_path is not None:
        bound_files.append(("rendered PDF", pdf_path))
        try:
            verify_pdf(pdf_path)
        except (OSError, ValueError) as exc:
            issues.append(str(exc))

    pages = payload.get("pages")
    if not isinstance(pages, list) or not pages:
        issues.append("renderer manifest has no pages")
        pages = []
    actual_numbers = [entry.get("page") for entry in pages if isinstance(entry, dict)]
    if actual_numbers != list(range(1, len(pages) + 1)):
        issues.append("renderer page sequence is non-contiguous")
    page_records: list[dict] = []
    for expected, entry in enumerate(pages, 1):
        if not isinstance(entry, dict):
            issues.append(f"renderer page {expected} entry is invalid")
            continue
        image_path, found = verify_record_fingerprint(
            entry.get("image") or {}, f"rendered page {expected}"
        )
        issues.extend(found)
        if image_path is None:
            continue
        bound_files.append((f"rendered page {expected}", image_path))
        if image_path.name.lower() != f"page-{expected}.png":
            issues.append(f"rendered page {expected} has an unexpected filename")
        try:
            width, height = verify_png(image_path)
            page_records.append(
                {
                    "page": expected,
                    "image": fingerprint(image_path),
                    "width_px": width,
                    "height_px": height,
                }
            )
        except (OSError, ValueError) as exc:
            issues.append(f"rendered page {expected} failed PNG decoding: {exc}")

    toolchain = payload.get("toolchain")
    if not isinstance(toolchain, dict) or not toolchain:
        issues.append("renderer toolchain evidence is missing")
        toolchain = {}
    for label, entry in toolchain.items():
        if not isinstance(entry, dict):
            issues.append(f"renderer toolchain entry is invalid: {label}")
            continue
        tool_path, found = verify_record_fingerprint(entry, f"renderer tool {label}")
        issues.extend(found)
        if tool_path is not None:
            bound_files.append((f"renderer tool {label}", tool_path))
    if renderer.get("executable"):
        executable_path, found = verify_record_fingerprint(
            renderer["executable"], "renderer executable"
        )
        issues.extend(found)
        if executable_path is not None:
            bound_files.append(("renderer executable", executable_path))
    try:
        ensure_distinct(bound_files)
    except ValueError as exc:
        issues.append(str(exc))

    try:
        output_dir = Path(payload["output_directory"]).resolve(strict=True)
        if not output_dir.is_dir():
            issues.append("renderer output_directory is not a directory")
        expected_files = {
            path.resolve()
            for name, path in bound_files
            if path is not None and (name.startswith("rendered ") or name == "renderer manifest")
        }
        actual_files = {path.resolve() for path in output_dir.rglob("*") if path.is_file()}
        if actual_files != expected_files:
            issues.append("renderer output directory contains unbound or missing files")
    except (OSError, KeyError, TypeError, ValueError) as exc:
        issues.append(f"renderer output directory is invalid: {exc}")

    return {
        "record": fingerprint(manifest),
        "document": fingerprint(document),
        "renderer": expected_renderer,
        "renderer_version": renderer.get("version"),
        "settings": settings,
        "application_open": application,
        "pdf": fingerprint(pdf_path) if pdf_path else None,
        "pages": page_records,
    }, issues


def parse_page_results(values: list[str], expected_pages: list[int]) -> tuple[list[dict], list[str]]:
    results: dict[int, dict] = {}
    issues: list[str] = []
    for raw in values:
        match = PAGE_RESULT.fullmatch(raw.strip())
        if not match:
            issues.append(f"invalid page result: {raw!r}")
            continue
        page = int(match.group(1))
        if page in results:
            issues.append(f"duplicate page result: {page}")
            continue
        results[page] = {
            "page": page,
            "status": match.group(2).upper(),
            "note": (match.group(3) or "").strip(),
        }
    if sorted(results) != expected_pages:
        issues.append(
            f"page-review coverage mismatch: expected {expected_pages}, got {sorted(results)}"
        )
    return [results[page] for page in sorted(results)], issues


def command_page_review(args: argparse.Namespace) -> int:
    document = args.document.resolve(strict=True)
    manifest = args.renderer_manifest.resolve(strict=True)
    output = args.output.resolve()
    ensure_distinct(
        [("document", document), ("renderer manifest", manifest), ("page review", output)]
    )
    manifest_entry, issues = validate_renderer_manifest(manifest, document)
    expected_pages = [entry["page"] for entry in manifest_entry["pages"]]
    page_results, found = parse_page_results(args.page_result, expected_pages)
    issues.extend(found)
    if not args.method.strip():
        issues.append("page-review method is empty")
    if issues:
        emit({"status": "NOT_RECORDED", "issues": issues}, error=True)
        return 2
    image_by_page = {entry["page"]: entry["image"] for entry in manifest_entry["pages"]}
    for result in page_results:
        result["image"] = image_by_page[result["page"]]
    status = "PASS" if all(item["status"] == "PASS" for item in page_results) else "FAIL"
    payload = {
        "schema_version": SCHEMA_VERSION,
        "record_type": "documents_page_review",
        "producer_skill": "documents-finalize",
        "created_at_utc": utc_now(),
        "status": status,
        "document": fingerprint(document),
        "renderer_manifest": manifest_entry,
        "method": args.method.strip(),
        "page_results": page_results,
        "judgment_boundary": (
            "Each visual judgment is recorded by the reviewer and bound to an exact page image; "
            "the software verifies coverage and identity, not visual correctness itself."
        ),
    }
    write_json(output, payload)
    emit({"status": status, "output": str(output), "pages": len(page_results)}, error=status != "PASS")
    return 0 if status == "PASS" else 2


def validate_page_review(
    review_path: Path, document_path: Path, manifest_path: Path
) -> tuple[dict, list[str]]:
    review = review_path.resolve(strict=True)
    document = document_path.resolve(strict=True)
    manifest = manifest_path.resolve(strict=True)
    payload = load_json(review)
    issues: list[str] = []
    if payload.get("schema_version") != SCHEMA_VERSION:
        issues.append("page review has unsupported schema_version")
    if payload.get("record_type") != "documents_page_review":
        issues.append("page review has unexpected record_type")
    if payload.get("producer_skill") != "documents-finalize":
        issues.append("page review has unexpected producer")
    if payload.get("status") != "PASS":
        issues.append("page review is not PASS")
    issues.extend(compare_fingerprint(payload.get("document") or {}, document, "page-review document"))
    manifest_entry, found = validate_renderer_manifest(manifest, document)
    issues.extend(found)
    stored_manifest = (payload.get("renderer_manifest") or {}).get("record") or {}
    issues.extend(compare_fingerprint(stored_manifest, manifest, "page-review renderer manifest"))
    expected_pages = [entry["page"] for entry in manifest_entry["pages"]]
    results = payload.get("page_results")
    if not isinstance(results, list):
        results = []
        issues.append("page review results must be a list")
    if [entry.get("page") for entry in results if isinstance(entry, dict)] != expected_pages:
        issues.append("page review does not cover every rendered page exactly once")
    image_by_page = {entry["page"]: entry["image"] for entry in manifest_entry["pages"]}
    for entry in results:
        if not isinstance(entry, dict):
            issues.append("page review contains an invalid entry")
            continue
        page = entry.get("page")
        if entry.get("status") != "PASS":
            issues.append(f"page {page} is not PASS")
        if page in image_by_page:
            image_path = Path(image_by_page[page]["path"]).resolve(strict=True)
            issues.extend(
                compare_fingerprint(entry.get("image") or {}, image_path, f"page-review image {page}")
            )
    if not str(payload.get("method", "")).strip():
        issues.append("page review method is empty")
    try:
        ensure_distinct(
            [("page review", review), ("document", document), ("renderer manifest", manifest)]
        )
    except ValueError as exc:
        issues.append(str(exc))
    return {"record": fingerprint(review), "payload": payload, "manifest": manifest_entry}, issues


def command_render_evidence(args: argparse.Namespace) -> int:
    document = args.document.resolve(strict=True)
    manifest = args.renderer_manifest.resolve(strict=True)
    page_review = args.page_review.resolve(strict=True)
    output = args.output.resolve()
    ensure_distinct(
        [
            ("document", document),
            ("renderer manifest", manifest),
            ("page review", page_review),
            ("render evidence", output),
        ]
    )
    page_entry, issues = validate_page_review(page_review, document, manifest)
    if issues:
        emit({"status": "NOT_READY", "issues": issues}, error=True)
        return 2
    payload = {
        "schema_version": SCHEMA_VERSION,
        "record_type": "documents_render_evidence",
        "producer_skill": "documents-finalize",
        "created_at_utc": utc_now(),
        "status": "FULL_PAGE_INSPECTION_RECORDED",
        "document": fingerprint(document),
        "renderer_manifest": page_entry["manifest"],
        "page_review": fingerprint(page_review),
        "rendered_pages": len(page_entry["manifest"]["pages"]),
        "visually_inspected_pages": [
            entry["page"] for entry in page_entry["manifest"]["pages"]
        ],
        "visual_defect_count": 0,
    }
    write_json(output, payload)
    emit({"status": payload["status"], "output": str(output)})
    return 0


def verify_render_evidence(
    evidence_path: Path, final_path: Path
) -> tuple[dict, list[str]]:
    evidence = evidence_path.resolve(strict=True)
    final = final_path.resolve(strict=True)
    payload = load_json(evidence)
    issues: list[str] = []
    if payload.get("schema_version") != SCHEMA_VERSION:
        issues.append("render evidence has unsupported schema_version")
    if payload.get("record_type") != "documents_render_evidence":
        issues.append("render evidence has unexpected record_type")
    if payload.get("producer_skill") != "documents-finalize":
        issues.append("render evidence has unexpected producer")
    if payload.get("status") != "FULL_PAGE_INSPECTION_RECORDED":
        issues.append("render evidence does not record completed full-page inspection")
    issues.extend(compare_fingerprint(payload.get("document") or {}, final, "render-evidence document"))
    manifest_path, found = verify_record_fingerprint(
        (payload.get("renderer_manifest") or {}).get("record") or {},
        "render-evidence renderer manifest",
    )
    issues.extend(found)
    page_review_path, found = verify_record_fingerprint(
        payload.get("page_review") or {}, "render-evidence page review"
    )
    issues.extend(found)
    manifest_entry: dict = {}
    if manifest_path and page_review_path:
        page_entry, found = validate_page_review(page_review_path, final, manifest_path)
        issues.extend(found)
        manifest_entry = page_entry["manifest"]
    expected_pages = [entry["page"] for entry in manifest_entry.get("pages", [])]
    if payload.get("rendered_pages") != len(expected_pages):
        issues.append("render evidence page count mismatch")
    if payload.get("visually_inspected_pages") != expected_pages:
        issues.append("render evidence does not cover every rendered page")
    if payload.get("visual_defect_count") != 0:
        issues.append("render evidence records visual defects")
    return {
        "record": fingerprint(evidence),
        "payload": payload,
        "manifest": manifest_entry,
        "manifest_path": manifest_path,
    }, issues


def _require_pass_entries(values: list[str], label: str) -> list[str]:
    cleaned = [value.strip() for value in values if value.strip()]
    if not cleaned:
        raise ValueError(f"at least one {label} is required")
    invalid = [value for value in cleaned if not PASS_RESULT.fullmatch(value)]
    if invalid:
        raise ValueError(f"{label} entries must use NAME=PASS[: evidence]: {invalid}")
    return cleaned


def _preflight(args: argparse.Namespace) -> tuple[Path, dict, dict]:
    path = args.preflight.resolve(strict=True)
    payload, context, issues = validate_preflight(path)
    if issues:
        raise ValueError("preflight failed live validation: " + "; ".join(issues))
    return path, payload, context


def _base_review(record_type: str, preflight_path: Path, context: dict) -> dict:
    return {
        "schema_version": SCHEMA_VERSION,
        "record_type": record_type,
        "producer_skill": "documents-finalize",
        "created_at_utc": utc_now(),
        "status": "PASS",
        "preflight": fingerprint(preflight_path),
        "source": fingerprint(context["source"]),
        "candidate": fingerprint(context["candidate"]),
        "final": fingerprint(context["final"]),
        "comparison": fingerprint(context["comparison_path"]),
    }


def command_scope_review(args: argparse.Namespace) -> int:
    preflight_path, _, context = _preflight(args)
    output = args.output.resolve()
    ensure_distinct(
        [
            ("preflight", preflight_path),
            ("source", context["source"]),
            ("candidate", context["candidate"]),
            ("final", context["final"]),
            ("scope review", output),
        ]
    )
    if not args.authorized_mode.strip() or not args.method.strip():
        raise ValueError("authorized mode and review method must be non-empty")
    requirement_results = _require_pass_entries(args.requirement_result, "requirement result")
    baseline_delta = context["comparison"]["comparisons"]["baseline_to_final"]
    text_reviews = [item.strip() for item in args.text_change_review if item.strip()]
    if baseline_delta["visible_content"]["changed"]:
        text_reviews = _require_pass_entries(text_reviews, "text-change review")
    payload = _base_review("documents_scope_review", preflight_path, context)
    payload.update(
        {
            "authorized_mode": args.authorized_mode.strip(),
            "review_method": args.method.strip(),
            "requirement_results": requirement_results,
            "baseline_to_final": baseline_delta,
            "text_change_reviews": text_reviews,
            "unapproved_change_count": 0,
            "unresolved_material_issues": 0,
            "notes": args.note,
            "judgment_boundary": (
                "The comparison is machine-derived; requirement and semantic-scope judgments "
                "are reviewer assertions bound to the exact files."
            ),
        }
    )
    write_json(output, payload)
    emit({"status": "PASS", "output": str(output)})
    return 0


def _risk_deltas(preflight_payload: dict) -> list[str]:
    baseline = preflight_payload["risk"].get("source") or preflight_payload["risk"]["candidate"]
    final = preflight_payload["risk"]["final"]
    deltas: list[str] = []
    for group in ("counts", "feature_flags"):
        left = baseline.get(group, {})
        right = final.get(group, {})
        for key in sorted(set(left) | set(right)):
            if left.get(key) != right.get(key):
                deltas.append(f"{group}.{key}")
    return deltas


def _changed_parts(context: dict) -> list[str]:
    package = context["comparison"]["comparisons"]["baseline_to_final"]["package"]
    return sorted(
        set(package["changed_parts"]) | set(package["added_parts"]) | set(package["removed_parts"])
    )


def command_structure_review(args: argparse.Namespace) -> int:
    preflight_path, preflight_payload, context = _preflight(args)
    output = args.output.resolve()
    if not args.method.strip():
        raise ValueError("structure-review method is empty")
    expected_parts = _changed_parts(context)
    reviewed_parts = sorted({item.strip() for item in args.reviewed_part if item.strip()})
    if reviewed_parts != expected_parts:
        raise ValueError(
            f"reviewed package-part coverage mismatch: expected {expected_parts}, got {reviewed_parts}"
        )
    risk_deltas = _risk_deltas(preflight_payload)
    explanations = [item.strip() for item in args.delta_explanation if item.strip()]
    if risk_deltas:
        explanations = _require_pass_entries(explanations, "delta explanation")
        explained_keys = {item.split("=", 1)[0].strip() for item in explanations}
        if explained_keys != set(risk_deltas):
            raise ValueError(
                f"risk-delta explanation coverage mismatch: expected {risk_deltas}, got {sorted(explained_keys)}"
            )
    payload = _base_review("documents_structure_review", preflight_path, context)
    payload.update(
        {
            "review_method": args.method.strip(),
            "reviewed_package_parts": reviewed_parts,
            "risk_delta_keys": risk_deltas,
            "delta_explanations": explanations,
            "unexplained_delta_count": 0,
            "notes": args.note,
        }
    )
    write_json(output, payload)
    emit({"status": "PASS", "output": str(output), "reviewed_parts": len(reviewed_parts)})
    return 0


def _policy_issues(preflight_payload: dict, policies: dict[str, str]) -> list[str]:
    baseline = preflight_payload["risk"].get("source") or preflight_payload["risk"]["candidate"]
    final = preflight_payload["risk"]["final"]
    bc, fc = baseline["counts"], final["counts"]
    bf, ff = baseline["feature_flags"], final["feature_flags"]
    issues: list[str] = []
    if policies["comments"] == "preserve" and fc["comments"] != bc["comments"]:
        issues.append("comments=preserve but comment count changed")
    if policies["comments"] == "remove" and fc["comments"] != 0:
        issues.append("comments=remove but comments remain")
    if policies["revisions"] == "preserve" and fc["revisions"] != bc["revisions"]:
        issues.append("revisions=preserve but revision count changed")
    if policies["revisions"] in {"accept", "reject"} and fc["revisions"] != 0:
        issues.append("revision resolution policy selected but revisions remain")
    if policies["fields"] == "preserve-editable":
        if fc["fields"] != bc["fields"]:
            issues.append("fields=preserve-editable but field count changed")
        if final.get("field_markers_detected") != baseline.get("field_markers_detected"):
            issues.append("fields=preserve-editable but field marker classes changed")
    if policies["fields"] in {"materialize", "remove"} and fc["fields"] != 0:
        issues.append("field removal/materialization policy selected but fields remain")
    citation_keys = ("zotero_or_csl_citations", "endnote_citations")
    if policies["citations"] == "preserve-editable":
        for key in citation_keys:
            if ff[key] != bf[key]:
                issues.append(f"citations=preserve-editable but {key} changed")
    if policies["citations"] in {"flatten", "none"}:
        if any(ff[key] for key in citation_keys):
            issues.append("citation flatten/none policy selected but editable citation markers remain")
    return issues


def command_native_policy_review(args: argparse.Namespace) -> int:
    preflight_path, preflight_payload, context = _preflight(args)
    output = args.output.resolve()
    if not args.method.strip():
        raise ValueError("native-policy review method is empty")
    policies = {
        "comments": args.comments_policy,
        "revisions": args.revisions_policy,
        "fields": args.fields_policy,
        "citations": args.citations_policy,
    }
    issues = _policy_issues(preflight_payload, policies)
    if issues:
        emit({"status": "FAIL", "issues": issues}, error=True)
        return 2
    payload = _base_review("documents_native_policy_review", preflight_path, context)
    payload.update(
        {
            "policies": policies,
            "review_method": args.method.strip(),
            "policy_violation_count": 0,
            "candidate_native_features": preflight_payload["risk"]["candidate"],
            "final_native_features": preflight_payload["risk"]["final"],
            "notes": args.note,
        }
    )
    write_json(output, payload)
    emit({"status": "PASS", "output": str(output)})
    return 0


def command_application_review(args: argparse.Namespace) -> int:
    preflight_path, preflight_payload, context = _preflight(args)
    final = context["final"]
    manifest = args.renderer_manifest.resolve(strict=True)
    output = args.output.resolve()
    ensure_distinct(
        [
            ("preflight", preflight_path),
            ("final", final),
            ("renderer manifest", manifest),
            ("application review", output),
        ]
    )
    manifest_entry, issues = validate_renderer_manifest(manifest, final)
    active_flags = sorted(
        key
        for key, value in preflight_payload["risk"]["final"]["feature_flags"].items()
        if value and key in WORD_REQUIRED_FLAGS
    )
    if active_flags and manifest_entry.get("renderer") != "microsoft-word":
        issues.append(
            "Microsoft Word evidence is required for active native features: "
            + ", ".join(active_flags)
        )
    if not args.method.strip():
        issues.append("application-review method is empty")
    if issues:
        emit({"status": "FAIL", "issues": issues}, error=True)
        return 2
    payload = _base_review("documents_application_review", preflight_path, context)
    payload.update(
        {
            "renderer_manifest": manifest_entry,
            "review_method": args.method.strip(),
            "word_required_feature_flags": active_flags,
            "application_open": manifest_entry["application_open"],
            "repair_warning_count": manifest_entry["application_open"]["repair_warning_count"],
            "repair_observation": manifest_entry["application_open"]["repair_observation"],
            "behavior_issue_count": 0,
            "notes": args.note,
        }
    )
    write_json(output, payload)
    emit({"status": "PASS", "output": str(output), "renderer": manifest_entry["renderer"]})
    return 0


def _validate_review_base(
    payload: dict, record_type: str, preflight_path: Path, context: dict
) -> list[str]:
    issues: list[str] = []
    if payload.get("schema_version") != SCHEMA_VERSION:
        issues.append(f"{record_type} has unsupported schema_version")
    if payload.get("record_type") != record_type:
        issues.append(f"unexpected review record_type; expected {record_type}")
    if payload.get("producer_skill") != "documents-finalize" or payload.get("status") != "PASS":
        issues.append(f"{record_type} is not a PASS record from documents-finalize")
    issues.extend(compare_fingerprint(payload.get("preflight") or {}, preflight_path, "review preflight"))
    for label in ("source", "candidate", "final"):
        path = context[label]
        entry = payload.get(label)
        if path is None:
            if entry is not None:
                issues.append(f"review unexpectedly stores {label}")
        else:
            issues.extend(compare_fingerprint(entry or {}, path, f"review {label}"))
    issues.extend(
        compare_fingerprint(
            payload.get("comparison") or {}, context["comparison_path"], "review comparison"
        )
    )
    return issues


def validate_scope_review(path: Path, preflight_path: Path) -> list[str]:
    payload = load_json(path.resolve(strict=True))
    _, context, preflight_issues = validate_preflight(preflight_path)
    issues = list(preflight_issues)
    issues.extend(_validate_review_base(payload, "documents_scope_review", preflight_path, context))
    if not str(payload.get("authorized_mode", "")).strip() or not str(
        payload.get("review_method", "")
    ).strip():
        issues.append("scope review lacks authorized mode or method")
    try:
        _require_pass_entries(payload.get("requirement_results") or [], "requirement result")
    except ValueError as exc:
        issues.append(str(exc))
    baseline = context.get("comparison", {}).get("comparisons", {}).get("baseline_to_final", {})
    if payload.get("baseline_to_final") != baseline:
        issues.append("scope review baseline comparison differs from live preflight")
    if baseline.get("visible_content", {}).get("changed"):
        try:
            _require_pass_entries(payload.get("text_change_reviews") or [], "text-change review")
        except ValueError as exc:
            issues.append(str(exc))
    if payload.get("unapproved_change_count") != 0 or payload.get("unresolved_material_issues") != 0:
        issues.append("scope review records unresolved or unapproved changes")
    return issues


def validate_structure_review(path: Path, preflight_path: Path) -> list[str]:
    payload = load_json(path.resolve(strict=True))
    preflight_payload, context, preflight_issues = validate_preflight(preflight_path)
    issues = list(preflight_issues)
    issues.extend(
        _validate_review_base(payload, "documents_structure_review", preflight_path, context)
    )
    expected_parts = _changed_parts(context)
    if payload.get("reviewed_package_parts") != expected_parts:
        issues.append("structure review does not cover every live-changed package part")
    expected_deltas = _risk_deltas(preflight_payload)
    if payload.get("risk_delta_keys") != expected_deltas:
        issues.append("structure review risk deltas differ from live preflight")
    if expected_deltas:
        try:
            entries = _require_pass_entries(
                payload.get("delta_explanations") or [], "delta explanation"
            )
            keys = {entry.split("=", 1)[0].strip() for entry in entries}
            if keys != set(expected_deltas):
                issues.append("structure review does not explain every live risk delta")
        except ValueError as exc:
            issues.append(str(exc))
    if payload.get("unexplained_delta_count") != 0:
        issues.append("structure review records unexplained deltas")
    if not str(payload.get("review_method", "")).strip():
        issues.append("structure review method is empty")
    return issues


def validate_native_review(path: Path, preflight_path: Path) -> list[str]:
    payload = load_json(path.resolve(strict=True))
    preflight_payload, context, preflight_issues = validate_preflight(preflight_path)
    issues = list(preflight_issues)
    issues.extend(
        _validate_review_base(payload, "documents_native_policy_review", preflight_path, context)
    )
    policies = payload.get("policies")
    if not isinstance(policies, dict):
        issues.append("native policy set is invalid")
    else:
        issues.extend(_policy_issues(preflight_payload, policies))
    if payload.get("policy_violation_count") != 0:
        issues.append("native policy review records violations")
    if not str(payload.get("review_method", "")).strip():
        issues.append("native policy review method is empty")
    return issues


def validate_application_review(path: Path, preflight_path: Path) -> list[str]:
    payload = load_json(path.resolve(strict=True))
    preflight_payload, context, preflight_issues = validate_preflight(preflight_path)
    issues = list(preflight_issues)
    issues.extend(
        _validate_review_base(payload, "documents_application_review", preflight_path, context)
    )
    manifest_path, found = verify_record_fingerprint(
        (payload.get("renderer_manifest") or {}).get("record") or {},
        "application-review renderer manifest",
    )
    issues.extend(found)
    manifest_entry: dict = {}
    if manifest_path:
        manifest_entry, found = validate_renderer_manifest(manifest_path, context["final"])
        issues.extend(found)
    active_flags = sorted(
        key
        for key, value in preflight_payload["risk"]["final"]["feature_flags"].items()
        if value and key in WORD_REQUIRED_FLAGS
    )
    if active_flags and manifest_entry.get("renderer") != "microsoft-word":
        issues.append("application review lacks required Microsoft Word evidence")
    if payload.get("word_required_feature_flags") != active_flags:
        issues.append("application review Word-required feature list differs from live preflight")
    if payload.get("application_open") != manifest_entry.get("application_open"):
        issues.append("application-open evidence differs from live renderer manifest")
    if payload.get("repair_observation") != manifest_entry.get("application_open", {}).get(
        "repair_observation"
    ):
        issues.append("application repair observation differs from live renderer manifest")
    if payload.get("repair_warning_count") != 0 or payload.get("behavior_issue_count") != 0:
        issues.append("application review records repair warnings or behavior issues")
    if not str(payload.get("review_method", "")).strip():
        issues.append("application review method is empty")
    return issues


def inventory_delivery(
    delivery_dir: Path, deliverables: list[Path], final: Path
) -> tuple[list[str], list[str]]:
    issues: list[str] = []
    directory = delivery_dir.resolve(strict=True)
    if not directory.is_dir():
        return [], ["delivery path is not a directory"]
    resolved_deliverables = [path.resolve(strict=True) for path in deliverables]
    if not any(same_file(path, final) for path in resolved_deliverables):
        issues.append("the exact final DOCX is not listed as a deliverable")
    for path in resolved_deliverables:
        if not path.is_relative_to(directory):
            issues.append(f"deliverable is outside delivery directory: {path}")
    if len({str(path) for path in resolved_deliverables}) != len(resolved_deliverables):
        issues.append("deliverable list contains duplicates")
    actual = sorted(
        str(path.resolve().relative_to(directory))
        for path in directory.rglob("*")
        if path.is_file()
    )
    expected = sorted(str(path.relative_to(directory)) for path in resolved_deliverables)
    if actual != expected:
        issues.append(f"delivery inventory mismatch: expected {expected}, got {actual}")
    return actual, issues


def validate_operation_completion(
    completion_path: Path, candidate: Path, final: Path
) -> tuple[dict, list[str]]:
    completion = completion_path.resolve(strict=True)
    payload = load_json(completion)
    issues: list[str] = []
    if payload.get("schema_version") != OPERATION_SCHEMA_VERSION:
        issues.append("operation completion has unsupported schema_version")
    if payload.get("record_type") != "documents_artifact_operation_completion":
        issues.append("operation completion has unexpected record_type")
    if payload.get("status") != "VERIFIED":
        issues.append("operation completion is not VERIFIED")

    start_path, found = verify_record_fingerprint(
        payload.get("start_receipt") or {}, "operation start receipt"
    )
    issues.extend(found)
    start_payload: dict = {}
    if start_path:
        start_payload = load_json(start_path)
        if start_payload.get("schema_version") != OPERATION_SCHEMA_VERSION:
            issues.append("operation start receipt has unsupported schema_version")
        if start_payload.get("record_type") != "documents_artifact_operation_start":
            issues.append("operation start receipt has unexpected record_type")
        if start_payload.get("status") != "STARTED":
            issues.append("operation start receipt is not STARTED")
        if start_payload.get("operation_id") != payload.get("operation_id"):
            issues.append("operation start/completion IDs differ")
        if start_payload.get("operation_kind") != "edit":
            issues.append("final release operation must be recorded as edit")
        if start_payload.get("output_format") != "docx":
            issues.append("final release operation format is not docx")

    def matching_entry(entries: list[dict], target: Path, label: str) -> dict | None:
        matches: list[dict] = []
        for entry in entries:
            try:
                entry_path = Path(entry["path"]).resolve(strict=True)
                if same_file(entry_path, target):
                    matches.append(entry)
            except (OSError, KeyError, TypeError, ValueError):
                continue
        if len(matches) != 1:
            issues.append(f"operation evidence must contain exactly one {label} entry")
            return None
        issues.extend(compare_fingerprint(matches[0], target, f"operation {label}"))
        return matches[0]

    matching_entry(payload.get("inputs") or [], candidate, "candidate input")
    matching_entry(payload.get("outputs") or [], final, "final output")
    planned = start_payload.get("planned_outputs") or []
    planned_matches = 0
    for entry in planned:
        try:
            if same_file(Path(entry["path"]).resolve(strict=True), final):
                planned_matches += 1
        except (OSError, KeyError, TypeError, ValueError):
            continue
    if planned_matches != 1:
        issues.append("operation start receipt does not target the exact final output once")
    if start_payload and len(planned) != start_payload.get("expected_output_count"):
        issues.append("operation start receipt planned-output count is inconsistent")

    verifier_script, found = verify_record_fingerprint(
        (payload.get("verifier") or {}).get("script") or {}, "operation verifier script"
    )
    issues.extend(found)
    if verifier_script is None:
        issues.append("operation completion lacks a live verifier script")
    return {"record": fingerprint(completion), "start_receipt": payload.get("start_receipt")}, issues


def command_qa_review(args: argparse.Namespace) -> int:
    preflight_path, preflight_payload, context = _preflight(args)
    source = context["source"]
    candidate = context["candidate"]
    final = context["final"]
    handoff = args.handoff.resolve(strict=True) if args.handoff else None
    render_evidence = args.render_evidence.resolve(strict=True)
    scope_review = args.scope_review.resolve(strict=True)
    structure_review = args.structure_review.resolve(strict=True)
    native_review = args.native_policy_review.resolve(strict=True)
    application_review = args.application_review.resolve(strict=True)
    operation_completion = args.operation_completion.resolve(strict=True)
    delivery_dir = args.delivery_dir.resolve(strict=True)
    deliverables = [path.resolve(strict=True) for path in args.deliverable]
    output = args.output.resolve()
    issues: list[str] = []
    if preflight_payload["provenance_mode"] == "new-document" and handoff is None:
        issues.append("new-document finalization requires a verified documents-fast handoff")
    try:
        ensure_distinct(
            [
                ("preflight", preflight_path),
                ("source", source),
                ("candidate", candidate),
                ("final", final),
                ("handoff", handoff),
                ("render evidence", render_evidence),
                ("scope review", scope_review),
                ("structure review", structure_review),
                ("native review", native_review),
                ("application review", application_review),
                ("operation completion", operation_completion),
                ("QA review", output),
            ]
        )
    except ValueError as exc:
        issues.append(str(exc))
    handoff_ref = None
    if handoff:
        _, found = verify_payload(handoff, candidate, source)
        issues.extend(f"handoff: {item}" for item in found)
        handoff_ref = fingerprint(handoff)
    render_entry, found = verify_render_evidence(render_evidence, final)
    issues.extend(f"visual QA: {item}" for item in found)
    issues.extend(f"scope: {item}" for item in validate_scope_review(scope_review, preflight_path))
    issues.extend(
        f"structure: {item}" for item in validate_structure_review(structure_review, preflight_path)
    )
    issues.extend(
        f"native policy: {item}" for item in validate_native_review(native_review, preflight_path)
    )
    issues.extend(
        f"application: {item}"
        for item in validate_application_review(application_review, preflight_path)
    )
    operation_entry, found = validate_operation_completion(operation_completion, candidate, final)
    issues.extend(f"operation: {item}" for item in found)
    inventory, found = inventory_delivery(delivery_dir, deliverables, final)
    issues.extend(found)
    evidence_paths = [
        preflight_path,
        handoff,
        render_evidence,
        scope_review,
        structure_review,
        native_review,
        application_review,
        operation_completion,
        render_entry.get("manifest_path"),
        output,
    ]
    for path in (item for item in evidence_paths if item is not None):
        if path.resolve().is_relative_to(delivery_dir):
            issues.append(f"QA evidence/output must be outside delivery: {path}")
    if issues:
        emit({"status": "NOT_READY", "issues": issues}, error=True)
        return 2
    review_refs = {
        "scope": fingerprint(scope_review),
        "structure": fingerprint(structure_review),
        "native_policy": fingerprint(native_review),
        "application": fingerprint(application_review),
    }
    gates = {
        "identity_provenance": {
            "status": "PASS",
            "evidence": ([handoff_ref] if handoff_ref else [fingerprint(preflight_path)])
            + [operation_entry["record"]],
        },
        "package_integrity": {
            "status": "PASS",
            "evidence": [preflight_payload["risk"]["final"]["report"]],
        },
        "authorized_scope_content": {"status": "PASS", "evidence": [review_refs["scope"]]},
        "protected_structures": {"status": "PASS", "evidence": [review_refs["structure"]]},
        "native_features_policy": {
            "status": "PASS",
            "evidence": [review_refs["native_policy"]],
        },
        "application_behavior": {
            "status": "PASS",
            "evidence": [review_refs["application"]],
        },
        "visual_qa": {"status": "PASS", "evidence": [fingerprint(render_evidence)]},
        "output_hygiene": {
            "status": "PASS",
            "evidence": [{"delivery_directory": str(delivery_dir), "files": inventory}],
        },
    }
    payload = {
        "schema_version": SCHEMA_VERSION,
        "record_type": "documents_qa_review",
        "producer_skill": "documents-finalize",
        "created_at_utc": utc_now(),
        "status": "READY_FOR_RELEASE",
        "provenance_mode": preflight_payload["provenance_mode"],
        "source": fingerprint(source),
        "candidate": fingerprint(candidate),
        "final": fingerprint(final),
        "preflight": fingerprint(preflight_path),
        "handoff": handoff_ref,
        "operation_completion": operation_entry["record"],
        "render_evidence": fingerprint(render_evidence),
        "reviews": review_refs,
        "delivery": {
            "directory": str(delivery_dir),
            "deliverables": [fingerprint(path) for path in deliverables],
            "inventory": inventory,
        },
        "unresolved_material_issues": 0,
        "gates": gates,
    }
    write_json(output, payload)
    emit({"status": "READY_FOR_RELEASE", "output": str(output)})
    return 0


def verify_qa_review(
    qa_path: Path, final_override: Path | None = None
) -> tuple[dict, list[str]]:
    qa = qa_path.resolve(strict=True)
    payload = load_json(qa)
    issues: list[str] = []
    if payload.get("schema_version") != SCHEMA_VERSION:
        issues.append("QA review has unsupported schema_version")
    if payload.get("record_type") != "documents_qa_review":
        issues.append("QA review has unexpected record_type")
    if payload.get("producer_skill") != "documents-finalize":
        issues.append("QA review has unexpected producer")
    if payload.get("status") != "READY_FOR_RELEASE":
        issues.append("QA review is not READY_FOR_RELEASE")
    if payload.get("unresolved_material_issues") != 0:
        issues.append("QA review has unresolved material issues")
    gates = payload.get("gates") or {}
    if set(gates) != set(REQUIRED_GATES):
        issues.append("QA review gate set is incomplete or unexpected")
    for gate in REQUIRED_GATES:
        if (gates.get(gate) or {}).get("status") != "PASS":
            issues.append(f"QA gate is not PASS: {gate}")

    preflight_path, found = verify_record_fingerprint(
        payload.get("preflight") or {}, "QA preflight"
    )
    issues.extend(found)
    context: dict = {}
    preflight_payload: dict = {}
    if preflight_path:
        preflight_payload, context, found = validate_preflight(preflight_path)
        issues.extend(f"QA preflight: {item}" for item in found)
    if not context:
        return payload, issues
    source, candidate, final = context["source"], context["candidate"], context["final"]
    for label, path in (("source", source), ("candidate", candidate), ("final", final)):
        entry = payload.get(label)
        if path is None:
            if entry is not None:
                issues.append(f"QA unexpectedly stores {label}")
        else:
            issues.extend(compare_fingerprint(entry or {}, path, f"QA {label}"))
    if final_override:
        supplied = final_override.resolve(strict=True)
        if not same_file(final, supplied):
            issues.append("supplied final differs from preflight/QA final")
        issues.extend(compare_fingerprint(payload.get("final") or {}, supplied, "supplied final"))

    handoff_path = None
    if payload.get("handoff"):
        handoff_path, found = verify_record_fingerprint(payload["handoff"], "QA handoff")
        issues.extend(found)
        if handoff_path:
            _, found = verify_payload(handoff_path, candidate, source)
            issues.extend(f"QA handoff: {item}" for item in found)
    elif preflight_payload.get("provenance_mode") == "new-document":
        issues.append("new-document QA lacks a typed documents-fast handoff")

    operation_path, found = verify_record_fingerprint(
        payload.get("operation_completion") or {}, "QA operation completion"
    )
    issues.extend(found)
    if operation_path:
        _, found = validate_operation_completion(operation_path, candidate, final)
        issues.extend(f"QA operation: {item}" for item in found)

    render_path, found = verify_record_fingerprint(
        payload.get("render_evidence") or {}, "QA render evidence"
    )
    issues.extend(found)
    if render_path:
        _, found = verify_render_evidence(render_path, final)
        issues.extend(f"QA visual evidence: {item}" for item in found)

    review_paths: dict[str, Path | None] = {}
    for label in ("scope", "structure", "native_policy", "application"):
        path, found = verify_record_fingerprint(
            (payload.get("reviews") or {}).get(label) or {}, f"QA {label} review"
        )
        review_paths[label] = path
        issues.extend(found)
    if review_paths["scope"]:
        issues.extend(validate_scope_review(review_paths["scope"], preflight_path))
    if review_paths["structure"]:
        issues.extend(validate_structure_review(review_paths["structure"], preflight_path))
    if review_paths["native_policy"]:
        issues.extend(validate_native_review(review_paths["native_policy"], preflight_path))
    if review_paths["application"]:
        issues.extend(validate_application_review(review_paths["application"], preflight_path))

    try:
        delivery_dir = Path(payload["delivery"]["directory"]).resolve(strict=True)
        deliverables = [
            Path(entry["path"]).resolve(strict=True)
            for entry in payload["delivery"]["deliverables"]
        ]
        for entry, path in zip(payload["delivery"]["deliverables"], deliverables):
            issues.extend(compare_fingerprint(entry, path, "QA deliverable"))
        inventory, found = inventory_delivery(delivery_dir, deliverables, final)
        issues.extend(found)
        if inventory != payload["delivery"].get("inventory"):
            issues.append("QA stored delivery inventory differs from live inventory")
        for evidence in [
            qa,
            preflight_path,
            handoff_path,
            operation_path,
            render_path,
            *review_paths.values(),
        ]:
            if evidence and evidence.resolve().is_relative_to(delivery_dir):
                issues.append(f"QA evidence is inside delivery directory: {evidence}")
    except (OSError, KeyError, TypeError, ValueError) as exc:
        issues.append(f"QA delivery evidence is invalid: {exc}")
    return payload, issues


def command_release(args: argparse.Namespace) -> int:
    final = args.final.resolve(strict=True)
    qa_review = args.qa_review.resolve(strict=True)
    output = args.output.resolve()
    ensure_distinct([("final", final), ("QA review", qa_review), ("release record", output)])
    qa_payload, issues = verify_qa_review(qa_review, final)
    try:
        delivery_dir = Path(qa_payload["delivery"]["directory"]).resolve(strict=True)
        if output.is_relative_to(delivery_dir):
            issues.append("release record must be outside the delivery directory")
    except (OSError, KeyError, TypeError, ValueError) as exc:
        issues.append(f"release delivery boundary is invalid: {exc}")
    if issues:
        emit({"status": "NOT_RELEASED", "issues": issues}, error=True)
        return 2
    release = {
        "schema_version": SCHEMA_VERSION,
        "record_type": "documents_release_record",
        "producer_skill": "documents-finalize",
        "created_at_utc": utc_now(),
        "artifact_status": "RELEASED",
        "release_eligible": True,
        "provenance_mode": qa_payload["provenance_mode"],
        "source": qa_payload.get("source"),
        "candidate": qa_payload["candidate"],
        "final": fingerprint(final),
        "preflight": qa_payload["preflight"],
        "qa_review": fingerprint(qa_review),
        "operation_completion": qa_payload["operation_completion"],
        "gates": qa_payload["gates"],
        "release_decision": {
            "status": "PASS",
            "unresolved_material_issues": 0,
            "all_required_gates_passed": True,
        },
    }
    write_json(output, release)
    emit({"status": "RELEASED", "output": str(output), "final_sha256": release["final"]["sha256"]})
    return 0


def verify_release_record(record_path: Path) -> tuple[dict, list[str]]:
    record = record_path.resolve(strict=True)
    payload = load_json(record)
    issues: list[str] = []
    if payload.get("schema_version") != SCHEMA_VERSION:
        issues.append("release record has unsupported schema_version")
    if payload.get("record_type") != "documents_release_record":
        issues.append("release record has unexpected record_type")
    if payload.get("producer_skill") != "documents-finalize":
        issues.append("release record has unexpected producer")
    if payload.get("artifact_status") != "RELEASED" or payload.get("release_eligible") is not True:
        issues.append("release record is not RELEASED")
    decision = payload.get("release_decision") or {}
    if (
        decision.get("status") != "PASS"
        or decision.get("unresolved_material_issues") != 0
        or decision.get("all_required_gates_passed") is not True
    ):
        issues.append("release decision is not PASS")

    final, found = verify_record_fingerprint(payload.get("final") or {}, "release final")
    issues.extend(found)
    qa_review, found = verify_record_fingerprint(
        payload.get("qa_review") or {}, "release QA review"
    )
    issues.extend(found)
    delivery_dir: Path | None = None
    if final and qa_review:
        qa_payload, found = verify_qa_review(qa_review, final)
        issues.extend(f"release QA: {item}" for item in found)
        for key in ("source", "candidate", "final", "preflight", "operation_completion", "gates"):
            if payload.get(key) != qa_payload.get(key):
                issues.append(f"release record {key} differs from its QA review")
        try:
            delivery_dir = Path(qa_payload["delivery"]["directory"]).resolve(strict=True)
            if record.is_relative_to(delivery_dir):
                issues.append("release record is inside the internal delivery directory")
        except (OSError, KeyError, TypeError, ValueError) as exc:
            issues.append(f"release record delivery boundary is invalid: {exc}")
    try:
        ensure_distinct(
            [("release record", record), ("release final", final), ("release QA review", qa_review)]
        )
    except ValueError as exc:
        issues.append(str(exc))
    return {
        "record": fingerprint(record),
        "payload": payload,
        "final": final,
        "delivery_directory": delivery_dir,
    }, issues


def command_deliver(args: argparse.Namespace) -> int:
    release_record = args.release_record.resolve(strict=True)
    input_at_start = args.input_at_start.expanduser().resolve(strict=False)
    destination = args.destination.expanduser().resolve(strict=False)
    output = args.output.expanduser().resolve(strict=False)
    release_entry, issues = verify_release_record(release_record)
    final = release_entry.get("final")
    internal_delivery = release_entry.get("delivery_directory")
    try:
        input_directory = input_at_start.parent.resolve(strict=True)
        if not input_directory.is_dir():
            issues.append("input-at-start parent is not a directory")
        if not same_file(destination.parent, input_directory):
            issues.append("delivery destination is not beside the input-at-start document")
    except OSError as exc:
        issues.append(f"input-at-start directory is invalid: {exc}")
        input_directory = input_at_start.parent.resolve(strict=False)
    if final and destination.suffix.lower() != final.suffix.lower():
        issues.append("delivery destination extension differs from the certified final")
    if internal_delivery and output.is_relative_to(internal_delivery):
        issues.append("delivery record must be outside the internal delivery directory")
    if issues:
        emit({"status": "NOT_DELIVERED", "issues": issues}, error=True)
        return 2

    ensure_new_file(
        destination,
        inputs=[release_record, final],
        other_outputs=[output],
        suffixes=[final.suffix],
        create_parent=False,
    )
    ensure_new_file(
        output,
        inputs=[release_record, final],
        other_outputs=[destination],
        suffixes=[".json"],
    )
    created_destination = False
    try:
        with final.open("rb") as source_handle, destination.open("xb") as destination_handle:
            created_destination = True
            shutil.copyfileobj(source_handle, destination_handle)
        post_release_entry, post_release_issues = verify_release_record(release_record)
        if post_release_issues:
            raise ValueError(
                "release chain changed during delivery: " + "; ".join(post_release_issues)
            )
        if post_release_entry["payload"].get("final") != release_entry["payload"].get("final"):
            raise ValueError("release final changed during delivery")
        certified = post_release_entry["payload"]["final"]
        delivered = fingerprint(destination)
        if (
            certified["sha256"] != delivered["sha256"]
            or certified["size_bytes"] != delivered["size_bytes"]
        ):
            raise ValueError("delivered copy is not byte-identical to the certified final")
        payload = {
            "schema_version": DELIVERY_SCHEMA_VERSION,
            "record_type": "documents_release_delivery",
            "producer_skill": "documents-finalize",
            "created_at_utc": utc_now(),
            "status": "DELIVERED",
            "release_record": post_release_entry["record"],
            "certified_final": certified,
            "input_at_start": str(input_at_start),
            "input_directory": str(input_directory),
            "delivered_final": delivered,
            "same_directory": True,
            "byte_identical": True,
        }
        write_json(output, payload)
    except BaseException:
        if created_destination:
            destination.unlink(missing_ok=True)
        raise
    emit(
        {
            "status": "DELIVERED",
            "output": str(output),
            "delivered_final": str(destination),
            "final_sha256": delivered["sha256"],
        }
    )
    return 0


def verify_delivery_record(record_path: Path) -> tuple[dict, list[str]]:
    record = record_path.resolve(strict=True)
    payload = load_json(record)
    issues: list[str] = []
    if payload.get("schema_version") != DELIVERY_SCHEMA_VERSION:
        issues.append("delivery record has unsupported schema_version")
    if payload.get("record_type") != "documents_release_delivery":
        issues.append("delivery record has unexpected record_type")
    if payload.get("producer_skill") != "documents-finalize":
        issues.append("delivery record has unexpected producer")
    if payload.get("status") != "DELIVERED":
        issues.append("delivery record is not DELIVERED")
    if payload.get("same_directory") is not True or payload.get("byte_identical") is not True:
        issues.append("delivery record does not assert same-directory byte identity")

    release_record, found = verify_record_fingerprint(
        payload.get("release_record") or {}, "delivery release record"
    )
    issues.extend(found)
    release_entry: dict = {}
    if release_record:
        release_entry, found = verify_release_record(release_record)
        issues.extend(f"delivery release: {item}" for item in found)
    final = release_entry.get("final")
    if final:
        issues.extend(
            compare_fingerprint(
                payload.get("certified_final") or {}, final, "delivery certified final"
            )
        )
    delivered, found = verify_record_fingerprint(
        payload.get("delivered_final") or {}, "delivered final"
    )
    issues.extend(found)
    try:
        input_at_start = Path(payload["input_at_start"]).expanduser().resolve(strict=False)
        input_directory = Path(payload["input_directory"]).expanduser().resolve(strict=True)
        if input_at_start.parent.resolve(strict=False) != input_directory:
            issues.append("delivery input-at-start path disagrees with its stored directory")
        if delivered and not same_file(delivered.parent, input_directory):
            issues.append("delivered final is no longer beside the input-at-start document")
    except (OSError, KeyError, TypeError, ValueError) as exc:
        issues.append(f"delivery input directory is invalid: {exc}")
    if final and delivered:
        certified = fingerprint(final)
        live_delivery = fingerprint(delivered)
        if (
            certified["sha256"] != live_delivery["sha256"]
            or certified["size_bytes"] != live_delivery["size_bytes"]
        ):
            issues.append("delivered final is not byte-identical to the certified final")
        try:
            ensure_distinct(
                [
                    ("delivery record", record),
                    ("release record", release_record),
                    ("certified final", final),
                    ("delivered final", delivered),
                ]
            )
        except ValueError as exc:
            issues.append(str(exc))
    return {"record": fingerprint(record), "payload": payload}, issues


def command_verify_delivery(args: argparse.Namespace) -> int:
    entry, issues = verify_delivery_record(args.delivery_record)
    if issues:
        emit({"status": "NOT_DELIVERED", "issues": issues}, error=True)
        return 2
    payload = entry["payload"]
    emit(
        {
            "status": "DELIVERED",
            "delivery_record": entry["record"]["path"],
            "delivered_final": payload["delivered_final"]["path"],
            "final_sha256": payload["delivered_final"]["sha256"],
        }
    )
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    page = sub.add_parser(
        "page-review", help="Record exactly one PASS/FAIL judgment for every rendered page"
    )
    page.add_argument("--document", type=Path, required=True)
    page.add_argument("--renderer-manifest", type=Path, required=True)
    page.add_argument("--output", type=Path, required=True)
    page.add_argument(
        "--page-result",
        action="append",
        required=True,
        help="PAGE=PASS[: note] or PAGE=FAIL[: defect]; repeat exactly once per page",
    )
    page.add_argument("--method", required=True)
    page.set_defaults(handler=command_page_review)

    render = sub.add_parser(
        "render-evidence", help="Bind a complete PASS page review to the live renderer manifest"
    )
    render.add_argument("--document", type=Path, required=True)
    render.add_argument("--renderer-manifest", type=Path, required=True)
    render.add_argument("--page-review", type=Path, required=True)
    render.add_argument("--output", type=Path, required=True)
    render.set_defaults(handler=command_render_evidence)

    scope = sub.add_parser(
        "scope-review", help="Record authorized-scope and semantic-content review"
    )
    scope.add_argument("--preflight", type=Path, required=True)
    scope.add_argument("--output", type=Path, required=True)
    scope.add_argument("--authorized-mode", required=True)
    scope.add_argument("--method", required=True)
    scope.add_argument("--requirement-result", action="append", required=True)
    scope.add_argument("--text-change-review", action="append", default=[])
    scope.add_argument("--note", action="append", default=[])
    scope.set_defaults(handler=command_scope_review)

    structure = sub.add_parser(
        "structure-review", help="Require review coverage for every changed package part"
    )
    structure.add_argument("--preflight", type=Path, required=True)
    structure.add_argument("--output", type=Path, required=True)
    structure.add_argument("--method", required=True)
    structure.add_argument("--reviewed-part", action="append", default=[])
    structure.add_argument("--delta-explanation", action="append", default=[])
    structure.add_argument("--note", action="append", default=[])
    structure.set_defaults(handler=command_structure_review)

    native = sub.add_parser(
        "native-policy-review", help="Derive native-feature policy compliance from fresh risk scans"
    )
    native.add_argument("--preflight", type=Path, required=True)
    native.add_argument("--output", type=Path, required=True)
    native.add_argument("--comments-policy", choices=["preserve", "remove"], required=True)
    native.add_argument(
        "--revisions-policy", choices=["preserve", "accept", "reject"], required=True
    )
    native.add_argument(
        "--fields-policy",
        choices=["preserve-editable", "materialize", "remove"],
        required=True,
    )
    native.add_argument(
        "--citations-policy",
        choices=["preserve-editable", "flatten", "none"],
        required=True,
    )
    native.add_argument("--method", required=True)
    native.add_argument("--note", action="append", default=[])
    native.set_defaults(handler=command_native_policy_review)

    application = sub.add_parser(
        "application-review", help="Derive application-open/export evidence from a live manifest"
    )
    application.add_argument("--preflight", type=Path, required=True)
    application.add_argument("--renderer-manifest", type=Path, required=True)
    application.add_argument("--output", type=Path, required=True)
    application.add_argument("--method", required=True)
    application.add_argument("--note", action="append", default=[])
    application.set_defaults(handler=command_application_review)

    qa = sub.add_parser("qa-review", help="Live-validate all evidence and derive release gates")
    qa.add_argument("--preflight", type=Path, required=True)
    qa.add_argument("--handoff", type=Path)
    qa.add_argument("--render-evidence", type=Path, required=True)
    qa.add_argument("--scope-review", type=Path, required=True)
    qa.add_argument("--structure-review", type=Path, required=True)
    qa.add_argument("--native-policy-review", type=Path, required=True)
    qa.add_argument("--application-review", type=Path, required=True)
    qa.add_argument("--operation-completion", type=Path, required=True)
    qa.add_argument("--delivery-dir", type=Path, required=True)
    qa.add_argument("--deliverable", type=Path, action="append", required=True)
    qa.add_argument("--output", type=Path, required=True)
    qa.set_defaults(handler=command_qa_review)

    release = sub.add_parser("release", help="Issue RELEASED only from live-valid QA evidence")
    release.add_argument("--final", type=Path, required=True)
    release.add_argument("--qa-review", type=Path, required=True)
    release.add_argument("--output", type=Path, required=True)
    release.set_defaults(handler=command_release)

    deliver = sub.add_parser(
        "deliver", help="Copy a live RELEASED file beside the input-at-start document"
    )
    deliver.add_argument("--release-record", type=Path, required=True)
    deliver.add_argument("--input-at-start", type=Path, required=True)
    deliver.add_argument("--destination", type=Path, required=True)
    deliver.add_argument("--output", type=Path, required=True)
    deliver.set_defaults(handler=command_deliver)

    verify_delivery = sub.add_parser(
        "verify-delivery", help="Live-validate a same-directory delivery record"
    )
    verify_delivery.add_argument("--delivery-record", type=Path, required=True)
    verify_delivery.set_defaults(handler=command_verify_delivery)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        return args.handler(args)
    except (
        OSError,
        ValueError,
        json.JSONDecodeError,
        KeyError,
        TypeError,
        AttributeError,
    ) as exc:
        emit({"status": "ERROR", "error": str(exc)}, error=True)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
