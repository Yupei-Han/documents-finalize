#!/usr/bin/env python3
"""Run dependency-light structural, routing, safety, and peer-drift checks."""

from __future__ import annotations

import argparse
import ast
import importlib.util
import json
import os
import re
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT_DIR = ROOT / "scripts"
TEXT_SUFFIXES = {".md", ".py", ".ps1", ".yaml", ".yml", ".txt"}
REFERENCE_RE = re.compile(
    r"(?<![A-Za-z0-9_])((?:scripts|tasks|references|ooxml|troubleshooting)/[A-Za-z0-9_.\-/]+)"
)
FAST_GUARDED_PYTHON = {
    "a11y_audit.py",
    "accept_tracked_changes.py",
    "apply_template_styles.py",
    "captions_and_crossrefs.py",
    "comments_add.py",
    "comments_apply_patch.py",
    "comments_extract.py",
    "comments_strip.py",
    "content_controls.py",
    "document_handoff.py",
    "docx_ooxml_patch.py",
    "docx_risk_scan.py",
    "docx_table_to_csv.py",
    "fields_materialize.py",
    "flatten_ref_fields.py",
    "google_docs_title_sanitize.py",
    "insert_note.py",
    "insert_ref_fields.py",
    "insert_toc.py",
    "internal_nav.py",
    "make_fixtures.py",
    "merge_docx_append.py",
    "privacy_scrub.py",
    "redact_docx.py",
    "render_and_diff.py",
    "set_protection.py",
    "style_lint.py",
    "style_normalize.py",
    "watermark_audit_remove.py",
    "xlsx_to_docx_table.py",
}
FINALIZE_GUARDED_PYTHON = {
    "document_handoff.py",
    "document_preflight.py",
    "document_release.py",
    "docx_compare.py",
    "docx_risk_scan.py",
}
FINALIZE_SCRIPTS = FINALIZE_GUARDED_PYTHON | {
    "e2e_smoke.py",
    "path_guard.py",
    "render_docx_libreoffice.py",
    "render_docx_word.ps1",
    "self_test.py",
}
SHARED_CORE = {
    "render_docx.py",
    "scripts/document_handoff.py",
    "scripts/docx_risk_scan.py",
    "scripts/path_guard.py",
    "scripts/self_test.py",
}


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def parse_frontmatter(path: Path) -> dict[str, str]:
    lines = read_text(path).splitlines()
    if not lines or lines[0].strip() != "---":
        raise ValueError("SKILL.md lacks opening YAML frontmatter")
    try:
        end = lines.index("---", 1)
    except ValueError as exc:
        raise ValueError("SKILL.md lacks closing YAML frontmatter") from exc
    values: dict[str, str] = {}
    for line in lines[1:end]:
        if not line.strip():
            continue
        if ":" not in line:
            raise ValueError(f"unsupported frontmatter line: {line!r}")
        key, value = line.split(":", 1)
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def check_frontmatter(issues: list[str]) -> None:
    try:
        values = parse_frontmatter(ROOT / "SKILL.md")
    except ValueError as exc:
        issues.append(str(exc))
        return
    if set(values) != {"name", "description"}:
        issues.append(f"frontmatter keys must be exactly name/description: {sorted(values)}")
    if values.get("name") != ROOT.name:
        issues.append(f"frontmatter name does not match directory: {values.get('name')!r}")
    description = values.get("description", "")
    if not description or len(description) > 1024:
        issues.append("frontmatter description is empty or longer than 1024 characters")
    skill_lines = read_text(ROOT / "SKILL.md").splitlines()
    if len(skill_lines) > 500:
        issues.append(f"SKILL.md is too long for progressive disclosure: {len(skill_lines)} lines")


def check_agent_metadata(issues: list[str]) -> None:
    path = ROOT / "agents" / "openai.yaml"
    if not path.is_file():
        issues.append("agents/openai.yaml is missing")
        return
    text = read_text(path)
    for key in ("display_name:", "short_description:", "default_prompt:", "allow_implicit_invocation:"):
        if key not in text:
            issues.append(f"agents/openai.yaml lacks {key}")
    if f"${ROOT.name}" not in text:
        issues.append("default_prompt does not explicitly invoke this skill")
    if "allow_implicit_invocation: true" not in text:
        issues.append("implicit invocation is not enabled")


def check_references(issues: list[str]) -> None:
    for markdown in ROOT.rglob("*.md"):
        text = read_text(markdown)
        for match in REFERENCE_RE.finditer(text):
            relative = match.group(1).rstrip(".,:;)")
            target = ROOT / Path(relative)
            if not target.exists():
                issues.append(
                    f"dangling reference in {markdown.relative_to(ROOT).as_posix()}: {relative}"
                )


def check_python(issues: list[str]) -> None:
    for path in ROOT.rglob("*.py"):
        if "__pycache__" in path.parts:
            continue
        try:
            ast.parse(read_text(path), filename=str(path))
        except SyntaxError as exc:
            issues.append(f"Python syntax error in {path.relative_to(ROOT)}: {exc}")
    guarded = FAST_GUARDED_PYTHON if ROOT.name == "documents-fast" else FINALIZE_GUARDED_PYTHON
    for name in sorted(guarded):
        path = SCRIPT_DIR / name
        if not path.is_file():
            issues.append(f"expected guarded helper is missing: scripts/{name}")
            continue
        text = read_text(path)
        if not any(token in text for token in ("ensure_new_file", "ensure_new_directory", "write_json")):
            issues.append(f"writer lacks shared create-new guard: scripts/{name}")
    root_renderer = read_text(ROOT / "render_docx.py")
    if "ensure_new_directory" not in root_renderer:
        issues.append("render_docx.py lacks create-new output-directory protection")
    if ROOT.name == "documents-finalize":
        for name in (
            "docx_compare.py",
            "document_preflight.py",
            "document_release.py",
            "e2e_smoke.py",
        ):
            if not (SCRIPT_DIR / name).is_file():
                issues.append(f"finalization evidence helper is missing: scripts/{name}")
        if "write_json" not in read_text(SCRIPT_DIR / "document_release.py"):
            issues.append("document_release.py does not use guarded JSON writes")


def check_role_surface(issues: list[str]) -> None:
    if ROOT.name != "documents-finalize":
        return
    actual = {
        path.name
        for path in SCRIPT_DIR.iterdir()
        if path.is_file() and path.suffix.lower() in {".py", ".ps1"}
    }
    if actual != FINALIZE_SCRIPTS:
        issues.append(
            "finalize script surface is not release-only; "
            f"missing={sorted(FINALIZE_SCRIPTS - actual)}, extra={sorted(actual - FINALIZE_SCRIPTS)}"
        )
    for directory in ("tasks", "ooxml", "troubleshooting"):
        path = ROOT / directory
        if path.is_dir() and any(item.is_file() for item in path.rglob("*")):
            issues.append(f"finalize retains authoring procedures: {directory}/")


def check_operation_receipts(issues: list[str]) -> None:
    path = ROOT / "container_tools" / "mark_artifact_operation_started.mjs"
    if not path.is_file():
        issues.append("operation receipt helper is missing")
        return
    text = read_text(path)
    for invariant in (
        "documents_artifact_operation_start",
        "documents_artifact_operation_completion",
        'flag: "wx"',
        "start_receipt:",
        "planned_outputs",
    ):
        if invariant not in text:
            issues.append(f"operation receipt helper lacks invariant: {invariant}")


def check_path_guard(issues: list[str]) -> None:
    spec = importlib.util.spec_from_file_location("documents_path_guard", SCRIPT_DIR / "path_guard.py")
    if spec is None or spec.loader is None:
        issues.append("cannot load path_guard.py")
        return
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    with tempfile.TemporaryDirectory(prefix="documents_skill_selftest_") as temp:
        root = Path(temp)
        input_file = root / "input.docx"
        input_file.write_bytes(b"input")
        try:
            module.ensure_new_file(input_file, inputs=[input_file], suffixes=[".docx"])
            issues.append("path_guard accepted an existing/same-path output")
        except (FileExistsError, ValueError):
            pass
        existing = root / "existing.docx"
        existing.write_bytes(b"existing")
        try:
            module.ensure_new_file(existing, inputs=[input_file], suffixes=[".docx"])
            issues.append("path_guard accepted an existing output")
        except FileExistsError:
            pass
        try:
            module.ensure_new_file(root / "wrong.txt", inputs=[input_file], suffixes=[".docx"])
            issues.append("path_guard accepted a wrong output extension")
        except ValueError:
            pass
        accepted = module.ensure_new_file(
            root / "nested" / "new.docx", inputs=[input_file], suffixes=[".docx"]
        )
        if accepted.name != "new.docx" or not accepted.parent.is_dir():
            issues.append("path_guard did not resolve/create a valid output parent")
        existing_dir = root / "existing-dir"
        existing_dir.mkdir()
        try:
            module.ensure_new_directory(existing_dir)
            issues.append("path_guard accepted an existing output directory")
        except FileExistsError:
            pass
        alias = root / "alias.docx"
        try:
            os.link(input_file, alias)
            if not module.same_file(input_file, alias):
                issues.append("path_guard failed to recognize a hard-link alias")
        except OSError:
            pass


def check_routing(issues: list[str]) -> None:
    skill = read_text(ROOT / "SKILL.md")
    description = parse_frontmatter(ROOT / "SKILL.md").get("description", "")
    if ROOT.name == "documents-fast":
        if "WORKING" not in description or "documents-finalize" not in description:
            issues.append("fast description does not clearly route release requests")
        for phrase in ("Never use `final`", "Status: WORKING", "bounded QA"):
            if phrase not in skill:
                issues.append(f"fast routing invariant is missing: {phrase}")
    elif ROOT.name == "documents-finalize":
        for phrase in (
            "FINAL",
            "Fail closed",
            "Inspect 100% of pages",
            "page-review",
            "Route every repair",
            "fresh finalization attempt",
        ):
            if phrase.lower() not in skill.lower():
                issues.append(f"finalize routing/release invariant is missing: {phrase}")
        if "$documents-fast" not in skill:
            issues.append("finalize does not route repairs to documents-fast")


def distribution_files(root: Path) -> set[str]:
    files: set[str] = set()
    for path in root.rglob("*"):
        if not path.is_file() or "__pycache__" in path.parts or path.suffix == ".pyc":
            continue
        relative = path.relative_to(root).as_posix()
        if relative == "manifest.txt":
            continue
        files.add(relative)
    return files


def check_manifest(issues: list[str]) -> None:
    path = ROOT / "manifest.txt"
    if not path.is_file():
        issues.append("manifest.txt is missing")
        return
    listed = {
        line.strip().replace("\\", "/")
        for line in read_text(path).splitlines()
        if line.strip()
    }
    expected = distribution_files(ROOT)
    if listed != expected:
        missing = sorted(expected - listed)
        stale = sorted(listed - expected)
        issues.append(f"manifest mismatch; missing={missing}, stale={stale}")


def _normalized_handoff(text: str) -> str:
    return re.sub(r"^ALLOW_CREATE\s*=\s*(?:True|False)\s*$", "ALLOW_CREATE = <ROLE>", text, flags=re.M)


def discover_peer() -> Path | None:
    peer_name = {
        "documents-fast": "documents-finalize",
        "documents-finalize": "documents-fast",
    }.get(ROOT.name)
    if peer_name is None:
        return None
    candidate = ROOT.parent / peer_name
    return candidate if candidate.is_dir() else None


def check_peer(peer: Path, issues: list[str]) -> None:
    peer = peer.resolve(strict=True)
    if peer == ROOT:
        issues.append("peer skill must differ from the current skill")
        return
    for relative in sorted(SHARED_CORE):
        here = ROOT / relative
        there = peer / relative
        if not here.is_file() or not there.is_file():
            issues.append(f"shared core is missing: {relative}")
            continue
        if relative == "scripts/document_handoff.py":
            left = _normalized_handoff(read_text(here))
            right = _normalized_handoff(read_text(there))
            if left != right:
                issues.append("shared document_handoff.py drift exceeds ALLOW_CREATE role difference")
        elif here.read_bytes() != there.read_bytes():
            issues.append(f"shared-file drift: {relative}")


def check_portability(issues: list[str]) -> None:
    for path in ROOT.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in TEXT_SUFFIXES or "__pycache__" in path.parts:
            continue
        text = read_text(path)
        if ("/mnt" + "/data") in text:
            issues.append(f"legacy mounted-workspace path remains in {path.relative_to(ROOT)}")
        if "C:\\Users\\Yupei" in text:
            issues.append(f"host-specific user path remains in {path.relative_to(ROOT)}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--peer",
        type=Path,
        help="Override the automatically discovered sibling skill root for shared-file drift checks",
    )
    args = parser.parse_args()
    issues: list[str] = []
    check_frontmatter(issues)
    check_agent_metadata(issues)
    check_references(issues)
    check_python(issues)
    check_role_surface(issues)
    check_operation_receipts(issues)
    check_path_guard(issues)
    check_routing(issues)
    check_manifest(issues)
    check_portability(issues)
    peer = args.peer or discover_peer()
    if peer is None:
        issues.append(
            "peer skill could not be auto-discovered beside the current skill; "
            "install the sibling or pass --peer"
        )
    else:
        check_peer(peer, issues)
    status = "PASS" if not issues else "FAIL"
    print(
        json.dumps(
            {
                "status": status,
                "skill": ROOT.name,
                "checks": [
                    "frontmatter",
                    "agent_metadata",
                    "references",
                    "python_ast",
                    "create_new_guards",
                    "role_surface",
                    "operation_receipts",
                    "routing",
                    "manifest",
                    "portability",
                    "peer_drift" if peer is not None else "peer_drift_missing",
                ],
                "peer": str(peer.resolve()) if peer is not None else None,
                "issues": issues,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0 if not issues else 2


if __name__ == "__main__":
    raise SystemExit(main())
