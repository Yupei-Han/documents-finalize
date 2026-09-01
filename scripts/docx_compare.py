#!/usr/bin/env python3
"""Create and live-validate a hash-bound semantic/package comparison for Word files."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

from path_guard import ensure_new_file, same_file


SCHEMA_VERSION = "1.0"
RECORD_TYPE = "documents_docx_comparison"
PRODUCER = "documents-finalize/docx_compare.py"
WORD_EXTENSIONS = {".docx", ".docm", ".dotx", ".dotm"}
STORY_PART_RE = re.compile(
    r"^word/(?:document|footnotes|endnotes|comments|header\d+|footer\d+)\.xml$",
    re.I,
)
REL_PART_RE = re.compile(r"(?:^|/)_rels/[^/]+\.rels$", re.I)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def fingerprint(path: Path) -> dict:
    resolved = path.resolve(strict=True)
    return {
        "path": str(resolved),
        "sha256": sha256_file(resolved),
        "size_bytes": resolved.stat().st_size,
    }


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _xml_root(data: bytes, part: str) -> ET.Element:
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise ValueError(f"invalid XML part {part}: {exc}") from exc


def _story_projection(parts: dict[str, bytes]) -> tuple[list[dict], list[str]]:
    stories: list[dict] = []
    instructions: list[str] = []
    for name in sorted(parts):
        if not STORY_PART_RE.fullmatch(name):
            continue
        root = _xml_root(parts[name], name)
        paragraphs: list[str] = []
        for paragraph in (node for node in root.iter() if _local(node.tag) == "p"):
            fragments: list[str] = []
            for node in paragraph.iter():
                local = _local(node.tag)
                if local in {"t", "delText", "tab", "br", "cr"}:
                    if local == "tab":
                        fragments.append("\t")
                    elif local in {"br", "cr"}:
                        fragments.append("\n")
                    elif node.text:
                        fragments.append(node.text)
            paragraphs.append("".join(fragments))
        for node in root.iter():
            local = _local(node.tag)
            if local == "instrText" and node.text:
                instructions.append(" ".join(node.text.split()))
            elif local == "fldSimple":
                instr = next(
                    (value for key, value in node.attrib.items() if _local(key) == "instr"),
                    "",
                )
                if instr:
                    instructions.append(" ".join(instr.split()))
        text = "\n".join(paragraphs)
        stories.append(
            {
                "part": name,
                "paragraphs": len(paragraphs),
                "characters": len(text),
                "sha256": sha256_bytes(text.encode("utf-8")),
            }
        )
    return stories, sorted(instructions)


def _relationship_projection(parts: dict[str, bytes]) -> list[dict]:
    projected: list[dict] = []
    for name in sorted(parts):
        if not REL_PART_RE.search(name):
            continue
        root = _xml_root(parts[name], name)
        relationships = []
        for node in root:
            if _local(node.tag) != "Relationship":
                continue
            relationships.append(
                {
                    "id": node.attrib.get("Id", ""),
                    "type": node.attrib.get("Type", ""),
                    "target": node.attrib.get("Target", ""),
                    "target_mode": node.attrib.get("TargetMode", ""),
                }
            )
        projected.append(
            {
                "part": name,
                "relationships": sorted(
                    relationships,
                    key=lambda item: (
                        item["id"],
                        item["type"],
                        item["target"],
                        item["target_mode"],
                    ),
                ),
            }
        )
    return projected


def snapshot(path: Path) -> dict:
    resolved = path.resolve(strict=True)
    if resolved.suffix.lower() not in WORD_EXTENSIONS:
        raise ValueError(f"unsupported Word extension: {resolved.suffix}")
    try:
        with zipfile.ZipFile(resolved, "r") as package:
            bad_member = package.testzip()
            if bad_member:
                raise ValueError(f"CRC failure in {bad_member}")
            names = package.namelist()
            if len(names) != len(set(names)):
                raise ValueError("duplicate package member names")
            parts = {name: package.read(name) for name in names if not name.endswith("/")}
    except zipfile.BadZipFile as exc:
        raise ValueError(f"invalid Word package: {exc}") from exc
    for required in ("[Content_Types].xml", "_rels/.rels", "word/document.xml"):
        if required not in parts:
            raise ValueError(f"missing required package part: {required}")
    stories, field_instructions = _story_projection(parts)
    story_digest_material = json.dumps(stories, ensure_ascii=False, sort_keys=True).encode("utf-8")
    field_digest_material = json.dumps(
        field_instructions, ensure_ascii=False, sort_keys=True
    ).encode("utf-8")
    relationships = _relationship_projection(parts)
    relationship_digest_material = json.dumps(
        relationships, ensure_ascii=False, sort_keys=True
    ).encode("utf-8")
    return {
        "file": fingerprint(resolved),
        "package": {
            "member_count": len(parts),
            "parts": {
                name: {"sha256": sha256_bytes(data), "size_bytes": len(data)}
                for name, data in sorted(parts.items())
            },
        },
        "visible_content": {
            "stories": stories,
            "sha256": sha256_bytes(story_digest_material),
            "paragraphs": sum(item["paragraphs"] for item in stories),
            "characters": sum(item["characters"] for item in stories),
        },
        "field_instructions": {
            "count": len(field_instructions),
            "sha256": sha256_bytes(field_digest_material),
        },
        "relationships": {
            "parts": relationships,
            "sha256": sha256_bytes(relationship_digest_material),
        },
    }


def _changed_story_parts(left: dict, right: dict) -> list[str]:
    left_map = {entry["part"]: entry["sha256"] for entry in left["visible_content"]["stories"]}
    right_map = {entry["part"]: entry["sha256"] for entry in right["visible_content"]["stories"]}
    return sorted(
        name
        for name in set(left_map) | set(right_map)
        if left_map.get(name) != right_map.get(name)
    )


def delta(left: dict, right: dict) -> dict:
    left_parts = left["package"]["parts"]
    right_parts = right["package"]["parts"]
    common = set(left_parts) & set(right_parts)
    changed = sorted(
        name for name in common if left_parts[name]["sha256"] != right_parts[name]["sha256"]
    )
    added = sorted(set(right_parts) - set(left_parts))
    removed = sorted(set(left_parts) - set(right_parts))
    return {
        "left_sha256": left["file"]["sha256"],
        "right_sha256": right["file"]["sha256"],
        "identical_bytes": left["file"]["sha256"] == right["file"]["sha256"],
        "package": {
            "changed_parts": changed,
            "added_parts": added,
            "removed_parts": removed,
            "changed_part_count": len(changed) + len(added) + len(removed),
        },
        "visible_content": {
            "changed": left["visible_content"]["sha256"]
            != right["visible_content"]["sha256"],
            "changed_story_parts": _changed_story_parts(left, right),
            "paragraph_delta": right["visible_content"]["paragraphs"]
            - left["visible_content"]["paragraphs"],
            "character_delta": right["visible_content"]["characters"]
            - left["visible_content"]["characters"],
        },
        "field_instructions": {
            "changed": left["field_instructions"]["sha256"]
            != right["field_instructions"]["sha256"],
            "count_delta": right["field_instructions"]["count"]
            - left["field_instructions"]["count"],
        },
        "relationships": {
            "changed": left["relationships"]["sha256"]
            != right["relationships"]["sha256"],
        },
    }


def build_report(source: Path | None, candidate: Path, final: Path) -> dict:
    candidate_resolved = candidate.resolve(strict=True)
    final_resolved = final.resolve(strict=True)
    source_resolved = source.resolve(strict=True) if source else None
    named = [
        ("source", source_resolved),
        ("candidate", candidate_resolved),
        ("final", final_resolved),
    ]
    concrete = [(name, path) for name, path in named if path is not None]
    for index, (left_name, left) in enumerate(concrete):
        for right_name, right in concrete[index + 1 :]:
            if same_file(left, right):
                raise ValueError(f"{left_name} and {right_name} must be separate files")
    documents = {
        "source": snapshot(source_resolved) if source_resolved else None,
        "candidate": snapshot(candidate_resolved),
        "final": snapshot(final_resolved),
    }
    comparisons = {
        "candidate_to_final": delta(documents["candidate"], documents["final"]),
        "baseline_to_final": delta(
            documents["source"] or documents["candidate"], documents["final"]
        ),
    }
    return {
        "schema_version": SCHEMA_VERSION,
        "record_type": RECORD_TYPE,
        "producer": PRODUCER,
        "created_at_utc": utc_now(),
        "provenance_mode": "source-edit" if source_resolved else "new-document",
        "documents": documents,
        "comparisons": comparisons,
    }


def _stable_payload(payload: dict) -> dict:
    return {key: value for key, value in payload.items() if key != "created_at_utc"}


def validate_report(
    report_path: Path,
    source: Path | None = None,
    candidate: Path | None = None,
    final: Path | None = None,
) -> tuple[dict, list[str]]:
    report = report_path.resolve(strict=True)
    payload = json.loads(report.read_text(encoding="utf-8"))
    issues: list[str] = []
    if payload.get("schema_version") != SCHEMA_VERSION:
        issues.append("comparison report has unsupported schema_version")
    if payload.get("record_type") != RECORD_TYPE:
        issues.append("comparison report has unexpected record_type")
    if payload.get("producer") != PRODUCER:
        issues.append("comparison report has unexpected producer")
    try:
        stored = payload["documents"]
        source_path = source
        if source_path is None and stored.get("source"):
            source_path = Path(stored["source"]["file"]["path"])
        candidate_path = candidate or Path(stored["candidate"]["file"]["path"])
        final_path = final or Path(stored["final"]["file"]["path"])
        live = build_report(source_path, candidate_path, final_path)
        if _stable_payload(payload) != _stable_payload(live):
            issues.append("comparison report differs from a fresh live comparison")
    except (OSError, KeyError, TypeError, ValueError, zipfile.BadZipFile) as exc:
        issues.append(f"comparison report cannot be live-validated: {exc}")
    return payload, issues


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--final", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        inputs = [args.candidate, args.final]
        if args.source:
            inputs.append(args.source)
        output = ensure_new_file(args.output, inputs=inputs, suffixes=[".json"])
        report = build_report(args.source, args.candidate, args.final)
        with output.open("x", encoding="utf-8", newline="\n") as handle:
            json.dump(report, handle, ensure_ascii=False, indent=2, sort_keys=True)
            handle.write("\n")
        print(
            json.dumps(
                {
                    "status": "CREATED",
                    "output": str(output),
                    "candidate_to_final_changed_parts": report["comparisons"][
                        "candidate_to_final"
                    ]["package"]["changed_part_count"],
                    "baseline_visible_content_changed": report["comparisons"][
                        "baseline_to_final"
                    ]["visible_content"]["changed"],
                },
                ensure_ascii=True,
            )
        )
        return 0
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        print(json.dumps({"status": "ERROR", "error": str(exc)}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
