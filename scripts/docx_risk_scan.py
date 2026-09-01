#!/usr/bin/env python3
"""Inventory preservation-sensitive structures in a DOCX package."""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from path_guard import ensure_new_file
from urllib.parse import unquote
from xml.etree import ElementTree as ET


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def count_tags(xml_parts: list[tuple[str, bytes]], names: set[str]) -> dict[str, int]:
    counts = {name: 0 for name in names}
    for _, data in xml_parts:
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        for element in root.iter():
            name = local_name(element.tag)
            if name in counts:
                counts[name] += 1
    return counts


def classify_field_instruction(text: str) -> str:
    normalized = " ".join(text.upper().split())
    for marker, label in (
        ("CSL_CITATION", "CSL_CITATION"),
        ("ADDIN EN.CITE", "ENDNOTE_CITATION"),
        ("BIBLIOGRAPHY", "BIBLIOGRAPHY"),
        ("PAGEREF", "PAGEREF"),
        ("NUMPAGES", "NUMPAGES"),
        ("HYPERLINK", "HYPERLINK"),
        ("CITATION", "CITATION"),
        ("TOC", "TOC"),
        ("REF", "REF"),
        ("PAGE", "PAGE"),
        ("SEQ", "SEQ"),
    ):
        if marker in normalized:
            return label
    return "OTHER" if normalized else "EMPTY"


def scan(path: Path, include_field_instructions: bool = False) -> dict:
    resolved = path.resolve(strict=True)
    if resolved.suffix.lower() not in {".docx", ".docm", ".dotx", ".dotm"}:
        raise ValueError(f"Unsupported Word package extension: {resolved.suffix}")

    xml_parts: list[tuple[str, bytes]] = []
    members: list[str] = []
    invalid_xml: list[str] = []
    crc_error: str | None = None
    required_parts = {"[Content_Types].xml", "_rels/.rels", "word/document.xml"}

    with zipfile.ZipFile(resolved) as package:
        members = package.namelist()
        crc_error = package.testzip()
        for name in members:
            if name.lower().endswith((".xml", ".rels")):
                data = package.read(name)
                xml_parts.append((name, data))
                try:
                    ET.fromstring(data)
                except ET.ParseError:
                    invalid_xml.append(name)

    member_set = set(members)
    duplicate_members = sorted({name for name in members if members.count(name) > 1})
    broken_relationships: list[dict[str, str]] = []
    missing_content_type_overrides: list[str] = []
    for name, data in xml_parts:
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        if name.lower().endswith(".rels"):
            rel_parent = posixpath.dirname(name)
            base_dir = "" if name == "_rels/.rels" else posixpath.dirname(rel_parent)
            for relationship in root.iter():
                if local_name(relationship.tag) != "Relationship":
                    continue
                attributes = {local_name(key): value for key, value in relationship.attrib.items()}
                if attributes.get("TargetMode", "").lower() == "external":
                    continue
                target = unquote(attributes.get("Target", "")).replace("\\", "/")
                target = target.split("#", 1)[0].split("?", 1)[0]
                if not target:
                    continue
                resolved_target = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join(base_dir, target))
                if resolved_target not in member_set:
                    broken_relationships.append({"relationships_part": name, "target": target, "resolved_target": resolved_target})
        if name == "[Content_Types].xml":
            for element in root.iter():
                if local_name(element.tag) != "Override":
                    continue
                part_name = next((value for key, value in element.attrib.items() if local_name(key) == "PartName"), "")
                normalized = part_name.lstrip("/")
                if normalized and normalized not in member_set:
                    missing_content_type_overrides.append(part_name)

    tag_names = {
        "p", "tbl", "sectPr", "drawing", "pict", "txbxContent", "oMath",
        "oMathPara", "sdt", "bookmarkStart", "hyperlink", "altChunk",
        "ins", "del", "moveFrom", "moveTo", "commentRangeStart",
        "fldSimple", "footnote", "endnote", "Relationship",
    }
    counts = count_tags(xml_parts, tag_names)

    instructions: list[str] = []
    comments = 0
    footnotes = 0
    endnotes = 0
    external_relationships = 0
    field_begins = 0
    field_separates = 0
    field_ends = 0
    document_protection = 0
    for name, data in xml_parts:
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        for element in root.iter():
            lname = local_name(element.tag)
            if lname == "instrText" and element.text:
                instructions.append(element.text)
            elif lname == "comment" and name.lower().endswith("comments.xml"):
                comments += 1
            elif lname == "footnote" and name.lower().endswith("footnotes.xml"):
                raw_id = next((v for k, v in element.attrib.items() if local_name(k) == "id"), "0")
                if int(raw_id) >= 0:
                    footnotes += 1
            elif lname == "endnote" and name.lower().endswith("endnotes.xml"):
                raw_id = next((v for k, v in element.attrib.items() if local_name(k) == "id"), "0")
                if int(raw_id) >= 0:
                    endnotes += 1
            elif lname == "Relationship":
                target_mode = next((v for k, v in element.attrib.items() if local_name(k) == "TargetMode"), "")
                if target_mode.lower() == "external":
                    external_relationships += 1
            elif lname == "fldChar":
                field_type = next((v for k, v in element.attrib.items() if local_name(k) == "fldCharType"), "")
                if field_type == "begin":
                    field_begins += 1
                elif field_type == "separate":
                    field_separates += 1
                elif field_type == "end":
                    field_ends += 1
            elif lname == "documentProtection":
                document_protection += 1

    instruction_text = " ".join(instructions)
    upper_instructions = instruction_text.upper()
    media_parts = [name for name in members if name.lower().startswith("word/media/") and not name.endswith("/")]
    embedding_parts = [name for name in members if name.lower().startswith("word/embeddings/") and not name.endswith("/")]
    custom_xml_parts = [name for name in members if name.lower().startswith("customxml/") and not name.endswith("/")]
    signature_parts = [name for name in members if name.lower().startswith("_xmlsignatures/") and not name.endswith("/")]
    header_parts = [name for name in members if re.fullmatch(r"word/header\d+\.xml", name, flags=re.I)]
    footer_parts = [name for name in members if re.fullmatch(r"word/footer\d+\.xml", name, flags=re.I)]

    revisions = counts["ins"] + counts["del"] + counts["moveFrom"] + counts["moveTo"]
    equations = counts["oMath"] + counts["oMathPara"]
    fields = counts["fldSimple"] + field_begins
    missing_required = sorted(required_parts - member_set)
    feature_flags = {
        "macros": any(name.lower().endswith("vbaproject.bin") for name in members),
        "digital_signatures": bool(signature_parts),
        "document_protection": document_protection > 0,
        "custom_xml": bool(custom_xml_parts),
        "embedded_objects": bool(embedding_parts),
        "drawings_or_legacy_pictures": counts["drawing"] + counts["pict"] > 0,
        "text_boxes": counts["txbxContent"] > 0,
        "equations": equations > 0,
        "fields": fields > 0,
        "zotero_or_csl_citations": "CSL_CITATION" in upper_instructions,
        "endnote_citations": "EN.CITE" in upper_instructions,
        "comments": comments > 0 or "word/comments.xml" in member_set,
        "tracked_changes": revisions > 0,
        "content_controls": counts["sdt"] > 0,
        "footnotes": footnotes > 0,
        "endnotes": endnotes > 0,
        "alt_chunks": counts["altChunk"] > 0,
        "external_relationships": external_relationships > 0,
        "multiple_sections": counts["sectPr"] > 1,
    }

    package_valid = not crc_error and not invalid_xml and not missing_required and not duplicate_members and not broken_relationships and not missing_content_type_overrides
    report = {
        "schema_version": "1.0",
        "record_type": "documents_docx_risk_report",
        "scanner": "documents-split/docx_risk_scan.py",
        "scanned_at_utc": datetime.now(timezone.utc).isoformat(),
        "path": str(resolved),
        "sha256": sha256(resolved),
        "size_bytes": resolved.stat().st_size,
        "package": {
            "valid": package_valid,
            "member_count": len(members),
            "crc_error_member": crc_error,
            "missing_required_parts": missing_required,
            "invalid_xml_parts": invalid_xml,
            "duplicate_members": duplicate_members,
            "broken_internal_relationships": broken_relationships,
            "missing_content_type_overrides": missing_content_type_overrides,
        },
        "counts": {
            "paragraphs": counts["p"],
            "tables": counts["tbl"],
            "sections": counts["sectPr"],
            "drawings": counts["drawing"],
            "legacy_pictures": counts["pict"],
            "text_boxes": counts["txbxContent"],
            "equations": equations,
            "comments": comments,
            "revisions": revisions,
            "content_controls": counts["sdt"],
            "bookmarks": counts["bookmarkStart"],
            "hyperlinks": counts["hyperlink"],
            "fields": fields,
            "field_begins": field_begins,
            "field_separates": field_separates,
            "field_ends": field_ends,
            "footnotes": footnotes,
            "endnotes": endnotes,
            "media_parts": len(media_parts),
            "embedded_object_parts": len(embedding_parts),
            "custom_xml_parts": len(custom_xml_parts),
            "header_parts": len(header_parts),
            "footer_parts": len(footer_parts),
            "external_relationships": external_relationships,
        },
        "feature_flags": feature_flags,
        "field_instruction_fragments": len(instructions),
        "field_markers_detected": sorted({classify_field_instruction(text) for text in instructions if text.strip()}),
    }
    if include_field_instructions:
        report["field_instructions"] = sorted(set(text.strip() for text in instructions if text.strip()))
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path, help="DOCX/DOCM/DOTX/DOTM package to inspect")
    parser.add_argument("--output", type=Path, help="Optional JSON output path")
    parser.add_argument("--compact", action="store_true", help="Emit compact JSON")
    parser.add_argument(
        "--include-field-instructions",
        action="store_true",
        help="Include raw field instructions; omitted by default to reduce sensitive/noisy output",
    )
    args = parser.parse_args()

    try:
        input_path = args.docx.resolve(strict=True)
        output_path = (
            ensure_new_file(args.output, inputs=[input_path], suffixes=[".json"])
            if args.output
            else None
        )
        report = scan(input_path, include_field_instructions=args.include_field_instructions)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        print(json.dumps({"status": "ERROR", "error": str(exc)}, ensure_ascii=True), file=sys.stderr)
        return 2

    rendered = json.dumps(report, ensure_ascii=False, indent=None if args.compact else 2, sort_keys=True)
    if output_path:
        with output_path.open("x", encoding="utf-8", newline="\n") as handle:
            handle.write(rendered + "\n")
    print(json.dumps(report, ensure_ascii=True, indent=None if args.compact else 2, sort_keys=True))
    return 0 if report["package"]["valid"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
