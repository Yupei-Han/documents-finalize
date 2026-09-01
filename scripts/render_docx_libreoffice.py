#!/usr/bin/env python3
"""Render a DOCX through LibreOffice and emit a hash-bound renderer manifest."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

from document_handoff import check_docx, fingerprint, same_file, sha256, utc_now, write_json


SCHEMA_VERSION = "2.0"
REPAIR_LANGUAGE = re.compile(r"\b(repair|repaired|recover|recovered|corrupt|corrupted)\b", re.I)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--dpi", type=int, default=180)
    args = parser.parse_args()
    try:
        document = args.input.resolve(strict=True)
        output_dir = args.output_dir.resolve()
        manifest = args.manifest.resolve()
        if args.dpi < 72 or args.dpi > 600:
            raise ValueError("DPI must be between 72 and 600")
        if manifest.suffix.lower() != ".json":
            raise ValueError("renderer manifest must use a .json extension")
        if manifest.exists():
            raise ValueError(f"renderer manifest already exists: {manifest}")
        if same_file(document, manifest):
            raise ValueError("renderer manifest must be separate from the document")
        if manifest.parent != output_dir:
            raise ValueError("renderer manifest must be inside --output-dir")
        package_issues = check_docx(document)
        if package_issues:
            raise ValueError("invalid Word package: " + "; ".join(package_issues))
        if output_dir.exists():
            raise ValueError(f"output directory must be new: {output_dir}")
        output_dir.parent.mkdir(parents=True, exist_ok=True)
        render_script = (Path(__file__).resolve().parents[1] / "render_docx.py").resolve(strict=True)
        source_hash_before = sha256(document)
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            raise ValueError("LibreOffice/soffice is not available on PATH")
        version_run = subprocess.run([soffice, "--version"], capture_output=True, text=True, check=False)
        renderer_version = (version_run.stdout or version_run.stderr).strip()
        if version_run.returncode != 0 or not renderer_version:
            raise ValueError("unable to record the LibreOffice version")
        run = subprocess.run(
            [sys.executable, str(render_script), str(document), "--output_dir", str(output_dir),
             "--dpi", str(args.dpi), "--emit_pdf"],
            capture_output=True,
            text=True,
            check=False,
        )
        if run.returncode != 0:
            detail = (run.stderr or run.stdout).strip()
            raise ValueError(f"render_docx.py failed with exit code {run.returncode}: {detail}")
        diagnostics = [
            line.strip()
            for line in ((run.stdout or "") + "\n" + (run.stderr or "")).splitlines()
            if line.strip()
        ]
        repair_lines = [line for line in diagnostics if REPAIR_LANGUAGE.search(line)]
        if repair_lines:
            raise ValueError(
                "LibreOffice diagnostics contain repair/recovery language: "
                + " | ".join(repair_lines)
            )
        source_hash_after = sha256(document)
        if source_hash_before != source_hash_after:
            raise ValueError("document SHA-256 changed during read-only rendering")
        pdf = (output_dir / f"{document.stem}.pdf").resolve(strict=True)
        pattern = re.compile(r"^page-(\d+)\.png$", re.IGNORECASE)
        numbered: dict[int, Path] = {}
        for path in output_dir.iterdir():
            match = pattern.fullmatch(path.name)
            if match and path.is_file():
                page = int(match.group(1))
                if page in numbered:
                    raise ValueError(f"duplicate rendered page number: {page}")
                numbered[page] = path.resolve(strict=True)
        actual = sorted(numbered)
        expected = list(range(1, len(numbered) + 1))
        if not actual or actual != expected:
            raise ValueError(f"rendered page sequence is empty or non-contiguous: {actual}")
        payload = {
            "schema_version": SCHEMA_VERSION,
            "record_type": "documents_renderer_manifest",
            "producer_script": "render_docx_libreoffice.py",
            "created_at_utc": utc_now(),
            "document": fingerprint(document),
            "source_sha256_before": source_hash_before,
            "source_sha256_after": source_hash_after,
            "renderer": {
                "id": "libreoffice",
                "version": renderer_version,
                "executable": fingerprint(Path(soffice)),
            },
            "settings": {"dpi": args.dpi, "markup_mode": "application-default", "include_markup": None},
            "application_open": {
                "status": "PASS",
                "opened_read_only": True,
                "open_mode": "headless-conversion-source-hash-preserved",
                "export_status": "PASS",
                "repair_warning_count": 0,
                "repair_observation": {
                    "status": "NO_DIAGNOSTIC_OBSERVED",
                    "observed_warning_count": 0,
                    "absence_proven": False,
                    "capture_scope": "captured LibreOffice/render_docx stdout and stderr",
                },
                "diagnostics": diagnostics,
            },
            "toolchain": {"render_docx_script": fingerprint(render_script)},
            "output_directory": str(output_dir),
            "pdf": fingerprint(pdf),
            "pages": [{"page": page, "image": fingerprint(numbered[page])} for page in actual],
        }
        write_json(manifest, payload)
        print(json.dumps({"status": "LIBREOFFICE_RENDER_COMPLETE", "manifest": str(manifest),
                          "document_sha256": source_hash_after, "pages": len(actual)}, indent=2))
        return 0
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        print(json.dumps({"status": "ERROR", "error": str(exc)}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
