---
name: documents-finalize
description: "Audit and release a FINAL DOCX through independent structural, content, application, and every-page visual QA. Use for release copies, submission/publication readiness, comprehensive verification, or finalizing a documents-fast candidate. Fail closed if any required evidence or material issue remains."
---

# Documents Finalize

Act as an independent, release-only gate. A working copy or handoff is input, not evidence that any check passed.

## Release authority

- Only this skill may describe a document artifact as `FINAL`, `RELEASED`, `submission-ready`, `publication-ready`, or fully checked.
- Grant `RELEASED` only when every applicable gate passes and no unresolved material issue remains.
- Fail closed when evidence is incomplete, identity is uncertain, a check cannot run, or a material defect remains. Report `Status: NOT RELEASED`; never weaken a gate silently.
- Never overwrite or modify the source or working candidate. Keep source, `working/`, `qa/`, and `delivery/` paths separate as defined in `references/workspace-layout.md`.
- This skill does not author or repair document content. Route every repair, direct final-document request without a candidate, and net-new document request through `$documents-fast` to produce a separate `WORKING` candidate. After any repair, discard prior preflight, render, and review evidence and begin a fresh finalization attempt.
- Finalization does not authorize new scientific claims, unsupported factual corrections, substantive omissions, or unrelated redesign.

## Required input and setup

Accept a source plus a working candidate and optional `.handoff.json`. A net-new document requires a `$documents-fast` handoff containing a typed task brief. Validate every supplied handoff with `scripts/document_handoff.py verify`, then independently recheck current paths, hashes, authorization, protected elements, comments policy, and revisions policy.

Before work:

1. Load the workspace document dependencies.
2. Read `references/workspace-layout.md` and `references/release.md`.
3. Record exact absolute input-at-start, source, candidate, QA, render, internal delivery, proposed final, and same-directory delivery paths.
4. Preserve the source and recoverable candidate until release completes.

Native Google Docs release is outside this DOCX gate. Require a native import connector and post-import validation, or return `NOT RELEASED`.

## Create the proposed final bytes

After accepting the candidate for release, create a persistent operation-start receipt exactly once, bound to the candidate hash and a new proposed final path. Create the final DOCX only as an exact byte-for-byte copy of that candidate, then create the completion receipt and confirm the candidate and final SHA-256 hashes match. Any content, package, field, metadata, or layout change is a repair and must return to `$documents-fast`; after it produces a new candidate, start again with a fresh receipt and preflight. Never reuse an output path or QA directory from a failed attempt.

Use the command sequence in `references/release.md`. A successful copy is not release evidence by itself.

## Independent release workflow

### 1. Establish provenance and authorization

- Confirm live source, candidate, handoff, receipt, and final identities. Reject stale or mismatched records.
- Confirm the requested operation and all protected elements. Do not infer ambiguous comments or revisions policy.
- Verify the delivery directory contains only declared deliverables; keep all QA evidence elsewhere.

### 2. Create a fresh preflight

- Run `scripts/document_preflight.py` only after the exact proposed final exists. It must create a new isolated QA directory with fresh source/candidate/final risk reports and a typed package/content comparison.
- Verify ZIP integrity, required parts, content types, relationships, XML well-formedness, and absence of accidental package members.
- Inventory and compare paragraphs, tables, drawings, equations, styles, numbering, themes, sections, headers/footers, notes, comments, revisions, fields, Zotero/EndNote data, bookmarks, hyperlinks, captions, cross-references, content controls, custom XML, embedded objects, macros, protection, and signatures.
- Treat count changes as prompts for targeted investigation, never proof of correctness or failure by themselves.

### 3. Verify authorized scope and protected structures

- Review every changed, added, and removed package part, visible-text digest, relationship digest, and field/citation digest against the authorized request.
- Confirm all requested changes are present and no unrelated content, formatting, metadata, relationship, or structure changed.
- For academic or scientific material, preserve claim scope, evidence relationships, conditions, certainty, numbers, units, symbols, equations, sample names, panel labels, references, and editable citation fields unless a supported material change was authorized.
- Route every discovered defect to `$documents-fast`. Do not patch it from this skill.

### 4. Render and inspect every page

- Use `scripts/render_docx_libreoffice.py` for ordinary DOCX files. Use `scripts/render_docx_word.ps1` when Word-native fields, citations, equations, revisions, comments, content controls, macros, or pagination are material.
- Microsoft Word is authoritative for Word-native behavior. LibreOffice is secondary compatibility evidence unless cross-application compatibility is an acceptance criterion.
- Use a new render directory with its manifest inside it. Each wrapper must preserve the DOCX hash and bind renderer/version/settings, open/export status, captured diagnostics, toolchain, PDF, and page hashes. Zero observed diagnostics is scoped evidence, not proof that silent behavior was impossible.
- Inspect 100% of pages at readable resolution after the last change; do not omit or sample pages. Check clipping, overlap, blank or duplicate pages, headings, lists, tables, figures, captions, equations, notes, headers/footers, page numbers, sections, margins, fonts, and symbols.
- Record exactly one `PASS` or `FAIL` for every page with `document_release.py page-review`, then generate hash-bound render evidence. Any missing, unreadable, failed, or mismatched page blocks release.

### 5. Apply all eight release gates

Use `references/release.md`. Require typed, live-revalidated evidence for:

1. identity and provenance;
2. package integrity;
3. authorized scope and content;
4. protected structures;
5. native-feature policy;
6. application behavior;
7. 100% page visual QA on the exact final bytes;
8. output hygiene and zero unresolved material issues.

Create scope, structure, native-policy, and application reviews from the fresh preflight. `structure-review` must cover every changed package part and risk delta; `scope-review` must record semantic review of visible-text changes; `application-review` derives status from live renderer and risk evidence. Generic notes and caller-supplied gate values are not evidence.

Run `document_release.py qa-review` to revalidate the exact candidate-to-final operation, all typed records, and delivery inventory. Run `document_release.py release` only from a live-valid QA review. Then run `document_release.py deliver` to create and hash-bind an exact copy beside the input-at-start document. Release and delivery evidence are create-new and remain outside `delivery/`.

## Failure and delivery

Return `Status: NOT RELEASED` when any identity, integrity, authority, renderer, page, protected-feature, scope, privacy, or evidence requirement fails. Preserve and link the best recoverable `WORKING` candidate or report, list the failed gates, and state that `$documents-fast` must produce a repaired candidate. Do not use a `_final` filename or issue a release record.

For success, lead with the live-verified same-directory copy and state `Status: RELEASED`, source and final SHA-256 hashes, inspected/total page counts, completed gates, and any non-material limitation. If same-directory delivery fails, report `Status: NOT DELIVERED`; the internal release record remains valid, but do not present an unbound copy as the user-facing release. Keep internal QA records separate and provide them only when requested.

## Included tools

- `container_tools/mark_artifact_operation_started.mjs`: create candidate/planned-output-bound start receipts and exact-output-hash-bound completion receipts.
- `scripts/document_handoff.py`: validate hash- and risk-bound working handoffs; creation is disabled in this skill.
- `scripts/document_preflight.py`, `scripts/docx_compare.py`, and `scripts/docx_risk_scan.py`: create and live-validate fresh package, content, and risk evidence.
- `scripts/render_docx_libreoffice.py`, `scripts/render_docx_word.ps1`, and `render_docx.py`: create manifest-bound release renders.
- `scripts/document_release.py`: create and validate page and typed review evidence, derive the eight gates, issue a release record, and create/live-verify its same-directory delivery record.
- `scripts/path_guard.py`: enforce create-new paths and reject aliases.
- `scripts/self_test.py`: check structure, routing, path safety, manifests, and shared-core drift.
- `scripts/e2e_smoke.py`: exercise positive and negative release behavior across both split skills.

Run command help before use. Stop on the first failed gate; repair belongs to `$documents-fast`, followed by a fresh finalization attempt.
