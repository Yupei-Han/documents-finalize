# DOCX release protocol

Read this file for every release attempt. Use absolute paths, new outputs, and a fresh `qa/` subdirectory. Run `--help` on each command before use.

## Command sequence

### 1. Bind and create the proposed final

Before copying the accepted candidate, create a persistent receipt bound to its hash and the exact planned final path. Copy the candidate bytes without modification, then create the completion receipt and confirm the candidate and final SHA-256 hashes match.

~~~text
node container_tools/mark_artifact_operation_started.mjs create --operation-kind edit --expected-output-count 1 --output-format docx --receipt <operation-start.json> --input <candidate.docx> --planned-output <release.docx>
# Create <release.docx> only after the command above succeeds.
node container_tools/mark_artifact_operation_started.mjs verify --receipt <operation-start.json> --require-outputs --verification <operation-complete.json>
~~~

Both receipts are create-new. The completion record rechecks the candidate, start receipt, marker script, planned path, and final output hash. Any content or package change belongs in `$documents-fast`; after repair, restart this sequence with new paths.

### 2. Create the release preflight

~~~text
python scripts/document_preflight.py --source <source.docx> --candidate <candidate.docx> --final <release.docx> --output-dir <new-qa-preflight-dir>
~~~

For a net-new document, replace `--source <source.docx>` with `--new-document`.

### 3. Render the exact final bytes

Use one manifest-producing wrapper. The render directory must be new, and the manifest must be inside it.

~~~text
python scripts/render_docx_libreoffice.py --input <release.docx> --output-dir <new-render-dir> --manifest <new-render-dir>/renderer.json
~~~

Use `render_docx_word.ps1` instead when the preflight shows active Word-native features or Word is the layout authority. Set `-IncludeMarkup` only when the authorized release view includes tracked markup.

~~~powershell
& "C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -File scripts\render_docx_word.ps1 `
  -InputPath "<release.docx>" -OutputDir "<new-render-dir>" `
  -ManifestPath "<new-render-dir>\renderer.json" -PdfToPpmPath "<pdftoppm.exe>"
~~~

### 4. Inspect and bind every rendered page

~~~text
python scripts/document_release.py page-review --document <release.docx> --renderer-manifest <renderer.json> --output <page-review.json> --page-result "1=PASS: checked at readable resolution" --page-result "2=PASS: checked at readable resolution" --method "<review method>"
python scripts/document_release.py render-evidence --document <release.docx> --renderer-manifest <renderer.json> --page-review <page-review.json> --output <render-evidence.json>
~~~

Record exactly one result for every rendered page. Do not omit or sample pages.

### 5. Create the four typed review records

Every structured result uses `NAME=PASS: evidence`.

~~~text
python scripts/document_release.py scope-review --preflight <preflight.json> --output <scope.json> --authorized-mode "<mode>" --method "<method>" --requirement-result "request=PASS: verified"
python scripts/document_release.py structure-review --preflight <preflight.json> --output <structure.json> --method "<method>"
python scripts/document_release.py native-policy-review --preflight <preflight.json> --output <native.json> --comments-policy preserve --revisions-policy preserve --fields-policy preserve-editable --citations-policy none --method "<method>"
python scripts/document_release.py application-review --preflight <preflight.json> --renderer-manifest <renderer.json> --output <application.json> --method "<method>"
~~~

Pass every baseline-to-final changed package path once with `--reviewed-part`. Explain every baseline-to-final risk delta once with `--delta-explanation "counts.tables=PASS: expected reason"`. When visible content changed, add one or more `--text-change-review "content=PASS: semantic review evidence"` entries.

### 6. Derive the gates and release

~~~text
python scripts/document_release.py qa-review --preflight <preflight.json> --operation-completion <operation-complete.json> --render-evidence <render-evidence.json> --scope-review <scope.json> --structure-review <structure.json> --native-policy-review <native.json> --application-review <application.json> --delivery-dir <delivery-dir> --deliverable <release.docx> --output <qa-ready.json>
python scripts/document_release.py release --final <release.docx> --qa-review <qa-ready.json> --output <release-record.json>
~~~

Add `--handoff <working.handoff.json>` whenever a fast-stage handoff exists and for every net-new document. Any nonzero command, stale hash, missing or failed page, incomplete changed-part coverage, unsupported policy state, missing Word authority, extra delivery file, or evidence inside `delivery/` blocks release.

### 7. Deliver beside the input-at-start document

The internal `delivery/` payload remains the certified release. Create the user-facing copy only after release, at a new path in the directory that contained the input document when the task began. `--input-at-start` records that original path and may refer to a file that was moved later; its parent directory must still exist.

~~~text
python scripts/document_release.py deliver --release-record <release-record.json> --input-at-start <original-input.docx> --destination <original-input-dir>/<released.docx> --output <delivery-record.json>
python scripts/document_release.py verify-delivery --delivery-record <delivery-record.json>
~~~

`deliver` live-validates the complete release chain, rejects an existing destination, copies the certified bytes without modification, requires matching SHA-256 and size, and writes a create-new delivery record under `qa/`, outside the internal `delivery/` directory. A failed or stale delivery is `NOT DELIVERED`; do not link it as the user-facing release. Exact byte identity means the existing application and page QA remain applicable without rerendering.

## Eight release gates

Every top-level gate applies and must be `PASS` with typed, hash-bound evidence. A feature-level subcheck may be `N/A` only after absence or irrelevance is evidenced.

### 1. Identity and provenance (`identity_provenance`)

- Source, candidate, final, and evidence paths are explicit and separate.
- Current SHA-256 hashes are recorded and rechecked live.
- Start and completion receipts bind the candidate, planned release path, marker script, and exact final hash.
- A fresh isolated preflight binds risk scans and the package/content comparison.
- Any handoff matches the candidate; a net-new document has a valid typed task brief.
- The authorized operation and protected elements are recorded.

### 2. Package integrity (`package_integrity`)

- The ZIP opens and every member passes CRC/read testing.
- Required content types, relationships, and main document parts exist.
- XML parts are well formed; internal targets and content-type overrides resolve.
- The preflight directory contains only its bound records.

### 3. Authorized scope and content (`authorized_scope_content`)

- Every requested change is present, and machine-derived visible-content and package deltas were reviewed.
- No unrelated text, claim, number, unit, citation marker, equation, label, formatting, metadata, relationship, or structure changed.
- Omission, correction, verification, and drafting stayed within the authorized mode.
- Unapproved changes and unresolved material issues both equal zero.

### 4. Protected structures (`protected_structures`)

Review every candidate-to-final changed, added, and removed package part. Explain every risk delta in sections, headers/footers, styles, numbering, tables, drawings, figures, text boxes, equations, notes, comments, revisions, fields, citations, bookmarks, links, content controls, custom XML, embedded objects, macros, signatures, and protection. Missing coverage and unexplained deltas both equal zero.

### 5. Native-feature policy (`native_features_policy`)

- Comments and revisions follow the authorized preserve/remove/accept/reject policy.
- Fields and Zotero/EndNote or other citations follow the authorized preservation/update policy.
- Supported policy values are checked automatically against fresh candidate/final scans.

### 6. Application behavior (`application_behavior`)

- The renderer manifest records successful read-only open/export, preserved source hash, and zero repair/recovery diagnostics observed within an explicit capture scope.
- Fields, links, references, citations, comments, and revisions behave according to policy.
- Microsoft Word validation is used when Word-native structures require it.
- Observed repair warnings and behavior issues equal zero. `absence_proven` remains false because silent behavior cannot be disproved.

### 7. Page-by-page visual QA (`visual_qa`)

- Exact final bytes were rendered after the last edit through a manifest-producing wrapper.
- Renderer identity/version, markup mode, toolchain, PDF, and page hashes are bound.
- Every PNG decodes fully, with exactly one hash-bound `PASS` or `FAIL` for every page.
- Visual defects equal zero.

### 8. Output hygiene (`output_hygiene`)

- `delivery/` contains exactly the declared deliverables and exact final DOCX.
- QA records, renders, risk reports, and release records remain outside `delivery/`.
- The filename follows the request; source and recoverable candidate remain intact.

`release_decision` is derived only after all eight exact keys pass, unresolved material issues equal zero, and the live QA record validates. Caller-supplied gate values, generic evidence flags, and arbitrary notes are not evidence.

## Typed QA evidence

Finalization records use schema version `2.0` unless noted. They bind absolute paths and SHA-256 hashes, refuse existing outputs, and are revalidated against live files.

| Record type | Producer | Binding |
|---|---|---|
| `documents_artifact_operation_start` (1.0) | receipt helper `create` | operation, inputs, exact planned outputs, marker hash, UUID |
| `documents_artifact_operation_completion` (1.0) | receipt helper `verify` | live start receipt, unchanged inputs and marker, exact output hashes |
| `documents_docx_risk_report` (1.0) | preflight / risk scan | package validity, counts, and native-feature flags for one exact DOCX |
| `documents_docx_comparison` (1.0) | preflight / compare | source/candidate/final bytes, package deltas, visible content, fields, relationships |
| `documents_preflight_bundle` (1.0) | preflight | isolated risk, comparison, and binding records |
| `documents_renderer_manifest` | render wrapper | DOCX hashes, renderer, settings, open/export status, diagnostics scope, toolchain, PDF, pages |
| `documents_page_review` | release `page-review` | one reviewer result bound to every page hash |
| `documents_render_evidence` | release `render-evidence` | live manifest, decoded PNGs, exact all-page coverage, zero defects |
| `documents_scope_review` | release `scope-review` | preflight, authorized mode, structured requirements, semantic review |
| `documents_structure_review` | release `structure-review` | every package-part change and risk delta |
| `documents_native_policy_review` | release `native-policy-review` | comments, revisions, fields, and citations policy against fresh scans |
| `documents_application_review` | release `application-review` | live renderer evidence and Word requirement for active native features |
| `documents_qa_review` | release `qa-review` | operation, all QA records, delivery inventory, eight derived gates |
| `documents_release_record` | release `release` | live-valid QA review and exact released DOCX |
| `documents_release_delivery` (1.0) | release `deliver` | live-valid release record, input-at-start directory, and byte-identical same-directory copy |

A net-new document also requires a schema-1.2 `documents_working_handoff` with a live typed task brief and input hashes. `repair_observation.status = NO_DIAGNOSTIC_OBSERVED` means only that no diagnostics appeared within the stated capture scope; `absence_proven` remains false. Manual semantic and visual judgments remain explicit reviewer assertions.
