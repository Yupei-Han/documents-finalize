# Documents Finalize

A Codex release gate for DOCX files that require a defensible **FINAL / RELEASED** status.

## What it does

- Independently audits a source document and a separate working candidate.
- Requires fresh provenance, package, scope, native-feature, application, and every-page visual evidence.
- Delivers an exact same-directory copy only after all applicable gates pass.
- Fails closed as `NOT RELEASED` when evidence is missing or a material issue remains.

## When to use it

Use this only for submission-ready, publication-ready, final, or comprehensively verified DOCX deliverables. It does not author or repair content. Start with [documents-fast](https://github.com/Yupei-Han/documents-fast) whenever a new candidate or repair is needed.

## Install

```powershell
git clone https://github.com/Yupei-Han/documents-finalize "$HOME/.agents/skills/documents-finalize"
```

Restart Codex if needed, confirm it with `/skills`, then invoke:

```text
$documents-finalize audit this working DOCX against its source and release it only if every required gate passes.
```

## Required inputs

- The immutable source document.
- A separate `WORKING` candidate, normally produced by `documents-fast`.
- An optional handoff JSON, which is verified rather than trusted.

## Release flow

1. Create an exact proposed-final copy from the candidate.
2. Run fresh preflight and package comparisons.
3. Verify authorized scope and protected Word structures.
4. Render and inspect 100% of pages.
5. Issue release and delivery records only after all eight gates pass.

## Validation

```powershell
python scripts/self_test.py
```

The included `e2e_smoke.py` covers the split fast-to-finalize release boundary for maintainers.

## Safety boundary

A successful save or render is not a release. Any repair returns to `documents-fast`, and finalization restarts with new evidence.
