# Document workspace layout

Use three separate directories for every document job:

```text
<job-root>/
|-- working/   Working baselines and WORKING candidates
|-- qa/        Risk reports, handoffs, comparisons, renders, and QA records
`-- delivery/  User-facing deliverables only
```

- Keep the immutable source outside these directories or treat its existing location as read-only.
- `documents-fast` writes candidates only to `working/`. Put a handoff in `qa/` when finalization is planned; never place QA evidence in `delivery/`.
- `documents-finalize` writes intermediate candidates and repairs to `working/`, all evidence to `qa/`, and the exact released DOCX to a new or empty `delivery/` directory.
- Every output file is create-new. Do not reuse a prior QA directory, render directory, report, candidate, or release path.
- Use absolute paths. In examples, `${WORK_DIR}` means an absolute writable directory. In PowerShell set `$WORK_DIR = 'D:\path\to\job'`; in POSIX shells set `WORK_DIR=/path/to/job`.
- In examples, `python` means the Python executable returned by the workspace dependency loader.
