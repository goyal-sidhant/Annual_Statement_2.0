# Working notes for the merge analysis (temporary)

These files are the raw evidence behind the two audits in `.build-docs/merge-analysis/`.
They are committed only so that hours of read-only analysis survive if the cloud computer
running this session is reclaimed. They are deleted with the rest of `merge-analysis/` in
Phase 6, once everything in them has moved into the audits, the comparison and Session 0.

- `notes/` — notes by 25 read-only reader agents (21-10-2026). Each covers one slice of
  project A (AnnualStatement_oldv1 at `bf2f1dd`) or project B (AnnualStatement_oldv2 at
  `b0e7d0e`): code read line by line, commit diffs, tests, older documents, Excel and Power
  Query safety, outside files, dead code, and pure-Python runs on made-up inputs.
- `A_registry.md`, `B_registry.md` — the feature and rule IDs (A-F#, A-R#, B-F#, B-R#)
  that the audits use.

They are working material, not the audits: where a note and an audit disagree, the audit
(which was checked against the code) wins. No client names, GSTINs, PANs, client file names
or figures are in them. Places in the old repositories that may hold client data are listed
by path and commit only.
