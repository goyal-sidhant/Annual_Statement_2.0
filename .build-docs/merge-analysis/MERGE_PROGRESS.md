# Merge progress — Annual Statement 2.0

> Read this first after any break, compaction or "continue". The full instruction for
> this session is `.build-docs/merge-analysis/MERGE_ANALYSIS_PROMPT.md` (copied byte for
> byte from the attachment, 34,093 bytes, 538 lines, SHA-256
> `ce093597f1196b920909b889188a3d66a08934c528e9dd37fc25ae989f5a551e`).

## Where we are

Phase 0 (set up and check access) — **done** (21-10-2026). Phase 1 (the two audits) is next.

The two Phase 0 blockers are cleared:

1. NEW had no commits. The author added a README.md on GitHub; `main` now starts at
   `328deb46f01702d3bc6131e68ae0f6150cb8a30b` ("Initial commit", README.md only). My
   branch `claude/pensive-feynman-0mz1aa` is built on top of it.
2. `rebuild_pipeline_FINAL.md` was missing (the author first attached `CLAUDE_TEMPLATE.md`
   by mistake). The author attached the right file on 21-10-2026; it matches its
   fingerprint exactly. `CLAUDE_TEMPLATE.md` is not kept in NEW.

Snapshots analysed (every document names these):

| Label | Repository | Branch | HEAD commit (the snapshot) | Date of HEAD |
|---|---|---|---|---|
| A | goyal-sidhant/AnnualStatement_oldv1 | main | `bf2f1dd56f088d019bed6022428aab9bf0eb423d` | 27-07-2026 |
| B | goyal-sidhant/AnnualStatement_oldv2 | main | `b0e7d0e8853845ce41ba0f0d2df849feed6bc88c` | 09-02-2026 |
| NEW | goyal-sidhant/Annual_Statement_2.0 | `claude/pensive-feynman-0mz1aa` on `main` `328deb4` | — | 21-10-2026 |

Where things are on this computer:

- A checkout: `/home/user/AnnualStatement_oldv1` (full history, not shallow, clean)
- B checkout: `/home/user/AnnualStatement_oldv2` (full history, not shallow, clean)
- NEW checkout: `/home/user/Annual_Statement_2.0`
- Combined scratch repository: `~/merge-scratch/combined` (remotes `A` and `B`)

If this computer was reclaimed: clone A and B again (read only), fetch NEW's branch, rebuild
`~/merge-scratch/combined` (`git init`, `git remote add A <A clone>`, `git remote add B <B
clone>`, fetch both `main`s).

## Exact next action

**Running all remaining items without pauses, as the author asked (23-10-2026 21:44 IST: "continue for all the items"). The author then allowed running workflows simultaneously "if required" (23-10-2026).** Now running in parallel: item 9 (B final) and Phase 2 part 1 (map + comparison sections 3-10). When both finish: Phase 2 final (summary, family tree, assembly, review against audit B's final text), then Phase 3, then the prompt's stop point (wait for the author's answers).

**Resumed on the author's "continue" (21-10-2026 11:30 UTC).** Queue item 1 is done; item 2 is running. Next: resume the next queue item (Workflow with the script below, `resumeFromRunId: wf_02b98075-ec9`, same args), then items 2–9 one at a time.

Phase 1 — audit A, then B, with `.build-docs/toolkit/Project_Audit_FINAL.md` (all 12
sections, every file line by line, `git show` for every commit, self-consistency review),
with the merge adjustments in the prompt's section 7. Outputs:
`.build-docs/merge-analysis/PROJECT_AUDIT_AnnualStatement_oldv1.md` and
`.build-docs/merge-analysis/PROJECT_AUDIT_AnnualStatement_oldv2.md`.

### Phase 1 working set (scratch, not committed)

- Scratch copies of each HEAD for reading and running: `~/merge-scratch/A-copy`, `~/merge-scratch/B-copy`
  (made with `git archive` from the combined repository, so line numbers match the repositories).
- Python 3.14.6 environments: `~/merge-scratch/venv-A` (tkinter, openpyxl 3.1.5, lxml, pytest — A's
  `requirements.txt`; pywin32 is Windows-only and not installable here) and `~/merge-scratch/venv-B`
  (PyQt5, openpyxl, lxml, pytest).
- A's test suite, first run (RAN, 21-10-2026): 288 tests collected. With a virtual display
  (`xvfb-run`): **1 failed, 287 passed**. Without a display: **14 failed, 241 passed, 33 skipped**.
  The audit explains each failure and skip.
- Notes from the readers: `~/merge-scratch/notes/<id>.md` — A-R1..A-R4 (A's code), A-X-DEAD, A-X-EXCEL,
  A-X-EXT, A-T (tests), A-D (older documents), A-TRACE (live traces); H-S (12 shared commits), H-A1..H-A3
  (A's 34 commits), H-B1..H-B2 (B's 5 commits); B-R1..B-R4 (B's code), B-D, B-X-EXCEL, B-X-EXT, B-X-DEAD,
  B-TRACE. If this computer is reclaimed before the audits are committed, these must be gathered again.

### Phase 1 writing runs (to resume after an interruption)

The audits are written by the workflow script
`/root/.claude/projects/-home-user/3434dc8c-237a-58f7-ab88-3f86af826ad2/workflows/scripts/merge-audit-write-wf_969e1fc4-64f.js`
(modes: `registry`, `sections`, `verifyShared`, `final`). Section drafts land in
`~/merge-scratch/audit/<A|B>/<section>.md`; the assembled audits in
`~/merge-scratch/audit/PROJECT_AUDIT_<name>.md`, then copied into this folder.

**The author's instructions (21-10-2026): run ONE workflow at a time; start the next only when
the previous one has finished.** Then, being short of session limit, the author asked to let only
A-S3a and A-S67 finish and to **pause everything else until the next session window**. Both
finished at about 07:15 UTC; the run was stopped at 07:17 UTC. **Nothing is running. Resume only
when the author says "continue".**

| # | What | Status | Run ID (resume with `resumeFromRunId`) |
|---|---|---|---|
| — | Registry A (31 features, 157 rules) | done | wf_d4384abb-add |
| — | Registry B (34 features, 144 rules) | done | wf_969e1fc4-64f |
| 1 | A sections S3a, S67, S12BF, X345 | **done** — written and adversarially checked (117 corrections in place), 21-10-2026 13:28 UTC | wf_02b98075-ec9 |
| 2 | A sections S3b, S45, S89 | **done** — written and checked, 21-10-2026 17:15 UTC | wf_ca6e1fb9-378; wf_b9ff97f3-7ab |
| 3 | A sections S2, S10, S12own, X12 | **done** — written and checked, 21-10-2026 | wf_d2f12f73-454 |
| 4 | Check of shared 12A (12 shared commits) | **done** — 13 corrections in place, 21-10-2026 | wf_ba37bbb8-fd1 |
| 5 | A final → commit audit A | S1 (title + executive summary) written; a usage limit (about 19:40 UTC, 21-10) stopped S11 and everything after; finishing with mode finalRest (S11 finish → check S11 and S1 → assemble + self-consistency → completeness) from 22-10-2026 01:30 UTC | wf_3971ed50-2f7 (S1); see next line |
| 5b | A final, finishing run | S11 written + checked; S1 checked; assembled; self-consistency review done (58 edits) — the computer restarted (about 02:45 UTC, 22-10) before the reviewer wrote its record, so the record was compiled from its log; completeness check next | wf_b46ac854-45c |
| 5c | A completeness check | **done** — 7 Excel touchpoints, 6 GST-neutrality rewordings, ~40 term explanations and 20 ⚠ marks added; audit A filed as `PROJECT_AUDIT_AnnualStatement_oldv1.md` (15,699 lines), 22-10-2026 | wf_5ac603f8-42c |
| 6 | B sections S3a, S67, S12BF, X345 | **done** — written and checked, 23-10-2026 | wf_c6d2eb54-2d9 |
| 7 | B sections S3b, S45, S89 | **done** — written and checked, 23-10-2026 | wf_9ea2a2b2-4e9 |
| 8 | B sections S2, S10, S12own, X12 | **done** — written and checked, 23-10-2026 16:03 UTC; paused here as asked | wf_2ce78fc9-5d2 |
| 9 | B final → file audit B | **done** — S11 and S1 written and checked; assembled; self-consistency review (58 corrections); completeness check (B's openpyxl path also drops drawn shapes, RAN); filed as `PROJECT_AUDIT_AnnualStatement_oldv2.md` (13,517 lines), 23-10-2026 | wf_05224ce4-4c2 |
| 10a | Phase 2 part 1: MAP + sections C3, C4, C5a, C5b, C6, C78, C910 | **done** — all written and checked, 23-10-2026 | wf_d430da19-c9c, wf_729e1365-a4a, wf_b77e6251-920, wf_1cc2bea7-bdb, wf_9632ef73-e0b |
| 10b | Phase 2 final: summary, family tree, assembly, review | **done** — `MERGE_COMPARISON.md` filed (3,844 lines), 23-10-2026; its review also corrected 4 small facts in the two audits (marked "[Corrected after the comparison review ...]") | wf_76948df9-a0c |
| 11a | Phase 3 drafting: D1 Part 1 (22 items) + D5 Part 11 | **done** — drafted and checked | wf_40ef9b99-b2c |
| 11b | Phase 3 drafting: D2 Parts 2+6 + D3 Parts 3+4+5 | running | wf_3b4e8d5a-cd5 |
| 11c | Phase 3 drafting: D4 Parts 7-10 | running | wf_478487dd-829 |
| 11d | Phase 3 assembly (renumber, closing list, coverage, checks) + final check → file MERGE_DECISIONS.md and MERGE_COMPARISON.md → post list + answer sheet → STOP | waits for 10b and 11a-c | |

Drafts so far (also copied to `working-notes/drafts/`): A — S3a.md (features A-F1..A-F15, 139 KB),
S67.md (Sections 6–7, gaps A-G1.., 208 KB), S10.md (partial); B — S3a.md, S3b.md, S45.md, S10.md
(all partial). A re-run reads an existing draft and finishes only what is missing.

Feature split for Section 3: A-F1..A-F15 / A-F16..A-F31; B-F1..B-F17 / B-F18..B-F34.

## Phase status

| Phase | Status | Date | Notes |
|---|---|---|---|
| 0 — Set up and check access | Done | 21-10-2026 | Blockers cleared the same day |
| 1 — Audit A and B | **Done** — both audits filed: A on 22-10-2026, B on 23-10-2026 | 21-23/10/2026 | `PROJECT_AUDIT_AnnualStatement_oldv1.md`, `PROJECT_AUDIT_AnnualStatement_oldv2.md` in this folder |
| 2 — Compare A and B | **Done** | 23-10-2026 | `MERGE_COMPARISON.md` in this folder |
| 3 — Numbered decision list | In progress (parts drafted in parallel; assembly next) | 23-10-2026 | |
| 4 — Taking the author's answers | Not started | | |
| 5 — Starting files | Not started | | |
| 6 — Final checks and hand-over | Not started | | |

## Facts verified

Checked on 21-10-2026 against the clones (read, not run, except where a git command is
named). "✓" = confirmed as stated; "⚠" = differs or needs a note.

| # | Fact in the prompt | Result | Evidence |
|---|---|---|---|
| 1 | A: 46 commits, last 27-07-2026, Tkinter. B: 17 commits, last 09-02-2026, "v4.0" PyQt5 rewrite | ✓ | `git rev-list --count`: A 46, B 17. A HEAD `bf2f1dd` 27-07-2026 07:48 +0530; B HEAD `b0e7d0e` 09-02-2026 08:42 +0530. A imports only `tkinter` (18 × `import tkinter as tk`); B imports only `PyQt5`. B calls itself v4.0 (`gui/__init__.py:11` `__version__ = '4.0.0'`, `main.py:42`, `gui/pyqt_main_window.py:71`). A's `requirements.txt` line 1 still says "v3.0" |
| 2 | Shared start `edac6a3` (31-08-2025); 12 shared commits; split after `47a1492` (30-10-2025); A then 34 (2 in March 2026, 32 in July 2026), B then 5 (07-02-2026 to 09-02-2026); one branch `main` each | ✓ (one note) | `git merge-base A/main B/main` = `47a1492`; `git rev-list --count 47a1492` = 12; A-only 34 (2 × 2026-03, 32 × 2026-07); B-only 5 (3 on 07-02, 1 on 08-02, 1 on 09-02-2026). Both root commits = `edac6a3`. `git ls-remote` shows only `refs/heads/main` on GitHub for each. **Note:** the clones also hold a local branch `claude/pensive-feynman-0mz1aa` and a matching local tracking ref, made by this session's setup before I started — they exist only on this computer, not on GitHub, and I have not touched them. One of the 12 shared commits is a merge (`9c56a1c`) |
| 3 | No third project; `9c56a1c` mentions "AnnualStatement", A's earlier name, which GitHub forwards to oldv1 | ✓ today, ⚠ for 2025 | Today `git ls-remote https://github.com/goyal-sidhant/AnnualStatement` returns A's HEAD `bf2f1dd`, so the name forwards to A. **But** GitHub's records say B's repository was created on 31-08-2025 (16:38 UTC, minutes before the first commit) and A's on 07-02-2026 — the same day as B's v4.0 commits. A repository's creation date does not change when it is renamed, so on 29-10-2025, when `9c56a1c` pulled from "AnnualStatement", that name most likely belonged to the repository now called **oldv2 (B)**. A probably took the name later, which is why it forwards to A now. Confidence MEDIUM–HIGH. Either way there is no third project: both share the same 12 commits. How to check: you will remember whether, in February 2026, you renamed the original repository and created a new one |
| 4 | Same four screens; extractor was separate (`launch_extractor.py`) before the split; A added a "Step 4: Extract" tab (July 2026), B an extraction view (February 2026) | ✓ (one note) | A: `gui/main_window.py` builds `SetupTab`, `ValidationTab`, `ProcessingTab`, `ExtractTab` (lines ~141–146); "Step 4: Extract" added in `bfcb24a` (26-07-2026). B: `gui/pyqt_main_window.py:90-93` tabs "Step 1: Setup", "Step 2: Validation", "Step 3: Processing", "Step 4: Extract && Refresh" (shown as "Extract & Refresh"), from `c258047` (07-02-2026). **Note:** `launch_extractor.py` still exists in both: in A it still opens the extractor as its own Tkinter window; in B it opens the main window already switched to Tab 4 |
| 5 | Documents in A: `PROJECT_AUDIT_ANNUALSTATEMENT.md` (March 2026), the design note, `tests/README.md`, `.claude/settings.local.json` | ✓ (one note) | Audit added `fc279af` 04-03-2026, revised `f7eb8f1` 05-03-2026; design note `b276542` 25-07-2026; `tests/README.md` `6c1e1c6` 23-07-2026. **Note:** `.claude/settings.local.json` is older than A's own commits — it came in with shared commit `bd176bf` (30-10-2025); B later deleted it. No HANDOVER.md or FILE_INDEX.md exists now or anywhere in either history |
| 6 | A has pytest (about 180 test functions), B none; A's `requirements.txt` makes pywin32 compulsory on Windows (July 2026 commit, so reports cannot lose Power Query); in B that line is commented out | ✓ | 183 `def test_` functions in 18 files under `tests/` (counted with `git grep`; the number pytest collects may differ — Phase 1 runs it). B has no test file in its tree or history after the split. A `requirements.txt` line 9: `pywin32>=305; sys_platform == 'win32'`; commit `be6418c` (26-07-2026) "Require pywin32 on Windows so reports cannot lose Power Query". B `requirements.txt` last line: `# pywin32>=306` |
| 7 | No Excel template, sample input, CSV, JSON data or PDF ever committed; only data-like file is `gst_organizer.log` (added in the first commit, deleted in the second, same day), which appears to hold client business names inside file names; both old repositories public on 21-10-2026 | ✓ (one addition) | Every file ever added in either history: 82 `.py`, 70 `.pyc`, 4 `.md`, 1 `.txt`, 1 `.ini`, 1 `.json` (`.claude/settings.local.json`), 1 `.gitignore`, 1 `.log`. No `.xlsx/.xls/.csv/.pdf`. Log added `edac6a3` 22:11 and deleted `cbfa4a5` 22:14 on 31-08-2025; 5,689 lines, all Python `logging` INFO lines (mostly from `core.file_parser`), dated 23-07-2025 to 31-08-2025; 2,852 lines name files — consistent with client names inside file names (content not read further, nothing quoted). 0 GSTIN-like and 0 PAN-like strings in it. Both repositories `"private": false` on GitHub. **Addition:** the first commit also held 70 compiled Python files (`__pycache__/*.pyc`, removed in `01862a0`, 29-10-2025). Compiled files carry the folder path of the computer they were made on; four such paths are inside them. They look like your own working-folder path, not client data, but I have not quoted them anywhere. One compiled file (`gui/__pycache__/app.*.pyc`) belongs to a `gui/app.py` that was never committed |
| 8 | NEW empty (no commits) and public | ✓ | GitHub lists no branches; `git ls-remote` empty; repository created 21-10-2026 02:02 UTC, `"private": false`. Later the same day the author added README.md, so `main` now exists (`328deb4`) |

### Toolkit files saved — fingerprint check

All five copied from the attachments on this computer (none retyped), then checked.

| File | Saved as | Bytes | Lines | SHA-256 matches the prompt? |
|---|---|---|---|---|
| `rebuild_pipeline_FINAL.md` | `.build-docs/toolkit/rebuild_pipeline_FINAL.md` | 76,645 | 1,594 | ✓ `a430b32c…03be` (attached second, after a wrong file) |
| `Project_Audit_FINAL.md` | `.build-docs/toolkit/Project_Audit_FINAL.md` | 22,846 | 259 | ✓ `7bfcc6a9…0770` |
| `REBUILD_BRIEF_TEMPLATE.md` | `.build-docs/toolkit/REBUILD_BRIEF_TEMPLATE.md` | 4,463 | 127 | ✓ `821a46f9…565a` |
| `REBUILD_BRIEF_ASSISTANT.md` | `.build-docs/toolkit/REBUILD_BRIEF_ASSISTANT.md` | 13,124 | 310 | ✓ `4db4a1cc…15d7` |
| `DECISIONS_LOG_TEMPLATE.md` | `.build-docs/toolkit/DECISIONS_LOG_TEMPLATE.md` | 3,478 | 83 | ✓ `0e486f67…9930` |

The CLAUDE.md copy inside `rebuild_pipeline_FINAL.md` was checked too (needed in Phase 5):
lines 189–1062 = 874 lines, 47,951 bytes, SHA-256 `73c8674d…42d9` ✓; its part from
`# HOW WE BUILD — MANDATORY FOR EVERY PROJECT` to the end = 822 lines, 45,352 bytes,
SHA-256 `4158e750…c471` ✓. Line 189 is `# Project: [PROJECT_NAME] — Rebuild`, line 1062 is
`---`, line 1063 is the closing fence, as the prompt says.

The wrongly attached `CLAUDE_TEMPLATE.md` (42,719 bytes, 752 lines, SHA-256
`d13bc99e…8ba0`) is a different version of the CLAUDE.md template — first line
`# Project: [PROJECT_NAME]`, with new-project lines in Reference Documents and shorter
wording in Standards 2–4. Its Standards 12 and 13 read the same as the pipeline's copy. It
was removed from NEW before the first commit.

### Phase 0 step 3 — the two current versions compared (summary)

`git diff --stat bf2f1dd b0e7d0e`: 82 files differ, 4,142 lines added and 8,797 removed
going from A to B. Of the 22 file paths both heads share, 7 are byte-identical
(`README.md`, `core/__init__.py`, `power_query_extractor/__init__.py`,
`power_query_extractor/config/__init__.py`, `power_query_extractor/config/cell_mappings.py`,
`power_query_extractor/core/__init__.py`, `power_query_extractor/gui/__init__.py`) — so the
cell mappings have not changed on either side since the split. Changed on both sides since
the split: `.gitignore`, `core/excel_handler.py`, `core/file_organizer.py`,
`core/file_parser.py`, `main.py`, `power_query_extractor/core/report_processor.py`,
`power_query_extractor/gui/extractor_window.py`, `requirements.txt`, `utils/helpers.py`,
plus GUI files that B deleted and A edited. File-by-file detail goes into the audits and
MERGE_COMPARISON.md.

### Carried forward from audit A (for audit B, the comparison and the decision list)

- New RAN finding in A's completeness check: openpyxl 3.1.5 also drops drawn shapes (text boxes,
  buttons) when it saves a workbook. Check B's openpyxl path for the same; ask the author whether the
  templates contain shapes (not added to A's Appendix X6, whose count of 88 is quoted elsewhere).
- Six GST-neutrality rewordings were made in A (7.7; 12D for 5c8ce95 and 33d0baf; "a correctly named
  .xlsm" in A-F7 and A-G37). Use the same neutral wording in B and in the comparison. The shared 12A
  text still says a hyphenated client name is "never read correctly" (a technical statement), and
  is identical in both audits.
- Appendix X1 of A now has 86 workbook touchpoints (rows 79-86 added).
- Possible client name to ask the author about: `PROJECT_AUDIT_ANNUALSTATEMENT.md` (A) lines 130-232
  use one company-like name in worked examples; treated as real, never copied.
