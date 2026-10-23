### 12B. File Evolution Map

This sub-section traces every file of project B — the 41 files at the snapshot `b0e7d0e` (09-02-2026), the 20 further files and 70 compiled files that existed only in the history, and one file that was never committed but left traces — through all 17 commits of B's branch `main`: the 12 commits shared with project A (always marked "shared - before the split on 30-10-2025") and B's own 5 (always marked "B only (after the split)").

> **How to read the evidence marks.** READ = established by reading the code or a diff (the exact lines a commit added and removed). RAN = established by running something. "RAN (me)" means I ran it for this section, on copies extracted with `git archive` into `/root/merge-scratch/run/B-S12BF/` (scripts in `/root/merge-scratch/run/B-S12BF/scripts/`, always `python -B`, so nothing was written into the repositories or into `/root/merge-scratch/B-copy`). "RAN (H-S / H-B1 / H-B2 / …)" means the reader whose working notes are named ran it. Inferences carry HIGH / MEDIUM / LOW confidence. A citation `path:line` is the line at the snapshot `b0e7d0e`; `path:line@commit` is the line in that older commit's version of the file.
>
> **Verification (task VERIFY-B-S12BF, 23-10-2026).** Every `file:line` citation was opened, and the most important RAN claims were re-run independently from fresh `git archive` copies in `/root/merge-scratch/run/VERIFY-B-S12BF/` (scripts in its `scripts/` folder, `python -B`): the size tables (`sizes.py`), the `.pyc` header fields (size and time only), the import chain with and without stand-in pywin32 modules at `47a1492`, `c258047`, `a62d3cc`, `c21f788`, `7e09e5c`, `b0e7d0e`, the dangling imports, the scan trace (`scan_trace.py`, same five-file result), the rule blame (`rule_blame.py`, output identical), the VERIFICATION RULE (`verify_rule.py`), the openpyxl merged-cell fallback (`fallback.py`), the consolidator mismatch (`consol.py`: 9 of 22 columns), the Step 4 disk tree (`disk.py`) and the project-wide outside-import sets (`stdlib_sets.py`). Where these runs confirm a claim it is marked "re-run by the verifier"; corrections are marked "[Corrected during verification …]" or worded in place.
>
> **Labels.** The 12 commits `edac6a3`, `cbfa4a5`, `6237597`, `9c56a1c`, `d674386`, `5ea77c2`, `1cb878b`, `01862a0`, `1c48722`, `4c42758`, `bd176bf`, `47a1492` are always "shared - before the split on 30-10-2025". The 5 commits `c258047`, `a62d3cc`, `c21f788`, `7e09e5c`, `b0e7d0e` are always "B only (after the split)". Where a table cell repeats a commit already labelled nearby without repeating its label, this rule supplies it.

**Terms used in this section (plain English).**
- *commit* = one saved change in the project's history; *diff* = the exact lines a commit added (+) and removed (−); *numstat* = git's count of lines added and removed per file.
- *branch* = one line of work. Project B has one branch, `main`. The combined repository also holds a reference `B/claude/pensive-feynman-0mz1aa`; RAN (me) `git for-each-ref`: it points at `b0e7d0e` itself, so it is not a separate line of work.
- *merge commit* = a commit that joins two lines of work. B has exactly one, `9c56a1c` (29-10-2025, shared - before the split on 30-10-2025), which joined the author's local commit `6237597` with `cbfa4a5`, a deletion made on GitHub's web page 2 minutes 13 seconds after the first commit's time stamp (`edac6a3` 22:11:56 → `cbfa4a5` 22:14:09 on 31-08-2025, RAN `git log --date=iso`; when the upload itself happened is not recorded). Both parents are part of `main`'s history.
- *blame* = git's report of which commit last changed each line of a file; *pickaxe* (`git log -S`) = git's search for the commits that added or removed a given piece of text.
- *rename detection* = git's check whether a file was moved or renamed (it compares content, not names).
- *.pyc / `__pycache__`* = compiled copies of Python files that Python writes by itself; they record the source file's size, last-modified time and the folder it came from.
- *module* = one `.py` file; *package* = a folder of modules with an `__init__.py` file; *import* = one file asking to use another file's code; *re-export* = a package's `__init__.py` importing names so they can be used from the package directly; *eager import* = an `__init__.py` that loads heavy modules the moment anything in the package is touched.
- *regex / pattern* = a text-matching rule for file names; *glob* = a wildcard file search such as `ITC_Report_*.xlsx`; *stat* = asking the operating system for a file's size and dates; *mtime* = a file's or folder's "last modified" time.
- *COM* = Windows' way for one program to drive another (here Python drives Excel); *pywin32* = the Python package that provides COM (`pythoncom`, `win32com`, `win32api`, `win32con`, `win32gui`); *openpyxl* = a Python library that reads and writes `.xlsx` files without Excel and drops Power Query when it saves; *Power Query* = Excel's data-import tool inside the ITC/Sales templates; *Links sheet* = the template sheet into which the program writes folder paths and file names; *SendKeys* = typing keystrokes into whatever window has the keyboard focus.
- *Tk / Tkinter* = Python's built-in window toolkit (used by every shared commit); *PyQt5* = a Python wrapper for the Qt 5 window toolkit (used by B from `c258047`); *widget* = one on-screen element; *view* = one tab's screen code; *controller* = code that runs a job (scan, processing, extraction) and reports back; *QThread / worker thread* = a second line of work so the window does not freeze; *signal / slot* = Qt's "this happened, now run that" wiring; *QSS* = Qt's CSS-like style rules for colours and fonts; *dataclass* = a simple Python container of named fields; *JSON* = a plain-text settings format.
- *Level 1 / Level 2 / Version folder* = B's output tree: `<target>/Annual Statement-<DDMMYY HHMM>/` (the run folder, "Level 1") → `<client>-<state code>/` (the client folder, "Level 2") → `Version-<DDMMYY HHMM>/`; *variations* = the code's word for problem files (unrecognised or invalid).
- *regression* = something that worked and was broken by a later change; *patch* = a fix that removes the symptom; *proper fix* = a fix that removes the cause; *dangling import* = an import of a file that has been deleted, which fails as soon as it runs.
- *trailer* = a tagged line at the end of a commit message (here "Co-Authored-By: …", naming the AI tool that helped); *shebang* = a first line `#!…` that tells Unix-type systems which program runs the file; *CRLF* = Windows-style line endings (two characters per line end), which make a file larger than its line count suggests; *lambda* = a small unnamed function written inline; *late binding* = a lambda made inside a loop reads the loop's variables only when it finally runs, so every copy sees the last loop values; *8.3 short path* = Windows' old short form of a long path (for example `PROGRA~1`); *ETA* = estimated time left; *first-substring look-up* = taking, for each output column, the first extracted value whose field name merely *contains* the column title (B-R83); *exit code 134* = the program was aborted by Qt after an error nobody caught.

**Scripts I ran for 12B (all RAN (me)).** `filehist.sh` → `filehist.txt` (every path that ever existed on `B/main`: each commit that touched it, oldest first, with lines added/removed; whether it exists at `b0e7d0e`); `scripts/sizes.py` → `sizes.txt` and `scripts/sizes_area.py` → `sizes_area.txt` (file and line counts at every commit and per folder); `git log -M --diff-filter=R` and `git log -C -C --diff-filter=C` (renames and copies: both empty); `scripts/verify_imports.py` → `verify_imports.txt` and `scripts/broken_imports_by_commit.py` → `broken_imports_by_commit.txt` (the method's VERIFICATION RULE and the reverse check, 12B.7); import-chain runs on extracted trees of `47a1492`, `c258047`, `a62d3cc`, `c21f788` and `b0e7d0e`, with and without stand-in ("stub") pywin32 modules (12B.7, 12F.5).

#### 12B.1 How the project grew (sizes at every point where the file set changed)

RAN (me) `scripts/sizes.py` (re-run by the verifier with its own script: every figure in both tables identical): counts from `git ls-tree`; line counts are git's own (a last line without a line ending counts as a line; `wc -l` gives one less for such files, which is why the registry's file list shows 7,805 lines at the snapshot where git counts 7,821).

| Point in history | Commit (date) | Files (without `.pyc`) | `.pyc` files | Python files / lines | What else is tracked |
|---|---|---|---|---|---|
| First commit | `edac6a3` (31-08-2025, shared - before the split on 30-10-2025) | 40 | 70 | 37 / 6,954 | README.md, requirements.txt, `gst_organizer.log` |
| Log deleted on GitHub | `cbfa4a5` (31-08-2025, shared) | 39 | 70 | 37 / 6,954 | README.md, requirements.txt |
| First review, merged | `6237597` → `9c56a1c` (29-10-2025, shared) | 40 → 39 | 70 | 37 / 6,968 | `6237597` still holds the log (its parent is `edac6a3`); `9c56a1c` brings in the deletion |
| `.gitignore`, `.pyc` removed | `1cb878b`, `01862a0` (29-10-2025, shared) | 40 | 70 → 0 (`1cb878b` still holds all 70; `01862a0` removes them) | 37 / 6,967 | `.gitignore` added |
| Extractor features | `bd176bf` (30-10-2025, shared) | 41 | 0 | 37 / 7,232 | `.claude/settings.local.json` added |
| **Split point** | `47a1492` (30-10-2025, shared) | 40 | 0 | 36 / 7,295 | old mapping copy deleted |
| PyQt5 rewrite added beside the Tk code | `c258047` (07-02-2026, B only (after the split)) | 57 | 0 | 54 / 9,829 | `.claude/settings.local.json` deleted |
| Tk code deleted | `a62d3cc` (07-02-2026, B only) | 40 | 0 | 37 / 6,531 | — |
| Dangling import fixed | `c21f788` (07-02-2026, B only) | 40 | 0 | 37 / 6,531 | — |
| v3 look restored | `7e09e5c` (08-02-2026, B only) | 41 | 0 | 38 / 7,609 | — |
| **Snapshot** | `b0e7d0e` (09-02-2026, B only) | 41 | 0 | 38 / 7,821 | — |

Per folder (RAN (me) `scripts/sizes_area.py`; Python files / lines):

| Folder | `edac6a3` | `47a1492` (split) | `c258047` | `a62d3cc` | `7e09e5c` | `b0e7d0e` (snapshot) |
|---|---|---|---|---|---|---|
| `core/` | 4 / 1,456 | 4 / 1,454 | 4 / 1,406 | 4 / 1,406 | 4 / 1,448 | 4 / 1,448 |
| `utils/` | 3 / 850 | 3 / 850 | 3 / 850 | 3 / 850 | 3 / 850 | 3 / 850 |
| `gui/` (top level: `__init__.py`, main window) | 2 / 376 | 2 / 376 | 3 / 816 | 2 / 451 | 2 / 536 | 2 / 619 |
| `gui/handlers/` + `gui/tabs/` + `gui/utils/` (Tk) | 13 / 1,991 | 13 / 2,001 | 13 / 1,993 | — | — | — |
| `gui/views/` | — | — | 5 / 830 | 5 / 830 | 5 / 1,622 | 5 / 1,721 |
| `gui/controllers/` | — | — | 4 / 549 | 4 / 549 | 4 / 551 | 4 / 580 |
| `gui/theme/` | — | — | 2 / 344 | 2 / 344 | 2 / 426 | 2 / 426 |
| `gui/widgets/` | 3 / 99 | 3 / 99 | 4 / 141 | 2 / 49 | 3 / 124 | 3 / 124 |
| `models/` | — | — | 5 / 429 | 5 / 429 | 5 / 429 | 5 / 430 |
| `power_query_extractor/` (all) | 10 / 2,069 | 9 / 2,403 | 9 / 2,380 | 8 / 1,532 | 8 / 1,532 | 8 / 1,532 |
| `main.py` + `launch_extractor.py` | 2 / 113 | 2 / 112 | 2 / 91 | 2 / 91 | 2 / 91 | 2 / 91 |

What the numbers show (READ of the diffs, HIGH):
- The program arrived in git already large and already split into folders (37 Python files in the first commit). Git holds no record of how it was built; the compiled leftovers in the first commit show about 2½ months of earlier work outside git (source files dated 19-06-2025 to 23-07-2025) and a single-file window `gui/app.py` that was never committed (RAN (H-S) on the `.pyc` header fields; see 12B.4).
- Between the first commit and the split the program barely changed (+341 Python lines): one feature commit for the extractor (`bd176bf`), one configuration commit (`47a1492`) and small fixes.
- B's own history is five commits in three days (07-02-2026 to 09-02-2026). The business core (`core/`, `utils/`, `power_query_extractor/core/` and `power_query_extractor/config/`) shrank slightly (logging clean-up in `c258047`) and then grew only in `core/file_parser.py` (`7e09e5c`); the `power_query_extractor/` total fell by 848 lines only because the Tk extractor window was deleted (`a62d3cc`); `utils/constants.py`, `power_query_extractor/config/cell_mappings.py` and every file-name pattern ⚠, folder template and cell map ⚠ (**[GST RULE]** values) are byte-for-byte the split-point versions (RAN (me) `filehist.txt`: no B-only commit touches either file).
- Almost all of B's growth is in the window layer: the Tk window code (2,476 lines in `gui/` at the split, plus the 848-line Tk extractor window) was replaced by a PyQt5 window layer of 3,900 lines (`gui/` + `models/` at the snapshot), of which 1,721 are in `gui/views/` (the four tab views and their 5-line `__init__.py`).
- `c258047` briefly held both window layers at once (54 Python files, 9,829 lines); two minutes later `a62d3cc` deleted the Tk layer (17 files, −3,298 lines).

#### 12B.2 The architecture's evolution in phases

| Phase | Commits | What happened to the file structure |
|---|---|---|
| 0. Before git (19-06-2025 to 31-08-2025) | none (evidence: 70 `.pyc` files committed in `edac6a3`) | A single window file `gui/app.py` (89,275 bytes, last edited 25-06-2025) was replaced on or shortly before 23-07-2025 by the modular Tk window (`gui/main_window.py` + `tabs/` + `handlers/` + `utils/` + `widgets/`); the extraction mapping had an older format (21-06-2025), kept as `cell_mappings copy.py`; the consolidator's source grew from 10,208 bytes (25-06-2025) to 17,711 bytes (14-07-2025). RAN (H-S), header fields of the `.pyc` files only, HIGH for the sizes and dates; MEDIUM for the day of the GUI split (the headers record only each source's last-edit time). |
| 1. First upload | `edac6a3`, `cbfa4a5` (31-08-2025, shared - before the split on 30-10-2025) | Two programs in one folder: the organiser (`main.py`, `core/`, `utils/`, `gui/`, Tkinter) and the separate Power Query Extractor (`launch_extractor.py`, `power_query_extractor/`, Tkinter). The run log and the compiled files were committed by accident; the log was deleted on GitHub's web page about two minutes later (commit times 22:11:56 → 22:14:09). |
| 2. Tidy-up and first review (29-10-2025) | `6237597`, `9c56a1c`, `d674386`, `5ea77c2`, `1cb878b`, `01862a0`, `1c48722`, `4c42758` (shared - before the split on 30-10-2025) | No new program files. `.gitignore` added; compiled files removed; four files touched by a review that was most likely AI-assisted (MEDIUM, from the message style); README rewritten; a real row-colour bug fixed under a "redundant code" message. |
| 3. Extractor features (30-10-2025) | `bd176bf`, `47a1492` (shared - before the split on 30-10-2025) | Extractor window and processor grow (wait time, refreshed-file suffix, per-report ticks); six new extraction output sheets; the old mapping copy deleted; the AI tool's permission file `.claude/settings.local.json` committed. **Split point.** |
| 4. PyQt5 rewrite, one window (07-02-2026, 11:27) | `c258047` (B only (after the split)) | 18 new files (the message's heading says "New Files (20)" but its own list and the diff have 18 — RAN (verifier) numstat): `models/` (settings, run history, scan result, signal hub), `gui/views/` (four tabs), `gui/controllers/` (three background workers), `gui/theme/` (style sheets), `gui/pyqt_main_window.py`, `gui/widgets/collapsible_group.py`. The two programs are **merged** into one four-tab window; `main.py` and `launch_extractor.py` rewritten; the Tk files kept "for reference". The core files get only logging and error-handling clean-ups. From this commit on the window cannot open without pywin32 (12B.7, 12F.5). Commit trailer "Co-Authored-By: Claude Opus 4.6" (RAN (me) `git log` trailers). |
| 5. Tk layer deleted, start-up repaired (07-02-2026, 11:29 and 11:57) | `a62d3cc`, `c21f788` (B only) | 17 Tk files deleted without updating their importers: the program could not start at `a62d3cc` (RAN (me), 12B.7); `c21f788` repointed `gui/__init__.py` 27 minutes later. Two other dangling imports stayed (`gui/widgets/__init__.py` until `7e09e5c`; `power_query_extractor/extractor_main.py` to this day). Both commits carry the same AI trailer. |
| 6. v3 look restored (08-02-2026) | `7e09e5c` (B only) | New `gui/widgets/colored_section.py` ("replicates the old Tkinter create_colored_section", `gui/widgets/colored_section.py:2`); Setup and Validation views rebuilt as coloured cards; scan validation **moved** from `utils/helpers.py` into `core/file_parser.py`; the Extract tab now filled from run history. No AI trailer; message prefixed "feat:". |
| 7. Extraction driven by the disk (09-02-2026) | `b0e7d0e` (B only) | The Extract tab rebuilt as a client → version tree read from the target folder (`gui/pyqt_main_window.py:419-504`); the consolidated report and a new extraction log go into the Level 1 folder; scan statistics and tab status dots. Deletes the run-history approach added the day before. **Snapshot.** |

#### 12B.3 Every file at the snapshot

Columns: **Created** = commit and date that added the file; **Original purpose** = what the file did when it first appeared (READ of that version); **Changed in** = every later commit that touched it, with lines added/removed (RAN (me) `filehist.txt`); **Role** = whether its job expanded, shrank or shifted; **Lines** = git's line count when created → at `b0e7d0e`. No file at the snapshot was ever renamed or copied from another path (RAN (me): rename and copy detection both empty). Moves of *content* between files are in 12B.5.

**Entry points, configuration, documents**

| File | Created | Original purpose | Changed in | Role | Lines |
|---|---|---|---|---|---|
| `main.py` | `edac6a3` (31-08-2025, shared) | Starts the Tk organiser: UTF-8 console wrapper (crashes without a console), log to `gst_organizer.log` in the start folder, creates `core/ utils/ gui/ logs/ temp/` and missing `__init__.py` files inside the program folder at every start, banner prints, `input("Press Enter…")` pauses, error text listing the non-existent `gui/app.py` (`main.py:72@edac6a3`). | `5ea77c2` (29-10-2025, shared; −1: malformed `# !/usr/bin/env python3` line removed); `c258047` (07-02-2026, B only; +27/−46: PyQt5 `QApplication`, name "GST File Organizer", version "4.0", console wrapper guarded by `hasattr(sys.stdout, 'buffer')` `main.py:12-13`, no folder creation, no pauses, ImportError advice "pip install PyQt5 openpyxl" `main.py:49-54`). | Same job (entry point), shifted from Tk to PyQt5; shrank (no folder creation, no console pauses). The 6237597 claim "Corrected non-existent app.py reference in main.py" was only made true by `c258047` (READ: `main.py:71@47a1492` still prints `app.py`). | 83 → 63 |
| `launch_extractor.py` | `edac6a3` (shared) | Starts the separate Tk Power Query Extractor (`power_query_extractor.extractor_main.PowerQueryExtractor`), Tk error box on failure. | `c258047` (B only; +15/−17): builds the same PyQt5 main window and switches to tab index 3 (`launch_extractor.py:13-19`); errors only printed (`:21-23`). | **Shifted**: from "launch the other program" to "open the one program on its fourth tab". The file name and the idea of a separate extractor survive only as a name (12E). | 30 → 28 |
| `requirements.txt` | `edac6a3` (shared) | `openpyxl>=3.1.2`, `lxml>=4.9.0`, commented `pytest`/`black`/`pylint`; header "v3.0". | `c258047` (B only; +6/−5): header "v4.0", adds `PyQt5>=5.15.0`, adds commented `# pywin32>=306`, drops the dev-tool comments. | Expanded by one live line. Never lists pywin32 as a real requirement although the window needs it from `c258047` (12F). | 13 → 14 |
| `.gitignore` | `1cb878b` (29-10-2025, shared) | Ignores compiled files, virtual environments, IDE files, logs (`*.log`, `gst_organizer.log`, `pq_extractor.log`), the old cache name, Excel lock files, `temp_gst_*`, `backup_*`, output folders `/Annual Statement-*/`, `/mnt/user-data/` (the upload folder of Claude's web sandbox, MEDIUM that the file was generated there), Power Query cache files. | `c258047` (B only; +4/−1): adds `#Claude` / `.claude` (`.gitignore:81-82`). | Stable; one line added when the AI tool's settings file was untracked. Several entries name files only the dead extractor or unused constants would create (12E.5). | 79 → 82 |
| `README.md` | `edac6a3` (shared) | v3.0 user guide written for the single-file `gui/app.py` era. | `d674386` (29-10-2025, shared; +239/−135): rewritten for the modular Tk organiser plus the separate extractor, with a "Recent Bug Fixes (Latest Commit)" list (most likely AI-written — MEDIUM, from the footer "Author: Advanced AI Assistant", `README.md:364`, and the list repeating `6237597`'s message) and the footer "Last Updated: 29th October 2025". **Never changed by any B-only commit** (RAN (me) `filehist.txt`). | Frozen at the split; describes a program B no longer is (12B.8). | 260 → 364 |

**`core/` and `utils/` (business core)**

| File | Created | Original purpose | Changed in | Role | Lines |
|---|---|---|---|---|---|
| `core/__init__.py` | `edac6a3` (shared) | Re-exports `FileParser`, `FileOrganizer`, `ExcelHandler`; `__version__ = '3.0.0'`. | never | Unchanged (VERIFICATION RULE checked, 12B.7). Still says 3.0.0 in a "v4.0" program (`core/__init__.py:12`). | 13 → 13 |
| `core/file_parser.py` | `edac6a3` (shared) | Scans the top level of the source folder, recognises the six GST file-name patterns **[GST RULE]** ⚠ (B-R5–B-R11), groups files by `<client>-<state code>`, works out missing/duplicate types and statistics. | `6237597` (29-10-2025, shared; +11/−13): progress callback stored (`core/file_parser.py:29`) and called through `_update_progress` inside try/except (`:48-54`), so a failing callback no longer aborts the scan; `c258047` (B only; +1/−15): four `[DEBUG]` INFO log lines and an `# ADD THIS` comment removed; `7e09e5c` (B only; +62/−20): `_process_file` does its own single `stat()` (`:174-179`), size check (`:181-191`) and 8-byte signature check (`:193-215`) and reports failures as variations; new progress messages "Finding Excel files..." (`:147`), "No Excel files found" (`:151`), "Found N files, processing..." (`:155`) and "[i/N] name" (`:158`, replacing "Processing <name>"); `datetime` imported (`:9`). | **Expanded**: took over file validation from `utils/helpers.py` and made rejected files visible (12C). | 373 → 399 |
| `core/file_organizer.py` | `edac6a3` (shared) | Creates the Level 1 / Level 2 / Version / category folders, copies files with cleaned names, backups in Fresh mode, writes the per-client `_Organization_Report_*.txt`. | `c258047` (B only; +1/−1: `# ADD THESE` comment removed). | Unchanged in behaviour since the first commit. | 409 → 409 |
| `core/excel_handler.py` | `edac6a3` (shared) | Makes the ITC and Sales reports from the templates through Excel COM (attaching to an open Excel with `GetActiveObject`, then always `Quit()`), with an openpyxl fallback off Windows; writes the processing summary workbook. | `c258047` (B only; +22/−56): 32 `[DEBUG]` INFO lines removed or turned into 3 `logger.debug` lines (the full folder-and-file dump no longer reaches the log); bare `except:` narrowed; in the openpyxl fallback a cell that cannot be written now fails the whole report instead of being skipped (`core/excel_handler.py:251-252`; RAN (H-B1) with a merged cell); a now-unused `state_code` variable left at `:550`. | Same job; shrank (logging). The COM practice is the original one (12F.7). | 661 → 627 |
| `utils/__init__.py` | `edac6a3` (shared) | Re-exports 9 of the constants and 21 of the helpers (`utils/__init__.py:6-20`; not, for example, the state table `STATE_CODE_MAPPING`, `get_state_code`, `get_short_path` or `create_client_state_key`, which callers import from the modules directly); `__version__ = '3.0.0'`. | never | Unchanged; still re-exports helpers that became dead later (`validate_excel_file`, `get_file_info`, dead since `7e09e5c`) (12B.7). | 38 → 38 |
| `utils/constants.py` | `edac6a3` (shared) | Every fixed value of the v3 design: the six file-name patterns **[GST RULE]** ⚠, expected types **[GST RULE]** ⚠, folder templates, the template Links cell map **[GST RULE]** ⚠, GUI config, processing modes, extensions/signatures, messages, `EXCEL_SAFETY`, the state-name → code table **[GST RULE]** ⚠. | never | **Unchanged since the first commit** (RAN (me) `filehist.txt`; VERIFICATION RULE 12B.7: it imports nothing internal). Its role shrank in practice: `GUI_CONFIG` and `MESSAGES` lost their last readers when the Tk files were deleted in `a62d3cc` (12E.5). | 275 → 275 |
| `utils/helpers.py` | `edac6a3` (shared) | Time stamps, file-name cleaning, copy/backup, Excel validation, folder listing, client-name checks, state codes, client-folder key, progress tracker, assorted unused helpers. | `6237597` (shared; +1/−1: `from .constants import STATE_CODE_MAPPING` relative import, `utils/helpers.py:491`); `c258047` (B only; +1/−1: `# ADD THIS LINE` comment removed); `7e09e5c` (B only; +8/−8): `find_excel_files` no longer validates each file and skips `~$` lock-file names (`:255-262`). | **Shrank**: validation moved out to the parser; `validate_excel_file` (`:210-247`) and `get_file_info` (`:115-136`) became dead. | 537 → 537 |

**`power_query_extractor/` (the former second program)**

| File | Created | Original purpose | Changed in | Role | Lines |
|---|---|---|---|---|---|
| `power_query_extractor/__init__.py` | `edac6a3` (shared) | `__version__ = "1.0.0"`. | never | Unchanged. | 3 → 3 |
| `power_query_extractor/extractor_main.py` | `edac6a3` (shared) | Start-up class of the separate Tk extractor: log to `pq_extractor.log`, cache lookup in three relative files, opens `PowerQueryExtractorApp`. | never | **Broken since `a62d3cc`**: line 9 imports the deleted `power_query_extractor/gui/extractor_window.py` (RAN (me): `ModuleNotFoundError: No module named 'power_query_extractor.gui.extractor_window'`). Nothing imports it; dead (registry B-F31, B-R128). | 79 → 79 |
| `power_query_extractor/config/__init__.py` | `edac6a3` (shared) | Docstring only. | never | Unchanged. | 1 → 1 |
| `power_query_extractor/config/cell_mappings.py` | `edac6a3` (shared) | [GST RULE] Which report cells are extracted into which output sheet and column (3 output sheets, 8 cells ⚠), extraction options, and a generated "backward compatibility" `CELL_MAPPINGS` table. | `47a1492` (30-10-2025, shared; +126): six more ITC output sheets, 14 more cells ⚠ (12C). | Expanded once, before the split; **never changed in B** (RAN (me) `filehist.txt`). | 152 → 278 |
| `power_query_extractor/core/__init__.py` | `edac6a3` (shared) | Docstring only. | never | Unchanged. | 1 → 1 |
| `power_query_extractor/core/report_processor.py` | `edac6a3` (shared) | Per client and report: copy the report to a "refreshed" name, open it in a visible Excel, refresh Power Query by SendKeys Alt+A, R, A, wait fixed times, close error dialogs, check `Info!BB2` **[GST RULE]** ⚠ (B-R71), save, then read the configured cells **[GST RULE]** ⚠ in a second hidden Excel. Imports pywin32 at module level with no fallback (`:11-15`). | `bd176bf` (30-10-2025, shared; +41/−17): configurable wait (`self.wait_time`, `:27`) and refreshed-file suffix (`:28`, `:91-94`), per-report selection with status "Not selected for processing" (`:39-70`); `c258047` (B only; +18/−26): stray `print` and an unreachable `return` removed, bare `except:` narrowed, paste-marker comments removed. | Same job; since `c258047` it is **loaded at window start-up** (through `gui/controllers/extraction_controller.py:10`), which made its unguarded pywin32 import a start-up requirement (12F.5). | 736 → 752 |
| `power_query_extractor/core/data_consolidator.py` | `edac6a3` (shared) | Writes the consolidated extraction workbook `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`: a Summary sheet plus one sheet per configured output sheet; two older layouts left as dead methods. | `c258047` (B only; +2/−17): a pasted duplicate header-writing block removed (no change to the output, READ), narrower `except`, paste-marker comments removed. | Unchanged in output. Its first-substring column look-up (`:130-138`) has put the "Diff GSTR-3B" ⚠ cells into three other sheets ever since `47a1492` added one-word column names (12C, B-R83) **[GST RULE]** — a technical mismatch; which cells are right is the author's call. | 432 → 417 |
| `power_query_extractor/gui/__init__.py` | `edac6a3` (shared) | Package of the Tk extractor window. | never | **Emptied by `a62d3cc`** (its only module deleted); now an empty package that only the broken `extractor_main.py` refers to. | 1 → 1 |

**`gui/` (window layer; everything below except `gui/__init__.py` and `gui/widgets/__init__.py` is B only (after the split))**

| File | Created | Original purpose | Changed in | Role | Lines |
|---|---|---|---|---|---|
| `gui/__init__.py` | `edac6a3` (shared) | Re-exports the Tk `GSTOrganizerApp` from `gui/main_window.py`; `__version__ = '3.0.0'`. | `c21f788` (07-02-2026, B only; +5/−5): imports `GSTOrganizerWindow` from `gui/pyqt_main_window.py`; `__version__ = '4.0.0'`. | **Shifted** from the Tk window to the PyQt5 window. Between `a62d3cc` and `c21f788` it pointed at a deleted file and the program could not start (RAN (me), 12B.7). Being an *eager* front door, it makes any import under `gui` load the whole window, all views, all controllers and the pywin32-dependent refresh module (RAN (me): tkinter loaded at `c258047`, not at `c21f788`). | 11 → 11 |
| `gui/pyqt_main_window.py` | `c258047` (B only) | The new main window: state objects, settings store, run history, migration of the old settings, three controllers, four tabs, menus, signal wiring, dry run, start checks, completion handling; the Extract tab filled only from the processing run just finished, with a whole-target name search (`_build_extraction_clients`, `rglob('*')`). | `7e09e5c` (B only; +132/−47): scan progress finally connected (`gui/pyqt_main_window.py:129, 238-242`), SCAN button disabled during a scan, Extract tab filled from run history at start-up and after processing, a two-level name search; `b0e7d0e` (B only; +203/−120): Extract tab rebuilt from a disk scan of the greatest-named `Annual Statement-*` folder (`:419-504`), scan statistics in the scan log and summary (`:244-316`), tab status dots (`:569-590`), switch to the Extract tab after processing (`:401`); the run-history and name-search code of the day before deleted. | **Expanded** from "thin wiring" (the `c258047` message's words) into the owner of the scan log text, statistics display, the disk scan for extraction and tab status. The source of the Extract list changed three times (12E). | 440 → 608 |
| `gui/controllers/__init__.py` | `c258047` (B only) | Re-exports the three controllers. | never | Unchanged. | 4 → 4 |
| `gui/controllers/scan_controller.py` | `c258047` (B only) | Runs `FileParser.scan_folder` in a QThread; refuses a second scan; returns a `ScanResult`. | `b0e7d0e` (B only; +3/−1): also passes `parser.get_statistics()` (`:31-37`). | Expanded slightly. Replaces the synchronous Tk scan (which froze the window). | 92 → 94 |
| `gui/controllers/processing_controller.py` | `c258047` (B only) | Port of the Tk processing flow into a QThread worker: folders, copies, ITC and Sales reports, organisation report, summary workbook, run-history records. Passes `'variations': []` to the summary (`:73`), so the summary's Variations sheet is always empty (regression against `gui/handlers/processing_handler.py:161@47a1492`). | `7e09e5c` (B only; +2/−0): `client_key` and `version_folder` added to each per-client summary entry (`:139-140`) — never read. | Unchanged in role. | 296 → 298 |
| `gui/controllers/extraction_controller.py` | `c258047` (B only) | Runs refresh + extraction in a QThread; writes the skip-refresh choice into the shared `EXTRACTION_OPTIONS`; consolidated report into the **target folder**. | `b0e7d0e` (B only; +42/−15): consolidated report into the **Level 1 folder** (`:92`), new `Extraction_Log_<YYYYMMDD_HHMM>.txt` (`:109-121`), emoji removed from its own messages. | Expanded (log file); output location moved back to where the old Tk extractor put it. | 157 → 184 |
| `gui/theme/__init__.py` | `c258047` (B only) | Re-exports `ThemeManager`. | never | Unchanged. | 2 → 2 |
| `gui/theme/theme_manager.py` | `c258047` (B only) | Light and dark application style sheets; `apply_theme`, unused `toggle`. Replaces the 212-line Tk recursive recolourer. | `7e09e5c` (B only; +82): unused `COLORS` table (`:10-22`), `#success` / `#warning` / `#info` button rules in both sheets (`#info` unused). | Expanded, but from `7e09e5c` many widgets carry their own inline light colours, so dark mode no longer reaches Setup and Validation (RAN (H-B2) screenshots). | 342 → 424 |
| `gui/views/__init__.py` | `c258047` (B only) | Re-exports the four views. | never | Unchanged. | 5 → 5 |
| `gui/views/setup_view.py` | `c258047` (B only) | Step 1: four path boxes with Browse buttons (template filter `*.xlsx *.xlsm`), three mode buttons with new wording, "Include client name in folder names" + Max length 5-100, Scan / Re-Scan, a help panel. | `7e09e5c` (B only; +522/−123): coloured cards, collapsible INSTRUCTIONS and EXPECTED FILE NAMES cards (`CollapsibleCard`, `:17-82`), scan log panel (`:519-588`), v3 wording restored for the mode buttons and the tick box, Max length 15-100 (`:375`), the v3 note "Names >10 chars will show warning…" (`:385`); `b0e7d0e` (B only; +10/−2): `collapse_info_cards()` (`:590-594`). | **Expanded** greatly (largest file in B). The redesign also made three Browse buttons white-on-white (RAN (H-B2), RAN (VW)). | 240 → 647 |
| `gui/views/validation_view.py` | `c258047` (B only) | Step 2: client tree with 9 columns, all clients ticked by default, Select All / Clear All / Complete Only, summary box, Dry Run, Export List (never connected), Start Processing. | `7e09e5c` (B only; +319/−72): two-column layout, instructions frame, coloured buttons; behaviour unchanged. | Same job, restyled. | 200 → 447 |
| `gui/views/processing_view.py` | `c258047` (B only) | Step 3: progress bar, status line, Start / Stop / Clear Log, HTML log (not escaped). | `7e09e5c` (B only; +67/−13): banner, status above the bar, inline styles. | Restyled only. | 112 → 166 |
| `gui/views/extraction_view.py` | `c258047` (B only) | Step 4: wait time 5-120, suffix, skip refresh; flat client list with ITC/Sales ticks (pre-ticked, always enabled) and a "Last Refresh" column from run history (always blank, key mismatch). | `7e09e5c` (B only; +102/−10): restyle; stores a version folder per row; `b0e7d0e` (B only; +162/−71): client → version tree (`populate_tree`, `:291-356`), "Refresh List" button, ticks start empty and only existing reports are tickable, "Refresh Status" from files. | **Shifted** from a list fed by the session to a tree fed by the disk. | 273 → 456 |
| `gui/widgets/__init__.py` | `edac6a3` (shared) | Re-exports the Tk `CollapsibleFrame` and `TitleBar`. | `7e09e5c` (B only; +3/−3): re-exports `CollapsibleGroup` and `ColoredSection` (`:4-7`). | **Shifted** Tk → PyQt5. At `a62d3cc` and `c21f788` (until `7e09e5c` repaired it) it imported two deleted files (RAN (me) at `c21f788`, re-run by the verifier: `ModuleNotFoundError: No module named 'gui.widgets.collapsible_frame'`; at `a62d3cc` the same import fails one step earlier, on `gui.main_window`, because `gui/__init__.py` runs first); harmless only because nothing imported `gui.widgets` then. `7e09e5c` had to fix it because the new views import `gui.widgets.colored_section` (an unannounced fix). | 7 → 7 |
| `gui/widgets/collapsible_group.py` | `c258047` (B only) | A checkable group box meant to replace the Tk `CollapsibleFrame` (imports animation classes it never uses, `:6`). | never | **Never used** by any view at any commit (READ `git grep` at `c258047` and `7e09e5c`; DC D-04); `setup_view.py` defines its own `CollapsibleCard`. Loaded at start-up only through the re-export added in `7e09e5c` (from `c258047` to `c21f788` nothing loaded it at all; RAN (verifier) `git grep` at every B-only commit). | 42 → 42 |
| `gui/widgets/colored_section.py` | `7e09e5c` (B only) | White card with a coloured header, "replicates the old Tkinter create_colored_section" (`:2`; the original was `gui/utils/ui_helpers.py:43@47a1492`). | never | Unchanged. Its inner rule `QWidget { background: white; }` (`:62`) is what hides the Browse buttons' colour. | 75 → 75 |

**`models/` (state layer; all B only (after the split))**

| File | Created | Original purpose | Changed in | Role | Lines |
|---|---|---|---|---|---|
| `models/__init__.py` | `c258047` (B only) | Re-exports `AppSettings`, `ScanResult`, `ProcessingState`, `ClientRunRecord`, `RunHistory`, `CacheManager` (not `AppSignals`). | never | Unchanged. | 4 → 4 |
| `models/app_state.py` | `c258047` (B only) | `AppSettings` (11 fields — the commit message says 12), `ScanResult`, `ProcessingState`. | `b0e7d0e` (B only; +1): `ScanResult.statistics` (`:55`). | Expanded by one field; `ScanResult.clear()` (`:70-74`) was not updated to clear it, but it is never called. | 87 → 88 |
| `models/cache_manager.py` | `c258047` (B only) | Per-user files in `~/.gst_organizer/` (`settings.json`, `run_history.json`, `session_state.json`) and the one-time import of the old `~/.gst_organizer_cache.json` ("Replaces the old gui/utils/cache_manager.py", `:3`). | never | Unchanged. The import reads `recent_source_folders` (`:122`) where the old writer stored `recent_folders` (`gui/utils/cache_manager.py:34-43@47a1492`): a writer/reader mismatch since its creation. | 152 → 152 |
| `models/run_history.py` | `c258047` (B only) | Run and per-client history, last 50 runs, per-client records never trimmed. | never | Unchanged; the extraction records it receives changed key twice (12C, B-R120). | 148 → 148 |
| `models/signals.py` | `c258047` (B only) | "Central signal hub - replaces direct self.app.xxx references" (`:10`), 18 signals (`models/signals.py:13-38`; RAN (verifier) count of `pyqtSignal(...)` assignments — [Corrected during verification — originally stated 20]). | never | **Never imported by anything, at any commit** (DC D-01; RAN (me) `verify_imports.txt` lists no importer). Each controller defines its own signals instead. | 38 → 38 |

#### 12B.4 Files that existed in the history but not at the snapshot

RAN (me) `filehist.txt` (20 non-compiled paths, 70 compiled paths). READ of every deleted Tk file is in H-B1 (Part "a62d3cc"), which I spot-checked at `47a1492` for the lines cited here.

| File | Created | Changed | Deleted | What it did and where its job went |
|---|---|---|---|---|
| `gst_organizer.log` | `edac6a3` (+5,689) | — | `cbfa4a5` (31-08-2025, on GitHub's web page); reached the local line through the merge `9c56a1c` | The program's run log, committed by accident. Described by kind only: client names, state names, source file names and Windows folder paths from real runs; no tax figures (RAN (H-S): shapes and counts only). **Still readable in the history of both projects** at `edac6a3`. |
| 70 × `**/__pycache__/*.pyc` (35 modules × Python 3.12 and 3.13) | `edac6a3` | — | `01862a0` (29-10-2025, shared) | Compiled leftovers; their headers are the only evidence of the pre-git history (12B.2 phase 0). They embed the developer computer's folder paths (not reproduced). |
| `gui/app.py` (never committed) | — | — | — | Known only from `gui/__pycache__/app.cpython-312.pyc` / `-313.pyc` in `edac6a3`: an 89,275-byte single-file GUI, last edited 25-06-2025 (RAN (H-S), HIGH). Replaced before git by the modular Tk GUI. |
| `power_query_extractor/config/cell_mappings copy.py` | `edac6a3` (63; already obsolete: its CRLF size matches a `.pyc` dated 21-06-2025, RAN (H-S)) | — | `47a1492` (shared) | The older mapping format (`CELL_MAPPINGS` per sheet and cell, same cells as the live file except `Diff - GSTR-1!I1` ⚠). The name contains a space, so it could never be imported. |
| `.claude/settings.local.json` | `bd176bf` (9) | — | `c258047` (B only (after the split)) | Claude Code (Anthropic's AI coding tool) permission file allowing `python -m py_compile` without asking; no secrets (READ (H-S), HIGH). B untracked it and added `.claude` to `.gitignore`; the `c258047` message says "already in .gitignore", which was not true before that commit (READ). |
| `gui/main_window.py` | `edac6a3` (365) | `c258047` (±2: two bare `except:` narrowed to `except tk.TclError:`) | `a62d3cc` (B only) | Tk `GSTOrganizerApp` "god object": window, variables, handlers, tabs, status bar, scan summary with statistics, the "Processing is still running. Stop and exit?" question (`gui/main_window.py:314-325@47a1492`), cache save/apply. Job → `gui/pyqt_main_window.py` (without the close question and without the statistics until `b0e7d0e`). |
| `gui/handlers/__init__.py` | `edac6a3` (8) | — | `a62d3cc` | Re-exports the three Tk handlers. |
| `gui/handlers/client_handler.py` | `edac6a3` (259) | `1c48722` (29-10-2025, shared; −6: row-colour bug fixed) | `a62d3cc` | Tk client tree, keyboard selection, per-client "include client name" toggle with a >10-character warning (`gui/handlers/client_handler.py:175-206@47a1492`), export of ticked clients to a text file (`:221-253`). Job → `gui/views/validation_view.py`; the per-client toggle and the export were **not** carried over (the EXPORT LIST button exists but is not connected, `gui/views/validation_view.py:186-221`). |
| `gui/handlers/file_handler.py` | `edac6a3` (151) | `6237597` (±1, no behaviour change) | `a62d3cc` | Browse dialogs (template filter `*.xltx *.xlsx`, `gui/handlers/file_handler.py:31, 42@47a1492`), saving settings on every browse, synchronous scan with status-bar progress and "No valid GST files found" (`:90-91`), input checks (`:120-151`). Job → `gui/views/setup_view.py`, `gui/controllers/scan_controller.py`, `gui/pyqt_main_window.py:218-225, 345-361`. |
| `gui/handlers/processing_handler.py` | `edac6a3` (447) | `6237597` (+46/−30: `root.after` wrappers, long-name threshold from the max-length setting); `c258047` (B only; +331/−339: split into 7 methods, 14 `[DEBUG]` log lines removed) | `a62d3cc` (two minutes after its refactor, so the refactored version never ran in a released state) | Tk dry run (with template/target checks), long-client-name confirmation (`gui/handlers/processing_handler.py:64-91@47a1492`), processing thread, summary with the real variations (`:161`), opening the output folder on Windows, macOS and Linux (`:442-449`). Job → `gui/controllers/processing_controller.py` + `gui/pyqt_main_window.py:326-343`; the long-name confirmation, the variations and the non-Windows folder opening were lost. |
| `gui/tabs/__init__.py`, `processing_tab.py`, `setup_tab.py`, `validation_tab.py` | `edac6a3` (8 / 117 / 377 / 247) | — | `a62d3cc` | The three Tk screens. Setup: template label "(.xlsx or .xltx)" (`gui/tabs/setup_tab.py:121@47a1492`), mode descriptions from `PROCESSING_MODES` (`:212-223`), "Include client name in brackets for Level 4 folders" (`:230`), note "Names >10 chars will show warning / This overrides individual client settings" (`:238-239`), max-length spin box 15-100 with a 10-100 validator (`:252-279`). Job → `gui/views/*.py`. |
| `gui/utils/__init__.py`, `cache_manager.py`, `status_bar.py`, `ui_helpers.py` | `edac6a3` (9 / 56 / 37 / 63) | — | `a62d3cc` | Old settings file `~/.gst_organizer_cache.json` with `recent_folders` (`gui/utils/cache_manager.py:15, 34-43@47a1492`); dark status bar "v3.0 \| Production Ready"; time-stamped coloured log insertion and `create_colored_section` (`gui/utils/ui_helpers.py:17-43@47a1492`). Jobs → `models/cache_manager.py`, `QStatusBar`, the views' own log functions (without time stamps), `gui/widgets/colored_section.py` (`7e09e5c`). |
| `gui/utils/dark_mode_manager.py` | `edac6a3` (212) | `c258047` (±4) | `a62d3cc` | Recursive Tk widget recolouring for dark mode. Job → `gui/theme/theme_manager.py` (QSS). |
| `gui/widgets/collapsible_frame.py`, `title_bar.py` | `edac6a3` (41 / 51) | — | `a62d3cc` | Tk collapsible frame; blue title bar with a Dark Mode tick box. Jobs → `CollapsibleCard` in `gui/views/setup_view.py:17-82` (`7e09e5c`; the intended replacement `collapsible_group.py` is unused); window title + View > Dark Mode menu. |
| `power_query_extractor/gui/extractor_window.py` | `edac6a3` (601) | `4c42758` (29-10-2025, shared; +30: mouse-wheel scrolling); `bd176bf` (shared; +362/−145: two-column layout, wait 5-60 s, suffix, per-report ticks disabled when the report is missing, "Last Refresh Status" per report type); `c258047` (±2) | `a62d3cc` (848 lines at deletion) | The separate Tk extractor window: found work on disk (newest `Annual Statement-*` by modification time, `power_query_extractor/gui/extractor_window.py:517@47a1492`; client folders not starting with `_`, `:525`; one version per client, `max` by name, `:529`), wrote the consolidated report into the Level 1 folder (`:808`) and offered to open it (`:818-820`). Job → `gui/views/extraction_view.py` + `gui/controllers/extraction_controller.py` + (from `b0e7d0e`) `gui/pyqt_main_window.py:419-504`. |

#### 12B.5 Renames, splits and merges

- **Renames and copies:** none in B's whole history. RAN (me): `git log -M --diff-filter=R B/main` and `git log -C -C --diff-filter=C B/main` both return nothing. Every move of responsibility happened by writing a new file and deleting an old one.
- **Merge of the two programs into one (the largest structural change).** At the split there were two programs with two entry points and two Tk windows (organiser: `main.py` → `gui/main_window.py`; extractor: `launch_extractor.py` → `power_query_extractor/extractor_main.py` → `power_query_extractor/gui/extractor_window.py`). `c258047` put both into one four-tab PyQt5 window; `a62d3cc` deleted both Tk windows. The extractor's former window code now lives in three places: the screen in `gui/views/extraction_view.py`, the background job in `gui/controllers/extraction_controller.py`, and (from `b0e7d0e`) the disk discovery in `gui/pyqt_main_window.py:419-504`. Its core (`power_query_extractor/core/*`, `config/cell_mappings.py`) stayed in place. `launch_extractor.py` became a shortcut to tab 4; `extractor_main.py` was left behind, broken.
- **Splits of the old Tk handlers (c258047 → a62d3cc).** `gui/handlers/processing_handler.py` → `gui/controllers/processing_controller.py` (the worker) + `gui/pyqt_main_window.py:326-343` (dry run) + `:345-371` (start checks). `gui/handlers/file_handler.py` → `gui/views/setup_view.py` (pickers) + `gui/controllers/scan_controller.py` (scan in a thread) + `gui/pyqt_main_window.py:218-243` (checks and progress). `gui/handlers/client_handler.py` → `gui/views/validation_view.py:363-447`. `gui/main_window.py` → `gui/pyqt_main_window.py` + `models/app_state.py` (the Tk variables became `AppSettings` fields) + `models/cache_manager.py`.
- **Replacements one-for-one.** `gui/tabs/*` → `gui/views/*`; `gui/utils/dark_mode_manager.py` → `gui/theme/theme_manager.py`; `gui/utils/cache_manager.py` → `models/cache_manager.py` (new folder, three JSON files, one-time import of the old file); `gui/utils/status_bar.py` → Qt's status bar; `gui/widgets/title_bar.py` → window title and menu; `gui/utils/ui_helpers.py:43@47a1492` `create_colored_section` → `gui/widgets/colored_section.py` (`7e09e5c`); `gui/widgets/collapsible_frame.py` → `CollapsibleCard` inside `gui/views/setup_view.py` (`7e09e5c`), while the replacement first written for it, `gui/widgets/collapsible_group.py` (`c258047`), was never used.
- **Merge of duplicated validation into the parser (`7e09e5c`).** File validation lived in `utils/helpers.py` (`validate_excel_file`, `:210-247`, called by `find_excel_files` and again by the parser through `get_file_info` + `validate_excel_file`); `7e09e5c` moved the size and signature checks into `core/file_parser.py:171-226` (one `stat`, one 8-byte read) and left both helpers in place, unused.
- **Split inside one file (`c258047`).** The 344-line Tk `process_files_thread` was cut into seven methods (`_initialize_processing` … `_cleanup_processing`); the file was deleted two minutes later, so the split survives only as the shape of `ProcessingWorker._process` / `_create_report` in `gui/controllers/processing_controller.py:42-239`.
- **Pre-git split** (evidence only): the single `gui/app.py` → the modular Tk GUI (`gui/main_window.py:4@edac6a3` "Modularized version - ties all components together").

#### 12B.6 How each core file changed, commit by commit

**`core/file_parser.py`** (4 commits)
1. `edac6a3` (31-08-2025, shared - before the split on 30-10-2025): `scan_folder` called `find_excel_files` (which already threw away files under 1 KB or without an Excel signature, silently) and then, for each file, `get_file_info` + `validate_excel_file` again (`_process_file`); a second check that could never fail, because the list was pre-filtered (READ `utils/helpers.py:249-266@edac6a3`). Many `[DEBUG]` INFO lines for Sales files (5,334 of the 5,689 lines of the committed run log came from this file, RAN (H-S) counts).
2. `6237597` (29-10-2025, shared): progress callback stored and wrapped (`core/file_parser.py:29, 48-54`); about a dozen blank lines removed.
3. `c258047` (07-02-2026, B only (after the split)): four `[DEBUG]` lines and one paste marker removed; parsing rules, client key and completeness untouched (READ diff).
4. `7e09e5c` (08-02-2026, B only): the validation moved in and became visible (12C B-R2, B-R3, B-R4); `datetime` import; finer progress messages. RAN (me) `scripts/scan_trace.py` on extracted trees with five made-up files (one valid, one 104-byte, one with a wrong signature, one valid `~$…` lock-file name, one unparseable name): at `47a1492` and `c21f788` the scan returned 3 scanned files and 2 variations (the 104-byte and wrong-signature files vanished without trace; the `~$` file appeared as "No matching pattern found"); at `7e09e5c` and `b0e7d0e` it returned 2 scanned files and 3 variations ("File too small to be valid Excel", "Invalid or corrupted Excel file", "No matching pattern found"), and the `~$` file was skipped by name.

**`core/excel_handler.py`** (2 commits) — `edac6a3` (shared): the whole Step 3 report mechanism (COM attach/hide/quit, copy template then fill, Links sheet, summary workbook). `c258047` (B only): logging diet (the full data-mapping dump with client folder paths and file names left the INFO log), narrower `except`, openpyxl fallback now fails a report on an unwritable cell, an unused `state_code` (`:550`) that still logs the "Unknown state" warning. Nothing about which cells are written, or how Excel is driven, changed in B (READ diff; RAN (me) `filehist.txt`).

**`core/file_organizer.py`** (2 commits) — behaviour fixed at `edac6a3`; `c258047` removed one comment. Every folder and copy rule in B is the split-point rule (12C category 6–7).

**`utils/helpers.py`** (4 commits) — `edac6a3`; `6237597` relative import of the state table; `c258047` comment; `7e09e5c` `find_excel_files` reduced to "list `*.xlsx`, `*.xls`, `*.xlsm` at the top level, drop `~$` names, sort" (`utils/helpers.py:249-266`).

**`utils/constants.py`** (1 commit) — never changed. **[GST RULE]** values: the file-name patterns ⚠, expected types ⚠, folder templates, template Links cell map ⚠, processing modes, `EXCEL_SAFETY` and the state table ⚠ at the snapshot are exactly those of 31-08-2025.

**`power_query_extractor/config/cell_mappings.py`** (2 commits) **[GST RULE]** — `edac6a3` (3 output sheets, 8 cells ⚠; re-run by the verifier by loading each commit's file); `47a1492` (9 output sheets, 22 cells ⚠). Nothing in B.

**`power_query_extractor/core/report_processor.py`** (3 commits) — `edac6a3` (keystroke refresh with fixed 10 s + 10 s waits); `bd176bf` (configurable wait and suffix, per-report selection); `c258047` (clean-up only). The refresh method itself (Dispatch, visible Excel, Alt+A, R, A, fixed waits) has not changed since the first commit (RAN (me) pickaxe, re-run by the verifier over the `.py` files: the text `SendKeys` (the refresh keystrokes, `power_query_extractor/core/report_processor.py:238, 243, 248`) and `GetActiveObject` (which belongs to Step 3, `core/excel_handler.py:111`, not to the refresh) were added in `edac6a3` and never added or removed since; `DispatchEx` and `RefreshAll` never appear. If the compiled `.pyc` files are included, the pickaxe also lists `01862a0`, the commit that deleted the `.pyc` files containing those words — [Corrected during verification — originally stated "appear only in `edac6a3`" without the `.py` restriction and implied `GetActiveObject` is part of the refresh]).

**`power_query_extractor/core/data_consolidator.py`** (2 commits) — `edac6a3`; `c258047` removed a duplicate header block (no output change).

**`gui/pyqt_main_window.py`** (3 commits) — `c258047` → `7e09e5c` → `b0e7d0e` as in 12B.3. RAN (me) pickaxe: `_populate_extraction_from_results` added `c258047`, removed `7e09e5c`; `_populate_extraction_from_history` added `7e09e5c`, removed `b0e7d0e`; `_build_extraction_clients` added `c258047`, removed `b0e7d0e`; `_find_latest_version_folder` added `7e09e5c`, removed `b0e7d0e`; `_populate_extraction_from_disk`, `_make_dot_icon` and `collapse_info_cards` added `b0e7d0e`.

**`main.py`** (3 commits) — `edac6a3` → `5ea77c2` (shebang) → `c258047` (PyQt5, guarded console wrapper, no folder creation, no pauses). Unchanged since 07-02-2026.

#### 12B.7 VERIFICATION RULE (Project_Audit_FINAL.md, 12B): does any "never changed" file import something added later?

What the rule asks: a file that imports something added or changed in a later commit must itself have changed. RAN (me) `scripts/verify_imports.py` → `verify_imports.txt`: for every Python file at the snapshot it finds the last commit that changed the file, reads every import of the project's own modules (`core`, `utils`, `gui`, `models`, `power_query_extractor`), checks that each imported module and name existed at that commit and at the snapshot, and finds the commit in which each imported name first appeared.

| File never changed since creation (at the snapshot) | Created | Internal imports | Newest imported name first appeared in | Result |
|---|---|---|---|---|
| `core/__init__.py` | `edac6a3` | 3 (`FileParser`, `FileOrganizer`, `ExcelHandler`) | `edac6a3` | consistent |
| `utils/__init__.py` | `edac6a3` | 30 (constants and helpers) | `edac6a3` | consistent (re-exports two helpers dead since `7e09e5c`; a dead re-export is not a contradiction) |
| `utils/constants.py` | `edac6a3` | 0 | — | consistent |
| `power_query_extractor/__init__.py`, `config/__init__.py`, `core/__init__.py`, `gui/__init__.py` | `edac6a3` | 0 | — | consistent |
| `power_query_extractor/extractor_main.py` | `edac6a3` | 1 (`.gui.extractor_window.PowerQueryExtractorApp`) | `edac6a3` | **reverse case**: the import resolved when the file was last changed (`edac6a3`) but the imported file was **deleted** in `a62d3cc`; the file was not updated, so it is broken at the snapshot (RAN (me) import: `ModuleNotFoundError`). |
| `gui/controllers/__init__.py`, `gui/theme/__init__.py`, `gui/views/__init__.py`, `models/__init__.py` | `c258047` | 3 / 1 / 4 / 6 | `c258047` | consistent |
| `gui/widgets/collapsible_group.py`, `models/run_history.py`, `models/signals.py` | `c258047` | 0 | — | consistent |
| `models/cache_manager.py` | `c258047` | 2 (`AppSettings`, `RunHistory`) | `c258047` | consistent; `models/app_state.py` changed later (`b0e7d0e` added `ScanResult.statistics`), but `cache_manager` does not use `ScanResult` |
| `gui/widgets/colored_section.py` | `7e09e5c` | 0 | — | consistent |

Files that *did* change were also checked: in every one, each imported name existed at the file's last change (no file imports a name first added after its own last change). The two files whose imports moved forward were changed in the same commit as the thing they import: `gui/views/setup_view.py` and `gui/views/validation_view.py` import `ColoredSection` (added `7e09e5c`) and were themselves changed in `7e09e5c`; `gui/widgets/__init__.py` re-exports `ColoredSection` and changed in `7e09e5c` — the rule is satisfied, and this is exactly why `gui/widgets/__init__.py` *had* to change in `7e09e5c` although the commit message does not mention it.

**Reverse check — imports of files that had been deleted** (RAN (me) `scripts/broken_imports_by_commit.py` over all 17 commits; the import results in the right-hand column were re-run by the verifier on fresh extracted trees with the same outcome, and the verifier's own VERIFICATION RULE check, `verify_rule.py`, found only the `extractor_main.py` case; the checker's reports about `FILE_PATTERNS` / `EXPECTED_FILE_TYPES` are false alarms, because those two are written with a type annotation, `utils/constants.py:13, 68`, which the checker did not read as a definition — confirmed by READ):

| Commit | Dangling imports at that commit | Effect (RAN (me) on extracted trees, with stand-in pywin32 modules where stated) |
|---|---|---|
| `edac6a3` … `c258047` | none | `47a1492`: `import gui.main_window` works without pywin32 (pythoncom not loaded; tkinter loaded). `c258047`: `import gui.pyqt_main_window` fails without pywin32 (`No module named 'pythoncom'`); with stand-ins it works and tkinter is loaded (because `gui/__init__.py` still imports the Tk window). |
| `a62d3cc` | `gui/__init__.py:7` → deleted `gui/main_window.py`; `gui/widgets/__init__.py:4-5` → deleted `collapsible_frame.py`, `title_bar.py`; `power_query_extractor/extractor_main.py:9` → deleted `extractor_window.py` | `import gui.pyqt_main_window` fails even with stand-ins: `No module named 'gui.main_window'` — **the program cannot start** (H-B1 RAN the same through `main.py`: it prints the misleading "pip install PyQt5 openpyxl"). |
| `c21f788` | `gui/widgets/__init__.py:4-5`; `extractor_main.py:9` | Window starts with stand-ins; tkinter no longer loaded. `import gui.widgets` fails (`No module named 'gui.widgets.collapsible_frame'`), harmless because nothing imported it yet. |
| `7e09e5c`, `b0e7d0e` | `extractor_main.py:9` only | Window starts with stand-ins; without pywin32 it still fails (`No module named 'pythoncom'`). `import power_query_extractor.extractor_main` fails (`No module named 'power_query_extractor.gui.extractor_window'`). |

#### 12B.8 Where the older document (README.md) disagrees with the file history

`README.md` is a source of intent, not of truth; it was last changed in `d674386` (29-10-2025, shared - before the split on 30-10-2025) and describes the v3.0 Tkinter program (READ).

| README says (line) | File history shows |
|---|---|
| Title and footer "v3.0 - Production Ready", "Version: 3.0", "Last Updated: 29th October 2025" (`README.md:1, 361-363`) | The window calls itself "GST File Organizer v4.0" since `c258047` (`gui/pyqt_main_window.py:71`); `core/__init__.py:12` and `utils/__init__.py:38` still say "3.0.0", `power_query_extractor/__init__.py:3` "1.0.0", `gui/__init__.py:11` "4.0.0" since `c21f788`. |
| Folder tree with `gui/handlers/`, `gui/tabs/`, `gui/utils/`, `gui/widgets/` and no `models/` (`README.md:71-95`, the `gui/` sub-folders at `:82-86`) | `gui/handlers/`, `gui/tabs/`, `gui/utils/` deleted in `a62d3cc`; `models/`, `gui/views/`, `gui/controllers/`, `gui/theme/` added in `c258047`; `gui/widgets/` holds different files since `c258047`/`7e09e5c`. |
| "python launch_extractor.py … The extractor auto-loads target folder from cache … Click **Refresh & Extract**" (`README.md:166-176`) | Since `c258047` `launch_extractor.py` opens the one main window on tab 4 (`launch_extractor.py:13-19`); the button is "Start Extraction"; the target folder comes from the saved settings and a disk scan (`b0e7d0e`). The old extractor's cache auto-load never worked even before (it looked for a relative `gst_organizer_cache.json`, `power_query_extractor/gui/extractor_window.py:28@47a1492`, which the organiser never wrote — RAN (H-S)). |
| "Requirements.txt Contents: `openpyxl>=3.0.0`, `pywin32>=305; sys_platform == 'win32'`" (`README.md:104-108`) | `requirements.txt` never contained either line in B: it had `openpyxl>=3.1.2` and `lxml>=4.9.0` from `edac6a3`, added `PyQt5>=5.15.0` and a commented `# pywin32>=306` in `c258047` (`requirements.txt:5, 14`). |
| "Logs are saved to: `gst_organizer.log` … `pq_extractor.log`" (`README.md:289-291`) | Only `main.py` sets up a log (`main.py:15-23`); `pq_extractor.log` is configured only by the broken `extractor_main.py:12-19`. Since `b0e7d0e` there is also a per-run `Extraction_Log_<YYYYMMDD_HHMM>.txt` in the Level 1 folder, which the README does not mention. |
| "Recent Bug Fixes (Latest Commit)" naming `utils/helpers.py`, `file_handler.py`, `processing_handler.py` (`README.md:320-327`) | Those are `6237597`'s claims (shared); `file_handler.py` and `processing_handler.py` were deleted in `a62d3cc`. Several of the claims did not match even `6237597`'s own diff (12D). |

### 12C. Logic and Rule Evolution

This sub-section follows every business rule and every GST rule of project B (the registry's 144 rules, B-R1 to B-R144, `/root/merge-scratch/audit/B_registry.md`) through the 17 commits, and lists the rules that existed earlier but are gone. For every rule tagged **[GST RULE]** it says only what the code did and when that changed — never whether it is right. **⚠** marks values that change over time (return and table names, template cell positions, sheet names, the state list, dates) so the author can check them.

#### 12C.1 How the evolution was established

- RAN (me) `scripts/rule_blame.py` → `rule_blame.txt`: for each registry rule, every `file:line` range in its "Implemented at" column was run through `git blame` on `B/main`, and the commits that last touched those lines were counted. Blame shows only the *last* change of each line; earlier changes were taken from the diffs (READ: every B-only diff, and the shared diffs `6237597`, `1c48722`, `bd176bf`, `47a1492` for the files concerned) and from pickaxe searches (RAN (me), 12B.6).
- RAN (me) `filehist.txt`: `utils/constants.py` (file-name patterns ⚠, expected types ⚠, folder templates, template Links cell map ⚠, modes, state table ⚠) and `power_query_extractor/config/cell_mappings.py` (extraction cells) were **never touched by any B-only commit**; the core rule files `core/file_organizer.py`, `core/excel_handler.py`, `power_query_extractor/core/report_processor.py` and `data_consolidator.py` received only logging, error-handling and comment changes in B (READ diffs, 12B.3).
- RAN (me) `scripts/scan_trace.py` on extracted trees of `47a1492`, `c21f788`, `7e09e5c`, `b0e7d0e` (made-up files) for the scan-acceptance rules B-R1 to B-R4 (12B.6).
- Earlier runs by other readers are cited by their note names (H-S, H-B1, H-B2, LT, DC, PQ, REGISTRY-B).

#### 12C.2 The overall picture: which rules were foundational and which were discovered

| Kind of rule | Where it came from | What happened to it in B |
|---|---|---|
| File-name patterns ⚠, expected file types ⚠, completeness, client key, state-name → code table and fallback ⚠, Pvt/Ltd abbreviations, folder layout and names, copy/backup rules, template Links cells ⚠, report names, the Info!BB2 refresh check ⚠, the first 8 extraction cells ⚠, the consolidated layout, all Excel handling (B-R5–B-R60, B-R66–B-R75, B-R82–B-R110 except as noted; the GST-tagged ones among them are **[GST RULE]**, see 12C.4) | **Foundational**: present in the first commit `edac6a3` (31-08-2025, shared - before the split on 30-10-2025) and, by the `.pyc` evidence, mostly settled before git (June–July 2025). | **Unchanged in value.** B-only commits only moved some call sites into the new PyQt5 controllers (`c258047`) and trimmed logging. |
| Refresh wait time, refreshed-file suffix, per-report (ITC / Sales) selection (B-R61, B-R62, B-R65, B-R68, B-R70) | **Discovered in real use** (`bd176bf`, 30-10-2025, shared): fixed 10-second waits and a fixed `_Refreshed_` name did not suit every report; some clients needed only one report refreshed (HIGH by content, H-S). | Kept; the wait-time box widened from 5–60 to 5–120 seconds in `c258047` (B only (after the split)). The two places that *look for* refreshed copies were never updated for a custom suffix (B-R63, B-R136). |
| Six more extraction output sheets, 14 cells ⚠ (B-R76–B-R81) | **Added after the template grew** (`47a1492`, 30-10-2025, shared, the split point). | Unchanged in B. The pre-existing first-substring look-up (B-R83) has filled 9 of the 22 output columns from other cells ever since (RAN (H-S, LT, DC, REGISTRY-B)) — a technical mismatch; which cells are right is the author's call. |
| Scan acceptance of small or non-Excel files (B-R2–B-R4) | Foundational (`edac6a3`), but the files were thrown away silently. | **Modified** in `7e09e5c` (B only): the same thresholds now produce visible "variations"; Excel lock files `~$…` are skipped by name. |
| Settings, run history, session state, migration of old settings (B-R115–B-R120) | **New design** in `c258047` (B only): replaces the single v3 settings file. | Extraction run-history keys changed twice (`7e09e5c`, `b0e7d0e`). |
| Which run folder and which versions Step 4 uses (B-R39, B-R136) | Foundational in the separate Tk extractor (newest `Annual Statement-*` by modification time, one version per client by greatest name). | **Replaced three times in three days** (`c258047` → `7e09e5c` → `b0e7d0e`), ending with "greatest name" for the run folder and "every version that holds a report" (12C.3). |
| Screen rules (B-R129–B-R144) | Rewritten from scratch for PyQt5 in `c258047`, restyled toward the v3 look in `7e09e5c`, extended in `b0e7d0e`. | Several v3 screen rules were lost (12C.5). |

Pattern (HIGH, READ): **no GST rule *value* — file-name pattern ⚠, expected type ⚠, state code ⚠, template Links cell ⚠, extraction cell ⚠, sheet name ⚠ or the BB2 check ⚠ — was modified or removed in the whole history**; new values were only *added* (the 14 cells of `47a1492`). GST-tagged *behaviour* did change, as the register below shows, and only in these places: per-report selection in `bd176bf` (shared; B-R61, B-R70, B-R88), the unrecognised-file hint no longer reaching the summary workbook from `c258047` (B-R12), and a client-folder length floor of 5 in `c258047` that could cut the state code off (B-R32, floor 15 restored in `7e09e5c`). [Corrected during verification — originally stated without the word *value*, which contradicted the B-R61, B-R70, B-R88, B-R12 and B-R32 rows.] None of this says whether any rule is right; that is the author's call. Business and screen rules in B were changed mainly by the PyQt5 rewrite and by its two follow-up days, not by reported real-world problems — none of B's five messages describes a user-reported failure (READ the messages; 12D).

#### 12C.3 Every rule change in date order

| Date | Commit (label) | Rule(s) | Old → new (exact) | Evidence |
|---|---|---|---|---|
| 31-08-2025 | `edac6a3` (shared - before the split on 30-10-2025) | all foundational rules | introduced (12C.2 first row) | READ (H-S §1) |
| 29-10-2025 | `6237597` (shared) | long-client-name warning before processing (removed later, 12C.5) | warn when a selected client name has more than 10 characters → more than the "Client folder name max length" setting (default 35; 10 if the value is ≤0 or unreadable) (`gui/handlers/processing_handler.py:64-91@47a1492`); the per-client toggle kept `> 10` (`gui/handlers/client_handler.py:188@47a1492`) | READ diff |
| 29-10-2025 | `6237597` (shared) | B-R113 (scan progress) | an error raised by the progress callback aborted the scan → it is caught and logged as a warning (`core/file_parser.py:48-54`) | READ diff |
| 29-10-2025 | `1c48722` (shared) | B-R27 (status colours, Tk version) | every row coloured "complete" in light mode and "incomplete" in dark mode → Complete green, anything else orange (`gui/handlers/client_handler.py:51-55@47a1492`) | RAN (H-S `test_tags.py`) |
| 30-10-2025 | `bd176bf` (shared) | B-R68 | waits 10 s, then up to 5 one-second error checks, then 10 s → *w* s (box 5–60, default 10), up to 5 checks, then max(10, *w*) s (`power_query_extractor/core/report_processor.py:253-258, 280-285`) | READ diff |
| 30-10-2025 | `bd176bf` (shared) | B-R62 | refreshed copy `<stem>_Refreshed_<DDMMYY_HHMM>.xlsx` → `<stem><suffix>.xlsx`, suffix default `_Refreshed_{timestamp}` (`report_processor.py:91-94`) | READ diff |
| 30-10-2025 | `bd176bf` (shared) | B-R61, B-R70, B-R88 | both reports always processed → per client ITC and/or Sales; an unticked report gets status "Not selected for processing" (`report_processor.py:39-70`), which the consolidated workbook shows as ✗ and "Missing" | READ diff |
| 30-10-2025 | `bd176bf` (shared) | B-R63, B-R136 (later) | refreshed-copy *search* stays fixed to `_Refreshed_` (`report_processor.py:102`) while the *name* becomes configurable — the mismatch begins here | READ diff (H-S) |
| 30-10-2025 | `47a1492` (shared, split point) | B-R76–B-R81 **[GST RULE] ⚠** | 3 output sheets / 8 cells → 9 output sheets / 22 cells: "Reco Diff" `Reco`!C37/D37/E37; "Table 6J Match" `T6 of R9`!K29; "Table 8D Limits" `T8 of R9`!G6; "IMS Reco" `IMS_Reco`!AA1/AD1/AG1; "Extra In Purchase" `ExtraInPurchase`!O1/P1/Q1; "Extra In GSTR2B" `ExtraInGSTR2B`!S1/T1/U1 (`power_query_extractor/config/cell_mappings.py:105-228`) | READ diff |
| 30-10-2025 | `47a1492` (shared) | B-R83 **[GST RULE]** | the first-substring look-up (unchanged code since `edac6a3`) now meets one-word column names and fills the "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" IGST/CGST/SGST columns from `Diff GSTR-3B`!J15/K15/L15 ⚠ whenever those hold a value — a technical mismatch; which cells are right is the author's call | RAN (H-S, LT, DC, REGISTRY-B); re-run by the verifier (`consol.py`, made-up values: 9 of 22 columns differ from the configured cell) |
| 07-02-2026 | `c258047` (B only (after the split)) | B-R115–B-R119 | one file `~/.gst_organizer_cache.json` (8 keys + `recent_folders`) → folder `~/.gst_organizer/` with `settings.json` (11 keys), `run_history.json`, `session_state.json`; one-time import of the old file (`models/cache_manager.py:16-152`) | READ |
| 07-02-2026 | `c258047` (B only) | B-R118 | — (new) the import reads `recent_source_folders` (`models/cache_manager.py:122`) while the old writer stored `recent_folders` (`gui/utils/cache_manager.py:34-43@47a1492`), so the old recent list is never carried over | RAN (H-B1, LT) |
| 07-02-2026 | `c258047` (B only) | B-R32 (screen side) | Max length box 15–100 (validator 10–100), always enabled → 5–100, enabled only while "include client name" is ticked, but applied to the client folder name in every case (`gui/pyqt_main_window.py:203`, `core/file_organizer.py:79-83`); values 5–8 cut the state code off and can merge two clients' folders | RAN (H-B1 `create_client_state_key`) |
| 07-02-2026 | `c258047` (B only) | B-R65 | wait-time box 5–60 → 5–120 seconds (`gui/views/extraction_view.py:127`); the three Step 4 options now saved in settings (`models/app_state.py:21-23`) | READ |
| 07-02-2026 | `c258047` (B only) | B-R48 | template picker "Excel Template (`*.xltx *.xlsx`)" (`gui/handlers/file_handler.py:31, 42@47a1492`) → "Excel Files (`*.xlsx *.xlsm`)" (`gui/views/setup_view.py:612, 619`) | READ |
| 07-02-2026 | `c258047` (B only) | B-R95, B-R12 | summary workbook's Variations sheet received the scan's problem files with their "Expected Pattern" hints (`gui/handlers/processing_handler.py:161@47a1492`) → always an empty list (`gui/controllers/processing_controller.py:73`), so "No file variations found" | RAN (H-B1, DC T8) |
| 07-02-2026 | `c258047` (B only) | B-R44 (screen texts) | mode labels taken from `PROCESSING_MODES` ("Add new version folders to existing structure", "Continue previously interrupted processing") → "Re-Run - Overwrite existing files", "Resume - Skip already processed"; help "Re-Run: Overwrites files in existing structure", "Resume: Skips clients already processed". Behaviour unchanged. | READ `gui/views/setup_view.py:103-104, 177-179@c258047` |
| 07-02-2026 | `c258047` (B only) | B-R134 | clients start unticked; "Missing" lists the missing type names; the 7th column is a per-client Yes/No toggle → all clients ticked; "Missing" and "Extra" are counts; "Folder Name" shows the scan key; new "Last Run" column | READ (H-B1 Part 3) |
| 07-02-2026 | `c258047` (B only) | B-R135 | Dry Run first checked templates and target (`gui/handlers/processing_handler.py:25-26@47a1492`) → no checks (`gui/pyqt_main_window.py:326-343`); the Tk late-binding bug (wrong client names in the preview, since `6237597`) disappears with the Tk code | READ; RAN (H-S `test_dryrun.py`) |
| 07-02-2026 | `c258047` (B only) | B-R39, B-R86, B-R87, B-R120 | Step 4 list: found on disk by the Tk extractor → only the clients processed in this session, located by a name search over the whole target tree for `<client>-<state as written>`; consolidated report: Level 1 folder → target folder; row label: client folder name → `<client> - <state>`; extraction run-history key `<client> - <state>` (never matches the processing key) | RAN (H-B1 `t_build`) |
| 07-02-2026 | `c258047` (B only) | B-R124 | the window can now start only with pywin32 installed (12F.5) | RAN (me), 12B.7 |
| 07-02-2026 | `c258047` (B only) | B-R97 | openpyxl fallback: a Links cell that cannot be written (e.g. inside a merged range) was skipped silently → the whole report fails and the unfilled template copy stays under the report name (`core/excel_handler.py:246-252`) — reachable only without pywin32, which since this commit cannot start the window | RAN (H-B1); re-run by the verifier (`fallback.py`, a made-up template whose Links!A2:B2 is merged: at `47a1492` the call returned True and the other cells were written; at `b0e7d0e` it returned False and the copy stayed unfilled) |
| 07-02-2026 | `c258047` (B only) | B-R105 | the unused temp folder `<temp>/gst_excel_temp` was created at every program start (the Tk window made one `ExcelHandler` at start-up, `gui/main_window.py:112@47a1492`) → at every processing run (`gui/controllers/processing_controller.py:54`) | READ |
| 07-02-2026 | `c258047` (B only) | B-R125 | start-up created `core/ utils/ gui/ logs/ temp/` in the program folder and needed a console → neither (`main.py:12-13, 26-29`) | RAN (H-B1 pythonw simulation) |
| 07-02-2026 | `a62d3cc` (B only) | B-R130, B-R131, B-R45 | `GUI_CONFIG`, four `MESSAGES` entries and the `PROCESSING_MODES` names/descriptions lose their last readers (the Tk files, e.g. `gui/main_window.py:58-63@47a1492`, `gui/handlers/file_handler.py:123-139@47a1492`, `gui/tabs/setup_tab.py:216, 221@47a1492`) → dead | RAN (me) grep on extracted trees |
| 07-02-2026 | `a62d3cc` (B only) | B-R36 | the per-client include-name toggle (writer) deleted → the organiser's look-up always gets False | READ |
| 08-02-2026 | `7e09e5c` (B only) | B-R1, B-R2, B-R3, B-R4, B-R29 | files under 1,024 bytes or without the `PK\x03\x04` / `\xd0\xcf\x11\xe0` signature vanished silently → listed as variations "File too small to be valid Excel" / "Invalid or corrupted Excel file" / "Cannot read file: …"; `~$…` names dropped before any check; "total files" therefore no longer equals "Found N files" | RAN (me) `scan_trace.py` |
| 08-02-2026 | `7e09e5c` (B only) | B-R32 (screen side) | Max length 5–100 → 15–100 (`gui/views/setup_view.py:375`) — the v3 floor restored | READ diff |
| 08-02-2026 | `7e09e5c` (B only) | B-R35, B-R44, B-R141 | tick-box label "Include client name in folder names" → "Include client name in brackets for Level 4 folders" (`setup_view.py:369`); mode buttons back to the `PROCESSING_MODES` wording, typed by hand (`:343-359`); help card "Re-Run: Overwrite existing", "Resume: Skip already processed" (`:510-512`), which contradicts the buttons; note "Names >10 chars will show warning. This overrides individual client settings if checked." (`:385`) for features B does not have; expected-file-name list retyped ⚠ (`:492-499`) | READ diff |
| 08-02-2026 | `7e09e5c` (B only) | B-R39, B-R120, B-R136 | Step 4 list → every client in run history (latest record), "latest" version = first sub-folder in reverse **name** order; fallback search two levels deep; extraction results stored under the display name, creating phantom rows | RAN (H-B2 `t_hist7e`) |
| 08-02-2026 | `7e09e5c` (B only) | B-R113, B-R138 | scan progress connected for the first time in PyQt5 (`gui/pyqt_main_window.py:129, 238-242`); SCAN FILES disabled during a scan (`:231`) | READ diff |
| 09-02-2026 | `b0e7d0e` (B only) | B-R39, B-R136 | Step 4 tree from the disk: run folder = the `Annual Statement-*` sub-folder with the greatest **name** (`gui/pyqt_main_window.py:432-448`); every version folder holding an `ITC_Report_*.xlsx` or `Sales_Report_*.xlsx`; ticks start empty, only existing reports tickable; status "Refreshed (dd/mm/yy HH:MM)" if any `*_Refreshed*` file exists, else "Pending" | RAN (H-B2, REGISTRY-B: `…-311225 1000` chosen over `…-211026 0335`) |
| 09-02-2026 | `b0e7d0e` (B only) | B-R86, B-R87, B-R120, B-R122 | consolidated report: target folder → Level 1 folder (`gui/controllers/extraction_controller.py:92`); row label → `<client folder> / <version folder>`, one row per ticked version; new `Extraction_Log_<YYYYMMDD_HHMM>.txt` (`:109-121`); extraction run-history key → the same label | RAN (H-B2 `t_worker`) |
| 09-02-2026 | `b0e7d0e` (B only) | B-R29, B-R140 | scan statistics shown again (Tk had shown them, `gui/main_window.py:267-301@47a1492`); tab dots green/orange (`gui/pyqt_main_window.py:569-590`); window moves to Step 4 after processing (`:401`) | READ diff |

#### 12C.4 Rule-by-rule register (all 144 registry rules)

Columns: **First** = commit in which the rule's behaviour first appeared; **Changes since** = every later commit that changed the rule's behaviour, wording or location (comment-only or logging-only changes are named as such); **Blame** = commits that last touched the rule's lines at the snapshot, with line counts (RAN (me) `rule_blame.txt`); **At the snapshot** = state at `b0e7d0e` (detail in the registry). "Relocated" means the same logic was moved into the PyQt5 code by `c258047`. All `edac6a3`, `6237597`, `1c48722`, `1cb878b`, `bd176bf`, `47a1492` entries are "shared - before the split on 30-10-2025"; all `c258047`, `a62d3cc`, `c21f788`, `7e09e5c`, `b0e7d0e` entries are "B only (after the split)".

**Category 1 – Input-file recognition**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R1 | Top-level `*.xlsx`/`*.xls`/`*.xlsm` listing, sorted; listing error → empty | — | `edac6a3` (with a validation filter inside `find_excel_files`) | `7e09e5c`: filter removed, `~$` names dropped (`utils/helpers.py:259-261`); `b0e7d0e`: a third copy of the extension list at `gui/pyqt_main_window.py:469` | edac6a3 15, 7e09e5c 11, b0e7d0e 1 | live |
| B-R2 | Under 1,024 bytes → variation "File too small to be valid Excel" | — | `edac6a3` as a **silent drop** (`utils/helpers.py:231-233`) | `7e09e5c`: moved into `core/file_parser.py:181-191` and made visible | edac6a3 8, 7e09e5c 6 | live; old copy dead |
| B-R3 | 8-byte signature `PK\x03\x04` / `\xd0\xcf\x11\xe0` → else variation | — | `edac6a3` as a silent drop (`utils/helpers.py:235-243`) | `7e09e5c`: `core/file_parser.py:193-215`, visible; new issue "Cannot read file: …" | edac6a3 13, 7e09e5c 23 | live |
| B-R4 | `stat` failure → file skipped, log only | — | `edac6a3` (silent drop, ERROR "Error validating Excel file") | `7e09e5c`: WARNING "Cannot stat …" (`core/file_parser.py:174-179`); still invisible on screen | 7e09e5c 6 | live |
| B-R5 | Six patterns in fixed order, case-insensitive, whole name | [GST RULE] | `edac6a3` | `c258047`: Sales debug logs removed only | edac6a3 119 | unchanged |
| B-R6 | "GSTR-2B-Reco" pattern | [GST RULE] ⚠ | `edac6a3` | none (help text retyped in `c258047`, `7e09e5c`) | edac6a3 35, 7e09e5c 1 | unchanged |
| B-R7 | "ImsReco" pattern, 8 digits | [GST RULE] ⚠ | `edac6a3` | none (help text retyped) | edac6a3 35, 7e09e5c 1 | unchanged |
| B-R8 | "GSTR3B" pattern with optional "(n) " prefix | [GST RULE] ⚠ | `edac6a3` | none (help text retyped) | edac6a3 35, 7e09e5c 1 | unchanged |
| B-R9 | "Sales" pattern, four parts | [GST RULE] ⚠ | `edac6a3` | none (help text retyped) | edac6a3 35, 7e09e5c 1 | unchanged |
| B-R10 | "SalesReco" pattern | [GST RULE] ⚠ | `edac6a3` | none (help text retyped) | edac6a3 35, 7e09e5c 1 | unchanged |
| B-R11 | "AnnualReport" pattern | [GST RULE] ⚠ | `edac6a3` | none (help text retyped) | edac6a3 35, 7e09e5c 1 | unchanged |
| B-R12 | Hint for unrecognised names | [GST RULE] ⚠ | `edac6a3`; reached the user in the summary's Variations sheet ("Expected Pattern") at the split | `c258047`: the summary receives an empty list → hint never shown | edac6a3 20 | dead in practice since `c258047` |
| B-R13 | Pattern `category` / `folder` / `description` unused | [GST RULE] ⚠ | `edac6a3` | none | edac6a3 20 | dead since `edac6a3` |
| B-R14 | `EXCEL_EXTENSIONS` unused | — | `edac6a3` | `b0e7d0e`: another hand-typed copy (`gui/pyqt_main_window.py:469`) | edac6a3 8, b0e7d0e 1 | dead |
| B-R15 | `EXCEL_SIGNATURES` unused | — | `edac6a3` | `7e09e5c`: signatures typed again in the parser | edac6a3 6, 7e09e5c 2 | dead |
| B-R16 | `MAX_FILE_SIZE_MB = 100` unused | — | `edac6a3` | none | edac6a3 1 | dead |

**Category 2 – Client identity and name cleaning, and Category 3 – State codes (B-R21, B-R22, which the registry files under category 3)**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R17 | Client key `<client as written>-<state code>` | [GST RULE] | `edac6a3` (`core/file_parser.py:258-278`) | `c258047`: new readers — Step 2 "Folder Name" column and processing run-history key (in v3 the 7th column was a per-client toggle) | edac6a3 22, c258047 14 | live |
| B-R18 | Client/state name checks, warnings never shown | — | `edac6a3` | none | edac6a3 34 | dead in practice |
| B-R19 | "Private"→"Pvt", "Limited"→"Ltd", two different ways | — | `edac6a3` | none | edac6a3 7 | unchanged |
| B-R20 | File-name cleaning | — | `edac6a3` | `c258047`: report-name call relocated (`gui/controllers/processing_controller.py:112`) | edac6a3 49, c258047 1 | unchanged |
| B-R21 | State name → two-letter code table (39 names, 36 codes) | [GST RULE] ⚠ | `edac6a3` | `6237597`: relative import of the table (`utils/helpers.py:491`), no behaviour change; `c258047`: an unused `state_code` call left at `core/excel_handler.py:550` (extra log warnings) | edac6a3 60, 6237597 1, c258047 1 | unchanged |
| B-R22 | Fallback code for unknown states | [GST RULE] ⚠ | `edac6a3` | none in value; `c258047` adds more evaluations, so the WARNING repeats more often | edac6a3 26, 6237597 1 | unchanged |

**Categories 4–5 – Financial year and period, expected files and completeness**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R23 | No FY / period logic; captured period text never read | [GST RULE] ⚠ | `edac6a3` | none | edac6a3 18 | dead (write-only) |
| B-R24 | Six expected file types | [GST RULE] ⚠ | `edac6a3` | none | edac6a3 16 | unchanged |
| B-R25 | GSTR-3B any number of files; others "Multiple …" | [GST RULE] ⚠ | `edac6a3` | none | edac6a3 10 | unchanged |
| B-R26 | Completeness % of types, 1 decimal | [GST RULE] | `edac6a3` | none | edac6a3 4 | unchanged |
| B-R27 | Status Complete / Missing n / Has duplicates; row colours | [GST RULE] | `edac6a3` (status); Tk colours fixed `1c48722` | `c258047`: PyQt5 colours green #107C10 / orange #FF8C00 / red-orange #D83B01 (`gui/views/validation_view.py:387-393`); `7e09e5c` comment only | edac6a3 13, c258047 21, 7e09e5c 1 | live |
| B-R28 | `validate_results` with "critical" types ⚠ | [GST RULE] ⚠ | `edac6a3` | none | edac6a3 34 | dead since `edac6a3` |
| B-R29 | Scan statistics | — | `edac6a3` (shown by the Tk window, `gui/main_window.py:267-301@47a1492`) | `c258047`: no longer shown (4-line summary); `7e09e5c`: "total files" excludes the newly visible rejected files; `b0e7d0e`: shown again in the scan log and summary (`gui/pyqt_main_window.py:244-316`) | edac6a3 29, 7e09e5c 10, b0e7d0e 48 | live |

**Category 6 – Folder layout and naming**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R30 | Time stamp `DDMMYY HHMM`, taken twice | — | `edac6a3` | `c258047`: worker's stamp relocated (`gui/controllers/processing_controller.py:44`) | edac6a3 4, c258047 1 | unchanged |
| B-R31 | Level 1 `Annual Statement-<DDMMYY HHMM>` | — | `edac6a3` | `1cb878b`: `.gitignore:56`; `b0e7d0e`: prefix typed again for Step 4 (`gui/pyqt_main_window.py:434`) | edac6a3 23, 1cb878b 1, b0e7d0e 1 | unchanged |
| B-R32 | Level 2 client folder name, max length | [GST RULE] (state-code part) | `edac6a3` (`utils/helpers.py:513-537`, default 35) | `c258047`: box 5–100, greyed unless ticked, always applied; `7e09e5c`: 15–100 | edac6a3 36, c258047 10, 7e09e5c 1 | live |
| B-R33 | `FOLDER_STRUCTURE['level2']` unused | — | `edac6a3` | none | edac6a3 1 | dead |
| B-R34 | `Version-<DDMMYY HHMM>` per run | — | `edac6a3` | none | edac6a3 8 | unchanged |
| B-R35 | Three category folders, optional " (<client>)" | — | `edac6a3` | `c258047`: tick-box label "Include client name in folder names"; `7e09e5c`: back to "…in brackets for Level 4 folders" | edac6a3 35, c258047 2, 7e09e5c 1 | unchanged in behaviour |
| B-R36 | Per-client include-name override | — | `edac6a3` (organiser look-up by `<client>-<state code>`; the Tk toggle stored `<client>-<state as shown>`, a key mismatch, RAN (H-S)) | `a62d3cc`: the toggle deleted → always False | edac6a3 10, c258047 10 | dead |
| B-R37 | Which category folder each type goes to | [GST RULE] | `edac6a3` | none | edac6a3 74, c258047 2 | unchanged |
| B-R38 | Re-Run/Resume reuse the newest run folder by modification time | — | `edac6a3` | none | edac6a3 12 | unchanged |
| B-R39 | Step 4 run folder = greatest name; all versions with reports | — | Tk extractor `edac6a3`: newest by modification time (`power_query_extractor/gui/extractor_window.py:517@47a1492`), one version per client by greatest name (`:529`), `_`-folders skipped (`:525`) | `c258047`: whole-target name search; `7e09e5c`: run history + reverse-name version; `b0e7d0e`: present rule (`gui/pyqt_main_window.py:419-504`) | b0e7d0e 30 | live; disagrees with B-R38 |
| B-R40 | Per-client `_Organization_Report_*.txt` | — | `edac6a3` | `c258047`: call relocated (`processing_controller.py:132-133`) | edac6a3 64, c258047 2 | unchanged |
| B-R41 | `cleanup_on_error` never called | — | `edac6a3` | none | edac6a3 27 | dead |

**Category 7 – Copying and versioning**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R42 | Copy, size check, source never changed | — | `edac6a3` | none | edac6a3 69 | unchanged |
| B-R43 | Existing destination: Fresh backup / Resume skip / Re-Run overwrite | — | `edac6a3` | none | edac6a3 30 | unchanged |
| B-R44 | Modes fresh / rerun / resume | — | `edac6a3` (keys and behaviour) | `c258047`: new on-screen wording; `7e09e5c`: buttons back to the constants' wording, a contradicting help card | edac6a3 74, c258047 7, 7e09e5c 19 | behaviour unchanged since `edac6a3` |
| B-R45 | `PROCESSING_MODES` names/descriptions/icons unused | — | `edac6a3` (names and descriptions **used** by the Tk Setup tab, `gui/tabs/setup_tab.py:216, 221@47a1492`) | `a62d3cc`: last readers deleted → dead | edac6a3 17 | dead since `a62d3cc` |
| B-R46 | `get_unique_path` unused | — | `edac6a3` | none | edac6a3 20 | dead |
| B-R47 | `calculate_file_hash` unused | — | `edac6a3` | none | edac6a3 19 | dead |

**Category 8 – Report creation and template Links cells**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R48 | Two templates; existence checks only | — | `edac6a3` (picker `*.xltx *.xlsx`; checks `gui/handlers/file_handler.py:132-151@47a1492`) | `c258047`: picker `*.xlsx *.xlsm`, checks rewritten (`gui/pyqt_main_window.py:353-358`, per client `processing_controller.py:210-218`); `7e09e5c`: label "(.xlsx or .xlsm)" | edac6a3 2, c258047 39, 7e09e5c 33 | live |
| B-R49 | `TEMPLATE_EXTENSIONS` unused | — | `edac6a3` | none | edac6a3 1 | dead |
| B-R50 | Every client always gets both reports | [GST RULE] | `edac6a3` (Tk flow) | `c258047`: relocated (`processing_controller.py:116-130`) | c258047 17 | unchanged in behaviour |
| B-R51 | Report file names `ITC_Report_…` / `Sales_Report_…` | [GST RULE] | `edac6a3` (`gui/handlers/processing_handler.py:230-234@47a1492`) | `c258047`: relocated | edac6a3 2, c258047 21 | unchanged |
| B-R52 | Template copied to the final name, then filled | — | `edac6a3` | none | edac6a3 8 | unchanged |
| B-R53 | Cells go into a sheet named "Links" ⚠ | [GST RULE] ⚠ | `edac6a3` | none | edac6a3 25 | unchanged |
| B-R54 | ITC Links cells B2, B4/B5, B7/B8, B10/B11 ⚠ | [GST RULE] ⚠ | `edac6a3` | none | edac6a3 48 | unchanged |
| B-R55 | Sales Links cells B2/B3, B5/B6, B8/B9 ⚠ | [GST RULE] ⚠ | `edac6a3` | none (comment line touched by `c258047`) | edac6a3 42, c258047 2 | unchanged |
| B-R56 | How Links values are formed (first file, stem, skip empty) | [GST RULE] | `edac6a3` | `c258047`: the `[DEBUG]` dump of all values removed from the log | edac6a3 88, c258047 5 | unchanged |
| B-R57 | Cleaned name only after a successful copy | [GST RULE] | `edac6a3` | none | edac6a3 73, c258047 2 | unchanged |
| B-R58 | `get_relative_excel_path` unused | — | `edac6a3` | none | edac6a3 18 | dead |
| B-R59 | "Created" = output file exists | — | `edac6a3` | none | edac6a3 8 | unchanged |
| B-R60 | Client "Success" even when reports fail | — | `edac6a3` (Tk: `gui/handlers/processing_handler.py:354-360@47a1492`) | `c258047`: relocated; `7e09e5c`: two unused keys added to each entry | c258047 88, 7e09e5c 2 | unchanged in behaviour |

**Category 9 – Power Query refresh**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R61 | First `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx`; "Not selected" | [GST RULE] | `edac6a3` (glob) | `bd176bf`: per-report selection | edac6a3 26, bd176bf 25 | unchanged since `bd176bf` |
| B-R62 | Refreshed copy name with suffix | — | `edac6a3` (fixed `_Refreshed_<DDMMYY_HHMM>`) | `bd176bf`: configurable; `c258047`: default stored in settings and typed in more places; `b0e7d0e`: constructor lines touched | edac6a3 4, bd176bf 3, c258047 2, b0e7d0e 2 | live |
| B-R63 | Skip refresh looks only for `_Refreshed_` | — | `edac6a3` | `bd176bf`: name configurable, search not; `c258047`: the tick reaches the code by overwriting the shared option at every run (`gui/controllers/extraction_controller.py:53-55`) | edac6a3 21, c258047 5 | live, mismatch since `bd176bf` |
| B-R64 | `auto_refresh_if_missing` unused | — | `edac6a3` | none | edac6a3 5 | dead |
| B-R65 | Step 4 options: wait 5–120 (default 10), suffix, skip | — | `bd176bf` (Tk: 5–60) | `c258047`: 5–120, saved in settings; `7e09e5c`, `b0e7d0e`: restyle / signature | bd176bf 2, c258047 47, 7e09e5c 15, b0e7d0e 3 | live |
| B-R66 | Refresh Excel via `Dispatch`, visible, maximised, retries | — | `edac6a3` | none | edac6a3 36 | unchanged |
| B-R67 | SendKeys Alt+A, R, A; no `RefreshAll` | — | `edac6a3` | `c258047`: except narrowing only | edac6a3 36, c258047 1 | unchanged |
| B-R68 | Fixed waits per report | — | `edac6a3` (10 s + checks + 10 s) | `bd176bf`: *w* + checks + max(10, *w*) | edac6a3 45, bd176bf 8, c258047 1 | unchanged since `bd176bf` |
| B-R69 | Always save, close, quit | — | `edac6a3` | `c258047`: unreachable `return` removed (no behaviour change) | edac6a3 61, c258047 6 | unchanged |
| B-R70 | Refresh success = no dialog + BB2 ⚠ check passed | [GST RULE] | `edac6a3` | `bd176bf`: "Not selected" status; `c258047`: one "Result: ITC ✅/❌" line; `b0e7d0e`: "OK/FAIL" | edac6a3 76, bd176bf 3, c258047 6, b0e7d0e 3 | unchanged in logic |
| B-R71 | Info!BB2 check ⚠ | [GST RULE] ⚠ | `edac6a3` | none | edac6a3 48 | unchanged |
| B-R72 | "Open method … failed" branch never runs | — | `edac6a3` | none | edac6a3 16 | dead |

**Category 10 – Extraction cell mappings**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R73 | "GSTR-3B Difference" `Diff GSTR-3B`!J15/K15/L15 ⚠ | [GST RULE] ⚠ | `edac6a3` (same cells in the pre-git format of 21-06-2025) | none | edac6a3 24 | unchanged |
| B-R74 | "GSTR-1 Difference" `Diff - GSTR-1`!N1/O1/P1/Q1 ⚠ | [GST RULE] ⚠ | `edac6a3` (same cells pre-git) | none | edac6a3 29 | unchanged |
| B-R75 | "Turnover GSTR-1" `Diff - GSTR-1`!I1 ⚠ | [GST RULE] ⚠ | `edac6a3` (new compared with the pre-git format) | none | edac6a3 11 | unchanged |
| B-R76 | "Reco Diff" `Reco`!C37/D37/E37 ⚠ | [GST RULE] ⚠ | `47a1492` | none | 47a1492 23 | unchanged |
| B-R77 | "Table 6J Match" `T6 of R9`!K29 ⚠ | [GST RULE] ⚠ | `47a1492` | none | 47a1492 11 | unchanged |
| B-R78 | "Table 8D Limits" `T8 of R9`!G6 ⚠ | [GST RULE] ⚠ | `47a1492` | none | 47a1492 11 | unchanged |
| B-R79 | "IMS Reco" `IMS_Reco`!AA1/AD1/AG1 ⚠ (shows other cells) | [GST RULE] ⚠ | `47a1492` | none | edac6a3 4, 47a1492 23 | unchanged |
| B-R80 | "Extra In Purchase" `ExtraInPurchase`!O1/P1/Q1 ⚠ (shows other cells) | [GST RULE] ⚠ | `47a1492` | none | edac6a3 4, 47a1492 23 | unchanged |
| B-R81 | "Extra In GSTR2B" `ExtraInGSTR2B`!S1/T1/U1 ⚠ (shows other cells) | [GST RULE] ⚠ | `47a1492` | none | edac6a3 4, 47a1492 23 | unchanged |
| B-R82 | Mapping table → cell reads | [GST RULE] | `edac6a3` | none (22 mappings since `47a1492`) | edac6a3 93 | unchanged |
| B-R83 | First field name *containing* the column label wins | [GST RULE] | `edac6a3` | none in code; harmful since `47a1492` | edac6a3 14 | unchanged |
| B-R84 | Calculated columns (subtract/add/multiply/divide) | [GST RULE] | `edac6a3` | `c258047`: inner `except` narrowed, same results | edac6a3 49, c258047 1 | dead in practice |
| B-R85 | `FIELD_DISPLAY_NAMES` used only by a dead method | — | `edac6a3` | none | edac6a3 4 | dead |

**Category 11 – Consolidated output layout**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R86 | `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` and its folder | — | `edac6a3` (Tk extractor: Level 1 folder, `power_query_extractor/gui/extractor_window.py:808@47a1492`) | `c258047`: target folder; `b0e7d0e`: Level 1 folder again (`gui/controllers/extraction_controller.py:92`) | edac6a3 15, c258047 5, b0e7d0e 5 | live |
| B-R87 | Sheet layout; column A label | — | `edac6a3` (label = client folder name) | `c258047`: pasted duplicate header block removed (no output change); label `<client> - <state>`; `7e09e5c`: `<client> - <state>` from history; `b0e7d0e`: `<client folder> / <version folder>` | edac6a3 90, c258047 1 | live |
| B-R88 | ✓/✗ "Successfully Run"; "Missing" when not success | [GST RULE] | `edac6a3` | `bd176bf`: unticked reports also ✗/"Missing" | edac6a3 53 | unchanged |
| B-R89 | `#,##0.00` for non-zero numbers | [GST RULE] | `edac6a3` | none | edac6a3 3 | unchanged |
| B-R90 | Consolidated Summary sheet | — | `edac6a3` | none | edac6a3 28 | unchanged |
| B-R91 | Optional sheet description lines | — | `edac6a3` | none | edac6a3 15 | dead in practice |
| B-R92 | `show_timestamp` unused | — | `edac6a3` | none | edac6a3 1 | dead |
| B-R93 | Two older layouts unused | — | before git (dead already in `edac6a3`; smaller 25-06-2025 `.pyc`, RAN (H-S)) | none | edac6a3 162 | dead |
| B-R94 | `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` | — | `edac6a3` | `c258047`: a failed summary is silent (the Tk flow logged "Failed to create summary report") | edac6a3 39, c258047 11 | live |
| B-R95 | Processing summary content (5 sheets) | — | `edac6a3` (Variations sheet filled) | `c258047`: Variations always empty; `7e09e5c`: unused keys | edac6a3 190, c258047 63, 7e09e5c 2 | live, Variations empty |

**Category 12 – Excel handling and safety**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R96 | Step 3 Excel: attach to open Excel, hide, always Quit | — | `edac6a3` | `c258047`: except narrowing only | edac6a3 115, c258047 8 | unchanged |
| B-R97 | openpyxl fallback | — | `edac6a3` | `c258047`: fails on unwritable cell; unreachable through the window | edac6a3 63, c258047 2 | dead in practice since `c258047` |
| B-R98 | Step 3 open with 8.3 short-path retry | — | `edac6a3` | none | edac6a3 54 | unchanged |
| B-R99 | Step 4 open, three attempts | — | `edac6a3` | none (except narrowing) | edac6a3 109, c258047 1 | unchanged |
| B-R100 | 218-character path limit | — | `edac6a3` | none | edac6a3 18 | unchanged |
| B-R101 | Step 4 reading Excel (read-write, says read-only) | — | `edac6a3` | none (except narrowing) | edac6a3 36, c258047 3 | unchanged |
| B-R102 | Error-dialog detection | — | `edac6a3` | `bd176bf`: whitespace | edac6a3 70, bd176bf 3, c258047 3 | unchanged |
| B-R103 | Error-dialog clicking | — | `edac6a3` | none (except narrowing) | edac6a3 81, c258047 2 | unchanged |
| B-R104 | All Python warnings switched off | — | `edac6a3` | none | edac6a3 2 | unchanged |
| B-R105 | Unused `gst_excel_temp` folder | — | `edac6a3` (created at every start) | `c258047`: created at every processing run | edac6a3 2, c258047 1 | dead |
| B-R106 | `EXCEL_SAFETY` retries unused | — | `edac6a3` | none | edac6a3 4 | dead |
| B-R107 | `EXCEL_SAFETY` prefixes unused | — | `edac6a3` | `1cb878b`: `.gitignore:52-53` lists them | edac6a3 2, 1cb878b 2 | dead |
| B-R108 | `EXCEL_SAFETY` 10 KB minimum unused | — | `edac6a3` | none | edac6a3 1 | dead |
| B-R109 | `EXCEL_SAFETY` save options unused | — | `edac6a3` | none | edac6a3 6 | dead |
| B-R110 | `CoInitialize` / `CoUninitialize` inside the Excel functions | — | `edac6a3` | none | edac6a3 15, c258047 3 | unchanged |

**Category 13 – Performance and operational constants**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R111 | Scan, processing and extraction in background threads; same-kind second job refused | — | Tk `edac6a3`: processing and extraction in `threading.Thread`, scan on the window thread, no "already running" guard (H-S O5) | `c258047`: QThread for all three, guards, but jobs of different kinds can now run together (one window) | c258047 132, b0e7d0e 4 | live |
| B-R112 | Stop only between items | — | `edac6a3` | `c258047`: relocated | c258047 15, b0e7d0e 1 | unchanged |
| B-R113 | Progress percentages, ETA, scan progress | — | `edac6a3` (`ProgressTracker`) | `6237597`: callback errors caught; `c258047`: controllers (scan progress not shown); `7e09e5c`: scan progress shown, new messages | edac6a3 56, c258047 24, 7e09e5c 9 | live |
| B-R114 | 0.5 s waits in Step 3 | — | `edac6a3` | none | edac6a3 2 | unchanged |

**Category 14 – Settings, cache, logs and run history**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R115 | `~/.gst_organizer/` folder | — | `c258047` (before: one file `~/.gst_organizer_cache.json`) | none | c258047 9 | live |
| B-R116 | `settings.json` 11 keys, when saved | — | `c258047` (Tk saved on every browse and option change) | none | c258047 113 | live |
| B-R117 | `session_state.json` write-only | — | `c258047` | none | c258047 27 | dead (write-only) |
| B-R118 | One-time import of the v3 settings | — | `c258047` | none | c258047 40 | live, recent list lost |
| B-R119 | `run_history.json` | — | `c258047` | none | c258047 163 | live |
| B-R120 | Run-history keys | — | `c258047` | `7e09e5c`: extraction key = display name; `b0e7d0e`: `<client folder> / <version folder>` | edac6a3 1, c258047 55, b0e7d0e 1 | live, keys never match |
| B-R121 | `gst_organizer.log` in the start folder | — | `edac6a3` | `c258047`: about 50 `[DEBUG]` INFO lines gone (RAN (verifier): 32 in `core/excel_handler.py`, 4 in `core/file_parser.py`, 14 in the Tk `processing_handler.py`), so less client detail is logged; still Links values, Info!BB2 ⚠ text, figures | edac6a3 19 | unchanged in location |
| B-R122 | `Extraction_Log_<YYYYMMDD_HHMM>.txt` | — | `b0e7d0e` | none | b0e7d0e 20 | live |
| B-R123 | Text cut-offs in logs | — | `edac6a3` | none | edac6a3 7 | unchanged |

**Category 15 – Start-up and dependencies**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R124 | Declared vs needed dependencies | — | `edac6a3` (openpyxl, lxml; pywin32 absent) | `c258047`: PyQt5 added, `# pywin32>=306` comment, pywin32 needed to start; `c21f788`: eager `gui/__init__.py` | edac6a3 32, c258047 8, c21f788 1 | live gap |
| B-R125 | `main.py` start-up | — | `edac6a3` | `5ea77c2`: shebang; `c258047`: PyQt5, guarded console wrapper, no folder creation | edac6a3 14, c258047 21 | live |
| B-R126 | Window start-up order | — | `c258047` | `7e09e5c`: Extract tab filled from history at start; `b0e7d0e`: disk scan at start on the window thread | c258047 49, b0e7d0e 2 | live |
| B-R127 | `launch_extractor.py` | — | `edac6a3` (launches the Tk extractor) | `c258047`: opens the main window on tab 4 | edac6a3 13, c258047 15 | live |
| B-R128 | Old extractor entry point | — | `edac6a3` | `a62d3cc`: broken by deleting its window | edac6a3 34, 1cb878b 2 | dead and broken |

**Category 16 – User-interface behaviour rules**

| ID | Rule (short) | Tags | First | Changes since | Blame | At the snapshot |
|---|---|---|---|---|---|---|
| B-R129 | Window title, size, tabs, menus, version labels | — | `c258047` (Tk window from `GUI_CONFIG`, "v3.0", 1200×800) | `c21f788`: `gui/__init__.py` "4.0.0" | edac6a3 3, c258047 49, c21f788 1 | live |
| B-R130 | `GUI_CONFIG` unused | — | `edac6a3` (used by the Tk window) | `a62d3cc`: dead | edac6a3 26 | dead |
| B-R131 | `MESSAGES` unused | — | `edac6a3` (four error texts used by Tk) | `a62d3cc`: dead | edac6a3 21 | dead |
| B-R132 | Theme leftovers (`COLORS`, `#info`, `toggle`) | — | `c258047` (`toggle`) | `7e09e5c`: `COLORS`, `#info` | c258047 4, 7e09e5c 31 | dead |
| B-R133 | Input checks and their messages | — | Tk `edac6a3` (`gui/handlers/file_handler.py:120-151@47a1492`) | `c258047`: new texts; `7e09e5c`: "Could not find output folders…"; `b0e7d0e`: Step 4 checks rewritten | c258047 33, 7e09e5c 5, b0e7d0e 20 | live |
| B-R134 | Step 2 table rules | — | `c258047` (Tk rules differed, 12C.5) | `7e09e5c`: restyle | c258047 60, 7e09e5c 16 | live |
| B-R135 | Dry run | — | Tk `edac6a3` (checked inputs first) | `6237597`: late-binding bug (Tk); `c258047`: PyQt5 version without checks | c258047 18 | live |
| B-R136 | Step 4 tree and "Refreshed" status | — | `c258047` (flat list) | `7e09e5c`: version per row; `b0e7d0e`: tree, status from files | c258047 59, 7e09e5c 25, b0e7d0e 196 | live |
| B-R137 | HTML logs, colours | — | `c258047` (Tk logs were time-stamped plain text, `gui/utils/ui_helpers.py:17-35@47a1492`) | `7e09e5c`: scan log | edac6a3 8, c258047 35, 7e09e5c 19 | live |
| B-R138 | Enable / disable rules | — | `c258047` (Tk buttons fired even when disabled, H-S O5) | `7e09e5c`: SCAN disabled during scan | c258047 37, 7e09e5c 5 | live |
| B-R139 | Completion and error feedback | — | `c258047` (Tk: folder opened on Windows, macOS and Linux; extractor asked "Open the consolidated report?") | `7e09e5c`; `b0e7d0e`: jump to Step 4 | c258047 55, 7e09e5c 2, b0e7d0e 13 | live |
| B-R140 | Tab dots | — | `b0e7d0e` | none | c258047 2, 7e09e5c 1, b0e7d0e 32 | live |
| B-R141 | Help texts (expected file names ⚠) | [GST RULE] ⚠ | Tk `edac6a3` (help column) | `c258047`: help panel; `7e09e5c`: cards and note; `b0e7d0e`: cards fold at each scan | c258047 2, 7e09e5c 292, b0e7d0e 10 | live |
| B-R142 | Size and time display formats | — | `edac6a3` | `c258047`: run-history times; `b0e7d0e`: `dd/mm/yy HH:MM` | edac6a3 14, c258047 5, b0e7d0e 1 | live |
| B-R143 | `truncate_path` unused | — | `edac6a3` | none | edac6a3 23 | dead |
| B-R144 | Light/dark style sheet | — | `c258047` (Tk recursive recolourer before) | `7e09e5c`: inline light styles defeat dark mode on Steps 1–2 | c258047 336, 7e09e5c 74 | live, partial |

#### 12C.5 Rules that existed earlier but are no longer in the code

Each existed at the split point `47a1492` (shared - before the split on 30-10-2025) unless stated, and was lost in B's own commits. None of these is judged; where it touches a GST rule it is tagged.

| Rule that existed | Where it was | Lost in (label) | What is left at the snapshot |
|---|---|---|---|
| Warn before processing when a selected client name is longer than the max-length setting (since `6237597`; before: 10 characters) and ask "Continue anyway?" | `gui/handlers/processing_handler.py:64-91@47a1492` | not ported in `c258047`; deleted with the file in `a62d3cc` (B only (after the split)) | only the note "Names >10 chars will show warning…" (`gui/views/setup_view.py:385`), added back as text in `7e09e5c` |
| Per-client "include client name in folder names" toggle (column 7 of the Step 2 tree) with a >10-character warning and "Global Setting Disabled" message | `gui/handlers/client_handler.py:175-219@47a1492` | `a62d3cc` (B only) | the plumbing `client_folder_settings` (always empty, `models/app_state.py:83`) and the organiser's look-up (`core/file_organizer.py:117-126`) — B-R36 |
| Export the ticked clients to a text file and open it | `gui/handlers/client_handler.py:221-253@47a1492` | `a62d3cc` (B only) | an EXPORT LIST button connected to nothing (`gui/views/validation_view.py:186-221`; RAN (MW e2): 0 receivers) |
| Summary workbook lists the scan's problem files with an "Expected Pattern" hint | `gui/handlers/processing_handler.py:161@47a1492` | `c258047` (B only) | Variations sheet always "No file variations found" (`gui/controllers/processing_controller.py:73`) — B-R95, B-R12 [GST RULE] ⚠ (the hint names expected return-file patterns) |
| Dry Run first checks templates and target | `gui/handlers/processing_handler.py:25-26@47a1492` | `c258047` (B only) | Dry Run without checks (`gui/pyqt_main_window.py:326-343`) |
| "No valid GST files found" warning instead of moving on when a scan finds no client | `gui/handlers/file_handler.py:89-91@47a1492` | `c258047` (B only) | the window always moves to Step 2 and turns Step 1 green (`gui/pyqt_main_window.py:304-307`) |
| Clients start unticked; "Missing" lists the type names | `gui/handlers/client_handler.py:31, 35, 46@47a1492` | `c258047` (B only) | all ticked; counts only (`gui/views/validation_view.py:363-421`) |
| Max-length box always editable; floor 15 (validator 10) | `gui/tabs/setup_tab.py:252-279@47a1492` | `c258047` (floor 5, greyed unless ticked); floor 15 restored in `7e09e5c` (B only) | 15–100, greyed unless ticked, always applied (`gui/views/setup_view.py:374-382`) — B-R32 [GST RULE] (state-code part) |
| Template picker offering `.xltx` | `gui/handlers/file_handler.py:31, 42@47a1492`; label "(.xlsx or .xltx)" `gui/tabs/setup_tab.py:121@47a1492` | `c258047` (B only) | `.xlsx` / `.xlsm` (`gui/views/setup_view.py:221, 612, 619`) |
| "Processing is still running. Stop and exit?" on closing | `gui/main_window.py:314-325@47a1492` | `c258047` (B only) | window closes at once (`gui/pyqt_main_window.py:601-608`) |
| Open the output folder on Windows, macOS and Linux | `gui/handlers/processing_handler.py:442-449@47a1492` | `c258047` (B only) | Windows only (`gui/pyqt_main_window.py:404-410`) |
| Settings saved at every Browse and option change | `gui/handlers/file_handler.py:19-55@47a1492` | `c258047` (B only) | saved at scan, processing complete, dark-mode toggle, File > Save Settings, close (B-R116) |
| Recent source folders kept in the settings file | `gui/utils/cache_manager.py:34-43@47a1492` | `c258047` (B only) | written to `session_state.json`, never read; the old list is not imported (B-R117, B-R118) |
| Per-file progress and per-file log lines during processing | `gui/handlers/processing_handler.py:203@47a1492` (`file_progress`) | `c258047` (B only) | one line per client; `organize_files` called without a callback (`gui/controllers/processing_controller.py:106`) |
| Time stamps on every log line "[HH:MM:SS]" | `gui/utils/ui_helpers.py:17-20@47a1492` | `c258047` (B only) | no time stamps (B-R137) |
| Start-up hint "Templates must have 'Links' sheet" ⚠; processing hint "Check if template has 'Links' sheet" ⚠ | Tk main window and processing handler (`gui/main_window.py:363@47a1492`; `gui/handlers/processing_handler.py:260, 334@47a1492`; READ (H-B1) at `c258047`, line numbers added by the verifier) | `c258047` / `a62d3cc` (B only) | none |
| Tk extractor: run folder = newest by modification time; one version per client; `_`-prefixed client folders skipped; status per report type from `*_Report_*_Refreshed_*` files shown as `YYYY-MM-DD HH:MM`; report box disabled when the report is missing; per-report result lines; "Open the consolidated report?" | `power_query_extractor/gui/extractor_window.py:517, 525, 529, 606, 616, 638-661, 808-820@47a1492` | `a62d3cc` (B only) (file deleted); partly rebuilt differently in `b0e7d0e` | run folder by name, all versions, no `_` skip, one status per version from any `*_Refreshed*` file as `dd/mm/yy HH:MM` (B-R39, B-R136); no prompt |
| Wait-time box maximum 60 seconds | `power_query_extractor/gui/extractor_window.py:182-194@47a1492` | `c258047` (B only) | 120 seconds (B-R65) |
| Unused temp folder created at every program start | `gui/main_window.py:112@47a1492` → `core/excel_handler.py:52-53` | `c258047` (B only) | created at every processing run instead (B-R105) |
| Program creates `core/ utils/ gui/ logs/ temp/` and `__init__.py` files in its own folder at every start | `main.py:29-43@47a1492` | `c258047` (B only) | none |
| Every row one colour in Tk (bug) | `gui/handlers/client_handler.py:52-61@edac6a3` (the overriding block is 54-61) | fixed by `1c48722` (shared) | PyQt5 three-colour rule (B-R27) |

Rules removed before the split (for completeness): the old mapping file `cell_mappings copy.py` (deleted `47a1492`, shared); the fixed 10-second waits and fixed `_Refreshed_` name (replaced `bd176bf`, shared); the single client tick box per client in the Tk extractor (replaced by two per-report boxes, `bd176bf`, shared).

#### 12C.6 Where the older document (README.md) disagrees with the rule history

| README says (line) | Rule history shows |
|---|---|
| Output names `GST_Processing_Summary_DDMMYYYY_HHMMSS`, `ITC_Report_ClientName_State_DDMMYYYY_HHMMSS` (`README.md:182-203`, from `d674386`) | Since `edac6a3` the code uses `DDMMYY_HHMM` and the **state code** (B-R30, B-R51); no commit ever used the README's format (READ, H-S). |
| "All file operations use temporary files and verification steps" (`README.md:7, 42`); "Backups created for existing files" (`:307`) | `EXCEL_SAFETY` (temp prefix, verification size, retries) was never read at any commit (B-R106–B-R109); the temp folder was never used (B-R105); backups only in Fresh mode for a destination that already exists in the same Version folder, which practically never happens (B-R43). |
| Processing modes "Re-Run: Add new versions", "Resume: Continue interrupted processing" (`README.md:142-145`) | Matches the `PROCESSING_MODES` texts and the snapshot's buttons, but not the snapshot's help card (`gui/views/setup_view.py:510-512`) nor the behaviour: Resume has always created a new Version folder, so it skips nothing unless repeated within the same minute (B-R44; RAN (LT, DC T9)). |
| "Review variations tab" (`README.md:156`); summary "Variations" sheet (`:200`) | There never was a variations tab in any version; the summary's Variations sheet lost its content in `c258047` (B-R95). |
| "Resolved thread safety issues in processing_handler.py", "Corrected client name validation logic indentation" (`README.md:324-325`, in the list at `:320-327`; [Corrected during verification — the first quotation originally read "Fixed thread safety issues", which is `6237597`'s commit-message wording, not the README's]) | `6237597`'s diff did not contain an indentation fix and its thread change introduced a wrong-name bug in the Tk Dry Run (RAN (H-S)); both concern Tk files deleted in `a62d3cc`. |
| Templates ".xlsx or .xltx" (`README.md:120-121`) | `.xltx` was offered by the Tk picker until `c258047`; B's picker offers `.xlsx` / `.xlsm`, and reports are always saved as `.xlsx` (B-R48, B-R51). |

#### 12C.7 What the rule history tells a rebuilder, and what only the author can confirm

- Rules that never moved in 17 commits — the six file-name patterns ⚠, the expected types ⚠, the state table ⚠, the Links cells ⚠, the BB2 check ⚠ and the 22 extraction cells ⚠ — are the stable heart of the tool. They are also the rules that carry ⚠ values; their unchanged state proves only that no one *changed* them, not that they still match today's templates and returns. **[GST RULE]**: the author must confirm each value.
- Rules that changed three times in three days (Step 4 run-folder choice, extraction run-history keys, the Extract list's source, the mode texts) show where B's rewrite was still unsettled at the snapshot; a rebuild should not copy any single one of them without the author choosing.
- Questions for the author (only the author can answer):
  1. Should the long-client-name warning, the per-client include-name choice and the client-list export come back (all lost in `a62d3cc`)?
  2. Which descriptions of Re-Run and Resume are the intended behaviour (constants/buttons, the `c258047` texts, or the snapshot's help card)? The code has behaved the same way since `edac6a3`.
  3. Which run folder should Step 4 use — newest by modification time (v3 extractor, and B's own Re-Run/Resume), greatest name (snapshot), the one just processed, or one the user picks?
  4. **[GST RULE]** Are the 22 extraction cells ⚠ and the template Links cells ⚠ still where today's templates keep those figures, and which cells should "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" show (B-R79–B-R81, B-R83)?
  5. Should problem files be listed (B since `7e09e5c`) or silently skipped (B before, and A)?

### 12D. Edge Cases Discovered in Production

This sub-section takes every commit in B's history that fixes a bug or handles a situation met in real use — including fixes hidden inside commits whose messages describe something else — and records what broke, what input or scenario caused it, how it was fixed (a *patch*, which removes the symptom, or a *proper fix*, which removes the cause), and whether similar cases are still unhandled at the snapshot `b0e7d0e` (with `file:line`). "Real use" is inferred: no commit message in B describes a user-reported failure (READ all 17 messages), so each case says how sure the inference is.

> **Important:** two of B's five own commits introduced start-up failures or regressions that later commits had to repair (`a62d3cc` → `c21f788`; `7e09e5c` → `b0e7d0e`), and `c258047` removed several v3 safeguards without saying so (12C.5). In B, "bug fix" commits are as much repairs of the rewrite as responses to real-world files.

#### 12D.1 Summary table

| # | Commit (date, label) | Kind | What broke / what situation | Fix | Similar cases still open at the snapshot |
|---|---|---|---|---|---|
| D1 | `cbfa4a5` + `9c56a1c` (31-08-2025 / 29-10-2025, shared - before the split on 30-10-2025) | data hygiene | The run log `gst_organizer.log` (client names, file names, folder paths) was committed in the first upload | Deleted from the tip (patch: still in history at `edac6a3`) | yes (D1 detail) |
| D2 | `1cb878b` + `01862a0` (29-10-2025, shared) | data hygiene | 70 compiled files with the developer computer's folder paths were committed | `.gitignore` + removal (patch: still in history) | partly |
| D3 | `5ea77c2` (29-10-2025, shared) | tidy-up | Malformed first line `# !/usr/bin/env python3` in `main.py` | removed (no behaviour effect) | — |
| D4 | `6237597` (29-10-2025, shared) | claimed bug fixes | Message claims six fixes; the diff contains two real changes, three changes with no effect, one claim with no change, and one new bug | mixed (D4 detail) | yes |
| D5 | `1c48722` (29-10-2025, shared) | real bug fix (message says "redundant code") | Tk client list: every row "complete"-coloured in light mode, "incomplete"-coloured in dark mode | proper fix | colour key on screen incomplete (D5) |
| D6 | `4c42758` (29-10-2025, shared) | usability, met in real use (MEDIUM) | Mouse wheel did not scroll the Tk extractor's client list | wheel binding added | not applicable in PyQt5 (D6) |
| D7 | `bd176bf` (30-10-2025, shared) | real-use edge cases (HIGH) | Fixed 10-second refresh waits did not fit every report; only one of ITC/Sales sometimes needed; refreshed copies needed another name; last refresh time wanted | configurable wait and suffix, per-report ticks, status column — a patch for the waiting problem | yes, and it created the suffix mismatch (D7) |
| D8 | `47a1492` (30-10-2025, shared, split point) | template grew (not a fix) | New ITC check sheets needed in the consolidated report | 14 cells added ⚠ — exposed the first-substring look-up | yes, unfixed (D8) |
| D9 | `c258047` (07-02-2026, B only (after the split)) | rewrite with embedded fixes | Start-up crash without a console; program writing folders into its own folder; window freezing during scans; Tk thread problems | proper fixes for these four | yes, and the rewrite introduced new gaps (D9) |
| D10 | `a62d3cc` (07-02-2026, B only) | regression | Deleting the Tk files left three dangling imports: the program could not start | — (introduced the failure) | — |
| D11 | `c21f788` (07-02-2026, B only) | real fix of D10 (HIGH) | Start-up failure "No module named 'gui.main_window'" | one-line patch | `power_query_extractor/extractor_main.py:9` still broken (D11) |
| D12 | `7e09e5c` (08-02-2026, B only) | unannounced fixes | Dangling `gui.widgets` imports; Extract tab unable to find clients whose state was spelled in full, and empty after a restart; scan progress never shown; SCAN re-entry; damaged files vanishing silently; Excel lock files counted | mostly proper; the Extract-tab fix was a patch replaced next day | yes (D12) |
| D13 | `b0e7d0e` (09-02-2026, B only) | fix of D12's patch | Extract list from run history: wrong "latest" version by name order, phantom rows after an extraction, empty on another PC | rebuilt from the disk — proper for those three | the same name-order problem moved to the run folder (D13) |

#### 12D.2 Each case in detail

**D1 — run log committed (`cbfa4a5`, merged into the local line by `9c56a1c`; shared - before the split on 30-10-2025).**
- *What broke:* `edac6a3` committed `gst_organizer.log` (5,689 lines). The file is described here by kind only: client names, state names, source file names and full Windows folder paths from runs between 23-07-2025 and 31-08-2025; no tax figures (RAN (H-S), counts and line shapes only).
- *Scenario:* `main.py` writes the log relative to the folder the program is started from (then `main.py:21-28@edac6a3`, now `main.py:15-23`); the author ran the program from the project folder and there was no `.gitignore` yet.
- *Fix:* deleted on GitHub's web page 2 minutes 13 seconds after the first commit's time stamp (22:11:56 → 22:14:09, RAN `git log --date=iso`); the local line received the deletion only on 29-10-2025 through the merge. **Patch**: the file stays readable with `git show edac6a3:gst_organizer.log` in both projects' histories (HIGH).
- *Similar cases at the snapshot:* `gst_organizer.log` is still written to the current folder (`main.py:15-23`; B-R121), now ignored by `.gitignore:39-40` when that folder is the repository; it still receives client folder names, written Links values, Info!BB2 ⚠ text and every extracted figure (B-R121, B-F32). Since `c258047` extracted figures are also kept indefinitely in `~/.gst_organizer/run_history.json` (`models/run_history.py:73-95`, B-R119); since `b0e7d0e` every extraction writes `Extraction_Log_<YYYYMMDD_HHMM>.txt` into the run folder (`gui/controllers/extraction_controller.py:109-121`). None of these is a repository leak, but each is client data at rest that the author should know about.

**D2 — compiled files committed (`1cb878b` + `01862a0`, shared).** 70 `.pyc` files (Python 3.12 and 3.13) were tracked from `edac6a3`; `.gitignore` (`1cb878b`, `.gitignore:2-3`) and the removal (`01862a0`) stopped it. **Patch** in the same sense as D1: the files, with the developer computer's folder paths embedded, remain in history (RAN (H-S), paths not reproduced). Similar today: none in the repository (RAN (me) `sizes.txt`: 0 `.pyc` from `01862a0` on). `.claude/settings.local.json` (an AI-tool permission file, no secrets) was committed in `bd176bf` and removed in `c258047` with `.claude` added to `.gitignore:81-82` — the same kind of slip, caught the same way.

**D3 — malformed shebang (`5ea77c2`, shared).** The first line `# !/usr/bin/env python3` had a space after `#`, so it never worked; it was removed. No behaviour change; the program is started with `python main.py` (README.md:112).

**D4 — "Fix critical import and threading bugs" (`6237597`, shared).** What the diff actually did (READ diff; RAN (H-S)):

| Claim in the message | What the diff does | Verdict | At the snapshot |
|---|---|---|---|
| "Improved progress_callback None handling in file_parser.py" | stores the callback and calls it inside try/except (`core/file_parser.py:29, 48-54`): an error in the progress display no longer aborts the scan | **real, proper fix** (the old code already checked for None) | present, unchanged |
| "Fixed thread safety issues in processing_handler.py GUI updates" | wraps part of the Tk worker's screen updates in `root.after(0, lambda: …)` | **partial, and a regression**: the lambdas read loop variables later, so the Tk Dry Run printed the last client's name and file for every line (RAN (H-S) `test_dryrun.py`) | gone: the Tk code was deleted in `a62d3cc`; the PyQt5 code calls the log directly and has no late-binding lambda (RAN+READ (DC §4.1)) |
| "Fixed indentation bug in long client name validation logic" | removed trailing spaces only; the real change is the threshold: 10 characters → the max-length setting (12C.3) | no indentation bug existed | the warning itself was lost in `a62d3cc` (12C.5) |
| "Fixed circular import error in utils/helpers.py (STATE_CODE_MAPPING)" | `from utils.constants import …` → `from .constants import …` (`utils/helpers.py:491`) | no effect (same module either way) | present; the comment "Import here to avoid circular import" (`utils/helpers.py:490`) describes a problem that does not exist (DC L-09) |
| "Fixed scanned_files attribute access error in file_handler.py (line 75)" | `len(self.app.file_parser.scanned_files)` → `len(self.app.scanned_files)` — the same object | no effect | gone with the Tk file |
| "Corrected non-existent app.py reference in main.py error messages" | `main.py` is **not in the diff** | claim false | made true only by `c258047`'s rewrite of `main.py` |

*Scenario (inferred, MEDIUM):* an AI-assisted code review on the author's first return to the code after two months, not observed crashes (RAN (H-S): no code path at `edac6a3` crashed on the listed points). *Similar cases at the snapshot:* none of the threading kind (workers only emit signals, DC §4.8); the README still repeats these claims as "Recent Bug Fixes (Latest Commit)" (`README.md:320-327`).

**D5 — wrong row colours (`1c48722`, shared).** *What broke:* in `edac6a3` a second `if/else` in `update_client_tree` overrode the status colour: light mode tagged every row "complete" (incomplete clients looked complete), dark mode every row "incomplete" (RAN (H-S) `test_tags.py`). *Scenario:* any scan with a mix of complete and incomplete clients. *Fix:* the overriding block removed — **proper fix** (Tk). *At the snapshot:* the PyQt5 list colours rows green (Complete), orange (status contains "Missing") or red-orange (anything else, i.e. "Has duplicates") (`gui/views/validation_view.py:387-393`), but the on-screen key mentions only two colours ("green = complete, orange = missing files", `gui/views/validation_view.py:72-115`; B-R141), and the row colours are fixed foreground colours whatever the theme (B-R144).

**D6 — mouse wheel in the extractor (`4c42758`, shared).** *What broke:* the Tk extractor's client list (a canvas of tick boxes) and its log box did not scroll with the wheel when there were many clients or log lines (`power_query_extractor/gui/extractor_window.py:216-235, 330-337@4c42758`; HIGH by content; MEDIUM that it was met with a real client count). *Fix:* `<MouseWheel>` bindings with Windows-style steps (`delta/120`) — fine on Windows; macOS reports much smaller wheel steps and Linux uses separate Button-4/5 events, so these bindings would scroll little or not at all there (MEDIUM, general Tk knowledge, not tested) — [Corrected during verification — originally "Windows/macOS style" and only the client list]. *At the snapshot:* not applicable; the Tk window was deleted in `a62d3cc`, and the PyQt5 views scroll natively (scroll areas `gui/views/setup_view.py:95-146`, `gui/views/validation_view.py:25-68`; the Step 4 list is a tree view).

**D7 — refresh timing and per-report choice (`bd176bf`, shared).** *What broke / situation (HIGH by content):* the keystroke refresh waited a fixed 10 s, then up to 5 one-second dialog checks, then another fixed 10 s; some Power Query refreshes need longer, so the copy could be saved before the refresh finished; some clients needed only the ITC or only the Sales report refreshed; the author wanted a different name for refreshed copies and to see when each report was last refreshed. *Fix:* wait *w* seconds (default 10) and then max(10, *w*) (`power_query_extractor/core/report_processor.py:253-258, 280-285`); suffix pattern (`:91-94`); "Not selected for processing" (`:39-70`). This is a **patch**: it tunes a guess instead of detecting when the refresh has finished (no `RefreshAll`, no `BackgroundQuery = False`, no `CalculateUntilAsyncQueriesDone`, no check of the connections' `Refreshing` flag — B-R67, B-R68; Standard 12 describes these, 12F.7). *Similar cases at the snapshot:*
- the suffix is honoured when *making* the copy but not when *looking for* copies: Skip refresh still searches `<stem>_Refreshed_*.xlsx` (`report_processor.py:102`, B-R63), and since `b0e7d0e` the Step 4 status searches `*_Refreshed*` (`gui/pyqt_main_window.py:468`, B-R136) — RAN (DC T6, LT §4.8): a custom suffix makes Skip refresh refresh again and the status stay "Pending";
- an empty suffix gives a copy with the same name as the original → "same file" error → report failed (`report_processor.py:91-94, 118-120`; RAN (PQ, XL E5)); a suffix without `{timestamp}` overwrites the previous copy (B-R62);
- the report search `ITC_Report_*.xlsx` also matches earlier refreshed copies (`report_processor.py:76-89`, B-R61; RAN (DC T6) on Linux; MEDIUM on Windows NTFS);
- the wait box became 5–120 seconds in `c258047` (`gui/views/extraction_view.py:127`); Qt's spin box cannot hold a non-number, so the Tk crash on a non-numeric entry (H-S, bd176bf item 3) cannot occur in B (READ).

**D8 — six new extraction sheets (`47a1492`, shared, the split point).** *Situation:* the ITC template gained check sheets (Reco, T6 of R9, T8 of R9, IMS_Reco, ExtraInPurchase, ExtraInGSTR2B ⚠) and the author wanted their cells ⚠ in the consolidated report. *What it exposed:* the consolidator's look-up takes, for each output column, the first extracted field whose name *contains* the column label (`power_query_extractor/core/data_consolidator.py:130-133`) before trying the exact name (`:135-138`); the new one-word labels "IGST", "CGST", "SGST" match `GSTR-3B_Difference_Diff_IGST/CGST/SGST` first, so the "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" sheets show the `Diff GSTR-3B`!J15/K15/L15 values (9 of 22 columns) whenever those cells hold a value (RAN (H-S) at `47a1492`; RAN (LT §4.6, DC T3, REGISTRY-B `check1.py`) at the snapshot). **[GST RULE]**: this records a technical mismatch between the configured cells and the written cells; which cells are right is the author's call. *Fix:* none in any commit. *Similar cases:* any future output column whose label is contained in an earlier field name (READ, HIGH); the figure written also depends on the order of the extracted data (RAN (LT §4.6) with reversed order).

**D9 — fixes inside the PyQt5 rewrite (`c258047`, B only (after the split)).**

| Problem at the split (`47a1492`) | Scenario | Fix in `c258047` | Patch / proper | Similar cases still open at the snapshot |
|---|---|---|---|---|
| `main.py` crashed before any window when there is no console (`sys.stdout` is `None` under `pythonw.exe`, which Standard 13's launcher uses) | starting the program without a console | `if hasattr(sys.stdout, 'buffer'):` (`main.py:12-13`) | proper (RAN (H-B1): simulated `pythonw` start passes at `c258047`, fails at the split) | every start-up error is still only `print()`ed and written to `gst_organizer.log` in the current folder (`main.py:49-59`: `logging.error` at `:50` and `:57`), never shown in a window, so without a console the user sees nothing; `launch_extractor.py:21-23` only prints (it sets up no log); no `sys.excepthook` — RAN (H-B1): an error inside a window handler aborts the program with exit code 134 (re-run by the verifier with a minimal PyQt5 5.15.11 window: an uncaught error in a slot printed the traceback and aborted with 134); the ImportError advice "pip install PyQt5 openpyxl" (`main.py:51-53`) is wrong when the missing module is `pythoncom` (RAN (me) 12B.7) |
| Every start created `core/ utils/ gui/ logs/ temp/` folders and `__init__.py` files inside the program folder (`main.py:29-43@47a1492`) | any start, including from a shared or read-only folder | the creation removed (`main.py:26-29`) | proper | the log file is still created in the *current* folder (`main.py:15-23`), which under Standard 13's launcher (`start "" /D "%~dp0" …`) is the shared program folder (12F.7) |
| The scan ran on the window's own thread, freezing the window on a big or network folder (`gui/handlers/file_handler.py:57-95@47a1492`) | scanning a large source folder | `ScanWorker` in a QThread (`gui/controllers/scan_controller.py:14-94`) | proper | since `b0e7d0e` the Step 4 disk scan runs on the window thread at start-up, after processing, after extraction and on Refresh List (`gui/pyqt_main_window.py:190, 384, 561, 419-504`; MEDIUM that it pauses the window on a large network target) |
| Tk screen updates from worker threads; late-binding wrong names (D4) | processing / Dry Run | workers only emit signals | proper | none found (DC §4.1, §4.8) |
| Unreachable `return`, stray `print`, duplicated header block | — | removed | clean-up | other dead code remains (12E.5) |

The same commit introduced or exposed new gaps, each still open at the snapshot: the window needs pywin32 to start although `requirements.txt` only mentions it in a comment (`requirements.txt:13-14`; RAN (me) 12B.7); the Variations sheet is always empty (`gui/controllers/processing_controller.py:73`); the long-name confirmation, the close-while-processing question, the client-list export and the per-client folder choice were lost (12C.5); the old recent-folder list is not imported (`models/cache_manager.py:122`); the Extract tab was filled only from the session and found folders by a name search that missed clients whose files spell the state in full (RAN (H-B1)) — replaced in `7e09e5c` / `b0e7d0e`; scan progress was not connected — fixed in `7e09e5c`; a Max length of 5–8 could cut the state code off client folders (RAN (H-B1)) — floor restored in `7e09e5c`.

**D10 / D11 — the start-up failure and its repair (`a62d3cc` → `c21f788`, B only).**
- *What broke:* `a62d3cc` deleted 17 Tk files but left `gui/__init__.py:7` importing the deleted `gui/main_window.py`. Because `main.py` imports `gui.pyqt_main_window`, Python first runs `gui/__init__.py`, so the program stopped at start with `No module named 'gui.main_window'` — even with pywin32 present (RAN (me) with stand-in pywin32 modules on an extracted `a62d3cc` tree; RAN (H-B1) through `main.py`, which printed the misleading "pip install PyQt5 openpyxl").
- *Scenario:* starting the program at all, 27 minutes after the deletion (MEDIUM that the author or the AI tool tried to run it, from the time gap).
- *Fix (`c21f788`):* `from .pyqt_main_window import GSTOrganizerWindow` (`gui/__init__.py:7`). A **patch** of one line: the commit did not look for other importers of the deleted files.
- *Similar cases:* `gui/widgets/__init__.py:4-5` kept importing the deleted `collapsible_frame` / `title_bar` until `7e09e5c` (RAN (me): `import gui.widgets` fails at `c21f788`); `power_query_extractor/extractor_main.py:9` still imports the deleted `extractor_window` at the snapshot (RAN (me): `ModuleNotFoundError`), harmless only because nothing imports it (B-F31, B-R128). B has no tests (DO), so nothing would catch the next dangling import.

**D12 — unannounced fixes in `7e09e5c` (B only).** The message speaks of layout; the diff also repairs six things (READ diff; RAN (H-B2) unless stated):

| What broke | Scenario | Fix | Patch / proper | Still open at the snapshot |
|---|---|---|---|---|
| `gui/widgets/__init__.py` imported two deleted files | any import under `gui.widgets` — needed from this commit because the new views import `gui.widgets.colored_section` | re-exports `CollapsibleGroup`, `ColoredSection` (`gui/widgets/__init__.py:4-7`) | proper (required) | `CollapsibleGroup` re-exported but unused (DC D-04) |
| The Extract tab found no folder for a client whose files spell the state in full (the search key used the state as written, folders use the state code), and was empty after a restart | extraction in a later session, or for "Maharashtra"-style names (RAN (H-B1, H-B2) `t_build`) | Extract list from run history, keyed by the scan key `<client>-<state code>`, with the version folder stored | **patch**, replaced the next day (D13) | — |
| Scan progress never shown in PyQt5 (`scan_progress` connected to nothing since `c258047`) | any scan | connected (`gui/pyqt_main_window.py:129, 238-242`) | proper | the scan log is HTML without escaping and slows down sharply on long scans (RAN (VW): 5,000 lines 136.6 s; B-R137) |
| SCAN could be pressed again during a scan | impatient double click | SCAN FILES disabled until the scan ends (`gui/pyqt_main_window.py:231`) | patch | RE-SCAN is not disabled; pressing it during a scan clears the log and is then silently refused (`gui/controllers/scan_controller.py:57-61`; RAN (MW e17)) |
| Files under 1 KB or without an Excel signature vanished without trace (filtered inside `find_excel_files` since `edac6a3`) | damaged downloads or non-Excel files with an Excel extension in the source folder (MEDIUM that this was met) | the parser lists them as variations (`core/file_parser.py:181-215`) | proper (RAN (me) `scan_trace.py`) | a file whose size cannot be read is still dropped with only a log line (`core/file_parser.py:174-179`, B-R4); `.xlsm` files are collected but no pattern accepts them (`utils/helpers.py:255` vs `utils/constants.py:15, 23, 31, 39, 47, 55`, B-R1, B-R6–B-R11); the variations list reaches the user only as a count (B-F9) |
| Excel's lock files `~$…` (made while a workbook is open) were scanned and listed as unrecognised when ≥ 1 KB with an Excel signature (the made-up case of the trace); real lock files are normally only a few hundred bytes and not zip files (MEDIUM, general knowledge, not tested here), so before `7e09e5c` they were normally dropped silently by the size check — the name skip is what keeps them from appearing as "File too small to be valid Excel" now that rejected files are listed | a source workbook open in Excel during the scan (MEDIUM) | names starting `~$` dropped (`utils/helpers.py:259-261`) | proper (RAN (me) `scan_trace.py`; re-run by the verifier with the same five made-up files and the same result) | Step 4's refreshed-file check also skips `~$` names (`gui/pyqt_main_window.py:468-470`); the report search `ITC_Report_*.xlsx` does not need to (a lock file starts with `~$`) |

The same commit introduced new problems (RAN (H-B2)): three Setup "Browse" buttons drawn white on white (`gui/widgets/colored_section.py:62`, `gui/views/setup_view.py:114`); dark mode no longer reaching Setup and Validation; Step 4's "latest" version by day-first name order; phantom rows after an extraction — the last two fixed by D13, the first two still open (B-F3, B-F28).

**D13 — the disk-driven Extract tab (`b0e7d0e`, B only).**
- *What broke (in D12's patch):* the run-history list picked the "latest" version by reverse *name* order of `Version-DDMMYY HHMM` (RAN (H-B2): `Version-150126 0800` chosen over `Version-010226 0900`); each extraction was stored under a display-name key that matched no processing record, adding a phantom row (RAN (H-B2) `t_hist7e`); clients processed on another PC never appeared.
- *Fix:* the list is read from the target folder: one tree row per client folder and one child per version folder that holds a report (`gui/pyqt_main_window.py:419-504`); the user ticks versions explicitly. **Proper** for the three symptoms (none can occur any more).
- *Similar cases still open:*
  - the **run folder** is now chosen by the same day-first name order (`gui/pyqt_main_window.py:432-448`; RAN (REGISTRY-B): `Annual Statement-311225 1000` (31-12-2025) chosen over `Annual Statement-211026 0335` (21-10-2026); re-run by the verifier through the real `_populate_extraction_from_disk` on a made-up target with these two run folders: the same choice, versions listed `Version-150126 0800` above `Version-010226 0900`, and a 1-byte `…_Refreshed_…` file shown as "Refreshed"); Re-Run/Resume choose by modification time (`core/file_organizer.py:158-169`), so Step 4 can show a different batch from the one just processed (B-R38 vs B-R39);
  - versions are listed in day-first name order too (`gui/pyqt_main_window.py:457`, cosmetic, DC S-10);
  - "Refreshed" means only that some `*_Refreshed*` file exists — also after a failed refresh, because the copy is made before refreshing and never removed (`report_processor.py:118-120`; RAN (H-B2) `t_worker`, RAN (MW e9));
  - the tree uses the last *saved* target folder, not the text in the Setup box (`gui/pyqt_main_window.py:421`; READ (H-B2));
  - extraction results are still written to run history under a key nothing reads (`gui/pyqt_main_window.py:554-558`, B-R120).

#### 12D.3 Edge cases visible in the first commit (met before git)

The first commit already contains defensive code that only makes sense after problems met in June–August 2025 (READ; MEDIUM that each was met in real use; the code is unchanged in B):

| Defensive code (at the snapshot) | Problem it answers |
|---|---|
| Up to 5 tries, 2 s apart, to set Excel's visibility and alerts (`power_query_extractor/core/report_processor.py:176-190`) | Excel not ready right after start |
| 3 tries when Excel answers "Call was rejected by callee" (`report_processor.py:194-205`) | Excel busy |
| Window activation retried 3 times (`report_processor.py:213-232`) | the Excel window not coming to the front before the keystrokes |
| Comments "Increased from 3 to 5", "Increased from 0.5" (`report_processor.py:209, 239, 244`); the `edac6a3` comment "WAIT 7 SECONDS AS REQUESTED" above a 10-second loop (replaced in `bd176bf`) | waits tuned by trial and error |
| Error-dialog search and clicking (`report_processor.py:260-298, 368-476`) | Power Query error dialogs blocking the refresh |
| Paths over 218 characters: 8.3 short path (`core/excel_handler.py:292-293`) or `\\?\` prefix (`report_processor.py:668-680`); advice "OneDrive/SharePoint sync issue" (`report_processor.py:741`) | long client folder paths; synced folders |
| Optional "(n) " prefix accepted only for GSTR-3B names (`utils/constants.py:31`) ⚠ | browser duplicate downloads of monthly GSTR-3B files |
| Excel COM instead of openpyxl for reports (`core/excel_handler.py:2-3, 83-93`); README "Excel corruption issues from previous versions" | earlier versions corrupting templates / losing Power Query (MEDIUM-HIGH, H-S) |

#### 12D.4 Fixes that needed a second fix

| First change | What it left or caused | Second change | Still open |
|---|---|---|---|
| `6237597` thread deferral (shared) | wrong names in the Tk Dry Run | none in Tk; the code was deleted in `a62d3cc` and PyQt5 has no such pattern | — |
| `bd176bf` configurable suffix (shared) | searches still fixed to `_Refreshed_` | none; `b0e7d0e` added a second fixed search (`gui/pyqt_main_window.py:468`) | yes (B-R63, B-R136) |
| `c258047` Extract tab (B only) | lost on restart; missed clients with full state names | `7e09e5c` run-history list → `b0e7d0e` disk tree | run folder by name order |
| `c258047` Max length floor 5 (B only) | state code cut off, folders merged (RAN (H-B1)) | `7e09e5c` floor 15 | the box is greyed out but still applied (B-R32) |
| `c258047` consolidated report into the target folder (B only) | outputs of one batch spread over two folders | `b0e7d0e` back into the Level 1 folder | — |
| `a62d3cc` Tk deletion (B only) | three dangling imports | `c21f788` (one), `7e09e5c` (two) | `extractor_main.py:9` |
| `7e09e5c` run-history Extract list (B only) | wrong latest version, phantom rows | `b0e7d0e` | the run-folder choice |
| `c258047` scan progress (B only) | not connected | `7e09e5c` | — |

#### 12D.5 Similar cases still unhandled at the snapshot — one list

1. Step 4 chooses the run folder by day-first name order (`gui/pyqt_main_window.py:432-448`) while Re-Run/Resume use modification time (`core/file_organizer.py:158-169`).
2. Custom refreshed-file suffix invisible to Skip refresh (`power_query_extractor/core/report_processor.py:102`) and to the Step 4 status (`gui/pyqt_main_window.py:468`).
3. "Refreshed" shown after a failed refresh (`report_processor.py:118-120` + `gui/pyqt_main_window.py:466-479`).
4. Consolidated look-up fills 9 columns from other cells (`power_query_extractor/core/data_consolidator.py:130-138`) **[GST RULE]** ⚠ — author's call.
5. Fixed waits instead of a completion check (`report_processor.py:207-298`).
6. The window cannot start without pywin32, which `requirements.txt:13-14` leaves commented out; the start-up advice names the wrong package (`main.py:51-53`).
7. Start-up errors only printed and logged to `gst_organizer.log`, never shown in a window; no `sys.excepthook` (`main.py:49-59`, `launch_extractor.py:21-23`).
8. `power_query_extractor/extractor_main.py:9` imports a deleted file.
9. A file whose size cannot be read vanishes from the scan (`core/file_parser.py:174-179`); `.xlsm` source files are collected but never recognised (`utils/helpers.py:255`).
10. Variations never shown as a list and the summary's Variations sheet always empty (`gui/controllers/processing_controller.py:73`).
11. RE-SCAN not disabled during a scan (`gui/pyqt_main_window.py:231`; `gui/controllers/scan_controller.py:57-61`).
12. Step 4 disk scan on the window thread (`gui/pyqt_main_window.py:419-504`).
13. Step 4 uses the last saved target, not the text in the box (`gui/pyqt_main_window.py:421`).
14. Three Setup Browse buttons invisible; dark mode partial (`gui/widgets/colored_section.py:62`, `gui/views/setup_view.py:114`).
15. Client data at rest: program log in the current folder, extraction log in the run folder, figures in `~/.gst_organizer/run_history.json` (B-R119, B-R121, B-R122); the old run log in history at `edac6a3`.

#### 12D.6 Where the older document (README.md) disagrees with the bug history

| README says (line) | Bug history shows |
|---|---|
| "📝 Recent Bug Fixes (Latest Commit)": circular import, `scanned_files` access, thread safety, indentation, progress callback, "Enhanced Excel corruption prevention mechanisms" (`README.md:320-327`) | Only the progress-callback change had an effect; "thread safety" introduced a regression; "indentation" and "circular import" changed nothing; no Excel-corruption change is in `6237597`'s diff (D4). The README itself was last changed in `d674386`, 12 commits before B's snapshot, and the "latest commit" whose fixes it lists is `6237597`, earlier still (RAN `git rev-list --count`: 12 commits after `d674386`; 15 commits of `main` are not in `6237597`'s own history, counting `cbfa4a5`). |
| "Check the log file: `gst_organizer.log`" (`README.md:333`) | True for the main program; nothing in the README warns that the log holds client data (D1). |
| "Known Limitations … Power Query refresh requires Windows and Excel" (`README.md:312-318`) | Since `c258047` the *whole window* requires Windows with pywin32, not just the refresh (D9, 12F). |

### 12E. Abandoned Approaches and Reversals

This sub-section lists everything in B's history that was tried and then removed, replaced or left half-built: code added and later deleted, swapped libraries and approaches, features started but never finished, and the dead methods, dead constants, commented-out code and paste markers that remain at the snapshot `b0e7d0e` as remnants. For each: what was tried, why it most likely failed (with confidence), and what replaced it. These are technical lessons only; nothing here judges a GST rule.

#### 12E.1 Approaches abandoned before git (known only from the first commit)

| What was tried | Evidence | What replaced it | Likely reason (confidence) |
|---|---|---|---|
| One large window file `gui/app.py` (89,275 bytes) | `gui/__pycache__/app.cpython-312.pyc` / `-313.pyc` in `edac6a3` (RAN (H-S), header fields); `main.py:72@edac6a3` still listed `gui/app.py`; the `edac6a3` README tree showed it | the modular Tk window (`gui/main_window.py:4@edac6a3` "Modularized version") on or before 23-07-2025 | size and maintainability (MEDIUM; the reasons are lost with the pre-git chats) |
| Report creation with openpyxl only | header `core/excel_handler.py:2-3` "WINDOWS VERSION - Uses Excel COM to preserve Power Query"; README "Excel corruption issues from previous versions have been specifically addressed" (`README.md:254@edac6a3`) | Excel through COM, openpyxl only as a fallback with a warning (`core/excel_handler.py:19-28, 83-93`) | openpyxl drops Power Query when it saves (MEDIUM-HIGH; the same fact is in Standard 12) |
| Old extraction mapping format (`CELL_MAPPINGS` per sheet/cell, display names like "Diff IGST (ITC)") | `power_query_extractor/config/cell_mappings copy.py@edac6a3` (63 lines; dated 21-06-2025 by the `.pyc` size match, RAN (H-S)) | `EXTRACTION_CONFIG` (output sheet → list of mappings) with a generated `CELL_MAPPINGS` "for backward compatibility" (`power_query_extractor/config/cell_mappings.py:255-278`); copy deleted in `47a1492` (shared - before the split on 30-10-2025) | one sheet per output table was wanted (MEDIUM) |
| Old consolidated layout (one sheet with ITC and Sales sections and a details sheet) | dead `_write_section` and `_create_details_sheet` (`power_query_extractor/core/data_consolidator.py:237-339, 341-399`); the consolidator's source was 10,208 bytes on 25-06-2025 and 17,711 bytes on 14-07-2025 (RAN (H-S)) | one sheet per configured output sheet plus a Summary (`create_report`, `:20-174`, and `_create_summary_sheet`, `:208-235`) | same as above (MEDIUM) |
| A 7-second refresh wait | comment "WAIT 7 SECONDS AS REQUESTED" above a 10-second loop (`power_query_extractor/core/report_processor.py@edac6a3`, removed in `bd176bf`) | fixed 10 s, then the user's wait (`bd176bf`) | refreshes took longer (MEDIUM: the comment itself shows the 7-second request was overruled by a 10-second loop, and the nearby comments "Increased from 3 to 5", "Increased from 0.5", `report_processor.py:209, 239, 244`, show other waits being lengthened by trial; that the refresh itself was too slow is an inference — [Corrected during verification — originally HIGH]) |
| `command=` button wiring in Tk | comments "working mouse events" / "proper event binding" (`gui/tabs/setup_tab.py:20, 103@edac6a3`) | `bind('<Button-1>')`, which fires even on a disabled button (H-S O5) | unknown (LOW); moot in B since `a62d3cc` |

#### 12E.2 Code added and later removed or replaced (in date order)

| # | What was tried | Added in (label) | Removed / replaced in (label) | Why it likely failed | What replaced it | Lesson |
|---|---|---|---|---|---|---|
| 1 | Committing the run log and compiled files | `edac6a3` (shared - before the split on 30-10-2025) | `cbfa4a5` / `9c56a1c` (log), `1cb878b` + `01862a0` (`.pyc`) (shared) | no `.gitignore`; the log held client data (HIGH) | `.gitignore` (`.gitignore:1-82`) | deleting in a later commit does not remove data from history (12D D1) |
| 2 | Direct Tk calls from the worker thread | `edac6a3` | `6237597` wrapped part of them in `root.after(0, lambda …)` (shared); all deleted in `a62d3cc` (B only (after the split)) | Tk may only be changed from its own thread; the wrappers read loop variables late (RAN (H-S): wrong names in Dry Run) | Qt signals from QThread workers (`gui/controllers/*.py`) | hand screen updates to the window as finished messages, never as lambdas over loop variables |
| 3 | Long-name warning at a fixed 10 characters | `edac6a3` | `6237597` (threshold = max-length setting, shared); the warning itself lost in `a62d3cc` (B only) | 10 was below the default folder-name limit of 35 (MEDIUM) | nothing in code; a stale note (`gui/views/setup_view.py:385`) | removing a confirmation step should be a decision, not a side effect of a rewrite |
| 4 | Fixed 10 s + 10 s refresh waits | `edac6a3` | `bd176bf` (shared) | some refreshes need longer (HIGH) | user wait *w* and max(10, *w*) (`report_processor.py:253-258, 280-285`) | still a guess (12D D7) |
| 5 | Fixed refreshed-copy name `_Refreshed_<DDMMYY_HHMM>` | `edac6a3` | `bd176bf` (shared) | the author wanted another name (HIGH) | user suffix (`report_processor.py:91-94`) — but the two searches stayed fixed (`:102`; `gui/pyqt_main_window.py:468`) | when a naming rule becomes configurable, every reader of that name must use the same rule |
| 6 | One "select client" tick box per client in the Tk extractor | `edac6a3` | `bd176bf` (shared) | some clients need only one report (HIGH) | two per-report boxes (now `gui/views/extraction_view.py:291-356`) | — |
| 7 | `.claude/settings.local.json` tracked | `bd176bf` (shared) | `c258047` (B only) | personal tool file (HIGH) | `.claude` in `.gitignore:81-82` | — |
| 8 | `[DEBUG]` diagnostics written at INFO level (about 50 lines, including a dump of every client folder and file name) | `edac6a3` | `c258047` (B only), turned into 3 `logger.debug` lines or deleted | log noise; client detail in the log (HIGH) | quieter log (`core/excel_handler.py:72, 100, 546`) | — |
| 9 | Tk window layer (13 files in `gui/handlers`, `gui/tabs`, `gui/utils`, 2 Tk widgets, `gui/main_window.py`) and the Tk extractor window | `edac6a3` | `a62d3cc` (B only) | a planned rewrite to PyQt5 (HIGH, `c258047` message "BREAKING CHANGE") | `gui/pyqt_main_window.py`, `gui/views/*`, `gui/controllers/*`, `gui/theme/*`, `models/*` | the rewrite re-implemented the screens from scratch and silently dropped several v3 behaviours (12C.5) |
| 10 | The refactored Tk `processing_handler.py` (7 methods) | `c258047` (B only) | `a62d3cc` (B only), two minutes later | it was already superseded by the PyQt5 worker (HIGH) | `gui/controllers/processing_controller.py` | work on a file about to be deleted is wasted (it never ran in a released state) |
| 11 | Extract tab filled only from the processing run just finished, folders found by a name search over the whole target tree (`rglob('*')`, key `<client>-<state as written>`) | `c258047` (B only) | `7e09e5c` (B only) | lost on restart; the key used the state as written while folders use the state code; whole-tree walk slow (RAN (H-B1, H-B2)) | run-history list (row 12) | look up folders by the same rule that names them |
| 12 | Extract tab filled from run history; "latest" version by reverse name order; two-level name search | `7e09e5c` (B only) | `b0e7d0e` (B only), next day | day-first names are not date order; extraction results saved under a different key made phantom rows; per-PC history misses work done elsewhere (RAN (H-B2)) | disk-driven tree (`gui/pyqt_main_window.py:419-504`) | — |
| 13 | Version folder passed through the processing summary (`client_key`, `version_folder` in each entry) | `7e09e5c` (B only) | never read at any commit: the only reader of the summary's client list, `_populate_extraction_from_results` (`c258047`), was deleted in the same commit that added the two keys (RAN (verifier) pickaxe and grep of the `7e09e5c` tree) | superseded by the run-history approach (MEDIUM that the keys were written for that reader) | nothing; dead data at `gui/controllers/processing_controller.py:139-140` | — |
| 14 | "Last Refresh" column read from run history | `c258047` (B only) | `b0e7d0e` (B only) | always blank: the extraction key never matched the processing key (RAN (H-B1)) | "Refresh Status" from `*_Refreshed*` files (`gui/pyqt_main_window.py:466-479`) | — |
| 15 | Consolidated report into the target folder | `c258047` (B only) | `b0e7d0e` (B only) | one batch's outputs ended up in two folders (MEDIUM) | back into the Level 1 folder (`gui/controllers/extraction_controller.py:92`) — the original Tk extractor's choice (`power_query_extractor/gui/extractor_window.py:808@47a1492`) | a reversal: `c258047` changed it without need |
| 16 | Max length floor 5 | `c258047` (B only) | `7e09e5c` (B only) | 5–8 cut the state code off client folders and could merge two clients (RAN (H-B1)) | floor 15 (`gui/views/setup_view.py:375`), the v3 value | a reversal |
| 17 | New mode wording "Re-Run - Overwrite existing files", "Resume - Skip already processed" | `c258047` (B only) | `7e09e5c` (B only) for the buttons | did not describe the code (READ, HIGH) | the `PROCESSING_MODES` wording, typed by hand (`gui/views/setup_view.py:343-359`); the help card kept a short form of the `c258047` wording (`:510-512`) | the reversal was only half done: buttons and card now disagree |
| 18 | Tick-box label "Include client name in folder names" | `c258047` (B only) | `7e09e5c` (B only) | reverted to the v3 wording (HIGH, docstrings "matching the original Tkinter UI") | "Include client name in brackets for Level 4 folders" (`gui/views/setup_view.py:369`) | a reversal |
| 19 | File validation inside `find_excel_files` (silent drop) | `edac6a3` | `7e09e5c` (B only) | slow on big or network folders (comments "fast glob, no I/O per file", `core/file_parser.py:146`) and invisible to the user (MEDIUM-HIGH) | single `stat` and 8-byte read in the parser, failures listed (`core/file_parser.py:171-226`) | the old helpers stayed as dead code (12E.5) |
| 20 | Emoji in the extraction worker's own messages | `c258047` (B only) | `b0e7d0e` (B only) | a plain-text log file was added the same day (LOW) | "OK/FAIL" etc. (`gui/controllers/extraction_controller.py:76-84`); the refresh module's messages still contain emoji (`report_processor.py:127, 198`) | — |
| 21 | Standalone extractor's disk discovery | `edac6a3` (Tk extractor) | deleted in `a62d3cc` (B only); **rebuilt differently** in `b0e7d0e` (B only) | the rewrite first replaced it with session results (rows 11–12) | `gui/pyqt_main_window.py:419-504`: run folder by greatest name (not modification time), all versions (not one), no `_` skip, one status per version | a reversal that did not bring back the original's run-folder rule (12D D13) |

#### 12E.3 Swapped libraries and approaches

| From | To | When (label) | Why (confidence) | Consequences still visible |
|---|---|---|---|---|
| Tkinter (Python's built-in window toolkit) | PyQt5 (Qt 5) | `c258047` / `a62d3cc` (B only (after the split)) | modernisation in one AI-assisted session ("Co-Authored-By: Claude Opus 4.6"), reasons not stated (MEDIUM: appearance, threading) | new dependency `PyQt5>=5.15.0` (`requirements.txt:5`); the author's Standard 12 names PySide6 (Qt 6) for desktop windows — a fact for the author, not a recommendation (12F.7) |
| Two programs (organiser + Power Query Extractor) | one four-tab window | `c258047` (B only) | one workflow, one settings store (HIGH, message "Unified two separate apps") | the extractor's pywin32 import became a start-up requirement of the whole program (12F.5) |
| Recursive widget recolouring for dark mode (`gui/utils/dark_mode_manager.py`, 212 lines) | application style sheets (QSS) | `c258047` (B only) | simpler (HIGH, message "replaces 212-line recursive widget walker") | from `7e09e5c` inline light styles override the dark sheet on Steps 1–2 (RAN (H-B2)) |
| One settings file `~/.gst_organizer_cache.json` | `~/.gst_organizer/settings.json` + `run_history.json` + `session_state.json` | `c258047` (B only) | run history and session state wanted (HIGH) | one-time import loses the recent-folder list (B-R118); `session_state.json` written but never read (B-R117) |
| Synchronous scan on the window thread | QThread worker | `c258047` (B only) | the window froze on big folders (HIGH) | the Step 4 disk scan (`b0e7d0e`) runs on the window thread again |
| A single `self.app` object used by every Tk handler | models + views + controllers + signals | `c258047` (B only) | clarity and separation of the state from the window (MEDIUM: the message says only what replaced it — "Framework-agnostic state layer", "Signal/slot architecture replaces 305 self.app.xxx callback references" — not why; [Corrected during verification — originally "testability and clarity (HIGH, message)"]) | the planned central signal hub `models/signals.py` was never used; each controller has its own signals |
| Template picker `*.xltx *.xlsx` | `*.xlsx *.xlsm` | `c258047` (B only) | not stated (LOW) | reports are always saved as `.xlsx` whatever the template (B-R51); whether Excel accepts an `.xlsm` saved as `.xlsx` was not testable here |
| Settings saved on every change | saved at a few moments | `c258047` (B only) | not stated | a crash before a save loses the changes (B-R116) |

#### 12E.4 Features started but never finished (partial implementations still in the code)

| Feature | What exists | What is missing | Since (label) | Evidence |
|---|---|---|---|---|
| Central signal hub | `AppSignals` with 18 signals (`models/signals.py:9-38`; [Corrected during verification — originally 20]) | never imported anywhere | `c258047` (B only (after the split)) | READ; DC D-01; RAN (me) `verify_imports.txt` (no importer) |
| Collapsible group widget with animation | `CollapsibleGroup` (`gui/widgets/collapsible_group.py:9-42`), animation classes imported but unused (`:6`) | never used by any view; `CollapsibleCard` written instead (`gui/views/setup_view.py:17-82`) | `c258047` (B only) | DC D-04, D-06 |
| Per-client include-name choice | look-up `core/file_organizer.py:117-126`, dictionary `models/app_state.py:83`, passed through `gui/pyqt_main_window.py:370` | any screen control to fill it (the Tk toggle was deleted in `a62d3cc`) | `edac6a3` / `a62d3cc` | B-F5, B-R36 |
| Export the client list | "EXPORT LIST" button (`gui/views/validation_view.py:186-200, 221`) | connected to nothing (RAN (MW e2): 0 receivers) | `c258047` (B only) | B-F10 |
| Long-client-name warning | the note "Names >10 chars will show warning. This overrides individual client settings if checked." (`gui/views/setup_view.py:385`) | no code warns about names, and there are no individual settings | note re-added `7e09e5c` (B only) | DC D-33 |
| Session restore | `active_tab`, `window_size`, `recent_source_folders` written (`gui/pyqt_main_window.py:604-607`, `models/cache_manager.py:45-52`) | `load_session_state` (`models/cache_manager.py:82-84`) and `get_recent_source_folders` (`:86-89`) never called | `c258047` (B only) | DC D-11, D-55 |
| Extraction history | `add_extraction_result` stores every extracted figure (`models/run_history.py:73-95`) | nothing reads `pq_refreshed`, `pq_refresh_timestamp`, `pq_extraction_data`; `itc_report_path`, `sales_report_path`, `errors` never filled (`models/run_history.py:27-28, 33`) | `c258047` (B only) | DC D-56; B-R119 |
| `ScanResult` helpers and `ProcessingState` flags | `total_files`, `total_clients`, `complete_clients`, `clear()` (`models/app_state.py:57-74`); `is_processing`, `stop_requested`, `selected_clients`, `reset()` (`:76-88`) | at the snapshot none is used. The three properties were read by the scan summary in `c258047` and `7e09e5c` (`gui/pyqt_main_window.py:231-233@c258047`, `:245-247@7e09e5c`) and lost their last reader in `b0e7d0e`, which switched the summary to `statistics`; `clear()` and the four `ProcessingState` items were never used at any commit (only `client_folder_settings` is passed on, `gui/pyqt_main_window.py:370`); `clear()` was not even updated for `statistics` (`b0e7d0e`). RAN (verifier) grep of the `c258047`, `7e09e5c` and `b0e7d0e` trees — [Corrected during verification — originally "never used" for all of them since `c258047`] | `c258047` (B only); the properties dead since `b0e7d0e` | DC D-12, D-13 |
| Tab "pending" state | grey `#a0a0a0` dot (`gui/pyqt_main_window.py:582-590`) | never set; dots never reset | `b0e7d0e` (B only) | B-R140 |
| Theme extras | `COLORS` (`gui/theme/theme_manager.py:10-22`), `#info` button rules (`:136-144, 319-327`), `toggle()` (`:421-424`) | never used | `c258047` / `7e09e5c` (B only) | DC D-15, D-44; B-R132 |
| Calculated output columns and sheet descriptions | `_calculate_value` (`power_query_extractor/core/data_consolidator.py:176-206`), description branch (`:51-65`), documented options (`power_query_extractor/config/cell_mappings.py:15-26`) | no mapping uses them (all 22 have `calculation: None`) | `edac6a3` (shared - before the split on 30-10-2025) | B-R84, B-R91 |
| Extraction options | `auto_refresh_if_missing`, `show_timestamp` (`power_query_extractor/config/cell_mappings.py:250-251`) | never read | `edac6a3` | B-R64, B-R92 |
| Excel safety settings | `EXCEL_SAFETY` retries, prefixes, 10 KB minimum, save options (`utils/constants.py:217-228`), copied to `self.safety_config` (`core/excel_handler.py:54`) | never read | `edac6a3` | B-R106–B-R109 |
| "Don't quit if Excel was already open" | comment and condition (`core/excel_handler.py:206-211`), `self.excel_app` (`:55`) | `self.excel_app` is never assigned, so the user's own Excel is always quit | `edac6a3` | B-R96; DC D-29 |
| Result validation and clean-up after a failed client | `validate_results` (`core/file_parser.py:366-399`, "critical" types ⚠ [GST RULE]), `cleanup_on_error` (`core/file_organizer.py:364-390`), `get_summary` (`:392-409`) | never called | `edac6a3` | B-R28, B-R41 |
| Unique names instead of overwriting; file hashes; relative Excel paths; path shortening for display | `get_unique_path`, `calculate_file_hash`, `get_relative_excel_path`, `truncate_path` (`utils/helpers.py:344, 422, 450, 287`) | never called | `edac6a3` | B-R46, B-R47, B-R58, B-R143 |
| Old entry point | `power_query_extractor/extractor_main.py` | imports a deleted file; nothing calls it | broken since `a62d3cc` (B only) | B-F31, B-R128 |

#### 12E.5 Remnants: dead methods, dead constants, commented-out code and paste markers

READ at the snapshot (checked against DC §2, §5); "since" says when each became a remnant.

| Remnant | Where | Since (label) | Origin |
|---|---|---|---|
| Paste marker from an AI chat "# Rest of your code remains same..." | `power_query_extractor/core/report_processor.py:251` | `edac6a3` (shared - before the split on 30-10-2025) | c258047's clean-up removed the other markers (`# ADD THIS`, `# ADD THE NEW CODE HERE …`, `# ... rest of the code`) but not this one |
| Change markers "# UPDATED:", "# NEW:", "# NEW" ×3, "# IMPORTANT:" | `utils/helpers.py:20, 470`; `core/excel_handler.py:386-388`; `core/file_organizer.py:216` | `edac6a3` | AI-assisted editing before git |
| Tuning comments "Increased from 3 to 5", "Increased from 0.5" ×2 | `power_query_extractor/core/report_processor.py:209, 239, 244` | `edac6a3` | pre-git trial and error (12D.3) |
| Commented example configuration block "ADD MORE SHEETS HERE AS NEEDED" | `power_query_extractor/config/cell_mappings.py:229-241` | `edac6a3` | intentional guidance, not dead code |
| Commented dependency `# pywin32>=306` | `requirements.txt:13-14` | `c258047` (B only (after the split)) | 12F |
| "Import here to avoid circular import" for a circular import that does not exist | `utils/helpers.py:490` | `edac6a3`; import made relative in `6237597` (shared) | 12D D4 |
| "Default to True for backward compatibility" — every caller now passes the values | `power_query_extractor/core/report_processor.py:51, 61` | `bd176bf` (shared); moot since `c258047`, whose callers always pass them (`gui/pyqt_main_window.py:526-527`) | |
| Dead helpers `get_date_only`, `get_file_info`, `validate_excel_file` (+ unused import in `core/excel_handler.py:38`) | `utils/helpers.py:27, 115, 210` | `get_file_info`, `validate_excel_file`: `7e09e5c` (B only); `get_date_only`: `edac6a3` | validation moved into the parser |
| Dead constants `GUI_CONFIG`, `MESSAGES`, `PROCESSING_MODES` names/descriptions/icons | `utils/constants.py:126-151, 191-211, 157-173` | `a62d3cc` (B only) — the Tk files were their last readers (`gui/main_window.py:58-63@47a1492`, `gui/handlers/file_handler.py:123-139@47a1492`, `gui/tabs/setup_tab.py:216, 221@47a1492`) | RAN (me) grep on extracted trees |
| Dead constants `EXCEL_EXTENSIONS`, `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB`, `EXCEL_SIGNATURES`, `FOLDER_STRUCTURE['level2']`, pattern `category`/`folder`/`description` | `utils/constants.py:179-185, 83, 17-59` | `edac6a3` | never used at any commit (H-S §1.14) |
| Dead methods `_write_section`, `_create_details_sheet` (and the imports `FIELD_DISPLAY_NAMES`, `CELL_MAPPINGS` they alone use) | `power_query_extractor/core/data_consolidator.py:237-339, 341-399, 11` | before git | 12E.1 |
| Dead `ReportProcessor.cleanup()` and `excel_instance` | `power_query_extractor/core/report_processor.py:749-751, 25` | `edac6a3` | |
| Unreachable "Open method … failed" branch | `power_query_extractor/core/report_processor.py:122-137` | `edac6a3` | B-R72 |
| Dead `ProcessingView.reset()` | `gui/views/processing_view.py:159-163` | `c258047` (B only) | |
| Unused imports left by the redesign: `QGroupBox` (`gui/views/setup_view.py:10`), `QFrame` (`gui/views/processing_view.py:7`, `gui/views/extraction_view.py:10`), `QMenuBar` and `RunHistory` (`gui/pyqt_main_window.py:12, 18`) | as listed | `c258047` / `7e09e5c` (B only) | DC §2.9 |
| Unused variable `state_code = get_state_code(...)` that still logs "Unknown state" warnings | `core/excel_handler.py:550` | `c258047` (B only), when the debug line that used it was removed | DC §2.10 |
| `.gitignore` entries for files B no longer makes: `pq_extractor.log` (`:41`), `config/last_session.json` (`:63`), `.gst_organizer_cache.json` (`:42`), `temp_gst_*`, `backup_*` (`:52-53`) | `.gitignore` | `1cb878b` (shared); the first three became moot in `a62d3cc`/`c258047` | DC L-14 |
| Version labels "3.0.0" (`core/__init__.py:12`, `utils/__init__.py:38`), "1.0.0" (`power_query_extractor/__init__.py:3`), "4.0.0" (`gui/__init__.py:11`), GUI_CONFIG title "GST File Organizer v3.0" (`utils/constants.py:128`) | as listed | `edac6a3` / `c21f788` | never read; the window title says v4.0 (`gui/pyqt_main_window.py:71`) |
| Names from the two-program era: `launch_extractor.py`, `power_query_extractor/` as a package name, the empty `power_query_extractor/gui/` | as listed | `c258047` / `a62d3cc` (B only) | the code they name now serves the one window |

No `TODO`, `FIXME`, `XXX` or `HACK` markers exist anywhere (RAN (DC) grep; confirmed READ by my own grep of the snapshot for the marker words above).

#### 12E.6 Lessons that should not be re-learned in a rebuild (technical only)

1. **Delete with a search.** `a62d3cc` deleted 17 files without checking their importers; three dangling imports followed, one of which stopped the program (12D D10). Before deleting a module, search for every import of it — or keep a small test that imports every module.
2. **A rewrite needs a feature list.** The PyQt5 rewrite was written from scratch, not ported; at least twelve v3 behaviours disappeared silently (12C.5). The registry's feature list (B-F1 … B-F34) is that list for the next rebuild.
3. **One naming rule, one finder.** Three times a name was *made* by one rule and *searched for* by another: the refreshed-copy suffix (`bd176bf`), the Extract tab's folder key (`c258047`), the run-history keys (`c258047`, `7e09e5c`, `b0e7d0e`). Each produced silent misses.
4. **Day-first dates are not sortable text.** `DDMMYY HHMM` names were sorted as text in the shared Tk extractor (versions), in `7e09e5c` (versions) and in `b0e7d0e` (run folders); each time the "latest" was wrong across month or year ends (RAN (H-S, H-B2, REGISTRY-B)).
5. **Write paths and read paths must be checked together.** The v3 settings import reads a key the v3 writer never wrote (`recent_source_folders` vs `recent_folders`); session state is written and never read.
6. **Do not rely on fixed waits for Power Query.** Every timing change in history (`edac6a3` comments, `bd176bf`) made the wait longer or configurable; none detected completion. The author's Standard 12 describes a completion-based method (12F.7) — whether it fits these templates is for the author to test.
7. **Platform-only imports decide who can start the program.** The unguarded pywin32 import in `report_processor.py:11-15` was harmless while the extractor was a separate program and became a start-up requirement when the programs were merged (12F.5).

#### 12E.7 Where the older document (README.md) disagrees with this history

| README says (line) | History shows |
|---|---|
| Two programs: `python main.py` and, after processing, `python launch_extractor.py` (`README.md:112, 166-176`) | Merged into one window in `c258047`; `launch_extractor.py` only opens tab 4. |
| Folder tree without `models/`, `gui/views/`, `gui/controllers/`, `gui/theme/` (`README.md:71-95`) | Those folders replaced the Tk ones in `c258047` / `a62d3cc`. |
| "Safe file operations: All Excel operations use temporary files" (`README.md:42`) | The temp-file settings (`EXCEL_SAFETY`) were never wired in at any commit (12E.4). |
| "The extractor auto-loads target folder from cache" (`README.md:171`) | The Tk extractor's cache look-up never found the organiser's file (RAN (H-S)); the approach was abandoned with the Tk extractor in `a62d3cc`; it survives only in the broken `extractor_main.py:36-60`. |

### 12F. Dependency Evolution

This sub-section follows every outside package and tool B needs — what `requirements.txt` declared at each commit, what the code actually imported, and when each changed and why — and lists what is used but not declared. *Dependency* here means a package or program the code needs that is not part of the project itself; *declared* means written in `requirements.txt` (the file `pip install -r requirements.txt` reads); *stdlib* (standard library) means the modules that come with Python itself; *version specifier* `>=3.1.2` means "this version or any newer one", as opposed to an exact pin `==3.1.2`.

Evidence: RAN (me) `scripts/imports_by_commit.py` → `imports_by_commit.txt` (every outside import at every commit, with `file:line`), `scripts/import_changes_per_file.py` → `import_changes_per_file.txt` (per-file changes), the import-chain runs of 12B.7; READ of every version of `requirements.txt` (`git log -- requirements.txt`: only `edac6a3` and `c258047` touch it).

#### 12F.1 `requirements.txt`, version by version

**Version 1 — `edac6a3` (31-08-2025, shared - before the split on 30-10-2025), unchanged to the split point `47a1492`:**

```text
# GST File Organizer v3.0 Requirements
# Python 3.7 or higher required

# Core dependency for Excel operations
openpyxl>=3.1.2

# Optional but recommended for better performance
lxml>=4.9.0

# Development dependencies (optional)
# pytest>=7.0.0
# black>=23.0.0
# pylint>=2.17.0
```

**Version 2 — `c258047` (07-02-2026, B only (after the split)), unchanged to the snapshot `b0e7d0e` (`requirements.txt:1-14`):**

```text
# GST File Organizer v4.0 Requirements
# Python 3.7 or higher required

# GUI framework
PyQt5>=5.15.0

# Core dependency for Excel operations
openpyxl>=3.1.2

# Optional but recommended for better performance
lxml>=4.9.0

# Windows-only: Excel COM automation for Power Query
# pywin32>=306
```

| Line | Version 1 (`edac6a3`) | Version 2 (`c258047`) | Why it changed (confidence) |
|---|---|---|---|
| Header | "v3.0" | "v4.0" | the PyQt5 rewrite was called v4.0 (HIGH) |
| `PyQt5>=5.15.0` | absent | **added** | the new window toolkit (HIGH, commit message "requirements.txt: added PyQt5>=5.15.0") |
| `openpyxl>=3.1.2` | present | unchanged | — |
| `lxml>=4.9.0` | present ("Optional but recommended") | unchanged | — |
| pywin32 | **absent**, although used | added **only as a comment** `# pywin32>=306` | probably meant as "install on Windows only" (MEDIUM); a comment is ignored by `pip`, so a plain `pip install -r requirements.txt` never installs it |
| `pytest`, `black`, `pylint` | commented "Development dependencies (optional)" | **removed** | B has no tests and no lint configuration (READ: no `tests/`, `pytest.ini`, `setup.cfg` or `pyproject.toml` at any commit; DO) |
| Exact versions | none (`>=` only) | none | never pinned at any commit (READ) |

**Version upgrades (the method's question "deliberate or routine?").** There were none to explain: no minimum version was raised or lowered in any commit — `openpyxl>=3.1.2` and `lxml>=4.9.0` are the same from `edac6a3` to the snapshot, and `PyQt5>=5.15.0` is unchanged since it was added in `c258047` (READ; `git log -- requirements.txt` lists only those two commits). The only version-like line ever added is the commented `# pywin32>=306` (`c258047`), which `pip` ignores. The README's different figures (`openpyxl>=3.0.0`, `pywin32>=305`, `README.md:106-107`) were never in the file (12F.8). **Removed or replaced dependencies:** Tkinter (standard library) was replaced by PyQt5 (`c258047`, imports gone in `a62d3cc`; 12E.3), and the three commented development tools were dropped in `c258047`.

#### 12F.2 Each dependency's history

| Dependency | First needed (label) | Declared? | Where used (snapshot) | History and why |
|---|---|---|---|---|
| **openpyxl** (reads/writes `.xlsx` without Excel) | `edac6a3` (shared - before the split on 30-10-2025) | yes, `>=3.1.2` since `edac6a3` | `core/excel_handler.py:31-33` (processing summary workbook; report fallback), `power_query_extractor/core/data_consolidator.py:8-10` (consolidated extraction workbook) | unchanged in every commit. Before git it was also used to make the reports themselves, which corrupted templates / lost Power Query (MEDIUM-HIGH, 12E.1); since then Excel COM makes the reports and openpyxl only writes workbooks the program creates itself — except the fallback path (B-R97), which saves over a copy of the template and has been unreachable through the window since `c258047`. This environment ran openpyxl 3.1.5 (RAN (H-B1)). |
| **lxml** (fast XML library) | — | yes, `>=4.9.0` since `edac6a3` | **never imported** by any file at any commit (RAN (me) `imports_by_commit.txt`) | listed as "Optional but recommended for better performance"; openpyxl uses it when it is installed (RAN (verifier): `openpyxl.xml.LXML` is `True` with openpyxl 3.1.5 and lxml 6.1.3 in this environment; whether it makes B noticeably faster was not measured). |
| **PyQt5** (Qt 5 window toolkit) | `c258047` (B only (after the split)) | yes, `>=5.15.0` since `c258047` | `PyQt5.QtWidgets` (`gui/pyqt_main_window.py:11`, all views, `main.py:37`, `launch_extractor.py:13`), `PyQt5.QtCore` (controllers, views, widgets, `models/signals.py:6`), `PyQt5.QtGui` (`gui/pyqt_main_window.py:15`, views) | added for the rewrite; later commits only used more classes from it: `7e09e5c` added `QScrollArea`, `QSizePolicy`, `QSplitter`, `QTextCursor`, `Qt` to the Setup view, `QFrame` and `QScrollArea` to the Validation view (dropping `QFont` there) and `QFrame` to the Processing view (dropping `QColor` there; the added `QFrame` is itself unused, 12E.5) — [Corrected during verification — originally said both views gained `QFrame` / `QScrollArea`]; `b0e7d0e` added `QPixmap`, `QPainter`, `QColor` to the main window (tab dots) and `QColor`, `QFont` to the Extract view (RAN (me) `import_changes_per_file.txt`). This environment ran PyQt 5.15.11 / Qt 5.15.14 on Python 3.14 (RAN (H-B1)). |
| **pywin32** (`pythoncom`, `win32com.client`, `win32api`, `win32con`, `win32gui` — Windows COM and window functions) | `edac6a3` (shared) | **never declared**; a comment since `c258047` | guarded (wrapped so that a missing package is tolerated): `core/excel_handler.py:19-28`, `utils/helpers.py:473-483`; **unguarded** (a missing package stops the import): `power_query_extractor/core/report_processor.py:11-15` (+ again inside a function, `:710`) | At the split, only the separate extractor needed it; the organiser fell back to openpyxl without it. Since `c258047` the window imports the extraction controller, which imports `report_processor`, so **the program cannot start without pywin32** (RAN (me) on extracted trees: `ModuleNotFoundError: No module named 'pythoncom'` at `c258047`, `c21f788`, `b0e7d0e`; `47a1492`'s organiser window imports without it). The README of `d674386` shows a line `pywin32>=305; sys_platform == 'win32'` that the real file never had (12F.8). |
| **Windows Script Host** (`WScript.Shell`, a Windows component reached through COM) | `edac6a3` | not a Python package | `power_query_extractor/core/report_processor.py:216` (bring the Excel window to the front, send the keystrokes) | unchanged; needs a logged-in Windows desktop. |
| **Microsoft Excel** | `edac6a3` | not a Python package | Step 3 report creation and Step 4 refresh/extraction | unchanged; README asks for "Microsoft Excel 2016+" (`README.md:55`), which the code never checks. |
| **tkinter** (stdlib window toolkit) | `edac6a3` | stdlib (but absent from python.org's embeddable package, MEDIUM general knowledge) | none at the snapshot | used by the whole v3 GUI; still **loaded at start-up at `c258047`** because `gui/__init__.py` imported the Tk window (RAN (me): `'tkinter' in sys.modules` True at `c258047`); imports removed in `a62d3cc`; no longer loaded from `c21f788` on (RAN (me)). |
| **pytest / black / pylint** (test runner, formatter, checker) | never used | commented in `edac6a3`, removed in `c258047` | none | B has never had tests (12F.6). |

#### 12F.3 What the code imports versus what is declared (at the snapshot)

RAN (me) `imports_by_commit.txt` (HEAD list), READ of each location.

| Imported package | Where (`file:line`) | In `requirements.txt`? | Gap |
|---|---|---|---|
| `PyQt5.QtWidgets`, `PyQt5.QtCore`, `PyQt5.QtGui` | 14 files (RAN (verifier): every file with a `from PyQt5 …` line — the 3 controller modules, the 4 view modules, the 2 widget modules, the main window, the theme manager, `models/signals.py`, `main.py`, `launch_extractor.py`; the package `__init__.py` files only mention PyQt5 in their docstrings; [Corrected during verification — originally 13]), e.g. `gui/pyqt_main_window.py:11, 14, 15`, `main.py:37`, `launch_extractor.py:13` | yes (`requirements.txt:5`) | — |
| `openpyxl`, `openpyxl.styles`, `openpyxl.utils` | `core/excel_handler.py:31-33`; `power_query_extractor/core/data_consolidator.py:8-10` | yes (`:8`) | — |
| `pythoncom`, `win32com.client` | `core/excel_handler.py:23-24` (guarded); `power_query_extractor/core/report_processor.py:11-12` (unguarded) | **no** (comment only, `:13-14`) | **documentation gap and start-up blocker**: needed to open the window at all (B-R124) |
| `win32api` | `report_processor.py:13, 710`; `utils/helpers.py:477` (guarded) | **no** | same |
| `win32con`, `win32gui` | `report_processor.py:14-15` | **no** | same |
| `lxml` | nowhere | yes (`:11`) | declared but never imported (harmless) |
| stdlib: `collections`, `dataclasses`, `datetime`, `hashlib`, `io`, `json`, `logging`, `os`, `pathlib`, `platform`, `re`, `shutil`, `sys`, `tempfile`, `time`, `typing`, `warnings` | many (list in `imports_by_commit.txt`) | not needed | `dataclasses` needs Python 3.7 or later, which matches the "Python 3.7 or higher" line (`requirements.txt:2`); B was not tested on 3.7 (it ran here on 3.14) |

> **The most important dependency fact in B's history:** the only package without which the window cannot open — pywin32 — has never been a real line in `requirements.txt`, and the start-up error message tells the user to install two *other* packages (`main.py:51-53`: "pip install PyQt5 openpyxl"). RAN (me) 12B.7; RAN (H-B1) through `main.py`.

#### 12F.4 Standard-library modules added and removed over time

RAN (me) `import_changes_per_file.txt` (per-file import changes between consecutive commits).

| Commit (label) | File | Added | Removed | Why |
|---|---|---|---|---|
| `c258047` (B only (after the split)) | new `models/app_state.py:6`, `models/run_history.py:8` | `dataclasses` | — | the new settings and history containers |
| `c258047` | `gui/handlers/processing_handler.py` (Tk, deleted 2 min later) | — | `traceback` | debug output removed in the refactor |
| `c258047` | `main.py` | `PyQt5.QtWidgets.QApplication` (`:37`) | `os` | rewrite; folder creation removed |
| `c258047` | `launch_extractor.py` | `PyQt5.QtWidgets.QApplication` (`:13`) | `tkinter`, `tkinter.messagebox` | opens the PyQt5 window |
| `a62d3cc` (B only) | (deleted Tk files) | — | `threading`, `tkinter` (project-wide) | Tk layer deleted |
| `7e09e5c` (B only) | `core/file_parser.py:9` | `datetime` | — | the parser builds its own file-info record (modified/created times) |
| `b0e7d0e` (B only) | `gui/pyqt_main_window.py:9` | `datetime` | — | refreshed-file time stamps in the Step 4 tree (`:476-478`) |
| `b0e7d0e` | `gui/views/extraction_view.py:6` | `pathlib.Path` | — | version-folder paths in the tree |

No other stdlib module was added or removed in any commit (RAN (me)); no third-party package other than PyQt5 was added after `edac6a3` (RAN (me) `imports_by_commit.txt`: the shared commits `cbfa4a5` … `47a1492` add or remove nothing). Project-wide, the verifier's re-run (`stdlib_sets.py`, every outside import of every `.py` file at every commit) finds exactly three changes: `c258047` added `PyQt5` and `dataclasses` and removed `traceback`; `a62d3cc` removed `threading` and `tkinter`. The `datetime` and `pathlib` rows above are new imports *in a file*, of modules the project already used elsewhere; the many new files of `c258047` likewise import already-used modules (`json`, `logging`, `pathlib` …) without adding any to the project.

#### 12F.5 Internal dependencies that changed (which part of the program needs which)

The *import chain* is the sequence of files Python loads because one file imports the next. Its change in `c258047` is what turned pywin32 from "needed by the extractor" into "needed to start".

| Commit (label) | Start command | Chain (READ; RAN (me) on extracted trees) | pywin32 needed to open the window? |
|---|---|---|---|
| `47a1492` (shared - before the split on 30-10-2025) | `python main.py` | `main.py` → `gui/__init__.py` → `gui/main_window.py` (Tk) → `core/*`, `utils/*` (`core/excel_handler.py:19-28` guarded) | **no** (RAN (me): `import gui.main_window` works; `pythoncom` not loaded) |
| `47a1492` | `python launch_extractor.py` | → `power_query_extractor/extractor_main.py` → `power_query_extractor/gui/extractor_window.py` → `power_query_extractor/core/report_processor.py:11-15` (unguarded) | yes, for the extractor only |
| `c258047` (B only (after the split)) | `python main.py` | `main.py:38` → `gui/__init__.py` (still the Tk window → tkinter loaded) → `gui/pyqt_main_window.py:26@c258047` → `gui/controllers/extraction_controller.py:10` → `report_processor.py:11-15` | **yes** (RAN (me): fails without; works with stand-ins; tkinter loaded) |
| `a62d3cc` (B only) | any | `gui/__init__.py:7` → deleted `gui/main_window.py` | the program cannot start at all (RAN (me)) |
| `c21f788` → `b0e7d0e` (B only) | `python main.py` or `python launch_extractor.py` | `main.py:38` / `launch_extractor.py:14` → `gui/__init__.py:7` (eager: `from .pyqt_main_window import GSTOrganizerWindow`) → `gui/pyqt_main_window.py:26@c21f788` / `:26@7e09e5c` / `:27` at the snapshot (one line lower since `b0e7d0e` added `from datetime import datetime`; [Corrected during verification — originally `:27` for every commit]) → `gui/controllers/extraction_controller.py:10` → `report_processor.py:11-15` | **yes** (RAN (me): fails without at `c21f788` and `b0e7d0e`; tkinter not loaded) |

Other internal-dependency changes (READ):
- `7e09e5c` made `gui/views/setup_view.py:14` and `gui/views/validation_view.py:12` depend on `gui/widgets/colored_section.py`, which forced the repair of `gui/widgets/__init__.py` (12B.7).
- `7e09e5c` removed `core/file_parser.py`'s dependency on `utils/helpers.validate_excel_file` and `get_file_info` (`core/file_parser.py:14-16`), leaving both helpers without callers.
- `models/*` depend only on the standard library (and `models/signals.py` on PyQt5); `core/*` and `utils/*` have never imported anything from `gui/` or `models/` (READ), so the business core can still be used without the window — only `power_query_extractor/core/report_processor.py` needs pywin32 to be imported at all.

#### 12F.6 Development tooling and the AI assistant

- **Claude Code permission file.** `.claude/settings.local.json` was committed in `bd176bf` (shared; it allowed the tool to run `python -m py_compile` without asking; no secrets, READ (H-S)) and untracked in `c258047` (B only), with `.claude` added to `.gitignore:81-82`.
- **AI trailers.** `c258047`, `a62d3cc` and `c21f788` end with "Co-Authored-By: Claude Opus 4.6" (RAN (me) `git log` trailers); `7e09e5c` and `b0e7d0e` carry none and use a "feat:" prefix (MEDIUM that their messages were written by a tool from the diff, H-B2). Shared commits before the split carry no trailer, but the first commit's code holds paste markers from chat answers (12E.5) and `.gitignore` names the web sandbox folder `/mnt/user-data/` (`.gitignore:61`, added in `1cb878b`).
- **No tests, no checks.** B has never had a test file, a `pytest.ini`, a linter configuration or a CI (continuous-integration) set-up (READ every commit's file list, RAN (me) `filehist.txt`); the commented dev tools were removed in `c258047`. Consequence seen in the history: the dangling imports of `a62d3cc` were found only by running the program (12D D10–D11).

#### 12F.7 Comparison with the author's Standards 12 and 13 (facts only; decisions are the author's)

Standard 12 (`rebuild_pipeline_FINAL.md:817-898`, "Choosing the Language and Libraries", including "Excel through COM"):

| Standard 12 says | B at the snapshot | Since (label) |
|---|---|---|
| Desktop windows: PySide6 (`PySide6-Essentials`, Qt 6) | PyQt5 5.15 (Qt 5) (`requirements.txt:5`) | `c258047` (B only (after the split)); before: Tkinter |
| COM with `pywin32` for workbooks with Power Query | yes, but pywin32 not declared (12F.3) | `edac6a3` (shared) |
| `DispatchEx("Excel.Application")`, `Visible = False`, `DisplayAlerts = False`; never take over the user's own Excel | Step 3: `GetActiveObject` first (attaches to an open Excel), else `Dispatch`; sets `Visible = False` on it (`core/excel_handler.py:110-120`). Step 4 refresh: `Dispatch`, **`Visible = True`**, maximised (`power_query_extractor/core/report_processor.py:170-190`). Step 4 reading: `Dispatch`, hidden (`:484-494`). `DispatchEx` never appears in any commit (RAN (me) pickaxe). | `edac6a3`, unchanged |
| Always `Quit()` in a `finally` block | Step 3: `Quit()` in `finally` (`core/excel_handler.py:203-211`) — also on an Excel that was already open, because `self.excel_app` is never set; `DisplayAlerts` is switched back to True just before. Step 4: `Quit()` in `finally` (`report_processor.py:356-360`). | `edac6a3`, unchanged |
| Open with `IgnoreReadOnlyRecommended=True`; check `workbook.ReadOnly` | Step 3: neither; Step 4: only in the third open attempt (`report_processor.py:638-747`); `ReadOnly` never checked (B-R98, B-R99) | `edac6a3`, unchanged |
| Refresh: `BackgroundQuery = False`, `RefreshAll()`, `CalculateUntilAsyncQueriesDone()`, check no connection is `Refreshing`, then save | keystrokes Alt+A, R, A to a visible Excel, then fixed waits, then save (B-R67, B-R68) | method `edac6a3`; waits `bd176bf` (shared) |
| In a worker thread: `CoInitialize`, Excel work in a separate function, `CoUninitialize` after that function returns | `CoInitialize` / `CoUninitialize` inside the same function, while Excel objects are still held in local variables (B-R110) | `edac6a3`, unchanged |
| Never let openpyxl save over a workbook with Power Query | the openpyxl fallback saves over a template copy without a Power Query check (`core/excel_handler.py:219-262`), but cannot run through the window since `c258047` (B-R97) | `edac6a3` / `c258047` |

Standard 13 (`rebuild_pipeline_FINAL.md:900-1061`, "How Users Start the App"):

| Standard 13 says | B at the snapshot | Since (label) |
|---|---|---|
| A `Run [App Name].bat` launcher with `pythonw.exe -B -s`, started from a desktop shortcut | no launcher; README says `python main.py` (`README.md:112`) | never had one |
| `sys.excepthook` and `threading.excepthook` that write the error to `%LOCALAPPDATA%\[App Name]\logs` and show a message box | none; start-up errors are `print()`ed and written to `gst_organizer.log` in the current folder (`main.py:49-59`), never shown in a window; an error inside a window handler aborts the program (exit code 134, D9) | never had one |
| Works without a console (no `print()` reliance) | the console wrapper no longer crashes without a console (`main.py:12-13`), but start-up errors are visible only on a console or in `gst_organizer.log` | crash fixed in `c258047` (B only) |
| Find the app's files from `APP_DIR`, never from the current folder | `sys.path` uses the file's own folder (`main.py:26-29`), but the log is opened relative to the current folder (`main.py:15-23`) | log location unchanged since `edac6a3` |
| Settings and logs in `%LOCALAPPDATA%\[App Name]\`, never in the shared app folder | settings in `~/.gst_organizer/` (`models/cache_manager.py:16-17`); program log in the current folder; per-run extraction log in the run folder (`gui/controllers/extraction_controller.py:109-121`) | settings folder `c258047` (before: `~/.gst_organizer_cache.json`); extraction log `b0e7d0e` |
| No writes into the shared app folder | no longer creates folders there (`c258047`); still writes `gst_organizer.log` to the current folder, which under the Standard's launcher (`start "" /D "%~dp0" …`) is the app folder | `c258047` partly |
| Each app's `requirements.txt` with exact versions of the packages it imports | `>=` versions; pywin32 missing; lxml listed but unused | unchanged since `c258047` |
| Testing used Python 3.14 | B ran here on Python 3.14 with PyQt5 5.15.11 (RAN (H-B1)); "Python 3.7 or higher" in `requirements.txt:2` | — |

#### 12F.8 Where the older document (README.md) disagrees with the dependency history

| README says (line) | Dependency history shows |
|---|---|
| "Requirements.txt Contents: `openpyxl>=3.0.0`, `pywin32>=305; sys_platform == 'win32'`" (`README.md:104-108`) | Neither line was ever in `requirements.txt` (12F.1); the README also omits `PyQt5>=5.15.0` (added `c258047`) and `lxml>=4.9.0`. |
| "Operating System: Windows 10/11 (best), macOS, Linux" (`README.md:52`) | At the split, the organiser could run on macOS/Linux through the openpyxl fallback; since `c258047` the window cannot start without pywin32, which exists only for Windows (RAN (me) on Linux: `No module named 'pythoncom'`). |
| "Python: 3.7 or higher (3.8+ recommended)" (`README.md:51`) | Consistent with `requirements.txt:2`; not tested by any commit (no tests); the developer used Python 3.12 and 3.13 in 2025 (RAN (H-S) `.pyc` headers). |
| Troubleshooting: "No module named 'win32com'" → `pip install pywin32` and the post-install script (`README.md:275`) | The advice is right for the error, but B's own start-up message advises "pip install PyQt5 openpyxl" instead (`main.py:51-53`). |
| Debug mode: "`logging.basicConfig(level=logging.DEBUG)` in main.py or launch_extractor.py" (`README.md:283-286`) | `launch_extractor.py` has had no logging set-up since `c258047` (`launch_extractor.py:1-28`). |

#### 12F.9 Questions for the author (dependencies)

1. Should pywin32 become a real requirement line (for example only on Windows), or should the window be able to open without it, as the v3 organiser could?
2. Should the window toolkit stay PyQt5 (Qt 5), or follow Standard 12's PySide6 (Qt 6)? (A fact for your decision: they are different packages; switching is a code change, not a version bump.)
3. Should `requirements.txt` pin exact versions (Standard 13)? Which versions are installed on the office PCs today (openpyxl, PyQt5, pywin32)?
4. Is lxml wanted (it is installed but never imported)?
5. Which Excel version and display language do the office PCs use? The Alt+A, R, A keystrokes depend on both (B-R67).

#### 12F.10 Limits of the evidence in 12B–12F

- Excel through COM, SendKeys and the Windows dialog handling could not be run on this Linux machine; everything about them is READ (stand-in modules only prove import order). Check on a Windows test PC with Excel and pywin32, using made-up clients only.
- The pre-git history (June–August 2025) is known only from `.pyc` header fields read by H-S; the reasons for those changes are lost.
- Whether the shared commits' `.pyc` paths or the committed run log contain anything beyond what H-S counted was deliberately not examined (client confidentiality).
- "Real use" behind each fix is inferred from the diffs; no commit message in B describes a user-reported failure.
