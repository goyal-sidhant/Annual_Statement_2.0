### 12B. File Evolution Map

This sub-section traces every file of project A — the 70 files at the snapshot `bf2f1dd` (27-07-2026), the 72 further files that existed only in the history, and one file that was never committed but left traces — through all 46 commits of A's branch `main`: the 12 commits shared with project B (marked "shared - before the split on 30-10-2025") and A's own 34 (marked "A only (after the split)").

> **How to read the evidence marks.** READ = established by reading the code or a diff (the exact lines a commit added and removed). RAN = established by running something; "RAN (me)" means I ran it for this section in `/root/merge-scratch/run/A-S12BF/` (scripts named below; nothing was written into the repositories or the scratch copies of A); "RAN (H-S / H-A1 / H-A2 / H-A3)" means the history reader whose working notes are named ran it. Inferences carry HIGH / MEDIUM / LOW confidence. A citation `path:line` is the line at the snapshot `bf2f1dd`; `path:line@commit` is the line in that older commit's version of the file.
>
> **Labels.** The 12 commits `edac6a3`, `cbfa4a5`, `6237597`, `9c56a1c`, `d674386`, `5ea77c2`, `1cb878b`, `01862a0`, `1c48722`, `4c42758`, `bd176bf`, `47a1492` are always "shared - before the split on 30-10-2025"; every other commit named in 12B–12F is "A only (after the split)". Where a table cell or a sentence repeats a commit already labelled nearby and does not repeat the label, this rule supplies it. [Added during verification — some cells carried no label.]

**Terms used in this section (plain English).**
- *commit* = one saved change in the project's history; *diff* = the exact lines a commit added (+) and removed (−); *numstat* = git's count of lines added and removed per file.
- *branch* = one line of work. Project A has one branch, `main`. (The combined repository also holds a ref `A/claude/pensive-feynman-0mz1aa`; RAN (me) `git for-each-ref`: it points at `bf2f1dd` itself, so it is not a separate line of work.)
- *merge commit* = a commit that joins two lines of work. A has exactly one: `9c56a1c` (29-10-2025), which joined the author's local commit `6237597` with `cbfa4a5`, a deletion made three minutes after the first upload on GitHub's web page. Both parents are part of `main`'s history.
- *blame* = git's report of which commit last changed each line of a file; *pickaxe* (`git log -S`) = git's search for the commits that added or removed a given piece of text.
- *rename detection* = git's check whether a file was moved or renamed (it compares content, not names).
- *.pyc / `__pycache__`* = compiled copies of Python files that Python writes by itself; they record the source file's size, its last-modified time and the folder it came from.
- *module* = one `.py` file; *package* = a folder of modules with an `__init__.py` file; *import* = one file asking to use another file's code; *re-export* = a package's `__init__.py` importing names so they can be used from the package directly.
- *regex / pattern* = a text-matching rule for file names (`[^-]+` means "one or more characters that are not a hyphen"); *file search* (glob) = a wildcard search such as `ITC_Report_*.xlsx`; *directory listing* (`os.scandir`) = reading a folder's whole list of files in one request.
- *COM* = Windows' way for one program to drive another (here Python drives Excel); *pywin32* = the Python package that provides COM (`pythoncom`, `win32com`, `win32api`, `win32con`, `win32gui`); *openpyxl* = a Python library that reads and writes `.xlsx` files without Excel and drops Power Query when it saves; *Power Query* = Excel's data-import tool inside the ITC/Sales templates; *Links sheet* = the template sheet into which the program writes folder paths and file names.
- *Tk / Tkinter* = Python's built-in window toolkit; *ttk* = its themed widgets, coloured through shared *styles* (a *style map* holds the colours for states such as "selected"); *widget* = one on-screen element; *binding* = telling Tk to call a function when an event (a click, a mouse-wheel turn) happens.
- *worker thread* = a second line of work running beside the window so the window does not freeze; Tk may only be changed from the window's own thread, so workers hand screen updates over with `after(0, …)` (*marshalling*).
- *network share* = a folder on another computer or server; *round-trip* = one request to that computer and its answer (slow over a network or a VPN); *staging file* = a local temporary copy that is worked on and then copied to its final place; *date modified* (mtime) = the time a file or folder was last changed.
- *regression* = something that worked and was broken by a later change; *characterization test* = a test that records what the code does today so later changes cannot alter it silently; *patch* = a fix that removes the symptom; *proper fix* = a fix that removes the cause.
- *pip* = Python's package installer; *callback* = a function handed to another part of the program so that part can call it back (for example to report progress); *lambda* = a small unnamed function written in one line; *late binding* = a lambda reads the variables it uses when it finally runs, not when it is created, so inside a loop it can see a later loop value; *docstring* = the description text at the top of a function; *re-entry guard* = a check that refuses to start something that is already running.
- *8.3 short path* = the old DOS-style short form of a Windows path (for example `PROGRA~1`), used to get under Excel's path-length limit; *extended path* (`\\?\…`) = a Windows prefix that lifts the usual 260-character path limit; *VPN* = a secure connection to the office network from outside (here the product Tailscale).
- *`Dispatch` / `GetActiveObject`* = the two COM ways to get an Excel: `Dispatch` asks Windows for an Excel and may be handed one that is already running, `GetActiveObject` attaches to the Excel the user already has open (`DispatchEx`, which Standard 12 names, always starts a separate one); *DataMashup* = the hidden part of a workbook (inside a `customXml` file) where Power Query keeps its queries.
- *weak reference* = a reference that does not keep a destroyed widget alive; *panel* = a reusable block of screen elements that can sit inside any window; *shell* = a small window whose only job is to host a panel; *`__getattr__` forwarder* = a method that passes requests for unknown names on to another object (here from the shell to the panel); *free-threaded build* = a special Python build (`python3.14t`) without the usual global lock, which some packages do not support; *mutation check* = deliberately breaking the code to confirm that a test then fails; *Tk variable* (`IntVar`, `StringVar`) = the object that holds the value of an on-screen box; *`stat`* = asking the file system for a file's size, dates and type in one request.

**Scripts I ran for 12B (all RAN (me)).** `filehist.sh` → `filehist.txt` (every file at `bf2f1dd`: each commit that touched it, oldest first, with lines added/removed); `verify_imports.py` → `verify_imports.txt` (the method's VERIFICATION RULE, see 12B.7); `git log -M --diff-filter=R` and `git log -C -C --diff-filter=C` (renames and copies: both empty); size counts at six commits (12B.1).

#### 12B.1 How the project grew (sizes at six points in its history)

RAN (me): file counts from `git ls-tree`, line counts with `wc -l` (which counts one line less for a file whose last line has no line ending; the per-file table in 12B.3 uses git's own counts instead).

| Point in history | Commit (date) | Files (without compiled `.pyc`) | `.pyc` files | Program `.py` files / lines | Test `.py` files / lines |
|---|---|---|---|---|---|
| First commit | `edac6a3` (31-08-2025) | 40 | 70 | 37 / 6,917 | 0 / 0 |
| Split point | `47a1492` (30-10-2025) | 40 | 0 | 36 / 7,259 | 0 / 0 |
| After the two March audit documents | `f7eb8f1` (05-03-2026) | 41 | 0 | 36 / 7,259 | 0 / 0 |
| End of the first July session | `ba2a753` (23-07-2026) | 51 | 0 | 37 / 7,450 | 7 / 515 |
| End of the extractor split | `2aaa2b1` (25-07-2026) | 60 | 0 | 40 / 8,010 | 12 / 1,329 |
| Snapshot | `bf2f1dd` (27-07-2026) | 70 | 0 | 43 / 8,896 | 19 / 2,163 |

What the numbers show (READ of the diffs, HIGH):
- The program arrived in git already large and already modular (37 Python files in the first commit). Git holds no record of how it was built; the compiled leftovers in the first commit show about 2½ months of earlier work outside git (source files dated 19-06-2025 to 23-07-2025, run log entries up to 31-08-2025, first commit 31-08-2025) and a single-file window `gui/app.py` that was never committed (12B.4).
- Between the first commit and the split, the program barely changed in size (+342 lines): one feature commit for the extractor (`bd176bf`), one configuration commit (`47a1492`) and small fixes.
- In March 2026 only documentation was added (`fc279af`, `f7eb8f1`).
- Almost all of A's own growth happened in five days, 23-07-2026 to 27-07-2026 (32 commits, every one carrying the trailer "Co-Authored-By: Claude Opus 4.8 (1M context)" — RAN (me): `git log --grep`, 32 commits, first `a603c02`). Program code grew by about 1,640 lines and a test suite of 19 files and about 2,160 lines was created from nothing.

#### 12B.2 The architecture's evolution in phases

| Phase | Commits | What happened to the file structure |
|---|---|---|
| 0. Before git (19-06-2025 to 31-08-2025) | none (evidence: 70 `.pyc` files committed in `edac6a3`) | A single window file `gui/app.py` (89,275 bytes, last edited 25-06-2025) was replaced on or shortly before 23-07-2025 by the modular window (`gui/main_window.py` + `tabs/` + `handlers/` + `utils/` + `widgets/`); the extraction mapping had an older format (21-06-2025) kept as `cell_mappings copy.py`; the consolidator grew by about 73% in June–July 2025 (source size recorded in its compiled copies: 10,208 bytes on 25-06-2025 → 17,711 bytes on 14-07-2025). RAN (H-S, header fields of the `.pyc` files only; re-checked RAN (me) `VERIFY-A-S12BF/pyc.py`), HIGH for the sizes and dates. The split date is MEDIUM: the headers record only each source's *last-edit* time — every modular `gui/` source is dated 23-07-2025 and none earlier, so the split happened on that day or the modular files were last touched then. [Corrected during verification — originally "replaced on 23-07-2025 … nearly doubled in July 2025", both stated as HIGH.] |
| 1. First upload | `edac6a3`, `cbfa4a5` (31-08-2025, shared - before the split on 30-10-2025) | Two programs in one folder: the organiser (`main.py`, `core/`, `utils/`, `gui/`) and the separate Power Query Extractor (`launch_extractor.py`, `power_query_extractor/`). The run log and the compiled files were committed by accident; the log was deleted on GitHub three minutes later. |
| 2. Tidy-up and first review (29-10-2025) | `6237597`, `9c56a1c`, `d674386`, `5ea77c2`, `1cb878b`, `01862a0`, `1c48722`, `4c42758` (shared - before the split on 30-10-2025) | No new program files. `.gitignore` added; compiled files removed; four files touched by a review that was most likely AI-assisted (MEDIUM, from the message style); README rewritten. |
| 3. Extractor features (30-10-2025) | `bd176bf`, `47a1492` (shared - before the split on 30-10-2025) | The extractor window and processor grow (wait time, suffix, per-report ticks); six new extraction sheets; the old mapping copy deleted; the AI tool's permission file `.claude/settings.local.json` committed. **Split point.** |
| 4. Documentation only (March 2026) | `fc279af`, `f7eb8f1` (A only (after the split)) | `PROJECT_AUDIT_ANNUALSTATEMENT.md` added and rewritten. No code change. |
| 5. Network speed and correctness (23-07-2026, 00:22–01:54) | `a603c02` … `0890022` (A only (after the split)) | Existing files only; no new program file. First changes ever to `core/excel_handler.py`, `core/file_organizer.py`, `utils/constants.py` and `power_query_extractor/extractor_main.py`. |
| 6. Tests, then guarded fixes (23-07-2026, 07:58–08:54) | `6c1e1c6` … `ba2a753` (A only (after the split)) | New `tests/` folder, `pytest.ini`; first shared helper module created by **merging** duplicated code out of two files (`utils/excel_com.py`). |
| 7. Scan speed, new Setup tab, merge preparation (25-07-2026) | `4219ffb` … `2aaa2b1` (A only (after the split)) | Setup tab rewritten and its logic **split out** into `gui/utils/setup_validation.py` plus a new widget `gui/widgets/tooltip.py`; design note added; the extractor window **split** (inside one file) into a panel and a window shell; refreshed-file naming **split out** into `power_query_extractor/core/refresh_naming.py`. |
| 8. The two programs merged into one window, dependencies, dark mode (26-07-2026) | `7347383` … `169c7b5` (A only (after the split)) | New `gui/tabs/extract_tab.py` makes the extractor the fourth tab of the organiser — the two programs become one (**merge**); start-up concerns **split out** of `main.py` into `utils/requirements_check.py` and `utils/dependency_installer.py`; dark-mode manager rewritten. |
| 9. Last fix (27-07-2026) | `bf2f1dd` (A only (after the split)) | Scrolling fix and hover help; a guard test file. |

#### 12B.3 Every file at the snapshot

Columns: **Created** = commit and date that added the file; **Original purpose** = what the file did when it first appeared (READ of that version); **Changed in** = every later commit that touched it, with lines added/removed (RAN (me) `filehist.txt`); **Role** = whether its job expanded, shrank or shifted; **Lines** = git's line count when created → at `bf2f1dd`. No file at the snapshot was ever renamed or copied from another path (RAN (me): rename and copy detection both empty). Splits and merges of *content* between files are in 12B.5.

**Entry points, configuration, documents and tooling**

| File | Created | Original purpose | Changed in | Role | Lines |
|---|---|---|---|---|---|
| `main.py` | `edac6a3` 31-08-2025 (shared - before the split on 30-10-2025) | Start the organiser: re-wrap the console as UTF-8, configure logging to `gst_organizer.log`, create folders, import and run the window; on an import error list the source files that should exist, including the never-committed `gui/app.py` (`main.py:72@edac6a3`). | `5ea77c2` (−1: a malformed first line `# !/usr/bin/env python3` removed; shared - before the split on 30-10-2025); `f363946` (+40/−18), `be6418c` (+15/−3), `a5ad487` (+22/−1) (all A only (after the split)) | **Expanded**: from a launcher into a start-up gate that checks packages, explains what is missing, offers to install them and restarts itself (A-F1, rules A-R124, A-R125, A-R127). | 83 → 137 |
| `launch_extractor.py` | `edac6a3` (shared - before the split on 30-10-2025) | Start the separate Power Query Extractor window; show a Tk error box if it cannot start. | never | **Unchanged file, shifted context**: it starts `PowerQueryExtractor` (`launch_extractor.py:14-18`, class at `power_query_extractor/extractor_main.py:24`), which opens the window class `PowerQueryExtractorApp` (`extractor_main.py:68`); since `a677670` that window class is a thin shell around the panel that is also Step 4 of the main window (A-F29). [Corrected during verification — originally said the class it starts is `PowerQueryExtractorApp`; it starts it only indirectly.] | 30 → 30 |
| `requirements.txt` | `edac6a3` (shared - before the split on 30-10-2025) | List `openpyxl>=3.1.2` and `lxml>=4.9.0`; development tools commented out. | `6c1e1c6` (+3/−1: pytest made active), `064c634` (+4: pywin32 for Windows) (both A only (after the split)) | **Expanded** (see 12F). | 13 → 19 |
| `README.md` | `edac6a3` (shared - before the split on 30-10-2025) | User guide written for the single-file `gui/app.py` era. | `d674386` (+239/−135, shared - before the split on 30-10-2025) | Rewritten once for the modular organiser plus the extractor; **never updated after the split**, so it describes neither Step 4, the workflow choice, the start-up check, the tests nor the short folder names (12B.8). | 260 → 364 |
| `.gitignore` | `1cb878b` 29-10-2025 (shared - before the split on 30-10-2025) | Stop committing compiled files, logs, caches, output folders. | `6c1e1c6` (+1: `.pytest_cache/`, A only (after the split)) | Stable. | 79 → 80 |
| `.claude/settings.local.json` | `bd176bf` 30-10-2025 (shared - before the split on 30-10-2025) | Permission file of the AI coding tool Claude Code: allows `python -m py_compile` (a syntax check) without asking. No secrets (READ). | never | Stable; not used by the program. (B deleted it after the split; A keeps it.) | 9 → 9 |
| `PROJECT_AUDIT_ANNUALSTATEMENT.md` | `fc279af` 04-03-2026 (A only (after the split)) | An AI-written audit of A as it stood at `47a1492`, following an earlier form of the author's audit method. | `f7eb8f1` (+1003/−677, A only (after the split)) — a rewrite the next morning | **Shrank in truth**: never updated after 05-03-2026, so at the snapshot it calls problems "open" that July fixed and says many files were "never modified" that were changed later (12B.8). | 856 → 1182 |
| `docs/superpowers/specs/2026-07-25-merge-extractor-design.md` | `b276542` 25-07-2026 (A only (after the split)) | Design note for merging the extractor into the main window, in five stages. | never | Stable record of intent; partly built (12E.4). | 149 → 149 |
| `pytest.ini` | `6c1e1c6` 23-07-2026 (A only (after the split)) | Test-runner settings (`testpaths = tests`, `addopts = -q`). | never | Stable. | 6 → 6 |

**Core logic (`core/`)**

| File | Created | Original purpose | Changed in | Role | Lines |
|---|---|---|---|---|---|
| `core/__init__.py` | `edac6a3` (shared - before the split on 30-10-2025) | Re-export the three core classes; version label `3.0.0`. | never | Stable. | 13 → 13 |
| `core/file_parser.py` | `edac6a3` (shared - before the split on 30-10-2025) | Scan the source folder, recognise the six file-name patterns, group files per client, compute completeness (A-F7, A-F8, A-F9). | `6237597` (+11/−13: progress callback stored and protected, shared - before the split on 30-10-2025); `a603c02` (+6/−11: second validity check removed), `064c634` (+4/−4: four "[DEBUG]" lines moved to debug level), `56aaad9` (+17/−12: uses `scan_excel_files` and passes the cached file details on) (A only (after the split)) | **Stable job, slimmer path**: it no longer validates files itself; validation lives only in `utils/helpers.py`. | 373 → 371 |
| `core/file_organizer.py` | `edac6a3` (shared - before the split on 30-10-2025) | Create the `Annual Statement-…`/client/`Version-…`/category tree, copy files, write the per-client text report (A-F13, A-F16). | `275d9a4` (+18/−1: Level-1 folder worked out once per run, A only (after the split)) | Stable. Untouched for almost 11 months, then one performance change. | 409 → 426 |
| `core/excel_handler.py` | `edac6a3` (shared - before the split on 30-10-2025) | Create ITC/Sales reports from templates through Excel (COM, Windows' way for one program to drive another) or openpyxl, and write the summary workbook (A-F14, A-F15, A-F16). | `12d505f` (+72/−24: reports built in a local temporary file, then copied once), `064c634` (+31/−31: debug lines), `ba2a753` (+10/−6: shared "already open?" helper) (all A only (after the split)) | Stable job; **shifted** where the work happens (local disk instead of the network share, A-R98). | 661 → 713 |

**Shared utilities (`utils/`)**

| File | Created | Original purpose | Changed in | Role | Lines |
|---|---|---|---|---|---|
| `utils/__init__.py` | `edac6a3` (shared - before the split on 30-10-2025) | Re-export constants and helpers; version label. | `56aaad9` (+2/−2: re-exports `scan_excel_files`, A only (after the split)) | Stable but **stale**: names added later (`WORKFLOW_MODES`, the three new `utils` modules) are not re-exported; nothing in the program imports from the package root anyway (READ, `utils/__init__.py:22-36`). | 38 → 38 |
| `utils/constants.py` | `edac6a3` (shared - before the split on 30-10-2025) | All patterns, folder templates, Links-cell map, GUI colours, modes, messages, state map. | `5c8ce95` (+1/−1: IMS Reco pattern), `7347383` (+9/−3: short category folder names), `510f7b4` (+11: `WORKFLOW_MODES`) (all A only (after the split)) | Stable central configuration; only three changes in the whole history. | 275 → 292 |
| `utils/helpers.py` | `edac6a3` (shared - before the split on 30-10-2025) | Time stamps, path helpers, file-name cleaning, copy/backup, Excel signature check, state code, client-state key. | `6237597` (+1/−1, shared - before the split on 30-10-2025); `a603c02` (+19/−8), `33d0baf` (+6/−3), `4219ffb` (+13/−4), `56aaad9` (+77/−30), `4f84ba9` (+59/−22) (A only (after the split)) | **Expanded and shifted**: from a grab-bag of helpers into also the performance-tuned network scan (`scan_excel_files` with parallel signature reads, `utils/helpers.py:309-365`). The most-changed `utils` file. | 537 → 644 |
| `utils/excel_com.py` | `ba2a753` 23-07-2026 (A only (after the split)) | One shared "is this workbook already open in Excel?" function, **merged** out of two identical loops. | never | Stable. | 29 → 29 |
| `utils/requirements_check.py` | `f363946` 26-07-2026 (A only (after the split)) | Start-up package check with plain-English report (pywin32 optional). | `be6418c` (+23/−3: pywin32 essential on Windows, A only (after the split)) | Stable after one reversal (12E.2). | 113 → 133 |
| `utils/dependency_installer.py` | `a5ad487` 26-07-2026 (A only (after the split)) | Offer to pip-install missing packages and restart. | never | Stable. | 158 → 158 |

**The organiser's window (`gui/`)**

| File | Created | Original purpose | Changed in | Role | Lines |
|---|---|---|---|---|---|
| `gui/__init__.py` | `edac6a3` (shared - before the split on 30-10-2025) | Re-export `GSTOrganizerApp`. | never | Stable; side effect grew: since `bfcb24a` importing `gui` also loads the whole extractor (READ). | 11 → 11 |
| `gui/main_window.py` | `edac6a3` (shared - before the split on 30-10-2025) | "Modularized version - ties all components together" (`gui/main_window.py:4`): builds the window, three tabs, handlers, cache, dark mode. | `bfcb24a` (+9/−2: fourth tab, safe tab names), `510f7b4` (+6: workflow setting), `c89c1d4` (+16: recompute Setup colours after a theme switch), `725fbaf` (+7/−3: a padding hack from the first commit removed) (all A only (after the split)) | **Expanded** from three steps to four; became the place that re-applies state colours after a theme change. | 365 → 398 |
| `gui/tabs/__init__.py` | `edac6a3` (shared - before the split on 30-10-2025) | Re-export the three tab classes. | never | **Stale**: its `__all__` (the list of names a module officially offers) lacks `ExtractTab` (READ, `gui/tabs/__init__.py`); harmless because nothing imports its listed names (one test reaches the submodule through the package, `from gui.tabs import extract_tab`, `tests/test_optional_pywin32.py:54`, which works whatever `__all__` says; RAN (me) grep during verification). | 8 → 8 |
| `gui/tabs/setup_tab.py` | `edac6a3` (shared - before the split on 30-10-2025) | Step 1 form: four paths with Browse buttons, modes, include-name tick, length box, Scan/Re-scan, help column with six hard-coded name patterns. | `322717f` (+370/−349: rewritten as a live checklist), `d64d3a9` (+59/−7: scrolling fixed), `510f7b4` (+39/−1: workflow card), `169c7b5` (+68/−18: dark-mode colours), `bf2f1dd` (+85/−8: hover help) (all A only (after the split)) | **Rewritten and expanded** (A-F3, A-F4, A-F5, A-F6, A-F31); logic **split out** to `gui/utils/setup_validation.py`. Most-changed tab. | 377 → 615 |
| `gui/tabs/validation_tab.py` | `edac6a3` (shared - before the split on 30-10-2025) | Step 2: summary, client table, Dry Run / Export / Start buttons (A-F8, A-F11, A-F12). | never | Stable (its `.bind('<Button-1>')` buttons are the old style that the newer tabs abandoned, A-R135). | 247 → 247 |
| `gui/tabs/processing_tab.py` | `edac6a3` (shared - before the split on 30-10-2025) | Step 3: progress, log, Start/Stop (A-F17). | never | Stable. | 117 → 117 |
| `gui/tabs/extract_tab.py` | `bfcb24a` 26-07-2026 (A only (after the split)) | Host the extractor panel as "Step 4: Extract", following Step 1's target folder (A-F18). | `c463256` (+44/−1: guarded import and "unavailable" notice, A only (after the split)) | Stable after one fix. | 88 → 131 |
| `gui/handlers/__init__.py` | `edac6a3` (shared - before the split on 30-10-2025) | Re-export the three handlers. | never | Stable. | 8 → 8 |
| `gui/handlers/file_handler.py` | `edac6a3` (shared - before the split on 30-10-2025) | Browse buttons, scan, input checks. | `6237597` (+1/−1, no behaviour change, shared - before the split on 30-10-2025) | Stable. | 151 → 151 |
| `gui/handlers/client_handler.py` | `edac6a3` (shared - before the split on 30-10-2025) | Client table, ticks, per-client folder-name toggle, export. | `1c48722` (−6: row colours fixed, shared - before the split on 30-10-2025); `33d0baf` (+6/−2: per-client key by state code, A only (after the split)) | Stable. | 259 → 257 |
| `gui/handlers/processing_handler.py` | `edac6a3` (shared - before the split on 30-10-2025) | Dry Run, Start, the Step 3 worker thread (a second line of work beside the window). | `6237597` (+46/−30, shared - before the split on 30-10-2025); `0890022` (+40/−25), `33d0baf` (+7), `510f7b4` (+46/−1) (A only (after the split)) | **Expanded**: gained the hand-over to Step 4 (A-F5). | 447 → 530 |
| `gui/utils/__init__.py` | `edac6a3` (shared - before the split on 30-10-2025) | Re-export four helper classes. | never | Stable. | 9 → 9 |
| `gui/utils/cache_manager.py` | `edac6a3` (shared - before the split on 30-10-2025) | Read/write `~/.gst_organizer_cache.json` (A-F27). | never | **Unchanged file, shifted role**: since `c1583d6` the extractor reads the file it writes (A-R118); before that nobody else could find it. | 56 → 56 |
| `gui/utils/dark_mode_manager.py` | `edac6a3` (shared - before the split on 30-10-2025) | Darken widgets class by class and guess the light colours back (A-F28). | `aeb83cf` (+258/−183: rewritten to capture and replay exact colours), `725fbaf` (+10/−3) (A only (after the split)) | **Replaced in place** (same file, new design). | 212 → 294 |
| `gui/utils/setup_validation.py` | `322717f` 25-07-2026 (A only (after the split)) | Pure-logic readiness rules for Step 1 (no Tkinter), **split out** of the Setup tab. | never | Stable. | 115 → 115 |
| `gui/utils/status_bar.py` | `edac6a3` (shared - before the split on 30-10-2025) | Status bar "v3.0 \| Production Ready". | never | Stable. | 37 → 37 |
| `gui/utils/ui_helpers.py` | `edac6a3` (shared - before the split on 30-10-2025) | Log-box helpers. | never | Stable. | 63 → 63 |
| `gui/widgets/__init__.py` | `edac6a3` (shared - before the split on 30-10-2025) | Re-export two widgets. | never | **Stale**: lacks `Tooltip` (READ); harmless. | 7 → 7 |
| `gui/widgets/collapsible_frame.py` | `edac6a3` (shared - before the split on 30-10-2025) | Collapsible box for Step 2 instructions. | never | Stable. | 41 → 41 |
| `gui/widgets/title_bar.py` | `edac6a3` (shared - before the split on 30-10-2025) | Blue title strip with the Dark Mode tick. | never | Stable. | 51 → 51 |
| `gui/widgets/tooltip.py` | `322717f` (A only (after the split)) | Hover tooltip. | never | **Use expanded** without the file changing: `322717f` attached it to the four path boxes (`gui/tabs/setup_tab.py:381`); `bf2f1dd` attached it to 29 more Step 1 widgets — 16 checklist parts, 2 workflow radios, 4 Browse buttons, 3 mode radios, the include-name tick, the length box, Scan and Re-scan (READ, `gui/tabs/setup_tab.py:314, 336, 386, 409, 423, 444, 462, 467`). | 69 → 69 |

**The Power Query Extractor (`power_query_extractor/`)**

| File | Created | Original purpose | Changed in | Role | Lines |
|---|---|---|---|---|---|
| `power_query_extractor/__init__.py` and `config/`, `core/`, `gui/` `__init__.py` | `edac6a3` (shared - before the split on 30-10-2025) | Package markers; version label `1.0.0`. | never | Stable. | 3/1/1/1 → same |
| `power_query_extractor/extractor_main.py` | `edac6a3` (shared - before the split on 30-10-2025) | Standalone start-up class; logging to `pq_extractor.log`; look for the main program's settings file (it never found it, A-R118). | `c1583d6` (+3: home-folder settings file tried first, A only (after the split)) | **Shrank in use**: since `bfcb24a` the main window bypasses it on purpose (it configures logging when loaded, `gui/tabs/extract_tab.py:8-10`); only `launch_extractor.py` uses it. | 79 → 82 |
| `power_query_extractor/config/cell_mappings.py` | `edac6a3` (shared - before the split on 30-10-2025) | Which report cells to extract (3 output sheets, 8 cells) and options [GST RULE] ⚠. | `47a1492` (+126: six ITC output sheets, shared - before the split on 30-10-2025) | Stable since the split point. | 152 → 278 |
| `power_query_extractor/core/report_processor.py` | `edac6a3` (shared - before the split on 30-10-2025) | Refresh each report in a visible Excel by key presses, check `Info!BB2`, read the configured cells (A-F21, A-F23). | `bd176bf` (+41/−17, shared - before the split on 30-10-2025); `064c634` (−2: stray `print` removed), `ba2a753` (+9/−5), `2aaa2b1` (+10/−2), `c463256` (+39/−5: pywin32 imports guarded) (A only (after the split)) | Stable engine; **shifted** from "Windows-only module that fails to load without pywin32" to "loads anywhere, refuses per client" (A-R64). | 736 → 804 |
| `power_query_extractor/core/data_consolidator.py` | `edac6a3` (shared - before the split on 30-10-2025) | Write the consolidated extraction workbook (A-F24). | never | **Unchanged file whose behaviour changed**: its column lookup became harmful when `47a1492` added columns named `IGST`/`CGST`/`SGST` (A-R88; RAN (me), 12C.2). | 432 → 432 |
| `power_query_extractor/core/refresh_naming.py` | `2aaa2b1` 25-07-2026 (A only (after the split)) | One place that decides "is this a refreshed copy?", **split out** of two hard-coded tests. | never | Stable. | 66 → 66 |
| `power_query_extractor/gui/extractor_window.py` | `edac6a3` (shared - before the split on 30-10-2025) | The standalone extractor window (a `tk.Tk` window class) with Setup and Processing tabs. | `4c42758` (+30), `bd176bf` (+362/−145) (shared - before the split on 30-10-2025); `275d9a4` (+72/−27), `c1583d6` (+10/−2), `c1f425e` (+58/−20), `2acffed` (+36/−2), `a677670` (+89/−43), `2aaa2b1` (+20/−8), `bfcb24a` (+10/−3), `bf2f1dd` (+42/−17) (A only (after the split)) | **The most-changed file of A (11 commits)**: from a separate program's window into a reusable panel (`ExtractorPanel`, `extractor_window.py:98`) plus a thin window shell (`PowerQueryExtractorApp`, `:1032`), hosted both standalone and as Step 4. | 601 → 1063 |

**Tests (`tests/`) — all created after the split (A only (after the split))**

| File | Created | Purpose | Changed in | Lines |
|---|---|---|---|---|
| `tests/README.md` | `6c1e1c6` 23-07-2026 | How to run the first four test files. | never — **stale**: lists 4 of today's 18 test files (12B.8) | 34 → 34 |
| `tests/conftest.py` | `6c1e1c6` | Path set-up, fake-Excel helper. | `a677670` (+23: one shared Tk root for all window tests) | 29 → 52 |
| `tests/test_file_patterns.py`, `tests/test_scan.py` | `6c1e1c6` | File-name patterns; end-to-end scan. | never | 94; 80 |
| `tests/test_helpers.py` | `6c1e1c6` | Pure helpers (cleaning, state codes, keys). | `33d0baf` (+6/−5), `4219ffb` (+26) | 154 → 181 |
| `tests/test_file_ops.py` | `6c1e1c6` | Validation and discovery against real temporary files. | `56aaad9` (+57/−1), `4f84ba9` (+24) | 89 → 169 |
| `tests/test_processing_handler.py` | `33d0baf` | The Start re-entry guard. | never | 36 |
| `tests/test_excel_com.py` | `ba2a753` | The shared "already open?" helper with fake objects. | never | 32 |
| `tests/test_setup_validation.py`, `tests/test_setup_tab_smoke.py` | `322717f` | Setup readiness rules; the real Setup tab. | smoke file: `d64d3a9` (+93/−1), `a677670` (+4/−18), `510f7b4` (+1) | 112; 131 → 210 |
| `tests/test_extractor_panel.py`, `tests/test_extractor_standalone.py` | `a677670` | Panel embedded anywhere; standalone shell in a separate process. | panel file: `2aaa2b1` (+29) | 108 → 137; 106 |
| `tests/test_refresh_naming.py` | `2aaa2b1` | Refreshed-name rules. | never | 121 |
| `tests/test_main_window_tabs.py` | `bfcb24a` | Four tabs, Step 4 pre-filled, no start-up scan. | never | 140 |
| `tests/test_workflow_mode.py` | `510f7b4` | Organise-only vs full pipeline. | never | 120 |
| `tests/test_dark_mode.py` | `aeb83cf` | Colour round trips. | `c89c1d4` (+33) | 165 → 198 |
| `tests/test_optional_pywin32.py` | `c463256` | App starts with pywin32 blocked. | never | 101 |
| `tests/test_requirements_check.py` | `f363946` | Start-up check. | `be6418c` (+10/−3: pywin32 "optional" test replaced by "required on Windows") | 88 → 95 |
| `tests/test_dependency_installer.py` | `a5ad487` | Install offer. | never | 115 |
| `tests/test_scroll_bindings.py` | `bf2f1dd` | Text checks that no window file removes other wheel bindings. | never | 64 |

#### 12B.4 Files that existed in the history but not at the snapshot

| File(s) | Added | Removed | What it was | What it reveals |
|---|---|---|---|---|
| `gst_organizer.log` (5,689 lines) | `edac6a3` 31-08-2025 22:11 (shared - before the split on 30-10-2025) | `cbfa4a5` 31-08-2025 22:14, made on GitHub's web page; reached the author's computer through the merge `9c56a1c` (29-10-2025) (shared - before the split on 30-10-2025) | The organiser's own run log, written to the folder the program was started from (`main.py:24`) and committed because no `.gitignore` existed yet. **Described by kind only** (not opened): it holds client names, state names, source file names and full Windows folder paths from real runs between 23-07-2025 and 31-08-2025, all at INFO level, no tax figures (the organiser never reads cell values) — per H-S §1.15, which counted message shapes only. | The log file's location in the start folder is a design that still exists (A-R120, gap A-G56 in Section 7). Deleting a file in a later commit does not remove it from history: it is still readable at `edac6a3` in both repositories (HIGH). |
| 70 compiled files `**/__pycache__/*.cpython-312.pyc` and `*.cpython-313.pyc` (35 modules × 2 Python versions) | `edac6a3` (shared - before the split on 30-10-2025) | `01862a0` 29-10-2025, right after `1cb878b` added `.gitignore` (shared - before the split on 30-10-2025) | Build leftovers for `core/` (4 modules), `gui/` (`__init__`, `app`, `main_window`), `gui/handlers/` (4), `gui/tabs/` (4), `gui/utils/` (5), `gui/widgets/` (3), `power_query_extractor/` (2), its `config/` (2), `core/` (3), `gui/` (2), and `utils/` (3) (RAN (me): `git ls-tree edac6a3`). | They embed the developer computer's folder paths (not reproduced here) and the source files' sizes and dates, which date the pre-git work (RAN (H-S), header fields only; HIGH): sources edited on Windows between 19-06-2025 and 23-07-2025; two interpreters, Python 3.12 and 3.13, in use (MEDIUM). |
| `gui/app.py` | **never committed** | — | Known only from `gui/__pycache__/app.cpython-312.pyc` / `-313.pyc` in `edac6a3`: a source of 89,275 bytes last edited 25-06-2025; its identifier strings include `GSTOrganizerApp`, `create_setup_tab`, `create_validation_tab`, `create_processing_tab`, `create_title_bar`, `restore_light_mode` (RAN (A-X-DEAD notes), `strings` on the compiled file). | The window began as one large file and was **split** into the modular `gui/` package before git (HIGH), on or shortly before 23-07-2025 (MEDIUM: the compiled headers give each modular source's last-edit time, all 23-07-2025, not its creation time). The README at `edac6a3` still showed `gui/app.py` in its folder tree, and `main.py` listed it in its error message until `f363946` (26-07-2026). |
| `power_query_extractor/config/cell_mappings copy.py` (63 lines) | `edac6a3` (shared - before the split on 30-10-2025) | `47a1492` 30-10-2025 (shared - before the split on 30-10-2025) | The older format of the extraction mapping (a sheet → cell → field dictionary), dated 21-06-2025 by the compiled-file evidence; its file name contains a space, so Python could never import it (READ). [GST RULE] ⚠ — it named the same cells `Diff GSTR-3B!J15/K15/L15` and `Diff - GSTR-1!N1/O1/P1/Q1` and four commented-out candidate cells (`'E5': 'reversed_itc'`, `'F5': 'net_itc'`, `'D10': 'export_sales'`, `'E10': 'exempt_sales'`) that were never active (READ `cell_mappings copy.py:10-31@edac6a3`). | A manual backup copy made before an edit, kept for two months, then deleted together with the configuration change that added six sheets. See 12E.1. |

#### 12B.5 Renames, splits and merges

**Renames and copies: none.** RAN (me): `git log -M50% --diff-filter=R` and `git log -C -C --diff-filter=C` over `A/main` return nothing. Every file kept its path from the day it was added.

**Splits and merges of content** (READ of the diffs; HIGH unless stated):

| Kind | When | From → to | Evidence |
|---|---|---|---|
| Split (before git) | 23-07-2025 or shortly before (last-edit dates only, MEDIUM) | one window file `gui/app.py` → `gui/main_window.py` + `gui/tabs/*` + `gui/handlers/*` + `gui/utils/*` + `gui/widgets/*` | compiled leftovers (12B.4); `gui/main_window.py:4` "Modularized version - ties all components together" |
| Split | `a603c02` (23-07-2026) | file validation done twice (in `utils/helpers.py` and again in `core/file_parser.py`) → only in `utils/helpers.py` | the deleted branch "Invalid or corrupted Excel file" in `core/file_parser.py` and the comment at `core/file_parser.py:190-193` |
| Split (one function into two) | `4f84ba9` (25-07-2026) | `validate_excel_file` → `_passes_metadata_checks` + `_has_excel_signature`, so only the network read runs in parallel | `utils/helpers.py:240-306` |
| Merge (duplicate code into one helper) | `ba2a753` (23-07-2026) | the same "already open?" loop inside `core/excel_handler.py` and `power_query_extractor/core/report_processor.py` → `utils/excel_com.py:find_open_workbook` | `utils/excel_com.py:1-29`; callers `core/excel_handler.py:333-340`, `power_query_extractor/core/report_processor.py:705-712` |
| Split (logic out of a window file) | `322717f` (25-07-2026) | Step 1 checks inside `gui/tabs/setup_tab.py` (pop-ups) → pure rules in `gui/utils/setup_validation.py` (no Tkinter) + a reusable `gui/widgets/tooltip.py` | `gui/utils/setup_validation.py:1-115` |
| Split (inside one file) | `a677670` (25-07-2026) | one window class `PowerQueryExtractorApp(tk.Tk)` → `ExtractorPanel(ttk.Frame)` (`extractor_window.py:98`) + a thin shell `PowerQueryExtractorApp` (`:1032`) + `configure_extractor_styles()` (`:85-95`) | "one implementation, two homes" (commit message) |
| Split (rule out of two places) | `2aaa2b1` (25-07-2026) | hard-coded `_Refreshed_` tests in `report_processor.py` and `extractor_window.py` → `power_query_extractor/core/refresh_naming.py` | `refresh_naming.py:8-11` explains the reason |
| **Merge of the two programs** | `bfcb24a` (26-07-2026) | the standalone extractor (`launch_extractor.py`) and the organiser (`main.py`) → one window with four steps (`gui/tabs/extract_tab.py`); the standalone path still works | `gui/main_window.py:144-146`; `launch_extractor.py` unchanged |
| Split (start-up concerns out of `main.py`) | `f363946`, `a5ad487` (26-07-2026) | inline import/error handling in `main.py` → `utils/requirements_check.py` (no Tkinter) and `utils/dependency_installer.py` | `main.py:66-106` |
| Split (one method into a cached wrapper and a body) | `275d9a4` (23-07-2026) | `FileOrganizer._get_level1_folder` → `_get_level1_folder` (cache) + `_compute_level1_folder` (old body) | `core/file_organizer.py:150-196` |

#### 12B.6 How each core file changed, commit by commit

**`main.py`** — `edac6a3`: 83 lines: UTF-8 console re-wrap at load time (`main.py:17`, still there), logging set-up (`:20-27`, still there), `setup_environment()` creating `core/`, `utils/`, `gui/`, `logs/`, `temp/` in the program folder (`:29-43`, still there), and an import-error message listing source files including `gui/app.py`. `5ea77c2` (shared - before the split on 30-10-2025): removed the malformed first line. `6237597`'s message (shared - before the split on 30-10-2025) claims to have "Corrected non-existent app.py reference in main.py error messages", but `main.py` is not in that commit (RAN (me): `git log --name-status`: `6237597` touched only `core/file_parser.py`, `gui/handlers/file_handler.py`, `gui/handlers/processing_handler.py`, `utils/helpers.py`) — a message-versus-diff mismatch (HIGH). `f363946` (A only (after the split)): the package check runs before the window is imported; the error message no longer lists source files. `be6418c`: `_pause()` (`main.py:45-54`) replaces bare `input()`. `a5ad487`: the install offer and restart (`main.py:75-95`). Unchanged since `a5ad487`.

**`core/file_parser.py`** — `edac6a3`: patterns applied in order; per-client grouping; completeness; a second `validate_excel_file` call inside `_process_file` that could add a variation "Invalid or corrupted Excel file" (unreachable, because the discovery function had already filtered the same files — RAN (H-A1) reading of the caller). `6237597`: the progress callback is stored on the object and called through `_update_progress` (`core/file_parser.py:48-54`), so a failing callback no longer aborts the scan. `a603c02`: the duplicate validation and its variation removed. `064c634`: four Sales "[DEBUG]" messages moved from INFO to DEBUG. `56aaad9`: imports `scan_excel_files` and passes each file's cached details to `get_file_info` (`core/file_parser.py:186-188`).

**`core/file_organizer.py`** — `edac6a3` until `275d9a4` (23-07-2026): unchanged. `275d9a4` caches the Level-1 folder per run (`core/file_organizer.py:45-47, 150-161`). Side effect (RAN (H-A1)): the run-wide "Folders Created" counter in each client's text report changed from 6/12/18 to 6/11/16 for three clients in fresh mode, because the Level-1 folder is no longer appended once per client (display only; `core/file_organizer.py:368`).

**`core/excel_handler.py`** — `edac6a3` until `12d505f`: unchanged. `12d505f`: the template is copied to `%TEMP%\gst_excel_temp\` (the user's temporary folder), edited there and copied to the client folder once (`core/excel_handler.py:135-155, 212-215`; openpyxl path `:269-322`). `064c634`: 31 "[DEBUG]" INFO lines → DEBUG (one `logger.error("[DEBUG] …")` remains, `core/excel_handler.py:83`). `ba2a753`: `find_open_workbook` replaces the inline loop (`:41`, `:333-340`). Never changed: attaching to a running Excel with `GetActiveObject` and always calling `Quit()` (`:123-128`, `:238-246`; A-R97).

**`utils/constants.py`** — three changes in 46 commits: `5c8ce95` IMS Reco pattern (`utils/constants.py:23`); `7347383` short category names plus a 6-line comment with the measured path lengths (`:85-95`); `510f7b4` `WORKFLOW_MODES` (`:163-172`). Everything else — the six patterns' other parts, the state map, the Links-cell map, the modes, the unused constants — is exactly as in `edac6a3` (RAN (me) blame, 12C).

**`utils/helpers.py`** — `6237597` relative import inside `get_state_code` (`utils/helpers.py:598`). `a603c02` one directory listing instead of three file searches; type from one file query. `33d0baf` and `4219ffb` the abbreviation rule (A-R19). `56aaad9` `scan_excel_files` with cached file details. `4f84ba9` parallel signature reads (`:263-365`).

**`gui/main_window.py`** — unchanged from `edac6a3` until 26-07-2026; then four commits all tied to the merge and to dark mode (table above). The tab-name list was a hard-coded three-item list that raised an error once a fourth tab existed; `bfcb24a` made the lookup defensive (`gui/main_window.py:331-338`).

**`gui/tabs/setup_tab.py`** — `edac6a3` form (two columns, `.bind('<Button-1>')` buttons, an application-wide mouse-wheel binding that was never released, a length validator allowing 10–100, a hard-coded list of six name patterns). `322717f` rewrote it as a checklist screen with `command=` buttons; `d64d3a9` replaced its broken Enter/Leave wheel toggle with one filtered handler; `510f7b4` added the workflow card; `169c7b5` dark colour pairs; `bf2f1dd` hover help.

**`gui/handlers/processing_handler.py`** — `6237597` wrapped part of the worker's screen updates in `root.after(0, lambda: …)` (which introduced wrong names in the log, 12D) and tied the long-name warning to the length setting; `0890022` replaced those wrappers with `_log_async` / `_progress_async` (`processing_handler.py:23-36`); `33d0baf` added the re-entry guard (`:71-76`); `510f7b4` added the Step 4 hand-over (`:136-175`, called at `:519`).

**`gui/utils/dark_mode_manager.py`** — `edac6a3` class-by-class darkening with guessed restore; untouched until `aeb83cf` rewrote it to capture each widget's colours and replay them; `725fbaf` also records empty style values (`dark_mode_manager.py:87-102`).

**`power_query_extractor/gui/extractor_window.py`** — `4c42758` wheel scrolling (Enter → bind for the whole application, Leave → unbind for the whole application); `bd176bf` two-column Setup tab with wait time, suffix, per-report ticks and "Last Refresh Status"; `275d9a4` two single-listing helpers (`:20-82`); `c1583d6` home-folder settings file (`:118`) and newest Version folder by date modified (`:648`); `c1f425e` every worker-thread screen update through `_ui` (`:989-1010`); `2acffed` refresh times recomputed after a run; `a677670` panel + shell split; `2aaa2b1` configured suffix; `bfcb24a` `auto_load` flag (`:107-139`); `bf2f1dd` the wheel handler rebuilt (`:434-478`).

**`power_query_extractor/core/report_processor.py`** — `bd176bf` wait time and suffix attributes (`:50-51`), per-report selection, configurable waits; `064c634` removed a `print` placed above the refresh function's docstring; `ba2a753` shared "already open?" helper (`:39`, `:709`); `2aaa2b1` skip-refresh lookup by configured suffix (`:137-159`); `c463256` guarded pywin32 imports and a per-client refusal (`:11-36`, `:62-77`).

**`power_query_extractor/config/cell_mappings.py`** — one change in its life: `47a1492` (shared - before the split on 30-10-2025) added six ITC output sheets (see 12C, A-R77 to A-R82) [GST RULE] ⚠.

**`power_query_extractor/core/data_consolidator.py`** — never changed since `edac6a3`. Its two dead methods (`:250-352`, `:354-414`) and its doubled header block with "# ... rest of the code" markers (`:73-83`) date from before git (12E.5).

#### 12B.7 VERIFICATION RULE (Project_Audit_FINAL.md, 12B): does any "never changed" file import something added later?

Method (RAN (me), `verify_imports.py` → `verify_imports.txt`): for every `.py` file at `bf2f1dd` I listed each import of a project module or name, found the first `A/main` commit in which the source module defines that name (by parsing every historical version of the source module), and compared it with the importing file's own change history. A file whose last change is older than a name it imports would contradict "never changed".

**Result: no contradiction.** All 19 imports of names that appeared after the importing file was created belong to files that were changed in exactly the commit that introduced the name:

| Importing file | Imported name (module) | Name introduced in | Importing file changed in that commit? |
|---|---|---|---|
| `main.py` | `check_requirements`, `format_report`, `interpreter_hint`, `missing_essential`, `missing_optional` (`utils/requirements_check.py`) | `f363946` | yes (`f363946`) |
| `main.py` | `offer_install`, `restart_application` (`utils/dependency_installer.py`) | `a5ad487` | yes (`a5ad487`) |
| `core/excel_handler.py` | `find_open_workbook` (`utils/excel_com.py`) | `ba2a753` | yes (`ba2a753`) |
| `core/file_parser.py` | `scan_excel_files` (`utils/helpers.py`) | `56aaad9` | yes (`56aaad9`) |
| `utils/__init__.py` | `scan_excel_files` | `56aaad9` | yes (`56aaad9`) |
| `tests/test_file_ops.py` | `scan_excel_files` | `56aaad9` | yes (`56aaad9`) |
| `gui/main_window.py` | `ExtractTab` (`gui/tabs/extract_tab.py`) | `bfcb24a` | yes (`bfcb24a`) |
| `gui/tabs/setup_tab.py` | `Tooltip` (`gui/widgets/tooltip.py`); module `gui/utils/setup_validation.py` | `322717f` | yes (`322717f`) |
| `gui/tabs/setup_tab.py` | `WORKFLOW_MODES` (`utils/constants.py`) | `510f7b4` | yes (`510f7b4`) |
| `power_query_extractor/core/report_processor.py` | `find_open_workbook` | `ba2a753` | yes (`ba2a753`) |
| `power_query_extractor/core/report_processor.py` | `DEFAULT_SUFFIX_PATTERN`, `is_refreshed_copy_of` (`refresh_naming.py`) | `2aaa2b1` | yes (`2aaa2b1`) |
| `power_query_extractor/gui/extractor_window.py` | `is_refreshed_name` (`refresh_naming.py`) | `2aaa2b1` | yes (`2aaa2b1`) |

The 36 `.py` files that were never changed after creation (RAN (me)): `core/__init__.py`, `gui/__init__.py`, `gui/handlers/__init__.py`, `gui/tabs/__init__.py`, `gui/tabs/processing_tab.py`, `gui/tabs/validation_tab.py`, `gui/utils/__init__.py`, `gui/utils/cache_manager.py`, `gui/utils/status_bar.py`, `gui/utils/ui_helpers.py`, `gui/widgets/__init__.py`, `gui/widgets/collapsible_frame.py`, `gui/widgets/title_bar.py`, `launch_extractor.py`, the four `power_query_extractor` package markers, `power_query_extractor/core/data_consolidator.py` (all since `edac6a3`); `gui/utils/setup_validation.py`, `gui/widgets/tooltip.py` (since `322717f`); `power_query_extractor/core/refresh_naming.py` (since `2aaa2b1`); `utils/excel_com.py` (since `ba2a753`); `utils/dependency_installer.py` (since `a5ad487`); and 12 test files (`tests/test_dependency_installer.py`, `test_excel_com.py`, `test_extractor_standalone.py`, `test_file_patterns.py`, `test_main_window_tabs.py`, `test_optional_pywin32.py`, `test_processing_handler.py`, `test_refresh_naming.py`, `test_scan.py`, `test_scroll_bindings.py`, `test_setup_validation.py`, `test_workflow_mode.py`). Each imports only names that existed when it was created. A second check (RAN (me)) of the attributes the never-changed tab files use on the main window (`self.app.<name>`) found every one either defined in `gui/main_window.py@edac6a3` or set by the tab itself.

**Stable files whose meaning changed around them** (the rule's spirit — a file can be textually unchanged yet behave differently because what it calls changed; READ, HIGH):
- `launch_extractor.py` (never changed) starts `extractor_main.PowerQueryExtractor`, which since `a677670` builds a shell around the same panel used by Step 4.
- `power_query_extractor/extractor_main.py` imports `PowerQueryExtractorApp` — a name that existed since `edac6a3` but became a different kind of object (full window → thin shell) in `a677670`.
- `power_query_extractor/core/data_consolidator.py` (never changed) began putting another cell's value into nine columns when `47a1492` changed the configuration it reads (A-R88; RAN (me), 12C.2).
- `gui/utils/cache_manager.py` (never changed) became the hand-over between the two programs when `c1583d6` pointed the extractor at its file.
- `gui/__init__.py` (never changed) imports `main_window`, so after `bfcb24a` importing the `gui` package loads the whole extractor; between `bfcb24a` and `c463256` that made the whole program fail to start without pywin32 (RAN (H-A3)).
- `gui/tabs/__init__.py`, `gui/widgets/__init__.py` and `utils/__init__.py` were not updated when `ExtractTab`, `Tooltip`, `WORKFLOW_MODES` and three new `utils` modules were added — harmless today because nothing imports from those package roots (READ; A-X-DEAD §2.7).

#### 12B.8 Where the older documents disagree with the file history

| Document and place | What it says | What the history shows (RAN (me) `git log`) |
|---|---|---|
| `PROJECT_AUDIT_ANNUALSTATEMENT.md:1093-1115` (its own 12B, written 05-03-2026) | `main.py` last modified `5ea77c2`; `requirements.txt`, `core/file_organizer.py`, `core/excel_handler.py`, `utils/constants.py`, `gui/main_window.py`, all `gui/tabs/*`, `gui/utils/*`, `gui/widgets/*`, `data_consolidator.py` "never modified"; `core/file_parser.py`, `utils/helpers.py`, `gui/handlers/file_handler.py`, `gui/handlers/processing_handler.py` last modified `6237597`; `gui/handlers/client_handler.py` last modified `1c48722`; `extractor_window.py` and `report_processor.py` last changed `bd176bf`; `cell_mappings.py` last changed `47a1492` | True on its date; **no longer true** for most rows: `main.py` (`f363946`, `be6418c`, `a5ad487`), `requirements.txt` (`6c1e1c6`, `064c634`), `core/file_parser.py` (`a603c02`, `064c634`, `56aaad9`), `core/file_organizer.py` (`275d9a4`), `core/excel_handler.py` (`12d505f`, `064c634`, `ba2a753`), `utils/constants.py` (`5c8ce95`, `7347383`, `510f7b4`), `utils/helpers.py` (`a603c02`, `33d0baf`, `4219ffb`, `56aaad9`, `4f84ba9`), `gui/main_window.py` (4 commits), `gui/handlers/client_handler.py` (`33d0baf`), `gui/handlers/processing_handler.py` (`0890022`, `33d0baf`, `510f7b4`), `gui/tabs/setup_tab.py` (5), `gui/utils/dark_mode_manager.py` (2), `extractor_window.py` (8 more), `report_processor.py` (4 more); and the folders `gui/tabs/`, `gui/utils/`, `gui/widgets/` gained new files (`extract_tab.py`, `setup_validation.py`, `tooltip.py`). Its table has no row for `extractor_main.py`, which has also changed since (`c1583d6`). Still true: `launch_extractor.py`, `gui/handlers/file_handler.py` (last `6237597`), `validation_tab.py`, `processing_tab.py`, `cache_manager.py`, `status_bar.py`, `ui_helpers.py`, `collapsible_frame.py`, `title_bar.py`, `cell_mappings.py` (last `47a1492`), `data_consolidator.py`. [Corrected during verification — originally counted `extractor_main.py` as one of its rows and left out the four rows whose "Last Modified" commit is also out of date.] |
| same file, `:1099` | `requirements.txt` "Modified on unmerged branch to add PyQt5" | The PyQt5 commits it means (`c258047`, `a62d3cc`, `c21f788`, 07-02-2026) are project B's history, not a branch of A (RAN (H-A1) `git log 47a1492..B/main`). |
| same file, `:968`, `:976`, `:1065` | `edac6a3` dated "2025-10-29"; `cbfa4a5` in a 29-10-2025 phase; PyQt5 commits "2026-03-04" | `edac6a3` and `cbfa4a5` are 31-08-2025; the PyQt5 commits are 07-02-2026 (RAN (A-D notes) — never true). |
| `README.md:72-95` (folder tree, last changed `d674386` 29-10-2025) [line range corrected during verification — originally `:69-97`] | Lists `core/`, `gui/`, `utils/`, `power_query_extractor/` with the original files | No `tests/`, `docs/`, `pytest.ini`, `utils/excel_com.py`, `utils/requirements_check.py`, `utils/dependency_installer.py`, `gui/tabs/extract_tab.py`, `gui/utils/setup_validation.py`, `gui/widgets/tooltip.py`, `power_query_extractor/core/refresh_naming.py` — all added after the split. |
| `README.md:180-190` (output tree) | `GSTR-3B Exports (ClientName)/`, `Other ITC related files (ClientName)/`, `Sales related files (ClientName)/` | Folder names shortened to `GSTR3B`, `ITC`, `Sales` in `7347383` (`utils/constants.py:91-95`). |
| `tests/README.md:16-23` | The test files are the four listed | 18 test files at the snapshot (RAN (A-D notes): 288 tests collected). |
| `docs/superpowers/specs/2026-07-25-merge-extractor-design.md:108-109` | `_classify_version_reports` and `_latest_refresh_times` "are unit-tested" | Not true at `b276542`; `_latest_refresh_times` got tests in `2aaa2b1` (`tests/test_refresh_naming.py:99-121`); `_classify_version_reports` has no test that names it even at the snapshot (RAN (H-A2), grep; re-run RAN (me) `git grep` at `b276542`, `2aaa2b1`, `bf2f1dd`). It is exercised only indirectly, when `tests/test_main_window_tabs.py:69` calls the panel's `scan_folder()`, which calls it (`power_query_extractor/gui/extractor_window.py:651`). [Refined during verification — originally "has none".] |

### 12C. Logic and Rule Evolution

This sub-section follows every business rule and GST rule of the registry (A-R1 … A-R157) back through the history: when it first appeared, whether it was changed or removed, and what the change suggests (foundational from the start, discovered in real use, found in a code review, a speed change, or a side effect of another change).

> **GST rules are described, never judged.** A rule tagged **[GST RULE]** decides a tax figure, which return / table / sheet / cell a figure belongs to, a rate, limit, rounding, period, GSTIN/PAN/state-code logic, ITC or a due date. For those this section says exactly what the code did at each commit and when that changed. It does not say whether any version is correct, does not recommend a version, and does not look up the law. Values that can change over time (table numbers, template cell positions, return and file-naming conventions, state lists, package versions, Excel limits) carry the symbol ⚠ so the author can check them.

#### 12C.1 How the evolution was established

- **Blame of every rule's code** (RAN (me), `rule_blame.py` → `rule_blame_final.txt`): for each registry rule, every `file:line` the registry cites was looked up with `git blame -w -M -C` (blame that ignores whitespace-only edits and follows lines that were moved or copied between files) and the commits that last touched those lines were counted. Summary in 12C.2; per rule in 12C.4.
- **Pickaxe searches** (RAN (me)) for the key literal of each changed rule — for example `1024`, `PK\x03\x04`, `"Invalid or corrupted Excel file"`, `replace('Private'`, `\bPrivate\b`, `[\s_]+', ' '`, `key=lambda x: x.name`, `_Refreshed_*.xlsx`, `WAIT 7 SECONDS`, `self.wait_time`, `local_temp`, `essential=False`, `essential=ON_WINDOWS`, `is_processing`, `get_state_code(values[1])`, `Not selected for processing`, `validatecommand`, `Names >10 chars` — to find the commit that added and the commit that removed each one.
- **Old and new code run side by side** (RAN (me), made-up inputs only): `rules_by_commit.py` (IMS Reco pattern, name cleaning, client folder key and category folder names at `edac6a3`, `47a1492`, `33d0baf`, `4219ffb`, `bf2f1dd`), `refresh_by_commit.py` (which file names count as refreshed copies at the split point and at the snapshot), `consol_by_commit.py` (the consolidated workbook's column lookup with the cell configuration of `edac6a3`, `47a1492` and `bf2f1dd`).
- **Every diff that changes a rule was read** (READ). Blame only shows the *last* commit that touched a line, so a cosmetic edit (for example `064c634` turning `logger.info("[DEBUG] …")` into `logger.debug(…)` inside the file-name matcher) makes a line look newer than the rule. Each blame hit was therefore checked against the diff; cosmetic touches are marked "cosmetic" below.

#### 12C.2 The overall picture: which rules were foundational and which were discovered

**Numbers** (RAN (me), `rule_stats.py` over the blame results):

| Group | Count | Meaning |
|---|---|---|
| All cited lines last changed in `edac6a3` | 74 | Rule present in the first commit and never edited since. |
| Some lines from `edac6a3`, some later | 67 | Present in the first commit; later commits edited some of its lines (substantively or cosmetically — see 12C.4). |
| No cited line from `edac6a3` | 14 | A-R9, A-R68, A-R78, A-R79, A-R106, A-R113, A-R124, A-R125, A-R129, A-R133, A-R144, A-R145, A-R146, A-R147. Of these, **A-R9** (1,024-byte minimum) and **A-R106** (openpyxl load options, `keep_links=True` at `core/excel_handler.py:242@edac6a3`) existed in `edac6a3` in lines that were later rewritten (pickaxe `git log -S'1024' -- utils/helpers.py` lists only `edac6a3`: the number was added there and its count of occurrences never changed afterwards, although the lines holding it were rewritten; re-run during verification), and **A-R146** (mouse wheel) began in `edac6a3` (`gui/tabs/setup_tab.py:66@edac6a3`); **A-R68** (recognising a refreshed file) also began in `edac6a3` in a hard-coded form — the Skip Refresh lookup `glob(f"{original_file.stem}_Refreshed_*.xlsx")` (`power_query_extractor/core/report_processor.py:80@edac6a3`) — and all its lines were rewritten by `2aaa2b1` (as its row in 12C.4 says). The other 10 (A-R78, A-R79, A-R113, A-R124, A-R125, A-R129, A-R133, A-R144, A-R145, A-R147) are genuinely new after the first commit. [Corrected during verification — originally "the other 11", which counted A-R68 as new and contradicted 12C.4.] |
| No code lines (rules about something that does not exist) | 2 | A-R24 (no GSTIN or PAN handling) and A-R123 (no run history). RAN (me): a case-insensitive search for "gstin" and the word "pan" in the non-test `.py` files of all 46 commits found nothing, so A-R24 held at every commit. |

**What this means for a rebuild** (inference, HIGH): the *business and GST rules* of project A — the six file-name patterns, the expected six file types, the completeness and duplicate rules, the state table and its fallback, the folder layout, the Links-sheet cell positions, the refresh success test (`Info!BB2`) and the original eight extraction cells — were all in the first commit and have essentially not changed since. They are **foundational**: they were settled during the 2½ months of pre-git work, whose only traces are tuning comments such as "Increased from 0.5" (`power_query_extractor/core/report_processor.py:283, 288`) and "WAIT 7 SECONDS AS REQUESTED" above a 10-second loop (`report_processor.py@edac6a3`, removed in `bd176bf`). What changed after the first commit is overwhelmingly **operational**: speed on a network share, Excel's path-length limit, refresh waiting and naming, how settings and packages are found, the window, and dark mode.

**The only changes to [GST RULE]-tagged rules in the whole history** (READ of every diff; HIGH):

| Commit | Label | Rule | What changed in the code (no judgement) |
|---|---|---|---|
| `47a1492` 30-10-2025 | shared - before the split on 30-10-2025 | A-R77 – A-R82 added ⚠ | Six new extraction output sheets, all from the ITC report: "Reco Diff" (`Reco` C37/D37/E37), "Table 6J Match" (`T6 of R9` K29), "Table 8D Limits" (`T8 of R9` G6), "IMS Reco" (`IMS_Reco` AA1/AD1/AG1), "Extra In Purchase" (`ExtraInPurchase` O1/P1/Q1), "Extra In GSTR2B" (`ExtraInGSTR2B` S1/T1/U1) (`power_query_extractor/config/cell_mappings.py:105-228`). The names "Table 6J" and "Table 8D" and the sheet names "T6 of R9", "T8 of R9" refer, by their wording, to tables of an annual return; only where the code reads is described here. |
| `47a1492` | shared - before the split on 30-10-2025 | A-R88 (effect, no code change) | The consolidator's column lookup has been the same since `edac6a3` (`power_query_extractor/core/data_consolidator.py:141-150`). With the `edac6a3` configuration it put every value in its configured column; with the columns added in `47a1492` it puts the values of `Diff GSTR-3B!J15/K15/L15` into nine columns. RAN (me), `consol_by_commit.py`, every extracted value set to its own label: **`edac6a3`: 8 columns, 0 show another cell; `47a1492`: 22 columns, 9 show another cell; `bf2f1dd`: 22 columns, the same 9** (re-run RAN (VERIFY-A-S12BF) on fresh `git archive` copies of the three commits: identical). The nine are the IGST/CGST/SGST columns of "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" ⚠ (template cell positions). A technical mismatch between configuration and code; which cell is right is the author's call. (Gap A-G1 in Section 7.) |
| `47a1492` | shared - before the split on 30-10-2025 | old-format mapping removed | The never-importable `cell_mappings copy.py` (with four commented-out candidate cells) deleted — see UNREGISTERED in 12C.5. |
| `bd176bf` 30-10-2025 | shared - before the split on 30-10-2025 | A-R89 | With per-report ticks, a report the user did not tick gets the status "Not selected for processing"; like any unsuccessful status, every column of that report then shows the "Missing" text in the consolidated workbook (`power_query_extractor/core/data_consolidator.py:128-129`; READ diff). |
| `33d0baf`, `510f7b4` (and `0890022`, logging only) | A only (after the split) | A-R60 | The per-client order (folders → copies → ITC report → Sales report → text report) and the way success is counted are unchanged since `edac6a3`; `33d0baf` makes Start ignore a second press during a run (`gui/handlers/processing_handler.py:75-76`); `510f7b4` appends the Step 4 hand-over after the summary (`:519`). |
| `322717f` 25-07-2026 | A only (after the split) | A-R1 – A-R6 (display only) ⚠ | The Setup card shows the patterns' `description` texts from `utils/constants.py` instead of a hand-typed list (`gui/tabs/setup_tab.py:362-369@edac6a3` → `gui/tabs/setup_tab.py:477-503`); the descriptions do not mention `.xls` or the optional "(N)" markers that the patterns accept (READ (H-A2)). The patterns themselves are unchanged. |
| `5c8ce95` 23-07-2026 | A only (after the split) | A-R2 ⚠ | IMS Reco file name `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$` → `^ImsReco-([^-]+)-([^-]+)-(\d{8})(?:\s*\(\d+\))?\.xlsx?$` (`utils/constants.py:23`): an optional "(N)" or " (N)" after the 8-digit date is now accepted; the captured date is unchanged. |
| `33d0baf` 23-07-2026 | A only (after the split) | A-R17, A-R23 | The per-client "include client name" toggle and the global tick now store their flag under the client key with the state CODE (`gui/handlers/client_handler.py:181, 218`), like the scan and the organiser; before, they used the state text as shown, so the flags were never found (READ diff). The key format itself (`<client>-<state code>`) is unchanged since `edac6a3`. |
| `7347383` 26-07-2026 | A only (after the split) | A-R54, A-R55 (values, not positions) ⚠ | The Links-sheet cell positions are unchanged, but the folder *text* written into ITC B2/B7/B10 and Sales B2/B8 changed because the category folders were renamed (`…/Other ITC related files` → `…/ITC`, etc., A-R39). Whether any template relies on the old folder names is a question for the author (12C.7). |
| `275d9a4` 23-07-2026 | A only (after the split) | A-R59 (recognition side) | Step 4 now recognises reports by a lower-case `itc_report_` / `sales_report_` prefix over one directory listing (`power_query_extractor/gui/extractor_window.py:20-42`) instead of the file search `ITC_Report_*.xlsx` (any letter case now; on Windows the old search was already case-insensitive). |
| `5c8ce95` (side effect) | A only (after the split) | A-R57 | No code change to A-R57, but a new case: when a folder holds both `…-01042025.xlsx` and `…-01042025 (1).xlsx` for one client, both are now IMS Reco files and, because the scan sorts by path and " (" sorts before ".", the "(1)" copy is the one linked into the ITC report (RAN (H-A1)). |

No [GST RULE] was ever removed from the live code; the state table (A-R21) ⚠, the state fallback (A-R22), the six expected types (A-R26) ⚠, the GSTR-3B "any number of files" exception (A-R27) ⚠, the completeness rules (A-R28, A-R29), the Links cell positions (A-R54, A-R55) ⚠, the `Info!BB2` test (A-R73) ⚠, the original eight extraction cells (A-R74 – A-R76) ⚠ and the number display (A-R90) are textually identical in `edac6a3` and `bf2f1dd` (RAN (me) blame: all their cited lines are `edac6a3`, apart from the cosmetic and import-line touches listed in 12C.4).

#### 12C.3 Every rule change in date order

| Date | Commit | Label | Rules | Old → new | Kind |
|---|---|---|---|---|---|
| 31-08-2025 | `edac6a3` | shared - before the split on 30-10-2025 | all foundational rules | introduced | foundational (pre-git) |
| 29-10-2025 | `6237597` | shared - before the split on 30-10-2025 | A-R136 | Start warning "client names longer than **10** characters" → "longer than the **max client folder name length** setting (default 35; 10 if the value is 0 or less or unreadable)" (`gui/handlers/processing_handler.py:86-113`); the per-client toggle keeps 10 (`gui/handlers/client_handler.py:191`) | review |
| 29-10-2025 | `6237597` | shared - before the split on 30-10-2025 | A-R139, A-R140, A-R141 | screen updates wrapped in `root.after(0, lambda: …)` — introduced wrong names in the Dry Run and possibly in the Step 3 log (12D) | review (with a regression) |
| 29-10-2025 | `1c48722` | shared - before the split on 30-10-2025 | A-R138 | row colours: before, every row got one colour (light mode all "complete", dark mode all "incomplete", RAN (H-S)) → complete green / other orange (`gui/handlers/client_handler.py:52-55`) | real use |
| 29-10-2025 | `4c42758` | shared - before the split on 30-10-2025 | A-R146 | extractor client list and log scroll with the wheel | real use |
| 30-10-2025 | `bd176bf` | shared - before the split on 30-10-2025 | A-R65, A-R67, A-R69, A-R72, A-R89 [GST RULE], A-R119, A-R143, A-R145 | both reports always → per-report ticks; fixed suffix → user pattern; fixed 10 s waits → N s (5–60, default 10) and max(10, N); "Not selected for processing" → every column "Missing"; "Last Refresh Status" column; options never stored | real use |
| 30-10-2025 | `47a1492` | shared - before the split on 30-10-2025 | A-R77 – A-R82 [GST RULE] ⚠; A-R88 (effect) | see 12C.2 | real use (template grew) |
| 23-07-2026 | `a603c02` | A only (after the split) | A-R8, A-R11, A-R109 | three file searches → one listing (upper-case extensions now found off Windows, RAN (H-A1)); "Invalid or corrupted Excel file" variation removed | speed |
| 23-07-2026 | `12d505f` | A only (after the split) | A-R98 (new), A-R97, A-R58, A-R106 | report built in place on the network target → built in `%TEMP%\gst_excel_temp\` and copied once | speed |
| 23-07-2026 | `275d9a4` | A only (after the split) | A-R111 (new), A-R59 [GST RULE], A-R62, A-R145 | Level-1 folder once per run; Step 4 report presence by one listing; "Folders Created" count changes in fresh mode | speed |
| 23-07-2026 | `5c8ce95` | A only (after the split) | A-R2 [GST RULE] ⚠ | see 12C.2 | real use |
| 23-07-2026 | `c1583d6` | A only (after the split) | A-R42, A-R114, A-R118 | newest Version folder by NAME → by date modified; extractor reads the organiser's home-folder settings file first | review |
| 23-07-2026 | `0890022` | A only (after the split) | A-R139, A-R140, A-R141 | values captured when scheduled (`gui/handlers/processing_handler.py:23-36`) | review |
| 23-07-2026 | `6c1e1c6` | A only (after the split) | A-R22 [GST RULE], A-R126 ⚠ (and pins A-R1 – A-R6 [GST RULE] ⚠, A-R19, A-R21 [GST RULE] ⚠) | tests pin current behaviour incl. "Unknown Place" → "UP" (`tests/test_helpers.py:96-97`); pytest made a listed requirement | tests |
| 23-07-2026 | `33d0baf` | A only (after the split) | A-R17 [GST RULE], A-R19, A-R20, A-R23 [GST RULE], A-R40, A-R135, A-R137 | per-client key by state code; "Private"/"Limited" in file names: case-sensitive substring → whole word, any case; Start ignored while a run is active (`processing_handler.py:75-76`) | review |
| 23-07-2026 | `064c634` | A only (after the split) | A-R122, A-R126 | 35 "[DEBUG]" INFO lines → DEBUG; pywin32 added to requirements | review |
| 23-07-2026 | `ba2a753` | A only (after the split) | A-R100 | two identical loops → one shared helper (behaviour unchanged) | refactor |
| 25-07-2026 | `4219ffb` | A only (after the split) | A-R19, A-R20 | separators normalised before the abbreviation (`utils/helpers.py:91`) — undoes `33d0baf`'s side effect on `X_Private_Limited` | regression fix |
| 25-07-2026 | `56aaad9`, `4f84ba9` | A only (after the split) | A-R8, A-R9, A-R10, A-R57 [GST RULE], A-R109, A-R110 (new) | cached file details; signature reads in parallel (16 at once, from 4 files) ⚠ (tuning numbers) | speed (measured on the real share) |
| 25-07-2026 | `322717f` | A only (after the split) | A-R36, A-R52, A-R113 (new), A-R133 (new), A-R147 (new), A-R146 | live checklist; typed length no longer validated; template list for the checklist | review / usability |
| 25-07-2026 | `d64d3a9` | A only (after the split) | A-R146 | one filtered wheel handler on Step 1 | regression fix |
| 25-07-2026 | `c1f425e`, `2acffed`, `a677670` | A only (after the split) | A-R143, A-R145 | worker updates marshalled; refresh times recomputed after a run; panel + shell | real use (2acffed reported) / merge plan |
| 25-07-2026 | `2aaa2b1` | A only (after the split) | A-R68 (new form), A-R69 | refreshed copies recognised by the configured suffix plus the legacy `_refreshed_` | real use (reported) |
| 26-07-2026 | `7347383` | A only (after the split) | A-R39, A-R104; A-R54/A-R55 values [GST RULE] ⚠ | category folders shortened | real use (measured) |
| 26-07-2026 | `bfcb24a` | A only (after the split) | A-R121, A-R131, A-R143, A-R144 (new) | fourth tab; Step 4 follows Step 1's target; no scan at start-up | merge plan |
| 26-07-2026 | `510f7b4` | A only (after the split) | A-R115, A-R116, A-R132, A-R144, A-R154 (new) | workflow mode `organize` / `full` | merge plan |
| 26-07-2026 | `aeb83cf`, `c89c1d4`, `725fbaf`, `169c7b5` | A only (after the split) | A-R148, A-R113 | dark mode capture/replay; checklist re-check on theme switch; empty style values recorded; dark checklist colours | merge plan / reported |
| 26-07-2026 | `c463256` | A only (after the split) | A-R64, A-R129 (new) | pywin32 imports guarded; per-client refusal; Step 4 notice | regression fix |
| 26-07-2026 | `f363946`, `be6418c`, `a5ad487` | A only (after the split) | A-R124 (new), A-R125 (new), A-R127 | start-up check (pywin32 optional) → essential on Windows → install offer | real use (crash on the author's PC) |
| 27-07-2026 | `bf2f1dd` | A only (after the split) | A-R146, A-R147, A-R154 | extractor wheel handler filtered; hover help | real use (reported) |

**RAN (me) — naming rules at five commits** (`rules_by_commit.py`, made-up names):

| Input | `edac6a3` / `47a1492` | `33d0baf` | `4219ffb` / `bf2f1dd` |
|---|---|---|---|
| IMS name `ImsReco-ClientA-Delhi-01042025 (1).xlsx` | not recognised | recognised (pattern from `5c8ce95`) | recognised |
| IMS name `(1) ImsReco-ClientA-Delhi-01042025.xlsx` | not recognised | not recognised | not recognised |
| `sanitize_filename('ClientA_Private_Limited')` | `ClientA_Pvt_Ltd` | `ClientA_Private_Limited` | `ClientA_Pvt_Ltd` |
| `sanitize_filename('CLIENTA PRIVATE LIMITED')` | `CLIENTA_PRIVATE_LIMITED` | `CLIENTA_Pvt_Ltd` | `CLIENTA_Pvt_Ltd` |
| `sanitize_filename('Privateer Traders')` | `Pvter_Traders` | `Privateer_Traders` | `Privateer_Traders` |
| `sanitize_filename('ClientAPrivateLimited')` | `ClientAPvtLtd` | `ClientAPrivateLimited` | `ClientAPrivateLimited` |
| Level-2 folder key `create_client_state_key('ClientA_Private_Limited', 'Delhi')` | `ClientA_Private_Limited-DL` | same | same (the separator step was never added here, `utils/helpers.py:620-644`) |
| Category folder templates | `GSTR-3B Exports ({client})`, `Other ITC related files ({client})`, `Sales related files ({client})` | same | `4219ffb`: same; `bf2f1dd`: `GSTR3B ({client})`, `ITC ({client})`, `Sales ({client})` |

**RAN (me) — refreshed-copy recognition** (`refresh_by_commit.py`; original report `ITC_Report_ClientA_Pvt_Ltd_DL_211026_1000.xlsx`, made up):

| "Refreshed File Suffix" | Split point `47a1492` (hard-coded `_Refreshed_` searches at `extractor_window.py:645@47a1492` and `report_processor.py:103@47a1492`) | Snapshot (`refresh_naming.py:43-66`) — copy recognised? | Snapshot — the ORIGINAL itself counted as a refreshed copy? |
|---|---|---|---|
| `_Refreshed_{timestamp}` (default) | found by status column and Skip Refresh | yes | no |
| `_Updated_{timestamp}` | not found by either | yes | no |
| `_R{timestamp}` | not found | yes | **yes** |
| `_{timestamp}` | not found | yes | **yes** |
| `_DL{timestamp}` | not found | yes | **yes** |

So `2aaa2b1` fixed the reported case (a custom suffix was invisible) and introduced over-matching for short suffixes (gap A-G16 in Section 7).

#### 12C.4 Rule-by-rule register (all 157 registry rules)

Columns: **GST** = tagged [GST RULE] in the registry; **Introduced** = first commit (all `edac6a3` entries are shared - before the split on 30-10-2025; all entries from `fc279af` on are A only (after the split)); **Later changes** = every substantive change, plus cosmetic touches found by blame; **Kind** = F foundational (first commit, settled before git), U discovered in real use (the commit message reports a real failure or measurement), R found in a code review (often AI-assisted), P speed change, M part of the 2026 merge plan, S side effect of another change. Evidence: blame and pickaxe RAN (me) for every row; diffs READ.

**Category 1 — Input-file recognition**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R1 | GSTR-2B Reco name pattern ⚠ | yes | `edac6a3` | Regex never changed (`utils/constants.py:15`). Shown on the Setup "EXPECTED FILE NAMES" card from `322717f` (before: a hand-typed list, `gui/tabs/setup_tab.py:362-369@edac6a3`). `064c634` cosmetic (two log lines of the matcher). | F |
| A-R2 | IMS Reco name pattern ⚠ | yes | `edac6a3` | `5c8ce95`: optional "(N)" after the date (see 12C.2). | U |
| A-R3 | GSTR-3B Export name pattern ⚠ | yes | `edac6a3`, already with the optional leading "(N) " for duplicate downloads | none | F |
| A-R4 | Sales name pattern ⚠ | yes | `edac6a3` | none | F |
| A-R5 | Sales Reco name pattern ⚠ | yes | `edac6a3` | none | F |
| A-R6 | Annual Report name pattern ⚠ | yes | `edac6a3` | none | F |
| A-R7 | First matching pattern wins; parts trimmed | yes | `edac6a3` | none to the matcher (`core/file_parser.py:82-128`): `6237597` changed the same file only above it (new `_update_progress`, `:48-54`) and below it (blank lines removed in `scan_folder`), RAN (me) `git show 6237597 -- core/file_parser.py`; `064c634` cosmetic (2 log lines). [Corrected during verification — originally "`6237597` removed blank lines nearby".] | F |
| A-R8 | Which files are looked at | no | `edac6a3`: three searches `*.xlsx`, `*.xls`, `*.xlsm` (letter case matters off Windows), then validation | `a603c02` one listing with a lower-case test; `56aaad9` cached details, regular-file test from the cached details; `4f84ba9` validation split in two. Same files on Windows (RAN (H-A1, H-A2) equivalence runs). | P |
| A-R9 | Minimum size 1,024 bytes | no | `edac6a3` (pickaxe) | moved into `_passes_metadata_checks` (`utils/helpers.py:303`) by `4f84ba9`; value unchanged | F |
| A-R10 | Excel signature bytes | no | `edac6a3` | moved into `_has_excel_signature` (`utils/helpers.py:275-287`) and read in parallel by `4f84ba9`; bytes unchanged | F, P |
| A-R11 | Unrecognised files become variations | no | `edac6a3` (three kinds) | `a603c02` removed the third kind, "Invalid or corrupted Excel file", which could not occur | F |
| A-R12 | Suggested pattern for an unrecognised name ⚠ | yes | `edac6a3` | none | F |
| A-R13 | `EXCEL_EXTENSIONS` (dead) | no | `edac6a3` | none; the number of mentions of the name is the same in every commit (pickaxe) — never used | F (dead from day one) |
| A-R14 | `MAX_FILE_SIZE_MB = 100` (dead) | no | `edac6a3` | none (never used) | F (dead) |
| A-R15 | `EXCEL_SIGNATURES` (dead) | no | `edac6a3` | none (never used) | F (dead) |
| A-R16 | Pattern `category`/`folder` labels (dead) | yes | `edac6a3` | none | F (dead) |

**Category 2 — Client identity and name cleaning**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R17 | Client key `<client>-<state code>` | yes | `edac6a3` (scan, organiser and `get_selected_clients` all used the code) | `33d0baf`: the folder-name toggle and the global tick now build the same key (before: `<client>-<state text>`); `6237597` cosmetic (the import line inside `get_state_code`, `utils/helpers.py:598`, found by blame during verification) | R |
| A-R18 | `validate_client_name` (output never read) | no | `edac6a3` | none | F (dead) |
| A-R19 | "Private" → "Pvt", "Limited" → "Ltd" | no | `edac6a3`: Level-2 key whole word any case (`utils/helpers.py:624-625`, unchanged); file names case-sensitive substring | `33d0baf` file names whole word any case; `4219ffb` spaces/underscores normalised first (RAN (me) table in 12C.3) | R, then regression fix |
| A-R20 | `sanitize_filename` steps | no | `edac6a3` | abbreviation part as A-R19 (`utils/helpers.py:91-93`); everything else unchanged | F + R |

**Category 3 — State, GSTIN and PAN codes**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R21 | State name → two-letter code table (39 names) ⚠ | yes | `edac6a3` | none to the table (`utils/constants.py:251-292`); `6237597` changed only the import line inside `get_state_code` (`utils/helpers.py:598`), no behaviour change (RAN (H-S)) | F |
| A-R22 | Fallback code for unknown names | yes | `edac6a3` | none to the rule; `6237597` cosmetic (the import line `utils/helpers.py:598`, found by blame during verification); pinned by a test since `6c1e1c6` (`tests/test_helpers.py:96-97`) | F |
| A-R23 | Where the code and where the text are used | yes | `edac6a3` | `33d0baf` toggle keys now use the code | F + R |
| A-R24 | No GSTIN or PAN handling | yes | absent at every commit (RAN (me)) | — | F |

**Category 4 — Financial year and tax period**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R25 | Period parts captured into `metadata`, never read ⚠ | yes | `edac6a3` | none to the reading side; `5c8ce95` kept the IMS date group; for the other five patterns a trailing "(N)" lands inside the captured period text (RAN (H-A1)) | F |

**Category 5 — Expected files and completeness**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R26 | Six expected file types ⚠ | yes | `edac6a3` | none | F |
| A-R27 | Duplicates; GSTR-3B Export exempt ("monthly filings - unlimited") ⚠ | yes | `edac6a3` | none | F |
| A-R28 | Completeness percentage | yes | `edac6a3` | none | F |
| A-R29 | Client status Complete / Missing N / Has duplicates | yes | `edac6a3` | none | F |
| A-R30 | "Critical" types in the never-called `validate_results` ⚠ | yes | `edac6a3` | none (never called at any commit) | F (dead) |
| A-R31 | Scan statistics and Step 2 summary text | no | `edac6a3` | none | F |

**Category 6 — Folder layout and naming**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R32 | Level-1 folder `Annual Statement-{timestamp}` | no | `edac6a3` | none | F |
| A-R33 | Run time stamps `%d%m%y %H%M` | no | `edac6a3` | none | F |
| A-R34 | `get_date_only` (dead) | no | `edac6a3` | none | F (dead) |
| A-R35 | Level-2 folder name via `create_client_state_key` | yes | `edac6a3` | none — it never received `4219ffb`'s separator step (RAN (me)) | F |
| A-R36 | Max client folder name length values | no | `edac6a3`: default 35, arrows 15–100, typed values checked 10–100, label "(default: 35, range: 15-100)" (`gui/tabs/setup_tab.py:252-279@edac6a3`) | `6237597` the value also drives the Start warning (A-R136); `322717f` validator removed (typed text now unchecked), hint "(default 35 · names over {limit} are shortened)"; `bf2f1dd` hover text | R |
| A-R37 | `FOLDER_STRUCTURE['level2']` (dead) | no | `edac6a3` | none | F (dead) |
| A-R38 | New `Version-{timestamp}` folder every run | no | `edac6a3` | none | F |
| A-R39 | Level-4 category folder names | no | `edac6a3`: `GSTR-3B Exports`, `Other ITC related files`, `Sales related files` (+ " (<client>)") | `7347383`: `GSTR3B`, `ITC`, `Sales` (`utils/constants.py:91-95`); the commit message reports 1,094 of 3,030 real output files over Excel's ~218-character limit ⚠ | U (measured) |
| A-R40 | Include-client-name decision | no | `edac6a3` (lookup side unchanged) | per-client flags work only since `33d0baf` fixed the writers' key | R |
| A-R41 | Which folder each file type goes into ⚠ | yes | `edac6a3` | none (the folder names themselves changed with A-R39) | F |
| A-R42 | Step 4: which run folder and which Version folder | no | `edac6a3`: run folder newest by date modified; Version folder highest NAME | `c1583d6` Version folder newest by date modified (RAN (me): by name `Version-300626 1000` beats `Version-150726 0900`); `275d9a4` report presence by one listing | R |

**Category 7 — Copying and versioning of files**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R43 | Processing modes fresh / rerun / resume | no | `edac6a3` | constants unchanged; `275d9a4` Level-1 choice once per run; help texts `322717f`, `bf2f1dd` (the Resume help text promises skipping that only happens within the same minute, gap A-G10) | F |
| A-R44 | Mode icons (dead) | no | `edac6a3` | none | F (dead) |
| A-R45 | Existing destination file by mode | no | `edac6a3` | none | F |
| A-R46 | Backup name `<stem>_backup_<DDMMYY_HHMM>` | no | `edac6a3` | none | F |
| A-R47 | Copy method (copy2 + size check) | no | `edac6a3` | none | F |
| A-R48 | Names of copies; in-memory rename | no | `edac6a3` | none | F |
| A-R49 | `get_unique_path` (dead) | no | `edac6a3` | none | F (dead) |
| A-R50 | `calculate_file_hash` (dead) | no | `edac6a3` | none | F (dead) |

**Category 8 — Report creation and template Links cells**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R51 | Processing needs both templates and a target | yes | `edac6a3` | none (the `322717f` checklist informs but does not gate) | F |
| A-R52 | Accepted template file types | no | `edac6a3`: dialogs `*.xltx *.xlsx`; report always `.xlsx` | `322717f` checklist list `.xlsx .xltx .xltm .xlsm` (`gui/utils/setup_validation.py:21`) — a third list | R |
| A-R53 | `TEMPLATE_EXTENSIONS` (dead) | no | `edac6a3` | none | F (dead) |
| A-R54 | ITC Links cells ⚠ | yes | `edac6a3` | positions unchanged; folder text in B2/B7/B10 changed with `7347383` | F (+ S) |
| A-R55 | Sales Links cells ⚠ | yes | `edac6a3` | positions unchanged; folder text in B2/B8 changed with `7347383`; `064c634` cosmetic (14 Sales debug lines) | F (+ S) |
| A-R56 | Form of the Links values | yes | `edac6a3` | none | F |
| A-R57 | First file of a type in sorted order feeds the report | yes | `edac6a3` | sorting kept through `a603c02`/`56aaad9`/`4f84ba9` (`utils/helpers.py:361`); new "(1)" case after `5c8ce95` | F (+ S) |
| A-R58 | Finding the Links sheet | no | `edac6a3` | `12d505f` now on the local copy (no rule change) | F |
| A-R59 | Report file names; Step 4 recognition | yes | `edac6a3` | `275d9a4` recognition by lower-case prefix; `2aaa2b1` one line | F + P |
| A-R60 | Step 3 order per client; what counts as success | yes | `edac6a3` | `0890022` logging only; `33d0baf` start guard; `510f7b4` hand-over appended at the end (`processing_handler.py:519`) | F |
| A-R61 | Step 3 summary workbook | no | `edac6a3` | none | F |
| A-R62 | Organisation text report | no | `edac6a3` | `275d9a4` changed the "Folders Created" number in fresh mode (S) | F |
| A-R63 | `get_relative_excel_path` (dead) | no | `edac6a3` | none | F (dead) |

**Category 9 — Power Query refresh**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R64 | Step 4 needs pywin32; no fallback | no | `edac6a3`: unconditional imports, the extractor could not even load without pywin32 | `c463256`: guarded imports; per-client refusal message (`report_processor.py:20-36, 62-77`) | U (regression from `bfcb24a`) |
| A-R65 | Which reports are processed | no | `edac6a3`: every ticked client, both reports | `bd176bf`: per-report ticks, "Not selected for processing" | U |
| A-R66 | The "original" = first search hit | no | `edac6a3` | none | F |
| A-R67 | Refreshed-copy name | no | `edac6a3` fixed `<stem>_Refreshed_<DDMMYY_HHMM>.xlsx` | `bd176bf` user suffix; `2aaa2b1` unused import of the default | U |
| A-R68 | Recognising a refreshed file | no | `edac6a3` skip lookup `<stem>_Refreshed_*.xlsx`; `bd176bf` status search `*_Refreshed_*`; `275d9a4` `'_refreshed_'` substring | `2aaa2b1` configured markers OR legacy marker (RAN (me) table in 12C.3) | U (reported twice) |
| A-R69 | Skip refresh | no | `edac6a3` | `bd176bf` tick moved; `2aaa2b1` configured-suffix matching | U |
| A-R70 | `auto_refresh_if_missing` (dead) | no | `edac6a3` | none | F (dead) |
| A-R71 | Refresh by key presses Alt+A, R, A ⚠ | no | `edac6a3` | `bd176bf` cosmetic (log text) | F |
| A-R72 | Refresh waits and retries | no | `edac6a3`: fixed 10 s (comment "WAIT 7 SECONDS AS REQUESTED"), 5 × 1 s checks, extended 10 s; tuning comments "Increased from 0.5", "Increased from 3 to 5" | `bd176bf`: N s (box 5–60, default 10); extended max(10, N) (`report_processor.py:297-330`) | U |
| A-R73 | Refresh success test `Info!BB2` ⚠ | yes | `edac6a3` | none | F |

**Category 10 — Extraction cell mappings** (all [GST RULE] ⚠ except where marked)

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R74 | "GSTR-3B Difference" ← `Diff GSTR-3B` J15/K15/L15 ⚠ | yes | before git (also in the 21-06-2025 copy file); in `edac6a3` | none | F |
| A-R75 | "GSTR-1 Difference" ← `Diff - GSTR-1` N1/O1/P1/Q1 ⚠ | yes | before git (also in the copy file); `edac6a3` | none | F |
| A-R76 | "Turnover GSTR-1" ← `Diff - GSTR-1` I1 ⚠ | yes | `edac6a3` (not in the older copy file) | none | F |
| A-R77 | "Reco Diff" ← `Reco` C37/D37/E37 ⚠ | yes | `47a1492` | none | U |
| A-R78 | "Table 6J Match" ← `T6 of R9` K29 ⚠ | yes | `47a1492` | none | U |
| A-R79 | "Table 8D Limits" ← `T8 of R9` G6 ⚠ | yes | `47a1492` | none | U |
| A-R80 | "IMS Reco" ← `IMS_Reco` AA1/AD1/AG1 ⚠ | yes | `47a1492` | none (shows another cell since day one, A-R88) | U |
| A-R81 | "Extra In Purchase" ← `ExtraInPurchase` O1/P1/Q1 ⚠ | yes | `47a1492` | none (same) | U |
| A-R82 | "Extra In GSTR2B" ← `ExtraInGSTR2B` S1/T1/U1 ⚠ | yes | `47a1492` | none (same) | U |
| A-R83 | How mappings are read (field names, exact sheet names) | yes | `edac6a3` | none | F |
| A-R84 | `FIELD_DISPLAY_NAMES` (dead) | no | `edac6a3` | none | F (dead) |
| A-R85 | Calculated columns (dormant) | yes | `edac6a3` | none (no mapping ever used a calculation) | F (dormant) |

**Category 11 — Consolidated output layout**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R86 | `PQ_Extraction_Report_<%Y%m%d_%H%M>.xlsx` in the last-scanned run folder | no | `edac6a3` | none | F |
| A-R87 | Sheet layout | no | `edac6a3` | none | F |
| A-R88 | Which extracted value goes into a column | yes | `edac6a3` | no code change; effect changed by `47a1492` (RAN (me), 12C.2) | F (+ S) |
| A-R89 | "Missing" when the refresh did not succeed | yes | `edac6a3` | `bd176bf`: unticked reports also give "Missing" in every column | F |
| A-R90 | Number display `#,##0.00` for non-zero numbers | yes | `edac6a3` | none | F |
| A-R91 | Consolidated Summary sheet | no | `edac6a3` | none | F |
| A-R92 | Description lines (dormant) and a doubled header write | no | `edac6a3` (with "# ... rest of the code" paste markers) | none | F (dormant) |
| A-R93 | `show_timestamp` (dead) | no | `edac6a3` | none | F (dead) |
| A-R94 | `_write_section` older layout (never called) | no | before git (older consolidated layout); in `edac6a3` | none | F (dead) |
| A-R95 | `_create_details_sheet` (never called) | no | before git; `edac6a3` | none | F (dead) |

**Category 12 — Excel handling and safety**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R96 | COM when Windows + pywin32, else openpyxl | no | `edac6a3` | function unchanged; `be6418c` added a start-up gate that refuses on Windows without pywin32 (A-R124) | F |
| A-R97 | Step 3 Excel session (attach, hide, always quit) | no | `edac6a3` | `12d505f` opens the local copy; the attach/quit pattern never changed (`core/excel_handler.py:123-128, 238-246`) | F |
| A-R98 | Local staging of Step 3 reports | no | `12d505f` (before: template copied straight to the network target and edited there) | none | P |
| A-R99 | Report counts as created if the file exists | no | `edac6a3` | none | F |
| A-R100 | "Already open" detection | no | `edac6a3` (two copies) | `ba2a753` one shared helper (`utils/excel_com.py:19-29`) | F + refactor |
| A-R101 | Step 3 open with one fallback (8.3 short path) | no | `edac6a3` | `ba2a753` (the open-check line only) | F |
| A-R102 | Step 4 open with several fallbacks | no | `edac6a3` | `ba2a753` (the open-check line only) | F |
| A-R103 | Step 4 Excel sessions (two per report, `Dispatch`) | no | `edac6a3` | none | F |
| A-R104 | 218-character threshold ⚠ (Excel limit, technical) | no | `edac6a3` | `7347383` comment with measured figures; `bf2f1dd` help texts | F |
| A-R105 | Error-dialog handling | no | `edac6a3` | `bd176bf` cosmetic | F |
| A-R106 | openpyxl load options (`keep_links=True`) | no | `edac6a3` (`core/excel_handler.py:242@edac6a3`) | `12d505f` loads the local copy | F |
| A-R107 | `EXCEL_SAFETY` (dead) | no | `edac6a3` | none | F (dead) |
| A-R108 | All Python warnings switched off | no | `edac6a3` | none | F |

**Category 13 — Performance and operational constants**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R109 | One directory listing per scan | no | `a603c02`, `56aaad9` (before: three searches and several file queries per file) | none | P |
| A-R110 | Parallel signature reads (16, from 4 files) ⚠ (tuning) | no | `4f84ba9` (message: 2,508 files on the live share; a cold sequential scan *projected* at about 677 s (about 11 minutes) from about 270 ms per file, against 40.78 s *measured* with the parallel reads; not verifiable here) | none | P (measured) |
| A-R111 | Level-1 folder once per run | no | `275d9a4` | none | P |
| A-R112 | Fixed 0.5 s pauses in Step 3 | no | `edac6a3` | none | F |
| A-R113 | Setup re-check 350 ms after typing; remembered answers | no | `322717f` | `c89c1d4` also forces a re-check on every theme switch | R |

**Category 14 — Settings, cache, logs and run history**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R114 | Settings file `~/.gst_organizer_cache.json` | no | `edac6a3` (writer unchanged) | `c1583d6` the extractor reads the same file | F |
| A-R115 | Settings keys and defaults | no | `edac6a3` | `510f7b4` `workflow_mode` | F |
| A-R116 | When settings are saved | no | `edac6a3` (Browse, dark mode, global tick, length arrows; the processing-mode radios never had a `command`, `gui/tabs/setup_tab.py:216@edac6a3`) | `510f7b4` workflow radios save | F |
| A-R117 | `recent_folders` written, never read | no | `edac6a3` | none | F (dead) |
| A-R118 | How the standalone extractor finds its folder | no | `edac6a3`: only current-folder files the organiser never writes | `c1583d6` the home-folder file first | R |
| A-R119 | Step 4 options not stored | no | `bd176bf` (when the options were added) | none | U |
| A-R120 | Main log in the current folder | no | `edac6a3` | none | F |
| A-R121 | Extractor log | no | `edac6a3` | `bfcb24a` embedded Step 4 bypasses it, so its lines go to `gst_organizer.log` | F + M |
| A-R122 | What is logged and at which level | no | `edac6a3` (many "[DEBUG]" messages at INFO) | `064c634` 35 moved to DEBUG; 14 "[DEBUG]" lines stay in the on-screen log (`gui/handlers/processing_handler.py:333-411`) | R |
| A-R123 | No run history | no | absent at every commit | — | F |

**Category 15 — Start-up and dependencies**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R124 | Which packages are required | no | `f363946` (pywin32 optional) — before, a missing package surfaced as an import error listing source files | `be6418c` pywin32 essential on Windows (`utils/requirements_check.py:57`) | U |
| A-R125 | Offer to install, then restart | no | `a5ad487` | none | U |
| A-R126 | `requirements.txt` contents ⚠ | no | `edac6a3` (openpyxl, lxml) | `6c1e1c6` pytest; `064c634` pywin32 (12F) | F |
| A-R127 | A console is assumed | no | `edac6a3` (UTF-8 re-wrap `main.py:17`) | `f363946` banner and report; `be6418c` `_pause()` | F |
| A-R128 | Program folder prepared at start | no | `edac6a3` | none | F |
| A-R129 | Step 4 loading guard and notice | no | `bfcb24a` (import) / `c463256` (guard, notice) | none | U |
| A-R130 | Version labels (dead) | no | `edac6a3` | none | F (dead) |

**Category 16 — User-interface behaviour rules**

| ID | Rule (short) | GST | Introduced | Later changes | Kind |
|---|---|---|---|---|---|
| A-R131 | Main window, tab names | no | `edac6a3` | `bfcb24a` fourth tab; safe tab-name lookup | F + M |
| A-R132 | Automatic tab switches | no | `edac6a3` | `510f7b4` Step 3 → Step 4 in "full" mode | F + M |
| A-R133 | Setup checklist evaluation | no | `322717f` | `169c7b5` dark colour pairs | R |
| A-R134 | Scan and Dry Run gates | no | `edac6a3` | none | F |
| A-R135 | Run control in Step 3 | no | `edac6a3` | `33d0baf` re-entry guard; `0890022` logging helpers | F + R |
| A-R136 | Long client-name warnings | no | `edac6a3` (both fixed at 10) | `6237597` Start warning uses the length setting; toggle still 10 | R |
| A-R137 | Include-name toggling | no | `edac6a3` | `33d0baf` key fix | F + R |
| A-R138 | Step 2 client table | no | `edac6a3` | `1c48722` row colours follow status | U |
| A-R139 | Dry-run text | no | `edac6a3` | `6237597` regression; `0890022` fixed | F + R |
| A-R140 | Step 3 log and completion | no | `edac6a3` | `0890022` | F |
| A-R141 | Progress and ETA | no | `edac6a3` | `0890022` | F |
| A-R142 | Exported client list | no | `edac6a3` | none | F |
| A-R143 | Step 4 panel behaviour | no | `edac6a3` (standalone window) | `bd176bf`, `c1583d6`, `c1f425e`, `2acffed`, `a677670`, `bfcb24a` (12B.6) | U + M |
| A-R144 | Step 4 folder follows Step 1 | no | `bfcb24a` | `510f7b4` "full" hand-over overrides it | M |
| A-R145 | "Last Refresh Status" text | no | `bd176bf` | `275d9a4`, `2acffed` (recomputed after a run), `2aaa2b1` (configured suffix) | U |
| A-R146 | Mouse-wheel scrolling | no | `edac6a3` (Setup: application-wide binding never released) | `4c42758`, `322717f`, `d64d3a9`, `bf2f1dd` (12E.2) | U |
| A-R147 | Tooltips | no | `322717f` | `bf2f1dd` | R / U |
| A-R148 | Dark-mode colours and capture | no | `edac6a3` | `aeb83cf`, `725fbaf`, `c89c1d4`, `169c7b5` | M / U |
| A-R149 | GUI colours | no | `edac6a3` | none | F |
| A-R150 | `GUI_CONFIG['fonts']` (dead) | no | `edac6a3` | none | F (dead) |
| A-R151 | `GUI_CONFIG['colors']['white']` (dead) | no | `edac6a3` | none | F (dead) |
| A-R152 | Message texts in use | no | `edac6a3` | none | F |
| A-R153 | Unused message texts (dead) | no | `edac6a3` | none | F (dead) |
| A-R154 | Workflow mode values | no | `510f7b4` | `bf2f1dd` hover texts (two disagree with the code, gap A-G66) | M |
| A-R155 | Title strip and status bar | no | `edac6a3` | none | F |
| A-R156 | Version label "v3.0" | no | `edac6a3` | none (unchanged in all 45 later commits, A's 34 own commits included; RAN (me) blame: every cited line `edac6a3`) | F |
| A-R157 | Shortened paths in the status bar | no | `edac6a3` | none | F |

#### 12C.5 Rules that existed earlier but are no longer in the code

These have no registry ID because the registry lists only what exists at the snapshot. Changes that replaced a rule by a registered successor are described in that rule's row above; the items below were removed without a successor.

- **UNREGISTERED: "Invalid or corrupted Excel file" variation** — `edac6a3` (shared - before the split on 30-10-2025) to `a603c02` (A only (after the split)). A second validity check inside `core/file_parser.py` that would have listed a failing file as a variation. It could never fire (the discovery function had already removed such files), so its removal changed nothing; but it means that since the first commit a too-small or wrong-signature file has vanished from the scan without any message (gap A-G37). READ (H-A1 analysis of the caller).
- **UNREGISTERED: old-format extraction mapping `cell_mappings copy.py`** [GST RULE] ⚠ — `edac6a3` to `47a1492` (both shared - before the split on 30-10-2025). A sheet → cell → field dictionary for `Diff GSTR-3B` J15/K15/L15 and `Diff - GSTR-1` N1/O1/P1/Q1, with display names such as "Diff IGST (ITC)" and four commented-out candidates (`E5` "reversed_itc", `F5` "net_itc", `D10` "export_sales", `E10` "exempt_sales") that were never active. It could not be imported (space in the file name). Described only.
- **UNREGISTERED: typed-value check of the length box (10–100) and the hint "⚠️ Names >10 chars will show warning"** — `edac6a3` to `322717f` (A only (after the split)). The validator (`gui/tabs/setup_tab.py:252-260@edac6a3`) refused letters and values outside 10–100; `322717f` dropped it, which is why text typed into the box now blocks saving and stops runs (gap A-G26). The hint had been misleading since `6237597` changed the Start warning to use the setting.
- **UNREGISTERED: a dark-mode switch permanently enlarged the tabs** — `edac6a3` to `725fbaf` (A only (after the split)): `toggle_dark_mode` ended with `style.configure('TNotebook.Tab', padding=[20, 10])` ("Trigger style update"); removed with an explanation (`gui/main_window.py:199-207`). The same padding is still set by the standalone extractor's own style function, which is not called inside the main window (`power_query_extractor/gui/extractor_window.py:85-95`).
- **UNREGISTERED: pywin32 optional on Windows** — `c463256`/`f363946` to `be6418c` (A only (after the split), 10 minutes). For ten minutes the program treated the openpyxl fallback as acceptable on Windows; `be6418c` reversed it because that fallback drops Power Query from the reports. Stale texts from those ten minutes remain (12E.2).

#### 12C.6 Where the older documents disagree with the rule history

| Document and place | Claim | What the history shows |
|---|---|---|
| `PROJECT_AUDIT_ANNUALSTATEMENT.md:1121` | "6 file patterns (regex) — No — Stable since creation" | True on 05-03-2026; **no longer true**: the IMS Reco pattern changed in `5c8ce95` (23-07-2026). |
| same, `:1124` | "Client name max length (35) — edac6a3 — 6237597 — Made configurable via GUI" | **Never true**: the length box and its setting existed in `edac6a3` (`gui/tabs/setup_tab.py:271-277@edac6a3`); `6237597` only made the Start warning use it (A-R136). |
| same, `:1125` | "Thread-safe UI updates — 6237597 — No [change]" | The `6237597` wrappers were partial and introduced wrong names; they were replaced by `_log_async`/`_progress_async` in `0890022` (A-R139). |
| same, `:501` | Name cleaning "Private → Pvt, Limited → Ltd (case-sensitive!)" | **No longer true** since `33d0baf` (whole word, any case) and `4219ffb` (separators normalised first). |
| `tests/README.md:28-31` | Example quirk: "the case-sensitive `Private`→`Pvt` substring replace mangling `Privateer`" | Removed the same morning by `33d0baf`; the test now asserts "Privateer Unlimited" stays as it is (`tests/test_helpers.py:43-45`). |
| `README.md:182-204` | Output names `GST_Processing_Summary_DDMMYYYY_HHMMSS.xlsx`, `ITC_Report_ClientName_State_DDMMYYYY_HHMMSS.xlsx` | **Never true at any commit**: the code has always used `%d%m%y_%H%M` and the state CODE (A-R33, A-R59; `gui/handlers/processing_handler.py:292-298`). |
| `README.md:171-173` | "The extractor auto-loads target folder from cache"; button "Refresh & Extract" | Auto-load worked only from `c1583d6` (23-07-2026) on (A-R118); the button has always been "🚀 Start Processing". |

#### 12C.7 What the rule history tells a rebuilder, and what only the author can confirm

- **Rules settled before git are the oldest and least-tested knowledge in the project.** The patterns, the state table, the Links map, the `BB2` test and the original extraction cells came complete in the first commit, with no history of how they were derived. The tests added in `6c1e1c6` pin many of them as "current behaviour", including at least one quirk the test itself flags (`tests/test_helpers.py:96-97`, "Unknown Place" → "UP", the code used for Uttar Pradesh) [GST RULE] ⚠. Pinning is not confirmation; the author should confirm each before a rebuild relies on it.
- **Rules added after real use came in bursts tied to one template or one share.** `47a1492` added 14 cells in one commit when the ITC template grew new check sheets; `7347383` renamed folders after a measured path-length failure. Both kinds of value change over time ⚠.
- **Review-driven changes were sometimes wrong the first time**: `6237597` (late-binding names), `33d0baf` (underscore names), `322717f` (wheel binding, lost validator), `aeb83cf`/`c89c1d4` (dark mode). Each needed a follow-up commit (12D.4).
- **Questions for the author** (no judgement on the rules): (1) Do the ITC/Sales templates' Power Queries read the folder paths from the Links cells, or do any rely on the old folder names `Other ITC related files` / `Sales related files` / `GSTR-3B Exports` (changed in `7347383`)? (2) Are the 14 extraction cells added on 30-10-2025 (A-R77 – A-R82) still at those sheet and cell positions in today's ITC template ⚠? (3) When both an IMS Reco file and its "(1)" duplicate exist, which should feed the report (since `5c8ce95` the "(1)" copy sorts first)? (4) Is the two-letter state list (A-R21) your reference list ⚠?

### 12D. Edge Cases Discovered in Production

"Bug-fix commits are gold." This sub-section takes every commit of A that fixes a bug or handles a new real-world situation — 36 of the 46, counting the four commits that removed accidentally committed data (the other ten are the first commit, documentation, a clean-up of a malformed first line, tests, a refactor and new features: `edac6a3`, `d674386`, `5ea77c2`, `fc279af`, `f7eb8f1`, `6c1e1c6`, `ba2a753`, `b276542`, `510f7b4`, `a5ad487`; `5ea77c2` is listed in 12E.2 #3 instead) [wording corrected during verification — `5ea77c2` is none of the kinds originally listed] — and records what broke, which input or situation caused it, how it was fixed (a quick **patch** that removes the symptom, or a **proper** fix that removes the cause), and whether similar cases are still unhandled in the snapshot code. Similar cases are cross-referenced to the gap IDs (A-G…) of Section 7.

**How each case was found** (from the commit message, checked against the diff; HIGH unless stated):
- **Reported** — the message says a user reported it, often "with screenshots": `2acffed`, `2aaa2b1`, `725fbaf`, `169c7b5`, `bf2f1dd`.
- **Measured on the real network share** — the message gives numbers from the office share: `4f84ba9` (2,508 files), `7347383` (3,030 files).
- **Hit on the author's computer** — `c463256`, `f363946`, `be6418c` (two Pythons, one without pywin32).
- **Found in a code review** (mostly AI-assisted; 32 July commits carry the trailer "Co-Authored-By: Claude Opus 4.8 (1M context)") — `6237597`, `c1583d6`, `0890022`, `33d0baf`, `064c634`, `322717f`, `c1f425e`, `aeb83cf`, `c89c1d4`. Several items were first listed in the March audit document (`f7eb8f1`), whose to-do list the July commits follow almost word for word (RAN (H-A1): `c1583d6`, `0890022`, `064c634`, `33d0baf` reproduce its items 1, 2, 5, 6 and 9).
- **Repairing a regression that an earlier change introduced** — `4219ffb` (of `33d0baf`), `d64d3a9` (of `322717f`), `c463256` (of `bfcb24a`), `169c7b5` (of `c89c1d4`), `bf2f1dd` (of a pattern from `4c42758` that became harmful with `bfcb24a`), `0890022` (of `6237597`).

#### 12D.1 Summary table

| Commit (date) | Label | What broke or was missing | Fix | Similar cases still open at the snapshot |
|---|---|---|---|---|
| `cbfa4a5`, `9c56a1c`, `1cb878b`, `01862a0` (31-08 / 29-10-2025) | shared - before the split on 30-10-2025 | The run log and compiled files were committed | Deleted; `.gitignore` added — patch (history keeps them) | log written to the start folder (A-G56); `.claude/` not ignored |
| `6237597` (29-10-2025) | shared - before the split on 30-10-2025 | Claimed crashes (most not reproducible); real gain: a failing progress callback no longer aborts a scan | Partial; introduced wrong names in the Dry Run | worker still touches the window (A-G70) |
| `1c48722` (29-10-2025) | shared - before the split on 30-10-2025 | Every client row had the same colour | Override block removed — proper | rows rebuilt in dark mode get light tints (A-G64) |
| `4c42758` (29-10-2025) | shared - before the split on 30-10-2025 | Extractor list did not scroll with the wheel | Wheel handlers — the pattern later broke Step 1 (`bf2f1dd`) | log box truncates small wheel steps; Linux wheel events (A-G68) |
| `bd176bf` (30-10-2025) | shared - before the split on 30-10-2025 | Fixed refresh waits too short/long; need one report only; custom names | Configurable — patch for the waits | no completion check (A-G3, A-G50); suffix not recognised until `2aaa2b1` |
| `47a1492` (30-10-2025) | shared - before the split on 30-10-2025 | ITC template grew new check sheets | Six mappings added — triggered the column cross-wiring | 9 columns show another cell (A-G1) |
| `a603c02` (23-07-2026) | A only (after the split) | Slow scans on the network share | Fewer file-system calls — proper (speed) | bad files vanish silently (A-G37) |
| `12d505f` | A only (after the split) | Slow report building; a failed build left a half-written file | Build locally, copy once — proper for speed, patch for safety | the final copy can still break half-way (A-G12) |
| `275d9a4` | A only (after the split) | Repeated folder scans per client | Cache per run — proper | — |
| `5c8ce95` | A only (after the split) | IMS downloads named "… (1).xlsx" not recognised | Pattern widened — proper | leading "(1) ", "- Copy"; the "(1)" copy is linked (A-G6, A-G38) |
| `c1583d6` | A only (after the split) | Extractor never found the settings file; wrong "latest" Version folder | Same path; date modified — proper | date modified moves when anything is saved in an old folder (A-G45) |
| `0890022` | A only (after the split) | Dry Run listed the last client for every client; log could show the next client | Values captured when scheduled — proper | other direct window calls from the worker (A-G70) |
| `33d0baf` | A only (after the split) | ALL-CAPS names not shortened, "Privateer" mangled; double Start; per-client toggles ignored | Regex; guard; same key — proper (first one caused a regression) | numeric-looking names (A-G25); toggle uses fixed 10 |
| `064c634` | A only (after the split) | Log file flooded; pywin32 missing from requirements | Debug level; requirement — patch | log location; 14 "[DEBUG]" lines on screen (A-G56, A-G69) |
| `4219ffb` (25-07-2026) | A only (after the split) | Underscore names no longer shortened (from `33d0baf`) | Separators normalised — proper for file names | folder key not fixed (A-G8, A-G35) |
| `56aaad9`, `4f84ba9` | A only (after the split) | Full scan of the real share projected at about 11 minutes (per the message: about 677 s from about 270 ms per file; not a measured run) | One listing, parallel reads — proper | scan still freezes the window (A-G48) |
| `322717f` | A only (after the split) | Disabled buttons still clickable; misleading hint; wheel hijacked | New tab — proper for Step 1 only | Step 2/3 buttons (A-G62); lost validator (A-G26); existence-only checks (A-G20) |
| `d64d3a9` | A only (after the split) | Step 1 scrolling dead over content and on touchpads (from `322717f`) | One filtered handler — proper | extractor kept the old pattern until `bf2f1dd` |
| `c1f425e` | A only (after the split) | Extractor updated the window from its worker thread | `_ui` marshalling — proper | two Tk-variable reads remain |
| `2acffed` | A only (after the split) | "Last refresh" column stale after a run (reported) | Recomputed after a run — proper | a scan during a run re-enables Start (A-G65) |
| `a677670` | A only (after the split) | Window tests skipped at random | One shared test root — proper | suite cannot load without pywin32 until `c463256`; start failures become skips |
| `2aaa2b1` | A only (after the split) | Custom suffix files never recognised (reported) | Configured markers — proper for the reported case | short suffixes over-match; suffix not saved (A-G16) |
| `7347383` (26-07-2026) | A only (after the split) | 36% of 3,030 output files over Excel's path limit (measured) | Shorter folder names — patch | no path-length guard (A-G41) |
| `bfcb24a` | A only (after the split) | Selecting the new fourth tab raised an error | Defensive lookup — proper | introduced the pywin32 start-up crash |
| `c463256` | A only (after the split) | App would not start without pywin32 | Guards — proper | notice unreachable; empty workbook "complete" (A-G7, A-G65) |
| `f363946` | A only (after the split) | Crash message blamed missing source files | Start-up check — proper | "found" is not "importable" (A-G7) |
| `be6418c` | A only (after the split) | Without pywin32 reports silently lost Power Query; spurious "EOF" traceback | Refuse to start on Windows — patch at start-up | no save-time Power Query check (A-G7); no-console start (A-G22) |
| `aeb83cf` | A only (after the split) | One dark-mode switch flattened colours permanently | Capture and replay — proper | style maps; widgets created while dark (A-G64) |
| `c89c1d4` | A only (after the split) | Stale state colours after leaving dark mode | Recompute Step 1 — patch (introduced light rows in dark) | other state-coloured widgets |
| `725fbaf` | A only (after the split) | Client list stayed dark; tabs grew (reported) | Record empty values; remove hack — proper | style maps |
| `169c7b5` | A only (after the split) | Step 1 unreadable in dark mode (reported; from `c89c1d4`) | Theme-aware colours for Step 1 — patch (one tab) | other tabs (A-G64) |
| `bf2f1dd` (27-07-2026) | A only (after the split) | Scrolling died after visiting Step 4 (reported) | Extractor handler filtered; guard tests — proper | log box wheel; help texts that disagree with code (A-G66, A-G68) |

#### 12D.2 Each case in detail

**`cbfa4a5` / `9c56a1c` / `1cb878b` / `01862a0` — data committed by accident** (shared - before the split on 30-10-2025)
- What broke: the first commit included the program's run log `gst_organizer.log` (client names, file names and folder paths — described by kind only) and 70 compiled files carrying the developer computer's folder paths, because there was no `.gitignore`.
- Fix: the log was deleted on GitHub three minutes later (`cbfa4a5`), the deletion reached the author's computer through the merge `9c56a1c`; `.gitignore` added (`1cb878b`) and compiled files removed (`01862a0`). A **patch**: both remain readable at `edac6a3` in A's (and B's) history (HIGH; removing them needs a history rewrite — the author's decision).
- Still open: the log is still written to whatever folder the program is started from (`main.py:24`), so starting it from a shared or synced folder puts client data there (A-G56); `.gitignore` does not list `.claude/` (READ `.gitignore`), which is why `.claude/settings.local.json` (harmless content) is tracked.

**`6237597` — "Fix critical import and threading bugs"** (shared - before the split on 30-10-2025)
- What the message claims: a circular import, an attribute error in the scan, thread-safety problems, an indentation bug, a wrong `app.py` reference in `main.py`. What the diff shows (RAN (H-S)): the import was already inside the function (no cycle exists — `utils/constants.py` imports nothing from `utils`); the "attribute error" could not occur (both names point to the same dictionary); no indentation bug existed; `main.py` is not in the commit. The real behaviour changes: (1) a progress callback that raises no longer aborts the scan (`core/file_parser.py:48-54`); (2) the Start warning threshold changed from 10 to the length setting (A-R136); (3) part of the worker's screen updates were deferred with `root.after(0, lambda: …)`.
- Regression introduced: those lambdas read the loop variables when they run, not when they are scheduled. RAN (H-S, fake app, two clients): the Dry Run printed the second client's name and file for every line. Fixed in `0890022`.
- Patch or proper: partial; the claims are a review's guesses rather than observed failures (MEDIUM).
- Still open: the rest of the worker thread still calls window functions and message boxes directly (`gui/handlers/processing_handler.py:319-523` log calls, `:498-502` and `:525` message boxes, `:529-530` button changes) (A-G70); the per-client toggle still warns above a fixed 10 characters (`gui/handlers/client_handler.py:191`).

**`1c48722` — row colours** (shared - before the split on 30-10-2025)
- What broke: in light mode every client row was tagged "complete" (incomplete clients looked green); in dark mode every row was tagged "incomplete" (RAN (H-S)). The message calls the removed lines "redundant", which understates a visible bug.
- Fix: the overriding block removed — **proper** (`gui/handlers/client_handler.py:52-55`).
- Still open: the tags' background colours are reset to the light colours on every rebuild without touching the text colour (`gui/handlers/client_handler.py:58-59`), so a rescan in dark mode gives white text on pale rows (RAN (H-A3), A-G64).

**`4c42758` — wheel scrolling in the extractor** (shared - before the split on 30-10-2025)
- What was missing: with many clients the extractor's client list (a canvas with tick boxes) did not scroll with the wheel.
- Fix: on pointer Enter, bind the wheel for the whole application; on Leave, unbind it for the whole application. Worked in a separate program; **harmful later**: once the extractor became a tab of the main window (`bfcb24a`), the Leave "unbind all" removed Step 1's wheel binding too — the bug `bf2f1dd` fixed.
- Still open: the extractor's log box still scrolls by `int(-1*(event.delta/120))` (`power_query_extractor/gui/extractor_window.py:559-565`), so small touch-pad movements scroll nothing; all wheel handlers listen only to `<MouseWheel>`, which Linux sends as other events (LOW; A targets Windows) (A-G68).

**`bd176bf` — configurable waits, suffix and per-report ticks** (shared - before the split on 30-10-2025)
- What was missing (real use, HIGH by content): the fixed 10-second wait after the refresh key presses was too short for some reports and too long for others; some clients needed only the ITC or only the Sales report; the author wanted other names for refreshed copies and to see when each report was last refreshed.
- Fix: wait N seconds (box 5–60, default 10) and an extended wait of max(10, N) (`power_query_extractor/core/report_processor.py:297-330`); a suffix pattern (`:51, 127-130`); per-report ticks with "Not selected for processing". A **patch** for the timing: the refresh is still assumed finished after fixed waits (no completion check).
- Introduced: the readers of refreshed copies still searched for `_Refreshed_`, so a custom suffix was never found (fixed `2aaa2b1`, RAN (me) 12C.3); an empty suffix makes the copy's name equal the original and the copy fails "are the same file" (RAN (TRACE)).
- Still open: typed values in the wait box are not checked (text makes `IntVar.get()` fail at `extractor_window.py:864`, A-G26); no detection that the refresh finished, no way to stop a long run (A-G3, A-G50).

**`47a1492` — six new ITC extraction sheets** [GST RULE] ⚠ (shared - before the split on 30-10-2025)
- Situation: the ITC template gained check sheets; their key cells were added to the consolidated workbook (A-R77 – A-R82).
- Effect: the consolidator's lookup "first field whose name contains the column title" (`power_query_extractor/core/data_consolidator.py:141-145`) now finds `GSTR-3B_Difference_Diff_IGST` for the column titled `IGST`, so three sheets show the values of `Diff GSTR-3B!J15/K15/L15` (RAN (me): 9 of 22 columns; 0 of 8 before). A technical mismatch; which cell is right is the author's call (A-G1).
- Still open: any future column title that is contained in an earlier field name would be affected the same way (READ, HIGH).

**`a603c02` — fewer round-trips when scanning** (A only (after the split))
- Situation: scanning a source folder on a network share was slow; every file-system call is a separate network round-trip (a request and its answer).
- Fix: one directory listing instead of three searches; `get_file_info` makes two file-system queries instead of four (`exists()` + `stat()`; the file/folder type is now read from the `stat()` answer instead of separate `is_file()`/`is_dir()` calls — the message says "4 calls -> 2") [corrected during verification — originally "a file's type from one query instead of four"]; the duplicate validation in `core/file_parser.py` removed — **proper** (speed). Off Windows it also changed behaviour: upper-case extensions are now found (RAN (H-A1)).
- Still open: files that fail the checks (under 1,024 bytes, wrong first bytes, Excel lock files `~$…`) disappear without any message (`utils/helpers.py:275-306`, A-G37).

**`12d505f` — build reports on local disk, copy once** (A only (after the split))
- Situation: Excel editing a workbook on the network share makes many small reads and writes, each a round-trip; and a failed build left a half-written template copy at the destination (RAN (H-A1): old code left `R.xlsx`, new code nothing).
- Fix: copy the template to `%TEMP%\gst_excel_temp\`, let Excel work there, copy the finished file to the client folder (`core/excel_handler.py:135-155, 212-215`; openpyxl path `:269-322`). **Proper** for speed; a **patch** for safety.
- Still open: the final copy (`core/excel_handler.py:215`, `:308`) writes straight to the final name, so a network drop during it leaves a broken file under the final name (A-G12); the commit asked for one report to be validated on Windows before bulk use — not verifiable here; Excel records the folder a workbook was last saved in (`x15ac:absPath` inside the file), which would now be the temporary folder (MEDIUM; how to check: unzip a report made on Windows after this commit and look at `xl/workbook.xml`).

**`275d9a4` — Level-1 folder cached; single-listing extractor helpers** (A only (after the split))
- Situation: in Re-run/Resume mode the target was re-scanned for every client.
- Fix: cached per run (`core/file_organizer.py:150-161`); the extractor's report checks use one listing (`power_query_extractor/gui/extractor_window.py:20-82`) — **proper**.
- Side effect: the run-wide "Folders Created" count in each text report changed in fresh mode (RAN (H-A1)); the count was, and still is, cumulative for the whole run, not per client (`core/file_organizer.py:362-368`, A-G19).

**`5c8ce95` — IMS Reco files with a "(N)" duplicate marker** [GST RULE] ⚠ (A only (after the split))
- What broke: browser downloads named `ImsReco-…-<8 digits> (1).xlsx` were listed as unrecognised.
- Fix: optional `\s*\(\d+\)` after the date (`utils/constants.py:23`) — **proper** at the pattern.
- Still open (RAN (me) and RAN (H-A1)): a leading "(1) " (accepted only for GSTR-3B), "- Copy", "(1) (1)" and `.xlsm` IMS files are still unrecognised (A-G38); for the other five types a trailing "(N)" is accepted but lands inside the captured period text; when an original and its "(1)" copy both exist, the client shows "Has duplicates" (if no file type is missing — a missing type takes precedence and the status reads "Missing N files"; RAN (me) `VERIFY-A-S12BF/ims_dup.py`, `ims_dup2.py`, made-up names) and the "(1)" copy — which sorts first (`utils/helpers.py:361`) — is the one written into the ITC report's Links cells (`core/excel_handler.py:647-651`) (A-G6).

**`c1583d6` — extractor settings path and Version-folder choice** (A only (after the split))
- What broke: (1) the organiser saves its settings in `~/.gst_organizer_cache.json`, but the extractor read `gst_organizer_cache.json` in the current folder, so "auto-load the last folder" never worked since the first commit; (2) the newest Version folder was chosen by name, and day-first names `Version-DDMMYY HHMM` do not sort by date across a month boundary (RAN (me): `Version-300626 1000` "beats" `Version-150726 0900`).
- Fix: both extractor readers use the organiser's path (`power_query_extractor/extractor_main.py:41`, `power_query_extractor/gui/extractor_window.py:118`); Version folder by date modified (`:648`) — **proper**.
- Still open: "date modified" changes when anything is saved directly inside an older folder (a refreshed copy, an Excel lock file, a consolidated workbook), which can make an older run or version "newest" (RAN (TRACE)/(PQ), A-G45); a FILE named `Annual Statement-…` is not excluded (`extractor_window.py:623`); the three current-folder fallbacks remain (`extractor_main.py:42-44`) [line range corrected during verification — originally `:41-44`; line 41 is the home-folder path].

**`0890022` — late-binding names in the log** (A only (after the split))
- What broke: the `6237597` wrappers. In the Dry Run (which runs on the window's own thread) every scheduled line ran after the loop, so every client heading and file line showed the last client (RAN (H-A1), two placeholder clients); in Step 3 the next client's name could appear whenever the window fell behind.
- Fix: `_log_async` / `_progress_async` bind the finished text when scheduling (`gui/handlers/processing_handler.py:23-36`) — **proper**.
- Still open: about 45 direct `self.app.log_message(...)` calls, two message boxes and the button resets still run on the worker thread (`processing_handler.py:319-530`) (A-G70); the worker still reads Tk variables (`:145, 190-194, 219-225, 301, 310, 332`).

**`33d0baf` — three fixes** (A only (after the split))
- (1) Name cleaning: "Privateer" became "Pvter" and "ALL-CAPS PRIVATE LIMITED" was not shortened → whole-word, any-case regex (`utils/helpers.py:92-93`). Introduced a regression for `X_Private_Limited` (RAN (me): `ClientA_Private_Limited` unchanged), fixed two days later in `4219ffb`.
- (2) Double Start: the Start buttons are wired with `.bind('<Button-1>')`, which fires even when a button is greyed out (RAN (H-A1)); two buttons (`gui/tabs/processing_tab.py:74`, `gui/tabs/validation_tab.py:142`) could start a second worker → guard `if getattr(self.app, 'is_processing', False): return` (`gui/handlers/processing_handler.py:75-76`) — **proper** for Step 3.
- (3) Per-client folder-name toggles were saved under `<client>-<state text>` but read as `<client>-<state code>`, so they were silently ignored since the first commit → same key everywhere (`gui/handlers/client_handler.py:181, 218`) — **proper**.
- Still open: the client table turns number-looking text into numbers ("0042" → 42), so the rebuilt keys (`client_handler.py:167-171, 177-181, 216-218`) do not match and Dry Run, Start (the whole run stops at `processing_handler.py:232`) and Export fail for such clients (A-G25); the toggle's warning uses a fixed 10 (`client_handler.py:191`); run-together names such as `ClientAPrivateLimited` are no longer shortened (pinned by a test with its own made-up name `ABCPrivateLimited`, `tests/test_helpers.py:68-74`) [line corrected during verification — originally `:70`]; the Step 2/3 buttons still cannot be used from the keyboard (A-G62).

**`064c634` — log hygiene; pywin32 declared** (A only (after the split))
- What broke: 35 `logger.info("[DEBUG] …")` lines (including full data-mapping dumps) flooded the log file at INFO; a stray `print` sat above a function's docstring; pywin32, imported by the extractor unconditionally, was missing from `requirements.txt` since the first commit.
- Fix: DEBUG level; `print` removed; `pywin32>=305; sys_platform == 'win32'` added — a **patch** (the log still records client names, paths and figures at INFO: `core/excel_handler.py:195`, `power_query_extractor/core/report_processor.py:56, 603-611`).
- Still open: the log file is in the start folder (`main.py:24`, A-G56); 14 "[DEBUG]" lines are still shown in the on-screen log (`gui/handlers/processing_handler.py:333-411`, "intentionally left" per the message, A-G69); one `logger.error("[DEBUG] …")` remains (`core/excel_handler.py:83`).

**`4219ffb` — underscore-joined names shortened again** (A only (after the split))
- What broke: after `33d0baf`, `X_Private_Limited` kept its full length because `_` counts as a letter for the regex word boundary; that adds 8 characters to every path that embeds the name — relevant to Excel's ~218-character limit ⚠.
- Fix: spaces and underscores turned into a space before the two substitutions (`utils/helpers.py:91`) — **proper** for `sanitize_filename`.
- Still open: the Level-2 client folder name is built by `create_client_state_key` (`utils/helpers.py:620-644`), which never got this step and never replaces spaces, so `ClientA_Private_Limited` gives the folder `ClientA_Private_Limited-DL` while the report name part is `ClientA_Pvt_Ltd` (RAN (me)); different spellings of one client can therefore become separate clients that share or split folders (A-G8, A-G35); names with dots or slashes are mangled by the extension split (A-G40).

**`56aaad9`, `4f84ba9` — scan of the real share from a projected 11 minutes to a measured 41 seconds** (A only (after the split))
- Situation (measured, per the messages; not verifiable here): 2,508 files on the office share over a VPN (Tailscale, about 45 ms per round-trip); `56aaad9` cut the scan from 6.2 to 1.2 file-system round-trips per file ("81% fewer"), after which the 8-byte signature read was the remaining cost (about 270 ms per file cold, `4f84ba9` message) [corrected during verification — originally "after `a603c02`"; both messages say it was after `56aaad9`].
- Fix: reuse the listing's file details (`utils/helpers.py:129-170`); read signatures 16 at a time (`:267-272, 346-361`) — **proper** (RAN (H-A2): same files; simulated latency 3.60 s → 0.26 s).
- Still open: the scan still runs on the window's own thread (`gui/handlers/file_handler.py:57-72`), so the window does not repaint for those 41 seconds (A-G48); a VPN user is the situation Standard 13 says to ask about (`rebuild_pipeline_FINAL.md:927-930`).

**`322717f` — new Setup tab** (A only (after the split))
- What broke (review): disabled Step 1 buttons still reacted (the `.bind` problem again); the hint "Names >10 chars will show warning" no longer matched the code; the Step 1 wheel binding was application-wide and never released, so the wheel scrolled Step 1 while the pointer was on another tab (`gui/tabs/setup_tab.py:66@edac6a3`).
- Fix: checklist, `command=` buttons (`gui/tabs/setup_tab.py:384`), honest length hint — **proper** for Step 1 only.
- Introduced: a broken Enter/Leave wheel toggle (fixed 7 minutes later, `d64d3a9`); the length box lost its validator (A-G26).
- Still open: a FILE passes as the source or target folder and a folder named `x.xlsx` as a template (existence only, `gui/utils/setup_validation.py:47-64`; RAN (TRACE)); a target folder that does not exist yet shows "Folder not found" although processing creates it (`core/file_organizer.py:65-69`); answers are remembered per exact text until a forced refresh (`gui/tabs/setup_tab.py:521-535`) (A-G20); the checklist does not gate Step 3 (A-R51); Step 2 and Step 3 buttons still use `.bind('<Button-1>')` (`gui/tabs/validation_tab.py:118, 130, 142, 190`; `gui/tabs/processing_tab.py:74, 87`) (A-G62).

**`d64d3a9` — Step 1 scrolling fixed** (A only (after the split))
- What broke (from `322717f`): Tk sends "Leave" to a container when the pointer moves onto one of its children, so hovering any card unbound the wheel; touch-pad steps smaller than one notch were truncated to zero.
- Fix: one application-wide handler added with `add='+'` (never replacing others), which scrolls only when the pointer is over the canvas and accumulates small steps (`gui/tabs/setup_tab.py:193-248`) — **proper**.
- Still open at the time: the extractor kept the Enter/Leave pattern (fixed `bf2f1dd`).

**`c1f425e` — extractor worker updates marshalled** (A only (after the split))
- Situation: preparing to share one Tk window with the organiser; the extractor's worker inserted log text, changed progress and raised message boxes directly from its thread (Tk is not safe to use from two threads).
- Fix: every worker update goes through `_ui` → `self.after(0, …)` (`power_query_extractor/gui/extractor_window.py:989-1010`) — **proper**. Two unmentioned changes (READ (H-A2)): Excel clean-up now runs while the "open the report?" question is still on screen; a failure to open the report is only logged.
- Still open: the worker still reads `wait_time_var` and `suffix_pattern_var` directly (`extractor_window.py:890-891`).

**`2acffed` — refresh status after a run (reported)** (A only (after the split))
- What broke: "the app does not pick up the refreshed state" — the "Last Refresh Status" column was computed only when the list was built.
- Fix: `refresh_status_display()` re-reads the times after every run, keeping the user's ticks (`extractor_window.py:759-808`) — **proper**.
- Still open: a Scan or Browse during a run rebuilds the list and re-enables "Start Processing" (`extractor_window.py:755-757`), and `start_processing` has no "already running" check (`:845-883`), so a second run can start on the same Excel (RAN (H-A2)) (A-G65, A-G71). The main window fixed the same hazard for its own Start in `33d0baf`.

**`a677670` — panel split; flaky window tests** (A only (after the split))
- What broke (tests): creating a new Tk root after destroying one made window tests skip at random.
- Fix: one session-wide test root (`tests/conftest.py:32-52`); the standalone window tested in a separate process — **proper** for the flakiness.
- Still open: from this commit until `c463256` the whole suite could not be collected on a PC without pywin32 (RAN (H-A2)); the separate-process tests turn a failed start into a SKIP (`tests/test_extractor_standalone.py:76`, `tests/test_main_window_tabs.py:100`), which is the "silently hiding" effect the commit wanted to remove.

**`2aaa2b1` — refreshed files recognised by the configured suffix (reported)** (A only (after the split))
- What broke: "the refresh state is still wrong" — with a custom suffix the client list said "Never refreshed" right after a refresh, and Skip Refresh refreshed everything again.
- Fix: `refresh_naming.py` derives the markers from the configured pattern and also accepts the old `_refreshed_` (`power_query_extractor/core/refresh_naming.py:23-66`) — **proper** for the reported case.
- Still open (RAN (me) 12C.3; RAN (H-A2)): short suffixes such as `_R{timestamp}`, `_{timestamp}`, `_DL{timestamp}` make the original report count as its own refreshed copy, so Skip Refresh can "reuse" the unrefreshed original and report success (A-G16, A-G2); the suffix is not saved between sessions (`extractor_window.py:297`, A-R119); the report to refresh is the first search hit (`report_processor.py:112-125`), which can be an earlier refreshed copy (A-G15); the default text is typed in three places (`refresh_naming.py:14`, `report_processor.py:51`, `extractor_window.py:297`).

**`7347383` — shorter category folders (measured)** ⚠ (A only (after the split))
- What broke: on a real run of 3,030 output files on the network share, the worst full path was 270 characters and 1,094 files (36%) exceeded Excel's ~218-character limit, "which is why long client names have been failing" (commit message).
- Fix: `Other ITC related files` → `ITC`, `Sales related files` → `Sales`, `GSTR-3B Exports` → `GSTR3B` (`utils/constants.py:91-95`) — a **patch**: the message itself says the margin is thin (215 vs 218, modelled together with a planned shorter share root whose name carries a financial-year label) and that "a path-length guard remains the durable fix". The code comment (`utils/constants.py:87-90`) calls the modelled figures "Measured" and drops the share-root assumption (READ (H-A3)).
- Still open: nothing measures a path before folders and files are created (`core/file_organizer.py:117-138, 270-291`) (A-G41); with "include client name" on, the category folder uses the full cleaned client name, not the shortened one (`core/file_organizer.py:90`); copied file names keep the full client name (`:271`).

**`bfcb24a` — Step 4 tab; tab-name error** (A only (after the split))
- What broke (found while adding the tab): `on_tab_changed` indexed a hard-coded three-item list, so selecting the fourth tab raised an error inside a Tk callback (the message says "on every tab change"; by reading, only when Step 4 became the current tab — READ (H-A3)).
- Fix: list extended and a bounds check (`gui/main_window.py:331-338`) — **proper** (RAN (H-A3) mutation check).
- Introduced: importing the main window now imported the extractor, whose pywin32 imports were unconditional — the program could no longer start on a Python without pywin32 (RAN (H-A3); fixed `c463256`). The new tests could only pass where pywin32 was installed, which is how it slipped through.

**`c463256` — start without pywin32** (A only (after the split))
- What broke: on the author's PC, `py` started a Python without pywin32 and `python` one with it; after `bfcb24a` the app died at start-up with "No module named 'pythoncom'".
- Fix: guarded imports with `WIN32_AVAILABLE`, a per-client refusal (`power_query_extractor/core/report_processor.py:20-36, 62-77`), a guarded Step 4 import with a notice (`gui/tabs/extract_tab.py:24-30, 84-105`) — **proper** for the crash.
- Still open: the notice cannot appear for a missing pywin32 any more (the extractor now loads without it; RAN (H-A3)) and would blame pywin32 for any other failure; without pywin32 Step 4 refuses every client yet still writes a consolidated workbook and says "Processing complete!" (`extractor_window.py:941-955`; RAN (H-A3)); the guard catches only `ImportError` (`report_processor.py:27`); the comments and notice still call pywin32 optional (`report_processor.py:11-13`, `extract_tab.py:21-23, 101-105`) although `be6418c` made it essential on Windows (12E.2).

**`f363946` — misleading crash message** (A only (after the split))
- What broke: the import error told the user to look for missing source files (including `gui/app.py`), when the cause was a missing package for the Python being used.
- Fix: a package check before the window is imported, a report naming the interpreter (`utils/requirements_check.py`, `main.py:62-73`) — **proper**.
- Still open: `find_spec` (a check that a module can be *found*, without loading it) reports a pywin32 that is present but cannot load as installed (RAN (H-A3)), so on Windows the openpyxl fallback could still run (A-G7); every message goes to `print()`, which shows nothing without a console (A-G22).

**`be6418c` — refuse to start on Windows without pywin32** (A only (after the split))
- What could break silently: without Excel COM, report creation falls back to openpyxl, which drops Power Query (Standard 12, `rebuild_pipeline_FINAL.md:846-850`), with only a log warning (`core/excel_handler.py:96-102`). Also: a blocked start printed its reason and then a spurious "EOF when reading a line" error.
- Fix: pywin32 essential when `sys.platform == 'win32'` (`utils/requirements_check.py:57`); `_pause()` (`main.py:45-54`). A **patch at start-up**, not at save time: the report writer still has its fallback, and nothing checks a workbook for Power Query before openpyxl saves it (RAN (H-A3): no "DataMashup" or "customXml" in any `.py` file).
- Still open: the `find_spec` false positive above (A-G7); `_pause()` does not catch the error `input()` raises when there is no input stream at all (`main.py:53`; RAN (H-A3)) — moot today because `main.py:17` fails first without a console (A-G22); two different Windows tests (`utils/requirements_check.py:17` vs `core/excel_handler.py:21`).

**`aeb83cf` — dark mode rebuilt** (A only (after the split))
- What broke: the first-commit dark mode guessed the light colours back (frames forced to white, button text to white), so one switch flattened every coloured surface; tree tags were never reset in light mode.
- Fix: capture every widget's colours before the first change and replay them; 11 ttk styles; weak references so destroyed widgets do not pile up — **proper** (RAN (H-A3): 241 widgets, 0 mismatches over 3 round trips at the snapshot).
- Still open: style *maps* (the "selected" colours of tabs and tree rows) are restored to fixed values, not the originals (`gui/utils/dark_mode_manager.py:164-172`; RAN (H-A3)); widgets created while dark mode is on are not themed until the next switch (RAN (H-A3)) (A-G64); the commit's claim that Tk "reuses path names" could not be reproduced with automatic names (MEDIUM; not sure — how to check: key a test build by `str(widget)` and rebuild the client rows).

**`c89c1d4` — stale state colours** (A only (after the split))
- What broke: a checklist row that turned invalid while dark came back with its old green tint after leaving dark mode.
- Fix: `toggle_dark_mode` re-runs the Step 1 check (`gui/main_window.py:180-194`) — a **patch** for one tab, which painted light rows on the dark window (RAN (H-A3)); fixed by `169c7b5`.
- Still open: other state-coloured widgets (the client table tags, `gui/handlers/client_handler.py:58-59`; the Step 1 "problem" label keeps the white text dark mode gave it, `gui/tabs/setup_tab.py:388-390`) (A-G64); its two tests do not exercise the behaviour (one checks two colour constants, one searches the function's source text for a word; READ (H-A3)).

**`725fbaf` — tree background and tab padding (reported with screenshots)** (A only (after the split))
- What broke: after leaving dark mode the client list stayed dark grey (an empty style value was never recorded) and the tabs stayed larger (the first-commit `padding=[20, 10]` hack).
- Fix: record empty values (`gui/utils/dark_mode_manager.py:87-102`); hack removed (`gui/main_window.py:199-207`) — **proper**.
- Still open: style maps (above).

**`169c7b5` — Step 1 readable in dark mode (reported with a screenshot)** (A only (after the split))
- What broke (from `c89c1d4`): light checklist rows on a dark window; accent stripes lost.
- Fix: light and dark colour pairs chosen by theme (`gui/tabs/setup_tab.py:29-67, 531-608`) — a **patch for one tab**.
- Still open: the other tabs (A-G64); two keys of the new colour sets are never read (`gui/tabs/setup_tab.py:56, 62`), so card borders keep the light colour (12E.5).

**`bf2f1dd` — scrolling died after visiting other tabs (reported); hover help** (A only (after the split))
- What broke: "scrolling worked on Step 1, but after hovering the other tabs and coming back it was dead until restart" — the extractor's Leave handler removed every application-wide wheel binding, including Step 1's (RAN (H-A3): 1 handler → 0).
- Fix: the extractor uses the Step 1 pattern (`power_query_extractor/gui/extractor_window.py:434-478`); `tests/test_scroll_bindings.py` forbids `unbind_all` and requires `add='+'` in window files — **proper**, and the first time a lesson was turned into an enforced test. The message says it plainly: "This is the second time this exact pattern has broken scrolling."
- Still open: the log box's wheel (above); several new hover texts disagree with the code — "Fast - no Excel automation involved" for organise-only (`gui/tabs/setup_tab.py:96-97`; on Windows reports are built through Excel, `core/excel_handler.py:90-95`), "then open each report in Excel to refresh" for the full pipeline (`:98-100`; the refresh is not started, `gui/handlers/processing_handler.py:139-142`), Resume "skipping files that are already in place" (`:108-109`; only within the same minute) (A-G66, A-G10).

#### 12D.3 Edge cases visible in the first commit (met before git)

The first commit already contains defences whose comments or structure show problems met during the 2½ months before git (READ; MEDIUM that each was a real incident):
- Excel busy: "Call was rejected by callee" retried 3 times (`power_query_extractor/core/report_processor.py:239-249`); Excel properties failing to set, retried 5 times ("Increased from 3 to 5", `:253`).
- Long paths: 8.3 short paths when a path exceeds 218 characters (`core/excel_handler.py:355-356`), `\\?\` extended paths (`report_processor.py:720-732`) ⚠.
- Files open elsewhere and OneDrive/SharePoint sync: guidance text in the Step 4 open routine (`report_processor.py:788-797`).
- Window activation failing before the key presses (`report_processor.py:258-276`); Excel error dialogs during a refresh (`:305-342, 417-525`).
- Duplicate downloads named `(2) GSTR3B-…` (`utils/constants.py:31`).
- Earlier versions corrupting Excel files / losing Power Query: the README's "Excel corruption issues from previous versions have been specifically addressed" (`README.md@edac6a3`) and `core/excel_handler.py:2-3` "WINDOWS VERSION - Uses Excel COM to preserve Power Query".
- Timing found by trial: "Increased from 0.5" (`report_processor.py:283, 288`), "WAIT 7 SECONDS AS REQUESTED" above a 10-second loop (`report_processor.py@edac6a3`, removed `bd176bf`).

#### 12D.4 Fixes that needed a second fix

| First change | What it broke | Second change | Time between |
|---|---|---|---|
| `6237597` (shared - before the split on 30-10-2025) deferred updates with lambdas | wrong client names in the Dry Run and log | `0890022` | 9 months |
| `33d0baf` word-boundary abbreviation | `X_Private_Limited` no longer shortened | `4219ffb` | 2 days |
| `322717f` Enter/Leave wheel toggle | Step 1 scrolling dead over content | `d64d3a9` | 7 minutes |
| `4c42758` (shared - before the split on 30-10-2025) extractor Enter/Leave with "unbind all" | harmless until `bfcb24a` merged the windows, then killed Step 1 scrolling | `bf2f1dd` | 9 months (about 11 hours after the merge: `bfcb24a` 26-07-2026 20:40, `bf2f1dd` 27-07-2026 07:48) [corrected during verification — originally "1 day after the merge"] |
| `bfcb24a` extractor imported by the main window | app would not start without pywin32 | `c463256` | 2½ hours |
| `c463256` / `f363946` pywin32 "optional" | on Windows reports could silently lose Power Query | `be6418c` | 10 minutes |
| `aeb83cf` capture once, replay | stale state colours | `c89c1d4` | 33 minutes |
| `c89c1d4` recompute Step 1 in light colours | light rows in dark mode | `169c7b5` | 13 minutes |
| `aeb83cf` skip empty style values | client list stayed dark | `725fbaf` | 41 minutes |

Pattern (inference, HIGH): changes to cross-cutting mechanisms — event bindings for the whole application, imports pulled in at start-up, theme switching — repeatedly broke something in another part of the window. The July work found these within minutes because the author tested the real app after each commit and reported with screenshots; tests caught only the logic that had been separated from Tk.

#### 12D.5 Similar cases still unhandled at the snapshot — one list

READ at the snapshot unless marked; each item is described in Section 7 under its gap ID.

1. Nine consolidated columns take another cell's value — `power_query_extractor/core/data_consolidator.py:141-150` [GST RULE] ⚠ (A-G1; RAN (me)).
2. Refresh not confirmed; a failed or never-refreshed copy can be reused as good — `report_processor.py:297-374`, `refresh_naming.py:43-66` (A-G2, A-G3, A-G16; RAN (me) for the suffix part).
3. No path-length guard before creating folders and files — `core/file_organizer.py:117-138, 270-291` ⚠ (A-G41).
4. Silent dropping of files that fail the checks — `utils/helpers.py:275-306` (A-G37).
5. Number-looking client names break selection — `gui/handlers/client_handler.py:167-171` (A-G25).
6. The Level-2 folder key never received the separator step — `utils/helpers.py:620-644` (A-G8, A-G35; RAN (me)).
7. "Newest" by date modified can pick an older run or version — `power_query_extractor/gui/extractor_window.py:631, 648`; `core/file_organizer.py:184` (A-G45).
8. Step 4 can start a second run — `extractor_window.py:755-757, 845-883` (A-G65, A-G71).
9. Step 3's worker still touches the window directly — `gui/handlers/processing_handler.py:319-530` (A-G70).
10. Both scans run on the window's thread — `gui/handlers/file_handler.py:57-72`; `extractor_window.py:611-668` (A-G48).
11. Disabled Step 2/3 buttons still react; no keyboard use — `gui/tabs/validation_tab.py:118, 130, 142, 190`; `gui/tabs/processing_tab.py:74, 87` (A-G62).
12. Text in the length or wait box blocks saving and stops runs — `gui/tabs/setup_tab.py:440-442`; `extractor_window.py:270-279, 864` (A-G26).
13. Power Query loss still possible through a "found but not importable" pywin32 — `utils/requirements_check.py:64-70`; `core/excel_handler.py:21-28, 96-102` (A-G7).
14. Start without a console dies silently — `main.py:17` (A-G22; RAN (H-A3) simulation).
15. Dark mode: style maps and widgets created while dark — `gui/utils/dark_mode_manager.py:164-172`; `gui/handlers/client_handler.py:58-59` (A-G64).
16. Help texts that promise behaviour the code does not have — `gui/tabs/setup_tab.py:96-100, 108-109, 119-123` (A-G66).
17. Duplicate downloads: the "(1)" copy is the one linked — `utils/helpers.py:361`, `core/excel_handler.py:628-700` [GST RULE] (A-G6).

#### 12D.6 Where the older documents disagree with the bug history

| Document and place | Claim | What the code and history show |
|---|---|---|
| `PROJECT_AUDIT_ANNUALSTATEMENT.md:1136` | `6237597`: "Circular import: `from utils.constants` fails when running as package" | No cycle exists (`utils/constants.py` imports nothing from `utils`); the import was already inside the function (READ (H-S)) — the claim was never true. |
| same, `:1138` | `6237597`: "AttributeError: `file_parser.scanned_files`" fixed | Both names refer to the same dictionary; no error was possible (READ (H-S)) — never true. |
| same, `:1139` | `1c48722`: "Dark mode tree colors incorrect after toggle" | The bug was in both modes: in light mode every row looked complete (RAN (H-S)). |
| same, `:1140` | `4c42758`: unhandled sibling "Setup tab canvas in main app also binds `bind_all` which may conflict" | The prediction came true only after `bfcb24a` put both in one window, and was fixed in `bf2f1dd` (`tests/test_scroll_bindings.py`). |
| same, `:1141` | `bd176bf`: extended wait still `max(10, self.wait_time)` | Still true (`report_processor.py:326`). |
| `README.md` (section "Recent Bug Fixes (Latest Commit)", added `d674386`) | Repeats `6237597`'s claims | Same as the first two rows. |

### 12E. Abandoned Approaches and Reversals

What was tried, why it probably did not work, what replaced it — and what is still left in the code as a remnant. "Remnant" means code that is still present at the snapshot but no longer does its original job: dead methods (defined but never called), dead constants, commented-out code, and texts that describe an approach that was dropped.

#### 12E.1 Approaches abandoned before git (known only from the first commit)

| What was tried | Evidence | What replaced it | Why it probably failed (inference) | Lesson |
|---|---|---|---|---|
| One large window file `gui/app.py` (89,275 bytes) | compiled leftovers in `edac6a3` (12B.4); `README.md@edac6a3` folder tree; `main.py:72@edac6a3` | the modular `gui/` package on 23-07-2025 (`gui/main_window.py:4` "Modularized version") | too large to change safely through AI-assisted edits (MEDIUM; reasons lost with the pre-git history) | Keep window code split by screen and by concern from the start. |
| Building reports with openpyxl alone | `core/excel_handler.py:2-3` "WINDOWS VERSION - Uses Excel COM to preserve Power Query"; `README.md` "Excel corruption issues from previous versions have been specifically addressed"; the fallback's own warnings `core/excel_handler.py:98, 267` | Excel through COM on Windows, openpyxl only as a fallback | openpyxl drops Power Query when it saves (Standard 12, `rebuild_pipeline_FINAL.md:827, 846-850`) (MEDIUM-HIGH) | A Power Query template must never be saved by openpyxl. In A this lesson is enforced only by a start-up gate since `be6418c` (12D). |
| An older extraction mapping format (sheet → cell → field), 21-06-2025 | `power_query_extractor/config/cell_mappings copy.py@edac6a3` (deleted `47a1492`) | `EXTRACTION_CONFIG` (output sheet → list of mappings) plus a generated "backward compatibility" copy in the old shape (`power_query_extractor/config/cell_mappings.py:259-278`) that the processor still reads (`report_processor.py:37`) | one consolidated sheet per output group was wanted (HIGH by structure) | The old shape survives as an internal format; a second mapping of the same cell silently replaces the first (A-G78). |
| An older consolidated layout (one sheet with an ITC section, a Sales section, a details sheet) | the never-called `_write_section` (`power_query_extractor/core/data_consolidator.py:250-352`) and `_create_details_sheet` (`:354-414`); the 3.12 compiled copy of the consolidator records a source of 10,208 bytes (25-06-2025), about 58% of the 17,711 bytes recorded by the 3.13 copy (14-07-2025), which matches the committed file once its 432 lines are counted with Windows line endings (17,280 + 432) (RAN (H-S), header fields; re-checked RAN (me) `VERIFY-A-S12BF/pyc.py`) [corrected during verification — originally "about half the size"] | one sheet per configured output group + Summary (`data_consolidator.py:20-186`) | layout grew with the number of checks (MEDIUM) | The dead layout is the only code that would write the failure REASON into the workbook (A-R94). |
| Shorter waits and fewer retries in the refresh | comments "Increased from 0.5" (`report_processor.py:283, 288`), "Increased from 3 to 5" (`:253`), "WAIT 7 SECONDS AS REQUESTED" above a 10-second loop (`report_processor.py@edac6a3`, removed `bd176bf`) | longer fixed waits, later a user-set wait (`bd176bf`) | key-press refresh with fixed waits cannot know when Excel has finished (HIGH) | Waiting longer is a tuning knob, not a completion check (gap A-G3). |

#### 12E.2 Code added and later removed or replaced (in date order)

| # | Tried (commit, label) | Removed or replaced by | Why it likely failed | What replaced it | Remnant at the snapshot |
|---|---|---|---|---|---|
| 1 | Committing the run log and compiled files (`edac6a3`, shared - before the split on 30-10-2025) | `cbfa4a5`, `9c56a1c`, `1cb878b`, `01862a0` (shared - before the split on 30-10-2025) | no `.gitignore`; the log holds client data | deletion + `.gitignore` | both still in history at `edac6a3` |
| 2 | A block that re-tagged every client row (`edac6a3`, shared - before the split on 30-10-2025) | `1c48722` (shared - before the split on 30-10-2025) | it overrode the status colour (all rows alike) | status-only tagging (`gui/handlers/client_handler.py:52-55`) | — |
| 3 | A malformed first line `# !/usr/bin/env python3` (`edac6a3`, shared - before the split on 30-10-2025) | `5ea77c2` (shared - before the split on 30-10-2025) | never valid (space after `#`) | removed | — |
| 4 | A fixed long-name threshold of 10 at Start (`edac6a3`, shared - before the split on 30-10-2025) | `6237597` (shared - before the split on 30-10-2025) | noisy when the default folder length is 35 | the length setting (fallback 10) | the toggle still uses 10 (`gui/handlers/client_handler.py:191`); the code comment "CHANGE: …" (`gui/handlers/processing_handler.py:86-89`) |
| 5 | Direct window calls from the Step 3 worker (`edac6a3`, shared - before the split on 30-10-2025) → `root.after(0, lambda: …)` (`6237597`, shared - before the split on 30-10-2025) | `_log_async` / `_progress_async` (`0890022`, A only (after the split)) | lambdas read loop variables late: wrong names | values captured at schedule time | about 45 direct calls remain (A-G70) |
| 6 | One "select client" tick per client in the extractor; both reports always refreshed (`edac6a3`, shared - before the split on 30-10-2025) | per-report ticks (`bd176bf`, shared - before the split on 30-10-2025) | some clients need only one report | ITC and Sales ticks | — |
| 7 | Fixed waits of 10 s and an extended 10 s (`edac6a3`, shared - before the split on 30-10-2025) | user wait N and max(10, N) (`bd176bf`, shared - before the split on 30-10-2025) | refresh time varies by report | a number box 5–60 | still no completion check |
| 8 | Fixed refreshed-copy name `_Refreshed_<ts>` (`edac6a3`, shared - before the split on 30-10-2025) → user suffix with readers still searching `_Refreshed_` (`bd176bf`, shared - before the split on 30-10-2025) → readers searching `_refreshed_` (`275d9a4`) | `power_query_extractor/core/refresh_naming.py` (`2aaa2b1`, A only (after the split)) | a custom suffix was invisible (reported twice) | configured markers + legacy marker | over-matching for short suffixes (A-G16); the default typed three times |
| 9 | The old-format mapping copy (`edac6a3`, shared - before the split on 30-10-2025) | deleted (`47a1492`) | superseded backup | — | the generated old-format `CELL_MAPPINGS` (`cell_mappings.py:259-278`) |
| 10 | Three file searches per scan and four file queries per file (`edac6a3`, shared - before the split on 30-10-2025) | one listing (`a603c02`) → cached details (`56aaad9`) → parallel signature reads (`4f84ba9`) (all A only (after the split)) | slow on a network share | `scan_excel_files` (`utils/helpers.py:309-365`) | `find_excel_files` survives only as a wrapper used by tests (`utils/helpers.py:368-373`; callers `tests/test_file_ops.py:58, 69, 73, 91`) and re-exported by `utils/__init__.py:16, 32`; `validate_excel_file` imported but unused in `core/excel_handler.py:38` |
| 11 | Validating each file twice, with an "Invalid or corrupted Excel file" variation (`edac6a3`, shared - before the split on 30-10-2025) | `a603c02` | the second check could never fail | one validation | files that fail now vanish silently (A-G37) |
| 12 | Building reports directly on the network target (`edac6a3`, shared - before the split on 30-10-2025) | local staging (`12d505f`) | many small network writes by Excel | `%TEMP%\gst_excel_temp\` | the temp folder is never removed (A-R98) |
| 13 | Re-resolving the Level-1 folder for every client (`edac6a3`, shared - before the split on 30-10-2025) | per-run cache (`275d9a4`) | repeated network scans | `_get_level1_folder` / `_compute_level1_folder` | — |
| 14 | Choosing the newest Version folder by name (`edac6a3`, shared - before the split on 30-10-2025) | by date modified (`c1583d6`) | day-first names do not sort by date | `max(..., key=st_mtime)` (`extractor_window.py:648`) | date modified has its own pitfalls (A-G45) |
| 15 | The extractor reading a settings file in the current folder (`edac6a3`, shared - before the split on 30-10-2025) | home-folder file first (`c1583d6`) | it never found the organiser's file | `Path.home() / '.gst_organizer_cache.json'` | three current-folder fallbacks and three extra key names that nothing writes (`power_query_extractor/extractor_main.py:41-55`) |
| 16 | Case-sensitive substring "Private"/"Limited" replacement (`edac6a3`, shared - before the split on 30-10-2025) | whole-word, any case (`33d0baf`) → plus separator normalisation (`4219ffb`) | mangled "Privateer"; missed capitals; then missed underscores | `utils/helpers.py:91-93` | the Level-2 key never got the separator step (`utils/helpers.py:620-644`) |
| 17 | "[DEBUG]" messages logged at INFO (`edac6a3`, shared - before the split on 30-10-2025) | DEBUG level (`064c634`) | flooded the log with path dumps | `logger.debug` | 14 on-screen "[DEBUG]" lines (`gui/handlers/processing_handler.py:333-411`); one `logger.error("[DEBUG] …")` (`core/excel_handler.py:83`) |
| 18 | A `print("PROCESSING: …")` placed above a function's docstring (`edac6a3`, shared - before the split on 30-10-2025) | `064c634` | turned the docstring into dead text | removed | — |
| 19 | Two copies of the "already open?" loop (`edac6a3`, shared - before the split on 30-10-2025) | `utils/excel_com.py` (`ba2a753`) | duplication | `find_open_workbook` | the two open-with-fallback strategies were kept apart on purpose ("Do NOT merge those two strategies without validating BOTH flows against real Microsoft Excel", `utils/excel_com.py:3-14`) |
| 20 | Setup tab v1: welcome card, instructions card, red "TARGET FOLDER ⭐ IMPORTANT!" panel, hand-typed pattern list, `.bind('<Button-1>')` buttons, a wheel binding for the whole application never released, a length validator 10–100 (`edac6a3`, shared - before the split on 30-10-2025) | rewritten (`322717f`) | disabled buttons still worked; wheel hijacked; nothing showed what was missing | checklist screen | the validator was lost, not replaced (A-G26) |
| 21 | Enter → bind / Leave → unbind wheel toggle on Step 1 (`322717f`) | one filtered handler (`d64d3a9`, 7 minutes later) | Tk sends Leave when the pointer moves onto a child | `gui/tabs/setup_tab.py:193-248` | — |
| 22 | Wheel steps `int(delta/120)` (`edac6a3` Step 1, `4c42758` extractor, both shared - before the split on 30-10-2025) | accumulated fractional steps (`d64d3a9` Step 1, `bf2f1dd` extractor list) | touch pads send smaller steps → 0 | `_wheel_accum` | still in the extractor's log box (`extractor_window.py:559-565`) |
| 23 | A new Tk root per test module (`322717f`) | one session root (`a677670`, about 3½ hours later: 19:24 → 23:01 on 25-07-2026) [corrected during verification — originally "1½ hours later"] | Tk cannot reliably create a root after one was destroyed | `tests/conftest.py:32-52` | — |
| 24 | The extractor as a whole window class (`edac6a3`, shared - before the split on 30-10-2025) | panel + shell (`a677670`) | a window cannot be a tab | `ExtractorPanel` | the shell's `__getattr__` forwarder, used only by tests (`extractor_window.py:1051-1063`) |
| 25 | A hard-coded three-name tab list (`edac6a3`, shared - before the split on 30-10-2025) | defensive lookup (`bfcb24a`) | broke with a fourth tab | bounds check | the list is still hard-coded (`gui/main_window.py:334`) |
| 26 | Unconditional pywin32 imports in the extractor (`edac6a3`, shared - before the split on 30-10-2025) | guarded (`c463256`) | harmless while separate; fatal once imported by the main window (`bfcb24a`) | `WIN32_AVAILABLE` | the unreachable Step 4 notice (`gui/tabs/extract_tab.py:67-69, 84-105`) |
| 27 | "pywin32 is optional; the openpyxl fallback degrades gracefully" (`c463256`, `f363946`, test `test_pywin32_is_optional_not_fatal`) | essential on Windows (`be6418c`, 10 minutes later) | the fallback strips Power Query from the reports | `essential=ON_WINDOWS` (`utils/requirements_check.py:57`) | stale comments and notice still say "optional" (`power_query_extractor/core/report_processor.py:11-13`; `gui/tabs/extract_tab.py:21-23, 101-105`) |
| 28 | Refuse to start (`be6418c`) | offer to install, then refuse (`a5ad487`) | typing the pip command | `utils/dependency_installer.py` | — |
| 29 | Bare `input("Press Enter…")` exits (`edac6a3`, shared - before the split on 30-10-2025; `f363946`) | `_pause()` (`be6418c`) | a spurious "EOF when reading a line" error | `main.py:45-54` | — |
| 30 | An import-error message listing source files incl. `gui/app.py` (`edac6a3`, shared - before the split on 30-10-2025) | real error + interpreter + pip hint (`f363946`) | pointed at the wrong cause | `main.py:114-127` | — |
| 31 | Class-by-class dark mode with guessed restore (`edac6a3`, shared - before the split on 30-10-2025) | capture and replay (`aeb83cf`) | one switch flattened coloured surfaces | `gui/utils/dark_mode_manager.py` | `_original_root_bg` written, never read (`:70, 83`); style maps still restored to fixed values (`:164-172`) |
| 32 | Inside `aeb83cf` itself (from its message, not in history): a one-pass capture and widget-path keys | dropped before committing | the root was recorded already dark; destroyed widgets' colours replayed onto new ones | two passes; weak keys by object | the comment at `dark_mode_manager.py:61-67` |
| 33 | Skipping empty ttk values when capturing (`aeb83cf`) | record them (`725fbaf`) | the client list stayed dark | `dark_mode_manager.py:87-102` | — |
| 34 | `TNotebook.Tab padding=[20, 10]` "to trigger a style update" on every theme switch (`edac6a3`, shared - before the split on 30-10-2025) | removed (`725fbaf`) | permanently enlarged the tabs | `update_idletasks` | the same padding in the standalone extractor's style function (`extractor_window.py:85-95`) |
| 35 | Light-only recompute of the Step 1 colours (`c89c1d4`) | theme-aware pairs (`169c7b5`, 13 minutes later) | light rows on a dark window | `STATE_STYLE_DARK`, `DARK_SURFACE` | the unused `card`/`border` keys (`gui/tabs/setup_tab.py:56, 62`) |
| 36 | Enter/Leave bind/unbind-all wheel binding in the extractor list (`4c42758`, shared - before the split on 30-10-2025) | one additive filtered handler (`bf2f1dd`) | "unbind all" deleted Step 1's handler once both shared a window | `extractor_window.py:434-478` + `tests/test_scroll_bindings.py` | — |
| 37 | Long category folder names (`edac6a3`, shared - before the split on 30-10-2025) | short names (`7347383`) | 36% of 3,030 real files over Excel's path limit ⚠ | `GSTR3B`, `ITC`, `Sales` | old names still in `README.md:185-187` and `PROJECT_AUDIT_ANNUALSTATEMENT.md:152-154, 163, 180, 231, 524-526` (RAN (me) grep for the three old names) [line list corrected during verification — originally `153-154, 163, 180, 525-526`] |
| 38 | The first audit text (`fc279af`, A only (after the split)) | rewritten next morning (`f7eb8f1`) | the method asked for more (dead code, traces, consistency) | v2 | v2 itself never updated after 05-03-2026; it dropped the v1 table of extraction cells [GST RULE] (RAN (H-A1): "K29", "T6 of R9", "T8 of R9" occur 0 times in v2) |

#### 12E.3 Swapped libraries and approaches

- **Third-party libraries: none was ever swapped in A.** RAN (me), `deps_by_commit.py`: the set of third-party modules imported by the program is the same in all 46 commits — `openpyxl`, `pythoncom`, `win32com`, `win32api`, `win32con`, `win32gui` — and `pytest` appears only in the tests from `6c1e1c6`. (Project B's switch to PyQt5 is not part of A's history; A's March audit calls those B commits an "unmerged branch … subsequently abandoned", `PROJECT_AUDIT_ANNUALSTATEMENT.md:8, 18, 1065-1091, 1145` — see 12E.6.)
- **Approaches swapped inside A** (READ): openpyxl → Excel through COM for writing reports (before git); `glob` → `os.scandir` (`a603c02`); sequential → a pool of 16 worker threads for signature reads (`4f84ba9`); per-file `stat` → the directory listing's cached details (`56aaad9`); inline lambdas → value-capturing helpers (`0890022`, `c1f425e`); "guess the colour back" → capture and replay with weak references (`aeb83cf`); two programs that talk through a settings file → one window (`bfcb24a`); popup-only validation → a live checklist (`322717f`); "refuse to start" → "offer to install, then refuse" (`a5ad487`).
- **Approaches deliberately kept** (recorded so a rebuild knows they were considered): the key-press refresh with fixed waits was declared out of scope (`docs/superpowers/specs/2026-07-25-merge-extractor-design.md:125-126`); the long `process_files_thread` method was left alone (`:127`); "any change to how reports are produced" was also out of scope (`:128`), although `7347383` changed the folder paths written into every new report; the two different "open a workbook with fallbacks" routines were kept apart until both can be tested against real Excel (`utils/excel_com.py:3-14`).

#### 12E.4 Features started but never finished (partial implementations still in the code)

| Item | Where (snapshot) | Since | State | Evidence |
|---|---|---|---|---|
| Calculated columns (subtract / add / multiply / divide of two cells) [GST RULE] | `power_query_extractor/core/data_consolidator.py:130-136, 188-218`; documented `power_query_extractor/config/cell_mappings.py:1-27` | `edac6a3` | no mapping has ever used a calculation (A-R85) | READ; RAN (TRACE) |
| Description lines above each consolidated sheet | `data_consolidator.py:52-65` | `edac6a3` | no output group has a `description` (A-R92) | READ |
| `auto_refresh_if_missing`, `show_timestamp` options | `cell_mappings.py:250-251` | `edac6a3` | never read (A-R70, A-R93) | pickaxe (RAN (me)) |
| `EXCEL_SAFETY` (retries, prefixes, minimum size, save options) | `utils/constants.py:234-245`; stored `core/excel_handler.py:55` | `edac6a3` | stored, never read (A-R107) | pickaxe |
| "Don't quit if Excel was already open" | `core/excel_handler.py:56, 241-246`; `__del__` `:58-64` | `edac6a3` | `self.excel_app` is never set, so `Quit()` always runs (A-R97, gap A-G54) | READ |
| "Critical" file types check | `core/file_parser.py:338-371` (`validate_results`) | `edac6a3` | never called (A-R30) [GST RULE] ⚠ | pickaxe |
| Clean-up of half-created folders after a client fails | `core/file_organizer.py:381-407` (`cleanup_on_error`, `_is_folder_empty`); `get_summary` `:409-426` | `edac6a3` | never called | pickaxe |
| Recent folders list | `gui/utils/cache_manager.py:32-43`; `get_cached_value` / `set_cached_value` `:50-56` | `edac6a3` | written, never read (A-R117) | pickaxe |
| `logs/` and `temp/` folders | `main.py:35-43` | `edac6a3` | created at every start, never used (A-R128) | READ |
| Client/state name warnings | `utils/helpers.py:428-449`; stored `core/file_parser.py:108-119` | `edac6a3` | computed, never shown (A-R18) | READ |
| Period / date / month / year parts of each file name [GST RULE] ⚠ | `core/file_parser.py:97-106` | `edac6a3` | captured, never read — no period check (A-R25, gap A-G39) | READ |
| Pattern `category` / `folder` labels [GST RULE] | `utils/constants.py:17-18, 25-26, 33-34, 41-42, 49-50, 57-58` | `edac6a3` | carried in every record, never read; the organiser uses its own map (A-R16) | READ |
| Unique-name, file-hash, relative-path and date-only helpers | `utils/helpers.py:451-470`, `:529-547`, `:557-574`, `:29-31` | `edac6a3` | never called (A-R49, A-R50, A-R63, A-R34) | pickaxe |
| Unused messages and fonts | `utils/constants.py:138-144, 214-227` | `edac6a3` | never read (A-R150, A-R153) | READ |
| A progress callback for `ProgressTracker` | `utils/helpers.py:476-523`; caller `gui/handlers/processing_handler.py:203` | `edac6a3` | never supplied | READ |
| Path-length guard ("the durable fix", `7347383` message) | — | named 26-07-2026 | never started (gap A-G41) | READ |
| Design note: "Continue to extraction" button in organise-only mode | `docs/superpowers/specs/2026-07-25-merge-extractor-design.md:78-81` | 25-07-2026 | not built; a log line instead (`gui/handlers/processing_handler.py:146-148`) | RAN (A-D): `git log -S'Continue to extraction'` → only `b276542` |
| Design note: move `logging.basicConfig` out of `extractor_main.py` | spec `:54-56` | 25-07-2026 | not done; side-stepped by importing `extractor_window` directly (`gui/tabs/extract_tab.py:8-10`) | READ |
| "Flattening" the notebook inside a notebook | `bfcb24a` message; `extractor_window.py:163-175` | 26-07-2026 | not done | READ |
| Step 4 notice for a missing pywin32 | `gui/tabs/extract_tab.py:67-69, 84-105` | `c463256` | cannot appear for the case it describes (RAN (H-A3)) | RAN |
| `offer_install(parent=…)` | `utils/dependency_installer.py:88` | `a5ad487` | never passed by `main.py:85-88` | READ |
| `ReportProcessor.cleanup()` | `power_query_extractor/core/report_processor.py:801-804` | `edac6a3` | called (`extractor_window.py:966`) but does nothing; `self.excel_instance` (`:48`) never read | READ |

#### 12E.5 Remnants: dead methods, dead constants, commented-out code and paste markers

**Dead from the first commit.** RAN (me): a pickaxe search for each of `EXCEL_EXTENSIONS`, `MAX_FILE_SIZE_MB`, `EXCEL_SIGNATURES`, `TEMPLATE_EXTENSIONS`, `EXCEL_SAFETY`, `get_date_only`, `get_unique_path`, `calculate_file_hash`, `get_relative_excel_path`, `validate_results`, `cleanup_on_error`, `get_summary`, `get_cached_value`, `set_cached_value`, `_write_section`, `_create_details_sheet`, `auto_refresh_if_missing`, `show_timestamp`, `recent_folders` in the program files returns only `edac6a3` — the number of mentions of each name has been identical in every commit since. None was ever wired in, and none was ever removed; they are remnants of the pre-git design.

**Dead or unreachable code added later (A only (after the split)):** `DarkModeManager._original_root_bg` (`gui/utils/dark_mode_manager.py:70, 83`, `aeb83cf`); `LIGHT_SURFACE` / `DARK_SURFACE` keys `card`, `border` (`gui/tabs/setup_tab.py:56, 62`, `169c7b5`); unused imports `os`, `Dict` (`gui/utils/setup_validation.py:8, 10`, `322717f`) and `DEFAULT_SUFFIX_PATTERN` (`power_query_extractor/core/report_processor.py:38`, `2aaa2b1`); the Step 4 notice (above); `find_excel_files` reduced to a test-only wrapper (`56aaad9`).

**Unreachable code from the first commit:** the `return {'success': True, 'error': None}` after a `try`/`except` that both return (`power_query_extractor/core/report_processor.py:383-386`); the `__del__` body that needs `self.excel_app` (`core/excel_handler.py:58-64`); a doubled header block with paste markers (`power_query_extractor/core/data_consolidator.py:73-83`).

**Commented-out code** (RAN (A-X-DEAD) grep, checked READ): only `power_query_extractor/config/cell_mappings.py:229-241` — a commented example sheet labelled "ADD MORE SHEETS HERE AS NEEDED", intended as documentation for the author (`edac6a3`) — and the two development tools in `requirements.txt:17-19` (`# black>=23.0.0`, `# pylint>=2.17.0`).

**Paste markers from AI chat answers** (READ; blame RAN (me)): all from `edac6a3` except one — `core/file_organizer.py:15` "# ADD THESE"; `core/excel_handler.py:39` "# ADD THESE TWO", `:449-451` "# NEW"; `utils/helpers.py:17` "# ADD THIS LINE", `:22` "# UPDATED:", `:577` "# NEW:"; `power_query_extractor/core/data_consolidator.py:75, 80` "# ... rest of the code" / "# ... rest of header formatting", `:222` "# ADD THIS LINE:", `:379, 402` "# ADD THESE LINES"; `power_query_extractor/core/report_processor.py:132` "# ADD THE NEW CODE HERE (notice the indentation - 3 levels):", `:183` "# ADD THIS: …", `:295` "# Rest of your code remains same..."; and from `6237597` (shared - before the split on 30-10-2025): `gui/handlers/processing_handler.py:86` "# CHANGE: …". They show that the pre-git code was assembled by pasting fragments of chat answers (HIGH). Commit `f7eb8f1`'s to-do item "remove dead code" (`PROJECT_AUDIT_ANNUALSTATEMENT.md`, Section 11 item 8) was not carried out in July: every dead item it named is still present (RAN (H-A1) grep).

#### 12E.6 Lessons that should not be re-learned in a rebuild (technical only)

1. **Application-wide event handling breaks other screens.** Twice (`322717f` → `d64d3a9`; `4c42758` → `bf2f1dd`) an "unbind all" or an Enter/Leave toggle killed scrolling elsewhere. The fix that held is one additive handler per scroll area that ignores events outside its area, enforced by `tests/test_scroll_bindings.py`.
2. **Optional imports must be optional at every level.** An import that is harmless in a separate program becomes fatal when another window imports it (`bfcb24a` → `c463256`).
3. **A fallback that silently damages the output is worse than stopping** — the reversal `c463256`/`f363946` → `be6418c`, in the author's own words "Producing damaged reports is worse than refusing to start" (`utils/requirements_check.py:52-53`). In A the protection is still only at start-up (12D, `be6418c`).
4. **Capture-once theming cannot follow colours the program changes later** (`aeb83cf` → `c89c1d4` → `169c7b5`); the per-tab patches do not cover the other tabs (A-G64).
5. **Names that sort as text are not dates** (`c1583d6`), and "date modified" is not "created" either (A-G45).
6. **Fixed waits do not detect completion** (pre-git tuning comments, `bd176bf`); Standard 12's refresh method (`rebuild_pipeline_FINAL.md:860-867`) was never tried in A's history (RAN (me) pickaxe: the texts `RefreshAll`, `CalculateUntilAsyncQueriesDone` and `DispatchEx` occur in no `.py` file at any of the 46 commits). Whether it was tried before git is a question for the author.
7. **Changing a shared rule in one place only leaves a second copy behind**: the "Private"/"Limited" rule exists twice (`utils/helpers.py:92-93`, `:624-625`) and only one copy got the separator step (`4219ffb`); the default refresh suffix is typed three times.

#### 12E.7 Where the older documents disagree with this history

| Document and place | Claim | What the history shows |
|---|---|---|
| `PROJECT_AUDIT_ANNUALSTATEMENT.md:1145` (its 12E item 1) and `:8, 18, 678, 1065-1091` | A PyQt5 migration on branch `copilot-worktree-2026-03-04T11-52-56` was "never merged … subsequently abandoned" | Those commits (`c258047`, `a62d3cc`, `c21f788`, 07-02-2026) are project B's `main` line, which continued (`7e09e5c`, `b0e7d0e`) — not an abandoned branch of A (RAN (H-A1) `git log 47a1492..B/main`). "Abandoned" describes only A's choice not to take them. |
| same, `:1151` (12E item 4) | "Only `save_options` is referenced (by `excel_handler.py:54` …)" | `excel_handler.py` stores the whole `EXCEL_SAFETY` dictionary (`core/excel_handler.py:55` today) and no key — `save_options` included — is ever read (RAN (A-D), grep) — never true. |
| same, `:96-112` (its dead-code table, Section 2) | 13 rows: 7 "dead" methods and constants (one of them, `extract_filename_without_extension`, its own text says IS called), a dead dictionary key, a stale reference to `gui/app.py`, four stale "ADD THIS" comments [range and count corrected during verification — originally "`:96-108` … lists 7 dead items"] | misses `cleanup_on_error`, `_is_folder_empty`, `get_summary`, three main-window delegators (`gui/main_window.py:244-245, 256-257, 275-276`), `UIHelpers.update_status`, `get_date_only`, `get_unique_path`, `get_relative_excel_path` — already dead on 05-03-2026 (RAN (A-X-DEAD) `git grep` at `f7eb8f1`); `validate_results` and `EXCEL_EXTENSIONS` (dead since `edac6a3`) are not mentioned anywhere in that document (RAN (me) grep). |
| `docs/superpowers/specs/2026-07-25-merge-extractor-design.md:78-81, 54-56` | "Continue to extraction" button; logging moved out of `extractor_main.py` | neither built (12E.4). |
| `README.md:185-187` | long category folder names | replaced in `7347383`. |

### 12F. Dependency Evolution

A *dependency* is a package the program needs that is not part of Python itself (installed with *pip*, Python's package installer, usually from a list in `requirements.txt`). This sub-section follows `requirements.txt` through the history, correlates each change with the commit that needed it, and compares what is declared with what the code actually imports.

#### 12F.1 `requirements.txt`, version by version

The file has three versions in 46 commits — created in `edac6a3` and changed twice, in `6c1e1c6` and `064c634` (RAN (me) `git show <commit>:requirements.txt`; re-run during verification with `git log --numstat -- requirements.txt`) [wording corrected during verification — originally "changed three times"]:

`edac6a3` (31-08-2025, shared - before the split on 30-10-2025) — 13 lines:
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

`6c1e1c6` (23-07-2026, A only (after the split)) — `pytest>=7.0.0` moved out of the commented block into an active entry "# Testing (run `pytest` from the repo root)", together with the first test suite.

`064c634` (23-07-2026, A only (after the split)) — added, between openpyxl and lxml:
```text
# Required on Windows for Excel COM automation / Power Query refresh.
# power_query_extractor and excel_handler import pythoncom/win32com/win32gui.
pywin32>=305; sys_platform == 'win32'
```
(`; sys_platform == 'win32'` is an *environment marker*: "install this only on Windows".) The file at the snapshot is exactly the `064c634` version (`requirements.txt:1-19`). Every entry is a minimum version (`>=`), never an exact one, at every commit ⚠.

#### 12F.2 Each dependency's history

| Package | Declared | Changed | Imported by the code | Why it was added or changed (correlated commit) |
|---|---|---|---|---|
| **openpyxl** (reads and writes `.xlsx` without Excel) | `>=3.1.2` since `edac6a3` ⚠ | never | `core/excel_handler.py:31-33` and `power_query_extractor/core/data_consolidator.py:8-10`, unconditionally, at every commit (RAN (me) `deps_by_commit.py`) | Core of the summary and consolidated workbooks and of the report fallback. Checked at start-up as essential since `f363946` (`utils/requirements_check.py:38-43`). |
| **lxml** (a fast XML library) | `>=4.9.0` since `edac6a3` ("Optional but recommended for better performance") ⚠ | never | **never imported by A** (RAN (me)) | openpyxl uses it by itself when lxml 3.3.1 or newer is installed and the environment variable `OPENPYXL_LXML` is either not set or set to exactly "True" (READ of the installed openpyxl 3.1.5, `openpyxl/xml/__init__.py:8-26`; re-read during verification). Not checked at start-up; installed anyway by the install offer (A-R125). |
| **pywin32** (Windows' COM bridge: modules `pythoncom`, `win32com`, `win32api`, `win32con`, `win32gui`) | **only from `064c634` (23-07-2026)** as `>=305; sys_platform == 'win32'` ⚠ | never after | since `edac6a3`: `core/excel_handler.py:21-28` (Windows only, inside a guard); `utils/helpers.py:584` (inside a function, guarded); `power_query_extractor/core/report_processor.py:20-36` (unconditional until `c463256`, guarded since) and `:762` (local re-import) | Used for 11 months without being declared. The March audit (`f7eb8f1`) listed it as critical gap 5 and proposed the exact line; `064c634` copied it character for character (RAN (H-A1)). Its importance then rose in two steps: `bfcb24a` made the main window load it at start-up (fatal without it, fixed `c463256`); `be6418c` made it essential on Windows (`utils/requirements_check.py:44-60`). |
| **pytest** (test runner) | commented in `edac6a3`; active `>=7.0.0` since `6c1e1c6` ⚠ | never after | tests only (`tests/conftest.py:11` and 14 test files; RAN (me)) | Added with the first tests. Side effect since `a5ad487`: accepting the start-up install offer installs the WHOLE file, so pytest (and lxml) land in the user's Python (`main.py:84-88`, `utils/dependency_installer.py:133-134`). |
| **black**, **pylint** (code formatter, code checker) | commented out since `edac6a3` | never | never | Development tools that were never used. |
| **tkinter** (Python's built-in window toolkit) | not listed — part of the standard library | — | every window file since `edac6a3`; `utils/dependency_installer.py:96-97` is the only guarded import | Not checked at start-up (`utils/requirements_check.py:37-61` knows only openpyxl and pywin32). Some Python distributions ship it separately. Whether the python.org "embeddable package" that Standard 13 uses for the shared runtime (`rebuild_pipeline_FINAL.md:913-914`) contains Tk is not sure — my recollection of Python's Windows documentation is that it does not (MEDIUM, not checked here); how to check: list the embeddable zip for `_tkinter.pyd` and a `tcl` folder. If it does not, a Tkinter program cannot run on that runtime as it stands. |
| **et_xmlfile** | not listed | — | not imported by A | Installed automatically by pip as openpyxl's own dependency (RAN (START notes): venv listing). |

**Python version.** "Python 3.7 or higher required" (`requirements.txt:2`) and "3.7 or higher (3.8+ recommended)" (`README.md:51`) have not changed since `edac6a3`. Evidence of the versions actually used: compiled files for Python 3.12 and 3.13 in the first commit (2025, RAN (H-S)); Python 3.14 and the free-threaded build `python3.14t` on the author's PC in July 2026 (`c463256` message: `py` → `python3.14t.exe` without pywin32, `python` → `python.exe` with it). The current package versions in this audit's environment declare minimums of Python 3.8 (openpyxl 3.1.5) and 3.10 (pytest 9.1.1) (RAN (START notes), package metadata), so the stated 3.7 cannot be met with today's newest packages ⚠. No syntax newer than Python 3.7 was found in the program code (RAN (A-D notes): parsed with the 3.7 grammar; MEDIUM — not exhaustive).

**Version upgrades: none, ever.** No minimum version number was raised in any commit; `openpyxl>=3.1.2` and `lxml>=4.9.0` are as in the first commit, `pywin32>=305` and `pytest>=7.0.0` as when added. Upgrades happen only implicitly, because `>=` lets pip install whatever is newest at install time (this environment has openpyxl 3.1.5, RAN (me)) ⚠.

#### 12F.3 What the code imports versus what is declared (at the snapshot)

RAN (me), `deps_by_commit.py` (every `import` and `from … import` in every `.py` file, compared with Python's own list of standard-library modules):

| Imported module | Package | Declared in `requirements.txt`? | Import sites at `bf2f1dd` | Guarded? |
|---|---|---|---|---|
| `openpyxl` (+ `.styles`, `.utils`) | openpyxl | yes, since `edac6a3` | `core/excel_handler.py:31-33`; `power_query_extractor/core/data_consolidator.py:8-10` | no (start-up check since `f363946`) |
| `pythoncom`, `win32com.client` | pywin32 | yes, since `064c634` (Windows only) | `core/excel_handler.py:23-24`; `power_query_extractor/core/report_processor.py:21-22` | yes (Windows check + `ImportError`; `ImportError` only) |
| `win32api` | pywin32 | same | `report_processor.py:23, 762`; `utils/helpers.py:584` | yes |
| `win32con`, `win32gui` | pywin32 | same | `report_processor.py:24-25` | yes (since `c463256`) |
| `pytest` | pytest | yes, since `6c1e1c6` | `tests/conftest.py:11` and 14 test files | — |
| `lxml` | lxml | yes | none | — |
| `tkinter` (`ttk`, `messagebox`, `filedialog`, `scrolledtext`) | standard library | not applicable | all `gui/` files; `power_query_extractor/gui/extractor_window.py`; `launch_extractor.py`; `utils/dependency_installer.py` | only in `utils/dependency_installer.py:95-99` |

Third-party modules imported but never declared at the snapshot: **none**. Over the history: **pywin32 was used but undeclared from `edac6a3` (31-08-2025) to `064c634` (23-07-2026)** — the documentation gap the method asks to flag; closed in A (still open in project B, whose `requirements.txt` has pywin32 only as a comment — RAN (H-A1)).

#### 12F.4 Standard-library modules added over time

RAN (me), first commit in which each standard-library module is imported (program files; tests separately):

| Commit | Label | Program | Tests | Why |
|---|---|---|---|---|
| `edac6a3` | shared - before the split on 30-10-2025 | `collections`, `datetime`, `hashlib`, `io`, `json`, `logging`, `os`, `pathlib`, `platform`, `re`, `shutil`, `sys`, `tempfile`, `threading`, `time`, `tkinter`, `traceback`, `typing`, `warnings` | — | the first program |
| `a603c02` | A only (after the split) | `stat` (`S_ISREG`, `S_ISDIR`) | — | file type from one query |
| `6c1e1c6` | A only (after the split) | — | `os`, `sys` (+ pytest) | first tests |
| `4f84ba9` | A only (after the split) | `concurrent.futures` (`ThreadPoolExecutor`) | — | parallel signature reads |
| `322717f` | A only (after the split) | — | `tkinter` | Setup tab smoke tests |
| `a677670` | A only (after the split) | — | `json`, `pathlib`, `subprocess`, `textwrap` | standalone window tested in a separate process |
| `aeb83cf` | A only (after the split) | `weakref` | — | colours keyed by widget object |
| `f363946` | A only (after the split) | `importlib` (`importlib.util.find_spec`) | — | start-up check without loading |
| `a5ad487` | A only (after the split) | `subprocess` | `builtins` | pip and restart from inside the program |
| `c89c1d4` | A only (after the split) | — | `inspect` | a test that reads a function's source text |
| `bf2f1dd` | A only (after the split) | — | `re` | the scroll-binding guard test |

`concurrent.futures` brings a rule from Standard 13 into play ("with `concurrent.futures`, always call `result()`", `rebuild_pipeline_FINAL.md:1037-1038`): `list(pool.map(...))` (`utils/helpers.py:354-355`) does collect every result, and `_has_excel_signature` catches every error itself (`:281-287`), so no worker error can be lost (READ (H-A2), HIGH).

#### 12F.5 Internal dependencies that changed (which part of the program needs which)

| Commit | Change | Consequence |
|---|---|---|
| `ba2a753` (A only (after the split)) | `power_query_extractor/core/report_processor.py:39` imports the top-level `utils.excel_com` | The extractor package is no longer self-contained; it works because both entry points put the program folder on Python's search path (`main.py:32`, `launch_extractor.py:10`) (READ (H-A1), HIGH). |
| `2aaa2b1` | `extractor_window.py:15` and `report_processor.py:38` import `power_query_extractor/core/refresh_naming.py` | One rule, two users. |
| `322717f` | `gui/tabs/setup_tab.py` imports `gui/utils/setup_validation.py` and `gui/widgets/tooltip.py` | Logic testable without a window — but importing `gui.utils.setup_validation` still loads the whole window, the extractor and openpyxl (222 modules, RAN (A-X-DEAD)), because `gui/__init__.py` imports the main window. |
| `bfcb24a` | `gui/main_window.py:29` imports `gui/tabs/extract_tab.py`, which imports `extractor_window` (`extract_tab.py:25`) | The organiser now depends on the extractor at start-up; this pulled pywin32 into start-up until `c463256`. |
| `f363946`, `a5ad487` | `main.py:68-70, 82` import `utils/requirements_check.py` and `utils/dependency_installer.py` | The start-up check needs no window toolkit; the installer imports tkinter inside a guard. A second, code-level list of dependencies now exists (`utils/requirements_check.py:37-61`) besides `requirements.txt` — two places to keep in step. |
| `a5ad487` | the program itself runs `<this python> -m pip install -r requirements.txt` | A run-time dependency on pip and on internet access, and the program can change the Python it runs under (gap A-G59). |

#### 12F.6 Development tooling and the AI assistant

- `.claude/settings.local.json` (`bd176bf`, shared - before the split on 30-10-2025) allows the AI coding tool to run `python -m py_compile` without asking — the first trace of Claude Code in the history; no secrets (READ).
- `.gitignore` (`1cb878b`) lists `/mnt/user-data/`, the file area of a web-based AI chat, which suggests the file was generated there (MEDIUM).
- `pytest.ini` (`6c1e1c6`): `testpaths = tests`, no `pythonpath = .`; `tests/conftest.py:9` inserts the program folder into the search path instead.
- All 32 July commits carry "Co-Authored-By: Claude Opus 4.8 (1M context)" (RAN (me)); the March audit document names "Claude Opus 4.6" as its author (`PROJECT_AUDIT_ANNUALSTATEMENT.md` header, READ (H-A1)).

#### 12F.7 Comparison with the author's Standards 12 and 13 (facts only; decisions are the author's)

| Standard says | A at the snapshot | Since |
|---|---|---|
| Each app's `requirements.txt` lists "the exact versions of the packages it imports (not pytest)" (`rebuild_pipeline_FINAL.md:937-938`) | minimum versions only; includes pytest and lxml (`requirements.txt:5-15`) ⚠ | `edac6a3` / `6c1e1c6` |
| `pytest.ini` with `pythonpath = .` and `testpaths = tests` (`:944-947`) | `testpaths = tests`; path inserted by `tests/conftest.py:9` | `6c1e1c6` |
| "Nothing is installed on users' PCs" (`:924`); packages go into the shared `_runtime` with `pip install --target …` (`:933-936`) | the program offers to pip-install `requirements.txt` into the Python it is running under (`main.py:75-95`) | `a5ad487` |
| Launcher starts `pythonw.exe -B -s` (`:962, 969-972`) | after installing, the program restarts itself with `[sys.executable] + sys.argv`, which drops `-B -s` (RAN (H-A3)); and `main.py:17` fails when there is no console (RAN (H-A3) simulation) | `a5ad487`; `edac6a3` |
| Desktop windows: PySide6 (`:833`) | Tkinter | `edac6a3` (before git) |
| Writing into or refreshing Power Query workbooks: Excel through COM with `DispatchEx`, never the user's own Excel (`:828, 851-856`) | COM through `GetActiveObject` / `Dispatch` (`core/excel_handler.py:123-128`; `report_processor.py:218, 536`) | `edac6a3` |
| openpyxl only for workbooks without Power Query; check for `DataMashup` before saving (`:827, 846-850`) | openpyxl remains the fallback writer for the templates; no `DataMashup` check anywhere (RAN (H-A3) grep); the start-up gate (`be6418c`) is the only protection | `edac6a3`; gate `be6418c` |

#### 12F.8 Where the older documents disagree with the dependency history

| Document and place | Claim | What the history shows |
|---|---|---|
| `README.md:104-108` ("Requirements.txt Contents", since `d674386` 29-10-2025) | `openpyxl>=3.0.0` and `pywin32>=305; sys_platform == 'win32'` | `openpyxl>=3.1.2` at every commit (never 3.0.0); pywin32 was not in the file until `064c634` (23-07-2026) — for nine months the README described a line that did not exist; lxml and pytest are not mentioned. |
| `README.md:275` | after `pip install pywin32` run `python Scripts/pywin32_postinstall.py -install` as admin | The install offer (`a5ad487`) runs only pip and then restarts; whether the post-install step is still needed cannot be checked here (Windows) — not sure; how to check: on a clean Windows PC, let the offer install pywin32 and start Step 3. |
| `PROJECT_AUDIT_ANNUALSTATEMENT.md:1153-1162` (its 12F) and `:643, 726` | "pywin32 NOT added to requirements.txt"; "MISSING from requirements.txt" | True on 05-03-2026; **no longer true** since `064c634`. Its row "c258047 (unmerged) PyQt5 added to requirements.txt" (`:1159`) concerns project B's history. |
| `requirements.txt:7-8` (comment) | "power_query_extractor and excel_handler import pythoncom/win32com/win32gui" | `core/excel_handler.py` imports only `win32com.client` and `pythoncom` (`:23-24`); `win32gui` is imported only by `report_processor.py:25`; and all those imports are guarded since `c463256`, although the comment was written while the extractor's were unconditional (READ). |
| `gui/tabs/extract_tab.py:21-23, 101-105`; `power_query_extractor/core/report_processor.py:11-13` | pywin32 "is optional - the organizer itself works without it"; "Steps 1-3 are unaffected … works without pywin32" | Since `be6418c` pywin32 is essential on Windows and the program refuses to start without it (`utils/requirements_check.py:44-60`, `main.py:97-102`). |

#### 12F.9 Questions for the author (dependencies)

1. Should the program's `requirements.txt` pin exact versions and leave out pytest (Standard 13, `rebuild_pipeline_FINAL.md:937-938`)? Today it lists minimums ⚠ and the install offer would put pytest on users' PCs.
2. Should the program ever run pip itself (`a5ad487`), given Standard 13's "nothing is installed on users' PCs"?
3. Which Python will users run — the normal 3.14 build, or the free-threaded `python3.14t` mentioned in `c463256` (for which pywin32 availability is unknown to me)? And will the shared runtime include tkinter?
4. Is the README's pywin32 post-install step still part of your set-up routine?

#### 12F.10 Limits of the evidence in 12B–12F

- Everything that needs Windows, Excel or pywin32 (COM sessions, key presses, pip installing pywin32, the restart making it importable, `pythonw` beyond the simulated missing console) was traced by reading only; where another reader simulated it, that is marked RAN (H-A3).
- Claims in commit messages that rest on measurements outside the repository — "11 min -> 41s", "3030 files", "220 widgets … no drift", "byte-identical characterization snapshot", "mutation-checked" — could not be verified here and are reported as the messages' claims.
- `gst_organizer.log` was not opened (described by kind only, from H-S's counts of message shapes); the compiled `.pyc` files were analysed only through their header fields and identifier strings by other readers, never decompiled.
- Blame and pickaxe identify the commits that touched a line or a text; that a change was "substantive" or "cosmetic" was decided by reading each diff.
