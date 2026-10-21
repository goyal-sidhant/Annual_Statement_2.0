> **About this part.** It covers the twelve commits that project A (goyal-sidhant/AnnualStatement_oldv1) and project B (goyal-sidhant/AnnualStatement_oldv2) have in common, from the first upload on 31-08-2025 to the split point 47a1492 on 30-10-2025. The same text appears unchanged in both audits. All twelve commits are on branch `main` in both repositories and every one is labelled "shared - before the split on 30-10-2025". Each audit continues with its own Phase 4.

**How to read the evidence in this part**

- Code is cited as `path:line@commit`, meaning "that line number in that commit's version of the file", for example `core/file_parser.py:254@edac6a3`. Files that did not change between edac6a3 and 47a1492 have the same line numbers at both. A citation without `@commit` means: inside the description of the first commit (sub-sections A to Y under edac6a3) the line at edac6a3; inside a later commit's entry the line at that commit (for bd176bf the files at bd176bf and 47a1492 are identical); in "Where the two projects started" the line at 47a1492. `:N` alone means "line N of the file named just before".
- Feature and rule IDs (A-F…, A-R…, B-F…, B-R…) are not used in this part: the same text serves both audits, and those IDs describe each project's latest snapshot, while this part describes code from before the split. Sub-sections of the first commit are referred to by their letters (A to Y).
- **READ** = I read the code or the diff (`git show <hash>`, and `git show -m --first-parent 9c56a1c` for the merge). **RAN** = I executed the code on made-up data, always on my own extracted copies under `/root/merge-scratch/run/SHARED-12A/` (Linux, Python 3.14, Tk under a virtual display, run with `python -B` so nothing is written into the copies). Anything that needs Excel or Windows was traced by reading only. Inferences carry **HIGH / MEDIUM / LOW** confidence.
- **[GST RULE]** marks code that decides a tax figure, which return, sheet or cell a figure belongs to, a state code, a period, rounding or display of figures. For these I only say what the code does; whether it is right is the author's call. **⚠** marks values that can change over time (state lists, return table numbers, template sheet names and cell positions) so the author can check them.
- Times are the commit times git recorded, in India Standard Time (+05:30). Times taken from compiled-file headers are in UTC and are labelled "UTC".
- No client name, GSTIN, PAN, real file name, folder path from the author's computer or figure appears in this part. Names such as `<Client A>`, "Alpha Traders" or "Beta Stores" are made up for the test runs.
- **Independent re-check (21-10-2026).** Every citation in this part was re-opened in the code at the cited commit, every commit count, date and line count was re-read from git, and the most important RAN claims were re-run with separately written scripts on fresh extracted copies under `/root/merge-scratch/run/VERIFY-A-shared12A/` (`v_dryrun.py`, `v_consolidator.py`, `v_consolidator_empty3b.py`, `v_suffix.py`, `v_rowtags.py`, `v_misc.py`, `v_window.py`, `v_hyphen.py`, `test_wheel_leave.py`, `pyc_hdr.py`, `pyc_paths.py`, and a no-console start of `main.py`). Corrections made during that check are written into the text where they belong.

**Words used in this part** (plain meanings of the technical terms used below)

| Term | Plain meaning |
|---|---|
| commit | one saved change in the git history |
| diff | the exact lines a commit added and removed |
| merge commit | a commit that joins two lines of work into one |
| split point | 47a1492, the last commit both projects share |
| regex (regular expression) | a text pattern used to recognise file names, for example "starts with `GSTR3B-`" |
| GUI | the program's windows, buttons and lists |
| Tkinter (Tk) | the window toolkit that comes with Python; every shared commit uses it |
| thread / worker thread | a second line of work running at the same time, so the window stays responsive; Tk windows should only be changed from the main thread |
| callback | a small piece of code handed to another part of the program to be run later |
| lambda / late binding | a lambda is a tiny inline function; "late binding" means it reads its variables when it finally runs, not when it was created, so inside a loop it can show the last value for every line |
| COM / pywin32 | COM is the Windows mechanism that lets a program drive Microsoft Excel; pywin32 is the Python package that provides it (`win32com`, `pythoncom`, `win32gui` ...) |
| openpyxl | a Python package that reads and writes .xlsx files without Excel |
| Power Query | Excel's built-in data-import engine; the ITC and Sales templates use it to pull data from the organised files |
| SendKeys | simulating key presses into whatever window is in front |
| cache file / JSON | a small settings file that remembers last-used folders; JSON is the text format it is written in |
| glob | a wildcard file search such as `ITC_Report_*.xlsx` |
| modification time (mtime) | a file's "last changed" date and time |
| `.pyc` / `__pycache__` | compiled copies of Python files that Python writes automatically into `__pycache__` folders |
| CRLF | Windows line endings (two characters per line end); git shows LF (one character) |
| Treeview | the table-like client list on the Validation screen |
| dead code | code that is never called |

**The twelve shared commits at a glance** (READ: `git log`, `git show --shortstat`, `git show -m --first-parent --shortstat 9c56a1c`)

| # | Commit | Date and time | Phase | Message (shortened) | Files | Lines added / removed | Kind |
|---|---|---|---|---|---|---|---|
| 1 | edac6a3 | 31-08-2025 22:11 | 1 | Initial commit | 110 | +12,916 / −0 | first upload, with an accidental run log and compiled files |
| 2 | cbfa4a5 | 31-08-2025 22:14 | 1 | Delete gst_organizer.log | 1 | +0 / −5,689 | clean-up, made on GitHub's web page |
| 3 | 6237597 | 29-10-2025 17:21 | 2 | Fix critical import and threading bugs in GST Organizer | 4 | +59 / −45 | claimed fixes; one rule change; one new bug |
| 4 | 9c56a1c | 29-10-2025 17:21 | 2 | Merge branch 'main' of https://github.com/goyal-sidhant/AnnualStatement | 1 (against its first parent) | +0 / −5,689 | merge that brought in commit 2 |
| 5 | d674386 | 29-10-2025 21:15 | 2 | Update README.md for clarity and formatting improvements | 1 | +239 / −135 | documentation |
| 6 | 5ea77c2 | 29-10-2025 21:20 | 2 | Remove shebang line from main.py | 1 | +0 / −1 | tidy-up |
| 7 | 1cb878b | 29-10-2025 21:43 | 2 | Add .gitignore file to exclude unnecessary files and directories | 1 | +79 / −0 | tooling |
| 8 | 01862a0 | 29-10-2025 21:46 | 2 | Remove __pycache__ directories from Git tracking | 70 | binary files only | clean-up |
| 9 | 1c48722 | 29-10-2025 22:18 | 3 | Remove redundant text color handling in client tree display | 1 | +0 / −6 | bug fix (the message calls it a clean-up) |
| 10 | 4c42758 | 29-10-2025 22:25 | 3 | Add mouse wheel scrolling support for main and log text areas | 1 | +30 / −0 | usability |
| 11 | bd176bf | 30-10-2025 09:04 | 3 | Enhance report processing UI with configurable wait time and file suffix options; add ITC/Sales report selection functionality | 3 | +412 / −162 | feature, plus an AI-tool settings file |
| 12 | 47a1492 | 30-10-2025 14:38 | 3 | Add new ITC report mappings for Reco Diff, Table 6J Match, Table 8D Limits, IMS Reco, Extra In Purchase, and Extra In GSTR2B sheets | 2 | +126 / −63 | configuration [GST RULE] and clean-up |

Shape of the history (READ: `git log --format='%h %p'`):

```
edac6a3 ──► cbfa4a5 ──────────────────┐          31-08-2025; cbfa4a5 made on GitHub's web page
   │                                  ▼
   └──► 6237597 ────────────────► 9c56a1c        29-10-2025; 6237597 made on the author's computer
                                      │
        9c56a1c ──► d674386 ──► 5ea77c2 ──► 1cb878b ──► 01862a0 ──► 1c48722 ──► 4c42758
        4c42758 ──► bd176bf ──► 47a1492           30-10-2025; 47a1492 = split point
```

**Scratch scripts I wrote and ran for this part** (all under `/root/merge-scratch/run/SHARED-12A/scripts/`, each on extracted copies `t_<commit>/` of the code, never on the original folders): `test_window.py` (organiser window builds), `test_dryrun.py` (Dry Run preview text), `test_row_tags.py` (row colours per client), `test_tree_tags_toggle.py` (row colours after switching dark mode on and off), `test_tree_values.py` (what the client list returns for number-like text; state-code helper), `test_misc.py` (disabled button with a click binding; "latest" version-folder choice), `test_consolidator.py` (consolidated extraction workbook with labelled values), `test_suffix.py` (refreshed-file suffix and Skip Refresh, Windows parts stubbed out), `pyc_headers.py`, `pyc_sizecheck.py` and `pyc_pathcount.py` (compiled-file header fields, a size comparison with the committed sources, and a count of distinct embedded folder locations; none prints any path). The run log was examined only with counting commands that print logger names, levels, dates and counts.

---

#### Phase 1: First upload of the finished v3.0 program (31-08-2025)

On the evening of 31-08-2025 a program that had already been built and used outside git for about ten weeks was uploaded in a single commit, together with its run log and its compiled leftovers. Three minutes later the run log was deleted on GitHub's web page. No further commit was made for the next 59 days.

##### edac6a3 - 31-08-2025 22:11 - shared - before the split on 30-10-2025

**Commit message:** "Initial commit" (no body). Author and committer: goyal-sidhant. No parent (it is the root of both histories). (READ)

**What changed - 110 files, +12,916 lines, nothing removed** (READ: `git show --numstat edac6a3`; 6,954 Python lines + 260 README + 13 requirements + 5,689 log = 12,916)

| Group | Files | Lines | Notes |
|---|---|---|---|
| Python source | 37 | 6,954 | listed in the next table |
| `README.md` | 1 | 260 | user guide, written for an older layout (see W below) |
| `requirements.txt` | 1 | 13 | `openpyxl>=3.1.2`, `lxml>=4.9.0`, three commented development tools |
| `gst_organizer.log` | 1 | 5,689 | the organiser's run log from real use (described by kind only, see X below) |
| compiled `__pycache__/*.pyc` | 70 | binary | 35 modules × 2 Python versions (see Y below) |

| Python file | Lines | Role |
|---|---|---|
| `main.py` | 83 | starts the organiser |
| `launch_extractor.py` | 30 | starts the separate Power Query Extractor |
| `core/__init__.py` | 13 | re-exports FileParser, FileOrganizer, ExcelHandler; `__version__ = '3.0.0'` |
| `core/file_parser.py` | 373 | scans a folder, recognises GST file names, groups files per client, checks completeness |
| `core/file_organizer.py` | 409 | creates output folders, copies files under cleaned names, writes a per-client text report |
| `core/excel_handler.py` | 661 | creates the ITC and Sales reports from templates (through Excel or openpyxl) and the summary workbook |
| `utils/__init__.py` | 38 | re-exports constants and helpers; `__version__ = '3.0.0'` |
| `utils/constants.py` | 275 | file-name patterns, folder names, template cell map, window settings, modes, messages, Excel safety values, state-name map |
| `utils/helpers.py` | 537 | timestamps, paths, file-name cleaning, copy and backup, Excel signature check, state code, client-state key |
| `gui/__init__.py` | 11 | re-exports GSTOrganizerApp |
| `gui/main_window.py` | 365 | main window; connects screens, button logic, settings cache, dark mode |
| `gui/tabs/__init__.py`, `setup_tab.py`, `validation_tab.py`, `processing_tab.py` | 8 / 377 / 247 / 117 | the three screens |
| `gui/handlers/__init__.py`, `file_handler.py`, `client_handler.py`, `processing_handler.py` | 8 / 151 / 259 / 447 | button logic: browse and scan, client selection, processing thread |
| `gui/utils/__init__.py`, `cache_manager.py`, `dark_mode_manager.py`, `status_bar.py`, `ui_helpers.py` | 9 / 56 / 212 / 37 / 63 | settings cache, theme switch, status bar, log helpers |
| `gui/widgets/__init__.py`, `collapsible_frame.py`, `title_bar.py` | 7 / 41 / 51 | small reusable window parts |
| `power_query_extractor/__init__.py` | 3 | `__version__ = "1.0.0"` |
| `power_query_extractor/extractor_main.py` | 79 | extractor start-up, its own log file, cache lookup |
| `power_query_extractor/config/__init__.py`, `core/__init__.py`, `gui/__init__.py` | 1 each | one-line descriptions only |
| `power_query_extractor/config/cell_mappings.py` | 152 | which report cells the extractor reads [GST RULE] |
| `power_query_extractor/config/cell_mappings copy.py` | 63 | an older format of the same mapping; the space in the file name means Python can never import it |
| `power_query_extractor/core/report_processor.py` | 736 | opens each report in Excel, refreshes Power Query by key presses, checks the refresh, reads cells |
| `power_query_extractor/core/data_consolidator.py` | 432 | writes the consolidated extraction workbook |
| `power_query_extractor/gui/extractor_window.py` | 601 | the extractor's window (Setup and Processing tabs) |

(Line counts are git's. A file whose last line has no line ending shows one line fewer in `wc -l`.)

**The original program in detail - where both projects started**

**A. Two separate programs**

- **The organiser** is started with `python main.py`. In order (READ):
  1. `main.py:18@edac6a3` replaces the console output (`sys.stdout`) with a UTF-8 version of it, so emoji in messages can be printed.
  2. `main.py:21-28` sends all log messages of level INFO and above both to a file `gst_organizer.log` and to the console. The file name has no folder, so the file is created in whatever folder the program is started from, and it is appended to on every start.
  3. `setup_environment()` (`main.py:30-44`) puts the program folder on Python's search path, creates the folders `core`, `utils`, `gui`, `logs` and `temp` inside the program folder if missing, and writes a one-line `__init__.py` into `core`, `utils` and `gui` if missing. `logs` and `temp` are never used afterwards.
  4. `main()` (`main.py:46-80`) imports the main window, prints a banner and runs the window. On a missing module it prints the list of expected files - which still includes **`gui/app.py`** (`main.py:72`), a file that is not in the commit (see Y) - and waits for Enter (`main.py:73`). On any other error it logs the details, prints them and waits for Enter.
  5. The first line `# !/usr/bin/env python3` (`main.py:1`) has a space after `#`, so it was never a working "shebang" (the line that tells Linux/macOS which interpreter to use) (READ, HIGH).
  6. The header says "Author: Advanced AI Assistant" and "Version: 3.0 (Production Ready)" (`main.py:6-7`).
  - RAN (`test_window.py`, `t_edac6a3`): the main window builds with title "GST File Organizer v3.0" and the tabs "📁 Step 1: Setup", "✅ Step 2: Validation", "🚀 Step 3: Processing".
  - Console dependence (READ, HIGH): the program prints banners and waits for Enter on errors (`main.py:54-56,63-79`), and step 1 assumes a console exists. Started without a console (for example with `pythonw.exe` on Windows, where `sys.stdout` is empty), line 18 fails with an AttributeError before anything is logged.
- **The Power Query Extractor** is a second, separate program started with `python launch_extractor.py` "after GST Organizer processing" (`launch_extractor.py:2-3`). The launcher puts its own folder on the search path (`:10`), imports and runs `PowerQueryExtractor` (`:14-18`), and on any error shows a Tk error box and re-raises it (`:20-27`). Importing `power_query_extractor/extractor_main.py` sets up its own log file `pq_extractor.log` (again relative to the start folder) plus console output (`extractor_main.py:12-19`), but only if no logging was set up before. The extractor's window class is `PowerQueryExtractorApp` (`power_query_extractor/gui/extractor_window.py:18`), title "Power Query Extractor", 1000×700 (`:43-45`).
  - The extractor is **Windows-only with pywin32 installed**: `power_query_extractor/core/report_processor.py:11-15` imports `pythoncom`, `win32com.client`, `win32api`, `win32con` and `win32gui` at the top with no fallback, so without them the program cannot even start. RAN: importing `power_query_extractor.extractor_main` from my copy stops with `ModuleNotFoundError: No module named 'pythoncom'`. The organiser, by contrast, imports pywin32 only on Windows and falls back to openpyxl (`core/excel_handler.py:19-33`). The two programs handle the same missing package in opposite ways (READ, HIGH).
- The two programs never call each other. They communicate only through the folders and file names the organiser creates (HIGH, READ of both).

**B. The organiser's three screens** (READ)

| Screen | What is on it | Code |
|---|---|---|
| Title bar | "🏢 GST FILE ORGANIZER & REPORT GENERATOR", subtitle, a "🌙 Dark Mode" checkbox placed at a fixed position | `gui/widgets/title_bar.py:17-51` |
| 📁 Step 1: Setup | welcome card; Source Folder (entry + Browse); Excel Templates: ITC Report Template and Sales Report Template (entry + Browse, file filter `*.xltx *.xlsx`); Target Folder (entry + Browse, marked "⭐ IMPORTANT!"); Processing Mode radio buttons (🆕 Fresh Run, 🔄 Re-Run, ▶️ Resume); checkbox "Include client name in brackets for Level 4 folders"; note "⚠️ Names >10 chars will show warning / ℹ️ This overrides individual client settings if checked"; "Client folder name max length" spin box 15-100 with label "(default: 35, range: 15-100)"; buttons 🔍 SCAN FILES and 🔄 RE-SCAN; a help column with five instructions and the six expected name patterns | `gui/tabs/setup_tab.py:19-377`; filters `gui/handlers/file_handler.py:31,42` |
| ✅ Step 2: Validation | collapsible instructions; "📊 SCAN SUMMARY" text box; buttons 🧪 DRY RUN, 📋 EXPORT LIST, 🚀 START; "👥 CLIENT SELECTION" with Select All / Clear All / Complete Only; the client list (a Treeview) with columns ✓, Client Name, State, Status, Files, Missing, Extra, "📁 Name" (include-name setting); keyboard ↑ ↓ and SPACE | `gui/tabs/validation_tab.py:18-247` |
| 🚀 Step 3: Processing | progress bar, progress label, current-operation label, 🚀 START PROCESSING and ⏹ STOP buttons, a dark "📝 PROCESSING LOG" with coloured lines | `gui/tabs/processing_tab.py:17-117` |
| Status bar | message on the left, "v3.0 \| Production Ready" on the right | `gui/utils/status_bar.py:15-37` |

- Almost every action button (Browse, Scan, Re-scan, Dry Run, Export, Start, Stop, Select All, Clear All, Complete Only) is wired with a mouse-click binding `bind('<Button-1>', ...)` instead of Tk's normal `command=` option (for example `setup_tab.py:115,161,199,306,318`, `validation_tab.py:118,130,142,190`, `processing_tab.py:74,87`). Comments call this "working mouse events" and "proper event binding" (`setup_tab.py:20,103`), which hints that `command=` had given trouble in an earlier version (LOW). Consequence: a button that is "disabled" still reacts to a click (RAN, `test_misc.py`: a disabled Tk button with such a binding still ran its handler), and pressing Space on a focused button does nothing, because there is no `command` to run (READ, MEDIUM). The extractor, written separately, uses `command=` throughout (`extractor_window.py:133,143,168,178,223`) - an inconsistency between the two programs (READ, HIGH).
- Mouse-wheel scrolling of the Setup screen is bound for the whole window with `bind_all("<MouseWheel>")` (`setup_tab.py:63-66`), the Windows/macOS wheel event.

**C. What SCAN FILES does** (READ; `gui/handlers/file_handler.py:57-95` → `core/file_parser.py:128-179`)

1. Checks that a source folder is entered and exists (`file_handler.py:120-130`).
2. `find_excel_files()` (`utils/helpers.py:249-266`) lists only the top level of the folder (no sub-folders) with the wildcards `*.xlsx`, `*.xls`, `*.xlsm`, and **silently drops** every file that fails `validate_excel_file()` (`helpers.py:210-247`): smaller than 1,024 bytes, or extension not one of .xlsx/.xls/.xlsm/.xltx/.xltm, or first bytes not the ZIP signature `PK\x03\x04` or the old-Excel signature `\xd0\xcf\x11\xe0`. Such files appear nowhere - not in the client list and not in "variations" (the "Invalid or corrupted Excel file" branch at `file_parser.py:188-195` can only be reached if a file changes between the two checks) (READ, HIGH). The list is sorted by name (`helpers.py:262`).
3. Each file name is tried against six patterns in a fixed order (D below). A match that yields a client and a state goes into the client data under the key `<client as written>-<state code>` (`file_parser.py:228-252`); everything else becomes a "variation" with a suggested pattern (`file_parser.py:212-222,290-308`).
4. Completeness is computed per client (E below), the summary text (`gui/main_window.py:267-301`) and the client list are refreshed, and the window switches to the Validation screen (`file_handler.py:78-88`).
- The scan runs in the main thread, so the window does not respond while a large folder is scanned; the per-file status messages it sets are not repainted until the scan ends (READ, MEDIUM).

**D. File-name patterns the scanner recognises** - `utils/constants.py:13-62@edac6a3`, unchanged up to the split point. [GST RULE] ⚠ (they decide which return or report each file is treated as)

All patterns ignore upper/lower case, must match the whole name, and accept only `.xls` or `.xlsx` (`\.xlsx?$`), so an `.xlsm` file is found by the scan but always ends up as a variation. `[^-]+` means "one or more characters that are not a hyphen", so **a client name or a state name containing a hyphen can never be read** (READ, HIGH).

| Order | Key | Pattern (quoted exactly) | Treated as | Copied to | Parts captured |
|---|---|---|---|---|---|
| 1 | `GSTR-2B-Reco` | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | "GSTR-2B Reco", category ITC | ITC folder | client, state, period (the period may contain hyphens) |
| 2 | `ImsReco` | `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$` | "IMS Reco", category ITC | ITC folder | client, state, date (exactly 8 digits; described as DDMMYYYY) |
| 3 | `GSTR3B` | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | "GSTR-3B Export", category GSTR3B | GSTR-3B folder | client, state, month; an optional prefix such as `(2) ` (probably a duplicate-download number) is allowed only here |
| 4 | `Sales` | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | "Sales", category Sales | Sales folder | client, state, start month, end month |
| 5 | `SalesReco` | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | "Sales Reco", category Sales | Sales folder | client, state, period |
| 6 | `AnnualReport` | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$` | "Annual Report", category Annual | the Version folder itself | client, state, year |

- [GST RULE] The period, date, month and year parts are stored (`file_parser.py:97-98`) and **never used anywhere** (READ, HIGH: no reader in any file). The code does not check a tax period or financial year at all; apart from the IMS date being 8 digits, these parts are free text.
- Client and state parts are checked by `validate_client_name()` (`helpers.py:321-342`: 2-100 characters, none of `< > : " / \ | ? *`); problems are stored as "warnings" (`file_parser.py:100-111`) that are **never shown** (READ, HIGH).
- Suggestions for unrecognised names (`file_parser.py:290-308`) are chosen by keywords in this order: "gstr" and "2b" → GSTR-2B pattern; "ims" → IMS; "gstr" and "3b" → GSTR-3B; "sales" and "reco" → Sales Reco; "sales" → Sales; "annual" → Annual Report; otherwise "Check file naming convention".

**E. When a client counts as "Complete"** - `core/file_parser.py:254-288@edac6a3`, list `utils/constants.py:68-75`. [GST RULE] ⚠ (it decides which returns and files a client must have)

- Six file types are expected: GSTR-2B Reco, IMS Reco, GSTR-3B Export, Sales, Sales Reco, Annual Report.
- "Missing" = expected types with no file. "Extra" = any type other than GSTR-3B Export with two or more files ("GSTR-3B can have multiple files (monthly filings - unlimited)", `:268-270`).
- Status: "Complete" ✅ when nothing is missing and nothing is extra; "Missing N files" ⚠️ when anything is missing; otherwise "Has duplicates" ⚠️. Completeness % = types found ÷ 6 × 100, rounded to one decimal (`:276-277`).
- A never-called method `validate_results()` (`:340-373`) would have flagged "Annual Report" and "GSTR-3B Export" as "critical" files - an unfinished idea.

**F. Client identity and state codes** (READ; RAN where marked)

- [GST RULE] ⚠ **State name → state code** (`utils/helpers.py:485-511`, list `utils/constants.py:234-275`): the state part of the file name is trimmed, lower-cased and looked up in a fixed list of 39 entries that map full state and union-territory names to two-letter **alphabetic** abbreviations (for example 'maharashtra' → MH, 'telangana' → TS, 'odisha' and 'orissa' → OD, 'uttarakhand' → UK, 'chhattisgarh' → CG, 'dadra and nagar haveli and daman and diu' → DD, 'delhi' and 'national capital territory of delhi' → DL, 'ladakh' → LA, 'puducherry' and 'pondicherry' → PY). If the name is not in the list, the code logs a warning and builds a fallback: with two or more words, the first letters of the first two words; with one word, the first three characters in capitals. The code never uses numeric state codes, and it does not read, build or check a GSTIN or PAN anywhere (READ, HIGH).
  - RAN (`test_tree_values.py`, the helper is identical at edac6a3 and 47a1492 apart from its import line): 'Maharashtra' → MH; ' tamil nadu ' → TN; 'MH' → MH (through the fallback, with a warning); 'NCT of Delhi' → NO; 'Jammu & Kashmir' → J&; 'Gujrat' (misspelt) → GUJ; '' → ''. So a misspelt state name produces a different client key and a different folder from the correctly spelt one.
- **Client key** = `<client exactly as written>-<state code>` (`core/file_parser.py:231-232`). It is case-sensitive: "Alpha" and "alpha" are two clients, and so are "X Private Limited" and "X Pvt Ltd" - although both later get the same output folder name (G below), so the second one's reports can overwrite the first one's (READ, MEDIUM).
- **Client-state folder name** `create_client_state_key()` (`helpers.py:513-537`): whole words "Private" → "Pvt" and "Limited" → "Ltd" (any case), then `<client>-<state code>`; if longer than the maximum (default 35, `core/file_organizer.py:42`; setting `gui/main_window.py:89`), the client part is cut to fit; if that would leave 5 characters or fewer the whole key is cut, which can cut off the state code. RAN: "Alpha Beta" with maximum 8 → "Alpha Be"; a long name can end with a space before the hyphen.
- **File-name cleaning** `sanitize_filename()` (`helpers.py:66-112`): replaces the text "Private" → "Pvt" and "Limited" → "Ltd" anywhere and case-sensitively (RAN: "Privately Limitedness" → "Pvtly_Ltdness"), replaces control characters and `<>:"/\|?*[]{}+=!@#$%^,;'"`~` with `_`, turns runs of spaces/underscores into one `_`, trims `_`, `.` and spaces at the ends, uses "unnamed" if nothing is left, and cuts the name to `200 - len(extension) - 10` characters (RAN: a 300-character name → 190 characters). Hyphens are kept on purpose (comment `:81`).

**G. Output folders and names** (READ; `core/file_organizer.py`, `gui/handlers/processing_handler.py`, `power_query_extractor/...`)

```
<Target folder>/
  Annual Statement-<DDMMYY HHMM>/                  Fresh Run: new each run. Re-Run/Resume: the newest existing one by modification time
    GST_Processing_Summary_<DDMMYY_HHMM>.xlsx      summary workbook (K below)
    PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx      written later by the extractor (Q below)
    <client-state key>/                            for example "Alpha Traders-MH" (spaces kept)
      Version-<DDMMYY HHMM>/                       always a new time stamp per run
        GSTR-3B Exports[ (<cleaned client>)]/
        Other ITC related files[ (<cleaned client>)]/
        Sales related files[ (<cleaned client>)]/
        <Annual Report file, cleaned name>
        ITC_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx
        Sales_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx
        ITC_Report_..._Refreshed_<DDMMYY_HHMM>.xlsx  written later by the extractor
        _Organization_Report_<client as written>_<state as written>.txt
```

- Time stamps: `get_timestamp()` = day, month, two-digit year, space, hours, minutes (`helpers.py:23-25`); report names replace the space with `_` (`processing_handler.py:215,383`). With minute precision, two runs in the same minute share the same folders.
- The bracketed client name in the three category folders ("Level 4" in the screen text) is added when the global checkbox is on, or when the per-client "📁 Name" setting for key `<client>-<state code>` is on (`file_organizer.py:113-135`). The bracket uses the cleaned client (spaces → `_`), the client-state folder keeps spaces.
- Which file goes where (`file_organizer.py:285-298`): GSTR-3B Export → GSTR-3B folder; GSTR-2B Reco and IMS Reco → ITC folder; Sales and Sales Reco → Sales folder; Annual Report → the Version folder itself.

**H. Copying the files and the three processing modes** (READ)

- Processing (`processing_handler.py:54-447@edac6a3`) runs in a background thread. Per selected client: create the folders → copy the files under cleaned names → ITC report → Sales report → per-client text report; after all clients the summary workbook; then an information box and the output folder is opened (`os.startfile` on Windows, `open` on macOS, `xdg-open` on Linux, `:425-437`). STOP is checked only between clients (`:149-151`).
- Each file is copied with `shutil.copy2` and the copy's size compared with the original; a size difference deletes the copy (`helpers.py:142-185`). After a successful copy the in-memory record gets the cleaned name, so the template is given the copied name (`file_organizer.py:216-229`).
- If the destination already exists (`file_organizer.py:259-270`): **Fresh Run** first makes a backup `<name>_backup_<DDMMYY_HHMM><ext>` (`helpers.py:187-204`); **Resume** skips the file; **Re-Run** overwrites it silently.
- Because the Version folder always gets the current time stamp (`file_organizer.py:43,98-104`), a destination can only already exist within the same minute or when two different source files clean to the same name. In practice **Resume behaves like Re-Run** (READ, HIGH). When two source files clean to the same name (for example names that differ only in "Private" versus "Pvt" or in a space versus an underscore), Re-Run keeps only the last one copied (READ, MEDIUM).
- Optional warning before processing: if "include client name" is on and any selected client name has **more than 10 characters**, a list of up to five names is shown with "Continue anyway?" (`processing_handler.py:64-75@edac6a3`); the per-client "📁 Name" toggle has its own >10 warning (`client_handler.py:193-202@edac6a3`).
- A client counts as "Success" whenever no exception was raised, even if both its reports failed (`processing_handler.py:337-350`) (READ, HIGH).

**I. Writing paths into the templates - the "Links" sheet** - `utils/constants.py:96-120@edac6a3`, values from `core/excel_handler.py:551-661`. [GST RULE] ⚠ (it decides which file feeds which input of the ITC and Sales reports; the cell positions depend on the templates)

| Report | Cell | What is written |
|---|---|---|
| ITC | B2 | path of the GSTR-3B folder (the whole folder, so all monthly GSTR-3B files are available) |
| ITC | B4 / B5 | path of the Version folder / Annual Report file name without extension |
| ITC | B7 / B8 | path of the ITC folder / GSTR-2B Reco file name without extension |
| ITC | B10 / B11 | path of the ITC folder / IMS Reco file name without extension |
| Sales | B2 / B3 | path of the Sales folder / Sales file name without extension |
| Sales | B5 / B6 | path of the Version folder / Annual Report file name without extension |
| Sales | B8 / B9 | path of the Sales folder / Sales Reco file name without extension |

- Paths are written with backslashes (`helpers.py:311-315`); file names lose their extension (`helpers.py:446-448`), so the template must add it itself (MEDIUM). Only the **first** file of each type is used (`annual_files[0]` and so on, `excel_handler.py:577,588,599,615,631,647`) - the alphabetically first name, because the scan sorts. Further files of the same type are copied but not linked.
- Empty values are not written (`if value:`, `excel_handler.py:173,260`), so when a file type is missing the template's existing cell content stays as it was (READ, HIGH; whether the template then shows stale data depends on the template - MEDIUM).
- The sheet is found by the exact name "Links" (any case), else any sheet with "link" in its name, else - only through Excel - a new sheet "Links" is added (`excel_handler.py:146-164`); the openpyxl route writes nothing when there is no Links sheet (`:255`). The organiser never reads any figure from any workbook (READ, HIGH).
- Report file names: `ITC_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx` and `Sales_Report_...` in the Version folder (`processing_handler.py:212-219,276-277@edac6a3`). The template is copied to that name whether it is an .xlsx or an .xltx template (`excel_handler.py:135,239`). Whether Excel opens an .xltx template copied under an .xlsx name without complaint was not testable here - not sure; how to check: run one client with an .xltx template on Windows and open the produced report.

**J. Excel through COM in the organiser** (`core/excel_handler.py:103-228@edac6a3`; READ only - cannot run here)

- `pythoncom.CoInitialize()` (start COM for this thread); then **`GetActiveObject("Excel.Application")` first** - this attaches to the user's own Excel if one is open - and only otherwise `Dispatch(...)` (`:121-126`). It sets `Visible=False`, `DisplayAlerts=False`, `ScreenUpdating=False` on that Excel (`:129-131`), so a user's own open Excel would be hidden.
- It copies the template to the output name, waits 0.5 s (`:133-136`), opens the copy (`_safe_open_workbook`, `:275-317`: reuses an already open workbook with the same path; on error -2147352567, an error mentioning "path", or a path longer than 218 characters, retries with the Windows 8.3 short path), writes the Links cells, saves, closes, checks the file exists.
- In the clean-up (`:206-228`) it sets `ScreenUpdating=True` and `DisplayAlerts=True` and calls `excel.Quit()` if `self.excel_app is None`. `self.excel_app` is set to None in `__init__` (`:55`) and never assigned anywhere, so **Quit is always called - also on a user's Excel attached through GetActiveObject**, although the comment says "Don't quit if Excel was already open" (`:217`) (READ, HIGH that Quit is always called). What Excel then does with the user's unsaved workbooks while hidden and with alerts switched back on is not sure (MEDIUM); how to check: on Windows, open Excel with an unsaved workbook, run one client, and watch whether that Excel closes or asks to save. The run log (X below) records only newly started Excel instances, never an attached one, so in the logged runs no open Excel was attached.
- Without pywin32 or off Windows, `_create_report_openpyxl` (`:230-273`) copies the template and saves it through openpyxl with `keep_links=True`, after logging "Power Query content may be lost!" (`:235`). The file header explains why the Excel route exists: "WINDOWS VERSION - Uses Excel COM to preserve Power Query and all Excel features" (`:2-3`).
- `warnings.filterwarnings('ignore')` (`:17`) silences every Python warning in the whole program, not only Excel ones (READ, HIGH).
- `ExcelHandler()` creates `<system temp>/gst_excel_temp` at every start (`:52-53`) and never uses it.

**K. Summary workbook, per-client text report, exported client list** (READ)

- Summary workbook `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` in the Annual Statement folder, always written with openpyxl (`excel_handler.py:319-357`; name `processing_handler.py:383-385`): sheets "Summary" (date, totals, reports created, errors, mode, target folder, include-name Yes/No, `:380-410`), "Client Status" (Client, State, Status, Files, Missing count, Complete %, `:412-439`), "File Mapping" (original file, client key, type, destination, status, `:441-467`), "Errors" (`:469-500`), "Variations" (file name, issue, expected pattern, size, "Rename file", `:502-533`).
- Per-client text report `_Organization_Report_<client>_<state>.txt` in the Version folder (`file_organizer.py:300-362`), using the client and state exactly as written in the file names. Its "Successfully Copied" and "Folders Created" figures count **everything done so far in the whole run**, not just this client, because one FileOrganizer object is used for all clients (`processing_handler.py:110-116`, `file_organizer.py:346-351`) (READ, HIGH).
- EXPORT LIST (`client_handler.py:227-259`) writes the selected clients as "client - state" into a .txt file the user chooses (default name `selected_clients_<YYYYMMDD_HHMMSS>.txt`), with the system's default text encoding (`:243`), then opens it with `os.startfile` (`:257`), which exists only on Windows; elsewhere the file is written and then "Failed to export" is shown.

**L. Look and memory** (READ; RAN where marked)

- Dark mode (`gui/utils/dark_mode_manager.py`, `gui/main_window.py:150-180`): recolours every widget. Switching back sets every frame and canvas to plain white (`dark_mode_manager.py:171-172`), not the original light grey `#F8F9FA`, so after dark → light the screens differ from start-up (MEDIUM).
- The client-list row colours after a theme switch are wrong in the original code. In `update_tree_tags()` the `else:` that should restore the light colours belongs to the wrong `if` (`dark_mode_manager.py:198-212`):

  ```python
  if hasattr(tree_widget, 'tag_configure'):
      if is_dark_mode:
          ... dark colours ...
  else:                       # belongs to "if hasattr", not to "if is_dark_mode"
      ... light colours ...
  ```

  RAN (`test_tree_tags_toggle.py`, identical result at edac6a3 and 47a1492): after dark → light the rows keep the dark backgrounds with white text; after a re-scan the light backgrounds return but the text stays white (white on light green/orange); in dark mode a re-scan gives light backgrounds with white text, because `update_client_tree()` re-applies only the light backgrounds (`client_handler.py:63-65`). This file did not change in the shared range, so the defect is present at the split point.
- Settings cache (`gui/utils/cache_manager.py`): `~/.gst_organizer_cache.json` in the user's home folder (`:15`), JSON with source_folder, itc_template, sales_template, target_folder, processing_mode, include_client_name, dark_mode, client_name_max_length (`main_window.py:328-340`) plus `recent_folders` (last five source folders, `cache_manager.py:32-43`), saved after each Browse, the dark-mode switch, the include-name checkbox and the max-length arrows. Per-client "📁 Name" choices are not saved and are reset to "No" at every scan (`client_handler.py:49`).
- The max-length setting has three different limits: spin box 15-100 (`setup_tab.py:271-272`), typing check 10-100 (`:252-260`), label "range: 15-100" (`:279`) (READ, HIGH).
- Closing the window during processing asks "Stop and exit?" and destroys the window one second later (`main_window.py:314-325`) while the background thread may be in the middle of a copy or an Excel call (MEDIUM).

**M. The Power Query Extractor - screens and flow** (READ; `power_query_extractor/gui/extractor_window.py@edac6a3`)

- Heading "🔄 Power Query Report Extractor" and two tabs. **Setup**: Target Folder (entry, 📁 Browse, 🔍 Scan); "Select Clients" with ✓ Select All, ✗ Deselect All and a checkbox "Skip Refresh (Use existing refreshed files)"; a scrollable list with one checkbox per client (all ticked) and "ITC ✓" / "Sales ✓" labels where a report exists; "🚀 Start Processing" (`:95-232,408-468`). **Processing**: progress bar and a log (`:245-313`).
- At start it tries to load the target folder from a cache (S below), then scans.
- **Finding the processed clients** (`:353-406`): in the target folder, the newest `Annual Statement-*` folder by modification time; inside it every sub-folder whose name does not start with `_`; inside each, the `Version-*` folder with the **largest name** (`:385`). Names are `Version-DDMMYY HHMM`, so day-first text order is not date order - RAN (`test_misc.py`): of "Version-310825 2211", "Version-151025 0900" and "Version-291025 2146" the code picks the 31-08-2025 one. The client name shown and used in the consolidated workbook is the client-state folder name.
- **Per client, per report type** (`core/report_processor.py:37-141`): take the first file matching `<ITC|Sales>_Report_*.xlsx` in that Version folder (`:54-67`; the wildcard also matches the extractor's own refreshed copies, and the order of the list depends on the operating system - on Windows NTFS the original usually comes first because "." sorts before "_" - LOW-MEDIUM); copy it to `<name>_Refreshed_<DDMMYY_HHMM>.xlsx` (`:69-71,96-98`); refresh the copy in Excel (N); wait 3 s (`:117-119`); read the configured cells from the copy in a second, hidden Excel (`:462-571`). "Extract data regardless of refresh status" (`:121-122`).
- **Skip Refresh** already exists here: the checkbox writes into the shared `EXTRACTION_OPTIONS` dictionary at run time (`extractor_window.py:234-243`); the processor then uses the newest `<name>_Refreshed_*.xlsx` by modification time, or refreshes if there is none (`report_processor.py:73-93`). It does not check whether that earlier refresh had passed validation (READ, HIGH).
- After all clients, the consolidated workbook is written (Q) and the user is asked whether to open it (`os.startfile`, `extractor_window.py:570-573`). The worker thread writes to the log, the progress bar and message boxes directly (`:516-601`), although Tk widgets should only be changed from the main thread (READ, HIGH).

**N. Refreshing Power Query by key presses** (`report_processor.py:143-350@edac6a3`; READ only)

- `print(f"PROCESSING: ...")` sits before the docstring (`:144-146`, "TEST: Add this line to check if logging works"), so the description is not a docstring and the line writes to the console.
- `Dispatch("Excel.Application")` (`:154`; this may return an Excel that is already running - MEDIUM), then up to five tries, 2 s apart, to set **`Visible=True`**, `DisplayAlerts=False` and `WindowState=-4137` (maximised) ("Keep it visible to avoid issues", `:156-170`); open the workbook with three tries on "Call was rejected by callee" (`:174-185`); wait 5 × 1 s ("Increased from 3 to 5", `:187-191`); bring the window to the front with `WScript.Shell.AppActivate` (three tries, `:193-212`); then **send the keys Alt+A, R, A** one second apart ("Increased from 0.5", `:214-229`) - Excel's ribbon shortcut for Data → Refresh All.
- Wait 10 × 1 s - the comment says "WAIT 7 SECONDS AS REQUESTED" (`:233-238`), so comment and code disagree. Check five times, one second apart, for an error dialog (`_find_error_window`, `:352-388`: any visible, enabled top-level window whose title contains 'microsoft excel', 'error', 'problem', 'cannot', 'failed', 'unable', 'could not', 'exception', 'invalid' or 'missing', and whose window class is `#32770`, `NUIDialog` or `bosa_sdm_msword`); read its text (`:390-425`) and press a button labelled ok/close/yes/continue, otherwise send it a close message (`:427-460`). Wait 10 s more "for complex queries" (`:260-265`), check once more, validate (O), save, close.
- There is no `RefreshAll()` call, no wait for asynchronous queries to finish and no check of whether a query is still refreshing; the refresh is assumed finished after the fixed waits (READ, HIGH). Because the keys go to whatever window is in front, the user must not use the computer while a refresh is running, and the program needs an interactive, logged-in desktop. Each report spends roughly 40 seconds in fixed waits alone (adding up the waits in the code; MEDIUM).
- The refreshed copy is saved even when validation fails (`:279-296`), and a later Skip Refresh run can reuse it (M above) (READ, HIGH).
- Clean-up closes without saving, quits Excel and stops COM (`:329-350`). A final `return {'success': True ...}` (`:318-321`) can never run, and the `except` in `_process_report` that looks for "Open method of Workbooks class failed" (`:101-115`) can never trigger, because `_refresh_power_query_simple` catches every error itself (READ, HIGH). Comments such as "ADD THE NEW CODE HERE (notice the indentation - 3 levels)" (`:73`) and "Rest of your code remains same..." (`:231`) are left-overs of code pasted from chat answers.
- Opening (`_safe_open_workbook`, `:622-731`): paths over 218 characters are converted to the Windows long-path form; then up to four further attempts (close same-named workbooks and retry; 8.3 short path; without the long-path prefix; `Open(..., UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, Notify=False)`), then advice in the log (file open elsewhere, damaged, path too long, OneDrive/SharePoint sync).

**O. When a refresh counts as successful** - `report_processor.py:573-620@edac6a3`. [GST RULE] ⚠ (a rule about the template; the sheet and cell depend on the template)

- A sheet named "Info" (any case) must exist and its cell **BB2** must be not empty, not blank, not "0" or "0.0", and not one of `#REF! #VALUE! #NAME? #NULL! #DIV/0! #N/A #NUM! #ERROR!`. The success message calls BB2 "Client Name" (`:616`).

**P. Which cells the extractor reads** - `power_query_extractor/config/cell_mappings.py@edac6a3`. [GST RULE] ⚠ (sheet names and cell positions depend on the templates; the GSTR-3B and GSTR-1 references belong to return formats that can change)

| Output sheet | Report | Template sheet!cell → output column | Lines |
|---|---|---|---|
| GSTR-3B Difference | ITC | `Diff GSTR-3B`!J15 → "Diff IGST"; !K15 → "Diff CGST"; !L15 → "Diff SGST" | `:35-58` |
| GSTR-1 Difference | Sales | `Diff - GSTR-1`!N1 → "Diff Taxable Value"; !O1 → "Diff IGST"; !P1 → "Diff CGST"; !Q1 → "Diff SGST" | `:61-89` |
| Turnover GSTR-1 | Sales | `Diff - GSTR-1`!I1 → "Taxable Value" | `:92-102` |

- The file is written to be edited by hand ("MASTER VERSION ... HOW TO EDIT THIS FILE", `:1-27`); every entry can also name a calculation (subtract, add, multiply, divide; divide by zero gives the text "Div/0", non-numbers "Error", `data_consolidator.py:188-218`). No entry uses one.
- `EXTRACTION_OPTIONS` (`:122-127`): skip_refresh False (changed by the checkbox), auto_refresh_if_missing True (never read), show_timestamp False (never read), missing_data_text "Missing".
- A "backward compatibility" block (`:133-152`) turns the list into the older shape `CELL_MAPPINGS[report][sheet][cell] = "<output sheet>_<output column>"` (spaces → `_`) that the reader still uses. If two output sheets ever named the same template sheet and cell, the later one would silently replace the earlier one (READ, HIGH; no such pair exists in the shared range).
- While reading, each value is stored twice: under its field name and under its bare cell reference such as "J15" "for calculations" (`report_processor.py:519-523`). Bare references from different template sheets of the same report would overwrite each other (READ, MEDIUM; no calculation uses them here).
- Sheet names must match exactly, including case (`report_processor.py:500`); a missing sheet gives empty values. The extraction log claims "Opening file in read-only mode" (`:476-478`) but the workbook is opened normally (read-write) and closed without saving (READ, HIGH). Every value read is written into `pq_extractor.log` and the window log (`:525,541-546`).
- `cell_mappings copy.py@edac6a3` holds the older format directly (`'Diff GSTR-3B': {'J15': 'itc_diff_igst', ...}`, `'Diff - GSTR-1': {'N1': 'sales_diff_tv', ...}`) with display names such as "Diff IGST (ITC)" and commented ideas `'E5': 'reversed_itc'`, `'F5': 'net_itc'`, `'D10': 'export_sales'`, `'E10': 'exempt_sales'`. It names the same cells as the current file except `Diff - GSTR-1`!I1, which is new. The compiled-file evidence (Y) dates it to 21-06-2025.

**Q. The consolidated extraction workbook** (`power_query_extractor/core/data_consolidator.py:20-186@edac6a3`; READ, RAN where marked)

- `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` (year first - a different date order from the organiser's folder and report names) in the newest Annual Statement folder; a "Summary" sheet first (date, number of clients, "ITC Reports Processed: n/total", list of sheets, `:220-248`), then one sheet per output sheet with columns Client Name, Successfully Run (✓/✗), then one column per mapping.
- [GST RULE] When a report's refresh did not succeed, every value cell of that report shows the text "Missing" (`:128-129`), even if values were read. Non-zero numbers get the Excel display format `#,##0.00` (`:161-162`); zeros keep the default format; the stored values are not rounded (READ, HIGH).
- Each column's value is looked up by a **substring test** first - "is the column name contained in a field name?" - and only then by the exact field name (`:141-150`):

  ```python
  for field_name, field_value in extracted_data.items():
      if mapping['output_column'] in field_name:     # substring test
          value = field_value
          break
  if value is None and mapping['input_cell']:        # exact name only if nothing was found
      key = f"{sheet_name}_{mapping['output_column']}".replace(' ', '_')
      value = extracted_data.get(key)
  ```

  At edac6a3 this is harmless by accident: every column name contains a space and every field name uses underscores, so the substring test never matches and the exact lookup finds every value. RAN (`test_consolidator.py`, each value set to its own "sheet!cell" label): every column at edac6a3 shows its own cell. It becomes harmful at 47a1492 (see that commit).
- Old methods `_write_section` (`:250-352`) and `_create_details_sheet` (`:354-414`) belong to an earlier single-sheet layout (ITC section, Sales section, details) and are never called; lines 73-83 build and write the header row a first time with "# ... rest of the code" place-holders, and lines 85-99 do it again properly.

**R. Every constant in `utils/constants.py@edac6a3`** (unchanged up to the split point; READ)

| Constant | Lines | Content | Used? |
|---|---|---|---|
| `FILE_PATTERNS` | 13-62 | the six patterns (D) | yes |
| `EXPECTED_FILE_TYPES` | 68-75 | the six types (E) | yes |
| `FOLDER_STRUCTURE` | 81-90 | `Annual Statement-{timestamp}`, `{client}-{state}`, `Version-{timestamp}`, `GSTR-3B Exports ({client})`, `Other ITC related files ({client})`, `Sales related files ({client})` | yes, except `level2` (the code uses `create_client_state_key`) |
| `EXCEL_TEMPLATE_MAPPING` | 96-120 | the Links cells (I) | yes |
| `GUI_CONFIG` | 126-151 | title "GST File Organizer v3.0", 1200x800, minimum 1000x700; Segoe UI and Consolas fonts; colours primary #0078D4, success #107C10, warning #FF8C00, danger #D83B01, info #0099BC, light #F8F9FA, dark #323130, white, complete #E8F5E8, incomplete #FFF4E5 | yes |
| `PROCESSING_MODES` | 157-173 | fresh "🆕 Fresh Run" (new folder structure), rerun "🔄 Re-Run" (new version folders in existing structure), resume "▶️ Resume" (continue interrupted processing) | yes |
| `EXCEL_EXTENSIONS`, `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB = 100`, `EXCEL_SIGNATURES` | 179-185 | extension sets, size limit, file signatures | no (the helpers hard-code their own copies) |
| `MESSAGES` | 191-211 | error, success and info texts | only no_source, no_templates, no_target, no_selection |
| `EXCEL_SAFETY` | 217-228 | max_retries 3, retry_delay 1, temp_prefix `temp_gst_`, backup_prefix `backup_`, verification_size_min 10,240, save_options all False | stored in `ExcelHandler.safety_config` (`excel_handler.py:54`) and never read |
| `STATE_CODE_MAPPING` | 234-275 | 39 state/territory names → two-letter codes (F) [GST RULE] ⚠ | yes |

**S. Who writes and who reads shared files** (READ, HIGH unless stated)

| Shared thing | Written by | Read by | Do they agree? |
|---|---|---|---|
| organiser settings | `gui/utils/cache_manager.py:15,29-48` → `~/.gst_organizer_cache.json` (home folder, name starts with a dot) | `gui/main_window.py:342-356` | yes |
| extractor "auto-load from cache" | nobody writes these | `extractor_main.py:36-60` tries (candidate list `:38-42`) `gst_organizer_cache.json`, `.cache/app_state.json`, `config/last_session.json` relative to the start folder, keys target_folder / target_path / output_folder / last_target; `extractor_window.py:28,329-341` tries `gst_organizer_cache.json` | **no** - different folder (start folder versus home folder) and different name (no leading dot). The extractor never finds the organiser's settings, so the promise "it will auto-load from cache" (`launch_extractor.py:16`) does not come true |
| log files | `main.py:25` → `./gst_organizer.log`; `extractor_main.py:16` → `./pq_extractor.log` | people | both relative to the start folder, appended for ever |
| organiser's reports | `processing_handler.py:218-219,276-277` (`ITC_Report_*`, `Sales_Report_*`) | `report_processor.py:54-55`, `extractor_window.py:388-389` (same wildcards) | yes, but the wildcard also matches refreshed copies (M) |
| refreshed copies | `report_processor.py:70-71,97-98` (`_Refreshed_`) | Skip Refresh lookup `:80` (`_Refreshed_`) | yes at edac6a3 (bd176bf breaks this) |
| `EXTRACTION_OPTIONS['skip_refresh']` | checkbox `extractor_window.py:237-238` | `report_processor.py:75-76` | yes (same module object in memory) |
| per-client "📁 Name" choice | toggle stores `<client>-<state as shown>` (`client_handler.py:184,220`) | organiser looks up `<client>-<state code>` (`file_organizer.py:118,126`) | **no** whenever the file names carry a full state name ("Maharashtra" versus "MH"): the toggle is then ignored. It works only when the state part is already the code (READ, HIGH) |

**T. Dependencies** (READ)

- `requirements.txt`: `openpyxl>=3.1.2`, `lxml>=4.9.0` (lxml is never imported; openpyxl can use it if installed), commented `pytest>=7.0.0`, `black>=23.0.0`, `pylint>=2.17.0`. No version is pinned. Comment "Python 3.7 or higher required".
- Used but not listed: **pywin32** (`win32com.client`, `pythoncom`, `win32api`, `win32con`, `win32gui`) - guarded in `core/excel_handler.py:19-28` and `utils/helpers.py:473-483`, unguarded in `power_query_extractor/core/report_processor.py:11-15`. Tkinter comes with Python.

**U. Dead code and unfinished ideas at the first commit** (READ; searched for callers)

- `utils/helpers.py`: `get_unique_path` (`:344`), `calculate_file_hash` (`:422`), `get_relative_excel_path` (`:450`) - re-exported, never called.
- `utils/constants.py`: the unused constants listed in R.
- `core/file_parser.py`: `validate_results` (`:340-373`), `self.errors`, the stored "warnings" and "metadata".
- `core/file_organizer.py`: `cleanup_on_error` (`:364-379`), `_is_folder_empty` (`:381-390`, used only by it), `get_summary` (`:392-409`).
- `core/excel_handler.py`: `self.temp_dir`, `self.safety_config`, `self.excel_app` (never assigned), so `__del__` (`:57-63`) never does anything.
- `gui/utils/ui_helpers.py`: `update_status` (`:11-14`); `gui/utils/cache_manager.py`: `get_cached_value`, `set_cached_value`.
- Extractor: `_write_section`, `_create_details_sheet`, the duplicate header code (Q); `self.excel_instance`, the unreachable return, the dead `except`, `cleanup()` that does nothing (`report_processor.py:733-736`); options auto_refresh_if_missing and show_timestamp; calculations (never configured); the whole `cell_mappings copy.py`.

**V. Defects already present in the first commit** (technical points only; all still present at the split point unless a later shared commit is named)

| # | Defect | Evidence | Fixed in the shared range? |
|---|---|---|---|
| 1 | Extractor cannot find the organiser's settings cache | S; READ, HIGH | no |
| 2 | Consolidator's substring lookup can put another sheet's value into a column | Q; RAN | harmless until 47a1492, which triggers it |
| 3 | Per-client "📁 Name" choice stored under a different key than the organiser reads | S; READ, HIGH | no |
| 4 | Every client row coloured "complete" in light mode and "incomplete" in dark mode | `client_handler.py:51-61@edac6a3`; RAN | yes, 1c48722 |
| 5 | Row colours not restored after dark → light; white text on light rows after a re-scan | L; RAN | no |
| 6 | Disabled START/STOP still react to clicks, and nothing stops a second processing thread from starting (the Validation tab's START is never disabled, `validation_tab.py:133-142`) | B; RAN for the binding, READ for the missing guard | no |
| 7 | Extractor picks the "latest" Version folder by name, which is not date order | M; RAN | no |
| 8 | Number-like client or state text breaks client selection: the client list gives "0123" back as the number 123 and "07" as 7, so the key no longer matches, and `get_state_code()` crashes on a number (`AttributeError: 'int' object has no attribute 'strip'`) in Dry Run, START and EXPORT LIST (`client_handler.py:167-179`) | RAN (`test_tree_values.py`) | no |
| 9 | Resume behaves like Re-Run | H; READ, HIGH | no |
| 10 | Program stops at the first line without a console | A; READ, HIGH | no |
| 11 | Background threads change Tk widgets directly (organiser processing, extractor processing) | H, M; READ | partly addressed by 6237597 (with a new bug) |
| 12 | Organiser may attach to the user's own Excel, hide it and quit it | J; READ | no |
| 13 | Refresh by key presses with fixed waits; no completion check | N; READ | waits made adjustable in bd176bf |
| 14 | "Read-only" log message while the file is opened read-write | P; READ, HIGH | no |
| 15 | Name variants of one client ("Private Limited" / "Pvt Ltd", different capitals) are separate clients at scan time but share one output folder; the later client's report copy replaces the earlier one without a backup (`excel_handler.py:135`) | F; READ, MEDIUM (for capital-letter variants this holds on Windows, whose file names ignore case) | no |
| 16 | Three different limits for the max-length setting | L; READ, HIGH | no |
| 17 | Window destroyed one second after "Stop and exit?" while work may be running | L; READ, MEDIUM | no |
| 18 | EXPORT LIST opens the file with a Windows-only call and writes with the default encoding | K; READ | no |
| 19 | Frames become white instead of light grey after dark → light | L; READ, MEDIUM | no |
| 20 | Files under 1 KB or with a wrong signature vanish from the scan without any message | C; READ, HIGH | no |
| 21 | Per-client text report shows run totals instead of per-client counts | K; READ, HIGH | no |
| 22 | A client is counted "Success" even when both reports failed | H; READ, HIGH | no |
| 23 | Skip Refresh can reuse a refreshed copy whose validation failed | M, N; READ, HIGH | no |
| 24 | Client or state names with a hyphen can never be recognised; `.xlsm` files always become variations | D; READ, HIGH | no |

**W. `README.md@edac6a3` compared with the code** (READ)

| README says | Code does |
|---|---|
| folder tree with `gui/app.py` and no tabs/handlers/widgets and no extractor (`README.md:36-54`) | the GUI is split into `main_window.py`, `tabs/`, `handlers/`, `utils/`, `widgets/`; `gui/app.py` is not in the commit; the extractor is not mentioned anywhere in this README |
| "Operating System: Windows 10/11, macOS, Linux" (`:18`) | the organiser runs elsewhere only through the openpyxl fallback; the extractor cannot start without pywin32 |
| names `Annual Statement-DDMMYYYY HHMMSS`, `Version-DDMMYYYY HHMMSS`, `ITC_Report_ClientName_State_DDMMYYYY_HHMMSS.xlsx` (`:111-123`) | `DDMMYY HHMM`, and the state **code** in report names (G) |
| "All Excel operations use temporary files", "Files are verified before finalization", "No VBA macros or external links preserved" (`:129-133`) | templates are copied straight to the output name; `EXCEL_SAFETY` is never read; the Excel route keeps whatever the template contains and the openpyxl route uses `keep_links=True` |
| "Backups created for existing files" (`:208`) | only in Fresh Run, only for copied source files |
| "View variations if files don't match naming convention" (`:101`) | variations appear only as a count in the scan summary and in the summary workbook |
| naming patterns (`:80-85`), Links cells (`:167-184`), summary sheets (`:196-202`) | match the code |
| "The Excel corruption issues from previous versions have been specifically addressed" (`:254`) | evidence of an earlier version that damaged Excel files (before git) |

**X. `gst_organizer.log` (added here, deleted in cbfa4a5) - described by kind only**

How I looked (RAN): counting commands only - line count, and counts of lines per logger name, level and date, and of lines beginning with fixed message texts taken from the code. Nothing from the file's content was printed or copied, and no figure from its messages is given here.

- **What it is**: the organiser's own run log (5,689 lines, the size git reports for the file). Every line has the standard form `<date time,milliseconds> - <logger name> - <level> - <message>`. All lines are level INFO (the "[DEBUG]" messages in the code are written at INFO level, which is why they appear). The writers are the program's own modules: `core.file_parser`, `core.excel_handler`, `core.file_organizer`, `utils.helpers` and the root logger. The entries span 23-07-2025 to the evening of 31-08-2025 (the computer's own clock), ending minutes before the commit.
- **What kind of events it records** (from the code's message texts the lines match; every line matches one): program start-ups on several days, folder scans, and one processing run on 23-07-2025 for a few clients (folders created, files copied, ITC and Sales reports created through a newly started Excel - never an already open one - with the Links cells written). There are no warning or error lines, so in these runs every scanned file matched a pattern and every state part was found in the state list (MEDIUM inference).
- **Kinds of information inside**: client names and state names; source file names; full Windows folder paths of the source folder, target folder, templates and outputs; output report names; file sizes and counts. **No tax figures** - the organiser never reads cell values (READ, HIGH).
- **How it was produced**: `logging.basicConfig(... FileHandler('gst_organizer.log') ...)` at `main.py:21-28@edac6a3`, a file relative to the start folder that is appended to on every start. The program was evidently started from the project folder, and with no `.gitignore` the file was committed with everything else (HIGH).

> **Warning - client data in the history of both repositories.** The log was deleted three minutes later (cbfa4a5), but deleting a file in a new commit does not remove it from history. Because edac6a3 is the first commit of both projects, anyone with either repository can still read the log at edac6a3 (READ, HIGH). Removing it would need a rewrite of both histories - the author's decision.

**Y. The 70 compiled `.pyc` files (added here, removed in 01862a0) and the lost `gui/app.py`**

How I looked (RAN): `pyc_headers.py` prints only each file's header fields (Python version marker, source size, source modification time); `pyc_pathcount.py` counts distinct embedded folder locations by a one-way fingerprint. Neither prints any path, and none is reproduced here.

- **What they are**: 70 files = 35 modules × 2 Python versions, `cpython-312` (Python 3.12) and `cpython-313` (Python 3.13). They cover every importable module of the commit plus `gui/app`; there are none for `main.py`, `launch_extractor.py` (run as scripts) or `cell_mappings copy.py` (not importable). They are present in edac6a3, cbfa4a5, 6237597, 9c56a1c, d674386, 5ea77c2 and 1cb878b.
- **They embed the developer computer's folder paths**: Python stores the full path of the source file inside each compiled file. All paths my scan matched are Windows drive-letter paths, with at least four different project-folder locations per Python version - the folder was copied or moved several times with its `__pycache__` carried along (MEDIUM). Folder names inside such paths could include personal or client names; I did not inspect them.
- **Header dates and sizes** (UTC): `core/__init__`, `utils/__init__` 19-06-2025; extractor package files and `extractor_main` 21-06-2025; `excel_handler`, `file_parser`, `constants` and `gui/app` 25-06-2025; `cell_mappings`, `report_processor`, `extractor_window` 28-06-2025 (3.13 copies); `data_consolidator` 14-07-2025 (3.13); the modular GUI files, `file_organizer` and `helpers` 23-07-2025 06:08-08:15 UTC (11:38-13:45 IST). For every module that has a committed source, the Python 3.13 copy records exactly the committed file's size with Windows line endings (for example `utils/constants.py`: 8,861 bytes + 274 line ends = 9,135) - so the code was edited on Windows and the compiled files match the committed text (HIGH).
- Four Python 3.12 copies record **older** sources: `cell_mappings` 1,938 bytes on 21-06-2025 - exactly `cell_mappings copy.py` (1,876 + 62 line ends), so the "copy" is the 21-06-2025 version (HIGH); `data_consolidator` 10,208 bytes on 25-06-2025 against 17,711 committed on 14-07-2025 (the consolidator nearly doubled in July, MEDIUM: probably the per-sheet layout replacing the dead `_write_section` layout); `report_processor` 29,229 (25-06) against 30,861 (28-06); `extractor_window` 20,875 (21-06) against 21,963 (28-06).
- **`gui/app` (98,247 and 97,844 bytes compiled)** belongs to a `gui/app.py` of **89,275 bytes**, last changed 25-06-2025 16:14 UTC, which is in no commit. The first GUI was one large file; on 23-07-2025 it was split into `gui/main_window.py` ("Modularized version - ties all components together", `gui/main_window.py:4`) plus the tab, handler, utility and widget files. `app.py` was deleted before the first commit but its compiled copies were committed. The README tree (`README.md:51-53`) and `main.py:72` still name it (HIGH).
- Both Python 3.12 and 3.13 compiled the modular GUI after 23-07-2025, so two interpreters - perhaps two computers - were in use (MEDIUM).

**Why it changed** (inferred)

- First upload of a program that had been developed and used for about ten weeks before git: compiled-file dates run from 19-06-2025, the log from 23-07-2025, the README already speaks of "previous versions" with corruption problems (HIGH). Everything in the folder was committed, including build leftovers and the run log, because there was no `.gitignore` (HIGH).
- The design history before 31-08-2025 - in particular why the GUI was split on 23-07-2025 and why the extractor refreshes by key presses - is lost; only the compiled files and comments hint at it (HIGH).

**What it reveals**

- **Data flow**: source folder → file-name recognition → per-client grouping keyed by client and state code → folders and cleaned copies → template copies with folder paths written into a "Links" sheet → (later, separate program) Excel refresh of those templates → a few cells read back → one consolidated workbook. The templates do all the GST work; the organiser never reads a figure, and the extractor only copies figures from fixed cells (READ, HIGH).
- **The naming convention is the core contract**: the six patterns decide client, state and return type. Everything later (folder names, keys, report names, the extractor's discovery) depends on them, and they forbid hyphens in names.
- **Rules introduced here**: the six patterns and their order [GST RULE] ⚠; the six expected files with the GSTR-3B exception [GST RULE] ⚠; the Links cell map [GST RULE] ⚠; the extraction cell map [GST RULE] ⚠; the state list and fallback [GST RULE] ⚠; the Info!BB2 refresh test [GST RULE] ⚠; "Missing" text and `#,##0.00` display [GST RULE]; Private/Limited abbreviations; 35-character folder key; the >10-character name warning; processing modes.
- **Edge cases met before git** (evidence in the code, MEDIUM): Excel busy ("Call was rejected by callee", `report_processor.py:181-185`); Excel properties failing to set (`:158-170`); paths over 218 characters (`:648-664`, `excel_handler.py:303-315`); OneDrive/SharePoint and files open elsewhere (`report_processor.py:720-729`); window activation failing (`:198-210`); Excel error dialogs during refresh (`:240-277`); duplicate downloads named `(2) GSTR3B-...` (`constants.py:31`); damaged Excel files and lost Power Query in earlier versions (README; `excel_handler.py:2-3`); refresh waits raised step by step ("Increased from 0.5", "Increased from 3 to 5", the "7 SECONDS" comment above a 10-second loop).
- **Design philosophy**: defensive code everywhere (try/except around almost every step, retries, fallbacks), very detailed logging (including file names and extracted values), configuration kept in Python files meant to be edited by the user ("EDIT THIS SECTION"), Windows and Excel as the real platform with a weaker openpyxl fallback for the organiser only, and two loosely coupled programs. Strongly AI-assisted authorship: "Author: Advanced AI Assistant" (`main.py:6`), comments "ADD THIS LINE" (`utils/helpers.py:15`), "ADD THIS" / "ADD THESE" (`core/file_parser.py:15`, `core/file_organizer.py:15`, `core/excel_handler.py:39`), "ADD THE NEW CODE HERE (notice the indentation - 3 levels)" and "Rest of your code remains same..." (`report_processor.py:73,231`), "# ... rest of the code" (`data_consolidator.py:75,80`) - code pasted from chat answers (HIGH).

##### cbfa4a5 - 31-08-2025 22:14 - shared - before the split on 30-10-2025

**Commit message:** "Delete gst_organizer.log" (no body). Committer "GitHub" with GitHub's no-reply address, author goyal-sidhant: the commit was made on GitHub's web page, not on the computer (READ, HIGH). Parent: edac6a3.

**What changed:** `gst_organizer.log` deleted (−5,689 lines). Nothing else (READ: `git show --stat cbfa4a5`). The log's content was not read; see X under edac6a3.

**Why it changed:** the author noticed, three minutes after the first upload, that the run log - which holds client names, file names and folder paths - had been uploaded, and deleted it in the browser (HIGH).

**What it reveals:**
- Awareness that the log is sensitive. But a deletion in a later commit leaves the file readable in history at edac6a3 in both repositories (READ, HIGH; see the warning under X).
- The computer did not receive this commit for 59 days: the next local commit (6237597) is still built on edac6a3, and the deletion arrived only through the merge 9c56a1c on 29-10-2025 (READ, HIGH).
- No code, rule or dependency changed.

---

#### Phase 2: Code review, README rewrite and repository clean-up (29-10-2025)

After 59 days without commits, the author returned to the code on the afternoon of 29-10-2025. An AI-assisted review produced one code commit; the pull made before pushing it brought in GitHub's log deletion; in the evening the README was rewritten, the broken shebang removed, a `.gitignore` added and the compiled files removed from the repository.

##### 6237597 - 29-10-2025 17:21 - shared - before the split on 30-10-2025

**Commit message:** "Fix critical import and threading bugs in GST Organizer", with the body: "Fixed circular import error in utils/helpers.py (STATE_CODE_MAPPING) / Fixed scanned_files attribute access error in file_handler.py (line 75) / Fixed thread safety issues in processing_handler.py GUI updates / Fixed indentation bug in long client name validation logic / Corrected non-existent app.py reference in main.py error messages / Improved progress_callback None handling in file_parser.py / These fixes resolve runtime crashes and import errors that prevented the application from launching and processing files correctly." Parent: edac6a3 (made on the computer, which did not yet have cbfa4a5). (READ)

**What changed - 4 files, +59 / −45** (READ: full diff)

| File | +/− | Change in plain English |
|---|---|---|
| `core/file_parser.py` | +11 / −13 | `__init__` adds `self.progress_callback = None` (`:30@6237597`); new method `_update_progress()` calls the stored progress callback inside a try/except and logs a warning if it fails (`:49-55`); `scan_folder()` stores the callback (`:150`) and calls `_update_progress()` instead of calling the callback directly; about a dozen blank lines inside `scan_folder()` are removed. Real effect: an error raised by the progress display no longer aborts the scan. Line numbers move: +9 from `parse_filename` (edac6a3 line 48 → 57) to `scan_folder`, then −2 from `_process_file` onward (181 → 179) |
| `gui/handlers/file_handler.py` | +1 / −1 | line 75: `len(self.app.file_parser.scanned_files)` → `len(self.app.scanned_files)` |
| `gui/handlers/processing_handler.py` | +46 / −30 | (1) In `dry_run()` every log call is wrapped as `self.app.root.after(0, lambda: self.app.log_message(...))` (`:36-52@6237597`). (2) In `start_processing()` the long-name warning is rewritten (`:64-91`): it reads `max_len = int(self.app.client_name_max_length.get())`, uses 10 if that is 0 or less or cannot be read, looks clients up with `.get()`, collects (name, length) pairs and shows "Long client names detected (exceeding {max_len} chars):" with up to five names and "If you continue, these client names will be used in folder names."; the "PROCESSING STARTED" line is wrapped in `root.after` (`:104`). (3) In the worker thread `process_files_thread()` the progress updates and log lines from the client header down to the ITC "Output:" line are wrapped in `root.after(0, lambda: ...)` (`:178-239`) |
| `utils/helpers.py` | +1 / −1 | inside `get_state_code()`: `from utils.constants import STATE_CODE_MAPPING` → `from .constants import STATE_CODE_MAPPING` (`:491`) |

**Each claim in the message checked against the diff** (READ unless marked)

| Claim | What the diff and the old code show | Confidence |
|---|---|---|
| "Fixed circular import error in utils/helpers.py" | The import was already inside the function with the comment "Import here to avoid circular import" (`utils/helpers.py:490-491@edac6a3`). `utils.constants` and `.constants` are the same module whenever helpers is loaded as `utils.helpers`, which is always the case in this program. No behaviour change. RAN (`test_tree_values.py` on both copies): `get_state_code()` gives identical results at edac6a3 and 47a1492 | HIGH |
| "Fixed scanned_files attribute access error in file_handler.py (line 75)" | `self.app.scanned_files` is the very dictionary returned by the parser (`file_handler.py:71-72`, `file_parser.py:174@edac6a3`), and the parser always defines `scanned_files` (`file_parser.py:34`). The old line could not raise an error; the change has no effect | HIGH |
| "Fixed thread safety issues in processing_handler.py GUI updates" | Only part of the GUI calls are wrapped: everything after line 239 of `process_files_thread()` - the ITC result lines, the whole Sales section, the organisation report, the summary workbook, the final messages, `messagebox.showinfo` (`:435`), `messagebox.showerror` (`:458`) and the button resets (`:462-463`) - still runs directly in the worker thread. The existing helpers already passed their work to the main thread (`UIHelpers.log_message` uses `log_widget.after(0, ...)`, `gui/utils/ui_helpers.py:23-25`; `update_progress` uses `root.after(0, ...)`, `gui/main_window.py:260-265`), so the new wrapping puts an `after` inside an `after`. And it introduced a bug (below) | HIGH |
| "Fixed indentation bug in long client name validation logic" | At edac6a3 `if long_names:` was already inside the `if include_client_name ...` block (12 spaces, `processing_handler.py:71@edac6a3`); the diff only removes trailing spaces from a blank line there. What really changed is the threshold (below) | HIGH |
| "Corrected non-existent app.py reference in main.py error messages" | `main.py` is **not in the diff**. It still prints "app.py" at this commit and at the split point (`main.py:71@47a1492`) | HIGH |
| "Improved progress_callback None handling in file_parser.py" | The old code already checked `if progress_callback:`; the real gain is that errors from the callback are caught | HIGH |
| "These fixes resolve runtime crashes and import errors that prevented the application from launching" | No code path at edac6a3 crashes on the points listed (RAN: the window builds at edac6a3, `test_window.py`; `get_state_code()` works). The claims read like the guesses of a code review, not observed failures (MEDIUM). Whether the author had seen a different crash (for example the console problem in A above, or a missing pywin32) cannot be told from the diff (LOW) | - |

**The new late-binding bug** (RAN, HIGH)

Each wrapped line looks like this (`processing_handler.py:42,46@6237597`):

```python
for client_key in selected:
    client_info = self.app.client_data[client_key]
    self.app.root.after(0, lambda: self.app.log_message(f"\n🏢 {client_info['client']} - {client_info['state']}", 'info'))
    for file_type, files in client_info['files'].items():
        for file_info in files:
            self.app.root.after(0, lambda: self.app.log_message(f"  📄 Would organize: {file_info['name']}", 'normal'))
```

The lambda reads `client_info` and `file_info` only when Tk runs it - after the loop has finished - so every line shows the last client and the last file. RAN (`test_dryrun.py`, two made-up clients "Alpha Traders" with two files and "Beta Stores" with one): at edac6a3 the Dry Run preview lists both clients with their own files; at 6237597 it shows "Beta Stores - Karnataka" twice and Beta's file name for all three "Would organize" lines. Dry Run runs entirely in the main thread, so it is always wrong. The same pattern in the worker thread (`:178-239`: client header, folder names "Level 1/2/3", the "📁 ... → ... folder" lines, ITC template and output names) shows wrong values whenever the worker runs ahead of the window (MEDIUM; needs real timing, not run). The per-file progress lambda inside `file_progress()` (`:203-206`) is not affected, because each call of that function creates its own variables.

**Why it changed** (inferred): an AI-assisted code review on the first day back after 59 days (HIGH for the gap; MEDIUM for the review: the message style with "(line 75)", the inline comment "CHANGE: ... This is safer and easier to review." at `processing_handler.py:64-67@6237597`, and claims that do not match the diff).

**What it reveals:**
- **Business rule change - long client-name warning before processing**: threshold "more than 10 characters" → "more than the 'Client folder name max length' setting" (default 35; 10 if the setting is 0 or less or unreadable) (`processing_handler.py:64-91@6237597`). Two other places keep 10: the per-client "📁 Name" toggle (`client_handler.py:194@6237597` = `:188@47a1492`) and the Setup label "⚠️ Names >10 chars will show warning" (`setup_tab.py:238`). So after this commit three places disagree (READ, HIGH). The new code's own comment still says "(falls back to 10)".
- **Threading**: the concern was real (Tk is not thread-safe) but the fix was partial and introduced the late-binding bug. The extractor's worker thread (`extractor_window.py:516-601@edac6a3`) was not touched.
- **Message versus diff**: of the six claims, only the progress-callback one roughly matches the diff; the thread-safety one is partial, three describe errors that could not happen, and the main.py one describes a change that is not in the diff. The README of d674386 repeats them as "Recent Bug Fixes (Latest Commit)".
- Abandoned: the direct log calls of edac6a3 in Dry Run and the first part of processing.
- No dependency changed.

##### 9c56a1c - 29-10-2025 17:21 - shared - before the split on 30-10-2025

**Commit message:** "Merge branch 'main' of https://github.com/goyal-sidhant/AnnualStatement" (no body) - git's standard text for a `git pull` that has to merge. Parents: 6237597 (first parent, the local work) and cbfa4a5 (second parent, GitHub's log deletion). Made 15 seconds after 6237597 (17:21:12 → 17:21:27). (READ)

**What changed:** `git show -m --first-parent 9c56a1c` shows a single change against the local line of work: **`gst_organizer.log` deleted (−5,689 lines)**. Against the GitHub side it adds 6237597's four files. The combined diff is empty, so there were no conflicts (READ). Merging the deletion also removed the log from the author's local project folder (git deletes a tracked file when it merges its deletion) (HIGH, git behaviour).

**Which repository it pulled from - what GitHub's records say** (dates supplied to this audit from GitHub's records; not re-checked here): the repository now called AnnualStatement_oldv2 (project B) was created on 31-08-2025, the day of edac6a3 and cbfa4a5; the repository now called AnnualStatement_oldv1 (project A) was created on 07-02-2026. On 29-10-2025 only the first of them existed, so this merge pulled from the repository that is now **AnnualStatement_oldv2**, which was then named "AnnualStatement" (HIGH, by dates). Today the old name "AnnualStatement" forwards to AnnualStatement_oldv1 - so the original repository was renamed, and AnnualStatement_oldv1 later carried the name "AnnualStatement" before being renamed itself (MEDIUM: GitHub forwards an old name to the repository that most recently gave it up; the order of renames is inferred, not recorded in git). Practical consequence: the URL written in this merge message now leads to AnnualStatement_oldv1, not to the repository the merge actually came from (MEDIUM). How to check: the "created" dates and rename history on each repository's GitHub page (or `gh api repos/goyal-sidhant/AnnualStatement`, which shows where the old name resolves).

**Why it changed:** a routine pull before pushing the new commit: GitHub already held cbfa4a5, which the computer lacked, so git had to create a merge (HIGH that this was a pull with an automatic merge; whether it was triggered by a refused push or by an editor's "sync" button is not recorded).

**What it reveals:**
- The author works on the computer and pushes to GitHub; the deletion made on GitHub sat unmerged for two months.
- Git refuses such a merge when the local copy of a file that the other side deletes has uncommitted changes. The merge went through, so the local log was unchanged since 31-08-2025 (or already deleted). Since every start of the organiser from that folder appends at least one line to the log (on Windows, importing `core/excel_handler.py` logs "Using win32com ..." or "win32com not available ..." at `:26` / `:28@edac6a3`, after `main.py:21-28` has opened the log), the organiser was most likely **not started from the project folder between 31-08-2025 and this merge** - which also means the 6237597 changes were probably committed without being run there (MEDIUM).
- No code, rule or dependency changed.

##### d674386 - 29-10-2025 21:15 - shared - before the split on 30-10-2025

**Commit message:** "Update README.md for clarity and formatting improvements" (no body). (READ)

**What changed - `README.md` +239 / −135 (260 → 364 lines)** (READ: full diff and the new file). The README is rewritten:
- New title "📊 GST File Organizer v3.0 - Production Ready"; new sections "⚠️ IMPORTANT: Prevents Excel Corruption" (`README.md:5-7@d674386`), "🎯 Overview" (organise, generate, **Refresh** Power Query, **Extract** cells, batch processing, prevent corruption, `:9-17`), "✨ Features" with Power Query support "(Windows only)", data extraction, dark mode, session memory (`:19-29`), "File Type Support" (six types; "GSTR-3B Export files (supports multiple monthly filings)", `:31-37`), "🔒 Excel Corruption Prevention" (`:39-47`).
- Requirements "Python 3.7 or higher (3.8+ recommended)", "Windows 10/11 (best), macOS, Linux", "Microsoft Excel 2016+ for Power Query features (Windows only)" (`:49-55`).
- Folder tree updated to the modular layout with `launch_extractor.py` and `power_query_extractor/` (no `gui/app.py`) (`:71-95`).
- A "Requirements.txt Contents" block (`:104-108`), a naming table with examples built on a generic placeholder company name (`:124-133`), usage for both programs (`:135-176`), output tree and file list (`:178-204`), Links cells (`:208-227`), an excerpt of `EXTRACTION_CONFIG` and `EXTRACTION_OPTIONS` (`:229-253`), a troubleshooting table including `pip install pywin32` and `pywin32_postinstall.py -install` (`:271-279`), "Debug Mode" (`:281-291`), "Known Limitations" (Excel files must be closed; templates need a specific structure; names must match exactly; Windows 260-character path limit; Power Query refresh needs Windows and Excel, `:312-318`), "📝 Recent Bug Fixes (Latest Commit)" (`:320-327`), "📄 License: This project is proprietary software for GST file management" (`:347-349`), footer "Version: 3.0 Production Ready / Status: Stable Release / Last Updated: 29th October 2025 / Author: Advanced AI Assistant" (`:361-364`).
- Removed: the "Version History v3.0" section and the old `gui/app.py` tree.

**README at d674386 compared with the code at d674386** (READ)

| README says | Code does |
|---|---|
| requirements.txt contains `openpyxl>=3.0.0` and `pywin32>=305; sys_platform == 'win32'` (`:104-108`) | the real file has `openpyxl>=3.1.2`, `lxml>=4.9.0` and no pywin32 - at every shared commit |
| "All file operations use temporary files and verification steps" (`:7`, `:42-43`), "Clean saves: No VBA macros or external links preserved" (`:44`), "Backup creation: Original files are never modified" (`:47`), "Multiple verification steps before file finalization" (`:310`) | same gaps as W under edac6a3: no temporary files, `EXCEL_SAFETY` unused, links kept; originals are indeed never modified (the program copies them), backups only in Fresh Run |
| `GST_Processing_Summary_DDMMYYYY_HHMMSS.xlsx`, `ITC_Report_ClientName_State_DDMMYYYY_HHMMSS.xlsx`, `Sales_Report_...` (`:182,188-189,195,202-203`) | `DDMMYY_HHMM` and the state **code**; the folder names `Annual Statement-DDMMYY HHMM`, `ClientName-StateCode`, `Version-DDMMYY HHMM` (`:181-184`) are now right |
| "The extractor auto-loads target folder from cache" (`:171`) | it never finds the organiser's cache (S under edac6a3) |
| Click **"Refresh & Extract"** (`:173`) | the button is "🚀 Start Processing" (`extractor_window.py:220-232@d674386`) |
| "Review variations tab for naming issues" (`:156`) | there is no variations tab |
| Dry Run described under the Processing tab (`:158-160`) | the DRY RUN button is on the Validation tab (`validation_tab.py:109-118`) |
| "Template not found → Ensure templates have 'Links' sheet" (`:278`) | "Template not found" is raised when the template file does not exist (`excel_handler.py:81-83`); a missing Links sheet is handled differently (I under edac6a3) |
| Debug Mode: add `logging.basicConfig(level=logging.DEBUG)` (`:283-287`) | `main.py` already calls `basicConfig` at INFO first; Python ignores a later `basicConfig` call once logging is set up, unless told to replace it |
| "Windows path limit of 260 characters applies" (`:317`) | the code has its own long-path handling from 218 characters (N under edac6a3) |
| "Recent Bug Fixes (Latest Commit)" (`:320-327`) | repeats 6237597's claims, which mostly do not match that diff (see 6237597); adds "Enhanced Excel corruption prevention mechanisms", which no commit did |
| Operating systems "macOS, Linux" (`:52`) | the extractor cannot start there (A under edac6a3) |

**Why it changed:** a documentation refresh after the code review earlier the same day, written by an AI assistant (footer "Author: Advanced AI Assistant") (HIGH).

**What it reveals:**
- The intended workflow: organise first, then refresh and extract with the second program; the extractor is now documented for the first time.
- The README describes the program as "Production Ready / Stable Release" while several statements do not match the code.
- This README text is the one both projects still carry: `git log -- README.md` on both main branches shows d674386 as its last change (READ, HIGH).
- No code, rule or dependency changed.

##### 5ea77c2 - 29-10-2025 21:20 - shared - before the split on 30-10-2025

**Commit message:** "Remove shebang line from main.py" (no body). (READ)

**What changed:** `main.py` −1 line: the first line `# !/usr/bin/env python3` is removed (READ: diff). Every later line of `main.py` moves up by one (for example the stdout wrapper 18 → 17, the "app.py" text 72 → 71).

**Why it changed:** the line was malformed (space after `#`) and has no meaning on Windows (MEDIUM: tidy-up).

**What it reveals:** nothing functional; the program is always started as `python main.py`. No rule or dependency changed.

##### 1cb878b - 29-10-2025 21:43 - shared - before the split on 30-10-2025

**Commit message:** "Add .gitignore file to exclude unnecessary files and directories" (no body). (READ)

**What changed:** new `.gitignore` +79 lines (a `.gitignore` lists files git should never add). Groups, quoted exactly (READ):
- Python build files: `__pycache__/`, `*.py[cod]`, `*$py.class`, `*.so`, `.Python`, `build/`, `develop-eggs/`, `dist/`, `downloads/`, `eggs/`, `.eggs/`, `lib/`, `lib64/`, `parts/`, `sdist/`, `var/`, `wheels/`, `*.egg-info/`, `.installed.cfg`, `*.egg`, `MANIFEST`.
- Virtual environments: `venv/`, `env/`, `ENV/`, `.venv`. Editors: `.vscode/`, `.idea/`, `*.swp`, `*.swo`, `*~`, `.DS_Store`.
- "Project-specific": `*.log`, `gst_organizer.log`, `pq_extractor.log`, `.gst_organizer_cache.json`, `*.cache`, `.cache/`, `temp/`, `logs/`.
- "Excel temporary files": `~$*.xlsx`, `~$*.xls`, `*.tmp`, `temp_gst_*`, `backup_*`.
- "Output directories (user data)": `/Annual Statement-*/`, `/output/`, `/target/`.
- "User data and sensitive information": `/mnt/user-data/`, `*.bak`, `config/last_session.json`.
- Test files: `test_files/`, `test_data/`, `*.test.xlsx`. Windows: `Thumbs.db`, `ehthumbs.db`, `Desktop.ini`. "Power Query cache": `*.odc`, `*.iqy`, `*.dqy`, `*.xlquery`. No line ending after the last line.

**Why it changed:** to stop logs and compiled files from being committed again; the compiled files were removed three minutes later (01862a0) (HIGH).

**What it reveals:**
- `/mnt/user-data/` is the upload and output folder of the sandbox used by Claude on claude.ai, so this file was most likely generated by that AI assistant (MEDIUM).
- Several entries are tied to this program: `temp/` and `logs/` (created by `main.py`), `temp_gst_*` and `backup_*` (the unused `EXCEL_SAFETY` prefixes - the real backups are named `<name>_backup_<time stamp>` and are **not** matched by `backup_*`), `.cache/` and `config/last_session.json` (two of the extractor's cache candidates), `.gst_organizer_cache.json` (the organiser's cache name, although that file lives in the home folder, not in the project). The extractor's own candidate `gst_organizer_cache.json` (without the dot) is not ignored. `lib/` would ignore any folder named "lib" anywhere (none exists). `.claude/` is not ignored, which is why bd176bf's `.claude/settings.local.json` was committed the next morning (READ, HIGH).
- Organised reports in a target folder outside the project are unaffected; only an `Annual Statement-*` folder at the project's top level is ignored.
- No code, rule or dependency changed.

##### 01862a0 - 29-10-2025 21:46 - shared - before the split on 30-10-2025

**Commit message:** "Remove __pycache__ directories from Git tracking" (no body). (READ)

**What changed:** the 70 compiled `.pyc` files are deleted from the repository (READ: `git show --name-status`: 70 × "D", nothing else; binary files, so no line counts; sizes 220 to 98,247 bytes).

**Why it changed:** follow-up to the new `.gitignore` (HIGH).

**What it reveals:** see Y under edac6a3 - the ten weeks of history before git, the never-committed single-file `gui/app.py`, two Python versions, editing on Windows, several project-folder locations. The files - with the developer computer's folder paths inside - stay readable in history from edac6a3 to 1cb878b (seven commits) (READ, HIGH). No code, rule or dependency changed.

---

#### Phase 3: Fixes and extensions driven by real use (29-10-2025 to 30-10-2025)

Late on 29-10-2025 and on 30-10-2025 the changes come from using the program: wrong row colours, a client list too long to scroll, Power Query refreshes needing different waits, a wish to refresh only one report, and new check sheets in the ITC template. The last of these commits is the split point.

##### 1c48722 - 29-10-2025 22:18 - shared - before the split on 30-10-2025

**Commit message:** "Remove redundant text color handling in client tree display" (no body). (READ)

**What changed:** `gui/handlers/client_handler.py` −6 lines (READ: diff). Inside `update_client_tree()` this block is removed (`client_handler.py:54-59@edac6a3`):

```python
# Ensure text color is correct
if not self.app.dark_mode.get():
    self.app.client_tree.item(item, tags=('complete',))
    # Force text to be black in light mode
    for col in self.app.client_tree['columns']:
        self.app.client_tree.set(item, col, self.app.client_tree.set(item, col))
```

The `else:` that followed it now belongs to `if client_info['status'] == 'Complete':`. Result: complete clients get the tag "complete", all others "incomplete" (`client_handler.py:51-55@1c48722`). The removed loop wrote each cell's own value back into it and did nothing.

**Why it changed:** the rows were coloured wrongly (HIGH, RAN below). The message calls it "redundant", which understates a visible fix.

**What it reveals:**
- What was broken - RAN (`test_row_tags.py`, real `ClientHandler`, one complete and one incomplete made-up client): at edac6a3, light mode tagged **both** rows "complete" (incomplete clients looked light green, like complete ones) and dark mode tagged both "incomplete"; at 1c48722 both modes give "complete" and "incomplete" correctly. A real-use observation that complete and incomplete clients looked the same.
- How it was fixed: properly (the overriding second `if/else` removed).
- Similar case still open: the theme switch itself does not restore the light row colours, and after a re-scan the text can stay white on light rows (L under edac6a3; RAN identical at 47a1492). Commit 1c48722 fixed the colour **choice**, not the colour **values** after a theme switch.
- UI rule now actually applied: "Complete = green row, otherwise orange row". No GST rule or dependency changed.

##### 4c42758 - 29-10-2025 22:25 - shared - before the split on 30-10-2025

**Commit message:** "Add mouse wheel scrolling support for main and log text areas" (no body). (READ)

**What changed:** `power_query_extractor/gui/extractor_window.py` +30 lines (READ: diff).
- In the Setup tab, after the client-list canvas: a function `_on_mousewheel()` scrolls the list by `-1*(event.delta/120)` units, but only if the list is taller than the visible area, and returns "break"; it is bound to the canvas and to the inner frame. Moving the mouse into the canvas (`<Enter>`) switches on a window-wide wheel binding (`bind_all`), leaving it (`<Leave>`) switches it off again (`unbind_all`), so the wheel also works over the client checkboxes.
- In the Processing tab: `_on_log_mousewheel()` scrolls the log by the same amount if the log is not empty, bound to the log box.

**Why it changed:** the extractor's client list (a canvas full of checkboxes) did not scroll with the mouse wheel (HIGH). The "main" area in the message is the client list.

**What it reveals:**
- The extractor was being used with enough clients to need scrolling (MEDIUM).
- `<MouseWheel>` with `delta/120` is the Windows/macOS wheel event; Linux sends different events, so there the wheel still does nothing (READ, HIGH; irrelevant on Windows).
- `unbind_all("<MouseWheel>")` on leaving removes every window-wide wheel binding in this window; the extractor has no other, so nothing is lost. The organiser's own window-wide wheel binding (`setup_tab.py:63-66`) belongs to a different program and window and is not affected (READ, HIGH).
- Not sure: when the pointer moves from the canvas onto a checkbox inside it, Tk may report a "leave" for the canvas, which would switch the wheel off again exactly over the checkboxes. How to check: on Windows, open the extractor with more clients than fit, rest the pointer on a client's checkbox and turn the wheel.
- No rule or dependency changed.

##### bd176bf - 30-10-2025 09:04 - shared - before the split on 30-10-2025

**Commit message:** "Enhance report processing UI with configurable wait time and file suffix options; add ITC/Sales report selection functionality" (no body). (READ)

**What changed - 3 files, +412 / −162** (READ: full diff and the files at bd176bf/47a1492, which are identical for these files)

1. **`.claude/settings.local.json`** +9 (new; shown here on one line):

   ```json
   {"permissions": {"allow": ["Bash(python -m py_compile:*)"], "deny": [], "ask": []}}
   ```

   This is the personal settings file of **Claude Code** (Anthropic's AI coding tool). It lets the tool run `python -m py_compile <file>` (a syntax check that compiles a file without running it) without asking each time. Nothing else is allowed; no password, token, key, path or personal data is in it (READ, HIGH). By that tool's convention a "settings.local.json" is personal and normally not committed; it was committed because `.claude/` is not in the `.gitignore`. Its presence shows that Claude Code was used for this commit's changes (MEDIUM-HIGH).

2. **`power_query_extractor/core/report_processor.py`** +41 / −17:
   - `__init__` adds `self.wait_time = 10` and `self.suffix_pattern = "_Refreshed_{timestamp}"` (`:27-28@47a1492`).
   - `process_client()` now honours `client_data['process_itc']` and `['process_sales']` (both default to True "for backward compatibility"); a report that is not selected gets the status `{'success': False, 'error': 'Not selected for processing'}` with no data (`:39-70`).
   - `_process_report()` builds the refreshed name as `<report name>` + the suffix, with `{timestamp}` replaced by day-month-year_hour-minute (`:91-94`).
   - In `_refresh_power_query_simple()`: the first wait after the key presses becomes `self.wait_time` seconds (was a fixed 10; log "Waiting... i/N seconds") (`:256-261`); the "extended wait" becomes `max(10, self.wait_time)` seconds (was 10) (`:283-289`). The five one-second error-dialog checks in between are unchanged. Some blank lines tidied.

3. **`power_query_extractor/gui/extractor_window.py`** +362 / −145 - the Setup tab is rebuilt in two columns:
   - **Left column** (fixed 400 pixels, `:101-104@47a1492`): Target Folder (label, entry, 📁 Browse, 🔍 Scan, `:106-157`); new "Processing Options" (`:159-236`): "Power Query Wait Time" spin box 5 to 60, default 10, label "seconds (Min: 5, Max: 60)" (`:170-198`) with no check of typed values; "Refreshed File Suffix" entry, default `_Refreshed_{timestamp}`, hint "Use {timestamp} for datetime" (`:200-223`) with no check; the "Skip Refresh (Use existing refreshed files)" checkbox moved here (`:225-236`); "🚀 Start Processing" at the bottom (`:238-251`).
   - **Right column** "Select Clients & Reports" (`:253-367`): buttons ✓ Select All, ✗ Deselect All, **ITC Only**, **Sales Only**; header row "Client Name | ITC | Sales | Last Refresh Status"; per client a name, an ITC checkbox and a Sales checkbox (ticked when that report exists, greyed out when it does not) and a status text (`display_clients()`, `:552-636`).
   - `get_refresh_status()` (`:638-660`): the newest `ITC_Report_*_Refreshed_*.xlsx` and `Sales_Report_*_Refreshed_*.xlsx` by modification time, shown as `YYYY-MM-DD HH:MM`, otherwise "Never"; `format_refresh_status()` (`:662-671`) gives "ITC: ... | Sales: ..." or "Never refreshed".
   - `start_processing()` (`:708-746`) collects clients with at least one ticked report, adds `process_itc` / `process_sales`, passes wait time and suffix to the processor; message "Please select at least one report to process". The worker (`:748-832`) logs the wait time and suffix, "Processing <name> - Reports: ITC, Sales" and "⏭️ ... Skipped" for unticked reports.
   - Removed: the single checkbox per client and the "ITC ✓ / Sales ✓" labels of edac6a3.

**Why it changed:** needs from real use (HIGH, from the content): some Power Query refreshes needed more (or less) time than the fixed 10 s; some clients needed only the ITC or only the Sales report refreshed; the author wanted another name for refreshed copies and to see when each report was last refreshed. Skip Refresh itself is **not** new - it existed at edac6a3 (`extractor_window.py:188-198,234-243@edac6a3`, `report_processor.py:73-93@edac6a3`); this commit only moves the checkbox.

**What it reveals:**
- **Business and technical rule changes** (none of them a GST rule): wait after the key presses "10 s + about 5 s of checks + 10 s" → "N s (default 10, 5-60 by the spin box) + about 5 s of checks + max(10, N) s"; refreshed-copy name `_Refreshed_<DDMMYY_HHMM>` → a user pattern (same default); which reports are processed: always both → per client ITC and/or Sales. [GST RULE] display: in the consolidated workbook a report that was not selected shows ✗ and "Missing" for all its values (`data_consolidator.py:110-129`).
- The fixed-wait key-press refresh (N under edac6a3) was the main pain point; the author tuned waits rather than detecting when the refresh finishes.
- **New mismatches introduced** (READ, HIGH; RAN where marked):
  1. **A custom suffix is not honoured by the code that looks for refreshed copies.** The Skip Refresh lookup still searches `<name>_Refreshed_*.xlsx` (`report_processor.py:103@47a1492`) and the status column searches `*_Refreshed_*` (`extractor_window.py:645,652@47a1492`). RAN (`test_suffix.py`, Windows modules stubbed, Excel steps replaced): with the suffix `_R_{timestamp}` a first run creates `..._R_<time>.xlsx`; a second run with Skip Refresh on does not find it and refreshes again; the status column would show "Never refreshed".
  2. **An empty suffix** makes the refreshed name equal to the original; the copy step fails with "... are the same file" and the report is marked failed (RAN). **A suffix without `{timestamp}`** makes every run overwrite the same refreshed file (RAN: `..._Done.xlsx`). Characters not allowed in Windows file names would make the copy fail (MEDIUM, not run).
  3. **The wait-time box accepts any typed value**: 0 or a negative number gives no first wait; a large number blocks the run that long; text that is not a number makes the start button fail with a Tk error and nothing starts (READ, MEDIUM).
  4. The wildcard `ITC_Report_*.xlsx` used to find the original report (`report_processor.py:76@47a1492`) also matches refreshed copies with any suffix, so after several runs the "original" may be a refreshed copy (READ, LOW-MEDIUM).
- Abandoned: the single "select client" checkbox per client.
- Dependencies: none (no new imports). The settings file concerns a development tool, not the running program.
- The message is accurate but does not mention the `.claude/settings.local.json` file.

##### 47a1492 - 30-10-2025 14:38 - shared - before the split on 30-10-2025

**Commit message:** "Add new ITC report mappings for Reco Diff, Table 6J Match, Table 8D Limits, IMS Reco, Extra In Purchase, and Extra In GSTR2B sheets" (no body). Author and committer recorded as "Sidhant" with the same e-mail address as "goyal-sidhant" - the only shared commit under this name, which suggests a different computer or git program for this one commit (LOW-MEDIUM). This is the **split point**: the last commit both projects share. (READ)

**What changed - 2 files, +126 / −63** (READ: full diff)

1. **`power_query_extractor/config/cell_mappings.py`** +126 (152 → 278 lines): six new output sheets in `EXTRACTION_CONFIG`, all for the ITC report and all without a calculation. [GST RULE] ⚠ (template sheet names and cell positions; the GSTR-9 table numbers in the names belong to the return format and can change):

   | New output sheet | Template sheet!cell → output column | Lines @47a1492 |
   |---|---|---|
   | Reco Diff | `Reco`!C37 → "IGST Diff"; !D37 → "CGST Diff"; !E37 → "SGST Diff" | 105-127 |
   | Table 6J Match | `T6 of R9`!K29 → "Table 6J Match" | 130-140 |
   | Table 8D Limits | `T8 of R9`!G6 → "Table 8D Limits" | 143-153 |
   | IMS Reco | `IMS_Reco`!AA1 → "IGST"; !AD1 → "CGST"; !AG1 → "SGST" | 156-178 |
   | Extra In Purchase | `ExtraInPurchase`!O1 → "IGST"; !P1 → "CGST"; !Q1 → "SGST" | 181-203 |
   | Extra In GSTR2B | `ExtraInGSTR2B`!S1 → "IGST"; !T1 → "CGST"; !U1 → "SGST" | 206-228 |

   The sheet names "T6 of R9" / "T8 of R9" and the column names refer to GSTR-9 Tables 6J and 8D ⚠ - the author's domain; the code only copies whatever those cells contain.

2. **`power_query_extractor/config/cell_mappings copy.py`** −63 (deleted): the old-format mapping of 21-06-2025 (P under edac6a3).

**Why it changed:** the ITC template had gained new check sheets (Reco, T6 of R9, T8 of R9, IMS_Reco, ExtraInPurchase, ExtraInGSTR2B) and the author wanted their key cells in the consolidated workbook (HIGH, from the content).

**What it reveals:**
- The ITC template's sheet names and check cells as of 30-10-2025; the consolidated workbook now has 9 data sheets plus Summary (22 cells in total).

> **Wrong values in three of the new sheets (RAN, HIGH).** The three new sheets IMS Reco, Extra In Purchase and Extra In GSTR2B have columns named exactly `IGST`, `CGST`, `SGST` - without a space. The consolidator's substring test (`data_consolidator.py:141-145`, shown under Q of edac6a3) finds `IGST` inside the first field name that contains it, `GSTR-3B_Difference_Diff_IGST`, before the exact lookup is ever tried. So the lookup puts **`Diff GSTR-3B`!J15, K15 and L15** into the IGST, CGST and SGST columns of all three sheets, although the mapping names `IMS_Reco`!AA1/AD1/AG1, `ExtraInPurchase`!O1/P1/Q1 and `ExtraInGSTR2B`!S1/T1/U1 - a technical mismatch between the configured cells and the cells shown; which cells are right is the author's call. RAN (`test_consolidator.py`, every value set to its own "sheet!cell" label): at 47a1492 those three sheets show `Diff GSTR-3B!J15`, `!K15`, `!L15`; the other six sheets show their own cells. The same run on copies of the latest snapshots of **both** projects gives the same wrong result. If a report has no `Diff GSTR-3B` sheet, its fields are empty, the exact lookup is reached, and the right cells appear (READ).

- Similar cases: any future column name that also occurs inside an earlier field name would be affected the same way (READ, HIGH).
- Rule change: [GST RULE] ⚠ 14 new extraction cells as listed; no existing cell changed. Removed: the dormant old-format mapping file.
- No dependency changed. The message lists the six sheets correctly but does not mention deleting the copy file.

##### Where the two projects started from (47a1492)

At 30-10-2025 14:38 both projects held exactly this program (READ: `git ls-tree` and the files at 47a1492; line counts are git's):

- **Files**: 36 Python files with 7,295 lines (one fewer than at edac6a3: `cell_mappings copy.py` is gone), plus `README.md` (364 lines, the d674386 text), `requirements.txt` (13, unchanged since edac6a3), `.gitignore` (79) and `.claude/settings.local.json` (9). No tests, no `docs/`, no `models/`, no `pytest.ini`. Changed since edac6a3: `main.py` (82), `core/file_parser.py` (371), `utils/helpers.py` (537, import line only), `gui/handlers/client_handler.py` (253), `gui/handlers/file_handler.py` (151, one line), `gui/handlers/processing_handler.py` (463), `power_query_extractor/core/report_processor.py` (760), `power_query_extractor/gui/extractor_window.py` (848), `power_query_extractor/config/cell_mappings.py` (278). Every other Python file is exactly as in edac6a3.
- **Organiser** (`python main.py`, Tkinter, window "GST File Organizer v3.0"): the same three screens and features as at edac6a3, except: Dry Run and the first part of the processing log go through `root.after` with the late-binding bug (`gui/handlers/processing_handler.py:36-52,104,178-239`); the long-name warning before processing uses the max-length setting (`:64-91`) while the per-client toggle and the Setup label still say 10; client rows are coloured by completeness (`gui/handlers/client_handler.py:51-55`).
- **Power Query Extractor** (`python launch_extractor.py`, Tkinter, Windows + pywin32 + Excel only): Setup tab with Target Folder, Processing Options (wait 5-60 s, default 10; refreshed-file suffix, default `_Refreshed_{timestamp}`; Skip Refresh), Start Processing, and a per-client ITC / Sales selection with "Last Refresh Status"; Processing tab with progress and a wheel-scrollable log. Per selected report: copy to the refreshed name → visible, maximised Excel → Alt+A, R, A → wait N s → five one-second error-dialog checks → wait max(10, N) s → final check → Info!BB2 test → save and close → wait 3 s → hidden Excel reads the configured cells → `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` in the newest Annual Statement folder.
- **Unchanged rules and formats since edac6a3**: the six file-name patterns and their order [GST RULE] ⚠; the six expected files with the GSTR-3B exception [GST RULE] ⚠; the Links cell map [GST RULE] ⚠; the 39-entry state list and its fallback [GST RULE] ⚠; the Info!BB2 test [GST RULE] ⚠; "Missing" and `#,##0.00` display [GST RULE]; Private/Limited abbreviations; 35-character folder key; all folder and file-name formats; the settings cache `~/.gst_organizer_cache.json` versus the extractor's relative `gst_organizer_cache.json`; the logs `./gst_organizer.log` and `./pq_extractor.log`.
- **Extraction map at the split point** [GST RULE] ⚠ (`power_query_extractor/config/cell_mappings.py:33-242@47a1492`): GSTR-3B Difference (ITC: `Diff GSTR-3B`!J15/K15/L15); GSTR-1 Difference (Sales: `Diff - GSTR-1`!N1/O1/P1/Q1); Turnover GSTR-1 (Sales: `Diff - GSTR-1`!I1); Reco Diff (ITC: `Reco`!C37/D37/E37); Table 6J Match (ITC: `T6 of R9`!K29); Table 8D Limits (ITC: `T8 of R9`!G6); IMS Reco (ITC: `IMS_Reco`!AA1/AD1/AG1 - shown with `Diff GSTR-3B`!J15/K15/L15 by the lookup bug); Extra In Purchase (ITC: `ExtraInPurchase`!O1/P1/Q1 - same bug); Extra In GSTR2B (ITC: `ExtraInGSTR2B`!S1/T1/U1 - same bug).
- **Known defects inherited by both projects**: all of V under edac6a3 except number 4 (fixed by 1c48722), with number 2 now active and wrong; the 6237597 late-binding bug and its partial threading change; the bd176bf suffix/lookup mismatch and the unchecked wait-time and suffix inputs; the Excel-through-COM practices (J and N under edac6a3); the console and file-location behaviour (A); the dead code (U); the README statements that do not match the code (d674386); `main.py:71` still printing "app.py".
- **Data in history**: the run log (edac6a3) and the 70 compiled files with the developer computer's folder paths (edac6a3 to 1cb878b) are ancestors of both projects' latest commits.
- **Tooling and authorship at the split**: AI-assisted throughout - an AI-written README and `.gitignore` (d674386, 1cb878b), a Claude Code settings file (bd176bf). Every later commit in either project starts from this state; the two histories are described separately from Phase 4 onward.
