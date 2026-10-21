# H-S — Working notes: the 12 SHARED commits (edac6a3 .. 47a1492)

Label for every commit in this file: **shared - before the split on 30-10-2025**. These notes feed BOTH audits (A = AnnualStatement_oldv1, B = AnnualStatement_oldv2).

## How to read these notes

- **Citations** are `repository-relative-path:line@commit`, for example `core/file_parser.py:79@47a1492`. The line number is the line in that commit's version of the file. "@edac6a3" = the first commit (the original v3.0 program). "@47a1492" = the split point. "A HEAD" = bf2f1dd (27-07-2026), "B HEAD" = b0e7d0e (09-02-2026). For files that did not change between edac6a3 and 47a1492 the line numbers are the same in both.
- **READ** = I read the code or diff. **RAN** = I executed something (always on copies under `/root/merge-scratch/run/H-S/`, never in A-copy, B-copy or /home/user). Inferences carry **HIGH / MEDIUM / LOW** confidence.
- **[GST RULE]** = anything deciding a tax figure, which return/table/sheet/cell a figure belongs to, rates/limits, rounding, period/financial year, GSTIN/PAN/state-code logic, ITC, due dates. I only describe what the code does; I never judge correctness. **TIME-VARYING** = a value that can change over time (table numbers, template cell positions, state lists, etc.).
- Client confidentiality: no client name, GSTIN, PAN, real file name or figure appears below. Examples use `<Client A>` or obviously generic names ("Alpha Traders", "ABC Ltd").
- Scratch scripts I wrote and ran (all under `/root/merge-scratch/run/H-S/`): `pyc_hdr.py`, `pyc_paths2.py` (compiled-file metadata, prints no paths), `logshape.py` (log message counts by code template, prints no log content), `test_dryrun.py`, `test_tags.py`, `test_consolidator.py`, `test_helpers.py`, `test_scan.py`, `test_folder_setting.py`, plus extracted trees `edac6a3/`, `47a1492/`, `c6237597/`, `c1c48722/`, `smoke47/`, `headA/`, `headB/`, `headA_utils/`.

### Plain-English glossary (terms used below)

- **Commit**: one saved change in git history. **Diff**: the exact lines added/removed by a commit. **Merge commit**: a commit that joins two lines of work. **HEAD**: the latest commit of a line of work. **Split point**: the last commit both projects share (47a1492).
- **Regex / pattern**: a text rule used to recognise file names (for example "starts with `GSTR3B-`").
- **GUI**: the program's windows and buttons. **Tkinter (Tk)**: Python's built-in toolkit for windows (used by A and by every shared commit). **PyQt5**: another window toolkit (used by B after the split).
- **Thread / worker thread**: a second line of execution so the window stays responsive while files are processed. Tk windows must only be changed from the main thread.
- **Callback**: a function handed to another piece of code to be called later. **lambda / closure**: a tiny inline function; a "late-binding" bug happens when such a function reads a loop variable later, after the loop has moved on, so it sees the wrong (last) value.
- **COM / pywin32 / win32com**: the Windows mechanism (and the Python package) that lets a program drive Microsoft Excel. **pythoncom.CoInitialize / CoUninitialize**: start/stop COM for the current thread.
- **openpyxl**: a Python package that reads/writes .xlsx files without Excel. It drops Power Query queries when it saves (Standard 12).
- **Power Query**: Excel's data-import engine; the ITC/Sales templates use it to pull data from the organised files.
- **SendKeys**: simulating keyboard presses into whatever window is in front.
- **Cache file**: a small settings file remembering last-used folders. **JSON**: the text format it uses.
- **glob**: a wildcard file search such as `ITC_Report_*.xlsx`. **mtime**: a file's "last modified" time.
- **.pyc / __pycache__**: compiled copies of Python files that Python writes automatically; they record the source file's size, modification time and full folder path.
- **CRLF**: Windows line endings (two bytes per line end), LF: Linux/git line endings (one byte).
- **Treeview**: the table-like client list on the Validation screen.
- **Dead code**: code that is never called.

---

## 0. Most important findings in this range (summary; details below)

1. **Wrong cells in the consolidated extraction report (RAN, HIGH)**. `data_consolidator.py:142-145@47a1492` looks up each output column by a *substring* match on field names. 47a1492 added three output sheets whose columns are just `IGST` / `CGST` / `SGST` (`cell_mappings.py:162,168,174,187,193,199,212,218,224@47a1492`). The substring `IGST` first matches the field of the *GSTR-3B Difference* sheet, so the "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" output sheets are filled with the values of `Diff GSTR-3B!J15/K15/L15` instead of their own configured cells. RAN with synthetic labelled values at 47a1492, A HEAD and B HEAD: same wrong result in all three. Technical bug, not a judgement of the mapping itself.
2. **6237597 introduced a late-binding regression** (RAN, HIGH): the Dry Run preview prints the LAST client's name and LAST file name for every line (`gui/handlers/processing_handler.py:36-52@6237597`). The same pattern is in the worker-thread logging (lines 178-239). A HEAD fixed it (`gui/handlers/processing_handler.py:23-31` A HEAD, `_log_async` captures the message at schedule time); B replaced the Tk code with PyQt.
3. **6237597's commit message does not match its diff** (READ, HIGH): it claims a main.py fix that is not in the diff (main.py still prints `app.py` at `main.py:71@47a1492`), claims an "attribute access error" that could not occur, and claims an indentation fix that only removed blank-line spaces.
4. **1c48722 is really a bug fix, not "redundant code removal"** (RAN, HIGH): before it, in light mode every client row was coloured as "complete" (even incomplete ones) and in dark mode every row as "incomplete".
5. **Client data sits in both repositories' history** (READ, HIGH): `gst_organizer.log` (5,689 lines; client names, state names, source file names, Windows folder paths) is in edac6a3, which is an ancestor of both HEADs; and 70 compiled `.pyc` files (edac6a3 to 1cb878b) embed the developer computer's folder paths.
6. **Pre-git history existed and is lost** (READ of pyc headers, HIGH): the `.pyc` headers show source edits from 19-06-2025 to 23-07-2025 and a monolithic `gui/app.py` (89,275 bytes, last edited 25-06-2025) that was replaced on 23-07-2025 by the modular GUI (main_window + tabs + handlers + utils + widgets) and never committed.
7. **Several original (edac6a3) defects survive to the split point**: extractor cannot find the main app's cache (different file name and folder); extractor picks the "latest" version folder by name `Version-DDMMYY HHMM`, which is not date order (RAN); the per-client "include client name in folder" toggle is stored under a key the organiser never looks up (RAN); bound START button still fires when "disabled" (RAN) and there is no "already processing" guard; numeric state codes in file names crash client selection (RAN at function level); `main.py` crashes when started without a console (RAN simulation; relevant to Standard 13). HEAD status of each is listed in section 1.13.
8. **Excel through COM in the original code differs from Standard 12 in several points** (READ): the main app attaches to the user's already-open Excel (`GetActiveObject`), hides it, and always calls `Quit()` on it; the extractor refreshes Power Query by sending keystrokes Alt-A, R, A to a visible maximised Excel and then waits fixed seconds (no `RefreshAll()` / `CalculateUntilAsyncQueriesDone()`).
9. **bd176bf's configurable "refreshed file suffix" is not honoured by the code that looks for refreshed files** (READ, HIGH): `report_processor.py:103@47a1492` and `extractor_window.py:645,652@47a1492` still search for the hard-coded `_Refreshed_` text. Fixed in A HEAD (`power_query_extractor/core/refresh_naming.py:61` describes the replacement); still present in B HEAD (`power_query_extractor/core/report_processor.py:102`).

---

## 1. The original program at edac6a3 (31-08-2025): "GST File Organizer v3.0"

### 1.1 What edac6a3 added (110 files, 12,916 lines)

- 37 Python files with 6,954 lines, README.md (260 lines), requirements.txt (13 lines), `gst_organizer.log` (5,689 lines, see 1.15) and 70 compiled `.pyc` files (see 1.16). (READ: `git show --stat edac6a3`, `--numstat`.)

| Path | Lines | Role (tag) |
|---|---|---|
| `main.py` | 83 | Entry point of the organiser [CORE/entry] |
| `launch_extractor.py` | 30 | Separate entry point for the Power Query Extractor [entry] |
| `requirements.txt` | 13 | `openpyxl>=3.1.2`, `lxml>=4.9.0`, dev tools commented out [CONFIG] |
| `README.md` | 260 | User guide (written for an older layout, see 1.17) [DOCS] |
| `core/__init__.py` | 13 | Re-exports FileParser, FileOrganizer, ExcelHandler; `__version__ = '3.0.0'` [CORE] |
| `core/file_parser.py` | 373 | Scans a folder, recognises GST file names, groups files per client, checks completeness [CORE] |
| `core/file_organizer.py` | 409 | Creates output folders, copies files, writes a per-client text report [CORE] |
| `core/excel_handler.py` | 661 | Creates ITC/Sales reports from templates (Excel COM or openpyxl) and the summary workbook [CORE] |
| `utils/__init__.py` | 38 | Re-exports constants and helpers; `__version__ = '3.0.0'` [UTIL] |
| `utils/constants.py` | 275 | File patterns, folder names, template cell map, GUI colours, modes, messages, Excel safety values, state-name map [CONFIG] |
| `utils/helpers.py` | 537 | Timestamps, paths, filename cleaning, copy/backup, Excel signature check, state code, client-state key [UTIL] |
| `gui/__init__.py` | 11 | Re-exports GSTOrganizerApp [UI] |
| `gui/main_window.py` | 365 | Main window; wires tabs, handlers, cache, dark mode [UI] |
| `gui/tabs/{__init__,setup_tab,validation_tab,processing_tab}.py` | 8 / 377 / 247 / 117 | The three screens [UI] |
| `gui/handlers/{__init__,file_handler,client_handler,processing_handler}.py` | 8 / 151 / 259 / 447 | Button logic: browse/scan, client selection, processing thread [UI] |
| `gui/utils/{__init__,cache_manager,dark_mode_manager,status_bar,ui_helpers}.py` | 9 / 56 / 212 / 37 / 63 | Settings cache, theme switch, status bar, log helpers [UI] |
| `gui/widgets/{__init__,collapsible_frame,title_bar}.py` | 7 / 41 / 51 | Small widgets [UI] |
| `power_query_extractor/__init__.py` | 3 | `__version__ = "1.0.0"` [CORE] |
| `power_query_extractor/extractor_main.py` | 79 | Extractor start-up class, logging to `pq_extractor.log`, cache lookup [CORE/entry] |
| `power_query_extractor/config/__init__.py`, `core/__init__.py`, `gui/__init__.py` | 1 each | Docstring only |
| `power_query_extractor/config/cell_mappings.py` | 152 | Which report cells to extract, options [CONFIG, GST RULE] |
| `power_query_extractor/config/cell_mappings copy.py` | 63 | Older format of the same mapping (file name contains a space; cannot be imported) [DEAD/ABANDONED] |
| `power_query_extractor/core/report_processor.py` | 736 | Opens each report in Excel, refreshes Power Query by keystrokes, validates, reads cells [CORE] |
| `power_query_extractor/core/data_consolidator.py` | 432 | Writes the consolidated extraction workbook [CORE] |
| `power_query_extractor/gui/extractor_window.py` | 601 | Extractor window (Setup + Processing tabs) [UI] |

(Line counts are git's per-file counts; `wc -l` shows one less where the last line has no newline.)

### 1.2 Entry points and start-up (READ)

- **Organiser**: `python main.py`. `main.py:18@edac6a3` replaces `sys.stdout` with a UTF-8 wrapper of `sys.stdout.buffer`; `main.py:21-28` configures logging at INFO level to `gst_organizer.log` (relative path = whatever folder the program is started from) and to the console; `setup_environment()` (`main.py:30-44`) puts the project folder on `sys.path`, **creates folders `core`, `utils`, `gui`, `logs`, `temp` in the project folder**, and writes a one-line `__init__.py` into core/utils/gui if missing; `main()` (`main.py:46-80`) imports `gui.main_window.GSTOrganizerApp`, prints a banner, runs the window. On ImportError it prints a list of expected files including **`gui/app.py`** (`main.py:72`) and waits for Enter (`input`, line 73); on any other error it logs with traceback, prints, waits for Enter, exits 1. The first line `# !/usr/bin/env python3` (`main.py:1`) is a malformed shebang (space after `#`), so it was never effective (HIGH).
  - RAN (split-point copy, Linux, Python 3.14, virtual display): the window builds and closes without error; `logs/` and `temp/` folders and an empty `gst_organizer.log` appear in the project/start folder; nothing is written to the home folder at start-up.
  - RAN (simulation): with `sys.stdout = None` (what `pythonw.exe` gives a program on Windows, Standard 13) `main.py` stops at the stdout wrapper with `AttributeError: 'NoneType' object has no attribute 'buffer'` (`main.py:17@47a1492` = line 18@edac6a3). Still unguarded at A HEAD `main.py:17`; guarded at B HEAD `main.py:12-13` (`if hasattr(sys.stdout, 'buffer')`).
- **Extractor**: `python launch_extractor.py` (`launch_extractor.py:10` appends the script folder to `sys.path`; `:14-18` imports `power_query_extractor.extractor_main.PowerQueryExtractor` and runs it; on any exception it shows a Tk error box and re-raises, `:20-27`). `extractor_main.py` also has its own `main()` (`:72-79`). At import, `extractor_main.py:12-19` configures logging to `pq_extractor.log` (relative path) and the console — but only if no logging was configured before.
  - RAN: on Linux the import fails at `power_query_extractor/core/report_processor.py:11` (`ModuleNotFoundError: No module named 'pythoncom'`), because `report_processor.py:11-15` imports pywin32 modules unconditionally. The extractor is therefore Windows-only with pywin32 installed; pywin32 is not in requirements.txt.

### 1.3 Features of the original program (READ unless marked)

**Organiser (main.py) — three tabs: "📁 Step 1: Setup", "✅ Step 2: Validation", "🚀 Step 3: Processing"** (`gui/tabs/setup_tab.py:22`, `validation_tab.py:21`, `processing_tab.py:20`).

- **F1 Setup screen** (`gui/tabs/setup_tab.py`): source folder, ITC template, Sales template, target folder (entries + Browse buttons; file filter `*.xltx *.xlsx`, `gui/handlers/file_handler.py:31,42`), processing-mode radio buttons (Fresh Run / Re-Run / Resume from `PROCESSING_MODES`), checkbox "Include client name in brackets for Level 4 folders" (`setup_tab.py:229-235`), label "⚠️ Names >10 chars will show warning / ℹ️ This overrides individual client settings if checked" (`:237-242`), spinbox "Client folder name max length" from 15 to 100 (`:271-277`) whose validator allows 10–100 (`:252-260`, inconsistent with the 15 lower bound and the label "(default: 35, range: 15-100)" `:279`), SCAN FILES and RE-SCAN buttons (`:297-318`), a help column listing the six expected name patterns (`:362-369`). Mouse-wheel scrolling bound globally with `bind_all("<MouseWheel>")` (`:64-66`, Windows-style wheel event only). The action buttons of the three screens (Browse, Scan, Re-scan, Dry Run, Export, Start, Stop, Select All/Clear All/Complete Only) are wired with `bind('<Button-1>', ...)` instead of `command=` (for example `setup_tab.py:115`, `validation_tab.py:118,130,142,190`, `processing_tab.py:74,87`); comments call this "working mouse events"/"proper event binding" (`setup_tab.py:20,103`), which hints at an earlier problem with `command=` (LOW). Only the collapsible-frame arrow uses `command=`.
- **F2 Scan** (`gui/handlers/file_handler.py:57-95` → `core/file_parser.py:128-179`): requires a source folder that exists (`file_handler.py:120-130`); `find_excel_files()` (`utils/helpers.py:249-266`) takes only the top level of the folder (no sub-folders), patterns `*.xlsx`, `*.xls`, `*.xlsm`, and **silently drops** files that fail `validate_excel_file()` (smaller than 1,024 bytes, or extension not one of .xlsx/.xls/.xlsm/.xltx/.xltm, or first bytes not the ZIP `PK\x03\x04` or OLE `\xd0\xcf\x11\xe0` signature; `helpers.py:210-247`). Each remaining file name is matched against the six patterns in order (1.4). Matches with client and state go into `client_data[<client>-<state code>]`; everything else goes into "variations" with a suggested pattern (`file_parser.py:290-308`). Completeness is computed per client (`file_parser.py:254-288`). The summary text (`gui/main_window.py:267-301`) and client tree are refreshed and the window switches to Validation.
  - RAN (split-point parser, synthetic generic names): a client with all six types plus a second GSTR-3B file (including a `(2) GSTR3B-...` prefix) → "Complete", 100.0 %; a client with two Sales files → "Missing 5 files" and extras "Multiple Sales (2 files)"; a client name containing a hyphen, an `.xlsm` file and a random name → variations "No matching pattern found"; a 104-byte file named like a GSTR-3B file → not listed anywhere (silently dropped); a file ending `.XLSX` (upper case) → not found on Linux (glob is case-sensitive there; on Windows it is case-insensitive).
- **F3 Validation screen** (`gui/tabs/validation_tab.py`): collapsible instructions, scan summary, buttons DRY RUN, EXPORT LIST, START (`:109-142`), client tree with columns Client, State, Status, Files, Missing, Extra, FolderName ("📁 Name") (`:197-221`), Select All / Clear All / Complete Only (`:174-190`), keyboard ↑ ↓ and SPACE (`:241-244`). Clicking the tree's first column toggles ☐/☑; clicking column 7 toggles the per-client "include name" setting (`client_handler.py:73-85`).
- **F4 Dry run** (`processing_handler.py:23-52@edac6a3`): lists, per selected client, each file "Would organize" and "Would create: ITC Report / Sales Report". Writes nothing.
- **F5 Processing** (`processing_handler.py:54-447@edac6a3`), in a background thread, per selected client: create folders (F6) → copy files with cleaned names (F7) → ITC report → Sales report (F8) → per-client text report → after all clients a summary workbook (F9) → info box → open the output folder (`os.startfile` on Windows, `open` on macOS, `xdg-open` on Linux, `:425-437`). Stop is checked only between clients (`:150-151`). Optional warning if "include client name" is on and any selected client name is longer than 10 characters (`:64-75@edac6a3`).
- **F6 Folder structure** (`core/file_organizer.py:67-179`): see 1.8.
- **F7 Copy with cleaned names** (`file_organizer.py:181-298`): each file is copied (`shutil.copy2`, then size compared; a size mismatch deletes the copy, `helpers.py:142-185`) to its category folder under a name cleaned by `sanitize_filename()` (1.9); the in-memory file record is updated to the cleaned name so the template gets the copied name (`file_organizer.py:216-229`). If the destination already exists: Fresh mode makes a backup copy `<stem>_backup_<DDMMYY_HHMM><ext>` first (`:260-265`, `helpers.py:187-204`); Resume mode skips the file; Re-Run mode overwrites silently.
- **F8 Report creation from templates** (`core/excel_handler.py:65-273, 551-661`): copies the template to `Version folder/ITC_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx` (and `Sales_Report_...`) (`processing_handler.py:212-219, 276-277@edac6a3`) and writes folder paths and file names into the template's "Links" sheet (1.6). On Windows with pywin32 it uses Excel through COM; otherwise openpyxl (with a warning that Power Query may be lost).
- **F9 Summary workbook** (`excel_handler.py:319-549`): `Annual Statement-.../GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` with sheets Summary, Client Status, File Mapping, Errors, Variations (always openpyxl).
- **F10 Per-client text report** (`file_organizer.py:300-362`): `_Organization_Report_<client>_<state as written in file name>.txt` in the Version folder (raw client/state, not cleaned), listing folders, files, counts.
- **F11 Export client list** (`client_handler.py:227-259`): writes selected clients ("client - state") to a .txt chosen by the user, then opens it with `os.startfile` (Windows only; on other systems the error is caught and shown as "Failed to export" after the file was written).
- **F12 Dark mode** (`gui/utils/dark_mode_manager.py`, `main_window.py:150-180`), **F13 settings cache** (1.10), **F14 status bar** "v3.0 | Production Ready" (`gui/utils/status_bar.py:28-33`).

**Power Query Extractor (launch_extractor.py) — two tabs "Setup" and "Processing"** (`power_query_extractor/gui/extractor_window.py:80-87@edac6a3`).

- **F15 Find processed clients** (`extractor_window.py:353-406@edac6a3`): in the chosen target folder, the newest `Annual Statement-*` folder by modification time; inside it every sub-folder not starting with `_`; inside each, the `Version-*` folder with the **largest name** (`:385`, see 1.13 O6); flags whether `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` exist. All clients pre-ticked (`:427`).
- **F16 Refresh + extract** (`report_processor.py:37-141`): per client and per report type: take the first `<type>_Report_*.xlsx` in the version folder (`:54-67`), copy it to `<stem>_Refreshed_<DDMMYY_HHMM>.xlsx` (`:70-71, 97-98`), refresh it in a visible Excel by keystrokes (1.11), validate via `Info!BB2` (1.9), save, close, wait 3 s, then open the refreshed copy in a second hidden Excel and read the configured cells (`:462-571`). Option "Skip Refresh (Use existing refreshed files)" reuses the newest existing `<stem>_Refreshed_*.xlsx` (`:75-93`; the checkbox writes into the shared `EXTRACTION_OPTIONS` dict at run time, `extractor_window.py:234-243`).
- **F17 Consolidated workbook** (`data_consolidator.py:20-186`): `Annual Statement-.../PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` with a "Summary" sheet first and one sheet per `EXTRACTION_CONFIG` entry; columns Client Name, Successfully Run (✓/✗), then one column per mapping. When the refresh status is not successful, every value cell shows "Missing" (`:128-129`). Afterwards asks whether to open it (`os.startfile`, `extractor_window.py:571-573`).

### 1.4 File-name patterns the scanner recognises (code, not client data) — `utils/constants.py:13-62@edac6a3` (unchanged at the split point)

All are case-insensitive (`re.IGNORECASE`), must match the whole name, and accept `.xls` or `.xlsx` only (`\.xlsx?$`), so `.xlsm` files are found by the scan but never match. `[^-]+` means "one or more characters other than a hyphen", so **client and state names cannot contain a hyphen** in any pattern (RAN: confirmed).

| Key (match order) | Regex | Type → category → folder | Captured groups |
|---|---|---|---|
| `GSTR-2B-Reco` | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | "GSTR-2B Reco" → ITC → `itc` | client, state, period (may contain hyphens) |
| `ImsReco` | `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$` | "IMS Reco" → ITC → `itc` | client, state, date (exactly 8 digits, described as DDMMYYYY) |
| `GSTR3B` | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | "GSTR-3B Export" → GSTR3B → `gstr3b` | client, state, month; an optional prefix like `(2) ` is allowed — only this pattern has it |
| `Sales` | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | "Sales" → Sales → `sales` | client, state, start_month, end_month |
| `SalesReco` | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | "Sales Reco" → Sales → `sales` | client, state, period |
| `AnnualReport` | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$` | "Annual Report" → Annual → `root` | client, state, year |

- The period/date/month/year groups are stored in `result['metadata']` (`core/file_parser.py:98`) and **never used anywhere** (READ: grep at 47a1492 finds no reader). [GST RULE] note: the code does not check tax periods or financial years at all; the period parts are free text except the 8-digit IMS date.
- `EXPECTED_FILE_TYPES` (`constants.py:68-75`): the six types above; a client is "Complete" only when all six are present and no type other than GSTR-3B has more than one file (`file_parser.py:254-288`). [GST RULE] (decides which returns/files a client must have): GSTR-3B may have any number of files ("monthly filings - unlimited", `file_parser.py:268-270`); every other type with 2+ files is an "extra" and makes the status "Has duplicates" if nothing is missing. Status text: "Complete" ✅, "Missing N files" ⚠️, "Has duplicates" ⚠️. Completeness % = found types / 6 × 100, rounded to 1 decimal.
- Suggestion text for unmatched names (`file_parser.py:290-308`): keyword checks in this order — contains "gstr" and "2b" → GSTR-2B pattern; "ims" → IMS; "gstr" and "3b" → GSTR3B; "sales" and "reco" → SalesReco; "sales" → Sales; "annual" → AnnualReport; else "Check file naming convention".
- `validate_client_name()` (`helpers.py:321-342`) is applied to both the client and the state part: 2–100 characters, none of `< > : " / \ | ? *`. Problems are stored in `result['warnings']` (`file_parser.py:101-111`) and **never displayed** (READ).
- Client key = `<client as written>-<state code>` (`file_parser.py:228-252`), case-sensitive: "Alpha" and "alpha" are two clients; "X Private Limited" and "X Pvt Ltd" are two clients at scan time although `create_client_state_key()` later gives both the same output folder (MEDIUM, by reading; see 1.13 O15).

### 1.5 Constants (all of `utils/constants.py@edac6a3`, unchanged at the split point)

- `FOLDER_STRUCTURE` (`:81-90`): level1 `Annual Statement-{timestamp}`, level2 `{client}-{state}` (unused; the code uses `create_client_state_key`), version `Version-{timestamp}`, categories gstr3b `GSTR-3B Exports ({client})`, itc `Other ITC related files ({client})`, sales `Sales related files ({client})`.
- `EXCEL_TEMPLATE_MAPPING` (`:96-120`): see 1.6.
- `GUI_CONFIG` (`:126-151`): title "GST File Organizer v3.0", size 1200x800, min (1000, 700); fonts Segoe UI 10/16/12/9, Consolas 9; colours primary #0078D4, success #107C10, warning #FF8C00, danger #D83B01, info #0099BC, light #F8F9FA, dark #323130, white, complete #E8F5E8, incomplete #FFF4E5.
- `PROCESSING_MODES` (`:157-173`): fresh "🆕 Fresh Run" (new folder structure), rerun "🔄 Re-Run" (new version folders in existing structure), resume "▶️ Resume" (continue interrupted processing).
- `EXCEL_EXTENSIONS = {'.xlsx','.xls','.xlsm'}` (`:179`, never used — `find_excel_files` hard-codes its own list), `TEMPLATE_EXTENSIONS = {'.xlsx','.xltx','.xltm'}` (`:180`, unused), `MAX_FILE_SIZE_MB = 100` (`:181`, unused), `EXCEL_SIGNATURES` (`:182-185`, unused — `validate_excel_file` hard-codes the same bytes).
- `MESSAGES` (`:191-211`): only the `errors` entries no_source, no_templates, no_target, no_selection are used (`file_handler.py:123,135,139`, `processing_handler.py:30,61`); `invalid_excel`, `template_error`, `processing_error`, all `success` and all `info` entries are unused.
- `EXCEL_SAFETY` (`:217-228`): max_retries 3, retry_delay 1 s, temp_prefix `temp_gst_`, backup_prefix `backup_`, verification_size_min 10,240, save_options data_only/keep_vba/keep_links False. Stored as `ExcelHandler.safety_config` (`excel_handler.py:54`) and **never read** (READ, HIGH). The README's "temporary files / clean saves / no external links" claims rest on this unused dict.
- `STATE_CODE_MAPPING` (`:234-275`): see 1.9.

### 1.6 Template "Links" sheet mapping (what the organiser writes into each new report) — `utils/constants.py:96-120@edac6a3`

[GST RULE] (borderline — it decides which return/file feeds which input of the ITC and Sales reports; TIME-VARYING if the templates change):

| Report | Cell | Value written (`excel_handler.py:551-661`) |
|---|---|---|
| ITC | B2 | GSTR-3B folder path (the whole folder, so every monthly GSTR-3B file is available) |
| ITC | B4 / B5 | Version folder path / Annual Report file name without extension |
| ITC | B7 / B8 | ITC folder path / GSTR-2B Reco file name without extension |
| ITC | B10 / B11 | ITC folder path / IMS Reco file name without extension |
| Sales | B2 / B3 | Sales folder path / Sales file name without extension |
| Sales | B5 / B6 | Version folder path / Annual Report file name without extension |
| Sales | B8 / B9 | Sales folder path / Sales Reco file name without extension |

- Paths are written with backslashes (`clean_windows_path`, `helpers.py:311-315`); file names lose their extension (`helpers.py:446-448`), so the template must add the extension itself (MEDIUM). Only the **first** file of each type is used (`annual_files[0]`, `gstr2b_files[0]`, ... `excel_handler.py:577,588,599,615,631,647`), i.e. the alphabetically first name because `find_excel_files` sorts (`helpers.py:262`); duplicates are copied but not linked. Empty strings are not written (`if value:` `excel_handler.py:173,260`), so a missing file leaves the template's existing cell value in place (MEDIUM).
- Sheet lookup: exact name "Links" (case-insensitive), else any sheet with "link" in its name, else (COM path only) a new sheet "Links" is added (`excel_handler.py:146-164`); the openpyxl path writes nothing if no Links sheet (`:255`).

### 1.7 Extraction cell mapping (what the extractor reads) — `power_query_extractor/config/cell_mappings.py@edac6a3`

[GST RULE] TIME-VARYING (template layout and GSTR table references can change):

| Output sheet | Report | Input sheet!cell → output column |
|---|---|---|
| `GSTR-3B Difference` (`:35-58`) | ITC | `Diff GSTR-3B`!J15 → "Diff IGST"; !K15 → "Diff CGST"; !L15 → "Diff SGST" |
| `GSTR-1 Difference` (`:61-89`) | Sales | `Diff - GSTR-1`!N1 → "Diff Taxable Value"; !O1 → "Diff IGST"; !P1 → "Diff CGST"; !Q1 → "Diff SGST" |
| `Turnover GSTR-1` (`:92-102`) | Sales | `Diff - GSTR-1`!I1 → "Taxable Value" |

- `EXTRACTION_OPTIONS` (`:122-127`): skip_refresh False, auto_refresh_if_missing True (never read), show_timestamp False (never read), missing_data_text "Missing".
- A "backward compatibility" block (`:133-152`) builds `CELL_MAPPINGS[report][input sheet][cell] = "<output sheet>_<output column>"` with spaces turned into underscores, plus `FIELD_DISPLAY_NAMES`. If two output sheets ever map the same input sheet and cell, the later one overwrites the earlier one and the earlier column would always show "Missing" (READ, HIGH; no such duplicate exists at edac6a3 or 47a1492).
- `calculation` types documented (`:21-26`): subtract, add, multiply, divide (divide by zero → text "Div/0"; non-numbers → "Error"; `data_consolidator.py:188-218`). No mapping uses a calculation in this range. If used, `cell1`/`cell2` are looked up by bare cell reference in the same report's extracted values, where cell references from different input sheets would collide (MEDIUM).
- `cell_mappings copy.py@edac6a3` (63 lines): the older format — `CELL_MAPPINGS = {'ITC': {'Diff GSTR-3B': {'J15': 'itc_diff_igst', 'K15': 'itc_diff_cgst', 'L15': 'itc_diff_sgst'}}, 'Sales': {'Diff - GSTR-1': {'N1': 'sales_diff_tv', 'O1': 'sales_diff_igst', 'P1': 'sales_diff_cgst', 'Q1': 'sales_diff_sgst'}}}` with display names like "Diff IGST (ITC)"; commented ideas `'E5': 'reversed_itc'`, `'F5': 'net_itc'`, `'D10': 'export_sales'`, `'E10': 'exempt_sales'`. The same cells as the new file except that `Diff - GSTR-1!I1` ("Taxable Value", Turnover) is new in the master version. The compiled-file evidence (1.16) dates this older format to 21-06-2025.

### 1.8 Output folders and file names (READ; RAN for the folder names)

```
<Target>/Annual Statement-<DDMMYY HHMM>/            (Fresh: new each run; Re-Run/Resume: newest existing by mtime)
    GST_Processing_Summary_<DDMMYY_HHMM>.xlsx
    PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx       (written later by the extractor)
    <client-state key>/                              (create_client_state_key, max 35 chars by default)
        Version-<DDMMYY HHMM>/                       (always a new timestamp per run)
            GSTR-3B Exports[ (<cleaned client>)]/
            Other ITC related files[ (<cleaned client>)]/
            Sales related files[ (<cleaned client>)]/
            <Annual Report file, cleaned name>
            ITC_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx
            Sales_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx
            ITC_Report_..._Refreshed_<DDMMYY_HHMM>.xlsx (written later by the extractor)
            _Organization_Report_<client>_<state>.txt
```

- Timestamps: `get_timestamp()` = `%d%m%y %H%M` (`helpers.py:23-25`); report and summary names replace the space with `_`. Minute precision: two runs in the same minute share the same Annual Statement and Version folders (RAN: a Re-Run in the same minute reused the Version folder). The README claims `DDMMYYYY HHMMSS` (`README.md:113-121@edac6a3`), which does not match.
- Brackets with the client name are added when the global checkbox is on, or the per-client setting for key `<client>-<state code>` is on (`file_organizer.py:117-130`). The bracket uses the *cleaned* client (spaces → `_`), the level-2 folder uses the *raw-abbreviated* client (spaces kept) (RAN: `Alpha Traders-MH` vs `(Alpha_Traders)`).
- `_get_destination_folder` (`file_organizer.py:285-298`): GSTR-3B Export → gstr3b; GSTR-2B Reco and IMS Reco → itc; Sales and Sales Reco → sales; Annual Report → the Version folder itself.

### 1.9 Business rules and [GST RULE]s in the original code (READ; RAN where noted)

- **[GST RULE] State name → state code** (`utils/helpers.py:485-511`, map `constants.py:234-275`). TIME-VARYING (states and union territories change). Lower-cased, trimmed name is looked up in 39 entries mapping full names to two-letter **alphabetic** abbreviations (for example 'maharashtra' → MH, 'telangana' → TS, 'odisha' and 'orissa' → OD, 'uttarakhand' → UK, 'chhattisgarh' → CG, 'dadra and nagar haveli and daman and diu' → DD, 'delhi' and 'national capital territory of delhi' → DL, 'jammu and kashmir' → JK, 'ladakh' → LA, 'puducherry'/'pondicherry' → PY, 'andaman and nicobar islands' → AN). These are not the numeric GST state codes; the code never uses numeric codes. Not in the map → fallback with a WARNING log: two or more words → first letters of the first two words; one word → first three characters upper-cased. RAN examples: 'Maharashtra' → MH; 'MH' → MH (via fallback, with warning); 'NCT of Delhi' → NO; 'Dadra and Nagar Haveli' → DA; 'Daman and Diu' → DA; 'Gujrat' → GUJ; 'Jammu & Kashmir' → J&; 'Andaman & Nicobar' → A&; '' → ''. A number (int) crashes it (RAN: `AttributeError: 'int' object has no attribute 'strip'`; see 1.13 O8). The code does not derive or check GSTINs or PANs anywhere (READ, HIGH).
- **Business abbreviations**: `create_client_state_key()` (`helpers.py:513-537`) replaces whole words "Private" → "Pvt" and "Limited" → "Ltd" case-insensitively; `sanitize_filename()` (`helpers.py:66-112`) replaces the substrings "Private" → "Pvt" and "Limited" → "Ltd" case-sensitively anywhere (RAN: "Privately Limitedness" → "Pvtly_Ltdness"), then replaces control characters and `<>:"/\|?*[]{}+=!@#$%^,;'"`~` with `_`, collapses spaces/underscores to one `_`, strips `_ . space` at the ends, uses "unnamed" if empty, and cuts the stem to `200 - len(extension) - 10` characters (RAN: a 300-character name → 190 characters).
- **Client-state key length**: default 35 (`file_organizer.py:42`, GUI default `main_window.py:89`); if longer, the client part is cut to `35 - len(state code) - 1`; if that would leave 5 or fewer characters the whole key is cut (and the state code can be lost) (RAN: max 8 → "Alpha Be").
- **Long-name warning**: >10 characters when "include client name" is on (`processing_handler.py:64-75@edac6a3`, `client_handler.py:194@edac6a3`).
- **Refresh validation rule (extractor)** (`report_processor.py:573-620@edac6a3`): a refresh counts as valid only if a sheet named "Info" (any case) exists and its cell **BB2** is not empty, not blank, not "0"/"0.0", and not one of `#REF! #VALUE! #NAME? #NULL! #DIV/0! #N/A #NUM! #ERROR!`; the success message calls BB2 "Client Name" (`:616`). Business rule about the template (TIME-VARYING with the template).
- **Display of figures** [GST RULE - display only]: consolidated values that are numbers other than 0 get the Excel number format `#,##0.00` (`data_consolidator.py:161-162`); zeros keep the default format; stored values are not rounded (READ, HIGH). Not-successful reports show "Missing" for every value (`:128-129`), even if values were read.
- **Completeness / duplicates**: see 1.4.

### 1.10 Shared state — who writes and who reads (READ, HIGH unless stated)

| State | Writer | Reader | Consistent? |
|---|---|---|---|
| Main settings cache | `gui/utils/cache_manager.py:15,29-48` writes `~/.gst_organizer_cache.json` (home folder, dot name); keys source_folder, itc_template, sales_template, target_folder, processing_mode, include_client_name, dark_mode, client_name_max_length (`main_window.py:328-340`) + recent_folders (last 5 source folders) | `main_window.py:342-356` | Yes within the organiser |
| Extractor "auto-load from cache" | — | `extractor_main.py:36-60` tries `./gst_organizer_cache.json`, `./.cache/app_state.json`, `./config/last_session.json` (relative to the start folder) with keys target_folder/target_path/output_folder/last_target; `extractor_window.py:28,329-341` tries `./gst_organizer_cache.json` | **No**: different folder (start folder vs home) and different name (no leading dot). The extractor never finds the organiser's cache, so the README promise "auto-loads target folder from cache" (`README.md:171@d674386`) does not happen (HIGH). Fixed at A HEAD (`power_query_extractor/extractor_main.py:41`, `gui/extractor_window.py:118` use `Path.home() / '.gst_organizer_cache.json'`); B HEAD's `power_query_extractor/extractor_main.py:39` still has the relative name but B's launcher no longer uses that module (see 10.6). |
| Logs | `main.py:25` → `./gst_organizer.log`; `extractor_main.py:16` → `./pq_extractor.log` | People | Both relative to the start folder, appended forever |
| `temp` folder | `excel_handler.py:52-53` creates `<system temp>/gst_excel_temp` on every start | nobody | Unused |
| Reports produced by organiser | `processing_handler.py:218-219,276-277@edac6a3` (`ITC_Report_*`, `Sales_Report_*` in Version folder) | extractor glob `ITC_Report_*.xlsx` (`report_processor.py:54-55`, `extractor_window.py:388-389`) | Yes, but the glob also matches the extractor's own `..._Refreshed_...` copies; on NTFS the original usually sorts first because `.` sorts before `_` (LOW-MEDIUM; order is not guaranteed by Python) |
| Refreshed copies | `report_processor.py:70-71,97-98` | skip-refresh lookup `:80` | Yes at edac6a3 (both `_Refreshed_`); broken by bd176bf (see section 2) |
| `EXTRACTION_OPTIONS['skip_refresh']` | `extractor_window.py:237-238` mutates the module dict | `report_processor.py:75-76` | Yes (same module object) |

### 1.11 Windows-only parts and Excel through COM (compare Standard 12, lines 844-878) (READ; cannot run here)

Main organiser, `core/excel_handler.py@edac6a3`:
- Import guarded: pywin32 used only if `platform.system() == 'Windows'` and the import works (`:20-28`); otherwise openpyxl (`:95-101`). `helpers.get_short_path` imports win32api lazily with fallback (`helpers.py:473-483`).
- `_create_report_win32com` (`:103-228`): `pythoncom.CoInitialize()`; **`GetActiveObject("Excel.Application")` first — attaches to the user's own running Excel if there is one** — else `Dispatch(...)` (not `DispatchEx`) (`:121-126`); sets `Visible=False`, `DisplayAlerts=False`, `ScreenUpdating=False` on that instance (`:129-131`) — so a user's own Excel would be hidden; copies the template to the output name and sleeps 0.5 s (`:135-136`); opens it (`_safe_open_workbook`, `:275-317`: reuses an already-open workbook with the same path; retries with the 8.3 short path on error -2147352567, "path" in the message, or path > 218 characters); writes Links cells; `wb.Save()` then `wb.Close(SaveChanges=True)`; checks the file exists. `finally`: closes the workbook if still open, sets `ScreenUpdating=True` and **`DisplayAlerts=True`**, and calls `excel.Quit()` if `self.excel_app is None` — `self.excel_app` is set to None in `__init__` (`:55`) and never assigned, so **Quit is always called, including on a user's Excel attached through GetActiveObject** (comment says "Don't quit if Excel was already open", `:217`) (HIGH by reading that the condition is always true; MEDIUM on what Excel then does with the user's unsaved workbooks, since alerts were just switched back on and the window is hidden); then `CoUninitialize()` while the local `excel`/`wb` objects are still referenced. No `IgnoreReadOnlyRecommended`, no `workbook.ReadOnly` check, no refresh (the organiser only writes paths). The log (1.15) shows 6 × "Created new Excel instance" and 0 × "Using existing Excel instance", so in the logged runs no user Excel was attached.
- `_create_report_openpyxl` (`:230-273`): copies the template and saves through openpyxl with `keep_links=True` — this is "openpyxl saving over a workbook the app did not create", without a Power Query check (Standard 12 first bullet). It only runs off Windows or without pywin32.
- `__del__` (`:57-63`) would Quit `self.excel_app` — never set, so it does nothing.

Extractor, `power_query_extractor/core/report_processor.py@edac6a3`:
- Imports pythoncom, win32com.client, win32api, win32con, win32gui at module level with no fallback (`:11-15`) — inconsistent with the organiser's guarded import.
- `_refresh_power_query_simple` (`:143-350`): `print("PROCESSING: ...")` placed before the docstring (`:145`, so the docstring is not a docstring and the line writes to the console); `CoInitialize`; `Dispatch("Excel.Application")` (not `DispatchEx`; may return a running instance — MEDIUM); sets **`Visible=True`**, `DisplayAlerts=False`, `WindowState=-4137` (maximised) with up to 5 tries 2 s apart ("Keep it visible to avoid issues", `:161`); opens the workbook (3 tries on "Call was rejected by callee"); sleeps 5 × 1 s; activates the window with `WScript.Shell.AppActivate(excel.Caption)` (3 tries); then **sends keystrokes `%a` (Alt+A), `r`, `a`** one second apart (`:216-229`) — Excel's ribbon shortcut for Data → Refresh All; waits 10 × 1 s (comment says "WAIT 7 SECONDS AS REQUESTED", `:235-238` — comment and code disagree); checks 5 times for an error dialog (`_find_error_window`, `:352-388`: any visible enabled top-level window whose title contains one of 'microsoft excel', 'error', 'problem', 'cannot', 'failed', 'unable', 'could not', 'exception', 'invalid', 'missing' and whose class is `#32770`, `NUIDialog` or `bosa_sdm_msword`); reads its text (`:390-425`) and presses a button whose label contains ok/close/yes/continue, else posts WM_CLOSE (`:427-460`); waits 10 more seconds; a final error check; `_validate_refresh` (BB2); `wb.Save()`; `wb.Close(SaveChanges=True)`. `finally`: close without saving, `Quit()`, `CoUninitialize()`. A trailing `return {'success': True, 'error': None}` (`:318-321`) can never run (both branches above it return). No `BackgroundQuery=False`, no `RefreshAll()`, no `CalculateUntilAsyncQueriesDone()`, no check of connections' `Refreshing` — the refresh is assumed finished after the fixed waits.
  - Consequence (READ, HIGH): keystrokes go to whichever window is in front; if the user clicks elsewhere during the ~25 s, the keys go to that window and the refresh does not happen. The program needs an interactive logged-in desktop.
- `_process_report` (`:50-141`): the `except` around the refresh that looks for "Open method of Workbooks class failed" (`:101-115`) can never trigger, because `_refresh_power_query_simple` catches every exception itself and returns a dict (HIGH). Extraction runs even when the refresh failed (`:121-122`); the consolidator then shows "Missing".
- `_extract_data` (`:462-571`): second `Dispatch` with `Visible=False`; logs "Opening file in read-only mode" and "✓ File opened in read-only mode" (`:476-478`) but opens read-write (`_safe_open_workbook` → `Workbooks.Open(file_str)` with no ReadOnly argument) — log text and code disagree (HIGH). Sheet names must match exactly (case-sensitive, `:500`); a missing sheet gives None for its fields. Each value is logged (`:525, 541-546`), so extracted figures go into `pq_extractor.log` and the GUI log.
- `_safe_open_workbook` (`:622-731`): for paths > 218 characters converts to `\\?\` or `\\?\UNC\` extended form; tries Open; then closes same-named workbooks (without saving) and retries; then 8.3 short path; then without the prefix; then `Workbooks.Open(Filename=..., UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, Notify=False)`; then prints guidance (file open elsewhere, corrupted, path too long, OneDrive/SharePoint sync).
- Other Windows-only calls without guard: `os.startfile` in `client_handler.py:257` (inside try), `extractor_window.py:573` (inside try of the worker), `processing_handler.py:429` (guarded by platform check).

### 1.12 Start-up, console and file locations (compare Standard 13) (READ; RAN where noted)

- Needs a console: `print()` banners and `input("Press Enter...")` on errors (`main.py:54-56,63-79`); the stdout wrapper crashes without a console (RAN simulation, 1.2).
- Writes into the app folder at start: `logs/`, `temp/`, missing `__init__.py` files (`main.py:36-44`; RAN), and `gst_organizer.log` in the start folder (RAN).
- Settings in the home folder (`~/.gst_organizer_cache.json`), not `%LOCALAPPDATA%`.
- No `sys.excepthook` or `threading.excepthook`; worker-thread errors are partly caught and shown with `messagebox` called from the worker thread (`processing_handler.py:439-442@edac6a3`, `extractor_window.py:575-578`).
- No pinned versions: `requirements.txt` uses `>=`.

### 1.13 Defects and edge cases in the original code, with status at the split point and at both HEADs

(Technical points only. "A HEAD"/"B HEAD" checks are spot checks by grep/read, marked with confidence; deeper HEAD analysis belongs to other readers.)

| # | What (evidence) | Split point 47a1492 | A HEAD | B HEAD |
|---|---|---|---|---|
| O1 | Extractor cannot find the organiser's cache (1.10) — READ, HIGH | still | fixed (`power_query_extractor/extractor_main.py:41`, `gui/extractor_window.py:118`) | relative path remains in unused `power_query_extractor/extractor_main.py:39` (MEDIUM) |
| O2 | Consolidator substring match (`data_consolidator.py:142-145`) can place another sheet's value under a column whose name is contained in an earlier field name — RAN | **triggered by 47a1492's new columns** (section 0 item 1) | still (`power_query_extractor/core/data_consolidator.py:143`; RAN wrong result) | still (`power_query_extractor/core/data_consolidator.py:131`; RAN wrong result) |
| O3 | Per-client "include name" toggle stored under `<client>-<state as shown>` (`client_handler.py:184,220@edac6a3`) but looked up as `<client>-<state code>` (`file_organizer.py:118,126`) — RAN: setting ignored when the file uses a full state name | still (`client_handler.py:178,214@47a1492`) | fixed (`gui/handlers/client_handler.py:181,218` use `get_state_code`) | Tk code replaced by PyQt; not checked (verify) |
| O4 | Client tree colours wrong (all "complete" in light mode, all "incomplete" in dark mode) — RAN | fixed by 1c48722 | n/a (fixed) | n/a |
| O5 | START/STOP use `bind('<Button-1>')` so a "disabled" button still fires (RAN), and `start_processing` has no "already running" guard — a second processing thread can start | still (`processing_tab.py:74`, `validation_tab.py:142`) | binding still (`gui/tabs/processing_tab.py:74`) but guard added (`gui/handlers/processing_handler.py:75`) | PyQt; not checked |
| O6 | Extractor chooses the "latest" Version folder with `max(..., key=name)`; names are `Version-DDMMYY HHMM`, so day-first text order is not date order (RAN: `Version-310825 ...` beats `Version-151025 ...`) | still (`extractor_window.py:529@47a1492`) | fixed (`power_query_extractor/gui/extractor_window.py:648` uses mtime) | replaced by explicit version selection in the PyQt tree (`gui/pyqt_main_window.py:517-528`; MEDIUM) |
| O7 | Files < 1 KB or with a wrong signature vanish silently from the scan (never reach "variations") — RAN | still | not checked | not checked |
| O8 | Numeric-looking values in the client tree come back as numbers (RAN: "0123" → 123, "07" → 7), so the key no longer matches `client_data` → KeyError; and a numeric state (for example a numeric GST state code in the file name) makes `get_state_code()` crash with AttributeError (RAN) in Dry Run / Start / Export | still (`client_handler.py:161-173@47a1492`) | `get_state_code` still crashes on int (RAN on a copy of A HEAD utils); A HEAD `gui/handlers/client_handler.py:169` still passes the tree value (MEDIUM) | PyQt; not checked |
| O9 | Resume mode always creates a new `Version-<now>` folder, so "skip if already exists" only works within the same minute: Resume behaves like Re-Run (READ, HIGH) | still | not checked | not checked |
| O10 | `main.py` stdout wrapper crashes without console (RAN simulation) | still (`main.py:17`) | still (`main.py:17`) | guarded (`main.py:12-13`) |
| O11 | Worker thread touches Tk widgets directly (processing_handler, extractor_window) | partly wrapped by 6237597 | `_log_async`/`_progress_async` (`gui/handlers/processing_handler.py:23-36`) | PyQt |
| O12 | Main app may attach to and Quit the user's own Excel (1.11) — READ | still | not checked | not checked |
| O13 | SendKeys refresh with fixed waits (1.11) — READ | still (`report_processor.py:241-251,256-289@47a1492`) | not checked | not checked |
| O14 | "Read-only" log message while opening read-write (1.11) — READ | still (`report_processor.py:500-502@47a1492`) | not checked | not checked |
| O15 | Name variants of one client ("Private Limited" vs "Pvt Ltd", different capitals) are separate clients at scan but get the same level-2 folder and the same report file names (for capital-letter variants this holds on Windows, whose file names ignore case); the second client's report copy overwrites the first's without backup (report creation copies the template over the output, `excel_handler.py:135`) — READ, MEDIUM | still | not checked | not checked |
| O16 | `toggle_folder_name_setting` length warning uses >10 while start-up warning (after 6237597) uses the configurable length; label still says ">10" | introduced by 6237597 | still (`gui/handlers/client_handler.py:191`; `gui/tabs/setup_tab.py:447` comment) | label `gui/views/setup_view.py:385` still ">10"; no check code found (MEDIUM) |
| O17 | Inconsistent max-length limits: setup spinbox 15–100 vs validator 10–100 vs label "range: 15-100" (`setup_tab.py:252-277`) | still | not checked | not checked |
| O18 | Closing the window during processing destroys it after 1 s while the daemon thread may be in the middle of a copy or an Excel call (`main_window.py:314-325`) | still | not checked | not checked |
| O19 | `export_client_list` opens with `os.startfile` (Windows only) and writes with the default encoding (`client_handler.py:243,257`) | still | not checked | not checked |
| O20 | `dark_mode_manager._restore_widget_colors` sets every Frame/Canvas to 'white' (`:171-172`), not the original light grey #F8F9FA — after dark→light the colours differ from start-up (MEDIUM) | still | not checked | not checked |

### 1.14 Dead code at edac6a3 (READ; grep for call sites at 47a1492 found none outside the definition and the `__init__` re-export)

- `utils/helpers.py`: `calculate_file_hash` (:422), `get_relative_excel_path` (:450), `get_unique_path` (:344) — defined and re-exported in `utils/__init__.py`, never called. (`format_duration` is used, by `ProgressTracker.get_eta`.)
- `utils/constants.py`: `EXCEL_EXTENSIONS`, `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB`, `EXCEL_SIGNATURES`, `EXCEL_SAFETY` (stored, never read), `FOLDER_STRUCTURE['level2']`, unused `MESSAGES` entries (1.5).
- `core/file_parser.py`: `validate_results` (:340-373, defines "critical" types Annual Report and GSTR-3B Export — never called), `self.errors` (collected, never shown), `result['warnings']`, `result['metadata']`.
- `core/file_organizer.py`: `cleanup_on_error` (:364-379), `_is_folder_empty` (only used by it), `get_summary` (:392-409), `self.errors` (only counted in get_summary).
- `core/excel_handler.py`: `self.temp_dir` (created, unused), `self.safety_config`, `self.excel_app` (never assigned), `__del__` effectively empty.
- `gui/utils/ui_helpers.py`: `update_status` (:11-14). `gui/utils/cache_manager.py`: `get_cached_value`, `set_cached_value`.
- `power_query_extractor/core/data_consolidator.py`: `_write_section` (:250-352) and `_create_details_sheet` (:354-414) — remnants of an older single-sheet "ITC section / Sales section / details" report; lines 73-83 write the header row twice (a pasted "# ... rest of the code" fragment followed by the real code at 85-99).
- `power_query_extractor/core/report_processor.py`: `self.excel_instance`, the unreachable return (:318-321), the dead `except` branch (:103-115), `cleanup()` (does nothing).
- `power_query_extractor/config/cell_mappings.py`: `auto_refresh_if_missing`, `show_timestamp`. `cell_mappings copy.py`: whole file (not importable).

### 1.15 What `gst_organizer.log` holds (special handling: shapes and counts only; nothing quoted)

How I looked (RAN): line count, the anonymised line-shape command given in the task, counts per logger name and level, date counts, and a script (`logshape.py`) that counts messages starting with each fixed text template taken from the code — it prints only the template text from the code and counts.

- **Size**: 5,689 lines, 803,193 bytes. Every line has the standard format `<date time,ms> - <logger name> - <level> - <message>`; 0 lines in other formats.
- **Levels**: all 5,689 lines are INFO (0 WARNING, 0 ERROR). The "[DEBUG]" messages are written with `logger.info(...)`, which is why they appear at INFO level.
- **Logger names**: core.file_parser 5,334; core.excel_handler 222; core.file_organizer 59; utils.helpers 52; root 22.
- **Dates**: 23-07-2025 (4,785 lines), 28-08-2025 (2), 30-08-2025 (7), 31-08-2025 (895). First entry 23-07-2025 11:47, last 31-08-2025 21:57 (local clock of the computer).
- **Every line matches a logging call in the edac6a3 code** (0 unmatched):
  - root 22 = "Using win32com for Excel operations (Power Query support)" from `core/excel_handler.py:26` — logged once per start of the organiser, and only on Windows with pywin32 → 22 start-ups (7 on 23-07, 2 on 28-08, 7 on 30-08, 6 on 31-08).
  - core.file_parser: "Found N Excel files to process" ×6 and "Scan complete" ×6 (6 scans: 5 on 23-07, 1 on 31-08); "[DEBUG] Parsing potential Sales file" ×1,398 (`file_parser.py:71`), "[DEBUG] Sales pattern matched" ×1,308 (`:117`), "[DEBUG] Extracted data: client=..., state=..., type=..." ×1,308 (`:118`), "[DEBUG] Sales file detected" ×1,308 (`:252`).
  - core.file_organizer: "FileOrganizer initialized" ×1 (one processing run, 23-07-2025), "Created folder structure for" ×3 and "Created organization report" ×3 (3 clients processed), "Copied ... to ..." ×52.
  - utils.helpers: "Successfully copied ... to ..." ×52.
  - core.excel_handler: "Created new Excel instance" ×6, "Successfully Created ... report" ×6 (3 ITC + 3 Sales), "Opening workbook" ×12, "Path length" ×6, "Updating cells" ×6, cell-write lines "  B<n> = ..." ×39 (13 per client = all 7 ITC and 6 Sales Links cells filled), "Updated N cells" ×6, "Saving workbook" ×6, "Copying template" ×6, and 129 "[DEBUG]" lines from `create_report_from_template` and `prepare_template_data` (template path, output path, data-mapping dictionary, client, state, state code, files available, folders and file names for Sales/Annual/Sales Reco).
- **Kinds of information inside** (from the templates above and the anonymised shapes): client names and state names; source file names; full Windows folder paths (drive-letter paths) of the source folder, target folder, templates and output folders; output report names; file sizes in bytes; counts. **No tax figures**: the organiser never reads cell values (READ, HIGH).
- **How it was produced**: `main.py:21-28@edac6a3` `logging.basicConfig(level=INFO, FileHandler('gst_organizer.log'), StreamHandler(stdout))`. The file is opened relative to the folder the program was started from (here evidently the project folder, since it was committed from there) and is appended to on every start, so it accumulated from 23-07-2025 to 31-08-2025 and was then committed by the first commit (there was no .gitignore yet).
- **Where it is now**: in history only. Added in edac6a3 (31-08-2025 22:11), deleted 3 minutes later in cbfa4a5 (31-08-2025 22:14, made on GitHub's web page); the local line of work kept it until the merge 9c56a1c (29-10-2025) brought that deletion in; `.gitignore` (1cb878b) then ignores `*.log` and `gst_organizer.log`. Because edac6a3 is an ancestor of both HEADs, anyone with either repository can still read it with `git show edac6a3:gst_organizer.log` (HIGH).
- Inference (MEDIUM): with zero WARNING lines, no "Could not parse filename" and no "Unknown state ... using fallback" happened in these runs, so every state part in the scanned file names was found in `STATE_CODE_MAPPING` (full state names rather than codes), and every Excel file found matched a pattern.
- Inference (MEDIUM): the log starts on 23-07-2025 at 11:47, the same day and hour in which the modular GUI files were being saved (11:38–13:45 IST, see 1.16); it was probably a fresh log started in a new project folder when the modular version was first run.

### 1.16 The 70 compiled files (`__pycache__/*.pyc`) and the never-committed `gui/app.py` (special handling: no embedded path reproduced)

- Added in edac6a3, removed in 01862a0 (29-10-2025 21:46). 70 files = 35 modules × 2 Python versions: `cpython-312` (Python 3.12, magic bytes `cb0d`) and `cpython-313` (Python 3.13, `f30d`). Modules: `core/{__init__,excel_handler,file_organizer,file_parser}`, `gui/{__init__,app,main_window}`, `gui/handlers/{__init__,client_handler,file_handler,processing_handler}`, `gui/tabs/{__init__,processing_tab,setup_tab,validation_tab}`, `gui/utils/{__init__,cache_manager,dark_mode_manager,status_bar,ui_helpers}`, `gui/widgets/{__init__,collapsible_frame,title_bar}`, `power_query_extractor/{__init__,extractor_main}`, `power_query_extractor/config/{__init__,cell_mappings}`, `power_query_extractor/core/{__init__,data_consolidator,report_processor}`, `power_query_extractor/gui/{__init__,extractor_window}`, `utils/{__init__,constants,helpers}`. No pyc for `main.py`, `launch_extractor.py` (run as scripts) or `cell_mappings copy.py` (not importable).
- Compiled files embed the full folder path of the computer they were made on (Python stores the source path in each file). RAN (`pyc_paths2.py`, which prints only hashes and lengths): all embedded paths are Windows-style; there are at least four different project-folder locations per Python version, i.e. the project folder was copied or moved several times, with its `__pycache__` carried along (MEDIUM). I did not reproduce any path. A company-suffix word scan of all pyc strings (counts only) found matches only in `utils/helpers` (the code's own "Private"/"Limited" abbreviation strings).
- RAN (`pyc_hdr.py`, header fields only): each pyc records its source file's size and modification time. 34 of the 35 modules have a committed source (all but `gui/app`). For all 34, the Python 3.13 pyc records exactly the committed file's size converted to Windows line endings (CRLF); for Python 3.12, 30 of 34 match and the other 4 are older versions of extractor files (next bullet). So the pyc files came from exactly the committed text and the code was edited on Windows (HIGH). Source last-modified dates (UTC): `core/__init__`, `utils/__init__` 19-06-2025; `power_query_extractor/__init__`, its three sub-package `__init__`s and `extractor_main` 21-06-2025; `excel_handler`, `file_parser`, `constants`, `gui/app` 25-06-2025; `cell_mappings`, `report_processor`, `extractor_window` 28-06-2025 (3.13); `data_consolidator` 14-07-2025 (3.13); modular GUI files, `file_organizer`, `helpers` 23-07-2025 06:08–08:15 UTC (11:38–13:45 IST).
- Older 3.12 copies of four extractor modules record older sources: `cell_mappings` 1,938 bytes (21-06-2025) — exactly the CRLF size of `cell_mappings copy.py` (1,876 bytes + 62 line ends), so the "copy" file is the 21-06-2025 version of the mapping (HIGH); `data_consolidator` 10,208 bytes (25-06-2025) vs 17,711 committed (14-07-2025) — the consolidator nearly doubled in July (MEDIUM: probably the per-sheet `EXTRACTION_CONFIG` layout replacing the dead `_write_section` layout); `report_processor` 29,229 (25-06) vs 30,861 (28-06); `extractor_window` 20,875 (21-06) vs 21,963 (28-06).
- **`gui/__pycache__/app.cpython-312.pyc` and `app.cpython-313.pyc`** (98,247 and 97,844 bytes) belong to a `gui/app.py` of **89,275 bytes** (Windows line endings), last modified 25-06-2025 16:14 UTC (21:44 IST), that is not in any commit. What it suggests (HIGH unless stated):
  1. The first GUI was one large file `gui/app.py` (roughly the size of the whole modular GUI today). On 23-07-2025 it was split into `gui/main_window.py` ("Modularized version - ties all components together", `gui/main_window.py:4`) plus `tabs/`, `handlers/`, `utils/`, `widgets/`.
  2. `app.py` was deleted (or left behind in another folder) before the first commit, but its compiled leftovers were committed because there was no `.gitignore`.
  3. The README of edac6a3 still shows `gui/app.py` in the folder tree (`README.md:51-53@edac6a3`) and `main.py:72@edac6a3` still lists it; both were written for the app.py era and not updated.
  4. The project had about 2.5 months of history (19-06-2025 to 31-08-2025) before git; that history (and the reasons for many design choices) is lost. The app.py pyc files could, if ever needed, be decompiled to recover the pre-split GUI's structure (they also contain embedded paths, so handle as sensitive) (LOW need).
  5. Both Python 3.12 and 3.13 ran the modular GUI after 23-07-2025 (both pyc sets have the 23-07 timestamps), so two interpreters (perhaps two computers) were in use (MEDIUM).
  6. Main-app source files did not change between 23-07-2025 and the first commit on 31-08-2025: the log shows start-ups up to 31-08 and Python rewrites a module's pyc whenever the source's time or size changes, yet the pyc timestamps stop at 23-07 and the sizes match the committed files (MEDIUM: assumes those runs used this same folder, which the relative log path suggests).

### 1.17 README.md at edac6a3 versus the code (READ)

- Claims: "Python 3.7 or higher"; folder tree with `gui/app.py` only (no tabs/handlers, no extractor) (`:36-54`); naming patterns (`:80-85`) match the code; output names `Annual Statement-DDMMYYYY HHMMSS`, `Version-DDMMYYYY HHMMSS`, `ITC_Report_ClientName_State_DDMMYYYY_HHMMSS` (`:111-123`) — the code uses DDMMYY HHMM and the state code; "All Excel operations use temporary files" / "Files are verified before finalization" / "No VBA macros or external links preserved" (`:129-133`) — the code copies templates straight to the output, `EXCEL_SAFETY` is unused, the COM path keeps whatever the template contains and the openpyxl path uses `keep_links=True`; "Backups created for existing files" — only in Fresh mode for copied files; Links cell list (`:167-184`) matches the code; summary sheets (`:196-202`) match; "Check the log file: gst_organizer.log" (`:216`); "Excel corruption issues from previous versions have been specifically addressed" (`:254`) — evidence of an earlier, corrupting version (pre-git).

---

## 2. Commit-by-commit (oldest first)

#### edac6a3 - 31-08-2025 22:11 - goyal-sidhant - shared - before the split on 30-10-2025
- **Message**: "Initial commit" (no body).
- **Files touched**: 110 files, +12,916 lines (READ: `git show --stat`). 37 Python files (6,954 lines), README.md (+260), requirements.txt (+13), `gst_organizer.log` (+5,689), 70 `__pycache__/*.pyc` (binary). Full description of every code file in section 1 (1.1–1.17).
- **Why**: first upload of a program that had already been developed for about 2.5 months outside git (HIGH: compiled-file timestamps 19-06-2025 to 23-07-2025, log entries from 23-07-2025, see 1.15–1.16). It was committed together with build leftovers and the run log because there was no `.gitignore` (HIGH).
- **What it reveals**: the v3.0 design (section 1): a Tk organiser (scan → validate → organise → template reports → summary) plus a separate Windows-only Power Query Extractor (refresh by keystrokes → validate Info!BB2 → read configured cells → consolidated workbook). Strongly AI-assisted authorship: "Author: Advanced AI Assistant" (`main.py:6`), comments such as "ADD THIS LINE" (`utils/helpers.py:15`), "ADD THIS" / "ADD THESE" (`core/file_parser.py:15`, `core/file_organizer.py:15`, `core/excel_handler.py:39`), "ADD THE NEW CODE HERE (notice the indentation - 3 levels)" (`report_processor.py:73`), "Rest of your code remains same..." (`report_processor.py:231`), "# ... rest of the code" (`data_consolidator.py:75,80`) — fragments pasted from chat answers (HIGH). Earlier versions corrupted Excel files and lost Power Query, which is why the organiser switched to Excel COM (header `core/excel_handler.py:2-3` "WINDOWS VERSION - Uses Excel COM to preserve Power Query", README `:3,129-133,254`) (MEDIUM-HIGH). Many tuning comments ("Increased from 0.5", "Increased from 3 to 5", "WAIT 7 SECONDS AS REQUESTED") show trial-and-error timing of the keystroke refresh before git (HIGH).
- **Type**: feature (initial import) + accidental data/tooling files.
- **If a bug fix**: n/a. The original defects are listed in 1.13.
- **Rule changes**: all rules are introduced here: file patterns (1.4), expected types/completeness (1.4) [GST RULE], Links mapping (1.6) [GST RULE borderline], extraction mapping (1.7) [GST RULE, TIME-VARYING], state-code map (1.9) [GST RULE, TIME-VARYING], abbreviations, key length 35, long-name warning >10, BB2 validation, number display.
- **Abandoned or reversed**: `cell_mappings copy.py` is already an abandoned older format at this point (deleted in 47a1492). `gst_organizer.log` is removed 3 minutes later (cbfa4a5); the pyc files are removed in 01862a0. Dead methods from older layouts (`_write_section`, `_create_details_sheet`).
- **Dependency changes**: requirements.txt created: `openpyxl>=3.1.2`, `lxml>=4.9.0` (lxml is never imported by the code; openpyxl can use it if present), commented `pytest>=7.0.0`, `black>=23.0.0`, `pylint>=2.17.0`. Third-party imports actually used: openpyxl; pywin32 modules `win32com.client`, `pythoncom` (guarded in `core/excel_handler.py:21-28`), `win32api` (lazy, guarded in `utils/helpers.py:477`), and unguarded `pythoncom`, `win32com.client`, `win32api`, `win32con`, `win32gui` in `power_query_extractor/core/report_processor.py:11-15`. **pywin32 is missing from requirements.txt.** tkinter is part of Python.
- **Discrepancies between message and diff**: none for the message; but README vs code discrepancies are listed in 1.17, and committed data files (log, pyc) were clearly not meant to be shared.

#### cbfa4a5 - 31-08-2025 22:14 - goyal-sidhant (committer: GitHub web interface) - shared - before the split on 30-10-2025
- **Message**: "Delete gst_organizer.log" (no body).
- **Files touched**: `gst_organizer.log` −5,689 (deleted). Nothing else (READ: `git show --stat`; content of the log not read, see 1.15).
- **Why**: the author noticed the run log (which holds client names, file names and folder paths) had been uploaded and deleted it on GitHub's web page 3 minutes after the first push (committer "GitHub", noreply e-mail) (HIGH).
- **What it reveals**: awareness that the log is sensitive; but deleting a file in a new commit does not remove it from history — it remains readable at edac6a3 in both repositories (HIGH). The local computer did not get this commit until 29-10-2025 (see 9c56a1c).
- **Type**: cleanup (data removal).
- **If a bug fix**: n/a.
- **Rule changes**: none.
- **Abandoned or reversed**: reverses the accidental addition of the log in edac6a3.
- **Dependency changes**: none.
- **Discrepancies**: none.

#### 6237597 - 29-10-2025 17:21 - goyal-sidhant - shared - before the split on 30-10-2025
- **Message**: "Fix critical import and threading bugs in GST Organizer". Body: "- Fixed circular import error in utils/helpers.py (STATE_CODE_MAPPING) - Fixed scanned_files attribute access error in file_handler.py (line 75) - Fixed thread safety issues in processing_handler.py GUI updates - Fixed indentation bug in long client name validation logic - Corrected non-existent app.py reference in main.py error messages - Improved progress_callback None handling in file_parser.py. These fixes resolve runtime crashes and import errors that prevented the application from launching and processing files correctly."
- **Parent**: edac6a3 (made on the local computer, which still had the log; 59 days after the first commit).
- **Files touched** (READ full diff):
  - `core/file_parser.py` (+11/−13): `__init__` adds `self.progress_callback = None` (`:30@6237597`); new method `_update_progress(current, total, message)` that calls the stored callback inside try/except and logs a warning if the callback fails (`:49-55`); `scan_folder` stores the callback (`self.progress_callback = progress_callback`) and calls `self._update_progress(...)` instead of `if progress_callback: progress_callback(...)`; about a dozen blank lines removed inside `scan_folder`. Behaviour change: an exception raised by the progress callback no longer aborts the scan (before, it propagated to `scan_folder`'s except and the scan failed).
  - `gui/handlers/file_handler.py` (+1/−1, line 75): `len(self.app.file_parser.scanned_files)` → `len(self.app.scanned_files)`. Both refer to the same dictionary object (`scan_folder` returns `self.scanned_files`, `file_parser.py:174@edac6a3`, and `FileParser.reset()` always defines it, `:34`), so the old line could not raise an attribute error and the change has no effect (READ, HIGH).
  - `gui/handlers/processing_handler.py` (+46/−30):
    - `dry_run()` (`:36-52@6237597`): every `self.app.log_message(...)` wrapped as `self.app.root.after(0, lambda: self.app.log_message(f"...", ...))`.
    - `start_processing()` (`:64-91`): long-name check now reads `max_len = int(self.app.client_name_max_length.get())`, falling back to 10 if ≤0 or on any error; uses `self.app.client_data.get(client_key, {})` and `.get('client', '')`; collects `(name, length)` pairs; message text now "Long client names detected (exceeding {max_len} chars):" + up to 5 names + "If you continue, these client names will be used in folder names." The "PROCESSING STARTED" log line wrapped in `root.after`.
    - `process_files_thread()` lines 178–239: the progress updates and log lines up to the ITC "Output:" line wrapped in `root.after(0, lambda: ...)`. Everything after line 239 (ITC success/failure lines, the whole Sales section, organisation report, summary report, final messages, `messagebox.showinfo` at 435, `messagebox.showerror` at 458, `start_btn.config` at 462) is still called directly from the worker thread.
  - `utils/helpers.py` (+1/−1, line 491): inside `get_state_code`, `from utils.constants import STATE_CODE_MAPPING` → `from .constants import STATE_CODE_MAPPING`.
- **Why** (inferred): an AI-assisted code review of the project (message style "(line 75)", "CHANGE:" comments in the code, "This is safer and easier to review") (MEDIUM), probably the first time the author returned to the code after August (HIGH: 59-day gap).
- **What it reveals**: (a) the long-name warning was meant to follow the "max length" setting rather than a fixed 10 — a business-rule change (below); (b) the threading concern was real (Tk is not thread-safe) but the fix was partial and the existing helpers already deferred work to the main thread (`UIHelpers.log_message` uses `log_widget.after(0, ...)`, `gui/utils/ui_helpers.py:23-25`; `main_window.update_progress` uses `root.after(0, ...)`, `gui/main_window.py:260-265`), so the new wrapping is an `after` inside an `after` (READ, HIGH).
- **Type**: bug fix (claimed) / refactor; in practice it also introduced a bug.
- **If a bug fix**:
  - Claimed "circular import": the import was already inside the function (comment "Import here to avoid circular import"), and `utils.constants` and `.constants` are the same module when helpers is imported as `utils.helpers` (always the case in this project). No behaviour change in normal use (READ, HIGH; RAN: `get_state_code` works the same at edac6a3 and 47a1492). Both HEADs keep the relative form (A HEAD `utils/helpers.py:598`, B HEAD `utils/helpers.py:491`).
  - Claimed "scanned_files attribute access error": not possible (above). Both HEADs keep the new form (A HEAD `gui/handlers/file_handler.py:75`; B replaced the Tk handler).
  - Claimed "thread safety": partial (above) and **regression**: each `lambda` reads its loop variables when Tk runs it later, not when it is scheduled. RAN (`test_dryrun.py`, fake app, two clients with 2 and 1 files): at 6237597 the Dry Run preview prints the second client's name twice and the second client's file name for all three "Would organize" lines; at edac6a3 it prints the correct names. Dry Run runs in the main thread, so all callbacks run after the loop finished and always show the last values (HIGH). The same pattern in the worker thread (`:178-239`) can show the wrong client/progress/file name whenever the worker gets ahead of the main loop (MEDIUM; not run, needs real timing). A HEAD fixed it properly with `_log_async`/`_progress_async`, which capture the finished message as a default argument (`gui/handlers/processing_handler.py:23-36` A HEAD). B HEAD replaced Tk with PyQt (`gui/pyqt_main_window.py:339` logs directly) — not applicable.
  - Claimed "indentation bug in long client name validation": at edac6a3 `if long_names:` was already inside the `if include_client_name...` block (12 spaces, `processing_handler.py:71@edac6a3`); the diff only removed trailing spaces from a blank line. No indentation bug existed (READ, HIGH). What did change is the threshold (below).
  - Claimed "Corrected non-existent app.py reference in main.py": **main.py is not in the diff**; `main.py` still prints "app.py" at 6237597 and at the split point (`main.py:71@47a1492`) (READ, HIGH). The reference is gone at both HEADs (A HEAD rewrote main.py in f363946, 26-07-2026; B HEAD main.py has no app.py line) — outside this range.
  - Claimed "progress_callback None handling": the old code already checked `if progress_callback:`; the real gain is that callback exceptions are caught (above). Present at both HEADs (`core/file_parser.py:29,48-52` in A and B).
  - Similar cases still unhandled at the split point: the rest of `process_files_thread` touches Tk directly; `extractor_window.py` worker thread also calls `log_message` / `update_progress` / `messagebox` directly (`:516-582@edac6a3`, `:748-832@47a1492`); the per-client toggle still uses its own `> 10` rule (`client_handler.py:188@47a1492`, still `gui/handlers/client_handler.py:191` at A HEAD) and the Setup label still says ">10" (`setup_tab.py:238@47a1492`).
- **Rule changes**: business rule "warn about long client names when 'include client name' is on": threshold **10 characters → the 'Client folder name max length' setting (default 35; fallback 10 if the value is ≤0 or unreadable)** (`processing_handler.py:64-91@6237597`). The per-client toggle warning and the label keep 10.
- **Abandoned or reversed**: the direct `log_message` calls (edac6a3) replaced by `root.after` wrappers; A later replaced these wrappers again (`_log_async`).
- **Dependency changes**: none.
- **Discrepancies between message and diff**: main.py claim false; "attribute access error" and "indentation bug" claims do not match the old code; "thread safety fixed" overstated (partial, and it introduced the late-binding bug). The README of d674386 repeats these claims as "Recent Bug Fixes (Latest Commit)".

#### 9c56a1c - 29-10-2025 17:21 - goyal-sidhant - shared - before the split on 30-10-2025
- **Message**: "Merge branch 'main' of https://github.com/goyal-sidhant/AnnualStatement" (no body).
- **Parents**: 6237597 (local work, first parent) and cbfa4a5 (the GitHub web deletion, second parent). Made 15 seconds after 6237597 (17:21:12 → 17:21:27): the author committed locally, then pulled, and git created this merge automatically (HIGH).
- **Files touched / what the merge brought in** (READ: `git show -m --first-parent --stat 9c56a1c`; the combined diff is empty, so there were no conflicts): relative to the local line, the only change is **`gst_organizer.log` deleted (−5,689)**. Relative to cbfa4a5 the merge adds 6237597's four files. Pulling this deletion also deleted the log file from the author's local project folder, because git removes a tracked file when it merges its deletion (HIGH, git behaviour).
- **Which repository was pulled**: on 29-10-2025 only one repository existed under that name. GitHub's records (checked by the coordinator) say the repository now named AnnualStatement_oldv2 (B) was created on 31-08-2025 — the day of edac6a3 and cbfa4a5 — and AnnualStatement_oldv1 (A) on 07-02-2026. So this merge pulled from the repository that is now B (HIGH by dates). Today the old name "AnnualStatement" forwards to A, which means the original was renamed (to …_oldv2) and A later held the name "AnnualStatement" before being renamed itself (MEDIUM; the forwarding rule is GitHub's, the sequence of renames is inferred).
- **Why**: routine pull before pushing (HIGH).
- **What it reveals**: the author works mostly locally and pushes to GitHub; the GitHub-side deletion had sat unmerged for two months.
- **Type**: merge (cleanup brought in).
- **If a bug fix**: n/a.
- **Rule changes**: none.
- **Abandoned or reversed**: completes the removal of the log from the working tree (it stays in history).
- **Dependency changes**: none.
- **Discrepancies**: none.

#### d674386 - 29-10-2025 21:15 - goyal-sidhant - shared - before the split on 30-10-2025
- **Message**: "Update README.md for clarity and formatting improvements" (no body).
- **Files touched**: `README.md` +239/−135 (READ full diff). New title "📊 GST File Organizer v3.0 - Production Ready"; new sections "⚠️ IMPORTANT: Prevents Excel Corruption", "🎯 Overview" (organise, generate, **Refresh** Power Query, **Extract** cells, batch, prevent corruption), "✨ Features" (adds Power Query support "(Windows only)", data extraction, dark mode, session memory), "File Type Support" (six types; "GSTR-3B Export files (supports multiple monthly filings)"), "🔒 Excel Corruption Prevention"; requirements now "Python 3.7 or higher (3.8+ recommended)", "Windows 10/11 (best)", "Microsoft Excel 2016+ for Power Query features (Windows only)"; folder tree updated to the modular layout and `launch_extractor.py`/`power_query_extractor/` (no more `gui/app.py`); a "Requirements.txt Contents" block (`README.md:104-108@d674386`) showing `openpyxl>=3.0.0` and `pywin32>=305; sys_platform == 'win32'`; naming-convention table with examples using the generic name "ABC Ltd" and "Maharashtra"; usage steps for both programs including "python launch_extractor.py … auto-loads target folder from cache … Click 'Refresh & Extract'"; output tree now `Annual Statement-DDMMYY HHMM` / `ClientName-StateCode` / `Version-DDMMYY HHMM` (correct) but still `GST_Processing_Summary_DDMMYYYY_HHMMSS` and `ITC_Report_ClientName_State_DDMMYYYY_HHMMSS` (wrong: code uses DDMMYY_HHMM and the state code); a "Cell Extraction Configuration" excerpt of `EXTRACTION_CONFIG`/`EXTRACTION_OPTIONS`; troubleshooting table (incl. `pip install pywin32` + `pywin32_postinstall.py -install`); "Debug Mode" (`logging.basicConfig(level=logging.DEBUG)`; logs `gst_organizer.log`, `pq_extractor.log`); "Known Limitations" (Excel files must be closed; templates must have specific structure; names must match exactly; Windows 260-character path limit; Power Query refresh requires Windows and Excel); "📝 Recent Bug Fixes (Latest Commit)" repeating 6237597's claims; "📄 License: This project is proprietary software for GST file management"; footer "Version: 3.0 Production Ready / Status: Stable Release / Last Updated: 29th October 2025 / Author: Advanced AI Assistant". Removed: "Version History v3.0" section.
- **Why**: documentation refresh after the code review earlier that day, AI-written (footer) (HIGH).
- **What it reveals**: the intended features of the extractor and the author's view of the workflow; also documents "Annual Statement-DDMMYY HHMM" naming.
- **Type**: docs.
- **If a bug fix**: n/a.
- **Rule changes**: none in code.
- **Abandoned or reversed**: drops the outdated `gui/app.py` tree.
- **Dependency changes**: none in requirements.txt; the README **misstates** its contents (shows openpyxl>=3.0.0 and a pywin32 line that the real file does not have).
- **Discrepancies (README vs code at this commit)**: requirements block wrong (above); "All file operations use temporary files and verification steps" / "Clean saves: No VBA macros or external links preserved" / "Backup creation: Original files are never modified" overstated (1.17); summary/report name formats wrong; extractor button is "🚀 Start Processing", not "Refresh & Extract" (`README.md:173@d674386` vs `extractor_window.py:220-232@d674386`); "auto-loads target folder from cache" does not work (1.10, O1); "Review variations tab" (`README.md:156@d674386`) — there is no variations tab in the window (variations appear only in the scan summary count and the summary workbook); Dry Run is described under the Processing tab but the button is on the Validation tab (`validation_tab.py:109-118`); "Recent Bug Fixes" repeats inaccurate claims (see 6237597).

#### 5ea77c2 - 29-10-2025 21:20 - goyal-sidhant - shared - before the split on 30-10-2025
- **Message**: "Remove shebang line from main.py" (no body).
- **Files touched**: `main.py` −1: removes line 1 `# !/usr/bin/env python3` (READ). All later lines move up by one.
- **Why**: the line was malformed (space after `#`) and meaningless on Windows (MEDIUM: tidy-up).
- **What it reveals**: nothing functional; the program is run as `python main.py`.
- **Type**: cleanup.
- **If a bug fix**: n/a. **Rule changes**: none. **Abandoned or reversed**: none. **Dependency changes**: none. **Discrepancies**: none (the removed line was not a working shebang anyway).

#### 1cb878b - 29-10-2025 21:43 - goyal-sidhant - shared - before the split on 30-10-2025
- **Message**: "Add .gitignore file to exclude unnecessary files and directories" (no body).
- **Files touched**: `.gitignore` +79 (new; READ). Groups: Python build artefacts (`__pycache__/`, `*.py[cod]`, `*$py.class`, `*.so`, `.Python`, `build/`, `develop-eggs/`, `dist/`, `downloads/`, `eggs/`, `.eggs/`, `lib/`, `lib64/`, `parts/`, `sdist/`, `var/`, `wheels/`, `*.egg-info/`, `.installed.cfg`, `*.egg`, `MANIFEST`); virtual environments (`venv/`, `env/`, `ENV/`, `.venv`); IDE (`.vscode/`, `.idea/`, `*.swp`, `*.swo`, `*~`, `.DS_Store`); project-specific (`*.log`, `gst_organizer.log`, `pq_extractor.log`, `.gst_organizer_cache.json`, `*.cache`, `.cache/`, `temp/`, `logs/`); Excel temporary files (`~$*.xlsx`, `~$*.xls`, `*.tmp`, `temp_gst_*`, `backup_*`); output folders (`/Annual Statement-*/`, `/output/`, `/target/`); "User data and sensitive information" (`/mnt/user-data/`, `*.bak`, `config/last_session.json`); test files (`test_files/`, `test_data/`, `*.test.xlsx`); Windows (`Thumbs.db`, `ehthumbs.db`, `Desktop.ini`); "Power Query cache" (`*.odc`, `*.iqy`, `*.dqy`, `*.xlquery`). No newline at end of file.
- **Why**: stop committing compiled files and logs (followed 3 minutes later by 01862a0) (HIGH).
- **What it reveals**: `/mnt/user-data/` is the upload/output folder of Claude's web sandbox, so this file was most likely generated by Claude on claude.ai (MEDIUM). Several entries match the program's names: `temp/`, `logs/` (created by `main.py`), `temp_gst_*` and `backup_*` (the unused `EXCEL_SAFETY` prefixes — the real backup files are named `<stem>_backup_<ts>` and are **not** matched by `backup_*`), `config/last_session.json` and `.cache/` (two of the extractor's cache candidates), `.gst_organizer_cache.json` (the organiser's cache name, although that file lives in the home folder, not the repository). The extractor's relative `gst_organizer_cache.json` (no dot) is not ignored. `.claude/` is not ignored, so bd176bf's `.claude/settings.local.json` gets committed the next morning.
- **Type**: tooling.
- **If a bug fix**: n/a. **Rule changes**: none. **Abandoned or reversed**: none. **Dependency changes**: none.
- **Discrepancies**: none (the message is accurate). Note `lib/` would ignore any folder named lib anywhere (none exists).
- **HEAD check**: B HEAD `.gitignore:82` adds `.claude` (B removed the settings file in c258047, 07-02-2026); A HEAD still tracks `.claude/settings.local.json` (unchanged since bd176bf).

#### 01862a0 - 29-10-2025 21:46 - goyal-sidhant - shared - before the split on 30-10-2025
- **Message**: "Remove __pycache__ directories from Git tracking" (no body).
- **Files touched**: 70 `.pyc` files deleted (binary; byte sizes 220 to 98,247; listed in 1.16) (READ: `--stat`, `--name-status`; numstat shows no text lines).
- **Why**: follow-up to the new `.gitignore`; `git rm --cached` style removal (HIGH).
- **What it reveals**: see 1.16 (the pre-git history, the never-committed `gui/app.py`, two Python versions, Windows development, several project-folder locations). The removed files remain in history (edac6a3 to 1cb878b) with their embedded folder paths (HIGH).
- **Type**: cleanup.
- **If a bug fix**: n/a. **Rule changes**: none. **Abandoned or reversed**: reverses the accidental inclusion in edac6a3. **Dependency changes**: none. **Discrepancies**: none.

#### 1c48722 - 29-10-2025 22:18 - goyal-sidhant - shared - before the split on 30-10-2025
- **Message**: "Remove redundant text color handling in client tree display" (no body).
- **Files touched**: `gui/handlers/client_handler.py` −6 (READ): inside `update_client_tree()` removes the block `# Ensure text color is correct / if not self.app.dark_mode.get(): tag 'complete' ... for col in columns: set(item, col, set(item, col)) / else: tag 'incomplete'` (edac6a3 lines 54-59; the `else:` at edac6a3 line 60 stays and now belongs to `if client_info['status'] == 'Complete':`). After the change: Complete → tag `complete`, anything else → tag `incomplete` (`client_handler.py:51-55@1c48722`).
- **Why**: rows were coloured wrongly (HIGH, see below); the commit message calls it "redundant", which understates it.
- **What it reveals**: a real-use observation that complete and incomplete clients looked the same.
- **Type**: bug fix (message says cleanup).
- **If a bug fix**: what broke — RAN (`test_tags.py`, real `ClientHandler` with one complete and one incomplete client): at edac6a3, light mode tagged both rows `complete` (incomplete clients appeared light green), dark mode tagged both `incomplete`; at 1c48722 both modes give complete/incomplete correctly. How fixed: proper fix (the second, overriding `if/else` removed; the no-op "force text black" loop also removed). Similar cases: the row colours are set only when the tree is rebuilt; `toggle_dark_mode` re-configures the tag colours (`main_window.py:169-171`, `dark_mode_manager.py:192-212`) — consistent. HEAD: A HEAD keeps the corrected logic (READ: `gui/handlers/client_handler.py:52-55` A HEAD); B uses PyQt (not applicable).
- **Rule changes**: UI rule "Complete = green row, otherwise orange row" now actually applied.
- **Abandoned or reversed**: removes edac6a3's overriding block.
- **Dependency changes**: none.
- **Discrepancies**: message says "redundant text color handling"; the diff changes visible behaviour (fixes row colours).

#### 4c42758 - 29-10-2025 22:25 - goyal-sidhant - shared - before the split on 30-10-2025
- **Message**: "Add mouse wheel scrolling support for main and log text areas" (no body).
- **Files touched**: `power_query_extractor/gui/extractor_window.py` +30 (READ): in `create_setup_tab()` after the client-list canvas: `_on_mousewheel(event)` scrolls the canvas by `-1*(event.delta/120)` units only if the canvas is shorter than the list, returns "break"; bound to the canvas and the scrollable frame; `<Enter>` on the canvas does `bind_all("<MouseWheel>")`, `<Leave>` does `unbind_all("<MouseWheel>")` (so the wheel works over the client check-boxes). In `create_processing_tab()`: `_on_log_mousewheel` scrolls the log text by the same amount if it is not empty; bound to the log widget.
- **Why**: the extractor's client list (a canvas with check-boxes) did not scroll with the wheel (HIGH). The message says "main and log text areas"; the "main" area is the client list.
- **What it reveals**: the extractor was used with many clients (a list long enough to need scrolling) (MEDIUM).
- **Type**: feature (usability).
- **If a bug fix**: n/a (usability gap). Note: `<MouseWheel>` with `delta/120` is the Windows/macOS event; Linux sends Button-4/5, so on Linux the wheel still does not scroll (READ, HIGH; irrelevant on Windows). `unbind_all("<MouseWheel>")` on leaving also removes any other global wheel binding in this window (there is none in the extractor).
- **Rule changes**: none. **Abandoned or reversed**: none. **Dependency changes**: none.
- **Discrepancies**: minor wording only.

#### bd176bf - 30-10-2025 09:04 - goyal-sidhant - shared - before the split on 30-10-2025
- **Message**: "Enhance report processing UI with configurable wait time and file suffix options; add ITC/Sales report selection functionality" (no body).
- **Files touched** (READ full diff):
  - `.claude/settings.local.json` +9 (new): `{"permissions": {"allow": ["Bash(python -m py_compile:*)"], "deny": [], "ask": []}}`. This is the per-user settings file of **Claude Code** (Anthropic's AI coding tool): it lets the tool run `python -m py_compile <anything>` (a syntax check that compiles a file without running it) without asking each time. Nothing else is allowed; deny/ask lists empty. **It holds no password, token, key, path or personal data** (READ, HIGH). By Claude Code convention "settings.local.json" is personal and usually not committed. Its presence shows Claude Code was used for this commit's changes (MEDIUM-HIGH). Status: unchanged at A HEAD; removed at B HEAD (c258047, 07-02-2026) and `.claude` ignored (B `.gitignore:82`).
  - `power_query_extractor/core/report_processor.py` +41/−17: `__init__` adds `self.wait_time = 10` and `self.suffix_pattern = "_Refreshed_{timestamp}"` (`:27-28@bd176bf`). `process_client()` now honours `client_data['process_itc']` / `['process_sales']` (default True "for backward compatibility"); a report not selected gets `{'status': {'success': False, 'error': 'Not selected for processing'}, 'data': {}, 'extracted_values': {}}` (`:39-70`). `_process_report()` builds the refreshed name as `<stem>` + suffix with `{timestamp}` replaced by `DDMMYY_HHMM` (`:91-94`). `_refresh_power_query_simple()`: first wait becomes `self.wait_time` seconds (was a fixed 10; log text "Waiting... i/N seconds") (`:256-261`); the "extended wait" becomes `max(10, self.wait_time)` seconds (was 10) (`:285-289`). The 5-second error-dialog check in between is unchanged. Whitespace tidied.
  - `power_query_extractor/gui/extractor_window.py` +362/−145: Setup tab re-laid out in two columns. Left column (fixed 400 px): Target Folder (label, entry, Browse, Scan); new "Processing Options": "Power Query Wait Time" `Spinbox` from 5 to 60, default 10, label "seconds (Min: 5, Max: 60)" (`:170-198@47a1492`), no input validation; "Refreshed File Suffix" entry, default `_Refreshed_{timestamp}`, hint "Use {timestamp} for datetime" (`:200-223`), no validation; "Skip Refresh (Use existing refreshed files)" checkbox moved here; "🚀 Start Processing" button moved to the bottom of the left column. Right column "Select Clients & Reports": buttons Select All, Deselect All, **ITC Only**, **Sales Only**; header row "Client Name | ITC | Sales | Last Refresh Status"; per client a name label, an ITC checkbox and a Sales checkbox (ticked if that report exists; disabled if it does not), and a status label from `get_refresh_status()` (`:638-660`: newest `ITC_Report_*_Refreshed_*.xlsx` / `Sales_Report_*_Refreshed_*.xlsx` by modification time, shown as `YYYY-MM-DD HH:MM`, else "Never") formatted by `format_refresh_status()` (`:662-671`). `start_processing()` (`:708-746`) collects clients with at least one ticked report, adds `process_itc`/`process_sales`, copies wait time and suffix into the processor; warning text "Please select at least one report to process". `_process_clients_thread()` (`:748-801`) logs wait time and suffix, "Processing <name> - Reports: ITC, Sales", and "⏭️ … Skipped" for unticked reports.
- **Why**: real-use needs (HIGH by content): some Power Query refreshes needed more (or less) than the fixed 10 s; some clients need only the ITC or only the Sales report refreshed; the author wanted a different name for refreshed copies and to see when each report was last refreshed.
- **What it reveals**: the fixed-wait keystroke refresh (1.11) was the main pain point; the author tuned waits instead of detecting completion. Report selection lets the author re-run a single report without redoing the other.
- **Type**: feature (+ tooling file).
- **If a bug fix**: n/a, but it introduces mismatches (READ, HIGH):
  1. A custom suffix is not honoured by the readers: the Skip-Refresh lookup still globs `<stem>_Refreshed_*.xlsx` (`report_processor.py:103@47a1492`) and the status column globs `*_Refreshed_*` (`extractor_window.py:645,652@47a1492`). With any suffix that does not contain `_Refreshed_`, refreshed copies are never found ("Never refreshed"; Skip Refresh refreshes again). A HEAD fixed this with a shared helper (`power_query_extractor/core/refresh_naming.py:61` "Replaces the hardcoded glob(f"{stem}_Refreshed_*.xlsx") lookup"); B HEAD still has the hard-coded glob (`power_query_extractor/core/report_processor.py:102`).
  2. An empty suffix makes the refreshed name equal to the original; `shutil.copy2` then raises "same file", the error is caught and the report is marked failed (`report_processor.py:94,120-121,157-165@47a1492`) (READ, HIGH; not run). A suffix without `{timestamp}` makes every run overwrite the same refreshed file (MEDIUM). Characters not allowed in file names cause a copy error (MEDIUM).
  3. The wait-time spinbox accepts typed values outside 5–60; an empty or non-numeric entry makes `self.wait_time_var.get()` raise inside `start_processing` and nothing starts (READ, MEDIUM — Tk IntVar raises TclError on non-integers).
  4. The `ITC_Report_*.xlsx` glob used to find the original report (`report_processor.py:76@47a1492`) also matches refreshed copies of any suffix (1.10).
- **Rule changes**: business rules: refresh waits **10 s + ~5 s checks + 10 s → N s (default 10, 5–60 by spinbox) + ~5 s checks + max(10, N) s**; refreshed-copy name `_Refreshed_{DDMMYY_HHMM}` → user pattern; per-report selection with "Not selected for processing" status (in the consolidated workbook such a report shows ✗ and "Missing" for all its values, `data_consolidator.py:110-129`).
- **Abandoned or reversed**: the single "select client" checkbox per client (edac6a3) replaced by two per-report checkboxes; the old "ITC ✓ / Sales ✓" labels removed.
- **Dependency changes**: none (no new imports).
- **Discrepancies**: the message is accurate; it does not mention `.claude/settings.local.json`.

#### 47a1492 - 30-10-2025 14:38 - Sidhant - shared - before the split on 30-10-2025 (the split point)
- **Message**: "Add new ITC report mappings for Reco Diff, Table 6J Match, Table 8D Limits, IMS Reco, Extra In Purchase, and Extra In GSTR2B sheets" (no body).
- **Author identity**: "Sidhant" (same e-mail as "goyal-sidhant"); the only commit under this name. After the split A uses "Sidhant Goyal" (32 commits) and "goyal-sidhant" (2); B uses "goyal-sidhant" (5). Suggests a different computer or git client for this one commit (LOW-MEDIUM).
- **Files touched** (READ full diff):
  - `power_query_extractor/config/cell_mappings.py` +126: six new output sheets in `EXTRACTION_CONFIG`, all `report_type 'ITC'`, all `calculation None`. [GST RULE] TIME-VARYING (template cell positions and GSTR-9 table references):
    - `Reco Diff` (`:105-127@47a1492`): `Reco`!C37 → "IGST Diff"; !D37 → "CGST Diff"; !E37 → "SGST Diff".
    - `Table 6J Match` (`:130-140`): `T6 of R9`!K29 → "Table 6J Match".
    - `Table 8D Limits` (`:143-153`): `T8 of R9`!G6 → "Table 8D Limits".
    - `IMS Reco` (`:156-178`): `IMS_Reco`!AA1 → "IGST"; !AD1 → "CGST"; !AG1 → "SGST".
    - `Extra In Purchase` (`:181-203`): `ExtraInPurchase`!O1 → "IGST"; !P1 → "CGST"; !Q1 → "SGST".
    - `Extra In GSTR2B` (`:206-228`): `ExtraInGSTR2B`!S1 → "IGST"; !T1 → "CGST"; !U1 → "SGST".
    - (The sheet names "T6 of R9" / "T8 of R9" and the column names refer to GSTR-9 Tables 6J and 8D — the author's domain; I describe only where the code reads.)
  - `power_query_extractor/config/cell_mappings copy.py` −63 (deleted): the old-format mapping (1.7).
- **Why**: the ITC template had grown new check sheets (Reco, T6 of R9, T8 of R9, IMS_Reco, ExtraInPurchase, ExtraInGSTR2B) and the author wanted their key cells in the consolidated workbook (HIGH by content).
- **What it reveals**: the ITC report template's sheet names and check cells as of 30-10-2025; the consolidated workbook now has 9 data sheets + Summary. **Silent wrong-value bug triggered** (RAN, HIGH): the new columns named exactly `IGST`, `CGST`, `SGST` are found by `data_consolidator.py:142-145` as substrings of the first field name containing them — `GSTR-3B_Difference_Diff_IGST` / `_CGST` / `_SGST` — so the IMS Reco, Extra In Purchase and Extra In GSTR2B sheets receive the values of `Diff GSTR-3B`!J15/K15/L15. The exact-name fallback (`:147-150`) is reached only when the substring loop found nothing or a None value. RAN (`test_consolidator.py`, every extracted value set to its own "sheet!cell" label): at 47a1492 the three sheets show `Diff GSTR-3B!J15`, `!K15`, `!L15`; the other six sheets show their own cells. Same result at A HEAD and B HEAD (copies of each HEAD's `power_query_extractor`; their `report_processor.py` builds the extracted-value keys the same way, A HEAD `:556-588`, B HEAD `:507-539`, and both `cell_mappings.py:276` use the same field-name rule). If the GSTR-3B Difference sheet is missing in a report, its fields are None and the fallback gives the right cells (READ).
- **Type**: feature (configuration) + cleanup.
- **If a bug fix**: n/a; it exposes O2 (above). Similar unhandled case: any future column name that is a substring of an earlier field name (for example "Taxable_Value" written without spaces) would be affected the same way (READ, HIGH).
- **Rule changes**: [GST RULE] TIME-VARYING: 14 new extraction cells as listed. Removed: the dormant old-format mapping file.
- **Abandoned or reversed**: `cell_mappings copy.py` (old format, from 21-06-2025 according to 1.16) finally deleted.
- **Dependency changes**: none.
- **Discrepancies**: the message lists the six new sheets correctly; it does not mention deleting the copy file.

---

## 3. File evolution (every file touched in edac6a3 .. 47a1492; no renames happened)

| File | Created | Changed in | Deleted | How its role shifted |
|---|---|---|---|---|
| `README.md` | edac6a3 (260 lines) | d674386 (+239/−135 → 364 lines) | — | v3.0 guide written for the single-file `gui/app.py` era → guide for the modular organiser plus the extractor, with an AI-written "Recent Bug Fixes" list and footer; several statements do not match the code (1.17, d674386) |
| `requirements.txt` | edac6a3 | — | — | Unchanged: openpyxl, lxml (pywin32 missing) |
| `main.py` | edac6a3 (83) | 5ea77c2 (−1, malformed shebang) | — | Unchanged role; still lists `gui/app.py` in its error text at the split |
| `launch_extractor.py` | edac6a3 (30) | — | — | Unchanged |
| `core/__init__.py`, `core/excel_handler.py`, `core/file_organizer.py` | edac6a3 | — | — | Unchanged |
| `core/file_parser.py` | edac6a3 (373) | 6237597 (+11/−13 → 371) | — | Progress callback stored and wrapped in try/except; blank lines removed (line numbers: +9 from `parse_filename` to `scan_folder` — e.g. edac6a3 line 48 → 57, 128 → 137 — then −2 from `_process_file` onward, edac6a3 line 181 → 179) |
| `utils/__init__.py`, `utils/constants.py` | edac6a3 | — | — | Unchanged (all patterns, mappings, state map identical at the split) |
| `utils/helpers.py` | edac6a3 (537) | 6237597 (+1/−1) | — | Relative import inside `get_state_code` only |
| `gui/__init__.py`, `gui/main_window.py`, `gui/tabs/*`, `gui/utils/*`, `gui/widgets/*`, `gui/handlers/__init__.py` | edac6a3 | — | — | Unchanged |
| `gui/handlers/client_handler.py` | edac6a3 (259) | 1c48722 (−6 → 253) | — | Row-colour bug fixed |
| `gui/handlers/file_handler.py` | edac6a3 (151) | 6237597 (+1/−1) | — | No behaviour change |
| `gui/handlers/processing_handler.py` | edac6a3 (447) | 6237597 (+46/−30 → 463) | — | Part of the GUI updates deferred with `root.after` (introducing the late-binding bug); long-name threshold now follows the max-length setting |
| `power_query_extractor/__init__.py`, `extractor_main.py`, `config/__init__.py`, `core/__init__.py`, `gui/__init__.py` | edac6a3 | — | — | Unchanged |
| `power_query_extractor/config/cell_mappings.py` | edac6a3 (152) | 47a1492 (+126 → 278) | — | 3 → 9 output sheets (all new ones ITC) |
| `power_query_extractor/config/cell_mappings copy.py` | edac6a3 (63; already obsolete, dated 21-06-2025 by the pyc evidence) | — | 47a1492 | Old-format mapping kept as a backup copy, then removed |
| `power_query_extractor/core/data_consolidator.py` | edac6a3 (432) | — | — | Unchanged, but its substring look-up becomes harmful after 47a1492 (O2) |
| `power_query_extractor/core/report_processor.py` | edac6a3 (736) | bd176bf (+41/−17 → 760) | — | Configurable waits and refreshed-file suffix; per-report selection |
| `power_query_extractor/gui/extractor_window.py` | edac6a3 (601) | 4c42758 (+30), bd176bf (+362/−145 → 848) | — | Single client list → two-column screen with options and per-report check-boxes plus last-refresh status |
| `gst_organizer.log` | edac6a3 (5,689) | — | cbfa4a5 (on GitHub); merged into the local line by 9c56a1c | Accidental run-log commit; remains in history |
| 70 × `**/__pycache__/*.pyc` | edac6a3 | — | 01862a0 | Accidental build leftovers; remain in history |
| `.gitignore` | 1cb878b (79) | — | — | New; ignores logs, caches, pyc, output folders, temp files |
| `.claude/settings.local.json` | bd176bf (9) | — | — (still at A HEAD; removed in B by c258047) | Claude Code permission file (allows `python -m py_compile`) |

Python totals: 37 files / 6,954 lines at edac6a3 → 36 files / 7,295 lines at 47a1492 (READ: git numstat; the deleted file is `cell_mappings copy.py`).

---

## 4. Rule evolution (edac6a3 .. 47a1492)

| Commit | Rule | Old → new | Where |
|---|---|---|---|
| edac6a3 | All original rules (file patterns, expected types, completeness/duplicates, Links mapping, extraction mapping, state-code map and fallback, abbreviations, key length 35, long-name warning >10, BB2 validation, number display, "Missing" text) | introduced | section 1.4–1.9 |
| 6237597 | Long-name warning before processing (business rule) | `> 10` characters → `> client_name_max_length` (default 35; 10 if ≤0 or unreadable) | `gui/handlers/processing_handler.py:64-91@6237597`; per-client toggle and Setup label still say 10 (`client_handler.py:188@47a1492`, `setup_tab.py:238@47a1492`) |
| 1c48722 | Row colour = completeness (UI rule) | effectively "all one colour" → Complete green / otherwise orange | `gui/handlers/client_handler.py:51-55@1c48722` |
| bd176bf | Power Query wait after the refresh keystrokes (business/technical rule) | fixed 10 s → user value N (spinbox 5–60, default 10) | `power_query_extractor/core/report_processor.py:256-261@47a1492` |
| bd176bf | Extended wait | fixed 10 s → `max(10, N)` s | `report_processor.py:285-289@47a1492` |
| bd176bf | Refreshed-copy name | `<stem>_Refreshed_<DDMMYY_HHMM>.xlsx` → `<stem><user suffix with {timestamp}>.xlsx` (default unchanged) | `report_processor.py:91-94@47a1492`; readers still hard-code `_Refreshed_` (`:103`, `extractor_window.py:645,652`) |
| bd176bf | Which reports are processed | always both → per client ITC and/or Sales; unselected → status "Not selected for processing" → ✗ and "Missing" in the consolidated sheets | `report_processor.py:39-70@47a1492`, `extractor_window.py:708-746@47a1492`, `data_consolidator.py:110-129` |
| 47a1492 | **[GST RULE] TIME-VARYING** extraction cells | 3 output sheets (8 cells) → 9 output sheets (22 cells): added Reco!C37/D37/E37 ("IGST/CGST/SGST Diff"), T6 of R9!K29 ("Table 6J Match"), T8 of R9!G6 ("Table 8D Limits"), IMS_Reco!AA1/AD1/AG1, ExtraInPurchase!O1/P1/Q1, ExtraInGSTR2B!S1/T1/U1 ("IGST/CGST/SGST") | `power_query_extractor/config/cell_mappings.py:105-228@47a1492` |
| 47a1492 | Old-format mapping (dormant) | removed | `cell_mappings copy.py` deleted |

Unchanged in the whole range (READ: `git diff edac6a3 47a1492` lists no change to these files): the six file-name patterns, `EXPECTED_FILE_TYPES`, the GSTR-3B "unlimited files" exception, the Links-sheet cell map, the folder structure and name formats, the [GST RULE] state-code map and fallback, the abbreviation rules, the BB2 validation rule, the original 8 extraction cells.

Pattern: rules were added incrementally after real use (bd176bf, 47a1492), not corrected; no GST rule was modified or removed in this range — only extended (47a1492) (HIGH).

---

## 5. Edge cases discovered in real use (from bug-fix and real-use commits, plus pre-git traces in the code)

- **Row colours misleading** (1c48722): incomplete clients looked complete in light mode. Proper fix. Still correct at A HEAD; n/a in B.
- **Wheel scrolling in the extractor's client list** (4c42758): many clients → list longer than the window. Windows-only event; fine on Windows.
- **Power Query refresh time varies** (bd176bf): fixed waits too short or too long → user-set wait. Underlying cause (no completion detection; keystroke refresh) not addressed (1.11). Still keystrokes at the split.
- **Need to refresh only one report per client** (bd176bf) and **see when each was last refreshed**; **need a different name for refreshed copies** — introduced the suffix/lookup mismatch (still in B HEAD `power_query_extractor/core/report_processor.py:102`; fixed in A HEAD).
- **Long client names in folders** (6237597): the default 35-character key length made a 10-character warning noisy → threshold tied to the setting (partial: two other places still use 10).
- **Template growth** (47a1492): new check sheets in the ITC template → new extraction cells; exposed the consolidator's substring look-up (O2), still wrong at both HEADs (RAN).
- **6237597's "runtime crashes and import errors that prevented the application from launching"**: I found no code path at edac6a3 that would crash on the points the message lists (RAN: the GUI builds on Python 3.14; `get_state_code` works); the claims look like a code-review's guesses rather than observed failures (MEDIUM). It is possible the author saw a crash from the stdout wrapper (no console) or from pywin32 — not provable from the diff (LOW).
- **Pre-git traces in the edac6a3 code** (evidence of problems met before 31-08-2025; MEDIUM): Excel busy "Call was rejected by callee" (retries, `report_processor.py:181-185@edac6a3`); Excel properties failing to set (5 retries, `:158-170`); long paths > 218 characters (`:648-664`, `excel_handler.py:303-315`); OneDrive/SharePoint sync and files open elsewhere (`report_processor.py:720-729`); window activation failing (`:198-210`); Excel error dialogs during refresh (`:240-277`); duplicate-download names like `(2) GSTR3B-...` (`constants.py:31`); Excel corruption in earlier versions (README); waits increased step by step ("Increased from 0.5", "Increased from 3 to 5", "WAIT 7 SECONDS AS REQUESTED" with a 10-second loop).

---

## 6. Abandoned approaches and reversals

| What was tried | Evidence | What replaced it | Lesson |
|---|---|---|---|
| One large GUI file `gui/app.py` (89,275 bytes) | pyc headers (1.16), README@edac6a3 tree, `main.py:72@edac6a3` | Modular GUI on 23-07-2025 (`gui/main_window.py:4` "Modularized version") | Pre-git; reasons lost |
| openpyxl-only report creation (earlier versions corrupted files / lost Power Query) | `excel_handler.py:2-3,30,96-101,230-235`, README "corruption issues from previous versions" | Excel COM on Windows; openpyxl kept only as a fallback with a warning | Matches Standard 12's openpyxl warning; MEDIUM-HIGH |
| Old extraction mapping format (`CELL_MAPPINGS` per sheet/cell, 21-06-2025) | `cell_mappings copy.py@edac6a3`, pyc size match (1.16) | `EXTRACTION_CONFIG` (output sheet → list of mappings) with a generated backward-compatible `CELL_MAPPINGS` (`cell_mappings.py:133-152@edac6a3`); copy deleted in 47a1492 | Old consumer code still reads `CELL_MAPPINGS` |
| Old consolidated layout (one sheet with an ITC section, a Sales section and a details sheet) | dead `_write_section`, `_create_details_sheet` (`data_consolidator.py:250-414`), smaller 3.12 pyc (25-06-2025) | One sheet per configured output sheet + Summary (`create_report`) | Dead methods remain at the split |
| Committing logs and pyc | edac6a3 | cbfa4a5 + 9c56a1c (log), 1cb878b + 01862a0 (pyc) | Data stays in history |
| Direct Tk calls from the worker thread | edac6a3 | `root.after(0, lambda ...)` in part of the code (6237597) — itself replaced in A by `_log_async` (A HEAD) | Late-binding regression |
| Fixed 10-second refresh wait (and an even earlier 7-second one, per the comment) | `report_processor.py:235-238@edac6a3` | User-set wait (bd176bf) | Still a guess, not a completion check |
| Fixed `_Refreshed_` suffix | edac6a3 | User pattern (bd176bf), readers not updated | Fixed in A later |
| One "select client" checkbox per client in the extractor | `extractor_window.py:425-464@edac6a3` | Two per-report check-boxes (bd176bf) | — |
| Fixed long-name threshold 10 | edac6a3 | Setting-based threshold in one of three places (6237597) | Inconsistent |
| Partial / unfinished ideas still in the code at the split | `EXTRACTION_OPTIONS['auto_refresh_if_missing']`, `['show_timestamp']` (never read), calculation support (never configured), `EXCEL_SAFETY` (never read), `validate_results` "critical files" (never called), `cleanup_on_error` (never called), `self.excel_app` "don't quit if Excel was already open" (never set) | — | Started, never wired in |

---

## 7. Dependency evolution

- `requirements.txt` is identical at edac6a3 and 47a1492: `openpyxl>=3.1.2`, `lxml>=4.9.0`; commented `pytest>=7.0.0`, `black>=23.0.0`, `pylint>=2.17.0`. No versions pinned (READ).
- Used but missing from requirements: **pywin32** (`win32com.client`, `pythoncom`, `win32api`, `win32con`, `win32gui`). Guarded import in `core/excel_handler.py:20-28` and `utils/helpers.py:476-483`; unguarded in `power_query_extractor/core/report_processor.py:11-15` (crashes on import without it; RAN on Linux). Inconsistent import handling between the two programs (HIGH).
- Listed but never imported: lxml (openpyxl may use it for speed if installed) (READ, HIGH).
- Standard library: tkinter, logging, threading, json, pathlib, shutil, hashlib, tempfile, platform, re, datetime, io, os, sys, time, warnings (`excel_handler.py:17` silences **all** Python warnings with `warnings.filterwarnings('ignore')`).
- No new third-party import was added by any commit after edac6a3 (READ: all diffs). `.claude/settings.local.json` (bd176bf) concerns a development tool, not a runtime dependency.
- README (d674386) states requirements.txt contains `openpyxl>=3.0.0` and `pywin32>=305; sys_platform == 'win32'` — false at every commit in this range.
- Python version: README says 3.7+; the developer machine(s) used Python 3.12 and 3.13 (pyc evidence, 1.16); RAN: the organiser's window builds on Python 3.14 (Linux).

---

## 8. Possible client names (locations only; nothing copied)

- `edac6a3:gst_organizer.log` — the whole file (5,689 lines) contains client names, state names, source file names and folder paths from real use (from the logging templates it matches: `core/file_parser.py:71,117,118,123,252@edac6a3`, `core/excel_handler.py:72-75,176,558-559,617-656@edac6a3`, `core/file_organizer.py:53,137,275,357@edac6a3`, `utils/helpers.py:180,199@edac6a3`). Not read as text.
- `edac6a3:**/__pycache__/*.pyc` (70 files, also present in cbfa4a5, 6237597, 9c56a1c, d674386, 5ea77c2, 1cb878b) — embedded Windows folder paths of the developer computer; folder names inside those paths could include client or personal names. Not reproduced or inspected beyond hashes and lengths.
- Code, comments, docs and commit messages in the range: no possible client or business name found. The only company-like name is the generic README example "ABC Ltd" (`README.md:128-133@d674386`), which is an obvious placeholder, not a client. No GSTIN- or PAN-shaped strings found (RAN: regex search of both trees).

---

## 9. Questions for the author

1. Are the GitHub repositories private? Both histories still contain `gst_organizer.log` (client names, file names, folder paths) and compiled files with your computer's folder paths at edac6a3. Removing them needs a history rewrite of both repositories — your decision.
2. Do you still have the pre-git versions (June–August 2025), in particular the single-file `gui/app.py`? Do you remember why the GUI was split on 23-07-2025?
3. The extractor refreshes Power Query by pressing Alt-A, R, A in a visible Excel ("Keep it visible to avoid issues"). Did you try Excel's own refresh command through COM (`RefreshAll`) before, and what went wrong? This decides whether Standard 12's refresh method can replace it.
4. Have you seen the consolidated extraction workbook show the same numbers in "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" as in "GSTR-3B Difference"? (The code does this since 47a1492, at both HEADs.)
5. [GST RULE] Please confirm the template cells the extractor reads (section 1.7 and 47a1492) are still where your templates keep those figures today (TIME-VARYING).
6. What is `Info!BB2` in your ITC/Sales templates (the code calls it "Client Name") and is "not empty / not 0 / not an error" the right test that a refresh worked?
7. In file names, do you write the state as a full name (for example "Maharashtra"), a letter code ("MH") or the numeric GST state code? The code maps full names; letter codes work through a fallback; numeric codes crash client selection.
8. Do any client names contain a hyphen? The file patterns cannot read them.
9. What should "Resume" do? The code always makes a new Version folder, so Resume behaves like Re-Run.
10. When the organiser ran, did your own open Excel ever disappear or close? (The code attaches to an open Excel, hides it and calls Quit on it.)
11. Do you use .xltx templates? The organiser copies the template file to an .xlsx name; whether Excel opens such a copy without complaint was not testable here.
12. Were the 29-10-2025 code review (6237597), the README rewrite and the `.gitignore` produced with an AI tool (the `.gitignore` names Claude's web sandbox folder `/mnt/user-data/`)? Was `.claude/settings.local.json` meant to be committed?
13. Commit 47a1492 is under the name "Sidhant" (all others "goyal-sidhant" or later "Sidhant Goyal") — was it made on a different computer?
14. Were two computers (Python 3.12 and 3.13) used in 2025?
15. Is the 10-character long-name warning in the per-client toggle still wanted, now that the start-up warning uses the max-length setting?

---

## 10. The program at the split point (47a1492, 30-10-2025)

What both A and B inherited. Everything in section 1 still applies except the changes listed here; line numbers below are @47a1492.

### 10.1 Files at the split point (36 Python files, 7,295 lines; plus README 364, requirements 13, `.gitignore` 79, `.claude/settings.local.json` 9)

`main.py` 82 · `launch_extractor.py` 30 · `core/__init__.py` 13 · `core/excel_handler.py` 661 · `core/file_organizer.py` 409 · `core/file_parser.py` 371 · `utils/__init__.py` 38 · `utils/constants.py` 275 · `utils/helpers.py` 537 · `gui/__init__.py` 11 · `gui/main_window.py` 365 · `gui/tabs/__init__.py` 8 · `gui/tabs/setup_tab.py` 377 · `gui/tabs/validation_tab.py` 247 · `gui/tabs/processing_tab.py` 117 · `gui/handlers/__init__.py` 8 · `gui/handlers/client_handler.py` 253 · `gui/handlers/file_handler.py` 151 · `gui/handlers/processing_handler.py` 463 · `gui/utils/__init__.py` 9 · `gui/utils/cache_manager.py` 56 · `gui/utils/dark_mode_manager.py` 212 · `gui/utils/status_bar.py` 37 · `gui/utils/ui_helpers.py` 63 · `gui/widgets/__init__.py` 7 · `gui/widgets/collapsible_frame.py` 41 · `gui/widgets/title_bar.py` 51 · `power_query_extractor/__init__.py` 3 · `power_query_extractor/extractor_main.py` 79 · `power_query_extractor/config/__init__.py` 1 · `power_query_extractor/config/cell_mappings.py` 278 · `power_query_extractor/core/__init__.py` 1 · `power_query_extractor/core/data_consolidator.py` 432 · `power_query_extractor/core/report_processor.py` 760 · `power_query_extractor/gui/__init__.py` 1 · `power_query_extractor/gui/extractor_window.py` 848. (git line counts.) No tests, no `docs/`, no `models/`, no `pytest.ini` (those appear only after the split).

### 10.2 Programs, screens and features

- **Organiser** (`python main.py`; window "GST File Organizer v3.0", 1200x800): tabs "📁 Step 1: Setup", "✅ Step 2: Validation", "🚀 Step 3: Processing" — identical to edac6a3 except: Dry Run and the worker thread's first log lines go through `root.after` (`gui/handlers/processing_handler.py:36-52,104,178-239`) with the late-binding bug; long-name warning uses the max-length setting (`:64-91`); client rows coloured correctly (`gui/handlers/client_handler.py:51-55`). Features F1–F14 as in 1.3.
- **Power Query Extractor** (`python launch_extractor.py`; window "Power Query Extractor", 1000x700, Windows + pywin32 + Excel only): tab "Setup" — left column: Target Folder (entry, Browse, Scan), Processing Options (wait time 5–60 s default 10, refreshed-file suffix default `_Refreshed_{timestamp}`, Skip Refresh), Start Processing; right column "Select Clients & Reports": Select All, Deselect All, ITC Only, Sales Only, and per client an ITC and a Sales checkbox plus "Last Refresh Status" (`power_query_extractor/gui/extractor_window.py:95-367,552-699`). Tab "Processing": progress bar, log with wheel scrolling (`:380-457`). Flow per selected report: copy to refreshed name → visible maximised Excel → Alt-A, R, A → wait N s → error-dialog checks (5 s) → wait max(10, N) s → final check → Info!BB2 validation → save/close → wait 3 s → hidden Excel reads the configured cells → consolidated workbook `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx` in the newest Annual Statement folder (`power_query_extractor/core/report_processor.py:39-375,486-596`, `power_query_extractor/core/data_consolidator.py:20-186`).

### 10.3 File-name patterns, constants, folder/name formats

Unchanged from edac6a3: patterns in `utils/constants.py:13-62` (table in 1.4), expected types `:68-75`, folder templates `:81-90`, Links mapping `:96-120` (table in 1.6), GUI config `:126-151`, modes `:157-173`, extensions/signatures `:179-185`, messages `:191-211`, `EXCEL_SAFETY` `:217-228` (unused), state map `:234-275`. Formats: `Annual Statement-DDMMYY HHMM`, `<client-state key>`, `Version-DDMMYY HHMM`, `ITC_Report_<client>_<state code>_<DDMMYY_HHMM>.xlsx`, `Sales_Report_...`, `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx`, `_Organization_Report_<client>_<state>.txt`, refreshed copies `<stem><suffix>` (default `_Refreshed_<DDMMYY_HHMM>`), `PQ_Extraction_Report_<YYYYMMDD_HHMM>.xlsx`. Cache `~/.gst_organizer_cache.json` (organiser) vs relative `gst_organizer_cache.json` (extractor). Logs `./gst_organizer.log`, `./pq_extractor.log`.

### 10.4 Rules at the split point

- [GST RULE] TIME-VARYING state-name map and fallback (`utils/helpers.py:485-511`, `utils/constants.py:234-275`) — unchanged.
- [GST RULE] expected six file types, GSTR-3B multiple files allowed, completeness % (`core/file_parser.py:252-286`) — unchanged.
- [GST RULE, borderline] Links-sheet input map (`utils/constants.py:96-120`; values from `core/excel_handler.py:551-661`) — unchanged.
- [GST RULE] TIME-VARYING extraction map, 9 output sheets / 22 cells (`power_query_extractor/config/cell_mappings.py:33-242`):

| Output sheet | Report | Cells (input sheet!cell → column) | Lines |
|---|---|---|---|
| GSTR-3B Difference | ITC | Diff GSTR-3B!J15 → Diff IGST; !K15 → Diff CGST; !L15 → Diff SGST | 35-58 |
| GSTR-1 Difference | Sales | Diff - GSTR-1!N1 → Diff Taxable Value; !O1 → Diff IGST; !P1 → Diff CGST; !Q1 → Diff SGST | 61-89 |
| Turnover GSTR-1 | Sales | Diff - GSTR-1!I1 → Taxable Value | 92-102 |
| Reco Diff | ITC | Reco!C37 → IGST Diff; !D37 → CGST Diff; !E37 → SGST Diff | 105-127 |
| Table 6J Match | ITC | T6 of R9!K29 → Table 6J Match | 130-140 |
| Table 8D Limits | ITC | T8 of R9!G6 → Table 8D Limits | 143-153 |
| IMS Reco | ITC | IMS_Reco!AA1 → IGST; !AD1 → CGST; !AG1 → SGST (**filled with Diff GSTR-3B J15/K15/L15 by the consolidator bug, RAN**) | 156-178 |
| Extra In Purchase | ITC | ExtraInPurchase!O1 → IGST; !P1 → CGST; !Q1 → SGST (**same bug**) | 181-203 |
| Extra In GSTR2B | ITC | ExtraInGSTR2B!S1 → IGST; !T1 → CGST; !U1 → SGST (**same bug**) | 206-228 |

- `EXTRACTION_OPTIONS` (`cell_mappings.py:248-253`): skip_refresh False (changed at run time by the checkbox), auto_refresh_if_missing True (unused), show_timestamp False (unused), missing_data_text "Missing". Backward-compatible `CELL_MAPPINGS`/`FIELD_DISPLAY_NAMES` built at `:260-278` (field name = `<output sheet>_<output column>` with spaces → `_`, `:276`).
- Refresh validation Info!BB2 (`report_processor.py:597-644`); number display `#,##0.00` for non-zero numbers (`data_consolidator.py:161-162`); "Missing" when not successful (`:128-129`).
- Business rules: abbreviations (`utils/helpers.py:66-112,513-537`); key length 35 (`core/file_organizer.py:42`); long-name warning: start-up uses the setting (`gui/handlers/processing_handler.py:64-91`), per-client toggle >10 (`gui/handlers/client_handler.py:188`); refresh waits N / max(10, N) (`report_processor.py:256-289`); per-report selection (`report_processor.py:39-70`).

### 10.5 Known defects at the split point (all inherited by A and B unless their later commits changed them)

O1, O2 (now active), O3, O5–O20 from 1.13; the 6237597 late-binding bug; the bd176bf suffix/lookup mismatch and unvalidated wait/suffix inputs; Excel COM practices (1.11); start-up/console behaviour (1.12); dead code (1.14, plus `extractor_main.py`'s relative cache list); README inaccuracies (d674386); `main.py:71` still printing `app.py`.

### 10.6 What each side changed afterwards (spot checks only; detail belongs to the A/B readers)

| Item at the split | A HEAD (bf2f1dd) | B HEAD (b0e7d0e) |
|---|---|---|
| GUI toolkit | Tkinter kept | PyQt5 ("v4.0"), Tk GUI removed (a62d3cc) |
| Late-binding logging (6237597) | fixed (`gui/handlers/processing_handler.py:23-36`) | replaced by PyQt (`gui/pyqt_main_window.py:339`) |
| Consolidator substring bug (O2) | still (RAN) | still (RAN) |
| Extractor cache path (O1) | fixed | extractor folded into the main window; old `power_query_extractor/extractor_main.py` imports a removed `gui/extractor_window` module (READ: only `power_query_extractor/gui/__init__.py` exists at B HEAD) |
| Version-folder choice (O6) | fixed (mtime) | explicit selection of version folders |
| Per-client folder toggle key (O3) | fixed | not checked |
| Refreshed-suffix lookup (bd176bf) | fixed (`refresh_naming.py`) | still (`report_processor.py:102`) |
| `main.py` stdout wrapper (O10) | still (`main.py:17`) | guarded (`main.py:12-13`) |
| `.claude/settings.local.json` | kept | removed, `.claude` ignored |
| `main.py` app.py text | removed (f363946) | removed |

---

## 11. INCOMPLETE

Nothing in my assigned range is missing: all 12 commits' diffs were read in full (`git show` per commit, `-m --first-parent` for the merge), every code file added in edac6a3 was read line by line, and the changed files were read at 47a1492. Limits (not omissions):
- Windows-only behaviour (Excel COM, SendKeys, win32gui dialogs, `os.startfile`) was traced by reading only; nothing in 1.11 was executed.
- The worker-thread variant of the late-binding bug (6237597, lines 178-239) was not executed (needs real timing); the Dry Run variant was (RAN).
- HEAD-status columns in 1.13 and 10.6 are spot checks (grep/read of a few lines, two RAN tests); "not checked" items need the A/B readers.
- The `.pyc` files were analysed only through their header fields and hashed path prefixes; their code objects were not decompiled (Python 3.14 cannot load 3.12/3.13 code objects, and decompiling `gui/app` was not needed for this task).
