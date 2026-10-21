## SECTION 3: FEATURE INVENTORY

This section lists every feature of project B (AnnualStatement_oldv2, branch main, snapshot b0e7d0e of 09-02-2026, the PyQt5 window that calls itself "GST File Organizer v4.0") and documents, feature by feature, what it does, where it lives, what goes in and what comes out, how the user starts it, and one concrete worked example ("trace") followed through the code to the real result. This file is the **first part** of the section: it gives the table of all 34 features and the full entries for **B-F1 to B-F17** (start-up, Step 1 Setup, Step 2 Validation, Step 3 Processing). B-F18 to B-F34 (Step 4 Extract & Refresh and the cross-cutting features) follow in the second part.

**How the traces were made (read this once).**

- Every feature ID and rule ID (B-F…, B-R…) is the one in the project-B registry (`/root/merge-scratch/audit/B_registry.md`). Locations are `file:line` in project B at b0e7d0e, repository-relative (for example `core/file_parser.py:243`). Older code is written `path:line@commit`.
- **RAN** = the code was executed on made-up inputs and the result is quoted from the output. **READ** = the statement comes from reading the code line by line. Every inference carries a confidence word: **HIGH**, **MEDIUM** or **LOW**.
- All runs used a scratch copy of the code (`/root/merge-scratch/run/B-S3a/proj`, byte-for-byte equal to the repository; checked with `diff -rq` after the runs: no difference, no `__pycache__` written) and `/root/merge-scratch/venv-B/bin/python -B` (Python 3.14 with PyQt5 5.15, openpyxl 3.1.5). The program's per-user folder was redirected to scratch folders (`/root/merge-scratch/run/B-S3a/home/...`) so nothing touched a real home folder. Scripts are in `/root/merge-scratch/run/B-S3a/scripts/`, their printed output in `/root/merge-scratch/run/B-S3a/work/*_out.txt`.
- **Why some runs needed "stand-ins".** This computer runs Linux and has neither Excel nor pywin32 (the Python add-on that lets Python drive Windows programs such as Excel through **COM**, Windows' remote-control mechanism). Project B's window cannot even be imported without pywin32 (see B-F1). To run the window at all, small fake modules with the pywin32 names (copied from the Excel reader's run folder into `/root/merge-scratch/run/B-S3a/fakewin/`) were put first on Python's search path. In the window runs they only make the import succeed; they decide nothing. Because the computer is Linux, `core/excel_handler.py:21` sends report creation down its **openpyxl** branch (openpyxl = a Python library that writes .xlsx files without Excel) instead of the Excel branch the author's Windows PC uses. Everything else in Steps 1-3 (scan, tables, dry run, folders, copies, summary workbook, run history) is the same code on both systems. The Excel (COM) branch itself is traced by **reading**; a separate run with *recording* stand-ins shows the **order** of the requests B makes to Excel (marked "fake-COM run") - that proves the Python control flow only, never how real Excel reacts.
- **Made-up clients** used in the traces (invented names, no client data): **<Client A>** = "ClientA Traders", Maharashtra, with a GSTR-2B Reco file, three GSTR-3B files and an Annual Report but **no** IMS Reco, Sales or Sales Reco file; **<Client B>** = "ClientB Private Limited", Gujarat, with all six file types and ten monthly GSTR-3B files (15 files); **<Client C>** = "ClientC Co", Tamil Nadu, with one GSTR-3B file, two Sales files and one more GSTR-3B file whose state is written "Tamilnadu"; **<Client D>** = "Client-D Exports", Kerala, with a hyphen inside its name; **<Client E>** = "ClientA Traders and Distributors", Maharashtra (a long name). Traps placed in the same folder: an `.xlsm` file, an unrecognised "Purchase Register", a 300-byte file, a file in a sub-folder and an Excel lock file (`~$...`). The builder is `/root/merge-scratch/run/B-S3a/scripts/mk_source.py`.
- **GST rules** (anything that decides a tax figure, which return, sheet or cell a thing belongs to, a rate, limit, rounding, period, GSTIN/PAN/state-code logic, ITC eligibility or a due date) are tagged **[GST RULE]**. For them this section only says what the code does; it never says whether that is right. Values that can change over time (return names, template cell positions, sheet names, the state list) carry **⚠** so the author can check them.
- Plain-English words used throughout: **thread** (a second line of work running at the same time, so the window does not freeze while a job runs); **signal** (Qt's message from one object to another, used to send progress from a background thread to the window); **regular expression / regex** (a text pattern used to recognise file names); **glob** (a wildcard file search such as `*.xlsx`); **stem** (a file name without its extension); **JSON** (a plain-text file format the program uses for its settings).

> **Two facts shape almost every trace below.** (1) Project B's window cannot start at all unless pywin32 is installed, yet `requirements.txt:14` leaves pywin32 commented out (B-F1). (2) On the author's Windows PC, Step 3 report creation takes over an Excel the user already has open, hides it and closes it at the end (B-F13). Both are described in full in their entries.

### Table of all registry features

| ID | Feature | Status (registry) | Where it mainly lives (file:line) |
|---|---|---|---|
| B-F1 | Start-up and dependency checks | works with defects | main.py:11-59; gui/pyqt_main_window.py:35-68; requirements.txt:1-14; power_query_extractor/core/report_processor.py:11-15 |
| B-F2 | One-time import of old v3 settings (migration) | works with defects | models/cache_manager.py:93-131; gui/pyqt_main_window.py:47 |
| B-F3 | Choose source folder, target folder and templates (Setup) | works with defects | gui/views/setup_view.py:192-330, 605-647; gui/pyqt_main_window.py:174-225, 345-361 |
| B-F4 | Include client name in category-folder names (global option) and client-folder length limit | works with defects | gui/views/setup_view.py:367-388; core/file_organizer.py:42, 79-135; utils/helpers.py:513-537 |
| B-F5 | Per-client option to include the client name in folder names | dead or unreachable | core/file_organizer.py:117-126; models/app_state.py:83 |
| B-F6 | Scan source folder and recognise client files | works with defects | core/file_parser.py:126-314; utils/helpers.py:249-266; utils/constants.py:13-75; gui/controllers/scan_controller.py:14-94 |
| B-F7 | Scan progress log and scan summary | works with defects | gui/pyqt_main_window.py:234-316; gui/views/setup_view.py:519-588; gui/views/validation_view.py:117-140 |
| B-F8 | Client list with completeness status (Validation) | works with defects | gui/views/validation_view.py:231-447; core/file_parser.py:280-314 |
| B-F9 | Unrecognised files / variations list | broken | core/file_parser.py:183-252, 316-334; gui/controllers/processing_controller.py:73; core/excel_handler.py:491-522 |
| B-F10 | Export client list (EXPORT LIST button) | dead or unreachable | gui/views/validation_view.py:186-200, 221 |
| B-F11 | Dry run (preview) | works with defects | gui/pyqt_main_window.py:326-343; gui/views/validation_view.py:169-184 |
| B-F12 | Organise files into client and category folders (fresh / re-run / resume, backups) | works with defects | core/file_organizer.py:26-298; gui/controllers/processing_controller.py:42-110; utils/helpers.py:142-204 |
| B-F13 | Create ITC report from template | works with defects | core/excel_handler.py:65-306, 540-588; gui/controllers/processing_controller.py:111-122, 204-239 |
| B-F14 | Create Sales report from template | works with defects | core/excel_handler.py:65-306, 590-622; gui/controllers/processing_controller.py:124-130 |
| B-F15 | Processing summary workbook and organisation report | works with defects | core/excel_handler.py:308-538; core/file_organizer.py:300-362; gui/controllers/processing_controller.py:58-76, 132-193 |
| B-F16 | Processing progress and log | works with defects | gui/views/processing_view.py:23-166; gui/controllers/processing_controller.py:80-110, 183-202, 242-298; gui/pyqt_main_window.py:137-152, 373-415 |
| B-F17 | Processing run history and "Last Run" column | works with defects | models/run_history.py:16-148; gui/controllers/processing_controller.py:149-160; gui/views/validation_view.py:381-385 |
| B-F18 | Find each client's reports for Step 4 (latest run folder, version folder) | works with defects | gui/pyqt_main_window.py:419-504; gui/views/extraction_view.py:147-247, 291-390 |
| B-F19 | Refresh reports through Power Query | works with defects | power_query_extractor/core/report_processor.py:39-366 |
| B-F20 | Refresh validation check (Info!BB2) | works with defects | power_query_extractor/core/report_processor.py:300-330, 589-636 |
| B-F21 | Skip refresh and refreshed-file naming | works with defects | gui/views/extraction_view.py:104-145; power_query_extractor/core/report_processor.py:91-120 |
| B-F22 | Extract cells from refreshed reports | works with defects | power_query_extractor/core/report_processor.py:478-587; power_query_extractor/config/cell_mappings.py:33-278 |
| B-F23 | Consolidated extraction workbook | works with defects | power_query_extractor/core/data_consolidator.py:20-235 |
| B-F24 | Extraction progress, on-screen log and Stop | works with defects | gui/views/extraction_view.py:28-102, 249-287, 392-416; gui/controllers/extraction_controller.py:43-107 |
| B-F25 | Extraction log, statistics and run history | works with defects | gui/controllers/extraction_controller.py:97-121; models/run_history.py:73-95 |
| B-F26 | Automatic closing of Excel error dialogs | works with defects | power_query_extractor/core/report_processor.py:260-298, 368-476 |
| B-F27 | Settings cache | works with defects | models/cache_manager.py:16-89, 135-152; models/app_state.py:10-46 |
| B-F28 | Dark mode / theme | works with defects | gui/theme/theme_manager.py:25-424; gui/pyqt_main_window.py:594-597 |
| B-F29 | Step-status dots on tabs and automatic tab switching | works with defects | gui/pyqt_main_window.py:304-307, 567-590 |
| B-F30 | Standalone extractor launcher | works with defects | launch_extractor.py:1-28 |
| B-F31 | Old stand-alone extractor entry point (extractor_main.py) | broken (dead or unreachable) | power_query_extractor/extractor_main.py:1-79 |
| B-F32 | Application log file | works with defects | main.py:15-23 |
| B-F33 | Help texts, tooltips, scrolling and collapsible sections | works with defects | gui/views/setup_view.py:17-82, 150-188, 469-517; gui/views/validation_view.py:72-115 |
| B-F34 | Background jobs and run guards | works with defects | gui/controllers/scan_controller.py:44-94; gui/controllers/processing_controller.py:242-298; gui/controllers/extraction_controller.py:124-184 |

Statuses are the registry's. The traces below confirmed every status in this part; where a trace found more than the registry says, the entry says so and the point is listed for Section 7 (Gaps).

---

### B-F1: Start-up and dependency checks

**What it does.** When the user starts the program, `main.py` makes the console accept any character, sets up the program's log file, tells Python where the program's own files are, builds the four-tab window and shows it. If something the program needs is not installed, it prints a message on the console and stops.

**Where it lives.**

| Step | file:line | What the code does |
|---|---|---|
| Console text | main.py:11-13 | If a console exists (`sys.stdout` has a `buffer`), re-wraps it as UTF-8 so symbols such as ✓ can be printed. |
| Program log | main.py:15-23 | `logging.basicConfig(level=INFO, ...)` with a `FileHandler('gst_organizer.log')` - a file name without a folder, so it lands in the folder the program was started from - and a copy to the console (details in B-F32). |
| Import path | main.py:26-29 | `setup_environment()` puts the folder of `main.py` first on Python's import path. |
| Start | main.py:32-47 | Imports PyQt5 (37) and the window (38), creates the application (40), names it "GST File Organizer" version "4.0" (41-42), builds the window (44), shows it (45), runs it (47). |
| Errors | main.py:49-59 | A missing module (ImportError) prints "Missing module: …", "Please install required packages:", "  pip install PyQt5 openpyxl" and exits with code 1 (49-54); any other start-up error prints "Critical error: …" and exits 1 (56-59). |
| Window build order | gui/pyqt_main_window.py:35-68 | Default settings (39), empty scan result (40), processing state (41), settings store (42, creates `~/.gst_organizer`), run history (43), theme manager (44), old-settings import (47, B-F2), saved settings (50-51), three controllers (54-56), window, tabs, menus, signal wiring, restoring the screen values (59-63; the restore at 174-190 also fills Step 4 from disk), theme (66), log line "GST File Organizer v4.0 initialized" (68). |
| The import that needs pywin32 | gui/__init__.py:7 → gui/pyqt_main_window.py:25 → gui/controllers/__init__.py:4 → gui/controllers/extraction_controller.py:10 → power_query_extractor/core/report_processor.py:11-15 | `report_processor.py` imports `pythoncom`, `win32com.client`, `win32api`, `win32con`, `win32gui` with no guard, and it is pulled in while the window module itself is being imported. |
| Declared dependencies | requirements.txt:1-14 | `PyQt5>=5.15.0` (5), `openpyxl>=3.1.2` (8), `lxml>=4.9.0` "Optional but recommended" (11), and `# pywin32>=306` **commented out** (13-14). |

**Input → Output.** Installed Python packages + the saved settings file `~/.gst_organizer/settings.json` → the open window, **or** console text and exit code 1.

**User interaction.** The user types `python main.py` in a console (README.md:110-113). The repository has no `.bat` launcher or shortcut; `launch_extractor.py` (B-F30) is a second way in that opens the same window on Step 4.

#### B-F1 concrete trace (RAN: `/root/merge-scratch/run/B-S3a/scripts/t_f1_startup.py`, output `work/f1_out.txt`)

Four realistic starts were run as separate programs, each from its own empty start folder.

**Case A - a fresh office PC set up exactly as README.md:97-102 says (`pip install -r requirements.txt`, which installs PyQt5, openpyxl and lxml but not pywin32).**

1. main.py:12 - a console exists, so main.py:13 re-wraps the output as UTF-8.
2. main.py:16-23 - logging is configured; the `FileHandler` immediately creates `gst_organizer.log` **in the start folder** (RAN: `work/f1/A/gst_organizer.log` appeared, 83 bytes).
3. main.py:63 calls `main()`; main.py:35 puts the program folder on the import path; main.py:37 imports PyQt5 (succeeds).
4. main.py:38 `from gui.pyqt_main_window import GSTOrganizerWindow`: Python first runs `gui/__init__.py`, whose line 7 imports the window module; the window module's line 25 imports `gui.controllers.scan_controller`, which first runs `gui/controllers/__init__.py`, whose line 4 imports the extraction controller, whose line 10 imports `report_processor`, whose line 11 is `import pythoncom` → `ModuleNotFoundError: No module named 'pythoncom'` (a kind of ImportError).
5. main.py:49 catches it; line 50 logs it; lines 51-53 print; line 54 exits with code 1. The settings folder `~/.gst_organizer` is never created, because the window constructor was never reached.

Actual output (RAN): exit code **1**; console:

```text
 Missing module: No module named 'pythoncom'

Please install required packages:
  pip install PyQt5 openpyxl
```

and the log file holds one line, `… - root - ERROR - Import Error: No module named 'pythoncom'`.

> **The advice does not fix the problem.** Both packages it names are already installed; the missing one is pywin32, which requirements.txt:14 leaves commented out. Following the README on a new PC therefore gives a program that cannot open. (HIGH, RAN.)

**Case B - the same PC with pywin32 present (here: the stand-ins), no saved settings.** The window constructor runs in the order listed above: `CacheManager()` creates `~/.gst_organizer` (models/cache_manager.py:31-33); `load_run_history` finds no file and returns an empty history (models/run_history.py:138-148); `migrate_from_old_cache` returns at once because there is no old file (models/cache_manager.py:101-102); `load_settings` finds no file and returns the defaults (models/cache_manager.py:57-62); `_restore_settings` fills the boxes with empty values and, because the target is empty, Step 4 shows "No target folder configured." (gui/pyqt_main_window.py:421-424). RAN result: window title **"GST File Organizer v4.0"**, size 1400 x 900, minimum 1200 x 800, tabs `['Step 1: Setup', 'Step 2: Validation', 'Step 3: Processing', 'Step 4: Extract && Refresh']` (the doubled `&&` shows as one "&" on screen), current tab 0, status bar "Ready", application name/version "GST File Organizer" / "4.0", exit code 0 after the window was closed. On closing, `closeEvent` (gui/pyqt_main_window.py:601-608) wrote `settings.json` (the 11 default keys) and `session_state.json` = `{"active_tab": 0, "window_size": [1400, 900]}`.

**Case C - a hand-edited settings file in which `pq_wait_time` was typed as text: `{"pq_wait_time": "10"}`.**

1. models/cache_manager.py:59-61 loads it; `AppSettings.from_dict` (models/app_state.py:41-46) keeps any key that is a field name and checks no types, so `pq_wait_time` becomes the text "10".
2. gui/pyqt_main_window.py:185 calls `spin_wait_time.setValue("10")`; Qt requires a whole number and raises `TypeError`.
3. The error leaves the window constructor and is caught by main.py:56 → "Critical error: setValue(self, val: int): argument 1 has unexpected type 'str'" printed and logged with the full traceback → exit code 1.

Actual output (RAN): exit 1, the window never appears. The only remedy is to repair or delete `~/.gst_organizer/settings.json` by hand.

**Case D - the per-user file `session_state.json` holds `"recent_source_folders": "C:/old"` (text instead of a list, e.g. after a hand edit), the window opens, and the user ticks View > Dark Mode.**

1. The menu action calls `_toggle_dark_mode` (gui/pyqt_main_window.py:594-597), which calls `_save_settings` (210-214), which calls `CacheManager.save_settings` (models/cache_manager.py:41-55).
2. Line 47 reads `recent = "C:/old"`; line 48: the source folder is not inside that text; line 49 `recent.insert(...)` → `AttributeError: 'str' object has no attribute 'insert'`.
3. Nothing catches it: B installs no `sys.excepthook` (a hook that receives errors nothing else caught; grep over all .py files: 0 hits). PyQt5 5.15 then **aborts the whole program**.

Actual output (RAN): the process ended with signal 6 (abort; a Linux shell reports it as exit code 134); the traceback appeared only on the error stream of the console; nothing was added to `gst_organizer.log`; no message box; the settings were not saved.

**What the trace shows against what the feature should do.** Bugs (all listed for Section 7):

1. The window cannot start without pywin32, while the dependency file comments pywin32 out and the error message recommends the wrong command (Case A, RAN). The window itself does not need Excel to scan or organise; the hard dependency comes only from the unguarded import in `report_processor.py:11-15`, while `core/excel_handler.py:19-28` and `utils/helpers.py:473-483` guard the same imports (an inconsistent pattern).
2. Start-up errors are only `print()`ed (main.py:51-53, 58). Started without a console (double-click with `pythonw`, as the author's Standard 13 requires), the user sees nothing at all. (READ; HIGH.)
3. One wrong value type in the settings file stops the program from starting (Case C, RAN).
4. Any error inside a button or menu handler kills the window without saving and without a message (Case D, RAN).
5. No check of the Python version, of Windows or of Excel, and no offer to install missing packages (READ: nothing in main.py:1-63 does this).

**Comparison with the author's Standard 13 (facts only).**

| Standard 13 (rebuild_pipeline_FINAL.md) | Project B |
|---|---|
| Users double-click a shortcut to `Run [App Name].bat`, which starts `pythonw.exe -B -s main.py` (lines 951-975) | No launcher; README tells users to type `python main.py` (README.md:110-113) |
| `main.py` installs `sys.excepthook` and `threading.excepthook` that write an error file and show a message box (lines 984-1029) | Neither hook; errors are printed (main.py:49-59) and an error in a handler aborts (Case D) |
| Logs and settings in `%LOCALAPPDATA%\[App Name]\` (lines 1035-1040) | Log in the start folder (main.py:20); settings in `~/.gst_organizer` (models/cache_manager.py:17) |
| Exact package versions in requirements.txt (lines 937-940) | Minimum versions only (`>=`), pywin32 commented out (requirements.txt:5-14) |
| Desktop windows with PySide6 (Standard 12 table, line 833) | PyQt5 (requirements.txt:5) |

**Older documents against the code.** README.md:52 says "Windows 10/11 (best), macOS, Linux": the program cannot open on macOS or Linux (Case A). README.md:104-108 shows the requirements file as `openpyxl>=3.0.0` and `pywin32>=305; sys_platform == 'win32'`: the real file has neither line. README.md:275 says the error to expect is "No module named 'win32com'": B reports `pythoncom` first. README.md:110-113 (`python main.py`) is still the way to start it, but only with pywin32 installed.

**Origin.** `main.py` was created in edac6a3 (shared - before the split on 30-10-2025) for the Tkinter version; it was rewritten for PyQt5 in c258047 (B only (after the split), 07-02-2026). The same commit made the unified window import the Step 4 extractor, which is what made pywin32 mandatory for start-up (at 47a1492 the organiser window did not import `report_processor`). `gui/__init__.py` was left importing a deleted module by a62d3cc and repaired in c21f788 (both B only (after the split)). Rules: B-R124, B-R125, B-R126.

---

### B-F2: One-time import of old v3 settings (migration)

**What it does.** The first time the new window opens on a PC where the old Tkinter version (v3.0) was used, it copies the old saved settings (folders, templates, processing mode, the include-name tick, the maximum length, dark mode) into the new settings file and leaves a marker file so that the import never runs again.

**Where it lives.** models/cache_manager.py:93-131 (`migrate_from_old_cache`); called once at gui/pyqt_main_window.py:47, after the run history is loaded (43) and before the settings are loaded (50). The old file was written by the v3 program: keys at gui/main_window.py:330-339@47a1492 and the recent-folders list at gui/utils/cache_manager.py:29-46@47a1492 (key `recent_folders`). Marker: `~/.gst_organizer/.migrated` (models/cache_manager.py:99, 127).

**Input → Output.** `~/.gst_organizer_cache.json` (v3) → `~/.gst_organizer/settings.json`, `~/.gst_organizer/session_state.json` and the marker `~/.gst_organizer/.migrated`.

**User interaction.** None; automatic when the window opens, if the old file exists and the marker does not.

#### B-F2 concrete trace (RAN: `scripts/t_f2_migration.py`, output `work/f2_out.txt`)

Input: an old v3 file written with exactly the keys the v3 program wrote - source folder `D:/Clients/FY-Source`, ITC and Sales template paths, target `D:/Clients/Output`, `processing_mode` "rerun", `include_client_name` true, `dark_mode` false, `client_name_max_length` **0**, and `recent_folders` = [`D:/Clients/FY-Source`, `D:/Clients/Older-Source`].

1. `CacheManager()` (models/cache_manager.py:31-37) creates the empty folder `~/.gst_organizer`.
2. models/cache_manager.py:98-99 - paths of the old file and of the marker. Line 101: the old file exists and the marker does not, so the import continues.
3. Lines 105-106 read the old JSON.
4. Lines 109-118 build the new settings: `source_folder`, `itc_template`, `sales_template`, `target_folder` copied; `processing_mode` "rerun"; `include_client_name` → `include_client_name_in_folders` = true; `client_name_max_length` = 0 (the default 35 at line 116 applies only when the key is missing); `dark_mode` false. The three Step 4 settings get their defaults.
5. Line 119 `save_settings` (lines 41-55): the session file is empty, so the recent list is `[]`; the source folder is not in it, so it is inserted and `session_state.json` is written (lines 48-52); then `settings.json` is written (line 54).
6. Line 122 looks for `recent_source_folders` in the old file. The v3 program wrote that list under `recent_folders`, so the result is `[]` and lines 123-124 are skipped.
7. Line 127 creates the marker; line 128 logs "Migrated settings from old cache file".

Actual output (RAN): files `['.migrated', 'session_state.json', 'settings.json']`; `settings.json` holds the 11 keys with `"processing_mode": "rerun"`, `"include_client_name_in_folders": true`, `"client_name_max_length": 0`; `session_state.json` = `{"recent_source_folders": ["D:/Clients/FY-Source"]}` - the second old folder is gone.

8. Second start with a changed old file: skipped at line 101 because the marker exists (RAN: the source folder stayed `D:/Clients/FY-Source`).
9. The window built from these settings (RAN): `window.settings.client_name_max_length` = 0, but the "Max length" box shows **15** (Qt keeps a spin box inside its range 15-100, gui/views/setup_view.py:375, 647), the tick is on and "Re-Run" is selected. When processing starts, `_collect_settings` reads the box (gui/pyqt_main_window.py:203), so processing receives **15**.
10. A damaged old file (cut-off JSON) on a fresh profile, started twice: each start logs the WARNING "Failed to migrate old cache: Expecting property name enclosed in double quotes: line 1 column 27 (char 26)"; no settings and no marker are written, so the import is retried at every start (RAN).

**What the trace shows against what the feature should do.** Bugs (for Section 7):

1. The old recent-folders list is never carried over: the writer used `recent_folders`, the reader asks for `recent_source_folders` (a write/read key mismatch; RAN). Low impact today because B never shows the recent list (B-F27).
2. **A saved maximum length of 0 becomes an effective limit of 15 characters**, not 35: the migration copies 0 (line 116), the spin box shows 15, and processing uses the box value. The v3 program replaced 0 by 35 when loading (gui/main_window.py:352-354@47a1492). So a user whose old file held 0 gets client folders cut to 15 characters without being told. (RAN; this adds to the registry, which records only that 0 is copied as 0. Likelihood that a real v3 file holds 0: LOW.)
3. A damaged old file is retried at every start with only a log warning (RAN). If writing `settings.json` fails, the marker is still created, because `_save_json` swallows write errors (models/cache_manager.py:141-142) before line 127 runs (READ; LOW).
4. The old file is left in place (READ).

**Older documents against the code.** README.md:27 ("Session Memory: Remembers your last used folders and settings") - the folders are remembered as settings, but the list of recent folders is neither migrated nor shown.

**Origin.** c258047 (B only (after the split), 07-02-2026). Rule: B-R118.

---

### B-F3: Choose source folder, target folder and templates (Setup)

**What it does.** On "Step 1: Setup" the user types or picks four things: the folder holding the client files (source), the ITC template workbook, the Sales template workbook, and the folder where the organised output goes (target). The values are remembered for the next session (see B-F27).

**Where it lives.**

| Part | file:line | What the code does |
|---|---|---|
| Source box + Browse | gui/views/setup_view.py:192-214 | Card "SOURCE FOLDER"; text box `source_folder` (200-202); "Browse" button, objectName "success" (204-208). |
| Template boxes + Browse | gui/views/setup_view.py:218-261 | Card "EXCEL TEMPLATES" with the caption "Select your Excel template files (.xlsx or .xlsm)" (221); boxes `itc_template` (231), `sales_template` (249); two "Browse" buttons, objectName "primary" (234-238, 252-256). |
| Target box + BROWSE | gui/views/setup_view.py:265-330 | Card "TARGET FOLDER - IMPORTANT!", box `target_folder` (297), red "BROWSE" button with its own style (308-324). |
| Pickers | gui/views/setup_view.py:605-627 | Folder pickers for source and target (606, 625); file pickers with the filter "Excel Files (\*.xlsx \*.xlsm)" (612, 619). None is given a starting folder. |
| Fill at start | gui/views/setup_view.py:635-647; gui/pyqt_main_window.py:174-187 | Writes the saved values into the boxes. |
| Read for saving and processing | gui/pyqt_main_window.py:194-214 | `_collect_settings` reads the boxes as they are; `_save_settings` stores them (models/cache_manager.py:41-55). |
| Checks before a scan | gui/pyqt_main_window.py:218-225 | Source must be non-empty ("Please select a source folder first.") and an existing folder ("Folder not found: <path>"). |
| Checks before processing | gui/pyqt_main_window.py:352-361 | ITC and Sales template paths must be non-empty and exist ("Please select a valid ITC template." / "... Sales template."); the target must be non-empty ("Please select a target folder."); the target need not exist (it is created later, core/file_organizer.py:61-65). |

**Input → Output.** Typed or picked paths → the settings fields `source_folder`, `itc_template`, `sales_template`, `target_folder`, used by the scan (B-F6), processing (B-F12 to B-F15) and Step 4 (B-F18, which reads the last **saved** target, gui/pyqt_main_window.py:421).

**User interaction.** Typing or pasting into the four boxes, or clicking Browse / BROWSE.

#### B-F3 concrete trace (RAN: `scripts/t_f3_setup.py`, output `work/f3_out.txt`)

Input: <Client A>'s source folder exists with one GSTR-3B file. The user (1) pastes the folder path the way Windows Explorer's "Copy as path" gives it, **with double quotes**; (2) types an ITC template path that is in fact a **folder** (`.../NotATemplate`), picks a real Sales template, and types a target folder that does not exist yet.

1. SCAN FILES → `_on_scan_requested` (gui/pyqt_main_window.py:218). Line 219 reads `"…/Source"` including the quotes; line 220 - not empty; line 223 `Path('"…/Source"').is_dir()` is false → line 224 shows the warning **"Folder not found: "/root/…/work/f3/Source""** and returns. No scan starts (RAN: the scan controller's thread stayed empty).
2. The same path with one trailing space → "Folder not found: …/Source " on this Linux machine (RAN). On Windows the file system normally ignores a trailing space at the end of a path, so there it would probably pass (MEDIUM; not testable here).
3. The path without quotes → the scan runs (B-F6); before it starts, line 227 saves all settings, so `settings.json` now holds the four boxes as they are at that moment.
4. START on Step 2 → `_on_start_processing` (345): one client ticked (346-350); line 352 collects the boxes; line 353 `Path(".../NotATemplate").exists()` is **true for a folder**, so the check passes; line 356 the Sales template exists; line 359 the target is not empty; lines 363-371 start processing.
5. In the worker, `_create_report` (gui/controllers/processing_controller.py:204-239): line 210 the path is not empty; line 215 it exists; line 220 output `ITC_Report_ClientA_Traders_MH_<DDMMYY_HHMM>.xlsx`; line 225 calls `create_report_from_template`, whose first step copies the "template" to the report path (core/excel_handler.py:228 on this Linux run; core/excel_handler.py:124 on Windows). Copying a folder raises an error; the openpyxl branch catches it (260-262) and returns False; on Windows the Excel branch catches it at 191-193 and also returns False (READ). Line 233 then logs "❌ Failed to create ITC report" - the reason is only in the program log.
6. The Sales report is made normally; the client is counted as successful (processing_controller.py:135).

Actual output (RAN): message box **"Successfully processed 1/1 clients"**; summary workbook: Successful 1, ITC Reports Created 0, Sales Reports Created 1, Report Creation Errors 1; "Client Status" says "Success"; the "Errors" sheet says "No errors encountered".

7. Pixel check of the four Browse buttons as drawn (RAN, `QWidget.grab()` of each button): the three "Browse" buttons for Source, ITC and Sales are **0 % non-white** (90 x 36 pixels: invisible white text on a white background); the Target "BROWSE" is 96 % non-white. Cause (READ): the card's content area has the style `QWidget { background: white; }` (gui/widgets/colored_section.py:62), which overrides the theme's coloured background for the buttons inside it, while the theme still gives these buttons white text and no border (gui/theme/theme_manager.py:92-97, 112-117). The buttons still work if clicked blind.

**What the trace shows against what the feature should do.** Bugs (for Section 7):

1. No clean-up of what is pasted: surrounding quotes (and on Linux a trailing space) make a correct folder fail with "Folder not found" (RAN).
2. A template is only checked to **exist**, not to be a file or an Excel workbook: a folder passes, and every client then gets "Failed to create ITC report" while the run still reports every client as successful (RAN).
3. Three of the four Browse buttons are invisible (RAN).
4. The pickers do not open at the current value (READ: no starting folder at 606, 612, 619, 625; Qt then opens at the program's current folder).
5. `.xlsm` (macro) templates are offered (612, 619), but every report is saved under a `.xlsx` name (processing_controller.py:220); Excel normally refuses to open a macro workbook with an `.xlsx` name (MEDIUM; not testable without Excel).
6. A source box changed after the last scan is not noticed: START processes the old scan's files (READ: gui/pyqt_main_window.py:366-371 uses `self.scan_result`, never the box).
7. Settings are saved when a scan starts, after processing, on File > Save Settings, on the dark-mode switch and on closing - not when START is pressed (READ; details in B-F27).

**Older documents against the code.** README.md:120-121 says templates may be ".xlsx or .xltx"; the screen says and offers ".xlsx or .xlsm" (setup_view.py:221, 612, 619), and the unused constant `TEMPLATE_EXTENSIONS = {'.xlsx', '.xltx', '.xltm'}` (utils/constants.py:180) gives a third answer. README.md:138-141 tells the user to "Click Browse" - three of the four are invisible.

**Origin.** Screen created in c258047 and restyled into coloured cards in 7e09e5c (both B only (after the split)); the invisible buttons appeared with the cards in 7e09e5c (VW notes, READ of the diff). Rules: B-R48, B-R133, B-R116.

---

### B-F4: Include client name in category-folder names (global option) and client-folder length limit

**What it does.** A tick box, "Include client name in brackets for Level 4 folders", adds " (<client name>)" to the three category folders inside each version folder ("GSTR-3B Exports", "Other ITC related files", "Sales related files"). A "Max length" box, greyed out unless the tick is on, limits the length of the client folder name ("Level 2").

**Where it lives.** gui/views/setup_view.py:367-388 (tick box 369; spin box 374-377: range 15-100, default 35, starts disabled; 382 the tick enables/disables the box; note text 385); gui/pyqt_main_window.py:202-203 (both values read for processing); gui/controllers/processing_controller.py:46-52 (passed to the organiser; `client_name_max_length` is **always** passed, line 51); core/file_organizer.py:42 (0 or less becomes 35), 79-83 (client folder name), 86 (cleaned name for the brackets), 113-135 (bracket decision); utils/helpers.py:513-537 (`create_client_state_key`) and 66-112 (`sanitize_filename`).

**Input → Output.** Tick (yes/no) + a number → category-folder names with or without the bracket; the client folder name cut to the number.

**User interaction.** The tick box and spin box on Step 1; their values are read at the moment processing starts (gui/pyqt_main_window.py:352).

#### B-F4 concrete trace (RAN: `scripts/t_f4_f5_f12.py`, output `work/f12_out.txt`)

Input: <Client E> ("ClientA Traders and Distributors", Maharashtra) with one GSTR-3B file. On screen the user once ticked the box, set "Max length" to 15, and then unticked the box again (a natural sequence).

1. Window (RAN): at start the tick is off and the box shows 35, disabled. After tick → 15 → untick, the box is disabled again, but `_collect_settings` (gui/pyqt_main_window.py:202-203) returns `include_client_name_in_folders = False` and `client_name_max_length = 15`.
2. The organiser is built with those values (processing_controller.py:46-52); core/file_organizer.py:42 keeps 15 (it is above 0).
3. `create_client_structure` (core/file_organizer.py:67-144): line 78 state "Maharashtra" → code "MH" [GST RULE] ⚠ (state-code table, utils/constants.py:234-275); lines 79-83 → `create_client_state_key` (utils/helpers.py:513-537): lines 517-518 find no "Private"/"Limited"; line 524 key "ClientA Traders and Distributors-MH" (35 characters); line 527: 35 > 15; line 529: room for the client part = 15 − 2 − 1 = 12; line 530: 12 > 5; line 531 cuts the client to "ClientA Trad"; line 532 key **"ClientA Trad-MH"**.
4. Line 86 cleans the client name for the brackets: "ClientA_Traders_and_Distributors". Lines 113-135, for each of the three category folders: line 121 the global tick is off; line 126 the per-client dictionary is empty (B-F5), so `include_name` is False; lines 129-130 remove " (ClientA_Traders_and_Distributors)" from the name.

Actual output (RAN): client folder **`ClientA Trad-MH`**; category folders `GSTR-3B Exports`, `Other ITC related files`, `Sales related files`.

5. Same client with the tick **on** and the box at 35 (RAN): client folder `ClientA Traders and Distributors-MH` (exactly 35, not cut); category folders **`GSTR-3B Exports (ClientA_Traders_and_Distributors)`**, `Other ITC related files (ClientA_Traders_and_Distributors)`, `Sales related files (ClientA_Traders_and_Distributors)` - the bracketed name is the full cleaned name and is not limited by "Max length".

**What the trace shows against what the feature should do.** Bugs (for Section 7):

1. The greyed-out "Max length" box still decides the client folder name (RAN): the screen suggests it has no effect when the tick is off.
2. "Max length" does not limit the bracketed name it sits next to; it limits a different folder (the client folder) (RAN).
3. The note under the option, "Names >10 chars will show warning. This overrides individual client settings if checked." (setup_view.py:385), describes a long-name warning and per-client settings that do not exist in B (READ: grep finds no such warning; the per-client setting is dead, B-F5).
4. Cutting can leave a space before the hyphen, and two different clients can be folded into one client folder by the cut or by the "Private"→"Pvt" / "Limited"→"Ltd" abbreviation (RAN by the live-trace reader, LT §2.4 and §6.6; see B-F12).

**Older documents against the code.** README.md:185-187 shows the category folders always as "GSTR-3B Exports (ClientName)" etc.; by default the code removes the bracket (B-R35). README.md:183 shows "ClientName-StateCode/" without the abbreviation and cut.

**Origin.** Folder logic edac6a3 (shared - before the split on 30-10-2025). In v3 the length limit was a separate, always-active setting "Client folder name max length:" (gui/tabs/setup_tab.py:248-279@47a1492, VW notes). c258047 (B only (after the split)) moved it next to the tick and tied its enabled state to it (range 5-100); 7e09e5c (B only (after the split)) changed the range to 15-100 and added the v3 note text. Rules: B-R32 ([GST RULE] for the state-code part), B-R35.

---

### B-F5: Per-client option to include the client name in folder names

**What it does.** Meant to let each client override the global "include client name" option. In B only the internal plumbing exists; there is no control that can switch it on.

**Where it lives.** core/file_organizer.py:117-126 (when the global tick is off, looks up `client_folder_settings["<client>-<state code>"]`, default False); models/app_state.py:83 (the dictionary, always created empty); gui/pyqt_main_window.py:370 (passed to processing); gui/controllers/processing_controller.py:26-32, 50 (passed on to the organiser).

**Input → Output.** None in practice (the dictionary is always empty).

**User interaction.** None - no screen element writes the dictionary.

#### B-F5 concrete trace (RAN: `scripts/t_f4_f5_f12.py`, output `work/f12_out.txt`; READ: grep)

1. READ (grep of all .py files for `client_folder_settings`): the only lines are models/app_state.py:83 (the empty default), gui/pyqt_main_window.py:370 (read), gui/controllers/processing_controller.py:27, 32, 50, 257, 266 (passed on) and core/file_organizer.py:29, 41, 126 (read). **No line writes into it.**
2. RAN: `ProcessingState().client_folder_settings` = `{}`; the live window's `processing_state.client_folder_settings` = `{}`.
3. RAN, to show the branch itself works: with the dictionary filled **by hand in the test** as `{'ClientA Traders-MH': True}` and the global tick off, <Client A>'s category folders became `GSTR-3B Exports (ClientA_Traders)`, `Other ITC related files (ClientA_Traders)`, `Sales related files (ClientA_Traders)`.

Actual output in the program: always the global tick's choice; the per-client branch is dead.

**What the trace shows.** A lost feature (for Section 7). In v3 (Tkinter) the user could tick "Folder Name" per client in the client tree (`toggle_folder_name_setting`, gui/handlers/client_handler.py:175-206@47a1492): it asked "Long Client Name Warning" for names over 10 characters, and unticking one client switched the global option off; `update_global_folder_setting` (208-219@47a1492) set every client when the global tick was turned on. The v3 key was `"<client>-<state as shown in the tree>"` (line 178@47a1492), while B looks up `"<client>-<state code>"` (core/file_organizer.py:118) - the two would match only if the tree showed the code (MEDIUM, READ). The control was not carried into the PyQt5 window (c258047) and the Tkinter handler was deleted in a62d3cc (both B only (after the split)).

**Origin.** Logic edac6a3 (shared - before the split on 30-10-2025); control lost in c258047 / a62d3cc. Rule: B-R36 (DEAD).

