## SECTION 3: FEATURE INVENTORY

This section lists every feature of project B (AnnualStatement_oldv2, branch main, snapshot b0e7d0e of 09-02-2026, the PyQt5 window that calls itself "GST File Organizer v4.0") and documents, feature by feature, what it does, where it lives, what goes in and what comes out, how the user starts it, and one concrete worked example ("trace") followed through the code to the real result. This file is the **first part** of the section: it gives the table of all 34 features and the full entries for **B-F1 to B-F17** (start-up, Step 1 Setup, Step 2 Validation, Step 3 Processing). B-F18 to B-F34 (Step 4 Extract & Refresh and the cross-cutting features) follow in the second part.

**How the traces were made (read this once).**

- Every feature ID and rule ID (B-F…, B-R…) is the one in the project-B registry (`/root/merge-scratch/audit/B_registry.md`). Locations are `file:line` in project B at b0e7d0e, repository-relative (for example `core/file_parser.py:243`). Older code is written `path:line@commit`.
- **RAN** = the code was executed on made-up inputs and the result is quoted from the output. **READ** = the statement comes from reading the code line by line. Every inference carries a confidence word: **HIGH**, **MEDIUM** or **LOW**.
- All runs used a scratch copy of the code (`/root/merge-scratch/run/B-S3a/proj`, byte-for-byte equal to the repository; checked with `diff -rq` after the runs: no difference, no `__pycache__` written) and `/root/merge-scratch/venv-B/bin/python -B` (Python 3.14 with PyQt5 5.15, openpyxl 3.1.5). The program's per-user folder was redirected to scratch folders (`/root/merge-scratch/run/B-S3a/home/...`) so nothing touched a real home folder. Scripts are in `/root/merge-scratch/run/B-S3a/scripts/`, their printed output in `/root/merge-scratch/run/B-S3a/work/*_out.txt`.
- **Independent re-run (verification).** Every script named in this part was run again by the verifier on a fresh copy of the code (`/root/merge-scratch/run/VERIFY-B-S3a/proj`, `diff -rq` against the repository: identical, no `__pycache__`), output in `/root/merge-scratch/run/VERIFY-B-S3a/work/*_out.txt`. All quoted results came out the same; only clock-based stamps differ (for example "231026 1216" instead of "211026 0640" in folder and file names, and the run times), because those runs took place on another day.
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
2. Start-up errors are only printed on the console (main.py:51-53, 58) and written to `gst_organizer.log` in the start folder (main.py:50, 57); no message box is shown. Started without a console (double-click with `pythonw`, as the author's Standard 13 requires), the user sees nothing at all on screen. (READ; HIGH.) [Corrected during verification - originally said the errors were "only `print()`ed"; main.py:50 and 57 also log them.]
3. One wrong value type in the settings file stops the program from starting (Case C, RAN).
4. Any error inside a button or menu handler kills the window without saving and without a message (Case D, RAN).
5. No check of the Python version, of Windows or of Excel, and no offer to install missing packages (READ: nothing in main.py:1-63 does this).

**Comparison with the author's Standard 13 (facts only).**

| Standard 13 (rebuild_pipeline_FINAL.md) | Project B |
|---|---|
| Users double-click a shortcut to `Run [App Name].bat`, which starts `pythonw.exe -B -s main.py` (lines 951-975) | No launcher; README tells users to type `python main.py` (README.md:110-113) |
| `main.py` installs `sys.excepthook` and `threading.excepthook` that write an error file and show a message box (lines 984-1029) | Neither hook; start-up errors are printed and logged to `gst_organizer.log` (main.py:49-59), no message box, and an error in a handler aborts (Case D) |
| Logs and settings in `%LOCALAPPDATA%\[App Name]\` (lines 1035-1040) | Log in the start folder (main.py:20); settings in `~/.gst_organizer` (models/cache_manager.py:17) |
| Find the app's own files from `APP_DIR` (the program's own folder), never from the current folder (line 1034) | The import path is taken from the program's own folder (main.py:28-29), but the log file name has no folder, so it follows the current folder (main.py:20) |
| Exact package versions in requirements.txt (lines 937-940) | Minimum versions only (`>=`), pywin32 commented out (requirements.txt:5-14) |
| Desktop windows with PySide6 (Standard 12 table, line 833) | PyQt5 (requirements.txt:5) |

**Older documents against the code.** README.md:52 says "Windows 10/11 (best), macOS, Linux": the program cannot open on macOS or Linux (Case A). README.md:104-108 shows the requirements file as `openpyxl>=3.0.0` and `pywin32>=305; sys_platform == 'win32'`: the real file has neither line. README.md:275 says the error to expect is "No module named 'win32com'": B reports `pythoncom` first. README.md:110-113 (`python main.py`) is still the way to start it, but only with pywin32 installed.

**Origin.** `main.py` was created in edac6a3 (shared - before the split on 30-10-2025) for the Tkinter version and lost its first "shebang" line (a Unix "run me with Python" line) in 5ea77c2 (shared - before the split on 30-10-2025); it was rewritten for PyQt5 in c258047 (B only (after the split), 07-02-2026). The same commit made the unified window import the Step 4 extractor, which is what made pywin32 mandatory for start-up (at 47a1492 the organiser window did not import `report_processor`). `gui/__init__.py` was left importing a deleted module by a62d3cc and repaired in c21f788 (both B only (after the split)). Rules: B-R124, B-R125, B-R126.

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
| Source box + Browse | gui/views/setup_view.py:192-214 | Card "SOURCE FOLDER"; text box `source_folder` (200-202); "Browse" button, objectName "success" (an internal name the theme uses to pick the button's colours) (204-208). |
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
5. In the worker, `_create_report` (gui/controllers/processing_controller.py:204-239): line 210 the path is not empty; line 215 it exists; line 220 output `ITC_Report_ClientA_Traders_MH_<DDMMYY_HHMM>.xlsx`; line 225 calls `create_report_from_template`; its existence check (core/excel_handler.py:78) passes because the folder exists, and the chosen branch then copies the "template" to the report path - on this Linux run as the openpyxl branch's first step (core/excel_handler.py:228); on Windows the Excel branch first attaches to or starts an Excel and hides it (110-120) and only then copies (124). Copying a folder raises an error; the openpyxl branch catches it (260-262) and returns False; on Windows the Excel branch catches it at 191-193 and also returns False, after Excel has already been started (READ). Line 233 then logs "❌ Failed to create ITC report" - the reason is only in the program log.
6. The Sales report is made normally; the client is counted as successful (processing_controller.py:135).

Actual output (RAN): message box **"Successfully processed 1/1 clients"**; summary workbook: Successful 1, ITC Reports Created 0, Sales Reports Created 1, Report Creation Errors 1; "Client Status" says "Success"; the "Errors" sheet says "No errors encountered".

7. Pixel check of the four Browse buttons as drawn (RAN, `QWidget.grab()` of each button): the three "Browse" buttons for Source, ITC and Sales are **0 % non-white** (90 x 36 pixels: invisible white text on a white background); the Target "BROWSE" is 96 % non-white. Cause (READ): the card's content area has the style `QWidget { background: white; }` (gui/widgets/colored_section.py:62), which overrides the theme's coloured background for the buttons inside it, while the theme still gives these buttons white text and no border (gui/theme/theme_manager.py:92-97, 112-117). The buttons still work if clicked blind.

**What the trace shows against what the feature should do.** Bugs (for Section 7):

1. No clean-up of what is pasted: surrounding quotes (and on Linux a trailing space) make a correct folder fail with "Folder not found" (RAN).
2. A template is only checked to **exist**, not to be a file or an Excel workbook: a folder passes, and every client then gets "Failed to create ITC report" while the run still reports every client as successful (RAN).
3. Three of the four Browse buttons are invisible (RAN).
4. The pickers do not open at the current value (READ: no starting folder at 606, 612, 619, 625). With no starting folder, Qt opens the folder last visited in a file dialog during this session, or the program's current folder for the first dialog (MEDIUM - Qt's own behaviour, not tested here; on Windows the native dialog may also remember a folder of its own; how to check: on the Windows PC, type a path into the box, click Browse and see where the dialog opens). [Corrected during verification - originally said Qt "then opens at the program's current folder", which holds only for the first dialog of a session.]
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

---

### B-F6: Scan source folder and recognise client files

**What it does.** When the user clicks SCAN FILES (or RE-SCAN), the program lists the Excel files lying directly in the source folder, rejects files that are too small or do not start like an Excel file, reads each remaining file *name* to decide which kind of GST file it is (GSTR-2B Reco, IMS Reco, GSTR-3B export, Sales, Sales Reco, Annual Report), which client it belongs to and which state, and groups the files per client with a completeness status. The work runs in a background **thread** so the window stays usable. The source files themselves are only read, never changed.

**Where it lives.**

| Step | file:line | What the code does |
|---|---|---|
| Buttons | gui/views/setup_view.py:421-437 (SCAN FILES → signal `scan_requested`), 439-456 (RE-SCAN → `rescan_requested`; disabled until the first scan) | Two buttons, one job. |
| Wiring | gui/pyqt_main_window.py:124-125 | Both signals go to the same handler `_on_scan_requested`. |
| Checks and start | gui/pyqt_main_window.py:218-232 | Source box must be non-empty (220-222) and an existing folder (223-225); saves all settings (227); collapses the two help cards (228); clears the scan log (229); writes "Scanning: <folder>" (230); disables SCAN FILES (231); starts the job (232). |
| Background job | gui/controllers/scan_controller.py:57-80 (refuses silently if a scan is running, 59-61), worker 24-41 | Runs `FileParser().scan_folder(...)` and `get_statistics()` in a QThread (Qt's background thread) and sends the result back as one `ScanResult` (models/app_state.py:49-55). |
| Folder listing | utils/helpers.py:249-266 `find_excel_files` | Wildcard searches (**glob**) `*.xlsx`, `*.xls`, `*.xlsm` in the top level only (255-257); drops names starting `~$` (Excel lock files, 260); sorts by full path (262); any error → empty list (264-266). |
| Per-file checks | core/file_parser.py:171-256 `_process_file` | File size and dates (175-179; failure → file silently skipped); under 1024 bytes → problem list (184-191); first 8 bytes must start `PK\x03\x04` or `\xd0\xcf\x11\xe0` (194-215); name recognition (229); stored by file name (240); to the client or to the problem list (243-252). |
| Name recognition [GST RULE] | core/file_parser.py:56-124 `parse_filename`; six patterns utils/constants.py:13-62 ⚠ | Tries the six **regular expressions** (text patterns) in a fixed order, ignoring letter case; the first that matches decides type, client, state and the period part. |
| Grouping [GST RULE] | core/file_parser.py:258-278 `_add_to_client_data`; state codes utils/helpers.py:485-511, table utils/constants.py:234-275 ⚠ | Client key = "<client text>-<state code>". |
| Completeness [GST RULE] | core/file_parser.py:280-314; expected types utils/constants.py:68-75 ⚠ | Missing types, duplicate types (GSTR-3B exempt), percentage, status. |
| Statistics | core/file_parser.py:336-364 | Counts shown in B-F7. |

**Input → Output.** The source folder → a `ScanResult` holding `scanned_files` (one record per accepted file, keyed by file name), `client_data` (one entry per client key), `variations` (the code's word for problem files) and `statistics`.

**User interaction.** Type or pick the source folder (B-F3), click SCAN FILES; RE-SCAN repeats the same job. There is no Cancel for a scan.

#### B-F6 concrete trace (RAN: `scripts/t_f6_scan.py`, output `work/f6_out.txt`; the same folder scanned through the real window in `scripts/t_window_flow.py`, output `work/flow_out.txt`)

Input: the made-up source folder of the introduction (built by `scripts/mk_source.py`): 30 entries - 27 genuine .xlsx-type files of 4,866-4,870 bytes each (about 4.8 KB in the program's own 1,024-byte units) for <Client A> to <Client D> (one of them with the extension `.xlsm`, one an unrecognised "Purchase Register"), the 300-byte `SalesReco-<Client C>-Tamil Nadu-FY2425.xlsx`, a lock file `~$GSTR3B-<Client A>-Maharashtra-Apr.xlsx` and a sub-folder `old/` holding one more GSTR-3B file.

1. SCAN FILES → gui/views/setup_view.py:437 → gui/pyqt_main_window.py:218. The box holds an existing folder (220, 223), so line 227 saves the settings (RAN: the first status-bar message is "Settings saved"), 231 disables SCAN FILES (RAN: `enabled right after click: False`) and 232 calls `scan_controller.scan_files`.
2. gui/controllers/scan_controller.py:59 - no scan running; 63 emits `scan_started`; 65-80 create the thread and start it. In the thread, lines 26-30 call `scan_folder` with a progress function that forwards every message as a Qt **signal**.
3. core/file_parser.py:138 empties the parser; 142-145 the folder exists and is a folder; 147 sends "Finding Excel files...".
4. utils/helpers.py:255-257 run the three globs **in the top folder only**, so the file in `old/` is never seen; 260 removes the lock file; 262 sorts. Result (RAN): **28 paths**, the `.xlsm` and the 300-byte file included.
5. core/file_parser.py:155 sends "Found 28 files, processing..."; the loop at 157-159 sends "[1/28] (1) GSTR3B-…" … "[28/28] SalesReco-…" and calls `_process_file` for each (RAN: 30 progress messages in all).
6. Inside `_process_file`, file by file (RAN):

| File (made-up) | Deciding line(s) | Outcome |
|---|---|---|
| `(1) GSTR3B-<Client A>-Maharashtra-Jun.xlsx` | size and signature pass (184, 199); 229 → pattern "GSTR3B" (utils/constants.py:31), the only pattern that allows a leading "(n) " | GSTR-3B Export of client "ClientA Traders", state "Maharashtra", month "Jun" |
| `GSTR3B-<Client A>-Maharashtra-Apr.xlsx`, `…-May.xlsx` | same | GSTR-3B Export |
| `GSTR-2B-Reco-<Client A>-Maharashtra-FY2425.xlsx` | pattern "GSTR-2B-Reco" (constants.py:15) | GSTR-2B Reco, period "FY2425" |
| `AnnualReport-<Client A>-Maharashtra-2024-25.xlsx` | pattern "AnnualReport" (constants.py:55); `(.+)` takes "2024-25" | Annual Report, year "2024-25" (kept as text, never checked - B-R23) |
| `GSTR3B-<Client A>-Maharashtra-Jul.xlsm` | every pattern ends `\.xlsx?$`, so nothing matches; 120-122, then 245-252 | problem file: "No matching pattern found", hint "GSTR3B-ClientName-State-Month.xlsx" |
| `<Client A> Purchase Register.xlsx` | nothing matches | problem file: "No matching pattern found", hint "Check file naming convention" |
| `SalesReco-<Client C>-Tamil Nadu-FY2425.xlsx` (300 bytes) | 184: 300 < 1024; returns at 191 before line 240 | problem file: "File too small to be valid Excel" (no hint); **not** a scanned file |
| `GSTR-2B-Reco-Client-D Exports-Kerala-FY2425.xlsx` | pattern "GSTR-2B-Reco": `([^-]+)` ("one or more characters that are not a hyphen") stops at the hyphen inside <Client D>'s name; the period part `(.+)` swallows the rest | client **"Client"**, state **"D Exports"**, period "Kerala-FY2425" - accepted with no warning on screen (only the program log notes the unknown state, step 7 below) |
| `GSTR3B-<Client C>-Tamil Nadu-Apr.xlsx` and `GSTR3B-<Client C>-Tamilnadu-May.xlsx` | both recognised | same client text, two different state texts |
| <Client B>'s 15 files (all six types, ten GSTR-3B months) | all recognised | - |

7. `_add_to_client_data` (core/file_parser.py:258-278) [GST RULE] ⚠: line 261 turns the state text into a code - "Maharashtra" → MH, "Gujarat" → GJ, "Tamil Nadu" → TN from the table (utils/constants.py:234-275); "Tamilnadu" is not in the table, so the one-word fallback takes its first three letters, **TAM** (utils/helpers.py:509), and "D Exports" (two words) becomes **DE** (utils/helpers.py:503-506); each fallback writes a WARNING to the program log (RAN: "Unknown state 'Tamilnadu', using fallback: TAM", "Unknown state 'D Exports', using fallback: DE"). Line 262 builds the key, so <Client C> becomes **two clients**, `ClientC Co-TN` and `ClientC Co-TAM`, and <Client D> becomes `Client-DE`. Lines 266-268 overwrite the client's state with that of the last file added.
8. `_analyze_client_completeness` (280-314) [GST RULE] ⚠, per client (RAN):
   - `ClientA Traders-MH`: types found GSTR-3B Export (3 files), Annual Report, GSTR-2B Reco → missing IMS Reco, Sales, Sales Reco (288); 3 of 6 types → 50.0 % (302-303); status "Missing 3 files" (309-310).
   - `ClientB Private Limited-GJ`: all six types; ten GSTR-3B files allowed (295-296) → 100.0 %, "Complete".
   - `ClientC Co-TN`: GSTR-3B 1, Sales 2 → remark "Multiple Sales (2 files)" (297); four types missing, and a missing type wins over a duplicate (306-311) → "Missing 4 files", 33.3 %.
   - `ClientC Co-TAM` and `Client-DE`: one type each → "Missing 5 files", 16.7 %.
9. `get_statistics` (336-364) → total_files 27, parsed 25, unparsed 2, clients 5, complete 1, incomplete 4, variations 3, total size "128.4 KB", parsing rate 92.59…, completion rate 20.0 (RAN). The worker sends the `ScanResult` to the window (scan_controller.py:32-38), which shows it (B-F7, B-F8).

Actual output (RAN, `work/f6_out.txt`):

```text
'Client-DE': 'Client' / 'D Exports' / {'GSTR-2B Reco': 1} / Missing 5 files
'ClientA Traders-MH': 'ClientA Traders' / 'Maharashtra' / {'GSTR-3B Export': 3, 'Annual Report': 1, 'GSTR-2B Reco': 1} / Missing 3 files
'ClientB Private Limited-GJ': … all six types, 'GSTR-3B Export': 10 / Complete
'ClientC Co-TAM': 'ClientC Co' / 'Tamilnadu' / {'GSTR-3B Export': 1} / Missing 5 files
'ClientC Co-TN': 'ClientC Co' / 'Tamil Nadu' / {'GSTR-3B Export': 1, 'Sales': 2} / extra ['Multiple Sales (2 files)'] / Missing 4 files
problem files: Purchase Register (No matching pattern found), …-Jul.xlsm (No matching pattern found), 300-byte SalesReco (File too small to be valid Excel)
```

10. Originals: SHA-256 fingerprint (a checksum that changes if even one byte of a file changes), size and modification time of all 30 source files were identical after a scan **and** a full processing run (RAN: `scripts/t_f12_source_unchanged.py`, `work/f12src_out.txt`: "identical after scan + processing: True").

**What the trace shows against what the feature should do.** The feature "recognises client files", but the trace shows several silent misrecognitions. Bugs (all listed for Section 7):

1. **A hyphen inside a client name silently splits the name** (RAN): <Client D>'s GSTR-2B Reco became client "Client", state "D Exports", key `Client-DE`; processing then makes a client folder `Client-DE` and reports `ITC_Report_Client_DE_…` (RAN, `work/flow_out.txt`). For GSTR3B, Sales and ImsReco names, whose last parts cannot take the extra hyphen (`([^-]+)` or `(\d{8})`), the same client name makes the file fail recognition instead (RAN: `scripts/t_f6_hyphen.py` → `work/f6h_out.txt` - with <Client D>'s name, GSTR-2B-Reco, SalesReco and AnnualReport names are accepted as client "Client" / state "D Exports" with the state name pushed into the period or year, while ImsReco, GSTR3B and Sales names give "No matching pattern found"). So one client's files are split between a mis-split client name ("Client" / "D Exports") and the problem list. [GST RULE] for the state-code part.
2. **One client becomes two when the state is spelled two ways** ("Tamil Nadu" / "Tamilnadu", RAN); "Delhi" and "DL" merge into one, and a different letter case of the client name makes a second client (RAN by the live-trace reader, LT §6.2). [GST RULE] (state-code logic) ⚠.
3. `.xlsm` source files are collected but can never be recognised; the hint tells the user to use a `.xlsx` name without saying the extension is the problem (RAN).
4. Files in sub-folders are ignored with no message (RAN: `old/` file never listed).
5. A file whose size cannot be read (for example a broken shortcut) is dropped from every list; only a log line remains (READ core/file_parser.py:175-179; RAN by LT §6.2).
6. Only the first 4 bytes are checked, so a fake file that merely starts with the right bytes is accepted (RAN by LT §6.2).
7. The name checks of `validate_client_name` are stored as `warnings` (core/file_parser.py:104-115) and the parser's `errors` list (46, 168, 256) - nothing reads either (READ: grep finds no reader).
8. From the second scan on, RE-SCAN stays clickable during a scan (it starts disabled, gui/views/setup_view.py:456, is enabled after the first finished scan, gui/pyqt_main_window.py:302, and is never disabled again; only SCAN FILES is disabled, 231); pressing it during a scan clears the scan log and is then silently refused (RAN by MW e17; READ gui/pyqt_main_window.py:229 runs before the guard at scan_controller.py:59-61). No cancel (READ).
9. A folder that vanishes while being listed (for example a network share that drops) is reported as "No Excel files found" rather than as an error (READ utils/helpers.py:264-266).
10. The period, month and year parts are stored but never checked or used (READ core/file_parser.py:102; B-R23) [GST RULE] ⚠.
11. Platform difference: on Linux the glob is case-sensitive (`.XLSX` not found, RAN LT §6.2); on Windows Python's glob ignores case (HIGH, by reasoning - not testable here).

**Older documents against the code.** README.md:124 "GST Files Naming Convention (MUST match exactly)": the code ignores letter case, accepts `.xls` as well as `.xlsx` and accepts a "(n) " prefix for GSTR3B names only (utils/constants.py:15-55). README.md:128-133 gives example names that parse as documented, but says nothing about hyphens or `.xlsm`. README.md:276 ("No Excel files found → Check file extensions (.xlsx, .xls)") omits that `.xlsm` is found but never recognised. README.md:22 ("automatically categorizes … by client and state") is true only when every file spells the client and state identically.

**Origin.** Parser, patterns and completeness logic: edac6a3 (shared - before the split on 30-10-2025); progress-callback guard 6237597 (shared - before the split on 30-10-2025); background thread c258047, size and signature checks moved into the parser plus per-file progress messages 7e09e5c, statistics sent with the result b0e7d0e (all B only (after the split)). Rules: B-R1 to B-R16, B-R17, B-R21, B-R22; related B-R23 to B-R26, B-R29.

---

### B-F7: Scan progress log and scan summary

**What it does.** While a scan runs, every message of the scanner becomes one line in the "SCAN LOG" box on Step 1, and the status bar at the bottom shows a percentage. When the scan ends, a block of counts is added to the scan log and a "SCAN SUMMARY" table appears on Step 2.

**Where it lives.**

| Part | file:line | What the code does |
|---|---|---|
| Start lines | gui/pyqt_main_window.py:229-230, 234-236 | Clears the log; writes "Scanning: <folder>" and "Scan started..."; status "Scanning files...". |
| Progress | gui/pyqt_main_window.py:238-242; messages from core/file_parser.py:147, 151, 155, 158 | One log line per message (colour "normal"); status bar "Scanning... P% (i/N)" with P = int(i ÷ N × 100), only when N > 0. |
| Result block | gui/pyqt_main_window.py:250-269 | "SCAN RESULTS", a line of 40 "=", Files found, Parsed (rate), Unparsed, Total size, Clients, Complete (rate), Incomplete, "Variations: n" only if n > 0, then "FILE TYPES:" sorted by name. |
| Summary table | gui/pyqt_main_window.py:276-299 → gui/views/validation_view.py:117-140, 423-425 | HTML (a web-page text format) table "Scan Summary" / "Clients" / "File Types" shown in a read-only box. |
| End | gui/pyqt_main_window.py:271, 302, 305-312 | Re-enables SCAN FILES, enables RE-SCAN, tab dots (B-F29), switches to Step 2, status "Scan complete: <n> clients, <m> files, <size>". |
| The log box | gui/views/setup_view.py:519-588 (`log_scan_message` 571-585) | Inserts each message as `<span style="color:…">message</span><br>` at the end of one text box (583). |
| Errors | gui/pyqt_main_window.py:313-316 (display error), 318-322 (scan error) | "Error displaying results: …" / "SCAN ERROR: …" + "Scan Error" box + status "Scan failed". |

**Input → Output.** Scanner messages and statistics → text in the Step 1 log, the status bar and the Step 2 summary box.

**User interaction.** None; automatic during and after every scan.

#### B-F7 concrete trace (RAN: `scripts/t_window_flow.py` through the real window, output `work/flow_out.txt`)

Input: the same source folder as B-F6, scanned by clicking SCAN FILES in the real window (offscreen).

1. gui/pyqt_main_window.py:229-230 → log "Scanning: …/work/flow/Source"; 236 → "Scan started...".
2. First message (0, 0, "Finding Excel files...") → logged; line 240 skips the status bar because N = 0.
3. (0, 28, "Found 28 files, processing...") → logged; status "Scanning... 0% (0/28)". Then 28 messages "[i/28] <name>": status "Scanning... 3% (1/28)" … "Scanning... 100% (28/28)" (RAN, status messages captured in order).
4. Completion (244-316): lines 253-269 send the counts with two leading spaces each; the HTML insertion at setup_view.py:583 removes the spaces. Line 261: variations = 3 > 0, so "Variations: 3" is written in the warning colour.
5. Line 274 fills the client table (B-F8); 277-299 build the HTML summary; 299 shows it; 302 enables RE-SCAN; 305-307 dots and switch to Step 2 (RAN: current tab 1); 308-312 status.

Actual output (RAN; scan log as displayed, after the 28 file lines):

```text
SCAN RESULTS
========================================
Files found: 27
Parsed: 25 (92.6%)
Unparsed: 2
Total size: 128.4 KB

Clients: 5
Complete: 1 (20.0%)
Incomplete: 4
Variations: 3

FILE TYPES:
Annual Report: 2
GSTR-2B Reco: 3
GSTR-3B Export: 15
IMS Reco: 1
Sales: 3
Sales Reco: 1
```

Status bar, last message: "Scan complete: 5 clients, 27 files, 128.4 KB". Step 2 summary box (plain text of the HTML): "Scan Summary / Total files: 27 / Parsed: 25 (92.6%) / Unparsed: 2 / Total size: 128.4 KB / Clients / Total clients: 5 / Complete: 1 (20.0%) / Incomplete: 4 / Variations: 3 / File Types / …" (RAN).

**What the trace shows against what the feature should do.** Bugs (for Section 7):

1. **The counts do not add up for the user** (RAN): the log counted "[1/28]" to "[28/28]", but the result says "Files found: 27" and the status bar "27 files"; the rejected 300-byte file is left out of "Files found" (core/file_parser.py:338 counts only files that passed the size and signature checks) and out of "Unparsed: 2", but is in "Variations: 3".
2. Only the **number** of problem files is shown; which files they are and why is never displayed (RAN; see B-F9).
3. Leading spaces vanish (RAN: the indented count lines appear flush left); any text between "<" and ">" in a message would vanish too, because messages are inserted as HTML (RAN for the same code in the Step 3 log, B-F16).
4. **The log slows the window down as it grows**: all lines are one text block, so inserting 1,000 lines took 5.6 s and 5,000 lines 136.6 s of window time (RAN by the views reader, VW t5_pixels). The scan runs in the background, but each per-file message is drawn on the window thread, so a source folder of a few thousand files freezes the window (MEDIUM for real folder sizes).
5. RE-SCAN during a scan clears the log while the first scan keeps writing into it (RAN by MW e17).
6. A scan error leaves the previous scan's summary and client table on Step 2 (READ gui/pyqt_main_window.py:318-322 touch neither).

**Older documents against the code.** README.md:149 ("Review scan summary and statistics") is true. README.md:156 ("Review variations tab for naming issues"): there is no variations tab - only the count (see B-F9).

**Origin.** Status-bar messages and a first, three-line summary HTML on Step 2 ("Total files", "Clients found", "Variations/issues") c258047; the Step 1 scan log with per-file progress (and a short "Scan complete!" block) 7e09e5c; the richer statistics, the "SCAN RESULTS" block in the scan log, the tabled summary with "File Types" and the card collapsing b0e7d0e (all B only (after the split); READ `git show c258047:gui/pyqt_main_window.py`, `git show 7e09e5c:gui/pyqt_main_window.py`, `git show b0e7d0e`). Rules: B-R29, B-R137, B-R113. [Corrected during verification - originally dated "the summary HTML" to b0e7d0e; a simpler summary HTML already existed in c258047.]

---

### B-F8: Client list with completeness status (Validation)

**What it does.** After each scan, Step 2 shows one row per client: tick box, client, state, status, number of files, number of missing types, number of duplicate remarks, a "Folder Name" and the time of the client's last processing run. Rows are coloured by status. "Select All", "Clear All" and "Complete Only" set the ticks; the ticked rows are what DRY RUN and START work on.

**Where it lives.**

| Part | file:line | What the code does |
|---|---|---|
| Table | gui/views/validation_view.py:321-337 | QTreeView (Qt's list-and-columns widget) with sorting by header click switched on (324); columns 'Select', 'Client', 'State', 'Status', 'Files', 'Missing', 'Extra', 'Folder Name', 'Last Run' (327-330). |
| Fill | gui/views/validation_view.py:363-410 `populate_clients`; called at gui/pyqt_main_window.py:274 | Deletes all rows (365); one row per client in key order (367); tick box **ticked** (370) holding the client key (371); numbers stored as text (376-378); "Folder Name" = the client key (379); "Last Run" from the run history (381-385, see B-F17); colours (388-393); count label (407-410). |
| Status rule [GST RULE] | core/file_parser.py:280-314 | Decides "Complete" / "Missing n files" / "Has duplicates" (B-F6 step 8). |
| Ticked keys | gui/views/validation_view.py:412-421 | Returns the keys of ticked rows **in the order shown**. |
| Buttons | gui/views/validation_view.py:282-295, 429-447 | Select All / Clear All / Complete Only (ticks rows whose status is exactly "Complete", unticks the rest). |
| Help text | gui/views/validation_view.py:96-113 (texts 97-101) | "Check client list - green = complete, orange = missing files" … "Click 'Start Processing' when ready". |

**Input → Output.** `client_data` from the scan + the run history → the table; the ticks → the list of client keys handed to DRY RUN (B-F11) and START (B-F12).

**User interaction.** The table fills itself after a scan; the user ticks rows, uses the three buttons, or clicks a column header to sort.

#### B-F8 concrete trace (RAN: `scripts/t_window_flow.py`, output `work/flow_out.txt`)

Input: the scan of B-F6 (five clients), first run ever (empty run history).

1. gui/pyqt_main_window.py:274 → `populate_clients(client_data, run_history)`. Line 367 sorts the keys: `Client-DE`, `ClientA Traders-MH`, `ClientB Private Limited-GJ`, `ClientC Co-TAM`, `ClientC Co-TN`.
2. For each key, lines 368-385 build the cells. "Last Run": `get_last_run_timestamp_for_client(key)` (models/run_history.py:102-107) finds no record → "".
3. Lines 388-393 colour: status exactly "Complete" → #107C10 (green); text containing "Missing" → #FF8C00 (orange); otherwise #D83B01 (red-orange).
4. Line 410: "5 clients (1 complete)".

Actual output (RAN):

| Ticked | Colour | Client | State | Status | Files | Missing | Extra | Folder Name | Last Run |
|---|---|---|---|---|---|---|---|---|---|
| yes | orange | Client | D Exports | Missing 5 files | 1 | 5 | 0 | Client-DE | |
| yes | orange | ClientA Traders | Maharashtra | Missing 3 files | 5 | 3 | 0 | ClientA Traders-MH | |
| yes | green | ClientB Private Limited | Gujarat | Complete | 15 | 0 | 0 | ClientB Private Limited-GJ | |
| yes | orange | ClientC Co | Tamilnadu | Missing 5 files | 1 | 5 | 0 | ClientC Co-TAM | |
| yes | orange | ClientC Co | Tamil Nadu | Missing 4 files | 3 | 4 | 1 | ClientC Co-TN | |

5. Sort by the "Files" column, ascending (the script called `sortByColumn(4, Qt.AscendingOrder)`, the same sort a header click triggers): the Files column became `1, 1, 15, 3, 5` and the processing order (gui/views/validation_view.py:412-421) became `Client-DE, ClientC Co-TAM, ClientB Private Limited-GJ, ClientC Co-TN, ClientA Traders-MH` (RAN) - text order, not number order.
6. "Complete Only" → ticked keys `['ClientB Private Limited-GJ']`; "Select All" → 5 ticked (RAN).
7. After processing (B-F12), the real client folder of <Client B> is `ClientB Pvt Ltd-GJ` (RAN, target tree), while the table's "Folder Name" still says `ClientB Private Limited-GJ`.

**What the trace shows against what the feature should do.** Bugs (for Section 7):

1. **Every client is ticked after a scan, complete or not** (RAN: all five ticked, four of them incomplete). Pressing START at once processes every client.
2. **Number columns sort as text** (RAN: 1, 1, 15, 3, 5), and processing follows the order on screen.
3. **"Folder Name" is not the real folder name**: it shows the scan key, while processing abbreviates "Private"/"Limited" and cuts to the maximum length (RAN: `ClientB Private Limited-GJ` shown, `ClientB Pvt Ltd-GJ` created; see B-F4, B-F12).
4. The lists behind "Missing" and "Extra" (which types are missing, "Multiple Sales (2 files)") are not shown anywhere on this screen - only their counts (RAN).
5. A mis-split name is shown as if it were right: <Client D> appears as client "Client" in state "D Exports" (RAN; cause in B-F6).
6. Two rows with the same client name "ClientC Co" and no hint that they are one client spelled two ways (RAN; cause in B-F6).
7. Ticks are lost at every scan (READ 365, 370). Nothing in this table or its buttons is disabled while processing runs (READ: no `setEnabled` on them; RAN by MW e4).
8. The help text explains only two of the three colours (READ validation_view.py:98 vs 393) and says "Click 'Start Processing'" while the button here reads "START" (READ 101 vs 202).
9. "Last Run" is filled only after the next scan (see B-F17).

**Older documents against the code.** README.md:150-152 says the list marks clients with "✅ = All files present" and "⚠️ = Missing some files": B shows no icons (the parser still computes `status_icon` at core/file_parser.py:308, 311, 314, but nothing displays it - READ grep); it colours the row text. README.md:153-155 (tick boxes, Select All / Clear All / Complete Only) is true.

**Origin.** Completeness logic edac6a3 (shared - before the split on 30-10-2025); the PyQt table c258047, restyled 7e09e5c (both B only (after the split)). Rules: B-R27, B-R134, B-R120; related B-R24 to B-R26.

---

### B-F9: Unrecognised files / variations list

**What it does (intended).** The scanner collects every file it could not use - too small, not an Excel file, unreadable, or a name that matches no pattern - with the reason and a suggested name pattern, so that the user can rename them. **In B only the number of such files reaches the user**; the list itself is never shown, and the processing summary's "Variations" sheet is always empty.

**Where it lives.**

| Part | file:line | What the code does |
|---|---|---|
| Collected | core/file_parser.py:183-215 (size, signature, unreadable), 245-252 (name not recognised, with `pattern_expected`) | Each entry: filename, path, issue, size (+ hint). |
| Hint [GST RULE] ⚠ | core/file_parser.py:316-334 `_suggest_pattern` | Picks one example name by simple word tests ("gstr"+"2b", "ims", "gstr"+"3b", "sales"+"reco", "sales", "annual"). |
| Kept | gui/controllers/scan_controller.py:27, 35 → models/app_state.py:54 (`ScanResult.variations`) | Stored in the window's `scan_result`; **no code reads it** (READ: grep for `variations` finds only these lines, the count lines below, the summary sheet, and `self.variations.clear()` at models/app_state.py:74 inside `ScanResult.clear()`, which nothing calls). |
| Shown | gui/pyqt_main_window.py:261-262 (scan log), 294 (summary box) | Only `statistics['variations']`, a number. |
| Summary sheet | gui/controllers/processing_controller.py:73 (`'variations': []`) → core/excel_handler.py:491-522 | The processing worker hands the summary an always-empty list, so the sheet prints "No file variations found". |

**Input → Output.** Problem files found by the scan → a count in the scan log and summary box; nothing else.

**User interaction.** None (automatic). There is no screen or button that lists the files.

#### B-F9 concrete trace (RAN: `scripts/t_window_flow.py` → `work/flow_out.txt`; `scripts/t_inspect_outputs.py` → `work/inspect_out.txt`)

Input: the same scan (B-F6) and then processing of all five clients (B-F12).

1. During the scan, three entries are collected (RAN, `w.scan_result.variations` in memory): `<Client A> Purchase Register.xlsx` "No matching pattern found" (hint "Check file naming convention"), `GSTR3B-<Client A>-Maharashtra-Jul.xlsm` "No matching pattern found" (hint "GSTR3B-ClientName-State-Month.xlsx"), the 300-byte SalesReco "File too small to be valid Excel".
2. The user sees "Variations: 3" in the scan log (gui/pyqt_main_window.py:262) and in the Step 2 summary (294) (RAN) - with no file names.
3. Processing: gui/controllers/processing_controller.py:73 sets `'variations': []` and nothing ever appends to it (READ grep). core/excel_handler.py:506-509: the list is empty → cell A2 "No file variations found", merged A2:E2.

Actual output (RAN, the "Variations" sheet of `GST_Processing_Summary_211026_0640.xlsx`):

```text
('Filename', 'Issue', 'Expected Pattern', 'Size', 'Action')
('No file variations found', None, None, None, None)
```

**What the trace shows against what the feature should do.** **Broken** (as the registry says). The three problem files were found but are never named to the user, and the summary workbook states the opposite of the truth ("No file variations found") (RAN). Listed for Section 7. The hints computed by `_suggest_pattern` (core/file_parser.py:316-334) are therefore never seen either; when they would be, plain word tests give odd hints (RAN by LT §1.2 row O54 and §1.3 item 11: "Claims" contains "ims" and gets the IMS hint).

**Older documents against the code.** README.md:156 ("Review variations tab for naming issues") and README.md:299 ("Review variations and fix naming issues"): no such tab exists. README.md:200 ("Variations: Files that don't match naming convention") describes the sheet that B always leaves empty.

**Origin.** Collection and the summary sheet edac6a3 (shared - before the split on 30-10-2025). At the split the Tkinter code passed the real list into the summary (`'variations': self.app.variations`, gui/handlers/processing_handler.py:161@47a1492 - READ); c258047 (B only (after the split)) replaced it with an empty list - a regression. Rules: B-R12, B-R2, B-R3, B-R95.

---

### B-F10: Export client list (EXPORT LIST button)

**What it does.** Nothing. Step 2 shows a teal button "EXPORT LIST" next to DRY RUN and START, but no code is connected to it.

**Where it lives.** gui/views/validation_view.py:186-200 (button and style) and 221 (placed in the row). No `clicked.connect` for `btn_export` anywhere (READ: grep), and gui/pyqt_main_window.py:120-172 wires no handler for it.

**Input → Output.** None.

**User interaction.** A click, with no effect.

#### B-F10 concrete trace (RAN: `scripts/t_window_flow.py` → `work/flow_out.txt`)

Input: after the B-F6 scan, the user clicks EXPORT LIST.

1. `btn_export.receivers(btn_export.clicked)` = **0** (RAN): nothing listens.
2. Click → no message box, the tab stays on Step 2, no new file in the work folder (RAN: "new message boxes = [] tab 1 -> 1 new files = set()").

**What the trace shows.** A visible button that does nothing (for Section 7). In the Tkinter version the same idea worked: `export_client_list` (gui/handlers/client_handler.py:221-253@47a1492, READ) asked for a file name `selected_clients_<YYYYMMDD_HHMMSS>.txt`, wrote "Selected Clients List", the date, the count and one "<client> - <state>" line per ticked client, showed "Export Complete" and opened the file. The button was created in c258047 without a connection and the handler was deleted in a62d3cc (both B only (after the split)).

**Older documents against the code.** README.md does not mention exporting the client list.

**Origin.** c258047 (B only (after the split)). No rule.

---

### B-F11: Dry run (preview)

**What it does.** DRY RUN on Step 2 switches to Step 3 and writes into its log, for each ticked client, which files would be organised and that an ITC report and a Sales report would be made. It touches no file.

**Where it lives.** Button gui/views/validation_view.py:169-184 → signal → gui/pyqt_main_window.py:135 → `_on_dry_run` gui/pyqt_main_window.py:326-343: at least one ticked client (327-330, "No clients selected."); switch to Step 3 (332); header line (333); per client "🏢 <client> - <state>" (336); per file "📄 Would organize: <name>" (337-339); always "📊 Would create: ITC Report" and "💰 Would create: Sales Report" (340-341); closing lines (342-343).

**Input → Output.** Ticked client keys + the scan data held by the window → preview lines appended to the Step 3 log.

**User interaction.** Click DRY RUN.

#### B-F11 concrete trace (RAN: `scripts/t_window_flow.py` → `work/flow_out.txt`; `scripts/t_f11_f15_extra.py` → `work/f11x_out.txt`)

**Case A - after the B-F6 scan, all five clients ticked, with the ITC, Sales and Target boxes still empty.**

1. gui/pyqt_main_window.py:327 → five keys; 332 → tab 2 (Step 3) (RAN).
2. Lines 334-341 loop over the five keys in table order; for each, the file types in the order they were found and each file's `name`.
3. No check of templates, target, mode or source files is made anywhere in 326-343 (READ); RAN: no message box although all three boxes were empty.

Actual output (RAN, excerpt as displayed - the leading spaces of lines 336-341 vanish, see B-F16):

```text
🧪 ════════ DRY RUN PREVIEW ════════
🏢 Client - D Exports
📄 Would organize: GSTR-2B-Reco-Client-D Exports-Kerala-FY2425.xlsx
📊 Would create: ITC Report
💰 Would create: Sales Report
🏢 ClientA Traders - Maharashtra
📄 Would organize: (1) GSTR3B-ClientA Traders-Maharashtra-Jun.xlsx
… (four more files) …
📊 Would create: ITC Report
💰 Would create: Sales Report
… (three more clients) …
🧪 ════════ DRY RUN COMPLETE ════════
💡 No files were actually moved or created
```

Target folder after the dry run: does not exist (RAN) - nothing was written.

**Case B - <Client A> processed once (B-F12), then DRY RUN pressed again without re-scanning.**

1. Processing rewrote the file names held in the window's scan data: core/file_organizer.py:216-229 replace each successfully copied file's `name` with the cleaned name and add `sanitized_name` (RAN: before, `'GSTR-2B-Reco-ClientA Traders-Maharashtra-FY2425.xlsx'`; after, `'GSTR-2B-Reco-ClientA_Traders-Maharashtra-FY2425.xlsx'`).
2. `_on_dry_run` reads the same dictionaries (gui/pyqt_main_window.py:335-339).

Actual output (RAN): every line now reads "📄 Would organize: …ClientA_Traders…" with underscores - names that do not exist in the source folder.

**What the trace shows against what the feature should do.** Bugs (for Section 7):

1. **The preview does not check what real processing checks**: with empty template and target boxes it still promised two reports per client, while START would refuse with "Please select a valid ITC template." (RAN; compare gui/pyqt_main_window.py:352-361). The Tkinter dry run validated the inputs first (gui/handlers/processing_handler.py:23-26@47a1492, READ by VW).
2. It always says "Would create: ITC Report / Sales Report" for every client, even <Client D> with one GSTR-2B file and <Client C>-TAM with one GSTR-3B file (RAN) - which does match what processing then does (B-R50) [GST RULE] (which reports a client gets), but gives no hint that most Links cells will stay unwritten.
3. "No files were actually moved" - real processing **copies**, never moves (B-F12); the word is misleading.
4. It ignores the processing mode: in Resume it still says "Would organize" for every file (READ).
5. After a processing run, without a re-scan, it shows cleaned names instead of the real source names (RAN, Case B).
6. The Step 3 log is not cleared first, so the preview and later real runs pile up in one log (RAN: the real run's lines followed the preview lines in `work/flow_out.txt`). One log line per file on the window thread makes a dry run over thousands of files slow (MEDIUM; timing mechanism RAN in B-F7).

**Older documents against the code.** README.md:158-159 places "Dry Run" on the Processing tab; the button is on Step 2 (Validation) (gui/views/validation_view.py:169-184) and only the output appears on Step 3.

**Origin.** The PyQt dry run c258047 (B only (after the split)); a Tkinter dry run with input checks existed at the split (47a1492). Rules: B-R135, B-R50.

---

### B-F12: Organise files into client and category folders (fresh / re-run / resume, backups)

**What it does.** For each ticked client, processing creates a dated folder tree in the target folder and **copies** (never moves) the client's recognised files into it under cleaned names: `<target>/Annual Statement-<DDMMYY HHMM>/<client>-<state code>/Version-<DDMMYY HHMM>/` with three category folders ("GSTR-3B Exports", "Other ITC related files", "Sales related files"); the Annual Report goes into the Version folder itself. The processing mode chosen on Step 1 decides whether a new run folder is made (Fresh) or the latest one is reused (Re-Run, Resume), and what happens when a copy of the same name already exists (backup and overwrite / overwrite / skip).

**Where it lives.**

| Part | file:line | What the code does |
|---|---|---|
| Mode buttons | gui/views/setup_view.py:334-359 (texts 343-347, Fresh ticked 352-353), 631-633 (position 0/1/2 → "fresh"/"rerun"/"resume"), 642-644 (restore) | Three radio buttons. |
| Start | gui/pyqt_main_window.py:345-371 (via START gui/views/validation_view.py:217 or Start Processing gui/views/processing_view.py:95) | Checks (B-F3), then hands settings, the window's scan data and the ticked keys to the controller. |
| Worker loop | gui/controllers/processing_controller.py:42-110, 173-181 | Builds one `FileOrganizer` per run (46-52); per client: folders (99), copies (106), counts (107-109); a raised error fails that client only (173-181). |
| Organiser set-up | core/file_organizer.py:26-65 | Max length 0 or less → 35 (42); its own time stamp (43); unknown mode → error (58-59); creates the target folder (61-65). |
| Folder tree | core/file_organizer.py:67-179 | Client folder name (79-83, utils/helpers.py:513-537, see B-F4); run folder: Fresh = new "Annual Statement-<stamp>" (148-156), Re-Run/Resume = the "Annual Statement-…" sub-folder with the newest modification time (158-169), else a new one (170-179); Version folder (98-104); category folders (106-135). Folder names utils/constants.py:81-90. |
| Copying | core/file_organizer.py:181-283; utils/helpers.py:142-204 | Destination by type (285-298) [GST RULE]; cleaned name (254, utils/helpers.py:66-112); existing copy: Fresh → backup `<stem>_backup_<DDMMYY_HHMM><ext>` then overwrite (261-265, helpers.py:187-204), Resume → "Skipped" (266-270), Re-Run → overwrite (no branch); `shutil.copy2` + size comparison (helpers.py:169-178); after a success the scan record's `name` is replaced by the cleaned name (216-229). |
| Never called | core/file_organizer.py:364-379 (`cleanup_on_error`) with its only helper `_is_folder_empty` (381-390, called only at 373), 392-409 (`get_summary`) | Dead (READ grep). |

**Input → Output.** Ticked client keys + scan data + the mode, tick box and max length from Step 1 → the folder tree with copies (and backups); a list of per-file results used by B-F15. Source files unchanged.

**User interaction.** Choose the mode on Step 1, then START on Step 2 (or Start Processing on Step 3).

#### B-F12 concrete trace

**Case A - through the real window, Fresh mode, all five clients of B-F6 ticked** (RAN: `scripts/t_window_flow.py` → `work/flow_out.txt`; on Linux the two reports take the openpyxl branch - B-F13 - but folder making and copying are the same code as on Windows).

1. START → gui/pyqt_main_window.py:345: five keys (346); `_collect_settings` (352) reads mode "fresh" (setup_view.py:632-633, radio 0), tick off, max 35; the templates exist (353-358); the target box is not empty (359); 364 switches to Step 3; 366-371 pass `self.scan_result.client_data` - the window's own dictionaries, not a copy.
2. gui/controllers/processing_controller.py:259 - no run active; 263 → Start Processing disabled, Stop enabled (RAN); the thread starts `_process`.
3. Line 44 time stamp "211026 0640"; 46-52 `FileOrganizer(target, "fresh", include_client_name=False, client_folder_settings={}, client_name_max_length=35)`: core/file_organizer.py:42 keeps 35; 43 takes its own stamp; 58 "fresh" is valid; 63 creates the target folder, which did not exist.
4. Loop, idx 0 = `Client-DE` (table order), idx 1 = `ClientA Traders-MH`. For <Client A>:
   - processing_controller.py:99 → `create_client_structure` (file_organizer.py:67-144): 78 "Maharashtra" → MH [GST RULE] ⚠; 79-83 key "ClientA Traders-MH" (18 characters, not cut); 86 cleaned name "ClientA_Traders"; 90 → `_get_level1_folder`: 148 Fresh → 150-155 "Annual Statement-211026 0640" (already made for the first client in this run; `mkdir` accepts an existing folder); 93-96 client folder; 98-104 "Version-211026 0640"; 113-135, three times: 121 tick off → 126 per-client dictionary empty (B-F5) → 130 removes " (ClientA_Traders)".
   - processing_controller.py:106 → `organize_files` (181-231): 199-202 collect the files in type order - GSTR-3B Export (3), Annual Report, GSTR-2B Reco. For each, `_organize_single_file` (233-283): 248 destination by type (GSTR-3B → "GSTR-3B Exports"; Annual Report → the Version folder; GSTR-2B Reco → "Other ITC related files") [GST RULE]; 254 cleaned name, e.g. "(1) GSTR3B-ClientA Traders-Maharashtra-Jun.xlsx" → "(1)_GSTR3B-ClientA_Traders-Maharashtra-Jun.xlsx"; 260 no such file yet; 273 `safe_copy_file` → `shutil.copy2` + equal sizes → "Success"; 217-229 the scan record's `name` becomes the cleaned name.
   - 107-109: 5 successes → log "✓ Organized 5 files".

Actual output for <Client A> (RAN, target tree):

```text
Annual Statement-211026 0640/
    ClientA Traders-MH/
        Version-211026 0640/
            AnnualReport-ClientA_Traders-Maharashtra-2024-25.xlsx
            GSTR-3B Exports/
                (1)_GSTR3B-ClientA_Traders-Maharashtra-Jun.xlsx
                GSTR3B-ClientA_Traders-Maharashtra-Apr.xlsx
                GSTR3B-ClientA_Traders-Maharashtra-May.xlsx
            ITC_Report_ClientA_Traders_MH_211026_0640.xlsx        (B-F13)
            Other ITC related files/
                GSTR-2B-Reco-ClientA_Traders-Maharashtra-FY2425.xlsx
            Sales related files/                                  (empty - no Sales files)
            Sales_Report_ClientA_Traders_MH_211026_0640.xlsx      (B-F14)
            _Organization_Report_ClientA Traders_Maharashtra.txt  (B-F15)
```

The other clients (RAN): `Client-DE/` (the mis-split <Client D>, B-F6), `ClientB Pvt Ltd-GJ/` (15 copies, names with "Pvt_Ltd"), `ClientC Co-TAM/` and `ClientC Co-TN/` (one client in two folders), and `GST_Processing_Summary_211026_0640.xlsx` in the run folder. Message box "Successfully processed 5/5 clients". Source files unchanged (RAN: `work/f12src_out.txt`, all 30 fingerprints identical).

**Case B - the three modes, on the organiser itself** (RAN: `scripts/t_f4_f5_f12.py` → `work/f12_out.txt`). Input: <Client A> with five files: GSTR-2B Reco, Annual Report and three GSTR-3B files named `…-Apr.xlsx`, `…-Apr_1.xlsx` and `…-Apr 1.xlsx` (the last two **clean to the same name** `…-Apr_1.xlsx`). The organiser's minute stamp was fixed in the test so that "same minute" and "next minute" could be shown on purpose; backups use the real clock (0641). Each run scanned the folder again first, as the user would.

| Run | Mode, minute | What the code did (lines) | Result (RAN) |
|---|---|---|---|
| 1 | Fresh, 1000 | New run folder (148-156). "Apr 1" copied as `…Apr_1.xlsx`; then "Apr_1" finds that name (260) → Fresh → backup of the first copy (263) → overwrite (273) | 5 × Success; GSTR-3B folder holds `…Apr.xlsx`, `…Apr_1.xlsx` and `…Apr_1_backup_211026_0641.xlsx`. By the copy order ("Apr 1" sorts before "Apr_1"), the backup holds "Apr 1" and `…Apr_1.xlsx` holds "Apr_1" (READ; the contents were checked by fingerprint in the same situation by the live-trace reader, LT §6.5) |
| 2 | Re-Run, 1001 | Latest run folder by modification time (167) = "…211026 1000"; new "Version-211026 1001" (98-104); the collision overwrites without backup (no Re-Run branch at 260-270) | 5 × Success; only **two** GSTR-3B files in the new version - by the same copy order the content of "Apr 1" is gone from it (READ; RAN with fingerprints by LT §6.5) |
| 3 | Resume, 1002 | New "Version-211026 1002"; files do not exist there yet, so nothing is skipped except the collision (266-270) | 4 × Success + 1 × Skipped - Resume copied everything again |
| 4 | Resume, **same** minute 1002 | Same Version folder; every destination exists → 266-270 | 5 × Skipped "File already exists". The ITC Links value prepared afterwards names the **uncleaned** file `GSTR-2B-Reco-ClientA Traders-Maharashtra-FY2425` while the folder holds `…ClientA_Traders…` (the name rewrite at 216-229 happens only after a Success) |
| 5 | Re-Run, same minute 1002 | Same Version folder; existing copies overwritten, no backup | 5 × Success, no backups |
| 6 | Fresh, 1000 again (same names as run 1) | Same run and Version folder names; every copy backed up then overwritten | 5 × Success; 5 backup operations (RAN list) gave 3 new backup files (Annual Report, GSTR-2B Reco, `…Apr`) - the existing `…Apr_1_backup_211026_0641.xlsx` was silently overwritten twice in the same minute (same backup name); by the copy order it ends holding "Apr 1" again (READ) |
| 7 | Re-Run, 1003, after a person saved a note into an older batch folder "Annual Statement-150926 0900" | 167 picks the newest **modification time** | New version written into the **older** batch "Annual Statement-150926 0900" |

**Case C - one client fails** (RAN: `scripts/t_f11_f15_extra.py` → `work/f11x_out.txt`; here the worker's `_process` was called directly, with the minute stamp fixed at "211026 1300", so that the stray file could be placed in advance; the code path inside the worker is the one the window uses). A stray *file* named `ClientA Traders-MH` sits where <Client A>'s client folder must be created. file_organizer.py:95 raises "[Errno 17] File exists"; 140-144 log and re-raise; processing_controller.py:173-181 count the client failed and log "❌ Failed ClientA Traders-MH: [Errno 17] File exists: …"; the next client continues. Result: 2 successful, 1 failed, 3 total. When a client fails *after* its folders exist, the partly made folders and copies stay, because `cleanup_on_error` (364-379, with its helper 381-390) is never called (READ grep; RAN by LT §6.7).

**On-screen texts about the modes, side by side (READ):**

| Where | Fresh | Re-Run | Resume |
|---|---|---|---|
| Mode buttons (gui/views/setup_view.py:344-346) | "Create new folder structure (recommended for first time)" | "Add new version folders to existing structure" | "Continue previously interrupted processing" |
| Help card (gui/views/setup_view.py:510-512) | "New timestamped structure" | "Overwrite existing" | "Skip already processed" |
| What the code does (Case B) | new run folder each run (same name within one minute) | reuses the most recently *modified* run folder, always adds a new minute-stamped Version folder, overwrites only within the same minute | same as Re-Run, except that within the **same minute** existing copies are skipped; every ticked client is processed again in full |

**What the trace shows against what the feature should do.** Bugs (for Section 7):

1. **Resume does not continue an interrupted run**: it creates a new Version folder and copies everything again; it skips only files copied in the same clock minute (RAN runs 3-4).
2. **Re-Run "overwrite existing" happens only within the same minute**; otherwise it is "reuse the batch folder, add a version" (RAN runs 2, 5). The three on-screen descriptions disagree with each other and with the code (table above).
3. **Two source files whose cleaned names coincide keep only one copy**, while both report "Success" (RAN run 1: the first survives only as a `_backup_`; run 2: lost entirely; run 3: "Skipped").
4. **Re-Run/Resume choose the batch folder by modification time** (RAN run 7): touching an older batch (saving a note into it) sends the next run there. Step 4 chooses by name instead (B-F18) - two different "latest" rules.
5. Same-minute Fresh runs reuse the same folders and add `_backup_` files that stay in the category folders (and, for the Annual Report, in the Version folder itself) for good (RAN run 6); a second backup in the same minute overwrites the first (RAN). If a template's Power Query reads whole folders, it would read the backups too (MEDIUM - the templates are not in the repository; how to check: open the ITC template's queries and see whether they filter file names).
6. **The window's scan data is rewritten** with cleaned names (RAN, B-F11 Case B), and after a skipped or failed copy the uncleaned name is what goes into the report's Links cell (RAN run 4) [GST RULE] (which file feeds the report).
7. Partly made folders stay when a client fails part-way (READ; RAN by LT §6.7). Two different clients can be folded into one client folder by the abbreviation or the cut (RAN by LT §6.6; see B-F4).
8. A copy is checked by size only; an interrupted copy can leave a partial file under the final name (READ utils/helpers.py:169-178; MEDIUM).
9. Nothing checks the length of the resulting paths (READ; RAN by the outside-files reader: 233-268 characters for a 44-character target and a 40-character client name, OF §5.4).
10. A source box changed after the last scan is not noticed (READ; see B-F3).

**Older documents against the code.** README.md:142-145 ("Re-Run: Add new versions to existing structure" - true; "Resume: Continue interrupted processing" - not what the code does). README.md:180-190: the tree is right in outline, but the brackets "(ClientName)" appear only when the tick box is on, the Annual Report copy (in the Version folder) is not shown, and the client folder name is abbreviated and cut (B-F4). README.md:306 ("Original files are never modified") is true (RAN). README.md:307 ("Backups created for existing files") holds only in Fresh mode and only for a same-minute collision.

**Origin.** Organiser, modes and backups edac6a3 (shared - before the split on 30-10-2025); since the split core/file_organizer.py has changed only by one removed comment in c258047 (B only (after the split); READ `git diff 47a1492 B/main -- core/file_organizer.py`); the PyQt screens and worker c258047 (B only (after the split)). Rules: B-R30 to B-R41, B-R42 to B-R47.

---

### B-F13: Create ITC report from template

**What it does.** For every processed client, right after its files are copied, the program copies the user's ITC template into the client's Version folder under a report name and writes into the template's "Links" sheet the folder paths and file names of the client's GSTR-3B exports, Annual Report, GSTR-2B Reco and IMS Reco, so that the template's own Power Query (Excel's built-in data-import feature) can later find those files. On the author's Windows PC this is done by driving Excel through **COM** (Windows' remote-control mechanism, via pywin32); without pywin32 or off Windows, a fallback writes the cells with openpyxl.

**Where it lives.**

| Step | file:line | What the code does |
|---|---|---|
| Report name | gui/controllers/processing_controller.py:111-122 | `ITC_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>` - client cleaned by `sanitize_filename` (112), state code from the scan or `get_state_code` (113; the fallback is evaluated every time), stamp with "_" (114). |
| Per-report checks | gui/controllers/processing_controller.py:204-239 | Template path empty (210-213) or missing (215-218) → error line, `report_errors` + 1; output `<Version folder>/<name>.xlsx` - always `.xlsx` (220); success / failure lines (225-239). |
| Links values [GST RULE] ⚠ | core/excel_handler.py:540-627 `prepare_template_data` (ITC part 553-588; `return data` 627); cell map utils/constants.py:96-108 | B2 = "GSTR-3B Exports" folder (always); B4 = Version folder, B5 = Annual Report name without extension (only if one exists); B7 = "Other ITC related files" folder, B8 = GSTR-2B Reco name (only if one exists); B10 = same folder, B11 = IMS Reco name (only if one exists). Only the **first** file of each type is used. Paths in Windows form (utils/helpers.py:311-315). |
| Branch | core/excel_handler.py:19-28, 65-93 | `USE_WIN32COM` is True only on Windows with pywin32 (21-25); template must exist (78-79). |
| COM branch | core/excel_handler.py:95-217, 264-306 | `CoInitialize` (107; switches COM on for this thread); `GetActiveObject` - attach to a running Excel - else `Dispatch` - get an Excel through COM, which may reuse one (110-115); hide, no alerts, no screen updates (118-120); copy template to the report path, wait 0.5 s (124-125); open (129, 264-306); find "Links" (any case), else a sheet containing "link", else **add** one (136-153); write non-empty values as text (159-167); Save, Close(SaveChanges=True), wait 0.5 s (173-180); success = file exists (183-189); `finally` (the clean-up block Python runs even after an error): screen updates and alerts back on, **always `Quit()`** (`self.excel_app` is never set, 203-211), `CoUninitialize` (213-217). |
| openpyxl branch | core/excel_handler.py:219-262 | Copy, load with `keep_links=True`, write into an exact (any-case) "Links" sheet only, save; success = file exists. Drops Power Query when it saves (B-R97). |

**Input → Output.** ITC template + the client's folders and (cleaned) file names → `Version-…/ITC_Report_<client>_<code>_<DDMMYY_HHMM>.xlsx` with the Links cells filled; `True`/`False` back to the worker.

**User interaction.** None; automatic during processing (one per client).

#### B-F13 concrete trace

**Case A - the author's Windows PC, Excel not already open, <Client A> (GSTR-3B, GSTR-2B Reco, Annual Report; no IMS Reco).** Traced by READ; the order of the requests B sends to Excel was recorded with stand-in COM objects (fake-COM run: `scripts/t_f13_f14_fakecom.py` → `work/f13_out.txt`; it proves Python's control flow only, never what real Excel does). Folder making, copying and `prepare_template_data` are B's real code.

1. processing_controller.py:112 `sanitize_filename("ClientA Traders")` → "ClientA_Traders"; 113 → "MH" [GST RULE] ⚠; 114 → "211026_0640"; 117-122 → `_create_report(…, 'ITC', itc_template, "ITC_Report_ClientA_Traders_MH_211026_0640", …)`.
2. 208 log "📊 Creating ITC report..."; 210 and 215 pass; 220 output path in the Version folder; 221 log "Output: …".
3. 224 `prepare_template_data` (excel_handler.py:553-588). Because copying ran first, the file names are the cleaned ones (B-F12 step 4). Values (fake-COM run printout, path shortened):

```text
gstr3b_folder   = '\…\Annual Statement-…\ClientA Traders-MH\Version-…\GSTR-3B Exports'
annual_folder   = '\…\ClientA Traders-MH\Version-…'
annual_filename = 'AnnualReport-ClientA_Traders-Maharashtra-2024-25'
gstr2b_folder   = '\…\Version-…\Other ITC related files'
gstr2b_filename = 'GSTR-2B-Reco-ClientA_Traders-Maharashtra-FY2425'
ims_folder      = ''      ims_filename = ''      (no IMS Reco file)
```

4. 225 → `create_report_from_template` → 78 template exists → 83 COM branch. Recorded calls, in order (fake-COM RAN):

```text
pythoncom.CoInitialize()
GetActiveObject('Excel.Application') -> fails (no Excel running)
Dispatch('Excel.Application') -> new instance APP1
APP1.Visible = False ; APP1.DisplayAlerts = False ; APP1.ScreenUpdating = False
APP1.Workbooks.Open(ITC_Report_ClientA_Traders_MH_….xlsx)
WRITE Links!B2, B4, B5, B7, B8          (B10, B11 skipped: empty values, line 162)
Workbook.Save() ; Workbook.Close(SaveChanges=True)
APP1.ScreenUpdating = True ; APP1.DisplayAlerts = True
APP1.Quit()
pythoncom.CoUninitialize()
```

5. 183-186: the file exists → True → processing_controller.py:226-230: `reports_generated` + 1, `itc_reports_created` + 1, log "✅ Success: ITC_Report_ClientA_Traders_MH_….xlsx".

**Case B - the same, but the user already has Excel open with an unsaved workbook** (fake-COM RAN, same output file): `GetActiveObject` **returns the user's Excel** (line 111); B sets `Visible = False` on it, writes and saves its report inside it, switches alerts back on and calls `Quit()` while the user's `Book1.xlsx` is still open in it (recorded: "USER_EXCEL.Quit() [open workbooks at quit: ['C:\Users\me\Documents\Book1.xlsx']]"). What real Excel then does - a "save changes?" question for Book1, possibly in a hidden window that B would wait on - could not be run here (MEDIUM/LOW; how to check: on a Windows test PC open Excel with an unsaved Book1 and process one dummy client).

**Case C - the template has no "Links" sheet** (fake-COM RAN): lines 142-153 find none and **add** a new empty sheet named "Links", fill it and report success ("✅ Success"); only the program log says "Links sheet not found! Creating one...". On the openpyxl branch nothing is written at all and success is still reported (READ 244-258).

**Case D - Excel dies while opening the copy** (fake-COM RAN): the open raises "The RPC server is unavailable."; line 191 catches it and would return False, but the `finally` block's `excel.ScreenUpdating = True` (204) raises again because Excel is gone, so `Quit()` and `CoUninitialize()` are skipped and that second error escapes to processing_controller.py:235-239 ("❌ ITC Report Error: …", `report_errors` + 1). **The unfilled template copy stays under the report's normal name** (RAN: file left `ITC_Report_ClientA_Traders_MH_211026_1000.xlsx` with the template's own text in Links!B2) - and Step 4 will later treat it as a real report (B-F18).

**Case E - what the cells really hold, Linux openpyxl branch** (RAN: `scripts/t_window_flow.py` + `scripts/t_inspect_outputs.py` → `work/inspect_out.txt`). The made-up template carries its own marker text in every Links cell B2-B11 ("TEMPLATE-ITC-B<n>"), so it is visible which cells B overwrote:

| Cell ⚠ | <Client A> | <Client C>-TN (GSTR-3B + 2 Sales only) |
|---|---|---|
| B2 | `\…\ClientA Traders-MH\Version-211026 0640\GSTR-3B Exports` | `\…\ClientC Co-TN\Version-211026 0640\GSTR-3B Exports` |
| B3 | TEMPLATE-ITC-B3 (never written) | TEMPLATE-ITC-B3 |
| B4 | `\…\Version-211026 0640` | TEMPLATE-ITC-B4 (no Annual Report) |
| B5 | `AnnualReport-ClientA_Traders-Maharashtra-2024-25` | TEMPLATE-ITC-B5 |
| B6 | TEMPLATE-ITC-B6 (never written) | TEMPLATE-ITC-B6 |
| B7 | `\…\Version-211026 0640\Other ITC related files` | TEMPLATE-ITC-B7 (no GSTR-2B Reco) |
| B8 | `GSTR-2B-Reco-ClientA_Traders-Maharashtra-FY2425` | TEMPLATE-ITC-B8 |
| B9 | TEMPLATE-ITC-B9 (never written) | TEMPLATE-ITC-B9 |
| B10 | TEMPLATE-ITC-B10 (no IMS Reco) | TEMPLATE-ITC-B10 |
| B11 | TEMPLATE-ITC-B11 | TEMPLATE-ITC-B11 |

(The Linux run's paths start with "\root\…" because the code only swaps "/" for "\"; on Windows they would be drive paths such as "D:\…".)

**What the trace shows against what the feature should do.** Bugs (for Section 7):

1. **Step 3 takes over an Excel the user already has open, hides it and closes it** (fake-COM RAN; the comment at core/excel_handler.py:206 "Don't quit if Excel was already open" says the opposite of what lines 207-211 do).
2. **When a file type is missing, the template's own cell content stays** (RAN Case E: every cell except B2 kept the template's text for <Client C>-TN; B10/B11 for <Client A>; B3, B6 and B9 are never written for anyone). Whatever path the template holds in those cells is what its queries would read [GST RULE] ⚠ - whether the cells should be cleared is the author's call.
3. **A report is made even when the client has none of the ITC inputs** (RAN: <Client C>-TN's ITC report, only B2 written) and is counted "Success" [GST RULE] (B-R50: which reports a client gets).
4. **No "Links" sheet → an empty one is created (COM) or nothing is written (openpyxl), and the run reports success** (fake-COM RAN, READ).
5. **If Excel dies mid-way, an unfilled template copy is left under the real report name**, Excel is not told to quit and COM is not released (fake-COM RAN Case D).
6. Only the first file of each type is linked; further files of that type are copied but not named in the template (READ excel_handler.py:561-562, 572-573, 583-584) [GST RULE] - which file should feed the report is the author's call.
7. After a skipped (Resume) or failed copy, the Links cell names the original uncleaned file, which is not in the folder (RAN B-F12 run 4) [GST RULE] (B-R57: which file feeds the report).
8. "Report created" means only "the file exists" (READ 183-189); a failed report does not make the client fail (RAN B-F3 trace: "Successfully processed 1/1 clients" with ITC Reports Created 0).
9. Two Excel starts and two `Quit()` per client (one per report) (fake-COM RAN, READ) - slow for many clients.
10. An `.xlsm` (or `.xltx`) template is saved under a `.xlsx` name (READ processing_controller.py:220); Excel normally refuses such a file (MEDIUM, not testable here).
11. A same-minute re-run overwrites the existing report without a backup (READ 124; RAN by the core reader).
12. Absolute paths are written, so moving or copying the output tree to another place leaves the Links pointing to the old place (READ utils/helpers.py:311-315; MEDIUM) [GST RULE] (B-R56: how the Links values are formed - which form is wanted is the author's call).
13. All program warnings are switched off for the whole program by core/excel_handler.py:17 (`warnings.filterwarnings('ignore')`), including openpyxl's own warnings about features it drops (READ).

**Comparison with the author's Standard 12, "Excel through COM" (facts only; rebuild_pipeline_FINAL.md:844-878).**

| Standard 12 says | B's Step 3 report creation does |
|---|---|
| Start a separate hidden Excel with `DispatchEx`; never take over the user's own open Excel | `GetActiveObject` first (attaches to the user's Excel), else `Dispatch` (core/excel_handler.py:110-115) |
| `Visible = False`, `DisplayAlerts = False` | Set (118-119), but on the user's instance when attached; alerts switched back **on** just before `Quit()` (205, 209) |
| `Quit()` in a `finally`, even after errors | In `finally` (209), but skipped when line 204 raises (Case D), and it also quits the user's instance (Case B) |
| Open with `IgnoreReadOnlyRecommended=True`; if `ReadOnly`, close and ask the user | Plain `Workbooks.Open(path)` (286, 301); `ReadOnly` never read |
| Before an openpyxl save into a workbook the app did not create, check for Power Query (`DataMashup`); never let openpyxl save over it | The openpyxl fallback (219-262) saves into the template copy with no such check; it cannot be reached through the app today because the window needs pywin32 to start (B-F1) |
| `CoInitialize` in worker threads; Excel work in a separate function; `CoUninitialize` after it returns | `CoInitialize` (107) and `CoUninitialize` (215) inside the same function while Excel objects are still held |

**Older documents against the code.** README.md:42 and 260 ("All Excel operations use temporary files") - the template is copied straight to the final report name (core/excel_handler.py:124). README.md:44 and 262 ("No VBA macros or external links preserved") - the COM branch saves the template as it is; the openpyxl branch opens with `keep_links=True` (231). README.md:210-219 (ITC Links cells B2, B4, B5, B7, B8, B10, B11) agrees with utils/constants.py:96-108 ⚠. README.md:188 and 202 (name "ITC_Report_ClientName_State_DDMMYYYY_HHMMSS.xlsx") - the real name is `ITC_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx`. README.md:278 ("Template not found → Ensure templates have 'Links' sheet") - "Template not found" means the template path is missing; a missing Links sheet gives no error.

**Origin.** edac6a3 (shared - before the split on 30-10-2025); c258047 (B only (after the split)) only removed debug log lines and narrowed `except` clauses (READ `git diff 47a1492 B/main -- core/excel_handler.py`); the PyQt worker that calls it c258047. Rules: B-R48 to B-R60, B-R96, B-R97, B-R98, B-R100, B-R114.

---

### B-F14: Create Sales report from template

**What it does.** Right after the ITC report, the same code makes a Sales report for the client from the Sales template, writing the folder paths and file names of the client's Sales file, Annual Report and Sales Reco file into its "Links" sheet.

**Where it lives.** gui/controllers/processing_controller.py:124-130 (name `Sales_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>`), 204-239 (same `_create_report`); core/excel_handler.py:590-622 (Sales values); the same COM / openpyxl code as B-F13 (65-306); cell map utils/constants.py:109-119 [GST RULE] ⚠ (B-R55): B2 = "Sales related files" folder and B3 = Sales file name (only if a Sales file exists); B5 = Version folder and B6 = Annual Report name (only if one exists); B8 = "Sales related files" folder and B9 = Sales Reco name (only if one exists); B4 and B7 never written.

**Input → Output.** Sales template + the client's folders and file names → `Version-…/Sales_Report_<client>_<code>_<DDMMYY_HHMM>.xlsx`.

**User interaction.** None; automatic, once per client.

#### B-F14 concrete trace

**Case A - Windows, <Client A> (no Sales and no Sales Reco file)** (READ + fake-COM RAN, `work/f13_out.txt`): `prepare_template_data` returns `sales_folder ''`, `sales_filename ''`, `annual_folder '\…\Version-…'`, `annual_filename 'AnnualReport-ClientA_Traders-Maharashtra-2024-25'`, `sales_reco_folder ''`, `sales_reco_filename ''`. Recorded calls: `CoInitialize`, `GetActiveObject` fails, `Dispatch` → a **second** new instance for this client, hide, Open `Sales_Report_…`, **WRITE Links!B5, B6 only**, Save, Close, alerts on, `Quit()`, `CoUninitialize`. Returned True → "✅ Success: Sales_Report_ClientA_Traders_MH_….xlsx".

**Case B - what the cells hold (Linux openpyxl branch, RAN `work/inspect_out.txt`), template marker text "TEMPLATE-SALES-B<n>":**

| Cell ⚠ | <Client A> | <Client C>-TN (two Sales files "Apr-Sep" and "Oct-Mar", no Annual Report, its only Sales Reco rejected as too small) |
|---|---|---|
| B2 | TEMPLATE-SALES-B2 (no Sales file) | `\…\ClientC Co-TN\Version-211026 0640\Sales related files` |
| B3 | TEMPLATE-SALES-B3 | `Sales-ClientC_Co-Tamil_Nadu-Apr-Sep` - the first Sales file only |
| B4 | TEMPLATE-SALES-B4 (never written) | TEMPLATE-SALES-B4 |
| B5 | `\…\ClientA Traders-MH\Version-211026 0640` | TEMPLATE-SALES-B5 (no Annual Report) |
| B6 | `AnnualReport-ClientA_Traders-Maharashtra-2024-25` | TEMPLATE-SALES-B6 |
| B7 | TEMPLATE-SALES-B7 (never written) | TEMPLATE-SALES-B7 |
| B8, B9 | TEMPLATE-SALES-B8 / B9 (no Sales Reco) | TEMPLATE-SALES-B8 / B9 |

The second Sales file of <Client C>-TN, `Sales-ClientC_Co-Tamil_Nadu-Oct-Mar.xlsx`, was copied into "Sales related files" (RAN target tree) but is named in no cell.

**What the trace shows against what the feature should do.** Bugs (for Section 7) - the same mechanism and defects as B-F13 (items 1, 2, 4, 5, 7-13 there apply unchanged), plus:

1. **A client with two Sales files gets only the first linked** (RAN: "Apr-Sep" linked, "Oct-Mar" not) [GST RULE] - whether both halves of the year should feed the report is the author's call.
2. **A Sales report is made for a client with no Sales data at all** (RAN: <Client A>'s Sales report has only B5/B6; B2/B3/B8/B9 keep the template's text) and is counted as created [GST RULE] (B-R50: which reports a client gets).
3. A second Excel start and `Quit()` per client, right after the ITC report's (fake-COM RAN). With an Excel already open, the Sales report's `GetActiveObject` (core/excel_handler.py:111) attaches to the user's Excel again **if it is still running** after the ITC report's `Quit()` - for example because it is waiting on a "save changes?" question for the user's own workbook (B-F13 Case B) - so the user's Excel can be taken over twice per client (MEDIUM: the fake-COM run proves only that B asks for the running Excel again; whether one is still running depends on real Excel - how to check: the Windows test of B-F13 Case B, watching whether the Sales report is made inside the user's Excel). [Corrected during verification - originally stated as certain that "the user's Excel is taken over twice".]

**Older documents against the code.** README.md:221-227 (Sales cells B2, B3, B5, B6, B8, B9) agrees with utils/constants.py:109-119 ⚠. README.md:189 and 203 give the name as "Sales_Report_ClientName_State_DDMMYYYY_HHMMSS.xlsx"; the real name is `Sales_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx`.

**Origin.** edac6a3 (shared - before the split on 30-10-2025). Rules: B-R55, B-R56, B-R51; shared with B-F13: B-R48, B-R50, B-R52, B-R53, B-R57, B-R59, B-R60, B-R96.

---

### B-F15: Processing summary workbook and organisation report

**What it does.** At the end of a processing run the program writes one Excel workbook summarising the run into the run folder (`GST_Processing_Summary_<DDMMYY_HHMM>.xlsx`, five sheets), and during the run it writes, for each processed client, a small text file listing that client's folders and files into its Version folder (`_Organization_Report_<client>_<state>.txt`).

**Where it lives.**

| Part | file:line | What the code does |
|---|---|---|
| Data collected | gui/controllers/processing_controller.py:58-76 (the `summary_data` dictionary; `'variations': []` at 73), 108, 135-147 (successful clients; `'status': 'Success'` typed in at 141), 162-169 (one file-mapping row per file result), 173-181 (failed clients → `errors`), 210-237 (report failures only raise `report_errors`) | |
| Summary written | gui/controllers/processing_controller.py:183-193 → core/excel_handler.py:308-346 | Only if at least one client succeeded (184); progress 95 "Creating summary report..." (185); path `<run folder>/GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` (187-188); "✅ Summary report created" only when the call returns True (190-191). `create_summary_report` catches every error itself and returns False (344-346), so the `except` at 192-193 cannot run. |
| Sheets | core/excel_handler.py:369-399 Summary; 401-428 Client Status; 430-456 File Mapping; 458-489 Errors; 491-522 Variations; 524-538 column widths (max 50); 348-367 styles (white bold 12 pt on #366092) | openpyxl, a brand-new workbook (no Power Query involved). |
| Organisation report | gui/controllers/processing_controller.py:132-133 → core/file_organizer.py:300-362 | Name from the raw client text and full state name (313); header, folders relative to the run folder (329-333), files per type (338-343), "Total Files", "Successfully Copied", "Folders Created" (345-351); UTF-8 (354); failures only logged (360-362). |

**Input → Output.** The run's results → `<run folder>/GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` and `<Version folder>/_Organization_Report_<client>_<state>.txt` per client.

**User interaction.** None; automatic.

#### B-F15 concrete trace

**Case A - the five-client Fresh run of B-F12** (RAN: `scripts/t_window_flow.py` + `scripts/t_inspect_outputs.py` → `work/inspect_out.txt`).

1. processing_controller.py:184 - five successes and a run folder → 188 path `…/Annual Statement-211026 0640/GST_Processing_Summary_211026_0640.xlsx` → 190 `create_summary_report`.
2. excel_handler.py:318-322 new workbook, default sheet removed; 328-332 the five sheets; 335-336 widths; 339 save; 342 exists → True → log "✅ Summary report created".

Actual output (RAN):

```text
Summary:       Processing Date '211026 0640' | Total Clients 5 | Successful 5 | Failed 0 | Total Files 25 |
               Reports Generated 10 | ITC Reports Created 5 | Sales Reports Created 5 | Report Creation Errors 0 |
               Processing Mode 'fresh' | Target Folder '…/work/flow/Target' | Include Client Name in Folders 'No'
Client Status: ('Client', 'D Exports', 'Success', 1, 5, '16.7%')
               ('ClientA Traders', 'Maharashtra', 'Success', 5, 3, '50.0%')
               ('ClientB Private Limited', 'Gujarat', 'Success', 15, 0, '100.0%')
               ('ClientC Co', 'Tamilnadu', 'Success', 1, 5, '16.7%')
               ('ClientC Co', 'Tamil Nadu', 'Success', 3, 4, '33.3%')
File Mapping:  25 rows: original name, client key, type, full destination path, 'Success'
Errors:        'No errors encountered'
Variations:    'No file variations found'      (the scan had found 3 problem files - B-F9)
```

3. Organisation report of <Client A>, the **second** client processed (file_organizer.py:300-362). Lines 346-351 count `self.copied_files` and `self.created_folders`, which belong to the organiser object shared by the whole run (created once at processing_controller.py:46-52). Actual output (RAN):

```text
GST FILE ORGANIZATION REPORT
==================================================
Generated: 2026-10-21 06:40:29
Client: ClientA Traders
State: Maharashtra
Processing Mode: fresh
FOLDER STRUCTURE: VERSION / GSTR3B / ITC / SALES paths relative to the run folder
FILES ORGANIZED:  GSTR-3B Export (3 files) … Annual Report (1 files) … GSTR-2B Reco (1 files) …
STATISTICS:
Total Files: 5
Successfully Copied: 6
Folders Created: 12
```

<Client A> copied 5 files and got 6 folders; the 6 and 12 include <Client D>'s copy and folders from earlier in the run. For <Client C>-TN, the fifth client, the report says "Total Files: 3 / Successfully Copied: 25 / Folders Created: 30" (RAN). File name: `_Organization_Report_ClientC Co_Tamil Nadu.txt` - raw names with spaces (RAN).

**Case B - one client fails** (RAN: `scripts/t_f11_f15_extra.py` → `work/f11x_out.txt`; the failure of B-F12 Case C). Summary sheet: Total Clients 3, Successful 2, **Failed 1**. Client Status: only `Client` (D Exports) and `ClientB Private Limited` - **the failed <Client A> is not listed at all**. Errors: `('11:32:50', 'ClientA Traders-MH', 'Processing Error', "[Errno 17] File exists: '…/ClientA Traders-MH'", 'Review')`. Message box text: "Successfully processed 2/3 clients".

**Case C - a report fails** (RAN in B-F3: ITC template path pointing to a folder): Summary sheet "ITC Reports Created 0", "Report Creation Errors 1"; Client Status "Success"; Errors sheet "No errors encountered".

**Case D - the summary cannot be written** (RAN: `scripts/t_f15_f16_f17.py` → `work/f16_out.txt`; a folder already has the summary's name). The program log gets "Error creating summary report: [Errno 21] Is a directory: …"; the worker's last log lines are "📊 Creating summary report..." then directly "🎉 Processing complete: 1/1 clients" - **nothing on screen says the summary failed**.

**What the trace shows against what the feature should do.** Bugs (for Section 7):

1. **The "Variations" sheet is always "No file variations found"**, even when the scan found problem files (RAN; cause in B-F9).
2. **"Client Status" leaves failed clients out and marks every listed client "Success"**, whatever happened to its copies and reports (RAN Cases B and C).
3. Report-creation failures appear only as a count on the Summary sheet, never in "Errors" (RAN Case C).
4. **Each organisation report's "Successfully Copied" and "Folders Created" add up the whole run so far**, not that client (RAN: 6/12 for a 5-file client; 25/30 for a 3-file client).
5. **A failed summary is silent on screen** (RAN Case D).
6. "Processing Date" is the folder stamp text "211026 0640" (DDMMYY HHMM), not a date Excel understands (RAN).
7. A second run in the same minute overwrites the summary and the organisation reports without warning (READ processing_controller.py:188, excel_handler.py:339, file_organizer.py:354).
8. The organisation report's name uses the raw client and full state text (RAN); like the client folder name (B-R32), it is not cleaned (`sanitize_filename` is not applied at file_organizer.py:313), while copies and report names are - so one client's outputs carry three spellings of its name (RAN: `_Organization_Report_ClientC Co_Tamil Nadu.txt`, `ClientC Co-TN/`, `ITC_Report_ClientC_Co_TN_…`). A failure to write it is only logged (READ 360-362).
9. "Total Files" on the Summary sheet counts successful copies (25), while "Files found" on screen counted accepted files (27) and the scan listed 28 - three different "file" counts for one run (RAN).

**Older documents against the code.** README.md:182 and 195 name the workbook "GST_Processing_Summary_DDMMYYYY_HHMMSS.xlsx"; the code writes `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` (RAN). README.md:197 ("Client Status: Status of each client") - only successful clients. README.md:199 ("Errors: Detailed error information") - client failures only. README.md:200 ("Variations: Files that don't match naming convention") - always empty. README.md:204 ("_Organization_Report… Processing log for each client") - its statistics are running totals of the batch.

**Origin.** Summary workbook and organisation report edac6a3 (shared - before the split on 30-10-2025); the empty variations list c258047; extra keys in each `summary_data['clients']` entry are written but never read: `itc_report_created` and `sales_report_created` (already in the Tkinter code at the split, gui/handlers/processing_handler.py:364-365@47a1492; read by the c258047 Step 4 list `_populate_extraction_from_results`) lost their only reader in 7e09e5c, which rebuilt that list from the run history instead, and `client_key`, `version_folder` were added in 7e09e5c without any reader (B only (after the split); READ `git show 7e09e5c -- gui/controllers/processing_controller.py` and `git show c258047:gui/pyqt_main_window.py`; MW notes). [Corrected during verification - originally said "since 7e09e5c/b0e7d0e" for all four keys; b0e7d0e did not touch them.] Rules: B-R40, B-R94, B-R95, B-R60.

---

### B-F16: Processing progress and log

**What it does.** During processing, Step 3 shows a progress bar, a status line ("Client i/N: <client> | ETA: <time left>") and a coloured log of each step. Stop asks the run to halt before the next client. At the end a message box gives the count of successful clients, the window jumps to Step 4 and, on Windows only, Explorer opens the run folder.

**Where it lives.**

| Part | file:line | What the code does |
|---|---|---|
| Screen | gui/views/processing_view.py:23-128 | Header "FILE PROCESSING", status label "Ready to process", progress bar 0-100, buttons Start Processing / Stop (red "danger" style) / Clear Log, dark log box. |
| Button states | gui/views/processing_view.py:132-135; gui/pyqt_main_window.py:138-139, 374, 413 | Start disabled and Stop enabled while running. |
| Progress | gui/controllers/processing_controller.py:86-90, 185, 198 → gui/views/processing_view.py:137-140 | Before each client: idx ÷ N × 100 and "Client i/N: <client> \| ETA: <t>" ("Calculating..." for the first; then `ProgressTracker.get_eta`, utils/helpers.py:390-403, format utils/helpers.py:405-416); 95 "Creating summary report..."; 100 "Complete \| Xm Ys". The bar drops fractions (`int`). |
| Log | worker lines processing_controller.py:82, 93, 103, 109, 171, 181, 186, 191, 199, 208-237 → gui/views/processing_view.py:142-157 | Each message inserted as HTML with a colour by level (normal, info, success, warning, error). |
| Stop | gui/views/processing_view.py:97-101 → gui/pyqt_main_window.py:151-152 → processing_controller.py:281-284 → checked at 81-83 | A flag read only before each client. |
| End | gui/pyqt_main_window.py:373-410 | Buttons back (374); status "Processing complete" (375); run history (B-F17, 378-380); settings saved (381 - status "Settings saved" for 3 s); Step 4 list rebuilt from disk (384); tab dots (391-392); box "Processing Complete" with "Successfully processed <ok>/<total> clients" and "Files organized in: <target box at this moment>" (394-398); switch to Step 4 (401); Windows only: Explorer on the run folder (403-410). |
| Error | gui/pyqt_main_window.py:412-415 | Box "Processing Error", status "Processing failed". |
| Dead | gui/views/processing_view.py:159-163 `reset()` | Never called (READ grep). |

**Input → Output.** Worker progress and log messages → bar, status line, log; the run summary → message box, tab switch, Explorer.

**User interaction.** Watch; Stop; Clear Log. Start Processing is a second way to start (same handler as START).

#### B-F16 concrete trace

**Case A - the five-client run of B-F12** (RAN: `scripts/t_window_flow.py` → `work/flow_out.txt`).

1. START → gui/pyqt_main_window.py:364 switches to Step 3 (RAN: tab 2); processing_controller.py:263 → `set_processing_state(True)` (RAN: "Start Processing enabled: False, Stop enabled: True").
2. Loop idx 0: 86 `tracker.update(0, …)`; 87 "Calculating..."; 89 0.0 % → "Client 1/5: Client | ETA: Calculating...". idx 1: `tracker.update(1)` → one client done; `get_eta` = remaining 4 ÷ rate → a fraction of a second → `format_duration` → "0s" → 20.0 % "Client 2/5: ClientA Traders | ETA: 0s". And so on to 80.0 %. Then 185 → 95.0 "Creating summary report..."; 198 → 100.0 "Complete | 0m 0s".
3. Each log message becomes `<span style="color:…">message</span><br>` (processing_view.py:155). The worker's leading "\n" and two or four spaces (e.g. "\n🏢 Processing …", "  ✓ Created folder structure", "    Output: …") are plain spaces in HTML, so they disappear.
4. End: 375 then 381 → status bar ends as "Settings saved"; 394-398 box; 401 → Step 4.

Actual output (RAN):

```text
progress signals: (0.0, 'Client 1/5: Client | ETA: Calculating...'), (20.0, 'Client 2/5: ClientA Traders | ETA: 0s'),
                  (40.0, 'Client 3/5: ClientB Private Limited | ETA: 0s'), (60.0, 'Client 4/5: ClientC Co | ETA: 0s'),
                  (80.0, 'Client 5/5: ClientC Co | ETA: 0s'), (95.0, 'Creating summary report...'), (100.0, 'Complete | 0m 0s')
log as displayed (first client):
🏢 Processing Client - D Exports
✓ Created folder structure
✓ Organized 1 files
📊 Creating ITC report...
Output: ITC_Report_Client_DE_211026_0640.xlsx
✅ Success: ITC_Report_Client_DE_211026_0640.xlsx
💰 Creating Sales report...
Output: Sales_Report_Client_DE_211026_0640.xlsx
✅ Success: Sales_Report_Client_DE_211026_0640.xlsx
✅ Completed Client
… four more clients …
📊 Creating summary report...
✅ Summary report created
🎉 Processing complete: 5/5 clients
message box: 'Processing Complete' / 'Successfully processed 5/5 clients\n\nFiles organized in:\n…/work/flow/Target'
status bar afterwards: 'Settings saved'; current tab: 3 (Step 4)
```

The log still held the earlier dry-run lines above this run (RAN; nothing clears it).

**Case B - Stop pressed while the first client is being processed** (RAN: `scripts/t_f15_f16_f17.py` → `work/f16_out.txt`; the test slowed `organize_files` by 0.4 s so that the click arrives during the first client, as a person's click would).

1. The click arrives while client 1 is copying (Stop enabled at that moment: True). processing_controller.py:284 sets the flag.
2. Client 1 finishes completely (folders, copies, both reports, organisation report). Line 81 at idx 1 sees the flag → 82 logs "Processing stopped by user" (warning) → `break`.
3. 184: one success → the summary is still written; 198 → 100 % "Complete | 0m 0s"; 199 → "🎉 Processing complete: 1/5 clients"; the window treats it as a normal end.

Actual output (RAN): message box "Processing Complete" / "Successfully processed 1/5 clients"; progress bar 100, status "Complete | 0m 0s"; the window moved to Step 4; run history entry `{'total_clients': 5, 'successful_clients': 1, 'failed_clients': 0, …}`; only one client folder created.

**Case C - an error message containing "<"** (RAN, same script; message shaped exactly like processing_controller.py:237): sent `"    ❌ ITC Report Error: '<' not supported between instances of 'str' and 'int'"`, displayed **`"❌ ITC Report Error: '"`** - the rest of the only on-screen record of the error is lost (the full text is still in the program log, processing_controller.py:238).

**What the trace shows against what the feature should do.** Bugs (for Section 7):

1. **Stop does not stop the client in progress, and a stopped run ends exactly like a finished one**: "Processing Complete", "Successfully processed 1/5", 100 %, move to Step 4; the four clients never started are recorded as neither successful nor failed (RAN Case B).
2. **The HTML log loses text and layout**: anything between "<" and ">" vanishes (RAN Case C), indentation and blank lines vanish (RAN Case A).
3. **The log is never cleared** between runs or after a dry run (RAN) and slows the window down as it grows (2,000 lines took about 49 s, RAN by VW).
4. The time-left estimate is per client and says "0s" when clients are quick (RAN); with slow Excel reports it is only as good as the average of the clients so far.
5. "✓ Organized N files" is always shown in the success colour, even when some copies failed (READ processing_controller.py:109; RAN by MW e12).
6. The "Processing complete" status text is replaced at once by "Settings saved" (RAN), which disappears after 3 s.
7. The message box names the **Target box as it is at the end**, not the folder used: editing the box during a run changes the message and the Step 4 list (RAN by MW e14).
8. A disabled Stop button looks enabled (RAN by MW e16: the theme's "danger" style has no disabled look).
9. During a run, Step 2's START and DRY RUN stay active; START is silently refused by the controller after the window has already switched tab and replaced its settings (RAN by VW; READ processing_controller.py:259-261).
10. Closing the window during a run is accepted at once and leaves partial output (RAN by MW e11).
11. Explorer opens the run folder only on Windows and only after the message box is closed (READ 403-410); on other systems nothing happens (no message).

**Older documents against the code.** README.md:160-162 ("Click Start Processing… Monitor progress bar and detailed log… Wait for completion message") is true. README.md:279 ("Processing seems frozen → Large files take time, check log for activity"): the log changes only between steps, and an Excel waiting on a hidden prompt (B-F13 Case B) would leave it unchanged with no way to stop the client in progress.

**Origin.** c258047 (worker messages, screen), 7e09e5c (header, layout), b0e7d0e (move to Step 4 after processing, the tab dots and the Step 4 list rebuilt from disk at the end) (all B only (after the split); READ `git show b0e7d0e -- gui/pyqt_main_window.py`). Rules: B-R113, B-R112, B-R139, B-R137, B-R138.

---

### B-F17: Processing run history and "Last Run" column

**What it does.** After each processing run the program appends a record of the run, and one record per processed client, to `run_history.json` in the per-user folder `~/.gst_organizer`. The Step 2 table's "Last Run" column shows, for each client, the time of its latest record.

**Where it lives.**

| Part | file:line | What the code does |
|---|---|---|
| Record types | models/run_history.py:16-33 (`ClientRunRecord`), 36-42 (`RunHistory`) | Client record fields: key, name, state, run time, mode, files organised, ITC / Sales created, ITC / Sales report path, output folder, PQ refreshed, PQ time, PQ extraction data, errors. |
| Client records made | gui/controllers/processing_controller.py:149-160 | Per successful client: key = the scan key, time = now (`%Y-%m-%d %H:%M:%S`), files organised, the two report flags, output folder = the client folder; report paths and errors are never filled. |
| Run recorded | gui/pyqt_main_window.py:377-380 → models/run_history.py:44-71 | Run entry with its own time stamp (46), five counts (total, successful and failed clients, total files, reports generated) and the processing mode (48-56); client records appended per key (62-67); only the last 50 runs kept (69-71) - client records are never trimmed. |
| Saved / loaded | models/cache_manager.py:36, 66-72; models/run_history.py:128-148 | JSON, indent 2, written in place; loaded once at start (gui/pyqt_main_window.py:43) - unreadable file → empty history (146-148); no type checks (119-126). |
| Shown | gui/views/validation_view.py:381-385 → models/run_history.py:97-107 | "Last Run" = `run_timestamp` of the newest record under the scan key; read only when the table is filled, i.e. at a scan. |

**Input → Output.** The run summary and client records → `~/.gst_organizer/run_history.json` → the "Last Run" column at the next scan.

**User interaction.** None; automatic after processing, visible after the next scan.

#### B-F17 concrete trace

**Case A - the five-client run** (RAN: `scripts/t_window_flow.py` → `work/flow_out.txt`).

1. processing_controller.py:150-160 builds five records; gui/pyqt_main_window.py:378 takes them out of the summary; 379 `add_run`; 380 save.
2. The Step 2 table is not refilled after processing (nothing calls `populate_clients` there); "Last Run" stays empty.
3. RE-SCAN → gui/pyqt_main_window.py:274 → `populate_clients(…, self.run_history)` → line 384 per key → the newest record's time.

Actual output (RAN):

```text
runs: [{'timestamp': '2026-10-21 06:40:29', 'total_clients': 5, 'successful_clients': 5, 'failed_clients': 0,
        'total_files': 25, 'reports_generated': 10, 'processing_mode': 'fresh'}]
client_history keys: Client-DE, ClientA Traders-MH, ClientB Private Limited-GJ, ClientC Co-TAM, ClientC Co-TN
record for ClientA Traders-MH: files_organized 5, itc_report_created true, sales_report_created true,
        itc_report_path '', sales_report_path '', output_folder '…/Annual Statement-211026 0640/ClientA Traders-MH',
        pq_refreshed false, pq_extraction_data {}, errors []
Last Run column right after processing: ['', '', '', '', '']
Last Run column after RE-SCAN: '2026-10-21 06:40:29' for all five clients
```

**Case B - 55 runs of one client** (RAN: `scripts/t_f15_f16_f17.py` → `work/f16_out.txt`): runs kept 50, `total_runs` 55, **55 client records kept** for the client; "Last Run" shows the 55th.

**Case C - a damaged history file** (RAN: `scripts/t_f17_badhistory.py` → `work/f17bad_out.txt`). `run_history.json` holds `{"runs": 5}` (valid JSON, wrong type - e.g. after a hand edit). The window opens normally: models/run_history.py:122 stores 5 as `runs` with no check. The user scans and processes one client; processing writes the folders, copies, both reports and the summary; then gui/pyqt_main_window.py:379 → models/run_history.py:58 `self.runs.append(...)` → `AttributeError: 'int' object has no attribute 'append'`. Nothing catches it inside this window handler, and B has no `sys.excepthook` (B-F1), so **PyQt5 aborts the whole program** (RAN: exit code 134). The completion message never appears, the settings are not saved again (line 381 is never reached), the history file is unchanged, and the window simply vanishes. (A file that is not valid JSON at all is handled differently: a log warning, an empty history, and the old file is overwritten at the next save - RAN by LT §5.2 #29-31, #36.)

**What the trace shows against what the feature should do.** Bugs (for Section 7):

1. **"Last Run" is not refreshed after processing**; it appears only after the next scan (RAN Case A).
2. **Client records are never trimmed** (RAN Case B: 55 kept while runs are capped at 50); with extraction records (B-F25) the file holds client figures indefinitely.
3. `itc_report_path`, `sales_report_path` and `errors` are always empty (RAN Case A).
4. **A wrongly typed value in the file crashes the program at the end of the next processing run** (RAN Case C); an unreadable file is silently replaced (RAN by LT).
5. "Last Run" is keyed by the scan key; extraction results are stored under a different key ("<client folder> / <version folder>"), so "Last Run" never reflects Step 4 (RAN by LT §5.2 #26; see B-F25). The Step 4 key starts with the real client folder name, which differs from the scan key whenever the name was abbreviated or cut (B-F4), so even without the "/ <version folder>" part the two keys would not match (READ: the Step 4 key is built at gui/views/extraction_view.py:384, carried as `client_data['name']` to power_query_extractor/core/report_processor.py:44 and stored at gui/pyqt_main_window.py:555-557; the processing key at gui/controllers/processing_controller.py:151).
6. With two copies of the program open, each saves its own in-memory history over the other's - the last to save wins (READ models/run_history.py:128-136; RAN by the models reader).
7. The run's time stamp (models/run_history.py:46, taken when the run is recorded) and each client's `run_timestamp` (processing_controller.py:154, taken when that client finished) can differ by minutes in a long run (READ); "Last Run" shows the client's.

**Older documents against the code.** README.md does not describe the run history. The commit message of c258047 says "Run history displayed in client tree ("Last Run" column)" - true only after a re-scan.

**Origin.** c258047 (B only (after the split)). Rules: B-R119, B-R120.

---

*End of the first part of Section 3 (B-F1 to B-F17). B-F18 to B-F34 follow in the second part.*
