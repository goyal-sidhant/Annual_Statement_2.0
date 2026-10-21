# B-X-DEAD: Project B dead code, pattern consistency, silent-bug scan, leftovers

**Project B** = goyal-sidhant/AnnualStatement_oldv2, snapshot `b0e7d0e` (09-02-2026), PyQt5 GUI, calls itself "v4.0".
Every `file:line` below is **repository-relative at `b0e7d0e`** (line numbers equal the repository's).
Finding IDs: **D-** dead code, **P-** pattern consistency, **S-** silent bug, **L-** leftovers / commented-out code, **R-** README versus code.

---

## 0. Method, evidence keys, terms

### 0.1 What was done
- **READ**: every one of the 38 `.py` files (7,805 lines), plus `README.md`, `requirements.txt` and `.gitignore`, was read start to end with the Read tool. Every function body, branch and return path was read; nothing was judged from its name.
- **RAN (mechanical candidate lists)**: three small scripts written for this task use Python's built-in `ast` module (Python's own code reader, which turns source code into a tree that a script can search):
  - `/root/merge-scratch/run/B-X-DEAD/dead_scan.py` → `scan_output.txt` (import graph, reachability from the two entry points, defs with no references, never-loaded module names, unused imports, unused parameters, unused local variables, attributes stored but never read, every Qt signal with its emit and connect sites, every `.connect(` target, getattr/QTimer/thread sites, mutable defaults, `==`/`is` with None/True/False, lambdas inside loops, every `except` clause).
  - `def_refs.py` → `def_refs_output.txt` (every class/function/method with **every** place its name is used, so that same-named methods on different classes, such as three different `reset` methods, could be told apart by hand).
  - `magic_scan.py` → `magic_output.txt` (every literal number and every non-message literal string in the logic files).
  - Every candidate was then checked by hand (READ) with `grep` and by reading the call sites.
- **RAN (behaviour tests)**: tests T1-T10 ran on a scratch copy `/root/merge-scratch/run/B-X-DEAD/Bcode` (confirmed byte-identical to B-copy with `diff -rq`), with `HOME` pointed at `/root/merge-scratch/run/B-X-DEAD/home` so that B's settings cache (`~/.gst_organizer/`) landed in scratch. Scripts: `t3_consolidator.py`, `t4_qt_wiring.py`, `t5_window.py`, `t6_suffix.py`; T1, T2, T7, T8, T9 and T10 were run inline (their code is visible in this session's history; their output is quoted below).
- **Test stand-ins ("stubs")**: Linux has no pywin32, so B cannot even be imported here (see S-02). For tests that needed B importable, `/root/merge-scratch/run/B-X-DEAD/stubs/` holds empty stand-in modules named `pythoncom`, `win32com.client`, `win32api`, `win32con` and `win32gui`. They do **not** imitate Excel; any real use raises an error. Each test says whether the stubs were on the path.
- **git**: `/root/merge-scratch/combined`, ref `B/main`, used only for "when did this appear" questions.
- Nothing was written in A-copy, B-copy, the original checkouts or anywhere under `/home/user`. `gst_organizer.log` was not opened.

### 0.2 Evidence keys
- **READ** = established by reading the code. **RAN** = established by executing code (test named).
- **HIGH / MEDIUM / LOW** = confidence of an *inference* (anything beyond what the line literally says).
- **[GST RULE]** = the code decides a tax-related matter (return, table, sheet, cell, period, state code, etc.). For these I only describe what the code does and never judge correctness. **TIME-VARYING** = a value that can change over time (table numbers, return names, state lists, dates).

### 0.3 Plain-English terms used below
- *Module*: one `.py` file. *Package*: a folder of modules with an `__init__.py` file. *Import*: one module loading another.
- *Class*: a blueprint for objects; *instantiate* = make an object from it. *Method*: a function that belongs to a class. *Parameter*: an input a function accepts.
- *Dead code*: code that never runs. *Write-only state*: data that is saved or stored but never read back by anything.
- *Signal / slot* (Qt's messaging system): a *signal* is an announcement ("scan finished"); a *slot* is the function that runs when it is announced; `.connect(...)` links them; `.emit(...)` makes the announcement.
- *Thread*: a separate line of work running at the same time. The *main thread* (also called the GUI thread) draws the window; *worker threads* do slow jobs. *QThread / moveToThread*: Qt's way to run an object's work in a worker thread.
- *COM*: the Windows mechanism Python uses to remote-control Excel. *pywin32*: the Python package that provides COM access (`pythoncom`, `win32com`, `win32api`, `win32con`, `win32gui`).
- *Re-export / `__all__`*: an `__init__.py` that imports names from its sub-modules so others could write `from utils import X`; `__all__` is the list of names offered that way.
- *glob*: a wildcard file search such as `ITC_Report_*.xlsx`. *regex*: a text pattern.
- *Lambda*: a small unnamed function written inline. *Closure / late binding*: a function that remembers a variable from where it was written; if made in a loop, it may see only the loop's last value.
- *Mutable default argument*: a list or dict used as a parameter's default value; it is shared between calls.
- *Dataclass*: a Python class mainly used to hold named fields.
- *openpyxl*: the Python library that reads and writes `.xlsx` files without Excel.
- *Magic number / string*: a literal value written directly into logic instead of being named once as a constant.

---

## 1. Headline findings (most important first)

| # | ID | Finding | Evidence | Confidence |
|---|----|---------|----------|-----------|
| 1 | S-01 | The consolidated extraction report (`PQ_Extraction_Report_*.xlsx`) puts the **wrong cell's value in 9 of its 22 data columns**: the IGST/CGST/SGST columns of the "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" sheets show the values configured for "Diff GSTR-3B"!J15/K15/L15. Cause: a substring match at `power_query_extractor/core/data_consolidator.py:130-133`. | RAN T3 | HIGH |
| 2 | S-02 / P-01 | **The whole application cannot start without pywin32.** `power_query_extractor/core/report_processor.py:11-15` imports five Windows-only modules unguarded; this file is imported at start-up through `gui/controllers/extraction_controller.py:10` ← `gui/pyqt_main_window.py:27`. Both `main.py` and `launch_extractor.py` fail. `requirements.txt:14` has pywin32 commented out, and `main.py:53` advises installing only "PyQt5 openpyxl". Every other Windows-only import in B is guarded. | RAN T1 | HIGH |
| 3 | S-03 | The Extract tab chooses the **"latest" Level-1 folder by sorting folder names** (`gui/pyqt_main_window.py:432-448`). Names carry `DDMMYY HHMM` (`utils/helpers.py:25`), so a December folder outranks a later February one. The re-run logic picks "latest" by modified time instead (`core/file_organizer.py:160-169`), so the two disagree. | RAN T5a | HIGH |
| 4 | S-04 | **Extraction results are saved in run history under keys that are never looked up.** They are stored under `"<client folder> / <version>"` (`report_processor.py:44` → `pyqt_main_window.py:555-557`), while the "Last Run" column looks up the scan key `"<client>-<state code>"` (`validation_view.py:384`). | RAN T5b | HIGH |
| 5 | S-05 | The processing summary workbook's **"Variations" sheet is always "No file variations found"**: `summary_data['variations']` is hard-coded `[]` (`gui/controllers/processing_controller.py:73`) and never filled. This regressed from the pre-split Tkinter code (`47a1492:gui/handlers/processing_handler.py:161` passed the scan's variations). | RAN T8 | HIGH |
| 6 | S-06 | The per-client `_Organization_Report_*.txt` shows **cumulative counts for the whole run so far** under "Successfully Copied" and "Folders Created", not per-client counts (`core/file_organizer.py:346-351`; one organizer per run at `processing_controller.py:46`). | RAN T10 | HIGH |
| 7 | S-07 | **"Resume" mode never resumes.** Each run creates a new minute-stamped `Version-DDMMYY HHMM` folder (`core/file_organizer.py:99-104`), so the "skip files that already exist" branch (`core/file_organizer.py:266-270`) only fires within the same clock minute. When it does fire, the template links receive the **un-sanitised** file name, which does not exist on disk. | RAN T9 | HIGH |
| 8 | S-08 | The Extract tab's configurable **"File suffix" is ignored by every search for refreshed files**: skip-refresh looks only for `_Refreshed_` (`report_processor.py:102`) and the tree's refresh status looks only for `*_Refreshed*` (`pyqt_main_window.py:468`). Separately, `ITC_Report_*.xlsx` also matches refreshed copies, and `report_files[0]` (`report_processor.py:89`) can pick a refreshed copy as the "original". | RAN T6 | HIGH (suffix); MEDIUM (which file Windows picks) |
| 9 | S-09 | A client is counted **"Success" even when its ITC and/or Sales report failed** (`processing_controller.py:135-147` run regardless of `itc_success` / `sales_success`). | RAN T5e | HIGH |
| 10 | D-01 to D-04 | Dead or broken modules: `models/signals.py` (central signal hub with 19 signals, never imported); `power_query_extractor/extractor_main.py` (Tkinter-era launcher that **crashes on import**, because `power_query_extractor/gui/extractor_window.py` was deleted in `a62d3cc`); the empty `power_query_extractor/gui/` package; and `gui/widgets/collapsible_group.py` (loaded at start-up, but its class is never used). | RAN T1d, T2; READ | HIGH |
| 11 | D-25 / D-33 / L-07 | The **per-client "include client name in folder" setting is dead** (no UI writes `processing_state.client_folder_settings`, so `core/file_organizer.py:124-126` always reads False). The Setup tab note at `gui/views/setup_view.py:385` still describes the removed Tkinter feature ("Names >10 chars will show warning ... overrides individual client settings"). | READ (+ git) | HIGH |
| 12 | P-05 | **Constants are largely bypassed.** 8 named constants and several sub-keys in `utils/constants.py` are never read (GUI_CONFIG, MESSAGES, EXCEL_EXTENSIONS, TEMPLATE_EXTENSIONS, MAX_FILE_SIZE_MB, EXCEL_SIGNATURES, EXCEL_SAFETY, FOLDER_STRUCTURE['level2']), while the same values are hard-coded and duplicated in logic files (Excel signatures, the 1 KB minimum size, Excel extension lists, the `Annual Statement-` prefix, the `_Refreshed_{timestamp}` suffix five times, the default length 35 eight times). | READ + RAN scanner | HIGH |

Other confirmed silent issues are listed in section 4: the numeric "Files", "Missing" and "Extra" columns sort as text (RAN T4d); the "Folder Name" column shows a different name from the folder actually created (RAN T5c); the max-length box is greyed out but still used (RAN T5d); the worker thread changes the window's scan data in place (RAN T5e); log lines with ✓/❌ symbols are dropped from the log file on a cp1252 Windows setup (RAN T7, simulated); the EXPORT LIST button does nothing (RAN T4c); and after Stop the app still says "Processing Complete".

---

## 2. Dead code

### 2.1 Module (file) level

Import graph and reachability: RAN `dead_scan.py` sections 1 and 2. "Reachable" means a static import chain leads from `main.py` or `launch_extractor.py` to the module (counting imports inside functions, and counting that importing `a.b` first runs `a/__init__.py`).

| ID | File | Status | Evidence | Confidence |
|----|------|--------|----------|-----------|
| D-01 | `models/signals.py` (lines 1-38; class `AppSignals` line 9) | **Dead.** Imported by nobody, not even `models/__init__.py:1-4`. Added in `c258047` (07-02-2026) as a "Central signal hub - replaces direct self.app.xxx references" (`models/signals.py:10`), but no commit ever imported it (`git log -S AppSignals B/main` → only `c258047`). Each controller defines its own signals instead (see 2.6). | RAN scanner (UNREACHABLE) + READ + git | HIGH |
| D-02 | `power_query_extractor/extractor_main.py` (1-79) | **Dead and broken.** Nobody imports it. Line 9 `from .gui.extractor_window import PowerQueryExtractorApp` points at a file deleted in `a62d3cc` (07-02-2026), so importing it raises `ModuleNotFoundError: No module named 'power_query_extractor.gui.extractor_window'` (RAN T2, stubs on the path, so pywin32 is not the cause). It is a Tkinter-era launcher: `app.mainloop()` at line 66 is a Tkinter call. It has its own `logging.basicConfig` writing `pq_extractor.log` (lines 12-19), and a cache lookup for `gst_organizer_cache.json`, `.cache/app_state.json` and `config/last_session.json` (lines 38-42) with keys `target_folder/target_path/output_folder/last_target` (line 52); none of these files are written by B, which writes `~/.gst_organizer/settings.json` (`models/cache_manager.py:17,35`). Unchanged since the initial commit `edac6a3` (`git log -- power_query_extractor/extractor_main.py` → only `edac6a3`). | RAN T2 + READ + git | HIGH |
| D-03 | `power_query_extractor/gui/__init__.py` (0 lines) | **Dead.** Empty package; its only content (`extractor_window.py`) was deleted in `a62d3cc`; only the dead D-02 refers to it. | RAN scanner + READ | HIGH |
| D-04 | `gui/widgets/collapsible_group.py` (1-42) | **Loaded but unused.** Python *does* load it at start-up, because importing `gui.widgets.colored_section` runs `gui/widgets/__init__.py:4`, which imports it (RAN T1d import trace shows `gui.widgets.collapsible_group` loaded). But the class `CollapsibleGroup` (line 9) is never instantiated: `git grep` at `c258047` and `7e09e5c` shows no view ever used it. `gui/views/setup_view.py:17-82` defines its own `CollapsibleCard` instead (added in `7e09e5c`). | RAN T1d + READ + git | HIGH |
| — | `launch_extractor.py` | **Live** entry point (documented at `README.md:166-169`). It is no longer an "extractor": it opens the normal main window and switches to tab index 3 (`launch_extractor.py:18`). The name is a leftover (see L-03). | READ | HIGH |
| — | `gui/controllers/__init__.py`, `gui/views/__init__.py`, `gui/theme/__init__.py`, `models/__init__.py`, `core/__init__.py`, `utils/__init__.py`, `gui/__init__.py`, `gui/widgets/__init__.py` | **Run, but their re-exports are never used.** No file anywhere uses a package-level import such as `from core import ...` or `from utils import ...` (RAN grep; zero matches). They still run as a side effect whenever a sub-module is imported. One serious side effect: `gui/__init__.py:7` imports `pyqt_main_window`, so importing **any** `gui.*` piece (even one view) drags in `report_processor` and therefore pywin32 (seen in T1; this is why tests needed stubs). See P-07. | RAN grep + T1d | HIGH |

### 2.2 Classes

| ID | Class | Status | Evidence |
|----|-------|--------|----------|
| D-05 | `AppSignals` (`models/signals.py:9`) | Never instantiated; module never imported (D-01). | RAN def_refs (0 references) + READ, HIGH |
| D-06 | `CollapsibleGroup` (`gui/widgets/collapsible_group.py:9`) | Never instantiated (only named in the re-export at `gui/widgets/__init__.py:4,7`). | RAN def_refs + READ, HIGH |
| D-07 | `PowerQueryExtractor` (`power_query_extractor/extractor_main.py:24`) | Used only by `main()` in the same dead module (line 74). | RAN def_refs + READ, HIGH |
| — | `ProcessingState` (`models/app_state.py:78`) | Instantiated (`gui/pyqt_main_window.py:41`), but three of its four fields are never read and its method is dead (D-13; fields in 2.7). Its fourth field, `client_folder_settings`, is read but never written (D-25). | READ, HIGH |
| — | `ScanResult` (`models/app_state.py:50`) | Instantiated and used for its four fields; its three properties and `clear()` are dead (D-12). | READ, HIGH |
| — | All other classes | Instantiated: CacheManager, ClientRunRecord, CollapsibleCard, ColoredSection, DataConsolidator, ExcelHandler, ExtractionController/Worker, ExtractionView, FileOrganizer, FileParser, GSTOrganizerWindow, ProcessingController/Worker, ProcessingView, ProgressTracker, ReportProcessor, RunHistory, ScanController/Worker, SetupView, ThemeManager, ValidationView, AppSettings. | RAN def_refs + READ, HIGH |

### 2.3 Functions and methods never called

Built from RAN `def_refs_output.txt`. Name collisions were resolved by reading each call site. For example, `reset` is defined three times; only `FileParser.reset` has callers (`core/file_parser.py:30,138`).

| ID | Function / method | Evidence | Notes |
|----|-------------------|----------|-------|
| D-08 | `FileOrganizer.cleanup_on_error` `core/file_organizer.py:364-379` | 0 references | Also makes `_is_folder_empty` (`:381-390`) dead, since it is its only caller, and the local `import shutil` at `:374`. Note that the processing worker never cleans up after a failed client. HIGH |
| D-09 | `FileOrganizer.get_summary` `core/file_organizer.py:392-409` | 0 references | The only reader of `FileOrganizer.errors` (so `errors` is write-only; see 2.7). HIGH |
| D-10 | `FileParser.validate_results` `core/file_parser.py:366-399` | 0 references | Contains `critical_types = ['Annual Report', 'GSTR-3B Export']` (`:387`) **[GST RULE]** (which returns count as "critical"; TIME-VARYING return names). It is never applied. HIGH |
| D-11 | `CacheManager.load_session_state` `models/cache_manager.py:82-84` and `get_recent_source_folders` `:86-89` | 0 references | So `session_state.json` is write-only (D-31). HIGH |
| D-12 | `ScanResult.total_files` `models/app_state.py:57-59`, `total_clients` `:61-63`, `complete_clients` `:65-68`, `clear` `:70-74` | No attribute access anywhere; every name match is a dict key string or a local variable (def_refs). `clear` name matches are all `QTextEdit.clear()` or `dict.clear()` calls. | HIGH |
| D-13 | `ProcessingState.reset` `models/app_state.py:85-88` | Not called; `reset` references are only on `FileParser`. | HIGH |
| D-14 | `ProcessingView.reset` `gui/views/processing_view.py:159-163` | Not called. Effect: the progress bar and status text keep the last run's values until the next run sends its first progress update. | HIGH (dead); cosmetic effect MEDIUM |
| D-15 | `ThemeManager.toggle` `gui/theme/theme_manager.py:421-424` | 0 references; the window calls `apply_theme` directly (`gui/pyqt_main_window.py:596`). | HIGH |
| D-16 | `CollapsibleCard.addLayout` `gui/views/setup_view.py:81-82` | The two cards only call `addWidget` (`:485, :503, :508, :515`). Every `.addLayout(` hit is on a `ColoredSection` or a Qt layout. | HIGH |
| D-17 | `ColoredSection.content_layout` `gui/widgets/colored_section.py:67-69`; `CollapsibleGroup.content_layout/addWidget/addLayout/_on_toggled` `gui/widgets/collapsible_group.py:28-42` | 0 references (the latter because the class is unused). | HIGH |
| D-18 | `DataConsolidator._write_section` `power_query_extractor/core/data_consolidator.py:237-339` | 0 references | It is the only user of the module-level imports `FIELD_DISPLAY_NAMES, CELL_MAPPINGS` (`:11`), so those imports are effectively dead too. The `c258047` message says "duplicated header-building code in data_consolidator.py" was fixed; this older section writer stayed. HIGH |
| D-19 | `DataConsolidator._create_details_sheet` `data_consolidator.py:341-399` | 0 references | HIGH |
| D-20 | `ReportProcessor.cleanup` `power_query_extractor/core/report_processor.py:749-751` | 0 references; its body is only a docstring and a comment ("Nothing to cleanup with this approach as we create/destroy Excel each time"). | HIGH |
| D-21 | `utils/helpers.py`: `get_date_only` `:27`, `get_file_info` `:115`, `validate_excel_file` `:210`, `truncate_path` `:287`, `get_unique_path` `:344`, `calculate_file_hash` `:422`, `get_relative_excel_path` `:450` | Only named in the unused re-export `utils/__init__.py:12-35`. `validate_excel_file` is also imported by `core/excel_handler.py:38` but never called there. | HIGH |
| D-22 | `extractor_main.py`: `load_cached_folder` `:36`, `run` `:62`, `main` `:72` | Dead module (D-02). | HIGH |

**Live, for contrast** (each has at least one real caller; READ): in `utils/helpers.py`, `get_timestamp`, `get_safe_timestamp` (only via `create_backup`, `:194`), `ensure_path_exists`, `safe_path_join`, `sanitize_filename`, `safe_copy_file`, `create_backup`, `find_excel_files`, `format_size`, `clean_windows_path`, `validate_client_name`, `ProgressTracker` (+`update`, `get_eta`, `get_elapsed_time`), `format_duration` (only via `get_eta`), `extract_filename_without_extension`, `get_short_path` (only on the Excel-COM path, `core/excel_handler.py:298`), `get_state_code` and `create_client_state_key`. `RunHistory.get_last_run_for_client` is live via `get_last_run_timestamp_for_client` (`models/run_history.py:104`).

### 2.4 Code that can run but does not, with the current configuration or UI ("dead in practice")

| ID | Where | Why it never runs | Evidence |
|----|-------|-------------------|----------|
| D-23 | `ProgressTracker` callback branch `utils/helpers.py:383-384` | The only caller creates `ProgressTracker(total_clients)` with no callback (`gui/controllers/processing_controller.py:56`). The `message` argument passed at `:86` therefore goes nowhere. | READ, HIGH |
| D-24 | `FileOrganizer.organize_files` progress branch `core/file_organizer.py:208-210` | The only caller passes no `progress_callback` (`processing_controller.py:106`). | READ, HIGH |
| D-25 | Per-client folder-name branch `core/file_organizer.py:124-126` | `client_folder_settings` is always `{}`: the window passes `self.processing_state.client_folder_settings` (`gui/pyqt_main_window.py:370`), and nothing ever writes to it (only `models/app_state.py:83` creates the empty dict). Only the global checkbox matters. In the pre-split Tkinter version the dict was filled by a per-client toggle (`47a1492:gui/handlers/client_handler.py:49,181-196,215`). | READ + git, HIGH |
| D-26 | "Resume: skip existing file" `core/file_organizer.py:266-270` and "Fresh: back up existing file" `:261-265` | The destination is inside a new `Version-<DDMMYY HHMM>` folder each run, so it only already exists when a run repeats in the same minute (or, for the backup branch, when two source names sanitise to the same name). | RAN T9, HIGH |
| D-27 | `create_client_state_key` truncate-everything branch `utils/helpers.py:533-535` | The UI limits max length to 15-100 (`gui/views/setup_view.py:375`); state codes are at most 3 characters (mapping values have 2, fallback has ≤3, `utils/helpers.py:504,509`), so `available ≥ 11 > 5`. A loaded value is clamped by the spin box before use (`pyqt_main_window.py:203`). | READ, HIGH |
| D-28 | Calculated columns `data_consolidator.py:118-124` and `_calculate_value` `:176-206`; the description-header branch `:53-65` | All 22 mappings in `power_query_extractor/config/cell_mappings.py:33-242` have `'calculation': None` and none has a `'description'` key, so these branches only run if the author edits the config. The `workbook` parameter of `_calculate_value` is never used inside it. | READ, HIGH |
| D-29 | `ExcelHandler.__del__` `core/excel_handler.py:57-63` | `self.excel_app` is set to `None` at `:55` and never assigned again anywhere (RAN grep), so `__del__` never quits Excel. **The same always-None value makes `:207` `if self.excel_app is None:` always True.** So `excel.Quit()` (`:209`) runs even when `:111` attached to an Excel that was **already open** (`GetActiveObject`), despite the comment at `:206` "Don't quit if Excel was already open". Before quitting, `:118` sets `Visible=False` on that instance. Windows-only; traced by reading. Cross-reference B-X-EXCEL. | READ, HIGH (logic); effect on a real user session MEDIUM |
| D-30 | `ExtractionWorker` fallback `gui/controllers/extraction_controller.py:29` (`level1_folder` falls back to `output_folder`) | The only caller always passes a non-empty Level-1 folder, since `pyqt_main_window.py:512-515` returns early if it is empty. So `output_folder` is in practice unused. | READ, HIGH |
| D-31 | Skip-refresh "use existing" branch `report_processor.py:103-111` | Reached only when existing copies use the default `_Refreshed_` suffix (S-08). | RAN T6, HIGH |

### 2.5 Signals (Qt announcements): who emits, who listens

RAN `dead_scan.py` section 8, then each line checked by hand (the scanner matches by name, so `AppSignals`' names falsely "matched" the controllers' signals of the same name).

| Declared at | Signals | Emitted at | Connected at | Verdict |
|-------------|---------|-----------|--------------|---------|
| `models/signals.py:13-38` (`AppSignals`) | settings_changed, scan_started/progress/completed/error, processing_started/progress/log/completed/error/stopped, extraction_started/progress/log/completed/error, status_message, switch_tab (19) | never | never | **All 19 dead** (D-01). Four of them (settings_changed, processing_stopped, status_message, switch_tab) have no namesake anywhere. HIGH |
| `scan_controller.py:16-18` ScanWorker finished/progress/error | | `:38`, `:29` (from a lambda in the worker thread), `:41` | `:71-73, :76-78` | live |
| `scan_controller.py:47-50` ScanController scan_started/progress/completed/error | | `:63`, `:72` (relay), `:83`, `:86` | `pyqt_main_window.py:128-131` | live |
| `processing_controller.py:21-24` worker progress/log/finished/error | | many (`:82-237`) | `:270-277` | live |
| `processing_controller.py:245-249` processing_started/progress/log/completed/error | | `:263`, `:270-271` (relay), `:287`, `:290` | `pyqt_main_window.py:138-147` | live |
| `extraction_controller.py:18-21` worker progress/log/finished/error | | `:67, :89, :97`, `:46`, `:103`, `:41` | `:156-163` | live |
| `extraction_controller.py:127-131` extraction_started/progress/log/completed/error | | `:146`, `:156-157` (relay), `:173`, `:176` | `pyqt_main_window.py:155-164` | live |
| `setup_view.py:88-89` scan_requested, rescan_requested | | `:437`, `:455` | `pyqt_main_window.py:124-125` | live. **Both go to the same handler** `_on_scan_requested`, so RE-SCAN does exactly what SCAN does. READ, HIGH |
| `validation_view.py:18-19` dry_run_requested, start_processing_requested | | `:184`, `:217` | `pyqt_main_window.py:134-135` | live |
| `processing_view.py:16-17` start_requested, stop_requested | | `:95`, `:101` | `pyqt_main_window.py:150-152` | live |
| `extraction_view.py:19-21` start/stop_extraction_requested, refresh_list_requested | | `:270`, `:276`, `:193` | `pyqt_main_window.py:167-172` | live |

- **Relay pattern check (RAN T4a, T4b):** `button.clicked.connect(self.some_signal.emit)` delivers correctly (`['scan', 'rescan']`). A worker-thread signal relayed through `controller.<signal>.emit` reaches its slots **in the main thread** (T4b: `progress slot ran in main thread: True | completed slot ran in main thread: True`), and the controller's `_cleanup` clears `_thread` and `_worker` afterwards.
- **No QTimer, no `threading.Thread`, no `getattr`/string dispatch** (other than reading exception `args` at `core/excel_handler.py:288` and `report_processor.py:687`) anywhere. RAN scanner section 10. So no method is reached through a hidden route that the reference search could have missed. HIGH
- `closeEvent` (`gui/pyqt_main_window.py:601`) has no explicit callers but is **live**: Qt calls it when the window closes. This was a scanner false positive.

### 2.6 Dead UI elements and UI text referring to removed features

| ID | Where | What | Evidence |
|----|-------|------|----------|
| D-32 | `gui/views/validation_view.py:186-200, 221` | The **"EXPORT LIST" button** is created, styled and placed, but `clicked` is never connected: 0 receivers. Added in `c258047` and never wired. | RAN T4c (`receivers of btn_export.clicked: 0`), HIGH |
| D-33 | `gui/views/setup_view.py:385` | Note "Names >10 chars will show warning. This overrides individual client settings if checked." Neither exists in B: the 10-character warning and the individual settings were in the Tkinter client handler (`47a1492:gui/handlers/client_handler.py:183-190`). | READ + git, HIGH |
| D-34 | `gui/views/setup_view.py:374-377, 382` | The "Max length" box is **disabled** unless "Include client name" is ticked, yet its value is **always** collected (`pyqt_main_window.py:203`) and always used to shorten the Level-2 `<client>-<state>` folder (`core/file_organizer.py:79-83` → `utils/helpers.py:513-537`). | RAN T5d (`spin enabled: False | client_name_max_length collected: 20`), HIGH |
| D-35 | `gui/views/setup_view.py:510-512` vs `:344-346` vs `utils/constants.py:163-171` | Three different descriptions of the same modes: the card says "Re-Run: Overwrite existing" and "Resume: Skip already processed"; the radio buttons say "Add new version folders to existing structure" and "Continue previously interrupted processing". Actual behaviour: see S-07. | READ, HIGH |
| D-36 | `gui/views/validation_view.py:329, 379` | The "Folder Name" column shows the scan key (`"<client>-<state code>"`), not the folder actually created (`create_client_state_key`, which abbreviates Private/Limited and truncates). In the Tkinter version this column was a Yes/No toggle (`47a1492:gui/handlers/client_handler.py:43`). | RAN T5c (`'Client A Private Limited-MH'` vs `'Client A Pvt Ltd-MH'`), HIGH |

### 2.7 Write-only state: values stored or saved but never read

**Module-level constants never read anywhere** (RAN scanner section 3, then grep):

| ID | Constant | Where | Notes |
|----|----------|-------|-------|
| D-37 | `GUI_CONFIG` | `utils/constants.py:126-151` | Title "GST File Organizer v3.0", size `1200x800`, min `(1000, 700)`, fonts, 10 colours. The real window uses 1200x800 min and 1400x900 (`pyqt_main_window.py:72-73`), and the views hard-code their colours. |
| D-38 | `EXCEL_EXTENSIONS`, `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB = 100`, `EXCEL_SIGNATURES` | `utils/constants.py:179-185` | **No maximum file size is checked anywhere.** Extension and signature checks are hard-coded elsewhere (P-05). |
| D-39 | `MESSAGES` | `utils/constants.py:191-211` | Every user message is written inline instead. |
| D-40 | `EXCEL_SAFETY` | `utils/constants.py:217-228` | Copied into `self.safety_config` (`core/excel_handler.py:54`), which is never read. So `max_retries 3`, `retry_delay 1`, `temp_prefix 'temp_gst_'`, `backup_prefix 'backup_'`, `verification_size_min 10240` and `save_options {data_only False, keep_vba False, keep_links False}` have **no effect**. The openpyxl path actually uses `keep_links=True` (`core/excel_handler.py:231`), contradicting the unused `keep_links: False`. The backup file name uses a *suffix* `_backup_` (`utils/helpers.py:195`), not the unused prefix. |
| D-41 | `FOLDER_STRUCTURE['level2'] = '{client}-{state}'` | `utils/constants.py:83` | Never read; the Level-2 name comes from `create_client_state_key`. |
| D-42 | `FILE_PATTERNS[*]['description']`, `['category']`, `['folder']` | `utils/constants.py:17-59` | `category` and `folder` are copied into each parse result (`core/file_parser.py:89-90`) and never read. The destination folder actually comes from a separate hard-coded map (`core/file_organizer.py:288-295`), which **disagrees for Annual Report** (`'folder': 'root'` in constants vs `'version'` in the map; the map wins). `description` duplicates hard-coded texts at `core/file_parser.py:322-332` and `gui/views/setup_view.py:493-498`. |
| D-43 | `PROCESSING_MODES` values (`name`, `description`, `icon`) | `utils/constants.py:157-173` | Only the keys are used, as a membership test (`core/file_organizer.py:58`). |
| D-44 | `COLORS` | `gui/theme/theme_manager.py:10-22` | Never read; the stylesheets repeat the hex values inline. |
| D-45 | `EXTRACTION_OPTIONS['auto_refresh_if_missing']`, `['show_timestamp']` | `power_query_extractor/config/cell_mappings.py:250-251` | Never read. In practice the code always refreshes if no refreshed file is found (`report_processor.py:112-115`) and never adds a timestamp option. |
| D-46 | `__all__`, `__version__`, `__author__` | `core/__init__.py:10-13`, `gui/__init__.py:9-11`, `gui/widgets/__init__.py:7`, `utils/__init__.py:22-38`, `power_query_extractor/__init__.py:3` | Never read (P-07, P-08). |

**Dictionary keys written but never read** (READ + RAN grep of every key):

| ID | Key(s) | Written at | Notes |
|----|--------|-----------|-------|
| D-47 | parse/scan result keys `pattern`, `category`, `folder`, `metadata`, `warnings`, `relative_path`, `extension`, `modified`, `created`, `size_formatted` | `core/file_parser.py:87-102, 110-115, 222-236` | `metadata` holds the period/date/month/year/start_month/end_month captured from file names **[GST RULE]** (tax-period text, TIME-VARYING). It is captured and never used by any later step. `warnings` holds client and state-name validation messages that nobody shows. |
| D-48 | client `status_icon` `'✅'`/`'⚠️'` and client `total_size` | `core/file_parser.py:308-314`, `:275` | Never read (statistics compute total size from `scanned_files`, `:340`). |
| D-49 | file-copy result `sanitized_filename`, `backup` | `core/file_organizer.py:257, 265` | Never read; the summary's File Mapping sheet does not show backups. |
| D-50 | `summary_data['variations']` | `processing_controller.py:73` | Always `[]` → S-05. |
| D-51 | extraction "finished" payload `total`, `processed` | `extraction_controller.py:105-106` | Never read (`pyqt_main_window.py:555` reads only `results`). |
| D-52 | extraction result `original_file`, `refreshed_file`, `timestamp` | `report_processor.py:45, 150-151` | Only persisted inside run history (D-55). |

**Object attributes stored but never read** (RAN scanner section 7 + READ):

| ID | Attribute | Where |
|----|-----------|-------|
| D-53 | `CollapsibleCard._bg`, `_border`, `_text_color` | `gui/views/setup_view.py:23-25` |
| D-54 | `ReportProcessor.excel_instance` | `report_processor.py:25` |
| — | `ExcelHandler.safety_config` | `core/excel_handler.py:54` (D-40) |
| — | `ExcelHandler.temp_dir` | `core/excel_handler.py:52-53`: a folder `gst_excel_temp` is **created in the system temp folder every time an ExcelHandler is made, and never used** (README.md:7,42 claims temporary files are used; R-02). |
| — | `ExcelHandler.excel_app` | `:55`, always None (D-29) |
| — | `FileParser.errors`, `FileOrganizer.errors` | `core/file_parser.py:46,168,256`, `core/file_organizer.py:48,143`: appended to, read only by the dead `get_summary`. |
| — | `ProcessingState.is_processing`, `stop_requested`, `selected_clients` | `models/app_state.py:80-82`: never read; the worker objects keep their own `stop_requested`. `client_folder_settings` (`:83`) is read but never written (D-25). |

**Persisted to disk but never read back** (READ + RAN T5f, which produced `~/.gst_organizer/{run_history.json, session_state.json, settings.json}` in scratch HOME):

| ID | Data | Writer | Reader |
|----|------|--------|--------|
| D-55 | `session_state.json`: `active_tab`, `window_size` (`pyqt_main_window.py:604-607`) and `recent_source_folders` (`models/cache_manager.py:45-52, 121-124`) | written on every save and close | none (D-11). HIGH |
| D-56 | `run_history.json`: `runs` (capped at the last 50, `models/run_history.py:69-71`), `last_run_timestamp`, `total_runs`; every `ClientRunRecord` field except `run_timestamp`. `itc_report_path` and `sales_report_path` (`run_history.py:27-28`) are **never even set** (always `''`). `pq_refreshed`, `pq_refresh_timestamp` and `pq_extraction_data` are written by `add_extraction_result` (`:81-95`). | written after each processing and extraction run | only `run_timestamp` of the last record is read (`run_history.py:102-107` → `validation_view.py:384`). HIGH |
| — | `pq_extraction_data` stores each client's **whole extraction result**, including every extracted cell value (`pyqt_main_window.py:555-557`), in `~/.gst_organizer/run_history.json`. `client_history` is never trimmed (only `runs` is capped). That means client figures accumulate in the user's home folder with no reader and no limit. Technical / confidentiality note; no GST judgement. | | READ, HIGH |

### 2.8 Parameters never used, and parameters never supplied

| ID | Where | Detail | Evidence |
|----|-------|--------|----------|
| D-57 | `data_consolidator.py:176` `_calculate_value(..., workbook)` | `workbook` is never used inside; the caller passes `wb` (`:123`) "for access to other data". | RAN scanner + READ, HIGH |
| — | `report_processor.py:451` `find_button(child_hwnd, unused)` | Unused **on purpose**: the Windows window-enumeration callback must accept two arguments. Not a defect. | READ, HIGH |
| D-58 | `processing_controller.py:42` `_process(self, selected_clients)` | Always called with `self.selected_keys` (`:37`); duplicates state the object already holds. | READ, HIGH |
| D-59 | Optional parameters that no caller ever supplies: `sanitize_filename(max_length=200)` `utils/helpers.py:66`; `ProgressTracker(callback=None)` `:372`; `organize_files(progress_callback=None)` `core/file_organizer.py:183`; `CollapsibleCard(parent=None)` `setup_view.py:20`. Always supplied with the same value: `safe_copy_file(verify=True)` (`core/file_organizer.py:273`). Defaults never used because every caller passes the value: `ExtractionWorker(wait_time=10, suffix_pattern=..., skip_refresh=False, level1_folder=None)` `extraction_controller.py:23-25` and `start_extraction(...)` `:138-140`. | | READ, HIGH |

### 2.9 Imports never used in the importing file

RAN scanner section 4; each checked by reading the file.

| File:line | Unused name(s) |
|-----------|----------------|
| `core/excel_handler.py:6` | `os` |
| `core/excel_handler.py:12` | `List`, `Optional` |
| `core/excel_handler.py:36-40` | `sanitize_filename`, `get_safe_timestamp`, `validate_excel_file` |
| `core/file_organizer.py:8` | `Tuple` |
| `core/file_parser.py:10` | `Optional` |
| `models/app_state.py:7` | `Optional` |
| `utils/constants.py:7` | `Tuple` |
| `gui/pyqt_main_window.py:12` | `QMenuBar` |
| `gui/pyqt_main_window.py:18` | `RunHistory` (the object is obtained from `cache_manager.load_run_history()`) |
| `gui/views/extraction_view.py:10` | `QFrame` |
| `gui/views/processing_view.py:7` | `QFrame` |
| `gui/views/setup_view.py:10` | `QGroupBox` (left over from the `c258047` group-box layout replaced in `7e09e5c`) |
| `gui/widgets/collapsible_group.py:6` | `Qt`, `QPropertyAnimation`, `QParallelAnimationGroup` (an animation was planned and never written) |
| `power_query_extractor/core/data_consolidator.py:6` | `Path` |
| `power_query_extractor/core/data_consolidator.py:11` | `FIELD_DISPLAY_NAMES`, `CELL_MAPPINGS`: used only by the dead `_write_section` (D-18) |
| `power_query_extractor/core/report_processor.py:5` | `os` |
| `power_query_extractor/extractor_main.py:5` | `sys` (dead module) |
| Re-export `__init__.py` files (every imported name) | `gui/controllers/__init__.py:2-4`, `gui/views/__init__.py:2-5`, `gui/theme/__init__.py:2`, `models/__init__.py:2-4`: imported only to re-export, never used (P-07) |

Redundant imports (imported twice): `report_processor.py:13` imports `win32api` at the top and again inside the function at `:710`. `utils/helpers.py:7` imports `re` at the top and again at `:516`.

### 2.10 Local variables assigned but never used

RAN scanner section 6; READ-verified.

| File:line | Variable | Note |
|-----------|----------|------|
| `core/excel_handler.py:550` | `state_code = get_state_code(client_info['state'])` | Never used. It has a side effect: for an unknown state it logs the "Unknown state ... using fallback" warning again, twice per client (ITC and Sales). In T8 one client produced that warning **6 times** (RAN). |
| `core/file_parser.py:282` | loop variable `client_key` | harmless |
| `gui/controllers/processing_controller.py:43` | `start_time = datetime.now()` | Elapsed time comes from the tracker instead (`:196`). |
| `gui/pyqt_main_window.py:337` | loop variable `ft` | harmless |
| `gui/views/setup_view.py:349` | loop variable `key` | The mode strings `'fresh'/'rerun'/'resume'` in the `modes` list (`:343-347`) are never used; the mode is decided later by position, `{0:'fresh',1:'rerun',2:'resume'}` (`:633`). |
| (false positive) `report_processor.py:460` | `clicked` | Correct use of `nonlocal`; it is read at `:469`. |

---

## 3. Pattern consistency (audit method step 8)

### P-01. Windows-only (OS-specific) imports

| File:line | How it imports | Follows "guarded" pattern? |
|-----------|---------------|---------------------------|
| `core/excel_handler.py:20-28` | `if platform.system() == 'Windows': try: import win32com.client, pythoncom ... except ImportError:` → sets `USE_WIN32COM` and falls back to openpyxl (`:83-93`) | **Yes** |
| `utils/helpers.py:476-483` | `win32api` imported inside the function, under a platform check and `try/except Exception` | **Yes** |
| `gui/pyqt_main_window.py:407-410` | `os.startfile` only under `platform.system() == 'Windows'`, inside `try` | **Yes** |
| `power_query_extractor/core/report_processor.py:11-15` | `import pythoncom`, `win32com.client`, `win32api`, `win32con`, `win32gui` at module top, **no guard** | **No (breaks the pattern)** |
| `gui/controllers/extraction_controller.py:10-11` | imports `ReportProcessor` at module top, no guard | **No**, which makes the break fatal |

**Consequence (RAN T1):**
- `python main.py` → `ERROR - Import Error: No module named 'pythoncom'`, then prints "Missing module ... pip install PyQt5 openpyxl" and exits with code 1.
- `python launch_extractor.py` → `Failed to launch: No module named 'pythoncom'`.

So the guarded openpyxl fallback in `core/excel_handler.py` can never actually be reached through the app on a machine without pywin32. That includes **Windows without pywin32**: `requirements.txt:14` has `# pywin32>=306` commented out, while README.md:107 lists pywin32 as required on Windows. HIGH.

`main.py:53` gives incomplete advice: it never names pywin32. READ, HIGH.

### P-02. Error handling

There are 91 `except` clauses in total (RAN scanner section 14): **0 bare `except:`** (this matches the `c258047` commit message "Fixed all 26 bare except: clauses") and **15 whose body is only `pass`**. All 15 silent ones are COM clean-up or column auto-fit: `core/excel_handler.py:62, 200, 210, 216, 534`; `report_processor.py:352, 359, 365, 395, 422, 462, 575, 581, 586`; `data_consolidator.py:412`.

Styles in use (READ):
1. **Log and return a "failure value" (main style).** `utils/helpers.py` (`safe_copy_file`→False, `create_backup`→None, `find_excel_files`→[], `get_file_info`→{'error'}, `calculate_file_hash`→None, `validate_excel_file`→False); `core/excel_handler.py` (`_create_report_win32com`/`_create_report_openpyxl`/`create_summary_report`→False, `prepare_template_data`→partial dict); `core/file_organizer.py` (`_organize_single_file`→result with error, `create_organization_report`→None); `models/cache_manager.py` (`_save_json` swallows, `_load_json`→{}); `models/run_history.py` (`save` swallows, `load`→empty); `report_processor.py` (status dicts); `data_consolidator._calculate_value`→'Error'.
2. **Log and re-raise.** `utils/helpers.py` `ensure_path_exists` `:45-47`, `safe_path_join` `:62-64`; `core/file_parser.py` `scan_folder` `:166-169`; `core/file_organizer.py` `create_client_structure` `:140-144`; `data_consolidator.create_report` `:172-174`; `report_processor._safe_open_workbook` `:732-747`.
3. **Convert to another error.** `core/file_organizer.py:62-65` turns any error into `ValueError`.
4. **Workers**: catch-all → `error` signal (`scan_controller.py:39-41`, `processing_controller.py:38-40`, `extraction_controller.py:39-41`). **GUI**: `QMessageBox` (`pyqt_main_window.py:221-565`).

Inconsistencies (READ):
- `ExcelHandler.create_report_from_template` **raises** `FileNotFoundError` for a missing template (`core/excel_handler.py:78-79`) but **returns False** for every other failure. Two contracts in one function.
- The two Excel modules report failure differently: `ExcelHandler` returns booleans; `ReportProcessor` returns `{'success', 'error', 'validation_message'}` dicts.
- In `report_processor.py` some warnings go **only to the log file**, never to the on-screen log: lines `186, 202, 226, 230, 272, 295, 504, 522, 543, 548` call `logger.warning/error` directly. Other messages use `self._log`, which goes to both. So the user does not see, for example, "Sheet 'X' not found" (`:522`) on screen.
- `ExtractionWorker._save_log_file` catches only `OSError` (`extraction_controller.py:120`). Processing writes no run log file at all (only the per-client `_Organization_Report_*.txt`), whereas extraction writes `Extraction_Log_*.txt`.

### P-03. Logging

- **Main pattern**: `logger = logging.getLogger(__name__)` in 13 modules. These modules have none, because they do not log: main.py, launch_extractor.py, all four views, the theme, both widgets, models/signals.py, models/app_state.py, cell_mappings.py, constants.py (RAN grep).
- **Breaks:** `core/excel_handler.py:26,28` calls the *root* logger (`logging.info/warning`) at import time; `main.py:50,57` does the same; `main.py:51-58` and `launch_extractor.py:23` use `print`.
- **Two entry points, two logging setups:**
  - `main.py:16-23` configures logging to `gst_organizer.log` **relative to the current working folder**, so the log goes wherever the program was started from, not the project folder.
  - `launch_extractor.py` configures **no** logging, so no log file is written when the app is started that way.
  - The dead `extractor_main.py:12-19` has a third setup (`pq_extractor.log`).
  - README.md:289-291 promises both files. (READ, HIGH)
- **Encoding (S-14):** `main.py:20` `logging.FileHandler('gst_organizer.log')` gives no encoding, so Python uses the system default. B logs ✓ ⚠️ ❌ ✅ 📌 💡 📋 through `report_processor._log` → `logger.info` (`report_processor.py:33`) and directly (`:272, :295, :333, :340, :504, :522, :548, :567`). RAN T7 (simulation with a cp1252 file handler): `UnicodeEncodeError: 'charmap' codec can't encode character '✓'`; logging prints "--- Logging error ---" and **drops that line**. Whether a given Windows PC uses cp1252 depends on its locale and on Python's UTF-8 mode (not the default in Python 3.14). MEDIUM that this happens on the author's machine; HIGH that it happens on any cp1252 setup.
- `report_processor._log` (`:30-37`) always logs at INFO, whatever `level` it is given; the level reaches only the screen. READ, HIGH.

### P-04. Threading (QThread + signals versus direct access)

- **Main pattern (followed by all three controllers, identically):** a `QObject` worker is moved to a new `QThread` with `moveToThread`. `started` is connected to `worker.run`, and the worker signals are relayed through the controller's own signals. `finished`/`error` → `thread.quit`, and `thread.finished` → `_cleanup` with `deleteLater`. A guard (`isRunning`) prevents a second concurrent run. See `scan_controller.py:57-94`, `processing_controller.py:256-298`, `extraction_controller.py:138-184`. RAN T4b confirms slots run in the main thread.
- **Workers never touch widgets directly** (READ: every worker → GUI path is a signal emit; `report_processor` → `log_callback` → `ExtractionWorker._log` → `self.log.emit`). HIGH.
- **Breaks:**
  1. `ProcessingWorker` **changes the window's own scan data in place** from the worker thread. `client_data` is the same object as `window.scan_result.client_data` (`pyqt_main_window.py:368` → `processing_controller.py:30`), and `core/file_organizer.py:224-227` rewrites `['name']` and adds `['sanitized_name']`. RAN T5e: before the run `'GSTR3B-Client A Private Limited-Maharashtra-Apr.xlsx'`, after it `'GSTR3B-Client_A_Pvt_Ltd-Maharashtra-Apr.xlsx'`. The change is intentional (comment `:216`), but no lock protects it. The Validation tab's DRY RUN button stays enabled during processing and reads this same data (`pyqt_main_window.py:335-339`). After a run, Dry Run shows the sanitised names. Risk of wrong output: LOW; pattern break: HIGH.
  2. `ExtractionWorker` sets the **module-wide** `EXTRACTION_OPTIONS['skip_refresh']` dict (`extraction_controller.py:54-55`) as a hidden way to pass a setting to `ReportProcessor` (`report_processor.py:97-98`). The other two settings, `wait_time` and `suffix_pattern`, are passed as object attributes (`extraction_controller.py:50-51`). The value stays set after the run. Only one extraction can run at a time, so this is a pattern break rather than a live race. READ, HIGH.
  3. Stop flags (`stop_requested`) are plain attributes set from the main thread (`processing_controller.py:284`, `extraction_controller.py:170`) and read in the worker. That is benign in CPython (no lock needed for a single boolean). The flag is only checked **between clients** (`processing_controller.py:81`, `extraction_controller.py:61`), so a long Excel refresh cannot be interrupted.
  4. **Disk scanning on the GUI thread**: `_populate_extraction_from_disk` (`pyqt_main_window.py:419-504`) walks the target folder with `iterdir`, `glob` and `stat` on the main thread, at start-up (`:190`), after processing (`:384`), after extraction (`:561`) and on "Refresh List". On a slow or network drive the window can freeze. Every other disk-heavy job is threaded. READ; freeze risk MEDIUM.
- **No locks or mutexes exist anywhere** (RAN grep: no `threading`, `Lock`, `QMutex`, `QTimer`).

### P-05. Constants versus hard-coded values

**(a) Constants defined, then bypassed by hard-coded copies:**

| Concern | Named constant (unused or partly used) | Hard-coded copies in logic |
|---------|----------------------------------------|----------------------------|
| Excel signatures | `EXCEL_SIGNATURES` `utils/constants.py:182-185` | `utils/helpers.py:240-241`, `core/file_parser.py:197-198` |
| Minimum Excel size | `EXCEL_SAFETY['verification_size_min'] = 10240` `:222` | `1024` at `utils/helpers.py:232` and `core/file_parser.py:184` (a different value) |
| Excel extensions | `EXCEL_EXTENSIONS` `:179`, `TEMPLATE_EXTENSIONS` `:180` | `utils/helpers.py:228` (5 extensions), `:255` (3 globs), `gui/pyqt_main_window.py:469` (3); file dialog `gui/views/setup_view.py:612,619` (`*.xlsx *.xlsm`); the FILE_PATTERNS regexes accept only `.xls`/`.xlsx` (`\.xlsx?$`, `utils/constants.py:15-55`), so `.xlsm` source files are found by `find_excel_files` but always land in "variations" |
| Level-1 prefix | `FOLDER_STRUCTURE['level1']` `:82` | `'Annual Statement-'` at `core/file_organizer.py:162` and `gui/pyqt_main_window.py:434` |
| Type → folder | `FILE_PATTERNS[*]['folder']` | `core/file_organizer.py:288-295` (disagrees for Annual Report; D-42) |
| File-name pattern texts | `FILE_PATTERNS[*]['description']` | `core/file_parser.py:322-332`, `gui/views/setup_view.py:493-498`, `README.md:126-133` |
| Mode names and descriptions | `PROCESSING_MODES` | `gui/views/setup_view.py:343-347, 510-512, 633` (D-35) |
| Window size, fonts, colours | `GUI_CONFIG`, `COLORS` | `gui/pyqt_main_window.py:72-73`; hex colours repeated in every view and in `pyqt_main_window.py:282-296, 585-589` |
| Messages | `MESSAGES` | inline everywhere |
| Refreshed-file suffix | none | `'_Refreshed_{timestamp}'` at `models/app_state.py:22`, `extraction_controller.py:24, 139`, `gui/views/extraction_view.py:136`, `report_processor.py:28`; searches hard-code `_Refreshed_` at `report_processor.py:102` and `*_Refreshed*` at `pyqt_main_window.py:468` (S-08) |
| Default max length 35 | none | `models/app_state.py:19`, `models/cache_manager.py:116`, `core/file_organizer.py:30, 42`, `utils/helpers.py:513`, `gui/views/setup_view.py:376, 636` |
| Default wait 10 s | none | `models/app_state.py:21`, `extraction_controller.py:23, 138`, `extraction_view.py:128`, `report_processor.py:27` |
| Status words | none | `'Success'/'Failed'/'Skipped'` (`core/file_organizer.py:217-397`, `processing_controller.py:107, 141`); `'Complete'`, `'Missing ...'`, `'Has duplicates'` (`core/file_parser.py:307-313`), re-tested as text at `validation_view.py:388-390, 409, 446` and `models/app_state.py:68` |
| Report-count keys | none | `f'{report_type.lower()}_reports_created'` (`processing_controller.py:227`) must spell exactly `itc_reports_created`/`sales_reports_created` (`:65-66`, `core/excel_handler.py:386-387`) |
| "Private"→"Pvt", "Limited"→"Ltd" | none | two **different** implementations: `utils/helpers.py:78-79` (case-sensitive, matches inside words) and `:517-518` (case-insensitive, whole words only). So an upper-case "PRIVATE LIMITED" is abbreviated in the folder key but not in copied file names. READ, HIGH |
| Tab positions | none | tab order is set by position (`pyqt_main_window.py:90-93`) and referred to by number at `:305-307, 332, 364, 391-392, 401, 551` and `launch_extractor.py:18` |

**(b) Every magic number and string in the logic files.** RAN `magic_scan.py`; the full raw list is in `/root/merge-scratch/run/B-X-DEAD/magic_output.txt`. Plain 0/1 indexes and openpyxl column/row counters are omitted below; f-string file-name pieces were added by reading.

| File | Magic values (line) |
|------|---------------------|
| `main.py` | `'gst_organizer.log'` relative to the current folder (20); log format (18); `'utf-8'` stdout wrapper (13); app name `'GST File Organizer'` (41); version `'4.0'` (42); exit code 1 (54, 59); advice text "pip install PyQt5 openpyxl" (53) |
| `launch_extractor.py` | tab index `3` (18) |
| `core/excel_handler.py` | temp folder `'gst_excel_temp'` (52); `'Excel.Application'` (111, 114); sleeps 0.5 s (125, 180); default sheet `'Links'` (133, 235) and fallback "any sheet containing `'link'`" (145) **[GST RULE adjacent: template contract]**; COM error `'-2147352567'` or the word `'path'` in the error, and path length `> 218` (292-293); default sheet `'Sheet'` (321-322); styles: header font size 12, white `'FFFFFF'`, fill `'366092'` (352-354), title size 16 (358); summary sheet names `'Summary'`, `'Client Status'`, `'File Mapping'`, `'Errors'`, `'Variations'` (371, 403, 432, 460, 493); `'A1'`, `'A1:B1'`, start row 3 (374-394); merge `'A2:E2'` (476, 509); `'Review'` (483), `'Rename file'` (516); column width `+2`, cap 50 (537); log truncations `[:50]` (165), `[-50:]` (300); template data keys and file-type names (553-622) **[GST RULE: which source file's folder and name go into which template link]** |
| `core/file_organizer.py` | default max length 35 (30, 42); `'Annual Statement-'` (162); type→folder map (288-295); report name `_Organization_Report_{client}_{state}.txt` using the **raw** client and full state name (313); `'=' * 50` (318-348); `'%Y-%m-%d %H:%M:%S'` (319); "files starting with `_` don't count" (386) |
| `core/file_parser.py` | minimum size 1024 (184); header 8 bytes (196); signatures (197-198); `'GSTR-3B Export'` may repeat (295) **[GST RULE]** (one file per period assumed; TIME-VARYING return name); status texts and icons (307-314); completeness rounded to 1 decimal (303); suggestion keywords `gstr`/`2b`/`ims`/`3b`/`sales`/`reco`/`annual` and pattern texts (321-334); critical types (387, dead) **[GST RULE]** |
| `utils/helpers.py` | `'%d%m%y %H%M'` (25), `'%d%m%y'` (29), `'%d%m%y_%H%M'` (33); `'Private'→'Pvt'`, `'Limited'→'Ltd'` (78-79, 517-518); invalid characters `<>:"/\|?*[]{}+=!@#$%^,;'"`~` (82); control characters <32 and 127-159 (86); max length 200, reserve 10 (66, 105-106); `'unnamed'` (71, 102); strip `'_. '` (98); extension set of 5 (228); minimum 1024 (232); signatures (240-241); globs `*.xlsx *.xls *.xlsm` (255); `'~$'` (260); size units and 1024 (277-283); truncate 60 and 8 (287, 304, dead); client name 2-100 characters and `[<>:"/\\|?*]` (331-338); unique-path limit 100 (362, dead); `'md5'` and 8192 (423, 433, dead); `'Windows'` (476); state fallback = initials of the first 2 words, else the first 3 characters upper-cased (503-509) **[GST RULE: state-code logic]**; key max 35, minimum 5 (513, 530); backup name `{stem}_backup_{timestamp}{suffix}` (195) |
| `utils/constants.py` (constants file) | `STATE_CODE_MAPPING` (234-275) **[GST RULE: state-code logic; TIME-VARYING list of states/UTs]**; `FILE_PATTERNS` regexes (15-55) **[GST RULE: return types and period formats; TIME-VARYING]**; `EXPECTED_FILE_TYPES` (68-75) **[GST RULE]**; `EXCEL_TEMPLATE_MAPPING` cells `B2/B4/B5/B7/B8/B10/B11` and `B2/B3/B5/B6/B8/B9` on sheet `'Links'` (96-120) (template contract: which file path goes into which cell; treated as **[GST RULE]** when in doubt) |
| `models/app_state.py` | defaults 35 (19), 10 (21), `'_Refreshed_{timestamp}'` (22), `'fresh'` (17) |
| `models/cache_manager.py` | `~/.gst_organizer` (17); `settings.json`, `run_history.json`, `session_state.json` (35-37); keep 5 recent folders (50); old cache `~/.gst_organizer_cache.json` (98); marker `.migrated` (99); old key `include_client_name` (115); default 35 (116); JSON indent 2 (140) |
| `models/run_history.py` | `'%Y-%m-%d %H:%M:%S'` (46, 75); keep the last 50 runs (70-71) |
| `gui/controllers/processing_controller.py` | report names `ITC_Report_{client}_{state}_{ts}` (120), `Sales_Report_...` (128); extension `.xlsx` **always** (220), even for a `.xlsm` template (P-06); `GST_Processing_Summary_{ts}.xlsx` (188); space→underscore in timestamp (114, 187); progress 95 and 100 (185, 198); `'Calculating...'` (87); `'%Y-%m-%d %H:%M:%S'` (154), `'%H:%M:%S'` (176); emoji labels (207) |
| `gui/controllers/extraction_controller.py` | defaults 10 and suffix (23-24, 138-139); progress 95 and 100 (89, 97); `Extraction_Log_{%Y%m%d_%H%M}.txt` (112-113); `'=' * 60` (116); global key `'skip_refresh'` (55) |
| `gui/pyqt_main_window.py` | minimum size 1200x800, starting size 1400x900 (72-73); status message 3000 ms (214); tab indexes (305-401, 551); HTML colours `#0078D4 #107C10 #FF8C00` (282-296); `'Annual Statement-'` (434); globs `ITC_Report_*.xlsx`, `Sales_Report_*.xlsx` (461-462), `*_Refreshed*` (468); extension set (469); `'~$'` (470); `'Pending'`/`'Refreshed'` (472, 479); `'%d/%m/%y %H:%M'` (478); dot icon 12 px and colours (571, 585-589); session keys (605-606); `'=' * 40` (252) |
| `power_query_extractor/core/report_processor.py` | default wait 10 and suffix (27-28); glob `{type}_Report_*.xlsx` (76); `'%d%m%y_%H%M'` (92); `'{timestamp}'` token (93); `_Refreshed_*.xlsx` (102); COM error texts `'Open method of Workbooks class failed'` and `'-2147352567'` (126, 736), truncations `[:100]` (127, 132, 701) and `[:200]` (689); fixed wait 3 s (141); `'Excel.Application'` (174, 487); property retries 5 × 2 s (178-188); `xlMaximized = -4137` (183); open retries 3 × 3 s on `'Call was rejected by callee'` (195-203); load wait 5 × 1 s (209-211); activate retries 3 × 1 s (220-227), extra 1 s (232); `'WScript.Shell'` (216); **SendKeys `'%a'`, `'r'`, `'a'` with 1 s gaps (238-249)**: the Excel ribbon key path for Data ▸ Refresh All (from general knowledge; depends on Excel's UI language and version, MEDIUM); error checks 5 × 1 s (266-275); extended wait `max(10, wait_time)` (282); save wait 2 s (307), close wait 1 s (319); error-window words `microsoft excel/error/problem/cannot/failed/unable/could not/exception/invalid/missing` (377-386) and dialog class names `#32770`, `NUIDialog`, `bosa_sdm_msword` (392); child text longer than 5 characters (420), cap 200 (435); default text `'Power Query error occurred'` (437, 441); button words `ok/close/yes/continue` (457) (substring match, so `'ok'` also matches e.g. "book"; LOW); 0.5 s (473); validation sheet `'info'` (597), cell `'BB2'` (607), `'0'`/`'0.0'` (621), Excel error list (626) (template contract: BB2 is expected to hold the client name, per `:605, :632`); path length 218 (668, 743); `\\?\` and `\\?\UNC\` prefixes (675-679, 719); waits 3 s (698), 2 s (706) |
| `power_query_extractor/core/data_consolidator.py` | header font size 12, white, fill `'0078D4'` (34-35); this **differs** from `excel_handler`'s `'366092'`, so the two generated workbooks use different header colours; section font 14 (36); description merge A:J (62); title `{sheet} Report`, `'A1:C1'`, header row 3 (68-71); `'Client Name'`, `'Successfully Run'` (74); `'✓'`/`'✗'`, `'107C10'`/`'D83B01'` (103-107); default text `'Missing'` (117, 142); number format `'#,##0.00'`, not applied to 0 (149-150); `PQ_Extraction_Report_{%Y%m%d_%H%M}.xlsx` (163-164); calculation names and `'Div/0'`, `'Unknown'`, `'Error'` (190-206); `'Summary'` sheet at position 0 (212); cells A1, A3-A8 (215-230); width cap 50 (415) |
| `power_query_extractor/config/cell_mappings.py` (configuration) | 8 output sheets, 22 input cells: `'Diff GSTR-3B'` J15/K15/L15; `'Diff - GSTR-1'` N1/O1/P1/Q1 and I1; `'Reco'` C37/D37/E37; `'T6 of R9'` K29 ("Table 6J Match"); `'T8 of R9'` G6 ("Table 8D Limits"); `'IMS_Reco'` AA1/AD1/AG1; `'ExtraInPurchase'` O1/P1/Q1; `'ExtraInGSTR2B'` S1/T1/U1 (33-242). **[GST RULE]**: which return table and which cell each figure comes from. **TIME-VARYING**: GSTR-9 table numbers 6J and 8D, return names. Not judged. Field-name rule `f"{sheet_name}_{output_column}"` with spaces → `_` (276). |
| Views (not logic files; listed for completeness) | Setup: max-length spin box 15-100, default 35 (375-376, 636); mode IDs 0/1/2 (354, 633); dialog filter `*.xlsx *.xlsm` (612, 619). Extraction: wait 5-120 s, default 10 (127-128). Validation: 9 column headers (327-330). Every view and the theme repeat hex colours, paddings and font sizes inline. These were read but not itemised line by line (see INCOMPLETE). |

### P-06. File paths and file names

- **pathlib versus os.path:** almost everything uses `pathlib.Path`. The exceptions are `launch_extractor.py:8` (`os.path`) and `main.py:28-29` (Path). `main.py` does `sys.path.insert(0, root)` while `launch_extractor.py` does `sys.path.append(dir)`. With `append`, an installed package with the same name (`utils`, `core`, `models`) could take precedence over B's own folder. LOW in practice, because Python already puts the script's folder first.
- **Folder and file naming is not uniform** (RAN T5c, T9):
  - Level-2 folder `create_client_state_key` keeps spaces and abbreviates whole words ("Client A Pvt Ltd-MH").
  - Category folders use `sanitize_filename(client)` (`core/file_organizer.py:86, 115`), which turns spaces into `_` ("Client_A_Pvt_Ltd").
  - Copied files are sanitised the same way (`:254`).
  - The organisation report name uses the **raw** client and the **full state name** (`:313`).
  - The report names use the sanitised client plus the state **code** (`processing_controller.py:112-113`).
- **Two different "long path" strategies for the same job** (opening a workbook in Excel): `core/excel_handler.py:264-306` tries an 8.3 short path; `report_processor.py:638-747` adds the `\\?\` prefix first, then closes same-named workbooks, then tries the short path, then tries without the prefix, then an open call with full options. Duplicated function, different behaviour. READ, HIGH.
- **Template extension versus output extension:** templates may be `.xlsm` (Setup dialog `setup_view.py:612, 619`), but the output is always named `.xlsx` (`processing_controller.py:220`) and created by copying the template file (`core/excel_handler.py:124, 228`). A macro workbook saved under an `.xlsx` name is normally refused by Excel ("file format or file extension is not valid"; general knowledge, MEDIUM; not run, since there is no Excel here). README.md:120-121 instead says `.xlsx or .xltx`, and `TEMPLATE_EXTENSIONS` (unused) lists `.xlsx/.xltx/.xltm`. Cross-reference B-X-EXCEL.
- `clean_windows_path` (`utils/helpers.py:311-315`) converts `/` to `\` for the template link cells, consistently applied in `prepare_template_data` (`core/excel_handler.py:555-616`). READ, HIGH.

### P-07. `__init__.py` re-exports and `__all__` lists

| Package | Style | Complete? | Used? |
|---------|-------|-----------|-------|
| `core/__init__.py` | re-export 3 classes + `__all__` + `__version__ '3.0.0'` + `__author__` | yes | no package-level importer |
| `utils/__init__.py` | re-export 9 constants + 21 helpers + `__all__` + `__version__ '3.0.0'` | **No.** It omits `STATE_CODE_MAPPING`, `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB`, `EXCEL_SIGNATURES`, `get_short_path`, `get_state_code` and `create_client_state_key`, which are exactly the newest additions (the `# NEW:` block at `utils/helpers.py:470`). | no |
| `gui/__init__.py` | re-export `GSTOrganizerWindow` + `__all__` + `__version__ '4.0.0'` | yes | No. Harmful side effect: importing any `gui.*` sub-module pulls in the whole main window, hence pywin32 (seen in T1d and in why the tests needed stubs). Fixed in `c21f788` after `a62d3cc` left it importing the deleted Tkinter `main_window`. |
| `gui/widgets/__init__.py` | re-export 2 + `__all__` | yes | No. Between `a62d3cc` and `7e09e5c` it still imported the deleted `collapsible_frame` and `title_bar`; that was latent, because nothing imported `gui.widgets` until `7e09e5c` (git, HIGH). |
| `gui/controllers/__init__.py`, `gui/views/__init__.py`, `gui/theme/__init__.py`, `models/__init__.py` | re-export, **no** `__all__` | `models` omits `AppSignals` (dead anyway) | no |
| `power_query_extractor/__init__.py` | `__version__ "1.0.0"` only | — | — |
| `power_query_extractor/config|core|gui/__init__.py` | empty | — | — |

Verdict: three different `__init__` styles, and none of the re-exports is used (RAN grep). The `__all__` lists are stale. HIGH.

### P-08. Version labels

The app says "v4.0": `main.py:2, 42`, `gui/pyqt_main_window.py:68, 71`, `gui/__init__.py:3, 11` (`4.0.0`). But these say "v3.0": `core/__init__.py:2, 12`, `core/excel_handler.py:2`, `core/file_parser.py:2`, `core/file_organizer.py:2`, `utils/__init__.py:2, 38`, `utils/helpers.py:2`, `utils/constants.py:2, 128` (GUI_CONFIG title), and README.md:1, 361 ("v3.0 Production Ready"). `power_query_extractor/__init__.py:3` says 1.0.0. READ, HIGH.

### P-09. Timestamp formats (six formats for one idea)

| Format | Example use |
|--------|-------------|
| `%d%m%y %H%M` | Level-1 and Version folder names (`utils/helpers.py:25` via `core/file_organizer.py:43`) and the processing summary timestamp (`processing_controller.py:44`) |
| `%d%m%y_%H%M` | backups (`utils/helpers.py:33`), refreshed-file suffix (`report_processor.py:92`) |
| `%Y%m%d_%H%M` | `PQ_Extraction_Report_*` (`data_consolidator.py:163`), `Extraction_Log_*` (`extraction_controller.py:112`) |
| `%Y-%m-%d %H:%M:%S` | run history (`models/run_history.py:46, 75`; `processing_controller.py:154`), organisation report (`core/file_organizer.py:319`) |
| `%d/%m/%y %H:%M` | Extract-tab refresh status (`pyqt_main_window.py:478`) |
| `%H:%M:%S` | error rows (`processing_controller.py:176`) |

The day-first formats do not sort by date, and B *does* sort folder names to find the newest one (S-03). READ, HIGH.

### P-10. Duplicate implementations of the same thing

- `_safe_open_workbook` ×2 (P-06).
- `_autofit_columns` ×2 (`core/excel_handler.py:524-538`, `data_consolidator.py:401-417`; near-identical).
- Header-style code ×2 with different colours.
- Coloured-log `color_map` + `insertHtml` ×3 (`setup_view.py:571-585`, `processing_view.py:142-157`, `extraction_view.py:402-416`; identical).
- Collapsible container ×2 (`CollapsibleCard` in `setup_view.py:17-82` in use; `CollapsibleGroup` unused).
- Excel signature check ×2.
- "Latest Level-1 folder" ×2 with **different rules** (name versus modified time; S-03).
- Private/Limited abbreviation ×2 with different rules.
- `get_state_code` called repeatedly for the same client (`core/file_parser.py:261`, `core/file_organizer.py:78` and again inside `create_client_state_key` `utils/helpers.py:521`, `core/excel_handler.py:550` (unused), `processing_controller.py:113` fallback).

READ, HIGH.

### P-11. Two entry points that differ

`main.py` sets UTF-8 stdout (`:12-13`), logging (`:16-23`), app name and version (`:41-42`), and catches ImportError separately (`:49-54`). `launch_extractor.py` does none of these and only switches tab (`:18`). READ, HIGH.

---

## 4. Silent-bug scan (audit method step 9)

For each category: what was looked for, what was found, and whether it is a real bug.

### 4.1 Closure / lambda late binding (very common with Qt `connect` in loops)
- **No `.connect(` call sits inside a loop** (RAN scanner section 9: all 74 connect sites are in constructors, builders or `start_*` methods; the loops in builders, `setup_view.py:349, 482, 500, 510` and `validation_view.py:103`, connect nothing).
- Lambdas or nested functions inside loops: `pyqt_main_window.py:475` (`key=lambda f: f.stat().st_mtime`, used immediately by `max`), `extraction_controller.py:70` (`log_callback`, which uses only `self` and is called synchronously within the same iteration), and the Windows callbacks `report_processor.py:370, 414, 451` (called synchronously by EnumWindows; `nonlocal clicked` used correctly).
- The connected lambdas `pyqt_main_window.py:139, 156` and `extraction_view.py:280` capture no loop variables.
- **Verdict: none found.** RAN + READ, HIGH.

### 4.2 Mutable default arguments
- None (RAN scanner section 11). The dataclasses use `field(default_factory=...)` (`models/app_state.py:52-55, 82-83`; `models/run_history.py:32-33, 39-40`). `FileOrganizer` uses `None` then `or {}` (`core/file_organizer.py:29, 41`).
- **Verdict: none.** HIGH.

### 4.3 Off-by-one in ranges or slices

| Where | What | Real? |
|-------|------|-------|
| `report_processor.py:283-285` | "Extended wait... {i}/{extended_wait}" prints 0, 2, 4, ... and never the final count (other loops print `i+1`) | Cosmetic only. READ, HIGH |
| `processing_controller.py:86-90`, `extraction_controller.py:65` | percent = `idx/total*100` (0 % at the first client) and the ETA uses completed clients | Intended; not a bug. READ |
| `utils/helpers.py:533-535` | truncating the whole key would cut the state code | Unreachable (D-27) |
| `utils/helpers.py:105-106` | length budget subtracts the extension length plus 10 | Intentional margin; not a bug |

**Verdict: no harmful off-by-one found.** HIGH.

### 4.4 String compared where a number or date was meant (or the reverse)

| ID | Where | What happens | Real? |
|----|-------|--------------|-------|
| **S-03** | `pyqt_main_window.py:432-436, 448` | Level-1 folders sorted by **name** (descending) to find the newest. Names use `DDMMYY HHMM`. RAN T5a: with folders "...-311225 1000" (older) and "...-010226 0900" (newer, also the newest by modified time), the Extract tab chose **`Annual Statement-311225 1000`**, while `FileOrganizer(rerun)` would reuse **`Annual Statement-010226 0900`**. The Extract tab then lists and extracts from the older run. Introduced in `b0e7d0e`. | **Real.** HIGH |
| S-10 | `pyqt_main_window.py:457` | Version folders listed sorted by name (descending); `Version-DDMMYY HHMM`, so not date order in the tree | Real, cosmetic (every version is still shown). HIGH |
| S-11 | `validation_view.py:324, 376-378` | Numeric columns "Files", "Missing" and "Extra" are stored as text with sorting enabled. RAN T4d: ascending sort gives `['10', '100', '2', '9']`. | Real, cosmetic. HIGH |
| — | `report_processor.py:621` | BB2 is treated as "0" only if its text is exactly `'0'` or `'0.0'` | Matches what COM returns for numbers (float). LOW concern |

### 4.5 `==` or `is` with None / True / False
- None (RAN scanner section 12 + grep). **Verdict: none.** HIGH.

### 4.6 Variables after a loop holding only the last value; stale accumulators

| ID | Where | What happens | Real? |
|----|-------|--------------|-------|
| **S-06** | `core/file_organizer.py:346-351` (with `self.copied_files` `:47, :214` and `self.created_folders` `:46, :96-178`, one organizer per run `processing_controller.py:46`) | The per-client report's "Successfully Copied" and "Folders Created" are running totals. RAN T10: client 1 `Total Files: 2, Successfully Copied: 2, Folders Created: 6`; client 2 `Total Files: 1, Successfully Copied: 3, Folders Created: 12`. Origin: `edac6a3` (shared commit, so A may share it). | **Real.** HIGH |
| **S-12** | `report_processor.py:77, 89` | `report_files[0]` from `glob("ITC_Report_*.xlsx")`, which **also matches refreshed copies**. RAN T6 on this Linux disk: the raw glob order put `..._Done_020226_1000.xlsx` first, so the code made a refreshed copy *of a refreshed copy* (`..._Done_020226_1000_Done_211026_0358.xlsx`). On Windows NTFS, folder listings come back in name order, where `.` sorts before `_`, so the original would probably come first by luck (general knowledge, MEDIUM; not run on Windows). Network shares may differ. | Real on non-NTFS; MEDIUM on Windows |
| — | `report_processor.py:195-205` | If all 3 open attempts get "Call was rejected by callee", `wb` stays None with no error raised. The code continues to `wb.Save()` (`:305`) → AttributeError → caught at `:332` → status error `"'NoneType' object has no attribute 'Save'"` | Real but contained: misleading message. READ, MEDIUM |
| — | `report_processor.py:72-152` `refresh_status` | Always assigned on every path (skip with existing file / refresh / exception re-raised) | Not a bug. READ, HIGH |
| — | `utils/helpers.py:395-396` `get_eta` | `rate = current / elapsed`; `elapsed` could be 0.0 only if two `time.time()` calls return the same value. A ZeroDivisionError at `processing_controller.py:87` sits **outside** the per-client `try` (`:92`) and would abort the whole run | Theoretical. LOW |

### 4.7 Thread-unsafe shared state

| ID | Where | Real? |
|----|-------|-------|
| S-13 | Worker changes `window.scan_result.client_data` in place (P-04 break 1). RAN T5e. | Pattern break; wrong-output risk LOW |
| — | Module-wide `EXTRACTION_OPTIONS['skip_refresh']` set from the worker (P-04 break 2) | Pattern break; no live race (single run guard). HIGH |
| — | `stop_requested` flags without locks | Benign in CPython. HIGH |

**Verdict: no lock anywhere.** One shared object is changed across threads by design.

### 4.8 Widgets touched from worker threads
- None: workers only emit signals. RAN T4b shows slots run in the main thread; READ of every worker path. **Verdict: none.** HIGH.
- Related, not a thread bug: disk scanning on the GUI thread (P-04 break 4).

### 4.9 Other silent bugs found while checking dead code and patterns ("actual output differs from apparent intent")

| ID | Where | What | Evidence |
|----|-------|------|----------|
| **S-01** | `data_consolidator.py:126-138` | For each configured column the code first searches every extracted field name for one that **contains** the column title (`if mapping['output_column'] in field_name`, `:131`) and takes the first hit, even if that value is None (`:132-133`). Only if that gives None does it try the exact key (`:136-138`). Field names are `"<output sheet>_<column>"` with spaces → `_` (`cell_mappings.py:276`), so titles that contain a space never substring-match and fall back to the exact key correctly. The single-word titles `IGST`, `CGST` and `SGST`, used by "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" (`cell_mappings.py:155-228`), match the first field containing those letters, which is `GSTR-3B_Difference_Diff_IGST/CGST/SGST` (extraction order follows the config). RAN T3 used a unique marker per cell: 9 columns WRONG, e.g. `sheet 'IMS Reco' column 'IGST': configured 'IMS_Reco!AA1' -> report shows 'Diff GSTR-3B!J15'`; the other 13 columns are OK. When "Diff GSTR-3B" values are missing (None, because the sheet was not found), the exact-key fallback happens to give the right value. The substring logic dates from `edac6a3`; the affected mappings were added in `47a1492` (30-10-2025, **before the split**, so A may be affected too; not checked here). The cell choice itself is **[GST RULE]**: not judged; only the mismatch between config and output is reported. | RAN T3, HIGH |
| **S-04** | `report_processor.py:44` (`'client': client_data['name']`, where name = `"<client folder> / <version>"`, `extraction_view.py:384`) → `pyqt_main_window.py:555-557` → `models/run_history.py:73-95` | Extraction data is stored under a key unrelated to the scan key `"<client>-<state code>"` used by the "Last Run" column (`validation_view.py:384`). The `client_key` field the window builds (`pyqt_main_window.py:524`, = folder name) is never used by `ReportProcessor`. Even the folder name would not match the scan key whenever abbreviation or truncation applies (D-36). RAN T5b: keys `['ClientA-XX / Version-TEST']`; lookup of `'ClientA-XX'` → `''`. | RAN, HIGH |
| **S-05** | `processing_controller.py:73` | Variations sheet always empty. RAN T8: scan found `['random name.xlsx']`; summary sheet rows `[['Filename', ...], ['No file variations found', None, ...]]`. Regression from `47a1492`; introduced in `c258047`. | RAN, HIGH |
| **S-07** | `core/file_organizer.py:99-104, 158-179, 259-270` | Resume (and Re-run) reuse the latest Level-1 folder, but always make a **new** Version folder named by the current minute, so nothing from an interrupted run is "continued" or skipped; every file is copied again. RAN T9: run 2 (resume, different minute) → new `Version-010226 0930`, status `['Success']`. Run 3 (resume, same minute) → `['Skipped']`, and the template data then holds **`'GSTR-2B-Reco-Client A Private Limited-Maharashtra-Q1'`** while the file on disk is `GSTR-2B-Reco-Client_A_Pvt_Ltd-Maharashtra-Q1.xlsx`. The name is only updated on `'Success'` (`:217`), so a skipped file keeps its original name and the Links cell would name a file that does not exist. | RAN, HIGH |
| **S-08** | `report_processor.py:102`, `pyqt_main_window.py:468` vs the user setting `extraction_view.py:136` → `report_processor.py:93` | A custom suffix such as `_Done_{timestamp}` is used for the file name but is invisible to both searches. RAN T6: with skip-refresh ON, "No existing refreshed file found - will refresh", and the refresh was still called; tree status glob saw `[]`, so it shows "Pending". The suffix became configurable in `bd176bf` (shared, 30-10-2025) without updating the hard-coded search (existing since `edac6a3`); the tree check was added in `b0e7d0e`. | RAN, HIGH |
| **S-09** | `processing_controller.py:117-147` | `successful_clients += 1` and status `'Success'` regardless of report results. RAN T5e (bad Sales template): `successful_clients: 1 | sales_report_created: False | report_errors: 1 | client status: Success`. The completion box says "Successfully processed 1/1 clients" (`pyqt_main_window.py:394-398`). Report failures are also not added to the summary's Errors sheet (only client-level exceptions are, `:175-180`), and failed clients are absent from the Client Status sheet (only successful ones are appended to `clients`, `:136`). | RAN + READ, HIGH |
| S-14 | `main.py:20` | Log lines containing ✓/❌ etc. are dropped on a cp1252 file encoding (P-03). | RAN T7 (simulated), MEDIUM for the author's PC |
| S-15 | `pyqt_main_window.py:373-401`, `processing_controller.py:81-83, 195-202` | After **Stop**, the worker breaks out of the loop and still emits `finished`. The window then shows "Processing Complete — Successfully processed k/N", switches to the Extract tab and opens the folder. There is no "stopped" state (the `processing_stopped` signal exists only in the dead `AppSignals`). The same applies to extraction (`extraction_controller.py:61-63, 97-107`). | READ, HIGH |
| S-16 | `validation_view.py:202-217` | The Validation tab's START button stays enabled during processing; a second click is silently ignored (`processing_controller.py:259-261` logs a warning only), after `_on_start_processing` has already replaced `self.settings` and switched tab (`pyqt_main_window.py:363-364`). The same applies to RE-SCAN during a scan (`scan_controller.py:59-61`). | READ, LOW |
| S-17 | `setup_view.py:583`, `processing_view.py:155`, `extraction_view.py:414` | Messages are inserted as HTML without escaping, so any text inside `<...>` in an error message (for example a Python `<class ...>` text or a `<Client>` placeholder) disappears from the on-screen log. | READ, LOW (cosmetic) |
| S-18 | `core/excel_handler.py:207-211` | Quits an Excel instance that was already open (D-29). | READ, HIGH (logic) |
| S-19 | `processing_controller.py:85` | `self.client_data[client_key]` is outside the per-client `try` (`:92`), so a missing key would abort the whole run. Keys come from the same scan as the tree, so this is LOW. | READ |

---

## 5. Commented-out code and leftovers

| ID | Where | What | Evidence |
|----|-------|------|----------|
| L-01 | `power_query_extractor/core/report_processor.py:251` | `# Rest of your code remains same...`: a fragment pasted from an AI chat. Code continues below it inside the same `try` (`:214`). | READ, HIGH |
| L-02 | `report_processor.py:209, 239, 244` | `# Increased from 3 to 5`, `# Increased from 0.5` ×2: tuning history kept as comments. | READ |
| L-03 | `launch_extractor.py` (whole) | Name and docstring from the separate Tkinter extractor era; it now opens the main app on tab 4 (rewritten in `c258047`). | READ + git |
| L-04 | `power_query_extractor/extractor_main.py` (whole) | Tkinter-era launcher, broken since `a62d3cc` (D-02). | RAN T2 |
| L-05 | `requirements.txt:13-14` | `# Windows-only: Excel COM automation for Power Query` / `# pywin32>=306`: a dependency that is required for start-up is commented out (P-01). | READ + RAN T1 |
| L-06 | `power_query_extractor/config/cell_mappings.py:229-241` | Commented example config block ("ADD MORE SHEETS HERE"); intentional user guidance, not dead code. `:255-278`: a "Backward compatibility / Build old format" block (`CELL_MAPPINGS`, `FIELD_DISPLAY_NAMES`). `CELL_MAPPINGS` is **live** (it drives what `_extract_data` reads, `report_processor.py:497`); `FIELD_DISPLAY_NAMES` is used only by the dead `_write_section`. | READ |
| L-07 | `gui/views/setup_view.py:385` | UI note describing the removed Tkinter per-client feature (D-33). | READ + git |
| L-08 | `utils/helpers.py:20, 470`; `core/excel_handler.py:386-388`; `core/file_organizer.py:216` | `# UPDATED:`, `# NEW:`, `# NEW` ×3, `# IMPORTANT:` change markers (the `c258047` message says 8 "# ADD THIS" comments were removed; these remained). | READ |
| L-09 | `utils/helpers.py:490` | `# Import here to avoid circular import` for `STATE_CODE_MAPPING`. `utils/constants.py` imports nothing from helpers (`constants.py:6-7`), so there is no circular import today. README.md:322 records the original fix. | READ, MEDIUM (that the reason is obsolete) |
| L-10 | `report_processor.py:51, 61` | `# Default to True for backward compatibility`; every caller now supplies `process_itc`/`process_sales` (`pyqt_main_window.py:526-527`). | READ |
| L-11 | `models/cache_manager.py:3, 93-131` | "Replaces the old gui/utils/cache_manager.py"; `migrate_from_old_cache` imports `~/.gst_organizer_cache.json` once (a one-time migration from the Tkinter version; live, called at `pyqt_main_window.py:47`). | READ |
| L-12 | `core/__init__.py`, `utils/__init__.py`, module docstrings | "v3.0" labels (P-08); `core/__init__.py:13` `__author__ = 'GST File Organizer Team'`. | READ |
| L-13 | `gui/widgets/collapsible_group.py:6` | Unused animation imports: a planned feature that was never built. | READ |
| L-14 | `.gitignore:40-42, 63` | `pq_extractor.log` and `config/last_session.json` relate only to the dead extractor (D-02). `.gst_organizer_cache.json` is the old Tkinter cache, which lives in the home folder, not the repo. Harmless. | READ |
| L-15 | `README.md` (whole) | Describes v3.0 Tkinter (see section 6). | READ + git (`README.md` last changed `d674386`, 29-10-2025, before B's split) |

No `TODO`, `FIXME`, `XXX` or `HACK` markers anywhere (RAN grep). The only commented-out code-like lines are the intentional example at L-06 (RAN grep).

---

## 6. README.md versus B's code (high-value discrepancies per the audit method)

B's `README.md` is unchanged since `d674386` (29-10-2025, before the split), so it describes the v3.0 Tkinter app. (The new project's `/home/user/Annual_Statement_2.0/README.md`, which the author added, contains only the title "# Annual_Statement_2.0" and has no bearing on these notes.)

| ID | README says | Code does |
|----|-------------|-----------|
| R-01 | Folder tree lists `gui/handlers/`, `gui/tabs/`, `gui/utils/` (README.md:82-86) | Deleted in `a62d3cc`; B has `gui/controllers`, `gui/views`, `gui/theme`, `gui/widgets`, `models` |
| R-02 | "All file operations use temporary files and verification steps" (7, 42-46, 260-262) | `temp_dir` is created and never used; `EXCEL_SAFETY` (temp prefix, verification size, retries) is never read (D-40). The only "verification" is a size comparison after copy (`utils/helpers.py:172-178`) and an "exists" check (`core/excel_handler.py:183-189`). |
| R-03 | "Clean saves: No VBA macros or external links preserved" (44, 262) | openpyxl path loads with `keep_links=True` (`core/excel_handler.py:231`); the COM path copies the template as-is (`:124`) |
| R-04 | requirements: openpyxl, pywin32 on Windows (105-108) | `requirements.txt` adds PyQt5 and lxml; pywin32 is commented out (L-05) |
| R-05 | "Review variations tab" (156); Summary has a Variations sheet (200) | No variations tab; the Variations sheet is always empty (S-05) |
| R-06 | `python launch_extractor.py` → "auto-loads target folder from cache ... Click **Refresh & Extract**" (166-176) | Opens the main window on tab 4; the button is "Start Extraction" (`extraction_view.py:254`); the folder comes from settings and a disk scan |
| R-07 | Logs `gst_organizer.log` and `pq_extractor.log` (289-291) | Only `main.py` writes a log, relative to the current folder; `pq_extractor.log` only from the dead module (P-03) |
| R-08 | "Re-Run: Add new versions"; "Resume: Continue interrupted processing" (142-145) | Resume does not resume (S-07) |
| R-09 | Recent fixes mention `file_handler.py`, `processing_handler.py` (320-327) | Those files were deleted in `a62d3cc` |
| R-10 | "Backups created for existing files" (307) | Backup branch effectively unreachable (D-26) |
| R-11 | Templates `.xlsx or .xltx` (120-121) | Dialog offers `.xlsx/.xlsm`; output always `.xlsx` (P-06) |
| R-12 | Python "3.7 or higher" (51); `requirements.txt:2` says the same | Not tested on 3.7; B ran here on 3.14 (tests). LOW relevance. |

---

## 7. [GST RULE] items touched in these notes (described, not judged)

- `power_query_extractor/config/cell_mappings.py:33-242`: sheet and cell sources for every extracted figure, including "Table 6J Match" from `'T6 of R9'!K29` and "Table 8D Limits" from `'T8 of R9'!G6`. TIME-VARYING (GSTR-9 table numbering; return names GSTR-1, GSTR-3B, GSTR-2B, IMS).
- `utils/constants.py:234-275` `STATE_CODE_MAPPING` and the fallback `utils/helpers.py:485-511` (initials of the first two words, or the first 3 letters). TIME-VARYING list of states and union territories. The codes are two-letter abbreviations, not the numeric codes found in a GSTIN; this is stated as a fact, not a judgement.
- `utils/constants.py:13-75` file patterns and expected types; `core/file_parser.py:295` (several GSTR-3B files allowed per client); `core/file_parser.py:387` (critical types, dead).
- `utils/constants.py:96-120` template link cells; `report_processor.py:589-636` (`Info!BB2` must hold the client name for a refresh to count as valid). These are template-layout contracts, tagged here because they decide which cell holds what.
- `core/file_parser.py:93-102` period, date, month and year text captured from file names and never used (D-47).

---

## 8. Scratch files produced (all under `/root/merge-scratch/run/B-X-DEAD/`)

- `dead_scan.py`, `scan_output.txt`
- `def_refs.py`, `def_refs_output.txt`
- `magic_scan.py`, `magic_output.txt`
- `t3_consolidator.py`, `t4_qt_wiring.py`, `t5_window.py`, `t6_suffix.py`
- `Bcode/` (scratch copy of B), `stubs/` (test stand-ins, with `README.txt`), `home/` (scratch HOME; holds a synthetic `~/.gst_organizer/` created by T5), `out/` (synthetic test folders and workbooks)

No venv was created and nothing was installed.

---

## 9. Possible client names

- None found in B's code, comments, configuration or strings at `b0e7d0e`.
- None found in the messages or author fields of B's 5 own commits (`c258047`, `a62d3cc`, `c21f788`, `7e09e5c`, `b0e7d0e`; the author is goyal-sidhant).
- `README.md:128-133` uses the example "ABC Ltd" with "Maharashtra". This is almost certainly a generic example; it is recorded only for completeness.
- The shared pre-split commits were looked at only through their subject lines and the specific diffs cited above. No names seen there.

---

## 10. Questions for the author

1. In the Extract tab, which "Annual Statement-..." folder should count as the latest: the newest by date or time, as the re-run logic assumes, or the one with the highest name? Today the two disagree (S-03).
2. Was "Resume" meant to continue inside the interrupted run's own Version folder? Today every run makes a new minute-stamped Version folder, so nothing is skipped (S-07).
3. Do you still want the per-client "include client name in the folder" toggle the Tkinter version had? B has no control for it, but the Setup tab note still mentions it (D-25, D-33).
4. Do you ever change the "File suffix" on the Extract tab? If yes, "Skip refresh" and the "Refreshed" status do not recognise those files (S-08).
5. Have you noticed that in `PQ_Extraction_Report_*.xlsx` the IGST/CGST/SGST columns of "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" equal the "GSTR-3B Difference" values? Which cells those sheets should read is your rule; the code currently does not read the cells your config lists (S-01).
6. What was the "EXPORT LIST" button on the Validation tab meant to export (D-32)?
7. Do you run B on a PC with pywin32 installed? Without it B does not start at all, and `requirements.txt` has it commented out (P-01).
8. Which template formats do you use: `.xlsx` only, or also `.xlsm` / `.xltx` (P-06)?
9. Should a client whose ITC or Sales report failed still count as "Successful", and should failed clients and report errors appear in the summary workbook (S-09)?
10. Should RE-SCAN do something different from SCAN (2.5)?
11. Is it acceptable that run history keeps every extracted figure for every client in your home folder (`~/.gst_organizer/run_history.json`), with no limit and nothing reading it (D-56)?
12. Can `models/signals.py`, `power_query_extractor/extractor_main.py`, the empty `power_query_extractor/gui/` and `gui/widgets/collapsible_group.py` be treated as discardable leftovers (D-01 to D-04)?

---

## INCOMPLETE

1. **Windows-only behaviour was traced by reading only**: Excel through COM, SendKeys, window enumeration and `os.startfile` (no Windows or Excel here). That covers D-29, S-18, the SendKeys key-path meaning (MEDIUM) and the `.xlsm`-as-`.xlsx` refusal (MEDIUM).
2. **Folder-listing order on Windows NTFS** (S-12) comes from general knowledge, not from a run.
3. **Project A was not checked.** S-01 (consolidator matching: `edac6a3` + `47a1492`) and S-06 (organisation-report counters: `edac6a3`) come from commits shared before the split, so A may have the same bugs. A later agent should verify in A.
4. **Styling literals in the views and the theme** (QSS colours, paddings, font sizes) were read but are summarised by category, not itemised line by line; the magic-value table itemises the logic files only.
5. The GUI was exercised only programmatically and off-screen (no human clicking). Visual issues, such as the check indicator shown on a non-checkable ITC/Sales cell (`extraction_view.py:327-339`), were not visually confirmed.
6. The S-14 encoding effect depends on the author's Windows locale and Python settings, which are unknown.
