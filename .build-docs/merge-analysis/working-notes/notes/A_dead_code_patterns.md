# A-X-DEAD — Project A: dead code, pattern consistency, silent bugs, leftovers

Working notes for the later audit writers. Project A = `goyal-sidhant/AnnualStatement_oldv1`,
snapshot HEAD `bf2f1dd` (27-07-2026), read from `/root/merge-scratch/A-copy` (line numbers =
repository line numbers). History from `/root/merge-scratch/combined`, ref `A/main`.

## 0. Legend, method and limits

**Markers.** READ = I read the code; RAN = I executed something (always in
`/root/merge-scratch/run/A-X-DEAD/`, never in A-copy). Confidence: HIGH (clear from code /
reproduced), MEDIUM (reasonable inference), LOW (speculation). [GST RULE] = the code decides
something the author owns (tax figure, which sheet/cell a figure goes to, state codes, tax
period, ...); for those I only describe, never judge. TIME-VARYING = a value that can change
over time (rates, table numbers, state lists, dates).

**Plain-English glossary (first use).**
- *module* = one `.py` file; *package* = a folder of modules with an `__init__.py` file.
- *import* = one file asking to use another file's code. *re-export* = a package's
  `__init__.py` importing names so they can be imported from the package directly.
  `__all__` = an optional list in a module naming what it "officially" offers.
- *class* = a blueprint for objects; *method* = a function that belongs to a class;
  *instance attribute* = a value stored on an object (`self.x`).
- *parameter* = an input a function accepts. *default* = the value used when the caller
  does not supply one.
- *dead code* = code that is never run in the real app. *test-only* = only the automated
  tests (in `tests/`) use it.
- *AST* = Python's own "parse tree" of a source file; used by my scripts to list every
  definition and reference without running the app.
- *Tk / Tkinter* = the desktop-window library project A uses. *widget* = an on-screen element
  (button, label...). *callback* = a function the window calls later (on click, on a timer).
- *thread* = a second line of work running at the same time as the window; *worker thread* =
  the one doing the slow processing. Tk is only safe to touch from its own (main) thread.
- *closure / lambda / late binding* = a small inline function that remembers variables; "late
  binding" is the bug where it remembers the variable, not its value at that moment.
- *COM* = Windows' way for Python to drive Excel (`pywin32` package). Cannot run here (Linux).
- *Treeview* = Tk's table widget (the client list). *IntVar* = a Tk variable holding a number.
- *glob* = a file-name pattern search such as `ITC_Report_*.xlsx`.

**Method (RAN).** Scripts in `/root/merge-scratch/run/A-X-DEAD/`:
- `dead_scan.py` → `scan_out.txt/.json`: parses all 62 `.py` files (43 non-test) with `ast`,
  lists every module, class, function, method, nested function, module variable, class
  attribute, instance attribute and parameter; counts references (plain name uses, attribute
  uses `x.name`, string literals equal to the name for `getattr`/`hasattr` dispatch, `__all__`
  entries, imports) split non-test vs test.
- `ref_detail.py` → `ref_detail.txt`: for each of the 626 definitions, every reference with
  its *receiver* (`self.app.x`, `self.file_handler.x`, ...) and enclosing function, so methods
  with the same name on different classes can be told apart by hand.
- `dict_keys.py` → `dict_keys.txt` (keys of dictionary constants never read by name),
  `unused_imports.py` → `unused_imports.txt`, `unused_locals.py` → `unused_locals.txt`,
  `magic.py` → `magic.txt` (hardcoded numbers and decision strings in logic files).
- Every candidate below was then checked by hand (READ), including callbacks passed as
  `command=`, `bind(`, `bind_all(`, `after(`, `trace_add(`, `protocol(`, lambdas, `getattr`/
  `hasattr` strings, the `__getattr__` forwarder (extractor_window.py:1051), thread targets
  (processing_handler.py:130, extractor_window.py:879, helpers.py:354), and Windows callback
  APIs (`EnumWindows`/`EnumChildWindows`, report_processor.py:450,475,515).
- I read every line of all 43 non-test `.py` files. Test files were searched (grep), not read
  line by line, except tests/test_extractor_standalone.py (read in full).
- Test suite (RAN, copy of A under xvfb): 287 passed, 1 failed —
  `tests/test_setup_validation.py:66` expects `itc.xltx` from `C:\gst\itc.xltx`; on Linux
  `Path(...).name` does not split on `\`, so this one fails only off-Windows (HIGH).
- Environment note (not a code finding): A-copy already contains `__pycache__/` folders with
  Python 3.14 compiled files dated 21-10-2026 03:21 — written by some earlier process after the
  copy was made; they are git-ignored and are not part of the snapshot. I did not create or
  remove them.

---

## 1. Dead code

### 1.1 Modules never imported — none dead (READ + RAN `dead_scan.py`)

| Module | Imported by (non-test) | Note |
|---|---|---|
| `main.py` | — | Entry point of the main app (`python main.py`). |
| `launch_extractor.py` | — | Entry point of the standalone extractor. |
| `power_query_extractor/extractor_main.py` | only `launch_extractor.py:14` | Live only on the standalone path. Its own `if __name__ == "__main__"` block (extractor_main.py:81-82) cannot run as a script: RAN `python power_query_extractor/extractor_main.py` → `ImportError: attempted relative import with no known parent package` (relative import at line 9). Works only as `python -m power_query_extractor.extractor_main`. HIGH. |
| every other module | ≥1 importer | See `scan_out.txt` module table. `gui/tabs/extract_tab.py` deliberately imports `extractor_window` and not `extractor_main` (extract_tab.py:8-10) because `extractor_main` runs `logging.basicConfig` at import (extractor_main.py:12-19). |
| package `__init__.py` files | run as a side effect of importing any submodule | Their re-exported *names* are never used by non-test code — see §2.7. |

No `gui/app.py` exists at HEAD and nothing imports `gui.app` (grep for `gui.app`, `from .app`,
`gui/app` across all `.py` files: no match). History of the compiled leftover in §4.

### 1.2 Classes never instantiated or subclassed — none (READ)
All 27 classes are created somewhere: e.g. FileParser (main_window.py:115), FileOrganizer (via
`self.app.FileOrganizer(...)`, processing_handler.py:189, class stored on the app at
main_window.py:110), ExcelHandler (main_window.py:116), ExtractorPanel (extract_tab.py:74 and
extractor_window.py:1048), PowerQueryExtractorApp (extractor_main.py:68), ProgressTracker
(processing_handler.py:203), Requirement/Result (requirements_check.py:38-60, 74).
`UIHelpers` (gui/utils/ui_helpers.py:8) is never instantiated — it is used only as a
namespace for static methods (ui_helpers.py:11-63); that is a style, not dead code.

### 1.3 Functions and methods never called anywhere (READ, after RAN scan) — HIGH unless noted

| # | Where | What | How searched / why dead | Tests use it? |
|---|---|---|---|---|
| D1 | core/file_organizer.py:381-396 | `FileOrganizer.cleanup_on_error` | name appears only at its definition (ref_detail); no `getattr` string. So on a per-client failure nothing removes half-created folders. | no |
| D2 | core/file_organizer.py:398-407 | `FileOrganizer._is_folder_empty` | only caller is D1 (file_organizer.py:390) → dead by cascade | no |
| D3 | core/file_organizer.py:409-426 | `FileOrganizer.get_summary` | no caller | no |
| D4 | core/file_parser.py:338-371 | `FileParser.validate_results` (contains `critical_types = ['Annual Report', 'GSTR-3B Export']` at :359 — [GST RULE], described only) | no caller | no |
| D5 | gui/main_window.py:244-245 | `GSTOrganizerApp.toggle_client_selection` (delegator) | ClientHandler calls its OWN `self.toggle_client_selection` (client_handler.py:75, 85); nothing calls `app.toggle_client_selection` | no |
| D6 | gui/main_window.py:256-257 | `GSTOrganizerApp.toggle_folder_name_setting` (delegator) | ClientHandler calls its own (client_handler.py:79) | no |
| D7 | gui/main_window.py:275-276 | `GSTOrganizerApp.process_files_thread` (delegator) | the thread target is the handler's own method (processing_handler.py:130) | no |
| D8 | gui/utils/ui_helpers.py:11-14 | `UIHelpers.update_status` | every `update_status(...)` call goes to `GSTOrganizerApp.update_status` (main_window.py:279) → `StatusBar.update_status` (status_bar.py:35); `UIHelpers.update_status` is never referenced (grep `UIHelpers.update_status`: none) | no |
| D9 | gui/utils/cache_manager.py:50-52 | `CacheManager.get_cached_value` | no caller | no |
| D10 | gui/utils/cache_manager.py:54-56 | `CacheManager.set_cached_value` | no caller | no |
| D11 | power_query_extractor/core/data_consolidator.py:250-352 | `DataConsolidator._write_section` | no caller (older report layout) | no |
| D12 | power_query_extractor/core/data_consolidator.py:354-414 | `DataConsolidator._create_details_sheet` | no caller | no |
| D13 | utils/helpers.py:29-31 | `get_date_only` | only re-exported in utils/__init__.py:13,29; nobody imports from the `utils` package root (§2.7) | no |
| D14 | utils/helpers.py:451-470 | `get_unique_path` | re-export only | no |
| D15 | utils/helpers.py:529-547 | `calculate_file_hash` | re-export only | no |
| D16 | utils/helpers.py:557-574 | `get_relative_excel_path` | re-export only | no |

Not dead although they look it: `ReportProcessor.cleanup` (report_processor.py:801-804) IS
called (extractor_window.py:966) but its body is only `pass` — a do-nothing call (HIGH).
`ExcelHandler.__del__` (excel_handler.py:58-64) runs automatically when the object is
destroyed, but its body can never do anything because `self.excel_app` is always `None` (§3.6).
`main()` in extractor_main.py:75 is reached only via `python -m` (see 1.1).

### 1.4 Test-only items (used by tests, not by the app) — READ + grep of tests/

| Where | Item | Non-test use | Test use |
|---|---|---|---|
| utils/helpers.py:368-373 | `find_excel_files` | re-export only (utils/__init__.py:16,32) | tests/test_file_ops.py:10,58,69,73,91 |
| utils/helpers.py:240-260 | `validate_excel_file` | imported by core/excel_handler.py:38 but never called there; core/file_parser.py:190-193 comment says scan_excel_files "already ran validate_excel_file()" — it does not call it; it calls the two halves `_passes_metadata_checks`/`_has_excel_signature` directly (helpers.py:343-357). Same checks, different call. | tests/test_file_ops.py (10 uses), tests/conftest.py:20 (docstring) |
| power_query_extractor/core/refresh_naming.py:14 | `DEFAULT_SUFFIX_PATTERN` | imported by report_processor.py:38 but unused there; the same text is hardcoded at report_processor.py:51 and extractor_window.py:297 | tests/test_refresh_naming.py |
| power_query_extractor/gui/extractor_window.py:1051-1063 | `PowerQueryExtractorApp.__getattr__` (forwards unknown attributes to the panel) | no non-test code reads a panel attribute through the app (extractor_main.py:68-69 only calls `app.mainloop()`, a Tk method) | tests/test_extractor_standalone.py:39-47 |
| gui/utils/setup_validation.py:36-44 | `_default_exists` | only used when no `exists` checker is passed (setup_validation.py:79); SetupTab always passes one (setup_tab.py:540) | tests/test_setup_validation.py |
| gui/utils/dark_mode_manager.py:278-279 | fallback `is_dark_mode = self.is_dark_mode` | the only caller always passes the flag (main_window.py:178) | tests pass it explicitly too (test_dark_mode.py:153,155) → fallback unused everywhere |

Note on D-items and tests: none of D1-D16 is referenced by any test (grep of tests/ for each
name: no match).

Technical side note (LOW impact): `PowerQueryExtractorApp.__getattr__` replaces Tk's own
`Tk.__getattr__`, which forwards unknown attributes to the Tcl interpreter object (RAN: Python
3.14 `tkinter.Tk.__getattr__` returns `getattr(self.tk, attr)`). Code that relied on that
(e.g. `app.eval(...)`) would now get an AttributeError. Nothing in A does that.

### 1.5 Unreachable code and branches that cannot run (READ)

| Where | What | Confidence |
|---|---|---|
| power_query_extractor/core/report_processor.py:383-386 | `return {'success': True, 'error': None}` after a try/except whose try ends in `return` (370) and whose except returns (378-381) → can never run | HIGH |
| core/excel_handler.py:60-64 | `__del__` body: `self.excel_app` is never set to anything but `None` (only assignment excel_handler.py:56; grep `excel_app`) | HIGH |
| core/excel_handler.py:242 | `if self.excel_app is None:` is always true → always `excel.Quit()` (see bug B1, §3.6) | HIGH |
| power_query_extractor/core/data_consolidator.py:53-65 | branch for a sheet `'description'` — no sheet in EXTRACTION_CONFIG (cell_mappings.py:33-242) has that key | HIGH (config-dependent) |
| power_query_extractor/core/data_consolidator.py:130-136 and 188-218 (`_calculate_value`) | every mapping has `'calculation': None` (cell_mappings.py), so the calculation branch never runs with the shipped config. [GST RULE]: `_calculate_value` computes subtract/add/multiply/divide of two cells, returns the text 'Error', 'Div/0' or 'Unknown' into the figure cell on failure (201-202, 212, 214, 218) — described only. | HIGH (config-dependent) |
| power_query_extractor/core/data_consolidator.py:73-83 | first header-writing block (headers 'Client Name','Successfully Run', `current_row`) is immediately redone by 85-101; lines 74, 78-79, 83 are dead writes; comments `# ... rest of the code` (75, 80) are paste leftovers | HIGH |
| power_query_extractor/config/cell_mappings.py:270 | `if mapping['input_sheet'] and mapping['input_cell']` — always true for the shipped config | HIGH (config-dependent) |
| gui/tabs/extract_tab.py:67-69, 84-105 | the "Power Query refresh is unavailable — needs pywin32" notice. RAN (no pywin32 here): `ExtractorPanel` imports fine (`EXTRACTOR_IMPORT_ERROR = None`) because report_processor.py:20-36 already guards pywin32. The notice therefore never shows for the pywin32 case it describes; it would only show for some other import failure, and would then wrongly blame pywin32. With pywin32 missing the panel builds and each report fails with "pywin32 is not installed..." (report_processor.py:66-77, RAN). | HIGH |
| core/file_parser.py:288-306 | `_suggest_pattern` builds a list but the `if/elif` chain adds at most one entry, so `" or ".join(...)` (306) never joins two | HIGH, harmless |
| core/excel_handler.py:617 | `state_code` computed in `prepare_template_data` but used only in a debug log line | HIGH, harmless |

### 1.6 Module-level constants / variables never referenced (READ + RAN scan)

| Where | Name | Status |
|---|---|---|
| utils/constants.py:197 | `TEMPLATE_EXTENSIONS = {'.xlsx','.xltx','.xltm'}` | never referenced. The real template check uses a different set, setup_validation.py:21 `{'.xlsx','.xltx','.xltm','.xlsm'}` |
| utils/constants.py:198 | `MAX_FILE_SIZE_MB = 100` | never referenced (no size limit is enforced) |
| utils/constants.py:199-202 | `EXCEL_SIGNATURES` | never referenced; helpers.py:264-265 re-hardcodes the same bytes |
| utils/constants.py:234-245 | `EXCEL_SAFETY` (max_retries 3, retry_delay 1, temp_prefix 'temp_gst_', backup_prefix 'backup_', verification_size_min 10240, save_options) | only stored into `ExcelHandler.safety_config` (excel_handler.py:55), which is never read → effectively dead, every key unused |
| utils/constants.py:196 | `EXCEL_EXTENSIONS` | only re-exported (utils/__init__.py:9,26); never used. helpers.py:299 and :327 hardcode their own (different) extension lists |
| power_query_extractor/config/cell_mappings.py:261, 278 | `FIELD_DISPLAY_NAMES` | built at import, read only by dead `_write_section` (data_consolidator.py:275) → effectively dead; the import at data_consolidator.py:11 (`FIELD_DISPLAY_NAMES, CELL_MAPPINGS`) serves only that dead method |
| power_query_extractor/config/cell_mappings.py:264-278 | loop variables `sheet_name, config, report_type, mapping, sheet, field_name` | left behind as module globals by the module-level loop; never used (harmless namespace litter) |
| gui/handlers/file_handler.py:10 | `logger` | never used in that file (0 `logger.` calls) |
| gui/main_window.py:38 | `logger` | never used; main_window.py:193 calls `logging.getLogger(__name__)` inline instead |
| core/__init__.py:12, gui/__init__.py:11, utils/__init__.py:38, power_query_extractor/__init__.py:3 | `__version__` ('3.0.0' ×3, '1.0.0') | never read |
| core/__init__.py:13 | `__author__ = 'GST File Organizer Team'` | never read |

### 1.7 Dictionary keys never read (RAN `dict_keys.py`, then READ each)

| Where | Key(s) never read | Note |
|---|---|---|
| utils/constants.py:138-144 | `GUI_CONFIG['fonts']` (all five: default, title, heading, small, code) | every font is hardcoded instead ('Segoe UI', 'Arial', 'Consolas' literals, §2.8) |
| utils/constants.py:153 | `GUI_CONFIG['colors']['white']` | grep `['white']`: none |
| utils/constants.py:83 | `FOLDER_STRUCTURE['level2'] = '{client}-{state}'` | Level-2 name is built by `create_client_state_key` (file_organizer.py:83-97), not this template |
| utils/constants.py:178,183,188 | `PROCESSING_MODES[*]['icon']` | never read |
| utils/constants.py:214-216 | `MESSAGES['errors']` invalid_excel, template_error, processing_error | only no_source / no_templates / no_target / no_selection are used (file_handler.py:123,135,139; processing_handler.py:45,83) |
| utils/constants.py:218-227 | `MESSAGES['success']` (all 3), `MESSAGES['info']` (all 3) | never read |
| utils/constants.py:234-245 | all `EXCEL_SAFETY` keys | see 1.6 |
| power_query_extractor/config/cell_mappings.py:250-251 | `EXTRACTION_OPTIONS['auto_refresh_if_missing']`, `['show_timestamp']` | never read; only skip_refresh and missing_data_text are used |
| gui/tabs/setup_tab.py:56,62 | `LIGHT_SURFACE` / `DARK_SURFACE` keys `'card'`, `'border'` | never read; `CARD_BORDER` (setup_tab.py:69, = light border) is used in both themes, so the dark border colour '#4A4A4A' is never applied (cards keep a light grey edge in dark mode — `_to_dark` never changes `highlightbackground`, dark_mode_manager.py:236-250). LOW, cosmetic |
| gui/utils/dark_mode_manager.py:25-29 | `THEMED_OPTIONS` entries `highlightbackground`, `selectbackground`, `selectforeground` | captured and restored but never darkened (no branch in `_to_dark`, 236-250) — harmless |

### 1.8 Instance attributes written but never read (READ)

| Where | Attribute | Note |
|---|---|---|
| core/excel_handler.py:55 | `ExcelHandler.safety_config` | never read |
| core/excel_handler.py:56 | `ExcelHandler.excel_app` | read (60, 242) but never set to anything except `None` — a stale variable that drives bug B1 |
| power_query_extractor/core/report_processor.py:48 | `ReportProcessor.excel_instance` | never read |
| gui/utils/dark_mode_manager.py:70, 83 | `DarkModeManager._original_root_bg` | written twice, never read (root colours are captured by the generic walk instead, 112) |
| gui/main_window.py:107 | `GSTOrganizerApp.widgets = {}` | never read |
| gui/main_window.py:142, 143 | `self.validation_tab`, `self.processing_tab` | stored, never read (harmless: keeps a reference) |
| power_query_extractor/gui/extractor_window.py:111 | `ExtractorPanel.target_folder` | stored, never read (the parameter is used directly at 133-136) |
| core/file_parser.py:46 (+176, 224) | `FileParser.errors` | appended to, never read by anyone |
| core/file_organizer.py:52 (+147) | `FileOrganizer.errors` | read only in dead `get_summary` (D3) |
| gui/utils/cache_manager.py:16, 23 | `CacheManager.cache_data` | read only in dead D9 (main_window keeps its own copy, main_window.py:121) |
| gui/utils/dark_mode_manager.py:59,106,118 | `DarkModeManager.is_dark_mode` | read only in the unused fallback (1.4) |
| gui/tabs/processing_tab.py:40; extractor_window.py:521 | `progress_bar` widgets; `client_name_check` (setup_tab.py:417) | used only once for layout right after creation (not really dead; listed for completeness) |

### 1.9 Data written into records but never read (READ; grep of each key)

Per-file record built in core/file_parser.py:199-204 and helpers.get_file_info (149-163):
`'relative_path'` (203), `'size_formatted'`, `'modified'`, `'created'`, `'extension'`,
`'is_file'`, `'is_dir'`, `'path'` (helpers.py:151-162), `'category'`, `'folder'`,
`'pattern'`, `'metadata'` (file_parser.py:91-106), `'warnings'` (114-119) — none is read
anywhere. Consequences worth recording:
- `FILE_PATTERNS[*]['folder']` (constants.py:18-60) is carried into every record and never
  used; FileOrganizer uses its own map `type_to_folder` (file_organizer.py:305-312). The two
  disagree for Annual Report: `'root'` (constants.py:58) vs `'version'` (file_organizer.py:311).
  Two sources of truth; the second wins. HIGH.
- The period / month / date / year parts of every file name are captured into `'metadata'`
  (file_parser.py:105-106) and never read: no tax-period or financial-year logic is applied by
  the organizer. [GST RULE]-related observation, described only.
- Client-name / state-name `'warnings'` from `validate_client_name` are never shown.
Other write-only data: op results `'sanitized_filename'` (file_organizer.py:274), `'backup'`
(282); summary rows `'itc_report_created'`, `'sales_report_created'` (processing_handler.py:427-428,
not printed by excel_handler.py:479-491); extractor results `'original_file'`,
`'refreshed_file'` (report_processor.py:194-195), `'timestamp'` (74, 81); `'validation_message'`
(154, 177, 373) is read only by dead D12; `'extracted_values'` read only by the never-run
calculation branch (data_consolidator.py:133); cache key `'recent_folders'`
(cache_manager.py:34-43) is written and only re-read by the writer itself.

### 1.10 Parameters accepted but never used (RAN scan + READ)

| Where | Parameter | Note |
|---|---|---|
| power_query_extractor/core/data_consolidator.py:188 | `_calculate_value(..., workbook)` | never used (caller passes `wb`, 135) |
| power_query_extractor/core/report_processor.py:500 | `find_button(child_hwnd, unused)` | named `unused`; required by the Windows callback signature — harmless |
| client_handler.py:81, 88, 100; main_window.py:331; tooltip.py:30, 62; extract_tab.py:121; setup_tab.py:512; lambdas validation_tab.py:118,130,142,190, processing_tab.py:74,87, setup_tab.py:168, extractor_window.py:428 | `event` / `e` / `_args` / `_event` | required by Tk callback signatures — harmless |

Optional parameters that NO non-test caller ever supplies (so only the default is ever used):
`sanitize_filename(max_length=200)` (helpers.py:68; callers processing_handler.py:292, file_organizer.py:90, 271);
`truncate_path(max_length=60)` (helpers.py:394); `Tooltip(delay=450, wraplength=520)`
(tooltip.py:12); `CacheManager(cache_file=None)` (cache_manager.py:14); `PowerQueryExtractor(
target_folder=None)` (extractor_main.py:27; callers launch_extractor.py:17, extractor_main.py:77);
`check_requirements(requirements=...)` (requirements_check.py:73); `offer_install(parent=None)`
(dependency_installer.py:88); `_short(width=46)` (setup_validation.py:67);
`ProgressTracker(callback=None)` (helpers.py:479; caller processing_handler.py:203) — so the
callback branch helpers.py:490-491 never runs and the `message` argument of `update()` is
always thrown away. All HIGH.

### 1.11 Unused imports (RAN `unused_imports.py`, checked by READ)

main.py:11 `os`; core/excel_handler.py:6 `os`, :12 `List`, `Optional`, :37-38
`sanitize_filename`, `get_safe_timestamp`, `validate_excel_file`; core/file_organizer.py:8
`Tuple`; core/file_parser.py:9 `Optional`; gui/main_window.py:7 `os`, :8 `sys`, :11 `Path`,
:20 `MESSAGES`, :21 `get_timestamp`, :25 `CollapsibleFrame`; gui/handlers/client_handler.py:6
`Path`; gui/utils/setup_validation.py:8 `os`, :10 `Dict`; power_query_extractor/extractor_main.py:5
`sys`; power_query_extractor/core/data_consolidator.py:6 `Path`;
power_query_extractor/core/report_processor.py:5 `os`, :38 `DEFAULT_SUFFIX_PATTERN`;
utils/constants.py:7 `Tuple`; utils/dependency_installer.py:18 `Path`.
Also redundant (not flagged by the script): report_processor.py:23 module-level `import
win32api` is shadowed by a local re-import at 762 inside `_safe_open_workbook` (the module-level
one still acts as part of the pywin32 availability check); helpers.py:623 re-imports `re`
locally although it is imported at line 7. In tests: test_dependency_installer.py:10 `pytest`,
test_extractor_panel.py:6 `tk`, :11 `PowerQueryExtractorApp`, test_workflow_mode.py:7 `pytest`.
Also typing `any` (lower-case, the built-in function) is used as a type at constants.py:13
`Dict[str, Dict[str, any]]` — works at run time, but is not the typing `Any`. LOW.

### 1.12 Unused local variables (RAN `unused_locals.py`)
gui/tabs/setup_tab.py:552 `emoji` (unpacked, never used); utils/dependency_installer.py:133
`output` (pip output discarded after logging); unused loop variables (normal style):
file_parser.py:254 `client_key`, processing_handler.py:59 `file_type`, extractor_window.py:849
`name`.

---

## 2. Pattern consistency (audit method step 8)

### 2.1 Error handling — no single project pattern (READ)
Five styles coexist; which one a function uses depends on when it was written:
1. **Status dictionaries** (`{'success': ..., 'error': ...}`): the whole extractor —
   report_processor.py:70-77, 90-104, 116-123, 151-155, 174-178, 199-207, 370-381, 389-393;
   also `FileOrganizer._organize_single_file` result dict with `'status'` =
   'Success'/'Failed'/'Skipped' (file_organizer.py:254-300), `parse_filename` result with
   `'error'` (file_parser.py:132), `get_file_info` returning `{'error': ...}` (helpers.py:146,166).
2. **True/False returns**: helpers.safe_copy_file (193-215), validate_excel_file and its halves
   (255-306); ExcelHandler.create_report_from_template / `_create_report_*` /
   create_summary_report (excel_handler.py:93-102, 221-228, 310-314, 405-409);
   dependency_installer.install returns `(bool, text)` (52-60), offer_install / restart_application
   return bool.
3. **None returns**: helpers.create_backup (222, 234), calculate_file_hash (535, 547),
   FileOrganizer.create_organization_report (379); get_short_path silently returns the input.
4. **Raising**: helpers.ensure_path_exists (49), safe_path_join (57, 66), get_unique_path (470);
   FileParser.scan_folder (153, 155, 177); FileOrganizer._validate_inputs (63, 69),
   create_client_structure (148); ExcelHandler.create_report_from_template raises
   FileNotFoundError for a missing template (84) although it otherwise returns bool;
   both `_safe_open_workbook`s (excel_handler.py:331, 369; report_processor.py:694, 703, 799);
   DataConsolidator.create_report (186).
5. **Swallow and log / silently pass**: 23 bare `except:` clauses (main_window.py:76, 173;
   data_consolidator.py:201, 427; report_processor.py:269, 401, 408, 414, 444, 452, 471, 511,
   624, 630, 635, 767; excel_handler.py:63, 126, 235, 245, 251, 302, 597) plus many
   `except Exception: pass` (setup_tab.py:350-351, 516-517, 567-568; extract_tab.py:111-112,
   118-119; tooltip.py:129-130, 148-150, 157-158; dependency_installer.py:157-158).
   `prepare_template_data` logs and returns a partial dict (excel_handler.py:710-713).
- **User-facing**: message boxes in the GUI layer (file_handler.py:95, 123-149;
  client_handler.py:192-207, 229, 251, 257; processing_handler.py:45, 83, 112, 498, 525;
  extractor_window.py:614, 860, 960, 970; launch_extractor.py:26; dependency_installer.py:116,
  143; main_window.py:347). Core modules never show dialogs (good separation).
- **Error text placed into figure cells**: data_consolidator.py:129, 154 ('Missing'
  placeholder from EXTRACTION_OPTIONS), 201-218 ('Error', 'Div/0', 'Unknown'). [GST RULE] —
  these land in figure columns; described only.
- Inconsistent misleading message: client_handler.py:251-257 shows "Export Complete", then
  `os.startfile` (Windows-only) — off Windows it raises AttributeError and the user then sees
  "Export Error: Failed to export" although the file was written. HIGH (READ).

### 2.2 Logging (READ + RAN count)
- **Main pattern**: module logger `logger = logging.getLogger(__name__)` with `logger.x(...)`
  calls — followed by core/*, utils/helpers.py, utils/dependency_installer.py,
  gui/handlers/processing_handler.py, gui/tabs/extract_tab.py, gui/utils/cache_manager.py,
  gui/utils/dark_mode_manager.py, all of power_query_extractor/*.
- **Breaks it**: main.py uses the root logger (`logging.error/warning`, 99, 105, 115, 131) and
  `print()` (62-132 — the only `print` calls in non-test code; console banner and dependency
  report shown before the window exists); core/excel_handler.py:26, 28 call the root logger at
  import time, before its own module logger is defined (43); main_window.py:193 calls
  `logging.getLogger(__name__)` inline while its module logger (38) is never used;
  report_processor.py:34 inline getLogger before its module logger (41); file_handler.py:10 has a
  logger that is never used. Files with no logging at all: client_handler.py, the four tab files
  except extract_tab, ui_helpers.py, status_bar.py, widgets, setup_validation.py,
  requirements_check.py, excel_com.py, refresh_naming.py.
- **Two log sinks**: GUI log box (`app.log_message`, processing_handler.py ~45 calls) vs log
  file. The processing worker writes almost everything only to the GUI log; the file gets only
  exceptions (406, 524) and the Step-4 hand-over error (174). `ReportProcessor._log`
  (report_processor.py:53-60) always logs to file at INFO, whatever `level` it is given — so
  "❌ ..." errors are filed as INFO.
- **[DEBUG] tags**: 14 GUI log lines still say `[DEBUG]` (processing_handler.py:333-411) —
  commit 064c634 says these were "intentionally left"; core/excel_handler.py:83 still logs
  `[DEBUG]` inside a `logger.error`.
- **Two `logging.basicConfig` calls**: main.py:20-27 and extractor_main.py:12-19, both writing
  a log file in the *current folder* ('gst_organizer.log', 'pq_extractor.log'); the second is a
  no-op if the first already ran (Python ignores basicConfig once handlers exist).
- **Global side effect**: core/excel_handler.py:17 `warnings.filterwarnings('ignore')` runs at
  import and silences ALL Python warnings for the whole process (openpyxl warnings included).
  HIGH (READ); impact MEDIUM (hides useful warnings).

### 2.3 OS-specific imports and calls (READ + RAN)
| Where | What | Guard |
|---|---|---|
| core/excel_handler.py:20-28 | `win32com.client`, `pythoncom` | platform check `platform.system() == 'Windows'` + `try/except ImportError` → flag `USE_WIN32COM`; falls back to openpyxl with a warning (97-102) |
| power_query_extractor/core/report_processor.py:20-36 | `pythoncom, win32com.client, win32api, win32con, win32gui` | `try/except ImportError` only (no platform check); names set to `None`; flag `WIN32_AVAILABLE` checked only in `process_client` (66) — the COM methods themselves assume it is True (fine, they are only reached from there) |
| utils/helpers.py:580-590 | `win32api.GetShortPathName` | platform check + local import inside try; returns original path on failure |
| report_processor.py:761-768 | local `import win32api` | inside try with bare except (redundant with line 23) |
| utils/requirements_check.py:17 | `ON_WINDOWS = sys.platform == 'win32'` | different idiom from the `platform.system()` used elsewhere |
| gui/tabs/extract_tab.py:24-30 | import of the extractor | `except Exception` — a second guard on top of report_processor's (see 1.5: the notice it enables is unreachable for the pywin32 case) |
| gui/handlers/processing_handler.py:507-512 | open output folder | platform check: `os.startfile` (Windows), `os.system('open "..."')` (macOS), `os.system('xdg-open "..."')` (Linux). The path is pasted into a shell command inside quotes; a folder name containing `"` would break the command. LOW |
| gui/handlers/client_handler.py:255; extractor_window.py:973 | `os.startfile` | NO platform check, only inside try (see 2.1 misleading export message) |
| utils/dependency_installer.py:95-99 | tkinter | the only guarded tkinter import (deliberate: the installer must work without Tk) |
| all other GUI files | tkinter | unguarded (main.py checks only openpyxl/pywin32, not tkinter) |
| main.py:17 | `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, ...)` at import | unguarded. RAN: with `sys.stdout = None` (what `pythonw.exe` gives a program with no console) main.py fails at import: `AttributeError: 'NoneType' object has no attribute 'buffer'`. HIGH |
| mouse wheel | `<MouseWheel>` only (setup_tab.py:201, extractor_window.py:478, 565) | Windows/macOS event; Linux sends Button-4/5 — not handled (A targets Windows; LOW) |

### 2.4 Config / constants access, and hardcoded values in logic files (READ + RAN `magic.py`)
**Main pattern**: shared settings live in `utils/constants.py` (FILE_PATTERNS,
EXPECTED_FILE_TYPES, FOLDER_STRUCTURE, EXCEL_TEMPLATE_MAPPING, GUI_CONFIG, WORKFLOW_MODES,
PROCESSING_MODES, MESSAGES, STATE_CODE_MAPPING); the extractor has its own config module
`power_query_extractor/config/cell_mappings.py` (EXTRACTION_CONFIG, EXTRACTION_OPTIONS) and never
imports utils/constants.py. Followed by: file_parser (FILE_PATTERNS, EXPECTED_FILE_TYPES),
file_organizer (FOLDER_STRUCTURE, PROCESSING_MODES), excel_handler (EXCEL_TEMPLATE_MAPPING),
setup_tab (FILE_PATTERNS for the reference card — explicitly "can never drift", setup_tab.py:478-479).
Broken by the hardcoded values below.

**Values that exist in the constants module but are re-hardcoded elsewhere (drift risk)**:
- Excel extensions: constants.py:196 `{'.xlsx','.xls','.xlsm'}` (unused) vs helpers.py:299
  `{'.xlsx','.xls','.xlsm','.xltx','.xltm'}` vs helpers.py:327 `('.xlsx','.xls','.xlsm')` vs
  setup_validation.py:21 `{'.xlsx','.xltx','.xltm','.xlsm'}` vs constants.py:197 (unused) vs the
  Browse dialogs `"*.xltx *.xlsx"` only (file_handler.py:31, 42). So a `.xlsm`/`.xltm` template
  passes the checklist but cannot be picked in the dialog; and FILE_PATTERNS regexes accept only
  `\.xlsx?$` (constants.py:15-55) so a `.xlsm` source file passes the scan but always lands in
  "variations". HIGH.
- Excel signature bytes: constants.py:199-202 (unused) vs helpers.py:264-265.
- Minimum file size: helpers.py:303 `>= 1024` vs EXCEL_SAFETY verification_size_min 10240 (unused).
- `'Annual Statement-'` prefix: file_organizer.py:179 and extractor_window.py:623 (`glob(
  "Annual Statement-*")`) hardcode what FOLDER_STRUCTURE['level1'] (constants.py:82) defines;
  `'Version-*'` extractor_window.py:641 vs FOLDER_STRUCTURE['version'] (84).
- Annual-report destination: file_organizer.py:305-312 vs FILE_PATTERNS 'folder' (see 1.9).
- Suggested file-name patterns: file_parser.py:293-304 repeat the exact FILE_PATTERNS
  'description' strings (constants.py:19-59) instead of reading them.
- Colours: GUI_CONFIG['colors'] is used for some widgets, but literal hex values repeat its
  entries elsewhere: '#323130' (= colors.dark) ui_helpers.py:60; '#FFF4E5' (= colors.incomplete)
  extract_tab.py:86-89, setup_tab.py:58-60; '#E8F5E8' (= colors.complete) setup_tab.py:30, 57;
  "0078D4"/"107C10"/"D83B01" (= primary/success/danger) data_consolidator.py:35, 119;
  '#0078D4' extractor_window.py:151, 332, 388, 399. Many more one-off colours in setup_tab.py
  (29-69), validation_tab.py (48, 64-65, 89, 165), processing_tab.py (106-117), status_bar.py:32,
  title_bar.py:50, tooltip.py:54, dark_mode_manager.py:72-78, 166-170, 284-287, extractor_window.py
  (151-586 passim).
- Default client-name length 35: file_organizer.py:30, 42; helpers.py:620; main_window.py:93,
  386-387; setup_tab.py:436, 614-615 — but the long-name warning thresholds are 10:
  client_handler.py:191 (per-client toggle, hardcoded) and processing_handler.py:95, 97 (fallback
  when the setting is invalid). setup_tab.py:446-447 notes the old text "hard-coded 10 chars";
  the per-client warning still does. HIGH.
- Default refresh suffix "_Refreshed_{timestamp}": report_processor.py:51 and
  extractor_window.py:297, while refresh_naming.py:14 defines DEFAULT_SUFFIX_PATTERN.
- Cache file name '.gst_organizer_cache.json': cache_manager.py:15, extractor_main.py:41,
  extractor_window.py:118 (three literals; consistent today).
- Excel path limit 218: excel_handler.py:356; report_processor.py:720, 795 (plus help texts
  setup_tab.py:115-122 and the comment constants.py:85-90).
- COM error code '-2147352567': excel_handler.py:355; report_processor.py:170, 788.
- 'Excel.Application': excel_handler.py:124, 127; report_processor.py:218, 536.
- Mouse-wheel step 3 units per notch: setup_tab.py:138 (named constant) vs extractor_window.py:471
  (literal 3).

**Other hardcoded numbers / decision strings in logic files** (mechanical list in
`/root/merge-scratch/run/A-X-DEAD/magic.txt`; behaviour-relevant ones):
- core/file_parser.py: 'GSTR-3B Export' multiple files always allowed (267) [GST RULE]
  TIME-VARYING (comment: "monthly filings - unlimited"); critical types (359, dead) [GST RULE];
  status texts 'Complete' / f'Missing {n} files' / 'Has duplicates' and icons (279-286);
  keyword tests 'gstr'+'2b', 'ims', 'gstr'+'3b', 'sales'+'reco', 'sales', 'annual' (293-304);
  percentage ×100 (274, 333-334).
- core/file_organizer.py: mode names 'fresh'/'resume' (165, 278, 283, 383); status strings
  'Success'/'Failed'/'Skipped' (259, 285, 291, 364, 412-414); report name
  '_Organization_Report_{client}_{state}.txt' (330); date format '%Y-%m-%d %H:%M:%S' (336);
  files starting with '_' ignored by the dead `_is_folder_empty` (403).
- core/excel_handler.py: temp folder 'gst_excel_temp' (53); sleeps 0.5 s (155, 210); default sheet
  'Links' (163, 286) and fallback "any sheet containing 'link'" (175, COM path only); 50-char
  log/slice limits (195, 363); summary sheet names and headers (434-579); header colour
  "366092" (416); autofit cap 50 (600); report types 'ITC'/'Sales' (620, 657).
- utils/helpers.py: timestamp formats '%d%m%y %H%M' / '%d%m%y' / '%d%m%y_%H%M' (27-35);
  sanitize max 200 and a 10-character reserve (68, 119-120); invalid characters (96); control
  character ranges 0-31 and 127-159 (100); 'Private'→'Pvt', 'Limited'→'Ltd' (92-93, repeated
  624-625); 16 parallel signature readers, parallel only from 4 files (271-272); 8 header bytes
  (283); client-name length 2..100 and invalid characters `<>:"/\|?*` (438-447); unique-path
  limit 100 (469, dead); md5, 8 KB chunks (530, 540, dead); state fallback: first letters of the
  first two words, else first 3 characters (608-618) [GST RULE]; client-state key max 35, min 5
  characters left for the name (620, 637).
- utils/dependency_installer.py: pip timeout 600 s (23); window "760x380" (123).
- gui/handlers/client_handler.py: Treeview column '#7' and value positions 0/1/2/6 (76, 152,
  169-170, 181, 189-201, 218-220); long-name warning > 10 (191); export file name
  'selected_clients_{%Y%m%d_%H%M%S}.txt' (236).
- gui/handlers/file_handler.py: tab index 1 (88); template filter "*.xltx *.xlsx" (31, 42).
- gui/handlers/processing_handler.py: tab indexes 2 and 3 (49, 116, 164); fallback max length 10
  (95, 97); first 5 long names shown (109); progress split 50/50 (267, 287), 95 and 100 (459,
  482); report names 'ITC_Report_{client}_{state}_{ts}.xlsx', 'Sales_Report_...',
  'GST_Processing_Summary_{ts}.xlsx' (297-298, 355-356, 463-464); workflow 'full' (145).
- gui/utils/cache_manager.py: last 5 recent folders (40).
- power_query_extractor/core/report_processor.py: wait 10 s default (50) and at least 10 s
  extra (326); copy wait 3 s (185); property retries 5 × 2 s (222-232); open retries 3 × 3 s
  (239-247); 5 × 1 s load wait (253-255); 3 activation attempts (264); key presses '%a', 'r',
  'a' = Alt, R, A (282-292); 5 × 1 s error checks (310-319); window-title words list (425-436)
  and dialog classes '#32770', 'NUIDialog', 'bosa_sdm_msword' (441); button words 'ok',
  'close', 'yes', 'continue' (506); validation sheet 'Info' cell 'BB2' and the values treated as
  failure: empty, '0', '0.0', error strings (646-678); xlMaximized -4137 (227); long-path
  prefixes `\\?\` and `\\?\UNC\` (727-731); report glob f"{type}_Report_*.xlsx" (112);
  timestamp '%d%m%y_%H%M' (128, duplicates helpers.get_safe_timestamp).
- power_query_extractor/core/data_consolidator.py: report name 'PQ_Extraction_Report_
  {%Y%m%d_%H%M}.xlsx' (175-176 — a different timestamp style from everywhere else); number
  format '#,##0.00' applied only to non-zero numbers (161-162) [GST RULE] (display rounding of
  figures; the stored value is not rounded) — described only; title merges 'A1:C1', 'A{i}:J{i}'.
- power_query_extractor/config/cell_mappings.py: sheet names, cell addresses and "Table 6J" /
  "Table 8D" labels (33-242) — [GST RULE] TIME-VARYING; this is the extractor's config module,
  i.e. the intended home; other notes cover their meaning.
- main.py: log file 'gst_organizer.log' (24); folders created at start 'core','utils','gui',
  'logs','temp' (35); extractor_main.py: 'pq_extractor.log' (16), three cwd-relative cache
  fallbacks and keys (41-44, 55).

### 2.5 Threading — how worker threads talk to the window (READ + RAN)
- **Stated project rule** (extractor_window.py:989-1000, commit c1f425e): "Tk is not
  thread-safe ... Every worker-thread UI update goes through here [`_ui`]". The extractor
  follows it: all log/progress/dialog/button work from `_process_clients_thread` goes through
  `_ui` → `self.after(0, ...)`, which also swallows TclError/RuntimeError (1003-1010). Two
  exceptions: extractor_window.py:890-891 read `wait_time_var.get()` and
  `suffix_pattern_var.get()` (Tk variables) directly from the worker.
- **Main app (ProcessingHandler) only partly follows it** (process_files_thread,
  processing_handler.py:182-530, run as a thread at 129-134):
  - 10 `_log_async` and 4 `_progress_async` calls (schedule via `root.after`, values captured
    by default arguments — commit 0890022);
  - 45 direct `self.app.log_message(...)` calls from the worker (319-523) — these end in
    `log_widget.after(0, ...)` (ui_helpers.py:23), i.e. a Tk call made from the worker thread;
  - `self.app.update_progress(...)` from the worker (459, 482) → `root.after` (main_window.py:289);
  - Tk variable reads from the worker: target_folder / processing_mode /
    include_client_name_in_folders / client_name_max_length (190-194, 219-225), itc_template
    (301, 310), sales_template (332), target_folder (501), workflow_mode (145);
  - modal dialogs raised from the worker: `messagebox.showinfo` (498-502), `messagebox.showerror`
    (525);
  - direct widget changes from the worker: `start_btn.config` / `stop_btn.config` in `finally`
    (529-530);
  - no TclError/RuntimeError protection (unlike `_ui`).
  Python's Tk can pass such calls to the main thread when Tcl is built with threads (the usual
  Windows build), so this mostly works, but it is the exact pattern c1f425e calls "a real cause
  of intermittent freezes". HIGH (READ) for the inconsistency; MEDIUM for real-world freezes.
  RAN ordering check: 3 × 200 interleaved `_log_async` + direct `log_message` calls from a
  worker on Linux/Tcl 9.0.3 kept their order every time — no ordering bug seen (LOW risk).
- **Synchronous work on the window thread**: the source scan (file_handler.py:71-72 →
  file_parser.scan_folder) and the extractor scan (extractor_window.py:611-668) run on the main
  thread, so the window freezes during a slow network scan; the progress callback only
  `config`s a label without repainting (file_handler.py:67-68). MEDIUM (UX).
- **Shared state without locks**: `self.app.client_data`, `variations`, `selected_clients`,
  `client_folder_settings` are read by the worker while the window can still change them (the
  Scan/Re-scan buttons, the per-client folder toggle and the Step 2 buttons stay enabled during
  processing; only Step 3's START is disabled, processing_handler.py:119-120). See bugs B9/B10.
- Thread pool for signature reads (helpers.py:351-357) touches no Tk — fine.
- Pip install (dependency_installer.py:133-134) runs on the main thread; the window repaints
  only through `update_idletasks` in `log` (127-130) — may show "Not responding" during long
  installs. LOW.

### 2.6 File paths (READ + RAN)
- **Main pattern**: `pathlib.Path` — followed by core/*, utils/helpers.py, gui/*, extractor.
- **Breaks it**: `os.path` in launch_extractor.py:10 and client_handler.py:252; `os.scandir`
  used on purpose for speed (helpers.py:328, extractor_window.py:33, 61); string surgery on
  paths (clean_windows_path helpers.py:422; report_processor.py:724-731).
- **App folder vs current folder vs user folder — mixed**:
  - app folder (`Path(__file__).parent`): main.py:31-43 creates `logs/`, `temp/` and missing
    `__init__.py` files inside the app folder at every start (`logs`/`temp` are never used);
    main.py:84 requirements.txt.
  - current folder (wherever the program was started from): both log files (main.py:24,
    extractor_main.py:16) and three extractor cache fallbacks (extractor_main.py:42-44). RAN:
    importing main.py created `gst_organizer.log` in the current folder. The file handler is
    created at import (main.py:20-27), outside main()'s try/except, so an unwritable current
    folder would crash before any message (READ; could not reproduce as root). MEDIUM.
  - user home: the settings cache (cache_manager.py:15, extractor_main.py:41,
    extractor_window.py:118).
  - system temp: excel_handler.py:53.
- `sys.path`: main.py:32 `insert(0, ...)` vs launch_extractor.py:10 `append(...)` (both
  redundant when run as scripts, since Python adds the script folder itself).
- truncate_path display quirk (helpers.py:414): RAN → '//.../Target' on Linux; on Windows the
  first part is 'C:\\' so it builds 'C:\\\\...\\Target' (doubled separator). Cosmetic. HIGH.

### 2.7 `__init__.py` re-exports and `__all__` lists (READ + RAN)
- **No code uses the re-exported names**: grep for `from utils import`, `from core import`,
  `from gui import`, `from gui.tabs import`, `from gui.utils import`, `from gui.widgets
  import`, `from gui.handlers import`: only `gui/tabs/setup_tab.py:21` (`from ..utils import
  setup_validation`, a *submodule*, not a re-exported name) and tests (test_dependency_installer.py:105,
  test_setup_tab_smoke.py:12, test_setup_validation.py:4, test_dark_mode.py:184,
  test_optional_pywin32.py:54,57). Every other import names the submodule directly.
- **Completeness** (RAN: compared each `__all__` with what its submodules define):
  - core/__init__.py:10 — complete.
  - gui/__init__.py:9 — complete (one class).
  - gui/handlers/__init__.py:8 — complete.
  - gui/tabs/__init__.py:8 — **missing `ExtractTab`** (extract_tab.py added in bfcb24a,
    26-07-2026; this `__init__` unchanged since edac6a3).
  - gui/widgets/__init__.py:7 — **missing `Tooltip`** (tooltip.py added in 322717f, 25-07-2026).
  - gui/utils/__init__.py:9 — lists 4 classes; `setup_validation` (322717f) not re-exported
    (it is imported as a submodule, so this is harmless).
  - utils/__init__.py:22-36 — updated once (56aaad9 added `scan_excel_files`) but **missing**
    constants WORKFLOW_MODES (510f7b4), TEMPLATE_EXTENSIONS, MAX_FILE_SIZE_MB, EXCEL_SIGNATURES,
    STATE_CODE_MAPPING; helpers get_short_path, get_state_code, create_client_state_key (all
    present since edac6a3); and the newer modules requirements_check, dependency_installer,
    excel_com. It re-exports dead items (D13-D16, find_excel_files, validate_excel_file).
  - power_query_extractor and its 3 sub-packages: no `__all__`; docstring / `__version__` only.
  - Every name in every `__all__` exists (no broken entries).
- **Side effect of eager imports** (RAN): importing the "Tkinter-free" module
  `gui.utils.setup_validation` (its docstring, setup_validation.py:3-4) loads tkinter,
  gui.main_window, the whole power_query_extractor package and openpyxl — 222 modules —
  because gui/__init__.py:7 imports main_window. Likewise core/__init__.py:6-8 imports
  excel_handler (and so the global `warnings.filterwarnings('ignore')`, 2.2) whenever any core
  module is imported. HIGH.
- The deferred import "to avoid circular import" (helpers.py:597-598) has no real cycle to avoid
  (constants.py imports nothing from utils); harmless (fix history: 6237597). LOW.

### 2.8 Other consistency points (READ)
- **Button wiring**: newer code uses `command=` (setup_tab.py:383-385 explains why: "so
  state='disabled' is honoured and the keyboard works"; collapsible_frame.py:21; title_bar.py:72;
  all extractor buttons); older tabs use `.bind('<Button-1>', ...)` (validation_tab.py:118, 130,
  142, 190; processing_tab.py:74, 87). Consequences: those buttons cannot be pressed with the
  keyboard; the disabled STOP button still fires (processing_tab.py:84-87 → logs "🛑 Stop
  requested by user" while nothing runs; harmless because start resets the flag,
  processing_handler.py:123); the Step-2 START is never disabled (re-entry is caught by the
  guard at processing_handler.py:75-76). HIGH.
- **Fonts**: two families — 'Segoe UI' (newer: setup_tab, extract_tab, tooltip, extractor) and
  'Arial' (older: validation_tab, processing_tab, status_bar, title_bar, ui_helpers,
  collapsible_frame); GUI_CONFIG['fonts'] unused.
- **Timestamp formats** for names: folders '%d%m%y %H%M' (helpers.py:27); report files the same
  with '_' (processing_handler.py:294); refreshed copies '%d%m%y_%H%M' (report_processor.py:128);
  consolidated report '%Y%m%d_%H%M' (data_consolidator.py:175); export list '%Y%m%d_%H%M%S'
  (client_handler.py:236); backups '%d%m%y_%H%M' (helpers.py:224). Not [GST RULE] (file naming).
- **Version strings**: 'v3.0' literals main.py:2,6,63; status_bar.py:29; main_window.py:393;
  constants.py:134; three `__version__ = '3.0.0'`; extractor '1.0.0'; requirements.txt:1.
- **Duplicated logic**: two near-identical `_autofit_columns` (excel_handler.py:587-601,
  data_consolidator.py:416-432); two mouse-wheel handlers (setup_tab.py:211-248,
  extractor_window.py:450-478) — and a third, the extractor log box (extractor_window.py:559-562),
  still uses `int(-1*(event.delta/120))`, the truncation that setup_tab.py:240-243 says makes
  precision-touchpad scrolling "appear dead" (small deltas round to 0). HIGH; two
  `_safe_open_workbook`s (deliberately different, utils/excel_com.py:3-14); two
  `_create_summary_sheet`s with different content; the Private/Limited abbreviation regex twice
  (helpers.py:92-93, 624-625).
- **Max client name length semantics**: the setting is passed as `max_length` of
  `create_client_state_key` (file_organizer.py:83-87), i.e. it limits the whole Level-2 folder
  name "client-STATE", not the client name; the category folders "ITC (client)" use the
  untruncated `sanitize_filename(client)` (file_organizer.py:90, 119). But the warning in
  start_processing compares the raw client name with the setting and only when "Include client
  name in Level 4 folder names" is on (processing_handler.py:90-113), and the help text says
  "Longest client name allowed in a folder name before it is shortened" (setup_tab.py:119-123).
  So the setting does not shorten the folders the warning is about. MEDIUM-HIGH (READ).

---

## 3. Silent-bug scan (audit method step 9)

### 3.1 Closure / lambda late binding in loops — no live bug (READ)
- validation_tab.py:190 `lambda e, cmd=command: cmd()` — value captured by default → correct.
- processing_handler.py:31, 36 — values captured by defaults (fix 0890022) → correct.
- processing_handler.py:266-269 `file_progress` is defined inside the per-client loop and
  reads `progress`, `idx`, `time_remaining` by reference, but it is called synchronously by
  `organize_files` in the same iteration (271-273) → harmless.
- main_window.py:289, ui_helpers.py:23-25, extractor_window.py:1002 — lambdas capture function
  parameters (new scope per call) → correct.
- setup_tab.py:167-171, 292-314; extractor_window.py:426-429, 688-753 — no loop-variable capture
  in deferred code → correct.

### 3.2 Mutable default arguments — none (RAN grep `def ...=[]|{}|set()`: no match).
`FileOrganizer.__init__(client_folder_settings=None)` uses `or {}` (file_organizer.py:41) —
correct. `check_requirements(requirements=REQUIREMENTS)` default is a tuple — fine.

### 3.3 Off-by-one
- **B-ETA** processing_handler.py:235, 457: the tracker is advanced by 0 for the first client
  and 1 for each later one and never for the last finished client, so after the loop `current =
  total-1`; the "Creating summary report | ETA: ..." message (458) shows one client's worth of
  time remaining although all clients are done; with a single client it shows "Unknown"
  (helpers.py:499-500). Display only. HIGH.
- setup_validation.py:67-72 `_short`: keeps 42 characters + '...' = 45, not 46; and it trims the
  START of the path although its docstring says "Trim the middle". Cosmetic. HIGH.
- No other off-by-one found in range/slice code (checked helpers.py:119-120, 404-416, 610-642;
  excel_handler.py:457-601; data_consolidator.py:52-166; report_processor.py:223-330).

### 3.4 String vs number comparison
- **B2 Treeview turns numeric-looking text into numbers** (RAN, tkprobe/probe_tree.py): a
  client name '007' comes back from `client_tree.item(...)['values']` as the integer 7; '123'
  as 123. Code rebuilds the client key from these values: client_handler.py:167-171
  (`get_selected_clients`), 177-181 (`toggle_folder_name_setting`), 216-218. Effects for names
  that are all digits with a leading zero: key "7-DL" ≠ stored "007-DL" → `KeyError` in
  dry_run (processing_handler.py:56, uncaught in a Tk callback), in processing
  (processing_handler.py:232 — outside the per-client try, so the outer handler stops the whole
  run with "Fatal error", 522-525), in export (client_handler.py:248, caught → "Export Error");
  per-client folder toggles silently ignored (FileOrganizer looks up the string key,
  file_organizer.py:122-130). For any all-digit name (even '123'), `len(values[0])`
  (client_handler.py:191, 194) raises TypeError after the setting was already flipped (185) but
  before the display updates (209) → screen and setting disagree. Rare in practice (client part
  of a file name being all digits). HIGH (mechanism), LOW (likelihood).
- **B3 COM error values** (READ, cannot RUN): report_processor.py:675-678 compares the cell text
  with '#REF!', '#VALUE!', ... but pywin32 normally returns an Excel error cell as a large
  negative integer (e.g. -2146826281 for #DIV/0!), not as that text. If so, BB2 holding an error
  passes validation, and an error in an extracted cell (584-585) lands in the consolidated report
  as a large negative number formatted '#,##0.00' (data_consolidator.py:157-162). MEDIUM — how
  to check: on Windows with Excel, put =1/0 into a mapped cell (and into Info!BB2), run Step 4,
  and look at the log line "✓ <field> (<cell>) = ..." and the output cell.
- report_processor.py:670 `cell_str == '0' or '0.0'` — matches how Python prints 0/0.0; fine.
- main_window.py:386-387 `cached_length > 0` — assumes the cached value is a number (it is
  written from IntVar.get) — fine unless the cache file is hand-edited. LOW.

### 3.5 `==`/`is` with None/True/False — none found (RAN grep: no `== None`, `!= None`,
`== True`, `== False`, `is True`, `is False` in non-test code). `sum(selected.values())` on
booleans (client_handler.py:130) is intentional and correct.

### 3.6 Stale variables / values that never change
- **B1 Excel the user already had open is hidden and closed** (core/excel_handler.py:123-131,
  238-246). `GetActiveObject("Excel.Application")` (124) attaches to an Excel the user already
  has open; the code then sets `Visible = False` (131) on it; in `finally` the comment says
  "Don't quit if Excel was already open", but the test is `if self.excel_app is None` and
  `excel_app` is never assigned (only line 56), so `excel.Quit()` (244) is always called — on
  the user's own Excel too, right after `DisplayAlerts = True` (240), so Excel may prompt about
  the user's unsaved workbooks or close them. Present since edac6a3. HIGH (code path); MEDIUM
  (exact Excel behaviour — how to check: open any workbook in Excel, run Step 3 for one client,
  watch the open workbook). Cross-ref Standard 12 line 853-854 ("Never take over the user's own
  open Excel"). The A-X-EXCEL notes cover COM in depth.
- **B4 "Resume" never resumes** (RAN, resume/probe_resume.py with placeholder client
  `<Client A>`): resume mode reuses the newest 'Annual Statement-*' folder
  (file_organizer.py:175-186) but every run makes a NEW 'Version-DDMMYY HHMM' folder
  (103-108), so `dest_path.exists()` (277) is false and nothing is skipped — a fresh run at
  10:00 then a resume at 10:15 gave two Version folders and the file copied again ('Success').
  Files are skipped only if two runs fall in the same minute. The help text says "Continue a
  run that was interrupted, skipping files that are already in place" (setup_tab.py:108-109).
  HIGH.
  Related: in resume mode a Skipped file's record keeps its original (unsanitized) name
  (organize_files updates the name only on 'Success', file_organizer.py:234-246), so the report
  links (excel_handler.py:628-700) would name a file that is on disk under its sanitized name
  — only matters in the same-minute case. MEDIUM.
- **B5 openpyxl fallback reports success without writing links** (excel_handler.py:289-310): if
  no sheet is named exactly 'Links' (case-insensitive), nothing is written, every per-cell
  error is swallowed by a bare except (302-303), the copy is still saved and the function
  returns `output_path.exists()` → True; processing counts it as a created report
  (processing_handler.py:309-319). The COM path at least tries a 'link' sheet or creates one
  (172-183). Only reachable where COM is unavailable (off Windows, or Windows without pywin32 if
  the user ignores the start-up block). HIGH (code), LOW (likelihood on Windows).
- **B6 Wrong figures routed to three extractor output sheets** [GST RULE] (RAN,
  consol/probe_consol.py, placeholder values 1001-1017 = one per source cell, in the order
  `_extract_data` stores them, report_processor.py:556-594): `create_report` first looks for the
  FIRST extracted field whose name merely CONTAINS the output column text
  (data_consolidator.py:141-145) and uses the exact key only as a fallback (147-150). Output
  columns with spaces ('Diff IGST', 'IGST Diff', 'Taxable Value', ...) never match a field name
  (fields use underscores) and correctly fall back. But the bare columns 'IGST', 'CGST', 'SGST'
  of the sheets 'IMS Reco', 'Extra In Purchase' and 'Extra In GSTR2B' (cell_mappings.py:155-228)
  match 'GSTR-3B_Difference_Diff_IGST' etc. first. Result of the run: 'IMS Reco': IGST=1001,
  CGST=1002, SGST=1003; 'Extra In Purchase' the same; 'Extra In GSTR2B' the same — i.e. the
  values of Diff GSTR-3B J15/K15/L15, not IMS_Reco AA1/AD1/AG1 (1009-1011), ExtraInPurchase
  O1/P1/Q1 (1012-1014) or ExtraInGSTR2B S1/T1/U1 (1015-1017) as configured. The other six sheets
  came out as configured. If the earlier-matched field is empty (`None`) the fallback finds the
  configured value; a 0 is kept. The substring logic dates from edac6a3; the bare 'IGST'/'CGST'/
  'SGST' columns were added in 47a1492 (30-10-2025, the last commit shared with B; B's
  data_consolidator.py:131 has the same line). I describe only; which cell is meant to feed each
  column is for the author. HIGH (reproduced).
- **B7 First file only** [GST RULE]: when a client has several files of one type (status 'Has
  duplicates'), `prepare_template_data` links only the first one (`[0]`, excel_handler.py:628,
  639, 650, 666, 682, 698; order = sorted path order from scan_excel_files, helpers.py:361); the
  others are copied but not linked. GSTR-3B is exempt because its folder is linked, not a file.
  Described only. HIGH.
- **B8 Workbook open retries can run out silently** (report_processor.py:239-249): if all three
  attempts fail with "Call was rejected by callee", the loop ends with `wb = None` and no error;
  the code then presses keys in Excel, waits, and fails later with "'NoneType' object has no
  attribute 'Save'" (caught at 376-381) — a misleading error message. HIGH (code).
- IntVar with bad text (RAN, tkprobe): `client_name_max_length.get()` raises TclError when the
  spin box is empty or not a number. It is read in `save_cache` (main_window.py:369), which every
  Browse button, the dark-mode toggle and the global folder checkbox call (file_handler.py:25,
  36, 47, 55; main_window.py:196; client_handler.py:223) → those callbacks error out and the
  settings are not saved; and in the worker (processing_handler.py:194) → "Fatal error" for the
  whole run. The spin box's own command (setup_tab.py:434-438) runs only on arrow clicks, not on
  typing. Same for the extractor wait time (extractor_window.py:864, 890). HIGH (mechanism).
- Summary label (main_window.py:312) "Clients with Missing Files" shows `incomplete_clients`
  (file_parser.py:328) = all clients not 'Complete', which also counts 'Has duplicates' clients
  that miss nothing. Display only. HIGH.
- Installer installs more than it offers: main.py:84-88 passes requirements.txt when it exists,
  and then `install(packages=None, requirements_path=...)` (dependency_installer.py:133-134)
  installs the WHOLE file (openpyxl, pywin32, lxml, pytest) while the prompt lists only the
  missing packages (110-115); the manual command shown on failure combines both (142). HIGH.
- Possible: `create_backup` in 'fresh' mode writes "<name>_backup_<ts>.xlsx" into the same
  category folder (helpers.py:224-228, file_organizer.py:277-282). If the ITC template's Power
  Query reads every file in the GSTR-3B folder (B2 holds the folder, constants.py:106), a backup
  would be read as an extra file. Happens only when the destination already exists (same-minute
  rerun or two names that sanitize to the same text). MEDIUM/LOW — how to check: open the ITC
  template's Power Query editor and see whether the GSTR-3B source is "Folder" with no file
  filter. Also: the .gitignore pattern `backup_*` (and EXCEL_SAFETY 'backup_prefix') does not
  match this naming.
- [GST RULE] state-code merging (described only): helpers.get_state_code maps names to the
  two-letter abbreviations in STATE_CODE_MAPPING (constants.py:251-292, TIME-VARYING list) and,
  for an unknown name, invents one from first letters / first 3 characters (helpers.py:608-618).
  The client key is "client-CODE" (file_parser.py:229-230), so two state spellings with the same
  code (e.g. the 'orissa'/'odisha', 'pondicherry'/'puducherry', 'delhi'/'national capital
  territory of delhi' pairs in the map) merge into one client entry whose displayed state is the
  last one seen (235), and two different unknown states whose invented codes coincide would
  merge too. HIGH (mechanism).

### 3.7 Variables used after a loop holding only the last value — none harmful (READ):
processing_handler.py:456-516 uses `level1_folder` (set once from the first client, 255-256 —
intended) and `timestamp`; excel_handler.py:166-183 and report_processor.py:564-576 use
`for ... break` lookups correctly; cell_mappings.py:264-278 leaves loop names as module globals
(1.6).

### 3.8 Thread-unsafe shared state (READ)
- **B9 Re-scan during processing**: Scan/Re-scan stay enabled during a run (only the Step-3
  START is disabled, processing_handler.py:119-120; setup_tab.py:585-591 enables them from the
  checklist). `rescan_files` clears the same dictionaries the worker reads (file_handler.py:102-105)
  and `scan_files` replaces them (71-72) on the window thread. A worker reading
  `self.app.client_data[client_key]` (processing_handler.py:232) in that window gets KeyError →
  whole run stops ("Fatal error"); `summary_data['variations']` holds the old list object
  (224), which `.clear()` empties → empty Variations sheet. MEDIUM.
- **B10 Settings changed mid-run apply mid-run**: `client_folder_settings` is passed by
  reference into FileOrganizer (processing_handler.py:193, file_organizer.py:41); toggling a
  client's folder-name cell during a run changes the folders of clients processed later. LOW.
- **B11 Closing during processing**: on_closing sets `stop_requested` and destroys the window
  after 1 s (main_window.py:351-353), but the worker checks the flag only between clients
  (processing_handler.py:229-230) and keeps calling Tk afterwards → TclError/RuntimeError in the
  worker, dialogs from a dead window, and an Excel instance possibly left open mid-report.
  MEDIUM.
- Extractor: `EXTRACTION_OPTIONS['skip_refresh']` is a module-level dictionary changed by the
  checkbox (extractor_window.py:486-487) and read by the worker (report_processor.py:134-135) —
  safe under Python's global lock, and both panel instances in one process would share it. LOW.
- `ReportProcessor.process_client` returns the same status dictionary object for ITC and Sales
  when pywin32 is missing (report_processor.py:70-77, `dict(empty)` is a shallow copy; RAN:
  `r['itc']['status'] is r['sales']['status']` → True). Never mutated later → harmless.

### 3.9 Widgets touched from worker threads — full list in 2.5. Summary: main app
processing_handler.py:145, 190-194, 219-225, 301, 310, 319-523 (45 log calls), 332, 459, 482,
498-502, 501, 525, 529-530; extractor 890-891. Everything else goes through `after`.

---

## 4. Commented-out code and leftovers

### 4.1 Commented-out code blocks (RAN grep for comment lines that look like code)
- Only power_query_extractor/config/cell_mappings.py:229-241 — a commented example sheet
  ("ADD MORE SHEETS HERE AS NEEDED"); intentional documentation for the author.
- requirements.txt: commented dev tools `# black>=23.0.0`, `# pylint>=2.17.0`.
No other commented-out code in the 43 non-test files.

### 4.2 AI-chat paste markers left in code (RAN grep)
data_consolidator.py:75, 80 `# ... rest of the code` (with the dead duplicate block 73-83),
:222 `# ADD THIS LINE:`, :379, :402 `# ADD THESE LINES`; report_processor.py:132 `# ADD THE NEW
CODE HERE (notice the indentation - 3 levels):`, :183 `# ADD THIS: ...`, :295 `# Rest of your
code remains same...`; file_organizer.py:15 `# ADD THESE`; excel_handler.py:39 `# ADD THESE
TWO`, :449-451 `# NEW`; helpers.py:17 `# ADD THIS LINE`, :22 `# UPDATED:`, :577 `# NEW:`;
processing_handler.py:86 `# CHANGE:`. No TODO/FIXME/XXX/HACK markers. main.py:5 "Author: Advanced
AI Assistant"; file_parser.py:4 "Fixed with proper Union import."

### 4.3 The `gui/app.py` leftover (RAN git)
- `git log --all -- gui/app.py` → no commits: the source was never committed.
- edac6a3 (31-08-2025, Initial commit) committed `gui/__pycache__/app.cpython-312.pyc` and
  `app.cpython-313.pyc` (the only compiled files in that commit with no matching source — RAN
  check over all 70 committed `.pyc` files). Their headers (RAN) record a source of 89,275 bytes with
  modification time 25-06-2025 16:14 UTC, compiled from a folder whose name ends "...data
  extractor v2.1\gui\app.py" (author's machine path; not client data). Identifier strings in it
  (RAN `strings`) include `GSTOrganizerApp`, `create_setup_tab`, `create_validation_tab`,
  `create_processing_tab`, `create_title_bar`, `restore_light_mode`, `update_colors`,
  `create_welcome_section`, `create_help_section` — i.e. the single-file GUI that was later split
  into gui/main_window.py + tabs/ + handlers/ + utils/ (HIGH that it is the monolithic
  predecessor; MEDIUM on the date, which is the source time stored in the compiled file).
  The committed `gui/__init__` compiled file (source time 23-07-2025, "...v2.2" folder) already
  imported `GSTOrganizerApp` from `main_window`.
- 01862a0 (29-10-2025) deleted all `__pycache__` files; .gitignore (1cb878b) now excludes them.
- Nothing imports `gui.app` at HEAD (grep: none).
- main.py kept telling users to check for "📄 app.py" in its ImportError message until f363946
  (26-07-2026). Commit 6237597 (29-10-2025) says "Corrected non-existent app.py reference in
  main.py error messages", but 6237597 did not touch main.py (its diff covers file_parser.py,
  file_handler.py, processing_handler.py, helpers.py only) — commit message vs diff mismatch.
  A's own audit file (PROJECT_AUDIT_ANNUALSTATEMENT.md:108, commit f7eb8f1, 05-03-2026) still
  flagged that message, correctly for its time.

### 4.4 Other leftovers
- `power_query_extractor/config/cell_mappings copy.py` — committed in edac6a3, deleted in
  47a1492 (30-10-2025).
- `gst_organizer.log` — committed in edac6a3, deleted in cbfa4a5 (not opened, per instructions).
- main.py:35-43 still creates `logs/` and `temp/` (never used) and writes a placeholder
  `__init__.py` into core/utils/gui if missing.
- `PROJECT_AUDIT_ANNUALSTATEMENT.md` (fc279af/f7eb8f1, March 2026) is an earlier audit kept in the
  repo; parts are stale (it predates the July 2026 commits). One claim was wrong even then:
  "Only `save_options` is referenced (by excel_handler.py:54 ...)" (its line 1151) — nothing reads
  `save_options`; the whole EXCEL_SAFETY dict is stored and never read (checked at f7eb8f1:
  `self.safety_config = EXCEL_SAFETY` at excel_handler.py:54 is its only use). Its dead list
  (its lines 96-108) has D9, D10, D11, D12, D15, TEMPLATE_EXTENSIONS and EXCEL_SIGNATURES, but
  misses D1-D8, D13, D14 and D16 — RAN `git grep` at f7eb8f1: those had exactly the same
  (missing) call sites then, so they were already dead when that audit was written.
- `docs/superpowers/specs/2026-07-25-merge-extractor-design.md` — design spec for the Step-4 merge
  (not read in detail; not code).
- `.claude/settings.local.json` — allows `python -m py_compile` for the AI tool (tooling).

---

## 5. Factual cross-references to the author's standards (no judgement on GST rules)
- Standard 12 "Excel through COM" (rebuild_pipeline_FINAL.md:851-856): asks for a separate hidden
  Excel via `DispatchEx` and "Never take over the user's own open Excel"; A uses
  `GetActiveObject` then `Dispatch` (excel_handler.py:124-127) and `Dispatch` (report_processor.py:218,
  536); `Quit()` is in `finally` in all three places (excel_handler.py:238-246;
  report_processor.py:404-409, 626-631) but with the B1 problem.
- Standard 13 (lines 1034-1040): "Find the app's own files from APP_DIR, never from the current
  folder"; "There is no console: print() shows nothing"; "Settings and logs go in ...
  %LOCALAPPDATA%..., never in the shared app folder". A writes logs to the current folder
  (main.py:24; extractor_main.py:16), creates folders in the app folder (main.py:35-43), prints
  its start-up dependency report to the console (main.py:62-73), and crashes at import under
  `pythonw.exe` (main.py:17, RAN). Settings go to the user's home folder (cache_manager.py:15).
- Standard 13 lines 1024-1025 (`sys.excepthook` / `threading.excepthook`): A sets neither; worker
  thread errors are caught by try/except inside the threads (processing_handler.py:522-525,
  extractor_window.py:957-960).

---

## Possible client names
None found. Searched: all 43 non-test `.py` files, all 19 test files (fixture names use only
generic 'ABC', 'XYZ', 'Client', 'ABC-DL', 'ABC Private Limited'), all A/main commit messages
(only the author's name/handle and generic 'ABC_Private_Limited' examples), and the strings in
the committed `gui/__pycache__/app.cpython-313.pyc` (no business names; it embeds the author's own
folder path). `gst_organizer.log` was not opened. Placeholders used in my probes: `<Client A>`,
'ClientA'.

## Questions for the author
1. Resume mode (B4): should "Resume" reuse the latest Version folder of the last run (so files
   already copied are skipped), or is a new Version folder per run intended?
2. [GST RULE] Extractor sheets 'IMS Reco', 'Extra In Purchase', 'Extra In GSTR2B' (B6): with the
   current code their IGST/CGST/SGST columns show the Diff GSTR-3B J15/K15/L15 values. Which cells
   are meant to feed those columns? (This affects A and B alike.)
3. Should Step 3 ever attach to an Excel the user already has open (B1)? Today it hides that Excel
   and then quits it.
4. "Max client folder name length": should it limit the Level-2 "client-STATE" folder (what it
   does), the Level-4 "ITC (client)" folders (what the warning talks about), or both? And should
   the per-client warning use the setting instead of the fixed 10?
5. Installer: should "Install" install only the missing package(s) shown in the prompt, or the
   whole requirements.txt (which includes pytest)?
6. Are the unused settings (EXCEL_SAFETY, MESSAGES success/info, GUI_CONFIG fonts, TEMPLATE_EXTENSIONS,
   MAX_FILE_SIZE_MB, EXTRACTION_OPTIONS auto_refresh_if_missing / show_timestamp) planned features
   to keep, or safe to drop in the merge?
7. [GST RULE] When a client has two files of the same type, only the first is linked into the
   ITC/Sales report (B7). Is that intended, or should duplicates block processing?
8. [GST RULE] Confirm the rules as coded: all six file types expected per client
   (constants.py:68-75); several GSTR-3B files always allowed (file_parser.py:266-268); the
   two-letter state list and its fallback (constants.py:251-292, helpers.py:592-618) —
   TIME-VARYING; which list is your reference?
9. Should "Has duplicates" clients be counted under "Clients with Missing Files" in the summary?
10. Does the ITC template's Power Query read every file in the GSTR-3B folder (relevant to the
    backup-file point in 3.6)?
11. Excel error cells in extracted data (B3): should they show as an error/blank rather than a
    number? (Needs a Windows test first.)

## INCOMPLETE
- Windows-only behaviour could not be executed here (no Excel / pywin32): B1, B3, B8 and every
  COM/`win32gui`/`os.startfile` path are traced by reading only.
- Test files were searched with grep for every dead-code candidate, not read line by line
  (except tests/test_extractor_standalone.py); the "test-only" verdicts rest on those searches.
- The "unwritable current folder" crash (2.6) was reasoned, not reproduced (running as root
  bypasses folder permissions).
- I did not re-check, commit by commit, when each dead item became dead (only where noted).
