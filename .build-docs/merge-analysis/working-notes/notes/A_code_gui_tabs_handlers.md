# A-R3 working notes - Project A (goyal-sidhant/AnnualStatement_oldv1 @ bf2f1dd) - main window, tabs, handlers

Reader: A-R3. Scope: `gui/__init__.py`, `gui/main_window.py`, `gui/tabs/__init__.py`, `gui/tabs/extract_tab.py`,
`gui/tabs/processing_tab.py`, `gui/tabs/setup_tab.py`, `gui/tabs/validation_tab.py`, `gui/handlers/__init__.py`,
`gui/handlers/client_handler.py`, `gui/handlers/file_handler.py`, `gui/handlers/processing_handler.py`.
Every line of all 11 files was read with the Read tool (full files). Other files (core/*, utils/*, gui/utils/*,
gui/widgets/*, power_query_extractor/gui/extractor_window.py, tests) were read only as far as needed to find
callers and understand calls; their full sections belong to A-R1 / A-R2 / A-R4 and the test reader.

## How to read these notes

- Paths are repository-relative (`gui/handlers/processing_handler.py:232`). Line numbers are those of the snapshot.
- **READ** = established by reading the code. **RAN** = I executed it (see "Experiments" below).
- Inferences carry **HIGH / MEDIUM / LOW** confidence.
- **[GST RULE]** = anything deciding a tax figure, which return/table/sheet/cell a figure belongs to, a rate, limit,
  threshold, rounding, financial year/tax period, GSTIN/PAN/state-code logic, ITC eligibility, due date. These are
  described only - never judged. **TIME-VARYING** marks values that change over time.
- No client data appears here. Experiments used only invented names ("ClientA", "ClientB", "ClientC", "ClientD",
  "007", "1234", "GenericLongClientNamePlaceholder Private Limited").

### Line counts

`wc -l` under-counts by one for 9 of the 11 files because they have no final newline character. The counts below
are the real number of lines (= last line number the Read tool shows):

| File | Lines | Final newline? |
|---|---|---|
| gui/__init__.py | 11 | no |
| gui/handlers/__init__.py | 8 | no |
| gui/handlers/client_handler.py | 257 | no |
| gui/handlers/file_handler.py | 151 | no |
| gui/handlers/processing_handler.py | 530 | no |
| gui/main_window.py | 398 | no |
| gui/tabs/__init__.py | 8 | no |
| gui/tabs/extract_tab.py | 131 | yes |
| gui/tabs/processing_tab.py | 117 | no |
| gui/tabs/setup_tab.py | 615 | yes |
| gui/tabs/validation_tab.py | 247 | no |

None of the files contains Windows CR LF line endings (RAN: byte check).

### Plain-English glossary (terms used throughout)

- **GUI** (graphical user interface: the window, tabs, buttons and boxes the user sees).
- **Tkinter / Tk** (Python's built-in toolkit for drawing windows; project A uses it, project B uses PyQt5).
- **Widget** (any single on-screen element: a button, label, text box, checkbox, table).
- **Notebook / tab** (Tk's container that shows the "Step 1 ... Step 4" tabs).
- **Treeview** (Tk's table widget; used for the Step 2 client list).
- **Callback / event handler** (a function the window calls when something happens, e.g. a click).
- **`command=` vs `.bind('<Button-1>')`** (two ways to connect a button to a function. `command=` runs on a normal
  click, respects a "disabled" button and works from the keyboard. `.bind('<Button-1>')` reacts to the raw mouse
  press, still fires when the button only *looks* disabled, and ignores the keyboard).
- **Thread / worker thread** (a second line of work running at the same time, so the window stays responsive).
  **Main thread / UI thread** (the one that draws the window). Tk is designed to be touched only from the main thread.
- **`after(0, f)`** (asks Tk to run `f` on the main thread as soon as it is free).
- **Daemon thread** (a worker thread that is killed instantly when the program exits, mid-task if need be).
- **Tk variable** (`StringVar`, `BooleanVar`, `IntVar`, `DoubleVar`: a value holder that widgets display/edit).
- **Trace** (an automatic call made whenever a Tk variable is written).
- **Debounce** (wait until typing pauses before re-checking).
- **Cache** (here: a small JSON settings file, `~/.gst_organizer_cache.json`, that remembers the last choices).
  **JSON** (a plain-text format for structured data).
- **COM** (Windows' mechanism for one program to drive another, here Excel). **pywin32** (the Python package that
  provides COM). **openpyxl** (a Python library that edits .xlsx files without Excel).
- **Modal dialog / message box** (a pop-up that blocks until answered).
- **Re-entrancy guard** (a check that stops the same action from starting twice at once).
- **Dead code** (code that nothing ever calls).
- **mtime** (a file's or folder's "last modified" time).
- **Regex** (a text pattern used to recognise file names).
- **Network share / UNC path** (a folder on a server, `\\server\share\...`; every access is a slow round-trip).

### Experiments (RAN)

Environment: Linux, Python 3.14.6, Tk/Tcl 9.0, under Xvfb (a virtual screen); pywin32 absent, so reports were built
by the openpyxl fallback path (core/excel_handler.py:96-102) and every Windows-only branch (Excel COM, `os.startfile`)
was traced by reading only. A full copy of the snapshot was placed in `/root/merge-scratch/run/A-R3/A/`
(without `__pycache__`), and runs used `python -B`, `PYTHONDONTWRITEBYTECODE=1`, `HOME=/root/merge-scratch/run/A-R3/home`
(so the settings cache was written only there) and `TMPDIR=/root/merge-scratch/run/A-R3/tmp` (so ExcelHandler's
staging folder `gst_excel_temp` was created only there). Nothing was written into A-copy (RAN: `find -newer` empty;
the `__pycache__` folders already in A-copy date from 03:21, when A-copy was made, before this task).
Scripts and data (all synthetic) are in `/root/merge-scratch/run/A-R3/`:

| Script | What it does |
|---|---|
| exp1_full_run.py `<mode>` | Builds the real `GSTOrganizerApp`, fills the 4 paths, scans 10 synthetic files (2 clients + 1 unparseable), selects all, presses Start, waits inside the real Tk main loop, dumps log/dialogs/output tree. Modes: `organize`, `full`, `globalname`, `stopfirst`, `darkscan`. Message boxes are replaced by recorders that also note which thread called them. |
| exp2_numeric_names.py | Digit-only client names ("007", "1234") and a digit-only state field ("07") through Step 2. |
| exp3_disabled_bind.py | Clicks the disabled Stop button; invokes the disabled Scan button. |
| exp4_close_midrun.py | Closes the window (answering Yes) while a run is in progress (report creation slowed to 2 s). |
| exp5_spinbox.py | Types '', 'abc', '0', '5', '500', '20.5' into the max-length spinbox variable and calls save_cache. |
| exp6_log_order.py | 200 back-to-back pairs of `_log_async` + direct `log_message` from a worker thread; checks order. |
| exp7_modes_length.py | Runs fresh, then resume, then rerun one minute apart, with include-client-name ON and max length 15. |
| exp8_midrun_changes.py | Export list on Linux; template path changed mid-run; Re-scan pressed mid-run; reads the summary workbook. |
| pytest_out.txt | Output of the existing tests that cover these files (126 passed, 1 failed - see below). |

Existing tests run (RAN, in the scratch copy): tests/test_processing_handler.py, test_workflow_mode.py,
test_main_window_tabs.py, test_setup_tab_smoke.py, test_dark_mode.py, test_scroll_bindings.py,
test_optional_pywin32.py, test_setup_validation.py: 126 passed, 1 failed. The failure is
`tests/test_setup_validation.py:66` (`test_ok_template_shows_filename_only`), which uses a Windows path
`C:\gst\itc.xltx`; on Linux `Path(...).name` returns the whole string - a platform-dependent test, not a code bug
(HIGH). (Test files belong to the test reader; noted here only because they exercise my files.)

---

## Per-file sections (path order)

### gui/__init__.py (11 lines) - [UI]

- **Purpose:** package marker for `gui`; re-exports the main window class and declares a version string. (READ)
- **Imports:** project only - `from .main_window import GSTOrganizerApp` (line 7). Side effect (READ, HIGH): importing
  *anything* under `gui` (for example `gui.tabs.setup_tab` in tests) first runs this file, which imports
  `gui/main_window.py` and through it core/*, utils/*, all four tabs, all three handlers and the extractor panel.
  No OS-specific modules.
- **Module-level names:** `__all__ = ['GSTOrganizerApp']` (line 9); `__version__ = '3.0.0'` (line 11).
- **Functions/classes:** none.
- **Callers:** no file does `from gui import GSTOrganizerApp` (READ, grep of all .py). `main.py:109` imports
  `from gui.main_window import GSTOrganizerApp` directly. So the re-export is unused; `__version__` is never read
  anywhere (the same unused `__version__ = '3.0.0'` exists in core/__init__.py:12 and utils/__init__.py:38).
- **Business rules:** none. **Edge cases:** none. **Excel / files:** none. **UI:** none.
- **Comments/version:** docstring lines 2-5 "GUI package for GST File Organizer v3.0 / Contains the graphical user
  interface components." - agrees with the window title `GUI_CONFIG['window']['title'] = 'GST File Organizer v3.0'`
  (utils/constants.py:134). Project A consistently calls itself v3.0 (B calls itself v4.0).
- **Pattern note:** re-export `__init__.py` files (this one, gui/tabs/__init__.py, gui/handlers/__init__.py,
  gui/widgets/__init__.py, gui/utils/__init__.py) are all bypassed by the real importers, which import submodules
  directly. They add maintenance (lists must be kept in step) without being used. (READ, HIGH)
- **History:** created in edac6a3 (31-08-2025); not changed since (READ: `git log A/main -- gui/__init__.py` lists
  only edac6a3).

### gui/handlers/__init__.py (8 lines) - [UI]

- **Purpose:** re-exports `FileHandler`, `ClientHandler`, `ProcessingHandler` (lines 4-6); `__all__` (line 8). (READ)
- **Imports:** project only (lines 4-6). Importing any one handler module runs this file and so imports all three.
- **Callers of the re-export:** none - `gui/main_window.py:30-32` imports each handler from its own module; tests
  import `gui.handlers.processing_handler` directly (tests/test_processing_handler.py:6, tests/test_workflow_mode.py:9).
  Dead re-export (READ, HIGH).
- Nothing else (no constants, functions, rules, files or UI). **History:** edac6a3 only.

### gui/handlers/client_handler.py (257 lines) - [UI]

- **Purpose:** everything the Step 2 client table does: fill it from scan results, tick/untick clients (mouse and
  keyboard), the per-client "include client name in folder names" toggle and the global toggle, rebuild the list of
  selected client keys for processing, and export the selection to a text file. (READ)
- **Imports:**
  - standard library: `os` (line 4; used for `os.path.basename` line 252 and `os.startfile` line 255),
    `datetime` (line 5), `pathlib.Path` (line 6 - **unused**, READ via AST scan), `tkinter.messagebox`,
    `tkinter.filedialog` (line 7).
  - project: `utils.constants.GUI_CONFIG` (line 8), `utils.helpers.get_state_code` (line 9).
  - OS-specific: `os.startfile` (Windows-only function) at line 255 - **not guarded** by a platform check; it sits
    inside the `try` at 240-257, so on a non-Windows system it raises `AttributeError`, which is caught and shown as
    "Export Error" *after* the "Export Complete" box (RAN exp8: both dialogs appeared, in that order).
- **Module-level constants/globals:** none.
- **Class `ClientHandler` (lines 12-257).** Instantiated once at gui/main_window.py:128 (`ClientHandler(self)`).
  - `__init__(self, app_instance)` (15-16): stores the main window object as `self.app`. Caller: main_window.py:128.
  - `update_client_tree(self)` (18-65). Steps: (1) returns at once if the app has no `client_tree` yet (20-21);
    (2) deletes every row (24-25); (3) replaces `app.selected_clients` with a new empty dict (27); (4) for each
    `client_key, client_info` in `app.client_data` (insertion order = order in which the scan first met each client;
    the scan sorts files by path, utils/helpers.py:361): builds `missing` = the missing file types joined by
    ", " or the text `'None'` (31); `extra` = number of "Multiple ..." notes as text, or `'0'` (32); inserts a row
    with checkbox text `'☐'` and values (client name, state *name*, `"<status_icon> <status>"`, file count, missing,
    extra, `'☐ No'`) (34-44); records the row as unselected (46); **resets that client's per-client folder-name setting
    to False** (49); tags the row `complete` if `status == 'Complete'`, else `incomplete` (52-55). (5) Sets the tag
    colours to `GUI_CONFIG['colors']['complete'] = '#E8F5E8'` and `['incomplete'] = '#FFF4E5'` (58-59) - background
    only, foreground untouched. (6) Selects and focuses the first row if any (62-65). Returns None. No exception
    handling (a missing key in `client_info` would raise into the Tk callback). Side effects: rebuilds the table;
    resets `selected_clients` and per-client settings. Callers: gui/main_window.py:230 (delegator), reached from
    gui/handlers/file_handler.py:80 (`self.app.update_client_tree()` after every scan).
  - `on_client_click(self, event)` (67-79): asks the table which region/column was clicked. Region `'tree'` (the
    first, checkbox column) -> `toggle_client_selection(row)` (72-75). Region `'cell'` in column `'#7'` (the 7th data
    column, "📁 Name" / FolderName) -> `toggle_folder_name_setting(row)` (76-79). Anything else does nothing here; the
    handler returns None, so the Treeview's own click behaviour (row highlight) still runs. Caller: bound at
    gui/tabs/validation_tab.py:241 via the delegator gui/main_window.py:232-233.
  - `on_space_key(self, event)` (81-86): toggles the first highlighted row (`selection()[0]`); returns `'break'`
    (stops Tk's default Space handling). With several rows highlighted (Shift/Ctrl-click) only the first toggles.
    Caller: validation_tab.py:242 via main_window.py:235-236.
  - `on_up_key(self, event)` (88-98): moves the highlight to the previous row, focuses it and scrolls it into view;
    returns `'break'` (replaces Tk's default Up). Nothing highlighted -> nothing happens. Caller: validation_tab.py:243
    via main_window.py:238-239.
  - `on_down_key(self, event)` (100-116): same downwards; if nothing is highlighted, highlights the first row
    (110-115). Returns `'break'`. Caller: validation_tab.py:244 via main_window.py:241-242.
  - `toggle_client_selection(self, item)` (118-131): flips `app.selected_clients[item]` (default False), sets the
    checkbox text to `'☑'` or `'☐'`, and puts "Selected N clients" in the status bar (N = number of True values).
    Callers: client_handler.py:75 and :85. The delegator gui/main_window.py:244-245 has **NO CALLERS FOUND**.
  - `select_all_clients(self)` (133-138): sets every row True / `'☑'`; status `"Selected all {len(app.client_data)} clients"`
    (counts the scan dictionary, not the rows - same number in practice). Caller: button at validation_tab.py:175
    via main_window.py:247-248.
  - `clear_all_clients(self)` (140-145): every row False / `'☐'`; status "Cleared all selections". Caller:
    validation_tab.py:176 via main_window.py:250-251.
  - `select_complete_clients(self)` (147-159): for each row, if the text `'Complete'` occurs in the Status column
    (`values[2]`, e.g. "✅ Complete") -> True/`'☑'`, else False/`'☐'` (it *replaces* the selection rather than
    adding to it); status "Selected N complete clients". Caller: validation_tab.py:177 via main_window.py:253-254.
  - `get_selected_clients(self)` (161-173): for every row marked True, reads the row's *displayed values*,
    converts the state name in column 2 to a code with `get_state_code(values[1])` and returns the list of
    `f"{values[0]}-{state_code}"` strings, in table order. Return: list of strings meant to equal the keys the
    scanner built (`f"{client}-{get_state_code(state)}"`, core/file_parser.py:229-230). No exception handling.
    Callers: gui/handlers/processing_handler.py:43 (dry run), :81 (start), client_handler.py:227 (export).
    **Silent/visible bug (RAN exp2, HIGH):** Tk's table returns any value that looks like a whole number as a Python
    number (tkinter/ttk.py `_convert_stringval` does `int(str(value))`, pure Python, so the same on Windows). So:
    a client name "007" comes back as `7` and the rebuilt key `"7-GA"` does not exist in `app.client_data`
    (later `KeyError`); a digit-only state field (for example "07") comes back as `7` and `get_state_code(7)` raises
    `AttributeError: 'int' object has no attribute 'strip'` (utils/helpers.py:601) - Dry Run, Start and Export then
    do nothing visible (the error goes to the console, which is hidden when started without one). A client "1234"
    survives here (`"1234-GA"` matches) but fails in `toggle_folder_name_setting` (below). Also `int()` ignores
    surrounding spaces and underscores ("1_000" -> 1000), though the scanner strips spaces first.
  - `toggle_folder_name_setting(self, item)` (175-209): rebuilds the key the same way as above (181); flips
    `app.client_folder_settings[key]` (184-185). If now ON: column 7 becomes `'☑ Yes'` (189); if the client name is
    longer than **10** characters (hard-coded, line 191) asks "Long Client Name Warning" / "Client name '<name>' has
    N characters.\n\nThis will create long folder names. Continue?" and on No reverts to False / `'☐ No'` (192-199).
    If now OFF: `'☐ No'` (201) and, if the global checkbox is ON, switches the global OFF and shows "Global Setting
    Disabled" / "Global folder name setting has been turned off\nsince you unchecked an individual client." (203-207).
    Finally writes the row values back (209). No exception handling, no cache save after switching the global off.
    Caller: client_handler.py:79. Delegator gui/main_window.py:256-257 has **NO CALLERS FOUND**.
    **Bug (RAN exp2):** for a digit-only name, `len(values[0])` at 191 raises `TypeError` *after* the setting was
    flipped at 185, so the setting becomes ON while the table still shows `'☐ No'`. For "007" the setting is stored
    under the wrong key `"7-GA"`.
  - `update_global_folder_setting(self)` (211-223): command of the Setup checkbox "Include client name in Level 4
    folder names" (gui/tabs/setup_tab.py:421 via main_window.py:259-260). If the global is ON, sets every row's
    setting True and shows `'☑ Yes'` (213-221). If OFF, deliberately changes nothing (comment line 222) - so after
    ON then OFF every per-client setting stays True and names are still added to folders (READ, HIGH; consequence of
    core/file_organizer.py:124-130). Always saves the cache (223).
  - `export_client_list(self)` (225-257): gets the selection (227); none -> "No Selection" / "Please select clients
    to export" (228-230). Save dialog: default extension `.txt`, filters "Text files (*.txt)" / "All files (*.*)",
    suggested name `selected_clients_YYYYmmdd_HHMMSS.txt` (233-237), no starting folder. If a name is chosen: opens
    the file with `open(filename, 'w')` - **no encoding given**, so on Windows the system code page is used and a
    client name with characters outside it would raise `UnicodeEncodeError` (caught, but the header lines already
    written stay in a half-written file) (MEDIUM, READ; not run on Windows). Writes "Selected Clients List",
    "Generated: YYYY-mm-dd HH:MM:SS", "Total: N clients", 50 "=" signs, a blank line, then one "<client> - <state>"
    line per selected client (241-249). Shows "Export Complete" / "Exported N clients to:\n<file name only>" (251-252)
    and opens the file with `os.startfile` (255). Any exception -> "Export Error" / "Failed to export: <error>"
    (256-257). A key from the "007" bug raises `KeyError` at 248 and is shown as "Failed to export: '7-GA'".
    Caller: validation_tab.py:130 via main_window.py:262-263. The export file contains client names and is written
    wherever the user chooses (confidentiality note only).
- **Business rules in this file:**
  - [GST RULE] state-code logic: the client key is "<client name>-<state code>", where the state code comes from
    `get_state_code` (lines 169-170, 181, 218) using the state-name table in utils/constants.py:251-292 (letter
    abbreviations such as 'telangana' -> 'TS', 'odisha'/'orissa' -> 'OD', 'dadra and nagar haveli and daman and diu'
    -> 'DD', 'ladakh' -> 'LA') and, for an unknown name, a fallback: first letters of the first two words, or the
    first three characters upper-cased (utils/helpers.py:610-618). TIME-VARYING (the list of states/union territories
    and their names). Described only.
  - "Complete" classification is consumed here (tags 52-55; "Complete Only" 147-159). It is decided in
    core/file_parser.py:252-286: Complete = all six expected file types present and no type appearing twice, except
    'GSTR-3B Export', which may appear any number of times. [GST RULE] (which returns are expected for a client;
    GSTR-3B treated as a per-period return). TIME-VARYING (the expected list in utils/constants.py:68-75 includes
    'IMS Reco').
  - UI rule: per-client long-name warning above 10 characters (191), independent of the configurable limit.
- **Edge cases handled:** no table yet (20-21); no rows -> no first selection (63); nothing highlighted for Space/Up
  (83-86, 90-98) and Down selects the first row (110-115); empty selection on export (228-230); export file errors
  shown (256-257).
- **Gaps, likely failures, silent bugs:**
  - Client keys rebuilt from displayed text (see `get_selected_clients`) - RAN exp2, HIGH. The row id (`item`) or a
    hidden column could have carried the real key.
  - After any scan, every per-client setting is reset to False and every row shows `'☐ No'` even when the global
    checkbox is ON (49, 43). The run then *does* add names to all folders (global wins, core/file_organizer.py:125-127)
    - display contradicts outcome (RAN exp1 `globalname`: table "☐ No" for both clients, folders created as
    "GSTR3B (ClientA)", "ITC (ClientA)", "Sales (ClientA)").
  - Rescanning in dark mode: line 58-59 re-applies light row backgrounds but leaves the dark-mode white text
    (set by gui/utils/dark_mode_manager.py:284-287), giving white text on pale green/amber rows (RAN exp1 `darkscan`:
    tag 'complete' background '#E8F5E8', foreground '#ffffff').
  - Hard-coded 10-character warning (191) vs configurable 35 default elsewhere; the Start-time warning uses the
    configurable value (processing_handler.py:92-105).
  - Switching the global OFF by un-ticking one client (204) is not saved to the cache until some other save happens.
  - Threads: all methods run on the UI thread; but `app.client_folder_settings` (the dict changed by 185/198/219) is
    passed *by reference* into the running FileOrganizer (processing_handler.py:193), so toggling a client while
    Step 3 runs changes the folder names of clients not yet processed (READ, MEDIUM-HIGH).
  - No closures in loops; no mutable default arguments; no `is`/`==` misuse found.
  - Windows-only: `os.startfile` (255).
- **Excel workbook touchpoints:** none.
- **Files read/written:** export text file at the user-chosen path (241), format described above; nothing else.
- **UI elements driven here:** the Step 2 table rows (checkbox text `'☐'`/`'☑'`, Name column `'☐ No'`/`'☑ Yes'`),
  row tags `complete`/`incomplete`, status-bar texts quoted above, dialogs quoted above.
- **Comments/docstrings:** comments at 178-180 and 217 (key by state CODE) agree with the code after 33d0baf.
  Comment 222 agrees with the code. No TODO/FIXME, no commented-out code.
- **Pattern notes:** no `logging` at all in this file; errors are either shown in a message box (export only) or
  left to propagate into Tk (printed to the console, invisible when there is no console). Delegation style: every
  public method is reached through a one-line wrapper on the main window.
- **History:** edac6a3 created it; 1c48722 (29-10-2025, shared with B) removed a block in `update_client_tree` that
  re-tagged rows and re-set every cell to force black text in light mode; 33d0baf (23-07-2026, A only) changed the
  two folder-setting keys from state *name* to state *code* (READ: diff 47a1492..A/main).

### gui/handlers/file_handler.py (151 lines) - [UI]

- **Purpose:** the four Browse dialogs, Scan / Re-scan, and the two input checks used before scanning and before
  processing. (READ)
- **Imports:** standard library `logging` (4), `pathlib.Path` (5), `tkinter.filedialog`, `tkinter.messagebox` (6);
  project `utils.constants.MESSAGES` (7), `utils.helpers.truncate_path` (8). No OS-specific modules.
- **Module-level globals:** `logger = logging.getLogger(__name__)` (10) - **never used** in this file (READ).
- **Class `FileHandler` (13-151).** Instantiated at gui/main_window.py:127.
  - `__init__` (16-17): stores the app. Caller main_window.py:127.
  - `browse_source_folder` (19-25): folder dialog titled "Select Source Folder" (no starting folder); if a folder is
    chosen: sets `app.source_folder`, status "Source: <path shortened to 60 chars>" (`truncate_path`), saves the cache.
    Cancel -> nothing. Caller: Browse button gui/tabs/setup_tab.py:360/384 via main_window.py:210-211.
  - `browse_itc_template` (27-36): file dialog "Select ITC Template", filter "Excel Template" = `*.xltx *.xlsx`
    only; sets `app.itc_template`, status "ITC template selected", saves cache. Caller setup_tab.py:362 via
    main_window.py:213-214.
  - `browse_sales_template` (38-47): same with "Select Sales Template", status "Sales template selected". Caller
    setup_tab.py:364 via main_window.py:216-217.
  - `browse_target_folder` (49-55): folder dialog "Select Target Folder"; status "Target: <shortened>"; saves cache.
    Caller setup_tab.py:366 via main_window.py:219-220.
  - `scan_files` (57-95). Steps: (1) `validate_scan_inputs()`; stop if False (59-60). (2) In a `try`: status
    "Scanning files...", log "🔍 Starting file scan..." (63-64). (3) Defines `progress_callback(current, total,
    message)` that sets status "Scanning: <message> (<current>/<total>)" (67-68). (4) Calls
    `app.file_parser.scan_folder(source, progress_callback)` and stores its three results in `app.scanned_files`,
    `app.client_data`, `app.variations` (71-72). (5) Status "🔍 Found N Excel files" (75-76, overwritten moments
    later). (6) `app.update_summary()` and `app.update_client_tree()` (79-80). (7) If at least one client: status and
    log "Found N clients" (success) and switch to Step 2 (`notebook.select(1)`) (82-88); else status "No valid GST
    files found" and log warning "⚠️ No valid GST files found" (89-91). (8) Any exception -> log "❌ Scan error:
    <e>" (error) and message box "Scan Error" with the error text (93-95). Return: None.
    Callers: Scan button setup_tab.py:459-460 via main_window.py:222-223; file_handler.py:118 (rescan).
    - **Runs on the UI thread** (READ, HIGH): the whole scan happens inside the button's callback; nothing calls
      `update()`, so the window does not repaint or respond until the scan ends ("Not Responding" on a large
      network share - commit 4f84ba9's message speaks of 11 min -> 41 s on the real share). The status texts from
      `progress_callback` are set but never painted while the scan runs; the log line from step 2 is only scheduled
      (`after(0)`) and appears afterwards.
    - Files that fail the size (>= 1024 bytes) or signature check are silently excluded and not counted
      (utils/helpers.py:290-306, 275-287); a listing error in `scan_excel_files` is swallowed and returns an empty
      list (utils/helpers.py:363-365), so an unreachable share shows "No valid GST files found" rather than an error (READ,
      HIGH; A-R1's file). Only the top level of the source folder is read (`os.scandir`, not recursive,
      utils/helpers.py:328-339).
    - Names of files that could not be parsed are never shown in the GUI - only their count in the summary
      (gui/main_window.py:306, 320-321); they appear in the summary workbook's Variations sheet after processing
      and in the log file (core/file_parser.py:131).
  - `rescan_files` (97-118): log "🔄 Re-scanning files..."; empties `app.scanned_files`, `app.client_data`,
    `app.selected_clients`, `app.variations` **in place** (102-105); clears the summary text and table (108-115);
    then `scan_files()` (118). If the source check then fails, the previous results are already gone. Caller:
    Re-scan button setup_tab.py:464-465 via main_window.py:225-226. **Mid-run hazard (RAN exp8):** Re-scan is not
    disabled during Step 3; emptying `app.variations` in place also empties the list the running job already handed
    to the summary workbook (processing_handler.py:224), so the workbook's Variations sheet said "No file variations
    found" although the scan had found one unparseable file.
  - `validate_scan_inputs` (120-130): empty source -> "Error" / "Please select a source folder containing GST files"
    (`MESSAGES['errors']['no_source']`, utils/constants.py:210); path not existing -> "Error" / "Source folder does
    not exist". Returns True otherwise. Does not check it is a folder (a file path passes here and fails in
    core/file_parser.py:154-155 as "Scan Error: Not a directory: ..."). Caller: file_handler.py:59.
  - `validate_processing_inputs` (132-151): either template empty -> "Please select both ITC and Sales template
    files"; target empty -> "Please select a target folder for organized files"; ITC template missing -> "ITC
    template file not found"; Sales template missing -> "Sales template file not found" (all titled "Error").
    Returns True otherwise. **Does not check:** that the target folder exists (FileOrganizer silently creates it,
    core/file_organizer.py:66-69 -> utils/helpers.py:45), the template file type, that the source folder still
    exists, or that the two templates differ. Callers: processing_handler.py:40 (dry run), :78 (start).
- **Business rules:** none of GST nature. UI rule: template file filter `*.xltx *.xlsx` (31, 42).
- **Edge cases handled:** cancelled dialogs (22, 33, 44, 52); empty/missing source (122-128); scan exceptions shown
  (93-95); no clients found (89-91); missing template files (143-149).
- **Gaps:** UI-thread scan (above); in-place clearing during a run (above); no is-a-folder check; validation weaker
  than the Setup checklist (see Cross-file D); Browse dialogs never start in the currently chosen folder.
  `save_cache` after each Browse raises if the max-length spinbox holds text that is not a number (RAN exp5:
  `TclError: expected floating-point number but got ""`), so the Browse result is set but the exception escapes into
  Tk and the cache is not written.
- **Excel touchpoints:** none directly (scan reads the first 8 bytes of each file - utils/helpers.py:281-284).
- **Files read/written:** reads the source folder listing (via core); writes the cache via `app.save_cache()`
  (25, 36, 47, 55).
- **UI elements:** dialogs and status/log texts quoted above; switches to Step 2 after a successful scan (88).
- **Comments:** "Show count while scanning" (74) is inaccurate - the count is shown after the scan finished.
- **Pattern notes:** errors to the user by message box; logs to the on-screen log, not to the log file (the module
  logger is unused).
- **History:** edac6a3; 6237597 (29-10-2025, shared) fixed line 75 to read `app.scanned_files` instead of
  `app.file_parser.scanned_files`. No A-only changes.

### gui/handlers/processing_handler.py (530 lines) - [UI] (it also orchestrates [CORE] work)

- **Purpose:** Dry Run, Start/Stop, the background job that organises files and builds the ITC/Sales reports and the
  summary workbook for each selected client, and the optional hand-over to Step 4. (READ)
- **Imports:** standard library `os` (4), `platform` (5), `logging` (6), `threading` (7), `datetime` (8), `pathlib.Path`
  (9), `tkinter.messagebox` (10); project `utils.constants.MESSAGES` (11), `utils.helpers.get_timestamp,
  sanitize_filename, get_state_code, ProgressTracker` (12). A lazy `import traceback` inside an `except` (405).
  OS-specific: `os.startfile` (508) is **guarded** by `platform.system() == 'Windows'` (507); macOS/Linux branches use
  `os.system('open "<folder>"')` / `os.system('xdg-open "<folder>"')` (510, 512).
- **Module-level globals:** `logger = logging.getLogger(__name__)` (14).
- **Operational constants inside the code:** progress split - the first half of each client's share of the bar for
  copying, the second half for reports (267, 287); summary step at 95% (459); finish at 100% (482). Long-name
  fallback limit 10 (95, 97). No waits, retries or timeouts in this file (the Excel waits are in core/excel_handler.py:155,210,
  0.5 s each; the Power Query wait time belongs to Step 4, see Cross-file A).
- **Class `ProcessingHandler` (17-530).** Instantiated at gui/main_window.py:129.
  - `__init__` (20-21): stores the app.
  - `_log_async(self, message, level='normal')` (23-31): schedules `app.log_message(message, level)` on the UI thread
    with `root.after(0, ...)`, capturing the already-built text through default arguments. Callers: 51-53, 57, 61,
    63-64, 66-67, 126, 146-148, 157, 244, 258, 261-263, 279, 284, 300-302. (Docstring: fixes wrong client names
    caused by late binding - agrees with commit 0890022.)
  - `_progress_async(self, progress, message)` (33-36): same for `app.update_progress`. Callers: 241, 252, 269, 289.
    Note (READ): both helpers add one extra hop - `app.log_message` and `app.update_progress` already schedule
    themselves with `after(0)` (gui/utils/ui_helpers.py:23; gui/main_window.py:289). RAN exp6: 200 back-to-back pairs
    "helper message then direct message" from a worker thread kept their order on Linux/Tk 9 (0 inversions), so the
    mix of two- and one-hop messages did not reorder the log there; not tested on Windows (LOW risk).
  - `dry_run(self)` (38-67): requires `validate_processing_inputs()` (templates and target, though a dry run uses
    neither) (40-41); gets the selection, none -> "Warning" / "Please select at least one client to process"
    (43-46); switches to Step 3 (49); logs "🧪 ════════ DRY RUN PREVIEW ════════", "🎯 Mode: <fresh|rerun|resume>",
    "👥 Selected: N clients", then per client "\n🏢 <client> - <state>", one "  📄 Would organize: <file name>" per
    file, "  📊 Would create: ITC Report", "  💰 Would create: Sales Report", then "\n🧪 ════════ DRY RUN COMPLETE
    ════════" and "💡 No files were actually moved or created" (51-67). Shows no destination folders, no workflow
    mode, no client-name setting. No `is_processing` check, so it can run during a real run and interleave with its
    log. `app.client_data[client_key]` (56) raises `KeyError` for the "007" case (READ; in RAN exp2, with all four
    rows ticked, the dry run stopped even earlier with the `AttributeError` of the "07" state case). Caller: DRY RUN button validation_tab.py:118 via main_window.py:266-267.
  - `start_processing(self)` (69-134). Steps: (1) **re-entrancy guard** - return if `app.is_processing` (75-76)
    (RAN exp1: a second click on the disabled Start button during a run started no second thread);
    (2) `validate_processing_inputs()` (78-79); (3) selection, none -> warning (81-84); (4) only if "Include client
    name" is ON: limit = `int(client_name_max_length.get())`, `<= 0` or any error -> **10** (92-97); lists names
    longer than the limit (up to 5 shown) and asks "Warning" / "Long client names detected (exceeding N chars):\n\n•
    <name> (L chars)...\n\nIf you continue, these client names will be used in folder names.\n\nContinue anyway?"
    - No cancels (99-113); (5) switches to Step 3 (116); Start button `disabled`, Stop `normal` (119-120);
    `stop_requested = False`, `is_processing = True` (123-124); logs "🚀 ════════ PROCESSING STARTED ════════"
    (126); (6) starts a **daemon** thread running `process_files_thread(selected)` and stores it in
    `app.processing_thread` (129-134; never read or joined anywhere - READ grep). Callers: START button
    validation_tab.py:142 and START PROCESSING button processing_tab.py:74, both via main_window.py:269-270; tests
    tests/test_processing_handler.py:27, :35.
  - `_maybe_continue_to_extraction(self)` (136-159): runs **on the worker thread** at the end of a successful run
    (519). Reads `app.workflow_mode` (a Tk variable read from the worker); any error -> silent return (150-151). Not
    'full' -> logs "ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query." and returns
    (145-149). No `extract_tab` attribute -> silent return (153-155). Else logs "\n🔄 Full pipeline: moving to Step 4
    to extract…" and schedules `_open_extraction_tab` on the UI thread (157-159). Callers: processing_handler.py:519;
    tests/test_workflow_mode.py:115-155.
  - `_open_extraction_tab(self)` (161-175), UI thread: selects Step 4 (164); takes `app.extract_tab.panel` (165);
    if the Step 1 target is set, writes it into the panel's folder box (166-168) - **overriding a folder the user may
    have chosen in Step 4**, and without updating ExtractTab's `_pushed_folder` bookkeeping
    (gui/tabs/extract_tab.py:41-43, 121-131); runs the panel's `scan_folder()` (169); logs success "✅ Step 4 ready -
    review the client list and press Process to refresh." (170-172). Any exception -> `logger.error(..., exc_info)`
    and log warning "Could not open Step 4 automatically: <e>" (173-175). If the extractor could not be imported,
    `panel` is None and the message is the technical "'NoneType' object has no attribute 'folder_path'" (READ, HIGH;
    the guard at 153-155 tests the tab, not the panel). The success line is logged even if the panel's scan found no
    "Annual Statement-*" folder (that scan reports into its own log and returns normally,
    power_query_extractor/gui/extractor_window.py:625-628). The text says "press Process" but the Step 4 button reads
    "🚀 Start Processing" (extractor_window.py:330). RAN exp1 `full`: Step 4 selected, folder = target, 2 clients
    listed, this log line shown. Caller: scheduled at 159.
  - `stop_processing(self)` (177-180): sets `app.stop_requested = True` and logs "🛑 Stop requested by user"
    (warning). Nothing else (button stays enabled, no "will stop after this client" text). Caller: STOP button
    processing_tab.py:87 via main_window.py:272-273. RAN exp3: clicking the *disabled* Stop button while idle still
    ran this (log showed the line, flag True) because the button uses `.bind('<Button-1>')`; harmless because
    start_processing resets the flag (123), but misleading.
  - `process_files_thread(self, selected_clients)` (182-530) - the worker. Step by step:
    1. `start_time = datetime.now()`; `timestamp = get_timestamp()` = "DDMMYY HHMM" (185-186; utils/helpers.py:25-27).
       This is a *run time*, not a tax period.
    2. Creates `app.file_organizer = app.FileOrganizer(target, processing_mode, include_client_name=<global>,
       client_folder_settings=<the live dict>, client_name_max_length=<IntVar value>)` (189-195). The FileOrganizer
       takes its *own* `get_timestamp()` (core/file_organizer.py:43): if the minute changes between line 186 and that
       call, folder names and report/summary file names carry different minutes (READ, LOW likelihood).
       Constructor errors (invalid mode; target cannot be created) jump to the outer `except` (522).
    3. Counters `processed`, `total_files`, `total_reports` (198-200); `ProgressTracker(total_clients)` (203);
       `level1_folder = None` (206); `summary_data` dict with timestamp, counts, mode, target, empty lists, a
       *reference* to `app.variations`, include-name flag (209-226).
    4. For each selected key (228): stop if `stop_requested` (229-230) - checked **only here**, so the current client
       always finishes, including slow Excel work. `client_info = app.client_data[client_key]` (232) is **outside**
       the per-client `try`: a missing key (the "007" case) ends the whole run in the outer `except` with "💥 Fatal
       error: '7-GA'" and a "Processing Error" box showing just the key; later clients are not processed and no summary
       workbook is written (READ, HIGH). Progress: tracker update (235; first client adds 0, later ones 1), ETA text
       "Calculating..." for the first client (236), bar `idx/total*100` with "Client i of n: <name> | ETA: <eta>" (239-241).
    5. Per-client `try` (243-453):
       - log "\n🏢 Processing <client> - <state>"; status "Client i/n: Creating folders | ETA: …" (244-252);
       - `folders = file_organizer.create_client_structure(client_info)` (253) - creates Level 1 / Level 2 / Version /
         3 category folders (core/file_organizer.py:71-148); first client's Level 1 kept (255-256);
       - logs "  ✓ Created folder structure", "    Level 1: <name>", "    Level 2: <name>", "    Level 3: <Version
         folder name>" (258-263);
       - nested `file_progress(current, total, message)` (266-269) moves the bar inside the first half of this
         client's share; called synchronously by `organize_files`, so its references to loop variables are current
         (no late binding);
       - `file_results = file_organizer.organize_files(client_info, folders, file_progress)` (271-273) copies every
         file (never moves; core/file_organizer.py:290, utils/helpers.py:199) and replaces the stored file name with
         the cleaned copy name;
       - counts successes, logs "  ✓ Organized N files" and **one line per successful file only** ("    📁 <name> →
         <type> folder") (275-284). Failed or skipped copies get no line in the log (READ, HIGH);
       - report names: `safe_client = sanitize_filename(client)`; `safe_state = client_info.get('state_code',
         get_state_code(state))`; `safe_timestamp = timestamp.replace(' ', '_')` (292-294);
       - **ITC report** (296-326): name `ITC_Report_<safe_client>_<state code>_<DDMMYY_HHMM>.xlsx` in the Version
         folder (297-298); logs "  📊 Creating ITC report...", "    Template: <template file name>" (template path
         **re-read from the Tk variable for every client**), "    Output: <name>" (300-302); `prepare_template_data(...,
         'ITC')` + `create_report_from_template(itc_template.get(), output, data, 'ITC')` (305-314); True -> counts and
         "    ✅ Success: <name>.xlsx"; False -> `report_errors += 1`, "    ❌ Failed to create ITC report", "    Check
         if template has 'Links' sheet" (warning); exception -> "    ❌ ITC Report Error: <e>" (315-326);
       - **Sales report** (328-411): "  💰 Starting Sales report creation..."; 14 lines tagged "[DEBUG]" are shown
         to the user (333, 342, 353, 362, 363, 367, 371, 374, 391, 393, 398, 403, 404, 411; commit 064c634 left them on
         purpose as "a separate UX decision"); empty template path -> error (335-337); file missing -> error
         (341-346); name `Sales_Report_<safe_client>_<state code>_<DDMMYY_HHMM>.xlsx` (355-356); extra existence checks
         of the output folder and the written file (363, 389-393 - extra network round-trips per client); same
         True/False/exception handling as ITC, plus a full traceback to the log file (405-406) and an outer "Sales
         Report Setup Error" catch (408-411). ITC and Sales code paths are not symmetrical (READ).
       - `file_organizer.create_organization_report(folders, client_info)` (414) writes
         `_Organization_Report_<client>_<state>.txt` into the Version folder with the *raw* (unsanitised) client and
         state names (core/file_organizer.py:330); its return value (None on failure) is ignored.
       - `processed += 1`, `successful_clients += 1` (416-417); appends a client record with status **'Success'**
         regardless of whether either report was created (420-429; `itc_report_created`/`sales_report_created` flags
         are recorded but the summary workbook never writes them - core/excel_handler.py:464-491 writes Client, State,
         Status, Files, Missing count, Complete %); appends one file-mapping row per copy result (431-438); logs
         "  ✅ Completed <client>" (440).
       - any exception -> `failed_clients += 1`, an error record (time, client key, 'Processing Error', message), log
         "  ❌ Failed <client key>: <error>" (442-453). Failed clients get no Client Status row.
    6. After the loop, only if `processed > 0` and a Level 1 folder exists (456): progress 95 "Creating summary report |
       ETA: <eta>" (the ETA then shows roughly one client's time although all clients are done - READ, LOW); log
       "\n📊 Creating Level 1 summary report..." and "  Output: GST_Processing_Summary_<DDMMYY_HHMM>.xlsx"; calls
       `create_summary_report(level1/…, summary_data)`; success "  ✅ Success: …" + "  Sheets: Summary, Client Status,
       File Mapping, Errors, Variations"; failure "  ❌ Failed to create summary report"; exception "  ❌ Summary Report
       Error: <e>" (456-475). **If every client failed, no workbook is written and the error list exists only in the
       on-screen log** (READ, HIGH).
    7. Finish: progress 100 "Processing complete | Total time: Xm Ys" (478-482); log "\n🎉 ════════ PROCESSING COMPLETE
       ════════", "✅ Processed: p/t clients", "📁 Files organized: n", "📊 Reports created: r", "📂 Output location:
       <Level 1>" (if any), "\n⏱️ Total processing time: X minutes Y seconds" (484-496) - two separate time
       calculations (tracker and datetime) for the same thing.
    8. `messagebox.showinfo("Processing Complete", "Successfully processed p/t clients\n\nFiles organized in:\n<target
       folder>")` (498-502) **called from the worker thread** (RAN exp1: recorder saw thread "Thread-1
       (process_files_thread)"). It blocks the worker until OK is pressed, so `is_processing` stays True, the Start
       button stays disabled, the output folder is not yet opened and the Step 4 hand-over waits until then. The text
       says "Successfully" even after a Stop or failures (RAN exp1 `stopfirst`: "Successfully processed 0/2 clients",
       bar 100%, log "PROCESSING COMPLETE"), and names the target folder, not the new Level 1 folder.
    9. Opens the Level 1 folder (505-516): Windows `os.startfile`; macOS `open`; other `xdg-open` via `os.system`,
       whose result is ignored, so "📂 Opened output folder" is logged even when it failed (RAN: "sh: 1: xdg-open: not
       found" yet the success line appeared). The folder name is put inside double quotes in a shell command (510, 512).
    10. `_maybe_continue_to_extraction()` (519) - also after a Stop.
    11. Outer `except` (522-525): log "💥 Fatal error: <e>", `logger.error(..., exc_info=True)`, message box
        "Processing Error" from the worker thread.
    12. `finally` (527-530): `is_processing = False`, Start `normal`, Stop `disabled` - widget changes made **from the
        worker thread**. If the window is gone these calls raise; the exception escapes the thread (RAN exp4 below).
    Callers: thread target at processing_handler.py:130. The delegator gui/main_window.py:275-276 has **NO CALLERS
    FOUND**.
- **Business rules in this file:**
  - [GST RULE] state-code logic in report file names: `safe_state = client_info.get('state_code',
    get_state_code(client_info['state']))` (293); state code from utils/constants.py:251-292 (TIME-VARYING list).
  - [GST RULE] (which return files feed which report/sheet/cell): every client always gets one "ITC" and one "Sales"
    report attempt (296-411); `prepare_template_data(…, 'ITC'|'Sales')` (305-307, 368-370) and the Links-sheet cell map
    in utils/constants.py:102-126 decide that the ITC report's Links sheet receives the GSTR-3B folder (B2), Annual
    Report folder/name (B4/B5), GSTR-2B Reco folder/name (B7/B8), IMS Reco folder/name (B10/B11), and the Sales report's
    Links sheet receives Sales folder/name (B2/B3), Annual Report folder/name (B5/B6), Sales Reco folder/name (B8/B9).
    Only the first file of each type is used (core/excel_handler.py:603-712). Described only; owned by A-R1's files.
    TIME-VARYING (return types, e.g. IMS).
  - Report naming: `ITC_Report_…`, `Sales_Report_…`, `GST_Processing_Summary_…` with DDMMYY_HHMM run time (297, 355,
    463). The Step 4 scanner finds reports by the prefixes `itc_report_` / `sales_report_` (extractor_window.py:36-39)
    - write and read sides agree (READ, HIGH).
  - Long-name warning threshold: configurable value, fallback 10 (92-105).
- **Edge cases handled:** re-entry (75-76); no selection (81-84, 43-46); invalid max-length value (92-97); empty or
  missing Sales template (335-346); ITC/Sales report failures counted separately (321, 337, 346, 395, 401, 409); per-client
  failures don't stop the run (442-453); no successful client -> no summary (456); summary failure caught (468-475);
  folder-open failure on Windows caught (515-516); hand-over failures caught (150-151, 173-175).
- **Gaps, likely failures, silent bugs:**
  - **Tk touched from the worker thread** (READ, HIGH; RAN on Linux): Tk-variable reads 190-194, 219-220, 225, 301, 310,
    332, 145, 501; `after()` scheduling from the worker 31, 36, 159 and inside `log_message`/`update_progress`;
    message boxes 498, 525; widget configuration 529-530. Tkinter forwards such calls to the main thread when Tcl is
    built with threads (the python.org Windows builds are, as far as I know - MEDIUM-HIGH, not checked on Windows;
    the Linux Tcl 9 used here is), but they fail with "RuntimeError: main thread is not in
    main loop" if the main loop is not running (RAN: exp1 first attempt and exp4).
  - **Closing the window during a run** (RAN exp4, HIGH): `on_closing` sets the stop flag and destroys the window
    after 1 s (gui/main_window.py:352-353). The worker was still inside the current client; it wrote that client's ITC
    report after the window had gone, then its next Tk call raised "main thread is not in main loop"; the `except`
    tried to log and failed again; the `finally` failed at line 529; the thread died with a chain of tracebacks. That
    client's Sales report, organisation report and the summary workbook were never written. In the real app,
    `main.py:112` returns once the window is destroyed, the interpreter exits and the **daemon** worker is killed
    wherever it is - possibly mid-copy or mid-Excel (MEDIUM for the Windows consequences: an invisible Excel left
    running, a staging file left in %TEMP%\gst_excel_temp; not run on Windows).
  - **Inputs re-read live during the run** (RAN exp8, HIGH): the ITC template path is read per client (301, 310) and the
    Sales template per client (332); changing either in Step 1 mid-run changed the template for the next client
    (log showed "Template: itc_template_OTHER.xlsx" for client 2). The per-client folder-name dict is live (193). The
    processing mode and global name flag are captured once (191-192).
  - **Nothing is locked during a run**: Setup Scan/Re-scan (enabled purely from the checklist,
    gui/tabs/setup_tab.py:585-591), Browse, the Validation buttons, Dry Run, and Step 4's own Start Processing remain
    usable. Step 3 and Step 4 can drive Excel at the same time; there is no cross-check of `is_processing` either way
    (READ, HIGH; Excel consequences MEDIUM - core/excel_handler.py:124 attaches to an already running Excel with
    `GetActiveObject` and quits it in `finally` 238-246 because `self.excel_app` is never set; report_processor uses
    `Dispatch`, power_query_extractor/core/report_processor.py:218, 536 - A-R1/A-R2 to confirm).
  - Stop is coarse and unreported (RAN exp1 `stopfirst`): no "stopped" wording anywhere; the hand-over to Step 4 still
    happens in 'full' mode after a Stop (519).
  - Success counting: a client whose two reports both failed is still "Success" and counted in "Successfully processed
    p/t" (417-429, 500).
  - Failed copies are not logged (282-284); the reports are still built and their Links cells then point at original
    (uncleaned) names that were never copied (core/file_organizer.py:233-246 only renames on success) - MEDIUM.
  - On the openpyxl fallback (no pywin32), a template without a 'Links' sheet is saved unchanged and reported as
    "✅ Success" (core/excel_handler.py:295-310), so the "Check if template has 'Links' sheet" hint never appears in
    that case (READ, HIGH; A-R1's file). On Windows with pywin32 the COM path adds a 'Links' sheet (excel_handler.py:180-183).
  - Two runs in the same minute (fresh mode) share the same Level 1/Version folders; copies of existing files are
    backed up first (core/file_organizer.py:277-282) but reports and the summary are overwritten without backup
    (READ, MEDIUM-HIGH). Rerun/resume add a new Version folder and a new summary workbook to the latest Level 1
    folder (RAN exp7: three GST_Processing_Summary files in one Level 1 folder).
  - Two copies of the app at once: shared cache file; shared local staging file names in %TEMP%\gst_excel_temp
    (core/excel_handler.py:147) when both process the same client in the same minute; shared Excel through COM
    (MEDIUM, READ).
  - Very large runs: the log keeps growing (one line per copied file) and is never cleared between runs.
  - Network share drop mid-run: per-client exceptions are logged and the run continues; per-file copy failures are
    silent; the summary is written at the end to the share - if it is unreachable then, only the on-screen log
    remains (READ, MEDIUM).
  - No closures created in loops capture loop variables by reference any more (fixed in 0890022); `file_progress`
    (266-269) is called synchronously. No mutable default arguments. `level1_folder` after the loop is the first
    client's Level 1, which FileOrganizer shares for the whole run (core/file_organizer.py:150-161) - correct.
- **Excel workbook touchpoints (indirect, via core/excel_handler.py):** ITC report created from the ITC template
  (309-314); Sales report from the Sales template (376-381); summary workbook created at Level 1 (469). Copies of the
  source workbooks are byte copies (no Excel). Step 4 (`scan_folder`, 169) only lists folders.
- **Files and folders written (via core):** under the target folder: `Annual Statement-<DDMMYY HHMM>` (fresh) or the
  newest existing `Annual Statement-*` by mtime (rerun/resume) / `<client-key shortened to max length>-<STATE>` /
  `Version-<DDMMYY HHMM>` / `GSTR3B`, `ITC`, `Sales` (or `GSTR3B (<client>)`, `ITC (<client>)`, `Sales (<client>)`);
  Annual Report copies go directly into the Version folder; report files and the organisation text file in the Version
  folder; summary workbook in Level 1 (RAN exp1: full tree listed in exp1 output; exp7). Backups `<name>_backup_<DDMMYY_HHMM><ext>`
  when a destination file already exists in fresh mode. Staging copies in `%TEMP%\gst_excel_temp\<report name>`
  (deleted after each report, best effort). Log file messages via `logger` (the file `gst_organizer.log` is
  configured in main.py:24, relative to the current folder - A-R4's file).
- **UI elements driven here:** Step 3 progress bar/labels, Start/Stop button states, Step 3 log lines (quoted above),
  tab switches to Step 3 (49, 116) and Step 4 (164), dialogs quoted above.
- **Comments/docstrings vs code:** 71-74 (guard) agrees; 86-89 says fallback 10 - agrees with code but conflicts with
  35 elsewhere; `_maybe_continue_to_extraction` docstring "In 'organize' mode nothing happens here" - it does log an
  info line (146-148); 158 "Marshal to the UI thread - this runs on the worker thread" agrees; "Debug checkpoint 1-6"
  comments (331, 340, 361, 366, 373, 388) are leftover debugging; "[DEBUG] Full traceback will be in log file" (404) is
  true only if logging to file is configured (main.py does). No TODO/FIXME, no commented-out code.
- **Pattern notes:** two log routes coexist - `_log_async` (lines up to 302 and 126/146/157) and direct
  `app.log_message` (319 onwards). History explains it: 6237597 wrapped only the first part of the loop in
  `root.after(0, lambda: ...)` (and the dry run), 0890022 replaced exactly those lambdas with the helpers, and the
  Sales part was never converted (READ diffs). User messages go to the on-screen log; technical detail to the module
  logger only in a few places (174, 406, 524). Exceptions are caught broadly (`except Exception`).
- **History:** edac6a3; 6237597 (shared) - after(0) wrapping, long-name check rewritten with the configurable limit
  and fallback 10; A-only: 0890022 (helpers for late binding), 33d0baf (re-entrancy guard), 510f7b4 (workflow
  hand-over to Step 4).

### gui/main_window.py (398 lines) - [UI]

- **Purpose:** builds the application window, owns all shared state (Tk variables, scan results, processing flags),
  creates the four tabs and three handlers, forwards button actions to the handlers, and saves/restores settings. (READ)
- **Imports:**
  - standard library: `os` (7, unused), `sys` (8, unused), `tkinter as tk` (9), `tkinter.ttk`, `tkinter.messagebox`
    (10), `pathlib.Path` (11, unused), `logging` (12).
  - project: core `FileParser` (15), `FileOrganizer` (16), `ExcelHandler` (17); `utils.constants.GUI_CONFIG` (20),
    `MESSAGES` (20, unused); `utils.helpers.get_timestamp` (21, unused); widgets `TitleBar` (24), `CollapsibleFrame`
    (25, unused here); tabs (26-29); handlers (30-32); gui utils `DarkModeManager`, `CacheManager`, `UIHelpers`,
    `StatusBar` (33-36). Unused imports found by an AST scan (READ/RAN): os, sys, Path, MESSAGES, get_timestamp,
    CollapsibleFrame.
  - OS-specific: none directly. Indirectly, `core/excel_handler.py:21-28` imports win32com/pythoncom only on Windows
    inside try/except; `gui/tabs/extract_tab.py:24-30` guards the extractor import.
- **Module-level globals:** `logger = logging.getLogger(__name__)` (38) - never used; line 193 calls
  `logging.getLogger(__name__)` again instead.
- **Class `GSTOrganizerApp` (41-398).** Instantiated at main.py:111 (and in tests via subprocess,
  tests/test_main_window_tabs.py:37, tests/test_optional_pywin32.py:57).
  - `__init__` (46-55): creates the Tk root (48), then in order `setup_window`, `setup_variables`, `setup_components`,
    `setup_handlers`, `create_gui`, `bind_events`, `apply_cached_values` (49-55). Consequence of the order: the tabs are
    built with empty paths and the cached values arrive afterwards through the variables' traces.
  - `setup_window` (57-77): title 'GST File Organizer v3.0', size '1200x800', minimum 1000×700, background
    '#F8F9FA' (59-64); centres the window from the screen size (66-70; RAN: 1200x800+360+140 on a 1920×1080 virtual
    screen); creates `self.style` and switches to the 'clam' theme, ignoring any error with a bare `except: pass`
    (73-77). Caller: 49.
  - `setup_variables` (79-110): `source_folder`, `itc_template`, `sales_template`, `target_folder` (StringVar, empty);
    `client_folder_settings = {}`; `processing_mode` = 'fresh'; `workflow_mode` = 'organize' (comment 88-89);
    `dark_mode` = False; `include_client_name_in_folders` = False; `client_name_max_length` = 35 (IntVar);
    `scanned_files = {}`, `client_data = {}`, `variations = []`, `selected_clients = {}`; `is_processing = False`,
    `stop_requested = False`, `processing_thread = None`; `widgets = {}` (**never used anywhere** - READ grep);
    `self.FileOrganizer = FileOrganizer` (110, the class, so the handler can create one). Caller: 50.
  - `setup_components` (112-123): `FileParser()`, `ExcelHandler()` (creates `%TEMP%\gst_excel_temp` at startup,
    core/excel_handler.py:53-54), `file_organizer = None`; `CacheManager()` and `cache_data = load_cache()` (reads
    `~/.gst_organizer_cache.json`); `DarkModeManager(root, style)` + `initialize()` (records original colours). Caller: 51.
  - `setup_handlers` (125-129): creates FileHandler, ClientHandler, ProcessingHandler with `self`. Caller: 52.
  - `create_gui` (131-149): TitleBar (title strip with the "🌙 Dark Mode" checkbox wired to `toggle_dark_mode`) (134);
    `ttk.Notebook` packed with padding 20 / (10, 20) (137-138); SetupTab, ValidationTab, ProcessingTab, ExtractTab in
    that order (141-146; comment 144-145 "Built without scanning, so startup does not touch the network"); StatusBar
    (149) packed after the notebook at the bottom. Caller: 53.
  - `bind_events` (151-154): window close button -> `on_closing`; notebook tab change -> `on_tab_changed`. Caller: 54.
  - `toggle_dark_mode` (157-207): called by the title-bar checkbox (gui/widgets/title_bar.py:32 via 134) and at startup
    if the cache says dark (389). Steps: apply dark or restore original colours (159-162); in light mode repaint the
    title strip and its direct children with the primary blue, ignoring errors (bare `except: pass`, 165-174);
    recolour the client-table row tags (177-178); recompute the Setup checklist colours with
    `setup_tab.refresh_status(force=True)` (188-194, errors only logged at debug level) - which also clears the
    existence cache and re-checks the four paths on disk; `save_cache()` (196); repaint (197, 206-207). The long comment
    180-187 explains the stale-tint problem (commit c89c1d4); 199-205 explains why the old tab-padding hack was removed
    (725fbaf). Dark mode themes only the widgets that exist at the moment of toggling; widgets created later (Step 2
    rows re-tagged by a scan, Step 4 client rows after a scan) are not themed (RAN for Step 2; READ for Step 4 -
    extractor_window.py:704-753 creates white rows).
  - Delegators (210-276): `browse_source_folder`, `browse_itc_template`, `browse_sales_template`,
    `browse_target_folder`, `scan_files`, `rescan_files` -> FileHandler; `update_client_tree`, `on_client_click`,
    `on_space_key`, `on_up_key`, `on_down_key`, `toggle_client_selection`, `select_all_clients`, `clear_all_clients`,
    `select_complete_clients`, `toggle_folder_name_setting`, `update_global_folder_setting`, `export_client_list` ->
    ClientHandler; `dry_run`, `start_processing`, `stop_processing`, `process_files_thread` -> ProcessingHandler.
    Callers: listed under each handler method. **NO CALLERS FOUND** for `toggle_client_selection` (244-245),
    `toggle_folder_name_setting` (256-257) and `process_files_thread` (275-276) (READ grep of all .py incl. tests).
    The key handlers return the handler's `'break'` (236, 239, 242), the others return None.
  - `update_status(message)` (279-281): status bar text "💡 <message>" (gui/utils/status_bar.py:35-37). Callers: the
    handlers (client_handler.py:131, 138, 145, 159; file_handler.py:24, 35, 46, 54, 63, 68, 76, 84, 90) and 338, 393.
    Called only from the UI thread (READ).
  - `log_message(message, level='normal')` (283-285): passes to `UIHelpers.log_message(self.log_text, …)`, which
    timestamps "[HH:MM:SS] " now and schedules the insert with `after(0)` (gui/utils/ui_helpers.py:17-33). Levels used:
    'normal', 'info', 'success', 'warning', 'error' (tag colours set in processing_tab.py:113-117). Callers: all
    handlers, 394-396, and `_log_async`.
  - `update_progress(value, message)` (287-292): schedules bar value, "📊 Progress: X.X%" and the operation text with
    `after(0)`; captures its own parameters (no loop, no late binding). Callers: processing_handler.py:36, 459, 482.
  - `update_summary` (294-328): returns if no summary box (296-297); takes `file_parser.get_statistics()` and writes a
    fixed text block into the Step 2 summary box (made editable, cleared, filled, made read-only again):
    "📊 SCAN RESULTS SUMMARY:", a rule line, "📁 FILES ANALYSIS:" with "Total Excel Files Found", "✅ Successfully
    Parsed", "❌ Unparsed/Variations", "📊 Parsing Success Rate: X.X%", "👥 CLIENTS ANALYSIS:" with "Total Clients
    Identified", "✅ Clients with All Files", "⚠️ Clients with Missing Files", "📈 Completeness Rate: X.X%", "📋 FILE
    TYPE DISTRIBUTION:" with one "   • <type>: N files" line per type, optional "🔍 FILE VARIATIONS: N files" and
    "💾 TOTAL DATA SIZE: <size>" (RAN exp1: first lines as expected). "Clients with Missing Files" = all clients that
    are not 'Complete', including clients with duplicates only (core/file_parser.py:328). Caller: file_handler.py:79.
  - `on_tab_changed(event)` (331-342): status "Step N: <Setup|Validation|Processing|Extract>" with an out-of-range
    fallback "Step N" (334-338; commit bfcb24a fixed the IndexError when Step 4 was added); repaints if not processing.
    Caller: binding at 154.
  - `on_closing` (344-355): if Step 3 is processing: "Processing Active" / "Processing is still running. Stop and
    exit?" -> Yes: stop flag + destroy after 1000 ms (see exp4 above); No: nothing. Not processing: destroy at once.
    **Does not check Step 4's own `is_processing`** (extractor_window.py:119, 872), so closing during a Step 4
    refresh gives no warning (READ, HIGH). Does not save settings. Caller: binding at 153.
  - `save_cache` (358-371): writes the nine keys `source_folder`, `itc_template`, `sales_template`, `target_folder`,
    `processing_mode`, `workflow_mode`, `include_client_name`, `dark_mode`, `client_name_max_length` through
    CacheManager (which adds `recent_folders`, gui/utils/cache_manager.py:29-48). `client_name_max_length.get()` raises
    if the spinbox holds non-numeric text (RAN exp5). Callers: file_handler.py:25/36/47/55, client_handler.py:223,
    setup_tab.py:349/438, 196. **Not called** when a path is typed (not browsed), when the processing-mode radio
    changes (setup_tab.py:404-407 has no `command`), or on closing - such changes are lost on restart unless another
    save happens (READ, HIGH).
  - `apply_cached_values` (373-389): if the cache is non-empty, sets the 4 paths, processing mode, include-name flag,
    workflow mode, dark mode, and max length (`cached if cached > 0 else 35`; a non-number in the cache would raise
    `TypeError` here and stop the app from starting - LOW, only a hand-edited/corrupt cache); no validation of the
    mode strings (RAN: a test value 'darkscan' was restored as the workflow mode; an invalid processing mode would
    make FileOrganizer raise at run time, core/file_organizer.py:62-63). Applies dark mode if cached (388-389).
    Caller: 55.
  - `run` (391-398): status "Welcome to GST File Organizer v3.0"; log "🎉 Application started" (info), "💡 Ready to
    organize your GST files!" (success), "⚠️ IMPORTANT: Templates must have 'Links' sheet for reports to work!"
    (warning); enters the Tk main loop. Caller: main.py:112.
- **Business rules:** none of GST nature here (counts/percentages in the summary are display of scan statistics).
- **Edge cases handled:** theme not available (74-77); title-bar children that refuse a colour (171-174); checklist
  refresh failure after a theme change (190-194); unknown tab index (337); cache missing (375); invalid cached length
  (386-387).
- **Gaps:** closing during a run (exp4); Step 4 activity ignored on close; settings not saved on typing/radio/close;
  bare `except:` clauses (76, 173) also swallow keyboard interrupts; unused imports/variables (above); `logger` unused.
  Startup checks the existence of the four cached paths on the UI thread (setup_tab.py:156 then traces, and 191 if
  dark) - on an unreachable share each check can block the window for the network timeout (READ, MEDIUM).
- **Excel touchpoints:** none directly; `ExcelHandler()` created here (116) is shared by the worker for every report.
- **Files/settings:** reads and writes `~/.gst_organizer_cache.json` (121, 371 via CacheManager). Creates
  `%TEMP%\gst_excel_temp` (116, via ExcelHandler).
- **UI elements:** main window (title, size, centre, 'clam' theme, background), title strip with Dark Mode checkbox
  (A-R4's widget), the notebook with four tabs ("📁 Step 1: Setup", "✅ Step 2: Validation", "🚀 Step 3: Processing",
  "🔄 Step 4: Extract" - RAN), status bar ("💡 …" left, "v3.0 | Production Ready" right - A-R4's widget), the close
  confirmation dialog.
- **Comments/version strings:** docstring "Main GUI Application Window for GST File Organizer v3.0 / Modularized
  version - ties all components together" (2-4); "Import core modules (these imports remain the same)" (14) refers to
  a past refactor; comments 88-89, 144-145, 180-187, 199-205, 335-336, 385 agree with code. Version texts "v3.0" (393).
- **Pattern notes:** "app object as shared state bag": tabs and handlers write attributes onto the app
  (`app.client_tree`, `app.summary_text`, `app.log_text`, `app.progress_var`, `app.start_btn`, `app.scan_btn`,
  `app.client_name_check`, …) and other modules read them - there is no declared interface, so the build order matters
  (e.g. `update_summary` and `update_client_tree` guard with `hasattr`). Delegator methods everywhere.
- **History:** edac6a3; A-only: bfcb24a (Step 4 tab, tab-name fix), 510f7b4 (workflow_mode variable and cache key),
  c89c1d4 (checklist refresh after theme change), 725fbaf (removed tab-padding hack). (READ diff 47a1492..A/main.)

### gui/tabs/__init__.py (8 lines) - [UI]

- **Purpose:** re-exports `SetupTab`, `ValidationTab`, `ProcessingTab` (lines 4-6) with `__all__` (8). (READ)
- **Gap:** `ExtractTab` (added in bfcb24a) is **not** re-exported - the list was not updated (READ, HIGH). Harmless
  because nobody imports from `gui.tabs` directly except tests/test_optional_pywin32.py:54 (`from gui.tabs import
  extract_tab`, which imports the submodule, not the re-export). `gui/main_window.py:26-29` imports each tab module
  directly. Dead re-export.
- **History:** edac6a3 only.

### gui/tabs/extract_tab.py (131 lines) - [UI]

- **Purpose:** Step 4 tab: hosts the Power Query Extractor panel inside the main window, shows a short explanation
  strip, shows a fallback notice if the extractor cannot be imported, and keeps the panel's folder in step with
  Step 1's target folder until the user picks a different one in Step 4. (READ)
- **Imports:** standard library `logging` (13), `tkinter as tk` (14), `tkinter.ttk` (15); project
  `utils.constants.GUI_CONFIG` (17); **guarded** import (24-30): `from power_query_extractor.gui.extractor_window
  import ExtractorPanel` inside `try … except Exception` - on failure `ExtractorPanel = None`,
  `EXTRACTOR_IMPORT_ERROR = <the exception>`, and a logger warning "Power Query Extractor unavailable: <e>".
  Deliberately imported from `extractor_window`, not `extractor_main`, because the latter calls
  `logging.basicConfig()` at import (docstring 8-10; confirmed power_query_extractor/extractor_main.py:12).
- **Module-level globals:** `logger` (19), `ExtractorPanel` / `EXTRACTOR_IMPORT_ERROR` (25-29), `COLORS =
  GUI_CONFIG['colors']` (32).
- **Class `ExtractTab` (35-131).** Instantiated at gui/main_window.py:146.
  - `__init__` (38-46): stores app/notebook; `_pushed_folder = None` (43; the folder value this tab last pushed in);
    `panel = None` (44); `create_tab()`; `_attach_target_folder_sync()`.
  - `create_tab` (48-82): adds tab "🔄 Step 4: Extract" (49-50); white header strip with "🔄  Step 4 — Refresh Power
    Query & extract data" (Segoe UI 12 bold, blue) and "Opens the ITC and Sales reports produced in Step 3, refreshes
    their Power Query connections, and collects the results into one workbook. Press Scan to list the clients in the
    folder below." (Segoe UI 9, grey, wrap 980 px) (54-65). If the import failed -> notice and return (67-69). Else
    builds `ExtractorPanel(tab, target_folder=<Step 1 target or None>, show_header=False, show_status_bar=False,
    auto_load=False)` (74-80) - no scan at startup (RAN by the existing test: `scans_at_startup == []`), packs it
    (81), records the panel's folder text as `_pushed_folder` (82; empty at startup, because cached values are applied
    after the tabs are built). Caller: 45.
  - `_create_unavailable_notice` (84-105): pale-orange card "⚠️  Power Query refresh is unavailable" and text "This
    step needs pywin32 (Excel automation), which is not installed for the Python running this app.\n\nInstall it
    with:    pip install pywin32\n\nNote that 'py' and 'python' can be different interpreters - install it for the one
    you launch the app with.\n\nDetails: <error>", plus an italic line "Steps 1-3 are unaffected: organising files and
    creating the reports works without pywin32." Caller: 68.
    **Stale message (READ + RAN, HIGH):** since c463256 the pywin32 imports are guarded inside
    power_query_extractor/core/report_processor.py:20-35, so a missing pywin32 no longer makes this import fail (RAN on
    Linux without pywin32: the panel was built, the log said "pywin32 not available … Power Query refresh is
    disabled"). This notice can now appear only for *other* import failures, where its pywin32 explanation would be
    wrong; the missing-pywin32 case is reported per client when the user presses Start Processing in Step 4
    (report_processor.py:66-77).
  - `_target_folder_value` (108-112): Step 1 target or None; errors -> None. Callers: 76, 129.
  - `_attach_target_folder_sync` (114-119): adds a write-trace on `app.target_folder`; errors swallowed. Caller: 46.
  - `_on_target_folder_changed(*_args)` (121-131): if no panel -> return. If the panel's box is non-empty and differs
    from what this tab last pushed, the user chose their own folder -> leave it (127-128). Otherwise copy Step 1's
    target (or '') into the panel and remember it (129-131). Fires on every write, including each keystroke typed in
    Step 1's target entry and the cached value at startup. Caller: trace at 117.
    Gap: `_open_extraction_tab` (processing_handler.py:166-168) writes the panel folder directly and bypasses this
    bookkeeping, so the "never clobber a folder the user chose in this tab" promise (comment 125-126) is broken by the
    full-pipeline hand-over (READ, HIGH).
- **Business rules:** none of GST nature here (the extractor's cell mappings are A-R2's, e.g. the merge-point commit
  47a1492 added Table 6J / 8D sheet mappings - not in this file).
- **Edge cases handled:** extractor import failure (24-30, 67-69); no panel for the folder sync (122-123); Step 1
  variable errors (109-112, 116-119).
- **Gaps:** stale notice (above); hand-over bypass (above); the panel lives in a notebook inside the notebook
  (commit bfcb24a: "Flattening is the next decision") - Step 4 has its own "Setup" and "Processing" sub-tabs and its
  own log, separate from Step 3's log; Step 4's processing state is invisible to `on_closing` and to Step 3.
- **Excel touchpoints:** none here (the panel refreshes reports through COM - A-R2).
- **Files:** none directly. The panel, when embedded with `auto_load=False`, never reads the settings cache
  (extractor_window.py:132-139; `load_from_cache` only runs standalone).
- **UI elements:** tab, header strip, notice card (quoted above); the embedded panel is described in Cross-file A.
- **Comments:** docstring 2-11 agrees; comment 21-23 ("The extractor pulls in pywin32 for Excel COM … Step 4 shows an
  explanation instead") is stale for the pywin32 case (above); 71-73 agrees; 41-42 and 125-126 describe intent that the
  hand-over breaks.
- **Pattern notes:** the only file in my set that uses a guarded import with a user-visible fallback; errors around
  the folder sync are swallowed silently (111, 118-119).
- **History:** created in bfcb24a (26-07-2026, A only); c463256 added the guard and the notice.

### gui/tabs/processing_tab.py (117 lines) - [UI]

- **Purpose:** builds the Step 3 screen: progress area with Start/Stop and the colour-coded processing log. (READ)
- **Imports:** `tkinter as tk`, `tkinter.ttk`, `tkinter.scrolledtext` (4-5); `utils.constants.GUI_CONFIG` (6). No
  OS-specific modules.
- **Class `ProcessingTab` (9-117).** Instantiated at gui/main_window.py:143.
  - `__init__` (12-15): stores app/notebook; `create_tab()`.
  - `create_tab` (17-117): tab "🚀 Step 3: Processing" (19-20). Progress card (white, solid border 2) with a blue
    header "📊 PROCESSING PROGRESS" (Arial 12 bold, white text) (23-33); `app.progress_var` (DoubleVar) and a
    determinate `ttk.Progressbar` (39-43); `app.progress_label` "⏳ Ready to start..." (Arial 11 bold) (46-51);
    `app.current_operation` empty label (Arial 10, blue) (53-58); buttons (61-87): `app.start_btn` "🚀 START
    PROCESSING" (green, white text, flat, hand cursor) bound with `.bind('<Button-1>', lambda e:
    app.start_processing())` (74) and `app.stop_btn` "⏹ STOP" (red, initially `state='disabled'`) bound with
    `.bind('<Button-1>', lambda e: app.stop_processing())` (87). Log card with a green header "📝 PROCESSING LOG"
    (90-100); `app.log_text` = ScrolledText, 18 lines high, Consolas 9, background '#1E1E1E', white text, read-only,
    word wrap (103-110); colour tags 'success' '#4CAF50', 'warning' '#FF9800', 'error' '#F44336', 'info' '#2196F3',
    'normal' '#FFFFFF' (113-117). Caller: 15.
- **Gaps:** both buttons use `.bind` instead of `command=` - the disabled look is cosmetic (RAN exp3 for Stop; Start is
  protected only by the re-entrancy guard), keyboard users cannot press them; there is no "clear log" or "open output
  folder" control; the log is never cleared between runs.
- **Excel / files / rules:** none.
- **Pattern notes:** Arial fonts here vs Segoe UI on Step 1; attributes are published onto the app object
  (`app.progress_bar` is never read elsewhere).
- **History:** edac6a3 only (no later commit touches it).

### gui/tabs/setup_tab.py (615 lines) - [UI]

- **Purpose:** builds the Step 1 screen - a scrollable page with a header and readiness pill, a live checklist of
  the four inputs, the workflow choice, the four path rows with Browse buttons, processing mode and folder-name
  options, the Scan / Re-scan buttons with a readiness strip, and a reference card of expected file names - and keeps
  the checklist and buttons up to date as the user picks or types paths. (READ)
- **Imports:** `tkinter as tk` (13), `pathlib.Path` (14), `tkinter.ttk` (15); project `utils.constants.FILE_PATTERNS,
  GUI_CONFIG, PROCESSING_MODES, WORKFLOW_MODES` (17-19), `..utils.ui_helpers.UIHelpers` (20), `..utils.setup_validation
  as sv` (21; Tk-free check logic, A-R4's file), `..widgets.tooltip.Tooltip` (22). No OS-specific modules (but Windows
  assumptions: fonts 'Segoe UI'/'Consolas'; wheel deltas of 120, see `_on_mousewheel`).
- **Module-level constants (operational and presentation):**
  - `COLORS = GUI_CONFIG['colors']` (24).
  - `STATE_STYLE` (29-33): OK -> ('✓', '#107C10', '#E8F5E8'); MISSING -> ('○', '#8A8F98', '#F1F3F4'); INVALID -> ('✕',
    '#D83B01', '#FDECEA'). `STATE_STYLE_DARK` (34-38): ('✓', '#7EE787', '#1B3326'), ('○', '#9AA0A6', '#3A3A3A'),
    ('✕', '#FF8A7A', '#4A2320').
  - `INPUT_ACCENT` (42-47): source green '📂', itc blue '📊', sales teal '💰', target orange '🎯'; `INPUT_ACCENT_DARK`
    (48-53) brightened.
  - `LIGHT_SURFACE` / `DARK_SURFACE` (56-67): 'muted' text colour and the pill/strip colour pairs; keys 'card' and
    'border' are **never read** (READ grep).
  - `CARD_BORDER = '#DADCE0'` (69).
  - Hover-help texts: `CHECKLIST_HELP` (73-86), `BROWSE_HELP` (88-93), `WORKFLOW_HELP` (95-101), `PROCESSING_HELP`
    (103-110), `CLIENT_NAME_HELP` (112-117; contains a generic placeholder company name as an example), `LENGTH_HELP`
    (119-123), `SCAN_HELP` (125-128), `RESCAN_HELP` (130-133). Quoted and checked against behaviour below.
  - `REVALIDATE_DELAY_MS = 350` (135): debounce after a path changes ("avoid stat'ing a UNC path on every keystroke").
  - `SCROLL_STEP_PX = 22` (137), `SCROLL_UNITS_PER_NOTCH = 3` (138).
- **Class `SetupTab` (141-615).** Instantiated at gui/main_window.py:141; tests/test_setup_tab_smoke.py:60.
  - `__init__` (144-156): `_rows` (key -> row widgets), `_path_tooltips`, `_revalidate_job`, `_exists_cache` (path ->
    True/False), `_canvas`, `_wheel_accum`, `_accent_bars`; `create_tab()`, `_attach_traces()`,
    `refresh_status(force=True)` (156; at this moment the paths are still empty, so every row shows "Not selected yet").
  - `create_tab` (159-201): tab "📁 Step 1: Setup"; a Canvas + vertical scrollbar + inner frame so the page scrolls
    (163-172); builds the seven parts in order (177-183); sets `yscrollincrement` 22 px (191); one application-wide
    wheel binding `bind_all('<MouseWheel>', _on_mousewheel, add='+')` (201; long comment 193-200 explains why not
    Enter/Leave and not `unbind_all` - commits d64d3a9 and bf2f1dd). Caller: 154.
  - `_is_dark` (204-209): current dark-mode flag; errors -> False. Callers: 542, 593.
  - `_is_in_scroll_area(widget)` (211-218): walks up the widget's parents; True if the canvas is among them. Caller: 232.
  - `_on_mousewheel(event)` (220-248): ignores the event if the canvas is gone (221-228), if the pointer is not over this
    tab's canvas (232-233), or if everything already fits (236-238) - returning None so others can use the wheel;
    otherwise accumulates `-delta/120*3` and scrolls by whole units, returning 'break' (243-248). Caller: binding at 201;
    tests (tests/test_setup_tab_smoke.py:164-208). Windows assumption: uses `event.delta` (Windows/macOS); with Tk 8.6 on
    Linux the wheel arrives as Button-4/5 and this handler never fires (READ, LOW impact - app targets Windows).
  - `_accent_card(parent, accent, pady)` (250-260): a bordered card with a 5-px coloured stripe on the left; records the
    stripe for repainting (257); returns the white inner frame. Callers: 263, 283, 318, 454, 480.
  - `_create_header` (262-280): "📁  Step 1 — Setup" (Segoe UI 15 bold, blue); `summary_label` readiness pill on the
    right (Segoe UI 9 bold); text "Choose where your GST files are, the two report templates, and where the organised
    output should go." Caller: 177.
  - `_create_checklist` (282-314): card titled "CHECKLIST"; for source / itc / sales / target (labels 'Source folder',
    'ITC template', 'Sales template', 'Target folder') a tinted row with an icon label ('○' initially), a name label
    "<emoji>  <label>" (Segoe UI 10 bold, width 20) and a detail label; each of the four widgets gets the
    `CHECKLIST_HELP` tooltip. No closures over the loop variable. Caller: 178.
  - `_create_workflow` (316-345): card "WHAT SHOULD THIS RUN DO?"; one radio per `WORKFLOW_MODES` entry bound to
    `app.workflow_mode` with `command=self._on_workflow_change` - "📁 Organise files & create reports" ("Sort the files
    and build the ITC/Sales reports, then stop.") and "🔄 Full pipeline" ("Also refresh Power Query and extract the
    data (slower).") (utils/constants.py:163-172); tooltips from `WORKFLOW_HELP`; footnote "Step 4 stays available
    either way - this only decides whether Step 3 continues into it automatically." (Segoe UI 8 italic). Caller: 179.
  - `_on_workflow_change` (347-351): `app.save_cache()`, all errors swallowed silently. Caller: radio command 334.
  - `_create_inputs` (353-366): section "📂 FOLDERS & TEMPLATES" with description "Pick each item - the checklist
    above updates as you go" (via `UIHelpers.create_colored_section`, blue header) and four `_path_row`s: "Source folder
    (files to organise)", "ITC report template", "Sales report template", "Target folder (where output is created)".
    Caller: 180.
  - `_path_row(parent, key, label, variable, command)` (368-391): bold label; an Entry bound to the Tk variable (user
    can type a path); a tooltip showing the full value (381); "Browse…" ttk button with `command=` (384; comment 383
    explains why); a red "problem" label under the entry (388-391). Callers: 359-366.
  - `_create_options` (393-451): section "⚙️ PROCESSING MODE" / "How the output folders are organised" (teal header);
    one radio per `PROCESSING_MODES` - "🆕 Fresh Run" ("Create new folder structure (recommended for first time)"),
    "🔄 Re-Run" ("Add new version folders to existing structure"), "▶️ Resume" ("Continue previously interrupted
    processing") (utils/constants.py:174-190) - **no `command`, so a change is not saved** (404-407); tooltips from
    `PROCESSING_HELP`; checkbox `app.client_name_check` "Include client name in Level 4 folder names" bound to
    `app.include_client_name_in_folders`, `command=app.update_global_folder_setting` (417-422) and note "Overrides the
    per-client setting on the next tab." (425-427); "Max client folder name length:" + Spinbox 15-100 bound to
    `app.client_name_max_length` with `command=on_length_change` (440-443); hint label `length_hint` (448-451).
    Nested `on_length_change` (434-438): if the value is <= 0 sets 35, updates the hint, saves the cache - runs only on
    the spinbox arrow buttons, **not when the user types** (no trace, no validation): typed values outside 15-100, 0,
    or text are accepted (RAN exp5: '' and 'abc' make `get()` raise TclError; '20.5' becomes 20; '500' accepted).
    Caller: 181.
  - `_create_actions` (453-475): card with `app.scan_btn` "🔍  Scan Files" (`command=app.scan_files`) and
    `app.rescan_btn` "🔄  Re-scan" (`command=app.rescan_files`), with tooltips; `action_bar` strip + `action_hint`
    label (Segoe UI 9 bold). Caller: 182. RAN exp3: a disabled ttk Scan button does not run its command.
  - `_create_reference` (477-503): purple card "📝  EXPECTED FILE NAMES" / "Files in the source folder must follow
    these patterns to be recognised."; one row per `FILE_PATTERNS` entry showing its `type` and `description`
    (alternate rows tinted '#FAF5FC'; descriptions in Consolas 8): GSTR-2B Reco 'GSTR-2B-Reco-ClientName-State-
    Period.xlsx'; IMS Reco 'ImsReco-ClientName-State-DDMMYYYY.xlsx'; GSTR-3B Export 'GSTR3B-ClientName-State-
    Month.xlsx'; Sales 'Sales-ClientName-State-StartMonth-EndMonth.xlsx'; Sales Reco 'SalesReco-ClientName-State-
    Period.xlsx'; Annual Report 'AnnualReport-ClientName-State-Year.xlsx' (utils/constants.py:13-62). These are
    simplified descriptions: the real patterns also accept '.xls', any letter case, an optional "(N) " prefix on GSTR3B
    names and an optional " (N)" suffix on ImsReco names, and require ImsReco's last part to be 8 digits
    (constants.py:15, 23, 31, 39, 47, 55). Caller: 183.
  - `_attach_traces` (506-510): write-trace on the four path variables -> `_schedule_refresh`. Caller: 155.
  - `_schedule_refresh(*_args)` (512-519): cancels the pending re-check (errors ignored) and schedules
    `refresh_status` in 350 ms. Caller: traces.
  - `_exists(path_str)` (521-529): existence check remembered per exact text; OSError/ValueError -> False. Caller:
    passed into `sv.evaluate_setup` (540). The memory is cleared only on `refresh_status(force=True)` (startup and
    dark-mode toggles): if a folder is created later, or a share comes back, the row keeps the old answer until the
    text changes (READ, HIGH). It only checks existence - a *file* path in the source/target row shows OK (folder check
    in gui/utils/setup_validation.py:47-53 does not test "is a folder").
  - `refresh_status(force=False)` (531-583): clears the pending-job marker; on force clears the existence memory;
    evaluates the four checks (sv.evaluate_setup: source/target -> 'Not selected yet' / 'Folder not found' / the path
    shortened to "…<last 42 chars>"; templates -> 'Not selected yet' / 'File not found' / 'Not an Excel template' (suffix
    not in .xlsx .xltx .xltm .xlsm) / the file name); picks light or dark styles; paints each row's icon, colours and
    detail; shows the message under the entry only when INVALID (557-559); repaints the accent stripes (564-568); updates
    the entry tooltips with the full paths (570-576); sets the pill text "N of 4 items set" (amber) or "Ready - all 4
    items set" (green) (579-581); calls `_update_actions`. Callers: 156, the timer from 518, gui/main_window.py:191,
    tests. Runs on the UI thread; each new path text costs one existence check, which on an unreachable share can
    block the window (MEDIUM).
  - `_update_actions(checks)` (585-608): Scan and Re-scan are enabled only when the source row is OK (586-591).
    Strip text: not ready -> "Select a valid source folder to scan" (red); everything OK -> "Ready to scan, and to
    process afterwards" (green); otherwise "Ready to scan · still needed for Step 3: <labels of every non-OK row>"
    (amber) (594-606). `blocking[0]` at 596 is always present when not ready (READ). Caller: 583. Nothing else in the
    app disables these buttons (they stay usable during Step 3 - RAN exp8).
  - `_update_length_hint` (610-615): "(default 35 · names over <limit> are shortened)"; error -> 35. Callers: 437, 451.
- **Business rules in this file:**
  - The expected-file-name reference card (477-503) shows the six return-file patterns [GST RULE] (which return files
    the tool expects and how the period is written in the name: Period / DDMMYYYY / Month / StartMonth-EndMonth /
    Year). The period parts are not validated anywhere in my files (READ; core/file_parser.py:97-106 stores them as
    metadata only). TIME-VARYING (the list of returns, e.g. IMS).
  - UI gating rules: Scan needs a valid source folder; the checklist treats templates (existence + suffix) and target
    (existence) as needed for Step 3 (sv.can_process) - but Step 3 does not enforce them (see Cross-file D).
- **Help texts compared with behaviour (READ unless marked):**
  - CHECKLIST_HELP 'source' (74-76): "They are read from here and copied into the target folder - nothing here is
    moved, renamed or deleted." - agrees (copies use shutil.copy2, utils/helpers.py:199; RAN: source untouched). The
    *copies* are renamed (spaces -> underscores, "Private"->"Pvt", "Limited"->"Ltd"; RAN exp7).
  - 'itc'/'sales' (77-82): "A copy is made per client and its Links sheet is filled in, so the template itself is never
    modified." - agrees (template copied to a staging file first, core/excel_handler.py:147-154, 273-279).
  - 'target' (83-85): "Each run adds a new timestamped folder, so previous runs are left untouched." - true only for
    'fresh' runs started in different minutes; 'rerun'/'resume' write into the newest existing Level 1 folder, and two
    fresh runs in the same minute share one folder and overwrite its reports (RAN exp7; READ).
  - PROCESSING_HELP 'resume' (108-109): "Continue a run that was interrupted, skipping files that are already in
    place." - **not what happens** (RAN exp7, HIGH): resume reuses the newest Level 1 folder but always creates a new
    `Version-<DDMMYY HHMM>` folder (core/file_organizer.py:103-108), so files are never "already in place" unless the
    run is in the same minute; the log showed "Organized 2 files" again into Version-…0341 after a fresh run at …0339.
  - 'rerun' (106-107): "Add a new version inside the most recent run's folder" - agrees (RAN exp7). "Most recent" =
    newest folder modification time (core/file_organizer.py:184), which changes whenever something is added inside it.
  - CLIENT_NAME_HELP (112-117) - agrees about Level 4 folder names (RAN). Its "about 218 characters" Excel limit is a
    technical constant also used in core/excel_handler.py:356 and utils/constants.py:85-90.
  - LENGTH_HELP (119-123) "Longest client name allowed in a folder name before it is shortened." and the hint "names
    over N are shortened" - **only the Level 2 folder "<client>-<STATE>" is shortened, and the limit counts the whole
    "<client>-<STATE>" text, not the name** (utils/helpers.py:620-644 via core/file_organizer.py:83-87); the Level 4
    folders "ITC (<client>)", the copied file names and the report file names keep the full cleaned name (RAN exp7,
    limit 15: Level 2 "GenericLongC-KL", Level 4 "ITC (GenericLongClientNamePlaceholder_Pvt_Ltd)"). HIGH.
  - SCAN_HELP (125-128) "Read-only - nothing is copied or changed. Results appear on Step 2." - agrees (the scan only
    lists and reads 8 bytes per file); results appear on Step 2 only when at least one client was found.
- **Edge cases handled:** canvas destroyed (225-228); pointer outside the tab (232-233); nothing to scroll (236-238);
  fractional wheel deltas (243-247); existence-check exceptions (527-528); unknown check keys (549-550); missing
  tooltip/attributes guarded with getattr (564, 588-591, 607); invalid length on the arrows (435-436).
- **Gaps:** typed paths and the processing-mode choice are not saved (no save on trace / no radio command); spinbox
  typing unvalidated (RAN exp5); existence memory not refreshed (above); help texts that disagree with behaviour
  (above); the checklist's "still needed for Step 3" is advisory only (Cross-file D); Scan/Re-scan not disabled during
  Step 3. No closures-in-loops issues; no mutable defaults.
- **Excel touchpoints:** none. **Files:** reads existence of the four paths; saves the cache via the app (349, 438).
- **UI elements:** all quoted above. Layout: scrollable single column of cards, each with a 5-px coloured stripe;
  padding 20/16; Segoe UI fonts (8-15), Consolas 8 for patterns; colours as in the constants above.
- **Comments/docstrings:** module docstring 2-11 agrees (debounce + cache; Scan enabled by the source folder alone).
  Comment 446-447 ("Message reflects the ACTUAL configured limit") is only true after an arrow click or restart - typing
  does not update the hint. Comment 383 agrees. Comment 135 agrees. No TODO/FIXME, no commented-out code.
- **Pattern notes:** this file established `command=` buttons, tooltips, theme-aware colour pairs and the
  "bind_all with add='+' and an inside-the-area test" wheel pattern (the commit bf2f1dd added a test that enforces the
  wheel pattern across GUI files, tests/test_scroll_bindings.py). Steps 2 and 3 still use the older `.bind` buttons and
  Arial fonts (inconsistent).
- **History (A only after the split):** 322717f rewrote the tab as a live checklist (and added
  gui/utils/setup_validation.py and gui/widgets/tooltip.py); d64d3a9 fixed scrolling; 510f7b4 added the workflow card;
  169c7b5 dark-mode colour pairs and stripe repainting; bf2f1dd hover help and the final wheel-binding fix. Earlier
  (shared) version had a "Welcome" panel, `.bind` buttons and a hard-coded "Names >10 chars will show warning" note
  (READ: removed lines in `git log -p`). I read these commits' messages and the removed/added UI strings, not every
  line of their diffs (see INCOMPLETE).

### gui/tabs/validation_tab.py (247 lines) - [UI]

- **Purpose:** builds the Step 2 screen: instructions, scan summary, the Dry Run / Export / Start buttons, and the
  client table with selection buttons and keyboard/mouse bindings. (READ)
- **Imports:** `tkinter as tk`, `ttk`, `scrolledtext` (4-5); `utils.constants.GUI_CONFIG` (6);
  `..widgets.collapsible_frame.CollapsibleFrame` (7). No OS-specific modules.
- **Class `ValidationTab` (10-247).** Instantiated at gui/main_window.py:142.
  - `__init__` (13-16): stores app/notebook; `create_tab()`.
  - `create_tab` (18-42): tab "✅ Step 2: Validation"; two columns - left fixed 450 px wide (`pack_propagate(False)`,
    comment says 40%), right expanding (comment 60%); left: instructions, summary, actions; right: client table.
  - `create_collapsible_instructions` (44-66): collapsible box "📋 VALIDATION INSTRUCTIONS" (orange tones '#FFF3E0' /
    '#E65100', Arial 10) with bullets "• Review the scan summary below", "• Check client list - ✅ = complete, ⚠️ =
    missing files", "• Select clients using click or keyboard (↑↓ + SPACE)", "• Try 'Dry Run' to preview without
    processing", "• Click 'Start Processing' when ready" (the button on this tab is labelled "🚀 START"; "⚠️" also
    marks duplicates-only clients, not only missing files). Clicking a row's "📁 Name" cell (folder-name toggle) is not
    mentioned anywhere on screen.
  - `create_summary_section` (68-92): blue header "📊 SCAN SUMMARY" (Arial 12 bold) and `app.summary_text`, a read-only
    ScrolledText (Consolas 9, '#f8f9fa', word wrap) inside a frame fixed at 200 px high.
  - `create_validation_actions` (94-142): "🎯 READY TO PROCESS?" (Arial 14 bold, green) and three flat buttons: "🧪 DRY
    RUN" (orange) -> `app.dry_run()`, "📋 EXPORT LIST" (teal) -> `app.export_client_list()`, "🚀 START" (green) ->
    `app.start_processing()`, each wired with `.bind('<Button-1>', lambda e: …)` (118, 130, 142).
  - `create_client_selection_section` (144-247): green header "👥 CLIENT SELECTION"; hint "⌨️ Use ↑↓ arrows + SPACE to
    select, or click with mouse" (light blue box); buttons "☑️ Select All" (green), "☐ Clear All" (dark), "✅ Complete
    Only" (blue), bound with `lambda e, cmd=command: cmd()` (190 - default-argument capture, so no late-binding bug);
    `app.client_tree` Treeview with columns ('Client', 'State', 'Status', 'Files', 'Missing', 'Extra', 'FolderName'),
    `show='tree headings'`, 20 rows high (197-201); widths #0 50, Client 250, State 100, Status 150, Files 60, Missing
    180, Extra 80, FolderName 100 (min 80) (204-211); headings '✓', '🏢 Client Name', '🌏 State', '📊 Status', '📁
    Files', '❌ Missing', '➕ Extra', '📁 Name' (214-221); vertical and horizontal scrollbars (224-238); bindings
    `<Button-1>` -> `app.on_client_click`, `<space>` -> `app.on_space_key`, `<Up>`/`<Down>` (241-244); initial focus on
    the table (247).
- **Business rules:** the ✅/⚠️ meaning (54) reflects core/file_parser.py:277-286 (see client_handler section) -
  [GST RULE] adjacent (which returns make a client "complete"); TIME-VARYING list of expected returns.
- **Edge cases:** none explicit (construction only).
- **Gaps:** all six buttons use `.bind` (no keyboard activation, act on mouse-press, cannot be disabled; Dry Run and
  Start remain clickable during a run - Start is protected by the guard, Dry Run is not); the "📁 Name" toggle is
  undiscoverable; layout fixed widths (450 px left, 200 px summary) on the 1000-px minimum window.
- **Excel / files:** none.
- **Pattern notes:** older visual style (Arial, solid borders) vs Step 1; widgets published on the app object.
- **History:** edac6a3 only (no later commit touches it).

---

## Cross-file observations

### A. The user's workflow, screen by screen (READ + RAN exp1)

**Start-up** (main.py:57-112 -> gui/main_window.py:46-55, 391-398). The window "GST File Organizer v3.0" opens
1200×800, centred, minimum 1000×700, theme 'clam'. Top: a blue title strip "🏢 GST FILE ORGANIZER & REPORT GENERATOR"
/ "Organize files and generate Excel reports automatically" with a "🌙 Dark Mode" checkbox (gui/widgets/title_bar.py).
Middle: four tabs. Bottom: dark status bar. Settings from `~/.gst_organizer_cache.json` are restored (paths, processing
mode, workflow mode, include-name flag, dark mode, max length). Step 3's log already contains "🎉 Application started",
"💡 Ready to organize your GST files!", "⚠️ IMPORTANT: Templates must have 'Links' sheet for reports to work!". Step 4
is built but does not scan.

**Step 1 - Setup** (gui/tabs/setup_tab.py). Top to bottom: header with readiness pill ("N of 4 items set" /
"Ready - all 4 items set"); CHECKLIST (four tinted rows ✓/○/✕ with detail text); "WHAT SHOULD THIS RUN DO?" (two
radios: organise-only or full pipeline; the choice is saved at once); "📂 FOLDERS & TEMPLATES" (four entries with
"Browse…"; typing is allowed; the checklist re-checks 350 ms after typing stops; red text under an entry when the path
is wrong); "⚙️ PROCESSING MODE" (Fresh / Re-Run / Resume; "Include client name in Level 4 folder names"; "Max client
folder name length" 15-100, default 35); action card with "🔍  Scan Files" and "🔄  Re-scan" (enabled only when the
source folder exists) and a coloured strip saying what is still missing; "📝  EXPECTED FILE NAMES". Hover help on
almost every element. The user clicks **Scan Files**: the window freezes while the source folder (top level only) is
read; on success the app jumps to Step 2.

**Step 2 - Validation** (gui/tabs/validation_tab.py, gui/handlers/client_handler.py). Left: collapsible instructions,
the scan summary text, "🧪 DRY RUN", "📋 EXPORT LIST", "🚀 START". Right: "☑️ Select All", "☐ Clear All", "✅ Complete
Only" and the client table (✓ / Client Name / State / Status / Files / Missing / Extra / Name). All clients start
unticked. The user ticks clients (click the ✓ column, or arrows + Space), optionally toggles a client's folder-name
column, optionally runs a Dry Run (a text preview in the Step 3 log) or exports the list to a .txt file, then clicks
**START**.

**Step 3 - Processing** (gui/tabs/processing_tab.py, gui/handlers/processing_handler.py). Checks run (templates set and
found, target set; at least one client; optional long-name confirmation). The app switches to Step 3, disables Start,
enables Stop and starts one background thread. For each client in table order: create folders -> copy files -> ITC
report -> Sales report -> organisation text file; progress bar and "Client i of n: <name> | ETA: …" update; the dark log
fills with coloured lines. Then the summary workbook at Level 1, the "Processing Complete" box, the output folder opens
in Explorer, and (only in 'full' workflow) the app switches to Step 4, fills its folder with the target folder and
scans it. **STOP** takes effect when the current client is finished.

**Step 4 - Extract** (gui/tabs/extract_tab.py hosting power_query_extractor/gui/extractor_window.py:98-1030, A-R2's
file). A header strip, then the panel's own two sub-tabs. *Setup* sub-tab: "Target Folder" box (pre-filled from
Step 1's target; "📁 Browse", "🔍 Scan"); "Processing Options": **"Power Query Wait Time:"** spinbox 5-60 seconds,
default 10, "seconds (Min: 5, Max: 60)" (extractor_window.py:259-287); **"Refreshed File Suffix:"** entry, default
"_Refreshed_{timestamp}", hint "Use {timestamp} for datetime" (289-312); checkbox "Skip Refresh\n(Use existing refreshed
files)" (314-325); "🚀 Start Processing" (disabled until a scan lists clients, 328-340, 755-757). Right column "Select
Clients & Reports": "✓ Select All", "✗ Deselect All", "ITC Only", "Sales Only"; one row per client folder of the
newest "Annual Statement-*" folder (by modification time) with an **ITC** and a **Sales** checkbox - each ticked and
enabled only if an `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` exists in that client's newest Version folder - and a
"Last Refresh Status" text (670-753). The wait time and suffix are handed to the refresher when Start Processing is
pressed (863-865); the work runs in the panel's own thread, logs into the panel's *Processing* sub-tab, and ends with
a consolidated workbook in the Level 1 folder and an offer to open it (885-975). In 'organize' mode the user opens
Step 4 and presses Scan; in 'full' mode Step 3 does that, but the refresh itself always needs the user's press
(RAN exp1 `full`). If the extractor module could not be imported, Step 4 shows the warning card instead.

### B. Shared state: write path vs read path

| Shared item | Written | Read | Consistent? |
|---|---|---|---|
| Settings cache `~/.gst_organizer_cache.json` (JSON) | gui/main_window.py:358-371 -> gui/utils/cache_manager.py:29-48 (adds `recent_folders`, keeps last 5 source folders) | gui/main_window.py:120-121, 373-389 (all nine keys); power_query_extractor/gui/extractor_window.py:118, 587-599 (only `target_folder`, and only when standalone); power_query_extractor/extractor_main.py:41-42 (A-R2) | Keys and path agree (READ). `recent_folders` is written but never read by anything else. Typed paths / processing-mode changes are written only incidentally. |
| Tk path variables (`source_folder`, `itc_template`, `sales_template`, `target_folder`) | Entries (setup_tab.py:378), Browse (file_handler.py:23, 34, 45, 53), cache restore (main_window.py:376-379) | Checklist (setup_tab.py:538-539), scan (file_handler.py:72), checks (file_handler.py:122-149), worker per client (processing_handler.py:190, 220, 301, 310, 332, 501), Step 4 sync (extract_tab.py:110), hand-over (processing_handler.py:166) | The worker re-reads templates per client - mid-run edits take effect (RAN exp8). |
| `app.client_data` / `app.scanned_files` / `app.variations` | file_handler.py:71-72 (replaced), 102-105 (emptied in place) | client_handler.py:30, 138, 248; processing_handler.py:56, 101, 232; main_window.py:320-321; summary data processing_handler.py:224 (reference to the list) | Emptying in place during a run corrupts the run's summary (RAN exp8) or raises KeyError (READ). |
| Client keys "<client>-<STATECODE>" | core/file_parser.py:229-230 | Rebuilt from table text: client_handler.py:169-170, 181, 218; core/file_organizer.py:122 | Mismatch for digit-only text (RAN exp2). |
| `app.client_folder_settings` | client_handler.py:49, 185, 198, 219 | core/file_organizer.py:130 (via processing_handler.py:193, live reference) | Keys agree since 33d0baf except the digit-only case. |
| `app.is_processing`, `app.stop_requested` | processing_handler.py:123-124, 179, 528; main_window.py:352 | processing_handler.py:75, 229; main_window.py:341, 346 | Plain booleans across threads (acceptable under Python's lock); Step 4 has its own separate `is_processing`. |
| Output tree (`Annual Statement-*` / `<client>-<STATE>` / `Version-*` / `ITC_Report_*`, `Sales_Report_*`) | core/file_organizer.py + core/excel_handler.py via processing_handler.py:253, 297-298, 355-356 | extractor_window.py:623-659 (`Annual Statement-*`, newest by mtime; client folders not starting with '_'; `Version-*` newest by mtime; prefixes `itc_report_`/`sales_report_`) | Names agree (RAN exp1 `full`: Step 4 found both clients). "Newest by mtime" on both sides. |
| Step 4 folder box | extract_tab.py:130 (sync), processing_handler.py:168 (hand-over), the user | extractor_window.py:613-623 | Hand-over bypasses the sync bookkeeping (READ). |
| Excel staging folder `%TEMP%\gst_excel_temp` | core/excel_handler.py:53-54, 147-154 | same | Shared by every copy of the app on the PC. |

### C. Threading model (READ; RAN on Linux)

- UI thread: everything in Steps 1-2, Scan (synchronous), Dry Run, Export, `_open_extraction_tab`.
- One daemon worker per Step 3 run (processing_handler.py:129-134), never joined; it calls Tk directly (variables,
  dialogs, buttons - list in the processing_handler section) and schedules UI updates through `after(0)`.
- COM in the worker: core/excel_handler.py:120 calls `pythoncom.CoInitialize()` per report on that thread (A-R1).
- Step 4 has its own daemon worker (extractor_window.py:878-883) that marshals all UI updates through `_ui()`
  (989-1010) - a cleaner pattern than Step 3's, inside the same app (inconsistent patterns).
- No locks anywhere; only booleans and dicts are shared.

### D. Inconsistent patterns and rules across files

- **What gates Step 3:** Setup checklist (gui/utils/setup_validation.py:94-97) requires templates that exist *and*
  have suffix .xlsx/.xltx/.xltm/.xlsm and a target that *exists*; `validate_processing_inputs`
  (file_handler.py:132-151) only checks the templates exist and the target is non-empty, and the target is then
  created if missing (core/file_organizer.py:66-69). Three template-type lists: Browse filter `*.xltx *.xlsx`
  (file_handler.py:31, 42), checklist `TEMPLATE_SUFFIXES` (setup_validation.py:21), unused
  `TEMPLATE_EXTENSIONS = {'.xlsx', '.xltx', '.xltm'}` (utils/constants.py:197; READ grep: no reader).
- **Long-name limits:** 10 hard-coded (client_handler.py:191); configurable with fallback 10
  (processing_handler.py:92-97); configurable with fallback 35 (main_window.py:93, 387; setup_tab.py:436, 614;
  core/file_organizer.py:42).
- **Buttons:** Step 1 uses `command=` (setup_tab.py:384, 459-465); Steps 2-3 use `.bind('<Button-1>')`
  (validation_tab.py:118, 130, 142, 190; processing_tab.py:74, 87). Commit 322717f's message names this hazard.
- **Log routes:** `_log_async` vs direct `log_message` (processing_handler.py), plus Step 4's own log.
- **Errors to the user:** message boxes (file_handler, export), on-screen log lines (processing), silent swallowing
  (setup_tab.py:350-351, extract_tab.py:111-112/118-119, processing_handler.py:150-151), or uncaught into Tk (most of
  client_handler.py) - the last two are invisible without a console.
- **Fonts/visuals:** Segoe UI cards on Steps 1 and 4 vs Arial solid-border boxes on Steps 2 and 3.
- **Worker-thread UI:** Step 3 touches Tk from the worker; Step 4 marshals everything (extractor_window.py:989-1010).

### E. Dead code found in my files (READ grep over all .py, tests included)

- Delegators with no callers: gui/main_window.py:244-245 `toggle_client_selection`, 256-257
  `toggle_folder_name_setting`, 275-276 `process_files_thread`.
- Unused attributes/variables: `self.widgets` (main_window.py:107); `app.processing_thread` is written (main_window.py:104,
  processing_handler.py:129) but never read; `app.progress_bar` (processing_tab.py:40) never read; `logger` in
  main_window.py:38 and file_handler.py:10 never used; `surface['card']`, `surface['border']` (setup_tab.py:56, 62)
  never read.
- Unused imports: main_window.py `os`, `sys`, `Path`, `MESSAGES`, `get_timestamp`, `CollapsibleFrame`;
  client_handler.py `Path`.
- Unused re-exports: gui/__init__.py, gui/tabs/__init__.py (also missing ExtractTab), gui/handlers/__init__.py;
  `__version__` never read.
- Collected but never output: `itc_report_created` / `sales_report_created` (processing_handler.py:427-428) are not
  written by the summary workbook (core/excel_handler.py:464-491).

### F. UI texts that disagree with behaviour (summary)

1. 'Resume' does not skip already-copied files (setup_tab.py:108-109, constants.py:187) - RAN exp7.
2. Max length shortens only the Level 2 folder (setup_tab.py:119-123, 615) - RAN exp7.
3. "previous runs are left untouched" (setup_tab.py:84-85) - not for rerun/resume or same-minute runs.
4. "press Process to refresh" (processing_handler.py:171) vs button "🚀 Start Processing".
5. Step 4 notice blames pywin32 although missing pywin32 no longer triggers it (extract_tab.py:92-97) - RAN.
6. "Successfully processed p/t clients" after Stop/failures (processing_handler.py:500) - RAN.
7. "📂 Opened output folder" when `os.system` failed (processing_handler.py:514) - RAN (Linux/macOS only).
8. Table shows "☐ No" while global include-name is ON after a scan (client_handler.py:43, 49) - RAN.
9. "Click 'Start Processing' when ready" vs "🚀 START" button (validation_tab.py:57).
10. "Show count while scanning" comment (file_handler.py:74) - shown after the scan.

### G. Facts relevant to the author's Standards (technical, no GST judgement)

- Standard 13 (no console; worker errors to the window; settings/logs in `%LOCALAPPDATA%\[App Name]\`; find files
  from the app folder, not the current folder): Step 3's worker errors that escape its `try` (exp4) and every
  exception left to Tk in client_handler.py go only to stderr, which is invisible under `pythonw.exe`; settings go to
  `~/.gst_organizer_cache.json` (gui/utils/cache_manager.py:15); the log file is `gst_organizer.log` in the current
  folder (main.py:24 - A-R4).
- Standard 12 "Excel through COM": the Step 3 report path started from my handler uses `GetActiveObject` (attaches to
  the user's open Excel) and quits it in `finally` (core/excel_handler.py:124-128, 238-246); CoInitialize is called on
  the worker thread (excel_handler.py:120). Details for A-R1 / the Excel cross-reader.

---

## Possible client names

None found in my 11 files, in their full git history (`git log -p` on these paths, searched for company-style words
and GSTIN-shaped text), or in the commit messages of the commits touching them. Locations of generic placeholders,
listed only for completeness (all look like obvious "ABC"-style examples, very likely not real clients):

- gui/tabs/setup_tab.py:114 - example company name inside the hover help `CLIENT_NAME_HELP` (added in bf2f1dd).
- Commit 33d0baf message - example company name used to illustrate the "Private Limited" abbreviation.
- tests/test_main_window_tabs.py:88-91 - folder/file names built from a generic "ABC" placeholder (test reader's file).

---

## Questions for the author

1. 'Resume': should it continue inside the interrupted run's existing Version folder and skip files already copied?
   Today it makes a new Version folder each run, so nothing is skipped (RAN).
2. 'Max client folder name length': should it also shorten the client name in the Level 4 folders ("ITC (<client>)"),
   the copied file names and the report file names? Today only the Level 2 "<client>-<STATE>" folder is shortened, and
   the limit counts "-<STATE>" too (RAN).
3. Do your file names ever use a digit-only client name (for example with leading zeros) or a numeric state field
   (such as the two-digit state code)? If so, Step 2 cannot process those clients today (RAN).
4. When Stop is pressed: is "finish the current client, then stop" right? Should the end message say it was stopped,
   and should the full pipeline still move to Step 4 after a Stop?
5. Should Step 1 (Scan, Re-scan, Browse, typed paths) and Step 2 be locked while Step 3 runs? Changing a template path
   mid-run currently changes the template for later clients (RAN).
6. Should Step 3 and Step 4 be allowed to run at the same time (both drive Excel)? Should closing the window warn
   during a Step 4 refresh?
7. A client is counted as "Success" even when both of its reports failed. Should the count, the "Processing
   Complete" box and the summary workbook show report failures per client?
8. Should failed or skipped file copies be listed in the Step 3 log (today only successful copies are listed)?
9. The 14 "[DEBUG]" lines shown in the Sales part of the log (left on purpose in 064c634): keep or remove?
10. The per-client folder-name toggle warns above 10 characters, the Start check uses your configured limit (default
    35, fallback 10). Which limit is intended?
11. After turning "Include client name" ON and then OFF, every client stays ticked, so names remain in folder names.
    Is that intended?
12. Should typed paths and the Fresh/Re-Run/Resume choice be remembered like browsed ones?
13. "Complete Only" selects only clients with all six file types and no duplicates (multiple GSTR-3B allowed). Is that
    the set of returns you expect for every client? [GST RULE - for you to decide]
14. Where should settings be stored - the current home-folder dotfile or `%LOCALAPPDATA%\[App Name]\` as in Standard 13?

---

## INCOMPLETE

- Windows-only behaviour was not executed (no Windows, Excel or pywin32 here): the Excel COM report path started by
  Step 3, `os.startfile` (output folder, export), message boxes raised from the worker thread under Windows Tk, the
  hidden-Excel consequences of closing mid-run, the Windows mouse-wheel deltas and the Segoe UI rendering. These were
  traced by reading only and are marked MEDIUM where they matter.
- Git history: for my files I read the commit list, the messages of every commit touching them, the complete diffs
  47a1492..A/main for gui/main_window.py, gui/handlers/client_handler.py and gui/handlers/processing_handler.py, and the
  diffs of 6237597 and 1c48722. I did **not** read line by line the full diffs of gui/tabs/setup_tab.py (322717f,
  d64d3a9, 510f7b4, 169c7b5, bf2f1dd) and gui/tabs/extract_tab.py (bfcb24a, c463256), nor the edac6a3 originals of the
  tab files - only their messages, stats and the removed/added UI strings. The history reader should cover them.
- The embedded Step 4 panel (power_query_extractor/gui/extractor_window.py) is described only as far as the user sees
  it from Step 4; its full section is A-R2's.
