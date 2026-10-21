# A-R4 working notes - project A (AnnualStatement_oldv1, snapshot bf2f1dd, 27-07-2026)

Reader: A-R4. Scope: gui/utils/*, gui/widgets/*, main.py, utils/dependency_installer.py,
utils/requirements_check.py, requirements.txt, pytest.ini, .gitignore, .claude/settings.local.json.
All 17 assigned files were read in full with the Read tool (every line). Other files were read only
to find callers or to understand a call.

## Legend and method

- READ = learnt by reading the code. RAN = I executed something (in /root/merge-scratch/run/A-R4/).
- HIGH / MEDIUM / LOW = my confidence in an inference (not needed for plain READ facts).
- [GST RULE] = decides a tax figure, return/table/sheet/cell, rate, limit, rounding, period, GSTIN/PAN/state logic,
  ITC eligibility or a due date. TIME-VARYING = value that can change over time.
- Paths are repository-relative (for example `main.py:17`). Commits are short hashes on A/main.
- Line counts are real line counts (a file without a final newline still counts its last line).
- Scratch work: a copy of A is at /root/merge-scratch/run/A-R4/A (with __pycache__ removed and a throw-away
  `git init` used only for `.gitignore` checks); harness scripts in /root/merge-scratch/run/A-R4/work/
  (`harness_main.py`, `dm_test.py`, `widgets_test.py`). HOME was redirected to scratch folders for runs that
  could write the settings file, so nothing was written to /root/.gst_organizer_cache.json.
- Environment used for RAN items: venv-A = Python 3.14.6 on Linux, Tcl/Tk 9.0.3 (RAN: `tkinter.TclVersion` = 9.0),
  openpyxl 3.1.5, lxml 6.1.3, pytest 9.1.1, NO pywin32. GUI runs used `xvfb-run` (a virtual screen).
- Note for later agents: /root/merge-scratch/A-copy contains `__pycache__` folders dated 21-10-2026. They are NOT
  part of the snapshot (`git ls-tree -r A/main` lists none; commit 01862a0 removed them from git). Ignore them. (READ/RAN)

Plain-English terms used below (first use):
- widget = any on-screen element (button, label, text box, frame).
- Tk / Tkinter = Python's built-in toolkit for windows and buttons; ttk = Tk's "themed" widgets whose colours come
  from a shared style table rather than from each widget.
- pack / place = two Tk ways of positioning widgets (pack stacks them; place puts them at fixed x/y coordinates).
- callback = a function the toolkit calls later, e.g. when a button is clicked.
- thread / worker thread = a second line of execution running at the same time as the window; the "main thread"
  runs the window.
- JSON = a plain-text data format of keys and values.
- pip = Python's package installer; PyPI = the public internet site pip downloads packages from.
- interpreter = the python.exe program that runs the code; `sys.executable` = its full path.
- COM = Windows' mechanism for one program to drive another (here: Python driving Excel); pywin32 = the Python
  package that provides COM (`pythoncom`, `win32com`).
- stdout = the console output stream; pythonw.exe = the Windows Python that runs with no console window.
- subprocess / Popen = starting another program from Python.

---

### .claude/settings.local.json (9 lines) - [TOOLING]

- Purpose: Claude Code (an AI coding assistant) permission file for this folder; it pre-approves one kind of shell
  command so the AI tool may run it without asking. (READ)
- Content (READ, lines 1-9):
  - `permissions.allow` = `["Bash(python -m py_compile:*)"]` (line 4): the AI tool may run any shell command that
    starts with `python -m py_compile` (Python's syntax check: it compiles a file to catch syntax errors and writes
    a .pyc cache file) without prompting.
  - `permissions.deny` = `[]` (line 6), `permissions.ask` = `[]` (line 7).
- Secrets: NONE. No tokens, passwords, keys, user names or paths. (READ, HIGH)
- History: added in bd176bf (30-10-2025, before the split, so shared with B), never changed on A. B deleted it in
  c258047 (07-02-2026). (RAN: `git log`)
- Gaps: `.gitignore` does not ignore `.claude/`, so this per-user "local" file is committed. By Claude Code
  convention a `settings.local.json` is personal and normally not committed (MEDIUM). Harmless content.
- Not used by the app at run time (no .py file reads it; grep). (READ/RAN)

### .gitignore (80 lines) - [CONFIG]

- Purpose: tells git which files never to commit. (READ)
- Content, grouped as in the file (READ):
  - Python build/cache (lines 1-23): `__pycache__/`, `.pytest_cache/` (line 3, added 6c1e1c6), `*.py[cod]`,
    `*$py.class`, `*.so`, `.Python`, `build/`, `develop-eggs/`, `dist/`, `downloads/`, `eggs/`, `.eggs/`, `lib/`,
    `lib64/`, `parts/`, `sdist/`, `var/`, `wheels/`, `*.egg-info/`, `.installed.cfg`, `*.egg`, `MANIFEST`.
  - Virtual environments (25-29): `venv/`, `env/`, `ENV/`, `.venv`.
  - IDE (31-37): `.vscode/`, `.idea/`, `*.swp`, `*.swo`, `*~`, `.DS_Store`.
  - Project-specific (39-47): `*.log`, `gst_organizer.log`, `pq_extractor.log`, `.gst_organizer_cache.json`,
    `*.cache`, `.cache/`, `temp/`, `logs/`.
  - Excel temporary files (49-54): `~$*.xlsx`, `~$*.xls`, `*.tmp`, `temp_gst_*`, `backup_*`.
  - Output directories "(user data)" (56-59): `/Annual Statement-*/`, `/output/`, `/target/`.
  - "User data and sensitive information" (61-64): `/mnt/user-data/`, `*.bak`, `config/last_session.json`.
  - Test files (66-69): `test_files/`, `test_data/`, `*.test.xlsx`.
  - Windows (71-74): `Thumbs.db`, `ehthumbs.db`, `Desktop.ini`.
  - Power Query cache (76-80): `*.odc`, `*.iqy`, `*.dqy`, `*.xlquery`.
- What each project-specific entry corresponds to in code (READ):
  - `gst_organizer.log` = main.py:24 log file (written in the current folder). `pq_extractor.log` =
    power_query_extractor/extractor_main.py:16. `.gst_organizer_cache.json` = gui/utils/cache_manager.py:15 - but that
    file is written in the user's HOME folder, not the repo, so this entry only matters if HOME is the repo (LOW).
  - `logs/`, `temp/` = folders main.py:35-37 creates in the app folder (unused - see main.py).
  - `temp_gst_*`, `backup_*` = utils/constants.py:237-238 prefixes (`'temp_prefix': 'temp_gst_'`, `'backup_prefix': 'backup_'`).
    Note utils/helpers.py:225 actually names backups `<stem>_backup_<timestamp><suffix>` (does NOT start with
    `backup_`), so `backup_*` does not match those backups. (READ, HIGH)
  - `/Annual Statement-*/` = Level-1 output folder pattern `'Annual Statement-{timestamp}'` (utils/constants.py:82),
    only when the output is created at the repo root.
  - `config/last_session.json`, `.cache/` = candidate cache files the standalone extractor looks for
    (power_query_extractor/extractor_main.py:42-44). NOTHING in A writes them (grep). Leftovers of an earlier design (MEDIUM).
  - `/mnt/user-data/` = the upload folder of a web-based AI chat environment; suggests files were once handled in
    such a chat (LOW-MEDIUM inference).
- RAN (`git check-ignore --no-index` in my scratch copy): none of the 70 tracked files is ignored; the folders
  `logs/` and `temp/` created by running main.py are ignored. Matches confirmed for gst_organizer.log (line 41),
  pq_extractor.log (42), .gst_organizer_cache.json (43), `~$x.xlsx` (50), `Annual Statement-<x>/` (57),
  config/last_session.json (64), .cache/app_state.json (45).
- Gaps (RAN): Excel lock files for macro/template workbooks (`~$*.xlsm`, `~$*.xltx`, `~$*.xltm`) are NOT ignored,
  although the app accepts .xlsm/.xltx/.xltm templates (gui/utils/setup_validation.py:21). `gst_organizer_cache.json`
  (no leading dot - extractor_main.py:42 candidate) is not ignored. `.claude/` not ignored (see above).
- History: created 1cb878b (29-10-2025, 79 lines); 6c1e1c6 (23-07-2026) added `.pytest_cache/`. B has its own copy (shared until 47a1492).
- No GST rules.

### gui/utils/__init__.py (9 lines) - [UI]

- Purpose: package file that re-exports four GUI helper classes. (READ)
- Imports (READ): project only - line 4 `from .dark_mode_manager import DarkModeManager`, line 5
  `from .cache_manager import CacheManager`, line 6 `from .ui_helpers import UIHelpers`, line 7
  `from .status_bar import StatusBar`. No OS-specific imports.
- Module-level: line 2 docstring "GUI Utilities Package"; line 9 `__all__ = ['DarkModeManager', 'CacheManager', 'UIHelpers', 'StatusBar']`.
- Callers of the re-exports: NONE. Every user imports the submodule directly (gui/main_window.py:33-36,
  gui/tabs/setup_tab.py:20-21, tests). `from gui.utils import X` / `from .utils import X` for these four names
  appears nowhere (grep). (READ/RAN, HIGH)
- `setup_validation` is NOT listed in `__all__` although it lives here (added 322717f) - re-export list not maintained. (READ)
- Side effect that matters (RAN): importing ANY submodule, e.g. `gui.utils.setup_validation`, first runs
  `gui/__init__.py` (which imports gui/main_window.py and therefore core/*, openpyxl, power_query_extractor, tkinter)
  and then this file (tkinter again). RAN: `import gui.utils.setup_validation` loaded 222 modules including
  tkinter, openpyxl, gui.main_window and power_query_extractor, and printed the pywin32 warning from
  power_query_extractor/core/report_processor.py:35. So "free of Tkinter" in setup_validation.py:3 holds for the
  file itself but not for importing it through its package. (HIGH)
- History: unchanged since edac6a3 (31-08-2025). Gaps: dead re-exports (maintenance only). No GST rules, no UI, no files.

### gui/utils/cache_manager.py (56 lines) - [UI] (settings persistence for the GUI)

- Purpose: reads and writes the app's remembered settings ("cache") as one JSON file in the user's home folder, and
  keeps a list of the last 5 source folders. (READ)
- Imports (READ): standard library `json` (4), `logging` (5), `pathlib.Path` (6). No third-party, no OS-specific.
- Module-level: `logger = logging.getLogger(__name__)` (8).
- class CacheManager (11-56):
  - `__init__(self, cache_file=None)` (14-16): `self.cache_file = cache_file or (Path.home() / '.gst_organizer_cache.json')`;
    `self.cache_data = {}`. Parameter `cache_file` is never passed by the app (only main_window.py:120 `CacheManager()`),
    so the path is always the default. The code assumes a Path object: a plain string would fail at `.exists()`
    (RAN: logged "Error loading cache: 'str' object has no attribute 'exists'").
    - Exact path: `Path.home()` = the user's home folder. On Windows that is `%USERPROFILE%`, so the file is
      `C:\Users\<user>\.gst_organizer_cache.json` (READ; Windows mapping HIGH by Python's documented behaviour).
  - `load_cache(self)` (18-27): if the file exists, opens it with `open(..., 'r')` (no encoding given = Windows'
    locale encoding), `json.load`s it into `self.cache_data` and returns that same object (line 24). If the file
    does not exist: returns a NEW `{}` and leaves `self.cache_data` unchanged. On any exception: `logger.error(f"Error loading cache: {e}")`
    and returns `{}`; `self.cache_data` keeps its previous value (RAN). Not shown to the user. Return shape: whatever
    the JSON top level is - not checked to be a dict (RAN: a file containing `[1,2]` is returned as a list).
  - `save_cache(self, data)` (29-48): (1) calls `self.load_cache()` again (line 33) - side effect: re-reads the file and
    replaces `self.cache_data` with the file's content; (2) `recent_folders = existing_cache.get('recent_folders', [])`;
    (3) if `data['source_folder']` is non-empty and not already in the list, inserts it at the front and keeps the first 5
    (lines 38-40); (4) MUTATES the caller's dict: `data['recent_folders'] = recent_folders` (43) (RAN: caller's dict gained
    the key); (5) overwrites the file with `json.dump(data, f)` (45-46) - one line, no indentation, ASCII-escaped
    (default `ensure_ascii=True`, RAN: non-ASCII letters stored as `\uXXXX`). Any exception: `logger.error(f"Error saving cache: {e}")`, swallowed.
    - Not "save current settings" in the sense of merging: keys not in `data` are DROPPED from the file each save
      (the file is replaced by `data` + `recent_folders`). (READ/RAN)
    - If the file is corrupt JSON: load fails, `recent_folders` restarts empty and the file is rewritten (RAN).
    - If the file holds valid JSON that is not an object (e.g. a list): `existing_cache.get` raises, the error is only
      logged, and the file is NEVER rewritten again - settings silently stop saving (RAN).
    - Already-listed folder chosen again is not moved to the front (not true most-recent-first order) (READ/RAN).
    - Called from: main_window.py:371 (inside `GSTOrganizerApp.save_cache`, which builds the dict). That method is
      triggered from main_window.py:196 (dark-mode toggle), gui/handlers/file_handler.py:25, 36, 47, 55 (the four
      Browse buttons after a choice), gui/handlers/client_handler.py:223 (global "include client name" checkbox),
      gui/tabs/setup_tab.py:349 (workflow radio buttons) and :438 (folder-name-length spinbox arrows). (READ)
  - `get_cached_value(self, key, default=None)` (50-52): returns `self.cache_data.get(key, default)`. NO CALLERS FOUND (dead).
  - `set_cached_value(self, key, value)` (54-56): sets `self.cache_data[key] = value` in memory only (no save).
    NO CALLERS FOUND (dead).
- Keys actually stored (READ main_window.py:360-370; RAN file content after a dark-mode toggle):
  `source_folder`, `itc_template`, `sales_template`, `target_folder` (strings, full paths), `processing_mode`
  (default 'fresh'), `workflow_mode` ('organize' or 'full'), `include_client_name` (bool), `dark_mode` (bool),
  `client_name_max_length` (int, default 35), `recent_folders` (list, max 5). RAN example (empty settings):
  `{"source_folder": "", "itc_template": "", "sales_template": "", "target_folder": "", "processing_mode": "fresh", "workflow_mode": "organize", "include_client_name": false, "dark_mode": true, "client_name_max_length": 35, "recent_folders": []}`.
- Readers of the same file (READ):
  - main_window.py:121 (via `load_cache`) and :375-389 `apply_cached_values` (all keys except `recent_folders`).
  - power_query_extractor/gui/extractor_window.py:118 (same path expression) and :590-593 (reads only `target_folder`).
  - power_query_extractor/extractor_main.py:41 (same path, first of four candidates) and :55-58 (tries keys
    `target_folder`, `target_path`, `output_folder`, `last_target`).
  - `recent_folders` is WRITTEN but never READ anywhere (grep) - write-only data. (HIGH)
- Business rules: keep last 5 source folders (line 40). Not a GST rule. Constant 5 is inline, not in constants.py.
- Edge cases handled: missing file (21), any read/write error (25-26, 47-48), blank source folder not added (38).
- Gaps / silent bugs:
  - Non-atomic write (open 'w' truncates first): a crash or full disk mid-write leaves a corrupt/empty file; next start
    silently loses every setting (only a log line). (HIGH by reading)
  - Startup crash from a valid-but-odd file (cross-file, main_window.py:375-387, READ): top-level JSON that is a
    non-empty list or string makes `self.cache_data.get` raise AttributeError; a string value for
    `client_name_max_length` makes `cached_length > 0` raise TypeError. Both happen inside `GSTOrganizerApp()` and end
    at main.py:130-134 "Critical error". Only on a hand-edited/corrupted file (LOW likelihood, HIGH certainty).
  - Two copies of the app at once share one file: last writer wins; each `save_cache` drops keys the other wrote only if
    they are not in its own dict (all writers send the full key set, so in practice last-writer-wins). (MEDIUM)
  - Location is the home folder root, not `%LOCALAPPDATA%\<App Name>\` (Standard 13). See cross-file table.
  - Merge note (READ, B-copy): B uses a different store, `Path.home() / '.gst_organizer'` (B models/cache_manager.py:17)
    and imports from the old `~/.gst_organizer_cache.json` (B models/cache_manager.py:95-98). Both apps on one PC would read
    the same old file. (READ, HIGH)
- Excel touchpoints: none. UI: none. Threads: none (always called on the main thread - all triggers are Tk callbacks).
- Comments: line 2 "Cache Manager for saving and loading application state"; line 32 "Preserve recent folders if they exist";
  line 40 "Keep only last 5". They agree with the code.
- Pattern notes: errors logged at ERROR and swallowed (same style as most of A's helpers). Tests: no test covers this file (grep).
- History: unchanged since edac6a3.

### gui/utils/dark_mode_manager.py (294 lines) - [UI]

- Purpose: switches every Tk widget and eleven ttk styles to a dark colour scheme and back, restoring the colours it
  recorded before changing them. (READ)
- Imports (READ): standard library `logging` (16), `tkinter as tk` (17), `weakref` (18); project `utils.constants.GUI_CONFIG` (20).
  No OS-specific imports.
- Module-level constants (READ):
  - `logger` (22).
  - `THEMED_OPTIONS` (25-29) = ('background', 'foreground', 'insertbackground', 'selectcolor', 'activebackground',
    'activeforeground', 'highlightbackground', 'troughcolor', 'selectbackground', 'selectforeground') - the widget colour
    options that are recorded. Only seven of them are ever changed by `_to_dark` (see below).
  - `FIELD_CLASSES` (32) = {'Entry', 'Text', 'Spinbox', 'Listbox', 'TCombobox'} - get the "field" dark colour.
  - `BUTTON_CLASSES` (33) = {'Button', 'Checkbutton', 'Radiobutton', 'Menubutton'} - get the "button" dark colour.
  - `TTK_STYLES` (37-49): style -> options recorded at start: TFrame(background), TLabel(background, foreground),
    TNotebook(background), TNotebook.Tab(background, foreground), TButton(background, foreground),
    TEntry(fieldbackground, foreground), TCombobox(fieldbackground, foreground), TScrollbar(background, troughcolor),
    TProgressbar(background, troughcolor), Treeview(background, foreground, fieldbackground),
    Treeview.Heading(background, foreground).
- class DarkModeManager (52-294):
  - `__init__(self, root, style)` (55-78): stores root window and ttk Style; `is_initialized=False`, `is_dark_mode=False`;
    `_original` = `weakref.WeakKeyDictionary()` (68: widget object -> {option: value}; entries vanish when the widget
    object is garbage-collected); `_original_ttk = {}` (69); `_original_root_bg = None` (70);
    `dark_colors` (72-78) = bg '#2b2b2b', fg '#ffffff', widget_bg '#3b3b3b', button_bg '#404040', trough '#333333'.
    Called from main_window.py:122 (`DarkModeManager(self.root, self.style)`).
  - `initialize(self)` (81-85): once only: saves `root.cget('bg')` into `_original_root_bg` and calls `_capture_ttk`.
    Called from main_window.py:123 - i.e. BEFORE any tab is built (create_gui is main_window.py:53, after
    setup_components :51). `_original_root_bg` is NEVER read anywhere (grep) - dead state. (HIGH)
  - `_capture_ttk(self)` (87-102): for each TTK_STYLES entry, `style.lookup(style, option)`; on exception skips that option;
    stores the value, storing '' when lookup gives None (so "unset" can be restored, fix 725fbaf).
    RAN (clam theme, Tk 9): captured TFrame bg '#dcdad5', TLabel '#dcdad5'/'black', TNotebook.Tab '#bab5ab'/'black',
    TEntry fieldbackground '' ... Treeview '#ffffff'/'black'/'' etc.
  - `apply_dark_mode(self)` (105-115): sets `is_dark_mode=True`; pass 1: walks the whole tree calling `_capture_widget`
    (record before mutate); then `_apply_ttk_dark`; pass 2: walks again calling `_to_dark`. Called from main_window.py:160.
  - `restore_original_colors(self)` (117-120): `is_dark_mode=False`; `_restore_ttk`; walks tree with `_to_original`.
    Called from main_window.py:162.
  - `_apply_ttk_dark(self)` (123-151): `style.configure` for the 11 styles with the dark colours (TProgressbar bar colour =
    `GUI_CONFIG['colors']['primary']` = '#0078D4'); failures logged at DEBUG. Then `style.map('TNotebook.Tab',
    background=[('selected', '#404040')], foreground=[('selected', '#ffffff')])` inside try/except: pass.
  - `_restore_ttk(self)` (153-172): re-applies every captured style (skips a style whose saved dict is empty); errors DEBUG.
    Then HARD-CODED maps (not captured values): `TNotebook.Tab` selected background '#f0f0f0', foreground 'black';
    `Treeview` selected foreground 'black', background `GUI_CONFIG['colors']['primary']` ('#0078D4'). Exceptions swallowed.
    - RAN: before any toggle the clam theme maps were TNotebook.Tab selected '#dcdad5' (+ default '#bab5ab') and Treeview
      selected bg '#4a6984' / fg '#ffffff'. After one dark->light round trip: TNotebook.Tab selected bg '#f0f0f0' (fallback
      entry dropped) and Treeview selected bg '#0078D4' with BLACK text. So the selected client row and selected tab look
      different after the first toggle than at start, and stay that way also in later dark mode (dark mode never sets a
      Treeview map). This contradicts the module docstring's "replays exactly those values" for ttk maps. These two maps are
      a remnant of the original edac6a3 code (READ: `git show edac6a3:gui/utils/dark_mode_manager.py`). (HIGH)
  - `_walk(self, widget, action)` (175-189): calls `action(widget)`; swallows `tk.TclError` silently and other exceptions
    at DEBUG; then recurses into `widget.winfo_children()` (returns if that fails). Depth-first recursion; includes the root
    and every Toplevel window that is a child of root (RAN: a Toplevel's label went dark).
  - `_supported(self, widget)` (191-197): returns the THEMED_OPTIONS present in `widget.keys()`; () on failure.
    ttk widgets mostly have none of these (they are skipped; their colours come from styles); ttk.Label does have
    background/foreground (RAN).
  - `_capture(self, widget, options)` (199-215): if the widget is ALREADY in `_original`, returns without recording
    (line 202-203) - the first capture is kept for the widget's whole life. TypeError (not weak-referenceable) -> return.
    Records `cget(option)` for each option, ignoring per-option errors.
  - `_capture_widget(self, widget)` (217-221): `_capture` with the widget's supported options. Used in pass 1.
  - `_to_dark(self, widget)` (223-253): captures again (no-op if already captured); chooses surface by class:
    FIELD_CLASSES -> '#3b3b3b', BUTTON_CLASSES -> '#404040', else '#2b2b2b'; sets background=surface, foreground='#ffffff',
    insertbackground='#ffffff', selectcolor='#3b3b3b', activebackground='#3b3b3b', activeforeground='#ffffff',
    troughcolor='#333333' (only those the widget supports). `highlightbackground`, `selectbackground`,
    `selectforeground` are recorded but never darkened (RAN: an Entry's highlight ring stayed '#d9d9d9' in dark mode).
  - `_to_original(self, widget)` (255-268): replays the recorded dict with `configure(**saved)`; TclError swallowed,
    others DEBUG. Widgets never recorded are left as they are.
  - `update_tree_tags(self, tree_widget, is_dark_mode=None)` (271-294): if `is_dark_mode` is None uses `self.is_dark_mode`;
    returns if the object has no `tag_configure`. Dark: tag 'complete' bg '#1a3d1a', 'incomplete' bg '#3d2f1a', both fg
    '#ffffff'. Light: 'complete' bg `GUI_CONFIG['colors']['complete']` ('#E8F5E8'), 'incomplete'
    `GUI_CONFIG['colors']['incomplete']` ('#FFF4E5'), both fg 'black'. Called from main_window.py:178 always WITH an explicit
    value, so the `None` default path (and therefore `self.is_dark_mode` as a read value) is used only by nobody in the app (tests only).
- Callers (READ/grep): only gui/main_window.py (lines 33, 122, 123, 160, 162, 178) and tests/test_dark_mode.py.
  The toggle path: TitleBar checkbox `command=` -> main_window.toggle_dark_mode (134 -> 157-207); also at start-up when the
  cached `dark_mode` is true (main_window.py:388-389).
- Business rules: none (colours only). No GST rules.
- Edge cases handled: widget destroyed mid-walk (179-180); non-weak-referenceable objects (204-205, 214-215, 259-260);
  options a widget lacks (generic discovery 191-197); unset ttk options restored to '' (95-101); widget without
  `tag_configure` (280-281); second `initialize` call (82).
- Gaps / silent bugs:
  1. STALE CAPTURE (RAN, HIGH): because `_capture` records a widget only once (202-203), any colour the app changes on a
     widget in light mode AFTER the first dark toggle is lost on the next restore. RAN: label bg red -> dark -> light (red) ->
     app sets blue -> dark -> light => red again (expected blue). In A this matters for widgets whose colours are set from
     live state: gui/tabs/setup_tab.py:553-556, 566, 606, 608 (checklist rows, accent bars, action strip). A works around it
     only for the Setup tab by calling `setup_tab.refresh_status(force=True)` after every toggle (main_window.py:188-194, commit c89c1d4).
  2. WIDGETS CREATED WHILE DARK ARE NOT THEMED (RAN, HIGH): nothing re-themes new widgets; they appear in light colours until
     the next "switch on". Commit c89c1d4's message says "widgets created during dark mode are themed and restore to their own
     colours" - only true after the next switch-on, not at creation (RAN). A widget created while dark WITH dark colours is
     recorded as "original" on the next switch-on and then stays dark after switching off (RAN: label made with '#2b2b2b' stayed
     '#2b2b2b').
  3. Custom ttk styles (any name not in TTK_STYLES, e.g. 'Accent.TButton') are never darkened (RAN). A itself defines no
     custom styles in the main window (grep: only power_query_extractor/gui/extractor_window.py:92-95, which is only called
     by the standalone extractor window).
  4. Any ttk style option in TTK_STYLES that some code configures after `initialize()` is reverted to the start-up value on
     restore (RAN with TLabel foreground). No current A code does this (grep), so latent only.
  5. Client tree tags (cross-file, READ, HIGH): gui/handlers/client_handler.py:58-59 re-sets the 'complete'/'incomplete' tag
     BACKGROUNDS to the light colours every time the tree is rebuilt (scan / rescan), even in dark mode, while the tag
     FOREGROUND stays '#ffffff' from the last dark `update_tree_tags` (tag_configure only changes the options given). Result in
     dark mode after a scan: white text on pale green / pale amber rows (hard to read) until the next toggle.
  6. Comment claim not reproduced (RAN, MEDIUM): lines 63-67 (and commit aeb83cf) say Tk reuses widget path names after a
     widget is destroyed. In Python 3.14 tkinter auto-names are counter-based (`.!frame.!label`, then `.!frame.!text`, new parent
     `.!frame2`) and A creates no explicitly named widgets (grep), so reuse could not be reproduced. Keying by object is still fine.
  7. Recursion over the whole tree twice per switch-on and once per switch-off: fine for hundreds of widgets; no measurement made.
  8. Thread safety: only called from Tk callbacks on the main thread (READ). OK.
- Comments/docstrings: module docstring (2-14) describes the previous implementation's bugs and claims exact replay - true for
  widget colours, NOT for the two ttk maps (point above). Lines 35-36 "The old version covered only four" - the edac6a3 version
  stored six style names (TFrame, TLabel, TNotebook, TNotebook.Tab, Treeview, Treeview.Heading = four widget families) (READ git
  show), so "four" is right only if counted by family (LOW importance). `update_tree_tags` docstring (272-277) describes the old else-branch bug - matches git history (aeb83cf).
- Pattern notes: DEBUG logging + swallow for theming failures; contrasts with ERROR logging in cache_manager. No user-visible errors.
- Tests: tests/test_dark_mode.py (11 tests; RAN: all pass on Linux).
- History: edac6a3 original (class-based guessing restore), aeb83cf rewrite (26-07-2026, +258/-183 lines in this file, plus a new
  165-line tests/test_dark_mode.py), 725fbaf (record empty ttk values). Related main_window changes c89c1d4, 725fbaf; setup_tab 169c7b5.

### gui/utils/setup_validation.py (115 lines) - [UI] (pure logic used by the Setup tab)

- Purpose: decides, for the four Setup inputs (source folder, ITC template, Sales template, target folder), whether each is
  OK / missing / invalid, whether Scan and Processing may proceed, and the one-line readiness text. (READ)
- Imports (READ): standard library `os` (8, UNUSED), `pathlib.Path` (9), `typing` Callable, Dict (UNUSED), List, NamedTuple,
  Optional (10). No OS-specific modules. (RAN: AST check found `os` and `Dict` unused.)
- Module-level constants (READ): states `OK='ok'` (13), `MISSING='missing'` (14), `INVALID='invalid'` (15); stages
  `FOR_SCAN='scan'` (18), `FOR_PROCESSING='processing'` (19); `TEMPLATE_SUFFIXES = {'.xlsx', '.xltx', '.xltm', '.xlsm'}` (21).
- `class Check(NamedTuple)` (24-33): fields key, label, state, message, blocks; property `is_ok` = `state == OK` (31-33).
- `_default_exists(path)` (36-44): False for empty/whitespace (39-40, because `Path('')` means the current folder); else
  `Path(path).exists()`, returning False on OSError/ValueError (ValueError = e.g. embedded NUL). Used only when no `exists`
  is injected (line 79); the app always injects setup_tab's own cached checker (gui/tabs/setup_tab.py:537-540), so in the app
  this function is used only by tests (tests/test_setup_validation.py:108-112). (READ)
- `_check_folder(key, label, value, blocks, exists)` (47-53): blank -> MISSING 'Not selected yet'; `not exists(value)` ->
  INVALID 'Folder not found'; else OK with message `_short(value)`. It does NOT check that the path is a folder (RAN: a FILE
  path is reported OK). Downstream, core/file_parser.py:154-155 raises "Not a directory" and file_handler shows a "Scan Error"
  box (READ, cross-file).
- `_check_template(...)` (56-64): blank -> MISSING 'Not selected yet'; missing -> INVALID 'File not found'; suffix (lower-cased)
  not in TEMPLATE_SUFFIXES -> INVALID 'Not an Excel template'; else OK with message = file name only. Does not check it is a
  file (RAN: a FOLDER named `x.xlsx` is OK). The check order means a missing file with a wrong suffix says 'File not found'.
- `_short(value, width=46)` (67-72): if longer than 46 chars returns `"..." + last 42 chars` = 45 chars (RAN). Docstring says
  "Trim the middle" - code trims the START (keeps the tail). `width` is never passed (always 46).
- `evaluate_setup(source_folder, itc_template, sales_template, target_folder, exists=None)` (75-85): returns 4 Checks in display
  order: ('source','Source folder', FOR_SCAN), ('itc','ITC template', FOR_PROCESSING), ('sales','Sales template', FOR_PROCESSING),
  ('target','Target folder', FOR_PROCESSING). Called from gui/tabs/setup_tab.py:537 (refresh_status); tests.
- `can_scan(checks)` (88-91): all FOR_SCAN checks OK. Called setup_tab.py:586.
- `can_process(checks)` (94-97): all checks OK. Called setup_tab.py:579 (pill colour) and :598 (hint text). Note: in A this result
  only colours the pill and chooses a hint; it does not enable/disable any processing button here (READ setup_tab 579-608).
- `blocking_labels(checks, stage)` (100-106): labels not OK; FOR_SCAN -> only scan items; ANY other value -> all items.
  Called setup_tab.py:595 (FOR_SCAN) and :602 (FOR_PROCESSING).
- `summary_line(checks)` (109-115): "Ready - all 4 items set" when all OK, else "N of 4 items set". Called setup_tab.py:580.
- Business rules (READ):
  - Scanning needs only a valid source folder (88-91, 81). Processing needs all four inputs (94-97).
  - Accepted template types .xlsx, .xltx, .xltm, .xlsm (21, 62). Workflow rules, not tagged [GST RULE]: they decide which
    inputs are required and which file types are accepted, not where any figure goes. (If the author treats "an ITC template
    and a Sales template are both required" as a GST rule, tag it.)
- Edge cases handled: blank/whitespace (39, 49, 58), bad path characters (43), case-insensitive suffix (62).
- Gaps:
  - Inconsistent accepted types (READ, HIGH): this file accepts 4 suffixes, but the Browse dialogs offer only
    `"*.xltx *.xlsx"` (gui/handlers/file_handler.py:31, 42). A .xlsm/.xltm can only be typed in. Whether core/excel_handler
    handles macro workbooks correctly is for A-R1 (question for the author).
  - Folder/file kind not checked (RAN, see above).
  - Staleness (cross-file, READ, MEDIUM-HIGH): setup_tab caches each path's existence (setup_tab.py:521-529) and clears the cache
    only at start (setup_tab.py:156) and on a dark-mode toggle (main_window.py:191). A folder created, deleted or disconnected
    (network drive) after the first check keeps its old state until restart, a theme toggle or a different path string.
  - Test failing on Linux (RAN): tests/test_setup_validation.py:65-66 expects `'itc.xltx'` from `C:\gst\itc.xltx`; on Linux
    `Path().name` keeps the whole string because backslash is not a separator. Windows-only test assumption; not a code bug on
    Windows (HIGH).
- UI text produced (READ): 'Not selected yet', 'Folder not found', 'File not found', 'Not an Excel template', the labels above,
  summary lines above.
- Comments: docstring (1-7) "Deliberately free of Tkinter" (true for the file; false via package import - see gui/utils/__init__.py);
  "~45ms" per existence check on a network share (author's measurement, not verifiable here). `can_scan` docstring "mirrors
  FileHandler.validate_scan_inputs" - true: file_handler.py:120-130 checks only non-empty + exists (READ).
- Files: none written; only `Path.exists()` reads (through the injected checker). Excel: none. Threads: none.
- History: added 322717f (25-07-2026), unchanged since. Tests: tests/test_setup_validation.py (RAN: 18 pass, 1 fails on Linux as above).

### gui/utils/status_bar.py (37 lines) - [UI]

- Purpose: the dark strip at the bottom of the main window with a status message on the left and a version text on the right. (READ)
- Imports (READ): `tkinter as tk` (4); project `utils.constants.GUI_CONFIG` (5).
- class StatusBar (8-37):
  - `__init__(self, parent)` (11-13): stores parent; calls `create_status_bar()`. Instantiated main_window.py:149 with `self.root`.
  - `create_status_bar(self)` (15-33): Frame bg `GUI_CONFIG['colors']['dark']` ('#323130'), height 30, packed `fill='x', side='bottom'`,
    `pack_propagate(False)` (fixed height). Left label `self.status_label`: text "💡 Ready to organize your GST files!", font
    ('Arial', 9), bg '#323130', fg 'white', padx 10, pady 6. Right label (local variable only): text "v3.0 | Production Ready",
    font ('Arial', 9), bg '#323130', fg '#BDBDBD'.
  - `update_status(self, message)` (35-37): sets label text to `f"💡 {message}"`. Called from main_window.py:281
    (`GSTOrganizerApp.update_status`), which is called from main_window.py:338 (tab change "Step N: <name>"), :393 ("Welcome to GST
    File Organizer v3.0"), gui/handlers/file_handler.py:24, 35, 46, 54, 63, 68, 76, 84, 90 and gui/handlers/client_handler.py:131,
    138, 145, 159. All are main-thread callers (READ: scan runs synchronously on the main thread, core/file_parser.py:165-166 calls
    the progress callback in the loop).
- Gaps:
  - RAN: at start-up the status already reads "💡 Step 1: Setup" - the "Welcome..." text set by run() (main_window.py:393) is
    immediately overwritten by the first tab-changed event (MEDIUM: observed on Linux/Xvfb).
  - RAN: at the minimum window size 1000x700 the status bar is NOT visible (`winfo_ismapped` = 0). Cause: it is packed after the
    notebook (main_window.py:138 then :149), and with pack the last-packed widget loses space first. At 1200x800 it is visible. (HIGH)
  - During a scan the label text is changed many times on the main thread without any repaint call, so "Scanning: ... (n/N)" is
    not seen until the scan returns (READ, MEDIUM-HIGH; file_handler.py:67-72).
  - Version text hard-coded "v3.0 | Production Ready" - see version notes in Cross-file.
  - In dark mode: frame and labels become '#2b2b2b', both texts white (version text loses its grey) (READ dark_mode_manager rules).
- No files, no Excel, no GST rules. No tests cover this file (grep). History: unchanged since edac6a3.

### gui/utils/ui_helpers.py (63 lines) - [UI]

- Purpose: static helper functions for the status label, the processing log box, the progress display and coloured section cards. (READ)
- Imports (READ): `tkinter as tk` (4), `datetime.datetime` (5).
- class UIHelpers (8-63), all `@staticmethod`:
  - `update_status(status_label, message)` (11-14): sets `text=f"💡 {message}"`. NO CALLERS FOUND (dead; duplicates
    StatusBar.update_status).
  - `log_message(log_widget, message, level='normal')` (16-25): takes the time NOW as HH:MM:SS, builds `"[HH:MM:SS] message\n"`,
    then schedules `_add_log_message` with `log_widget.after(0, lambda: ...)`. The lambda captures this call's own local values,
    so there is NO late-binding problem (READ, HIGH). Called from main_window.py:285 (`GSTOrganizerApp.log_message`), which is
    called (among others) DIRECTLY from the processing worker thread (gui/handlers/processing_handler.py:319-523, inside
    `process_files_thread` started at :129-134) and also via `root.after` (processing_handler.py:31). So `log_widget.after(...)`
    itself runs on the worker thread: this relies on Tkinter forwarding calls made from other threads to the main thread
    (works with a "threaded" Tcl build). RAN: three calls from a worker thread on Linux/Tk 9 worked. Windows python.org builds:
    not verifiable here (MEDIUM).
  - `_add_log_message(log_widget, message, level)` (27-33): enables the Text widget, inserts the message with tag `level`,
    scrolls to end, disables it again. Tags configured in gui/tabs/processing_tab.py:113-117 ('success' '#4CAF50', 'warning'
    '#FF9800', 'error' '#F44336', 'info' '#2196F3', 'normal' '#FFFFFF'). An unknown level just inserts unstyled text. If the
    widget was destroyed (app closing) Tk prints a TclError to the console (READ, MEDIUM). The log box is never trimmed (grows
    for the whole session).
  - `update_progress(progress_var, progress_label, current_op_label, value, message)` (35-40): sets the progress variable,
    label `f"📊 Progress: {value:.1f}%"`, and the current-operation label text. Called from main_window.py:289 inside
    `root.after(0, ...)`. RAN: a string value raises ValueError ("Unknown format code 'f'"); numbers are fine.
  - `create_colored_section(parent, title, description, color)` (42-63): white Frame, relief solid, border 2, PACKED immediately
    `fill='x', pady=10`; a header Frame bg=color height 40 (fixed); title Label ('Arial', 12, 'bold') white on color, centred;
    description Label ('Arial', 10), bg white, fg '#323130', wraplength 500, pady 10. Returns the section frame (RAN: children
    Frame + Label). Called from gui/tabs/setup_tab.py:354 ("📂 FOLDERS & TEMPLATES", COLORS['primary']) and :394
    ("⚙️ PROCESSING MODE", COLORS['info']). In dark mode the coloured header becomes '#2b2b2b' (generic theming) (READ).
- Gaps: dead `update_status`; helper packs itself (callers cannot use grid); timestamp is local time with no date.
- No files, no Excel, no GST rules. No tests directly (setup_tab smoke tests use create_colored_section indirectly). History: unchanged since edac6a3.

### gui/widgets/__init__.py (7 lines) - [UI]

- Purpose: package file re-exporting CollapsibleFrame and TitleBar. (READ)
- Imports: line 4 `from .collapsible_frame import CollapsibleFrame`, line 5 `from .title_bar import TitleBar`; line 7
  `__all__ = ['CollapsibleFrame', 'TitleBar']`.
- Callers of the re-exports: NONE (all users import submodules: main_window.py:24-25, validation_tab.py:7, setup_tab.py:22). (HIGH)
- `Tooltip` (gui/widgets/tooltip.py, added 322717f) is not re-exported - list not maintained. No GST rules. Unchanged since edac6a3.

### gui/widgets/collapsible_frame.py (41 lines) - [UI]

- Purpose: a frame with a small ▼/▶ button that hides or shows its inner area. (READ)
- Imports: `tkinter as tk` (4).
- class CollapsibleFrame(tk.Frame) (7-41):
  - `__init__(self, parent, title="", **kwargs)` (9-30): outer Frame gets `**kwargs`; `self.show = tk.IntVar()` set to 1
    (expanded); `title_frame` (default colours) packed fill x expand 1; `toggle_button` text "▼", width 2, relief flat, font
    ('Arial', 10), `command=self.toggle`; title Label font ('Arial', 11, 'bold'), padx 5; `sub_frame` relief sunken, border 1,
    packed fill both expand 1, padx 10, pady 5.
  - `toggle(self)` (32-41): if shown -> `pack_forget()` the sub_frame, button "▶", show=0; else re-pack with same options,
    button "▼", show=1. RAN: collapse gives ('▶', not mapped, 0); expand gives ('▼', mapped, 1).
- Callers: gui/tabs/validation_tab.py:46-48 (`CollapsibleFrame(parent, "📋 VALIDATION INSTRUCTIONS", bg='#FFF3E0', relief='solid',
  borderwidth=2)`), instructions added to `.sub_frame` (validation_tab.py:59-65). main_window.py:25 imports it but never uses it (READ).
- Gaps (READ, HIGH): `bg` passed by the caller colours only the outer frame; the title strip, button, title label and inner frame
  keep the default system grey - visual mismatch on the orange card. `IntVar()` without a master uses the default Tk root (fine
  with one root).
- No files, no Excel, no GST rules. No tests. Unchanged since edac6a3.

### gui/widgets/title_bar.py (51 lines) - [UI]

- Purpose: the blue banner at the top of the main window with the app title, subtitle and the "Dark Mode" checkbox. (READ)
- Imports: `tkinter as tk` (4); `utils.constants.GUI_CONFIG` (5).
- class TitleBar (8-51):
  - `__init__(self, parent, dark_mode_var, toggle_callback)` (11-15): stores them; calls `create_title_bar()`.
    Instantiated main_window.py:134 with (root, app.dark_mode BooleanVar, app.toggle_dark_mode).
  - `create_title_bar(self)` (17-51): `title_frame` bg primary '#0078D4', height 70 fixed (pack_propagate False), pack fill x.
    `dark_mode_frame` bg primary, placed with `place(x=self.parent.winfo_width()-150, y=15)` (27). Checkbutton "🌙 Dark Mode",
    variable = dark_mode_var, `command=self.toggle_callback`, font ('Arial', 10, 'bold'), bg primary, fg 'white', selectcolor primary.
    Title Label "🏢 GST FILE ORGANIZER & REPORT GENERATOR", ('Arial', 18, 'bold'), white on primary, `pack(expand=True)`.
    Subtitle "Organize files and generate Excel reports automatically", ('Arial', 11), fg '#E3F2FD'.
- Gaps:
  - FIXED POSITION (RAN, HIGH for logic; Windows pixel values not verified): x is computed once from the window width at creation.
    RAN on Xvfb: x = 1050 (1200-150). Resizing to 1600 wide left it at 1050 (not at the right edge); resizing to the minimum
    1000 wide put it at 1050 > 1000, i.e. OFF-WINDOW - the Dark Mode checkbox cannot be seen or clicked at minimum width.
    If on some PC the width were not yet known at creation (winfo_width = 1), x would be -149 (off the left edge) (LOW-MEDIUM; on
    Xvfb the width was already 1200 because main_window.py:67 forces a layout pass first).
  - Mixed `colors[...]` local and `GUI_CONFIG['colors'][...]` access in the same function (style only).
  - Dark mode: main_window.py:165-174 additionally forces the title frame and its DIRECT children back to primary on switch-off
    (redundant with the manager's restore; the checkbox is a grandchild and is restored by the manager) (READ).
- UI texts as quoted above. No files, no Excel, no GST rules. No tests. Unchanged since edac6a3.

### gui/widgets/tooltip.py (69 lines) - [UI]

- Purpose: small yellow pop-up help text shown when the mouse rests on a widget. (READ)
- Imports: `tkinter as tk` (2).
- class Tooltip (5-69):
  - `__init__(self, widget, text='', delay=450, wraplength=520)` (12-22): stores settings; `_after_id=None`, `_window=None`;
    binds `<Enter>` -> `_schedule`, `<Leave>` -> `_hide`, `<ButtonPress>` -> `_hide`, all with `add='+'` (adds to, does not
    replace, other bindings on the widget - important because setup_tab uses Enter/Leave for scrolling, commit bf2f1dd).
    `delay` and `wraplength` are never passed by callers (always 450 ms / 520 px).
  - `set_text(self, text)` (24-27): `self.text = text or ''`; if the pop-up is visible, hides it (does not re-show).
    Called from gui/tabs/setup_tab.py:576 (path entries, on every refresh).
  - `_schedule(self, _event=None)` (30-33): cancels any pending show; if text is non-empty, `after(delay, _show)`.
  - `_cancel(self)` (35-41): `after_cancel` (errors swallowed), clears id.
  - `_show(self)` (43-60): returns if already shown or text empty; position x = widget's screen x + 12, y = widget's screen y +
    height + 6; borderless Toplevel (`wm_overrideredirect(True)`); Label justify left, bg '#FFFFE0', fg '#333333', relief solid,
    border 1, font ('Segoe UI', 9), wraplength, padx 6, pady 4. Any exception -> `_window=None` ("A tooltip must never break the app").
    RAN: pop-up appeared with bg '#FFFFE0'; empty text never schedules.
  - `_hide(self, _event=None)` (62-69): cancel + destroy window (errors swallowed).
- Callers (all gui/tabs/setup_tab.py): 314 (each checklist row and its 3 labels), 336 (workflow radios), 381 (path entries,
  stored in `_path_tooltips`), 386 (Browse buttons), 409 (processing-mode radios), 423 (client-name checkbox), 444 (length
  spinbox), 462 (Scan), 467 (Re-scan). Tooltip objects are kept alive by Tk's bindings even when not stored (READ, HIGH).
- Gaps (READ): no clamping to the screen edge (can open partly off-screen near the bottom/right); 'Segoe UI' is a Windows font.
  Pop-ups are created fresh each time in light colours (fine in dark mode).
- No files, no Excel, no GST rules. Tests: exercised indirectly by tests/test_setup_tab_smoke.py (RAN: 13 pass).
  History: added 322717f (25-07-2026), unchanged since.

### main.py (137 lines) - [CORE] (entry point)

- Purpose: the program you run to start the main app: it sets up console output and logging, creates some folders, checks
  dependencies, offers to install missing ones (then restarts), and opens the main window. (READ)
- How the app is started (READ): `python main.py` / `py main.py` from a console, or double-clicking main.py (Windows opens it with
  the console python.exe). There is no .bat launcher, no shortcut and no pythonw use in A. A second, separate entry point
  `launch_extractor.py` (A-R2's file) starts only the extractor and does NOT run these checks.
- Imports (READ): `sys` (9), `io` (10), `os` (11, UNUSED - RAN AST check), `logging` (12), `pathlib.Path` (13). Later, inside
  `main()`: `utils.requirements_check` (68-70, 120), `utils.dependency_installer` (82), `gui.main_window.GSTOrganizerApp` (109).
  No OS-specific modules here.
- Module-level code (runs on import, before `main()`):
  - Line 17: `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')` (re-wraps console output as UTF-8 so emoji print).
    - Under pythonw.exe (no console) `sys.stdout` is None, so this line raises AttributeError BEFORE logging exists: the app dies
      with no window, no message and no log. RAN: with `sys.stdout = None`, running main.py ends in
      "AttributeError: 'NoneType' object has no attribute 'buffer'" and no log file is created. (HIGH) => A REQUIRES a console.
    - The new wrapper is not line-buffered: RAN in a pseudo-console, text printed before a 2-second pause appeared only at exit.
      So the banner and the requirements report (lines 62-73) may stay invisible in the console while the install dialog is shown;
      they appear at the next flush (a logging call, `input()`, or exit). (RAN on Linux, HIGH that Windows behaves the same:
      buffering is in the Python layer.)
  - Lines 20-27: `logging.basicConfig(level=INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler('gst_organizer.log'), logging.StreamHandler(sys.stdout)])`.
    - Log file path = `gst_organizer.log` in the CURRENT working folder (wherever the console was, or the folder of main.py when
      double-clicked, or a shortcut's "Start in"). RAN: the log appeared in my scratch working folder, not in the app folder.
    - Append mode, no size limit/rotation, no encoding given (Windows locale encoding, typically cp1252). Log lines containing
      emoji exist (power_query_extractor/core/report_processor.py:316, 339, 377, 389, 553, 571, 592, 597, 616). RAN (simulated
      cp1252): such a line is NOT written to the file and Python prints "--- Logging error ---" with a traceback to the console;
      plain lines are written. (MEDIUM that the PC's locale is cp1252; HIGH for the effect.)
    - Created at import time even if the app then fails; if the current folder is read-only, `FileHandler` raises at import and the
      app dies with a raw traceback (READ, HIGH).
- `setup_environment()` (29-43): `project_root = Path(__file__).parent`; `sys.path.insert(0, str(project_root))`; for 'core',
  'utils', 'gui', 'logs', 'temp': `mkdir(exist_ok=True)` under project_root; for core/utils/gui, writes `__init__.py` with
  `"""Package initialization"""` if missing. Side effects: creates `<app folder>\logs` and `<app folder>\temp` on first run (RAN:
  both appeared in my scratch copy; they are .gitignored so a fresh copy lacks them). NOTHING uses `logs/` or `temp/` (grep: the log
  goes to the current folder; core/excel_handler.py:53 uses the system temp folder `gst_excel_temp`). On a read-only app folder
  the `mkdir` raises PermissionError -> caught at 130-134 -> "Critical error" and exit 1 (READ, HIGH). Called only from main():60.
- `_pause(message="\nPress Enter to exit...")` (45-54): `input(message)`, ignoring EOFError, KeyboardInterrupt, OSError (added
  be6418c). Called from main.py:94, 101, 126, 133.
- `main()` (57-134), step by step (READ; flow RAN on Linux with patched dialogs):
  1. `setup_environment()` (60).
  2. Prints a 60-char "=" line, "🚀 GST File Organizer v3.0 - Production Ready", "=" line (62-64).
  3. `check_requirements()` and prints `format_report(results)` then "=" line (68-73).
  4. `blocking = missing_essential(results)`, `optional = missing_optional(results)` (75-76).
  5. If either is non-empty (78-95): `offer_install(blocking + optional, requirements_path=<app folder>\requirements.txt if it exists
     else None)` (84-88). If it returns True: prints "\n✅ Packages installed - restarting…"; if `restart_application()` succeeds ->
     `sys.exit(0)`; else prints "Could not restart automatically. Please start the app again.", pauses, `sys.exit(0)` (exit code 0
     although the app did not start).
  6. If `blocking` (97-102): `logging.error("Required package missing: %s", package)` per item, pause, `sys.exit(1)`.
  7. For each optional missing: `logging.warning("%s not installed - %s", package, purpose)` (104-106).
  8. Imports the GUI (109), builds `GSTOrganizerApp()` (111) and calls `.run()` (112) which enters the Tk event loop.
  - `except ImportError` (114-127): logs "Import Error: {e}" with traceback; prints "\n❌ Could not start: {e}", "\nThis usually
    means a required package is missing for THIS Python,", "or a source file has been moved.", then (if possible)
    "\n   Running: Python X.Y.Z at <exe>", "\n   Try:  pip install -r requirements.txt", "   Note 'py' and 'python' can be
    different interpreters."; pause; exit 1. The suggested command uses a RELATIVE requirements.txt (works only if the console is
    in the app folder).
  - `except Exception` (130-134): logs "Critical error: {e}" with traceback; prints "\n💥 Critical error: {e}"; pause; exit 1.
  - `sys.exit()` inside the try raises SystemExit, which is not an `Exception`, so it passes through (READ, HIGH).
  - Errors inside button clicks after the window opens are NOT caught here (Tk prints them to the console and carries on) and
    errors in worker threads are not caught here either; there is no `sys.excepthook` / `threading.excepthook` (READ, HIGH).
- RAN observations (Linux, venv-A without pywin32, dialogs patched):
  - pywin32 counts as "optional" off Windows, so the install offer appeared. Offer text (exact): title "Install missing packages?",
    body "The app needs these packages:\n\n• pywin32 - Excel automation: preserves Power Query when creating reports, and powers the
    refresh in Step 4\n    Reports would be created WITHOUT their Power Query connections, and Step 4 could not refresh them.\n\n
    Install them now for:\n<python path>\n\nThe app will restart afterwards."
  - Accepting ran exactly `[<sys.executable>, '-m', 'pip', 'install', '-r', '<app folder>/requirements.txt']` with
    `capture_output=True, text=True, timeout=600`, then `Popen([<sys.executable>, '<app folder>/main.py'])`, then exit code 0.
  - Declining continued into the GUI (warning logged). A second Tk root created after the dialog's root was destroyed worked here.
  - Non-Windows loop (READ+RAN, HIGH): pip skips pywin32 off Windows (marker `sys_platform == 'win32'`), reports success, the app
    restarts, pywin32 is still missing, and the offer appears again - on every start for as long as the user accepts.
- What it changes on the PC (READ): creates `logs\` and `temp\` (and possibly `__init__.py` files) in the app folder; appends to
  `gst_organizer.log` in the current folder; (after consent) installs packages from the internet into the running Python - see
  utils/dependency_installer.py; starts a new copy of itself.
- Business rules: none, no GST rules.
- Comments/docstrings: docstring (1-7) "GST File Organizer & Report Generator v3.0", "Author: Advanced AI Assistant",
  "Version: 3.0 (Production Ready)". Comment at 66-67 (check before importing GUI) agrees with code. Comment 79-81 ("pywin32 only
  becomes importable in a freshly started interpreter") is the author's rationale, not verifiable here.
- Pattern notes: user messages are console `print`s plus `logging`; only the install offer uses message boxes. Logging calls
  use the root logger (`logging.error`) while library modules use `logging.getLogger(__name__)`.
- Tests: none import or run main.py (grep: no `import main`/`runpy` in tests). (HIGH)
- History: edac6a3 original (same stdout/log/setup lines; ImportError branch listed source files incl. `gui/app.py`, which did not
  exist even then - edac6a3 has gui/__init__.py and gui/main_window.py only at that level, RAN `git ls-tree`); 5ea77c2 removed the shebang line `# !/usr/bin/env python3` (which was malformed anyway: space after `#`);
  f363946 added the requirement check and new ImportError text; be6418c added `_pause`; a5ad487 added the install offer.

### pytest.ini (6 lines) - [CONFIG] (test settings)

- Content (READ): `[pytest]`; `testpaths = tests`; `python_files = test_*.py`; `python_classes = Test*`;
  `python_functions = test_*`; `addopts = -q` (quiet output; also cancels a `-v` given on the command line - RAN).
- No `pythonpath = .`: the project folder is put on the import path by tests/conftest.py:9 instead (READ). Standard 13 asks for
  `pythonpath = .` and `testpaths = tests` because the portable Python does not add the project folder itself.
- History: added 6c1e1c6 (23-07-2026), unchanged. B has no pytest.ini (RAN: `git ls-tree B/main`).

### requirements.txt (19 lines) - [CONFIG]

- Content (READ): line 1 "# GST File Organizer v3.0 Requirements"; line 2 "# Python 3.7 or higher required";
  line 5 `openpyxl>=3.1.2` (comment line 4 "Core dependency for Excel operations");
  line 9 `pywin32>=305; sys_platform == 'win32'` (comments 7-8 "Required on Windows for Excel COM automation / Power Query refresh.
  power_query_extractor and excel_handler import pythoncom/win32com/win32gui.");
  line 12 `lxml>=4.9.0` ("Optional but recommended for better performance");
  line 15 `pytest>=7.0.0` ("Testing (run `pytest` from the repo root)");
  lines 17-19 commented out `# black>=23.0.0`, `# pylint>=2.17.0` ("Development dependencies (optional)"). No final newline.
- Every dependency, purpose and where used (READ/grep):
  | Package | Spec | Purpose | Imported where |
  |---|---|---|---|
  | openpyxl | >=3.1.2 | read/write Excel without Excel | core/excel_handler.py:31-33, power_query_extractor/core/data_consolidator.py:8-10 (unconditional) |
  | pywin32 | >=305, Windows only | Excel through COM, window handling | core/excel_handler.py:22-24 (Windows check + try/except), power_query_extractor/core/report_processor.py:20-37 (try/except), report_processor.py:762 and utils/helpers.py:584 (inside functions) |
  | lxml | >=4.9.0 | NOT imported by A; openpyxl uses it automatically when present (openpyxl/xml/__init__.py `LXML = lxml_available() and OPENPYXL_LXML env == "True"`, RAN read of installed openpyxl 3.1.5) | none directly |
  | pytest | >=7.0.0 | runs tests/ | tests only |
  | tkinter | (standard library, not listed) | the GUI | gui/*, utils/dependency_installer.py:96-97 |
  Indirect packages pip pulls in (RAN: venv-A listing): et_xmlfile (for openpyxl); iniconfig, packaging, pluggy, pygments (for pytest).
- The comment on line 8 is only partly accurate: core/excel_handler.py imports win32com.client and pythoncom (not win32gui);
  report_processor imports pythoncom, win32com.client, win32api, win32con, win32gui; helpers imports win32api. All are guarded
  since c463256, although commit 064c634 described them as unconditional at the time. (READ)
- Python version: "3.7 or higher" (line 2). No syntax newer than 3.7 found by a grep of non-test code (walrus, `list[...]`
  annotations, `match`, `X | None`) (RAN grep, MEDIUM - not exhaustive). The installed test versions declare Requires-Python:
  openpyxl 3.1.5 >=3.8, lxml 6.1.3 >=3.8, pytest 9.1.1 >=3.10 (RAN: read package metadata). Commit messages say the author's PC
  has Python 3.14 (python.exe and python3.14t.exe) (READ f363946, be6418c).
- Gaps: only minimum versions (no upper limits, no exact pins) - a future major openpyxl could be installed by the install offer;
  includes pytest, so the install offer puts a test tool on users' PCs; Standard 13 asks for exact versions and no pytest.
- History: edac6a3 (openpyxl, lxml; pytest commented); 6c1e1c6 uncommented pytest; 064c634 added pywin32 line.

### utils/dependency_installer.py (158 lines) - [UTIL]

- Purpose: shows a Yes/No offer to install missing packages with pip into the running Python, runs pip, shows its output, and can
  relaunch the app. (READ)
- Imports (READ): `logging` (15), `subprocess` (16), `sys` (17), `pathlib.Path` (18, UNUSED - RAN AST). Inside `offer_install`
  (96-97): `tkinter as tk`, `tkinter.messagebox`, `tkinter.scrolledtext`, guarded by try/except returning False.
- Module-level: `logger` (20); `INSTALL_TIMEOUT_SECONDS = 600` (23) - pip is killed after 10 minutes.
- `build_pip_command(packages=None, requirements_path=None)` (26-38): `[sys.executable, '-m', 'pip', 'install']` + `['-r', path]` if
  given + package names if given. Called from install:43 and offer_install:142 (manual command text); tests.
- `install(packages=None, requirements_path=None, log=None)` (41-60): logs "Running: <command joined by spaces>\n" to the callback;
  `subprocess.run(command, capture_output=True, text=True, timeout=600)`; TimeoutExpired -> (False, "pip did not finish within
  600 seconds."); any other exception -> (False, "Could not run pip: {exc}"); else sends stdout+stderr to the callback and returns
  (returncode == 0, output). Called from offer_install:133; tests.
  - Output is collected and shown only AFTER pip ends - there is no live streaming (READ, HIGH). Commit a5ad487 says "A window
    streams pip's output so it never looks frozen"; the test `test_progress_is_streamed_to_the_callback`
    (tests/test_dependency_installer.py:78-84) only checks that output reaches the callback after completion.
  - pip runs on the window's own thread with no event processing, so the "Installing…" window cannot repaint or respond while pip
    runs (up to 600 s); on Windows it is likely to show "Not Responding" (READ HIGH for the blocking; MEDIUM for how Windows shows it).
  - `text=True` decodes pip output with the locale encoding; unusual characters could raise UnicodeDecodeError inside
    `subprocess.run`, which the generic except turns into "Could not run pip: ..." (LOW).
- `restart_application()` (63-74): `subprocess.Popen([sys.executable] + sys.argv)`; True on success; on exception logs
  "Could not restart automatically: %s" and returns False. Called from main.py:91.
  - The new process gets the script path and arguments but NOT interpreter options (e.g. `-B`, `-s`), because `sys.argv` does not
    contain them (READ, HIGH). It inherits the current folder and the console.
- `describe_missing(results)` (77-85): lines "• {package} - {purpose}" plus "    {consequence}" when set. Called offer_install:112; tests.
- `offer_install(results, requirements_path=None, parent=None)` (88-158):
  1. Imports tkinter (failure -> return False) (95-99).
  2. `packages` = package names; `blocking` = any essential (101-102).
  3. Creates a hidden Tk root when no parent (106-108). `parent` is never passed by any caller (main.py:85-88) - always None.
  4. Yes/No box "Install missing packages?" (text quoted under main.py) (110-118). No -> return False.
  5. Toplevel "Installing…", 760x380, ScrolledText font ('Consolas', 9), wrap word (121-125); inner `log()` (127-130) inserts text,
     scrolls, `update_idletasks()`.
  6. Logs "Installing into: <exe>\n\n"; runs `install(packages=None if requirements_path else packages, requirements_path=...)`
     (133-134). With requirements.txt present (always, in a normal copy) it installs EVERYTHING in requirements.txt that is not
     already satisfied - including pytest and lxml - not just the missing package.
  7. Success: logs "\n\n✅ Installed successfully. Restarting the app…" and returns True (136-139) (the window is destroyed almost
     at once in `finally`).
  8. Failure: logs "\n\n❌ Installation failed."; error box "Installation failed": "The packages could not be installed
     automatically.\n\nPlease run this yourself:\n\n{manual}\n\n" + "The app cannot start without them." when blocking; returns False.
     `manual` = `build_pip_command(packages, requirements_path)` joined with spaces (142) - (a) it contains BOTH `-r requirements.txt`
     AND the package names, unlike the command actually tried; (b) paths are not quoted, so a Python or app path containing spaces
     breaks when pasted into a console (READ, HIGH).
  9. Any exception -> `logger.error("Install offer failed: %s", ..., exc_info=True)`, return False (150-152). `finally` destroys
     the hidden root (153-158).
- What it changes on the PC (READ, after consent): runs `<python> -m pip install -r <app folder>\requirements.txt`, which downloads
  from PyPI (internet/proxy needed) and installs into that Python's site-packages (if not writable, pip normally falls back to the
  user's own site-packages - pip behaviour, MEDIUM). Packages: openpyxl, pywin32 (Windows), lxml, pytest and their dependencies.
  Then starts a second process.
- Standard 13 says "Nothing is installed on users' PCs. Every app shares the one `_runtime` folder." This module's purpose is the
  opposite (installs into the user's Python). Factual difference; decision for the author.
- Excel: none. GST rules: none. Threads: none (blocking on the main thread).
- Comments: docstring 1-14 (offer, not silent; restart needed because pywin32 adds its DLL folder via a start-up hook) agree with
  the code, except "streams" (commit message) as above.
- Tests: tests/test_dependency_installer.py, 13 tests (RAN: all pass). Dialog flow not tested beyond the Tk-unavailable fallback.
- History: added a5ad487 (26-07-2026), unchanged.

### utils/requirements_check.py (133 lines) - [UTIL]

- Purpose: at start-up, checks whether openpyxl and pywin32 are present for the running Python, classifies each as essential or
  optional, and builds the console report. (READ)
- Imports (READ): `importlib.util` (12), `sys` (13), `typing` List, NamedTuple (14). No tkinter, no third-party.
  RAN: with openpyxl blocked, importing this module (and its package `utils`, whose __init__ imports only utils/constants.py and
  utils/helpers.py, standard library only) worked and reported openpyxl missing; no third-party module was loaded. (HIGH)
- Module-level: `ON_WINDOWS = sys.platform == 'win32'` (17).
- `class Requirement(NamedTuple)` (20-25): module (import name), package (pip name), purpose, essential, consequence=''.
- `class Result(NamedTuple)` (28-34): requirement, installed; property `essential` -> requirement.essential.
- `REQUIREMENTS` (37-61):
  - openpyxl / openpyxl, "Reading and writing Excel files", essential=True.
  - module 'pythoncom' / package 'pywin32', purpose "Excel automation: preserves Power Query when creating reports, and powers the
    refresh in Step 4", essential = ON_WINDOWS (True on Windows, False elsewhere), consequence "Reports would be created WITHOUT
    their Power Query connections, and Step 4 could not refresh them."
  - lxml is NOT checked (it is optional in requirements.txt).
- `_is_installed(module_name)` (64-70): True if already imported; else `importlib.util.find_spec(name) is not None`
  (finds the module WITHOUT importing it); ImportError/ValueError -> False.
  - GAP (RAN, MEDIUM-HIGH): "found" is not "importable". RAN: a stand-in `pythoncom.py` that raises ImportError on import was
    reported installed=True. A broken pywin32 (e.g. DLL not loadable) would pass this check; then core/excel_handler.py:21-27
    silently sets USE_WIN32COM=False and reports are built with openpyxl - the very outcome be6418c meant to block
    (comments 49-53). Also: whether `pythoncom` is findable depends on pywin32's .pth start-up file being processed; under an
    embeddable Python (Standard 13's `_runtime`) that depends on its `._pth` settings - not checked here (LOW).
- `check_requirements(requirements=REQUIREMENTS)` (73-74): list of Results. Called main.py:71.
- `missing_essential(results)` (77-79) / `missing_optional(results)` (82-84): filters. Called main.py:75-76 and inside format_report.
- `interpreter_hint()` (87-90): "Python X.Y.Z at <sys.executable>". Called main.py:121 and format_report:128.
- `format_report(results)` (93-133): "Checking requirements..." then one line per requirement
  `"   [<mark>] <package padded to 10> - <purpose>"` with mark 'OK     ', 'MISSING' (essential) or 'not found' (optional; 9 chars,
  so columns misalign - RAN); then for each optional missing: "", "⚠️  <pkg> is not installed for this Python.", "   Without it:
  <purpose>.", consequence line, "   Install with:  pip install <pkg>"; for each essential missing: "", "❌ <pkg> is REQUIRED and is
  not installed.", "   Needed for: <purpose>.", consequence + "   The app stops here rather than produce damaged output." (only when a
  consequence exists), "   Install with:  pip install <pkg>"; and when anything is missing: "", "   Running: <hint>", "   Note 'py'
  and 'python' can be different interpreters - install", "   for the one you start the app with." Called main.py:72.
  - Wording: "Without it: Excel automation: preserves Power Query..." reads as if that is what you get WITHOUT it (UX, LOW).
  - The report suggests `pip install <pkg>` while the offer installs `-r requirements.txt` (inconsistent advice, LOW).
- Business rules (not GST): openpyxl always essential; pywin32 essential only on Windows (57). Rationale in comments 49-56
  (damaged reports worse than refusing to start).
- Excel/files: none. Threads: none.
- Tests: tests/test_requirements_check.py, 15 tests (RAN: all pass), including "does not import the module to check it".
- History: f363946 created (pywin32 optional), be6418c made pywin32 essential on Windows and added `consequence`.

---

## Cross-file observations

### 1. Shared files, caches, settings and logs - write path vs read path

| Item | Written by | Read by | Consistent? |
|---|---|---|---|
| Settings file `Path.home() / '.gst_organizer_cache.json'` (Windows: `%USERPROFILE%\.gst_organizer_cache.json`), one-line JSON | gui/utils/cache_manager.py:45-46 (dict built at gui/main_window.py:360-370; triggers listed under cache_manager) | gui/utils/cache_manager.py:21-23 via main_window.py:121, used at main_window.py:375-389; power_query_extractor/gui/extractor_window.py:118, 590-593 (key `target_folder`); power_query_extractor/extractor_main.py:41, 50-58 (keys `target_folder`, `target_path`, `output_folder`, `last_target`) | Path identical in all three (READ). Key `target_folder` agrees. Encoding: all open with the default encoding; the file is pure ASCII (json escapes), so no mismatch (RAN). `recent_folders` written, never read. |
| Other extractor cache candidates `gst_organizer_cache.json`, `.cache/app_state.json`, `config/last_session.json` (relative to the CURRENT folder) | NOBODY in A (grep) | power_query_extractor/extractor_main.py:42-44 | Dead read paths (A-R2's file). |
| Main log `gst_organizer.log` in the CURRENT folder | main.py:24 (root logger, so every module's logger ends here when started via main.py) | humans only; README.md:290, 333 mention it (docs readers) | Location depends on how the app was started. |
| Extractor log `pq_extractor.log` in the current folder | power_query_extractor/extractor_main.py:16 (only when that module is imported - extract_tab.py:8-10 deliberately avoids it) | humans | - |
| `<app folder>\logs\`, `<app folder>\temp\` | created by main.py:35-37 | nobody | Dead folders. |
| Excel scratch folder `%TEMP%\gst_excel_temp` | core/excel_handler.py:53 (A-R1) | A-R1 | Not my file; noted because main.py's `temp\` is unrelated to it. |

### 2. Standard 13 comparison (facts only; decisions are the author's)

| Standard 13 says | A does | Evidence |
|---|---|---|
| No console window; pythonw.exe; launcher .bat | Requires a console: crashes silently under pythonw; prints status and waits for Enter | main.py:17 (RAN), 62-73, 45-54 |
| Settings in `%LOCALAPPDATA%\<App Name>\` | `%USERPROFILE%\.gst_organizer_cache.json` | cache_manager.py:15 |
| Logs in `%LOCALAPPDATA%\<App Name>\` (error files) | `gst_organizer.log` in the current folder; no per-crash error file | main.py:24 (RAN) |
| Uncaught errors -> plain-English message box + log, via sys.excepthook and threading.excepthook | Start-up errors -> console text + log; no excepthooks; errors in clicks/threads go to the console only | main.py:114-134 (READ) |
| Nothing installed on users' PCs | Offers to pip-install requirements.txt (incl. pytest) into the running Python from the internet | dependency_installer.py:26-60, 88-158 (RAN command) |
| Users' PCs never write into the shared app folder (`-B`) | Creates `logs\`, `temp\` (and possibly `__init__.py`) in the app folder; may write the log there; .pyc files unless -B | main.py:29-43 (RAN) |
| Find own files from APP_DIR, not the current folder | requirements.txt and sys.path from `Path(__file__).parent` (good); log file and the ImportError advice use the current folder | main.py:31-32, 84 vs 24, 124 |
| requirements.txt: exact versions, no pytest | `>=` minimums, includes pytest | requirements.txt:5-15 |
| pytest.ini with `pythonpath = .` and `testpaths = tests` | `testpaths = tests`, no `pythonpath` (conftest adds the path) | pytest.ini, tests/conftest.py:9 |
| (Standard 12) desktop windows: PySide6 | Tkinter | all gui/* |

### 3. Inconsistent patterns

- Restart drops interpreter options (dependency_installer.py:70) - relevant if the app is ever started with `-B -s`.
- Template types: setup_validation.py:21 accepts .xlsx/.xltx/.xltm/.xlsm; Browse dialogs offer only .xltx/.xlsx
  (file_handler.py:31, 42).
- Error reporting styles in my files: cache_manager logs ERROR and swallows; dark_mode_manager logs DEBUG and swallows; tooltip
  swallows silently; dependency_installer logs ERROR or shows message boxes; main.py prints to the console. None shows a message
  box for start-up failures (except the install flow).
- Version strings all say 3.0 although A has 32 commits of new features in July 2026: main.py:2, 6, 63; status_bar.py:29
  ("v3.0 | Production Ready"); requirements.txt:1; utils/constants.py:134 window title 'GST File Organizer v3.0';
  gui/__init__.py:11 and utils/__init__.py:38 `__version__ = '3.0.0'`; main_window.py:393. (B calls itself v4.0.) (READ)
- Re-export `__init__.py` files exist (gui/utils, gui/widgets) but nobody imports from them, and new modules
  (setup_validation, tooltip) were not added - the established practice is "import the submodule directly". (HIGH)
- Thread use of Tk: processing_handler calls `app.log_message` and `app.update_progress` directly from the worker thread
  (processing_handler.py:319-523, 459, 482) and also through `_log_async`/`_progress_async` (23-36) - mixed. UIHelpers.log_message
  and main_window.update_progress then call `after()` from that thread (relies on Tk forwarding; RAN OK on Linux). (MEDIUM)
- Settings saving is event-driven and partial (cross-file, READ): processing-mode radio buttons have no `command`
  (setup_tab.py:404-407) so a change is saved only when some other save happens; paths typed into the entries are not saved until
  a later save event; nothing saves on closing (main_window.py:344-355). A non-numeric value typed into the length spinbox makes
  `IntVar.get()` raise TclError inside main_window.save_cache (main_window.py:369), so every save fails while it stays invalid (MEDIUM).

### 4. Dead code found in my files

- gui/utils/cache_manager.py:50-52 `get_cached_value`, :54-56 `set_cached_value` (no callers).
- gui/utils/ui_helpers.py:11-14 `UIHelpers.update_status` (no callers).
- gui/utils/dark_mode_manager.py:70, 83 `_original_root_bg` (set, never read); `is_dark_mode` attribute read only on the
  unused default path of update_tree_tags (279).
- gui/utils/setup_validation.py:36-44 `_default_exists` (only tests in practice); unused imports `os` (8), `Dict` (10).
- main.py:11 unused `import os`; utils/dependency_installer.py:18 unused `Path`.
- Unused parameters in practice: CacheManager `cache_file`, offer_install `parent`, Tooltip `delay`/`wraplength`, `_short` `width`.
- Package re-exports in gui/utils/__init__.py:9 and gui/widgets/__init__.py:7 (never used).
- main.py:35-37 creates `logs`/`temp` that nothing uses.
- Outside my files but found on the way: gui/main_window.py:25 imports CollapsibleFrame and never uses it.

### 5. GST rules

- None found in any of my 17 files (READ). The only domain-flavoured items are the labels "ITC template" / "Sales template" and
  the requirement that both are present before processing (setup_validation.py:82-83, 94-97) - treated as workflow rules.

### 6. Things that could only be traced by reading (Windows-only)

- pythonw behaviour (simulated by setting stdout to None), Windows console buffering, cp1252 default log encoding (simulated with
  an explicit cp1252 handler), "Not Responding" during pip, Tk version and threading build of the Windows python.org Python,
  console behaviour after Popen restart, Windows font rendering (Segoe UI/Consolas/emoji), exact pixel position of the Dark Mode
  checkbox on Windows.

## Possible client names

- None found in my 17 assigned files, nor in the commit messages/diffs I read (edac6a3, 5ea77c2, 1cb878b, 6c1e1c6, 064c634,
  bd176bf, f363946, be6418c, a5ad487, aeb83cf, 725fbaf, c89c1d4, 169c7b5). The tests I glanced at use the generic folder
  "ABC-DL" (tests/test_extractor_panel.py:27, tests/test_extractor_standalone.py:63, tests/test_main_window_tabs.py:88) -
  looks generic, listed only so the test reader can confirm.

## Questions for the author

1. Where should settings and logs live in the merged app - `%LOCALAPPDATA%\<App Name>\` (Standard 13), A's
   `%USERPROFILE%\.gst_organizer_cache.json` + log in the current folder, or B's `~/.gst_organizer/`?
2. A needs a console window and offers to install packages with pip. Standard 13 says no console and nothing installed on PCs.
   Should the install offer be dropped in the merged app?
3. Should `recent_folders` (last 5 source folders, saved but never shown) become a visible feature or be removed?
4. Should macro-enabled templates (.xlsm/.xltm) be accepted? The Setup checklist accepts them; the Browse dialog does not.
5. Is dark mode wanted in the merged app? (A spent four commits fixing its restore; known remaining issues: stale colours on
   widgets that change colour in light mode, widgets created while dark stay light, client rows become white-on-pale after a scan
   in dark mode, selected-row colours change after the first toggle.)
6. Should the version text ("v3.0 | Production Ready") and "Author: Advanced AI Assistant" be changed?
7. Is it intentional that `.claude/settings.local.json` is committed? (Content harmless.)
8. Were the `logs\` and `temp\` folders meant to be used for something?

## INCOMPLETE

- Nothing missing: all 17 assigned files read line by line; every function's callers searched across all .py files of the snapshot.
  Windows-only behaviours are listed in section 6 above as traced by reading/simulation only.
