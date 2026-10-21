## SECTION 3: FEATURE INVENTORY

This section lists every feature of project A (`goyal-sidhant/AnnualStatement_oldv1`, branch `main`, snapshot `bf2f1dd` of 27-07-2026, the Tkinter program that calls itself "GST File Organizer v3.0") and, for each one, says what it does, where it lives, what goes in and what comes out, how the user starts it, and what really happens when one realistic example is pushed through it. **This file is the first part: features A-F1 to A-F15** (start-up, Step 1 Setup, Step 2 Validation and Step 3 Processing up to the two reports). Features A-F16 to A-F31 (summary workbook, progress and log, Step 4 extraction, settings, dark mode, logging, help texts) follow in the second part.

The feature IDs are those of the project registry (`/root/merge-scratch/audit/A_registry.md`); rule IDs (A-R1 ...) refer to the same registry and are explained in full in Section 5. Gaps found here are listed for Section 7 (they will get A-G numbers there).

### How to read this section

- **READ** = established by reading the code. **RAN** = the code was executed on made-up inputs in a scratch copy of the project (`/root/merge-scratch/run/A-S3a/proj`, identical to the snapshot; checksums of the original copy were taken before and after all runs and are identical). "RAN (me)" = run for this section; scripts are in `/root/merge-scratch/run/A-S3a/scripts/`, the long outputs in `/root/merge-scratch/run/A-S3a/out/`. "RAN (TRACE)", "RAN (CORE)", "RAN (GUI)", "RAN (START)", "RAN (XL)" = run earlier by the reader whose working notes carry that short name (list in the registry); I re-checked each such claim against the code.
- **Confidence** of every inference: **HIGH / MEDIUM / LOW**. Where I am not sure I say so and give a way to check.
- **[GST RULE]** marks anything that decides which return, table, sheet or cell a file or figure belongs to, a state code, a period, and the like. For these I only say what the code does; I never say whether it is right. The symbol **⚠** marks values that can change over time (return names, file-naming conventions, template cell positions, state lists) so the author can check them.
- **Excel through COM** (Windows' way for one program to drive another; here Python drives Microsoft Excel through the **pywin32** package) cannot run on this Linux machine. Those parts are traced **by reading**. For report creation I additionally drove the Python code with **fake COM objects** (stand-ins that only record which calls are made); that proves the order of the Python calls, not what real Excel would do.
- Window tests ran under **Xvfb** (a virtual screen, so Tk windows can be built without a monitor). Message boxes, file dialogs and `os.system` were replaced by recorders so the runs need no clicks.
- All client names are invented. In the runs, **`<Client A>` = the made-up name "ClientA Traders", state Maharashtra** and **`<Client B>` = the made-up name "ClientB Exports Private Limited", state Gujarat**. They appear inside made-up file names because the code reads the client from the file name; they are not real clients.

### Plain-English terms used in this section (first use)

| Term | Meaning |
|---|---|
| Tk / Tkinter | Python's built-in toolkit for windows, buttons and boxes. **widget** = any one on-screen element. |
| ttk | Tk's "themed" widgets, whose colours come from a shared style table. |
| Treeview | Tk's table widget; it shows the Step 2 client list. |
| Tk variable (`StringVar`, `IntVar`, `BooleanVar`) | a value holder that a widget shows and edits; an `IntVar` must hold a whole number. |
| trace (on a Tk variable) | an automatic call made every time the variable is written. |
| debounce | wait until typing pauses before re-checking. |
| thread / worker thread | a second line of work running beside the window, so the window does not freeze. **daemon thread** = a worker that is killed instantly, without clean-up, when the program exits. |
| main loop | Tk's endless "wait for a click, handle it" loop; a worker thread can only talk to Tk while it runs. |
| regex / pattern | a text-matching rule used to recognise file names; `[^-]+` means "one or more characters that are not a hyphen". |
| openpyxl | a Python library that reads and writes `.xlsx` files without Excel. When it saves a workbook it drops Power Query. |
| Power Query | Excel's data-import tool; its queries live in a hidden part of the workbook (DataMashup); "refresh" re-runs them. |
| Links sheet | the sheet in the ITC/Sales templates into which the program writes folder paths and file names that the template's queries read. |
| staging file | a temporary local copy that is worked on and then copied to its final place. |
| stem | a file name without its extension (`Sales-X.xlsx` → `Sales-X`). |
| mtime | a file's or folder's "date modified". |
| JSON / settings cache | a plain-text settings format; the program keeps its remembered settings in `~/.gst_organizer_cache.json` (the user's home folder). |
| CWD (current folder) | the folder a program was started from; a file name with no folder part is created there. |
| `find_spec` | a Python check that a module file exists, without loading it. |
| pip | Python's package installer (downloads packages from the internet). |
| UNC path | a network path `\\server\share\...`. |

### The made-up test bed used by the traces (RAN, me)

Built by `scripts/common.py` with openpyxl (every workbook is a small genuine `.xlsx`, about 5 KB):

| Item | Contents |
|---|---|
| Source folder, `<Client A>` (complete) | `GSTR-2B-Reco-ClientA Traders-Maharashtra-Apr25-Mar26.xlsx`, `ImsReco-ClientA Traders-Maharashtra-31032026.xlsx`, `GSTR3B-ClientA Traders-Maharashtra-Apr.xlsx`, `GSTR3B-ClientA Traders-Maharashtra-May.xlsx`, `Sales-ClientA Traders-Maharashtra-Apr-Mar.xlsx`, `SalesReco-ClientA Traders-Maharashtra-FY 2025-26.xlsx`, `AnnualReport-ClientA Traders-Maharashtra-2025-26.xlsx` |
| Source folder, `<Client B>` (only GSTR-2B and GSTR-3B, no Sales file) | `GSTR-2B-Reco-ClientB Exports Private Limited-Gujarat-Q1.xlsx`, `GSTR3B-ClientB Exports Private Limited-Gujarat-Apr.xlsx` |
| Source folder, names no pattern accepts | `ClientB notes.xlsx`, `Sales Reco-ClientA Traders-Maharashtra-Q1.xlsx` (space instead of `SalesReco`) |
| Source folder, files that should never reach the parser | `tiny.xlsx` (500 bytes), `GSTR3B-ClientA Traders-Maharashtra-Jun.csv`, sub-folder `old/` holding `GSTR3B-ClientA Traders-Maharashtra-Mar.xlsx` |
| Templates | `ITC_Template.xlsx` and `Sales_Template.xlsx`, each with a sheet `Links` whose cells B2 to B11 hold the placeholder text `OLD-B2` ... `OLD-B11`, plus a sheet `Info`. No Power Query (openpyxl cannot create one). |

The main end-to-end run (`scripts/t_pipeline.py`, output `out/pipeline.txt`) builds the **real** main window under Xvfb and goes: Browse the four paths → Scan Files → Step 2 table, Complete Only, a per-client folder-name click → Dry Run → Export List → Start (run 1: Fresh, organise-only, folder stamp `211026 1030`) → Start (run 2: Resume, full pipeline, folder stamp `211026 1045`) → Start with the length box emptied. The two folder stamps were fixed by replacing `get_timestamp` in `core/file_organizer.py` and `gui/handlers/processing_handler.py` so a "later" run could be simulated without waiting. Because this is Linux without pywin32, the reports in this run were built by the **openpyxl fallback** (`core/excel_handler.py:262-322`), not by Excel.

### All features of project A (registry list)

Status words: "Works with defects" = the feature does its main job, but the traces or the code show defects (listed in the feature's subsection or in the second part). No feature of A is missing or wholly broken at the snapshot; every one of the 31 carries defects.

| ID | Name | Status | Where (main files and lines at `bf2f1dd`) |
|---|---|---|---|
| A-F1 | Start-up and dependency checks (incl. offer to install missing packages) | Works with defects | `main.py:17-134`; `utils/requirements_check.py:17-133`; `utils/dependency_installer.py:26-158`; `requirements.txt:1-19` |
| A-F2 | Main window with four step tabs, title strip and status bar (incl. close confirmation) | Works with defects | `gui/main_window.py:41-398`; `gui/widgets/title_bar.py:8-51`; `gui/utils/status_bar.py:8-37`; `utils/constants.py:132-157` |
| A-F3 | Choose source folder, target folder and templates (Setup) | Works with defects | `gui/tabs/setup_tab.py:353-391`; `gui/handlers/file_handler.py:19-55`; `gui/main_window.py:80-84, 210-220` |
| A-F4 | Live setup checklist | Works with defects | `gui/tabs/setup_tab.py:282-314, 506-608`; `gui/utils/setup_validation.py:13-115` |
| A-F5 | Workflow mode (organise-only vs full pipeline) | Works with defects | `utils/constants.py:163-172`; `gui/tabs/setup_tab.py:316-351`; `gui/handlers/processing_handler.py:136-175, 519` |
| A-F6 | Maximum client folder name length setting | Works with defects | `gui/tabs/setup_tab.py:429-451, 610-615`; `core/file_organizer.py:42, 83-87`; `utils/helpers.py:620-644` |
| A-F7 | Scan source folder and recognise client files (Scan / Re-scan) | Works with defects | `gui/handlers/file_handler.py:57-130`; `core/file_parser.py:56-306`; `utils/helpers.py:263-365`; `utils/constants.py:13-75` |
| A-F8 | Client list with completeness status (Validation) | Works with defects | `gui/tabs/validation_tab.py:18-247`; `gui/handlers/client_handler.py:18-173`; `gui/main_window.py:294-328` |
| A-F9 | Unrecognised files / variations list | Works with defects | `core/file_parser.py:130-132, 209-220, 288-306`; `gui/main_window.py:320-321`; `core/excel_handler.py:554-585` |
| A-F10 | Per-client option to include the client name in folder names (with global override) | Works with defects | `gui/tabs/setup_tab.py:417-427`; `gui/handlers/client_handler.py:49, 76-79, 175-223`; `core/file_organizer.py:90, 117-139` |
| A-F11 | Dry run (preview) | Works with defects | `gui/handlers/processing_handler.py:38-67`; `gui/tabs/validation_tab.py:109-118` |
| A-F12 | Export selected client list to a text file | Works with defects | `gui/handlers/client_handler.py:225-257`; `gui/tabs/validation_tab.py:121-130` |
| A-F13 | Organise files into client and category folders (fresh / re-run / resume, backups) | Works with defects | `core/file_organizer.py:26-379`; `utils/helpers.py:68-234, 620-644`; `gui/handlers/processing_handler.py:189-284` |
| A-F14 | Create ITC report from template | Works with defects | `gui/handlers/processing_handler.py:291-326`; `core/excel_handler.py:66-369, 603-656`; `utils/excel_com.py:19-29` |
| A-F15 | Create Sales report from template | Works with defects | `gui/handlers/processing_handler.py:328-411`; `core/excel_handler.py:66-369, 657-708` |
| A-F16 | Processing summary workbook and organisation report | Works with defects | `core/file_organizer.py:317-379`; `core/excel_handler.py:371-601`; `gui/handlers/processing_handler.py:413-475` |
| A-F17 | Processing progress and log (Start / Stop, ETA, completion) | Works with defects | `gui/tabs/processing_tab.py:17-117`; `gui/handlers/processing_handler.py:23-36, 69-134, 177-530`; `gui/utils/ui_helpers.py:16-40` |
| A-F18 | Step 4 tab hosting the extractor (folder follows Step 1; fallback notice) | Works with defects | `gui/tabs/extract_tab.py:1-131`; `power_query_extractor/gui/extractor_window.py:98-139` |
| A-F19 | Find each client's reports for Step 4 (latest run folder, version folder) | Works with defects | `power_query_extractor/gui/extractor_window.py:20-42, 587-668` |
| A-F20 | Step 4 client and report selection with last-refresh status | Works with defects | `power_query_extractor/gui/extractor_window.py:45-82, 343-432, 670-836` |
| A-F21 | Refresh reports through Power Query | Works with defects | `power_query_extractor/core/report_processor.py:62-415, 638-799` |
| A-F22 | Skip refresh and refreshed-file naming | Works with defects | `power_query_extractor/core/refresh_naming.py:14-66`; `power_query_extractor/core/report_processor.py:51, 127-159` |
| A-F23 | Extract cells from refreshed reports | Works with defects | `power_query_extractor/core/report_processor.py:527-636`; `power_query_extractor/config/cell_mappings.py:33-278` |
| A-F24 | Consolidated extraction workbook | Works with defects | `power_query_extractor/core/data_consolidator.py:20-248, 416-432` |
| A-F25 | Extraction log, statistics and run history | Works with defects | `power_query_extractor/gui/extractor_window.py:494-571, 885-1029` |
| A-F26 | Automatic closing of Excel error dialogs | Works with defects | `power_query_extractor/core/report_processor.py:305-342, 417-525` |
| A-F27 | Settings cache | Works with defects | `gui/utils/cache_manager.py:11-56`; `gui/main_window.py:358-389`; `power_query_extractor/extractor_main.py:36-63` |
| A-F28 | Dark mode / theme | Works with defects | `gui/utils/dark_mode_manager.py:16-294`; `gui/main_window.py:157-207` |
| A-F29 | Standalone extractor launcher | Works with defects | `launch_extractor.py:1-30`; `power_query_extractor/extractor_main.py:1-82` |
| A-F30 | Application log file | Works with defects | `main.py:20-27`; `power_query_extractor/extractor_main.py:12-19` |
| A-F31 | Help texts, tooltips, scrolling and collapsible sections | Works with defects | `gui/widgets/tooltip.py:5-69`; `gui/tabs/setup_tab.py:71-138, 193-248, 477-503`; `gui/widgets/collapsible_frame.py:7-41` |

### A-F1: Start-up and dependency checks (incl. offer to install missing packages)

#### What it does

When the user runs `python main.py`, the program first re-wraps the console output as UTF-8 (so emoji print), sets up the log file, makes sure its own folder is on Python's search path, creates a few folders, prints a banner and a short report of which required packages are installed, and — if anything is missing — offers in a Yes/No box to install the missing packages with pip and restart itself. If an essential package is still missing it stops with exit code 1; otherwise it opens the main window (A-F2).

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Console re-wrap to UTF-8 (runs the moment `main.py` is loaded) | `main.py:17` |
| Log set-up: `logging.basicConfig(level=INFO, ...)` with `FileHandler('gst_organizer.log')` + console | `main.py:20-27` (see A-F30) |
| `setup_environment()`: program folder first on the search path; creates `core`, `utils`, `gui`, `logs`, `temp` in the program folder; writes a placeholder `__init__.py` into core/utils/gui if missing | `main.py:29-43` |
| `_pause()`: waits for Enter, ignores end-of-input / Ctrl+C / OS errors | `main.py:45-54` |
| `main()`: banner, check, report, offer, block or warn, open the window, two error handlers | `main.py:57-134` |
| Which packages are checked: `REQUIREMENTS` (openpyxl essential; module `pythoncom` = package pywin32, essential only when `sys.platform == 'win32'`) | `utils/requirements_check.py:17, 37-61` (essential flag `:57`) |
| Presence test without importing (`find_spec`) | `utils/requirements_check.py:64-70` |
| Split into essential / optional missing; interpreter hint; console report | `utils/requirements_check.py:73-133` |
| pip command, install with 600 s time limit, restart, offer dialog | `utils/dependency_installer.py:23, 26-38, 41-60, 63-74, 88-158` |
| Package list the installer uses | `requirements.txt:1-19` (`openpyxl>=3.1.2`, `pywin32>=305; sys_platform == 'win32'`, `lxml>=4.9.0`, `pytest>=7.0.0`) |

#### Input → Output

The Python that runs the program and the packages installed in it → console banner and report; possibly a Yes/No box "Install missing packages?"; possibly a pip run and a restart; exit code 1 when an essential package is missing; otherwise the main window. Side effects on the PC: `logs\` and `temp\` folders (and possibly `__init__.py` files) in the program folder; `gst_organizer.log` in the current folder; with the user's consent, packages downloaded from the internet into that Python.

#### User interaction

Automatic when the user starts `python main.py` (or double-clicks it, which opens a console). The only choices are "Yes/No" in the install box and pressing Enter at "Press Enter to exit...". There is no launcher file and no `pythonw` (no-console) start in A (READ; START notes).

#### Concrete trace

Script `scripts/t_f1_startup.py` (RAN, me). The main window was replaced by a recorder so no window opened; `offer_install` was replaced by a recorder that answers "No" in scenarios 1-2; in scenario 3 only the Yes/No box, `subprocess.run` and `subprocess.Popen` were replaced, so the real offer code ran. Each run started from an empty work folder (the current folder).

**Scenario 1 — a Linux PC with openpyxl but without pywin32; the user answers "No" (RAN).**

1. `main.py:17` re-wraps `sys.stdout`; `main.py:20-27` creates `gst_organizer.log` **in the current folder** (RAN: it appeared in `work/f1_linux_decline`, not in the program folder).
2. `main()` → `setup_environment()` (`main.py:29-43`) puts the program folder first on the search path and creates `logs/` and `temp/` in the program folder (RAN: both appeared in the scratch program folder; nothing in A ever writes into them — READ, grep).
3. Banner (`main.py:62-64`).
4. `check_requirements()` (`utils/requirements_check.py:73-74`) calls `_is_installed` (`:64-70`) for each entry: `openpyxl` → found; `pythoncom` → `find_spec` returns nothing → not installed. `ON_WINDOWS` is False (`:17`), so pywin32 is **optional** here (`:57`).
5. `format_report()` (`:93-133`) printed exactly:
   ```
   Checking requirements...
      [OK     ] openpyxl   - Reading and writing Excel files
      [not found] pywin32    - Excel automation: preserves Power Query when creating reports, and powers the refresh in Step 4

   ⚠️  pywin32 is not installed for this Python.
      Without it: Excel automation: preserves Power Query when creating reports, and powers the refresh in Step 4.
      Reports would be created WITHOUT their Power Query connections, and Step 4 could not refresh them.
      Install with:  pip install pywin32
   ```
   (the "[not found]" mark is 9 characters against 7, so the columns do not line up — cosmetic).
6. `blocking = []`, `optional = [pywin32]` (`main.py:75-76`) → `offer_install(...)` is called (`:78-88`) with `requirements_path = <program folder>/requirements.txt` because that file exists (`:84, 87`). Recorded call: `[('pywin32', essential=False, installed=False)]`, path `.../proj/requirements.txt`. The answer "No" returns False.
7. Nothing is blocking (`:97`), so `:104-106` logs `WARNING - pywin32 not installed - Excel automation: preserves Power Query ...` (RAN: this line is in the log file).
8. `:109-112` builds and runs the main window (recorded).

**Scenario 2 — the same PC pretending to be Windows (`sys.platform = 'win32'` set before `utils/requirements_check.py` was loaded) (RAN, simulated).**

Point 4 of scenario 1 now marks pywin32 **essential** (`utils/requirements_check.py:57`); the report prints `[MISSING] pywin32 ...`, then `❌ pywin32 is REQUIRED and is not installed.` and `The app stops here rather than produce damaged output.` (`:114-123`). The offer is still made (`main.py:78`); after "No", `main.py:97-102` logs `ERROR - Required package missing: pywin32`, prints `Press Enter to exit...`, `input()` meets end-of-input and is ignored (`:51-54`), and the program ends with **exit code 1**. The window is never built.

**Scenario 3 — Linux, the user answers "Yes" (RAN, Xvfb).**

The real `offer_install` (`utils/dependency_installer.py:88-158`) showed this Yes/No text (recorded verbatim):
```
The app needs these packages:

• pywin32 - Excel automation: preserves Power Query when creating reports, and powers the refresh in Step 4
    Reports would be created WITHOUT their Power Query connections, and Step 4 could not refresh them.

Install them now for:
<path of the running python>

The app will restart afterwards.
```
After "Yes" it opened the "Installing…" window (`:121-125`) and called `install(packages=None, requirements_path=...)` (`:133-134`), which ran (recorded):
```
[<python>, '-m', 'pip', 'install', '-r', '<program folder>/requirements.txt']   capture_output=True, text=True, timeout=600
```
My stand-in reported success, so `:136-139` returned True; `main.py:90` printed `✅ Packages installed - restarting…`; `restart_application()` (`utils/dependency_installer.py:63-74`) called `Popen([sys.executable] + sys.argv)` and `main.py:92` ended with **exit code 0**.

**Scenario 4 — no console at all (what `pythonw.exe` gives: `sys.stdout` is None) (RAN).**

`main.py:17` raised `AttributeError: 'NoneType' object has no attribute 'buffer'` before line 20, so **no log file was created**, no window and no message appeared (the work folder stayed empty).

#### What the trace shows against what the feature claims

> **BUG (for Section 7): without a console the program dies silently.** `main.py:17` assumes a console; under a no-console start nothing is shown and nothing is logged (RAN). HIGH.

> **BUG (for Section 7): the box names one package, the command installs the whole `requirements.txt`.** The box lists only pywin32, but the command is `pip install -r requirements.txt` (RAN), which also installs or upgrades openpyxl, lxml and **pytest** (a test tool) on the user's PC (`requirements.txt:5-15`). Installing the whole file is deliberate according to the commit message of `a5ad487` ("Installs from requirements.txt when present, so a fresh machine is set up in one go"), but the user is not told. The manual command shown after a failure combines `-r requirements.txt` **and** the package names, with paths not quoted (`utils/dependency_installer.py:142`; READ), so a path with spaces breaks when pasted. HIGH.

- **Off Windows the offer comes back at every start** (READ + RAN START): pip skips pywin32 because of its marker `sys_platform == 'win32'` (`requirements.txt:9`), reports success, the program restarts, pywin32 is still missing, and the offer appears again. HIGH.
- **"Found" is not "importable"** (RAN XL E2, simulated): `find_spec` (`utils/requirements_check.py:68`) only checks that a `pythoncom` file exists. A pywin32 that is present but cannot load passes the check on Windows, and report creation then falls back silently to openpyxl, losing Power Query (`core/excel_handler.py:21-28, 96-102`; see A-F14). The check also never looks at `win32com`, `win32gui`, `win32api` or `lxml` (`:37-61`), nor whether Excel itself is installed. MEDIUM likelihood, HIGH for the mechanism.
- **The banner may appear late** (RAN START in a pseudo-console): the UTF-8 wrapper made at `main.py:17` is not line-buffered, so text printed before the Yes/No box can stay invisible until a later flush. MEDIUM for the Windows console.
- **Exit code 0 although the app did not start** when the automatic restart fails (`main.py:93-95`; READ).
- The restart drops interpreter options such as `-B` because `sys.argv` does not contain them (`utils/dependency_installer.py:70`; READ, HIGH).
- pip runs on the window's own thread (`:47-50` called from `:133`), so the "Installing…" window cannot repaint for up to 600 s; its output appears only after pip ends (`:57-59`; READ). The commit message of `a5ad487` says the window "streams pip's output" — the code shows it only at the end.
- The `ImportError` advice `pip install -r requirements.txt` (`main.py:124`) uses a relative path, which works only when the console is in the program folder (READ).
- `logs/` and `temp/` are created at every start (`main.py:35-37`) and never used (READ, grep). A read-only program folder makes `mkdir` fail → "💥 Critical error" and exit 1 (`main.py:130-134`; READ, HIGH).
- Errors in button clicks and in worker threads after the window opens are **not** caught by `main()` (there is no `sys.excepthook` / `threading.excepthook`; READ): they go to the console only.

Rules involved: A-R124 (which packages, essential vs optional), A-R125 (offer to install), A-R126 (requirements.txt ⚠ package versions), A-R127 (a console is assumed), A-R128 (program folder preparation), A-R120 (log file). None is a [GST RULE].

#### Older documents against the code

- `README.md:110-113` (RM15) still says only "python main.py"; the start-up check and install offer are not documented. `README.md:104-108` (RM14) shows `openpyxl>=3.0.0` and no lxml/pytest — not the real file.
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:20, 726` (S1.11, DEP2) say requirements.txt lacks pywin32 — no longer true since `064c634` (23-07-2026).
- Comment `requirements.txt:8` says "power_query_extractor and excel_handler import pythoncom/win32com/win32gui"; `core/excel_handler.py:23-24` imports only `win32com.client` and `pythoncom` (READ) — partly inaccurate.
- `main.py:1-7` docstring "Author: Advanced AI Assistant", "Version: 3.0 (Production Ready)" (version label unchanged through 34 later commits; A-R156).

#### History

The console re-wrap, the log set-up and `setup_environment()` are from `edac6a3` (31-08-2025, shared - before the split on 30-10-2025); `5ea77c2` removed a malformed shebang line (shared - before the split on 30-10-2025). The requirement check (`f363946`), "pywin32 essential on Windows" plus `_pause()` (`be6418c`) and the install offer (`a5ad487`) were all added on 26-07-2026 between 23:15 and 23:27 (A only (after the split)). The pywin32 line in `requirements.txt` came with `064c634` and the un-commented pytest line with `6c1e1c6` (both 23-07-2026, A only (after the split)).

Fact for later sections (no recommendation): the author's Standard 13 (`rebuild_pipeline_FINAL.md:900-1061`) describes a start without a console window and says nothing is installed on users' PCs; A requires a console (scenario 4) and offers to pip-install into the running Python (scenario 3).

### A-F2: Main window with four step tabs, title strip and status bar (incl. close confirmation)

#### What it does

Builds the program's one window: a blue title strip with the title, a subtitle and a "🌙 Dark Mode" tick box; four tabs "Step 1" to "Step 4"; a dark status bar at the bottom. It owns all shared state (the Tk variables for the paths and options, the scan results, the "processing" flags), creates the three handlers that do the work, restores the remembered settings, and asks for confirmation before closing — but only while Step 3 is running.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Build order: Tk root → window → variables → components → handlers → GUI → events → cached settings | `gui/main_window.py:46-55` |
| Title "GST File Organizer v3.0", size 1200×800, minimum 1000×700, background `#F8F9FA`, centring, ttk theme 'clam' | `gui/main_window.py:57-77`; values `utils/constants.py:133-137, 151` |
| Tk variables and shared data (`processing_mode='fresh'`, `workflow_mode='organize'`, `client_name_max_length=35`, `client_data`, `is_processing`, …) | `gui/main_window.py:79-110` |
| `FileParser()`, `ExcelHandler()` (creates `%TEMP%\gst_excel_temp` at start-up), settings cache load, dark-mode manager | `gui/main_window.py:112-123` |
| Handlers `FileHandler`, `ClientHandler`, `ProcessingHandler` | `gui/main_window.py:125-129` |
| Title strip, notebook (the tab container), four tabs, status bar | `gui/main_window.py:131-149` |
| Close button and tab-change bindings | `gui/main_window.py:151-154` |
| One-line forwarders to the handlers | `gui/main_window.py:210-276` |
| Status text "Step N: <name>" on every tab change | `gui/main_window.py:331-342` |
| Close confirmation "Processing Active" | `gui/main_window.py:344-355` |
| `run()`: welcome status, three start-up log lines, `mainloop()` | `gui/main_window.py:391-398` |
| Title strip widget (70 px, Dark Mode tick placed once at x = window width − 150) | `gui/widgets/title_bar.py:17-51` (position `:27`) |
| Status bar widget ("💡 <message>" left, "v3.0 \| Production Ready" right, 30 px) | `gui/utils/status_bar.py:15-37` |

#### Input → Output

Restored settings (A-F27) → the window with its tabs; the status bar shows "Step N: <Setup|Validation|Processing|Extract>" whenever the tab changes.

#### User interaction

Opens automatically after A-F1. The user clicks tabs, the Dark Mode tick (A-F28) and the window's close button.

#### Concrete trace

Scripts `scripts/t_f2_window.py` and `scripts/t_f2b_window.py` (RAN, me, Xvfb with a 1920×1080 virtual screen, empty settings file):

1. `GSTOrganizerApp()` (`gui/main_window.py:46-55`) → title `GST File Organizer v3.0`, geometry `1200x800+360+140` (centred by `:66-70`), minimum `(1000, 700)`, theme `clam`.
2. Tabs (from the four tab classes): `['📁 Step 1: Setup', '✅ Step 2: Validation', '🚀 Step 3: Processing', '🔄 Step 4: Extract']`.
3. `ExtractTab` is built without scanning (`:144-146`); `ExcelHandler()` (`:116`) created the folder `gst_excel_temp` in the temp folder at start-up (RAN: it appeared in my scratch `TMPDIR`).
4. `run()` sets the status to "Welcome to GST File Organizer v3.0" (`:393`). RAN in the same order as `run()`: right after that the status read `💡 Welcome to GST File Organizer v3.0`, and as soon as Tk processed its first events it read **`💡 Step 1: Setup`** — the first tab-change event (`:331-338`) overwrote the welcome text before the user could see it.
5. Selecting tab 2 → `💡 Step 2: Validation`; tab 4 → `💡 Step 4: Extract` (`:334-338`; the list of four names was fixed in `bfcb24a`).
6. Dark Mode tick: placed at x = **1050** in a 1200-pixel window (`gui/widgets/title_bar.py:27`).
7. Window resized to the minimum **1000×700**: tick still at x = 1050 → **outside the window**; status bar **not shown** (`winfo_ismapped` = 0). Resized to 1600×900: tick still at x = 1050 (it does not follow the right edge); status bar shown.
8. With `is_processing` set to True (as during Step 3), the close button (`on_closing`, `:344-355`) showed "Processing Active" / "Processing is still running. Stop and exit?"; after "Yes" `stop_requested` became True and the window was destroyed **1.04 s** later (`:353`, `after(1000, destroy)`).

#### What the trace shows against what the feature claims

> **BUG (for Section 7): at the minimum window size the Dark Mode tick and the status bar are not visible** (RAN). The tick is placed once from the width at creation (`gui/widgets/title_bar.py:27`); the status bar is packed after the notebook (`gui/main_window.py:138, 149`), and with Tk's `pack` the last-packed widget loses its space first. HIGH.

> **BUG (for Section 7): closing during Step 3 kills the work mid-client; closing during Step 4 gives no warning at all.** `on_closing` checks only Step 3's flag (`gui/main_window.py:346`), not Step 4's own `is_processing` (`power_query_extractor/gui/extractor_window.py:119, 872`). During Step 3 the window is destroyed after 1 s while the daemon worker is still inside the current client (RAN GUI exp4: the client's Sales report, text report and the summary workbook were never written, and the worker died with "main thread is not in main loop"). In the real program the process ends and the daemon thread is killed wherever it is, possibly mid-copy or with a hidden Excel left running (MEDIUM for the Excel after-effects; READ).

- The welcome status text is overwritten at once (RAN, step 4) — cosmetic, MEDIUM-HIGH.
- Closing does not save settings (`:344-355`; READ) — see A-F27.
- The tabs are built before the cached settings are applied (`:53` then `:55`), so the restored paths reach the tabs through the variables' traces (READ); at start-up the Setup checklist therefore checks the four remembered paths on the window's thread (`gui/tabs/setup_tab.py:156` then the traces), which can stall the window on an unreachable network share (READ, MEDIUM).
- Unused leftovers here: imports `os`, `sys`, `Path`, `MESSAGES`, `get_timestamp`, `CollapsibleFrame` (`gui/main_window.py:7-25`), `self.widgets` (`:107`), the module `logger` (`:38`), and three forwarders nobody calls (`:244-245, 256-257, 275-276`) (READ; detail for Section 2).

Rules involved: A-R131 (window, tabs, status per tab), A-R132 (automatic tab switches), A-R155 (title strip and status bar), A-R156 (version label "v3.0" ⚠ — not a GST value, but it never changed). None is a [GST RULE].

#### Older documents against the code

`PROJECT_AUDIT_ANNUALSTATEMENT.md:771-779, 817-820` (UI1, UI2, UI7) still describe the window, title strip and status bar correctly. `README.md:1, 361-364` (RM1) "Version 3.0 … Last Updated 29th October 2025" — the version label is unchanged although A has 34 commits of its own since.

#### History

Window, title strip and status bar from `edac6a3` (shared - before the split on 30-10-2025), unchanged except: `bfcb24a` (26-07-2026) added the Step 4 tab and made the tab-name lookup safe for four tabs; `510f7b4` added `workflow_mode`; `c89c1d4` and `725fbaf` changed the dark-mode switch (all A only (after the split)).

### A-F3: Choose source folder, target folder and templates (Setup)

#### What it does

Step 1 has four path rows — source folder, ITC report template, Sales report template, target folder. Each row is a text box the user can type into and a "Browse…" button that opens a folder or file dialog. A Browse choice is written into the box, announced in the status bar and saved in the settings file.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| The four rows, each an Entry bound to a Tk variable plus a ttk "Browse…" button with `command=` | `gui/tabs/setup_tab.py:353-366` (`_create_inputs`), `:368-391` (`_path_row`; entry `:378`, button `:384`, red problem label `:388-391`) |
| Source dialog "Select Source Folder"; status "Source: <path shortened to 60>"; save | `gui/handlers/file_handler.py:19-25` |
| ITC dialog "Select ITC Template", filter `Excel Template` = `*.xltx *.xlsx`; status "ITC template selected"; save | `gui/handlers/file_handler.py:27-36` (filter `:31`) |
| Sales dialog, same filter; status "Sales template selected"; save | `gui/handlers/file_handler.py:38-47` (filter `:42`) |
| Target dialog "Select Target Folder"; status "Target: <shortened>"; save | `gui/handlers/file_handler.py:49-55` |
| The four Tk variables | `gui/main_window.py:82-86` |
| Forwarders | `gui/main_window.py:210-220` |
| Path shortening for the status bar | `utils/helpers.py:394-416` (`truncate_path`) |
| What a save writes | `gui/main_window.py:358-371` → `gui/utils/cache_manager.py:29-48` |

#### Input → Output

The user's choice (dialog) or typing → the four Tk variables. A Browse also writes the settings file `~/.gst_organizer_cache.json` and a status-bar message. Typing only changes the variable (the checklist, A-F4, reacts to it).

#### User interaction

"Browse…" buttons or typing in the boxes. The processing step later re-reads the template boxes for every client (see A-F15).

#### Concrete trace

Script `scripts/t_f3_paths.py` (RAN, me, Xvfb, empty home folder; the dialogs were replaced by fixed answers and recorded):

1. No settings file at first.
2. "Browse…" for the source → `askdirectory(title='Select Source Folder')` (no starting folder is passed) → variable = `<base>/source`; status `💡 Source: <base>/source` (shorter than 60 characters, so not shortened); settings file now holds `source_folder = <base>/source` (`gui/handlers/file_handler.py:21-25`).
3. "Browse…" for the ITC and Sales templates → both dialogs called with `filetypes=[('Excel Template', '*.xltx *.xlsx')]` (`:31, :42`); status `💡 Sales template selected`.
4. The target folder **typed** into its box (variable set, no Browse) → the settings file still held `target_folder = ''`. Only after "Browse…" for the target did it hold the path (`:53-55`).
5. Keys in the settings file: `client_name_max_length, dark_mode, include_client_name, itc_template, processing_mode, recent_folders, sales_template, source_folder, target_folder, workflow_mode` (writer `gui/main_window.py:360-370`, `recent_folders` added by `gui/utils/cache_manager.py:32-43`).
6. The "Max client folder name length" box emptied (as a user deleting its text would), then "Browse…" for the source → the new folder **was** put into the box, but saving raised `TclError: expected floating-point number but got ""` (from `self.client_name_max_length.get()` at `gui/main_window.py:369`) and the settings file kept the **old** source folder.

#### What the trace shows against what the feature claims

> **BUG (for Section 7): typed paths are not remembered, and one bad value in the length box stops all saving.** A typed path is saved only if some later Browse/tick happens to save (step 4); when the length box does not hold a whole number, every Browse fails to save (step 6) — in the real window the error goes to Tk's error printout on the console, so the user sees nothing. HIGH (RAN).

- The ITC/Sales dialogs offer only `.xltx` and `.xlsx`, while the checklist (A-F4) also accepts `.xltm` and `.xlsm` (`gui/utils/setup_validation.py:21`) — a macro template can only be typed in (READ; A-R52). Whether macro templates work at all is open (see A-F14).
- The dialogs never open in the folder already chosen (no `initialdir`; RAN: only `title` and `filetypes` were passed).
- `truncate_path` can return text longer than its 60-character limit and doubles a separator on a root such as `C:\` (RAN CORE exp3, TRACE T8 §7) — cosmetic.
- The template boxes are re-read for every client during Step 3, so changing a template mid-run changes it for later clients (RAN GUI exp8; see A-F15).

Rules involved: A-R51 (what Step 3 later requires), A-R52 (accepted template types), A-R114 to A-R116 (settings file, keys, when saved), A-R157 (shortened paths in the status bar). None is a [GST RULE].

#### Older documents against the code

`README.md:119-122` (RM16) says templates are `.xlsx` or `.xltx` — matches the dialog filter but not the checklist (which also accepts `.xltm`/`.xlsm`). `README.md:137-146` (RM18) lists the Setup steps but not the workflow card or the checklist. The hover help (`gui/tabs/setup_tab.py:77-82`) says "(.xlsx / .xltx)" — matches the filter.

#### History

`gui/handlers/file_handler.py` is from `edac6a3` (shared - before the split on 30-10-2025; `6237597`, also shared, fixed only the scan count line). The rows themselves were rebuilt in `322717f` (25-07-2026, A only (after the split)) with `command=` buttons and tooltips.

### A-F4: Live setup checklist

#### What it does

At the top of Step 1, a four-row checklist shows for each input a ✓ (set and found), ○ (not chosen yet) or ✕ (chosen but not found / wrong kind) with a short message. A "pill" in the header reads "N of 4 items set" or "Ready - all 4 items set", and a coloured strip under the Scan buttons says what is still needed. The rows are re-checked 350 ms after the user stops typing. "🔍 Scan Files" and "🔄 Re-scan" are enabled only when the source folder exists.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Checklist card (four tinted rows, tooltips) | `gui/tabs/setup_tab.py:282-314` |
| Traces on the four path variables → re-check scheduled 350 ms later (`REVALIDATE_DELAY_MS = 350`) | `gui/tabs/setup_tab.py:135, 506-519` |
| Existence check remembered per exact path text, cleared only on a forced refresh | `gui/tabs/setup_tab.py:521-529, 534-535` |
| Repaint rows, problem labels, accent bars, tooltips, pill | `gui/tabs/setup_tab.py:531-583` |
| Enable/disable Scan and Re-scan; strip text | `gui/tabs/setup_tab.py:585-608` |
| Tk-free decision logic: states, folder and template checks, template suffixes `{'.xlsx', '.xltx', '.xltm', '.xlsm'}`, `can_scan`, `can_process`, labels, summary line | `gui/utils/setup_validation.py:13-115` (suffixes `:21`; folder check `:47-53`; template check `:56-64`; shortening `:67-72`) |
| Forced refresh at start-up and after a dark-mode switch | `gui/tabs/setup_tab.py:156`; `gui/main_window.py:188-194` |

#### Input → Output

The four path texts → row icons and messages, a red message under a wrong entry, the pill text, the strip text, and the enabled/disabled state of Scan and Re-scan.

#### User interaction

Automatic: whenever a path changes (typed, browsed or restored from settings), and forced at start-up and after a dark-mode switch.

#### Concrete trace

Script `scripts/t_f4_checklist.py` (RAN, me, real window under Xvfb). Scenario: the user picks a real source folder and a real ITC template (`ITC_Template.xlsx`), but points the Sales template at a file named `Sales_Template.xls` and types a target folder that does not exist yet.

1. At start-up (`gui/tabs/setup_tab.py:156`, forced): all four rows `○ Not selected yet`; pill `0 of 4 items set`; strip `Select a valid source folder to scan`; Scan and Re-scan **disabled** (`:586-591`, `can_scan` false).
2. The four variables are set. Each write fires the trace (`:506-510`), which cancels and re-schedules the check (`:512-519`). **100 ms later** the pill still read `0 of 4 items set` (debounce working).
3. After the 350 ms pause `refresh_status()` (`:531-583`) called `evaluate_setup` (`gui/utils/setup_validation.py:75-85`) with the remembered-existence checker (`setup_tab.py:521-529`):
   - source → exists → `✓ <base>/source` (`:47-53`; shown whole because it is ≤ 46 characters);
   - ITC → exists, suffix `.xlsx` in the list → `✓ ITC_Template.xlsx` (`:56-64`);
   - Sales → exists, suffix `.xls` not in `{'.xlsx', '.xltx', '.xltm', '.xlsm'}` → `✕ Not an Excel template`, and the same text in red under the entry (`setup_tab.py:557-559`);
   - target → not on disk → `✕ Folder not found`.
   Pill `2 of 4 items set`; strip `Ready to scan · still needed for Step 3: Sales template, Target folder` (`:601-604`); Scan and Re-scan **enabled** (only the source matters, `setup_validation.py:88-91`).
4. The target folder was then **created on disk** and the Sales box corrected to `Sales_Template.xlsx`. After the pause: Sales `✓`, but target **still `✕ Folder not found`** — its text had not changed, so the remembered "does not exist" answer was reused (`setup_tab.py:524-529`). Pill `3 of 4 items set`.
5. A forced refresh (`refresh_status(force=True)`, what start-up and a dark-mode switch do) cleared the memory (`:534-535`): target `✓ ...e-scratch/run/A-S3a/work/f4/target_not_yet` (the start of the path cut off, `setup_validation.py:67-72`), pill `Ready - all 4 items set`, strip `Ready to scan, and to process afterwards`.
6. The source box pointed at a **file** (`.../source/ClientB notes.xlsx`): row `✓`, pill still `Ready - all 4 items set`, Scan enabled. Pressing Scan with a file as the source ends in the error box "Scan Error: Not a directory: …" (`core/file_parser.py:154-155`; RAN TRACE T7).

#### What the trace shows against what the feature claims

> **BUG (for Section 7): the checklist can show a stale ✕ (or ✓) for a path whose state changed on disk.** The yes/no answer is remembered per exact text and forgotten only at start-up and on a dark-mode switch (RAN, step 4). A network share that comes back, or a folder created after the first check, stays ✕ until the text changes or the program restarts. HIGH.

> **BUG (for Section 7): "set and present" only means "exists".** A file passes as the source or target folder, and a folder named `x.xlsx` passes as a template (RAN step 6 and TRACE T7); the checklist then says "Ready". HIGH.

- "Ready - all 4 items set" only colours the pill and the strip. Nothing disables Step 3's Start buttons; Step 3's own gate (`gui/handlers/file_handler.py:132-151`, A-R51) is weaker: it does not check the target exists or the template type (READ; RAN TRACE T7: a `.txt` "template" and a missing target pass that gate).
- Three different template-type lists exist: the Browse filter `*.xltx *.xlsx` (`gui/handlers/file_handler.py:31, 42`), this checklist `{.xlsx, .xltx, .xltm, .xlsm}` (`gui/utils/setup_validation.py:21`), and the unused `TEMPLATE_EXTENSIONS = {'.xlsx', '.xltx', '.xltm'}` (`utils/constants.py:197`) (READ).
- The existence checks run on the window's own thread (`setup_tab.py:521-529`), so on an unreachable network path each new text can freeze the window for the network time-out (READ, MEDIUM).
- `_short` keeps the **end** of a long path ("..." + last 42 characters) although its docstring says "Trim the middle" (`setup_validation.py:67-72`; RAN step 5) — cosmetic.

Rules involved: A-R113 (350 ms delay, remembered answers), A-R133 (row states, pill, strip texts), A-R52 (accepted template types). None is a [GST RULE].

#### Older documents against the code

The checklist is not described in `README.md` at all (docs notes §4.1 item 8). `PROJECT_AUDIT_ANNUALSTATEMENT.md:781-794` (UI3) describes the old two-column Setup tab with a welcome banner — no longer true since `322717f`.

#### History

Added in `322717f` (25-07-2026, together with `gui/utils/setup_validation.py` and the tooltips); forced refresh after a theme switch in `c89c1d4`; dark-mode colours in `169c7b5` (all A only (after the split)).

### A-F5: Workflow mode (organise-only vs full pipeline)

#### What it does

A card on Step 1, "WHAT SHOULD THIS RUN DO?", has two radio buttons: "📁 Organise files & create reports" (value `organize`, the default) and "🔄 Full pipeline" (value `full`). The choice is saved at once. It only matters at the very end of a Step 3 run: in `organize` mode Step 3 writes one hint line into its log; in `full` mode Step 3 switches to Step 4, puts the target folder into Step 4's folder box and scans it, so the client list is ready — but the Power Query refresh itself still has to be started by the user.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Values, names and descriptions | `utils/constants.py:163-172` |
| The card and radios (`command=self._on_workflow_change`), hover help, footnote "Step 4 stays available either way …" | `gui/tabs/setup_tab.py:316-345`; help texts `:95-101` |
| Save on change (errors swallowed) | `gui/tabs/setup_tab.py:347-351` |
| The Tk variable, default `organize` | `gui/main_window.py:88-90` |
| Hand-over decision (runs on the worker thread) | `gui/handlers/processing_handler.py:136-159` |
| Hand-over action on the window's thread: select Step 4, set folder, scan, log | `gui/handlers/processing_handler.py:161-175` |
| Called at the end of every run, also after Stop | `gui/handlers/processing_handler.py:519` |

#### Input → Output

The radio choice → saved setting `workflow_mode`; at the end of Step 3: a log line (`organize`) or Step 4 selected, its folder set to the target and its client list filled (`full`).

#### User interaction

Radio button on Step 1; the effect appears automatically when a Step 3 run ends.

#### Concrete trace

From the end-to-end run `scripts/t_pipeline.py` (RAN, me, real window and real worker thread inside Tk's main loop; output `out/pipeline.txt`, sections S6 and S7).

**Run 1, `organize`:** at the end of `process_files_thread` (`gui/handlers/processing_handler.py:519`) `_maybe_continue_to_extraction()` read `workflow_mode` (a Tk variable read from the worker thread, `:145`), found `organize`, and logged:
```
ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query.
```
The tab stayed on Step 3 (index 2).

**Run 2, `full`:** the worker logged `🔄 Full pipeline: moving to Step 4 to extract…` (`:157`) and scheduled `_open_extraction_tab` on the window's thread (`:159`). That function (`:161-175`):
1. selected Step 4 (RAN: tab index 3);
2. took `self.app.extract_tab.panel` and set its folder box to the Step 1 target (`:165-168`) → RAN: `<base>/target`;
3. called the panel's `scan_folder()` (`power_query_extractor/gui/extractor_window.py:611-668`), which chose the newest `Annual Statement-*` folder by date modified → `Annual Statement-211026 1030`, and listed both clients, each with its newest `Version-*` folder `Version-211026 1045` and both report ticks ticked;
4. logged `✅ Step 4 ready - review the client list and press Process to refresh.` (`:170-172`).
Step 4's "🚀 Start Processing" button was enabled; **no refresh was started** (RAN: the panel log ends with `Found 2 clients`).

#### What the trace shows against what the feature claims

> **BUG (for Section 7): the on-screen descriptions promise more, or less, than the code does.** (a) "Full pipeline – Also refresh Power Query and extract the data (slower)" (`utils/constants.py:170`) and the hover help "…and then open each report in Excel to refresh its Power Query and pull the figures out" (`gui/tabs/setup_tab.py:98-100`): the refresh is **not** started (RAN; deliberate per the docstring `processing_handler.py:139-142`). (b) "Organise … Fast - no Excel automation involved" (`setup_tab.py:96-97`): on Windows with pywin32, Step 3 builds every report by driving Excel through COM (`core/excel_handler.py:90-95`; READ). (c) The log says "press Process" but the button is "🚀 Start Processing" (`extractor_window.py:330`). HIGH.

- The hand-over writes the target into Step 4's folder box unconditionally (`processing_handler.py:166-168`), overwriting a folder the user chose in Step 4, and without updating Step 4's own bookkeeping `_pushed_folder` (`gui/tabs/extract_tab.py:121-131`); after that, later changes of the Step 1 target may stop flowing into Step 4 (READ, MEDIUM).
- The hand-over also happens after the user pressed **Stop** (`:519` runs after the loop's `break`; READ, HIGH).
- If the extractor panel could not be loaded, `panel` is None and the log shows the technical message "Could not open Step 4 automatically: 'NoneType' object has no attribute 'folder_path'" — the guard at `:153-155` tests the tab, not the panel (READ, HIGH).
- The success line is logged even when Step 4's scan found nothing (the scan reports into Step 4's own log and returns normally, `extractor_window.py:625-628`; READ).

Rules involved: A-R154 (workflow values and hand-over). Not a [GST RULE].

#### Older documents against the code

The design note `docs/superpowers/specs/2026-07-25-merge-extractor-design.md:73-81, 85-87` (P14-P18): the setting and its saving were built as planned (P14); "`full` — also refresh PQ and extract" **differs** (only switch + scan, P15); "In `organize` mode Step 3 finishes with a 'Continue to extraction' button" was **never built** (P17; only a log line); "folder follows Setup until overridden" is built except for this hand-over (P18).

#### History

Added in `510f7b4` (26-07-2026 22:58, A only (after the split)). Covered by `tests/test_workflow_mode.py` (logic with fakes, incl. "test_full_pipeline_does_not_start_the_refresh").

### A-F6: Maximum client folder name length setting

#### What it does

A number box on Step 1, "Max client folder name length:" (arrows from 15 to 100, default 35), with the hint "(default 35 · names over <limit> are shortened)". The value limits the length of the client folder name `<client>-<STATE CODE>` (the "Level 2" folder inside each run folder). Nothing else is shortened by it.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Label, number box (`Spinbox` from 15 to 100, bound to the `IntVar`), arrow-click handler (0 or less → 35, update hint, save) and hint | `gui/tabs/setup_tab.py:429-451` (handler `:434-438`, box `:440-443`), hint text `:610-615`; hover help `:119-123` |
| The variable (default 35), saved and restored (0 or less → 35) | `gui/main_window.py:93, 369, 386-387` |
| Passed to the organiser when Step 3 starts | `gui/handlers/processing_handler.py:194` |
| Stored (0 or less → 35) | `core/file_organizer.py:30, 42` |
| Used only for the Level-2 folder name | `core/file_organizer.py:83-87, 97` → `utils/helpers.py:620-644` (`create_client_state_key`) |
| Also used (with fallback 10) for the long-name warning at Start, only when the global "include client name" tick is on | `gui/handlers/processing_handler.py:90-113` |

#### Input → Output

A whole number → the maximum length of the Level-2 folder name `<client>-<STATE CODE>`.

#### User interaction

Arrow buttons (these save the setting) or typing (neither checked nor saved).

#### Concrete trace

Script `scripts/t_edges_f6_f8_f10.py` (RAN, me): `FileOrganizer.create_client_structure` for `<Client B>` (client text "ClientB Exports Private Limited", state "Gujarat"), with "include client name" on so the Level-4 folders show the name:

1. `create_client_state_key` (`utils/helpers.py:620-644`): "Private" → "Pvt", "Limited" → "Ltd" (`:624-625`), state code `GJ` (`:628`), key `ClientB Exports Pvt Ltd-GJ` (26 characters, `:631`).
2. Limit 35: 26 ≤ 35 → unchanged. Limit 20: 26 > 20 → room for the client part = 20 − 2 − 1 = 17 > 5 → client cut to 17 characters (`:634-639`). Limit 8: room = 8 − 3 = 5, not > 5 → the **whole key** is cut to 8 characters (`:640-642`) and the state code is lost.
3. The Level-4 folder uses `sanitize_filename(client)` (`core/file_organizer.py:90, 119`), which the limit never touches.

RAN output:
```
limit  35: Level 2 = 'ClientB Exports Pvt Ltd-GJ'     Level 4 = 'ITC (ClientB_Exports_Pvt_Ltd)'
limit  20: Level 2 = 'ClientB Exports P-GJ'           Level 4 = 'ITC (ClientB_Exports_Pvt_Ltd)'
limit   8: Level 2 = 'ClientB '                       Level 4 = 'ITC (ClientB_Exports_Pvt_Ltd)'
limit typed as text "35" -> TypeError '>' not supported between instances of 'str' and 'int'
```

From the end-to-end run (`out/pipeline.txt`, S8, RAN): with the box **emptied** and "🚀 START PROCESSING" pressed, the worker failed at `gui/handlers/processing_handler.py:194` (`IntVar.get()` on an empty text) and the Step 3 log showed only:
```
🚀 ════════ PROCESSING STARTED ════════
💥 Fatal error: expected floating-point number but got ""
```
plus an error box "Processing Error" raised **from the worker thread** (recorded thread name `Thread-3 (process_files_thread)`). No folder was created. The progress area still showed the **previous** run's `📊 Progress: 100.0%` and `Processing complete | Total time: 0m 0s`.

#### What the trace shows against what the feature claims

> **BUG (for Section 7): the setting does not do what its help text says.** "Longest client name allowed in a folder name before it is shortened" (`gui/tabs/setup_tab.py:119-120`) and the hint "names over <limit> are shortened" (`:615`): only the Level-2 folder is shortened, and the limit counts the whole `<client>-<STATE>` text, not the name; the Level-4 folders `ITC (<client>)`, the copied file names and the report names keep the full name (RAN). The long-name warning at Start (`processing_handler.py:90-113`) appears only when the **global** tick is on and compares the raw name with this limit, so it warns about folders the limit does not shorten; with a per-client tick (A-F10) no Start warning appears at all (RAN: `<Client B>`, 31 characters, got named folders without a Start warning). HIGH.

> **BUG (for Section 7): a typed value is never checked.** An empty or non-numeric box makes every settings save fail (A-F3) and ends any Step 3 run at once with "Fatal error" (RAN). A typed value of 8 or less (for a two-letter state code) drops the state code from the folder name (RAN); values above 100 are accepted. The hint does not change when typing (`_update_length_hint` is called only at creation and by the arrows, `setup_tab.py:437, 451`). HIGH (mechanism), MEDIUM (likelihood of typing).

- After a fatal error at the start of a run the progress label keeps the previous run's "100% / Processing complete" text (RAN) — a detail of A-F17.
- A cut client part may end with a space (`'ClientB '`, or `'ClientA Pvt -DL'` in CORE exp4), which becomes part of the folder name (RAN).

Rules involved: A-R35 (Level-2 name, [GST RULE] only because it carries the state code ⚠ state list), A-R36 (length values), A-R136 (long-name warnings), A-R116 (when saved).

#### Older documents against the code

`PROJECT_AUDIT_ANNUALSTATEMENT.md:511-514` (R12) says the client part is cut "keeping -STATE" with a "minimum 5 chars" — the code cuts only if more than 5 characters remain and otherwise loses the state code (PARTLY true). The old hint "Names >10 chars will show warning" was replaced by the live hint in `322717f`.

#### History

`create_client_state_key` and the default 35 are from `edac6a3`; the configurable limit in the Start warning (fallback 10) from `6237597` (both shared - before the split on 30-10-2025). The hint showing the real limit came with `322717f` (A only (after the split)).

### A-F7: Scan source folder and recognise client files (Scan / Re-scan)

#### What it does

"🔍 Scan Files" reads the **top level only** of the source folder, keeps genuine Excel files, matches each file name against six name patterns (one per document type), groups recognised files by "client + state code", works out per client which of the six types are missing or duplicated, fills Step 2 (A-F8) and switches to it when at least one client was found. "🔄 Re-scan" empties the previous results first and then scans again.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Buttons (ttk, `command=`), enabled by A-F4 | `gui/tabs/setup_tab.py:459-467` |
| Scan entry point: checks, status/log, scan, summary, table, switch to Step 2, error box | `gui/handlers/file_handler.py:57-95` |
| Re-scan: empties scan results **in place**, clears summary and table, scans | `gui/handlers/file_handler.py:97-118` (emptying `:102-105`) |
| Scan input check (source set, exists) | `gui/handlers/file_handler.py:120-130` |
| `scan_folder`: reset, folder checks, list files, process each, completeness | `core/file_parser.py:136-177` |
| One listing with `os.scandir`, extension filter `.xlsx/.xls/.xlsm`, metadata check (regular file, suffix in `.xlsx/.xls/.xlsm/.xltx/.xltm`, ≥ 1024 bytes), 8-byte signature check (16 parallel reads when ≥ 4 files), sort by path, any error → empty list | `utils/helpers.py:309-365` (extensions `:327`, size `:303`, signatures `:264-265, 275-287`, parallel `:267-272, 351-357`, sort `:361`, error `:363-365`) |
| File-name matching (six patterns, first match wins, parts trimmed) | `core/file_parser.py:56-134`; patterns `utils/constants.py:13-62` |
| Per file: details, parse, store, client data or variation | `core/file_parser.py:179-224` |
| Client key `"<client as written>-<state code>"` | `core/file_parser.py:226-250` (key `:229-230`); state code `utils/helpers.py:592-618`, table `utils/constants.py:251-292` |
| Missing / duplicate types, completeness %, status | `core/file_parser.py:252-286`; expected types `utils/constants.py:68-75` |
| Statistics for the Step 2 summary | `core/file_parser.py:308-336` |

#### Input → Output

The source folder → `scanned_files` (one record per listed file, keyed by file name), `client_data` (one record per client key with its files per type, missing and duplicate types, completeness, status) and `variations` (unrecognised files, A-F9); the Step 2 summary text and table; a switch to Step 2.

#### User interaction

"🔍 Scan Files" / "🔄 Re-scan" on Step 1 (enabled only when the source folder exists, A-F4).

#### Concrete trace

End-to-end run `scripts/t_pipeline.py`, section S2 (RAN, me, real window). Source folder as in the test bed: 14 entries at the top level (11 workbooks with recognisable or near-miss names, `tiny.xlsx`, one `.csv`, the sub-folder `old/`).

1. The Scan button's own `command=` was invoked → `GSTOrganizerApp.scan_files` (`gui/main_window.py:222-223`) → `FileHandler.scan_files` (`gui/handlers/file_handler.py:57`).
2. `validate_scan_inputs` (`:120-130`): source set and existing → go on. Status "Scanning files..."; log line "🔍 Starting file scan..." (`:63-64`).
3. `FileParser.scan_folder` (`core/file_parser.py:136`): `reset()` (`:148`), folder exists and is a folder (`:152-155`), then `scan_excel_files` (`utils/helpers.py:309`):
   - the one listing returns 14 entries; the extension test (`:327, 330`) keeps the 11 `.xlsx` names **and** `tiny.xlsx` (12); the `.csv` and the folder `old` are skipped — and the workbook inside `old/` is never seen (no sub-folders);
   - the metadata check (`:290-306`) drops `tiny.xlsx` (500 bytes < 1024, `:303`) → 11;
   - 11 ≥ 4, so the 8-byte signature reads run in a pool of 11 threads (`:351-355`); all 11 start with `PK\x03\x04`;
   - sorted by full path (`:361`).
4. For each of the 11 files (`core/file_parser.py:165-167`): a progress message is set in the status bar (`file_handler.py:67-68`, not painted while the scan runs on the window's thread), then `_process_file` (`:179`) → `get_file_info` with the listing's cached details (`:188`) → `parse_filename` (`:196`):
   - `GSTR-2B-Reco-ClientB Exports Private Limited-Gujarat-Q1.xlsx` → first pattern `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` ⚠ matches; client `ClientB Exports Private Limited`, state `Gujarat`, period `Q1` kept in `metadata` (`:98-106`); both texts pass `validate_client_name` (`:108-119`); `break` (`:128`);
   - `GSTR3B-ClientA Traders-Maharashtra-Apr.xlsx` → pattern `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` ⚠ → type "GSTR-3B Export", month `Apr`;
   - `ClientB notes.xlsx` and `Sales Reco-ClientA Traders-Maharashtra-Q1.xlsx` match nothing → WARNING "Could not parse filename: <name>" goes to the log (`:131`; RAN: printed with the file name) → error "No matching pattern found" (`:132`).
   - Recognised files go to `_add_to_client_data` (`:210-211`): key = client + `-` + `get_state_code(state)` (`:229-230`) → `Maharashtra` → `MH`, `Gujarat` → `GJ` from the state table ⚠ (`utils/constants.py:251-292`). The other two become variations (`:212-220`, see A-F9).
5. `_analyze_client_completeness` (`core/file_parser.py:252-286`) against the six expected types ⚠ (`utils/constants.py:68-75`):
   - `ClientA Traders-MH`: all six types found; GSTR-3B has 2 files but that type may repeat (`:266-268`) → no duplicates → **Complete ✅**, completeness 100.0;
   - `ClientB Exports Private Limited-GJ`: found GSTR-2B Reco and GSTR-3B Export → missing `['Annual Report', 'IMS Reco', 'Sales', 'Sales Reco']` (sorted, `:260`) → **Missing 4 files ⚠️**, completeness 2 ÷ 6 × 100 = 33.3 (`:274-275`).
6. Back in `scan_files`: status `🔍 Found 11 Excel files` (`file_handler.py:75-76`); `update_summary()` and `update_client_tree()` (`:79-80`); clients found → status "Found 2 clients" and log "✅ Found 2 clients" (`:82-85`); `notebook.select(1)` (`:88`) → Step 2. RAN: the tab index became 1 and the status bar read **`💡 Step 2: Validation`** — the tab-change event (`gui/main_window.py:331-338`) overwrote "Found 2 clients" at once.

The Step 2 summary text produced (RAN, `gui/main_window.py:294-328`):
```
📊 SCAN RESULTS SUMMARY:
════════════════════════════════════════════
📁 FILES ANALYSIS:
   Total Excel Files Found: 11
   ✅ Successfully Parsed: 9
   ❌ Unparsed/Variations: 2
   📊 Parsing Success Rate: 81.8%

👥 CLIENTS ANALYSIS:
   Total Clients Identified: 2
   ✅ Clients with All Files: 1
   ⚠️ Clients with Missing Files: 1
   📈 Completeness Rate: 50.0%

📋 FILE TYPE DISTRIBUTION:
   • Annual Report: 1 files
   • GSTR-2B Reco: 2 files
   • GSTR-3B Export: 3 files
   • IMS Reco: 1 files
   • Sales: 1 files
   • Sales Reco: 1 files

🔍 FILE VARIATIONS: 2 files

💾 TOTAL DATA SIZE: 56.5 KB
```

#### What the trace shows against what the feature claims

> **BUG (for Section 7): files are dropped without any message.** `tiny.xlsx`, the `.csv` and the workbook in the sub-folder `old/` appear nowhere — not in the count ("Total Excel Files Found: 11"), not as variations, not in the log (RAN). The same happens to `.xlsb`, `.xltx`, Excel lock files `~$…` and any file whose first bytes are not an Excel signature (RAN TRACE T6, T8). A correctly named `.xlsm` is listed but no pattern accepts it (every pattern ends `\.xlsx?$`), so it always becomes a variation (RAN TRACE T1 row 7). HIGH.

> **BUG (for Section 7): a hyphen inside a client name silently files the document under a wrong client and state** for the GSTR-2B Reco, Sales Reco and Annual Report patterns (`[^-]+` groups); RAN TRACE T1: `GSTR-2B-Reco-Alpha-Test-Traders-Maharashtra-Q1.xlsx` → client `Alpha`, state `Test`, key `Alpha-TES`; `Sales-Reco-…` is read as a **Sales** file of client `Reco`. The GSTR-3B and Sales patterns reject such names instead. HIGH.

> **BUG (for Section 7): a scan cannot be told apart from an empty folder when the share fails.** Any listing error makes `scan_excel_files` return an empty list (`utils/helpers.py:363-365`), and the screen says "No valid GST files found" (`file_handler.py:89-91`; READ, RAN TRACE T8 §6 for a missing folder). HIGH.

> **BUG (for Section 7): Re-scan stays usable during Step 3 and pulls the data from under the running job.** `rescan_files` empties `client_data` and `variations` in place (`file_handler.py:102-105`). The worker's next `self.app.client_data[client_key]` (`gui/handlers/processing_handler.py:232`) sits **outside** the per-client error handling, so the KeyError ends the whole run with "Fatal error" (READ, HIGH); the summary workbook's Variations sheet comes out empty (RAN GUI exp8). MEDIUM likelihood.

- The scan runs on the window's own thread (`file_handler.py:71-72`): the window does not repaint or respond until the scan ends; the progress texts are set but never painted (READ, HIGH; commit `4f84ba9` measured about 41 s for 2508 files on the office share).
- "Found 2 clients" is replaced by "Step 2: Validation" in the status bar the moment Step 2 opens (RAN, step 6); the "✅ Found 2 clients" log line is on Step 3, not visible on Step 2.
- Spelling decides identity: `<Client A> Private Limited`, `<Client A> Pvt Ltd` and a lower-case spelling become three clients; two state spellings with the same code merge, and the client then shows the state text of the last file read (RAN TRACE T6, T8; A-R17).
- The parts after client and state (period, IMS date, month, start/end month, year) are kept in `metadata` and never read (`core/file_parser.py:106`; READ, grep) — files of different periods are grouped together [GST RULE] ⚠ (A-R25). The IMS date only has to be 8 digits (`99999999` accepted; RAN TRACE T1).
- An unknown state text gets an invented code: first letters of the first two words (`Unknown Place` → `UP`, the same as Uttar Pradesh) or the first three characters (`TamilNadu` → `TAM`, `27` → `27`) [GST RULE] ⚠ (RAN TRACE T2; A-R22).
- `validate_client_name` results are stored as `warnings` and never shown (`core/file_parser.py:108-119`; READ); errors collected in `self.errors` are never read outside the parser (READ).
- File names of unparsed files go into the log file at WARNING level (`core/file_parser.py:131`) — confidentiality note, see A-F30.

Rules involved (all [GST RULE] where marked): A-R1 to A-R7 (the six patterns ⚠ and how matching works) [GST RULE]; A-R8 to A-R10 (scope, size, signature); A-R11, A-R12 (variations, suggestions ⚠) — A-R12 [GST RULE]; A-R17 (client key) [GST RULE]; A-R21, A-R22 (state table ⚠ and fallback) [GST RULE]; A-R25 (periods ignored ⚠) [GST RULE]; A-R26 to A-R31 (expected types ⚠, duplicates ⚠, completeness, status, statistics) — A-R26 to A-R30 [GST RULE]; A-R109, A-R110 (one listing, parallel reads).

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:131` (F1.5) "`scan_folder()` calls `find_excel_files()`" — no longer true (`scan_excel_files`, `a603c02`, `56aaad9`).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:597` (E4) "Corrupted Excel → excluded and added to variations" — never true: such files vanish silently (`tests/test_scan.py:52` pins this).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:609` (E10) "Empty client name … logged as warning, still processed" — never true: a blank client goes to the variations (see A-F9).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:474, 614` (R7, E15) "a hyphen inside a client name breaks parsing" — partly: for three of the six patterns the file is accepted under a wrong client (above).
- `README.md:124-133` (RM17) "MUST match exactly" — the code also accepts any letter case, `.xls`, a leading "(N) " for GSTR-3B and a trailing "(N)" for IMS.
- `README.md:276` (RM35) "No Excel files found → check extensions (.xlsx, .xls)" — the message is "No valid GST files found", `.xlsm` is also listed, and too-small/invalid files are dropped silently.
- The Step 1 card "📝 EXPECTED FILE NAMES" (`gui/tabs/setup_tab.py:477-503`) shows the six descriptions from `utils/constants.py:19-59` ⚠; it does not mention `.xls`, the "(N)" markers or that the IMS date must be exactly 8 digits (see A-F31).

#### History

Patterns, grouping and completeness from `edac6a3` (shared - before the split on 30-10-2025); `6237597` (shared) made the progress callback safe and fixed the file count line. A only (after the split): `a603c02` (one listing), `56aaad9` (reuse listing details), `4f84ba9` (parallel signature reads), `5c8ce95` (IMS "(N)" suffix accepted), `322717f` (buttons now enabled by the checklist).

### A-F8: Client list with completeness status (Validation)

#### What it does

Step 2 shows, on the left, collapsible instructions, the scan summary text (A-F7) and three action buttons (Dry Run, Export List, Start); on the right, a table with one row per client — tick box, client name, state, status ("✅ Complete" / "⚠️ Missing N files" / "⚠️ Has duplicates"), number of files, missing types, number of duplicated types, and the per-client folder-name toggle "📁 Name" (A-F10). The user ticks clients by clicking the first column, with Space and the arrow keys, or with "☑️ Select All", "☐ Clear All", "✅ Complete Only". The ticked rows become the list of client keys that Dry Run, Start and Export use.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Tab layout (left column fixed 450 px, summary box fixed 200 px high) | `gui/tabs/validation_tab.py:18-42` (`:30`), `:68-92` (`:83`) |
| Collapsible instructions | `gui/tabs/validation_tab.py:44-66`; widget `gui/widgets/collapsible_frame.py:7-41` |
| Dry Run / Export / Start buttons wired with `.bind('<Button-1>')` | `gui/tabs/validation_tab.py:94-142` (`:118, 130, 142`) |
| Selection buttons (also `.bind`), table columns, widths, headings, scrollbars, key and click bindings | `gui/tabs/validation_tab.py:144-247` (buttons `:174-190`, table `:197-221`, bindings `:241-244`) |
| Fill the table from `client_data`; reset ticks and per-client name flags; colour tags `#E8F5E8` / `#FFF4E5`; highlight first row | `gui/handlers/client_handler.py:18-65` |
| Click: first column toggles the tick; 7th column toggles the name flag | `gui/handlers/client_handler.py:67-79` |
| Space / Up / Down | `gui/handlers/client_handler.py:81-116` |
| Toggle, Select All, Clear All, Complete Only (status text contains "Complete") | `gui/handlers/client_handler.py:118-159` |
| Ticked rows → client keys, **rebuilt from the displayed text** `f"{values[0]}-{get_state_code(values[1])}"` | `gui/handlers/client_handler.py:161-173` |
| Summary text | `gui/main_window.py:294-328` |
| Status and completeness as decided by the scan | `core/file_parser.py:252-286, 308-336` |

#### Input → Output

`client_data` from the scan → table rows; the user's ticks → a list of client keys such as `ClientA Traders-MH`, used by A-F11, A-F12 and A-F13.

#### User interaction

The table is filled after a scan; the user clicks, uses Space/arrows, or the three selection buttons.

#### Concrete trace

End-to-end run, section S3 (RAN, me, real `ttk.Treeview`), continuing from the A-F7 scan:

1. `update_client_tree` (`gui/handlers/client_handler.py:18-65`) deleted old rows, made a new empty tick dictionary (`:27`) and inserted one row per client in scan order (`:30-44`). Rows as Tk returned them:
   ```
   I001  ☐  ['ClientA Traders', 'Maharashtra', '✅ Complete', 7, 'None', 0, '☐ No']                                 tag complete
   I002  ☐  ['ClientB Exports Private Limited', 'Gujarat', '⚠️ Missing 4 files', 2,
             'Annual Report, IMS Reco, Sales, Sales Reco', 0, '☐ No']                                         tag incomplete
   ```
   Tag colours `#E8F5E8` (complete) and `#FFF4E5` (incomplete) (`:58-59`); first row `I001` highlighted (`:62-65`). Every per-client name flag was reset to False (`:49`).
2. "✅ Complete Only" (`:147-159`) checked whether the Status text contains "Complete" (`:152`): `I001` ticked, `I002` unticked (it **replaces** the selection); status bar `💡 Selected 1 complete clients`. `get_selected_clients()` (`:161-173`) → `['ClientA Traders-MH']`.
3. "☑️ Select All" (`:133-138`) → both ticked; `get_selected_clients()` → `['ClientA Traders-MH', 'ClientB Exports Private Limited-GJ']`.

Edge run, script `scripts/t_edges_f6_f8_f10.py` (RAN, me): three made-up clients `ClientA Traders` (Maharashtra), a client written `0042` (Maharashtra) and `ClientC Mock` with the state written as `24`:
```
row values as Tk returns them: ['ClientA Traders', 'Maharashtra']
row values as Tk returns them: [42, 'Maharashtra']
row values as Tk returns them: ['ClientC Mock', 24]
   tick only ['ClientA Traders', 'Maharashtra'] -> key ['ClientA Traders-MH'] | exists in client_data: True
   tick only [42, 'Maharashtra'] -> key ['42-MH'] | exists in client_data: False
   tick only ['ClientC Mock', 24] -> AttributeError: 'int' object has no attribute 'strip'
```

#### What the trace shows against what the feature claims

> **BUG (for Section 7): the selected client keys are rebuilt from the table's displayed text, and Tk hands back number-looking text as numbers.** `0042` comes back as `42`, so the key `42-MH` does not exist; a state written as digits comes back as an integer and `get_state_code` crashes at `.strip()` (`utils/helpers.py:601`) (RAN). Consequences: with a numeric state, Dry Run, Start, Export and the folder-name toggle do nothing visible (the error goes to the console); with `0042`, Start reaches `client_data['42-MH']` at `gui/handlers/processing_handler.py:232`, **outside** the per-client error handling, so the whole run stops with "💥 Fatal error: '42-MH'" and no summary workbook (RAN TRACE T8 §2). The table row id (`item`) could have carried the real key. HIGH (mechanism), LOW (likelihood: an all-digit client or a numeric state in the file name).

- "⚠️ Clients with Missing Files" counts every client that is not Complete, including clients that miss nothing but have duplicates (`core/file_parser.py:328`; RAN TRACE T6). The status text shows only "Missing N files" when a client has both missing and duplicated types; the "Extra" column shows the **number of duplicated types**, not of extra files (`client_handler.py:32`; READ).
- The six buttons on this tab are wired with `.bind('<Button-1>')` (`validation_tab.py:118, 130, 142, 190`): they cannot be used from the keyboard and react to the raw mouse press (READ, HIGH).
- Space toggles only the first highlighted row when several are highlighted (`client_handler.py:83-85`; READ).
- After a rescan in dark mode the rows show white text on the pale light-mode tints (`client_handler.py:58-59` resets only the backgrounds; RAN GUI darkscan) — see A-F28.
- The instructions say "⚠️ = missing files" (also used for duplicates) and "Click 'Start Processing' when ready" (the button reads "🚀 START") (`validation_tab.py:52-58`; READ). The folder-name toggle (7th column) is not mentioned anywhere on screen.
- Fixed widths (450 px left column, 200 px summary) on a window whose minimum width is 1000 px (READ).

Rules involved: A-R26 (six expected types ⚠) [GST RULE], A-R27 (duplicates; GSTR-3B may repeat ⚠) [GST RULE], A-R28 (completeness %, rounded to 1 decimal) [GST RULE], A-R29 (status) [GST RULE], A-R31 (summary text), A-R17 (client key) [GST RULE], A-R21/A-R22 (state code ⚠) [GST RULE], A-R138 (table and selection behaviour). Described only; whether "Complete" means the right set of returns is for the author.

#### Older documents against the code

`PROJECT_AUDIT_ANNUALSTATEMENT.md:263-277` (F7.1-F7.3) and `README.md:148-155` (RM19) still describe the table and the selection buttons correctly; neither mentions the number-looking-text problem.

#### History

From `edac6a3` (shared - before the split on 30-10-2025); `1c48722` (29-10-2025, shared) removed a block that re-tagged rows and forced black text; `33d0baf` (23-07-2026, A only (after the split)) changed the folder-name keys to the state **code**. `gui/tabs/validation_tab.py` itself is unchanged since `edac6a3`.

### A-F9: Unrecognised files / variations list

#### What it does

Excel files that pass the scan's checks but whose names match no pattern — or match with an empty client or state — are collected as "variations", each with an issue text and a suggested correct pattern. The window shows only how many there are (Step 2 summary); the list itself appears only later, as the "Variations" sheet of the Step 3 summary workbook.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| No match → error "No matching pattern found" (and a WARNING log line with the file name) | `core/file_parser.py:130-132` |
| A file becomes a variation when not parsed or when client/state is empty: `{filename, path, issue, size, pattern_expected}` | `core/file_parser.py:209-220` |
| Suggested pattern by keyword (first hit wins) | `core/file_parser.py:288-306` |
| Count in the Step 2 summary ("❌ Unparsed/Variations", "🔍 FILE VARIATIONS") | `gui/main_window.py:306, 320-321`; `core/file_parser.py:311, 325, 329` |
| Passed by reference into the Step 3 summary data | `gui/handlers/processing_handler.py:224` |
| Written as the "Variations" sheet (Filename, Issue, Expected Pattern, Size, Action "Rename file") | `core/excel_handler.py:554-585` |

#### Input → Output

Unrecognised files → variation records → a count on Step 2 and, after processing, rows in the "Variations" sheet of `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx`. Variations are never copied.

#### User interaction

None directly: produced by the scan, visible after processing.

#### Concrete trace

End-to-end run, sections S2 and S6 (RAN, me):

1. During the scan, `ClientB notes.xlsx` and `Sales Reco-ClientA Traders-Maharashtra-Q1.xlsx` matched no pattern (`core/file_parser.py:130-132`); `_process_file` added them as variations (`:212-220`) with `_suggest_pattern` (`:288-306`):
   ```
   ClientB notes.xlsx                               | No matching pattern found | Check file naming convention           | 5238
   Sales Reco-ClientA Traders-Maharashtra-Q1.xlsx   | No matching pattern found | SalesReco-ClientName-State-Period.xlsx | 5260
   ```
   (the second name contains "sales" and "reco", `:299-300`; the first contains no keyword, `:306`).
2. Step 2 summary: `❌ Unparsed/Variations: 2` and `🔍 FILE VARIATIONS: 2 files` — numbers only, no names (`gui/main_window.py:306, 320-321`).
3. After run 1 the summary workbook's "Variations" sheet held (RAN, read back with openpyxl):
   ```
   Filename | Issue | Expected Pattern | Size | Action
   ClientB notes.xlsx | No matching pattern found | Check file naming convention | 5238 | Rename file
   Sales Reco-ClientA Traders-Maharashtra-Q1.xlsx | No matching pattern found | SalesReco-ClientName-State-Period.xlsx | 5260 | Rename file
   ```
4. `tiny.xlsx`, the `.csv` and the file in `old/` (dropped by the scan, A-F7) appear neither in the count nor in the sheet.

#### What the trace shows against what the feature claims

> **BUG (for Section 7): the user cannot see which files were not recognised until after processing, and some unrecognised files never appear at all.** There is no on-screen list (only a count); files dropped by the extension, size or signature checks appear nowhere (RAN). HIGH.

- A name with a blank client (e.g. `GSTR3B- -Maharashtra-Apr.xlsx`) is counted as "Successfully Parsed" in the summary but listed as a variation with issue "Could not parse filename" (RAN TRACE T11: parsed 2, unparsed 1, variations 2) — the summary numbers disagree with the list. HIGH.
- The suggestion is a plain substring test: any name containing "ims" (e.g. "claims") gets the IMS pattern (`core/file_parser.py:295`; READ).
- A Re-scan during Step 3 empties the same list the summary workbook later writes, so the sheet says "No file variations found" (RAN GUI exp8; see A-F7).
- Unparsed file names are written to the log at WARNING level (`core/file_parser.py:131`).

Rules involved: A-R11 (what becomes a variation), A-R12 (suggestion texts ⚠ return names) [GST RULE].

#### Older documents against the code

`README.md:156` (RM20) "Review variations tab for naming issues" — never true: there has never been a variations tab (`git grep` of the tabs at `47a1492`: nothing). `PROJECT_AUDIT_ANNUALSTATEMENT.md:597` (E4) says files with a bad signature are "added to variations" — never true at the snapshot (they vanish).

#### History

From `edac6a3` (shared - before the split on 30-10-2025). `a603c02` (A only (after the split)) removed a variations branch for invalid files that could no longer run.

### A-F10: Per-client option to include the client name in folder names (with global override)

#### What it does

Normally the three category folders inside each Version folder are called `GSTR3B`, `ITC` and `Sales`. The client's cleaned name can be added in brackets — `GSTR3B (<client>)`, `ITC (<client>)`, `Sales (<client>)` — either for every client (the Step 1 tick "Include client name in Level 4 folder names", which wins) or for single clients (a click in the "📁 Name" column of the Step 2 table). Long names trigger warnings.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Global tick (`command=app.update_global_folder_setting`) and the note "Overrides the per-client setting on the next tab." | `gui/tabs/setup_tab.py:417-427` |
| Every scan resets each client's flag to False and shows "☐ No" | `gui/handlers/client_handler.py:43, 49` |
| Click in the 7th column → toggle | `gui/handlers/client_handler.py:76-79` |
| Per-client toggle: key rebuilt from the row text, flip, "☑ Yes"/"☐ No", warning above **10** characters (fixed), unticking one client while the global tick is on turns the global tick off with a message | `gui/handlers/client_handler.py:175-209` (10 at `:191`) |
| Global tick ON → every row ON; OFF → nothing changes; always saves settings | `gui/handlers/client_handler.py:211-223` |
| Decision per client: global ON → name; else per-client flag under key `"<client>-<state code>"` | `core/file_organizer.py:117-139` (`:122-134`) |
| Folder name templates ⚠ | `utils/constants.py:91-95` |
| Global flag read once; the per-client dictionary passed **by reference** (live) | `gui/handlers/processing_handler.py:192-193` |
| Start-time warning (only when the global tick is on; configured limit, fallback 10) | `gui/handlers/processing_handler.py:86-113` |

#### Input → Output

Global flag + per-client flags keyed `"<client>-<STATE CODE>"` → the names of the three Level-4 folders of each client (and therefore the folder paths written into the reports' Links sheets, A-F14/A-F15).

#### User interaction

The Step 1 tick (saved to settings) and clicks in the 7th column of the Step 2 table (not saved).

#### Concrete trace

End-to-end run, sections S3 and S6 (RAN, me):

1. After the scan, the per-client flags were `{'ClientA Traders-MH': False, 'ClientB Exports Private Limited-GJ': False}` and both rows showed "☐ No" (`client_handler.py:43, 49`).
2. A click in the "📁 Name" cell of the `<Client B>` row (`:76-79` → `toggle_folder_name_setting`, `:175-209`): key rebuilt as `ClientB Exports Private Limited-GJ` (`:181`), flag flipped to True (`:184-185`), cell "☑ Yes" (`:189`); the name has 31 characters > 10 (`:191`), so the box "Long Client Name Warning" asked: "Client name 'ClientB Exports Private Limited' has 31 characters. This will create long folder names. Continue?" — answered Yes, flag kept.
3. Run 1 (global tick off): `FileOrganizer(..., include_client_name=False, client_folder_settings=<the live dictionary>)` (`processing_handler.py:189-195`). In `create_client_structure` (`core/file_organizer.py:117-139`):
   - `<Client A>`: flag False → " (ClientA_Traders)" removed (`:133-134`) → `GSTR3B`, `ITC`, `Sales`;
   - `<Client B>`: flag True → `GSTR3B (ClientB_Exports_Pvt_Ltd)`, `ITC (ClientB_Exports_Pvt_Ltd)`, `Sales (ClientB_Exports_Pvt_Ltd)` (the name cleaned by `sanitize_filename`: spaces → `_`, "Private Limited" → "Pvt_Ltd").
   RAN tree:
   ```
   ClientA Traders-MH/Version-211026 1030/  GSTR3B/  ITC/  Sales/
   ClientB Exports Pvt Ltd-GJ/Version-211026 1030/  GSTR3B (ClientB_Exports_Pvt_Ltd)/  ITC (ClientB_Exports_Pvt_Ltd)/  Sales (ClientB_Exports_Pvt_Ltd)/
   ```
   `<Client B>`'s ITC report Links B2 and B7 point into the named folders (RAN, see A-F14). No long-name warning appeared at Start, because the global tick was off (`processing_handler.py:90`).
4. The run's summary workbook says `Include Client Name in Folders | No` (RAN) — it reports only the global tick (`processing_handler.py:225`).

Edge run, script `scripts/t_edges_f6_f8_f10.py` (RAN, me): global tick ON then OFF:
```
after scan     : {'ClientA Traders-MH': False, ..., 'ClientB Exports Private Limited-GJ': False}  ['☐ No', '☐ No']
global ticked  : {'ClientA Traders-MH': True,  ..., 'ClientB Exports Private Limited-GJ': True}   ['☑ Yes', '☑ Yes']
global unticked: {'ClientA Traders-MH': True,  ..., 'ClientB Exports Private Limited-GJ': True}   ['☑ Yes', '☑ Yes']
ClientA Level-4 folders with the global tick now OFF: GSTR3B (ClientA_Traders) | ITC (ClientA_Traders) | Sales (ClientA_Traders)
```
The dictionary also still held the keys of clients from an **earlier** scan (`0042-MH`, `ClientC Mock-24`): `update_client_tree` sets flags for the current clients but never clears old ones (`client_handler.py:49`; RAN).

#### What the trace shows against what the feature claims

> **BUG (for Section 7): what the screen shows and what the folders get can disagree.** (a) Ticking the global box ON and then OFF leaves every per-client flag ON, so names are still added to all folders (RAN; `client_handler.py:222` "Don't change individual settings when global is turned OFF"). (b) After any scan the rows show "☐ No" while the global tick is still ON, and the run adds names to every folder (RAN GUI exp1 "globalname"). (c) The summary workbook says "Include Client Name in Folders: No" while per-client names were used (RAN). HIGH.

- The per-client warning uses a fixed 10 characters (`client_handler.py:191`), the Start warning the configured length with fallback 10 (`processing_handler.py:92-97`), and only for the global tick; a client switched on per row gets no Start warning (RAN, step 3). See A-F6.
- Unticking one client while the global tick is on switches the global tick off with the box "Global Setting Disabled", but that change is not saved (`client_handler.py:203-207`; READ).
- For an all-digit client name the toggle crashes at `len(values[0])` **after** the flag was flipped, so screen and setting disagree; for `0042` the flag is stored under the wrong key `42-MH` (RAN TRACE T7; see A-F8).
- The dictionary is handed live to the running job, so a click during Step 3 changes the folders of clients not yet processed (`processing_handler.py:193`; READ, MEDIUM-HIGH).
- The added name is the full cleaned name and is never shortened by the length setting (A-F6); it lengthens every file path inside these folders against Excel's ~218-character limit, which the hover help itself warns about (`gui/tabs/setup_tab.py:112-117`).

Rules involved: A-R39 (Level-4 names ⚠), A-R40 (decision), A-R136 (warnings), A-R137 (toggling behaviour). Not [GST RULE].

#### Older documents against the code

`PROJECT_AUDIT_ANNUALSTATEMENT.md:156, 702` (F2.4, DD9) describe the option as working; at the time of that audit (March 2026) per-client flags were silently ignored because the GUI stored them under the state **name** while the organiser looked them up under the state **code** — fixed in `33d0baf` (23-07-2026). The hover help example "'ITC (ABC Pvt Ltd)'" (`gui/tabs/setup_tab.py:113-114`) shows spaces; real folders use underscores (`ITC (ClientB_Exports_Pvt_Ltd)`, RAN).

#### History

From `edac6a3`; the configurable Start warning from `6237597` (both shared - before the split on 30-10-2025). A only (after the split): `33d0baf` (key by state code), `7347383` (short folder names `GSTR3B/ITC/Sales (…)` instead of "GSTR-3B Exports (…)", "Other ITC related files (…)", "Sales related files (…)").

### A-F11: Dry run (preview)

#### What it does

"🧪 DRY RUN" on Step 2 switches to Step 3 and writes into its log, for each ticked client, the files that would be organised and the two reports that would be created — without creating or copying anything.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Button (`.bind('<Button-1>')`) | `gui/tabs/validation_tab.py:109-118` |
| `dry_run()`: processing gate, selection, switch to Step 3, log lines | `gui/handlers/processing_handler.py:38-67` |
| Gate used: both templates set and existing, target set | `gui/handlers/file_handler.py:132-151` |
| Log lines go through `_log_async` (message frozen at call time) | `gui/handlers/processing_handler.py:23-31` |

#### Input → Output

Ticked clients → log lines on Step 3; nothing on disk.

#### User interaction

"🧪 DRY RUN" button on Step 2.

#### Concrete trace

End-to-end run, section S4 (RAN, me), both clients ticked, mode Fresh:

1. `validate_processing_inputs()` (`processing_handler.py:40` → `file_handler.py:132-151`): both templates set and present, target set → go on.
2. `get_selected_clients()` (`:43`) → `['ClientA Traders-MH', 'ClientB Exports Private Limited-GJ']`.
3. `notebook.select(2)` (`:49`) → RAN: Step 3 selected.
4. Lines `:51-67`, each through `_log_async`. The Step 3 log then held (RAN; time stamps and empty lines left out):
   ```
   🧪 ════════ DRY RUN PREVIEW ════════
   🎯 Mode: fresh
   👥 Selected: 2 clients
   🏢 ClientA Traders - Maharashtra
     📄 Would organize: AnnualReport-ClientA Traders-Maharashtra-2025-26.xlsx
     📄 Would organize: GSTR-2B-Reco-ClientA Traders-Maharashtra-Apr25-Mar26.xlsx
     📄 Would organize: GSTR3B-ClientA Traders-Maharashtra-Apr.xlsx
     📄 Would organize: GSTR3B-ClientA Traders-Maharashtra-May.xlsx
     📄 Would organize: ImsReco-ClientA Traders-Maharashtra-31032026.xlsx
     📄 Would organize: Sales-ClientA Traders-Maharashtra-Apr-Mar.xlsx
     📄 Would organize: SalesReco-ClientA Traders-Maharashtra-FY 2025-26.xlsx
     📊 Would create: ITC Report
     💰 Would create: Sales Report
   🏢 ClientB Exports Private Limited - Gujarat
     📄 Would organize: GSTR-2B-Reco-ClientB Exports Private Limited-Gujarat-Q1.xlsx
     📄 Would organize: GSTR3B-ClientB Exports Private Limited-Gujarat-Apr.xlsx
     📊 Would create: ITC Report
     💰 Would create: Sales Report
   🧪 ════════ DRY RUN COMPLETE ════════
   💡 No files were actually moved or created
   ```
5. Nothing was created under the target (RAN: the target folder did not even exist yet).

#### What the trace shows against what the feature claims

> **BUG (for Section 7): the preview says "Would create: Sales Report" for `<Client B>`, who has no Sales, Sales Reco or Annual Report file.** The real run does create that report, but it is an unchanged copy of the template with none of its Links cells filled (RAN, see A-F15). The preview gives no hint of that, nor of the destination folders, the effect of the mode (Fresh/Re-Run/Resume), the folder-name option, or which of several files of one type would feed a report. HIGH (RAN for the Sales report).

- Dry Run refuses to run unless both templates and the target are set and the templates exist, although it uses none of them (`processing_handler.py:40`; READ).
- There is no "already processing" check: a Dry Run during a real run interleaves its lines with the run's log (READ).
- After a run, the file names shown are the **cleaned** copy names, because the run renames the records in place (`core/file_organizer.py:233-246`; RAN: after run 1 every record of the end-to-end run held the cleaned name, e.g. `AnnualReport-ClientA_Traders-Maharashtra-2025-26.xlsx`) — so a second Dry Run would list names that are not in the source folder (HIGH, inference from the RAN data).
- A digit-only client key (`42-MH`, A-F8) raises KeyError at `:56` inside the click handler; the user sees nothing (READ).
- Each client heading starts with a line break after the time stamp (`:57`), so the time stamp stands alone on its line (RAN) — cosmetic.

Rules involved: A-R134 (gates), A-R139 (dry-run texts). Not [GST RULE].

#### Older documents against the code

`PROJECT_AUDIT_ANNUALSTATEMENT.md:257-259` (F6.2) reported that the dry-run log repeated the last client because of late binding — fixed in `0890022` (23-07-2026) exactly as that audit recommended. `README.md:158-159` (RM21) says "Step 3: Processing Tab — 1. Click 'Dry Run'" — never true: the button is on Step 2.

#### History

From `edac6a3`; `6237597` wrapped the lines in `root.after` (both shared - before the split on 30-10-2025); `0890022` (A only (after the split)) froze the message text per call.

### A-F12: Export selected client list to a text file

#### What it does

"📋 EXPORT LIST" on Step 2 asks where to save, writes the ticked clients into a text file with a short header, shows "Export Complete" and opens the file with Windows' default program.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Button (`.bind('<Button-1>')`) | `gui/tabs/validation_tab.py:121-130` |
| `export_client_list()` | `gui/handlers/client_handler.py:225-257` (dialog `:233-237`, write `:241-249`, boxes `:251-252, 256-257`, open `:255`) |

#### Input → Output

Ticked rows → `selected_clients_<YYYYmmdd_HHMMSS>.txt` (suggested name) containing "Selected Clients List", "Generated: <date time>", "Total: N clients", 50 "=" signs, a blank line, then one "<client> - <state>" line per client.

#### User interaction

"📋 EXPORT LIST" button, then the save dialog.

#### Concrete trace

End-to-end run, section S5 (RAN, me, Linux), both clients ticked:

1. `get_selected_clients()` (`:227`) → two keys.
2. Save dialog with suggested name `selected_clients_20261021_064056.txt` (`:233-237`; the recorder answered a scratch path).
3. `open(filename, 'w')` with no encoding (`:241`) and the lines `:242-249`. File content (RAN):
   ```
   Selected Clients List
   Generated: 2026-10-21 06:40:56
   Total: 2 clients
   ==================================================

   ClientA Traders - Maharashtra
   ClientB Exports Private Limited - Gujarat
   ```
4. Box "Export Complete" / "Exported 2 clients to: selected_clients.txt" (`:251-252`).
5. `os.startfile(filename)` (`:255`) does not exist on Linux → AttributeError → box **"Export Error" / "Failed to export: module 'os' has no attribute 'startfile'"** (`:256-257`), right after the "Export Complete" box (RAN).

#### What the trace shows against what the feature claims

> **BUG (for Section 7, cross-platform only): on any non-Windows system the user gets "Export Complete" followed by "Export Error"** although the file was written (RAN). `os.startfile` is Windows-only and not guarded by a platform check (`client_handler.py:255`), unlike Step 3's folder opening (`processing_handler.py:507-512`). LOW impact for an office that uses Windows; HIGH certainty.

- The file is written with the system's default encoding (on Windows usually the "ANSI" code page): a client name with characters outside it (for example Devanagari) would raise an error half-way, leaving a file with only the header lines and showing "Export Error" (READ, MEDIUM — not run on Windows).
- The list shows the client and state **as written in the file names**, not the client keys, files, status or missing types (READ) — informational only.
- A digit-only client key (A-F8) ends in "Failed to export: '42-MH'" (KeyError at `:248`; READ).
- The exported file contains client names and is written wherever the user chooses (confidentiality note only).

Rule involved: A-R142 (export format). Not [GST RULE].

#### Older documents against the code

`PROJECT_AUDIT_ANNUALSTATEMENT.md:354-358` (F13.1-F13.2): the feature description is right; "os.startfile … will crash on macOS/Linux" is not — it is caught and shown as "Export Error". Its recommendation K4 (platform check, doc line 910) was not done.

#### History

From `edac6a3` (shared - before the split on 30-10-2025); unchanged since.

### A-F13: Organise files into client and category folders (fresh / re-run / resume, backups)

#### What it does

When the user presses Start, a background worker (a daemon thread) handles the ticked clients one after another. For each client it creates the folder tree under the target folder — `Annual Statement-<DDMMYY HHMM>` / `<client>-<STATE CODE>` / `Version-<DDMMYY HHMM>` / `GSTR3B`, `ITC`, `Sales` — and **copies** (never moves) each of the client's source files into the right folder under a cleaned name. What happens when a file of that name is already there depends on the processing mode (Fresh / Re-Run / Resume). The reports (A-F14, A-F15), the text report and the summary workbook (A-F16) follow in the same run.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Mode radios (no `command`, so not saved) and their help texts | `gui/tabs/setup_tab.py:393-412`; help `:103-110`; values `utils/constants.py:174-190` |
| Start: re-entrancy guard, gate, selection, Start-time warning, buttons, daemon worker | `gui/handlers/processing_handler.py:69-134` |
| Worker: organiser created, per-client loop, folders, copies, log lines | `gui/handlers/processing_handler.py:182-284` (organiser `:189-195`, client lookup `:232`, per-client `try` from `:243`, folders `:253`, copies `:271-284`) |
| Organiser set-up (default length 35, run time stamp, mode check, target created) | `core/file_organizer.py:26-69` (`:42, 43, 62-63, 66-69`) |
| Folder tree for one client | `core/file_organizer.py:71-148` |
| Level-1 (run) folder: Fresh = new; Re-Run/Resume = newest existing `Annual Statement-*` folder by date modified; worked out once per run | `core/file_organizer.py:150-196` (`:184`) |
| Copy all files of a client; rename the in-memory record after a successful copy | `core/file_organizer.py:198-248` (`:233-246`) |
| One file: destination by type, cleaned name, existing-file handling by mode, copy | `core/file_organizer.py:250-300` (`:277-287`) |
| Which type goes into which folder | `core/file_organizer.py:302-315` |
| Name cleaning (`sanitize_filename`), copy with size check, backup | `utils/helpers.py:68-126, 172-215, 217-234` |
| Level-2 name, run time stamps | `utils/helpers.py:620-644, 25-35` |
| Folder name templates ⚠ (short names since `7347383`, with the measured reason in the comment) | `utils/constants.py:81-96` |

#### Input → Output

Ticked clients (`client_data` records) + target folder + mode → folders and copied files under the target; one result per file (Success / Failed / Skipped); the client records renamed in memory to the copies' names (used by the reports).

#### User interaction

"🚀 START" on Step 2 or "🚀 START PROCESSING" on Step 3 (both `.bind('<Button-1>')`); the mode radio on Step 1; "⏹ STOP" (takes effect between clients, A-F17).

#### Concrete trace

End-to-end run, sections S6 and S7 (RAN, me, real window, real worker thread inside Tk's main loop), both clients ticked, `<Client B>`'s folder-name flag on (A-F10), length 35.

**Run 1 — Fresh, folder stamp `211026 1030`:**

1. `start_processing()` (`processing_handler.py:69-134`): not already processing (`:75-76`); gate passed (`:78-79`); two keys (`:81`); global name tick off, so no Start warning (`:90`); Step 3 selected (`:116`); Start greyed, Stop active (`:119-120`); `is_processing = True` (`:124`); log "🚀 ════════ PROCESSING STARTED ════════" (`:126`); daemon thread started (`:129-134`). RAN: `is_processing = True | Start btn disabled | Stop btn normal | tab 2`.
2. Worker (`:182`): `timestamp = '211026 1030'` (`:186`); `FileOrganizer(target, 'fresh', include_client_name=False, client_folder_settings=<live dict>, client_name_max_length=35)` (`:189-195`) → `core/file_organizer.py:26-58`: stamp `'211026 1030'` (`:43`, a **second** clock reading), mode accepted (`:62-63`), target folder **created** because it did not exist (`:66-69`).
3. Client 1, `ClientA Traders-MH`: not stopped (`:229`); record fetched (`:232`); progress "Client 1 of 2: ClientA Traders | ETA: Calculating..." (`:235-241`).
4. `create_client_structure` (`core/file_organizer.py:71-148`): state code `MH` (`:82`); Level-2 name `ClientA Traders-MH` (`:83-87`); cleaned name `ClientA_Traders` (`:90`); Level 1 computed once (`:150-161` → `:165-173`, Fresh) = `Annual Statement-211026 1030`, created; Level 2 and `Version-211026 1030` created (`:97-108`); category folders `GSTR3B`, `ITC`, `Sales` (flag off → " (ClientA_Traders)" removed, `:117-139`).
5. Log: "✓ Created folder structure", "Level 1: Annual Statement-211026 1030", "Level 2: ClientA Traders-MH", "Level 3: Version-211026 1030" (`processing_handler.py:258-263`).
6. `organize_files` (`core/file_organizer.py:198-248`) took the files in the order their types were first met in the (sorted) scan: Annual Report, GSTR-2B Reco, GSTR-3B ×2, IMS Reco, Sales, Sales Reco. For each, `_organize_single_file` (`:250-300`):
   - destination by type ⚠ (`:302-315`): Annual Report → the Version folder itself; GSTR-2B Reco and IMS Reco → `ITC`; GSTR-3B Export → `GSTR3B`; Sales and Sales Reco → `Sales`;
   - cleaned name (`:271`, `utils/helpers.py:68-126`): `AnnualReport-ClientA Traders-Maharashtra-2025-26.xlsx` → `AnnualReport-ClientA_Traders-Maharashtra-2025-26.xlsx`; `SalesReco-… FY 2025-26.xlsx` → `…FY_2025-26.xlsx` (spaces → `_`);
   - destination did not exist, so no backup/skip branch (`:277-287`);
   - `safe_copy_file` (`utils/helpers.py:172-215`): source exists, folder made, `shutil.copy2` (content and dates), sizes equal → Success;
   - the record in `client_data` renamed to the copy's name (`:233-246`).
7. Log "✓ Organized 7 files", then one line per **successful** copy, e.g. `📁 GSTR-2B-Reco-ClientA Traders-Maharashtra-Apr25-Mar26.xlsx → GSTR-2B Reco folder` (`processing_handler.py:279-284`).
8. The two reports (A-F14, A-F15), the text report (A-F16), "✅ Completed ClientA Traders".
9. Client 2, `ClientB Exports Private Limited-GJ`: Level 2 `ClientB Exports Pvt Ltd-GJ` (abbreviated); category folders **with** the name (flag on); 2 files copied; the `Sales (ClientB_Exports_Pvt_Ltd)` folder created empty.
10. Summary workbook at Level 1 (A-F16); final log lines; box "Processing Complete" / "Successfully processed 2/2 clients … Files organized in: <base>/target" shown **from the worker thread** (`:498-502`; RAN: recorded thread `Thread-1 (process_files_thread)`); output folder opened (`:505-516`; on Linux `xdg-open "<base>/target/Annual Statement-211026 1030"`, recorded); `organize` hint (A-F5).

Target after run 1 (RAN; sizes omitted):
```
Annual Statement-211026 1030/
  ClientA Traders-MH/
    Version-211026 1030/
      AnnualReport-ClientA_Traders-Maharashtra-2025-26.xlsx
      GSTR3B/   GSTR3B-ClientA_Traders-Maharashtra-Apr.xlsx, GSTR3B-ClientA_Traders-Maharashtra-May.xlsx
      ITC/      GSTR-2B-Reco-ClientA_Traders-Maharashtra-Apr25-Mar26.xlsx, ImsReco-ClientA_Traders-Maharashtra-31032026.xlsx
      ITC_Report_ClientA_Traders_MH_211026_1030.xlsx
      Sales/    Sales-ClientA_Traders-Maharashtra-Apr-Mar.xlsx, SalesReco-ClientA_Traders-Maharashtra-FY_2025-26.xlsx
      Sales_Report_ClientA_Traders_MH_211026_1030.xlsx
      _Organization_Report_ClientA Traders_Maharashtra.txt
  ClientB Exports Pvt Ltd-GJ/
    Version-211026 1030/
      GSTR3B (ClientB_Exports_Pvt_Ltd)/   GSTR3B-ClientB_Exports_Pvt_Ltd-Gujarat-Apr.xlsx
      ITC (ClientB_Exports_Pvt_Ltd)/      GSTR-2B-Reco-ClientB_Exports_Pvt_Ltd-Gujarat-Q1.xlsx
      ITC_Report_ClientB_Exports_Pvt_Ltd_GJ_211026_1030.xlsx
      Sales (ClientB_Exports_Pvt_Ltd)/    (empty)
      Sales_Report_ClientB_Exports_Pvt_Ltd_GJ_211026_1030.xlsx
      _Organization_Report_ClientB Exports Private Limited_Gujarat.txt
  GST_Processing_Summary_211026_1030.xlsx
```

**Run 2 — Resume, folder stamp `211026 1045` (a quarter of an hour later):**

11. `_compute_level1_folder` (`core/file_organizer.py:175-186`): lists the target, keeps folders starting with "Annual Statement-", takes the newest by date modified → reuses `Annual Statement-211026 1030`.
12. But the Version folder is named after the **current** minute (`:102-108`): `Version-211026 1045` — new and empty. So for every file `dest_path.exists()` (`:277`) is False, the Resume branch "Skipped" (`:283-287`) never applies, and **all 7 + 2 files were copied again** ("✓ Organized 7 files", "✓ Organized 2 files"; RAN). A second summary `GST_Processing_Summary_211026_1045.xlsx` was written into the same Level-1 folder.
13. The run-2 log lists the **cleaned** names (e.g. `📁 AnnualReport-ClientA_Traders-…`), because run 1 renamed the records in place (`:233-246`).

**Originals (RAN, me, `scripts/t_f13_originals.py`):** the source folder was snapshotted (names, sha256, modification time to the nanosecond) before and after organising all clients: identical. A copy has the same bytes **and** the same date modified as its original (`shutil.copy2`).

#### What the trace shows against what the feature claims

> **BUG (for Section 7): "Resume" does not resume.** Its texts say "Continue previously interrupted processing" (`utils/constants.py:187`) and "Continue a run that was interrupted, skipping files that are already in place" (`gui/tabs/setup_tab.py:108-109`). Every run makes a new `Version-<current minute>` folder, so Resume copies everything again unless it is started in the **same minute** as the interrupted run (RAN, run 2; RAN TRACE T6 runs 4-5). HIGH.

> **BUG (for Section 7): runs in the same minute overwrite each other.** Fresh in the same minute reuses the same folders: an existing copy is first backed up as `<stem>_backup_<DDMMYY_HHMM><ext>` **in the same category folder** and then overwritten; the reports and the Level-1 summary are overwritten without backup, and the summary then lists only the last run's clients. Re-Run overwrites copies without any backup (no branch for it at `core/file_organizer.py:277-287`). Two backups in one minute overwrite each other. (RAN TRACE T6 runs 2-4, T8 §4.) HIGH.

> **BUG (for Section 7): two different "clients" can end up in one folder and overwrite each other's files.** The same client spelt "Private Limited" in one file name and "Pvt Ltd" in another gives two clients in Step 2 but one Level-2 folder and identical cleaned file names, so the second copy overwrites the first (RAN TRACE T8 §1). Names agreeing in their first 32 characters share one folder at the default length 35 (RAN TRACE T6 §7, T9), and on Windows names that differ only in capitals would share a folder too (MEDIUM, reasoned). HIGH.

- In Resume within the same minute, a skipped file keeps its original (uncleaned) name in memory, so the report's Links cell names a file that is on disk under its cleaned name (RAN CORE exp8, EXT t12; `core/file_organizer.py:234`). MEDIUM.
- "Newest Level-1 folder" means "most recently modified": writing any file into an older `Annual Statement-*` folder redirects Re-Run and Resume to it (RAN TRACE T3c). HIGH (Linux), MEDIUM (Windows folder-date behaviour).
- No path-length check exists before folders and files are created; commit `7347383` itself calls a path-length guard "the durable fix" (READ). Folder names alone take "Annual Statement-DDMMYY HHMM" (28 characters) and "Version-DDMMYY HHMM" (19).
- A copy broken by a network drop is left in place (`utils/helpers.py:213-215` returns False without deleting), and its file is not mentioned in the log because only successful copies are logged (`processing_handler.py:282-284`; READ, HIGH).
- If the backup fails, the existing file is overwritten anyway (`core/file_organizer.py:280-282`; READ).
- The log line `→ <type> folder` names the document **type** ("GSTR-2B Reco folder", "Annual Report folder"), not the folder the file went to (`ITC`, the Version folder) (`processing_handler.py:284`; RAN) — cosmetic, UNREGISTERED UI text (see return list).
- The per-client text report shows run-wide running totals: `<Client B>`'s report said "Total Files: 2 / Successfully Copied: 9 / Folders Created: 11" (RAN; `core/file_organizer.py:362-368`) — detail of A-F16.
- Folder and report names come from two separate clock readings (`processing_handler.py:186`, `core/file_organizer.py:43`), so a minute boundary between them gives different minutes (READ, LOW).
- The processing-mode radios have no `command`, so a changed mode is not saved (`gui/tabs/setup_tab.py:404-407`; READ).

Rules involved: A-R32 (Level-1 name), A-R33 (run time stamps), A-R35 (Level 2) [GST RULE], A-R38 (Version folder), A-R39 (Level-4 names ⚠), A-R40 (name decision), A-R41 (which type goes into which folder ⚠) [GST RULE], A-R42 (newest by date modified), A-R43 (processing modes), A-R45 (existing destination), A-R46 (backup name), A-R47 (copy method), A-R48 (copy names and in-memory rename), A-R111 (Level 1 worked out once). Dead: A-R37, A-R44, A-R49, A-R50.

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:152-154, 518-527` (F2.3, R13) and `README.md:185-187` (RM27): category folders "GSTR-3B Exports (…)", "Other ITC related files (…)", "Sales related files (…)" — no longer true since `7347383` (26-07-2026).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:559-562` (R17) describes Resume as "skip existing files" — PARTLY: correct line by line, but it cannot resume a run started in an earlier minute.
- `README.md:7, 42, 260` (RM3) "All file operations use temporary files and verification steps" — never true for the copies (straight `copy2` plus a size check).
- `README.md:45, 47, 306-307` (RM8) "Original files are never modified" — true (RAN).
- `README.md:178-190` (RM26-RM29): folder names right; the summary name is `GST_Processing_Summary_DDMMYY_HHMM.xlsx`, not `…DDMMYYYY_HHMMSS`; report names use the state **code**.
- Hover help for the target (`gui/tabs/setup_tab.py:83-85`) "Each run adds a new timestamped folder, so previous runs are left untouched" — true only for Fresh runs in different minutes.

#### History

From `edac6a3` (shared - before the split on 30-10-2025). A only (after the split): `275d9a4` (Level 1 worked out once per run), `33d0baf` and `4219ffb` (abbreviation rules in `sanitize_filename`), `7347383` (short category folder names), `0890022` (log messages frozen per call), `33d0baf` (re-entrancy guard on Start).

### A-F14: Create ITC report from template

#### What it does

For every processed client, right after its files are copied, the program makes an ITC report: it copies the ITC template to a local temporary file, writes into the template's `Links` sheet the folder paths and file names of the client's copied GSTR-3B, Annual Report, GSTR-2B Reco and IMS Reco files (through Excel by COM on Windows; with openpyxl elsewhere), and copies the result into the client's Version folder as `ITC_Report_<client>_<STATE CODE>_<DDMMYY_HHMM>.xlsx`. The template's Power Query queries are expected to read those Links cells when the report is later refreshed (Step 4).

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Report name, log lines, data preparation, call, success/failure counting | `gui/handlers/processing_handler.py:291-326` (name `:292-298`, template path re-read `:301, 310`) |
| Links values for the ITC report | `core/excel_handler.py:603-656` (`prepare_template_data`) |
| Links cell map ⚠ (ITC: B2, B4, B5, B7, B8, B10, B11) | `utils/constants.py:102-114` |
| Entry: template must exist (else raises), output folder made, COM or openpyxl | `core/excel_handler.py:66-102` |
| Choice made at load time: COM only if Windows **and** `win32com.client` and `pythoncom` import | `core/excel_handler.py:19-28` |
| COM path: attach to a running Excel or start one, hide it, stage locally, open, find/create `Links`, write, save, close, copy to destination, clean-up and Quit | `core/excel_handler.py:104-260` |
| openpyxl path: stage, load, find `Links` (exact name only), write, save, copy, clean-up | `core/excel_handler.py:262-322` |
| Open with one fallback (8.3 short path when the error mentions a path, or the path is > 218 characters) | `core/excel_handler.py:324-369`; `utils/helpers.py:580-590` |
| "Already open in this Excel?" check | `utils/excel_com.py:19-29` |
| Path form written into Links (forward slashes turned into backslashes), stem without extension | `utils/helpers.py:418-422, 553-555` |

#### Input → Output

ITC template + the client's folders and copied files → `ITC_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx` in the client's Version folder; a "✅ Success" or "❌ Failed" line in the Step 3 log; counters for the summary workbook.

#### User interaction

None: automatic for each client during Step 3.

#### Concrete trace — openpyxl path (RAN, me, Linux; end-to-end run S6)

For `<Client B>` (GSTR-2B Reco and GSTR-3B only; folder-name flag on):

1. Name (`processing_handler.py:292-298`): `sanitize_filename('ClientB Exports Private Limited')` → `ClientB_Exports_Pvt_Ltd`; state code from the record `GJ` (`:293`) [GST RULE] ⚠ state table; stamp `211026_1030` → `ITC_Report_ClientB_Exports_Pvt_Ltd_GJ_211026_1030.xlsx` in `Version-211026 1030`.
2. Log "📊 Creating ITC report...", "Template: ITC_Template.xlsx" (template path read again from the Step 1 box, `:301`), "Output: …" (`:300-302`).
3. `prepare_template_data(client_info, folders, 'ITC')` (`core/excel_handler.py:603-656`) [GST RULE] ⚠:
   - `state_code` computed and never used (`:616-617`);
   - `gstr3b_folder` = the GSTR-3B folder, **always** (`:622`) → `...\Version-211026 1030\GSTR3B (ClientB_Exports_Pvt_Ltd)`;
   - no Annual Report file → `annual_folder = ''`, `annual_filename = ''` (`:625-633`);
   - GSTR-2B Reco: folder = the ITC folder, file = stem of the **first** file of that type (`:636-641`) → `GSTR-2B-Reco-ClientB_Exports_Pvt_Ltd-Gujarat-Q1` (already the cleaned copy name, because A-F13 renamed the record);
   - no IMS Reco file → `ims_folder = ''`, `ims_filename = ''` (`:647-655`).
4. `create_report_from_template(<ITC template>, output, data, 'ITC')` (`processing_handler.py:309-314` → `core/excel_handler.py:66-102`): template exists (`:82-84`); output folder ensured (`:87`); `USE_WIN32COM` is False on Linux → log WARNING "Not on Windows - Power Query may not be preserved!" (`:98`) → `_create_report_openpyxl` (`:262-322`):
   - WARNING "Using openpyxl - Power Query content may be lost!" (`:267`);
   - staging file `<temp>/gst_excel_temp/ITC_Report_ClientB_…_211026_1030.xlsx` (`:273`), old one deleted (`:274-278`), template copied there (`:279`);
   - `load_workbook(..., data_only=False, keep_links=True)` (`:282`);
   - sheet `Links` found by exact name ignoring case (`:289-293`);
   - for B2, B4, B5, B7, B8, B10, B11 only **non-empty** values are written, as text (`:297-303`) → B2, B7, B8 written; B4, B5, B10, B11 skipped;
   - saved locally, copied to the destination, `True` because the file exists (`:306-310`); staging file deleted (`:315-322`; RAN: the staging folder was empty afterwards).
5. `True` → counters up, log "✅ Success: ITC_Report_ClientB_Exports_Pvt_Ltd_GJ_211026_1030.xlsx" (`processing_handler.py:315-319`).

The `Links` sheets read back (RAN; `\…\` = the target path, written with backslashes by `clean_windows_path`):

| Cell ⚠ | Meaning (from the map) | `<Client A>` report | `<Client B>` report |
|---|---|---|---|
| B2 | GSTR-3B folder | `\…\ClientA Traders-MH\Version-211026 1030\GSTR3B` | `\…\ClientB Exports Pvt Ltd-GJ\Version-211026 1030\GSTR3B (ClientB_Exports_Pvt_Ltd)` |
| B3 | (not in the map) | `OLD-B3` (template text) | `OLD-B3` |
| B4 | Annual Report folder (= Version folder) | `\…\ClientA Traders-MH\Version-211026 1030` | **`OLD-B4`** (template text kept) |
| B5 | Annual Report file stem | `AnnualReport-ClientA_Traders-Maharashtra-2025-26` | **`OLD-B5`** |
| B6 | (not in the map) | `OLD-B6` | `OLD-B6` |
| B7 | GSTR-2B Reco folder (= ITC folder) | `\…\Version-211026 1030\ITC` | `\…\Version-211026 1030\ITC (ClientB_Exports_Pvt_Ltd)` |
| B8 | GSTR-2B Reco file stem | `GSTR-2B-Reco-ClientA_Traders-Maharashtra-Apr25-Mar26` | `GSTR-2B-Reco-ClientB_Exports_Pvt_Ltd-Gujarat-Q1` |
| B9 | (not in the map) | `OLD-B9` | `OLD-B9` |
| B10 | IMS Reco folder (= ITC folder) | `\…\Version-211026 1030\ITC` | **`OLD-B10`** |
| B11 | IMS Reco file stem | `ImsReco-ClientA_Traders-Maharashtra-31032026` | **`OLD-B11`** |

#### Concrete trace — Excel through COM (READ; control flow RAN with fake COM objects)

Script `scripts/t_f14_fakecom.py` (RAN, me): fake `win32com`/`pythoncom` modules recorded every call while the real `core/excel_handler.py` ran as if on Windows; the same `<Client B>` data. **This shows only the order of the Python calls, not what real Excel does.** Recorded sequence when the user already has Excel open:

```
CoInitialize()
GetActiveObject('Excel.Application')                 # the USER'S running Excel is taken  (core/excel_handler.py:124)
USER'S Excel.Visible = False                         # user's Excel windows hidden          (:131)
USER'S Excel.DisplayAlerts = False                   # (:132)
USER'S Excel.ScreenUpdating = False                  # (:133)
Workbooks.Open('ITC_Report_ClientB_Exports_Pvt_Ltd_GJ_211026_1030.xlsx')   # the local staging copy (:159 -> :349)
Range(B2).Value = '…\GSTR3B'                         # only non-empty values (:189-197)
Range(B7).Value = '…\ITC'
Range(B8).Value = '…GSTR-2B-Reco-ClientB_Exports_Pvt_Ltd-Gujarat-Q1'
wb.Save()                                            # (:203)
wb.Close(SaveChanges=True)                           # (:206)
USER'S Excel.ScreenUpdating = True                   # finally (:239)
USER'S Excel.DisplayAlerts = True                    # finally (:240) - alerts back ON ...
USER'S Excel.Quit()                                  # ... then Quit, ALWAYS (:242-246)
CoUninitialize()                                     # (:250)
returned: True        report at destination: True   staging folder afterwards: []   Excel 'Visible' at the end: False
```

Other recorded scenarios:
- **No Excel running:** `GetActiveObject` fails → `Dispatch('Excel.Application')` (`:126-128`) → same sequence on the new Excel → returns True.
- **Template without a `Links` sheet** (sheets `Data`, `Info`): no sheet named `Links`, none containing "link" (`:166-178`) → log ERROR "Links sheet not found! Creating one..." → `Sheets.Add()` and rename to `Links` (`:180-183`) → the three values written into the **new, empty** sheet → saved → **returned True**.
- **Excel stops answering after the workbook was closed:** in `finally`, `excel.ScreenUpdating = True` (`:239`, not protected) raised "The RPC server is unavailable" → the exception **left the function**: no `Quit()`, no `CoUninitialize()`, staging file left in `gst_excel_temp`; the report **was already at the destination** (copied at `:215`). The caller would log "❌ ITC Report Error: …" and count a report error (`processing_handler.py:324-326`).

What real Excel would do here cannot be run on this machine (READ, with confidence):
- Taking the user's running Excel, hiding it (`Visible` is never set back) and then quitting it closes the user's own Excel; with `DisplayAlerts` switched back on just before `Quit()`, Excel may show a "Save changes?" prompt in a window that is now invisible, and the worker would wait on it (MEDIUM-HIGH; how to check: open Excel with an unsaved workbook, run Step 3 for one client, watch that window). The comment "Don't quit if Excel was already open" (`:241`) is not what happens: the test is `self.excel_app is None` (`:242`), and `self.excel_app` is set only once, to None (`:56`; READ, grep).
- `Dispatch` may itself attach to a running Excel (MEDIUM; general pywin32 behaviour).
- Each report is one Excel start/attach and one `Quit()`, so two per client, plus 1 s of fixed pauses per report (`:155, 210`).
- With an `.xltx` template, the staging copy is named `.xlsx`; whether Excel opens and saves it as a normal workbook is not known (LOW-MEDIUM). On the openpyxl path the output keeps the template content type inside an `.xlsx` name (RAN CORE exp5). How to check: run one client with an `.xltx` template on Windows and look for `template.main` in the report's `[Content_Types].xml`.

#### What the trace shows against what the feature claims

> **BUG (for Section 7): when the client has no file of a type, the template's existing text stays in those Links cells.** Empty values are simply not written (`core/excel_handler.py:192, 300`), so `<Client B>`'s report kept `OLD-B4`, `OLD-B5`, `OLD-B10`, `OLD-B11` (RAN). If the real templates hold sample paths there, the report's queries would point at another client's or a sample file. Whether the real templates contain such text is unknown — the templates are not in the repository (how to check: open each template's `Links` sheet). HIGH (mechanism).

> **BUG (for Section 7): a template without a `Links` sheet still gives "✅ Success".** The COM path adds a new empty `Links` sheet and writes into it (RAN with fake COM); the openpyxl path writes nothing at all and returns True (RAN CORE exp5, EXT t6). The hint "Check if template has 'Links' sheet" (`processing_handler.py:323`) can therefore never appear for that cause. HIGH.

> **BUG (for Section 7): Step 3 takes over, hides and closes the user's own Excel** (control flow RAN with fake COM; real effect MEDIUM-HIGH, see above), and **an Excel that stops answering skips the clean-up** (RAN with fake COM): Excel and the staging file are left behind, and a report that exists is counted as an error. HIGH (code).

> **BUG (for Section 7): on the openpyxl path the report silently loses Power Query.** A template with a DataMashup part came out without it, and the function returned True; the GUI shows "✅ Success" (RAN XL E1). On Windows this path is reached when pywin32 is present but cannot load (A-F1, RAN XL E2 simulated). HIGH (mechanism), LOW-MEDIUM (likelihood on Windows).

- **Only the first file of a type is linked** [GST RULE]: `files[0]` in path-sorted order (`core/excel_handler.py:628-630, 639-641, 650-652`; `utils/helpers.py:361`). A `… (1).xlsx` copy sorts before `….xlsx`, a `DDMMYYYY` date sorts by day, `2023-24` before `2024-25` (RAN TRACE T3b). The other files of that type are copied but not linked; nothing tells the user. GSTR-3B is linked by folder only (B2), so every file in that folder — including any `_backup_` copy from a same-minute Fresh run (A-F13) — is in the linked folder; whether the template reads the whole folder is unknown (how to check: open the template's Power Query source step). Described only.
- Folder paths are written exactly as the target was chosen (mapped drive letter or `\\server\share`), so a report refreshed on another PC needs the same path (READ, MEDIUM — depends on how the queries use the cells). File names are written **without** extension, so the template must supply it; `.xls` inputs are linked the same way (READ).
- The output overwrites an existing file of the same name without warning (`:215, 308`); a copy broken by a network drop is left at the destination (`:215` raises → `:226-228` returns False; no clean-up) and Step 4 would later find it by its name (READ, HIGH).
- `warnings.filterwarnings('ignore')` at load time (`core/excel_handler.py:17`) switches off all Python warnings for the whole program (READ).
- Every written Links value (first 50 characters, which include client names and paths) is logged at INFO (`:195`) — confidentiality note, A-F30.
- The template-missing case raises `FileNotFoundError` (`:84`) instead of returning False; for ITC it shows as "❌ ITC Report Error: Template not found: …" (`processing_handler.py:324-326`; READ).

Facts against the author's Standard 12 "Excel through COM" (`rebuild_pipeline_FINAL.md:817-898`; no recommendation, detail in Section 9): uses `GetActiveObject` then `Dispatch`, never `DispatchEx` (`:123-128`); `DisplayAlerts` set back to True before `Quit()` (`:240-244`); no `IgnoreReadOnlyRecommended` and no read-only check on open (`:349, 364`); `CoUninitialize()` runs in the same function while Excel objects are still held (`:250`); the openpyxl fallback saves a template that may contain Power Query with only a log warning, without checking for DataMashup (`:262-322`).

Rules involved: A-R54 (ITC Links cells ⚠ template cell positions) [GST RULE]; A-R56 (form of the values) [GST RULE]; A-R57 (first file of a type) [GST RULE]; A-R58 (finding the Links sheet); A-R59 (report names) [GST RULE]; A-R60 (order and success counting) [GST RULE]; A-R96 to A-R101 (COM or openpyxl, the COM session, staging, success test, already-open check, open fallback); A-R104 (218-character threshold); A-R106 to A-R108; A-R112 (fixed pauses). Dead: A-R63, A-R107.

#### Older documents against the code

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:220-221` (F4.5) "copies template to output location" — no longer true since `12d505f` (local staging).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:222` (F4.7) mentions the "contains link" fallback but not that the COM path **creates** a sheet and reports success.
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:231-232` (F4.10, F4.11): B2 path now ends `\GSTR3B`; the B5 example with a state code inside the file stem was never true (the stem keeps the state **name** from the original file name).
- `PROJECT_AUDIT_ANNUALSTATEMENT.md:661` (G12) noted `GetActiveObject` but not the always-`Quit()`.
- `README.md:188-189, 202-203` (RM28) report name with `DDMMYYYY_HHMMSS` and "State" — never true (`DDMMYY_HHMM`, state code). `README.md:208-227` (RM32) cell map — still true. `README.md:44, 262` (RM7) "No VBA macros or external links preserved" — never true (`keep_links=True`, `:282`; COM keeps everything). `README.md:278` (RM36) "Template not found → Ensure templates have Links sheet" — mismatched advice. `README.md:301, 314` "Close all Excel files" — not checked or enforced.
- Code comment `core/excel_handler.py:143-146` "Saving locally then copying yields a byte-identical file" — plausible for the copy step (HIGH); whether every query uses only the absolute Links paths is not verifiable without the templates (LOW).

#### History

The COM session (incl. `GetActiveObject`, `Visible = False`, the always-`Quit()` clean-up), the openpyxl fallback, the Links map and `prepare_template_data` are from `edac6a3` (shared - before the split on 30-10-2025). A only (after the split): `12d505f` (23-07-2026, stage on local disk, copy to the network once), `064c634` (debug lines downgraded to `logger.debug`), `ba2a753` (shared "already open" helper `utils/excel_com.py`).

### A-F15: Create Sales report from template

#### What it does

Right after the ITC report, the same mechanism (A-F14) makes a Sales report from the Sales template: it writes the folder paths and file names of the client's copied Sales, Annual Report and Sales Reco files into the template's `Links` sheet and saves the result as `Sales_Report_<client>_<STATE CODE>_<DDMMYY_HHMM>.xlsx` in the Version folder. The Sales part of the Step 3 code also prints 14 extra "[DEBUG]" lines into the user's log.

#### Where it lives

| What | File and lines (READ) |
|---|---|
| Sales part of the worker: template path read again, empty/missing checks, name, 14 "[DEBUG]" log lines, data, call, counting, extra existence checks, full traceback to the log file | `gui/handlers/processing_handler.py:328-411` ("[DEBUG]" lines at `:333, 342, 353, 362, 363, 367, 371, 374, 391, 393, 398, 403, 404, 411`) |
| Links values for the Sales report | `core/excel_handler.py:657-708` |
| Links cell map ⚠ (Sales: B2, B3, B5, B6, B8, B9) | `utils/constants.py:115-125` |
| Report creation (shared with A-F14) | `core/excel_handler.py:66-369` |

#### Input → Output

Sales template + the client's folders and copied files → `Sales_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx` in the Version folder; log lines; counters.

#### User interaction

None: automatic for each client during Step 3, after the ITC report.

#### Concrete trace (RAN, me, openpyxl path, end-to-end run S6; Excel-through-COM part as in A-F14, READ)

1. "💰 Starting Sales report creation..." (`processing_handler.py:329`); the Sales template path is read again from the Step 1 box (`:332`) and printed: `[DEBUG] Sales template path: <base>/templates/Sales_Template.xlsx` (`:333`).
2. Path not empty (`:335`); file exists → `[DEBUG] Template file exists: True` (`:341-342`); name `Sales_Report_ClientB_Exports_Pvt_Ltd_GJ_211026_1030.xlsx` (`:355-356`); "Template: …", "Output: …", `[DEBUG] Full output path: …`, `[DEBUG] Output directory exists: True` (`:358-363`).
3. `prepare_template_data(..., 'Sales')` (`:368-370` → `core/excel_handler.py:657-708`) [GST RULE] ⚠:
   - `<Client A>`: Sales folder + stem `Sales-ClientA_Traders-Maharashtra-Apr-Mar`; Version folder + stem `AnnualReport-ClientA_Traders-Maharashtra-2025-26`; Sales folder + stem `SalesReco-ClientA_Traders-Maharashtra-FY_2025-26`;
   - `<Client B>`: **no** Sales, Annual Report or Sales Reco file → all six values `''` (`:671-673, 687-689, 703-705`).
   `[DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']` (`:371`).
4. `create_report_from_template(sales_template_path, …, 'Sales')` (`:376-381`) → openpyxl path as in A-F14 → for `<Client B>` **no cell is written** (every value empty, `core/excel_handler.py:300`) → the template copy is saved unchanged → True.
5. "✅ Success: Sales_Report_ClientB_Exports_Pvt_Ltd_GJ_211026_1030.xlsx" (`:386`), then a second existence check of the output on the share and `[DEBUG] Output file created, size: 5462 bytes` (`:389-391`).

`Links` sheets read back (RAN):

| Cell ⚠ | Meaning (from the map) | `<Client A>` report | `<Client B>` report |
|---|---|---|---|
| B2 | Sales folder | `\…\ClientA Traders-MH\Version-211026 1030\Sales` | **`OLD-B2`** (template text) |
| B3 | Sales file stem | `Sales-ClientA_Traders-Maharashtra-Apr-Mar` | **`OLD-B3`** |
| B4 | (not in the map) | `OLD-B4` | `OLD-B4` |
| B5 | Annual Report folder (= Version folder) | `\…\ClientA Traders-MH\Version-211026 1030` | **`OLD-B5`** |
| B6 | Annual Report file stem | `AnnualReport-ClientA_Traders-Maharashtra-2025-26` | **`OLD-B6`** |
| B7 | (not in the map) | `OLD-B7` | `OLD-B7` |
| B8 | Sales Reco folder (= Sales folder) | `\…\Version-211026 1030\Sales` | **`OLD-B8`** |
| B9 | Sales Reco file stem | `SalesReco-ClientA_Traders-Maharashtra-FY_2025-26` | **`OLD-B9`** |

The summary workbook then said "Sales Reports Created: 2", "Report Creation Errors: 0", and both clients "Success" (RAN). In run 2 the Step 4 scan listed `<Client B>` with `has_sales = True` and its Sales tick **pre-ticked** (RAN, end-to-end S7), so a Step 4 refresh would open and refresh this unfilled template copy.

#### What the trace shows against what the feature claims

> **BUG (for Section 7): a client with no Sales-type files still gets a "successful" Sales report that is an unchanged copy of the template.** None of its Links cells are filled; whatever the template already holds in B2-B9 stays (RAN: `OLD-…` everywhere). It is logged "✅ Success", counted in "Sales Reports Created", and offered pre-ticked in Step 4 (RAN). Every client always gets both report attempts, whatever files it has (A-R60). Whether the real Sales template's cells hold sample paths cannot be checked here (templates not in the repository). HIGH (mechanism).

> **BUG (for Section 7): 14 "[DEBUG]" lines appear in the user's Step 3 log for every client** (RAN), e.g. full output paths with client names. The message of commit `064c634` says they "are intentionally left - they touch nested control flow and removing them is a separate UX decision". HIGH (fact), decision for the author.

- The ITC and Sales code paths are not symmetrical: the Sales part checks again for an empty or missing template (`processing_handler.py:335-346`) and writes a full traceback to the log file (`:405-406`); the ITC part relies on `create_report_from_template` raising (A-F14) (READ).
- The template path is re-read for every client, so changing the Sales template in Step 1 during a run changes it for the later clients (RAN GUI exp8: the log showed the other template's name for client 2) (`:332`). HIGH.
- Extra existence checks of the output folder and the new file (`:363, 389-393`) are extra network round-trips per client (READ).
- All other defects of the shared mechanism apply (A-F14): user's Excel taken over and closed, missing `Links` sheet reported as success, openpyxl path drops Power Query, only the first file of a type linked [GST RULE], output overwritten without warning.

Rules involved: A-R55 (Sales Links cells ⚠ template cell positions) [GST RULE]; A-R56, A-R57, A-R59, A-R60 [GST RULE]; A-R58; A-R96 to A-R101 (shared with A-F14).

#### Older documents against the code

`PROJECT_AUDIT_ANNUALSTATEMENT.md:194-215` (F4.3) and `README.md:208-227` (RM32) list the Sales cells B2, B3, B5, B6, B8, B9 correctly. Neither says that a client without Sales-type files gets an unfilled report. `PROJECT_AUDIT_ANNUALSTATEMENT.md:920` (K9) recommended removing the "[DEBUG]" logging — done in `core/` (`064c634`), not in these 14 user-visible lines.

#### History

From `edac6a3` (shared - before the split on 30-10-2025); the 14 "[DEBUG]" lines were already present at the split (RAN `git show … | grep -c`: 14 at `47a1492`, 14 at `bf2f1dd`). A only (after the split): `12d505f` (staging, shared mechanism), `064c634` (left these lines on purpose).

### Findings from the traces in this part, for Section 7

Each item below was shown by a concrete trace (RAN unless marked) and contradicts what the feature, its label or its help text claims. They are listed here so the Section 7 writer can number them (A-G…); the evidence is in the feature subsections above.

| # | Feature | Finding (short) | Evidence |
|---|---|---|---|
| 1 | A-F1 | No console → program dies at `main.py:17` with no message and no log | RAN scenario 4 |
| 2 | A-F1 | Install box names one package; the command installs all of `requirements.txt` (incl. pytest); manual command mixes both, unquoted | RAN scenario 3; READ `utils/dependency_installer.py:142` |
| 3 | A-F1 | Off Windows the install offer returns at every start; `find_spec` passes a pywin32 that cannot load | READ + RAN START; RAN XL E2 (simulated) |
| 4 | A-F2 | Dark Mode tick off-window and status bar hidden at the 1000×700 minimum size | RAN |
| 5 | A-F2 | Closing during Step 3 kills the worker after 1 s; closing during Step 4 has no warning | RAN (1.04 s); RAN GUI exp4; READ `gui/main_window.py:346` |
| 6 | A-F2, A-F7 | Status-bar messages ("Welcome…", "Found N clients") overwritten at once by the tab-change text | RAN |
| 7 | A-F3, A-F6 | Typed paths and the processing mode are not saved; a non-numeric length box makes every save fail and every run end with "Fatal error" | RAN |
| 8 | A-F4 | Remembered existence answers go stale; a file passes as a folder and a folder as a template; "Ready" does not gate Step 3 | RAN |
| 9 | A-F5 | "Full pipeline" does not start the refresh; "organise … no Excel automation" is untrue on Windows; hand-over runs after Stop and overwrites a Step 4 folder | RAN + READ |
| 10 | A-F6 | The length setting shortens only the Level-2 folder; ≤ 8 drops the state code; no Start warning for per-client names | RAN |
| 11 | A-F6 / A-F17 | After a fatal error at the start of a run the progress area still shows the previous run's "100% … Processing complete" | RAN |
| 12 | A-F7 | Silent drops (small, non-signature, `.csv`, `.xlsb`, `.xltx`, sub-folder files); `.xlsm` never matches; hyphenated client names mis-filed; listing error looks like an empty folder; scan freezes the window | RAN (me + TRACE T1/T6/T8); READ |
| 13 | A-F7 | Re-scan during Step 3 empties the running job's data: whole run "Fatal error" or empty Variations sheet | RAN GUI exp8; READ `processing_handler.py:232` |
| 14 | A-F8 | Client keys rebuilt from table text; number-looking names/states → wrong key, crash, or whole-run abort | RAN |
| 15 | A-F8 | "Clients with Missing Files" includes duplicates-only clients; Step 2/3 buttons not usable from the keyboard | RAN TRACE T6; READ |
| 16 | A-F9 | No on-screen list of unrecognised files; dropped files appear nowhere; blank-client files counted as parsed and as variations | RAN (me + TRACE T11) |
| 17 | A-F10 | Global tick OFF leaves per-client flags ON; table shows "☐ No" after a scan while the global tick adds names; summary says "No" while per-client names were used; stale keys kept across scans | RAN |
| 18 | A-F11 | Dry run promises a Sales report for a client with no Sales-type files; gated by templates it does not use; no "already processing" check | RAN + READ |
| 19 | A-F12 | "Export Complete" followed by "Export Error" off Windows; default (code-page) encoding | RAN; READ |
| 20 | A-F13 | Resume does not resume; same-minute runs overwrite (Re-Run without backup); two clients can share one folder and overwrite each other; no path-length guard; failed copies not logged | RAN (me + TRACE T6/T8) |
| 21 | A-F14, A-F15 | Empty values leave the template's own text in Links cells; a client with no Sales-type files gets an unfilled "successful" Sales report, pre-ticked in Step 4 | RAN |
| 22 | A-F14, A-F15 | Missing `Links` sheet reported as success (COM adds an empty sheet, openpyxl writes nothing) | RAN (fake COM; CORE exp5) |
| 23 | A-F14, A-F15 | The user's running Excel is taken, hidden and quit; an unresponsive Excel skips the clean-up and a created report is counted as an error | control flow RAN (fake COM); real effect READ, MEDIUM-HIGH |
| 24 | A-F14, A-F15 | openpyxl fallback drops Power Query and still reports success | RAN XL E1 |
| 25 | A-F15 | 14 "[DEBUG]" lines in the user's log per client | RAN |

