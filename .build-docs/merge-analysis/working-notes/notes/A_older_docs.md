# A-D — Project A's older documents checked against A's code today (snapshot bf2f1dd)

Working notes for the merge analysis. Reader: later agents who will write formal audits; final reader: the author (not a programmer).

## 0. How to read these notes

**Scope (task A-D).** Every line of these five documents in project A (`goyal-sidhant/AnnualStatement_oldv1`, HEAD = bf2f1dd, 27-07-2026) was read:

| Document | Lines | Written / last changed | Note |
|---|---|---|---|
| `PROJECT_AUDIT_ANNUALSTATEMENT.md` | 1,182 (READ in full) | fc279af (04-03-2026, "Project Audit", 856 lines) then f7eb8f1 (05-03-2026, "Project Audit 2.0", rewritten) | The file at HEAD is identical to the f7eb8f1 version (RAN: `diff` of `git show A/main:...` against A-copy, no output). It was written **before** A's 32 July-2026 commits. |
| `docs/superpowers/specs/2026-07-25-merge-extractor-design.md` | 149 (READ in full) | b276542 (25-07-2026) | A plan ("design spec") for joining the Power Query Extractor into the main window. |
| `tests/README.md` | 34 (READ in full) | 6c1e1c6 (23-07-2026) | Describes the test suite as first added. |
| `README.md` | 363 newline-terminated lines + 1 unterminated = 364 (READ in full) | d674386 (29-10-2025, a commit shared by A and B) | Identical in A and B (RAN: `diff <(git show A/main:README.md) <(git show B/main:README.md)` → no difference). |
| `.gitignore` | 80 lines (READ in full) | 1cb878b (29-10-2025, 79 lines) + 6c1e1c6 (added line 3 `.pytest_cache/`) | Differs from B's: both start from the 1cb878b file; A added `.pytest_cache/` (6c1e1c6), B added `#Claude` / `.claude` at the end (c258047) (RAN `git diff A/main B/main -- .gitignore`). |

The first audit version (fc279af, 856 lines) was also read in full (`git show fc279af:PROJECT_AUDIT_ANNUALSTATEMENT.md`, saved for reading as `/root/merge-scratch/run/A-D/audit_v1_fc279af.md`), and the f7eb8f1 diff was examined (it is a near-total rewrite: 1,003 lines added, 677 removed; only 165 unchanged lines survive, nearly all headings and blank lines).

To check the claims, every source file of A at HEAD was read (main.py, launch_extractor.py, requirements.txt, pytest.ini, .claude/settings.local.json, all of core/, utils/, gui/, power_query_extractor/), plus tests/conftest.py and the test names of the test files the documents mention.

**Labels used.**
- **READ** = established by reading code or git output. **RAN** = I executed something (scripts are in `/root/merge-scratch/run/A-D/`, outputs quoted in Section 9).
- **HIGH / MEDIUM / LOW** = confidence in an inference.
- Verdicts: **STILL TRUE** (true at the document's date and true at HEAD) / **NO LONGER TRUE** (true then, changed since — commit named) / **NEVER TRUE** (wrong even when written) / **CANNOT VERIFY** (needs Windows, Excel, real data or the author) / **PARTLY** (part right, part wrong — the split is spelled out).
- **[GST RULE]** = the code decides a tax figure, a return/table/sheet/cell, a state-code, a period etc. These are described, never judged. **TIME-VARYING** = a value that can change over time (table numbers, state lists, years).
- File references are repository-relative (`core/file_parser.py:123` = line 123 of that file at bf2f1dd unless a commit is named). "@47a1492" means "in that file as it stood at commit 47a1492", the split point where A and B part ways (30-10-2025); the March audit was written against that code.

**Plain-English glossary (used below).**
- *commit* = one saved change in the project history; *HEAD* = the latest commit; *diff* = the exact lines a commit added/removed.
- *regex (regular expression)* = a text pattern the program uses to recognise file names.
- *Excel COM / win32com / pywin32* = the Windows-only way a Python program remote-controls the real Microsoft Excel. *openpyxl* = a Python library that reads/writes .xlsx files without Excel (it cannot run Power Query).
- *Power Query* = Excel's built-in data-import feature; the report templates use it to pull figures from the client files.
- *Tk / Tkinter* = the window toolkit project A uses. *ttk style* = a shared look applied to many Tk widgets at once.
- *thread / worker thread* = a second line of work running in the background so the window does not freeze; Tk windows must only be touched from the main thread.
- *lambda / closure capture* = a small deferred function; if it is created in a loop without "freezing" the loop's current values, it later sees only the last values.
- *cache file* = the small settings file the app remembers between runs (`~/.gst_organizer_cache.json`, i.e. in the user's home folder).
- *mtime* = a file or folder's "last modified" time.
- *dead code* = code that exists but is never used.
- *xvfb* = a fake screen used here (Linux, no monitor) so Tk windows can be created in tests.

---

## 1. `PROJECT_AUDIT_ANNUALSTATEMENT.md` (the March-2026 audit, f7eb8f1 version = file at HEAD)

### 1.1 Header and Section 1 (Executive summary), doc lines 1-20

| # | Claim | Doc line | Code / git evidence | Verdict | Commit |
|---|---|---|---|---|---|
| H1 | Audit date 2026-03-05 | 3 | f7eb8f1 author date 05-03-2026 08:44 +0530 (READ `git log`) | STILL TRUE | f7eb8f1 |
| H2 | Project root `d:\OneDrive 2\OneDrive\Python 2026\AnnualStatement` | 5 | The author's own folder, not in the repo | CANNOT VERIFY | — |
| H3 | "v3.0 (Production Ready) — Tkinter GUI on main" | 6 | `utils/constants.py:134` title 'GST File Organizer v3.0'; `gui/utils/status_bar.py:29` "v3.0 \| Production Ready"; `main.py:2,6,63`; Tkinter throughout `gui/` (READ) | STILL TRUE — but the code has changed a great deal since without the version number changing (32 July commits) | — |
| H4 | "Active Branch: main at commit fc279af" | 7 | f7eb8f1's parent is fc279af (READ `git log --format=%P`) | STILL TRUE as a statement about that day | — |
| H5 | "Unmerged branch `copilot-worktree-2026-03-04T11-52-56` — PyQt5 migration, 3 commits ahead of main, NOT merged" | 8 | No branch of that name exists in either original checkout or in the combined repository (READ `git branch -a` / `for-each-ref`: only `main` and `claude/pensive-feynman-0mz1aa` in each). Its three commits c258047, a62d3cc, c21f788 are the first three of **project B's** five own commits (READ `git log B/main ^A/main`). | PARTLY: the commits are real and were never merged into A; the branch name CANNOT VERIFY; they became project B, not an abandoned branch | c258047, a62d3cc, c21f788 |
| S1.1 | Tool for CA/GST consultants; receives GSTR-2B reco, IMS reco, GSTR-3B exports, Sales, Sales reco, Annual Reports per client; organises, links to templates, refreshes Power Query, extracts values | 14 | Six file types `utils/constants.py:13-62`, `:68-75`; templates `core/excel_handler.py:66-102`; refresh/extract `power_query_extractor/core/report_processor.py` (READ) | STILL TRUE (now also inside the main window as Step 4, bfcb24a) | bfcb24a |
| S1.2 | Stack: Python 3.7+, Tkinter, openpyxl, win32com/pywin32 | 16 | `requirements.txt:2` "Python 3.7 or higher required"; imports as stated (READ). A Python-3.7-grammar parse of every non-test source file found no newer syntax (RAN, Section 9 exp-5; MEDIUM — grammar check only, no 3.7 interpreter available) | STILL TRUE as stated; actual 3.7 compatibility CANNOT VERIFY | — |
| S1.3 | Two entry points: `main.py`, `launch_extractor.py` | 16 | Both exist; `launch_extractor.py` unchanged since edac6a3 (READ `git log -- launch_extractor.py`) | STILL TRUE; additionally the extractor now also lives inside `main.py`'s window as tab 4 (`gui/main_window.py:146`, `gui/tabs/extract_tab.py`) | bfcb24a |
| S1.4 | Git history shows real-world fixes (circular imports, threading, long path handling) | 18 | 6237597 (imports, threading) exists. **Long-path handling was never a fix commit**: it is present from edac6a3 (READ `git show edac6a3:core/excel_handler.py` lines 304-309 and `report_processor.py` lines 652-727: 218-character checks, short paths, `\\?\` prefixes) | PARTLY (long-path part NEVER TRUE as "a fix in history") | 6237597 |
| S1.5 | PyQt5 migration "attempted (2026-03-04)" | 18 | c258047 author date **07-02-2026 11:27**, a62d3cc 07-02-2026 11:29, c21f788 07-02-2026 11:57 (READ). The first audit (fc279af line 17) had the right date 2026-02-07. | NEVER TRUE (date) | — |
| S1.6 | "branch was subsequently abandoned" | 18 | It continued as project B (two more commits 08-02 and 09-02-2026: 7e09e5c, b0e7d0e). Whether the author "abandoned" it is intent. | CANNOT VERIFY (intent); factually it is project B | — |
| S1.7 | Most recent commit on main fc279af added the first audit | 18 | True on 05-03-2026 (READ) | STILL TRUE as history | fc279af |
| S1.8 | "Functional but with significant issues" | 20 | Opinion | CANNOT VERIFY | — |
| S1.9 | PQ Extractor imports `pythoncom`/`win32com.client` at module level without fallback → crashes on import off Windows | 20 | @47a1492 `report_processor.py:11-15` plain imports (READ). At HEAD guarded: `power_query_extractor/core/report_processor.py:20-36` (try/except ImportError → `WIN32_AVAILABLE=False`), `process_client` refuses cleanly at `:66-77`; `gui/tabs/extract_tab.py:24-30` also guards the panel import (READ). RAN: whole test suite imports the extractor on Linux without pywin32 and passes (Section 9 exp-0). | NO LONGER TRUE | c463256 (26-07-2026) |
| S1.10 | main.py docstring says v3.0, UI says v3.0 — consistent | 20 | `main.py:2,6`; `constants.py:134`; `status_bar.py:29` | STILL TRUE | — |
| S1.11 | `requirements.txt` lists only openpyxl and lxml, omits pywin32 | 20 | Now `requirements.txt:9` `pywin32>=305; sys_platform == 'win32'` and `:15` `pytest>=7.0.0` (READ) | NO LONGER TRUE | 064c634 (pywin32), 6c1e1c6 (pytest) |

### 1.2 Section 2 (Project structure map), doc lines 24-112

| # | Claim | Doc line | Code / git evidence | Verdict | Commit |
|---|---|---|---|---|---|
| S2.1 | requirements.txt "INCOMPLETE - missing pywin32" | 30 | see S1.11 | NO LONGER TRUE | 064c634 |
| S2.2 | README.md "Comprehensive usage documentation (364 lines)" | 31 | 364 lines counting the unterminated last line (READ) | STILL TRUE (length); content outdated — see Section 4 of these notes | — |
| S2.3 | `.gitignore` "(79 lines)" | 33 | At f7eb8f1: 79 lines (RAN `git show 47a1492:.gitignore | wc -l` = 78 newlines + last line). Now 80 lines: `.pytest_cache/` inserted at `.gitignore:3` | NO LONGER TRUE | 6c1e1c6 |
| S2.4 | `.claude/settings.local.json` permits `python -m py_compile` via Bash | 35-36 | `.claude/settings.local.json` `"allow": ["Bash(python -m py_compile:*)"]` (READ) | STILL TRUE (still tracked in A; B stopped tracking it in c258047) | bd176bf (added) |
| S2.5 | `core/__init__.py` exports FileParser, FileOrganizer, ExcelHandler | 39 | `core/__init__.py:6-10` (READ) | STILL TRUE | — |
| S2.6 | file_parser / file_organizer / excel_handler descriptions | 40-42 | Match (READ) | STILL TRUE | — |
| S2.7 | `utils/__init__.py` "Package init (minimal, just docstring)" | 45 | It re-exports 9 constants and 22 helpers (`utils/__init__.py:6-36`); same @47a1492 (READ). The first audit (fc279af line 38) said correctly "re-exports everything". | NEVER TRUE (now also re-exports `scan_excel_files`, added 56aaad9) | — |
| S2.8 | constants.py / helpers.py descriptions | 46-47 | Match (READ) | STILL TRUE | — |
| S2.9 | `gui/__init__.py` exports GSTOrganizerApp | 50 | `gui/__init__.py:7` (READ) | STILL TRUE | — |
| S2.10 | `gui/handlers/__init__.py`, `gui/tabs/__init__.py`, `gui/utils/__init__.py`, `gui/widgets/__init__.py` "(empty docstring)" | 53, 58, 63, 69 | Each imports and lists its modules (`gui/handlers/__init__.py:4-8`, `gui/tabs/__init__.py:4-8`, `gui/utils/__init__.py:4-9`, `gui/widgets/__init__.py:4-7`); same @47a1492 (READ). At HEAD `gui/tabs/__init__.py` does **not** list the new `ExtractTab`, and `gui/widgets/__init__.py` does not list `Tooltip` (the code imports them directly, so nothing breaks). | NEVER TRUE | — |
| S2.11 | Module roles in gui/ (handlers, tabs Step 1-3, cache_manager at `~/.gst_organizer_cache.json`, dark_mode_manager, status_bar, ui_helpers, collapsible_frame, title_bar with dark-mode toggle) | 51-71 | Match (READ). But the map is now incomplete: missing `gui/tabs/extract_tab.py` (Step 4, bfcb24a), `gui/utils/setup_validation.py` and `gui/widgets/tooltip.py` (322717f) | NO LONGER complete | bfcb24a, 322717f |
| S2.12 | power_query_extractor tree: `__init__` "version only", extractor_main "entry point class", cell_mappings "editable by user", report_processor "Excel COM for PQ refresh", data_consolidator, extractor_window "Tkinter window" | 73-85 | `power_query_extractor/__init__.py:1-3` (docstring + `__version__`); `extractor_main.py:24` class `PowerQueryExtractor`; others match. extractor_window is no longer just a window: now `ExtractorPanel(ttk.Frame)` (`extractor_window.py:98`) + thin `PowerQueryExtractorApp(tk.Tk)` (`:1032`). Missing from map: `power_query_extractor/core/refresh_naming.py` (2aaa2b1) | PARTLY / NO LONGER complete | a677670, 2aaa2b1 |
| S2.13 | (Map omits) — files that exist at HEAD and are not in the map | — | `utils/excel_com.py` (ba2a753), `utils/requirements_check.py` (f363946), `utils/dependency_installer.py` (a5ad487), `pytest.ini` + `tests/` (6c1e1c6 onward), `docs/superpowers/specs/...` (b276542) (READ) | NO LONGER complete | as listed |
| S2.14 | Entry points table (`python main.py`, `python launch_extractor.py`) | 91-94 | `main.py:136-137`, `launch_extractor.py:29-30` (READ) | STILL TRUE | — |

**Dead-code table (doc lines 96-112):**

| # | Claim | Doc line | Evidence at HEAD | Verdict | Commit |
|---|---|---|---|---|---|
| D1 | `helpers.py:422 calculate_file_hash()` never called | 100 | Now `utils/helpers.py:529`; only other appearances are the re-export in `utils/__init__.py:19,35` (RAN `grep -rn calculate_file_hash`) | STILL TRUE (line moved) | — |
| D2 | `extract_filename_without_extension()` IS called from `excel_handler.py:576,588,598,614,631,646` | 101 | @47a1492 the calls are at 576, 587, 598, 614, 630, 646 (READ) — two off by one. At HEAD `core/excel_handler.py:628, 639, 650, 666, 682, 698` (READ) | STILL TRUE (line numbers moved) | — |
| D3 | `data_consolidator.py:250 _write_section()` never called | 102 | `power_query_extractor/core/data_consolidator.py:250`, no callers (RAN grep) | STILL TRUE | — |
| D4 | `data_consolidator.py:354 _create_details_sheet()` never called | 103 | `:354`, no callers (RAN grep) | STILL TRUE | — |
| D5 | `constants.py:180 TEMPLATE_EXTENSIONS` unreferenced | 104 | now `utils/constants.py:197`, unreferenced (RAN grep) | STILL TRUE | — |
| D6 | `constants.py:182 EXCEL_SIGNATURES` unreferenced; real check uses hard-coded bytes at `helpers.py:240-241` | 105 | now `utils/constants.py:199-202`; the real check uses `_XLSX_SIG`/`_XLS_SIG` at `utils/helpers.py:264-265` (READ) | STILL TRUE | — |
| D7 | `cache_manager.py:50-55 get_cached_value()/set_cached_value()` never called | 106 | `gui/utils/cache_manager.py:50-56`, no callers (RAN grep) | STILL TRUE | — |
| D8 | `parsing_rate` IS used, at `main_window.py:280` | 107 | computed `core/file_parser.py:333`, shown `gui/main_window.py:307` (READ) | STILL TRUE (line moved) | — |
| D9 | `main.py:72` error message references non-existent `gui/app.py` | 108 | @47a1492 the text is at `main.py:71` ("📄 app.py" under "📁 gui/") (READ). Removed when the ImportError handler was rewritten (`main.py:114-127` now prints the interpreter hint) (READ `git log -S'app.py' -- main.py` → edac6a3, f363946) | NO LONGER TRUE | f363946 (26-07-2026) |
| D10 | `helpers.py:15 # ADD THIS LINE` | 109 | now `utils/helpers.py:17` `import platform  # ADD THIS LINE` | STILL TRUE (line moved) | — |
| D11 | `file_parser.py:16 # ADD THIS` | 110 | gone from `core/file_parser.py:13-16` (READ; `git log -S'# ADD THIS' -- core/file_parser.py` → 56aaad9) | NO LONGER TRUE | 56aaad9 |
| D12 | `file_organizer.py:15 # ADD THESE` | 111 | `core/file_organizer.py:15` | STILL TRUE | — |
| D13 | `excel_handler.py:39 # ADD THESE TWO` | 112 | `core/excel_handler.py:39` | STILL TRUE | — |
| D14 | (omitted by the audit) other leftover AI-instruction comments | — | `report_processor.py:132` "# ADD THE NEW CODE HERE (notice the indentation - 3 levels):", `:183` "# ADD THIS:", `:295` "# Rest of your code remains same...", `data_consolidator.py:222` "# ADD THIS LINE:", `:379`, `:402` "# ADD THESE LINES", `:75`/`:80` "# ... rest of the code / header formatting" (READ) | audit incomplete (all existed @47a1492 — READ `git show 47a1492:...` grep, HIGH) | — |
| D15 | (omitted) other dead items at HEAD | — | `FileParser.validate_results()` `core/file_parser.py:338-371` (no callers); `get_unique_path()` `utils/helpers.py:451-470` (no callers, so the "safety limit 100" constant the audit lists at doc line 583 is never exercised); `get_relative_excel_path()` `:557-574`; `EXCEL_EXTENSIONS` `constants.py:196`; `UIHelpers.update_status` `gui/utils/ui_helpers.py:12-14`; `ReportProcessor.cleanup()` does nothing (`report_processor.py:801-803`); `find_excel_files()` now only used by tests (`utils/helpers.py:368-373`); `recent_folders` is written to the cache (`gui/utils/cache_manager.py:32-43`) but never read anywhere (RAN grep) | audit incomplete | — |

### 1.3 Section 3 (Feature inventory), doc lines 116-358

Note: the audit's examples use a real-looking company name (see Section 10). Below it is written `<Client A>`.

| # | Claim | Doc line | Code evidence at HEAD | Verdict | Commit |
|---|---|---|---|---|---|
| F1.1 | Scans folder, matches 6 regexes, extracts client & state, groups by client | 120 | `core/file_parser.py:136-177`, `:56-134`; `utils/constants.py:13-62` | STILL TRUE [GST RULE: classification of files into return types] | — |
| F1.2 | Lives in `core/file_parser.py`, FILE_PATTERNS lines 13-62 | 122 | `utils/constants.py:13-62` | STILL TRUE | — |
| F1.3 | Output: scanned_files, client_data keyed "ClientName-StateCode", variations | 125-126 | `core/file_parser.py:173, 229-230, 214-220` | STILL TRUE | — |
| F1.4 | User clicks "Scan Files" on Setup tab; progress in status bar | 128 | `gui/tabs/setup_tab.py:459-462` ("🔍  Scan Files", ttk button); progress callback → status bar `gui/handlers/file_handler.py:67-68` | STILL TRUE (button now enabled only when the source folder is valid, `setup_tab.py:585-591`) | 322717f |
| F1.5 | Trace step 1: `scan_folder()` calls `find_excel_files()` which globs `*.xlsx/*.xls/*.xlsm` | 131 | Now calls `scan_excel_files()` (`core/file_parser.py:158`), one `os.scandir` listing with a lower-case `.endswith(('.xlsx','.xls','.xlsm'))` filter (`utils/helpers.py:327-339`) | NO LONGER TRUE (same extensions; mechanism changed) | a603c02, 56aaad9 |
| F1.6 | Trace step 2: each file validated (extension, ≥1KB, ZIP/OLE signature) | 132 | Same checks, split into `_passes_metadata_checks` (`utils/helpers.py:290-306`, extensions .xlsx/.xls/.xlsm/.xltx/.xltm, size ≥1024) and `_has_excel_signature` (`:275-287`), signature reads run 16-at-a-time for ≥4 files (`:271-272, 351-357`) | STILL TRUE (behaviour); mechanism changed | 56aaad9, 4f84ba9 |
| F1.7 | Trace step 3: GSTR3B pattern matches `GSTR3B-<Client A>-Maharashtra-Apr24.xlsx` → client, state "Maharashtra" | 133 | RAN with generic name `GSTR3B-ABC Ltd-Maharashtra-Apr24.xlsx` → type 'GSTR-3B Export', client 'ABC Ltd', state 'Maharashtra' (Section 9 exp-1) | STILL TRUE | — |
| F1.8 | Trace step 4: "Maharashtra" → "MH", key "<Client A>-MH" | 134 | `core/file_parser.py:229-230`; `constants.py:266`; RAN key 'ABC Ltd-MH' (exp-2) | STILL TRUE [GST RULE: state-code logic] | — |
| F1.9 | Trace step 5: 3 of 6 types → "Missing 3 files" | 135 | `core/file_parser.py:259-283` | STILL TRUE [GST RULE: which documents a client "must" have] | — |
| F1.10 | Trace step 6: GSTR3B regex accepts a leading "(1) " | 136 | `constants.py:31`; RAN `(1) GSTR3B-...` parses (exp-1) | STILL TRUE | — |
| F1.11 | "Bug found": line 278 `if not missing and not extras`; extra unrecognised file still "Complete" because it goes to variations | 138 | `core/file_parser.py:278` | STILL TRUE (as described; the audit itself calls it arguably correct) | — |
| F2.1 | Creates the hierarchy; inputs: target, mode, client info | 142-147 | `core/file_organizer.py:71-148` | STILL TRUE | — |
| F2.2 | level1 `Annual Statement-DDMMYY HHMM/`, level2 `ClientName-StateCode/`, version `Version-DDMMYY HHMM/` | 149-151 | `constants.py:82-84`; timestamp format `'%d%m%y %H%M'` `utils/helpers.py:25-27`; RAN (exp-2: `Annual Statement-211026 0413/ABC Ltd-MH/Version-211026 0413`) | STILL TRUE | — |
| F2.3 | Category folders `GSTR-3B Exports (ClientName)`, `Other ITC related files (ClientName)`, `Sales related files (ClientName)` (or without the bracket) | 152-154 | Now `GSTR3B ({client})`, `ITC ({client})`, `Sales ({client})` (`utils/constants.py:91-95`); RAN folders `GSTR3B`, `ITC`, `Sales` (exp-2) | NO LONGER TRUE | 7347383 (26-07-2026) |
| F2.4 | Client name in brackets controlled by a checkbox | 156 | Global checkbox `gui/tabs/setup_tab.py:417-423` plus per-client toggle `gui/handlers/client_handler.py:175-209`; decision `core/file_organizer.py:124-134` | STILL TRUE. Note: per-client toggles were **silently ignored** until 33d0baf (key used state NAME vs CODE) — the audit did not notice | 33d0baf |
| F2.5 | Trace: `_get_level1_folder()` creates the timestamped folder; `create_client_state_key(..., max_length=35)` → "<Client A>-MH"; version folder; categories without name | 158-163 | `core/file_organizer.py:150-196` (Level 1 is now computed once per run and cached, `:45-47, 159-161`); `utils/helpers.py:620-644` | STILL TRUE except category names (F2.3) | 275d9a4, 7347383 |
| F3.1 | Copies and sanitises: "Private"→"Pvt", "Limited"→"Ltd", invalid characters removed | 167 | `core/file_organizer.py:271-272`; `utils/helpers.py:68-126` | STILL TRUE in outline (rules changed — see R-SAN) | 33d0baf, 4219ffb |
| F3.2 | `organize_files` lines 181-231 | 169 | now `core/file_organizer.py:198-248` | STILL TRUE (moved) | 275d9a4 |
| F3.3 | Results Success/Failed/Skipped | 173 | `core/file_organizer.py:254-300` | STILL TRUE | — |
| F3.4 | After copying, `client_info['files'][type][i]['name']` is replaced with the sanitised name, which the Excel step then writes into templates | 175 | `core/file_organizer.py:233-246`; RAN: annual filename written = `AnnualReport-ABC_Ltd-Maharashtra-2024` (exp-2) | STILL TRUE | — |
| F3.5 | Trace: `Sales-<Client A> Private Limited-Maharashtra-Apr-Jun.xlsx` → `Sales-<Client A>_Pvt_Ltd-Maharashtra-Apr-Jun.xlsx` | 177-179 | RAN with generic `ABC`: `Sales-ABC_Pvt_Ltd-Maharashtra-Apr-Jun.xlsx` (exp-1) | STILL TRUE (result) | — |
| F3.6 | Trace step 3: copied into `Sales related files/...` | 180 | folder is now `Sales/` (exp-2) | NO LONGER TRUE | 7347383 |
| F3.7 | Trace step 4: size compared; mismatch → destination deleted, False | 181 | `utils/helpers.py:202-208` | STILL TRUE | — |
| F4.1 | Copies template, writes folder paths/filenames into "Links" sheet; COM on Windows (keeps Power Query), openpyxl elsewhere (loses it) | 186 | `core/excel_handler.py:90-102`. New since be6418c: on Windows the app refuses to start without pywin32 (`utils/requirements_check.py:44-60` `essential=ON_WINDOWS`; `main.py:97-102`), so the openpyxl fallback is only reachable on Windows if someone bypasses main.py | STILL TRUE (code paths); context changed | be6418c |
| F4.2 | ExcelHandler lines 65-274 | 188 | now `core/excel_handler.py:46-713` | moved | — |
| F4.3 | ITC Links cells B2 gstr3b_folder, B4 annual_folder, B5 annual_filename, B7 gstr2b_folder, B8 gstr2b_filename, B10 ims_folder, B11 ims_filename; Sales B2, B3, B5, B6, B8, B9 | 194-215 | `utils/constants.py:102-126` (exact); values from `core/excel_handler.py:603-713`; B4/B5 = Version folder + annual file stem; B7 and B10 both = the ITC folder (READ) | STILL TRUE (template wiring; these are folder/file names, not tax figures) | — |
| F4.4 | COM step 1 `CoInitialize`; 2 `GetActiveObject` then `Dispatch` | 218-219 | `core/excel_handler.py:120, 123-128` | STILL TRUE | — |
| F4.5 | COM step 3 "Copies template to output location via shutil.copy2"; step 4 opens the copied file | 220-221 | Now copies the template to a **local** staging file `%TEMP%\gst_excel_temp\<output name>` (`core/excel_handler.py:147-154`, temp dir `:53`), opens, edits and saves that, then copies the finished file to the output path (`:212-215`), deleting the staging file in `finally` (`:254-260`). @47a1492 it copied the template straight to the output path and opened it there (READ 12d505f diff: removed lines `shutil.copy2(str(template_path), str(output_path))` and `wb = self._safe_open_workbook(excel, output_path)`). | NO LONGER TRUE | 12d505f (23-07-2026) |
| F4.6 | `_safe_open_workbook()` handles >218-char paths with 8.3 short-path fallback | 221 | `core/excel_handler.py:324-369` (threshold `:356`); "already open" detection moved to shared `utils/excel_com.py:19-29` | STILL TRUE (but it now opens a short local temp path, so the fallback rarely matters — MEDIUM) | ba2a753 |
| F4.7 | Finds "Links" case-insensitively, then any sheet containing "link" | 222 | `core/excel_handler.py:166-178` | STILL TRUE; not stated by audit here: if none is found the COM path **creates** a new "Links" sheet (`:180-183`) and reports success; the openpyxl path writes nothing and still returns True (`:295-310`; RAN exp-2) | — |
| F4.8 | Writes `ws.Range(cell).Value = str(value)`; saves, closes; `CoUninitialize` in finally | 223-225 | `core/excel_handler.py:189-206, 248-252`; only non-empty values are written (`:192`) | STILL TRUE | — |
| F4.9 | Trace: template copied to `Version-.../ITC_Report_<Client A>_MH_050326_1430.xlsx` | 228 | Name built `gui/handlers/processing_handler.py:292-298` = `ITC_Report_{sanitised client}_{state code}_{DDMMYY_HHMM}.xlsx` | STILL TRUE (format) — but it is now built locally first (F4.5) | 12d505f |
| F4.10 | Trace: B2 = `...\Version-050326 1430\GSTR-3B Exports` | 231 | now `...\GSTR3B` (RAN exp-2) | NO LONGER TRUE | 7347383 |
| F4.11 | Trace: B5 = `AnnualReport-<Client A>-MH-2024` | 232 | The stem keeps the **state name** from the original file name (sanitised), e.g. `AnnualReport-ABC_Ltd-Maharashtra-2024` (RAN exp-2); nothing converts the state to a code inside file names (`utils/helpers.py:68-126`) | NEVER TRUE | — |
| F5.1 | Summary workbook with 5 sheets Summary, Client Status, File Mapping, Errors, Variations at Level 1 | 237 | `core/excel_handler.py:371-409`, sheets `:391-395`; saved at Level 1 `processing_handler.py:462-469` | STILL TRUE | — |
| F5.2 | method lines 319-357 | 239 | now `:371-409` | moved | — |
| F5.3 | Output name `GST_Processing_Summary_DDMMYY_HHMM.xlsx` | 243 | `processing_handler.py:186, 462-464` (timestamp `DDMMYY HHMM` with space→underscore) | STILL TRUE | — |
| F5.4 | Automatic at end of processing | 245 | `processing_handler.py:455-475` (only if ≥1 client succeeded, `:456`) | STILL TRUE | — |
| F6.1 | Dry run shows what would happen | 249-255 | `gui/handlers/processing_handler.py:38-67`; button on Validation tab `gui/tabs/validation_tab.py:109-118` | STILL TRUE | — |
| F6.2 | "Bug: lines 41-48 lambda captures client_info/file_info by reference → dry-run log repeats the last client/file" (and same for processing log) | 257-259 | @47a1492/6237597 the lambdas were `root.after(0, lambda: ...)` without frozen values (READ 6237597 diff). Now every such call goes through `_log_async`/`_progress_async`, which freeze the message with default arguments (`processing_handler.py:23-36`; dry run `:51-67`) | NO LONGER TRUE (fixed exactly as the audit recommended) | 0890022 (23-07-2026) |
| F7.1 | Tree with checkboxes, keyboard (arrows + space), Select All / Clear All / Complete Only, per-client folder-name column | 263 | `gui/tabs/validation_tab.py:144-247`; `gui/handlers/client_handler.py:18-223` | STILL TRUE | — |
| F7.2 | Click column #0 toggles; FolderName column is `#7`; space toggles | 271 | `client_handler.py:69-86` (region 'tree' and column '#7') | STILL TRUE | — |
| F7.3 | Trace: `select_complete_clients()` checks Status (index 2) for "Complete"; status "Selected 6 complete clients" | 273-277 | `client_handler.py:147-159` | STILL TRUE | — |
| F8.1 | Full pipeline: folders, copies, ITC & Sales reports, organisation report, master summary | 281 | `processing_handler.py:182-530` | STILL TRUE; plus, in 'full' workflow, it now hands over to Step 4 (`:519`, `:136-175`) | 510f7b4 |
| F8.2 | `process_files_thread` lines 119-463 | 283 | now `:182-530` | moved | — |
| F8.3 | Start button, daemon thread, Stop button, progress & log via `root.after(0, ...)` | 289 | thread `:129-134`; stop flag checked once per client `:229`; Stop only takes effect between clients | STILL TRUE | — |
| F8.4 | "All UI updates from the processing thread use `self.app.root.after(0, lambda: ...)`" (added in 6237597) | 291 | @47a1492 only 24 calls were wrapped; the worker also called `self.app.log_message(...)` directly 45 times (RAN count from line 119 on) (Sales section, summary, completion), `messagebox.showinfo/showerror` and `start_btn.config/stop_btn.config` (READ `ph_47a1492.py` lines 256-463). At HEAD the same direct calls remain: `processing_handler.py:319-516` (`log_message`), `:459, :482` (`update_progress`), `:498-502` (`messagebox.showinfo` from the worker), `:525` (`messagebox.showerror`), `:528-530` (button config in `finally`). `log_message` itself schedules via `widget.after` (`gui/utils/ui_helpers.py:22-25`), so most lines still reach the screen; calling Tk from a worker thread is still the pattern the design note calls unsafe (MEDIUM — exact Windows behaviour CANNOT VERIFY) | NEVER TRUE | — |
| F9.1 | Dark mode: walks all widgets with `winfo_children()`, stores original colours on first run, applies `#2b2b2b`/`#ffffff`/`#3b3b3b`, or restores originals | 303 | @47a1492 (unchanged until July): stored only ttk style colours and the root background on first run (`dm_b276542.py:27-63`); on restore it **guessed** (Frames/Canvas forced to white `:171-172`, button text white `:182, 184`, Entry/Text white). Rebuilt: two-pass capture of every widget's real colours, replay on restore (`gui/utils/dark_mode_manager.py:105-120, 199-268`); palette now also `#404040` and `#333333` (`:72-78`) | PARTLY then ("restores originals" NEVER TRUE for plain Tk widgets); NO LONGER the implementation | aeb83cf, c89c1d4, 725fbaf, 169c7b5 |
| F9.2 | Setting persisted to cache | 301 | `gui/main_window.py:196, 368, 384, 388-389` | STILL TRUE | — |
| F10.1 | Saves/loads source, templates, target, mode, dark mode, client-name settings to JSON in home dir | 307-311 | `gui/utils/cache_manager.py:14-48`; keys `gui/main_window.py:360-370` | STILL TRUE; key `workflow_mode` added (`main_window.py:366, 382-383`) | 510f7b4 |
| F10.2 | Writer `main_window.py:328-340`, reader `:342-356` | 314-315 | now `:358-371` and `:373-389` | moved | — |
| F10.3 | Cross-app readers `extractor_main.py:36-60` and `extractor_window.py:473-485` | 316 | now `power_query_extractor/extractor_main.py:36-63` and `power_query_extractor/gui/extractor_window.py:587-599` | moved | — |
| F10.4 | MISMATCH: extractor looks for `gst_organizer_cache.json` in the current folder (`extractor_main.py:39`, `extractor_window.py:28`) → never finds the main app's cache — "confirmed bug" | 317 | True @47a1492 (READ: `extractor_main.py:39`, `extractor_window.py:28`). Now `extractor_main.py:41` checks `Path.home()/'.gst_organizer_cache.json'` **first** (old relative paths kept as fallbacks `:42-44`), and `extractor_window.py:118` uses the home path (READ c1583d6 diff) | NO LONGER TRUE | c1583d6 (23-07-2026) |
| F11.1 | Opens report via COM, Alt+A,R,A via SendKeys, waits, checks error dialogs, validates BB2 in "Info", extracts cells | 321 | `report_processor.py:209-415` (SendKeys `:260, 282-292`), error windows `:417-525`, validation `:638-685`, extraction `:527-636` | STILL TRUE | — |
| F11.2 | (1) copy with suffix (default `_Refreshed_{timestamp}`) | 330 | `report_processor.py:51, 127-130, 164`; timestamp `'%d%m%y_%H%M'` | STILL TRUE | — |
| F11.3 | (2) visible Excel instance | 331 | `:218, 225` | STILL TRUE | — |
| F11.4 | (4) waits configurable time (default 10, range 5-60) | 333 | `report_processor.py:50, 297-302`; spinbox `extractor_window.py:270-279`. The spinbox range only limits the arrow buttons; a typed value is passed unchecked (`:864`) (READ). Real total per report is much longer: 5 s load (`:253-255`) + 1 s + 2 s keys + wait + up to 5 s error checks + `max(10, wait)` extended wait (`:326`) + 2 s + 1 s + 3 s release wait (`:185`) (READ) | STILL TRUE as stated; understates elapsed time | — |
| F11.5 | (5) error dialogs found by enumerating windows | 334 | `:417-453` (classes `#32770`, `NUIDialog`, `bosa_sdm_msword`, title keywords `:425-436`) | STILL TRUE | — |
| F11.6 | (6) BB2 in "Info" non-empty/non-zero/non-error | 335 | `:638-685` ("Info" matched case-insensitively `:646`) | STILL TRUE | — |
| F11.7 | (8) a NEW hidden Excel instance for extraction; (9) reads cells per CELL_MAPPINGS | 337-338 | `:535-595` (log text says "read-only mode" `:541, 543` but `_safe_open_workbook` opens normally `:737`) | STILL TRUE (log wording misleading) | — |
| F11.8 | "Bug: unreachable `return {'success': True...}` at 342-345 after the inner try/except, which always returns at 329 / 337-340" | 340 | now `report_processor.py:383-386`; inner try `:258` returns at `:370` or `:378-381` (READ) | STILL TRUE (line moved) | — |
| F12.1 | Multi-sheet consolidated workbook, one sheet per EXTRACTION_CONFIG entry + Summary, `PQ_Extraction_Report_YYYYMMDD_HHMM.xlsx` | 344-350 | `power_query_extractor/core/data_consolidator.py:45-182` (name `:175-176`); RAN: 9 sheets + Summary (exp-3) | STILL TRUE as a description. **Not mentioned by any document:** for 9 output columns the workbook shows the wrong cell's value — see discrepancy X1 [GST RULE] | — |
| F13.1 | Export selected clients to a text file and open it | 354 | `gui/handlers/client_handler.py:225-257` | STILL TRUE | — |
| F13.2 | lines 221-253; `os.startfile` at 251 is Windows-only and "will crash" on macOS/Linux | 356-358 | now `os.startfile` at `client_handler.py:255`. It sits inside `try ... except Exception` (`:240, 256`), so off Windows the file is written, "Export Complete" shows, then an "Export Error: ... has no attribute 'startfile'" box appears — no crash. Same try/except @47a1492 (READ `client_handler.py:236-253`). RAN: `hasattr(os,'startfile')` is False on Linux (exp-1) | PARTLY: Windows-only — STILL TRUE; "will crash" — NEVER TRUE | — |

### 1.4 Section 4 (Data flow), doc lines 362-452

| # | Claim | Doc line | Code evidence at HEAD | Verdict | Commit |
|---|---|---|---|---|---|
| S4.1 | `find_excel_files()` → list of .xlsx/.xls paths | 371 | `scan_excel_files()` (`utils/helpers.py:309-365`) returns (path, stat) pairs for .xlsx/.xls/.xlsm | NO LONGER TRUE (function); .xlsm omitted by the doc even then | 56aaad9 |
| S4.2 | `validate_excel_file()` checks `PK\x03\x04` / `\xd0\xcf\x11\xe0` | 373 | `utils/helpers.py:264-265, 275-287` | STILL TRUE | — |
| S4.3 | parse → group by key; unmatched → variations with suggested pattern; completeness | 375-381 | `core/file_parser.py:196-220, 252-306` | STILL TRUE | — |
| S4.4 | Per client: create_client_structure, organize_files (sanitize + safe_copy), prepare_template_data, ITC & Sales reports, organisation report; summary at Level 1; `os.startfile(level1_folder)` | 389-403 | `processing_handler.py:253-414, 455-516`; startfile only on Windows, `open`/`xdg-open` otherwise (`:505-516`) | STILL TRUE | — |
| S4.5 | "[Windows] win32com: copy template, open, write Links, save / [Other] openpyxl: copy, open, write, save" | 396-397 | both now stage in local temp first (`core/excel_handler.py:147-215, 269-322`) | NO LONGER TRUE (detail) | 12d505f |
| S4.6 | PQ flow: "Loads target_folder from cache (BUG: wrong path)" | 411 | fixed (F10.4) | NO LONGER TRUE | c1583d6 |
| S4.7 | Scans for `Annual Statement-*` → latest | 413 | `extractor_window.py:622-632`: latest **by modified time** inside the chosen folder | STILL TRUE | — |
| S4.8 | For each client: latest `Version-*`; checks `ITC_Report_*.xlsx`, `Sales_Report_*.xlsx` | 416-417 | `:636-659`; latest version chosen by mtime since c1583d6 (before: by name, which mis-sorts day-first names across months — a bug the audit did not report); report check `_classify_version_reports` `:20-42` (one listing, case-insensitive) | STILL TRUE (selection rule changed) | c1583d6, 275d9a4 |
| S4.9 | Copy + `_Refreshed_{timestamp}`, visible Excel, SendKeys, wait, error check, BB2, save/close, new hidden Excel, extract | 424-433 | as F11 | STILL TRUE | — |
| S4.10 | `DataConsolidator.create_report()` → `PQ_Extraction_Report_*.xlsx` at Level 1; offer to open | 435-437 | `extractor_window.py:945, 955, 968-975` (open uses `os.startfile`, inside try/except) | STILL TRUE | — |
| S4.11 | Main app writes `~/.gst_organizer_cache.json`, keys incl. `recent_folders[]` | 443-447 | `cache_manager.py:15, 32-46`; `recent_folders` is written but never read anywhere (RAN grep) | STILL TRUE (write side); key list now also `workflow_mode` | 510f7b4 |
| S4.12 | PQ Extractor reads `./gst_organizer_cache.json` ← MISMATCH | 449-451 | fixed (F10.4) | NO LONGER TRUE | c1583d6 |

### 1.5 Section 5 (Business rules), doc lines 456-584

| # | Claim | Doc line | Code evidence at HEAD | Verdict | Commit |
|---|---|---|---|---|---|
| R1 | Pattern table: GSTR-2B-Reco `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` ITC; GSTR3B `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` GSTR3B; Sales `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$`; SalesReco `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$`; AnnualReport `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$` category Annual | 460-467 | `utils/constants.py:15, 31, 39, 47, 55` (exact) | STILL TRUE [GST RULE: file → return-type classification] | — |
| R2 | ImsReco `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$` | 463 | now `^ImsReco-([^-]+)-([^-]+)-(\d{8})(?:\s*\(\d+\))?\.xlsx?$` (`constants.py:23`) — accepts a trailing " (1)" duplicate marker; RAN `ImsReco-ABC Ltd-Maharashtra-30042024 (1).xlsx` parses with date '30042024' (exp-1) | NO LONGER TRUE | 5c8ce95 (23-07-2026) [GST RULE] |
| R3 | All patterns case-insensitive (`constants.py:15-61`) | 470 | `re.IGNORECASE` on each (`:15, 23, 31, 39, 47, 55`); RAN lower-case name parses (exp-1) | STILL TRUE | — |
| R4 | GSTR3B `(1)` prefix at `constants.py:32` | 471 | at `:31` | STILL TRUE (line off by one even @47a1492) | — |
| R5 | ImsReco needs exactly 8 digits (`constants.py:24`) | 472 | `(\d{8})` at `:23` | STILL TRUE [GST RULE: date/period field format]. The 8-digit date, and every other period/month/year group, is stored in `result['metadata']` (`core/file_parser.py:105-106`) and **never used anywhere** (RAN grep for `metadata`) — the app does not look at which period a file covers | — |
| R6 | Sales needs exactly 4 groups (`constants.py:39-44`) | 473 | `constants.py:39, 44` | STILL TRUE | — |
| R7 | `[^-]+` → a hyphen inside a client name breaks parsing | 474 | RAN `GSTR3B-ABC-XYZ Ltd-Maharashtra-Apr24.xlsx` → not parsed (variations) (exp-1) | STILL TRUE | — |
| R8 | "Complete" needs all six types (`constants.py:68-75`) | 478-484 | `constants.py:68-75`; `file_parser.py:259, 278` | STILL TRUE [GST RULE] | — |
| R9 | GSTR-3B Export allows unlimited duplicates (`file_parser.py:267-268`); others flagged | 486 | `core/file_parser.py:264-271` | STILL TRUE [GST RULE] | — |
| R10 | State map at `constants.py:234-274`; orissa→OD, pondicherry→PY, "national capital territory of delhi"→DL, "dadra and nagar haveli and daman and diu"→DD | 490-494 | now `utils/constants.py:251-292` (lines 272, 291, 286, 284) | STILL TRUE [GST RULE: state-code logic] TIME-VARYING (the list of states/UTs and names). The codes are two-letter abbreviations used for folder names, not numeric codes | — |
| R11 | Unknown state fallback `helpers.py:502-510`: 2+ words → first letters, 1 word → first 3 chars | 496 | `utils/helpers.py:608-618` (first letters of the first **two** words; else first 3 characters upper-cased) RAN: 'Unknown Place'→'UP', 'Atlantis'→'ATL' (exp-1). Note 'UP' equals the code the map gives Uttar Pradesh (`constants.py:279`) — the test suite pins this as known behaviour (`tests/test_helpers.py:96`) | STILL TRUE [GST RULE] (collision described, not judged) | — |
| R-SAN | Sanitisation (`helpers.py:66-112`): (1) Private→Pvt, Limited→Ltd **case-sensitive**; (2) invalid chars list; (3) control chars 0-31,127-159; (4) hyphens kept; (5) runs of spaces/underscores → `_`; (6) strip `_`, `.`, space; (7) max 200 | 500-507 | now `utils/helpers.py:68-126`. (1) changed: separators normalised to spaces first (`:91`), then whole-word, case-insensitive `\bPrivate\b`/`\bLimited\b` (`:92-93`); RAN `ABC PRIVATE LIMITED`→`ABC_Pvt_Ltd`, `ABC_Private_Limited`→`ABC_Pvt_Ltd`, `Privateer Unlimited` unchanged (exp-1). (2) `:96` same list; (3) `:99-102`; (4) yes; (5) `:109`; (6) `:112`; (7) default 200, base name cut to `200 - len(ext) - 10` (`:119-120`) | (1) NO LONGER TRUE; (2)-(7) STILL TRUE | 33d0baf, 4219ffb |
| R12 | Client-state key (`helpers.py:513-537`): default 35 (spinbox 15-100); over the limit the client part is cut keeping "-STATE"; minimum 5 chars | 511-514 | `utils/helpers.py:620-644`; spinbox `gui/tabs/setup_tab.py:440` (from 15 to 100). Cutting happens only if the space left is **more than 5** (`:637`), otherwise the whole key is cut to the limit and the state code is lost (RAN `max_length=7` → 'AAAAAAA', exp-1). This key function does **not** normalise underscores, so `ABC_Private_Limited` stays unabbreviated in the Level-2 folder name while file names get `ABC_Pvt_Ltd` (RAN exp-1) | PARTLY ("minimum 5" imprecise) | — |
| R13 | Folder template (`constants.py:81-90`) incl. "GSTR-3B Exports ({client})", "Other ITC related files ({client})", "Sales related files ({client})" | 518-527 | `constants.py:81-96` → "GSTR3B ({client})", "ITC ({client})", "Sales ({client})" | NO LONGER TRUE (level-4 names) | 7347383 |
| R14 | Type → folder map (`file_organizer.py:288-295`); Annual Report → Version folder | 531-539 | `core/file_organizer.py:305-312` | STILL TRUE | — |
| R15 | Excel validation (`helpers.py:210-247`): exists & file; ext in .xlsx/.xls/.xlsm/.xltx/.xltm; ≥1024 bytes; signature | 543-547 | `utils/helpers.py:290-306, 275-287`. Note the scan only *lists* .xlsx/.xls/.xlsm (`:327`), so .xltx/.xltm in the source folder are never considered | STILL TRUE | — |
| R16 | Client name validation (`helpers.py:321-342`): not empty, ≥2, ≤100, no `<>:"/\|?*` | 551-555 | `utils/helpers.py:428-449` | STILL TRUE | — |
| R17 | Modes (`constants.py:157-173`): fresh = new folder, backup if destination exists; rerun = latest Annual Statement folder, new Version folder; resume = latest folder, skip existing files | 559-562 | `constants.py:174-190`; `core/file_organizer.py:163-196` ("latest" = newest modified time, `:184`), `:276-287`. **But** every run, including resume, creates a new `Version-<current minute>` folder (`:102-108`), so in resume mode the destination files only "already exist" if the earlier run started in the same minute; otherwise everything is copied again into a new Version folder. RAN: resume with a different minute → new folder `Version-010199 0000`, all files 'Success' (copied); same minute → 'Skipped' (exp-2). In rerun mode an existing destination is simply overwritten (no branch for it, `:277-287`) | PARTLY — the text is a faithful reading of each line, but "resume" cannot resume an interrupted run started in an earlier minute (HIGH) | — |
| R18 | PQ validation (`report_processor.py:597-644`): Info sheet; BB2 fails if None, blank, "0", "0.0", or #REF!/#VALUE!/#NAME?/#NULL!/#DIV/0!/#N/A/#NUM!/#ERROR!; success = client name expected | 566-570 | `report_processor.py:638-685` (exact list at `:675`) | STILL TRUE (template convention, technical) | — |
| R19 | MAX_FILE_SIZE_MB 100 never enforced | 576 | `constants.py:198`, unreferenced | STILL TRUE | — |
| R20 | EXCEL_SAFETY max_retries 3, retry_delay 1, temp_prefix 'temp_gst_', verification_size_min 10240 — not used | 577-580 | `constants.py:234-245`; whole dict stored as `self.safety_config` (`core/excel_handler.py:55`) and never read | STILL TRUE | — |
| R21 | wait_time 10 (`report_processor.py:27`) | 581 | `:50` | STILL TRUE (moved) | — |
| R22 | Long-path threshold 218 (`excel_handler.py:304`, `report_processor.py:676`) | 582 | `core/excel_handler.py:356`, `report_processor.py:720` | STILL TRUE (moved) | — |
| R23 | Unique-path safety limit 100 (`helpers.py:363`) | 583 | `utils/helpers.py:469`; `get_unique_path` itself is never called (D15) | STILL TRUE but dead | — |
| R24 | Short path via `win32api.GetShortPathName`, Windows only, graceful (`helpers.py:476-478`) | 584 | `utils/helpers.py:580-590` | STILL TRUE | — |

### 1.6 Section 6 (Edge cases handled), doc lines 588-627

| # | Claim | Doc line | Code evidence | Verdict | Commit |
|---|---|---|---|---|---|
| E1 | Source folder missing → FileNotFoundError (`file_parser.py:154`) | 594 | `core/file_parser.py:152-153` (the GUI checks first, `gui/handlers/file_handler.py:126-128`) | STILL TRUE | — |
| E2 | Source is a file → ValueError (`:156`) | 595 | `:154-155` | STILL TRUE | — |
| E3 | No Excel files → empty result, warning (`:159-161`) | 596 | `:159-161` | STILL TRUE | — |
| E4 | Corrupted Excel (bad signature) → excluded and **added to variations** (`file_parser.py:186-194`) | 597 | @47a1492 `find_excel_files()` already dropped invalid files (`helpers.py:259-260` @47a1492), so the variations branch in `_process_file` could not normally run; a603c02 removed that branch. At HEAD invalid files vanish silently: not in variations, not in the scan count. RAN: an invalid `Sales-...xlsx` (bad signature, 2 KB) appeared nowhere (exp-2); `tests/test_scan.py:52` "test_invalid_excel_is_skipped_entirely" pins this | NEVER TRUE (the "added to variations" part) | a603c02 removed the dead branch |
| E5 | File <1KB → treated as corrupted (`helpers.py:232`) | 598 | `utils/helpers.py:303` | STILL TRUE (silently dropped, as E4) | — |
| E6 | Destination exists, fresh → backup with timestamp (`file_organizer.py:260-265`) | 599 | `core/file_organizer.py:277-282`; backup name `<stem>_backup_<DDMMYY_HHMM><ext>` (`utils/helpers.py:224-226`) | STILL TRUE | — |
| E7 | Destination exists, resume → skipped (`:266-269`) | 600 | `:283-287` (see R17 limitation) | STILL TRUE (literal) | — |
| E8 | Long paths >218 → 8.3 short path (`excel_handler.py:296-317`) | 601 | `core/excel_handler.py:342-369` | STILL TRUE | — |
| E9 | Network disconnect, locked files — not handled (FRAGILE) | 602-603 | no specific handling in organiser/copy (`utils/helpers.py:186-215` generic except). The extractor does try closing same-name workbooks and other path forms (`report_processor.py:743-783`) | STILL TRUE | — |
| E10 | Empty client name from parsing → "logged as warning, still processed" (`file_parser.py:114-116`) | 609 | Validation runs only when both client and state are non-empty (`core/file_parser.py:109`); an empty (whitespace-only) client makes `_process_file` send the file to variations with issue "Could not parse filename" (`:210-220`). RAN: `GSTR3B- -Maharashtra-Apr24.xlsx` → parsed=True, client='' , no warning (exp-1). Same code @47a1492 | NEVER TRUE | — |
| E11 | Unknown state → fallback (`helpers.py:502-510`) | 610 | R11 | STILL TRUE [GST RULE] | — |
| E12 | Client name >100 → error, warning only, doesn't block (`helpers.py:334-335`) | 611 | `utils/helpers.py:441-442`; only appended to `warnings` (`core/file_parser.py:113-115`), which nothing displays or reads (RAN `grep -rn warnings`: only `core/file_parser.py:114-119` write it — HIGH). Side note: `core/excel_handler.py:17` `warnings.filterwarnings('ignore')` silences all Python warnings for the whole app as soon as that module is imported | STILL TRUE | — |
| E13 | Empty filename → "unnamed" (`helpers.py:70-71`) | 612 | `:72-73` (also `:115-116`) | STILL TRUE | — |
| E14 | Unicode: control chars replaced (`helpers.py:85-88`) | 613 | `:99-102` | STILL TRUE | — |
| E15 | Hyphen in client name breaks parsing | 614 | R7 | STILL TRUE | — |
| E16 | Duplicate client-state → merged entry (`file_parser.py:226-246`) | 615 | `:226-250` | STILL TRUE | — |
| E17 | Bare `except:` at `excel_handler.py:62,124,222-228,263,545` | 623 | now `core/excel_handler.py:63, 126, 235, 245, 251, 302, 597` (READ) | STILL TRUE (moved) | — |
| E18 | Thread exception → messagebox (`processing_handler.py:455-458`) | 624 | `:522-525` — the messagebox is shown from the worker thread | STILL TRUE (moved) | — |
| E19 | COM cleanup in finally (`excel_handler.py:206-228`, `report_processor.py:353-374`) | 625 | `core/excel_handler.py:230-260`, `report_processor.py:394-415` | STILL TRUE (moved) | — |

### 1.7 Section 7 (Gaps and vulnerabilities), doc lines 631-671 — was each fixed in July?

| # | Gap/bug as listed in March | Doc line | Status at HEAD (evidence) | Verdict | Commit |
|---|---|---|---|---|---|
| G1 | Cache path mismatch (extractor reads `./gst_organizer_cache.json`) | 635 | Fixed: `extractor_main.py:41`, `extractor_window.py:118` read `~/.gst_organizer_cache.json` (READ). Not "centralised in a shared constant" as recommended — the path is written out separately in `gui/utils/cache_manager.py:15`, `extractor_main.py:41`, `extractor_window.py:118` | FIXED (differently from the recommendation) | c1583d6 |
| G2 | Lambda capture in loops (`processing_handler.py` 41-48, 178, 181, 189, 195-200, 206, 216, 221, 237-239) | 637 | Fixed via `_log_async`/`_progress_async` default-argument capture (`processing_handler.py:23-36`) — the exact recommended technique. Not covered by this fix: the 45 direct `self.app.log_message(...)` calls (RAN count) and the messageboxes from the worker thread (F8.4) — those never had the capture problem, but they are Tk calls from a worker thread | FIXED | 0890022 |
| G3 | PQ Extractor imports win32 modules unguarded (`report_processor.py:11-14`) | 639 | Guarded `report_processor.py:20-36`; extractor panel import guarded `gui/tabs/extract_tab.py:24-30` | FIXED | c463256 |
| G4 | `os.startfile()` without platform check in `client_handler.py:251` | 641 | Still unguarded at `gui/handlers/client_handler.py:255` (inside try/except, so it shows an error box rather than crashing — F13.2). Also unguarded: `extractor_window.py:973` (inside try/except, logs error) | STILL PRESENT | — |
| G5 | requirements.txt missing pywin32 | 643 | `requirements.txt:9` `pywin32>=305; sys_platform == 'win32'` (exactly the recommended line) | FIXED | 064c634 |
| G6 | Worker thread touches `client_data`, `client_folder_settings` etc. without locks; rescan during processing could break | 645 | Unchanged: worker reads `self.app.client_data[client_key]` each loop (`processing_handler.py:232`) and receives `client_folder_settings` by reference (`:193`). Scan/Re-scan buttons are not disabled during processing (their state depends only on the Setup checklist, `gui/tabs/setup_tab.py:585-591`); Re-scan first empties the same dict (`gui/handlers/file_handler.py:102-105`) (READ). A rescan mid-run would make later clients fail with a KeyError, which is caught and counted as a failed client (`processing_handler.py:442-453`) (MEDIUM, by reading). Only a *second Start* is now blocked (`:71-76`) | STILL PRESENT (partly mitigated) | 33d0baf (re-entrancy guard) |
| G7 | `sanitize_filename` case-sensitive (`helpers.py:78-79`) vs `create_client_state_key` case-insensitive | 649 | Both now `\b...\b` + IGNORECASE (`utils/helpers.py:92-93`, `:624-625`). A new, smaller inconsistency remains: only `sanitize_filename` first turns underscores into spaces (`:91`), so underscore-joined names abbreviate in file names but not in the Level-2 client folder name (RAN exp-1) | FIXED (residual difference) | 33d0baf, 4219ffb |
| G8 | `validate_client_name` used to validate the state (`file_parser.py:113`) | 651 | `core/file_parser.py:111`; RAN state 'M' → warning text "State name issue: Client name too short" (exp-1) — and that warning is never shown (E12) | STILL PRESENT | — |
| G9 | Working directory: `main.py:36` creates logs/ and temp/ relative to `__file__` (correct); `extractor_main.py:39` reads cache relative to CWD | 655 | `main.py:31-37` unchanged (note: these `logs/` and `temp/` folders are created on every start but nothing writes into them — RAN grep for `logs`/`temp` users: none besides `main.py:35`; HIGH). Extractor: home path now first (`extractor_main.py:41`), CWD-relative fallbacks remain (`:42-44`). Log files are still CWD-relative: `main.py:24` `gst_organizer.log`, `extractor_main.py:16` `pq_extractor.log` | PARTLY FIXED | c1583d6 |
| G10 | Second fresh run creates a new timestamped folder | 657 | `core/file_organizer.py:165-173` | STILL TRUE (by design) | — |
| G11 | Two instances: same cache, same timestamps, Excel COM conflicts | 659 | No locking anywhere (READ) | STILL PRESENT | — |
| G12 | Excel already open: `GetActiveObject` reuses the user's Excel (`excel_handler.py:122-123`); `_safe_open_workbook` (line 285) reuses an already-open workbook | 661 | `core/excel_handler.py:123-128`; detection now shared `utils/excel_com.py:19-29` (`excel_handler.py:337`). **Not noticed by the audit:** the cleanup always calls `excel.Quit()` because its guard `if self.excel_app is None` (`:242`) is always true — `self.excel_app` is set to None in `__init__` (`:56`) and never assigned anywhere else (RAN grep). The comment "Don't quit if Excel was already open" (`:241`) does not match what the code does; on Windows this would ask the user's own running Excel to quit after each report (CANNOT VERIFY on Windows; logic HIGH by reading). Present @47a1492 too | STILL PRESENT (and worse than described) | — |
| G13 | OneDrive path with spaces | 663 | environmental | CANNOT VERIFY | — |
| G14 | No template structure validation; error "only surfaces during report creation, after folders are created" | 667 | Still no check: `validate_processing_inputs` checks existence only (`gui/handlers/file_handler.py:132-151`); the Setup checklist checks extension only (`gui/utils/setup_validation.py:56-64`). Worse than described: no error surfaces at all — COM path creates a new empty "Links" sheet and reports success (`core/excel_handler.py:180-183, 218-221`); openpyxl path writes nothing and returns True (RAN exp-2: no-Links template → `True`, sheets `['Data']`, nothing written) | STILL PRESENT; the "error surfaces" part NEVER TRUE | — |
| G15 | No disk-space check | 669 | none (READ) | STILL PRESENT | — |
| G16 | No concurrent-access protection | 671 | none (READ) | STILL PRESENT | — |

### 1.8 Section 8 (Design decisions), doc lines 675-714

| # | Claim | Doc line | Evidence | Verdict | Commit |
|---|---|---|---|---|---|
| DD1 | Tkinter chosen; PyQt5 branch "reconsidered later but abandoned" | 678 | Tkinter still used in A; the PyQt5 commits are project B (H5) | PARTLY ("abandoned" CANNOT VERIFY) | — |
| DD2 | COM because openpyxl cannot keep/refresh Power Query; `excel_handler.py:3` says so; fallback lines 230-273 warns | 681 | `core/excel_handler.py:3`, fallback now `:262-322` (warnings `:98, 267`). Strengthened since: pywin32 now required on Windows "so reports cannot lose Power Query" (`utils/requirements_check.py:49-57`) | STILL TRUE | be6418c |
| DD3 | Separate PQ Extractor app — user may run it days later | 684 | Now hosted as Step 4 in the main window as well (bfcb24a); the design note cites D3 as the reason Step 4 keeps its own folder picker | NO LONGER the whole truth (both now) | bfcb24a |
| DD4 | JSON cache in home dir | 687 | `gui/utils/cache_manager.py:15` | STILL TRUE | — |
| DD5 | Regex on file names (MEDIUM) | 690 | inference | CANNOT VERIFY (intent) | — |
| DD6 | State codes to keep paths short; `file_organizer.py:87` comment "Use state code instead of full name" | 693 | `core/file_organizer.py:91` | STILL TRUE (moved) — and July went further (7347383 shortened category folders; its message records measured path lengths) | 7347383 |
| DD7 | Daemon thread + `root.after` | 696 | `processing_handler.py:129-134` | STILL TRUE | — |
| DD8 | Handlers take `app_instance` | 699 | `gui/handlers/*.py:16-17, 15-16, 20-21` | STILL TRUE | — |
| DD9 | Client name in brackets optional (checkbox + per-client toggle) | 702 | F2.4 | STILL TRUE | — |
| PC1 | Pattern table: win32 import guarding INCONSISTENT (`report_processor.py:11-14`) | 708 | guarded now | NO LONGER TRUE | c463256 |
| PC2 | `os.startfile` INCONSISTENT (`client_handler.py:251`) | 709 | `client_handler.py:255` | STILL TRUE | — |
| PC3 | "Thread-safe UI updates — CONSISTENT across processing_handler" | 710 | F8.4: many direct Tk calls from the worker | NEVER TRUE | — |
| PC4 | "Logging pattern: All modules use `logger = logging.getLogger(__name__)` — CONSISTENT" | 711 | Several modules that do log use the root logger instead: `core/excel_handler.py:26, 28` (`logging.info/warning` at import), `main.py:99, 105, 115, 131`, `gui/main_window.py:193` (inline getLogger) (RAN grep); many GUI modules have no logger (they do not log). Same @47a1492 | PARTLY (overstated) | — |
| PC5 | COM try/finally CONSISTENT | 712 | E19 | STILL TRUE | — |
| PC6 | Cache path INCONSISTENT | 713 | G1 | NO LONGER TRUE | c1583d6 |
| PC7 | Case sensitivity INCONSISTENT (`helpers.py:78-79` vs `:517`) | 714 | G7 | NO LONGER TRUE (residual underscore difference) | 33d0baf, 4219ffb |

### 1.9 Section 9 (Dependencies and environment), doc lines 718-763

| # | Claim | Doc line | Evidence | Verdict | Commit |
|---|---|---|---|---|---|
| DEP1 | openpyxl ≥3.1.2 required; lxml ≥4.9.0 recommended | 724-725 | `requirements.txt:5, 12` | STILL TRUE | — |
| DEP2 | pywin32 "Not specified … MISSING from requirements.txt" | 726 | `requirements.txt:9` | NO LONGER TRUE | 064c634 |
| DEP3 | (not mentioned) pytest | — | `requirements.txt:15` `pytest>=7.0.0`. Side effect: the start-up install offer runs `pip install -r requirements.txt` (`main.py:84-88`, `utils/dependency_installer.py:133-134`), so accepting it also installs pytest and lxml on the user's PC (READ) | new | 6c1e1c6, a5ad487 |
| DEP4 | Minimum 3.7 (requirements comment and README) | 730 | `requirements.txt:2`, `README.md:51` | STILL TRUE as text (S1.2) | — |
| DEP5 | No dataclass, no match/case, no `|` unions | 731-732 | none found (RAN grep; 3.7 grammar parse passed) | STILL TRUE | — |
| DEP6 | OS table: win32com/pythoncom fallback openpyxl (excel_handler) / NONE (report_processor); win32api graceful (helpers) / CRASHES (report_processor); win32con & win32gui CRASH on import | 738-742 | report_processor now guarded (`:20-36`) | NO LONGER TRUE for report_processor rows | c463256 |
| DEP7 | `os.startfile` used in processing_handler, client_handler, extractor_window; platform check only in processing_handler | 743 | `processing_handler.py:507-512` (checked), `client_handler.py:255`, `extractor_window.py:973` (unchecked, both in try/except) | STILL TRUE | — |
| DEP8 | (not mentioned) new startup dependency check | — | `main.py:66-106`, `utils/requirements_check.py` (pythoncom essential on Windows, optional elsewhere `:57`), `utils/dependency_installer.py` (offer to pip-install then restart) | new | f363946, be6418c, a5ad487 |
| FS1 | Windows path separators for Excel via `clean_windows_path()` | 748 | `utils/helpers.py:418-422` | STILL TRUE | — |
| FS2 | 260-char Windows limit; 218 for COM | 749 | E8, R22 | STILL TRUE | — |
| FS3 | OneDrive expected; home writable; temp writable (`%TEMP%/gst_excel_temp/`) | 750-752 | `core/excel_handler.py:53` — and since 12d505f every report is staged there, so temp must also have room for each report | STILL TRUE (more important now) | 12d505f |
| SET1 | Setup: install Python 3.7+ with PATH; pip install openpyxl, lxml, pywin32; run `pywin32_postinstall.py -install` as admin; download; `python main.py`; `python launch_extractor.py` | 756-763 | Commands still valid; `main.py` now detects a missing pywin32 and offers to install it (but does not run the post-install script) (READ `utils/dependency_installer.py`). Whether the post-install step is needed CANNOT VERIFY (Windows) | STILL TRUE / partly superseded | a5ad487 |

### 1.10 Section 10 (UI), doc lines 767-853

| # | Claim | Doc line | Evidence | Verdict | Commit |
|---|---|---|---|---|---|
| UI1 | Main window 1200×800, min 1000×700; clam theme; primary #0078D4, success #107C10; title bar / notebook / status bar | 771-773 | `constants.py:134-136, 146-147`; `gui/main_window.py:57-77, 131-149` | STILL TRUE | — |
| UI2 | Title bar #0078D4, 70 px, "🏢 GST FILE ORGANIZER & REPORT GENERATOR" 18pt bold white, subtitle 11pt, "🌙 Dark Mode" checkbox top-right | 776-779 | `gui/widgets/title_bar.py:21-51` | STILL TRUE | — |
| UI3 | Tab 1 Setup: two columns 60/40; welcome banner; Source (green), Templates (blue), Target (red, "🚨 IMPORTANT!"), Processing Mode (teal) with 3 radios, client-name checkbox, spinbox 15-100 default 35; "🔍 SCAN FILES" (green) + "🔄 RE-SCAN" (orange); right column instructions card + file-patterns card | 781-794 | True @47a1492 (READ `setup_tab.py:44, 76, 90-91, 119-120, 165-176, 203-204, 271-272, 297, 309, 325, 353` @47a1492). At HEAD one scrolling column: header card with readiness pill, live **checklist** card, **"WHAT SHOULD THIS RUN DO?"** workflow card, "📂 FOLDERS & TEMPLATES" section, "⚙️ PROCESSING MODE" section (3 radios, "Include client name in Level 4 folder names", spinbox 15-100 with hint), actions card with ttk "🔍  Scan Files"/"🔄  Re-scan" (disabled until the source folder is valid) and a status strip, "📝 EXPECTED FILE NAMES" card generated from FILE_PATTERNS; hover tooltips everywhere (`gui/tabs/setup_tab.py:159-615`) | NO LONGER TRUE | 322717f, d64d3a9, 510f7b4, 169c7b5, bf2f1dd |
| UI4 | Tab 2 Validation: 40/60; collapsible orange instructions; Scan Summary (Consolas 9); "🧪 DRY RUN" orange, "📋 EXPORT LIST" teal, "🚀 START" green; right: hint "⌨️ Use ↑↓ arrows + SPACE to select", three buttons, tree columns ✓/Client Name/State/Status/Files/Missing/Extra/FolderName; #E8F5E8 / #FFF4E5 | 796-809 | `gui/tabs/validation_tab.py:18-247` (hint text continues ", or click with mouse" `:163`; left column fixed 450 px `:30`; FolderName heading text is "📁 Name" `:221`) | STILL TRUE (file unchanged since edac6a3) | — |
| UI5 | Tab 3 Processing: progress bar, % label, current-operation label; START PROCESSING green, STOP red disabled; dark log #1E1E1E Consolas 9 with success/warning/error/info colours | 811-815 | `gui/tabs/processing_tab.py:17-117` | STILL TRUE (unchanged since edac6a3) | — |
| UI6 | (not in audit) Tab 4 "🔄 Step 4: Extract" | — | `gui/tabs/extract_tab.py:48-105`: explanatory strip + embedded ExtractorPanel (its own Setup/Processing sub-tabs, i.e. a notebook inside a notebook) or a "Power Query refresh is unavailable" notice if pywin32 is missing | new | bfcb24a |
| UI7 | Status bar #323130, 30 px, "💡" prefix, "v3.0 \| Production Ready" | 817-820 | `gui/utils/status_bar.py:15-37` | STILL TRUE | — |
| UI8 | PQ window 1000×700, min 900×600, clam; title bar #0078D4 60 px "🔄 Power Query Report Extractor" | 824-829 | `extractor_window.py:1042-1046, 150-161` — banner only when standalone (`show_header`) | STILL TRUE (standalone) | a677670 |
| UI9 | PQ Setup tab: left 400 px fixed: folder entry + Browse + Scan; wait spinbox 5-60 default 10; suffix default `_Refreshed_{timestamp}`; "Skip Refresh"; "🚀 Start Processing" full width, disabled initially; right: headers Client Name / ITC / Sales / Last Refresh Status; ✓ Select All, ✗ Deselect All, ITC Only, Sales Only; scrollable list | 831-845 | `extractor_window.py:184-481` | STILL TRUE | — |
| UI10 | PQ Processing tab: label + bar; log Consolas 9 light; info black, success green, warning yellow, error red | 847-850 | `:494-571` | STILL TRUE | — |
| UI11 | PQ status bar #e0e0e0 25 px | 852-853 | `:573-585` (standalone only) | STILL TRUE | — |

### 1.11 Section 11 (Reconstruction blueprint), doc lines 857-948 — what happened to each recommendation

| # | Recommendation | Doc line | Outcome at HEAD | Commit |
|---|---|---|---|---|
| K1 | Fix cache path; centralise in a shared constant | 904 | Fixed, not centralised (G1) | c1583d6 |
| K2 | Fix lambda capture with default-argument binding | 906 | Done exactly so | 0890022 |
| K3 | Guard PQ imports like `excel_handler.py:20-28` | 908 | Done (`report_processor.py:20-36`) | c463256 |
| K4 | Platform check for `client_handler.py:251` | 910 | Not done (`client_handler.py:255`) | — |
| K5 | `pywin32>=305; sys_platform == 'win32'` in requirements | 912 | Done exactly so (`requirements.txt:9`) | 064c634 |
| K6 | Case-insensitive `\bPrivate\b` in `sanitize_filename` | 914 | Done (plus underscore handling) | 33d0baf, 4219ffb |
| K7 | Validate templates have "Links" before processing | 916 | Not done (G14) | — |
| K8 | Remove dead code (`calculate_file_hash`, `_write_section`, `_create_details_sheet`, `TEMPLATE_EXTENSIONS`, `EXCEL_SIGNATURES`, `MAX_FILE_SIZE_MB`, `# ADD THIS` comments) | 918 | Not done except the file_parser `# ADD THIS` comment (D11) | 56aaad9 (one comment) |
| K9 | Remove `[DEBUG]` logging in excel_handler and processing_handler | 920 | Partly: 35 `logger.info("[DEBUG] ...")` in excel_handler/file_parser downgraded to `logger.debug` (064c634 message). Left: `core/excel_handler.py:83` `logger.error(f"[DEBUG] Template file does not exist...")` and 14 user-visible "[DEBUG]" lines in the processing log (`processing_handler.py:333-411`) — 064c634's message says these were left on purpose (RAN `grep -c "\[DEBUG\]"`: 1 and 14) | 064c634 |
| K10 | Template validation at scan time | 922 | Not done | — |
| K11 | Foundation: Python 3.8+, `pip install openpyxl>=3.1.2 pywin32>=305`; build order constants → helpers → parser → organiser → excel_handler → gui → extractor; feature priority and complexity tables; skill ideas | 859-948 | Advice / estimates | CANNOT VERIFY (opinion) |

### 1.12 Section 12 (Changelog archaeology), doc lines 952-1162

| # | Claim | Doc line | Evidence (READ `git log`/`git show`) | Verdict |
|---|---|---|---|---|
| C1 | Branch map: edac6a3 → 6237597 → 9c56a1c (merge with cbfa4a5) → d674386 → 5ea77c2 → 1cb878b → 01862a0 → 1c48722 → 4c42758 → bd176bf → 47a1492 → fc279af; copilot-worktree: 47a1492 → c258047 → a62d3cc → c21f788 | 956-963 | Parents: cbfa4a5←edac6a3, 6237597←edac6a3, 9c56a1c = merge(6237597, cbfa4a5), then linear; c258047←47a1492 (READ `%P`). The "worktree" line is B's main | STILL TRUE as history (branch name CANNOT VERIFY) |
| C2 | Phase 1 "edac6a3 — 2025-10-29" | 968 | author date **2025-08-31 22:11** | NEVER TRUE (fc279af version had 2025-08-31) |
| C3 | edac6a3: 110 files, 12,916 insertions; includes `__pycache__` and `gst_organizer.log` (5,689 lines) | 972 | RAN `git show --stat edac6a3`: 110 files, 12916 insertions, 70 `.pyc`, `gst_organizer.log` 5689 lines (log not opened), `cell_mappings copy.py` 63 lines | STILL TRUE |
| C4 | "The log file shows real-world processing of GST files with actual client names" | 974 | not opened (confidentiality rule) | CANNOT VERIFY |
| C5 | Phase 2 "(cbfa4a5, 6237597, 9c56a1c — 2025-10-29)" | 976 | cbfa4a5 is **2025-08-31 22:14**; 6237597 and 9c56a1c 2025-10-29 | NEVER TRUE for cbfa4a5 |
| C6 | cbfa4a5 removed the 5,689-line log | 980 | RAN `--stat`: 5689 deletions | STILL TRUE |
| C7 | 6237597: 4 files, +59 −45; (1) file_parser progress_callback + `_update_progress`; (2) file_handler `self.app.file_parser.scanned_files` → `self.app.scanned_files`; (3) processing_handler wrapped "ALL" UI calls in `root.after`, configurable max_len; (4) helpers `from utils.constants` → `from .constants` | 985-989 | All four diffs confirmed (READ `git show 6237597`), except (3) "ALL": only some calls were wrapped (F8.4). The commit message also claims "Corrected non-existent app.py reference in main.py" — main.py is not in the commit (D9) | PARTLY |
| C8 | "Why": circular import would crash on startup; thread issue caused random crashes; "the scanned_files attribute error would crash on scan completion" | 990-993 | `FileParser` has always had `scanned_files` (set in `reset()` called from `__init__`, edac6a3 `file_parser.py:32-34` READ), and `scan_folder` returns that same dict, so the old line could not raise an AttributeError (HIGH). The `get_state_code` import was already inside the function; whether the absolute import ever failed depends on how the program was started (LOW) | NEVER TRUE (attribute crash); CANNOT VERIFY (import crash) |
| C9 | 9c56a1c merge of local fixes + remote log deletion | 996-999 | parents 6237597 + cbfa4a5 | STILL TRUE |
| C10 | d674386: README overhaul +239 −135 | 1003-1006 | RAN `--stat`: 239/135 | STILL TRUE |
| C11 | 5ea77c2 removed `# !/usr/bin/env python3` (with a space) | 1008-1011 | diff shows `-# !/usr/bin/env python3` (the fc279af version wrongly wrote `#!/...`) | STILL TRUE |
| C12 | 1cb878b: 79-line .gitignore | 1013-1016 | RAN: 79 insertions | STILL TRUE (now 80 lines, S2.3) |
| C13 | 01862a0: removed 70 .pyc | 1018-1021 | RAN: 70 files | STILL TRUE |
| C14 | 1c48722: removed 6 lines from client_handler (redundant tag colours) | 1025-1028 | RAN: 6 deletions | STILL TRUE ("Why" = inference) |
| C15 | 4c42758: mousewheel for canvas (Enter/Leave) and log text in extractor_window | 1032-1035 | +30 lines (RAN). The Enter/Leave + `unbind_all` approach was later found to kill the Setup tab's scrolling once both lived in one window, and replaced (`extractor_window.py:434-478`, bf2f1dd) | STILL TRUE as history; NO LONGER the code | 
| C16 | bd176bf: 3 files, +412 −162; `.claude/settings.local.json` new; wait time, suffix, selective ITC/Sales; extractor two-column UI etc. | 1037-1048 | RAN `--stat`: .claude 9, report_processor 58, extractor_window 507; 412/162 | STILL TRUE |
| C17 | 47a1492: 2 files, +126 −63; deleted `cell_mappings copy.py`; added 6 configs Reco Diff, Table 6J Match, Table 8D Limits, IMS Reco, Extra In Purchase, Extra In GSTR2B ("Reco" C37 → "IGST Diff") | 1050-1056 | RAN `--stat` matches; configs at `power_query_extractor/config/cell_mappings.py:104-228`; file unchanged since 47a1492 | STILL TRUE [GST RULE] TIME-VARYING (table numbers 6J/8D, sheet "T6 of R9"/"T8 of R9") |
| C18 | fc279af: 856-line audit, 2026-03-04 | 1060-1063 | 856 insertions, 04-03-2026 22:16 | STILL TRUE |
| C19 | c258047 / a62d3cc / c21f788 "— 2026-03-04" | 1065 | all 2026-02-07 | NEVER TRUE (date) |
| C20 | c258047: 33 files, +3,070 −541; new models/, gui/pyqt_main_window.py, gui/views/, gui/controllers/, gui/theme/, gui/widgets/collapsible_group.py; core cleanup; main.py & launch_extractor.py for PyQt5; PyQt5 in requirements | 1067-1079 | RAN `--stat` 33 files 3070/541; commit message lists the same (READ) | STILL TRUE (as B's history) |
| C21 | a62d3cc deleted 17 Tkinter GUI files | 1081-1084 | RAN: 17 files, 3298 deletions | STILL TRUE |
| C22 | c21f788 fixed `gui/__init__.py` import | 1086-1089 | 1 file, +5 −5 | STILL TRUE |
| C23 | "This branch was NEVER merged to main" | 1091 | not in A's history (READ `git log A/main`) | STILL TRUE |
| C24 | 12B file evolution table (main.py last 5ea77c2; launch_extractor never modified; requirements never; file_parser 6237597; file_organizer/excel_handler/constants never; helpers 6237597; main_window never; file_handler 6237597; client_handler 1c48722; processing_handler 6237597; tabs/utils/widgets never; extractor_window bd176bf; report_processor bd176bf; cell_mappings 47a1492; data_consolidator never) | 1095-1115 | True on 05-03-2026. At HEAD (RAN `git log --format=%h A/main -- <file>`): main.py a5ad487/be6418c/f363946; requirements 064c634/6c1e1c6; file_parser 56aaad9/064c634/a603c02; file_organizer 275d9a4; excel_handler ba2a753/064c634/12d505f; constants 510f7b4/7347383/5c8ce95; helpers 4f84ba9/56aaad9/4219ffb/33d0baf/a603c02; main_window 725fbaf/c89c1d4/510f7b4/bfcb24a; client_handler 33d0baf; processing_handler 510f7b4/33d0baf/0890022; setup_tab bf2f1dd/169c7b5/510f7b4/d64d3a9/322717f; dark_mode_manager 725fbaf/aeb83cf; extractor_window bf2f1dd/bfcb24a/2aaa2b1/a677670/2acffed/c1f425e/c1583d6/275d9a4; report_processor c463256/2aaa2b1/ba2a753/064c634; extractor_main c1583d6. **Still never changed since edac6a3 (or since the date given):** launch_extractor.py, gui/handlers/file_handler.py (6237597), gui/tabs/validation_tab.py, gui/tabs/processing_tab.py, gui/utils/cache_manager.py, status_bar.py, ui_helpers.py, gui/widgets/collapsible_frame.py, title_bar.py, data_consolidator.py, cell_mappings.py (47a1492) | NO LONGER TRUE for most rows |
| C25 | 12C: "Client name max length (35) | edac6a3 | 6237597 | Made configurable via GUI" | 1124 | The spinbox (15-100) already existed in edac6a3 (`setup_tab.py:271-273` @edac6a3, IntVar `main_window.py:89` @edac6a3, READ). 6237597 only made the long-name warning use the setting instead of a fixed 10. (fc279af version said "added in c258047" — also wrong; that is a B commit) | NEVER TRUE |
| C26 | 12C other rows (6 patterns stable; state map stable; GSTR-3B duplicates; thread-safe updates 6237597; ITC/Sales selective, wait time, skip refresh bd176bf; ITC mappings 47a1492; Sales mappings stable) | 1121-1130 | True on 05-03-2026. Since then: ImsReco pattern changed (5c8ce95); category folders changed (7347383); sanitisation rule changed (33d0baf, 4219ffb) | NO LONGER TRUE for "6 file patterns — No change" |
| C27 | 12D rows (6237597 ×3, 1c48722, 4c42758 — "Setup tab canvas also binds bind_all which may conflict", bd176bf — "extended wait still `max(10, self.wait_time)` at line 285") | 1134-1141 | 4c42758 prediction came true only after the two windows were merged (bfcb24a) and was fixed in bf2f1dd (`gui/tabs/setup_tab.py:193-201`, `extractor_window.py:434-478`, `tests/test_scroll_bindings.py`); extended wait still `report_processor.py:326` | 4c42758 row: materialised then FIXED (bf2f1dd); bd176bf row STILL TRUE |
| C28 | 12E: PyQt5 abandoned; `cell_mappings copy.py` backup; dead consolidator methods; EXCEL_SAFETY "only `save_options` is referenced (by `excel_handler.py:54` which stores it)" | 1145-1151 | `excel_handler.py:54` @47a1492 (now `:55`) stores the **whole** `EXCEL_SAFETY` dict as `self.safety_config`; no key, `save_options` included, is ever read (RAN grep) | PARTLY — the EXCEL_SAFETY wording NEVER TRUE; the rest STILL TRUE (abandonment CANNOT VERIFY) |
| C29 | 12F: requirements created edac6a3 with openpyxl ≥3.1.2, lxml ≥4.9.0; pywin32 never added; PyQt5 in c258047 (unmerged) | 1155-1162 | edac6a3 requirements (READ) matches; pywin32 added 064c634, pytest 6c1e1c6 | NO LONGER TRUE ("never added") |

### 1.13 Self-consistency notes, doc lines 1166-1182

| # | Claim | Doc line | Evidence | Verdict |
|---|---|---|---|---|
| SC1 | F6 and Gap 2 agree on the lambda bug | 1168 | both described a bug since fixed | NO LONGER TRUE (bug fixed 0890022) |
| SC2-5 | Section 8 vs 11 consistency; dead-code agreement; cache mismatch agreement | 1170-1176 | internally consistent at the time | STILL TRUE as statements about the document |
| SC6 | "12B never-modified claims verified" | 1178 | true at the time; most now changed (C24) | NO LONGER TRUE |
| SC7-8 | Corrections about `extract_filename_without_extension` and `parsing_rate` | 1180-1182 | D2, D8 | STILL TRUE |

### 1.14 What changed between "Project Audit" (fc279af) and "Project Audit 2.0" (f7eb8f1)

The f7eb8f1 diff rewrites the document (1,003 added, 677 removed; 165 unchanged lines, almost all headings/blank lines — RAN `git show f7eb8f1 | grep -c '^ '`). Version 1 (856 lines) is reachable only in history. Substantive differences, with which version is right at HEAD:

| Topic | fc279af (v1) line | f7eb8f1 (v2, current) line | Which is right (evidence) |
|---|---|---|---|
| Header | "Active Branch: main at 47a1492", no branch line (6-7) | "at fc279af", names the copilot-worktree branch (7-8) | both describe their day |
| PyQt5 commit date | 2026-02-07 (17, 782-786) | 2026-03-04 (18, 1065) | v1 right (C19) |
| edac6a3 / cbfa4a5 dates | 2025-08-31 (719-726) | 2025-10-29 (968, 976) | v1 right (C2, C5) |
| Dead code | "None detected … All files are imported and used" (85-87) | dead-code table (96-112) | v2 right (D1-D15) |
| `utils/__init__.py` | "re-exports everything" (38) | "minimal, just docstring" (45) | v1 right (S2.7) |
| Cache handoff | extractor loads from `~/.gst_organizer_cache.json` (239) | mismatch bug (317, 449-451) | v2 right on 05-03-2026 (true @47a1492); fixed since (c1583d6) |
| pywin32 absence | "code handles its absence gracefully" (504) | PQ extractor crashes on import (20, 639) | v2 right on 05-03-2026; fixed since (c463256) |
| 6237597 | "corrected reference to non-existent app.py", "fixed indentation bug" (734); helpers fix "changed import to lazy import inside function" (830) | lists the 4 real file changes (985-989) | v2 closer. The app.py claim was copied from the commit message but main.py is not in that commit (D9); the indentation: the long-name check's logic indentation is identical before/after (RAN `git show edac6a3:gui/handlers/processing_handler.py`, lines 64-75 — only a whitespace-only blank line differs); the import was already inside the function — only `utils.constants` → `.constants` changed (READ diff) |
| Configurable max length | "added in commit c258047" (483) | "6237597 made configurable via GUI" (1124) | both wrong — present from edac6a3 (C25) |
| Sanitisation | Private→Pvt "applied first" (352), no case remark | "case-sensitive!" (501) + gap 7 | v2 right on 05-03-2026 (@47a1492 `helpers.py:78-79` plain `.replace`); changed since (33d0baf) |
| Lambda capture bug | not mentioned | F6 + gap 2 (257, 637) | v2 right then; fixed since (0890022) |
| Unreachable return, os.startfile, cache mismatch, EXCEL_SAFETY | not mentioned | listed | v2 (with the corrections above) |
| Extraction mappings table | full table (332-349) | only the 6 new ones in 12A | v1 table matches `cell_mappings.py` exactly (READ) [GST RULE] |
| Long path handling | `report_processor.py` 646-755, `excel_handler.py` 275-317; `\\?\` prefix; retries (378-382) | threshold only (582) | v1 accurate @47a1492 and still true (now `report_processor.py:687-799`, `excel_handler.py:324-369`) |
| Edge cases v1 listed and v2 dropped | "Empty client/state names: validation in file_parser.py flags these as errors (343-356)" (411) | — | `validate_results()` (`core/file_parser.py:338-371`) is never called (RAN grep), so nothing flags them — v1 claim NEVER TRUE |
| v1 gaps not in v2 | No GSTIN validation (428) [GST RULE]; only top-level folder scanned (437); lower-case glob may miss upper-case extensions on Linux (438); `..` traversal (448); progress per client not per file (454); no undo (453) | — | Still: no GSTIN logic anywhere in A (RAN `grep -rni gstin`: none — see 9 exp-6); top-level only (`os.scandir`, `utils/helpers.py:328`); extension match is now lower-cased (`:330`) so upper-case extensions ARE found (RAN exp-1 `.XLSX` parsed) → that v1 gap NO LONGER TRUE (a603c02); progress is per file during copying (`processing_handler.py:266-269`) → v1 "per-client only" NEVER TRUE |
| v1 "Why Keyboard Simulation" (471) | SendKeys rationale MEDIUM | dropped | inference; CANNOT VERIFY |
| v1 recommendations | QThread/PyQt5, `RefreshAll()` instead of SendKeys, recursive scan, GSTIN validation, progress persistence, merge PQ into main app as 4th tab, structured config dir, disk check, hyphen names (677-685) | v2 deliberately does not recommend replacing Tkinter or COM (902) | Of v1's list only "PQ extractor as a fourth tab" was done in A (bfcb24a); B did the PyQt5/structured-config items (c258047, by its message) |

---

## 2. `docs/superpowers/specs/2026-07-25-merge-extractor-design.md` (design note, b276542, 25-07-2026) — plan vs what was built

The note's own code references were checked against the code **as it stood at b276542** (RAN `git show b276542:...`): `extractor_window.py:109-111` title/geometry/minsize, `:118-124` `theme_use('clam')` + TNotebook styles, `:801` `_process_clients_thread`, `:887` `update_progress` (with `update_idletasks()` at `:892`), `:894` `log_message`; `extractor_main.py:12-19` `logging.basicConfig(... FileHandler('pq_extractor.log'))`; `dark_mode_manager.py:143` four-colour whitelist, `:172` Frame/Canvas → white, `:182`/`:184` button text white, `:198-206` `else` bound to `if hasattr(...)`. **All of these references were exact** (READ `dm_b276542.py`, `ew_b276542.py`). So the note's description of the starting point is accurate; the table below is about what was then built.

| # | Plan item | Note line | What was built (evidence at HEAD) | Verdict | Commit |
|---|---|---|---|---|---|
| P1 | Context: two programs that communicate only through `~/.gst_organizer_cache.json` | 8-12 | True on 25-07 (after c1583d6). Now they also share the same window | STILL TRUE as context | c1583d6 |
| P2 | Hard constraint: app must not break; first two steps pure refactors | 19-21 | Step 1 = c1f425e ("Behaviour-preserving hardening"), Step 2 = a677670 ("Pure refactor") per their messages; a regression did occur at Step 3: hosting the extractor made the whole app fail to start without pywin32 — fixed next day (c463256 message: "Regression from hosting the extractor as Step 4 (bfcb24a)") | PARTLY (one breakage, fixed) | bfcb24a → c463256 |
| P3 | Decision: a fourth tab "Step 4: Extract" | 27 | `gui/tabs/extract_tab.py:50` "🔄 Step 4: Extract"; added in `gui/main_window.py:144-146`; `on_tab_changed` made safe for 4 tabs (`:331-338`) | BUILT as planned | bfcb24a |
| P4 | Decision: workflow choice up front on Setup | 28 | Card "WHAT SHOULD THIS RUN DO?" `gui/tabs/setup_tab.py:316-345`; values `utils/constants.py:163-172` | BUILT as planned | 510f7b4 |
| P5 | Decision: Step 4 always accessible, own folder picker | 29 | Tab always added (`main_window.py:146`); panel has its own folder entry, Browse and Scan (`extractor_window.py:196-246`) | BUILT | bfcb24a |
| P6 | Decision: `launch_extractor.py` keeps working | 30 | `launch_extractor.py` unchanged since edac6a3; `PowerQueryExtractorApp` shell (`extractor_window.py:1032-1062`); RAN `tests/test_extractor_standalone.py` passes under xvfb (Section 9 exp-0) | BUILT | a677670 |
| P7 | Split: `ExtractorPanel(ttk.Frame)` (all UI + logic, parent-agnostic) | 39 | `extractor_window.py:98-1029` | BUILT | a677670 |
| P8 | `PowerQueryExtractorApp(tk.Tk)` thin shell | 40 | `:1032-1062` — plus an extra not in the plan: `__getattr__` forwards unknown attributes to the panel (`:1051-1062`) | BUILT (+extra) | a677670 |
| P9 | `gui/tabs/extract_tab.py` hosts the panel | 41 | `gui/tabs/extract_tab.py:74-82` (`show_header=False, show_status_bar=False, auto_load=False`). Extra: guarded import with an "unavailable" notice (`:24-30, 84-105`) | BUILT (+extra) | bfcb24a, c463256 |
| P10 | Move window concerns to the shell: (1) title/geometry/minsize | 49 | `extractor_window.py:1042-1044` | BUILT | a677670 |
| P11 | (2) `theme_use('clam')` + global TNotebook styles | 50-53 | moved into `configure_extractor_styles()` (`:85-95`), called only by the shell (`:1046`) | BUILT | a677670 |
| P12 | (3) `logging.basicConfig(...)` in `extractor_main.py:12-19` "moves into the standalone entry point" | 54-56 | **Not moved**: still at module level in `power_query_extractor/extractor_main.py:12-19` (RAN `git log -G basicConfig -- power_query_extractor/` → only edac6a3). The hazard is avoided another way: `extract_tab.py` imports from `extractor_window`, not `extractor_main` (`gui/tabs/extract_tab.py:8-10, 25`). Consequence: when the extractor runs inside the main window its messages go to `gst_organizer.log`, not `pq_extractor.log` (READ `main.py:20-27`) | DIFFERS (goal met by a workaround) | bfcb24a |
| P13 | Thread safety: all worker-thread UI updates through `root.after(0, ...)` with values bound at schedule time; `update_idletasks()` removed from the worker | 60-69 | `_ui()` (`extractor_window.py:989-1010`) → `self.after(0, lambda: func(*args, **kwargs))`; `log_message`/`update_progress` marshal (`:1012-1029`); messagebox, status and button reset marshalled (`:952, 955, 960, 963`); `_apply_progress` has no `update_idletasks` (`:1016-1019`). Residual: the worker still **reads** Tk variables `wait_time_var`/`suffix_pattern_var` (`:890-891`) (reads, not updates; MEDIUM risk). Note `_ui` uses the panel's `after`, not `root.after` (same effect) | BUILT | c1f425e |
| P14 | New setting `workflow_mode`, chosen on Setup, persisted in the cache | 73 | `main_window.py:90, 366, 382-383`; saved on change `setup_tab.py:347-351` | BUILT | 510f7b4 |
| P15 | `organize` — organise + reports, then stop; `full` — also refresh PQ and extract | 75-76 | Labels `constants.py:165-170` ("Also refresh Power Query and extract the data (slower)"), help `setup_tab.py:95-101` ("…and then open each report in Excel to refresh its Power Query and pull the figures out"). **Built behaviour for `full`:** after Step 3 the app switches to Step 4, fills in the target folder and scans the client list — it does **not** start the refresh; the user must press "🚀 Start Processing" in Step 4 (`gui/handlers/processing_handler.py:136-175`, comment `:139-142`; `tests/test_workflow_mode.py:85` "test_full_pipeline_does_not_start_the_refresh"). So the on-screen description of "full" promises more than happens automatically | DIFFERS | 510f7b4 |
| P16 | "The mode does not disable Step 4" | 78-80 | Step 4 tab exists in both modes (`main_window.py:146`) | BUILT | 510f7b4 |
| P17 | "In `organize` mode Step 3 finishes with a 'Continue to extraction' button" | 80-81 | **Never built.** Organise mode only writes a log line "ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query." (`processing_handler.py:145-149`). RAN `git log -S'Continue to extraction' A/main` → only b276542 (the note itself) | NOT BUILT | — |
| P18 | Step 4 folder pre-filled from Setup's target folder, follows it until the user overrides it | 85-87 | `extract_tab.py:76, 82, 108-131`. Exception: in `full` mode the hand-off sets the Step 4 folder to the target folder **unconditionally** (`processing_handler.py:166-168`), overriding a folder the user chose in Step 4, and does not update `_pushed_folder`, so afterwards Setup changes may stop propagating (READ; MEDIUM) | BUILT (with that exception) | bfcb24a, 510f7b4 |
| P19 | "pointing at an `Annual Statement-...` folder from a previous run still works" | 87-89 | The scan looks for `Annual Statement-*` folders **inside** the chosen folder and uses the newest by modified time (`extractor_window.py:622-632`). RAN: choosing an older `Annual Statement-...` folder itself → "No Annual Statement folders found", 0 clients; choosing the parent → the newest run only (Section 9 exp-4a). An older batch can only be processed if it is the most recently modified `Annual Statement-*` folder in the chosen parent (MEDIUM — mtime can change, e.g. when a consolidated report is saved into it) | NOT AS BUILT (claim does not hold for pointing at the run folder itself) | — |
| P20 | Staging 1: thread-safety fix while still standalone | 95-96 | c1f425e (25-07 22:46) precedes the split | BUILT in order | c1f425e |
| P21 | Staging 2: extract `ExtractorPanel`; pure refactor | 97 | a677670 (25-07 23:01) | BUILT | a677670 |
| P22 | Staging 3: Step 4 tab | 98-99 | bfcb24a (26-07 20:40) | BUILT | bfcb24a |
| P23 | Staging 4: workflow mode + auto-advance from Step 3 | 100 | 510f7b4 (26-07 22:58); auto-advance = switch + scan only (P15) | BUILT (partly) | 510f7b4 |
| P24 | Staging 5: dark mode rebuilt last | 101 | aeb83cf (26-07 23:04), then c89c1d4, 725fbaf, 169c7b5 | BUILT | aeb83cf … 169c7b5 |
| P25 | Unplanned commits during the work | — | 2acffed (refresh timestamps update after a run), 2aaa2b1 (configured suffix recognised), 7347383 (shorter category folders — this changes the folder paths written into every new report's Links sheet, although the note lists "any change to how reports are produced" as out of scope), c463256 / f363946 / be6418c / a5ad487 (pywin32 handling, start-up check, install offer), bf2f1dd (scrolling fix + tooltips) | outside the plan | as listed |
| P26 | Testing: "Existing coverage: `_classify_version_reports` and `_latest_refresh_times` are unit-tested (`tests/test_extractor_helpers` equivalents live in the suite)" | 108-109 | At b276542 the suite had no test of either (RAN `git grep ... b276542 -- tests/` → nothing; test files then: test_excel_com, test_file_ops, test_file_patterns, test_helpers, test_processing_handler, test_scan, test_setup_tab_smoke, test_setup_validation). Today `_latest_refresh_times` is tested in `tests/test_refresh_naming.py:99-121` (added 2aaa2b1, after the note); `_classify_version_reports` is still untested (RAN grep); there has never been a `test_extractor_helpers` file | NEVER TRUE (at the note's date) | — |
| P27 | Smoke test: ExtractorPanel builds inside a plain `ttk.Frame` | 113 | `tests/test_extractor_panel.py:36-52` | BUILT | a677670 |
| P28 | Smoke test: standalone shell builds and hosts the panel | 114 | `tests/test_extractor_standalone.py` (6 tests, run in a subprocess) | BUILT | a677670 |
| P29 | Smoke test: main app builds with four tabs, Step 4 reachable | 115 | `tests/test_main_window_tabs.py` (7 tests, incl. "test_has_four_steps_ending_with_extract") | BUILT | bfcb24a |
| P30 | Unit tests: workflow-mode gating as pure logic, no Tkinter | 116-117 | `tests/test_workflow_mode.py` (7 tests, fakes only, no Tk import — READ `:1-69`) | BUILT | 510f7b4 |
| P31 | Thread-safety test: assert worker-thread updates are *scheduled* rather than applied directly, "same approach as the existing `_log_async` tests" | 118-119 | No test asserts "scheduled, not applied"; `test_extractor_panel.py:70-77` only checks that messages appear after `root.update()`. There are no `_log_async` tests (RAN grep: none at b276542 or HEAD; `tests/test_processing_handler.py` only tests the re-entrancy guard) | PARTLY BUILT; the "existing `_log_async` tests" NEVER TRUE | — |
| P32 | Manual check per step | 121 | — | CANNOT VERIFY | — |
| P33 | Out of scope: SendKeys refresh and fixed `time.sleep` waits stay | 125-126 | unchanged (`report_processor.py:253-330`) | AS PLANNED | — |
| P34 | Out of scope: `process_files_thread` god-method refactor | 127 | still one method `processing_handler.py:182-530` | AS PLANNED | — |
| P35 | Out of scope: any change to how reports are produced | 128 | see P25 (7347383 changed folder paths in new reports) | not fully respected | 7347383 |
| P36 | Dark-mode defect: `_restore_widget_colors` forces Frames/Canvas white; originals only for Labels/Buttons | 136-139 | replaced by exact capture/replay (`dark_mode_manager.py:199-268`) | FIXED | aeb83cf |
| P37 | `update_tree_tags` else-binding bug | 140-142 | fixed (`:271-294`) | FIXED | aeb83cf |
| P38 | Buttons restored with `fg='white'` | 143 | fixed (replay) | FIXED | aeb83cf |
| P39 | Labels only darkened if bg in a 4-item whitelist | 144-145 | generic option discovery (`:191-197, 223-253`) | FIXED | aeb83cf |
| P40 | Only TFrame/TLabel/TNotebook/Treeview ttk styles handled | 146-147 | 11 styles (`:37-49`) | FIXED | aeb83cf |
| P41 | "Widgets created after a toggle are never themed" | 148 | **Still true.** Nothing re-applies dark mode when new widgets are built (only `toggle_dark_mode`/`apply_cached_values` call it, `main_window.py:157-163, 388-389`). RAN: a Label created after `apply_dark_mode()` keeps `bg='white'` (Section 9 exp-4b); `tests/test_dark_mode.py:129-141` itself calls `apply_dark_mode()` a second time to theme the late widget. (c89c1d4's message says "widgets created during dark mode are themed" — not what happens without a second toggle.) Exceptions: the Setup checklist and accent bars, which recolour themselves for dark mode (`setup_tab.py:542-608`, 169c7b5) | STILL PRESENT | — |
| P42 | "both failure paths use bare `except: pass`" | 148-149 | `dark_mode_manager.py` now catches `tk.TclError`/`TypeError`/`Exception` with debug logging (`:175-189, 199-215, 255-268`), no bare `except:` (RAN grep) | FIXED | aeb83cf |

Additional plan-vs-built notes:
- bfcb24a's message: "the panel keeps its own Setup/Processing sub-tabs for now, i.e. a notebook inside a notebook. Flattening is the next decision" — still nested at HEAD (`extractor_window.py:163-175`). Not in the note; not done.
- The Step 4 "unavailable" notice says "Steps 1-3 are unaffected: organising files and creating the reports works without pywin32" (`gui/tabs/extract_tab.py:101-105`) and `report_processor.py:11-13` calls pywin32 "genuinely optional". Since be6418c (same day, later) the app refuses to start on Windows without pywin32 (`utils/requirements_check.py:44-60`, `main.py:97-102`), so on Windows those two texts are out of date (READ; HIGH).

---

## 3. `tests/README.md` (6c1e1c6, 23-07-2026)

| # | Claim | Line | Evidence | Verdict | Commit |
|---|---|---|---|---|---|
| T1 | Regression tests for data-integrity logic (misfiling) | 3-4 | Suite now also covers GUI, dark mode, extractor, start-up checks: 18 test files, 288 tests (RAN full `pytest` under xvfb: 1 failed + 287 passed = 288) | NO LONGER complete | 322717f onward |
| T2 | Run `pytest` from the repository root (or `python -m pytest`) | 8-14 | `pytest.ini:1-6` (`testpaths = tests`, `addopts = -q`); RAN works (Section 9 exp-0) | STILL TRUE | — |
| T3 | Requires pytest — `pip install -r requirements.txt` | 14 | `requirements.txt:15` | STILL TRUE | 6c1e1c6 |
| T4 | `test_file_patterns.py`: every report type, case-insensitivity, `.xls`, `(N)` markers (trailing + GSTR-3B leading) | 20 | `tests/test_file_patterns.py:34, 57, 63, 71, 90` | STILL TRUE | — |
| T5 | `test_helpers.py`: sanitize_filename, get_state_code, create_client_state_key, validate_client_name, format_size, format_duration, path helpers | 21 | `tests/test_helpers.py:24-181` | STILL TRUE | — |
| T6 | `test_file_ops.py`: validate_excel_file, find_excel_files, get_file_info against real temp files | 22 | plus `scan_excel_files`, cached-stat equivalence, parallel branch (`tests/test_file_ops.py:76-150`) | STILL TRUE (now more) | 56aaad9, 4f84ba9 |
| T7 | `test_scan.py`: end-to-end scan_folder (grouping, missing, variations, invalid skipping) | 23 | `tests/test_scan.py:18-80` | STILL TRUE | — |
| T8 | The table lists the test files | 16-23 | 14 more exist: test_dark_mode, test_dependency_installer, test_excel_com, test_extractor_panel, test_extractor_standalone, test_main_window_tabs, test_optional_pywin32, test_processing_handler, test_refresh_naming, test_requirements_check, test_scroll_bindings, test_setup_tab_smoke, test_setup_validation, test_workflow_mode | NO LONGER complete | 33d0baf … bf2f1dd |
| T9 | Characterisation tests; `NOTE` comments mark known quirks, "e.g. the case-sensitive `Private`→`Pvt` substring replace mangling `Privateer`" | 27-31 | That quirk was removed the same day: `tests/test_helpers.py:43-45` now asserts `"Privateer Unlimited"` stays `"Privateer_Unlimited"`; RAN exp-1 confirms. Remaining NOTEs: `:70` (run-together names), `:96` ('Unknown Place'→'UP' collides with Uttar Pradesh) [GST RULE], `:163` (hours format drops seconds) | NO LONGER TRUE (the example) | 33d0baf |
| T10 | "No GUI, Excel COM, or network access is exercised, so the suite runs fast (<1s) and anywhere." | 33-34 | The four documented files: 95 tests in 0.30 s (RAN). Whole suite: GUI tests exist (Tk windows built in-process and in subprocesses). RAN under xvfb: 287 passed, 1 failed in 3.57 s — the failure `tests/test_setup_validation.py:66` expects `'itc.xltx'` from a Windows path `C:\gst\itc.xltx` and gets the whole string on Linux. RAN without a display: 241 passed, 14 failed (13 GUI tests that need a display + the same path test), 33 skipped, 1.40 s | NO LONGER TRUE (true only for the four original files) | 322717f, a677670, bfcb24a, aeb83cf, bf2f1dd |
| T11 | (commit message) "81 tests" | — | the four original files now hold 95 tests (RAN) | NO LONGER TRUE (count) | 33d0baf, 4219ffb, 56aaad9, 4f84ba9 |

---

## 4. `README.md` (d674386, 29-10-2025; identical in A and B) — checked against A at HEAD

| # | Claim | Line | Evidence (A HEAD) | Verdict | Commit |
|---|---|---|---|---|---|
| RM1 | "GST File Organizer v3.0 - Production Ready"; "Version 3.0 Production Ready, Stable Release, Last Updated 29th October 2025, Author: Advanced AI Assistant" | 1, 361-364 | version strings unchanged (`constants.py:134`, `status_bar.py:29`, `main.py:2-6`); but A's code changed in 32 commits up to 27-07-2026 | "Last Updated" NO LONGER TRUE | July commits |
| RM2 | "enterprise-grade", "Bulletproof" | 3, 28 | wording | CANNOT VERIFY (opinion) | — |
| RM3 | "All file operations use temporary files and verification steps" / "Safe file operations: All Excel operations use temporary files" | 7, 42, 260 | Client files are copied straight to their destination with `shutil.copy2` + a size check, no temp file (`utils/helpers.py:196-208`). ITC/Sales reports: staged in a local temp file only since 12d505f (`core/excel_handler.py:147-215, 269-322`). The summary workbook (`excel_handler.py:402`) and the consolidated workbook (`data_consolidator.py:179`) are saved directly. The refresh works on a copy of the report (`report_processor.py:164`) | NEVER TRUE as "all" (partly true for reports since 12d505f) | 12d505f |
| RM4 | Overview/features: organise, ITC & Sales reports from templates, PQ refresh (Windows), extraction into consolidated sheets, batch, dark mode, session memory, 5-sheet summary | 11-29 | all exist (Sections 1.3, 1.10) | STILL TRUE | — |
| RM5 | Six file types; GSTR-3B "supports multiple monthly filings" | 31-37 | `constants.py:68-75`; `file_parser.py:264-268` | STILL TRUE [GST RULE] | — |
| RM6 | "Verification: Files are verified before finalization" | 43, 261, 310 | copies: size check (`helpers.py:202-208`); reports: only "does the output file exist" (`excel_handler.py:218-224, 310`) | PARTLY | — |
| RM7 | "Clean saves: No VBA macros or external links preserved" | 44, 262 | The intended settings (`EXCEL_SAFETY['save_options']` keep_vba/keep_links False, `constants.py:240-244`) are never used. The openpyxl path loads with `keep_links=True` (`core/excel_handler.py:282`); the COM path saves whatever the template contains (`:203`) | NEVER TRUE | — |
| RM8 | "Robust sanitization"; "Backup creation: Original files are never modified" | 45, 47, 306-307 | sanitisation `helpers.py:68-126`; sources are only read/copied (READ organiser); backups only in fresh mode when the destination exists (`file_organizer.py:277-282`) | STILL TRUE | — |
| RM9 | Python 3.7+ (3.8+ recommended); Windows best, macOS, Linux | 51-52 | S1.2. On Linux the main window starts (RAN: `tests/test_main_window_tabs.py` passes under xvfb). macOS CANNOT VERIFY | STILL TRUE (Linux RAN) | — |
| RM10 | 4GB RAM, 500MB, Excel 2016+ | 53-55 | — | CANNOT VERIFY | — |
| RM11 | Step 1: install Python with "Add to PATH" | 59-66 | general | CANNOT VERIFY | — |
| RM12 | Step 2 folder layout: main.py, launch_extractor.py, requirements.txt, core/(4 files), gui/(handlers, tabs, utils, widgets), utils/(`__init__`, constants, helpers), power_query_extractor/(core, gui, config) | 68-95 | `utils/` now also **needs** `excel_com.py` (imported by `core/excel_handler.py:41` and `report_processor.py:39`), `requirements_check.py` (`main.py:68`) and `dependency_installer.py` (`main.py:82`); `gui/tabs/extract_tab.py` is part of tabs. Copying only the listed utils files makes `main.py` stop with "Could not start" (`main.py:114-127`) (READ). (`gui/__init__.py` is not listed, but `main.py:39-43` creates a placeholder if missing) | NO LONGER TRUE | ba2a753, f363946, a5ad487 |
| RM13 | Step 3: `pip install -r requirements.txt` | 97-102 | still the way; it now installs openpyxl, pywin32 (Windows), lxml **and pytest** (`requirements.txt:5-15`) | STILL TRUE (installs more) | 6c1e1c6, 064c634 |
| RM14 | "Requirements.txt Contents: `openpyxl>=3.0.0` / `pywin32>=305; sys_platform == 'win32'`" | 104-108 | Real file: `openpyxl>=3.1.2` (since edac6a3), `pywin32>=305; sys_platform == 'win32'` (only since 064c634), `lxml>=4.9.0`, `pytest>=7.0.0` (READ `requirements.txt`, `git show edac6a3:requirements.txt`) | NEVER TRUE (openpyxl version, lxml missing); pywin32 line became true 23-07-2026 | 064c634 |
| RM15 | Step 4: `python main.py` | 110-113 | `main.py:136-137`. New: start-up dependency report and install offer (`main.py:66-106`) not mentioned | STILL TRUE (incomplete) | f363946, a5ad487 |
| RM16 | Templates: .xlsx or .xltx; "must have a 'Links' sheet" | 119-122 | File dialogs offer `*.xltx *.xlsx` (`gui/handlers/file_handler.py:31, 42`); the checklist also accepts .xltm/.xlsm (`gui/utils/setup_validation.py:21`). A missing Links sheet is not enforced (G14) | PARTLY | — |
| RM17 | Naming table and examples (`GSTR-2B-Reco-{Client}-{State}-{Period}.xlsx` … `AnnualReport-{Client}-{State}-{Year}.xlsx`); "MUST match exactly" | 124-133 | RAN all six README examples parse to the right types (exp-1). Matching is case-insensitive, and accepts a leading "(N) " for GSTR-3B and a trailing " (N)" for all types (IMS since 5c8ce95) — so "exactly" is stricter than the code | STILL TRUE (examples) [GST RULE] | 5c8ce95 |
| RM18 | Setup tab: Browse Source; ITC template; Sales template; Target; choose Fresh / Re-Run / Resume ("Resume: Continue interrupted processing"); click "Scan Files" | 137-146 | Steps exist (`gui/tabs/setup_tab.py:353-391, 401-412, 459-462`). Missing: the workflow choice card (510f7b4). "Resume" cannot pick up a run started in an earlier minute (R17, RAN) | PARTLY | 510f7b4 |
| RM19 | Validation tab: summary; ✅ all files / ⚠️ missing; checkbox selection; Select All, Clear All, Complete Only | 148-155 | `validation_tab.py`, `client_handler.py`, status icons `file_parser.py:278-286` (⚠️ also for "Has duplicates") | STILL TRUE | — |
| RM20 | "Review variations tab for naming issues" | 156 | There is no variations tab (never was: RAN `git grep -i variation 47a1492 -- gui/tabs/` → nothing). Variations appear only as a count in the scan summary (`gui/main_window.py:320-321`) and as the "Variations" sheet of the summary workbook (`core/excel_handler.py:554-585`) | NEVER TRUE | — |
| RM21 | "Step 3: Processing Tab — 1. Click 'Dry Run' to preview" | 158-159 | The Dry Run button is on the Validation tab (`gui/tabs/validation_tab.py:109-118`); the Processing tab has only Start/Stop (`gui/tabs/processing_tab.py:64-87`) | NEVER TRUE | — |
| RM22 | "2. Click Start Processing; 3. Monitor progress bar and log; 4. Wait for completion message" | 160-162 | `processing_tab.py:64-74`; completion box `processing_handler.py:498-502` | STILL TRUE | — |
| RM23 | Power Query Extractor: run `python launch_extractor.py` after processing | 164-169 | works (P6); additionally available as Step 4 inside the main window (not in README) | STILL TRUE (incomplete) | bfcb24a |
| RM24 | "The extractor auto-loads target folder from cache" | 171 | Before c1583d6 it looked in the wrong place (G1), so this did not work from d674386 until 23-07-2026. Standalone now: cache read (`extractor_main.py:36-63`) → folder set and scanned (`extractor_window.py:135-137`). Inside the main window (Step 4) it is pre-filled from Setup but **not** scanned until the user presses Scan (`extract_tab.py:71-79`, `auto_load=False`) | NEVER TRUE until c1583d6; now TRUE standalone, PARTLY in Step 4 | c1583d6, bfcb24a |
| RM25 | "Click 'Refresh & Extract'" | 173 | The button is "🚀 Start Processing" (`extractor_window.py:328-340`); no "Refresh & Extract" text ever existed in code (RAN `git log --all -S'Refresh & Extract'` → only d674386 README and f7eb8f1 audit) | NEVER TRUE | — |
| RM26 | Output tree: `Annual Statement-DDMMYY HHMM/`, `GST_Processing_Summary_DDMMYYYY_HHMMSS.xlsx`, `ClientName-StateCode/Version-DDMMYY HHMM/` | 178-184 | folder names STILL TRUE (`constants.py:82-84`); summary file name is `GST_Processing_Summary_DDMMYY_HHMM.xlsx` (`processing_handler.py:186, 462-464`, timestamp `'%d%m%y %H%M'` → `_`) | summary name NEVER TRUE | — |
| RM27 | Category folders `GSTR-3B Exports (ClientName)/`, `Other ITC related files (ClientName)/`, `Sales related files (ClientName)/` | 185-187 | `GSTR3B (…)`, `ITC (…)`, `Sales (…)` (`constants.py:91-95`), and without "(ClientName)" unless enabled (default off, `main_window.py:92`) (RAN exp-2) | NO LONGER TRUE | 7347383 |
| RM28 | `ITC_Report_ClientName_State_DDMMYYYY_HHMMSS.xlsx`, `Sales_Report_…` | 188-189, 202-203 | real: `ITC_Report_{sanitised client}_{state CODE}_{DDMMYY_HHMM}.xlsx` (`processing_handler.py:292-298, 355-356`) | NEVER TRUE (date format; "State" is the code) | — |
| RM29 | `_Organization_Report_ClientName_State.txt` in the Version folder | 190, 204 | `core/file_organizer.py:330-331` — uses the raw client name and the full state **name** | STILL TRUE | — |
| RM30 | (missing from tree) Annual Report file at Version-folder root; refreshed copies `…_Refreshed_DDMMYY_HHMM.xlsx` next to the reports; `PQ_Extraction_Report_YYYYMMDD_HHMM.xlsx` at Level 1 | — | `file_organizer.py:311`; `report_processor.py:128-130`; `data_consolidator.py:175-177` | README incomplete | — |
| RM31 | Summary workbook sheets: Summary, Client Status, File Mapping, Errors, Variations | 195-200 | `excel_handler.py:391-395` | STILL TRUE | — |
| RM32 | Template cell mappings ITC B2,B4,B5,B7,B8,B10,B11; Sales B2,B3,B5,B6,B8,B9 | 208-227 | `constants.py:102-126` (B4 = Version folder; B7 & B10 = ITC folder) | STILL TRUE | — |
| RM33 | Cell extraction example (`'GSTR-3B Difference'`, ITC, `'Diff GSTR-3B'`, `J15`, `'Diff IGST'`) and `EXTRACTION_OPTIONS` (skip_refresh False, auto_refresh_if_missing True, show_timestamp False, missing_data_text 'Missing') | 229-253 | `power_query_extractor/config/cell_mappings.py:33-58, 248-253` exact. Not said: `auto_refresh_if_missing` and `show_timestamp` are never read (RAN grep: only defined); the UI checkbox overwrites `skip_refresh` in memory (`extractor_window.py:483-487`); and some configured cells are not the values that reach the output (discrepancy X1) | STILL TRUE as text [GST RULE] TIME-VARYING | — |
| RM34 | Troubleshooting: "No module named 'win32com'" → `pip install pywin32` + `python Scripts/pywin32_postinstall.py -install` as admin | 275 | `main.py` now checks for `pythoncom` itself and offers the pip install (`utils/requirements_check.py:44-60`, `utils/dependency_installer.py`); the post-install step CANNOT VERIFY | STILL USEFUL / partly superseded | a5ad487 |
| RM35 | "No Excel files found" → check extensions (.xlsx, .xls) | 276 | the GUI's message is "No valid GST files found" (`file_handler.py:90-91`); .xlsm also scanned (`helpers.py:327`); invalid/too-small files are dropped silently (E4) | PARTLY | — |
| RM36 | "Template not found" → "Ensure templates have Links sheet" | 278 | the template-missing messages are "ITC/Sales template file not found" (`file_handler.py:143-149`) and are about the file path; a missing Links sheet produces no error (G14) | NEVER TRUE (mismatched advice) | — |
| RM37 | Debug mode: add `logging.basicConfig(level=logging.DEBUG)` in main.py or launch_extractor.py | 283-287 | `main.py:20-27` already calls `basicConfig(level=INFO, ...)` at import; a second call does nothing unless the first is changed (standard logging behaviour; MEDIUM — not RAN). In `launch_extractor.py` a call placed before the import of `extractor_main` would win and stop `pq_extractor.log` being created (`extractor_main.py:12-19`) (MEDIUM) | PARTLY | — |
| RM38 | Logs: `gst_organizer.log` (main), `pq_extractor.log` (extractor) | 289-291, 333 | `main.py:24`, `extractor_main.py:16`; both relative to the folder the program is started from. Step 4 messages go to `gst_organizer.log` (P12) | PARTLY (Step 4 case) | bfcb24a |
| RM39 | Best practices (backup, test small, close Excel, …) | 293-301 | advice. "Close all Excel files before processing" matters more than it says: report creation reuses a running Excel and then calls `Quit()` on it (G12) | CANNOT VERIFY (advice) | — |
| RM40 | "No data is sent outside your computer" | 305 | No network calls with client data (RAN grep: no `requests`/`urllib`/`socket`/`http` in source). Since a5ad487 the app can, if the user agrees, run `pip install` (`utils/dependency_installer.py:41-60`), which downloads packages from the internet | STILL TRUE for client data; NO LONGER literally true | a5ad487 |
| RM41 | "All operations are logged"; "Extensive error checking"; "Multiple verification steps" | 308-310 | mostly logged; see RM6 | PARTLY | — |
| RM42 | Known limitations (Excel closed; templates structure; patterns exact; 260-char limit; PQ needs Windows+Excel) | 314-318 | The code itself treats ~218 characters as the practical limit (`constants.py:85-90`, `excel_handler.py:356`, `report_processor.py:720`) | PARTLY (218 vs 260) | — |
| RM43 | "Recent Bug Fixes (Latest Commit)": circular import (STATE_CODE_MAPPING); scanned_files attribute; thread safety; "client name validation logic indentation"; progress callback; "Enhanced Excel corruption prevention" | 320-327 | Matches 6237597's message, not fully its diff: import path, scanned_files line, partial `root.after` wrapping and `_update_progress` are real (READ); the "indentation" fix and any Excel-corruption change are not in the diff (6237597 touches no Excel code; the long-name block's logic indentation is identical in edac6a3 — RAN `cat -A` of `processing_handler.py:64-75` @edac6a3); "(Latest Commit)" is long out of date | PARTLY / NEVER TRUE for the last two items | — |
| RM44 | Support / reporting / licence / conclusion | 329-357 | — | CANNOT VERIFY (non-technical) | — |

### 4.1 README instructions that are wrong for A today (for a user following them)

1. Folder layout to copy (lines 68-95) omits `utils/excel_com.py`, `utils/requirements_check.py`, `utils/dependency_installer.py` and `gui/tabs/extract_tab.py` → the app will not start (RM12).
2. `requirements.txt` contents shown (104-108) do not match the real file (RM14).
3. There is no "variations tab" (156) (RM20).
4. "Dry Run" is on the Validation tab, not the Processing tab (158-159) (RM21).
5. The extractor button is "🚀 Start Processing", not "Refresh & Extract" (173) (RM25).
6. Output names: category folders are `GSTR3B`, `ITC`, `Sales` (185-187); report and summary file names use `DDMMYY_HHMM` and the state *code* (182, 188-189, 195, 202-203) (RM26-RM28).
7. "Template not found → ensure Links sheet" (278) points at the wrong cause (RM36).
8. Not documented at all: the Step 4 tab, the workflow choice (organise-only / full), the start-up dependency check and install offer, the Setup checklist and tooltips.
9. "Resume: Continue interrupted processing" (145) does not behave as a resume across minutes (R17).

---

## 5. `.gitignore` (80 lines) — checked against what A actually writes

| # | Pattern(s) | Line | What A writes / reads (evidence) | Verdict |
|---|---|---|---|---|
| GI1 | `__pycache__/`, `.pytest_cache/`, `*.py[cod]` … build/dist/eggs | 2-23 | Python byte-code; `.pytest_cache/` from pytest (added 6c1e1c6). Note A-copy currently contains `__pycache__` folders created by other readers' runs (not in git) | STILL APPLICABLE |
| GI2 | venv/IDE patterns | 25-37 | generic | n/a |
| GI3 | `*.log`, `gst_organizer.log`, `pq_extractor.log` | 40-42 | logs written to the start folder (`main.py:24`, `extractor_main.py:16`); `gst_organizer.log` was committed in edac6a3 before this file existed and deleted in cbfa4a5 (READ `--stat`) | STILL APPLICABLE |
| GI4 | `.gst_organizer_cache.json` | 43 | the cache is written to the **home folder** (`gui/utils/cache_manager.py:15`), never inside the repo, so this pattern only matters if someone copies it in | harmless / not needed |
| GI5 | `*.cache`, `.cache/` | 44-45 | `extractor_main.py:43` *reads* `.cache/app_state.json` (relative to the start folder) as a fallback; nothing in A writes it (RAN grep) | consistent (read-only fallback) |
| GI6 | `temp/`, `logs/` | 46-47 | `main.py:35-37` creates `logs/` and `temp/` in the program folder on every start; nothing writes into them (RAN grep, G9) | STILL APPLICABLE |
| GI7 | `~$*.xlsx`, `~$*.xls`, `*.tmp`, `temp_gst_*`, `backup_*` | 50-54 | `temp_gst_` / `backup_` are only dead constants (`constants.py:237-238`); real backups are named `<stem>_backup_<timestamp><ext>` (`utils/helpers.py:225`) — not matched by `backup_*` — and are written in the target folder, not the repo; the report staging folder is `%TEMP%\gst_excel_temp` (`excel_handler.py:53`) | `temp_gst_*`/`backup_*` do not match anything A creates |
| GI8 | `/Annual Statement-*/`, `/output/`, `/target/` | 57-59 | output only lands in the repo if the user picks the repo as target folder (`file_organizer.py:165-173`) | protective |
| GI9 | `/mnt/user-data/`, `*.bak`, `config/last_session.json` | 62-64 | `extractor_main.py:44` reads `config/last_session.json` as a fallback; nothing writes it; nothing in A uses `/mnt/user-data/` (RAN grep) | consistent / unexplained (`/mnt/user-data/`) |
| GI10 | `test_files/`, `test_data/`, `*.test.xlsx` | 67-69 | the test suite uses pytest temporary folders (`tests/conftest.py`, `tmp_path`), not these | not needed by the suite |
| GI11 | Windows and Power Query cache patterns | 72-80 | generic | n/a |
| GI12 | (absent) `.claude` | — | A tracks `.claude/settings.local.json` (bd176bf); B ignores `.claude` and untracked it (c258047) (RAN `git diff A/main B/main -- .gitignore`) | difference A vs B |
| GI13 | (absent) client Excel files in general | — | nothing stops a client `.xlsx` dropped in the repo root from being committed (only `/Annual Statement-*/` folders, `*.test.xlsx` and temp/lock files are ignored) | observation (no claim) |

---

## 6. HANDOVER.md / FILE_INDEX.md

RAN `git -C /root/merge-scratch/combined log --all --name-only | grep -i -E 'handover|file_index'` → no output (exit status 1). Neither file exists now or anywhere in either project's history. The audit method (`Project_Audit_FINAL.md:11-12`) therefore has no developer-diary documents to compare against; the documents in this task are the only written statements of intent.

---

## 7. High-value discrepancies (most important first)

**X1 — The consolidated workbook puts the wrong figures in 9 columns. [GST RULE]** (not mentioned by any document; READ + RAN, HIGH)
`DataConsolidator.create_report` fills each output column by searching the extracted values for the first field whose name *contains* the column title (`power_query_extractor/core/data_consolidator.py:139-145`), and only falls back to the exact configured field if nothing was found (`:147-150`). The extracted values are stored sheet by sheet in configuration order, so the first ITC field containing "IGST" is always the "Diff GSTR-3B" one (field `GSTR-3B_Difference_Diff_IGST`, cell J15). Columns whose title is a single word — "IGST", "CGST", "SGST" — therefore receive J15/K15/L15 of sheet "Diff GSTR-3B" instead of their configured cells: sheet "IMS Reco" (configured IMS_Reco!AA1/AD1/AG1), "Extra In Purchase" (ExtraInPurchase!O1/P1/Q1), "Extra In GSTR2B" (ExtraInGSTR2B!S1/T1/U1). RAN with synthetic numbers (Section 9 exp-3): all 9 MISMATCH, all other 13 columns OK. Titles containing spaces ("Diff IGST", "IGST Diff", "Table 6J Match", …) never match the underscore field names and fall through to the correct exact lookup. The wrong value appears whenever the Diff GSTR-3B cell holds any value, including 0; only if that cell is empty does the correct value show. Present since the three sheets were added in 47a1492 (30-10-2025); `data_consolidator.py` is unchanged since edac6a3. This describes what the code does; whether any figure is right for a return is for the author.
Related: when a report's refresh is not marked successful (e.g. the BB2 check fails), every configured column for that report shows "Missing" (`data_consolidator.py:128-129`), although values were extracted and the log says "(Data extracted despite error)" (`extractor_window.py:924-925, 935-936`).

**X2 — Creating a report asks the user's own running Excel to quit.** (READ, HIGH on the logic; behaviour on Windows CANNOT VERIFY)
`core/excel_handler.py:123-128` reuses an Excel that is already open; the `finally` block's guard `if self.excel_app is None:` (`:242`) is always true because `self.excel_app` is only ever set to None (`:56`); so `excel.Quit()` runs after every report, contrary to the comment "Don't quit if Excel was already open" (`:241`). Present @47a1492. The March audit (G12) noted only the reuse.

**X3 — A template without a "Links" sheet produces no error.** (READ + RAN)
COM path creates an empty "Links" sheet and reports success (`excel_handler.py:180-183, 218-221`); openpyxl path writes nothing and returns True (`:295-310`; RAN exp-2). The audit (G14) says the error "surfaces during report creation"; the README tells users templates "must" have it. No check anywhere before processing (`file_handler.py:132-151`, `setup_validation.py:56-64`).

**X4 — "Resume" cannot resume an interrupted run.** (READ + RAN, HIGH)
Every run creates `Version-<current minute>` (`core/file_organizer.py:102-108`); resume only skips files already present in *that* folder (`:283-287`). RAN: resume a minute later copies everything again into a new Version folder. README line 145 and the audit (R17, E7) describe it as continuing interrupted work.

**X5 — Invalid or too-small Excel files vanish silently.** (READ + RAN)
They are filtered out during listing (`utils/helpers.py:343-359`) and appear neither in variations nor in the counts. The audit (E4) says they are "added to variations" — that branch was unreachable even @47a1492 and was removed in a603c02. RAN exp-2.

**X6 — The main processing thread still touches the window directly.** (READ; effect on Windows MEDIUM)
The audit says "All UI updates … use root.after" (F8.4, PC3) — never true. At HEAD: 45 direct `log_message` calls (RAN count), two `update_progress` calls, `messagebox.showinfo`/`showerror` and button configuration from the worker (`gui/handlers/processing_handler.py:319-530`). The July thread-safety work (c1f425e) was done for the extractor only; the design note calls this pattern "a genuine cause of freezes and crashes" for the extractor (note lines 60-64).

**X7 — The design note and the build differ in four places.** (Section 2)
'full' mode does not start the refresh (P15); no "Continue to extraction" button in organise mode (P17); extractor logging set-up not moved, avoided by import path (P12); pointing Step 4 at an older `Annual Statement-…` folder finds nothing (P19, RAN). Also its claimed existing test coverage did not exist (P26, P31).

**X8 — README instructions wrong for A today** — nine items, Section 4.1.

**X9 — Widgets built while dark mode is on stay light** (P41, RAN) — the design note listed it as a defect; still present; c89c1d4's message says otherwise.

**X10 — tests/README.md** — "no GUI … runs anywhere (<1s)" no longer true; one test fails on Linux, 13 more fail without a display (T10, RAN).

**X11 — Commit message vs diff (6237597).** Claims "Corrected non-existent app.py reference in main.py" and "Fixed indentation bug"; main.py is not in the commit and the long-name block's indentation is unchanged (D9, C7, RM43). The first audit (fc279af line 734) and the README "Recent Bug Fixes" (320-327) repeat the message. The `app.py` reference was actually removed in f363946 (26-07-2026).

**X12 — pywin32 is mandatory on Windows, but two texts still call it optional** (`gui/tabs/extract_tab.py:101-105`, `report_processor.py:11-13` vs `utils/requirements_check.py:44-60`, be6418c).

**X13 — Period information in file names is never used. [GST RULE]** The month/period/year/date parts captured by the patterns are stored (`core/file_parser.py:105-106`) and read nowhere (RAN grep); e.g. which GSTR-3B months a client has is not checked. Described, not judged.

**X14 — Things the March audit missed that July fixed:** per-client folder-name toggles silently ignored (33d0baf); newest Version folder picked by name, wrong across months (c1583d6); refresh detection hard-coded to `_Refreshed_` (2aaa2b1); refresh timestamps not updated after a run (2acffed); fourth-tab IndexError (bfcb24a); scrolling killed by `unbind_all` (bf2f1dd) — each from the commit message and diff (READ).

**X15 — Dates and branch facts in the March audit** — PyQt5 commits dated 2026-03-04 instead of 2026-02-07; edac6a3/cbfa4a5 dated 2025-10-29 instead of 2025-08-31; the "unmerged branch" is project B (H5, S1.5, C2, C5, C19). The first version (fc279af) had these dates right; v2 introduced the errors.

**X16 — 'full' hand-off overrides a Step 4 folder the user chose** (`processing_handler.py:166-168`; P18, READ, MEDIUM).

**X17 — Small inconsistencies that remain:** Level-2 folder names keep "Private"/"Limited" for underscore-joined names while file names abbreviate (R12, RAN); `create_client_state_key` drops the state code if the limit leaves ≤5 characters (R12, RAN); state "M"-style warnings are worded as "Client name too short" and never shown (G8, E12).

---

## 8. Items tagged [GST RULE] in these documents (described only — no judgement)

| Rule | Where in code | What the code does | Doc reference | Note |
|---|---|---|---|---|
| File classification by name | `utils/constants.py:13-62` | Six regexes decide whether a file is GSTR-2B Reco, IMS Reco, GSTR-3B Export, Sales, Sales Reco or Annual Report, and which folder it goes to (`folder` key; `core/file_organizer.py:305-312`) | audit R1-R7; README 124-133 | IMS pattern changed 5c8ce95 |
| Required documents per client | `utils/constants.py:68-75`; `core/file_parser.py:252-286` | A client is "Complete" only with all six types; GSTR-3B may have any number of files; any other type with 2+ files → "Has duplicates" | audit R8-R9; README 31-37 | — |
| "Critical" files | `core/file_parser.py:359` (`['Annual Report', 'GSTR-3B Export']`) | Only inside `validate_results()`, which is never called | (v1 audit 411) | dead code |
| State name → two-letter code | `utils/constants.py:251-292`; `utils/helpers.py:592-618` | 39 name entries (36 distinct codes; RAN count of `constants.py:253-291`) incl. alternatives (orissa, pondicherry, NCT of Delhi, merged DNH-DD) → 2-letter codes used in folder names and keys; unknown names → initials of first two words or first 3 letters ('Unknown Place' → 'UP', same as Uttar Pradesh's entry) | audit R10-R11, F1.8 | TIME-VARYING (list of states/UTs, names) |
| Period fields in names | `utils/constants.py:15-55` groups; `core/file_parser.py:105-106` | Captured (period, DDMMYYYY date, month, start/end month, year) but never used | audit R5-R6 | X13 |
| Which sheet/cell each extracted figure comes from | `power_query_extractor/config/cell_mappings.py:33-228` | GSTR-3B Difference ← 'Diff GSTR-3B' J15/K15/L15; GSTR-1 Difference ← 'Diff - GSTR-1' N1/O1/P1/Q1; Turnover GSTR-1 ← 'Diff - GSTR-1' I1; Reco Diff ← 'Reco' C37/D37/E37; Table 6J Match ← 'T6 of R9' K29; Table 8D Limits ← 'T8 of R9' G6; IMS Reco ← 'IMS_Reco' AA1/AD1/AG1; Extra In Purchase ← 'ExtraInPurchase' O1/P1/Q1; Extra In GSTR2B ← 'ExtraInGSTR2B' S1/T1/U1 | audit C17, fc279af 332-349; README 229-253 | TIME-VARYING (table numbers "6J", "8D", sheet names "T6 of R9"/"T8 of R9" — the fc279af audit (line 780) reads these as GSTR-9 tables; that is the audit's interpretation, not checked) |
| Which figure actually reaches the consolidated workbook | `power_query_extractor/core/data_consolidator.py:124-165` | substring match → 9 columns take Diff GSTR-3B values (X1); refresh not successful → "Missing" for all of that report's columns; numbers formatted `#,##0.00` unless 0 (`:161-162`) | not in any document | X1 |

No GSTIN, PAN, rate, threshold, rounding, financial-year or due-date logic exists anywhere in A's source (RAN `grep -rni "gstin\|\bpan\b"` → nothing). The only financial-year text found in my scope is inside commit 7347383's message (a share folder name) — not code.

---

## 9. Experiments run (all in `/root/merge-scratch/run/A-D/`, against a copy `A-run/` of A-copy without `__pycache__`; `python -B`, `PYTHONDONTWRITEBYTECODE=1`, `-p no:cacheprovider`)

- **exp-0 (pytest).** Under `xvfb-run`: `1 failed, 287 passed in 3.57s` (failure `tests/test_setup_validation.py:66`, Windows path `C:\gst\itc.xltx` → name not split on Linux). Without display: `14 failed, 241 passed, 33 skipped in 1.40s` (13 GUI tests error "no display name and no $DISPLAY", plus the same path test). The four files in tests/README: `95 passed in 0.30s`.
- **exp-1 (`exp1_parse.py`).** All six README examples parse to the right type/client/state; `(1) GSTR3B-…` and `… (1).xlsx` parse; lower-case + `.XLSX` parses; `GSTR3B-ABC-XYZ Ltd-…` not parsed; `GSTR3B- -Maharashtra-Apr24.xlsx` → parsed=True, client='' (goes to variations); state 'M' → warning "State name issue: Client name too short". Sanitise: `Sales-ABC Private Limited-…` → `Sales-ABC_Pvt_Ltd-…`; `ABC_Private_Limited` → `ABC_Pvt_Ltd`; `ABC PRIVATE LIMITED` → `ABC_Pvt_Ltd`; `Privateer Unlimited` unchanged. Key: `ABC_Private_Limited` → `ABC_Private_Limited-MH` (not abbreviated); `max_length=7` → `AAAAAAA` (state code lost). Fallback: 'Unknown Place'→'UP', 'Atlantis'→'ATL'. `hasattr(os,'startfile')` False on Linux.
- **exp-2 (`exp2_organize.py`).** One generic client, 7 real .xlsx + 1 bad-signature file: bad file absent from scanned files and variations; status Complete; folders `…/Version-211026 0413/GSTR3B`, `ITC`, `Sales`; all copies Success; `USE_WIN32COM=False`; ITC data `annual_filename='AnnualReport-ABC_Ltd-Maharashtra-2024'`, B2 ends `\GSTR3B`; template without Links → returned True, sheets ['Data'], nothing written; resume with a different minute → new `Version-010199 0000`, status Success (re-copied); resume same minute → Skipped.
- **exp-3 (`exp3_consolidate.py`).** Real EXTRACTION_CONFIG, synthetic numbers per configured cell in `_extract_data` order: 13 columns OK, 9 MISMATCH (IMS Reco, Extra In Purchase, Extra In GSTR2B × IGST/CGST/SGST all got the Diff GSTR-3B J15/K15/L15 numbers).
- **exp-4 (`exp4_extractor_scan.py`, xvfb).** Parent folder with two runs → only the newer run's client listed; pointing at the older `Annual Statement-…` folder → "No Annual Statement folders found", 0 clients. Dark mode: label created before toggle → `#2b2b2b`; created after toggle → `white`.
- **exp-5.** `ast.parse(..., feature_version=(3,7))` on every non-test `.py` → no syntax rejected (best-effort check, not a 3.7 run).
- **exp-6 (greps).** `calculate_file_hash`, `_write_section`, `_create_details_sheet`, `get/set_cached_value`, `TEMPLATE_EXTENSIONS`, `EXCEL_SIGNATURES`, `MAX_FILE_SIZE_MB`, `validate_results`, `get_relative_excel_path`, `get_unique_path`, `EXCEL_EXTENSIONS`, `recent_folders` (read side), `metadata` (read side), parsed `warnings` (read side), `auto_refresh_if_missing`, `show_timestamp`, `excel_app` assignments, `gstin`/`pan`, `/mnt/user-data`, network modules, users of `logs/`/`temp/` — results as cited above.

---

## 10. Possible client names (locations only — nothing copied)

- `PROJECT_AUDIT_ANNUALSTATEMENT.md` lines 130, 133, 134, 158, 160, 161, 162, 177, 179, 180, 182, 227, 228, 231, 232 (A HEAD; introduced by commit f7eb8f1, absent from fc279af): one business name that looks like a real, well-known company, used in the worked examples (client part of example file names, folder paths and keys). Written as `<Client A>` in these notes.
- `gst_organizer.log` in commit edac6a3 (deleted in cbfa4a5): the audit (line 974) says it shows "actual client names" — not opened, per instructions.
- Commit 7347383 message: contains the author's network share path (server and folder names, including a financial-year folder) — infrastructure, not a client name; noted only because it is environment detail.
- Not client data: the README's "ABC Ltd"; the audit's "ABC-XYZ Ltd"; v1's "ABCLtd", "ABC Private Limited Industries", "ABC-DEF Corp"; test fixtures "ABC-DL" (`tests/test_extractor_panel.py:28`) — generic placeholders.

---

## 11. Questions for the author

1. X1: In a past `PQ_Extraction_Report_*.xlsx`, do the "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" sheets show the same three numbers as "GSTR-3B Difference" for each client? (Open one and compare row by row.) Was this ever noticed?
2. Also X1: when the BB2 check fails, should the extracted figures still be shown instead of "Missing"?
3. X2: When you create reports while Excel is open with your own work, does Excel close at the end?
4. X4: What should "Resume" do after an interrupted run — reuse the last Version folder?
5. Design note P15/P17: should 'full' start the refresh by itself (as the Setup card says), and do you still want a "Continue to extraction" button in organise-only mode?
6. P19: How do you currently refresh an older batch in Step 4 or the standalone extractor?
7. X3: Should processing stop when a template has no "Links" sheet?
8. X5: Should unreadable/too-small Excel files be listed (e.g. under variations) instead of disappearing?
9. X12: Is pywin32 meant to be required on Windows (be6418c), and should the Step 4 notice say so?
10. H5/X15: Is the March "copilot-worktree" branch the origin of project B? Were B's commits of 08-02 and 09-02-2026 made on that branch before the March audit?
11. The business name used in the audit's examples (Section 10): is it a real client? If yes, it is in A's git history (f7eb8f1).
12. `.gitignore` line 62 `/mnt/user-data/`: where does this come from, and is anything ever written there?

---

## INCOMPLETE

- Windows-only behaviour (Excel COM, SendKeys, `Quit()` on a reused Excel, 8.3 short paths, `os.startfile`, pywin32 post-install) was traced by reading only; X2 and G12 consequences on a real PC are unverified.
- Python 3.7 compatibility claimed by README/requirements was checked only with a grammar parse (exp-5), not by running 3.7.
- Line numbers quoted by the March audit were spot-checked against 47a1492 for about 30 references (mostly exact, a few one line off); not every one of its (many) line references was individually re-checked at 47a1492 — each claim was instead checked at HEAD as recorded above.
- The README was checked against A only (task scope); it is identical in B but B's code was not examined here.
- `gst_organizer.log` (edac6a3) deliberately not opened.
