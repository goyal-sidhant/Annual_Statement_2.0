# B-TRACE - Project B (AnnualStatement_oldv2, snapshot b0e7d0e, 09-02-2026): running the pure-Python parts on made-up inputs

Working notes for the later audit of project B. Every claim is marked **RAN** (I executed the code and saw the output) or **READ** (from reading the code), and every inference carries a confidence of **HIGH / MEDIUM / LOW**. File paths are repository-relative (for example `core/file_parser.py:123`); line numbers are those of snapshot `b0e7d0e`. GST rules are tagged **[GST RULE]**: these notes describe what the code does and never judge it. Values that may change over time are marked **TIME-VARYING**.

All client names below are invented (Alpha Test Traders, Beta Demo Industries, Gamma Sample LLP, Delta Example & Co, Epsilon Trial Enterprises, Zeta Mock Pvt. Ltd.). All figures are invented.

---

## 0. How this trace was done

### 0.1 Words used in these notes (plain English)

- **function / method**: a named block of code that does one job; a *method* belongs to a *class* (a bundle of code and data, for example `FileParser`).
- **regex (regular expression)**: a text pattern the code uses to recognise file names, for example "starts with `GSTR3B-`, then some text without hyphens, ...".
- **glob**: asking the operating system for "all files whose names fit `*.xlsx`" in one folder.
- **dict / list**: Python's built-in containers. A *dict* holds named entries ("key -> value"); a *list* holds entries in order.
- **JSON**: a plain-text file format for settings. The program stores its memory between sessions in JSON files.
- **checksum (SHA-256)**: a fingerprint of a file's bytes. If the fingerprint is the same before and after, not a single byte changed.
- **COM**: Microsoft's way of letting a program drive Excel remotely on Windows (through the `pywin32` package: modules `win32com`, `pythoncom`, `win32gui` ...). It cannot exist on this Linux machine.
- **stand-in (stub)**: a small fake module I wrote so that the real code can be imported and run here. My stand-ins pretend to be Excel but read cells with openpyxl (a Python library that reads and writes .xlsx files without Excel). Wherever a result depends on a stand-in, I say so.
- **cached value**: an .xlsx file stores each formula *and* the last result Excel calculated. openpyxl cannot calculate. With `data_only=True` it returns the stored result (or nothing, if Excel never calculated the file); with `data_only=False` it returns the formula text.
- **signal / thread**: PyQt5 (the window toolkit B uses) runs slow work in a background *thread* (a second line of work running at the same time) and reports back through *signals* (messages between objects).
- **log**: the running diary the program writes (warnings, errors) for the developer, not shown to the user as a message box.
- **sanitise**: clean a name so it is safe to use as a file or folder name.

### 0.2 Code read in full, line by line (READ)

| File (B @ b0e7d0e) | Lines | Last commits that touched it on B/main |
|---|---|---|
| `core/file_parser.py` | 1-399 | 7e09e5c (08-02-2026: single `stat`, <1 KB check and signature check moved into the parser), c258047 (07-02-2026: debug logging removed), 6237597, edac6a3 |
| `core/file_organizer.py` | 1-409 | c258047 (comment only), edac6a3 |
| `utils/helpers.py` | 1-537 | 7e09e5c (`find_excel_files` no longer opens each file; skips `~$` files), c258047, 6237597, edac6a3 |
| `utils/constants.py` | 1-275 | edac6a3 only (never changed since the first commit) |
| `power_query_extractor/config/cell_mappings.py` | 1-278 | 47a1492 (30-10-2025, last shared commit: six ITC mappings added), edac6a3 |
| `power_query_extractor/core/data_consolidator.py` | 1-417 | c258047 (duplicate header block and bare `except:` removed), edac6a3 |
| `models/cache_manager.py` | 1-152 | c258047 (created) |
| `models/run_history.py` | 1-148 | c258047 (created) |
| `gui/controllers/scan_controller.py` | 1-94 | b0e7d0e (statistics added to the result), c258047 (created) |

Also read because the assigned code depends on it: `models/app_state.py` (1-88), `power_query_extractor/core/report_processor.py` (1-751, needed to feed the consolidator), `gui/controllers/extraction_controller.py` (1-184), `gui/controllers/processing_controller.py` (1-298), `main.py` (1-63), `launch_extractor.py` (1-28), `power_query_extractor/extractor_main.py` (1-78), all `__init__.py` files of `core`, `models`, `utils`, `gui`, `gui/controllers`, `power_query_extractor`. I also read these targeted line ranges: `gui/pyqt_main_window.py:30-80, 200-230, 360-608`, `gui/views/extraction_view.py:360-400`, `gui/views/validation_view.py:370-395`, `gui/views/setup_view.py:374-379`, `core/excel_handler.py:14-80` plus its list of method names. I did not read the rest of `core/excel_handler.py` or the GUI views; other tasks cover them.

### 0.3 Where and how it ran

- Run folder `/root/merge-scratch/run/B-TRACE/`: `proj/` is a full copy of B-copy; `scripts/` holds my scripts; `stubs/` holds the stand-ins; `work/` holds the fake trees; `out/` holds every raw result as JSON.
- Python `/root/merge-scratch/venv-B/bin/python -B` (Python 3.14.6, openpyxl 3.1.5, PyQt5 5.15.11 / Qt 5.15.14), with `PYTHONDONTWRITEBYTECODE=1` and `QT_QPA_PLATFORM=offscreen`. `HOME` was pointed at `/root/merge-scratch/run/B-TRACE/fakehome` so that the cache manager (which writes into the user's home folder) never touched a real home folder. `TMPDIR` was pointed at `run/B-TRACE/tmp`, because `ExcelHandler()` creates `<temp>/gst_excel_temp` (`core/excel_handler.py:52-53`).
- **RAN:** B-copy is byte-for-byte unchanged. I compared SHA-256 of all 41 files before and after (`out/Bcopy_sha256_before.txt` = `out/Bcopy_sha256_after.txt`), and no `__pycache__` folder or `.pyc` file exists in B-copy or in the run folder. Nothing was written under `/home/user`. No git command was run against the original checkouts. I used only read-only `git log`/`git show`/`git diff` on `/root/merge-scratch/combined`.
- **Stand-ins** (`stubs/pythoncom.py`, `win32api.py`, `win32con.py`, `win32gui.py`, `win32com/client.py`). `win32com.client.Dispatch("Excel.Application")` returns a fake Excel that opens workbooks with openpyxl and answers `ws.Range("J15").Value`. It records every open, save, close, quit and "SendKeys" so the call order can be checked. No Power Query refresh happens and the fake Save writes nothing. Known differences from real Excel COM (each is reasoned, not run): real COM returns every number as a decimal (`100` -> `100.0`); it returns error cells such as `#N/A` as large negative whole numbers (for example `-2146826246`) where my stand-in gives the text `'#N/A'`; and real Excel recalculates formulas when it opens a file, while my stand-in cannot. That last point: **MEDIUM** confidence that a "never-calculated" formula reads as its calculated value under real Excel.

### 0.4 Scripts (all in `/root/merge-scratch/run/B-TRACE/scripts/`)

| Script | Covers | Raw output |
|---|---|---|
| `common.py` | puts `proj/` on the import path, captures log lines | - |
| `t1_filenames.py` | Part 1: 12 fixed names + 58 own names through `FileParser.parse_filename`, every regex separately, `_suggest_pattern` | `out/t1_filenames.json`, `out/t1_filenames.md` |
| `t2_names.py` | Part 2: `sanitize_filename` (36 inputs + 5 extra), `create_client_state_key` (17), `get_state_code` (30 variants + 18 fallbacks), `validate_client_name` (20) | `out/t2_names.json` |
| `t3_periods.py` | Part 3: timestamps, period text kept by the parser, two-year client, folder-name ordering | `out/t3_periods.json` |
| `t4_extract.py` | Part 4: mapping tables, made-up ITC/Sales reports, real `ReportProcessor` + `DataConsolidator` with stand-ins, refreshed-copy naming, refresh validation | `out/t4_extract.json`, `out/t4_stdout.txt` |
| `t4b_consolidator.py` | Part 4: consolidator alone on synthetic inputs, calculation and description branches, mapping build loop | `out/t4b_consolidator.json` |
| `t5_cache.py` | Part 5: cache manager, run history, migration (42 steps) | `out/t5_cache.json` |
| `t6_scan_organise.py` | Part 6: fake source tree, threaded `ScanController`, real `ProcessingWorker` in six runs (fresh/rerun/resume), checksums | `out/t6_scan_organise.json` |
| `t6b_organiser_edges.py` | Part 6: remaining organiser branches | `out/t6b_organiser_edges.json` |
| `t7_helpers.py` | the remaining small functions of `utils/helpers.py` | `out/t7_helpers.json` |
| `make_fragments.py` | turns the JSON results into the tables pasted below (runs no project code) | `out/fragments/*.md` |

Extra one-line runs (RAN, shown in sections 4 and 9): `sanitize_filename` on five dotted names; `python -B proj/main.py` without stand-ins; `import power_query_extractor.extractor_main`.

---

## 1. File-name recognition (Part 1)

### 1.1 What the code does (READ)

`FileParser.parse_filename` (`core/file_parser.py:56-124`) tries the six regexes of `FILE_PATTERNS` (`utils/constants.py:13-62`) in this order, using `re.match` (the pattern must match from the first character). It stops at the first match (`:118`). All six are case-insensitive (`re.IGNORECASE`).

| order | key | regex (exact) | type | category | folder key | captured groups |
|---|---|---|---|---|---|---|
| 1 | `GSTR-2B-Reco` | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | GSTR-2B Reco | ITC | itc | client, state, period |
| 2 | `ImsReco` | `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$` | IMS Reco | ITC | itc | client, state, date |
| 3 | `GSTR3B` | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | GSTR-3B Export | GSTR3B | gstr3b | client, state, month |
| 4 | `Sales` | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | Sales | Sales | sales | client, state, start_month, end_month |
| 5 | `SalesReco` | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | Sales Reco | Sales | sales | client, state, period |
| 6 | `AnnualReport` | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$` | Annual Report | Annual | root | client, state, year |

`[^-]+` means "one or more characters that are not a hyphen", so a client or state name can never contain a hyphen. `\.xlsx?$` means the name must end in `.xls` or `.xlsx` (any case), so `.xlsm` never matches even though the scanner collects `.xlsm` files (`utils/helpers.py:255`). Captured groups are trimmed of spaces (`core/file_parser.py:98,100,102`). A group is only stored if it is non-empty *before* trimming (`:96`), so a group of only spaces becomes `''`. Client and state are each checked with `validate_client_name` (`:105-115`). This happens only when *both* are non-empty after trimming, and the state is checked with the same function, so its message reads "State name issue: Client name ..." (RAN, section 2.6). The parser keeps period, date, month, start_month, end_month and year as text in `metadata` (`:102`); nothing else in B ever reads `metadata` (READ: `grep metadata` finds only `core/file_parser.py:74,102`; HIGH).

`_suggest_pattern` (`core/file_parser.py:316-334`) produces the "expected pattern" hint for an unrecognised file. It checks the lower-cased name for these substrings, in this order (first true one wins): `'gstr' and '2b'` -> `GSTR-2B-Reco-ClientName-State-Period.xlsx`; `'ims'` -> `ImsReco-ClientName-State-DDMMYYYY.xlsx`; `'gstr' and '3b'` -> `GSTR3B-ClientName-State-Month.xlsx`; `'sales' and 'reco'` -> `SalesReco-...`; `'sales'` -> `Sales-ClientName-State-StartMonth-EndMonth.xlsx`; `'annual'` -> `AnnualReport-ClientName-State-Year.xlsx`; otherwise `Check file naming convention`. Because the `if ... elif` chain always produces at most one suggestion, the `" or ".join(...)` at `:334` never actually joins two (READ, HIGH).

### 1.2 Results (RAN, `scripts/t1_filenames.py`, function `core/file_parser.py:56-124` and `:316-334`)

F01-F12 are the 12 fixed comparison names, exactly as given. O01-O58 are my own. "Recognised" is what `parse_filename` returned. The column "`_suggest_pattern` text" is computed for every row: the scanner only shows it for unrecognised files (`core/file_parser.py:251`).

| # | Input file name | What it probes | Recognised? (pattern) | type / category / folder | client | state | metadata kept | warnings / error | `_suggest_pattern` text |
|---|---|---|---|---|---|---|---|---|---|
| F01 | 'Alpha Test Traders GSTR-2B Apr-24.xlsx' | fixed comparison name | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| F02 | 'Alpha Test Traders GSTR2B APR 2024.xlsx' | fixed comparison name | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| F03 | 'ALPHA TEST TRADERS gstr-1 (1).xlsx' | fixed comparison name | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | Check file naming convention |
| F04 | 'Beta Demo Industries Pvt Ltd_GSTR3B_2024-25.xlsx' | fixed comparison name | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |
| F05 | 'Beta Demo Industries Private Limited Purchase Register.xlsx' | fixed comparison name | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | Check file naming convention |
| F06 | 'Gamma Sample LLP Sales Register FY 2024-25.xls' | fixed comparison name | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | Sales-ClientName-State-StartMonth-EndMonth.xlsx |
| F07 | 'Gamma Sample LLP IMS Reco (2).xlsx' | fixed comparison name | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | ImsReco-ClientName-State-DDMMYYYY.xlsx |
| F08 | 'Delta Example & Co Reco Diff.xlsx' | fixed comparison name | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | Check file naming convention |
| F09 | 'Delta Example & Co ITC Report_refreshed.xlsx' | fixed comparison name | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | Check file naming convention |
| F10 | 'Epsilon Trial Enterprises gstr9.json' | fixed comparison name | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | Check file naming convention |
| F11 | 'Epsilon Trial Enterprises Table 8D Limits.xlsx' | fixed comparison name | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | Check file naming convention |
| F12 | 'Zeta Mock Pvt. Ltd. Extra In GSTR2B.xlsx' | fixed comparison name | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| O01 | 'GSTR-2B-Reco-Alpha Test Traders-Delhi-Apr24.xlsx' | canonical GSTR-2B-Reco | YES (GSTR-2B-Reco) | GSTR-2B Reco / ITC / itc | 'Alpha Test Traders' | 'Delhi' | {"period": "Apr24"} | - | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| O02 | 'GSTR-2B-Reco-Alpha Test Traders-Delhi-Apr24-Mar25.xlsx' | period containing a hyphen | YES (GSTR-2B-Reco) | GSTR-2B Reco / ITC / itc | 'Alpha Test Traders' | 'Delhi' | {"period": "Apr24-Mar25"} | - | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| O03 | 'GSTR-2B-Reco-Alpha Test Traders-Delhi-FY.xls' | .xls extension | YES (GSTR-2B-Reco) | GSTR-2B Reco / ITC / itc | 'Alpha Test Traders' | 'Delhi' | {"period": "FY"} | - | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| O04 | 'gstr-2b-reco-alpha test traders-delhi-apr24.XLSX' | all lower case + upper-case extension | YES (GSTR-2B-Reco) | GSTR-2B Reco / ITC / itc | 'alpha test traders' | 'delhi' | {"period": "apr24"} | - | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| O05 | 'GSTR-2B-Reco-Alpha-Test-Traders-Delhi-Apr24.xlsx' | client name containing hyphens | YES (GSTR-2B-Reco) | GSTR-2B Reco / ITC / itc | 'Alpha' | 'Test' | {"period": "Traders-Delhi-Apr24"} | - | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| O06 | 'GSTR-2B-Reco-Alpha:Test-Delhi-Apr24.xlsx' | client with ':' (invalid on Windows) | YES (GSTR-2B-Reco) | GSTR-2B Reco / ITC / itc | 'Alpha:Test' | 'Delhi' | {"period": "Apr24"} | Client name issue: Client name contains invalid characters | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| O07 | 'GSTR-2B-Reco-Alpha Test Traders-Delhi-Apr24_Refreshed_211026_1100.xlsx' | refreshed suffix on a source name | YES (GSTR-2B-Reco) | GSTR-2B Reco / ITC / itc | 'Alpha Test Traders' | 'Delhi' | {"period": "Apr24_Refreshed_211026_1100"} | - | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| O08 | 'GSTR-2B-Reco-Alpha Test Traders-Delhi-Apr24 (1).xlsx' | browser duplicate suffix ' (1)' | YES (GSTR-2B-Reco) | GSTR-2B Reco / ITC / itc | 'Alpha Test Traders' | 'Delhi' | {"period": "Apr24 (1)"} | - | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| O09 | 'ImsReco-Alpha Test Traders-Delhi-31032025.xlsx' | canonical ImsReco | YES (ImsReco) | IMS Reco / ITC / itc | 'Alpha Test Traders' | 'Delhi' | {"date": "31032025"} | - | ImsReco-ClientName-State-DDMMYYYY.xlsx |
| O10 | 'ImsReco-Alpha Test Traders-Delhi-99999999.xlsx' | 8 digits that are not a real date | YES (ImsReco) | IMS Reco / ITC / itc | 'Alpha Test Traders' | 'Delhi' | {"date": "99999999"} | - | ImsReco-ClientName-State-DDMMYYYY.xlsx |
| O11 | 'ImsReco-Alpha Test Traders-Delhi-3103202.xlsx' | 7 digits | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | ImsReco-ClientName-State-DDMMYYYY.xlsx |
| O12 | 'ImsReco-Alpha Test Traders-Delhi-31-03-2025.xlsx' | date with hyphens | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | ImsReco-ClientName-State-DDMMYYYY.xlsx |
| O13 | 'ImsReco-Alpha Test Traders-Delhi-31032025 (1).xlsx' | duplicate suffix after date | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | ImsReco-ClientName-State-DDMMYYYY.xlsx |
| O14 | 'IMS Reco-Alpha Test Traders-Delhi-31032025.xlsx' | 'IMS Reco' with space | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | ImsReco-ClientName-State-DDMMYYYY.xlsx |
| O15 | 'GSTR3B-Alpha Test Traders-Delhi-Apr.xlsx' | canonical GSTR3B | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Alpha Test Traders' | 'Delhi' | {"month": "Apr"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O16 | '(1) GSTR3B-Alpha Test Traders-Delhi-May.xlsx' | '(1) ' prefix (allowed for GSTR3B only) | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Alpha Test Traders' | 'Delhi' | {"month": "May"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O17 | '(12)GSTR3B-Alpha Test Traders-Delhi-Jun.xlsx' | '(12)' prefix, no space | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Alpha Test Traders' | 'Delhi' | {"month": "Jun"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O18 | 'GSTR3B-Alpha Test Traders-Delhi-Jul (1).xlsx' | ' (1)' duplicate suffix | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Alpha Test Traders' | 'Delhi' | {"month": "Jul (1)"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O19 | 'GSTR3B-Alpha Test Traders-Delhi-Aug - Copy.xlsx' | Windows ' - Copy' suffix | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |
| O20 | 'GSTR3B-Alpha Test Traders-Delhi-Apr-2024.xlsx' | extra hyphen in month | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |
| O21 | 'GSTR-3B-Alpha Test Traders-Delhi-Apr.xlsx' | 'GSTR-3B' with hyphen | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |
| O22 | 'GSTR3B_Alpha Test Traders_Delhi_Apr.xlsx' | underscores instead of hyphens | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |
| O23 | 'GSTR3B-Alpha Test Traders-Delhi-Apr_Refreshed_211026_1100.xlsx' | refreshed suffix on GSTR3B | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Alpha Test Traders' | 'Delhi' | {"month": "Apr_Refreshed_211026_1100"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O24 | 'GSTR3B-Alpha Test Traders-Delhi-Apr.xlsx.xlsx' | double extension | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Alpha Test Traders' | 'Delhi' | {"month": "Apr.xlsx"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O25 | 'GSTR3B-Alpha Test Traders-Delhi-Apr.xlsm' | .xlsm (globbed by scanner) | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |
| O26 | 'GSTR3B-Alpha Test Traders-Delhi-Apr.csv' | .csv | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |
| O27 | 'GSTR3B-Alpha Test Traders-Delhi-Apr.xlsx ' | trailing space after extension | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |
| O28 | ' GSTR3B-Alpha Test Traders-Delhi-Apr.xlsx' | leading space | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |
| O29 | '~$GSTR3B-Alpha Test Traders-Delhi-Apr.xlsx' | Excel lock/temp file | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |
| O30 | 'GSTR3B-Alpha Test Traders-Delhi-.xlsx' | empty month | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |
| O31 | 'GSTR3B- - -Apr.xlsx' | client and state are only spaces | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | '' | '' | {"month": "Apr"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O32 | 'GSTR3B-A-Delhi-Apr.xlsx' | one-letter client | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'A' | 'Delhi' | {"month": "Apr"} | Client name issue: Client name too short | GSTR3B-ClientName-State-Month.xlsx |
| O33 | 'GSTR3B-Alpha Test Traders-Tamil Nadu-Apr.xlsx' | two-word state | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Alpha Test Traders' | 'Tamil Nadu' | {"month": "Apr"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O34 | 'GSTR3B-Alpha Test Traders-DL-Apr.xlsx' | state written as code | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Alpha Test Traders' | 'DL' | {"month": "Apr"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O35 | 'GSTR3B-Alpha Test Traders-New Delhi-Apr.xlsx' | unknown two-word state | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Alpha Test Traders' | 'New Delhi' | {"month": "Apr"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O36 | 'GSTR3B-Alpha Test Traders-Tamilnadu-Apr.xlsx' | unknown one-word state | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Alpha Test Traders' | 'Tamilnadu' | {"month": "Apr"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O37 | 'GSTR3B–Alpha Test Traders–Delhi–Apr.xlsx' | en-dash instead of hyphen | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |
| O38 | 'GSTR3B-Ālpha Tēst Traders-Delhi-Apr.xlsx' | accented Latin letters | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Ālpha Tēst Traders' | 'Delhi' | {"month": "Apr"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O39 | 'GSTR3B-अल्फा ट्रेडर्स-Delhi-Apr.xlsx' | Devanagari client name | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'अल्फा ट्रेडर्स' | 'Delhi' | {"month": "Apr"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O40 | 'GSTR3B-Alpha Test Traders-Delhi-Apr.XLS' | upper-case .XLS | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Alpha Test Traders' | 'Delhi' | {"month": "Apr"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O41 | GSTR3B-('Alpha Test Traders ' x8, stripped)-Delhi-Apr.xlsx  [len 173] | client of 151 chars | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | '<151 chars>' | 'Delhi' | {"month": "Apr"} | Client name issue: Client name too long | GSTR3B-ClientName-State-Month.xlsx |
| O42 | GSTR-2B-Reco-Alpha Test Traders-Delhi-('Apr24 ' x45, stripped).xlsx  [len 312] | file name of ~300 chars | YES (GSTR-2B-Reco) | GSTR-2B Reco / ITC / itc | 'Alpha Test Traders' | 'Delhi' | {"period": "<269 chars>"} | - | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| O43 | 'Sales-Alpha Test Traders-Delhi-Apr24-Mar25.xlsx' | canonical Sales | YES (Sales) | Sales / Sales / sales | 'Alpha Test Traders' | 'Delhi' | {"start_month": "Apr24", "end_month": "Mar25"} | - | Sales-ClientName-State-StartMonth-EndMonth.xlsx |
| O44 | 'Sales-Alpha Test Traders-Delhi-FY2024-25.xlsx' | FY with hyphen read as start/end | YES (Sales) | Sales / Sales / sales | 'Alpha Test Traders' | 'Delhi' | {"start_month": "FY2024", "end_month": "25"} | - | Sales-ClientName-State-StartMonth-EndMonth.xlsx |
| O45 | 'Sales-Alpha Test Traders-Delhi-Apr24.xlsx' | Sales with only one period part | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | Sales-ClientName-State-StartMonth-EndMonth.xlsx |
| O46 | 'Sales-Alpha Test Traders-Delhi-Apr-Jun (1).xlsx' | Sales with duplicate suffix | YES (Sales) | Sales / Sales / sales | 'Alpha Test Traders' | 'Delhi' | {"start_month": "Apr", "end_month": "Jun (1)"} | - | Sales-ClientName-State-StartMonth-EndMonth.xlsx |
| O47 | 'SalesReco-Alpha Test Traders-Delhi-Q1.xlsx' | canonical SalesReco | YES (SalesReco) | Sales Reco / Sales / sales | 'Alpha Test Traders' | 'Delhi' | {"period": "Q1"} | - | SalesReco-ClientName-State-Period.xlsx |
| O48 | 'Sales Reco-Alpha Test Traders-Delhi-Q1.xlsx' | 'Sales Reco' with space | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | SalesReco-ClientName-State-Period.xlsx |
| O49 | 'AnnualReport-Alpha Test Traders-Delhi-2024-25.xlsx' | canonical AnnualReport, year with hyphen | YES (AnnualReport) | Annual Report / Annual / root | 'Alpha Test Traders' | 'Delhi' | {"year": "2024-25"} | - | AnnualReport-ClientName-State-Year.xlsx |
| O50 | 'Annual Report-Alpha Test Traders-Delhi-2024.xlsx' | 'Annual Report' with space | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | AnnualReport-ClientName-State-Year.xlsx |
| O51 | 'Sales-Zeta Mock Pvt. Ltd.-Maharashtra-Apr-Mar.xlsx' | client with dots | YES (Sales) | Sales / Sales / sales | 'Zeta Mock Pvt. Ltd.' | 'Maharashtra' | {"start_month": "Apr", "end_month": "Mar"} | - | Sales-ClientName-State-StartMonth-EndMonth.xlsx |
| O52 | 'GSTR3B-Beta Demo Industries Private Limited-Maharashtra-Apr.xlsx' | 'Private Limited' client | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Beta Demo Industries Private Limited' | 'Maharashtra' | {"month": "Apr"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O53 | 'GSTR3B-Delta Example & Co-Gujarat-Apr.xlsx' | client with '&' | YES (GSTR3B) | GSTR-3B Export / GSTR3B / gstr3b | 'Delta Example & Co' | 'Gujarat' | {"month": "Apr"} | - | GSTR3B-ClientName-State-Month.xlsx |
| O54 | 'Alpha Claims Register.xlsx' | 'ims' hidden inside 'Claims' | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | ImsReco-ClientName-State-DDMMYYYY.xlsx |
| O55 | 'Alpha Annual Return Data.xlsx' | 'annual' keyword | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | AnnualReport-ClientName-State-Year.xlsx |
| O56 | 'Alpha Sales Reco.xlsx' | 'sales' + 'reco' keywords | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | SalesReco-ClientName-State-Period.xlsx |
| O57 | 'ITC_Report_Alpha_Test_Traders_DL_211026_1000_Refreshed_211026_1100.xlsx' | app's own refreshed report | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | Check file naming convention |
| O58 | 'GSTR3B-Alpha Test Traders-Delhi-Apr' | no extension at all | no | Unknown / Unknown / Unknown | '' | '' | {} | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx |

**Two patterns matching one name: impossible with these six regexes (RAN).** For all 70 names I tested every regex separately (column "Recognised" uses the first match; the JSON field `all_patterns_matching` lists every match). No name matched more than one pattern, because each pattern begins with a different fixed word (`GSTR-2B-Reco-`, `ImsReco-`, `GSTR3B-` with an optional `(n)` in front, `Sales-`, `SalesReco-`, `AnnualReport-`). `Sales-` cannot match `SalesReco-...`, because `Sales-` needs a hyphen right after "Sales". (HIGH)

### 1.3 Surprises in Part 1, and why

1. **All 12 fixed comparison names are unrecognised (RAN, F01-F12).** B only knows names that *begin* with its fixed prefix (`re.match`, `core/file_parser.py:79`), so a name that begins with the client's name ("Alpha Test Traders GSTR-2B ...") never matches. F10 (`.json`) would not even be looked at by a folder scan (section 6). (HIGH)
2. **Hyphens split names in the wrong place (RAN, O05).** `GSTR-2B-Reco-Alpha-Test-Traders-Delhi-Apr24.xlsx` gives client `'Alpha'`, state `'Test'`, period `'Traders-Delhi-Apr24'`. It is recognised without any warning and becomes a client "Alpha" in a state "Test" (state code fallback `TES`, section 2.5). Cause: `[^-]+` for client and state, `(.+)` for the period. (HIGH)
3. **Whitespace-only client/state is "recognised" but then dropped (RAN, O31).** `GSTR3B- - -Apr.xlsx` -> `parsed=True`, client `''`, state `''`, and no warning, because the validation at `:105` needs both to be non-empty. In a folder scan, `_process_file` (`:243-252`) then puts it in the variations list with issue `Could not parse filename`. That text is the fallback at `:249`, because `parse_filename` sets no `error` key when a pattern matched. (HIGH)
4. **Windows " - Copy" duplicates are rejected; " (1)" duplicates are accepted (RAN, O18, O19, O08, O13, O46).** " (1)" contains no hyphen, so it becomes part of the month/period text (`'Jul (1)'`, `'Apr24 (1)'`, end_month `'Jun (1)'`). " - Copy" adds a hyphen, so the GSTR3B regex fails. ImsReco needs exactly 8 digits at the end, so `31032025 (1)` fails. A `(1)` *prefix* is allowed only for GSTR3B (`^(?:\(\d+\)\s*)?`). (HIGH)
5. **Refreshed suffixes ride along in the period/month (RAN, O07, O23).** `..._Refreshed_211026_1100` is kept inside the period/month text, so a refreshed copy dropped into the source folder is treated as just another GSTR-2B-Reco/GSTR3B file of that client. The app's own refreshed report names (`ITC_Report_..._Refreshed_...`, O57) are not recognised. (HIGH)
6. **`.xlsm` and `.csv` are not recognised (RAN, O25, O26)**, and neither are a trailing space, a leading space, an en-dash instead of a hyphen, `GSTR-3B-` with a hyphen, underscores instead of hyphens, `Sales Reco`/`IMS Reco`/`Annual Report` with a space, or a missing extension (O27, O28, O37, O21, O22, O48, O14, O50, O58). Upper-case extensions are accepted by the regex (O04, O40). Section 6 shows the scanner's own case rule.
7. **Double extension accepted (RAN, O24):** `...-Apr.xlsx.xlsx` -> month `'Apr.xlsx'`. (HIGH)
8. **[GST RULE] Period text is not checked at all (RAN, O10, O44 and section 3).** ImsReco accepts `99999999` and `00000000` as the 8-digit date. Sales `...-FY2024-25.xlsx` splits into start_month `'FY2024'`, end_month `'25'`. Nothing reads these values afterwards. **TIME-VARYING** (periods, years).
9. **Length warnings only (RAN, O41, O42):** a 151-character client is recognised with the warning "Client name too long" and still used. A 312-character file name is recognised normally (the parser does not check length; sections 2 and 6 cover length in folder and file names). (HIGH)
10. **Unicode accepted (RAN, O38, O39):** accented and Devanagari client names are recognised and kept as-is.
11. **Suggestion hint can mislead (RAN, O54):** `Alpha Claims Register.xlsx` is told to follow the `ImsReco-...` pattern, because "Cla**ims**" contains `ims` (`core/file_parser.py:323`). F03 (`gstr-1`) and F08/F09/F11 get only "Check file naming convention". (HIGH)

---

## 2. Name cleaning (Part 2)

### 2.1 `sanitize_filename` (`utils/helpers.py:66-112`) - what it does (READ)

1. Empty input -> `"unnamed"` (`:70-71`).
2. Splits the name with Python's `Path(...)`: `stem` (the part before the last dot) and `suffix` (from the last dot on) (`:74-75`). **The suffix is put back at the end completely unchanged** (`:109-110`).
3. On the stem only: text replacement `'Private'` -> `'Pvt'`, then `'Limited'` -> `'Ltd'` (`:78-79`). This is case-sensitive and does not respect word boundaries.
4. Control characters (codes 0-31 and 127-159) -> `_` (`:85-88`).
5. Each of the characters `` <>:"/\|?*[]{}+=!@#$%^,;'"`~ `` (the string at `:82`) -> `_` (`:91-92`). Hyphen, dot, `&` and round brackets are kept.
6. Any run of whitespace or underscores -> one `_` (`:95`). In Python 3 this also catches the non-breaking space.
7. Strips leading/trailing `_`, `.` and spaces (`:98`); if nothing is left -> `"unnamed"` (`:101-102`).
8. Cuts the stem to `max_length - len(extension) - 10` characters, which is 190 with no extension and 185 for `.xlsx` (`:105-106`). This cut happens *after* step 7, so the result can end in `_`.

Where it is used (READ): every copied data file name (`core/file_organizer.py:254`); the client part of the three category folder names (`:86, :115`); and the client part of the ITC/Sales report file names (`gui/controllers/processing_controller.py:112, 120, 128`).

### 2.2 `sanitize_filename` results (RAN, `scripts/t2_names.py`)

| # | Input | What it probes | `Path(...).stem` / `.suffix` as Python sees it | Output of `sanitize_filename` | sanitising twice gives the same? |
|---|---|---|---|---|---|
| S01 | 'Alpha Test Traders' (len 18) | plain client name | 'Alpha Test Traders' / '' | 'Alpha_Test_Traders' (len 18) | yes |
| S02 | 'Beta Demo Industries Private Limited' (len 36) | Private Limited | 'Beta Demo Industries Private Limited' / '' | 'Beta_Demo_Industries_Pvt_Ltd' (len 28) | yes |
| S03 | 'Beta Demo Industries PRIVATE LIMITED' (len 36) | upper case PRIVATE LIMITED | 'Beta Demo Industries PRIVATE LIMITED' / '' | 'Beta_Demo_Industries_PRIVATE_LIMITED' (len 36) | yes |
| S04 | 'Beta Demo Industries private limited' (len 36) | lower case private limited | 'Beta Demo Industries private limited' / '' | 'Beta_Demo_Industries_private_limited' (len 36) | yes |
| S05 | 'Beta Demo Industries Pvt Ltd' (len 28) | already abbreviated | 'Beta Demo Industries Pvt Ltd' / '' | 'Beta_Demo_Industries_Pvt_Ltd' (len 28) | yes |
| S06 | 'Beta Demo Industries Pvt. Ltd.' (len 30) | Pvt. Ltd. with dots | 'Beta Demo Industries Pvt. Ltd' / '.' | 'Beta_Demo_Industries_Pvt._Ltd.' (len 30) | yes |
| S07 | 'Zeta Mock Pvt. Ltd' (len 18) | dot inside name, no trailing dot | 'Zeta Mock Pvt' / '. Ltd' | 'Zeta_Mock_Pvt. Ltd' (len 18) | yes |
| S08 | 'Gamma Sample LLP' (len 16) | LLP | 'Gamma Sample LLP' / '' | 'Gamma_Sample_LLP' (len 16) | yes |
| S09 | 'Delta Example & Co' (len 18) | ampersand | 'Delta Example & Co' / '' | 'Delta_Example_&_Co' (len 18) | yes |
| S10 | 'Delta Example, Sons & Co.' (len 25) | comma + ampersand + trailing dot | 'Delta Example, Sons & Co' / '.' | 'Delta_Example_Sons_&_Co.' (len 24) | yes |
| S11 | 'Alpha_Test__Traders' (len 19) | underscores | 'Alpha_Test__Traders' / '' | 'Alpha_Test_Traders' (len 18) | yes |
| S12 | 'Alpha   Test\tTraders' (len 20) | multiple spaces and a tab | 'Alpha   Test\tTraders' / '' | 'Alpha_Test_Traders' (len 18) | yes |
| S13 | 'Alpha\xa0Test Traders' (len 18) | non-breaking space | 'Alpha\xa0Test Traders' / '' | 'Alpha_Test_Traders' (len 18) | yes |
| S14 | 'Ālpha Tēst Traders' (len 18) | accented letters | 'Ālpha Tēst Traders' / '' | 'Ālpha_Tēst_Traders' (len 18) | yes |
| S15 | 'अल्फा ट्रेडर्स' (len 14) | Devanagari | 'अल्फा ट्रेडर्स' / '' | 'अल्फा_ट्रेडर्स' (len 14) | yes |
| S16 | 'Alpha (India) Traders [Unit-2] {North}' (len 38) | brackets of three kinds | 'Alpha (India) Traders [Unit-2] {North}' / '' | 'Alpha_(India)_Traders_Unit-2_North' (len 34) | yes |
| S17 | 'Alpha+Beta=Gamma! @Delhi #1 $ % ^ ~ `' (len 37) | many symbols | 'Alpha+Beta=Gamma! @Delhi #1 $ % ^ ~ `' / '' | 'Alpha_Beta_Gamma_Delhi_1' (len 24) | yes |
| S18 | 'Alpha\x85Traders\x07' (len 14) | control characters (U+0085, BEL) | 'Alpha\x85Traders\x07' / '' | 'Alpha_Traders' (len 13) | yes |
| S19 | '' (len 0) | empty string | '' / '' | 'unnamed' (len 7) | yes |
| S20 | '   ' (len 3) | only spaces | '   ' / '' | 'unnamed' (len 7) | yes |
| S21 | '___' (len 3) | only underscores | '___' / '' | 'unnamed' (len 7) | yes |
| S22 | 'Unlimited Privateer Traders' (len 27) | words containing Limited/Private | 'Unlimited Privateer Traders' / '' | 'Unlimited_Pvter_Traders' (len 23) | yes |
| S23 | <285 chars: 'Alpha Test Traders Alpha Test '...' Alpha Test Traders '> (len 285) | 285-char name | <285 chars: 'Alpha Test Traders Alpha Test '...' Alpha Test Traders '> / '' | <190 chars: 'Alpha_Test_Traders_Alpha_Test_'...'_Alpha_Test_Traders_'> (len 190) | NO |
| S24 | 'Alpha/Beta Traders' (len 18) | forward slash | 'Beta Traders' / '' | 'Beta_Traders' (len 12) | yes |
| S25 | 'Alpha\\Beta Traders' (len 18) | back slash | 'Alpha\\Beta Traders' / '' | 'Alpha_Beta_Traders' (len 18) | yes |
| S26 | 'Alpha: Beta? Traders*' (len 21) | Windows-invalid : ? * | 'Alpha: Beta? Traders*' / '' | 'Alpha_Beta_Traders' (len 18) | yes |
| S27 | 'GSTR3B-Alpha Test Traders-Delhi-Apr (1).xlsx' (len 44) | file name with (1) | 'GSTR3B-Alpha Test Traders-Delhi-Apr (1)' / '.xlsx' | 'GSTR3B-Alpha_Test_Traders-Delhi-Apr_(1).xlsx' (len 44) | yes |
| S28 | 'GSTR3B-Alpha Test Traders-Delhi-Apr 24.xlsx' (len 43) | file name with space | 'GSTR3B-Alpha Test Traders-Delhi-Apr 24' / '.xlsx' | 'GSTR3B-Alpha_Test_Traders-Delhi-Apr_24.xlsx' (len 43) | yes |
| S29 | 'GSTR3B-Alpha Test Traders-Delhi-Apr_24.xlsx' (len 43) | file name with underscore (collides with S28) | 'GSTR3B-Alpha Test Traders-Delhi-Apr_24' / '.xlsx' | 'GSTR3B-Alpha_Test_Traders-Delhi-Apr_24.xlsx' (len 43) | yes |
| S30 | 'GSTR3B-Beta Demo Industries Private Limited-Maharashtra-Apr.xlsx' (len 64) | file name with Private Limited | 'GSTR3B-Beta Demo Industries Private Limited-Maharashtra-Apr' / '.xlsx' | 'GSTR3B-Beta_Demo_Industries_Pvt_Ltd-Maharashtra-Apr.xlsx' (len 56) | yes |
| S31 | '.xlsx' (len 5) | only an extension | '.xlsx' / '' | 'xlsx' (len 4) | yes |
| S32 | 'Report.' (len 7) | trailing dot | 'Report' / '.' | 'Report.' (len 7) | yes |
| S33 | 'Alpha.Test.Traders' (len 18) | dots between words | 'Alpha.Test' / '.Traders' | 'Alpha.Test.Traders' (len 18) | yes |
| S34 | <313 chars: 'GSTR-2B-Reco-Alpha Test Trader'...'24 Apr24 Apr24 .xlsx'> (len 313) | ~300-char file name | <308 chars: 'GSTR-2B-Reco-Alpha Test Trader'...'4 Apr24 Apr24 Apr24 '> / '.xlsx' | <190 chars: 'GSTR-2B-Reco-Alpha_Test_Trader'...'Apr24_Apr24_Apr.xlsx'> (len 190) | yes |
| S35 | 'Alpha Test Traders.' (len 19) | client ending in a dot | 'Alpha Test Traders' / '.' | 'Alpha_Test_Traders.' (len 19) | yes |
| S36 | 'Ltd.' (len 4) | only 'Ltd.' | 'Ltd' / '.' | 'Ltd.' (len 4) | yes |

Extra run (RAN, one-liner): `'Alpha. Beta: Gamma?'` -> `'Alpha. Beta: Gamma?'` (unchanged); `'Alpha Traders. Unit*2'` -> `'Alpha_Traders. Unit*2'`; `'Alpha.Beta Traders Private Limited'` -> `'Alpha.Beta Traders Private Limited'` (unchanged); `'Beta Demo Private Ltd.'` -> `'Beta_Demo_Pvt_Ltd.'`; `'GSTR3B-Alpha. Traders-Delhi-Apr.xlsx'` -> `'GSTR3B-Alpha._Traders-Delhi-Apr.xlsx'`. A Windows-style path check (RAN, `PureWindowsPath`) gives stem `'Beta Traders'` for `'Alpha\Beta Traders'`: on Windows a back-slash is treated as a folder separator, like the forward slash in S24. (MEDIUM: shown with Python's Windows path rules on Linux, not on Windows.)

### 2.3 Surprises in sanitising, and why

1. **For a client name containing a dot, everything after the last dot escapes cleaning (RAN, S07, S10, S33, extra run).** Client names have no extension, but `Path()` still treats ". Ltd" (S07), "." (S06, S10, S35, S36), ".Traders" (S33) or ". Beta: Gamma?" as an "extension". That part is re-attached untouched, with spaces, `:` `?` `*` and "Private Limited" left unchanged. So `Zeta Mock Pvt. Ltd` -> `Zeta_Mock_Pvt. Ltd` (the space survives), and `Alpha.Beta Traders Private Limited` is returned completely unchanged. This affects category folder names and report file names for such clients. **Python-version note:** on this Python 3.14.6 a trailing dot counts as a suffix (`Path('Report.').suffix == '.'`, RAN). I am not sure older Python versions behave the same; the README says "Python 3.7 or higher". To check: run `python -c "from pathlib import Path; print(repr(Path('Report.').suffix))"` on the author's Python. (HIGH for the behaviour on 3.14; LOW for other versions.)
2. **"Private"/"Limited" replacement is case-sensitive and not word-based (RAN, S03, S04, S22).** `PRIVATE LIMITED` and `private limited` are not abbreviated, while `Privateer` becomes `Pvter`. `create_client_state_key` (section 2.4) does the same job differently: case-insensitive and whole words only (`utils/helpers.py:517-518`). One client can therefore show `Pvt Ltd` in its Level-2 folder (K03) but `PRIVATE_LIMITED` in its category folders and report names. (HIGH)
3. **Two different names can clean to the same name (RAN, S28/S29).** `...Apr 24.xlsx` and `...Apr_24.xlsx` both give `...Apr_24.xlsx`. Section 6.5 shows what that does to copied files.
4. **Text before a `/` is dropped (RAN, S24):** `Alpha/Beta Traders` -> `Beta_Traders`. On Windows the same happens with `\` (MEDIUM, see 2.2). Neither character can appear in a real Windows file name, so this matters only if the input comes from somewhere else. (LOW practical impact)
5. **Long names are cut to 190 characters total and may end in `_` (RAN, S23, S34).** Sanitising twice is then not stable (S23: second pass removes the trailing `_`). (HIGH)
6. **`.xlsx` alone becomes `xlsx` (RAN, S31)**, because Python treats a name starting with a dot as having no suffix.

### 2.4 Level-2 folder key, `create_client_state_key` (`utils/helpers.py:513-537`) (READ + RAN)

READ: it replaces whole-word `Private` -> `Pvt` and `Limited` -> `Ltd`, ignoring case (`:517-518`), and appends `-` plus `get_state_code(state)` (`:521-524`). If the key is longer than `max_length` (default 35; the Setup screen spin box allows 15-100, `gui/views/setup_view.py:375`), the client part is cut to `max_length - len(code) - 1` characters (`:527-532`). If that leaves 5 characters or fewer, the whole key is cut instead (`:533-535`), which can drop the state code. **It does not remove invalid characters** (no sanitising), so the Level-2 folder keeps spaces, `&`, commas and dots. `FileOrganizer` replaces a `max_length` of 0 or less with 35 (`core/file_organizer.py:42`).

| # | client | state | max_length | `create_client_state_key` result (Level-2 folder name) | `sanitize_filename(client)` (used in category-folder and report names) | warning logged |
|---|---|---|---|---|---|---|
| K01 | 'Alpha Test Traders' | 'Delhi' | 35 | 'Alpha Test Traders-DL' (len 21) | 'Alpha_Test_Traders' | - |
| K02 | 'Beta Demo Industries Private Limited' | 'Maharashtra' | 35 | 'Beta Demo Industries Pvt Ltd-MH' (len 31) | 'Beta_Demo_Industries_Pvt_Ltd' | - |
| K03 | 'Beta Demo Industries PRIVATE LIMITED' | 'Maharashtra' | 35 | 'Beta Demo Industries Pvt Ltd-MH' (len 31) | 'Beta_Demo_Industries_PRIVATE_LIMITED' | - |
| K04 | 'Unlimited Privateer Traders' | 'Gujarat' | 35 | 'Unlimited Privateer Traders-GJ' (len 30) | 'Unlimited_Pvter_Traders' | - |
| K05 | 'Alpha Test Traders and Distributors North' | 'Delhi' | 35 | 'Alpha Test Traders and Distribut-DL' (len 35) | 'Alpha_Test_Traders_and_Distributors_North' | - |
| K06 | 'Alpha Test Traders and Distributors South' | 'Delhi' | 35 | 'Alpha Test Traders and Distribut-DL' (len 35) | 'Alpha_Test_Traders_and_Distributors_South' | - |
| K07 | 'Alpha Test Traders' | 'Tamil Nadu' | 35 | 'Alpha Test Traders-TN' (len 21) | 'Alpha_Test_Traders' | - |
| K08 | 'Alpha Test Traders' | 'Tamilnadu' | 35 | 'Alpha Test Traders-TAM' (len 22) | 'Alpha_Test_Traders' | Unknown state 'Tamilnadu', using fallback: TAM |
| K09 | 'Alpha Test Traders' | 'New Delhi' | 35 | 'Alpha Test Traders-ND' (len 21) | 'Alpha_Test_Traders' | Unknown state 'New Delhi', using fallback: ND |
| K10 | 'Alpha Test Traders' | 'DL' | 35 | 'Alpha Test Traders-DL' (len 21) | 'Alpha_Test_Traders' | Unknown state 'DL', using fallback: DL |
| K11 | 'Alpha Test Traders' | 'Delhi' | 10 | 'Alpha T-DL' (len 10) | 'Alpha_Test_Traders' | - |
| K12 | 'Alpha Test Traders' | 'Dadra and Nagar Haveli and Daman and Diu' | 35 | 'Alpha Test Traders-DD' (len 21) | 'Alpha_Test_Traders' | - |
| K13 | 'Zeta Mock Pvt. Ltd.' | 'Maharashtra' | 35 | 'Zeta Mock Pvt. Ltd.-MH' (len 22) | 'Zeta_Mock_Pvt._Ltd.' | - |
| K14 | 'Alpha Test Traders' | '' | 35 | 'Alpha Test Traders-' (len 19) | 'Alpha_Test_Traders' | - |
| K15 | 'Alpha Test Traders' | 'Some Unknown Long Statename Here' | 6 | 'Alpha ' (len 6) | 'Alpha_Test_Traders' | Unknown state 'Some Unknown Long Statename Here', using fallback: SU |
| K16 | 'Alpha Test Traders' | 'Delhi' | 0 | '' (len 0) | 'Alpha_Test_Traders' | - |
| K17 | 'Delta Example, Sons & Co.' | 'Gujarat' | 35 | 'Delta Example, Sons & Co.-GJ' (len 28) | 'Delta_Example_Sons_&_Co.' | - |

Surprises:
- **Two different clients share one Level-2 folder (RAN, K05/K06; section 6.6 shows real files):** "...Distributors North" and "...Distributors South" in Delhi both become `Alpha Test Traders and Distribut-DL`. (HIGH)
- K15 (`max_length` 6) gives `'Alpha '`, with a trailing space and no state code. K16 (`max_length` 0) gives `''`. Neither can be reached from the app: processing reads the spin box (limited to 15-100) at `gui/pyqt_main_window.py:352`, and `FileOrganizer` replaces 0 or less with 35 (`core/file_organizer.py:42`). With 15 and a 3-letter code the client part keeps 11 characters, so the "cut the whole key" branch (`utils/helpers.py:533-535`) is unreachable from the app (READ, HIGH).

### 2.5 [GST RULE] State-name-to-code table and fallback, `get_state_code` (`utils/helpers.py:485-511`, table `utils/constants.py:234-275`) - TIME-VARYING (the list of states and union territories)

READ: the state text is trimmed and lower-cased, then looked up in `STATE_CODE_MAPPING` (`:494-499`). If it is not found: with two or more words, the first letters of the first two words, upper-cased (`:502-506`); with one word, the first three characters of the *original* text, upper-cased (`:508-511`). In both fallback cases a warning goes to the log only. Empty text is returned as-is (`:487-488`). The codes are two-letter abbreviations. Described only, not judged. The full table as it stands in the code (RAN: printed from the imported module; 39 names, 36 distinct codes):

| name key (lower case) in STATE_CODE_MAPPING | code |
|---|---|
| andhra pradesh | AP |
| arunachal pradesh | AR |
| assam | AS |
| bihar | BR |
| chhattisgarh | CG |
| goa | GA |
| gujarat | GJ |
| haryana | HR |
| himachal pradesh | HP |
| jharkhand | JH |
| karnataka | KA |
| kerala | KL |
| madhya pradesh | MP |
| maharashtra | MH |
| manipur | MN |
| meghalaya | ML |
| mizoram | MZ |
| nagaland | NL |
| odisha | OD |
| orissa | OD |
| punjab | PB |
| rajasthan | RJ |
| sikkim | SK |
| tamil nadu | TN |
| telangana | TS |
| tripura | TR |
| uttar pradesh | UP |
| uttarakhand | UK |
| west bengal | WB |
| andaman and nicobar islands | AN |
| chandigarh | CH |
| dadra and nagar haveli and daman and diu | DD |
| delhi | DL |
| national capital territory of delhi | DL |
| jammu and kashmir | JK |
| ladakh | LA |
| lakshadweep | LD |
| puducherry | PY |
| pondicherry | PY |

Codes shared by several names (RAN): `OD` (odisha, orissa), `DL` (delhi, national capital territory of delhi), `PY` (puducherry, pondicherry).

Variants (RAN):

| state text in file name | code returned by `get_state_code` | found in the table? | warning logged |
|---|---|---|---|
| 'Delhi' | 'DL' | yes | - |
| 'delhi' | 'DL' | yes | - |
| ' DELHI ' | 'DL' | yes | - |
| 'NCT of Delhi' | 'NO' | no (fallback) | Unknown state 'NCT of Delhi', using fallback: NO |
| 'National Capital Territory of Delhi' | 'DL' | yes | - |
| 'New Delhi' | 'ND' | no (fallback) | Unknown state 'New Delhi', using fallback: ND |
| 'DL' | 'DL' | no (fallback) | Unknown state 'DL', using fallback: DL |
| '07' | '07' | no (fallback) | Unknown state '07', using fallback: 07 |
| 'Tamil Nadu' | 'TN' | yes | - |
| 'Tamilnadu' | 'TAM' | no (fallback) | Unknown state 'Tamilnadu', using fallback: TAM |
| 'Tamil  Nadu' | 'TN' | no (fallback) | Unknown state 'Tamil  Nadu', using fallback: TN |
| 'Orissa' | 'OD' | yes | - |
| 'Odisha' | 'OD' | yes | - |
| 'Pondicherry' | 'PY' | yes | - |
| 'Puducherry' | 'PY' | yes | - |
| 'Uttaranchal' | 'UTT' | no (fallback) | Unknown state 'Uttaranchal', using fallback: UTT |
| 'Telangana' | 'TS' | yes | - |
| 'Jammu & Kashmir' | 'J&' | no (fallback) | Unknown state 'Jammu & Kashmir', using fallback: J& |
| 'Jammu and Kashmir' | 'JK' | yes | - |
| 'J&K' | 'J&K' | no (fallback) | Unknown state 'J&K', using fallback: J&K |
| 'Andaman & Nicobar' | 'A&' | no (fallback) | Unknown state 'Andaman & Nicobar', using fallback: A& |
| 'Dadra and Nagar Haveli' | 'DA' | no (fallback) | Unknown state 'Dadra and Nagar Haveli', using fallback: DA |
| 'Daman and Diu' | 'DA' | no (fallback) | Unknown state 'Daman and Diu', using fallback: DA |
| 'Ladakh' | 'LA' | yes | - |
| 'Other Territory' | 'OT' | no (fallback) | Unknown state 'Other Territory', using fallback: OT |
| 'Centre Jurisdiction' | 'CJ' | no (fallback) | Unknown state 'Centre Jurisdiction', using fallback: CJ |
| '' | '' | no (fallback) | - |
| 'X' | 'X' | no (fallback) | Unknown state 'X', using fallback: X |
| 'Maharashtra-West' | 'MAH' | no (fallback) | Unknown state 'Maharashtra-West', using fallback: MAH |
| 'Gujarat' | 'GJ' | yes | - |

Fallbacks for other unknown names (RAN):

| unknown state text | fallback code | same as the code of a listed state? |
|---|---|---|
| Delta | DEL | no |
| Goa Region | GR | no |
| Mahe | MAH | no |
| Karaikal | KAR | no |
| Yanam | YAN | no |
| Andhra | AND | no |
| Tamil | TAM | no |
| Uttar | UTT | no |
| Madhya | MAD | no |
| West | WES | no |
| Bengal | BEN | no |
| Kerala State | KS | no |
| Himachal | HIM | no |
| Arunachal | ARU | no |
| Gujarat State | GS | no |
| Delhi NCR | DN | no |
| Tamilnadu | TAM | no |
| Telangana State | TS | telangana |

Technical consequences (RAN): the client key is `client-code` (`core/file_parser.py:261-262`). `Delhi` and `DL` give the same key and are merged into one client (section 6). `Tamil Nadu` and `Tamilnadu` give `TN` and `TAM`, so the same client is split into two (section 6). `Dadra and Nagar Haveli` and `Daman and Diu` both fall back to `DA`. The two-digit text `07` stays `07`. Whether any of this is wanted is for the author.

### 2.6 `validate_client_name` (`utils/helpers.py:321-342`) (RAN)

It is used for both client and state (`core/file_parser.py:106-107`). Its result is only a warning stored on the parsed record, and the file is still used.

| input | length | valid? | message |
|---|---|---|---|
| 'Alpha Test Traders' | 18 | True | - |
| '' | 0 | False | Client name cannot be empty |
| '   ' | 3 | False | Client name cannot be empty |
| 'A' | 1 | False | Client name too short |
| 'AB' | 2 | True | - |
| <100 chars: 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'...'xxxxxxxxxxxxxxxxxxxx'> | 100 | True | - |
| <101 chars: 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'...'xxxxxxxxxxxxxxxxxxxx'> | 101 | False | Client name too long |
| 'Alpha<Beta' | 10 | False | Client name contains invalid characters |
| 'Alpha>Beta' | 10 | False | Client name contains invalid characters |
| 'Alpha"Beta' | 10 | False | Client name contains invalid characters |
| 'Alpha/Beta' | 10 | False | Client name contains invalid characters |
| 'Alpha\\Beta' | 10 | False | Client name contains invalid characters |
| 'Alpha\|Beta' | 10 | False | Client name contains invalid characters |
| 'Alpha?Beta' | 10 | False | Client name contains invalid characters |
| 'Alpha*Beta' | 10 | False | Client name contains invalid characters |
| 'Alpha:Beta' | 10 | False | Client name contains invalid characters |
| 'Alpha & Co.' | 11 | True | - |
| ' A' | 2 | True | - |
| 'Alpha\tTraders' | 13 | True | - |
| 'अल' | 2 | True | - |

The same function applied to the state (RAN): `GSTR3B-Alpha Test Traders-D-Apr.xlsx` -> warning `State name issue: Client name too short`; `...-Del:hi-...` -> `State name issue: Client name contains invalid characters`. Tabs and leading spaces pass. The length check runs on the untrimmed text (`' A'` passes as 2 characters).

---

## 3. Financial year and tax period logic (Part 3) - [GST RULE], TIME-VARYING

**No financial-year or tax-period calculation exists anywhere in B's Python code (READ, HIGH).** I searched every `.py` file, case-insensitively, for `financial|fiscal|FY|fy_|tax_period|period|2023|2024|2025|april|march|month|year|refreshed`, and for every `strptime/strftime/fromtimestamp/datetime.now/time.time`. Period words appear only as regex group names and help texts (`utils/constants.py:19-60`, `core/file_parser.py:322-332`, `gui/views/setup_view.py:493-498`). The Excel templates (not in the repository) may contain such logic; I could not check them.

What does exist:

| Item | Where | What it does (RAN unless marked) |
|---|---|---|
| [GST RULE] period/date/month/year text from file names | `core/file_parser.py:93-102`, groups at `utils/constants.py:20,28,36,44,52,60` | Stored as text in `metadata`. Never validated, never compared and never read again (READ: grep). RAN examples: `ImsReco-...-31022025` -> date `'31022025'`; `00000000` accepted; GSTR3B month `'13'` and `'March 2030'` accepted; Sales `Mar25-Apr24` (end before start) accepted; AnnualReport year `'1999'` accepted. |
| [GST RULE] GSTR-3B may have several files | `core/file_parser.py:292-297` (comment `:294` "monthly filings - unlimited") | Every other type with two or more files gets "Multiple <type> (n files)" and the status "Has duplicates". `GSTR-3B Export` is skipped in that check. RAN: a client with GSTR-3B files for "Apr 2023" and "Apr 2024" gets no remark, while two AnnualReport files for `2023-24` and `2024-25` give `Multiple Annual Report (2 files)`. TIME-VARYING (how often a return is filed). |
| [GST RULE] the six "expected" file types | `utils/constants.py:68-75`, used at `core/file_parser.py:284-303` | Missing types are listed, and completeness % = (number of types found / 6) x 100, rounded to one decimal place (`:302-303`). |
| [GST RULE] "critical" types `Annual Report`, `GSTR-3B Export` | `core/file_parser.py:387` in `validate_results` | `validate_results` is **never called** (READ: grep, HIGH). RAN directly, it reports "Missing Critical Files" warnings (section 6.2). |
| Folder/backup/report time stamps | `utils/helpers.py:23-33`, `power_query_extractor/core/report_processor.py:92`, `power_query_extractor/core/data_consolidator.py:163` | Day first: `DDMMYY HHMM` (folders), `DDMMYY_HHMM` (backups, refreshed copies); the extraction report uses `YYYYmmdd_HHMM`. RAN: 31-03-2025 23:59:59 -> `310325 2359`; 01-04-2025 00:00:01 -> `010425 0000`; 01-01-2026 09:05 -> `010126 0905`. These are clock times, not tax periods. Resolution is one minute (consequences in sections 4 and 6). |
| Which Level-1 folder counts as "latest" | `core/file_organizer.py:160-169` versus `gui/pyqt_main_window.py:432-448` | Two different rules. The organiser (rerun/resume) uses the newest *modification time*; the Extract tab uses the *name* sorted from Z to A. Day-first names do not sort by date. RAN: with `Annual Statement-311225 1000` (older) and `Annual Statement-010126 0900` (newer), the organiser picks `010126 0900` and the Extract tab picks `311225 1000`. Version folders in the Extract tree are also sorted by name, Z to A (`gui/pyqt_main_window.py:457`). RAN order: `311225, 280226, 150126, 010326, 010126` against the true newest-first order `010326, 280226, 150126, 010126, 311225`. |

---

## 4. Cell mappings and the consolidated extraction report (Part 4)

### 4.1 [GST RULE] Every mapping (RAN: printed from the imported module; source `power_query_extractor/config/cell_mappings.py:33-242`) - TIME-VARYING (sheet names `T6 of R9` / `T8 of R9`, table labels 6J / 8D, column labels)

| # | Output sheet in the consolidated report | report_type | input sheet (in ITC/Sales report) | input cell | output column header | calculation | internal field name (`CELL_MAPPINGS`) |
|---|---|---|---|---|---|---|---|
| 1 | GSTR-3B Difference | ITC | Diff GSTR-3B | J15 | Diff IGST | None | GSTR-3B_Difference_Diff_IGST |
| 2 | GSTR-3B Difference | ITC | Diff GSTR-3B | K15 | Diff CGST | None | GSTR-3B_Difference_Diff_CGST |
| 3 | GSTR-3B Difference | ITC | Diff GSTR-3B | L15 | Diff SGST | None | GSTR-3B_Difference_Diff_SGST |
| 4 | GSTR-1 Difference | Sales | Diff - GSTR-1 | N1 | Diff Taxable Value | None | GSTR-1_Difference_Diff_Taxable_Value |
| 5 | GSTR-1 Difference | Sales | Diff - GSTR-1 | O1 | Diff IGST | None | GSTR-1_Difference_Diff_IGST |
| 6 | GSTR-1 Difference | Sales | Diff - GSTR-1 | P1 | Diff CGST | None | GSTR-1_Difference_Diff_CGST |
| 7 | GSTR-1 Difference | Sales | Diff - GSTR-1 | Q1 | Diff SGST | None | GSTR-1_Difference_Diff_SGST |
| 8 | Turnover GSTR-1 | Sales | Diff - GSTR-1 | I1 | Taxable Value | None | Turnover_GSTR-1_Taxable_Value |
| 9 | Reco Diff | ITC | Reco | C37 | IGST Diff | None | Reco_Diff_IGST_Diff |
| 10 | Reco Diff | ITC | Reco | D37 | CGST Diff | None | Reco_Diff_CGST_Diff |
| 11 | Reco Diff | ITC | Reco | E37 | SGST Diff | None | Reco_Diff_SGST_Diff |
| 12 | Table 6J Match | ITC | T6 of R9 | K29 | Table 6J Match | None | Table_6J_Match_Table_6J_Match |
| 13 | Table 8D Limits | ITC | T8 of R9 | G6 | Table 8D Limits | None | Table_8D_Limits_Table_8D_Limits |
| 14 | IMS Reco | ITC | IMS_Reco | AA1 | IGST | None | IMS_Reco_IGST |
| 15 | IMS Reco | ITC | IMS_Reco | AD1 | CGST | None | IMS_Reco_CGST |
| 16 | IMS Reco | ITC | IMS_Reco | AG1 | SGST | None | IMS_Reco_SGST |
| 17 | Extra In Purchase | ITC | ExtraInPurchase | O1 | IGST | None | Extra_In_Purchase_IGST |
| 18 | Extra In Purchase | ITC | ExtraInPurchase | P1 | CGST | None | Extra_In_Purchase_CGST |
| 19 | Extra In Purchase | ITC | ExtraInPurchase | Q1 | SGST | None | Extra_In_Purchase_SGST |
| 20 | Extra In GSTR2B | ITC | ExtraInGSTR2B | S1 | IGST | None | Extra_In_GSTR2B_IGST |
| 21 | Extra In GSTR2B | ITC | ExtraInGSTR2B | T1 | CGST | None | Extra_In_GSTR2B_CGST |
| 22 | Extra In GSTR2B | ITC | ExtraInGSTR2B | U1 | SGST | None | Extra_In_GSTR2B_SGST |

All 22 mappings have `calculation: None`. None of the 9 output sheets has a `description` key, so both of those code paths are unused by the shipped configuration (READ; RAN in 4.10 with a custom configuration). Options (`:248-253`): `skip_refresh False` (overwritten from the screen's check box at `gui/controllers/extraction_controller.py:55`), `auto_refresh_if_missing True` and `show_timestamp False`. The last two are **never read** anywhere (READ: grep, HIGH). `missing_data_text 'Missing'` is read at `power_query_extractor/core/data_consolidator.py:117,142`.

The old-format table built at import time (`cell_mappings.py:259-277`, RAN):
`CELL_MAPPINGS = {'ITC': {'Diff GSTR-3B': {J15,K15,L15}, 'Reco': {C37,D37,E37}, 'T6 of R9': {K29}, 'T8 of R9': {G6}, 'IMS_Reco': {AA1,AD1,AG1}, 'ExtraInPurchase': {O1,P1,Q1}, 'ExtraInGSTR2B': {S1,T1,U1}}, 'Sales': {'Diff - GSTR-1': {N1,O1,P1,Q1,I1}}}`. The field names are `<output sheet>_<output column>` with spaces replaced by `_`, as in the last column above.

### 4.2 How the data flows (READ)

1. The Extract tab builds one entry per ticked version folder: `{'name': '<Level-2 folder> / <Version folder>', 'client_key': <Level-2 folder>, 'latest_version': <path>, 'process_itc', 'process_sales'}` (`gui/pyqt_main_window.py:517-528`, names from `gui/views/extraction_view.py:384-385`).
2. `ReportProcessor.process_client` (`power_query_extractor/core/report_processor.py:39-70`). For each report it calls `_process_report` (`:72-163`), which:
   - takes the *first* file found by glob `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` in the version folder (`:76-89`);
   - copies it to `<stem><suffix>.xlsx`, where the suffix is `_Refreshed_{timestamp}` with `{timestamp}` = `DDMMYY_HHMM` (`:92-94, 119-120`);
   - refreshes the **copy** through COM and keystrokes, Alt, R, A (`:165-366`);
   - checks the refresh with `_validate_refresh` (`:589-636`, cell `Info!BB2`);
   - reads the mapped cells from the copy (`_extract_data`, `:478-587`). Each value is stored twice, under its field name and under its bare cell reference (`:536, 539`). A missing sheet gives `None` for its fields (`:521-527`).
3. `DataConsolidator.create_report(results, level1_folder)` (`power_query_extractor/core/data_consolidator.py:20-174`, called at `gui/controllers/extraction_controller.py:92`) writes `PQ_Extraction_Report_YYYYmmdd_HHMM.xlsx` into the Level-1 folder, with one sheet per output sheet plus `Summary` first.
   - Column A is `result['client']` (the "Level-2 / Version" text). Column B is ✓/✗ from `status['success']` (`:97-109`); that status is the **refresh** status, not "values found".
   - If ✗, every value cell is `Missing` (`:116-117`).
   - Otherwise each value is looked up by **substring**: the first field name in the extracted data that *contains* the output column text (`:130-133`). Only if that gives `None` is the exact field name tried (`:136-138`). `None` or `''` becomes `Missing` (`:141-142`).
   - Numbers other than 0 get the format `#,##0.00` (`:149-150`).

### 4.3 The made-up workbooks (created with openpyxl in `work/t4/...`; fake numbers)

ITC report (sheets as named in the mappings):

| sheet | cells set | purpose |
|---|---|---|
| `Diff GSTR-3B` | J15=1500.5, K15=-200, L15=0 | normal number, negative, zero |
| `Reco` | C37=`'12,345'` (text), D37 empty, E37=`=1+1` (formula never calculated: no stored result) | text in a number cell, empty cell, never-calculated formula |
| `T6 of R9` | **sheet missing** | missing sheet |
| `T8 of R9` | G6=`=100*43.21` with a stored result 4321 injected into the file, as if Excel had saved it | calculated formula |
| `IMS_Reco` | AA1=111, AD1=222, AG1=333 | distinct values to trace |
| `ExtraInPurchase` | O1=10, P1=20, Q1=30 | |
| `ExtraInGSTR2B` | S1 empty, T1=5, U1=6 | |
| `Info` | BB2 = client name text (client 3: empty) | refresh validation |

Sales report: `Diff - GSTR-1` I1=987654.32, N1=0, O1=`'abc'` (text), P1 empty, Q1=-0.01; sheet `info` (lower case) BB2 = text (client 3: `'#N/A'` text).

Four fake clients in `Annual Statement-211026 1000`:
- C1 `Alpha Test Traders-DL`: both reports complete.
- C2 `Beta Demo Industries Pvt Ltd-MH`: ITC report without the `Diff GSTR-3B` sheet, IMS 7/8/9; Sales not ticked.
- C3 `Gamma Sample LLP-TN`: ITC `Info!BB2` empty, Sales `Info!BB2` = `#N/A`.
- C4 `Delta Example & Co-GJ`: version folder with no reports.

What openpyxl itself reads from the C1 ITC file (RAN): with `data_only=True`, `Reco!C37='12,345'`, `D37=None`, `E37=None`, `T8 of R9!G6=4321`; with `data_only=False`, `E37='=1+1'`, `G6='=100*43.21'`.

### 4.4 Results with the stand-in reading stored results (`data_only=True`, closest to COM) (RAN, `scripts/t4_extract.py`)

Status and data per client (RAN):
- C1 ITC: `{'success': True, 'error': None, 'validation_message': 'Validated - Client Name: Alpha Test Traders'}`; data `GSTR-3B_Difference_Diff_IGST 1500.5, ..._CGST -200, ..._SGST 0, Reco_Diff_IGST_Diff '12,345', Reco_Diff_CGST_Diff None, Reco_Diff_SGST_Diff None, Table_6J_Match_Table_6J_Match None, Table_8D_Limits_Table_8D_Limits 4321, IMS_Reco_IGST 111, _CGST 222, _SGST 333, Extra_In_Purchase_* 10/20/30, Extra_In_GSTR2B_IGST None, _CGST 5, _SGST 6`, plus the same values under the bare cell references.
- C1 Sales: success; `..._Diff_Taxable_Value 0, ..._Diff_IGST 'abc', ..._Diff_CGST None, ..._Diff_SGST -0.01, Turnover_GSTR-1_Taxable_Value 987654.32`.
- C2 ITC: success; the three `GSTR-3B_Difference_*` fields are `None` (sheet missing; no bare-cell keys for them); IMS 7/8/9. C2 Sales: `{'success': False, 'error': 'Not selected for processing'}`, data `{}`.
- C3 ITC: `{'success': False, 'error': 'Refresh validation failed: Cell BB2 is empty', ...}`, **but all values were extracted** (same as C1). C3 Sales: `'Refresh validation failed: Cell BB2 contains error: #N/A'`, values extracted.
- C4: `'No ITC report found'` / `'No Sales report found'`, data `{}`.
- Log warnings (RAN): `Sheet 'T6 of R9' not found` (x3 clients), `Sheet 'Diff GSTR-3B' not found` (C2).
- Fake-COM call order per report (RAN): open copy -> SendKeys `%a`, `r`, `a` -> Save -> Close(SaveChanges=True) -> Close(SaveChanges=False) (the second close comes from the `finally` block at `:345-353`; with real COM this second close of an already-closed workbook would raise an error that is silently swallowed: READ, MEDIUM) -> Quit -> open again for extraction -> Close(False) -> Quit. Only the `_Refreshed_` copies were ever opened. Original reports: unchanged checksums (RAN: `existing_files_changed: []`).
- New files (RAN): one `<original stem>_Refreshed_211026_HHMM.xlsx` per processed report next to the original (5 files), plus `PQ_Extraction_Report_20261021_HHMM.xlsx` in the Level-1 folder.

The consolidated report it wrote (RAN; row 1 title, row 3 headers, rows 4-7 = C1-C4; `[fmt ...]` = number format applied):

**Sheet `Summary`** (inserted first): 'Extraction Summary'; 'Generated: 2026-10-21 03:55:24'; 'Total Clients: 4'; 'ITC Reports Processed: 2/4'; 'Sales Reports Processed: 1/4'; 'Sheets Created:'; '• GSTR-3B Difference'; '• GSTR-1 Difference'; '• Turnover GSTR-1'; '• Reco Diff'; '• Table 6J Match'; '• Table 8D Limits'; '• IMS Reco'; '• Extra In Purchase'; '• Extra In GSTR2B'

**Sheet `GSTR-3B Difference`** (merged cells: A1:C1)

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'GSTR-3B Difference Report' |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Diff IGST' | 'Diff CGST' | 'Diff SGST' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 1500.5 [fmt #,##0.00] | -200 [fmt #,##0.00] | 0 |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | 'Missing' | 'Missing' | 'Missing' |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |

**Sheet `GSTR-1 Difference`** (merged cells: A1:C1)

| row | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| 1 | 'GSTR-1 Difference Report' |  |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Diff Taxable Value' | 'Diff IGST' | 'Diff CGST' | 'Diff SGST' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 0 | 'abc' | 'Missing' | -0.01 [fmt #,##0.00] |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' | 'Missing' |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' | 'Missing' |

**Sheet `Turnover GSTR-1`** (merged cells: A1:C1)

| row | A | B | C |
|---|---|---|---|
| 1 | 'Turnover GSTR-1 Report' |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Taxable Value' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 987654.32 [fmt #,##0.00] |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✗' | 'Missing' |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' |

**Sheet `Reco Diff`** (merged cells: A1:C1)

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'Reco Diff Report' |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'IGST Diff' | 'CGST Diff' | 'SGST Diff' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | '12,345' | 'Missing' | 'Missing' |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | '12,345' | 'Missing' | 'Missing' |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |

**Sheet `Table 6J Match`** (merged cells: A1:C1)

| row | A | B | C |
|---|---|---|---|
| 1 | 'Table 6J Match Report' |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Table 6J Match' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 'Missing' |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | 'Missing' |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' |

**Sheet `Table 8D Limits`** (merged cells: A1:C1)

| row | A | B | C |
|---|---|---|---|
| 1 | 'Table 8D Limits Report' |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Table 8D Limits' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 4321 [fmt #,##0.00] |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | 4321 [fmt #,##0.00] |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' |

**Sheet `IMS Reco`** (merged cells: A1:C1)

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'IMS Reco Report' |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'IGST' | 'CGST' | 'SGST' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 1500.5 [fmt #,##0.00] | -200 [fmt #,##0.00] | 0 |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | 7 [fmt #,##0.00] | 8 [fmt #,##0.00] | 9 [fmt #,##0.00] |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |

**Sheet `Extra In Purchase`** (merged cells: A1:C1)

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'Extra In Purchase Report' |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'IGST' | 'CGST' | 'SGST' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 1500.5 [fmt #,##0.00] | -200 [fmt #,##0.00] | 0 |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | 10 [fmt #,##0.00] | 20 [fmt #,##0.00] | 30 [fmt #,##0.00] |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |

**Sheet `Extra In GSTR2B`** (merged cells: A1:C1)

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'Extra In GSTR2B Report' |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'IGST' | 'CGST' | 'SGST' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 1500.5 [fmt #,##0.00] | -200 [fmt #,##0.00] | 0 |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | 'Missing' | 5 [fmt #,##0.00] | 6 [fmt #,##0.00] |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |

### 4.5 What happened to each edge case (RAN, from the table above)

| Edge case | What was in the file | What the report shows | Why |
|---|---|---|---|
| Missing sheet (`T6 of R9`) | no sheet | `Missing`, and column B still ✓ | `_extract_data` stores `None` (`report_processor.py:521-527`); ✓ is the refresh status (`data_consolidator.py:97-98`) |
| Missing sheet `Diff GSTR-3B` (C2) | no sheet | `GSTR-3B Difference` row: ✓ with `Missing` x3 | same |
| Empty cell (`Reco!D37`, `Diff - GSTR-1!P1`, `ExtraInGSTR2B!S1`) | empty | `Missing` | `None` -> `Missing` (`data_consolidator.py:141-142`) |
| Text in a number cell (`Reco!C37 '12,345'`, `Diff - GSTR-1!O1 'abc'`) | text | the text itself, written as text, no number format | value written as-is (`:145`), format only for int/float (`:149`) |
| Zero (`L15=0`, `N1=0`) | 0 | `0`, without the `#,##0.00` format | `value != 0` test at `:149` |
| Never-calculated formula (`Reco!E37 =1+1`) | formula, no stored result | `Missing` (stand-in). With real Excel: probably `2`, because Excel recalculates on opening (not run; MEDIUM) | openpyxl `data_only=True` gives `None` |
| Calculated formula with a stored result (`T8 of R9!G6`) | 4321 stored | `4321` | stored value read |
| Refresh validation failed (C3) | values present | ✗ and `Missing` everywhere, although the values were extracted | `:116-117` ignores the data when the refresh status is unsuccessful |
| Report not ticked (C2 Sales) | - | ✗ and `Missing`: looks the same as a failure | `process_client` sets `success False` (`report_processor.py:60-68`) |
| No report file (C4) | - | ✗ and `Missing` | `:79-87` |

### 4.6 The substring look-up writes the wrong figures into three sheets (RAN, HIGH)

The configuration says the `IMS Reco` sheet takes `IMS_Reco!AA1/AD1/AG1`, `Extra In Purchase` takes `ExtraInPurchase!O1/P1/Q1` and `Extra In GSTR2B` takes `ExtraInGSTR2B!S1/T1/U1` (`cell_mappings.py:156-228`). In the run above, **for C1 all three sheets show 1500.5 / -200 / 0, the values of `Diff GSTR-3B!J15/K15/L15`**, instead of 111/222/333, 10/20/30 and Missing/5/6. For C2, whose ITC report has no `Diff GSTR-3B` sheet, the same three sheets show the configured values (7/8/9, 10/20/30, Missing/5/6).

Why (technical): these three sheets use the output-column names `IGST`, `CGST`, `SGST`, with no space. The look-up at `data_consolidator.py:130-133` takes the **first** extracted field whose name *contains* that text. The extracted data is filled in mapping order, so `GSTR-3B_Difference_Diff_IGST` comes first and contains `IGST`. Only when that first match is `None` (sheet missing or cell empty) does the exact-name fallback at `:136-138` find the right field. Output columns that contain a space (`Diff IGST`, `IGST Diff`, `Taxable Value`, `Table 6J Match` ...) never match by substring, because the field names have `_` instead of spaces, so those sheets always fall through to the exact name and are correct.

Confirmed a second way (RAN, `scripts/t4b_consolidator.py` part C), with distinct values 1..45 per field: `IMS Reco`, `Extra In Purchase` and `Extra In GSTR2B` all showed 1/2/3, the GSTR-3B difference values. Feeding the **same data with the keys in reverse order** made all three show 31/32/33, the `Extra_In_GSTR2B` values. Every other sheet showed its own values in both orders. The figure written into these three sheets depends on the order of the extracted data, not on the configured cell. [GST RULE] context: which return figure ends up under which heading. I describe it only; the author decides what the sheets should contain.

### 4.7 Results with the stand-in returning formula text (`data_only=False`) and the second run (RAN)

The same four clients were extracted a second time in the same minute, with formulas returned as text. The report differs only here: `Reco Diff` E4/E5 = `'=1+1'` and `Table 8D Limits` C4/C5 = `'=100*43.21'`, and openpyxl stores these strings as **live formulas** in the output report (RAN, see the value-type table in 4.10: a text value starting with `=` is written with cell type `f`, formula). Real COM returns values, not formula text, so this mode shows what happens if a cell's *text* starts with `=` (LOW likelihood).

Full second-run report:

**Sheet `Summary`** (inserted first): 'Extraction Summary'; 'Generated: 2026-10-21 03:55:24'; 'Total Clients: 4'; 'ITC Reports Processed: 2/4'; 'Sales Reports Processed: 1/4'; 'Sheets Created:'; '• GSTR-3B Difference'; '• GSTR-1 Difference'; '• Turnover GSTR-1'; '• Reco Diff'; '• Table 6J Match'; '• Table 8D Limits'; '• IMS Reco'; '• Extra In Purchase'; '• Extra In GSTR2B'

**Sheet `GSTR-3B Difference`** (merged cells: A1:C1)

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'GSTR-3B Difference Report' |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Diff IGST' | 'Diff CGST' | 'Diff SGST' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 1500.5 [fmt #,##0.00] | -200 [fmt #,##0.00] | 0 |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | 'Missing' | 'Missing' | 'Missing' |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |

**Sheet `GSTR-1 Difference`** (merged cells: A1:C1)

| row | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| 1 | 'GSTR-1 Difference Report' |  |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Diff Taxable Value' | 'Diff IGST' | 'Diff CGST' | 'Diff SGST' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 0 | 'abc' | 'Missing' | -0.01 [fmt #,##0.00] |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' | 'Missing' |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' | 'Missing' |

**Sheet `Turnover GSTR-1`** (merged cells: A1:C1)

| row | A | B | C |
|---|---|---|---|
| 1 | 'Turnover GSTR-1 Report' |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Taxable Value' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 987654.32 [fmt #,##0.00] |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✗' | 'Missing' |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' |

**Sheet `Reco Diff`** (merged cells: A1:C1)

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'Reco Diff Report' |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'IGST Diff' | 'CGST Diff' | 'SGST Diff' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | '12,345' | 'Missing' | '=1+1' |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | '12,345' | 'Missing' | '=1+1' |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |

**Sheet `Table 6J Match`** (merged cells: A1:C1)

| row | A | B | C |
|---|---|---|---|
| 1 | 'Table 6J Match Report' |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Table 6J Match' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 'Missing' |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | 'Missing' |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' |

**Sheet `Table 8D Limits`** (merged cells: A1:C1)

| row | A | B | C |
|---|---|---|---|
| 1 | 'Table 8D Limits Report' |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Table 8D Limits' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | '=100*43.21' |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | '=100*43.21' |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' |

**Sheet `IMS Reco`** (merged cells: A1:C1)

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'IMS Reco Report' |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'IGST' | 'CGST' | 'SGST' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 1500.5 [fmt #,##0.00] | -200 [fmt #,##0.00] | 0 |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | 7 [fmt #,##0.00] | 8 [fmt #,##0.00] | 9 [fmt #,##0.00] |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |

**Sheet `Extra In Purchase`** (merged cells: A1:C1)

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'Extra In Purchase Report' |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'IGST' | 'CGST' | 'SGST' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 1500.5 [fmt #,##0.00] | -200 [fmt #,##0.00] | 0 |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | 10 [fmt #,##0.00] | 20 [fmt #,##0.00] | 30 [fmt #,##0.00] |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |

**Sheet `Extra In GSTR2B`** (merged cells: A1:C1)

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'Extra In GSTR2B Report' |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'IGST' | 'CGST' | 'SGST' |
| 4 | 'Alpha Test Traders-DL / Version-211026 1000' | '✓' | 1500.5 [fmt #,##0.00] | -200 [fmt #,##0.00] | 0 |
| 5 | 'Beta Demo Industries Pvt Ltd-MH / Version-211026 1000' | '✓' | 'Missing' | 5 [fmt #,##0.00] | 6 [fmt #,##0.00] |
| 6 | 'Gamma Sample LLP-TN / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Example & Co-GJ / Version-211026 1000' | '✗' | 'Missing' | 'Missing' | 'Missing' |

**The second run picked up the first run's refreshed copies as its "original" reports (RAN).** The glob `ITC_Report_*.xlsx` also matches `ITC_Report_..._Refreshed_211026_0355.xlsx`, and `_process_report` takes element `[0]` of a list whose order is whatever the operating system returns (`report_processor.py:77, 89`). On this Linux machine the refreshed copy came first for C2 and C3, and the second run created `..._Refreshed_211026_0355_Refreshed_211026_0355.xlsx`. For C1 the original came first. Its new copy had the same name as the first run's copy (same minute), so `shutil.copy2` silently overwrote it. On Windows with NTFS (the Windows disk format) the folder listing is normally in alphabetical order, and with the default suffix `.` (in `..._1000.xlsx`) sorts before `_` (in `..._1000_Refreshed_...`), so the original usually comes first. With a suffix that starts with a space or `-`, or on a network/USB drive, it may not (MEDIUM, not run on Windows).

### 4.8 Refreshed-copy naming, skip-refresh and custom suffixes (RAN, `scripts/t4_extract.py` section 4)

| Situation (version folder contents) | Setting | Original chosen | File read / written | Status |
|---|---|---|---|---|
| `ITC_Report_..._1000.xlsx` + an older `..._1000_Refreshed_201026_0900.xlsx` | skip off | **the older refreshed copy** (Linux order: refreshed listed first) | new `..._Refreshed_201026_0900_Refreshed_211026_0355.xlsx` | refresh stand-in |
| same | skip **on** | the older refreshed copy | the newest `<that stem>_Refreshed_*.xlsx` = the double-suffixed file | `{'success': True, 'error': None, 'validation_message': 'Skipped - Using existing refreshed file'}` |
| `ITC_Report_X.xlsx` + `ITC_Report_X_PQ_201026_0900.xlsx`, suffix pattern `_PQ_{timestamp}` | skip on | `ITC_Report_X.xlsx` | existing `_PQ_` copy **not recognised** (skip looks only for `_Refreshed_*`, `report_processor.py:102`), so a new `ITC_Report_X_PQ_211026_0355.xlsx` was made and refreshed | - |
| `ITC_Report_Y.xlsx`, suffix `_Refreshed` (no `{timestamp}`) run twice | skip off | `ITC_Report_Y.xlsx` both times | `ITC_Report_Y_Refreshed.xlsx` overwritten by the second run (folder still holds 2 files) | - |
| only `ITC_Report_Z_Refreshed_201026_0900.xlsx` (original deleted) | skip off | the refreshed file | `..._Refreshed_201026_0900_Refreshed_211026_0355.xlsx` | - |

The Extract tab's "Refreshed/Pending" status also looks only for `*_Refreshed*` (`gui/pyqt_main_window.py:467-479`, READ), so a custom suffix always shows "Pending" (HIGH by reading).

### 4.9 Refresh check `_validate_refresh` (`report_processor.py:589-636`) (RAN with minimal fake workbooks)

| `Info` sheet / `BB2` content | Result |
|---|---|
| no sheet named info (any case) | `(False, 'Info sheet not found')` |
| None | `(False, 'Cell BB2 is empty')` |
| `'   '` | `(False, 'Cell BB2 is blank')` |
| 0 or 0.0 | `(False, 'Cell BB2 contains 0')` |
| `'#N/A'` / `'#n/a'` | `(False, 'Cell BB2 contains error: #N/A')` / `(... #n/a)` |
| `-2146826246` (the number real COM returns for `#N/A`, from my knowledge of pywin32: MEDIUM) | **`(True, 'Validated - Client Name: -2146826246')`** |
| `False` | `(True, 'Validated - Client Name: False')` |
| `'Alpha Test Traders'` | `(True, 'Validated - Client Name: Alpha Test Traders')` |

The check only tests that BB2 is filled; it never compares BB2 with the client (READ `:605-632`, RAN: C2's file held another client's name and passed).

### 4.10 Consolidator alone on synthetic inputs (RAN, `scripts/t4b_consolidator.py`)

- **Empty results list**: report created; `Summary` says `Total Clients: 0`, `ITC Reports Processed: 0/0`, and every sheet has headers only. (In the app, `create_report` is only called when there is at least one result: `extraction_controller.py:87`.)
- **Value types** written into `GSTR-3B Difference / Diff IGST`:

| value given | cell written (type) | number format |
|---|---|---|
| `0`, `0.0` | `0` (number) | none |
| `''` | `'Missing'` | - |
| `'   '` | `'   '` (text, not treated as missing) | - |
| `'abc'` | `'abc'` | - |
| `-5`, `1e12` | number | `#,##0.00` |
| `True` | `TRUE` (boolean) | `#,##0.00` (because Python counts True as a whole number) |
| `False` | `FALSE` (boolean) | none (equals 0) |
| `None` | `'Missing'` | - |
| `datetime(2025,3,31)` | date | `yyyy-mm-dd h:mm:ss` |
| `'#N/A'` text | written as an Excel **error** cell | - |
| `-2146826246` (COM-style error number) | `-2,146,826,246.00` (number) | `#,##0.00` |
| `'=SUM(A1:A2)'` text | written as a **live formula** | - |
| `'1500.50'` text | text, not converted | - |
| tuple `(1, 2)` (what COM returns for a multi-cell range) | **whole report fails**: `ValueError: Cannot convert (1, 2) to Excel`. In the app this is caught at `extraction_controller.py:94-95` as "Report error", so no report is produced | - |

- **`description` and `calculation` branches** (custom configuration put into the module for this test only): three description lines go in A1:A3, each merged A:J, with the headers on row 5 (`start_row = len + 2`, `:65`). Calculations with `A1=100, A2=40`: subtract 60, add 140, multiply 4000, divide 2.5; divide by a missing cell -> `'Div/0'`; unknown type `'percent'` -> `'Unknown'`; text operand `'1,000'` -> `'Error'`; result 0 written without number format. With the refresh failed, all `Missing`. Direct calls of `_calculate_value` (`:176-206`): missing operand -> 0; `''` -> 0; `'12.5'` text -> converted (16.5); `True` -> 1; calculation given as text instead of a dict -> `'Error'`. Calculations look values up by **bare cell reference** (`extracted_values.get('A1')`, `:182-183`); `_extract_data` stores bare references for every sheet in one shared dict (`report_processor.py:539`), so the same cell reference on two sheets of one report would collide. The shipped configuration has none (READ, HIGH).
- **Same minute, two reports**: both named `PQ_Extraction_Report_20261021_0356.xlsx`; the second **overwrote** the first (only one file left, and A4 holds the second run's client).
- `output_folder` given as text instead of a path -> `TypeError` (the app always passes a path: `extraction_controller.py:28-29`). A result whose `itc` is `None` -> `AttributeError` (`process_client` never produces that).
- **Unused helpers** `_write_section` (`:237-339`) and `_create_details_sheet` (`:341-399`) are never called (READ: grep). Run directly, both still work; `_write_section` shows duplicate headers (`CGST`, `IGST`, `SGST`, `CGST`...) because several fields share one display name.
- **Mapping build loop on a custom configuration** (RAN by running the real file's text with only `EXTRACTION_CONFIG` replaced):
  - Two output sheets mapping the same input sheet and cell -> only the **last** field name survives in `CELL_MAPPINGS` (`Sheet_Two_Same_Cell_Again`), so the first sheet's column could only ever get `Missing`.
  - A mapping with a blank `input_sheet` is left out of `CELL_MAPPINGS` (`:270`) but still gets a column in the report.
  - `report_type: 'itc'` (lower case) creates a separate `'itc'` group that `_extract_data` never asks for (it asks for `'ITC'`).

---

## 5. Cache manager and run history (Part 5)

### 5.1 Files, writers and readers (READ)

| File | Written by | Read by |
|---|---|---|
| `~/.gst_organizer/settings.json` | `CacheManager.save_settings` (`models/cache_manager.py:41-55`), called from `gui/pyqt_main_window.py:213` (`_save_settings`: menu "Save" `:102`, on scan `:227`, after processing `:381`, dark-mode toggle `:597`, window close `:603`) and from `migrate_from_old_cache` (`:119`) | `load_settings` (`:57-62`) at start-up (`pyqt_main_window.py:50`) |
| `~/.gst_organizer/session_state.json` | `save_settings` (recent folders, only when a new folder is added `:48-52`), `save_session_state` (`:76-80`, from `pyqt_main_window.py:604-607`: `active_tab`, `window_size`), migration (`:124`) | `get_recent_source_folders` (`:86-89`) and `load_session_state` (`:82-84`): **neither is called anywhere** (READ: grep, HIGH), so this file is written but never used |
| `~/.gst_organizer/run_history.json` | `save_run_history` -> `RunHistory.save` (`models/run_history.py:128-136`), after processing (`pyqt_main_window.py:379-380`) and after extraction (`:555-558`) | `load_run_history` -> `RunHistory.load` (`:138-148`) at start-up (`pyqt_main_window.py:43`); shown as "last run" in the Validation tab (`gui/views/validation_view.py:384`) |
| `~/.gst_organizer/.migrated` | migration (`:127`) | migration (`:101`) |
| `~/.gst_organizer_cache.json` (old cache from the Tkinter version) | not written by B | migration only (`:98-106`) |

`_DEFAULT_CACHE_DIR` is fixed when the module is imported (`cache_manager.py:17`). RAN: with my fake `HOME` it became `/root/merge-scratch/run/B-TRACE/fakehome/.gst_organizer`. `power_query_extractor/extractor_main.py:36-60` reads yet another set of names (`gst_organizer_cache.json`, `.cache/app_state.json`, `config/last_session.json` in the *current* folder) that nothing in B writes (READ). That file cannot be imported anyway (section 9).

### 5.2 Results (RAN, `scripts/t5_cache.py`; JSON `out/t5_cache.json`)

| # | Action | Result |
|---|---|---|
| 01 | `CacheManager(empty dir)` | folder created, no files written yet |
| 02-04 | load settings / history / session with no files | defaults: `AppSettings()` (processing_mode `fresh`, client_name_max_length 35, pq_wait_time 10, pq_suffix_pattern `_Refreshed_{timestamp}` ...); empty history `{runs: [], client_history: {}, last_run_timestamp: '', total_runs: 0}`; `{}` |
| 05 | save settings 8 times with source folders one..six, three again, `''` | recent list newest-first, capped at 5: after "six" `[six, five, four, three, two]`; **re-using "three" does not move it to the front**; `''` is not added. `settings.json` holds the 11 settings fields |
| 06 | reload | same 11 values |
| 07 | settings.json with an unknown key and wrong types (`client_name_max_length "abc"`, `dark_mode "yes"`, `processing_mode "banana"`, `pq_wait_time -5`) | unknown key dropped; **wrong types and values accepted as-is** (`AppSettings.from_dict`, `models/app_state.py:41-46`, filters names only) |
| 08 | `FileOrganizer(..., client_name_max_length="abc")` | `TypeError: '>' not supported between instances of 'str' and 'int'` (`core/file_organizer.py:42`). In the app, processing uses settings freshly collected from the screen (`pyqt_main_window.py:352`, `:194-208`), where the spin box limits the value to 15-100, so MEDIUM that this cannot happen there |
| 09 | `FileOrganizer(..., "banana")` | `ValueError: Invalid processing mode: banana` |
| 10, 11, 15 | settings.json truncated / empty / UTF-16 | warning `Failed to load settings.json: ...`, **defaults silently used** |
| 12 | settings.json = `null` | defaults, no warning |
| 13, 14 | settings.json = a list `[1, 2]` / a text `"hello"` | **crash**: `AttributeError: 'list' object has no attribute 'items'`. `load_settings` has no protection (`cache_manager.py:57-62`), and it runs in the main window's start-up (`pyqt_main_window.py:50`), where it would end in main.py's catch-all "Critical error" (`main.py:56-59`, READ, MEDIUM) |
| 16 | save over a corrupt settings.json | file simply overwritten with good content (the old content is lost) |
| 17-19 | session_state.json = a list | `save_settings`, `save_session_state` and `get_recent_source_folders` all **crash** (`AttributeError ... 'get'` / `'update'`) |
| 20 | `recent_source_folders` = a text | `save_settings` crashes: `AttributeError: 'str' object has no attribute 'insert'` |
| 21 | all three files deleted | defaults again, no warning |
| 22 | whole cache folder deleted, then save | folder recreated by `_save_json` (`:138`); `settings.json` and `session_state.json` written |
| 23 | cache "folder" path is actually a file | `FileExistsError` in the constructor (`:33`): start-up crash |
| 24 | `add_run(summary, [record])` | one run entry with **its own** timestamp (`run_history.py:46`) and the 6 listed counts; extra summary keys ignored; per-client record stored by `client_key`. `to_dict()` returns the live internal lists, not copies (RAN: `to_dict()['runs'] is h.runs` -> True) |
| 25 | save + reload | identical |
| 26 | `add_extraction_result(result['client'], result)` as the app does (`pyqt_main_window.py:556-557`) | **stored under a new key `'Alpha Test Traders-DL / Version-211026 1000'`** with an otherwise empty record. The processing record under `'Alpha Test Traders-DL'` keeps `pq_refreshed False`, and its "last run" stays `2026-10-21 10:00:00`. The `datetime` in the result is saved as text (`default=str`) and comes back as text |
| 27 | two extraction results on the same key | the second **replaces** the first in the latest record (only `{'second': 2}` left) |
| 28 | 60 runs | `runs` capped at 50 (the oldest 10 dropped), `total_runs` 60, per-client history **not capped** (60 records) |
| 29-31 | run_history.json truncated / a list / null | warning, empty history (protected: `run_history.py:140-147`) |
| 32, 33 | `{"runs": 5}` / `{"client_history": []}` | loaded **without** a warning; the next `add_run` / `add_extraction_result` **crashes** (`AttributeError: 'int' object has no attribute 'append'` / `TypeError: list indices must be integers...`) (34, 35) |
| 36 | corrupt history, then a new run saved | the old history is **silently replaced** by a one-run history |
| 37 | odd values (path, set, bytes) in extraction data | saved as text (`'/a/b'`, `'{1, 2}'`, `"b'x'"`) |
| 38 | migration, first time (old cache with source folder, ITC template, include_client_name True, max length 40, dark mode, recent `[a, b]`) | settings.json gets those values; `.migrated` marker made; session_state recent = `['/old/a', '/old/b']`. **The migrated source folder itself is not in the recent list**: `save_settings` put it first, then `save_session_state` replaced the list with the old one (`cache_manager.py:119-124`) |
| 39 | migration again after the old cache changed | skipped (marker present) |
| 40, 41 | old cache corrupt | warning every time, no marker, so it is **retried at every start** |
| 42 | `CacheManager()` with no folder given | uses `<HOME>/.gst_organizer` |

Why the protection differs (technical): `_load_json` (`cache_manager.py:144-152`) protects only the *reading*; any valid JSON that is not a dict goes back to callers that assume a dict. `RunHistory.load` runs its `from_dict` inside the protection, so a list there is caught, but wrong inner types are not checked. (HIGH)

`ClientRunRecord` fields `itc_report_path`, `sales_report_path` and `errors` are never filled by the processing worker (`gui/controllers/processing_controller.py:150-160`, READ: always `''` / `[]`).

---

## 6. Folder scanning and organising on a fake tree (Part 6)

### 6.1 The fake source folder (`work/t6/source`, made by `scripts/t6_scan_organise.py`)

Three planned clients:
- **complete**: Alpha Test Traders, Delhi; 9 files, including four GSTR-3B files, one of which uses the state written as `DL`;
- **missing files**: Beta Demo Industries Private Limited, Maharashtra; GSTR-3B + Sales + a SalesReco `.xls` that has only the old-Excel header bytes, and an AnnualReport that is really a text file;
- **duplicates + unrecognised**: Gamma Sample LLP, Tamil Nadu; two GSTR-2B-Reco, two Sales, two GSTR-3B whose names differ only by space/underscore, an unrecognised "Purchase Register", and one file with the state written `Tamilnadu`.

Traps: the 12 fixed names, a lower-case client variant, `.xlsm`, upper-case `.XLSX`, a 404-byte "xlsx", an Excel lock file `~$...`, a `.json`, a `.csv`, a file in a sub-folder `old/`, and a broken shortcut (a link pointing to a file that does not exist). Real `.xlsx` files were made with openpyxl (about 4.85 KB each, each with different content).

### 6.2 What the scan did with every file (RAN, threaded `ScanController.scan_files`, `gui/controllers/scan_controller.py:57-80`; parser `core/file_parser.py:126-256`)

| source file (fake) | size (bytes) | found by `find_excel_files`? | outcome of the scan |
|---|---|---|---|
| (1) GSTR3B-Alpha Test Traders-Delhi-Jun.xlsx | 4850 | yes | client `Alpha Test Traders-DL`, type `GSTR-3B Export` |
| ALPHA TEST TRADERS gstr-1 (1).xlsx | 4851 | yes | variation: No matching pattern found; suggestion: Check file naming convention |
| Alpha Test Traders GSTR-2B Apr-24.xlsx | 4851 | yes | variation: No matching pattern found; suggestion: GSTR-2B-Reco-ClientName-State-Period.xlsx |
| Alpha Test Traders GSTR2B APR 2024.xlsx | 4851 | yes | variation: No matching pattern found; suggestion: GSTR-2B-Reco-ClientName-State-Period.xlsx |
| AnnualReport-Alpha Test Traders-Delhi-2024-25.xlsx | 4850 | yes | client `Alpha Test Traders-DL`, type `Annual Report` |
| AnnualReport-Beta Demo Industries Private Limited-Maharashtra-2025.xlsx | 1700 | yes | variation: Invalid or corrupted Excel file |
| AnnualReport-Gamma Sample LLP-Tamil Nadu-2025.xlsx | 4851 | yes | client `Gamma Sample LLP-TN`, type `Annual Report` |
| Beta Demo Industries Private Limited Purchase Register.xlsx | 4851 | yes | variation: No matching pattern found; suggestion: Check file naming convention |
| Beta Demo Industries Pvt Ltd_GSTR3B_2024-25.xlsx | 4850 | yes | variation: No matching pattern found; suggestion: GSTR3B-ClientName-State-Month.xlsx |
| Delta Example & Co ITC Report_refreshed.xlsx | 4851 | yes | variation: No matching pattern found; suggestion: Check file naming convention |
| Delta Example & Co Reco Diff.xlsx | 4851 | yes | variation: No matching pattern found; suggestion: Check file naming convention |
| Epsilon Trial Enterprises Table 8D Limits.xlsx | 4851 | yes | variation: No matching pattern found; suggestion: Check file naming convention |
| Epsilon Trial Enterprises gstr9.json | 14 | no | not looked at |
| GSTR-2B-Reco-Alpha Test Traders-Delhi-Apr24-Mar25.xlsx | 4850 | yes | client `Alpha Test Traders-DL`, type `GSTR-2B Reco` |
| GSTR-2B-Reco-Gamma Sample LLP-Tamil Nadu-H1.xlsx | 4851 | yes | client `Gamma Sample LLP-TN`, type `GSTR-2B Reco` |
| GSTR-2B-Reco-Gamma Sample LLP-Tamil Nadu-H2.xlsx | 4851 | yes | client `Gamma Sample LLP-TN`, type `GSTR-2B Reco` |
| GSTR3B-Alpha Test Traders-DL-Jul.xlsx | 4850 | yes | client `Alpha Test Traders-DL`, type `GSTR-3B Export` |
| GSTR3B-Alpha Test Traders-Delhi-Apr.xlsx | 4850 | yes | client `Alpha Test Traders-DL`, type `GSTR-3B Export` |
| GSTR3B-Alpha Test Traders-Delhi-Aug.xlsx | 404 | yes | variation: File too small to be valid Excel |
| GSTR3B-Alpha Test Traders-Delhi-Dec.xlsx | symlink->missing | yes | dropped silently (stat failed; only a log warning), in neither list |
| GSTR3B-Alpha Test Traders-Delhi-May.xlsx | 4850 | yes | client `Alpha Test Traders-DL`, type `GSTR-3B Export` |
| GSTR3B-Alpha Test Traders-Delhi-Oct.xlsm | 4851 | yes | variation: No matching pattern found; suggestion: GSTR3B-ClientName-State-Month.xlsx |
| GSTR3B-Alpha Test Traders-Delhi-Sep.XLSX | 4852 | no | not looked at |
| GSTR3B-Beta Demo Industries Private Limited-Maharashtra-Apr.xlsx | 4850 | yes | client `Beta Demo Industries Private Limited-MH`, type `GSTR-3B Export` |
| GSTR3B-Gamma Sample LLP-Tamil Nadu-Apr 24.xlsx | 4851 | yes | client `Gamma Sample LLP-TN`, type `GSTR-3B Export` |
| GSTR3B-Gamma Sample LLP-Tamil Nadu-Apr_24.xlsx | 4851 | yes | client `Gamma Sample LLP-TN`, type `GSTR-3B Export` |
| GSTR3B-Gamma Sample LLP-Tamilnadu-May.xlsx | 4851 | yes | client `Gamma Sample LLP-TAM`, type `GSTR-3B Export` |
| GSTR3B-alpha test traders-Delhi-Feb.xlsx | 4851 | yes | client `alpha test traders-DL`, type `GSTR-3B Export` |
| Gamma Sample LLP IMS Reco (2).xlsx | 4851 | yes | variation: No matching pattern found; suggestion: ImsReco-ClientName-State-DDMMYYYY.xlsx |
| Gamma Sample LLP Purchase Register.xlsx | 4851 | yes | variation: No matching pattern found; suggestion: Check file naming convention |
| Gamma Sample LLP Sales Register FY 2024-25.xls | 3008 | yes | variation: No matching pattern found; suggestion: Sales-ClientName-State-StartMonth-EndMonth.xlsx |
| ImsReco-Alpha Test Traders-Delhi-31032025.xlsx | 4850 | yes | client `Alpha Test Traders-DL`, type `IMS Reco` |
| ImsReco-Gamma Sample LLP-Tamil Nadu-31032025.xlsx | 4851 | yes | client `Gamma Sample LLP-TN`, type `IMS Reco` |
| Sales-Alpha Test Traders-Delhi-Apr24-Mar25.xlsx | 4850 | yes | client `Alpha Test Traders-DL`, type `Sales` |
| Sales-Beta Demo Industries Private Limited-Maharashtra-Apr-Mar.xlsx | 4851 | yes | client `Beta Demo Industries Private Limited-MH`, type `Sales` |
| Sales-Gamma Sample LLP-Tamil Nadu-Apr-Sep.xlsx | 4851 | yes | client `Gamma Sample LLP-TN`, type `Sales` |
| Sales-Gamma Sample LLP-Tamil Nadu-Oct-Mar.xlsx | 4850 | yes | client `Gamma Sample LLP-TN`, type `Sales` |
| SalesReco-Alpha Test Traders-Delhi-FY.xlsx | 4849 | yes | client `Alpha Test Traders-DL`, type `Sales Reco` |
| SalesReco-Beta Demo Industries Private Limited-Maharashtra-H1.xls | 3008 | yes | client `Beta Demo Industries Private Limited-MH`, type `Sales Reco` |
| SalesReco-Gamma Sample LLP-Tamil Nadu-FY.xlsx | 4851 | yes | client `Gamma Sample LLP-TN`, type `Sales Reco` |
| Zeta Mock Pvt. Ltd. Extra In GSTR2B.xlsx | 4851 | yes | variation: No matching pattern found; suggestion: GSTR-2B-Reco-ClientName-State-Period.xlsx |
| notes.csv | 8 | no | not looked at |
| old/GSTR3B-Alpha Test Traders-Delhi-Nov.xlsx | 4851 | no | not looked at |
| ~$GSTR3B-Alpha Test Traders-Delhi-Apr.xlsx | 165 | no | not looked at |

Counts (RAN): `find_excel_files` returned 39 paths. The glob is `*.xlsx`, `*.xls`, `*.xlsm` in that one folder only, minus `~$` files, sorted (`utils/helpers.py:249-266`). Of those 39, 36 became "scanned files"; 2 went straight to variations (the too-small file, `<1024` bytes, `core/file_parser.py:184-191`, and the fake text "xlsx", signature check `:194-206`); and 1, the broken link, was **dropped silently**, with only the log line `Cannot stat ...: [Errno 2] No such file or directory` (`:175-179`), not in any list the user sees. The `.XLSX` file was not found because the glob is case-sensitive on Linux; on Windows Python's glob ignores case, so it would be found and recognised (READ/knowledge, HIGH). The sub-folder file, `.json`, `.csv` and `~$` file were not looked at.

Clients found (RAN):

| client key | client | state shown | state_code | files per type | missing_files | extra_files | completeness % | status |
|---|---|---|---|---|---|---|---|---|
| Alpha Test Traders-DL | Alpha Test Traders | Delhi | DL | GSTR-3B Export: 4; Annual Report: 1; GSTR-2B Reco: 1; IMS Reco: 1; Sales: 1; Sales Reco: 1 | - | - | 100.0 | Complete |
| Gamma Sample LLP-TN | Gamma Sample LLP | Tamil Nadu | TN | Annual Report: 1; GSTR-2B Reco: 2; GSTR-3B Export: 2; IMS Reco: 1; Sales: 2; Sales Reco: 1 | - | Multiple GSTR-2B Reco (2 files), Multiple Sales (2 files) | 100.0 | Has duplicates |
| Beta Demo Industries Private Limited-MH | Beta Demo Industries Private Limited | Maharashtra | MH | GSTR-3B Export: 1; Sales: 1; Sales Reco: 1 | Annual Report, GSTR-2B Reco, IMS Reco | - | 50.0 | Missing 3 files |
| Gamma Sample LLP-TAM | Gamma Sample LLP | Tamilnadu | TAM | GSTR-3B Export: 1 | Annual Report, GSTR-2B Reco, IMS Reco, Sales, Sales Reco | - | 16.7 | Missing 5 files |
| alpha test traders-DL | alpha test traders | Delhi | DL | GSTR-3B Export: 1 | Annual Report, GSTR-2B Reco, IMS Reco, Sales, Sales Reco | - | 16.7 | Missing 5 files |

What to notice (RAN):
- **Delhi + DL merged**: `GSTR3B-...-DL-Jul.xlsx` joined `Alpha Test Traders-DL`. The client's `state` is overwritten by whichever file is processed last (`core/file_parser.py:266-268`); here `Delhi`, because the `DL` file sorts before the `Delhi` files.
- **Tamil Nadu + Tamilnadu split** into `Gamma Sample LLP-TN` and `Gamma Sample LLP-TAM` (log `Unknown state 'Tamilnadu', using fallback: TAM`).
- **Upper/lower case split**: `alpha test traders-DL` is a separate client from `Alpha Test Traders-DL` (the key is case-sensitive, `:262`).
- **Duplicates**: Gamma TN shows `Multiple GSTR-2B Reco (2 files)`, `Multiple Sales (2 files)`, completeness 100.0, status "Has duplicates". Its two GSTR-3B files are not reported (the GSTR-3B exception).
- **Fake files counted**: Beta's SalesReco `.xls` (header bytes only) was accepted. The scanner checks only the first 4 bytes (`:197-199`).
- **Statistics** (RAN, `get_statistics` `:336-364`): total_files 36, parsed 23, unparsed 13, total_clients 5, complete 1, incomplete 4, variations 15, total_size 170937 (`166.9 KB`), file_type_distribution `{GSTR-3B Export: 9, Annual Report: 2, GSTR-2B Reco: 3, IMS Reco: 2, Sales: 4, Sales Reco: 3}`, parsing_rate 63.888888888888886 (not rounded), completion_rate 20.0. "total_files" leaves out the two invalid files, but "variations" includes them.
- **Progress messages** (RAN): `(0,0,'Finding Excel files...')`, `(0,39,'Found 39 files, processing...')`, then `(n,39,'[n/39] <file name>')` for each of the 39 (41 messages in all).
- **Second scan while the first is running** (RAN): ignored, with the log line `Scan already in progress` (`scan_controller.py:59-61`); `scan_started` was emitted only once.
- **Errors** (RAN): missing folder -> `FileNotFoundError: Folder not found: ...` (and through the controller, the `scan_error` signal carries `Folder not found: ...`); a file instead of a folder -> `ValueError: Not a directory: ...`; empty folder -> returns empty results with the log line `No Excel files found`.
- **`validate_results`** (never called by the app; RAN directly): `Missing Critical Files` / `Warning` / `Annual Report` for `Beta Demo Industries Private Limited-MH`, `Gamma Sample LLP-TAM` and `alpha test traders-DL`.
- **Originals after scanning: unchanged** (RAN: SHA-256, size and modification time of every source file identical).

### 6.3 Organising: six runs through the real `ProcessingWorker._process` (RAN)

Set-up: the five scanned clients, all selected; templates left empty, so each ITC/Sales report stops at "template path is empty" (`gui/controllers/processing_controller.py:210-213`; 10 report errors per run, and **each client is still counted as successful**, `:135`). Note: the real window refuses to start processing unless both template files exist (`gui/pyqt_main_window.py:352-358`). I called `ProcessingWorker` directly so that organising could be tested without templates. The organising steps are the same ones the window triggers (`create_client_structure` -> `organize_files` -> reports -> `create_organization_report`, `processing_controller.py:99-133`). The minute stamp was fixed by replacing `get_timestamp` in `core/file_organizer.py` and `processing_controller.py` for the test. The backup stamp (`create_backup` -> `get_safe_timestamp`) used the real clock (`0359`).

| Run | Mode, minute | Level-1 folder used | Effect on data files (23 per run) | Other files |
|---|---|---|---|---|
| 1 | fresh, 1000 | new `Annual Statement-211026 1000` | 23 Success; 22 distinct copies + 1 `_backup_` (Gamma TN GSTR-3B name collision, 6.5) | per client `_Organization_Report_<client>_<state>.txt` (raw names, with spaces); `GST_Processing_Summary_211026_1000.xlsx` |
| 2 | fresh, **same** minute 1000 | the same folder (same name) | every existing copy first copied to `<name>_backup_211026_0359.<ext>`, then overwritten. 23 backup operations gave 21 new backup files: the Gamma TN collision file's backup name already existed from run 1 (same minute), so it was overwritten twice | reports rewritten |
| 3 | rerun, 1001 | **existing** `...1000` (newest modification time, `core/file_organizer.py:160-169`) | new `Version-211026 1001` in every client folder, 23 Success | new summary `..._1001.xlsx` |
| 4 | rerun, **same** minute 1001 | same | all 23 copies **overwritten silently**: no backup, no skip (`_organize_single_file` has no branch for rerun, `:260-270`) | reports rewritten |
| 5 | resume, same minute 1001 | same | all 23 `Skipped` / `File already exists` (`:266-270`); summary `total_files 0` | organisation reports **rewritten** with "Successfully Copied: 0" |
| 6 | fresh 1002, first client only, "include client name" on | new `Annual Statement-211026 1002` | category folders named `GSTR-3B Exports (Alpha_Test_Traders)`, `Other ITC related files (Alpha_Test_Traders)`, `Sales related files (Alpha_Test_Traders)` | - |

Folder tree made by run 1 (RAN), per client:

```
<target>/Annual Statement-211026 1000/
    GST_Processing_Summary_211026_1000.xlsx
    Alpha Test Traders-DL/                      <- create_client_state_key (no sanitising)
        Version-211026 1000/
            AnnualReport-Alpha_Test_Traders-Delhi-2024-25.xlsx      <- Annual Report goes in the version folder itself
            _Organization_Report_Alpha Test Traders_Delhi.txt
            GSTR-3B Exports/            (1)_GSTR3B-Alpha_Test_Traders-Delhi-Jun.xlsx, GSTR3B-Alpha_Test_Traders-DL-Jul.xlsx, ...-Apr.xlsx, ...-May.xlsx
            Other ITC related files/    GSTR-2B-Reco-Alpha_Test_Traders-Delhi-Apr24-Mar25.xlsx, ImsReco-Alpha_Test_Traders-Delhi-31032025.xlsx
            Sales related files/        Sales-...-Apr24-Mar25.xlsx, SalesReco-...-FY.xlsx
    Beta Demo Industries Pvt Ltd-MH/Version-211026 1000/ ... (empty "Other ITC related files" folder is still created)
    Gamma Sample LLP-TN/ ..., Gamma Sample LLP-TAM/ ..., alpha test traders-DL/ ...
```

Copied file names are sanitised (spaces -> `_`, `Private Limited` -> `Pvt_Ltd`, `(1) ` -> `(1)_`). Category folders are always created, even when empty (`core/file_organizer.py:113-135`). Partly created folders stay behind when a client fails, because `cleanup_on_error` (`:364-379`) is never called (READ: grep; RAN in 6.7).

### 6.4 Are the ORIGINAL files untouched? (RAN)

**Yes.** After the scan, after all six processing runs and after the collision runs, the SHA-256, size and modification time of every file in `work/t6/source` were identical to the starting snapshot (`source_unchanged_after_all_organising: true`). Files are copied, never moved (`shutil.copy2` in `utils/helpers.py:169`), and copies keep the source's modification time (RAN: true for every non-backup copy; 75 target files have content equal to some source file). The copy check compares **sizes only** (`utils/helpers.py:172-178`), not content (READ).

### 6.5 Two source files whose cleaned names collide (RAN)

`GSTR3B-Gamma Sample LLP-Tamil Nadu-Apr 24.xlsx` (content fingerprint `2c05fccb...`) and `...-Apr_24.xlsx` (`9acaf979...`) both become `GSTR3B-Gamma_Sample_LLP-Tamil_Nadu-Apr_24.xlsx`.
- **fresh** (run 1): the folder holds `...Apr_24.xlsx` = `9acaf979` (the second file) and `...Apr_24_backup_211026_0359.xlsx` = `2c05fccb` (the first file). Both "Success".
- **rerun** (run 3, new version folder): only `...Apr_24.xlsx` = `9acaf979`. **The first file's content is not in the organised output at all**, yet both were reported "Success".
- The organisation report lists `GSTR3B-Gamma_Sample_LLP-Tamil_Nadu-Apr_24.xlsx` twice. `organize_files` overwrites the in-memory scan data with the cleaned names (`core/file_organizer.py:216-229`, RAN: both entries now carry the same name).
- Backup names have minute resolution: a second backup in the same minute **overwrites** the first (RAN, `t6b` 7d: two `create_backup` calls gave one file `a_backup_211026_0401.xlsx`).
- MEDIUM, cannot check here: the report templates get *folder* paths (`utils/constants.py:96-120`: `B2 gstr3b_folder`, `B7 gstr2b_folder` ...), so if their Power Query reads every file in a folder, the `_backup_` copies would be read as well. To check: open the ITC template's Power Query and see whether it filters file names.

### 6.6 Two clients folded into one Level-2 folder (RAN, `work/t6/source2`)

"Alpha Test Traders and Distributors North" and "... South" (Delhi) are two separate clients in the scan (keys `...North-DL`, `...South-DL`). After processing, all four of their files plus both organisation reports sit in **one** folder `Alpha Test Traders and Distribut-DL/Version-211026 1100/`, in shared `GSTR-3B Exports` and `Other ITC related files` folders. The file names differ, so nothing was overwritten here. The same MEDIUM caveat about folder-reading templates applies, now for two clients' data in one folder.

### 6.7 Remaining organiser branches (RAN, `scripts/t6b_organiser_edges.py`)

| Case | Result |
|---|---|
| source file deleted between scan and processing | that file `Failed` / `Copy operation failed`, log `Source file not found: ...`; the others continue |
| unknown file type | `Failed` / `No destination folder for type: Purchase Register` |
| a *file* sitting where the Level-2 folder must go | that client counted failed: `[Errno 17] File exists: ...`; the next client completes |
| target folder path is a file | `ValueError: Cannot create target folder: [Errno 17] File exists: ...` |
| per-client "include name" setting `{'Beta ...-MH': True}` | Beta's category folders get `(Beta_Demo_Industries_Pvt_Ltd)`; Alpha's do not. **In the app this dict is never filled**: `ProcessingState.client_folder_settings` stays `{}` (`models/app_state.py:83`; READ: grep finds no writer), so only the global check box works |
| `cleanup_on_error` (never called) | a version folder holding only `_`-files counts as "empty" and is deleted; one with real files is kept |
| `get_summary` (never called) | works: `folders_created 6, files_processed 2, successful 1, failed 1, skipped 0` |
| `safe_copy_file` missing source -> False; into new sub-folders -> True; `create_backup` of a missing file -> None; `get_unique_path` (never called) -> `u_2.xlsx`; `safe_path_join('', '')` -> `ValueError: No valid path components provided` | as listed |

### 6.8 The per-client organisation report adds up the whole batch (RAN)

`create_organization_report` (`core/file_organizer.py:300-362`) writes "Successfully Copied" and "Folders Created" from the organiser's running lists (`:345-351`). Those lists cover **every client processed so far in that run**, not just this client. RAN, run 1:
- `Gamma Sample LLP-TN` (2nd client): `Total Files: 9`, `Successfully Copied: 18`, `Folders Created: 12`;
- `Gamma Sample LLP-TAM` (4th client): `Total Files: 1`, `Successfully Copied: 22`, `Folders Created: 24`.

"Folders Created" also counts the Level-1 folder again for every client in fresh mode (`:155`).

---

## 7. Remaining helper functions (RAN, `scripts/t7_helpers.py`)

| Function | Inputs -> outputs |
|---|---|
| `format_size` (`utils/helpers.py:272-285`) | -1 `Unknown`; 0 `0.0 B`; 1023 `1023.0 B`; 1024 `1.0 KB`; 1048575 `1024.0 KB`; 1048576 `1.0 MB`; 10^12 `931.3 GB`; 10^19 `9094947.0 TB` |
| `format_duration` (`:405-416`) | 0.9 `0s`; 59.99 `59s`; 60 `1m 0s`; 3599 `59m 59s`; 3600 `1h 0m`; 86400 `24h 0m`; -5 `-5s` |
| `truncate_path` (`:287-309`, only exported, never called) | when shortening, keeps only the last part: `.../file_name_that_is_long.xlsx`; the first-part branch never includes the middle |
| `clean_windows_path` (`:311-315`) | `''`/`None` -> `''`; `/a/b/c.xlsx` -> `\a\b\c.xlsx` |
| `extract_filename_without_extension` (`:446-448`) | `a.b.c.xlsx` -> `a.b.c`; `Zeta Mock Pvt. Ltd` -> `Zeta Mock Pvt` (same dot effect as 2.3) |
| `get_relative_excel_path` (`:450-467`, never called) | inside -> `GSTR-3B Exports`; outside -> `\other\x.xlsx` |
| `get_short_path` (`:473-483`) | on Linux returns the path unchanged |
| `calculate_file_hash` (`:422-440`, never called) | md5 text; missing file -> None; unknown algorithm -> None |
| `validate_excel_file` (`:210-247`, used by excel_handler only) | accepts `.xlsx .xls .xlsm .xltx .xltm` in any case, ≥1 KB and correct first 4 bytes; `.csv` -> False |
| `get_file_info` (`:115-136`) | dict with name, size, formatted size ...; missing -> `{'error': 'File not found'}` |
| `find_excel_files` on a missing folder or on a file | `[]` (no error) |
| `ProgressTracker` (`:369-403`) | callback gets `(current, total, percent, message)`; `get_eta` before any update `Unknown`; **`ZeroDivisionError`** if called after an update with zero elapsed time (`:396`). In the app it is only called from the second client on (`processing_controller.py:87`), so this is unlikely (LOW) |

Constants defined but never used (READ: grep outside `utils/constants.py` and `utils/__init__.py`): `EXCEL_EXTENSIONS`, `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB`, `EXCEL_SIGNATURES` (the parser re-types the same signatures at `core/file_parser.py:197-198`), `GUI_CONFIG` (still says "v3.0" and Tk-style sizes), `MESSAGES`. Used: `FILE_PATTERNS`, `EXPECTED_FILE_TYPES`, `FOLDER_STRUCTURE`, `PROCESSING_MODES`, `STATE_CODE_MAPPING`, `EXCEL_TEMPLATE_MAPPING` and `EXCEL_SAFETY` (the last two by `core/excel_handler.py`).

---

## 8. The most important findings, ranked (technical; GST rules described, not judged)

1. **[GST RULE context] Three output sheets of the PQ extraction report carry the wrong cells (RAN, 4.6).** `IMS Reco`, `Extra In Purchase` and `Extra In GSTR2B` show the `Diff GSTR-3B` J15/K15/L15 values whenever that sheet exists and those cells are filled. Cause: the substring look-up at `power_query_extractor/core/data_consolidator.py:130-133`. Confirmed twice, including with reversed key order.
2. **The app cannot start without pywin32 (RAN, 9.1).** `report_processor.py:11-15` imports pywin32 unguarded and is pulled in by `gui/__init__.py`, so even scanning needs it. `requirements.txt` has pywin32 commented out, and `main.py`'s message suggests installing PyQt5/openpyxl instead.
3. **Re-running extraction can treat its own `_Refreshed_` copy as the source report (RAN on Linux, 4.7/4.8).** The glob `ITC_Report_*.xlsx` plus `[0]` (`report_processor.py:76-89`). On Windows/NTFS the default suffix usually avoids it (MEDIUM). Skip-refresh and the tab's status look only for `_Refreshed_` even when the user changed the suffix.
4. **Silent overwrites caused by minute-resolution names (RAN):**
   - same-minute fresh runs make `_backup_` copies of every file, and a later same-minute backup overwrites an earlier one;
   - same-minute rerun overwrites with no backup;
   - same-minute extraction overwrites the refreshed copy and the `PQ_Extraction_Report`;
   - two source names that clean to the same name keep only one copy in rerun mode, while both report "Success" (6.5).
5. **One client becomes two, or two clients become one (RAN):**
   - state-name variants: `Tamil Nadu`/`Tamilnadu` split, `Delhi`/`DL` merge;
   - client-name case: `Alpha ...`/`alpha ...` split;
   - the 35-character Level-2 truncation folds two clients into one folder (6.6);
   - hyphenated client names are cut at the first hyphen (1.3 #2).
6. **"Successfully Run" ✓ means only "refresh validated" (RAN, 4.5).** ✓ appears with all values `Missing` (missing sheet). ✗ hides values that were extracted (validation failed). "Not ticked" looks the same as "failed". The validation accepts any non-empty BB2, including a COM error number (MEDIUM) and another client's name.
7. **Extraction results are stored under a different key from processing records (RAN, 5.2 #26).** The Validation tab's "last run" never reflects an extraction.
8. **Settings files of the wrong shape crash start-up or saving (RAN, 5.2 #13/14/17-20/23).** Corrupt history is silently replaced (#36). Wrong types are accepted (#07).
9. **The scanner trusts the first 4 bytes and drops broken links silently (RAN, 6.2).** Upper-case extensions behave differently on Linux and Windows (HIGH by knowledge).
10. **Per-client organisation reports add up the whole batch (RAN, 6.8).** The per-client "include name" setting has no writer in the app (READ).
11. **`sanitize_filename` leaves everything after a dot in a client name uncleaned, and its Private/Limited rule differs from the Level-2 rule (RAN, 2.3).**

---

## 9. Side observations made while running (outside the six parts; brief)

1. **Import chain (RAN).** `from gui.controllers.scan_controller import ScanController` failed with `ModuleNotFoundError: No module named 'pythoncom'`. The chain: `gui/__init__.py:7` -> `gui/pyqt_main_window.py:25` -> `gui/controllers/__init__.py:4` -> `gui/controllers/extraction_controller.py:10` -> `power_query_extractor/core/report_processor.py:11`. `core/excel_handler.py:19-28` guards the same import with a Windows check and `try/except`; `report_processor.py:11-15` does not. That is an inconsistent pattern (HIGH). Running `python -B proj/main.py` without stand-ins (RAN) printed `Missing module: No module named 'pythoncom'` / `Please install required packages:` / `pip install PyQt5 openpyxl`, exit code 1. It also created `gst_organizer.log` in the *current* folder (`main.py:20`).
2. **`power_query_extractor/extractor_main.py` cannot be imported (RAN).** `ModuleNotFoundError: No module named 'power_query_extractor.gui.extractor_window'`; `power_query_extractor/gui/` holds only `__init__.py`. Nothing imports `extractor_main.py` (READ: grep). `launch_extractor.py` opens the main window on tab 4 instead.
3. **Wrong log message (READ).** `_extract_data` logs "File opened in read-only mode" (`report_processor.py:492-494`), but `_safe_open_workbook` opens without `ReadOnly=True` (`:685`); it closes without saving (`:573`). (HIGH)

---

## Possible client names

None found. I checked the assigned files, the other files listed in 0.2, B's README and the subjects and bodies of B's 17 commit messages (`git log B/main`; grep for Pvt/Ltd/LLP/Traders/Enterprises/Industries/& Co). The only company-like name is the documentation example `ABC Ltd` at `README.md:128-133`. I judge it an obvious generic placeholder and list it here only so the author can confirm. All names in these notes are invented test names. I did not open `gst_organizer.log` from commits edac6a3/6237597; the only `gst_organizer.log` involved is the 83-byte file my own `main.py` run created in `work/mainrun/`. I did not open it; its size matches the single error line printed on screen.

## Questions for the author

1. The configuration says the `IMS Reco`, `Extra In Purchase` and `Extra In GSTR2B` report sheets take `IMS_Reco!AA1/AD1/AG1`, `ExtraInPurchase!O1/P1/Q1` and `ExtraInGSTR2B!S1/T1/U1`. B's report shows the `Diff GSTR-3B` J15/K15/L15 figures there (4.6). Which figures did you expect to see in those three sheets?
2. Should "Successfully Run ✓" mean "the refresh worked" or "the figures were found"? Should a report that was *not ticked* look different from one that *failed*?
3. Is `Info!BB2` meant to hold the client's name? Should the tool compare it with the client being processed?
4. In your real files, is the state written as a full name (`Delhi`, `Tamil Nadu`) or sometimes as a code (`DL`) or joined up (`Tamilnadu`)? B treats `Tamil Nadu` and `Tamilnadu` as two clients. [GST RULE] Do you want two-letter codes, two-digit codes or something else in folder names? (I am not judging the current table, only asking what you want.)
5. Do client names in your files ever contain hyphens (`-`) or dots (`Pvt. Ltd.`)? Hyphens break B's recognition; dots escape name cleaning.
6. Do your ITC/Sales templates' Power Queries read *every* file in the GSTR-3B / ITC / Sales folders? If so, `_backup_` copies, and a second client folded into the same folder, would be read too (6.5, 6.6). Could you provide the templates (they are not in the repository)?
7. Do the templates contain any financial-year or period logic? B's Python code has none (section 3).
8. Should two GSTR-2B-Reco or two Sales files for one client (for example two half-years) be normal, or an error? Should several GSTR-3B files from *different* years be allowed? [GST RULE] (described in section 3; your decision)
9. Should a re-run in the same minute be prevented, or should it create a new, distinct folder?
10. Should `pywin32` be a required install (it is commented out in `requirements.txt`, yet B cannot start without it)?

## INCOMPLETE

All six assigned parts were run. What could **not** be run on this machine, and so rests on reading or knowledge:
- Real Excel COM: the actual Power Query refresh, real cell value types (decimals, error numbers, dates) and recalculation of never-calculated formulas. I used openpyxl stand-ins (`stubs/`). Section 4 marks these MEDIUM.
- Windows-only behaviour: case-insensitive glob (`.XLSX`), NTFS folder order (which report `[0]` picks), trailing dots/spaces in folder names, long paths (>218 / 260 characters, `report_processor.py:668-680`), `win32api.GetShortPathName`, `os.startfile`.
- Permission-denied cases (read-only cache folder, locked source file): this machine runs as the administrator account, which ignores file permissions, so I did not test them.
- `core/excel_handler.py` report creation from templates (no templates in the repository; another task covers excel_handler). The processing summary workbook it creates was produced in my runs, but I did not inspect its contents.
- The GUI views themselves (Setup, Validation, Processing, Extract screens) were not driven; only the controllers and models were run.
