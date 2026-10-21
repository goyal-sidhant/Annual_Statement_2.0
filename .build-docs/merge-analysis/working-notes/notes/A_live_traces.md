# A-TRACE - Project A (AnnualStatement_oldv1, snapshot bf2f1dd) - running the pure-Python parts on made-up inputs

Working notes for the later audit writers. Author of the code: Adv. Sidhant Goyal ("the author"). Project A = `goyal-sidhant/AnnualStatement_oldv1`, HEAD `bf2f1dd56f088d019bed6022428aab9bf0eb423d` (27-07-2026), Tkinter GUI. Everything below is about project A only.

## 0. How to read these notes

**Markers**

- **READ** = I learned it by reading the code. **RAN** = I executed the code and the output shown is what it printed.
- **HIGH / MEDIUM / LOW** = my confidence in an inference (anything that goes beyond what was literally read or printed).
- **[GST RULE]** = code that decides a tax-related matter (which return, table, sheet or cell a figure belongs to; period/financial year; state code logic; which return files a client needs). For these I only describe what the code does. I do not judge correctness. **TIME-VARYING** = a value that can change over time (table numbers, state lists, return periodicity, report layouts).
- Citations are repository-relative `file:line` at snapshot bf2f1dd (the scratch copy has identical line numbers). Commits are cited where they matter.

**Plain-English glossary (first use of each term)**

- *function / method*: a named block of code that does one job; a *method* belongs to a *class* (a bundle of related functions and data).
- *branch*: one of the paths through an `if` / `else` decision.
- *regex / pattern*: a text-matching rule, e.g. "starts with `GSTR3B-`, then some text without a hyphen ...". `[^-]+` means "one or more characters that are not a hyphen". `re.IGNORECASE` means upper/lower case does not matter.
- *parse*: split a file name into its parts (client, state, period).
- *dictionary / key*: a lookup table; the *key* is the label used to find an entry (here a client key looks like `Alpha Test Traders-MH`).
- *sanitize*: clean a name so it is safe as a file or folder name.
- *openpyxl*: a Python library that reads and writes .xlsx files without Excel.
- *COM / pywin32*: the Windows mechanism (and the Python package for it) that lets a program drive the real Excel. Not available on this Linux machine.
- *stub / fake / shim*: a stand-in object I wrote so a piece of code can run without the real thing (no Excel, no window).
- *Tk / Treeview*: Tk is the window toolkit A uses; a *Treeview* is the table widget on the Validation screen.
- *glob*: a wildcard file search such as `ITC_Report_*.xlsx`; *scandir*: listing the entries of one folder.
- *mtime*: a file's or folder's "last modified" time stamp.
- *checksum (sha256)*: a fingerprint of a file's bytes; if it is unchanged, the file's content is unchanged.
- *substring*: a piece of text found inside a longer text (`"IGST"` is a substring of `"GSTR-3B_Difference_Diff_IGST"`).
- *thread*: a second line of work running at the same time as the window.

## 1. Set-up, safety and what was run

- Copy: `cp -a /root/merge-scratch/A-copy /root/merge-scratch/run/A-TRACE/proj`, `__pycache__` folders removed. Verified identical to `git archive A/main` of bf2f1dd (RAN: `diff -r` clean).
- All runs: `/root/merge-scratch/venv-A/bin/python -B` (Python 3.14.6, openpyxl 3.1.5) with `PYTHONDONTWRITEBYTECODE=1`; scripts that create Excel reports ran with `TMPDIR` and `HOME` pointed into the run folder, because `ExcelHandler.__init__` makes a folder in the system temp directory (`core/excel_handler.py:53-54`) and the extractor's cache lives in the home folder (`power_query_extractor/gui/extractor_window.py:118`). The Treeview test ran under `xvfb-run -a`. The project tests ran with `pytest -p no:cacheprovider`.
- `/root/merge-scratch/A-copy` checksums of all 129 files before and after all runs: identical (RAN; files `Acopy_checksums_before.txt` / `Acopy_checksums_after.txt` in the run folder). No `__pycache__` was created in the run copy either.
- One disclosure: I ran `git -C /home/user/AnnualStatement_oldv1 status --short` once (a read-only command, empty output). Git can refresh its own index time-stamp cache during `status`; nothing else touched the original checkouts.
- No Windows-only part was executed. Where the extractor needs Excel COM I used a fake COM layer backed by openpyxl (`scripts/fakecom.py`), clearly marked; it can only show *cached* cell values, not what real Excel would calculate.

**Scripts (all in `/root/merge-scratch/run/A-TRACE/scripts/`) and their outputs (in `/root/merge-scratch/run/A-TRACE/outputs/`)**

| script | what it runs | output |
|---|---|---|
| `common.py` | puts the run copy on the import path; captures log messages | - |
| `fakecom.py` | openpyxl-backed stand-in for `win32com`/`pythoncom` (only for `_extract_data`) | - |
| `t1_filenames.py` | `FileParser.parse_filename` on 60 made-up names | `t1_filenames.md` |
| `t2_names.py` | `sanitize_filename`, `create_client_state_key`, `get_state_code`, `validate_client_name` | `t2_names.md` |
| `t3_periods_time.py` | period text in names, which duplicate feeds the template, run-timestamp names, mtime choice | `t3_periods_time.md` |
| `t4_cells.py` | every cell mapping; `_extract_data` through the fake COM on 4 made-up workbooks; `DataConsolidator.create_report`; `_calculate_value` | `t4_cells.md` |
| `t4b_config_variants.py` | consolidator branches the shipped config never reaches (description, calculation, same cell mapped twice) | `t4b_config_variants.md` |
| `t5_refresh_naming.py` | `refresh_markers`, `is_refreshed_name`, `is_refreshed_copy_of`, `_process_report` naming on fake folders, extractor list helpers | `t5_refresh_naming.md` |
| `t6_scan_organise.py` | fake source tree (3 clients + extras) -> `scan_folder` -> the real orchestrator `ProcessingHandler.process_files_thread` with a stub app, 9 runs in different modes, checksums, Step-4 scan | `t6_scan_organise.md` |
| `t7_setup_and_handlers.py` | `setup_validation`, `FileHandler` scan/validate, `ClientHandler` with a real Tk Treeview | `t7_setup_and_handlers.md` |
| `t8_edges.py` | inconsistent client spellings, a selected key missing from the data, organiser failure branches, backups, helper edge cases | `t8_edges.md` |
| `t9_extract_on_collision.py` | Step 4 on the merged folder from t6 section 7 | `t9_extract_on_collision.md` |
| `t10_failed_refresh_leftover.py` | what a failed refresh leaves behind and how it is then treated | `t10_failed_refresh_leftover.md` |
| `t11_blank_client_count.py` | how a parsed-but-blank-client file is counted | `t11_blank_client_count.md` |
| `t12_treeview_odd_text.py` | Treeview round-trip of text Tcl treats specially (under `xvfb-run`) | `t12_treeview_odd_text.md` |
| `assemble_notes.py` | builds this notes file from `notes_parts/` and `outputs/` | `/root/merge-scratch/notes/A-TRACE.md` |

Work folders (fake inputs and everything the code produced) are under `/root/merge-scratch/run/A-TRACE/work/t3 ... t11` (t1, t2 and t12 need no files). Run order matters only for t9 (it reads t6's output).

## 2. What I read before running (every line of these)

`core/file_parser.py` (1-370), `core/file_organizer.py` (1-425), `core/__init__.py`, `utils/helpers.py` (1-643), `utils/constants.py` (1-291), `utils/__init__.py`, `utils/excel_com.py`, `power_query_extractor/config/cell_mappings.py` (1-277), `power_query_extractor/core/refresh_naming.py` (1-66), `power_query_extractor/core/data_consolidator.py` (1-431), `power_query_extractor/core/report_processor.py` (1-803, needed because the consolidator only consumes what `_extract_data` produces), `gui/handlers/__init__.py`, `gui/handlers/client_handler.py` (1-256), `gui/handlers/file_handler.py` (1-150), `gui/handlers/processing_handler.py` (1-529), `gui/utils/setup_validation.py` (1-115), and the scanning parts of `power_query_extractor/gui/extractor_window.py` (1-140, 255-325, 480-495, 580-980), `core/excel_handler.py` (1-110, 260-325, 600-713: imports, the openpyxl fallback and `prepare_template_data`), `gui/main_window.py` (1-135, 280-330), `gui/tabs/setup_tab.py` (430-451, 609-615), `gui/tabs/validation_tab.py` (197-243), and the tests `tests/conftest.py`, `tests/test_file_patterns.py`, `tests/test_scan.py`, `tests/test_helpers.py`, `tests/test_refresh_naming.py`. Standards 12 (rebuild_pipeline_FINAL.md:817-898) and 13 (900-1061) and `Project_Audit_FINAL.md` were read once.

Key facts from reading that shape every run (READ):

- **The scan only looks at the top level of the source folder.** `scan_excel_files` lists one folder with `os.scandir` and never descends into sub-folders (`utils/helpers.py:328-339`). Candidates must end in `.xlsx`, `.xls` or `.xlsm` (`utils/helpers.py:327,330`), be a regular file, be at least 1024 bytes (`utils/helpers.py:303`), and start with a ZIP or OLE signature (`utils/helpers.py:264-265,284`). The result is sorted by full path (`utils/helpers.py:361`).
- **Patterns** (`utils/constants.py:13-62`), tried in this order, first match wins (`core/file_parser.py:82-128`). Every pattern ends `\.xlsx?$`, i.e. `.xls` or `.xlsx` only, any case.

| pattern name | regex (constants.py line) | type / category / folder | groups kept |
|---|---|---|---|
| GSTR-2B-Reco | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` (15) | GSTR-2B Reco / ITC / itc | client, state, period |
| ImsReco | `^ImsReco-([^-]+)-([^-]+)-(\d{8})(?:\s*\(\d+\))?\.xlsx?$` (23) | IMS Reco / ITC / itc | client, state, date |
| GSTR3B | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` (31) | GSTR-3B Export / GSTR3B / gstr3b | client, state, month |
| Sales | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` (39) | Sales / Sales / sales | client, state, start_month, end_month |
| SalesReco | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$` (47) | Sales Reco / Sales / sales | client, state, period |
| AnnualReport | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$` (55) | Annual Report / Annual / root | client, state, year |

- [GST RULE] `EXPECTED_FILE_TYPES` (`utils/constants.py:68-75`) = the six types above; a client is "Complete" only with at least one of each and no duplicates except GSTR-3B (`core/file_parser.py:252-286`; GSTR-3B exemption at 266-268 with the comment "monthly filings - unlimited"). TIME-VARYING (which returns exist and how often they are filed can change).
- Client grouping key = `"<client text>-<state code>"` (`core/file_parser.py:229-230`), case-sensitive and space-sensitive; the state code comes from `get_state_code` (`utils/helpers.py:592-618`) which maps 39 lower-case names (`utils/constants.py:251-291`) and otherwise invents a code (first letters of the first two words, or the first three characters). [GST RULE] state-code logic; TIME-VARYING list.
- Organising copies, never moves: `shutil.copy2` (`utils/helpers.py:199`).

## 3. T1 - File-name recognition (RAN)

Function under test: `FileParser.parse_filename` (`core/file_parser.py:56-134`). For each name the script also reports: whether `scan_excel_files` would even list the file (extension test at `utils/helpers.py:327-331`), how many of the six patterns match (to look for two patterns matching one name), the client key the scan would group it under (`core/file_parser.py:229-230`), and `_suggest_pattern` (`core/file_parser.py:288-306`) for names that do not parse. Script `scripts/t1_filenames.py`, output `outputs/t1_filenames.md`, reproduced exactly below. All client names are made up.



<!-- RAN output: outputs/t1_filenames.md (headings demoted by two levels) -->

| # | input file name | why | scan_excel_files lists it? | patterns matching | parsed | type / category / folder | client | state | metadata | warnings / error | client key | suggestion (if unparsed) |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | `'GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Apr24-Mar25.xlsx'` | 2B happy; period contains a hyphen | yes | GSTR-2B-Reco | True | GSTR-2B Reco / ITC / itc | `'Alpha Test Traders'` | `'Maharashtra'` | `{'period': 'Apr24-Mar25'}` |  | `Alpha Test Traders-MH` |  |
| 2 | `'gstr-2b-reco-alpha test traders-maharashtra-FY2425.XLSX'` | 2B lower-case prefix, upper-case extension | yes | GSTR-2B-Reco | True | GSTR-2B Reco / ITC / itc | `'alpha test traders'` | `'maharashtra'` | `{'period': 'FY2425'}` |  | `alpha test traders-MH` |  |
| 3 | `'GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Apr24 (1).xlsx'` | 2B with (1) duplicate suffix | yes | GSTR-2B-Reco | True | GSTR-2B Reco / ITC / itc | `'Alpha Test Traders'` | `'Maharashtra'` | `{'period': 'Apr24 (1)'}` |  | `Alpha Test Traders-MH` |  |
| 4 | `'GSTR-2B-Reco-Beta Demo Industries Pvt Ltd-Gujarat-Q1.xls'` | 2B .xls | yes | GSTR-2B-Reco | True | GSTR-2B Reco / ITC / itc | `'Beta Demo Industries Pvt Ltd'` | `'Gujarat'` | `{'period': 'Q1'}` |  | `Beta Demo Industries Pvt Ltd-GJ` |  |
| 5 | `'GSTR-2B-Reco- Alpha Test Traders - Maharashtra -Q1.xlsx'` | 2B extra spaces around parts | yes | GSTR-2B-Reco | True | GSTR-2B Reco / ITC / itc | `'Alpha Test Traders'` | `'Maharashtra'` | `{'period': 'Q1'}` |  | `Alpha Test Traders-MH` |  |
| 6 | `'GSTR-2B-Reco-Alpha-Test-Traders-Maharashtra-Q1.xlsx'` | 2B hyphenated client name | yes | GSTR-2B-Reco | True | GSTR-2B Reco / ITC / itc | `'Alpha'` | `'Test'` | `{'period': 'Traders-Maharashtra-Q1'}` |  | `Alpha-TES` |  |
| 7 | `'GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Q1.xlsm'` | 2B .xlsm extension | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| 8 | `'GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Q1.xlsx.xlsx'` | 2B double extension | yes | GSTR-2B-Reco | True | GSTR-2B Reco / ITC / itc | `'Alpha Test Traders'` | `'Maharashtra'` | `{'period': 'Q1.xlsx'}` |  | `Alpha Test Traders-MH` |  |
| 9 | `'GSTR-2B-Reco-Alpha Test Traders-Maharashtra-.xlsx'` | 2B empty period | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| 10 | `'GSTR2B-Reco-Alpha Test Traders-Maharashtra-Q1.xlsx'` | near-miss: GSTR2B without hyphen | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | GSTR-2B-Reco-ClientName-State-Period.xlsx |
| 11 | `'ImsReco-Alpha Test Traders-Maharashtra-31032025.xlsx'` | IMS happy | yes | ImsReco | True | IMS Reco / ITC / itc | `'Alpha Test Traders'` | `'Maharashtra'` | `{'date': '31032025'}` |  | `Alpha Test Traders-MH` |  |
| 12 | `'ImsReco-Alpha Test Traders-Maharashtra-31032025 (1).xlsx'` | IMS with ' (1)' | yes | ImsReco | True | IMS Reco / ITC / itc | `'Alpha Test Traders'` | `'Maharashtra'` | `{'date': '31032025'}` |  | `Alpha Test Traders-MH` |  |
| 13 | `'ImsReco-Alpha Test Traders-Maharashtra-31032025(2).xlsx'` | IMS with '(2)' no space | yes | ImsReco | True | IMS Reco / ITC / itc | `'Alpha Test Traders'` | `'Maharashtra'` | `{'date': '31032025'}` |  | `Alpha Test Traders-MH` |  |
| 14 | `'IMSRECO-ALPHA TEST TRADERS-MAHARASHTRA-01042024.XLSX'` | IMS all upper-case | yes | ImsReco | True | IMS Reco / ITC / itc | `'ALPHA TEST TRADERS'` | `'MAHARASHTRA'` | `{'date': '01042024'}` |  | `ALPHA TEST TRADERS-MH` |  |
| 15 | `'ImsReco-Alpha Test Traders-Maharashtra-310325.xlsx'` | IMS 6-digit date | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | ImsReco-ClientName-State-DDMMYYYY.xlsx |
| 16 | `'ImsReco-Alpha Test Traders-Maharashtra-Mar25.xlsx'` | IMS month text instead of date | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | ImsReco-ClientName-State-DDMMYYYY.xlsx |
| 17 | `'ImsReco-Alpha Test Traders-Maharashtra-99999999.xlsx'` | IMS impossible date | yes | ImsReco | True | IMS Reco / ITC / itc | `'Alpha Test Traders'` | `'Maharashtra'` | `{'date': '99999999'}` |  | `Alpha Test Traders-MH` |  |
| 18 | `'IMS Reco-Alpha Test Traders-Maharashtra-31032025.xlsx'` | near-miss: 'IMS Reco' with space | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | ImsReco-ClientName-State-DDMMYYYY.xlsx |
| 19 | `'GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx'` | 3B happy | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'Maharashtra'` | `{'month': 'Apr'}` |  | `Alpha Test Traders-MH` |  |
| 20 | `'(1) GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx'` | 3B leading '(1) ' | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'Maharashtra'` | `{'month': 'Apr'}` |  | `Alpha Test Traders-MH` |  |
| 21 | `'(12)GSTR3B-Alpha Test Traders-Maharashtra-May.xlsx'` | 3B leading '(12)' no space | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'Maharashtra'` | `{'month': 'May'}` |  | `Alpha Test Traders-MH` |  |
| 22 | `'GSTR3B-Alpha Test Traders-Maharashtra-Apr (1).xlsx'` | 3B trailing ' (1)' | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'Maharashtra'` | `{'month': 'Apr (1)'}` |  | `Alpha Test Traders-MH` |  |
| 23 | `'GSTR3B-Alpha Test Traders-Maharashtra-Apr-2024.xlsx'` | 3B month with year after hyphen | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | GSTR3B-ClientName-State-Month.xlsx |
| 24 | `'GSTR3B-Alpha Test Traders-Maharashtra-Apr .xlsx'` | 3B space before extension | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'Maharashtra'` | `{'month': 'Apr'}` |  | `Alpha Test Traders-MH` |  |
| 25 | `'GSTR-3B-Alpha Test Traders-Maharashtra-Apr.xlsx'` | near-miss: GSTR-3B with hyphen | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | GSTR3B-ClientName-State-Month.xlsx |
| 26 | `'GSTR3B_Alpha Test Traders_Maharashtra_Apr.xlsx'` | near-miss: underscores | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | GSTR3B-ClientName-State-Month.xlsx |
| 27 | `'Copy of GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx'` | near-miss: 'Copy of' prefix | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | GSTR3B-ClientName-State-Month.xlsx |
| 28 | `' GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx'` | near-miss: leading space | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | GSTR3B-ClientName-State-Month.xlsx |
| 29 | `'GSTR3B- -Maharashtra-Apr.xlsx'` | 3B client is only a space | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `''` | `'Maharashtra'` | `{'month': 'Apr'}` |  | `` |  |
| 30 | `'GSTR3B-Alpha Test Traders-A-Apr.xlsx'` | 3B one-letter state | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'A'` | `{'month': 'Apr'}` | State name issue: Client name too short | `Alpha Test Traders-A` |  |
| 31 | `'GSTR3B-Alpha:Test-Maharashtra-Apr.xlsx'` | 3B client with ':' (string only) | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha:Test'` | `'Maharashtra'` | `{'month': 'Apr'}` | Client name issue: Client name contains invalid characters | `Alpha:Test-MH` |  |
| 32 | `'Sales-Alpha Test Traders-Maharashtra-Apr24-Mar25.xlsx'` | Sales happy | yes | Sales | True | Sales / Sales / sales | `'Alpha Test Traders'` | `'Maharashtra'` | `{'start_month': 'Apr24', 'end_month': 'Mar25'}` |  | `Alpha Test Traders-MH` |  |
| 33 | `'sales-beta demo industries pvt ltd-gujarat-apr-mar.XLS'` | Sales lower-case, .XLS | yes | Sales | True | Sales / Sales / sales | `'beta demo industries pvt ltd'` | `'gujarat'` | `{'start_month': 'apr', 'end_month': 'mar'}` |  | `beta demo industries pvt ltd-GJ` |  |
| 34 | `'Sales-Alpha Test Traders-Maharashtra-Apr-Mar (1).xlsx'` | Sales with ' (1)' | yes | Sales | True | Sales / Sales / sales | `'Alpha Test Traders'` | `'Maharashtra'` | `{'start_month': 'Apr', 'end_month': 'Mar (1)'}` |  | `Alpha Test Traders-MH` |  |
| 35 | `'Sales-Alpha Test Traders-Maharashtra-Apr2024.xlsx'` | Sales with only one period part | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | Sales-ClientName-State-StartMonth-EndMonth.xlsx |
| 36 | `'Sales-Alpha Test Traders-Maharashtra-Apr-Jun-Extra.xlsx'` | Sales with one part too many | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | Sales-ClientName-State-StartMonth-EndMonth.xlsx |
| 37 | `'Sales-Reco-Alpha Test Traders-Maharashtra-Q1.xlsx'` | 'Sales-Reco-' (hyphen) instead of SalesReco | yes | Sales | True | Sales / Sales / sales | `'Reco'` | `'Alpha Test Traders'` | `{'start_month': 'Maharashtra', 'end_month': 'Q1'}` |  | `Reco-AT` |  |
| 38 | `'SalesReco-Alpha Test Traders-Maharashtra-FY 2024-25.xlsx'` | SalesReco happy, period with hyphen | yes | SalesReco | True | Sales Reco / Sales / sales | `'Alpha Test Traders'` | `'Maharashtra'` | `{'period': 'FY 2024-25'}` |  | `Alpha Test Traders-MH` |  |
| 39 | `'Sales Reco-Alpha Test Traders-Maharashtra-Q1.xlsx'` | near-miss: 'Sales Reco' with space | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | SalesReco-ClientName-State-Period.xlsx |
| 40 | `'AnnualReport-Alpha Test Traders-Maharashtra-2024-25.xlsx'` | Annual happy | yes | AnnualReport | True | Annual Report / Annual / root | `'Alpha Test Traders'` | `'Maharashtra'` | `{'year': '2024-25'}` |  | `Alpha Test Traders-MH` |  |
| 41 | `'AnnualReport-Alpha Test Traders-Maharashtra-2024-25 (1).xlsx'` | Annual with ' (1)' | yes | AnnualReport | True | Annual Report / Annual / root | `'Alpha Test Traders'` | `'Maharashtra'` | `{'year': '2024-25 (1)'}` |  | `Alpha Test Traders-MH` |  |
| 42 | `'AnnualReport-Alpha Test Traders-Maharashtra.xlsx'` | Annual with no year part | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | AnnualReport-ClientName-State-Year.xlsx |
| 43 | `'Annual Report-Alpha Test Traders-Maharashtra-2024.xlsx'` | near-miss: 'Annual Report' with space | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | AnnualReport-ClientName-State-Year.xlsx |
| 44 | `'AnnualReport-Alpha Test Traders-Maharashtra-2024_Refreshed_211026_1100.xlsx'` | Annual with refreshed suffix | yes | AnnualReport | True | Annual Report / Annual / root | `'Alpha Test Traders'` | `'Maharashtra'` | `{'year': '2024_Refreshed_211026_1100'}` |  | `Alpha Test Traders-MH` |  |
| 45 | `'ITC_Report_Alpha_Test_Traders_MH_211026_1030_Refreshed_211026_1100.xlsx'` | app's own refreshed ITC report | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | Check file naming convention |
| 46 | `'Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx'` | app's own Sales report | yes | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | Sales-ClientName-State-StartMonth-EndMonth.xlsx |
| 47 | `'GSTR3B-Alpha Test Traders-Maharashtra-Apr.csv'` | non-Excel .csv | NO | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | GSTR3B-ClientName-State-Month.xlsx |
| 48 | `'GSTR3B-Alpha Test Traders-Maharashtra-Apr.pdf'` | non-Excel .pdf | NO | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | GSTR3B-ClientName-State-Month.xlsx |
| 49 | `'GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsb'` | Excel binary .xlsb | NO | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | GSTR3B-ClientName-State-Month.xlsx |
| 50 | `'GSTR3B-Alpha Test Traders-Maharashtra-Apr.xltx'` | Excel template .xltx | NO | - | False | Unknown / Unknown / Unknown | `''` | `''` | `{}` | No matching pattern found | `` | GSTR3B-ClientName-State-Month.xlsx |
| 51 | `'Sales-Alpha Test Traders Alpha Test Trad...ders -Maharashtra-Apr-Mar.xlsx'` | very long client (114 chars) | yes | Sales | True | Sales / Sales / sales | `'Alpha Test Traders Alpha Test Traders Alpha Test Traders Alpha Test Traders Alpha Test Traders Alpha Test Traders'` | `'Maharashtra'` | `{'start_month': 'Apr', 'end_month': 'Mar'}` | Client name issue: Client name too long | `Alpha Test Traders Alpha Test Traders Alpha Test Traders Alpha Test Traders Alpha Test Traders Alpha Test Traders-MH` |  |
| 52 | `'GSTR3B-Ålpha Tëst Tradérs-Mahārāshtra-Apr.xlsx'` | Unicode accents in client and state | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Ålpha Tëst Tradérs'` | `'Mahārāshtra'` | `{'month': 'Apr'}` |  | `Ålpha Tëst Tradérs-MAH` |  |
| 53 | `'GSTR3B-Alpha Test Traders-Mumbai-Apr.xlsx'` | city instead of state | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'Mumbai'` | `{'month': 'Apr'}` |  | `Alpha Test Traders-MUM` |  |
| 54 | `'GSTR3B-Alpha Test Traders-Unknown Place-Apr.xlsx'` | unknown two-word state | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'Unknown Place'` | `{'month': 'Apr'}` |  | `Alpha Test Traders-UP` |  |
| 55 | `'GSTR3B-Alpha Test Traders-Jammu & Kashmir-Apr.xlsx'` | state with '&' | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'Jammu & Kashmir'` | `{'month': 'Apr'}` |  | `Alpha Test Traders-J&` |  |
| 56 | `'GSTR3B-Alpha Test Traders-27-Apr.xlsx'` | numeric state code | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'27'` | `{'month': 'Apr'}` |  | `Alpha Test Traders-27` |  |
| 57 | `'GSTR3B-Alpha Test Traders-MH-Apr.xlsx'` | two-letter state code | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'MH'` | `{'month': 'Apr'}` |  | `Alpha Test Traders-MH` |  |
| 58 | `'GSTR3B-Alpha Test Traders-Orissa-Apr.xlsx'` | old state name Orissa | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'Orissa'` | `{'month': 'Apr'}` |  | `Alpha Test Traders-OD` |  |
| 59 | `'GSTR3B-Alpha Test Traders-Telangana-Apr.xlsx'` | Telangana | yes | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'Telangana'` | `{'month': 'Apr'}` |  | `Alpha Test Traders-TS` |  |
| 60 | `'GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx\n'` | trailing newline (Linux-only name) | NO | GSTR3B | True | GSTR-3B Export / GSTR3B / gstr3b | `'Alpha Test Traders'` | `'Maharashtra'` | `{'month': 'Apr'}` |  | `Alpha Test Traders-MH` |  |

Max patterns matching any single name: 1
Logged state-code fallbacks (get_state_code warnings) seen while computing keys are listed in t2.

<!-- end of outputs/t1_filenames.md -->

### T1 observations (each RAN unless marked)

1. **No name matched two patterns** (max 1 over 60 names). READ reason, HIGH: every pattern starts with a different fixed text (`GSTR-2B-Reco-`, `ImsReco-`, optional `(n)` + `GSTR3B-`, `Sales-`, `SalesReco-`, `AnnualReport-`); `Sales-` and `SalesReco-` cannot both match because the 6th character must be `-` for one and `R` for the other. So the pattern order in `utils/constants.py:13-62` does not matter today.
2. **A hyphen inside a client name silently mis-splits the name** (row 6): `GSTR-2B-Reco-Alpha-Test-Traders-Maharashtra-Q1.xlsx` -> client `Alpha`, state `Test`, period `Traders-Maharashtra-Q1`, client key `Alpha-TES`. The file is accepted as parsed and would be filed under a client "Alpha" in a state "TES". No warning (the state "Test" passes `validate_client_name`). Same for row 37: `Sales-Reco-...` (hyphen instead of `SalesReco-`) is accepted as a **Sales** file of client `Reco`, state `Alpha Test Traders`, key `Reco-AT`. Cause: `[^-]+` groups (`utils/constants.py:15-55`). Technical point; HIGH.
3. **`.xlsm` is listed by the scanner but no pattern accepts it** (row 7): the scanner includes `.xlsm` (`utils/helpers.py:327`) but every pattern ends `\.xlsx?$`, so a correctly named `.xlsm` file always lands in "variations" with the suggestion to rename. `.xlsb`, `.xltx`, `.csv`, `.pdf` are never listed at all (rows 47-50) and are not reported anywhere (no count, no variation entry).
4. **Duplicate markers `(1)` / `(2)`**: accepted for every type. For IMS the marker is cut off the date (`(?:\s*\(\d+\))?`, rows 12-13, metadata date `31032025`). For the other types the marker stays inside the last group (`'Apr24 (1)'`, `'Apr (1)'`, `'Mar (1)'`, `'2024-25 (1)'`). The metadata is never read anywhere else (READ: `grep` for `metadata` finds only `core/file_parser.py:74,106`), so this has no effect except that both files count as two files of that type (duplicates, see T6).
5. A leading `(n)` prefix is accepted **only for GSTR-3B** (rows 20-21); `(1) Sales-...` would not parse. A leading space or `Copy of ` prefix is not accepted (rows 27-28). A space before `.xlsx` is accepted and trimmed (row 24).
6. **A blank client is "parsed"** (row 29, `GSTR3B- -Maharashtra-Apr.xlsx`): `parsed=True`, client `''`. `_process_file` then sends it to variations with the issue text "Could not parse filename" (`core/file_parser.py:210-220`) while the summary still counts it as "Successfully Parsed" (`core/file_parser.py:311`). RAN in t11: 3 files -> stats `parsed 2, unparsed 1`, but `variations 2`. The Validation summary ("Unparsed/Variations: 1") therefore disagrees with the variations list (2). HIGH.
7. `validate_client_name` (`utils/helpers.py:428-449`) is applied to the **state** as well (`core/file_parser.py:111`), so a one-letter state gives the warning text "State name issue: Client name too short" (row 30). Warnings are stored in the result but nothing reads or displays them (RAN grep: `'warnings'` appears only at `core/file_parser.py:114-119`; HIGH).
8. The IMS date part only has to be 8 digits: `99999999` and `20250331` are accepted as dates; `310325` (6 digits) and `Mar25` are rejected (rows 15-17). [GST RULE]-adjacent (period text); described only.
9. Names that the app itself creates (`ITC_Report_...`, `Sales_Report_...`) are not recognised, so if the target folder were used as a source they would all be variations (rows 45-46). `AnnualReport-...-2024_Refreshed_211026_1100.xlsx` (row 44) is accepted as an Annual Report with year text `2024_Refreshed_211026_1100`.
10. Very long client (114 chars, row 51): parsed, warning "Client name too long" stored, nothing blocks it. Unicode accents parse (row 52); the accented state `Mahārāshtra` is not in the state table and becomes code `MAH`.
11. State spelling decides the client key: `Mumbai` -> `MUM`, `Unknown Place` -> `UP` (same code as Uttar Pradesh), `Jammu & Kashmir` -> `J&`, `27` -> `27`, `MH` -> `MH`, `Orissa` -> `OD`, `Telangana` -> `TS` (rows 53-59). [GST RULE] state-code logic: described only.
12. Python's `$` also matches just before a trailing newline, so a (Linux-only) name ending in `\n` parses (row 60) - but the scanner would never list it (extension test fails). Harmless.

## 4. T2 - Name cleaning (RAN)

Functions: `sanitize_filename` (`utils/helpers.py:68-126`), `create_client_state_key` (`utils/helpers.py:620-644`), `get_state_code` (`utils/helpers.py:592-618`), `validate_client_name` (`utils/helpers.py:428-449`).

Where each is used (READ):

| function | used at | produces |
|---|---|---|
| `create_client_state_key(client, state, max_length=client_name_max_length)` | `core/file_organizer.py:83-87,97` | the **Level-2 (client) folder name**, e.g. `Beta Demo Industries Pvt Ltd-GJ`. Only abbreviates Private/Limited and truncates; does NOT remove invalid characters. |
| `sanitize_filename(client)` | `core/file_organizer.py:90` (category folder label `ITC (<client>)`, used only when "include client name" is on), `gui/handlers/processing_handler.py:292` (report file names `ITC_Report_<client>_<state>_<time>.xlsx`) | a cleaned client text |
| `sanitize_filename(file name)` | `core/file_organizer.py:271` | the **destination file name** of every copied source file |
| `get_state_code(state)` | `core/file_parser.py:229`, `core/file_organizer.py:82`, `gui/handlers/client_handler.py:169,181,218`, `gui/handlers/processing_handler.py:293`, `core/excel_handler.py:616` | the 2-letter (or invented) code |

How `sanitize_filename` works, step by step (READ): empty -> `unnamed` (72-73); split off the "extension" with `Path(name).stem` / `.suffix` (76-77); turn runs of spaces/underscores into one space (91); `Private`->`Pvt`, `Limited`->`Ltd` as whole words, any case (92-93); control characters -> `_` (99-102); each of `` <>:"/\|?*[]{}+=!@#$%^,;'"`~ `` -> `_` (96,105-106) - hyphen, dot, `&`, `(`, `)` are kept; runs of spaces/underscores -> one `_` (109); strip `_`, `.`, space at both ends (112); empty -> `unnamed` (115-116); cut to `max_length - len(extension) - 10` characters (119-120, default 200 -> 190 for `.xlsx`); add the extension back unchanged (123-124).

How `create_client_state_key` works (READ): `Private`/`Limited` abbreviation on the raw client text (624-625, spaces kept), state code (628), `client-CODE` (631); if longer than `max_length` (default 35): if `max_length - len(code) - 1 > 5` cut the client part to fit (637-639), otherwise cut the whole key to `max_length` characters (641-642).

Script `scripts/t2_names.py`, output `outputs/t2_names.md`, reproduced exactly:



<!-- RAN output: outputs/t2_names.md (headings demoted by two levels) -->

#### sanitize_filename / create_client_state_key on client names (state = 'Maharashtra')

| # | input (repr) | sanitize_filename(name) | Path(name).stem / .suffix | create_client_state_key(name,'Maharashtra') (max 35) | len | validate_client_name |
|---|---|---|---|---|---|---|
| 1 | `'Alpha Test Traders'` | `'Alpha_Test_Traders'` | `'Alpha Test Traders' / ''` | `'Alpha Test Traders-MH'` | 21 | (True, '') |
| 2 | `'Beta Demo Industries Private Limited'` | `'Beta_Demo_Industries_Pvt_Ltd'` | `'Beta Demo Industries Private Limited' / ''` | `'Beta Demo Industries Pvt Ltd-MH'` | 31 | (True, '') |
| 3 | `'Beta Demo Industries Pvt Ltd'` | `'Beta_Demo_Industries_Pvt_Ltd'` | `'Beta Demo Industries Pvt Ltd' / ''` | `'Beta Demo Industries Pvt Ltd-MH'` | 31 | (True, '') |
| 4 | `'Beta Demo Industries Pvt. Ltd.'` | `'Beta_Demo_Industries_Pvt._Ltd.'` | `'Beta Demo Industries Pvt. Ltd' / '.'` | `'Beta Demo Industries Pvt. Ltd.-MH'` | 33 | (True, '') |
| 5 | `'BETA DEMO INDUSTRIES PRIVATE LIMITED'` | `'BETA_DEMO_INDUSTRIES_Pvt_Ltd'` | `'BETA DEMO INDUSTRIES PRIVATE LIMITED' / ''` | `'BETA DEMO INDUSTRIES Pvt Ltd-MH'` | 31 | (True, '') |
| 6 | `'Gamma Sample Enterprises LLP'` | `'Gamma_Sample_Enterprises_LLP'` | `'Gamma Sample Enterprises LLP' / ''` | `'Gamma Sample Enterprises LLP-MH'` | 31 | (True, '') |
| 7 | `'Delta & Sons Mock Co.'` | `'Delta_&_Sons_Mock_Co.'` | `'Delta & Sons Mock Co' / '.'` | `'Delta & Sons Mock Co.-MH'` | 24 | (True, '') |
| 8 | `'Epsilon, Fake Exports (India) Pvt. Ltd.'` | `'Epsilon_Fake_Exports_(India)_Pvt._Ltd.'` | `'Epsilon, Fake Exports (India) Pvt. Ltd' / '.'` | `'Epsilon, Fake Exports (India) Pv-MH'` | 35 | (True, '') |
| 9 | `'Zeta_Private_Limited'` | `'Zeta_Pvt_Ltd'` | `'Zeta_Private_Limited' / ''` | `'Zeta_Private_Limited-MH'` | 23 | (True, '') |
| 10 | `'Eta-Private-Limited'` | `'Eta-Pvt-Ltd'` | `'Eta-Private-Limited' / ''` | `'Eta-Pvt-Ltd-MH'` | 14 | (True, '') |
| 11 | `'ThetaPrivateLimited'` | `'ThetaPrivateLimited'` | `'ThetaPrivateLimited' / ''` | `'ThetaPrivateLimited-MH'` | 22 | (True, '') |
| 12 | `'Privateer Unlimited Test Co'` | `'Privateer_Unlimited_Test_Co'` | `'Privateer Unlimited Test Co' / ''` | `'Privateer Unlimited Test Co-MH'` | 30 | (True, '') |
| 13 | `'Iota   Lots   Of   Spaces'` | `'Iota_Lots_Of_Spaces'` | `'Iota   Lots   Of   Spaces' / ''` | `'Iota   Lots   Of   Spaces-MH'` | 28 | (True, '') |
| 14 | `'  Kappa leading and trailing  '` | `'Kappa_leading_and_trailing'` | `'  Kappa leading and trailing  ' / ''` | `'  Kappa leading and trailing  -MH'` | 33 | (True, '') |
| 15 | `'Lambda.Dots.In.Name'` | `'Lambda.Dots.In.Name'` | `'Lambda.Dots.In' / '.Name'` | `'Lambda.Dots.In.Name-MH'` | 22 | (True, '') |
| 16 | `'Mu Test Traders.'` | `'Mu_Test_Traders.'` | `'Mu Test Traders' / '.'` | `'Mu Test Traders.-MH'` | 19 | (True, '') |
| 17 | `'Ålpha Tëst Tradérs'` | `'Ålpha_Tëst_Tradérs'` | `'Ålpha Tëst Tradérs' / ''` | `'Ålpha Tëst Tradérs-MH'` | 21 | (True, '') |
| 18 | `'अल्फा टेस्ट ट्रेडर्स'` | `'अल्फा_टेस्ट_ट्रेडर्स'` | `'अल्फा टेस्ट ट्रेडर्स' / ''` | `'अल्फा टेस्ट ट्रेडर्स-MH'` | 23 | (True, '') |
| 19 | `'Nu+Xi=Omicron! @Pi #Rho $Sigma %Tau ^Upsilon'` | `'Nu_Xi_Omicron_Pi_Rho_Sigma_Tau_Upsilon'` | `'Nu+Xi=Omicron! @Pi #Rho $Sigma %Tau ^Upsilon' / ''` | `'Nu+Xi=Omicron! @Pi #Rho $Sigma %-MH'` | 35 | (True, '') |
| 20 | `'Phi [Chi] {Psi} ~Omega` \'Quote\' "Dbl"'` | `'Phi_Chi_Psi_Omega_Quote_Dbl'` | `'Phi [Chi] {Psi} ~Omega` \'Quote\' "Dbl"' / ''` | `"Phi [Chi] {Psi} ~Omega` 'Quote' -MH"` | 35 | (False, 'Client name contains invalid characters') |
| 21 | `'Tab\tAnd\x07Bell'` | `'Tab_And_Bell'` | `'Tab\tAnd\x07Bell' / ''` | `'Tab\tAnd\x07Bell-MH'` | 15 | (True, '') |
| 22 | `'A Very Long Fake Client Name P...(99 chars)'` | `'A_Very_Long_Fake_Client_Name_P...(91 chars)'` | `'A Very Long Fake Client Name Private Limited Of Imaginary Industries And Trading Company Of Nowhere' / ''` | `'A Very Long Fake Client Name Pvt-MH'` | 35 | (True, '') |
| 23 | `'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx...(250 chars)'` | `'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx...(190 chars)'` | `'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx' / ''` | `'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx-MH'` | 35 | (False, 'Client name too long') |
| 24 | `''` | `'unnamed'` | `n/a` | `'-MH'` | 3 | (False, 'Client name cannot be empty') |
| 25 | `'...'` | `'unnamed'` | `'...' / ''` | `'...-MH'` | 6 | (True, '') |
| 26 | `'_'` | `'unnamed'` | `'_' / ''` | `'_-MH'` | 4 | (False, 'Client name too short') |
| 27 | `'Pvt Ltd'` | `'Pvt_Ltd'` | `'Pvt Ltd' / ''` | `'Pvt Ltd-MH'` | 10 | (True, '') |
| 28 | `'Gamma Sample Manufacturing Company of India One'` | `'Gamma_Sample_Manufacturing_Company_of_India_One'` | `'Gamma Sample Manufacturing Company of India One' / ''` | `'Gamma Sample Manufacturing Compa-MH'` | 35 | (True, '') |
| 29 | `'Gamma Sample Manufacturing Company of India Two'` | `'Gamma_Sample_Manufacturing_Company_of_India_Two'` | `'Gamma Sample Manufacturing Company of India Two' / ''` | `'Gamma Sample Manufacturing Compa-MH'` | 35 | (True, '') |

#### sanitize_filename on whole file names (what file_organizer.py:271 uses as the destination name)

| # | input | output | changed? |
|---|---|---|---|
| 1 | `'GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx'` | `'GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx'` | yes |
| 2 | `'GSTR-2B-Reco-Beta Demo Industries Private Limited-Gujarat-Apr24-Mar25.xlsx'` | `'GSTR-2B-Reco-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr24-Mar25.xlsx'` | yes |
| 3 | `'Sales-Beta Demo Industries Pvt. Ltd.-Gujarat-Apr-Mar.xlsx'` | `'Sales-Beta_Demo_Industries_Pvt._Ltd.-Gujarat-Apr-Mar.xlsx'` | yes |
| 4 | `'(1) GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx'` | `'(1)_GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx'` | yes |
| 5 | `'ImsReco-Alpha Test Traders-Maharashtra-31032025 (1).xlsx'` | `'ImsReco-Alpha_Test_Traders-Maharashtra-31032025_(1).xlsx'` | yes |
| 6 | `'SalesReco-Delta & Sons Mock Co.-Karnataka-FY 2024-25.xlsx'` | `'SalesReco-Delta_&_Sons_Mock_Co.-Karnataka-FY_2024-25.xlsx'` | yes |
| 7 | `'AnnualReport-Epsilon, Fake Exports-Delhi-2024-25.XLSX'` | `'AnnualReport-Epsilon_Fake_Exports-Delhi-2024-25.XLSX'` | yes |
| 8 | `'Sales-Zeta_Private_Limited-Delhi-Apr-Mar.xls'` | `'Sales-Zeta_Pvt_Ltd-Delhi-Apr-Mar.xls'` | yes |
| 9 | `'GSTR3B-Ålpha Tëst Tradérs-Maharashtra-Apr.xlsx'` | `'GSTR3B-Ålpha_Tëst_Tradérs-Maharashtra-Apr.xlsx'` | yes |
| 10 | `'GSTR3B-Alpha Test Traders-Maharashtra-Apr .xlsx'` | `'GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx'` | yes |
| 11 | `'GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx.xlsx'` | `'GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx.xlsx'` | yes |
| 12 | `'Sales-Long Fake Name Long Fake Name Long...(256 chars)'` | `"Sales-Long_Fake_Name_Long_Fake_Name_Long...(190 chars) ends 'ke_Name.xlsx'"` | yes |

#### max_length edge: sanitize_filename(name, max_length) with short limits

- max_length=200: sanitize_filename('Alpha Test Traders Report.xlsx', 200) -> 'Alpha_Test_Traders_Report.xlsx'
- max_length=20: sanitize_filename('Alpha Test Traders Report.xlsx', 20) -> 'Alpha.xlsx'
- max_length=15: sanitize_filename('Alpha Test Traders Report.xlsx', 15) -> '.xlsx'
- max_length=10: sanitize_filename('Alpha Test Traders Report.xlsx', 10) -> 'Alpha_Test_Traders_R.xlsx'
- max_length=5: sanitize_filename('Alpha Test Traders Report.xlsx', 5) -> 'Alpha_Test_Trad.xlsx'

#### create_client_state_key with different max_length values (client_name_max_length setting)

- max_length=35: 'Beta Demo Industries Private Limited'+Maharashtra -> 'Beta Demo Industries Pvt Ltd-MH' (len 31); 'Beta Demo Industries'+'Unknown Long State Name' -> 'Beta Demo Industries-UL'
- max_length=10: 'Beta Demo Industries Private Limited'+Maharashtra -> 'Beta De-MH' (len 10); 'Beta Demo Industries'+'Unknown Long State Name' -> 'Beta De-UL'
- max_length=8: 'Beta Demo Industries Private Limited'+Maharashtra -> 'Beta Dem' (len 8); 'Beta Demo Industries'+'Unknown Long State Name' -> 'Beta Dem'
- max_length=7: 'Beta Demo Industries Private Limited'+Maharashtra -> 'Beta De' (len 7); 'Beta Demo Industries'+'Unknown Long State Name' -> 'Beta De'
- max_length=6: 'Beta Demo Industries Private Limited'+Maharashtra -> 'Beta D' (len 6); 'Beta Demo Industries'+'Unknown Long State Name' -> 'Beta D'
- max_length=5: 'Beta Demo Industries Private Limited'+Maharashtra -> 'Beta ' (len 5); 'Beta Demo Industries'+'Unknown Long State Name' -> 'Beta '
- max_length=3: 'Beta Demo Industries Private Limited'+Maharashtra -> 'Bet' (len 3); 'Beta Demo Industries'+'Unknown Long State Name' -> 'Bet'
- max_length=1: 'Beta Demo Industries Private Limited'+Maharashtra -> 'B' (len 1); 'Beta Demo Industries'+'Unknown Long State Name' -> 'B'
- max_length=0: 'Beta Demo Industries Private Limited'+Maharashtra -> '' (len 0); 'Beta Demo Industries'+'Unknown Long State Name' -> ''
- max_length=-5: 'Beta Demo Industries Private Limited'+Maharashtra -> 'Beta Demo Industries Pvt L' (len 26); 'Beta Demo Industries'+'Unknown Long State Name' -> 'Beta Demo Industri'

#### get_state_code (utils/helpers.py:592) on state strings  [GST RULE: state-code logic; TIME-VARYING list]

| input (repr) | output | log warning |
|---|---|---|
| `'Maharashtra'` | `'MH'` |  |
| `'maharashtra'` | `'MH'` |  |
| `'  MAHARASHTRA  '` | `'MH'` |  |
| `'MH'` | `'MH'` | Unknown state 'MH', using fallback: MH |
| `'mh'` | `'MH'` | Unknown state 'mh', using fallback: MH |
| `'Mh.'` | `'MH.'` | Unknown state 'Mh.', using fallback: MH. |
| `'Orissa'` | `'OD'` |  |
| `'Odisha'` | `'OD'` |  |
| `'Pondicherry'` | `'PY'` |  |
| `'Puducherry'` | `'PY'` |  |
| `'Delhi'` | `'DL'` |  |
| `'NCT of Delhi'` | `'NO'` | Unknown state 'NCT of Delhi', using fallback: NO |
| `'National Capital Territory of Delhi'` | `'DL'` |  |
| `'New Delhi'` | `'ND'` | Unknown state 'New Delhi', using fallback: ND |
| `'Delhi NCR'` | `'DN'` | Unknown state 'Delhi NCR', using fallback: DN |
| `'Jammu and Kashmir'` | `'JK'` |  |
| `'Jammu & Kashmir'` | `'J&'` | Unknown state 'Jammu & Kashmir', using fallback: J& |
| `'J&K'` | `'J&K'` | Unknown state 'J&K', using fallback: J&K |
| `'Andaman and Nicobar Islands'` | `'AN'` |  |
| `'Andaman & Nicobar'` | `'A&'` | Unknown state 'Andaman & Nicobar', using fallback: A& |
| `'Dadra and Nagar Haveli and Daman and Diu'` | `'DD'` |  |
| `'Daman and Diu'` | `'DA'` | Unknown state 'Daman and Diu', using fallback: DA |
| `'Telangana'` | `'TS'` |  |
| `'Uttar Pradesh'` | `'UP'` |  |
| `'Unknown Place'` | `'UP'` | Unknown state 'Unknown Place', using fallback: UP |
| `'Uttarakhand'` | `'UK'` |  |
| `'Uttaranchal'` | `'UTT'` | Unknown state 'Uttaranchal', using fallback: UTT |
| `'Tamil Nadu'` | `'TN'` |  |
| `'TamilNadu'` | `'TAM'` | Unknown state 'TamilNadu', using fallback: TAM |
| `'Tamil-Nadu'` | `'TAM'` | Unknown state 'Tamil-Nadu', using fallback: TAM |
| `'27'` | `'27'` | Unknown state '27', using fallback: 27 |
| `'07'` | `'07'` | Unknown state '07', using fallback: 07 |
| `'Other Territory'` | `'OT'` | Unknown state 'Other Territory', using fallback: OT |
| `'Mahārāshtra'` | `'MAH'` | Unknown state 'Mahārāshtra', using fallback: MAH |
| `'A'` | `'A'` | Unknown state 'A', using fallback: A |
| `''` | `''` |  |
| `None` | `None` |  |

#### Client names containing a dot: Path(name).suffix treats the text after the LAST dot as an 'extension' (helpers.py:76-77, 123-124)

| client name | sanitize_filename | Path.stem | Path.suffix |
|---|---|---|---|
| `Omega Co. Private Limited` | `'Omega_Co. Private Limited'` | `'Omega Co'` | `'. Private Limited'` |
| `Omega Co. Pvt. Ltd` | `'Omega_Co._Pvt. Ltd'` | `'Omega Co. Pvt'` | `'. Ltd'` |
| `Sigma Inc. & Co. Private Limited` | `'Sigma_Inc._&_Co. Private Limited'` | `'Sigma Inc. & Co'` | `'. Private Limited'` |
| `Tau Enterprises Pvt.Ltd` | `'Tau_Enterprises_Pvt.Ltd'` | `'Tau Enterprises Pvt'` | `'.Ltd'` |
| `Upsilon Traders (P) Ltd.` | `'Upsilon_Traders_(P)_Ltd.'` | `'Upsilon Traders (P) Ltd'` | `'.'` |
| `Phi Agencies Pvt. Ltd.-Extra` | `'Phi_Agencies_Pvt._Ltd.-Extra'` | `'Phi Agencies Pvt. Ltd'` | `'.-Extra'` |

Report name built the way processing_handler.py:292-298 does (safe_client + state code + timestamp):
- 'Omega Co. Private Limited' -> ITC_Report_Omega_Co. Private Limited_MH_211026_1030.xlsx
- 'Upsilon Traders (P) Ltd.' -> ITC_Report_Upsilon_Traders_(P)_Ltd._MH_211026_1030.xlsx

(Python 3.14.6)

<!-- end of outputs/t2_names.md -->

### T2 observations

1. **The folder name and the file names are cleaned by two different functions**, so they differ: folder `Beta Demo Industries Pvt Ltd-GJ` (spaces kept, `create_client_state_key`) vs files `GSTR3B-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr.xlsx` (underscores, `sanitize_filename`). `Zeta_Private_Limited` becomes `Zeta_Pvt_Ltd` in file names but the folder keeps `Zeta_Private_Limited-MH` (row 9) because `create_client_state_key` uses `\bPrivate\b` on the raw text, where `_` is a word character. RAN, HIGH.
2. **Client names containing a dot** (common: "Co.", "Pvt. Ltd."): `sanitize_filename` treats everything after the LAST dot as a file extension and re-attaches it untouched (`utils/helpers.py:76-77,123-124`). `Omega Co. Private Limited` -> `Omega_Co. Private Limited` (spaces kept, not abbreviated); a report would be named `ITC_Report_Omega_Co. Private Limited_MH_211026_1030.xlsx` (built exactly as `gui/handlers/processing_handler.py:292-298` does). On this Python 3.14.6 a trailing dot also counts as an extension (`'Mu Test Traders.'` -> suffix `'.'`), so the result can end with a dot. I did not check older Python versions (LOW confidence that they behave the same for the trailing-dot case; the "text after the last dot" behaviour is long-standing, HIGH). Only affects client-name uses (report names, category labels); for whole file names the suffix is `.xlsx` so they are fine.
3. **Truncation collisions**: with the default 35, `Gamma Sample Manufacturing Company of India One` and `...Two` both become `Gamma Sample Manufacturing Compa-MH` (rows 28-29). T6 section 7 shows what that does (two clients share one folder). RAN, HIGH.
4. **Small limits drop the state code**: for a 2-letter code, any `max_length` of 8 or less gives a key with no state at all (`'Beta Dem'`, `'Bet'`, `''` for 0). `FileOrganizer` turns 0 and negatives into 35 (`core/file_organizer.py:42`), and the Setup spin box offers 15-100 (`gui/tabs/setup_tab.py:440`), so 1-8 needs a typed value - Tk spin boxes accept typed values unless validation is set, and none is set there (READ, MEDIUM). With an unknown long state name the invented code is 2 letters (`UL`) so the same applies.
5. `sanitize_filename` with `max_length` below `len(extension)+10` slices with a negative or zero number and returns a name LONGER than the limit or with an empty stem (`'.xlsx'` at 15). No caller passes a small limit (all callers use the default 200; READ), so this is unreachable today. LOW impact.
6. `create_client_state_key` keeps control characters, quotes, braces and leading/trailing spaces (rows 14, 20, 21). Real file names on Windows cannot contain `<>:"/\|?*` or control characters, and the parser strips spaces (`core/file_parser.py:102-104`), so from real scans only `&`, `,`, `(`, `)`, `.`, `'`, `{`, `}`, `[`, `]`, `!`, `@`, `#`, `$`, `%`, `^`, `~`, `` ` ``, `+`, `=` and Unicode can reach the folder name; all are legal in Windows folder names (READ, HIGH).
7. `'Pvt Ltd'` vs `'Private Limited'` spellings of one client produce two different parser keys but the SAME folder name and the SAME cleaned file names - see T8 section 1 for the overwrite this causes.
8. [GST RULE] `get_state_code` (table at the end of the output): known names map to 2-letter codes; `MH`/`mh` pass through as `MH` (fallback, with a warning in the log); `Mh.` -> `MH.`; `NCT of Delhi` -> `NO`, `New Delhi` -> `ND`, `Delhi NCR` -> `DN`; `Jammu & Kashmir` -> `J&`, `J&K` -> `J&K`, `Andaman & Nicobar` -> `A&`; `Daman and Diu` -> `DA`; `Uttaranchal` -> `UTT`; `TamilNadu`/`Tamil-Nadu` -> `TAM`; numbers pass through (`27`, `07`); `Unknown Place` -> `UP` (same as Uttar Pradesh); `''` and `None` are returned unchanged. TIME-VARYING: the list of states/UTs and their names (`utils/constants.py:251-291`). The code uses letter codes only; it has no numeric GST state-code table (READ, HIGH). Whether any of these mappings is right is for the author.

## 5. T3 - Financial year and tax period logic

**Finding first: project A contains no financial-year or tax-period logic at all.** READ + RAN (search): at HEAD, a search of all Python files for financial/fiscal/FY/April/March/tax period/strptime finds only run time stamps and the comment in `power_query_extractor/gui/extractor_window.py:645-647`; a history search (`git log -G` on A/main) finds such words only in test file names (commit 56aaad9, 6c1e1c6). What exists instead:

- The parser keeps the period part of a name as plain text in `metadata` (`core/file_parser.py:98-106`): `period` (GSTR-2B Reco, Sales Reco), `date` (IMS Reco, must be 8 digits), `month` (GSTR-3B), `start_month`/`end_month` (Sales), `year` (Annual Report). [GST RULE] area (tax period); TIME-VARYING. **Nothing ever reads `metadata`** (RAN grep: only `core/file_parser.py:74,106`). So the code never checks that files belong to the same year, never orders them by period, and never rejects an out-of-year file.
- When a client has two files of one type, the template is given `files[0]` (`core/excel_handler.py:629,640,651,667,683,699`), i.e. the first in **path-sorted order** (`utils/helpers.py:361`). Which period's file feeds the ITC/Sales report is therefore decided by alphabetical order of the file names. [GST RULE] (which period's figures are used) - described only, see the table in (b).
- Run time stamps (not tax periods): `get_timestamp()` = `DDMMYY HHMM`, `get_date_only()` = `DDMMYY`, `get_safe_timestamp()` = `DDMMYY_HHMM` (`utils/helpers.py:25-35`); the consolidated extraction report uses `YYYYMMDD_HHMM` (`power_query_extractor/core/data_consolidator.py:175`); the refreshed-copy suffix uses `DDMMYY_HHMM` (`power_query_extractor/core/report_processor.py:128`). All are local clock time with minute resolution.

Script `scripts/t3_periods_time.py`, output `outputs/t3_periods_time.md`, reproduced exactly (the "refreshed suffix" and "consolidated report" columns apply the same format strings the code uses to the fixed moments; the other columns call the real functions with the clock replaced):



<!-- RAN output: outputs/t3_periods_time.md (headings demoted by two levels) -->

#### (a) Period text in file names -> what parse_filename keeps  [GST RULE area: tax period / FY; TIME-VARYING]

| file name | type | metadata kept (never read anywhere else) |
|---|---|---|
| `GSTR3B-Alpha Test Traders-Maharashtra-Mar.xlsx` | GSTR-3B Export | `{'month': 'Mar'}` |
| `GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx` | GSTR-3B Export | `{'month': 'Apr'}` |
| `GSTR3B-Alpha Test Traders-Maharashtra-March 2025.xlsx` | GSTR-3B Export | `{'month': 'March 2025'}` |
| `GSTR3B-Alpha Test Traders-Maharashtra-042025.xlsx` | GSTR-3B Export | `{'month': '042025'}` |
| `GSTR3B-Alpha Test Traders-Maharashtra-13.xlsx` | GSTR-3B Export | `{'month': '13'}` |
| `GSTR3B-Alpha Test Traders-Maharashtra-Mar25.xlsx` | GSTR-3B Export | `{'month': 'Mar25'}` |
| `GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Apr24-Mar25.xlsx` | GSTR-2B Reco | `{'period': 'Apr24-Mar25'}` |
| `GSTR-2B-Reco-Alpha Test Traders-Maharashtra-2024-25.xlsx` | GSTR-2B Reco | `{'period': '2024-25'}` |
| `GSTR-2B-Reco-Alpha Test Traders-Maharashtra-24-25.xlsx` | GSTR-2B Reco | `{'period': '24-25'}` |
| `GSTR-2B-Reco-Alpha Test Traders-Maharashtra-FY2024-2025.xlsx` | GSTR-2B Reco | `{'period': 'FY2024-2025'}` |
| `ImsReco-Alpha Test Traders-Maharashtra-31032025.xlsx` | IMS Reco | `{'date': '31032025'}` |
| `ImsReco-Alpha Test Traders-Maharashtra-01042025.xlsx` | IMS Reco | `{'date': '01042025'}` |
| `ImsReco-Alpha Test Traders-Maharashtra-31022025.xlsx` | IMS Reco | `{'date': '31022025'}` |
| `ImsReco-Alpha Test Traders-Maharashtra-20250331.xlsx` | IMS Reco | `{'date': '20250331'}` |
| `Sales-Alpha Test Traders-Maharashtra-Apr-Mar.xlsx` | Sales | `{'start_month': 'Apr', 'end_month': 'Mar'}` |
| `Sales-Alpha Test Traders-Maharashtra-Mar-Apr.xlsx` | Sales | `{'start_month': 'Mar', 'end_month': 'Apr'}` |
| `Sales-Alpha Test Traders-Maharashtra-Apr24-Mar25.xlsx` | Sales | `{'start_month': 'Apr24', 'end_month': 'Mar25'}` |
| `Sales-Alpha Test Traders-Maharashtra-2024-2025.xlsx` | Sales | `{'start_month': '2024', 'end_month': '2025'}` |
| `Sales-Alpha Test Traders-Maharashtra-0424-0325.xlsx` | Sales | `{'start_month': '0424', 'end_month': '0325'}` |
| `SalesReco-Alpha Test Traders-Maharashtra-FY 2024-25.xlsx` | Sales Reco | `{'period': 'FY 2024-25'}` |
| `SalesReco-Alpha Test Traders-Maharashtra-2025.xlsx` | Sales Reco | `{'period': '2025'}` |
| `AnnualReport-Alpha Test Traders-Maharashtra-2024-25.xlsx` | Annual Report | `{'year': '2024-25'}` |
| `AnnualReport-Alpha Test Traders-Maharashtra-24-25.xlsx` | Annual Report | `{'year': '24-25'}` |
| `AnnualReport-Alpha Test Traders-Maharashtra-2024.xlsx` | Annual Report | `{'year': '2024'}` |
| `AnnualReport-Alpha Test Traders-Maharashtra-FY.xlsx` | Annual Report | `{'year': 'FY'}` |
| `AnnualReport-Alpha Test Traders-Maharashtra-x.xlsx` | Annual Report | `{'year': 'x'}` |

#### (b) Two periods of the same type for one client -> which one prepare_template_data passes to the template

files per type in the order stored (scan_excel_files sorts by full path, utils/helpers.py:361):
- Annual Report: ['AnnualReport-Alpha Test Traders-Maharashtra-2023-24.xlsx', 'AnnualReport-Alpha Test Traders-Maharashtra-2024-25.xlsx']
- GSTR-2B Reco: ['GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Apr24-Mar25.xlsx', 'GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Apr25-Mar26.xlsx']
- GSTR-3B Export: ['GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx', 'GSTR3B-Alpha Test Traders-Maharashtra-Mar.xlsx']
- IMS Reco: ['ImsReco-Alpha Test Traders-Maharashtra-01042025.xlsx', 'ImsReco-Alpha Test Traders-Maharashtra-31032025.xlsx']
- Sales: ['Sales-Alpha Test Traders-Maharashtra-Apr-Sep.xlsx', 'Sales-Alpha Test Traders-Maharashtra-Oct-Mar.xlsx']
- Sales Reco: ['SalesReco-Alpha Test Traders-Maharashtra-FY 2024-25 (1).xlsx', 'SalesReco-Alpha Test Traders-Maharashtra-FY 2024-25.xlsx']
status: 'Has duplicates'; extra_files: ['Multiple Annual Report (2 files)', 'Multiple GSTR-2B Reco (2 files)', 'Multiple IMS Reco (2 files)', 'Multiple Sales (2 files)', 'Multiple Sales Reco (2 files)']; missing: []; completeness: 100.0

prepare_template_data(..., 'ITC') -> {'gstr3b_folder': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t3\\V\\GSTR3B', 'annual_folder': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t3\\V', 'annual_filename': 'AnnualReport-Alpha Test Traders-Maharashtra-2023-24', 'gstr2b_folder': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t3\\V\\ITC', 'gstr2b_filename': 'GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Apr24-Mar25', 'ims_folder': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t3\\V\\ITC', 'ims_filename': 'ImsReco-Alpha Test Traders-Maharashtra-01042025'}
prepare_template_data(..., 'Sales') -> {'sales_folder': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t3\\V\\Sales', 'sales_filename': 'Sales-Alpha Test Traders-Maharashtra-Apr-Sep', 'annual_folder': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t3\\V', 'annual_filename': 'AnnualReport-Alpha Test Traders-Maharashtra-2023-24', 'sales_reco_folder': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t3\\V\\Sales', 'sales_reco_filename': 'SalesReco-Alpha Test Traders-Maharashtra-FY 2024-25 (1)'}
prepare_template_data(..., 'Other') -> {}

#### (c) Run timestamps (utils/helpers.py:25-35) and the names built from them

| moment | get_timestamp() | get_date_only() | get_safe_timestamp() | Level-1 folder | Version folder | ITC report name (processing_handler.py:294-298) | refreshed suffix (report_processor.py:128) | consolidated report (data_consolidator.py:175-176) |
|---|---|---|---|---|---|---|---|---|
| 2025-03-31 23:59:30 | `310325 2359` | `310325` | `310325_2359` | `Annual Statement-310325 2359` | `Version-310325 2359` | `ITC_Report_Alpha_Test_Traders_MH_310325_2359.xlsx` | `_Refreshed_310325_2359` | `PQ_Extraction_Report_20250331_2359.xlsx` |
| 2025-04-01 00:00:10 | `010425 0000` | `010425` | `010425_0000` | `Annual Statement-010425 0000` | `Version-010425 0000` | `ITC_Report_Alpha_Test_Traders_MH_010425_0000.xlsx` | `_Refreshed_010425_0000` | `PQ_Extraction_Report_20250401_0000.xlsx` |
| 2025-12-31 23:59:00 | `311225 2359` | `311225` | `311225_2359` | `Annual Statement-311225 2359` | `Version-311225 2359` | `ITC_Report_Alpha_Test_Traders_MH_311225_2359.xlsx` | `_Refreshed_311225_2359` | `PQ_Extraction_Report_20251231_2359.xlsx` |
| 2026-01-01 00:00:00 | `010126 0000` | `010126` | `010126_0000` | `Annual Statement-010126 0000` | `Version-010126 0000` | `ITC_Report_Alpha_Test_Traders_MH_010126_0000.xlsx` | `_Refreshed_010126_0000` | `PQ_Extraction_Report_20260101_0000.xlsx` |
| 2026-01-02 09:05:00 | `020126 0905` | `020126` | `020126_0905` | `Annual Statement-020126 0905` | `Version-020126 0905` | `ITC_Report_Alpha_Test_Traders_MH_020126_0905.xlsx` | `_Refreshed_020126_0905` | `PQ_Extraction_Report_20260102_0905.xlsx` |
| 2026-02-01 09:05:00 | `010226 0905` | `010226` | `010226_0905` | `Annual Statement-010226 0905` | `Version-010226 0905` | `ITC_Report_Alpha_Test_Traders_MH_010226_0905.xlsx` | `_Refreshed_010226_0905` | `PQ_Extraction_Report_20260201_0905.xlsx` |

##### Sorting day-first names (why extractor_window.py:643-648 uses mtime instead)

names: ['Version-300626 1000', 'Version-020726 0900', 'Version-311225 2359', 'Version-010126 0000', 'Version-150126 1200']
max() by text: Version-311225 2359 | sorted by text: ['Version-010126 0000', 'Version-020726 0900', 'Version-150126 1200', 'Version-300626 1000', 'Version-311225 2359']
chronological: ['Version-311225 2359', 'Version-010126 0000', 'Version-150126 1200', 'Version-300626 1000', 'Version-020726 0900'] | latest: Version-020726 0900

##### 'Latest by modification time' (file_organizer.py:184, extractor_window.py:631,648) after touching an OLDER folder

before: Annual Statement-150426 0900
after a file is written into the older folder: Annual Statement-010426 0900

<!-- end of outputs/t3_periods_time.md -->

### T3 observations

1. (a) All period strings - including impossible ones (`13`, `31022025`, `x`, `FY`, `Mar-Apr`) - are accepted and kept verbatim. Two-digit and four-digit year forms (`24-25`, `2024-25`, `FY2024-2025`) are all accepted; nothing compares them. [GST RULE]; TIME-VARYING. RAN, HIGH.
2. (b) With two files of each type for one client, the template data received (RAN): Annual Report `...-2023-24` (not `...-2024-25`), GSTR-2B Reco `...-Apr24-Mar25` (not `...-Apr25-Mar26`), IMS Reco `...-01042025` (not `...-31032025`: a `DDMMYYYY` text sorts by day first, so it is not date order), Sales `...-Apr-Sep` (not `...-Oct-Mar`), Sales Reco `...-FY 2024-25 (1)` (the "(1)" copy, because a space sorts before a dot, so `x (1).xlsx` comes before `x.xlsx`). The client's status is "Has duplicates" with 100% completeness; nothing tells the user which of the two files went into the report. [GST RULE] (which period/file feeds the report) - described only; whether the earliest-by-name is intended is a question for the author.
3. (c) Folder and file names use day-first stamps. Text order is not date order: `max()` of five `Version-...` names picks `Version-311225 2359` although `Version-020726 0900` is the newest. The code avoids sorting by name and picks the newest folder by modification time instead (`core/file_organizer.py:184`; `power_query_extractor/gui/extractor_window.py:631,648`).
4. **"Newest by modification time" is not "most recently created"**: RAN - after one file is written into an older `Annual Statement-...` folder, that older folder becomes the one chosen. Any file created, renamed or deleted directly inside an old run folder (for example a summary workbook opened and saved there, which makes Excel write a temporary file and rename it) redirects later "Re-Run"/"Resume" runs and the Step 4 scan to that older folder. HIGH that the code behaves this way; MEDIUM that Windows updates folder times the same way (NTFS updates a folder's time when entries are added or renamed in it; I could not run Windows).
5. Minute resolution: two runs inside the same minute get the same folder names - see T6 section 5 (backups and silent overwrites).
6. The two clock reads that name one run are separate: `ProcessingHandler.process_files_thread` takes its own stamp at `gui/handlers/processing_handler.py:186` (used for report and summary names), while `FileOrganizer` takes another at `core/file_organizer.py:43` (used for folder names). If the minute changes between those two lines the report names carry a different minute from their Version folder. READ; the window is microseconds, LOW likelihood.

## 6. T4 - Cell mappings and the consolidator

**How the data flows (READ):**

1. `cell_mappings.py` defines `EXTRACTION_CONFIG` (`power_query_extractor/config/cell_mappings.py:33-242`): one entry per output sheet, each with a `report_type` (ITC or Sales) and a list of mappings (input sheet, input cell, output column, calculation). At import time it builds `CELL_MAPPINGS[report_type][input_sheet][input_cell] = field_name` with `field_name = "<output sheet>_<output column>"` with spaces turned into `_` (`cell_mappings.py:260-278`), and `FIELD_DISPLAY_NAMES`.
2. Step 4 refreshes each report copy (Windows/Excel only) and then calls `ReportProcessor._extract_data(file, report_type)` (`power_query_extractor/core/report_processor.py:527-636`). For every sheet in `CELL_MAPPINGS[report_type]` it looks for a sheet whose name is exactly equal (`report_processor.py:564-566`, case-sensitive); if missing, every field of that sheet is set to `None` (570-576); otherwise it reads `ws.Range(cell).Value` (584) and stores it twice: under the field name and under the bare cell reference such as `J15` (585-588).
3. `DataConsolidator.create_report(results, level1_folder)` (`power_query_extractor/core/data_consolidator.py:20-186`) writes one sheet per `EXTRACTION_CONFIG` entry: a title in A1 merged A1:C1 (66-71) or description lines (53-65), a header row (`Client Name`, `Successfully Run`, then one column per mapping; 86-99), and one row per client (104-166). For each mapping the value is: `Missing` if the refresh status was not a success (128-129); else the result of `_calculate_value` if a calculation is configured (130-136); else **the first entry in the extracted data whose key CONTAINS the output column's text** (139-145); and only if that gave `None`, the exact key (148-150). `None` or `''` -> `Missing` (153-154). Non-zero numbers get the format `#,##0.00` (161-162). A `Summary` sheet is inserted first (172, 220-248). The file is `PQ_Extraction_Report_YYYYMMDD_HHMM.xlsx` in the folder passed in (175-179), which the extractor sets to the newest `Annual Statement-*` folder (`power_query_extractor/gui/extractor_window.py:631-632,945`). The "Client Name" written is the Level-2 folder name (`extractor_window.py:654`), e.g. `Alpha Test Traders-MH`.

**How I ran it:** `_extract_data` needs Excel COM, which does not exist here. I replaced the module's `pythoncom`/`win32com` with `scripts/fakecom.py`, which opens the workbook with openpyxl and answers `wb.Sheets`, `.Name`, `ws.Range(ref).Value`, `Close`, `Quit`. With `data_only=True` it returns what is cached in the file. **A formula that Excel has never calculated has no cached value, so openpyxl returns `None`; real Excel through COM would return the calculated number** (not verifiable here). Made-up workbooks with fake numbers were created with openpyxl in `work/t4/`:

- `itc1`: all ITC sheets except `T8 of R9` (missing sheet); `Diff GSTR-3B`!K15 holds the TEXT `'1,234.00'`; `Reco`!C37 holds the never-calculated formula `=SUM(C1:C3)`; `Reco`!E37 is empty; `Reco`!D37 is 0; `T6 of R9`!K29 is the text `Yes`.
- `sales1`: `Diff - GSTR-1` with I1, N1, O1 numbers, P1 a formula, Q1 empty.
- `sales_case`: the same Sales sheet but named `diff - gstr-1` (lower case).
- `itc_nulls`: all ITC sheets except `Diff GSTR-3B`.

Results fed to the consolidator: client 1 both successes; client 2 ITC refresh **failed** but data extracted, Sales from `sales_case`; client 3 ITC from `itc_nulls`, Sales "Not selected"; client 4 = exactly what `process_client` returns on a machine without pywin32 (`report_processor.py:66-77`).

[GST RULE] The whole mapping table below decides which report cell each figure is read from and under which label it is reported (including the labels "Table 6J Match" and "Table 8D Limits" and the input sheets `T6 of R9` / `T8 of R9`). TIME-VARYING: table numbers, sheet layouts and cell addresses depend on the report templates and the return formats of the time. Described only.

Script `scripts/t4_cells.py`, output `outputs/t4_cells.md`, reproduced exactly:



<!-- RAN output: outputs/t4_cells.md (headings demoted by two levels) -->

#### Part A - every mapping (EXTRACTION_CONFIG, cell_mappings.py:33-242)

| # | output sheet (consolidated report) | report_type | input_sheet | input_cell | output_column | calculation | derived field name (cell_mappings.py:276) |
|---|---|---|---|---|---|---|---|
| 1 | GSTR-3B Difference | ITC | `Diff GSTR-3B` | J15 | Diff IGST | None | `GSTR-3B_Difference_Diff_IGST` |
| 2 | GSTR-3B Difference | ITC | `Diff GSTR-3B` | K15 | Diff CGST | None | `GSTR-3B_Difference_Diff_CGST` |
| 3 | GSTR-3B Difference | ITC | `Diff GSTR-3B` | L15 | Diff SGST | None | `GSTR-3B_Difference_Diff_SGST` |
| 4 | GSTR-1 Difference | Sales | `Diff - GSTR-1` | N1 | Diff Taxable Value | None | `GSTR-1_Difference_Diff_Taxable_Value` |
| 5 | GSTR-1 Difference | Sales | `Diff - GSTR-1` | O1 | Diff IGST | None | `GSTR-1_Difference_Diff_IGST` |
| 6 | GSTR-1 Difference | Sales | `Diff - GSTR-1` | P1 | Diff CGST | None | `GSTR-1_Difference_Diff_CGST` |
| 7 | GSTR-1 Difference | Sales | `Diff - GSTR-1` | Q1 | Diff SGST | None | `GSTR-1_Difference_Diff_SGST` |
| 8 | Turnover GSTR-1 | Sales | `Diff - GSTR-1` | I1 | Taxable Value | None | `Turnover_GSTR-1_Taxable_Value` |
| 9 | Reco Diff | ITC | `Reco` | C37 | IGST Diff | None | `Reco_Diff_IGST_Diff` |
| 10 | Reco Diff | ITC | `Reco` | D37 | CGST Diff | None | `Reco_Diff_CGST_Diff` |
| 11 | Reco Diff | ITC | `Reco` | E37 | SGST Diff | None | `Reco_Diff_SGST_Diff` |
| 12 | Table 6J Match | ITC | `T6 of R9` | K29 | Table 6J Match | None | `Table_6J_Match_Table_6J_Match` |
| 13 | Table 8D Limits | ITC | `T8 of R9` | G6 | Table 8D Limits | None | `Table_8D_Limits_Table_8D_Limits` |
| 14 | IMS Reco | ITC | `IMS_Reco` | AA1 | IGST | None | `IMS_Reco_IGST` |
| 15 | IMS Reco | ITC | `IMS_Reco` | AD1 | CGST | None | `IMS_Reco_CGST` |
| 16 | IMS Reco | ITC | `IMS_Reco` | AG1 | SGST | None | `IMS_Reco_SGST` |
| 17 | Extra In Purchase | ITC | `ExtraInPurchase` | O1 | IGST | None | `Extra_In_Purchase_IGST` |
| 18 | Extra In Purchase | ITC | `ExtraInPurchase` | P1 | CGST | None | `Extra_In_Purchase_CGST` |
| 19 | Extra In Purchase | ITC | `ExtraInPurchase` | Q1 | SGST | None | `Extra_In_Purchase_SGST` |
| 20 | Extra In GSTR2B | ITC | `ExtraInGSTR2B` | S1 | IGST | None | `Extra_In_GSTR2B_IGST` |
| 21 | Extra In GSTR2B | ITC | `ExtraInGSTR2B` | T1 | CGST | None | `Extra_In_GSTR2B_CGST` |
| 22 | Extra In GSTR2B | ITC | `ExtraInGSTR2B` | U1 | SGST | None | `Extra_In_GSTR2B_SGST` |

Total mappings: 22; output sheets: 9

CELL_MAPPINGS (built at import, cell_mappings.py:260-278):
```
{
 "ITC": {
  "Diff GSTR-3B": {
   "J15": "GSTR-3B_Difference_Diff_IGST",
   "K15": "GSTR-3B_Difference_Diff_CGST",
   "L15": "GSTR-3B_Difference_Diff_SGST"
  },
  "Reco": {
   "C37": "Reco_Diff_IGST_Diff",
   "D37": "Reco_Diff_CGST_Diff",
   "E37": "Reco_Diff_SGST_Diff"
  },
  "T6 of R9": {
   "K29": "Table_6J_Match_Table_6J_Match"
  },
  "T8 of R9": {
   "G6": "Table_8D_Limits_Table_8D_Limits"
  },
  "IMS_Reco": {
   "AA1": "IMS_Reco_IGST",
   "AD1": "IMS_Reco_CGST",
   "AG1": "IMS_Reco_SGST"
  },
  "ExtraInPurchase": {
   "O1": "Extra_In_Purchase_IGST",
   "P1": "Extra_In_Purchase_CGST",
   "Q1": "Extra_In_Purchase_SGST"
  },
  "ExtraInGSTR2B": {
   "S1": "Extra_In_GSTR2B_IGST",
   "T1": "Extra_In_GSTR2B_CGST",
   "U1": "Extra_In_GSTR2B_SGST"
  }
 },
 "Sales": {
  "Diff - GSTR-1": {
   "N1": "GSTR-1_Difference_Diff_Taxable_Value",
   "O1": "GSTR-1_Difference_Diff_IGST",
   "P1": "GSTR-1_Difference_Diff_CGST",
   "Q1": "GSTR-1_Difference_Diff_SGST",
   "I1": "Turnover_GSTR-1_Taxable_Value"
  }
 }
}
```
FIELD_DISPLAY_NAMES:
```
{
 "GSTR-3B_Difference_Diff_IGST": "Diff IGST",
 "GSTR-3B_Difference_Diff_CGST": "Diff CGST",
 "GSTR-3B_Difference_Diff_SGST": "Diff SGST",
 "GSTR-1_Difference_Diff_Taxable_Value": "Diff Taxable Value",
 "GSTR-1_Difference_Diff_IGST": "Diff IGST",
 "GSTR-1_Difference_Diff_CGST": "Diff CGST",
 "GSTR-1_Difference_Diff_SGST": "Diff SGST",
 "Turnover_GSTR-1_Taxable_Value": "Taxable Value",
 "Reco_Diff_IGST_Diff": "IGST Diff",
 "Reco_Diff_CGST_Diff": "CGST Diff",
 "Reco_Diff_SGST_Diff": "SGST Diff",
 "Table_6J_Match_Table_6J_Match": "Table 6J Match",
 "Table_8D_Limits_Table_8D_Limits": "Table 8D Limits",
 "IMS_Reco_IGST": "IGST",
 "IMS_Reco_CGST": "CGST",
 "IMS_Reco_SGST": "SGST",
 "Extra_In_Purchase_IGST": "IGST",
 "Extra_In_Purchase_CGST": "CGST",
 "Extra_In_Purchase_SGST": "SGST",
 "Extra_In_GSTR2B_IGST": "IGST",
 "Extra_In_GSTR2B_CGST": "CGST",
 "Extra_In_GSTR2B_SGST": "SGST"
}
```
EXTRACTION_OPTIONS: {'skip_refresh': False, 'auto_refresh_if_missing': True, 'show_timestamp': False, 'missing_data_text': 'Missing'}

##### Which field names contain each output_column as a substring (data_consolidator.py:142-145 uses `in`)

| output sheet | output_column | field names in the SAME report type that contain it (in CELL_MAPPINGS insertion order) |
|---|---|---|
| GSTR-3B Difference | `Diff IGST` | - |
| GSTR-3B Difference | `Diff CGST` | - |
| GSTR-3B Difference | `Diff SGST` | - |
| GSTR-1 Difference | `Diff Taxable Value` | - |
| GSTR-1 Difference | `Diff IGST` | - |
| GSTR-1 Difference | `Diff CGST` | - |
| GSTR-1 Difference | `Diff SGST` | - |
| Turnover GSTR-1 | `Taxable Value` | - |
| Reco Diff | `IGST Diff` | - |
| Reco Diff | `CGST Diff` | - |
| Reco Diff | `SGST Diff` | - |
| Table 6J Match | `Table 6J Match` | - |
| Table 8D Limits | `Table 8D Limits` | - |
| IMS Reco | `IGST` | `GSTR-3B_Difference_Diff_IGST`, `Reco_Diff_IGST_Diff`, `IMS_Reco_IGST`, `Extra_In_Purchase_IGST`, `Extra_In_GSTR2B_IGST` |
| IMS Reco | `CGST` | `GSTR-3B_Difference_Diff_CGST`, `Reco_Diff_CGST_Diff`, `IMS_Reco_CGST`, `Extra_In_Purchase_CGST`, `Extra_In_GSTR2B_CGST` |
| IMS Reco | `SGST` | `GSTR-3B_Difference_Diff_SGST`, `Reco_Diff_SGST_Diff`, `IMS_Reco_SGST`, `Extra_In_Purchase_SGST`, `Extra_In_GSTR2B_SGST` |
| Extra In Purchase | `IGST` | `GSTR-3B_Difference_Diff_IGST`, `Reco_Diff_IGST_Diff`, `IMS_Reco_IGST`, `Extra_In_Purchase_IGST`, `Extra_In_GSTR2B_IGST` |
| Extra In Purchase | `CGST` | `GSTR-3B_Difference_Diff_CGST`, `Reco_Diff_CGST_Diff`, `IMS_Reco_CGST`, `Extra_In_Purchase_CGST`, `Extra_In_GSTR2B_CGST` |
| Extra In Purchase | `SGST` | `GSTR-3B_Difference_Diff_SGST`, `Reco_Diff_SGST_Diff`, `IMS_Reco_SGST`, `Extra_In_Purchase_SGST`, `Extra_In_GSTR2B_SGST` |
| Extra In GSTR2B | `IGST` | `GSTR-3B_Difference_Diff_IGST`, `Reco_Diff_IGST_Diff`, `IMS_Reco_IGST`, `Extra_In_Purchase_IGST`, `Extra_In_GSTR2B_IGST` |
| Extra In GSTR2B | `CGST` | `GSTR-3B_Difference_Diff_CGST`, `Reco_Diff_CGST_Diff`, `IMS_Reco_CGST`, `Extra_In_Purchase_CGST`, `Extra_In_GSTR2B_CGST` |
| Extra In GSTR2B | `SGST` | `GSTR-3B_Difference_Diff_SGST`, `Reco_Diff_SGST_Diff`, `IMS_Reco_SGST`, `Extra_In_Purchase_SGST`, `Extra_In_GSTR2B_SGST` |

#### Part B - the made-up workbooks (fake numbers)

- `ITC_Report_Alpha_Test_Traders_MH_211026_1030_Refreshed_211026_1100.xlsx`: sheets ['Info', 'Diff GSTR-3B', 'Reco', 'T6 of R9', 'IMS_Reco', 'ExtraInPurchase', 'ExtraInGSTR2B']
    - `Info`: {'BB2': '<Client A>'}
    - `Diff GSTR-3B`: {'J15': 1500.5, 'K15': '1,234.00', 'L15': -75}
    - `Reco`: {'C1': 10, 'C2': 20, 'C3': 30, 'C37': '=SUM(C1:C3)', 'D37': 0}
    - `T6 of R9`: {'K29': 'Yes'}
    - `IMS_Reco`: {'AA1': 111, 'AD1': 222, 'AG1': 333}
    - `ExtraInPurchase`: {'O1': 444, 'P1': 555, 'Q1': 666}
    - `ExtraInGSTR2B`: {'S1': 777, 'T1': 888, 'U1': 999}
- `Sales_Report_Alpha_Test_Traders_MH_211026_1030_Refreshed_211026_1100.xlsx`: sheets ['Diff - GSTR-1']
    - `Diff - GSTR-1`: {'I1': 50000, 'N1': 10000, 'O1': 1800, 'P1': '=N1*0.09'}
- `Sales_Report_Beta_Demo_GJ_211026_1030_Refreshed_211026_1100.xlsx`: sheets ['diff - gstr-1']
    - `diff - gstr-1`: {'I1': 50000, 'N1': 10000, 'O1': 1800, 'P1': '=N1*0.09'}
- `ITC_Report_Gamma_KA_211026_1030_Refreshed_211026_1100.xlsx`: sheets ['Info', 'Reco', 'T6 of R9', 'T8 of R9', 'IMS_Reco', 'ExtraInPurchase', 'ExtraInGSTR2B']
    - `Info`: {'BB2': '<Client A>'}
    - `Reco`: {'C1': 10, 'C2': 20, 'C3': 30, 'C37': '=SUM(C1:C3)', 'D37': 0}
    - `T6 of R9`: {'K29': 'Yes'}
    - `T8 of R9`: {'G6': 98765.4321}
    - `IMS_Reco`: {'AA1': 111, 'AD1': 222, 'AG1': 333}
    - `ExtraInPurchase`: {'O1': 444, 'P1': 555, 'Q1': 666}
    - `ExtraInGSTR2B`: {'S1': 777, 'T1': 888, 'U1': 999}

#### Part B2 - ReportProcessor._extract_data through the fake COM (data_only=True = cached values)

##### itc1 (T8 of R9 missing) -> ITC

| key in extracted_data | value (repr) |
|---|---|
| `GSTR-3B_Difference_Diff_IGST` | `1500.5` |
| `J15` | `1500.5` |
| `GSTR-3B_Difference_Diff_CGST` | `'1,234.00'` |
| `K15` | `'1,234.00'` |
| `GSTR-3B_Difference_Diff_SGST` | `-75` |
| `L15` | `-75` |
| `Reco_Diff_IGST_Diff` | `None` |
| `C37` | `None` |
| `Reco_Diff_CGST_Diff` | `0` |
| `D37` | `0` |
| `Reco_Diff_SGST_Diff` | `None` |
| `E37` | `None` |
| `Table_6J_Match_Table_6J_Match` | `'Yes'` |
| `K29` | `'Yes'` |
| `Table_8D_Limits_Table_8D_Limits` | `None` |
| `IMS_Reco_IGST` | `111` |
| `AA1` | `111` |
| `IMS_Reco_CGST` | `222` |
| `AD1` | `222` |
| `IMS_Reco_SGST` | `333` |
| `AG1` | `333` |
| `Extra_In_Purchase_IGST` | `444` |
| `O1` | `444` |
| `Extra_In_Purchase_CGST` | `555` |
| `P1` | `555` |
| `Extra_In_Purchase_SGST` | `666` |
| `Q1` | `666` |
| `Extra_In_GSTR2B_IGST` | `777` |
| `S1` | `777` |
| `Extra_In_GSTR2B_CGST` | `888` |
| `T1` | `888` |
| `Extra_In_GSTR2B_SGST` | `999` |
| `U1` | `999` |

logger warnings: ["  ⚠️ Sheet 'T8 of R9' not found"]
COM calls: Dispatch=['Excel.Application'], Visible=False, DisplayAlerts=False, opened=['ITC_Report_Alpha_Test_Traders_MH_211026_1030_Refreshed_211026_1100.xlsx'], Quit=True, pythoncom=['CoInitialize', 'CoUninitialize']

##### sales1 -> Sales

| key in extracted_data | value (repr) |
|---|---|
| `GSTR-1_Difference_Diff_Taxable_Value` | `10000` |
| `N1` | `10000` |
| `GSTR-1_Difference_Diff_IGST` | `1800` |
| `O1` | `1800` |
| `GSTR-1_Difference_Diff_CGST` | `None` |
| `P1` | `None` |
| `GSTR-1_Difference_Diff_SGST` | `None` |
| `Q1` | `None` |
| `Turnover_GSTR-1_Taxable_Value` | `50000` |
| `I1` | `50000` |

logger warnings: []
COM calls: Dispatch=['Excel.Application'], Visible=False, DisplayAlerts=False, opened=['Sales_Report_Alpha_Test_Traders_MH_211026_1030_Refreshed_211026_1100.xlsx'], Quit=True, pythoncom=['CoInitialize', 'CoUninitialize']

##### sales_case (sheet named diff - gstr-1) -> Sales

| key in extracted_data | value (repr) |
|---|---|
| `GSTR-1_Difference_Diff_Taxable_Value` | `None` |
| `GSTR-1_Difference_Diff_IGST` | `None` |
| `GSTR-1_Difference_Diff_CGST` | `None` |
| `GSTR-1_Difference_Diff_SGST` | `None` |
| `Turnover_GSTR-1_Taxable_Value` | `None` |

logger warnings: ["  ⚠️ Sheet 'Diff - GSTR-1' not found"]
COM calls: Dispatch=['Excel.Application'], Visible=False, DisplayAlerts=False, opened=['Sales_Report_Beta_Demo_GJ_211026_1030_Refreshed_211026_1100.xlsx'], Quit=True, pythoncom=['CoInitialize', 'CoUninitialize']

##### itc_nulls (Diff GSTR-3B missing) -> ITC

| key in extracted_data | value (repr) |
|---|---|
| `GSTR-3B_Difference_Diff_IGST` | `None` |
| `GSTR-3B_Difference_Diff_CGST` | `None` |
| `GSTR-3B_Difference_Diff_SGST` | `None` |
| `Reco_Diff_IGST_Diff` | `None` |
| `C37` | `None` |
| `Reco_Diff_CGST_Diff` | `0` |
| `D37` | `0` |
| `Reco_Diff_SGST_Diff` | `None` |
| `E37` | `None` |
| `Table_6J_Match_Table_6J_Match` | `'Yes'` |
| `K29` | `'Yes'` |
| `Table_8D_Limits_Table_8D_Limits` | `98765.4321` |
| `G6` | `98765.4321` |
| `IMS_Reco_IGST` | `111` |
| `AA1` | `111` |
| `IMS_Reco_CGST` | `222` |
| `AD1` | `222` |
| `IMS_Reco_SGST` | `333` |
| `AG1` | `333` |
| `Extra_In_Purchase_IGST` | `444` |
| `O1` | `444` |
| `Extra_In_Purchase_CGST` | `555` |
| `P1` | `555` |
| `Extra_In_Purchase_SGST` | `666` |
| `Q1` | `666` |
| `Extra_In_GSTR2B_IGST` | `777` |
| `S1` | `777` |
| `Extra_In_GSTR2B_CGST` | `888` |
| `T1` | `888` |
| `Extra_In_GSTR2B_SGST` | `999` |
| `U1` | `999` |

logger warnings: ["  ⚠️ Sheet 'Diff GSTR-3B' not found"]
COM calls: Dispatch=['Excel.Application'], Visible=False, DisplayAlerts=False, opened=['ITC_Report_Gamma_KA_211026_1030_Refreshed_211026_1100.xlsx'], Quit=True, pythoncom=['CoInitialize', 'CoUninitialize']

##### Same ITC file read with data_only=False (what openpyxl sees as the cell CONTENT)

{'Reco_Diff_IGST_Diff': '=SUM(C1:C3)', 'C37': '=SUM(C1:C3)'}

#### Part C - process_client on a machine without pywin32 (report_processor.py:66-77) returns:

```
{'client': 'Delta Mock Co-DL', 'timestamp': '<datetime>', 'itc': {'status': {'success': False, 'error': 'pywin32 is not installed for this Python, so Power Query cannot be refreshed. Install it with: pip install pywin32'}, 'data': {}, 'extracted_values': {}}, 'sales': {'status': {'success': False, 'error': 'pywin32 is not installed for this Python, so Power Query cannot be refreshed. Install it with: pip install pywin32'}, 'data': {}, 'extracted_values': {}}}
```

#### Part D - DataConsolidator.create_report(results, folder) wrote: `PQ_Extraction_Report_20261021_0420.xlsx` into the folder it was given

logger output during create_report: []

Sheets in order: ['Summary', 'GSTR-3B Difference', 'GSTR-1 Difference', 'Turnover GSTR-1', 'Reco Diff', 'Table 6J Match', 'Table 8D Limits', 'IMS Reco', 'Extra In Purchase', 'Extra In GSTR2B']

##### Sheet `Summary` (dimensions A1:A17; merged [])

| row | A |
|---|---|
| 1 | 'Extraction Summary' |
| 2 |  |
| 3 | 'Generated: 2026-10-21 04:20:58' |
| 4 | 'Total Clients: 4' |
| 5 | 'ITC Reports Processed: 2/4' |
| 6 | 'Sales Reports Processed: 2/4' |
| 7 |  |
| 8 | 'Sheets Created:' |
| 9 | '• GSTR-3B Difference' |
| 10 | '• GSTR-1 Difference' |
| 11 | '• Turnover GSTR-1' |
| 12 | '• Reco Diff' |
| 13 | '• Table 6J Match' |
| 14 | '• Table 8D Limits' |
| 15 | '• IMS Reco' |
| 16 | '• Extra In Purchase' |
| 17 | '• Extra In GSTR2B' |

##### Sheet `GSTR-3B Difference` (dimensions A1:E7; merged ['A1:C1'])

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'GSTR-3B Difference Report' |  |  |  |  |
| 2 |  |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Diff IGST' | 'Diff CGST' | 'Diff SGST' |
| 4 | 'Alpha Test Traders-MH' | '✓' | 1500.5 [fmt #,##0.00] | '1,234.00' | -75 [fmt #,##0.00] |
| 5 | 'Beta Demo Industries Pvt Ltd-GJ' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 6 | 'Gamma Sample Enterprises LLP-KA' | '✓' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Mock Co-DL' | '✗' | 'Missing' | 'Missing' | 'Missing' |

##### Sheet `GSTR-1 Difference` (dimensions A1:F7; merged ['A1:C1'])

| row | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| 1 | 'GSTR-1 Difference Report' |  |  |  |  |  |
| 2 |  |  |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Diff Taxable Value' | 'Diff IGST' | 'Diff CGST' | 'Diff SGST' |
| 4 | 'Alpha Test Traders-MH' | '✓' | 10000 [fmt #,##0.00] | 1800 [fmt #,##0.00] | 'Missing' | 'Missing' |
| 5 | 'Beta Demo Industries Pvt Ltd-GJ' | '✓' | 'Missing' | 'Missing' | 'Missing' | 'Missing' |
| 6 | 'Gamma Sample Enterprises LLP-KA' | '✗' | 'Missing' | 'Missing' | 'Missing' | 'Missing' |
| 7 | 'Delta Mock Co-DL' | '✗' | 'Missing' | 'Missing' | 'Missing' | 'Missing' |

##### Sheet `Turnover GSTR-1` (dimensions A1:C7; merged ['A1:C1'])

| row | A | B | C |
|---|---|---|---|
| 1 | 'Turnover GSTR-1 Report' |  |  |
| 2 |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Taxable Value' |
| 4 | 'Alpha Test Traders-MH' | '✓' | 50000 [fmt #,##0.00] |
| 5 | 'Beta Demo Industries Pvt Ltd-GJ' | '✓' | 'Missing' |
| 6 | 'Gamma Sample Enterprises LLP-KA' | '✗' | 'Missing' |
| 7 | 'Delta Mock Co-DL' | '✗' | 'Missing' |

##### Sheet `Reco Diff` (dimensions A1:E7; merged ['A1:C1'])

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'Reco Diff Report' |  |  |  |  |
| 2 |  |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'IGST Diff' | 'CGST Diff' | 'SGST Diff' |
| 4 | 'Alpha Test Traders-MH' | '✓' | 'Missing' | 0 | 'Missing' |
| 5 | 'Beta Demo Industries Pvt Ltd-GJ' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 6 | 'Gamma Sample Enterprises LLP-KA' | '✓' | 'Missing' | 0 | 'Missing' |
| 7 | 'Delta Mock Co-DL' | '✗' | 'Missing' | 'Missing' | 'Missing' |

##### Sheet `Table 6J Match` (dimensions A1:C7; merged ['A1:C1'])

| row | A | B | C |
|---|---|---|---|
| 1 | 'Table 6J Match Report' |  |  |
| 2 |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Table 6J Match' |
| 4 | 'Alpha Test Traders-MH' | '✓' | 'Yes' |
| 5 | 'Beta Demo Industries Pvt Ltd-GJ' | '✗' | 'Missing' |
| 6 | 'Gamma Sample Enterprises LLP-KA' | '✓' | 'Yes' |
| 7 | 'Delta Mock Co-DL' | '✗' | 'Missing' |

##### Sheet `Table 8D Limits` (dimensions A1:C7; merged ['A1:C1'])

| row | A | B | C |
|---|---|---|---|
| 1 | 'Table 8D Limits Report' |  |  |
| 2 |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'Table 8D Limits' |
| 4 | 'Alpha Test Traders-MH' | '✓' | 'Missing' |
| 5 | 'Beta Demo Industries Pvt Ltd-GJ' | '✗' | 'Missing' |
| 6 | 'Gamma Sample Enterprises LLP-KA' | '✓' | 98765.4321 [fmt #,##0.00] |
| 7 | 'Delta Mock Co-DL' | '✗' | 'Missing' |

##### Sheet `IMS Reco` (dimensions A1:E7; merged ['A1:C1'])

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'IMS Reco Report' |  |  |  |  |
| 2 |  |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'IGST' | 'CGST' | 'SGST' |
| 4 | 'Alpha Test Traders-MH' | '✓' | 1500.5 [fmt #,##0.00] | '1,234.00' | -75 [fmt #,##0.00] |
| 5 | 'Beta Demo Industries Pvt Ltd-GJ' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 6 | 'Gamma Sample Enterprises LLP-KA' | '✓' | 111 [fmt #,##0.00] | 222 [fmt #,##0.00] | 333 [fmt #,##0.00] |
| 7 | 'Delta Mock Co-DL' | '✗' | 'Missing' | 'Missing' | 'Missing' |

##### Sheet `Extra In Purchase` (dimensions A1:E7; merged ['A1:C1'])

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'Extra In Purchase Report' |  |  |  |  |
| 2 |  |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'IGST' | 'CGST' | 'SGST' |
| 4 | 'Alpha Test Traders-MH' | '✓' | 1500.5 [fmt #,##0.00] | '1,234.00' | -75 [fmt #,##0.00] |
| 5 | 'Beta Demo Industries Pvt Ltd-GJ' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 6 | 'Gamma Sample Enterprises LLP-KA' | '✓' | 444 [fmt #,##0.00] | 555 [fmt #,##0.00] | 666 [fmt #,##0.00] |
| 7 | 'Delta Mock Co-DL' | '✗' | 'Missing' | 'Missing' | 'Missing' |

##### Sheet `Extra In GSTR2B` (dimensions A1:E7; merged ['A1:C1'])

| row | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | 'Extra In GSTR2B Report' |  |  |  |  |
| 2 |  |  |  |  |  |
| 3 | 'Client Name' | 'Successfully Run' | 'IGST' | 'CGST' | 'SGST' |
| 4 | 'Alpha Test Traders-MH' | '✓' | 1500.5 [fmt #,##0.00] | '1,234.00' | -75 [fmt #,##0.00] |
| 5 | 'Beta Demo Industries Pvt Ltd-GJ' | '✗' | 'Missing' | 'Missing' | 'Missing' |
| 6 | 'Gamma Sample Enterprises LLP-KA' | '✓' | 777 [fmt #,##0.00] | 888 [fmt #,##0.00] | 999 [fmt #,##0.00] |
| 7 | 'Delta Mock Co-DL' | '✗' | 'Missing' | 'Missing' | 'Missing' |

#### Part E - configured cell vs value written, client 'Alpha Test Traders-MH' (all statuses success)

| output sheet | column | configured source | value in the source workbook (extracted) | value written by consolidator | same? |
|---|---|---|---|---|---|
| GSTR-3B Difference | Diff IGST | `Diff GSTR-3B`!J15 | `1500.5` | `1500.5` | yes |
| GSTR-3B Difference | Diff CGST | `Diff GSTR-3B`!K15 | `'1,234.00'` | `'1,234.00'` | yes |
| GSTR-3B Difference | Diff SGST | `Diff GSTR-3B`!L15 | `-75` | `-75` | yes |
| GSTR-1 Difference | Diff Taxable Value | `Diff - GSTR-1`!N1 | `10000` | `10000` | yes |
| GSTR-1 Difference | Diff IGST | `Diff - GSTR-1`!O1 | `1800` | `1800` | yes |
| GSTR-1 Difference | Diff CGST | `Diff - GSTR-1`!P1 | `None` | `'Missing'` | yes |
| GSTR-1 Difference | Diff SGST | `Diff - GSTR-1`!Q1 | `None` | `'Missing'` | yes |
| Turnover GSTR-1 | Taxable Value | `Diff - GSTR-1`!I1 | `50000` | `50000` | yes |
| Reco Diff | IGST Diff | `Reco`!C37 | `None` | `'Missing'` | yes |
| Reco Diff | CGST Diff | `Reco`!D37 | `0` | `0` | yes |
| Reco Diff | SGST Diff | `Reco`!E37 | `None` | `'Missing'` | yes |
| Table 6J Match | Table 6J Match | `T6 of R9`!K29 | `'Yes'` | `'Yes'` | yes |
| Table 8D Limits | Table 8D Limits | `T8 of R9`!G6 | `None` | `'Missing'` | yes |
| IMS Reco | IGST | `IMS_Reco`!AA1 | `111` | `1500.5` | **NO** |
| IMS Reco | CGST | `IMS_Reco`!AD1 | `222` | `'1,234.00'` | **NO** |
| IMS Reco | SGST | `IMS_Reco`!AG1 | `333` | `-75` | **NO** |
| Extra In Purchase | IGST | `ExtraInPurchase`!O1 | `444` | `1500.5` | **NO** |
| Extra In Purchase | CGST | `ExtraInPurchase`!P1 | `555` | `'1,234.00'` | **NO** |
| Extra In Purchase | SGST | `ExtraInPurchase`!Q1 | `666` | `-75` | **NO** |
| Extra In GSTR2B | IGST | `ExtraInGSTR2B`!S1 | `777` | `1500.5` | **NO** |
| Extra In GSTR2B | CGST | `ExtraInGSTR2B`!T1 | `888` | `'1,234.00'` | **NO** |
| Extra In GSTR2B | SGST | `ExtraInGSTR2B`!U1 | `999` | `-75` | **NO** |

#### Part F - DataConsolidator._calculate_value (data_consolidator.py:188-218), called directly

- {'type': 'subtract', 'cell1': 'J15', 'cell2': 'L15'} -> 1575.5  logs=[]
- {'type': 'add', 'cell1': 'J15', 'cell2': 'X'} -> 1513.0  logs=[]
- {'type': 'multiply', 'cell1': 'J15', 'cell2': 'D37'} -> 0.0  logs=[]
- {'type': 'divide', 'cell1': 'J15', 'cell2': 'D37'} -> 'Div/0'  logs=[]
- {'type': 'divide', 'cell1': 'J15', 'cell2': 'C37'} -> 'Div/0'  logs=[]
- {'type': 'subtract', 'cell1': 'J15', 'cell2': 'K15'} -> 'Error'  logs=[]
- {'type': 'subtract', 'cell1': 'J15', 'cell2': 'Z99'} -> 1500.5  logs=[]
- {'type': 'power', 'cell1': 'J15', 'cell2': 'L15'} -> 'Unknown'  logs=[]
- {'cell1': 'J15', 'cell2': 'L15'} -> 'Unknown'  logs=[]
- calculation=None-like dict {} is falsy so the mapping never reaches _calculate_value (data_consolidator.py:130 `elif mapping.get('calculation')`)

<!-- end of outputs/t4_cells.md -->

### T4b - the consolidator branches the shipped configuration never reaches (RAN)

Script `scripts/t4b_config_variants.py` re-executes the module's own builder (`cell_mappings.py:260-278`) on a made-up configuration with (i) a `description` list, (ii) a `calculation`, (iii) a calculation that refers to a cell nobody maps, and (iv) two output sheets that read the SAME input cell (`Diff GSTR-3B`!J15). Output `outputs/t4b_config_variants.md`, reproduced exactly:



<!-- RAN output: outputs/t4b_config_variants.md (headings demoted by two levels) -->

Builder source executed (cell_mappings.py:260-278):
```
CELL_MAPPINGS = {}
FIELD_DISPLAY_NAMES = {}

# Build old format for compatibility
for sheet_name, config in EXTRACTION_CONFIG.items():
    report_type = config['report_type']
    if report_type not in CELL_MAPPINGS:
        CELL_MAPPINGS[report_type] = {}
    
    for mapping in config['mappings']:
        if mapping['input_sheet'] and mapping['input_cell']:
            sheet = mapping['input_sheet']
            if sheet not in CELL_MAPPINGS[report_type]:
                CELL_MAPPINGS[report_type][sheet] = {}
            
            # Create unique field name
            field_name = f"{sheet_name}_{mapping['output_column']}".replace(' ', '_')
            CELL_MAPPINGS[report_type][sheet][mapping['input_cell']] = field_name
            FIELD_DISPLAY_NAMES[field_name] = mapping['output_column']
```

CELL_MAPPINGS built from CUSTOM: {'ITC': {'Diff GSTR-3B': {'J15': 'Sheet_Two_Value_B', 'L15': 'Sheet_Two_Value_C'}}}
FIELD_DISPLAY_NAMES built from CUSTOM: {'Sheet_One_Value_A': 'Value A', 'Sheet_Two_Value_B': 'Value B', 'Sheet_Two_Value_C': 'Value C'}

_extract_data -> {'Sheet_Two_Value_B': 1500.5, 'J15': 1500.5, 'Sheet_Two_Value_C': -75, 'L15': -75}

##### `Summary` merged=[]
row 1: 'Extraction Summary'
row 2: None
row 3: 'Generated: 2026-10-21 04:21:30'
row 4: 'Total Clients: 1'
row 5: 'ITC Reports Processed: 1/1'
row 6: 'Sales Reports Processed: 1/1'
row 7: None
row 8: 'Sheets Created:'
row 9: '• Sheet One'
row 10: '• Sheet Two'

##### `Sheet One` merged=['A1:J1', 'A2:J2']
row 1: 'Custom title line' | None | None | None | None | None | None | None | None | None
row 2: 'second line of description' | None | None | None | None | None | None | None | None | None
row 3: None | None | None | None | None | None | None | None | None | None
row 4: 'Client Name' | 'Successfully Run' | 'Value A' | None | None | None | None | None | None | None
row 5: 'Alpha Test Traders-MH' | '✓' | 'Missing' | None | None | None | None | None | None | None

##### `Sheet Two` merged=['A1:C1']
row 1: 'Sheet Two Report' | None | None | None | None | None
row 2: None | None | None | None | None | None
row 3: 'Client Name' | 'Successfully Run' | 'Value B' | 'Value C' | 'J15 minus L15' | 'J15 minus unmapped Z9'
row 4: 'Alpha Test Traders-MH' | '✓' | 1500.5 | -75 | 1575.5 | 1500.5

<!-- end of outputs/t4b_config_variants.md -->

### T4 observations

1. **The consolidated report puts the GSTR-3B Difference figures into the IGST/CGST/SGST columns of three other sheets.** RAN (Part E, rows marked **NO**): for client 1, `IMS Reco` shows 1500.5 / '1,234.00' / -75 (the values of `Diff GSTR-3B`!J15/K15/L15) instead of 111/222/333 (`IMS_Reco`!AA1/AD1/AG1); `Extra In Purchase` shows the same instead of 444/555/666; `Extra In GSTR2B` the same instead of 777/888/999. Cause (READ, HIGH): the lookup at `data_consolidator.py:142-145` takes the first extracted key that *contains* the output column text; the columns of those three sheets are just `IGST`, `CGST`, `SGST`, and the first key containing `IGST` is `GSTR-3B_Difference_Diff_IGST` (it is inserted first because `GSTR-3B Difference` is the first config entry). Columns whose names contain a space (`Diff IGST`, `IGST Diff`, `Table 6J Match`...) never match any key (keys have underscores), so they fall through to the exact lookup and are right. When the `Diff GSTR-3B` sheet is absent (client 3), the first match is `None`, so the exact lookup runs and the right values appear (111/222/333, 444..., 777...). So the wrong figures appear exactly in the normal case. The substring loop dates from the initial commit edac6a3; the three affected output sheets were added in 47a1492 (30-10-2025, the last commit shared with project B), so the wrong output exists since 47a1492. This is a technical finding: the output does not follow the code's own configuration. Which cells *should* be read is [GST RULE] for the author. No test covers the consolidator (RAN grep of `tests/`: no reference to `DataConsolidator`, `create_report` or `_extract_data`).
2. **A failed refresh blanks every figure even when the data was extracted**: client 2's ITC status is failure, its extracted data is complete, yet every ITC column shows `Missing` and the status ✗ (`data_consolidator.py:128-129`). Meanwhile the extractor's log says "(Data extracted despite error)" (`extractor_window.py:924-925`). RAN, HIGH.
3. **Sheet names must match exactly, including case**: `diff - gstr-1` was "not found" (`report_processor.py:565`), every Sales value became `Missing`, but the row still shows ✓ because ✓ only reflects the refresh status. RAN (through the fake COM), HIGH for the comparison itself.
4. **Text in a number cell is copied as text** (`'1,234.00'` lands as text, no number format, so Excel would not add it up). **Zero** is written as 0 without the `#,##0.00` format (`data_consolidator.py:161`). **Empty cell** -> `Missing`. **Missing sheet** -> `Missing` for all its columns (plus a log warning). **Formula without cached value** -> `None` -> `Missing` through openpyxl; through real Excel the computed value would be read (not verifiable here). RAN.
5. The `Summary` sheet counts "ITC Reports Processed: 2/4" from refresh statuses only (`data_consolidator.py:235-239`), not from what was extracted. RAN.
6. The header row is written twice: once at 77-80 with only two headers and default formatting, then again at 92-99 with all headers and formatting. The first write is fully overwritten; harmless leftover. READ, HIGH.
7. T4b: `description` lines go to A1, A2... merged across A:J, header row moves to `len(description)+2` (RAN). A calculation works on the bare cell-reference keys (`J15`, `L15`) that `_extract_data` also stores - so it only works for cells that some mapping already reads, and a reference to an unmapped cell silently counts as 0 (`J15 minus unmapped Z9` = 1500.5; `data_consolidator.py:194-195`). Cell references are not sheet-qualified, so two sheets of one report type that read the same address would overwrite each other's cell-reference key (READ, HIGH; not triggered by the shipped config, which uses distinct addresses within each report type - RAN Part A).
8. T4b: when two output sheets read the same input cell, the builder keeps only the LAST field name for that cell (`cell_mappings.py:277`), so the first output sheet always shows `Missing` (`Sheet One`/`Value A`). Not triggered by the shipped config (RAN Part A: no duplicate cell within a report type).
9. `_calculate_value` (RAN Part F): `subtract/add/multiply/divide`; text that is not a plain number -> `'Error'`; divide by 0 or by an empty cell -> `'Div/0'`; unknown or missing type -> `'Unknown'`; empty/None/0 inputs count as 0 (`if value1` at 199-200). No shipped mapping uses a calculation (all `None`), so this code is reachable only by editing the config.
10. Dead code (READ, HIGH; grep found no caller): `_write_section` (250-352), `_create_details_sheet` (354-414). `EXTRACTION_OPTIONS['auto_refresh_if_missing']` and `['show_timestamp']` (`cell_mappings.py:250-251`) are never read; the "refresh if no refreshed file exists" behaviour is hard-coded (`report_processor.py:156-159`). `EXTRACTION_OPTIONS['skip_refresh']` is changed at run time by the extractor's checkbox, which mutates the module-level dictionary (`extractor_window.py:486-487`).
11. The consolidated report's name has minute resolution (`%Y%m%d_%H%M`); two extractions in the same minute into the same folder overwrite the earlier report without warning (`wb.save`, `data_consolidator.py:179`). READ, HIGH.

## 7. T5 - Refresh naming

**The rules (READ, `power_query_extractor/core/refresh_naming.py`):**

- Default suffix pattern `_Refreshed_{timestamp}` (line 14); legacy marker `_refreshed_` (20).
- `refresh_markers(pattern)` (23-40): lower-case the pattern, split it on `{timestamp}`, keep the non-empty pieces; if the pattern is empty/None or has no fixed text, use `['_refreshed_']`.
- `is_refreshed_name(name, pattern)` (43-55): the lower-cased name must end with `.xlsx`; then it is "refreshed" if it contains ALL the markers, OR if it contains the legacy marker `_refreshed_` anyway.
- `is_refreshed_copy_of(name, stem, pattern)` (58-66): the name must start with the original's stem (any case) and pass `is_refreshed_name`.

**Where they are used (READ):** the client list's "last refreshed" column (`power_query_extractor/gui/extractor_window.py:45-82`: newest modification time among `itc_report_*` / `sales_report_*` files that pass `is_refreshed_name`) and the "skip refresh" option (`report_processor.py:137-159`: reuse the newest existing refreshed copy of the original).

**How a refreshed copy is named and made (READ, `report_processor.py:108-196`):** `glob("<ITC|Sales>_Report_*.xlsx")` in the client's newest Version folder (112-113) -> the FIRST hit is taken as the original (125) -> name = `<original stem><pattern with {timestamp} replaced by DDMMYY_HHMM>.xlsx` (128-130; plain `str.replace`, so only the exact lower-case `{timestamp}` is replaced) -> `shutil.copy2(original, copy)` (164) -> refresh the copy through Excel (168) -> wait 3 s (185) -> extract from the copy (188).

**How I ran it:** the pure functions directly; `_process_report` for real on fake folders under `work/t5/`, with only `_refresh_power_query_simple` and `_extract_data` replaced by stubs that record which file they were given (no Excel here) and `time.sleep` disabled. Script `scripts/t5_refresh_naming.py`, output `outputs/t5_refresh_naming.md`, reproduced exactly (times in names such as `_211026_0422` are the real clock of the run):



<!-- RAN output: outputs/t5_refresh_naming.md (headings demoted by two levels) -->

#### refresh_markers(pattern)  (refresh_naming.py:23-40)

| pattern | markers | the name report_processor.py:128-130 would build for stem `ITC_Report_X` at 21-10-2026 11:00 |
|---|---|---|
| `None` | `['_refreshed_']` | `(processor default pattern is never None: report_processor.py:51)` |
| `''` | `['_refreshed_']` | `ITC_Report_X.xlsx` |
| `'_Refreshed_{timestamp}'` | `['_refreshed_']` | `ITC_Report_X_Refreshed_211026_1100.xlsx` |
| `'_Updated_{timestamp}'` | `['_updated_']` | `ITC_Report_X_Updated_211026_1100.xlsx` |
| `'_v{timestamp}_final'` | `['_v', '_final']` | `ITC_Report_X_v211026_1100_final.xlsx` |
| `'{timestamp}'` | `['_refreshed_']` | `ITC_Report_X211026_1100.xlsx` |
| `'_Refreshed'` | `['_refreshed']` | `ITC_Report_X_Refreshed.xlsx` |
| `'_R_{timestamp}_{timestamp}'` | `['_r_', '_']` | `ITC_Report_X_R_211026_1100_211026_1100.xlsx` |
| `' Refreshed {timestamp}'` | `[' refreshed ']` | `ITC_Report_X Refreshed 211026_1100.xlsx` |
| `'_REFRESHED_{TIMESTAMP}'` | `['_refreshed_']` | `ITC_Report_X_REFRESHED_{TIMESTAMP}.xlsx` |
| `'_v{timestamp}'` | `['_v']` | `ITC_Report_X_v211026_1100.xlsx` |

#### is_refreshed_name(name, pattern)  (refresh_naming.py:43-55)

| file name | None | '_Refreshed_{timestamp}' | '_Updated_{timestamp}' | '_v{timestamp}_final' | '{timestamp}' | '_v{timestamp}' |
|---|---|---|---|---|---|---|
| `ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx` | False | False | False | False | False | False |
| `ITC_Report_Alpha_Test_Traders_MH_211026_1030_Refreshed_211026_1100.xlsx` | True | True | True | True | True | True |
| `ITC_Report_Alpha_Test_Traders_MH_211026_1030_Refreshed_211026_1100_Refreshed_211026_1130.xlsx` | True | True | True | True | True | True |
| `ITC_REPORT_ALPHA_MH_211026_1030_REFRESHED_211026_1100.XLSX` | True | True | True | True | True | True |
| `ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1100.xlsm` | False | False | False | False | False | False |
| `ITC_Report_Alpha_MH_211026_1030_Updated_211026_1100.xlsx` | False | False | True | False | False | False |
| `ITC_Report_Alpha_MH_211026_1030_v211026_1100_final.xlsx` | False | False | False | True | False | True |
| `ITC_Report_Alpha_MH_211026_1030_v211026_1100.xlsx` | False | False | False | False | False | True |
| `ITC_Report_Alpha_MH_211026_1030211026_1100.xlsx` | False | False | False | False | False | False |
| `ITC_Report_Alpha_MH_211026_1030_Refreshed.xlsx` | False | False | False | False | False | False |
| `ITC_Report_Alpha_MH_211026_1030 Refreshed 211026_1100.xlsx` | False | False | False | False | False | False |
| `ITC_Report_Alpha_MH_211026_1030_REFRESHED_{TIMESTAMP}.xlsx` | True | True | True | True | True | True |
| `ITC_Report_Refreshed_Test_Traders_MH_211026_1030.xlsx` | True | True | True | True | True | True |
| `ITC_Report_Vega_Test_Traders_MH_211026_1030.xlsx` | False | False | False | False | False | True |
| `~$ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1100.xlsx` | True | True | True | True | True | True |

#### is_refreshed_copy_of(name, stem, pattern)  (refresh_naming.py:58-66)

- name `ITC_Report_Alpha_Test_Traders_MH_211026_1030_Refreshed_211026_1100.xlsx`, stem `ITC_Report_Alpha_Test_Traders_MH_211026_1030`, pattern `_Refreshed_{timestamp}` -> True
- name `ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx`, stem `ITC_Report_Alpha_Test_Traders_MH_211026_1030`, pattern `_Refreshed_{timestamp}` -> False
- name `ITC_Report_Alpha_Test_Traders_MH_211026_1030_Refreshed_211026_1100_Refreshed_211026_1130.xlsx`, stem `ITC_Report_Alpha_Test_Traders_MH_211026_1030`, pattern `_Refreshed_{timestamp}` -> True
- name `itc_report_alpha_test_traders_mh_211026_1030_refreshed_x.xlsx`, stem `ITC_Report_Alpha_Test_Traders_MH_211026_1030`, pattern `_Refreshed_{timestamp}` -> True
- name `ITC_Report_Alpha_Test_Traders_MH_211026_1030_Refreshed_211026_1100.xlsx`, stem `ITC_Report_Alpha_Test_Traders_MH_211026_1030`, pattern `_Updated_{timestamp}` -> True
- name `ITC_Report_Alpha_Test_Traders_MH_211026_1030_Updated_211026_1100.xlsx`, stem `ITC_Report_Alpha_Test_Traders_MH_211026_1030`, pattern `_Refreshed_{timestamp}` -> False
- name `Sales_Report_Alpha_Test_Traders_MH_211026_1030_Refreshed_211026_1100.xlsx`, stem `ITC_Report_Alpha_Test_Traders_MH_211026_1030`, pattern `_Refreshed_{timestamp}` -> False
- name `ITC_Report_Refreshed_Test_Traders_MH_211026_1030.xlsx`, stem `ITC_Report_Refreshed_Test_Traders_MH_211026_1030`, pattern `_Refreshed_{timestamp}` -> True

#### ReportProcessor._process_report naming, run for real on fake folders (refresh + extraction stubbed)

##### only the original

- suffix `'_Refreshed_{timestamp}'`, skip_refresh=False
- files before: ['ITC_Report_Alpha_MH_211026_1030.xlsx']
- glob('ITC_Report_*.xlsx') order on this disk: ['ITC_Report_Alpha_MH_211026_1030.xlsx']  -> report_files[0] = `ITC_Report_Alpha_MH_211026_1030.xlsx`
- result status: {'success': True, 'error': None, 'validation_message': 'stub'}
- original_file: `ITC_Report_Alpha_MH_211026_1030.xlsx`  refreshed_file: `ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_0422.xlsx`
- stub calls: [('refresh', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_0422.xlsx'), ('extract', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_0422.xlsx')]
- GUI log lines: []
- logger errors: []
- files after: ['ITC_Report_Alpha_MH_211026_1030.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_0422.xlsx']

##### original plus an older refreshed copy

- suffix `'_Refreshed_{timestamp}'`, skip_refresh=False
- files before: ['ITC_Report_Alpha_MH_211026_1030.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx']
- glob('ITC_Report_*.xlsx') order on this disk: ['ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx', 'ITC_Report_Alpha_MH_211026_1030.xlsx']  -> report_files[0] = `ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx`
- result status: {'success': True, 'error': None, 'validation_message': 'stub'}
- original_file: `ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx`  refreshed_file: `ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040_Refreshed_211026_0422.xlsx`
- stub calls: [('refresh', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040_Refreshed_211026_0422.xlsx'), ('extract', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040_Refreshed_211026_0422.xlsx')]
- GUI log lines: []
- logger errors: []
- files after: ['ITC_Report_Alpha_MH_211026_1030.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040_Refreshed_211026_0422.xlsx']

##### original plus refreshed copy, names chosen so the copy lists first

- suffix `'_Refreshed_{timestamp}'`, skip_refresh=False
- files before: ['ITC_Report_Alpha_MH_211026_1030.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx']
- glob('ITC_Report_*.xlsx') order on this disk: ['ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx', 'ITC_Report_Alpha_MH_211026_1030.xlsx']  -> report_files[0] = `ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx`
- result status: {'success': True, 'error': None, 'validation_message': 'stub'}
- original_file: `ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx`  refreshed_file: `ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040_Refreshed_211026_0422.xlsx`
- stub calls: [('refresh', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040_Refreshed_211026_0422.xlsx'), ('extract', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040_Refreshed_211026_0422.xlsx')]
- GUI log lines: []
- logger errors: []
- files after: ['ITC_Report_Alpha_MH_211026_1030.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040_Refreshed_211026_0422.xlsx']

##### skip_refresh with two refreshed copies

- suffix `'_Refreshed_{timestamp}'`, skip_refresh=True
- files before: ['ITC_Report_Alpha_MH_211026_1030.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1050.xlsx']
- glob('ITC_Report_*.xlsx') order on this disk: ['ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1050.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx', 'ITC_Report_Alpha_MH_211026_1030.xlsx']  -> report_files[0] = `ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1050.xlsx`
- result status: {'success': True, 'error': None, 'validation_message': 'Skipped - Using existing refreshed file'}
- original_file: `ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1050.xlsx`  refreshed_file: `ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1050.xlsx`
- stub calls: [('extract', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1050.xlsx')]
- GUI log lines: ['Skip refresh enabled - using existing: ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1050.xlsx']
- logger errors: []
- files after: ['ITC_Report_Alpha_MH_211026_1030.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1040.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1050.xlsx']

##### skip_refresh with no refreshed copy

- suffix `'_Refreshed_{timestamp}'`, skip_refresh=True
- files before: ['ITC_Report_Alpha_MH_211026_1030.xlsx']
- glob('ITC_Report_*.xlsx') order on this disk: ['ITC_Report_Alpha_MH_211026_1030.xlsx']  -> report_files[0] = `ITC_Report_Alpha_MH_211026_1030.xlsx`
- result status: {'success': True, 'error': None, 'validation_message': 'stub'}
- original_file: `ITC_Report_Alpha_MH_211026_1030.xlsx`  refreshed_file: `ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_0422.xlsx`
- stub calls: [('refresh', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_0422.xlsx'), ('extract', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_0422.xlsx')]
- GUI log lines: ['No existing refreshed file found - will refresh']
- logger errors: []
- files after: ['ITC_Report_Alpha_MH_211026_1030.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_0422.xlsx']

##### skip_refresh, client name contains Refreshed

- suffix `'_Refreshed_{timestamp}'`, skip_refresh=True
- files before: ['ITC_Report_Refreshed_Test_Traders_MH_211026_1030.xlsx']
- glob('ITC_Report_*.xlsx') order on this disk: ['ITC_Report_Refreshed_Test_Traders_MH_211026_1030.xlsx']  -> report_files[0] = `ITC_Report_Refreshed_Test_Traders_MH_211026_1030.xlsx`
- result status: {'success': True, 'error': None, 'validation_message': 'Skipped - Using existing refreshed file'}
- original_file: `ITC_Report_Refreshed_Test_Traders_MH_211026_1030.xlsx`  refreshed_file: `ITC_Report_Refreshed_Test_Traders_MH_211026_1030.xlsx`
- stub calls: [('extract', 'ITC_Report_Refreshed_Test_Traders_MH_211026_1030.xlsx')]
- GUI log lines: ['Skip refresh enabled - using existing: ITC_Report_Refreshed_Test_Traders_MH_211026_1030.xlsx']
- logger errors: []
- files after: ['ITC_Report_Refreshed_Test_Traders_MH_211026_1030.xlsx']

##### skip_refresh, custom suffix _v{timestamp}, client starting with V

- suffix `'_v{timestamp}'`, skip_refresh=True
- files before: ['ITC_Report_Vega_Test_Traders_MH_211026_1030.xlsx']
- glob('ITC_Report_*.xlsx') order on this disk: ['ITC_Report_Vega_Test_Traders_MH_211026_1030.xlsx']  -> report_files[0] = `ITC_Report_Vega_Test_Traders_MH_211026_1030.xlsx`
- result status: {'success': True, 'error': None, 'validation_message': 'Skipped - Using existing refreshed file'}
- original_file: `ITC_Report_Vega_Test_Traders_MH_211026_1030.xlsx`  refreshed_file: `ITC_Report_Vega_Test_Traders_MH_211026_1030.xlsx`
- stub calls: [('extract', 'ITC_Report_Vega_Test_Traders_MH_211026_1030.xlsx')]
- GUI log lines: ['Skip refresh enabled - using existing: ITC_Report_Vega_Test_Traders_MH_211026_1030.xlsx']
- logger errors: []
- files after: ['ITC_Report_Vega_Test_Traders_MH_211026_1030.xlsx']

##### empty suffix

- suffix `''`, skip_refresh=False
- files before: ['ITC_Report_Alpha_MH_211026_1030.xlsx']
- glob('ITC_Report_*.xlsx') order on this disk: ['ITC_Report_Alpha_MH_211026_1030.xlsx']  -> report_files[0] = `ITC_Report_Alpha_MH_211026_1030.xlsx`
- result status: {'success': False, 'error': "PosixPath('/root/merge-scratch/run/A-TRACE/work/t5/empty_suffix/ITC_Report_Alpha_MH_211026_1030.xlsx') and PosixPath('/root/merge-scratch/run/A-TRACE/work/t5/empty_suffix/ITC_Report_Alpha_MH_211026_1030.xlsx') are the same file"}
- original_file: ``  refreshed_file: ``
- stub calls: []
- GUI log lines: []
- logger errors: ["Error processing ITC report: PosixPath('/root/merge-scratch/run/A-TRACE/work/t5/empty_suffix/ITC_Report_Alpha_MH_211026_1030.xlsx') and PosixPath('/root/merge-s"]
- files after: ['ITC_Report_Alpha_MH_211026_1030.xlsx']

##### suffix without timestamp, run twice (2nd run)

- suffix `'_Refreshed'`, skip_refresh=False
- files before: ['ITC_Report_Alpha_MH_211026_1030.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed.xlsx']
- glob('ITC_Report_*.xlsx') order on this disk: ['ITC_Report_Alpha_MH_211026_1030_Refreshed.xlsx', 'ITC_Report_Alpha_MH_211026_1030.xlsx']  -> report_files[0] = `ITC_Report_Alpha_MH_211026_1030_Refreshed.xlsx`
- result status: {'success': True, 'error': None, 'validation_message': 'stub'}
- original_file: `ITC_Report_Alpha_MH_211026_1030_Refreshed.xlsx`  refreshed_file: `ITC_Report_Alpha_MH_211026_1030_Refreshed_Refreshed.xlsx`
- stub calls: [('refresh', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_Refreshed.xlsx'), ('extract', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_Refreshed.xlsx')]
- GUI log lines: []
- logger errors: []
- files after: ['ITC_Report_Alpha_MH_211026_1030.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_Refreshed.xlsx']

##### suffix that is only {timestamp}

- suffix `'{timestamp}'`, skip_refresh=False
- files before: ['ITC_Report_Alpha_MH_211026_1030.xlsx']
- glob('ITC_Report_*.xlsx') order on this disk: ['ITC_Report_Alpha_MH_211026_1030.xlsx']  -> report_files[0] = `ITC_Report_Alpha_MH_211026_1030.xlsx`
- result status: {'success': True, 'error': None, 'validation_message': 'stub'}
- original_file: `ITC_Report_Alpha_MH_211026_1030.xlsx`  refreshed_file: `ITC_Report_Alpha_MH_211026_1030211026_0422.xlsx`
- stub calls: [('refresh', 'ITC_Report_Alpha_MH_211026_1030211026_0422.xlsx'), ('extract', 'ITC_Report_Alpha_MH_211026_1030211026_0422.xlsx')]
- GUI log lines: []
- logger errors: []
- files after: ['ITC_Report_Alpha_MH_211026_1030.xlsx', 'ITC_Report_Alpha_MH_211026_1030211026_0422.xlsx']

##### no ITC report at all

- suffix `'_Refreshed_{timestamp}'`, skip_refresh=False
- files before: ['Sales_Report_Alpha_MH_211026_1030.xlsx']
- glob('ITC_Report_*.xlsx') order on this disk: []  -> report_files[0] = `None`
- result status: {'success': False, 'error': 'No ITC report found'}
- original_file: ``  refreshed_file: ``
- stub calls: []
- GUI log lines: []
- logger errors: []
- files after: ['Sales_Report_Alpha_MH_211026_1030.xlsx']

##### upper-case extension .XLSX only

- suffix `'_Refreshed_{timestamp}'`, skip_refresh=False
- files before: ['ITC_Report_Alpha_MH_211026_1030.XLSX']
- glob('ITC_Report_*.xlsx') order on this disk: []  -> report_files[0] = `None`
- result status: {'success': False, 'error': 'No ITC report found'}
- original_file: ``  refreshed_file: ``
- stub calls: []
- GUI log lines: []
- logger errors: []
- files after: ['ITC_Report_Alpha_MH_211026_1030.XLSX']

#### extractor_window._classify_version_reports and _latest_refresh_times on one folder

files: ['GSTR3B-Alpha-MH-Apr.xlsx', 'ITC_Report_Alpha_MH_211026_1030.xlsx', 'ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_1100.xlsx', 'Sales_Report_Alpha_MH_211026_1030_Updated_211026_1110.xlsx', 'sales_report_alpha_mh_211026_1030.XLSX']
_classify_version_reports -> (True, True)
_latest_refresh_times(folder, None) -> ('2026-10-21 03:52', None)
_latest_refresh_times(folder, '_Refreshed_{timestamp}') -> ('2026-10-21 03:52', None)
_latest_refresh_times(folder, '_Updated_{timestamp}') -> ('2026-10-21 03:52', '2026-10-21 04:02')
_latest_refresh_times(folder, '{timestamp}') -> ('2026-10-21 03:52', None)
_classify_version_reports(missing folder) -> (False, False)
_latest_refresh_times(missing folder) -> (None, None)

<!-- end of outputs/t5_refresh_naming.md -->

### T10 - what a failed refresh leaves behind (RAN)

Script `scripts/t10_failed_refresh_leftover.py`: an original report dated three days earlier; the refresh stub returns failure (as `_refresh_power_query_simple` does on any Excel error, `report_processor.py:376-393`); then a second run with "skip refresh" on. Output `outputs/t10_failed_refresh_leftover.md`, reproduced exactly:



<!-- RAN output: outputs/t10_failed_refresh_leftover.md (headings demoted by two levels) -->

run 1 status: {'success': False, 'error': 'Excel error: (simulated) Call was rejected by callee'}
  ITC_Report_Alpha_MH_211026_1030.xlsx  mtime=2026-10-18 04:29
  ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_0429.xlsx  mtime=2026-10-18 04:29
client list 'last refreshed' (ITC, Sales): ('2026-10-18 04:29', None)
run 2 with skip_refresh: status {'success': True, 'error': None, 'validation_message': 'Skipped - Using existing refreshed file'} | extracted from ['ITC_Report_Alpha_MH_211026_1030_Refreshed_211026_0429.xlsx']

<!-- end of outputs/t10_failed_refresh_leftover.md -->

### T5 / T10 observations

1. **The "original" report can be a refreshed copy.** `glob` returns files in whatever order the file system lists them, and the code takes the first (`report_processor.py:113,125`). On this Linux disk the refreshed copy was listed first, so the run refreshed a copy of the copy and created `..._Refreshed_211026_1040_Refreshed_211026_0422.xlsx`. RAN, HIGH that the order is not controlled. MEDIUM: on a local NTFS disk Windows normally lists names in sorted order, and `.` sorts before `_`, so the plain original would usually come first there; on network shares served by other file systems the order can be arbitrary, as here. With "skip refresh", if the older of two refreshed copies happened to be listed first, its stem would be used as the "original" stem and only copies of THAT file would qualify, so an older refreshed copy could be reused instead of the newest (READ, MEDIUM; in my run the newest was listed first, so it was not shown).
2. **A failed refresh leaves a file that later counts as a good refresh** (T10). `shutil.copy2` keeps the original's modification time, and on failure nothing is saved, so the "_Refreshed_" copy is an unrefreshed duplicate dated like the original. The client list then shows "ITC: 2026-10-18 04:29" (the original's date) as the last refresh, and the next run with "skip refresh" reuses it and reports `success: True, 'Skipped - Using existing refreshed file'`; the consolidated report would then show ✓ and the unrefreshed figures. RAN (stubs for Excel), HIGH. Nothing removes failed copies (READ).
3. **Names that contain the marker are mistaken for refreshed files**: a client whose cleaned name contains `Refreshed` (made-up `Refreshed Test Traders`) has an ORIGINAL report `ITC_Report_Refreshed_Test_Traders_...xlsx` that passes `is_refreshed_name` under every pattern (legacy marker); with "skip refresh" the unrefreshed original is used as the refreshed file. With a custom suffix `_v{timestamp}`, any client whose cleaned name starts with V (made-up `Vega Test Traders`) behaves the same. RAN, HIGH.
4. `is_refreshed_name` always also accepts the legacy marker, so with a custom pattern both new-style and old-style files count (by design, per the docstring at 17-19). `{timestamp}` alone or an empty pattern fall back to the legacy marker; but the processor then builds names WITHOUT `_refreshed_` (`ITC_Report_X211026_1100.xlsx`), so files made with the pattern `{timestamp}` are never recognised: the list says "Never refreshed" and skip-refresh never finds them. RAN, HIGH.
5. **Empty suffix**: the copy's name equals the original's, `shutil.copy2` raises "are the same file", the report fails with that message and no extraction happens. RAN, HIGH. **Suffix without `{timestamp}`** (e.g. `_Refreshed`): every refresh writes the same name, and once a `_Refreshed` copy exists the glob can pick it as the original and produce `_Refreshed_Refreshed` (RAN). **Upper-case `{TIMESTAMP}`** is not replaced and appears literally in the file name (RAN Part 1).
6. A refreshed copy has minute resolution: two refreshes of the same report in the same minute overwrite the earlier copy (`shutil.copy2` over an existing file, line 164). READ, HIGH.
7. An upper-case `.XLSX` report is not found by the glob on Linux (RAN: "No ITC report found"); on Windows glob is case-insensitive (READ comment at `extractor_window.py:27-28`, MEDIUM). The app itself always writes `.xlsx` (`gui/handlers/processing_handler.py:298,356`), so this matters only for hand-made files.
8. `_classify_version_reports` (`extractor_window.py:20-42`) ticks ITC/Sales when ANY `itc_report_*.xlsx` / `sales_report_*.xlsx` exists, refreshed copies included, any case (RAN: `(True, True)` with a lower-case `.XLSX` Sales original present). `_latest_refresh_times` shows the newest modification time of recognised refreshed copies; missing folder -> `(None, None)` (RAN).
9. Excel's own lock files (`~$...`) pass `is_refreshed_name` (RAN) but are excluded by the callers because they require the name to start with `itc_report_`/`sales_report_` or with the original's stem (READ, HIGH).

## 8. T6 - Folder scanning and organising on a fake tree (RAN)

**What was built (`work/t6/source/`, made-up names, every `.xlsx` a real tiny workbook made by openpyxl):**

- Client 1 **complete**: `Alpha Test Traders`, Maharashtra - one file of each of the six types plus four GSTR-3B months (one with a leading `(1) `, one with upper-case `.XLSX`).
- Client 2 **missing files**: `Beta Demo Industries Private Limited`, Gujarat - only a GSTR-3B and a Sales `.xls` (OLE signature).
- Client 3 **duplicates + unrecognised**: `Gamma Sample Enterprises LLP`, Karnataka - two Sales (`... (1)`), two IMS (`... (1)`), a GSTR-3B whose state is written `KA`, a GSTR-2B Reco, an Annual Report whose client is written in CAPITALS, a Sales Reco with a hyphenated client name, `Gamma notes.xlsx`, `Gamma GSTR3B Jul.xlsm`.
- Files that should never reach the parser: `tiny.xlsx` (500 bytes), `fake.xlsx` (text), a `.csv`, a `.xlsb`, an Excel lock file `~$GSTR3B-...xlsx` (165 bytes), a sub-folder `Alpha extra/` containing a valid GSTR-3B, and a FOLDER named `Sales-Fake Folder-Delhi-Apr-Mar.xlsx`.
- Templates `work/t6/templates/ITC_Template.xlsx` and `Sales_Template.xlsx`, each with a `Links` sheet (no Power Query - openpyxl cannot make one).

**What ran:** `FileParser.scan_folder` (`core/file_parser.py:136-177`), the Validation summary text builder `GSTOrganizerApp.update_summary` (`gui/main_window.py:294-329`, on a stub), then the real orchestrator `ProcessingHandler.process_files_thread` (`gui/handlers/processing_handler.py:182-530`) with a stub app object instead of the window (Tk variables replaced by simple objects; `root.after` runs the callback at once; message boxes recorded; `os.system` recorded instead of opening a file browser). It creates `FileOrganizer` (`core/file_organizer.py`), copies files, builds reports with `ExcelHandler` - **on Linux this is the openpyxl fallback `_create_report_openpyxl` (`core/excel_handler.py:262-322`), not the Windows COM path** - writes the per-client text report and the Level-1 summary workbook. The folder time stamp was fixed per run by replacing `get_timestamp` in `core/file_organizer.py` and `gui/handlers/processing_handler.py` (so runs 2-4 could be "in the same minute"). Backup names use the real clock (`get_safe_timestamp`, `utils/helpers.py:224`). Before/after snapshots of the source tree record every file's sha256, size and modification time (nanoseconds) and every folder's modification time.

Script `scripts/t6_scan_organise.py`, output `outputs/t6_scan_organise.md`, reproduced exactly (long: it contains the full GUI log of each run):



<!-- RAN output: outputs/t6_scan_organise.md (headings demoted by two levels) -->

#### 1. Source tree created (made-up names, fake contents)

| path (relative to source) | sha256[:16] | bytes |
|---|---|---|
| `(1) GSTR3B-Alpha Test Traders-Maharashtra-Jun.xlsx` | 4066f5a683ca318c | 4849 |
| `Alpha extra/` | dir | 0 |
| `Alpha extra/GSTR3B-Alpha Test Traders-Maharashtra-Aug.xlsx` | 4066f5a683ca318c | 4849 |
| `AnnualReport-Alpha Test Traders-Maharashtra-2024-25.xlsx` | 4066f5a683ca318c | 4849 |
| `AnnualReport-GAMMA SAMPLE ENTERPRISES LLP-Karnataka-2024-25.xlsx` | 4066f5a683ca318c | 4849 |
| `GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Apr24-Mar25.xlsx` | 4066f5a683ca318c | 4849 |
| `GSTR-2B-Reco-Gamma Sample Enterprises LLP-Karnataka-Q1.xlsx` | 4066f5a683ca318c | 4849 |
| `GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx` | 4066f5a683ca318c | 4849 |
| `GSTR3B-Alpha Test Traders-Maharashtra-Jul.csv` | 492d5ea496056f1a | 8 |
| `GSTR3B-Alpha Test Traders-Maharashtra-May.xlsx` | 4066f5a683ca318c | 4849 |
| `GSTR3B-Alpha Test Traders-Maharashtra-Oct.xlsb` | 4066f5a683ca318c | 4849 |
| `GSTR3B-Alpha Test Traders-Maharashtra-Sep.XLSX` | 4066f5a683ca318c | 4849 |
| `GSTR3B-Beta Demo Industries Private Limited-Gujarat-Apr.xlsx` | 4066f5a683ca318c | 4849 |
| `GSTR3B-Gamma Sample Enterprises LLP-KA-Apr.xlsx` | 4066f5a683ca318c | 4849 |
| `Gamma GSTR3B Jul.xlsm` | 4066f5a683ca318c | 4849 |
| `Gamma notes.xlsx` | 4066f5a683ca318c | 4849 |
| `ImsReco-Alpha Test Traders-Maharashtra-31032025.xlsx` | 4066f5a683ca318c | 4849 |
| `ImsReco-Gamma Sample Enterprises LLP-Karnataka-31032025 (1).xlsx` | 4066f5a683ca318c | 4849 |
| `ImsReco-Gamma Sample Enterprises LLP-Karnataka-31032025.xlsx` | 4066f5a683ca318c | 4849 |
| `Sales-Alpha Test Traders-Maharashtra-Apr-Mar.xlsx` | 4066f5a683ca318c | 4849 |
| `Sales-Beta Demo Industries Private Limited-Gujarat-Apr-Mar.xls` | 351608981e4b29d9 | 4096 |
| `Sales-Fake Folder-Delhi-Apr-Mar.xlsx/` | dir | 0 |
| `Sales-Gamma Sample Enterprises LLP-Karnataka-Apr-Mar (1).xlsx` | 4066f5a683ca318c | 4849 |
| `Sales-Gamma Sample Enterprises LLP-Karnataka-Apr-Mar.xlsx` | 4066f5a683ca318c | 4849 |
| `SalesReco-Alpha Test Traders-Maharashtra-FY 2024-25.xlsx` | 4066f5a683ca318c | 4849 |
| `SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1.xlsx` | 4066f5a683ca318c | 4849 |
| `fake.xlsx` | 7e3c46189bc6011f | 3600 |
| `tiny.xlsx` | 739897c7c698d5b3 | 500 |
| `~$GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx` | fc45a6ec2ad955e5 | 165 |

#### 2. FileParser.scan_folder(source)  (core/file_parser.py:136-177)

progress callback calls: 21; first [(1, 21, 'Processing (1) GSTR3B-Alpha Test Traders-Maharashtra-Jun.xlsx'), (2, 21, 'Processing AnnualReport-Alpha Test Traders-Maharashtra-2024-25.xlsx')]; last [(21, 21, 'Processing SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1.xlsx')]

scanned_files: 21 entries (keyed by file name):

| file | parsed | type | client | state | size | ext |
|---|---|---|---|---|---|---|
| `(1) GSTR3B-Alpha Test Traders-Maharashtra-Jun.xlsx` | True | GSTR-3B Export | `Alpha Test Traders` | `Maharashtra` | 4849 | .xlsx |
| `AnnualReport-Alpha Test Traders-Maharashtra-2024-25.xlsx` | True | Annual Report | `Alpha Test Traders` | `Maharashtra` | 4849 | .xlsx |
| `AnnualReport-GAMMA SAMPLE ENTERPRISES LLP-Karnataka-2024-25.xlsx` | True | Annual Report | `GAMMA SAMPLE ENTERPRISES LLP` | `Karnataka` | 4849 | .xlsx |
| `GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Apr24-Mar25.xlsx` | True | GSTR-2B Reco | `Alpha Test Traders` | `Maharashtra` | 4849 | .xlsx |
| `GSTR-2B-Reco-Gamma Sample Enterprises LLP-Karnataka-Q1.xlsx` | True | GSTR-2B Reco | `Gamma Sample Enterprises LLP` | `Karnataka` | 4849 | .xlsx |
| `GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx` | True | GSTR-3B Export | `Alpha Test Traders` | `Maharashtra` | 4849 | .xlsx |
| `GSTR3B-Alpha Test Traders-Maharashtra-May.xlsx` | True | GSTR-3B Export | `Alpha Test Traders` | `Maharashtra` | 4849 | .xlsx |
| `GSTR3B-Alpha Test Traders-Maharashtra-Sep.XLSX` | True | GSTR-3B Export | `Alpha Test Traders` | `Maharashtra` | 4849 | .xlsx |
| `GSTR3B-Beta Demo Industries Private Limited-Gujarat-Apr.xlsx` | True | GSTR-3B Export | `Beta Demo Industries Private Limited` | `Gujarat` | 4849 | .xlsx |
| `GSTR3B-Gamma Sample Enterprises LLP-KA-Apr.xlsx` | True | GSTR-3B Export | `Gamma Sample Enterprises LLP` | `KA` | 4849 | .xlsx |
| `Gamma GSTR3B Jul.xlsm` | False | Unknown | `` | `` | 4849 | .xlsm |
| `Gamma notes.xlsx` | False | Unknown | `` | `` | 4849 | .xlsx |
| `ImsReco-Alpha Test Traders-Maharashtra-31032025.xlsx` | True | IMS Reco | `Alpha Test Traders` | `Maharashtra` | 4849 | .xlsx |
| `ImsReco-Gamma Sample Enterprises LLP-Karnataka-31032025 (1).xlsx` | True | IMS Reco | `Gamma Sample Enterprises LLP` | `Karnataka` | 4849 | .xlsx |
| `ImsReco-Gamma Sample Enterprises LLP-Karnataka-31032025.xlsx` | True | IMS Reco | `Gamma Sample Enterprises LLP` | `Karnataka` | 4849 | .xlsx |
| `Sales-Alpha Test Traders-Maharashtra-Apr-Mar.xlsx` | True | Sales | `Alpha Test Traders` | `Maharashtra` | 4849 | .xlsx |
| `Sales-Beta Demo Industries Private Limited-Gujarat-Apr-Mar.xls` | True | Sales | `Beta Demo Industries Private Limited` | `Gujarat` | 4096 | .xls |
| `Sales-Gamma Sample Enterprises LLP-Karnataka-Apr-Mar (1).xlsx` | True | Sales | `Gamma Sample Enterprises LLP` | `Karnataka` | 4849 | .xlsx |
| `Sales-Gamma Sample Enterprises LLP-Karnataka-Apr-Mar.xlsx` | True | Sales | `Gamma Sample Enterprises LLP` | `Karnataka` | 4849 | .xlsx |
| `SalesReco-Alpha Test Traders-Maharashtra-FY 2024-25.xlsx` | True | Sales Reco | `Alpha Test Traders` | `Maharashtra` | 4849 | .xlsx |
| `SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1.xlsx` | True | Sales Reco | `Gamma` | `Sample` | 4849 | .xlsx |

files in the source tree that never reached the parser: ['Alpha extra/GSTR3B-Alpha Test Traders-Maharashtra-Aug.xlsx', 'GSTR3B-Alpha Test Traders-Maharashtra-Jul.csv', 'GSTR3B-Alpha Test Traders-Maharashtra-Oct.xlsb', 'fake.xlsx', 'tiny.xlsx', '~$GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx']

client_data keys: ['Alpha Test Traders-MH', 'GAMMA SAMPLE ENTERPRISES LLP-KA', 'Gamma Sample Enterprises LLP-KA', 'Beta Demo Industries Private Limited-GJ', 'Gamma-SAM']

| client key | client | state | state_code | status | icon | file_count | completeness | missing_files | extra_files | files per type |
|---|---|---|---|---|---|---|---|---|---|---|
| `Alpha Test Traders-MH` | `Alpha Test Traders` | `Maharashtra` | MH | Complete | ✅ | 9 | 100.0 | [] | [] | {'GSTR-3B Export': ['(1) GSTR3B-Alpha Test Traders-Maharashtra-Jun.xlsx', 'GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx', 'GSTR3B-Alpha Test Traders-Maharashtra-May.xlsx', 'GSTR3B-Alpha Test Traders-Maharashtra-Sep.XLSX'], 'Annual Report': ['AnnualReport-Alpha Test Traders-Maharashtra-2024-25.xlsx'], 'GSTR-2B Reco': ['GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Apr24-Mar25.xlsx'], 'IMS Reco': ['ImsReco-Alpha Test Traders-Maharashtra-31032025.xlsx'], 'Sales': ['Sales-Alpha Test Traders-Maharashtra-Apr-Mar.xlsx'], 'Sales Reco': ['SalesReco-Alpha Test Traders-Maharashtra-FY 2024-25.xlsx']} |
| `GAMMA SAMPLE ENTERPRISES LLP-KA` | `GAMMA SAMPLE ENTERPRISES LLP` | `Karnataka` | KA | Missing 5 files | ⚠️ | 1 | 16.7 | ['GSTR-2B Reco', 'GSTR-3B Export', 'IMS Reco', 'Sales', 'Sales Reco'] | [] | {'Annual Report': ['AnnualReport-GAMMA SAMPLE ENTERPRISES LLP-Karnataka-2024-25.xlsx']} |
| `Gamma Sample Enterprises LLP-KA` | `Gamma Sample Enterprises LLP` | `Karnataka` | KA | Missing 2 files | ⚠️ | 6 | 66.7 | ['Annual Report', 'Sales Reco'] | ['Multiple IMS Reco (2 files)', 'Multiple Sales (2 files)'] | {'GSTR-2B Reco': ['GSTR-2B-Reco-Gamma Sample Enterprises LLP-Karnataka-Q1.xlsx'], 'GSTR-3B Export': ['GSTR3B-Gamma Sample Enterprises LLP-KA-Apr.xlsx'], 'IMS Reco': ['ImsReco-Gamma Sample Enterprises LLP-Karnataka-31032025 (1).xlsx', 'ImsReco-Gamma Sample Enterprises LLP-Karnataka-31032025.xlsx'], 'Sales': ['Sales-Gamma Sample Enterprises LLP-Karnataka-Apr-Mar (1).xlsx', 'Sales-Gamma Sample Enterprises LLP-Karnataka-Apr-Mar.xlsx']} |
| `Beta Demo Industries Private Limited-GJ` | `Beta Demo Industries Private Limited` | `Gujarat` | GJ | Missing 4 files | ⚠️ | 2 | 33.3 | ['Annual Report', 'GSTR-2B Reco', 'IMS Reco', 'Sales Reco'] | [] | {'GSTR-3B Export': ['GSTR3B-Beta Demo Industries Private Limited-Gujarat-Apr.xlsx'], 'Sales': ['Sales-Beta Demo Industries Private Limited-Gujarat-Apr-Mar.xls']} |
| `Gamma-SAM` | `Gamma` | `Sample` | SAM | Missing 5 files | ⚠️ | 1 | 16.7 | ['Annual Report', 'GSTR-2B Reco', 'GSTR-3B Export', 'IMS Reco', 'Sales'] | [] | {'Sales Reco': ['SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1.xlsx']} |

variations (2):

| filename | issue | size | pattern_expected |
|---|---|---|---|
| `Gamma GSTR3B Jul.xlsm` | No matching pattern found | 4849 | GSTR3B-ClientName-State-Month.xlsx |
| `Gamma notes.xlsx` | No matching pattern found | 4849 | Check file naming convention |

parser.errors: []

logger WARNING/ERROR during scan: [('WARNING', "Unknown state 'KA', using fallback: KA"), ('WARNING', 'Could not parse filename: Gamma GSTR3B Jul.xlsm'), ('WARNING', 'Could not parse filename: Gamma notes.xlsx'), ('WARNING', "Unknown state 'Sample', using fallback: SAM")]

get_statistics(): {
 "total_files": 21,
 "parsed_files": 19,
 "unparsed_files": 2,
 "total_clients": 5,
 "complete_clients": 1,
 "incomplete_clients": 4,
 "variations": 2,
 "total_size": 101076,
 "total_size_formatted": "98.7 KB",
 "file_type_distribution": {
  "GSTR-3B Export": 6,
  "Annual Report": 2,
  "GSTR-2B Reco": 2,
  "IMS Reco": 3,
  "Sales": 4,
  "Sales Reco": 2
 },
 "parsing_rate": 90.47619047619048,
 "completion_rate": 20.0
}

validate_results() (never called by the app): [{'type': 'Missing Critical Files', 'client': 'GAMMA SAMPLE ENTERPRISES LLP-KA', 'severity': 'Warning', 'details': 'GSTR-3B Export'}, {'type': 'Missing Critical Files', 'client': 'Gamma Sample Enterprises LLP-KA', 'severity': 'Warning', 'details': 'Annual Report'}, {'type': 'Missing Critical Files', 'client': 'Beta Demo Industries Private Limited-GJ', 'severity': 'Warning', 'details': 'Annual Report'}, {'type': 'Missing Critical Files', 'client': 'Gamma-SAM', 'severity': 'Warning', 'details': 'Annual Report, GSTR-3B Export'}]

Validation-screen summary text (GSTOrganizerApp.update_summary, gui/main_window.py:294-329):
```
📊 SCAN RESULTS SUMMARY:
════════════════════════════════════════════
📁 FILES ANALYSIS:
   Total Excel Files Found: 21
   ✅ Successfully Parsed: 19
   ❌ Unparsed/Variations: 2
   📊 Parsing Success Rate: 90.5%

👥 CLIENTS ANALYSIS:
   Total Clients Identified: 5
   ✅ Clients with All Files: 1
   ⚠️ Clients with Missing Files: 4
   📈 Completeness Rate: 20.0%

📋 FILE TYPE DISTRIBUTION:
   • GSTR-3B Export: 6 files
   • Annual Report: 2 files
   • GSTR-2B Reco: 2 files
   • IMS Reco: 3 files
   • Sales: 4 files
   • Sales Reco: 2 files

🔍 FILE VARIATIONS: 2 files

💾 TOTAL DATA SIZE: 98.7 KB
```

#### 3. Organise: ProcessingHandler.process_files_thread with a stub app (fresh mode)


##### Run 1: mode='fresh', timestamp='211026 1030', clients=['Alpha Test Traders-MH', 'GAMMA SAMPLE ENTERPRISES LLP-KA', 'Gamma Sample Enterprises LLP-KA', 'Beta Demo Industries Private Limited-GJ', 'Gamma-SAM']

GUI log (level, message):
```
[info] 
🏢 Processing Alpha Test Traders - Maharashtra
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1030
[normal]     Level 2: Alpha Test Traders-MH
[normal]     Level 3: Version-211026 1030
[success]   ✓ Organized 9 files
[normal]     📁 (1) GSTR3B-Alpha Test Traders-Maharashtra-Jun.xlsx → GSTR-3B Export folder
[normal]     📁 GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx → GSTR-3B Export folder
[normal]     📁 GSTR3B-Alpha Test Traders-Maharashtra-May.xlsx → GSTR-3B Export folder
[normal]     📁 GSTR3B-Alpha Test Traders-Maharashtra-Sep.XLSX → GSTR-3B Export folder
[normal]     📁 AnnualReport-Alpha Test Traders-Maharashtra-2024-25.xlsx → Annual Report folder
[normal]     📁 GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Apr24-Mar25.xlsx → GSTR-2B Reco folder
[normal]     📁 ImsReco-Alpha Test Traders-Maharashtra-31032025.xlsx → IMS Reco folder
[normal]     📁 Sales-Alpha Test Traders-Maharashtra-Apr-Mar.xlsx → Sales folder
[normal]     📁 SalesReco-Alpha Test Traders-Maharashtra-FY 2024-25.xlsx → Sales Reco folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[success]     ✅ Success: ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[normal]     [DEBUG] Output file created, size: 5458 bytes
[success]   ✅ Completed Alpha Test Traders
[info] 
🏢 Processing GAMMA SAMPLE ENTERPRISES LLP - Karnataka
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1030
[normal]     Level 2: GAMMA SAMPLE ENTERPRISES LLP-KA
[normal]     Level 3: Version-211026 1030
[success]   ✓ Organized 1 files
[normal]     📁 AnnualReport-GAMMA SAMPLE ENTERPRISES LLP-Karnataka-2024-25.xlsx → Annual Report folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1030.xlsx
[success]     ✅ Success: ITC_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1030.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1030.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/GAMMA SAMPLE ENTERPRISES LLP-KA/Version-211026 1030/Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1030.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1030.xlsx
[normal]     [DEBUG] Output file created, size: 5417 bytes
[success]   ✅ Completed GAMMA SAMPLE ENTERPRISES LLP
[info] 
🏢 Processing Gamma Sample Enterprises LLP - Karnataka
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1030
[normal]     Level 2: Gamma Sample Enterprises LLP-KA
[normal]     Level 3: Version-211026 1030
[success]   ✓ Organized 6 files
[normal]     📁 GSTR-2B-Reco-Gamma Sample Enterprises LLP-Karnataka-Q1.xlsx → GSTR-2B Reco folder
[normal]     📁 GSTR3B-Gamma Sample Enterprises LLP-KA-Apr.xlsx → GSTR-3B Export folder
[normal]     📁 ImsReco-Gamma Sample Enterprises LLP-Karnataka-31032025 (1).xlsx → IMS Reco folder
[normal]     📁 ImsReco-Gamma Sample Enterprises LLP-Karnataka-31032025.xlsx → IMS Reco folder
[normal]     📁 Sales-Gamma Sample Enterprises LLP-Karnataka-Apr-Mar (1).xlsx → Sales folder
[normal]     📁 Sales-Gamma Sample Enterprises LLP-Karnataka-Apr-Mar.xlsx → Sales folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1030.xlsx
[success]     ✅ Success: ITC_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1030.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1030.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Gamma Sample Enterprises LLP-KA/Version-211026 1030/Sales_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1030.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1030.xlsx
[normal]     [DEBUG] Output file created, size: 5414 bytes
[success]   ✅ Completed Gamma Sample Enterprises LLP
[info] 
🏢 Processing Beta Demo Industries Private Limited - Gujarat
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1030
[normal]     Level 2: Beta Demo Industries Pvt Ltd-GJ
[normal]     Level 3: Version-211026 1030
[success]   ✓ Organized 2 files
[normal]     📁 GSTR3B-Beta Demo Industries Private Limited-Gujarat-Apr.xlsx → GSTR-3B Export folder
[normal]     📁 Sales-Beta Demo Industries Private Limited-Gujarat-Apr-Mar.xls → Sales folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_Beta_Demo_Industries_Pvt_Ltd_GJ_211026_1030.xlsx
[success]     ✅ Success: ITC_Report_Beta_Demo_Industries_Pvt_Ltd_GJ_211026_1030.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_Beta_Demo_Industries_Pvt_Ltd_GJ_211026_1030.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Beta Demo Industries Pvt Ltd-GJ/Version-211026 1030/Sales_Report_Beta_Demo_Industries_Pvt_Ltd_GJ_211026_1030.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_Beta_Demo_Industries_Pvt_Ltd_GJ_211026_1030.xlsx
[normal]     [DEBUG] Output file created, size: 5410 bytes
[success]   ✅ Completed Beta Demo Industries Private Limited
[info] 
🏢 Processing Gamma - Sample
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1030
[normal]     Level 2: Gamma-SAM
[normal]     Level 3: Version-211026 1030
[success]   ✓ Organized 1 files
[normal]     📁 SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1.xlsx → Sales Reco folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_Gamma_SAM_211026_1030.xlsx
[success]     ✅ Success: ITC_Report_Gamma_SAM_211026_1030.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_Gamma_SAM_211026_1030.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Gamma-SAM/Version-211026 1030/Sales_Report_Gamma_SAM_211026_1030.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_Gamma_SAM_211026_1030.xlsx
[normal]     [DEBUG] Output file created, size: 5401 bytes
[success]   ✅ Completed Gamma
[info] 
📊 Creating Level 1 summary report...
[normal]   Output: GST_Processing_Summary_211026_1030.xlsx
[success]   ✅ Success: GST_Processing_Summary_211026_1030.xlsx
[normal]   Sheets: Summary, Client Status, File Mapping, Errors, Variations
[success] 
🎉 ════════ PROCESSING COMPLETE ════════
[info] ✅ Processed: 5/5 clients
[info] 📁 Files organized: 19
[info] 📊 Reports created: 10
[info] 📂 Output location: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030
[info] 
⏱️ Total processing time: 0 minutes 0 seconds
[success] 📂 Opened output folder
[info] ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query.
```
progress updates: 36 (last: (100, 'Processing complete | Total time: 0m 0s'))
pop-ups: [('info', ('Processing Complete', 'Successfully processed 5/5 clients\n\nFiles organized in:\n/root/merge-scratch/run/A-TRACE/work/t6/target'))]; os.system calls: ['xdg-open "/root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030"']
logger WARNING/ERROR: [('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'utils.helpers', "Unknown state 'Sample', using fallback: SAM"), ('WARNING', 'utils.helpers', "Unknown state 'Sample', using fallback: SAM"), ('WARNING', 'utils.helpers', "Unknown state 'Sample', using fallback: SAM"), ('WARNING', 'utils.helpers', "Unknown state 'Sample', using fallback: SAM"), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'utils.helpers', "Unknown state 'Sample', using fallback: SAM"), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!')]
file_organizer.get_summary(): {'timestamp': '211026 1030', 'processing_mode': 'fresh', 'target_folder': '/root/merge-scratch/run/A-TRACE/work/t6/target', 'folders_created': 26, 'files_processed': 19, 'successful_copies': 19, 'failed_copies': 0, 'skipped_files': 0, 'errors': 0}

Target tree after run 1:
```
Annual Statement-211026 1030/
  Alpha Test Traders-MH/
    Version-211026 1030/
      AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25.xlsx  (4849 B)
      GSTR3B/
        (1)_GSTR3B-Alpha_Test_Traders-Maharashtra-Jun.xlsx  (4849 B)
        GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx  (4849 B)
        GSTR3B-Alpha_Test_Traders-Maharashtra-May.xlsx  (4849 B)
        GSTR3B-Alpha_Test_Traders-Maharashtra-Sep.XLSX  (4849 B)
      ITC/
        GSTR-2B-Reco-Alpha_Test_Traders-Maharashtra-Apr24-Mar25.xlsx  (4849 B)
        ImsReco-Alpha_Test_Traders-Maharashtra-31032025.xlsx  (4849 B)
      ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx  (5483 B)
      Sales/
        Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx  (4849 B)
        SalesReco-Alpha_Test_Traders-Maharashtra-FY_2024-25.xlsx  (4849 B)
      Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx  (5458 B)
      _Organization_Report_Alpha Test Traders_Maharashtra.txt  (1302 B)
  Beta Demo Industries Pvt Ltd-GJ/
    Version-211026 1030/
      GSTR3B/
        GSTR3B-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr.xlsx  (4849 B)
      ITC/
      ITC_Report_Beta_Demo_Industries_Pvt_Ltd_GJ_211026_1030.xlsx  (5377 B)
      Sales/
        Sales-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr-Mar.xls  (4096 B)
      Sales_Report_Beta_Demo_Industries_Pvt_Ltd_GJ_211026_1030.xlsx  (5410 B)
      _Organization_Report_Beta Demo Industries Private Limited_Gujarat.txt  (873 B)
  GAMMA SAMPLE ENTERPRISES LLP-KA/
    Version-211026 1030/
      AnnualReport-GAMMA_SAMPLE_ENTERPRISES_LLP-Karnataka-2024-25.xlsx  (4849 B)
      GSTR3B/
      ITC/
      ITC_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1030.xlsx  (5432 B)
      Sales/
      Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1030.xlsx  (5417 B)
      _Organization_Report_GAMMA SAMPLE ENTERPRISES LLP_Karnataka.txt  (801 B)
  GST_Processing_Summary_211026_1030.xlsx  (9262 B)
  Gamma Sample Enterprises LLP-KA/
    Version-211026 1030/
      GSTR3B/
        GSTR3B-Gamma_Sample_Enterprises_LLP-KA-Apr.xlsx  (4849 B)
      ITC/
        GSTR-2B-Reco-Gamma_Sample_Enterprises_LLP-Karnataka-Q1.xlsx  (4849 B)
        ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025.xlsx  (4849 B)
        ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025_(1).xlsx  (4849 B)
      ITC_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1030.xlsx  (5464 B)
      Sales/
        Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar.xlsx  (4849 B)
        Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar_(1).xlsx  (4849 B)
      Sales_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1030.xlsx  (5414 B)
      _Organization_Report_Gamma Sample Enterprises LLP_Karnataka.txt  (1175 B)
  Gamma-SAM/
    Version-211026 1030/
      GSTR3B/
      ITC/
      ITC_Report_Gamma_SAM_211026_1030.xlsx  (5360 B)
      Sales/
        SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1.xlsx  (4849 B)
      Sales_Report_Gamma_SAM_211026_1030.xlsx  (5401 B)
      _Organization_Report_Gamma_Sample.txt  (676 B)
```

##### Links sheet contents of every created report (openpyxl fallback path, core/excel_handler.py:262-322)

- `Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx` sheets=['Links', 'Info']
    - Links: {'A1': 'ITC template - links sheet', 'B2': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Alpha Test Traders-MH\\Version-211026 1030\\GSTR3B', 'B4': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Alpha Test Traders-MH\\Version-211026 1030', 'B5': 'AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25', 'B7': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Alpha Test Traders-MH\\Version-211026 1030\\ITC', 'B8': 'GSTR-2B-Reco-Alpha_Test_Traders-Maharashtra-Apr24-Mar25', 'B10': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Alpha Test Traders-MH\\Version-211026 1030\\ITC', 'B11': 'ImsReco-Alpha_Test_Traders-Maharashtra-31032025'}
- `Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx` sheets=['Links', 'Info']
    - Links: {'A1': 'Sales template - links sheet', 'B2': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Alpha Test Traders-MH\\Version-211026 1030\\Sales', 'B3': 'Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar', 'B5': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Alpha Test Traders-MH\\Version-211026 1030', 'B6': 'AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25', 'B8': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Alpha Test Traders-MH\\Version-211026 1030\\Sales', 'B9': 'SalesReco-Alpha_Test_Traders-Maharashtra-FY_2024-25'}
- `Annual Statement-211026 1030/Beta Demo Industries Pvt Ltd-GJ/Version-211026 1030/ITC_Report_Beta_Demo_Industries_Pvt_Ltd_GJ_211026_1030.xlsx` sheets=['Links', 'Info']
    - Links: {'A1': 'ITC template - links sheet', 'B2': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Beta Demo Industries Pvt Ltd-GJ\\Version-211026 1030\\GSTR3B'}
- `Annual Statement-211026 1030/Beta Demo Industries Pvt Ltd-GJ/Version-211026 1030/Sales_Report_Beta_Demo_Industries_Pvt_Ltd_GJ_211026_1030.xlsx` sheets=['Links', 'Info']
    - Links: {'A1': 'Sales template - links sheet', 'B2': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Beta Demo Industries Pvt Ltd-GJ\\Version-211026 1030\\Sales', 'B3': 'Sales-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr-Mar'}
- `Annual Statement-211026 1030/GAMMA SAMPLE ENTERPRISES LLP-KA/Version-211026 1030/ITC_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1030.xlsx` sheets=['Links', 'Info']
    - Links: {'A1': 'ITC template - links sheet', 'B2': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\GAMMA SAMPLE ENTERPRISES LLP-KA\\Version-211026 1030\\GSTR3B', 'B4': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\GAMMA SAMPLE ENTERPRISES LLP-KA\\Version-211026 1030', 'B5': 'AnnualReport-GAMMA_SAMPLE_ENTERPRISES_LLP-Karnataka-2024-25'}
- `Annual Statement-211026 1030/GAMMA SAMPLE ENTERPRISES LLP-KA/Version-211026 1030/Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1030.xlsx` sheets=['Links', 'Info']
    - Links: {'A1': 'Sales template - links sheet', 'B5': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\GAMMA SAMPLE ENTERPRISES LLP-KA\\Version-211026 1030', 'B6': 'AnnualReport-GAMMA_SAMPLE_ENTERPRISES_LLP-Karnataka-2024-25'}
- `Annual Statement-211026 1030/Gamma Sample Enterprises LLP-KA/Version-211026 1030/ITC_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1030.xlsx` sheets=['Links', 'Info']
    - Links: {'A1': 'ITC template - links sheet', 'B2': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Gamma Sample Enterprises LLP-KA\\Version-211026 1030\\GSTR3B', 'B7': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Gamma Sample Enterprises LLP-KA\\Version-211026 1030\\ITC', 'B8': 'GSTR-2B-Reco-Gamma_Sample_Enterprises_LLP-Karnataka-Q1', 'B10': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Gamma Sample Enterprises LLP-KA\\Version-211026 1030\\ITC', 'B11': 'ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025_(1)'}
- `Annual Statement-211026 1030/Gamma Sample Enterprises LLP-KA/Version-211026 1030/Sales_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1030.xlsx` sheets=['Links', 'Info']
    - Links: {'A1': 'Sales template - links sheet', 'B2': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Gamma Sample Enterprises LLP-KA\\Version-211026 1030\\Sales', 'B3': 'Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar_(1)'}
- `Annual Statement-211026 1030/Gamma-SAM/Version-211026 1030/ITC_Report_Gamma_SAM_211026_1030.xlsx` sheets=['Links', 'Info']
    - Links: {'A1': 'ITC template - links sheet', 'B2': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Gamma-SAM\\Version-211026 1030\\GSTR3B'}
- `Annual Statement-211026 1030/Gamma-SAM/Version-211026 1030/Sales_Report_Gamma_SAM_211026_1030.xlsx` sheets=['Links', 'Info']
    - Links: {'A1': 'Sales template - links sheet', 'B8': '\\root\\merge-scratch\\run\\A-TRACE\\work\\t6\\target\\Annual Statement-211026 1030\\Gamma-SAM\\Version-211026 1030\\Sales', 'B9': 'SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1'}

`Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/_Organization_Report_Alpha Test Traders_Maharashtra.txt`:
```
GST FILE ORGANIZATION REPORT
==================================================
Generated: 2026-10-21 04:24:54
Client: Alpha Test Traders
State: Maharashtra
Processing Mode: fresh

FOLDER STRUCTURE:
==================================================
VERSION: Alpha Test Traders-MH/Version-211026 1030
GSTR3B: Alpha Test Traders-MH/Version-211026 1030/GSTR3B
ITC: Alpha Test Traders-MH/Version-211026 1030/ITC
SALES: Alpha Test Traders-MH/Version-211026 1030/Sales


FILES ORGANIZED:
==================================================

GSTR-3B Export (4 files):
  - (1)_GSTR3B-Alpha_Test_Traders-Maharashtra-Jun.xlsx
  - GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx
  - GSTR3B-Alpha_Test_Traders-Maharashtra-May.xlsx
  - GSTR3B-Alpha_Test_Traders-Maharashtra-Sep.XLSX

Annual Report (1 files):
  - AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25.xlsx

GSTR-2B Reco (1 files):
  - GSTR-2B-Reco-Alpha_Test_Traders-Maharashtra-Apr24-Mar25.xlsx

IMS Reco (1 files):
  - ImsReco-Alpha_Test_Traders-Maharashtra-31032025.xlsx

Sales (1 files):
  - Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx

Sales Reco (1 files):
  - SalesReco-Alpha_Test_Traders-Maharashtra-FY_2024-25.xlsx


STATISTICS:
==================================================
Total Files: 9
Successfully Copied: 9
Folders Created: 6

```

`Annual Statement-211026 1030/Beta Demo Industries Pvt Ltd-GJ/Version-211026 1030/_Organization_Report_Beta Demo Industries Private Limited_Gujarat.txt`:
```
GST FILE ORGANIZATION REPORT
==================================================
Generated: 2026-10-21 04:24:54
Client: Beta Demo Industries Private Limited
State: Gujarat
Processing Mode: fresh

FOLDER STRUCTURE:
==================================================
VERSION: Beta Demo Industries Pvt Ltd-GJ/Version-211026 1030
GSTR3B: Beta Demo Industries Pvt Ltd-GJ/Version-211026 1030/GSTR3B
ITC: Beta Demo Industries Pvt Ltd-GJ/Version-211026 1030/ITC
SALES: Beta Demo Industries Pvt Ltd-GJ/Version-211026 1030/Sales


FILES ORGANIZED:
==================================================

GSTR-3B Export (1 files):
  - GSTR3B-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr.xlsx

Sales (1 files):
  - Sales-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr-Mar.xls


STATISTICS:
==================================================
Total Files: 2
Successfully Copied: 18
Folders Created: 21

```

`Annual Statement-211026 1030/GAMMA SAMPLE ENTERPRISES LLP-KA/Version-211026 1030/_Organization_Report_GAMMA SAMPLE ENTERPRISES LLP_Karnataka.txt`:
```
GST FILE ORGANIZATION REPORT
==================================================
Generated: 2026-10-21 04:24:54
Client: GAMMA SAMPLE ENTERPRISES LLP
State: Karnataka
Processing Mode: fresh

FOLDER STRUCTURE:
==================================================
VERSION: GAMMA SAMPLE ENTERPRISES LLP-KA/Version-211026 1030
GSTR3B: GAMMA SAMPLE ENTERPRISES LLP-KA/Version-211026 1030/GSTR3B
ITC: GAMMA SAMPLE ENTERPRISES LLP-KA/Version-211026 1030/ITC
SALES: GAMMA SAMPLE ENTERPRISES LLP-KA/Version-211026 1030/Sales


FILES ORGANIZED:
==================================================

Annual Report (1 files):
  - AnnualReport-GAMMA_SAMPLE_ENTERPRISES_LLP-Karnataka-2024-25.xlsx


STATISTICS:
==================================================
Total Files: 1
Successfully Copied: 10
Folders Created: 11

```

`Annual Statement-211026 1030/Gamma Sample Enterprises LLP-KA/Version-211026 1030/_Organization_Report_Gamma Sample Enterprises LLP_Karnataka.txt`:
```
GST FILE ORGANIZATION REPORT
==================================================
Generated: 2026-10-21 04:24:54
Client: Gamma Sample Enterprises LLP
State: Karnataka
Processing Mode: fresh

FOLDER STRUCTURE:
==================================================
VERSION: Gamma Sample Enterprises LLP-KA/Version-211026 1030
GSTR3B: Gamma Sample Enterprises LLP-KA/Version-211026 1030/GSTR3B
ITC: Gamma Sample Enterprises LLP-KA/Version-211026 1030/ITC
SALES: Gamma Sample Enterprises LLP-KA/Version-211026 1030/Sales


FILES ORGANIZED:
==================================================

GSTR-2B Reco (1 files):
  - GSTR-2B-Reco-Gamma_Sample_Enterprises_LLP-Karnataka-Q1.xlsx

GSTR-3B Export (1 files):
  - GSTR3B-Gamma_Sample_Enterprises_LLP-KA-Apr.xlsx

IMS Reco (2 files):
  - ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025_(1).xlsx
  - ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025.xlsx

Sales (2 files):
  - Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar_(1).xlsx
  - Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar.xlsx


STATISTICS:
==================================================
Total Files: 6
Successfully Copied: 16
Folders Created: 16

```

`Annual Statement-211026 1030/Gamma-SAM/Version-211026 1030/_Organization_Report_Gamma_Sample.txt`:
```
GST FILE ORGANIZATION REPORT
==================================================
Generated: 2026-10-21 04:24:54
Client: Gamma
State: Sample
Processing Mode: fresh

FOLDER STRUCTURE:
==================================================
VERSION: Gamma-SAM/Version-211026 1030
GSTR3B: Gamma-SAM/Version-211026 1030/GSTR3B
ITC: Gamma-SAM/Version-211026 1030/ITC
SALES: Gamma-SAM/Version-211026 1030/Sales


FILES ORGANIZED:
==================================================

Sales Reco (1 files):
  - SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1.xlsx


STATISTICS:
==================================================
Total Files: 1
Successfully Copied: 19
Folders Created: 26

```

Summary workbook `Annual Statement-211026 1030/GST_Processing_Summary_211026_1030.xlsx` sheets: ['Summary', 'Client Status', 'File Mapping', 'Errors', 'Variations']

###### Summary
    GST FILE PROCESSING SUMMARY | 
    Processing Date | 211026 1030
    Total Clients | 5
    Successful | 5
    Failed | 0
    Total Files | 19
    Reports Generated | 10
    ITC Reports Created | 5
    Sales Reports Created | 5
    Report Creation Errors | 0
    Processing Mode | fresh
    Target Folder | /root/merge-scratch/run/A-TRACE/work/t6/target
    Include Client Name in Folders | No

###### Client Status
    Client | State | Status | Files | Missing | Complete %
    Alpha Test Traders | Maharashtra | Success | 9 | 0 | 100.0%
    GAMMA SAMPLE ENTERPRISES LLP | Karnataka | Success | 1 | 5 | 16.7%
    Gamma Sample Enterprises LLP | Karnataka | Success | 6 | 2 | 66.7%
    Beta Demo Industries Private Limited | Gujarat | Success | 2 | 4 | 33.3%
    Gamma | Sample | Success | 1 | 5 | 16.7%

###### File Mapping
    Original File | Client | Type | Destination | Status
    (1) GSTR3B-Alpha Test Traders-Maharashtra-Jun.xlsx | Alpha Test Traders-MH | GSTR-3B Export | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/GSTR3B/(1)_GSTR3B-Alpha_Test_Traders-Maharashtra-Jun.xlsx | Success
    GSTR3B-Alpha Test Traders-Maharashtra-Apr.xlsx | Alpha Test Traders-MH | GSTR-3B Export | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/GSTR3B/GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx | Success
    GSTR3B-Alpha Test Traders-Maharashtra-May.xlsx | Alpha Test Traders-MH | GSTR-3B Export | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/GSTR3B/GSTR3B-Alpha_Test_Traders-Maharashtra-May.xlsx | Success
    GSTR3B-Alpha Test Traders-Maharashtra-Sep.XLSX | Alpha Test Traders-MH | GSTR-3B Export | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/GSTR3B/GSTR3B-Alpha_Test_Traders-Maharashtra-Sep.XLSX | Success
    AnnualReport-Alpha Test Traders-Maharashtra-2024-25.xlsx | Alpha Test Traders-MH | Annual Report | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25.xlsx | Success
    GSTR-2B-Reco-Alpha Test Traders-Maharashtra-Apr24-Mar25.xlsx | Alpha Test Traders-MH | GSTR-2B Reco | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/ITC/GSTR-2B-Reco-Alpha_Test_Traders-Maharashtra-Apr24-Mar25.xlsx | Success
    ImsReco-Alpha Test Traders-Maharashtra-31032025.xlsx | Alpha Test Traders-MH | IMS Reco | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/ITC/ImsReco-Alpha_Test_Traders-Maharashtra-31032025.xlsx | Success
    Sales-Alpha Test Traders-Maharashtra-Apr-Mar.xlsx | Alpha Test Traders-MH | Sales | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/Sales/Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx | Success
    SalesReco-Alpha Test Traders-Maharashtra-FY 2024-25.xlsx | Alpha Test Traders-MH | Sales Reco | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/Sales/SalesReco-Alpha_Test_Traders-Maharashtra-FY_2024-25.xlsx | Success
    AnnualReport-GAMMA SAMPLE ENTERPRISES LLP-Karnataka-2024-25.xlsx | GAMMA SAMPLE ENTERPRISES LLP-KA | Annual Report | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/GAMMA SAMPLE ENTERPRISES LLP-KA/Version-211026 1030/AnnualReport-GAMMA_SAMPLE_ENTERPRISES_LLP-Karnataka-2024-25.xlsx | Success
    GSTR-2B-Reco-Gamma Sample Enterprises LLP-Karnataka-Q1.xlsx | Gamma Sample Enterprises LLP-KA | GSTR-2B Reco | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Gamma Sample Enterprises LLP-KA/Version-211026 1030/ITC/GSTR-2B-Reco-Gamma_Sample_Enterprises_LLP-Karnataka-Q1.xlsx | Success
    GSTR3B-Gamma Sample Enterprises LLP-KA-Apr.xlsx | Gamma Sample Enterprises LLP-KA | GSTR-3B Export | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Gamma Sample Enterprises LLP-KA/Version-211026 1030/GSTR3B/GSTR3B-Gamma_Sample_Enterprises_LLP-KA-Apr.xlsx | Success
    ImsReco-Gamma Sample Enterprises LLP-Karnataka-31032025 (1).xlsx | Gamma Sample Enterprises LLP-KA | IMS Reco | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Gamma Sample Enterprises LLP-KA/Version-211026 1030/ITC/ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025_(1).xlsx | Success
    ImsReco-Gamma Sample Enterprises LLP-Karnataka-31032025.xlsx | Gamma Sample Enterprises LLP-KA | IMS Reco | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Gamma Sample Enterprises LLP-KA/Version-211026 1030/ITC/ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025.xlsx | Success
    Sales-Gamma Sample Enterprises LLP-Karnataka-Apr-Mar (1).xlsx | Gamma Sample Enterprises LLP-KA | Sales | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Gamma Sample Enterprises LLP-KA/Version-211026 1030/Sales/Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar_(1).xlsx | Success
    Sales-Gamma Sample Enterprises LLP-Karnataka-Apr-Mar.xlsx | Gamma Sample Enterprises LLP-KA | Sales | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Gamma Sample Enterprises LLP-KA/Version-211026 1030/Sales/Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar.xlsx | Success
    GSTR3B-Beta Demo Industries Private Limited-Gujarat-Apr.xlsx | Beta Demo Industries Private Limited-GJ | GSTR-3B Export | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Beta Demo Industries Pvt Ltd-GJ/Version-211026 1030/GSTR3B/GSTR3B-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr.xlsx | Success
    Sales-Beta Demo Industries Private Limited-Gujarat-Apr-Mar.xls | Beta Demo Industries Private Limited-GJ | Sales | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Beta Demo Industries Pvt Ltd-GJ/Version-211026 1030/Sales/Sales-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr-Mar.xls | Success
    SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1.xlsx | Gamma-SAM | Sales Reco | /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Gamma-SAM/Version-211026 1030/Sales/SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1.xlsx | Success

###### Errors
    Time | Client | Type | Error | Action
    No errors encountered |  |  |  | 

###### Variations
    Filename | Issue | Expected Pattern | Size | Action
    Gamma GSTR3B Jul.xlsm | No matching pattern found | GSTR3B-ClientName-State-Month.xlsx | 4849 | Rename file
    Gamma notes.xlsx | No matching pattern found | Check file naming convention | 4849 | Rename file

##### client_data after run 1 (organize_files renames entries in place, file_organizer.py:233-246)

- Alpha Test Traders-MH: [('(1)_GSTR3B-Alpha_Test_Traders-Maharashtra-Jun.xlsx', '(1)_GSTR3B-Alpha_Test_Traders-Maharashtra-Jun.xlsx'), ('GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx', 'GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx'), ('GSTR3B-Alpha_Test_Traders-Maharashtra-May.xlsx', 'GSTR3B-Alpha_Test_Traders-Maharashtra-May.xlsx'), ('GSTR3B-Alpha_Test_Traders-Maharashtra-Sep.XLSX', 'GSTR3B-Alpha_Test_Traders-Maharashtra-Sep.XLSX'), ('AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25.xlsx', 'AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25.xlsx'), ('GSTR-2B-Reco-Alpha_Test_Traders-Maharashtra-Apr24-Mar25.xlsx', 'GSTR-2B-Reco-Alpha_Test_Traders-Maharashtra-Apr24-Mar25.xlsx'), ('ImsReco-Alpha_Test_Traders-Maharashtra-31032025.xlsx', 'ImsReco-Alpha_Test_Traders-Maharashtra-31032025.xlsx'), ('Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx', 'Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx'), ('SalesReco-Alpha_Test_Traders-Maharashtra-FY_2024-25.xlsx', 'SalesReco-Alpha_Test_Traders-Maharashtra-FY_2024-25.xlsx')]
- GAMMA SAMPLE ENTERPRISES LLP-KA: [('AnnualReport-GAMMA_SAMPLE_ENTERPRISES_LLP-Karnataka-2024-25.xlsx', 'AnnualReport-GAMMA_SAMPLE_ENTERPRISES_LLP-Karnataka-2024-25.xlsx')]
- Gamma Sample Enterprises LLP-KA: [('GSTR-2B-Reco-Gamma_Sample_Enterprises_LLP-Karnataka-Q1.xlsx', 'GSTR-2B-Reco-Gamma_Sample_Enterprises_LLP-Karnataka-Q1.xlsx'), ('GSTR3B-Gamma_Sample_Enterprises_LLP-KA-Apr.xlsx', 'GSTR3B-Gamma_Sample_Enterprises_LLP-KA-Apr.xlsx'), ('ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025_(1).xlsx', 'ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025_(1).xlsx'), ('ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025.xlsx', 'ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025.xlsx'), ('Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar_(1).xlsx', 'Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar_(1).xlsx'), ('Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar.xlsx', 'Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar.xlsx')]
- Beta Demo Industries Private Limited-GJ: [('GSTR3B-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr.xlsx', 'GSTR3B-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr.xlsx'), ('Sales-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr-Mar.xls', 'Sales-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr-Mar.xls')]
- Gamma-SAM: [('SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1.xlsx', 'SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1.xlsx')]

#### 4. Source tree after run 1 vs before

identical (names, sha256, size, mtime_ns, plus directory mtimes): True
differences: {}

#### 5. Running again


##### Run 2 - same minute, fresh: mode='fresh', timestamp='211026 1030', clients=['Alpha Test Traders-MH']

GUI log (level, message):
```
[info] 
🏢 Processing Alpha Test Traders - Maharashtra
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1030
[normal]     Level 2: Alpha Test Traders-MH
[normal]     Level 3: Version-211026 1030
[success]   ✓ Organized 9 files
[normal]     📁 (1)_GSTR3B-Alpha_Test_Traders-Maharashtra-Jun.xlsx → GSTR-3B Export folder
[normal]     📁 GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx → GSTR-3B Export folder
[normal]     📁 GSTR3B-Alpha_Test_Traders-Maharashtra-May.xlsx → GSTR-3B Export folder
[normal]     📁 GSTR3B-Alpha_Test_Traders-Maharashtra-Sep.XLSX → GSTR-3B Export folder
[normal]     📁 AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25.xlsx → Annual Report folder
[normal]     📁 GSTR-2B-Reco-Alpha_Test_Traders-Maharashtra-Apr24-Mar25.xlsx → GSTR-2B Reco folder
[normal]     📁 ImsReco-Alpha_Test_Traders-Maharashtra-31032025.xlsx → IMS Reco folder
[normal]     📁 Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx → Sales folder
[normal]     📁 SalesReco-Alpha_Test_Traders-Maharashtra-FY_2024-25.xlsx → Sales Reco folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[success]     ✅ Success: ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[normal]     [DEBUG] Output file created, size: 5458 bytes
[success]   ✅ Completed Alpha Test Traders
[info] 
📊 Creating Level 1 summary report...
[normal]   Output: GST_Processing_Summary_211026_1030.xlsx
[success]   ✅ Success: GST_Processing_Summary_211026_1030.xlsx
[normal]   Sheets: Summary, Client Status, File Mapping, Errors, Variations
[success] 
🎉 ════════ PROCESSING COMPLETE ════════
[info] ✅ Processed: 1/1 clients
[info] 📁 Files organized: 9
[info] 📊 Reports created: 2
[info] 📂 Output location: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030
[info] 
⏱️ Total processing time: 0 minutes 0 seconds
[success] 📂 Opened output folder
[info] ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query.
```
progress updates: 14 (last: (100, 'Processing complete | Total time: 0m 0s'))
pop-ups: [('info', ('Processing Complete', 'Successfully processed 1/1 clients\n\nFiles organized in:\n/root/merge-scratch/run/A-TRACE/work/t6/target'))]; os.system calls: ['xdg-open "/root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030"']
logger WARNING/ERROR: [('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!')]
file_organizer.get_summary(): {'timestamp': '211026 1030', 'processing_mode': 'fresh', 'target_folder': '/root/merge-scratch/run/A-TRACE/work/t6/target', 'folders_created': 6, 'files_processed': 9, 'successful_copies': 9, 'failed_copies': 0, 'skipped_files': 0, 'errors': 0}

##### Run 3 - same minute, rerun: mode='rerun', timestamp='211026 1030', clients=['Alpha Test Traders-MH']

GUI log (level, message):
```
[info] 
🏢 Processing Alpha Test Traders - Maharashtra
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1030
[normal]     Level 2: Alpha Test Traders-MH
[normal]     Level 3: Version-211026 1030
[success]   ✓ Organized 9 files
[normal]     📁 (1)_GSTR3B-Alpha_Test_Traders-Maharashtra-Jun.xlsx → GSTR-3B Export folder
[normal]     📁 GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx → GSTR-3B Export folder
[normal]     📁 GSTR3B-Alpha_Test_Traders-Maharashtra-May.xlsx → GSTR-3B Export folder
[normal]     📁 GSTR3B-Alpha_Test_Traders-Maharashtra-Sep.XLSX → GSTR-3B Export folder
[normal]     📁 AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25.xlsx → Annual Report folder
[normal]     📁 GSTR-2B-Reco-Alpha_Test_Traders-Maharashtra-Apr24-Mar25.xlsx → GSTR-2B Reco folder
[normal]     📁 ImsReco-Alpha_Test_Traders-Maharashtra-31032025.xlsx → IMS Reco folder
[normal]     📁 Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx → Sales folder
[normal]     📁 SalesReco-Alpha_Test_Traders-Maharashtra-FY_2024-25.xlsx → Sales Reco folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[success]     ✅ Success: ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[normal]     [DEBUG] Output file created, size: 5458 bytes
[success]   ✅ Completed Alpha Test Traders
[info] 
📊 Creating Level 1 summary report...
[normal]   Output: GST_Processing_Summary_211026_1030.xlsx
[success]   ✅ Success: GST_Processing_Summary_211026_1030.xlsx
[normal]   Sheets: Summary, Client Status, File Mapping, Errors, Variations
[success] 
🎉 ════════ PROCESSING COMPLETE ════════
[info] ✅ Processed: 1/1 clients
[info] 📁 Files organized: 9
[info] 📊 Reports created: 2
[info] 📂 Output location: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030
[info] 
⏱️ Total processing time: 0 minutes 0 seconds
[success] 📂 Opened output folder
[info] ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query.
```
progress updates: 14 (last: (100, 'Processing complete | Total time: 0m 0s'))
pop-ups: [('info', ('Processing Complete', 'Successfully processed 1/1 clients\n\nFiles organized in:\n/root/merge-scratch/run/A-TRACE/work/t6/target'))]; os.system calls: ['xdg-open "/root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030"']
logger WARNING/ERROR: [('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!')]
file_organizer.get_summary(): {'timestamp': '211026 1030', 'processing_mode': 'rerun', 'target_folder': '/root/merge-scratch/run/A-TRACE/work/t6/target', 'folders_created': 5, 'files_processed': 9, 'successful_copies': 9, 'failed_copies': 0, 'skipped_files': 0, 'errors': 0}

##### Run 4 - same minute, resume: mode='resume', timestamp='211026 1030', clients=['Alpha Test Traders-MH']

GUI log (level, message):
```
[info] 
🏢 Processing Alpha Test Traders - Maharashtra
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1030
[normal]     Level 2: Alpha Test Traders-MH
[normal]     Level 3: Version-211026 1030
[success]   ✓ Organized 0 files
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[success]     ✅ Success: ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1030/Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx
[normal]     [DEBUG] Output file created, size: 5458 bytes
[success]   ✅ Completed Alpha Test Traders
[info] 
📊 Creating Level 1 summary report...
[normal]   Output: GST_Processing_Summary_211026_1030.xlsx
[success]   ✅ Success: GST_Processing_Summary_211026_1030.xlsx
[normal]   Sheets: Summary, Client Status, File Mapping, Errors, Variations
[success] 
🎉 ════════ PROCESSING COMPLETE ════════
[info] ✅ Processed: 1/1 clients
[info] 📁 Files organized: 0
[info] 📊 Reports created: 2
[info] 📂 Output location: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030
[info] 
⏱️ Total processing time: 0 minutes 0 seconds
[success] 📂 Opened output folder
[info] ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query.
```
progress updates: 14 (last: (100, 'Processing complete | Total time: 0m 0s'))
pop-ups: [('info', ('Processing Complete', 'Successfully processed 1/1 clients\n\nFiles organized in:\n/root/merge-scratch/run/A-TRACE/work/t6/target'))]; os.system calls: ['xdg-open "/root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030"']
logger WARNING/ERROR: [('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!')]
file_organizer.get_summary(): {'timestamp': '211026 1030', 'processing_mode': 'resume', 'target_folder': '/root/merge-scratch/run/A-TRACE/work/t6/target', 'folders_created': 5, 'files_processed': 9, 'successful_copies': 0, 'failed_copies': 0, 'skipped_files': 9, 'errors': 0}

##### Run 5 - later minute, resume: mode='resume', timestamp='211026 1045', clients=['Alpha Test Traders-MH']

GUI log (level, message):
```
[info] 
🏢 Processing Alpha Test Traders - Maharashtra
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1030
[normal]     Level 2: Alpha Test Traders-MH
[normal]     Level 3: Version-211026 1045
[success]   ✓ Organized 9 files
[normal]     📁 (1)_GSTR3B-Alpha_Test_Traders-Maharashtra-Jun.xlsx → GSTR-3B Export folder
[normal]     📁 GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx → GSTR-3B Export folder
[normal]     📁 GSTR3B-Alpha_Test_Traders-Maharashtra-May.xlsx → GSTR-3B Export folder
[normal]     📁 GSTR3B-Alpha_Test_Traders-Maharashtra-Sep.XLSX → GSTR-3B Export folder
[normal]     📁 AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25.xlsx → Annual Report folder
[normal]     📁 GSTR-2B-Reco-Alpha_Test_Traders-Maharashtra-Apr24-Mar25.xlsx → GSTR-2B Reco folder
[normal]     📁 ImsReco-Alpha_Test_Traders-Maharashtra-31032025.xlsx → IMS Reco folder
[normal]     📁 Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx → Sales folder
[normal]     📁 SalesReco-Alpha_Test_Traders-Maharashtra-FY_2024-25.xlsx → Sales Reco folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_Alpha_Test_Traders_MH_211026_1045.xlsx
[success]     ✅ Success: ITC_Report_Alpha_Test_Traders_MH_211026_1045.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_Alpha_Test_Traders_MH_211026_1045.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/Alpha Test Traders-MH/Version-211026 1045/Sales_Report_Alpha_Test_Traders_MH_211026_1045.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_Alpha_Test_Traders_MH_211026_1045.xlsx
[normal]     [DEBUG] Output file created, size: 5459 bytes
[success]   ✅ Completed Alpha Test Traders
[info] 
📊 Creating Level 1 summary report...
[normal]   Output: GST_Processing_Summary_211026_1045.xlsx
[success]   ✅ Success: GST_Processing_Summary_211026_1045.xlsx
[normal]   Sheets: Summary, Client Status, File Mapping, Errors, Variations
[success] 
🎉 ════════ PROCESSING COMPLETE ════════
[info] ✅ Processed: 1/1 clients
[info] 📁 Files organized: 9
[info] 📊 Reports created: 2
[info] 📂 Output location: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030
[info] 
⏱️ Total processing time: 0 minutes 0 seconds
[success] 📂 Opened output folder
[info] ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query.
```
progress updates: 14 (last: (100, 'Processing complete | Total time: 0m 0s'))
pop-ups: [('info', ('Processing Complete', 'Successfully processed 1/1 clients\n\nFiles organized in:\n/root/merge-scratch/run/A-TRACE/work/t6/target'))]; os.system calls: ['xdg-open "/root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030"']
logger WARNING/ERROR: [('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!')]
file_organizer.get_summary(): {'timestamp': '211026 1045', 'processing_mode': 'resume', 'target_folder': '/root/merge-scratch/run/A-TRACE/work/t6/target', 'folders_created': 5, 'files_processed': 9, 'successful_copies': 9, 'failed_copies': 0, 'skipped_files': 0, 'errors': 0}

##### Run 6 - later minute, rerun: mode='rerun', timestamp='211026 1050', clients=['GAMMA SAMPLE ENTERPRISES LLP-KA']

GUI log (level, message):
```
[info] 
🏢 Processing GAMMA SAMPLE ENTERPRISES LLP - Karnataka
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1030
[normal]     Level 2: GAMMA SAMPLE ENTERPRISES LLP-KA
[normal]     Level 3: Version-211026 1050
[success]   ✓ Organized 1 files
[normal]     📁 AnnualReport-GAMMA_SAMPLE_ENTERPRISES_LLP-Karnataka-2024-25.xlsx → Annual Report folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1050.xlsx
[success]     ✅ Success: ITC_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1050.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1050.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030/GAMMA SAMPLE ENTERPRISES LLP-KA/Version-211026 1050/Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1050.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1050.xlsx
[normal]     [DEBUG] Output file created, size: 5419 bytes
[success]   ✅ Completed GAMMA SAMPLE ENTERPRISES LLP
[info] 
📊 Creating Level 1 summary report...
[normal]   Output: GST_Processing_Summary_211026_1050.xlsx
[success]   ✅ Success: GST_Processing_Summary_211026_1050.xlsx
[normal]   Sheets: Summary, Client Status, File Mapping, Errors, Variations
[success] 
🎉 ════════ PROCESSING COMPLETE ════════
[info] ✅ Processed: 1/1 clients
[info] 📁 Files organized: 1
[info] 📊 Reports created: 2
[info] 📂 Output location: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030
[info] 
⏱️ Total processing time: 0 minutes 0 seconds
[success] 📂 Opened output folder
[info] ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query.
```
progress updates: 6 (last: (100, 'Processing complete | Total time: 0m 0s'))
pop-ups: [('info', ('Processing Complete', 'Successfully processed 1/1 clients\n\nFiles organized in:\n/root/merge-scratch/run/A-TRACE/work/t6/target'))]; os.system calls: ['xdg-open "/root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1030"']
logger WARNING/ERROR: [('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!')]
file_organizer.get_summary(): {'timestamp': '211026 1050', 'processing_mode': 'rerun', 'target_folder': '/root/merge-scratch/run/A-TRACE/work/t6/target', 'folders_created': 5, 'files_processed': 1, 'successful_copies': 1, 'failed_copies': 0, 'skipped_files': 0, 'errors': 0}

Target tree after runs 2-6:
```
Annual Statement-211026 1030/
  Alpha Test Traders-MH/
    Version-211026 1030/
      AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25.xlsx  (4849 B)
      AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25_backup_211026_0424.xlsx  (4849 B)
      GSTR3B/
        (1)_GSTR3B-Alpha_Test_Traders-Maharashtra-Jun.xlsx  (4849 B)
        (1)_GSTR3B-Alpha_Test_Traders-Maharashtra-Jun_backup_211026_0424.xlsx  (4849 B)
        GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx  (4849 B)
        GSTR3B-Alpha_Test_Traders-Maharashtra-Apr_backup_211026_0424.xlsx  (4849 B)
        GSTR3B-Alpha_Test_Traders-Maharashtra-May.xlsx  (4849 B)
        GSTR3B-Alpha_Test_Traders-Maharashtra-May_backup_211026_0424.xlsx  (4849 B)
        GSTR3B-Alpha_Test_Traders-Maharashtra-Sep.XLSX  (4849 B)
        GSTR3B-Alpha_Test_Traders-Maharashtra-Sep_backup_211026_0424.XLSX  (4849 B)
      ITC/
        GSTR-2B-Reco-Alpha_Test_Traders-Maharashtra-Apr24-Mar25.xlsx  (4849 B)
        GSTR-2B-Reco-Alpha_Test_Traders-Maharashtra-Apr24-Mar25_backup_211026_0424.xlsx  (4849 B)
        ImsReco-Alpha_Test_Traders-Maharashtra-31032025.xlsx  (4849 B)
        ImsReco-Alpha_Test_Traders-Maharashtra-31032025_backup_211026_0424.xlsx  (4849 B)
      ITC_Report_Alpha_Test_Traders_MH_211026_1030.xlsx  (5483 B)
      Sales/
        Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx  (4849 B)
        Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar_backup_211026_0424.xlsx  (4849 B)
        SalesReco-Alpha_Test_Traders-Maharashtra-FY_2024-25.xlsx  (4849 B)
        SalesReco-Alpha_Test_Traders-Maharashtra-FY_2024-25_backup_211026_0424.xlsx  (4849 B)
      Sales_Report_Alpha_Test_Traders_MH_211026_1030.xlsx  (5458 B)
      _Organization_Report_Alpha Test Traders_Maharashtra.txt  (1303 B)
    Version-211026 1045/
      AnnualReport-Alpha_Test_Traders-Maharashtra-2024-25.xlsx  (4849 B)
      GSTR3B/
        (1)_GSTR3B-Alpha_Test_Traders-Maharashtra-Jun.xlsx  (4849 B)
        GSTR3B-Alpha_Test_Traders-Maharashtra-Apr.xlsx  (4849 B)
        GSTR3B-Alpha_Test_Traders-Maharashtra-May.xlsx  (4849 B)
        GSTR3B-Alpha_Test_Traders-Maharashtra-Sep.XLSX  (4849 B)
      ITC/
        GSTR-2B-Reco-Alpha_Test_Traders-Maharashtra-Apr24-Mar25.xlsx  (4849 B)
        ImsReco-Alpha_Test_Traders-Maharashtra-31032025.xlsx  (4849 B)
      ITC_Report_Alpha_Test_Traders_MH_211026_1045.xlsx  (5484 B)
      Sales/
        Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx  (4849 B)
        SalesReco-Alpha_Test_Traders-Maharashtra-FY_2024-25.xlsx  (4849 B)
      Sales_Report_Alpha_Test_Traders_MH_211026_1045.xlsx  (5459 B)
      _Organization_Report_Alpha Test Traders_Maharashtra.txt  (1303 B)
  Beta Demo Industries Pvt Ltd-GJ/
    Version-211026 1030/
      GSTR3B/
        GSTR3B-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr.xlsx  (4849 B)
      ITC/
      ITC_Report_Beta_Demo_Industries_Pvt_Ltd_GJ_211026_1030.xlsx  (5377 B)
      Sales/
        Sales-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr-Mar.xls  (4096 B)
      Sales_Report_Beta_Demo_Industries_Pvt_Ltd_GJ_211026_1030.xlsx  (5410 B)
      _Organization_Report_Beta Demo Industries Private Limited_Gujarat.txt  (873 B)
  GAMMA SAMPLE ENTERPRISES LLP-KA/
    Version-211026 1030/
      AnnualReport-GAMMA_SAMPLE_ENTERPRISES_LLP-Karnataka-2024-25.xlsx  (4849 B)
      GSTR3B/
      ITC/
      ITC_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1030.xlsx  (5432 B)
      Sales/
      Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1030.xlsx  (5417 B)
      _Organization_Report_GAMMA SAMPLE ENTERPRISES LLP_Karnataka.txt  (801 B)
    Version-211026 1050/
      AnnualReport-GAMMA_SAMPLE_ENTERPRISES_LLP-Karnataka-2024-25.xlsx  (4849 B)
      GSTR3B/
      ITC/
      ITC_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1050.xlsx  (5434 B)
      Sales/
      Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1050.xlsx  (5419 B)
      _Organization_Report_GAMMA SAMPLE ENTERPRISES LLP_Karnataka.txt  (799 B)
  GST_Processing_Summary_211026_1030.xlsx  (8489 B)
  GST_Processing_Summary_211026_1045.xlsx  (8490 B)
  GST_Processing_Summary_211026_1050.xlsx  (8172 B)
  Gamma Sample Enterprises LLP-KA/
    Version-211026 1030/
      GSTR3B/
        GSTR3B-Gamma_Sample_Enterprises_LLP-KA-Apr.xlsx  (4849 B)
      ITC/
        GSTR-2B-Reco-Gamma_Sample_Enterprises_LLP-Karnataka-Q1.xlsx  (4849 B)
        ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025.xlsx  (4849 B)
        ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025_(1).xlsx  (4849 B)
      ITC_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1030.xlsx  (5464 B)
      Sales/
        Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar.xlsx  (4849 B)
        Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar_(1).xlsx  (4849 B)
      Sales_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1030.xlsx  (5414 B)
      _Organization_Report_Gamma Sample Enterprises LLP_Karnataka.txt  (1175 B)
  Gamma-SAM/
    Version-211026 1030/
      GSTR3B/
      ITC/
      ITC_Report_Gamma_SAM_211026_1030.xlsx  (5360 B)
      Sales/
        SalesReco-Gamma-Sample-Enterprises-LLP-Karnataka-Q1.xlsx  (4849 B)
      Sales_Report_Gamma_SAM_211026_1030.xlsx  (5401 B)
      _Organization_Report_Gamma_Sample.txt  (676 B)
```

Source tree still identical after all runs: True

#### 6. 'Include client name in folders' ON (file_organizer.py:117-139)


##### Run 7 - include_client_name=True: mode='fresh', timestamp='211026 1100', clients=['GAMMA SAMPLE ENTERPRISES LLP-KA']

GUI log (level, message):
```
[info] 
🏢 Processing GAMMA SAMPLE ENTERPRISES LLP - Karnataka
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1100
[normal]     Level 2: GAMMA SAMPLE ENTERPRISES LLP-KA
[normal]     Level 3: Version-211026 1100
[success]   ✓ Organized 1 files
[normal]     📁 AnnualReport-GAMMA_SAMPLE_ENTERPRISES_LLP-Karnataka-2024-25.xlsx → Annual Report folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1100.xlsx
[success]     ✅ Success: ITC_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1100.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1100.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1100/GAMMA SAMPLE ENTERPRISES LLP-KA/Version-211026 1100/Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1100.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1100.xlsx
[normal]     [DEBUG] Output file created, size: 5415 bytes
[success]   ✅ Completed GAMMA SAMPLE ENTERPRISES LLP
[info] 
📊 Creating Level 1 summary report...
[normal]   Output: GST_Processing_Summary_211026_1100.xlsx
[success]   ✅ Success: GST_Processing_Summary_211026_1100.xlsx
[normal]   Sheets: Summary, Client Status, File Mapping, Errors, Variations
[success] 
🎉 ════════ PROCESSING COMPLETE ════════
[info] ✅ Processed: 1/1 clients
[info] 📁 Files organized: 1
[info] 📊 Reports created: 2
[info] 📂 Output location: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1100
[info] 
⏱️ Total processing time: 0 minutes 0 seconds
[success] 📂 Opened output folder
[info] ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query.
```
progress updates: 6 (last: (100, 'Processing complete | Total time: 0m 0s'))
pop-ups: [('info', ('Processing Complete', 'Successfully processed 1/1 clients\n\nFiles organized in:\n/root/merge-scratch/run/A-TRACE/work/t6/target'))]; os.system calls: ['xdg-open "/root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1100"']
logger WARNING/ERROR: [('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!')]
file_organizer.get_summary(): {'timestamp': '211026 1100', 'processing_mode': 'fresh', 'target_folder': '/root/merge-scratch/run/A-TRACE/work/t6/target', 'folders_created': 6, 'files_processed': 1, 'successful_copies': 1, 'failed_copies': 0, 'skipped_files': 0, 'errors': 0}

```
GAMMA SAMPLE ENTERPRISES LLP-KA/
  Version-211026 1100/
    AnnualReport-GAMMA_SAMPLE_ENTERPRISES_LLP-Karnataka-2024-25.xlsx  (4849 B)
    GSTR3B (GAMMA_SAMPLE_ENTERPRISES_LLP)/
    ITC (GAMMA_SAMPLE_ENTERPRISES_LLP)/
    ITC_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1100.xlsx  (5440 B)
    Sales (GAMMA_SAMPLE_ENTERPRISES_LLP)/
    Sales_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_211026_1100.xlsx  (5415 B)
    _Organization_Report_GAMMA SAMPLE ENTERPRISES LLP_Karnataka.txt  (892 B)
GST_Processing_Summary_211026_1100.xlsx  (8172 B)
```

Run 8 - global OFF, per-client setting keyed by the parser key:

##### Run 8 - per-client folder setting: mode='fresh', timestamp='211026 1105', clients=['Gamma Sample Enterprises LLP-KA']

GUI log (level, message):
```
[info] 
🏢 Processing Gamma Sample Enterprises LLP - Karnataka
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1105
[normal]     Level 2: Gamma Sample Enterprises LLP-KA
[normal]     Level 3: Version-211026 1105
[success]   ✓ Organized 6 files
[normal]     📁 GSTR-2B-Reco-Gamma_Sample_Enterprises_LLP-Karnataka-Q1.xlsx → GSTR-2B Reco folder
[normal]     📁 GSTR3B-Gamma_Sample_Enterprises_LLP-KA-Apr.xlsx → GSTR-3B Export folder
[normal]     📁 ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025_(1).xlsx → IMS Reco folder
[normal]     📁 ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025.xlsx → IMS Reco folder
[normal]     📁 Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar_(1).xlsx → Sales folder
[normal]     📁 Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar.xlsx → Sales folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1105.xlsx
[success]     ✅ Success: ITC_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1105.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1105.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1105/Gamma Sample Enterprises LLP-KA/Version-211026 1105/Sales_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1105.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1105.xlsx
[normal]     [DEBUG] Output file created, size: 5417 bytes
[success]   ✅ Completed Gamma Sample Enterprises LLP
[info] 
📊 Creating Level 1 summary report...
[normal]   Output: GST_Processing_Summary_211026_1105.xlsx
[success]   ✅ Success: GST_Processing_Summary_211026_1105.xlsx
[normal]   Sheets: Summary, Client Status, File Mapping, Errors, Variations
[success] 
🎉 ════════ PROCESSING COMPLETE ════════
[info] ✅ Processed: 1/1 clients
[info] 📁 Files organized: 6
[info] 📊 Reports created: 2
[info] 📂 Output location: /root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1105
[info] 
⏱️ Total processing time: 0 minutes 0 seconds
[success] 📂 Opened output folder
[info] ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query.
```
progress updates: 11 (last: (100, 'Processing complete | Total time: 0m 0s'))
pop-ups: [('info', ('Processing Complete', 'Successfully processed 1/1 clients\n\nFiles organized in:\n/root/merge-scratch/run/A-TRACE/work/t6/target'))]; os.system calls: ['xdg-open "/root/merge-scratch/run/A-TRACE/work/t6/target/Annual Statement-211026 1105"']
logger WARNING/ERROR: [('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!')]
file_organizer.get_summary(): {'timestamp': '211026 1105', 'processing_mode': 'fresh', 'target_folder': '/root/merge-scratch/run/A-TRACE/work/t6/target', 'folders_created': 6, 'files_processed': 6, 'successful_copies': 6, 'failed_copies': 0, 'skipped_files': 0, 'errors': 0}

```
GST_Processing_Summary_211026_1105.xlsx  (8394 B)
Gamma Sample Enterprises LLP-KA/
  Version-211026 1105/
    GSTR3B (Gamma_Sample_Enterprises_LLP)/
      GSTR3B-Gamma_Sample_Enterprises_LLP-KA-Apr.xlsx  (4849 B)
    ITC (Gamma_Sample_Enterprises_LLP)/
      GSTR-2B-Reco-Gamma_Sample_Enterprises_LLP-Karnataka-Q1.xlsx  (4849 B)
      ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025.xlsx  (4849 B)
      ImsReco-Gamma_Sample_Enterprises_LLP-Karnataka-31032025_(1).xlsx  (4849 B)
    ITC_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1105.xlsx  (5471 B)
    Sales (Gamma_Sample_Enterprises_LLP)/
      Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar.xlsx  (4849 B)
      Sales-Gamma_Sample_Enterprises_LLP-Karnataka-Apr-Mar_(1).xlsx  (4849 B)
    Sales_Report_Gamma_Sample_Enterprises_LLP_KA_211026_1105.xlsx  (5417 B)
    _Organization_Report_Gamma Sample Enterprises LLP_Karnataka.txt  (1266 B)
```

#### 7. Two different long client names that share their first 32 characters

client keys: ['Gamma Sample Manufacturing Company of India One-MH', 'Gamma Sample Manufacturing Company of India Two-MH']

##### Run 9 - collision: mode='fresh', timestamp='211026 1110', clients=['Gamma Sample Manufacturing Company of India One-MH', 'Gamma Sample Manufacturing Company of India Two-MH']

GUI log (level, message):
```
[info] 
🏢 Processing Gamma Sample Manufacturing Company of India One - Maharashtra
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1110
[normal]     Level 2: Gamma Sample Manufacturing Compa-MH
[normal]     Level 3: Version-211026 1110
[success]   ✓ Organized 2 files
[normal]     📁 AnnualReport-Gamma Sample Manufacturing Company of India One-Maharashtra-2024-25.xlsx → Annual Report folder
[normal]     📁 Sales-Gamma Sample Manufacturing Company of India One-Maharashtra-Apr-Mar.xlsx → Sales folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110.xlsx
[success]     ✅ Success: ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target2/Annual Statement-211026 1110/Gamma Sample Manufacturing Compa-MH/Version-211026 1110/Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110.xlsx
[normal]     [DEBUG] Output file created, size: 5460 bytes
[success]   ✅ Completed Gamma Sample Manufacturing Company of India One
[info] 
🏢 Processing Gamma Sample Manufacturing Company of India Two - Maharashtra
[success]   ✓ Created folder structure
[normal]     Level 1: Annual Statement-211026 1110
[normal]     Level 2: Gamma Sample Manufacturing Compa-MH
[normal]     Level 3: Version-211026 1110
[success]   ✓ Organized 2 files
[normal]     📁 AnnualReport-Gamma Sample Manufacturing Company of India Two-Maharashtra-2024-25.xlsx → Annual Report folder
[normal]     📁 Sales-Gamma Sample Manufacturing Company of India Two-Maharashtra-Apr-Mar.xlsx → Sales folder
[info]   📊 Creating ITC report...
[normal]     Template: ITC_Template.xlsx
[normal]     Output: ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110.xlsx
[success]     ✅ Success: ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110.xlsx
[info]   💰 Starting Sales report creation...
[normal]     [DEBUG] Sales template path: /root/merge-scratch/run/A-TRACE/work/t6/templates/Sales_Template.xlsx
[normal]     [DEBUG] Template file exists: True
[normal]     Template: Sales_Template.xlsx
[normal]     Output: Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110.xlsx
[normal]     [DEBUG] Full output path: /root/merge-scratch/run/A-TRACE/work/t6/target2/Annual Statement-211026 1110/Gamma Sample Manufacturing Compa-MH/Version-211026 1110/Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110.xlsx
[normal]     [DEBUG] Output directory exists: True
[normal]     [DEBUG] Preparing template data...
[normal]     [DEBUG] Template data keys: ['sales_folder', 'sales_filename', 'annual_folder', 'annual_filename', 'sales_reco_folder', 'sales_reco_filename']
[normal]     [DEBUG] Calling create_report_from_template...
[success]     ✅ Success: Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110.xlsx
[normal]     [DEBUG] Output file created, size: 5460 bytes
[success]   ✅ Completed Gamma Sample Manufacturing Company of India Two
[info] 
📊 Creating Level 1 summary report...
[normal]   Output: GST_Processing_Summary_211026_1110.xlsx
[success]   ✅ Success: GST_Processing_Summary_211026_1110.xlsx
[normal]   Sheets: Summary, Client Status, File Mapping, Errors, Variations
[success] 
🎉 ════════ PROCESSING COMPLETE ════════
[info] ✅ Processed: 2/2 clients
[info] 📁 Files organized: 4
[info] 📊 Reports created: 4
[info] 📂 Output location: /root/merge-scratch/run/A-TRACE/work/t6/target2/Annual Statement-211026 1110
[info] 
⏱️ Total processing time: 0 minutes 0 seconds
[success] 📂 Opened output folder
[info] ℹ️ Workflow is 'organise only'. Open Step 4 if you also want to refresh Power Query.
```
progress updates: 12 (last: (100, 'Processing complete | Total time: 0m 0s'))
pop-ups: [('info', ('Processing Complete', 'Successfully processed 2/2 clients\n\nFiles organized in:\n/root/merge-scratch/run/A-TRACE/work/t6/target2'))]; os.system calls: ['xdg-open "/root/merge-scratch/run/A-TRACE/work/t6/target2/Annual Statement-211026 1110"']
logger WARNING/ERROR: [('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!'), ('WARNING', 'core.excel_handler', 'Not on Windows - Power Query may not be preserved!'), ('WARNING', 'core.excel_handler', 'Using openpyxl - Power Query content may be lost!')]
file_organizer.get_summary(): {'timestamp': '211026 1110', 'processing_mode': 'fresh', 'target_folder': '/root/merge-scratch/run/A-TRACE/work/t6/target2', 'folders_created': 6, 'files_processed': 4, 'successful_copies': 4, 'failed_copies': 0, 'skipped_files': 0, 'errors': 0}

```
Annual Statement-211026 1110/
  GST_Processing_Summary_211026_1110.xlsx  (8255 B)
  Gamma Sample Manufacturing Compa-MH/
    Version-211026 1110/
      AnnualReport-Gamma_Sample_Manufacturing_Company_of_India_One-Maharashtra-2024-25.xlsx  (4879 B)
      AnnualReport-Gamma_Sample_Manufacturing_Company_of_India_Two-Maharashtra-2024-25.xlsx  (4879 B)
      GSTR3B/
      ITC/
      ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110.xlsx  (5445 B)
      ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110.xlsx  (5445 B)
      Sales/
        Sales-Gamma_Sample_Manufacturing_Company_of_India_One-Maharashtra-Apr-Mar.xlsx  (4879 B)
        Sales-Gamma_Sample_Manufacturing_Company_of_India_Two-Maharashtra-Apr-Mar.xlsx  (4879 B)
      Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110.xlsx  (5460 B)
      Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110.xlsx  (5460 B)
      _Organization_Report_Gamma Sample Manufacturing Company of India One_Maharashtra.txt  (958 B)
      _Organization_Report_Gamma Sample Manufacturing Company of India Two_Maharashtra.txt  (959 B)
```
- `Annual Statement-211026 1110/Gamma Sample Manufacturing Compa-MH/Version-211026 1110/Sales/Sales-Gamma_Sample_Manufacturing_Company_of_India_One-Maharashtra-Apr-Mar.xlsx` A1 = 'Gamma Sample Manufacturing Company of India One'
- `Annual Statement-211026 1110/Gamma Sample Manufacturing Compa-MH/Version-211026 1110/Sales/Sales-Gamma_Sample_Manufacturing_Company_of_India_Two-Maharashtra-Apr-Mar.xlsx` A1 = 'Gamma Sample Manufacturing Company of India Two'

#### 8. Step 4 (Extract) scan of the target: ExtractorPanel.scan_folder (power_query_extractor/gui/extractor_window.py:611-668)

level1_folder chosen: Annual Statement-211026 1105; logs: [('info', 'Scanning folder for clients...'), ('info', 'Using folder: Annual Statement-211026 1105'), ('success', 'Found 1 clients')]; status: Found 1 clients
- name='Gamma Sample Enterprises LLP-KA' latest_version='Version-211026 1105' has_itc=True has_sales=True

<!-- end of outputs/t6_scan_organise.md -->

### T9 - Step 4 on the merged folder from section 7 (RAN)

Script `scripts/t9_extract_on_collision.py` runs `ExtractorPanel.scan_folder` (stub panel) on `work/t6/target2`, then `ReportProcessor.process_client` with refresh/extract stubbed. Output `outputs/t9_extract_on_collision.md`, reproduced exactly (the "files glob would offer" line was printed after the run, so it already includes the refreshed copy the run created):



<!-- RAN output: outputs/t9_extract_on_collision.md (headings demoted by two levels) -->

clients found: [('Gamma Sample Manufacturing Compa-MH', True, True)]
process_client result client field: Gamma Sample Manufacturing Compa-MH
files glob would offer: ['ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110.xlsx', 'ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110.xlsx', 'ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110_Refreshed_211026_0428.xlsx']
calls: [('refresh', 'ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110_Refreshed_211026_0428.xlsx'), ('extract', 'ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110_Refreshed_211026_0428.xlsx'), ('refresh', 'Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110_Refreshed_211026_0428.xlsx'), ('extract', 'Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110_Refreshed_211026_0428.xlsx')]
files now in version folder: ['ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110.xlsx', 'ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110.xlsx', 'ITC_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110_Refreshed_211026_0428.xlsx', 'Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110.xlsx', 'Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_One_MH_211026_1110_Refreshed_211026_0428.xlsx', 'Sales_Report_Gamma_Sample_Manufacturing_Company_of_India_Two_MH_211026_1110.xlsx']

<!-- end of outputs/t9_extract_on_collision.md -->

### T11 - counting a parsed file with a blank client (RAN)

Output `outputs/t11_blank_client_count.md`, reproduced exactly:



<!-- RAN output: outputs/t11_blank_client_count.md (headings demoted by two levels) -->

clients: ['Alpha Test Traders-MH']
variations: [('GSTR3B- -Maharashtra-Apr.xlsx', 'Could not parse filename', 'GSTR3B-ClientName-State-Month.xlsx'), ('Random.xlsx', 'No matching pattern found', 'Check file naming convention')]
stats: total 3 parsed 2 unparsed 1 variations 2

<!-- end of outputs/t11_blank_client_count.md -->

### T6 observations - what is scanned, reported, created, and whether originals are touched

**Scanned (RAN):** 21 of the 27 files (plus one folder) reached the parser. Never reached it, and never reported anywhere (no count, no variation, no log line at all - `scan_excel_files` drops them silently, `utils/helpers.py:329-359`): the file inside the sub-folder (scan is top-level only, `utils/helpers.py:328`), the `.csv`, the `.xlsb`, `fake.xlsx` (no signature), `tiny.xlsx` and the `~$` lock file (under 1024 bytes, `utils/helpers.py:303`), and the folder named `...xlsx` (not a regular file, `utils/helpers.py:296`). HIGH.

**Reported (RAN):**

- 5 "clients" from 3 real ones: `Alpha Test Traders-MH` (Complete), `Beta Demo Industries Private Limited-GJ` (Missing 4 files), `Gamma Sample Enterprises LLP-KA` (Missing 2 files, plus `Multiple IMS Reco (2 files)`, `Multiple Sales (2 files)`), and two extra rows created by naming variants: `GAMMA SAMPLE ENTERPRISES LLP-KA` (the capitals Annual Report; Missing 5) and `Gamma-SAM` (the hyphenated Sales Reco: client `Gamma`, state `Sample`; Missing 5). So client 3 looks as if it lacks its Annual Report and Sales Reco although both were supplied. The state written as `KA` joined the right client (both give code `KA`).
- The status text shows only "Missing N files" when files are missing even if there are also duplicates (`core/file_parser.py:281-283`); duplicates appear only in the "Extra" column, which shows the NUMBER OF TYPES with duplicates (`gui/handlers/client_handler.py:32`), not the number of extra files.
- Two variations (`Gamma notes.xlsx` -> "Check file naming convention"; the `.xlsm` -> GSTR-3B suggestion). Summary text: 21 files, 19 parsed, 2 unparsed, 5 clients, 1 complete, "Clients with Missing Files: 4" (this count is simply "not Complete", so duplicate-only clients are counted as "missing", `core/file_parser.py:328`), completeness rate 20.0%, distribution by type, 98.7 KB.
- `validate_results()` would have flagged missing critical files (`Annual Report`, `GSTR-3B Export`) but the app never calls it (RAN grep: no caller). [GST RULE] (which files are "critical", `core/file_parser.py:359`).

**Created (RAN, run 1, fresh mode):** one `Annual Statement-211026 1030/` folder; per client `<Level-2 key>/Version-211026 1030/` with `GSTR3B/`, `ITC/`, `Sales/` (always created, even when empty), the Annual Report copied into the Version folder itself, GSTR-2B and IMS into `ITC/`, Sales and Sales Reco into `Sales/`, GSTR-3B into `GSTR3B/` (`core/file_organizer.py:302-315`); every copied file renamed by `sanitize_filename` (spaces -> `_`, Private Limited -> Pvt_Ltd, `(1) ` -> `(1)_`); an `ITC_Report_<client>_<code>_<stamp>.xlsx` and `Sales_Report_...` from the templates; `_Organization_Report_<client>_<state>.txt`; and `GST_Processing_Summary_211026_1030.xlsx` in the Level-1 folder with sheets Summary, Client Status, File Mapping, Errors, Variations. Every client - including the five incomplete ones - got both reports and status "Success". The GUI then tried to open the folder (`xdg-open ...` on Linux, `os.startfile` on Windows, `processing_handler.py:505-516`).

**Links sheet contents (RAN):** cells B2/B4/B5/B7/B8/B10/B11 (ITC) and B2/B3/B5/B6/B8/B9 (Sales) per `utils/constants.py:102-126`, holding folder paths with backslashes and file names without extension (`core/excel_handler.py:603-713`). Cells for missing types are simply not written (`core/excel_handler.py:299-301`), so whatever the template already has in them stays. For client 3 the Sales link names `Sales-..._(1)` and the IMS link `ImsReco-..._(1)` - the "(1)" copies, because they sort first (see T3 b).

**Per-client text report (RAN):** "Successfully Copied" and "Folders Created" are running totals for the whole run, not for that client (e.g. Beta: "Total Files: 2, Successfully Copied: 18, Folders Created: 21") - `core/file_organizer.py:363-368` counts `self.copied_files` / `self.created_folders`, which accumulate across clients. HIGH.

**Originals (RAN):** after run 1 and again after all nine runs, the source tree was identical: same names, same sha256, same sizes, same modification times to the nanosecond, same folder modification times. Copying uses `shutil.copy2` (`utils/helpers.py:199`), which also carries the original's modification time onto the copy. Nothing is moved or deleted in the source. HIGH. (Reading files does update their last-access time on some systems; not checked, not relevant.)

**Running again (RAN, section 5):**

- Run 2, fresh, **same minute**: same Level-1 and Version folders; each existing copy is first backed up as `<name>_backup_<DDMMYY_HHMM><ext>` then overwritten (`core/file_organizer.py:277-282`); the reports and the Level-1 summary are overwritten without backup - the summary that listed 5 clients now lists only the 1 client of run 2 (size 9262 -> 8489 bytes).
- Run 3, rerun, same minute: files silently overwritten (no backup, no skip: `rerun` has no branch at 277-287).
- Run 4, **resume**, same minute: all 9 files "Skipped - File already exists", "Organized 0 files", but both reports are rebuilt and the summary overwritten.
- Run 5, **resume, a later minute**: a NEW `Version-211026 1045` folder and all 9 files copied again. "Resume" never continues an earlier Version folder: the Version folder is always named after the current minute (`core/file_organizer.py:103-108`), so skipping only happens within the same minute. HIGH. (The mode's description is "Continue previously interrupted processing", `utils/constants.py:185-188`.)
- Runs 5 and 6 (resume/rerun) reused the newest `Annual Statement-...` folder (by modification time) and wrote a new summary there.
- `create_backup` twice in one minute writes the same backup name, so the earlier backup is lost (T8 section 4, RAN).

**"Include client name" (RAN, section 6):** with the global option on, category folders are `GSTR3B (<sanitized client>)` etc.; with it off, a per-client setting keyed `client-CODE` works (run 8). The client-name length limit applies only to the Level-2 folder, not to these labels (`core/file_organizer.py:90,119`).

**Two clients merged into one folder (RAN, section 7 + T9):** `Gamma Sample Manufacturing Company of India One` and `... Two` share the Level-2 folder `Gamma Sample Manufacturing Compa-MH` and Version folder; both clients' files, reports and text reports sit together (no overwrite here because the full names differ inside the file names). Step 4 then sees ONE client with that folder name, and in my run refreshed the ITC report of "Two" and the Sales report of "One" into the same consolidated row; the other two reports are never refreshed or extracted. Which file is picked depends on listing order (T5 obs 1). HIGH for the merge, MEDIUM for which file is picked on Windows. If the templates' Power Query reads every file in the linked folder (I cannot see the templates), each report would also read the other client's files - LOW, question for the author.

**Step 4 scan (RAN, section 8):** `ExtractorPanel.scan_folder` looks only in the newest `Annual Statement-*` folder by modification time (`power_query_extractor/gui/extractor_window.py:623-632`), skips folders starting with `_` (639), uses each client folder's newest `Version-*` (641-648). After runs 7 and 8 created new Level-1 folders, Step 4 showed only the one client of run 8.

**Case-only differences on Windows (MEDIUM, reasoning, not runnable here):** on Linux `GAMMA SAMPLE ENTERPRISES LLP-KA` and `Gamma Sample Enterprises LLP-KA` became two folders. Windows file names are case-insensitive by default, so on Windows both clients would use the SAME Level-2 and Version folder, and their report names `ITC_Report_GAMMA_SAMPLE_ENTERPRISES_LLP_KA_<stamp>.xlsx` / `ITC_Report_Gamma_Sample_Enterprises_LLP_KA_<stamp>.xlsx` would be the same file - the second would overwrite the first (in the openpyxl path `shutil.copy2` overwrites, `core/excel_handler.py:308`; the COM path was not traced by me).

## 9. T7 - Setup checklist, scan/validate handlers, and the Validation table keys (RAN)

**Code (READ):** `gui/utils/setup_validation.py` is Tk-free: `evaluate_setup` returns four checks (source folder, ITC template, Sales template, target folder; 75-85). A folder check only tests "set" and "exists" (47-53) - not "is a folder". A template check tests "set", "exists" and the suffix being one of `.xlsx .xltx .xltm .xlsm` (21, 56-64) - not "is a file". `can_scan` needs only the source (88-91); `can_process` needs all four (94-97). `FileHandler.validate_scan_inputs` repeats "set" and "exists" for the source (`gui/handlers/file_handler.py:120-130`); `validate_processing_inputs` checks both templates set and existing and the target set (132-151) - not the target's existence (FileOrganizer creates it, `core/file_organizer.py:66-69`) and not the template suffix. `ClientHandler.update_client_tree` puts each client's text values into the Treeview (`gui/handlers/client_handler.py:34-44`); `get_selected_clients` reads them BACK from the Treeview and rebuilds the key as `f"{values[0]}-{get_state_code(values[1])}"` (161-173); `toggle_folder_name_setting` and `update_global_folder_setting` do the same (175-223).

**How I ran it:** `setup_validation` directly; `FileHandler` with a stub app and recorded message boxes; `ClientHandler` with a REAL `ttk.Treeview` under `xvfb-run` (no window shown), message boxes recorded, synthetic client entries including made-up numeric-looking names. Script `scripts/t7_setup_and_handlers.py`, output `outputs/t7_setup_and_handlers.md`, reproduced exactly:



<!-- RAN output: outputs/t7_setup_and_handlers.md (headings demoted by two levels) -->

#### setup_validation.evaluate_setup (gui/utils/setup_validation.py:75-85) with the default exists()

| case | checks (key: state - message) | can_scan | can_process | blocking(scan) | blocking(processing) | summary_line |
|---|---|---|---|---|---|---|
| all blank | source: missing - Not selected yet; itc: missing - Not selected yet; sales: missing - Not selected yet; target: missing - Not selected yet | False | False | ['Source folder'] | ['Source folder', 'ITC template', 'Sales template', 'Target folder'] | 0 of 4 items set |
| whitespace only | source: missing - Not selected yet; itc: missing - Not selected yet; sales: missing - Not selected yet; target: missing - Not selected yet | False | False | ['Source folder'] | ['Source folder', 'ITC template', 'Sales template', 'Target folder'] | 0 of 4 items set |
| all good (.xltx + .XLSX) | source: ok - /root/merge-scratch/run/A-TRACE/work/t7/src; itc: ok - ITC.xltx; sales: ok - Sales.XLSX; target: ok - /root/merge-scratch/run/A-TRACE/work/t7/tgt | True | True | [] | [] | Ready - all 4 items set |
| source does not exist | source: invalid - Folder not found; itc: ok - ITC.xltx; sales: ok - Sales.XLSX; target: ok - /root/merge-scratch/run/A-TRACE/work/t7/tgt | False | False | ['Source folder'] | ['Source folder'] | 3 of 4 items set |
| template is .xls | source: ok - /root/merge-scratch/run/A-TRACE/work/t7/src; itc: invalid - Not an Excel template; sales: ok - Sales.XLSX; target: ok - /root/merge-scratch/run/A-TRACE/work/t7/tgt | True | False | [] | ['ITC template'] | 3 of 4 items set |
| template is .xltm | source: ok - /root/merge-scratch/run/A-TRACE/work/t7/src; itc: ok - macro.xltm; sales: ok - Sales.XLSX; target: ok - /root/merge-scratch/run/A-TRACE/work/t7/tgt | True | True | [] | [] | Ready - all 4 items set |
| template is .txt | source: ok - /root/merge-scratch/run/A-TRACE/work/t7/src; itc: invalid - Not an Excel template; sales: ok - Sales.XLSX; target: ok - /root/merge-scratch/run/A-TRACE/work/t7/tgt | True | False | [] | ['ITC template'] | 3 of 4 items set |
| template is a FOLDER named folder.xlsx | source: ok - /root/merge-scratch/run/A-TRACE/work/t7/src; itc: ok - folder.xlsx; sales: ok - Sales.XLSX; target: ok - /root/merge-scratch/run/A-TRACE/work/t7/tgt | True | True | [] | [] | Ready - all 4 items set |
| source is a FILE, target is a FILE | source: ok - ...run/A-TRACE/work/t7/a_file_not_folder.xlsx; itc: ok - ITC.xltx; sales: ok - Sales.XLSX; target: ok - ...erge-scratch/run/A-TRACE/work/t7/notes.txt | True | True | [] | [] | Ready - all 4 items set |
| target does not exist yet | source: ok - /root/merge-scratch/run/A-TRACE/work/t7/src; itc: ok - ITC.xltx; sales: ok - Sales.XLSX; target: invalid - Folder not found | True | False | [] | ['Target folder'] | 3 of 4 items set |
| relative path '.' | source: ok - .; itc: ok - ITC.xltx; sales: ok - Sales.XLSX; target: ok - . | True | True | [] | [] | Ready - all 4 items set |
| NUL character in path | source: invalid - Folder not found; itc: ok - ITC.xltx; sales: ok - Sales.XLSX; target: ok - /root/merge-scratch/run/A-TRACE/work/t7/tgt | False | False | ['Source folder'] | ['Source folder'] | 3 of 4 items set |

_short() on long paths: '...2024-25\\GST\\Annual Statement Source Folder' 'short'
injected exists() was called with: ['a', 'b.xlsx', 'c.xlsx', 'd']

#### FileHandler.scan_files / validate_* (gui/handlers/file_handler.py:57-151) with a stub app

- blank source: message boxes=[('error', ('Error', 'Please select a source folder containing GST files'))]; status=[]; log=[]; tab selected=[]
- missing source: message boxes=[('error', ('Error', 'Source folder does not exist'))]; status=[]; log=[]; tab selected=[]
- source is a file: message boxes=[('error', ('Scan Error', 'Not a directory: /root/merge-scratch/run/A-TRACE/work/t7/a_file_not_folder.xlsx'))]; status=['Scanning files...']; log=[('info', '🔍 Starting file scan...'), ('error', '❌ Scan error: Not a directory: /root/merge-scratch/run/A-TRACE/work/t7/a_file_not_folder.xlsx')]; tab selected=[]
- empty source folder: message boxes=[]; status=['Scanning files...', '🔍 Found 0 Excel files', 'No valid GST files found']; log=[('info', '🔍 Starting file scan...'), ('call', 'update_summary'), ('call', 'update_client_tree'), ('warning', '⚠️ No valid GST files found')]; tab selected=[]
- validate_processing_inputs, no templates: returned False; boxes=[('error', ('Error', 'Please select both ITC and Sales template files'))]
- validate_processing_inputs, templates, no target: returned False; boxes=[('error', ('Error', 'Please select a target folder for organized files'))]
- validate_processing_inputs, missing ITC template file: returned False; boxes=[('error', ('Error', 'ITC template file not found'))]
- validate_processing_inputs, target missing on disk (not checked): returned True; boxes=[]
- validate_processing_inputs, template is .txt (not checked here): returned True; boxes=[]

#### ClientHandler tree and keys with a real ttk.Treeview (gui/handlers/client_handler.py:18-223)

| parser key (client_data) | tree values as Tk returns them (repr) | get_selected_clients() result for this row alone | matches a client_data key? |
|---|---|---|---|
| `Alpha Test Traders-MH` | `['Alpha Test Traders', 'Maharashtra', ' Complete', 1, 'None', 0, '☐ No']` | `['Alpha Test Traders-MH']` | yes |
| `1001-MH` | `[1001, 'Maharashtra', ' Complete', 1, 'None', 0, '☐ No']` | `['1001-MH']` | yes |
| `0042-MH` | `[42, 'Maharashtra', ' Complete', 1, 'None', 0, '☐ No']` | `['42-MH']` | **NO - KeyError later** |
| `12.50-DL` | `['12.50', 'Delhi', ' Complete', 1, 'None', 0, '☐ No']` | `['12.50-DL']` | yes |
| `Beta  Demo-GJ` | `['Beta  Demo', 'Gujarat', ' Missing 1 files', 1, 'Sales', 0, '☐ No']` | `['Beta  Demo-GJ']` | yes |
| `Gamma {Test}-KA` | `['Gamma {Test}', 'Karnataka', ' Has duplicates', 1, 'None', 1, '☐ No']` | `['Gamma {Test}-KA']` | yes |
| `Delta Mock-MH` | `['Delta Mock', 'MH', ' Complete', 1, 'None', 0, '☐ No']` | `['Delta Mock-MH']` | yes |
| `Epsilon Fake-27` | `['Epsilon Fake', 27, ' Complete', 1, 'None', 0, '☐ No']` | EXCEPTION AttributeError: 'int' object has no attribute 'strip' | **crash** |

select_complete_clients():
  selected: ['Alpha Test Traders', 1001, 42, '12.50', 'Delta Mock', 'Epsilon Fake'] | status: Selected 6 complete clients

toggle_folder_name_setting on each row (asks only when the name is >10 chars):
  toggle on [1001, 'Maharashtra'] -> EXCEPTION TypeError object of type 'int' has no len()
  toggle on [42, 'Maharashtra'] -> EXCEPTION TypeError object of type 'int' has no len()
  toggle on ['Epsilon Fake', 27] -> EXCEPTION AttributeError 'int' object has no attribute 'strip'
  client_folder_settings now: {'Alpha Test Traders-MH': True, '1001-MH': True, '0042-MH': False, '12.50-DL': True, 'Beta  Demo-GJ': True, 'Gamma {Test}-KA': True, 'Delta Mock-MH': True, 'Epsilon Fake-27': False, '42-MH': True}
  keys that do NOT exist in client_data: ['42-MH']
  dialogs: ['Long Client Name Warning', 'Long Client Name Warning']
  untick one row while global ON -> global now False dialogs: ['info']

<!-- end of outputs/t7_setup_and_handlers.md -->

### T12 - more Treeview round-trips (RAN)

Script `scripts/t12_treeview_odd_text.py` (under `xvfb-run`): what comes back from a Treeview for text that Tcl treats specially. Output `outputs/t12_treeview_odd_text.md`, reproduced exactly. Only whole-number-looking text changes (`'+12'` -> 12, `'07'` -> 7, earlier `'0042'` -> 42, `'27'` -> 27); braces, brackets, `$`, quotes, spaces, `1e3`, `0x1F`, `1.5` come back unchanged.



<!-- RAN output: outputs/t12_treeview_odd_text.md (headings demoted by two levels) -->

inserted ('Gamma {Test', 'Karnataka') -> read back ['Gamma {Test', 'Karnataka'] -> key Gamma {Test-Karnataka | same text? True
inserted ('Gamma Test}', 'Karnataka') -> read back ['Gamma Test}', 'Karnataka'] -> key Gamma Test}-Karnataka | same text? True
inserted ('Delta [Mock] Co', 'Delhi') -> read back ['Delta [Mock] Co', 'Delhi'] -> key Delta [Mock] Co-Delhi | same text? True
inserted ('Eta $x Co', 'Goa') -> read back ['Eta $x Co', 'Goa'] -> key Eta $x Co-Goa | same text? True
inserted ('Theta "Q" Co', 'Goa') -> read back ['Theta "Q" Co', 'Goa'] -> key Theta "Q" Co-Goa | same text? True
inserted ('1e3', 'Goa') -> read back ['1e3', 'Goa'] -> key 1e3-Goa | same text? True
inserted ('0x1F', 'Goa') -> read back ['0x1F', 'Goa'] -> key 0x1F-Goa | same text? True
inserted ('+12', 'Goa') -> read back [12, 'Goa'] -> key 12-Goa | same text? False
inserted ('007 Traders', 'Goa') -> read back ['007 Traders', 'Goa'] -> key 007 Traders-Goa | same text? True
inserted ('  Iota  ', 'Goa') -> read back ['  Iota  ', 'Goa'] -> key   Iota  -Goa | same text? True
inserted ('', 'Goa') -> read back ['', 'Goa'] -> key -Goa | same text? True
inserted ('Kappa', '07') -> read back ['Kappa', 7] -> key Kappa-7 | same text? False
inserted ('Lambda', '1.5') -> read back ['Lambda', '1.5'] -> key Lambda-1.5 | same text? True
inserted ('Mu', 'true') -> read back ['Mu', 'true'] -> key Mu-true | same text? True

<!-- end of outputs/t12_treeview_odd_text.md -->

### T7 observations

1. **Tk turns number-looking cell values into numbers, which breaks the client keys.** A Treeview gives values back as Python numbers when they look like whole numbers: client `1001` -> `1001` (key still `1001-MH`, works by luck), client `0042` -> `42` (key `42-MH`, which is not in the client data), state `27` -> the number 27. RAN, HIGH. Consequences (RAN in T7 and T8 section 2):
   - a state written as a number makes `get_selected_clients` crash (`AttributeError: 'int' object has no attribute 'strip'` in `get_state_code`, `utils/helpers.py:601`, called from `gui/handlers/client_handler.py:169`) - so Dry Run and Start Processing fail whenever such a client is ticked; `toggle_folder_name_setting` and `update_global_folder_setting` crash the same way (181, 218; READ for 218);
   - a client like `0042` gives key `42-MH`; `process_files_thread` looks it up at `gui/handlers/processing_handler.py:232`, OUTSIDE the per-client error handler that starts at 243, so the `KeyError` jumps to the outer handler (522-525): "Fatal error: '42-MH'", an error box, and every client after it in the list is not processed and no summary workbook is written (T8 section 2: Alpha processed, Gamma never reached). `dry_run` would fail at `processing_handler.py:56` the same way (READ);
   - `toggle_folder_name_setting` on a numeric client crashes at `len(values[0])` (`client_handler.py:191`) after already flipping the setting, and the stored key `42-MH` never matches what `FileOrganizer` looks up (`core/file_organizer.py:122,130`).
   Decimal-looking text (`12.50`) came back as text in my run (RAN).
2. The Status column text starts with the icon and a space (`client_handler.py:39`); `select_complete_clients` tests `'Complete' in` that text (152) - works for the statuses the parser produces (RAN: 6 rows with status Complete selected).
3. The long-name question is asked only when switching a client's folder-name setting ON and the name is over 10 characters (hard-coded 10 at `client_handler.py:191`), while Start Processing warns using the configured limit (`processing_handler.py:90-113`). RAN for the first; READ for the second.
4. Setup checklist (RAN): a template that is a FOLDER named `folder.xlsx` is "ok"; a source or target that is a FILE is "ok" (the checklist then says "Ready - all 4 items set"), and pressing Scan with a file as source gives the error box "Not a directory: ..." from `scan_folder` (`core/file_parser.py:154-155`). `.xls` templates are refused, `.xltm` accepted. A blank or whitespace value is "missing"; `.` (current folder) counts as an existing folder. A path containing a NUL character is "Folder not found" rather than a crash (`setup_validation.py:41-44`).
5. `validate_processing_inputs` lets through a target that does not exist (it is created later) and a `.txt` "template" (RAN: both returned True). The `.txt` would fail later inside report creation.
6. The project's own test `tests/test_setup_validation.py::TestEvaluateSetup::test_ok_template_shows_filename_only` fails on Linux (RAN: expects `itc.xltx`, gets `C:\gst\itc.xltx`) because the test uses a Windows path, which Linux does not split at backslashes. It is a platform-dependent test, not a code failure; on Windows it would pass (MEDIUM).

## 10. T8 - Edge and failure branches (RAN)

Script `scripts/t8_edges.py` (stub app as in T6 where needed; folder stamp fixed to `211026 1200` in `core/file_organizer.py`, report/backup names use the real clock, so the report names below carry `0427`). Output `outputs/t8_edges.md`, reproduced exactly:



<!-- RAN output: outputs/t8_edges.md (headings demoted by two levels) -->

#### 1. Same client written two ways: 'Pvt Ltd' in some names, 'Private Limited' in others

- parser client key `Beta Demo Industries Private Limited-GJ`: status 'Missing 4 files', files ['GSTR3B-Beta Demo Industries Private Limited-Gujarat-Apr.xlsx', 'Sales-Beta Demo Industries Private Limited-Gujarat-Apr-Mar.xlsx'], extra []
- parser client key `Alpha Test Traders-MH`: status 'Missing 5 files', files ['Sales-Alpha Test Traders-Maharashtra-Apr-Mar .xlsx', 'Sales-Alpha Test Traders-Maharashtra-Apr-Mar.xlsx'], extra ['Multiple Sales (2 files)']
- parser client key `Beta Demo Industries Pvt Ltd-GJ`: status 'Missing 5 files', files ['Sales-Beta Demo Industries Pvt Ltd-Gujarat-Apr-Mar.xlsx'], extra []
- organise `Beta Demo Industries Private Limited-GJ` -> level2 `Beta Demo Industries Pvt Ltd-GJ`; results: [('GSTR3B-Beta Demo Industries Private Limited-Gujarat-Apr.xlsx', 'GSTR3B-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr.xlsx', 'Success', None), ('Sales-Beta Demo Industries Private Limited-Gujarat-Apr-Mar.xlsx', 'Sales-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr-Mar.xlsx', 'Success', None)]
- organise `Alpha Test Traders-MH` -> level2 `Alpha Test Traders-MH`; results: [('Sales-Alpha Test Traders-Maharashtra-Apr-Mar .xlsx', 'Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx', 'Success', None), ('Sales-Alpha Test Traders-Maharashtra-Apr-Mar.xlsx', 'Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx', 'Success', 'Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar_backup_211026_0427.xlsx')]
- organise `Beta Demo Industries Pvt Ltd-GJ` -> level2 `Beta Demo Industries Pvt Ltd-GJ`; results: [('Sales-Beta Demo Industries Pvt Ltd-Gujarat-Apr-Mar.xlsx', 'Sales-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr-Mar.xlsx', 'Success', 'Sales-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr-Mar_backup_211026_0427.xlsx')]

resulting files and their A1 note:
- `Annual Statement-211026 1200/Alpha Test Traders-MH/Version-211026 1200/Sales/Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx` A1='no trailing space'
- `Annual Statement-211026 1200/Alpha Test Traders-MH/Version-211026 1200/Sales/Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar_backup_211026_0427.xlsx` A1='trailing space before .xlsx'
- `Annual Statement-211026 1200/Beta Demo Industries Pvt Ltd-GJ/Version-211026 1200/GSTR3B/GSTR3B-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr.xlsx` A1='fake'
- `Annual Statement-211026 1200/Beta Demo Industries Pvt Ltd-GJ/Version-211026 1200/Sales/Sales-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr-Mar.xlsx` A1='version written with Pvt Ltd'
- `Annual Statement-211026 1200/Beta Demo Industries Pvt Ltd-GJ/Version-211026 1200/Sales/Sales-Beta_Demo_Industries_Pvt_Ltd-Gujarat-Apr-Mar_backup_211026_0427.xlsx` A1='version written with Private Limited'

#### 2. A selected key that is not in client_data (what the '0042' -> 42 tree conversion produces)

client_data keys: ['0042-MH', 'Alpha Test Traders-MH', 'Gamma Sample Enterprises LLP-KA'] | selected keys passed in: ['Alpha Test Traders-MH', '42-MH', 'Gamma Sample Enterprises LLP-KA']
GUI log (errors/summary lines): ['  ✅ Completed Alpha Test Traders', "💥 Fatal error: '42-MH'"]
message boxes: [('error', ('Processing Error', "'42-MH'"))]
target tree: ['Annual Statement-211026 1200/Alpha Test Traders-MH/Version-211026 1200/ITC_Report_Alpha_Test_Traders_MH_211026_0427.xlsx', 'Annual Statement-211026 1200/Alpha Test Traders-MH/Version-211026 1200/Sales/Sales-Alpha_Test_Traders-Maharashtra-Apr-Mar.xlsx', 'Annual Statement-211026 1200/Alpha Test Traders-MH/Version-211026 1200/Sales_Report_Alpha_Test_Traders_MH_211026_0427.xlsx', 'Annual Statement-211026 1200/Alpha Test Traders-MH/Version-211026 1200/_Organization_Report_Alpha Test Traders_Maharashtra.txt']
is_processing reset: False

#### 3. FileOrganizer construction and single-file branches

- mode 'fresh': OK
- mode 'rerun': OK
- mode 'resume': OK
- mode 'Fresh': ValueError: Invalid processing mode: Fresh
- mode 'dry': ValueError: Invalid processing mode: dry
- target folder under a FILE: ValueError: Cannot create target folder: [Errno 20] Not a directory: '/root/merge-scratch/run/A-TRACE/work/t8/a_file/sub'
- client_name_max_length=0 -> stored 35
- client_name_max_length=-3 -> stored 35
- client_name_max_length=5 -> stored 5
- client_name_max_length='35' -> TypeError: '>' not supported between instances of 'str' and 'int'
- _organize_single_file unknown type -> {'filename': 'x.xlsx', 'file_type': 'Unknown Type', 'source': '/root/merge-scratch/run/A-TRACE/work/t8/x.xlsx', 'destination': '', 'status': 'Failed', 'error': 'No destination folder for type: Unknown Type'}
- _organize_single_file missing source -> status 'Failed' error 'Copy operation failed'; dest dir created? False

#### 4. create_backup twice within one minute

- backups: Sales-X-MH-Apr-Mar_backup_211026_0427.xlsx, Sales-X-MH-Apr-Mar_backup_211026_0427.xlsx; same path: True; backup content now 'v2' (v1 lost)
- create_backup(missing) -> None

#### 5. cleanup_on_error (never called by the app) on a version folder that holds only '_' files

- version folder still exists? False
- with a real file: still exists? True

#### 6. scan_folder failure paths and the parallel-read threshold

- scan_folder(does_not_exist) -> FileNotFoundError: Folder not found: /root/merge-scratch/run/A-TRACE/work/t8/does_not_exist
- scan_folder(a_file) -> ValueError: Not a directory: /root/merge-scratch/run/A-TRACE/work/t8/a_file
- 3 files (sequential path): ['GSTR3B-A Test-MH-Apr.xlsx', 'GSTR3B-A Test-MH-Jun.xlsx', 'GSTR3B-A Test-MH-May.xlsx']
- 5 files (thread pool path): ['GSTR3B-A Test-MH-Apr.xlsx', 'GSTR3B-A Test-MH-Aug.xlsx', 'GSTR3B-A Test-MH-Jul.xlsx', 'GSTR3B-A Test-MH-Jun.xlsx', 'GSTR3B-A Test-MH-May.xlsx']
- progress callback that raises: logged ['Progress callback failed: progress widget gone'] (x5), scan continues
- validate_excel_file(1023 bytes, PK header) -> False
- validate_excel_file(1024 bytes, PK header) -> True
- validate_excel_file(a plain .zip renamed .xlsx) -> True
- scan_excel_files(missing folder) -> [] (error only logged)

#### 7. Small helpers

- truncate_path('C:\\Users\\someone\\Clients\\Annual Statement-211026 1030\\Alpha Test Traders-MH\\Version-211026 1030',) -> '...ent-211026 1030\\Alpha Test Traders-MH\\Version-211026 1030'
- truncate_path('/a/b',) -> '/a/b'
- truncate_path('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx',) -> '...xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
- truncate_path('/root/merge-scratch/run/A-TRACE/work/t8/some/very/long/path/to/a/file/that/is/long.xlsx', 30) -> '//.../long.xlsx'
- clean_windows_path('') -> ''; clean_windows_path(None) -> ''; clean_windows_path('a/b c/d.xlsx') -> 'a\\b c\\d.xlsx'
- get_relative_excel_path('/x/y/report.xlsx', '/x/y/Sales/a.xlsx') -> 'Sales\\a.xlsx'
- get_relative_excel_path('/x/y/report.xlsx', '/z/a.xlsx') -> '\\z\\a.xlsx'
- get_unique_path(u.xlsx with u_1 existing) -> u_2.xlsx; for a folder -> bk_1
- get_short_path on Linux -> '/a/b'
- calculate_file_hash(u.xlsx) -> 9dd4e461268c8034f5c8564e155c67a6; missing -> None; bad algorithm -> None
- format_size(1536)='1.5 KB', format_size(1024**5)='1024.0 TB', format_size(1024**6)='1048576.0 TB'
- format_duration(59.9)='59s', (3600)='1h 0m', (-5)='-5s'
- ProgressTracker(total=0).update() ok; get_eta() -> '0s'
- ProgressTracker(4) after 1 in ~10s: get_eta() -> '30s'
- safe_path_join('', 'a', None, 'b') -> PosixPath('a/b')
- safe_path_join('', None) -> ValueError: No valid path components provided
- get_file_info(missing) -> {'error': 'File not found'}

<!-- end of outputs/t8_edges.md -->

### T8 observations

1. **One client written two ways ("Pvt Ltd" in one file name, "Private Limited" in another)**: the scan shows two clients (`Beta Demo Industries Private Limited-GJ`, `Beta Demo Industries Pvt Ltd-GJ`), each incomplete. Organising both puts them into the same folder `Beta Demo Industries Pvt Ltd-GJ` and both Sales files get the same cleaned name, so the second one processed overwrites the first; in fresh mode the first survives only as `..._backup_<time>.xlsx`, and the ITC/Sales reports of the first client now link to a file that holds the second file's content. RAN, HIGH. In rerun mode there would be no backup (READ, `core/file_organizer.py:277-287`).
2. **Two files of one client that clean to the same name** (here `...-Apr-Mar.xlsx` and `...-Apr-Mar .xlsx`, the second with a space before the extension): the scan reports `Multiple Sales (2 files)`, both copies report "Success", but only one file remains under the real name and the other becomes a backup. RAN, HIGH.
3. A selected key missing from the client data aborts the whole run (see T7 obs 1). RAN.
4. `FileOrganizer` rejects unknown modes (case-sensitive: `Fresh` is refused) and a target that cannot be created (`ValueError: Cannot create target folder`). A text `'35'` as length limit raises `TypeError` at `core/file_organizer.py:42` (the GUI passes an `IntVar` value, which is a number, `gui/main_window.py:93`; a non-numeric typed value would make `IntVar.get()` itself fail inside the worker's error handler - READ, MEDIUM). RAN.
5. `_organize_single_file`: unknown type -> "No destination folder for type"; missing source -> "Copy operation failed" (no exception). RAN.
6. `create_backup` twice in one minute: the same backup name, the earlier backup's content is lost (`v1 lost`). RAN, HIGH.
7. `cleanup_on_error` (never called anywhere, RAN grep) would delete a whole Version folder if it contains only files whose names start with `_` (the text report) - RAN. Dead code.
8. `scan_folder` on a missing folder raises `FileNotFoundError`; on a file raises `ValueError` (caught by `FileHandler.scan_files` and shown in an error box, T7). The sequential path (fewer than 4 files) and the 16-thread path give the same sorted order (RAN). A progress callback that fails is logged and ignored (RAN). Files of 1023 bytes are rejected, 1024 accepted. **Any ZIP file renamed `.xlsx` is accepted as Excel** (only the first 4 bytes are checked, `utils/helpers.py:282-284`). A missing source folder inside `scan_excel_files` returns `[]` and only logs (RAN).
9. Helpers (RAN): `truncate_path` keeps first and last parts only (`'//.../long.xlsx'` for a Linux path - cosmetic); `clean_windows_path(None)` -> `''`; `get_relative_excel_path` falls back to the absolute path; `get_unique_path` appends `_1`, `_2` (unused by the app, RAN grep); `get_short_path` returns the path unchanged off Windows; `calculate_file_hash` (unused) returns `None` on any error; `format_size` stops at TB; `format_duration(-5)` = `'-5s'`; `ProgressTracker(total=0)` does not divide by zero; `safe_path_join` skips empty parts and raises when none are left.
10. Dead constants (RAN grep, no use outside `utils/constants.py`/`utils/__init__.py`/tests): `EXCEL_EXTENSIONS`, `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB` (so no maximum file size is enforced anywhere), `EXCEL_SIGNATURES` (the signatures are repeated inline at `utils/helpers.py:264-265`). `EXCEL_SAFETY` is stored as `ExcelHandler.safety_config` (`core/excel_handler.py:55`) but that attribute is never read (RAN grep).

## 11. The project's own tests for these modules (RAN)

`pytest -p no:cacheprovider tests/test_file_patterns.py tests/test_helpers.py tests/test_refresh_naming.py tests/test_scan.py tests/test_file_ops.py tests/test_setup_validation.py tests/test_processing_handler.py` in the run copy under `xvfb-run`: **134 passed, 1 failed** (`test_ok_template_shows_filename_only`, a Windows-path test, see T7 obs 6). No test in `tests/` refers to `FileOrganizer`, `organize_files`, `prepare_template_data`, `DataConsolidator`, `create_report`, `_extract_data`, `get_selected_clients` or `update_client_tree` (RAN grep), so the problems in T4, T6, T7 and T8 are not covered by the suite.

## 12. All surprising or possibly wrong outputs in one list (technical; GST rules only described)

| # | what happens (plain English) | where | evidence | confidence |
|---|---|---|---|---|
| TR-01 | Consolidated report: the "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" sheets show the GSTR-3B Difference figures (Diff GSTR-3B!J15/K15/L15) instead of their own configured cells, whenever the Diff GSTR-3B sheet exists. Output does not follow the code's own configuration. Present since 47a1492 (shared with B). [GST RULE] area (which figure is reported where). | `power_query_extractor/core/data_consolidator.py:139-150`; config `power_query_extractor/config/cell_mappings.py:155-228` | RAN T4 Part E | HIGH |
| TR-02 | A failed Power Query refresh leaves an unrefreshed `_Refreshed_` copy with the original's date; the client list then shows it as the last refresh and "skip refresh" reuses it as a success (✓ in the report). | `power_query_extractor/core/report_processor.py:164,137-155,376-393`; `power_query_extractor/gui/extractor_window.py:45-82` | RAN T10 (Excel stubbed) | HIGH |
| TR-03 | The "original" report is the first file the folder listing returns; it can be an earlier refreshed copy, giving double suffixes and refreshing a copy of a copy. | `report_processor.py:112-125` | RAN T5 | HIGH (order uncontrolled); MEDIUM (NTFS usually sorted) |
| TR-04 | An original report whose name contains the refresh marker (client name with "Refreshed", or "_v..." with a `_v{timestamp}` suffix) is treated as already refreshed; skip-refresh then skips the refresh. | `power_query_extractor/core/refresh_naming.py:49-55` | RAN T5 | HIGH |
| TR-05 | A hyphen in a client name shifts the parts: wrong client, wrong state, separate client row, files filed under the fragment. `Sales-Reco-...` is filed as Sales of client "Reco". | `utils/constants.py:15-55` (`[^-]+`) | RAN T1, T6 | HIGH |
| TR-06 | One client spelled differently (capitals, Pvt Ltd/Private Limited, extra space, state name vs code or number) becomes several client rows, each "incomplete"; Pvt Ltd/Private Limited variants then share one folder and overwrite each other's files (backup only in fresh mode). On Windows, capital-only differences also share one folder and one report name. | `core/file_parser.py:229-230` vs `utils/helpers.py:620-631,68-126` | RAN T6, T8; Windows part reasoned | HIGH; MEDIUM for Windows |
| TR-07 | Long client names that share their first 32 characters share one folder; Step 4 treats them as one client and refreshes/extracts only one report of each kind, possibly from different clients in one row. | `utils/helpers.py:634-639`; `report_processor.py:125`; `extractor_window.py:638-659` | RAN T6 s7, T9 | HIGH |
| TR-08 | Number-looking client or state text (`0042`, `07`, `27`, `+12`) comes back from the Validation table as a number: wrong key (`42-MH`) -> the whole processing run aborts at that client with "Fatal error" and no summary; a numeric state crashes Dry Run / Start / the folder-name toggle. | `gui/handlers/client_handler.py:166-171,181,191,218`; `gui/handlers/processing_handler.py:232,522-525` | RAN T7, T8, T12 | HIGH |
| TR-09 | "Resume" does not continue an interrupted run: it always makes a new Version folder for the current minute and copies everything again (it only skips within the same minute). | `core/file_organizer.py:103-108,283-287` | RAN T6 runs 4-5 | HIGH |
| TR-10 | Same-minute re-runs: fresh backs up then overwrites; rerun overwrites silently; reports and the Level-1 summary are overwritten without backup (the summary then lists only the last run's clients); two backups in one minute overwrite each other. | `core/file_organizer.py:277-287`; `utils/helpers.py:224-228`; `processing_handler.py:462-464` | RAN T6 runs 2-4, T8 s4 | HIGH |
| TR-11 | When a type has several files, the report template gets the alphabetically first (a "(1)" copy beats the plain name; `01042025` beats `31032025`; `2023-24` beats `2024-25`). Nothing tells the user. [GST RULE] (which period's file is used). | `core/excel_handler.py:629-699`; `utils/helpers.py:361` | RAN T3 b, T6 | HIGH |
| TR-12 | Per-client text report shows run-wide running totals for "Successfully Copied" and "Folders Created". | `core/file_organizer.py:363-368` | RAN T6 | HIGH |
| TR-13 | Scan ignores sub-folders, `.xlsb`, `.xltx`, `.csv`, files under 1 KB and files without an Excel signature, silently; `.xlsm` files are listed but can never match a pattern. | `utils/helpers.py:299-361`; `utils/constants.py` patterns end `\.xlsx?$` | RAN T1, T6 | HIGH |
| TR-14 | A name with a blank client counts as "Successfully Parsed" in the summary but is listed as a variation "Could not parse filename" - the summary's numbers disagree with the variation list. | `core/file_parser.py:210-220,311` | RAN T11 | HIGH |
| TR-15 | Client names with a dot ("Co.", "Pvt. Ltd.") are only partly cleaned: everything after the last dot is kept as-is (spaces, unabbreviated words) in report names and category labels. | `utils/helpers.py:76-77,123-124`; used at `processing_handler.py:292`, `core/file_organizer.py:90` | RAN T2 | HIGH |
| TR-16 | Consolidated report: a failed refresh shows `Missing` for every figure even when the figures were extracted; ✓/✗ and the Summary counts reflect only the refresh status. | `data_consolidator.py:109-110,128-129,235-239` | RAN T4 | HIGH |
| TR-17 | Sheet names in the report must match the mapping exactly, including capitals; otherwise all its figures are `Missing` while the row shows ✓. | `report_processor.py:564-576` | RAN T4 (fake COM) | HIGH |
| TR-18 | Text in a number cell is copied as text; 0 left unformatted; formulas without a cached value read as empty through openpyxl (Excel would compute them). | `data_consolidator.py:153-162` | RAN T4 | HIGH (formula-under-Excel part not verifiable) |
| TR-19 | Folder-name key with a length limit of 8 or less loses the state code entirely. | `utils/helpers.py:636-642` | RAN T2 | HIGH (reachable only by typing a small value: MEDIUM) |
| TR-20 | Refreshed-suffix setting edge cases: empty -> "same file" error; `{timestamp}` alone -> copies never recognised; no `{timestamp}` -> `_Refreshed_Refreshed` growth; `{TIMESTAMP}` stays literal. | `report_processor.py:128-130,164`; `refresh_naming.py:37-40` | RAN T5 | HIGH |
| TR-21 | "Newest folder" means most recently modified, so writing into an old run folder makes it the target of Re-Run/Resume and of the Step 4 scan. | `core/file_organizer.py:184`; `extractor_window.py:631,648` | RAN T3 c | HIGH (Linux); MEDIUM (Windows) |
| TR-22 | Setup checklist accepts a folder as a template and a file as source/target; the scan then fails with "Not a directory". | `gui/utils/setup_validation.py:47-64` | RAN T7 | HIGH |
| TR-23 | Any ZIP file renamed `.xlsx` passes the Excel check. | `utils/helpers.py:282-284` | RAN T8 | HIGH |
| TR-24 | Consolidated report, refreshed copies and backups all have minute-resolution names; a second run in the same minute overwrites without warning. | `data_consolidator.py:175-179`; `report_processor.py:128-164`; `utils/helpers.py:224-228` | RAN (backups), READ (others) | HIGH |
| TR-25 | Dead code/settings: `validate_results`, `cleanup_on_error`, `get_summary`, `_write_section`, `_create_details_sheet`, `get_unique_path`, `calculate_file_hash`, `get_relative_excel_path`, `get_date_only`, `validate_excel_file` and `find_excel_files` (both only called by tests; the scan uses their inner pieces); constants `EXCEL_EXTENSIONS`, `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB`, `EXCEL_SIGNATURES`; `EXCEL_SAFETY` values; options `auto_refresh_if_missing`, `show_timestamp`; parse `metadata` and `warnings`. | various (see sections 6, 10) | RAN grep | HIGH |
| TR-26 | The worker thread updates the log and shows message boxes directly at many points instead of through `_log_async` (Tk is not thread-safe). Not exercised with a real window. | `processing_handler.py:319-440,459-516` | READ | MEDIUM (risk), HIGH (that it happens) |

## 13. [GST RULE] register for what I traced (described only; no judgement)

| # | rule as coded | where | TIME-VARYING? |
|---|---|---|---|
| G1 | A client is "Complete" only with all six file types: GSTR-2B Reco, IMS Reco, GSTR-3B Export, Sales, Sales Reco, Annual Report. | `utils/constants.py:68-75`; `core/file_parser.py:252-286` | yes (which returns/reports exist) |
| G2 | Several GSTR-3B files per client are always allowed (comment: monthly filings); several files of any other type are flagged as duplicates. | `core/file_parser.py:264-269` | yes (filing periodicity) |
| G3 | "Critical" files = Annual Report and GSTR-3B Export (in `validate_results`, which is never called). | `core/file_parser.py:359` | yes |
| G4 | File naming conventions per type, incl. IMS date as exactly 8 digits (`DDMMYYYY` per the suggestion text) and a free-text period/year for the others. | `utils/constants.py:13-62`; `core/file_parser.py:288-306` | yes |
| G5 | State -> code table (39 names, two alternative spellings) and invented codes for anything else. | `utils/constants.py:251-291`; `utils/helpers.py:592-618` | yes |
| G6 | Which file type goes to which folder: GSTR-2B Reco and IMS Reco -> ITC; Sales and Sales Reco -> Sales; GSTR-3B -> GSTR3B; Annual Report -> the Version folder. | `utils/constants.py:16-58`; `core/file_organizer.py:305-312` | possibly |
| G7 | No financial-year / tax-period logic; when several periods exist, the alphabetically first file feeds the report. | `core/excel_handler.py:625-699` | yes |
| G8 | Extraction map: Diff GSTR-3B!J15/K15/L15 -> "Diff IGST/CGST/SGST"; Diff - GSTR-1!N1/O1/P1/Q1 -> "Diff Taxable Value/IGST/CGST/SGST"; Diff - GSTR-1!I1 -> "Taxable Value"; Reco!C37/D37/E37 -> "IGST/CGST/SGST Diff"; T6 of R9!K29 -> "Table 6J Match"; T8 of R9!G6 -> "Table 8D Limits"; IMS_Reco!AA1/AD1/AG1, ExtraInPurchase!O1/P1/Q1, ExtraInGSTR2B!S1/T1/U1 -> "IGST/CGST/SGST" (but see TR-01). | `power_query_extractor/config/cell_mappings.py:33-242` | yes (table numbers, layouts) |

## 14. Touch points with the author's Standards (only what my runs or reads showed)

- Standard 12, "Excel through COM" (rebuild_pipeline_FINAL.md:844-878): the extraction opens Excel with `Dispatch` (`report_processor.py:536`) and the refresh with `Dispatch`, `Visible = True` and keyboard shortcuts sent with `SendKeys` (`report_processor.py:218-227,260-293`), fixed waits instead of `CalculateUntilAsyncQueriesDone`; `Quit()` is in `finally` (404-409, 626-631); `CoInitialize`/`CoUninitialize` are inside the same functions that hold the Excel objects (215/412, 533/633). On a PC without pywin32 the main app builds reports with openpyxl (`core/excel_handler.py:96-102,282,306`), which the Standard says drops Power Query on save; the code only logs a warning (RAN: "Using openpyxl - Power Query content may be lost!") and does not check the template for a `DataMashup` part. READ; detail belongs to the Excel-specific notes.
- Standard 13: the extractor reads its cache from a dotfile in the user's home folder (`power_query_extractor/gui/extractor_window.py:118`) and the report builder stages files in the system temp folder (`core/excel_handler.py:53`), not under `%LOCALAPPDATA%\[App Name]\`. READ.

## 15. Possible client names

None found. In everything I read (the files listed in section 2, the test files listed there, the A/main commit subjects) the only names are the author's account and generic examples (`ABC`, `XYZ`, `ABC Private Limited`). All names in my scripts and outputs are invented (`Alpha Test Traders`, `Beta Demo Industries Private Limited`, `Gamma Sample Enterprises LLP`, `Gamma Sample Manufacturing Company of India One/Two`, `Delta Mock Co`, `Omega Co.` ...) and `<Client A>`. I did not open `gst_organizer.log`.

## 16. Questions for the author

1. When a client has two files of one type (two years, two IMS dates, or a "(1)" copy), which one should feed the ITC/Sales report? Today the alphabetically first name is used and nothing is shown (TR-11). [GST RULE]
2. For the consolidated report's "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" sheets, are the configured cells (IMS_Reco!AA1/AD1/AG1, ExtraInPurchase!O1/P1/Q1, ExtraInGSTR2B!S1/T1/U1) the ones you expect to see? Today those sheets show the Diff GSTR-3B cells instead (TR-01).
3. Is "Resume" meant to continue the earlier Version folder of an interrupted run (TR-09)?
4. Do file names ever carry a hyphen inside the client name, a number as the client name, or the state as a number / short form (e.g. "27", "J&K", "KA")? Each of these behaves differently today (TR-05, TR-06, TR-08). [GST RULE] for the state-code table: is `utils/constants.py:251-291` the list you want?
5. Do the ITC/Sales templates' Power Query queries read every file in the linked folder, or only the named file? This decides how much the folder-merging cases (TR-06, TR-07) matter.
6. After a failed refresh, should the half-made `_Refreshed_` copy be kept, and should the figures that were still extracted appear in the consolidated report (TR-02, TR-16)?
7. Should the scan look inside sub-folders of the source, and should `.xlsm` / `.xlsb` source files be accepted (TR-13)?
8. Should a client with only duplicate files (nothing missing) count under "Clients with Missing Files"?

## 17. INCOMPLETE

- Not executable here, so only read: everything that needs Windows or Excel - the Power Query refresh (`report_processor.py:209-416`), the COM report builder (`core/excel_handler.py:104-260`), error-dialog handling with `win32gui`, `os.startfile`. Real Excel's values for formula cells (TR-18) are not observable with openpyxl. Windows-specific behaviour (NTFS listing order, case-insensitive names, folder time updates) is reasoned, not run (marked MEDIUM).
- `core/excel_handler.py` was read only at 1-110, 260-325 and 600-713; `create_summary_report` and its sheet builders (371-601) were exercised (RAN, output in T6) but not read line by line by me.
- The real Tk windows (Setup, Validation, Processing, Extract tabs) were not opened; only a bare Treeview was used (T7, T12). Thread-safety (TR-26) not exercised.
- Not run: the extractor's `start_processing` / `_process_clients_thread` GUI paths and its cache loading (they need the full panel); I called `ExtractorPanel.scan_folder` and `ReportProcessor` directly instead.
