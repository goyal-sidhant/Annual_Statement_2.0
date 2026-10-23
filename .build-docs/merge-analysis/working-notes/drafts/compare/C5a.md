## 5. Rules side by side

This section puts every business rule (a fixed decision the program makes on its own, such as "a file whose name starts with `ImsReco-` is an IMS reconciliation") of project A next to the matching rule of project B. This first part covers the rule categories 1 to 8 of the two registries (the numbered lists of all features and rules made for each audit): input-file recognition, client identity and name cleaning, state codes, financial year and tax period, expected files and completeness, folder layout and naming, copying and versioning, and report creation with the template "Links" cells. Categories 9 to 16 (Power Query refresh onwards) follow in the next part.

Reminder of the two projects: **A** = `AnnualStatement_oldv1` at `bf2f1dd` (27-07-2026, Tkinter window, "v3.0", 288 tests); **B** = `AnnualStatement_oldv2` at `b0e7d0e` (09-02-2026, PyQt5 window, "v4.0", no tests). Despite the names, A is not the older one. Both share their first 12 commits (`edac6a3` 31-08-2025 … `47a1492` 30-10-2025) and split after `47a1492`.

### How to read this section

Each category starts with one table that lists **every** rule of both projects in that category, in A's order, followed by the B rules that no A rule covers. The columns are:

| Column | What it holds |
|---|---|
| A rule / B rule | The rule's ID in each registry (A-R#, B-R#). A rule that one registry files under another category is shown where its partner sits, with a pointer in the other category. |
| What it decides | The rule in plain words. |
| A's version | What A does, with the place in A's code (`A:path:line`). |
| B's version | What B does, with the place in B's code (`B:path:line`). |
| Relation | **same** = the same behaviour and values in both; **different** = the behaviour or values differ (a numbered worked example follows the table); **only A** / **only B** = the rule exists in one project only. |
| Evidence | How it was checked (see the legend below). |

Every **different** row is followed, under the table, by a numbered block (for example "1.1") with: what each project does; one realistic, made-up example that was **run through the real code of both projects** with both results shown; what the difference changes in a client's annual working (which figure, report cell, folder or file); and, for every difference, when each version came in (commit, date, message) and any hint in the code or messages about why.

| Mark | Meaning |
|---|---|
| **[GST RULE]** | The rule decides something only the author can decide: a tax figure; which return, table, sheet or cell a figure belongs to; a rate, limit or threshold; rounding; the financial year or tax period; GSTIN, PAN or state-code logic; ITC eligibility or reversal; a due date. When in doubt the rule is tagged (if either registry tags it, it is tagged here). For these rules this section only describes; it never says which version is right and never recommends one. |
| **⚠** | A value that can change over time and must be checked by the author: return and table names, file-naming conventions of the portal, the state list, date formats, template cell positions, Excel limits. |
| **READ** | Established by reading the code of both projects. |
| **RAN** | The real functions of both projects were executed on the same made-up inputs (see "How the runs were done"). |
| **HIGH / MEDIUM / LOW** | Confidence of a statement that could not be run here. |
| `A:path:line`, `B:path:line` | A place in A at `bf2f1dd` or in B at `b0e7d0e`. `path:line@commit` is a place in older code. |
| "technical point" | A remark about crashes, file safety, speed or clarity of messages. Such remarks never judge a GST rule. |

### How the runs were done

- Both projects were copied to `/root/merge-scratch/run/C5a/A` and `/root/merge-scratch/run/C5a/B` and run there with `python -B` (Python 3.14; project A with its own `venv-A`, project B with `venv-B`). Nothing was run inside the original repositories. The scripts are in `/root/merge-scratch/run/C5a/scripts/` and their outputs in `/root/merge-scratch/run/C5a/out/`:
  - `core_probe.py` - file names through both name parsers (the code that splits a file name into type, client, state and period), a source folder with unusual entries, the name cleaner, the client-folder name builder, state codes, the folder builder ("organiser") and the Links values with three test templates;
  - `merged_probe.py` - a Links cell inside a merged cell range;
  - `step3_probe.py` - each project's real Step 3 worker (the code behind "Start Processing") on three made-up clients, including a change of template during the run (`out/step3_A.json`, `out/step3_B.json`) and a damaged template (`out/step3broken_*.json`);
  - `step4_probe.py` and `step4_probe_newonly.py` - each project's real Step 4 code that chooses the run folder and the version folders;
  - `gui_probe_A.py`, `gui_probe_B.py`, `display_probe.py` - the real Step 1 and Step 2 screen code (A's Tkinter screens under `xvfb-run`, a virtual screen; B's Qt screens in Qt's "offscreen" mode, which draws without a monitor);
  - `key42_probe.py`, `collide_probe.py`, `modes_probe.py`, `ims_only1_probe.py` - smaller single-question runs named in the blocks below.
- The clock was fixed at 21-10-2026 10:30 inside each run by replacing the time-stamp function in memory (the code on disk was not changed), so folder names read `Annual Statement-211026 1030` and `Version-211026 1030`.
- All clients are invented: `ClientA` (Delhi), `ClientB PRIVATE LIMITED` (Maharashtra), `ClientD` (Tamil Nadu), `ClientE` and `ClientF` (Delhi), `0042` (Maharashtra; a client written as a number), `Zeta Traders` and `Alpha Traders` (Delhi). The "Excel files" are 2 KB stand-ins that start with the bytes of a real Excel file; they hold no return data.
- This machine is Linux without Excel. Everything Excel itself does through COM (COM is Windows' way for one program to drive another; here Python drives Excel) could only be **READ**. Report building was **RAN** through each project's openpyxl path (openpyxl is a Python library that writes `.xlsx` files without Excel; both projects fall back to it when Excel cannot be driven). Where a result could differ on Windows this is said, with a confidence mark.
- B's window code imports pywin32 (the Python add-on for driving Excel on Windows) when its modules are loaded, so on Linux those modules cannot even be imported. For the B screen runs, empty stand-in modules named `pythoncom`, `win32com`, `win32api`, `win32con`, `win32gui` were placed in memory first (`scripts/fake_win32.py`) only so that the import succeeds. None of their functions is called on the code paths that were run.
- Function-by-function comparison: `scripts/funcdiff.py` compared the source text of every function in `utils/helpers.py`, `core/file_parser.py`, `core/file_organizer.py` and `core/excel_handler.py` of both projects (RAN). It is cited below as "funcdiff: IDENTICAL" or "funcdiff: DIFF".

### Category 1: Input-file recognition

This category decides which files in the source folder (the folder of downloaded returns and working files for the season) are looked at, and which client, state, return type and period each file name is read as. All six file-name patterns are regular expressions (text patterns the computer matches against a file name) kept in one table, `FILE_PATTERNS`, in each project.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R1 | B-R6 | GSTR-2B Reco file name ⚠ [GST RULE] | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$`, any letter case; parts client, state, period. A:utils/constants.py:14-21 (pattern :15) | Identical. B:utils/constants.py:14-21 (pattern :15) | same | RAN core_probe: `GSTR-2B-Reco-ClientA-Delhi-Apr25-Mar26.xlsx` gives period `Apr25-Mar26` in both; with " (1)" before `.xlsx` the period becomes `Apr25-Mar26 (1)` in both. Since `edac6a3` (shared). |
| A-R2 | B-R7 | IMS Reco file name ⚠ [GST RULE] | `^ImsReco-([^-]+)-([^-]+)-(\d{8})(?:\s*\(\d+\))?\.xlsx?$` - also accepts one "(N)" or " (N)" after the 8-digit date. A:utils/constants.py:23 | `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$` - nothing may follow the date. B:utils/constants.py:23 | **different** | 1.1 below (RAN) |
| A-R3 | B-R8 | GSTR-3B Export file name ⚠ [GST RULE] | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$`; a leading "(1) " is accepted, a trailing " (1)" ends up inside the month. A:utils/constants.py:30-37 (pattern :31) | Identical. B:utils/constants.py:30-37 (pattern :31) | same | RAN core_probe: `(1) GSTR3B-ClientA-Delhi-Apr.xlsx` gives month `Apr`; `GSTR3B-ClientA-Delhi-Apr (1).xlsx` gives month `Apr (1)` - in both. |
| A-R4 | B-R9 | Sales file name ⚠ [GST RULE] | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$`; exactly four parts (client, state, start month, end month). A:utils/constants.py:38-45 (pattern :39) | Identical. B:utils/constants.py:38-45 (pattern :39) | same | RAN core_probe: `Sales-ClientA-Delhi-Apr25-Mar26 (1).xlsx` gives start `Apr25`, end `Mar26 (1)` in both. |
| A-R5 | B-R10 | Sales Reco file name ⚠ [GST RULE] | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$`. A:utils/constants.py:46-53 (pattern :47) | Identical. B:utils/constants.py:46-53 (pattern :47) | same | RAN core_probe: period `FY2025-26` in both. |
| A-R6 | B-R11 | Annual Report file name ⚠ [GST RULE] | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$`. A:utils/constants.py:54-61 (pattern :55) | Identical. B:utils/constants.py:54-61 (pattern :55) | same | RAN core_probe: year `2025-26`, and `2025-26 (1)` for a " (1)" name, in both. |
| A-R7 | B-R5 | How a name is matched [GST RULE]: patterns tried in a fixed order, first match wins, parts trimmed, empty parts skipped | A:core/file_parser.py:82-134 | B:core/file_parser.py:78-124 | same | READ: the two functions differ only in A's extra debug lines for Sales files (A:core/file_parser.py:77-79, 123-126). RAN core_probe: 15 names give the same results except the IMS "(N)" names of 1.1. |
| A-R8 | B-R1 | Which files the scan looks at | One folder listing; names ending `.xlsx`, `.xls`, `.xlsm` in any letter case; regular files only (a folder named like a workbook is skipped). A:utils/helpers.py:309-365 (extensions :327, test :330, regular file :296) | Three wildcard searches (`*.xlsx`, `*.xls`, `*.xlsm`; case-sensitive on Linux), names starting `~$` skipped, folders not excluded. B:utils/helpers.py:249-266 | **different** | 1.2 below (RAN) |
| A-R9 | B-R2 | Minimum size 1024 bytes | Smaller files dropped with no message and no count. A:utils/helpers.py:302-303 | Smaller files listed as the variation (the code's word for a problem file) "File too small to be valid Excel". B:core/file_parser.py:181-191 | **different** | 1.3 below (RAN) |
| A-R10 | B-R3 | Excel signature (first 8 bytes must be `PK\x03\x04` or `\xd0\xcf\x11\xe0`) | Wrong signature dropped with no message and no count. A:utils/helpers.py:263-287 | Listed as the variation "Invalid or corrupted Excel file" (unreadable: "Cannot read file: <error>"). B:core/file_parser.py:193-215 | **different** | 1.4 below (RAN) |
| (no A rule; A skips silently at A:utils/helpers.py:332-338) | B-R4 | A file whose size and dates cannot be read | Skipped; nothing on screen, nothing in the log file | Skipped; nothing on screen; the log file (`gst_organizer.log`) gets "Cannot stat <name>: <error>". B:core/file_parser.py:174-179 | **different** | 1.5 below (RAN) |
| A-R11 | B-R5 | Unrecognised names become variations ("No matching pattern found"; a match with an empty client or state gives "Could not parse filename") | A:core/file_parser.py:130-132, 209-220 | B:core/file_parser.py:120-122, 245-252 | same (for names); B's list additionally holds the rejects of 1.3 to 1.5 | RAN core_probe: `Notes for ClientA.xlsx` → "No matching pattern found" in both. funcdiff: `_process_file` DIFF only in the size/signature part. |
| A-R12 | B-R12 | Hint for an unrecognised name ⚠ [GST RULE] | Same hint texts and order; the hint is shown in the Step 3 summary workbook's "Variations" sheet. A:core/file_parser.py:288-306; list handed over at A:gui/handlers/processing_handler.py:224; sheet A:core/excel_handler.py:554-585 | Same hint computed, never shown: the summary workbook receives an empty list. B:core/file_parser.py:316-334; empty list at B:gui/controllers/processing_controller.py:73 | **different** | 1.6 below (RAN) |
| A-R13 | B-R14 | `EXCEL_EXTENSIONS = {'.xlsx', '.xls', '.xlsm'}` - never used | A:utils/constants.py:196 | B:utils/constants.py:179 | same (dead in both) | RAN: a search of both copies finds no use outside the definition |
| A-R14 | B-R16 | `MAX_FILE_SIZE_MB = 100` - never used; no upper size limit anywhere | A:utils/constants.py:198 | B:utils/constants.py:181 | same (dead in both) | RAN: a search of both copies finds no use outside the definition |
| A-R15 | B-R15 | `EXCEL_SIGNATURES` - never used (the bytes are typed again elsewhere) | A:utils/constants.py:199-202 | B:utils/constants.py:182-185 | same (dead in both) | RAN: a search of both copies finds no use outside the definition |
| A-R16 | B-R13 (with B-R141) | The patterns' extra labels ⚠ [GST RULE]: `category` and `folder` are copied into every file record and never read in either; the `description` (example name) is read by A's Step 1 card and by nothing in B | A:utils/constants.py:17-19, 25-27, 33-35, 41-43, 49-51, 57-59; card A:gui/tabs/setup_tab.py:477-503 | B:utils/constants.py:17-19, 25-27, 33-35, 41-43, 49-51, 57-59; B's card is typed by hand at B:gui/views/setup_view.py:488-499 | **different** (only the example names shown on Step 1) | 1.7 below (RAN) |

#### 1.1 IMS Reco names with a browser "(N)" after the date (A-R2 / B-R7) · [GST RULE] · ⚠

**What A does.** A name such as `ImsReco-ClientA-Delhi-31032026 (1).xlsx` or `…31032026(2).xlsx` (the "(1)" a web browser adds when the same file is downloaded twice) is an IMS Reco file of ClientA; the date part stays `31032026`. Only one such marker is allowed. A:utils/constants.py:23.

**What B does.** Nothing may follow the 8 digits; the same name is unrecognised ("No matching pattern found"). B:utils/constants.py:23.

**Worked example (RAN, `core_probe.py` and `ims_only1_probe.py`).**

| Input | A | B |
|---|---|---|
| `ImsReco-ClientA-Delhi-31032026 (1).xlsx` | IMS Reco, client `ClientA`, state `Delhi`, date `31032026` | variation "No matching pattern found" |
| `ImsReco-ClientA-Delhi-31032026(2).xlsx` | IMS Reco, date `31032026` | variation "No matching pattern found" |
| `ImsReco-ClientA-Delhi-31032026 (1) (2).xlsx` | variation (only one marker allowed) | variation |
| ClientA's folder holds both `ImsReco-ClientA-Delhi-31032026.xlsx` and the " (1)" copy, plus the other five types: Step 2 status | "Has duplicates", remark "Multiple IMS Reco (2 files)" | "Complete" |
| Same folder: files copied for ClientA | 8 (the duplicate as `ImsReco-ClientA-Delhi-31032026_(1).xlsx` in the ITC folder) | 7 |
| Same folder: ITC report Links cell B11 (the IMS file name the template reads) | `ImsReco-ClientA-Delhi-31032026_(1)` - the duplicate, because a space sorts before a dot (rule A-R57) | `ImsReco-ClientA-Delhi-31032026` |
| ClientF has **only** `ImsReco-ClientF-Delhi-31032026 (1).xlsx` (plus the other five types): status | "Complete"; B10 and B11 written (`…_(1)`) | "Missing 1 files" (IMS Reco); B10 and B11 not written, so the template's own text stays in them; the file counted only as a variation |
| Step 3 summary workbook for the same three-client folder (`step3_probe.py`) | "Total Files" 10; ClientA "Files" 7 | "Total Files" 9; ClientA "Files" 6 |

**What it changes in the annual working.** Which IMS reconciliation file feeds the ITC report's IMS part (Links B10/B11, and from there the IMS figures the template's queries compute). With a re-downloaded IMS file in the folder, A links the "(1)" copy (whether that is the newer download is not checked by either project), and Step 2 flags the client "Has duplicates"; B ignores the "(1)" copy and links the original. With only a "(1)" download, A treats IMS Reco as present; B reports the client as missing IMS Reco and leaves the IMS cells of the ITC report as the template had them. For all other types a trailing " (1)" is accepted by both and simply becomes part of the period text (RAN, rows A-R1, A-R3 to A-R6).

**When each version came in.** B's pattern is the original one from `edac6a3` (31-08-2025, "Initial commit"), never changed in B. A's came in `5c8ce95` (23-07-2026, "Detect IMS Reco files with a trailing (N) duplicate suffix"). Its message says such names "were not detected: the pattern required exactly 8 digits immediately before the extension", that "the date group still captures exactly the 8 digits, so metadata is unchanged", and that "all other types already tolerated the suffix via their flexible trailing group". A's tests pin the new behaviour (A:tests/test_file_patterns.py:47-52, 63-65); run against B's code these tests fail (four failures, recorded in `notes2/A_tests.md`). Neither project's Step 1 help mentions the "(N)" form (1.7). Which version to keep is the author's decision.

#### 1.2 Which files the scan looks at (A-R8 / B-R1)

**What A does.** One listing of the top level of the source folder. A name is kept if, in lower case, it ends in `.xlsx`, `.xls` or `.xlsm`; it must be a regular file (A:utils/helpers.py:296), at least 1 KB (1.3) and carry an Excel signature (1.4). Anything else is dropped without a message. A:utils/helpers.py:309-365.

**What B does.** Three wildcard searches (`*.xlsx`, `*.xls`, `*.xlsm`) of the top level; names starting `~$` (Excel's lock files, made while a workbook is open) are dropped; everything else found goes to the parser, which applies 1.3 to 1.5. On Linux the wildcard search respects letter case; on Windows Python's search ignores it (HIGH). B:utils/helpers.py:249-266; parser B:core/file_parser.py:171-256.

**Worked example (RAN, `core_probe.py`, folder `src_discovery`; B ran on Linux).**

| Entry in the source folder | A | B |
|---|---|---|
| `GSTR3B-ClientA-Delhi-Apr.xlsx`, `GSTR3B-ClientA-Delhi-May.xls` | listed, GSTR-3B | listed, GSTR-3B |
| `GSTR3B-ClientA-Delhi-Jun.XLSX` (capital extension) | listed, GSTR-3B | not found at all on Linux; on Windows found (HIGH, not run) |
| `~$GSTR3B-ClientA-Delhi-Apr.xlsx` (165-byte lock file) | dropped (under 1 KB), no message | skipped by its `~$` name |
| `~$Big-lock.xlsx` (2 KB, Excel signature) | listed → variation "No matching pattern found" | skipped by its `~$` name |
| `GSTR3B-ClientA-Delhi-Sep.xlsm` | listed → variation (no pattern accepts `.xlsm`) | listed → the same variation |
| `Old folder.xlsx` (a **folder**, not a file) | skipped, no message | listed → variation "Cannot read file: Is a directory: …" (on Linux). On Windows the size of a folder is reported as 0, so the variation would probably read "File too small to be valid Excel" (MEDIUM, not run) |
| `GSTR3B-ClientA-Delhi-Dec.csv`, `sub/GSTR3B-ClientA-Delhi-Nov.xlsx` | ignored (wrong extension; sub-folder) | ignored |
| Result for ClientA | three GSTR-3B months (Apr, May, Jun) | two (Apr, May) on Linux |

**What it changes in the annual working.** On the office's Windows PCs the capital-letter case does not apply (HIGH), so in practice the two scans differ in three places: a lock-named file of 1 KB or more (rare) is a variation in A and invisible in B; a folder whose name ends in `.xlsx` adds a variation in B and is invisible in A; and files that are too small or not Excel are invisible in A but counted in B (1.3, 1.4). None of this changes which genuine return file goes where.

**When each version came in.** At the split both projects used B's three wildcard searches and dropped small and non-Excel files silently (`47a1492:utils/helpers.py:249-266`, where `validate_excel_file` filtered them). A replaced the searches with one listing in `a603c02` (23-07-2026, "Speed up network file scanning (behavior-preserving)"), then reused the listing's file details in `56aaad9` (25-07-2026, "Reuse directory-listing metadata: 81% fewer round-trips per file") and read the signatures 16 at a time in `4f84ba9` (25-07-2026, "Read Excel signatures concurrently: 11 min -> 41s on the real share"); the messages give speed on the network share as the reason. B added the `~$` filter and moved the size and signature checks into the parser in `7e09e5c` (08-02-2026, "feat: Refactor Setup and Validation Views with Collapsible Cards and Colored Sections"; the message says only "Added functionality to handle Excel file selection and validation more efficiently"). Not a GST rule.

#### 1.3 Files under 1 KB (A-R9 / B-R2)

**What A does.** Drops them with no message, no count and no log line. A:utils/helpers.py:302-303.
**What B does.** Lists each one as a variation "File too small to be valid Excel" with its size. B:core/file_parser.py:181-191.

**Worked example (RAN, `core_probe.py`).** `GSTR3B-ClientA-Delhi-Jul.xlsx` of 500 bytes (for example a download that broke off): A - not in any list or count; B - variation `{"filename": "GSTR3B-ClientA-Delhi-Jul.xlsx", "issue": "File too small to be valid Excel", "size": 500}`, counted under "Variations" on Steps 1 and 2.

**What it changes.** In A a broken July GSTR-3B download disappears without trace; ClientA simply has one month fewer in the GSTR-3B folder. In B the variation count rises by one, but the name of the file is shown nowhere on screen, because B never displays its variation list (1.6). Technical point: B at least counts the file; neither tells the user which file it was.

**When.** Silent drop since `edac6a3` (shared; at the split `47a1492:utils/helpers.py:260`). B's variation since `7e09e5c` (08-02-2026). A unchanged in behaviour. Not a GST rule.

#### 1.4 Files without an Excel signature (A-R10 / B-R3)

**What A does.** Reads the first 8 bytes; anything that does not start like a ZIP-based Excel file (`PK\x03\x04`) or an old `.xls` (`\xd0\xcf\x11\xe0`) is dropped with no message. A:utils/helpers.py:263-287.
**What B does.** The same test, but the file becomes the variation "Invalid or corrupted Excel file"; if it cannot be read, "Cannot read file: <error>". B:core/file_parser.py:193-215.

**Worked example (RAN, `core_probe.py`).** `GSTR3B-ClientA-Delhi-Aug.xlsx` that is really a 2,240-byte text file: A - invisible; B - variation "Invalid or corrupted Excel file", size 2240.

**What it changes.** As 1.3: in A a mis-saved August GSTR-3B (for example a portal error page saved as `.xlsx`) vanishes from the client's working; in B it is counted but not named on screen.

**When.** As 1.3: silent in the shared code (`47a1492`), variation in B since `7e09e5c` (08-02-2026). Not a GST rule.

#### 1.5 Entries whose size cannot be read (no A rule / B-R4)

**What A does.** If the listing finds a name but its details cannot be read (for example a link whose target has gone, or a file deleted from the share between listing and reading), the entry is skipped with no message and no log line. A:utils/helpers.py:332-338.
**What B does.** The same skip, but the log file gets "Cannot stat <name>: <error>". B:core/file_parser.py:174-179. Any other per-file error goes into a list nothing reads, in both (A:core/file_parser.py:222-224; B:core/file_parser.py:254-256).

**Worked example (RAN, `core_probe.py`).** A broken link named `GSTR3B-ClientA-Delhi-Oct.xlsx`: A - nothing anywhere; B - nothing on screen, log line "WARNING core.file_parser: Cannot stat GSTR3B-ClientA-Delhi-Oct.xlsx: [Errno 2] No such file or directory: …".

**What it changes.** Only whether a trace is left in the log file for later investigation; on screen both are silent.

**When.** A's silent skip came with the single listing (`a603c02`, 23-07-2026; at the split such a file was dropped silently inside `validate_excel_file`, `47a1492:utils/helpers.py:260`). B's log line since `7e09e5c` (08-02-2026). Not a GST rule.

#### 1.6 Showing the hint for an unrecognised file name (A-R12 / B-R12) · [GST RULE] · ⚠

**What both do.** For every unrecognised name a hint is chosen by plain text tests on the lower-cased name, first hit wins: contains "gstr" and "2b" → `GSTR-2B-Reco-ClientName-State-Period.xlsx`; "ims" → `ImsReco-ClientName-State-DDMMYYYY.xlsx` ⚠; "gstr" and "3b" → `GSTR3B-ClientName-State-Month.xlsx`; "sales" and "reco" → `SalesReco-ClientName-State-Period.xlsx`; "sales" → `Sales-ClientName-State-StartMonth-EndMonth.xlsx`; "annual" → `AnnualReport-ClientName-State-Year.xlsx`; otherwise "Check file naming convention". funcdiff: `_suggest_pattern` IDENTICAL.

**What differs.** A hands the scan's variation list to the summary workbook (A:gui/handlers/processing_handler.py:224), which writes it to the "Variations" sheet with the hint as "Expected Pattern" (A:core/excel_handler.py:554-585). B's worker builds its summary with `'variations': []` (B:gui/controllers/processing_controller.py:73); the scan's list stays in memory and is never displayed.

**Worked example (RAN, `core_probe.py` and `step3_probe.py`).** Source folder with `Claims register ClientA.xlsx` (and, for B, the IMS "(1)" file of 1.1):

| | A | B |
|---|---|---|
| Hint computed for `Claims register ClientA.xlsx` | `ImsReco-ClientName-State-DDMMYYYY.xlsx` (the word "claims" contains "ims") | the same |
| Hint for `IMS Reco-ClientA-Delhi-31032026.xlsx` (space in "IMS Reco") | `ImsReco-ClientName-State-DDMMYYYY.xlsx` | the same |
| Summary workbook, sheet "Variations" | one row: `Claims register ClientA.xlsx` · "No matching pattern found" · `ImsReco-ClientName-State-DDMMYYYY.xlsx` · 2004 · "Rename file" | "No file variations found" (although B's scan held two variations) |

**What it changes.** In A, after Step 3 the staff can open `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` and see which source files were not used and how they should have been named (for example an IMS download saved as `IMS Reco-…` with a space). In B the only trace is the number "Variations: n" on Steps 1 and 2; the file names and hints are not shown anywhere, so a mis-named file of a client is noticed only if someone spots the missing type.

**When.** The hint and the Variations sheet date from `edac6a3` (31-08-2025, shared); at the split the Tk code passed the real list (`47a1492:gui/handlers/processing_handler.py:161`). B's empty list came with `c258047` (07-02-2026, "Refactor GST File Organizer v3.0 -> v4.0: PyQt5 migration, unified app, enhanced caching"); its message does not mention variations. The new worker is given only the client data and the ticked keys (B:gui/pyqt_main_window.py:366-371), not the scan's variation list (READ, HIGH that this is how the list was lost; why is not stated anywhere). Which version to keep is the author's decision.

#### 1.7 Example file names shown on Step 1 (A-R16 / B-R13, with B-R141) · [GST RULE] · ⚠

**What A does.** The "EXPECTED FILE NAMES" card on Step 1 is built in a loop over `FILE_PATTERNS` and shows each pattern's type and `description` text (A:gui/tabs/setup_tab.py:477-503; texts at A:utils/constants.py:19, 27, 35, 43, 51, 59).
**What B does.** The card "EXPECTED FILE NAMES / MODES" shows a list typed directly in the screen code (B:gui/views/setup_view.py:488-499); the patterns' `description` texts are never read.

**Worked example (RAN, `gui_probe_A.py`, `gui_probe_B.py` - the text of the real widgets).**

| Type | A's card | B's card |
|---|---|---|
| GSTR-2B Reco | `GSTR-2B-Reco-ClientName-State-Period.xlsx` | `GSTR-2B-Reco-Client-State-Period` |
| IMS Reco | `ImsReco-ClientName-State-DDMMYYYY.xlsx` | `ImsReco-Client-State-DDMMYYYY` |
| GSTR-3B Export | `GSTR3B-ClientName-State-Month.xlsx` | `GSTR3B-Client-State-Month` |
| Sales | `Sales-ClientName-State-StartMonth-EndMonth.xlsx` | `Sales-Client-State-Start-End` |
| Sales Reco | `SalesReco-ClientName-State-Period.xlsx` | `SalesReco-Client-State-Period` |
| Annual Report | `AnnualReport-ClientName-State-Year.xlsx` | `AnnualReport-Client-State-Year` |

Both name the same six types in the same order. A shows `.xlsx` and the type label; B shows no extension. Neither mentions that `.xls` is also accepted, that a hyphen inside a client or state name breaks the reading (rule A-R7/B-R5), or the "(1)" forms each project accepts (1.1). Technical point: A's code comment says the card "can never drift from the regexes the parser actually uses" (A:gui/tabs/setup_tab.py:478-479), but the card shows the hand-written `description` text kept next to each pattern, not the pattern itself; the IMS description (A:utils/constants.py:27) does not reflect the "(N)" that A's IMS pattern (A:utils/constants.py:23) accepts.

**What it changes.** Only what staff are told when they rename downloads at the start of the season; the recognition itself is 1.1 and the other pattern rows.

**When.** B's list is the shared original text: the Tk card at the split held the same six hand-typed lines (`47a1492:gui/tabs/setup_tab.py:361-368`, from `edac6a3`). B's first PyQt5 screen (`c258047`, 07-02-2026) showed an HTML list with `.xlsx` names (`c258047:gui/views/setup_view.py:169` for IMS); `7e09e5c` (08-02-2026) brought back the short hand-typed list. A replaced the list in `322717f` (25-07-2026, "Restructure the Setup tab: live validation checklist"). Which wording to keep is the author's decision.

### Category 2: Client identity and name cleaning

This category decides which files belong to one client, and how a client's name is cleaned before it is used in folder and file names.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R17 | B-R17 | The client key [GST RULE]: files are grouped under "<client as written in the file name>-<state code>"; the key handed from Step 2 to Step 3 | Built at the scan, A:core/file_parser.py:226-236. Step 2 rebuilds the key from the text of the table row, A:gui/handlers/client_handler.py:161-173 | Built the same way, B:core/file_parser.py:258-268. Step 2 stores the key inside the row and hands it back unchanged, B:gui/views/validation_view.py:371, 412-421 | **different** (only the key Step 2 hands back) | 2.1 below (RAN) |
| A-R18 | B-R18 | Client and state text checks (empty, shorter than 2, longer than 100, characters `< > : " / \ ? *` and the vertical bar); the result is stored and never read | A:utils/helpers.py:428-449; called A:core/file_parser.py:108-119 | B:utils/helpers.py:321-342; called B:core/file_parser.py:104-115 | same (no effect in either) | funcdiff: `validate_client_name` IDENTICAL. RAN core_probe: `''`, `'A'`, 101 characters, `'a<b'`, `'Delhi'` give the same five answers in both. |
| A-R19 | B-R19 | "Private" → "Pvt", "Limited" → "Ltd" | File names (copies, report names, the bracketed name in category folders): whole words, any letter case, after turning underscores into spaces, A:utils/helpers.py:91-93. Client folder: whole words, any case, A:utils/helpers.py:624-625 | File names: plain, case-sensitive replacement anywhere in the text, B:utils/helpers.py:77-79. Client folder: whole words, any case (identical to A), B:utils/helpers.py:515-518 | **different** for file names; same for the client folder | 2.2 below (RAN) |
| A-R20 | B-R20 | Name cleaning (`sanitize_filename`): characters replaced, runs of spaces and underscores joined, length cut to 185 + extension, "unnamed" for empty text | A:utils/helpers.py:68-126 | B:utils/helpers.py:66-112 | **different** only through the abbreviation step of A-R19 / B-R19 | funcdiff: DIFF. RAN core_probe: the other test names give the same result in both, e.g. `ClientA/Delhi Branch` → `Delhi_Branch` (text before a slash is lost in both), `ClientB Pvt. Ltd` → `ClientB_Pvt. Ltd` (text after the last dot treated as an extension in both), `ClientA's #1, Shop!` → `ClientA_s_1_Shop`, `''` and `'___'` → `unnamed`. |

#### 2.1 The client key handed back by Step 2 (A-R17 / B-R17) · [GST RULE]

**What both do at the scan.** A client is "<client text exactly as written>-<state code>". For `AnnualReport-0042-Maharashtra-2025-26.xlsx` (a client written as an office code number) the key is `0042-MH` in both (RAN).

**What A does when the user ticks clients.** Step 2's table (a Tk "Treeview") gives back the cell texts; Tk turns number-looking text into numbers, so `0042` comes back as the number 42. A rebuilds the key as `f"{values[0]}-{state_code}"` and gets `42-MH`, which does not exist (A:gui/handlers/client_handler.py:167-170). Step 3 looks the key up outside its per-client protection (A:gui/handlers/processing_handler.py:232), so the run stops.

**What B does.** The key is stored inside the row when the table is filled and handed back as it is (B:gui/views/validation_view.py:371, 418).

**Worked example (RAN, `gui_probe_A.py`, `gui_probe_B.py`, `key42_probe.py`).**

| | A | B |
|---|---|---|
| Value read back from the Step 2 table for client `0042` | `42` (a number) | `0042` |
| Keys handed to Step 3 after ticking all three clients | `42-MH`, `ClientA-DL`, `ClientE-DL` (`42-MH` is not a scanned client) | `0042-MH`, `ClientA-DL`, `ClientE-DL` |
| Start Processing with `42-MH` first | "💥 Fatal error: '42-MH'" and an error box; nothing written for any client, no summary workbook | (not affected) |

**What it changes.** A client whose name in the file names is a pure number with leading zeros cannot be processed in A, and the run stops at that client: clients ticked after it are not processed and no summary workbook is written (clients before it are processed; READ, HIGH, from the loop at A:gui/handlers/processing_handler.py:228-232). Other number-looking names were not run. Technical point: this is a crash in A's screen code; the key-building rule itself is the same in both projects, and nothing here judges it.

**When.** Key building since `edac6a3` (shared). A's rebuild from the table text is the shared Tk code (`47a1492:gui/handlers/client_handler.py:161-170`). B's stored key came with the PyQt5 screen in `c258047` (07-02-2026); the message does not mention it.

#### 2.2 "Private"/"Limited" in file names (A-R19, A-R20 / B-R19, B-R20)

**What A does.** In every cleaned file name: runs of spaces and underscores become one space, then whole words "Private" and "Limited" in any letter case become "Pvt" and "Ltd" (A:utils/helpers.py:91-93). Words that merely contain them ("Privateer", "Unlimited") are left alone.
**What B does.** `base_name.replace('Private', 'Pvt')` and `.replace('Limited', 'Ltd')`: exact capitals only, anywhere in the text (B:utils/helpers.py:77-79).
**Same in both.** The client-folder name (Level 2) uses whole words in any case and does not first turn underscores into spaces (A:utils/helpers.py:624-625 = B:utils/helpers.py:515-518).

**Worked example (RAN, `core_probe.py`, `collide_probe.py`, `step3_probe.py`).**

| Client text in the file names | A: cleaned name | B: cleaned name | Client folder (both) |
|---|---|---|---|
| `ClientB PRIVATE LIMITED` | `ClientB_Pvt_Ltd` | `ClientB_PRIVATE_LIMITED` | `ClientB Pvt Ltd-MH` |
| `ClientB Private Limited` | `ClientB_Pvt_Ltd` | `ClientB_Pvt_Ltd` | `ClientB Pvt Ltd-MH` |
| `ClientB_Private_Limited` | `ClientB_Pvt_Ltd` | `ClientB_Pvt_Ltd` | `ClientB_Private_Limited-MH` |
| `Privateer Traders Unlimited` | `Privateer_Traders_Unlimited` | `Pvter_Traders_Unlimited` | `Privateer Traders Unlimited-DL` |
| `Delimited Exports Private Limited` | `Delimited_Exports_Pvt_Ltd` | `Delimited_Exports_Pvt_Ltd` | - |

Where the cleaned name is used, for `ClientB PRIVATE LIMITED` (Maharashtra):

| Use | A | B |
|---|---|---|
| Copy of the April GSTR-3B | `GSTR3B-ClientB_Pvt_Ltd-Maharashtra-Apr.xlsx` | `GSTR3B-ClientB_PRIVATE_LIMITED-Maharashtra-Apr.xlsx` |
| ITC report name | `ITC_Report_ClientB_Pvt_Ltd_MH_211026_1030.xlsx` | `ITC_Report_ClientB_PRIVATE_LIMITED_MH_211026_1030.xlsx` |
| Category folder with the include-name tick on | `GSTR3B (ClientB_Pvt_Ltd)` | `GSTR-3B Exports (ClientB_PRIVATE_LIMITED)` |
| ITC Links cell B8 (GSTR-2B file name) | `GSTR-2B-Reco-ClientB_Pvt_Ltd-Maharashtra-FY2025-26` | `GSTR-2B-Reco-ClientB_PRIVATE_LIMITED-Maharashtra-FY2025-26` |

Two spellings of one client in one state (`ClientB PRIVATE LIMITED` with an April GSTR-3B, `ClientB Private Limited` with April and May; different contents), Fresh mode:

| | A | B |
|---|---|---|
| Client folders | both clients in `ClientB Pvt Ltd-MH` | both clients in `ClientB Pvt Ltd-MH` |
| April copies | both cleaned to `GSTR3B-ClientB_Pvt_Ltd-Maharashtra-Apr.xlsx`: the second overwrites the first, which is kept as `…-Apr_backup_211026_1030.xlsx`; both reported "Success" | `GSTR3B-ClientB_PRIVATE_LIMITED-Maharashtra-Apr.xlsx` and `GSTR3B-ClientB_Pvt_Ltd-Maharashtra-Apr.xlsx` side by side |
| ITC report names | identical for both clients (`ITC_Report_ClientB_Pvt_Ltd_MH_211026_1030.xlsx`), so the second client's report replaces the first's (the A audit ran this: `v_abbrev_merge.py`) | different names, both kept |

**What it changes in the annual working.** The names of copied files, report files and bracketed folder suffixes for clients written in capitals (common for company names) or containing words like "Privateer". Within each project the Links file-name cells match that project's copy names for every file copied successfully in the run (a skipped or failed file keeps its original name, rule A-R48 / B-R57, the same in both), so the templates find the copied files in both (RAN, the B8 row above). In B the same client is "Pvt Ltd" in its client folder but "PRIVATE_LIMITED" in its file and report names, and such names are 8 characters longer each time they appear (the A commits cite Excel's ~218-character path limit ⚠). In A, two spellings of one client in one state merge into the same file and report names and overwrite each other; in B only the client folder is shared.

**When.** B keeps the original text replacement from `edac6a3` (shared; `47a1492:utils/helpers.py:78-79`). A changed it in `33d0baf` (23-07-2026, "Fix sanitize abbreviation, processing re-entrancy, folder-setting key"); the message says the old "case-sensitive substring replace mangled words that merely CONTAIN them ('Privateer'->'Pvter') and missed ALL-CAPS names". Because "_" counts as part of a word, `33d0baf` stopped abbreviating underscore-joined names; `4219ffb` (25-07-2026, "Fix sanitize_filename: abbreviate underscore-joined names again") restored that, citing "path length (Excel's ~218 char limit, and there is no full-path guard anywhere)". A's tests pin the new behaviour (A:tests/test_helpers.py:43-47, 69-71); three of them fail against B's code (`notes2/A_tests.md`). Not a GST rule.

### Category 3: State, GSTIN and PAN codes

The state part of each file name is turned into a two-letter code that goes into the client key, the client folder name and the report names.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R21 | B-R21 | State name → two-letter code table ⚠ [GST RULE]: 39 lower-case names, 36 codes (e.g. delhi DL, maharashtra MH, tamil nadu TN, odisha/orissa OD, puducherry/pondicherry PY); text trimmed and lower-cased first; letter codes only, no numeric state codes | A:utils/constants.py:251-292; look-up A:utils/helpers.py:592-606 | B:utils/constants.py:234-275; look-up B:utils/helpers.py:485-499 | same | RAN core_probe: 39 names and 36 codes in both; `Delhi`, `" delhi "`, `DL`, `Tamil Nadu`, `Orissa`, `National Capital Territory of Delhi`, `Ladakh` give DL, DL, DL, TN, OD, DL, LA in both. funcdiff: `get_state_code` IDENTICAL. Since `edac6a3` (shared); B changed no state value after the split. |
| A-R22 | B-R22 | Made-up code for a state not in the table ⚠ [GST RULE] (A's registry does not mark ⚠; B's does): two or more words → first letters of the first two words; one word → first three characters of the original text, upper-cased | A:utils/helpers.py:608-618 | B:utils/helpers.py:501-511 | same | RAN core_probe: `Unknown Place` → `UP` (the same code as Uttar Pradesh), `Tamilnadu` → `TAM`, `Daman and Diu` → `DA`, `27` → `27`, `" Xyz"` → `" XY"`, `''` → `''`, in both. |
| A-R23 | B-R21 (use sites) | Where the code and where the state text are used [GST RULE]: the CODE in the client key, client folder and report names; the state TEXT as written in the organisation text report, the Step 2 table and the summary workbook | A:core/file_parser.py:229-236; A:utils/helpers.py:628-631; A:gui/handlers/processing_handler.py:293; A:core/file_organizer.py:82, 330 | B:core/file_parser.py:261-268; B:utils/helpers.py:521; B:gui/controllers/processing_controller.py:113; B:core/file_organizer.py:78, 313 | same | READ; RAN step3_probe: report names `…_MH_…`, `…_TN_…`, summary "State" column `Maharashtra`, `Tamil Nadu` in both. |
| A-R24 | B-R21 (note) | No GSTIN or PAN is read, checked or derived [GST RULE] | none | none | same | RAN: a whole-word, any-case search for "gstin" or "pan" over every `.py` file of both copies (tests excluded) finds nothing. |

There are no differences in this category.

### Category 4: Financial year and tax period

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R25 | B-R23 | Period parts of the file names ⚠ [GST RULE]: the period (GSTR-2B Reco, Sales Reco), the 8-digit date (IMS Reco), the month (GSTR-3B), start and end month (Sales) and the year (Annual Report) are kept as plain text and never read again; no financial-year or tax-period check exists | stored at A:core/file_parser.py:97-106 | stored at B:core/file_parser.py:93-102 | same | RAN: a search for `metadata` finds only the two lines that create and fill it in each project (A:core/file_parser.py:74, 106; B:core/file_parser.py:74, 102). RAN core_probe: periods such as `Apr25-Mar26 (1)`, `Apr (1)`, `2025-26 (1)` are kept as text in both. The time stamps in folder names are clock times, not tax periods (Category 6). |

There are no differences in this category. In both projects, files of different years for one client are grouped together; which of several files of one type is linked is decided by alphabetical order (rule A-R57 / B-R56, Category 8).

### Category 5: Expected files and completeness

This category decides which return files every client is expected to have and how complete each client is shown to be on Step 2.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R26 | B-R24 | Six expected types for every client ⚠ [GST RULE]: GSTR-2B Reco, IMS Reco, GSTR-3B Export, Sales, Sales Reco, Annual Report | A:utils/constants.py:68-75; A:core/file_parser.py:252-260 | B:utils/constants.py:68-75; B:core/file_parser.py:280-288 | same | funcdiff: `_analyze_client_completeness` IDENTICAL; RAN core_probe: `0042-MH` missing GSTR-2B Reco, GSTR-3B Export, IMS Reco, Sales, Sales Reco in both. |
| A-R27 | B-R25 | Duplicates ⚠ [GST RULE]: GSTR-3B may have any number of files ("monthly filings - unlimited"); any other type with 2 or more files gets "Multiple <type> (<n> files)" | A:core/file_parser.py:262-271 | B:core/file_parser.py:290-299 | same | RAN core_probe: ClientD with two GSTR-2B Reco files → "Multiple GSTR-2B Reco (2 files)" in both. (The IMS "(1)" file of 1.1 makes a duplicate in A only.) |
| A-R28 | B-R26 | Completeness = types found ÷ 6 × 100, one decimal [GST RULE] | A:core/file_parser.py:273-275 | B:core/file_parser.py:301-303 | same | RAN core_probe: 100.0, 33.3, 16.7 in both. |
| A-R29 | B-R27 | Client status and how Step 2 shows it [GST RULE]: "Complete" / "Missing <n> files" (wins over duplicates) / "Has duplicates" | Status A:core/file_parser.py:277-286; display A:gui/handlers/client_handler.py:37-59 (row tint), 147-159 ("Complete Only") | Status B:core/file_parser.py:305-314; display B:gui/views/validation_view.py:387-410 (text colour), 441-447 ("Complete Only") | **different** (display only) | 5.1 below (RAN) |
| A-R30 | B-R28 | "Critical" types ⚠ [GST RULE] `['Annual Report', 'GSTR-3B Export']` in `validate_results`, which nothing calls | A:core/file_parser.py:338-371 (list :359) | B:core/file_parser.py:366-399 (list :387) | same (dead in both) | funcdiff: `validate_results` IDENTICAL; RAN: a search of both copies finds no caller. |
| A-R31 | B-R29 | Scan statistics and how they are shown | Statistics A:core/file_parser.py:308-336; Step 2 text box A:gui/main_window.py:294-328 | Statistics B:core/file_parser.py:336-364; Step 1 scan log, Step 2 summary and status bar B:gui/pyqt_main_window.py:247-312 | **different** (input and display) | 5.2 below (RAN) |

#### 5.1 How Step 2 shows a client's status (A-R29 / B-R27) · [GST RULE]

**What both do.** The status texts and the order of the tests are identical (funcdiff IDENTICAL). "Complete Only" ticks only clients whose status is "Complete".
**What A does.** Shows the status with an icon ("✅ Complete", "⚠️ Missing 5 files", "⚠️ Has duplicates"; A:gui/handlers/client_handler.py:39) and tints the whole row: light green `#E8F5E8` for Complete, light orange `#FFF4E5` for every other status (A:gui/handlers/client_handler.py:52-59; colours A:utils/constants.py:154-155). "Complete Only" tests "status contains Complete" (A:gui/handlers/client_handler.py:152).
**What B does.** Shows the status without the icon and colours the text of the row: green `#107C10` for Complete, orange `#FF8C00` when the status contains "Missing", red-orange `#D83B01` otherwise (B:gui/views/validation_view.py:387-398). "Complete Only" tests "status equals Complete" (B:gui/views/validation_view.py:446). A count label reads "3 clients (1 complete)" (B:gui/views/validation_view.py:407-410).

**Worked example (RAN, `gui_probe_A.py`, `gui_probe_B.py`; the same three clients).**

| Client | Status (both) | A on Step 2 | B on Step 2 |
|---|---|---|---|
| `ClientA` (all six types) | Complete | "✅ Complete", row tinted `#E8F5E8` | "Complete", text `#107C10` |
| `0042` (Annual Report only) | Missing 5 files | "⚠️ Missing 5 files", row tinted `#FFF4E5` | "Missing 5 files", text `#FF8C00` |
| `ClientE` (all six types, two GSTR-2B Reco) | Has duplicates | "⚠️ Has duplicates", row tinted `#FFF4E5` (same as a client with missing files) | "Has duplicates", text `#D83B01` |
| "Complete Only" | | ticks ClientA only | ticks ClientA only |

Neighbouring screen differences seen in the same runs (registered under category 16 as A-R138 / B-R134): A's rows start unticked and the "Missing" column lists the missing type names; B's rows start all ticked and "Missing" shows a count.

**What it changes.** Only how quickly staff can tell a client with duplicate files (for example two GSTR-2B reconciliations for the year) from a client with missing files. Which clients are complete, and which are ticked by "Complete Only", is the same.

**When.** Status logic since `edac6a3` (shared). A's tints are the shared Tk code (colours in `GUI_CONFIG`, the same values in both projects' `utils/constants.py`; B no longer reads them). B's colours came with `c258047` (07-02-2026) and `7e09e5c` (08-02-2026); no message explains the choice. Tagged [GST RULE] because the status rule is tagged; the difference is display only, and nothing here judges the rule.

#### 5.2 Scan statistics and where they are shown (A-R31 / B-R29)

**What both do.** The same statistics function (funcdiff `get_statistics` IDENTICAL).
**What differs.** The input: B's "variations" include the files rejected as too small, not Excel or unreadable (1.3 to 1.5), which A never sees, while B's "total files" counts only the files that passed those checks. The display: A writes one text box on Step 2 (A:gui/main_window.py:294-328); B writes a log on Step 1, an HTML summary on Step 2 and a status-bar line (B:gui/pyqt_main_window.py:250-269, 281-298, 308-312).

**Worked example (RAN, `display_probe.py`, the `src_discovery` folder of 1.2).**

| Shown to the user | A | B |
|---|---|---|
| Files found | "Total Excel Files Found: 6" | "Files found: 4" (status bar: "Scan complete: 1 clients, 4 files, 7.8 KB") |
| Parsed | "✅ Successfully Parsed: 3"; "📊 Parsing Success Rate: 50.0%" | "Parsed: 2 (50.0%)" |
| Unparsed | "❌ Unparsed/Variations: 3" | "Unparsed: 2" |
| Variations | "🔍 FILE VARIATIONS: 3 files" | "Variations: 5" (shown next to "Unparsed: 2") |
| Clients | "Total Clients Identified: 1"; "⚠️ Clients with Missing Files: 1"; "📈 Completeness Rate: 0.0%" | "Clients: 1"; "Incomplete: 1"; "Complete: 0 (0.0%)" |
| GSTR-3B files | "• GSTR-3B Export: 3 files" | "GSTR-3B Export: 2" |
| Size | "💾 TOTAL DATA SIZE: 11.7 KB" | "Total size: 7.8 KB" |

**What it changes.** Only the counts the staff read after a scan. Technical point (clarity of messages): B shows "Unparsed: 2" and "Variations: 5" for the same scan without saying why they differ; A labels every client that is not Complete (including duplicate-only clients) "Clients with Missing Files".

**When.** The statistics function is shared (`edac6a3`). A's text box is the shared Tk code. B's display came with `c258047` / `b0e7d0e` (09-02-2026, "feat: Enhance extraction process with logging and statistics tracking", which added the statistics to the scan result), and the extra variations with `7e09e5c` (08-02-2026). Not a GST rule.

### Category 6: Folder layout and naming

This category decides the folder tree Step 3 builds in the target folder: a run folder `Annual Statement-<DDMMYY HHMM>` (Level 1), one client folder `<client>-<state code>` per client (Level 2), a `Version-<DDMMYY HHMM>` folder for each run (Level 3) and three category folders inside it (Level 4); and which run and version Step 4 later works on.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R32 | B-R31 | Run folder name `Annual Statement-<DDMMYY HHMM>`; a second Fresh run in the same minute reuses it | A:utils/constants.py:82; A:core/file_organizer.py:163-196 (worked out once per run, A-R111) | B:utils/constants.py:82; B:core/file_organizer.py:146-179 (worked out per client, same result) | same | RAN core_probe and step3_probe: `Annual Statement-211026 1030` in both. |
| A-R33 | B-R30 | Time stamps: `%d%m%y %H%M` for folders, the same with `_` for report and summary names; two separate clock readings per run | A:utils/helpers.py:25-35; A:core/file_organizer.py:43; A:gui/handlers/processing_handler.py:186 | B:utils/helpers.py:23-33; B:core/file_organizer.py:43; B:gui/controllers/processing_controller.py:44 | same | funcdiff: `get_timestamp`, `get_safe_timestamp` IDENTICAL. These are clock times, not tax periods. |
| A-R34 | B-R30 | `get_date_only()` (`%d%m%y`) - never called | A:utils/helpers.py:29-31 | B:utils/helpers.py:27-29 | same (dead in both) | funcdiff IDENTICAL; RAN: a search of both copies finds no caller |
| A-R35 | B-R32 | Client folder name [GST RULE] (the state-code part): abbreviations (2.2), "-<state code>", cut to the maximum length; if fewer than 6 characters would remain for the client, the whole name is cut and the state code is lost | A:utils/helpers.py:620-644 | B:utils/helpers.py:513-537 | same | funcdiff: `create_client_state_key` IDENTICAL. RAN core_probe: `ClientB PRIVATE LIMITED`/Maharashtra/35 → `ClientB Pvt Ltd-MH`; `ClientA Private Limited`/Delhi/15 → `ClientA Pvt -DL`; maximum 8 → `ClientA ` (code lost) - in both. |
| A-R36 | B-R32, B-R138 | The maximum client-folder length on Step 1 (default 35; 0 or less becomes 35 inside the organiser, A:core/file_organizer.py:42 = B:core/file_organizer.py:42) | Spin box with arrows 15-100, always active; typed values not checked. A:gui/tabs/setup_tab.py:434-443 | Qt spin box held within 15-100; greyed out unless "Include client name" is ticked, but its value is used anyway. B:gui/views/setup_view.py:374-382; B:gui/pyqt_main_window.py:203 | **different** | 6.1 below (RAN) |
| A-R37 | B-R33 | `FOLDER_STRUCTURE['level2'] = '{client}-{state}'` - never used | A:utils/constants.py:83 | B:utils/constants.py:83 | same (dead in both) | RAN: a search of both copies finds no use outside the definition |
| A-R38 | B-R34 | A new `Version-<DDMMYY HHMM>` folder in every run and mode | A:utils/constants.py:84; A:core/file_organizer.py:102-108 | B:utils/constants.py:84; B:core/file_organizer.py:98-104 | same | funcdiff: `create_client_structure` IDENTICAL; RAN modes_probe. |
| A-R39 | B-R35 | Names of the three category folders | `GSTR3B ({client})`, `ITC ({client})`, `Sales ({client})`. A:utils/constants.py:85-95 | `GSTR-3B Exports ({client})`, `Other ITC related files ({client})`, `Sales related files ({client})`. B:utils/constants.py:85-89 | **different** | 6.2 below (RAN) |
| A-R40 | B-R35, B-R36 | Whether a client's name goes into its category folders: the global tick for all clients, else a per-client flag | Decision A:core/file_organizer.py:121-134; per-client flags set on Step 2, A:gui/handlers/client_handler.py:175-223 | Decision B:core/file_organizer.py:117-130 (identical); nothing in B ever sets a per-client flag (dictionary defined at B:models/app_state.py:83 and only passed on) | **different** | 6.3 below (RAN) |
| A-R41 | B-R37 | Which folder each file type goes into ⚠ [GST RULE] (A's registry marks ⚠, B's does not) | GSTR-3B Export → GSTR-3B folder; GSTR-2B Reco and IMS Reco → ITC folder; Sales and Sales Reco → Sales folder; Annual Report → the Version folder. A:core/file_organizer.py:302-315 | The same mapping. B:core/file_organizer.py:285-298 | **different** only in the folder names of 6.2 | 6.4 below (RAN) |
| A-R42 | B-R39 | Step 4: which run folder and which version folders are offered | Run folder: the `Annual Statement-*` entry with the newest date modified (when the system last recorded a change). In it: client folders not starting with `_`; per client only its newest `Version-*` folder by date modified. A:power_query_extractor/gui/extractor_window.py:611-668 | Run folder: the highest name (sorted Z to A). In it: every client folder; every version folder that holds an `ITC_Report_*.xlsx` or `Sales_Report_*.xlsx`, highest name first (Z to A). B:gui/pyqt_main_window.py:419-495 | **different** | 6.5 below (RAN) |
| A-R62 (A files it under category 8) | B-R40 | Organisation text report `_Organization_Report_<client>_<state>.txt` in the Version folder, with whole-run running totals | A:core/file_organizer.py:317-379 | B:core/file_organizer.py:300-362 | same (only the folder names written inside differ, 6.2) | funcdiff: `create_organization_report` IDENTICAL. |
| (no A rule; the same unused method at A:core/file_organizer.py:381-396) | B-R41 | `cleanup_on_error` - never called, so a client that fails midway leaves its folders | A:core/file_organizer.py:381-396 | B:core/file_organizer.py:364-390 | same (dead in both) | funcdiff: `cleanup_on_error` IDENTICAL; RAN: a search of both copies finds no caller. |
| A-R43 (A files it under category 7) | B-R38 | Re-Run and Resume reuse the `Annual Statement-*` folder with the newest date modified | A:core/file_organizer.py:174-196 (newest at :184) | B:core/file_organizer.py:158-179 (newest at :167) | same | RAN modes_probe: with `Annual Statement-311225 1000` (modified 31-12-2025) and `Annual Statement-211026 0335` (modified 21-10-2026), a Re-run used `Annual Statement-211026 0335` in both. Note: B's own Step 4 would show the other folder (6.5). |

#### 6.1 The maximum client-folder length box (A-R36 / B-R32, B-R138)

**What A does.** A Tk spin box whose arrows run from 15 to 100 (A:gui/tabs/setup_tab.py:440-442), always active. A value typed into it is not checked: an arrow click turns 0 or less into 35 (A:gui/tabs/setup_tab.py:434-438), but a typed value goes to Step 3 as it is (A:gui/handlers/processing_handler.py:194).
**What B does.** A Qt spin box with range 15-100 (B:gui/views/setup_view.py:374-377), which cannot hold a value outside the range. It is greyed out until "Include client name in brackets for Level 4 folders" is ticked (B:gui/views/setup_view.py:377, 382), but Start reads its value whether greyed or not (B:gui/pyqt_main_window.py:203).

**Worked example (RAN, `gui_probe_A.py`, `gui_probe_B.py`, `core_probe.py`).**

| Typed or stored value | A reads | B holds |
|---|---|---|
| `8` | 8 | 15 |
| `500` | 500 | 100 |
| `20.5` | 20 | (cannot be typed) |
| `abc` | reading fails (Tcl error) | ignored, keeps the previous value (20) |
| a saved 0 brought back at start-up | 35 (A:gui/main_window.py:386-387) | shown as 15 |
| box enabled at start | yes | no (greyed) |

Folder names with a maximum of 8 (only possible in A): `ClientA`/Delhi → `ClientA-`; `ClientA`/Maharashtra → `ClientA-` (the same folder for both states); `ClientB PRIVATE LIMITED` → `ClientB ` (ending in a space); `ClientD`/Tamil Nadu → `ClientD-`. With B's lowest value 15: `ClientA Private Limited`/Maharashtra → `ClientA Pvt -MH`. (RAN, the builder is identical in both, rule A-R35.)

**What it changes.** In A, a typed value of 8 or less removes the state code from client folder names, so one client's registrations in different states would share a client folder; a typed text such as `abc` makes the Step 3 run stop with "Fatal error" (READ: the value is read at A:gui/handlers/processing_handler.py:194, inside the run's outer protection only; the A audit RAN this and also found that every settings save then fails). In B the value always stays within 15-100. Technical point: B's box cannot produce the values that drop the state code; A's can. B's greyed-out look suggests the value is not used, while it is.

**When.** A's box (15-100 arrows, typed values unchecked) is the shared Tk code since `edac6a3` (`47a1492:gui/tabs/setup_tab.py:271-272`). B's PyQt5 box came in `c258047` (07-02-2026) with range 5-100 and the greyed-out state; `7e09e5c` (08-02-2026) set the range to 15-100. No message gives a reason. Not a GST rule (the state-code part of the folder name is rule A-R35, which is the same in both).

#### 6.2 Category folder names (A-R39 / B-R35)

**What A does.** `GSTR3B`, `ITC`, `Sales`, each followed by " (<cleaned client name>)" when the client's name is to be included (A:utils/constants.py:91-95).
**What B does.** `GSTR-3B Exports`, `Other ITC related files`, `Sales related files`, with the same optional suffix (B:utils/constants.py:85-89).
**Same in both.** All three are created even when empty; the suffix is removed by text replacement when not wanted (A:core/file_organizer.py:117-139 = B:core/file_organizer.py:113-135).

**Worked example (RAN, `core_probe.py`, Fresh run at 21-10-2026 10:30).**

| | A | B |
|---|---|---|
| ClientA's GSTR-2B Reco copy | `…\ClientA-DL\Version-211026 1030\ITC\GSTR-2B-Reco-ClientA-Delhi-Apr25-Mar26.xlsx` | `…\ClientA-DL\Version-211026 1030\Other ITC related files\GSTR-2B-Reco-ClientA-Delhi-Apr25-Mar26.xlsx` |
| ClientA's GSTR-3B copy | `…\Version-211026 1030\GSTR3B\GSTR3B-ClientA-Delhi-Apr.xlsx` | `…\Version-211026 1030\GSTR-3B Exports\GSTR3B-ClientA-Delhi-Apr.xlsx` |
| ClientA's Sales copy | `…\Version-211026 1030\Sales\Sales-ClientA-Delhi-Apr25-Mar26.xlsx` | `…\Version-211026 1030\Sales related files\Sales-ClientA-Delhi-Apr25-Mar26.xlsx` |
| With the include-name tick | `ITC (ClientA)`, `GSTR3B (ClientA)`, `Sales (ClientA)` | `Other ITC related files (ClientA)`, `GSTR-3B Exports (ClientA)`, `Sales related files (ClientA)` |
| Path length | A's paths are 20 (ITC), 14 (Sales) and 9 (GSTR-3B) characters shorter | |

**What it changes.** Every path below a Version folder, and therefore the folder text written into the template Links cells (8.3). Staff opening a client's working folder see different folder names; folders made by older runs keep their old names. Excel and Power Query fail on full paths above about 218 characters ⚠ (Excel behaviour, may change), so the shorter names leave more room for long client names and share paths.

**When.** B's names are the shared originals (`edac6a3`; `47a1492:utils/constants.py:86-88`). A shortened them in `7347383` (26-07-2026, "Shorten category folder names to fit Excel's path limit"). The message reports a real run of 3030 output files on the network share whose worst path was 270 characters, with 1094 files (36 %) over Excel's ~218 limit; it models the short names together with a planned shorter share root at a worst case of 215 characters with none over, against 235 characters and 27 files over for the shorter root alone, and warns that "the margin is thin" because copied file names still carry the full client name. The code comment that accompanies the change (A:utils/constants.py:85-90) calls the 235 / 27 and 215 / 0 figures "Measured on a real run" and leaves out that they assume the shorter share root. Not a GST rule in either registry; its effect on the Links cells is 8.3.

#### 6.3 Per-client choice to include the client's name (A-R40 / B-R35, B-R36)

**What both do in the organiser.** If the global tick is on, every client's category folders carry the name; otherwise a per-client flag looked up under "<client>-<state code>" decides, default off (A:core/file_organizer.py:121-134 = B:core/file_organizer.py:117-130).
**What A does on screen.** Step 2 has a "Name" column; clicking it toggles the client's flag (A:gui/handlers/client_handler.py:175-209), with a warning for names over 10 characters; the global tick sets all flags (A:gui/handlers/client_handler.py:211-223).
**What B does on screen.** Nothing sets a per-client flag: the dictionary is created empty (B:models/app_state.py:83) and only passed to the organiser (B:gui/pyqt_main_window.py:370; B:gui/controllers/processing_controller.py:27-32, 50). Step 1 still says "Names >10 chars will show warning. This overrides individual client settings if checked." (B:gui/views/setup_view.py:385).

**Worked example (RAN, `core_probe.py`, `gui_probe_A.py`).** With the flag `{'ClientA-DL': True}` handed to the organiser, both projects name only ClientA's folders `… (ClientA)` and leave the other clients' plain (identical result). On A's Step 2, clicking ClientA's "Name" cell produced exactly that flag; in B no screen action can produce it (search RAN: no line writes the dictionary).

**What it changes.** In B the client's name can only be added to the category folders of all clients or none; in A it can be added for single clients (for example only for the two clients whose folders staff open often).

**When.** The per-client control is shared Tk code, but at the split it stored flags under the state NAME while the organiser looked them up under the state CODE, so per-client choices were silently ignored in both lines of work (`47a1492:gui/handlers/client_handler.py:178, 214`). A fixed the key in `33d0baf` (23-07-2026; the message says "per-client folder-name toggles were silently ignored"). B dropped the control with the Tk screens (`c258047`, `a62d3cc`, 07-02-2026). Not a GST rule.

#### 6.4 Which folder each file type goes into (A-R41 / B-R37) · [GST RULE] · ⚠

**What both do.** The same fixed map (funcdiff `_get_destination_folder` IDENTICAL): GSTR-3B Export → GSTR-3B folder; GSTR-2B Reco and IMS Reco → ITC folder; Sales and Sales Reco → Sales folder; Annual Report → the Version folder itself; any other type → "No destination folder for type: <type>". Only the folder names differ (6.2).

**Worked example (RAN, `core_probe.py`).** Every one of ClientA's eight files landed in the same category in both projects; for example `GSTR3B-ClientA-Delhi-Apr.xlsx` → A `…\Version-211026 1030\GSTR3B\`, B `…\Version-211026 1030\GSTR-3B Exports\`; `AnnualReport-ClientA-Delhi-2025-26.xlsx` → `…\Version-211026 1030\` in both.

**What it changes.** Nothing about which return goes with which other return; only the folder names that the Links cells point to (8.3).

**When.** The map since `edac6a3` (shared); the folder names as 6.2 (`7347383` in A). Tagged because the registry of A tags it (it would need changing if the set of return types changes ⚠); nothing here judges it.

#### 6.5 Step 4: which run folder and which version folders are offered (A-R42 / B-R39)

**What A does.** In the folder shown on Step 4, every entry named `Annual Statement-*` is a candidate and the one with the newest date modified is used. Inside it, every sub-folder not starting with `_` is a client; its newest `Version-*` folder by date modified is used; a client is listed if it has any `Version-*` folder, with the ITC and Sales ticks set only if such a report is present (names starting `itc_report_` / `sales_report_` and ending `.xlsx`, any letter case). A:power_query_extractor/gui/extractor_window.py:20-42, 611-668.
**What B does.** It reads the target folder last saved in the settings (B:gui/pyqt_main_window.py:421), sorts the `Annual Statement-*` folders by name from Z to A and uses the first (B:gui/pyqt_main_window.py:432-448). Inside it, it lists every client folder and every version folder that holds a file matching `ITC_Report_*.xlsx` or `Sales_Report_*.xlsx`, sorted by name Z to A (B:gui/pyqt_main_window.py:453-495).

**Worked example (RAN, `step4_probe.py` and `step4_probe_newonly.py`; dates modified set on purpose).** Target folder:
- `Annual Statement-311225 1000` (run of 31-12-2025, last modified 31-12-2025), with ClientA's ITC and Sales reports;
- `Annual Statement-211026 0335` (run of 21-10-2026, last modified 02-11-2026) holding: `ClientA-DL` with `Version-211026 0335` and a re-run `Version-021126 1000` (both with reports); `ClientB Pvt Ltd-MH` whose only report was renamed by hand to `….XLSX`; `_ClientC-DL` with an ITC report; `ClientD-TN` whose Version folder holds no report.

| | A | B |
|---|---|---|
| Run folder used | `Annual Statement-211026 0335` (newest date modified) | `Annual Statement-311225 1000` (highest name: "31…" sorts after "21…"), so last December's run is offered |
| If only the October 2026 run exists: ClientA | one line: `Version-021126 1000` (newest by date modified), ITC and Sales ticked | two lines: `Version-211026 0335` first, then `Version-021126 1000` (name order, day first) |
| `_ClientC-DL` | left out | listed |
| `ClientD-TN` (Version folder without reports) | listed, both ticks off | left out |
| `ClientB Pvt Ltd-MH` (report named `.XLSX`) | listed with ITC ticked; the search for the original report then finds nothing on Linux ("No ITC report found"); on Windows it is found (HIGH, not run) | left out on Linux; on Windows listed (HIGH) |

**What it changes.** Which run's reports Step 4 refreshes and extracts, and therefore which run's figures go into the consolidated workbook. Day-first names do not sort by date, so B's choice of run folder is the run with the highest day number, not the latest run, whenever runs fall in different months or years; inside a run B offers every version (a re-run and its predecessor side by side, older-named first when the days differ). A's choice follows the last change on disk, so writing anything into an older run folder (for example opening a report there, which makes Excel write a lock file) can make that older folder "newest" (the A audit ran this on Linux; that Windows updates folder dates the same way is MEDIUM). Within B, Step 3's Re-Run picks the run folder by date modified (row A-R43 / B-R38) while Step 4 picks it by name, so the two steps of B can work on different runs. Technical point: both methods can pick an older run; neither shows the user which date it used.

**When.** At the split the Tk extractor chose the run folder by date modified and the version by name (`47a1492:power_query_extractor/gui/extractor_window.py:517, 529`). A changed the version choice to date modified in `c1583d6` (23-07-2026, "Fix extractor cache path and version-folder selection"; the message says names "are 'Version-DDMMYY HHMM' and day-first text does not sort chronologically across month boundaries") and replaced the report searches with the prefix test in `275d9a4` (23-07-2026, "Cache level-1 folder; single-listing scans in PQ extractor", for speed). B wrote its own Step 4 list in `b0e7d0e` (09-02-2026, "feat: Enhance extraction process with logging and statistics tracking"), choosing both by name; no reason is given. Not tagged as a GST rule in either registry; it decides which data set is used, not a tax rule.

### Category 7: Copying and versioning of files

This category decides how source files are copied into the tree, what happens when a copy already exists, and what the three processing modes do.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R43 | B-R44 | The three processing modes (fresh, rerun, resume) and how they are described on screen | Logic A:core/file_organizer.py:60-63, 163-196, 276-287. Screen A:gui/tabs/setup_tab.py:401-412 (radio buttons) and 103-110 (hover help) | Logic B:core/file_organizer.py:56-59, 146-179, 259-270. Screen B:gui/views/setup_view.py:343-359 (radio buttons), 510-512 (help card) | **different** (screen texts only; logic the same) | 7.1 below (RAN) |
| A-R44 | B-R45 | `PROCESSING_MODES` name, description and icon | Name and description shown on Step 1 (A:gui/tabs/setup_tab.py:404, 410); icon never read (A:utils/constants.py:178, 183, 188) | None of the three is read (B:utils/constants.py:157-173) | **different** | 7.1 below (RAN) |
| A-R45 | B-R43 | A destination file that already exists: Fresh → backup then overwrite; Resume → "Skipped", "File already exists"; Re-run → overwrite without backup | A:core/file_organizer.py:276-287 | B:core/file_organizer.py:259-270 | same | funcdiff: `_organize_single_file` IDENTICAL; RAN modes_probe (7.1). |
| A-R46 | B-R43 | Backup name `<stem>_backup_<DDMMYY_HHMM><ext>` in the same folder | A:utils/helpers.py:217-234 | B:utils/helpers.py:187-204 | same | funcdiff IDENTICAL; RAN modes_probe: `GSTR3B-ClientA-Delhi-Apr_backup_211026_1030.xlsx` in both. |
| A-R47 | B-R42 | Copy only (never move or delete the source), sizes compared, a partial copy left on error | A:utils/helpers.py:172-215 | B:utils/helpers.py:142-185 | same | funcdiff: `safe_copy_file` IDENTICAL. |
| A-R48 | B-R57 (B files it under category 8; [GST RULE] in B's registry) | After a successful copy the client's file record takes the cleaned copy name (used by the Links cells); skipped or failed files keep the original name | A:core/file_organizer.py:233-246, 270-274 | B:core/file_organizer.py:216-229, 253-257 | same | funcdiff: `organize_files` IDENTICAL. |
| A-R49 | B-R46 | `get_unique_path` - never called | A:utils/helpers.py:451-470 | B:utils/helpers.py:344-363 | same (dead in both) | funcdiff IDENTICAL; RAN: a search of both copies finds no caller |
| A-R50 | B-R47 | `calculate_file_hash` - never called | A:utils/helpers.py:529-547 | B:utils/helpers.py:422-440 | same (dead in both) | funcdiff IDENTICAL; RAN: a search of both copies finds no caller |

#### 7.1 How the processing modes are described on screen (A-R43, A-R44 / B-R44, B-R45)

**What both do.** The same logic (funcdiff: `_organize_single_file`, `organize_files`, `create_client_structure` IDENTICAL; the run-folder choice differs only in being worked out once per run in A). Fresh makes a new run folder; Re-Run and Resume reuse the run folder with the newest date modified; every mode makes a new Version folder named by the current minute, so Resume skips files only when it is started in the same minute as the earlier run. A mode text other than exactly `fresh`, `rerun` or `resume` is refused.

**Worked example (RAN, `modes_probe.py`).** ClientA's two files, four runs in the same minute:

| Run | A | B |
|---|---|---|
| 1 Fresh | both files "Success" | the same |
| 2 Resume | both "Skipped", "File already exists" | the same |
| 3 Re-run | both "Success" (overwritten, no backup) | the same |
| 4 Fresh | both "Success", with `…_backup_211026_1030.xlsx` copies beside them | the same (in `GSTR-3B Exports` instead of `GSTR3B`) |
| Mode text `Fresh` (capital F) | "Invalid processing mode: Fresh" | the same |

**What differs: the texts on Step 1 (RAN, `gui_probe_A.py`, `gui_probe_B.py`).**

| Mode | A (radio button + line below + hover help) | B (radio button) | B (help card on the same page) |
|---|---|---|---|
| Fresh | "🆕 Fresh Run" / "Create new folder structure (recommended for first time)" / "Create a brand new timestamped folder for this run. Use this normally - previous runs are left alone." | "Fresh - Create new folder structure (recommended for first time)" | "Fresh: New timestamped structure" |
| Re-Run | "🔄 Re-Run" / "Add new version folders to existing structure" / "Add a new version inside the most recent run's folder instead of starting a new one. Use when redoing work for the same batch." | "Re-Run - Add new version folders to existing structure" | "Re-Run: Overwrite existing" |
| Resume | "▶️ Resume" / "Continue previously interrupted processing" / "Continue a run that was interrupted, skipping files that are already in place." | "Resume - Continue previously interrupted processing" | "Resume: Skip already processed" |

**What it changes.** Only what staff expect when they pick a mode, for example when a run broke off half-way through the client list. Technical point (clarity of messages): B's help card says "Re-Run: Overwrite existing", while its own button and the code add a new Version folder; no text in either project says that Resume skips files only within the same minute.

**When.** A shows the constant's name and description since the shared Tk code (`47a1492:gui/tabs/setup_tab.py:212-221`, from `edac6a3`); A added the hover help in `322717f` (25-07-2026). B's first PyQt5 texts (`c258047`, 07-02-2026) were "Re-Run - Overwrite existing files" and "Resume - Skip already processed"; the present button and card texts came in `7e09e5c` (08-02-2026). Not a GST rule.

### Category 8: Report creation and template Links cells

This category decides how Step 3 makes each client's ITC report and Sales report: what must be chosen first, which template is copied, and what is written into the template's "Links" sheet (the sheet into which the program writes folder paths and file names for the template's Power Query queries to read; Power Query is Excel's data-import tool, whose queries live inside the workbook). It also covers the Step 3 summary workbook. The program never writes a tax figure into a report; the figures are produced later by the template's own queries.

| A rule | B rule | What it decides | A's version | B's version | Relation | Evidence |
|---|---|---|---|---|---|---|
| A-R51 | B-R48, B-R133 | What must be chosen before Dry Run and Start [GST RULE] (A's registry tags it; B's does not) | Dry Run and Start: both templates set and existing, and a target typed. A:gui/handlers/file_handler.py:132-151; called at A:gui/handlers/processing_handler.py:40 (Dry Run) and :78 (Start) | Start: the same three checks with other wording, B:gui/pyqt_main_window.py:345-361, and again per client, B:gui/controllers/processing_controller.py:210-218. Dry Run: no check, B:gui/pyqt_main_window.py:326-343 | **different** (Dry Run) | 8.1 below (RAN) |
| A-R52 | B-R48, B-R51 | Which template file types are offered or accepted; the report is always named `.xlsx` | Browse filter `*.xltx *.xlsx`, A:gui/handlers/file_handler.py:31, 42; Step 1 checklist accepts `.xlsx .xltx .xltm .xlsm`, A:gui/utils/setup_validation.py:21, 62-63; Start accepts any existing file | Browse filter `*.xlsx *.xlsm` and the label "(.xlsx or .xlsm)", B:gui/views/setup_view.py:221, 612, 619; Start accepts any existing file | **different** | 8.2 below (RAN) |
| A-R53 | B-R49 | `TEMPLATE_EXTENSIONS = {'.xlsx', '.xltx', '.xltm'}` - never used | A:utils/constants.py:197 | B:utils/constants.py:180 | same (dead in both) | RAN: a search of both copies finds no use outside the definition |
| A-R54 | B-R54 | ITC report Links cells ⚠ [GST RULE]: B2 GSTR-3B folder (always); B4/B5 Version folder and Annual Report name; B7/B8 ITC folder and GSTR-2B Reco name; B10/B11 ITC folder and IMS Reco name; B3, B6, B9 never written | Map A:utils/constants.py:102-114; values A:core/excel_handler.py:603-656 | Map B:utils/constants.py:96-108; values B:core/excel_handler.py:540-588 | **different** (only the folder text, through 6.2; and B11 through 1.1) | 8.3 below (RAN) |
| A-R55 | B-R55 | Sales report Links cells ⚠ [GST RULE]: B2/B3 Sales folder and Sales name; B5/B6 Version folder and Annual Report name; B8/B9 Sales folder and Sales Reco name; B4, B7 never written | Map A:utils/constants.py:115-125; values A:core/excel_handler.py:657-708 | Map B:utils/constants.py:109-119; values B:core/excel_handler.py:590-622 | **different** (only the folder text, through 6.2) | 8.3 below (RAN) |
| A-R56 | B-R56 | Form of the Links values [GST RULE]: full folder paths with `/` turned into `\`; file names without extension; written as text; an empty value is not written, so the template's own text stays | A:utils/helpers.py:418-422, 553-555; A:core/excel_handler.py:189-197 (Excel), 297-303 (openpyxl) | B:utils/helpers.py:311-315, 446-448; B:core/excel_handler.py:159-167 (Excel), 246-252 (openpyxl) | same | funcdiff: `clean_windows_path`, `extract_filename_without_extension` IDENTICAL; RAN core_probe: B3 keeps "PLACEHOLDER-B3" and the file names have no extension in both. |
| A-R57 | B-R56 | Which file of a type feeds a Links cell [GST RULE]: the first in sorted full-path order | Sort A:utils/helpers.py:361; first file A:core/excel_handler.py:629, 640, 651, 667, 683, 699 | Sort B:utils/helpers.py:262; first file B:core/excel_handler.py:561-562, 572-573, 583-584, 595-596, 606-607, 617-618 | same (the outcome for IMS "(1)" files differs through 1.1) | RAN core_probe: ClientD's two GSTR-2B files `…-FY2025-26 (1).xlsx` and `…-FY2025-26.xlsx` → B8 = `GSTR-2B-Reco-ClientD-Tamil_Nadu-FY2025-26_(1)` in both (a space sorts before a dot). |
| A-R58 | B-R53 | Finding the Links sheet ⚠ [GST RULE] (B's registry tags it; A's does not) | Excel path: `Links` in any letter case, else a sheet whose name contains "link", else a new empty `Links` sheet, and success either way, A:core/excel_handler.py:162-183. openpyxl path: exact name only; a Links cell inside a merged range is skipped silently, A:core/excel_handler.py:285-303 | Excel path identical, B:core/excel_handler.py:131-153. openpyxl path: exact name only; a Links cell inside a merged range makes the report fail, B:core/excel_handler.py:233-252 | same on the Excel path; **different** on the openpyxl path | 8.4 below (RAN) |
| A-R59 | B-R51, B-R61 | Report file names and how Step 4 recognises them [GST RULE] | `ITC_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx`, `Sales_Report_…`, A:gui/handlers/processing_handler.py:291-298, 355-356. Step 4 ticks by names starting `itc_report_` / `sales_report_` and ending `.xlsx` in any letter case (A:power_query_extractor/gui/extractor_window.py:20-42) but finds the original with the search `ITC_Report_*.xlsx` (A:power_query_extractor/core/report_processor.py:111-125) | The same names, B:gui/controllers/processing_controller.py:111-130, 220. Step 4 uses the search `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` for the list (B:gui/pyqt_main_window.py:461-462) and for the original (B:power_query_extractor/core/report_processor.py:72-89) | **different** (names only through 2.2; recognition) | 8.5 below (RAN) |
| A-R60 | B-R50, B-R60 | What one Step 3 run does per client [GST RULE] (A-R60 and B-R50 tagged): order folders → copies → ITC report → Sales report → text report; both reports tried for every client; a client is "Success" even if its reports fail | A:gui/handlers/processing_handler.py:182-453; template paths and per-client flags read from the screen during the run (:193, 301, 310, 332) | B:gui/controllers/processing_controller.py:42-181, 204-239; template paths and options fixed when Start is pressed (B:gui/pyqt_main_window.py:352, 363-371) | **different** (live reads; client order); success counting the same | 8.6 below (RAN) |
| A-R61 | B-R94, B-R95 (B files them under category 11) | Step 3 summary workbook `GST_Processing_Summary_<DDMMYY_HHMM>.xlsx` (sheets Summary, Client Status, File Mapping, Errors, Variations) | Variations sheet filled from the scan; a failure is shown as a log line. A:gui/handlers/processing_handler.py:209-226, 455-475; A:core/excel_handler.py:371-601 | Variations sheet always "No file variations found"; a failure shows nothing. B:gui/controllers/processing_controller.py:58-76, 183-193; B:core/excel_handler.py:308-538 | **different** | 8.7 below (RAN) |
| A-R62 | B-R40 | Organisation text report | see Category 6 | see Category 6 | same | funcdiff IDENTICAL |
| A-R63 | B-R58 | `get_relative_excel_path` - never called (Links values are always full paths) | A:utils/helpers.py:557-574 | B:utils/helpers.py:450-467 | same (dead in both) | funcdiff IDENTICAL; RAN: a search of both copies finds no caller |
| A-R48 | B-R57 | Copy names used in the Links cells | see Category 7 | see Category 7 | same | funcdiff IDENTICAL |
| A-R98 (A files it under category 12) | B-R52 | Where the report is assembled: a local temporary copy, or the template copied straight to the final report name | Excel path: works on a copy in `%TEMP%\gst_excel_temp` and copies the finished report to the target once, A:core/excel_handler.py:147-155, 215 (removal of the temporary copy :254-260); openpyxl path the same, A:core/excel_handler.py:273-279, 306-308, 315-320 | Copies the template to the final report name first and fills it there, B:core/excel_handler.py:122-129 (Excel), 227-231 (openpyxl) | **different** | 8.8 below (RAN on the openpyxl path) |
| A-R99 (A files it under category 12) | B-R59 | A report counts as "created" if the output file exists | A:core/excel_handler.py:217-224 (Excel), 310 (openpyxl) | B:core/excel_handler.py:182-189 (Excel), 258 (openpyxl) | same | READ; RAN step3_probe: a Sales report with no Links cell written counted as created in both. |

#### 8.1 What must be chosen before a Dry Run (A-R51 / B-R48, B-R133) · [GST RULE]

**What A does.** Dry Run (a preview that lists what Start would do, without copying anything) and Start use the same check: both template paths set ("Please select both ITC and Sales template files"), a target typed ("Please select a target folder for organized files"), and both template files present ("ITC template file not found" / "Sales template file not found"). A:gui/handlers/file_handler.py:132-151.
**What B does.** Start checks the same three things with its own words ("Please select a valid ITC template." / "Please select a valid Sales template." / "Please select a target folder."; B:gui/pyqt_main_window.py:352-361). Dry Run checks only that a client is ticked (B:gui/pyqt_main_window.py:326-330).

**Worked example (RAN, `gui_probe_A.py`, `gui_probe_B.py`).** No templates chosen, three clients ticked, Dry Run pressed:

| | A | B |
|---|---|---|
| Result | error box "Please select both ITC and Sales template files"; no preview | the preview is printed: for each client "📄 Would organize: …" for every file, then "📊 Would create: ITC Report" and "💰 Would create: Sales Report", ending "🧪 ════════ DRY RUN COMPLETE ════════" and "💡 No files were actually moved or created" |

**What it changes.** Only the preview: in B a Dry Run can promise ITC and Sales reports that Start would then refuse to make until templates are chosen. Start behaves alike in both: neither checks the template's type or contents (8.2).

**When.** The shared Dry Run called the check (`47a1492:gui/handlers/processing_handler.py:25`, from `edac6a3`). B's PyQt5 Dry Run came without it in `c258047` (07-02-2026); no reason is given. Tagged [GST RULE] by A's registry (it gates whether returns are processed); described only.

#### 8.2 Template file types (A-R52 / B-R48, B-R51)

**What A does.** The Browse dialog offers `*.xltx *.xlsx`; the Step 1 checklist marks other types "Not an Excel template", but this only colours the readiness pill; Start accepts any existing file.
**What B does.** The Browse dialog offers `*.xlsx *.xlsm`; Start accepts any existing file; there is no checklist.
**Same in both.** The report is always named `.xlsx` whatever the template's type (A:gui/handlers/processing_handler.py:298, 356; B:gui/controllers/processing_controller.py:220), and the template's contents (Links sheet, queries) are not checked.

**Worked example (RAN, inline runs in `run/C5a` on `tmp/tplnames`).**

| ITC template chosen | A checklist | A Start | B Start |
|---|---|---|---|
| `ITC_template.xlsx` | ok | accepted | accepted |
| `ITC_template.xltx` | ok | accepted | accepted |
| `ITC_template.xlsm` | ok | accepted | accepted |
| `ITC_template.xltm` | ok | accepted | (not run; accepted by the same test) |
| `ITC_template.xlsb` | "Not an Excel template" | accepted | accepted |
| `ITC_template.txt` | "Not an Excel template" | accepted | accepted |
| none | "Not selected yet" | refused | refused ("Please select a valid ITC template.") |

**What it changes.** Which template files staff can pick with the mouse: an `.xltx` template in A, an `.xlsm` in B (both can be typed in either). An `.xlsm` or `.xltx` template gives a report named `.xlsx` whose inside is a macro workbook or a template; whether Excel then opens it normally is not sure (MEDIUM; how to check: run one client on Windows with such a template and open the report). Not a GST rule.

**When.** The shared Browse filter was `*.xltx *.xlsx` (`47a1492:gui/handlers/file_handler.py:31, 42`), kept by A; A added the checklist in `322717f` (25-07-2026). B's filter `*.xlsx *.xlsm` came in `c258047` (07-02-2026) and its label in `7e09e5c` (08-02-2026). No reason is given.

#### 8.3 The folder text in the Links cells (A-R54, A-R55 / B-R54, B-R55) · [GST RULE] · ⚠

**What both do.** The same cells ⚠ (template cell positions) from the same sources (funcdiff: `prepare_template_data` differs only in removed debug lines). What differs is the folder text, because of the category folder names (6.2), and for IMS the file chosen (1.1).

**Worked example (RAN, `core_probe.py`, a placeholder template whose cells B2-B11 hold "PLACEHOLDER-B…", openpyxl path; Linux paths appear with backslashes as `\root\…`, shortened here to `…`).** ClientA, Delhi:

| ITC report cell | A | B |
|---|---|---|
| B2 (GSTR-3B folder) | `…\Annual Statement-211026 1030\ClientA-DL\Version-211026 1030\GSTR3B` | `…\Annual Statement-211026 1030\ClientA-DL\Version-211026 1030\GSTR-3B Exports` |
| B3 | `PLACEHOLDER-B3` (template text kept) | `PLACEHOLDER-B3` |
| B4 (Version folder) | `…\ClientA-DL\Version-211026 1030` | the same |
| B5 (Annual Report) | `AnnualReport-ClientA-Delhi-2025-26` | the same |
| B7 (GSTR-2B folder) | `…\Version-211026 1030\ITC` | `…\Version-211026 1030\Other ITC related files` |
| B8 (GSTR-2B file) | `GSTR-2B-Reco-ClientA-Delhi-Apr25-Mar26` | the same |
| B10 (IMS folder) | `…\Version-211026 1030\ITC` | `…\Version-211026 1030\Other ITC related files` |
| B11 (IMS file) | `ImsReco-ClientA-Delhi-31032026_(1)` (1.1) | `ImsReco-ClientA-Delhi-31032026` |

| Sales report cell | A | B |
|---|---|---|
| B2 and B8 (Sales folder) | `…\Version-211026 1030\Sales` | `…\Version-211026 1030\Sales related files` |
| B3 (Sales file) | `Sales-ClientA-Delhi-Apr25-Mar26` | the same |
| B5 / B6 | `…\Version-211026 1030` / `AnnualReport-ClientA-Delhi-2025-26` | the same |
| B9 (Sales Reco file) | `SalesReco-ClientA-Delhi-FY2025-26` | the same |

For `ClientB PRIVATE LIMITED` the file-name cells differ as well, through 2.2 (B8 `GSTR-2B-Reco-ClientB_Pvt_Ltd-Maharashtra-FY2025-26` in A, `…ClientB_PRIVATE_LIMITED…` in B). For a client with only GSTR-3B and GSTR-2B files (ClientD), only ITC B2 and B7/B8 are written and the Sales report gets no cell, in both.

**What it changes.** The folder each template query is pointed at. If the templates' queries take the folder from these cells, both projects' reports find their files. If any query contains a fixed folder name such as "Other ITC related files", it works only with B's names (and with A's runs made before `7347383`). The templates are not in either repository, so this is NOT SURE. How to check: open each template in Excel, Data > Get Data > Launch Power Query Editor, and for each query open Advanced Editor; search for "GSTR-3B Exports", "Other ITC related files" and "Sales related files", and see whether the source step refers to the Links cells.

**When.** Cells and sources since `edac6a3` (shared; B changed no Links value after the split). A's folder text changed with the folder names in `7347383` (26-07-2026, "Shorten category folder names to fit Excel's path limit"; reasons in 6.2); the commit changed only `utils/constants.py`, and the A audit (Section 5, A-R54) asks the author the same template question. Which version to keep is the author's decision.

#### 8.4 A Links cell inside a merged cell range, openpyxl path (A-R58 / B-R53) · [GST RULE] · ⚠

**What both do on the Excel (COM) path** - the path used on office PCs with pywin32 - identical code (A:core/excel_handler.py:162-183 = B:core/excel_handler.py:131-153; READ only, Excel cannot run here).
**What differs on the openpyxl path.** Writing into a cell that lies inside a merged range, but is not its first cell, raises an error in openpyxl. A catches every error for that cell and moves on (bare `except: pass`, A:core/excel_handler.py:302-303). B catches only `KeyError` and `ValueError` (B:core/excel_handler.py:251), so the error ends the whole report.

**Worked example (RAN, `merged_probe.py`).**

| Template Links sheet | A | B |
|---|---|---|
| B2 is the first cell of a merged B2:C2 | B2 written, all other cells written, report "created" | the same |
| B2 lies inside a merged A2:B2 | B2 left empty, all other cells written (B4, B5, B11 …), report "created", no message | report fails; log file "Fallback method error: 'MergedCell' object attribute 'value' is read-only"; the unfilled template copy stays on disk under the report name, with B4, B5, B11 still showing the template's placeholders |

**What it changes.** Nothing for reports made through Excel (identical code). The openpyxl path runs in A on any PC where pywin32 cannot be loaded (A:core/excel_handler.py:20-28; normally off Windows, since A's start-up check requires pywin32 on Windows, rule A-R124); in B the window cannot start without pywin32 (rule B-R124), so B's openpyxl path cannot be reached through the app at all (HIGH, READ). The openpyxl path also loses the template's Power Query (rule A-R96 / B-R97, category 12).

**When.** Shared bare `except` from `edac6a3`. B narrowed it in `c258047` (07-02-2026); the message lists "Fixed all 26 bare except: clauses with specific exception types". Tagged [GST RULE] and ⚠ by B's registry (the sheet name `Links` belongs to the templates); described only.

#### 8.5 Report names and how Step 4 recognises them (A-R59 / B-R51, B-R61) · [GST RULE]

**What both do.** Name the reports `ITC_Report_<cleaned client>_<state code>_<DDMMYY_HHMM>.xlsx` and `Sales_Report_…` in the client's Version folder; the cleaned client name differs through 2.2 (RAN step3_probe: `ITC_Report_ClientB_Pvt_Ltd_MH_211026_1030.xlsx` in A, `ITC_Report_ClientB_PRIVATE_LIMITED_MH_211026_1030.xlsx` in B; ClientA and ClientD identical). Refreshed copies (`…_Refreshed_…xlsx`) count as reports in both.
**What differs in Step 4.** A ticks a report when a file name, in lower case, starts `itc_report_` / `sales_report_` and ends `.xlsx`; it then looks for the original with the search `ITC_Report_*.xlsx`, which respects letter case on Linux. B uses that search for both the list and the original.

**Worked example (RAN, `step4_probe.py`, on Linux).** ClientB's only report renamed by hand to `ITC_Report_ClientB_Pvt_Ltd_MH_211026_0335.XLSX`:

| | A | B |
|---|---|---|
| Step 4 list | ClientB listed, ITC ticked | ClientB not listed |
| Processing that report | "No ITC report found" | (not offered) |

On Windows both find the file (HIGH, Python's search ignores letter case there; not run). The program itself always writes `.xlsx` in lower case, so only renamed files are affected.

**What it changes.** Only reports whose names were changed by hand; on the office's Windows PCs the two behave alike (HIGH). The name differences of 2.2 decide which file a staff member opens for `ClientB PRIVATE LIMITED`.

**When.** Names since `edac6a3` (shared). A's prefix test came in `275d9a4` (23-07-2026, "Cache level-1 folder; single-listing scans in PQ extractor"), for speed on the network share; its message calls the new helpers "proven equivalent to the original glob logic", which holds on Windows but not on Linux (RAN above). B's search came with its Step 4 page in `b0e7d0e` (09-02-2026). Described only.

#### 8.6 What one Step 3 run does per client (A-R60 / B-R50, B-R60) · [GST RULE]

**What both do.** For each ticked client: folders → copies → ITC report → Sales report → organisation text report; both reports are tried for every client whatever files it has; a client is counted "Success" unless an error escapes the client's steps, even if its reports failed; the look-up of the client's data sits outside the per-client protection (A:gui/handlers/processing_handler.py:232; B:gui/controllers/processing_controller.py:85).
**What differs (1): when the template paths are read.** A reads the ITC and Sales template paths and the per-client name flags from the Step 1/Step 2 screen while the run goes on (A:gui/handlers/processing_handler.py:193, 301, 310, 332). B collects all settings once when Start is pressed and hands that snapshot to the worker (B:gui/pyqt_main_window.py:352, 363-371; used at B:gui/controllers/processing_controller.py:119, 127).
**What differs (2): the order of clients.** A processes clients in the order of its Step 2 table, which is the order in which the scan first met each client in the alphabetically sorted file list (A:gui/handlers/client_handler.py:30, 165). B's table, and so its run, is sorted by client key (B:gui/views/validation_view.py:367). (Registered in neither registry; the A writers flagged A's order as "UNREGISTERED" in `A_cross_flags.md`.)

**Worked example (RAN, `step3_probe.py`, `step3broken`, inline order run).**

| | A | B |
|---|---|---|
| Three clients (ClientA, ClientB PRIVATE LIMITED, ClientD); after ClientA is completed the ITC template on Step 1 is changed from "v1" to "v2" | ClientA's ITC report from v1; ClientB's and ClientD's from **v2** (Links A1 marker "ITC TEMPLATE v2") | all three ITC reports from **v1** |
| ClientD has only one GSTR-3B file | its Sales report is still made and counted "created" (no Links cell written) | the same |
| The ITC template is a damaged file | client "Success"; summary: Successful 1, ITC Reports Created 0, Report Creation Errors 1 | the same counts |
| Clients `Zeta Traders` (first file `AnnualReport-Zeta Traders-…`) and `Alpha Traders` (first file `GSTR3B-Alpha Traders-…`), both ticked | Zeta Traders first, then Alpha Traders | Alpha Traders first, then Zeta Traders |

**What it changes.** If the author changes a template on Step 1 while a batch of clients is running (for example after spotting that the previous year's template was selected), A's batch silently ends up with reports from two templates, split at the client being processed at that moment; B's whole batch uses the template that was chosen when Start was pressed. The client order decides which client's files win when two clients share a folder (2.2) and how the whole-run running totals in each organisation text report read (rule A-R62).

**When.** The live reads are the shared Tk code (`47a1492:gui/handlers/processing_handler.py:238, 247, 269`, from `edac6a3`); A's table order is the shared Tk code. B's snapshot and sorted table came with `c258047` (07-02-2026); no reason is given. Described only.

#### 8.7 The Step 3 summary workbook (A-R61 / B-R94, B-R95)

**What both do.** The same workbook code (funcdiff: `create_summary_report` and its five sheet functions IDENTICAL; the column-width helper `_autofit_columns` differs only in the kind of error it catches): sheets Summary, Client Status, File Mapping, Errors, Variations; written when at least one client was processed (A) or succeeded (B), which mean the same here because A counts both together (A:gui/handlers/processing_handler.py:416-417).
**What differs.** A hands over the scan's variation list (A:gui/handlers/processing_handler.py:224); B hands over an empty list (B:gui/controllers/processing_controller.py:73). If writing the workbook fails, A shows "❌ Failed to create summary report" or "❌ Summary Report Error: …" in the Step 3 log (A:gui/handlers/processing_handler.py:468-475); B shows a line only when it succeeds (B:gui/controllers/processing_controller.py:189-193; the B audit ran the failure case and saw nothing on screen, `audit/B/S3a.md` Case D).

**Worked example (RAN, `step3_probe.py`; the same three-client folder with `Claims register ClientA.xlsx` and ClientA's IMS "(1)" file).**

| Sheet | A | B |
|---|---|---|
| Summary: Total Files / Reports Generated | 10 / 6 | 9 / 6 |
| Client Status: ClientA "Files" | 7 | 6 |
| Variations | `Claims register ClientA.xlsx` · "No matching pattern found" · `ImsReco-ClientName-State-DDMMYYYY.xlsx` · 2004 · "Rename file" | "No file variations found" |

**What it changes.** The summary workbook is the only place where A lists the unused source files; B has no such place (1.6). The counts differ only through 1.1. Not a GST rule.

**When.** As 1.6: the real list at the split (`47a1492:gui/handlers/processing_handler.py:161`), the empty list in B since `c258047` (07-02-2026).

#### 8.8 Where the report is assembled (A-R98 / B-R52)

**What A does.** Copies the template to a local temporary folder (`%TEMP%\gst_excel_temp`), fills and saves it there, and copies the finished report to the client's Version folder once; the temporary copy is removed afterwards (Excel path A:core/excel_handler.py:147-155, 215, 254-260; openpyxl path :273-279, 306-308, 315-320).
**What B does.** Copies the template straight to the final report name, waits 0.5 s on the Excel path, and fills the copy in place (B:core/excel_handler.py:122-129, 227-231). B also creates the temporary folder and never uses it (B:core/excel_handler.py:52-53).

**Worked example (RAN on the openpyxl path, `core_probe.py` and `step3broken`; the Excel path is READ only).** The ITC template is a damaged file:

| | A | B |
|---|---|---|
| `_create_report_openpyxl` returns | False ("Fallback method error: File is not a zip file" in the log file) | False (the same log line) |
| File left under the report name | none | `ITC_Report_ClientA_DL_211026_1030.xlsx`, which cannot be opened (not a valid workbook) |
| Temporary folder afterwards | empty | (not used) |
| Step 4 later | nothing to list for ITC | the file matches `ITC_Report_*.xlsx`, so B's Step 4 lists it as ClientA's ITC report (B:gui/pyqt_main_window.py:461) |

**What it changes.** After a failed report, B leaves a broken or unfilled "report" in the client's Version folder under the normal name, which Step 4 then offers for refreshing; A leaves nothing there. Technical point (file safety): A's way never leaves a half-made report under the final name; B's does. A's commit also gives speed on the network share as the main reason.

**When.** B keeps the shared way (`edac6a3`). A changed it in `12d505f` (23-07-2026, "Build Excel reports on local disk, copy to network once"); the message explains that Excel's many small reads and writes are slow over a network share, and notes "On failure the destination is no longer left with a half-written template copy (minor improvement)" and that the Excel path "could not be executed in CI; validate one report on a Windows machine before bulk processing". Not a GST rule.

### Categories 1 to 8 at a glance

| Category | Rule pairs (rows) | same | different | GST-rule pairs that differ |
|---|---|---|---|---|
| 1 Input-file recognition | 17 | 10 | 7 (1.1-1.7) | 1.1 IMS "(N)" ⚠; 1.6 hint shown ⚠; 1.7 example names ⚠ |
| 2 Client identity and name cleaning | 4 | 1 | 3 (2.1, 2.2 covering A-R19 and A-R20) | 2.1 client key handed back by Step 2 |
| 3 State, GSTIN and PAN codes | 4 | 4 | 0 | - |
| 4 Financial year and tax period | 1 | 1 | 0 | - |
| 5 Expected files and completeness | 6 | 4 | 2 (5.1, 5.2) | 5.1 status display |
| 6 Folder layout and naming | 14 | 9 | 5 (6.1-6.5) | 6.4 type → folder ⚠ (folder names only) |
| 7 Copying and versioning | 8 | 6 | 2 (7.1, covering A-R43 and A-R44) | - |
| 8 Report creation and Links cells | 14 (plus two pointer rows) | 5 | 9 (8.1-8.8, 8.3 covering A-R54 and A-R55) | 8.1 Dry Run check; 8.3 Links folder text ⚠; 8.4 merged Links cell ⚠; 8.5 report recognition; 8.6 live template reads and client order |
| **Total** | **68** | **40** | **28** | **12 pairs** (A-R2, A-R12, A-R16, A-R17, A-R29, A-R41, A-R51, A-R54, A-R55, A-R58, A-R59, A-R60 with their B partners) |

Every GST rule that is the same in both projects in these categories comes from the shared commits and was not changed by B after the split: the five other file-name patterns and the matching method (A-R1, A-R3 to A-R7), the state table, fallback codes and their uses (A-R21 to A-R23), the absence of GSTIN, PAN and period logic (A-R24, A-R25), the expected types, duplicate, completeness and "critical" rules (A-R26 to A-R28, A-R30), the client-folder name (A-R35), the copy rename (A-R48 / B-R57), and the form of the Links values and the "first file in sorted order" rule (A-R56, A-R57).
