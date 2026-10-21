# A-X-EXT: Project A, everything the code needs from outside the repository, and all shared state

Project A = goyal-sidhant/AnnualStatement_oldv1, snapshot HEAD bf2f1dd (27-07-2026), Tkinter GUI.
All paths below are repository-relative (for example `core/file_parser.py:123`). Line numbers are those of
bf2f1dd. Scratch work lives in `/root/merge-scratch/run/A-X-EXT/` (scripts `t1_...py` to `t14_...py`
and their saved outputs `t*_output.txt`; see the Appendix).

---

## 0. Method, legend and scope

- **READ** = I read the code (every line of the function named). **RAN** = I executed the code (a copy of
  it, in the scratch folder, on invented names and invented numbers only). **HIGH / MEDIUM / LOW** =
  my confidence in an inference.
- **[GST RULE]** = the code decides something the author's GST knowledge owns (which return/table/sheet/
  cell a figure belongs to, state codes, which documents a client needs, tax periods...). For those I only
  describe what the code does. I do not judge whether it is right. **TIME-VARYING** = a value that can
  change over time (table numbers, return types, state list, file-naming conventions of the portal).
- Files read in full for this task (every line): `main.py`, `launch_extractor.py`, `requirements.txt`,
  `pytest.ini`, `.gitignore`, `.claude/settings.local.json`, `README.md`, `utils/__init__.py`,
  `utils/constants.py`, `utils/helpers.py`, `utils/excel_com.py`, `utils/requirements_check.py`,
  `utils/dependency_installer.py`, `core/__init__.py`, `core/file_parser.py`, `core/file_organizer.py`,
  `core/excel_handler.py`, `gui/__init__.py`, `gui/main_window.py`, all of `gui/handlers/*`, `gui/tabs/*`,
  `gui/utils/*`, `gui/widgets/*`, all of `power_query_extractor/*` (including the 1062-line
  `gui/extractor_window.py`), `docs/superpowers/specs/2026-07-25-merge-extractor-design.md`,
  `tests/README.md`, `tests/conftest.py`, `tests/test_file_patterns.py`, `tests/test_scan.py`,
  `tests/test_extractor_standalone.py`, the first 80 lines of `tests/test_main_window_tabs.py`.
  Commits read in full (message + diff): c1583d6, 7347383, 5c8ce95, 12d505f (excel_handler part).
  Commit messages read: a603c02, 56aaad9, 4f84ba9, 275d9a4, 2aaa2b1, 47a1492 (stat).
  The repository's own audit `PROJECT_AUDIT_ANNUALSTATEMENT.md` (f7eb8f1, March 2026) was only searched,
  not read in full.

### Plain-English glossary (terms used below)

- **Template**: an Excel file the author prepares once; the app makes a copy of it for every client.
- **Power Query (PQ)**: Excel's built-in data-import tool. Its queries are stored inside the workbook file
  (in a hidden part called `customXml/item*.xml` containing the word `DataMashup`). A "refresh" re-runs
  the queries, re-reading the source files.
- **COM / pywin32**: COM is Windows' way for one program to drive another; "Excel through COM" means
  Python starts (or grabs) a real Excel and gives it commands. `pywin32` is the Python package that
  provides it (`win32com`, `pythoncom`, `win32api`, `win32gui`, `win32con`).
- **openpyxl**: a Python library that reads and writes `.xlsx` files without Excel. When it SAVES a
  workbook it silently drops the Power Query part (RAN below, also the author's Standard 12).
- **Regex (regular expression)**: a text pattern used to recognise file names.
- **Stem**: a file name without its extension (`Report.xlsx` -> `Report`).
- **Sanitise**: replace characters the code considers unsafe in a file name.
- **mtime**: a file's or folder's "last modified" time stored by the file system. A folder's mtime changes
  whenever something is added to, removed from or renamed directly inside it.
- **glob / scandir**: ways for a program to list the files in a folder.
- **JSON**: a plain-text format for settings (`{"key": "value"}`).
- **CWD (current working directory)**: the folder a program "stands in" when it starts. A file name with
  no folder part (for example `gst_organizer.log`) is created in the CWD, which depends on HOW the app
  was started, not on where the app is.
- **Home folder / %USERPROFILE%**: `C:\Users\<name>` on Windows. Python's `Path.home()` returns it.
- **%LOCALAPPDATA%**: `C:\Users\<name>\AppData\Local`, the per-user, per-PC folder the author's Standard
  13 requires for settings and logs.
- **%TEMP%**: the user's temporary folder (`...\AppData\Local\Temp`), returned by `tempfile.gettempdir()`.
- **UNC path**: a network path written `\\server\share\folder`. A **mapped drive** is the same share shown
  as a letter, e.g. `Z:\folder`.
- **pythonw.exe**: the Windows Python that opens no console window (Standard 13 launches apps with it).
  Under it `sys.stdout` and `sys.stdin` are `None`.
- **8.3 short path**: an old Windows alias such as `C:\PROGRA~1`; not always available (often switched
  off on data volumes and on network shares).
- **`\\?\` extended path**: a Windows prefix that lifts the 260-character limit for programs that support
  it.

---

## 1. Templates the code expects

### 1.1 How many, how the user chooses them, where the path is stored (READ)

| Item | Detail | Evidence |
|---|---|---|
| Number | Exactly two: an **ITC template** and a **Sales template**. | `gui/main_window.py:83-84`, `gui/handlers/file_handler.py:27-47` |
| How chosen | "Browse..." buttons on Step 1 open a file dialog filtered to `*.xltx *.xlsx` (label "Excel Template"). The path can also be typed into the entry box. | `gui/handlers/file_handler.py:29-32, 40-43`; `gui/tabs/setup_tab.py:361-364, 378-385` |
| Accepted extensions (Step 1 checklist) | `.xlsx`, `.xltx`, `.xltm`, `.xlsm` (case-insensitive); anything else shows "Not an Excel template". | `gui/utils/setup_validation.py:21, 56-64` |
| Other extension list (unused) | `TEMPLATE_EXTENSIONS = {'.xlsx', '.xltx', '.xltm'}` is defined but never used anywhere (grep). | `utils/constants.py:197` |
| Gate before processing | Both template paths must be non-empty and exist; target folder must be non-empty. Nothing checks the template's CONTENT (sheets, queries) before processing. | `gui/handlers/file_handler.py:132-150` |
| Where the path is stored | In the settings file `~/.gst_organizer_cache.json` under keys `itc_template` and `sales_template` (full path as text, exactly as the dialog returned it - on Windows Tk returns forward slashes, e.g. `Z:/Templates/ITC.xltx`). Section 4.1. | `gui/main_window.py:360-371, 377-378` |
| When saved | Only when a browse dialog returns a file (and on the other save triggers listed in 4.1). A path TYPED into the box is not saved unless some other action saves later. | `gui/handlers/file_handler.py:33-36, 44-47`; HIGH |

### 1.2 What happens to the template file (READ; RAN for the non-Excel path)

For every selected client, Step 3 creates two reports (`gui/handlers/processing_handler.py:292-414`):

1. **Names** (READ `processing_handler.py:292-298, 355-356`):
   `ITC_Report_<safe_client>_<state_code>_<DDMMYY_HHMM>.xlsx` and
   `Sales_Report_<safe_client>_<state_code>_<DDMMYY_HHMM>.xlsx`, in the client's version folder.
   `<safe_client>` = `sanitize_filename(client)` (full name, not shortened, `utils/helpers.py:68-126`);
   `<state_code>` = section 2.6; the time is the run's timestamp with the space replaced by `_`.
   The output is ALWAYS named `.xlsx`, whatever the template's extension.
2. **Staging on the local disk** (READ `core/excel_handler.py:53-54, 147-159, 212-215, 254-260`, commit
   12d505f): the template is copied with `shutil.copy2` to `%TEMP%\gst_excel_temp\<output file name>`,
   Excel (or openpyxl) edits that local copy, then the finished file is copied to the network target with
   `shutil.copy2`, then the local copy is deleted (best effort). The commit message says this was done to
   avoid many slow network round-trips while Excel has the file open.
3. **Two code paths** (READ `core/excel_handler.py:19-28, 90-102`):
   - **Windows with pywin32** (`USE_WIN32COM = True`): Excel through COM (`_create_report_win32com`,
     104-260). It first tries `GetActiveObject("Excel.Application")` (the user's own running Excel), else
     `Dispatch` (123-128). Details of the Excel risks are in the A-X-EXCEL notes; not repeated here.
   - **Anything else** (no pywin32, or not Windows): openpyxl (`_create_report_openpyxl`, 262-322).
     RAN t6: openpyxl removes the Power Query part. With a test template holding a fake
     `customXml/item1.xml` (DataMashup), the output had NO `customXml` part (`t6_output.txt`). HIGH.
     On Windows, `utils/requirements_check.py:73-89` makes pywin32 ESSENTIAL so the app refuses to start
     rather than fall back; off Windows it is only a warning, so a Mac/Linux run silently produces
     reports without Power Query.

### 1.3 What the code assumes is INSIDE each template

#### 1.3.1 A sheet called "Links" with fixed cells [GST RULE: which report each source file feeds] (READ `utils/constants.py:102-126`)

The code writes **text** into these cells of the template copy (`core/excel_handler.py:189-197` COM,
`297-303` openpyxl; values are `str(value)`):

| Template | Cell | What is written (exact source) |
|---|---|---|
| ITC | B2 | `gstr3b_folder`: full path of the client's `GSTR3B` category folder, Windows backslashes (`clean_windows_path`, `utils/helpers.py:418-422`). Written even when the client has no GSTR-3B file (the folder always exists). `core/excel_handler.py:622` |
| ITC | B4 | `annual_folder`: full path of the version folder (the Annual Report is copied there), only if the client has an Annual Report, else blank. `:625-633` |
| ITC | B5 | `annual_filename`: stem (no extension) of the FIRST Annual Report file of the client, sanitised name. `:628-630` |
| ITC | B7 | `gstr2b_folder`: path of the `ITC` category folder, only if a GSTR-2B Reco file exists. `:636-644` |
| ITC | B8 | `gstr2b_filename`: stem of the FIRST GSTR-2B Reco file. |
| ITC | B10 | `ims_folder`: path of the `ITC` category folder, only if an IMS Reco file exists. `:647-655` |
| ITC | B11 | `ims_filename`: stem of the FIRST IMS Reco file. |
| Sales | B2 | `sales_folder`: path of the `Sales` category folder, only if a Sales file exists. `:661-674` |
| Sales | B3 | `sales_filename`: stem of the FIRST Sales file. |
| Sales | B5 | `annual_folder` (version folder), only if an Annual Report exists. `:677-690` |
| Sales | B6 | `annual_filename`: stem of the FIRST Annual Report file. |
| Sales | B8 | `sales_reco_folder`: path of the `Sales` category folder, only if a Sales Reco file exists. `:693-706` |
| Sales | B9 | `sales_reco_filename`: stem of the FIRST Sales Reco file. |

Cells B3, B6, B9 of the ITC template and B4, B7 of the Sales template are not touched.
"FIRST" = first in sorted path order (section 2.6.3); when a client has two files of one type only one is
linked. Only the GSTR-3B link is a whole folder (so the template presumably reads every file in that
folder - MEDIUM, the code does not show the query).

**Important details (all READ, RAN where marked):**

- **The extension is never written**, only the stem (`extract_filename_without_extension`,
  `utils/helpers.py:553-555`). The template's own queries must add the extension. Input files may be
  `.xlsx` or `.xls` (section 2.2), so an `.xls` input would be linked by stem only - whether the template
  then finds it depends on the template (question for the author). MEDIUM.
- **Blank values are skipped, so the template's existing cell content stays.** `if value:` at
  `core/excel_handler.py:192` (COM) and `:300` (openpyxl). RAN t7: a template whose Links cells held
  `PLACEHOLDER-B4/B5/B10/B11`, run for a client without Annual Report and IMS files, kept those
  placeholders in the output (`t7_output.txt`). So whatever the author left in those cells of the template
  (for example the paths of the client used to build the template) is what the report will query for a
  client missing that file. HIGH.
- **Paths are absolute** and are whatever the user picked as target (mapped drive letter or UNC). Commit
  12d505f relies on this ("data sources are ABSOLUTE paths written into the Links sheet"). A report
  created on one PC through `Z:` refreshes on another PC only if that PC maps the same letter. MEDIUM.
- Values pass through `clean_windows_path` (backslashes) only for folders; file stems are written as is.

#### 1.3.2 How the "Links" sheet is found, and what happens if it is missing (READ; RAN for openpyxl)

| Path | Rule | What if no match | Evidence |
|---|---|---|---|
| COM | 1st: a sheet whose name equals `Links` ignoring case. 2nd: the first sheet whose name CONTAINS `link` (so `Hyperlinks` would be taken). | Adds a NEW sheet named `Links` and writes the cells into it; the report is reported as success. | `core/excel_handler.py:162-183` |
| openpyxl | Only a sheet named `Links` ignoring case. | Writes nothing, saves, returns True (success). RAN t6: `MyLinks` and `Data`-only templates both "succeeded" with no cell written. | `core/excel_handler.py:285-303`; `t6_output.txt` |

The Step 3 log then says "Success" in both cases; the only hint is the general warning printed at start-up
"Templates must have 'Links' sheet" (`gui/main_window.py:396`) and the message after a FAILED report
(`processing_handler.py:323, 397`). HIGH.

#### 1.3.3 Power Query queries (READ; nothing in the code)

The code never reads, adds, edits or checks any query. It assumes (comments only) that the template
contains Power Query queries that take their file/folder locations from the Links cells
(`core/excel_handler.py:3, 66-72, 143-146`; commit 12d505f). How the queries read the cells (named ranges?
`Excel.CurrentWorkbook()`?), and whether they filter the folders by file name, cannot be seen from the
repository (it holds no template). Question for the author. HIGH that the code is silent.

#### 1.3.4 What the extractor (Step 4) assumes is inside the REFRESHED report

1. **Sheet `Info`, cell `BB2`** (READ `power_query_extractor/core/report_processor.py:638-685`): after the
   refresh and BEFORE saving, the code looks for a sheet named `info` (ignoring case) and reads `BB2`.
   The refresh is called valid only if BB2 is not empty, not blank, not `0`/`0.0`, and not one of
   `#REF! #VALUE! #NAME? #NULL! #DIV/0! #N/A #NUM! #ERROR!`. The success message calls the value
   "Client Name", so BB2 is expected to be filled by a query with the client's name (comment line 654).
   If the sheet is missing the refresh is marked failed ("Info sheet not found") but the file is still
   saved and the values are still extracted (`:345-374`, `:187-188`).
2. **Sheets and cells to extract** [GST RULE] TIME-VARYING (table numbers, sheet names)
   (READ `power_query_extractor/config/cell_mappings.py:33-242`, unchanged since 47a1492, 30-10-2025, a
   commit shared with project B). Sheet names must match EXACTLY, case included
   (`report_processor.py:564-566`, `sheet.Name == sheet_name`):

   | Output sheet (consolidated report) | Report | Input sheet (in the report) | Input cell -> output column |
   |---|---|---|---|
   | GSTR-3B Difference | ITC | `Diff GSTR-3B` | J15 -> Diff IGST; K15 -> Diff CGST; L15 -> Diff SGST |
   | GSTR-1 Difference | Sales | `Diff - GSTR-1` | N1 -> Diff Taxable Value; O1 -> Diff IGST; P1 -> Diff CGST; Q1 -> Diff SGST |
   | Turnover GSTR-1 | Sales | `Diff - GSTR-1` | I1 -> Taxable Value |
   | Reco Diff | ITC | `Reco` | C37 -> IGST Diff; D37 -> CGST Diff; E37 -> SGST Diff |
   | Table 6J Match | ITC | `T6 of R9` | K29 -> Table 6J Match |
   | Table 8D Limits | ITC | `T8 of R9` | G6 -> Table 8D Limits |
   | IMS Reco | ITC | `IMS_Reco` | AA1 -> IGST; AD1 -> CGST; AG1 -> SGST |
   | Extra In Purchase | ITC | `ExtraInPurchase` | O1 -> IGST; P1 -> CGST; Q1 -> SGST |
   | Extra In GSTR2B | ITC | `ExtraInGSTR2B` | S1 -> IGST; T1 -> CGST; U1 -> SGST |

   So the ITC template must contain (after refresh) sheets `Info`, `Diff GSTR-3B`, `Reco`, `T6 of R9`,
   `T8 of R9`, `IMS_Reco`, `ExtraInPurchase`, `ExtraInGSTR2B`; the Sales template `Info` and
   `Diff - GSTR-1`. A missing sheet gives `None` for its fields (`report_processor.py:570-576`), shown as
   "Missing" in the consolidated report (`data_consolidator.py:152-154`, text from
   `EXTRACTION_OPTIONS['missing_data_text']`). Note: a WRONG-COLUMN bug in how these values are read back
   is described in 4.9.3 (RAN).
3. **Report file names**: the extractor only looks at `ITC_Report_*.xlsx` and `Sales_Report_*.xlsx` in
   the newest version folder (section 3.5).

### 1.4 Template file-format pitfalls (RAN / READ)

- **`.xltx` template, non-Excel path - RAN t6:** the output is named `.xlsx` but keeps the TEMPLATE
  content type (`spreadsheetml.template.main+xml`) because openpyxl keeps `wb.template` from the source.
  A `.xlsx` file whose inside says "template" is commonly refused by Excel ("file format or file extension
  is not valid"). HIGH that the mismatch exists (RAN); MEDIUM that Excel refuses it (cannot test here;
  how to check: open such an output in Excel on Windows).
- **`.xltx` template, COM path - READ only:** the `.xltx` is copied under an `.xlsx` name
  (`core/excel_handler.py:147, 154`) and opened with `Workbooks.Open` (`:349`). Whether Excel opens a
  template that has been renamed `.xlsx`, and in which format `wb.Save()` (`:203`) then writes it, I am
  not sure. LOW. How to check: on Windows run Step 3 once with an `.xltx` template and open the report.
- **`.xlsm`/`.xltm` (macro) template:** accepted by the Step 1 checklist (`setup_validation.py:21`) but
  renamed `.xlsx`. Excel refuses macro content under `.xlsx`; openpyxl (keep_vba not set) drops macros.
  Not tested. LOW-MEDIUM.
- **Read-only template file - READ:** `shutil.copy2` copies the read-only flag on Windows (Python's
  `copystat` uses `chmod`, which on Windows sets/clears only the read-only attribute). So a template marked
  read-only gives a read-only staging copy; COM `wb.Save()` (`:203`) would then fail, and the next run's
  `local_temp.unlink()` (`:149-152`) fails silently (`except: pass`) and the following `copy2` onto the
  read-only file fails. The final `copy2` (`:215`) would also make the finished report read-only, which
  the extractor's own `copy2` (`report_processor.py:164`) passes on to the refreshed copy that Excel must
  save. MEDIUM (mechanism HIGH; only matters if the author keeps templates read-only - question).
- **Staging name collision - READ:** the staging file is `%TEMP%\gst_excel_temp\<output name>`
  (`:147, 273`). Two copies of the app run by the same Windows user, for the same client in the same
  minute, would use the same staging file. LOW (rare).

---

## 2. Input files the code recognises

### 2.1 Which files are looked at at all (READ `utils/helpers.py:240-373`, `core/file_parser.py:136-177`; RAN t2)

- **Only the top level of the chosen source folder.** `os.scandir` is not recursive
  (`helpers.py:328`); files in subfolders are ignored without any message. RAN t2 (`sub/` file absent).
- **Extension filter 1 (listing):** name ends with `.xlsx`, `.xls` or `.xlsm`, ignoring case
  (`helpers.py:327, 330`). `.xltx`/`.xltm`/`.csv`/`.pdf` are never looked at and are NOT reported
  anywhere. RAN t2 (`.xltx` file absent from every list).
- **Extension filter 2 (metadata):** `.xlsx .xls .xlsm .xltx .xltm` (`helpers.py:299`) - redundant after
  filter 1.
- **Must be a regular file** (`S_ISREG`, `helpers.py:296`) - a folder named `x.xlsx` is skipped.
- **Size at least 1024 bytes** (`helpers.py:302-303`). Smaller files are dropped silently (this also
  drops Excel's `~$` lock files). `MAX_FILE_SIZE_MB = 100` (`utils/constants.py:198`) is defined but never
  enforced: there is no upper limit. RAN t2 (1000-byte file absent).
- **Signature check:** first bytes must be `PK\x03\x04` (zip-based: xlsx/xlsm/xltx) or `\xd0\xcf\x11\xe0`
  (old .xls) (`helpers.py:264-265, 275-287`). A file with another header is dropped silently, not
  reported as a variation. RAN t2 ("NOPE" header file absent). The read uses up to 16 threads when there
  are 4 or more candidates (`helpers.py:271-272, 351-357`, commit 4f84ba9: 11 minutes -> 41 s on the real
  share over a VPN).
- **Sorted by full path** (`helpers.py:361`). On Windows, Python sorts paths ignoring case; on Linux case
  matters. RAN (`t2b_sort_output.txt`). This order decides "which file wins" (2.6.3).
- Files that pass are stored in `scanned_files` keyed by file NAME (`core/file_parser.py:207`).

### 2.2 File-name patterns, exactly as coded [GST RULE: which report/category a file belongs to] TIME-VARYING (portal file-naming conventions, return types)

`utils/constants.py:13-62`, all compiled with `re.IGNORECASE`, tried in this dict order
(`core/file_parser.py:82-128`), first match wins (`break` at 128):

| # | Key | Regex (verbatim) | type | category | folder key | groups captured |
|---|---|---|---|---|---|---|
| 1 | `GSTR-2B-Reco` | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | `GSTR-2B Reco` | ITC | `itc` | client, state, period |
| 2 | `ImsReco` | `^ImsReco-([^-]+)-([^-]+)-(\d{8})(?:\s*\(\d+\))?\.xlsx?$` | `IMS Reco` | ITC | `itc` | client, state, date |
| 3 | `GSTR3B` | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | `GSTR-3B Export` | GSTR3B | `gstr3b` | client, state, month |
| 4 | `Sales` | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | `Sales` | Sales | `sales` | client, state, start_month, end_month |
| 5 | `SalesReco` | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | `Sales Reco` | Sales | `sales` | client, state, period |
| 6 | `AnnualReport` | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$` | `Annual Report` | Annual | `root` | client, state, year |

Display descriptions shown on the Setup tab (`constants.py:19, 27, 35, 43, 51, 59`, rendered by
`gui/tabs/setup_tab.py:477-503`): `GSTR-2B-Reco-ClientName-State-Period.xlsx`,
`ImsReco-ClientName-State-DDMMYYYY.xlsx`, `GSTR3B-ClientName-State-Month.xlsx`,
`Sales-ClientName-State-StartMonth-EndMonth.xlsx`, `SalesReco-ClientName-State-Period.xlsx`,
`AnnualReport-ClientName-State-Year.xlsx`.

Where each type is copied (`core/file_organizer.py:302-315`, not from the `folder` key above, which is
unused): GSTR-3B Export -> `GSTR3B` folder; GSTR-2B Reco and IMS Reco -> `ITC` folder; Sales and Sales
Reco -> `Sales` folder; Annual Report -> the version folder itself.

**Rules that follow from the regexes (READ; all RAN in t1, see `t1_output.txt`):**

- **Extension:** only `.xlsx` or `.xls` (`\.xlsx?$`), any case. `.xlsm` passes the scan (2.1) but no
  pattern matches, so it becomes a "variation". RAN.
- **Case:** the whole name is matched ignoring case (`gstr-2b-reco-...XLSX` matches), BUT the captured
  client and state keep their typed case, and the client key is case-sensitive (2.6), so `ClientA` and
  `clienta` become TWO clients. RAN t2.
- **Spaces** around the client/state are stripped (`file_parser.py:102-106`); spaces inside are kept.
- **Order / ties between patterns:** the six literal prefixes are mutually exclusive (`Sales-` cannot
  match a name starting `SalesReco-`), so at most one pattern can match; the order does not matter in
  practice. HIGH (READ + RAN).
- **"(N)" duplicate markers** (Windows adds ` (1)` when a downloaded file already exists):
  - IMS Reco: optional `(N)` after the 8 digits, with or without a space, ONE marker only
    (commit 5c8ce95, 23-07-2026). `...01012024 (1).xlsx` and `...01012024(2).xlsx` match;
    `...01012024 (1) (2).xlsx` does not. The date group stays the 8 digits. RAN.
  - GSTR-3B: an optional LEADING `(N)` with optional space (`(1) GSTR3B-...` and `(1)GSTR3B-...` match;
    present since the first commit edac6a3). A trailing ` (1)` also matches because the month group
    `[^-]+` swallows it (month = `Apr24 (1)`). RAN.
  - GSTR-2B Reco, Sales, Sales Reco, Annual Report: a trailing ` (N)` is swallowed by the last group
    (period / end_month / year becomes e.g. `Q1 (1)`). A LEADING `(N)` is not accepted for these
    (`(1) Sales-...` -> variation). RAN.
- **IMS date:** exactly 8 digits (`\d{8}`; described as DDMMYYYY but the code never checks it is a date).
  7 digits or `01-01-2024` -> variation. RAN.
- **Hyphens are separators**, so a client or state containing a hyphen breaks recognition (RAN t1):
  - GSTR-2B Reco, Sales Reco, Annual Report (last group `(.+)` takes the rest): the name is MISREAD
    without warning except a length warning - `GSTR-2B-Reco-Client-A-Delhi-Apr24.xlsx` gives client
    `Client`, state `A`, period `Delhi-Apr24`.
  - IMS Reco, GSTR-3B, Sales (fixed number of hyphen-free groups): not recognised -> variation.
- **Near-miss spellings are not recognised:** `GSTR-3B-...`, `GSTR2B-Reco-...`, `GSTR-2B Reco-...`,
  `IMS-Reco-...`, `Annual Report-...`, too few parts (`GSTR3B-ClientA-Delhi.xlsx`,
  `Sales-ClientA-Delhi-Apr.xlsx`). RAN.
- **`Sales-Reco-...` is MISREAD as a Sales file:** `Sales-Reco-ClientA-Delhi-Q1.xlsx` -> type `Sales`,
  client `Reco`, state `ClientA`, start `Delhi`, end `Q1`. RAN t1. HIGH.
- **Validation warnings never block:** client or state shorter than 2 characters, longer than 100, or
  containing `<>:"/\|?*` adds a warning to the parse result (`file_parser.py:108-119`, rules
  `utils/helpers.py:428-449`) but the file is still accepted and grouped. The state warning reuses the
  client wording ("State name issue: Client name too short"). The warnings are never shown in the GUI
  (grep: `warnings` is not read anywhere). RAN t1 (`GSTR3B-A-D-Apr24.xlsx` accepted).
- **Captured period/date/month/year are never used** (grep: `metadata` is only written,
  `file_parser.py:74, 106`). There is NO financial-year or tax-period logic anywhere in A: files of
  different periods/years for the same client are grouped together. [GST RULE] (absence of period logic).
  HIGH.

### 2.3 Hint text for unrecognised names (READ `core/file_parser.py:288-306`; RAN)

Checked in this order on the lower-cased name: contains `gstr` and `2b` -> GSTR-2B hint; else contains
`ims` -> IMS hint; else `gstr` and `3b` -> GSTR-3B hint; else `sales` and `reco` -> Sales Reco hint; else
`sales` -> Sales hint; else `annual` -> Annual hint; else "Check file naming convention". Only one hint is
ever produced (an `if/elif` chain), although the code joins hints with " or ".

### 2.4 What happens to an unrecognised file (READ `core/file_parser.py:209-220`; RAN t2)

- Valid Excel files (2.1) that match no pattern, OR match but give an empty client/state, go to the
  **variations** list with `filename, path, issue ('No matching pattern found'), size, pattern_expected`.
- They are never copied. The Validation tab shows only their COUNT (`gui/main_window.py:320-321`); the
  list itself appears only in the Step 3 summary workbook, sheet "Variations"
  (`core/excel_handler.py:554-585`, data passed at `processing_handler.py:224`). README.md:156 mentions a
  "variations tab" that does not exist.
- Files dropped by the size/signature/extension checks (2.1) do not appear anywhere.

### 2.5 State names to state codes [GST RULE] state-code logic, TIME-VARYING (states/UTs change)

READ `utils/helpers.py:592-618`, `utils/constants.py:251-292`; RAN t1.

- The state text from the file name is trimmed and lower-cased and looked up in `STATE_CODE_MAPPING`
  (39 entries: full state/UT names, plus `orissa`, `pondicherry`, `national capital territory of delhi`
  as alternatives). The codes are two-letter LETTER codes (for example `delhi -> DL`,
  `telangana -> TS`, `dadra and nagar haveli and daman and diu -> DD`), not the numeric GST state codes.
- Not found: if the text has 2 or more words, the code becomes the first letters of the first two words,
  upper-cased (`Unknown Place -> UP`, `New Delhi -> ND`, `Dadra and Nagar Haveli -> DA`,
  `Daman and Diu -> DA`); if one word, its first 3 characters upper-cased (`TamilNadu -> TAM`,
  `Mah -> MAH`, `MH -> MH`, `TG -> TG`). A warning is logged. RAN.
- Consequence (technical): two different state texts can give the same code and so the same client key
  (RAN t2: `Unknown Place` and `Uttar Pradesh` merged under `ClientB-UP`); the same state written two ways
  can give different codes (`TamilNadu` vs `Tamil Nadu`).

### 2.6 How recognised files become "clients" (READ `core/file_parser.py:226-286`; RAN t2)

#### 2.6.1 Client key
`"<client as typed>-<state code>"` (`file_parser.py:229-230`). Same key = same client. `client` and
`state` stored on the client are those of the LAST file processed for that key (`:234-235`), so the state
shown can be the code (`DL`) or the name (`Delhi`) depending on file order. RAN t2.

#### 2.6.2 Completeness [GST RULE] (which documents a client must have) TIME-VARYING
- Expected types (`utils/constants.py:68-75`): `GSTR-2B Reco, IMS Reco, GSTR-3B Export, Sales,
  Sales Reco, Annual Report`. Missing = expected minus found (`file_parser.py:255-260`).
- More than one file of a type -> "Multiple <type> (n files)" EXCEPT `GSTR-3B Export`, where any number
  is allowed (comment "monthly filings - unlimited", `:264-269`).
- Status: `Complete` (nothing missing, no duplicates), `Missing n files`, or `Has duplicates`
  (`:277-286`). Completeness % = types found / 6 (`:274-275`). RAN t2: a client with all six types and
  two GSTR-2B files shows `Has duplicates`, 100%.
- `validate_results()` (`:338-371`) would flag missing "critical" types `Annual Report` and
  `GSTR-3B Export`, but it is never called (grep) - dead code.

#### 2.6.3 Which file wins when a client has two files of one type (RAN t2 + `t2b_sort_output.txt`)
Report links use `files[0]` (`core/excel_handler.py:625-706`), i.e. the first in the scan's sorted order.
Because a space sorts before a dot, `...Apr24 (1).xlsx` sorts BEFORE `...Apr24.xlsx` on both Windows and
Linux, so the Windows "duplicate download" copy is the one linked. `(1) GSTR3B-...` also sorts first (but
all GSTR-3B files are copied and the folder is linked). HIGH (RAN on both sort rules).

#### 2.6.4 Client key consistency across the GUI (READ; RAN t4)
The same key `"<client>-<state code>"` is rebuilt in `gui/handlers/client_handler.py:169-170, 181, 218`
from the Validation tree's displayed values and used to look up `client_data` and
`client_folder_settings` (`core/file_organizer.py:122, 130`). It agrees with the parser's key EXCEPT
when Tk changes the displayed value: RAN t4 showed a ttk Treeview hands back `'007'` as the integer `7`,
so a purely numeric client name with leading zeros would give key `7-DL`, not found in `client_data`
(KeyError: Dry Run fails; in Step 3 `processing_handler.py:232` is outside the per-client `try`, so the
whole run stops with "Fatal error"). LOW likelihood, HIGH mechanism.

### 2.7 RAN summary (t1, t2)

See `t1_output.txt` (44 invented names) and `t2_output.txt` (an invented folder). Every claim in 2.1-2.6
marked RAN is visible there.

---

## 3. Folder layout

### 3.1 What Step 1 READS (source folder) (READ; RAN t2)

One flat folder holding all clients' files together (no client subfolders, no period/FY folders - they
would be ignored). No other input folder is read by Steps 1-3.

### 3.2 What Step 3 CREATES (READ `core/file_organizer.py`, `utils/constants.py:81-96`,
`gui/handlers/processing_handler.py`; RAN t3)

```
<Target folder>                                   (created if missing: file_organizer.py:65-69)
└── Annual Statement-DDMMYY HHMM                  Level 1 (constants.py:82; one per run in 'fresh')
    ├── GST_Processing_Summary_DDMMYY_HHMM.xlsx   (processing_handler.py:462-469; openpyxl, 5 sheets)
    ├── PQ_Extraction_Report_YYYYMMDD_HHMM.xlsx   (written later by Step 4; data_consolidator.py:175-179)
    └── <Client key>                              Level 2 (create_client_state_key, see below)
        └── Version-DDMMYY HHMM                   Level 3 (constants.py:84; new each run)
            ├── GSTR3B                            Level 4 categories (constants.py:91-95)
            │   └── <sanitised GSTR-3B file names>
            ├── ITC
            │   └── <sanitised GSTR-2B Reco and IMS Reco file names>
            ├── Sales
            │   └── <sanitised Sales and Sales Reco file names>
            ├── <sanitised Annual Report file name>
            ├── ITC_Report_<safe client>_<state code>_DDMMYY_HHMM.xlsx
            ├── Sales_Report_<safe client>_<state code>_DDMMYY_HHMM.xlsx
            ├── ITC_Report_..._Refreshed_DDMMYY_HHMM.xlsx     (Step 4, section 3.6)
            └── _Organization_Report_<client>_<state>.txt   (file_organizer.py:330-372, UTF-8)
```

Exact name rules:

- **Timestamps:** `DDMMYY HHMM` (`utils/helpers.py:25-27`) for Level 1 and Version folders, taken ONCE
  when the FileOrganizer is created (`file_organizer.py:43`), minute resolution. Reports and summary use
  the same timestamp with `_` (`processing_handler.py:186, 294, 462`). Backups use `DDMMYY_HHMM`
  (`helpers.py:33-35, 224-225`). The Step 4 report uses `YYYYMMDD_HHMM` (a different format). These are
  run times, not tax periods.
- **Level 2 = `create_client_state_key(client, state, max_length)`** (`utils/helpers.py:620-644`):
  `Private` -> `Pvt` and `Limited` -> `Ltd` (whole words, any case), then `"<client>-<state code>"`. If
  longer than `max_length` (setting "Max client folder name length", default 35, spinbox 15-100,
  `gui/tabs/setup_tab.py:440-442`; 0 or less becomes 35 at `file_organizer.py:42`), the client part is cut
  to `max_length - len(code) - 1` characters; if that would leave 5 or fewer characters the whole key is
  cut to `max_length` (and the state code can be lost: RAN t3 `('ABCDEFGHIJ','Delhi',8) -> 'ABCDEFGH'`).
  **No character sanitising** at Level 2: RAN t3 kept `#`, `[`, `]` in the folder name while the file
  names had them replaced. Whether `[`/`]` in a folder name upsets Excel or Power Query I am not sure
  (LOW; check by refreshing one report for such a client).
- **Level 4 category names** (commit 7347383, 26-07-2026): `GSTR3B`, `ITC`, `Sales` (were
  `GSTR-3B Exports`, `Other ITC related files`, `Sales related files`). Why: the commit measured a real
  run of 3030 output files on the network share - worst path 270 characters, 1094 files (36%) over
  Excel's ~218-character limit; the short names plus a planned shorter share root (an `FY 2025-26` folder
  on the office server, as written in the commit message) gave a modelled worst case of 215 characters
  with none over 218. The commit itself warns the margin is thin and that the copied file name still
  carries the full client name. Existing output folders keep their old names. READ HIGH.
- **"Include client name" option** (global checkbox `setup_tab.py:417-423`, or per-client "Name" column
  on the Validation tab `client_handler.py:175-209`): adds ` (<sanitised client>)` to each Level 4 name,
  e.g. `ITC (Client_A_Pvt_Ltd)` (`file_organizer.py:117-139`). The global setting overrides the per-client
  ones (`:124-130`). Default off. RAN t3.
- **Copied input file names** = `sanitize_filename(original name)` (`file_organizer.py:271`,
  `helpers.py:68-126`): runs of spaces/underscores -> one space; `Private`->`Pvt`, `Limited`->`Ltd` (whole
  words); control characters and each of `<>:"/\|?*[]{}+=!@#$%^,;'"`~` -> `_`; runs of spaces/underscores
  -> `_`; leading/trailing `_ . space` removed; empty -> `unnamed`; stem cut to
  `200 - len(extension) - 10` characters. Hyphens and `&()` are kept. RAN t3, e.g.
  `ImsReco-Client'C-Delhi-01012024 (1).xlsx -> ImsReco-Client_C-Delhi-01012024_(1).xlsx`.
- **`_Organization_Report_<client>_<state>.txt`** uses the RAW client and state text (not sanitised, not
  abbreviated), so it is often the longest path in the tree (RAN t3: 201 characters versus 196 for the
  longest data file). It is plain text, not opened by Excel.
- **Backups**: when a destination file already exists in `fresh` mode, the old file is first copied to
  `<stem>_backup_DDMMYY_HHMM<ext>` IN THE SAME FOLDER (`file_organizer.py:276-282`,
  `helpers.py:217-234`). See 3.3.

### 3.3 The three processing modes (READ `file_organizer.py:150-196, 276-287`; RAN t12, t13)

| Mode | Level 1 used | Version folder | Existing destination file |
|---|---|---|---|
| fresh | new `Annual Statement-<now>` | new `Version-<now>` | backed up next to it, then overwritten |
| rerun | the existing `Annual Statement-*` folder with the newest mtime (new one if none) | new `Version-<now>` | overwritten, no backup |
| resume | same as rerun | new `Version-<now>` | skipped ("File already exists") |

RAN t12 findings:
- **"Resume" does not resume.** The version folder name is the CURRENT minute, so a resume one minute
  later creates a brand-new version folder and copies everything again; it only skips files when started
  in the same minute as the interrupted run. HIGH.
- **Same-minute resume links a non-existent file name:** skipped files keep their ORIGINAL name in
  memory (the rename to the sanitised name happens only on success, `file_organizer.py:233-246`), so the
  Links cell gets e.g. `GSTR-2B-Reco-Client A Private Limited-Delhi-Apr24` while the file in the folder
  is `GSTR-2B-Reco-Client_A_Pvt_Ltd-Delhi-Apr24.xlsx`. RAN t12. HIGH (writer/reader mismatch).
- **"Newest" Level 1 is by mtime**, so adding anything to an older run's folder makes it the "newest".
  RAN t14 (same rule in the extractor).

RAN t13: two `fresh` runs in the same minute reuse the same folders and leave `_backup_` copies INSIDE
the `GSTR3B` category folder, which is linked as a whole folder (B2). If the template reads every file in
that folder, the backup copies would be read too. MEDIUM (depends on the template's query).

### 3.4 Two clients sharing one folder (RAN t5)

With the default 35-character limit, two clients in the same state whose names share the first 32
characters get the SAME Level 2 folder and the same version folder (same run timestamp). RAN t5: both
clients' GSTR-3B files ended up in one `GSTR3B` folder, and both ITC reports' B2 point to it. The reports
themselves get different names (full client name), but Step 4 later sees ONE client folder and processes
only `report_files[0]` per type (section 3.5). HIGH (mechanism); likelihood depends on client names.

### 3.5 What Step 4 (Power Query Extractor) READS (READ `power_query_extractor/gui/extractor_window.py:611-668`; RAN t14)

1. The folder in its "Target Folder" box (pre-filled from Step 1's target folder, section 4.1).
2. `glob("Annual Statement-*")` in that folder (files or folders); the one with the newest mtime is used
   (`:623-632`). Pointing the box at an `Annual Statement-...` folder itself finds nothing ("No Annual
   Statement folders found") - RAN t14. The design spec (`docs/superpowers/specs/...:86-89`) says pointing
   at an older run's folder "still works"; the code does not do that. HIGH.
3. Every subfolder whose name does not start with `_` is a client; its `Version-*` folder with the newest
   mtime is used (`:638-648`, commit c1583d6 switched from name order to mtime because `DDMMYY` text does
   not sort by date). Clients with no `Version-*` folder are skipped silently.
4. In that version folder: any file whose lower-cased name starts `itc_report_` / `sales_report_` and ends
   `.xlsx` sets "has ITC"/"has Sales" (`:20-42`); refreshed copies count too.
5. Processing a report picks `glob("<ITC|Sales>_Report_*.xlsx")[0]` (`report_processor.py:112-125`).
   That glob ALSO matches earlier refreshed copies, and the order is whatever the file system lists. RAN
   t9 on Linux: the refreshed copy came first and was treated as the "original", producing
   `..._Refreshed_020126_0900_Refreshed_211026_0359.xlsx`. On Windows NTFS the listing is alphabetical, and
   `..._1000.xlsx` sorts before `..._1000_Refreshed_...` (dot before underscore), so the true original
   is normally first there; on a non-Windows file server it may not be. MEDIUM.

### 3.6 What Step 4 CREATES (READ; RAN t9, t10)

- **Refreshed copy** next to the original: `<original stem><suffix>.xlsx`, where `<suffix>` is the
  "Refreshed File Suffix" box (default `_Refreshed_{timestamp}`, `{timestamp}` replaced by `DDMMYY_HHMM`)
  (`report_processor.py:128-130, 164`). The refreshed file is ~22 characters longer than the original
  (matters for the 218 limit). A blank suffix makes the copy's name equal the original's: RAN t9 gave the
  error "...are the same file" and the report was not processed.
- **`PQ_Extraction_Report_YYYYMMDD_HHMM.xlsx`** in the Level 1 folder found at scan time
  (`extractor_window.py:945`, `data_consolidator.py:175-179`), written directly with openpyxl. Two runs in
  the same minute overwrite it.

### 3.7 Path-length arithmetic (READ; RAN t3 lengths)

Full path of a copied ITC file with defaults =
`len(target) + 1 + 28 ("Annual Statement-DDMMYY HHMM") + 1 + len(Level 2, at most 35) + 1 + 19
("Version-DDMMYY HHMM") + 1 + 3 ("ITC") + 1 + len(sanitised file name)`
= `len(target) + 90 + len(file name)` at most. The client name appears twice (shortened in Level 2, full
in the file name) and a third time in Level 4 when "include client name" is on. No code checks the length
before creating anything; the 218 figure only appears in warnings/fallbacks
(`core/excel_handler.py:356`, `report_processor.py:720, 795`, `gui/tabs/setup_tab.py:112-123`). HIGH.

### 3.8 Documentation that disagrees with the code (READ)

- `README.md:179-191` shows the OLD long category names and report names `..._DDMMYYYY_HHMMSS.xlsx`;
  the code writes short names and `DDMMYY_HHMM` (7347383, `processing_handler.py:294-298`).
- `README.md:104-108` lists `openpyxl>=3.0.0`; `requirements.txt` says `openpyxl>=3.1.2`.
- `docs/superpowers/specs/2026-07-25-merge-extractor-design.md:80-81` says organise-only mode ends with a
  "Continue to extraction" button; the code only logs a hint (`processing_handler.py:145-149`).

---

## 4. The app's own files: settings, cache, logs, temp, run history

### 4.1 Inventory

| # | File / folder | Exact path expression (relative to what) | Written by | Read by | Format |
|---|---|---|---|---|---|
| 1 | Settings ("cache") | `Path.home() / '.gst_organizer_cache.json'` = `%USERPROFILE%\.gst_organizer_cache.json` | `gui/utils/cache_manager.py:15, 29-48` via `gui/main_window.py:358-371` | `cache_manager.py:18-27` -> `main_window.py:120-121, 373-389`; `power_query_extractor/extractor_main.py:36-63`; `power_query_extractor/gui/extractor_window.py:118, 587-599` | JSON |
| 2 | Main log | `logging.FileHandler('gst_organizer.log')` - relative to the CWD | `main.py:20-27` (runs at import of main.py) | nobody (people) | text, appended |
| 3 | Extractor log | `logging.FileHandler('pq_extractor.log')` - relative to the CWD | `power_query_extractor/extractor_main.py:12-19` (runs at import of that module, i.e. only the standalone launcher) | nobody | text, appended |
| 4 | `logs\` and `temp\` folders | `Path(__file__).parent / 'logs'`, `/ 'temp'` - in the APP folder | `main.py:29-43` at every start | nobody (grep) | empty folders |
| 5 | Missing `__init__.py` | app folder `core\`, `utils\`, `gui\` | `main.py:39-43` (only if missing) | Python | text |
| 6 | Report staging | `Path(tempfile.gettempdir()) / "gst_excel_temp"` = `%TEMP%\gst_excel_temp\` | `core/excel_handler.py:53-54` (created when the main window starts), files `:147-159, 273-279` | same | xlsx, deleted after use |
| 7 | Extractor fallback caches | `Path("gst_organizer_cache.json")`, `Path(".cache/app_state.json")`, `Path("config/last_session.json")` - relative to the CWD | NOBODY in A (git: never written in any commit) | `extractor_main.py:42-44` | JSON |
| 8 | Exported client list | user-chosen `.txt` (default name `selected_clients_YYYYMMDD_HHMMSS.txt`) | `gui/handlers/client_handler.py:233-255` | user | text, Windows default encoding |
| 9 | Output files in the target folder | section 3 | Steps 3 and 4 | Step 4, Excel | xlsx / txt |

There is no run-history file, no database, no registry use, and no per-run log other than items 2-3 and
the output files of section 3 (the summary workbook and `_Organization_Report_*.txt` act as the only
record of a run).

### 4.2 Settings file: write path vs read paths (audit method step 4)

**Writer** (READ `gui/main_window.py:358-371` -> `gui/utils/cache_manager.py:29-48`):
- Keys written, exactly: `source_folder`, `itc_template`, `sales_template`, `target_folder`,
  `processing_mode` (`fresh|rerun|resume`), `workflow_mode` (`organize|full`), `include_client_name`
  (true/false), `dark_mode` (true/false), `client_name_max_length` (integer), plus `recent_folders`
  (list) added by CacheManager. RAN t8 (`t8_output.txt`).
- The whole file is rewritten each time (`open(..., 'w')`, `json.dump`): no temporary file, so a crash
  mid-write leaves broken JSON. RAN t8: a broken file loads as `{}` (defaults) and the next save starts
  `recent_folders` again from one entry.
- Encoding: `open()` without encoding, but `json.dump` writes pure ASCII (non-ASCII as `\u` escapes), so
  readers with any encoding agree. RAN t8 ("file is pure ASCII: True").
- `recent_folders`: the current source folder is added at the FRONT only if not already in the list; the
  list is cut to 5. An already-listed folder is not moved to the front, so it is not a true
  most-recent list. RAN t8. It is never read by anything (grep) - write-only data.
- **When it is saved** (READ): Browse on any of the four paths (`file_handler.py:25, 36, 47, 55`); a
  workflow radio button (`setup_tab.py:347-351`); the length spinbox arrows (`setup_tab.py:434-438`); the
  global "include client name" checkbox (`client_handler.py:211-223`); the dark-mode toggle
  (`main_window.py:196`). **Not saved:** typed paths, the processing-mode radio buttons
  (`setup_tab.py:401-412` has no command), per-client name toggles, and closing the window
  (`main_window.py:344-355`). So the processing mode and typed paths are remembered only if one of the
  save triggers happens afterwards. HIGH.
- If the length spinbox holds non-numeric text, `IntVar.get()` in `save_cache` raises (`main_window.py:369`)
  and that save is lost. MEDIUM.

**Reader 1 - main window at start-up** (READ `main_window.py:120-121, 373-389`): same file, same keys,
defaults `''`, `fresh`, `False`, `organize`, `False`, `35` (0 or negative -> 35). If `dark_mode` is true it
switches dark mode on, which itself calls `save_cache()` (`:196`) - so start-up can rewrite the file.
Agrees with the writer. HIGH.

**Reader 2 - standalone extractor** (`launch_extractor.py` -> `extractor_main.py:36-63`): tries, in
order, `~/.gst_organizer_cache.json` (added FIRST by commit c1583d6), then three CWD-relative files that
nothing writes; in each, keys `target_folder`, `target_path`, `output_folder`, `last_target` (only
`target_folder` is ever written). RAN t8: the home file's `target_folder` is found from any CWD; with
`target_folder` blank, a CWD file holding `output_folder` was used instead. Agrees with the writer for
`target_folder`. HIGH.

**Reader 3 - extractor panel** (`extractor_window.py:118, 587-599`): same home file, key `target_folder`;
used only when the panel is built with no folder and `auto_load=True` (the standalone path when
Reader 2 found nothing, `:132-139`). Embedded in the main window (Step 4) the panel never reads the file;
it takes the folder from `app.target_folder` (`gui/tabs/extract_tab.py:74-82, 108-131`), which came from
the same file through Reader 1. HIGH.

**Check of the c1583d6 fix ("Fix extractor cache path", 23-07-2026):** before it, both extractor readers
used `Path("gst_organizer_cache.json")` in the CWD while the writer used the home dotfile (edac6a3:
`extractor_window.py:28`, `extractor_main.py:39`). After it, both readers point at the same home dotfile
and key - write and read now agree (READ + RAN t8). Leftovers: the three CWD fallbacks and three unused
keys in `extractor_main.py:42-55` remain (dead, but they could pick up a stray file in whatever folder the
extractor is started from); `.gitignore:43, 64` still list `.gst_organizer_cache.json` and
`config/last_session.json`. The repository's own audit (`PROJECT_AUDIT_ANNUALSTATEMENT.md:316-317, 450,
635, 713`) still describes the mismatch as current (written before the fix). HIGH.

Other notes on this file:
- Two copies of the main app run by the same user overwrite each other's settings (last save wins); the
  extractor never writes. LOW impact.
- The test `tests/test_main_window_tabs.py:36-38` builds the real main window, which reads the
  developer's REAL home settings file, and replaces `save_cache` only after construction, so a cached
  `dark_mode: true` would rewrite the real file during tests. READ, MEDIUM.
- The file holds folder and template paths; if folders are named after clients, client names end up in
  `%USERPROFILE%`. LOW.

### 4.3 Logs (READ; RAN t11 and a Windows-encoding simulation)

- `gst_organizer.log` is created in the CWD the moment `main.py` is imported (RAN t11: created in the
  scratch CWD, not in the app folder). With Standard 13's launcher (`start "" /D "%~dp0"`,
  rebuild_pipeline_FINAL.md:962) the CWD is the APP folder on the share, so every user would append to
  one shared log file. Started by double-clicking `main.py`, Windows also uses the script's folder. From a
  shortcut, it is the shortcut's "Start in" folder. Never rotated, never trimmed.
- `pq_extractor.log`: same, but only when `extractor_main` is imported (standalone launcher). RAN t8
  (created in the scratch CWD by the import alone). When Step 4 runs inside the main window,
  `gui/tabs/extract_tab.py:8-11, 24-25` deliberately imports `extractor_window` instead, so extractor
  messages go to `gst_organizer.log`.
- Encoding: `FileHandler` without encoding uses the Windows default code page (typically cp1252). Many
  messages contain symbols (`✓ ❌ ⚠️ 🏢 ...`, e.g. `report_processor.py:242, 389, 539`). RAN t15 (a
  simulation with cp1252, `t15_output.txt`): such lines raise "--- Logging error ---" on stderr and are
  MISSING from the file. MEDIUM (depends on the PC's code page / Python UTF-8 mode).
- No log is written at all if start-up fails at `main.py:17` (4.6 / 5.2).

### 4.4 `logs\` and `temp\` folders in the app folder (READ; RAN t11)

`setup_environment()` (`main.py:29-43`) creates `logs` and `temp` next to `main.py` at every start, and
writes `__init__.py` files if missing. Nothing uses either folder (grep). RAN t11: both were created in the
copied app folder. On a share where users may not write into the app folder, `mkdir` raises, the error is
caught by `main()` (`main.py:130-134`) and the app stops with "Critical error". HIGH (READ), not run on a
read-only share.

### 4.5 Report staging folder (READ; RAN t3, t6)

`%TEMP%\gst_excel_temp\` (per user, per PC). RAN with `TMPDIR` pointing at the scratch folder: the folder
was created there and the staging files were removed after each report (`t6_output.txt`: "staging folder
left behind? []"). A crash between copy and delete leaves a file behind; the next run deletes a same-named
one first (`excel_handler.py:148-152`).

### 4.6 Settings that are NOT stored anywhere (READ)

- Extractor "Power Query Wait Time" (default 10, 5-60, `extractor_window.py:270-279`), "Refreshed File
  Suffix" (default `_Refreshed_{timestamp}`, `:297-304`) and "Skip Refresh" (`:315-325`) live only in
  memory and reset every start.
- "Skip Refresh" works by changing the module-level dictionary `EXTRACTION_OPTIONS['skip_refresh']`
  (`extractor_window.py:483-487`) that `report_processor.py:134-135` reads - shared by every extractor
  panel in the process. `auto_refresh_if_missing` and `show_timestamp` (`cell_mappings.py:250-251`) are
  never read (grep).
- The Validation tab's per-client "include name" toggles are reset to No at every scan
  (`client_handler.py:48-49`).

### 4.7 Shared state between the app's own steps (READ; RAN where marked)

#### 4.7.1 Step 3 report names -> Step 4 readers
Writer `processing_handler.py:297-298, 355-356` (`ITC_Report_...xlsx`, `Sales_Report_...xlsx` in the
version folder). Readers: `extractor_window.py:36-39` (prefix + `.xlsx`, ignoring case) and
`report_processor.py:112-113` (`glob`, case-insensitive on Windows only). Agree on the prefix. HIGH.

#### 4.7.2 Refreshed-copy names: one writer, three readers (RAN t9)
Writer `report_processor.py:128-130`. Readers: `_latest_refresh_times` (`extractor_window.py:45-82`,
via `refresh_naming.is_refreshed_name`), the skip-refresh lookup (`report_processor.py:142-146`, via
`is_refreshed_copy_of`), and the report picker (`report_processor.py:112-125`, plain glob, does NOT
exclude refreshed copies - 3.5 item 5). Commit 2aaa2b1 made the first two follow the configured suffix.
Remaining disagreements (RAN t9):
- Suffix `_R`: marker `_r` is found inside every `itc_report_...` name, so every ORIGINAL counts as
  refreshed (Last Refresh shows the original's time; skip-refresh may "reuse" an unrefreshed original).
- Suffix `_{timestamp}`: marker `_` - same problem.
- Suffix `{timestamp}` or blank: falls back to `_refreshed_`, so files written with that suffix are never
  recognised; a blank suffix also makes the copy target the original itself (error "are the same file").
- Suffixes are not checked for characters Windows forbids in file names.
HIGH (RAN).

#### 4.7.3 Extracted values -> consolidated report: WRONG-COLUMN bug (RAN t10)
Writer `report_processor.py:556-594` stores each value under `"<output sheet>_<output column>"` with
spaces turned into `_` (key built in `cell_mappings.py:264-278`) and under the bare cell reference.
Reader `data_consolidator.py:138-150` first looks for the first stored key that CONTAINS the output
column name, and only then the exact key. Column names with a space never match (keys have `_`), so they
fall through to the exact key and are correct. Column names without spaces - `IGST`, `CGST`, `SGST` of
the sheets "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" - match the FIRST key containing them,
which is `GSTR-3B_Difference_Diff_IGST/CGST/SGST`. RAN t10 with invented marker numbers: those 9 cells of
the consolidated report showed the values read from `Diff GSTR-3B` J15/K15/L15 instead of `IMS_Reco`
AA1/AD1/AG1, `ExtraInPurchase` O1/P1/Q1 and `ExtraInGSTR2B` S1/T1/U1; the other 13 columns were right.
(If J15/K15/L15 are empty the wrong match finds `None` and the exact key then gives the right value.)
The reader is unchanged since edac6a3; the affected columns arrived with 47a1492 (shared with B). This is
a technical key-matching defect; which cell is the correct source is the author's call. HIGH (RAN).

#### 4.7.4 Sanitised names inside the client data (READ; RAN t12)
`organize_files` renames the in-memory file entries to the sanitised name after a successful copy
(`file_organizer.py:233-246`); `prepare_template_data` then links those names. Agrees in fresh/rerun;
disagrees for skipped files in same-minute resume (3.3). The entries are the SAME objects as in
`scanned_files` (`file_parser.py:199-211`), so after a run the Dry Run shows sanitised names until a
re-scan. LOW impact.

#### 4.7.5 `.gitignore` versus real file names (READ `.gitignore:40-57, 64`)
Lists `gst_organizer.log`, `pq_extractor.log`, `.gst_organizer_cache.json` (only matters if the home
folder were the repository), `temp/`, `logs/`, `temp_gst_*`, `backup_*`, `/Annual Statement-*/`. The real
backup names are `<stem>_backup_<ts>` (not `backup_*`); `temp_gst_` is an unused constant
(`constants.py:237`).

---

## 5. Environment

### 5.1 Interpreter and packages (READ)

- `requirements.txt`: comment "Python 3.7 or higher"; `openpyxl>=3.1.2`; `pywin32>=305; sys_platform ==
  'win32'`; `lxml>=4.9.0` ("optional"); `pytest>=7.0.0` (tests). Minimum versions only, not exact
  (Standard 13, rebuild_pipeline_FINAL.md:919, 937, asks for exact versions and no pytest). I found no
  syntax newer than 3.7 (grep for `:=`, `match`, `X | None`, `list[...]`); not run on 3.7. LOW-MEDIUM.
- Tkinter must be present (part of python.org Windows installers; the embeddable package does not include
  it). The author's Standard 12 names PySide6 for new desktop apps.
- Start-up check (`utils/requirements_check.py:66-90`): openpyxl essential; pywin32 (`pythoncom`)
  essential ON WINDOWS ONLY (commit be6418c). Missing packages trigger an OFFER to run
  `<this python> -m pip install -r requirements.txt` (or the package names), 600-second timeout, then a
  restart with `sys.executable + sys.argv` (`utils/dependency_installer.py:23, 26-74, 88-158`, `main.py:78-95`).
  Under Standard 13 that would install into the shared `_runtime` from a user's PC
  (rebuild_pipeline_FINAL.md:924, 937-940). HIGH (READ).
- `pytest.ini` has no `pythonpath = .`; `tests/conftest.py:9` adds the project folder itself.
- `launch_extractor.py:10` APPENDS the app folder to `sys.path` (main.py inserts it first,
  `main.py:31-32`). With a portable Python that does not add the script folder itself, another package
  called `utils` earlier on the path would be imported instead (`report_processor.py:39` imports
  `utils.excel_com`). LOW.

### 5.2 Start-up and the console (RAN t11)

- `main.py:17` replaces `sys.stdout` with a UTF-8 wrapper of `sys.stdout.buffer`. Under `pythonw.exe`
  `sys.stdout` is `None`. RAN t11 (simulated with `sys.stdout = None`): `AttributeError: 'NoneType' object
  has no attribute 'buffer'` at import - before logging is set up and outside any `try`, so the app would
  die with no window, no message and no log. HIGH.
- `main.py` reports problems with `print()` and waits with `input()` (`_pause`, `main.py:45-54`); under
  pythonw `input()` raises `RuntimeError`, which `_pause` does not catch. READ, MEDIUM (unreachable today
  because of line 17).
- No `sys.excepthook` / `threading.excepthook`; worker threads catch their own errors
  (`processing_handler.py:522-530`, `extractor_window.py:957-966`).

### 5.3 Windows-only calls (READ; cannot run here)

| Where | Call | Guard |
|---|---|---|
| `core/excel_handler.py:21-28` | `import win32com.client, pythoncom` | `platform.system()=='Windows'` + `try/except ImportError` -> `USE_WIN32COM` |
| `core/excel_handler.py:120-252` | `CoInitialize`, `GetActiveObject`/`Dispatch("Excel.Application")`, `Workbooks.Open`, `Range.Value`, `Sheets.Add`, `Save`, `Close`, `Quit`, `CoUninitialize` | only when `USE_WIN32COM` |
| `utils/helpers.py:580-590` | `win32api.GetShortPathName` | Windows check + broad `except` |
| `power_query_extractor/core/report_processor.py:20-36` | `import pythoncom, win32com.client, win32api, win32con, win32gui` | `try/except ImportError` -> `WIN32_AVAILABLE`; `process_client` returns early (`:66-77`) |
| `report_processor.py:218-227` | `Dispatch("Excel.Application")`, `Visible=True`, `WindowState=-4137` (maximised) | inside try |
| `report_processor.py:260-293` | `Dispatch("WScript.Shell")` (Windows Script Host), `AppActivate`, `SendKeys("%a")`, `"r"`, `"a"` (Alt, A, R, A = Data > Refresh All, by keyboard) | inside try |
| `report_processor.py:417-525` | `win32gui.EnumWindows/GetWindowText/GetClassName/EnumChildWindows/SendMessage(BM_CLICK)/PostMessage(WM_CLOSE)` to find and close error dialogs | inside try |
| `report_processor.py:720-731, 762-766` | `\\?\` and `\\?\UNC\` long-path prefixes; `GetShortPathName` | inside try |
| `gui/handlers/client_handler.py:255` | `os.startfile` (open the exported list) | inside try; off Windows the file is written but an "Export Error" is shown |
| `gui/handlers/processing_handler.py:505-516` | `os.startfile` / `open` / `xdg-open` for the output folder | platform check |
| `power_query_extractor/gui/extractor_window.py:973` | `os.startfile(report)` | try/except, logs only |

No `ctypes.windll`, `msvcrt` or `winreg` anywhere (grep). The Excel-specific risks (taking over the
user's Excel, hiding it, Quit, SendKeys focus) are in the A-X-EXCEL notes.

### 5.4 Environment variables, user profile, registry (READ)

- No `os.environ` / `getenv` anywhere. Indirect: `Path.home()` (`USERPROFILE` on Windows, `HOME`
  elsewhere) for the settings file; `tempfile.gettempdir()` (`TMP`/`TEMP`/`TMPDIR`) for staging. RAN t3/t8
  with `HOME`/`TMPDIR` redirected.
- No registry reads or writes. README.md:275 tells users to run `pywin32_postinstall.py -install` as
  administrator (a registry/system change outside the app).

### 5.5 Network shares, UNC paths and drive letters (READ)

- Designed for source and target folders on a network share (comments and commits a603c02, 56aaad9,
  4f84ba9 measure an SMB share reached over a VPN; `setup_tab.py:135` "avoid stat'ing a UNC path").
- Folder paths are stored and written into reports exactly as picked (drive letter or UNC). Tk returns
  forward slashes; `clean_windows_path` turns them into backslashes for the Links cells only.
- The extractor converts long UNC paths to `\\?\UNC\server\share\...` before `Workbooks.Open`
  (`report_processor.py:727-731`). Whether Excel accepts `\\?\` paths I am not sure (LOW; check on
  Windows with one long path). 8.3 short names are often unavailable on network shares, so the
  `GetShortPathName` fallbacks may change nothing (MEDIUM).
- Folder "newest" choices use mtime, which on a share is the server's clock/semantics.

### 5.6 Path-length limits (READ)

- ~218 characters: the figure used for Excel/Power Query (`constants.py:85-90`, `excel_handler.py:356`,
  `report_processor.py:720`). Windows' classic 260 limit (`MAX_PATH`, README.md:317) applies to Python's
  own copies unless Windows long paths are enabled. No pre-check exists (3.7).

### 5.7 Current working directory (READ; RAN t8, t11)

Depends on the CWD: both log files (4.3) and the three extractor fallback caches (4.1 item 7). Everything
else uses the app folder (`main.py:31, 84`; `launch_extractor.py:10`), the home folder, `%TEMP%`, or paths
the user chose.

### 5.8 Excel and desktop requirements (READ)

- Microsoft Excel on the PC that runs Step 3 (COM path) and Step 4. README.md:55 says "Excel 2016+"; the
  code checks no version.
- Step 4 needs an unlocked, interactive desktop: Excel is made visible and maximised and keystrokes are
  typed into the focused window (`report_processor.py:225-227, 258-293`); the Alt-A-R-A key tips and the
  English dialog words/button texts (`:425-436, 506`) assume an English Excel/Windows UI. MEDIUM.
- Fixed waits: 3 s after each report, 5 s loading, the configured wait (10 s default), 5 s error checks,
  at least 10 s extended wait (`report_processor.py:185, 253-255, 300-302, 310-319, 326-330`).

### 5.9 Fonts and mouse (READ)

Fonts `Segoe UI`, `Consolas`, `Arial` (Windows fonts; Tk substitutes elsewhere). Mouse-wheel handlers use
the Windows `<MouseWheel>`/`delta/120` convention (`setup_tab.py:201, 220-248`,
`extractor_window.py:458-478, 559-565`); X11 Linux sends Button-4/5 instead, so wheel scrolling would not
work there. LOW relevance.

---

## 6. Storage locations compared with the author's Standard 13 (rebuild_pipeline_FINAL.md:900-1061)

| A's location | Standard 13 rule | Verdict |
|---|---|---|
| Settings `%USERPROFILE%\.gst_organizer_cache.json` | "Settings and logs go in the user's own folder (`%LOCALAPPDATA%\[App Name]\`), never in the shared app folder" (1039-1040) | Per-user (good) but the wrong folder: profile root, not `%LOCALAPPDATA%\<App Name>\`. |
| `gst_organizer.log`, `pq_extractor.log` in the CWD | Same rule; logs in `%LOCALAPPDATA%\[App Name]\` (1035-1036, 997); "Find the app's own files from `APP_DIR`, never from the current folder" (1034) | Breaks both: CWD-relative, and with the Standard 13 launcher the CWD is the shared app folder, so all users write one file there. |
| One log file written by several people | "Never keep ... any file several people write at the same time, in the shared folder" (1046-1047) | Breaks it (same reason). |
| `logs\`, `temp\` (and possibly `__init__.py`) created in the app folder at every start | Users' PCs must not write into the shared folder (`-B`, 969-970; 1039-1040) | Breaks it; also stops the app on a read-only share (4.4). |
| Extractor fallback caches in the CWD | 1034 | Breaks it (dead, but still consulted). |
| `%TEMP%\gst_excel_temp\` | Not covered explicitly | Per-user and local - compatible in spirit. |
| Fatal errors | `report_fatal_error` writes `%LOCALAPPDATA%\<App>\logs\error_*.txt` and shows a message box (980-1023) | Not present; under pythonw the app dies silently at `main.py:17` (5.2). |
| `print()` / `input()` for problems | "There is no console: print() shows nothing" (1035) | Breaks it (`main.py:62-127`). |
| Offer to `pip install` into the running Python | "Nothing is installed on users' PCs"; never upgrade `_runtime` packages without asking (924, 937-940) | Breaks it if used with the shared runtime. |
| `requirements.txt` with `>=` and pytest | Exact versions of imported packages, no pytest (919, 937) | Does not match. |
| Finding its own files | `main.py:31-32, 84` and `launch_extractor.py:10` use the script's folder | Matches (except logs/caches above). |
| `pytest.ini` | `[pytest]`, `pythonpath = .`, `testpaths = tests` (945-946) | Partly: no `pythonpath = .`. |
| Passwords / tokens | none kept | Nothing to compare. |

---

## 7. Checklist: what the author needs to supply for a rebuild

1. **ITC template** (the real file, `.xltx` or `.xlsx` - say which is used in practice), containing:
   - a sheet `Links`; cells B2, B4, B5, B7, B8, B10, B11 and what each query does with them (folder vs
     file stem; which extension the query adds; whether folder queries filter by name);
   - the Power Query queries themselves (names and what files they read), their privacy-level setting,
     and whether they read a whole folder (B2) or one file;
   - after refresh: sheet `Info` with the client name in `BB2`; sheets `Diff GSTR-3B` (J15:L15), `Reco`
     (C37:E37), `T6 of R9` (K29), `T8 of R9` (G6), `IMS_Reco` (AA1, AD1, AG1), `ExtraInPurchase`
     (O1:Q1), `ExtraInGSTR2B` (S1:U1) - confirm names and cells [GST RULE];
   - what the template's Links cells contain as SHIPPED (because blank values leave them unchanged, 1.3.1);
   - whether the file is marked read-only on the share (1.4).
2. **Sales template**, same questions: `Links` B2, B3, B5, B6, B8, B9; after refresh `Info!BB2` and
   `Diff - GSTR-1` (I1, N1:Q1).
3. **One anonymised sample of each input type**, named exactly as the portal/office produces them, with
   real-looking but invented contents:
   GSTR-2B Reco, IMS Reco (including a ` (1)` duplicate), GSTR-3B export (one month; plus a `(1) GSTR3B-`
   prefixed one if that is how the portal names repeats), Sales (start and end month), Sales Reco, Annual
   Report; plus one `.xls` if old-format files still occur. Each must be a genuine Excel file of at least
   1 KB.
4. **A sample source folder** listing (names only, invented clients) showing: client names with spaces,
   `Private Limited`, hyphens, apostrophes, `&`, long names; the different ways the state is written
   (full name vs code); any subfolders people use.
5. **A sample output tree** from a real run (names only, anonymised), with the real target root (drive
   letter or UNC, and the `FY ...` level), to re-check the 218-character margin.
6. **One refreshed ITC and one refreshed Sales report** (anonymised) to confirm the extractor cells.
7. **PC facts**: Windows and Excel versions and UI language; whether Step 4 runs on the user's own
   desktop; whether everyone maps the share to the same drive letter.

---

## 8. Other technical findings noticed while tracing (for the audit writers)

- `FileOrganizer` creates the target folder if it does not exist (`file_organizer.py:60-69`) while Step 1
  only requires it to be non-empty for processing (`file_handler.py:138-140`): a typo creates a new folder.
- `cleanup_on_error`, `get_summary` (`file_organizer.py:381-425`) and `validate_results`
  (`file_parser.py:338-371`) are never called; `EXCEL_SAFETY`, `EXCEL_SIGNATURES`,
  `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB` are unused (grep).
- `processing_handler.py:319-414` calls `self.app.log_message` directly from the worker thread (not via
  `_log_async`); `UIHelpers.log_message` itself schedules with `after(0)`, so it is marshalled - noted
  only because the same file mixes both styles.
- `export_client_list` writes with the Windows default encoding (`client_handler.py:241`); a client name
  outside that code page raises and shows "Export Error".

---

## Possible client names

- `PROJECT_AUDIT_ANNUALSTATEMENT.md:130, 133, 177, 179, 180, 182, 232` - example file names use a
  company name (may be a well-known company used as a generic example, or a real client). Not copied.
- No other possible client names found in the files I read, the tests I read (`ABC`, `XYZ`, `Client`
  only) or the commit messages I read. Commit messages 7347383 and 4f84ba9 contain the office server and
  share folder path (infrastructure names, not clients); not copied beyond "an `FY 2025-26` folder".

## Questions for the author

1. Which template format do you actually use, `.xltx` or `.xlsx`? (An `.xltx` copied under an `.xlsx`
   name may not open in Excel - 1.4.)
2. How do the template's queries read the Links cells, which extension do they add to the file stems, and
   do the folder queries (GSTR-3B folder B2) read every file or filter by name? (Backups and second
   clients can land in that folder - 3.3, 3.4.)
3. What do the Links cells of your templates contain before the app touches them? Blank values leave them
   unchanged (1.3.1).
4. Are templates kept read-only on the share (1.4)?
5. Do all PCs reach the share by the same drive letter, or by UNC (`\\server\...`)? Reports carry
   absolute paths.
6. Is `Info!BB2` meant to hold the client name after refresh (the only refresh check)?
7. Is Step 4 run on the user's own desktop, and in which Excel UI language? (Keyboard key tips.)
8. Do client names ever contain hyphens, or start with digits/leading zeros? (2.2, 2.6.4.)
9. Do two clients in one state ever share a long name prefix (3.4)?
10. Should files of different periods/years ever sit in the same source folder? The code has no period
    logic (2.2) - this is your rule to set.
11. [GST RULE] Are the extraction sheets/cells in 1.3.4 and the six expected file types in 2.6.2 still
    current?

## INCOMPLETE

- Excel-through-COM behaviour (opening a renamed `.xltx`, `\\?\` paths, read-only copies, SendKeys) could
  only be traced by reading - no Windows/Excel here. Each such point says how to check.
- `PROJECT_AUDIT_ANNUALSTATEMENT.md` was searched, not read in full; the test files other than those named
  in section 0 were not read line by line (only searched).
- Not run on Python 3.7 or on a read-only share.

## Appendix: RAN scripts (all under /root/merge-scratch/run/A-X-EXT/, invented data only)

| Script | What it shows | Output |
|---|---|---|
| `t1_recogniser.py` | 44 invented names through `FileParser.parse_filename`; state-code function | `t1_output.txt` |
| `t2_scan.py` | `scan_folder` on an invented folder: filters, sort order, grouping, duplicates, variations | `t2_output.txt`, `t2b_sort_output.txt` |
| `t3_organise.py` | Created tree, path lengths, Links values, include-client-name, sanitise/key examples | `t3_output.txt` |
| `t4_tk_values.py` | Treeview turns `'007'` into `7` (xvfb) | `t4_output.txt` |
| `t5_collision.py` | Two clients sharing one Level 2 folder | `t5_output.txt` |
| `t6_template_fallback.py` | openpyxl path: Links lookup, missing sheet, `.xltx` content type, Power Query part dropped | `t6_output.txt` |
| `t7_blank_cells.py` | Blank mapping values leave template placeholders | `t7_output.txt` |
| `t8_cache.py` | Settings file writer vs three readers, recent_folders, corrupt file, CWD fallback, pq_extractor.log | `t8_output.txt` |
| `t9_refresh_names.py` | Refreshed-file naming vs readers; glob order; blank suffix | `t9_output.txt` |
| `t10_consolidator.py` | Wrong-column key matching in the consolidated report | `t10_output.txt` |
| `t11_startup.py` | `main.py` under `sys.stdout=None`; log in CWD; `logs`/`temp` in app folder | `t11_output.txt` |
| `t12_modes.py` | fresh / resume / rerun / same-minute resume | `t12_output.txt` |
| `t13_fresh_same_minute.py` | Backups inside the linked GSTR3B folder | `t13_output.txt` |
| `t14_extractor_scan.py` | Step 4 folder choice by mtime; pointing at a run folder finds nothing (xvfb) | `t14_output.txt` |
| `t15_log_encoding.py` | Log lines with symbols vanish from a cp1252 log file (simulation of the Windows default) | `t15_output.txt` |

Runs used `python -B`, `PYTHONDONTWRITEBYTECODE=1`, `HOME` and `TMPDIR` redirected into the scratch
folder; nothing was written to A-copy, B-copy or /home/user.
