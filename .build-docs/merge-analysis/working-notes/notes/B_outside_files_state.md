# B-X-EXT: Project B, everything the code needs from outside the repository, and all shared state

Project B = goyal-sidhant/AnnualStatement_oldv2, snapshot HEAD b0e7d0e (09-02-2026), PyQt5 GUI, calls itself "v4.0".
Read from /root/merge-scratch/B-copy (line numbers = repository line numbers). History from /root/merge-scratch/combined (ref B/main, split point 47a1492).
All paths below are repository-relative.

## 0. Method, legend and scope

- Every `.py` file in B (41 files, 7,805 lines), plus `README.md`, `requirements.txt` and `.gitignore`, was read in full with the Read tool. Post-split diffs (47a1492..B/main) were read for the files that touch external inputs or state: core/*, utils/helpers.py, power_query_extractor/*, main.py, launch_extractor.py, .gitignore, requirements.txt, models/*, gui/controllers/extraction_controller.py. The commit messages of c258047 and b0e7d0e were also read. Pre-split writers were checked where B reads their files: 47a1492:gui/utils/cache_manager.py, 47a1492:gui/main_window.py:328-356 and 47a1492:power_query_extractor/gui/extractor_window.py:20-45 and 473-486.
- The user said a README.md had been added to "the repo". Both were read. B's `README.md` (364 lines) is the v3.0 document, last changed in the shared commit d674386; B never updated it after the split. `/home/user/Annual_Statement_2.0/README.md` holds a single heading line ("# Annual_Statement_2.0") and nothing to act on.
- **READ** = established by reading the code. **RAN** = executed. RAN checks used copies under /root/merge-scratch/run/B-X-EXT/ with `/root/merge-scratch/venv-B/bin/python -B` (Python 3.14.6, PyQt5 5.15.11 / Qt 5.15.14, openpyxl 3.1.5, Linux, no Excel). HOME and TMPDIR were redirected into that folder. Nothing was written to B-copy or to /home/user. One exception to note: a read-only `git status --porcelain` was run on /home/user/AnnualStatement_oldv2 to confirm it was clean (it printed nothing). That command changed only the `.git` folder's own timestamp, by creating and removing a temporary lock file. `.git/index`, HEAD, refs and all working files are unchanged (checked with `ls --time-style=full-iso`).
- Windows-only behaviour (Excel through COM (Windows' way for one program to drive another, here Python driving Excel), SendKeys, win32gui, os.startfile, Windows path rules) could only be traced by reading. For one test (t6, t8), small **stub** modules (empty stand-ins) named pythoncom/win32com/win32api/win32con/win32gui were placed in the run folder. Without them the app cannot even be imported on Linux (see 5.1). They only let the non-Windows logic run.
- Confidence for inferences: HIGH / MEDIUM / LOW.
- **[GST RULE]** marks anything that decides a tax figure, a return/table/sheet/cell, a rate/limit/threshold, rounding, a period, GSTIN/PAN/state-code logic, ITC eligibility or a due date. For these the notes only describe what the code does, with no judgement. **TIME-VARYING** marks values that can change over time.
- Plain-English glossary of terms used below:
  - **regex**: a pattern rule for matching text.
  - **glob**: a wildcard file search such as `*.xlsx`.
  - **JSON**: a plain-text format for settings.
  - **Power Query (PQ)**: Excel's built-in data-import and refresh feature.
  - **cwd**: the current working directory, the folder a program was started "in".
  - **UNC path**: a network path such as `\\server\share\...`.
  - **8.3 short path**: the old DOS-style short name Windows can give a long path.
  - **QThread**: a background worker thread in Qt.
  - **stem**: a file name without its extension.

---

## 1. Templates the code expects

### 1.1 How many, how chosen, where the path is stored
- Exactly **two** templates: an **ITC template** and a **Sales template** (READ gui/views/setup_view.py:218-261).
- **How chosen.** A file dialog with filter `"Excel Files (*.xlsx *.xlsm)"` (setup_view.py:610-622), or by typing a path into the text box. The label says "(.xlsx or .xlsm)" (setup_view.py:221).
  - The constant `TEMPLATE_EXTENSIONS = {'.xlsx', '.xltx', '.xltm'}` (utils/constants.py:180) is **never used** (grep: only its definition).
  - README.md:120-121 says ".xlsx or .xltx". The README, the dialog and the dead constant disagree (READ, HIGH).
- **Where the path is stored.** In memory: `AppSettings.itc_template` and `AppSettings.sales_template` (models/app_state.py:14-15). On disk: keys `"itc_template"` and `"sales_template"` in `~/.gst_organizer/settings.json` (models/cache_manager.py:35, 41-55; see section 4). The path is stored as typed or as returned by the dialog; on Windows Qt's dialog returns forward slashes, e.g. `C:/...` (MEDIUM, Qt convention).
- **Checks before use.**
  - pyqt_main_window.py:353-358: the template must be non-empty and must exist, else a warning box.
  - processing_controller.py:210-218: repeats the check per client (logs an error, counts `report_errors`).
  - excel_handler.py:78-79: raises FileNotFoundError.
  - Nothing checks the extension, the sheets or the queries before processing (READ, HIGH).

### 1.2 What happens to the template file
- For each selected client, the template is **copied** with `shutil.copy2` to `<Version folder>/ITC_Report_<client>_<statecode>_<DDMMYY_HHMM>.xlsx`, and likewise `Sales_Report_...xlsx` (gui/controllers/processing_controller.py:112-130, 220; core/excel_handler.py:124 or 228). The copy is then opened and cells are written. The template itself is never modified (READ, HIGH).
- The output extension is **always `.xlsx`** whatever the template's extension (processing_controller.py:220).
  - RAN t4: an invented `.xltx` template put through the openpyxl path produced a file named `.xlsx` whose internal content type is still "template" (`spreadsheetml.template.main`).
  - In the COM path the raw file is copied byte-for-byte (excel_handler.py:124), so an `.xlsm` or `.xltx` template becomes a macro or template package with an `.xlsx` name.
  - Not sure whether Excel then refuses to open it. To check: on an office PC copy a real `.xlsm` template to a `.xlsx` name and double-click it (MEDIUM that Excel shows "file format or extension is not valid").

### 1.3 Two code paths for writing the template copy (READ)
- **COM path** (`_create_report_win32com`, excel_handler.py:95-217) is used when `platform.system() == 'Windows'` and `import win32com.client, pythoncom` succeeds (excel_handler.py:20-28).
- **openpyxl fallback** (`_create_report_openpyxl`, excel_handler.py:219-262) is used otherwise. It loads with `keep_links=True` and saves. It logs "Power Query content may be lost!" (excel_handler.py:224); the author's Standard 12 notes openpyxl drops Power Query on save.
- In the running app the fallback is **effectively unreachable**. On any machine without pywin32 the app cannot start at all, because report_processor.py imports pythoncom and the win32 modules unconditionally (RAN, see 5.1). So the app only runs where COM is available. The fallback can only run when core/excel_handler.py is used on its own, as in RAN t4 (HIGH).

### 1.4 What the code assumes is inside the templates

**A. A "Links" sheet** ([GST RULE] where it decides which file feeds which report; otherwise technical):
- **Sheet lookup** in the COM path (excel_handler.py:131-153), in this order:
  1. a sheet whose name equals `Links`, case-insensitive;
  2. else the first sheet whose name contains `link`, case-insensitive (logs a warning);
  3. else **a new sheet named `Links` is added and the cells are written into it** (excel_handler.py:150-153). The report is then reported as a success, but nothing in the workbook reads the new sheet (READ, HIGH).
- In the openpyxl path only the exact case-insensitive `Links` is tried. If it is missing, nothing is written and the function still returns True (excel_handler.py:238-258). RAN t4: a template without a Links sheet gave `ok=True` with no cells written.
- **Cells written** (utils/constants.py:96-120). The value is always written as **text**, `str(value)`, and only if non-empty (excel_handler.py:159-165 / 246-252).

  | Template | Cell | Key | Value written (excel_handler.py:540-627) |
  |---|---|---|---|
  | ITC | B2 | gstr3b_folder | full Windows path of the Version's GSTR-3B category folder (always written, even with no 3B files) |
  | ITC | B4 | annual_folder | full path of the Version folder (only if an Annual Report exists) |
  | ITC | B5 | annual_filename | **stem** of the first Annual Report's copied (sanitised) name |
  | ITC | B7 | gstr2b_folder | full path of the "Other ITC related files" folder (only if a GSTR-2B Reco exists) |
  | ITC | B8 | gstr2b_filename | stem of the first GSTR-2B Reco |
  | ITC | B10 | ims_folder | full path of the "Other ITC related files" folder (only if an IMS Reco exists) |
  | ITC | B11 | ims_filename | stem of the first IMS Reco |
  | Sales | B2 | sales_folder | full path of the "Sales related files" folder (only if a Sales file exists) |
  | Sales | B3 | sales_filename | stem of the first Sales file |
  | Sales | B5 | annual_folder | full path of the Version folder (only if an Annual Report exists) |
  | Sales | B6 | annual_filename | stem of the first Annual Report |
  | Sales | B8 | sales_reco_folder | full path of the "Sales related files" folder (only if a Sales Reco exists) |
  | Sales | B9 | sales_reco_filename | stem of the first Sales Reco |

- RAN t4 confirmed the exact cell values on invented data, e.g. B5 = `AnnualReport-ClientA-Maharashtra-2024` with no extension, B2 = a backslash path.
- **Cells B3, B6 and B9 (ITC) and B4 and B7 (Sales) are never written.** If the template keeps other wiring there (for example the extension, or a combined path formula), it is untouched. Not sure; question for the author.
- "First file" = `files[0]`: the first file of that type in scan order, which is sorted by full path (utils/helpers.py:262). Any second Annual Report, 2B or IMS file is copied but not linked (READ, HIGH).
- The **template's own queries** must turn these cells into file paths. The code assumes, but never checks, that the queries:
  - read cell B2 as a **folder** (probably all files in it), since only a folder is given for GSTR-3B;
  - read the other cells as folder + stem, adding the extension themselves, since only stems are written. If an input is `.xls`, the query would need to know that. Not sure; question for the author.
  - It cannot be verified from the repository, which holds no templates (MEDIUM).
- **Cells left empty:** when a file type is missing, its folder and stem keys are `''` and are **not written**. The template's own previous contents of those cells stay in place (excel_handler.py:161-162).

**B. An "Info" sheet with cell BB2** (used after the Power Query refresh, report_processor.py:589-636):
- The sheet name is matched case-insensitively (`'info'`).
- The refresh counts as "validated" only if BB2 is not empty, not blank, not `0` or `0.0`, and not one of `#REF! #VALUE! #NAME? #NULL! #DIV/0! #N/A #NUM! #ERROR!`.
- The log text calls BB2 the "Client Name" (report_processor.py:605-606, 632). So the template's queries are expected to place the client name in Info!BB2 when the refresh works (MEDIUM).
- If the Info sheet is missing, validation fails, the refresh is marked failed, and **all extracted values for that report are shown as "Missing"** in the consolidated report (data_consolidator.py:116-117) (READ, HIGH).

**C. Sheets and cells read by the Extract step** [GST RULE] TIME-VARYING. Which sheet and cell a figure comes from, and the "Table 6J" and "Table 8D" labels, are GST table references (power_query_extractor/config/cell_mappings.py:33-242):
- Sheet names are matched **exactly, case-sensitively** (report_processor.py:515-519). Values are read through COM (`ws.Range(cell).Value`, report_processor.py:535) from the **refreshed copy**.

  | Report | Input sheet (exact) | Cells | Output sheet / column (consolidated report) |
  |---|---|---|---|
  | ITC | `Diff GSTR-3B` | J15, K15, L15 | "GSTR-3B Difference": Diff IGST / Diff CGST / Diff SGST |
  | ITC | `Reco` | C37, D37, E37 | "Reco Diff": IGST Diff / CGST Diff / SGST Diff |
  | ITC | `T6 of R9` | K29 | "Table 6J Match": Table 6J Match |
  | ITC | `T8 of R9` | G6 | "Table 8D Limits": Table 8D Limits |
  | ITC | `IMS_Reco` | AA1, AD1, AG1 | "IMS Reco": IGST / CGST / SGST |
  | ITC | `ExtraInPurchase` | O1, P1, Q1 | "Extra In Purchase": IGST / CGST / SGST |
  | ITC | `ExtraInGSTR2B` | S1, T1, U1 | "Extra In GSTR2B": IGST / CGST / SGST |
  | Sales | `Diff - GSTR-1` | N1, O1, P1, Q1 | "GSTR-1 Difference": Diff Taxable Value / Diff IGST / Diff CGST / Diff SGST |
  | Sales | `Diff - GSTR-1` | I1 | "Turnover GSTR-1": Taxable Value |

- RAN t7 printed the generated `CELL_MAPPINGS` and matches this table.
- A missing sheet gives `None` for its fields, then "Missing" in the output (report_processor.py:521-527; data_consolidator.py:141-142).
- **Technical defect** (RAN t7). For the three ITC sheets whose output columns are single words (`IGST`, `CGST`, `SGST`: "IMS Reco", "Extra In Purchase", "Extra In GSTR2B"), the consolidator fills the column with the first extracted field whose name contains that word (data_consolidator.py:129-133). That is the `Diff GSTR-3B` J15/K15/L15 value. With distinct invented marker values, all three output sheets showed `<Diff GSTR-3B!J15>`, `<Diff GSTR-3B!K15>` and `<Diff GSTR-3B!L15>` instead of the IMS_Reco, ExtraInPurchase or ExtraInGSTR2B cells.
  - Multi-word columns (e.g. "Diff IGST") are not affected. Their space never matches the underscore in field names, so they fall through to the exact-key lookup at data_consolidator.py:136-138.
  - Which cell should feed which column is the author's GST call. The technical point is only that the lookup does not use the configured cell.
- **Number display** [GST RULE] (rounding/display): numeric non-zero values get Excel format `#,##0.00` (data_consolidator.py:149-150). The value itself is not rounded.

**D. The Excel ribbon**, which the template does not control but the refresh depends on:
- The refresh sends the keystrokes `Alt+A`, then `R`, then `A` to the Excel window (report_processor.py:236-249). This presses "Data > Refresh All" through ribbon key tips.
- This assumes an Excel with the ribbon and **English key tips** (MEDIUM). Section 5 has more.

**E. Any Power Query queries at all**: never checked. Standard 12 asks to check for a `customXml/item*.xml` part containing `DataMashup`; B does no such check (READ, HIGH).

---

## 2. Input files recognised

### 2.1 Which files are looked at (READ utils/helpers.py:249-266; core/file_parser.py:126-169)
- Only the **top level of the chosen source folder** is searched; sub-folders are ignored. The search is `Path.glob` (not recursive) for `*.xlsx`, `*.xls` and `*.xlsm`, in that order, then sorted (helpers.py:255-262). RAN t2: a file in `sub/` was not seen.
- Files whose name starts with `~$` (Excel's lock files) are dropped silently (helpers.py:260). RAN t2: confirmed.
- Other extensions (`.csv`, `.xlsb`, `.zip`, `.pdf`, ...) are **silently ignored**: not counted and not listed as variations. RAN t2: the `.csv` was not mentioned anywhere.
- **Extension case.**
  - On Linux the glob is case-sensitive: `X.XLSX` was not found (RAN t2).
  - On Windows, Python's `Path.glob` matches case-insensitively by default, so `.XLSX` or `.Xls` would be found (HIGH, Python documentation behaviour; not runnable here).
- For each file found (file_parser.py:171-256):
  1. `stat()` fails: a warning in the log, and the file is skipped. It appears nowhere in the UI.
  2. Size < 1024 bytes: added to "variations" (the code's word for problem files) with issue `File too small to be valid Excel`, and not parsed.
  3. The first 8 bytes do not start with `PK\x03\x04` (zip, used by xlsx/xlsm) or `\xd0\xcf\x11\xe0` (old OLE format, used by xls): variation `Invalid or corrupted Excel file`.
  4. The file cannot be read: variation `Cannot read file: <error>`.
  5. Otherwise the file name is parsed (2.2). Parsed-but-invalid and unparsed files go to variations with issue `No matching pattern found` and a `pattern_expected` hint (2.4).
- Files in steps 2-4 are **not** in `scanned_files`, so the "Files found" statistic excludes them. RAN t2: 16 files found by the glob, statistic `total_files` = 14.
- **The file's content is never read beyond the 8-byte signature.** What is inside each input file matters only to the templates' Power Query, not to the Python code (READ, HIGH).

### 2.2 File-name patterns, exactly as coded (utils/constants.py:13-62) [GST RULE] TIME-VARYING (return and file types)
- All patterns are compiled with `re.IGNORECASE`. Each must match the **whole** name (`^...$`).
- They are tried **in this order**, and the **first match wins** (`break` at file_parser.py:118): GSTR-2B-Reco, ImsReco, GSTR3B, Sales, SalesReco, AnnualReport.
- The extension part `\.xlsx?$` accepts only `.xls` or `.xlsx`. **`.xlsm` files are found by the glob but never match any pattern**, so they always become variations. RAN t1/t2: confirmed.

| # | Key | Regex (verbatim) | Type | Category | "folder" field | Groups | Copied to (file_organizer.py:288-295) |
|---|---|---|---|---|---|---|---|
| 1 | GSTR-2B-Reco | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | GSTR-2B Reco | ITC | itc | client, state, period | Other ITC related files |
| 2 | ImsReco | `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$` | IMS Reco | ITC | itc | client, state, date (8 digits) | Other ITC related files |
| 3 | GSTR3B | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | GSTR-3B Export | GSTR3B | gstr3b | client, state, month | GSTR-3B Exports |
| 4 | Sales | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | Sales | Sales | sales | client, state, start_month, end_month | Sales related files |
| 5 | SalesReco | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | Sales Reco | Sales | sales | client, state, period | Sales related files |
| 6 | AnnualReport | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$` | Annual Report | Annual | root | client, state, year | the Version folder itself |

Rules that follow from the regexes (READ, confirmed by RAN t1 on 66 invented names):
- **Hyphen is the separator.** Client and state are `[^-]+` (no hyphen allowed). A hyphen inside a client name breaks recognition, or silently mis-splits it:
  - `GSTR3B-Client-A-...` is not recognised (pattern 3 allows exactly 3 groups).
  - `GSTR-2B-Reco-Client-A-Maharashtra-Apr24.xlsx` **is recognised** with client=`Client`, state=`A`, period=`Maharashtra-Apr24`, because the period group `(.+)` absorbs extra hyphens. The only sign is a warning "State name issue: Client name too short", and it is not shown in the UI (RAN t1).
  - The same absorption applies to SalesReco and AnnualReport. `AnnualReport-ClientA-Maharashtra-FY-2024-25.xlsx` gave year=`FY-2024-25`.
- **The last group** differs by pattern:
  - GSTR3B and Sales use `[^-]+`: no hyphen in month or start/end month. `Sales-ClientA-Maharashtra-Apr24.xlsx` (only one period part) and `Sales-...-Apr-May-Jun.xlsx` are not recognised.
  - IMS needs **exactly 8 digits**: `30-04-2024`, `Apr24`, 7 digits or 9 digits all fail.
- **Ties and look-alikes** (RAN t1):
  - `SalesReco-...` never matches pattern 4, because pattern 4 needs `Sales-` with a hyphen.
  - `Sales-Reco-ClientA-Maharashtra-Q1.xlsx` **matches pattern 4 (Sales)** with client=`Reco`, state=`ClientA` (state code fallback `CLI`), start_month=`Maharashtra`, end_month=`Q1`. It is filed silently under a wrong "client".
  - `Sales Reco-...` (with a space), `GSTR2B-Reco-...`, `GSTR-3B-...`, `GSTR3B_..._...` (underscores), `Annual Report-...` (space) and `Ims-Reco-...` are not recognised.
  - `IMSReco-...` and `gstr-2b-reco-...XLSX` are recognised, because matching is case-insensitive.
- **Duplicate-download markers** (RAN t1):
  - Only GSTR3B accepts a leading counter such as `(1) ` or `(12)`, via `(?:\(\d+\)\s*)?`.
  - A trailing ` (1)` or `(1)` before the extension is **accepted by every pattern except ImsReco**, because it becomes part of the last group: month `Apr24 (1)`, period `Apr24 (1)`, end_month `Jun (1)`, year `2024 - Copy`.
  - `ImsReco-...-30042024 (1).xlsx` and `(1) GSTR-2B-Reco-...` are **not** recognised.
  - So browser-duplicated GSTR-3B downloads count as extra months: three GSTR-3B files for one client in RAN t2. Duplicated 2B, Sales, Sales Reco or Annual files raise "Multiple ..." (see 2.3).
- **Spaces.** Spaces around the hyphens (`GSTR3B - ClientA - ...`) are not recognised. A leading space before `GSTR3B` is not recognised. Spaces inside client or state are fine, and client/state are `.strip()`ped (file_parser.py:98-100).
- **Client and state validation** (file_parser.py:104-115, using utils/helpers.py:321-342 for **both** client and state):
  - Empty, shorter than 2 characters, longer than 100 characters, or containing `< > : " / \ | ? *` gives a **warning only**. The file is still used, and warnings are not shown anywhere in the B UI (READ, HIGH).
  - RAN t1: client `A`, state `M` and a 120-character client were all accepted with warnings.

### 2.3 How a recognised file becomes a "client" (core/file_parser.py:258-314)
- **Client key** = `f"{client}-{state_code}"` (file_parser.py:261-262), using the raw client text from the file name. [GST RULE] state-code logic (see 2.5).
  - The key is **case-sensitive**. RAN t2: `GSTR3B-ClientB-...` and `GSTR3B-clientb-...` became two separate clients, `ClientB-KA` and `clientb-KA`.
  - On Windows those two would get Level-2 folders `ClientB-KA` and `clientb-KA`, which **are the same folder** on a case-insensitive disk, so their copies and reports would land together. Their report files `ITC_Report_ClientB_...` and `ITC_Report_clientb_...` are the same file on Windows, so one overwrites the other (MEDIUM; Windows-only reasoning, not run).
- **State spelling.** Full name and code both map to the same key: `Karnataka` and `KA` both give `-KA` (RAN t2). The displayed `state` is whichever file was processed **last** (file_parser.py:267). RAN t2 showed `karnataka` (lower case) for ClientB.
- **"Complete" status** [GST RULE] (which documents a client must have), TIME-VARYING:
  - Expected types are `EXPECTED_FILE_TYPES` = GSTR-2B Reco, IMS Reco, GSTR-3B Export, Sales, Sales Reco, Annual Report (utils/constants.py:68-75).
  - Status values:
    - "Complete": none missing and no "Multiple ...".
    - "Missing N files": at least one missing.
    - "Has duplicates": nothing missing, but some type has more than one file.
  - **GSTR-3B Export may have any number of files** ("monthly filings - unlimited", file_parser.py:294-296). Any other type with 2 or more files gives `Multiple <type> (<n> files)`.
  - Completeness % = found types / 6 × 100, rounded to 1 decimal (file_parser.py:302-303).
- `validate_results()` (file_parser.py:366-399) treats Annual Report and GSTR-3B Export as "critical" [GST RULE]. It is **never called** (grep), so this rule is dead.

### 2.4 What happens to an unrecognised file
- It goes into `variations` with `issue` and a `pattern_expected` hint from `_suggest_pattern` (file_parser.py:316-334). The hint uses lower-case keyword tests in this order:
  1. contains `gstr` and `2b`: the 2B pattern;
  2. else `ims`: the IMS pattern;
  3. else `gstr` and `3b`: the 3B pattern;
  4. else `sales` and `reco`: the SalesReco pattern;
  5. else `sales`: the Sales pattern;
  6. else `annual`: the AnnualReport pattern;
  7. else "Check file naming convention".
- **Where the user sees it in B: only the count.** The setup scan log shows `Variations: N` (pyqt_main_window.py:261-262) and the validation summary shows the count (pyqt_main_window.py:294). The file names and issues are **never displayed**.
- The summary workbook's "Variations" sheet is **always "No file variations found"**, because processing_controller.py:73 sets `'variations': []` and nothing fills it (grep). This is a regression from v3: 47a1492:gui/handlers/processing_handler.py:161 passed the real list. The regression came in c258047 (READ, HIGH).
- An unrecognised file is not copied anywhere (READ, HIGH).

### 2.5 State names to state codes [GST RULE] state-code logic, TIME-VARYING (states and union territories change)
- `STATE_CODE_MAPPING` (utils/constants.py:234-275) maps lower-case full names to **two-letter alphabetic codes**, e.g. `'maharashtra': 'MH'`, `'telangana': 'TS'`, `'uttarakhand': 'UK'`, `'odisha'/'orissa': 'OD'`, `'dadra and nagar haveli and daman and diu': 'DD'`, `'delhi'/'national capital territory of delhi': 'DL'`, `'puducherry'/'pondicherry': 'PY'`, `'ladakh': 'LA'`. There are 38 entries. These are **not** the two-digit numeric GST state codes; described only, no judgement.
- `get_state_code` (utils/helpers.py:485-511) works like this:
  - It strips and lower-cases the name, then looks it up.
  - If not found and the name has 2 or more words, it uses the first letters of the first two words, upper-cased.
  - If one word, it uses the first 3 characters of the original text, upper-cased.
- RAN t1/t3:
  - `MH` gives `MH` (fallback);
  - `Dadra and Nagar Haveli` and `Daman and Diu` **both** give `DA`;
  - `Other Territory` gives `OT`; `New Delhi` gives `ND`; `Delhi NCR` gives `DN`;
  - `Karnatak` (misspelt) gives `KAR`; `Westbengal` gives `WES`.
  - Two different states can therefore collide on one key.
- The code is used in the client key, the Level-2 folder name, and the report file names (sections 2.3 and 3).

---

## 3. Folder layout read and created

### 3.1 Read
- **Source folder**: a single flat folder chosen in Setup (setup_view.py:605-608). It must exist (pyqt_main_window.py:219-225). Only top-level `*.xlsx|*.xls|*.xlsm` files are read (2.1).
- **Target folder**:
  - Re-Run and Resume list it for sub-folders whose name starts with `Annual Statement-` and take the one with the newest **modification time** (`st_mtime`, core/file_organizer.py:158-169).
  - The Extract tab lists the same sub-folders and takes the one with the **highest name** (`sorted(..., key=name, reverse=True)`, pyqt_main_window.py:431-448).
  - **These disagree.** Names are `Annual Statement-DDMMYY HHMM`, and sorting by DDMMYY is not date order. RAN t5: with `Annual Statement-311225 0900` (Dec 2025) and `Annual Statement-020126 0900` (Jan 2026):
    - Re-Run/Resume used the Jan-2026 folder;
    - the Extract tab showed the Dec-2025 folder.
  - Version folders in the Extract tree are ordered the same way, by name descending (pyqt_main_window.py:457) (HIGH).
- **Extract tab, inside the chosen Level-1 folder** (pyqt_main_window.py:450-504). For each sub-folder (Level 2), then each sub-folder of that (Version), it keeps only Version folders containing `ITC_Report_*.xlsx` or `Sales_Report_*.xlsx`.
  - "Refreshed" status = any file matching `*_Refreshed*` with extension `.xlsx|.xls|.xlsm` that does not start with `~$`. The newest file's modification time is shown as `DD/MM/YY HH:MM` (pyqt_main_window.py:467-479).
  - Folder names are **not** checked against the naming rules. Any folder with a report inside is shown.

### 3.2 Created (READ core/file_organizer.py and utils/constants.py:81-90; RAN t4/t5 tree on invented data)
```
<Target folder>/                                         (created if missing: file_organizer.py:61-65)
└── Annual Statement-DDMMYY HHMM/                        Level 1 (constants.py:82; timestamp helpers.py:23-25)
    ├── GST_Processing_Summary_DDMMYY_HHMM.xlsx          (processing_controller.py:184-193; only if ≥1 client succeeded)
    ├── PQ_Extraction_Report_YYYYMMDD_HHMM.xlsx          (Extract step; data_consolidator.py:163-167)
    ├── Extraction_Log_YYYYMMDD_HHMM.txt                 (Extract step; extraction_controller.py:109-121)
    └── <Client>-<SC>/                                   Level 2 = create_client_state_key (helpers.py:513-537)
        └── Version-DDMMYY HHMM/                         (constants.py:84)
            ├── AnnualReport-...xlsx                     (Annual Report goes here, file_organizer.py:294)
            ├── ITC_Report_<sc>_<SC>_DDMMYY_HHMM.xlsx    (processing_controller.py:117-122, 220)
            ├── Sales_Report_<sc>_<SC>_DDMMYY_HHMM.xlsx
            ├── ITC_Report_..._Refreshed_DDMMYY_HHMM.xlsx  (Extract step copy; report_processor.py:92-94)
            ├── _Organization_Report_<client>_<state>.txt  (file_organizer.py:313-355)
            ├── GSTR-3B Exports[ (<sc>)]/                (constants.py:86)
            ├── Other ITC related files[ (<sc>)]/        (constants.py:87) - GSTR-2B Reco + IMS Reco
            └── Sales related files[ (<sc>)]/            (constants.py:88) - Sales + Sales Reco
```
Exact name rules:
- `DDMMYY HHMM` = `datetime.now().strftime("%d%m%y %H%M")` (helpers.py:25). It is taken **once per run**: `FileOrganizer.timestamp` (file_organizer.py:43) for the folders, and separately `ProcessingWorker._process` (processing_controller.py:44) for the report and summary names. Two separate `now()` calls a moment apart could fall in different minutes (LOW).
  - **Same minute = same folders.** `mkdir(exist_ok=True)` reuses them. RAN t4: two runs in the same minute wrote into the same Level-1 and Version folders.
- **Level 2 `<Client>-<SC>`** = `create_client_state_key(client, state, max_length)` (helpers.py:513-537):
  - It replaces the whole words `Private` with `Pvt` and `Limited` with `Ltd`, **case-insensitively**, then appends `-` and the state code.
  - If longer than `client_name_max_length` (Setup spin box, 15-100, default 35, setup_view.py:374-377), the client part is cut so the whole key fits.
  - **No other cleaning**: commas, `&`, brackets and dots stay. RAN t3: `Client & Sons (India)-MH`, `Client A LLP, Unit 2-MH`.
  - The spin box is greyed out unless "Include client name in brackets" is ticked (setup_view.py:377, 382), **but the value is used for Level 2 in every case** (file_organizer.py:79-83) (READ, HIGH).
- **`<sc>`** (category-folder bracket and report file name) = `sanitize_filename(client)` (helpers.py:66-112):
  - It replaces `Private` with `Pvt` and `Limited` with `Ltd` **case-sensitively and anywhere in the text**, not as whole words.
  - It changes control characters and ``<>:"/\|?*[]{}+=!@#$%^,;'"`~`` to `_`; collapses runs of spaces and underscores to one `_`; strips `_`, `.` and spaces from the ends; and limits the length to 200 minus 10 minus the extension length.
  - It treats the client name as a file name, so text after the last dot is treated as an "extension". RAN t3 (Python 3.14): `Client A Co. Ltd` became `Client_A_Co. Ltd` (space kept, dot kept); `Client A private limited` became `Client_A_private_limited` (no abbreviation, while Level 2 gives `Client A Pvt Ltd-MH`).
  - So Level 2, the category bracket, the report name and the organisation-report name can show **different spellings** of the same client (READ+RAN, HIGH).
- **Category folders:** the ` (<sc>)` suffix is added only if the global "Include client name..." box is ticked, or if the per-client setting says so (file_organizer.py:113-135).
  - **The per-client setting can never be set in B.** `ProcessingState.client_folder_settings` (models/app_state.py:83) is always `{}`: nothing writes it (grep). The v3 per-client toggle lived in 47a1492:gui/handlers/client_handler.py:175-219, which was deleted in a62d3cc.
  - The Setup note "Names >10 chars will show warning..." (setup_view.py:385) refers to that lost feature (READ, HIGH).
- **Input copies** are named `sanitize_filename(original name)`, e.g. `(1) GSTR3B-...` became `(1)_GSTR3B-...` and spaces became `_` (RAN t3). The copy uses `shutil.copy2`, which keeps the source's timestamps, and checks that the sizes match (helpers.py:142-185). Sources are never changed or moved (READ, HIGH).
- **Existing destination file** (file_organizer.py:259-270):
  - Fresh: a backup `<stem>_backup_DDMMYY_HHMM<ext>` is made next to it (helpers.py:187-204), then the file is overwritten.
  - Re-Run: overwritten, no backup.
  - Resume: skipped ("File already exists").
  - Report files (template copies) are always overwritten without backup (excel_handler.py:124/228).
  - Backups are written into the same category folders. If a template query reads a whole folder (B2), backups and `(1)` duplicates would be read too (MEDIUM; depends on the template; question for the author).
- **Re-Run / Resume behaviour versus the on-screen text:**
  - The Setup card says "Re-Run: Overwrite existing" and "Resume: Skip already processed" (setup_view.py:510-512).
  - What the code does: both re-use the latest Level-1 folder by mtime but always create a **new `Version-<now>` folder** (file_organizer.py:98-104). RAN t5: Resume in a new minute copied every file again ("Success", nothing skipped).
  - Skipping or overwriting only happens within the same minute (READ+RAN, HIGH).
  - With no `Annual Statement-*` folder, Re-Run/Resume create a new one (file_organizer.py:170-179).
- **Organisation report** `_Organization_Report_<client>_<state>.txt` uses the **raw** client and **raw full state name** (file_organizer.py:313). These are not sanitised, but they came from a valid Windows file name, so they hold no forbidden characters (MEDIUM).
  - Its "Successfully Copied" and "Folders Created" figures are **cumulative for the whole run**, not per client: one organiser serves all clients (processing_controller.py:46-52; file_organizer.py:346-351).
  - RAN t5: the second client's report said "Total Files: 3 / Successfully Copied: 11 / Folders Created: 12".
- **Summary workbook** (excel_handler.py:308-346) has the sheets Summary, Client Status, File Mapping, Errors and Variations (Variations is always empty, see 2.4).
- **Extract outputs**:
  - Refreshed copy: `<report stem><suffix>.xlsx` in the same Version folder, where suffix = the user's "File suffix" text with `{timestamp}` replaced by `DDMMYY_HHMM` (report_processor.py:92-94; default `_Refreshed_{timestamp}`, extraction_view.py:136).
    - An empty suffix makes the copy the same file as the original: shutil raises "are the same file" and the report fails (RAN t8).
    - A suffix without `{timestamp}` gives a fixed name, overwritten every time (RAN t8).
  - Consolidated report and log in the Level-1 folder: `PQ_Extraction_Report_%Y%m%d_%H%M.xlsx` and `Extraction_Log_%Y%m%d_%H%M.txt` (note the different date order from the folders). The same minute overwrites.
  - Until b0e7d0e the consolidated report went to the target folder itself (c258047:extraction_controller); b0e7d0e moved it to the Level-1 folder and added the log file.
- `cleanup_on_error` (file_organizer.py:364-379) would delete an empty Version folder in Fresh mode. It is **never called** (grep).
- **Path lengths** (RAN, computed for an invented 44-character target path and a 40-character client name):
  - GSTR-2B copy 233 characters (268 with "Include client name");
  - ITC report 195;
  - refreshed ITC report 217.
  - The bracketed category name uses the full `<sc>` and is **not** limited by the 35-character setting. See 5.4.

---

## 4. The app's own files: settings, cache, run history, logs, theme, temp

### 4.1 Inventory (each file traced both ways)

| # | File (exact path expression) | Relative to | Format / keys | WRITTEN by | READ by | Agree? |
|---|---|---|---|---|---|---|
| S1 | `Path.home() / '.gst_organizer' / 'settings.json'` (models/cache_manager.py:17, 32-35) | user home folder; on Windows `Path.home()` = `%USERPROFILE%` (Python ≥3.8 ignores HOME on Windows; HIGH) | JSON, indent 2, UTF-8, `default=str`. 11 keys from `AppSettings.to_dict` (app_state.py:25-39): `source_folder`, `itc_template`, `sales_template`, `target_folder` (str); `processing_mode` ('fresh'/'rerun'/'resume'); `include_client_name_in_folders` (bool); `client_name_max_length` (int); `dark_mode` (bool); `pq_wait_time` (int); `pq_suffix_pattern` (str); `pq_skip_refresh` (bool) | `CacheManager.save_settings` (cache_manager.py:41-55), called by `_save_settings` (pyqt_main_window.py:210-214). That is called by: Scan (227), processing completed (381), dark-mode toggle (597), File > Save Settings (101-103), window close (603), and old-cache migration (cache_manager.py:119) | `load_settings` (cache_manager.py:57-62), then `AppSettings.from_dict` (app_state.py:41-46), at startup (pyqt_main_window.py:50-51), then put into the views (174-187) | **Yes** (same keys; unknown keys dropped; missing keys default). RAN t6 round trip OK. Types are not validated: a hand-edited wrong type would reach `QSpinBox.setValue` (LOW). Values outside the spin ranges are clamped at save (RAN t6: 5 saved as 15, 500 as 120). Settings are **not** saved when Start is pressed, only when processing completes (pyqt_main_window.py:352-363 versus 381) |
| S2 | `~/.gst_organizer/session_state.json` (cache_manager.py:37) | home | JSON. Keys: `recent_source_folders` (list, newest first, max 5), `active_tab` (int), `window_size` ([w, h]) | `save_settings` (recent list, only when the source folder is new, cache_manager.py:46-52); `save_session_state` on close (pyqt_main_window.py:604-607) and in migration (cache_manager.py:124) | **Nobody.** `load_session_state` (82-84) and `get_recent_source_folders` (86-89) are never called (grep). Only `save_settings` re-reads the recent list to update it | **Write-only** (dead state). RAN t6: the file holds `active_tab` and `window_size`, which are never restored |
| S3 | `~/.gst_organizer/run_history.json` (cache_manager.py:36) | home | JSON. `runs` (list of {timestamp, total_clients, successful_clients, failed_clients, total_files, reports_generated, processing_mode}, **trimmed to 50**); `client_history` ({key: [ClientRunRecord dicts]}, **never trimmed**); `last_run_timestamp`; `total_runs` (run_history.py:44-71, 109-116) | `save_run_history`, after processing (pyqt_main_window.py:378-380) and after extraction (555-558) | `RunHistory.load` at startup (pyqt_main_window.py:43). Only use: the Validation tab's "Last Run" column = `client_history[client_key][-1]['run_timestamp']` (validation_view.py:381-385; run_history.py:97-107) | **Partly.** Processing writes under the parser key `<client>-<SC>`, and Validation reads the same key: agree. **Extraction writes under `result['client']`**, which is `"<Level-2 folder name> / <Version folder name>"` (extraction_view.py:384; report_processor.py:44; pyqt_main_window.py:555-557). No reader uses that key, so extraction history is **orphaned**. RAN t6: `client_history` keys became `['ClientA-MH', 'ClientA-MH / Version-020126 0900']`. Even without the "/ Version" part, the Level-2 name (Pvt/Ltd, truncated) can differ from the parser key. Fields `itc_report_path`, `sales_report_path` and `errors` (run_history.py:27-28, 33) are never filled |
| S4 | `~/.gst_organizer/.migrated` (cache_manager.py:99) | home | empty marker file | `migrate_from_old_cache` (127) | same function (101) | Yes |
| S5 | `~/.gst_organizer_cache.json` (cache_manager.py:98), the **old v3 cache** | home | JSON written by 47a1492:gui/utils/cache_manager.py:29-47 with keys `source_folder`, `itc_template`, `sales_template`, `target_folder`, `processing_mode`, `include_client_name`, `dark_mode`, `client_name_max_length`, **`recent_folders`** (47a1492:gui/main_window.py:328-340) | not written by B | `migrate_from_old_cache` (cache_manager.py:93-131), once, at startup (pyqt_main_window.py:47) | **Mismatch: B reads `recent_source_folders` (cache_manager.py:122), but v3 wrote `recent_folders`**, so the old recent-folder list is never migrated. RAN t6: an invented old cache with 2 recent folders gave `recent_source_folders: ['/x/src']`, the current folder only, added by `save_settings`. Other keys map correctly (`include_client_name` to `include_client_name_in_folders`). The old file was written ASCII-only (`json.dump` default), so reading it as UTF-8 is safe (HIGH) |
| S6 | `gst_organizer.log` (main.py:20) | **the current working directory** (relative path), not the app folder | text log, root logger, INFO+, format `%(asctime)s - %(name)s - %(levelname)s - %(message)s`, append mode, no rotation, default (locale) encoding | `logging.basicConfig` in main.py:16-23. Every module's logger propagates to it | nothing in code (README.md:289-291, 333 tells the user to look at it) | n/a. RAN: started from another folder, main.py created `gst_organizer.log` in that folder. **Not created when started via launch_extractor.py**, which has no `basicConfig`, so there is no log file at all and only warnings go to stderr (READ, HIGH). Log messages contain symbols such as `✓ ⚠️ ❌ 📌`; with a non-UTF-8 Windows code page each such line could fail to write ("--- Logging error ---" to stderr) (MEDIUM) |
| S7 | `pq_extractor.log` (power_query_extractor/extractor_main.py:16) | cwd | text | would be `basicConfig` in extractor_main.py | — | **Never created.** extractor_main.py imports `.gui.extractor_window` (line 9), which a62d3cc deleted, so the import fails before line 12. RAN: `ModuleNotFoundError: No module named 'power_query_extractor.gui.extractor_window'`. The whole file is dead; its cwd-relative cache readers (`gst_organizer_cache.json`, `.cache/app_state.json`, `config/last_session.json`, extractor_main.py:38-58) never run. Even in v3 these never matched the real cache, which was in home, not cwd |
| S8 | `<Level-1>/Extraction_Log_%Y%m%d_%H%M.txt` (extraction_controller.py:109-121) | data folder | UTF-8 text: header plus every Extract-tab log line | `ExtractionWorker._save_log_file` | nobody | n/a |
| S9 | `Path(tempfile.gettempdir()) / "gst_excel_temp"` (excel_handler.py:52-53) | `%TEMP%` (per user) | empty folder | created at every `ExcelHandler()` (once per processing run) | **never used**, never deleted (grep `temp_dir`) | dead. RAN t4: created under the redirected TMPDIR |
| S10 | Theme | inside S1 | `dark_mode` bool | `_toggle_dark_mode`, then `_save_settings` (pyqt_main_window.py:594-597) | startup `apply_theme(self.settings.dark_mode)` (66) and the menu tick (116) | Yes (RAN t6: migrated `dark_mode: True` ticked the menu). The stylesheets are Python strings in gui/theme/theme_manager.py:25-405, not files |
| S11 | Shared in-memory state `EXTRACTION_OPTIONS['skip_refresh']` | module-level dict in cell_mappings.py:248-253 | bool | `ExtractionWorker._process` overwrites it each run (extraction_controller.py:54-55) | report_processor.py:97-98 | Agree (same object). It is a mutable global, so the last run's value persists in the process until the next run (LOW). `auto_refresh_if_missing` and `show_timestamp` are never read (grep) |
| S12 | `ProcessingState.client_folder_settings` | in memory | {client_key: bool} | **nobody** | FileOrganizer (file_organizer.py:126) | **No write path.** Always empty (see 3.2) |

The cache folder `~/.gst_organizer` is created every time the window opens (cache_manager.py:33), even before the user does anything (READ, HIGH).

### 4.2 Shared on-disk state between the app's own steps (data folder)

| What | Writer | Reader(s) | Agree? |
|---|---|---|---|
| Latest Level-1 folder | FileOrganizer: Fresh creates `Annual Statement-<now>`; Re-Run/Resume pick the newest **mtime** (file_organizer.py:146-179) | Extract tab picks the highest **name** (pyqt_main_window.py:431-448) | **Mismatch** (RAN t5) |
| Report files | `ITC_Report_*.xlsx`, `Sales_Report_*.xlsx` (processing_controller.py:117-130, 220) | Extract tree (`ITC_Report_*.xlsx` / `Sales_Report_*.xlsx`, pyqt_main_window.py:461-462); `ReportProcessor._process_report` (`f"{report_type}_Report_*.xlsx"`, takes `report_files[0]`, report_processor.py:76-89) | Prefix agrees, **but the glob also matches the refreshed copies** (they keep the prefix). Whichever file the folder listing returns first is used as the "original". RAN t8 on Linux: a previous `_Refreshed_` copy was picked and copied to `..._Refreshed_<a>_Refreshed_<b>.xlsx`. On Windows NTFS (Windows' disk format), alphabetical listing would normally put the original first, because `.` sorts before `_` (MEDIUM; listing order on network shares may differ) |
| Refreshed copies | `stem + <user suffix with {timestamp}> + .xlsx` (report_processor.py:92-94) | "Skip refresh": `f"{stem}_Refreshed_*.xlsx"` (report_processor.py:102); Extract tree status: `*_Refreshed*` (pyqt_main_window.py:467-471) | **Mismatch when the suffix is changed.** Both readers hard-code `_Refreshed`. RAN t8: with suffix `_R{timestamp}`, "skip refresh" did not find the earlier copy and refreshed again, and the tree shows "Pending". The tree glob also counts unrelated files such as `old_Refreshed_copy.xlsx` (RAN t8). Status is per Version folder, not per report: one refreshed ITC marks the whole version "Refreshed" (READ) |
| Links cells | excel_handler.py (section 1.4A) | the template's queries (outside the repo) | cannot verify |
| Info!BB2 | the template's queries (outside the repo) | report_processor.py:589-636 | cannot verify |
| Extracted cells | report_processor.py:478-564 (dict keyed by `<OutputSheet>_<Column>` with spaces as `_`, **and** by bare cell reference) | data_consolidator.py:113-142 (substring match, then exact key) | **Mismatch for single-word columns** (RAN t7; section 1.4C) |
| Variations list | file_parser.py (2.4) | summary workbook "Variations" sheet | **Broken link**: always `[]` (processing_controller.py:73) |

### 4.3 Confidentiality note on the app's own files
- `run_history.json` (S3) stores, per extraction, the **whole extraction result**: client folder name, refreshed file paths, and every extracted value (tax figures). It is stored in the user's profile folder in plain text (pyqt_main_window.py:555-558 into run_history.py:73-95).
- `gst_organizer.log` (S6) and `Extraction_Log_*.txt` (S8) also contain client names, full paths and extracted figures (report_processor.py:541, 557-562). If S6 lands in a shared folder (see 6), other users can read them (READ, HIGH).

---

## 5. Environment

### 5.1 Interpreter, packages, start-up
- `requirements.txt`: `PyQt5>=5.15.0`, `openpyxl>=3.1.2`, `lxml>=4.9.0` ("optional"). **`pywin32` is commented out** (`# pywin32>=306`, line 14). There are no exact versions. It says "Python 3.7 or higher".
- **pywin32 is in fact mandatory for the app to start.**
  - The chain is: power_query_extractor/core/report_processor.py:11-15 imports `pythoncom`, `win32com.client`, `win32api`, `win32con` and `win32gui` **unconditionally**; gui/controllers/extraction_controller.py:10 imports report_processor; gui/pyqt_main_window.py:27 imports extraction_controller; main.py:38 imports the main window. gui/__init__.py:7 also imports it.
  - RAN: `import gui.pyqt_main_window` gives `ModuleNotFoundError: No module named 'pythoncom'`. `python main.py` printed "Missing module: No module named 'pythoncom' / Please install required packages: pip install PyQt5 openpyxl" (a misleading hint) and exited 1.
  - **Inconsistent import pattern**: excel_handler.py:20-28 guards the same imports with a platform check and try/except, and helpers.py:476-483 guards `win32api`; report_processor does not (HIGH).
  - The README (README.md:107) lists `pywin32>=305; sys_platform == 'win32'` in its "Requirements.txt Contents" block, but the actual file comments it out (README vs file mismatch).
- **Entry points:**
  - `main.py`: sets the UTF-8 stdout wrapper if a console exists (12-13), configures logging (16-23), inserts the project folder into `sys.path` (26-29), opens the window.
  - `launch_extractor.py`: appends its folder to `sys.path` (8), opens the same window on tab index 3 (18). It has **no logging config**.
  - `power_query_extractor/extractor_main.py` is dead (S7).
- **Errors are reported only with `print()` to the console** (main.py:49-59; launch_extractor.py:22-24). Started with `pythonw.exe` (no console, as Standard 13 requires), an import error or crash at start-up shows **nothing**. There is no `sys.excepthook` and no message box (READ, HIGH).
- **Finding the app's own files:**
  - The Python modules are found through `sys.path` set from `Path(__file__).parent` (main.py:28). That does not depend on the cwd.
  - The **log** does depend on the cwd (S6). No other own-file path depends on the cwd.
- Fonts: "Segoe UI" and "Consolas", hard-coded in stylesheets (theme_manager.py:30, 165; views). These are Windows fonts; other systems use fallbacks (LOW importance).

### 5.2 Windows-only calls (all READ; cannot run here)

| Call | Where | Guarded? | Notes |
|---|---|---|---|
| `win32com.client.GetActiveObject("Excel.Application")`, else `Dispatch(...)` | excel_handler.py:110-115 | platform + try (20-28) | **Attaches to the user's already-open Excel** if there is one. It then sets `Visible=False`, `ScreenUpdating=False` and `DisplayAlerts=False` (118-120) on it, so the user's own Excel windows are hidden during report creation. In `finally` it sets `DisplayAlerts=True` (205) and, because `self.excel_app` is always `None` (55; never assigned), **calls `Quit()` on that Excel** (207-211). That closes the user's own Excel, with alerts back on, so "save changes?" prompts may appear in a hidden Excel (HIGH on the sequence; the effect is MEDIUM) |
| `Workbooks.Open(path)`; on error, if the message has `-2147352567` or `path` or the path is longer than 218, it retries with `win32api.GetShortPathName` (8.3) | excel_handler.py:264-306; helpers.py:473-483 | yes | 218 = the code's assumed Excel path limit. 8.3 names work only if the disk has 8.3 name creation on (MEDIUM) |
| `pythoncom.CoInitialize/CoUninitialize` | excel_handler.py:107, 215; report_processor.py:171, 363, 484, 584 | report_processor: no | Called in QThread workers (correct place). Excel objects are still held when `CoUninitialize` runs. Standard 12 asks for the Excel work in a separate function (MEDIUM) |
| `win32com.client.Dispatch("Excel.Application")` for the refresh, with `Visible=True` and `WindowState=-4137` (maximised) | report_processor.py:174-191 | no | `Dispatch` may attach to an already-running Excel instead of starting a new one (MEDIUM; Standard 12 prescribes `DispatchEx`). `Quit()` in `finally` (355-360) would then close it |
| `WScript.Shell` `AppActivate(excel.Caption)`, then `SendKeys("%a")`, `SendKeys("r")`, `SendKeys("a")` | report_processor.py:216-249 | no | Needs an unlocked, interactive desktop with keyboard focus on Excel. A click in another window during the wait sends the keys elsewhere. It depends on the Excel UI language and ribbon key tips (English: Data = A, Refresh All = R then A) (MEDIUM). Fixed waits: 5 s load; the user wait time (Setup range 5-120 s, extraction_view.py:126-128); 5 s of error checks; `max(10, wait)` s extended wait; plus 2+1+3 s (report_processor.py:207-319, 141). That is about 2×wait + 20 s per report |
| `win32gui.EnumWindows` + `GetClassName` in `['#32770','NUIDialog','bosa_sdm_msword']` + title containing any of `microsoft excel, error, problem, cannot, failed, unable, could not, exception, invalid, missing` | report_processor.py:368-404 | no | Looks at **every top-level window on the desktop**, not only this Excel's. A matching dialog from any program can be "closed": the code clicks a child button whose text contains `ok`, `close`, `yes` or `continue` (BM_CLICK), else posts WM_CLOSE (report_processor.py:443-476) (HIGH that it scans all windows; MEDIUM that it hits other programs in practice) |
| Long-path handling: if longer than 218 characters, forward slashes become backslashes and the `\\?\` or `\\?\UNC\` prefix is added; on failure it closes same-named workbooks and retries; then the 8.3 path; then without the prefix; then `Open(Filename=..., UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, Notify=False)` | report_processor.py:638-747 | inner try only | Whether Excel accepts `\\?\` paths is not known here. LOW that it helps; check on a PC with a 230-character path. The "read-only" log message at 492-494 is not true: `_safe_open_workbook` opens read-write and closes without saving |
| `os.startfile(level1)` (opens Explorer on the Level-1 folder after processing) | pyqt_main_window.py:404-410 | `platform.system()=='Windows'` | — |

### 5.3 Environment variables, registry, network
- **Environment variables read directly: none** (grep `environ|getenv` gives nothing).
- Indirect use:
  - `USERPROFILE` (Windows) or `HOME` (other systems) through `Path.home()` (cache_manager.py:17, 98);
  - `TEMP`/`TMP` through `tempfile.gettempdir()` (excel_handler.py:52);
  - Qt's own variables.
- **Registry: none** (grep `winreg|registry` gives nothing). Indirectly, Windows long-path support (the `LongPathsEnabled` registry setting) decides whether `shutil.copy2` and `mkdir` succeed for paths over 260 characters. Excel has its own lower limit regardless (MEDIUM).
- **UNC and network drives:** no special handling, except report_processor's `\\?\UNC\` prefix for paths longer than 218 characters (report_processor.py:676-677). Links cells get whatever path form the target has, mapped drive letter or `\\server\share`, through `clean_windows_path` (helpers.py:311-315). Folder selection relies on `st_mtime` (file_organizer.py:167), which works on network drives (MEDIUM). There is **no locking**: two users processing the same target in the same minute write into the same folders and files (READ, HIGH).
- **OneDrive:** not referenced in the code, except a hint string at report_processor.py:741.

### 5.4 Path-length limits
- The code assumes a 218-character Excel limit (excel_handler.py:293; report_processor.py:668, 743). The README says "Windows path limit of 260 characters applies" (README.md:317).
- No code shortens paths beforehand. The only limit is the 35-character (15-100) Level-2 key; category-bracket names and copied file names use up to about 190 characters each (helpers.py:104-106).
- RAN example (invented 44-character target, 40-character client): copies reach 233 to 268 characters. That is above 218 and, with "Include client name", above 260.

### 5.5 Excel version
- No version check in code. The README says "Microsoft Excel 2016+ for Power Query features (Windows only)" (README.md:55).
- The refresh relies on the ribbon key-tip sequence (5.2). Extraction relies on COM `Range().Value`. Templates are assumed to hold Power Query queries (Excel 2016+ has Power Query built in; MEDIUM).

### 5.6 Current working directory
- It matters only for `gst_organizer.log` (S6) and for the dead extractor_main.py readers (S7).
- The Qt file dialogs start in the default folder (empty string passed, setup_view.py:606, 612, 619, 625). They do not remember the last folder except through the stored text.

---

## 6. Storage locations compared with the author's Standard 13 (rebuild_pipeline_FINAL.md:900-1061)

| B location | Standard 13 rule | Comparison |
|---|---|---|
| Settings, session and run history in `%USERPROFILE%\.gst_organizer\` (S1-S4) | "Settings and logs go in the user's own folder (`%LOCALAPPDATA%\[App Name]\`), never in the shared app folder" (lines 1039-1040; LOG_DIR at 997) | **Per-user: matches the intent.** **Base folder and name differ**: profile root plus a dot-folder instead of `%LOCALAPPDATA%\<App Name>\`. In a roaming-profile domain the profile root may roam, while AppData\Local does not (MEDIUM). run_history.json holds client figures (4.3); the Standard has no specific rule on that beyond per-user storage |
| `gst_organizer.log` in the **cwd** (S6) | Logs in `%LOCALAPPDATA%\[App Name]\`; "Find the app's own files from APP_DIR, never from the current folder" (1034); launcher `start "" /D "%~dp0"` makes the cwd **the app folder in the shared share** (962); "Never keep … any file several people write at the same time, in the shared folder" (1046-1047) | **Conflicts.** Under the Standard 13 launcher, every user would append to one `gst_organizer.log` inside `\\server\share\Apps\<App>\`: logs in the shared folder, written by several people at once, and holding client names and figures (READ+RAN that the log follows the cwd; HIGH) |
| `pq_extractor.log` (S7) | same | Never created (dead code) |
| `Extraction_Log_*.txt` in the Level-1 data folder (S8) | logs in `%LOCALAPPDATA%` | **Differs**: a log in the client data folder. It is arguably an output record, not an app log; the author decides |
| Output workbooks and text reports in the target folder (3.2) | The Standard has no rule on output data location | Outside Standard 13's scope; no lock against two users writing (5.3) |
| `%TEMP%\gst_excel_temp` (S9) | — | Per user; unused |
| Old `~/.gst_organizer_cache.json` (S5) | — | Read-only migration source |
| No error log file, no `sys.excepthook`, no `threading.excepthook`, no message box; errors go to `print()` (5.1) | `report_fatal_error` writes `%LOCALAPPDATA%\<App>\logs\error_*.txt` and shows a MessageBoxW (995-1025); "There is no console: print() shows nothing" (1035) | **Missing.** Under `pythonw.exe` a start-up ImportError, for example missing pywin32 (5.1), would leave the user with nothing on screen |
| `sys.path.insert(0, project_root)` (main.py:26-29) | `sys.path.insert(0, str(APP_DIR))` (1026) | **Similar.** B uses `Path(__file__).parent` (not `.resolve()`) |
| No `.bat` launcher; README says `python main.py` (README.md:112) | `Run [App Name].bat` with `pythonw.exe -B -s` (951-963) | **Missing.** `python.exe` opens a console. Without `-B`, `__pycache__` is written into the app folder (B's .gitignore lists `__pycache__/`) |
| requirements.txt with `>=` ranges and pywin32 commented out | "the packages this app imports, with exact versions" (919, 937-938) | **Differs** (no exact versions; a required package left out) |
| No passwords or tokens stored | 1041-1045 | Nothing to compare (B stores no credentials) |

Standard 12 "Excel through COM" (lines 844-878), for context only, since these are environment rules:
- B does **not** use `DispatchEx`. It uses `GetActiveObject` or `Dispatch` and so may take over the user's Excel.
- It quits that Excel with `DisplayAlerts` switched back on (excel_handler.py:204-211).
- It keeps Excel **visible** for the refresh (report_processor.py:181).
- It refreshes by **SendKeys** instead of `BackgroundQuery=False` + `RefreshAll()` + `CalculateUntilAsyncQueriesDone()`, and relies on fixed waits.
- It does not check `workbook.ReadOnly`.
- It never checks for `DataMashup` before saving with openpyxl, although that path is effectively unreachable (1.3).
- It does call `CoInitialize` in workers and `Quit()` in `finally`.
- Standard 12 also prefers PySide6; B uses PyQt5.

---

## 7. Checklist: what the author needs to supply for a rebuild

**Templates (none are in either repository):**
- [ ] **ITC template** in its real format (.xlsx or .xlsm or .xltx; say which), containing:
  - a `Links` sheet. Mark which of B2/B4/B5/B7/B8/B10/B11 the queries read, and what B3/B6/B9 hold;
  - an `Info` sheet whose BB2 the queries fill (with what?);
  - the sheets `Diff GSTR-3B`, `Reco`, `T6 of R9`, `T8 of R9`, `IMS_Reco`, `ExtraInPurchase` and `ExtraInGSTR2B`, with the cells listed in 1.4C;
  - the Power Query queries themselves. A short note of how each query builds its file path from the Links cells (folder + stem + ".xlsx"? whole folder for GSTR-3B?).
- [ ] **Sales template**, likewise: `Links` (B2/B3/B5/B6/B8/B9), `Info`!BB2, and `Diff - GSTR-1` (I1, N1, O1, P1, Q1).
- [ ] One example of each template **after a successful refresh** (with dummy data), so the extracted cells can be checked.
- [ ] Which Excel version and UI language the office PCs run (the refresh uses English ribbon key tips).

**One sample of each input type** (dummy or anonymised data, real column layout, real naming as the portal or office produces it):
- [ ] GSTR-2B Reco (`GSTR-2B-Reco-<Client>-<State>-<Period>.xlsx`)
- [ ] IMS Reco (`ImsReco-<Client>-<State>-<DDMMYYYY>.xlsx`)
- [ ] GSTR-3B export, two months, including one browser-duplicated download such as `(1) ...` or `... (1).xlsx`
- [ ] Sales (`Sales-<Client>-<State>-<Start>-<End>.xlsx`)
- [ ] Sales Reco (`SalesReco-<Client>-<State>-<Period>.xlsx`)
- [ ] Annual Report (`AnnualReport-<Client>-<State>-<Year>.xlsx`)
- [ ] Any `.xls` or `.xlsm` inputs that occur in practice (`.xlsm` is never recognised by B)
- [ ] A few **real-life awkward names**, anonymised: client names with hyphens, dots, `&` or "Private Limited"; states written as codes; files with spaces around hyphens.

**Sample folder layout:**
- [ ] A source folder exactly as an intern prepares it: flat or with sub-folders? Several clients in one folder?
- [ ] A target folder after one real run, with one `Annual Statement-...` folder, one client and one version, including the refreshed copies. File names only, with dummy content, to confirm the expected tree in 3.2.
- [ ] Where target folders normally live (local, mapped drive, UNC share, OneDrive) and typical path lengths.

**Settings and history:**
- [ ] Whether `~/.gst_organizer/settings.json` and `run_history.json` from a real PC should be migrated. B already holds client figures in `run_history.json` (4.3).

---

## 8. Other technical findings noticed while tracing (for the audit writers)
- **Dead or never-connected items found while tracing state:**
  - The Validation "EXPORT LIST" button has no handler (validation_view.py:186-200; grep `btn_export`).
  - Never used: `models/signals.py` `AppSignals` (never instantiated); `FileParser.validate_results`; `FileOrganizer.cleanup_on_error` and `get_summary`; `DataConsolidator._write_section` and `_create_details_sheet` (data_consolidator.py:237-399); `CollapsibleGroup` (only exported); `utils.helpers` `get_unique_path`, `calculate_file_hash` and `get_relative_excel_path`; the `EXCEL_SAFETY` contents (`self.safety_config` is stored, excel_handler.py:54, never read); `TEMPLATE_EXTENSIONS`, `MAX_FILE_SIZE_MB` and `EXCEL_SIGNATURES` (constants.py:180-185); `GUI_CONFIG` and `MESSAGES` (only re-exported in utils/__init__.py).
  - All of these were confirmed by grep in B-copy.
- The "Folder Name" column in Validation shows the parser key `<client>-<SC>` (validation_view.py:379), **not** the Level-2 folder actually created (create_client_state_key: Pvt/Ltd and truncation) (READ, HIGH).
- The Extract tab uses `self.settings.target_folder` (pyqt_main_window.py:421). That is the last **saved** value, so after editing the target box without scanning or saving, "Refresh List" still reads the old folder (READ, MEDIUM).
- After a failed refresh validation (no Info sheet, BB2 empty, error dialog seen), the consolidated report shows "Missing" for every value of that report, even though the values were extracted (data_consolidator.py:96-117) (READ, HIGH).
- B's README.md is the unchanged v3.0 text: Tkinter folder list (README.md:72-95), `DDMMYYYY_HHMMSS` file names (README.md:182-204) against the real `DDMMYY_HHMM`, `pq_extractor.log`, and "Author: Advanced AI Assistant" (README.md:364). It does not describe B's v4.0 code.
- `.gitignore` lists files tied to the state above: `gst_organizer.log`, `pq_extractor.log`, `.gst_organizer_cache.json` (repo-relative, while the real one is in home), `config/last_session.json`, `temp_gst_*` and `backup_*` (unused prefixes), and `.claude`. B deleted the tracked `.claude/settings.local.json` in c258047; it was added pre-split in bd176bf and only allowed `Bash(python -m py_compile:*)` (tooling).

---

## Possible client names
- **None found in B.** The search covered all code, README.md and B's commit messages (`git log B/main`) for company-like words (Pvt, Ltd, Limited, LLP, Traders, Industries, Enterprises, Sons, Corp, Inc, a GSTIN-shaped regex).
- The only company-like text is the generic example "ABC Ltd" in README.md:128-133 (obviously generic).
- Commit authorship is the author (goyal-sidhant).
- gst_organizer.log (commits edac6a3, 6237597) was **not** opened.
- All names in this file are invented placeholders (ClientA, ClientB, Client A Private Limited, etc.).

## Questions for the author
1. How do the templates' Power Query queries use the Links cells? Folder + stem + ".xlsx"? Does B2 read **every** file in the GSTR-3B folder? This decides whether `.xls` inputs, browser duplicates `(1) ...` and the `_backup_` copies made in Fresh mode would be picked up.
2. What do Links B3/B6/B9 (ITC) and B4/B7 (Sales) contain? The code never writes them.
3. Which template format do you use (.xlsx, .xlsm or .xltx)? B always names the output `.xlsx`.
4. What fills Info!BB2 and what should it contain? The code treats it as "Client Name".
5. Which Excel version and display language are on the office PCs? The refresh presses Alt+A, R, A.
6. [GST RULE] Which source cell should feed the "IGST/CGST/SGST" columns of the "IMS Reco", "Extra In Purchase" and "Extra In GSTR2B" output sheets? Today a lookup defect fills them from Diff GSTR-3B J15/K15/L15 (1.4C). This is your call; the notes only report the mechanics.
7. [GST RULE] The folders and keys use two-letter alphabetic state abbreviations with a fallback (2.5). Is that what you want in folder names, or something else?
8. Do real client names contain hyphens? The file-name rules cannot handle them and sometimes silently mis-split them (2.2).
9. Are source files always in one flat folder, or in sub-folders? B ignores sub-folders.
10. Do you still need the v3 per-client "include client name in brackets" switch? B lost it (3.2).
11. What should Re-Run and Resume mean? Today both create a new Version folder each minute; Resume only skips within the same minute (3.2).
12. Where do target folders live (local, network, OneDrive)? Do two people ever process the same target at the same time?
13. Should run history, which today stores extracted figures in the user profile, be kept at all in a rebuild?

## INCOMPLETE
- **Windows-only behaviour was traced by reading only:**
  - COM report creation, the SendKeys refresh, error-dialog handling, the Excel path limits and the `\\?\` prefix, 8.3 short paths, `os.startfile`;
  - Windows case-insensitive glob and NTFS listing order;
  - `%USERPROFILE%` resolution, and log encoding under a Windows code page.
  - Each is marked READ with a confidence. Checking them needs a Windows PC with Excel and pywin32.
- **Template contents are unknown** (not in either repository). Everything in 1.4 is what the code expects, not what the templates contain.
- The full GUI processing and extraction flow was **not** clicked through end to end. The components were run directly (t2-t8) and the window was built offscreen with stub win32 modules (t6). The real `ProcessingWorker` thread and the real Excel steps were not run.
- Project A was not examined (outside this task).

## Appendix: RAN scripts (all under /root/merge-scratch/run/B-X-EXT/, invented data only)
- `t1_recogniser.py`: parse_filename, state code and suggestion for 66 names.
- `t2_scan.py`: scan_folder on `fake_source/`.
- `t3_names.py`: sanitize_filename, create_client_state_key, get_state_code and name templates.
- `t4_organize.py`: FileOrganizer, ExcelHandler (openpyxl path) and summary/organisation reports into `fake_target/`.
- `t5_modes.py`: Fresh/Resume in a different minute, Level-1 choice by mtime versus name, cumulative report statistics.
- `t6_state.py`: offscreen main window with `HOME` redirected (stub win32 modules in `stubs/`). Settings, session and run-history round trip; old-cache migration.
- `t7_consolidator.py`: consolidated report column filling with marker values.
- `t8_refresh_files.py`: report-file pick, refreshed-copy names, skip-refresh and suffix handling (stub win32 modules).
- `code/` is a copy of B-copy used for imports. The main.py launch test ran from `cwdtest/` (the log appeared there).
