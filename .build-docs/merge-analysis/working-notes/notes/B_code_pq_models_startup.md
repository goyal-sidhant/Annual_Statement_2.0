# B-R2 working notes - project B (AnnualStatement_oldv2, snapshot b0e7d0e) - Power Query Extractor, models, entry points, requirements, .gitignore

Reader: B-R2. Files covered (every line read, with the Read tool or `cat -n`):
`power_query_extractor/__init__.py`, `power_query_extractor/config/__init__.py`, `power_query_extractor/config/cell_mappings.py`,
`power_query_extractor/core/__init__.py`, `power_query_extractor/core/data_consolidator.py`, `power_query_extractor/core/report_processor.py`,
`power_query_extractor/extractor_main.py`, `power_query_extractor/gui/__init__.py`, `launch_extractor.py`, `models/__init__.py`,
`models/app_state.py`, `models/cache_manager.py`, `models/run_history.py`, `models/signals.py`, `main.py`, `requirements.txt`, `.gitignore`.
Other files were read only to find callers or understand a call: `gui/pyqt_main_window.py` (all 608 lines), `gui/controllers/extraction_controller.py`
(all 184 lines), parts of `gui/controllers/processing_controller.py`, `gui/controllers/scan_controller.py`, `gui/views/extraction_view.py`,
`gui/views/validation_view.py`, `core/file_parser.py`, `core/file_organizer.py`, `core/excel_handler.py`, `utils/helpers.py`, `utils/constants.py`,
`README.md` (only the passages about my files), and history via `git -C /root/merge-scratch/combined` (refs `B/main`, `A/main`).

## How to read these notes

- **READ** = I read the code and it says this. **RAN** = I executed it (on copies in `/root/merge-scratch/run/B-R2/`) and saw this result.
- Confidence for every inference: **HIGH / MEDIUM / LOW**.
- **[GST RULE]** = the code decides something the author owns (which return / sheet / cell a figure comes from or goes to, what counts as a
  successful report, display rounding ...). I only describe these; I do not judge them. **TIME-VARYING** = a value that may change over time
  (table numbers, template layouts, cell positions).
- Runs: Python 3.14.6, PyQt5 5.15.11 (Qt 5.15.19, PyQt5_sip 12.20.0), openpyxl 3.1.5, lxml 6.1.3, et_xmlfile 2.0.0, all in
  `/root/merge-scratch/venv-B` (RAN, `importlib.metadata`). **pywin32 is NOT installed** there and there is no Excel. Everything that talks to
  Excel was therefore (a) read line by line and (b) driven with home-made **fake** pywin32 modules in `/root/merge-scratch/run/B-R2/stubs/`
  (`pythoncom.py`, `win32com/client.py`, `win32api.py`, `win32con.py`, `win32gui.py`, `fakecom_state.py`). The fakes record every call and read
  cell values from real .xlsx files made with openpyxl. **They prove only the Python control flow** (which calls happen, in which order, what is
  returned to the caller), NOT how real Excel or Windows behave.
- Scripts: `/root/merge-scratch/run/B-R2/scripts/t_pq.py` (scenarios `normal`, `no_info`, `errint`, `j15none`, `rejected`, `openfail`,
  `errwin`, `skip`, `emptysuffix`, `setprop`, `attached`), `t_models.py`, `t_start.py`, `t_slotexc.py`. HOME was pointed at scratch folders
  (`/root/merge-scratch/run/B-R2/home`, `.../start6_*/home`) so nothing was written to a real home folder. All test data is invented
  (`ClientA-07`, `<Client A>`, numbers 1..1000). No client data was opened.
- Line counts: I give the last line number you can cite (the Read-tool count). Files that do not end with a line-break show one more line than
  `wc -l`: `power_query_extractor/__init__.py`, the three docstring-only `__init__.py` files, `cell_mappings.py`, `data_consolidator.py`,
  `report_processor.py`, `extractor_main.py` and `.gitignore` (RAN, `tail -c1`).

### Plain-English glossary (terms used below)

- **COM**: Windows' mechanism that lets one program (Python) drive another (Excel) by remote control.
- **pywin32**: the Python add-on that provides COM and other Windows calls. Its modules used here: `pythoncom` (COM start/stop per thread),
  `win32com.client` (talk to Excel / WScript.Shell), `win32gui` (list and inspect windows on the screen), `win32con` (Windows message numbers),
  `win32api` (misc Windows calls, here the short "8.3" file name).
- **Dispatch / DispatchEx**: two pywin32 ways to get an Excel object. `DispatchEx` always starts a new, separate Excel; `Dispatch` may connect to
  an Excel that is already running (MEDIUM - see report_processor section). The author's Standard 12 requires `DispatchEx`.
- **SendKeys / AppActivate**: Windows Script Host ("WScript.Shell") functions that bring a window to the front (AppActivate) and type keystrokes
  into whatever window is in front (SendKeys). `%a` in SendKeys means "Alt+A".
- **Power Query (PQ)**: Excel's data-import feature; its queries live inside the workbook and are re-run by "Refresh All".
- **openpyxl**: Python library that reads/writes .xlsx files directly, without Excel.
- **glob**: wildcard file search such as `ITC_Report_*.xlsx`. **stem**: file name without extension.
- **mtime**: a file's "last modified" time stamp.
- **dict**: Python key-to-value table. **dataclass**: a Python class that is mainly a bundle of named fields.
- **JSON**: a plain-text file format for settings (`{"key": value}`).
- **QThread / worker thread**: a background line of execution so the window does not freeze. **signal / slot**: Qt's way of passing a message
  from one object (or thread) to a function in another; the function receiving it is a "slot".
- **exception**: Python's error object. "Caught" = handled by code; "swallowed" = caught and ignored without telling anyone; "re-raised" = passed on.
- **dead code**: code that nothing ever calls.
- **hwnd**: a Windows "window handle" (a number identifying one window on screen). **#32770**: the Windows class name of a standard dialog box.
- **excepthook**: a Python hook that receives any error nothing else caught.
- **pythonw.exe**: the Windows Python that opens no black console window (Standard 13).

---

## Top findings (quick list for later agents; details in the file sections)

1. **Wrong figures in three output sheets (RAN, HIGH).** `data_consolidator.py:130-133` fills each output column by searching for the *first*
   extracted field whose name *contains* the column label. For the one-word labels `IGST`, `CGST`, `SGST` (output sheets `IMS Reco`,
   `Extra In Purchase`, `Extra In GSTR2B`, configured in `cell_mappings.py:155-228`) the first match is always the GSTR-3B difference field, so
   these 9 columns show `Diff GSTR-3B` J15/K15/L15 instead of their own configured cells, whenever J15/K15/L15 are not empty. RAN with invented
   values: expected 10/20/30, 40/50/60, 70/80/90; the report showed 111/222/333 in all three sheets. [GST RULE] material (which cell a figure
   comes from) - described only. The same line exists in project A (`A/main:power_query_extractor/core/data_consolidator.py:143`).
2. **The whole app cannot start without pywin32 (RAN, HIGH)**, although `requirements.txt:13-14` lists pywin32 only as a commented-out
   "Windows-only" line. `report_processor.py:11-15` imports pywin32 unconditionally; it is imported at start-up through
   `gui/controllers/extraction_controller.py:10` <- `gui/pyqt_main_window.py:27` <- `main.py:38`. main.py then tells the user to
   `pip install PyQt5 openpyxl` (`main.py:51-53`), which does not fix it. Other files guard the same import (`core/excel_handler.py:20-27`,
   `utils/helpers.py:476-477`).
3. **Refresh is done by typing keystrokes into a visible, maximised Excel (READ, HIGH)** obtained with `Dispatch` (not `DispatchEx`), then fixed
   sleeps (about 39 s per report at the default 10 s), then save. No `RefreshAll()`, no `BackgroundQuery = False`, no
   `CalculateUntilAsyncQueriesDone()`, no check that queries finished (Standard 12 lines 851-875). Keystrokes are sent even if the Excel window
   could not be brought to the front (`report_processor.py:229-248`).
4. **Error-dialog detection looks at every window on the desktop (READ, HIGH)** and clicks the first button whose text *contains* `ok`, `close`,
   `yes` or `continue` (`report_processor.py:368-476`). It can close dialogs of other programs, and any dialog titled "Microsoft Excel" counts
   as an error.
5. **What the report says vs what happened (RAN, HIGH):** the "Successfully Run" tick is the *refresh* status; when it is false every figure is
   replaced by `Missing` although the figures were read (`data_consolidator.py:116-117`); "skip refresh" re-uses an old refreshed copy and
   reports success even if that old refresh failed (`report_processor.py:100-111`); a custom file suffix breaks "skip refresh" and the
   "Refreshed" status in the tree (hard-coded `_Refreshed_` at `report_processor.py:102` and `gui/pyqt_main_window.py:468`).
6. **Run history never links extraction to processing (RAN, HIGH):** extraction results are saved under the key `"<client folder> / <version
   folder>"` (`gui/pyqt_main_window.py:556`), processing records under `"<client>-<state code>"` (`core/file_parser.py:262`), so
   `add_extraction_result` (`run_history.py:73-95`) always creates a separate record; `pq_refreshed` is never true on a processing record.
   `client_history` grows without limit; extracted figures stay in `~/.gst_organizer/run_history.json` for ever.
7. **`power_query_extractor/extractor_main.py` is broken and dead (RAN, HIGH):** line 9 imports `power_query_extractor/gui/extractor_window.py`,
   which B deleted in a62d3cc (07-02-2026): `ModuleNotFoundError`. Nothing calls it.
8. **Dead or write-only state (READ/RAN, HIGH):** `models/signals.py` is never imported; `ProcessingState` is mostly unused and
   `client_folder_settings` is never filled; `session_state.json` (`active_tab`, `window_size`, `recent_source_folders`) is written but never
   read; the v3 cache migration loses the old recent-folders list (`recent_folders` vs `recent_source_folders`).
9. **Start-up vs Standard 13 (READ/RAN):** no `.bat` launcher / `pythonw`, errors only `print()`ed, log file `gst_organizer.log` written to the
   current folder, no `sys.excepthook` - and RAN: an uncaught error inside any PyQt5 slot aborts the whole app (exit code 134) with no message.
   Settings live in `~/.gst_organizer`, not `%LOCALAPPDATA%\[App Name]\`. GUI is PyQt5, Standard 12 names PySide6.
10. **Excel error values may pass as figures (MEDIUM):** through COM an error cell (`#N/A`, `#DIV/0!`) usually arrives as a large negative whole
    number, not as text, so `_validate_refresh` (`report_processor.py:626-629`) would not recognise it and the consolidated report would show
    it as a number (simulated RAN with assumed values).

---

### power_query_extractor/__init__.py (3 lines) - [CORE]

- Purpose: marks `power_query_extractor` as a Python package (a folder Python can import from) and sets a version string. READ.
- Imports: none.
- Module-level constants: `__version__ = "1.0.0"` (line 3). Never read anywhere in B (grep for `__version__` finds only this line,
  `gui/__init__.py:11` = `'4.0.0'`, `core/__init__.py:12` = `'3.0.0'`, `utils/__init__.py:38` = `'3.0.0'`) -> dead constant. READ, HIGH.
- Classes/functions: none.
- Comments/docstrings: line 1 `"""Power Query Extractor Module for GST File Organizer"""` - agrees with content.
- Version strings: four different version numbers coexist (1.0.0 here, 4.0.0 in gui, 3.0.0 in core and utils, "4.0" in `main.py:2,42` and
  window title `gui/pyqt_main_window.py:71`). None is used for logic. READ.
- History: created edac6a3 (31-08-2025), never changed (git log B/main). READ.
- Gaps: none (no code).

### power_query_extractor/config/__init__.py (1 line) - [CONFIG]

- Purpose: package marker only; one docstring `"""Configuration for Power Query Extractor"""`, no code, no re-exports. READ.
- Imports / constants / functions: none. Callers import `power_query_extractor.config.cell_mappings` directly.
- History: edac6a3, unchanged.

### power_query_extractor/config/cell_mappings.py (278 lines) - [CONFIG]

- Purpose: the "what to extract" table: which sheet and cell of the ITC / Sales report each figure is read from, and which output sheet and
  column it goes to in the consolidated report; plus four options; plus an automatically built older-format copy of the same table. READ.
- Imports: none (pure data + one loop that runs at import time).

#### Module-level constants and globals

| Name | Lines | Value / shape | What it controls | Who reads it |
|---|---|---|---|---|
| `EXTRACTION_CONFIG` | 33-242 | dict: output-sheet name -> `{'report_type': 'ITC'/'Sales', 'mappings': [ {input_sheet, input_cell, output_column, calculation}, ... ]}` | Which output sheets exist (one per key, in this order), their column headers, which cells feed them | `data_consolidator.py:24, 45` (sheets), `:210, 233` (summary list); indirectly everything via `CELL_MAPPINGS` |
| `EXTRACTION_OPTIONS` | 248-253 | `{'skip_refresh': False, 'auto_refresh_if_missing': True, 'show_timestamp': False, 'missing_data_text': 'Missing'}` | see table below | `report_processor.py:97-98`, `data_consolidator.py:24, 117, 142`; written by `gui/controllers/extraction_controller.py:54-55` |
| `CELL_MAPPINGS` | 260, filled 264-278 | dict: report type -> input sheet -> cell -> field name | Which cells `_extract_data` actually reads (`report_processor.py:497`) | `report_processor.py:16, 497`; `data_consolidator.py:11, 243` (only inside dead `_write_section`) |
| `FIELD_DISPLAY_NAMES` | 261, filled 278 | dict: field name -> output_column label | Labels for the dead `_write_section` | `data_consolidator.py:11, 262` (dead code only) |
| loop variables `sheet_name`, `config`, `report_type`, `mapping`, `sheet`, `field_name` | 264-277 | left behind as module globals after the loop | nothing (harmless leftovers) | nobody |

`EXTRACTION_OPTIONS` one by one (READ, grep across all .py files):

| Option | Line | Value | Effect in the running app |
|---|---|---|---|
| `skip_refresh` | 249 | `False` | Read at `report_processor.py:98`. But `extraction_controller.py:55` overwrites it with the Tab 4 checkbox value at the start of every extraction, so the value in this file has no effect in the app (READ, HIGH). |
| `auto_refresh_if_missing` | 250 | `True` | **Never read anywhere** (dead option). The behaviour it describes ("refresh if no refreshed file found") is hard-coded at `report_processor.py:112-115` and cannot be switched off. READ, HIGH. |
| `show_timestamp` | 251 | `False` | **Never read anywhere** (dead option; no timestamp feature exists). READ, HIGH. |
| `missing_data_text` | 252 | `'Missing'` | Text written into the consolidated report instead of a figure (`data_consolidator.py:117, 142`). |

#### EVERY mapping - [GST RULE], TIME-VARYING (template layouts, sheet names, cell positions and table numbers can change between years or template versions)

Output layout of every output sheet (from `data_consolidator.py:68-89`, RAN): A1 = "`<output sheet> Report`" (merged A1:C1), row 3 = headers
(A `Client Name`, B `Successfully Run`, then one column per mapping in list order starting at C), data rows from row 4, one row per extracted
client/version. "Field name" = the key built at line 276 (`f"{output sheet}_{output_column}"` with spaces replaced by `_`), RAN.
All 22 mappings have `'calculation': None` (plain copy). No entry has a `description` key.

| # | Output sheet (line) | Report type | Input sheet | Input cell | Output column label | Output col. | Field name (CELL_MAPPINGS) | Lines | What the code actually writes (RAN) |
|---|---|---|---|---|---|---|---|---|---|
| 1 | `GSTR-3B Difference` (35) | ITC | `Diff GSTR-3B` | J15 | `Diff IGST` | C | `GSTR-3B_Difference_Diff_IGST` | 38-43 | J15 (correct, via exact-key fallback) |
| 2 | `GSTR-3B Difference` | ITC | `Diff GSTR-3B` | K15 | `Diff CGST` | D | `GSTR-3B_Difference_Diff_CGST` | 44-49 | K15 |
| 3 | `GSTR-3B Difference` | ITC | `Diff GSTR-3B` | L15 | `Diff SGST` | E | `GSTR-3B_Difference_Diff_SGST` | 50-55 | L15 |
| 4 | `GSTR-1 Difference` (61) | Sales | `Diff - GSTR-1` | N1 | `Diff Taxable Value` | C | `GSTR-1_Difference_Diff_Taxable_Value` | 64-69 | N1 |
| 5 | `GSTR-1 Difference` | Sales | `Diff - GSTR-1` | O1 | `Diff IGST` | D | `GSTR-1_Difference_Diff_IGST` | 70-75 | O1 |
| 6 | `GSTR-1 Difference` | Sales | `Diff - GSTR-1` | P1 | `Diff CGST` | E | `GSTR-1_Difference_Diff_CGST` | 76-81 | P1 |
| 7 | `GSTR-1 Difference` | Sales | `Diff - GSTR-1` | Q1 | `Diff SGST` | F | `GSTR-1_Difference_Diff_SGST` | 82-87 | Q1 |
| 8 | `Turnover GSTR-1` (92) | Sales | `Diff - GSTR-1` | I1 | `Taxable Value` | C | `Turnover_GSTR-1_Taxable_Value` | 95-100 | I1 |
| 9 | `Reco Diff` (105) | ITC | `Reco` | C37 | `IGST Diff` | C | `Reco_Diff_IGST_Diff` | 108-113 | C37 |
| 10 | `Reco Diff` | ITC | `Reco` | D37 | `CGST Diff` | D | `Reco_Diff_CGST_Diff` | 114-119 | D37 |
| 11 | `Reco Diff` | ITC | `Reco` | E37 | `SGST Diff` | E | `Reco_Diff_SGST_Diff` | 120-125 | E37 |
| 12 | `Table 6J Match` (130) | ITC | `T6 of R9` | K29 | `Table 6J Match` | C | `Table_6J_Match_Table_6J_Match` | 133-138 | K29 |
| 13 | `Table 8D Limits` (143) | ITC | `T8 of R9` | G6 | `Table 8D Limits` | C | `Table_8D_Limits_Table_8D_Limits` | 146-151 | G6 |
| 14 | `IMS Reco` (156) | ITC | `IMS_Reco` | AA1 | `IGST` | C | `IMS_Reco_IGST` | 159-164 | **`Diff GSTR-3B`!J15** unless J15 is empty, then AA1 |
| 15 | `IMS Reco` | ITC | `IMS_Reco` | AD1 | `CGST` | D | `IMS_Reco_CGST` | 165-170 | **`Diff GSTR-3B`!K15** unless K15 is empty, then AD1 |
| 16 | `IMS Reco` | ITC | `IMS_Reco` | AG1 | `SGST` | E | `IMS_Reco_SGST` | 171-176 | **`Diff GSTR-3B`!L15** unless L15 is empty, then AG1 |
| 17 | `Extra In Purchase` (181) | ITC | `ExtraInPurchase` | O1 | `IGST` | C | `Extra_In_Purchase_IGST` | 184-189 | **J15** unless empty, then O1 |
| 18 | `Extra In Purchase` | ITC | `ExtraInPurchase` | P1 | `CGST` | D | `Extra_In_Purchase_CGST` | 190-195 | **K15** unless empty, then P1 |
| 19 | `Extra In Purchase` | ITC | `ExtraInPurchase` | Q1 | `SGST` | E | `Extra_In_Purchase_SGST` | 196-201 | **L15** unless empty, then Q1 |
| 20 | `Extra In GSTR2B` (206) | ITC | `ExtraInGSTR2B` | S1 | `IGST` | C | `Extra_In_GSTR2B_IGST` | 209-214 | **J15** unless empty, then S1 |
| 21 | `Extra In GSTR2B` | ITC | `ExtraInGSTR2B` | T1 | `CGST` | D | `Extra_In_GSTR2B_CGST` | 215-220 | **K15** unless empty, then T1 |
| 22 | `Extra In GSTR2B` | ITC | `ExtraInGSTR2B` | U1 | `SGST` | E | `Extra_In_GSTR2B_SGST` | 221-226 | **L15** unless empty, then U1 |

Totals (RAN): 9 output sheets, 22 mappings (ITC 17 = 3+3+1+1+3+3+3; Sales 5 = 4+1), 22 field names, no output-sheet name longer than 31
characters (Excel's limit). The input sheet `Diff - GSTR-1` feeds two output sheets (rows 4-8). The "what the code writes" column is explained
under data_consolidator.py (`create_report`, lines 130-138); RAN with the fake-COM script (`t_pq.py normal`, `j15none`).

Derived older-format table `CELL_MAPPINGS` (RAN, printed): `ITC` -> `Diff GSTR-3B` {J15,K15,L15}, `Reco` {C37,D37,E37}, `T6 of R9` {K29},
`T8 of R9` {G6}, `IMS_Reco` {AA1,AD1,AG1}, `ExtraInPurchase` {O1,P1,Q1}, `ExtraInGSTR2B` {S1,T1,U1}; `Sales` -> `Diff - GSTR-1`
{N1,O1,P1,Q1,I1}, each cell mapped to the field name in the table above. Sheet order = first appearance in `EXTRACTION_CONFIG`; this order is
the order in which `_extract_data` reads the cells and therefore the order of keys in the extracted dict - which decides the wrong matches in
rows 14-22.

#### The builder loop (lines 259-278), step by step (READ, RAN)

1. For each output sheet in `EXTRACTION_CONFIG` (dict order): take its `report_type`; create `CELL_MAPPINGS[report_type]` if missing (266-267).
2. For each mapping: only if both `input_sheet` and `input_cell` are truthy (non-empty) (270): create `CELL_MAPPINGS[rt][sheet]` if missing;
   field name = `f"{output sheet}_{output_column}"` with spaces -> `_` (276); `CELL_MAPPINGS[rt][sheet][cell] = field name` (277);
   `FIELD_DISPLAY_NAMES[field name] = output_column` (278).
3. Runs once, when the module is first imported. Editing the file needs an app restart.

#### Business rules in this file

- Every row of the mapping table above is a [GST RULE] (which return-related figure is taken from which template cell, and under which label it
  is reported). TIME-VARYING: sheet names, cell positions, table numbers (`Table 6J`, `Table 8D`, `T6 of R9`, `T8 of R9`), return names
  (GSTR-3B, GSTR-1, GSTR-2B), the `IMS` reference. Described only.
- `report_type` decides which report file a figure is read from: `ITC` -> files matching `ITC_Report_*.xlsx`, `Sales` -> `Sales_Report_*.xlsx`
  (`report_processor.py:76`). [GST RULE - when in doubt].
- Calculation types documented at lines 21-26 (`subtract`, `add`, `multiply`, `divide`) - implemented in `data_consolidator.py:176-206`, used
  by no current mapping. [GST RULE] mechanism (would compute figures if configured).

#### Edge cases explicitly handled

- Empty `input_sheet` or `input_cell` -> mapping not read from the report (line 270). Such a mapping can still produce an output column (a
  calculated one, or `Missing`).

#### Gaps, likely failures and silent bugs

- Two mappings with the same input sheet AND cell (in different output sheets) would overwrite each other at line 277: only the later field
  name would be read; the earlier output column would then be filled by the substring search or show `Missing`. Not the case today (RAN: no
  duplicates). LATENT, READ, HIGH.
- Two output-sheet/column combinations that differ only by space vs underscore would produce the same field name (276). Latent. READ.
- An output-sheet name over 31 characters or containing `[ ] : * ? / \` would break or alter `wb.create_sheet` in the consolidator. Latent.
  MEDIUM (openpyxl behaviour for long titles not tested here).
- A calculation's `cell1`/`cell2` are bare cell references with no sheet name (docstring line 15). They are looked up in the extracted values,
  which contain only cells listed in `CELL_MAPPINGS`; an unlisted cell silently counts as 0 (`data_consolidator.py:182-188`). Latent (no
  calculations configured). READ, HIGH.
- The module-level dict `EXTRACTION_OPTIONS` is changed at run time by `extraction_controller.py:55` (a global, shared, mutable setting).
  Harmless today because every run sets it first. READ.

#### Comments, docstrings, TODO, version strings

- Docstring 1-27: "MASTER VERSION", how to edit, calculation types. Agrees with the code, except it does not mention the optional
  `description` key that the consolidator supports (`data_consolidator.py:53-65`).
- Lines 229-241: commented-out example entry `'My Custom Sheet'` (template for adding sheets; not dead logic).
- Line 255 "DON'T EDIT BELOW THIS LINE"; line 259 "Backward compatibility" - the older format is the one of the deleted file
  `power_query_extractor/config/cell_mappings copy.py` (removed in 47a1492; it had a direct `CELL_MAPPINGS` dict with field names like
  `itc_diff_igst` and a `FIELD_DISPLAY_NAMES` dict). READ (git show 47a1492^).

#### History

- edac6a3 (31-08-2025): created with the first three output sheets (rows 1-8 of the table) - `git diff edac6a3 47a1492^` on this file is empty.
- 47a1492 (30-10-2025, last shared commit): added the six ITC output sheets `Reco Diff`, `Table 6J Match`, `Table 8D Limits`, `IMS Reco`,
  `Extra In Purchase`, `Extra In GSTR2B` (rows 9-22) and deleted `cell_mappings copy.py`.
- Not changed after the split in either project (`git diff --stat 47a1492 A/main` and `... B/main` do not list it). READ, HIGH.

#### Pattern notes

- Configuration lives in Python code that the user is told to edit (docstring), not in a data file. Options are read with `.get(key, default)`.

### power_query_extractor/core/__init__.py (1 line) - [CORE]

- Purpose: package marker; one docstring `"""Core processing components"""`, no code, no re-exports. READ. History: edac6a3, unchanged.

### power_query_extractor/core/data_consolidator.py (417 lines) - [CORE]

- Purpose: builds the consolidated Excel report `PQ_Extraction_Report_<YYYYmmdd_HHMM>.xlsx` (one sheet per `EXTRACTION_CONFIG` entry plus a
  `Summary` sheet) from the per-client extraction results, using openpyxl. READ, RAN.
- Imports: standard library `logging` (5), `pathlib.Path` (6, **unused** - no `Path(` in the file), `datetime` (7); third-party openpyxl
  `Workbook` (8), `Font, PatternFill, Alignment, Border, Side` (9), `get_column_letter` (10); project `FIELD_DISPLAY_NAMES, CELL_MAPPINGS`
  from `..config.cell_mappings` (11; used only by the dead `_write_section`). Function-local imports: `EXTRACTION_CONFIG, EXTRACTION_OPTIONS`
  (24), `EXTRACTION_CONFIG` (210). No OS-specific modules.
- Module-level globals: `logger = logging.getLogger(__name__)` (13).

#### class DataConsolidator (16-417) - no `__init__`, no stored state

**`create_report(self, results, output_folder)` - lines 20-174**
- Parameters: `results` = list of per-client dicts as returned by `ReportProcessor.process_client` (`{'client', 'timestamp', 'itc', 'sales'}`);
  `output_folder` = a `pathlib.Path` (it is joined with `/` at line 165, so a plain string would fail; the caller passes a Path,
  `extraction_controller.py:29, 92`).
- Steps (READ, RAN):
  1. 24: import the config dicts. 27-31: new empty workbook, remove its default sheet.
  2. 34-42: styles - header font bold white size 12; header fill solid `0078D4` (blue); section font bold 14; thin border on all four sides.
  3. 45-157: for each output sheet in `EXTRACTION_CONFIG` order: log "Creating sheet: ..."; `wb.create_sheet(name)` (49).
     - 53-65: if the entry has a `description` list: each line in A1, A2 ... (first bold 14, others italic 11), each merged A:J, start row =
       lines + 2. **No current entry has `description` -> this branch is dead with the current config** (READ, HIGH).
     - 66-71 (always taken today): A1 = `"<name> Report"`, bold 14, merge A1:C1, start row 3.
     - 74-78: headers `['Client Name', 'Successfully Run']` + each mapping's `output_column`; 81-86 write them in the start row with the header style.
     - 89-154: one row per result, from start row + 1:
       - 94: column A = `result['client']` (for B this is `"<client folder> / <version folder>"`, see report_processor 44 and
         `gui/views/extraction_view.py:384`), with border.
       - 97: `report_data = result.get('itc' or 'sales', {})` (`report_type.lower()`).
       - 98: `is_success = report_data['status']['success']` (via `.get`, default False). **This is the refresh status** set by
         `report_processor.py` (refresh + Info!BB2 check), not "values were read".
       - 100-109: column B = `'✓'` (green `107C10`, bold) or `'✗'` (orange-red `D83B01`, bold).
       - 112-152: for each mapping, column C onwards:
         - 116-117: if not success -> value = `missing_data_text` (`'Missing'`) **for every column, even though `_extract_data` did read the
           values** (RAN `t_pq.py no_info`: all 22 columns `Missing`, while `res['itc']['data']` held the figures). [GST RULE - when in doubt:
           decides whether a figure is reported].
         - 118-124: elif the mapping has a truthy `calculation` -> `_calculate_value(report_data['extracted_values'], calc, wb)`.
           No mapping has one today -> unused path.
         - 125-138: else: `extracted_data = report_data.get('data', {})`; **130-133: loop over every extracted key in insertion order and take
           the first one where `mapping['output_column'] in field_name` (substring test)**; 136-138: only if that gave `None`, try the exact key
           `f"{sheet_name}_{output_column}"` with spaces -> `_`.
         - 141-142: `None` or `''` -> `'Missing'`.
         - 145-146: write value with border. 149-150: number format `'#,##0.00'` only if the value is `int`/`float` and not 0 (RAN: `0` stays
           `General`; `Decimal('12.50')` written as 12.5 with `General`; Python `True`/`False` would count as int). Display only - the stored
           value is not rounded. [GST RULE - when in doubt: display rounding].
  4. 157: `_autofit_columns(ws)` for each output sheet. 160: `_create_summary_sheet(wb, results)` (inserted as first sheet).
  5. 163-167: name `PQ_Extraction_Report_<YYYYmmdd_HHMM>.xlsx` (local time, minute precision), saved in `output_folder` with `wb.save`.
     **A second run in the same minute silently overwrites the first report** (RAN: five `create_report` calls within one minute left one file).
  6. 170: returns the saved `Path`.
- Exceptions: 172-174 catch everything, log with traceback, **re-raise**. The caller (`extraction_controller.py:90-95`) catches it and writes
  `"  Report error: <message>"` to the Tab 4 log (no message box). READ.
- Side effects: one new .xlsx in the Level 1 folder; log lines.
- Callers: `gui/controllers/extraction_controller.py:92` (only when `results` is non-empty, line 87). READ (grep).

**THE SUBSTRING BUG (lines 130-133) - RAN, HIGH.** Extracted keys for an ITC report come in this order (RAN): `GSTR-3B_Difference_Diff_IGST`,
`J15`, `GSTR-3B_Difference_Diff_CGST`, `K15`, `GSTR-3B_Difference_Diff_SGST`, `L15`, `Reco_Diff_IGST_Diff`, `C37`, ... `IMS_Reco_IGST`, `AA1`, ...
For output column `IGST` of `IMS Reco`, the first key containing the text `IGST` is `GSTR-3B_Difference_Diff_IGST`, so the value of
`Diff GSTR-3B`!J15 is written. Labels that contain a space (`Diff IGST`, `IGST Diff`, `Taxable Value`, `Table 6J Match` ...) never match a key
(keys have underscores), so they fall through to the correct exact-key lookup. Result (RAN, invented values J15/K15/L15 = 111/222/333,
IMS_Reco 10/20/30, ExtraInPurchase 40/50/60, ExtraInGSTR2B 70/80/90): sheets `IMS Reco`, `Extra In Purchase`, `Extra In GSTR2B` all showed
111 / 222 / 333. With J15 empty (RAN `j15none`): IGST columns showed 10 / 40 / 70 (correct, fallback) but CGST/SGST still 222/333. The
substring line has existed since edac6a3; it became harmful when 47a1492 added one-word labels. Project A has the identical line
(`A/main:power_query_extractor/core/data_consolidator.py:143`). The figures are [GST RULE] material; the bug is that the output does not
follow the configuration.

**`_calculate_value(self, extracted_values, calculation, workbook)` - lines 176-206**
- Parameter `workbook` is **never used** (dead parameter).
- Steps: 179 type; 182-183 `cell1`/`cell2` values from `extracted_values` (default 0); 186-190 convert with `float(v) if v else 0` - any
  falsy value (None, 0, '') becomes 0; a non-numeric text (for example `'1,000'`) -> returns `'Error'`; 193-202: `subtract` / `add` /
  `multiply` / `divide` (division by 0 -> text `'Div/0'`), any other type -> `'Unknown'`.
- Exceptions: 204-206 any other error -> logs "Calculation error", returns `'Error'` (swallowed into a text result).
- Return: a float or one of the texts `'Error'`, `'Div/0'`, `'Unknown'`, which would then be written as the figure.
- Callers: line 120 only, and only for mappings with a `calculation` - none today -> unused in practice. READ, HIGH.
- [GST RULE] mechanism (would compute figures). Gap: missing cells silently count as 0 (no warning).

**`_create_summary_sheet(self, wb, results)` - lines 208-235**
- Creates sheet `Summary` at position 0 (212). A1 `"Extraction Summary"` (bold 16); A3 `"Generated: <YYYY-mm-dd HH:MM:SS>"`; A4
  `"Total Clients: <n>"`; A5 `"ITC Reports Processed: <x>/<n>"`; A6 `"Sales Reports Processed: <y>/<n>"` where x / y count results whose
  refresh status is success (222-223); A8 `"Sheets Created:"` (bold); A9 onward `"• <output sheet>"` for every `EXTRACTION_CONFIG` key.
  RAN (empty and one-client inputs).
- "Processed" means "refresh succeeded"; a report not selected (`process_itc` False) counts as not processed. No column widths set.
- If an output-sheet key were `Summary`, openpyxl would rename one of them (latent, not tested).
- Callers: line 160.

**`_write_section(self, ws, results, report_key, report_type, start_row, header_font, header_fill, border, subsection_font)` - lines 237-339**
- **NO CALLERS FOUND (dead)** - grep in B, and also no callers in edac6a3 or 47a1492 (it was already dead in the initial commit). READ, HIGH.
- What it would do: collect all field names from `CELL_MAPPINGS[report_type]` plus all keys of the results' data (242-252), sorted; if none,
  write an italic grey `"No cell mappings defined for ..."` (254-258); headers `Client Name`, `Refresh Status`, display names, `Error Details`;
  per client `'✓ Success'`/`'✗ Failed'`, each value or `'-'`, error text; a `"Summary:"` row with `"<s>/<n> successful"`; returns the last row.
  Would raise `KeyError` on a result without `status`. Remnant of an older one-sheet layout (MEDIUM inference).

**`_create_details_sheet(self, wb, results)` - lines 341-399**
- **NO CALLERS FOUND (dead)** (also dead in edac6a3 / 47a1492). Would write an `Extraction Details` sheet: per client `Client: <name>`, ITC
  and Sales status, validation message, and every extracted `field: value` line.

**`_autofit_columns(self, ws)` - lines 401-417**
- For each column: longest `str(value)` among truthy values (0 and '' are ignored) (403-413), width = `min(longest + 2, 50)`, set if > 0
  (always at least 2). `column[0].column` works on merged cells too (RAN, no error). The merged A1 title counts toward column A's width (RAN:
  `GSTR-3B Difference` sheet column A = 27). Exceptions `TypeError`/`ValueError` per cell swallowed (412-413).
- Callers: line 157.

#### Business rules in this file

- [GST RULE - when in doubt] A figure is reported only when the report's refresh status is success; otherwise `Missing` (116-117).
- [GST RULE - when in doubt] Empty (`None` / `''`) figures are shown as `Missing`; 0 is shown as 0 (141-142).
- [GST RULE - when in doubt] Display format `#,##0.00` (two decimals, thousands separator) for non-zero int/float (149-150); not applied to 0,
  to text, or to `Decimal` values.
- Column B header `Successfully Run`, values ✓ / ✗ (74, 100-109).

#### Edge cases explicitly handled

- Report not selected / not found / refresh failed -> `Missing` in every column (116-117). Empty results list -> report with headers only
  and `Total Clients: 0` (RAN; but the caller never calls it with an empty list). Division by zero in calculations -> `'Div/0'` (200).

#### Gaps, likely failures and silent bugs

- Substring bug (above) - RAN, HIGH.
- Figures hidden as `Missing` when refresh/validation failed although they were read (RAN). The user cannot tell "not found" from "found but
  refresh doubtful".
- Same-minute overwrite of the report (RAN). Running twice on the same data creates a second report file (new minute) - no clean-up.
- A date value that arrives as a time-zone-aware datetime makes the **whole report fail to save** (RAN with a tz-aware datetime: openpyxl
  `TypeError: Excel does not support timezones in datetimes`, re-raised; GUI log shows "Report error"). Whether pywin32 delivers dates as
  tz-aware is not verifiable here (MEDIUM: recent pywin32 builds return tz-aware datetimes; check on Windows by mapping a date cell).
- Excel error cells arriving as big negative integers would be written as figures with `#,##0.00` (simulated RAN `errint` with the assumed
  value -2146826246 for `#N/A`; MEDIUM that real COM returns integers - check on Windows with `=NA()` in a mapped cell).
- `Decimal` (what COM may return for cells formatted as Currency, MEDIUM) is written as a number but without the `#,##0.00` format (RAN).
- No check that `output_folder` exists or is writable; a network-drive disconnect raises at `wb.save` -> re-raised -> "Report error" in GUI log;
  the per-client figures are then lost except in run history (`pyqt_main_window.py:557`).
- Not thread-related: called inside the extraction worker thread, touches no widgets. OK.

#### Excel workbook touchpoints

- 27 `Workbook()` create (new consolidated report, openpyxl); 31 remove default sheet; 49, 212 create sheets; 167 `wb.save(report_path)`.
  The report is a new workbook with no Power Query, so openpyxl saving is in line with Standard 12 line 827. READ.

#### Files written

- `<Level 1 folder>/PQ_Extraction_Report_<YYYYmmdd_HHMM>.xlsx` (163-167). The Level 1 folder is the one shown in Tab 4
  (`gui/pyqt_main_window.py:512, 543` -> `extraction_controller.py:29, 92`).

#### Comments, docstrings, history

- Docstrings agree with code except `create_report` "multiple custom sheets" (fine). Comment 23 "Import the new configuration" - fine.
- c258047 (07-02-2026, B only): removed a duplicated header-building block with "# ... rest of the code" placeholders, two bare `except:`
  narrowed to `(TypeError, ValueError)`, removed "# ADD THIS" comments. No behaviour change (READ, git show).

#### Pattern notes

- Config imported both at module level (11) and inside functions (24, 210) - inconsistent but harmless. Errors are logged and re-raised in
  `create_report`, swallowed into text results in `_calculate_value`.

### power_query_extractor/core/report_processor.py (752 lines) - [CORE]

- Purpose: for one client/version folder, finds the ITC and/or Sales report, makes a time-stamped copy, refreshes the copy's Power Query by
  driving a visible Excel with keystrokes, checks `Info!BB2`, saves, then opens the copy again in a second (hidden) Excel and reads the
  configured cells. READ; Python control flow RAN with fake pywin32.
- Imports (all at module level, **unconditional** - no try/except, no platform check):
  - standard library: `os` (5, **unused**), `time` (6), `shutil` (7), `logging` (8), `Path` (9), `datetime` (10);
  - Windows-only (pywin32): `pythoncom` (11), `win32com.client` (12), `win32api` (13, used only at 711 where it is imported again locally at
    710), `win32con` (14), `win32gui` (15). On a PC without pywin32 the import fails with `ModuleNotFoundError: No module named 'pythoncom'`
    (RAN). Because `gui/controllers/extraction_controller.py:10` imports this module at module level, and `gui/pyqt_main_window.py:27` imports
    that, **the whole application fails to start** (RAN: `main.py` prints "Missing module: No module named 'pythoncom'" and exits 1;
    `launch_extractor.py` prints "Failed to launch: No module named 'pythoncom'"). Inconsistent with the project's own pattern in
    `core/excel_handler.py:20-27` (platform check + `try/except ImportError` + flag) and `utils/helpers.py:476-477` (platform check). HIGH.
  - project: `CELL_MAPPINGS` (16); function-local `EXTRACTION_OPTIONS` (97).
- Module-level globals: `logger` (18).

#### Operational / safety constants (all hard-coded, READ)

| What | Where | Value |
|---|---|---|
| Default refresh wait | 27 | `self.wait_time = 10` s (overwritten by Tab 4 spin box, range 5-120, `gui/views/extraction_view.py:126-128`, via `extraction_controller.py:50`) |
| Default copy suffix | 28 | `"_Refreshed_{timestamp}"` (overwritten by Tab 4 text box, `extraction_controller.py:51`) |
| Report file pattern | 76 | `f"{report_type}_Report_*.xlsx"` -> `ITC_Report_*.xlsx`, `Sales_Report_*.xlsx` |
| Copy time stamp | 92 | `datetime.now().strftime("%d%m%y_%H%M")` (DDMMYY_HHMM, minute precision) |
| Existing-copy pattern (skip mode) | 102 | `f"{stem}_Refreshed_*.xlsx"` (hard-coded, ignores the suffix setting) |
| Errors treated as "could not open" | 126 | text contains `"Open method of Workbooks class failed"` or `"-2147352567"` (0x80020009, COM's generic "exception occurred") |
| File-release wait after refresh | 141 | 3 s |
| Retries for Excel properties | 178-190 | 5 attempts, 2 s apart, then continue anyway |
| `excel.WindowState` | 183 | `-4137` (xlMaximized) |
| Open retries | 195-205 | 3 attempts, 3 s apart, only for `"Call was rejected by callee"` (Excel busy) |
| Load wait | 209-211 | 5 x 1 s |
| AppActivate retries | 220-227 | 3 attempts, 1 s apart; then 1 s (232) |
| Keystrokes | 238-248 | `"%a"` (Alt+A), 1 s, `"r"`, 1 s, `"a"` = Data tab -> Refresh All -> Refresh All (key tips; MEDIUM, depends on Excel version/language) |
| Refresh wait | 256-258 | `wait_time` x 1 s |
| Error-dialog checks | 266-275 | 5 checks, 1 s apart |
| Extra wait | 282-286 | `max(10, wait_time)` x 1 s |
| After save / after close | 307, 319 | 2 s, 1 s |
| Error-window title words | 376-387 | `microsoft excel`, `error`, `problem`, `cannot`, `failed`, `unable`, `could not`, `exception`, `invalid`, `missing` |
| Dialog class names | 392 | `#32770`, `NUIDialog`, `bosa_sdm_msword` |
| Child-text minimum length | 420 | more than 5 characters |
| Error text maximum | 435 | 200 characters; fallback text `"Power Query error occurred"` (437, 441) |
| Buttons clicked | 457 | button text containing `ok`, `close`, `yes`, `continue` (substring, lower-case) |
| Pause after closing dialog | 473 | 0.5 s |
| Validation cell | 597-607 | sheet named `info` (any case), cell `BB2` |
| Values rejected in BB2 | 610-629 | `None`, blank/whitespace, `'0'`, `'0.0'`, `#REF!`, `#VALUE!`, `#NAME?`, `#NULL!`, `#DIV/0!`, `#N/A`, `#NUM!`, `#ERROR!` (text, upper-cased) |
| Long-path threshold | 668, 743 | more than 218 characters ("Excel's typical limit") |
| Open waits | 698, 706 | 3 s, 2 s |

Time per report at the default `wait_time` 10, with no errors and Excel answering at once (READ, adding the sleeps): 5 + 1 + 2 + 10 + 5 + 10
+ 2 + 1 = 36 s in the refresh, + 3 s release = **about 39 s plus Excel start/open times**; formula `16 + w + max(10, w)` + 3. At w = 120:
about 259 s per report. A client with both reports takes twice that. The Stop button is checked only between clients
(`extraction_controller.py:61`). HIGH.

#### class ReportProcessor (21-752)

**`__init__(self)` - 24-28**: sets `excel_instance = None` (**never read anywhere** - dead attribute), `log_callback = None`,
`wait_time = 10`, `suffix_pattern = "_Refreshed_{timestamp}"`. Created at `extraction_controller.py:49` (one processor for the whole run).

**`_log(self, message, level='info')` - 30-37**: always `logger.info(message)` (ignores `level`); if a callback is set, calls
`callback("  " + message, level)`. All 96 `self._log(...)` calls in this file use the default level `'info'` (RAN grep), so in Tab 4 every
processor line is coloured as "info". The 23 `logger.warning/error/debug` calls (for example "Sheet not found" 522, "Excel busy, retrying"
202, "Error detected" 272) go only to the log file / console, **not to the Tab 4 log and not to `Extraction_Log_*.txt`**. Callers: every
step below.

**`process_client(self, client_data, log_callback=None)` - 39-70**
- `client_data` keys used: `name` (44), `process_itc` (51, default True), `process_sales` (61, default True), and `latest_version` (in
  `_process_report`). Built at `gui/pyqt_main_window.py:522-528`: `name` = `"<client folder> / <version folder>"`, `client_key` = client folder
  name (unused here), `latest_version` = version folder `Path`.
- Stores the callback (41). Result dict: `client`, `timestamp` (a `datetime`), `itc`, `sales`. For each report type: if selected ->
  `_process_report(...)`; else `{'status': {'success': False, 'error': 'Not selected for processing'}, 'data': {}, 'extracted_values': {}}`.
- Returns the dict. No exceptions caught here (`_process_report` catches its own). Caller: `extraction_controller.py:73`.

**`_process_report(self, client_data, report_type)` - 72-163**, step by step (READ; RAN `normal`, `skip`, `emptysuffix`, `openfail`)
1. 76-77: `glob("ITC_Report_*.xlsx")` (or Sales) in the version folder. 79-87: none -> status `"No ITC report found"`, empty data.
2. 89: **`original_file = report_files[0]`** - the first file the folder listing returns. The refreshed copies made by earlier runs also match
   this pattern (RAN: after one run the glob returned the original and `..._Refreshed_211026_0355.xlsx`). With the default suffix the original
   sorts first on NTFS (Windows' normal disk format lists names in sorted order, and `.` sorts before `_`) (MEDIUM). On other file systems
   (for example a network share from a Linux NAS) the order can be arbitrary, and with a custom suffix that starts with a character sorting
   before `.` (space, `-`, `(` ...) the copy sorts first: the next run would refresh a copy of the copy (`..._x_y.xlsx`). If a version folder
   ever held two original reports, only the first is used, silently. MEDIUM.
3. 92-94: copy name = `<stem><suffix>.xlsx`, suffix = the user's pattern with `{timestamp}` replaced by DDMMYY_HHMM. Always `.xlsx`.
4. 97-98: read `EXTRACTION_OPTIONS['skip_refresh']` (set from the Tab 4 checkbox).
5. 100-115 (skip mode): look for `<stem>_Refreshed_*.xlsx`; if any, take the newest by mtime, status `{'success': True, 'error': None,
   'validation_message': 'Skipped - Using existing refreshed file'}` - **success even if that earlier refresh had failed** (READ, HIGH); if
   none, log and fall through to a normal refresh. With a custom suffix (no `_Refreshed_`) skip mode never finds the copies and refreshes every
   time (RAN `skip`: three runs, three refreshes, the same-minute copy overwritten).
6. 118-137 (refresh): `shutil.copy2(original, copy)` (120) - overwrites a same-minute copy; the **original report is never modified** (only the
   copy is opened and saved) (READ, RAN). Then `_refresh_power_query_simple(copy)` (124). Lines 125-137 catch "could not open" errors and
   continue - but `_refresh_power_query_simple` catches every `Exception` itself (lines 339-344) and returns a status dict, so **lines 125-137
   are effectively dead** (RAN `openfail`: the error arrived as a returned status, the special message at 127-128 never appeared). HIGH.
7. 140-141: log "Waiting for file to be released...", sleep 3 s.
8. 144: `_extract_data(copy, report_type)` - **always, whatever the refresh status** (comment 143).
9. 146-152: return `{'status': refresh_status, 'data': values, 'extracted_values': values (the same dict object), 'original_file': str,
   'refreshed_file': str}`.
10. 154-163: any exception (for example copy failure, a suffix with characters Windows forbids, a network disconnect) -> logged with traceback,
    status `{'success': False, 'error': str(e)}`, empty data. RAN `emptysuffix`: an empty suffix makes the copy name equal to the original;
    `shutil.SameFileError` is caught here, nothing is damaged, the error text (full path) is shown as the status.

**`_refresh_power_query_simple(self, excel_file)` - 165-366** - how a report is refreshed through Excel COM, step by step (READ; order of
calls confirmed RAN with fakes)
1. 171: `pythoncom.CoInitialize()` (prepares COM for this worker thread).
2. 174: `excel = win32com.client.Dispatch("Excel.Application")` - **`Dispatch`, not `DispatchEx`**. MEDIUM (cannot test here): when the user
   already has Excel open, `Dispatch` commonly hands back that running Excel instead of starting a new one. Then the user's own Excel is made
   visible/maximised, receives the keystrokes, and is closed by `excel.Quit()` at 357 with `DisplayAlerts = False` (Excel would not ask to save
   the user's other open workbooks). Check on Windows: open Excel with an unsaved workbook, run one extraction, see whether that Excel closes.
   Standard 12 (lines 851-855) requires `DispatchEx` and "Never take over the user's own open Excel".
3. 178-190: up to 5 tries: `Visible = True` ("Keep it visible to avoid issues"), `DisplayAlerts = False`, `WindowState = xlMaximized`; on
   repeated failure logs "Could not set Excel properties, continuing anyway..." (RAN `setprop`).
4. 192-205: open the copy via `_safe_open_workbook` up to 3 times while the error text contains "Call was rejected by callee"; other errors
   re-raise to the outer handler (339). **If all 3 tries are rejected, `wb` stays `None` and the code carries on** (RAN `rejected`: up to 9
   `Workbooks.Open` calls, then keystrokes sent anyway, then `'NoneType' object has no attribute 'Save'` returned as the error).
5. 208-211: wait 5 s ("Loading... n/5 seconds").
6. 214-227: `shell = Dispatch("WScript.Shell")`; up to 3 x `shell.AppActivate(excel.Caption)` (bring the window whose title matches Excel's
   caption to the front). If it never succeeds: a warning in the log file only, **and the keystrokes are still sent** to whatever window is in
   front - possibly the user's e-mail or document (READ, HIGH; effect on real Windows MEDIUM). Whether `excel.Caption` matches the window
   title in current Excel versions is not verifiable here (LOW).
7. 234-249: `SendKeys("%a")`, 1 s, `SendKeys("r")`, 1 s, `SendKeys("a")`: Excel's keyboard route Data -> Refresh All -> Refresh All. The log
   says "Sending Alt key..." but `%a` is Alt+A in one chord (comment slightly off). Needs: Excel window focused, an unlocked screen with a
   logged-in user, nobody touching keyboard or mouse, ribbon key tips the same as on the author's PC (other Excel language/version could map
   A/R/A differently - MEDIUM).
8. 253-258: wait `wait_time` seconds. **The code never asks Excel whether the refresh finished**; it just waits.
9. 261-278: 5 checks, 1 s apart, `_find_error_window()`; on the first hit: read its text (`_get_window_text`), log it as a warning (file
   only), close it (`_close_error_window`), stop checking.
10. 281-286: extra wait `max(10, wait_time)` seconds (log every 2 s, shows 0/10, 2/10 ...).
11. 289-298: if no error yet, one more `_find_error_window()` check ("Late error").
12. 301: `_validate_refresh(wb)` -> (ok, message) - checks `Info!BB2`.
13. 304-307: **`wb.Save()` - always, even if an error dialog was found or BB2 failed** (RAN `errwin`, `no_info`). Then 2 s.
14. 310-313: if validation failed -> error = `"Refresh validation failed: <message>"`.
15. 316-319: `wb.Close(SaveChanges=True)` (saves again), 1 s.
16. 321-330: log result; return `{'success': not error, 'error': text or None, 'validation_message': message}`.
17. 332-337 (inner `except`, covers steps 6-16): any error -> `{'success': False, 'error': str(e)}` (no `validation_message` key).
18. 339-344 (outer `except`, covers steps 1-5): any error -> logged with traceback, `{'success': False, 'error': str(e)}`.
19. 345-366 `finally`: `wb.Close(SaveChanges=False)` again if `wb` is set (after step 15 the workbook is already closed, so this call fails
    and the error is swallowed - RAN with fakes: a second Close was made); `excel.Quit()` (swallowed errors); `pythoncom.CoUninitialize()`
    (swallowed errors). Each step is logged ("Cleaning up Excel instance..." etc.).
- Return shape: dict `success` / `error` / optionally `validation_message`. Never raises an `Exception` (only `BaseException` such as
  KeyboardInterrupt could escape). Caller: line 124 only.
- Comment line 251 `# Rest of your code remains same...` - a leftover from a pasted chat answer; the indentation after it puts the whole
  refresh, save and close inside the `try` started at 214 (READ).

**Values read back** happen in `_extract_data` (below) in a **separate, hidden Excel** after the refreshing Excel has quit. The refreshed
values therefore count only if Excel had finished the refresh before `wb.Save()` at step 13. Standard 12 (lines 860-867) records a test where
saving straight after `RefreshAll()` stored the old 2 rows instead of 60,000 new ones; here nothing waits for completion except the fixed
sleeps. MEDIUM risk that a slow query is saved half-done; the `Info!BB2` check may still pass if BB2 already had a value (MEDIUM - depends on the
template; question for the author).

**`_find_error_window(self)` - 368-404**
- Inner `callback(hwnd, windows)` (370-397): for every **top-level window on the whole desktop** (`win32gui.EnumWindows`), if visible and
  enabled and its lower-case title contains any of the 10 words (376-387) **and** its class is `#32770`, `NUIDialog` or `bosa_sdm_msword`,
  add it. Per-window errors swallowed (395-396). Always returns True (keep listing).
- 399-404: returns the first match or `None`; any error -> `None` (swallowed).
- Gaps (READ, HIGH; real-world frequency MEDIUM): not limited to the Excel process - another program's dialog titled e.g. "... error ..." or
  any dialog titled "Microsoft Excel" (including a harmless prompt in the user's own Excel) counts as a Power Query error and will be closed;
  RAN `errwin` with two fake dialogs (one titled "Microsoft Excel", one from "Some other app - error"): both qualified, the first was taken.
  `bosa_sdm_msword` is a Word-style dialog class (LOW: why it is listed is unclear).
- Callers: 268, 291.

**`_get_window_text(self, hwnd)` - 406-441**: title + text of child windows whose class contains `static` or `text` and whose text is longer
than 5 characters (`EnumChildWindows`), joined, whitespace collapsed, cut to 200 characters; empty -> `"Power Query error occurred"`; any
error -> same text (logged at debug). Callers: 271, 294. The text becomes the status `error` and ends up in run history.

**`_close_error_window(self, hwnd)` - 443-476**: walks the dialog's child windows; the first whose class contains `button` and whose
lower-case text contains `ok`, `close`, `yes` or `continue` gets `BM_CLICK` (`SendMessage`), and enumeration stops (callback returns False;
LOW: pywin32's `EnumChildWindows` is understood to ignore that return value, so it may not stop - harmless). If nothing was clicked:
`PostMessage(WM_CLOSE)`. Then 0.5 s. Errors -> debug log only. **Substring matching: `ok` also matches texts like "Look up" / "Book...";
`yes` means a "Yes" button is pressed on whatever question the dialog asked** (RAN `errwin`: the fake "&Yes" button was clicked). READ, HIGH.
Callers: 273, 296.

**`_extract_data(self, excel_file, report_type)` - 478-587** - how values are read back (READ; RAN with fakes)
1. 484: `CoInitialize()`. 487: `Dispatch("Excel.Application")` again (same Dispatch caveat), `Visible = False`, `DisplayAlerts = False` (no
   retries here; a "Call was rejected" at this point goes to 566 and returns `{}` - RAN `setprop`: status stayed success, data empty, the
   report then shows ✓ with `Missing` everywhere).
2. 492-494: logs "Opening file in read-only mode" / "File opened in read-only mode", but `_safe_open_workbook` opens it **read-write**
   (first attempt `Workbooks.Open(path)` with no `ReadOnly`; RAN fake log `Workbooks.Open ('...xlsx',) {}`). Comment disagrees with code.
   Nothing is saved (closed with `SaveChanges=False`), so no harm except a lock on the file while reading.
3. 497-505: `mappings = CELL_MAPPINGS[report_type]`; none -> warning and `{}`.
4. 507-552: for each input sheet: find it by exact name among `wb.Sheets` (515-519; case-sensitive, unlike `Info` at 597); not found ->
   every field of that sheet = `None` (521-527; warning in log file; the Tab 4 log shows "`<field> = None (sheet not found)`"). Found -> for each
   cell: `ws.Range(cell).Value` (535) stored twice: under the field name (536) and under the bare cell reference (539) "for calculations"; a
   read error -> both `None` (542-545). A sheet-level error -> its fields `None` (547-552).
5. 554-562: log the count and every `field: value` (these lines - client figures - go to Tab 4, `Extraction_Log_*.txt` in the Level 1 folder,
   and `gst_organizer.log`).
6. 564: return the dict (field names and cell references -> values). 566-568: any other error -> logged, `{}`.
7. 569-587 `finally`: `wb.Close(SaveChanges=False)`, `excel.Quit()`, `CoUninitialize()`, errors swallowed.
- `.Value` (not `.Value2`): Excel's `.Value` turns Currency-formatted cells into a currency type and date cells into dates (Excel object model
  behaviour; through pywin32 currency usually arrives as `Decimal` and dates as datetimes - MEDIUM). Error cells arrive as integers
  (MEDIUM). Effects on the report: see data_consolidator gaps.
- Cell-reference keys can collide across sheets of the same report type (two sheets both mapping `O1` would share one `O1` key). Not the case
  today (RAN: no duplicates within ITC). Latent.
- Caller: line 144.

**`_validate_refresh(self, wb)` - 589-636** - [GST RULE - when in doubt: decides whether a report counts as successfully refreshed and
therefore whether its figures are shown]
- Finds a sheet whose name lower-cased is `info` (596-599). Missing -> `(False, "Info sheet not found")`.
- Reads `Info!BB2` (607; the comment at 605 says it holds the client name). `None` -> `(False, "Cell BB2 is empty")`; blank after `strip()` ->
  `"Cell BB2 is blank"`; text `'0'` or `'0.0'` -> `"Cell BB2 contains 0"`; upper-cased text in the error list (626) ->
  `"Cell BB2 contains error: ..."`; otherwise `(True, "Validated - Client Name: <BB2 text>")` (632).
- Any exception -> `(False, "Validation error: ...")` (634-636). Called with `wb = None` (step 4 above) this is the path taken.
- Gaps: Excel error cells read through COM are usually integers, not text (MEDIUM) -> `str(-2146826281)` passes as a valid "client name"
  (simulated RAN `errint`: status success, message "Validated - Client Name: -2146826281"). Newer Excel error texts (`#SPILL!`, `#CALC!`,
  `#GETTING_DATA` ...) are not in the list; `#ERROR!` is not an Excel error value (harmless). `Decimal('0.00')` would pass (text `'0.00'`). Only
  "not empty" is tested - BB2 is never compared with the client the folder belongs to, and nothing writes BB2 in B's code (grep: the app writes
  file paths into the `Links` sheet, `core/excel_handler.py:133-153`), so BB2 is presumably filled by the template's own formulas or queries
  (MEDIUM - question for the author).
- The client name text in BB2 is written into the status / Tab 4 log / run history.
- Caller: 301.

**`_safe_open_workbook(self, excel, file_path)` - 638-747**
1. 640: absolute path (`Path.absolute()` uses the current folder only if the path is relative; here it is always absolute).
2. 643-645: missing -> log and `FileNotFoundError`.
3. 648-654: open the file in Python and read 1 byte (detects a file Windows will not let us read); errors logged and re-raised.
4. 657-660: if a workbook with the same full path is already open **in this Excel** -> return it (only possible if `Dispatch` attached to an
   existing Excel - suggests the author met that case; MEDIUM).
5. 665-680: path longer than 218 characters -> slashes to backslashes and the `\\?\` (or `\\?\UNC\`) long-path prefix. LOW: whether Excel's
   `Workbooks.Open` accepts that prefix is not verifiable here; the third attempt has a fallback without it.
6. 683-689: attempt 1 `excel.Workbooks.Open(file_str)` (no options). On error logs the code and text. LOW: line 687
   `getattr(e, 'args', [None])[0]` raises `IndexError` if an exception has empty `args`, which would skip attempts 2 and 3 (COM errors always
   carry args, so practically irrelevant).
7. 692-701: attempt 2: **close every workbook with the same file name in this Excel, without saving** (`wb.Close(SaveChanges=False)`), wait
   3 s, open again. If Excel was the user's own, an unsaved user workbook with the same name would lose its changes (LOW likelihood - names are
   time-stamped copies; READ).
8. 704-731: attempt 3: wait 2 s; try the 8.3 short name (`win32api.GetShortPathName`) if it differs; else, if the long-path prefix was added,
   open the plain path; else open with `Filename=..., UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, Notify=False`. **These
   safer options (Standard 12 line 857 asks for `IgnoreReadOnlyRecommended=True`) are used only on this last path**, and
   `workbook.ReadOnly` is never checked after any open (Standard 12 lines 857-859). READ, HIGH.
9. 732-747: all failed -> logs advice (file open elsewhere, corrupted, path too long, "OneDrive/SharePoint sync issue", long-path support)
   and re-raises.
- Callers: 197 (refresh), 493 (extract).

**`cleanup(self)` - 749-752**: does nothing (`pass`). **NO CALLERS FOUND in B (dead)**; it was called by the deleted Tk extractor window
(`47a1492:power_query_extractor/gui/extractor_window.py:832`). READ, HIGH.

#### Excel workbook touchpoints (all COM, all on the time-stamped copy, never on the original)

| Line | Action | Workbook |
|---|---|---|
| 120 | `shutil.copy2` (file copy, not Excel) | original -> copy |
| 174 | Dispatch Excel #1 (visible, maximised) | - |
| 197 -> 685 / 699 / 714 / 722 / 725 | `Workbooks.Open` | copy |
| 238-248 | SendKeys Alt+A, R, A (Refresh All) | copy (whatever window is in front) |
| 305 | `wb.Save()` | copy |
| 317 | `wb.Close(SaveChanges=True)` | copy |
| 350 | `wb.Close(SaveChanges=False)` (second close, fails silently) | copy |
| 357 | `excel.Quit()` | Excel #1 |
| 487 | Dispatch Excel #2 (hidden) | - |
| 493 -> 685 ... | `Workbooks.Open` (read-write, despite "read-only" log) | copy (or an older `_Refreshed_` copy in skip mode) |
| 535 | `Range(cell).Value` reads | copy |
| 573 | `wb.Close(SaveChanges=False)` | copy |
| 579 | `excel.Quit()` | Excel #2 |
| 697 | `Close(SaveChanges=False)` of any same-name workbook | any workbook open in that Excel |

#### Comparison with Standard 12 "Excel through COM" (rebuild_pipeline_FINAL.md lines 844-878) - facts only

| Standard 12 rule (line) | What B does | Line(s) |
|---|---|---|
| COM only for Power Query workbooks the app writes/refreshes (846-850) | COM for the refresh (yes) and also a second Excel just to read cells (the standard says reading a PQ workbook with openpyxl is fine, line 827) | 174, 487 |
| `DispatchEx`, `Visible = False`, `DisplayAlerts = False`; never take over the user's Excel (851-855) | `Dispatch`; refresh Excel `Visible = True` + maximised; extraction Excel hidden; `DisplayAlerts = False` in both | 174, 181-183, 487-489 |
| `excel.Quit()` in `finally` (856) | Yes, both | 355-360, 577-582 |
| Open with `IgnoreReadOnlyRecommended=True`; if `ReadOnly`, close and ask the user (857-859) | Only on the last-resort attempt; `ReadOnly` never checked | 725-731 |
| `BackgroundQuery = False` on each connection, `RefreshAll()`, `CalculateUntilAsyncQueriesDone()`, check no `Refreshing`, then save (860-867) | None of these; keystrokes + fixed sleeps + `Info!BB2` check, then save | 238-305 |
| Never use `RefreshDate` as proof (866-867) | Not used | - |
| Worker thread: `CoInitialize()` first, Excel work in a separate function, `CoUninitialize()` after it returns (872-875) | `CoInitialize()` first (yes); `CoUninitialize()` is called inside the same function's `finally` while `excel`, `wb`, `shell`, sheet objects are still referenced by local variables | 171/363, 484/584 |
| Excel installed, logged-in user, never a scheduled job (876-877) | Implied; SendKeys additionally needs the window in front and an unlocked screen | - |

#### Business rules in this file

- [GST RULE - when in doubt] Report files are found by name pattern `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` in the version folder (76);
  the first match is used (89).
- [GST RULE - when in doubt] A refresh "succeeds" when no error dialog was seen AND `Info!BB2` is non-empty, non-zero, not one of the listed
  error texts (262-330, 589-636).
- [GST RULE] Which cells are read: `CELL_MAPPINGS` (497-552) - see the mapping table. TIME-VARYING.
- Technical: refreshed copies are named `<stem><suffix>.xlsx`, default `_Refreshed_DDMMYY_HHMM` (92-94).

#### Edge cases explicitly handled

- No report found (79-87); report not selected (53-58, 63-68); skip mode with no existing copy -> refresh (112-115); Excel busy on open (201-
  203); properties cannot be set (185-190); window cannot be activated (229-230); error dialog (266-298); sheet missing (521-527); cell
  unreadable (542-545); Info sheet missing / BB2 empty, blank, 0 or error text (601-629); file missing or unreadable (643-654); file already
  open in this Excel (657-660); long path (668-680); three open strategies (683-731).

#### Gaps, likely failures and silent bugs (summary; details above)

- Unconditional pywin32 import stops the whole app (RAN). HIGH.
- Possible take-over of the user's Excel by `Dispatch` + `Quit()` (MEDIUM, Windows check needed).
- Keystrokes sent even without focus; user must not touch the PC during extraction; screen lock or another window popping up breaks the
  refresh silently (fixed waits, no completion check). HIGH (READ), real-world MEDIUM.
- Error-dialog scan covers all programs; clicks "Yes"/"OK" by substring. HIGH (READ).
- Workbook saved even after an error dialog / failed validation (RAN).
- `wb = None` path when Excel stays busy -> keystrokes into an Excel with no workbook, then a confusing `'NoneType' ... 'Save'` error (RAN).
- Lines 125-137 dead (RAN). `os` import, `excel_instance`, `cleanup()` dead.
- Skip mode reuses copies regardless of whether they refreshed successfully; custom suffix breaks skip mode (RAN).
- `report_files[0]` choice depends on folder-listing order (MEDIUM).
- Same-minute re-run overwrites the earlier copy (`copy2`), then refreshes it again - harmless; many runs leave many copies (no clean-up).
- Two copies of the app (or two users on one PC) running extractions at the same time: both type keystrokes into whichever Excel is in front,
  and `Dispatch` may give both the same Excel; one `Quit()` ends the other's work (MEDIUM).
- Network drive disconnect: copy, open or save fails -> caught, status error; nothing retried. A half-written copy may remain (LOW).
- `CoUninitialize()` while COM objects are still referenced (Standard 12 deviation); possible crash or hang when Python later releases them
  (LOW-MEDIUM, cannot test here).
- Thread safety: runs entirely inside the extraction worker thread; it touches no Qt widgets (only the callback, which emits a Qt signal -
  safe). The module-level `EXTRACTION_OPTIONS` is changed from this thread (`extraction_controller.py:55`) and read here; single worker, so OK.

#### Files, folders, logs read or written

- Reads: `<version folder>/ITC_Report_*.xlsx`, `Sales_Report_*.xlsx`, `<stem>_Refreshed_*.xlsx` (skip mode).
- Writes: `<version folder>/<stem><suffix>.xlsx` (copy, then saved by Excel). Log lines (logger -> handlers configured by `main.py:16-23`).
- Through the callback: every `_log` line ends up in Tab 4 and in `<Level 1>/Extraction_Log_<YYYYmmdd_HHMM>.txt`
  (`extraction_controller.py:109-121`), including client folder names, BB2 text and all extracted figures.

#### Comments, docstrings, history

- Comments that disagree with the code: 486/492/494 "read-only mode" (opens read-write); 237 "Sending Alt key" (sends Alt+A); 251 "Rest of
  your code remains same..." (pasted-chat leftover); 605 "client name" (only non-empty is checked).
- edac6a3: created. bd176bf (30-10-2025, shared): `process_itc` / `process_sales` selection, configurable `wait_time` (the old fixed loop had
  the comment "WAIT 7 SECONDS AS REQUESTED" while waiting 10) and `suffix_pattern`, `extended_wait = max(10, wait_time)`. c258047 (07-02-2026,
  B only): removed a stray `print(f"PROCESSING: ...")` placed above the docstring, removed the unreachable `return {'success': True ...}` after
  the inner try/except (the commit message calls it "unreachable return" - correct), narrowed 12 bare `except:` to `except Exception:` (RAN: `git show c258047` count), removed
  "# ADD THIS" comments. No behaviour change. Project A changed this file differently after the split (72 changed lines, plus a new 66-line
  `power_query_extractor/core/refresh_naming.py`, `git diff --stat 47a1492 A/main`) - for the comparison phase.

#### Pattern notes

- Error handling: every Excel step is wrapped; most failures become a returned status dict (never an exception to the caller); clean-up
  errors are swallowed silently. Logging: two channels (`logger` = file/console; `_log` = file/console + GUI) used inconsistently - warnings
  mostly only to the file.
- Waits are fixed sleeps everywhere; no polling of Excel state.

### power_query_extractor/extractor_main.py (79 lines) - [UI] (old stand-alone entry point; broken and dead)

- Purpose: was the stand-alone start file of the old Tkinter (Tk = Python's built-in window toolkit) Power Query Extractor window; in B it
  cannot even be imported. RAN, HIGH.
- Imports: `sys` (5, **unused**), `logging` (6), `json` (7), `Path` (8); project `from .gui.extractor_window import PowerQueryExtractorApp` (9).
  `power_query_extractor/gui/extractor_window.py` was **deleted in a62d3cc** (07-02-2026, "Remove old Tkinter GUI files replaced by PyQt5 in
  v4.0", 848 lines removed). RAN: `import power_query_extractor.extractor_main` and `python -m power_query_extractor.extractor_main` both
  stop at line 9 with `ModuleNotFoundError: No module named 'power_query_extractor.gui.extractor_window'`.
- Module-level side effects (12-19): `logging.basicConfig(level=INFO, format=..., handlers=[FileHandler('pq_extractor.log'), StreamHandler()])`
  - a log file in the **current folder**. Never reached in B because line 9 fails first (RAN: no `pq_extractor.log` was created). If it were
  imported after `main.py` had configured logging, `basicConfig` would do nothing (Python ignores a second `basicConfig`).
- Globals: `logger` (21).
- `class PowerQueryExtractor` (24-69):
  - `__init__(self, target_folder=None)` (27-34): keeps the folder; if empty, `load_cached_folder()`; logs it.
  - `load_cached_folder(self)` (36-60): tries, in order, `gst_organizer_cache.json`, `.cache/app_state.json`, `config/last_session.json`, **all
    relative to the current folder**; for the first existing one, `json.load` and return the first non-empty value of keys `target_folder`,
    `target_path`, `output_folder`, `last_target`; errors -> warning, try next; nothing -> `''`. None of these files is written by B (B writes
    `~/.gst_organizer/settings.json`, see cache_manager). The old Tk main app wrote `~/.gst_organizer_cache.json` (home folder, leading dot,
    `c258047^:gui/utils/cache_manager.py:15`), the old extractor window used `Path("gst_organizer_cache.json")` in the current folder
    (`47a1492:power_query_extractor/gui/extractor_window.py:28`) - a write/read mismatch that already existed before the split. READ, HIGH.
  - `run(self)` (62-69): `PowerQueryExtractorApp(folder).mainloop()` (Tk API); errors logged and re-raised.
- `main()` (72-75), `if __name__ == "__main__"` (78-79).
- Callers: **NO CALLERS FOUND (dead)**. Until c258047, `launch_extractor.py` imported `PowerQueryExtractor` from here (git show c258047).
- `.gitignore` still lists `pq_extractor.log` (41) and `config/last_session.json` (63), and README.md:164-171 / 290-291 still describe this
  extractor ("auto-loads target folder from cache", `pq_extractor.log`) - stale documentation (README has its own reader).
- History: edac6a3 created; unchanged in B. Project A changed it after the split (+3 lines) and still has `extractor_window.py`.
- Gaps: everything in it is unusable in B; delete or rewrite candidate (technical point).

### power_query_extractor/gui/__init__.py (1 line) - [UI]

- Purpose: package marker; docstring `"""GUI components for Power Query Extractor"""`; no code. The package is now **empty** (its only
  module `extractor_window.py` was deleted in a62d3cc) - a leftover. READ, HIGH. History: edac6a3, unchanged.

### launch_extractor.py (28 lines) - [UI] (second entry point)

- Purpose: opens the same main window as `main.py`, switched to Tab 4 ("Step 4: Extract && Refresh"). READ, RAN.
- Imports: `sys` (5), `os` (6); inside `main()`: `PyQt5.QtWidgets.QApplication` (13), `gui.pyqt_main_window.GSTOrganizerWindow` (14) - which
  pulls in pywin32 through the extraction controller (see report_processor).
- Line 8: `sys.path.append(<folder of this file>)` - **append** (lowest priority), unlike `main.py:29` which inserts at position 0. When run
  as `python launch_extractor.py` Python already puts the script folder first, so this line is redundant (READ, MEDIUM).
- `main()` (11-24): inside `try`: create `QApplication(sys.argv)`, the window, `window.tabs.setCurrentIndex(3)` (Tab 4) before `show()`,
  `sys.exit(app.exec_())`. `except Exception` -> `print(f"Failed to launch: {e}")`, `sys.exit(1)` (SystemExit is not an `Exception`, so the
  normal exit is not caught).
- Differences from `main.py` (READ, RAN): no `logging.basicConfig` -> module INFO messages are dropped and nothing is written to
  `gst_organizer.log` (RAN: the start folder stayed empty), WARNING and above go to the console only; no application name/version
  (`main.py:41-42`); no UTF-8 console wrapper; no special message for a missing module. Under `pythonw` the `print` shows nothing.
- RAN without pywin32: prints `Failed to launch: No module named 'pythoncom'`, exit 1. RAN with fake pywin32: window opens on tab index 3
  "Step 4: Extract && Refresh"; on close it wrote `~/.gst_organizer/settings.json` and `session_state.json` (`"active_tab": 3`).
- Callers: none (entry point). README.md:166-169 tells users to run it.
- History: edac6a3 created (Tk launcher that imported `power_query_extractor.extractor_main.PowerQueryExtractor` and showed a Tk error box on
  failure). c258047 rewrote it for PyQt5; the error box was replaced by `print()` - for users without a console this is a silent failure
  (READ, HIGH).
- Standard 13: same deviations as main.py (no `.bat`/`pythonw`, no excepthook, `print` for errors).

### models/__init__.py (4 lines) - [CORE]

- Purpose: re-exports (makes available directly from `models`) `AppSettings`, `ScanResult`, `ProcessingState` (line 2), `ClientRunRecord`,
  `RunHistory` (3), `CacheManager` (4). Docstring line 1 `"""Data models for GST File Organizer"""`. READ.
- `AppSignals` (models/signals.py) is **not** re-exported. Importing any `models.*` submodule runs this file, which imports `app_state`,
  `run_history`, `cache_manager` (no Qt) - so the models can be used without PyQt5 (RAN: `import models` works; `dir(models)` lists the six
  names plus the submodules).
- Callers of the re-exports: none - every user imports the submodules directly (`gui/pyqt_main_window.py:17-19`,
  `gui/controllers/processing_controller.py:12-13`, `gui/controllers/scan_controller.py:9`). Re-export layer unused (READ, HIGH).
- History: c258047 created; unchanged.

### models/app_state.py (88 lines) - [CORE]

- Purpose: three dataclasses holding the app's settings, the last scan result and processing flags (docstring 1-4: "Framework-agnostic ... state
  that was previously scattered across the Tkinter GSTOrganizerApp god object" - agrees). READ, RAN.
- Imports: `dataclass`, `field` (6); `Dict, List, Any, Optional` (7; `Optional` **unused**). No OS-specific modules.

**`class AppSettings` (10-46)** - 11 fields (RAN; the c258047 commit message says "12-field dataclass" - it is 11):

| Field | Line | Default | Filled from (window) | Used by |
|---|---|---|---|---|
| `source_folder` | 13 | `''` | Setup tab `source_folder` (`pyqt_main_window.py:197`) | scan (219), recent list (cache_manager 48-52) |
| `itc_template` | 14 | `''` | Setup (198) | processing (`processing_controller.py:119`), check at `pyqt_main_window.py:353` |
| `sales_template` | 15 | `''` | Setup (199) | processing (`processing_controller.py:127`), check at 356 |
| `target_folder` | 16 | `''` | Setup (200) | processing (`processing_controller.py:47`); Tab 4 tree root (`pyqt_main_window.py:421`) |
| `processing_mode` | 17 | `'fresh'` | Setup (201) | `FileOrganizer` (`processing_controller.py:48`), run history |
| `include_client_name_in_folders` | 18 | `False` | Setup checkbox (202) | `FileOrganizer` (49) |
| `client_name_max_length` | 19 | `35` | Setup spin box (203) | `FileOrganizer` (51) |
| `dark_mode` | 20 | `False` | View menu (595) | theme (66, 596) |
| `pq_wait_time` | 21 | `10` | Tab 4 spin box (205) | restored at 185; the extraction itself reads the spin box (544) |
| `pq_suffix_pattern` | 22 | `'_Refreshed_{timestamp}'` | Tab 4 text box (206) | restored at 186; extraction reads the box (545) |
| `pq_skip_refresh` | 23 | `False` | Tab 4 checkbox (207) | restored at 187; extraction reads the box (546) |

- `to_dict(self)` (25-39): returns the 11 fields as a plain dict, **listed by hand** (a new field must be added here too, while `from_dict`
  picks fields automatically - maintenance trap). Caller: `cache_manager.py:43`.
- `from_dict(cls, data)` (41-46): keeps only keys that are field names (44-45) and builds the object. **No type checks** (RAN:
  `{'pq_wait_time': '10', 'dark_mode': 'no'}` accepted as text; `'10'` would later be passed to `QSpinBox.setValue` at
  `pyqt_main_window.py:185`, which expects a number - MEDIUM that this raises at start-up; only possible with a hand-edited file). Caller:
  `cache_manager.py:61`.
- Instances created: `pyqt_main_window.py:39` (default, immediately replaced at 50-51), `:196` (`_collect_settings`), `cache_manager.py:62,
  109`.
- No GST rules. `client_name_max_length = 35` and `processing_mode = 'fresh'` are folder-naming / processing defaults (B-R1 covers their use).

**`class ScanResult` (49-74)**: fields `scanned_files` (dict), `client_data` (dict), `variations` (list), `statistics` (dict, added in b0e7d0e
09-02-2026), all with `default_factory` (no shared mutable defaults - good). Created empty at `pyqt_main_window.py:40`, filled at
`gui/controllers/scan_controller.py:32-37`, read at `pyqt_main_window.py:246-299, 335, 368`.
- Properties `total_files` (57-59), `total_clients` (61-63), `complete_clients` (65-68, counts `status == 'Complete'` as text): **NO CALLERS
  FOUND (dead)** - the window uses `statistics` from the parser instead. RAN: they compute correctly (1, 2, 1 on a test object).
- `clear(self)` (70-74): **NO CALLERS FOUND (dead)**; it would not clear `statistics` (RAN) - it was not updated when b0e7d0e added the field.

**`class ProcessingState` (77-88)**: `is_processing`, `stop_requested`, `selected_clients` - **never read or written anywhere outside this file
(dead fields)**; `client_folder_settings` (83) - passed to processing at `pyqt_main_window.py:370`, but **nothing ever fills it** (grep: no
writer in B; the validation view has no per-client option) -> always `{}`, so the per-client "include client name" override at
`core/file_organizer.py:126` is always False and only the global checkbox has an effect. READ, HIGH. The v3 Tk app did have a per-client
toggle that filled this dict (`c258047^:gui/handlers/client_handler.py:175-195` `toggle_folder_name_setting`, `:208-219` global update);
the v4 rewrite kept the slot but not the control - a lost feature (READ, HIGH).
- `reset(self)` (85-88): **NO CALLERS FOUND (dead)**; it would not reset `selected_clients` / `client_folder_settings` (RAN).

- Gaps: none that crash in normal use. Edge cases handled: unknown keys in saved settings are ignored (45).
- History: c258047 created; b0e7d0e added `statistics`.

### models/cache_manager.py (152 lines) - [CORE]

- Purpose: saves and loads settings, run history and "session" values as JSON files in a per-user folder, and imports settings from the v3
  (Tkinter) app's cache once. READ, RAN.
- Imports: `json`, `logging`, `Path`, `Any, Dict, Optional`; project `AppSettings` (11), `RunHistory` (12). No OS-specific modules.
- Constant `_DEFAULT_CACHE_DIR = Path.home() / '.gst_organizer'` (17) - on Windows normally `C:\Users\<user>\.gst_organizer` (MEDIUM: Python
  takes the home folder from the user profile). Standard 13 (lines 1036-1040) puts settings and logs in `%LOCALAPPDATA%\[App Name]\`.
- Files (all UTF-8 JSON, `indent=2`, non-JSON values turned into text by `default=str`):
  - `~/.gst_organizer/settings.json` (35) - keys = the 11 `AppSettings` fields.
  - `~/.gst_organizer/run_history.json` (36) - keys `runs`, `client_history`, `last_run_timestamp`, `total_runs` (see run_history).
  - `~/.gst_organizer/session_state.json` (37) - keys `recent_source_folders` (list, max 5), `active_tab` (int), `window_size` ([w, h]).
  - `~/.gst_organizer/.migrated` (99) - empty marker file.
  - reads `~/.gst_organizer_cache.json` (98) - the v3 app's file.

Functions (READ; RAN `t_models.py`, `t_start.py`):
- `__init__(self, cache_dir=None)` (31-37): folder = argument or default; **`mkdir(parents=True, exist_ok=True)` at construction**, no
  try/except - if the home folder is not writable the window constructor (`pyqt_main_window.py:42`) raises and the app does not start
  (`main.py:56-59` prints it). RAN: folder created at start-up even before anything is saved. Only caller: `pyqt_main_window.py:42` (default
  folder; the `cache_dir` parameter is used only by my tests).
- `save_settings(self, settings)` (41-55): `to_dict()`; load `session_state.json`; if `source_folder` is non-empty and **not already** in the
  recent list -> insert at the front, keep 5, save session; save `settings.json`. An existing entry is not moved to the front (RAN: after saving
  B,C,D,E,F,A the list was A,F,E,D,C). Callers: `pyqt_main_window.py:213` (`_save_settings`, which runs from File > Save Settings 102, a scan
  227, processing completed 381, dark-mode toggle 597, window close 603) and `cache_manager.py:119`.
- `load_settings(self)` (57-62): `{}` or missing -> defaults. Corrupt JSON -> warning in the log, defaults (RAN), and the next save
  **overwrites the corrupt file - the old settings are lost silently**. Caller: `pyqt_main_window.py:50`.
- `save_run_history` (66-68) -> `RunHistory.save`; callers `pyqt_main_window.py:380, 558`. `load_run_history` (70-72) -> `RunHistory.load`;
  caller `pyqt_main_window.py:43`.
- `save_session_state(self, state)` (76-80): load, merge (`update`), save. Callers: `pyqt_main_window.py:604` (on close: `active_tab`,
  `window_size`) and `:124` (migration).
- `load_session_state(self)` (82-84): **NO CALLERS FOUND (dead)** -> `active_tab` and `window_size` are written but never restored.
- `get_recent_source_folders(self)` (86-89): **NO CALLERS FOUND (dead)** -> the recent-folders list is maintained but never shown.
- `migrate_from_old_cache(self)` (93-131): returns at once if `~/.gst_organizer_cache.json` is missing or the `.migrated` marker exists (101-102).
  Otherwise: read the old file (105-106); build `AppSettings` from old keys `source_folder`, `itc_template`, `sales_template`, `target_folder`,
  `processing_mode` (default `'fresh'`), `include_client_name` -> `include_client_name_in_folders`, `client_name_max_length` (default 35),
  `dark_mode` (109-118) - these match what the v3 app wrote (`c258047^:gui/main_window.py:330-339`); save (119); copy
  `old_data['recent_source_folders']` (122-124) - **but the v3 app stored that list under `recent_folders`**
  (`c258047^:gui/utils/cache_manager.py:34, 43`), so the old list is never carried over (RAN: only the current source folder appeared, added by
  `save_settings`). Touch the marker (127). Any error -> warning, no marker (retried at next start). The v3 app replaced
  `client_name_max_length` 0 by 35 when loading (`c258047^:gui/main_window.py:353-354`); the migration copies 0 through (RAN). Because
  `_save_json` swallows write errors, the marker can be created even when `settings.json` was not written (LOW). The old file is left in place.
  Caller: `pyqt_main_window.py:47`, between `load_run_history` (43) and `load_settings` (50) - order correct.
- `_save_json(self, path, data)` (135-142): create the parent folder, overwrite the file. **Not atomic** (writes the file in place: a crash or
  power cut mid-write leaves a truncated file, which then loads as empty and is overwritten - see above). Errors -> `logger.error` only; the
  user is not told that settings were not saved.
- `_load_json(self, path)` (144-152): `{}` when missing or unreadable (warning). If a file holds a JSON list instead of an object, callers'
  `.get` would raise `AttributeError` (only with a hand-edited file; LOW). Such an error in a slot would abort the app (see main.py).

- Two copies of the app open at the same time (or one user on two PCs sharing a roaming profile): each holds its own copy in memory and the
  last one to save wins - the other's settings / run records are lost (RAN with two `CacheManager` objects: only the second instance's record
  survived). HIGH.
- Edge cases handled: missing files, corrupt files (defaults), missing parent folder (138).
- History: c258047 created (docstring line 3: "Replaces the old gui/utils/cache_manager.py" - correct; old file deleted in a62d3cc).

### models/run_history.py (148 lines) - [CORE]

- Purpose: keeps a history of processing runs and per-client records, and attaches extraction results, saved as JSON. READ, RAN.
- Imports: `json`, `logging`, `dataclass, field, asdict`, `datetime`, `Path`, `Dict, List, Any, Optional`.

**`class ClientRunRecord` (16-33)** - fields: `client_key`, `client_name`, `state`, `run_timestamp`, `processing_mode` (str, default `''`);
`files_organized` (0); `itc_report_created`, `sales_report_created` (False); `itc_report_path`, `sales_report_path`, `output_folder` (`''`);
`pq_refreshed` (False); `pq_refresh_timestamp` (`''`); `pq_extraction_data` (dict); `errors` (list). Created at
`gui/controllers/processing_controller.py:150-160` (fills key, name, state, run time, mode, files organised, both "created" flags,
`output_folder` = Level 2 folder) and at `run_history.py:88-94`. **`itc_report_path`, `sales_report_path` and `errors` are never set anywhere
(always empty - dead fields)**. READ, HIGH.

**`class RunHistory` (36-148)** - fields `runs` (list), `client_history` (dict: client key -> list of record dicts), `last_run_timestamp`,
`total_runs`.
- `add_run(self, run_summary, client_records)` (44-71): time stamp `YYYY-mm-dd HH:MM:SS`; append a run entry with 7 keys (`timestamp`,
  `total_clients`, `successful_clients`, `failed_clients`, `total_files`, `reports_generated`, `processing_mode` - all present in the processing
  summary, `processing_controller.py:58-75`); `total_runs += 1`; append each record (as a dict) under its `client_key`; keep only the last 50
  `runs` (69-71). **`client_history` is never trimmed** (RAN: 61 records kept after 61 runs) - the comment "prevent unbounded growth" covers
  only `runs`. Caller: `pyqt_main_window.py:379`.
- `add_extraction_result(self, client_key, extraction_data)` (73-95): if the key has records, **change the latest one in place**:
  `pq_refreshed = True` (**always True, even when the refresh failed** - the status is not looked at), `pq_refresh_timestamp`,
  `pq_extraction_data` = the whole result; else create a new record with `run_timestamp` = now. Caller: `pyqt_main_window.py:557`, which
  passes `result['client']` = `"<client folder> / <version folder>"` (`gui/views/extraction_view.py:384` -> `report_processor.py:44`). Processing
  records use `client_key` = `f"{client}-{state_code}"` (`core/file_parser.py:262`, passed at `processing_controller.py:151`). The two never
  match, so **extraction results always go into separate records and processing records never show `pq_refreshed`** (RAN: `client_history`
  keys `['ClientA-07', 'ClientA-07 / Version-211026 1000']`, processing record `pq_refreshed` False). A second extraction of the same version
  overwrites the first one's data in that separate record. READ+RAN, HIGH.
- `get_last_run_for_client` (97-100): last record or None; caller 104. `get_last_run_timestamp_for_client` (102-107): its `run_timestamp`
  or `''`; caller `gui/views/validation_view.py:384` ("Last Run" column). Unaffected by extraction because of the key mismatch.
- `to_dict` (109-116), `from_dict` (118-126; **no type checks**: RAN `{"runs": null}` loads, then the next `add_run` raises
  `AttributeError: 'NoneType' object has no attribute 'append'` - inside the processing-completed slot this would abort the app).
- `save(self, file_path)` (128-136): same in-place, non-atomic write as `_save_json`; errors logged only. Caller: `cache_manager.py:68`.
- `load(cls, file_path)` (138-148): corrupt -> warning, empty history (RAN) -> the next save **overwrites all old history** silently.
  Caller: `cache_manager.py:72`.
- Data kept (confidentiality, READ): every extraction result is stored whole: client folder / version names, a time stamp, both report
  statuses (including BB2 text and error-dialog text), all extracted figures twice (`data` and `extracted_values`), and full paths of the
  original and refreshed files - in each user's profile, indefinitely. Dates and other non-JSON values become text (`default=str`, RAN: the
  extraction `timestamp` was stored as text).
- Docstring 1-4 ("Stores per-client processing history, PQ refresh status, and extraction results") - true for storage, but the refresh status
  is never linked to the processing history (key mismatch).
- History: c258047 created; unchanged.

### models/signals.py (38 lines) - [CORE] (dead)

- Purpose (claimed by the docstring, lines 1-4 and 10): "Central signal hub ... All inter-component communication goes through these signals" /
  "replaces direct self.app.xxx references". **In fact the class is never imported or created anywhere** (grep `AppSignals`, `signals` across
  B: only this file). The controllers define their own signals (`extraction_controller.py:18-21, 127-131`, similarly the scan and processing
  controllers). Docstring disagrees with the code. READ, HIGH.
- Imports: `PyQt5.QtCore.QObject, pyqtSignal` (6) - unconditional, but irrelevant because nothing imports the module.
- `class AppSignals(QObject)` (9-38), signals: `settings_changed()` 13; `scan_started()` 16, `scan_progress(int, int, str)` 17,
  `scan_completed(object)` 18, `scan_error(str)` 19; `processing_started()` 22, `processing_progress(float, str)` 23,
  `processing_log(str, str)` 24, `processing_completed(dict)` 25, `processing_error(str)` 26, `processing_stopped()` 27;
  `extraction_started()` 30, `extraction_progress(float, str)` 31, `extraction_log(str, str)` 32, `extraction_completed(dict)` 33,
  `extraction_error(str)` 34; `status_message(str)` 37, `switch_tab(int)` 38. Several (`settings_changed`, `processing_stopped`,
  `status_message`, `switch_tab`) have no counterpart in any controller.
- History: c258047 created ("Signal/slot architecture replaces 305 self.app.xxx callback references") - planned design, never wired in.

### main.py (63 lines) - [UI] (main entry point)

- Purpose: start file of the PyQt5 app: sets up console encoding and logging, puts the project folder on Python's import path, opens the main
  window. READ, RAN.
- Imports: `sys`, `io`, `logging`, `Path` (6-9); inside `main()`: `PyQt5.QtWidgets.QApplication` (37), `gui.pyqt_main_window.GSTOrganizerWindow`
  (38). No OS-specific module directly, but line 38 imports pywin32 indirectly and unconditionally (report_processor.py:11-15).
- Module-level code (runs at import, before `main()`):
  - 12-13: if `sys.stdout` has a `buffer` (a real console), replace it by a UTF-8 text wrapper (so emoji in log lines can be printed). Under
    `pythonw` `sys.stdout` is `None`, the check is False and nothing happens (RAN simulation with `sys.stdout = None`: no crash). Good guard.
  - 16-23: `logging.basicConfig(level=INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', handlers=[FileHandler('gst_organizer.log'),
    StreamHandler(sys.stdout)])`.
    - `gst_organizer.log` is relative -> it is created/appended in the **current folder** (the folder the app was started from, e.g. a
      shortcut's "Start in" folder, or the shared app folder when main.py is double-clicked there) (RAN: created in the start folder). Never
      rotated (grows for ever). Contains client folder names, BB2 text and extracted figures (from report_processor `_log`) - confidentiality.
    - `FileHandler` gets no `encoding`, so it uses the PC's default text encoding. The code logs many non-ASCII symbols (✓ ⚠️ ❌ ✅ 📌 💡 📋 in
      report_processor, emoji in other modules). If the default is cp1252 (usual for English Windows, Python 3.14 is not in UTF-8 mode by
      default there - MEDIUM), each such line fails to be written and Python prints "--- Logging error ---" instead (RAN with a handler forced to
      cp1252: the `✓` line was missing from the file). Check on a Windows PC: open `gst_organizer.log` after an extraction and look for the
      ✓ lines.
    - `StreamHandler(None)` under pythonw falls back to `sys.stderr`, also `None`; Python's logging then skips the console silently (RAN
      simulation: no crash).
- `setup_environment()` (26-29): `sys.path.insert(0, <folder of main.py>)` (uses `__file__`, not the current folder - good). Caller: 35.
- `main()` (32-59): `setup_environment()`; import PyQt5 and the window; `QApplication(sys.argv)`; name "GST File Organizer", version "4.0"
  (41-42); create and show the window; `sys.exit(app.exec_())`.
  - `except ImportError` (49-54): logs, prints `"\n Missing module: <e>"`, `"\nPlease install required packages:"`, `"  pip install PyQt5
    openpyxl"`, exit 1. RAN without pywin32: the message names `pythoncom` but the advice does not mention pywin32. No `input()` pause any more,
    so if started by double-click with `python.exe` the console window closes at once and the message is unreadable; with `pythonw` nothing is
    shown at all.
  - `except Exception` (56-59): logs with traceback, prints `"\n Critical error: <e>"`, exit 1.
  - These handlers cover only start-up. Errors inside the running window (Qt slots) never reach them: **RAN - an uncaught exception in a
    PyQt5 slot aborts the process (exit code 134)**, printing the traceback to the console only. No `sys.excepthook` is installed, so there
    is no log file entry and no message box for such crashes.
- Lines 62-63: `if __name__ == "__main__": main()`.
- RAN with fake pywin32 (`t_start.py main.py`): window "GST File Organizer v4.0", 4 tabs `Step 1: Setup`, `Step 2: Validation`,
  `Step 3: Processing`, `Step 4: Extract && Refresh` (the `&&` shows as one `&` on screen), start-up log line
  `GST File Organizer v4.0 initialized`; on close it wrote `settings.json` and `session_state.json` (`active_tab 0`, `window_size [1400, 900]`).

#### Comparison with Standard 13 (rebuild_pipeline_FINAL.md lines 900-1061) - facts only

| Standard 13 rule (lines) | B | Evidence |
|---|---|---|
| Users double-click a shortcut to `Run [App Name].bat`, which starts `..\_runtime\python\pythonw.exe -B -s main.py` with `start "" /D` (951-975) | No `.bat`, no shared runtime; README tells users to run `python main.py` / `python launch_extractor.py` in a console (README.md:101, 111-112, 166-169) | file list; README |
| Shared `_runtime` + per-app `requirements.txt` with **exact** versions (913-915, 937-940) | Minimum versions only (`>=`), pywin32 commented out | requirements.txt:5, 8, 11, 14 |
| `main.py` first lines: `APP_DIR`, `LOG_DIR = %LOCALAPPDATA%\APP_NAME\logs`, `report_fatal_error` with log file + `MessageBoxW`, `sys.excepthook`, `threading.excepthook`, `sys.path.insert(0, APP_DIR)` (984-1029) | Only the `sys.path.insert` exists (inside `main()`, 26-29); no excepthooks, no message box; logging to the current folder | main.py:16-29 |
| Find the app's files from `APP_DIR`, never the current folder (1034) | Imports use `__file__` (good); the log file uses the current folder; `extractor_main.py:38-42` (dead) reads cache files from the current folder | main.py:20, 28 |
| No console: `print()` shows nothing; report problems in the window; log to `%LOCALAPPDATA%` (1035-1038) | Start-up errors only `print()`ed (51-53, 58); launch_extractor also `print()`s (23) | main.py:49-59; launch_extractor.py:22-24 |
| Worker threads send errors back to the window (1036-1038) | Yes for the extraction worker (`extraction_controller.py:36-41` -> `error` signal -> message box `pyqt_main_window.py:563-565`) | READ |
| Settings and logs in `%LOCALAPPDATA%\[App Name]\`, never in the shared app folder (1039-1040) | Settings in `~/.gst_organizer` (user profile root, per user - not shared); log in the current folder (may be the shared folder) | cache_manager.py:17; main.py:20 |
| No passwords in code (1041-1045) | None present | READ |
| Never a file several people write at once in the shared folder (1046-1047) | `gst_organizer.log` in the shared folder would be such a file if started from there | main.py:20 |
| `-B` so users' PCs write no `.pyc` into the shared folder (969) | Not used; `.gitignore` ignores `__pycache__/` | .gitignore:2 |
| Desktop windows with PySide6 (Standard 12 line 833) | PyQt5 5.15 (Qt 5) | requirements.txt:5 |

- Comments/docstrings: docstring 1-4 "v4.0 ... PyQt5 unified application" - agrees. Comment 11 "Set UTF-8 encoding for stdout" - agrees.
- History: edac6a3 (v3.0, Tk; printed an emoji banner, created folders `core`, `utils`, `gui`, `logs`, `temp` and missing `__init__.py` files at
  every start, and paused with `input("Press Enter to exit...")` on errors); 5ea77c2 (29-10-2025) removed the shebang line; c258047 rewrote it for
  PyQt5 (folder creation and `input()` removed, `hasattr` guard added). Project A changed main.py substantially after the split (95-line diff) -
  for the comparison phase.

### requirements.txt (14 lines) - [CONFIG]

| Line | Entry | Purpose (from code) | Imported where | Installed in venv-B (RAN) |
|---|---|---|---|---|
| 1-2 | comments: "GST File Organizer v4.0 Requirements", "Python 3.7 or higher required" | - | - | Python 3.14.6 used here; 3.7 not tested (LOW) |
| 4-5 | `PyQt5>=5.15.0` ("GUI framework") | all windows, threads, signals | `main.py:37`, `launch_extractor.py:13`, all `gui/*`, `models/signals.py:6` | 5.15.11 (+ PyQt5-Qt5 5.15.19, PyQt5_sip 12.20.0, pulled in automatically) |
| 7-8 | `openpyxl>=3.1.2` ("Core dependency for Excel operations") | consolidated report (`data_consolidator.py:8-10`), report creation in core (B-R1) | `data_consolidator.py`, `core/excel_handler.py` | 3.1.5 (+ et_xmlfile 2.0.0) |
| 10-11 | `lxml>=4.9.0` ("Optional but recommended for better performance") | never imported by B's code; openpyxl uses it automatically when present (RAN: openpyxl's traceback went through `lxml_write_cell`) | indirect | 6.1.3 |
| 13-14 | `# pywin32>=306` (commented out, "Windows-only: Excel COM automation for Power Query") | COM, window handling, short paths | `report_processor.py:11-15` (unconditional), `core/excel_handler.py:21-24` and `utils/helpers.py:476-477` (guarded) | NOT installed |

- **Dependency used in code but missing from the dependency file**: pywin32 (`pythoncom`, `win32com`, `win32api`, `win32con`, `win32gui`) is
  required for the app to start at all (RAN), yet it is commented out. `pip install -r requirements.txt` gives an app that cannot start.
  HIGH. (Also needs a working "WScript.Shell" COM object - Windows Script Host - for `report_processor.py:216`.)
- No pins (exact versions) - Standard 13 asks for exact versions. No pytest (B has no tests).
- README.md:104-108 shows different contents (`openpyxl>=3.0.0`, `pywin32>=305; sys_platform == 'win32'`, no PyQt5) - stale.
- History: edac6a3 created (v3.0: openpyxl, lxml, commented dev tools pytest/black/pylint); c258047 added PyQt5 and the commented pywin32 line,
  removed the dev-tool comments. A changed its own copy after the split (8-line diff) - for the comparison phase.

### .gitignore (82 lines) - [TOOLING]

Tells git which files never to track. READ every line; notes on how each group relates to B's real files:
- 1-22 Python build artefacts: `__pycache__/`, `*.py[cod]`, `*$py.class`, `*.so`, `.Python`, `build/`, `develop-eggs/`, `dist/`, `downloads/`,
  `eggs/`, `.eggs/`, `lib/`, `lib64/`, `parts/`, `sdist/`, `var/`, `wheels/`, `*.egg-info/`, `.installed.cfg`, `*.egg`, `MANIFEST` (standard
  template). `__pycache__` folders had been committed in edac6a3 and were removed in 01862a0 (29-10-2025) right after this file was added
  (1cb878b, same day).
- 24-28 virtual environments `venv/`, `env/`, `ENV/`, `.venv`. 30-36 editor/OS: `.vscode/`, `.idea/`, `*.swp`, `*.swo`, `*~`, `.DS_Store`.
- 38-46 project-specific: `*.log`, `gst_organizer.log` (written by main.py:20 when started from the project folder; a copy was committed in
  edac6a3, still present on the 6237597 side, and deleted in cbfa4a5 on the other side of merge 9c56a1c - file not opened), `pq_extractor.log` (only the dead extractor_main.py:16 would write it), `.gst_organizer_cache.json` (the v3
  cache lived in the home folder, not the project folder, so this line never matters - READ), `*.cache`, `.cache/` (extractor_main.py:40, dead),
  `temp/`, `logs/` (the v3 main.py created these folders; B no longer does).
- 48-53 Excel temporary/lock files `~$*.xlsx`, `~$*.xls`, `*.tmp`, `temp_gst_*`, `backup_*` (no code in B creates `temp_gst_*` or `backup_*`
  names - grep; possibly from core's temporary-file logic in older versions - LOW).
- 55-58 "Output directories (user data)": `/Annual Statement-*/`, `/output/`, `/target/` - protects client output if the target folder is the
  project folder itself. The consolidated report and `Extraction_Log_*.txt` live inside `Annual Statement-*` folders, so they are covered only
  in that case.
- 60-63 "User data and sensitive information": `/mnt/user-data/` (the path where a Claude.ai chat keeps uploaded files - a sign of AI-assisted
  development; MEDIUM inference), `*.bak`, `config/last_session.json` (read only by the dead extractor_main.py:41).
- 65-68 test data `test_files/`, `test_data/`, `*.test.xlsx` (no tests exist).
- 70-73 Windows `Thumbs.db`, `ehthumbs.db`, `Desktop.ini`. 75-79 Power Query / connection files `*.odc`, `*.iqy`, `*.dqy`, `*.xlquery`.
- 81-82 `#Claude` / `.claude` (added in c258047; `.claude/settings.local.json` had been committed in bd176bf and was removed from tracking in
  c258047). Last line has no line-break (RAN).
- Not ignored: `~/.gst_organizer/` lives outside the project, so nothing needed.

---

## Cross-file observations

### Shared files / settings / caches - write path vs read path

| Shared thing | Written at | Read at | Consistent? |
|---|---|---|---|
| `~/.gst_organizer/settings.json` (11 `AppSettings` keys) | `cache_manager.py:54` via `save_settings` (from `pyqt_main_window.py:213`; migration `cache_manager.py:119`) | `cache_manager.py:59` (`load_settings`, from `pyqt_main_window.py:50`) -> `app_state.py:44-46` | Yes - same path, keys from `to_dict` (27-39) match the field names `from_dict` accepts (RAN round-trip) |
| `~/.gst_organizer/session_state.json` `recent_source_folders` | `cache_manager.py:49-52`, `:124` | `cache_manager.py:47` (internal), `:89` (`get_recent_source_folders` - no callers) | Written, never used by the app |
| `~/.gst_organizer/session_state.json` `active_tab`, `window_size` | `pyqt_main_window.py:604-607` -> `cache_manager.py:76-80` | `cache_manager.py:84` (`load_session_state` - no callers) | Write-only |
| `~/.gst_organizer/run_history.json` | `run_history.py:132-133` via `cache_manager.py:68` (from `pyqt_main_window.py:380` processing, `:558` extraction) | `run_history.py:143-145` via `cache_manager.py:72` (`pyqt_main_window.py:43`); records read by `validation_view.py:384` | Format consistent; **keys inconsistent**: processing key `"<client>-<state code>"` (`file_parser.py:262`, `processing_controller.py:151`) vs extraction key `"<client folder> / <version folder>"` (`extraction_view.py:384` -> `report_processor.py:44` -> `pyqt_main_window.py:556`) (RAN) |
| `~/.gst_organizer/.migrated` | `cache_manager.py:127` | `cache_manager.py:101` | Yes |
| `~/.gst_organizer_cache.json` (v3) | v3 app (deleted): `c258047^:gui/utils/cache_manager.py:45-46`, keys from `c258047^:gui/main_window.py:330-339` + `recent_folders` | `cache_manager.py:98-124` | Settings keys match; recent list key **mismatch** (`recent_folders` written vs `recent_source_folders` read) (RAN) |
| `gst_organizer_cache.json`, `.cache/app_state.json`, `config/last_session.json` (current folder) | nobody in B (old extractor window used the first, in the current folder) | `extractor_main.py:38-55` (dead) | Mismatch; dead |
| `gst_organizer.log` (current folder) | `main.py:20` (all modules' loggers) | people (README.md:290, 333) | launch_extractor.py writes no log at all (RAN) |
| `pq_extractor.log` | `extractor_main.py:16` (never reached) | README.md:291, `.gitignore:41` | Never created in B (RAN) |
| `<Level 1>/Extraction_Log_<YYYYmmdd_HHMM>.txt` | `extraction_controller.py:109-121` (GUI-callback lines only) | people | Warnings logged only through `logger` in report_processor are not in it |
| `<Level 1>/PQ_Extraction_Report_<YYYYmmdd_HHMM>.xlsx` | `data_consolidator.py:163-167` | people | Same-minute overwrite (RAN) |
| Report files `ITC_Report_*.xlsx` / `Sales_Report_*.xlsx` | `processing_controller.py:116-130` (names `ITC_Report_<client>_<state code>_<DDMMYY_HHMM>` at 120 and `Sales_Report_...` at 128, created by core/excel_handler - B-R1) | `report_processor.py:76-77`; Tab 4 tree `pyqt_main_window.py:461-462` | Yes; but the refreshed copies also match the pattern (RAN) |
| Refreshed copies `<stem><suffix>.xlsx` | `report_processor.py:92-94, 120` (user suffix, default `_Refreshed_DDMMYY_HHMM`) | skip mode `report_processor.py:102` (`<stem>_Refreshed_*.xlsx`); Tab 4 status `pyqt_main_window.py:467-479` (`*_Refreshed*` with .xlsx/.xls/.xlsm, not `~$`) | **Mismatch with a custom suffix** (RAN). Also the tree shows "Refreshed (date)" for any copy, even one whose refresh failed (copy exists, READ, HIGH) |
| `EXTRACTION_OPTIONS['skip_refresh']` (in memory) | `extraction_controller.py:55` (worker thread) | `report_processor.py:98` (same thread) | Yes; the file default is never effective |
| `CELL_MAPPINGS` (in memory) | `cell_mappings.py:264-278` | `report_processor.py:497`; `data_consolidator.py:243` (dead) | Yes |
| `EXTRACTION_CONFIG` (in memory) | `cell_mappings.py:33-242` | `data_consolidator.py:45, 233` | Yes, but see the substring bug |
| `Info!BB2` in each report | not written by B's code (grep); presumably the template | `report_processor.py:607` | Unknown writer - question for the author |
| PQ settings `pq_wait_time`, `pq_suffix_pattern`, `pq_skip_refresh` | saved from Tab 4 widgets (`pyqt_main_window.py:205-207`) | restored into the widgets (185-187); the extraction reads the widgets directly (544-546) | Yes |

### Other cross-file observations

- **Windows-module import pattern is inconsistent** (Audit method step 8): guarded in `core/excel_handler.py:20-27` and `utils/helpers.py:476-477`;
  unguarded in `report_processor.py:11-15`; and `os.startfile` is guarded by a platform check in `pyqt_main_window.py:407-408`. The unguarded
  one decides whether the app starts (RAN). HIGH.
- **"Latest" Level 1 folder for Tab 4 is chosen by name** (`pyqt_main_window.py:432-436`, sorted descending), and folder names use DDMMYY
  (`utils/constants.py:82` `'Annual Statement-{timestamp}'`, `utils/helpers.py:23-25` `%d%m%y %H%M`). Name order is not date order across
  months/years (e.g. a 31-12 folder sorts after a 01-01 folder of the next year), so Tab 4 can show an older Level 1 folder as "latest". READ,
  HIGH (owner: B-R3's file; noted because it feeds my extraction code).
- **Time-stamp formats differ across the feature**: copies `DDMMYY_HHMM` (`report_processor.py:92`), consolidated report and extraction log
  `YYYYmmdd_HHMM` (`data_consolidator.py:163`, `extraction_controller.py:112`), run history `YYYY-mm-dd HH:MM:SS` (`run_history.py:46, 75`),
  summary sheet `YYYY-mm-dd HH:MM:SS` (`data_consolidator.py:218`), tree status `dd/mm/yy HH:MM` (`pyqt_main_window.py:478`), Level 1 / version
  folders `DDMMYY HHMM`. READ.
- **Error reporting styles differ**: `print()` at start-up (main.py, launch_extractor.py), message boxes in the window
  (`pyqt_main_window.py:221-565`), status dicts in report_processor, re-raise in data_consolidator, log-only in cache_manager/run_history.
- **Unicode symbols in log text + log file without explicit encoding** (main.py:20) - see main.py (MEDIUM on Windows).
- **The Tab 4 log is HTML**: `extraction_view.py:402-418` (`log_message`) inserts each message with `insertHtml` (414), so the `\n` at the start of many
  report_processor messages (508, 554, 570) does not start a new line, and a `<` or `&` in a folder name would be treated as HTML (owner B-R4;
  LOW).
- **Run-time per client** (~80 s with both reports at default wait) and **Stop only between clients** (`extraction_controller.py:61`).
- **Dead code found in my files**: `power_query_extractor/extractor_main.py` (whole file, broken import), `power_query_extractor/gui/` (empty
  package), `power_query_extractor/__init__.py:3 __version__`, `cell_mappings.py` options `auto_refresh_if_missing`, `show_timestamp`, the
  `description` feature, `FIELD_DISPLAY_NAMES` (used only by dead code), `data_consolidator.py` `_write_section`, `_create_details_sheet`,
  unused import `Path` (6), unused parameter `workbook` (176), `_calculate_value` in practice; `report_processor.py` unused import `os` (5),
  `excel_instance` (25), `cleanup()` (749-752), lines 125-137 (effectively unreachable); `models/signals.py` (whole file);
  `models/__init__.py` re-exports (unused); `app_state.py` `Optional` import, `ScanResult.total_files/total_clients/complete_clients/clear`,
  `ProcessingState.is_processing/stop_requested/selected_clients/reset`, `client_folder_settings` never filled; `cache_manager.py`
  `load_session_state`, `get_recent_source_folders`; `run_history.py` fields `itc_report_path`, `sales_report_path`, `errors`.
- **Project A has the same consolidator substring line** (`A/main:power_query_extractor/core/data_consolidator.py:143`) and the same
  `cell_mappings.py` (unchanged since 47a1492 in both) - relevant to the comparison phase.
- **AI-assisted commits**: c258047, a62d3cc and c21f788 carry "Co-Authored-By: Claude Opus 4.6" in their messages (the whole v4.0 PyQt5
  rewrite of B was done with an AI agent; `.gitignore:60-61, 81-82` also point to AI-assisted work). READ.
- **Commit-message claims checked**: c258047 "AppSettings: 12-field dataclass" - 11 fields (RAN); "unreachable return in report_processor.py" -
  correct; "Signal/slot architecture replaces 305 self.app.xxx callback references" - true for the controllers' own signals, but the
  `models/signals.py` hub it created is unused; "Run history displayed in client tree ('Last Run' column)" - only processing runs appear
  (extraction key mismatch).

## Possible client names

- None found. Checked: the current content of my 17 assigned files; the diffs I read (c258047 for my files, 47a1492 for `cell_mappings.py`,
  bd176bf for `report_processor.py`, b0e7d0e for `app_state.py`); the deleted `power_query_extractor/config/cell_mappings copy.py` (47a1492^);
  the old v3 cache code (`c258047^:gui/utils/cache_manager.py`, `c258047^:gui/main_window.py:328-365`); and all 17 B/main commit messages
  (full bodies). The only names present are the author's (git author / repository owner) and generic examples. `gst_organizer.log` was not
  opened.
- Runtime note (not code): the consolidated report, `Extraction_Log_*.txt`, `gst_organizer.log` and `~/.gst_organizer/run_history.json` will
  contain real client folder names and figures when the app is used.

## Questions for the author

1. Output sheets `IMS Reco`, `Extra In Purchase`, `Extra In GSTR2B`: the configuration says their IGST/CGST/SGST come from `IMS_Reco`
   AA1/AD1/AG1, `ExtraInPurchase` O1/P1/Q1 and `ExtraInGSTR2B` S1/T1/U1, but the code writes the `Diff GSTR-3B` J15/K15/L15 figures there
   (RAN). Have you noticed identical figures in these sheets? Which cells do you want (please confirm the whole 22-row mapping table for the
   current templates - TIME-VARYING)?
2. `Info!BB2`: what fills this cell in the ITC and Sales templates (a formula, a Power Query result)? Is "BB2 not empty" the test you want for
   "refresh succeeded", and should it be compared with the client the folder belongs to?
3. When a refresh fails or BB2 is empty, the report shows `Missing` for every figure although the figures were read. Do you want to see the
   figures with a warning instead, or is hiding them intended?
4. During extraction Excel opens visibly and the app types Alt+A, R, A into it. Do users know not to touch the PC for the ~40 s per report?
   Which Excel version and language do the office PCs use (the key route can differ)?
5. Have you seen the app close your own open Excel during an extraction (possible effect of `Dispatch` + `Quit`)?
6. Is pywin32 installed on every PC that runs B? Without it the app does not open at all.
7. Is it acceptable that client folder names and extracted figures are kept indefinitely in `C:\Users\<user>\.gst_organizer\run_history.json`,
   in `gst_organizer.log` in the start folder, and in `Extraction_Log_*.txt` in the client output folder?
8. Does anyone change the "suffix" box on Tab 4? A suffix without `_Refreshed_` breaks "Skip refresh" and the "Refreshed" status.
9. `auto_refresh_if_missing` and `show_timestamp` in `cell_mappings.py` do nothing. Were these planned features?
10. Is `launch_extractor.py` (or the old stand-alone extractor, `extractor_main.py`) still used by anyone?
11. v3 let users switch "include client name in folder names" per client in the client tree; v4 keeps a place for that choice
    (`ProcessingState.client_folder_settings`) but has no control to set it, so only the global checkbox works. Was dropping the per-client
    choice intended? (Also: the v3 recent-folder list is not carried over by the migration.)
12. The "Last Run" column in the Validation tab shows processing runs only; should it also show the last extraction/refresh?

## Not verifiable on this Linux machine (how to check on a Windows PC with Excel)

- Whether `win32com.client.Dispatch("Excel.Application")` attaches to an already open Excel (open Excel with an unsaved workbook, run one
  extraction, watch whether that Excel is closed).
- Whether SendKeys Alt+A, R, A triggers "Refresh All" in the office Excel version/language, and what happens if another window has focus.
- How error cells, currency cells and date cells come back from `Range.Value` (put `=NA()`, a Currency-formatted number and a date in mapped
  cells; look at the consolidated report and at `run_history.json`).
- Whether `AppActivate(excel.Caption)` finds the window; whether Excel accepts `\\?\` long-path prefixes; behaviour of `CoUninitialize()`
  with live COM objects; whether `gst_organizer.log` loses the ✓ lines (default file encoding).
- The fake-COM runs prove only the Python control flow.

## INCOMPLETE

- Nothing missing from the assigned reading: all 17 files were read in full. The Windows-only behaviours listed above could only be traced by
  reading and simulation, as the task anticipated.
