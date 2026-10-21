# H-A1 - History notes: A-only commits, part 1 of 3 (fc279af .. ba2a753)

Task: read every line of the twelve commits that exist only in project A
(goyal-sidhant/AnnualStatement_oldv1) after the split from project B at commit 47a1492
(30-10-2025), oldest first: fc279af, f7eb8f1, a603c02, 12d505f, 275d9a4, 5c8ce95, c1583d6,
0890022, 6c1e1c6, 33d0baf, 064c634, ba2a753.

## How these notes were made

- Source of truth: the git repository /root/merge-scratch/combined, ref A/main. For every commit
  I ran `git show --stat` and read the full `git show` output (saved under
  /root/merge-scratch/run/H-A1/<hash>.diff). For f7eb8f1 (a near-total rewrite of one Markdown
  file, 1,862 diff lines, only 166 lines unchanged) I read the complete old version (856 lines,
  from fc279af) and the complete new version (1,182 lines, from f7eb8f1) with the Read tool;
  together they contain every removed and every added line of that diff.
- For each code change I also read the surrounding code at the commit (not only the changed
  lines) and the same code at A's HEAD (bf2f1dd), and where useful at B's HEAD (b0e7d0e).
- All 12 commits: confirmed RAN that each is an ancestor of A/main and not of B/main, and that
  none is a merge commit (each has one parent). So every heading below is labelled
  "A only (after the split)".
- Dates/times are the author dates, Indian time (+0530).
- Paths are repository-relative. Line numbers "at HEAD" are A's HEAD unless "B" is said.
- Marks: READ = from reading code/diff; RAN = I executed something (scratch copies under
  /root/merge-scratch/run/H-A1/, Python 3.14 from venv-A, `-B`, nothing written into A-copy or
  B-copy). Inferences carry HIGH / MEDIUM / LOW confidence.
- [GST RULE] marks anything that decides a tax figure, return/table/sheet/cell, rate, limit,
  period, state-code logic and similar. For those I only describe what the code does.
  TIME-VARYING marks values that change over time.
- No client data is copied here. Example names in these notes are placeholders (<Client A>, ABC,
  CA) or the generic examples the code/tests themselves use (ABC, XYZ).

Plain-English terms used below (explained once here):
- commit = one saved change in the project history; diff = the exact lines added (+) and removed (-).
- regex / pattern = a text rule the program uses to recognise file names.
- COM = Windows' mechanism that lets Python remote-control Excel ("Excel through COM").
- openpyxl = a Python library that reads/writes .xlsx files without Excel; it cannot run or keep
  Power Query (Excel's built-in data-import/refresh feature) when it saves a workbook.
- glob = "list the files whose names match *.xlsx"; os.scandir = "list everything in this folder
  once", returning each entry with some details already attached.
- stat = asking the file system for a file's size/dates/type; round-trip = one request to the
  file server and its answer (slow on a network share, a folder on another computer or on a
  server, sometimes reached by a "UNC" path like \\server\share).
- mtime = a file's or folder's "last modified" time.
- cache file = a small settings file the app writes so it remembers your last choices.
- lambda / closure = a tiny function created on the fly that "remembers" variables; "late
  binding" = it reads the variable's value when it finally runs, not when it was created.
- Tk event loop = the part of a Tkinter (the GUI toolkit A uses) program that waits for clicks
  and runs scheduled screen updates; worker thread = a second line of execution that does the
  long processing while the window stays responsive.
- pytest = the testing tool; characterization test = a test that records what the code does
  today so that later changes cannot silently alter it; fixture = a helper that pytest hands to
  tests.
- re-entrancy = the same action being started a second time while the first is still running.
- word boundary (\b in a regex) = the edge between a letter/digit/underscore and anything else.
- logger levels: INFO messages go to the log file at the app's current setting; DEBUG messages
  are hidden unless someone turns on debug logging.
- environment marker (`; sys_platform == 'win32'`) = a note in requirements.txt meaning
  "install this only on Windows".
- docstring = the explanatory text at the top of a Python function.

Environment observation (not caused by me): /root/merge-scratch/A-copy/tests/__pycache__/ exists
(compiled test files dated 21-10-2026 03:21 UTC). It is not part of the repository (RAN:
`git ls-tree -r A/main` lists no __pycache__). It does not change any line numbers.

---

## Commit-by-commit notes (oldest first)

#### fc279af - 04-03-2026 22:16 - goyal-sidhant - A only (after the split)

- Message: "Project Audit" (no body, no Co-Authored-By trailer).
- Files touched:
  - PROJECT_AUDIT_ANNUALSTATEMENT.md +856/-0 (new file). READ in full.
    What it is: an AI-written "Project Archaeology" audit of A as it stood at 47a1492. Its header
    says: audit date 2026-03-04, "Auditor: Claude Opus 4.6 (AI-assisted reverse engineering)",
    project root `d:\OneDrive 2\OneDrive\Python 2026\AnnualStatement`, "Current Version: v3.0
    (Production Ready) - Tkinter GUI", "Active Branch: main at commit 47a1492".
    Contents, section by section (the other agent checks these claims against the code; I only
    record what the document says and what later changed):
    - S1 Executive summary: tool for CA/GST consultants; Python 3.7+, Tkinter, openpyxl,
      win32com/pywin32; two entry points main.py and launch_extractor.py; "actively maintained
      and production-used"; a PyQt5 migration "recently attempted (2026-02-07) on a separate
      branch but not merged".
    - S2 Structure map (annotated tree of core/, utils/, gui/, power_query_extractor/); "Dead/Orphan
      Files: None detected"; notes `cell_mappings copy.py` existed until 47a1492.
    - S3 Feature inventory F1-F15: F1 scanning/pattern matching, F2 completeness analysis, F3
      folder structure, F4 ITC/Sales report generation, F5 processing summary report, F6
      organization text report, F7 Power Query refresh (SendKeys Alt+A,R,A; Info!BB2 check), F8
      cell value extraction (9 configurations), F9 consolidated extraction report, F10 session
      cache at ~/.gst_organizer_cache.json, F11 dark mode, F12 client list export, F13 dry run,
      F14 state name to code conversion, F15 configurable client folder names (default 35,
      range 15-100, warning above 10 characters).
    - S4 Data flow diagrams (main app and extractor).
    - S5 Business rules: the six file-name regexes with generic examples ("ABCLtd"); category
      routing; completeness rules (GSTR-3B exempt from duplicate check); ITC/Sales template
      "Links" sheet cell map (ITC B2,B4,B5,B7,B8,B10,B11; Sales B2,B3,B5,B6,B8,B9);
      a table of the Power Query Extractor's cell mappings from power_query_extractor/config/
      cell_mappings.py [GST RULE] TIME-VARYING (these name GSTR-9 tables): "GSTR-3B Difference"
      from sheet 'Diff GSTR-3B' J15/K15/L15; "GSTR-1 Difference" from 'Diff - GSTR-1'
      N1/O1/P1/Q1; "Turnover GSTR-1" from 'Diff - GSTR-1' I1; "Reco Diff" from 'Reco'
      C37/D37/E37; "Table 6J Match" from 'T6 of R9' K29; "Table 8D Limits" from 'T8 of R9' G6;
      "IMS Reco" from 'IMS_Reco' AA1/AD1/AG1; "Extra In Purchase" from 'ExtraInPurchase'
      O1/P1/Q1; "Extra In GSTR2B" from 'ExtraInGSTR2B' S1/T1/U1 (as the document states them;
      the code itself is not touched in my range); sanitization rules; client-state key rules;
      Excel validation rules (extensions, >=1,024 bytes, PK or OLE signature); refresh
      validation (Info!BB2 not blank/"0"/error); long-path handling (218 characters, 8.3 short
      paths, \\?\ prefix).
    - S6 Edge cases handled, rated "MODERATE".
    - S7 Gaps: no concurrency protection, no disk-space check, shallow template check, "No GSTIN
      validation", network disconnection, Excel zombie, MAX_FILE_SIZE_MB never enforced, regex
      backtracking, no subfolder scan, lower-case glob on Linux, Unicode normalisation, hyphens
      in client names, sequential processing, no progress persistence, SendKeys fragility,
      directory traversal, os.startfile, plaintext cache, no undo, per-client progress, excessive
      [DEBUG] logging.
    - S8 Design decisions (Why Tkinter, why two apps, why COM, why keyboard simulation for
      refresh (MEDIUM), why state codes, why copy, why keep hyphens, why max length, "evidence of
      iteration" including the "6 numbered debug checkpoints" in the Sales report code).
    - S9 Dependencies (openpyxl >=3.1.2, lxml >=4.9.0, pywin32 "not in requirements.txt").
    - S10 UI description of all tabs of both windows.
    - S11 Reconstruction blueprint, including "Known Improvements": QThread/PyQt5, COM
      RefreshAll() instead of SendKeys, recursive scanning, add GSTIN validation, progress
      persistence, "Separate the PQ Extractor into the main app as a fourth tab", structured
      config directory, disk-space check, hyphenated client names.
    - S12 Changelog archaeology of the 12 shared commits plus three "unmerged" PyQt5 commits
      (c258047, a62d3cc, c21f788) said to sit on branch `copilot-worktree-2026-03-04T11-52-56`.
- Why: the author ran his own "Project Archaeology" method (an earlier form of the method now kept in
  /home/user/Annual_Statement_2.0/.build-docs/toolkit/Project_Audit_FINAL.md; MEDIUM that it was
  an earlier version - the 12-section layout matches) on A to document
  it, roughly four months after the last code change. HIGH (the document follows that method's
  12-section layout).
- What it reveals: on 04-03-2026 the author treated A (Tkinter) as the production line and B's
  PyQt5 commits as an unmerged experiment (the document says so twice). The commits c258047,
  a62d3cc and c21f788 that the document calls "unmerged branch" are, in the combined repository,
  B's first three commits after the split (RAN: `git log 47a1492..B/main`). B/main also has two
  more commits (7e09e5c 08-02-2026, b0e7d0e 09-02-2026) that the document does not mention.
  LOW-MEDIUM: the local branch the AI saw may have had only those three commits.
- Type: docs.
- Rule changes: none in code. The document records [GST RULE] cell mappings (above) as they
  stood at 47a1492.
- Abandoned or reversed: the whole file is rewritten one day later by f7eb8f1.
- Dependency changes: none.
- Discrepancies between message and diff: none (message is one word). Note: the document says
  "Auditor: Claude Opus 4.6" but the commit has no Co-Authored-By trailer (READ).

#### f7eb8f1 - 05-03-2026 08:44 - goyal-sidhant - A only (after the split)

- Message: "Project Audit 2.0" (no body, no trailer).
- Files touched:
  - PROJECT_AUDIT_ANNUALSTATEMENT.md +1003/-677 (file grows 856 -> 1,182 lines). Only 166 lines
    are unchanged, so this is a rewrite, not an edit. READ both full versions.
    The file is identical at A's HEAD (RAN: `git diff f7eb8f1 A/main -- PROJECT_AUDIT_ANNUALSTATEMENT.md`
    is empty), so the HEAD copy still describes the code as of 47a1492/fc279af.

  What the revision ADDED (READ):
  - Header: audit date 2026-03-05; "Active Branch: main at commit fc279af"; new line "Unmerged
    Branch: copilot-worktree-2026-03-04T11-52-56 - attempted PyQt5 migration (3 commits ahead of
    main, NOT merged)".
  - S1: a new paragraph "Current state: Functional but with significant issues" (extractor
    imports pythoncom/win32com without fallback so it "will crash on import on non-Windows";
    stale version reference; requirements.txt omits pywin32). The PyQt5 attempt is now dated
    "2026-03-04" and described as "subsequently abandoned".
  - S2: adds PROJECT_AUDIT_ANNUALSTATEMENT.md, `.claude/settings.local.json` [TOOLING] ("Permits
    `python -m py_compile` via Bash"), README "364 lines", .gitignore "79 lines", requirements
    marked "INCOMPLETE - missing pywin32"; an Entry Points table; and a new "Dead Code Analysis"
    table replacing "None detected": helpers.py `calculate_file_hash()` dead;
    data_consolidator.py `_write_section()` and `_create_details_sheet()` dead; constants
    `TEMPLATE_EXTENSIONS` and `EXCEL_SIGNATURES` dead; cache_manager `get_cached_value()` /
    `set_cached_value()` dead; main.py:72 refers to non-existent `gui/app.py`; four stale
    "# ADD THIS" comments (helpers.py:15, file_parser.py:16, file_organizer.py:15,
    excel_handler.py:39); plus two rows that say "NOT dead" (extract_filename_without_extension,
    parsing_rate).
  - S3: features renumbered F1-F13 and given "Concrete trace" walk-throughs (F1, F2, F3, F4, F7)
    that use a real-company-like example client name (locations only, see "Possible client
    names"). New "bugs found": F1 completeness note (judged "arguably correct"); F6 dry run
    lambda late-binding bug, "confirmed silent bug" (later fixed by 0890022); F10 "Shared state
    verification" finding the cache path mismatch, "confirmed bug" (later fixed by c1583d6);
    F11 an "unreachable return" in report_processor `_refresh_power_query_simple()` lines
    342-345; F13 os.startfile at client_handler.py:251 is Windows-only.
  - S4: a new "Settings/Cache Data Flow" diagram showing writer ~/.gst_organizer_cache.json vs
    reader ./gst_organizer_cache.json ("MISMATCH!").
  - S5: new sub-sections Expected File Types, State Code Mapping [GST RULE] (state-code logic;
    alternative names orissa->OD, pondicherry->PY, "national capital territory of delhi"->DL,
    "dadra and nagar haveli and daman and diu"->DD; fallback initials) TIME-VARYING (state/UT
    names and codes can change), Client Name Length Limit, Folder Structure Template, File Type
    -> Destination Folder, Client Name Validation, Processing Modes, and an "Operational
    Constants" table (MAX_FILE_SIZE_MB 100 never enforced; EXCEL_SAFETY max_retries 3,
    retry_delay 1, temp_prefix "temp_gst_", verification_size_min 10240 "not used"; wait_time
    10; 218-character threshold; unique-path limit 100). The sanitization rule now says
    "Private -> Pvt, Limited -> Ltd (case-sensitive!)" (v1 did not say this; 33d0baf later
    changed exactly this).
  - S6: rewritten as tables with ROBUST/MODERATE/FRAGILE ratings per edge case; a list of bare
    `except:` locations (excel_handler.py:62,124,222-228,263,545).
  - S7: new "Critical Gaps" 1-6 (cache path mismatch; lambda capture; unguarded win32 imports in
    report_processor.py:11-14; os.startfile without platform check; requirements.txt missing
    pywin32; shared mutable state without locks), "Silent Logic Bugs" 7-8 (case-sensitive
    abbreviation; validate_client_name used to validate the state), "Scenario-Based Failure
    Analysis" 9-13 (working directory, second run, two instances, Excel already open via
    GetActiveObject, OneDrive path with spaces), "Missing Validations" 14-16.
  - S8: decisions D1-D9 with confidence, plus a new "Pattern Consistency Check" table (platform
    guard, os.startfile, thread-safe updates, logging, COM cleanup, cache path, case sensitivity).
  - S9: an "OS-Specific Dependencies" table (which files import win32com, pythoncom, win32api,
    win32con, win32gui, os.startfile, and whether there is a fallback).
  - S10: more UI detail (pixel heights, colours, button texts with emoji).
  - S11: "Known Improvements" now explicitly constrained by S8 ("not recommending replacing
    Tkinter (D1) or Win32COM (D2)") and lists ten concrete fixes: 1 central cache path
    `Path.home() / '.gst_organizer_cache.json'`; 2 lambda default-argument binding; 3 guard
    report_processor imports; 4 platform check for client_handler.py:251; 5
    `pywin32>=305; sys_platform == 'win32'` in requirements.txt; 6
    `re.sub(r'\bPrivate\b', 'Pvt', ..., flags=re.IGNORECASE)` in sanitize_filename(); 7 validate
    templates before processing; 8 remove dead code; 9 clean [DEBUG] logging; 10 template
    validation at scan time.
  - S12: a "Branch Map" diagram; the phases re-cut into 6 phases plus "Unmerged Branch"; a
    Phase 6 entry for fc279af; File Evolution table extended (launch_extractor.py,
    requirements.txt, gui/main_window.py, tabs/utils/widgets all "never modified"); Logic and
    Rule Evolution turned into a table; Edge Cases table adds 1c48722 and the remark that
    6237597 introduced the lambda-capture bug; Abandoned Approaches adds the dead
    data_consolidator methods and the mostly-unused EXCEL_SAFETY constants.
  - A new closing "Self-Consistency Review Notes" (8 items; two "[Corrected during self-review]").

  What the revision REMOVED (READ):
  - The v1 S5 table of Power Query Extractor cell mappings ([GST RULE] TIME-VARYING: the
    'T6 of R9' K29 "Table 6J Match", 'T8 of R9' G6 "Table 8D Limits" and the other rows listed
    under fc279af). RAN: the strings "K29", "T6 of R9", "T8 of R9" occur 0 times in the new
    version. The new version therefore no longer documents where any extracted GST figure is
    read from.
  - v1 S5 "File Categorization Rules" and "Long Path Handling" sub-sections (the 218-character
    threshold survives only in the constants table).
  - v1 features F2 (completeness), F6 (organization text report), F14 (state codes) and F15
    (configurable folder names) as separate features; v1 F7/F8 merged into new F11.
  - v1 S7 items: "No GSTIN validation", "Excel process zombie", regex backtracking, subfolders
    not scanned (rglob), lower-case glob on Linux, Unicode normalisation (as a gap), sequential
    processing, no progress persistence, directory traversal, plaintext cache, no undo,
    per-client progress bar.
  - v1 S8 "Why Keyboard Simulation for Power Query Refresh?" (MEDIUM), "Why Copy Files",
    "Why ... Preserve Hyphens", "Evidence of Iteration" (the "6 numbered debug checkpoints").
  - v1 S11 improvements: QThread/PyQt5, RefreshAll() instead of SendKeys, recursive scanning,
    GSTIN validation, progress persistence, "PQ Extractor ... as a fourth tab", structured
    config directory, hyphenated names.
  - v1 S12 statements that c258047 was "Co-authored with Claude Opus 4.6" and its cleanup
    numbers (50 debug lines, 26 bare excepts, 344-line function) - these came from c258047's own
    message (RAN `git show -s c258047`).

  What the revision CHANGED that can be checked against git (RAN `git log`):
  - edac6a3: v1 said 2025-08-31 (correct: 31-08-2025 22:11); v2 says "Phase 1: Initial Release
    (edac6a3 - 2025-10-29)" (wrong date).
  - c258047/a62d3cc/c21f788: v1 said 2026-02-07 (correct: 07-02-2026); v2 says 2026-03-04 in
    S1, S12 heading and branch name context (wrong date; 2026-03-04 is only the date in the
    branch's name).
- Why: a second pass of the same audit method the next morning, apparently to satisfy the
  method's stricter steps (dead code at every level, shared-state write/read tracing, concrete
  traces, pattern-consistency check, self-consistency review). HIGH (the added sections map
  one-to-one onto Project_Audit_FINAL.md process steps 4, 8, 9 and the Self-Consistency Review).
- What it reveals: f7eb8f1's S11 list is the to-do list that the July commits worked through.
  Several July changes reproduce its wording almost exactly: c1583d6 (cache path, item 1),
  0890022 (lambda default arguments, item 2), 064c634 (`pywin32>=305; sys_platform == 'win32'`,
  item 5, character for character; [DEBUG] clean-up, item 9), 33d0baf (`re.sub(r'\bPrivate\b'
  ... flags=re.IGNORECASE)`, item 6, character for character). HIGH. Items still open at A's
  HEAD (READ, grep): item 4 (gui/handlers/client_handler.py:255 still calls os.startfile with no
  platform check); item 8 (utils/helpers.py:529 calculate_file_hash, utils/constants.py:197
  TEMPLATE_EXTENSIONS and :199 EXCEL_SIGNATURES, gui/utils/cache_manager.py:50
  get_cached_value, power_query_extractor/core/data_consolidator.py:250 `_write_section` and
  :354 `_create_details_sheet` are all still present; "# ADD THIS" comments remain at
  core/file_organizer.py:15, core/excel_handler.py:39, utils/helpers.py:17,
  power_query_extractor/core/data_consolidator.py:222/379/402,
  power_query_extractor/core/report_processor.py:132/183).
- Type: docs.
- Rule changes: none in code. Documentation of [GST RULE] cell mappings removed (above).
- Abandoned or reversed: reverses (rewrites) fc279af's text. Never updated after 05-03-2026, so at
  HEAD it is stale (it still calls the cache mismatch, lambda bug and missing pywin32 "confirmed"
  open problems, all fixed on 23-07-2026). HIGH.
- Dependency changes: none.
- Discrepancies between message and diff: none (message is a title). The date errors above are
  inside the document.

#### a603c02 - 23-07-2026 00:22 - Sidhant Goyal - A only (after the split)

- Message: "Speed up network file scanning (behavior-preserving)". Body (quoted):
  "Cut redundant filesystem round-trips when scanning source folders that live on
  network/UNC/OneDrive shares, where every metadata call is a separate slow round-trip. No change
  to which files are found, validated, or reported — verified byte-identical against a
  before/after characterization snapshot covering case-insensitive extensions, invalid files,
  templates, and directories.
  - find_excel_files: enumerate the folder once with os.scandir instead of three separate glob()
    passes (*.xlsx/*.xls/*.xlsm).
  - get_file_info: derive is_file/is_dir from a single stat() instead of
    exists()+stat()+is_file()+is_dir() (4 calls -> 2).
  - file_parser._process_file: remove the duplicate validate_excel_file() call; files reaching it
    were already validated by find_excel_files().
  Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
  (First commit in the history with an AI co-author trailer; all ten July commits in my range
  carry the same trailer.)
- Files touched:
  - core/file_parser.py +6/-11: in `_process_file` (at the commit, lines ~179-215) the block
    "if not validate_excel_file(file_path): append a variation 'Invalid or corrupted Excel file'
    with size; return" is deleted and replaced by a NOTE comment explaining why. READ: the only
    caller of `_process_file` is `scan_folder`, which loops over `find_excel_files(folder)`, and
    `find_excel_files` already returns only files for which `validate_excel_file` is true
    (utils/helpers.py at the parent commit, `valid_files = [f for f in excel_files if
    validate_excel_file(f)]`). So the deleted branch could only ever fire if a file changed
    between the two checks. HIGH.
  - utils/helpers.py +19/-8: new import `from stat import S_ISREG, S_ISDIR`.
    `get_file_info`: `'is_file': path.is_file()` -> `S_ISREG(stat.st_mode)` and
    `'is_dir': path.is_dir()` -> `S_ISDIR(stat.st_mode)` (both still follow symbolic links, like
    before). The function still calls `path.exists()` and `path.stat()` first, so 4 file-system
    queries become 2 (READ, matches the message). `find_excel_files`: the three
    `folder_path.glob('*.xlsx' / '*.xls' / '*.xlsm')` passes are replaced by one
    `os.scandir(folder_path)` loop that keeps entries whose lower-cased name ends with
    ('.xlsx', '.xls', '.xlsm'); the validity filter and `sorted(...)` are unchanged; on any
    exception it still logs and returns [] (READ).
- Why: scanning a source folder on a network share (the author's office share / OneDrive) was
  slow because every file-system call is a round-trip. HIGH (message and comments; later commits
  56aaad9 and 4f84ba9 continue the same work, the latter quoting "11 min -> 41s on the real
  share").
- What it reveals: the app is used against network folders, not just local disk. Files that fail
  validation (too small, wrong first bytes, not a file) are silently left out of the scan: they do
  not appear in "Variations" and nothing is logged (READ: before this commit the variations branch
  was unreachable anyway; at HEAD utils/helpers.py:275-366 `_has_excel_signature` /
  `_passes_metadata_checks` return False without a log line except on exceptions, and grep finds no
  "invalid"/"corrupt" reporting in core/file_parser.py, gui/main_window.py,
  gui/handlers/file_handler.py). Effect: e.g. a 0-byte download, an Excel lock file (~$...), or
  an ".xls" that is really an HTML/XML export would vanish from the scan with no message. MEDIUM
  (the HTML/XML-as-.xls case is an inference about real-world exports; how to check: put such a
  file in a source folder and scan).
- Type: refactor (performance).
- Bug fix? No (performance only). But see Discrepancies: on non-Windows systems it changes which
  files are found.
- Rule changes: file discovery rule unchanged on Windows (extensions .xlsx/.xls/.xlsm discovered;
  .xltx/.xltm accepted by `validate_excel_file` but never discovered; >=1,024 bytes; PK or OLE
  signature). Not a GST rule.
- Abandoned or reversed: the "Invalid or corrupted Excel file" variation entry (present since
  edac6a3) is removed for good. `find_excel_files` itself is later reduced to a thin wrapper over
  the new `scan_excel_files` (56aaad9, 4f84ba9; HEAD utils/helpers.py:309-373), and
  `get_file_info` later takes an optional pre-fetched stat (HEAD utils/helpers.py:129-163).
- Dependency changes: none (stat and os are standard library).
- Discrepancies between message and diff:
  1. "No change to which files are found" is true on Windows only. RAN
     (/root/merge-scratch/run/H-A1/t_a603): on Linux the old code found
     ['.hidden.xlsx', 'Lower.xlsx', 'macro.xlsm', 'old.xls'] and the new code additionally found
     'Mixed.Xlsx' and 'UPPER.XLSX' (pathlib glob is case-sensitive on Linux, the new lower-case
     match is not). On Windows the old glob was already case-insensitive (the code comment says
     so). Harmless for the Windows office use; it matters only for Linux/macOS runs and for tests
     (6c1e1c6's test expecting "UPPER.XLSX" would fail on Linux with the old code).
  2. The "byte-identical characterization snapshot" is not in the repository; unverifiable.
- B status (for the merge): B's HEAD still has the three glob passes (B utils/helpers.py:~255-257),
  the old `path.is_file()`/`path.is_dir()` (B utils/helpers.py:131-132) and the unreachable
  "Invalid or corrupted Excel file" branch (B core/file_parser.py:203). READ.

#### 12d505f - 23-07-2026 00:34 - Sidhant Goyal - A only (after the split)

- Message: "Build Excel reports on local disk, copy to network once". Body (quoted):
  "Report creation previously copied the template to the network target and then
  opened/edited/saved the workbook in place via Excel COM. While a workbook is open Excel
  performs many small random-access reads/writes; over a network share each is a slow
  round-trip.
  Now the report is staged on local disk (self.temp_dir): copy template -> local temp, let Excel
  do all its work locally, then copy the finished file to the network destination in a single
  sequential write. Applied to both the win32com path (production) and the openpyxl fallback.
  Behavior preserved:
  - Same output report - mapped Links-sheet cells written identically, all other cells/sheets
    untouched (verified byte-identical on the openpyxl path via a before/after characterization
    snapshot).
  - Power Query is preserved: report data sources are ABSOLUTE paths stored in the Links sheet,
    so they resolve the same regardless of where the report file is stored; saving locally then
    copying yields the same file.
  - On failure the destination is no longer left with a half-written template copy (minor
    improvement).
  Note: the win32com path requires real Excel and could not be executed in CI; validate one
  report on a Windows machine before bulk processing.
  Co-Authored-By: Claude Opus 4.8 (1M context)"
- Files touched:
  - core/excel_handler.py +72/-24 (several of the +/- pairs are only trailing-whitespace
    removal). Changes (READ, line numbers at the commit):
    - `_create_report_win32com` (103-259): new local `local_temp = None` (115). Instead of
      `shutil.copy2(template, output_path)` it now sets `local_temp = self.temp_dir /
      output_path.name` (146), deletes a leftover file of that name if present (147-151, errors
      ignored), copies the template there (153), keeps `time.sleep(0.5)` (154), opens the local
      copy with `_safe_open_workbook(excel, local_temp)` (158). After the unchanged
      write-cells / `wb.Save()` / `wb.Close(SaveChanges=True)` / `time.sleep(0.5)` it copies the
      finished file to the destination with `shutil.copy2(local_temp, output_path)` (214), then
      the unchanged existence/size check. The `finally` block gains a best-effort delete of the
      local file after COM clean-up (253-259).
    - `_create_report_openpyxl` (261-321): same staging: copy template to `local_temp`,
      `load_workbook(local_temp, data_only=False, keep_links=True)`, write mapped cells, save to
      `local_temp`, close, then copy to `output_path`; new `finally` deletes `local_temp`.
    - `self.temp_dir` is unchanged from before: `Path(tempfile.gettempdir()) / "gst_excel_temp"`,
      created in `__init__` (line 52-53), i.e. %TEMP%\gst_excel_temp on Windows.
- Why: report creation was slow because Excel was editing the file directly on the network
  share. HIGH (message, comments).
- What it reveals:
  - The ITC and Sales templates' Power Query reads its source files through paths written into
    the "Links" sheet (the message's stated assumption). Not verifiable here (no template in the
    repository). MEDIUM. How to check: open an ITC template's Power Query editor and confirm every
    source path comes from the Links cells and none from the workbook's own location.
  - Excel instance handling unchanged and worth recording for the audits (READ, still at HEAD):
    the code first tries `win32com.client.GetActiveObject("Excel.Application")` (HEAD
    core/excel_handler.py:124), i.e. it attaches to the user's own open Excel if there is one;
    it sets `excel.Visible = False` on that instance; in `finally` it sets
    `excel.DisplayAlerts = True` (HEAD :240) and then calls `excel.Quit()` whenever
    `self.excel_app is None` (HEAD :242-244). `self.excel_app` is never assigned anything but
    None (only at HEAD :56; RAN grep), so the comment "Don't quit if Excel was already open"
    (HEAD :241) is not what happens: the app always quits the Excel it used, including a user's
    already-open Excel, and does so with alerts switched back on. HIGH (reading); the effect on a
    real desktop (whether a "save changes?" prompt appears for the user's other workbooks) is
    MEDIUM, needs a Windows test. Compare the author's Standard 12 "Excel through COM" rules
    (rebuild_pipeline_FINAL.md lines 851-856: DispatchEx, never take over the user's Excel,
    DisplayAlerts = False so Quit cannot hang, Quit in finally) and lines 872-875
    (CoUninitialize only after the function holding Excel objects has returned; here
    `pythoncom.CoUninitialize()` runs inside the same function's `finally`, HEAD :248-251, while
    the local `excel` variable still refers to Excel). Technical observations only.
  - The openpyxl fallback still saves a copy of a template that normally contains Power Query;
    openpyxl drops the queries on save (the code's own warning, line 266). That path runs when
    `USE_WIN32COM` is False (non-Windows, or pywin32 failed to import). Standard 12 lines 846-850
    say never let openpyxl save over such a workbook. A later A commit (be6418c "Require pywin32
    on Windows so reports cannot lose Power Query", outside my range) addresses the Windows case.
- Type: refactor (performance) with a small robustness gain.
- Bug fix? Partly: "on failure the destination is no longer left with a half-written template
  copy". RAN (/root/merge-scratch/run/H-A1/t_12d5): with a corrupt template on the openpyxl path,
  the old code returned False and left `R.xlsx` (a raw copy of the template) in the destination;
  the new code returned False and left nothing; the temp folder was empty afterwards in both.
  Remaining similar cases (READ): if the final `shutil.copy2(local_temp, output_path)` itself
  fails half-way (network drop), a partial destination file can remain; if Excel still holds the
  local file when `finally` runs, the delete fails silently and leaves a file in
  %TEMP%\gst_excel_temp (the next run deletes a same-named leftover first, lines 147-151).
- Rule changes: none (the Links cell map EXCEL_TEMPLATE_MAPPING in utils/constants.py is
  untouched).
- Abandoned or reversed: the old "copy template straight to output_path, open there" approach
  (since edac6a3) is replaced. Still in place at HEAD (core/excel_handler.py:147-159, :215,
  :255-258, :273-320). A side effect (inference, MEDIUM): because Excel now opens a short local
  temp path instead of the long network path, the long-path fallback in `_safe_open_workbook`
  (8.3 short names above 218 characters) is rarely needed for report creation any more.
- Dependency changes: none.
- Discrepancies between message and diff:
  1. RAN: on the openpyxl path, old and new outputs had the same sheets and the same cell values
     (Links!A1 kept, B2 and B5 written, Other!C3 kept). I did not compare bytes (openpyxl writes
     timestamps); the "byte-identical" snapshot is not in the repository.
  2. The code comment "Saving locally then copying yields a byte-identical file at the
     destination" (line 145) is only trivially true (the copy equals the local file). Compared
     with the old behaviour, an Excel-saved .xlsx normally records the folder it was last saved
     in inside the workbook (the `x15ac:absPath` element of workbook.xml), so reports created by
     the COM path would now record the %TEMP% staging folder instead of the client's folder.
     MEDIUM (general Excel behaviour, not testable here). How to check: unzip a report made on
     Windows after this commit and look at xl/workbook.xml. Practical effect believed nil unless
     something in the template uses the file's own saved location.
  3. The win32com path was never executed (the message says so; nor could I on Linux).
- B status: B's HEAD still copies the template straight to the output path and edits it there
  (B core/excel_handler.py:124 and :228). READ.

#### 275d9a4 - 23-07-2026 00:46 - Sidhant Goyal - A only (after the split)

- Message: "Cache level-1 folder; single-listing scans in PQ extractor". Body (quoted):
  "More network round-trip reduction, behavior-preserving (verified via characterization +
  equivalence tests).
  FileOrganizer._get_level1_folder: resolve the Level 1 folder once per run and reuse it for
  every client. Previously create_client_structure re-resolved it per client - in rerun/resume
  mode that re-scanned the target (iterdir + a stat per entry) for every client; in fresh mode it
  re-issued a mkdir per client. Result is identical within a run. Verified: folder tree, returned
  paths, and summary byte-identical across fresh and rerun modes with a repeated client.
  PQ extractor (extractor_window): replace per-client glob passes with two pure helpers that each
  do a single os.scandir:
  - _classify_version_reports -> (has_itc, has_sales)
  - _latest_refresh_times -> (itc_status, sales_status), also removing the redundant double
    stat() per winner.
  scan_folder + get_refresh_status now do fewer directory listings per client. Helpers proven
  equivalent to the original glob logic across 9 folder layouts
  (empty/itc/sales/both/refreshed-newest/mtime-tie/mixed-case/near-misses/refreshed-only), with
  coverage asserted.
  Per-file copy metadata trims were evaluated and deliberately NOT changed: the file transfer
  itself dominates that cost and the checks guard copy verification.
  Co-Authored-By: Claude Opus 4.8 (1M context)"
- Files touched:
  - core/file_organizer.py +18/-1: `__init__` gains `self._level1_folder = None` (line 47).
    `_get_level1_folder()` (150-161) becomes a caching wrapper: compute once via the renamed
    old body `_compute_level1_folder()` (163-196), then return the stored value. The old body is
    unchanged: 'fresh' -> create "Annual Statement-{timestamp}" under the target and append it
    to `created_folders`; 'rerun'/'resume' -> pick the existing "Annual Statement-*" folder with
    the newest mtime, or create a new one if none exists. READ.
  - power_query_extractor/gui/extractor_window.py +72/-27: new `import os`; two new module-level
    functions:
    - `_classify_version_reports(version_folder)` (lines 19-41): one `os.scandir`; has_itc if
      any entry name lower-cased starts with 'itc_report_' and ends with '.xlsx'; has_sales
      likewise with 'sales_report_' (elif); OSError -> (False, False) so far.
    - `_latest_refresh_times(version_folder)` (44-78): one `os.scandir`; skips names without
      '_refreshed_' or not ending '.xlsx'; keeps the largest `entry.stat().st_mtime` for
      'itc_report_' and for 'sales_report_' names; formats each as "%Y-%m-%d %H:%M" or None.
    - `scan_folder` (560-612): `list(latest_version.glob("ITC_Report_*.xlsx"))` /
      `glob("Sales_Report_*.xlsx")` replaced by `_classify_version_reports(latest_version)`.
    - `get_refresh_status` (700-705): body replaced by `return
      _latest_refresh_times(client['latest_version'])`, with `except Exception: return None,
      None` (old: two globs, `max(..., key=stat().st_mtime)`, a second `stat()`, bare `except:
      pass`, then return whatever had been found).
- Why: same network-speed drive as a603c02. HIGH.
- What it reveals: in rerun/resume mode the "current" Annual Statement folder is the newest by
  modified time (core/file_organizer.py:177-186, unchanged); the extractor likewise. A folder's
  modified time changes whenever something is created inside it directly (on Windows NTFS this
  includes Excel's "~$" lock file when someone opens a workbook stored directly in that folder),
  so opening, e.g., an older run's summary workbook could make an older Annual Statement folder
  "newest". MEDIUM (file-system behaviour, not testable here; how to check: on the share, open an
  old GST_Processing_Summary workbook, then look at the folder's Date modified).
- Type: refactor (performance).
- Bug fix? No.
- Rule changes: none. Refreshed-file recognition still hard-codes '_refreshed_' (case-insensitive)
  regardless of the user's configurable suffix (later fixed in 2aaa2b1, outside my range; HEAD
  power_query_extractor/gui/extractor_window.py:45-82 now calls `is_refreshed_name(entry.name,
  suffix_pattern)`).
- Abandoned or reversed: the per-client glob approach. The helper docstring "Reproduces the old
  checks exactly" survives at HEAD (extractor_window.py:21-29).
- Dependency changes: none.
- Discrepancies between message and diff:
  1. "Result is identical within a run ... summary byte-identical": the Level 1 path is
     identical, but the per-client text report `_Organization_Report_<client>_<state>.txt`
     changes in fresh mode. Its "Folders Created" line prints `len(self.created_folders)`
     (core/file_organizer.py:368, no de-duplication), and before this commit fresh mode appended
     the Level 1 folder once per client. RAN (/root/merge-scratch/run/H-A1/t_275d, three
     placeholder clients, fresh mode): old reports said 6 / 12 / 18, new say 6 / 11 / 16; rerun
     mode 6 / 11 / 16 in both. `get_summary()` (which uses `len(set(...))`) gives 16 in all
     cases, but `get_summary` is never called (RAN grep: only its definition at :409). Note the
     count is run-wide and cumulative in both versions (client 3's report counts clients 1-2's
     folders too); same for "Successfully Copied" (:363-367). Minor, display only.
  2. "Reproduces the old checks exactly": RAN (/root/merge-scratch/run/H-A1/t_275d/pq) - a name
     like "ITC_Report_Refreshed_.xlsx" is treated as refreshed by the new helper (the "_" is
     shared between the prefix and "_refreshed_") but not by the old glob
     "ITC_Report_*_Refreshed_*.xlsx". Also on Linux the old globs were case-sensitive. Neither
     matters for real file names (which always have client/state/timestamp parts) on Windows.
  3. The "9 folder layouts" equivalence tests are not in the repository (RAN: no test at HEAD
     references `_classify_version_reports`; only tests/test_refresh_naming.py, added later,
     references `_latest_refresh_times`).
- B status: B's core/file_organizer.py:146 `_get_level1_folder` exists without caching (no
  `_level1_folder` attribute; READ). B has no extractor_window.py (its extraction tab is PyQt;
  B gui/pyqt_main_window.py:461-462 uses `any(version_dir.glob("ITC_Report_*.xlsx"))`).

#### 5c8ce95 - 23-07-2026 01:02 - Sidhant Goyal - A only (after the split)

- Message: "Detect IMS Reco files with a trailing (N) duplicate suffix". Body (quoted; the
  example name is generic):
  "IMS Reco filenames ending in a duplicate marker like " (1)" / "(2)" (e.g.
  ImsReco-ABC-Delhi-01012024 (1).xlsx) were not detected: the pattern required exactly 8 digits
  immediately before the extension. Allow an optional "(N)" (with or without a leading space)
  after the 8-digit date.
  The date group still captures exactly the 8 digits, so metadata is unchanged for existing
  files. All other types already tolerated the suffix via their flexible trailing group; only
  IMS Reco needed this.
  Verified: plain IMS Reco parses identically; suffixed IMS Reco now parses with the same
  client/state and date '01012024'; other types and the full scan flow unchanged.
  Co-Authored-By: Claude Opus 4.8 (1M context)"
- Files touched:
  - utils/constants.py +1/-1, FILE_PATTERNS['ImsReco']['pattern'] (line 23):
    old `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$`
    new `^ImsReco-([^-]+)-([^-]+)-(\d{8})(?:\s*\(\d+\))?\.xlsx?$` (both re.IGNORECASE).
    The description 'ImsReco-ClientName-State-DDMMYYYY.xlsx' and groups ['client','state','date']
    are unchanged.
- Why: real downloads of the IMS reconciliation file arrived with a browser duplicate marker
  " (1)" and were being shown as unrecognised ("variations") instead of IMS Reco. HIGH (the
  example and the wording "were not detected").
- What it reveals: [GST RULE] (file-type recognition that routes the file into the ITC report's
  Links cells B10/B11 and captures an 8-digit date, described by the code as DDMMYYYY). RAN
  (regex test): plain name -> same groups; "... (1).xlsx", "...(2).xlsx", "...  (12).XLS" -> now
  IMS Reco with the same 8-digit date; still NOT recognised: a leading "(1) ImsReco-..." (only
  the GSTR-3B pattern accepts a leading "(N) "), "... - Copy.xlsx", "... (1) (1).xlsx",
  "... ().xlsx", any ".xlsm" IMS file, and 7-digit dates. For the other five types the "(N)" is
  accepted but becomes part of the captured period/month/year text (RAN: GSTR3B month
  "Jan (1)", Sales end month "Jun (2)", AnnualReport year "2024 (1)", GSTR-2B period
  "Apr24 (1)", SalesReco period "Q1 (1)").
- Type: bug fix.
- If a bug fix: broke for IMS file names with a trailing "(N)"; fixed properly at the pattern
  (not a patch elsewhere). Consequences now (READ): if a folder holds both "X.xlsx" and
  "X (1).xlsx" for the same client, the client gets two IMS Reco files, the completeness check
  reports "Multiple IMS Reco (2 files)" and status "Has duplicates" (or "Missing N files" if
  something else is missing) (core/file_parser.py:252-286 at HEAD), and the ITC report links
  `ims_files[0]` (core/excel_handler.py:647-651 at HEAD). Files are processed in sorted path
  order (utils/helpers.py scan, `sorted(...)`), and RAN `sorted()` puts "... (1).xlsx" and
  "...(1).xlsx" BEFORE the plain "....xlsx" (space and "(" sort before "."), so the duplicate
  copy, not the original, is the one written into the ITC template. HIGH (reading + sort test);
  whether the "(1)" file is the newer or older download is unknown (question for the author).
  The same ordering applies to every type that accepts "(N)". Still the same at HEAD
  (utils/constants.py:23).
- Rule changes: [GST RULE] IMS Reco recognition widened as above (utils/constants.py:23,
  old -> new quoted). Not TIME-VARYING in itself.
- Abandoned or reversed: none.
- Dependency changes: none.
- Discrepancies: none found; "All other types already tolerated the suffix" is true (RAN), with
  the side effect on captured metadata noted above (not mentioned in the message).
- B status: B's HEAD still has the old pattern (B utils/constants.py:23). READ.

#### c1583d6 - 23-07-2026 01:49 - Sidhant Goyal - A only (after the split)

- Message: "Fix extractor cache path and version-folder selection". Body (quoted):
  "Two correctness bugs in the Power Query Extractor:
  1) Cache handoff was broken. The main organizer saves its cache to ~/.gst_organizer_cache.json
     (gui/utils/cache_manager.py), but the extractor read a differently-named file in the current
     working directory ("gst_organizer_cache.json"), so auto-loading the last-used folder never
     worked. Point both extractor readers (extractor_main load_cached_folder + extractor_window)
     at the same home-dir dotfile. Additive for extractor_main (home path checked first, old
     paths kept).
  2) Wrong "latest version" folder. scan_folder chose the newest Version-* folder with
     max(key=name), but names are "Version-DDMMYY HHMM" and day-first text does not sort
     chronologically across month boundaries (e.g. "300626" sorts after "020726"), so re-runs
     could pick an older version. Now sorts by st_mtime, matching the level-1 folder selection
     right above it.
  Verified: mtime-sort picks the chronologically newest folder where name-sort picked the wrong
  one; extractor now loads the folder from the same cache file the main app writes.
  Co-Authored-By: Claude Opus 4.8 (1M context)"
- Files touched:
  - power_query_extractor/extractor_main.py +3/-0: `load_cached_folder` list of candidate files
    gets `Path.home() / '.gst_organizer_cache.json'` as the FIRST entry (line 41), before the old
    CWD-relative "gst_organizer_cache.json", ".cache/app_state.json", "config/last_session.json".
    The rest (try keys 'target_folder', 'target_path', 'output_folder', 'last_target'; return ''
    if none) is unchanged (READ, lines 36-63).
  - power_query_extractor/gui/extractor_window.py +10/-2: `self.cache_file` changed from
    `Path("gst_organizer_cache.json")` to `Path.home() / '.gst_organizer_cache.json'` (line 94);
    in `scan_folder`, `latest_version = max(version_folders, key=lambda x: x.name)` ->
    `key=lambda x: x.stat().st_mtime` (line ~600), with a comment.
- Write/read path check (READ): writer gui/utils/cache_manager.py:15 `Path.home() /
  '.gst_organizer_cache.json'`, `save_cache` writes the dict from gui/main_window.py (key
  'target_folder' at main_window.py:334 at the commit). Readers: extractor_main.py:41 (first
  candidate, key 'target_folder' tried first) and extractor_window.py `load_from_cache` (reads
  'target_folder', lines 539-551). Neither extractor file writes the cache. The value is the base
  target folder; the extractor then looks for "Annual Statement-*" inside it - consistent. HIGH.
- Why: the extractor never auto-loaded the last target folder, and could pick an older Version
  folder after a re-run that crossed a month boundary. HIGH. Both bugs had been listed in the
  f7eb8f1 audit (cache mismatch as "Critical Gap 1"); the version-name sort bug was not in it.
- What it reveals: run folders are named day-first ("DDMMYY HHMM") so they cannot be compared as
  text; the code relies on folder modified times instead. RAN: `max(['Version-300626 1000',
  'Version-020726 1000'])` returns the 30-06 folder although 02-07 is later.
- Type: bug fix (two).
- If a bug fix: (1) proper fix (both readers now use the writer's path); the stale CWD
  candidates in extractor_main are kept as fallbacks (could pick up an unrelated old file in
  whatever folder the app is started from, only if the home file is missing; LOW).
  (2) Fixed by switching to mtime. Remaining similar cases (READ, HEAD):
  - mtime is "last changed", not "created": anything written directly inside an older Version
    folder (an extractor "_Refreshed_" copy, a "~$" Excel lock file from opening a report
    there, a manual copy) makes that older folder "newest". MEDIUM (see 275d9a4 note).
  - The Level 1 choice in the extractor uses `base_path.glob("Annual Statement-*")` and the
    Version choice `client_folder.glob("Version-*")` without checking they are folders
    (extractor_window.py:623 and :641 at HEAD); a FILE whose name starts "Annual Statement-"
    with the newest mtime would be chosen and `iterdir()` would fail into "Error scanning
    folder". LOW.
  - B (for the merge): B picks the latest Level 1 folder by NAME, descending
    (B gui/pyqt_main_window.py:432-436 `sorted(..., key=lambda p: p.name, reverse=True)`) and
    lists Version folders by name descending (B gui/pyqt_main_window.py:457), i.e. the same
    day-first ordering problem class this commit fixed in A. READ, HIGH. B's
    power_query_extractor/extractor_main.py:39 still reads only the CWD
    "gst_organizer_cache.json" (B's own cache lives elsewhere; B models/cache_manager.py:95-98
    mentions migrating from ~/.gst_organizer_cache.json).
- Rule changes: business rule "which Version folder is current" changed from name order to
  newest modified time (extractor_window.py:648 at HEAD). Not a GST rule.
- Abandoned or reversed: name-based version choice (since edac6a3) abandoned. The extractor
  window no longer reads the CWD cache file at all (a replacement, not an addition, for that
  file - the message only claims "additive" for extractor_main, which is accurate).
- Dependency changes: none.
- Discrepancies: none found.

#### 0890022 - 23-07-2026 01:54 - Sidhant Goyal - A only (after the split)

- Message: "Fix late-binding closures in processing log/progress messages". Body (quoted):
  "The per-client processing loop scheduled UI updates via root.after(0, lambda: ...
  client_info/folders/result ...). Those lambdas captured the loop variables by reference, so if
  the Tk event loop drained them a beat late they could render the NEXT client's data - wrong
  client names in the log/progress line. (Display only: the actual folder/copy/report work reads
  the values directly in the worker thread and was always correct.)
  Route every root.after log/progress call through two helpers - _log_async and _progress_async
  - that capture the already-formatted message/values via default arguments at schedule time.
  Same messages, same order, same levels; values are now bound when scheduled instead of when
  run.
  Verified with a test that reproduces the bug (deferred callbacks all showed the last client)
  and confirms both helpers preserve the correct per-iteration values and ordering.
  Co-Authored-By: Claude Opus 4.8 (1M context)"
- Files touched:
  - gui/handlers/processing_handler.py +40/-25: two new methods `_log_async(message,
    level='normal')` (lines 23-31: `root.after(0, lambda m=message, lv=level:
    self.app.log_message(m, lv))`) and `_progress_async(progress, message)` (33-36). Every
    `self.app.root.after(0, lambda: ...)` call is replaced: in `dry_run` (9 calls), in
    `start_processing` (1, "PROCESSING STARTED") and in `process_files_thread` (per-client
    progress line, "Processing <client>", "Creating folders", "Created folder structure", the
    three "Level 1/2/3" lines, the nested `file_progress` callback, "Organized N files", one
    line per copied file, "Creating Excel reports", "Creating ITC report...", template name,
    output name). RAN grep: after the commit the only `root.after` calls left in the file are
    inside the two helpers.
- Why: the processing log and the dry-run preview showed wrong client/file names. HIGH. Listed
  in the f7eb8f1 audit as a "confirmed silent bug" (F6, Critical Gap 2) with this exact remedy.
- What it reveals: `self.app.log_message` (gui/main_window.py:256-258 at the commit) already
  schedules its own screen update through `UIHelpers.log_message`, which formats the timestamped
  text immediately and calls `log_widget.after(0, ...)` with that text bound
  (gui/utils/ui_helpers.py:17-25); `update_progress` (main_window.py:260-265) also captures its
  arguments. So the many direct `self.app.log_message(...)` calls later in the worker (ITC result
  onwards, lines ~271-471 at the commit, and `messagebox.showinfo` at ~450) never had the
  late-binding bug. The extra `root.after` hop only existed in the replaced calls; it is kept
  (so the old ordering is kept). HIGH (READ).
- Type: bug fix.
- If a bug fix: RAN (/root/merge-scratch/run/H-A1/t_0890, Tk under xvfb, two placeholder
  clients, fake app object): with the old code the dry run printed "<Client B>" for BOTH client
  headings and "file-b1" for all three file lines; with the new code each client and file is
  correct. The dry run runs on the main (GUI) thread, so none of its scheduled callbacks can run
  until the loop has finished: the old dry run was ALWAYS wrong for every client except the last,
  not only "a beat late". In the worker-thread processing loop the wrong value appeared only when
  the GUI fell behind (timing-dependent). Proper fix (values bound at scheduling time). Similar
  cases still present (READ, HEAD): worker-thread code still touches Tk directly in places that
  this commit did not target - `messagebox.showinfo(...)` (gui/handlers/processing_handler.py:498 at HEAD) and
  `messagebox.showerror(...)` (:525) near the end of `process_files_thread`, and
  `self.app.start_btn.config(...)` / `stop_btn.config(...)` in its `finally` (:527-530); calling Tkinter from a
  non-GUI thread is not officially supported (works in practice on Windows' threaded Tcl). MEDIUM.
- Rule changes: none.
- Abandoned or reversed: the inline `root.after(0, lambda: ...)` pattern introduced by 6237597
  (shared history, 29-10-2025) is retired in this file.
- Dependency changes: none.
- Discrepancies between message and diff:
  1. Under-describes the dry-run case (deterministically wrong, see RAN above).
  2. "Verified with a test that reproduces the bug": no such test is committed (tests/ does not
     exist until 6c1e1c6; at HEAD no test references `_log_async`; RAN grep).
  3. "Same messages, same order": true relative to the old code (same double hop) - READ.
- B status: not applicable as code (B is PyQt5 with signals); B's own logging path is another
  reader's subject.

#### 6c1e1c6 - 23-07-2026 07:58 - Sidhant Goyal - A only (after the split)

- Message: "Add pytest suite for data-integrity logic (81 tests)". Body (quoted):
  "First automated tests for the repo. Covers the logic whose silent failure would misfile a
  client's data, with zero GUI/COM/network coupling so the suite runs in <1s anywhere:
  - test_file_patterns.py: parse_filename / FILE_PATTERNS - every report type,
    case-insensitivity, .xls extension, and the (N) duplicate markers (trailing suffix incl. the
    IMS fix, GSTR-3B leading prefix).
  - test_helpers.py: pure helpers - sanitize_filename, get_state_code, create_client_state_key,
    validate_client_name, format_size/duration, path helpers. Known quirks (substring
    'Private'->'Pvt' mangling, 'UP' fallback collision) are pinned with NOTE comments as current
    behaviour.
  - test_file_ops.py: validate_excel_file / find_excel_files / get_file_info against real temp
    files (valid/invalid signatures, sizes, extensions, case-insensitive discovery, directories).
  - test_scan.py: end-to-end FileParser.scan_folder - client grouping, completeness/missing
    reporting, variations, invalid-file skipping.
  Tooling: pytest.ini (testpaths=tests), tests/conftest.py (path setup + fake-excel fixture),
  tests/README.md, pytest uncommented in requirements, .pytest_cache/ gitignored.
  Characterization tests (pin current behaviour); mutation-checked to confirm they fail when the
  logic regresses.
  Co-Authored-By: Claude Opus 4.8 (1M context)"
- Files touched (all READ line by line):
  - .gitignore +1/-0: adds `.pytest_cache/` after `__pycache__/`.
  - pytest.ini +6 (new): `[pytest]`, `testpaths = tests`, `python_files = test_*.py`,
    `python_classes = Test*`, `python_functions = test_*`, `addopts = -q`. No `pythonpath`.
  - requirements.txt +3/-1: `pytest>=7.0.0` moved from the commented "Development dependencies"
    block into an active "# Testing (run `pytest` from the repo root)" entry.
  - tests/README.md +34 (new): how to run (`pytest` / `python -m pytest`), a table of the four
    test files, and "Nature of these tests" (characterization; NOTE-marked quirks are "not
    endorsements"; no GUI/COM/network; "<1s").
  - tests/conftest.py +29 (new): inserts the project root into `sys.path`; constants
    PK_SIG = b"PK\x03\x04", OLE_SIG = b"\xd0\xcf\x11\xe0"; `_write_fake_excel(path, header,
    size=2048)` writes the signature plus zero padding (not a real workbook); fixture
    `make_excel`.
  - tests/test_file_ops.py +89 (new, 14 tests): validate_excel_file accepts .xlsx (PK), .xls
    (OLE), .xltx; rejects 500-byte file, wrong signature, .txt, missing file, a directory named
    "adir.xlsx". find_excel_files: returns exactly the four valid .xlsx/.xls/.xlsm files
    including "UPPER.XLSX", excludes tiny, wrong-signature, .xltx, .txt and a directory named
    "looks.xlsx"; returns sorted; empty folder -> []. get_file_info: metadata for a 2,048-byte
    file (is_file True, is_dir False, extension ".xlsx"); a directory -> is_dir True; missing
    file -> {"error": "File not found"}.
  - tests/test_file_patterns.py +94 (new, 21 tests): eight "happy" names (one per type, an .xls
    Sales file, an all-lower-case GSTR-3B name) checked for type/client/state; six trailing
    "(N)" names (GSTR-2B, three IMS variants, Sales, AnnualReport) must parse with client "ABC";
    the IMS date stays "01012024"; "(1) GSTR3B-..." leading prefix parses; five non-matching
    names (too few / too many dash groups, wrong extension, random names) give parsed False,
    type "Unknown", client "".
  - tests/test_helpers.py +154 (new, 41 tests): sanitize_filename ("ABC Private Limited" ->
    "ABC_Pvt_Ltd"; "<>:*" removed; whitespace collapsed; "" -> "unnamed"; "(1)" kept;
    NOTE quirk "Privateer Unlimited" -> "Pvter_Unlimited"; 250 x "x" + ".xlsx" -> 185 x "x" +
    ".xlsx"); get_state_code [GST RULE] (state-code logic) TIME-VARYING (state/UT list):
    "Delhi"/"delhi" -> "DL", "  Maharashtra " -> "MH", "Tamil Nadu" -> "TN", "West Bengal" ->
    "WB", unknown two words "Unknown Place" -> "UP" (NOTE: same as the code the mapping uses for
    Uttar Pradesh), unknown one word "Xyz" -> "XYZ", "" -> ""; create_client_state_key ("ABC
    Private Limited","Delhi") -> "ABC Pvt Ltd-DL", ("XYZ","Tamil Nadu") -> "XYZ-TN", a long name
    + "Maharashtra" -> exactly 35 characters ending "-MH"; validate_client_name messages ("" ->
    "Client name cannot be empty", "A" -> "too short", 101 chars -> "too long", "AB<C" ->
    "contains invalid characters"); format_size (0 -> "0.0 B", 1023 -> "1023.0 B", 1024 ->
    "1.0 KB", 1536 -> "1.5 KB", 1 MiB -> "1.0 MB", 1 GiB -> "1.0 GB", -5 -> "Unknown");
    format_duration (0 "0s", 30 "30s", 59 "59s", 90 "1m 30s", 3599 "59m 59s", 3661 "1h 1m" with
    NOTE "hours format drops the seconds"); clean_windows_path("a/b/c") -> r"a\b\c";
    extract_filename_without_extension ("report.xlsx"->"report", "X.Y.xlsx"->"X.Y",
    "noext"->"noext").
  - tests/test_scan.py +80 (new, 5 tests): two files of one client -> key "ABC-DL", client "ABC",
    state "Delhi", state_code "DL", two types, four missing, status "Missing 4 files"; an
    unparseable valid file -> variations; a bad-signature file -> absent from `scanned` (only
    `scanned` is asserted); empty folder -> nothing; one file of each of the six types ->
    missing [] and status "Complete".
- RAN (/root/merge-scratch/run/H-A1/t_6c1e, `git archive 6c1e1c6`): 81 passed in 0.23 s on Linux,
  Python 3.14; per file 14 + 21 + 41 + 5 = 81 (matches the title).
- Why: to protect the scan/naming logic before more refactoring (the July work was being done
  with an AI assistant; tests pin behaviour). HIGH.
- What it reveals: the author's/assistant's view of which logic is "data-integrity" critical
  (file classification, naming, state codes, validation). Two quirks were consciously pinned
  rather than fixed (the 'Private' substring and the 'UP' collision); the first is fixed one
  commit later (33d0baf), the second is still present (READ: tests/test_helpers.py:96-97 at HEAD still pins
  "Unknown Place" -> "UP"). The invalid-file test pins that
  corrupted/tiny files simply disappear from the scan (see a603c02).
- Type: tests + tooling.
- Rule changes: none in code. The tests freeze these [GST RULE]-adjacent behaviours: state-code
  mapping and fallback (above), IMS date capture.
- Abandoned or reversed: test `test_substring_abbreviation_quirk` is replaced in 33d0baf;
  tests/conftest.py is changed again in a677670, tests/test_file_ops.py in 56aaad9 and 4f84ba9,
  tests/test_helpers.py in 33d0baf and 4219ffb (RAN `git log`).
- Dependency changes: pytest>=7.0.0 becomes a listed requirement (requirements.txt). Compared
  with the author's Standard 13 (rebuild_pipeline_FINAL.md lines 937-938 "exact versions of the
  packages it imports (not pytest)" and 944-947 `pytest.ini` with `pythonpath = .` and
  `testpaths = tests`): this commit adds pytest to the app's requirements with a ">=" range and
  uses a conftest.py sys.path insert instead of `pythonpath = .`. Factual comparison only.
- Discrepancies between message and diff:
  1. "mutation-checked" - no trace in the repository; unverifiable.
  2. Otherwise matches (81 tests, <1 s, tooling list).

#### 33d0baf - 23-07-2026 08:48 - Sidhant Goyal - A only (after the split)

- Message: "Fix sanitize abbreviation, processing re-entrancy, folder-setting key". Body (quoted):
  "Three correctness fixes (tests added where the logic is unit-testable):
  1) sanitize_filename business abbreviation now uses word-boundary, case-insensitive matching
     (re.sub \bPrivate\b / \bLimited\b), matching create_client_state_key. Previously a
     case-sensitive substring replace mangled words that merely CONTAIN them
     ('Privateer'->'Pvter') and missed ALL-CAPS names. Common case
     ('ABC Private Limited'->'ABC_Pvt_Ltd') unchanged; verified by the existing + new tests.
  2) ProcessingHandler.start_processing now guards on is_processing. The Start button is bound
     via .bind('<Button-1>'), which fires even when the button is 'disabled', so a click mid-run
     could launch a second worker thread. New tests assert the guard blocks re-entry and is
     transparent when idle.
  3) client_folder_settings key mismatch: toggle_folder_name_setting and
     update_global_folder_setting keyed by state NAME, but FileOrganizer and get_selected_clients
     key by state CODE - so per-client folder-name toggles were silently ignored. Now all sites
     use client-STATECODE.
  Co-Authored-By: Claude Opus 4.8 (1M context)"
- Files touched:
  - utils/helpers.py +6/-3, `sanitize_filename` (lines ~75-83): `base_name.replace('Private',
    'Pvt')` / `.replace('Limited', 'Ltd')` -> `re.sub(r'\bPrivate\b', 'Pvt', base_name,
    flags=re.IGNORECASE)` / same for Limited -> 'Ltd'. (`import re` already at line 7.)
  - gui/handlers/processing_handler.py +7/-0: at the top of `start_processing` (lines 71-76):
    `if getattr(self.app, 'is_processing', False): return`.
  - gui/handlers/client_handler.py +6/-2: in `toggle_folder_name_setting` and
    `update_global_folder_setting`, `client_key = f"{values[0]}-{values[1]}"` ->
    `f"{values[0]}-{get_state_code(values[1])}"` (values[0] = client name, values[1] = state
    NAME as shown in the tree).
  - tests/test_helpers.py +6/-5: quirk test replaced by `test_abbreviation_uses_word_boundaries`
    ("Privateer Unlimited" -> "Privateer_Unlimited") and
    `test_abbreviation_is_case_insensitive` ("ABC PRIVATE LIMITED" -> "ABC_Pvt_Ltd").
  - tests/test_processing_handler.py +36 (new, 2 tests): fake app with `is_processing` True ->
    `validate_processing_inputs` not called; False -> called.
- Verification of fix 3, write/read of `client_folder_settings` (READ at the commit):
  writers - client_handler.py:49 (`update_client_tree` initialises every client to False using
  the client_data key), :185/:198 (toggle), :219 (global ON); reader - core/file_organizer.py:122
  builds `f"{client_info['client']}-{state_code}"` and reads :130. The client_data key itself is
  `f"{file_data['client']}-{state_code}"` (core/file_parser.py:225), and `get_selected_clients`
  builds `f"{values[0]}-{get_state_code(values[1])}"` (client_handler.py:~165-170). After the
  fix all sites agree (raw client name + 2-letter state code). Before it, toggles were stored
  under "name-StateName", which FileOrganizer never looked up, so with the global checkbox OFF a
  per-client "Yes" had no effect (the global checkbox ON still worked, because FileOrganizer
  checks `self.include_client_name` first). HIGH.
- Verification of fix 2 (RAN /root/merge-scratch/run/H-A1/tkcheck.py under xvfb): a
  `tk.Button(state='disabled')` with `.bind('<Button-1>', ...)` still runs the binding (output
  ['bind']), confirming the message. There are two Start buttons bound this way:
  gui/tabs/processing_tab.py:74 (`self.app.start_btn`, disabled during a run) and
  gui/tabs/validation_tab.py:142 (a local `start_btn` that is never disabled). Both call
  `self.app.start_processing()`, so before this commit either could start a second worker thread
  during a run. `is_processing` is set True only after validation and the optional long-name
  confirmation (processing_handler.py:124) and reset in the worker's `finally` (HEAD :528). HIGH.
- Why: three defects found in review (fix 1 matches f7eb8f1 S11 item 6 character for character;
  fixes 2-3 are new). HIGH.
- What it reveals: folder/file names are built from client names with "Private"/"Limited"
  abbreviated (to keep paths short); the per-client "FolderName" toggle exists to add the client
  name to category folder names one client at a time.
- Type: bug fix (three) + tests.
- If a bug fix:
  1. Abbreviation: RAN (/root/merge-scratch/run/H-A1/t_33d0, old vs new vs HEAD):
     "ABC Private Limited" -> "ABC_Pvt_Ltd" in all three; "ABC PRIVATE LIMITED" and
     "ABC private limited": old kept them unabbreviated, new and HEAD give "ABC_Pvt_Ltd";
     "Privateer Unlimited": old "Pvter_Unlimited", new/HEAD "Privateer_Unlimited".
     REGRESSION introduced here (not mentioned in the message): names already joined with
     underscores, e.g. "ABC_Private_Limited", were abbreviated by the old substring code
     ("ABC_Pvt_Ltd") but NOT by the new code ("ABC_Private_Limited"), because "_" counts as a
     word character so \b does not match there; run-together names "ABCPrivateLimited" also
     stopped being abbreviated ("ABCPvtLtd" before). The underscore case was fixed two days later
     by 4219ffb (25-07-2026 18:45, A only, outside my range; HEAD utils/helpers.py:79-93 now
     turns runs of spaces/underscores into a space before matching); the run-together case is
     still unabbreviated at HEAD (RAN). Effect while the regression lasted: longer file and
     report names for such clients (Excel's ~218-character path limit is the stated concern in
     HEAD's comment). The report file names also use `sanitize_filename(client_info['client'])`
     (processing_handler.py "Prepare safe filenames"), so they changed the same way.
  2. Re-entrancy: proper guard. Still present at HEAD (processing_handler.py:75).
  3. Folder-setting key: proper fix; still present at HEAD (client_handler.py:181, :218).
     Similar edge still unhandled (RAN tkcheck.py): a Tk Treeview turns a value that looks like a
     whole number into an int, e.g. a client name "007" comes back as 7. All tree-derived keys
     (`get_selected_clients`, toggle, global setting) would then become "7-XX" while client_data
     holds "007-XX", so processing would fail to find that client. Applies only to purely numeric
     client names. LOW likelihood; HIGH confidence in the mechanism. The long-name warning in the
     toggle still uses a fixed 10 characters (client_handler.py:191 at HEAD), not the
     configurable maximum.
- Rule changes (business, not GST): file-name abbreviation rule old "case-sensitive substring
  'Private'->'Pvt', 'Limited'->'Ltd'" -> new "whole word, any case" (utils/helpers.py; HEAD
  :92-93 plus the separator normalisation from 4219ffb).
- Abandoned or reversed: reverses the quirk pinned by 6c1e1c6's test; partly reversed in turn by
  4219ffb (underscore handling).
- Dependency changes: none.
- Discrepancies between message and diff:
  1. "Common case ... unchanged; verified by the existing + new tests" is true for space-separated
     names, but the commit silently changed underscore-joined and run-together names (above).
- B status: B's HEAD still uses the substring replace (B utils/helpers.py:78) while its
  create_client_state_key uses the regex (B utils/helpers.py:517). READ.

#### 064c634 - 23-07-2026 08:49 - Sidhant Goyal - A only (after the split)

- Message: "Log hygiene: downgrade [DEBUG] to logger.debug; drop stray print; add pywin32". Body
  (quoted):
  "#4 hygiene sweep (no logic changes):
  - Convert 35 `logger.info("[DEBUG] ...")` calls in core/excel_handler.py and
    core/file_parser.py to `logger.debug(...)` (drop the [DEBUG] tag). These flooded the file log
    at INFO level, including full data-mapping dumps.
  - Remove the stray `print("PROCESSING: ...")` in report_processor _refresh_power_query_simple -
    it was also sitting ABOVE the docstring, turning the docstring into a dead string expression;
    removing it restores the docstring.
  - Add pywin32 to requirements.txt (win32 marker). It was a hard runtime dependency
    (pythoncom/win32com/win32gui imported unconditionally) that was missing from the manifest.
  The 14 user-visible [DEBUG] lines in the GUI processing log (via app.log_message) are
  intentionally left - they touch nested control flow and removing them is a separate UX
  decision.
  Co-Authored-By: Claude Opus 4.8 (1M context)"
- Files touched:
  - core/excel_handler.py +31/-31: 31 `logger.info(f"[DEBUG] ...")` -> `logger.debug(...)` in
    `create_report_from_template` (including "Data mappings: {data_mappings}"),
    `_create_report_win32com`, and `prepare_template_data` (Sales branch, including "Final
    Sales data prepared: {data}"). One `logger.error(f"[DEBUG] Template file does not exist:
    ...")` (line 82 at the commit, :83 at HEAD) keeps its "[DEBUG]" tag.
  - core/file_parser.py +4/-4: four Sales-related `logger.info("[DEBUG] ...")` -> `logger.debug`
    (`parse_filename` lines ~80, ~126-127; `_add_to_client_data` ~245).
  - power_query_extractor/core/report_processor.py +0/-2: deletes the comment "# TEST: Add this
    line to check if logging works" and `print(f"PROCESSING: {excel_file.name}")` that preceded
    the docstring of `_refresh_power_query_simple`.
  - requirements.txt +4/-0: comment lines and `pywin32>=305; sys_platform == 'win32'`.
- RAN counts: 35 removed `logger.info(...[DEBUG]...)` lines and 35 added `logger.debug` lines in
  the diff; at the commit `git grep "log_message.*\[DEBUG\]"` gives 14 in
  gui/handlers/processing_handler.py (matches the message). At A's HEAD 15 "[DEBUG]" strings
  remain in .py files (the 14 GUI lines + the logger.error above).
- Why: the log file was flooded and contained full path/file-name dumps at INFO level; pywin32
  was undeclared. HIGH. Matches f7eb8f1 S11 items 5 and 9.
- What it reveals: logging is configured at INFO in main.py:20-25 and
  power_query_extractor/extractor_main.py:12-18, writing `gst_organizer.log` / `pq_extractor.log`
  in the CURRENT WORKING FOLDER (READ, unchanged in my range, still at HEAD). If the app is
  started from a network folder, the log (which still contains client names and paths at INFO,
  e.g. excel_handler.py's "{cell_ref} = {value[:50]}..." lines) lands in that folder. Compare
  Standard 13 lines 1035-1040 (logs in %LOCALAPPDATA%\[App Name]\). Technical observation.
  The excel_handler win32 imports are guarded (try/except inside `if platform.system() ==
  'Windows'`, excel_handler.py:20-28), while report_processor.py:11-15 imports pythoncom,
  win32com.client, win32api, win32con, win32gui unconditionally (READ at the commit).
- Type: cleanup + dependency.
- Rule changes: none.
- Abandoned or reversed: the "[DEBUG]" INFO logging added before git (edac6a3) is mostly
  retired; B had already done the same kind of clean-up on 07-02-2026 in c258047 (its message:
  "Removed 50 [DEBUG] logger.info() lines ... Fixed dead code: unreachable return ...
  stray print()"; B's HEAD has 0 "[DEBUG]" strings, RAN grep). So A re-did in July, in Tkinter,
  part of what B did in February. HIGH.
- Dependency changes: pywin32>=305 added for Windows only. Version is a ">=" range, not an exact
  pin (Standard 13 lines 937-938 ask for exact versions). B's requirements.txt has pywin32 only
  as a comment ("# pywin32>=306"). READ.
- Discrepancies between message and diff:
  1. "drop the [DEBUG] tag": one logger.error line still has it.
  2. requirements comment "power_query_extractor and excel_handler import
     pythoncom/win32com/win32gui": excel_handler imports pythoncom and win32com (guarded), not
     win32gui. Minor.
  3. "no logic changes": true (READ).

#### ba2a753 - 23-07-2026 08:54 - Sidhant Goyal - A only (after the split)

- Message: "De-duplicate the shared 'already-open workbook' detection". Body (quoted):
  "#2, done conservatively. The two _safe_open_workbook methods (ExcelHandler, ReportProcessor)
  shared an identical "is this file already open in Excel?" loop but otherwise diverge for good
  reason - ExcelHandler opens a freshly-copied LOCAL temp file, ReportProcessor opens a NETWORK
  file that may still be locked from a Power Query refresh (needs extended \?\ paths +
  close-by-name retries).
  Extract only the genuinely-identical, testable piece into utils/excel_com.find_open_workbook
  and call it from both, so a fix to the detection now applies to both. The divergent
  open-with-fallback strategies are intentionally left in place (and cross-referenced with NOTE
  comments) - merging untestable COM code that solves two different problems would risk
  breaking the working flows; that would need real-Excel validation.
  Added tests/test_excel_com.py (3 tests, fake COM objects).
  Co-Authored-By: Claude Opus 4.8 (1M context)"
- Files touched:
  - utils/excel_com.py +29 (new): module docstring explaining why the two open strategies must
    stay different ("Do NOT merge those two strategies without validating BOTH flows against real
    Microsoft Excel"); `find_open_workbook(excel, abs_path)`: loops over `excel.Workbooks`,
    returns the first whose `Path(wb.FullName).absolute() == abs_path`, else None. READ: this is
    the same loop both methods had inline (identical comparison).
  - core/excel_handler.py +10/-6: new import (line 41 at the commit and HEAD); in
    `_safe_open_workbook` the inline loop is replaced by `find_open_workbook` plus a NOTE
    comment (HEAD :333-340).
  - power_query_extractor/core/report_processor.py +9/-5: new import
    `from utils.excel_com import find_open_workbook` (line 17 at the commit, :39 at HEAD); same
    replacement in its `_safe_open_workbook` (line ~663-670 at the commit, :709 at HEAD).
  - tests/test_excel_com.py +32 (new, 3 tests with fake Excel/Workbook objects: match found,
    not open -> None, no workbooks -> None).
- RAN (/root/merge-scratch/run/H-A1/t_ba2a, `git archive ba2a753`): 87 passed (3 + 14 + 21 + 42 +
  2 + 5).
- Why: a numbered review item ("#2") to remove duplicated code, done in the smallest safe way.
  HIGH.
- What it reveals: READ at the commit, report_processor `_safe_open_workbook` (lines 645-755)
  really does differ: it checks the file can be read, converts paths over 218 characters to
  "\\?\" or "\\?\UNC\" form, and on failure closes ANY open workbook whose file NAME equals the
  target's name (`wb.Close(SaveChanges=False)`, lines 701-709) before retrying, then tries the 8.3
  short path, then a plain path with `IgnoreReadOnlyRecommended=True`. Combined with the fact
  that the extractor obtains Excel with `win32com.client.Dispatch("Excel.Application")` (HEAD
  report_processor.py:218 and :536), which can attach to a user's already-running Excel, the
  close-by-name retry could close a user's own same-named workbook without saving. MEDIUM
  (whether Dispatch attaches to the running instance depends on the Excel setup; how to check:
  on Windows, open a workbook with the same file name from another folder, then run the
  extractor on a report that fails its first open). Compare Standard 12 line 853-854 ("Never
  take over the user's own open Excel"; use DispatchEx). Technical observation; this code is not
  changed by the commit.
  The import `from utils.excel_com import ...` inside power_query_extractor is a new
  dependency of the extractor package on the main app's top-level `utils` package; it works
  because launch_extractor.py adds the project folder to `sys.path` and main.py runs from it.
  HIGH (READ).
- Type: refactor + tests.
- Rule changes: none.
- Abandoned or reversed: the two inline loops (since edac6a3) are replaced. utils/excel_com.py and
  tests/test_excel_com.py are unchanged up to A's HEAD (RAN `git log`).
- Dependency changes: none (no new third-party import).
- Discrepancies between message and diff: none found ("mirrors, exactly" is true). The
  explanation "ExcelHandler opens a freshly-copied LOCAL temp file" became true only with
  12d505f, the same night (consistent).
- B status: B has no utils/excel_com.py (B utils/ holds __init__.py, constants.py, helpers.py).

---

## File evolution (every file touched in this range)

History per file (RAN `git log A/main -- <file>`, newest first; my range in bold):

| File | Commits (newest first) | What happened in this range, and role shift |
|---|---|---|
| PROJECT_AUDIT_ANNUALSTATEMENT.md | **f7eb8f1**, **fc279af** | Created 04-03-2026 (856 lines), rewritten 05-03-2026 (1,182 lines). Never touched again; at HEAD it documents the March state and calls bugs "confirmed" that July fixed. Role: from "audit of production code" to stale record. |
| core/file_parser.py | 56aaad9, **064c634**, **a603c02**, 6237597, edac6a3 | a603c02 removed the duplicate validation (and with it the "Invalid or corrupted Excel file" variation); 064c634 moved four Sales debug lines to DEBUG. Later 56aaad9 passes a cached stat through `_process_file`. |
| utils/helpers.py | 4f84ba9, 56aaad9, 4219ffb, **33d0baf**, **a603c02**, 6237597, edac6a3 | a603c02: scandir-based `find_excel_files`, stat-derived is_file/is_dir. 33d0baf: regex abbreviation in `sanitize_filename`. Later: scan_excel_files with parallel signature reads (4f84ba9), separator normalisation (4219ffb). Role grew from "helpers" to "performance-tuned network scanning". |
| core/excel_handler.py | **ba2a753**, **064c634**, **12d505f**, edac6a3 | First changes since the initial commit. 12d505f: local staging of reports; 064c634: DEBUG downgrade; ba2a753: shared already-open check. |
| core/file_organizer.py | **275d9a4**, edac6a3 | First change since the initial commit: Level 1 folder cached per run. Unchanged since. |
| utils/constants.py | 510f7b4, 7347383, **5c8ce95**, edac6a3 | First change since the initial commit: IMS Reco pattern accepts "(N)". Later 7347383 shortens category folder names, 510f7b4 adds workflow settings (outside range). |
| power_query_extractor/extractor_main.py | **c1583d6**, edac6a3 | First change since the initial commit: home-folder cache path first. Unchanged since. |
| power_query_extractor/gui/extractor_window.py | bf2f1dd, bfcb24a, 2aaa2b1, a677670, 2acffed, c1f425e, **c1583d6**, **275d9a4**, bd176bf, 4c42758, edac6a3 | 275d9a4: two single-listing helpers; c1583d6: cache path + mtime version choice. Later split into ExtractorPanel + window shell (a677670) and embedded as "Step 4: Extract" (bfcb24a). Most-changed file of A. |
| power_query_extractor/core/report_processor.py | c463256, 2aaa2b1, **ba2a753**, **064c634**, bd176bf, edac6a3 | 064c634: stray print removed; ba2a753: shared helper import (new dependency on top-level utils). |
| gui/handlers/processing_handler.py | 510f7b4, **33d0baf**, **0890022**, 6237597, edac6a3 | 0890022: _log_async/_progress_async; 33d0baf: re-entrancy guard. |
| gui/handlers/client_handler.py | **33d0baf**, 1c48722, edac6a3 | 33d0baf: folder-setting keys use the state code. |
| utils/excel_com.py | **ba2a753** | Created; unchanged since. |
| requirements.txt | **064c634**, **6c1e1c6**, edac6a3 | 6c1e1c6: pytest active; 064c634: pywin32 (Windows). These are the only changes after edac6a3; HEAD content = 064c634's. |
| .gitignore | **6c1e1c6**, 1cb878b | `.pytest_cache/` added. |
| pytest.ini | **6c1e1c6** | Created; unchanged since. |
| tests/README.md | **6c1e1c6** | Created; unchanged since (so at HEAD it still lists only the first four test files although HEAD has 18 test_*.py files - RAN `ls`). |
| tests/conftest.py | a677670, **6c1e1c6** | Created; changed again in a677670. |
| tests/test_file_ops.py | 4f84ba9, 56aaad9, **6c1e1c6** | Created; extended later. |
| tests/test_file_patterns.py | **6c1e1c6** | Created; unchanged since. |
| tests/test_helpers.py | 4219ffb, **33d0baf**, **6c1e1c6** | Created; quirk test replaced (33d0baf); underscore test added later (4219ffb). |
| tests/test_scan.py | **6c1e1c6** | Created; unchanged since. |
| tests/test_processing_handler.py | **33d0baf** | Created; unchanged since. |
| tests/test_excel_com.py | **ba2a753** | Created; unchanged since. |

Phase reading (inference, HIGH): March 2026 = documentation (two audit passes). 23-07-2026
00:22-01:54 = "network speed + correctness" pass (a603c02 .. 0890022). 07:58-08:54 = "safety net
then clean-up" pass (tests first, then fixes guarded by tests). The "#2" / "#4" numbering in
ba2a753 / 064c634 points to a numbered review list that is not in the repository.

## Rule evolution (this range)

- [GST RULE] IMS Reco file recognition (utils/constants.py:23): old required exactly 8 digits
  then ".xls"/".xlsx"; new also accepts an optional "(N)" (spaces allowed before it) after the 8
  digits; the captured 8-digit date (code calls it DDMMYYYY) is unchanged (5c8ce95). Other five
  patterns unchanged; they already accepted "(N)" inside their last group.
- [GST RULE] state-code logic: not changed in code; frozen by tests (6c1e1c6): Delhi->DL,
  Maharashtra->MH, Tamil Nadu->TN, West Bengal->WB; unknown multi-word name -> initials (can
  collide, e.g. "UP"); unknown single word -> first three letters upper-case; "" -> "".
  TIME-VARYING (state/UT names and codes can change). These are the app's 2-letter folder codes.
- [GST RULE] Power Query Extractor cell mappings (power_query_extractor/config/cell_mappings.py,
  including "Table 6J Match" 'T6 of R9' K29 and "Table 8D Limits" 'T8 of R9' G6, TIME-VARYING):
  unchanged in code in this range; documented in fc279af, dropped from the document in f7eb8f1.
- Business (not GST) naming rule "Private"->"Pvt", "Limited"->"Ltd" in file names
  (utils/helpers.py `sanitize_filename`): case-sensitive substring (edac6a3 .. 6c1e1c6) ->
  whole word any case (33d0baf) -> whole word any case after turning "_" into spaces (4219ffb,
  later). `create_client_state_key` (folder names) used whole word any case all along.
- Per-client "include client name in category folder" setting key: "client-StateName" (broken)
  -> "client-STATECODE" (33d0baf).
- "Latest Version folder" in the extractor: highest name (text order of "Version-DDMMYY HHMM")
  -> newest modified time (c1583d6).
- "Latest Annual Statement folder" for rerun/resume in the organizer: newest modified time
  (unchanged), now computed once per run (275d9a4).
- Which files a scan accepts: unchanged on Windows; case-insensitive extensions everywhere
  (a603c02); invalid/corrupted files still silently skipped.
- Report building location: in place on the target -> local %TEMP%\gst_excel_temp then one copy
  to the target (12d505f).
- Start-processing rule: may now start only when no run is in progress (33d0baf).

## Edge cases discovered in real use (from this range)

| Commit | What broke / the scenario | How fixed | Similar cases still open (HEAD) |
|---|---|---|---|
| a603c02, 12d505f, 275d9a4 | Slow scans and report building on a network/OneDrive/UNC share (each file-system call a round-trip; Excel editing in place over the network). | Fewer calls; local staging; per-run caching. Behaviour kept. | Silent dropping of invalid files (utils/helpers.py:275-366); partial destination file if the final copy fails; temp leftovers if Excel holds the file. |
| 5c8ce95 | IMS Reco downloads named "... (1).xlsx" not recognised. | Pattern widened (proper fix). | Leading "(1) ImsReco-..." and "- Copy" names still unrecognised; when original and "(1)" copy both exist, the "(1)" copy is linked into the ITC report (sort order). |
| c1583d6 | Extractor never remembered the target folder (cache path mismatch; found by the f7eb8f1 audit rather than reported from use - MEDIUM). | Read the writer's path. | Stale CWD fallbacks kept in extractor_main. |
| c1583d6 | After a re-run crossing a month boundary the extractor could open an older Version folder. | Choose by modified time. | Modified time moves when anything is written in an old folder (lock files, refreshed copies); B still chooses Level 1 by name. |
| 0890022 | Dry-run preview listed the last client's name/files for every client; processing log could show the next client's name. | Values bound at scheduling time. | Worker thread still calls messagebox and button `.config` directly. |
| 33d0baf | Double/second click on Start (two Start buttons, bindings fire when disabled) could start a second processing thread. | is_processing guard. | Guard depends on is_processing being reset in `finally` (it is). |
| 33d0baf | Per-client "FolderName" toggles ignored. | Same key everywhere. | Numeric-looking client names change type in the Tk tree (e.g. "007" -> 7). |
| 33d0baf | ALL-CAPS "PRIVATE LIMITED" not abbreviated; "Privateer" mangled. | Regex with word boundaries. | Introduced: underscore-joined names not abbreviated (fixed 4219ffb); run-together "ABCPrivateLimited" still not abbreviated. |
| 064c634 | Log file flooded with INFO "[DEBUG]" lines incl. path dumps; pywin32 missing from requirements. | Downgraded to DEBUG; requirement added. | Log files still written to the working folder; 14 "[DEBUG]" lines still shown in the GUI log. |

## Abandoned approaches (this range)

- Validating each Excel file twice (once in discovery, once in `_process_file`) and the
  "Invalid or corrupted Excel file" variation it produced - removed (a603c02).
- Three glob passes for discovery - replaced by one directory listing (a603c02), itself later
  replaced by `scan_excel_files` (outside range).
- Creating/editing reports directly on the network target - replaced by local staging (12d505f).
- Per-client recomputation of the Level 1 folder - replaced by a per-run cache (275d9a4).
- Per-client glob passes in the extractor - replaced by single-listing helpers (275d9a4).
- Name-ordered Version folder choice - replaced by modified time (c1583d6).
- CWD-relative cache file in the extractor window - replaced (c1583d6).
- Inline `root.after(0, lambda: ...)` UI updates - replaced by binding helpers (0890022).
- Case-sensitive substring abbreviation in `sanitize_filename` - replaced (33d0baf), then
  refined (4219ffb).
- `logger.info("[DEBUG] ...")` file logging - downgraded (064c634).
- Duplicated "already open?" loops - shared helper (ba2a753).
- Documentation: the v1 audit text (fc279af) - rewritten (f7eb8f1); the v2 audit then left
  stale (never updated after the July fixes).

## Dependency evolution (this range)

| Commit | Change | Reason |
|---|---|---|
| 6c1e1c6 | `pytest>=7.0.0` moved from a comment to an active line in requirements.txt | New test suite. (Standard 13 says the app's requirements.txt should not list pytest and should use exact versions.) |
| 064c634 | `pywin32>=305; sys_platform == 'win32'` added | It was a hard runtime need of the extractor (unconditional imports at report_processor.py:11-15) and of the COM report path; missing since edac6a3; recommended verbatim by f7eb8f1 S11 item 5. |
| a603c02 | standard-library `stat` (S_ISREG, S_ISDIR), `os.scandir` | Performance; no new third-party package. |
| 275d9a4 | `import os` in extractor_window.py | os.scandir. |
| ba2a753 | extractor imports the main app's `utils.excel_com` | Shared helper; internal coupling, not third-party. |

Unchanged: openpyxl>=3.1.2, lxml>=4.9.0. At B's HEAD: PyQt5>=5.15.0 active, pywin32 only as a
comment, no pytest line (READ B requirements.txt).

## Observations against the author's standards (technical only, for the audit writers)

- Standard 12 "Excel through COM" (rebuild_pipeline_FINAL.md 844-878) vs code touched in this
  range: report creation attaches to a running Excel via GetActiveObject and always quits it
  (core/excel_handler.py:124, :240-244, `self.excel_app` never set); DisplayAlerts is switched
  back on before Quit (:240); CoUninitialize runs in the same function that still holds Excel
  objects (:248-251); no `IgnoreReadOnlyRecommended`/`ReadOnly` check when opening the staged
  copy; the extractor uses `Dispatch` not `DispatchEx` (report_processor.py:218, :536) and can
  close same-named workbooks unsaved (close-by-name retry). The openpyxl fallback saves a copy of
  a Power Query template (Standard 12 846-850).
- Standard 13 (900-1061): logs go to the working folder (main.py:24,
  power_query_extractor/extractor_main.py:16); requirements use ">=" ranges and include pytest;
  pytest.ini lacks `pythonpath = .` (conftest.py inserts the path instead).

## Possible client names (locations only; nothing copied)

- PROJECT_AUDIT_ANNUALSTATEMENT.md as added by f7eb8f1 (identical at A's HEAD), lines 130, 133,
  134, 158, 160, 161, 162, 177, 179, 180, 182, 227, 228, 231, 232: the "Concrete trace" examples
  use one company-like name that is also the name of a well-known real business group. Probably
  an illustrative example chosen by the AI auditor, but it is not an obviously generic name.
  (fc279af's version does not contain it; RAN grep count 0.)
- No other candidates found in the diffs or messages of the twelve commits: the tests and
  messages use generic "ABC", "XYZ", "CA"-style names, "Privateer Unlimited", "Unknown Place",
  "A Very Long Client Name That Exceeds The Limit For Sure"; the v1 audit uses "ABCLtd",
  "ABC Pvt Ltd", "ABC Private Limited Industries", "ABC-DEF Corp"; the v2 audit "ABC-XYZ Ltd".

## Questions for the author

1. The PROJECT_AUDIT_ANNUALSTATEMENT.md "Concrete trace" examples (lines listed above) use a
   real-business-like name. Is that one of your clients, or just an example the AI picked?
2. That audit document has not been updated since 05-03-2026 and now describes problems fixed in
   July as open. Keep it as a historical record, update it, or retire it in the merged tool?
3. Both audits call c258047/a62d3cc/c21f788 an "unmerged branch"
   (copilot-worktree-2026-03-04T11-52-56). Is that the same work as AnnualStatement_oldv2 (B)?
   Did your local branch have only those three commits, or also the 08-02 and 09-02 ones?
4. 12d505f asked you to "validate one report on a Windows machine before bulk processing". Did
   you, and did the ITC/Sales reports' Power Query still refresh correctly afterwards? Do any
   queries in your templates use the workbook's own location (rather than the Links-sheet
   paths)?
5. When both "ImsReco-...-<date>.xlsx" and "ImsReco-...-<date> (1).xlsx" exist for a client,
   which one should be linked into the ITC report? (Today the "(1)" copy sorts first and is used.)
6. Do you or your team open files that sit directly in an older "Annual Statement-..." or
   "Version-..." folder before re-running or using the extractor? (Modified time decides which
   folder counts as latest, and opening an Excel file there can change it.)
7. Between 23-07-2026 08:48 and 25-07-2026 18:45 underscore-joined names like
   "<Client A>_Private_Limited" were not shortened to "Pvt"/"Ltd". Did any real runs happen in
   that window (folder/file names would be longer)? Should run-together names
   ("<Client A>PrivateLimited") also be shortened?
8. Should the "Folders Created" and "Successfully Copied" counts in each client's
   _Organization_Report text file be per client (today they are cumulative for the whole run)?
9. Are any client names purely numeric (e.g. starting with zeros)? The client list widget turns
   them into numbers, which would break selection for that client.
10. Files that fail the Excel checks (too small, not a real .xlsx/.xls inside) are silently left
    out of a scan. Would you want them listed (e.g. in Variations) so you notice them?
11. Do you want pytest listed in the app's requirements.txt, and should versions be exact pins
    (your Standard 13), rather than ">=" ranges?
12. The 14 "[DEBUG]" lines still shown in the on-screen processing log (Sales report steps): keep
    or remove?

## Limitations (not INCOMPLETE - everything assigned was read)

- The Windows-only parts (Excel through COM in core/excel_handler.py and
  power_query_extractor/core/report_processor.py) were traced by reading only.
- Claims of "byte-identical characterization snapshots", "9 folder layouts" equivalence tests,
  "mutation-checked" tests and the closure-reproduction test are not in the repository and could
  not be verified.
- For f7eb8f1 I read both complete versions of the file instead of paging through the 1,862-line
  diff text; together they contain every removed and added line.
