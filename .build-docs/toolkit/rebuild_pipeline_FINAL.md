# The Complete Rebuild Pipeline — FINAL VERSION

> **When to use this:** You have an existing tool that you want to rebuild
> from scratch with new perspective. You'll first audit the old code, then
> make decisions about what to keep/change/drop, then build fresh.
>
> This document covers:
> 1. What documents you prepare BEFORE giving the rebuild prompt
> 2. The Rebuild Prompt itself — what goes in it, including the Build Plan
>    (always saved, even in auto mode) and the nine-step Build Loop
> 3. The Structured Prompting Style for mid-build changes

---

# PART 1: DOCUMENTS YOU PREPARE BEFORE THE BUILD

You will provide Claude Code with **four documents** when starting a rebuild.
Three you prepare upfront, one grows automatically during the build.

**If you talk the rebuild through with Claude first,** the Rebuild Brief is
compiled from your answers (for example with the Rebuild Brief Assistant), and
Claude can pre-fill the project sections of CLAUDE.md and Session 0 of the
Decisions Log from that discussion. You check every line. The *How We Build*
and *Coding Standards* sections of CLAUDE.md are copied word for word, never
shortened.

**Claude Code then creates and keeps two more** at Step B of the Rebuild Prompt:
`.build-docs/BUILD_PLAN.md` (the Build Plan — always saved, even in auto mode)
and `.build-docs/RESUME_CONTEXT.md` (where we are and the exact next action).
Each sub-phase adds a few files of its own (see *Files the loop produces* in
CLAUDE.md).

---

## Document 1: The Audit Report (already done)

**File:** `.build-docs/PROJECT_AUDIT_[PROJECT_NAME].md`
**Created by:** The Project Archaeology prompt (`Project_Audit_FINAL.md`)
**Your role:** Read it carefully. Highlight or note sections you agree with,
disagree with, or want to change. You don't modify this file — it's your
reference material.

---

## Document 2: The Rebuild Brief (YOU write this)

**File:** `.build-docs/REBUILD_BRIEF_[PROJECT_NAME].md`
**Created by:** You, after reading the audit
**Purpose:** This is your decision document. It tells Claude Code what to keep,
what to change, what to drop, and what to add new.

**Reading guide for the audit:** Start with Section 1 (Executive Summary) for
the overview. Then go directly to Section 7 (Gaps) — that's where the rebuild
value is. Then read Section 5 (Business Rules) to verify your domain knowledge.
Then Section 11 (Reconstruction Blueprint) for the suggested build order.
Skim the rest as needed.

### Template:

```markdown
# Rebuild Brief: [PROJECT_NAME]
# Date: [TODAY'S DATE]
# Author: Adv. Sidhant Goyal

---

## 1. WHAT THIS TOOL DOES (in my own words)

[Write 3-4 sentences describing the tool's purpose. Don't copy from the audit —
restate it in YOUR understanding. If your understanding differs from the audit's
description, that gap itself is valuable information.]

---

## 2. WHO USES THIS AND HOW

- **Primary user:** [e.g., "Me and my team at Goyal Tax Services"]
- **Technical comfort:** [e.g., "Non-coders. Must work with simple clicks, no terminal commands."]
- **Frequency of use:** [e.g., "Daily during filing season, monthly otherwise"]
- **Environment:** [e.g., "Windows 10/11 desktops, files on local drives and sometimes network server"]

---

## 3. FEATURE DECISIONS

For each feature from the audit (Section 3), mark your decision:

| Feature | Audit Ref | Decision | Notes |
|---------|-----------|----------|-------|
| [Feature name] | F1 | KEEP AS-IS | [Why — what works well] |
| [Feature name] | F2 | KEEP BUT IMPROVE | [What to change and why] |
| [Feature name] | F3 | DROP | [Why — no longer needed or was never useful] |
| [Feature name] | F4 | REBUILD DIFFERENTLY | [What was wrong and what you want instead] |

---

## 4. NEW FEATURES TO ADD

Features that didn't exist in the original but you want in the rebuild:

| New Feature | What it should do | Why I need it | Priority |
|-------------|-------------------|---------------|----------|
| [Name] | [Plain English description] | [The problem it solves] | MUST HAVE / NICE TO HAVE |

---

## 5. BUSINESS RULES — KEEP / CHANGE / ADD

Reference audit Section 5. For each rule:

### Rules to KEEP exactly as-is:
- [Rule from audit] — (Audit: file.py, line XX)

### Rules to CHANGE:
- **Old rule:** [What the audit found]
- **New rule:** [What it should be now]
- **Why:** [Regulation changed / original was wrong / better understanding now]

### Rules to ADD (not in original):
- [New rule and why it's needed]

### Contracts — promises the rebuild must always keep:
Numbered promises that hold on every run, whatever the input. Each one gets
a contract test, and every sub-phase plan names the contracts it relies on.
- C1: [e.g., "Every invoice row in the input ends up in the output or in the error report — never silently dropped"]
- C2: [e.g., "Output totals plus the error-report totals equal the input totals; any rounding difference is shown, never hidden"]
- C3: [e.g., "The tool never changes a client's figure without listing the change in the report"]

---

## 6. EDGE CASES — LESSONS FROM THE AUDIT

Reference audit Sections 6 and 7.

### Must handle (were handled in original — keep):
- [Edge case from audit Section 6]

### Must handle (were NOT handled — gaps from Section 7 that bit you):
- [Gap that caused real problems]

### Not worth handling (over-engineering for my use case):
- [Edge case that's theoretically possible but practically irrelevant]

---

## 7. DESIGN PREFERENCES

- **Framework:** [e.g., "Move to PySide6" or "Keep PyQt5 as it is"]
- **Interface style:** [e.g., "Single window, big buttons, minimal tabs"]
- **Output format:** [e.g., "Excel (.xlsx) with formatting, not just raw CSV"]
- **File handling:** [e.g., "Drag-and-drop preferred over file browse dialogs"]
- **Error display:** [e.g., "Show errors in the main window, not pop-up dialogs"]

---

## 8. WHAT I LEARNED (from reading the audit)

Write 3-5 sentences about what surprised you, what you'd do differently,
and what principles you want the rebuild to follow. This is your "new
perspective" section.

---

## 9. PARKING LOT

[Leave empty. Ideas that pop up during the build go here.]

| Idea | When it came up | Priority | Status |
|------|----------------|----------|--------|
| | | | |

---
```

---

## Document 3: CLAUDE.md (project memory, how we build, coding standards)

**File:** `CLAUDE.md` (placed in the project root — MUST be root, not a subfolder)
**Created by:** You, before starting the build — or Claude pre-fills the project
sections from your discussion
**Purpose:** Claude Code reads this automatically at the start of every session.
It contains your identity, project context, how we build (the Build Plan rule and
the nine-step Build Loop), coding standards, and documentation rules.

### Template:

```markdown
# Project: [PROJECT_NAME] — Rebuild

> **Filling this template (for Claude):** Fill the parts in [square brackets] in the
> sections above *How We Build* from our project discussion: project name, tech stack,
> reference documents and current phase. Copy *How We Build* and *Coding Standards*
> word for word — never shorten, summarise or reword them. Delete this note once the
> file is filled.

---

## About Me
I am a CA/GST consultant, not a programmer. I build tools through
conversation with AI. Explain technical decisions in plain English.
Ask before making architectural choices. Never assume I understand
technical jargon — explain it on first use.

## Tech Stack
- **Language:** Python 3 — the default for every app (Standard 12 says when
  another language is allowed)
- **Window toolkit:** PySide6 (the `PySide6-Essentials` package, Qt 6)
  [— or PyQt5, only for an existing PyQt5 app that is being kept as it is]
- **Other libraries:** openpyxl for Excel files without Power Query; Excel itself
  through COM (`pywin32`) for workbooks with Power Query queries that the app
  writes into or refreshes (Standard 12); [others from Standard 12's table]
- **Runtime:** the office's shared portable Python in `_runtime\python\`,
  next to the apps in the shared apps folder — nothing is installed on
  users' PCs (Standard 13)
- **Shared apps folder:** [e.g. \\server\share\Apps]
- **How users start it:** a desktop shortcut to `Run [App Name].bat`, set to
  Run: Minimized (Standard 13)

## Reference Documents
- `.build-docs/PROJECT_AUDIT_[name].md` — Full audit of the original tool
- `.build-docs/REBUILD_BRIEF_[name].md` — My decisions for the rebuild
- `.build-docs/BUILD_PLAN.md` — The Build Plan: tracker, structure, phases and sub-phases,
  change log. Always written and saved before any code — even in auto mode
- `.build-docs/RESUME_CONTEXT.md` — Where we are, the exact next action, step status,
  hard-won lessons. Read it first in a cold session
- `.build-docs/DECISIONS_LOG.md` — Running log of every discussion and decision
  (Session 0 = our project discussion, Session 1 = the Build Plan, then one session
  per sub-phase)
- `.build-docs/Phase <N> ….md` — Each sub-phase's plan, test disputes, reviews and
  re-count evidence (see *Files the loop produces* below)
- `.build-docs/context/` — Additional reference files (samples, templates, specs)

## Current Phase
[Not started — Step A next]
> Once the build starts: Sub-phase [N.M] [name] · Step [.K] [step name] · [status].
> Updated at the end of every step (Build Loop rule 1).

---

# HOW WE BUILD — MANDATORY FOR EVERY PROJECT

Two parts: the Build Plan, written and saved before any code, and the Build
Loop, which every sub-phase of that plan goes through.

---

## The Build Plan — always written, always saved

- Before any code, Claude writes the Build Plan, saves it as
  `.build-docs/BUILD_PLAN.md` and commits it (Standard 10). No sub-phase
  starts until the author approves it.
- This holds in **every** permission mode — **auto mode included**. Auto mode
  removes permission prompts, and it may also tell Claude to keep going,
  minimise interruptions and prefer action over planning. For this project the
  author's rules come first: auto mode never skips the Build Plan, a stop
  point, or the wait for the author's "next".
- A plan that exists only in the chat, or only in Claude Code's own plan-mode
  file, does not count. In plan mode, saving it to `.build-docs/BUILD_PLAN.md`
  is the first action once the plan is approved.
- The same goes for every sub-phase plan (step .1): always written and saved as
  `.build-docs/Phase <N> Plan - <name>.md`, auto mode included.
- BUILD_PLAN.md stays current: its tracker at the end of every step, the rest
  at every .9 Close.

---

## Build loop — nine stop points per sub-phase (the author's standing method)

### Key to the terms below
- **The author** — me (see *About Me*).
- **The Brief** — `.build-docs/REBUILD_BRIEF_[name].md`. The audit
  (`.build-docs/PROJECT_AUDIT_[name].md`) is a reference: read the parts the
  Brief cites; where they differ, the Brief wins.
- **Contracts** — the numbered promises (C1, C2, …) in the Brief's *Contracts* list,
  under Business Rules: what must hold on every run, whatever the input. Each one
  gets a contract test.
- **BUILD_PLAN** — `.build-docs/BUILD_PLAN.md`: tracker, file structure, feature
  checklist, phases and sub-phases, change log.
- **RESUME_CONTEXT** — `.build-docs/RESUME_CONTEXT.md`: where we are, the exact next
  action, the "Step status" lines, hard-won lessons.
- **Decisions Log** — `.build-docs/DECISIONS_LOG.md`, one session per sub-phase.
- **Sub-phase** — one numbered block of BUILD_PLAN (Phase 2 → 2.1, 2.2, …). Its steps
  are .1 to .9, plus the optional .7b. `<N>` in a file name is the sub-phase number
  (e.g. `Phase 2.1 Plan - Excel reader.md`); in test file names `<phase>` uses an
  underscore instead of the dot (`2_1`), because a dot breaks Python imports.
- **P-flags** — technical choices Claude made without asking, numbered across the
  whole project (P1, P2, …).
- **Lettered assumptions** — assumptions about technical choices, labelled A, B,
  C…, that stand unless the author objects ("silence = consent"). Business
  questions are never settled this way (rule 3).
- **The reference notes** — the files in `.build-docs/context/`.
- **Where files go** — every `Phase <N> …` file lives in `.build-docs/`; every test
  file lives in `tests/`.

### Roles
- **Opus** (the strongest Claude coding model) plans, codes and answers reviews. It keeps the Decisions Log,
  RESUME_CONTEXT and the git commits, and runs every multi-agent workflow.
- **Fable** (the second Claude model) writes the tests **before** any code exists and settles every disagreement.
  A different model writes the tests, so a misreading the coder shares gets caught. Keep its role narrow:
  it costs more and hits usage limits sooner.
- **Codex** (ChatGPT, on the author's OpenAI quota) checks each plan, writes the break-it tests and reviews the code.
  - It must run the model and effort set in `~/.codex/config.toml`. **Never pass `--model` or `--effort`.**
    If Codex reports any other model or effort, the break-it tests go back to Fable.
  - Drive Codex **from Bash through the codex-companion script** (or through the codex-rescue agent with its
    model set to Opus). Never through a Sonnet wrapper. Of the Claude models, only Opus and Fable are used,
    sub-agents included.
  - **Never rush Codex.** Run it in the background with no time limit, and poll
    `status <job> --wait --timeout-ms 540000` until it ends. The 9-minute poll is only a tool limit,
    never a limit on Codex. Fetch the output with `result <job>`.
  - Verify the model and effort Codex actually used from its session log (`~/.codex/sessions/**/rollout-*.jsonl`
    or `~/.codex/logs_2.sqlite`), and record it.
  - **Word Codex briefs plainly** ("damaged or unusual files", not "hostile", "attack" or "zip bomb"),
    or OpenAI's content filter may stop the run. If it does stop, resume the same session once.
- **Opus may never edit, skip, weaken or delete a Fable or Codex test.** If it believes a test is wrong,
  it leaves the test red, logs a dispute, and Fable decides.

### Rules for every step
1. **One step per "next".** Each step ends with: a checkpoint commit (Standard 10; the commit guarded so it never
   runs if the status update failed, e.g. PowerShell `if ($LASTEXITCODE -ne 0) { exit 1 }`); one status line in
   RESUME_CONTEXT ("Step status"), CLAUDE.md *Current Phase*, the Decisions Log session and the BUILD_PLAN tracker;
   a plain-English report; then STOP until the author says "next".
2. **Empty steps skip themselves** (.4 with no disputes, .7 with no findings, .8 with nothing to settle).
3. **Business questions** (anything that changes what the app does with client work) are asked one at a time, in
   plain English, with a recommended default, before the plan depends on them. **Technical choices** are made,
   logged as P-flags or lettered assumptions ("silence = consent") and reported. Never decide a business question
   in code.
4. **Phase plans are never trimmed for size.** Precision wins over length; no context lost, nothing left to guesswork.
5. **Same conversation.** "Start afresh" means a new usage window, never a new conversation. Never tell the author
   to open one; the conversation compacts itself. RESUME_CONTEXT and CLAUDE.md must still let a cold session pick up.
6. **If a usage limit stops a run:** finished files stay on disk. Resume with "finish, do not restart" prompts,
   reusing cached agents where possible. Never redo finished work.
7. **Real data:** never treat a folder as a reference without asking the author what it is and how it was made.
   When real files are needed, name exactly which (company, export, period) in a checklist file with Done / Notes /
   Answer columns, and the author downloads them by hand.
8. **Heavy steps may be split** (e.g. .3a code to green, .3b real-data gate), with a commit and a stop between.
9. Before reporting "done", run the suite yourself and show the exact count. Never claim a pass you did not see.

### The nine steps

| Step | Who | What happens | Output (done when) | Skips when |
|---|---|---|---|---|
| **.1 Plan + plan check** | Opus → Codex → Opus | Ask any business question first. Read to the last line: CLAUDE.md, the Brief, the newest Decisions Log session, RESUME_CONTEXT, BUILD_PLAN's block for this sub-phase, and the reference notes. Optionally, read-only fact-gatherers first check the code, the rules and the real files. Opus writes `.build-docs/Phase <N> Plan - <name>.md`: proof of reading; plain-English context (and what it is NOT); prerequisites, spikes and questions; the Brief contracts it relies on; the edge cases it owns, each with a test; the **locked `file ← test_file` order**; for each file the exact public API (signatures, return shapes, codes, exact texts, constants), its rules and edge cases; contract tests; test data; the real-data gate; the exit gate; open flags and P-flags (numbered on from the last used). Codex checks it read-only. Opus verifies each Codex point, then fixes it or rebuts it with evidence, and appends an answers table. | Plan status "approved for tests"; every Codex point fixed or answered; commit | — |
| **.2 Tests first** | Fable | Writes every test file from the approved plan and the Brief, in the locked order, before any code. Imports modules under test **lazily inside fixtures**, so a missing module fails only its own tests and never breaks collection of the suite. Synthetic files go in temp folders; real fixtures are read-only and the test SKIPS when the data is absent. A coverage pass checks every plan rule has a test, removes over-specification and duplicates, and checks the test files agree with each other. | New tests red **for the right reason only** (the code is missing); every earlier test still green; commit | — |
| **.3 Code** | Opus (Fable for the real-data gate) | Codes in the locked order, each file's tests green before the next, to every coding standard in CLAUDE.md. A test Opus believes wrong is left red and logged in `.build-docs/Phase <N> test disputes.md` (what it asserts, why it is wrong, exact quotes, the code as built, the change if Fable rules for the test). Technical choices made while coding are saved verbatim at the end of the plan. **Real-data gate**, if the plan has one: (1) the app runs in print/preview mode; (2) Fable re-counts independently with its own script, without reading the app's code; (3) every discrepancy is investigated file by file (a fault in the re-count is fixed; a fault in the app goes to Opus; never freeze a wrong number); (4) Fable freezes the anchors in its own test; (5) the gate passes in hard-assert mode, timed. The evidence goes in an evidence file, and the re-count script is preserved. | Suite green except logged disputes; gate passed; commit | — |
| **.4 Disputes** | Fable | Rules on each dispute from primary evidence (the spec, the real files, library source): "test is wrong" (Fable corrects its own test, minimally, keeping the guard's purpose) or "test stands" (Opus changes the code). The ruling is appended to the disputes file. | Every dispute ruled; suite green; commit | No disputes |
| **.5 Break-it tests** | Codex | Writes **new files only** (`tests/test_breakit_<phase>_*.py`), asking "how could this fail silently?" against the Brief's edge cases and contracts. Plan-silent cases are marked `# SPEC-SILENT`. Opus then runs the suite (Codex's sandbox may not be able to run tests) and classifies every failure: real defect, spec-silent, or test error. | New tests added; failures classified; commit | — |
| **.6 Fixes + code review** | Opus → Codex | Opus fixes the real defects. Spec-silent failures are either hardened (refusing an ambiguous input rather than guessing is usually right) or disputed for .8. Profile anything slow before blaming the disk. Then Codex reviews all the sub-phase's code, read-only, with no time limit; the findings ([BLOCKING] / [should-fix] / [minor]) are saved verbatim in `.build-docs/Phase <N> Codex code review.md`. **Ask the author** whether a second review in adversarial mode is wanted (step .7b). | Fixes in; review saved; commit | — |
| **.7 Answers** | Opus | Reproduces each finding first, then either fixes it test-first (a regression test in `tests/test_regressions_<phase>.py` that fails on the old code and passes on the new), rebuts it with evidence, or hands it to Fable when the fix belongs in Fable's or Codex's test files. An independent Opus checker rebuilds each scenario and confirms each fix on the old and the new code. An answers table is appended to the review file. Any change to what the app does with client work goes to the author as a question. | Every finding fixed, rebutted or handed on; suite re-run; commit | Codex found nothing |
| **.7b Adversarial review** (optional, the author's call) | Codex → Opus | Codex's adversarial-review mode (`adversarial-review --background --base <ref> "<plain focus text; known deferred items excluded>"`), over this sub-phase or over several together. Saved verbatim; Opus answers as in .7, with an independent checker. When a fix overturns a plan rule, the plan is rewritten in place, and tests pinned to the old rule become disputes for .8. | Findings answered; commit | Author declines |
| **.8 Settle** | Fable | Settles every rebuttal and every remaining dispute from primary evidence, and applies the review items handed to it. Opus applies any ruling that goes against the code. | Suite fully green; rulings logged; commit | Nothing to settle |
| **.9 Close** | Opus (+ author if a live sitting) | The live sitting first, if the sub-phase has one. Then a read-only gate check with evidence (suite, real-data gate timed, contract tests, fixture tests ran and did not skip). In parallel, on separate files: (a) the Decisions Log session entry (built; decisions; P-flags from the plan and the coders' saved choices; assumptions; disputes; break-it, review and adversarial outcomes; superseded items; open items; status with the suite count and version); (b) small tidy-ups the checkers listed; (c) BUILD_PLAN (tracker, the sub-phase's status, carry-forwards into later blocks, change log) **and the next sub-phase broken into its nine steps, not started**; (d) RESUME_CONTEXT (where we are, the exact next action, new hard-won lessons) and CLAUDE.md *Current Phase*. A cross-document consistency critic then hunts contradictions; a fixer applies the real ones. Bump the version if the author can now see or run something new. | Sub-phase closed; commit; the author pauses before the next sub-phase | — |

### Files the loop produces (per sub-phase)
`Phase <N> Plan - <name>.md` · `Phase <N> test disputes.md` · `Phase <N> Codex code review.md` ·
(optional) `Phase <N> Codex adversarial review.md` · (if a real-data gate) `Phase <N> <gate> re-count.md` plus the
preserved re-count script · `tests/test_breakit_<phase>_*.py` · `tests/test_regressions_<phase>.py` ·
one Decisions Log session per sub-phase · RESUME_CONTEXT "Step status" lines.

---

# CODING STANDARDS — MANDATORY FOR ALL CODE

These standards apply to EVERY file and EVERY function in this project.
No exceptions. No shortcuts. These exist so that a non-programmer can
read the code six months later and understand what everything does,
why it exists, how it connects to everything else, and what assumptions
were made.

---

## Standard 1: Modular Folder Structure From Day One

NEVER build in a single file. NEVER start flat and split later.
Always use folder-based structure from the very first line of code:

```
my-tool/
├── Run My Tool.bat              ← What the desktop shortcut runs (Standard 13)
├── main.py                      ← Entry point ONLY. Started by the .bat, never by users.
│
├── config/                      ← All settings, constants, mappings
│   ├── __init__.py
│   ├── settings.py              ← Paths, thresholds, environment config
│   ├── constants.py             ← Fixed values (tax rates, state codes, limits)
│   └── mappings.py              ← Lookup tables (state code → name, HSN → rate)
│
├── core/                        ← Business logic, one file per domain area
│   ├── __init__.py
│   ├── gstin_validator.py
│   ├── tax_calculator.py
│   ├── invoice_validator.py
│   └── return_classifier.py
│
├── models/                      ← Data structures / object definitions
│   ├── __init__.py
│   ├── invoice.py               ← What an Invoice object looks like
│   ├── error_entry.py           ← What an Error record looks like
│   └── client.py                ← What a Client record looks like
│
├── readers/                     ← Input handling, one file per format
│   ├── __init__.py
│   ├── excel_reader.py
│   ├── csv_reader.py
│   └── json_reader.py
│
├── writers/                     ← Output generation, one file per type
│   ├── __init__.py
│   ├── excel_writer.py
│   ├── report_writer.py
│   └── json_writer.py
│
├── gui/                         ← Interface, one file per screen/dialog
│   ├── __init__.py
│   ├── main_window.py
│   ├── settings_dialog.py
│   ├── error_viewer.py
│   └── progress_bar.py
│
└── utils/                       ← Shared helpers, split by concern
    ├── __init__.py
    ├── file_helpers.py
    └── format_helpers.py
```

Not every project needs every folder. Propose only what's relevant.
But NEVER combine responsibilities into one file.

Each file does ONE thing. If you can't describe its purpose in one
sentence, split it. Before creating any new file or reorganising
the structure, tell me and get approval.

---

## Standard 2: Naming Conventions

### Files and Folders
- All lowercase with underscores: `gstin_validator.py`, `excel_reader.py`
- Folder names short and descriptive: `core/`, `gui/`, `config/`
- NO abbreviations unless universally known: `gst` is fine, `inv_val` is not
  (use `invoice_validator` instead)

### Functions and Variables
- All lowercase with underscores (snake_case):
  `validate_gstin()`, `tax_rate`, `input_file_path`
- Function names start with a VERB: `validate_`, `calculate_`, `read_`,
  `write_`, `export_`, `parse_`, `format_`, `check_`, `get_`, `set_`
- Boolean variables start with `is_` or `has_`:
  `is_valid`, `has_errors`, `is_interstate`
- NEVER use single-letter variables except `i`, `j` in short loops
- NEVER use vague names: `data`, `result`, `temp`, `info`, `stuff`
  Instead: `invoice_data`, `validation_result`, `temp_file_path`

### Classes
- PascalCase (capitalise each word): `InvoiceValidator`, `ExcelReader`,
  `MainWindow`, `GSTINResult`

### Constants
- ALL_UPPERCASE with underscores: `MAX_FILE_SIZE`, `GST_RATES`,
  `STATE_CODES`, `B2CL_THRESHOLD`
- ALL constants live in `config/constants.py` — NEVER hardcode values
  directly in logic files

---

## Standard 3: File-Level Documentation

Every Python file must start with this header block:

```python
"""
FILE: core/gstin_validator.py

PURPOSE: Validates GSTIN (Goods and Services Tax Identification Number)
         format, structure, and check digit.

CONTAINS:
- validate_gstin()     — Main validation, returns valid/invalid with details
- extract_components() — Breaks GSTIN into state, PAN, entity, check digit
- verify_checksum()    — Runs the government-specified check digit algorithm

DEPENDS ON:
- config/constants.py  → STATE_CODES, GSTIN_LENGTH, CHECKSUM_CHARS
- utils/format_helpers.py → clean_string()

USED BY:
- core/invoice_validator.py → calls validate_gstin() for each invoice row
- gui/error_viewer.py → displays validation error messages

CHANGE LOG:
| Date       | Change                                    | Why                              |
|------------|-------------------------------------------|----------------------------------|
| DD-MM-YYYY | Created — initial validation rules        | Section 25 CGST Act compliance   |
| DD-MM-YYYY | Added old state code "00" handling         | Client data contained code "00"  |
| DD-MM-YYYY | Added composition taxpayer identification  | Needed for GSTR-4 projects       |
"""
```

---

## Standard 4: Function-Level Documentation

Every function MUST have a documentation block with ALL of these fields:

```python
def validate_gstin(gstin_string, strict_mode=False):
    """
    WHAT:
        Validates a GSTIN string against format rules and checksum.
        Returns a result dict with valid/invalid status, error details,
        and extracted components (state, PAN, entity type).

    WHY ADDED:
        Clients frequently submit invoices with malformed GSTINs —
        wrong length, invalid state codes, typos in PAN portion.
        This catches errors before they reach the GST Portal upload.
        Legal basis: Section 25 CGST Act, Rule 10 CGST Rules.

    CALLED BY:
        → core/invoice_validator.py → validate_invoice_row()
        → gui/main_window.py → on_single_gstin_check()

    CALLS:
        → config/constants.py → STATE_CODES, GSTIN_LENGTH
        → utils/format_helpers.py → clean_string()
        → self → verify_checksum() (private helper in this file)

    EDGE CASES HANDLED:
        - None or empty input → returns invalid, message: "GSTIN is empty"
        - Leading/trailing spaces → auto-trimmed before validation
        - Lowercase input → auto-converted to uppercase
        - Old state code "00" → flagged: "use state-specific code"
        - Valid format but wrong check digit → specific error message
        - Non-string input (int, float) → converted to string first

    ASSUMPTIONS:
        - State code list is current as of [date]. New codes require
          updating config/constants.py → STATE_CODES
        - Check digit algorithm follows government specification v1.0.
          If algorithm changes, this function needs manual update.
        *** ASSUMPTION: strict_mode defaults to False because most
            client data has minor formatting issues that should be
            auto-corrected, not rejected. Decided in Session 2. ***

    PARAMETERS:
        gstin_string (str): The GSTIN to validate. May contain spaces,
                            lowercase, or other formatting issues.
        strict_mode (bool): If True, rejects auto-correctable issues
                            instead of fixing them. Default: False.

    RETURNS:
        dict: {
            "valid": bool,          — True if GSTIN passes all checks
            "error": str or None,   — Error message if invalid
            "original": str,        — Input as received
            "cleaned": str,         — Input after auto-correction
            "state_code": str,      — First 2 characters
            "state_name": str,      — Full state name from lookup
            "pan": str,             — Characters 3-12
            "entity_code": str,     — Character 13
            "entity_type": str,     — "Regular"/"Composition"/"Government"/etc.
            "check_digit": str      — Last character
        }
    """
```

For small helper functions (under 5 lines), a shortened format is acceptable:

```python
def clean_string(text):
    """
    WHAT: Strips whitespace and converts to uppercase.
    CALLED BY: core/gstin_validator.py, core/invoice_validator.py
    RETURNS: Cleaned string, or empty string if input is None.
    """
    if text is None:
        return ""
    return str(text).strip().upper()
```

---

## Standard 5: In-Line Code Comments

Comments inside the code follow these rules:

```python
# GOOD — explains WHY (not obvious from the code)
# GST Portal rejects amounts with more than 2 decimal places
amount = round(amount, 2)

# GOOD — cites business rule
# B2CL threshold: Section 37 read with Table 5 of GSTR-1
# Invoices above ₹2,50,000 to unregistered persons → B2CL
if amount > B2CL_THRESHOLD and not recipient_gstin:
    table = "B2CL"

# GOOD — warns about non-obvious behaviour
# WARNING: openpyxl reads merged cells as None for all but the
# top-left cell. We handle this in _unmerge_and_fill() below.

# BAD — describes WHAT (obvious from the code, adds no value)
# increment counter by 1
counter += 1

# BAD — vague
# fix the thing
data = clean_data(data)
```

### Special Comment Tags (searchable markers):

```python
# TODO(Sidhant, DD-MM-YYYY): Add support for GSTR-2B reconciliation
#   Context: Client requested this, deferred to Phase 2.
#   See Parking Lot entry #3 in Rebuild Brief.

# FIXME(DD-MM-YYYY): This breaks if Excel file has merged headers.
#   Workaround: manually unmerge before processing.
#   Proper fix: implement auto-unmerge in readers/excel_reader.py

# HACK(DD-MM-YYYY): Using string comparison for dates because
#   openpyxl returns mixed types (datetime vs string) depending
#   on Excel cell formatting. Proper fix: normalise in reader.

# ASSUMPTION(DD-MM-YYYY): Assuming all input files use UTF-8.
#   If client sends ANSI-encoded files, this will break.
#   Decided in Session 3 — revisit if encoding errors appear.

# LEGAL(Section 16(2) CGST Act): ITC available only if supplier
#   has filed GSTR-1 and tax is reflected in recipient's GSTR-2B.
```

---

## Standard 6: No Magic Numbers

NEVER hardcode values directly in logic. Always use named constants
from `config/constants.py`:

```python
# BAD — magic number, no one knows what 250000 means
if amount > 250000:
    table = "B2CL"

# GOOD — named constant with clear meaning
# In config/constants.py:
B2CL_THRESHOLD = 250000  # ₹2,50,000 — Section 37 / GSTR-1 Table 5

# In core/return_classifier.py:
from config.constants import B2CL_THRESHOLD
if amount > B2CL_THRESHOLD:
    table = "B2CL"
```

This applies to ALL values: tax rates, thresholds, file size limits,
retry counts, timeout values, column names, sheet names, date formats,
state codes — everything. If a value might ever change or needs
explanation, it belongs in config/.

---

## Standard 7: Consistent Error Handling

Every function that can fail follows this pattern:

```python
def read_excel_file(file_path):
    """
    [Standard docstring as per Standard 4]
    """
    # --- Input validation ---
    if not file_path:
        return {"success": False, "error": "No file path provided", "data": None}

    if not os.path.exists(file_path):
        return {"success": False, "error": f"File not found: {file_path}", "data": None}

    if not file_path.endswith(('.xlsx', '.xls')):
        return {"success": False, "error": f"Not an Excel file: {file_path}", "data": None}

    # --- Main logic ---
    try:
        workbook = openpyxl.load_workbook(file_path)
        # ... processing ...
        return {"success": True, "error": None, "data": processed_data}

    except PermissionError:
        # File is open in another program
        return {"success": False,
                "error": f"Cannot read file — it may be open in Excel: {file_path}",
                "data": None}

    except Exception as e:
        # Unexpected error — log details, show friendly message
        return {"success": False,
                "error": f"Unexpected error reading {file_path}: {str(e)}",
                "data": None}
```

Rules:
- NEVER let exceptions crash the application
- ALWAYS return a consistent result format: `{"success": bool, "error": str, "data": ...}`
- ALWAYS validate inputs at the top of the function, before any processing
- ALWAYS catch specific exceptions first (PermissionError, FileNotFoundError),
  then catch general Exception as a safety net
- Error messages must be in PLAIN ENGLISH — the user is not a programmer.
  "File not found" not "FileNotFoundError". "File may be open in Excel" not
  "PermissionError: [Errno 13]"

---

## Standard 8: Type Hints

Every function signature must include type hints. These make the code
self-documenting — even a non-programmer can see what goes in and what
comes out:

```python
# Without type hints — unclear what this expects or returns
def validate_gstin(gstin_string, strict_mode):
    ...

# With type hints — immediately clear
def validate_gstin(gstin_string: str, strict_mode: bool = False) -> dict:
    ...

# More examples:
def read_excel_file(file_path: str) -> dict:
def calculate_tax(amount: float, rate: float, is_interstate: bool) -> dict:
def format_date(date_value: str, output_format: str = "DD-MM-YYYY") -> str:
def get_state_name(state_code: str) -> str | None:
```

---

## Standard 9: Import Organisation

Every file organises imports in this order, with blank lines between groups:

```python
# Group 1: Python standard library
import os
import sys
from datetime import datetime
from pathlib import Path

# Group 2: Third-party libraries
import openpyxl
from PySide6.QtWidgets import QMainWindow, QPushButton

# Group 3: This project's modules
from config.constants import STATE_CODES, B2CL_THRESHOLD
from config.settings import OUTPUT_PATH
from core.gstin_validator import validate_gstin
from utils.format_helpers import clean_string
```

NEVER use `from module import *` — always import specific names.
NEVER use relative imports (`.module`) — always use full paths from
the project root.

---

## Standard 10: Git Commit Messages

Every git commit follows this format:

```
[TYPE] Short description (what changed)

WHY: Why this change was made
AFFECTS: Which files/features are impacted
EDGE CASES: Any new edge cases handled (if applicable)
DECIDED: Reference to decision (if applicable)
```

Types:
- `[FEATURE]` — New feature added
- `[FIX]` — Bug fix
- `[REFACTOR]` — Code restructure without behaviour change
- `[CONFIG]` — Settings, constants, dependency changes
- `[DOCS]` — Documentation only
- `[STYLE]` — Formatting, naming (no logic change)

Example:
```
[FEATURE] Add GSTIN checksum validation

WHY: Format-only validation missed GSTINs with valid structure but wrong
     check digit. 3 client files had this issue in January batch.
AFFECTS: core/gstin_validator.py (new verify_checksum function),
         config/constants.py (added CHECKSUM_CHARS mapping)
EDGE CASES: Handles old-format GSTINs from pre-2019 registrations
DECIDED: Session 2, Message 14 — chose algorithmic check over API lookup
```

---

## Standard 11: Assumptions Must Be Visible

If Claude Code makes ANY assumption — about data format, user behaviour,
business rules, file structure, edge case handling, or anything else —
it MUST appear in THREE places:

1. **In the code** — in the ASSUMPTIONS field of the function docstring,
   or as an `# ASSUMPTION(date):` inline comment
2. **In the DECISIONS_LOG** — in that message's log entry
3. **Verbally in the chat** — "I'm assuming X because Y. Is that correct?"

No silent assumptions. Ever. If the Brief doesn't cover something and
Claude Code has to make a judgment call, that judgment must be visible
and traceable. Silent assumptions become invisible bugs.

---

## Standard 12: Choosing the Language and Libraries

Python is the default for every app: in a Microsoft study, AI agents did better
in Python than in C#; the leading open-source GST tool (India Compliance) is
mostly Python; and Python covers every task below except the few marked. Use
another language only where this table says so, and only after asking me: a
different language changes how the app is started, tested and maintained.

| Task | Use | Watch out for |
|---|---|---|
| Excel/CSV processing and reconciliations (e.g. GSTR-2B vs purchase register), in workbooks that have no Power Query | `openpyxl` (no Excel needed); `rapidfuzz` for fuzzy matching | openpyxl silently drops Power Query queries (tested) and shapes when it saves a workbook; macros survive only with `keep_vba=True` and still cannot be edited. Reading a Power Query workbook with openpyxl is fine: only saving loses the queries |
| Writing into, or refreshing, a workbook that has Power Query queries | Excel itself, driven through COM with `pywin32` | Follow "Excel through COM" below |
| GST portal JSON files (create or read) | Python's built-in `json` | — |
| Text PDFs (invoices, notices) | `pdfplumber` (MIT licence) | Not PyMuPDF: its licence is AGPL (or paid) |
| Scanned PDFs (OCR) | `OCRmyPDF` with Tesseract, or Windows' built-in OCR through `winrt-Windows.Media.Ocr` | Tesseract must be installed on each PC — ask me first. Windows OCR needs the language pack, and Microsoft supports it only for packaged apps |
| Word letters and replies | `python-docx`, with `docxtpl` for templates (no Word needed) | Word-to-PDF (`docx2pdf`) needs Word installed and a logged-in user — never a scheduled job |
| Desktop windows and animations | PySide6 (the `PySide6-Essentials` package, Qt 6) | Pin it to 6.12.x while any PC runs Windows 10 (Qt 6.12 is the last Qt for Windows 10) |
| Tally Prime | XML over HTTP to port 9000 (TallyPrime running, company loaded) | JSON only where every PC runs TallyPrime 7.0 or later; ODBC is read-only |
| Email (reading mail, saving attachments, sending) | Python talking to the mail server itself, so it works whichever mail app people use (new Outlook, Thunderbird or webmail): IMAP to read (`imaplib`), SMTP to send (`smtplib`), or Microsoft Graph for Microsoft 365 mailboxes | Follow "Email" below |
| GST portal downloads | Do not automate: the app works on files a person downloads | Login needs a CAPTCHA and an OTP, and the portal's terms forbid scraping; automated access is only through GSP/ASP APIs |
| Other websites | Playwright for Python, driving the installed Edge (`channel="msedge"`) | Officially supports Windows 11, not Windows 10 |
| Data that several people change | Ask me first | Never an SQLite file in the shared folder used from several PCs at once (it can corrupt); SQLite is fine for one user on one PC |
| Scheduled jobs | Windows Task Scheduler, under the user's own account (no admin needed) | A job that reads the shared folder must "Run only when user is logged on" |
| Summaries and dashboards | An Excel report written with `openpyxl` (charts included), or charts in the app's own window | Nothing that needs a web server kept running. Qt's chart module is not in PySide6-Essentials: it needs PySide6-Addons (about 105 MB more in the shared runtime), so ask me first, or prefer an Excel report |
| Tools that must live inside Excel | VBA — only after asking me | Macros are blocked in files on an untrusted shared folder and in downloaded or emailed files; Python in Excel runs in Microsoft's cloud and cannot reach local files or Tally |
| Small helper scripts | Python | PowerShell `.ps1` scripts are blocked by default on Windows 10/11 |

### Excel through COM

- Use it only for a workbook with Power Query queries that the app writes into
  or refreshes. Before saving with openpyxl into a workbook the app did not
  create, check it for Power Query (a `customXml/item*.xml` part containing
  `DataMashup`). If one is found, use COM, or stop and tell the user. Never let
  openpyxl save over such a workbook.
- Start a separate, hidden Excel and switch off its prompts (tested):
  `excel = win32com.client.DispatchEx("Excel.Application")`, then
  `excel.Visible = False` and `excel.DisplayAlerts = False`. Never take over the
  user's own open Excel. Without `DisplayAlerts = False`, `Quit()` can wait
  forever on an invisible "save changes?" prompt.
- Always call `excel.Quit()` in a `finally` block, even after an error.
- Open the workbook with `IgnoreReadOnlyRecommended=True`. If `workbook.ReadOnly`
  is then true, someone else has it open: close it, and ask the user to close the
  workbook in Excel first.
- To refresh: set `BackgroundQuery = False` on each of the workbook's
  connections, call `workbook.RefreshAll()`, then
  `excel.CalculateUntilAsyncQueriesDone()`, and only then save. Tested on a
  query loaded to a sheet: this saved all 60,000 new rows, while a plain
  `RefreshAll()` followed at once by a save stored the old 2 rows, because the
  refresh was still running. Before saving, check that no connection's
  `Refreshing` is still true. A connection's `RefreshDate` stayed empty, so
  never use it as proof of a refresh.
- Adding a query through COM works (`workbook.Queries.Add`), but loading it
  into a new Excel table failed here (`ListObjects.Add`: "invalid argument",
  in hidden and in visible Excel). If an app needs a new query loaded into a
  table, ask me first.
- In a worker thread, call `pythoncom.CoInitialize()` first (tested: without
  it, COM calls fail), do the Excel work in a separate function, and call
  `pythoncom.CoUninitialize()` after that function returns, so no Excel object
  is still held.
- Excel must be installed on every PC that runs the app, and it needs a
  logged-in user: never a scheduled job.

### Email

- **Microsoft 365:** Microsoft Graph `sendMail` with delegated sign-in (the
  sender signs in once through MSAL; permission `Mail.Send`). Ask me first: it
  needs a one-time app registration in Microsoft Entra, which I do in the admin
  centre. Password sign-in for sending (SMTP AUTH) is switched off by default
  from the end of December 2026. Attachments over 3 MB need a draft plus an
  upload session. Use app-only sending only with `Mail.Send` limited to that one
  mailbox.
- **Other providers (IMAP and SMTP):** which provider, and how it signs in (app
  password or OAuth), is a question for me. After sending over SMTP, save a copy
  in the Sent folder with IMAP `APPEND`, unless the provider does that itself
  (Gmail does); otherwise there is no record of what went to clients.
- **Mail programs:** classic Outlook can be automated through `pywin32`, with the
  same limits as Excel above. The new Outlook for Windows cannot be automated at
  all (no COM, no VBA). Thunderbird can be automated only through a JavaScript
  add-on — ask me first.
- Passwords and sign-in tokens: Standard 13.

---

## Standard 13: How Users Start the App (No Console Window)

The people who use these apps never run code. They double-click a desktop
shortcut, the app's window opens, and no black console window appears. The
launcher and the `main.py` lines below were tested on Windows 11: started
through a Run: Minimized desktop shortcut from a network path, including a
crash, an error in a worker thread, and a missing shared Python.

### Where everything lives

~~~text
\\server\share\Apps\                  ← the office shared apps folder (example name)
├── _runtime\
│   ├── python\                       ← ONE portable Python for every app: python.org's
│   │                                    official "embeddable package" plus the packages
│   └── requirements.txt              ← the Python version, then every package with its exact version
├── GSTR-2B Reconciler\               ← one released app
│   ├── Run GSTR-2B Reconciler.bat    ← what the desktop shortcut runs
│   ├── main.py
│   ├── requirements.txt              ← the packages this app imports, with exact versions
│   └── config\  core\  gui\  ...
└── [the next app]\
~~~

- Nothing is installed on users' PCs. Every app shares the one `_runtime` folder.
- The shared apps folder sits on a server share that OneDrive does not sync:
  OneDrive would upload the runtime's 4,000 files and keep them in sync.
- The shared Python is for PCs in the office. Tested from outside the office,
  over a VPN (Tailscale, 74 ms ping): loading Python and Qt from the shared
  folder took 50–100 seconds, against 1–2 seconds from a local copy. Ask me
  before giving an app to someone who works outside the office.
- All testing used Python 3.14. Move `_runtime` to a newer Python only after
  checking that every package listed has a Windows wheel for it — ask me first.
- Packages go into `_runtime` with
  `pip install --target _runtime\python\Lib\site-packages <package>==<version>`,
  run by a Python of exactly the same X.Y version as the runtime (tested). Then
  import each new package with `_runtime\python\python.exe -s` before calling it ready.
- Each app keeps its own `requirements.txt` with the exact versions of the
  packages it imports (not pytest). When releasing, install any that are missing
  from the shared `_runtime` and add them to `_runtime\requirements.txt`. Never
  upgrade or remove a package already there without asking me: every app uses it.
- Claude builds each app in a local apps folder outside OneDrive (for example
  `C:\Apps-dev\`), laid out the same way: a copy of the shared `_runtime` (plus
  `pytest`) next to the project folders.
- Tests run from the project folder with `..\_runtime\python\python.exe -s -m pytest`.
  The project root has a `pytest.ini` with `[pytest]`, `pythonpath = .` and
  `testpaths = tests`, because the portable Python never adds the project folder
  itself (tested).
- Releasing an app means copying its folder into the shared apps folder without
  `.build-docs`, `tests` and `.git`, then compiling it (see Updating below).

### The launcher: `Run [App Name].bat`, in the app's own folder

~~~bat
@echo off
rem Starts GSTR-2B Reconciler with no console window, from the shared folder or a local copy.
rem Desktop shortcut: point it at this file and set "Run" to "Minimized".
set "PYW=%~dp0..\_runtime\python\pythonw.exe"
if not exist "%PYW%" (
  powershell -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [void][System.Windows.MessageBox]::Show('GSTR-2B Reconciler could not start: the shared Python folder was not found. Please tell the office in-charge.', 'GSTR-2B Reconciler')"
  exit /b 1
)
start "" /D "%~dp0" "%PYW%" -B -s "%~dp0main.py"
~~~

Why each part is there (each one was tested):
- `start "" /D "%~dp0"` with full paths built from `%~dp0` works from a network
  path. Never `pushd` + a relative path + `popd`: from a network path the app
  silently fails to start.
- `pythonw.exe` opens no console. `-B` stops users' PCs writing files into the
  shared folder. `-s` stops packages from a person's own Python folder (in their
  user profile) leaking into the shared Python — without it, a developer PC's own
  packages showed up, so an app could work there and fail on every intern's PC.
- The message box replaces a silent failure when the shared Python is missing.
- Never a `.vbs` launcher (Microsoft is switching VBScript off, around 2027) and
  never a `.ps1` launcher (blocked by default).

**The desktop shortcut:** right-click the .bat → Send to → Desktop (create
shortcut); rename it to the app's name; in its Properties set **Run** to
**Minimized**. Tested: no console window appears, and the app's own window
opens normally (not minimized).

### Rules for the app itself

~~~python
# main.py — right after the file header (Standard 3), before any other code (tested).
import ctypes
import os
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from types import TracebackType

APP_DIR = Path(__file__).resolve().parent
APP_NAME = APP_DIR.name                      # the app's folder is named after the app
LOG_DIR = Path(os.environ["LOCALAPPDATA"]) / APP_NAME / "logs"
MESSAGE_BOX_ERROR_ICON = 0x10                # Windows MB_ICONERROR


def report_fatal_error(exc_type: type[BaseException], exc_value: BaseException,
                       exc_traceback: TracebackType | None) -> None:
    """
    WHAT: Writes the full error to a log file and shows a plain-English message box.
    CALLED BY: Python itself, for any error nothing else caught (sys.excepthook),
               and report_thread_error() for errors in worker threads.
    RETURNS: Nothing.
    """
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOG_DIR / f"error_{datetime.now():%Y-%m-%d_%H%M%S}.txt"
    log_file.write_text("".join(traceback.format_exception(exc_type, exc_value, exc_traceback)), encoding="utf-8")
    ctypes.windll.user32.MessageBoxW(None, f"Something went wrong in {APP_NAME}.\n\nPlease send this file to the office in-charge:\n{log_file}", APP_NAME, MESSAGE_BOX_ERROR_ICON)


def report_thread_error(args: threading.ExceptHookArgs) -> None:
    """
    WHAT: Passes an uncaught error from a worker thread to report_fatal_error().
    CALLED BY: Python itself (threading.excepthook).
    RETURNS: Nothing.
    """
    report_fatal_error(args.exc_type, args.exc_value, args.exc_traceback)


sys.excepthook = report_fatal_error
threading.excepthook = report_thread_error
sys.path.insert(0, str(APP_DIR))             # the shared portable Python does not add the app's folder itself

# Only now import PySide6 and this project's own modules (config, core, gui, ...).
~~~

- These lines come first, straight after the file header (an exception to
  Standards 2, 6 and 9: these few values stay in `main.py` because `config/`
  cannot be imported yet — never move them into `config/`).
- Find the app's own files from `APP_DIR`, never from the current folder.
- There is no console: `print()` shows nothing. Report problems in the window,
  and log them to files in `%LOCALAPPDATA%\[App Name]\`. Worker threads send
  their errors back to the window; with `concurrent.futures`, always call
  `result()`.
- Settings and logs go in the user's own folder (`%LOCALAPPDATA%\[App Name]\`),
  never in the shared app folder.
- Passwords, app passwords, client secrets and tokens never go in code, in
  `config/` or in the shared folder (another exception to Standard 6). A password
  goes in Windows Credential Manager (the `keyring` package, per user); Microsoft
  sign-in tokens go in an encrypted per-user token cache (`msal-extensions`)
  under `%LOCALAPPDATA%\[App Name]\`. Ask me before any app keeps a password.
- Never keep a database, or any file several people write at the same time, in
  the shared folder (Standard 12).

### Updating

- Add an app's new packages to `_runtime` first — only when nobody has an app
  open, because Windows locks `pythonw.exe` and the Qt files while they run —
  and only then copy in the new app version. After a runtime change, open every
  app once.
- App files can be updated at any time; users get the change the next time they
  open the app.
- After every release or update, compile the app's Python files once, from its
  folder: `..\_runtime\python\python.exe -s -m compileall -q .` — start-up stays fast
  even though users' PCs never write into the shared folder. After changing the
  Python version in `_runtime`, re-run this in every app folder.

---
```

---

## Document 4: DECISIONS_LOG.md (auto-maintained by Claude Code)

**File:** `.build-docs/DECISIONS_LOG.md`
**Created by:** You copy the template below. Claude can pre-fill Session 0 from
your discussion; Claude Code fills the rest automatically after every single message.
**Purpose:** This is your project's memory. Every exchange gets logged here in
real time — Session 0 for the project discussion, Session 1 for the Build Plan,
then one session per sub-phase. When a session ends and the chat history is gone,
this file preserves what happened.

**Why this exists:** Human memory is not a hard drive. You will forget why you
chose dropdown over radio buttons. You will forget which edge case prompted that
if-statement. This file remembers so you don't have to.

### Template:

```markdown
# Decisions Log: [PROJECT_NAME]

> **What is this file?** This is your project's memory. After EVERY message
> exchange during the build, Claude Code appends a 2-3 line summary below.
> When a chat session ends and the history is gone, this file preserves
> what happened — what was discussed, what was decided, and what's still open.
>
> **How it is organised:**
> - **Session 0** — what we decided while discussing the project, before the
>   build. Claude pre-fills it from that discussion.
> - **Session 1** — Steps A and B of the Build Prompt: confirming understanding
>   and writing the Build Plan.
> - **Session 2 onwards** — one session per sub-phase of the Build Plan (the
>   Build Loop in CLAUDE.md), laid out as in *Format for each sub-phase session*
>   below.
>
> **Rules:**
> - Auto-maintained by Claude Code — do not edit manually unless correcting an error
> - Claude Code appends after EVERY response, no exceptions
> - Each new sub-phase gets a new session header — not each new chat or usage window
> - The "Open" field carries forward unresolved questions
> - P-flags (technical choices Claude made without asking) are numbered across the
>   whole project — P1, P2, … — so continue from the last number used in this log
> - Lettered assumptions (A, B, C…) are for technical choices only; they stand
>   unless the author objects — silence = consent. Business questions are always asked.

---

## Format for each sub-phase session

Copy this layout when a sub-phase starts. Add a step status line as each step
ends, an exchange entry after every response, and the session entry at .9 Close.

~~~markdown
## Session [N] — Sub-phase [N.M] [name] — [DATE]

### Step status
- [DATE] | [N.M].[K] [step name] | [done / skipped — why] | [one line]

### Exchanges
#### [DATE] | Session [N] | Message [N]
- **Discussed:** [What we talked about in this exchange]
- **Decided/Built:** [What was decided or what code was written]
- **Open:** [Any unresolved question or next step]

### Session entry (written at [N.M].9 Close)
- **Built:** [What now exists that did not before]
- **Decisions:** [What was decided, and why]
- **P-flags:** [From the plan and the coders' saved choices]
- **Assumptions:** [Lettered — A, B, C…]
- **Disputes:** [Each dispute, with Fable's ruling]
- **Break-it tests:** [What Codex's tests found, and what was done]
- **Code review:** [Codex's findings, and the answers]
- **Adversarial review:** [Findings and answers — or "declined"]
- **Superseded:** [Earlier decisions this sub-phase replaced]
- **Open items:** [Carried forward]
- **Status:** [Suite count, e.g. "412 passed, 3 skipped"] · version [x.y]
~~~

---

## Session 0 — Project discussion — [DATE]

> Pre-filled by Claude from our project discussion, before the build starts.

- **Discussed:** [What the tool is for and how it will be used, as we talked it through]
- **Decided:** [Each decision we reached — one line each]
- **Assumptions:** [Technical choices only — A. …  B. … — they stand unless the
  author objects. Anything that changes what the app does with client work goes
  under Open, as a question.]
- **Open:** [Business questions still to answer — asked one at a time before the plan depends on them]

---

## Session 1 — Build Plan (Steps A and B) — [DATE]

### Step status

(One line when Step A is confirmed; one when BUILD_PLAN.md is saved)

### Exchanges

(Entries will be added automatically below)

```

---

# PART 2: THE REBUILD PROMPT

This is what you paste into Claude Code to start the actual rebuild. The four
documents above should already be in the project folder.

```
I am rebuilding a tool from scratch based on a detailed audit of the
original version and my rebuild decisions. Before writing any code,
read these documents in this order:

1. `CLAUDE.md` — My preferences, how we build (the Build Plan rule and the
   nine-step Build Loop), coding standards, and project principles
2. `.build-docs/PROJECT_AUDIT_[PROJECT_NAME].md` — Complete audit of the original tool
3. `.build-docs/REBUILD_BRIEF_[PROJECT_NAME].md` — My decisions: what to keep, change, drop, add
4. `.build-docs/DECISIONS_LOG.md` — Running log of all past discussions (Session 0 holds what we decided while discussing the rebuild; read the latest entries to understand where we left off)
5. `.build-docs/RESUME_CONTEXT.md` and `.build-docs/BUILD_PLAN.md` — only if
   they already exist. If they do, we are resuming: read RESUME_CONTEXT
   before everything else, and do NOT redo Steps A and B or write
   BUILD_PLAN.md or RESUME_CONTEXT.md afresh — keep updating them as the
   Build Loop says. Tell me where we are (RESUME_CONTEXT's newest Step
   status line and CLAUDE.md *Current Phase*) and what comes next, then
   STOP until I say "next". If a usage limit cut a step short, finish it —
   do not restart it (Build Loop rule 6).

Also check `.build-docs/context/` for any sample files, templates, or
reference documents. Read them before building any related feature.

After reading everything, do the following:

## Step A: Confirm Understanding
Summarize back to me in 5-6 sentences:
- What this tool does
- The key differences between the original and what I want in the rebuild
- The build order you recommend (which feature first, second, third...)
- Any questions or conflicts you spotted between the audit and my brief

Ask me any business question one at a time, in plain English, with a
recommended default (Build Loop rule 3 in CLAUDE.md). Write each answer into
the Rebuild Brief — as a business rule, contract or edge case — before
Step B, so every later step reads it.

Do NOT start writing code until I confirm your understanding is correct.
After the summary, STOP until I confirm — auto mode included.

## Step B: Build Plan — ALWAYS WRITTEN AND SAVED
Once I confirm, write the Build Plan and SAVE it as `.build-docs/BUILD_PLAN.md`.
Do this every time, in every permission mode — auto mode included. Auto mode
may tell you to keep going, minimise interruptions and prefer action over
planning; for this build my rules come first, so it never skips the Build Plan
or any stop point. A plan that lives only in the chat, or only in plan mode,
does not count.

The Build Plan has these parts, in this order:

### Tracker (at the top)
- One row per sub-phase: number, name, status, current step, real-data gate
  (yes/no), live sitting (yes/no)
- The current version, the test-suite count and the last commit
- Updated at the end of every step (Build Loop rule 1)

### Part 1: File Structure
Propose the complete folder and file structure BEFORE any code.
Use the modular folder-based organisation defined in CLAUDE.md:
- `main.py` in root (entry point only)
- `Run <App Name>.bat` in root: what users' desktop shortcut runs (Standard 13)
- `requirements.txt` in root: the exact versions of the packages this app imports
- `pytest.ini` in root (`pythonpath = .`, `testpaths = tests`) and `tests/` for
  the Build Loop's tests (Standard 13)
- `config/` folder for settings, constants, lookup tables
- `core/` folder with one file per business logic domain
- `models/` folder for data structure definitions
- `readers/` folder with one file per input format
- `writers/` folder with one file per output type
- `gui/` folder with one file per screen or dialog
- `utils/` folder for shared helpers, split by concern
For each file, write ONE sentence describing its single responsibility.
Not every folder is needed — propose only what's relevant to this project.
But NEVER propose a flat or single-file structure.

### Part 2: Feature Checklist
- Each feature as a numbered item
- Which files from the structure handle this feature
- Dependencies noted (e.g., "F3 requires F1 to be complete first")
- Edge cases from the Rebuild Brief listed under each feature

### Part 3: Phases and Sub-phases
Group the features into phases, and each phase into sub-phases
(Phase 2 → Sub-phases 2.1, 2.2, …). For each sub-phase:
- The features (F-numbers) and files it covers
- The business rules, contracts and edge cases from the Rebuild Brief it owns
- Which earlier sub-phases it depends on
- Its real-data gate, if any: which real files, and what gets re-counted
- Its live sitting, if any: what I will try
- Carry-forwards from earlier sub-phases (added at each .9 Close)
Break Sub-phase 1.1 into its nine steps (.1 to .9), not started. Each later
sub-phase is broken into its nine steps at the .9 Close of the one before it.

### Change Log (at the bottom)
One row for every change to the plan: date, what changed, why.

Also create `.build-docs/RESUME_CONTEXT.md` with four headings: Where we
are · Exact next action · Step status · Hard-won lessons. Set CLAUDE.md
*Current Phase* to "Build Plan — waiting for approval". Commit (Standard 10),
then STOP.

Wait for my approval of the Build Plan before Sub-phase 1.1 starts.

## Step C: Build Each Sub-phase Through the Build Loop
After I approve the Build Plan:
- Run every sub-phase through the nine-step Build Loop in CLAUDE.md, from
  .1 Plan + plan check to .9 Close
- One step per "next": finish the step, update the status lines, make the
  checkpoint commit (guarded so it never runs if the status update failed —
  Build Loop rule 1), report in plain English, then STOP until I say "next"
- If you meet a decision my brief does not cover, follow Build Loop rule 3:
  business questions come to me one at a time, with a recommended default,
  and each answer goes into the Rebuild Brief as at Step A; technical choices
  are made, logged as P-flags or lettered assumptions, and reported. Never
  decide a business question in code.
- **Put each piece of logic in the correct file from the approved structure.
  Do NOT dump new logic into main.py or combine unrelated logic in one file.
  If a new file is needed that wasn't in the approved structure, tell me
  WHY before creating it. Never reorganise or move files without my approval.**
- **Follow EVERY coding standard in CLAUDE.md — file headers, function
  docstrings with CALLED BY/CALLS/EDGE CASES/ASSUMPTIONS, type hints,
  named constants, consistent error handling. No exceptions. No shortcuts.**

## Step D: After Each Step and Each Sub-phase
After every step (Build Loop rule 1):
- One status line each in RESUME_CONTEXT ("Step status"), CLAUDE.md
  *Current Phase*, the Decisions Log session and the BUILD_PLAN tracker
- Then a checkpoint commit following Standard 10 in CLAUDE.md, guarded so
  it never runs if the status update failed
- A plain-English report, then STOP until I say "next"

At every .9 Close, as well as the loop's own updates:
- Update CLAUDE.md with any new conventions or decisions we made
- Add any new edge cases we discovered to the Rebuild Brief's Edge Cases
  section (Section 6); only ideas left for later go to the Parking Lot

## Step E: MANDATORY — Log Every Exchange
**After EVERY response you give me — no exceptions — append a 2-3 line entry
under "Exchanges" in the current session of `.build-docs/DECISIONS_LOG.md`,
using this format:**

#### [DATE] | Session [N] | Message [N]
- **Discussed:** [What we talked about in this exchange]
- **Decided/Built:** [What was decided or what code was written]
- **Open:** [Any unresolved question or next step]

**Do this silently after every single response. Do not ask permission.
Do not skip it because the exchange seemed minor. Do not announce that
you are doing it. Every exchange matters — the minor ones are often the
ones we forget first.**

**When a new sub-phase starts, add a new session** (full layout in the
Decisions Log's "Format for each sub-phase session"):

## Session [N] — Sub-phase [N.M] [name] — [DATE]

## Step F: Testing Guidance
You run the test suite yourself and show me the exact count before you
report "done" (Build Loop rule 9). I never run tests.

Whenever a sub-phase gives me something new to see or run, and at every
live sitting, tell me how to test it:
- Start the app the way users will: double-click its .bat (or a desktop
  shortcut to it), never main.py
- Give me specific steps: "Click this button, select this file, check this output"
- Tell me what the correct result should look like
- Tell me what to try that SHOULD fail (to verify error handling works)
- If I can provide a sample file for testing, tell me what it should contain

I am not a programmer — do not tell me to "run the test suite" or
"check the console." Tell me what to CLICK and what to LOOK AT.

## Step G: Final Documentation
When all features are complete and tested:
- Generate a `README.md` using the README template at `.build-docs/readme_template.md`
- Fill in every section based on what was actually built (not what was planned)
- Include the 3-5 most likely problems a user will encounter
- Include a version history starting at 1.0
- Commit the README as the final git commit with message:
  "[DOCS] Add user-facing README — project complete"
```

---

# PART 3: STRUCTURED PROMPTING STYLE FOR MID-BUILD CHANGES

When something pops up during the build — a new idea, a correction, a change of
mind — use these templates. The structure ensures Claude Code gets complete
context every time.

---

## Template 1: NEW FEATURE REQUEST (mid-build)

```
## New Feature Request

**What I want:**
[Describe in plain English]

**Why I need it:**
[The problem this solves or the scenario that triggered this idea]

**Where it fits:**
[Which existing feature does it relate to? Or is it standalone?]

**Input → Output:**
[What goes in, what should come out]

**Edge cases to handle:**
[Any scenarios you already know about]

**Priority:**
MUST HAVE — block everything until done
NICE TO HAVE — add to Parking Lot for later

**Scope check:**
Is this listed in the Brief's future possibilities? If yes, add to Parking Lot.

**If it goes ahead:**
Add it to the Brief (its rules, edge cases and any new contract) and to
BUILD_PLAN — a new or changed sub-phase, plus a Change Log row — before any
code. It then runs through the Build Loop like any other sub-phase.
```

---

## Template 2: CHANGE EXISTING FEATURE

```
## Change Request

**Which feature:**
[Feature name or F-number from the build plan]

**What it currently does:**
[Brief description of current behaviour]

**What I want instead:**
[Describe the new behaviour]

**Why the change:**
[What triggered this — a real-world scenario, a better idea, user feedback?]

**Impact check:**
Does this change affect any other feature? If unsure, check before implementing.

**Tests and plans:**
First update the Brief — the rule, and any contract the change touches — so
later tests are written from the new rule. If the change belongs to the
sub-phase in progress, rewrite its plan in place and add a BUILD_PLAN Change
Log row; tests that pin the old behaviour go to Fable as disputes, settled
before that sub-phase closes. If its sub-phase has already closed, add a
change sub-phase to BUILD_PLAN (with a Change Log row) and run it through
the Build Loop. Opus never edits Fable's or Codex's tests. End as usual:
status lines, commit, plain-English report, STOP.
```

---

## Template 3: BUG OR UNEXPECTED BEHAVIOUR

```
## Bug Report

**What I did:**
[Step by step — what buttons I clicked, what file I selected, etc.]

**What I expected:**
[What should have happened]

**What actually happened:**
[What went wrong — error message, wrong output, crash, etc.]

**Sample data:**
[If possible, describe or attach the input that caused the problem]

**Reproducible?**
[Does it happen every time, or only sometimes?]

**How Claude handles it:**
Reproduce it first. Then fix it test-first: a regression test that fails on
the old code and passes on the new. If the fix changes what the app does
with client work, ask me first.
```

---

## Template 4: DESIGN DECISION CHECKPOINT

```
## Design Decision

**Decision:** [What was decided]

**Options considered:**
1. [Option A] — [Pro / Con]
2. [Option B] — [Pro / Con]
3. [Option C] — [Pro / Con]

**Chosen:** [Which option and why]

**Trade-off accepted:** [What downside we're living with]

> Add this decision to CLAUDE.md so future sessions remember it.
```

---

## Template 5: PARKING LOT IDEA (not urgent, save for later)

```
## Parking Lot Entry

**Idea:** [Brief description]
**Triggered by:** [What made you think of this during the current work]
**Priority:** HIGH — raise it at the next .9 Close (Template 1) / LOW — revisit at end
**Notes:** [Any additional context]

> Add to the Parking Lot table in REBUILD_BRIEF.
> Do NOT implement now — continue with the current step.
```

---

## Template 6: SESSION HANDOVER (when starting a new chat/session)

> Only for a cold start. Normally the build stays in one conversation:
> "start afresh" means a new usage window, never a new conversation
> (Build Loop rule 5).

```
## Session Continuation

Read .build-docs/RESUME_CONTEXT.md first: its newest Step status line shows
the step we are on ("Where we are" and "Exact next action" are rewritten at
each .9 Close). Then read CLAUDE.md for project context, how we build and
coding standards, this sub-phase's block in .build-docs/BUILD_PLAN.md, and
the newest session in .build-docs/DECISIONS_LOG.md.

**Last session we completed:**
[Sub-phase and step, or what was done]

**Current state:**
[What's working, what's partially done]

**This session I want to:**
[What to build or fix next — usually the step after RESUME_CONTEXT's newest Step status line]

**Any open decisions from last session:**
[Check the "Open" field in the last few DECISIONS_LOG entries]
```

---

## Template 7: I DON'T UNDERSTAND WHAT YOU BUILT

```
## Explain This Differently

I don't follow what you just built or why. Please explain it as if I'm
a CA reviewing a client's books — walk me through what each part does
using a real-world analogy from accounting or GST practice.

Specifically:
- What does this code do in plain English?
- Why was it needed?
- How does it connect to the features I asked for?
- Walk me through one example with real data
```

---

## Template 8: FINISH, DO NOT RESTART (after a usage limit)

```
## Finish, Do Not Restart

A usage limit stopped the last run partway through step [N.M.K].
Finished files are already on disk — do not redo them.

Read .build-docs/RESUME_CONTEXT.md, then check what is already done:
files, test results, commits. Finish only what is missing, reusing cached
agents where you can. Then end the step as usual: status lines, guarded
commit, plain-English report, and STOP until I say "next".
```

---

# THE COMPLETE PIPELINE — VISUAL SUMMARY

```
STAGE 1: ARCHAEOLOGY      STAGE 2: DECISIONS             STAGE 3: REBUILD
┌────────────────┐        ┌─────────────────────┐        ┌────────────────────────┐
│ Run Audit      │        │ Read Audit          │        │ Provide 4 docs         │
│ Prompt (FINAL) │───────→│ Write Rebuild Brief │───────→│ Run the Rebuild Prompt │
│ on old project │        │ (with contracts)    │        │                        │
│                │        │ Write CLAUDE.md     │        │ Step B: BUILD_PLAN.md  │
│ OUTPUT:        │        │ Copy DECISIONS_LOG  │        │   saved every time,    │
│ Audit .md file │        │ (Session 0)         │        │   auto mode too        │
│                │        │ Add context/        │        │ Step C: each sub-phase │
│                │        │ sample files        │        │   runs the Build Loop, │
│                │        │                     │        │   one step per "next"  │
│                │        │ YOUR DECISIONS:     │        │                        │
│                │        │ Keep / Change /     │        │ DECISIONS_LOG grows    │
│                │        │ Drop / Add          │        │ after EVERY message    │
│                │        │                     │        │                        │
│                │        │                     │        │ Code follows EVERY     │
│                │        │                     │        │ coding standard        │
└────────────────┘        └─────────────────────┘        └────────────────────────┘
    You observe               You decide                     You build
                                                             (Claude remembers)

BUILD LOOP — every sub-phase, one step per "next"

 .1 Plan + plan check → .2 Tests first → .3 Code → .4 Disputes
 → .5 Break-it tests → .6 Fixes + code review → .7 Answers
 → (.7b Adversarial review, if you want it) → .8 Settle → .9 Close
```
