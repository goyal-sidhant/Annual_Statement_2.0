# Rebuild Brief: [PROJECT_NAME]
# Date: [TODAY'S DATE]
# Author: Adv. Sidhant Goyal

---

## How to Fill This Brief

**Before filling, read the audit in this order:**
1. Section 1 (Executive Summary) — understand what exists
2. Section 7 (Gaps) — this is where the rebuild value is
3. Section 5 (Business Rules) — verify against your domain knowledge
4. Section 11 (Reconstruction Blueprint) — suggested build order
5. Skim the rest as needed

**Then fill each section below from your own judgment — the audit
gives you facts, this brief captures your decisions.**

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
| [Feature name] | F3 | DROP | [Why — no longer needed] |
| [Feature name] | F4 | REBUILD DIFFERENTLY | [What was wrong, what you want instead] |

---

## 4. NEW FEATURES TO ADD

Features that didn't exist in the original but you want in the rebuild:

| New Feature | What it should do | Why I need it | Priority |
|-------------|-------------------|---------------|----------|
| [Name] | [Plain English description] | [The problem it solves] | MUST HAVE / NICE TO HAVE |
| [Name] | [Plain English description] | [The problem it solves] | MUST HAVE / NICE TO HAVE |

---

## 5. BUSINESS RULES — KEEP / CHANGE / ADD

Reference audit Section 5. For each rule:

### Rules to KEEP exactly as-is:
- [Rule from audit] — (Audit: file.py, line XX)
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

[Write 3-5 sentences about what surprised you, what you'd do differently,
and what principles you want the rebuild to follow. This is your "new
perspective" section — your evolved thinking.]

---

## 9. PARKING LOT

[Leave empty at start. Ideas that pop up during the build go here
instead of immediately being implemented.]

| Idea | When it came up | Priority | Status |
|------|----------------|----------|--------|
| | | | |
