# Coordinator notes for the assembler (Phase 3)

Questions and hand-overs the part drafters raised. Resolve each one when assembling: decide placement (GST-rule items go to Part 2 with no recommendation - when in doubt, it is a GST rule), merge duplicates, and record what you did in "## Checks before posting".

## From D1 (/root/merge-scratch/decisions/part1.md, 22 items)

### Questions for the coordinator
- Borderline GST classification that I kept in Part 1 with a recommendation: P1-10, Private/Limited name cleaning (A-R19/A-R20 against B-R19/B-R20). C3, C4 and C5a all say it is not a GST rule, but C910's decision candidates fold 'Private/Pvt' into a GST-tagged 'client identity' line. If Part 2 also covers it, one of the two must go.
- Borderline: P1-9, the per-client 'client name in folder names' toggle (A-F10/B-F5, A-R40/B-R35, B-R36). C3 and C5a say it is not a GST rule. C4 4.4.4 tags it GST only for its state-code key, which is the same reader in both. I kept it in Part 1 and said the key form is not a choice there.
- Borderline: P1-14, Step 4's choice of batch and Version folders (A-R42/B-R39). No registry tags it, and C3/C5a call it technical. It interlocks with the consolidated rows question (one row per client or per version, A-R87/B-R87), which C3 flags as a possible GST rule. I left that question for Part 2 and added a consistency note to P1-14.
- Items I sent to Part 2 that are not among MAP 2.1's eleven GST pairs, so the Part 2 drafter may not have them: Skip refresh and refreshed-copy recognition (A-R68, A-R69/B-R63; C5b 9.4 'treated as one'); template file types offered (A-R52/B-R48, B-R51; C5a 8.2); lowest client-folder length (A-R36/B-R32, B-R138; C5a 6.1, C4 4.4.3); the Step 2 'Missing' column names-versus-count and the 'Folder Name' column (C6 G3, G4); consolidated rows per client or per version (A-R87, A-R91/B-R87, B-R90). Please confirm Part 2 covers them, or move any back to Part 1.
- P1-1 (start-up), P1-4 (settings), P1-17 (run log) and P1-19 (unexpected errors) touch areas Part 8 also covers through Standards 12 and 13: launcher, settings/log location, error hooks and Power Query protection. I limited them to user-visible behaviour and pointed to Part 8. The assembler may want to merge or cross-reference.
- B-F20 (BB2 check), A-F23/B-F22 (cell extraction) and A-F26/B-F26 (dialog clicking) are the same in both but defective, so they are not in my closing list. They are listed under 'Pointers to other parts' for Parts 6 and 7.

### Checker: unresolved
- CONFLICT for the assembler: P1-9 (per-client client-name toggle) recommends 'the tick for all clients only', while part3_4_5.md P3-6 (same question) recommends 'Keep'. P3-7 (long-name warning) overlaps too. Merge them into one item with one recommendation. Note that comparison C4 row 4A-12 tags this question '[GST RULE] (state-code key)', whereas the registries (A-R40 'no') and C5a 6.3 say it is not a GST rule.
- CONFLICT: P1-21 (dark mode) recommends keeping it now, built on B's single colour sheet, while P5-1 recommends 'Defer' and P3-20 asks the theme-file question. Merge these into one answer.
- CONFLICT: P1-22 recommends B's way (keep 'open on Step 4' as a second start), while P4-7 recommends dropping all three start files and using only the Standard 13 launcher. Merge these into one answer.
- P1-13 no longer conflicts with P3-16: the start-up fill question now sits in Part 3. P3-16 could also cite A's deliberate reason from bfcb24a ('scanning there would hit the network on every launch').
- Duplicates with consistent recommendations, which the assembler should merge or cross-reference: P1-1 with P3-2/P3-3 (start-up check, install offer); P1-17 with P3-14 (Step 4 run log); P1-18 (Step 4 Stop) with P3-15; P1-6 with P3-12/P3-17 (scan speed and live progress); P1-11 with P3-12 (run folder worked out once per run, A-R111); P1-2 (tab dots) with P3-18; P1-8 (export) with P3-8; P1-7 with P3-9 (showing rejected files); P1-20 (hover tips) with P3-10; the closing-list line A-R112/B-R114 with P6-8.
- Not verifiable from the code: whether the office used A's per-client toggle (P1-9's 'little used' is a LOW inference); whether Windows updates folder dates the way the Linux runs showed (P1-14, MEDIUM); whether a %TEMP%-built report records that folder inside the workbook (12d505f asks for a Windows check, P1-11).

## From D5 (/root/merge-scratch/decisions/part11.md, 40 items)

### Questions for the coordinator
- Possible overlaps for the assembler to merge: P11-35 (is the company-like name in A's PROJECT_AUDIT_ANNUALSTATEMENT.md a real client; is the share path confidential) may duplicate a Part 10 item on the old public history. P11-13 (shared apps folder path) may overlap Part 8 or Part 10. P11-27 and P11-28 (facts about the templates, including the 'how to check for hard-coded folder names' step) may overlap Part 9 (templates) or the Part 1 item on category folder names. P11-9 (Windows 10 or 11) may overlap Part 8's PySide6 pin. P11-19 (fact side of H1, the IMS/Extra columns) pairs with the GST-rule item in Part 2 or 6. Suggestion: keep Part 11's fact questions and have the other parts point to them.
- These Appendix X6 questions are decision questions that I left to Parts 1-10. Please check they land somewhere. A X6: Q2, Q6, Q8, Q10, Q12, Q13, Q14, Q15, Q30, Q34, Q35, Q36, Q37, Q38, Q40, Q49, Q50, Q51, Q52, Q53, Q54, Q55, Q56, Q63, Q64, Q66, Q70, Q71, Q72, Q73, Q74, Q79, Q80, Q82, Q83, Q87. B X6: Q4, Q5, Q8, Q10, Q11, Q12, Q14, Q15, Q21, Q23, Q30, Q31, Q33, Q40, Q41, Q43, Q44, Q45, Q46, Q49, Q52, Q53, Q55, Q56, Q62, Q64, Q65.
- P11-13 tells the author that the shared-apps-folder path will be written into CLAUDE.md in NEW, so Part 10's private/public answer matters. If Part 10 recommends keeping NEW public, that wording may need a cross-reference.

### Checker: unresolved
- Part 11 items carry several lettered sub-questions. Each sub-question can be answered in a word or a line, and the intro now says so, but no item was split into separate items. The assembler may prefer one question per item.
- Cross-references between items use local ids (P11-9, P11-10, 'the next item', 'the previous item', P2-11 by topic). The assembler must remap them when it renumbers.
- The run figures (2,508 / 3,030 / 1,094 files, 270 characters, 41 s) are kept. They are office operational detail, not client data, per C910 D5. P11-35(b) now asks the author whether he regards them as confidential; if he says yes, they must be removed from P11-2, P11-12, P11-14, P11-15 and P11-21.
- Parts 3-5 and 7-10 do not exist yet, so I could not check whether the Appendix X6 questions not in Part 11 (all decision-type) are covered elsewhere. Part 11's own fact-question coverage was checked against all X6 lists and the DECISION_CANDIDATES question lists.

## From D4 (/root/merge-scratch/decisions/part7_8_9_10.md, 43 items)

### Questions for the coordinator
- Duplicates to merge when assembling. One answer each is enough: P10-7 (is the March-audit name a real client) = P11-35(a); P8-2 (Windows 10/11) = P11-9(a); P8-11 (testing) and P3-1 (A's tests) both recommend a reference note of A's test scenarios. Partial overlaps, where the contract adds a test to a Part 1 choice and the items point to each other by title: P7-14 C14 with P1-18; P7-15 C15 with P1-19; P7-12 C12 with P1-11; P7-2 C2 with P1-7 and P3-9; P8-9 (where settings live) with P1-4 (what is remembered).
- Conflict between other drafters: P1-22 recommends B's way (one window opened on Step 4), while P4-7 recommends dropping all three launcher files. My P8-8 recommends Standard 13's single launcher and leaves this question to those two items.
- C8 (P7-8, consolidated columns follow the approved cell list) is given 'No recommendation - your call', because P6-6 row 11 treats the 'contains' look-up as a GST rule. Please confirm this is the treatment you want. If the assembler instead judges C8 technical, it would become 'Agree'.
- How Step 4 reads the figures (openpyxl or a second Excel) is deliberately left to P5-9 (defer). P8-5 points there rather than recommending openpyxl now.
- P10-5 proposes a .gitignore, which is not in Phase 5's file list. It is created only if the author says yes.

### Checker: unresolved
- P11-35 (part11.md, which I may not edit) now carries the merged question about the company-like name; its drafter or the assembler may want to add the exact line list (130, 133, 134, 158, 160-162, 177, 179, 180, 182, 227, 228, 231, 232) and '33 of A's commits from f7eb8f1' that I kept in the Part 10 note.
- part1.md P1-19 says B 'ends with error code 134' with no qualifier. That figure is the Linux exit code for an abort (B audit line 976), not what Windows records; the Part 1 checker should qualify it.
- Some shared gaps land in no draft at all: A-G58 (L6, text starting with '=' written as a formula; B-G65 appears only in a part1.md refs list), B-G75 (L1), B-G42 (L14) and B-G76 (L15). They are LOW and outside Parts 7-10, but the assembler's coverage check should place them or list them.
- The 'Std:NNN', 'Section N' and other shorthands rely on this file's glossary table at the top. The assembler must keep that glossary, or move its rows to the document-wide glossary, or many terms lose their first-use explanation.
- C13 (path-length check) keeps a recommendation as a technical promise, while P6-8 row 5 asks the author to confirm the 218-character value as a Part 6 '⚠' rule. This is consistent as written, but the assembler should keep both cross-references.
