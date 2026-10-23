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

## From D2 (/root/merge-scratch/decisions/part2_6.md, 22 items)

### Questions for the coordinator
- Cross-references inside the items use local ids (P2-1, P2-13, P6-3, P6-5 ...). The assembler must renumber these together with the item headings.
- Possible duplicates with Part 1 that must not carry a recommendation on the GST side: (a) P2-2 (hint shown / Variations sheet, A-R12/B-R12, A-F9/B-F9, A-G67/B-G20). If Part 1 has a 'show the problem-file list' item, merge it, or limit it to the technical side: rejected small or non-Excel files, on-screen display. (b) P2-4 (client key '0042', A-R17, A-G25). Any technical item on A-G25 should point here. (c) P2-6 (length box, A-R36/B-R32/B-R138). C3 row 6 judges 'B, for input safety'. If Part 1 keeps an item on A-G26 (text in number boxes), it must not settle the 15-100 range or the state-code cut. (d) P2-13 (Skip refresh, A-R68/A-R69/B-R63). C5b tags it a GST rule, while the MAP candidates call it technical. The status display (A-R145/B-R136) can stay technical in Part 1. (e) P2-14 (rows per client or per Version folder, A-R87/B-R87). I listed it 'when in doubt' per C3's flag, although sections C3 and C5b call the layout technical. Any Part 1 item on consolidated-workbook location or labels should exclude the row choice. Which batch and Version folders Step 4 offers (A-R42/B-R39) is left to Part 1.
- Per-client 'include client name in folders' toggle (A-R40/B-R35/B-R36, A-R137, A-F10/B-F5): C4 section 4.4.4 tags it [GST RULE] (state-code key), but C5a 6.3 and C3 row 10 say it is technical. I left it out of Part 2 because the key form is identical in both; the difference is only whether a screen control exists, which makes it effectively only-A (Part 3). If the coordinator follows C4's tag ('when in doubt'), the Part 3 item must say 'No recommendation - your call'.
- P2-9 (template file types, A-R52/B-R48/B-R51) is tagged a GST rule only through B-R51, the report name. If the coordinator decides it is not a GST rule, it can move to Part 1 with a recommendation.
- P2-10 (merged Links cell on the openpyxl path) depends on the Part 8 decision on whether any openpyxl report path survives. If Part 8 drops it, P2-10 can be removed or marked 'falls away'.
- Dormant GST rules are rows in Part 6: critical types A-R30/B-R28 (P6-2 row 5) and calculated columns A-R85/B-R84 (P6-6 row 12), each asking 'drop' or 'keep'. Part 4 must not recommend dropping them; it may point to these rows instead.
- Shared rules I did not place, left for Parts 4 and 8 and the shared-gap items, so the coverage table should check they land somewhere. (a) Excel handling: A-R97, A-R100 to A-R103, A-R106 to A-R108 / B-R96, B-R98, B-R99, B-R101, B-R104 to B-R110. (b) Logs: A-R120, A-R122 / B-R121, B-R123. (c) Modes and backup behaviour: A-R45/B-R43. (d) 'Created if the file exists': A-R99/B-R59. (e) Progress tracker: A-R141/B-R113. (f) Shared dead code: A-R13 to A-R15, A-R18, A-R34, A-R37, A-R49, A-R50, A-R53, A-R63, A-R70, A-R84, A-R92 to A-R95, A-R107, A-R150, A-R151, A-R153 / B-R14 to B-R16, B-R18, B-R33, B-R41, B-R46, B-R47, B-R49, B-R58, B-R64, B-R72, B-R85, B-R91 to B-R93, B-R106 to B-R109, B-R130, B-R131.
- P2-5 cites A-R138/B-R134 for the Missing column only, and P2-11 cites B-R134 for the order only. The rest of those rules (default ticks, sortable columns, Last Run) is technical and belongs in Part 1.

### Checker: unresolved
- C4 section 4.4.4 tags the per-client "client name in folder names" option as [GST RULE], because the flag is stored under "<client>-<state code>". It sits in Part 1 as P1-9, with a recommendation, while C5a 6.3 and C6 call it not a GST rule. I did not add a duplicate Part 2 item. The assembler or Part 1 checker must decide: either keep P1-9 as it is, or move the state-code key part into Part 2 and drop its recommendation.
- C910's candidate list groups the "Private"/"Limited" shortening under "Client identity ... GST RULE", while C5a 2.2 and MAP call it technical. It sits in Part 1 as P1-10, with a recommendation. P6-4 now points to P1-10. The assembler should confirm that classification.
- Several shared, non-GST rules from section 5 are not in this file and need a home in another part (Part 8 or the closing list): A-R45/B-R43 (what each mode does with an existing file), A-R99/B-R59 (a report counts as made if the file exists), A-R100 to A-R102/B-R98, B-R99 (open retries and long-path handling), A-R106/B-R97 (keep_links), A-R108/B-R104 (all warnings switched off). Also shared dead rules (A-R13 to A-R15, A-R18, A-R34, A-R37, A-R49, A-R50, A-R53, A-R63 and their B partners), presumably for Part 4. I could not check the other part files because only part1.md and part11.md exist so far.


## From D3 (/root/merge-scratch/decisions/part3_4_5.md, 39 items)

### Questions for the coordinator
- P3-6 (A's per-client 'include client name in folder names' switch: A-F10, A-R137, A-R40 / B-F5, B-R36) has a recommendation (Keep). MAP, C3 and C5a treat it as technical, but the C4 decision candidates tagged it [GST RULE], because the folder text also reaches the Links cells. Under the 'when in doubt' rule, should the recommendation be removed?
- Some items are likely to duplicate Part 1 items and should be merged by the assembler: P3-2 start-up check (A-F1/B-F1), P3-9 Variations list (A-F9/B-F9; its hint wording is MAP 2.1 item 2 in Part 2), P3-13/P3-14/P3-15 Step 4 history, log file and Stop (C3 row 25), P3-16 Step 4 fill at start-up (A-R144/B-R126), P3-17 scan log (B-F7 vs A-F8), and P3-20 plus P5-1 theme and dark mode (A-F28/B-F28).
- P4-8 (old documents, README, .claude/settings.local.json) may overlap Part 10 (housekeeping and old public history). Cross-reference or merge.
- P3-1 recommends turning A's test scenarios into a context note. The Part 9 drafter should offer that note so the two items agree.
- A-R117/B-R117 (recent folders and session state) appear in P4-5/P4-6 (drop the half-built code) and in P5-2 (defer the feature). The two are consistent; keep both or merge.

### Checker: unresolved
- The GST status of the per-client 'client name in folder names' setting is inconsistent. MERGE_COMPARISON.md 4.4.4 and C4's candidate tag it [GST RULE]; section 5, 6.3 and both registries say it is not. P1-9 (in part1.md, not my file) gives it a recommendation. The Part 1 checker or the assembler needs to settle it (when in doubt it is a GST rule).
- part7_8_9_10.md P8-8 still refers to Part 4 ('The separate launchers'). P4-7 now covers only extractor_main.py, and the launcher question is P1-22. The assembler should re-point P8-8.
- The pointer list asks the assembler to fold minor points into P1-18 (Stop checked inside every wait; a stopped run reported as 'stopped') and P1-20 (A's tooltips are not kept on screen near its edges). I cannot edit part1.md.
- part1.md was being edited at the same time as this check. I kept P3-3 (install offer), P3-8 (export) and P3-16 (Step 4 start-up fill) because P1-1, P1-8 and P1-13 now send those questions to Part 3. If part1.md changes again, the 'Asked in another part' list must be checked again.
- P3-1's recommendation and P9-7 note (4) both rely on the same reference note of A's test scenarios. They agree, but the assembler should keep them pointing at each other.
