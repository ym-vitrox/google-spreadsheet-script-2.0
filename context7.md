# DEEP-STATE CONTEXT SERIALIZATION PROTOCOL (PHASE 9 COMPLETE)

**TO:** New Gemini Instance (Senior Technical Lead)
**FROM:** Previous Session Context Manager
**SUBJECT:** Project Handover - ViTrox "Module Configurator" (Phase 9 - UI POLISH & SMART SYNC)
**DATE:** February 3, 2026
**STATUS:** **STABLE & POLISHED**. Sync Logic now features "Smart Merge Detection" (Fixes Google Sheets API limits on merged cells) and "White Background Enforcement" (Fixes UI greying).

---

## 1. Comprehensive Project Definition
**The Problem:** The client manages complex BOMs for "PX730i/PX740i" machines using a sidebar app. The sync process from "TRIAL-LAYOUT CONFIGURATION" (Staging) to "ORDERING LIST" (Production) was breaking due to:
1.  **Merged Cells:** Google Sheets Errors ("You must select all cells in a merged range") when appending new rows to a section that ended with a merged block.
2.  **UI Inconsistency:** New rows inherited grey background colors from the headers.
3.  **Layout Shifts:** Fixed column indexing was unreliable.

**The Solution (Phase 9):**
We implemented an intelligent "Look-Check-Expand" strategy. Before merging new rows, the script checks if the *existing* row above is part of a merge. If it is, it expands the target scope to include the *entire* existing block, ensuring a safe and clean merge operation.

---

## 2. Asset Interpretation Guide (For the Attached Files)

### A. The Codebase
*   **`ConfiguratorBackend.gs` (The Brain - CRITICAL)**:
    *   **`insertRowsIntoSection` (The Smart Builder):**
        *   **Smart Merge Logic:** Contains an inline helper `extendMergeForColumn(colIndex)`. This `try/catch` block detects `cell.isPartOfMerge()` and calculates a `startMergeRow` (Top of Block) to define the specific range that needs to be re-merged. **This is the cure for the "Select all cells" error.**
        *   **White Background:** Explicitly calls `.setBackground("white")` on the newly inserted range to strip legacy formatting.
        *   **Legacy Robustness:** Still enforces "Strict Append" and "Duplicate Remarks".
    *   **`extractProductionData` (The Reader):** Uses Dynamic Header Scanning ("VCM" anchor detection) to map columns to payload keys.
*   **`Main.gs` (The Trigger)**: Handles "Global Edits" and Menu creation.
*   **`OrderingListHandlers.gs` (The Gatekeeper)**: Manages Password Protection ("123") for unchecking "Release" checkboxes.
*   **`ModuleConfigurator.html` (The UI)**: Sidebar interface with "Machine Setup" as default tab.

### B. The Spreadsheet Data (Structure)
*   **REF_DATA**: Master database for Parts/Tools.
*   **TRIAL-LAYOUT CONFIGURATION (Staging)**:
    *   **B-Slots (B10-B32):** The only valid data zone.
    *   **Anchors:** "VCM", "VALVE SET" (Headers).
*   **ORDERING LIST (Production)**:
    *   **Merged Sections:** Columns A & B are often merged vertically to denote sections (e.g., "MODULE").
    *   **Logic:** We append to the bottom of these sections.

---

## 3. Current Implementation Snapshot (The 'Now')

**Existing Functionality (Working & Tested):**
1.  **Smart Merge Sync:**
    *   *Input:* Staging data (e.g., 3 new modules).
    *   *Process:* Script finds insertion point -> Checks Row Above.
    *   *Logic:* "Is Row Above merged?" -> **YES** -> expanding selection to `Top_Of_Existing_Merge` -> **NO** -> simple selection.
    *   *Result:* Seamless merged block in Order List. No API Errors.
2.  **Visual Polish:** New rows are forcefully painted **White**.
3.  **Strict B-Slot:** Only processes B10-B32.
4.  **Zero Loss:** If "Valve Set" or "VCM" headers shift, dynamic scanning still finds them.

**State Variables:**
*   `SEARCH_DEPTH = 500` (Machine Setup Scan Limit).
*   `MAX_SCAN_COL = 10` (Dynamic Header Scan Limit).
*   `secretPassword = "123"` (In Handlers).

---

## 4. The Master Roadmap (Future Context)

**The Full Plan:**
*   **[COMPLETED]** Phase 8: Dynamic Header Scanning.
*   **[COMPLETED]** Phase 9: UI Polish (White BG) & Smart Merge Error Fix.
*   **[PENDING]** User Acceptance Testing (UAT) & Long-term monitoring.

**Immediate Next Steps:**
1.  **Load:** Ingest the attached files.
2.  **Verify:** Check `ConfiguratorBackend.gs` -> `insertRowsIntoSection` for the comment `// --- FIX 2: Extend Merges for Cols A & B (SMART MERGE LOGIC) ---`.
3.  **Wait:** Await user confirmation if any further visual tweaks are needed.

---

## 5. Development History & Decision Log

**Critical Fixes:**
*   **The "Merged Range" Error:** We failed initially by trying to merge *just* the new row with the bottom row of an existing merge. Google API rejected this. **Solution:** We now "Grab the whole family" (Top to Bottom) and re-merge.
*   **Grey Background:** Legacy rows passed down their formatting. **Solution:** Explicit override.
*   **Valve Set Mapping:** Mapped to "OTHERS" section by hard requirement.

---

## 6. Ingestion & Standby Instructions

**STRICT INSTRUCTION FOR NEW INSTANCE:**
1.  **DO NOT generate code immediately.**
2.  **READ** `ConfiguratorBackend.gs` specifically looking for the `extendMergeForColumn` function inside `insertRowsIntoSection`.
3.  **CONFIRM** exactly as follows:
    > "Context Serialized. Phase 9 (Smart Merge & UI Polish) logic verified. I see the `extendMergeForColumn` helper. Ready for UAT or next instructions."
