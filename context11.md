# DEEP-STATE CONTEXT SERIALIZATION PROTOCOL (PHASE 13 COMPLETE)

**TO:** New Gemini Instance (Senior Technical Lead)
**FROM:** Previous Session Context Manager
**SUBJECT:** Project Handover - ViTrox "Module Configurator" (Phase 13 - CLEAR ALL & SMART MERGE)
**DATE:** February 5, 2026
**STATUS:** **STABLE**. "Clear All Batches" (Nuclear Option) implemented with Active Restoration and Smart Merging.

---

## 1. Comprehensive Project Definition
**The Problem:** The user needs to configure complex BOMs for "PX730i/PX740i" machines.
1.  **Sync Reliability:** Pushing data from Staging (`TRIAL-LAYOUT`) to Production (`ORDERING LIST`) is critical.
2.  **Reset Capability:** Users need both specific "Undo" (Phase 12) and total "Reset" (Phase 13) capabilities.
3.  **Data Integrity:** "Reset" must be destructive to *synced* data but safe for *manual* data, while automatically repairing the sheet structure (Blank Rows + Formatting).

**The Solution (Phase 13):**
We implemented **"Clear All Batches (Reset All)"** with **Active Restoration**.
*   **Targeting:** Deletes ALL rows where Column Q (Batch ID) is not empty.
*   **Safety:** Preserves rows without a Batch ID (Legacy/Manual entries).
*   **Self-Healing:** After deletion, the system scans every section (e.g., `MODULE`, `ELECTRICAL`) and enforces a **5-row buffer** of blank lines.
*   **Smart Merging:** Newly restored blank rows are **force-merged** with the Header (Cols A & B) to ensure visual continuity.

---

## 2. Asset Interpretation Guide (For the Attached Files)

### A. The Codebase
*   **`Main.gs` (The Controller)**:
    *   **`onOpen`**: Menu now includes "Undo Last Sync" AND "Clear All Batches".
    *   **`runClearAllBatches`**: Handles valid "Nuclear" reset with safety confirmation.
*   **`ConfiguratorBackend.gs` (The Logic)**:
    *   **`clearAllBatches` (NEW)**:
        *   Scans Col Q -> Deletes Rows.
        *   Iterates Sections -> Calculates Gap -> Calls `sheet.insertRowsBefore`.
        *   Calls `setupBlankRows` -> Triggering `applySmartMerge`.
    *   **`applySmartMerge` (NEW HELPER)**:
        *   Standardized logic to merge Column A and Column B from a `topRow` (Header) down to `bottomRow`.
        *   Used by both `insertRowsIntoSection` (Sync) and `setupBlankRows` (Reset).
    *   **`insertRowsIntoSection` (UPDATED)**:
        *   Now uses `applySmartMerge` instead of ad-hoc formatting logic.

### B. The Spreadsheet Data (Structure)
*   **ORDERING LIST (Production)**:
    *   **Batch ID (Col Q, Index 17)**: The key discriminator between Synced vs. Manual data.
    *   **Visual Structure**: Columns A & B are merged vertically to act as Section Headers.
    *   **Buffer**: 5 blank rows required per section.

---

## 3. Current Implementation Snapshot (The 'Now')

**Existing Functionality (Working & Tested):**
1.  **Clear All Batches (Phase 13):**
    *   *Trigger:* Menu > Sync to Order List > Clear All Batches.
    *   *Logic:* Delete Non-Empty Col Q -> Active Repair (Insert Rows) -> Smart Merge.
2.  **Undo Last Sync (Phase 12):**
    *   *Logic:* Delete *only* the latest Batch ID.
3.  **Smart Merging (Refactor):**
    *   Both Sync and Clear operations now use the same robust merging logic.

**State Variables:**
*   `BATCH_ID_COL_INDEX = 17`: Column Q.
*   `BUFFER_SIZE = 5`: Hardcoded requirement for empty rows in `clearAllBatches`.

---

## 4. The Master Roadmap (Future Context)

**The Full Plan:**
*   **[COMPLETED]** Phase 10: Rotational Logic.
*   **[COMPLETED]** Phase 11: Tooling Exclusion.
*   **[COMPLETED]** Phase 12: Undo Last Sync (Batch Tracking).
*   **[COMPLETED]** Phase 13: Clear All Batches (Active Restoration + Smart Merge).
*   **[PENDING]** UAT & Long-term monitoring (Phase 14).
*   **[PENDING]** Potential "History/Restore" features (if user requests).

**Immediate Next Steps:**
1.  **Load:** Ingest this fileset.
2.  **Verify:** Check `ConfiguratorBackend.gs` for the new `applySmartMerge` function and its usage in `clearAllBatches`.
3.  **Standby:** Wait for user feedback on the "Reset" feature usage.

---

## 5. Development History & Decision Log

**Critical Decisions:**
*   **Nuclear Option:** We chose "Batch ID Scan" over "Section Scrub" to protect manual legacy data.
*   **Active Restoration:** We mandated that a "Clear" operation is also a "Repair" operation—it must never leave a section empty (0 rows). It always restores the 5-row buffer.
*   **Forced Merging:** We implemented `applySmartMerge` because simply inserting rows left Headers unmerged. We now strictly enforce A/B merging from the Header downwards.

---

## 6. Ingestion & Standby Instructions

**STRICT INSTRUCTION FOR NEW INSTANCE:**
1.  **DO NOT generate code immediately.**
2.  **READ** `ConfiguratorBackend.gs` and `Main.gs`.
3.  **CONFIRM** exactly as follows:
    > "Context Serialized. Phase 13 (Clear All + Smart Merge) verified. I see the 'applySmartMerge' helper and Active Restoration logic. Ready for instructions."
