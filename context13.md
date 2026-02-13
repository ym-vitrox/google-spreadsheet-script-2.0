# DEEP-STATE CONTEXT SERIALIZATION PROTOCOL (PHASE 13 REFINED - TARGETED RESET & REF_DATA HEADERS)

**TO:** New Gemini Instance (Senior Technical Lead)
**FROM:** Previous Session Context Manager
**SUBJECT:** Project Handover - ViTrox "Module Configurator" (Phase 13 - REFINED RESET & REF_DATA)
**DATE:** February 9, 2026
**STATUS:** **STABLE**. "Clear All Batches" now uses a targeted approach to preserve manual entries. `REF_DATA` sync now supports a header row.

---

## 1. Comprehensive Project Definition
**The Problem:** The user needs to configure complex BOMs for "PX730i/PX740i" machines.
1.  **Sync Reliability:** Pushing data from Staging (`TRIAL-LAYOUT`) to Production (`ORDERING LIST`) must be traceable.
2.  **Reset Capability:** "Undo" and "Clear All" features must effectively "un-sync" data, allowing re-syncing without duplication or errors.
3.  **Data Integrity:** We must differentiate between machine-synced rows and user-manually-entered rows in Staging.
4.  **Reference Data:** The source `REF_DATA` sheet now has a header row that must be preserved during syncs.

**The Solution (Phase 13 Refined):**
We implemented **Refined Batch Logic & Header Support**.
*   **Targeted Reset (Clear All Batches):** Instead of wiping entire columns, the system scans for `Batch ID`s. If a row has a Batch ID, it clears both the ID and the "SYNCED" status. If no ID exists, the row is treated as manual and preserved.
*   **REF_DATA Header Support:** All sync and read operations now strictly start from Row 2, preserving the header in Row 1.
*   **Dual-Write Storage (Unchanged):**
    *   **Production:** `ORDERING LIST` stores Batch ID in **Column Q (Index 17)**.
    *   **Staging:** `TRIAL-LAYOUT` stores Batch ID in **Column S (Index 19)** (and marks "SYNCED" in Col J).

---

## 2. Asset Interpretation Guide (For the Attached Files)

### A. The Codebase
*   **`Main.gs` (The Controller)**:
    *   **`runMasterSync`**: Updated to preserve Row 1 of `REF_DATA`.
    *   **`runProductionSync`**: Generates `batchID` centrally.
    *   **`runClearAllBatches`**: Triggers the refined reset logic.
    *   **New Utility**: `insertShoppingList` now references `I2:I` and `K2:K` for validation to exclude headers.
*   **`ConfiguratorBackend.gs` (The Logic)**:
    *   **`resetAllSyncStatus`**: **CRITICAL**. This function now iterates through Column S. It *only* clears Col S and Col J if a Batch ID is found.
    *   **`clearAllBatches`**: Calls `resetAllSyncStatus` after clearing Production rows and restoring the buffer.
    *   **REF_DATA Getters**: Functions like `getModuleList`, `getVisionPCOptions` now start reading from Row 2.

### B. The Spreadsheet Data (Structure)
*   **ORDERING LIST (Production)**:
    *   **Batch ID (Col Q, Index 17)**: The primary key for "Undo".
    *   **Visual Structure**: 5 blank rows buffer required per section.
*   **TRIAL-LAYOUT CONFIGURATION (Staging)**:
    *   **Sync Status (Col J, Index 10)**: "SYNCED" marker.
    *   **Batch ID (Col S, Index 19)**: Stores the ID to allow targeted un-syncing.
*   **REF_DATA (Reference)**:
    *   **Header Row (Row 1)**: **DO NOT TOUCH**.
    *   **Data (Row 2+)**: All valid data starts here.

---

## 3. Current Implementation Snapshot (The 'Now')

**Existing Functionality (Working & Tested):**
1.  **Sync REF_DATA (DB Only):**
    *   Fetches data from external source.
    *   Clears `REF_DATA` from **Row 2** downwards.
    *   Writes new data starting at **Row 2**.
2.  **Clear All Batches (Targeted):**
    *   Deletes all rows in Production that have a Batch ID in Col Q.
    *   Restores the 5-row buffer in Production.
    *   Scans Staging Col S. **IF** Batch ID exists -> Clear Col S & Col J. **ELSE** -> Skip.
3.  **Undo Last Sync:**
    *   Removes latest batch from Production.
    *   Removes sync status for that specific batch in Staging.

**State Variables:**
*   `BATCH_ID_COL_INDEX = 17` (Ordering List).
*   `STAGING_BATCH_COL = 19` (Trial Layout - Implicitly handled in backend).

---

## 4. The Master Roadmap (Future Context)

**The Full Plan:**
*   **[COMPLETED]** Phase 12: Undo Last Sync (Basic).
*   **[COMPLETED]** Phase 13: Robust Sync & Reset (Targeted Logic).
*   **[COMPLETED]** REF_DATA Header Support.
*   **[PENDING]** User Acceptance Testing (UAT).
*   **[PENDING]** "History" Mode (Optional future discussion).

**Immediate Next Steps:**
1.  **Load:** Ingest this file set.
2.  **Verify:** Check `ConfiguratorBackend.gs` -> `resetAllSyncStatus`. Ensure it has the `if (batchID !== "")` check.
3.  **Wait:** User will likely test the "Clear All Batches" to confirm manual entries are safe.

---

## 5. Development History & Decision Log

**Critical Decisions:**
*   **Targeted vs. Bulk Clear:** We switched from bulk clearing Columns J & S to a targeted approach. This was to protect "Manual Entries" where a user might manually type "SYNCED" or have legacy data without a Batch ID.
*   **Header Preservation:** We explicitly shifted all `REF_DATA` operations to `Row 2` to allow for a permanent header row, improving readability for the user.

---

## 6. Ingestion & Standby Instructions

**STRICT INSTRUCTION FOR NEW INSTANCE:**
1.  **DO NOT generate code immediately.**
2.  **READ** `ConfiguratorBackend.gs` and check `resetAllSyncStatus`. verify it loops and checks for `batchID`.
3.  **CONFIRM** exactly as follows:
    > "Context Serialized. 'Targeted Reset' logic confirmed in ConfiguratorBackend.gs. REF_DATA Header support confirmed. Ready for instructions."
