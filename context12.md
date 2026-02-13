# DEEP-STATE CONTEXT SERIALIZATION PROTOCOL (PHASE 13 COMPLETE - ROBUST SYNC)

**TO:** New Gemini Instance (Senior Technical Lead)
**FROM:** Previous Session Context Manager
**SUBJECT:** Project Handover - ViTrox "Module Configurator" (Phase 13 - ROBUST SYNC & RESET)
**DATE:** February 5, 2026
**STATUS:** **STABLE**. "Clear All Batches" and "Undo Last Sync" now feature robust "Sync Status Reset" and "Batch ID Tracking" across Production and Staging.

---

## 1. Comprehensive Project Definition
**The Problem:** The user needs to configure complex BOMs for "PX730i/PX740i" machines.
1.  **Sync Reliability:** Pushing data from Staging (`TRIAL-LAYOUT`) to Production (`ORDERING LIST`) must be traceable.
2.  **Reset Capability:** "Undo" and "Clear All" features must effectively "un-sync" data, allowing re-syncing without duplication or errors.
3.  **Data Integrity:** We must track exactly *which* rows belong to *which* sync operation to handle partial undo operations safely.

**The Solution (Phase 13):**
We implemented **Robust Batch ID Tracking**.
*   **Central ID Generation:** A unique `Batch ID` (Timestamp) is generated at the start of every Sync.
*   **Dual-Write Storage:**
    *   **Production:** `ORDERING LIST` stores Batch ID in **Column Q (Index 17)**.
    *   **Staging:** `TRIAL-LAYOUT` stores Batch ID in **Column S (Index 19)** (and marks "SYNCED" in Col J).
*   **Active Reset:**
    *   **Undo Last Sync:** Deletes Production rows -> Scans Staging Col S -> Clears Staging status for *only* that batch.
    *   **Clear All Batches:** Deletes Production Synced rows -> Restores Buffer -> Clears *ALL* Staging status (Cols J & S).

---

## 2. Asset Interpretation Guide (For the Attached Files)

### A. The Codebase
*   **`Main.gs` (The Controller)**:
    *   **`runProductionSync`**: Now generates `batchID` centrally and passes it to *two* backend functions: `inject` and `mark`.
    *   **`runClearAllBatches`**: Triggers the "Nuclear" reset.
*   **`ConfiguratorBackend.gs` (The Logic)**:
    *   **`injectProductionData(payload, batchID)`**: Receives and writes the ID to Col Q.
    *   **`markRowsAsSynced(rows, batchID)`**: Receives and writes the ID to Staging Col S.
    *   **`clearLatestBatch`**: Includes valid `resetSyncStatusForBatch` call.
    *   **`clearAllBatches`**: Includes valid `resetAllSyncStatus` call.
    *   **`applySmartMerge`**: The shared logic ensuring Col A/B visual continuity.

### B. The Spreadsheet Data (Structure)
*   **ORDERING LIST (Production)**:
    *   **Batch ID (Col Q, Index 17)**: The primary key for "Undo".
    *   **Visual Structure**: 5 blank rows buffer required per section.
*   **TRIAL-LAYOUT CONFIGURATION (Staging)**:
    *   **Sync Status (Col J, Index 10)**: "SYNCED" marker.
    *   **Batch ID (Col S, Index 19)**: **NEW**. Stores the ID to allow targeted un-syncing.

---

## 3. Current Implementation Snapshot (The 'Now')

**Existing Functionality (Working & Tested):**
1.  **Sync to Order List (Phase 13):**
    *   Validates Staging -> Generates ID -> Writes to Production (with ID) -> Marks Staging (with ID).
2.  **Undo Last Sync (Phase 13):**
    *   Finds latest ID in Production -> Deletes Rows -> Finds same ID in Staging -> Clears Status.
3.  **Clear All Batches (Phase 13):**
    *   Deletes all ID rows in Production -> Repairs 5-row buffer -> Wipes Col J & S in Staging.

**State Variables:**
*   `BATCH_ID_COL_INDEX = 17` (Ordering List).
*   `STAGING_BATCH_COL = 19` (Trial Layout - Implicitly handled in `markRowsAsSynced`).

---

## 4. The Master Roadmap (Future Context)

**The Full Plan:**
*   **[COMPLETED]** Phase 12: Undo Last Sync (Basic).
*   **[COMPLETED]** Phase 13: Robust Sync (Batch ID in Staging + Active Reset).
*   **[PENDING]** User Acceptance Testing (UAT).
*   **[PENDING]** "History" Mode (Optional future discussion).

**Immediate Next Steps:**
1.  **Load:** Ingest this file set.
2.  **Verify:** Check `ConfiguratorBackend.gs` for the helper functions `resetSyncStatusForBatch` and `resetAllSyncStatus`.
3.  **Wait:** User will likely test the Sync/Undo cycle.

---

## 5. Development History & Decision Log

**Critical Decisions:**
*   **Column S for Staging ID:** We moved from R to S (Index 19) per user request to avoid potential conflict with future Spare Parts expansion or "Other" columns.
*   **Central ID Generation:** Moving ID generation to `Main.gs` was crucial to ensure the *exact same string* reached both sheets. Using `new Date()` inside separate functions would have caused millisecond mismatches, breaking the "Undo" link.

---

## 6. Ingestion & Standby Instructions

**STRICT INSTRUCTION FOR NEW INSTANCE:**
1.  **DO NOT generate code immediately.**
2.  **READ** `ConfiguratorBackend.gs` and `Main.gs` to confirm the column indices (Q=17, S=19).
3.  **CONFIRM** exactly as follows:
    > "Context Serialized. Phase 13 (Robust Sync & Reset) verified. I see the Batch ID tracking in Column S (Staging) and Column Q (Production). Ready for instructions."
