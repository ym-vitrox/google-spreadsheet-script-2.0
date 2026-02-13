# DEEP-STATE CONTEXT SERIALIZATION PROTOCOL (PHASE 7c COMPLETE)

**TO:** New Gemini Instance (Senior Technical Lead)
**FROM:** Previous Session Context Manager
**SUBJECT:** Project Handover - ViTrox "Module Configurator" (Phase 7c - STRICT APPEND & REMARK LOGIC)
**DATE:** January 29, 2026
**STATUS:** **STABLE & ENHANCED**. Sync Logic now supports "Strict Append" for full history and "Duplicate Remarking" for neglected items.

---

## 1. Comprehensive Project Definition
**The Problem:** The client manages complex BOMs for "PX730i/PX740i" machines. They use a custom Google Sheets Sidebar App to configure machine modules (Vision PC, Base Modules, Tooling).
**The Workflow:** User Configures (Sidebar) -> Payload sent to Backend -> Script writes to Staging Sheet ("TRIAL-LAYOUT CONFIGURATION") -> User clicks "Sync" -> Script updates Production Sheet ("ORDERING LIST").
**Current Phase:** Phase 7c - **Sync Logic Finalization**. We have implemented a "Strict Append" strategy for *all* sections (Vision PC, Config Modules, Tooling, Core) to ensure a complete history log, combined with a "Remark Logic" to warn about unreleased duplicates.

---

## 2. Asset Interpretation Guide (For the Attached Files)

### A. The Codebase
*   **`ModuleConfigurator.html` (Frontend)**:
    *   **Layout:** "Machine Setup" is the default tab.
    *   **Transport:** Uses manual `JSON.stringify` to avoid Code 10 errors.
*   **`ConfiguratorBackend.gs` (The Brain)**:
    *   **`extractProductionData` (The Reader):**
        *   **Sources:**
            *   **Vision PC:** Cell C2.
            *   **Configurable Base Modules:** Block between Header & "Base Module Tooling".
            *   **Base Tooling:** Block between "Base Module Tooling" & "Comment".
            *   **CORE Items:** Fetched externally via `fetchCoreItemsFromExternal`.
            *   **Modules:** B-Slot Matrix configuration.
    *   **`injectProductionData` (The Writer):**
        *   **Strategy:** "Strict Append" (Additive Only).
        *   **Payload Struct:** `{ PC: [], CONFIG: [], CORE: [], TOOLING: [], MODULE: [] ... }`.
    *   **`insertRowsIntoSection` (The Helper):**
        *   **Additive Flow:** 
            1.  Finds the active section (Header -> Next Header).
            2.  **Scan Logic:** Scans *existing* section data. If it finds a row with the **Same ID** as a new item AND `Released == False`, it writes a **Remark** in Column J: *"haven't release, please take note"*.
            3.  **Append Logic:** Always inserts new rows at the bottom of the list.
*   **`Main.gs`**:
    *   **`runProductionSync` (The Trigger):**
        *   **Gatekeeper Update:** Now runs if *any* payload section (PC, CONFIG, TOOLING, CORE, MODULE) has data. It no longer relies solely on `rowsToMarkSynced`.
    *   **`processToolingOptions`:** Syncs tooling data to `REF_DATA` (relaxed block logic).

### B. The Spreadsheet Data
*   **REF_DATA**: Master database. Columns A:B contain Base Data (Vision PC descs). Columns U:V contain Tooling Descs.
*   **TRIAL-LAYOUT CONFIGURATION (Staging)**: Source for the Sync.
*   **ORDERING LIST (Production)**: Target. Contains headers: `PC`, `CONFIG`, `CORE`, `MODULE`.

---

## 3. Current Implementation Snapshot (The 'Now')

**Existing Functionality (Working & Tested):**
1.  **Vision PC Sync:** Singleton Logic removed. Now follows **Strict Append**.
2.  **Duplicate Marking:** If you sync an unreleased item twice, the first one gets a warning remark.
3.  **Comprehensive Extraction:** Config Modules, Base Tooling, and External CORE items are now fully extracted and synced.
4.  **Gatekeeper:** Sync button works even if only Machine Setup data is new.

**Logic Flow:**
*   User clicks "Sync".
*   `runProductionSync` checks: "Is there *any* new data?" -> Yes.
*   `extractProductionData` pulls all sections (including Vision PC desc lookup from REF_DATA A:B).
*   `injectProductionData` iterates sections.
*   `insertRowsIntoSection` scans for unreleased duplicates -> Marks them -> Appends new rows.
*   Success.

---

## 4. The Master Roadmap (Future Context)

**The Full Plan:**
*   **[COMPLETED]** Phase 7a/b: Layout & Tooling Fixes.
*   **[COMPLETED]** Phase 7c: Sync Logic (Strict Append + Remarks).
    *   7c.1 Vision PC (Append).
    *   7c.2 Configurable Base Modules (Append).
    *   7c.3 Base Module Tooling (Append).
    *   7c.4 External CORE (Append).
*   **[PENDING]** Phase 8: Final Verified Deployment. User validation of the new Append Logic.

**Immediate Next Steps:**
1.  **Verify:** User to confirm the "Duplicate Remark" behaves exactly as requested.
2.  **Cleanup:** Remove any debug logging if present.

---

## 5. Development History & Decision Log

**Key Decisions:**
*   **Switch to Append:** We moved from "Singleton/Smart Replace" to **"Strict Append"** to preserve a complete history log.
*   **Duplicate Handling:** Instead of preventing duplicates, we **Flag** them (Column J Remark) if they are unreleased. This enforces user attention.
*   **Tooling Strategy:** Base Tooling and Module Tooling share the `TOOLING` section. Both use Append logic.
*   **External CORE:** Sourced from specific Sheet ID `1nTSOqK4nGRkUEHGFnUF30gRCGFQMo6I2l8vhZB-NkSA`.

---

## 6. Ingestion & Standby Instructions

**STRICT INSTRUCTION FOR NEW INSTANCE:**
1.  **DO NOT generate code immediately.**
2.  Wait for the user to upload the files (`ConfiguratorBackend.gs`, `Main.gs`, `ModuleConfigurator.html`).
3.  **CRITICAL:** Read `insertRowsIntoSection` in `ConfiguratorBackend.gs`. Verify the "Remark Scanning" block is present before generating any new code.
4.  **REPLY ONLY:** "Context Serialized. I have verified the Phase 7c 'Strict Append & Remark' Logic. Ready for Phase 8 Verification."
