# DEEP-STATE CONTEXT SERIALIZATION PROTOCOL

**TO:** New Gemini Instance (Senior Technical Lead)  
**FROM:** Previous Session Context Manager  
**SUBJECT:** Project Handover - ViTrox "Module Configurator" (Phase 6 - STABILIZED)  
**DATE:** January 27, 2026  
**STATUS:** **STABLE** (Transport Error, Duplication, and Ghost Rows RESOLVED).

---

## 1. Comprehensive Project Definition
**The Problem:** The client manages complex BOMs for "PX730i/PX740i" machines. They use a custom Google Sheets Sidebar App to configure machine modules (Vision PC, Base Modules, Tooling).  
**The Workflow:** User Configures (Sidebar) -> Payload sent to Backend -> Script writes to Staging Sheet ("TRIAL-LAYOUT CONFIGURATION") -> User clicks "Sync" -> Script updates Production Sheet ("ORDERING LIST").  
**Current Phase:** Phase 6 - **Stability & Optimization**. We have just resolved a critical "Transport Error (Code 10)" and a persistent "Duplication/Gap" bug in the Staging Sheet generation.

---

## 2. Asset Interpretation Guide (For the Attached Files)

### A. The Codebase
*   **`ModuleConfigurator.html` (Frontend)**:
    *   **CRITICAL:** The `saveMachineSetup` function now **manual stringifies** the JSON payload (`JSON.stringify(payload)`) before sending it to `google.script.run`. This is a specific fix for the Transport Error.
    *   **CRITICAL:** The selector `document.querySelectorAll("#baseToolingList > .row-item")` uses a **strict child selector (`>`)**. This prevents "Add Another Option" sub-rows from being counted as duplicate tools.
*   **`ConfiguratorBackend.gs` (The Brain)**:
    *   **`saveMachineSetup(payloadRaw)`**: This is the main handler.
        *   **Step 1:** Manually parses `JSON.parse(payloadRaw)` (handling the stringified transport).
        *   **Step 2:** Uses **"Safe Mode V2"** logic: Instead of using `createTextFinder` or `findAll` (which caused timeouts), it performs a **Bulk Read** of columns A:B into memory and scans them using JavaScript loops to find anchors (`configModule`, `baseTooling`, `comment`).
    *   **`buildMasterDictionary` / `extractProductionData`**: Helper functions for the Sync process.

### B. The Spreadsheet Data (CSVs)
*   **REF_DATA**: Master database for Parts/Modules.
*   **TRIAL-LAYOUT CONFIGURATION (Staging)**:
    *   **Target Area:** Top Section ("Machine Setup").
    *   **Anchors:** "Vision PC", "Configurable Base Module", "Base Module Tooling", "Comment" (or "Comment:").
    *   *Note: Anchors may have trailing colons (e.g., "Base Module Tooling:"). Code handles this.*
*   **ORDERING LIST**: Production output (not focus of current task).

---

## 3. Current Implementation Snapshot (The 'Now')

**Existing Functionality (Working & Tested):**
1.  **Transport Stability:** The "String Payload" strategy (Frontend Stringify -> Backend Parse) has successfully eliminated the Transport Error (Code 10).
2.  **Section Reset Strategy:** The backend uses a "Wipe & Write" approach to prevent gaps/duplication:
    *   *Step A:* Calculate Gap between anchors.
    *   *Step B:* **Delete** the entire gap (cleaning the section).
    *   *Step C:* **Insert** exactly `N` rows required for the new data.
    *   *Step D:* Write data starting at `StartRow + 1` (preserving headers).
3.  **Phantom Row Prevention:** Frontend strictly filters out nested child rows. Backend strictly ignores any tool item without an `id`.

**Logic Flow:**
*   Frontend collects data -> Stringifies -> Send to Backend.
*   Backend Parses -> Bulk Reads "TRIAL-LAYOUT CONFIGURATION" (A:B) -> Finds Anchors in Memory.
*   Backend resets "Base Modules" section (Delete Gap -> Insert) -> Updates Anchors.
*   Backend resets "Tooling" section (Delete Gap -> Insert) -> Writes Data.
*   Returns `{ status: "success", log: "..." }`.

---

## 4. The Master Roadmap (Future Context)

**The Full Plan:**
*   **[COMPLETED]** Phase 6a: Fix Transport Error (Code 10) via String Payload.
*   **[COMPLETED]** Phase 6b: Fix Duplication/Gap Logic via "Section Reset".
*   **[COMPLETED]** Phase 6c: Fix "Add Another Option" Phantom Rows via Strict Selectors.
*   **[NEXT]** Phase 7: **Sync Logic Verification**. We need to ensure the `extractProductionData` function (which reads the Staging sheet) can correctly parse the *new* flattened format we are generating.
*   **[PENDING]** Phase 8: Final UI Polish & Comments Section.

**Immediate Next Steps:**
1.  Analyze `extractProductionData` in `ConfiguratorBackend.gs`.
2.  Verify if it uses the same "Safe Mode" (Memory Scan) or if it still relies on risky `createTextFinder` calls that might timeout.
3.  Test the "Sync to Production" button with the new data structure.

---

## 5. Development History & Decision Log

**What we have tried (and discarded):**
*   **FAIL:** Passing raw JSON objects to `google.script.run`. **Result:** Transport Error (Warden issues with complex/large payloads). **Decision:** ALWAYS stringify payloads.
*   **FAIL:** Using `sheet.createTextFinder("Comment").findAll()`. **Result:** Timeouts/Crashes on large sheets. **Decision:** ALWAYS use `sheet.getRange(1,1,500,2).getValues()` and scan in JavaScript (Safe Mode).
*   **FAIL:** "Diffing" logic (Insert only if `Need > Gap`). **Result:** Complex off-by-one errors and shifting headers. **Decision:** ALWAYS use "Section Reset" (Delete All -> Insert Needed).

**User Preferences:**
*   **Operating System:** Windows.
*   **Style:** Robust functionality over brevity. Prefer explicit logging for debugging.
*   **Critical Rule:** **NO** `createTextFinder` or `findAll` in the main flow. Use Memory Scanning.

---

## 6. Ingestion & Standby Instructions

**STRICT INSTRUCTION FOR NEW INSTANCE:**
1.  **DO NOT generate code immediately.**
2.  Wait for the user to upload `ModuleConfigurator.html`, `ConfiguratorBackend.gs`, and the CSV exports.
3.  Cross-reference the uploaded code with the **"Current Implementation Snapshot"** above. Verify that the `saveMachineSetup` function contains the **Manual Parse**, **Safe Mode Scan**, and **Section Reset** logic.
4.  **REPLY ONLY:** "Context Serialized. I have verified the Phase 6 'Stable State' (String Transport + Section Reset). I am ready to proceed with Phase 7 (Sync Logic Verification)."
