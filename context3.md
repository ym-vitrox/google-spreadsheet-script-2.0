# DEEP-STATE CONTEXT SERIALIZATION PROTOCOL

**TO:** New Gemini Instance (Senior Technical Lead)
**FROM:** Previous Session Context Manager
**SUBJECT:** Project Handover - ViTrox "Module Configurator" (Phase 7 - STABILIZED & REFINED)
**DATE:** January 27, 2026
**STATUS:** **STABLE** (Transport Error, Duplication, Ghost Rows, Empty Header Rows RESOLVED).

---

## 1. Comprehensive Project Definition
**The Problem:** The client manages complex BOMs for "PX730i/PX740i" machines. They use a custom Google Sheets Sidebar App to configure machine modules (Vision PC, Base Modules, Tooling).
**The Workflow:** User Configures (Sidebar) -> Payload sent to Backend -> Script writes to Staging Sheet ("TRIAL-LAYOUT CONFIGURATION") -> User clicks "Sync" -> Script updates Production Sheet ("ORDERING LIST").
**Current Phase:** Phase 7 - **Polish & Optimization**. We have resolved critical errors (Transport, Duplication) and implemented refined UX features (Smart Overwrite, Parent Descriptions, All Borders).

---

## 2. Asset Interpretation Guide (For the Attached Files)

### A. The Codebase
*   **`ModuleConfigurator.html` (Frontend)**:
    *   **CRITICAL:** The `saveMachineSetup` function now **manual stringifies** the JSON payload (`JSON.stringify(payload)`) before sending it to `google.script.run`. This is a specific fix for the Transport Error (Code 10).
    *   **CRITICAL:** The selector `document.querySelectorAll("#baseToolingList > .row-item")` uses a **strict child selector (`>`)** to prevent "Add Another Option" phantom rows.
*   **`ConfiguratorBackend.gs` (The Brain)**:
    *   **`saveMachineSetup(payloadRaw)`**: This is the main handler.
        *   **Step 1:** Manually parses `JSON.parse(payloadRaw)` to handle stringified transport.
        *   **Step 2:** Uses **"Safe Mode V2"** (Bulk Read A:B + JS Scan) to find anchors (`configModule`, `baseTooling`, `comment`) without `createTextFinder` timeouts.
        *   **Step 3:** Uses **"Smart Overwrite"** logic: Overwrites existing rows in place, inserts/deletes only at the tail end, and **includes the Header Row** as the first valid data slot effectively filling the "empty gap".
    *   **`getBaseModuleToolingList`**:
        *   **Feature:** Imports Parent Descriptions from `REF_DATA` Column U (format: `ParentID | Description`).
        *   **Logic:** Prioritizes this description over child-derived ones.
    *   **`fetchOptionsForTool(menuData, parentID)`**:
        *   **Fix:** Cleanly splits strings in Column U to match `ID` correctly against inputs.
*   **`Main.gs`**:
    *   **`processToolingOptions`**:
        *   **Data Source:** Reads "BOM Structure Tree Diagram" (Cols O:P) to get Parent IDs and Descriptions.
        *   **Output:** Writes to `REF_DATA` Column U in the `ID | Description` format.

### B. The Spreadsheet Data
*   **REF_DATA**: Master database. Column U now stores `ID | Description` pairs.
*   **TRIAL-LAYOUT CONFIGURATION (Staging)**:
    *   **Target Area:** Top Section ("Machine Setup").
    *   **Anchors:** "Configurable Base Module", "Base Module Tooling", "Comment".
    *   **Format:** All generated rows now have **Black Solid Borders** (`setBorder(true, true, true, true, true, true, "black", ...)`).

---

## 3. Current Implementation Snapshot (The 'Now')

**Existing Functionality (Working & Tested):**
1.  **Transport Stability:** String Payload strategy eliminates Code 10 errors.
2.  **Parent Description Import:** Tooling parents now show accurate descriptions sourced from BOM Tree (via `Main.gs` -> `REF_DATA`).
3.  **Smart Overwrite:** The script **never deletes** the first row (Header Row). It overwrites existing data and expands/contracts the bottom of the list only.
4.  **Header Row Utilization:** Data writing starts **at** the Header Row index (Cols C+), filling the visual gap next to the "Base Module Tooling:" label.
5.  **Visuals:** All generated output has strict Black Borders for a clean grid look.

**Logic Flow:**
*   Frontend Stringifies -> Backend Parses.
*   Backend Scans Anchors (Safe Mode).
*   Backend performs Smart Overwrite on Base Modules (Overwrite + Tail Adjustment).
*   Backend performs Smart Overwrite on Tooling (Overwrite + Tail Adjustment).
*   Returns `{ status: "success", log: "..." }`.

---

## 4. The Master Roadmap (Future Context)

**The Full Plan:**
*   **[COMPLETED]** Phase 6a: Fix Transport Error (String Payload).
*   **[COMPLETED]** Phase 6b: Fix Duplication/Gap Logic (Section Reset / Smart Overwrite).
*   **[COMPLETED]** Phase 6c: Parent Description Import (`ID | Desc` format).
*   **[COMPLETED]** Phase 6d: Smart Overwrite + Header Row Inclusion.
*   **[COMPLETED]** Phase 6e: UI Polish (All Borders).
*   **[NEXT]** Phase 7: **Sync Logic Verification**. We need to ensure `extractProductionData` (which pushes Staging -> Production) handles this refined, gap-less format correctly.
*   **[PENDING]** Phase 8: Final Deployment checks.

**Immediate Next Steps:**
1.  Verify `extractProductionData` in `ConfiguratorBackend.gs`.
2.  Ensure it uses "Safe Mode" scanning and correctly indexes the new "gap-less" data structure (where data starts on the Header Row).
3.  Test the "Sync" button.

---

## 5. Development History & Decision Log

**What we have tried (and kept/discarded):**
*   **KEPT:** Manual Stringify/Parse for Transport (Essential).
*   **KEPT:** Smart Overwrite (Prevents "flicker" and emptying of headers).
*   **DISCARDED:** `createTextFinder` (Unreliable/Slow).
*   **DISCARDED:** "Wipe & Replace" (Caused visual emptying of the list before refill).

**User Preferences:**
*   **Style:** Detailed, robust logic.
*   **Visuals:** "All Borders" (Black) is mandatory for the output table.
*   **Format:** `ParentID | Description` is the standard for `REF_DATA` Column U.

---

## 6. Ingestion & Standby Instructions

**STRICT INSTRUCTION FOR NEW INSTANCE:**
1.  **DO NOT generate code immediately.**
2.  Wait for the user to upload the files.
3.  Cross-reference `ConfiguratorBackend.gs` with this protocol. Identify the `saveMachineSetup` "Smart Overwrite" block and the `fetchOptionsForTool` split logic.
4.  **REPLY ONLY:** "Context Serialized. I have verified the Phase 7 'Refined State' (Smart Overwrite, Borders, Parent Parsing). I am ready to proceed with Phase 7 (Sync Logic Verification)."
