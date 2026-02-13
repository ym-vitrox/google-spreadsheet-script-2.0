# DEEP-STATE CONTEXT SERIALIZATION PROTOCOL

**TO:** New Gemini Instance (Senior Technical Lead)
**FROM:** Previous Session Context Manager
**SUBJECT:** Project Handover - ViTrox "Module Configurator" (Phase 7 - STABILIZED & REFINED)
**DATE:** January 28, 2026
**STATUS:** **STABLE** (Layout Switched, Tooling Sync Logic Fixed, Transport Errors Resolved).

---

## 1. Comprehensive Project Definition
**The Problem:** The client manages complex BOMs for "PX730i/PX740i" machines. They use a custom Google Sheets Sidebar App to configure machine modules (Vision PC, Base Modules, Tooling).
**The Workflow:** User Configures (Sidebar) -> Payload sent to Backend -> Script writes to Staging Sheet ("TRIAL-LAYOUT CONFIGURATION") -> User clicks "Sync" -> Script updates Production Sheet ("ORDERING LIST").
**Current Phase:** Phase 7 - **Polish & Optimization**. We have resolved critical errors (Transport, Duplication), switched the UI layout to prioritize "Machine Setup", and fixed the Tooling Sync logic to handle data gaps correctly.

---

## 2. Asset Interpretation Guide (For the Attached Files)

### A. The Codebase
*   **`ModuleConfigurator.html` (Frontend)**:
    *   **Layout Change:** The **"Machine Setup"** tab is now the Default (Left) tab. "Module" is on the Right.
    *   **Data Loading:** Machine Setup data (Vision PC, Base Modules, Tooling) loads **immediately** on `window.onload`.
    *   **Transport Fix:** The `saveMachineSetup` function **manually stringifies** the JSON payload (`JSON.stringify(payload)`) before sending it to `google.script.run` to prevent Transport Error (Code 10).
    *   **Deduplication:** The selector `document.querySelectorAll("#baseToolingList > .row-item")` uses a **strict child selector (`>`)** to prevent "Add Another Option" phantom rows.
*   **`ConfiguratorBackend.gs` (The Brain)**:
    *   **`saveMachineSetup(payloadRaw)`**: This is the main handler.
        *   **Step 1:** Manually parses `JSON.parse(payloadRaw)` to handle stringified transport.
        *   **Step 2:** Uses **"Safe Mode V2"** (Bulk Read A:B + JS Scan) to find anchors (`configModule`, `baseTooling`, `comment`) without `createTextFinder` timeouts.
        *   **Step 3:** Uses **"Smart Overwrite"** logic: Overwrites existing rows in place, inserts/deletes only at the tail end, and **includes the Header Row** as the first valid data slot effectively filling the "empty gap".
    *   **`getBaseModuleToolingList`**:
        *   **Feature:** Imports Parent Descriptions from `REF_DATA` Column U (format: `ParentID | Description`).
        *   **Logic:** Prioritizes this description over child-derived ones.
*   **`Main.gs`**:
    *   **`processToolingOptions`**: 
        *   **Role:** Syncs "Tooling Illustration" sheet to `REF_DATA`.
        *   **Strict Block Logic (Relaxed for Gaps):**
            *   **Start Block:** `[Parent ID]` found in Col A.
            *   **Continue Block:** Persistence across Empty Rows is **ALLOWED** (to handle gaps in source data).
            *   **Stop/Reset Block:** Only stops if **Non-Bracketed Text** is found in Col A (Visual Break like "Notes").

### B. The Spreadsheet Data
*   **REF_DATA**: Master database. Column U now stores `ID | Description` pairs.
*   **TRIAL-LAYOUT CONFIGURATION (Staging)**:
    *   **Target Area:** Top Section ("Machine Setup").
    *   **Anchors:** "Configurable Base Module", "Base Module Tooling", "Comment".
    *   **Format:** All generated rows now have **Black Solid Borders**.

---

## 3. Current Implementation Snapshot (The 'Now')

**Existing Functionality (Working & Tested):**
1.  **Transport Stability:** String Payload strategy eliminates Code 10 errors.
2.  **UI Layout:** "Machine Setup" is the default view. Setup data loads instantly.
3.  **Tooling Sync:** `Main.gs` correctly syncs tooling data, respecting visual blocks but allowing empty row gaps.
4.  **Smart Overwrite:** The script overwrites existing data and expands/contracts the bottom of the list only.
5.  **Visuals:** All generated output has strict Black Borders.

**Logic Flow:**
*   Frontend Stringifies -> Backend Parses.
*   Backend Scans Anchors (Safe Mode).
*   Backend performs Smart Overwrite on Base Modules.
*   Backend performs Smart Overwrite on Tooling (including H2 header row).
*   Returns `{ status: "success", log: "..." }`.

---

## 4. The Master Roadmap (Future Context)

**The Full Plan:**
*   **[COMPLETED]** Phase 6: Core Stability (Transport, Overwrite, Borders).
*   **[COMPLETED]** Phase 7a: Layout Switch (Machine Setup Default).
*   **[COMPLETED]** Phase 7b: Tooling Sync Logic Fix (Relaxed Block Logic).
*   **[NEXT]** Phase 7c: **Sync Logic Verification**. We need to ensure `extractProductionData` (which pushes Staging -> Production) handles this refined, gap-less format correctly.
*   **[PENDING]** Phase 8: Final Deployment checks.

**Immediate Next Steps:**
1.  Verify `extractProductionData` in `ConfiguratorBackend.gs`.
2.  Ensure it uses "Safe Mode" scanning and correctly indexes the new "gap-less" data structure (where data starts on the Header Row).
3.  Test the "Sync" button (Staging -> Production).

---

## 5. Development History & Decision Log

**What we have tried (and kept/discarded):**
*   **KEPT:** Manual Stringify/Parse for Transport (Essential).
*   **KEPT:** Smart Overwrite (Prevents "flicker" and emptying of headers).
*   **KEPT:** Relaxed Sync Logic (Allowed empty rows in Tooling Illustration to serve as continuations, not stops).
*   **DISCARDED:** Strict "Stop on Empty Row" Logic (Caused missing data because source data has gaps).
*   **DISCARDED:** `createTextFinder` (Unreliable/Slow).

**User Preferences:**
*   **Style:** Detailed, robust logic.
*   **Visuals:** "All Borders" (Black) is mandatory for the output table.
*   **Format:** `ParentID | Description` is the standard for `REF_DATA` Column U.

---

## 6. Ingestion & Standby Instructions

**STRICT INSTRUCTION FOR NEW INSTANCE:**
1.  **DO NOT generate code immediately.**
2.  Wait for the user to upload the files.
3.  Cross-reference `Main.gs` with this protocol. Identify the `processToolingOptions` logic (look for the "Stop Signal" comments).
4.  **REPLY ONLY:** "Context Serialized. I have verified the Phase 7 'Refined State' (Layout Switch, Tooling Sync Fix). I am ready to proceed with Phase 7c (Sync Logic Verification)."
