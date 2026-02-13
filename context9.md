# DEEP-STATE CONTEXT SERIALIZATION PROTOCOL (PHASE 11 COMPLETE)

**TO:** New Gemini Instance (Senior Technical Lead)
**FROM:** Previous Session Context Manager
**SUBJECT:** Project Handover - ViTrox "Module Configurator" (Phase 11 - TOOLING EXCLUSION & ROTATIONAL LOGIC)
**DATE:** February 4, 2026
**STATUS:** **STABLE**. Tooling Exclusion for `430001-A688` implemented. Rotational Logic active.

---

## 1. Comprehensive Project Definition
**The Problem:** The user needs to configure complex BOMs for "PX730i/PX740i" machines.
1.  **Electrical Rotation:** Modules have 4 specific electrical parts. Logic must scan `TRIAL-LAYOUT` to find gaps and assign the next available instance.
2.  **Tooling Exclusion:** Certain legacy tools (e.g., `430001-A688`) appear in standard lists but must sometimes be explicitly excluded by the user so they don't appear in the BOM.
3.  **Machine Setup:** A complex hierarchical form to separate Machine Setup from Module Configuration.

**The Solution (Phase 11):**
We added a "Tooling Exclusion" feature. For specific IDs (currently `430001-A688`), the UI injects an "excluded this tooling" option. If selected, the frontend logic **filters this item out** of the payload entirely, ensuring backend "Smart Overwrite" logic (which wipes and rewrites the section) effectively deletes/omits the row.

---

## 2. Asset Interpretation Guide (For the Attached Files)

### A. The Codebase
*   **`ModuleConfigurator.html` (The UI - UPDATED Phase 11)**:
    *   **`renderBaseModuleTooling` (UPDATED):**
        *   Contains logic to target `tool.id === "430001-A688"`.
        *   Injects `<option value="EXCLUDE_TOOLING">excluded this tooling</option>` into the dropdown.
    *   **`saveMachineSetup` (UPDATED):**
        *   **Payload Filtering:** Before pushing to `payload.baseTooling`, it checks `if (select.value === 'EXCLUDE_TOOLING')`. If true, it **skips** the push.
    *   **`saveConfiguration`:** Handles the standard module save loop with Auto-Refresh.
*   **`ConfiguratorBackend.gs` (The Brain)**:
    *   **`scanUsedElectrical(moduleID)`:** Returns a Set of used electrical IDs by scanning `TRIAL-LAYOUT` Cols K & L.
    *   **`saveMachineSetup(payloadRaw)`:** Parses JSON payload (sent as string) and uses "Smart Overwrite" (Insert/Delete rows) to sync the lists. It relies on the frontend to provide the *correct* list; it does not filter defaults itself.

### B. The Spreadsheet Data (Structure)
*   **REF_DATA**:
    *   **Cols I & J:** Base Modules.
    *   **Cols W & X:** Electrical Parts (Rotational Source).
    *   **Cols U & V:** Tooling Master List.
*   **TRIAL-LAYOUT CONFIGURATION**:
    *   **"Machine Setup" Section:** Top of the sheet (Rows ~1-100). Validated by Anchors ("Configurable Base Module", "Base Module Tooling", "Comment").
    *   **"Module Configuration" Section:** Starts after "CONFIGURATION" header. Validated by "B10" slot anchor.

---

## 3. Current Implementation Snapshot (The 'Now')

**Existing Functionality (Working & Tested):**
1.  **Tooling Exclusion (Phase 11):**
    *   *Target:* `430001-A688 | List-Tooling Flipper`.
    *   *Action:* User selects "excluded this tooling".
    *   *Result:* Item is removed from `TRIAL-LAYOUT` upon save.
    *   *Recovery:* User selects normal option -> Item reappears.
2.  **Smart Rotational Logic (Phase 10):**
    *   Fills gaps in electrical sequence (e.g., if #2 is deleted, next add re-uses #2).
    *   Auto-Refreshes UI on save.
3.  **Machine Setup Sync (Phase 6):**
    *   Uses 3-Anchor System.
    *   Sends payload as String to avoid Transport Error 10.

---

## 4. The Master Roadmap (Future Context)

**The Full Plan:**
*   **[COMPLETED]** Phase 10: Rotational Logic & Auto-Refresh.
*   **[COMPLETED]** Phase 11: Tooling Exclusion logic.
*   **[PENDING]** Expansion of Exclusion Logic to other tools (if requested).
*   **[PENDING]** UAT & Long-term monitoring.

**Immediate Next Steps:**
1.  **Load:** Ingest this fileset.
2.  **Verify:** Check `ModuleConfigurator.html` for `EXCLUDE_TOOLING` string.
3.  **Standby:** Wait for `430001-A688` testing feedback or new requests.

---

## 5. Development History & Decision Log

**Critical Decisions:**
*   **Frontend vs Backend Filtering:** We chose **Frontend Filtering** for exclusion. The backend is designed to be a "dumb writer" of whatever payload it receives. This keeps the complex UI state logic (what "excluded" means) in the UI layer.
*   **Exclusion Phrasing:** User specifically requested lowercase "excluded this tooling".

---

## 6. Ingestion & Standby Instructions

**STRICT INSTRUCTION FOR NEW INSTANCE:**
1.  **DO NOT generate code immediately.**
2.  **READ** `ModuleConfigurator.html`. Search for `EXCLUDE_TOOLING`.
3.  **CONFIRM** exactly as follows:
    > "Context Serialized. Phase 11 (Tooling Exclusion) verified. I see the specific logic for 430001-A688 and the payload filtering. Ready for instructions."
