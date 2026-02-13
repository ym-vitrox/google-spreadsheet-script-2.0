# DEEP-STATE CONTEXT SERIALIZATION PROTOCOL (PHASE 8 COMPLETE)

**TO:** New Gemini Instance (Senior Technical Lead)
**FROM:** Previous Session Context Manager
**SUBJECT:** Project Handover - ViTrox "Module Configurator" (Phase 8 - DYNAMIC HEADER & SYNC OPTIMIZATION)
**DATE:** January 30, 2026
**STATUS:** **STABLE & ROBUST**. Sync Logic now features "Dynamic Merge Detection", "Strict Append (B-Slots 10-32)", and "Grouped Tooling Numbering".

---

## 1. Comprehensive Project Definition
**The Problem:** The client manages complex BOMs for "PX730i/PX740i" machines. They use a custom Google Sheets Sidebar App to configure machine modules (Vision PC, Base Modules, Tooling). The layout is dynamic, involving merged headers and variable column widths.
**The Workflow:** User Configures (Sidebar) -> Payload sent to Backend -> Script writes to Staging Sheet ("TRIAL-LAYOUT CONFIGURATION") -> User clicks "Sync" -> Script updates Production Sheet ("ORDERING LIST").
**Current Phase:** Phase 8 - **Dynamic Robustness**. We have moved away from rigid column indexing. The system now *scans* the sheet structure to find headers like "VCM" and "VALVE SET" dynamically. We also enforced strict item numbering rules for Tooling (Grouping) and strict data boundaries (B10-B32).

---

## 2. Asset Interpretation Guide (For the Attached Files)

### A. The Codebase
*   **`ModuleConfigurator.html` (Frontend)**:
    *   **Layout:** "Machine Setup" is the default tab.
    *   **Transport:** Uses manual `JSON.stringify` to avoid Code 10 errors.
*   **`ConfiguratorBackend.gs` (The Brain)**:
    *   **`extractProductionData` (The Reader - HEAVILY MODIFIED):**
        *   **Dynamic Header Scan:** Finds "VCM" anchor -> Detects merged columns -> Maps subsequent headers (e.g., "VALVE SET").
        *   **Dynamic Option Read:** Reads the specific Option Name/ID from the row *immediately below* the detected header.
        *   **Strict B-Slot Scan:** Forces data extraction to occur ONLY between rows where Col G matches `B10` through `B32`.
        *   **Empty Module Skip:** Strictly ignores rows where Module ID (Col K) is empty, preventing "phantom syncs" even if checkboxes are ticked.
    *   **`injectProductionData` (The Writer):**
        *   **Strategy:** "Strict Append" (Additive Only).
        *   **Header Mapping:** Explicitly maps "VALVE SET" items to the **"OTHERS"** section (User Request).
    *   **`insertRowsIntoSection` (The Helper):**
        *   **Remark Logic:** Scans existing data for unreleased duplicates. If found, marks the duplicate with *"haven't release, please take note"*.
*   **`Main.gs`**:
    *   **`runProductionSync`:** The main trigger.
    *   **`processToolingOptions`:** Syncs tooling data to `REF_DATA` (supports hierarchy).

### B. The Spreadsheet Data
*   **REF_DATA**: Master database.
*   **TRIAL-LAYOUT CONFIGURATION (Staging)**:
    *   **Key Anchors:** "VCM" (Header), "VALVE SET" (Header).
    *   **Data Zone:** Rows identified by Col G containing `B10`...`B32`.
*   **ORDERING LIST (Production)**: Target Sheet. Sections: `PC`, `CONFIG`, `CORE`, `MODULE`, `OTHERS`, `TOOLING`.

---

## 3. Current Implementation Snapshot (The 'Now')

**Existing Functionality (Working & Tested):**
1.  **Dynamic Header Detection:** The script no longer breaks if "VCM" or "VALVE SET" columns are shifted or resized (merged). It detects the merge range automatically.
2.  **Grouped Tooling Numbering:** In the "TOOLING" section, multiple options under the same Parent Tool share a single Item Number (e.g., "1. Option A, Option B").
3.  **Strict B-Slot Enforcement:** Logic completely ignores data outside the B10-B32 range.
4.  **Empty Row Safeguard:** Rows with no Module ID are skipped during sync.
5.  **Valve Set Mapping:** "VALVE SET" checkboxes now correctly populate the "OTHERS" section in production.

**Logic Flow:**
*   User clicks "Sync".
*   Script finds "VCM" -> Maps headers (cols) + options (row below).
*   Script scans B10-B32.
*   If Row has Module ID AND is not "SYNCED":
    *   Extract Module, Vision, Tooling.
    *   Check flagged columns against the Dynamic Header Map.
    *   Push items to respective arrays (MODULE, TOOLING, OTHERS).
*   Script appends to "ORDERING LIST".
*   Script marks processed rows as "SYNCED" in Staging.

---

## 4. The Master Roadmap (Future Context)

**The Full Plan:**
*   **[COMPLETED]** Phase 7c: Sync Logic (Strict Append + Remarks).
*   **[COMPLETED]** Phase 8: Dynamic Header Scanning & Tooling Grouping.
*   **[PENDING]** Phase 9: Final Deployment & User Acceptance Testing (UAT).
*   **[PENDING]** Future: Potential "Edit Mode" (currently only Append is supported).

**Immediate Next Steps:**
1.  **Ingest:** Load the attached files into the new instance.
2.  **Verify:** Read `ConfiguratorBackend.gs` to confirm the `while (colCursor <= MAX_SCAN_COL)` loop exists (proof of dynamic scanning).
3.  **Standby:** Await user instructions for UAT or further refinement.

---

## 5. Development History & Decision Log

**Key Decisions:**
*   **Dynamic vs Fixed:** We switched from Fixed Column Indexing (fragile) to Dynamic Merge Detection (robust) to handle layout changes.
*   **Valve Set Mapping:** User explicitly requested "VALVE SET" to map to "OTHERS", overruling previous logic.
*   **Empty Skip:** We introduced a strict check for `moduleID === ""` to prevent the sync from processing empty rows that might have accidental checkbox ticks.

---

## 6. Ingestion & Standby Instructions

**STRICT INSTRUCTION FOR NEW INSTANCE:**
1.  **DO NOT generate code immediately.**
2.  Wait for the user to upload the files (`ConfiguratorBackend.gs`, `Main.gs`, `ModuleConfigurator.html`).
3.  **CRITICAL:** Read `extractProductionData` in `ConfiguratorBackend.gs`. Confirm you see the **Dynamic Header Scanning Strategy** comment block.
4.  **REPLY ONLY:** "Context Serialized. Phase 8 (Dynamic Headers & Sync Optimization) logic confirmed. Ready for next instructions."
