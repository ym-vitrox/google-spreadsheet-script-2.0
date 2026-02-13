DEEP-STATE CONTEXT SERIALIZATION PROTOCOL

TO: New Gemini Instance (Senior Technical Lead) FROM: Previous Session
Context Manager SUBJECT: Project Handover - ViTrox \"Module
Configurator\" (Phase 6 - Critical \"Gap & Duplication\" Bug & Transport
Error) DATE: January 27, 2026 PRIORITY: CRITICAL

SYSTEM INSTRUCTION: INGESTION PROTOCOL You are inheriting an active,
high-complexity Google Apps Script project called the \"Module
Configurator\". YOUR IMMEDIATE TASK: READ this entire protocol. WAIT for
the user to upload the codebase (.gs, .html) and Spreadsheet CSV
exports. ANALYZE the files against the \"Asset Interpretation Guide\"
below. REPLY ONLY: \"Context Serialized. I have analyzed the Phase 6
codebase, the Hybrid Tree logic, and the history of the
\'Gap/Duplication\' bug. I am ready to proceed.\" DO NOT GENERATE CODE
YET.

1\. Comprehensive Project Definition The Problem: The client manages
complex BOMs for \"PX730i/PX740i\" machines. They need a custom Google
Sheets Sidebar App to configure machine modules and \"Machine Setup\"
components (Vision PC, Base Modules, Tooling). The Workflow: User
Configures (Sidebar) -\> Payload sent to Backend -\> Script writes to
Staging Sheet (TRIAL-LAYOUT CONFIGURATION) -\> User clicks \"Sync\" -\>
Script extracts data and updates Production Sheet (ORDERING LIST).
Current Phase: Phase 6 - Machine Setup & Staging Restructuring. We are
strictly focused on the \"Base Module Tooling\" section of the Staging
Sheet. 2. Asset Interpretation Guide (For Attached Files) A. The
Codebase ModuleConfigurator.html (Frontend): Sends a JSON payload. The
baseTooling array contains objects with {id, desc, structure}. Note: It
currently sends empty objects if the user adds a row but leaves it
blank. ConfiguratorBackend.gs (The Brain - CRITICAL):
saveMachineSetup(payload): The function causing the current bug. It
attempts to save tooling to the Staging Sheet using a \"Collapse &
Rebuild\" strategy. extractProductionData(): Scans Staging to prepare
for Sync. Updated to scan the top section (Machine Setup).
getBaseModuleToolingList(): Fetches data from REF_DATA (Cols U:V).
Main.gs: Menu triggers. B. The Spreadsheet Data (CSVs) REF_DATA
(Database): Source of truth for modules and tooling. TRIAL-LAYOUT
CONFIGURATION (Staging - The Target): Top Section: \"Machine Setup\".
Anchors: VCM, Configurable Base Module, Base Module Tooling, Comment (or
Comment:). Layout: Col C (Parent ID), Col D (Child ID), Col E
(Description). ORDERING LIST (Production): Where data is synced to. 3.
Current Implementation Snapshot (The \'Now\') Architecture: We recently
moved from a 3-Level nested tree to a 2-Level Flattened Tree. Level 1:
Writes to Col C. Level 2: Writes to Col D. Description: Writes to Col E.
Current State: The code contains a \"Nuclear\" write strategy intended
to delete all rows between anchors and write fresh data. The Bug: \"The
Gap & Duplication Issue\" leading to \"Transport Error\". When saving,
the script needs to identify the \"Base Module Tooling\" section and
clear it before writing new data. The Anchor Detection Logic is failing,
leading to either massive deletions (Gap) or appending data without
clearing (Duplication). Recent attempts to fix this logic have
introduced a \"Transport Error (Code 10)\", meaning the server-side
script is crashing or timing out before returning a response to the
client. 4. The Master Roadmap Phase 6 (Current): \[x\] View Switcher &
Complex Tooling Logic. \[x\] 2-Level Flattening Logic (Coded). \[ \]
CRITICAL: Fix the \"Gap + Duplication\" bug in saveMachineSetup without
causing Transport Errors. \[ \] Verify Sync Logic for new flattened
structure. Phase 7: Refinement & Comments Section. 5. Development
History & Decision Log (READ CAREFULLY) We have attempted multiple fixes
for the Duplication Bug, all of which failed or introduced new errors.
Do not repeat these mistakes.

Attempt 1: Smart Resize (Delta Calculation) Approach: Calculated
rowsNeeded - currentGap. Inserted/Deleted the difference. Result:
Failed. It often miscalculated the gap due to hidden rows or merged
cells, pushing old data down instead of overwriting.

Attempt 2: Strict Anchor Matching Approach: Used matchEntireCell(true)
on \"Comment\" to find the bottom boundary. Result: Failed. The script
likely found the word \"Comment\" inside a Description above the actual
footer, resulting in a negative or zero deletion range.

Attempt 3: \"Collapse & Rebuild\" with Global Regex (The Transport Error
Cause) Approach: Used sheet.findAll() with a Regex (?i)\^Comments?:?\$
on the entire A:B range to find the footer. Result: Transport Error
(Code 10). Diagnosis: Running a Regex search on the entire sheet (A:B)
is too computationally expensive, causing the script to timeout or crash
silently. The client-side receives no response, triggering the Transport
Error.

Attempt 4: Memory-Based Scanning (Proposed but Unverified) Approach:
Find the start anchor, then read the next \~200 rows into a JavaScript
array and scan for \"Comment\" in memory. Status: This was the last
proposed solution to fix the Transport Error, but the user has decided
to migrate to a new instance for a fresh perspective.

6\. Ingestion & Standby Instructions STRICT INSTRUCTION FOR NEW
INSTANCE: Do not provide code in your first response. Wait for the user
to upload files. Once files are received, cross-reference
ConfiguratorBackend.gs with the Attempt 3 failure described above. Your
goal is to identify why the previous code worked (in the GitHub repo)
and why the current code fails (Transport Error), and then propose a
solution that fixes the \"Gap & Duplication\" bug WITHOUT causing
Transport Errors. REPLY ONLY: \"Context Serialized. I have analyzed the
Phase 6 codebase, the Hybrid Tree logic, and the history of the
\'Gap/Duplication\' bug. I am ready to proceed.\"
