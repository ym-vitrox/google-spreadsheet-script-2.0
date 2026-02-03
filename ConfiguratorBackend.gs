/**
 * ConfiguratorBackend.gs
 * Server-side logic for the Module Configurator UI.
 * UPDATED: Phase 6 (Machine Setup - 2-Level Flattened Hierarchy & Sync Update)
 * FIX: "Collapse & Rebuild" strategy + Strict Sanitation to eliminate duplication bugs.
 * FIX: Robust Anchor Detection (Regex) + Safety Check for Anchor Order.
 */

// --- CONSTANTS ---
var RUBBER_TIP_PARENTS_BACKEND = ["430001-A689", "430001-A690", "430001-A691", "430001-A692"];
var RUBBER_TIP_SOURCE_ID_BACKEND = "430001-A380";

// Vision PC Target IDs
var VISION_PC_IDS = ["430001-A366", "430001-A367"];

// Complex Tooling Rules (Parent ID -> Mandatory Child ID)
var COMPLEX_TOOL_RULES = {
  "430001-A490": "430000-A748" // PUH Interface -> Assy-PUH Phoenix-v2.0
};

// Multi-Select Tools (Parent IDs that allow multiple option rows)
var MULTI_SELECT_TOOLS = ["430001-A380", "430001-A495"];

// NEW: Categorized Multi Tools (Complex Grouping)
var CATEGORIZED_MULTI_TOOLS = ["430001-A494"]; // Die Ejector Tooling

/**
 * 1. Get List of Modules (ID and Description)
 */
function getModuleList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];

  var rawValues = sheet.getRange(1, 3, lastRow, 2).getValues();
  var modules = [];

  for (var i = 0; i < rawValues.length; i++) {
    var id = String(rawValues[i][0]).trim();
    var desc = String(rawValues[i][1]).trim();

    if (id !== "" && id !== "Part ID") {
      modules.push({ id: id, desc: desc });
    }
  }
  return modules;
}

/**
 * 1.5 Get Vision PC Options (Hybrid Strategy)
 * Fetches descriptions for specific hardcoded IDs from REF_DATA
 */
function getVisionPCOptions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  // Columns A & B are Part ID & Description (Cols 1 & 2)
  var rawValues = sheet.getRange(1, 1, lastRow, 2).getValues();

  var options = [];

  // Create a quick lookup map for performance
  var lookupMap = {};
  for (var i = 0; i < rawValues.length; i++) {
    var rId = String(rawValues[i][0]).trim();
    var rDesc = String(rawValues[i][1]).trim();
    if (rId) lookupMap[rId] = rDesc;
  }

  // Build result based on HARDCODED Target IDs
  for (var j = 0; j < VISION_PC_IDS.length; j++) {
    var targetID = VISION_PC_IDS[j];
    var foundDesc = lookupMap[targetID] || "Description not found";
    options.push({ id: targetID, desc: foundDesc });
  }

  return options;
}

/**
 * 1.6 Get Configurable Base Module Options
 * Fetches from REF_DATA Columns I & J
 */
function getBaseModuleOptions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];

  // Columns I (9) and J (10)
  var rawValues = sheet.getRange(1, 9, lastRow, 2).getValues();
  var options = [];

  for (var i = 0; i < rawValues.length; i++) {
    var id = String(rawValues[i][0]).trim();
    var desc = String(rawValues[i][1]).trim();

    // Filter Logic:
    // 1. Exclude empty IDs
    // 2. Exclude "Part ID" header
    // 3. Exclude "List-" or "---" headers
    if (id && id !== "Part ID" && !id.toUpperCase().startsWith("LIST-") && id.indexOf("---") === -1) {
      options.push({ id: id, desc: desc });
    }
  }
  return options;
}

/**
 * 1.7 Get Base Module Tooling List (Hybrid Interceptor Logic)
 * Logic: Fetch U:V, Exclude IDs found in Y:Z, Apply Complex Rules, Multi-Select & Categorized Logic
 */
function getBaseModuleToolingList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];

  // 1. Fetch Data
  var toolingRange = sheet.getRange(1, 21, lastRow, 2).getValues(); // U:V
  var exclusionRange = sheet.getRange(1, 25, lastRow, 2).getValues(); // Y:Z

  // 2. Build Exclusion Set
  var excludedIDs = new Set();
  for (var i = 0; i < exclusionRange.length; i++) {
    var row = exclusionRange[i];
    // Y (row[0]) and Z (row[1])
    [row[0], row[1]].forEach(function (cellVal) {
      if (cellVal) {
        var str = String(cellVal).trim();
        // Handle delimited IDs (e.g., "A; B; C")
        var parts = str.split(/[;\n\r]+/);
        parts.forEach(function (p) {
          var clean = p.trim();
          if (clean) excludedIDs.add(clean);
        });
      }
    });
  }

  // 3. Process Tooling List (U:V)
  var uniqueTools = [];
  var seenParents = new Set();
  var menuData = toolingRange; // Reuse the U:V values

  for (var j = 0; j < toolingRange.length; j++) {
    var rawParentField = String(toolingRange[j][0]).trim();
    var partsParent = rawParentField.split("|");
    var parentID = partsParent[0].trim();
    var parentDescFromColU = (partsParent.length > 1) ? partsParent[1].trim() : "";

    var rawDescField = String(toolingRange[j][1]).trim(); // Col V

    // Filter Logic
    if (!parentID || parentID === "Part ID" || parentID.indexOf("---") > -1 || parentID.toUpperCase().startsWith("LIST-")) {
      continue;
    }

    // EXCLUSION LOGIC
    if (excludedIDs.has(parentID)) {
      continue;
    }

    if (seenParents.has(parentID)) {
      continue; // Already processed this parent
    }
    seenParents.add(parentID);

    // 4. Determine Description
    // Priority: Description from Col U (Master) > Description derived from Child (Legacy)
    var finalDesc = parentDescFromColU;

    if (!finalDesc || finalDesc === "") {
      // Fallback: Using old logic (taking description from child row if format is "ID::Desc")
      finalDesc = rawDescField;
      if (rawDescField.indexOf("::") > -1) {
        var parts = rawDescField.split("::");
        if (parts.length > 1) {
          finalDesc = parts[1].trim();
        }
      }
    }

    // 5. Fetch Options & Apply Interceptor

    // --- CASE A: CATEGORIZED MULTI (e.g., Die Ejector) ---
    if (CATEGORIZED_MULTI_TOOLS.includes(parentID)) {
      // Fetch raw options but preserve structure
      var groups = [];
      var currentGroup = null;

      // Manually scan the menuData for this parentID
      for (var m = 0; m < menuData.length; m++) {
        var rowParentRaw = String(menuData[m][0]);
        var rowID = rowParentRaw.split("|")[0].trim();
        if (rowID === parentID) {
          var line = String(menuData[m][1]).trim();

          if (line.startsWith("---")) {
            // New Group Detected (e.g., --- OPTIONAL EJECTOR NEEDLE [430001-A381] ---)
            var rawHeader = line.replace(/---/g, "").trim();

            // EXTRACT SUB-ID FROM HEADER IF PRESENT
            var groupName = rawHeader;
            var groupID = null;
            var match = rawHeader.match(/\[(.*?)\]/);
            if (match && match[1]) {
              groupID = match[1];
              groupName = rawHeader.replace(/\[.*?\]/, "").trim();
            }

            var groupType = (rawHeader.indexOf("OPTIONAL") > -1) ? "multi" : "single";

            currentGroup = {
              name: groupName,
              id: groupID, // Pass this to Frontend for Sub-Grouping
              type: groupType,
              options: []
            };
            groups.push(currentGroup);

          } else if (currentGroup && line) {
            // Add Item to Current Group
            var parts = line.split("::");
            var id = parts[0].trim();
            var desc = (parts.length > 1) ? parts[1].trim() : "";
            currentGroup.options.push({ id: id, desc: desc });
          }
        }
      }

      uniqueTools.push({
        id: parentID,
        desc: finalDesc,
        type: 'CATEGORIZED_MULTI',
        groups: groups
      });
    }

    // --- CASE B: COMPLEX RULES (e.g., PUH) ---
    else if (COMPLEX_TOOL_RULES[parentID]) {
      var mandatoryID = COMPLEX_TOOL_RULES[parentID];
      var allOptions = fetchOptionsForTool(menuData, parentID);

      var mandatoryItem = null;
      var selectableOptions = [];

      for (var k = 0; k < allOptions.length; k++) {
        if (allOptions[k].id === mandatoryID) {
          mandatoryItem = allOptions[k];
        } else {
          selectableOptions.push(allOptions[k]);
        }
      }

      uniqueTools.push({
        id: parentID,
        desc: finalDesc,
        type: 'COMPLEX', // Flag for UI
        mandatoryItem: mandatoryItem, // The Fixed Item
        options: selectableOptions // The Dropdown Items
      });

    }
    // --- CASE C: MULTI-SELECT TOOLS ---
    else if (MULTI_SELECT_TOOLS.includes(parentID)) {
      var options = fetchOptionsForTool(menuData, parentID);
      uniqueTools.push({
        id: parentID,
        desc: finalDesc,
        type: 'MULTI_SELECT', // Flag for UI
        options: options
      });
    }
    // --- CASE D: STANDARD ---
    else {
      var options = fetchOptionsForTool(menuData, parentID);
      uniqueTools.push({
        id: parentID,
        desc: finalDesc,
        type: 'STANDARD',
        options: options
      });
    }
  }

  return uniqueTools;
}

/**
 * 2. Get Full Details for Selected Module
 */
function getModuleDetails(moduleID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var refSheet = ss.getSheetByName("REF_DATA");
  var dataRange = refSheet.getRange("C:AF").getValues();
  var rowData = null;

  for (var i = 0; i < dataRange.length; i++) {
    if (String(dataRange[i][0]) === moduleID) {
      rowData = dataRange[i];
      break;
    }
  }

  if (!rowData) return { error: "Module ID not found." };

  // --- ELECTRICAL ---
  var elecResult = null;
  var elecIdsStr = rowData[20];
  var elecDescStr = rowData[21];

  if (elecIdsStr) {
    var eIds = String(elecIdsStr).split(';').map(function (s) { return s.trim(); });
    var eDescs = String(elecDescStr).split(';').map(function (s) { return s.trim(); });

    if (eIds.length > 0) {
      elecResult = {
        id: eIds[0],
        desc: eDescs[0] || "",
        note: "Instance 1 (Rotation A)"
      };
    }
  }

  // --- TOOLING ---
  var toolsResult = [];
  var toolIdsStr = rowData[22];
  var toolDescStr = rowData[23];

  if (toolIdsStr) {
    var tIds = String(toolIdsStr).split(';').map(function (s) { return s.trim(); });
    var tDescs = String(toolDescStr).split(';').map(function (s) { return s.trim(); });
    var menuData = refSheet.getRange("U:V").getValues();

    for (var k = 0; k < tIds.length; k++) {
      var tID = tIds[k];
      if (!tID) continue;

      var toolObj = {
        id: tID,
        desc: tDescs[k] || "",
        standardOptions: [],
        rubberTipOptions: [],
        requiresRubberTip: false
      };

      var stdOpts = fetchOptionsForTool(menuData, tID);
      if (stdOpts.length > 0) toolObj.standardOptions = stdOpts;

      if (RUBBER_TIP_PARENTS_BACKEND.includes(tID)) {
        var tipOpts = fetchOptionsForTool(menuData, RUBBER_TIP_SOURCE_ID_BACKEND);
        if (tipOpts.length > 0) {
          toolObj.requiresRubberTip = true;
          toolObj.rubberTipOptions = tipOpts;
        }
      }
      toolsResult.push(toolObj);
    }
  }

  // --- VISION ---
  var visionResult = { type: 'none', options: [] };
  var visionIdsStr = rowData[26];
  var visionDescStr = rowData[27];

  if (visionIdsStr) {
    var vIds = String(visionIdsStr).split(';').map(function (s) { return s.trim(); });
    var vDescs = String(visionDescStr).split(';').map(function (s) { return s.trim(); });

    var vOptions = [];
    for (var v = 0; v < vIds.length; v++) {
      if (vIds[v]) {
        vOptions.push({ id: vIds[v], desc: vDescs[v] || "" });
      }
    }

    if (vOptions.length === 1) {
      visionResult.type = 'fixed';
      visionResult.options = vOptions;
    } else if (vOptions.length > 1) {
      visionResult.type = 'select';
      visionResult.options = vOptions;
    }
  }

  // --- JIGS ---
  var jigsResult = [];
  var jigIdsStr = rowData[24];
  var jigDescStr = rowData[25];
  if (jigIdsStr) {
    var jIds = String(jigIdsStr).split(';').map(function (s) { return s.trim(); });
    var jDescs = String(jigDescStr).split(';').map(function (s) { return s.trim(); });
    for (var j = 0; j < jIds.length; j++) {
      if (jIds[j]) jigsResult.push({ id: jIds[j], desc: jDescs[j] || "" });
    }
  }

  // --- SPARE PARTS ---
  var sparePartsResult = [];
  var spIdsStr = rowData[28];
  var spDescStr = rowData[29];

  if (spIdsStr) {
    var spIds = String(spIdsStr).split(';').map(function (s) { return s.trim(); });
    var spDescs = String(spDescStr).split(';').map(function (s) { return s.trim(); });

    for (var s = 0; s < spIds.length; s++) {
      if (spIds[s]) {
        sparePartsResult.push({ id: spIds[s], desc: spDescs[s] || "" });
      }
    }
  }

  return {
    moduleID: moduleID,
    electrical: elecResult,
    tools: toolsResult,
    jigs: jigsResult,
    vision: visionResult,
    spareParts: sparePartsResult
  };
}

function fetchOptionsForTool(menuData, parentID) {
  var options = [];
  var foundParent = false;
  var currentCategory = "Standard Options";

  for (var i = 0; i < menuData.length; i++) {
    var rawParent = String(menuData[i][0]);
    var parentIDCheck = rawParent.split("|")[0].trim();
    var rowChild = String(menuData[i][1]);

    if (parentIDCheck === parentID) {
      foundParent = true;
      if (rowChild) {
        if (rowChild.indexOf("---") > -1) {
          currentCategory = rowChild.replace(/---/g, '').trim();
        }
        else {
          var parts = rowChild.split('::');
          var id = parts[0].trim();
          var desc = "";
          if (parts.length > 1) desc = parts[1].trim();

          options.push({ id: id, desc: desc, category: currentCategory });
        }
      }
    } else if (foundParent) {
      if (rawParent !== "") break;
    }
  }
  return options;
}

/**
 * 2.5 Get Layout Header Labels
 * UPDATED Phase 6: Dynamic Anchor Scanning
 */
function getLayoutHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) return null;

  // Find VCM Anchor Row dynamically
  var vcmFinder = sheet.getRange("B:B").createTextFinder("VCM").matchEntireCell(true).findNext();
  if (!vcmFinder) {
    throw new Error("Critical Error: 'VCM' anchor not found in Column B of Trial Layout. Please ensure headers exist.");
  }
  var vcmRowIndex = vcmFinder.getRow();

  // Read Headers (Anchor Row) and Options (Row below Anchor)
  // B:F = 5 columns
  var rawData = sheet.getRange(vcmRowIndex, 2, 2, 5).getValues();
  var rowCategories = rawData[0]; // Header Row
  var rowOptions = rawData[1];    // Option Row

  return {
    group1: {
      title: rowCategories[0], // B "VCM"
      options: [
        { colIndex: 2, label: rowOptions[0] }, // B
        { colIndex: 3, label: rowOptions[1] }  // C
      ]
    },
    group2: {
      title: rowCategories[2], // D "VALVE SET"
      options: [
        { colIndex: 4, label: rowOptions[2] }, // D
        { colIndex: 5, label: rowOptions[3] }, // E
        { colIndex: 6, label: rowOptions[4] }  // F
      ]
    }
  };
}

/**
 * 3. SAVE CONFIGURATION TO TRIAL LAYOUT
 * UPDATED Phase 6: "Look Before You Leap" Strategy (Correctly finds B10)
 */
function saveConfiguration(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) throw new Error("Sheet 'TRIAL-LAYOUT CONFIGURATION' not found.");

  var lastRow = sheet.getLastRow();

  // STEP 1: Find the "CONFIGURATION" Fence
  var configHeaderFinder = sheet.getRange("A:B").createTextFinder("CONFIGURATION").matchEntireCell(true).findNext();

  if (!configHeaderFinder) {
    throw new Error("Critical Error: 'CONFIGURATION' header not found in Column A or B. Cannot safely locate list start.");
  }

  var startScanRow = configHeaderFinder.getRow();

  // Scan Column G (Turret Location) starting from the header row downwards
  var scanRange = sheet.getRange(startScanRow, 7, lastRow - startScanRow + 1, 2).getValues();

  var targetRowIndex = -1;
  var startScanning = false;

  for (var i = 0; i < scanRange.length; i++) {
    var slotID = String(scanRange[i][0]).trim(); // Col G
    var desc = String(scanRange[i][1]).trim();   // Col H

    // STRICT VALIDATION: Is this a "B" slot?
    var isBSlot = slotID.startsWith("B") && !isNaN(parseInt(slotID.substring(1)));

    if (!startScanning) {
      if (isBSlot) {
        startScanning = true; // Found the first valid slot (e.g., B10)
      }
    }

    if (startScanning) {
      if (slotID === "B33" || (slotID.startsWith("B") && parseInt(slotID.substring(1)) > 32)) {
        break; // End of list (B33 is stop)
      }

      // If desc is empty, we found our target
      if (desc === "") {
        targetRowIndex = startScanRow + i; // Convert relative index i back to absolute row
        break;
      }
    }
  }

  if (targetRowIndex === -1) {
    throw new Error("Configuration List is full (B10-B32) or could not be located at 'CONFIGURATION' header.");
  }

  // WRITE DATA
  sheet.getRange(targetRowIndex, 8).setValue(payload.moduleDesc);
  sheet.getRange(targetRowIndex, 9).setValue(payload.numberVision || 1);
  sheet.getRange(targetRowIndex, 11).setValue(payload.moduleID);
  sheet.getRange(targetRowIndex, 12).setValue(payload.elecID);
  sheet.getRange(targetRowIndex, 13).setValue(payload.visionID);
  sheet.getRange(targetRowIndex, 14).setValue(payload.toolOptionID);
  sheet.getRange(targetRowIndex, 15).setValue(payload.rubberTipID);
  sheet.getRange(targetRowIndex, 16).setValue(payload.jigID);
  sheet.getRange(targetRowIndex, 17).setValue(payload.sparePartsID);

  if (payload.layoutFlags) {
    for (var col = 2; col <= 6; col++) {
      var val = payload.layoutFlags[String(col)] === true;
      var cell = sheet.getRange(targetRowIndex, col);
      var rule = cell.getDataValidation();
      var isCheckbox = (rule != null && rule.getCriteriaType() == SpreadsheetApp.DataValidationCriteria.CHECKBOX);
      if (!isCheckbox) { cell.insertCheckboxes(); }
      cell.setValue(val);
    }
  }

  var slotLabel = sheet.getRange(targetRowIndex, 7).getValue();
  return { status: "success", slot: slotLabel };
}

/**
 * 3.5 SAVE MACHINE SETUP (Phase 6 - SAFE MODE V2 + STRING TRANSPORT)
 * FIX: Payload sent as String to bypass Transport Error. Parsed manually.
 * FIX: Replaced ALL TextFinder calls with Bulk-Read + JS Scan to eliminate Timeouts.
 */
function saveMachineSetup(payloadRaw) {
  var log = ["Init"]; // Trace Log

  // 1. Manual Parsing (Fix for Transport Error)
  var payload = null;
  try {
    payload = (typeof payloadRaw === 'string') ? JSON.parse(payloadRaw) : payloadRaw;
  } catch (e) {
    return { status: "error", message: "Server Payload Parse Failed: " + e.message };
  }

  // GUARD: Basic Payload Validation
  if (!payload) {
    return { status: "error", message: "No payload received by server.", log: "No Payload" };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
    if (!sheet) throw new Error("Sheet 'TRIAL-LAYOUT CONFIGURATION' not found.");

    // 1. Save Vision PC & Top Section (No scanning needed here, fixed cells)
    if (payload.visionPC) {
      sheet.getRange(2, 3).setValue(payload.visionPC);
    }

    // --- SAFE MODE SCANNING (SINGLE READ) ---
    var SEARCH_DEPTH = 500; // Look in top 500 rows only
    var lastSheetRow = sheet.getLastRow();
    var actualDepth = Math.min(SEARCH_DEPTH, lastSheetRow);

    log.push("Reading Top " + actualDepth);
    var bigData = sheet.getRange(1, 1, actualDepth, 2).getValues(); // Read A:B

    // Locate Anchors in Memory
    var anchorMap = {
      configModule: -1,
      baseTooling: -1,
      comment: -1
    };

    for (var i = 0; i < bigData.length; i++) {
      var rowNum = i + 1;
      var valA = String(bigData[i][0]).trim();
      var valB = String(bigData[i][1]).trim();

      // Match 1: Configurable Base Module
      // Allow for optional trailing colon
      var cleanA = valA.replace(/:$/, "");
      var cleanB = valB.replace(/:$/, "");

      if (cleanA === "Configurable Base Module" || cleanB === "Configurable Base Module") {
        anchorMap.configModule = rowNum;
      }

      // Match 2: Base Module Tooling
      if (cleanA === "Base Module Tooling" || cleanB === "Base Module Tooling") {
        anchorMap.baseTooling = rowNum;
      }

      // Match 3: Comment (Strict Regex) - Only first valid footer after tooling
      if (anchorMap.baseTooling !== -1 && anchorMap.comment === -1) {
        if (/^comments?:?$/i.test(valA) || /^comments?:?$/i.test(valB)) {
          anchorMap.comment = rowNum;
        }
      }
    }
    log.push("Anchors Found: " + JSON.stringify(anchorMap));

    // 2. Save Configurable Base Modules (SECTION RESET STRATEGY)
    if (payload.baseModules && Array.isArray(payload.baseModules)) {
      if (anchorMap.configModule !== -1 && anchorMap.baseTooling !== -1) {
        var startRow = anchorMap.configModule;
        var nextSectionRow = anchorMap.baseTooling;
        var currentGap = nextSectionRow - startRow; // Gap includes Header Row now
        var itemsToSave = payload.baseModules;
        var requiredSlots = Math.max(itemsToSave.length, 3);

        log.push("BaseMod Update: Gap=" + currentGap + ", Need=" + requiredSlots);

        // --- SMART OVERWRITE STRATEGY (INCL HEADER ROW) ---
        // 1. Adjust Row Count (Insert or Delete at tail only)
        if (requiredSlots > currentGap) {
          var rowsToAdd = requiredSlots - currentGap;
          // Insert after the last existing slot (startRow + currentGap - 1)
          sheet.insertRowsAfter(startRow + currentGap - 1, rowsToAdd);
          // Update Map immediately (shifted down)
          anchorMap.baseTooling += rowsToAdd;
          anchorMap.comment += rowsToAdd;
        } else if (currentGap > requiredSlots) {
          var rowsToDelete = currentGap - requiredSlots;
          // Delete starting from the first extra row (startRow + requiredSlots)
          // Note: In 1-based indexing, if we filled 0..req-1 (rows start..start+req-1), next is start+req.
          sheet.deleteRows(startRow + requiredSlots, rowsToDelete);
          // Update Map immediately (shifted up)
          anchorMap.baseTooling -= rowsToDelete;
          anchorMap.comment -= rowsToDelete;
        }

        // 2. Overwrite Data
        for (var i = 0; i < requiredSlots; i++) {
          var targetRow = startRow + i; // Write STARTING AT header row
          var val = (i < itemsToSave.length) ? itemsToSave[i] : "";

          var cell = sheet.getRange(targetRow, 3);
          cell.setValue(val);

          // Formatting (Safe Merge)
          var mergeRange = sheet.getRange(targetRow, 3, 1, 5);
          try {
            // Only merge if not already merged or if we just created it
            mergeRange.merge();
            mergeRange.setVerticalAlignment("middle");
          } catch (e) { }

          // Borders (All Borders)
          sheet.getRange(targetRow, 3, 1, 5).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
        }
      }
    }

    // 3. Save Base Module Tooling (SECTION RESET STRATEGY)
    if (payload.baseTooling && Array.isArray(payload.baseTooling)) {
      if (anchorMap.baseTooling === -1) throw new Error("Anchor 'Base Module Tooling' not found.");
      if (anchorMap.comment === -1) throw new Error("Anchor 'Comment' footer not found.");

      var startRow = anchorMap.baseTooling;
      var endRow = anchorMap.comment;

      // Build Queue - SAFE MODE: Filter only valid items
      var writeQueue = [];
      var sanitizedTooling = payload.baseTooling.filter(function (item) {
        // STRICT VALIDATION: Must have an ID to be a valid tool row
        return (item.id && (item.desc || (item.structure && item.structure.length > 0)));
      });

      for (var p = 0; p < sanitizedTooling.length; p++) {
        var tool = sanitizedTooling[p];
        writeQueue.push({ level: 1, id: tool.id, desc: tool.desc });
        if (tool.structure) {
          for (var s = 0; s < tool.structure.length; s++) {
            var item = tool.structure[s];
            if (!item.id && !item.desc) continue;
            if (item.type === 'option') writeQueue.push({ level: 2, id: item.id, desc: item.desc });
            else if (item.type === 'group' && item.children) {
              for (var c = 0; c < item.children.length; c++) {
                var child = item.children[c];
                writeQueue.push({ level: 2, id: child.id, desc: child.desc });
              }
            }
          }
        }
      }

      var requiredSlots = Math.max(writeQueue.length, 1);
      var currentGap = endRow - startRow; // Gap includes Header Row

      log.push("Tooling Update: Gap=" + currentGap + ", Need=" + requiredSlots);

      // --- SMART OVERWRITE STRATEGY (INCL HEADER ROW) ---
      // 1. Adjust Row Count
      if (requiredSlots > currentGap) {
        var rowsToAdd = requiredSlots - currentGap;
        // Insert after the last existing slot
        sheet.insertRowsAfter(startRow + currentGap - 1, rowsToAdd);
        var newRange = sheet.getRange(startRow + currentGap, 1, rowsToAdd, sheet.getLastColumn());
        newRange.setTextRotation(0).setVerticalAlignment("middle").setFontWeight("normal").setFontStyle("normal");
      } else if (currentGap > requiredSlots) {
        var rowsToDelete = currentGap - requiredSlots;
        // Delete extra rows
        sheet.deleteRows(startRow + requiredSlots, rowsToDelete);
      }

      // 2. Overwrite Data
      var outputRange = sheet.getRange(startRow, 3, requiredSlots, 3); // Start at Header Row (C,D,E)
      var values = [];
      var fontWeights = [];
      for (var k = 0; k < requiredSlots; k++) {
        if (k < writeQueue.length) {
          var d = writeQueue[k];
          // Col C = Lvl 1, Col D = Lvl 2
          values.push([(d.level === 1 ? d.id : ""), (d.level === 2 ? d.id : ""), d.desc || ""]);
          fontWeights.push(["bold", "bold", "bold"]);
        } else {
          values.push(["", "", ""]);
          fontWeights.push(["normal", "normal", "normal"]);
        }
      }
      outputRange.setValues(values);
      outputRange.setFontWeights(fontWeights);

      // Cleanup Col F just to be safe
      sheet.getRange(startRow, 6, requiredSlots, 1).clearContent();

      // Borders
      sheet.getRange(startRow, 3, requiredSlots, 4).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    }

    log.push("Success");
    return { status: "success", log: log.join(" -> ") };

  } catch (e) {
    return { status: "error", message: e.toString(), log: log.join(" -> ") + " -> ERROR" };
  }
}

// =======================================================
// EXTRACTION LOGIC
// =======================================================

function buildMasterDictionary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var refSheet = ss.getSheetByName("REF_DATA");
  if (!refSheet) return {};

  var lastRow = refSheet.getLastRow();
  var dictionary = {};

  function addToDict(id, desc) {
    if (id && id !== "" && id !== "Part ID") {
      var cleanId = String(id).trim();
      if (!dictionary[cleanId]) {
        dictionary[cleanId] = String(desc).trim();
      }
    }
  }

  // 0. Base Data (Cols A & B) - Includes Vision PCs
  var baseData = refSheet.getRange(1, 1, lastRow, 2).getValues();
  for (var z = 0; z < baseData.length; z++) {
    addToDict(baseData[z][0], baseData[z][1]);
  }

  // 1. Modules (Cols C & D)
  var moduleData = refSheet.getRange(1, 3, lastRow, 2).getValues();
  for (var i = 0; i < moduleData.length; i++) {
    addToDict(moduleData[i][0], moduleData[i][1]);
  }

  // 2. Tooling Menu
  var menuData = refSheet.getRange(1, 21, lastRow, 2).getValues();
  for (var j = 0; j < menuData.length; j++) {
    var packed = String(menuData[j][1]);
    if (packed.indexOf("::") > -1) {
      var parts = packed.split("::");
      addToDict(parts[0].trim(), parts[1].trim());
    }
  }

  // 3. Manual Mapping Block
  var manualData = refSheet.getRange(1, 23, lastRow, 10).getValues();
  for (var k = 0; k < manualData.length; k++) {
    var row = manualData[k];
    for (var p = 0; p < 10; p += 2) {
      var rawId = row[p];
      var rawDesc = row[p + 1];
      if (rawId) {
        var ids = String(rawId).split(";");
        var descs = String(rawDesc).split(";");
        for (var x = 0; x < ids.length; x++) {
          var cleanId = ids[x].trim();
          var cleanDesc = (descs[x] || "").trim();
          addToDict(cleanId, cleanDesc);
        }
      }
    }
  }

  return dictionary;
}

/**
 * B. Extract Data for Production
 * UPDATED: Includes Machine Setup (Top Section) Extraction
 */
function extractProductionData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) throw new Error("Staging Sheet Missing");

  var masterDict = buildMasterDictionary();
  var lastRow = sheet.getLastRow();

  // Initialize Payload
  // Initialize Payload
  var payload = {
    PC: [], // Phase 7c.1
    CONFIG: [], // Phase 7c (Configurable Base Module -> CONFIG)
    MODULE: [],
    ELECTRICAL: [],
    VISION: [],
    TOOLING: [],
    JIG: [],
    SPARES: [],
    VCM: [],
    OTHERS: [],
    CORE: [], // Phase 7c (External BOM -> CORE)
    rowsToMarkSynced: []
  };

  // --- SECTION 0: VISION PC EXTRACTION (Cell C2) ---
  try {
    var visionPCRaw = sheet.getRange(2, 3).getValue(); // C2
    if (visionPCRaw) {
      var vId = String(visionPCRaw).trim();
      if (vId) {
        payload.PC.push({
          requiresNumbering: true,
          id: vId,
          desc: masterDict[vId] || "Vision PC (Check REF_DATA)",
          qty: 1
        });
      }
    }
  } catch (e) {
    console.warn("Vision PC Extraction Failed: " + e.message);
  }

  // --- SECTION 1: MACHINE SETUP EXTRACTION (Cols C-E) ---
  // Scan between "Base Module Tooling" and "Comment"
  // --- SECTION 1: MACHINE SETUP EXTRACTION (Cols C-E) ---
  // Scan between "Base Module Tooling" and "Comment"
  try {
    // SAFE MODE SCANNING (SINGLE READ)
    var SEARCH_DEPTH = 500;
    var actualDepth = Math.min(SEARCH_DEPTH, lastRow);
    var bigData = sheet.getRange(1, 1, actualDepth, 2).getValues(); // Read A:B

    // --- SECTION 1: CONFIGURABLE BASE MODULE EXTRACTION (Cols A-C) ---
    // Scan between "Configurable Base Module" and "Base Module Tooling"
    var configStartR = -1;
    var baseToolingStartR = -1;
    var commentStartR = -1;

    for (var i = 0; i < bigData.length; i++) {
      var rowNum = i + 1;
      var valA = String(bigData[i][0]).trim().replace(/:$/, "");
      var valB = String(bigData[i][1]).trim().replace(/:$/, "");

      if (valA === "Configurable Base Module" || valB === "Configurable Base Module") configStartR = rowNum;
      if (valA === "Base Module Tooling" || valB === "Base Module Tooling") baseToolingStartR = rowNum;
      if ((/^comments?:?$/i.test(valA) || /^comments?:?$/i.test(valB)) && baseToolingStartR !== -1 && commentStartR === -1) commentStartR = rowNum;
    }

    // A. EXTRACT CONFIGURABLE BASE MODULES -> payload.CONFIG
    if (configStartR !== -1 && baseToolingStartR !== -1 && baseToolingStartR > configStartR + 1) {
      var configRange = sheet.getRange(configStartR + 1, 3, baseToolingStartR - configStartR - 1, 1).getValues(); // Col C only
      for (var c = 0; c < configRange.length; c++) {
        var cId = String(configRange[c][0]).trim();
        if (cId) {
          payload.CONFIG.push({
            requiresNumbering: true,
            id: cId,
            desc: masterDict[cId] || "Base Module (Check REF_DATA)",
            qty: 1
          });
        }
      }
    }

    // B. EXTRACT BASE MODULE TOOLING -> payload.TOOLING
    if (baseToolingStartR !== -1 && commentStartR !== -1 && commentStartR > baseToolingStartR + 1) {
      var setupRange = sheet.getRange(baseToolingStartR + 1, 3, commentStartR - baseToolingStartR - 1, 3).getValues(); // C, D, E
      var isNewGroup = true; // Default true so orphans get numbered

      for (var m = 0; m < setupRange.length; m++) {
        var rowData = setupRange[m];
        var idC = String(rowData[0]).trim(); // Parent (Group Header)
        var idD = String(rowData[1]).trim(); // Child (Option)
        var descE = String(rowData[2]).trim(); // Description in E

        // Signal new group if Parent ID is present
        if (idC !== "") {
          isNewGroup = true;
        }

        // Logic: Only take Valid Options (Col D)
        if (idD) {
          payload.TOOLING.push({
            requiresNumbering: isNewGroup,
            id: idD,
            desc: descE || masterDict[idD] || "Tooling Option",
            qty: 1
          });
          // Consume the numbering flag for this group
          if (isNewGroup) isNewGroup = false;
        }
      }
    }
  } catch (e) {
    console.warn("Machine Setup Extraction Failed: " + e.message);
  }

  // --- SECTION 1.5: EXTERNAL CORE EXTRACTION (BOM Tree) ---
  try {
    var coreItems = fetchCoreItemsFromExternal();
    if (coreItems && coreItems.length > 0) {
      payload.CORE = coreItems;
    }
  } catch (e) {
    console.warn("CORE Extraction Failed: " + e.message);
  }



  // --- SECTION 2: MODULE CONFIGURATION EXTRACTION (B-Slots) ---
  // Dynamic Header Reading Strategy
  var vcmFinder = sheet.getRange("B:B").createTextFinder("VCM").matchEntireCell(true).findNext();
  if (!vcmFinder) throw new Error("Critical Error: 'VCM' anchor not found. Cannot extract configuration.");

  var vcmRowIndex = vcmFinder.getRow();

  // 1. Detect Header Structure (Dynamic Merge Scanning)
  // We scan columns starting from B (Col 2)
  // Logic: Current Header -> Detect Merge -> Get Range -> Next Header -> Repeat
  var columnMap = {}; // Key: ColIndex (1-based), Value: Category (e.g., "VCM")
  var colCursor = 2; // Start at B
  var MAX_SCAN_COL = 10; // Failsafe limit (Scanning B->K should be enough)

  // We expect "VCM" at Col 2. Let's see how wide it is.
  while (colCursor <= MAX_SCAN_COL) {
    var cell = sheet.getRange(vcmRowIndex, colCursor);
    var headerVal = String(cell.getValue()).trim();

    if (headerVal === "") {
      // If we hit an empty header, we stop scanning headers
      break;
    }

    var startCol = colCursor;
    var endCol = colCursor;

    if (cell.isPartOfMerge()) {
      // If merged, find the bounds
      var range = cell.getMergedRanges()[0];
      // Note: getMergedRanges returns the merge compatible with the cell
      startCol = range.getColumn();
      endCol = range.getLastColumn();
    }

    // Map these columns to the found header value
    // NOTE: If header is "Valve Set", we might want to map it to "VCM" or "OTHERS" depending on legacy logic?
    // Using EXACT header value for now, will map during Push
    for (var c = startCol; c <= endCol; c++) {
      columnMap[c] = headerVal;
    }

    // Move cursor to next block
    colCursor = endCol + 1;
  }

  // 2. Read Option Row (Immediately below Header)
  // We need to know what "Option" corresponds to each column (e.g., "Standard", "High Force", ID...)
  var optionRowIndex = vcmRowIndex + 1;
  var optionData = sheet.getRange(optionRowIndex, 1, 1, colCursor).getValues()[0]; // Read A..EndCol

  // 3. Scan Data (Strict B10 - B32 Range)
  // Find B10 Row first to start scanning
  var bSlotFinder = sheet.getRange("G:G").createTextFinder("B10").matchEntireCell(true).findNext();
  if (!bSlotFinder) {
    // Fallback: Just start searching below header if B10 explicitly missing?
    // No, strictly enforcing B10 per user request.
    console.warn("B10 Start Anchor not found. Configuring Module Extraction skipped.");
    return payload;
  }

  var startDataRow = bSlotFinder.getRow();

  // Read Data Block (From B10 downwards, plenty of rows)
  // We need to read Columns A through K (Col 11) or further if headers go further
  var maxDataCol = Math.max(18, colCursor);
  var rawData = sheet.getRange(startDataRow, 1, lastRow - startDataRow + 1, maxDataCol).getValues();

  for (var i = 0; i < rawData.length; i++) {
    var row = rawData[i];
    var turretName = String(row[6]).trim(); // Col G (Index 6)
    var syncStatus = String(row[9]).trim(); // Col J (Index 9)
    var moduleID = String(row[10]).trim();  // Col K (Index 10)

    // STOP CONDITION: B33 or higher / End of List
    if (turretName === "B33") break;
    if (turretName.startsWith("B")) {
      var num = parseInt(turretName.substring(1));
      if (!isNaN(num) && num > 32) break;
    }

    // SKIP CONDITION: Not a B-Slot (e.g. empty row or garbage)
    if (!turretName.startsWith("B")) continue;

    // SKIP CONDITION: Already Synced
    if (syncStatus === "SYNCED") continue;

    // STRICT SKIP CONDITION: Empty Module ID (Crucial Fix)
    // If no Module is assigned, we do NOT process this row, even if flags are checked.
    if (moduleID === "") continue;

    // VALID DATA FOUND -> PROCESS
    payload.rowsToMarkSynced.push({ row: startDataRow + i }); // Absolute Row

    var globalQty = parseInt(row[8]); // Col I
    if (isNaN(globalQty) || globalQty < 1) globalQty = 1;

    // Helper Push Function
    function pushItem(category, id, descOverride, isPrimary) {
      if (!id || id === "") return;
      var cleanId = id.trim();
      var finalDesc = descOverride || masterDict[cleanId] || "Check REF_DATA";

      payload[category].push({
        requiresNumbering: isPrimary,
        id: cleanId,
        desc: finalDesc,
        qty: globalQty
      });
    }

    // 1. Core Module (Col K)
    pushItem("MODULE", moduleID, String(row[7]), true); // Col H Desc

    // 2. Fixed Columns (Electrical L, Vision M, etc.) -> Indices 11, 12...
    pushItem("ELECTRICAL", String(row[11]), null, true);
    pushItem("VISION", String(row[12]), null, true);
    pushItem("TOOLING", String(row[13]), null, true);
    pushItem("TOOLING", String(row[14]), null, false);
    pushItem("JIG", String(row[15]), null, true);

    // Spares (Col 16)
    var sparesRaw = String(row[16]);
    if (sparesRaw) {
      var spareIds = sparesRaw.split(";");
      for (var s = 0; s < spareIds.length; s++) {
        pushItem("SPARES", spareIds[s].trim(), null, (s === 0));
      }
    }

    // 3. DYNAMIC COLUMNS (The Checkboxes)
    // Iterate through our mapped columns (from 'columnMap')
    // We check row[c-1] because row array is 0-indexed, columnMap is 1-based
    for (var c in columnMap) {
      var colIdx = parseInt(c); // 1-based index
      var isChecked = row[colIdx - 1]; // 0-based index

      if (isChecked === true) {
        var headerCategory = columnMap[colIdx];
        var optionValue = String(optionData[colIdx - 1]).trim(); // From Option Row

        // Parse ID/Desc from Option Cell (e.g. "430001-A689" or "Standard")
        // Attempt to extract ID if present
        var parsedID = "";
        var parsedDesc = "";

        var match = optionValue.match(/(\d{4,}-\w+)/);
        if (match) {
          parsedID = match[0];
          parsedDesc = optionValue.replace(parsedID, "").trim();
        } else {
          parsedID = optionValue;
        }

        // MAP HEADERS TO PAYLOAD SECTIONS
        var targetSection = "OTHERS"; // Default catch-all

        // Explicit Mappings
        if (headerCategory.toUpperCase().includes("VCM")) targetSection = "VCM";

        // FIXED: Valve Set now maps to OTHERS per user request
        if (headerCategory.toUpperCase().includes("VALVE")) targetSection = "OTHERS";

        pushItem(targetSection, parsedID, parsedDesc, true);
      }
    }
  }

  return payload;
}

/**
 * C. Inject Data to Production
 */
function injectProductionData(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ORDERING LIST");
  if (!sheet) throw new Error("ORDERING LIST Sheet Missing");

  var itemsAdded = 0;

  var sections = [
    { key: "PC", header: "PC" }, // Phase 7c.1
    { key: "CONFIG", header: "CONFIG" }, // Phase 7c.2
    { key: "CORE", header: "CORE" }, // Phase 7c.2
    { key: "MODULE", header: "MODULE" },
    { key: "ELECTRICAL", header: "ELECTRICAL" },
    { key: "VISION", header: "VISION" },
    { key: "VCM", header: "VCM" },
    { key: "OTHERS", header: "OTHERS" },
    { key: "SPARES", header: "SPARES" },
    { key: "JIG", header: "JIG/CALIBRATION" },
    { key: "TOOLING", header: "TOOLING" }
  ];

  for (var i = 0; i < sections.length; i++) {
    var sec = sections[i];
    var items = payload[sec.key];
    if (items && items.length > 0) {
      // Pass the entire section object 'sec' as options
      itemsAdded += insertRowsIntoSection(sheet, sec.header, items, sec);
    }
  }

  return itemsAdded;
}

/**
 * D. The Smart Fill Helper
 */
function insertRowsIntoSection(sheet, sectionHeader, items, options) {
  var lastRow = sheet.getLastRow();
  var rangeValues = sheet.getRange("A1:D" + lastRow).getValues();

  var startRowIndex = -1;
  var anchorRowIndex = -1;

  // 1. Find Start Header
  for (var i = 0; i < rangeValues.length; i++) {
    if (String(rangeValues[i][0]).trim().toUpperCase() === sectionHeader.toUpperCase()) {
      startRowIndex = i;
      break;
    }
  }

  if (startRowIndex === -1) {
    console.warn("Section Header '" + sectionHeader + "' not found in ORDERING LIST.");
    return 0;
  }

  // 2. Find Next Header (Anchor)
  for (var j = startRowIndex + 1; j < rangeValues.length; j++) {
    var cellVal = String(rangeValues[j][0]).trim();
    if (cellVal !== "") {
      anchorRowIndex = j;
      break;
    }
  }

  if (anchorRowIndex === -1) {
    anchorRowIndex = rangeValues.length;
  }

  // --- Find Current Max Item Number in this Zone ---
  var currentMaxNum = 0;
  for (var r = startRowIndex + 1; r < anchorRowIndex; r++) {
    var val = rangeValues[r][2]; // Column C
    if (typeof val === 'number') {
      if (val > currentMaxNum) currentMaxNum = val;
    } else if (val) {
      var parsed = parseInt(val);
      if (!isNaN(parsed) && parsed > currentMaxNum) {
        currentMaxNum = parsed;
      }
    }
  }

  // 3. Find Write Cursor
  var writeCursorIndex = -1;

  for (var k = startRowIndex + 1; k < anchorRowIndex; k++) {
    var partId = String(rangeValues[k][3]).trim();
    var itemNum = String(rangeValues[k][2]).trim();

    if (partId === "" && itemNum === "") {
      writeCursorIndex = k;
      break;
    }
  }

  if (writeCursorIndex === -1) {
    writeCursorIndex = anchorRowIndex;
  }

  if (writeCursorIndex === -1) {
    writeCursorIndex = anchorRowIndex;
  }

  // --- ADDITIVE LOGIC (Phase 7c): Scan for Duplicates & Mark Remarks ---
  if (writeCursorIndex > startRowIndex + 1) {
    var existingDataRange = sheet.getRange(startRowIndex + 2, 4, writeCursorIndex - (startRowIndex + 1), 4).getValues(); // Cols D, E, F, G

    // Create a Map of Existing Unreleased Items
    // Key: PartID, Value: Array of Row Indices (Relative to startRowIndex + 2)
    var unreleasedMap = {};

    for (var e = 0; e < existingDataRange.length; e++) {
      var eID = String(existingDataRange[e][0]).trim(); // Col D
      var isReleased = existingDataRange[e][3] === true; // Col G

      if (eID && !isReleased) {
        if (!unreleasedMap[eID]) unreleasedMap[eID] = [];
        unreleasedMap[eID].push(e);
      }
    }

    // Check Payload Items against Map
    for (var n = 0; n < items.length; n++) {
      var newID = items[n].id;
      if (unreleasedMap[newID]) {
        // Found unreleased duplicates! Mark them.
        var indices = unreleasedMap[newID];
        for (var k = 0; k < indices.length; k++) {
          var rowToMark = (startRowIndex + 2) + indices[k];
          sheet.getRange(rowToMark, 10).setValue("haven't release, please take note"); // Col J
        }
      }
    }
  }

  // 4. Calculate Capacity & Deficit
  var availableSlots = anchorRowIndex - writeCursorIndex;
  var itemsNeeded = items.length;
  var deficit = itemsNeeded - availableSlots;

  // 5. Expand if necessary
  if (deficit > 0) {
    // Determine Insertion Point
    var insertRowStart = anchorRowIndex + 1;
    sheet.insertRowsBefore(insertRowStart, deficit);

    // --- FIX 1: Set Background to White (Remove Grey) ---
    // We target the entire row(s) we just added (Cols 1 to Last)
    sheet.getRange(insertRowStart, 1, deficit, sheet.getLastColumn()).setBackground("white");

    // --- FIX 2: Extend Merges for Cols A & B (SMART MERGE LOGIC) ---
    // Safely handles cases where "Row Above" is already part of a multi-row merge.

    if (insertRowStart > 1) {
      try {
        var rowAbove = insertRowStart - 1;
        var lastNewRow = insertRowStart + deficit - 1;

        // Helper Logic for Extending Merge
        var extendMergeForColumn = function (colIndex) {
          var cellAbove = sheet.getRange(rowAbove, colIndex);
          var startMergeRow = rowAbove;

          // Check if "Row Above" is already merged
          if (cellAbove.isPartOfMerge()) {
            // Get the TOP of the existing block
            var existingRange = cellAbove.getMergedRanges()[0];
            startMergeRow = existingRange.getRow();
          }

          // Define the NEW total range (From Top of Block to Bottom of New Rows)
          var totalRows = lastNewRow - startMergeRow + 1;

          // Apply Merge (Force Re-merge over the expanded area)
          var targetRange = sheet.getRange(startMergeRow, colIndex, totalRows, 1);
          targetRange.merge();
          targetRange.setVerticalAlignment("middle");
        };

        // Execute for Column A (1) and Column B (2)
        extendMergeForColumn(1);
        extendMergeForColumn(2);

      } catch (e) {
        console.warn("Smart Merge Extension warning: " + e.message);
      }
    }
  }

  // 6. Write Data
  var startWriteRow = writeCursorIndex + 1;
  var output = [];
  var numberingCounter = currentMaxNum;

  for (var x = 0; x < items.length; x++) {
    var itemLabel = "";
    if (items[x].requiresNumbering) {
      numberingCounter++;
      itemLabel = numberingCounter;
    }

    output.push([
      itemLabel,    // Col C: Item No
      items[x].id,  // Col D: Part ID
      items[x].desc,// Col E: Description
      items[x].qty, // Col F: Qty
      false,        // Col G: Released (Checkbox Default False)
      "",           // Col H: Date (Empty)
      ""            // Col I: Type (Empty)
    ]);
  }

  // Write 7 columns (C through I)
  sheet.getRange(startWriteRow, 3, itemsNeeded, 7).setValues(output);

  // 7. Apply Data Validation (Checkbox & Dropdown)
  sheet.getRange(startWriteRow, 7, itemsNeeded, 1).insertCheckboxes();

  var typeRange = sheet.getRange(startWriteRow, 9, itemsNeeded, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['CHARGE OUT', 'MRP'], true)
    .setAllowInvalid(true)
    .build();
  typeRange.setDataValidation(rule);

  return itemsNeeded;
}

function markRowsAsSynced(rowsToSync) {
  if (!rowsToSync || rowsToSync.length === 0) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) return;

  for (var i = 0; i < rowsToSync.length; i++) {
    var r = rowsToSync[i].row;
    sheet.getRange(r, 10).setValue("SYNCED");
  }
}

/**
 * OPTIONAL: Retroactive Fix Tool (Option B support tool, kept for safety)
 */
function initializeReleaseColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ORDERING LIST");
  var lastRow = sheet.getLastRow();

  // Apply to G:G (Checkboxes)
  sheet.getRange(7, 7, lastRow - 6, 1).insertCheckboxes();

  // Apply to I:I (Dropdowns)
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['CHARGE OUT', 'MRP'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(7, 9, lastRow - 6, 1).setDataValidation(rule);

  SpreadsheetApp.getUi().alert("Columns Initialized.");
}

// =========================================
// EXTERNAL DATA HELPERS
// =========================================
function fetchCoreItemsFromExternal() {
  var externalID = "1nTSOqK4nGRkUEHGFnUF30gRCGFQMo6I2l8vhZB-NkSA";
  var ss = SpreadsheetApp.openById(externalID);
  var sheet = ss.getSheetByName("BOM Structure Tree Diagram");
  if (!sheet) return [];

  var finder = sheet.getRange("C:D").createTextFinder("CORE :430000-A557").matchEntireCell(false).findNext();
  if (!finder) return [];

  var startRow = finder.getRow() + 1;
  var lastRow = sheet.getLastRow();

  // Read C (ID) and D (Desc)
  var range = sheet.getRange(startRow, 3, lastRow - startRow + 1, 2).getValues();
  var coreItems = [];

  for (var i = 0; i < range.length; i++) {
    var id = String(range[i][0]).trim();
    var desc = String(range[i][1]).trim();

    // Stop at empty row
    if (id === "") break;

    coreItems.push({
      requiresNumbering: true,
      id: id,
      desc: desc,
      qty: 1
    });
  }
  return coreItems;
}
