/**
 * ConfiguratorBackend.gs
 * Server-side logic for the Module Configurator UI.
 * UPDATED: Phase 4.4 (Smart Fill Logic)
 */

// --- CONSTANTS ---
var RUBBER_TIP_PARENTS_BACKEND = ["430001-A689", "430001-A690", "430001-A691", "430001-A692"];
var RUBBER_TIP_SOURCE_ID_BACKEND = "430001-A380";

/**
 * 1. Get List of Modules (ID and Description)
 * Reads REF_DATA Cols C (ID) and D (Description).
 */
function getModuleList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  
  // Read Cols C and D
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
 * 2. Get Full Details for Selected Module
 * Fetches IDs and Descriptions for all components.
 */
function getModuleDetails(moduleID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var refSheet = ss.getSheetByName("REF_DATA");
  
  // Fetch C:AF range (Extended to include Spare Parts)
  // C is Col 3, AF is Col 32. Total cols = 30.
  var dataRange = refSheet.getRange("C:AF").getValues();
  var rowData = null;
  
  for (var i = 0; i < dataRange.length; i++) {
    if (String(dataRange[i][0]) === moduleID) {
      rowData = dataRange[i];
      break;
    }
  }
  
  if (!rowData) return { error: "Module ID not found." };
  
  // Array Indices (0-based from C):
  // W=20 (Elec ID), X=21 (Elec Desc)
  // Y=22 (Tool ID), Z=23 (Tool Desc)
  // AA=24 (Jig ID), AB=25 (Jig Desc)
  // AC=26 (Vision ID), AD=27 (Vision Desc)
  // AE=28 (Spare ID), AF=29 (Spare Desc) -> NEW
  
  // --- ELECTRICAL ---
  var elecResult = null;
  var elecIdsStr = rowData[20];
  var elecDescStr = rowData[21];
  
  if (elecIdsStr) {
    var eIds = String(elecIdsStr).split(';').map(function(s){ return s.trim(); });
    var eDescs = String(elecDescStr).split(';').map(function(s){ return s.trim(); });
    
    // Default to Instance 1
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
    var tIds = String(toolIdsStr).split(';').map(function(s){ return s.trim(); });
    var tDescs = String(toolDescStr).split(';').map(function(s){ return s.trim(); });
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
      
      // Standard Options
      var stdOpts = fetchOptionsForTool(menuData, tID);
      if (stdOpts.length > 0) toolObj.standardOptions = stdOpts;
      
      // Rubber Tip Dependency
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
    var vIds = String(visionIdsStr).split(';').map(function(s){ return s.trim(); });
    var vDescs = String(visionDescStr).split(';').map(function(s){ return s.trim(); });
    
    // Create objects for each vision option
    var vOptions = [];
    for(var v=0; v<vIds.length; v++){
      if(vIds[v]) {
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
    var jIds = String(jigIdsStr).split(';').map(function(s){ return s.trim(); });
    var jDescs = String(jigDescStr).split(';').map(function(s){ return s.trim(); });
    for (var j = 0; j < jIds.length; j++) {
      if(jIds[j]) jigsResult.push({ id: jIds[j], desc: jDescs[j] || "" });
    }
  }

  // --- SPARE PARTS (NEW) ---
  var sparePartsResult = [];
  var spIdsStr = rowData[28]; // Col AE
  var spDescStr = rowData[29]; // Col AF
  
  if (spIdsStr) {
    var spIds = String(spIdsStr).split(';').map(function(s){ return s.trim(); });
    var spDescs = String(spDescStr).split(';').map(function(s){ return s.trim(); });
    
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
    var rowParent = String(menuData[i][0]);
    var rowChild = String(menuData[i][1]); 

    if (rowParent === parentID) {
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
      if (rowParent !== "") break; 
    }
  }
  return options;
}

/**
 * 2.5 Get Layout Header Labels
 */
function getLayoutHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) return null;

  var rawData = sheet.getRange(5, 2, 2, 5).getValues();
  var rowCategories = rawData[0]; // Row 5
  var rowOptions = rawData[1];    // Row 6
  
  return {
    group1: {
      title: rowCategories[0], // B5 "VCM"
      options: [
        { colIndex: 2, label: rowOptions[0] }, // B6
        { colIndex: 3, label: rowOptions[1] }  // C6
      ]
    },
    group2: {
      title: rowCategories[2], // D5 "VALVE SET"
      options: [
        { colIndex: 4, label: rowOptions[2] }, // D6
        { colIndex: 5, label: rowOptions[3] }, // E6
        { colIndex: 6, label: rowOptions[4] }  // F6
      ]
    }
  };
}

/**
 * 3. SAVE CONFIGURATION TO TRIAL LAYOUT
 */
function saveConfiguration(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) throw new Error("Sheet 'TRIAL-LAYOUT CONFIGURATION' not found.");

  var lastRow = sheet.getLastRow();
  var scanRange = sheet.getRange(1, 7, lastRow, 2).getValues();

  var targetRowIndex = -1;
  var foundSlotLabel = "";
  var startScanning = false;

  for (var i = 0; i < scanRange.length; i++) {
    var slotID = String(scanRange[i][0]).trim();
    var desc = String(scanRange[i][1]).trim();

    if (slotID === "B10") {
      startScanning = true;
    }

    if (startScanning) {
      if (slotID === "B33" || (slotID.indexOf("B") === 0 && parseInt(slotID.substring(1)) > 32)) {
         break; 
      }

      if (desc === "") {
        targetRowIndex = i + 1; 
        foundSlotLabel = slotID;
        break; 
      }
    }
  }

  if (targetRowIndex === -1) {
    throw new Error("Configuration List (B10-B32) is full. Please clear some rows.");
  }

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

  return { status: "success", slot: foundSlotLabel };
}

// =======================================================
// PHASE 4.2: EXTRACTION LOGIC ("The Brain")
// =======================================================

/**
 * A. Build Master Dictionary
 */
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

  // 1. Modules
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
      addToDict(parts[0], parts[1]);
    }
  }

  // 3. Manual Mapping Block
  var manualData = refSheet.getRange(1, 23, lastRow, 10).getValues(); 
  for (var k = 0; k < manualData.length; k++) {
    var row = manualData[k];
    for (var p = 0; p < 10; p += 2) {
       var rawId = row[p];
       var rawDesc = row[p+1];
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
 */
function extractProductionData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) throw new Error("Staging Sheet Missing");

  var masterDict = buildMasterDictionary();
  var lastRow = sheet.getLastRow();
  
  var headerRow = sheet.getRange(6, 1, 1, 18).getValues()[0];
  var rawData = sheet.getRange(1, 1, lastRow, 18).getValues();

  var payload = {
    MODULE: [],
    ELECTRICAL: [],
    VISION: [],
    TOOLING: [],
    JIG: [], // "JIG/CALIBRATION"
    SPARES: [],
    VCM: [],
    OTHERS: [], 
    rowsToMarkSynced: [] 
  };

  var startScanning = false;
  
  for (var i = 0; i < rawData.length; i++) {
    var row = rawData[i];
    var turretName = String(row[6]).trim(); // Col G
    var syncStatus = String(row[9]).trim(); // Col J
    var moduleID = String(row[10]).trim();  // Col K
    
    if (turretName === "B10") startScanning = true;
    
    if (turretName === "B33" || (turretName.startsWith("B") && parseInt(turretName.substring(1)) > 32)) {
      break;
    }

    if (startScanning) {
      if (moduleID === "") continue; 
      if (syncStatus === "SYNCED") continue; 

      payload.rowsToMarkSynced.push({row: i + 1});
      
      var globalQty = parseInt(row[8]); 
      if (isNaN(globalQty) || globalQty < 1) globalQty = 1;
      
      var itemIndex = payload.rowsToMarkSynced.length; 

      function pushItem(category, id, descOverride, isPrimary) {
        if (!id || id === "") return;
        var cleanId = id.trim();
        var finalDesc = descOverride || masterDict[cleanId] || "Check REF_DATA";
        
        payload[category].push({
          itemIdx: isPrimary ? itemIndex : "", 
          id: cleanId,
          desc: finalDesc,
          qty: globalQty 
        });
      }

      pushItem("MODULE", moduleID, String(row[7]), true);
      pushItem("ELECTRICAL", String(row[11]), null, true);
      pushItem("VISION", String(row[12]), null, true);

      // Tooling
      pushItem("TOOLING", String(row[13]), null, true); 
      pushItem("TOOLING", String(row[14]), null, false); 

      pushItem("JIG", String(row[15]), null, true);

      // Spares
      var sparesRaw = String(row[16]);
      if (sparesRaw) {
        var spareIds = sparesRaw.split(";");
        for (var s = 0; s < spareIds.length; s++) {
          var sId = spareIds[s].trim();
          pushItem("SPARES", sId, null, (s === 0)); 
        }
      }

      function extractIdFromHeader(headerText) {
        var match = headerText.match(/(\d{4,}-\w+)/); 
        return match ? match[0] : "";
      }

      if (row[1] === true) pushItem("VCM", extractIdFromHeader(headerRow[1]), null, true);
      if (row[2] === true) pushItem("VCM", extractIdFromHeader(headerRow[2]), null, true);

      if (row[3] === true) pushItem("OTHERS", extractIdFromHeader(headerRow[3]), null, true);
      if (row[4] === true) pushItem("OTHERS", extractIdFromHeader(headerRow[4]), null, true);
      if (row[5] === true) pushItem("OTHERS", extractIdFromHeader(headerRow[5]), null, true);

    } 
  } 

  return payload;
}

// =======================================================
// PHASE 4.3: PRODUCTION INJECTION & STATUS UPDATE
// =======================================================

/**
 * C. Main Injection Controller
 */
function injectProductionData(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ORDERING LIST");
  if (!sheet) throw new Error("ORDERING LIST Sheet Missing");

  var itemsAdded = 0;

  // Map Payload Key to Sheet Header Name
  var sections = [
    { key: "MODULE", header: "MODULE" },
    { key: "ELECTRICAL", header: "ELECTRICAL" },
    { key: "VISION", header: "VISION" },
    { key: "VCM", header: "VCM" },
    { key: "OTHERS", header: "OTHERS" },
    { key: "SPARES", header: "SPARES" },
    { key: "JIG", header: "JIG/CALIBRATION" },
    { key: "TOOLING", header: "TOOLING" } // Will default to OTHERS if missing, or alert
  ];

  for (var i = 0; i < sections.length; i++) {
    var sec = sections[i];
    var items = payload[sec.key];
    if (items && items.length > 0) {
      itemsAdded += insertRowsIntoSection(sheet, sec.header, items);
    }
  }

  return itemsAdded;
}

/**
 * D. The Smart Fill Helper (Fill then Expand) - UPDATED PHASE 4.4
 */
function insertRowsIntoSection(sheet, sectionHeader, items) {
  var lastRow = sheet.getLastRow();
  // Fetch columns A (Headers) and D (Part ID) to analyze structure
  // We grab A1:D(LastRow) to have context of headers and content
  var rangeValues = sheet.getRange("A1:D" + lastRow).getValues();

  var startRowIndex = -1; // 0-based index of the Header row
  var anchorRowIndex = -1; // 0-based index of the Next Header row

  // 1. Find Start Header
  for (var i = 0; i < rangeValues.length; i++) {
    if (String(rangeValues[i][0]).trim().toUpperCase() === sectionHeader.toUpperCase()) {
      startRowIndex = i;
      break;
    }
  }

  if (startRowIndex === -1) {
    console.warn("Section Header '" + sectionHeader + "' not found in ORDERING LIST.");
    return 0; // Skip this section
  }

  // 2. Find Next Header (Anchor)
  // Scan downwards from the row AFTER startRowIndex.
  // The logic: The next "Header" is the first non-empty cell in Column A.
  for (var j = startRowIndex + 1; j < rangeValues.length; j++) {
    var cellVal = String(rangeValues[j][0]).trim(); // Col A
    if (cellVal !== "") {
       anchorRowIndex = j;
       break;
    }
  }

  // If no next header found, the Anchor is effectively "after the last row"
  if (anchorRowIndex === -1) {
     anchorRowIndex = rangeValues.length; // Points to the row AFTER the last data row
  }

  // 3. Find Write Cursor (First Empty Row in Zone)
  // Zone is between (startRowIndex + 1) and (anchorRowIndex - 1)
  var writeCursorIndex = -1;

  for (var k = startRowIndex + 1; k < anchorRowIndex; k++) {
    // Check Column D (Index 3) for content (Part ID)
    // Also check Col C (Index 2) just in case D is empty but C has Item #
    var partId = String(rangeValues[k][3]).trim();
    var itemNum = String(rangeValues[k][2]).trim();

    if (partId === "" && itemNum === "") {
      writeCursorIndex = k;
      break;
    }
  }

  // If section is full (no empty rows found), the cursor is at the anchor
  if (writeCursorIndex === -1) {
    writeCursorIndex = anchorRowIndex;
  }

  // 4. Calculate Capacity & Deficit
  // writeCursorIndex is 0-based index.
  // anchorRowIndex is 0-based index.
  // Available rows = anchorRowIndex - writeCursorIndex
  var availableSlots = anchorRowIndex - writeCursorIndex;
  var itemsNeeded = items.length;
  var deficit = itemsNeeded - availableSlots;

  // 5. Expand if necessary (Scenario B)
  if (deficit > 0) {
    // We insert rows BEFORE the anchor row.
    // anchorRowIndex is 0-based, so (anchorRowIndex + 1) is the 1-based row number.
    // This pushes the Anchor down, creating 'deficit' amount of new blank rows.
    sheet.insertRowsBefore(anchorRowIndex + 1, deficit);
  }

  // 6. Write Data
  // We write starting at writeCursorIndex.
  // Converting to 1-based: writeCursorIndex + 1
  var startWriteRow = writeCursorIndex + 1;

  var output = [];
  for (var x = 0; x < items.length; x++) {
    output.push([
      items[x].itemIdx, // Col C
      items[x].id,      // Col D
      items[x].desc,    // Col E
      items[x].qty      // Col F
    ]);
  }

  // Range: Row, Col 3 (C), Height, Width 4
  sheet.getRange(startWriteRow, 3, itemsNeeded, 4).setValues(output);

  return itemsNeeded;
}

/**
 * E. Update Status in Staging
 */
function markRowsAsSynced(rowsToSync) {
  if (!rowsToSync || rowsToSync.length === 0) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) return;

  for (var i = 0; i < rowsToSync.length; i++) {
     var r = rowsToSync[i].row;
     // Set Column J (10)
     sheet.getRange(r, 10).setValue("SYNCED");
  }
}
