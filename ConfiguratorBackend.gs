/**
 * ConfiguratorBackend.gs
 * Server-side logic for the Module Configurator UI.
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
  
  // Read Cols C and D (Index 3 and 4 in 1-based notation? No, getRange uses 1-based)
  // Col C = 3, Col D = 4.
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
    spareParts: sparePartsResult // Return the parsed list
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
 * 2.5 (NEW) Get Layout Header Labels
 * Reads Row 5 (Category) and Row 6 (Options) from TRIAL-LAYOUT CONFIGURATION
 * to populate the VCM/Valve UI dynamically.
 */
function getLayoutHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) return null;

  // Read Row 5 & 6, Cols B-F (Indices 2-6)
  // getRange(row, col, numRows, numCols) -> getRange(5, 2, 2, 5)
  // Row 5 = Categories, Row 6 = Options
  var rawData = sheet.getRange(5, 2, 2, 5).getValues();
  var rowCategories = rawData[0]; // Row 5
  var rowOptions = rawData[1];    // Row 6
  
  // Mapping:
  // VCM Group: Cols 2(B) & 3(C). Category Name from B5 (Index 0).
  // Valve Group: Cols 4(D), 5(E), 6(F). Category Name from D5 (Index 2).
  
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
 * Scans for first empty slot between B10 and B32.
 */
function saveConfiguration(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TRIAL-LAYOUT CONFIGURATION");
  if (!sheet) throw new Error("Sheet 'TRIAL-LAYOUT CONFIGURATION' not found.");

  // Get data for scanning: Col G (Slot IDs) and Col H (Descriptions)
  // Assuming headers are top rows, we grab mostly everything
  var lastRow = sheet.getLastRow();
  // Fetch G:H. Index 1 = Col G, Index 2 = Col H in the resulting 2D array
  var scanRange = sheet.getRange(1, 7, lastRow, 2).getValues();

  var targetRowIndex = -1;
  var foundSlotLabel = "";
  var startScanning = false;

  for (var i = 0; i < scanRange.length; i++) {
    var slotID = String(scanRange[i][0]).trim(); // Col G
    var desc = String(scanRange[i][1]).trim();   // Col H

    // START Condition
    if (slotID === "B10") {
      startScanning = true;
    }

    // CHECK Condition (Only if started)
    if (startScanning) {
      // STOP Condition
      if (slotID === "B33" || (slotID.indexOf("B") === 0 && parseInt(slotID.substring(1)) > 32)) {
         break; // Went past B32
      }

      // EMPTY DETECT
      // UPDATED: Now strictly checks for empty strings only. 
      // Removed the check for "---" to ensure rows marked with dashes are skipped.
      if (desc === "") {
        targetRowIndex = i + 1; // 1-based index
        foundSlotLabel = slotID;
        break; // Found it!
      }
    }
  }

  if (targetRowIndex === -1) {
    throw new Error("Configuration List (B10-B32) is full. Please clear some rows.");
  }

  // --- WRITE DATA ---
  // Col H (8)  : Description
  // Col K (11) : Module ID
  // Col L (12) : Elec ID
  // Col M (13) : Vision ID
  // Col N (14) : Tool Option ID
  // Col O (15) : Rubber Tip ID
  // Col P (16) : Jig ID
  // Col Q (17) : Spare Part ID

  sheet.getRange(targetRowIndex, 8).setValue(payload.moduleDesc);
  sheet.getRange(targetRowIndex, 11).setValue(payload.moduleID);
  sheet.getRange(targetRowIndex, 12).setValue(payload.elecID);
  sheet.getRange(targetRowIndex, 13).setValue(payload.visionID);
  sheet.getRange(targetRowIndex, 14).setValue(payload.toolOptionID);
  sheet.getRange(targetRowIndex, 15).setValue(payload.rubberTipID);
  sheet.getRange(targetRowIndex, 16).setValue(payload.jigID);
  sheet.getRange(targetRowIndex, 17).setValue(payload.sparePartsID);
  
  // (NEW) Write VCM/Valve Flags (Cols 2-6)
  // Payload.layoutFlags = { "2": true, "3": false ... }
  if (payload.layoutFlags) {
     for (var col = 2; col <= 6; col++) {
        var val = payload.layoutFlags[String(col)] === true; // Ensure boolean
        var cell = sheet.getRange(targetRowIndex, col);
        
        // CHECKBOX LOGIC:
        // 1. Check if the cell already has Checkbox validation.
        // 2. If not (it's empty or text), insert checkboxes first.
        // 3. Set the Boolean value (which will check/uncheck the box).
        
        var rule = cell.getDataValidation();
        var isCheckbox = (rule != null && rule.getCriteriaType() == SpreadsheetApp.DataValidationCriteria.CHECKBOX);
        
        if (!isCheckbox) {
           cell.insertCheckboxes();
        }
        
        cell.setValue(val);
     }
  }

  return { status: "success", slot: foundSlotLabel };
}
