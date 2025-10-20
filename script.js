// Global state
let workbook = null;
let dataHeaders = [];
let rawData = [];
let isProcessing = false;

// --- Firestore Global Variables Setup (Mandatory) ---
// Note: These are placeholder vars as you included them in the original code
const appId = typeof __app_id !== "undefined" ? __app_id : "default-app-id";
const firebaseConfig =
  typeof __firebase_config !== "undefined" ? JSON.parse(__firebase_config) : {};
const initialAuthToken =
  typeof __initial_auth_token !== "undefined" ? __initial_auth_token : null;
let db = null;
let auth = null;

// --- Utility Functions ---

/** Shows a message box (toast replacement) */
function showMessage(text, type = "info") {
  const box = $("#messageBox");
  let classes = "bg-box text-white border-subtle";
  if (type === "success") {
    // Success uses primary color for strong indication
    classes = "bg-primary text-white border-primary";
  } else if (type === "error") {
    // Error uses white text on dark box with subtle border
    classes = "bg-box text-white border-subtle";
  }
  box
    .removeClass()
    .addClass("p-3 rounded-lg text-sm mb-4 border " + classes)
    .text(text)
    .slideDown(200);
  setTimeout(() => box.slideUp(200), 5000);
}

/** Converts cell value based on detected type or selected type */
function convertValue(value, type) {
  if (value === null || value === undefined) return null;
  if (String(value).trim() === "" && type !== "string") return null;

  switch (type) {
    case "number":
      const num = parseFloat(value);
      return isNaN(num) ? String(value) : num;
    case "date":
      if (typeof value === "number" && value > 1) {
        try {
          // Excel date to Date object
          const date = XLSX.SSF.parse_date_code(value);
          const dateObj = new Date(Date.UTC(date.y, date.m - 1, date.d));
          return dateObj.toISOString().split("T")[0];
        } catch (e) {
          return String(value);
        }
      }
      const dateVal = new Date(value);
      return isNaN(dateVal.getTime())
        ? String(value)
        : dateVal.toISOString().split("T")[0];
    case "boolean":
      const str = String(value).toLowerCase().trim();
      if (str === "true" || str === "1" || str === "yes") return true;
      if (str === "false" || str === "0" || str === "no") return false;
      return value;
    case "string":
      return String(value).trim();
    case "auto":
    default:
      if (
        value === "TRUE" ||
        value === "FALSE" ||
        value === "true" ||
        value === "false"
      )
        return String(value).toLowerCase() === "true";
      if (!isNaN(parseFloat(value)) && isFinite(value))
        return parseFloat(value);
      return String(value);
  }
}

/** Core function to process one data record (row) against the mapping and build nested JSON. */
function processRecord(record, mapping, customKeys) {
  const result = {};
  const arrayMap = new Map();

  // 1. Process Column Mappings
  for (let i = 0; i < mapping.length; i++) {
    const map = mapping[i];
    const originalKey = map.originalKey;
    const targetPath = map.targetPath.trim();
    const dataType = map.dataType;
    const transformRule = map.transformRule;

    if (map.action === "skip" || !targetPath) continue;

    let value = record[originalKey];

    // --- Custom Value Transformation Logic ---
    if (transformRule && value !== null && value !== undefined) {
      let stringValue = String(value).trim();

      // Rule format: FIND1=REPLACE1|FIND2=REPLACE2
      const rules = transformRule.split("|");

      for (const rule of rules) {
        const parts = rule.trim().split("=");
        if (parts.length === 2) {
          const find = parts[0].trim();
          const replace = parts[1].trim();

          // Case-insensitive comparison for finding the value
          if (stringValue.toUpperCase() === find.toUpperCase()) {
            value = replace; // Apply replacement
            break; // Stop after the first match
          }
        }
      }
    }
    // --- End Custom Value Transformation Logic ---

    value = convertValue(value, dataType); // Convert after potential transformation

    if (value === null && $("#skipEmptyRows").is(":checked")) continue;

    const isArrayPath =
      targetPath.includes("[]") || targetPath.match(/\[\d+\]/);

    if (isArrayPath) {
      // --- Array Grouping Logic (e.g., items[].id) ---

      const parts = targetPath.split(".");
      let arrayPath = [];
      let arrayRootName = "";
      let keyPathInArray = "";

      let foundArrayPart = false;
      for (let part of parts) {
        const match = part.match(/(\w+)\[(\d*|\])/);
        if (match) {
          arrayRootName = match[1];
          keyPathInArray = targetPath.substring(
            targetPath.indexOf(part) + part.length + 1
          );
          foundArrayPart = true;
          break;
        }
        arrayPath.push(part);
      }

      if (!foundArrayPart) continue;

      let arrayParent = result;
      if (arrayPath.length > 0) {
        arrayParent = arrayPath.reduce((acc, part) => {
          acc[part] = acc[part] || {};
          return acc[part];
        }, result);
      }

      arrayParent[arrayRootName] = arrayParent[arrayRootName] || [];
      const array = arrayParent[arrayRootName];

      let currentArrayObject = arrayMap.get(arrayRootName);
      if (!currentArrayObject) {
        currentArrayObject = {};
        arrayMap.set(arrayRootName, currentArrayObject);
        array.push(currentArrayObject);
      }

      let currentNested = currentArrayObject;
      const nestedParts = keyPathInArray.split(".");
      for (let k = 0; k < nestedParts.length; k++) {
        const part = nestedParts[k];
        const isLast = k === nestedParts.length - 1;

        if (!isLast) {
          currentNested[part] = currentNested[part] || {};
          currentNested = currentNested[part];
        } else {
          currentNested[part] = value;
        }
      }
    } else {
      // --- Standard Object Nesting Logic (e.g., user.name) ---
      let current = result;
      const pathParts = targetPath.split(".");

      for (let j = 0; j < pathParts.length; j++) {
        let part = pathParts[j];
        const isLast = j === pathParts.length - 1;

        if (!isLast) {
          if (
            !current[part] ||
            typeof current[part] !== "object" ||
            Array.isArray(current[part])
          ) {
            current[part] = {};
          }
          current = current[part];
        } else {
          current[part] = value;
        }
      }
    }
  }

  // 2. Process Custom Keys
  customKeys.forEach((custom) => {
    if (custom.key && custom.value !== undefined) {
      let current = result;
      const parts = custom.key.split(".");
      let value = convertValue(custom.value, custom.type);

      for (let i = 0; i < parts.length; i++) {
        const part = parts[i];
        const isLast = i === parts.length - 1;

        if (!isLast) {
          current[part] = current[part] || {};
          current = current[part];
        } else {
          current[part] = value;
        }
      }
    }
  });

  return result;
}

// --- UI/Handler Functions ---

/** Handles file selection and reading */
function handleFile(file) {
  if (isProcessing) return;
  isProcessing = true;
  showMessage(`Reading file: ${file.name}...`, "info");
  $("#fileInfo")
    .text(`File: ${file.name} | Size: ${(file.size / 1024).toFixed(2)} KB`)
    .removeClass("hidden");
  $("#convertBtn").prop("disabled", true).text("Reading...");
  $("#savePresetBtn").prop("disabled", true);
  $("#progressBarContainer").removeClass("hidden");
  $("#progressBar").css("width", "10%");

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = e.target.result;
    try {
      workbook = XLSX.read(data, { type: "binary", cellDates: true });
      populateSheetSelect(workbook);
      isProcessing = false;
      $("#convertBtn").prop("disabled", false).text("Generate JSON");
      $("#savePresetBtn").prop("disabled", false);
      $("#progressBarContainer").addClass("hidden");
      showMessage("File loaded successfully. Please map columns.", "success");
      loadSheetData();
    } catch (error) {
      showMessage(`Error reading file: ${error.message}`, "error");
      isProcessing = false;
      $("#convertBtn").prop("disabled", true).text("Generate JSON");
      $("#savePresetBtn").prop("disabled", true);
      $("#progressBarContainer").addClass("hidden");
    }
  };
  reader.readAsBinaryString(file);
}

/** Populates the sheet dropdown */
function populateSheetSelect(wb) {
  const select = $("#sheetSelect");
  select.empty().prop("disabled", false);
  wb.SheetNames.forEach((name) => {
    select.append(`<option value="${name}" class="bg-box">${name}</option>`);
  });
  select.off("change").on("change", loadSheetData);
}

/** Reads data from the selected sheet and updates previews/mappings */
/** Reads data from the selected sheet and updates previews/mappings */
function loadSheetData() {
  if (!workbook) return;

  const sheetName = $("#sheetSelect").val();
  const ws = workbook.Sheets[sheetName];
  if (!ws) return;

  // Read data with headers as JSON keys
  const sheetData = XLSX.utils.sheet_to_json(ws, {
    raw: false,
    defval: null,
    header: 1, // Read as array of arrays first
  });

  // Use first row as headers, handle empty/null headers
  let headers = [];
  if (sheetData.length > 0) {
    headers = sheetData[0].map((h) =>
      h !== null && h !== undefined ? String(h).trim() : "Unnamed_Column"
    );
    // Ensure unique headers by appending index if duplicate found
    const uniqueHeaders = [];
    const headerCounts = {};
    headers.forEach((h) => {
      if (headerCounts[h]) {
        headerCounts[h]++;
        uniqueHeaders.push(`${h}_${headerCounts[h]}`);
      } else {
        headerCounts[h] = 1;
        uniqueHeaders.push(h);
      }
    });
    headers = uniqueHeaders;
  }

  // Convert array of arrays to array of objects using unique headers
  rawData = [];
  if (sheetData.length > 1) {
    for (let i = 1; i < sheetData.length; i++) {
      const row = {};
      headers.forEach((h, index) => {
        row[h] = sheetData[i][index] !== undefined ? sheetData[i][index] : null;
      });
      rawData.push(row);
    }
  }

  dataHeaders = headers;

  // ✅ Preview aur Mapping update
  populateSheetPreview(rawData, dataHeaders);
  populateColumnMapping(dataHeaders);

  // ✅ Auto-expand Sheet Preview panel jab data load ho jae
  const previewPanel = $("#sheetPreviewPanel .panel-content");
  if (rawData.length > 0 && previewPanel.is(":hidden")) {
    previewPanel.slideDown(300);
    $("#sheetPreviewPanel .toggle-btn svg")
      .removeClass("rotate-0")
      .addClass("rotate-180")
      .attr("data-state", "expanded");
  }
}

/** Populates the Sheet Preview table */
function populateSheetPreview(data, headers) {
  const table = $("#sheetPreviewTable");
  const thead = table.find("thead");
  const tbody = table.find("tbody");

  thead.empty();
  tbody.empty();

  if (headers.length === 0) {
    tbody.html(
      `<tr><td colspan="100" class="p-4 text-center text-white opacity-50">No data found in this sheet.</td></tr>`
    );
    return;
  }

  // Header Row
  let headerRowHtml =
    '<tr class="bg-box border-b border-subtle bg-[#141414] sticky top-0">';
  headers.forEach((h) => {
    headerRowHtml += `<th class="px-4 py-3 text-left text-sm font-semibold text-white uppercase tracking-wider">${h}</th>`;
  });
  headerRowHtml += "</tr>";
  thead.html(headerRowHtml);

  // Data Rows
  data.forEach((row) => {
    let rowHtml =
      '<tr class="border-b border-subtle hover:bg-box/50 transition duration-150">';
    headers.forEach((h) => {
      const value = row[h] !== null && row[h] !== undefined ? row[h] : "";
      rowHtml += `<td class="px-4 py-3 whitespace-nowrap text-sm text-white/70">${
        value !== "" ? value : "<em>-</em>"
      }</td>`;
    });
    rowHtml += "</tr>";
    tbody.append(rowHtml);
  });
}

/** Populates the Column Mapping table */
function populateColumnMapping(headers) {
  const tbody = $("#mappingBody");
  tbody.empty();

  if (headers.length === 0) {
    tbody.html(
      `<tr><td colspan="5" class="p-4 text-center text-white opacity-50">No headers found to map.</td></tr>`
    );
    return;
  }

  const typeOptions = [
    { val: "auto", text: "Auto" },
    { val: "string", text: "String" },
    { val: "number", text: "Number" },
    { val: "date", text: "Date (YYYY-MM-DD)" },
  ];

  const actionOptions = [
    { val: "keep", text: "Keep" },
    { val: "skip", text: "Skip" },
  ];

  headers.forEach((h, index) => {
    // Generate default key: clean, snake_case, unique
    const defaultKey = h
      .replace(/[^a-zA-Z0-9]/g, "_")
      .toLowerCase()
      .replace(/_+/g, "_")
      .replace(/^_|_$/g, "");
    const initialPath = defaultKey || `column_${index + 1}`;

    const $row = $(
      `<tr id="map-row-${index}" class="border-b border-subtle hover:bg-box/50 transition duration-150"></tr>`
    );

    // Original Header
    $row.append(`<td class="p-3 font-mono text-sm text-white">${h}</td>`);

    // Mapped JSON Key
    $row.append(`
        <td class="p-3">
            <input type="text" data-original="${h}" value="${initialPath}"
                    class="w-full input-style text-sm p-1.5 focus:ring-primary focus:border-primary transition duration-200 mapping-target-path"
                    placeholder="e.g., user.name or items[].id">
        </td>
    `);

    // Data Type
    let typeSelect = `<select class="w-28 input-style text-sm p-1.5 focus:ring-primary focus:border-primary transition duration-200 mapping-data-type">`;
    typeOptions.forEach((opt) => {
      const selected = opt.val === "auto" ? "selected" : "";
      typeSelect += `<option value="${opt.val}" ${selected} class="bg-box">${opt.text}</option>`;
    });
    typeSelect += `</select>`;
    $row.append(`<td class="p-3">${typeSelect}</td>`);

    // NEW: Transform Rule
    $row.append(`
        <td class="p-3">
            <input type="text" 
                    class="w-full input-style text-sm p-1.5 focus:ring-primary focus:border-primary transition duration-200 mapping-transform-rule"
                    placeholder="e.g., PO=Purchase Order|SO=Sales Order">
        </td>
    `);

    // Action (Keep/Skip)
    let actionSelect = `<select class="w-24 input-style text-sm p-1.5 focus:ring-primary focus:border-primary transition duration-200 mapping-action">`;
    actionOptions.forEach((opt) => {
      const selected = opt.val === "keep" ? "selected" : "";
      actionSelect += `<option value="${opt.val}" ${selected} class="bg-box">${opt.text}</option>`;
    });
    actionSelect += `</select>`;
    $row.append(`<td class="p-3">${actionSelect}</td>`);

    tbody.append($row);
  });
}

/** Extracts the current column mapping from the UI */
function getMapping() {
  const mapping = [];
  $("#mappingBody tr").each(function () {
    const $row = $(this);
    const originalKey = $row.find(".mapping-target-path").data("original");
    const targetPath = $row.find(".mapping-target-path").val();
    const dataType = $row.find(".mapping-data-type").val();
    const action = $row.find(".mapping-action").val();
    const transformRule = $row.find(".mapping-transform-rule").val();

    if (originalKey) {
      mapping.push({
        originalKey: originalKey,
        targetPath: targetPath,
        dataType: dataType,
        action: action,
        transformRule: transformRule,
      });
    }
  });
  return mapping;
}

/** Extracts custom keys from the UI */
function getCustomKeys() {
  const customKeys = [];
  $("#customKeysList .custom-key-row").each(function () {
    const $row = $(this);
    const key = $row.find(".custom-key-input").val();
    const value = $row.find(".custom-value-input").val();
    const type = $row.find(".custom-type-select").val();

    if (key && value !== "") {
      customKeys.push({ key, value, type });
    }
  });
  return customKeys;
}

/** Main conversion logic */
$("#convertBtn").on("click", function () {
  if (isProcessing || rawData.length === 0) {
    showMessage("Please upload a file first.", "error");
    return;
  }

  isProcessing = true;
  $("#convertBtn").prop("disabled", true).text("Converting...");
  $("#downloadJsonBtn").prop("disabled", true);
  $("#progressBarContainer").removeClass("hidden");

  const mapping = getMapping();
  const customKeys = getCustomKeys();
  const wrapArray = $("#wrapArray").is(":checked");
  const prettyPrint = $("#prettyPrint").is(":checked");
  const jsonOutputElement = $("#jsonOutput");
  let finalJson = [];

  // Simulate processing time with progress bar update
  const totalRows = rawData.length;
  let processedRows = 0;

  setTimeout(() => {
    try {
      rawData.forEach((record, index) => {
        const result = processRecord(record, mapping, customKeys);
        if (Object.keys(result).length > 0) {
          finalJson.push(result);
        }
        processedRows++;
        const progress = Math.min(
          100,
          Math.floor((processedRows / totalRows) * 90) + 10
        );
        $("#progressBar").css("width", `${progress}%`);
      });

      let outputData = wrapArray ? finalJson : finalJson[0] || {};
      const jsonString = JSON.stringify(outputData, null, prettyPrint ? 2 : 0);
      jsonOutputElement.text(jsonString);

      showMessage(
        `Conversion complete! ${finalJson.length} records processed.`,
        "success"
      );

      // ✅ JSON panel auto expand hone ka code
      const jsonPanel = $("#jsonOutputPanel .panel-content");
      if (jsonPanel.is(":hidden")) {
        jsonPanel.slideDown(300);
        $("#jsonOutputPanel .toggle-btn svg")
          .removeClass("rotate-0")
          .addClass("rotate-180")
          .attr("data-state", "expanded");
      }

      $("#downloadJsonBtn").prop("disabled", false);
    } catch (error) {
      showMessage(`Conversion Error: ${error.message}`, "error");
      jsonOutputElement.text(`Error: ${error.message}`);
    } finally {
      isProcessing = false;
      $("#convertBtn").prop("disabled", false).text("Generate JSON");
      $("#progressBar").css("width", "100%");
      setTimeout(() => $("#progressBarContainer").addClass("hidden"), 500);
    }
  }, 50);
});

/** Add Custom Key Row */
let customKeyCounter = 0;
function addCustomKeyRow(key = "", value = "", type = "string") {
  const rowId = `custom-key-${customKeyCounter++}`;
  const typeOptions = [
    { val: "string", text: "String" },
    { val: "number", text: "Number" },
    { val: "boolean", text: "Boolean (true/false)" },
  ];

  let typeSelect = `<select class="custom-type-select w-28 input-style text-sm p-1.5 focus:ring-primary focus:border-primary transition duration-200">`;
  typeOptions.forEach((opt) => {
    const selected = opt.val === type ? "selected" : "";
    typeSelect += `<option value="${opt.val}" ${selected} class="bg-box">${opt.text}</option>`;
  });
  typeSelect += `</select>`;

  const $row = $(`
        <div id="${rowId}" class="custom-key-row flex space-x-2 items-center">
            <input type="text" value="${key}" placeholder="JSON Key (e.g., meta.version)"
                class="custom-key-input flex-1 input-style text-sm p-1.5">
            <input type="text" value="${value}" placeholder="Static Value"
                class="custom-value-input flex-1 input-style text-sm p-1.5">
            ${typeSelect}
            <button class="remove-key-btn text-white/50 hover:text-red-500 transition duration-200" title="Remove">
                <i class="ti ti-trash w-5 h-5"></i>
            </button>
        </div>
    `);

  $row.find(".remove-key-btn").on("click", function () {
    $row.remove();
  });

  $("#customKeysList").append($row);
}

// Initial Custom Key Row
// addCustomKeyRow();

$("#addCustomKeyBtn").on("click", () => addCustomKeyRow());

/** Download JSON file */
$("#downloadJsonBtn").on("click", function () {
  const jsonString = $("#jsonOutput").text();
  const blob = new Blob([jsonString], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "converted_data.json";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

/** Save/Load Preset Functions (Placeholder/Simplified Local Storage) */

function saveMappingPreset() {
  const mapping = getMapping();
  const customKeys = getCustomKeys();
  const preset = { mapping, customKeys };
  localStorage.setItem("jsonMappingPreset", JSON.stringify(preset));
  showMessage("Mapping saved successfully to local storage!", "success");
}

function loadMappingPreset() {
  const presetString = localStorage.getItem("jsonMappingPreset");
  if (!presetString) {
    showMessage("No saved mapping found.", "error");
    return;
  }

  try {
    const preset = JSON.parse(presetString);

    // 1. Load Column Mapping (requires headers to be present)
    if (dataHeaders.length > 0) {
      populateColumnMapping(dataHeaders); // Re-render initial mapping
      preset.mapping.forEach((savedMap) => {
        const $row = $(`#mappingBody tr`).filter(function () {
          return (
            $(this).find(".mapping-target-path").data("original") ===
            savedMap.originalKey
          );
        });
        if ($row.length) {
          $row.find(".mapping-target-path").val(savedMap.targetPath || "");
          $row.find(".mapping-data-type").val(savedMap.dataType || "auto");
          $row.find(".mapping-action").val(savedMap.action || "keep");
          $row
            .find(".mapping-transform-rule")
            .val(savedMap.transformRule || "");
        }
      });
      showMessage("Column mapping loaded.", "info");
    } else {
      showMessage(
        "Mapping loaded, but requires file to be uploaded to apply column settings.",
        "info"
      );
    }

    // 2. Load Custom Keys
    $("#customKeysList").empty();
    if (preset.customKeys.length > 0) {
      preset.customKeys.forEach((key) =>
        addCustomKeyRow(key.key, key.value, key.type)
      );
    } else {
      addCustomKeyRow(); // Add one empty row if no custom keys
    }

    showMessage("Mapping preset loaded successfully!", "success");
  } catch (e) {
    showMessage("Error loading mapping preset.", "error");
  }
}

$("#savePresetBtn").on("click", saveMappingPreset);
$("#loadPresetBtn").on("click", loadMappingPreset);

// --- Initial Setup and Event Listeners ---
$(document).ready(function () {
  // File Input Change
  $("#excelFile").on("change", function (e) {
    if (this.files.length > 0) {
      handleFile(this.files[0]);
    }
  });

  // Drag and Drop Handlers
  const dropzone = $("#dropzone");
  dropzone.on("dragover", function (e) {
    e.preventDefault();
    e.stopPropagation();
    $(this).addClass("border-primary bg-box/70");
  });
  dropzone.on("dragleave", function (e) {
    e.preventDefault();
    e.stopPropagation();
    $(this).removeClass("border-primary bg-box/70");
  });
  dropzone.on("drop", function (e) {
    e.preventDefault();
    e.stopPropagation();
    $(this).removeClass("border-primary bg-box/70");
    const files = e.originalEvent.dataTransfer.files;
    if (files.length > 0) {
      handleFile(files[0]);
    }
  });

  // Panel Toggle Logic
  $(".panel-header").on("click", function () {
    const $header = $(this);
    const $content = $header.next(".panel-content");
    const $icon = $header.find(".toggle-btn svg");

    $content.slideToggle(300, function () {
      if ($content.is(":visible")) {
        $icon
          .removeClass("rotate-0")
          .addClass("rotate-180")
          .attr("data-state", "expanded");
      } else {
        $icon
          .removeClass("rotate-180")
          .addClass("rotate-0")
          .attr("data-state", "collapsed");
      }
    });
  });
});
