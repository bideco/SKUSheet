/******************************************************
 * BARCODE TOOLS — FINAL CONSOLIDATED SCRIPT
 ******************************************************/

/***********************
 * GLOBAL CONSTANTS
 ***********************/
const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const CACHE = CacheService.getScriptCache();

const DEFAULTS = {
  BOX_WIDTH: 4,
  BOX_HEIGHT: 5,
  BOXES_PER_ROW: 3,   // ✅ default 3
  ROWS_PER_PAGE: 2,  // ✅ default 2
  INFO_FONT_SIZE: 12,
  SHOW_IMAGES: true
};

const PANEL_BG = "#f3f3f3"; // light gray (Sheets palette-like)

const STATUS_COLORS = {
  IDLE: "white",
  QUEUED: "#fff2cc",
  BUSY: "#cfe2f3",
  SUCCESS: "#d9ead3",
  ERROR: "#f4cccc"
};

/***********************
 * SIMPLE TRIGGERS
 ***********************/
function onOpen() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName("Sheet1");
    if (!sheet) return;

    // Compact A column (checkbox-sized)
    sheet.setColumnWidth(1, 48);

    // Reasonable defaults for barcode area columns
    for (let c = 2; c <= 30; c++) {
      sheet.setColumnWidth(c, 90);
    }

    sheet.setHiddenGridlines(false);

    // Clean up any orphaned settings/help panel
    cleanupPanelIfPresent_(sheet);

    // Ensure A-column controls exist
    setupAColumnControls_(sheet);

    // Reset A1 status
    setStatus_("", STATUS_COLORS.IDLE);

    SpreadsheetApp.getActive().toast("Barcode Tools ready", "Ready", 3);
  } catch (err) {
    Logger.log("onOpen error: " + err);
  }
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== "Sheet1") return;

    const row = e.range.getRow();
    const col = e.range.getColumn();

    // A1 — SKU input
    if (row === 1 && col === 1) {
      handleSkuInput_(e.value);
      return;
    }

    // Column A checkboxes (A2–A8)
    if (col === 1 && row >= 2 && row <= 8) {
      handleCheckbox_(row);
    }
  } catch (err) {
    Logger.log("onEdit error: " + err);
  }
}

/***********************
 * SKU INPUT + QUEUE
 ***********************/
function handleSkuInput_(value) {
  if (!value) return;

  setStatus_("Queued", STATUS_COLORS.QUEUED);

  const lines = String(value)
    .split("\n")
    .map(v => v.trim())
    .filter(Boolean);

  lines.forEach(line => {
    const barcodeData = formatSKU_(line);
    const skuForQueue =
      barcodeData.type === "ean13"
        ? normalizeEAN_(line)
        : line.trim();

    enqueueSku_(skuForQueue);
  });

  // Never hold SKU text in A1
  SpreadsheetApp.getActive()
    .getSheetByName("Sheet1")
    .getRange("A1")
    .clearContent();
}

/***********************
 * CHECKBOX HANDLER
 ***********************/
function handleCheckbox_(row) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");
  const cell = sheet.getRange(row, 1);
  if (!cell.getValue()) return;

  // Checkboxes behave like buttons
  cell.setValue(false);

  switch (row) {
    case 2: // Process Queue
      processQueue_();
      break;
    case 3: // Queue Items (display-only placeholder)
      break;
    case 4: // Force Refresh
      SCRIPT_PROPS.setProperty("FORCE_REFRESH", "true");
      SpreadsheetApp.getActive().toast("Force refresh enabled (EAN only)");
      break;
    case 5: // Reset Page
      resetPage_();
      break;
    case 6: // Print Layout
      togglePrintLayout_();
      break;
    case 7: // Settings
      openSettingsPanel_();
      break;
    case 8: // Help
      openHelpPanel_();
      break;
  }
}

/***********************
 * QUEUE PROCESSING
 ***********************/
function processQueue_() {
  const queue = getQueue_();
  if (!queue.length) return;

  setStatus_("Busy", STATUS_COLORS.BUSY);

  while (queue.length) {
    const sku = queue.shift();
    processSku_(sku);
  }

  saveQueue_([]);
  setStatus_("", STATUS_COLORS.SUCCESS);
}

/***********************
 * SKU PROCESSOR
 ***********************/
function processSku_(inputSku) {
  const rawSku = String(inputSku).trim();
  const barcodeData = formatSKU_(rawSku);

  const normalized =
    barcodeData.type === "ean13"
      ? normalizeEAN_(rawSku)
      : rawSku;

  const productData =
    barcodeData.type !== "ean13"
      ? { info: "SKU: " + rawSku, imageUrl: null }
      : getProductData_(normalized);

  const pos = getNextPosition_();
  drawBarcodeBox_(pos, barcodeData, productData);
}

/***********************
 * BARCODE DRAWING
 ***********************/
function drawBarcodeBox_(pos, barcodeData, productData) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");

  const bw = getNum_("BOX_WIDTH");
  const bh = getNum_("BOX_HEIGHT");

  // Barcode box
  const barcodeRange = sheet.getRange(pos.row, pos.col, bh, bw);
  barcodeRange.merge();
  barcodeRange
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("white");

  let fontSize = 40;

  if (barcodeData.type === "ean13") {
    // Dynamic scaling, 190 baseline at width 4
    const BASE_FONT = 190;
    const BASE_WIDTH = 4;
    fontSize = Math.round(BASE_FONT * (bw / BASE_WIDTH));
    fontSize = Math.max(120, Math.min(260, fontSize));
    barcodeRange.setFontFamily("Libre Barcode EAN13 Text");
  } else if (barcodeData.type === "code39") {
    barcodeRange.setFontFamily("Libre Barcode 39");
  } else {
    barcodeRange.setFontFamily("Libre Barcode 128");
  }

  barcodeRange
    .setFontSize(fontSize)
    .setValue("'" + barcodeData.value);

  // Product info box — SAME HEIGHT & WIDTH
  const infoRange = sheet.getRange(pos.row, pos.col + bw, bh, bw);
  infoRange.merge();
  infoRange
    .setValue(productData.info || "")
    .setWrap(true)
    .setFontSize(getNum_("INFO_FONT_SIZE"))
    .setVerticalAlignment("top")
    .setBackground("#f0f0f0");

  if (productData.imageUrl && getBool_("SHOW_IMAGES")) {
    infoRange.setFormula(`=IMAGE("${productData.imageUrl}",1)`);
  }
}

/***********************
 * POSITIONING
 ***********************/
function getNextPosition_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");
  const bw = getNum_("BOX_WIDTH");
  const bh = getNum_("BOX_HEIGHT");
  const perRow = getNum_("BOXES_PER_ROW");

  let count = 0;
  for (let r = 1; r <= sheet.getMaxRows(); r += bh) {
    for (let c = 2; c < 2 + perRow * bw * 2; c += bw * 2) {
      if (sheet.getRange(r, c).getValue()) count++;
    }
  }

  const row = 1 + Math.floor(count / perRow) * bh;
  const col = 2 + (count % perRow) * bw * 2;

  return { row, col };
}

/***********************
 * SKU FORMATTING
 ***********************/
function formatSKU_(sku) {
  const s = String(sku).trim();

  if (/^\d{12,13}$/.test(s)) {
    return { type: "ean13", value: s.length === 12 ? "0" + s : s };
  }
  if (/^[A-Z0-9\-.$/%+ ]+$/i.test(s)) {
    return { type: "code39", value: s.toUpperCase() };
  }
  return { type: "code128", value: s };
}

function normalizeEAN_(sku) {
  const digits = String(sku).replace(/\D/g, "");
  return digits.length === 12 ? "0" + digits : digits;
}

/***********************
 * PRODUCT DATA (EAN ONLY)
 ***********************/
function getProductData_(ean) {
  const force = SCRIPT_PROPS.getProperty("FORCE_REFRESH") === "true";
  const cacheKey = "EAN_" + ean;

  if (!force) {
    const cached = CACHE.get(cacheKey);
    if (cached) return JSON.parse(cached);
  }

  const apiKey = SCRIPT_PROPS.getProperty("SERPAPI_KEY");
  if (!apiKey) return { info: "SKU: " + ean, imageUrl: null };

  const url =
    "https://serpapi.com/search.json?engine=amazon&amazon_domain=amazon.com&k=" +
    encodeURIComponent(ean) +
    "&api_key=" +
    apiKey;

  const res = UrlFetchApp.fetch(url);
  const json = JSON.parse(res.getContentText());
  const item = json.organic_results?.[0];

  const data = {
    info: item?.title || "SKU: " + ean,
    imageUrl: item?.thumbnail || null
  };

  CACHE.put(cacheKey, JSON.stringify(data), 21600);
  SCRIPT_PROPS.deleteProperty("FORCE_REFRESH");
  return data;
}

/***********************
 * QUEUE STORAGE
 ***********************/
function getQueue_() {
  return JSON.parse(SCRIPT_PROPS.getProperty("QUEUE") || "[]");
}
function saveQueue_(q) {
  SCRIPT_PROPS.setProperty("QUEUE", JSON.stringify(q));
}
function enqueueSku_(sku) {
  const q = getQueue_();
  q.push(sku);
  saveQueue_(q);
}

/***********************
 * SETTINGS / HELP PANEL
 ***********************/
function openSettingsPanel_() {
  openPanel_([
    ["Box Width", getNum_("BOX_WIDTH")],
    ["Box Height", getNum_("BOX_HEIGHT")],
    ["Boxes Per Row", getNum_("BOXES_PER_ROW")],
    ["Rows Per Page", getNum_("ROWS_PER_PAGE")],
    ["Info Font Size", getNum_("INFO_FONT_SIZE")],
    ["Show Images", getBool_("SHOW_IMAGES")]
  ]);
}

function openHelpPanel_() {
  openPanel_([
    ["Usage", "Enter SKU(s) in A1, one per line."],
    ["Process", "Use A2 to process the queue."],
    ["Reset", "A5 clears all barcodes."],
    ["Print", "A6 toggles print layout."],
    ["Mobile", "Checkboxes act as buttons."]
  ]);
}

function openPanel_(rows) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");

  cleanupPanelIfPresent_(sheet);

  const panelRows = rows.length + 4;
  sheet.insertRowsBefore(1, panelRows);
  SpreadsheetApp.flush();

  const panelRange = sheet.getRange(1, 1, panelRows, 2);
  panelRange.clearFormat().setBackground(PANEL_BG).setFontSize(24);

  rows.forEach((r, i) => {
    sheet.getRange(3 + i, 1).setValue(r[0]);
    sheet.getRange(3 + i, 2).setValue(r[1]);
  });

  sheet.getRange("A1").setNote("__PANEL_OPEN__");
  setStatus_("", STATUS_COLORS.BUSY);
}

/***********************
 * PANEL CLEANUP
 ***********************/
function cleanupPanelIfPresent_(sheet) {
  const note = sheet.getRange("A1").getNote();
  if (note === "__PANEL_OPEN__") {
    sheet.getRange(1, 1, 40, 10).clearContent().clearFormat();
    sheet.deleteRows(1, 40);
    sheet.getRange("A1").setNote("");
  }
}

/***********************
 * RESET PAGE
 ***********************/
function resetPage_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");
  sheet.getRange("B1:Z2000").clearContent().clearFormat();
  SpreadsheetApp.getActive().toast("Page reset");
}

/***********************
 * PRINT LAYOUT
 ***********************/
function togglePrintLayout_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");
  const hidden = sheet.isColumnHiddenByUser(1);
  hidden ? sheet.showColumns(1) : sheet.hideColumns(1);
  sheet.setHiddenGridlines(!sheet.hasHiddenGridlines());
}

/***********************
 * UI HELPERS
 ***********************/
function setStatus_(text, color) {
  const cell = SpreadsheetApp.getActive()
    .getSheetByName("Sheet1")
    .getRange("A1");
  cell.setBackground(color);
  if (text !== undefined) cell.setValue(text);
}

/***********************
 * UTIL
 ***********************/
function getNum_(k) {
  return Number(SCRIPT_PROPS.getProperty(k) || DEFAULTS[k]);
}
function getBool_(k) {
  return SCRIPT_PROPS.getProperty(k) !== "false";
}
