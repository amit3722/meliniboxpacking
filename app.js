const tabButtons = document.querySelectorAll(".tab-btn");
const distributionTab = document.getElementById("distributionTab");
const orderTab = document.getElementById("orderTab");

const distributionInfoForm = document.getElementById("distributionInfoForm");
const itemNoInput = document.getElementById("itemNo");
const batchNoInput = document.getElementById("batchNo");
const packerNameInput = document.getElementById("packerName");
const distDateInput = document.getElementById("distDate");
const distHeaderRow = document.getElementById("distHeaderRow");
const distributionInfoBlock = document.getElementById("distributionInfoBlock");
const distributionCalcBlock = document.getElementById("distributionCalcBlock");
const distributionStepError = document.getElementById("distributionStepError");

const distributionForm = document.getElementById("distributionForm");
const distTotalColorsInput = document.getElementById("distTotalColors");
const createColorFieldsBtn = document.getElementById("createColorFieldsBtn");
const colorFieldsWrap = document.getElementById("colorFieldsWrap");
const distSizeInput = document.getElementById("distSize");
const distMinColorInBoxInput = document.getElementById("distMinColorInBox");
const distMaxPcsOfColorInput = document.getElementById("distMaxPcsOfColor");
const distMinPcsInput = document.getElementById("distMinPcs");
const distMaxPcsInput = document.getElementById("distMaxPcs");
const distributionClearBtn = document.getElementById("distributionClearBtn");
const distributionCalcError = document.getElementById("distributionCalcError");
const balancePcsModeInputs = document.querySelectorAll('input[name="balancePcsMode"]');

const metaItemNoEl = document.getElementById("metaItemNo");
const metaBatchNoEl = document.getElementById("metaBatchNo");
const metaPackerNameEl = document.getElementById("metaPackerName");
const metaDistDateEl = document.getElementById("metaDistDate");

const metricTotalColorsEl = document.getElementById("metricTotalColors");
const metricTotalQtyEl = document.getElementById("metricTotalQty");
const metricSizeEl = document.getElementById("metricSize");
const qtyReceivedValueEl = document.getElementById("qtyReceivedValue");
const qtyLooseValueEl = document.getElementById("qtyLooseValue");
const distributionTablesContainer = document.getElementById("distributionTablesContainer");
const a4SheetEl = document.querySelector(".a4-sheet");
const saveOutputBtn = document.getElementById("saveOutputBtn");
const exportFormatSelect = document.getElementById("exportFormat");
const exportOutputBtn = document.getElementById("exportOutputBtn");
const printOutputBtn = document.getElementById("printOutputBtn");

const orderForm = document.getElementById("orderForm");
const customerInput = document.getElementById("customerName");
const productInput = document.getElementById("productName");
const orderUnitsInput = document.getElementById("orderUnits");
const unitPriceInput = document.getElementById("unitPrice");
const deliveryNotesInput = document.getElementById("deliveryNotes");

const metaCustomerEl = document.getElementById("metaCustomer");
const metaProductEl = document.getElementById("metaProduct");
const metaUnitsEl = document.getElementById("metaUnits");
const metaPriceEl = document.getElementById("metaPrice");
const metaSubtotalEl = document.getElementById("metaSubtotal");
const metaNotesEl = document.getElementById("metaNotes");

const ALLOWED_SIZES = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "7XL"];
const MAX_COLOR_FIELDS = 10;
let distributionTableCount = 0;
const distributionTotals = {
  boxQty: 0,
  receivedQty: 0,
  packedPcs: 0,
  loose: 0,
};

function toInt(value) {
  const n = Number.parseInt(value, 10);
  return Number.isNaN(n) ? 0 : Math.max(0, n);
}

function toFloat(value) {
  const n = Number.parseFloat(value);
  return Number.isNaN(n) ? null : Math.max(0, n);
}

function formatMoney(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}

function getTodayISODate() {
  const now = new Date();
  now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
  return now.toISOString().slice(0, 10);
}

function formatDisplayDate(isoDate) {
  if (!isoDate) {
    return "-";
  }
  const parts = isoDate.split("-");
  if (parts.length !== 3) {
    return isoDate;
  }
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

function normalizeDistributionHeaderInputs() {
  itemNoInput.value = itemNoInput.value.toUpperCase();
  batchNoInput.value = batchNoInput.value.toUpperCase();
  packerNameInput.value = packerNameInput.value.toUpperCase();
}

function isDistributionHeaderValid() {
  return (
    itemNoInput.value.trim() !== "" &&
    batchNoInput.value.trim() !== "" &&
    packerNameInput.value.trim() !== "" &&
    distDateInput.value.trim() !== ""
  );
}

function unlockDistributionForm() {
  distributionInfoBlock.classList.add("is-hidden");
  distributionCalcBlock.classList.remove("is-hidden");
  distributionStepError.classList.add("is-hidden");
}

function normalizeNumericInput(inputEl) {
  inputEl.value = inputEl.value.replace(/[^\d]/g, "");
}

function validateSizeField() {
  distSizeInput.value = distSizeInput.value.trim().toUpperCase();
  if (distSizeInput.value === "" || ALLOWED_SIZES.includes(distSizeInput.value)) {
    distSizeInput.setCustomValidity("");
    return true;
  }

  distSizeInput.setCustomValidity("Size must be between XS and 7XL.");
  return false;
}

function showDistributionCalcError(message) {
  distributionCalcError.textContent = message;
  distributionCalcError.classList.remove("is-hidden");
}

function hideDistributionCalcError() {
  distributionCalcError.classList.add("is-hidden");
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function renderColorFieldRows(totalColors) {
  const count = toInt(totalColors);
  if (count <= 0) {
    colorFieldsWrap.innerHTML = '<p class="mini-help">Color and Qty fields will generate based on Total No. Of Color.</p>';
    return;
  }

  const rows = [];
  for (let i = 1; i <= count; i += 1) {
    rows.push(`
      <div class="color-field-row">
        <label>
          Color ${i}
          <input class="color-name" type="text" placeholder="Color ${i}" required />
        </label>
        <label>
          Qty ${i}
          <input class="color-qty" type="number" min="0" step="1" inputmode="numeric" required />
        </label>
      </div>
    `);
  }

  colorFieldsWrap.innerHTML = rows.join("");

  colorFieldsWrap.querySelectorAll(".color-name").forEach((inputEl) => {
    inputEl.addEventListener("input", () => {
      inputEl.value = inputEl.value.toUpperCase();
    });
  });

  colorFieldsWrap.querySelectorAll(".color-qty").forEach((inputEl) => {
    inputEl.addEventListener("input", () => {
      normalizeNumericInput(inputEl);
    });
  });
}

function getBalancePcsMode() {
  const selected = Array.from(balancePcsModeInputs).find((inputEl) => inputEl.checked);
  return selected ? selected.value : "yes";
}

function buildDistributionTableMarkup(size, colorNames, packedColorTotals, patternRows, totalBoxQty, totalLooseQty) {
  const safeSize = size ? escapeHtml(size) : "-";
  const safeColorHeaders = colorNames.map((color) => `<th>${escapeHtml(color)}</th>`).join("");
  const safePackedTotals = colorNames
    .map((_, index) => `<th>${packedColorTotals[index] ?? 0}</th>`)
    .join("");
  const totalCols = 2 + Math.max(colorNames.length, 1);

  if (!patternRows.length) {
    return `
      <table>
        <thead>
          <tr><th colspan="${totalCols}" class="size-head">SIZE : ${safeSize}</th></tr>
          <tr class="shade-qty-row">
            <th colspan="2">Used Qty (Excl. Loose)</th>
            ${safePackedTotals || "<th>0</th>"}
          </tr>
          <tr>
            <th>Pcs/Box</th>
            <th>Box_Qty</th>
            ${safeColorHeaders || "<th>Color</th>"}
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>0</td>
            <td>0</td>
            ${safeColorHeaders ? colorNames.map(() => "<td>0</td>").join("") : "<td>0</td>"}
          </tr>
        </tbody>
        <tfoot>
          <tr class="table-total-row">
            <td>Total Box</td>
            <td>0</td>
            ${safeColorHeaders ? colorNames.map(() => "<td></td>").join("") : "<td></td>"}
          </tr>
          <tr class="table-total-row">
            <td>Loose</td>
            <td>${totalLooseQty}</td>
            ${safeColorHeaders ? colorNames.map(() => "<td></td>").join("") : "<td></td>"}
          </tr>
        </tfoot>
      </table>
    `;
  }

  const bodyRows = patternRows
    .map((pattern) => {
      const colorCells = pattern.allocations
        .map((value) => `<td>${value > 0 ? value : 0}</td>`)
        .join("");
      return `<tr><td>${pattern.pcsPerBox}</td><td>${pattern.boxQty}</td>${colorCells}</tr>`;
    })
    .join("");

  const footerColorCells = colorNames.map(() => "<td></td>").join("");

  return `
    <table>
      <thead>
        <tr><th colspan="${totalCols}" class="size-head">SIZE : ${safeSize}</th></tr>
        <tr class="shade-qty-row">
          <th colspan="2">Used Qty (Excl. Loose)</th>
          ${safePackedTotals}
        </tr>
        <tr>
          <th>Pcs/Box</th>
          <th>Box_Qty</th>
          ${safeColorHeaders}
        </tr>
      </thead>
      <tbody>
        ${bodyRows}
      </tbody>
      <tfoot>
        <tr class="table-total-row">
          <td>Total Box</td>
          <td>${totalBoxQty}</td>
          ${footerColorCells}
        </tr>
        <tr class="table-total-row">
          <td>Loose</td>
          <td>${totalLooseQty}</td>
          ${footerColorCells}
        </tr>
      </tfoot>
    </table>
  `;
}

function appendDistributionTable(size, colorNames, packedColorTotals, patternRows, totalBoxQty, totalLooseQty) {
  if (distributionTableCount === 0) {
    distributionTablesContainer.innerHTML = "";
  }
  distributionTableCount += 1;
  const tableBlock = document.createElement("div");
  tableBlock.className = "table-wrap history-table";
  tableBlock.innerHTML = `
    <p class="table-caption">Table ${distributionTableCount}</p>
    ${buildDistributionTableMarkup(size, colorNames, packedColorTotals, patternRows, totalBoxQty, totalLooseQty)}
  `;
  distributionTablesContainer.appendChild(tableBlock);
}

function renderInitialDistributionTable() {
  distributionTableCount = 0;
  distributionTablesContainer.innerHTML = "";
  distributionTotals.boxQty = 0;
  distributionTotals.receivedQty = 0;
  distributionTotals.packedPcs = 0;
  distributionTotals.loose = 0;
}

function resetDistributionOutput() {
  renderInitialDistributionTable();
  metricTotalColorsEl.textContent = "0";
  metricTotalQtyEl.textContent = "0";
  metricSizeEl.textContent = "0";
  qtyReceivedValueEl.textContent = "0";
  qtyLooseValueEl.textContent = "0";
}

function tryBuildBoxForTarget(remaining, options, targetPcs) {
  const { minColorInBox, maxPcsColor, minPcs, maxPcs } = options;
  if (targetPcs < minPcs || targetPcs > maxPcs || targetPcs <= 0) {
    return { success: false };
  }

  const minRequiredColors = Math.max(1, minColorInBox);
  const temp = remaining.slice();
  const allocations = Array(temp.length).fill(0);
  let total = 0;

  const availableIndices = temp
    .map((qty, idx) => ({ qty, idx }))
    .filter((item) => item.qty > 0)
    .sort((a, b) => b.qty - a.qty)
    .map((item) => item.idx);

  if (availableIndices.length < minRequiredColors || minRequiredColors > targetPcs) {
    return { success: false };
  }

  for (let i = 0; i < minRequiredColors; i += 1) {
    const idx = availableIndices[i];
    temp[idx] -= 1;
    allocations[idx] += 1;
    total += 1;
  }

  while (total < targetPcs) {
    let selected = -1;
    let selectedQty = -1;
    for (let i = 0; i < temp.length; i += 1) {
      if (temp[i] > 0 && allocations[i] < maxPcsColor && temp[i] > selectedQty) {
        selected = i;
        selectedQty = temp[i];
      }
    }

    if (selected === -1) {
      return { success: false };
    }

    temp[selected] -= 1;
    allocations[selected] += 1;
    total += 1;
  }

  const usedColors = allocations.filter((value) => value > 0).length;
  if (usedColors < minRequiredColors) {
    return { success: false };
  }

  return {
    success: true,
    remaining: temp,
    allocations,
    pcsPerBox: targetPcs,
  };
}

function distributeBoxes(colorData, options) {
  const remaining = colorData.map((item) => item.qty);
  const boxes = [];

  // Phase 1: Build max-filled valid boxes to minimize loose.
  while (true) {
    const result = tryBuildBoxForTarget(remaining, options, options.maxPcs);
    if (!result.success) {
      break;
    }
    for (let i = 0; i < remaining.length; i += 1) {
      remaining[i] = result.remaining[i];
    }
    boxes.push({
      allocations: result.allocations,
      pcsPerBox: result.pcsPerBox,
    });
  }

  // Phase 2 (NOT mode): keep forming valid boxes with smaller targets until loose can be 0.
  if (options.mode === "not") {
    while (remaining.some((qty) => qty > 0)) {
      const remainingTotal = remaining.reduce((sum, qty) => sum + qty, 0);
      const upperTarget = Math.min(options.maxPcs, remainingTotal);
      const lowerTarget = Math.min(options.minPcs, upperTarget);

      let found = null;
      for (let target = upperTarget; target >= lowerTarget; target -= 1) {
        const tryResult = tryBuildBoxForTarget(remaining, options, target);
        if (tryResult.success) {
          found = tryResult;
          break;
        }
      }

      if (!found) {
        break;
      }

      for (let i = 0; i < remaining.length; i += 1) {
        remaining[i] = found.remaining[i];
      }
      boxes.push({
        allocations: found.allocations,
        pcsPerBox: found.pcsPerBox,
      });
    }
  }

  const looseQty = remaining.reduce((sum, qty) => sum + qty, 0);
  return { boxes, looseQty };
}

function groupBoxPatterns(boxes) {
  const map = new Map();
  boxes.forEach((box) => {
    const key = box.allocations.join("|");
    if (!map.has(key)) {
      map.set(key, {
        allocations: box.allocations,
        pcsPerBox: box.pcsPerBox,
        boxQty: 0,
      });
    }
    map.get(key).boxQty += 1;
  });
  return Array.from(map.values()).sort((a, b) => b.pcsPerBox - a.pcsPerBox);
}

function calculateDistributionForm() {
  hideDistributionCalcError();
  validateSizeField();

  const totalColors = toInt(distTotalColorsInput.value);
  if (totalColors <= 0) {
    showDistributionCalcError("Total No. Of Color must be a numeric value greater than 0.");
    return;
  }
  if (totalColors > MAX_COLOR_FIELDS) {
    showDistributionCalcError(`Maximum ${MAX_COLOR_FIELDS} colors allowed.`);
    return;
  }

  const currentRows = colorFieldsWrap.querySelectorAll(".color-field-row").length;
  if (currentRows !== totalColors) {
    showDistributionCalcError("Please click Create Fields after entering Total No. Of Color.");
    return;
  }

  if (!validateSizeField()) {
    showDistributionCalcError("Please choose a valid size from XS to 7XL.");
    distSizeInput.reportValidity();
    return;
  }

  const minColorInBox = toInt(distMinColorInBoxInput.value);
  const maxPcsColor = toInt(distMaxPcsOfColorInput.value);
  const minPcs = toInt(distMinPcsInput.value);
  const maxPcs = toInt(distMaxPcsInput.value);

  if (minPcs > maxPcs) {
    showDistributionCalcError("Min Pcs cannot be greater than Max Pcs.");
    return;
  }

  const colorRows = Array.from(colorFieldsWrap.querySelectorAll(".color-field-row"));
  const data = [];
  for (let i = 0; i < colorRows.length; i += 1) {
    const row = colorRows[i];
    const colorInput = row.querySelector(".color-name");
    const qtyInput = row.querySelector(".color-qty");

    colorInput.value = colorInput.value.trim().toUpperCase();
    normalizeNumericInput(qtyInput);

    if (colorInput.value === "") {
      showDistributionCalcError(`Please enter color name for row ${i + 1}.`);
      return;
    }

    if (qtyInput.value.trim() === "") {
      showDistributionCalcError(`Please enter numeric Qty for row ${i + 1}.`);
      return;
    }

    data.push({
      color: colorInput.value,
      qty: toInt(qtyInput.value),
    });
  }

  const totalQty = data.reduce((sum, item) => sum + item.qty, 0);

  if (maxPcsColor <= 0) {
    showDistributionCalcError("Max. Pcs/Color must be greater than 0.");
    return;
  }

  if (maxPcs <= 0) {
    showDistributionCalcError("Max Pcs must be greater than 0.");
    return;
  }

  if (minColorInBox > totalColors) {
    showDistributionCalcError("Min. Color/Box cannot be greater than Total No. Of Color.");
    return;
  }

  if (maxPcsColor > maxPcs) {
    showDistributionCalcError("Max. Pcs/Color cannot be greater than Max Pcs.");
    return;
  }

  if (minColorInBox > maxPcs) {
    showDistributionCalcError("Min. Color/Box cannot be greater than Max Pcs.");
    return;
  }

  const balanceMode = getBalancePcsMode();
  const { boxes, looseQty } = distributeBoxes(data, {
    minColorInBox,
    maxPcsColor,
    minPcs,
    maxPcs,
    mode: balanceMode,
  });

  const totalBoxQty = boxes.length;
  const totalLooseQty = looseQty;
  const totalPackedQty = totalQty - totalLooseQty;
  if (balanceMode === "not" && totalLooseQty > 0) {
    showDistributionCalcError("Unable to force Loose = 0 with current Max Pcs value.");
    return;
  }

  const patternRows = groupBoxPatterns(boxes);
  const packedColorTotals = Array(data.length).fill(0);
  boxes.forEach((box) => {
    box.allocations.forEach((qty, index) => {
      packedColorTotals[index] += qty;
    });
  });

  appendDistributionTable(
    distSizeInput.value,
    data.map((item) => item.color),
    packedColorTotals,
    patternRows,
    totalBoxQty,
    totalLooseQty
  );

  distributionTotals.boxQty += totalBoxQty;
  distributionTotals.receivedQty += totalQty;
  distributionTotals.packedPcs += totalPackedQty;
  distributionTotals.loose += totalLooseQty;

  metricTotalColorsEl.textContent = String(distributionTotals.boxQty);
  metricTotalQtyEl.textContent = String(distributionTotals.packedPcs);
  metricSizeEl.textContent = String(distributionTotals.loose);
  qtyReceivedValueEl.textContent = String(distributionTotals.receivedQty);
  qtyLooseValueEl.textContent = String(distributionTotals.loose);

  distSizeInput.value = "";
  distSizeInput.setCustomValidity("");
  colorFieldsWrap.querySelectorAll(".color-qty").forEach((qtyInput) => {
    qtyInput.value = "";
  });
}

function createColorFields() {
  hideDistributionCalcError();
  const totalColors = toInt(distTotalColorsInput.value);
  if (totalColors <= 0) {
    showDistributionCalcError("Enter Total No. Of Color first, then click Create Fields.");
    return;
  }
  if (totalColors > MAX_COLOR_FIELDS) {
    showDistributionCalcError(`You can create maximum ${MAX_COLOR_FIELDS} color fields only.`);
    return;
  }

  renderColorFieldRows(totalColors);
}

function clearDistributionForm() {
  distributionForm.reset();
  hideDistributionCalcError();
  distSizeInput.setCustomValidity("");
  balancePcsModeInputs.forEach((inputEl) => {
    inputEl.checked = inputEl.value === "yes";
  });
  renderColorFieldRows(0);
  resetDistributionOutput();
}

function switchTab(tabName) {
  const isDistribution = tabName === "distribution";
  distributionTab.classList.toggle("is-hidden", !isDistribution);
  orderTab.classList.toggle("is-hidden", isDistribution);
  tabButtons.forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.tab === tabName);
  });
}

function updateDistributionHeaderPreview() {
  metaItemNoEl.textContent = itemNoInput.value.trim() || "-";
  metaBatchNoEl.textContent = batchNoInput.value.trim() || "-";
  metaPackerNameEl.textContent = packerNameInput.value.trim() || "-";
  metaDistDateEl.textContent = formatDisplayDate(distDateInput.value.trim());
}

function updateOrderPreview() {
  const units = toInt(orderUnitsInput.value);
  const unitPrice = toFloat(unitPriceInput.value);
  const customer = customerInput.value.trim();
  const product = productInput.value.trim();
  const notes = deliveryNotesInput.value.trim();

  metaCustomerEl.textContent = customer || "-";
  metaProductEl.textContent = product || "-";
  metaUnitsEl.textContent = String(units);
  metaPriceEl.textContent = unitPrice === null ? "-" : formatMoney(unitPrice);
  metaNotesEl.textContent = notes || "-";

  if (unitPrice === null || units === 0) {
    metaSubtotalEl.textContent = "-";
  } else {
    metaSubtotalEl.textContent = formatMoney(units * unitPrice);
  }
}

function makeSafeFileNamePart(value) {
  const text = value.trim().toUpperCase();
  if (!text) {
    return "";
  }
  return text.replace(/[\\/:*?"<>|]+/g, "_").replace(/\s+/g, "_");
}

function getOutputBaseFileName() {
  const itemPart = makeSafeFileNamePart(itemNoInput.value);
  const batchPart = makeSafeFileNamePart(batchNoInput.value);
  if (itemPart && batchPart) {
    return `${itemPart}_${batchPart}`;
  }
  if (itemPart) {
    return itemPart;
  }
  if (batchPart) {
    return batchPart;
  }
  return "DISTRIBUTION_OUTPUT";
}

function getOutputSheetMarkup() {
  const sheetClone = a4SheetEl.cloneNode(true);
  sheetClone.querySelectorAll(".no-print").forEach((node) => node.remove());
  return sheetClone.outerHTML;
}

function getEmbeddedStyles() {
  const cssBlocks = [];
  Array.from(document.styleSheets).forEach((styleSheet) => {
    try {
      const rules = styleSheet.cssRules;
      if (!rules || rules.length === 0) {
        return;
      }
      const text = Array.from(rules)
        .map((rule) => rule.cssText)
        .join("\n");
      if (text.trim() !== "") {
        cssBlocks.push(text);
      }
    } catch (error) {
      // Ignore external stylesheets that the browser blocks for security.
    }
  });

  return cssBlocks.join("\n").replace(/<\/style/gi, "<\\/style");
}

function buildPrintableDocument(title, mode = "save") {
  const sheetMarkup = getOutputSheetMarkup();
  const embeddedStyles = getEmbeddedStyles();
  const modeStyles = mode === "excel"
    ? "body{margin:0;background:#ffffff;}"
    : `
      body{margin:0;min-height:100vh;background:#ffffff;display:grid;place-items:center;}
      @media print{
        @page{size:A4;margin:0;}
        body{min-height:auto;}
      }
    `;

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${escapeHtml(title)}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Source+Serif+4:wght@500;600&display=swap" rel="stylesheet" />
  <style>
    ${embeddedStyles}
    ${modeStyles}
  </style>
</head>
<body>
  ${sheetMarkup}
</body>
</html>`;
}

function getPdfCellStyle(cellEl, rowEl) {
  const isHeaderCell = cellEl.tagName === "TH";
  const isSizeHead = cellEl.classList.contains("size-head");
  const isShadeQtyRow = rowEl.classList.contains("shade-qty-row");
  const isTotalRow = rowEl.classList.contains("table-total-row");

  const style = {
    fontSize: 8,
    textColor: [17, 24, 39],
    halign: "center",
    valign: "middle",
    cellPadding: 1.5,
    lineColor: [209, 213, 219],
    lineWidth: 0.1,
  };

  if (isSizeHead) {
    style.fontStyle = "bold";
    style.fillColor = [238, 242, 255];
  } else if (isShadeQtyRow || isTotalRow) {
    style.fontStyle = "bold";
    style.fillColor = [248, 250, 252];
  } else if (isHeaderCell) {
    style.fontStyle = "bold";
    style.fillColor = [249, 250, 251];
  }

  return style;
}

function buildPdfBodyFromHtmlTable(tableEl) {
  const rows = Array.from(tableEl.querySelectorAll("tr"));
  return rows.map((rowEl) => {
    const cellElements = Array.from(rowEl.children).filter((node) => node.tagName === "TH" || node.tagName === "TD");
    return cellElements.map((cellEl) => {
      const cell = {
        content: (cellEl.textContent || "").trim() || " ",
        styles: getPdfCellStyle(cellEl, rowEl),
      };
      const colSpan = Number.parseInt(cellEl.getAttribute("colspan") || "1", 10);
      const rowSpan = Number.parseInt(cellEl.getAttribute("rowspan") || "1", 10);
      if (colSpan > 1) {
        cell.colSpan = colSpan;
      }
      if (rowSpan > 1) {
        cell.rowSpan = rowSpan;
      }
      return cell;
    });
  });
}

async function exportOutputAsPdf(fileBase) {
  if (!window.jspdf || typeof window.jspdf.jsPDF !== "function") {
    throw new Error("jsPDF is not available.");
  }

  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({
    orientation: "portrait",
    unit: "mm",
    format: "a4",
    compress: true,
  });

  if (typeof pdf.autoTable !== "function") {
    throw new Error("jsPDF autoTable plugin is not available.");
  }

  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const margin = 2;
  const contentWidth = pageWidth - (margin * 2);
  let currentY = 7;

  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(12);
  pdf.text("BOX DISTRIBUTION OUTPUT", margin, currentY);
  currentY += 4;

  const headerInfoRows = [
    [{ content: "DATE", styles: { fontStyle: "bold", fillColor: [249, 250, 251] } }, { content: metaDistDateEl.textContent.trim() || "-" }],
    [{ content: "ITEM NO.", styles: { fontStyle: "bold", fillColor: [249, 250, 251] } }, { content: metaItemNoEl.textContent.trim() || "-" }],
    [{ content: "BATCH NO.", styles: { fontStyle: "bold", fillColor: [249, 250, 251] } }, { content: metaBatchNoEl.textContent.trim() || "-" }],
    [{ content: "PACKER", styles: { fontStyle: "bold", fillColor: [249, 250, 251] } }, { content: metaPackerNameEl.textContent.trim() || "-" }],
  ];

  pdf.autoTable({
    startY: currentY,
    body: headerInfoRows,
    theme: "grid",
    tableWidth: contentWidth,
    margin: { left: margin, right: margin },
    styles: {
      fontSize: 8,
      textColor: [17, 24, 39],
      lineColor: [209, 213, 219],
      lineWidth: 0.1,
      cellPadding: 1.6,
    },
    columnStyles: {
      0: { cellWidth: contentWidth * 0.28 },
      1: { cellWidth: contentWidth * 0.72 },
    },
  });
  currentY = pdf.lastAutoTable.finalY + 2;

  const qtyRows = Array.from(a4SheetEl.querySelectorAll(".qty-table tbody tr")).map((rowEl) => {
    const tds = rowEl.querySelectorAll("td");
    return [
      { content: tds[0] ? tds[0].textContent.trim() : "", styles: { fontStyle: "bold", fillColor: [249, 250, 251] } },
      { content: tds[1] ? tds[1].textContent.trim() : "", styles: { halign: "center" } },
    ];
  });

  pdf.autoTable({
    startY: currentY,
    body: qtyRows,
    theme: "grid",
    tableWidth: contentWidth,
    margin: { left: margin, right: margin },
    styles: {
      fontSize: 8,
      textColor: [17, 24, 39],
      lineColor: [209, 213, 219],
      lineWidth: 0.1,
      cellPadding: 1.6,
    },
    columnStyles: {
      0: { cellWidth: contentWidth * 0.78 },
      1: { cellWidth: contentWidth * 0.22, halign: "center" },
    },
  });
  currentY = pdf.lastAutoTable.finalY + 2;

  const metricsRow = [[
    { content: "Total Box Qty", styles: { fontStyle: "bold", fillColor: [248, 250, 252] } },
    { content: metricTotalColorsEl.textContent.trim() || "0", styles: { fontStyle: "bold" } },
    { content: "Total pcs", styles: { fontStyle: "bold", fillColor: [248, 250, 252] } },
    { content: metricTotalQtyEl.textContent.trim() || "0", styles: { fontStyle: "bold" } },
    { content: "Total Loose", styles: { fontStyle: "bold", fillColor: [248, 250, 252] } },
    { content: metricSizeEl.textContent.trim() || "0", styles: { fontStyle: "bold" } },
  ]];

  pdf.autoTable({
    startY: currentY,
    body: metricsRow,
    theme: "grid",
    tableWidth: contentWidth,
    margin: { left: margin, right: margin },
    styles: {
      fontSize: 8,
      textColor: [17, 24, 39],
      lineColor: [209, 213, 219],
      lineWidth: 0.1,
      cellPadding: 1.6,
      halign: "center",
    },
  });
  currentY = pdf.lastAutoTable.finalY + 3;

  const tableBlocks = Array.from(distributionTablesContainer.querySelectorAll(".history-table"));
  let tablesOnCurrentPage = 0;

  tableBlocks.forEach((tableBlock, index) => {
    if (tablesOnCurrentPage >= 4 || currentY > pageHeight - 25) {
      pdf.addPage();
      currentY = 8;
      tablesOnCurrentPage = 0;
    }

    const captionEl = tableBlock.querySelector(".table-caption");
    const tableEl = tableBlock.querySelector("table");
    const captionText = captionEl ? captionEl.textContent.trim() : `Table ${index + 1}`;

    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(9);
    pdf.text(captionText, margin, currentY);
    currentY += 2;

    if (tableEl) {
      const bodyData = buildPdfBodyFromHtmlTable(tableEl);
      const pagesBefore = pdf.internal.getNumberOfPages();
      pdf.autoTable({
        startY: currentY,
        body: bodyData,
        theme: "grid",
        tableWidth: contentWidth,
        margin: { left: margin, right: margin },
        rowPageBreak: "avoid",
        styles: {
          fontSize: 7.7,
          textColor: [17, 24, 39],
          lineColor: [209, 213, 219],
          lineWidth: 0.1,
          cellPadding: 1.4,
        },
      });

      currentY = pdf.lastAutoTable.finalY + 3;
      const pagesAfter = pdf.internal.getNumberOfPages();
      if (pagesAfter > pagesBefore) {
        tablesOnCurrentPage = 1;
      } else {
        tablesOnCurrentPage += 1;
      }
    }
  });

  pdf.save(`${fileBase}.pdf`);
}

function parseExcelCellValue(text) {
  const trimmed = text.trim();
  if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
    return Number(trimmed);
  }
  return trimmed;
}

function setWorksheetCellBorder(cell) {
  cell.border = {
    top: { style: "thin", color: { argb: "FFD1D5DB" } },
    left: { style: "thin", color: { argb: "FFD1D5DB" } },
    bottom: { style: "thin", color: { argb: "FFD1D5DB" } },
    right: { style: "thin", color: { argb: "FFD1D5DB" } },
  };
}

function applyHtmlCellStyleToWorksheetCell(worksheetCell, htmlCell) {
  const rowEl = htmlCell.closest("tr");
  const isHeader = htmlCell.tagName === "TH";
  const isTotalRow = rowEl && rowEl.classList.contains("table-total-row");
  const isSizeHead = htmlCell.classList.contains("size-head");
  const isShadeQtyRow = rowEl && rowEl.classList.contains("shade-qty-row");

  worksheetCell.alignment = {
    horizontal: "center",
    vertical: "middle",
    wrapText: true,
  };
  worksheetCell.font = {
    name: "Calibri",
    size: isSizeHead ? 11 : 10,
    bold: isHeader || isTotalRow,
    color: { argb: "FF111827" },
  };
  setWorksheetCellBorder(worksheetCell);

  if (isSizeHead) {
    worksheetCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFEEF2FF" },
    };
  } else if (isTotalRow || isShadeQtyRow) {
    worksheetCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF8FAFC" },
    };
  } else if (isHeader) {
    worksheetCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF9FAFB" },
    };
  }
}

function writeHtmlTableToWorksheet(worksheet, tableEl, startRow, startCol) {
  const tableRows = Array.from(tableEl.querySelectorAll("tr"));
  const grid = [];
  const merges = [];
  let maxCol = 0;

  tableRows.forEach((rowEl, rowIdx) => {
    if (!grid[rowIdx]) {
      grid[rowIdx] = [];
    }
    let colIdx = 0;
    const rowCells = Array.from(rowEl.children).filter((node) => node.tagName === "TH" || node.tagName === "TD");
    rowCells.forEach((cellEl) => {
      while (grid[rowIdx][colIdx]) {
        colIdx += 1;
      }

      const rowSpan = Math.max(1, Number.parseInt(cellEl.getAttribute("rowspan") || "1", 10));
      const colSpan = Math.max(1, Number.parseInt(cellEl.getAttribute("colspan") || "1", 10));

      for (let r = 0; r < rowSpan; r += 1) {
        if (!grid[rowIdx + r]) {
          grid[rowIdx + r] = [];
        }
        for (let c = 0; c < colSpan; c += 1) {
          grid[rowIdx + r][colIdx + c] = r === 0 && c === 0 ? cellEl : "__SPAN__";
        }
      }

      if (rowSpan > 1 || colSpan > 1) {
        merges.push({
          r1: rowIdx,
          c1: colIdx,
          r2: rowIdx + rowSpan - 1,
          c2: colIdx + colSpan - 1,
        });
      }

      colIdx += colSpan;
      maxCol = Math.max(maxCol, colIdx);
    });
  });

  for (let r = 0; r < grid.length; r += 1) {
    worksheet.getRow(startRow + r).height = 20;
    for (let c = 0; c < maxCol; c += 1) {
      const marker = grid[r] ? grid[r][c] : null;
      if (!marker || marker === "__SPAN__") {
        continue;
      }

      const worksheetCell = worksheet.getCell(startRow + r, startCol + c);
      worksheetCell.value = parseExcelCellValue(marker.textContent || "");
      applyHtmlCellStyleToWorksheetCell(worksheetCell, marker);
      const currentWidth = worksheet.getColumn(startCol + c).width || 10;
      worksheet.getColumn(startCol + c).width = Math.max(currentWidth, 12);
    }
  }

  merges.forEach((mergeRange) => {
    worksheet.mergeCells(
      startRow + mergeRange.r1,
      startCol + mergeRange.c1,
      startRow + mergeRange.r2,
      startCol + mergeRange.c2
    );
  });

  return {
    rowCount: grid.length,
    colCount: maxCol,
  };
}

async function exportOutputAsExcel(fileBase) {
  if (!window.ExcelJS) {
    throw new Error("ExcelJS is not available.");
  }

  const workbook = new window.ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Distribution Output");
  for (let col = 1; col <= 12; col += 1) {
    worksheet.getColumn(col).width = 14;
  }

  let rowCursor = 1;
  worksheet.getCell(rowCursor, 1).value = "BOX DISTRIBUTION OUTPUT";
  worksheet.getCell(rowCursor, 1).font = { name: "Calibri", size: 14, bold: true, color: { argb: "FF0F172A" } };
  worksheet.getCell(rowCursor, 1).alignment = { vertical: "middle", horizontal: "left" };
  worksheet.mergeCells(rowCursor, 1, rowCursor, 6);
  rowCursor += 2;

  const headerRows = [
    ["DATE", metaDistDateEl.textContent.trim() || "-"],
    ["ITEM NO.", metaItemNoEl.textContent.trim() || "-"],
    ["BATCH NO.", metaBatchNoEl.textContent.trim() || "-"],
    ["PACKER", metaPackerNameEl.textContent.trim() || "-"],
  ];

  const qtyRows = Array.from(a4SheetEl.querySelectorAll(".qty-table tbody tr")).map((rowEl) => {
    const cells = rowEl.querySelectorAll("td");
    return [
      cells[0] ? cells[0].textContent.trim() : "",
      cells[1] ? cells[1].textContent.trim() : "",
    ];
  });

  const topBlockRows = Math.max(headerRows.length, qtyRows.length);
  for (let i = 0; i < topBlockRows; i += 1) {
    const rowNumber = rowCursor + i;
    worksheet.getRow(rowNumber).height = 20;

    const left = headerRows[i];
    if (left) {
      worksheet.getCell(rowNumber, 1).value = left[0];
      worksheet.getCell(rowNumber, 1).font = { bold: true, name: "Calibri", size: 10 };
      worksheet.getCell(rowNumber, 2).value = left[1];
      setWorksheetCellBorder(worksheet.getCell(rowNumber, 1));
      setWorksheetCellBorder(worksheet.getCell(rowNumber, 2));
    }

    const right = qtyRows[i];
    if (right) {
      worksheet.getCell(rowNumber, 5).value = right[0];
      worksheet.getCell(rowNumber, 5).font = { bold: true, name: "Calibri", size: 10 };
      worksheet.getCell(rowNumber, 6).value = parseExcelCellValue(right[1] || "");
      setWorksheetCellBorder(worksheet.getCell(rowNumber, 5));
      setWorksheetCellBorder(worksheet.getCell(rowNumber, 6));
    }
  }

  rowCursor += topBlockRows + 1;

  const metricRows = [
    ["Total Box Qty", metricTotalColorsEl.textContent.trim() || "0"],
    ["Total pcs", metricTotalQtyEl.textContent.trim() || "0"],
    ["Total Loose", metricSizeEl.textContent.trim() || "0"],
  ];

  metricRows.forEach((metric, index) => {
    const labelCol = 1 + (index * 2);
    worksheet.getCell(rowCursor, labelCol).value = metric[0];
    worksheet.getCell(rowCursor, labelCol).font = { bold: true, name: "Calibri", size: 10 };
    worksheet.getCell(rowCursor, labelCol).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF8FAFC" },
    };
    worksheet.getCell(rowCursor, labelCol + 1).value = parseExcelCellValue(metric[1]);
    worksheet.getCell(rowCursor, labelCol + 1).font = { bold: true, name: "Calibri", size: 10 };
    setWorksheetCellBorder(worksheet.getCell(rowCursor, labelCol));
    setWorksheetCellBorder(worksheet.getCell(rowCursor, labelCol + 1));
  });

  rowCursor += 2;
  const tableBlocks = Array.from(distributionTablesContainer.querySelectorAll(".history-table"));
  tableBlocks.forEach((tableBlock, index) => {
    const caption = tableBlock.querySelector(".table-caption");
    const tableEl = tableBlock.querySelector("table");
    const captionText = caption ? caption.textContent.trim() : `Table ${index + 1}`;

    worksheet.getCell(rowCursor, 1).value = captionText;
    worksheet.getCell(rowCursor, 1).font = { name: "Calibri", size: 11, bold: true, color: { argb: "FF334155" } };
    worksheet.getCell(rowCursor, 1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFF8FAFC" },
    };
    worksheet.mergeCells(rowCursor, 1, rowCursor, 12);
    rowCursor += 1;

    if (tableEl) {
      const written = writeHtmlTableToWorksheet(worksheet, tableEl, rowCursor, 1);
      rowCursor += written.rowCount + 2;
    }
  });

  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.pageSetup = {
    paperSize: 9,
    orientation: "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.3,
      bottom: 0.3,
      header: 0.15,
      footer: 0.15,
    },
  };

  const buffer = await workbook.xlsx.writeBuffer();
  downloadFileContent(buffer, `${fileBase}.xlsx`, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

function downloadFileContent(content, fileName, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = objectUrl;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(objectUrl);
}

function openPrintWindow(docTitle, shouldAutoPrint) {
  const printWindow = window.open("", "_blank", "noopener,noreferrer,width=980,height=900");
  if (!printWindow) {
    showDistributionCalcError("Popup blocked. Allow popups to use Print.");
    return;
  }

  const html = buildPrintableDocument(docTitle, "print");
  printWindow.document.open();
  printWindow.document.write(html);
  printWindow.document.close();

  if (shouldAutoPrint) {
    let printed = false;
    const triggerPrint = () => {
      if (printed) {
        return;
      }
      printed = true;
      printWindow.focus();
      printWindow.print();
    };

    // Wait for styles to load before invoking print.
    printWindow.addEventListener("load", triggerPrint, { once: true });
    setTimeout(triggerPrint, 500);
  }
}

function handleSaveOutput() {
  hideDistributionCalcError();
  const fileBase = getOutputBaseFileName();
  const html = buildPrintableDocument(fileBase, "save");
  downloadFileContent(html, `${fileBase}.html`, "text/html;charset=utf-8");
}

async function handleExportOutput() {
  hideDistributionCalcError();
  const fileBase = getOutputBaseFileName();
  const format = exportFormatSelect.value;
  const hasOutput = distributionTableCount > 0;

  if (!hasOutput) {
    showDistributionCalcError("Pehle Calculate karo, phir Export use karo.");
    return;
  }

  const originalLabel = exportOutputBtn.textContent;
  exportOutputBtn.disabled = true;
  exportOutputBtn.textContent = "Exporting...";

  try {
    if (format === "excel") {
      await exportOutputAsExcel(fileBase);
      return;
    }

    if (format === "pdf") {
      await exportOutputAsPdf(fileBase);
    }
  } catch (error) {
    if (format === "pdf") {
      showDistributionCalcError("PDF export failed. Ek baar page reload karke dobara try karo.");
    } else {
      showDistributionCalcError("Excel export failed. Ek baar page reload karke dobara try karo.");
    }
  } finally {
    exportOutputBtn.disabled = false;
    exportOutputBtn.textContent = originalLabel;
  }
}

function handlePrintOutput() {
  hideDistributionCalcError();
  openPrintWindow(getOutputBaseFileName(), true);
}

distributionForm.addEventListener("submit", (event) => {
  event.preventDefault();
  calculateDistributionForm();
});

[itemNoInput, batchNoInput, packerNameInput, distDateInput].forEach((input) => {
  input.addEventListener("input", () => {
    normalizeDistributionHeaderInputs();
    if (isDistributionHeaderValid()) {
      distributionStepError.classList.add("is-hidden");
    }
  });
});

distributionInfoForm.addEventListener("submit", (event) => {
  event.preventDefault();
  normalizeDistributionHeaderInputs();
  updateDistributionHeaderPreview();
  if (!isDistributionHeaderValid()) {
    distributionStepError.classList.remove("is-hidden");
    return;
  }
  distHeaderRow.classList.remove("is-hidden");
  unlockDistributionForm();
});

[distTotalColorsInput, distMinColorInBoxInput, distMaxPcsOfColorInput, distMinPcsInput, distMaxPcsInput].forEach((inputEl) => {
  inputEl.addEventListener("input", () => {
    normalizeNumericInput(inputEl);
    if (inputEl === distTotalColorsInput && toInt(inputEl.value) > MAX_COLOR_FIELDS) {
      inputEl.value = String(MAX_COLOR_FIELDS);
    }
  });
});

createColorFieldsBtn.addEventListener("click", () => {
  createColorFields();
});

distSizeInput.addEventListener("input", () => {
  distSizeInput.value = distSizeInput.value.toUpperCase();
  validateSizeField();
});

distSizeInput.addEventListener("blur", () => {
  validateSizeField();
});

distributionClearBtn.addEventListener("click", () => {
  clearDistributionForm();
});

saveOutputBtn.addEventListener("click", () => {
  handleSaveOutput();
});

exportOutputBtn.addEventListener("click", () => {
  void handleExportOutput();
});

printOutputBtn.addEventListener("click", () => {
  handlePrintOutput();
});

orderForm.addEventListener("submit", (event) => {
  event.preventDefault();
  updateOrderPreview();
});

[customerInput, productInput, orderUnitsInput, unitPriceInput, deliveryNotesInput].forEach((input) => {
  input.addEventListener("input", updateOrderPreview);
});

tabButtons.forEach((btn) => {
  btn.addEventListener("click", () => {
    switchTab(btn.dataset.tab);
  });
});

if (!distDateInput.value) {
  distDateInput.value = getTodayISODate();
}

renderColorFieldRows(0);
resetDistributionOutput();
updateOrderPreview();
