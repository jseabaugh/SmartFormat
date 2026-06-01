/* ============================================================
   Smart Format – Non-Volatile Conditional Formatting Add-In
   All formatting is applied as static cell properties via the
   Office JS Range API, so Excel never re-evaluates conditions.
   ============================================================ */

"use strict";

// ── State ─────────────────────────────────────────────────────────────────────
let rules = [];           // Array of rule objects
let ruleIdCounter = 0;

// ── Office Initialization ─────────────────────────────────────────────────────
Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) return;
  bindUI();
  renderRuleList();
});

// ── UI Bindings ───────────────────────────────────────────────────────────────
function bindUI() {
  document.getElementById("btnPickRange").onclick    = pickSelection;
  document.getElementById("btnAddRule").onclick      = addRule;
  document.getElementById("btnApply").onclick        = applyFormatting;
  document.getElementById("btnClearFormat").onclick  = clearFormatting;
  document.getElementById("btnClearRules").onclick   = clearAllRules;
  document.getElementById("ruleType").onchange       = updateValueFields;
  updateValueFields();
}

// ── Condition type → value field labels ──────────────────────────────────────
const TYPE_META = {
  greaterThan:   { label: "Value",      v2: false },
  lessThan:      { label: "Value",      v2: false },
  equalTo:       { label: "Value",      v2: false },
  notEqualTo:    { label: "Value",      v2: false },
  between:       { label: "Min value",  v2: true  },
  textContains:  { label: "Text",       v2: false },
  blank:         { label: null,         v2: false },
  notBlank:      { label: null,         v2: false },
  top10:         { label: "Top N",      v2: false },
  bottom10:      { label: "Bottom N",   v2: false },
  aboveAverage:  { label: null,         v2: false },
  belowAverage:  { label: null,         v2: false },
  duplicate:     { label: null,         v2: false },
  unique:        { label: null,         v2: false },
  colorScale:    { label: null,         v2: false, noColor: true },
  dataBar:       { label: null,         v2: false, noColor: true },
};

function updateValueFields() {
  const type = document.getElementById("ruleType").value;
  const meta = TYPE_META[type] || {};
  const v1Row = document.getElementById("value1Row");
  const v2Row = document.getElementById("value2Row");
  const colorSection = document.getElementById("colorSection");

  v1Row.style.display     = meta.label  ? "" : "none";
  v2Row.style.display     = meta.v2     ? "" : "none";
  colorSection.style.display = meta.noColor ? "none" : "";

  if (meta.label) {
    document.getElementById("value1Label").textContent = meta.label;
  }
}

// ── Pick selection ─────────────────────────────────────────────────────────────
async function pickSelection() {
  try {
    await Excel.run(async ctx => {
      const range = ctx.workbook.getSelectedRange();
      range.load("address");
      await ctx.sync();
      // Strip sheet name prefix (Sheet1!A1:B5 → A1:B5)
      const addr = range.address.replace(/^[^!]+!/, "");
      document.getElementById("rangeInput").value = addr;
    });
  } catch(e) {
    showStatus("Could not read selection: " + e.message, "error");
  }
}

// ── Add a rule to the list ────────────────────────────────────────────────────
function addRule() {
  const range  = document.getElementById("rangeInput").value.trim();
  const type   = document.getElementById("ruleType").value;
  const value1 = document.getElementById("ruleValue1").value.trim();
  const value2 = document.getElementById("ruleValue2").value.trim();
  const bg     = document.getElementById("bgColor").value;
  const font   = document.getElementById("fontColor").value;
  const bold   = document.getElementById("boldCheck").checked;

  if (!range) { showStatus("Please enter a target range.", "error"); return; }

  const meta = TYPE_META[type] || {};
  if (meta.label && !value1) { showStatus("Please enter a value for the condition.", "error"); return; }
  if (meta.v2 && !value2)    { showStatus("Please enter the second value (max).", "error"); return; }

  const rule = { id: ++ruleIdCounter, range, type, value1, value2, bg, font, bold };
  rules.push(rule);
  renderRuleList();
  showStatus(`Rule #${rule.id} added. Click "Apply" to apply formatting.`, "info");
}

// ── Render rule chips ─────────────────────────────────────────────────────────
function renderRuleList() {
  const list  = document.getElementById("ruleList");
  const empty = document.getElementById("emptyState");
  document.getElementById("ruleCount").textContent = rules.length;

  if (!rules.length) {
    empty.style.display = "";
    list.innerHTML = "";
    return;
  }
  empty.style.display = "none";

  list.innerHTML = rules.map(r => `
    <div class="rule-item" data-id="${r.id}">
      <div class="rule-swatch" style="background:${r.bg};"></div>
      <div class="rule-desc">
        <div class="rule-name">${r.range} — ${friendlyType(r.type)}</div>
        <div class="rule-detail">${ruleDetail(r)}</div>
      </div>
      <button class="rule-delete" title="Remove rule" onclick="removeRule(${r.id})">✕</button>
    </div>
  `).join("");
}

function friendlyType(t) {
  const map = {
    greaterThan:"Greater than", lessThan:"Less than", equalTo:"Equal to",
    between:"Between", notEqualTo:"Not equal to", textContains:"Text contains",
    blank:"Blank", notBlank:"Non-blank", top10:"Top N", bottom10:"Bottom N",
    aboveAverage:"Above avg", belowAverage:"Below avg",
    duplicate:"Duplicate", unique:"Unique",
    colorScale:"Color scale", dataBar:"Data bar"
  };
  return map[t] || t;
}

function ruleDetail(r) {
  if (r.type === "between")        return `${r.value1} – ${r.value2}`;
  if (r.type === "colorScale" || r.type === "dataBar") return "visual scale";
  if (r.value1)                    return r.value1;
  return "";
}

function removeRule(id) {
  rules = rules.filter(r => r.id !== id);
  renderRuleList();
}

function clearAllRules() {
  rules = [];
  renderRuleList();
  showStatus("All rules cleared.", "info");
}

// ── Presets ───────────────────────────────────────────────────────────────────
const PRESETS = {
  red:    { bg: "#FFC7CE", font: "#9C0006" },
  yellow: { bg: "#FFEB9C", font: "#9C5700" },
  green:  { bg: "#C6EFCE", font: "#276221" },
  blue:   { bg: "#DDEBF7", font: "#1F497D" },
  none:   { bg: "#FFFFFF", font: "#000000" },
};

function applyPreset(name) {
  const p = PRESETS[name];
  if (!p) return;
  document.getElementById("bgColor").value   = p.bg;
  document.getElementById("fontColor").value = p.font;
}

// ── Apply formatting (non-volatile) ──────────────────────────────────────────
async function applyFormatting() {
  if (!rules.length) { showStatus("No rules to apply.", "error"); return; }
  showStatus("Applying formatting…", "info");
  document.getElementById("btnApply").disabled = true;

  try {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();

      for (const rule of rules) {
        const target = sheet.getRange(rule.range);
        target.load(["values", "address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
        await ctx.sync();

        const values = target.values;
        const rows   = target.rowCount;
        const cols   = target.columnCount;
        const baseRow = target.rowIndex;
        const baseCol = target.columnIndex;

        // ── Compute statistics for avg / top-bottom rules ──────────────────
        let avg = 0, flatNums = [], sorted = [];
        if (["aboveAverage","belowAverage","top10","bottom10","duplicate","unique","colorScale","dataBar"].includes(rule.type)) {
          flatNums = values.flat().filter(v => typeof v === "number");
          avg      = flatNums.length ? flatNums.reduce((a,b)=>a+b,0) / flatNums.length : 0;
          sorted   = [...flatNums].sort((a,b) => a - b);
        }

        // Thresholds for top/bottom N
        const n = parseInt(rule.value1) || 10;
        let topThreshold    = sorted.length ? sorted[Math.max(0, sorted.length - n)] : Infinity;
        let bottomThreshold = sorted.length ? sorted[Math.min(sorted.length-1, n-1)]  : -Infinity;

        // Frequency map for duplicate/unique
        const freq = {};
        if (rule.type === "duplicate" || rule.type === "unique") {
          for (const v of values.flat()) {
            const k = String(v);
            freq[k] = (freq[k] || 0) + 1;
          }
        }

        // Color-scale: map numeric range to gradient
        const csMin = sorted[0];
        const csMax = sorted[sorted.length - 1];

        // ── Walk each cell ────────────────────────────────────────────────
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const cell  = target.getCell(r, c);
            const val   = values[r][c];
            const match = evaluateCondition(rule, val, { avg, topThreshold, bottomThreshold, freq });

            if (rule.type === "colorScale") {
              const color = colorScaleHex(val, csMin, csMax);
              cell.format.fill.color = color;
            } else if (rule.type === "dataBar") {
              // Simulate data bar with fill opacity via a solid gradient substitute
              const pct = typeof val === "number" && csMax !== csMin
                ? Math.max(0, Math.min(1, (val - csMin) / (csMax - csMin)))
                : 0;
              // Encode intensity into color lightness
              const hex = dataBarHex(pct);
              cell.format.fill.color = hex;
            } else if (match) {
              cell.format.fill.color = rule.bg;
              cell.format.font.color = rule.font;
              cell.format.font.bold  = rule.bold;
            } else {
              // Clear format for unmatched cells so stale formatting is removed
              cell.format.fill.clear();
              cell.format.font.color = "#000000";
              cell.format.font.bold  = false;
            }
          }
        }
      }

      await ctx.sync();
    });

    showStatus("✅ Formatting applied successfully!", "success");
  } catch(e) {
    showStatus("Error: " + e.message, "error");
  } finally {
    document.getElementById("btnApply").disabled = false;
  }
}

// ── Condition evaluator ───────────────────────────────────────────────────────
function evaluateCondition(rule, val, stats) {
  const { avg, topThreshold, bottomThreshold, freq } = stats;
  const num = parseFloat(rule.value1);
  const num2 = parseFloat(rule.value2);

  switch (rule.type) {
    case "greaterThan":   return typeof val === "number" && val > num;
    case "lessThan":      return typeof val === "number" && val < num;
    case "equalTo":       return String(val) === String(rule.value1) || val == num;
    case "notEqualTo":    return String(val) !== String(rule.value1) && val != num;
    case "between":       return typeof val === "number" && val >= num && val <= num2;
    case "textContains":  return typeof val === "string" && val.toLowerCase().includes(rule.value1.toLowerCase());
    case "blank":         return val === null || val === "" || val === undefined;
    case "notBlank":      return val !== null && val !== "" && val !== undefined;
    case "top10":         return typeof val === "number" && val >= topThreshold;
    case "bottom10":      return typeof val === "number" && val <= bottomThreshold;
    case "aboveAverage":  return typeof val === "number" && val > avg;
    case "belowAverage":  return typeof val === "number" && val < avg;
    case "duplicate":     return (freq[String(val)] || 0) > 1;
    case "unique":        return (freq[String(val)] || 0) === 1;
    default:              return false;
  }
}

// ── Color scale helper: green (low) → yellow → red (high) ───────────────────
function colorScaleHex(val, min, max) {
  if (typeof val !== "number" || min === max) return "#FFFFFF";
  const t = Math.max(0, Math.min(1, (val - min) / (max - min)));
  // Low = green (#63BE7B), Mid = yellow (#FFEB84), High = red (#F8696B)
  if (t < 0.5) {
    const s = t * 2;
    return lerpColor("#63BE7B", "#FFEB84", s);
  } else {
    const s = (t - 0.5) * 2;
    return lerpColor("#FFEB84", "#F8696B", s);
  }
}

function dataBarHex(pct) {
  // Simple blue shade: light → dark proportional to value
  return lerpColor("#D9E8F7", "#2E75B6", pct);
}

function lerpColor(a, b, t) {
  const ra = parseInt(a.slice(1,3),16), ga = parseInt(a.slice(3,5),16), ba2 = parseInt(a.slice(5,7),16);
  const rb = parseInt(b.slice(1,3),16), gb = parseInt(b.slice(3,5),16), bb2 = parseInt(b.slice(5,7),16);
  const r = Math.round(ra + (rb-ra)*t).toString(16).padStart(2,"0");
  const g = Math.round(ga + (gb-ga)*t).toString(16).padStart(2,"0");
  const bv = Math.round(ba2 + (bb2-ba2)*t).toString(16).padStart(2,"0");
  return `#${r}${g}${bv}`;
}

// ── Clear all formatting on the active sheet ──────────────────────────────────
async function clearFormatting() {
  try {
    await Excel.run(async ctx => {
      for (const rule of rules) {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(rule.range);
        range.format.fill.clear();
        range.format.font.color = "#000000";
        range.format.font.bold  = false;
      }
      await ctx.sync();
    });
    showStatus("Formatting cleared.", "info");
  } catch(e) {
    showStatus("Error: " + e.message, "error");
  }
}

// ── Status toast ──────────────────────────────────────────────────────────────
let statusTimer;
function showStatus(msg, type = "info") {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className   = `show ${type}`;
  clearTimeout(statusTimer);
  statusTimer = setTimeout(() => el.className = "", 4000);
}
