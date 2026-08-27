import * as XLSX from "xlsx";
import Papa from "papaparse";

export const normalizeText = value => String(value ?? "").trim().toLowerCase().replace(/\s+/g, " ");
export const normalizeTarget = value => normalizeText(value).replace(/[“”]/g, '"').replace(/\s*=\s*/g, "=");

export function toNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (value == null || value === "") return 0;
  let text = String(value).trim();
  if (!text) return 0;
  const negative = /^\(.*\)$/.test(text);
  text = text.replace(/[(),$%\s]/g, "").replace(/[^0-9eE+\-.]/g, "");
  const n = Number(text);
  return Number.isFinite(n) ? (negative ? -Math.abs(n) : n) : 0;
}

export function toRate(value) {
  if (value == null || value === "") return null;
  if (typeof value === "string" && value.includes("%")) return toNumber(value) / 100;
  const n = toNumber(value);
  if (!Number.isFinite(n)) return null;
  return n > 1 ? n / 100 : n;
}

function parseCsv(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: ({ data, errors }) => {
        const serious = errors.find(e => e.type === "Quotes");
        if (serious) reject(new Error(serious.message));
        else resolve({ fileName: file.name, sheets: [{ name: "CSV", rows: data }] });
      },
      error: reject,
    });
  });
}

// Amazon-generated workbooks sometimes declare !ref=A1 despite containing thousands
// of populated cells. SheetJS then converts only A1 unless we repair the range.
export function repairWorksheetRef(ws) {
  const addresses = Object.keys(ws).filter(k => !k.startsWith("!"));
  if (!addresses.length) return ws;
  let minR = Infinity, minC = Infinity, maxR = -1, maxC = -1;
  for (const address of addresses) {
    const cell = XLSX.utils.decode_cell(address);
    minR = Math.min(minR, cell.r); minC = Math.min(minC, cell.c);
    maxR = Math.max(maxR, cell.r); maxC = Math.max(maxC, cell.c);
  }
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
  return ws;
}

export async function parseWorkbook(file) {
  if (!file) throw new Error("Missing file.");
  const ext = file.name.split(".").pop()?.toLowerCase();
  if (!["csv", "xlsx", "xls"].includes(ext)) throw new Error(`Unsupported file type: .${ext}`);
  if (ext === "csv") return parseCsv(file);

  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: "array", cellDates: true, dense: false });
  const sheets = wb.SheetNames.map(name => {
    const ws = repairWorksheetRef(wb.Sheets[name]);
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
    return { name, rows };
  });
  return { fileName: file.name, sheets };
}

function sheetScore(sheet, kind) {
  const sample = sheet.rows[0] || {};
  const keys = Object.keys(sample).map(normalizeText);
  const has = re => keys.some(k => re.test(k));
  let score = 0;
  if (kind === "searchTerm") {
    if (/sp search term report/i.test(sheet.name)) score += 5;
    if (has(/customer search term|search term/)) score += 8;
    if (has(/^spend$/)) score += 2;
    if (has(/orders/)) score += 2;
  }
  if (kind === "targeting") {
    if (has(/^targeting$/)) score += 8;
    if (has(/match type/)) score += 2;
    if (has(/^spend$/)) score += 2;
    if (has(/top-of-search impression share/)) score += 2;
  }
  if (kind === "bulk") {
    if (/sponsored products campaigns/i.test(sheet.name)) score += 8;
    if (has(/^entity$/)) score += 8;
    if (has(/^campaign id$/)) score += 3;
    if (has(/^ad group id$/)) score += 3;
    if (has(/^operation$/)) score += 1;
  }
  if (kind === "bulkSearch") {
    if (/sp search term report/i.test(sheet.name)) score += 8;
    if (has(/^campaign id$/)) score += 3;
    if (has(/customer search term/)) score += 6;
  }
  return score;
}

export function pickBestSheet(workbook, kind, required = true) {
  const scored = workbook.sheets
    .map(s => ({ ...s, score: sheetScore(s, kind) }))
    .sort((a, b) => b.score - a.score);
  if (!scored.length || scored[0].score < 5) {
    if (required) throw new Error(`Could not identify a valid ${kind} sheet in ${workbook.fileName}.`);
    return null;
  }
  return scored[0];
}

function findColumn(sample, patterns, label, required = true) {
  const keys = Object.keys(sample || {});
  const found = keys.find(k => patterns.some(re => re.test(k.trim())));
  if (!found && required) throw new Error(`Required column not found: ${label}`);
  return found || null;
}

export function mapReportColumns(rows, kind) {
  const s = rows[0] || {};
  const common = {
    startDate: findColumn(s, [/^start date$/i], "Start Date", false),
    endDate: findColumn(s, [/^end date$/i], "End Date", false),
    currency: findColumn(s, [/^currency$/i], "Currency", false),
    campaignId: findColumn(s, [/^campaign id$/i], "Campaign ID", false),
    campaignName: findColumn(s, [/^campaign name$/i, /campaign name \(informational only\)/i], "Campaign Name", false),
    adGroupId: findColumn(s, [/^ad group id$/i], "Ad Group ID", false),
    adGroupName: findColumn(s, [/^ad group name$/i, /ad group name \(informational only\)/i], "Ad Group Name", false),
    spend: findColumn(s, [/^spend$/i], "Spend"),
    orders: findColumn(s, [/^orders$/i, /total orders/i], "Orders"),
    sales: findColumn(s, [/^sales$/i, /total sales/i], "Sales"),
    clicks: findColumn(s, [/^clicks$/i], "Clicks"),
    impressions: findColumn(s, [/^impressions$/i], "Impressions", false),
    matchType: findColumn(s, [/^match type$/i], "Match Type", false),
    cpc: findColumn(s, [/^cpc$/i, /cost per click/i], "CPC", false),
  };
  if (kind === "searchTerm" || kind === "bulkSearch") {
    common.term = findColumn(s, [/customer search term/i, /^search term$/i], "Customer Search Term");
    common.targeting = findColumn(s, [/^targeting$/i, /^keyword text$/i, /^product targeting expression$/i], "Targeting", false);
    common.keywordId = findColumn(s, [/^keyword id$/i], "Keyword ID", false);
    common.productTargetId = findColumn(s, [/^product targeting id$/i], "Product Targeting ID", false);
    common.bid = findColumn(s, [/^bid$/i], "Bid", false);
  }
  if (kind === "targeting") {
    common.targeting = findColumn(s, [/^targeting$/i], "Targeting");
    common.tosShare = findColumn(s, [/top-of-search impression share/i], "Top-of-search impression share", false);
  }
  return common;
}

export function reportDateRange(rows, cols) {
  if (!cols.startDate && !cols.endDate) return null;
  const dates = [];
  for (const r of rows) {
    for (const c of [cols.startDate, cols.endDate]) {
      if (!c || !r[c]) continue;
      const d = new Date(r[c]);
      if (!Number.isNaN(d.getTime())) dates.push(d);
    }
  }
  if (!dates.length) return null;
  return { start: new Date(Math.min(...dates)), end: new Date(Math.max(...dates)) };
}

export function bulkDateRangeFromFileName(fileName) {
  const m = String(fileName || "").match(/-(20\d{6})-(20\d{6})-/);
  if (!m) return null;
  const parse = s => new Date(`${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6,8)}T00:00:00`);
  return { start: parse(m[1]), end: parse(m[2]) };
}
