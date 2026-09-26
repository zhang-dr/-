import { getStore } from "@netlify/blobs";
import ExcelJS from "exceljs";
import { Readable } from "node:stream";
import * as XLSX from "xlsx";

const OPERATING_HASH = process.env.OPERATING_PASSWORD_HASH || "95f8ab63f4e9b5e1e8b9b85baf09d8c598fde734efc79f6291b2904ea180e04d";
const OPERATING_SECRET = process.env.OPERATING_SESSION_SECRET || "operating-session-secret";
const VOLUME_BASE = "https://zhangxihao.netlify.app/Volume";
const ORIGIN_CENTER = "Lutterworth HUB";
const SCHEMA_VERSION = 2;
const PLATFORMS = ["SF", "TT", "TEMU", "CERTIFIED_WAREHOUSE"];
const CENTERS = [
  { center: "Uxbridge Depot", shortName: "UXB", chinese: "阿克斯布里奇", prefix: "10" },
  { center: "Lutterworth HUB", shortName: "LTW", chinese: "拉特沃思", prefix: "20" },
  { center: "Rochdale Depot", shortName: "ROC", chinese: "罗奇代尔", prefix: "30" },
  { center: "Glasgow Depot", shortName: "GLA", chinese: "格拉斯哥", prefix: "40" },
];
const CODE_TO_OUTLET = {
  UK100001: "LONDON001", UK100002: "C.LONDON002", UK100003: "LONDON003", UK100004: "C.LONDON004", UK100005: "C.LONDON005", UK100006: "C.LONDON006", UK100007: "C.LONDON007", UK100008: "LONDON008", UK100009: "LONDON009", UK100010: "C.LONDON010", UK100011: "C.WATFORD001", UK100012: "C.LONDON012", UK100013: "C.LONDON013", UK100014: "C.LON.SE",
  UK200001: "C.BHX001", UK200002: "C.BHX002", UK200003: "C.LEICESTER001", UK200004: "C.COVENTRY001", UK200005: "C.WALSALL001", UK200006: "C.DUDLEY001", UK200012: "C.NOTTINGHAM001", UK200013: "C.WOLVERHAMPTON001", UK200014: "C.COVENTRY002",
  UK300001: "C.PRESTON001", UK300002: "C.MAN001", UK300003: "C.STOCKPORT001", UK300004: "C.LIV001", UK300005: "C.ROCHDALE001", UK300006: "C.CREWE001",
  UK400001: "C.GLASGOW001", UK400004: "C.EDINBURGH001", UK400005: "C.EDINBURGH002",
};
const OUTLET_TO_CODE = Object.fromEntries(Object.entries(CODE_TO_OUTLET).map(([code, outlet]) => [outlet, code]));
const DEFAULT_SETTINGS = { cubeFactor: 0.00957897245703636, vehicleOverrides: {} };
const FIELD_ALIASES = {
  pickupDate: ["提货时间", "提货日期", "Pickup Time", "Pickup Date", "pickup date"],
  routeCode: ["二段码", "route code", "Route Code", "线路编码", "路由编码"],
  outlet: ["末端网点", "目的网点", "网点名称", "outlet", "destination outlet", "末端网点编码"],
  outletCode: ["网点编码", "目的网点编码", "末端网点编码", "outlet code", "destination outlet code"],
  depot: ["末端机构", "目的机构", "Depot", "Hub", "末端中心/集散"],
  depotCode: ["末端机构编码", "目的机构编码", "depot code", "hub code"],
  adjustedPostcode: ["调整后邮编", "outward postcode", "postcode prefix"],
  receiverPostcode: ["consigneezip", "consignee zip", "收件邮编", "目的邮编", "destination postcode", "receiver postcode", "邮编"],
  status: ["启用状态", "status"],
  source: ["订单来源", "来源平台", "平台", "order source", "source", "platform"],
};

const json = (data, status = 200, headers = {}) => new Response(JSON.stringify(data), { status, headers: { "content-type": "application/json; charset=utf-8", ...headers } });
const clean = (value) => {
  if (value == null) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object") {
    if ("result" in value) return clean(value.result);
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || "").join("").trim();
    if ("text" in value) return clean(value.text);
    if ("hyperlink" in value) return clean(value.text || value.hyperlink);
    return "";
  }
  const text = String(value).trim();
  return /^#(?:REF|VALUE|N\/A|DIV\/0|NAME|NULL|NUM)!?$/i.test(text) ? "" : text;
};
const excelCellValue = (value) => {
  if (value == null) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object") {
    if ("result" in value) return excelCellValue(value.result);
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || "").join("");
    if ("text" in value) return value.text;
    if ("hyperlink" in value) return value.text || value.hyperlink;
    return "";
  }
  return value;
};
const normalizeHeader = (value) => clean(value).toLowerCase().replace(/[\s_./\\\-:：()（）]+/g, "");
const fieldValue = (row, field) => {
  const normalized = new Map(Object.keys(row).map((key) => [normalizeHeader(key), row[key]]));
  for (const alias of FIELD_ALIASES[field]) {
    const value = normalized.get(normalizeHeader(alias));
    if (clean(value)) return value;
  }
  return "";
};
const hasField = (row, field) => {
  const normalized = new Set(Object.keys(row).map(normalizeHeader));
  return FIELD_ALIASES[field].some((alias) => normalized.has(normalizeHeader(alias)));
};
const normalizeRouteCode = (value) => {
  const match = clean(value).toUpperCase().replace(/^UK/, "").match(/^(\d{2})\D*(\d{1,4})/);
  return match ? `${match[1]}-${String(Number(match[2])).padStart(3, "0")}` : "";
};
const codeFromSegment = (value) => {
  const route = normalizeRouteCode(value);
  return route ? "UK" + route.replace(/^(\d{2})-(\d{3})$/, (_, prefix, value) => `${prefix}${String(Number(value)).padStart(4, "0")}`) : "";
};
const codeFrom = (row) => codeFromSegment(fieldValue(row, "routeCode")) || OUTLET_TO_CODE[clean(fieldValue(row, "outlet")).toUpperCase()] || "";
const regionFrom = (code) => code.startsWith("UK10") ? "South" : code.startsWith("UK20") ? "Midlands" : code.startsWith("UK30") ? "Northwest" : code.startsWith("UK40") ? "North" : "Unmapped";
const centerFromCode = (code) => CENTERS.find((center) => code.startsWith("UK" + center.prefix))?.center || "Unmapped";
const centerFromDepot = (value, code) => {
  const text = clean(value).toLowerCase();
  return CENTERS.find((center) => text.includes(center.center.toLowerCase().replace(" depot", "").replace(" hub", "")))?.center || centerFromCode(code);
};
const shortNameForCenter = (centerName, centerCode = "") => {
  const explicit = clean(centerCode).toUpperCase().replace(/\d+$/g, "");
  if (explicit) return explicit;
  const known = CENTERS.find((center) => center.center === centerName)?.shortName;
  if (known) return known;
  return clean(centerName).split(/\s+/).map((part) => part[0] || "").join("").slice(0, 4).toUpperCase() || "CTR";
};
const centerDefinitions = (map, rows = []) => {
  const definitions = new Map();
  const add = (centerName, centerCode = "") => {
    const center = clean(centerName);
    if (!center || center === "Unmapped") return;
    const known = CENTERS.find((item) => item.center === center);
    const current = definitions.get(center);
    definitions.set(center, { center, shortName: shortNameForCenter(center, centerCode || current?.shortName), chinese: known?.chinese || current?.chinese || "" });
  };
  for (const entry of Object.values(map?.entries || {})) if (entry?.enabled) add(entry.centerName, entry.centerCode);
  for (const row of rows) { add(row.destinationCenter, row.destinationCenterCode); add(row.originCenter); }
  add(ORIGIN_CENTER);
  return [...definitions.values()].sort((a, b) => a.shortName.localeCompare(b.shortName));
};
const routeCode = (code) => code.replace(/^UK/, "").replace(/^(\d{2})(\d{4})$/, (_, prefix, value) => `${prefix}-${String(Number(value)).padStart(3, "0")}`);
const normalizePostcode = (value) => {
  const compact = clean(value).toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (!compact) return "Unknown";
  return compact.length >= 5 ? compact.slice(0, -3) + " " + compact.slice(-3) : compact;
};
const postcodeParts = (value, outwardFallback = "") => {
  const full = normalizePostcode(value);
  const district = full === "Unknown" ? clean(outwardFallback).toUpperCase().replace(/\s/g, "") || "Unknown" : full.split(" ")[0];
  const area = district.match(/^[A-Z]+/)?.[0] || "Unknown";
  const inward = full === "Unknown" ? "" : full.split(" ")[1] || "";
  return { area, district, sector: inward ? `${district} ${inward[0]}` : district, full };
};
const pickupToOperatingDate = (value) => {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "number" && value > 20000 && value < 80000) {
    const date = new Date(Date.UTC(1899, 11, 30 + Math.floor(value)));
    return date.toISOString().slice(0, 10);
  }
  const text = clean(value);
  if (!text) return "";
  if (/^\d{5}(?:\.\d+)?$/.test(text)) {
    const serial = Number(text);
    if (serial > 20000 && serial < 80000) {
      const date = new Date(Date.UTC(1899, 11, 30 + Math.floor(serial)));
      return date.toISOString().slice(0, 10);
    }
  }
  let year = new Date().getFullYear(), month = 0, day = 0;
  let match = text.match(/(20\d{2})\D+(\d{1,2})\D+(\d{1,2})/);
  if (match) { year = Number(match[1]); month = Number(match[2]); day = Number(match[3]); }
  if (!month) {
    match = text.match(/(?:20\d{2}\D*)?(\d{1,2})\s*月\s*(\d{1,2})\s*(?:日|号)?/);
    if (match) { month = Number(match[1]); day = Number(match[2]); }
  }
  if (!month) {
    match = text.match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
    if (match) {
      const first = Number(match[1]), second = Number(match[2]), parsedYear = Number(match[3]);
      year = parsedYear < 100 ? 2000 + parsedYear : parsedYear;
      if (first > 12) { day = first; month = second; }
      else { month = first; day = second; }
    }
  }
  if (!month || !day || month > 12 || day > 31) return "";
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
};
const addDays = (date, days) => {
  const [year, month, day] = date.split("-").map(Number);
  const next = new Date(Date.UTC(year, month - 1, day + days));
  return `${next.getUTCFullYear()}-${String(next.getUTCMonth() + 1).padStart(2, "0")}-${String(next.getUTCDate()).padStart(2, "0")}`;
};
const platformCounts = () => Object.fromEntries(PLATFORMS.map((platform) => [platform, 0]));
const addPlatformCount = (target, platform, count = 1) => {
  if (PLATFORMS.includes(platform)) target[platform] = (target[platform] || 0) + count;
  return target;
};
const mergePlatformCounts = (...counts) => counts.reduce((total, current) => {
  for (const platform of PLATFORMS) total[platform] += Number(current?.[platform] || 0);
  return total;
}, platformCounts());
const estimatedParcelsFor = (platforms = {}) => Number(platforms.TT || 0) / 50 + Number(platforms.SF || 0) / 50 + Number(platforms.TEMU || 0) / 21;
const roundParcels = (platforms) => Math.round(estimatedParcelsFor(platforms) * 100) / 100;
const retentionCutoffDate = (days = 7) => {
  const now = new Date();
  const start = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate() - days + 1));
  return start.toISOString().slice(0, 10);
};
const platformFromText = (value) => {
  const text = clean(value).toUpperCase();
  if (text.includes("TEMU") || text.includes("特木")) return "TEMU";
  if (text.includes("TIKTOK") || text.includes("TIK TOK") || text.includes("抖音") || /(^|[^A-Z])TT([^A-Z]|$)/.test(text)) return "TT";
  if (text.includes("顺丰") || text.includes("SHUNFENG") || /(^|[^A-Z])SF([^A-Z]|$)/.test(text)) return "SF";
  return "OTHER";
};
const sourceType = (sheet, row) => {
  const source = fieldValue(row, "source");
  if (hasField(row, "source")) return platformFromText(source);
  return platformFromText(sheet);
};
const warehouseType = (sheet, row) => {
  const source = hasField(row, "source") ? fieldValue(row, "source") : sheet;
  return clean(source).toUpperCase().includes("TEMU-CW") ? "CERTIFIED_WAREHOUSE" : "";
};
const pickupOriginCenter = (sheet) => clean(sheet).toUpperCase().includes("PIK") ? "Glasgow Depot" : ORIGIN_CENTER;
const sourcePlatform = (sheet, row) => {
  const explicit = hasField(row, "source") ? clean(fieldValue(row, "source")) : "";
  const value = (explicit || clean(sheet)).toUpperCase();
  if (value.includes("TEMU-CW")) return "TEMU-CW";
  if (value.includes("TEMU-CB")) return "TEMU-CB";
  if (value.includes("TEMU") || value.includes("特木")) return "TEMU";
  if (value.includes("TIKTOK") || value.includes("TIK TOK") || value.includes("抖音") || /(^|[^A-Z])TT([^A-Z]|$)/.test(value)) return explicit ? value : "TT";
  if (value.includes("顺丰") || value.includes("SHUNFENG") || /(^|[^A-Z])SF([^A-Z]|$)/.test(value)) return explicit ? value : "SF";
  return explicit || "OTHER";
};
const addSourcePlatformCount = (target, platform, count = 1) => {
  const key = clean(platform).toUpperCase() || "OTHER";
  target[key] = (target[key] || 0) + count;
  return target;
};
const mergeSourcePlatformCounts = (...counts) => counts.reduce((total, current) => {
  for (const [platform, count] of Object.entries(current || {})) total[platform] = (total[platform] || 0) + Number(count || 0);
  return total;
}, {});
const capacityFor = (region) => region === "South" ? 40 : ["Midlands", "Northwest"].includes(region) ? 70 : null;
const vehicleFor = (center, cubes, settings, date) => settings.vehicleOverrides?.[`${date}:${center}`] || settings.vehicleOverrides?.[center] || "";
const CENTER_SHORT_TO_NAME = Object.fromEntries(CENTERS.map((center) => [center.shortName, center.center]));

async function sha256(value) {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(value));
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}
async function sign(value) {
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(OPERATING_SECRET), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const digest = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(value));
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}
function cookie(request, name) {
  return (request.headers.get("cookie") || "").split(";").map((part) => part.trim()).find((part) => part.startsWith(name + "="))?.split("=").slice(1).join("=") || "";
}
async function isOperatingAdmin(request) {
  const [expires, signature] = cookie(request, "operating_admin").split(".");
  return Number(expires) > Date.now() && signature === await sign(expires);
}
async function sessionCookie(request) {
  const expires = String(Date.now() + 8 * 60 * 60 * 1000);
  const secure = new URL(request.url).protocol === "https:" ? "; Secure" : "";
  return `operating_admin=${expires}.${await sign(expires)}; HttpOnly; SameSite=Strict; Path=/; Max-Age=28800${secure}`;
}
async function settings(store) {
  return { ...DEFAULT_SETTINGS, ...await store.get("settings.json", { type: "json" }).catch(() => null) || {} };
}
async function differenceReasonsForDate(store, date) {
  const legacy = await store.get(`difference-reasons/${date}.json`, { type: "json" }).catch(() => null);
  const reasons = { ...(legacy?.reasons || {}) };
  const prefix = `difference-reason-cells/${date}/`;
  const { blobs } = await store.list({ prefix });
  for (const blob of blobs) {
    const cell = await store.get(blob.key, { type: "json" }).catch(() => null);
    if (cell?.key && cell.value) reasons[cell.key] = cell.value;
  }
  return reasons;
}
async function vehicleOverridesForDate(store, date) {
  const overrides = {};
  const prefix = `vehicle-override-cells/${date}/`;
  const { blobs } = await store.list({ prefix });
  for (const blob of blobs) {
    const cell = await store.get(blob.key, { type: "json" }).catch(() => null);
    if (cell?.center && cell.vehicle) overrides[`${date}:${cell.center}`] = cell.vehicle;
  }
  return overrides;
}
async function routeMap(store) {
  const map = await store.get("route-map.json", { type: "json" }).catch(() => null) || { updatedAt: null, count: 0, entries: {} };
  const postcodeIndex = {};
  for (const entry of Object.values(map.entries || {})) {
    if (!entry?.enabled) continue;
    const postcode = clean(entry.postcodeOutward).toUpperCase().replace(/\s/g, "");
    (postcodeIndex[postcode] ||= []).push(entry);
  }
  for (const [postcode, entries] of Object.entries(postcodeIndex)) {
    postcodeIndex[postcode] = entries.slice().sort((a, b) => clean(b.updatedAt).localeCompare(clean(a.updatedAt)));
  }
  return { ...map, postcodeIndex };
}
function routeMapKey(postcodeOutward, route) {
  return `${clean(postcodeOutward).toUpperCase().replace(/\s/g, "")}\u0001${normalizeRouteCode(route) || clean(route).toUpperCase()}`;
}
const routeMapSummary = (map) => ({ updatedAt: map.updatedAt || null, count: Object.keys(map.entries || {}).length });
function parseRouteMapWorkbook(book, existingMap) {
  const updatedAt = new Date().toISOString();
  const sourceUploadId = crypto.randomUUID();
  const next = { updatedAt, count: 0, entries: { ...(existingMap.entries || {}) } };
  let rows = 0, enabled = 0, skipped = 0, duplicates = 0, updated = 0;
  const seen = new Set();
  for (const sheetName of book.SheetNames) {
    for (const row of XLSX.utils.sheet_to_json(book.Sheets[sheetName], { defval: "", raw: false })) {
      rows += 1;
      if (clean(fieldValue(row, "status")) && clean(fieldValue(row, "status")) !== "启用") { skipped += 1; continue; }
      const postcodeOutward = clean(fieldValue(row, "receiverPostcode")).toUpperCase().replace(/\s/g, "");
      const route = normalizeRouteCode(fieldValue(row, "routeCode"));
      const outletCode = clean(fieldValue(row, "outletCode")).toUpperCase() || codeFromSegment(route);
      const outletName = clean(fieldValue(row, "outlet")).toUpperCase() || CODE_TO_OUTLET[outletCode] || "Unknown Outlet";
      const centerName = clean(fieldValue(row, "depot")) || centerFromCode(outletCode);
      if (!postcodeOutward || !route || !outletCode || !outletName) { skipped += 1; continue; }
      const key = routeMapKey(postcodeOutward, route);
      if (seen.has(key)) duplicates += 1;
      seen.add(key);
      enabled += 1; updated += 1;
      for (const existingKey of Object.keys(next.entries)) {
        if (clean(next.entries[existingKey]?.postcodeOutward).toUpperCase().replace(/\s/g, "") === postcodeOutward && existingKey !== key) delete next.entries[existingKey];
      }
      next.entries[key] = { postcodeOutward, routeCode: route, outletCode, outletName, centerCode: clean(fieldValue(row, "depotCode")), centerName, enabled: true, updatedAt, sourceUploadId };
    }
  }
  next.count = Object.keys(next.entries).length;
  return { map: next, summary: { rows, enabled, skipped, duplicates, updated, count: next.count, updatedAt } };
}
function resolveRoute(row, postal, map, sheetName) {
  const normalizedRoute = normalizeRouteCode(fieldValue(row, "routeCode"));
  let correctedRoute = false;
  const candidates = (map.postcodeIndex?.[postal.district] || Object.values(map.entries || {}).filter((entry) => entry?.enabled && entry.postcodeOutward === postal.district)).filter((entry) => entry?.enabled);
  const mapped = candidates[0] || null;
  correctedRoute = Boolean(mapped?.routeCode && mapped.routeCode !== normalizedRoute);
  const source = { sourceType: sourceType(sheetName, row), warehouseType: warehouseType(sheetName, row), sourcePlatform: sourcePlatform(sheetName, row), originCenter: pickupOriginCenter(sheetName) };
  if (mapped?.enabled) return { matched: true, correctedRoute, originalRouteCode: correctedRoute ? normalizedRoute : "", code: mapped.outletCode, routeCode: mapped.routeCode, region: regionFrom(mapped.outletCode), outlet: mapped.outletName, destinationCenter: mapped.centerName || centerFromCode(mapped.outletCode), destinationCenterCode: mapped.centerCode || "", ...source };
  return { matched: false, code: "", routeCode: normalizedRoute, region: "Unmapped", outlet: "未映射", destinationCenter: "Unmapped", ...source };
}
function parseOrderRows(book, map) {
  const rowsByDate = new Map(), missingByDate = new Map(), routeMissingByDate = new Map(), routeMissingKeysByDate = new Map(), routeMatchedByDate = new Map(), detailSheets = [], skippedSheets = [];
  let missingPickupDates = 0, rowCount = 0;
  for (const sheetName of book.SheetNames) {
    const records = XLSX.utils.sheet_to_json(book.Sheets[sheetName], { defval: "", raw: false });
    const first = records[0] || {};
    const hasPickupDate = hasField(first, "pickupDate");
    const hasOrderPostcode = hasField(first, "receiverPostcode") || hasField(first, "adjustedPostcode");
    if (!hasPickupDate || !hasOrderPostcode) { skippedSheets.push(sheetName); continue; }
    detailSheets.push(sheetName);
    records.forEach((row, index) => {
      const date = pickupToOperatingDate(fieldValue(row, "pickupDate"));
      if (!date) { missingPickupDates += 1; return; }
      const postal = postcodeParts(fieldValue(row, "receiverPostcode"), fieldValue(row, "adjustedPostcode"));
      if (postal.full === "Unknown") missingByDate.set(date, (missingByDate.get(date) || 0) + 1);
      const resolved = resolveRoute(row, postal, map, sheetName);
      if (resolved.matched) routeMatchedByDate.set(date, (routeMatchedByDate.get(date) || 0) + 1);
      else {
        routeMissingByDate.set(date, (routeMissingByDate.get(date) || 0) + 1);
        const missingKeys = routeMissingKeysByDate.get(date) || new Map();
        const missingKey = `${postal.district} + ${resolved.routeCode || "无二段码"}`;
        missingKeys.set(missingKey, (missingKeys.get(missingKey) || 0) + 1);
        routeMissingKeysByDate.set(date, missingKeys);
      }
      const rows = rowsByDate.get(date) || [];
      rows.push({ waybill: `ROW:${sheetName}:${index + 2}`, source: sheetName, ...resolved, ...postal });
      rowsByDate.set(date, rows); rowCount += 1;
    });
  }
  return { rowsByDate, missingByDate, routeMissingByDate, routeMissingKeysByDate, routeMatchedByDate, detailSheets, skippedSheets, missingPickupDates, rowCount };
}
function headerFieldIndexes(values) {
  const headers = new Map(Array.from(values.entries(), ([index, value]) => [normalizeHeader(excelCellValue(value)), index]));
  return Object.fromEntries(Object.entries(FIELD_ALIASES).map(([field, aliases]) => [field, aliases.map(normalizeHeader).map((alias) => headers.get(alias)).find((index) => index != null) ?? -1]));
}
function compactOrderRow(values, indexes) {
  const row = {};
  for (const field of ["pickupDate", "routeCode", "receiverPostcode", "adjustedPostcode", "source"]) {
    if (indexes[field] >= 0) row[FIELD_ALIASES[field][0]] = excelCellValue(values[indexes[field]]);
  }
  return row;
}
function collectOrderRow(parsed, sheetName, row, index, map) {
  const date = pickupToOperatingDate(fieldValue(row, "pickupDate"));
  if (!date) { parsed.missingPickupDates += 1; return; }
  const postal = postcodeParts(fieldValue(row, "receiverPostcode"), fieldValue(row, "adjustedPostcode"));
  if (postal.full === "Unknown") parsed.missingByDate.set(date, (parsed.missingByDate.get(date) || 0) + 1);
  const resolved = resolveRoute(row, postal, map, sheetName);
  if (resolved.matched) parsed.routeMatchedByDate.set(date, (parsed.routeMatchedByDate.get(date) || 0) + 1);
  else {
    parsed.routeMissingByDate.set(date, (parsed.routeMissingByDate.get(date) || 0) + 1);
    const missingKeys = parsed.routeMissingKeysByDate.get(date) || new Map();
    const missingKey = `${postal.district} + ${resolved.routeCode || "无二段码"}`;
    missingKeys.set(missingKey, (missingKeys.get(missingKey) || 0) + 1);
    parsed.routeMissingKeysByDate.set(date, missingKeys);
  }
  const rows = parsed.rowsByDate.get(date) || [];
  rows.push({ waybill: `ROW:${sheetName}:${index}`, source: sheetName, ...resolved, ...postal });
  parsed.rowsByDate.set(date, rows);
  parsed.rowCount += 1;
}
async function parseOrderFile(file, map) {
  const parsed = { rowsByDate: new Map(), missingByDate: new Map(), routeMissingByDate: new Map(), routeMissingKeysByDate: new Map(), routeMatchedByDate: new Map(), detailSheets: [], skippedSheets: [], missingPickupDates: 0, rowCount: 0 };
  const cbtPickup = new Map();
  const stream = Readable.fromWeb(file.stream());
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(stream, { entries: "emit", sharedStrings: "cache", styles: "ignore", hyperlinks: "ignore", worksheets: "emit" });
  for await (const worksheet of reader) {
    let indexes = null;
    let detail = false;
    const isCbt = clean(worksheet.name).toUpperCase() === "CBT";
    for await (const excelRow of worksheet) {
      const values = excelRow.values || [];
      if (!indexes) {
        indexes = headerFieldIndexes(values);
        // Route code is optional: postcode + the maintained route map are authoritative.
        detail = indexes.pickupDate >= 0 && (indexes.receiverPostcode >= 0 || indexes.adjustedPostcode >= 0);
        if (detail) parsed.detailSheets.push(worksheet.name);
        else if (!isCbt) parsed.skippedSheets.push(worksheet.name);
        continue;
      }
      if (detail && (indexes.receiverPostcode < 0 || clean(excelCellValue(values[indexes.receiverPostcode])) || indexes.adjustedPostcode < 0 || clean(excelCellValue(values[indexes.adjustedPostcode])))) collectOrderRow(parsed, worksheet.name, compactOrderRow(values, indexes), excelRow.number, map);
      if (isCbt) {
        const date = pickupToOperatingDate(excelCellValue(values[1]));
        if (!date) continue;
        const centers = cbtPickup.get(date) || new Map();
        for (let column = 2; column < values.length; column += 2) {
          const center = CENTER_SHORT_TO_NAME[clean(excelCellValue(values[column])).toUpperCase()];
          const value = Number(clean(excelCellValue(values[column + 1])).replace(/,/g, ""));
          if (center && Number.isFinite(value)) centers.set(center, (centers.get(center) || 0) + value);
        }
        cbtPickup.set(date, centers);
      }
    }
  }
  return { parsed, cbtPickup };
}
function summaryTargetDate(book) {
  const sheetName = book.SheetNames.find((name) => name.includes("提货订单汇总"));
  if (!sheetName) return "";
  const rows = XLSX.utils.sheet_to_json(book.Sheets[sheetName], { header: 1, defval: "", raw: false, blankrows: false });
  return pickupToOperatingDate((rows[0] || []).map(clean).join(" "));
}
function parseCbtPickup(book) {
  if (!book.Sheets.CBT) return new Map();
  const rows = XLSX.utils.sheet_to_json(book.Sheets.CBT, { header: 1, defval: "", raw: false, blankrows: false });
  const byDate = new Map();
  for (const row of rows.slice(1)) {
    const date = pickupToOperatingDate(row[0]);
    if (!date) continue;
    const centers = byDate.get(date) || new Map();
    for (let index = 1; index < row.length; index += 2) {
      const center = CENTER_SHORT_TO_NAME[clean(row[index]).toUpperCase()];
      const value = Number(clean(row[index + 1]).replace(/,/g, ""));
      if (center && Number.isFinite(value)) centers.set(center, (centers.get(center) || 0) + value);
    }
    byDate.set(date, centers);
  }
  return byDate;
}
function aggregateVolume(date, rows, existing, importedAt, qualityPatch = {}) {
  const volumeRows = rows;
  const outlets = new Map(), fullPostcodeMap = new Map();
  const addCount = (row, metric, count = 1) => {
    if (!count) return;
    const oKey = [row.region, row.code, row.outlet].join("\u0001");
    const outlet = outlets.get(oKey) || { region: row.region, code: row.code, outlet: row.outlet, due: 0, backlog: 0, total: 0, platforms: platformCounts(), sourcePlatforms: {}, capacity: capacityFor(row.region), drivers: null, backlogShare: 0 };
    outlet[metric] += count;
    outlet.total += count;
    if (metric === "due") { addPlatformCount(outlet.platforms, row.sourceType, count); addPlatformCount(outlet.platforms, row.warehouseType, count); addSourcePlatformCount(outlet.sourcePlatforms, row.sourcePlatform || row.sourceType, count); }
    outlets.set(oKey, outlet);
  };
  const addFullPostcodeCount = (row, metric, count = 1) => {
    if (!count) return;
    const key = [row.region, row.code, row.outlet, row.full].join("\u0001");
    const bucket = fullPostcodeMap.get(key) || { region: row.region, code: row.code, outlet: row.outlet, postal: row.full, due: 0, backlog: 0, total: 0 };
    bucket[metric] += count; bucket.total += count;
    if (row.detailMissing || row.full === "Postcode detail missing / 邮编明细缺失") bucket.detailMissing = true;
    fullPostcodeMap.set(key, bucket);
  };
  const add = (row, metric) => {
    addCount(row, metric);
    addFullPostcodeCount(row, metric);
  };
  volumeRows.forEach((row) => add(row, "due"));
  for (const row of existing?.outlets || []) addCount(row, "backlog", row.backlog || 0);
  for (const row of existing?.views?.["Full Postcode"] || []) addFullPostcodeCount({ ...row, full: row.postal }, "backlog", row.backlog || 0);
  const outletRows = [...outlets.values()].map((row) => ({ ...row, drivers: row.capacity ? Math.ceil(row.total / row.capacity) : null, backlogShare: row.total ? row.backlog / row.total : 0 }));
  const views = { District: [], Area: [], Sector: [], "Full Postcode": completePostcodeCoverage(outletRows, [...fullPostcodeMap.values()]) };
  const regions = ["South", "Midlands", "Northwest", "North", "Unmapped"].map((region) => {
    const rows = outletRows.filter((row) => row.region === region);
    return rows.length ? { region, due: rows.reduce((s, r) => s + r.due, 0), backlog: rows.reduce((s, r) => s + r.backlog, 0), total: rows.reduce((s, r) => s + r.total, 0), platforms: mergePlatformCounts(...rows.map((row) => row.platforms)), drivers: rows.reduce((s, r) => s + (r.drivers || 0), 0), outlets: rows.length } : null;
  }).filter(Boolean);
  const due = volumeRows.length, backlog = outletRows.reduce((s, r) => s + r.backlog, 0);
  return { schemaVersion: SCHEMA_VERSION, date, importedAt, backlogEffectiveDate: existing?.backlogEffectiveDate || null, updatedAt: { inbound: importedAt, backlog: existing?.updatedAt?.backlog || existing?.importedAt || null }, sourceCounts: { inbound: due, backlog }, quality: { inboundDuplicates: 0, backlogDuplicates: existing?.quality?.backlogDuplicates || 0, inboundMissingPostcodes: qualityPatch.missingPostcodes || 0, backlogMissingPostcodes: existing?.quality?.backlogMissingPostcodes || 0, missingPostcodes: (qualityPatch.missingPostcodes || 0) + (existing?.quality?.backlogMissingPostcodes || 0), missingPickupDates: qualityPatch.missingPickupDates || 0, routeMapMatched: qualityPatch.routeMapMatched || 0, routeMapMissing: qualityPatch.routeMapMissing || 0, routeMapMissingExamples: qualityPatch.routeMapMissingExamples || [], routeMapDuplicates: qualityPatch.routeMapDuplicates || 0, routeMapUpdated: qualityPatch.routeMapUpdated || null }, totals: { due, backlog, total: due + backlog, platforms: mergePlatformCounts(...outletRows.map((row) => row.platforms)), sourcePlatforms: mergeSourcePlatformCounts(...outletRows.map((row) => row.sourcePlatforms)), drivers: outletRows.reduce((s, r) => s + (r.drivers || 0), 0), outlets: outletRows.length }, regions, outlets: outletRows, views };
}
function completePostcodeCoverage(outlets, fullRows) {
  const rows = (fullRows || []).map((row) => ({ ...row }));
  const totals = new Map();
  const keyFor = (row) => [row.region, row.code, row.outlet].join("\u0001");
  for (const row of rows) {
    const key = keyFor(row);
    const current = totals.get(key) || { due: 0, backlog: 0 };
    current.due += Number(row.due || 0);
    current.backlog += Number(row.backlog || 0);
    totals.set(key, current);
  }
  for (const outlet of outlets || []) {
    const current = totals.get(keyFor(outlet)) || { due: 0, backlog: 0 };
    const due = Number(outlet.due || 0) - current.due;
    const backlog = Number(outlet.backlog || 0) - current.backlog;
    if (due <= 0 && backlog <= 0) continue;
    rows.push({
      region: outlet.region,
      code: outlet.code,
      outlet: outlet.outlet,
      postal: "Postcode detail missing / 邮编明细缺失",
      due: Math.max(0, due),
      backlog: Math.max(0, backlog),
      total: Math.max(0, due) + Math.max(0, backlog),
      detailMissing: true,
    });
  }
  return rows.sort((a, b) => b.total - a.total);
}
function aggregateOperating(date, rows, appSettings, quality, importedAt, cbtPickupByCenter = new Map(), map = null) {
  const centerList = centerDefinitions(map, rows);
  const centerMetrics = new Map(centerList.map((center) => [center.center, { outbound: 0, inbound: 0, sameCenter: 0, cbtPickup: cbtPickupByCenter.get(center.center) || 0 }]));
  const outletMetrics = new Map();
  for (const row of rows) {
    const destination = row.destinationCenter || centerFromCode(row.code);
    if (destination === "Unmapped") continue;
    const origin = row.originCenter || ORIGIN_CENTER;
    const originMetrics = centerMetrics.get(origin) || { outbound: 0, inbound: 0, sameCenter: 0, cbtPickup: 0 };
    const destinationMetrics = centerMetrics.get(destination) || { outbound: 0, inbound: 0, sameCenter: 0, cbtPickup: 0 };
    if (origin === destination) destinationMetrics.sameCenter += 1;
    else { originMetrics.outbound += 1; destinationMetrics.inbound += 1; }
    centerMetrics.set(origin, originMetrics);
    centerMetrics.set(destination, destinationMetrics);
    const key = `${destination}\u0001${row.code}\u0001${row.outlet}`;
    const outlet = outletMetrics.get(key) || { center: destination, outlet: row.outlet, routeCode: row.routeCode || routeCode(row.code), deliveryVolume: 0, types: {}, sourcePlatforms: {} };
    outlet.deliveryVolume += 1; outlet.types[row.sourceType] = (outlet.types[row.sourceType] || 0) + 1; if (row.warehouseType) outlet.types[row.warehouseType] = (outlet.types[row.warehouseType] || 0) + 1; addSourcePlatformCount(outlet.sourcePlatforms, row.sourcePlatform || row.sourceType); outletMetrics.set(key, outlet);
  }
  const outlets = [...outletMetrics.values()].sort((a, b) => a.routeCode.localeCompare(b.routeCode)).map((o) => {
    const platforms = Object.fromEntries(PLATFORMS.map((platform) => [platform, o.types[platform] || 0]));
    return { center: o.center, outlet: o.outlet, routeCode: o.routeCode, deliveryVolume: o.deliveryVolume, platforms, sourcePlatforms: o.sourcePlatforms, estimatedParcels: roundParcels(platforms) };
  });
  const centers = centerDefinitions(map, rows).map((center) => {
    const metrics = centerMetrics.get(center.center) || { outbound: 0, inbound: 0, sameCenter: 0, cbtPickup: 0 };
    const centerOutlets = outlets.filter((o) => o.center === center.center);
    const platforms = mergePlatformCounts(...centerOutlets.map((outlet) => outlet.platforms));
    const estimatedParcels = Math.round(estimatedParcelsFor(platforms));
    const operation = metrics.outbound + metrics.inbound + metrics.sameCenter + metrics.cbtPickup;
    const cubeVolume = metrics.inbound || metrics.sameCenter;
    const estimatedCubes = Math.round(cubeVolume * appSettings.cubeFactor * 100) / 100;
    return { center: center.center, shortName: center.shortName, chinese: center.chinese, operation, outbound: metrics.outbound, inbound: metrics.inbound, sameCenter: metrics.sameCenter, cbtPickup: metrics.cbtPickup, platforms, estimatedParcels, estimatedCubes, vehicle: vehicleFor(center.center, estimatedCubes, appSettings, date), outletCount: centerOutlets.length };
  });
  const totalPlatforms = mergePlatformCounts(...outlets.map((outlet) => outlet.platforms));
  const totals = { operation: centers.reduce((s, r) => s + r.operation, 0), outbound: centers.reduce((s, r) => s + r.outbound, 0), inbound: centers.reduce((s, r) => s + r.inbound, 0), sameCenter: centers.reduce((s, r) => s + r.sameCenter, 0), cbtPickup: centers.reduce((s, r) => s + r.cbtPickup, 0), platforms: totalPlatforms, sourcePlatforms: mergeSourcePlatformCounts(...outlets.map((outlet) => outlet.sourcePlatforms)), estimatedParcels: Math.round(estimatedParcelsFor(totalPlatforms)), estimatedCubes: Math.round(centers.reduce((s, r) => s + r.estimatedCubes, 0) * 100) / 100, outletCount: outlets.length, centers: centers.length };
  return { schemaVersion: SCHEMA_VERSION, date, importedAt, title: "提货单量预测", settings: appSettings, totals, centers, outlets, quality };
}
function alignOutletDatasets(operating, volume) {
  const operatingOutlets = (operating.outlets || []).filter((outlet) => Number(outlet.deliveryVolume || 0) > 0).map((outlet) => ({ ...outlet, platforms: { ...platformCounts(), ...(outlet.platforms || {}) } }));
  const volumeOutlets = (volume.outlets || [])
    .filter((outlet) => outlet.code && outlet.region !== "Unmapped" && outlet.outlet !== "未映射" && outlet.outlet !== "Unknown Outlet")
    .map((outlet) => ({ ...outlet, platforms: { ...platformCounts(), ...(outlet.platforms || {}) } }));
  const operatingKeys = new Set(operatingOutlets.map((outlet) => `${codeFromSegment(outlet.routeCode)}\u0001${outlet.outlet}`));
  const volumeKeys = new Set(volumeOutlets.map((outlet) => `${outlet.code}\u0001${outlet.outlet}`));

  for (const outlet of operatingOutlets) {
    const code = codeFromSegment(outlet.routeCode);
    const key = `${code}\u0001${outlet.outlet}`;
    if (!code || volumeKeys.has(key)) continue;
    const due = Number(outlet.deliveryVolume || 0);
    const region = regionFrom(code);
    const capacity = capacityFor(region);
    volumeOutlets.push({
      region, code, outlet: outlet.outlet, due, backlog: 0, total: due,
      platforms: { ...platformCounts(), ...(outlet.platforms || {}) }, capacity,
      drivers: due && capacity ? Math.ceil(due / capacity) : 0, backlogShare: 0,
    });
    volumeKeys.add(key);
  }

  for (const outlet of volumeOutlets) {
    const key = `${outlet.code}\u0001${outlet.outlet}`;
    const due = Number(outlet.due || 0);
    if (!outlet.code || !due || operatingKeys.has(key)) continue;
    const platforms = { ...platformCounts(), ...(outlet.platforms || {}) };
    operatingOutlets.push({
      center: centerFromCode(outlet.code), outlet: outlet.outlet, routeCode: routeCode(outlet.code),
      deliveryVolume: due, platforms, estimatedParcels: roundParcels(platforms),
    });
    operatingKeys.add(key);
  }

  operatingOutlets.sort((a, b) => a.routeCode.localeCompare(b.routeCode) || a.outlet.localeCompare(b.outlet));
  volumeOutlets.sort((a, b) => a.code.localeCompare(b.code) || a.outlet.localeCompare(b.outlet));
  const volumeDue = volumeOutlets.reduce((sum, row) => sum + Number(row.due || 0), 0);
  const volumeBacklog = volumeOutlets.reduce((sum, row) => sum + Number(row.backlog || 0), 0);
  const volumeDrivers = volumeOutlets.reduce((sum, row) => sum + Number(row.drivers || 0), 0);
  const volumePlatforms = mergePlatformCounts(...volumeOutlets.map((row) => row.platforms));
  const centers = (operating.centers || []).map((center) => {
    const centerOutlets = operatingOutlets.filter((outlet) => outlet.center === center.center);
    return { ...center, outletCount: centerOutlets.length };
  });
  const regions = ["South", "Midlands", "Northwest", "North", "Unmapped"].map((region) => {
    const rows = volumeOutlets.filter((outlet) => outlet.region === region);
    if (!rows.length) return null;
    return {
      region,
      due: rows.reduce((sum, row) => sum + Number(row.due || 0), 0),
      backlog: rows.reduce((sum, row) => sum + Number(row.backlog || 0), 0),
      total: rows.reduce((sum, row) => sum + Number(row.total || 0), 0),
      platforms: mergePlatformCounts(...rows.map((row) => row.platforms)),
      drivers: rows.reduce((sum, row) => sum + Number(row.drivers || 0), 0),
      outlets: rows.length,
    };
  }).filter(Boolean);
  return {
    operating: { ...operating, centers, outlets: operatingOutlets, totals: { ...operating.totals, outletCount: operatingOutlets.length } },
    volume: {
      ...volume,
      sourceCounts: { ...(volume.sourceCounts || {}), inbound: volumeDue },
      regions,
      outlets: volumeOutlets,
      views: { ...(volume.views || {}), "Full Postcode": completePostcodeCoverage(volumeOutlets, volume.views?.["Full Postcode"] || []) },
      totals: { ...volume.totals, due: volumeDue, backlog: volumeBacklog, total: volumeDue + volumeBacklog, platforms: volumePlatforms, drivers: volumeDrivers, outlets: volumeOutlets.length },
    },
  };
}
function hasMappableVolume(dataset) {
  return Number(dataset?.totals?.due || 0) > 0;
}
function expandPackedPostcodes(volume) {
  if (!Array.isArray(volume?.fullPostcodesPacked)) return volume;
  const full = volume.fullPostcodesPacked.map(([region, code, outlet, postal, due, backlog, detailMissing]) => ({
    region, code, outlet, postal,
    due: Number(due || 0),
    backlog: Number(backlog || 0),
    total: Number(due || 0) + Number(backlog || 0),
    ...(detailMissing ? { detailMissing: true } : {}),
  }));
  const { fullPostcodesPacked, ...rest } = volume;
  return { ...rest, views: { ...(volume.views || {}), "Full Postcode": full } };
}
function packPostcodes(volume) {
  const fullPostcodesPacked = (volume?.views?.["Full Postcode"] || []).map((row) => [
    row.region, row.code, row.outlet, row.postal, Number(row.due || 0), Number(row.backlog || 0), row.detailMissing ? 1 : 0,
  ]);
  return { ...volume, views: { ...(volume.views || {}), "Full Postcode": [] }, fullPostcodesPacked };
}
async function volumeFetch(path, options = {}) {
  const response = await fetch(VOLUME_BASE + path, options);
  if (!response.ok) throw new Error(`Volume API ${path} failed: ${response.status}`);
  return response;
}
async function datasetFetch(date, kind = "forecast") {
  const query = kind === "pickup" ? "?kind=pickup&full=1" : "?full=1";
  return fetch(`${VOLUME_BASE}/api/datasets/${date}${query}`).then((r) => r.ok ? r.json() : undefined);
}
async function syncDataset(date, dataset, volumeCookie, kind = "forecast") {
  const query = kind === "pickup" ? "?kind=pickup&side=inbound" : "?side=inbound";
  const postcodeCounts = new Map();
  for (const row of dataset?.views?.["Full Postcode"] || []) {
    const count = Number(row.due || 0);
    if (count <= 0) continue;
    const postcode = String(row.postal || "Unknown");
    postcodeCounts.set(postcode, (postcodeCounts.get(postcode) || 0) + count);
  }
  const compact = {
    date,
    importedAt: dataset.importedAt,
    updatedAt: dataset.updatedAt,
    sourceCounts: dataset.sourceCounts,
    quality: dataset.quality,
    outlets: dataset.outlets,
    totals: { due: Number(dataset.totals?.due || 0), backlog: 0, platforms: dataset.totals?.platforms, sourcePlatforms: dataset.totals?.sourcePlatforms },
    sidePostcodesPacked: [...postcodeCounts],
  };
  await volumeFetch(`/api/datasets/${date}${query}`, { method: "PUT", headers: { "content-type": "application/json", cookie: volumeCookie }, body: JSON.stringify(compact) });
}
async function clearSyncedDataset(sourceDate, kind, volumeCookie) {
  const date = addDays(sourceDate, 1);
  const pickup = kind === "actual";
  const current = await datasetFetch(date, pickup ? "pickup" : "forecast");
  if (!current) return { sourceDate, date, kind, cleared: false };
  const importedAt = new Date().toISOString();
  const empty = {
    date,
    importedAt,
    updatedAt: { ...(current.updatedAt || {}), inbound: importedAt },
    sourceCounts: { ...(current.sourceCounts || {}), inbound: 0 },
    quality: {
      ...(current.quality || {}),
      inboundDuplicates: 0,
      inboundMissingPostcodes: 0,
      routeMapMatched: 0,
      routeMapMissing: 0,
      routeMapMissingExamples: [],
    },
    outlets: [],
    totals: { due: 0, backlog: 0, total: 0, platforms: platformCounts(), sourcePlatforms: {} },
    sidePostcodesPacked: [],
  };
  const query = pickup ? "?kind=pickup&side=inbound" : "?side=inbound";
  await volumeFetch(`/api/datasets/${date}${query}`, {
    method: "PUT",
    headers: { "content-type": "application/json", cookie: volumeCookie },
    body: JSON.stringify(empty),
  });
  return { sourceDate, date, kind, cleared: true };
}
async function loginVolume() {
  if (!process.env.VOLUME_ADMIN_PASSWORD) throw new Error("VOLUME_ADMIN_PASSWORD is not configured");
  const response = await volumeFetch("/api/admin/login", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ password: process.env.VOLUME_ADMIN_PASSWORD }) });
  return response.headers.get("set-cookie") || "";
}
function normalizeOperatingDataset(item, currentSettings = null) {
  const appSettings = {
    ...DEFAULT_SETTINGS,
    ...(item.settings || {}),
    ...(currentSettings || {}),
    vehicleOverrides: {
      ...(item.settings?.vehicleOverrides || {}),
      ...(currentSettings?.vehicleOverrides || {}),
    },
  };
  const platformsAvailable = Boolean(item.totals?.platforms);
  if ((item.schemaVersion || 0) >= SCHEMA_VERSION) {
    const centersWithVehicles = (item.centers || []).map((center) => ({
      ...center,
      vehicle: vehicleFor(center.center, center.estimatedCubes || 0, appSettings, item.date),
    }));
    if (!platformsAvailable) return { ...item, centers: centersWithVehicles, platformsAvailable: false };
    const outlets = (item.outlets || []).filter((outlet) => Number(outlet.deliveryVolume || 0) > 0).map((outlet) => {
      const platforms = { ...platformCounts(), ...(outlet.platforms || {}) };
      return { ...outlet, platforms, estimatedParcels: roundParcels(platforms) };
    });
    const centers = centersWithVehicles.map((center) => {
      const centerOutlets = outlets.filter((outlet) => outlet.center === center.center);
      const platforms = mergePlatformCounts(...centerOutlets.map((outlet) => outlet.platforms));
      return { ...center, platforms, estimatedParcels: Math.round(estimatedParcelsFor(platforms)) };
    });
    const platforms = mergePlatformCounts(...outlets.map((outlet) => outlet.platforms));
    return { ...item, platformsAvailable: true, outlets, centers, totals: { ...item.totals, outletCount: outlets.length, platforms, estimatedParcels: Math.round(estimatedParcelsFor(platforms)) } };
  }
  const centers = (item.centers || []).map((center) => {
    const outbound = center.center === ORIGIN_CENTER ? center.outbound + center.sameCenter : center.outbound;
    const cbtPickup = center.cbtPickup || Math.max(0, center.operation - outbound - center.inbound);
    return { ...center, outbound, cbtPickup, platforms: center.platforms || platformCounts(), vehicle: vehicleFor(center.center, center.estimatedCubes || 0, appSettings, item.date) };
  });
  return {
    ...item,
    platformsAvailable: false,
    centers,
    outlets: (item.outlets || []).filter((outlet) => Number(outlet.deliveryVolume || 0) > 0).map((outlet) => ({ ...outlet, platforms: outlet.platforms || platformCounts() })),
    totals: {
      ...item.totals,
      outbound: centers.reduce((sum, center) => sum + center.outbound, 0),
      cbtPickup: centers.reduce((sum, center) => sum + center.cbtPickup, 0),
      platforms: item.totals?.platforms || platformCounts(),
    },
  };
}
const platformComparison = (forecast = platformCounts(), actual = null) => Object.fromEntries(PLATFORMS.map((platform) => {
  const forecastValue = Number(forecast?.[platform] || 0);
  const actualValue = actual ? Number(actual?.[platform] || 0) : null;
  return [platform, { forecast: forecastValue, actual: actualValue, diff: actualValue == null ? null : actualValue - forecastValue, completionRate: actualValue == null || !forecastValue ? null : actualValue / forecastValue }];
}));
const legacySourcePlatforms = (platforms = {}) => {
  const certified = Number(platforms.CERTIFIED_WAREHOUSE || 0);
  const result = { SF: Number(platforms.SF || 0), TT: Number(platforms.TT || 0) };
  const temuCb = Math.max(0, Number(platforms.TEMU || 0) - certified);
  if (temuCb) result["TEMU-CB"] = temuCb;
  if (certified) result["TEMU-CW"] = certified;
  if (platforms.OTHER) result.OTHER = Number(platforms.OTHER);
  return result;
};
const sourcePlatformGroup = (platform) => {
  const value = clean(platform).toUpperCase();
  if (value.includes("TEMU-CW")) return "TEMU-CW";
  if (value.includes("TEMU")) return "TEMU-CB";
  const type = platformFromText(value);
  if (type === "SF" || type === "TT") return type;
  return value || "OTHER";
};
const groupedSourcePlatforms = (platforms = {}) => {
  const result = {};
  for (const [platform, count] of Object.entries(platforms || {})) {
    const group = sourcePlatformGroup(platform);
    result[group] = (result[group] || 0) + Number(count || 0);
  }
  return result;
};
const sourcePlatformComparison = (forecast = {}, actual = null) => {
  const groupedForecast = groupedSourcePlatforms(forecast);
  const groupedActual = actual ? groupedSourcePlatforms(actual) : null;
  const preferred = ["SF", "TT", "TEMU-CB", "TEMU-CW"];
  const keys = [...new Set([...preferred, ...Object.keys(groupedForecast), ...Object.keys(groupedActual || {})])].filter((key) => Number(groupedForecast[key] || 0) || Number(groupedActual?.[key] || 0));
  return Object.fromEntries(keys.map((platform) => {
    const forecastValue = Number(groupedForecast[platform] || 0);
    const actualValue = groupedActual ? Number(groupedActual[platform] || 0) : null;
    return [platform, { forecast: forecastValue, actual: actualValue, diff: actualValue == null ? null : actualValue - forecastValue, completionRate: actualValue == null || !forecastValue ? null : actualValue / forecastValue }];
  }));
};
function buildComparison(rawItem, pickup, currentSettings = null) {
  const item = normalizeOperatingDataset(rawItem, currentSettings);
  const pickupDate = item.date;
  const pickupPlatformsAvailable = pickup?.platformsAvailable !== false && Boolean(pickup?.totals?.platforms);
  const forecastByKey = new Map(item.outlets.map((outlet) => [`${codeFromSegment(outlet.routeCode)}\u0001${outlet.outlet}`, outlet]));
  const actualRows = (pickup?.outlets || []).filter((outlet) => (outlet.due || 0) > 0);
  const actualByKey = new Map(actualRows.map((outlet) => [`${outlet.code || ""}\u0001${outlet.outlet}`, outlet]));
  const comparisonOutlets = [];
  for (const [key, forecast] of forecastByKey) {
    const actual = actualByKey.get(key);
    const actualDue = pickup ? actual?.due || 0 : null;
    comparisonOutlets.push({
      ...forecast,
      region: regionFrom(codeFromSegment(forecast.routeCode)),
      comparison: {
        actualDue,
        diff: actualDue == null ? null : actualDue - forecast.deliveryVolume,
        completionRate: actualDue == null || !forecast.deliveryVolume ? null : actualDue / forecast.deliveryVolume,
        platforms: platformComparison(forecast.platforms, pickupPlatformsAvailable ? actual?.platforms || platformCounts() : null),
      },
    });
    actualByKey.delete(key);
  }
  for (const actual of actualByKey.values()) {
    comparisonOutlets.push({
      center: centerFromCode(actual.code || ""),
      region: actual.region || regionFrom(actual.code || ""),
      outlet: actual.outlet,
      routeCode: actual.code ? routeCode(actual.code) : "-",
      deliveryVolume: 0,
      estimatedParcels: 0,
      platforms: platformCounts(),
      comparison: {
        actualDue: actual.due || 0,
        diff: actual.due || 0,
        completionRate: null,
        platforms: platformComparison(platformCounts(), pickupPlatformsAvailable ? actual.platforms || platformCounts() : null),
      },
    });
  }
  const outlets = item.outlets.map((outlet) => comparisonOutlets.find((row) => row.outlet === outlet.outlet && row.routeCode === outlet.routeCode) || outlet);
  const centers = item.centers.map((center) => {
    const centerOutlets = comparisonOutlets.filter((outlet) => outlet.center === center.center);
    const forecastDue = centerOutlets.reduce((sum, row) => sum + row.deliveryVolume, 0);
    const actualDue = pickup ? centerOutlets.reduce((sum, row) => sum + (row.comparison?.actualDue || 0), 0) : null;
    const platformParcels = centerOutlets.reduce((sum, row) => sum + row.estimatedParcels, 0);
    return { ...center, comparison: { forecastDue, platformParcels, actualDue, diff: actualDue == null ? null : actualDue - forecastDue, completionRate: actualDue == null || !forecastDue ? null : actualDue / forecastDue } };
  });
  const forecastDue = item.outlets.reduce((sum, row) => sum + row.deliveryVolume, 0);
  const actualDue = pickup ? Number(pickup.totals?.due || 0) : null;
  return {
    ...item,
    comparisonDate: pickupDate,
    pickupAvailable: Boolean(pickup),
    forecastPlatformsAvailable: item.platformsAvailable,
    actualPlatformsAvailable: pickupPlatformsAvailable,
    platformsAvailable: item.platformsAvailable && pickupPlatformsAvailable,
    comparisonQuality: {
      forecastRouteMapMissing: item.quality?.routeMapMissing || 0,
      actualRouteMapMissing: pickup?.quality?.routeMapMissing || 0,
      forecastRouteMapMissingExamples: item.quality?.routeMapMissingExamples || [],
      actualRouteMapMissingExamples: pickup?.quality?.routeMapMissingExamples || [],
    },
    platformComparison: sourcePlatformComparison(item.totals.sourcePlatforms || legacySourcePlatforms(item.totals.platforms), pickupPlatformsAvailable ? pickup.totals.sourcePlatforms || legacySourcePlatforms(pickup.totals.platforms) : null),
    comparisonOutlets,
    outlets,
    centers,
    totals: { ...item.totals, comparison: { forecastDue, platformParcels: item.outlets.reduce((sum, row) => sum + row.estimatedParcels, 0), actualDue, diff: actualDue == null ? null : actualDue - forecastDue, completionRate: actualDue == null || !forecastDue ? null : actualDue / forecastDue } },
  };
}
async function appendComparison(item) {
  const pickupDate = item.date;
  const store = getStore("operating", { consistency: "strong" });
  const [pickup, differenceReasons, currentSettings, datedVehicleOverrides] = await Promise.all([
    datasetFetch(addDays(pickupDate, 1), "pickup").catch(() => null),
    differenceReasonsForDate(store, pickupDate),
    settings(store),
    vehicleOverridesForDate(store, pickupDate),
  ]);
  currentSettings.vehicleOverrides = { ...(currentSettings.vehicleOverrides || {}), ...datedVehicleOverrides };
  return { ...buildComparison(item, pickup, currentSettings), differenceReasons };
}

async function route(request) {
  const url = new URL(request.url);
  const path = url.pathname.replace(/^\/(?:\.netlify\/functions\/operating-api|operating\/api)/i, "") || "/";
  if (path === "/admin/session") return json({ authenticated: await isOperatingAdmin(request) });
  if (path === "/admin/logout") return json({ authenticated: false }, 200, { "set-cookie": "operating_admin=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0; Secure" });
  if (path === "/admin/login" && request.method === "POST") {
    const body = await request.json();
    return await sha256(body.password || "") === OPERATING_HASH ? json({ authenticated: true }, 200, { "set-cookie": await sessionCookie(request) }) : json({ error: "Invalid password" }, 401);
  }
  const store = getStore("operating", { consistency: "strong" });
  const vehicleMatch = path.match(/^\/vehicle-overrides\/(\d{4}-\d{2}-\d{2})$/);
  if (vehicleMatch) {
    if (request.method !== "PUT") return json({ error: "Method not allowed" }, 405);
    const date = vehicleMatch[1];
    const body = await request.json().catch(() => ({}));
    const center = clean(body.center);
    const vehicle = clean(body.vehicle).slice(0, 30);
    const map = await routeMap(store);
    if (!centerDefinitions(map).some((item) => item.center === center)) return json({ error: "Unknown transfer centre" }, 400);
    const cellKey = `vehicle-override-cells/${date}/${encodeURIComponent(center)}.json`;
    if (vehicle) await store.setJSON(cellKey, { date, center, vehicle, updatedAt: new Date().toISOString() });
    else await store.delete(cellKey);
    return json({ date, center, vehicle, overridden: Boolean(vehicle) });
  }
  if (path === "/settings") {
    if (request.method === "GET") return json({ settings: await settings(store) });
    if (!await isOperatingAdmin(request)) return json({ error: "Operating authentication required" }, 401);
    const current = await settings(store);
    const next = { ...DEFAULT_SETTINGS, ...await request.json(), vehicleOverrides: current.vehicleOverrides || {} };
    await store.setJSON("settings.json", next);
    return json({ settings: next });
  }
  const reasonMatch = path.match(/^\/difference-reasons\/(\d{4}-\d{2}-\d{2})$/);
  if (reasonMatch) {
    const reasonStore = getStore("operating", { consistency: "strong" });
    const date = reasonMatch[1];
    const cellPrefix = `difference-reason-cells/${date}/`;
    if (request.method === "GET") {
      return json({ date, reasons: await differenceReasonsForDate(reasonStore, date) });
    }
    if (request.method !== "PUT") return json({ error: "Method not allowed" }, 405);
    const body = await request.json().catch(() => ({}));
    if (typeof body.key === "string") {
      const key = body.key.slice(0, 240);
      const value = String(body.value ?? "").slice(0, 500);
      const cellKey = `${cellPrefix}${encodeURIComponent(key)}.json`;
      if (value) await reasonStore.setJSON(cellKey, { date, key, value, updatedAt: new Date().toISOString() });
      else await reasonStore.delete(cellKey);
      return json({ date, key, value });
    }
    const reasons = Object.fromEntries(Object.entries(body.reasons || {}).slice(0, 500).map(([key, value]) => [String(key).slice(0, 240), String(value ?? "").slice(0, 500)]));
    await reasonStore.setJSON(`difference-reasons/${date}.json`, { date, reasons, updatedAt: new Date().toISOString() });
    return json({ date, reasons });
  }
  if (path === "/route-map") {
    if (request.method === "GET") return json({ routeMap: routeMapSummary(await routeMap(store)) });
    if (!await isOperatingAdmin(request)) return json({ error: "Operating authentication required" }, 401);
    const form = await request.formData();
    const file = form.get("file");
    if (!(file instanceof File)) return json({ error: "Missing route map Excel file" }, 400);
    const parsed = parseRouteMapWorkbook(XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false }), await routeMap(store));
    await store.setJSON("route-map.json", parsed.map);
    const volumeCookie = await loginVolume();
    const reclassified = await volumeFetch("/api/route-map/reclassify", { method: "POST", headers: { cookie: volumeCookie } }).then((response) => response.json()).catch(() => ({ changed: [] }));
    return json({ routeMap: routeMapSummary(parsed.map), summary: parsed.summary, reclassified: reclassified.changed || [] });
  }
  if (path === "/route-map/export") {
    if (!await isOperatingAdmin(request)) return json({ error: "Operating authentication required" }, 401);
    if (request.method !== "GET") return json({ error: "Method not allowed" }, 405);
    const map = await routeMap(store);
    return json({ routeMap: { updatedAt: map.updatedAt || null, count: Object.keys(map.entries || {}).length, entries: map.entries || {} } });
  }
  if (path === "/datasets") {
    const { blobs } = await store.list({ prefix: "datasets/" });
    const includeArchived = url.searchParams.get("all") === "1" && await isOperatingAdmin(request);
    const cutoff = retentionCutoffDate();
    const datasets = [];
    for (const blob of blobs) {
      const item = await store.get(blob.key, { type: "json" });
      if (item && (includeArchived || item.date >= cutoff)) {
        const normalized = normalizeOperatingDataset(item);
        datasets.push({ date: normalized.date, importedAt: normalized.importedAt, ...normalized.totals });
      }
    }
    return json({ datasets: datasets.sort((a, b) => b.date.localeCompare(a.date)) });
  }
  if (path === "/clearable-dates") {
    if (!await isOperatingAdmin(request)) return json({ error: "Operating authentication required" }, 401);
    const kind = clean(url.searchParams.get("kind")).toLowerCase() === "actual" ? "actual" : "forecast";
    if (kind === "forecast") {
      const { blobs } = await store.list({ prefix: "datasets/" });
      const dates = blobs.map((blob) => blob.key.match(/^datasets\/(\d{4}-\d{2}-\d{2})\.json$/)?.[1]).filter(Boolean).sort().reverse();
      return json({ kind, dates });
    }
    const response = await volumeFetch("/api/datasets?kind=pickup");
    const data = await response.json();
    const dates = (data.datasets || []).filter((item) => Number(item?.sourceCounts?.inbound ?? item?.totals?.due ?? item?.due ?? 0) > 0).map((item) => addDays(item.date, -1)).sort().reverse();
    return json({ kind, dates: [...new Set(dates)] });
  }
  if (path === "/clear-data" && request.method === "POST") {
    if (!await isOperatingAdmin(request)) return json({ error: "Operating authentication required" }, 401);
    const body = await request.json().catch(() => ({}));
    const kind = clean(body.kind).toLowerCase();
    const date = clean(body.date);
    if (!new Set(["forecast", "actual"]).has(kind)) return json({ error: "Invalid data type" }, 400);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json({ error: "Invalid operation date" }, 400);
    const result = await clearSyncedDataset(date, kind, await loginVolume());
    if (kind === "forecast") {
      await store.delete(`datasets/${date}.json`);
      await store.delete(`difference-reasons/${date}.json`).catch(() => undefined);
    }
    return json(result);
  }
  if (path === "/archive/prune" && request.method === "POST") {
    if (!await isOperatingAdmin(request)) return json({ error: "Operating authentication required" }, 401);
    const body = await request.json().catch(() => ({}));
    const keepDays = Math.max(1, Math.min(60, Number(body.keepDays || 7) || 7));
    const cutoff = retentionCutoffDate(keepDays);
    const { blobs } = await store.list({ prefix: "datasets/" });
    const removed = [];
    const kept = [];
    for (const blob of blobs) {
      const date = blob.key.match(/^datasets\/(\d{4}-\d{2}-\d{2})\.json$/)?.[1];
      if (!date) continue;
      if (date < cutoff) {
        await store.delete(blob.key);
        removed.push(date);
      } else kept.push(date);
    }
    return json({ cutoff, keepDays, removed: removed.sort(), kept: kept.sort() });
  }
  const dateMatch = path.match(/^\/datasets\/(\d{4}-\d{2}-\d{2})$/);
  if (dateMatch) {
    const item = await store.get(`datasets/${dateMatch[1]}.json`, { type: "json" });
    return item ? json(await appendComparison(item)) : json({ error: "Dataset not found" }, 404);
  }
  if (path === "/import" && request.method === "POST") {
    if (!await isOperatingAdmin(request)) return json({ error: "Operating authentication required" }, 401);
    const form = await request.formData();
    const file = form.get("file");
    if (!(file instanceof File)) return json({ error: "Missing Excel file" }, 400);
    const kind = clean(form.get("kind")).toLowerCase() === "actual" ? "actual" : "forecast";
    const selectedSourceDate = clean(form.get("sourceDate"));
    if (selectedSourceDate && !/^\d{4}-\d{2}-\d{2}$/.test(selectedSourceDate)) return json({ error: "Invalid operation date" }, 400);
    const map = await routeMap(store);
    const { parsed, cbtPickup } = await parseOrderFile(file, map);
    const appSettings = await settings(store);
    const importedAt = new Date().toISOString();
    const volumeCookie = await loginVolume();
    const summaries = [];
    if (selectedSourceDate) {
      const selectedRows = [...parsed.rowsByDate.values()].flat();
      parsed.rowsByDate = new Map([[selectedSourceDate, selectedRows]]);
    }
    const importDates = [...parsed.rowsByDate.keys()].sort();
    for (const date of importDates) {
      const rows = parsed.rowsByDate.get(date);
      if (!rows?.length) continue;
      const targetDate = addDays(date, 1);
      // Inbound/due imports must never inherit backlog from another business date.
      // Only preserve the independent backlog side when replacing this exact date.
      const existing = await datasetFetch(targetDate, kind === "actual" ? "pickup" : "forecast");
      const routeMapMissingExamples = [...(parsed.routeMissingKeysByDate.get(date) || new Map()).entries()].sort((a, b) => b[1] - a[1]).slice(0, 12).map(([key, count]) => ({ key, count }));
      const qualityPatch = { missingPostcodes: parsed.missingByDate.get(date) || 0, missingPickupDates: parsed.missingPickupDates, routeMapMatched: parsed.routeMatchedByDate.get(date) || 0, routeMapMissing: parsed.routeMissingByDate.get(date) || 0, routeMapMissingExamples, routeMapDuplicates: 0, routeMapUpdated: map.updatedAt };
      const dataset = aggregateVolume(targetDate, rows, existing, importedAt, qualityPatch);
      if (!hasMappableVolume(dataset)) {
        summaries.push({ date, ...(kind === "actual" ? { pickupDate: targetDate } : { volumeDate: targetDate }), skipped: true, reason: "No mapped outlets; existing data was preserved" });
        continue;
      }
      if (kind === "actual") {
        await syncDataset(targetDate, dataset, volumeCookie, "pickup");
        summaries.push({ date, pickupDate: targetDate, due: dataset.totals.due, outlets: dataset.totals.outlets });
      } else {
        const operating = aggregateOperating(date, rows, appSettings, { rows: rows.length, skippedSheets: parsed.skippedSheets, detailSheets: parsed.detailSheets, targetDate: targetDate || null, missingPickupDates: parsed.missingPickupDates, missingPostcodes: parsed.missingByDate.get(date) || 0, routeMapMatched: parsed.routeMatchedByDate.get(date) || 0, routeMapMissing: parsed.routeMissingByDate.get(date) || 0, routeMapMissingExamples, routeMapDuplicates: 0, routeMapUpdated: map.updatedAt }, importedAt, cbtPickup.get(date), map);
        const aligned = alignOutletDatasets(operating, dataset);
        await store.setJSON(`datasets/${date}.json`, aligned.operating);
        await syncDataset(targetDate, aligned.volume, volumeCookie, "forecast");
        summaries.push({ date, volumeDate: targetDate, operation: aligned.operating.totals.operation, due: aligned.volume.totals.due, outlets: aligned.operating.totals.outletCount });
      }
    }
    return json({ importedAt, kind, detailSheets: parsed.detailSheets, skippedSheets: parsed.skippedSheets, dates: summaries });
  }
  if (path === "/save-dataset" && request.method === "POST") {
    if (!await isOperatingAdmin(request)) return json({ error: "Operating authentication required" }, 401);
    const payload = await request.json();
    payload.volume = expandPackedPostcodes(payload.volume);
    if (!payload?.operating?.date || !payload?.volume?.date || payload.volume.date !== addDays(payload.operating.date, 1)) return json({ error: "Invalid dataset payload" }, 400);
    if (!hasMappableVolume(payload.volume)) return json({ error: "No mapped outlets; existing data was preserved" }, 422);
    const aligned = alignOutletDatasets(payload.operating, payload.volume);
    await syncDataset(aligned.volume.date, aligned.volume, await loginVolume(), "forecast");
    await store.setJSON(`datasets/${aligned.operating.date}.json`, aligned.operating);
    return json({ date: aligned.operating.date, operation: aligned.operating.totals.operation, due: aligned.volume.totals.due, outlets: aligned.operating.totals.outletCount });
  }
  if (path === "/save-pickup" && request.method === "POST") {
    if (!await isOperatingAdmin(request)) return json({ error: "Operating authentication required" }, 401);
    const payload = await request.json();
    payload.pickup = expandPackedPostcodes(payload.pickup);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(payload?.sourceDate || "") || !payload?.pickup?.date || payload.pickup.date !== addDays(payload.sourceDate, 1)) return json({ error: "Invalid pickup dataset payload" }, 400);
    if (!hasMappableVolume(payload.pickup)) return json({ error: "No mapped outlets; existing pickup data was preserved" }, 422);
    await syncDataset(payload.pickup.date, payload.pickup, await loginVolume(), "pickup");
    return json({ sourceDate: payload.sourceDate, pickupDate: payload.pickup.date, due: payload.pickup.totals.due, outlets: payload.pickup.totals.outlets });
  }
  return json({ error: "Not found" }, 404);
}

export default route;
export { parseRouteMapWorkbook, parseOrderRows, parseOrderFile, parseCbtPickup, summaryTargetDate, aggregateOperating, aggregateVolume, alignOutletDatasets, clearSyncedDataset, completePostcodeCoverage, differenceReasonsForDate, expandPackedPostcodes, hasMappableVolume, buildComparison, resolveRoute, sourceType, sourcePlatform, sourcePlatformGroup, pickupOriginCenter, vehicleOverridesForDate, warehouseType, vehicleFor, platformCounts, estimatedParcelsFor, addDays, DEFAULT_SETTINGS, PLATFORMS };
