import * as XLSX from "/Volume/assets/xlsx-Ul8hOgQU.js";
import { parseXlsxOrderDetails } from "/Operating/xlsx-stream.js";

window.downloadOperatingWorkbook = (sheets, filename) => {
  const workbook = XLSX.utils.book_new();
  for (const item of sheets) {
    const worksheet = XLSX.utils.aoa_to_sheet(item.rows);
    const columnCount = Math.max(...item.rows.map((row) => row.length));
    worksheet["!cols"] = Array.from({ length: columnCount }, (_, index) => ({
      wch: Math.min(index === columnCount - 1 ? 42 : 24, Math.max(10, ...item.rows.map((row) => String(row[index] ?? "").length + 2)))
    }));
    XLSX.utils.book_append_sheet(workbook, worksheet, String(item.name || "Data").slice(0, 31));
  }
  XLSX.writeFile(workbook, filename);
};

const ORIGIN_CENTER = "Lutterworth HUB";
const SCHEMA_VERSION = 2;
const DEFAULT_SETTINGS = {
  cubeFactor: 0.00957897245703636,
  vehicleOverrides: {}
};
const CENTERS = [
  { center: "Uxbridge Depot", shortName: "UXB", chinese: "阿克斯布里奇", prefix: "10" },
  { center: "Lutterworth HUB", shortName: "LTW", chinese: "拉特沃思", prefix: "20" },
  { center: "Rochdale Depot", shortName: "ROC", chinese: "罗奇代尔", prefix: "30" },
  { center: "Glasgow Depot", shortName: "GLA", chinese: "格拉斯哥", prefix: "40" }
];
const CODE_TO_OUTLET = {
  UK100001: "LONDON001", UK100002: "C.LONDON002", UK100003: "LONDON003", UK100004: "C.LONDON004", UK100005: "C.LONDON005", UK100006: "C.LONDON006", UK100007: "C.LONDON007", UK100008: "LONDON008", UK100009: "LONDON009", UK100010: "C.LONDON010", UK100011: "C.WATFORD001", UK100012: "C.LONDON012", UK100013: "C.LONDON013", UK100014: "C.LON.SE",
  UK200001: "C.BHX001", UK200002: "C.BHX002", UK200003: "C.LEICESTER001", UK200004: "C.COVENTRY001", UK200005: "C.WALSALL001", UK200006: "C.DUDLEY001", UK200012: "C.NOTTINGHAM001", UK200013: "C.WOLVERHAMPTON001", UK200014: "C.COVENTRY002",
  UK300001: "C.PRESTON001", UK300002: "C.MAN001", UK300003: "C.STOCKPORT001", UK300004: "C.LIV001", UK300005: "C.ROCHDALE001", UK300006: "C.CREWE001",
  UK400001: "C.GLASGOW001", UK400004: "C.EDINBURGH001", UK400005: "C.EDINBURGH002"
};
const OUTLET_TO_CODE = Object.fromEntries(Object.entries(CODE_TO_OUTLET).map(([code, outlet]) => [outlet.toUpperCase(), code]));
const FIELD_ALIASES = {
  pickupDate: ["提货时间", "提货日期", "Pickup Time", "Pickup Date", "pickup date"],
  routeCode: ["二段码", "route code", "Route Code"],
  outlet: ["末端网点", "目的网点", "outlet"],
  depot: ["末端机构", "目的机构", "Depot", "Hub"],
  adjustedPostcode: ["调整后邮编"],
  receiverPostcode: ["收件邮编", "consigneezip", "consignee zip", "consignee zip code"],
  source: ["订单来源", "来源平台", "平台", "order source", "source", "platform"],
  trackingNumber: ["运单号", "TrackingNumber", "tracking number", "物流单号", "waybill number", "订单号", "reference", "主运单号", "Carton No."]
};

const $ = (id) => document.getElementById(id);
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
const normalizeHeader = (value) => clean(value).toLowerCase().replace(/[\s_./\\\-:：()（）]+/g, "");
const normalizedKeyMap = (row) => new Map(Object.keys(row).map((key) => [normalizeHeader(key), key]));
function fieldValue(row, field) {
  const keys = normalizedKeyMap(row);
  for (const alias of FIELD_ALIASES[field]) {
    const key = keys.get(normalizeHeader(alias));
    if (key && clean(row[key])) return row[key];
  }
  return "";
}
function hasField(row, field) {
  const keys = normalizedKeyMap(row);
  return FIELD_ALIASES[field].some((alias) => keys.has(normalizeHeader(alias)));
}
function normalizeRouteCode(value) {
  const text = clean(value).toUpperCase().replace(/^UK/, "");
  const match = text.match(/^(\d{2})\D*(\d{1,4})/);
  return match ? `UK${match[1]}${String(Number(match[2])).padStart(4, "0")}` : "";
}
function codeFrom(row) {
  return normalizeRouteCode(fieldValue(row, "routeCode")) || OUTLET_TO_CODE[clean(fieldValue(row, "outlet")).toUpperCase()] || "";
}
function routeCode(code) {
  return clean(code).replace(/^UK/, "").replace(/^(\d{2})(\d{4})$/, (_, prefix, number) => `${prefix}-${String(Number(number)).padStart(3, "0")}`);
}
function centerFromCode(code) {
  return CENTERS.find((item) => clean(code).startsWith(`UK${item.prefix}`))?.center || "Unmapped";
}
function centerFromDepot(value, code) {
  const text = clean(value).toLowerCase();
  const byName = CENTERS.find((item) => text.includes(item.center.toLowerCase().replace(" depot", "").replace(" hub", "")));
  return byName?.center || centerFromCode(code);
}
function shortNameForCenter(centerName, centerCode = "") {
  const explicit = clean(centerCode).toUpperCase().replace(/\d+$/g, "");
  if (explicit) return explicit;
  const known = CENTERS.find((item) => item.center === centerName)?.shortName;
  if (known) return known;
  return clean(centerName).split(/\s+/).map((part) => part[0] || "").join("").slice(0, 4).toUpperCase() || "CTR";
}
function centerDefinitions(routeMap, rows = []) {
  const definitions = new Map();
  const add = (centerName, centerCode = "") => {
    const center = clean(centerName);
    if (!center || center === "Unmapped") return;
    const known = CENTERS.find((item) => item.center === center);
    const current = definitions.get(center);
    definitions.set(center, {
      center,
      shortName: shortNameForCenter(center, centerCode || current?.shortName),
      chinese: known?.chinese || current?.chinese || "",
    });
  };
  for (const entry of Object.values(routeMap?.entries || {})) if (entry?.enabled) add(entry.centerName, entry.centerCode);
  for (const row of rows) { add(row.destinationCenter, row.destinationCenterCode); add(row.originCenter); }
  add(ORIGIN_CENTER);
  return [...definitions.values()].sort((a, b) => a.shortName.localeCompare(b.shortName));
}
function regionFromCode(code) {
  if (clean(code).startsWith("UK10")) return "South";
  if (clean(code).startsWith("UK20")) return "Midlands";
  if (clean(code).startsWith("UK30")) return "Northwest";
  if (clean(code).startsWith("UK40")) return "North";
  return "Unmapped";
}
function postcodeParts(receiverPostcode, fallbackPostcode = "") {
  const compact = clean(receiverPostcode).toUpperCase().replace(/[^A-Z0-9]/g, "");
  const fallback = clean(fallbackPostcode).toUpperCase().replace(/\s/g, "");
  const full = compact ? compact.length >= 5 ? `${compact.slice(0, -3)} ${compact.slice(-3)}` : compact : "Unknown";
  const district = full === "Unknown" ? fallback || "Unknown" : full.split(" ")[0];
  const area = district.match(/^[A-Z]+/)?.[0] || "Unknown";
  const inward = full === "Unknown" ? "" : full.split(" ")[1] || "";
  return { area, district, sector: inward ? `${district} ${inward[0]}` : district, full };
}
function pickupDate(value) {
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
  if (match) [year, month, day] = [Number(match[1]), Number(match[2]), Number(match[3])];
  if (!month) {
    match = text.match(/(?:20\d{2}\D*)?(\d{1,2})\s*月\s*(\d{1,2})\s*(?:日|号)?/);
    if (match) [month, day] = [Number(match[1]), Number(match[2])];
  }
  if (!month) {
    match = text.match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
    if (match) {
      const first = Number(match[1]), second = Number(match[2]), rawYear = Number(match[3]);
      year = rawYear < 100 ? 2000 + rawYear : rawYear;
      [month, day] = first > 12 ? [second, first] : [first, second];
    }
  }
  if (!month || !day || month > 12 || day > 31) return "";
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}
function addDays(date, days) {
  const [year, month, day] = date.split("-").map(Number);
  const dt = new Date(Date.UTC(year, month - 1, day + days));
  return `${dt.getUTCFullYear()}-${String(dt.getUTCMonth() + 1).padStart(2, "0")}-${String(dt.getUTCDate()).padStart(2, "0")}`;
}
function sourceType(sheet, row) {
  const platformFromText = (value) => {
    const text = clean(value).toUpperCase();
    if (text.includes("TEMU") || text.includes("特木")) return "TEMU";
    if (text.includes("TIKTOK") || text.includes("TIK TOK") || text.includes("抖音") || /(^|[^A-Z])TT([^A-Z]|$)/.test(text)) return "TT";
    if (text.includes("顺丰") || text.includes("SHUNFENG") || /(^|[^A-Z])SF([^A-Z]|$)/.test(text)) return "SF";
    return "OTHER";
  };
  if (clean(fieldValue(row, "source"))) return platformFromText(fieldValue(row, "source"));
  return platformFromText(sheet);
}
function warehouseType(sheet, row) {
  const source = clean(fieldValue(row, "source")) ? fieldValue(row, "source") : sheet;
  return clean(source).toUpperCase().includes("TEMU-CW") ? "CERTIFIED_WAREHOUSE" : "";
}
function pickupOriginCenter(sheet) {
  return clean(sheet).toUpperCase().includes("PIK") ? "Glasgow Depot" : ORIGIN_CENTER;
}
function sourcePlatform(sheet, row) {
  const explicit = clean(fieldValue(row, "source"));
  const value = (explicit || clean(sheet)).toUpperCase();
  if (value.includes("TEMU-CW")) return "TEMU-CW";
  if (value.includes("TEMU-CB")) return "TEMU-CB";
  if (value.includes("TEMU") || value.includes("特木")) return "TEMU";
  if (value.includes("TIKTOK") || value.includes("TIK TOK") || value.includes("抖音") || /(^|[^A-Z])TT([^A-Z]|$)/.test(value)) return explicit ? value : "TT";
  if (value.includes("顺丰") || value.includes("SHUNFENG") || /(^|[^A-Z])SF([^A-Z]|$)/.test(value)) return explicit ? value : "SF";
  return explicit || "OTHER";
}
function routeMapKey(postcodeOutward, route) {
  return `${clean(postcodeOutward).toUpperCase().replace(/\s/g, "")}\u0001${routeCode(normalizeRouteCode(route)) || normalizeRouteCode(route) || clean(route).toUpperCase()}`;
}
function withPostcodeIndex(routeMap) {
  const entries = routeMap?.entries || {};
  const postcodeIndex = {};
  for (const entry of Object.values(entries)) {
    if (!entry?.enabled) continue;
    (postcodeIndex[clean(entry.postcodeOutward).toUpperCase().replace(/\s/g, "")] ||= []).push(entry);
  }
  for (const [postcode, items] of Object.entries(postcodeIndex)) {
    postcodeIndex[postcode] = items.slice().sort((a, b) => clean(b.updatedAt).localeCompare(clean(a.updatedAt)));
  }
  return { entries, postcodeIndex, updatedAt: routeMap?.updatedAt || null };
}
function resolveRoute(row, postal, map, sheet) {
  const normalizedRoute = routeCode(normalizeRouteCode(fieldValue(row, "routeCode"))) || normalizeRouteCode(fieldValue(row, "routeCode"));
  const candidates = (map.postcodeIndex?.[postal.district] || []).filter((entry) => entry?.enabled);
  const mapped = candidates[0] || null;
  const source = { sourceType: sourceType(sheet, row), warehouseType: warehouseType(sheet, row), sourcePlatform: sourcePlatform(sheet, row), originCenter: pickupOriginCenter(sheet) };
  if (mapped?.enabled) return { matched: true, code: mapped.outletCode, routeCode: mapped.routeCode, region: regionFromCode(mapped.outletCode), outlet: mapped.outletName, destinationCenter: mapped.centerName || centerFromCode(mapped.outletCode), destinationCenterCode: mapped.centerCode || "", ...source };
  return { matched: false, code: "", routeCode: normalizedRoute, region: "Unmapped", outlet: "未映射", destinationCenter: "Unmapped", ...source };
}
function platformCounts() {
  return { SF: 0, TT: 0, TEMU: 0, OTHER: 0, CERTIFIED_WAREHOUSE: 0 };
}
function addPlatformCount(target, source, count = 1) {
  target[source in target ? source : "OTHER"] += count;
}
function mergePlatformCounts(...items) {
  const counts = platformCounts();
  for (const item of items) {
    for (const key of Object.keys(counts)) counts[key] += Number(item?.[key] || 0);
  }
  return counts;
}
function addSourcePlatformCount(target, source, count = 1) {
  const key = clean(source).toUpperCase() || "OTHER";
  target[key] = (target[key] || 0) + count;
}
function mergeSourcePlatformCounts(...items) {
  const counts = {};
  for (const item of items) for (const [key, value] of Object.entries(item || {})) counts[key] = (counts[key] || 0) + Number(value || 0);
  return counts;
}
function estimatedParcelsFor(platforms = {}) {
  return Number(platforms.TT || 0) / 50 + Number(platforms.SF || 0) / 50 + Number(platforms.TEMU || 0) / 21;
}
function roundParcels(platforms) {
  return Math.round(estimatedParcelsFor(platforms) * 100) / 100;
}
function capacityFor(region) {
  if (region === "South") return 40;
  if (region === "Midlands" || region === "Northwest") return 70;
  return null;
}
function vehicleFor(center, cubes, settings, date) {
  return settings.vehicleOverrides?.[`${date}:${center}`] || settings.vehicleOverrides?.[center] || "";
}
function buildVolume(date, rows, existing, importedAt, qualityPatch) {
  const volumeRows = rows;
  const outletMap = new Map(), fullPostcodeMap = new Map();
  const addOutlet = (row, metric, count = 1) => {
    if (!count) return;
    const key = [row.region, row.code, row.outlet].join("\u0001");
    const target = outletMap.get(key) || { region: row.region, code: row.code, outlet: row.outlet, due: 0, backlog: 0, total: 0, platforms: platformCounts(), sourcePlatforms: {}, capacity: capacityFor(row.region), drivers: null, backlogShare: 0 };
    target[metric] += count;
    target.total += count;
    if (metric === "due") { addPlatformCount(target.platforms, row.sourceType, count); addPlatformCount(target.platforms, row.warehouseType, count); addSourcePlatformCount(target.sourcePlatforms, row.sourcePlatform || row.sourceType, count); }
    outletMap.set(key, target);
  };
  const addFullPostcode = (row, metric, count = 1) => {
    if (!count) return;
    const key = [row.region, row.code, row.outlet, row.full].join("\u0001");
    const target = fullPostcodeMap.get(key) || { region: row.region, code: row.code, outlet: row.outlet, postal: row.full, due: 0, backlog: 0, total: 0 };
    target[metric] += count;
    target.total += count;
    if (row.detailMissing || row.full === "Postcode detail missing / 邮编明细缺失") target.detailMissing = true;
    fullPostcodeMap.set(key, target);
  };
  for (const row of volumeRows) {
    addOutlet(row, "due");
    addFullPostcode(row, "due");
  }
  for (const row of existing?.outlets || []) addOutlet(row, "backlog", row.backlog || 0);
  for (const row of existing?.views?.["Full Postcode"] || []) addFullPostcode({ ...row, full: row.postal }, "backlog", row.backlog || 0);
  const outlets = [...outletMap.values()].map((row) => ({ ...row, drivers: row.capacity ? Math.ceil(row.total / row.capacity) : null, backlogShare: row.total ? row.backlog / row.total : 0 }));
  const views = { District: [], Area: [], Sector: [], "Full Postcode": completePostcodeCoverage(outlets, [...fullPostcodeMap.values()]) };
  const regions = ["South", "Midlands", "Northwest", "North", "Unmapped"].map((region) => {
    const regionRows = outlets.filter((row) => row.region === region);
    return regionRows.length ? { region, due: regionRows.reduce((sum, row) => sum + row.due, 0), backlog: regionRows.reduce((sum, row) => sum + row.backlog, 0), total: regionRows.reduce((sum, row) => sum + row.total, 0), platforms: mergePlatformCounts(...regionRows.map((row) => row.platforms)), drivers: regionRows.reduce((sum, row) => sum + (row.drivers || 0), 0), outlets: regionRows.length } : null;
  }).filter(Boolean);
  const due = volumeRows.length, backlog = outlets.reduce((sum, row) => sum + row.backlog, 0);
  return { schemaVersion: SCHEMA_VERSION, date, importedAt, backlogEffectiveDate: existing?.backlogEffectiveDate || null, updatedAt: { inbound: importedAt, backlog: existing?.updatedAt?.backlog || existing?.importedAt || null }, sourceCounts: { inbound: due, backlog }, quality: { inboundDuplicates: qualityPatch.duplicates || 0, backlogDuplicates: existing?.quality?.backlogDuplicates || 0, inboundMissingPostcodes: qualityPatch.missingPostcodes || 0, backlogMissingPostcodes: existing?.quality?.backlogMissingPostcodes || 0, missingPostcodes: (qualityPatch.missingPostcodes || 0) + (existing?.quality?.backlogMissingPostcodes || 0), missingPickupDates: qualityPatch.missingPickupDates || 0, routeMapMatched: qualityPatch.routeMapMatched || 0, routeMapMissing: qualityPatch.routeMapMissing || 0, routeMapMissingExamples: qualityPatch.routeMapMissingExamples || [], routeMapDuplicates: 0, routeMapUpdated: qualityPatch.routeMapUpdated || null }, totals: { due, backlog, total: due + backlog, platforms: mergePlatformCounts(...outlets.map((row) => row.platforms)), sourcePlatforms: mergeSourcePlatformCounts(...outlets.map((row) => row.sourcePlatforms)), drivers: outlets.reduce((sum, row) => sum + (row.drivers || 0), 0), outlets: outlets.length }, regions, outlets, views };
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
    rows.push({ region: outlet.region, code: outlet.code, outlet: outlet.outlet,
      postal: "Postcode detail missing / 邮编明细缺失", due: Math.max(0, due), backlog: Math.max(0, backlog),
      total: Math.max(0, due) + Math.max(0, backlog), detailMissing: true });
  }
  return rows.sort((a, b) => b.total - a.total);
}
function packVolumeForUpload(volume) {
  const fullPostcodesPacked = (volume.views?.["Full Postcode"] || []).map((row) => [
    row.region, row.code, row.outlet, row.postal, Number(row.due || 0), Number(row.backlog || 0), row.detailMissing ? 1 : 0,
  ]);
  return { ...volume, views: { ...(volume.views || {}), "Full Postcode": [] }, fullPostcodesPacked };
}
function buildOperating(date, rows, settings, quality, importedAt, routeMap) {
  const centerList = centerDefinitions(routeMap, rows);
  const centerMetrics = new Map(centerList.map((item) => [item.center, { outbound: 0, inbound: 0, sameCenter: 0, cbtPickup: 0 }]));
  const outletMetrics = new Map();
  for (const row of rows) {
    const destination = row.destinationCenter || centerFromCode(row.code);
    if (destination === "Unmapped") continue;
    const origin = row.originCenter || ORIGIN_CENTER;
    const originMetrics = centerMetrics.get(origin) || { outbound: 0, inbound: 0, sameCenter: 0, cbtPickup: 0 };
    const destinationMetrics = centerMetrics.get(destination) || { outbound: 0, inbound: 0, sameCenter: 0, cbtPickup: 0 };
    if (destination === origin) destinationMetrics.sameCenter += 1;
    else {
      originMetrics.outbound += 1;
      destinationMetrics.inbound += 1;
    }
    centerMetrics.set(origin, originMetrics);
    centerMetrics.set(destination, destinationMetrics);
    const outletKey = `${destination}\u0001${row.code}\u0001${row.outlet}`;
    const outlet = outletMetrics.get(outletKey) || { center: destination, outlet: row.outlet, routeCode: row.routeCode, deliveryVolume: 0, types: platformCounts(), sourcePlatforms: {} };
    outlet.deliveryVolume += 1;
    addPlatformCount(outlet.types, row.sourceType);
    addPlatformCount(outlet.types, row.warehouseType);
    addSourcePlatformCount(outlet.sourcePlatforms, row.sourcePlatform || row.sourceType);
    outletMetrics.set(outletKey, outlet);
  }
  const outlets = [...outletMetrics.values()].filter((row) => row.center !== "Unmapped").sort((a, b) => a.routeCode.localeCompare(b.routeCode)).map((row) => ({ center: row.center, outlet: row.outlet, routeCode: row.routeCode, deliveryVolume: row.deliveryVolume, estimatedParcels: roundParcels(row.types), platforms: row.types, sourcePlatforms: row.sourcePlatforms }));
  const centers = centerDefinitions(routeMap, rows).map((item) => {
    const metrics = centerMetrics.get(item.center) || { outbound: 0, inbound: 0, sameCenter: 0, cbtPickup: 0 };
    const centerOutlets = outlets.filter((row) => row.center === item.center);
    const platforms = mergePlatformCounts(...centerOutlets.map((outlet) => outlet.platforms));
    const estimatedParcels = Math.round(estimatedParcelsFor(platforms));
    const operation = metrics.outbound + metrics.inbound + metrics.sameCenter + metrics.cbtPickup;
    const cubeVolume = metrics.inbound + metrics.sameCenter;
    const estimatedCubes = Math.round(cubeVolume * settings.cubeFactor * 100) / 100;
    return { center: item.center, shortName: item.shortName, chinese: item.chinese, operation, outbound: metrics.outbound, inbound: metrics.inbound, sameCenter: metrics.sameCenter, cbtPickup: metrics.cbtPickup, platforms, estimatedParcels, estimatedCubes, vehicle: vehicleFor(item.center, estimatedCubes, settings, date), outletCount: centerOutlets.length };
  });
  const totalPlatforms = mergePlatformCounts(...outlets.map((row) => row.platforms));
  const totals = { operation: centers.reduce((sum, row) => sum + row.operation, 0), outbound: centers.reduce((sum, row) => sum + row.outbound, 0), inbound: centers.reduce((sum, row) => sum + row.inbound, 0), sameCenter: centers.reduce((sum, row) => sum + row.sameCenter, 0), cbtPickup: 0, estimatedParcels: Math.round(estimatedParcelsFor(totalPlatforms)), estimatedCubes: Math.round(centers.reduce((sum, row) => sum + row.estimatedCubes, 0) * 100) / 100, outletCount: outlets.length, centers: centers.length, platforms: totalPlatforms, sourcePlatforms: mergeSourcePlatformCounts(...outlets.map((row) => row.sourcePlatforms)) };
  return { schemaVersion: SCHEMA_VERSION, date, importedAt, title: "提货单量预测", settings, totals, centers, outlets, quality, forecastPlatformsAvailable: true };
}
function parseWorkbook(arrayBuffer, routeMap, selectedDate = "") {
  const book = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
  const map = withPostcodeIndex(routeMap);
  const byDate = new Map(), skippedSheets = [], detailSheets = [], routeMapMissingKeys = new Map();
  const seenOrders = new Set();
  let rows = 0, duplicates = 0, missingPickupDates = 0, missingPostcodes = 0, routeMapMissing = 0, routeMapMatched = 0;
  for (const sheet of book.SheetNames) {
    const records = XLSX.utils.sheet_to_json(book.Sheets[sheet], { defval: "", raw: false });
    const first = records[0] || {};
    if (!records.length || !hasField(first, "pickupDate") || (!hasField(first, "receiverPostcode") && !hasField(first, "adjustedPostcode"))) {
      skippedSheets.push(sheet);
      continue;
    }
    detailSheets.push(sheet);
    for (const record of records) {
      // Template tabs can contain a date-only placeholder row. It is not an order.
      if (!fieldValue(record, "receiverPostcode") && !fieldValue(record, "adjustedPostcode")) continue;
      const detectedDate = pickupDate(fieldValue(record, "pickupDate"));
      if (selectedDate && detectedDate && detectedDate !== selectedDate) continue;
      const date = selectedDate || detectedDate;
      if (!date) {
        missingPickupDates += 1;
        continue;
      }
      const trackingNumber = clean(fieldValue(record, "trackingNumber")).toUpperCase();
      if (trackingNumber && seenOrders.has(trackingNumber)) {
        duplicates += 1;
        continue;
      }
      if (trackingNumber) seenOrders.add(trackingNumber);
      const postcodes = postcodeParts(fieldValue(record, "receiverPostcode"), fieldValue(record, "adjustedPostcode"));
      if (postcodes.full === "Unknown") missingPostcodes += 1;
      const resolved = resolveRoute(record, postcodes, map, sheet);
      if (resolved.matched) routeMapMatched += 1;
      else {
        routeMapMissing += 1;
        const missingKey = `${postcodes.district} + ${resolved.routeCode || "无二段码"}`;
        routeMapMissingKeys.set(missingKey, (routeMapMissingKeys.get(missingKey) || 0) + 1);
      }
      const parsed = { ...resolved, ...postcodes };
      const dateRows = byDate.get(date) || [];
      dateRows.push(parsed);
      byDate.set(date, dateRows);
      rows += 1;
    }
  }
  const routeMapMissingExamples = [...routeMapMissingKeys.entries()].sort((a, b) => b[1] - a[1]).slice(0, 12).map(([key, count]) => ({ key, count }));
  return { byDate, quality: { rows, duplicates, skippedSheets, detailSheets, missingPickupDates, missingPostcodes, routeMapMissing, routeMapDuplicates: 0, routeMapMatched, routeMapMissingExamples, routeMapUpdated: map.updatedAt } };
}

function parsedStreamedWorkbook(streamed, routeMap, selectedDate) {
  const map = withPostcodeIndex(routeMap);
  const routeMapMissingKeys = new Map();
  const parsedRows = [];
  let routeMapMatched = 0, routeMapMissing = 0;
  for (const item of streamed.records) {
    const record = {
      "提货日期": item.pickupDate,
      "收件邮编": item.receiverPostcode,
      "调整后邮编": item.adjustedPostcode,
      "二段码": item.routeCode,
      "末端网点": item.outlet,
      "末端机构": item.depot,
      "订单来源": item.source,
      "运单号": item.trackingNumber,
    };
    const postcodes = postcodeParts(item.receiverPostcode, item.adjustedPostcode);
    const resolved = resolveRoute(record, postcodes, map, item.sheet);
    if (resolved.matched) routeMapMatched += 1;
    else {
      routeMapMissing += 1;
      const missingKey = `${postcodes.district} + ${resolved.routeCode || "无二段码"}`;
      routeMapMissingKeys.set(missingKey, (routeMapMissingKeys.get(missingKey) || 0) + 1);
    }
    parsedRows.push({ ...resolved, ...postcodes });
  }
  const routeMapMissingExamples = [...routeMapMissingKeys.entries()].sort((a, b) => b[1] - a[1]).slice(0, 12).map(([key, count]) => ({ key, count }));
  return {
    byDate: parsedRows.length ? new Map([[selectedDate, parsedRows]]) : new Map(),
    platformsAvailable: true,
    quality: { ...streamed.quality, routeMapMatched, routeMapMissing, routeMapMissingExamples, routeMapDuplicates: 0, routeMapUpdated: map.updatedAt },
  };
}

function positiveWholeNumber(value) {
  const number = Number(clean(value).replace(/,/g, ""));
  return Number.isFinite(number) && number > 0 ? Math.round(number) : 0;
}

function actualSummaryRouteIndex(routeMap) {
  const result = new Map();
  for (const entry of Object.values(routeMap?.entries || {})) {
    if (!entry?.enabled) continue;
    const code = normalizeRouteCode(entry.outletCode || entry.routeCode);
    const key = routeCode(code);
    if (!key) continue;
    const current = result.get(key);
    if (!current || clean(entry.updatedAt).localeCompare(clean(current.updatedAt)) >= 0) result.set(key, entry);
  }
  return result;
}

function parseActualSummary(arrayBuffer, routeMap, selectedDate = "") {
  const directory = XLSX.read(arrayBuffer, { type: "array", bookSheets: true, bookProps: true });
  const candidates = directory.SheetNames.filter((name) => /提货订单汇总|实际操作/i.test(clean(name)));
  const byDate = new Map();
  if (!candidates.length) return { byDate, quality: {}, platformsAvailable: true };

  const book = XLSX.read(arrayBuffer, { type: "array", cellDates: false, sheets: candidates });
  const mappedByRoute = actualSummaryRouteIndex(routeMap);
  let best = null;
  for (const sheet of candidates) {
    if (!book.Sheets[sheet]) continue;
    const matrix = XLSX.utils.sheet_to_json(book.Sheets[sheet], { header: 1, defval: "", raw: false, blankrows: false });
    const headerRow = matrix.findIndex((row) => {
      const headers = row.map(normalizeHeader);
      return headers.includes(normalizeHeader("派件网点")) && headers.includes(normalizeHeader("二段码")) && headers.includes(normalizeHeader("派件量"));
    });
    if (headerRow < 0) continue;
    const heading = matrix.slice(0, headerRow).flat().map(clean).filter(Boolean).join(" ");
    if (!/实际操作/i.test(heading)) continue;

    const headers = matrix[headerRow].map(normalizeHeader);
    const outletColumn = headers.indexOf(normalizeHeader("派件网点"));
    const routeColumn = headers.indexOf(normalizeHeader("二段码"));
    const volumeColumn = headers.indexOf(normalizeHeader("派件量"));
    const operationColumn = headers.indexOf(normalizeHeader("操作量"));
    const sourceDate = selectedDate || pickupDate(heading);
    if (!sourceDate) continue;

    const rows = [];
    let operationTotal = 0;
    let outletRows = 0;
    for (const row of matrix.slice(headerRow + 1)) {
      if (row.some((value) => clean(value) === "合计")) operationTotal = Math.max(operationTotal, positiveWholeNumber(row[operationColumn]));
      const displayRoute = routeCode(normalizeRouteCode(row[routeColumn]));
      const deliveryVolume = positiveWholeNumber(row[volumeColumn]);
      if (!displayRoute || !deliveryVolume) continue;
      const mapped = mappedByRoute.get(displayRoute);
      const code = normalizeRouteCode(mapped?.outletCode || displayRoute);
      const outlet = clean(mapped?.outletName) || CODE_TO_OUTLET[code] || clean(row[outletColumn]);
      if (!code || !outlet || outlet === "合计") continue;
      const destinationCenter = clean(mapped?.centerName) || centerFromCode(code);
      const template = {
        code,
        routeCode: displayRoute,
        region: regionFromCode(code),
        outlet,
        destinationCenter,
        sourceType: "OTHER",
        warehouseType: "",
        area: "Unknown",
        district: "Unknown",
        sector: "Unknown",
        full: "Postcode detail missing / 邮编明细缺失",
        detailMissing: true,
      };
      for (let index = 0; index < deliveryVolume; index += 1) rows.push(template);
      outletRows += 1;
    }
    if (!rows.length || (best && best.rows.length >= rows.length)) continue;
    best = { sourceDate, rows, sheet, operationTotal, outletRows };
  }

  if (!best) return { byDate, quality: {}, platformsAvailable: true };
  byDate.set(best.sourceDate, best.rows);
  return {
    byDate,
    platformsAvailable: false,
    quality: {
      rows: best.rows.length,
      skippedSheets: directory.SheetNames.filter((name) => name !== best.sheet),
      detailSheets: [best.sheet],
      summarySheet: best.sheet,
      summaryOperationTotal: best.operationTotal,
      summaryOutletRows: best.outletRows,
      missingPickupDates: 0,
      missingPostcodes: best.rows.length,
      routeMapMissing: 0,
      routeMapDuplicates: 0,
      routeMapMatched: best.rows.length,
      routeMapMissingExamples: [],
      routeMapUpdated: routeMap?.updatedAt || null,
    },
  };
}
async function fetchJson(path, fallback) {
  try {
    const response = await fetch(path);
    if (!response.ok) return fallback;
    return response.json();
  } catch {
    return fallback;
  }
}
async function latestExistingVolume(volumeDate, kind = "forecast") {
  const base = kind === "actual" ? "/Pickup/api" : "/Volume/api";
  return fetchJson(`${base}/datasets/${volumeDate}?full=1`, null);
}
function translate(zh, en) {
  return document.documentElement.lang === "en" ? en : zh;
}
function setNotice(text) {
  const notice = $("notice");
  notice.textContent = text;
  notice.classList.toggle("open", Boolean(text));
}
async function localForecastImport() {
  const file = $("forecastInput")?.files?.[0];
  if (!file) {
    setNotice(translate("请先选择文件。", "Choose a file first."));
    return null;
  }
  const selectedDate = clean($("forecastDate")?.value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(selectedDate)) {
    setNotice(translate("请先选择预测数据的操作日期。", "Choose the forecast operation date first."));
    return null;
  }
  const started = Date.now();
  const progress = window.setInterval(() => {
    setNotice(translate(`正在本地解析 ${file.name} · ${Math.floor((Date.now() - started) / 1000)}s`, `Parsing ${file.name} locally · ${Math.floor((Date.now() - started) / 1000)}s`));
  }, 1000);
  try {
    setNotice(translate("正在本地解析 Excel，请不要关闭页面。", "Parsing Excel locally. Please keep this page open."));
    const routeMapResult = await fetchJson("/Operating/api/route-map/export", { routeMap: null });
    const arrayBuffer = await file.arrayBuffer();
    const parsed = file.size >= 20 * 1024 * 1024
      ? parsedStreamedWorkbook(await parseXlsxOrderDetails(arrayBuffer, selectedDate), routeMapResult.routeMap, selectedDate)
      : parseWorkbook(arrayBuffer, routeMapResult.routeMap, selectedDate);
    if (!parsed.byDate.size) throw new Error(translate("没有识别到可导入的订单明细表。", "No importable order detail sheets were detected."));
    const settingsResult = await fetchJson("/Operating/api/settings", { settings: DEFAULT_SETTINGS });
    const settings = { ...DEFAULT_SETTINGS, ...(settingsResult.settings || {}) };
    const importedAt = new Date().toISOString();
    const done = [];
    for (const date of [...parsed.byDate.keys()].sort()) {
      const volumeDate = addDays(date, 1);
      setNotice(translate(`正在同步 ${date}，并更新 Volume ${volumeDate}。`, `Syncing ${date} and updating Volume ${volumeDate}.`));
      const rows = parsed.byDate.get(date);
      const quality = { ...parsed.quality, rows: rows.length };
      const existing = await latestExistingVolume(volumeDate);
      const operating = buildOperating(date, rows, settings, quality, importedAt, routeMapResult.routeMap);
      const volume = buildVolume(volumeDate, rows, existing, importedAt, quality);
      if (!Number(volume.totals?.due || 0)) {
        done.push(translate(`${date} 已跳过（没有匹配到有效网点）`, `${date} skipped (no mapped outlets)`));
        continue;
      }
      const response = await fetch("/Operating/api/save-dataset", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ operating, volume: packVolumeForUpload(volume) })
      });
      const result = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(result.error || translate("同步失败，请稍后重试。", "Sync failed. Please try again later."));
      done.push(`${date} -> ${volumeDate}`);
    }
    setNotice(translate(`导入成功：${done.join("，")}。`, `Import complete: ${done.join(", ")}.`));
    window.setTimeout(() => window.location.reload(), 1200);
    return { done };
  } finally {
    window.clearInterval(progress);
  }
}

async function localActualImport() {
  const file = $("actualInput")?.files?.[0];
  if (!file) {
    setNotice(translate("请先选择文件。", "Choose a file first."));
    return null;
  }
  const selectedDate = clean($("actualDate")?.value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(selectedDate)) {
    setNotice(translate("请先选择实际数据的操作日期。", "Choose the actual operation date first."));
    return null;
  }
  const started = Date.now();
  const progress = window.setInterval(() => {
    setNotice(translate(`正在本地解析实际数据 ${file.name} · ${Math.floor((Date.now() - started) / 1000)}s`, `Parsing actual data ${file.name} locally · ${Math.floor((Date.now() - started) / 1000)}s`));
  }, 1000);
  try {
    setNotice(translate("正在本地解析实际操作 Excel，请不要关闭页面。", "Parsing the actual operations Excel locally. Please keep this page open."));
    const routeMapResult = await fetchJson("/Operating/api/route-map/export", { routeMap: null });
    const arrayBuffer = await file.arrayBuffer();
    let parsed;
    try {
      parsed = parsedStreamedWorkbook(await parseXlsxOrderDetails(arrayBuffer, selectedDate), routeMapResult.routeMap, selectedDate);
    } catch (error) {
      if (/\.xlsx$/i.test(file.name)) throw error;
      parsed = parseWorkbook(arrayBuffer, routeMapResult.routeMap, selectedDate);
    }
    if (!parsed.byDate.size && parsed.quality?.detailSheets?.length) {
      throw new Error(translate(`明细表中没有 ${selectedDate} 的有效订单，请检查所选操作日期。`, `No valid orders dated ${selectedDate} were found in the detail sheets. Check the selected operation date.`));
    }
    if (!parsed.byDate.size) parsed = parseActualSummary(arrayBuffer, routeMapResult.routeMap, selectedDate);
    if (!parsed.byDate.size) throw new Error(translate("没有识别到可导入的实际操作明细表。", "No importable actual operation detail sheets were detected."));
    const importedAt = new Date().toISOString();
    const done = [];
    for (const date of [...parsed.byDate.keys()].sort()) {
      const pickupDate = addDays(date, 1);
      setNotice(translate(`正在同步 ${date} 的实际操作数据，并更新 Pickup ${pickupDate}。`, `Syncing actual operations for ${date} and updating Pickup ${pickupDate}.`));
      const rows = parsed.byDate.get(date);
      const quality = { ...parsed.quality, rows: rows.length };
      const existing = await latestExistingVolume(pickupDate, "actual");
      const pickup = buildVolume(pickupDate, rows, existing, importedAt, quality);
      pickup.platformsAvailable = parsed.platformsAvailable !== false;
      if (!Number(pickup.totals?.due || 0)) {
        done.push(translate(`${date} 已跳过（没有匹配到有效网点）`, `${date} skipped (no mapped outlets)`));
        continue;
      }
      const response = await fetch("/Operating/api/save-pickup", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ sourceDate: date, pickup: packVolumeForUpload(pickup) })
      });
      const result = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(result.error || translate("实际数据同步失败，请稍后重试。", "Actual data sync failed. Please try again later."));
      if (pickup.platformsAvailable === false) done.push(`${date} (${pickup.totals.due} ${translate("票，汇总表无平台明细", "shipments, no platform detail in summary")})`);
      else {
        const platforms = pickup.totals?.platforms || platformCounts();
        done.push(`${date} (SF ${platforms.SF || 0}, TT ${platforms.TT || 0}, TEMU ${platforms.TEMU || 0})`);
      }
    }
    setNotice(translate(`实际数据导入成功：${done.join("，")}。`, `Actual import complete: ${done.join(", ")}.`));
    window.setTimeout(() => window.location.reload(), 1200);
    return { done };
  } finally {
    window.clearInterval(progress);
  }
}

const forecastButton = $("forecastBtn");
if (forecastButton) {
  forecastButton.onclick = () => localForecastImport().catch((error) => setNotice(error.message || translate("导入失败。", "Import failed.")));
}
const actualButton = $("actualBtn");
if (actualButton) {
  actualButton.onclick = () => localActualImport().catch((error) => setNotice(error.message || translate("实际数据导入失败。", "Actual import failed.")));
}
