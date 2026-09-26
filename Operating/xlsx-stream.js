import JSZip from "../Volume/assets/jszip.min-UlDkZ5Fx.js";

const FIELD_ALIASES = {
  pickupDate: ["提货时间", "提货日期", "Pickup Time", "Pickup Date", "pickup date"],
  routeCode: ["二段码", "route code", "Route Code", "线路编码", "路由编码"],
  outlet: ["末端网点", "目的网点", "网点名称", "outlet", "destination outlet", "末端网点编码"],
  depot: ["末端机构", "目的机构", "Depot", "Hub", "末端中心/集散"],
  adjustedPostcode: ["调整后邮编", "outward postcode", "postcode prefix"],
  receiverPostcode: ["consigneezip", "consignee zip", "收件邮编", "目的邮编", "destination postcode", "receiver postcode", "邮编"],
  source: ["订单来源", "来源平台", "平台", "order source", "source", "platform"],
  trackingNumber: ["运单号", "TrackingNumber", "tracking number", "物流单号", "waybill number", "订单号", "reference", "主运单号", "Carton No."],
};

const clean = (value) => {
  const text = value == null ? "" : String(value).trim();
  return /^#(?:REF|VALUE|N\/A|DIV\/0|NAME|NULL|NUM)!?$/i.test(text) ? "" : text;
};

const normalizeHeader = (value) => clean(value).toLowerCase().replace(/[\s_./\\\-:：()（）]+/g, "");

function decodeXml(value = "") {
  return value.replace(/&(?:#x[\da-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (entity) => {
    if (entity === "&amp;") return "&";
    if (entity === "&lt;") return "<";
    if (entity === "&gt;") return ">";
    if (entity === "&quot;") return '"';
    if (entity === "&apos;") return "'";
    const hex = entity.match(/^&#x([\da-f]+);$/i);
    const decimal = entity.match(/^&#(\d+);$/);
    const codePoint = hex ? Number.parseInt(hex[1], 16) : decimal ? Number(decimal[1]) : 0;
    return codePoint ? String.fromCodePoint(codePoint) : entity;
  });
}

function xmlText(xml = "") {
  return [...xml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((match) => decodeXml(match[1])).join("");
}

function streamEntry(entry, onText) {
  if (!entry) return Promise.reject(new Error("Required XLSX entry is missing."));
  return new Promise((resolve, reject) => {
    const decoder = new TextDecoder();
    const stream = entry.internalStream("uint8array");
    let settled = false;
    const stop = (result) => {
      if (settled) return;
      settled = true;
      stream.pause();
      resolve(result);
    };
    stream.on("data", (chunk) => {
      if (settled) return;
      try {
        if (onText(decoder.decode(chunk, { stream: true })) === false) stop({ stopped: true });
      } catch (error) {
        settled = true;
        stream.pause();
        reject(error);
      }
    });
    stream.on("error", (error) => {
      if (settled) return;
      settled = true;
      reject(error);
    });
    stream.on("end", () => {
      if (settled) return;
      try {
        onText(decoder.decode());
        settled = true;
        resolve({ stopped: false });
      } catch (error) {
        settled = true;
        reject(error);
      }
    });
    stream.resume();
  });
}

async function readSharedStrings(entry, onProgress) {
  if (!entry) return [];
  const values = [];
  let buffer = "";
  let parsed = 0;
  await streamEntry(entry, (chunk) => {
    buffer += chunk;
    for (;;) {
      const start = buffer.indexOf("<si");
      if (start < 0) {
        if (buffer.length > 256) buffer = buffer.slice(-256);
        break;
      }
      const end = buffer.indexOf("</si>", start);
      if (end < 0) {
        if (start > 0) buffer = buffer.slice(start);
        break;
      }
      values.push(xmlText(buffer.slice(start, end + 5)));
      buffer = buffer.slice(end + 5);
      parsed += 1;
      if (parsed % 100000 === 0) onProgress?.({ phase: "sharedStrings", rows: parsed });
    }
  });
  return values;
}

function worksheetDirectory(workbookXml, relationshipsXml) {
  const relationships = new Map();
  for (const match of relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)) {
    const attrs = match[1];
    const id = attrs.match(/\bId="([^"]+)"/)?.[1];
    const target = attrs.match(/\bTarget="([^"]+)"/)?.[1];
    if (id && target) relationships.set(decodeXml(id), decodeXml(target));
  }
  const sheets = [];
  for (const match of workbookXml.matchAll(/<sheet\b([^>]*)\/?\s*>/g)) {
    const attrs = match[1];
    const name = attrs.match(/\bname="([^"]+)"/)?.[1];
    const relationshipId = attrs.match(/\br:id="([^"]+)"/)?.[1];
    const target = relationships.get(decodeXml(relationshipId || ""));
    if (!name || !target) continue;
    const normalizedTarget = target.replace(/^\/?xl\//, "").replace(/^\//, "");
    sheets.push({ name: decodeXml(name), path: `xl/${normalizedTarget}` });
  }
  return sheets;
}

function parseCells(rowXml, sharedStrings) {
  const cells = new Map();
  const populatedCells = rowXml.replace(/<c\b[^>]*\/>/g, "");
  const cellPattern = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
  for (const match of populatedCells.matchAll(cellPattern)) {
    const attributes = match[1];
    const column = attributes.match(/(?:^|\s)r="([A-Z]+)\d+"/)?.[1];
    if (!column) continue;
    const body = match[2] || "";
    const type = attributes.match(/\bt="([^"]+)"/)?.[1] || "n";
    const raw = body.match(/<v>([\s\S]*?)<\/v>/)?.[1] || "";
    let value = decodeXml(raw);
    if (type === "s" && /^\d+$/.test(value)) value = sharedStrings[Number(value)] || "";
    else if (type === "inlineStr") value = xmlText(body);
    cells.set(column, clean(value));
  }
  return cells;
}

function fieldColumns(headers, field) {
  const columns = [];
  for (const alias of FIELD_ALIASES[field].map(normalizeHeader)) {
    for (const [column, header] of headers) {
      if (normalizeHeader(header) === alias && !columns.includes(column)) columns.push(column);
    }
  }
  return columns;
}

function firstCellValue(cells, columns) {
  for (const column of columns) {
    const value = clean(cells.get(column));
    if (value) return value;
  }
  return "";
}

function excelDate(value) {
  const text = clean(value);
  if (/^\d{5}(?:\.\d+)?$/.test(text)) {
    const serial = Number(text);
    if (serial > 20000 && serial < 80000) return new Date(Date.UTC(1899, 11, 30 + Math.floor(serial))).toISOString().slice(0, 10);
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

async function forEachWorksheetRow(entry, sharedStrings, callback) {
  let buffer = "";
  return streamEntry(entry, (chunk) => {
    buffer += chunk;
    for (;;) {
      const start = buffer.indexOf("<row");
      if (start < 0) {
        if (buffer.length > 256) buffer = buffer.slice(-256);
        break;
      }
      const end = buffer.indexOf("</row>", start);
      if (end < 0) {
        if (start > 0) buffer = buffer.slice(start);
        break;
      }
      const rowXml = buffer.slice(start, end + 6);
      buffer = buffer.slice(end + 6);
      const rowNumber = Number(rowXml.match(/^<row\b[^>]*\br="(\d+)"/)?.[1] || 0);
      if (callback(parseCells(rowXml, sharedStrings), rowNumber) === false) return false;
    }
    return true;
  });
}

function sheetCanContainOrders(name) {
  return !/(?:二段码|汇总|summary|无段码|异常|航班|\bCBT\b)/i.test(clean(name));
}

export async function parseXlsxOrderDetails(arrayBuffer, selectedDate, { onProgress } = {}) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const workbookEntry = zip.file("xl/workbook.xml");
  const relationshipsEntry = zip.file("xl/_rels/workbook.xml.rels");
  if (!workbookEntry || !relationshipsEntry) throw new Error("Unsupported XLSX workbook structure.");
  const [workbookXml, relationshipsXml] = await Promise.all([workbookEntry.async("string"), relationshipsEntry.async("string")]);
  const directory = worksheetDirectory(workbookXml, relationshipsXml);
  onProgress?.({ phase: "sharedStrings", rows: 0 });
  const sharedStrings = await readSharedStrings(zip.file("xl/sharedStrings.xml"), onProgress);
  const records = [];
  const detailSheets = [];
  const skippedSheets = [];
  const seenOrders = new Set();
  let duplicates = 0;
  let missingPickupDates = 0;
  let missingPostcodes = 0;

  for (const sheet of directory) {
    if (!sheetCanContainOrders(sheet.name) || !zip.file(sheet.path)) {
      skippedSheets.push(sheet.name);
      continue;
    }
    onProgress?.({ phase: "sheet", sheet: sheet.name, rows: records.length });
    let headers = null;
    let columns = null;
    let headerCandidates = 0;
    let orderRowsSeen = 0;
    let emptyTail = 0;
    await forEachWorksheetRow(zip.file(sheet.path), sharedStrings, (cells, rowNumber) => {
      if (!headers) {
        headerCandidates += 1;
        const candidate = new Map([...cells].map(([column, value]) => {
          const resolved = /^\d+$/.test(value) ? sharedStrings[Number(value)] || value : value;
          return [column, resolved];
        }).filter(([, value]) => clean(value)));
        const hasDate = fieldColumns(candidate, "pickupDate").length > 0;
        const hasPostcode = fieldColumns(candidate, "receiverPostcode").length > 0 || fieldColumns(candidate, "adjustedPostcode").length > 0;
        if (hasDate && hasPostcode) {
          headers = candidate;
          columns = Object.fromEntries(Object.keys(FIELD_ALIASES).map((field) => [field, fieldColumns(headers, field)]));
          detailSheets.push(sheet.name);
          onProgress?.({ phase: "headers", sheet: sheet.name, headers: Object.fromEntries(candidate), columns });
        }
        return headerCandidates < 50;
      }

      const rowDate = excelDate(firstCellValue(cells, columns.pickupDate));
      const trackingNumber = firstCellValue(cells, columns.trackingNumber);
      const receiverPostcode = firstCellValue(cells, columns.receiverPostcode);
      const adjustedPostcode = firstCellValue(cells, columns.adjustedPostcode);
      const hasIdentityColumns = columns.trackingNumber.length > 0;
      if (trackingNumber || receiverPostcode || adjustedPostcode) orderRowsSeen += 1;
      if (orderRowsSeen && !trackingNumber && !receiverPostcode && !adjustedPostcode) {
        emptyTail += 1;
        if (emptyTail >= 1000) return false;
      } else if (trackingNumber || receiverPostcode || adjustedPostcode) {
        emptyTail = 0;
      }

      if (rowDate && rowDate !== selectedDate) return true;
      if (hasIdentityColumns && !trackingNumber) return true;
      if (!receiverPostcode && !adjustedPostcode) {
        if (trackingNumber) missingPostcodes += 1;
        return true;
      }
      if (!rowDate) missingPickupDates += 1;

      const identity = trackingNumber ? clean(trackingNumber).toUpperCase() : `${sheet.name}\u0001${rowNumber}`;
      if (seenOrders.has(identity)) {
        duplicates += 1;
        return true;
      }
      seenOrders.add(identity);
      records.push({
        sheet: sheet.name,
        pickupDate: selectedDate,
        trackingNumber,
        receiverPostcode,
        adjustedPostcode,
        routeCode: firstCellValue(cells, columns.routeCode),
        outlet: firstCellValue(cells, columns.outlet),
        depot: firstCellValue(cells, columns.depot),
        source: firstCellValue(cells, columns.source),
      });
      if (records.length % 5000 === 0) onProgress?.({ phase: "orders", sheet: sheet.name, rows: records.length });
      return true;
    });
    if (!headers) skippedSheets.push(sheet.name);
  }

  return {
    records,
    quality: { rows: records.length, duplicates, missingPickupDates, missingPostcodes, detailSheets, skippedSheets },
  };
}

export { excelDate, normalizeHeader };
