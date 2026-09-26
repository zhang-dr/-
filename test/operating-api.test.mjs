import test from "node:test";
import assert from "node:assert/strict";
import * as XLSX from "xlsx";
import {
  DEFAULT_SETTINGS,
  aggregateOperating,
  aggregateVolume,
  alignOutletDatasets,
  completePostcodeCoverage,
  differenceReasonsForDate,
  expandPackedPostcodes,
  hasMappableVolume,
  buildComparison,
  estimatedParcelsFor,
  resolveRoute,
  sourcePlatform,
  sourcePlatformGroup,
  sourceType,
  pickupOriginCenter,
  warehouseType,
  vehicleFor,
  vehicleOverridesForDate,
  parseRouteMapWorkbook,
} from "../netlify/functions/operating-api.mjs";

const mappedRow = (overrides = {}) => ({
  code: "UK200002",
  routeCode: "20-002",
  region: "Midlands",
  outlet: "C.BHX002",
  destinationCenter: "Lutterworth HUB",
  sourceType: "SF",
  area: "B",
  district: "B13",
  sector: "B13 8",
  full: "B13 8XX",
  ...overrides,
});

test("platform comes from Order Source before sheet name", () => {
  assert.equal(sourceType("SF提货订单明细", { 订单来源: "TEMU" }), "TEMU");
  assert.equal(sourceType("TT正式分单", { 订单来源: "" }), "OTHER");
  assert.equal(sourceType("TT正式分单", { 其他字段: "" }), "TT");
  assert.equal(sourceType("TEMU正式分单", { 订单来源: "抖音" }), "TT");
  assert.equal(sourceType("TEMU正式分单", { 订单来源: "顺丰" }), "SF");
  assert.equal(sourceType("TEMU正式分单", { 订单来源: "SF-CB" }), "SF");
  assert.equal(sourceType("TEMU正式分单", { 订单来源: "TT-CB" }), "TT");
  assert.equal(sourceType("TT正式分单", { 订单来源: "TEMU-CB" }), "TEMU");
});

test("platform comparison keeps exact Order Source values", () => {
  assert.equal(sourcePlatform("TEMU正式分单", { 订单来源: "TEMU-CB" }), "TEMU-CB");
  assert.equal(sourcePlatform("TEMU正式分单", { 订单来源: "TEMU-CW" }), "TEMU-CW");
  assert.equal(sourcePlatform("混合数据", { 订单来源: "TT-CB" }), "TT-CB");
});

test("platform comparison groups equivalent cross-border source labels", () => {
  assert.equal(sourcePlatformGroup("SF"), "SF");
  assert.equal(sourcePlatformGroup("SF-CB"), "SF");
  assert.equal(sourcePlatformGroup("TT-CB"), "TT");
  assert.equal(sourcePlatformGroup("TEMU-CB"), "TEMU-CB");
  assert.equal(sourcePlatformGroup("TEMU-CW"), "TEMU-CW");
  const result = buildComparison({
    schemaVersion: 2,
    date: "2026-09-10",
    quality: {},
    centers: [],
    outlets: [],
    totals: {
      platforms: { SF: 5741, TT: 0, TEMU: 0 },
      sourcePlatforms: { SF: 4891, "SF-CB": 850 },
    },
  }, null);
  assert.deepEqual(Object.keys(result.platformComparison), ["SF"]);
  assert.equal(result.platformComparison.SF.forecast, 5741);
});

test("PIK pickup belongs to Glasgow and does not add Lutterworth operation", () => {
  assert.equal(pickupOriginCenter("PIK提单"), "Glasgow Depot");
  const result = aggregateOperating("2026-09-09", [mappedRow({
    code: "UK400001",
    routeCode: "40-001",
    region: "North",
    outlet: "C.GLASGOW001",
    destinationCenter: "Glasgow Depot",
    originCenter: "Glasgow Depot",
    sourceType: "TEMU",
    sourcePlatform: "TEMU-CB",
  })], DEFAULT_SETTINGS, { rows: 1 }, "2026-09-09T00:00:00Z");
  const gla = result.centers.find((center) => center.shortName === "GLA");
  const ltw = result.centers.find((center) => center.shortName === "LTW");
  assert.equal(gla.operation, 1);
  assert.equal(gla.sameCenter, 1);
  assert.equal(gla.platforms.TEMU, 1);
  assert.equal(result.outlets.find((outlet) => outlet.outlet === "C.GLASGOW001").deliveryVolume, 1);
  assert.equal(ltw.operation, 0);
  assert.equal(ltw.outbound, 0);
});

test("parcel estimate uses TT and SF divided by 50 and TEMU divided by 21", () => {
  assert.equal(estimatedParcelsFor({ TT: 50, SF: 100, TEMU: 42 }), 5);
});

test("TEMU-CW is included in TEMU and separately counted as certified warehouse", () => {
  assert.equal(sourceType("TEMU正式分单", { 订单来源: "TEMU-CW" }), "TEMU");
  assert.equal(warehouseType("TEMU正式分单", { 订单来源: "TEMU-CW" }), "CERTIFIED_WAREHOUSE");
  const row = mappedRow({ sourceType: "TEMU", warehouseType: "CERTIFIED_WAREHOUSE" });
  const result = aggregateOperating("2026-08-28", [row], DEFAULT_SETTINGS, { rows: 1 }, "2026-08-28T00:00:00Z");
  assert.equal(result.totals.platforms.TEMU, 1);
  assert.equal(result.totals.platforms.CERTIFIED_WAREHOUSE, 1);
});

test("missing postcode + route mapping stays Unmapped", () => {
  const result = resolveRoute({ 二段码: "20-002", 订单来源: "SF" }, { district: "B13" }, { entries: {} }, "SF提货订单明细");
  assert.equal(result.matched, false);
  assert.equal(result.destinationCenter, "Unmapped");
  assert.equal(result.outlet, "未映射");
  assert.equal(result.code, "");
});

test("unique postcode mapping corrects a stale route code", () => {
  const entry = { postcodeOutward: "B31", routeCode: "20-002", outletCode: "UK200002", outletName: "C.BHX002", centerName: "Lutterworth HUB", enabled: true, updatedAt: "2026-08-29T00:00:00Z" };
  const result = resolveRoute({ 二段码: "20-020", 订单来源: "TEMU" }, { district: "B31" }, { entries: { "B31\u000120-002": entry }, postcodeIndex: { B31: [entry] } }, "TEMU正式分单");
  assert.equal(result.matched, true);
  assert.equal(result.correctedRoute, true);
  assert.equal(result.originalRouteCode, "20-020");
  assert.equal(result.routeCode, "20-002");
  assert.equal(result.outlet, "C.BHX002");
});

test("postcode conflict uses the latest uploaded mapping", () => {
  const first = { postcodeOutward: "B31", routeCode: "20-002", outletCode: "UK200002", outletName: "C.BHX002", centerName: "Lutterworth HUB", enabled: true, updatedAt: "2026-08-29T00:00:00Z" };
  const second = { postcodeOutward: "B31", routeCode: "20-003", outletCode: "UK200003", outletName: "C.LEICESTER001", centerName: "Lutterworth HUB", enabled: true, updatedAt: "2026-08-30T00:00:00Z" };
  const result = resolveRoute({ 二段码: "20-020" }, { district: "B31" }, { entries: {}, postcodeIndex: { B31: [second, first] } }, "TEMU正式分单");
  assert.equal(result.matched, true);
  assert.equal(result.correctedRoute, true);
  assert.equal(result.outlet, "C.LEICESTER001");
});

test("new route prefixes and transfer centres come from route maintenance", () => {
  const dartford = {
    postcodeOutward: "E16",
    routeCode: "13-012",
    outletCode: "UK100012",
    outletName: "C.LONDON012",
    centerCode: "DFD01",
    centerName: "Dartford Depot",
    enabled: true,
    updatedAt: "2026-09-24T00:00:00Z",
  };
  const map = { entries: { "E16\u000113-012": dartford }, postcodeIndex: { E16: [dartford] } };
  const resolved = resolveRoute({ 二段码: "10-012", 订单来源: "SF-CB" }, { district: "E16" }, map, "系统订单明细");
  assert.equal(resolved.matched, true);
  assert.equal(resolved.routeCode, "13-012");
  assert.equal(resolved.destinationCenter, "Dartford Depot");
  assert.equal(resolved.destinationCenterCode, "DFD01");

  const result = aggregateOperating("2026-09-27", [{ ...mappedRow(), ...resolved }], DEFAULT_SETTINGS, { rows: 1 }, "2026-09-27T00:00:00Z", new Map(), map);
  const dfd = result.centers.find((center) => center.center === "Dartford Depot");
  assert.equal(dfd.shortName, "DFD");
  assert.equal(dfd.inbound, 1);
  assert.equal(result.outlets[0].center, "Dartford Depot");
});

test("route maintenance accepts a new two-digit route prefix", () => {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.json_to_sheet([{
    网点编码: "UK100012",
    网点名称: "C.LONDON012",
    二段码: "13-012",
    末端机构编码: "DFD01",
    末端机构: "Dartford Depot",
    邮编: "E16",
    启用状态: "启用",
  }]), "二段码");
  const parsed = parseRouteMapWorkbook(book, { entries: {} });
  assert.equal(parsed.summary.enabled, 1);
  assert.equal(parsed.map.entries["E16\u000113-012"].centerName, "Dartford Depot");
});

test("vehicle is blank until manually maintained", () => {
  assert.equal(vehicleFor("Lutterworth HUB", 31, DEFAULT_SETTINGS, "2026-08-28"), "");
  assert.equal(vehicleFor("Lutterworth HUB", 100, DEFAULT_SETTINGS, "2026-08-28"), "");
  assert.equal(vehicleFor("Lutterworth HUB", 31, { ...DEFAULT_SETTINGS, vehicleOverrides: { "2026-08-28:Lutterworth HUB": "2 x 18T" } }, "2026-08-28"), "2 x 18T");
});

test("current date vehicle override is applied to stored datasets", () => {
  const date = "2026-09-10";
  const result = buildComparison({
    schemaVersion: 2,
    date,
    quality: {},
    centers: [{ center: "Lutterworth HUB", shortName: "LTW", estimatedCubes: 25, vehicle: "7.5T" }],
    outlets: [],
    totals: { platforms: { SF: 0, TT: 0, TEMU: 0, CERTIFIED_WAREHOUSE: 0 }, sourcePlatforms: {} },
  }, null, { ...DEFAULT_SETTINGS, vehicleOverrides: { [`${date}:Lutterworth HUB`]: "2 x 18T" } });
  assert.equal(result.centers[0].vehicle, "2 x 18T");
});

test("Operating separates outbound and same-centre without changing operation volume", () => {
  const rows = [
    mappedRow({ sourceType: "SF" }),
    mappedRow({ sourceType: "TT" }),
    mappedRow({ code: "UK100002", routeCode: "10-002", region: "South", outlet: "C.LONDON002", destinationCenter: "Uxbridge Depot", sourceType: "TEMU" }),
    mappedRow({ code: "", routeCode: "20-999", region: "Unmapped", outlet: "未映射", destinationCenter: "Unmapped", sourceType: "SF" }),
  ];
  const settings = { ...DEFAULT_SETTINGS, cubeFactor: 10 };
  const result = aggregateOperating("2026-08-28", rows, settings, { rows: rows.length, routeMapMissing: 1 }, "2026-08-29T00:00:00Z", new Map([["Lutterworth HUB", 5]]));
  const ltw = result.centers.find((center) => center.shortName === "LTW");
  const uxb = result.centers.find((center) => center.shortName === "UXB");
  assert.equal(ltw.outbound, 1);
  assert.equal(ltw.sameCenter, 2);
  assert.equal(ltw.operation, 8);
  assert.equal(uxb.inbound, 1);
  assert.equal(result.totals.operation, 9);
  assert.equal(result.totals.cbtPickup, 5);
  assert.deepEqual(result.totals.platforms, { SF: 1, TT: 1, TEMU: 1, CERTIFIED_WAREHOUSE: 0 });
});

test("Volume and Pickup aggregates retain platform counts and backlog", () => {
  const rows = [mappedRow({ sourceType: "SF" }), mappedRow({ sourceType: "TEMU" })];
  const existing = { outlets: [{ region: "Midlands", code: "UK200002", outlet: "C.BHX002", backlog: 7 }], views: {}, updatedAt: { backlog: "2026-08-27T00:00:00Z" } };
  const result = aggregateVolume("2026-08-29", rows, existing, "2026-08-29T00:00:00Z");
  assert.equal(result.totals.due, 2);
  assert.equal(result.totals.backlog, 7);
  assert.equal(result.totals.total, 9);
  assert.deepEqual(result.totals.platforms, { SF: 1, TT: 0, TEMU: 1, CERTIFIED_WAREHOUSE: 0 });
});

test("actual summary postcode placeholders retain their outlet allocation", () => {
  const row = mappedRow({
    sourceType: "OTHER",
    area: "Unknown",
    district: "Unknown",
    sector: "Unknown",
    full: "Postcode detail missing / 邮编明细缺失",
    detailMissing: true,
  });
  const result = aggregateVolume("2026-09-08", [row], null, "2026-09-07T20:00:00Z", { missingPostcodes: 1 });
  assert.equal(result.totals.due, 1);
  assert.equal(result.outlets[0].outlet, "C.BHX002");
  assert.equal(result.views["Full Postcode"][0].detailMissing, true);
});

test("missing inherited backlog postcode detail is preserved without blocking a new forecast", () => {
  const rows = completePostcodeCoverage(
    [{ region: "Midlands", code: "UK200002", outlet: "C.BHX002", due: 2, backlog: 7 }],
    [{ region: "Midlands", code: "UK200002", outlet: "C.BHX002", postal: "B13 8XX", due: 2, backlog: 0, total: 2 }],
  );
  assert.equal(rows.reduce((sum, row) => sum + row.due, 0), 2);
  assert.equal(rows.reduce((sum, row) => sum + row.backlog, 0), 7);
  assert.equal(rows.find((row) => row.detailMissing).postal, "Postcode detail missing / 邮编明细缺失");
});

test("packed postcode rows expand without losing due, backlog, or missing-detail state", () => {
  const result = expandPackedPostcodes({
    views: { District: [] },
    fullPostcodesPacked: [["Midlands", "UK200002", "C.BHX002", "B13 8XX", 2, 7, 0], ["Midlands", "UK200002", "C.BHX002", "Postcode detail missing / 邮编明细缺失", 0, 3, 1]],
  });
  assert.equal(result.views["Full Postcode"].reduce((sum, row) => sum + row.due, 0), 2);
  assert.equal(result.views["Full Postcode"].reduce((sum, row) => sum + row.backlog, 0), 10);
  assert.equal(result.views["Full Postcode"][1].detailMissing, true);
  assert.equal("fullPostcodesPacked" in result, false);
});

test("Operating hides zero-forecast backlog outlets while Volume keeps them", () => {
  const operating = {
    date: "2026-09-02",
    centers: [
      { center: "Uxbridge Depot", outletCount: 1 },
      { center: "Glasgow Depot", outletCount: 0 },
    ],
    outlets: [
      { center: "Uxbridge Depot", routeCode: "10-001", outlet: "LONDON001", deliveryVolume: 0, estimatedParcels: 0, platforms: { SF: 0, TT: 0, TEMU: 0 } },
    ],
    totals: { outletCount: 1 },
  };
  const volume = {
    date: "2026-09-03",
    outlets: [
      { region: "North", code: "UK400002", outlet: "C.GLASGOW002", due: 0, backlog: 11, total: 11, drivers: 0, platforms: { SF: 0, TT: 0, TEMU: 0 } },
      { region: "Unmapped", code: "", outlet: "Unknown Outlet", due: 0, backlog: 0, total: 0, drivers: 0, platforms: { SF: 0, TT: 0, TEMU: 0 } },
    ],
    totals: { due: 0, backlog: 11, total: 11, outlets: 2 },
  };
  const result = alignOutletDatasets(operating, volume);
  assert.deepEqual(result.operating.outlets, []);
  assert.deepEqual(result.volume.outlets.map((row) => row.outlet), ["C.GLASGOW002"]);
  assert.equal(result.operating.totals.outletCount, 0);
  assert.equal(result.volume.totals.outlets, 1);
});

test("zero mapped forecast groups cannot overwrite an existing business date", () => {
  assert.equal(hasMappableVolume({ totals: { due: 0 } }), false);
  assert.equal(hasMappableVolume({ totals: { due: 23060 } }), true);
});

test("comparison includes actual-only outlets and platform totals", () => {
  const operating = {
    schemaVersion: 2,
    date: "2026-08-28",
    settings: DEFAULT_SETTINGS,
    quality: {},
    centers: [{ center: "Uxbridge Depot", shortName: "UXB", operation: 3, outbound: 0, inbound: 3, sameCenter: 0, cbtPickup: 0, estimatedParcels: 0, estimatedCubes: 1, vehicle: "7.5T", outletCount: 1 }],
    outlets: [{ center: "Uxbridge Depot", outlet: "C.LONDON002", routeCode: "10-002", deliveryVolume: 3, estimatedParcels: 0, platforms: { SF: 2, TT: 1, TEMU: 0 } }],
    totals: { operation: 3, outbound: 0, inbound: 3, sameCenter: 0, cbtPickup: 0, estimatedParcels: 0, estimatedCubes: 1, outletCount: 1, platforms: { SF: 2, TT: 1, TEMU: 0 } },
  };
  const pickup = {
    schemaVersion: 2,
    quality: {},
    outlets: [
      { region: "South", code: "UK100002", outlet: "C.LONDON002", due: 4, platforms: { SF: 3, TT: 1, TEMU: 0 } },
      { region: "South", code: "UK100017", outlet: "C.LUTON001", due: 2, platforms: { SF: 0, TT: 0, TEMU: 2 } },
    ],
    totals: { due: 6, platforms: { SF: 3, TT: 1, TEMU: 2 } },
  };
  const result = buildComparison(operating, pickup);
  assert.equal(result.comparisonDate, "2026-08-28");
  assert.equal(result.totals.comparison.actualDue, 6);
  assert.equal(result.comparisonOutlets.length, 2);
  assert.equal(result.platformComparison.SF.forecast, 2);
  assert.equal(result.platformComparison.SF.actual, 3);
  assert.equal(result.platformComparison["TEMU-CB"].actual, 2);
  assert.equal(result.comparisonOutlets.find((row) => row.outlet === "C.LUTON001").comparison.diff, 2);
});

test("actual summary totals remain comparable without inventing a platform split", () => {
  const operating = {
    schemaVersion: 2,
    date: "2026-09-07",
    settings: DEFAULT_SETTINGS,
    quality: {},
    centers: [{ center: "Lutterworth HUB", shortName: "LTW", operation: 10, outbound: 0, inbound: 0, sameCenter: 10, cbtPickup: 0, estimatedParcels: 1, estimatedCubes: 1, vehicle: "7.5T", outletCount: 1 }],
    outlets: [{ center: "Lutterworth HUB", outlet: "C.BHX002", routeCode: "20-002", deliveryVolume: 10, estimatedParcels: 1, platforms: { SF: 5, TT: 3, TEMU: 2, CERTIFIED_WAREHOUSE: 0 } }],
    totals: { operation: 10, outbound: 0, inbound: 0, sameCenter: 10, cbtPickup: 0, estimatedParcels: 1, estimatedCubes: 1, outletCount: 1, platforms: { SF: 5, TT: 3, TEMU: 2, CERTIFIED_WAREHOUSE: 0 } },
  };
  const pickup = {
    platformsAvailable: false,
    quality: { summarySheet: "提货订单汇总（25个网点）" },
    outlets: [{ region: "Midlands", code: "UK200002", outlet: "C.BHX002", due: 12, platforms: { SF: 0, TT: 0, TEMU: 0, CERTIFIED_WAREHOUSE: 0 } }],
    totals: { due: 12, platforms: { SF: 0, TT: 0, TEMU: 0, CERTIFIED_WAREHOUSE: 0 } },
  };
  const result = buildComparison(operating, pickup);
  assert.equal(result.totals.comparison.actualDue, 12);
  assert.equal(result.actualPlatformsAvailable, false);
  assert.equal(result.platformComparison.SF.actual, null);
  assert.equal(result.comparisonOutlets[0].comparison.actualDue, 12);
});

test("legacy datasets discard calculated vehicles unless manually maintained", () => {
  const legacy = {
    date: "2026-08-27",
    settings: DEFAULT_SETTINGS,
    quality: {},
    centers: [{ center: "Lutterworth HUB", shortName: "LTW", operation: 15, outbound: 10, inbound: 0, sameCenter: 2, cbtPickup: 0, estimatedParcels: 20, estimatedCubes: 40, vehicle: "7.5T", outletCount: 0 }],
    outlets: [],
    totals: { operation: 15, outbound: 10, inbound: 0, sameCenter: 2, cbtPickup: 0, estimatedParcels: 20, estimatedCubes: 40, outletCount: 0 },
  };
  const result = buildComparison(legacy, null);
  assert.equal(result.centers[0].outbound, 12);
  assert.equal(result.centers[0].cbtPickup, 3);
  assert.equal(result.centers[0].vehicle, "");
});

test("dated manual fields merge legacy values with independently saved cells", async () => {
  const values = new Map([
    ["difference-reasons/2026-09-14.json", { reasons: { "platform:TT": "legacy", "platform:SF": "kept" } }],
    ["difference-reason-cells/2026-09-14/tt.json", { key: "platform:TT", value: "latest" }],
    ["vehicle-override-cells/2026-09-14/ltw.json", { center: "Lutterworth HUB", vehicle: "2 x 44T" }],
  ]);
  const store = {
    get: async (key) => values.get(key) || null,
    list: async ({ prefix }) => ({ blobs: [...values.keys()].filter((key) => key.startsWith(prefix)).map((key) => ({ key })) }),
  };
  assert.deepEqual(await differenceReasonsForDate(store, "2026-09-14"), { "platform:TT": "latest", "platform:SF": "kept" });
  assert.deepEqual(await vehicleOverridesForDate(store, "2026-09-14"), { "2026-09-14:Lutterworth HUB": "2 x 44T" });
});
