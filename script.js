const languageButtons = document.querySelectorAll("[data-lang-switch]");
const translatable = document.querySelectorAll("[data-en][data-zh]");
const menuToggle = document.querySelector("[data-menu-toggle]");
const siteNav = document.querySelector("[data-site-nav]");
const year = document.querySelector("[data-year]");
const footprintMap = document.querySelector("[data-footprint-map]");
const footprintTooltip = document.querySelector("[data-footprint-tooltip]");

function getSavedLanguage() {
  try {
    return window.localStorage.getItem("site-language") || "en";
  } catch {
    return "en";
  }
}

function saveLanguage(language) {
  try {
    window.localStorage.setItem("site-language", language);
  } catch {
    // Language switching should still work if local storage is unavailable.
  }
}

const savedLanguage = getSavedLanguage();
let currentLanguage = savedLanguage === "zh" ? "zh" : "en";
const footprintPlaces = [
  {
    x: 49.2,
    y: 40.8,
    countryEn: "United Kingdom & Ireland Area",
    countryZh: "英国及周边",
    placeEn: "United Kingdom Footprints",
    placeZh: "英国足迹",
    noteEn: "Study life and travel across UK counties and cities, including Liverpool and wider northern England.",
    noteZh: "英国学习生活与旅行足迹，覆盖利物浦及英格兰北部、伦敦等城市与郡。",
    typeEn: "Study · Travel",
    typeZh: "学习 · 旅行",
    spreadX: 5.8,
    spreadY: 5.2,
    cities: ["牛津郡", "德比郡", "北约克郡", "东北林肯郡", "格洛斯特郡", "北林肯郡", "诺丁汉郡", "东约克郡", "南约克郡", "赫尔", "西约克郡", "柴郡", "白金汉郡", "伦敦", "贝德福德郡", "卢顿", "米尔顿凯恩斯", "曼彻斯特", "默西赛德郡"]
  },
  {
    x: 50.3,
    y: 45.2,
    countryEn: "Europe",
    countryZh: "欧洲",
    placeEn: "Continental Europe",
    placeZh: "欧洲大陆",
    noteEn: "Travel footprints across France, the Netherlands and Spain-linked islands.",
    noteZh: "法国、荷兰及西班牙相关岛屿等欧洲旅行足迹。",
    typeEn: "Travel",
    typeZh: "旅行",
    spreadX: 5,
    spreadY: 3.8,
    cities: ["芒什省", "奥恩省", "巴黎省", "塞纳-马恩省", "北荷兰省", "拉斯帕尔马斯"]
  },
  {
    x: 72,
    y: 46.5,
    countryEn: "China",
    countryZh: "中国",
    placeEn: "North China & Northeast",
    placeZh: "华北与东北",
    noteEn: "A dense cluster of education, family, travel, and regional business footprints across northern China.",
    noteZh: "覆盖华北、东北的学习、生活、旅行与区域商业实践足迹。",
    typeEn: "Education · Work · Travel",
    typeZh: "教育 · 工作 · 旅行",
    spreadX: 8.5,
    spreadY: 7.4,
    cities: ["北京", "天津", "石家庄市", "唐山市", "承德市", "廊坊市", "秦皇岛市", "衡水市", "太原市", "晋中市", "忻州市", "通辽市", "哈尔滨市", "长春市", "延边朝鲜族自治州", "沈阳市", "大连市", "盘锦市", "营口市", "锦州市", "鞍山市", "辽阳市", "本溪市", "朝阳市"]
  },
  {
    x: 73.8,
    y: 52,
    countryEn: "China",
    countryZh: "中国",
    placeEn: "East, Central & South China",
    placeZh: "华东、华中与华南",
    noteEn: "City badges across commercial, cultural and coastal regions, including Shandong, Jiangsu, Zhejiang, Shanghai and Guangdong.",
    noteZh: "覆盖山东、江苏、浙江、上海、广东及华中等商业、文化与沿海地区的城市徽章。",
    typeEn: "Business · Travel",
    typeZh: "商业 · 旅行",
    spreadX: 9,
    spreadY: 8,
    cities: ["济南市", "青岛市", "烟台市", "淄博市", "德州市", "滨州市", "临沂市", "泰安市", "枣庄市", "苏州市", "徐州市", "宿州市", "杭州市", "上海市", "南昌市", "武汉市", "长沙市", "郑州市", "深圳市", "广州市", "南充市", "成都市", "绵阳市"]
  },
  {
    x: 69,
    y: 54.8,
    countryEn: "China",
    countryZh: "中国",
    placeEn: "Northwest & Southwest China",
    placeZh: "西北与西南",
    noteEn: "Long-distance travel and cultural landscapes across western China, including Tibetan and Qinghai regions.",
    noteZh: "西部中国的长距离旅行与文化地理足迹，包括西藏、青海、甘肃、陕西、云南等区域。",
    typeEn: "Travel · Culture",
    typeZh: "旅行 · 文化",
    spreadX: 8,
    spreadY: 8,
    cities: ["重庆市", "榆林市", "商洛市", "拉萨市", "那曲市", "林芝市", "山南市", "海西蒙古族藏族自治州", "西宁市", "海北藏族自治州", "海东市", "兰州市", "定西市", "天水市", "西安市", "咸阳市"]
  },
  {
    x: 78.2,
    y: 49.6,
    countryEn: "Japan & Korea",
    countryZh: "日本与韩国",
    placeEn: "Japan & Jeju",
    placeZh: "日本与济州",
    noteEn: "East Asian city footprints from Japan and Jeju.",
    noteZh: "日本与济州相关的东亚城市足迹。",
    typeEn: "Travel",
    typeZh: "旅行",
    spreadX: 4,
    spreadY: 3,
    cities: ["济州市", "千叶县", "神奈川县", "东京都"]
  },
  {
    x: 76.1,
    y: 63.5,
    countryEn: "Hong Kong & Southeast Asia",
    countryZh: "香港与东南亚",
    placeEn: "Hong Kong",
    placeZh: "香港",
    noteEn: "Asian city and travel footprints connected to global business and urban life.",
    noteZh: "与全球商业和城市生活相关的亚洲城市与旅行足迹。",
    typeEn: "Travel · Business Context",
    typeZh: "旅行 · 商业语境",
    spreadX: 4,
    spreadY: 3.2,
    cities: ["香港"]
  },
  {
    x: 62,
    y: 57.5,
    countryEn: "Middle East",
    countryZh: "中东",
    placeEn: "Abu Dhabi",
    placeZh: "阿布扎比",
    noteEn: "A long-distance footprint represented in the badge map.",
    noteZh: "徽章地图中的远距离足迹。",
    typeEn: "Travel",
    typeZh: "旅行",
    spreadX: 5.5,
    spreadY: 4,
    cities: ["阿布扎比"]
  }
];

const cityCoordinates = {
  "牛津郡": [51.75, -1.25], "德比郡": [53.1, -1.55], "北约克郡": [54.15, -1.35], "东北林肯郡": [53.55, -0.08], "格洛斯特郡": [51.86, -2.24], "北林肯郡": [53.6, -0.65], "诺丁汉郡": [53.1, -1.0], "东约克郡": [53.84, -0.43], "南约克郡": [53.48, -1.35], "赫尔": [53.75, -0.34], "西约克郡": [53.8, -1.55], "柴郡": [53.2, -2.6], "白金汉郡": [51.82, -0.82], "伦敦": [51.51, -0.13], "贝德福德郡": [52.14, -0.47], "卢顿": [51.88, -0.42], "米尔顿凯恩斯": [52.04, -0.76], "曼彻斯特": [53.48, -2.24], "默西赛德郡": [53.41, -2.99],
  "芒什省": [49.05, -1.25], "奥恩省": [48.57, 0.08], "巴黎省": [48.86, 2.35], "塞纳-马恩省": [48.84, 2.65], "北荷兰省": [52.52, 4.95], "拉斯帕尔马斯": [28.12, -15.43],
  "北京": [39.9, 116.4], "天津": [39.08, 117.2], "石家庄市": [38.04, 114.51], "唐山市": [39.63, 118.18], "承德市": [40.95, 117.96], "廊坊市": [39.52, 116.7], "秦皇岛市": [39.94, 119.6], "衡水市": [37.74, 115.67], "太原市": [37.87, 112.55], "晋中市": [37.69, 112.75], "忻州市": [38.42, 112.73], "通辽市": [43.65, 122.24], "哈尔滨市": [45.8, 126.53], "长春市": [43.82, 125.32], "延边朝鲜族自治州": [42.9, 129.5], "沈阳市": [41.8, 123.43], "大连市": [38.91, 121.61], "盘锦市": [41.12, 122.07], "营口市": [40.67, 122.23], "锦州市": [41.1, 121.13], "鞍山市": [41.11, 122.99], "辽阳市": [41.27, 123.17], "本溪市": [41.29, 123.77], "朝阳市": [41.57, 120.45],
  "济南市": [36.65, 117.12], "青岛市": [36.07, 120.38], "烟台市": [37.46, 121.45], "淄博市": [36.81, 118.05], "德州市": [37.45, 116.31], "滨州市": [37.38, 117.97], "临沂市": [35.1, 118.35], "泰安市": [36.2, 117.09], "枣庄市": [34.81, 117.32], "苏州市": [31.3, 120.58], "徐州市": [34.26, 117.2], "宿州市": [33.65, 116.96], "杭州市": [30.27, 120.16], "上海市": [31.23, 121.47], "南昌市": [28.68, 115.86], "武汉市": [30.59, 114.3], "长沙市": [28.23, 112.94], "郑州市": [34.75, 113.62], "深圳市": [22.54, 114.06], "广州市": [23.13, 113.26], "南充市": [30.84, 106.11], "成都市": [30.67, 104.06], "绵阳市": [31.47, 104.68],
  "重庆市": [29.56, 106.55], "榆林市": [38.29, 109.73], "商洛市": [33.87, 109.94], "拉萨市": [29.65, 91.13], "那曲市": [31.48, 92.06], "林芝市": [29.65, 94.36], "山南市": [29.24, 91.77], "海西蒙古族藏族自治州": [37.38, 97.37], "西宁市": [36.62, 101.78], "海北藏族自治州": [36.95, 100.9], "海东市": [36.5, 102.1], "兰州市": [36.06, 103.84], "定西市": [35.58, 104.63], "天水市": [34.58, 105.72], "西安市": [34.34, 108.94], "咸阳市": [34.33, 108.71],
  "济州市": [33.5, 126.53], "千叶县": [35.6, 140.12], "神奈川县": [35.45, 139.64], "东京都": [35.68, 139.76],
  "香港": [22.32, 114.17],
  "阿布扎比": [24.45, 54.38]
};

function getCityPoint(city, place, index, projection) {
  const coordinate = cityCoordinates[city];
  if (coordinate) {
    const projected = projection([coordinate[1], coordinate[0]]);
    if (projected) {
      return projected;
    }
  }

  const angle = index * 2.399963;
  const radius = 0.4 + Math.sqrt(index + 1) * 0.28;
  return [place.x + Math.cos(angle) * radius, place.y + Math.sin(angle) * radius];
}

let worldAtlasPromise;

function loadWorldAtlas() {
  if (window.__XIHAO_WORLD_ATLAS__) {
    return Promise.resolve(window.__XIHAO_WORLD_ATLAS__);
  }

  if (!worldAtlasPromise) {
    worldAtlasPromise = fetch("assets/countries-110m.json").then((response) => response.json());
  }

  return worldAtlasPromise;
}

function markerOffset(index) {
  const angle = index * 2.399963;
  const radius = 1.2 + (index % 5) * 0.55;
  return [Math.cos(angle) * radius, Math.sin(angle) * radius];
}

function scaledMarkerRadius(baseRadius, zoomScale) {
  return baseRadius / Math.pow(zoomScale, 1.2);
}

function getFootprintCities() {
  return footprintPlaces.flatMap((place) =>
    place.cities.map((city, cityIndex) => ({
      city,
      cityIndex,
      place,
      label: currentLanguage === "zh" ? place.placeZh : place.placeEn
    }))
  );
}

function moveFootprintTooltip(event) {
  if (!footprintTooltip) return;

  const mapBounds = footprintMap.getBoundingClientRect();
  const x = event.clientX - mapBounds.left;
  const y = event.clientY - mapBounds.top;
  footprintTooltip.style.left = `${x}px`;
  footprintTooltip.style.top = `${y}px`;
}

function showFootprintTooltip(text, event) {
  if (!footprintTooltip) return;

  footprintTooltip.textContent = text;
  footprintTooltip.classList.add("is-visible");
  if (event) moveFootprintTooltip(event);
}

function hideFootprintTooltip() {
  footprintTooltip?.classList.remove("is-visible");
}

async function renderFootprints() {
  if (!footprintMap || !window.d3 || !window.topojson) return;

  const bounds = footprintMap.getBoundingClientRect();
  const width = Math.max(320, Math.round(bounds.width || 960));
  const height = Math.round(width * 1163 / 1920);
  const atlas = await loadWorldAtlas().catch(() => null);
  if (!atlas) return;
  const land = topojson.feature(atlas, atlas.objects.countries);
  const projection = d3.geoNaturalEarth1().fitExtent([[10, 8], [width - 10, height - 8]], { type: "Sphere" });
  const path = d3.geoPath(projection);
  const svg = d3.select(footprintMap);
  const cities = getFootprintCities();

  svg.attr("viewBox", `0 0 ${width} ${height}`).attr("preserveAspectRatio", "xMidYMid meet");
  svg.selectAll("*").remove();
  svg.on(".zoom", null);

  const zoomLayer = svg.append("g").attr("class", "map-zoom-layer");

  zoomLayer.append("path")
    .datum({ type: "Sphere" })
    .attr("class", "map-outline")
    .attr("d", path);

  zoomLayer.append("g")
    .selectAll("path")
    .data(land.features)
    .join("path")
    .attr("class", "map-land")
    .attr("d", path);

  zoomLayer.append("g")
    .selectAll("circle")
    .data(cities)
    .join("circle")
    .attr("class", "city-marker")
    .style("fill", "var(--ink)")
    .attr("data-base-radius", width < 720 ? 2.4 : 3.2)
    .attr("cx", (entry, index) => {
      const [x] = getCityPoint(entry.city, entry.place, entry.cityIndex, projection);
      const [offsetX] = markerOffset(index);
      return x + offsetX;
    })
    .attr("cy", (entry, index) => {
      const [, y] = getCityPoint(entry.city, entry.place, entry.cityIndex, projection);
      const [, offsetY] = markerOffset(index);
      return y + offsetY;
    })
    .attr("r", width < 720 ? 2.4 : 3.2)
    .attr("tabindex", 0)
    .attr("aria-label", (entry) => `${entry.city} · ${entry.label}`)
    .on("mouseenter", (event, entry) => showFootprintTooltip(entry.city, event))
    .on("mousemove", (event) => moveFootprintTooltip(event))
    .on("mouseleave", hideFootprintTooltip)
    .on("focus", function (event, entry) {
      const mapBounds = footprintMap.getBoundingClientRect();
      const markerBounds = this.getBoundingClientRect();
      showFootprintTooltip(entry.city);
      footprintTooltip.style.left = `${markerBounds.left + markerBounds.width / 2 - mapBounds.left}px`;
      footprintTooltip.style.top = `${markerBounds.top - mapBounds.top}px`;
    })
    .on("blur", hideFootprintTooltip)
    .append("title")
    .text((entry) => `${entry.city} · ${entry.label}`);

  const zoomBehavior = d3.zoom()
    .scaleExtent([1, 7])
    .translateExtent([[0, 0], [width, height]])
    .extent([[0, 0], [width, height]])
    .on("start", () => {
      hideFootprintTooltip();
      footprintMap.classList.add("is-panning");
    })
    .on("zoom", (event) => {
      zoomLayer.attr("transform", event.transform);
      zoomLayer.selectAll(".city-marker")
        .attr("r", function () {
          return scaledMarkerRadius(Number(this.dataset.baseRadius), event.transform.k);
        })
        .style("fill", "var(--ink)")
        .attr("stroke-width", 1.2 / Math.pow(event.transform.k, 1.2));
    })
    .on("end", () => {
      footprintMap.classList.remove("is-panning");
    });

  svg.call(zoomBehavior);
}

function setLanguage(language) {
  const activeLanguage = language === "zh" ? "zh" : "en";
  currentLanguage = activeLanguage;
  document.documentElement.lang = activeLanguage === "zh" ? "zh-CN" : "en-GB";
  document.body.dataset.language = activeLanguage;

  translatable.forEach((element) => {
    element.textContent = element.dataset[activeLanguage];
  });

  document.querySelectorAll("[data-title-en][data-title-zh]").forEach((element) => {
    element.setAttribute("title", element.dataset[`title${activeLanguage === "zh" ? "Zh" : "En"}`]);
  });

  languageButtons.forEach((button) => {
    const isActive = button.dataset.langSwitch === activeLanguage;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
  });

  saveLanguage(activeLanguage);
  renderFootprints().catch(() => {});
}

languageButtons.forEach((button) => {
  button.addEventListener("click", () => setLanguage(button.dataset.langSwitch));
});

menuToggle?.addEventListener("click", () => {
  const isOpen = siteNav?.classList.toggle("is-open");
  menuToggle.setAttribute("aria-expanded", String(Boolean(isOpen)));
});

document.querySelectorAll("[data-site-nav] a").forEach((link) => {
  link.addEventListener("click", () => {
    siteNav?.classList.remove("is-open");
    menuToggle?.setAttribute("aria-expanded", "false");
  });
});

if (year) {
  year.textContent = new Date().getFullYear();
}

setLanguage(savedLanguage);

if (footprintMap && "ResizeObserver" in window) {
  const mapResizeObserver = new ResizeObserver(() => {
    renderFootprints().catch(() => {});
  });
  mapResizeObserver.observe(footprintMap);
}
