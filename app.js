const DATA_URL = "data/aaslt-db.json";
const WEATHER_LOCATION = "West Point, NY";
const WEATHER_URL =
  "https://api.open-meteo.com/v1/forecast?latitude=41.3915&longitude=-73.9559&current=temperature_2m,relative_humidity_2m,weather_code&daily=temperature_2m_max,temperature_2m_min,weather_code&hourly=temperature_2m,relative_humidity_2m&temperature_unit=fahrenheit&timezone=America%2FNew_York&forecast_days=16";
const TASKS_KEY = "aaslt.tasks.v1";
const S4_KEY = "aaslt.s4.items.v3";
const DELETED_SOURCE_KEY = "aaslt.deletedSource.v1";
const SETTINGS_KEY = "aaslt.settings.v2";
const EVENT_OVERRIDES_KEY = "aaslt.eventOverrides.v2";
const RECEIPT_DB = "aaslt-receipts";
const RECEIPT_STORE = "files";
const CANONICAL_LRTC_FILE = "Total AASLT Cadre LRTC (1).xlsx";
const CLOUD_SYNC_SCHEMA = 1;

const TRACKED_PEOPLE = [
  "Unassigned",
  "Noah Ginty",
  "Thomas Gargan",
  "Massimo Luciano",
  "Brian Samuel",
  "Anant Sabata",
  "Chinmay Satpute",
  "Chadwick Nash",
  "Patrick Dunn",
  "MAJ Liesenfelt",
  "2LT Schneider",
  "2LT Goldstien",
];
const ASSIGNABLE_PEOPLE = TRACKED_PEOPLE.filter((person) => person !== "Unassigned");

const S4_LANES = ["Draw", "Water Buffalo", "LMTV TMR", "FLA Request"];
const S4_STATUS_OPTIONS = {
  Draw: ["Planned", "Drawn", "Complete"],
  "Water Buffalo": ["Planned", "Sent", "Complete"],
  "LMTV TMR": ["Planned", "Sent", "Complete"],
  "FLA Request": ["Planned", "Sent", "Complete"],
};
const S4_FIELD_CONFIG = {
  Draw: {
    visible: ["item", "qty", "needBy", "status", "notes"],
    itemLabel: "Draw Item",
    itemPlaceholder: "e.g. Skedco litter, radios, cones",
    qtyLabel: "Qty",
    timeLabel: "Draw Time",
    required: ["item"],
  },
  "Water Buffalo": {
    visible: ["action", "location", "needBy", "status", "notes"],
    timeLabel: "TMR Time",
    required: ["location"],
  },
  "LMTV TMR": {
    visible: ["item", "pax", "needBy", "pickup", "return", "status", "notes"],
    itemLabel: "Mission / Pickup Point",
    itemPlaceholder: "e.g. Enabler PLT movement",
    timeLabel: "Request Time",
    required: ["item"],
  },
  "FLA Request": {
    visible: ["qty", "placement", "needBy", "status", "notes"],
    qtyLabel: "FLAs",
    timeLabel: "Request Time",
    required: ["placement"],
  },
};

const CATEGORY_ORDER = ["Training", "Instruction", "Medical", "Meals", "Cadre", "Logistics", "Operations"];
const CATEGORY_CLASS = {
  Training: "cat-training",
  Instruction: "cat-instruction",
  Medical: "cat-medical",
  Meals: "cat-meals",
  Cadre: "cat-cadre",
  Logistics: "cat-logistics",
  Operations: "cat-operations",
};

const MEAL_TIMES = {
  Breakfast: "07:30",
  Lunch: "11:30",
  Dinner: "17:30",
};

const CADET_MESS_TIMES = {
  Breakfast: "07:30",
  Lunch: "12:00",
  Dinner: "18:00",
};

const WEATHER_CODE_LABELS = {
  0: "Clear",
  1: "Mainly clear",
  2: "Partly cloudy",
  3: "Overcast",
  45: "Fog",
  48: "Rime fog",
  51: "Light drizzle",
  53: "Drizzle",
  55: "Heavy drizzle",
  56: "Freezing drizzle",
  57: "Freezing drizzle",
  61: "Light rain",
  63: "Rain",
  65: "Heavy rain",
  66: "Freezing rain",
  67: "Freezing rain",
  71: "Light snow",
  73: "Snow",
  75: "Heavy snow",
  77: "Snow grains",
  80: "Rain showers",
  81: "Rain showers",
  82: "Heavy showers",
  85: "Snow showers",
  86: "Heavy snow showers",
  95: "Thunderstorms",
  96: "Storms with hail",
  99: "Storms with hail",
};

const state = {
  data: null,
  rawEvents: [],
  events: [],
  selectedDate: null,
  viewMonth: null,
  sourceFilter: "all",
  search: "",
  dayView: "tracks",
  assignmentView: "list",
  tasks: [],
  editingTaskId: null,
  taskFormOpen: false,
  deletedSource: { taskings: [], supportItems: [] },
  eventOverrides: {},
  s4Items: [],
  receipts: [],
  weather: {
    status: "loading",
    days: [],
    updatedAt: "",
    error: "",
  },
  cloud: {
    enabled: false,
    ready: false,
    applying: false,
    saveTimer: null,
    docRef: null,
    storage: null,
    setDoc: null,
    serverTimestamp: null,
    storageRef: null,
    uploadBytes: null,
    getDownloadURL: null,
    deleteObject: null,
    lastError: "",
  },
  sourceIndex: 0,
  sheetIndex: 0,
  sourceSearch: "",
  classLookup: new Map(),
  classDayLabels: new Map(),
};

const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => Array.from(root.querySelectorAll(selector));

function createElement(tag, options = {}) {
  const el = document.createElement(tag);
  if (options.className) el.className = options.className;
  if (options.text !== undefined) el.textContent = options.text;
  if (options.html !== undefined) el.innerHTML = options.html;
  if (options.attrs) Object.entries(options.attrs).forEach(([key, value]) => el.setAttribute(key, value));
  if (options.children) el.append(...options.children);
  return el;
}

function refreshIcons() {
  if (window.lucide) window.lucide.createIcons();
}

function inflateSourcesFromTables(tables = {}) {
  if (!tables.sources || !tables.sheets || !tables.cells) return [];
  const sources = tables.sources.map((source) => ({ ...source, sheets: [] }));
  const sourceById = new Map(sources.map((source) => [source.id, source]));
  const sheetById = new Map();
  tables.sheets.forEach((sheet) => {
    const inflatedSheet = {
      name: sheet.name,
      maxRow: sheet.maxRow,
      maxColumn: sheet.maxColumn,
      cells: [],
    };
    sheetById.set(sheet.id, inflatedSheet);
    sourceById.get(sheet.sourceId)?.sheets.push(inflatedSheet);
  });
  tables.cells.forEach((cell) => {
    sheetById.get(cell.sheetId)?.cells.push({
      r: cell.row,
      c: cell.col,
      a: cell.address,
      v: cell.value,
      f: cell.formula,
    });
  });
  sources.forEach((source) => source.sheets.forEach((sheet) => sheet.cells.sort((a, b) => a.r - b.r || a.c - b.c)));
  return sources;
}

function sourceRefToLegacy(ref = {}, lookup = {}) {
  const source = lookup.sourcesById?.get(ref.sourceId);
  const sheet = lookup.sheetsById?.get(ref.sheetId);
  return {
    file: ref.file || source?.name || "",
    sheet: ref.sheet || sheet?.name || "",
    cell: ref.address || "",
    row: ref.row,
    col: ref.col,
    sourceId: ref.sourceId,
    sheetId: ref.sheetId,
    cellId: ref.cellId,
    notesCell: ref.notesCell,
    originalNotes: ref.originalNotes,
  };
}

function normalizeLoadedData(payload) {
  if (!payload.tables) return { ...payload, diagnostics: payload.diagnostics || [] };
  const tables = payload.tables;
  const sourcesById = new Map((tables.sources || []).map((source) => [source.id, source]));
  const sheetsById = new Map((tables.sheets || []).map((sheet) => [sheet.id, sheet]));
  const lookup = { sourcesById, sheetsById };
  return {
    generatedAt: payload.generatedAt,
    operation: payload.operation,
    summary: payload.summary,
    events: (tables.events || []).map((event) => {
      const relatedSources = (event.sourceRefs || []).map((ref) => sourceRefToLegacy(ref, lookup));
      return {
        ...event,
        source: relatedSources[0] || null,
        relatedSources,
      };
    }),
    taskings: (tables.taskings || []).map((tasking) => {
      const relatedSources = (tasking.sourceRefs || []).map((ref) => sourceRefToLegacy(ref, lookup));
      return {
        ...tasking,
        source: relatedSources[0] || null,
        relatedSources,
      };
    }),
    supportItems: (tables.supportItems || []).map((item) => {
      const relatedSources = (item.sourceRefs || []).map((ref) => sourceRefToLegacy(ref, lookup));
      return {
        ...item,
        source: relatedSources[0] || null,
        relatedSources,
      };
    }),
    notes: (tables.notes || []).map((note) => ({
      ...note,
      sourceRefs: (note.sourceRefs || []).map((ref) => sourceRefToLegacy(ref, lookup)),
    })),
    diagnostics: (tables.diagnostics || []).map((item) => ({
      ...item,
      sourceRefs: (item.sourceRefs || []).map((ref) => sourceRefToLegacy(ref, lookup)),
    })),
    sources: inflateSourcesFromTables(tables),
    tables,
  };
}

function parseDate(iso) {
  const [year, month, day] = iso.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function toIso(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function addDays(iso, amount) {
  const date = parseDate(iso);
  date.setDate(date.getDate() + amount);
  return toIso(date);
}

function formatFullDate(iso) {
  return parseDate(iso).toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric", year: "numeric" });
}

function formatCompactDate(iso) {
  return parseDate(iso).toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

function formatMonth(date) {
  return date.toLocaleDateString(undefined, { month: "long", year: "numeric" });
}

function formatShortWeekday(iso) {
  return parseDate(iso).toLocaleDateString(undefined, { weekday: "short" });
}

function formatTemp(value) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) return "--";
  return `${Math.round(Number(value))}\u00b0`;
}

function weatherCodeLabel(code) {
  return WEATHER_CODE_LABELS[Number(code)] || "Forecast";
}

function heatIndexF(temp, humidity) {
  const t = Number(temp);
  const rh = Number(humidity);
  if (Number.isNaN(t) || Number.isNaN(rh)) return undefined;
  if (t < 80) return t;
  let hi =
    -42.379 +
    2.04901523 * t +
    10.14333127 * rh -
    0.22475541 * t * rh -
    0.00683783 * t * t -
    0.05481717 * rh * rh +
    0.00122874 * t * t * rh +
    0.00085282 * t * rh * rh -
    0.00000199 * t * t * rh * rh;
  if (rh < 13 && t >= 80 && t <= 112) hi -= ((13 - rh) / 4) * Math.sqrt((17 - Math.abs(t - 95)) / 17);
  if (rh > 85 && t >= 80 && t <= 87) hi += ((rh - 85) / 10) * ((87 - t) / 5);
  return hi;
}

function humidityByDate(hourly = {}) {
  const totals = {};
  (hourly.time || []).forEach((time, index) => {
    const humidity = Number((hourly.relative_humidity_2m || [])[index]);
    if (!time || Number.isNaN(humidity)) return;
    const iso = String(time).slice(0, 10);
    totals[iso] ||= { sum: 0, count: 0 };
    totals[iso].sum += humidity;
    totals[iso].count += 1;
  });
  return Object.fromEntries(Object.entries(totals).map(([iso, value]) => [iso, Math.round(value.sum / value.count)]));
}

function heatIndexByDate(hourly = {}) {
  const maxByDate = {};
  (hourly.time || []).forEach((time, index) => {
    if (!time) return;
    const heatIndex = heatIndexF((hourly.temperature_2m || [])[index], (hourly.relative_humidity_2m || [])[index]);
    if (heatIndex === undefined) return;
    const iso = String(time).slice(0, 10);
    maxByDate[iso] = Math.max(maxByDate[iso] ?? heatIndex, heatIndex);
  });
  return Object.fromEntries(Object.entries(maxByDate).map(([iso, value]) => [iso, Math.round(value)]));
}

function normalizeWeather(payload = {}) {
  const daily = payload.daily || {};
  const humidity = humidityByDate(payload.hourly || {});
  const heatIndex = heatIndexByDate(payload.hourly || {});
  const todayIso = String(payload.current?.time || "").slice(0, 10);
  const currentHumidity = Number(payload.current?.relative_humidity_2m);
  const currentHeatIndex = heatIndexF(payload.current?.temperature_2m, payload.current?.relative_humidity_2m);
  return (daily.time || []).map((iso, index) => ({
    iso,
    label: index === 0 ? "Today" : index === 1 ? "Tomorrow" : formatShortWeekday(iso),
    weekday: formatShortWeekday(iso),
    condition: weatherCodeLabel((daily.weather_code || [])[index]),
    high: Math.round(Number((daily.temperature_2m_max || [])[index])),
    low: Math.round(Number((daily.temperature_2m_min || [])[index])),
    humidity: iso === todayIso && !Number.isNaN(currentHumidity) ? Math.round(currentHumidity) : humidity[iso],
    heatIndex: iso === todayIso && currentHeatIndex !== undefined ? Math.max(Math.round(currentHeatIndex), heatIndex[iso] || Math.round(currentHeatIndex)) : heatIndex[iso],
  }));
}

function escapeRegex(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char]);
}

function safeClassName(value) {
  return (
    String(value || "none")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-|-$/g, "") || "none"
  );
}

function cleanTitle(text) {
  return String(text || "")
    .replace(/\s+/g, " ")
    .replace(/\s*\/\s*/g, "/")
    .replace(/\bRecieve\b/gi, "Receive")
    .replace(/\bClassrom\b/gi, "Classroom")
    .replace(/\bRapelling\b/gi, "Rappelling")
    .replace(/\bAssualt\b/gi, "Assault")
    .trim();
}

function normalizeTime(value) {
  if (!value) return "";
  const match = String(value).match(/(\d{1,2}):?(\d{2})/);
  if (!match) return "";
  return `${match[1].padStart(2, "0")}:${match[2]}`;
}

function mealNameFromText(text) {
  const upper = String(text || "").toUpperCase();
  if (upper.includes("BREAKFAST")) return "Breakfast";
  if (upper.includes("LUNCH")) return "Lunch";
  if (upper.includes("DINNER")) return "Dinner";
  return "";
}

function mealStatusFromText(text) {
  const upper = String(text || "").toUpperCase();
  if (upper.includes("MESS HALL") || upper.includes("CADET MESS") || /\(CM\)/.test(upper) || upper.endsWith(": CM")) return "CM";
  if (upper.includes("MRE") || /\(M\)/.test(upper) || /\(M\+\)/.test(upper) || /\(HM\+\)/.test(upper)) return "MRE";
  if (upper.includes("A-FRAME") || /\(AF\)/.test(upper)) return "A-Frame";
  return "";
}

function mealStatusLabel(status) {
  if (status === "CM") return "Cadet Mess";
  if (status === "MRE") return "MRE";
  if (status === "A-Frame") return "A-Frame";
  return status || "TBD meal type";
}

function isCadetMessStatus(status) {
  return status === "CM";
}

function eventSourceFiles(event) {
  return (event.relatedSources || [event.source].filter(Boolean)).map((source) => source?.file).filter(Boolean);
}

function isClassesMealSource(event) {
  return event.sourceKind === "class-calendar" || eventSourceFiles(event).includes("CLASSES.xlsx");
}

function mealScheduleScore(event) {
  let score = 0;
  if (isClassesMealSource(event)) score += 100;
  if (event.start) score += 20;
  score += mealLocationScore(event);
  return score;
}

function bestScheduleMeal(events) {
  return [...events].sort((a, b) => mealScheduleScore(b) - mealScheduleScore(a) || (a.start || "99:99").localeCompare(b.start || "99:99"))[0];
}

function mergeScheduleMealGroup(events) {
  const best = bestScheduleMeal(events);
  return {
    ...best,
    relatedSources: dedupeSources(events.flatMap((event) => event.relatedSources || [event.source].filter(Boolean))),
  };
}

function mealStartFor(status, meal, schedule) {
  if (isCadetMessStatus(status)) return CADET_MESS_TIMES[meal] || MEAL_TIMES[meal] || "";
  return schedule?.start || MEAL_TIMES[meal] || "";
}

function mealLocationFor(status, schedule) {
  if (isCadetMessStatus(status)) return "Cadet Mess Hall";
  return schedule?.location || (status ? mealStatusLabel(status) : "");
}

function mealNotesFor(status, meal, schedule) {
  if (isCadetMessStatus(status)) return `Meal type from mess matrix; Cadet Mess fixed ${meal.toLowerCase()} time.`;
  if (status === "MRE") {
    return schedule?.start
      ? "Meal type from mess matrix; MRE time from CLASSES schedule."
      : "Meal type from mess matrix; no CLASSES schedule time found, so a planning default is shown.";
  }
  if (status) return "Meal type from mess matrix; time/location from schedule when available.";
  return "Time/location from schedule; no mess matrix status found.";
}

function inferClassKey(event) {
  if (event.classKey) return event.classKey;
  const text = `${event.title || ""} ${event.group || ""} ${event.notes || ""}`.toUpperCase();
  const airMatch = text.match(/\bAIR\s*ASSAULT\s*(\d)\b/);
  if (airMatch) return `Air Assault ${airMatch[1]}`;
  const aaMatch = text.match(/\bAA\s*(\d)\b/);
  if (aaMatch) return `Air Assault ${aaMatch[1]}`;
  const source = event.source || {};
  const lookup = state.classLookup.get(`${source.file}|${source.sheet}|${source.row}|${source.col}`);
  return lookup || "";
}

function buildClassLookup() {
  state.classLookup = new Map();
  const classes = state.data.sources.find((source) => source.name === "CLASSES.xlsx");
  const sheet = classes?.sheets.find((item) => item.name === "SCHED WORKING");
  if (!sheet) return;
  const headings = sheet.cells.filter((cell) => /DAY\s+[-\d]+\s+WP\d+/i.test(String(cell.v || "")));
  for (const heading of headings) {
    const wp = String(heading.v).match(/WP(\d+)/i);
    if (!wp) continue;
    const number = Number(wp[1]) - 699;
    if (number < 1 || number > 9) continue;
    for (let row = heading.r + 2; row <= heading.r + 12; row += 1) {
      state.classLookup.set(`CLASSES.xlsx|SCHED WORKING|${row}|${heading.c}`, `Air Assault ${number}`);
      state.classLookup.set(`CLASSES.xlsx|SCHED WORKING|${row}|${heading.c + 1}`, `Air Assault ${number}`);
    }
  }
}

function normalizeSourceEvent(event) {
  const sourceRefs = event.relatedSources || event.sources || [event.source].filter(Boolean);
  const classKey = inferClassKey(event);
  const normalized = {
    ...event,
    id: event.id || crypto.randomUUID(),
    title: cleanTitle(event.title),
    category: event.category || "Operations",
    start: normalizeTime(event.start),
    end: normalizeTime(event.end),
    location: cleanTitle(event.location),
    notes: cleanTitle(event.notes),
    people: [],
    classKey,
    sourceKind: event.sourceKind || "schedule",
    relatedSources: sourceRefs,
  };
  if (normalized.category === "Medical" && /RUCK/i.test(`${normalized.title} ${normalized.location}`)) {
    normalized.notes = normalized.notes.replace(/X\d+\s*FLA/i, "X3 FLA") || "X3 FLA W/ CREW";
  }
  return normalized;
}

function eventText(event) {
  return [
    event.title,
    event.category,
    event.group,
    event.location,
    event.notes,
    event.classKey,
    event.source?.file,
    event.source?.sheet,
  ]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();
}

function eventSort(a, b) {
  return (a.start || "99:99").localeCompare(b.start || "99:99") || trackSort(a.classKey).localeCompare(trackSort(b.classKey)) || a.title.localeCompare(b.title);
}

function classNumberFromKey(classKey) {
  return Number(String(classKey || "").match(/\d+/)?.[0] || 0);
}

function trackSort(value) {
  return value || "Shared";
}

function dayLabelFromTitle(title) {
  const match = String(title || "").match(/\bDAY\s+(-?\d+)\b/i);
  return match ? `Day ${match[1]}` : "";
}

function classDayLabelForTrack(iso, track) {
  const storedLabel = state.classDayLabels.get(`${iso}|${track}`);
  if (storedLabel) return storedLabel;
  const dayEvent = state.events.find((event) => event.date === iso && event.classKey === track && dayLabelFromTitle(event.title));
  return dayLabelFromTitle(dayEvent?.title);
}

function buildClassDayLabels(events) {
  state.classDayLabels = new Map();
  events.forEach((event) => {
    const label = dayLabelFromTitle(event.title);
    if (!label || !event.date || !event.classKey) return;
    state.classDayLabels.set(`${event.date}|${event.classKey}`, label);
  });
}

function mergeKey(event) {
  return [
    event.date,
    event.start || "",
    event.end || "",
    event.category,
    event.classKey || "Shared",
    cleanTitle(event.title).toUpperCase().replace(/\b(HARBER|HARRIS|MICEK|OLSON|FRANK|BERLIN|JOHNSON|MAYFIELD|BED)\b/g, "").replace(/\s+/g, " "),
  ].join("|");
}

function mergeEvents(events) {
  const byKey = new Map();
  for (const event of events) {
    const key = mergeKey(event);
    if (!byKey.has(key)) {
      byKey.set(key, { ...event, relatedSources: [...(event.relatedSources || [])] });
      continue;
    }
    const existing = byKey.get(key);
    existing.notes = Array.from(new Set([existing.notes, event.notes].filter(Boolean))).join(" / ");
    existing.location = existing.location || event.location;
    existing.group = existing.group || event.group;
    existing.relatedSources.push(...(event.relatedSources || []));
  }
  return Array.from(byKey.values()).sort(eventSort).map((event, index) => ({ ...event, id: event.id || `op-${index}` }));
}

function medicalCoverageLocation(event) {
  const titleMatch = String(event.title || "").match(/Medical coverage\s*-\s*(.+)$/i);
  return cleanTitle(event.location || titleMatch?.[1] || "Medical coverage");
}

function medicalCoverageTitle(event, locations) {
  const prefix = String(event.title || "").match(/^\s*(\d{3,4}-UTC)\b/i)?.[1] || (event.start ? `${event.start.replace(":", "")}-UTC` : "");
  return `${prefix ? `${prefix} ` : ""}Medical coverage - ${locations.join(" + ")}`;
}

function dedupeSources(sources) {
  const byKey = new Map();
  sources.filter(Boolean).forEach((source) => {
    const key = `${source.file || ""}|${source.sheet || ""}|${source.row || ""}|${source.col || ""}`;
    byKey.set(key, source);
  });
  return Array.from(byKey.values());
}

function southDockPeKey(event) {
  const text = `${event.title || ""} ${event.location || ""}`.toUpperCase();
  if (/\bCLEAN\s*PE\b/.test(text)) return "clean-pe";
  if (/\b(?:AA|AIR\s*ASSU?ALT)\s*PT\b/.test(text)) return "aa-pt";
  const peMatch = text.match(/\bPE\s*([123])\b/);
  return peMatch ? `pe-${peMatch[1]}` : "";
}

function normalizeSouthDockPeEvents(events) {
  const byWindow = new Map();
  events.forEach((event) => {
    const peKey = southDockPeKey(event);
    if (!peKey) return;
    event.location = "South Dock";
    const key = [event.date, event.classKey || "", event.start || "", event.end || "", peKey].join("|");
    if (!byWindow.has(key)) byWindow.set(key, []);
    byWindow.get(key).push(event);
  });

  const removeIds = new Set();
  byWindow.forEach((items) => {
    const keeper = items.find((event) => event.sourceKind === "lrtc") || items[0];
    const duplicateSources = items.filter((event) => event !== keeper).flatMap((event) => event.relatedSources || []);
    keeper.relatedSources = dedupeSources([...(keeper.relatedSources || []), ...duplicateSources]);
    items.forEach((event) => {
      if (event !== keeper && event.sourceKind === "class-calendar") removeIds.add(event.id);
    });
  });

  return events.filter((event) => !removeIds.has(event.id));
}

function minutesFromTime(value) {
  const normalized = normalizeTime(value);
  if (!normalized) return null;
  const [hour, minute] = normalized.split(":").map(Number);
  return hour * 60 + minute;
}

function canonicalActivityText(text) {
  return cleanTitle(text)
    .toUpperCase()
    .replace(/\b[0-2]?\d[0-5]\d\s*-\s*(?:[0-2]?\d[0-5]\d|UTC)\b/g, " ")
    .replace(/\b[0-2]?\d[0-5]\d\b/g, " ")
    .replace(/\bP\s*([123])\b/g, "PHASE $1")
    .replace(/\bCLASSROM\b/g, "CLASSROOM")
    .replace(/\bRAPELLING\b/g, "RAPPELLING")
    .replace(/\bAIR\s+ASSUALT\b/g, "AIR ASSAULT")
    .replace(/\bGRAD\b/g, "GRADUATION")
    .replace(/[^\w\s/]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function activityEndMinutes(event, start) {
  const explicitEnd = minutesFromTime(event.end);
  if (explicitEnd !== null) return explicitEnd;
  return /\bUTC\b/i.test(event.title || "") ? start + 240 : start + 45;
}

function activityWindowsOverlap(a, b) {
  const aStart = minutesFromTime(a.start);
  const bStart = minutesFromTime(b.start);
  if (aStart === null || bStart === null) return false;
  const aEnd = activityEndMinutes(a, aStart);
  const bEnd = activityEndMinutes(b, bStart);
  return Math.max(aStart, bStart) < Math.min(aEnd, bEnd);
}

function isClassEventCoveredByLrtc(classEvent, lrtcEvent) {
  if (!["Training", "Instruction"].includes(classEvent.category) || !["Training", "Instruction"].includes(lrtcEvent.category)) return false;
  if (classEvent.date !== lrtcEvent.date || (classEvent.classKey || "") !== (lrtcEvent.classKey || "")) return false;
  const classTitle = canonicalActivityText(classEvent.title);
  const lrtcTitle = canonicalActivityText(lrtcEvent.title);
  if (!classTitle || !lrtcTitle) return false;
  const sameStart = classEvent.start && lrtcEvent.start && classEvent.start === lrtcEvent.start;
  const sameWindow = sameStart && (classEvent.end || "") === (lrtcEvent.end || "");
  const overlappingWindow = activityWindowsOverlap(classEvent, lrtcEvent);
  const similarTitle = lrtcTitle.includes(classTitle) || classTitle.includes(lrtcTitle) || (classTitle.includes("PHASE 1 TEST") && lrtcTitle.includes("PHASE 1 TEST"));
  const ambiguousClassTitle = /^(\d+\s*HR|PHASE\s+\d+\s+TEST|COLD\s+LOAD|FORMATION|AIR|LZ\/PZ|SLING)/.test(classTitle);
  const samePlace = classEvent.location && lrtcEvent.location && canonicalActivityText(classEvent.location) === canonicalActivityText(lrtcEvent.location);
  return (overlappingWindow && similarTitle) || (sameWindow && (ambiguousClassTitle || samePlace || lrtcTitle.length > classTitle.length + 6));
}

function removeClassEventsCoveredByLrtc(events) {
  const lrtcEvents = events.filter((event) => event.sourceKind === "lrtc" && event.classKey);
  if (!lrtcEvents.length) return events;
  return events.filter((event) => {
    if (event.sourceKind !== "class-calendar") return true;
    return !lrtcEvents.some((lrtcEvent) => isClassEventCoveredByLrtc(event, lrtcEvent));
  });
}

function removeUntimedEvents(events) {
  return events.filter((event) => event.start);
}

function addDayMinusTwoMarneMedicalSupportEvents(events) {
  const existing = new Set(
    events
      .filter((event) => event.category === "Medical" && /MARNE/i.test(`${event.title || ""} ${event.location || ""} ${event.notes || ""}`))
      .map((event) => `${event.date}|${event.classKey}|${event.start || ""}|${event.end || ""}`)
  );
  const dayMinusTwoEvents = events.filter((event) => event.date && event.classKey && dayLabelFromTitle(event.title) === "Day -2");
  dayMinusTwoEvents.forEach((dayEvent) => {
    const key = `${dayEvent.date}|${dayEvent.classKey}|08:00|11:00`;
    if (existing.has(key)) return;
    existing.add(key);
    events.push({
      id: `generated-day-minus-two-marne-fla-${dayEvent.date}-${dayEvent.classKey.replace(/\s+/g, "-").toLowerCase()}`,
      date: dayEvent.date,
      start: "08:00",
      end: "11:00",
      title: "0800-1100 Medical coverage - MARNE",
      category: "Medical",
      group: "Medical Coverage",
      classKey: dayEvent.classKey,
      sourceKind: "generated-medical",
      location: "MARNE",
      notes: "X1 FLA support",
      relatedSources: [...(dayEvent.relatedSources || [])],
    });
  });
  return events;
}

function isCanonicalLrtcEvent(event) {
  if (event.sourceKind !== "lrtc") return true;
  const sources = event.relatedSources || [event.source].filter(Boolean);
  return sources.some((source) => source?.file === CANONICAL_LRTC_FILE);
}

function coalesceSharedMedicalEvents(events) {
  const output = [];
  const byCoverageWindow = new Map();
  for (const event of events) {
    if (event.category !== "Medical" || event.classKey || !event.start) {
      output.push(event);
      continue;
    }

    const key = `${event.date}|${event.start}|${event.end || ""}`;
    const location = medicalCoverageLocation(event);
    if (!byCoverageWindow.has(key)) {
      const copy = {
        ...event,
        _medicalItems: [event],
        _medicalLocations: location ? [location] : [],
      };
      byCoverageWindow.set(key, copy);
      output.push(copy);
      continue;
    }

    const existing = byCoverageWindow.get(key);
    existing._medicalItems.push(event);
    if (location && !existing._medicalLocations.includes(location)) existing._medicalLocations.push(location);
    existing.relatedSources = dedupeSources([...(existing.relatedSources || []), ...(event.relatedSources || [])]);
  }

  return output
    .map((event) => {
      if (!event._medicalItems) return event;
      const { _medicalItems, _medicalLocations, ...cleanEvent } = event;
      if (_medicalItems.length === 1) return cleanEvent;
      const locations = _medicalLocations.length ? _medicalLocations : ["Medical coverage"];
      cleanEvent.id = `medical-${cleanEvent.date}-${cleanEvent.start}-${locations.join("-")}`;
      cleanEvent.title = medicalCoverageTitle(cleanEvent, locations);
      cleanEvent.location = locations.join(" + ");
      cleanEvent.notes = _medicalItems
        .map((item) => {
          const itemLocation = medicalCoverageLocation(item);
          return item.notes ? `${itemLocation}: ${item.notes}` : itemLocation;
        })
        .join(" / ");
      cleanEvent.relatedSources = dedupeSources(_medicalItems.flatMap((item) => item.relatedSources || []));
      return cleanEvent;
    })
    .sort(eventSort);
}

function mealLocationScore(event) {
  if (!event.location) return 0;
  if (/^(Cadet Mess|MRE|A-Frame|TBD meal type)$/i.test(event.location)) return 1;
  return 2;
}

function bestMealEvent(events) {
  return [...events].sort((a, b) => {
    const score = mealLocationScore(b) - mealLocationScore(a);
    if (score) return score;
    return (a.start || "99:99").localeCompare(b.start || "99:99");
  })[0];
}

function coalesceMealEvents(events) {
  const output = [];
  const byMeal = new Map();
  for (const event of events) {
    const meal = event.category === "Meals" ? mealNameFromText(event.title) : "";
    if (!meal || !event.classKey) {
      output.push(event);
      continue;
    }

    const key = `${event.date}|${event.classKey}|${meal}`;
    if (!byMeal.has(key)) {
      const copy = { ...event, _mealItems: [event] };
      byMeal.set(key, copy);
      output.push(copy);
      continue;
    }

    const existing = byMeal.get(key);
    existing._mealItems.push(event);
    existing.relatedSources = dedupeSources([...(existing.relatedSources || []), ...(event.relatedSources || [])]);
  }

  return output
    .map((event) => {
      if (!event._mealItems) return event;
      const { _mealItems, ...cleanEvent } = event;
      if (_mealItems.length === 1) return cleanEvent;
      const best = bestMealEvent(_mealItems);
      const times = Array.from(new Set(_mealItems.map((item) => item.start).filter(Boolean))).sort();
      return {
        ...cleanEvent,
        ...best,
        id: `meal-${cleanEvent.date}-${cleanEvent.classKey}-${mealNameFromText(cleanEvent.title)}`,
        relatedSources: dedupeSources(_mealItems.flatMap((item) => item.relatedSources || [])),
        notes:
          times.length > 1
            ? "Meal type from mess matrix; duplicate schedule rows merged."
            : best.notes,
      };
    })
    .sort(eventSort);
}

function buildOperationalEvents(rawEvents) {
  const normalized = rawEvents.map(normalizeSourceEvent).filter(isCanonicalLrtcEvent);
  buildClassDayLabels(normalized);
  const messMap = new Map();
  const messEvents = normalized.filter((event) => event.sourceKind === "mess");
  for (const event of messEvents) {
    const meal = mealNameFromText(event.title);
    const classKey = inferClassKey(event);
    const status = mealStatusFromText(event.title);
    if (!meal || !classKey || !status) continue;
    messMap.set(`${event.date}|${classKey}|${meal}`, { status, event });
  }

  const operational = [];
  const scheduleMealByKey = new Map();
  for (const event of normalized) {
    if (event.sourceKind === "mess") continue;
    if (event.category === "Meals") {
      const meal = mealNameFromText(event.title);
      if (!meal) continue;
      const classKey = event.classKey || "";
      const key = `${event.date}|${classKey || "Shared"}|${meal}`;
      if (!scheduleMealByKey.has(key)) scheduleMealByKey.set(key, []);
      scheduleMealByKey.get(key).push(event);
      continue;
    }
    operational.push(event);
  }

  const usedMessKeys = new Set();
  for (const schedules of scheduleMealByKey.values()) {
    const schedule = mergeScheduleMealGroup(schedules);
    const meal = mealNameFromText(schedule.title);
    const candidateKeys = schedule.classKey
      ? [`${schedule.date}|${schedule.classKey}|${meal}`]
      : Array.from(messMap.keys()).filter((key) => key.startsWith(`${schedule.date}|`) && key.endsWith(`|${meal}`));
    const targets = candidateKeys.length ? candidateKeys : [`${schedule.date}|${schedule.classKey || "Shared"}|${meal}`];
    for (const key of targets) {
      const mess = messMap.get(key);
      const classKey = key.split("|")[1] === "Shared" ? schedule.classKey : key.split("|")[1];
      const status = mess?.status || mealStatusFromText(schedule.title) || "";
      const start = mealStartFor(status, meal, schedule);
      usedMessKeys.add(key);
      operational.push({
        ...schedule,
        id: `meal-${key}-${start}`,
        title: `${meal}: ${mealStatusLabel(status)}`,
        category: "Meals",
        start,
        end: isCadetMessStatus(status) ? "" : schedule.end,
        classKey,
        sourceKind: "meal-merged",
        location: mealLocationFor(status, schedule),
        notes: mealNotesFor(status, meal, schedule),
        relatedSources: dedupeSources([...(schedule.relatedSources || []), mess?.event?.source].filter(Boolean)),
      });
    }
  }

  for (const [key, mess] of messMap.entries()) {
    if (usedMessKeys.has(key)) continue;
    const [date, classKey, meal] = key.split("|");
    const start = mealStartFor(mess.status, meal, null);
    operational.push({
      ...mess.event,
      id: `meal-mess-only-${key}`,
      date,
      start,
      end: "",
      title: `${meal}: ${mealStatusLabel(mess.status)}`,
      category: "Meals",
      classKey,
      sourceKind: "meal-merged",
      location: mealLocationFor(mess.status, null),
      notes: mealNotesFor(mess.status, meal, null),
      relatedSources: [mess.event.source].filter(Boolean),
    });
  }

  const lrtcPreferred = removeClassEventsCoveredByLrtc(operational);
  const southDockNormalized = normalizeSouthDockPeEvents(lrtcPreferred);
  const withDayMinusTwoSupport = addDayMinusTwoMarneMedicalSupportEvents(southDockNormalized);
  const timedOperational = removeUntimedEvents(withDayMinusTwoSupport);
  return coalesceSharedMedicalEvents(coalesceMealEvents(mergeEvents(timedOperational)));
}

function filteredEvents(events = state.events) {
  const query = state.search.trim().toLowerCase();
  return events.filter((event) => {
    if (state.sourceFilter !== "all" && !event.relatedSources.some((source) => source?.file === state.sourceFilter)) return false;
    return !query || eventText(event).includes(query);
  });
}

function eventsForDate(iso, events = filteredEvents()) {
  return events.filter((event) => event.date === iso);
}

function manualEventFromOverride(id, override = {}) {
  return {
    id,
    date: override.date || state.selectedDate,
    start: override.start || "",
    end: override.end || "",
    title: override.title || "New timeline event",
    category: override.category || "Operations",
    group: override.category || "Operations",
    location: override.location || "",
    people: [],
    notes: override.notes || "",
    classKey: override.classKey || "",
    sourceKind: "manual-event",
    relatedSources: [],
    source: null,
    isEdited: true,
    isManual: true,
  };
}

function createManualEventId() {
  const randomId = globalThis.crypto?.randomUUID ? globalThis.crypto.randomUUID() : `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  return `manual-event-${randomId}`;
}

function applyEventOverrides(events) {
  const sourceIds = new Set(events.map((event) => event.id));
  const editedSourceEvents = events.map((event) => {
    const override = state.eventOverrides[event.id];
    if (!override) return event;
    return {
      ...event,
      ...override,
      id: event.id,
      relatedSources: event.relatedSources,
      source: event.source,
      sourceKind: event.sourceKind,
      isEdited: true,
    };
  });
  const manualEvents = Object.entries(state.eventOverrides)
    .filter(([id, override]) => override?.sourceKind === "manual-event" && !sourceIds.has(id))
    .map(([id, override]) => manualEventFromOverride(id, override));
  return [...editedSourceEvents, ...manualEvents].filter((event) => event.start);
}

function rebuildEvents() {
  state.events = applyEventOverrides(buildOperationalEvents(state.rawEvents));
}

function eventById(id) {
  return state.events.find((event) => event.id === id);
}

function countBy(items, getter) {
  return items.reduce((acc, item) => {
    const key = getter(item);
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function loadJson(key, fallback) {
  try {
    const value = JSON.parse(localStorage.getItem(key));
    return value ?? fallback;
  } catch {
    return fallback;
  }
}

function normalizeDeletedSource(value = {}) {
  return {
    taskings: Array.from(new Set(Array.isArray(value.taskings) ? value.taskings : [])),
    supportItems: Array.from(new Set(Array.isArray(value.supportItems) ? value.supportItems : [])),
  };
}

function saveDeletedSource() {
  state.deletedSource = normalizeDeletedSource(state.deletedSource);
  localStorage.setItem(DELETED_SOURCE_KEY, JSON.stringify(state.deletedSource));
  scheduleCloudSave("deleted-source");
}

function setSyncStatus(label, status = "local") {
  const el = $("#syncStatus");
  const card = el?.closest(".sync-card");
  if (!el || !card) return;
  el.textContent = label;
  card.classList.toggle("is-online", status === "online");
  card.classList.toggle("is-saving", status === "saving");
  card.classList.toggle("is-error", status === "error");
}

function cloudConfig() {
  return window.AASLT_FIREBASE_CONFIG || {};
}

function isCloudConfigured(config = cloudConfig()) {
  return Boolean(config.enabled && config.firebase?.apiKey && config.firebase?.projectId && config.firebase?.appId);
}

function persistLocalStateSnapshot() {
  localStorage.setItem(TASKS_KEY, JSON.stringify(state.tasks));
  localStorage.setItem(S4_KEY, JSON.stringify(state.s4Items));
  localStorage.setItem(EVENT_OVERRIDES_KEY, JSON.stringify(state.eventOverrides));
  state.deletedSource = normalizeDeletedSource(state.deletedSource);
  localStorage.setItem(DELETED_SOURCE_KEY, JSON.stringify(state.deletedSource));
}

function cloudStatePayload() {
  return {
    schemaVersion: CLOUD_SYNC_SCHEMA,
    sourceGeneratedAt: state.data?.generatedAt || "",
    clientUpdatedAt: new Date().toISOString(),
    tasks: consolidateTasks(state.tasks),
    s4Items: state.s4Items,
    eventOverrides: state.eventOverrides,
    deletedSource: normalizeDeletedSource(state.deletedSource),
    receiptRegister: state.receipts.map(({ blob, ...meta }) => meta),
  };
}

function scheduleCloudSave(reason = "state") {
  if (!state.cloud.ready || state.cloud.applying || !state.cloud.docRef || !state.cloud.setDoc) return;
  window.clearTimeout(state.cloud.saveTimer);
  state.cloud.saveTimer = window.setTimeout(() => writeCloudState(reason), 450);
}

async function writeCloudState(reason = "state") {
  if (!state.cloud.ready || state.cloud.applying || !state.cloud.docRef || !state.cloud.setDoc) return;
  try {
    setSyncStatus("Saving", "saving");
    const payload = cloudStatePayload();
    await state.cloud.setDoc(
      state.cloud.docRef,
      {
        ...payload,
        writeReason: reason,
        serverUpdatedAt: state.cloud.serverTimestamp ? state.cloud.serverTimestamp() : payload.clientUpdatedAt,
      },
      { merge: false }
    );
    setSyncStatus("Synced", "online");
  } catch (error) {
    state.cloud.lastError = error.message;
    setSyncStatus("Sync error", "error");
    console.error("Cloud sync write failed", error);
  }
}

function applyCloudState(payload = {}) {
  state.cloud.applying = true;
  try {
    state.deletedSource = normalizeDeletedSource(payload.deletedSource || {});
    state.eventOverrides = payload.eventOverrides || {};
    state.tasks = mergeSourceTaskings(Array.isArray(payload.tasks) ? payload.tasks : [], state.data.taskings || []);
    state.s4Items = mergeSourceSupportItems(migrateS4(Array.isArray(payload.s4Items) ? payload.s4Items : []), state.data.supportItems || []);
    if (Array.isArray(payload.receiptRegister)) {
      const localReceipts = new Map(state.receipts.map((receipt) => [receipt.id, receipt]));
      state.receipts = payload.receiptRegister.map((receipt) => ({ ...receipt, blob: localReceipts.get(receipt.id)?.blob }));
    }
    persistLocalStateSnapshot();
    rebuildEvents();
    renderAll();
    setSyncStatus("Synced", "online");
  } finally {
    state.cloud.applying = false;
  }
}

async function initCloudSync() {
  const config = cloudConfig();
  if (!isCloudConfigured(config)) {
    setSyncStatus("Local only");
    return;
  }

  try {
    setSyncStatus("Connecting", "saving");
    const version = config.firebaseVersion || "10.12.5";
    const [{ initializeApp }, firestore, storageModule] = await Promise.all([
      import(`https://www.gstatic.com/firebasejs/${version}/firebase-app.js`),
      import(`https://www.gstatic.com/firebasejs/${version}/firebase-firestore.js`),
      import(`https://www.gstatic.com/firebasejs/${version}/firebase-storage.js`),
    ]);
    const app = initializeApp(config.firebase);
    const db = firestore.getFirestore(app);
    const storage = config.firebase.storageBucket ? storageModule.getStorage(app) : null;
    const path = config.firestorePath || `dashboards/${config.dashboardId || "aaslt-control"}/state/main`;
    const pathParts = path.split("/").filter(Boolean);
    if (pathParts.length % 2 !== 0) throw new Error(`Firestore path must point to a document: ${path}`);

    state.cloud.enabled = true;
    state.cloud.ready = true;
    state.cloud.docRef = firestore.doc(db, ...pathParts);
    state.cloud.storage = storage;
    state.cloud.setDoc = firestore.setDoc;
    state.cloud.serverTimestamp = firestore.serverTimestamp;
    state.cloud.storageRef = storageModule.ref;
    state.cloud.uploadBytes = storageModule.uploadBytes;
    state.cloud.getDownloadURL = storageModule.getDownloadURL;
    state.cloud.deleteObject = storageModule.deleteObject;

    firestore.onSnapshot(
      state.cloud.docRef,
      (snapshot) => {
        if (!snapshot.exists()) {
          writeCloudState("initialize");
          return;
        }
        applyCloudState(snapshot.data());
      },
      (error) => {
        state.cloud.lastError = error.message;
        setSyncStatus("Sync error", "error");
        console.error("Cloud sync listener failed", error);
      }
    );
  } catch (error) {
    state.cloud.lastError = error.message;
    setSyncStatus("Sync error", "error");
    console.error("Cloud sync failed to start", error);
  }
}

function saveTasks() {
  state.tasks = consolidateTasks(state.tasks);
  localStorage.setItem(TASKS_KEY, JSON.stringify(state.tasks));
  scheduleCloudSave("tasks");
}

function saveEventOverrides() {
  localStorage.setItem(EVENT_OVERRIDES_KEY, JSON.stringify(state.eventOverrides));
  scheduleCloudSave("event-overrides");
}

function saveS4() {
  localStorage.setItem(S4_KEY, JSON.stringify(state.s4Items));
  scheduleCloudSave("s4");
}

function saveSettings() {
  localStorage.setItem(SETTINGS_KEY, JSON.stringify({ selectedDate: state.selectedDate, sourceFilter: state.sourceFilter, dayView: state.dayView, assignmentView: state.assignmentView }));
}

function peopleRequiredFromTitle(title) {
  const matches = Array.from(String(title || "").matchAll(/\b(\d+)\s*x\s*Cadre\b/gi));
  return matches.length ? matches.reduce((sum, match) => sum + Number(match[1]), 0) : 1;
}

function canonicalTaskTitle(title) {
  return cleanTitle(title)
    .replace(/^Cadre coverage\s*-\s*(?:\d+\s*x\s*Cadre\s*)+$/i, "Cadre coverage")
    .replace(/\bRecieve\b/gi, "Receive")
    .replace(/\bClassrom\b/gi, "Classroom");
}

function taskOwners(task) {
  const owners = Array.isArray(task.owners) ? task.owners : task.owner && task.owner !== "Unassigned" ? [task.owner] : [];
  return Array.from(new Set(owners.filter((owner) => ASSIGNABLE_PEOPLE.includes(owner))));
}

function ownersLabel(task) {
  const owners = taskOwners(task);
  return owners.length ? owners.join(", ") : "Unassigned";
}

function taskRequiredPeople(task) {
  return Math.max(1, Number(task.requiredPeople || 0), peopleRequiredFromTitle(task.title), taskOwners(task).length);
}

function taskTaskees(task) {
  return Math.max(0, Number(task.taskees || task.taskeeCount || 0));
}

function taskStaffingText(task) {
  const cadreText = `${taskOwners(task).length} of ${taskRequiredPeople(task)} cadre`;
  const taskees = taskTaskees(task);
  return `${cadreText} / ${taskees} taskee${taskees === 1 ? "" : "s"}`;
}

function compactTaskNotes(notes, limit = 6) {
  const parts = String(notes || "")
    .split(" / ")
    .map((part) => part.trim())
    .filter(Boolean);
  if (!parts.length) return "";
  const unique = [];
  const seen = new Set();
  parts.forEach((part) => {
    const key = part.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    unique.push(part);
  });
  if (unique.length <= limit) return unique.join(" / ");
  return `${unique.slice(0, limit).join(" / ")} / +${unique.length - limit} more`;
}

function selectedOwnersFromForm() {
  return Array.from($$("#taskOwners input[type='checkbox']:checked")).map((input) => input.value);
}

function setSelectedOwners(container, owners = []) {
  const selected = new Set(owners);
  $$("input[type='checkbox']", container).forEach((input) => {
    input.checked = selected.has(input.value);
  });
}

function renderPersonOption(person, checked, onChange) {
  const input = createElement("input", { attrs: { type: "checkbox", value: person } });
  input.checked = checked;
  input.addEventListener("change", onChange);
  const label = createElement("label", { className: "person-option" });
  label.addEventListener("click", (event) => {
    if (event.target === input) return;
    event.preventDefault();
    input.checked = !input.checked;
    onChange(event);
  });
  label.append(input, createElement("span", { text: person }));
  return label;
}

function selectedOwnersFromPicker(container) {
  return Array.from($$("input[type='checkbox']:checked", container)).map((input) => input.value);
}

function statusWithOwners(status, owners) {
  if (status === "done") return "done";
  return owners.length ? "claimed" : "open";
}

function normalizeTaskRecord(task) {
  const owners = taskOwners(task);
  const title = canonicalTaskTitle(task.title);
  return {
    ...task,
    title,
    owners,
    owner: owners[0] || "Unassigned",
    requiredPeople: taskRequiredPeople({ ...task, title, owners }),
    taskees: taskTaskees(task),
    status: statusWithOwners(task.status || "open", owners),
    notes: compactTaskNotes(task.notes),
  };
}

function taskConsolidationKey(task) {
  const titleKey = canonicalTaskTitle(task.title).toUpperCase();
  return [task.date, task.start || "", task.end || "", task.track || "Shared", titleKey].join("|");
}

function taskSourceSeedIds(task) {
  return Array.from(new Set([...(task.sourceSeedIds || []), task.sourceSeedId].filter(Boolean)));
}

function mergeTaskStatuses(tasks) {
  if (tasks.every((task) => task.status === "done")) return "done";
  if (tasks.some((task) => task.status === "claimed" || taskOwners(task).length)) return "claimed";
  return "open";
}

function consolidateTasks(tasks) {
  const byKey = new Map();
  tasks.map(normalizeTaskRecord).forEach((task) => {
    const key = taskConsolidationKey(task);
    if (!byKey.has(key)) {
      byKey.set(key, task);
      return;
    }

    const existing = byKey.get(key);
    const owners = Array.from(new Set([...taskOwners(existing), ...taskOwners(task)]));
    const notes = compactTaskNotes([existing.notes, task.notes].filter(Boolean).join(" / "));
    const sourceSeedIds = Array.from(new Set([...(existing.sourceSeedIds || [existing.sourceSeedId].filter(Boolean)), ...(task.sourceSeedIds || [task.sourceSeedId].filter(Boolean))]));
    byKey.set(key, {
      ...existing,
      owners,
      owner: owners[0] || "Unassigned",
      requiredPeople: Math.max(taskRequiredPeople(existing), taskRequiredPeople(task)),
      taskees: Math.max(taskTaskees(existing), taskTaskees(task)),
      status: mergeTaskStatuses([existing, task]),
      notes,
      sourceSeedIds,
      sourceKind: existing.sourceKind || task.sourceKind,
      updatedAt: [existing.updatedAt, task.updatedAt].filter(Boolean).sort().pop() || new Date().toISOString(),
    });
  });
  return Array.from(byKey.values());
}

function taskFromSourceTasking(tasking) {
  const requiredPeople = peopleRequiredFromTitle(tasking.title);
  return {
    id: `seed-${tasking.id}`,
    sourceSeedId: tasking.id,
    sourceKind: "lrtc-tasking",
    title: canonicalTaskTitle(tasking.title),
    date: tasking.date,
    start: tasking.start || "",
    end: tasking.end || "",
    track: tasking.track || "Shared",
    owners: [],
    owner: "Unassigned",
    requiredPeople,
    taskees: 0,
    status: "open",
    notes: [tasking.location, tasking.role, sourceSummary(tasking)].filter(Boolean).join(" / "),
    updatedAt: state.data?.generatedAt || new Date().toISOString(),
  };
}

function mergeSourceTaskings(existingTasks, sourceTaskings = []) {
  const deleted = new Set(state.deletedSource.taskings || []);
  const seeds = sourceTaskings.map(taskFromSourceTasking).filter((task) => !deleted.has(task.sourceSeedId));
  const seedIds = new Set(seeds.map((task) => task.sourceSeedId));
  const existingBySeed = new Map();
  const manualTasks = [];

  existingTasks.forEach((task) => {
    if (taskSourceSeedIds(task).some((seedId) => deleted.has(seedId))) return;
    if (task.sourceKind === "lrtc-tasking" && task.sourceSeedId) {
      if (seedIds.has(task.sourceSeedId)) existingBySeed.set(task.sourceSeedId, task);
      return;
    }
    manualTasks.push(task);
  });

  const mergedSeeds = seeds.map((seed) => {
    const existing = existingBySeed.get(seed.sourceSeedId);
    if (!existing) return seed;
    const owners = taskOwners(existing);
    return {
      ...seed,
      id: existing.id || seed.id,
      owners,
      owner: owners[0] || "Unassigned",
      requiredPeople: Math.max(seed.requiredPeople || 1, existing.requiredPeople || 1, owners.length),
      taskees: taskTaskees(existing),
      status: existing.status || seed.status,
      notes: existing.notes || seed.notes,
      updatedAt: existing.updatedAt || seed.updatedAt,
    };
  });

  return consolidateTasks([...manualTasks, ...mergedSeeds]);
}

function normalizeS4Lane(type) {
  if (type === "FLA") return "FLA Request";
  if (type === "Transportation") return "LMTV TMR";
  return S4_LANES.includes(type) ? type : "Draw";
}

function inferS4Location(text = "") {
  const upper = String(text).toUpperCase();
  if (upper.includes("SOUTH DOCK")) return "South Dock";
  if (upper.includes("DAVIS")) return "Davis Shelf";
  if (upper.includes("MARNE")) return "MARNE";
  return "";
}

function inferWaterAction(text = "") {
  const upper = String(text).toUpperCase();
  if (upper.includes("DRAIN")) return "Drain";
  if (upper.includes("PICKUP") || upper.includes("PICK UP")) return "Pickup";
  if (upper.includes("PLACE")) return "Placement";
  return "Refill";
}

function inferFlaPlacement(text = "") {
  return String(text || "")
    .replace(/^\s*\d+\s*x?\s*/i, "")
    .replace(/\bFLA\b/gi, "")
    .replace(/\b(at|on|to|for)\b/gi, "")
    .replace(/^[\s:;-]+|[\s:;-]+$/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeS4Status(type, status = "") {
  const options = S4_STATUS_OPTIONS[type] || S4_STATUS_OPTIONS.Draw;
  if (options.includes(status)) return status;
  if (["Requested", "Dispatched", "Issued"].includes(status)) return "Sent";
  if (["Refilled", "Returned"].includes(status)) return "Complete";
  return options[0];
}

function normalizeS4Item(item = {}) {
  const type = normalizeS4Lane(item.type || item.lane);
  const title = item.item || item.title || "";
  const needBy = item.needBy || item.datetime || item.requestTime || item.drawTime || "";
  const normalized = {
    ...item,
    type,
    item: title || (type === "FLA Request" ? "FLA Request" : type),
    qty: Math.max(1, Number(item.qty || item.count || 1)),
    needBy,
    requestTime: item.requestTime || needBy,
    drawTime: item.drawTime || needBy,
    status: normalizeS4Status(type, item.status || "Planned"),
    owner: item.owner || "",
    notes: item.notes || "",
  };

  if (type === "Water Buffalo") {
    normalized.requestKind = item.requestKind || inferWaterAction(title);
    normalized.location = item.location || inferS4Location(title) || "Davis Shelf";
    normalized.item = `${normalized.requestKind} Water Buffalo`;
  }

  if (type === "LMTV TMR") {
    normalized.pax = Math.max(0, Number(item.pax || item.qty || 0));
    normalized.location = "MARNE";
    normalized.requestKind = "Round Trip";
    normalized.pickupTime = item.pickupTime || item.pickup || "";
    normalized.returnTime = item.returnTime || item.return || "";
    normalized.item = title || "LMTV Round Trip";
  }

  if (type === "FLA Request") {
    normalized.placement = item.placement || item.location || inferFlaPlacement(title);
    normalized.location = normalized.placement;
    normalized.item = `${normalized.qty}x FLA${normalized.placement ? ` - ${normalized.placement}` : ""}`;
  }

  if (type === "Draw") {
    normalized.drawTime = item.drawTime || needBy;
  }

  return normalized;
}

function s4FromSourceSupport(item) {
  return normalizeS4Item({
    id: `seed-${item.id}`,
    sourceSeedId: item.id,
    sourceKind: "lrtc-support",
    type: item.lane || item.type || "Draw",
    item: item.item,
    qty: Number(item.qty || 1),
    needBy: item.datetime || "",
    status: item.status || "Planned",
    owner: item.owner || "",
    notes: [item.track && item.track !== "Shared" ? item.track : "", sourceSummary(item)].filter(Boolean).join(" / "),
  });
}

function mergeSourceSupportItems(existingItems, sourceItems = []) {
  const deleted = new Set(state.deletedSource.supportItems || []);
  const seeds = sourceItems.map(s4FromSourceSupport).filter((item) => !deleted.has(item.sourceSeedId));
  const seedIds = new Set(seeds.map((item) => item.sourceSeedId));
  const existingBySeed = new Map();
  const manualItems = [];

  existingItems.forEach((item) => {
    if (item.sourceSeedId && deleted.has(item.sourceSeedId)) return;
    if (item.sourceKind === "lrtc-support" && item.sourceSeedId) {
      if (seedIds.has(item.sourceSeedId)) existingBySeed.set(item.sourceSeedId, item);
      return;
    }
    manualItems.push(item);
  });

  const mergedSeeds = seeds.map((seed) => {
    const existing = existingBySeed.get(seed.sourceSeedId);
    if (!existing) return seed;
    return normalizeS4Item({
      ...seed,
      id: existing.id || seed.id,
      status: existing.status || seed.status,
      owner: existing.owner || seed.owner,
      notes: existing.notes || seed.notes,
      requestKind: existing.requestKind || seed.requestKind,
      location: existing.location || seed.location,
      placement: existing.placement || seed.placement,
      pax: existing.pax ?? seed.pax,
      pickupTime: existing.pickupTime || seed.pickupTime,
      returnTime: existing.returnTime || seed.returnTime,
    });
  });

  return [...manualItems.map(normalizeS4Item), ...mergedSeeds];
}

async function init() {
  state.tasks = loadJson(TASKS_KEY, []);
  state.deletedSource = normalizeDeletedSource(loadJson(DELETED_SOURCE_KEY, {}));
  state.eventOverrides = loadJson(EVENT_OVERRIDES_KEY, {});
  state.s4Items = migrateS4(loadJson(S4_KEY, []));
  const response = await fetch(DATA_URL);
  state.data = normalizeLoadedData(await response.json());
  state.rawEvents = state.data.events || [];
  state.tasks = mergeSourceTaskings(state.tasks, state.data.taskings || []);
  state.s4Items = mergeSourceSupportItems(state.s4Items, state.data.supportItems || []);
  saveTasks();
  saveS4();
  buildClassLookup();
  rebuildEvents();
  initSettings();
  bindEvents();
  renderStaticControls();
  await loadReceipts();
  renderAll();
  loadWeather();
  await initCloudSync();
}

async function loadWeather() {
  state.weather = { status: "loading", days: [], updatedAt: "", error: "" };
  renderWeather();
  try {
    const response = await fetch(WEATHER_URL);
    if (!response.ok) throw new Error(`Weather request failed: ${response.status}`);
    const payload = await response.json();
    state.weather = {
      status: "ready",
      days: normalizeWeather(payload),
      updatedAt: payload.current?.time || "",
      error: "",
    };
  } catch (error) {
    state.weather = {
      status: "error",
      days: [],
      updatedAt: "",
      error: error.message || "Weather unavailable",
    };
  }
  renderWeather();
}

function initSettings() {
  const settings = loadJson(SETTINGS_KEY, {});
  const todayIso = toIso(new Date());
  const eventDates = new Set(state.events.map((event) => event.date));
  const [start] = state.data.summary.dateRange;
  state.selectedDate = eventDates.has(settings.selectedDate) ? settings.selectedDate : eventDates.has(todayIso) ? todayIso : start;
  state.viewMonth = parseDate(state.selectedDate);
  const sourceNames = new Set((state.data.sources || []).map((source) => source.name));
  state.sourceFilter = settings.sourceFilter && (settings.sourceFilter === "all" || sourceNames.has(settings.sourceFilter)) ? settings.sourceFilter : "all";
  state.dayView = settings.dayView || "tracks";
  state.assignmentView = settings.assignmentView || "list";
}

function bindEvents() {
  $("#globalSearch").addEventListener("input", (event) => {
    state.search = event.target.value;
    renderAll();
  });
  $("#todayButton").addEventListener("click", () => {
    const today = toIso(new Date());
    selectDate(state.events.some((event) => event.date === today) ? today : state.data.summary.dateRange[0]);
  });
  $("#prevMonthButton").addEventListener("click", () => {
    state.viewMonth.setMonth(state.viewMonth.getMonth() - 1);
    renderCalendar();
  });
  $("#nextMonthButton").addEventListener("click", () => {
    state.viewMonth.setMonth(state.viewMonth.getMonth() + 1);
    renderCalendar();
  });
  $("#prevDayButton").addEventListener("click", () => selectDate(addDays(state.selectedDate, -1)));
  $("#nextDayButton").addEventListener("click", () => selectDate(addDays(state.selectedDate, 1)));
  $("#addEventButton").addEventListener("click", openNewEventEditor);
  $("#sourceFilter").addEventListener("change", (event) => {
    state.sourceFilter = event.target.value;
    renderAll();
  });
  $$("#openNextPaneButton, #closeNextPaneButton, #drawerBackdrop").forEach((el) => el.addEventListener("click", () => toggleNextDrawer()));
  $$("#quickTaskButton").forEach((el) => el.addEventListener("click", openNewTaskForm));
  $$(".segment[data-day-view]").forEach((button) => {
    button.addEventListener("click", () => {
      state.dayView = button.dataset.dayView;
      renderDay();
      saveSettings();
    });
  });
  $$(".segment[data-assignment-view]").forEach((button) => {
    button.addEventListener("click", () => {
      state.assignmentView = button.dataset.assignmentView;
      renderAssignments();
      saveSettings();
    });
  });
  $$(".tab").forEach((button) => button.addEventListener("click", () => selectTab(button.dataset.tab)));
  $("#taskForm").addEventListener("submit", saveTaskFromForm);
  $("#closeTaskFormButton").addEventListener("click", closeTaskForm);
  $("#s4Type").addEventListener("change", configureS4Form);
  $("#s4Form").addEventListener("submit", addS4Item);
  $("#clearCompletedS4").addEventListener("click", () => {
    state.s4Items.forEach((item) => {
      if (normalizeS4Item(item).status === "Complete") rememberDeletedSupportSource(item);
    });
    state.s4Items = state.s4Items.filter((item) => normalizeS4Item(item).status !== "Complete");
    saveS4();
    renderS4();
  });
  $("#receiptInput").addEventListener("change", (event) => addReceipts(Array.from(event.target.files)));
  $("#dropZone").addEventListener("dragover", (event) => {
    event.preventDefault();
    $("#dropZone").classList.add("is-dragging");
  });
  $("#dropZone").addEventListener("dragleave", () => $("#dropZone").classList.remove("is-dragging"));
  $("#dropZone").addEventListener("drop", (event) => {
    event.preventDefault();
    $("#dropZone").classList.remove("is-dragging");
    addReceipts(Array.from(event.dataTransfer.files));
  });
  $("#sourceSelect").addEventListener("change", (event) => {
    state.sourceIndex = Number(event.target.value);
    state.sheetIndex = 0;
    renderSourceControls();
    renderSourceTable();
  });
  $("#sheetSelect").addEventListener("change", (event) => {
    state.sheetIndex = Number(event.target.value);
    renderSourceTable();
  });
  $("#sourceSearch").addEventListener("input", (event) => {
    state.sourceSearch = event.target.value;
    renderSourceTable();
  });
  $("#closeDialog").addEventListener("click", () => $("#eventDialog").close());
  $("#eventEditForm").addEventListener("submit", saveEventEdit);
  $("#closeEditDialog").addEventListener("click", () => $("#eventEditDialog").close());
  $("#cancelEventEdit").addEventListener("click", () => $("#eventEditDialog").close());
  $("#resetEventEdit").addEventListener("click", resetEditedEvent);
  $("#printDayButton").addEventListener("click", printSelectedDay);
  $("#exportStateButton").addEventListener("click", exportLocalState);
}

function renderStaticControls() {
  const sourceFilter = $("#sourceFilter");
  for (const source of state.data.sources) sourceFilter.append(createElement("option", { text: source.name, attrs: { value: source.name } }));
  sourceFilter.value = state.sourceFilter;
  const ownerSelect = $("#taskOwners");
  ownerSelect.replaceChildren();
  ASSIGNABLE_PEOPLE.forEach((person) => ownerSelect.append(renderPersonOption(person, false, () => {})));
  $("#taskDate").value = state.selectedDate;
  configureS4Form();
  renderSourceControls();
}

function renderAll() {
  renderSummary();
  renderCalendar();
  renderWeather();
  renderDay();
  renderAssignments();
  renderNextDay();
  renderMealStatus();
  renderS4();
  renderReceipts();
  renderNotes();
  saveSettings();
  refreshIcons();
}

function renderSummary() {
  const dayEvents = eventsForDate(state.selectedDate);
  const tracks = activeTracksForDate(state.selectedDate);
  const openTasks = tasksForDate(state.selectedDate).filter((task) => task.status !== "done").length;
  const nextMeal = dayEvents.find((event) => event.category === "Meals" && (event.start || "99:99") >= new Date().toTimeString().slice(0, 5)) || dayEvents.find((event) => event.category === "Meals");
  $("#summaryTracks").textContent = tracks.length ? tracks.map((track) => track.replace("Air Assault ", "AA")).join(" / ") : "Shared";
  $("#summaryOpenTasks").textContent = `${openTasks} open`;
  $("#summaryNextMeal").textContent = nextMeal ? `${nextMeal.start || ""} ${nextMeal.title}`.trim() : "No meal found";
}

function renderCalendar() {
  $("#monthLabel").textContent = formatMonth(state.viewMonth);
  const grid = $("#calendarGrid");
  grid.replaceChildren();
  const year = state.viewMonth.getFullYear();
  const month = state.viewMonth.getMonth();
  const first = new Date(year, month, 1);
  const start = new Date(first);
  start.setDate(first.getDate() - first.getDay());
  const counts = countBy(filteredEvents(), (event) => event.date);
  for (let i = 0; i < 42; i += 1) {
    const date = new Date(start);
    date.setDate(start.getDate() + i);
    const iso = toIso(date);
    const button = createElement("button", { className: "calendar-day compact", attrs: { type: "button", "aria-label": formatFullDate(iso) } });
    if (date.getMonth() !== month) button.classList.add("is-muted");
    if (iso === state.selectedDate) button.classList.add("is-selected");
    if (iso === toIso(new Date())) button.classList.add("is-today");
    button.append(createElement("span", { className: "day-number", text: String(date.getDate()) }));
    button.append(createElement("strong", { text: counts[iso] ? String(counts[iso]) : "" }));
    button.addEventListener("click", () => selectDate(iso));
    grid.append(button);
  }
}

function formatWeatherUpdated(value) {
  if (!value) return "Live forecast";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Live forecast";
  return `Updated ${date.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" })}`;
}

function renderWeather() {
  renderTodayWeatherStrip();
  const panel = $("#weatherPanel");
  if (!panel) return;
  panel.replaceChildren();
  const header = createElement("div", { className: "weather-head" });
  const title = createElement("div");
  title.append(createElement("span", { className: "weather-title", text: "Weather" }), createElement("small", { text: WEATHER_LOCATION }));
  header.append(title, createElement("span", { className: "weather-updated", text: formatWeatherUpdated(state.weather.updatedAt) }));

  if (state.weather.status === "loading") {
    panel.append(header, createElement("div", { className: "weather-empty", text: "Loading West Point forecast..." }));
    return;
  }

  if (state.weather.status === "error") {
    panel.append(header, createElement("div", { className: "weather-empty", text: "Weather unavailable." }));
    return;
  }

  const days = state.weather.days || [];
  if (!days.length) {
    panel.append(header, createElement("div", { className: "weather-empty", text: "No weather data." }));
    return;
  }

  const focus = createElement("div", { className: "weather-focus" });
  days.slice(0, 2).forEach((day) => {
    const card = createElement("article", { className: "weather-card" });
    const cardHead = createElement("div", { className: "weather-card-head" });
    cardHead.append(createElement("span", { text: day.label }), createElement("strong", { text: `${formatTemp(day.high)} / ${formatTemp(day.low)}` }));
    card.append(cardHead, createElement("p", { text: day.condition }));
    const humidityText = day.humidity === undefined ? "Humidity TBD" : `Humidity ${day.humidity}%`;
    const heatIndexText = day.heatIndex === undefined ? "HI TBD" : `HI ${formatTemp(day.heatIndex)}`;
    card.append(createElement("small", { text: `${humidityText} / ${heatIndexText}` }));
    focus.append(card);
  });

  const week = createElement("div", { className: "weather-week", attrs: { "aria-label": "Seven day temperatures" } });
  days.slice(0, 7).forEach((day) => {
    const chip = createElement("div", { className: "weather-day" });
    chip.append(createElement("span", { text: day.weekday }), createElement("strong", { text: formatTemp(day.high) }), createElement("small", { text: formatTemp(day.low) }));
    week.append(chip);
  });

  panel.append(header, focus, week);
}

function selectedWeatherLabel() {
  if (!state.selectedDate) return "Selected Weather";
  const todayIso = toIso(new Date());
  if (state.selectedDate === todayIso) return "Weather Today";
  if (state.selectedDate === addDays(todayIso, 1)) return "Weather Tomorrow";
  return `Weather ${formatCompactDate(state.selectedDate)}`;
}

function renderTodayWeatherStrip() {
  const strip = $("#todayWeatherStrip");
  if (!strip) return;
  strip.replaceChildren();
  const label = selectedWeatherLabel();
  if (state.weather.status === "loading") {
    strip.append(createElement("span", { className: "today-weather-label", text: label }), createElement("strong", { className: "today-weather-chip", text: "Loading forecast" }));
    return;
  }
  if (state.weather.status === "error") {
    strip.append(createElement("span", { className: "today-weather-label", text: label }), createElement("strong", { className: "today-weather-chip", text: "Weather unavailable" }));
    return;
  }
  const selectedWeather = (state.weather.days || []).find((day) => day.iso === state.selectedDate);
  if (!selectedWeather) {
    strip.append(createElement("span", { className: "today-weather-label", text: label }), createElement("strong", { className: "today-weather-chip", text: "Forecast unavailable" }));
    return;
  }
  const chips = [
    selectedWeather.condition,
    `${formatTemp(selectedWeather.high)} / ${formatTemp(selectedWeather.low)}`,
    selectedWeather.heatIndex === undefined ? "HI TBD" : `HI ${formatTemp(selectedWeather.heatIndex)}`,
    selectedWeather.humidity === undefined ? "Humidity TBD" : `Humidity ${selectedWeather.humidity}%`,
  ];
  strip.append(createElement("span", { className: "today-weather-label", text: label }));
  chips.forEach((chip, index) => {
    strip.append(createElement(index === 1 ? "strong" : "span", { className: "today-weather-chip", text: chip }));
  });
}

function selectDate(iso) {
  state.selectedDate = iso;
  state.viewMonth = parseDate(iso);
  if (!state.taskFormOpen || !state.editingTaskId) $("#taskDate").value = iso;
  renderAll();
}

function renderDay() {
  const dayEvents = eventsForDate(state.selectedDate).sort(eventSort);
  const byTrack = eventsByTrack(dayEvents);
  $("#selectedDateLabel").textContent = formatFullDate(state.selectedDate);
  $("#selectedDaySubtitle").textContent = `${dayEvents.length} merged event${dayEvents.length === 1 ? "" : "s"} from ${state.sourceFilter === "all" ? "all sources" : state.sourceFilter}`;
  $("#dayTrackOne").textContent = (byTrack[0]?.events.length || 0).toString();
  $("#dayTrackTwo").textContent = (byTrack[1]?.events.length || 0).toString();
  $("#dayShared").textContent = byTrack.shared.length.toString();
  $("#dayMeals").textContent = dayEvents.filter((event) => event.category === "Meals").length.toString();
  $$(".segment[data-day-view]").forEach((button) => button.classList.toggle("is-active", button.dataset.dayView === state.dayView));
  $("#trackTimelines").classList.toggle("is-hidden", state.dayView !== "tracks");
  $("#dayTimeline").classList.toggle("is-hidden", state.dayView !== "all");
  $("#dayTableWrap").classList.toggle("is-hidden", state.dayView !== "table");
  renderTrackTimelines(byTrack);
  renderTimeline(dayEvents, $("#dayTimeline"));
  renderDayTable(dayEvents);
}

function activeTracksForDate(iso) {
  const tracks = Array.from(new Set(state.events.filter((event) => event.date === iso && event.classKey).map((event) => event.classKey)));
  return tracks.sort((a, b) => (classNumberFromKey(a) || 99) - (classNumberFromKey(b) || 99)).slice(0, 2);
}

function eventsByTrack(dayEvents) {
  const tracks = activeTracksForDate(state.selectedDate);
  const byTrack = tracks.map((track) => ({ track, dayLabel: classDayLabelForTrack(state.selectedDate, track), events: dayEvents.filter((event) => event.classKey === track) }));
  const shared = dayEvents.filter((event) => !event.classKey || !tracks.includes(event.classKey));
  return { tracks, 0: byTrack[0], 1: byTrack[1], shared };
}

function renderTrackTimelines(byTrack) {
  const wrap = $("#trackTimelines");
  wrap.replaceChildren();
  const hasTwoTracks = byTrack.tracks.length > 1;
  const columns = byTrack.tracks.map((track, index) => ({
    title: [track, byTrack[index]?.dayLabel].filter(Boolean).join(" - "),
    events: byTrack[index]?.events || [],
    className: hasTwoTracks ? "track-column" : "track-column full-column",
  }));
  if (!columns.length) wrap.append(createElement("div", { className: "empty", text: "No active Air Assault tracks for this day." }));
  columns.forEach((column) => {
    const pane = createElement("section", { className: column.className });
    pane.append(createElement("h3", { text: column.title }));
    renderTimeline(column.events.sort(eventSort), pane);
    wrap.append(pane);
  });
}

function renderTimeline(events, wrap) {
  const heading = wrap.matches?.(".track-column") ? wrap.querySelector("h3") : null;
  wrap.replaceChildren();
  if (heading) wrap.append(heading);
  if (!events.length) {
    wrap.append(createElement("div", { className: "empty", text: "No matching events for this lane." }));
    return;
  }
  const grouped = events.reduce((acc, event) => {
    const key = event.start || "Unscheduled";
    acc[key] ||= [];
    acc[key].push(event);
    return acc;
  }, {});
  Object.entries(grouped).forEach(([time, items]) => {
    const group = createElement("section", { className: "time-group" });
    group.append(createElement("div", { className: "time-label", text: time }));
    const list = createElement("div", { className: "event-stack" });
    items.forEach((event) => list.append(renderEventCard(event)));
    group.append(list);
    wrap.append(group);
  });
}

function renderEventCard(event) {
  const editedClass = event.isEdited ? " is-edited" : "";
  const card = createElement("button", { className: `event-card ${CATEGORY_CLASS[event.category] || "cat-operations"}${editedClass}`, attrs: { type: "button" } });
  const top = createElement("div", { className: "event-top" });
  top.append(createElement("span", { className: "badge", text: event.category }));
  const chips = createElement("span", { className: "event-chips" });
  if (event.isEdited) chips.append(createElement("span", { className: "source-chip edited-chip", text: "Edited" }));
  if (event.classKey) chips.append(createElement("span", { className: "source-chip", text: event.classKey.replace("Air Assault ", "AA") }));
  if (chips.children.length) top.append(chips);
  card.append(top);
  card.append(createElement("strong", { text: event.title }));
  const meta = [event.location, event.notes].filter(Boolean).join(" / ");
  if (meta) card.append(createElement("p", { text: meta }));
  card.addEventListener("click", () => openEvent(event));
  return card;
}

function renderDayTable(events) {
  const wrap = $("#dayTableWrap");
  wrap.replaceChildren();
  const table = createElement("table");
  table.append(row(["Time", "Track", "Category", "Task", "Location", "Source"], "th"));
  const body = createElement("tbody");
  events.forEach((event) => {
    body.append(row([event.start || "", event.classKey || "Shared", event.category, event.title, event.location || "", sourceSummary(event)]));
  });
  table.append(body);
  wrap.append(table);
}

function row(values, cellTag = "td") {
  const tr = createElement("tr");
  values.forEach((value) => tr.append(createElement(cellTag, { text: value })));
  return tr;
}

function taskTitleFromEvent(event) {
  return `${event.start || ""} ${event.classKey ? `${event.classKey}: ` : ""}${event.title}`.trim();
}

function tasksForDate(iso) {
  return state.tasks.filter((task) => task.date === iso).sort((a, b) => (a.start || "99:99").localeCompare(b.start || "99:99") || a.title.localeCompare(b.title));
}

function resetTaskForm(date = state.selectedDate) {
  $("#taskForm").reset();
  $("#taskDate").value = date || state.selectedDate;
  $("#taskSlots").value = 1;
  $("#taskTaskees").value = 0;
  $("#taskTrack").value = "Shared";
  $("#taskStatus").value = "open";
  setSelectedOwners($("#taskOwners"), []);
}

function setTaskFormOpen(open, mode = "new") {
  state.taskFormOpen = open;
  const form = $("#taskForm");
  form.classList.toggle("is-hidden", !open);
  form.setAttribute("aria-hidden", String(!open));
  $("#quickTaskButton").classList.toggle("is-active", open && mode === "new");
  $("#taskFormMode").textContent = mode === "edit" ? "Editing Task" : "New Task";
  $("#taskFormTitle").textContent = mode === "edit" ? "Editing Task" : "New Task";
  $("#taskSubmitLabel").textContent = mode === "edit" ? "Save Changes" : "Save Task";
  refreshIcons();
}

function openNewTaskForm() {
  state.editingTaskId = null;
  resetTaskForm(state.selectedDate);
  setTaskFormOpen(true, "new");
  $("#taskTitle").focus();
}

function closeTaskForm() {
  state.editingTaskId = null;
  resetTaskForm(state.selectedDate);
  setTaskFormOpen(false);
}

function saveTaskFromForm(event) {
  event.preventDefault();
  const existingTask = state.editingTaskId ? state.tasks.find((task) => task.id === state.editingTaskId) : null;
  const record = {
    id: state.editingTaskId || crypto.randomUUID(),
    title: $("#taskTitle").value.trim(),
    date: $("#taskDate").value || state.selectedDate,
    start: $("#taskStart").value,
    end: $("#taskEnd").value,
    track: $("#taskTrack").value,
    owners: selectedOwnersFromForm(),
    requiredPeople: Number($("#taskSlots").value || 1),
    taskees: Number($("#taskTaskees").value || 0),
    status: statusWithOwners($("#taskStatus").value, selectedOwnersFromForm()),
    notes: $("#taskNotes").value.trim(),
    updatedAt: new Date().toISOString(),
  };
  if (existingTask) {
    ["sourceKind", "sourceSeedId", "sourceSeedIds", "sourceEventId"].forEach((key) => {
      if (existingTask[key] !== undefined) record[key] = existingTask[key];
    });
  }
  record.owner = record.owners[0] || "Unassigned";
  state.tasks = state.tasks.filter((task) => task.id !== record.id);
  state.tasks.push(record);
  state.editingTaskId = null;
  saveTasks();
  closeTaskForm();
  renderAll();
}

function editTask(task) {
  state.editingTaskId = task.id;
  setTaskFormOpen(true, "edit");
  $("#taskTitle").value = task.title;
  $("#taskDate").value = task.date;
  $("#taskStart").value = task.start || "";
  $("#taskEnd").value = task.end || "";
  $("#taskTrack").value = task.track || "Shared";
  $("#taskSlots").value = taskRequiredPeople(task);
  $("#taskTaskees").value = taskTaskees(task);
  setSelectedOwners($("#taskOwners"), taskOwners(task));
  $("#taskStatus").value = task.status || "open";
  $("#taskNotes").value = task.notes || "";
  $("#taskTitle").focus();
}

function createTaskFromEvent(event) {
  const task = {
    id: crypto.randomUUID(),
    title: taskTitleFromEvent(event),
    date: event.date,
    start: event.start || "",
    end: event.end || "",
    track: event.classKey || "Shared",
    owners: [],
    owner: "Unassigned",
    requiredPeople: 1,
    taskees: 0,
    status: "open",
    notes: [event.location, event.notes, sourceSummary(event)].filter(Boolean).join(" / "),
    sourceEventId: event.id,
    updatedAt: new Date().toISOString(),
  };
  state.tasks.push(task);
  saveTasks();
  $("#eventDialog").close();
  renderAll();
}

function openEventEditor(event) {
  $("#eventDialog").close();
  $("#eventEditForm").dataset.mode = "edit";
  $("#eventEditId").value = event.id;
  $("#eventEditModeLabel").textContent = event.isManual ? "Edit Manual Event" : "Edit Timeline Event";
  $("#eventEditHeading").textContent = event.title;
  $("#eventEditTitle").value = event.title || "";
  $("#eventEditDate").value = event.date || state.selectedDate;
  $("#eventEditStart").value = event.start || "";
  $("#eventEditEnd").value = event.end || "";
  $("#eventEditTrack").value = event.classKey || "";
  $("#eventEditCategory").value = event.category || "Operations";
  $("#eventEditLocation").value = event.location || "";
  $("#eventEditNotes").value = event.notes || "";
  $("#resetEventEdit").hidden = !state.eventOverrides[event.id];
  $("#resetEventEdit").textContent = event.isManual ? "Delete Event" : "Reset Source";
  $("#eventEditSubmitLabel").textContent = "Save Event";
  $("#eventEditDialog").showModal();
  refreshIcons();
}

function openNewEventEditor() {
  const tracks = activeTracksForDate(state.selectedDate);
  $("#eventDialog").close();
  $("#eventEditForm").dataset.mode = "create";
  $("#eventEditId").value = "";
  $("#eventEditModeLabel").textContent = "New Timeline Event";
  $("#eventEditHeading").textContent = "New Event";
  $("#eventEditTitle").value = "";
  $("#eventEditDate").value = state.selectedDate;
  $("#eventEditStart").value = "";
  $("#eventEditEnd").value = "";
  $("#eventEditTrack").value = tracks[0] || "";
  $("#eventEditCategory").value = "Training";
  $("#eventEditLocation").value = "";
  $("#eventEditNotes").value = "";
  $("#resetEventEdit").hidden = true;
  $("#resetEventEdit").textContent = "Reset Source";
  $("#eventEditSubmitLabel").textContent = "Add Event";
  $("#eventEditDialog").showModal();
  $("#eventEditTitle").focus();
  refreshIcons();
}

function saveEventEdit(event) {
  event.preventDefault();
  const existingId = $("#eventEditId").value;
  const id = existingId || createManualEventId();
  const isManual = !existingId || state.eventOverrides[existingId]?.sourceKind === "manual-event";
  const override = {
    title: $("#eventEditTitle").value.trim(),
    date: $("#eventEditDate").value,
    start: $("#eventEditStart").value,
    end: $("#eventEditEnd").value,
    classKey: $("#eventEditTrack").value,
    category: $("#eventEditCategory").value,
    location: $("#eventEditLocation").value.trim(),
    notes: $("#eventEditNotes").value.trim(),
    updatedAt: new Date().toISOString(),
  };
  if (isManual) {
    override.sourceKind = "manual-event";
    state.sourceFilter = "all";
    $("#sourceFilter").value = "all";
    saveSettings();
  }
  state.eventOverrides[id] = override;
  saveEventOverrides();
  $("#eventEditDialog").close();
  rebuildEvents();
  state.selectedDate = override.date || state.selectedDate;
  state.viewMonth = parseDate(state.selectedDate);
  $("#taskDate").value = state.selectedDate;
  renderAll();
}

function resetEditedEvent() {
  const id = $("#eventEditId").value;
  if (!id) return;
  const wasManual = state.eventOverrides[id]?.sourceKind === "manual-event";
  delete state.eventOverrides[id];
  saveEventOverrides();
  $("#eventEditDialog").close();
  rebuildEvents();
  const restored = wasManual ? null : eventById(id);
  if (restored) {
    state.selectedDate = restored.date;
    state.viewMonth = parseDate(restored.date);
    $("#taskDate").value = restored.date;
  }
  renderAll();
}

function renderAssignments() {
  const dayTasks = tasksForDate(state.selectedDate);
  $("#assignmentCount").textContent = `${dayTasks.filter((task) => task.status !== "done").length} open`;
  $$(".segment[data-assignment-view]").forEach((button) => button.classList.toggle("is-active", button.dataset.assignmentView === state.assignmentView));
  $("#assignmentTimeline").classList.toggle("is-hidden", state.assignmentView !== "timeline");
  $("#assignmentBoard").classList.toggle("is-hidden", state.assignmentView !== "list");
  renderAssignmentTimeline(dayTasks);
  renderAssignmentList(dayTasks);
  refreshIcons();
}

function renderAssignmentTimeline(tasks) {
  const wrap = $("#assignmentTimeline");
  wrap.replaceChildren();
  if (!tasks.length) {
    wrap.append(createElement("div", { className: "empty", text: "No cadre taskings yet. Add one above or create one from a schedule event." }));
    return;
  }
  const laneTasks = new Map(TRACKED_PEOPLE.map((person) => [person, []]));
  tasks.forEach((task) => {
    const owners = taskOwners(task);
    const people = owners.length ? owners : ["Unassigned"];
    people.forEach((person) => laneTasks.get(person)?.push(task));
  });
  TRACKED_PEOPLE.forEach((person) => {
    const tasksForPerson = laneTasks.get(person) || [];
    if (!tasksForPerson.length) return;
    const lane = createElement("article", { className: "person-lane" });
    lane.append(createElement("strong", { text: person }));
    tasksForPerson.forEach((task) => lane.append(renderTaskMini(task)));
    wrap.append(lane);
  });
}

function renderAssignmentList(tasks) {
  const list = $("#assignmentBoard");
  list.replaceChildren();
  if (!tasks.length) {
    list.append(createElement("div", { className: "empty", text: "No taskings for this day." }));
    return;
  }
  tasks.forEach((task) => list.append(renderTaskRow(task)));
}

function rememberDeletedTaskSource(task) {
  const ids = taskSourceSeedIds(task);
  if (!ids.length) return;
  state.deletedSource.taskings = Array.from(new Set([...(state.deletedSource.taskings || []), ...ids]));
  saveDeletedSource();
}

function deleteTask(id) {
  const task = state.tasks.find((candidate) => candidate.id === id);
  if (task) rememberDeletedTaskSource(task);
  state.tasks = state.tasks.filter((candidate) => candidate.id !== id);
  if (state.editingTaskId === id) closeTaskForm();
  saveTasks();
  renderAll();
}

function renderTaskMini(task) {
  const item = createElement("article", { className: `task-mini status-${task.status}`, attrs: { "data-task-id": task.id } });
  const main = createElement("button", { className: "task-mini-main", attrs: { type: "button", "aria-label": `Open ${task.title}` } });
  main.append(createElement("span", { text: task.start || "Unscheduled" }));
  main.append(createElement("strong", { text: task.title }));
  main.append(createElement("small", { text: `${task.track || "Shared"} / ${task.status} / ${taskStaffingText(task)}` }));
  main.addEventListener("click", () => editTask(task));

  const actions = createElement("div", { className: "task-mini-actions" });
  const remove = createElement("button", { className: "icon-button", html: '<i data-lucide="trash-2"></i>', attrs: { type: "button", title: "Remove", "aria-label": `Remove ${task.title}` } });
  remove.addEventListener("click", () => deleteTask(task.id));
  actions.append(remove);
  item.append(main, actions);
  return item;
}

function renderTaskRow(task) {
  const rowEl = createElement("article", { className: `task-row status-${task.status}`, attrs: { "data-task-id": task.id } });
  const main = createElement("div", { className: "task-main", attrs: { role: "button", tabindex: "0", "aria-label": `Open ${task.title}` } });
  main.addEventListener("click", (event) => {
    if (event.target.closest("button")) return;
    editTask(task);
  });
  main.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    event.preventDefault();
    editTask(task);
  });
  main.append(createElement("strong", { text: task.title }));
  main.append(createElement("small", { className: "task-line", text: `${task.start || "Unscheduled"}${task.end ? `-${task.end}` : ""} / ${task.track || "Shared"} / ${task.status || "open"} / ${taskStaffingText(task)}` }));
  main.append(createElement("small", { className: "task-owner-line", text: `Owners: ${ownersLabel(task)}` }));
  const controls = createElement("div", { className: "task-row-controls" });
  const remove = createElement("button", { className: "icon-button", html: '<i data-lucide="trash-2"></i>', attrs: { type: "button", title: "Remove", "aria-label": "Remove" } });
  remove.addEventListener("click", () => deleteTask(task.id));
  controls.append(remove);
  rowEl.append(main, controls);
  return rowEl;
}

function updateTask(id, patch) {
  state.tasks = state.tasks.map((task) => (task.id === id ? { ...task, ...patch, updatedAt: new Date().toISOString() } : task));
  saveTasks();
  renderAll();
}

function renderNextDay() {
  const nextIso = addDays(state.selectedDate, 1);
  const nextEvents = eventsForDate(nextIso).sort(eventSort);
  const nextTasks = tasksForDate(nextIso).filter((task) => task.status !== "done");
  $("#nextDrawerDate").textContent = formatFullDate(nextIso);
  $("#nextDayCount").textContent = nextEvents.length;
  $("#nextTaskCount").textContent = nextTasks.length;
  const list = $("#nextDayList");
  list.replaceChildren();
  nextEvents.slice(0, 18).forEach((event) => list.append(renderCompactEvent(event)));
  if (!nextEvents.length) list.append(createElement("div", { className: "empty", text: "No events found for next day." }));
  const taskList = $("#nextTaskList");
  taskList.replaceChildren();
  nextTasks.forEach((task) => taskList.append(renderTaskMini(task)));
  if (!nextTasks.length) taskList.append(createElement("div", { className: "empty", text: "No open taskings for next day." }));
}

function renderMealStatus() {
  const nextIso = addDays(state.selectedDate, 1);
  const meals = state.events.filter((event) => event.date === nextIso && event.category === "Meals").sort(eventSort);
  const board = $("#mealStatusBoard");
  board.replaceChildren();
  $("#mealStatusCount").textContent = meals.length;
  if (!meals.length) {
    board.append(createElement("div", { className: "empty", text: "No Air Assault meal status for next day." }));
    return;
  }
  const byClass = {};
  meals.forEach((event) => {
    const key = event.classKey || "Shared";
    byClass[key] ||= [];
    byClass[key].push(event);
  });
  Object.entries(byClass).slice(0, 2).forEach(([unit, unitMeals]) => {
    const card = createElement("article", { className: "meal-card" });
    card.append(createElement("strong", { text: unit }));
    ["Breakfast", "Lunch", "Dinner"].forEach((mealName) => {
      const event = unitMeals.find((item) => item.title.includes(mealName));
      const line = createElement("button", { className: "meal-line", attrs: { type: "button" } });
      line.append(createElement("span", { text: mealName }));
      if (event) {
        const status = mealStatusFromText(event.title);
        line.append(createElement("b", { className: status === "MRE" ? "meal-mre" : status === "CM" ? "meal-cm" : "meal-other", text: status || "TBD" }));
        line.append(createElement("small", { text: [event.start, event.location].filter(Boolean).join(" / ") }));
        line.addEventListener("click", () => openEvent(event));
      } else {
        line.append(createElement("b", { className: "meal-empty", text: "No row" }));
      }
      card.append(line);
    });
    board.append(card);
  });
}

function renderCompactEvent(event) {
  const item = createElement("button", { className: "compact-event", attrs: { type: "button" } });
  item.append(createElement("span", { text: event.start || "Unscheduled" }));
  item.append(createElement("strong", { text: `${event.classKey ? `${event.classKey}: ` : ""}${event.title}` }));
  item.addEventListener("click", () => openEvent(event));
  return item;
}

function toggleNextDrawer(force) {
  const drawer = $("#nextDrawer");
  const open = typeof force === "boolean" ? force : drawer.getAttribute("aria-hidden") === "true";
  drawer.setAttribute("aria-hidden", open ? "false" : "true");
  drawer.classList.toggle("is-open", open);
  document.body.classList.toggle("drawer-open", open);
  $("#drawerBackdrop").hidden = !open;
}

function migrateS4(items) {
  return items
    .map(normalizeS4Item)
    .filter((item) => item.item);
}

function configureS4Form() {
  const type = $("#s4Type")?.value || "Draw";
  const config = S4_FIELD_CONFIG[type] || S4_FIELD_CONFIG.Draw;
  const visible = new Set(config.visible);
  $$(".s4-field[data-s4-field]").forEach((field) => {
    const key = field.dataset.s4Field;
    const isVisible = visible.has(key);
    field.hidden = !isVisible;
    $$("input, select, textarea", field).forEach((input) => {
      input.disabled = !isVisible;
      input.required = isVisible && (config.required || []).includes(key);
    });
  });
  $("#s4ItemLabel").textContent = config.itemLabel || "Item";
  $("#s4Item").placeholder = config.itemPlaceholder || "";
  $("#s4QtyLabel").textContent = config.qtyLabel || "Qty";
  $("#s4NeedByLabel").textContent = config.timeLabel || "Date / Time";
  const statusSelect = $("#s4Status");
  const current = statusSelect.value;
  statusSelect.replaceChildren();
  (S4_STATUS_OPTIONS[type] || S4_STATUS_OPTIONS.Draw).forEach((status) => statusSelect.append(createElement("option", { text: status, attrs: { value: status } })));
  statusSelect.value = (S4_STATUS_OPTIONS[type] || []).includes(current) ? current : (S4_STATUS_OPTIONS[type] || S4_STATUS_OPTIONS.Draw)[0];
}

function addS4Item(event) {
  event.preventDefault();
  const type = $("#s4Type")?.value || "Draw";
  const qty = Number($("#s4Qty")?.value || 1);
  const placement = $("#s4Placement")?.value.trim() || "";
  const baseItem = $("#s4Item")?.value.trim() || "";
  const item = {
    id: crypto.randomUUID(),
    type,
    item: baseItem,
    qty,
    needBy: $("#s4NeedBy")?.value || "",
    requestTime: $("#s4NeedBy")?.value || "",
    drawTime: $("#s4NeedBy")?.value || "",
    status: $("#s4Status")?.value || "Planned",
    owner: "",
    notes: $("#s4Notes")?.value.trim() || "",
    requestKind: $("#s4RequestKind")?.value || "",
    location: $("#s4Location")?.value || "",
    placement,
    pax: Number($("#s4Pax")?.value || 0),
    pickupTime: $("#s4PickupTime")?.value || "",
    returnTime: $("#s4ReturnTime")?.value || "",
  };
  state.s4Items.unshift(item);
  state.s4Items = migrateS4(state.s4Items);
  saveS4();
  event.target.reset();
  $("#s4Type").value = type;
  $("#s4Qty").value = 1;
  $("#s4Pax").value = 0;
  configureS4Form();
  renderS4();
}

function rememberDeletedSupportSource(item) {
  if (!item?.sourceSeedId) return;
  state.deletedSource.supportItems = Array.from(new Set([...(state.deletedSource.supportItems || []), item.sourceSeedId]));
  saveDeletedSource();
}

function deleteS4Item(item) {
  rememberDeletedSupportSource(item);
  state.s4Items = state.s4Items.filter((candidate) => candidate.id !== item.id);
  saveS4();
  renderS4();
}

function formatSupportDateTime(value) {
  if (!value) return "TBD";
  const [datePart, timePart = ""] = String(value).split("T");
  const dateLabel = datePart ? formatCompactDate(datePart) : "";
  const timeLabel = normalizeTime(timePart) || "";
  return [dateLabel, timeLabel].filter(Boolean).join(" ");
}

function renderS4() {
  const list = $("#s4List");
  list.replaceChildren();
  S4_LANES.forEach((type) => {
    const laneItems = state.s4Items.filter((item) => normalizeS4Lane(item.type) === type);
    const lane = createElement("section", { className: "s4-lane", attrs: { "data-lane": type } });
    const laneHead = createElement("div", { className: "s4-lane-head" });
    laneHead.append(createElement("h3", { text: type }), renderS4LaneSummary(type, laneItems));
    const laneBody = createElement("div", { className: "s4-lane-body" });
    if (!laneItems.length) laneBody.append(createElement("div", { className: "empty compact", text: "Nothing tracked." }));
    laneItems
      .sort((a, b) => (s4PrimaryTime(a) || "9999").localeCompare(s4PrimaryTime(b) || "9999"))
      .forEach((item) => laneBody.append(renderS4Item(item)));
    lane.append(laneHead, laneBody);
    list.append(lane);
  });
  refreshIcons();
}

function renderS4LaneSummary(type, items) {
  const wrap = createElement("div", { className: "s4-lane-summary" });
  const options = S4_STATUS_OPTIONS[type] || S4_STATUS_OPTIONS.Draw;
  options.forEach((status) => {
    const count = items.filter((item) => normalizeS4Item(item).status === status).length;
    wrap.append(createElement("span", { text: `${status} ${count}` }));
  });
  return wrap;
}

function s4PrimaryTime(item) {
  item = normalizeS4Item(item);
  if (item.type === "Draw") return item.drawTime || item.needBy;
  if (item.type === "LMTV TMR") return item.requestTime || item.needBy || item.pickupTime;
  return item.requestTime || item.needBy;
}

function s4Title(item) {
  item = normalizeS4Item(item);
  if (item.type === "Water Buffalo") return `${item.requestKind || "Refill"} Water Buffalo`;
  if (item.type === "LMTV TMR") return item.item || "LMTV Round Trip";
  if (item.type === "FLA Request") return item.item || `${item.qty || 1}x FLA Request`;
  return item.item || "Draw";
}

function s4Meta(item) {
  item = normalizeS4Item(item);
  if (item.type === "Water Buffalo") {
    return [`TMR ${item.status}`, item.requestKind, item.location, formatSupportDateTime(item.requestTime || item.needBy)].filter(Boolean);
  }
  if (item.type === "LMTV TMR") {
    return [`TMR ${item.status}`, `${item.pax || 0} PAX`, "Round trip to MARNE", formatSupportDateTime(item.requestTime || item.needBy)].filter(Boolean);
  }
  if (item.type === "FLA Request") {
    return [`${item.qty || 1} FLA`, formatSupportDateTime(item.requestTime || item.needBy), item.placement || item.location || "Placement TBD"].filter(Boolean);
  }
  return [`${item.qty || 1} item${Number(item.qty || 1) === 1 ? "" : "s"}`, formatSupportDateTime(item.drawTime || item.needBy), item.status].filter(Boolean);
}

function s4DetailLines(item) {
  item = normalizeS4Item(item);
  const lines = [];
  if (item.type === "LMTV TMR") {
    if (item.pickupTime) lines.push(`Pickup: ${formatSupportDateTime(item.pickupTime)}`);
    if (item.returnTime) lines.push(`Return: ${formatSupportDateTime(item.returnTime)}`);
  }
  if (item.owner) lines.push(`Owner: ${item.owner}`);
  if (item.notes) lines.push(item.notes);
  return lines;
}

function renderS4Item(item) {
  const normalized = normalizeS4Item(item);
  const rowEl = createElement("article", { className: `s4-item status-${safeClassName(normalized.status)}` });
  const main = createElement("div", { className: "s4-main" });
  main.append(createElement("strong", { text: s4Title(normalized) }));
  const meta = createElement("div", { className: "s4-meta" });
  s4Meta(normalized).forEach((line) => meta.append(createElement("span", { text: line })));
  main.append(meta);
  s4DetailLines(normalized).forEach((line) => main.append(createElement("small", { text: line })));
  const status = createElement("select", { attrs: { "aria-label": `Status for ${normalized.item}` } });
  const statusOptions = S4_STATUS_OPTIONS[normalized.type] || S4_STATUS_OPTIONS.Draw;
  Array.from(new Set([...statusOptions, normalized.status])).forEach((value) => status.append(createElement("option", { text: value, attrs: { value } })));
  status.value = normalized.status;
  status.addEventListener("change", () => {
    item.status = status.value;
    item.type = normalized.type;
    if (status.value === "Sent" || status.value === "Drawn" || status.value === "Complete") item.sentAt = item.sentAt || new Date().toISOString();
    saveS4();
    renderS4();
  });
  const remove = createElement("button", { className: "icon-button", html: '<i data-lucide="trash-2"></i>', attrs: { type: "button", title: "Remove item", "aria-label": "Remove item" } });
  remove.addEventListener("click", () => deleteS4Item(item));
  rowEl.append(main, status, remove);
  return rowEl;
}

function sourceSummary(event) {
  const sources = event.relatedSources || [event.source].filter(Boolean);
  return Array.from(new Set(sources.map((source) => `${source.file} ${source.sheet} ${source.cell}`))).slice(0, 3).join(" / ");
}

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(RECEIPT_DB, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(RECEIPT_STORE, { keyPath: "id" });
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function storeReceipt(file) {
  const db = await openDb();
  const record = { id: crypto.randomUUID(), name: file.name, type: file.type, size: file.size, addedAt: new Date().toISOString(), blob: file };
  if (state.cloud.storage && state.cloud.storageRef && state.cloud.uploadBytes && state.cloud.getDownloadURL) {
    try {
      const safeName = file.name.replace(/[^\w.\-]+/g, "_");
      record.storagePath = `receipts/${record.id}-${safeName}`;
      const fileRef = state.cloud.storageRef(state.cloud.storage, record.storagePath);
      await state.cloud.uploadBytes(fileRef, file);
      record.downloadUrl = await state.cloud.getDownloadURL(fileRef);
      record.synced = true;
    } catch (error) {
      console.warn("Firebase Storage upload failed; keeping receipt local", error);
      delete record.storagePath;
      record.synced = false;
      record.syncError = "Local file only";
    }
  } else {
    record.synced = false;
  }
  return new Promise((resolve, reject) => {
    const tx = db.transaction(RECEIPT_STORE, "readwrite");
    tx.objectStore(RECEIPT_STORE).put(record);
    tx.oncomplete = () => resolve(record);
    tx.onerror = () => reject(tx.error);
  });
}

async function loadReceipts() {
  const db = await openDb();
  state.receipts = await new Promise((resolve, reject) => {
    const tx = db.transaction(RECEIPT_STORE, "readonly");
    const request = tx.objectStore(RECEIPT_STORE).getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
  });
}

async function addReceipts(files) {
  if (!files.length) return;
  await Promise.all(files.map(storeReceipt));
  await loadReceipts();
  scheduleCloudSave("receipts");
  renderReceipts();
}

async function deleteReceipt(id) {
  const record = state.receipts.find((receipt) => receipt.id === id);
  if (record?.storagePath && state.cloud.storage && state.cloud.storageRef && state.cloud.deleteObject) {
    try {
      await state.cloud.deleteObject(state.cloud.storageRef(state.cloud.storage, record.storagePath));
    } catch (error) {
      console.warn("Could not delete receipt from Firebase Storage", error);
    }
  }
  const db = await openDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(RECEIPT_STORE, "readwrite");
    tx.objectStore(RECEIPT_STORE).delete(id);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
  state.receipts = state.receipts.filter((receipt) => receipt.id !== id);
  scheduleCloudSave("receipts");
  renderReceipts();
}

function renderReceipts() {
  const list = $("#receiptList");
  list.replaceChildren();
  $("#receiptCount").textContent = state.receipts.length;
  if (!state.receipts.length) {
    list.append(createElement("div", { className: "empty", text: "No hand receipts uploaded." }));
    return;
  }
  state.receipts.sort((a, b) => b.addedAt.localeCompare(a.addedAt)).forEach((record) => {
    const item = createElement("article", { className: "receipt-item" });
    const info = createElement("div");
    info.append(createElement("strong", { text: record.name }));
    info.append(createElement("span", { text: `${(record.size / 1024).toFixed(1)} KB / ${new Date(record.addedAt).toLocaleString()}` }));
    info.append(createElement("small", { text: record.downloadUrl ? "Synced file" : record.blob ? "Local file in this browser" : "Metadata only" }));
    const download = createElement("button", { className: "icon-button", html: '<i data-lucide="download"></i>', attrs: { type: "button", title: "Download", "aria-label": "Download" } });
    if (!record.blob && !record.downloadUrl) {
      download.disabled = true;
      download.title = "File is not available in this browser";
    }
    download.addEventListener("click", () => {
      const url = record.blob ? URL.createObjectURL(record.blob) : record.downloadUrl;
      if (!url) return;
      const a = createElement("a", { attrs: { href: url, download: record.name, target: "_blank", rel: "noopener" } });
      a.click();
      if (record.blob) URL.revokeObjectURL(url);
    });
    const remove = createElement("button", { className: "icon-button", html: '<i data-lucide="trash-2"></i>', attrs: { type: "button", title: "Delete", "aria-label": "Delete" } });
    remove.addEventListener("click", () => deleteReceipt(record.id));
    item.append(info, download, remove);
    list.append(item);
  });
  refreshIcons();
}

function selectTab(tab) {
  $$(".tab").forEach((button) => button.classList.toggle("is-active", button.dataset.tab === tab));
  $$(".tab-panel").forEach((panel) => panel.classList.toggle("is-active", panel.id === `tab-${tab}`));
  if (tab === "sources") renderSourceTable();
}

function renderSourceControls() {
  const sourceSelect = $("#sourceSelect");
  const sheetSelect = $("#sheetSelect");
  sourceSelect.replaceChildren();
  state.data.sources.forEach((source, index) => sourceSelect.append(createElement("option", { text: source.name, attrs: { value: String(index) } })));
  sourceSelect.value = String(state.sourceIndex);
  sheetSelect.replaceChildren();
  const source = state.data.sources[state.sourceIndex];
  source.sheets.forEach((sheet, index) => sheetSelect.append(createElement("option", { text: sheet.name, attrs: { value: String(index) } })));
  sheetSelect.value = String(state.sheetIndex);
}

function renderSourceTable() {
  const wrap = $("#sourceTable");
  wrap.replaceChildren();
  const source = state.data.sources[state.sourceIndex];
  const sheet = source?.sheets[state.sheetIndex];
  if (!sheet) return;
  const query = state.sourceSearch.trim();
  const matcher = query ? new RegExp(escapeRegex(query), "i") : null;
  const cells = sheet.cells.filter((cell) => !matcher || matcher.test(String(cell.v ?? "")) || matcher.test(cell.a)).slice(0, 350);
  const table = createElement("table");
  table.append(row(["Cell", "Value", "Formula"], "th"));
  const body = createElement("tbody");
  cells.forEach((cell) => body.append(row([cell.a, String(cell.v ?? ""), cell.f || ""])));
  table.append(body);
  wrap.append(table);
  if (sheet.cells.length > cells.length) wrap.append(createElement("p", { className: "table-note", text: `Showing ${cells.length} of ${sheet.cells.length} populated cells.` }));
}

function renderNotes() {
  const list = $("#notesList");
  const notes = state.data.notes || [];
  const diagnostics = state.data.diagnostics || [];
  $("#noteCount").textContent = diagnostics.length + notes.length;
  list.replaceChildren();
  if (!diagnostics.length && !notes.length) {
    list.append(createElement("div", { className: "empty", text: "No data diagnostics or Air Assault notes found." }));
    return;
  }

  diagnostics.forEach((item) => {
    const note = createElement("article", { className: `note-item severity-${item.severity || "info"}` });
    const label = [item.severity?.toUpperCase(), item.kind].filter(Boolean).join(" / ");
    note.append(createElement("strong", { text: label }));
    note.append(createElement("p", { text: item.message }));
    const detail = [item.date, item.classKey, item.sourceRefs?.[0] ? `${item.sourceRefs[0].file} ${item.sourceRefs[0].sheet} ${item.sourceRefs[0].cell}` : ""]
      .filter(Boolean)
      .join(" / ");
    if (detail) note.append(createElement("span", { text: detail }));
    list.append(note);
  });

  notes.forEach((note) => {
    const item = createElement("article", { className: "note-item" });
    item.append(createElement("strong", { text: `${note.sheet} / ${note.cell}` }));
    item.append(createElement("p", { text: note.text }));
    list.append(item);
  });
}

function openEvent(event) {
  $("#dialogCategory").textContent = event.category;
  $("#dialogTitle").textContent = event.title;
  const details = $("#dialogDetails");
  details.replaceChildren();
  [
    ["Date", formatFullDate(event.date)],
    ["Time", event.start ? `${event.start}${event.end ? `-${event.end}` : ""}` : "Unscheduled"],
    ["Track", event.classKey || "Shared"],
    ["Location", event.location],
    ["Notes", event.notes],
    ["Source", sourceSummary(event)],
    ["Local Edit", event.isEdited ? "Edited in this browser" : ""],
  ]
    .filter(([, value]) => value)
    .forEach(([label, value]) => details.append(createElement("dt", { text: label }), createElement("dd", { text: value })));
  const actions = createElement("div", { className: "dialog-action-row" });
  const taskAction = createElement("button", { className: "task-submit dialog-action", html: '<i data-lucide="clipboard-plus"></i><span>Create Tasking</span>', attrs: { type: "button" } });
  const editAction = createElement("button", { className: "secondary-button dialog-action", html: '<i data-lucide="pencil"></i><span>Edit Event</span>', attrs: { type: "button" } });
  taskAction.addEventListener("click", () => createTaskFromEvent(event));
  editAction.addEventListener("click", () => openEventEditor(event));
  actions.append(taskAction, editAction);
  details.append(createElement("dt", { text: "Action" }), createElement("dd", { children: [actions] }));
  $("#eventDialog").showModal();
  refreshIcons();
}

function printValue(value, fallback = "TBD") {
  const text = String(value ?? "").trim();
  return text ? escapeHtml(text) : escapeHtml(fallback);
}

function printTimeRange(item) {
  if (!item.start) return "Unscheduled";
  return `${item.start}${item.end ? `-${item.end}` : ""}`;
}

function printStat(label, value) {
  return `<article class="print-stat"><span>${escapeHtml(label)}</span><strong>${printValue(value)}</strong></article>`;
}

function printEventCard(event) {
  const meta = [event.location, event.notes].filter(Boolean).join(" / ");
  const chips = [event.classKey ? event.classKey.replace("Air Assault ", "AA") : "", event.isEdited ? "Edited" : ""].filter(Boolean);
  return `
    <article class="print-card print-event print-${safeClassName(event.category)}">
      <div class="print-card-top">
        <span class="print-badge">${printValue(event.category, "Operations")}</span>
        ${chips.length ? `<span class="print-chips">${chips.map((chip) => `<b>${escapeHtml(chip)}</b>`).join("")}</span>` : ""}
      </div>
      <strong>${printValue(event.title, "Untitled event")}</strong>
      ${meta ? `<p>${escapeHtml(meta)}</p>` : ""}
    </article>
  `;
}

function printEventMini(event) {
  const track = event.classKey ? event.classKey.replace("Air Assault ", "AA") : "";
  const meta = [event.category, event.location].filter(Boolean).join(" / ");
  return `
    <article class="print-event-mini print-${safeClassName(event.category)}">
      <strong>${printValue(event.title, "Untitled event")}</strong>
      <p>${[track, meta, event.notes].filter(Boolean).map(escapeHtml).join(" / ")}</p>
    </article>
  `;
}

function printTimelineMatrix(dayEvents) {
  if (!dayEvents.length) return `<div class="print-empty">No timeline items for this day.</div>`;
  const byTrack = eventsByTrack(dayEvents);
  const columns = byTrack.tracks.map((track, index) => {
    const dayLabel = byTrack[index]?.dayLabel;
    return { title: [track, dayLabel].filter(Boolean).join(" - "), events: byTrack[index]?.events || [] };
  });
  if (byTrack.shared.length) columns.push({ title: "Shared", events: byTrack.shared });
  const times = Array.from(new Set(dayEvents.map((event) => event.start || "Unscheduled"))).sort((a, b) => (a === "Unscheduled" ? "99:99" : a).localeCompare(b === "Unscheduled" ? "99:99" : b));
  const header = [`<div class="print-matrix-head print-matrix-time">Time</div>`, ...columns.map((column) => `<div class="print-matrix-head">${escapeHtml(column.title)}</div>`)].join("");
  const rows = times
    .map((time) => {
      const cells = columns
        .map((column) => {
          const items = column.events.filter((event) => (event.start || "Unscheduled") === time).sort(eventSort);
          return `<div class="print-matrix-cell">${items.length ? items.map(printEventMini).join("") : ""}</div>`;
        })
        .join("");
      return `<div class="print-matrix-time">${escapeHtml(time)}</div>${cells}`;
    })
    .join("");
  return `<div class="print-timeline-matrix" style="--timeline-cols: ${columns.length};">${header}${rows}</div>`;
}

function printTaskCard(task) {
  const notes = String(task.notes || "").trim();
  return `
    <article class="print-card print-task status-${safeClassName(task.status)}">
      <div class="print-card-top">
        <span class="print-time-chip">${escapeHtml(printTimeRange(task))}</span>
        <span class="print-badge">${printValue(task.status || "open")}</span>
      </div>
      <strong>${printValue(task.title, "Untitled task")}</strong>
      <p>${printValue(task.track || "Shared")} / ${escapeHtml(taskStaffingText(task))}</p>
      <p>Owners: ${escapeHtml(ownersLabel(task))}</p>
      ${notes ? `<p>${escapeHtml(notes)}</p>` : ""}
    </article>
  `;
}

function printAssignmentsByTime(tasks) {
  if (!tasks.length) return `<div class="print-empty">No cadre taskings for this day.</div>`;
  const laneTasks = new Map(TRACKED_PEOPLE.map((person) => [person, []]));
  tasks.forEach((task) => {
    const owners = taskOwners(task);
    const people = owners.length ? owners : ["Unassigned"];
    people.forEach((person) => laneTasks.get(person)?.push(task));
  });
  const lanes = TRACKED_PEOPLE.map((person) => {
    const tasksForPerson = laneTasks.get(person) || [];
    if (!tasksForPerson.length) return "";
    return `
      <section class="print-person-lane">
        <h3>${escapeHtml(person)}</h3>
        <div class="print-task-stack">${tasksForPerson.map(printTaskLine).join("")}</div>
      </section>
    `;
  }).filter(Boolean);
  return `<div class="print-person-grid">${lanes.join("")}</div>`;
}

function printTaskLine(task) {
  return `
    <article class="print-task-line status-${safeClassName(task.status)}">
      <span>${escapeHtml(printTimeRange(task))}</span>
      <strong>${printValue(task.title, "Untitled task")}</strong>
      <small>${printValue(task.track || "Shared")} / ${escapeHtml(taskStaffingText(task))}</small>
    </article>
  `;
}

function printAssignmentsList(tasks) {
  if (!tasks.length) return `<div class="print-empty">No cadre taskings for this day.</div>`;
  return `<div class="print-task-list">${tasks.map(printTaskCard).join("")}</div>`;
}

function printAssignments(tasks) {
  const done = tasks.filter((task) => task.status === "done").length;
  const claimed = tasks.filter((task) => task.status === "claimed").length;
  const open = tasks.filter((task) => task.status !== "done").length;
  const viewLabel = state.assignmentView === "list" ? "By Task" : "By Person";
  return `
    <section class="print-section">
      <div class="print-section-head">
        <div>
          <p class="print-eyebrow">Current Task List</p>
          <h2>Cadre Tasking - ${escapeHtml(viewLabel)}</h2>
        </div>
        <span>${open} open / ${claimed} claimed / ${done} done</span>
      </div>
      ${state.assignmentView === "list" ? printAssignmentsList(tasks) : printAssignmentsByTime(tasks)}
    </section>
  `;
}

function printMealSummary(dayEvents) {
  const meals = dayEvents.filter((event) => event.category === "Meals").sort(eventSort);
  if (!meals.length) return `<div class="print-empty">No meals found for this day.</div>`;
  return `
    <div class="print-meal-strip">
      ${meals
        .map((meal) => {
          const mealName = mealNameFromText(meal.title) || "Meal";
          const status = mealStatusFromText(meal.title) || String(meal.title || "").split(":").pop().trim() || "TBD";
          const unit = meal.classKey ? meal.classKey.replace("Air Assault ", "AA") : "Shared";
          return `
            <article class="print-meal">
              <span>${escapeHtml(unit)}</span>
              <strong>${escapeHtml(mealName)}: ${escapeHtml(status)}</strong>
              <em>${escapeHtml(printTimeRange(meal))}</em>
              <p>${printValue(meal.location, "Location TBD")}</p>
            </article>
          `;
        })
        .join("")}
    </div>
  `;
}

function printS4Item(item) {
  const normalized = normalizeS4Item(item);
  const details = s4DetailLines(normalized);
  return `
    <article class="print-card print-s4-item status-${safeClassName(normalized.status)}">
      <div class="print-card-top">
        <span class="print-badge">${printValue(normalized.status || "Planned")}</span>
        <span class="print-time-chip">${printValue(formatSupportDateTime(s4PrimaryTime(normalized)), "TBD")}</span>
      </div>
      <strong>${printValue(s4Title(normalized), "Support item")}</strong>
      <p>${s4Meta(normalized).map(escapeHtml).join(" / ")}</p>
      ${details.length ? `<p>${details.map(escapeHtml).join(" / ")}</p>` : ""}
    </article>
  `;
}

function printS4Page() {
  const total = state.s4Items.length;
  return `
    <section class="sheet s4-page">
      <header class="print-page-head">
        <div>
          <p class="print-eyebrow">Support Control</p>
          <h1>S4 Tracker</h1>
        </div>
        <span>${total} active support item${total === 1 ? "" : "s"}</span>
      </header>
      <div class="print-s4-grid">
        ${S4_LANES.map((lane) => {
          const laneItems = state.s4Items.filter((item) => normalizeS4Lane(item.type) === lane).sort((a, b) => (s4PrimaryTime(a) || "9999").localeCompare(s4PrimaryTime(b) || "9999"));
          return `
            <section class="print-s4-lane">
              <div class="print-section-head">
                <h2>${escapeHtml(lane)}</h2>
                <span>${laneItems.length}</span>
              </div>
              ${laneItems.length ? laneItems.map(printS4Item).join("") : `<div class="print-empty">Nothing tracked.</div>`}
            </section>
          `;
        }).join("")}
      </div>
    </section>
  `;
}

function buildPrintHtml(options = {}) {
  const autoPrint = Boolean(options.autoPrint);
  const iso = state.selectedDate;
  const dayEvents = eventsForDate(iso).sort(eventSort);
  const dayTasks = tasksForDate(iso);
  const trackLabels = activeTracksForDate(iso).map((track) => [track, classDayLabelForTrack(iso, track)].filter(Boolean).join(" - "));
  const activeTrackLabel = trackLabels.length ? trackLabels.join(" / ") : "Shared";

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>AASLT Print - ${escapeHtml(formatFullDate(iso))}</title>
    <style>
      @page { size: letter portrait; margin: 0.28in; }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        background: #e8eef0;
        color: #162326;
        font: 8.6px/1.18 Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }
      .print-now {
        position: fixed;
        right: 14px;
        top: 14px;
        z-index: 2;
        border: 1px solid #255f66;
        border-radius: 6px;
        background: #2e7076;
        color: #fff;
        padding: 7px 10px;
        font-weight: 800;
      }
      .sheet {
        width: min(100%, 8.5in);
        margin: 14px auto;
        padding: 0.2in;
        border: 1px solid #c7d3d6;
        background: #fff;
        box-shadow: 0 14px 35px rgba(20, 33, 37, 0.14);
      }
      .s4-page { break-before: page; page-break-before: always; }
      .print-page-head,
      .print-section-head,
      .print-card-top {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 10px;
      }
      .print-page-head {
        margin-bottom: 6px;
        padding-bottom: 5px;
        border-bottom: 1.5px solid #244f55;
      }
      .print-page-head h1 {
        margin: 0;
        font-size: 16px;
        line-height: 1;
      }
      .print-page-head span,
      .print-section-head span {
        color: #52666b;
        font-weight: 800;
        text-align: right;
      }
      .print-eyebrow {
        margin: 0;
        color: #66787d;
        font-size: 7px;
        font-weight: 900;
        letter-spacing: 0.07em;
        text-transform: uppercase;
      }
      .print-section {
        margin-top: 6px;
      }
      .print-section-head {
        margin-bottom: 4px;
        padding-bottom: 3px;
        border-bottom: 1px solid #c8d5d8;
      }
      h2 {
        margin: 0;
        font-size: 10.5px;
      }
      h3 {
        margin: 0 0 3px;
        font-size: 8.8px;
      }
      p {
        margin: 1px 0 0;
      }
      .print-stat,
      .print-meal,
      .print-card,
      .print-event-mini,
      .print-task-line,
      .print-empty {
        border: 1px solid #c8d5d8;
        border-radius: 4px;
        background: #f8fbfa;
      }
      .print-stat {
        padding: 4px 5px;
      }
      .print-stat span,
      .print-card p,
      .print-meal p,
      .print-event-mini p,
      .print-task-line small {
        color: #52666b;
      }
      .print-stat span {
        display: block;
        font-size: 6.8px;
        font-weight: 900;
        letter-spacing: 0.06em;
        text-transform: uppercase;
      }
      .print-stat strong {
        display: block;
        margin-top: 1px;
        font-size: 9px;
      }
      .print-meal-strip {
        display: grid;
        grid-template-columns: repeat(6, minmax(0, 1fr));
        gap: 3px;
      }
      .print-person-grid,
      .print-s4-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 5px;
      }
      .print-meal {
        min-width: 0;
        padding: 3px 4px;
      }
      .print-meal span {
        display: block;
        color: #244f55;
        font-size: 6.8px;
        font-weight: 950;
        letter-spacing: 0.05em;
        text-transform: uppercase;
      }
      .print-meal strong {
        display: block;
        margin-top: 1px;
        font-size: 8px;
      }
      .print-meal em {
        display: inline-block;
        color: #225b61;
        font-style: normal;
        font-weight: 900;
      }
      .print-timeline-matrix {
        display: grid;
        grid-template-columns: 0.42in repeat(var(--timeline-cols), minmax(0, 1fr));
        border-top: 1px solid #9fb2b7;
        border-left: 1px solid #c8d5d8;
      }
      .print-matrix-head,
      .print-matrix-time,
      .print-matrix-cell {
        min-width: 0;
        border-right: 1px solid #c8d5d8;
        border-bottom: 1px solid #c8d5d8;
      }
      .print-matrix-head {
        background: #edf4f4;
        color: #244f55;
        padding: 3px 4px;
        font-size: 7.4px;
        font-weight: 950;
        text-transform: uppercase;
      }
      .print-matrix-time {
        padding: 4px 3px;
        color: #225b61;
        font-size: 7.5px;
        font-weight: 950;
        text-align: center;
      }
      .print-matrix-cell {
        display: grid;
        gap: 2px;
        align-content: start;
        min-height: 0.18in;
        padding: 2px;
      }
      .print-event-mini {
        padding: 2px 4px;
        border-left-width: 3px;
        break-inside: avoid;
        page-break-inside: avoid;
      }
      .print-event-mini strong {
        display: block;
        font-size: 7.8px;
        line-height: 1.12;
      }
      .print-event-mini p {
        display: -webkit-box;
        overflow: hidden;
        -webkit-line-clamp: 2;
        -webkit-box-orient: vertical;
        font-size: 6.8px;
      }
      .print-person-lane,
      .print-s4-lane {
        border: 1px solid #c8d5d8;
        border-radius: 4px;
        padding: 4px;
        break-inside: avoid;
        page-break-inside: avoid;
      }
      .print-stack,
      .print-task-stack {
        display: grid;
        gap: 2px;
      }
      .print-card {
        padding: 4px 5px;
        break-inside: avoid;
        page-break-inside: avoid;
      }
      .print-card strong {
        display: block;
        margin-top: 2px;
        font-size: 8.6px;
      }
      .print-badge,
      .print-time-chip,
      .print-chips b {
        display: inline-flex;
        align-items: center;
        min-height: 13px;
        border-radius: 999px;
        padding: 1px 4px;
        font-size: 6.8px;
        font-weight: 900;
      }
      .print-badge { background: #e0eced; color: #225b61; }
      .print-time-chip { background: #fff; border: 1px solid #c8d5d8; color: #225b61; }
      .print-chips {
        display: inline-flex;
        flex-wrap: wrap;
        gap: 3px;
        justify-content: flex-end;
      }
      .print-chips b { background: #eef4f4; color: #53686d; }
      .print-training { border-left: 4px solid #667d34; }
      .print-instruction { border-left: 4px solid #386da3; }
      .print-medical { border-left: 4px solid #a3483e; }
      .print-meals { border-left: 4px solid #a7641b; }
      .print-cadre { border-left: 4px solid #6e628f; }
      .print-logistics,
      .print-operations { border-left: 4px solid #2e7076; }
      .print-task-list {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 3px;
      }
      .print-task-line {
        display: grid;
        grid-template-columns: 0.36in minmax(0, 1fr);
        gap: 1px 4px;
        padding: 2px 4px;
      }
      .print-task-line span {
        grid-row: span 2;
        color: #225b61;
        font-size: 7px;
        font-weight: 950;
      }
      .print-task-line strong {
        font-size: 7.7px;
      }
      .print-task-line small {
        font-size: 6.7px;
      }
      .print-empty {
        padding: 5px;
        color: #6a7a7f;
        font-weight: 700;
        text-align: center;
      }
      .print-s4-lane {
        min-height: 1.2in;
      }
      .print-s4-lane .print-card + .print-card {
        margin-top: 3px;
      }
      .status-done { opacity: 0.72; }
      @media print {
        body { background: #fff; }
        .print-now { display: none; }
        .sheet {
          width: auto;
          min-height: auto;
          margin: 0;
          padding: 0;
          border: 0;
          box-shadow: none;
        }
      }
    </style>
  </head>
  <body>
    <button class="print-now" type="button" onclick="window.print()">Print</button>
    <section class="sheet ops-page">
      <header class="print-page-head">
        <div>
          <h1>${escapeHtml(formatFullDate(iso))}</h1>
        </div>
        <span>Active Tracks: ${escapeHtml(activeTrackLabel)}</span>
      </header>

      <section class="print-section">
        <div class="print-section-head">
          <div>
            <p class="print-eyebrow">Meal Snapshot</p>
            <h2>CM / MRE</h2>
          </div>
        </div>
        ${printMealSummary(dayEvents)}
      </section>

      <section class="print-section">
        <div class="print-section-head">
          <div>
            <p class="print-eyebrow">Timeline</p>
            <h2>Selected Day Schedule</h2>
          </div>
          <span>${state.dayView === "all" ? "All Events view" : state.dayView === "table" ? "Table view" : "Tracks view"}</span>
        </div>
        ${printTimelineMatrix(dayEvents)}
      </section>

      ${printAssignments(dayTasks)}
    </section>

    ${printS4Page()}
    ${
      autoPrint
        ? `<script>
      window.addEventListener("load", () => {
        window.setTimeout(() => window.print(), 300);
      });
    <\/script>`
        : ""
    }
  </body>
</html>`;
}

function printSelectedDay() {
  const previewOnly = new URLSearchParams(window.location.search).has("printPreview");
  const blob = new Blob([buildPrintHtml({ autoPrint: !previewOnly })], { type: "text/html" });
  const url = URL.createObjectURL(blob);
  const printWindow = window.open(url, "_blank");
  if (!printWindow) {
    URL.revokeObjectURL(url);
    window.alert("Print window was blocked. Allow pop-ups for this site and try again.");
    return;
  }
  printWindow.focus();
  window.setTimeout(() => URL.revokeObjectURL(url), 60000);
}

function exportLocalState() {
  const payload = {
    exportedAt: new Date().toISOString(),
    selectedDate: state.selectedDate,
    tasks: state.tasks,
    deletedSource: state.deletedSource,
    eventOverrides: state.eventOverrides,
    s4Items: state.s4Items,
    receiptRegister: state.receipts.map(({ blob, ...meta }) => meta),
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = createElement("a", { attrs: { href: url, download: "aaslt-dashboard-state.json" } });
  a.click();
  URL.revokeObjectURL(url);
}

init().catch((error) => {
  console.error(error);
  document.body.replaceChildren(createElement("main", { className: "fatal", text: `Dashboard failed to load: ${error.message}` }));
});
