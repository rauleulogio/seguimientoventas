/* ============================================================
   SALES CONTROL - OCT 26
   Google Sheets -> Web Dashboard
============================================================ */

"use strict";

/* ============================================================
   CONFIGURACIÓN
============================================================ */

let CSV_URL = "";

const ROWS_PER_PAGE = 10;
const REFRESH_INTERVAL = 60000;
const SEARCH_DEBOUNCE = 200;
const FETCH_TIMEOUT = 20000;

const PROJECTION_USES_FILTERS = true;

/* ============================================================
   ACCESO (LOGIN) - Contraseñas directas en texto plano
============================================================ */

const USERS_CONFIG = {
    "elva": {
        pass: "elva2026",
        name: "Elva Edith Curo Salazar",
        csvUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vS2dS4SShe8mxhPXlWqkzaYjhLzztRdNPaSy8shwiCx8U_bO_EGp3W8dbSYJfrc5qUmXGYHX9x87Y7I/pub?output=csv"
    },
    "gianella": {
        pass: "gianella2026",
        name: "Gianella Mendiguete Navarro",
        csvUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vTpDEkiE8mYRUvwsYlzzcPpd54EQCYF-7yIVOd-D14hmGvUosSbzocGvd7jIsVxb9Seffj3aSZ_2UKQ/pub?output=csv"
    },
    "raul": {
        pass: "raul2026",
        name: "Raúl Hugo Eulogio Terranova",
        csvUrl: "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ-0sM4BhtW0jFQUkPZr4fkExixvpNEGe1GPisZTtFsJC-pBtpOePh0uigNq7iFAw/pub?output=csv"
    }
};

const SESSION_KEY = "seguimiento_ventas_sesion";
const MAX_LOGIN_ATTEMPTS = 5;
const LOGIN_LOCK_MS = 30000;

let currentUser = null;


/* ============================================================
   ESTADO GLOBAL
============================================================ */

let headers = [];
let sales = [];
let filteredSales = [];

let currentPage = 1;
let isLoading = false;
let refreshTimer = null;
let toastTimer = null;
let searchTimer = null;
let lastFocusedElement = null;

const columnCache = new Map();


/* ============================================================
   COLUMNAS Y TABLAS DE CONFIGURACIÓN
============================================================ */

const COLUMN_ALIASES = {
    n: ["N°", "Nº", "NO", "N", "NUMERO", "NÚMERO"],
    fechaVenta: ["FECHA VENTA"],
    horaVenta: ["HORA VENTA"],
    dni: ["DNI"],
    nombre: ["NOMBRE", "NOMBRE CLIENTE"],
    tipoDespacho: ["TIPO DESPAACHO", "TIPO DESPACHO", "TIPO DE DESPACHO"],
    rango: ["RANGO"],
    rangoEntrega: ["RANGO ENTREGA"],
    departamento: ["DEPARTAMENTO"],
    provincia: ["PROVINCIA"],
    distrito: ["DISTRITO"],
    tipoOfrec: ["TIPO OFREC.", "TIPO OFREC", "TIPO OFRECIMIENTO"],
    tipoVenta: ["TIPO VENTA"],
    lineaPortar: ["LINEA A PORTAR", "LÍNEA A PORTAR"],
    tipoProducto: ["TIPO DE PROD.", "TIPO DE PROD", "TIPO DE PRODUCTO"],
    supervisor: ["SUPERVISOR"],
    estadoBO: ["ESTADO BO"],
    subestado: ["SUBESTADO"],
    ordenOT: ["ORDEN OT"],
    atencionFinal: ["ATENCIÓN FINAL", "ATENCION FINAL"],
    boValidador: ["BO VALIDADOR"],
    dniVendedor: ["DNI VENDEDOR"],
    vendedor: ["VENDEDOR"],
    estado: ["ESTADO"],
    fechaActivacion: ["FECHA ACTIVACIÓN", "FECHA ACTIVACION"],
    modalidad: ["MODALIDAD"],
    planDescuento: ["PLAN CON DESC.", "PLAN CON DESC", "PLAN CON DESCUENTO"],
    clienteEntel: ["CLIENTE ENTEL"],
    status: ["STATUS"]
};

const REQUIRED_COLUMNS = ["n", "nombre", "dni", "estado"];

const MAIN_COLUMNS = [
    { key: "n", label: "N°" },
    { key: "fechaVenta", label: "FECHA VENTA" },
    { key: "dni", label: "DNI" },
    { key: "nombre", label: "NOMBRE" },
    { key: "rango", label: "RANGO" },
    { key: "distrito", label: "DISTRITO" },
    { key: "tipoOfrec", label: "TIPO OFREC." },
    { key: "lineaPortar", label: "LINEA A PORTAR" },
    { key: "estadoBO", label: "ESTADO BO" },
    { key: "ordenOT", label: "ORDEN OT", occurrence: 1 },
    { key: "estado", label: "ESTADO" },
    { key: "fechaActivacion", label: "FECHA ACTIVACIÓN" },
    { key: "modalidad", label: "MODALIDAD" }
];

const DATE_KEYS = ["fechaVenta", "fechaActivacion"];
const STATUS_KEYS = ["estado"];

const DETAIL_SECTIONS = [
    {
        title: "Cliente y venta",
        fields: [
            ["N°", "n"],
            ["Fecha de venta", "fechaVenta"],
            ["Hora de venta", "horaVenta"],
            ["DNI", "dni"],
            ["Nombre", "nombre"]
        ]
    },
    {
        title: "Entrega",
        fields: [
            ["Rango", "rango"],
            ["Rango entrega", "rangoEntrega"],
            ["Distrito", "distrito"],
            ["Provincia", "provincia"],
            ["Departamento", "departamento"],
            ["Tipo despacho", "tipoDespacho"]
        ]
    },
    {
        title: "Producto",
        fields: [
            ["Tipo ofrecimiento", "tipoOfrec"],
            ["Tipo de venta", "tipoVenta"],
            ["Línea a portar", "lineaPortar"],
            ["Tipo de producto", "tipoProducto"],
            ["Modalidad", "modalidad"],
            ["Plan con descuento", "planDescuento"],
            ["Cliente Entel", "clienteEntel"]
        ]
    },
    {
        title: "Gestión comercial",
        fields: [
            ["Supervisor", "supervisor"],
            ["Estado BO", "estadoBO"],
            ["Subestado", "subestado"],
            ["Orden OT", "ordenOT", 1],
            ["Atención final", "atencionFinal"],
            ["BO validador", "boValidador"]
        ]
    },
    {
        title: "Vendedor y activación",
        fields: [
            ["DNI vendedor", "dniVendedor"],
            ["Vendedor", "vendedor"],
            ["Estado", "estado"],
            ["Fecha activación", "fechaActivacion"],
            ["Status", "status"],
            ["Orden OT 2", "ordenOT", 2]
        ]
    }
];


/* ============================================================
   COMISIONES
============================================================ */

const COMMISSION_RANGES = [
    { range: 1,  pos: 5.00,  alta: 2.00,  prepago: 2.50,  multi: 8.00,  from: 20,  to: 30 },
    { range: 2,  pos: 12.50, alta: 4.00,  prepago: 6.00,  multi: 14.50, from: 31,  to: 40 },
    { range: 3,  pos: 13.50, alta: 4.50,  prepago: 6.50,  multi: 15.50, from: 41,  to: 50 },
    { range: 4,  pos: 14.50, alta: 5.00,  prepago: 7.00,  multi: 16.50, from: 51,  to: 60 },
    { range: 5,  pos: 15.50, alta: 5.50,  prepago: 7.50,  multi: 17.50, from: 61,  to: 70 },
    { range: 6,  pos: 17.00, alta: 6.00,  prepago: 8.50,  multi: 19.00, from: 71,  to: 80 },
    { range: 7,  pos: 19.00, alta: 6.00,  prepago: 10.00, multi: 22.00, from: 81,  to: 95 },
    { range: 8,  pos: 21.50, alta: 7.00,  prepago: 12.00, multi: 25.00, from: 96,  to: 110 },
    { range: 9,  pos: 24.00, alta: 8.00,  prepago: 14.00, multi: 27.00, from: 111, to: 150 },
    { range: 10, pos: 27.00, alta: 9.00,  prepago: 16.00, multi: 30.00, from: 151, to: 199 },
    { range: 11, pos: 30.00, alta: 12.00, prepago: 18.00, multi: 33.00, from: 200, to: Infinity }
];

const TOP_RANGE = COMMISSION_RANGES[COMMISSION_RANGES.length - 1];
const TARGET_SALES = TOP_RANGE.from;

const TABLEU_RULES = [
    { min: 80,        pct: 15,  label: "≥ 80%" },
    { min: 75,        pct: 5,   label: "≥ 75%" },
    { min: 70,        pct: -5,  label: "< 75%" },
    { min: -Infinity, pct: -15, label: "< 70%" }
];

const MULTI_RULES = [
    { min: 35,        pct: 0,   label: "≥ 35%" },
    { min: 30,        pct: -5,  label: "< 35%" },
    { min: -Infinity, pct: -10, label: "< 30%" }
];

function round2(value) {
    return Math.round(value * 100) / 100;
}

function findRule(rules, value) {
    return rules.find(rule => round2(value) >= rule.min);
}


/* ============================================================
   UTILIDADES
============================================================ */

const $ = id => document.getElementById(id);

function normalize(value) {
    return String(value ?? "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toUpperCase()
        .trim();
}

function normalizeHeader(value) {
    return normalize(value)
        .replace(/[°º]/g, "")
        .replace(/[.:;,()[\]{}]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function cleanValue(value) {
    return String(value ?? "").replace(/\r/g, "").trim();
}

function escapeHTML(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function setText(id, value) {
    const element = $(id);
    if (element) {
        element.textContent = value;
    }
}

function clamp(value, min = 0, max = 100) {
    return Math.max(min, Math.min(max, Number(value) || 0));
}

function setWidth(id, percentageValue) {
    const element = $(id);
    if (element) {
        element.style.width = `${clamp(percentageValue)}%`;
    }
}

function percentage(value, total) {
    return total ? (value / total) * 100 : 0;
}

function formatPercentage(value) {
    if (!Number.isFinite(value)) return "0%";
    return `${value.toFixed(2).replace(/\.00$/, "").replace(/(\.\d)0$/, "$1")}%`;
}

const moneyFormatter = new Intl.NumberFormat("es-PE", {
    style: "currency",
    currency: "PEN",
    minimumFractionDigits: 2
});

function formatMoney(value) {
    return moneyFormatter.format(value || 0);
}

function formatSignedMoney(value) {
    if (value > 0) return `+${formatMoney(value)}`;
    if (value < 0) return `−${formatMoney(Math.abs(value))}`;
    return formatMoney(0);
}

function formatDateValue(value) {
    const text = cleanValue(value);
    const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    return match ? `${match[3]}/${match[2]}/${match[1]}` : text;
}


/* ============================================================
   CSV
============================================================ */

function detectDelimiter(text) {
    const firstLine = text.split(/\r?\n/, 1)[0] || "";
    let comma = 0, semicolon = 0, insideQuotes = false;

    for (const char of firstLine) {
        if (char === '"') {
            insideQuotes = !insideQuotes;
        } else if (!insideQuotes) {
            if (char === ",") comma++;
            if (char === ";") semicolon++;
        }
    }
    return semicolon > comma ? ";" : ",";
}

function parseCSV(text, delimiter) {
    const rows = [];
    let row = [], value = "", insideQuotes = false;

    const pushRow = () => {
        row.push(value);
        if (row.some(cell => cleanValue(cell) !== "")) {
            rows.push(row);
        }
        row = [];
        value = "";
    };

    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        const next = text[i + 1];

        if (char === '"') {
            if (insideQuotes && next === '"') {
                value += '"';
                i++;
            } else {
                insideQuotes = !insideQuotes;
            }
        } else if (char === delimiter && !insideQuotes) {
            row.push(value);
            value = "";
        } else if ((char === "\n" || char === "\r") && !insideQuotes) {
            if (char === "\r" && next === "\n") i++;
            pushRow();
        } else {
            value += char;
        }
    }
    if (value !== "" || row.length > 0) pushRow();
    return rows;
}


/* ============================================================
   ACCESO A COLUMNAS
============================================================ */

function computeColumnIndex(aliases, occurrence) {
    const wanted = aliases.map(normalizeHeader);
    const normalizedHeaders = headers.map(normalizeHeader);
    const exact = [], partial = [];

    normalizedHeaders.forEach((header, index) => {
        if (!header) return;
        if (wanted.includes(header)) exact.push(index);
        const isPartial = wanted.some(alias =>
            alias.length >= 4 && header.length >= 4 &&
            (header.includes(alias) || alias.includes(header))
        );
        if (isPartial) partial.push(index);
    });

    if (occurrence === "last") {
        const list = exact.length ? exact : partial;
        return list.length ? list[list.length - 1] : -1;
    }

    if (exact.length >= occurrence) return exact[occurrence - 1];
    if (partial.length >= occurrence) return partial[occurrence - 1];
    return -1;
}

function findColumnIndex(aliases, occurrence = 1) {
    if (!headers.length) return -1;
    const cacheKey = `${aliases.join("|")}#${occurrence}`;
    if (!columnCache.has(cacheKey)) {
        columnCache.set(cacheKey, computeColumnIndex(aliases, occurrence));
    }
    return columnCache.get(cacheKey);
}

function resolveAliases(keyOrAliases) {
    return Array.isArray(keyOrAliases) ? keyOrAliases : (COLUMN_ALIASES[keyOrAliases] || [keyOrAliases]);
}

function getValue(row, keyOrAliases, occurrence = 1) {
    const index = findColumnIndex(resolveAliases(keyOrAliases), occurrence);
    return index < 0 || index >= row.length ? "" : cleanValue(row[index]);
}

function missingColumns() {
    return REQUIRED_COLUMNS.filter(key => findColumnIndex(resolveAliases(key)) < 0);
}


/* ============================================================
   CLASIFICACIÓN DE VENTAS
============================================================ */

function isRealSale(row) {
    const hasNumber = getValue(row, "n").replace(/\D/g, "") !== "";
    if (!hasNumber) return false;
    const hasName = normalize(getValue(row, "nombre")).length > 0;
    const hasDni = getValue(row, "dni").replace(/\D/g, "").length >= 6;
    return hasName || hasDni;
}

function classifyStatusText(text) {
    if (/OBSERV/.test(text)) return "observed";
    if (/\b(ACTIV|VALID)/.test(text) && !/\bNO (ACTIV|VALID)/.test(text)) return "active";
    if (/PROGRES|PENDIENT|EN PROCESO|EN CURSO/.test(text)) return "progress";
    return null;
}

function getOperationalStatus(row) {
    const estado = normalize(getValue(row, "estado"));
    if (/OBSERV/.test(estado)) return "observed";

    const combined = [
        estado,
        ...["estadoBO", "subestado", "atencionFinal", "status"].map(key => normalize(getValue(row, key)))
    ].join(" | ");

    if (/\b(CANCEL|RECHAZ)/.test(combined)) return "cancelled";
    if (/NO ?RECOGE/.test(combined)) return "noRecoge";

    return classifyStatusText(estado) || classifyStatusText(combined) || "progress";
}

function getOtCount(row) {
    const raw = getValue(row, "ordenOT", "last");
    if (!raw) return 0;
    const small = raw.match(/^\s*(\d{1,3})(?!\d)/);
    if (small) return parseInt(small[1], 10);
    const ids = raw.match(/\d{6,}/g);
    return ids ? ids.length : 0;
}

function getChannel(row) {
    const combined = normalize(`${getValue(row, "tipoDespacho")} ${getValue(row, "modalidad")}`);
    if (combined.includes("TIENDA") || combined.includes("STORE")) return "tienda";
    if (combined.includes("DELIVERY") || combined.includes("DOMICILIO")) return "delivery";
    return "other";
}

function getProductText(row) {
    return normalize([
        getValue(row, "tipoVenta"),
        getValue(row, "tipoProducto"),
        getValue(row, "tipoOfrec"),
        getValue(row, "modalidad"),
        getValue(row, "lineaPortar")
    ].join(" "));
}

function isMultipedido(row) {
    return /MULTI\s*LINEA/.test(normalize(getValue(row, "tipoOfrec")));
}

function getCommissionType(productText, isMulti) {
    if (isMulti) return "multi";
    if (productText.includes("PREPAGO")) return "prepago";
    if (/\bALTA\b|LLAA/.test(productText)) return "alta";
    return "pos";
}

function buildSale(row) {
    const productText = getProductText(row);
    const isMulti = isMultipedido(row);

    const searchText = [
        "n", "fechaVenta", "horaVenta", "dni", "nombre", "distrito",
        "tipoOfrec", "lineaPortar", "estado", "vendedor", "supervisor"
    ].map(key => getValue(row, key)).concat(getValue(row, "ordenOT", 1)).map(normalize).join(" ");

    const fields = {};
    FILTERS.forEach(filter => {
        const raw = getValue(row, filter.key, filter.occurrence || 1);
        fields[filter.key] = filter.type === "date" ? toISODate(raw) : raw;
    });

    return {
        row,
        status: getOperationalStatus(row),
        channel: getChannel(row),
        isMulti,
        commissionType: getCommissionType(productText, isMulti),
        otCount: getOtCount(row),
        fields,
        searchText
    };
}

function summarize(list) {
    const result = { total: list.length, active: 0, progress: 0, observed: 0, cancelledOnly: 0, noRecoge: 0 };
    for (const sale of list) {
        if (sale.status === "active") result.active++;
        else if (sale.status === "progress") result.progress++;
        else if (sale.status === "observed") result.observed++;
        else if (sale.status === "cancelled") result.cancelledOnly++;
        else if (sale.status === "noRecoge") result.noRecoge++;
    }
    result.cancelled = result.cancelledOnly + result.noRecoge;
    return result;
}


/* ============================================================
   CARGA DE DATOS
============================================================ */

function buildCsvUrl() {
    const url = new URL(CSV_URL);
    url.searchParams.set("t", Date.now());
    return url.toString();
}

async function cargarDatos(showToastMessage = false) {
    if (isLoading) return;
    isLoading = true;

    const refreshButton = $("refreshButton");
    if (refreshButton) refreshButton.disabled = true;

    setConnection("loading", "Actualizando...");

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), FETCH_TIMEOUT);

    try {
        const response = await fetch(buildCsvUrl(), {
            cache: "no-store",
            signal: controller.signal
        });

        if (!response.ok) throw new Error(`HTTP ${response.status}`);

        const text = (await response.text()).replace(/^\uFEFF/, "");
        if (text.trim().length < 10) throw new Error("Google Sheets no devolvió datos.");

        const parsed = parseCSV(text, detectDelimiter(text));
        if (parsed.length < 2) throw new Error("No se encontraron filas.");

        headers = parsed[0].map(cleanValue);
        columnCache.clear();

        sales = parsed.slice(1).filter(isRealSale).map(buildSale);

        cargarFiltros();
        aplicarFiltros();

        setConnection("connected", "Conectado");
        setText("lastUpdate", `Actualizado ${new Date().toLocaleString("es-PE", { dateStyle: "short", timeStyle: "short" })}`);

        if (showToastMessage) mostrarToast(`Datos actualizados: ${sales.length} ventas.`);

    } catch (error) {
        console.error("Error cargando Google Sheets:", error);
        setConnection("error", "Error de conexión");
        if (!sales.length) {
            $("tableEmpty").hidden = false;
        }
        mostrarToast("No se pudieron actualizar los datos.");
    } finally {
        clearTimeout(timeout);
        isLoading = false;
        if (refreshButton) refreshButton.disabled = false;
    }
}

function setConnection(type, text) {
    setText("connectionStatus", text);
    const dot = $("connectionDot");
    if (dot) dot.className = type;
}


/* ============================================================
   FILTROS
============================================================ */

const FILTERS = [
    { key: "fechaVenta", label: "Fecha de venta", type: "date" },
    { key: "dni", label: "DNI", type: "text" },
    { key: "rango", label: "Rango", type: "select" },
    { key: "distrito", label: "Distrito", type: "select" },
    { key: "tipoOfrec", label: "Tipo ofrecimiento", type: "select" },
    { key: "lineaPortar", label: "Línea a portar", type: "select" },
    { key: "estadoBO", label: "Estado BO", type: "select" },
    { key: "ordenOT", label: "Orden OT", type: "text", occurrence: 1 },
    { key: "estado", label: "Estado", type: "select" },
    { key: "fechaActivacion", label: "Fecha de activación", type: "date" },
    { key: "modalidad", label: "Modalidad", type: "select" }
];

const EMPTY_OPTION = "__EMPTY__";
const filterId = (filter, suffix = "") => `filter-${filter.key}${suffix}`;

function toISODate(value) {
    const text = cleanValue(value);
    const pad = number => String(number).padStart(2, "0");
    let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (match) return `${match[1]}-${pad(match[2])}-${pad(match[3])}`;
    match = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);
    if (match) {
        const year = match[3].length === 2 ? `20${match[3]}` : match[3];
        return `${year}-${pad(match[2])}-${pad(match[1])}`;
    }
    return "";
}

function createFilterField(filter) {
    const id = filterId(filter);
    const label = escapeHTML(filter.label);
    const field = document.createElement("div");

    field.className = filter.type === "date" ? "field field-date" : "field";

    if (filter.type === "date") {
        field.innerHTML = `
            <span class="field-label" id="${id}-label">${label}</span>
            <div class="date-range" role="group" aria-labelledby="${id}-label">
                <label>Desde <input type="date" id="${id}-from"></label>
                <label>Hasta <input type="date" id="${id}-to"></label>
            </div>
        `;
    } else if (filter.type === "text") {
        field.innerHTML = `
            <label for="${id}">${label}</label>
            <input type="text" id="${id}" autocomplete="off" placeholder="Contiene...">
        `;
    } else {
        field.innerHTML = `
            <label for="${id}">${label}</label>
            <select id="${id}"><option value="ALL">Todos</option></select>
        `;
    }
    return field;
}

function construirFiltros() {
    const container = $("advancedFilters");
    if (container) container.replaceChildren(...FILTERS.map(createFilterField));
}

function cargarFiltros() {
    FILTERS.filter(filter => filter.type === "select").forEach(filter => {
        const select = $(filterId(filter));
        if (!select) return;

        const current = select.value;
        const values = new Set();
        let hasEmpty = false;

        sales.forEach(sale => {
            const value = sale.fields[filter.key];
            if (value) values.add(value);
            else hasEmpty = true;
        });

        select.replaceChildren(new Option("Todos", "ALL"));
        [...values].sort((a, b) => a.localeCompare(b, "es", { numeric: true })).forEach(value => select.add(new Option(value, value)));
        if (hasEmpty) select.add(new Option("(Sin dato)", EMPTY_OPTION));

        select.value = current || "ALL";
    });
}

let appliedValues = {};

function leerControles() {
    const values = {};
    $("advancedFilters")?.querySelectorAll("input, select").forEach(control => {
        values[control.id] = control.value;
    });
    return values;
}

function escribirControles(values) {
    $("advancedFilters")?.querySelectorAll("input, select").forEach(control => {
        const isSelect = control.tagName === "SELECT";
        control.value = values[control.id] ?? (isSelect ? "ALL" : "");
    });
}

function leerFiltros() {
    const active = [];
    FILTERS.forEach(filter => {
        const id = filterId(filter);
        if (filter.type === "date") {
            const from = appliedValues[`${id}-from`] || "";
            const to = appliedValues[`${id}-to`] || "";
            if (from || to) {
                active.push({ filter, from, to, ids: [`${id}-from`, `${id}-to`], display: `${from} a ${to}` });
            }
        } else if (filter.type === "text") {
            const raw = cleanValue(appliedValues[id]);
            if (raw) active.push({ filter, text: normalize(raw), ids: [id], display: raw });
        } else {
            const value = appliedValues[id] || "ALL";
            if (value !== "ALL") active.push({ filter, value, ids: [id], display: value });
        }
    });
    return active;
}

function matchesFilter(sale, active) {
    const { filter } = active;
    const value = sale.fields[filter.key];
    if (filter.type === "date") return Boolean(value) && (!active.from || value >= active.from) && (!active.to || value <= active.to);
    if (filter.type === "text") return normalize(value).includes(active.text);
    return active.value === EMPTY_OPTION ? !value : value === active.value;
}

function renderFiltrosActivos(active) {
    const container = $("activeFilters");
    if (!container) return;

    container.replaceChildren(...active.map(item => {
        const chip = document.createElement("span");
        chip.className = "filter-chip";
        chip.innerHTML = `<span>${item.filter.label}: ${item.display}</span><button type="button" class="filter-chip-remove">×</button>`;
        chip.querySelector("button").addEventListener("click", () => {
            item.ids.forEach(id => { appliedValues[id] = "ALL"; });
            escribirControles(appliedValues);
            resetPageAndFilter();
        });
        return chip;
    }));
}

function limpiarFiltro(ids) {
    ids.forEach(id => {
        appliedValues[id] = $(id)?.tagName === "SELECT" ? "ALL" : "";
    });
    escribirControles(appliedValues);
    resetPageAndFilter();
}

function aplicarFiltros() {
    const search = normalize($("searchInput")?.value);
    const active = leerFiltros();

    filteredSales = sales.filter(sale =>
        (!search || sale.searchText.includes(search)) &&
        active.every(filter => matchesFilter(sale, filter))
    );

    renderFiltrosActivos(active);
    renderAll();
}

function resetPageAndFilter() {
    currentPage = 1;
    aplicarFiltros();
}


/* ============================================================
   RENDER Y TABLA
============================================================ */

function renderAll() {
    actualizarKPIs();
    renderTabla();
    renderAnalytics();
    calculateChannelKPIs(filteredSales);
    calculateProjection(PROJECTION_USES_FILTERS ? filteredSales : sales);
}

function actualizarKPIs() {
    const stats = summarize(filteredSales);
    setText("kpiTotal", stats.total);
    setText("kpiActive", stats.active);
    setText("kpiProgress", stats.progress);
    setText("kpiObserved", stats.observed);
    setText("kpiCancelled", stats.cancelled);
}

function renderStatusBadge(value) {
    const text = cleanValue(value);
    if (!text) return `<span class="status-badge status-neutral">—</span>`;
    const normalized = normalize(text);
    let className = "status-neutral";

    if (/\b(CANCEL|RECHAZ)/.test(normalized)) className = "status-cancelled";
    else if (/\b(ACTIV|VALID)/.test(normalized)) className = "status-active";
    else if (/OBSERV/.test(normalized)) className = "status-observed";
    else if (/PROGRES|PENDIENT/.test(normalized)) className = "status-progress";

    return `<span class="status-badge ${className}">${escapeHTML(text)}</span>`;
}

function renderTableHead() {
    const head = $("tableHead");
    if (!head) return;
    head.innerHTML = `<tr>${MAIN_COLUMNS.map(col => `<th>${col.label}</th>`).join("")}<th>ACCIÓN</th></tr>`;
}

function renderTabla() {
    const body = $("tableBody");
    const empty = $("tableEmpty");
    if (!body) return;

    renderTableHead();
    const totalPages = Math.max(1, Math.ceil(filteredSales.length / ROWS_PER_PAGE));
    currentPage = clamp(currentPage, 1, totalPages);

    const start = (currentPage - 1) * ROWS_PER_PAGE;
    const pageRows = filteredSales.slice(start, start + ROWS_PER_PAGE);

    setText("paginationInfo", `${filteredSales.length} registros`);
    body.replaceChildren();

    if (!pageRows.length) {
        if (empty) empty.hidden = false;
        actualizarPaginacion(0, 0);
        return;
    }
    if (empty) empty.hidden = true;

    const fragment = document.createDocumentFragment();
    pageRows.forEach(sale => {
        const tr = document.createElement("tr");
        MAIN_COLUMNS.forEach(column => {
            const td = document.createElement("td");
            let value = getValue(sale.row, column.key, column.occurrence || 1);
            if (DATE_KEYS.includes(column.key)) value = formatDateValue(value);

            if (STATUS_KEYS.includes(column.key)) {
                td.innerHTML = renderStatusBadge(value);
            } else {
                td.textContent = value || "—";
            }
            tr.appendChild(td);
        });

        const actionTd = document.createElement("td");
        const btn = document.createElement("button");
        btn.className = "action-button";
        btn.textContent = "Ver detalle";
        btn.addEventListener("click", () => mostrarDetalle(sale));
        actionTd.appendChild(btn);
        tr.appendChild(actionTd);

        fragment.appendChild(tr);
    });

    body.appendChild(fragment);
    actualizarPaginacion(currentPage, totalPages);
}

function actualizarPaginacion(page, totalPages) {
    const pagination = $("pagination");
    if (!pagination) return;
    pagination.replaceChildren();
    if (totalPages <= 1) return;

    for (let i = 1; i <= totalPages; i++) {
        const button = document.createElement("button");
        button.className = `page-button ${i === page ? "active" : ""}`;
        button.textContent = i;
        button.addEventListener("click", () => {
            currentPage = i;
            renderTabla();
        });
        pagination.appendChild(button);
    }
}


/* ============================================================
   ANALYTICS Y KPIs DE CANAL
============================================================ */

function renderAnalytics() {
    const stats = summarize(filteredSales);
    setText("analyticsTotal", stats.total);

    const rows = [
        ["Active", stats.active],
        ["Progress", stats.progress],
        ["Observed", stats.observed],
        ["Cancelled", stats.cancelled]
    ];

    rows.forEach(([name, count]) => {
        const pct = percentage(count, stats.total);
        setText(`analytics${name}Pct`, formatPercentage(pct));
        setWidth(`analyticsBar${name}`, pct);
    });
}

function setPie(id, percentageValue, centerValue, label) {
    const element = $(id);
    if (!element) return;
    element.style.setProperty("--value", clamp(percentageValue));

    const value = element.querySelector(".pie-value");
    const labelElement = element.querySelector(".pie-label");
    if (value) value.textContent = centerValue;
    if (labelElement) labelElement.textContent = label;
}

function setLine(prefix, count, total) {
    setText(`${prefix}Q`, count);
    setText(`${prefix}Pct`, formatPercentage(percentage(count, total)));
}

function calculateChannelKPIs(currentSales) {
    const tiendaSales = currentSales.filter(s => s.channel === "tienda");
    const tiendaStats = summarize(tiendaSales);
    setLine("tiendaActivas", tiendaStats.active, tiendaStats.total);
    setLine("tiendaNoRecoge", tiendaStats.cancelled, tiendaStats.total);
    setLine("tiendaProgreso", tiendaStats.progress, tiendaStats.total);
    setText("tiendaEfectividadQ", tiendaStats.total);
    setText("tiendaEfectividadPct", formatPercentage(percentage(tiendaStats.active, tiendaStats.total)));
    setPie("chartTienda", percentage(tiendaStats.active, tiendaStats.total), formatPercentage(percentage(tiendaStats.active, tiendaStats.total)), "ACTIVAS");
    setPie("chartTiendaTotal", tiendaStats.total > 0 ? 100 : 0, tiendaStats.total, "TOTAL");

    const multiSales = currentSales.filter(s => s.isMulti);
    const multiStats = summarize(multiSales);
    setLine("multiActivas", multiStats.active, multiStats.total);
    setLine("multiCanceladas", multiStats.cancelled, multiStats.total);
    setLine("multiProgreso", multiStats.progress, multiStats.total);
    setText("multiTotalQ", multiStats.total);
    setText("multiTotalPct", multiStats.total ? "100%" : "0%");
    setPie("chartMulti", percentage(multiStats.active, multiStats.total), formatPercentage(percentage(multiStats.active, multiStats.total)), "ACTIVAS");
    setPie("chartMultiTotal", multiStats.total > 0 ? 100 : 0, multiStats.total, "TOTAL");

    const delSales = currentSales.filter(s => s.channel === "delivery");
    const delStats = summarize(delSales);
    setLine("deliveryActivas", delStats.active, delStats.total);
    setLine("deliveryCanceladas", delStats.cancelled, delStats.total);
    setLine("deliveryProgreso", delStats.progress, delStats.total);
    setText("deliveryTotalQ", delStats.total);
    setText("deliveryEfectividadPct", formatPercentage(percentage(delStats.active, delStats.total)));
    setPie("chartDelivery", percentage(delStats.active, delStats.total), formatPercentage(percentage(delStats.active, delStats.total)), "ACTIVAS");
    setPie("chartDeliveryTotal", delStats.total > 0 ? 100 : 0, delStats.total, "TOTAL");

    let tableuTotal = 0, tableuActive = 0, tableuCancel = 0, tableuProgress = 0;
    currentSales.forEach(s => {
        const orders = s.otCount;
        if (!orders) return;
        tableuTotal += orders;
        if (s.status === "active") tableuActive += orders;
        else if (s.status === "progress") tableuProgress += orders;
        else tableuCancel += orders;
    });

    setLine("tableuActivas", tableuActive, tableuTotal);
    setLine("tableuCanceladas", tableuCancel, tableuTotal);
    setLine("tableuProgreso", tableuProgress, tableuTotal);
    setText("tableuTotalQ", tableuTotal);
    setText("tableuTotalPct", tableuTotal ? "100%" : "0%");
    const tableuAlcance = percentage(tableuActive, tableuTotal);
    setPie("chartTableu", tableuAlcance, formatPercentage(tableuAlcance), "ACTIVAS");
    setPie("chartTableuTotal", tableuTotal > 0 ? 100 : 0, tableuTotal, "TOTAL OT");
}


/* ============================================================
   PROYECCIÓN DE COMISIONES
============================================================ */

function calculateProjection(currentSales) {
    const activeSales = currentSales.filter(sale => sale.status === "active");
    const activeCount = activeSales.length;

    const currentRange = COMMISSION_RANGES.find(r => activeCount >= r.from && activeCount <= r.to);
    const pricingRange = currentRange || COMMISSION_RANGES[0];

    const totals = { pos: 0, alta: 0, prepago: 0, multi: 0 };
    const counts = { pos: 0, alta: 0, prepago: 0, multi: 0 };

    activeSales.forEach(sale => {
        totals[sale.commissionType] += pricingRange[sale.commissionType];
        counts[sale.commissionType]++;
    });

    const baseCommission = totals.pos + totals.alta + totals.prepago + totals.multi;

    setText("projectionActiveSales", activeCount);
    setText("projectionCurrentRange", currentRange ? `Rango ${currentRange.range}` : "Aún no alcanzado");
    setText("projectionMissingRange", Math.max(TARGET_SALES - activeCount, 0));
    setText("projectionTotal", formatMoney(baseCommission));

    setText("projectionPos", formatMoney(totals.pos));
    setText("projectionAlta", formatMoney(totals.alta));
    setText("projectionPrepago", formatMoney(totals.prepago));
    setText("projectionMulti", formatMoney(totals.multi));

    setText("projectionPosCount", `${counts.pos} ventas`);
    setText("projectionAltaCount", `${counts.alta} ventas`);
    setText("projectionPrepagoCount", `${counts.prepago} ventas`);
    setText("projectionMultiCount", `${counts.multi} ventas`);

    setWidth("rangeProgressBar", Math.min(100, percentage(activeCount, TARGET_SALES)));
    setText("rangeProgressCurrent", `${activeCount} / ${TARGET_SALES}`);

    renderCommissionTable(currentRange, activeCount);
}

function renderCommissionTable(currentRange, activeCount) {
    const tbody = $("commissionTableBody");
    if (!tbody) return;

    tbody.innerHTML = COMMISSION_RANGES.map(range => {
        let rowClass = "";
        if (currentRange && range.range === currentRange.range) {
            rowClass = "current-range";
        }
        return `
            <tr class="${rowClass}">
                <td>Rango ${range.range}</td>
                <td>${formatMoney(range.pos)}</td>
                <td>${formatMoney(range.alta)}</td>
                <td>${formatMoney(range.prepago)}</td>
                <td>${formatMoney(range.multi)}</td>
                <td>${range.from}</td>
                <td>${range.to === Infinity ? "∞" : range.to}</td>
            </tr>
        `;
    }).join("");
}


/* ============================================================
   DETALLE Y TOAST
============================================================ */

function mostrarDetalle(sale) {
    const body = $("detailBody");
    const overlay = $("detailOverlay");
    if (!sale || !body || !overlay) return;

    const row = sale.row;
    setText("detailTitle", getValue(row, "nombre") || "Venta sin nombre");
    setText("detailSubtitle", `Venta #${getValue(row, "n") || "—"} · DNI ${getValue(row, "dni") || "—"}`);

    body.innerHTML = DETAIL_SECTIONS.map(section => {
        const fieldsHTML = section.fields.map(([label, key, occurrence = 1]) => {
            let value = getValue(row, key, occurrence);
            if (DATE_KEYS.includes(key)) value = formatDateValue(value);
            return `
                <div class="detail-item">
                    <span class="detail-label">${escapeHTML(label)}</span>
                    <span class="detail-value">${value ? escapeHTML(value) : "—"}</span>
                </div>
            `;
        }).join("");

        return `
            <section class="detail-section">
                <h4 class="detail-section-title">${escapeHTML(section.title)}</h4>
                <div class="detail-grid">${fieldsHTML}</div>
            </section>
        `;
    }).join("");

    overlay.hidden = false;
}

function cerrarDetalle() {
    const overlay = $("detailOverlay");
    if (overlay) overlay.hidden = true;
}

function mostrarToast(message) {
    const toast = $("toast");
    if (!toast) return;
    toast.textContent = message;
    toast.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { toast.hidden = true; }, 4000);
}


/* ============================================================
   AUTENTICACIÓN Y LOGIN (Texto plano directo)
============================================================ */

function findUser(id) {
    return Object.prototype.hasOwnProperty.call(USERS_CONFIG, id)
        ? { id, ...USERS_CONFIG[id] }
        : null;
}

function validarCredenciales(usuario, clave) {
    const user = findUser(String(usuario).trim().toLowerCase());
    if (!user || user.pass !== clave) {
        return null;
    }
    return user;
}

function configurarLogin() {
    const loginForm = $("loginForm");
    const loginOverlay = $("loginScreen");
    const logoutButton = $("logoutButton");

    const savedUser = sessionStorage.getItem(SESSION_KEY);
    if (savedUser && USERS_CONFIG[savedUser]) {
        const user = findUser(savedUser);
        if (user) {
            CSV_URL = user.csvUrl;
            setText("userName", user.name);
            if (loginOverlay) loginOverlay.hidden = true;
            $("appShell").hidden = false;
            iniciarDashboard();
        }
    }

    loginForm?.addEventListener("submit", event => {
        event.preventDefault();
        const userInput = $("loginUser").value;
        const passInput = $("loginPass").value;
        const errorBox = $("loginError");

        const user = validarCredenciales(userInput, passInput);

        if (user) {
            sessionStorage.setItem(SESSION_KEY, user.id);
            CSV_URL = user.csvUrl;
            setText("userName", user.name);
            if (loginOverlay) loginOverlay.hidden = true;
            $("appShell").hidden = false;
            if (errorBox) errorBox.hidden = true;
            iniciarDashboard();
        } else if (errorBox) {
            errorBox.textContent = "Usuario o contraseña incorrectos.";
            errorBox.hidden = false;
        }
    });

    logoutButton?.addEventListener("click", () => {
        sessionStorage.removeItem(SESSION_KEY);
        window.location.reload();
    });

    $("togglePass")?.addEventListener("click", () => {
        const input = $("loginPass");
        const show = input.type === "password";
        input.type = show ? "text" : "password";
        $("togglePass").textContent = show ? "Ocultar" : "Mostrar";
    });
}


/* ============================================================
   EVENTOS E INICIALIZACIÓN
============================================================ */

function configurarEventos() {
    $("searchInput")?.addEventListener("input", () => {
        clearTimeout(searchTimer);
        searchTimer = setTimeout(resetPageAndFilter, SEARCH_DEBOUNCE);
    });

    $("refreshButton")?.addEventListener("click", () => cargarDatos(true));
    $("closeDetail")?.addEventListener("click", cerrarDetalle);
    $("detailOverlay")?.addEventListener("click", event => {
        if (event.target === event.currentTarget) cerrarDetalle();
    });
}

function iniciarDashboard() {
    construirFiltros();
    appliedValues = leerControles();
    configurarEventos();
    cargarDatos();
}

document.addEventListener("DOMContentLoaded", () => {
    configurarLogin();
});
