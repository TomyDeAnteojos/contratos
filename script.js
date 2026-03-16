const downloadButton = document.getElementById("download-button");
const statusEl = document.getElementById("status");
const previewSelector = document.getElementById("preview-selector");
const previewTableBody = document.querySelector("#preview-table tbody");
const fileInput = document.getElementById("excel-input");

const PT = {
  obra: null,
  servicio: null,
};

const DAY_MS = 86400 * 1000;
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const DATE_FIELDS = new Set(["INICIO", "FIN"]);

function resolveDate(value) {
  if (value == null) return null;
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = new Date(EXCEL_EPOCH + Math.round(value * DAY_MS));
    return isNaN(date.getTime()) ? null : date;
  }
  const text = value.toString().trim();
  if (!text) return null;
  const parts = text.split(/[\/\.\-]/).map((part) => part.trim()).filter(Boolean);
  if (parts.length < 3) return null;
  const [dayPart, monthPart, yearPart] = parts;
  const day = Number(dayPart);
  const month = Number(monthPart);
  const year = Number(yearPart);
  if ([day, month, year].some((num) => Number.isNaN(num))) return null;
  const date = new Date(Date.UTC(year, month - 1, day));
  return isNaN(date.getTime()) ? null : date;
}

function toDateParts(date) {
  return {
    day: String(date.getUTCDate()).padStart(2, "0"),
    month: String(date.getUTCMonth() + 1).padStart(2, "0"),
    year: String(date.getUTCFullYear()),
    monthIndex: date.getUTCMonth() + 1,
    yearNum: date.getUTCFullYear(),
  };
}

function parseDate(value) {
  const date = resolveDate(value);
  return date ? toDateParts(date) : null;
}

function formatDate(value) {
  const date = resolveDate(value);
  if (!date) return null;
  const { day, month, year } = toDateParts(date);
  return `${day}/${month}/${year}`;
}

function formatFieldValue(column, value) {
  if (DATE_FIELDS.has(column)) {
    const formatted = formatDate(value);
    if (formatted) {
      return formatted;
    }
  }
  if (value == null) {
    return "";
  }
  return value.toString().trim();
}

function replacePlaceholdersInXml(xmlString, replacements) {
  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  const xmlDoc = parser.parseFromString(xmlString, "application/xml");
  const paragraphs = Array.from(xmlDoc.getElementsByTagName("*")).filter((node) => node.localName === "p");
  const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";

  const setWordText = (node, text) => {
    node.textContent = text;
    if (/^\s|\s$/.test(text)) {
      node.setAttributeNS(XML_NAMESPACE, "xml:space", "preserve");
    } else {
      node.removeAttributeNS(XML_NAMESPACE, "space");
    }
  };

  const applyReplacement = (segments, start, end, value) => {
    let inserted = false;
    for (const segment of segments) {
      const segStart = segment.start;
      const segEnd = segment.start + segment.length;
      if (segEnd <= start || segStart >= end) continue;
      const from = Math.max(start, segStart) - segStart;
      const to = Math.min(end, segEnd) - segStart;
      const current = segment.node.textContent || "";
      if (!inserted) {
        const before = current.slice(0, from);
        const after = current.slice(to);
        setWordText(segment.node, before + value + after);
        inserted = true;
      } else {
        setWordText(segment.node, current.slice(0, from) + current.slice(to));
      }
    }
  };

  paragraphs.forEach((paragraph) => {
    const textNodes = Array.from(paragraph.getElementsByTagName("*")).filter((node) => node.localName === "t");
    let flatText = "";
    const segments = [];

    textNodes.forEach((node) => {
      const value = node.textContent || "";
      segments.push({
        node,
        start: flatText.length,
        length: value.length,
      });
      flatText += value;
    });
    const occupied = [];
    const placeholders = Object.keys(replacements).sort((a, b) => b.length - a.length);
    const matches = [];

    placeholders.forEach((placeholder) => {
      let searchFrom = 0;
      while (searchFrom < flatText.length) {
        const matchIndex = flatText.indexOf(placeholder, searchFrom);
        if (matchIndex === -1) break;
        const matchEnd = matchIndex + placeholder.length;
        const overlaps = occupied.some(([start, end]) => !(matchEnd <= start || matchIndex >= end));
        if (!overlaps) {
          matches.push({
            start: matchIndex,
            end: matchEnd,
            value: replacements[placeholder].toString(),
          });
          occupied.push([matchIndex, matchEnd]);
        }
        searchFrom = matchIndex + 1;
      }
    });

    matches
      .sort((a, b) => b.start - a.start)
      .forEach((match) => {
        applyReplacement(segments, match.start, match.end, match.value);
      });
  });

  return serializer.serializeToString(xmlDoc);
}

const previewOrder = [
  "%apellido",
  "%nombre",
  "%nombre_completo",
  "%TRABAJO",
  "%domicilio",
  "%mail",
  "%telefono",
  "%cuit",
  "%dni",
  "%forma_pago",
  "%tipo_contrato",
  "%sueldo",
  "%total",
  "%duracion",
  "%inicio",
  "%fin",
  "%idia",
  "%imes",
  "%ianio",
  "%fdia",
  "%fmes",
  "%fanio",
  "%objetivo",
];

const state = {
  rows: [],
};

let templatesReady = false;
let readyPromise = null;

async function bootstrap() {
  statusEl.textContent = "Cargando plantillas...";
  readyPromise = loadTemplates();
  try {
    await readyPromise;
    templatesReady = true;
    statusEl.textContent = "Listo para cargar Excel.";
  } catch (error) {
    statusEl.textContent = "No se pudieron cargar las plantillas de Word.";
    console.error(error);
  }
}

async function loadTemplates() {
  const keys = Object.keys(PT);
  await Promise.all(
    keys.map(async (key) => {
      const response = await fetch(`${key}.docx`);
      if (!response.ok) {
        throw new Error(`No se pudo obtener ${key}.docx`);
      }
      PT[key] = await response.arrayBuffer();
    })
  );
}

function setStatus(message, variant = "normal") {
  statusEl.textContent = message;
  statusEl.dataset.status = variant;
}

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  if (!templatesReady) {
    setStatus("Cargando plantillas, espera un momento...");
    await readyPromise;
  }
  setStatus("Procesando Excel...");
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      throw new Error("El archivo no tiene hojas.");
    }
    const sheet = workbook.Sheets[firstSheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    const normalized = rawRows.map(normalizeRow);
    const built = normalized.map(buildRowModel).filter(Boolean);
    state.rows = built;
    if (!built.length) {
      setStatus("No se detectaron filas con datos.");
      updatePreviewOptions();
      refreshPreview();
      downloadButton.disabled = true;
      return;
    }
    setStatus(`Se cargaron ${built.length} registros. Listo para descargar.`);
    downloadButton.disabled = false;
    updatePreviewOptions();
    refreshPreview(0);
  } catch (error) {
    console.error(error);
    setStatus("No se pudo leer el archivo. Asegurate de que el Excel tenga las columnas correctas.");
    downloadButton.disabled = true;
    state.rows = [];
    updatePreviewOptions();
    refreshPreview();
  }
});

function normalizeRow(row) {
  const normalized = {};
  for (const [key, value] of Object.entries(row)) {
    const cleanKey = key
      .toString()
      .replace(/\u00A0/g, " ")
      .trim()
      .toUpperCase()
      .replace(/[^\w]+/g, "_")
      .replace(/^_+|_+$/g, "");
    if (!cleanKey) continue;
    normalized[cleanKey] = typeof value === "string" ? value.trim() : value;
  }
  return normalized;
}

function expandReplacements(replacements) {
  const expanded = {};
  for (const [key, value] of Object.entries(replacements)) {
    expanded[key] = value;
    expanded[key.toLowerCase()] = value;
    expanded[key.toUpperCase()] = value;
  }
  return expanded;
}

function buildRowModel(row) {
  const sexo = (row.SEXO || "").toString().trim().toUpperCase();
  const isFemale = sexo === "F";
  const folder = (row.TRABAJO || "").toString().trim().toLowerCase();
  const template = folder === "obra" ? "obra" : "servicio";
  const inicioDate = parseDate(row.INICIO);
  const finDate = parseDate(row.FIN);
  const sueldoRaw = formatFieldValue("PESOS", row.PESOS);
  const sueldoNumeric = Number(sueldoRaw.replace(/[^0-9.-]/g, "")) || 0;
  const monthsBetween = Math.max(0, monthDifference(inicioDate, finDate));
  const durationDays = Math.max(0, dayDifference(inicioDate, finDate));
  const totalValue = sueldoNumeric * monthsBetween;
  const totalFormatted = formatFinancial(totalValue);
  const replacements = {
    "%el": isFemale ? "la" : "el",
    "%sr": isFemale ? "Sra" : "Sr",
    "%EL": isFemale ? "LA" : "EL",
    "%al": isFemale ? "a la" : "al",
    "%TRABAJO": isFemale ? "PRESTADORA" : "PRESTADOR",
    "%apellido": formatFieldValue("APELLIDO", row.APELLIDO),
    "%nombre": formatFieldValue("NOMBRE", row.NOMBRE),
    "%dni": formatFieldValue("DNI", row.DNI),
    "%cuit": formatFieldValue("CUIT", row.CUIT),
    "%calle": formatFieldValue("CALLE", row.CALLE),
    "%localidad": formatFieldValue("LOCALIDAD", row.LOCALIDAD),
    "%mail": formatFieldValue("MAIL", row.MAIL),
    "%telefono": formatFieldValue("TELEFONO", row.TELEFONO),
    "%PESOS": sueldoRaw,
    "%sueldo": sueldoRaw,
    "%total": totalFormatted,
    "%duracion": durationDays ? `${durationDays}` : "",
    "%forma_pago": formatFieldValue("FORMA_PAGO", row.FORMA_PAGO),
    "%tipo_contrato": formatFieldValue("TIPO_CONTRATO", row.TIPO_CONTRATO),
    "%objetivo": formatFieldValue("OBJETIVO", row.OBJETIVO),
    "%inicio": formatFieldValue("INICIO", row.INICIO),
    "%fin": formatFieldValue("FIN", row.FIN),
    "%nombre_completo": [formatFieldValue("APELLIDO", row.APELLIDO), formatFieldValue("NOMBRE", row.NOMBRE)]
      .filter(Boolean)
      .join(", "),
    "%domicilio": [formatFieldValue("CALLE", row.CALLE), formatFieldValue("LOCALIDAD", row.LOCALIDAD)]
      .filter(Boolean)
      .join(", "),
  };

  if (inicioDate) {
    replacements["%idia"] = inicioDate.day;
    replacements["%imes"] = inicioDate.month;
    replacements["%ianio"] = inicioDate.year;
  }
  if (finDate) {
    replacements["%fdia"] = finDate.day;
    replacements["%fmes"] = finDate.month;
    replacements["%fanio"] = finDate.year;
  }

  Object.entries(row).forEach(([key, value]) => {
    const normalizedKey = key.toString().trim().toLowerCase().replace(/\s+/g, "_");
    if (!normalizedKey) return;
    const placeholder = `%${normalizedKey}`;
    const hasEquivalentPlaceholder = Object.keys(replacements).some(
      (existingKey) => existingKey.toLowerCase() === placeholder.toLowerCase()
    );
    if (!hasEquivalentPlaceholder) {
      replacements[placeholder] = formatFieldValue(key, value);
    }
  });

  return {
    template,
    replacements: expandReplacements(replacements),
    label: `${replacements["%apellido"]} ${replacements["%nombre"]}`.trim() || `Fila ${Math.random().toString(36).slice(2, 6)}`,
  };
}

function monthDifference(start, end) {
  if (!start || !end) return 0;
  const totalEnd = end.yearNum * 12 + end.monthIndex;
  const totalStart = start.yearNum * 12 + start.monthIndex;
  return Math.max(0, totalEnd - totalStart + 1);
}

function dayDifference(start, end) {
  if (!start || !end) return 0;
  const startDate = Date.UTC(start.yearNum, start.monthIndex - 1, Number(start.day));
  const endDate = Date.UTC(end.yearNum, end.monthIndex - 1, Number(end.day));
  return Math.max(0, Math.floor((endDate - startDate) / DAY_MS) + 1);
}

function formatFinancial(value) {
  if (!Number.isFinite(value)) return "0";
  return value % 1 === 0 ? `${value}` : value.toFixed(2);
}

function updatePreviewOptions() {
  if (!state.rows.length) {
    previewSelector.innerHTML = "<option>Sin registros</option>";
    previewSelector.disabled = true;
    return;
  }
  previewSelector.disabled = false;
  previewSelector.innerHTML = state.rows
    .map((row, index) => `<option value="${index}">${row.label || `Registro ${index + 1}`}</option>`)
    .join("");
}

function refreshPreview(index = 0) {
  const row = state.rows[index];
  if (!row) {
    previewTableBody.innerHTML = `<tr><td colspan="2" class="muted">Carga un Excel para ver valores.</td></tr>`;
    return;
  }
  const rowsHtml = previewOrder
    .map((key) => {
      if (!(key in row.replacements)) return "";
      const value = row.replacements[key] || "";
      const label = key.replace("%", "");
      return `<tr><td>${label}</td><td>${value}</td></tr>`;
    })
    .filter(Boolean)
    .join("");
  previewTableBody.innerHTML = rowsHtml || `<tr><td colspan="2" class="muted">No hay valores disponibles.</td></tr>`;
}

previewSelector.addEventListener("change", (event) => {
  const index = Number(event.target.value);
  refreshPreview(index);
});

downloadButton.addEventListener("click", async () => {
  if (!state.rows.length) return;
  downloadButton.disabled = true;
  setStatus("Creando contratos...", "busy");
  try {
    const zip = new JSZip();
    for (const row of state.rows) {
      const templateKey = row.template === "obra" ? "obra" : "servicio";
      const buffer = PT[templateKey];
      if (!buffer) throw new Error(`Falta plantilla ${templateKey}`);
      const contractBlob = await fillTemplate(buffer, row.replacements);
      const fileName = `${row.replacements["%apellido"]}_${row.replacements["%nombre"]}`.replace(/\s+/g, "_") || `contrato-${Math.random().toString(36).slice(2, 6)}`;
      zip.file(`${fileName}.docx`, contractBlob);
    }
    const finalBlob = await zip.generateAsync({ type: "blob" });
    saveAs(finalBlob, "contratos.zip");
    setStatus("Archivo ZIP descargado." , "success");
  } catch (error) {
    console.error(error);
    setStatus("Hubo un error al generar los contratos.");
  } finally {
    downloadButton.disabled = false;
  }
});

async function fillTemplate(buffer, replacements) {
  const zip = await JSZip.loadAsync(buffer);
  const tasks = [];
  zip.forEach((relativePath, file) => {
    if (file.dir) return;
    if (!relativePath.endsWith(".xml") && !relativePath.endsWith(".rels")) return;
    tasks.push(
      file.async("string").then((content) => {
        if (relativePath.endsWith(".xml")) {
          zip.file(relativePath, replacePlaceholdersInXml(content, replacements));
          return;
        }
        let modified = content;
        const safeReplacements = {};
        Object.entries(replacements).forEach(([key, value]) => {
          safeReplacements[key] = escapeXml(value);
        });
        Object.entries(safeReplacements).forEach(([key, value]) => {
          modified = modified.split(key).join(value);
        });
        zip.file(relativePath, modified);
      })
    );
  });
  await Promise.all(tasks);
  return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

function escapeXml(value) {
  const text = value == null ? "" : value;
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

updatePreviewOptions();
refreshPreview();
bootstrap();
