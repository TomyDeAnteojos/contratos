/* Generador de Contratos (Word) + Vista previa
   Plantilla: template.docx con tags &APELLIDO&, &NOMBRE&, etc.
*/
const $ = (s) => document.querySelector(s);

const state = {
  templateArrayBuffer: null,
  extractedData: null,
  ocrText: "",
  paragraphs: ["Entre la UNIVERSIDAD NACIONAL DE PILAR, representada en este acto por su Rectora, Lic. Elizabeth Diana Wanger, DNI 18.287.351, en adelante LA \"COMITENTE\", por una parte, y el/la Sr/a &APELLIDO&, &NOMBRE&, DNI &DNI&, con domicilio en la calle &CALLE& &NUMERO& de la Localidad de &LOCALIDAD&, en adelante el/la \"PRESTADOR/A DE SERVICIOS\" y en forma conjunta las \"PARTES’’, convienen en celebrar el presente Contrato de Locación de Servicios Profesionales, en adelante el \"Contrato\", el que quedará encuadrado bajo las disposiciones del Código Civil y Comercial de la Nación y estará sujeto a las siguientes términos y condiciones:", "PRIMERA - OBJETO: LA COMITENTE encomienda al/la PRESTADOR/A DE SERVICIOS y este/a acepta en calidad de trabajador/a autónomo/a, la realización de las actividades que se hallan descriptas en la especificaciones técnicas, que como anexo forman parte del presente contrato, para la prestación de un servicio de asesoramiento técnico / formulación de proyectos / servicio de consultoría para proyectos de acreditación de carreras / para la complementación de actividades esenciales de enseñanza y educación / servicio de apoyo estudiantil / otros.", "SEGUNDA – ANTECEDENTES PROFESIONALES: EL/LA PRESTADOR/A DE SERVICIOS es un/a destacado/a profesional en su rubro. Además, posee la formación, experiencia y herramientas técnicas necesarias para la realización de las tareas y/o actividades que comprenden la prestación del servicio.", "TERCERA - SUBCONTRATOS: Como regla general, EL/LA PRESTADOR/A DE SERVICIOS deberá cumplir con sus obligaciones contractuales de manera directa debido a que resulta elegido por sus cualidades para realizarlo personalmente. En el supuesto de requerir una subcontratación, notificará fehacientemente a LA COMITENTE y con la debida antelación. La subcontratación no dará derecho al/la PRESTADOR/A DE SERVICIOS a solicitar modificaciones en el precio del contrato, conforme lo dispuesto en la cláusula QUINTA. De cualquier manera, el/la PRESTADOR/A DE SERVICIOS conservará la dirección y responsabilidad de la ejecución.", "EL/LA PRESTADOR/A DE SERVICIOS preservará y protegerá los derechos de LA COMITENTE respecto de la ejecución bajo subcontratos que pudiera firmar con terceros y deberá:", "Asumir responsabilidad solidaria ante LA COMITENTE para todas las obligaciones y responsabilidades que pudieran originarse de los actos y omisiones generados por su subcontratado y empleados.", "Hacer cumplir a sus subcontratados y/o empleados toda la normativa imperante, reglamentos emitidos por LA COMITENTE y cualquier normativa aplicable.", "EL/LA PRESTADOR/A DE SERVICIOS se constituirá como el único responsable ante sus subcontratados y/o empleados, y los contratos que celebre con éstos nunca le serán oponibles a LA COMITENTE, ni lo podrán alcanzar o afectar, siendo de exclusiva responsabilidad del/la PRESTADOR/A DE SERVICIOS.", "CUARTA – COMPROMISOS DEL PRESTADOR DE SERVICIOS: EL/LA PRESTADOR/A DE SERVICIOS manifiesta: que posee CUIT N°&CUIT& y que se compromete a realizar el servicio para el cual es contratado con total profesionalidad, actuando dentro de las prescripciones éticas y legales que hacen a su disciplina;", "Asimismo, manifiesta que se hará cargo bajo su exclusiva responsabilidad de sus  aportes previsionales y se comprometerá al estricto cumplimiento de los deberes y obligaciones derivadas de las aplicaciones de la legislación vigente, con especial atención a las reglamentaciones de seguridad e higiene;", "Que responderá por el cumplimiento de todas las leyes, ordenanzas, reglamentos y demás disposiciones nacionales y la reglamentación universitaria en materia de locación de servicios profesionales;", "Que conoce la misión y objetivos principales del/la COMITENTE, su estructura académica, sus autoridades y que no encuentra objeción para la ejecución de los servicios por los que se lo ha contratado, así como tampoco ninguna circunstancia que de algún modo impida su avance.", "QUINTA - PRECIO: Se determina un valor total del servicio hasta su total culminación de PESOS ($0.000.000,00) incluyendo IVA -de corresponder-, en adelante el “Precio”, que será abonado por LA COMITENTE al/la PRESTADOR/A DE SERVICIOS, no admitiéndose costos adicionales para financiar tareas o materiales excluidos de los alcances del presente Contrato. EL/LA PRESTADOR/A DE SERVICIOS, que declara haber estudiado y analizado lo requerido, pondrá a disposición los recursos y medios necesarios para la correcta realización del servicio contratado, de acuerdo con los términos y condiciones aquí pactadas, aun cuando ellos no hayan sido expresamente detallados en el presente y su documentación complementaria. Estos elementos que, sin ser mencionados, pudieran ser necesarios para dar cumplimiento al objeto integrante del presente Contrato, no darán derecho al/la PRESTADOR/A DE SERVICIOS a requerir modificaciones en el precio del Contrato.", "SEXTA – MODALIDAD DE PAGO: El precio total y absoluto del servicio contratado deberá pagarse a nombre de &NOMBRE&, &APELLIDO&, CUIT N°&CUIT& y de la siguiente manera: PAGOS SEMANALES / QUINCENALES / MENSUALES / BIMESTRALES / que El COMITENTE abonará a razón de PESOS ($000.000,00) hasta alcanzar la suma de PESOS ($0.000.000,00) con fecha límite del (día) de (mes) de 202x.", "SEPTIMA - PLAZOS DE EJECUCIÓN: Se estipula un período para la ejecución de las obligaciones contractuales de XXXX días corridos, contados desde el (día) de (mes) de 202x hasta el (día) de (mes) de 202x; prorrogables únicamente por estipulación expresa contractual.", "El COMITENTE procederá a la cancelación total de la contraprestación una vez canceladas todas las obligaciones asumidas por EL/LA PRESTADOR/A DE SERVICIOS.", "OCTAVA - RESCISIÓN DEL CONTRATO: Sin perjuicio de lo establecido en la normativa nacional vigente, en caso de incumplimiento de las condiciones pactadas en el presente Contrato por cualquiera de las partes, la otra podrá optar por rescindir el mismo intimando previamente a la contraparte al cumplimiento de la prestación debida, en forma fehaciente con cinco (5) días de anticipación. En caso de persistir el incumplimiento, el Contrato quedará rescindido de pleno derecho, pudiendo la parte cumplidora reclamar por los daños y perjuicios ocasionados.", "Asimismo, LA COMITENTE podrá, en cualquier etapa y sin notificación previa, rescindir unilateralmente y de manera anticipada el presente contrato, en uso de las atribuciones conferidas por el artículo 11 y 12, inciso a) del Decreto N°1023/2001 del RÉGIMEN DE CONTRATACIONES DE LA ADMINISTRACIÓN NACIONAL.", "NOVENA - CLAUSULA DE INDEMNIDAD: El/la PRESTADOR/A DE SERVICIOS deberá indemnizar y mantener informada a LA COMITENTE de cualquier reclamo, acción legal u otros procedimientos realizados por terceras partes que fueran atribuibles a cualquier acto u omisión que surja de la realización del contrato. Se conviene que dicha indemnización y deber de información se aplicará a los reclamos presentados por Entidades Gubernamentales Federales, Provinciales o Municipales. Asimismo, tal protección se aplicará a todos los reclamos que resulten aún después de prestado el servicio, como así también los presentados por sus empleados o subcontratistas, autoridades fiscales, impositivas o de la seguridad social respecto a la ejecución de tareas que estén bajo la órbita de su realización.", "DÉCIMA - CONFIDENCIALIDAD: EL/LA PRESTADOR/A DE SERVICIOS deberá, de conformidad con lo dispuesto en la Ley N°24.766, guardar estricta confidencialidad sobre las informaciones no publicadas, o de carácter reservado o confidencial contenidas en documentación, informes, expedientes, sistemas informáticos y cualquier otro medio en posesión del COMITENTE, con motivo de la ejecución de las obligaciones emanadas del presente Contrato, salvo autorización expresa de LA COMITENTE. Esta obligación de confidencialidad permanecerá en vigencia aún después de su cumplimiento, rescisión o resolución, siendo responsable EL/LA PRESTADOR/A DE SERVICIOS de los daños y perjuicios que pudiera ocasionar la difusión indebida.", "DÉCIMA PRIMERA - PROPIEDAD INTELECTUAL: Los derechos de propiedad de autor y de reproducción, así como cualquier otro derecho intelectual de cualquier naturaleza sobre informes, estudios, etc. producidos como consecuencia del cumplimiento de las obligaciones contractuales, pertenecerán exclusivamente a LA COMITENTE.", "DÉCIMA SEGUNDA - DERECHO DE INSPECCIÓN: LA COMITENTE puede inspeccionar los servicios prestados a través de algún representante o por sí, todas las veces que lo crea necesario, con motivo de constatar el fiel cumplimiento de todas las obligaciones que EL/LA PRESTADOR/A DE SERVICIOS asume.", "DÉCIMA TERCERA - ADENDAS AL CONTRATO: En el supuesto de que surjan propuestas de adendas que LAS PARTES deseen efectuar sobre las cláusulas que rigen la presente contratación, deberán comunicarse en forma previa y fehaciente para su análisis y la realización del procedimiento administrativo requerido para su eventual aprobación.", "DECIMA CUARTA - DOMICILIOS DE PAGO: Se determina como domicilio de pago expreso de las sumas reconocidas y estipuladas en la cláusula CUARTA los denunciados por LAS PARTES en el encabezamiento del presente contrato, o en su defecto los denunciados en las Especificaciones Técnicas que integran el Anexo referido anteriormente.", "DÉCIMA QUINTA - COMPETENCIA JUDICIAL: En caso de controversia judicial o prejudicial producto de la aplicación y/o interpretación de alguna de las cláusulas que rigen la presente contratación, LAS PARTES se someten a la Jurisdicción de los Juzgados Federales de Campana.", "En prueba de conformidad, en la ciudad de Pilar, Provincia de Buenos Aires y a los …… días del mes de ………….. de 20.., se suscriben dos ejemplares de un mismo tenor y a un solo efecto.", "Anexo", "ESPECIFICACIONES TÉCNICAS", "PRESTADOR DE SERVICIOS:: &APELLIDO&, &NOMBRE &", "CUIT: &CUIT&", "DOMICILIO: &DOMICILIO&", "TELÉFONO: &TELEFONO&", "E MAIL: &EMAIL&", "DURACIÓN:", "FECHA DE INICIO: 00/00/2025", "FECHA DE FINALIZACIÓN: 00/00/2025", "OBJETIVO GENERAL DE LA CONTRATACIÓN", "__________________________________________________________________________________________________________________________________________", "REQUERIMIENTOS Y CONDICIONES DE TRABAJO", "“El PRESTADOR DE SERVICIOS” deberá realizar tareas encomendadas en el período establecido en la cláusula SÉPTIMA.", "Las tareas por realizar por “El PRESTADOR DE SERVICIOS” podrán sufrir modificaciones por indicación de “LA COMITENTE”, para ser adecuada a las variaciones que puedan experimentar el desarrollo de los objetivos para los que fue contratada y el mejor logro de estos.", "FIRMA DEL PRESTADOR DE SERVICIOS: \t\t\tFIRMA COMITENTE", "ACLARACIÓN: \t\t\t\t\t\t\tSELLO", "DNI:"],
  title: "CONTRATO DE LOCACIÓN DE SERVICIOS PROFESIONALES"
};

const ui = {
  templateStatus: $("#templateStatus"),
  pdfFiles: $("#pdfFiles"),
  ocrStatus: $("#ocrStatus"),
  ocrProgress: $("#ocrProgress"),
  ocrText: $("#ocrText"),
  btnGenerate: $("#btnGenerate"),
  btnClear: $("#btnClear"),
  preview: $("#preview"),
  btnTogglePreview: $("#btnTogglePreview"),
  progressWrap: document.querySelector(".progress"),
  dropzone: $("#dropzone"),
};

function setStatus(el, msg) {
  if (!el) return;
  el.textContent = msg || "";
}
function setMain(msg) {
  setStatus(ui.ocrStatus, msg);
}
function setOcrText(text) {
  state.ocrText = text || "";
  if (ui.ocrText) ui.ocrText.value = state.ocrText;
}
function setProgress(value) {
  if (!ui.ocrProgress) return;
  const clamped = Math.max(0, Math.min(100, value));
  ui.ocrProgress.style.width = `${clamped}%`;
  if (ui.progressWrap) {
    ui.progressWrap.classList.toggle("active", clamped > 0 && clamped < 100);
  }
}

function setProcessing(isProcessing) {
  if (ui.progressWrap) {
    ui.progressWrap.classList.toggle("active", isProcessing);
  }
}

function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function sanitizeFilename(s) {
  return String(s || "")
    .normalize("NFKD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9._-]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

function emptyData() {
  return {
    apellido: "",
    nombre: "",
    dni: "",
    calle: "",
    numero: "",
    localidad: "",
    cuit: "",
    genero: "",
    domicilio: "",
    telefono: "",
    email: "",
  };
}

function getData() {
  return state.extractedData || emptyData();
}

function setData(data) {
  state.extractedData = data;
  if (ui.btnGenerate) ui.btnGenerate.disabled = false;
}

function resetExtractedData() {
  state.extractedData = emptyData();
  state.ocrText = "";
  if (ui.btnGenerate) ui.btnGenerate.disabled = true;
}

function buildDomicilio(calle, numero, localidad) {
  return [calle, numero, localidad].filter(Boolean).join(" ").trim();
}

function onlyDigits(value) {
  return String(value || "").replace(/\D+/g, "");
}

function normalizeLine(line) {
  return line.replace(/\s+/g, " ").trim();
}

function findAfterLabel(lines, labels) {
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    for (const label of labels) {
      const idx = line.indexOf(label);
      if (idx !== -1) {
        const raw = line.slice(idx + label.length).replace(/[:\-]/g, " ").trim();
        if (raw) return raw;
        if (i + 1 < lines.length) {
          return lines[i + 1].replace(/[:\-]/g, " ").trim();
        }
      }
    }
  }
  return "";
}

function cleanNameValue(value) {
  return String(value || "")
    .replace(/^[/\s]+/, "")
    .replace(/\b(APELLIDO[S]?|NOMBRE[S]?|SURNAME|NAME)\b/gi, "")
    .replace(/[:\-]/g, " ")
    .replace(/^[^A-ZÁÉÍÓÚÜÑ]+/gi, "")
    .replace(/[^\p{L}\s]+$/giu, "")
    .replace(/\bL+\b/gi, "")
    .replace(/\bK+\b/gi, "")
    .replace(/\s+/g, " ")
    .trim();
}

function findNameAfterLabel(lines, labels) {
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    for (const label of labels) {
      const idx = line.indexOf(label);
      if (idx !== -1) {
        const raw = line.slice(idx + label.length).replace(/[:\-]/g, " ").trim();
        const cleaned = cleanNameValue(raw);
        if (cleaned && !isBadNameValue(cleaned) && isLikelyName(cleaned)) return cleaned;
        if (i + 1 < lines.length) {
          const next = cleanNameValue(lines[i + 1]);
          if (next && !isBadNameValue(next) && isLikelyName(next)) return next;
        }
      }
    }
  }
  return "";
}

function pickFirstMatch(text, regex) {
  const match = text.match(regex);
  return match ? match[1] : "";
}

function pickEmail(text) {
  const clean = String(text || "").replace(/\s+@\s+/g, "@");
  const match = clean.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return match ? match[0] : "";
}

function parseMrz(text) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const line1 = lines.find((line) => line.startsWith("IDARG"));
  if (!line1) return {};
  const dniMatch = line1.match(/IDARG(\d{7,9})/);

  const candidate =
    lines.find(
      (line) =>
        !line.startsWith("IDARG") &&
        /K/.test(line) &&
        !/\d/.test(line) &&
        line.length >= 10
    ) ||
    lines.find(
      (line) =>
        !line.startsWith("IDARG") &&
        /[A-Z]/.test(line) &&
        !/\d/.test(line)
    );

  let apellido = "";
  let nombre = "";
  if (candidate) {
    let normalized = candidate.replace(/[^A-Z<]+/g, "");
    normalized = normalized.replace(/K+/g, "<");
    const parts = normalized.split(/<+/);
    const a = (parts[0] || "").replace(/</g, " ").trim();
    const n = parts.slice(1).join(" ").replace(/</g, " ").replace(/\s+/g, " ").trim();
    if (a && a.length >= 3 && !a.startsWith("REPUBLICA") && !isBadNameValue(a)) apellido = a;
    if (n && n.length >= 2 && !n.startsWith("REPUBLICA") && !isBadNameValue(n)) {
      nombre = cleanNameValue(n);
    }
  }
  return { apellido, nombre, dni: dniMatch ? dniMatch[1] : "" };
}

function extractDniFromText(text, cuitDigits) {
  const matches = String(text || "").match(/\b\d{7,9}\b/g) || [];
  for (const match of matches) {
    if (cuitDigits && cuitDigits.includes(match)) continue;
    return match;
  }
  return "";
}

function findValueBelowLabel(upperLines, rawLines, labels, skipWords) {
  for (let i = 0; i < upperLines.length; i++) {
    const line = upperLines[i];
    for (const label of labels) {
      if (line.includes(label)) {
        for (let j = i + 1; j < upperLines.length; j++) {
          const candidateUpper = upperLines[j];
          const candidateRaw = rawLines[j];
          if (!candidateUpper) continue;
          if (skipWords && skipWords.some((w) => candidateUpper.includes(w))) {
            continue;
          }
          return normalizeLine(candidateRaw);
        }
      }
    }
  }
  return "";
}

function parseGenero(value) {
  const upper = String(value || "").toUpperCase();
  const match = upper.match(/\b([MF])\b/);
  return match ? match[1] : "";
}

function isLikelyName(value) {
  const s = String(value || "").trim();
  if (!s) return false;
  if (/^\d+$/.test(s)) return false;
  if (s.length < 2) return false;
  return true;
}

function isBadNameValue(value) {
  const upper = String(value || "").toUpperCase();
  const bad = [
    "FECHA",
    "DATE",
    "ISSUE",
    "EXPIRY",
    "VENCIMIENTO",
    "NACIONALIDAD",
    "NATIONALITY",
    "SEXO",
    "DOCUMENTO",
    "DOCUMENT",
    "DOMICILIO",
    "LOCALIDAD",
    "PROVINCIA",
    "ARGENTINA",
  ];
  if (bad.some((w) => upper.includes(w))) return true;
  if (upper.length <= 3 && (upper === "ARA" || upper === "ARG")) return true;
  return false;
}

function normalizeCalle(value) {
  let s = String(value || "").toUpperCase();
  s = s.replace(/-+/g, " ").replace(/\s+/g, " ").trim();
  if (!s.includes(" ") && s.startsWith("EL") && s.length > 2) {
    s = `EL ${s.slice(2)}`.trim();
  }
  return s;
}

function parseAddressFromText(rawLines) {
  for (const line of rawLines) {
    const upper = line.toUpperCase().replace(/\s+/g, " ").replace(/\.+/g, ".").trim();
    if (
      upper.includes("IDARG") ||
      upper.includes("DOCUMENTO") ||
      upper.includes("DOCUMENT") ||
      upper.includes("TRAMITE") ||
      upper.includes("NACIONALIDAD")
    ) {
      continue;
    }
    const match = upper.match(/([A-Z\- ]+)\s+(\d{1,5})[.,]\s*([A-Z ]+)/);
    if (match) {
      let calle = normalizeCalle(match[1]);
      calle = calle.replace(/\bOUSNER\b/gi, "").replace(/\s+/g, " ").trim();
      const numero = match[2].trim();
      const localidad = match[3].trim();
      return { calle, numero, localidad };
    }
  }
  return { calle: "", numero: "", localidad: "" };
}

function parseOcrText(text) {
  const rawText = String(text || "");
  const rawLines = rawText.split(/\r?\n/).map(normalizeLine);
  const upperLines = rawLines.map((line) => line.toUpperCase());
  const lines = upperLines.filter(Boolean);

  const mrz = parseMrz(rawText.toUpperCase());
  let apellido =
    findNameAfterLabel(lines, ["APELLIDO", "APELLIDOS", "SURNAME"]) || "";
  let nombre =
    findNameAfterLabel(lines, ["NOMBRE", "NOMBRES", "NAME"]) || "";
  const apellidoBelow = findValueBelowLabel(
    upperLines,
    rawLines,
    ["APELLIDO", "SURNAME"],
    ["APELLIDO", "SURNAME", "NOMBRE", "NAME", "FECHA", "DATE", "ISSUE", "EXPIRY", "SEXO", "NACIONALIDAD", "ARGENTINA"]
  );
  const nombreBelow = findValueBelowLabel(
    upperLines,
    rawLines,
    ["NOMBRE", "NAME"],
    ["APELLIDO", "SURNAME", "NOMBRE", "NAME", "FECHA", "DATE", "ISSUE", "EXPIRY", "SEXO", "NACIONALIDAD", "ARGENTINA"]
  );
  if (apellidoBelow) {
    const cleaned = cleanNameValue(apellidoBelow);
    if (!isBadNameValue(cleaned)) apellido = cleaned;
  }
  if (nombreBelow) {
    const cleaned = cleanNameValue(nombreBelow);
    if (!isBadNameValue(cleaned)) nombre = cleaned;
  }
  if (mrz.apellido) apellido = mrz.apellido;
  if (mrz.nombre) nombre = mrz.nombre;
  if (!apellido && nombre) {
    const swap = cleanNameValue(nombre);
    apellido = swap;
    nombre = "";
  }
  if (!isLikelyName(apellido) || isBadNameValue(apellido)) apellido = "";
  if (!isLikelyName(nombre) || isBadNameValue(nombre)) nombre = "";
  const nameTokens = `${apellido} ${nombre}`.toUpperCase().split(/\s+/).filter(Boolean);

  const upper = rawText.toUpperCase();
  const cuitRaw = pickFirstMatch(upper, /CUIT\s*[:\-]?\s*([0-9.\-\s]{11,13})/);
  const cuit = onlyDigits(cuitRaw);

  const documentoLine = findValueBelowLabel(
    upperLines,
    rawLines,
    ["DOCUMENTO", "DOCUMENT"],
    ["DOCUMENTO", "DOCUMENT", "NOMBRE", "NAME", "APELLIDO", "SURNAME"]
  );
  const dniRaw =
    pickFirstMatch(upper, /DNI\s*[:\-]?\s*([0-9.\s]{7,10})/) ||
    pickFirstMatch(upper, /NRO\s*[:\-]?\s*([0-9.\s]{7,10})/) ||
    pickFirstMatch(upper, /NUMERO\s*[:\-]?\s*([0-9.\s]{7,10})/);
  const dniAlt =
    dniRaw ||
    pickFirstMatch(upper, /DOCUMENTO\s*[:\-]?\s*([0-9.\s]{7,10})/) ||
    pickFirstMatch(upper, /DOC(?:UMENTO)?\s*[:\-]?\s*([0-9.\s]{7,10})/);
  const docDigits = onlyDigits(documentoLine);
  const dni = docDigits || mrz.dni || onlyDigits(dniAlt) || extractDniFromText(upper, cuit);
  if (dni && apellido && onlyDigits(apellido) === dni) apellido = "";
  if (dni && nombre && onlyDigits(nombre) === dni) nombre = "";

  const sexoLine = findValueBelowLabel(
    upperLines,
    rawLines,
    ["SEXO", "SEX"],
    ["SEXO", "SEX"]
  );
  const sexo =
    pickFirstMatch(upper, /SEXO\s*[:\-]?\s*([MF])/i) ||
    pickFirstMatch(upper, /SEXO\s*[:\-]?\s*([MF])\b/i) ||
    parseGenero(sexoLine);
  const genero = sexo ? sexo.trim() : "";

  let domicilioLine = findValueBelowLabel(
    upperLines,
    rawLines,
    ["DOMICILIO", "ADDRESS"],
    ["DOMICILIO", "ADDRESS", "NOMBRE", "NAME", "APELLIDO", "SURNAME"]
  ) || pickFirstMatch(upper, /DOMICILIO\s*[:\-]?\s*([^\n]+)/);
  if (domicilioLine && nameTokens.some((t) => domicilioLine.toUpperCase().includes(t))) {
    domicilioLine = "";
  }
  let localidad = findValueBelowLabel(
    upperLines,
    rawLines,
    ["LOCALIDAD", "LOCALITY", "MUNICIPIO"],
    ["LOCALIDAD", "LOCALITY", "MUNICIPIO"]
  ) || pickFirstMatch(upper, /LOCALIDAD\s*[:\-]?\s*([^\n]+)/);
  localidad = String(localidad || "").replace(/PROVINCIA.*$/i, "").trim();

  let calle = "";
  let numero = "";
  if (domicilioLine) {
    const domicilioClean = domicilioLine.replace(/\s+\bLOCALIDAD\b.*$/i, "").trim();
    const match = domicilioClean.match(/^(.*?)(?:\s+(\d{1,6}[A-Z]?))$/);
    if (match) {
      calle = normalizeCalle(match[1]);
      numero = match[2].trim();
    } else {
      calle = normalizeCalle(domicilioClean);
    }
  }
  const calleLine = findValueBelowLabel(
    upperLines,
    rawLines,
    ["CALLE", "STREET"],
    ["CALLE", "STREET", "NOMBRE", "NAME", "APELLIDO", "SURNAME"]
  );
  if (calleLine && !nameTokens.some((t) => calleLine.toUpperCase().includes(t))) {
    calle = normalizeCalle(calleLine);
  }
  const numeroLine = findValueBelowLabel(
    upperLines,
    rawLines,
    ["NUMERO", "NRO", "NO", "NUM"],
    ["NUMERO", "NRO", "NO", "NUM", "NOMBRE", "NAME", "APELLIDO", "SURNAME"]
  );
  if (numeroLine) {
    const numMatch = numeroLine.match(/\d{1,6}[A-Z]?/);
    if (numMatch) numero = numMatch[0];
  }

  if (!calle && domicilioLine) {
    const parts = domicilioLine.split(/\s+/);
    if (parts.length >= 2) {
      const last = parts[parts.length - 1];
      if (/^\d{1,6}[A-Z]?$/.test(last)) {
        numero = numero || last;
        calle = parts.slice(0, -1).join(" ");
      }
    }
  }

  const email = pickEmail(rawText) || "";
  const telefono = pickFirstMatch(rawText, /(\+?\d[\d\s().-]{6,}\d)/) || "";

  if (!calle || !numero || !localidad) {
    const parsed = parseAddressFromText(rawLines);
    if (!calle) calle = parsed.calle;
    if (!numero) numero = parsed.numero;
    if (!localidad) localidad = parsed.localidad;
  }

  const domicilio = buildDomicilio(calle, numero, localidad);

  return {
    apellido,
    nombre,
    dni,
    calle,
    numero,
    localidad,
    cuit,
    genero,
    domicilio,
    telefono: telefono.trim(),
    email: email.trim(),
  };
}

const OCR_MAX_PAGES = 2;
const OCR_SCALE_DNI = 3.0;
const OCR_SCALE_DEFAULT = 1.4;
const OCR_PSM = 6;

function preprocessCanvas(canvas, { strong } = {}) {
  const ctx = canvas.getContext("2d");
  const { width, height } = canvas;
  const img = ctx.getImageData(0, 0, width, height);
  const data = img.data;
  const contrast = strong ? 1.35 : 1.15;
  const intercept = 128 * (1 - contrast);
  for (let i = 0; i < data.length; i += 4) {
    const r = data[i];
    const g = data[i + 1];
    const b = data[i + 2];
    let v = 0.2126 * r + 0.7152 * g + 0.0722 * b;
    v = v * contrast + intercept;
    if (strong) {
      v = v > 140 ? 255 : 0;
    }
    v = Math.max(0, Math.min(255, v));
    data[i] = v;
    data[i + 1] = v;
    data[i + 2] = v;
  }
  ctx.putImageData(img, 0, 0);
}

async function recognizeWithConfig(canvas, lang, config, onProgress) {
  const result = await window.Tesseract.recognize(canvas, lang, {
    logger: (m) => onProgress?.(m),
    ...config,
  });
  return result.data?.text || "";
}

async function extractTextFromPdf(file, onProgress, { isDni } = {}) {
  console.log("[OCR] Abriendo PDF:", file.name, file.size);
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  console.log("[OCR] Paginas:", pdf.numPages, "PDF:", file.name);
  const effectiveIsDni = isDni || pdf.numPages >= 2;
  let text = "";

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const content = await page.getTextContent();
    const pageText = content.items.map((item) => item.str).join(" ").trim();
    if (pageText.length >= 20) {
      console.log("[OCR] Texto extraido (PDF text layer):", file.name, "pag", pageNum);
      text += "\n" + pageText;
      if (!effectiveIsDni) continue;
    }
    if (pageNum > OCR_MAX_PAGES) {
      console.log("[OCR] Saltando OCR por limite de paginas:", file.name, "pag", pageNum);
      continue;
    }

    console.log("[OCR] OCR en pagina:", file.name, "pag", pageNum);
    const viewport = page.getViewport({ scale: effectiveIsDni ? OCR_SCALE_DNI : OCR_SCALE_DEFAULT });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    context.imageSmoothingEnabled = false;
    canvas.width = viewport.width;
    canvas.height = viewport.height;

    await page.render({ canvasContext: context, viewport }).promise;
    preprocessCanvas(canvas, { strong: effectiveIsDni });
    const lang = effectiveIsDni ? "spa+eng" : "spa";
    let ocrText = await recognizeWithConfig(
      canvas,
      lang,
      { tessedit_pageseg_mode: OCR_PSM },
      onProgress
    );
    if (effectiveIsDni) {
      const digitsText = await recognizeWithConfig(
        canvas,
        "eng",
        {
          tessedit_pageseg_mode: OCR_PSM,
          tessedit_char_whitelist: "0123456789",
        },
        onProgress
      );
      const mrzText = await recognizeWithConfig(
        canvas,
        "eng",
        {
          tessedit_pageseg_mode: OCR_PSM,
          tessedit_char_whitelist: "ABCDEFGHIJKLMNOPQRSTUVWXYZ<0123456789",
        },
        onProgress
      );
      ocrText += "\n" + digitsText + "\n" + mrzText;
    }
    console.log("[OCR] OCR listo:", file.name, "pag", pageNum);
    text += "\n" + ocrText;
  }

  console.log("[OCR] Texto total:", file.name, "chars", text.length);
  return text;
}

async function ocrTextFromFiles(files) {
  let combinedText = "";
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const fileName = String(file.name || "").toLowerCase();
    const isDni = fileName.includes("dni") || fileName.includes("documento");
    setMain(`Procesando ${file.name} (${i + 1}/${files.length})...`);
    console.log("[OCR] Archivo:", file.name, "DNI:", isDni);
    const text = await extractTextFromPdf(
      file,
      (m) => {
        if (m?.status) {
          setMain(`OCR ${file.name}: ${m.status} ${Math.round((m.progress || 0) * 100)}%`);
        }
      },
      { isDni }
    );
    combinedText += "\n" + text;
    setProgress(((i + 1) / files.length) * 100);
  }
  return combinedText.trim();
}

async function processFiles(files) {
  console.log("[OCR] processFiles start");
  if (!files || !files.length) {
    setMain("No se seleccionaron PDFs.");
    console.log("[OCR] Sin archivos");
    return;
  }
  if (!window.pdfjsLib || !window.Tesseract) {
    console.log("[OCR] Falta pdfjsLib o Tesseract");
    throw new Error("No se cargaron las librerias de OCR/PDF.");
  }

  try {
    console.log("[OCR] Procesando archivos:", files.length);
    setMain(`Procesando ${files.length} PDF(s)... esto puede demorar.`);
    setProcessing(true);
    setProgress(5);
    let current = getData();
    const combinedText = await ocrTextFromFiles(files);
    setOcrText(combinedText);
    console.log("[OCR] Texto combinado chars", combinedText.length);
    const partial = parseOcrText(combinedText);
    current = mergeData(current, partial);
    setData(current);
    renderPreview();
    if (hasAllRequired(current)) {
      setMain("OCR finalizado. Datos completos.");
      setProgress(100);
      console.log("[OCR] Datos completos");
    }
    if (!hasAllRequired(current)) {
      setMain("OCR finalizado. Revisa la vista previa.");
      console.log("[OCR] OCR finalizado con faltantes");
    }
  } finally {
    setProcessing(false);
    console.log("[OCR] processFiles end");
  }
}

function validate(raw) {
  const required = ["apellido","nombre","dni","calle","numero","localidad","cuit","genero","telefono","email"];
  const missing = required.filter((k) => !raw[k]);
  if (missing.length) throw new Error("Faltan datos OCR: " + missing.join(", "));
}

function mergeData(base, extra) {
  const merged = { ...base };
  for (const key of Object.keys(extra)) {
    if (!merged[key] && extra[key]) {
      merged[key] = extra[key];
    }
  }
  if (!merged.domicilio) {
    merged.domicilio = buildDomicilio(merged.calle, merged.numero, merged.localidad);
  }
  return merged;
}

function hasAllRequired(raw) {
  const required = ["apellido","nombre","dni","calle","numero","localidad","cuit","genero","telefono","email"];
  return required.every((k) => !!raw[k]);
}
function buildDataObject(raw) {
  // Docxtemplater es case-sensitive. En el modelo aparecen tags como &APELLIDO& ... y en el Anexo &NOMBRE & (con espacio).
  const genero = (raw.genero || "").toUpperCase();
  const esMasculino = genero === "M";
  const esFemenino = genero === "F";
  const pick = (m, f) => (esMasculino ? m : esFemenino ? f : "");
  const articulo = pick("el", "la");
  const tratamiento = pick("Sr", "Sra");
  const contraccion = pick("del", "de la");
  const rol = pick("PRESTADOR", "PRESTADORA");
  const pronombre = pick("este", "esta");
  const trabajador = pick("trabajador", "trabajadora");
  const autonomo = pick("autonomo", "autonoma");
  const articuloIndef = pick("un", "una");
  const destacado = pick("destacado", "destacada");
  const contraccionA = pick("al", "a la");
  return {
    APELLIDO: raw.apellido,
    NOMBRE: raw.nombre,
    "NOMBRE ": raw.nombre, // para el tag &NOMBRE &
    DNI: raw.dni,
    CALLE: raw.calle,
    NUMERO: raw.numero,
    LOCALIDAD: raw.localidad,
    CUIT: raw.cuit,
    GENERO: articulo,
    GENERO1: tratamiento,
    GENERO2: contraccion,
    GENERO3: rol,
    GENERO4: pronombre,
    GENERO5: trabajador,
    GENERO6: autonomo,
    GENERO7: articuloIndef,
    GENERO8: destacado,
    GENERO9: contraccionA,
    DOMICILIO: raw.domicilio,
    TELEFONO: raw.telefono,
    EMAIL: raw.email,
  };
}






async function loadTemplateFromFetch() {
  const res = await fetch("template.docx");
  if (!res.ok) throw new Error("No se pudo descargar template.docx");
  const buffer = await res.arrayBuffer();
  state.templateArrayBuffer = normalizeTemplate(buffer);
  setStatus(ui.templateStatus, "Plantilla incluida cargada: template.docx");
}

function normalizeTemplate(arrayBuffer) {
  try {
    const zip = new PizZip(arrayBuffer);
    const path = "word/document.xml";
    const docXml = zip.file(path)?.asText();
    if (!docXml) return arrayBuffer;
    const fixed = docXml
      .replaceAll("%GENERO&amp;", "&amp;GENERO&amp;")
      .replaceAll("%GENERO1&amp;", "&amp;GENERO1&amp;")
      .replaceAll("%GENERO2&amp;", "&amp;GENERO2&amp;");
    if (fixed !== docXml) {
      zip.file(path, fixed);
      return zip.generate({ type: "arraybuffer" });
    }
  } catch (err) {
    console.error(err);
  }
  return arrayBuffer;
}

function assertTemplate() {
  if (!state.templateArrayBuffer) {
    throw new Error("No se pudo cargar la plantilla incluida.");
  }
}

function replaceForPreview(text, dataObj) {
  // Reemplaza tags del modelo (formato &&TAG&&).
  // Para evitar XSS: escapamos primero texto fijo, luego insertamos valores (escapados) en <b>.
  let s = String(text)
    .replaceAll("&&APELLIDO&&", "__APELLIDO__")
    .replaceAll("&&NOMBRE&&", "__NOMBRE__")
    .replaceAll("&&NOMBRE &&", "__NOMBRE__")
    .replaceAll("&&DNI&&", "__DNI__")
    .replaceAll("&&CALLE&&", "__CALLE__")
    .replaceAll("&&NUMERO&&", "__NUMERO__")
    .replaceAll("&&LOCALIDAD&&", "__LOCALIDAD__")
    .replaceAll("&&CUIT&&", "__CUIT__")
    .replaceAll("&&GENERO&&", "__GENERO__")
    .replaceAll("&&GENERO1&&", "__GENERO1__")
    .replaceAll("&&GENERO2&&", "__GENERO2__")
    .replaceAll("&&GENERO3&&", "__GENERO3__")
    .replaceAll("&&GENERO4&&", "__GENERO4__")
    .replaceAll("&&GENERO5&&", "__GENERO5__")
    .replaceAll("&&GENERO6&&", "__GENERO6__")
    .replaceAll("&&GENERO7&&", "__GENERO7__")
    .replaceAll("&&GENERO8&&", "__GENERO8__")
    .replaceAll("&&GENERO9&&", "__GENERO9__")
    .replaceAll("&&DOMICILIO&&", "__DOMICILIO__")
    .replaceAll("&&TELEFONO&&", "__TELEFONO__")
    .replaceAll("&&EMAIL&&", "__EMAIL__")
    .replaceAll("&APELLIDO&", "__APELLIDO__")
    .replaceAll("&NOMBRE&", "__NOMBRE__")
    .replaceAll("&NOMBRE &", "__NOMBRE__")
    .replaceAll("&DNI&", "__DNI__")
    .replaceAll("&CALLE&", "__CALLE__")
    .replaceAll("&NUMERO&", "__NUMERO__")
    .replaceAll("&LOCALIDAD&", "__LOCALIDAD__")
    .replaceAll("&CUIT&", "__CUIT__")
    .replaceAll("&GENERO&", "__GENERO__")
    .replaceAll("&GENERO1&", "__GENERO1__")
    .replaceAll("&GENERO2&", "__GENERO2__")
    .replaceAll("&GENERO3&", "__GENERO3__")
    .replaceAll("&GENERO4&", "__GENERO4__")
    .replaceAll("&GENERO5&", "__GENERO5__")
    .replaceAll("&GENERO6&", "__GENERO6__")
    .replaceAll("&GENERO7&", "__GENERO7__")
    .replaceAll("&GENERO8&", "__GENERO8__")
    .replaceAll("&GENERO9&", "__GENERO9__")
    .replaceAll("&DOMICILIO&", "__DOMICILIO__")
    .replaceAll("&TELEFONO&", "__TELEFONO__")
    .replaceAll("&EMAIL&", "__EMAIL__");

  s = escapeHtml(s);

  const val = (k) => escapeHtml(dataObj[k] ?? "");
  return s
    .replaceAll("__APELLIDO__", `<b>${val("APELLIDO")}</b>`)
    .replaceAll("__NOMBRE__", `<b>${val("NOMBRE")}</b>`)
    .replaceAll("__DNI__", `<b>${val("DNI")}</b>`)
    .replaceAll("__CALLE__", `<b>${val("CALLE")}</b>`)
    .replaceAll("__NUMERO__", `<b>${val("NUMERO")}</b>`)
    .replaceAll("__LOCALIDAD__", `<b>${val("LOCALIDAD")}</b>`)
    .replaceAll("__CUIT__", `<b>${val("CUIT")}</b>`)
    .replaceAll("__GENERO__", `<b>${val("GENERO")}</b>`)
    .replaceAll("__GENERO1__", `<b>${val("GENERO1")}</b>`)
    .replaceAll("__GENERO2__", `<b>${val("GENERO2")}</b>`)
    .replaceAll("__GENERO3__", `<b>${val("GENERO3")}</b>`)
    .replaceAll("__GENERO4__", `<b>${val("GENERO4")}</b>`)
    .replaceAll("__GENERO5__", `<b>${val("GENERO5")}</b>`)
    .replaceAll("__GENERO6__", `<b>${val("GENERO6")}</b>`)
    .replaceAll("__GENERO7__", `<b>${val("GENERO7")}</b>`)
    .replaceAll("__GENERO8__", `<b>${val("GENERO8")}</b>`)
    .replaceAll("__GENERO9__", `<b>${val("GENERO9")}</b>`)
    .replaceAll("__DOMICILIO__", `<b>${val("DOMICILIO")}</b>`)
    .replaceAll("__TELEFONO__", `<b>${val("TELEFONO")}</b>`)
    .replaceAll("__EMAIL__", `<b>${val("EMAIL")}</b>`);
}

function renderPreview() {
  const raw = getData();
  const dataObj = buildDataObject(raw);

  const items = [
    ["Apellido", dataObj.APELLIDO],
    ["Nombre", dataObj.NOMBRE],
    ["DNI", dataObj.DNI],
    ["CUIT", dataObj.CUIT],
    ["Calle", dataObj.CALLE],
    ["Numero", dataObj.NUMERO],
    ["Localidad", dataObj.LOCALIDAD],
    ["Domicilio", dataObj.DOMICILIO],
    ["Telefono", dataObj.TELEFONO],
    ["Email", dataObj.EMAIL],
    ["Genero", raw.genero || ""],
    ["GENERO", dataObj.GENERO],
    ["GENERO1", dataObj.GENERO1],
    ["GENERO2", dataObj.GENERO2],
    ["GENERO3", dataObj.GENERO3],
    ["GENERO4", dataObj.GENERO4],
    ["GENERO5", dataObj.GENERO5],
    ["GENERO6", dataObj.GENERO6],
    ["GENERO7", dataObj.GENERO7],
    ["GENERO8", dataObj.GENERO8],
    ["GENERO9", dataObj.GENERO9],
  ];

  let html = `<h2>Resumen de datos</h2><div class="preview-list">`;
  for (const [label, value] of items) {
    html += `<div class="preview-item"><span>${escapeHtml(label)}</span><b>${escapeHtml(value)}</b></div>`;
  }
  html += `</div>`;
  ui.preview.innerHTML = html;
}

let previewVisible = true;
function setPreviewVisible(show) {
  previewVisible = show;
  if (ui.preview) ui.preview.style.display = show ? "block" : "none";
  if (ui.btnTogglePreview) {
    ui.btnTogglePreview.textContent = show ? "Ocultar vista previa" : "Mostrar vista previa";
  }
}

async function generateWord() {
  const raw = getData();
  if (!raw.domicilio) {
    raw.domicilio = buildDomicilio(raw.calle, raw.numero, raw.localidad);
  }
  validate(raw);
  const dataObj = buildDataObject(raw);

  assertTemplate();

  setMain("Generando Word...");
  const zip = new PizZip(state.templateArrayBuffer);
  const doc = new window.docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "&", end: "&" },
  });

  doc.render(dataObj);

  const blob = doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });

  const baseName = sanitizeFilename(`Contrato_${dataObj.APELLIDO}_${dataObj.NOMBRE}_${dataObj.DNI}`) || "Contrato";
  saveAs(blob, `${baseName}.docx`);
  setMain("Listo.");
}

/* Eventos */


ui.btnGenerate.addEventListener("click", async () => {
  try {
    await generateWord();
  } catch (err) {
    console.error(err);
    setMain("Error: " + err.message);
  }
});

ui.pdfFiles?.addEventListener("change", async () => {
  const count = ui.pdfFiles.files?.length || 0;
  if (count) {
    setMain(`Archivos seleccionados: ${count}. Procesando...`);
    try {
      console.log("[OCR] change input files:", count);
      await processFiles(ui.pdfFiles.files);
    } catch (err) {
      console.error(err);
      setMain("Error: " + err.message);
    }
  } else {
    setMain("Esperando archivos PDF...");
  }
});

ui.dropzone?.addEventListener("dragover", (e) => {
  e.preventDefault();
  ui.dropzone.classList.add("dragover");
});

ui.dropzone?.addEventListener("dragleave", () => {
  ui.dropzone.classList.remove("dragover");
});

ui.dropzone?.addEventListener("drop", async (e) => {
  e.preventDefault();
  ui.dropzone.classList.remove("dragover");
  const files = Array.from(e.dataTransfer?.files || []).filter(
    (file) => file.type === "application/pdf"
  );
  if (!files.length) {
    setMain("No se detectaron PDFs.");
    return;
  }
  if (ui.pdfFiles) {
    const dt = new DataTransfer();
    files.forEach((file) => dt.items.add(file));
    ui.pdfFiles.files = dt.files;
  }
  setMain(`Archivos seleccionados: ${files.length}. Procesando...`);
  try {
    await processFiles(files);
  } catch (err) {
    console.error(err);
    setMain("Error: " + err.message);
  }
});

ui.btnClear.addEventListener("click", () => {
  if (ui.pdfFiles) ui.pdfFiles.value = "";
  resetExtractedData();
  renderPreview();
  setProgress(0);
  setProcessing(false);
  setOcrText("");
  setMain("Listo.");
});

ui.btnTogglePreview?.addEventListener("click", () => {
  setPreviewVisible(!previewVisible);
});

async function init() {
  resetExtractedData();
  renderPreview();
  setPreviewVisible(true);
  setProgress(0);
  setOcrText("");
  if (window.pdfjsLib?.GlobalWorkerOptions) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
  }
  try {
    await loadTemplateFromFetch();
  } catch (err) {
    console.error(err);
    setStatus(
      ui.templateStatus,
      "No se pudo cargar la plantilla incluida. Abr\u00ed con servidor local (ver README). Detalle: " + err.message
    );
    setMain("Error cargando la plantilla incluida.");
  }
}

// Init
init();
