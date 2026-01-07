/* Generador de Contratos (Word) + Vista previa
   Plantilla: template.docx con tags &APELLIDO&, &NOMBRE&, etc.
*/
const $ = (s) => document.querySelector(s);

const state = {
  templateArrayBuffer: null,
  paragraphs: ["Entre la UNIVERSIDAD NACIONAL DE PILAR, representada en este acto por su Rectora, Lic. Elizabeth Diana Wanger, DNI 18.287.351, en adelante LA \"COMITENTE\", por una parte, y el/la Sr/a &APELLIDO&, &NOMBRE&, DNI &DNI&, con domicilio en la calle &CALLE& &NUMERO& de la Localidad de &LOCALIDAD&, en adelante el/la \"PRESTADOR/A DE SERVICIOS\" y en forma conjunta las \"PARTES’’, convienen en celebrar el presente Contrato de Locación de Servicios Profesionales, en adelante el \"Contrato\", el que quedará encuadrado bajo las disposiciones del Código Civil y Comercial de la Nación y estará sujeto a las siguientes términos y condiciones:", "PRIMERA - OBJETO: LA COMITENTE encomienda al/la PRESTADOR/A DE SERVICIOS y este/a acepta en calidad de trabajador/a autónomo/a, la realización de las actividades que se hallan descriptas en la especificaciones técnicas, que como anexo forman parte del presente contrato, para la prestación de un servicio de asesoramiento técnico / formulación de proyectos / servicio de consultoría para proyectos de acreditación de carreras / para la complementación de actividades esenciales de enseñanza y educación / servicio de apoyo estudiantil / otros.", "SEGUNDA – ANTECEDENTES PROFESIONALES: EL/LA PRESTADOR/A DE SERVICIOS es un/a destacado/a profesional en su rubro. Además, posee la formación, experiencia y herramientas técnicas necesarias para la realización de las tareas y/o actividades que comprenden la prestación del servicio.", "TERCERA - SUBCONTRATOS: Como regla general, EL/LA PRESTADOR/A DE SERVICIOS deberá cumplir con sus obligaciones contractuales de manera directa debido a que resulta elegido por sus cualidades para realizarlo personalmente. En el supuesto de requerir una subcontratación, notificará fehacientemente a LA COMITENTE y con la debida antelación. La subcontratación no dará derecho al/la PRESTADOR/A DE SERVICIOS a solicitar modificaciones en el precio del contrato, conforme lo dispuesto en la cláusula QUINTA. De cualquier manera, el/la PRESTADOR/A DE SERVICIOS conservará la dirección y responsabilidad de la ejecución.", "EL/LA PRESTADOR/A DE SERVICIOS preservará y protegerá los derechos de LA COMITENTE respecto de la ejecución bajo subcontratos que pudiera firmar con terceros y deberá:", "Asumir responsabilidad solidaria ante LA COMITENTE para todas las obligaciones y responsabilidades que pudieran originarse de los actos y omisiones generados por su subcontratado y empleados.", "Hacer cumplir a sus subcontratados y/o empleados toda la normativa imperante, reglamentos emitidos por LA COMITENTE y cualquier normativa aplicable.", "EL/LA PRESTADOR/A DE SERVICIOS se constituirá como el único responsable ante sus subcontratados y/o empleados, y los contratos que celebre con éstos nunca le serán oponibles a LA COMITENTE, ni lo podrán alcanzar o afectar, siendo de exclusiva responsabilidad del/la PRESTADOR/A DE SERVICIOS.", "CUARTA – COMPROMISOS DEL PRESTADOR DE SERVICIOS: EL/LA PRESTADOR/A DE SERVICIOS manifiesta: que posee CUIT N°&CUIT& y que se compromete a realizar el servicio para el cual es contratado con total profesionalidad, actuando dentro de las prescripciones éticas y legales que hacen a su disciplina;", "Asimismo, manifiesta que se hará cargo bajo su exclusiva responsabilidad de sus  aportes previsionales y se comprometerá al estricto cumplimiento de los deberes y obligaciones derivadas de las aplicaciones de la legislación vigente, con especial atención a las reglamentaciones de seguridad e higiene;", "Que responderá por el cumplimiento de todas las leyes, ordenanzas, reglamentos y demás disposiciones nacionales y la reglamentación universitaria en materia de locación de servicios profesionales;", "Que conoce la misión y objetivos principales del/la COMITENTE, su estructura académica, sus autoridades y que no encuentra objeción para la ejecución de los servicios por los que se lo ha contratado, así como tampoco ninguna circunstancia que de algún modo impida su avance.", "QUINTA - PRECIO: Se determina un valor total del servicio hasta su total culminación de PESOS ($0.000.000,00) incluyendo IVA -de corresponder-, en adelante el “Precio”, que será abonado por LA COMITENTE al/la PRESTADOR/A DE SERVICIOS, no admitiéndose costos adicionales para financiar tareas o materiales excluidos de los alcances del presente Contrato. EL/LA PRESTADOR/A DE SERVICIOS, que declara haber estudiado y analizado lo requerido, pondrá a disposición los recursos y medios necesarios para la correcta realización del servicio contratado, de acuerdo con los términos y condiciones aquí pactadas, aun cuando ellos no hayan sido expresamente detallados en el presente y su documentación complementaria. Estos elementos que, sin ser mencionados, pudieran ser necesarios para dar cumplimiento al objeto integrante del presente Contrato, no darán derecho al/la PRESTADOR/A DE SERVICIOS a requerir modificaciones en el precio del Contrato.", "SEXTA – MODALIDAD DE PAGO: El precio total y absoluto del servicio contratado deberá pagarse a nombre de &NOMBRE&, &APELLIDO&, CUIT N°&CUIT& y de la siguiente manera: PAGOS SEMANALES / QUINCENALES / MENSUALES / BIMESTRALES / que El COMITENTE abonará a razón de PESOS ($000.000,00) hasta alcanzar la suma de PESOS ($0.000.000,00) con fecha límite del (día) de (mes) de 202x.", "SEPTIMA - PLAZOS DE EJECUCIÓN: Se estipula un período para la ejecución de las obligaciones contractuales de XXXX días corridos, contados desde el (día) de (mes) de 202x hasta el (día) de (mes) de 202x; prorrogables únicamente por estipulación expresa contractual.", "El COMITENTE procederá a la cancelación total de la contraprestación una vez canceladas todas las obligaciones asumidas por EL/LA PRESTADOR/A DE SERVICIOS.", "OCTAVA - RESCISIÓN DEL CONTRATO: Sin perjuicio de lo establecido en la normativa nacional vigente, en caso de incumplimiento de las condiciones pactadas en el presente Contrato por cualquiera de las partes, la otra podrá optar por rescindir el mismo intimando previamente a la contraparte al cumplimiento de la prestación debida, en forma fehaciente con cinco (5) días de anticipación. En caso de persistir el incumplimiento, el Contrato quedará rescindido de pleno derecho, pudiendo la parte cumplidora reclamar por los daños y perjuicios ocasionados.", "Asimismo, LA COMITENTE podrá, en cualquier etapa y sin notificación previa, rescindir unilateralmente y de manera anticipada el presente contrato, en uso de las atribuciones conferidas por el artículo 11 y 12, inciso a) del Decreto N°1023/2001 del RÉGIMEN DE CONTRATACIONES DE LA ADMINISTRACIÓN NACIONAL.", "NOVENA - CLAUSULA DE INDEMNIDAD: El/la PRESTADOR/A DE SERVICIOS deberá indemnizar y mantener informada a LA COMITENTE de cualquier reclamo, acción legal u otros procedimientos realizados por terceras partes que fueran atribuibles a cualquier acto u omisión que surja de la realización del contrato. Se conviene que dicha indemnización y deber de información se aplicará a los reclamos presentados por Entidades Gubernamentales Federales, Provinciales o Municipales. Asimismo, tal protección se aplicará a todos los reclamos que resulten aún después de prestado el servicio, como así también los presentados por sus empleados o subcontratistas, autoridades fiscales, impositivas o de la seguridad social respecto a la ejecución de tareas que estén bajo la órbita de su realización.", "DÉCIMA - CONFIDENCIALIDAD: EL/LA PRESTADOR/A DE SERVICIOS deberá, de conformidad con lo dispuesto en la Ley N°24.766, guardar estricta confidencialidad sobre las informaciones no publicadas, o de carácter reservado o confidencial contenidas en documentación, informes, expedientes, sistemas informáticos y cualquier otro medio en posesión del COMITENTE, con motivo de la ejecución de las obligaciones emanadas del presente Contrato, salvo autorización expresa de LA COMITENTE. Esta obligación de confidencialidad permanecerá en vigencia aún después de su cumplimiento, rescisión o resolución, siendo responsable EL/LA PRESTADOR/A DE SERVICIOS de los daños y perjuicios que pudiera ocasionar la difusión indebida.", "DÉCIMA PRIMERA - PROPIEDAD INTELECTUAL: Los derechos de propiedad de autor y de reproducción, así como cualquier otro derecho intelectual de cualquier naturaleza sobre informes, estudios, etc. producidos como consecuencia del cumplimiento de las obligaciones contractuales, pertenecerán exclusivamente a LA COMITENTE.", "DÉCIMA SEGUNDA - DERECHO DE INSPECCIÓN: LA COMITENTE puede inspeccionar los servicios prestados a través de algún representante o por sí, todas las veces que lo crea necesario, con motivo de constatar el fiel cumplimiento de todas las obligaciones que EL/LA PRESTADOR/A DE SERVICIOS asume.", "DÉCIMA TERCERA - ADENDAS AL CONTRATO: En el supuesto de que surjan propuestas de adendas que LAS PARTES deseen efectuar sobre las cláusulas que rigen la presente contratación, deberán comunicarse en forma previa y fehaciente para su análisis y la realización del procedimiento administrativo requerido para su eventual aprobación.", "DECIMA CUARTA - DOMICILIOS DE PAGO: Se determina como domicilio de pago expreso de las sumas reconocidas y estipuladas en la cláusula CUARTA los denunciados por LAS PARTES en el encabezamiento del presente contrato, o en su defecto los denunciados en las Especificaciones Técnicas que integran el Anexo referido anteriormente.", "DÉCIMA QUINTA - COMPETENCIA JUDICIAL: En caso de controversia judicial o prejudicial producto de la aplicación y/o interpretación de alguna de las cláusulas que rigen la presente contratación, LAS PARTES se someten a la Jurisdicción de los Juzgados Federales de Campana.", "En prueba de conformidad, en la ciudad de Pilar, Provincia de Buenos Aires y a los …… días del mes de ………….. de 20.., se suscriben dos ejemplares de un mismo tenor y a un solo efecto.", "Anexo", "ESPECIFICACIONES TÉCNICAS", "PRESTADOR DE SERVICIOS:: &APELLIDO&, &NOMBRE &", "CUIT: &CUIT&", "DOMICILIO: &DOMICILIO&", "TELÉFONO: &TELEFONO&", "E MAIL: &EMAIL&", "DURACIÓN:", "FECHA DE INICIO: 00/00/2025", "FECHA DE FINALIZACIÓN: 00/00/2025", "OBJETIVO GENERAL DE LA CONTRATACIÓN", "__________________________________________________________________________________________________________________________________________", "REQUERIMIENTOS Y CONDICIONES DE TRABAJO", "“El PRESTADOR DE SERVICIOS” deberá realizar tareas encomendadas en el período establecido en la cláusula SÉPTIMA.", "Las tareas por realizar por “El PRESTADOR DE SERVICIOS” podrán sufrir modificaciones por indicación de “LA COMITENTE”, para ser adecuada a las variaciones que puedan experimentar el desarrollo de los objetivos para los que fue contratada y el mejor logro de estos.", "FIRMA DEL PRESTADOR DE SERVICIOS: \t\t\tFIRMA COMITENTE", "ACLARACIÓN: \t\t\t\t\t\t\tSELLO", "DNI:"],
  title: "CONTRATO DE LOCACIÓN DE SERVICIOS PROFESIONALES"
};

const ui = {
  templateStatus: $("#templateStatus"),

  form: $("#form"),
  btnGenerate: $("#btnGenerate"),
  btnClear: $("#btnClear"),
  mainStatus: $("#mainStatus"),
  preview: $("#preview"),
  btnTogglePreview: $("#btnTogglePreview"),
};

function setStatus(el, msg) {
  if (!el) return;
  el.textContent = msg || "";
}
function setMain(msg) {
  setStatus(ui.mainStatus, msg);
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

function coerce(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

function getFormData() {
  const fd = new FormData(ui.form);
  const raw = Object.fromEntries(fd.entries());

  // domicilio opcional: si está vacío lo armamos
  let domicilio = coerce(raw.domicilio);
  if (!domicilio) {
    const parts = [coerce(raw.calle), coerce(raw.numero), coerce(raw.localidad)].filter(Boolean);
    domicilio = parts.join(" ");
  }

  return {
    apellido: coerce(raw.apellido),
    nombre: coerce(raw.nombre),
    dni: coerce(raw.dni),
    calle: coerce(raw.calle),
    numero: coerce(raw.numero),
    localidad: coerce(raw.localidad),
    cuit: coerce(raw.cuit),
    domicilio,
    telefono: coerce(raw.telefono),
    email: coerce(raw.email),
  };
}

function validate(raw) {
  const required = ["apellido","nombre","dni","calle","numero","localidad","cuit","telefono","email"];
  const missing = required.filter(k => !raw[k]);
  if (missing.length) throw new Error("Faltan campos obligatorios: " + missing.join(", "));
}

function buildDataObject(raw) {
  // Docxtemplater es case-sensitive. En el modelo aparecen tags como &APELLIDO& ... y en el Anexo &NOMBRE & (con espacio).
  return {
    APELLIDO: raw.apellido,
    NOMBRE: raw.nombre,
    "NOMBRE ": raw.nombre, // para el tag &NOMBRE &
    DNI: raw.dni,
    CALLE: raw.calle,
    NUMERO: raw.numero,
    LOCALIDAD: raw.localidad,
    CUIT: raw.cuit,
    DOMICILIO: raw.domicilio,
    TELEFONO: raw.telefono,
    EMAIL: raw.email,
  };
}

async function loadTemplateFromFetch() {
  const res = await fetch("template.docx");
  if (!res.ok) throw new Error("No se pudo cargar la plantilla incluida.");
  state.templateArrayBuffer = await res.arrayBuffer();
  setStatus(ui.templateStatus, "Plantilla incluida cargada: template.docx");
}

function assertTemplate() {
  if (!state.templateArrayBuffer) {
    throw new Error("No se pudo cargar la plantilla incluida.");
  }
}

function replaceForPreview(text, dataObj) {
  // Reemplaza tags del modelo (formato &TAG&).
  // Para evitar XSS: escapamos primero texto fijo, luego insertamos valores (escapados) en <b>.
  let s = String(text)
    .replaceAll("&APELLIDO&", "__APELLIDO__")
    .replaceAll("&NOMBRE&", "__NOMBRE__")
    .replaceAll("&NOMBRE &", "__NOMBRE__")
    .replaceAll("&DNI&", "__DNI__")
    .replaceAll("&CALLE&", "__CALLE__")
    .replaceAll("&NUMERO&", "__NUMERO__")
    .replaceAll("&LOCALIDAD&", "__LOCALIDAD__")
    .replaceAll("&CUIT&", "__CUIT__")
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
    .replaceAll("__DOMICILIO__", `<b>${val("DOMICILIO")}</b>`)
    .replaceAll("__TELEFONO__", `<b>${val("TELEFONO")}</b>`)
    .replaceAll("__EMAIL__", `<b>${val("EMAIL")}</b>`);
}

function renderPreview() {
  const raw = getFormData();
  const dataObj = buildDataObject(raw);

  const isHeading = (s) => s && s.length <= 80 && s === s.toUpperCase();

  let html = `<h1>${escapeHtml(state.title)}</h1>`;
  for (const p of state.paragraphs) {
    const replaced = replaceForPreview(p, dataObj);
    if (isHeading(p)) html += `<h2>${replaced}</h2>`;
    else html += `<p>${replaced}</p>`;
  }
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
  const raw = getFormData();
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
  setMain("Listo ✅");
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

ui.btnClear.addEventListener("click", () => {
  ui.form.reset();
  renderPreview();
  setMain("Listo.");
});

ui.form.addEventListener("input", () => {
  renderPreview();
});

ui.btnTogglePreview?.addEventListener("click", () => {
  setPreviewVisible(!previewVisible);
});

async function init() {
  renderPreview();
  setPreviewVisible(true);
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
