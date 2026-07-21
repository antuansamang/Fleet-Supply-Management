import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import {
  Anchor, Upload, Receipt, CheckCircle2, Circle, Trash2, ChevronDown,
  ChevronRight, Loader2, Package, Utensils, Wrench, BedDouble, Waves, Compass,
  AlertTriangle, X, FileText, ClipboardList, LayoutDashboard, Boxes, Tags,
  Plus, RefreshCw, CalendarDays, ArrowUpRight, ShieldCheck, Download, Search, Smartphone,
} from "lucide-react";
import { supabase } from "./supabaseClient";

/* ============================================================
   EXPLORER FLEET CONTROL · Abastecimiento y Requisiciones
   Flota Galápagos: Humboldt Explorer · Tiburón Explorer · Grand Majestic
   ============================================================ */

const F_MONO = { fontFamily: "'IBM Plex Mono', ui-monospace, monospace" };
const F_BASE = { fontFamily: "'Inter', system-ui, sans-serif" };

const BARCOS = [
  { id: "humboldt", nombre: "Humboldt Explorer" },
  { id: "tiburon", nombre: "Tiburón Explorer" },
  { id: "majestic", nombre: "Grand Majestic" },
];
const BARCO_IDS = BARCOS.map((b) => b.id);

const DEPARTAMENTOS = [
  { id: "ABASTECIMIENTO", nombre: "Abastecimiento · Alimentos y Víveres", corto: "Abastecimiento", Icon: Utensils, dot: "bg-teal-500", barra: "bg-teal-500" },
  { id: "MAQUINAS", nombre: "Máquinas", corto: "Máquinas", Icon: Wrench, dot: "bg-amber-500", barra: "bg-amber-500" },
  { id: "CUBIERTA", nombre: "Cubierta", corto: "Cubierta", Icon: Anchor, dot: "bg-sky-500", barra: "bg-sky-500" },
  { id: "HOTELERIA", nombre: "Hotelería", corto: "Hotelería", Icon: BedDouble, dot: "bg-violet-500", barra: "bg-violet-500" },
  { id: "BUCEO", nombre: "Buceo", corto: "Buceo", Icon: Waves, dot: "bg-cyan-500", barra: "bg-cyan-500" },
  { id: "PUENTE", nombre: "Puente", corto: "Puente", Icon: Compass, dot: "bg-slate-400", barra: "bg-slate-400" },
  { id: "OTRO", nombre: "Otro", corto: "Otro", Icon: Package, dot: "bg-stone-400", barra: "bg-stone-400" },
];
const depInfo = (id) => DEPARTAMENTOS.find((d) => d.id === id) || DEPARTAMENTOS[DEPARTAMENTOS.length - 1];

/* ---------------- Utilidades ---------------- */

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

const num = (v) => {
  if (v === null || v === undefined || v === "") return null;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : null;
};

const round2 = (v) => (v == null ? null : Math.round(v * 100) / 100);

const fmtMoney = (v) =>
  v == null ? "—" : "$" + Number(v).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

const fmtFecha = (iso) => {
  try { return new Date(iso).toLocaleDateString("es-EC", { day: "2-digit", month: "short", year: "numeric" }); }
  catch (e) { return iso || "—"; }
};

const fmtFechaHora = (iso) => {
  try { return new Date(iso).toLocaleString("es-EC", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }); }
  catch (e) { return iso || "—"; }
};

const normaliza = (s) => String(s || "").toLowerCase();

/* ---------------- Semanas de crucero (lunes a lunes, 8 días; rango editable) ---------------- */

const MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
const MES_CORTO = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];

const isoLocal = (d) =>
  d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");

const parseISO = (s) => {
  const [y, m, d] = String(s).split("-").map(Number);
  return new Date(y, (m || 1) - 1, d || 1);
};

const esISO = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s));

const lunesDe = (fecha) => {
  const d = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
  const dia = d.getDay();
  d.setDate(d.getDate() + (dia === 0 ? -6 : 1 - dia));
  return d;
};

const sumarDias = (iso, n) => {
  const d = parseISO(iso);
  d.setDate(d.getDate() + n);
  return isoLocal(d);
};

const semanaActual = () => isoLocal(lunesDe(new Date()));

const claveSemana = (r) => String(r.semana) + "|" + String(r.semanaFin || "");

const finDe = (iso, finIso) => {
  const f = esISO(finIso) ? parseISO(finIso) : parseISO(iso);
  if (!esISO(finIso)) f.setDate(f.getDate() + 7);
  return f;
};

const nombreSemana = (iso, finIso) => {
  if (!esISO(iso)) return String(iso ?? "—");
  const ini = parseISO(iso);
  const fin = finDe(iso, finIso);
  let etiqueta =
    ini.getMonth() === fin.getMonth()
      ? ini.getDate() + " al " + fin.getDate() + " de " + MESES[ini.getMonth()]
      : ini.getDate() + " de " + MESES[ini.getMonth()] + " al " + fin.getDate() + " de " + MESES[fin.getMonth()];
  const hoy = new Date().getFullYear();
  if (ini.getFullYear() !== hoy || fin.getFullYear() !== hoy) etiqueta += " " + fin.getFullYear();
  return etiqueta;
};

const rangoCorto = (iso, finIso) => {
  if (!esISO(iso)) return "";
  const ini = parseISO(iso);
  const fin = finDe(iso, finIso);
  return ini.getMonth() === fin.getMonth()
    ? ini.getDate() + "–" + fin.getDate() + " " + MES_CORTO[ini.getMonth()]
    : ini.getDate() + " " + MES_CORTO[ini.getMonth()] + " – " + fin.getDate() + " " + MES_CORTO[fin.getMonth()];
};

const numeroSemana = (iso) => {
  if (!esISO(iso)) return null;
  const t = parseISO(iso);
  t.setDate(t.getDate() + 3 - ((t.getDay() + 6) % 7));
  const w1 = new Date(t.getFullYear(), 0, 4);
  return 1 + Math.round(((t - w1) / 86400000 - 3 + ((w1.getDay() + 6) % 7)) / 7);
};

const generarSemanas = () => {
  const base = lunesDe(new Date());
  const lista = [];
  for (let i = -16; i <= 36; i++) {
    const d = new Date(base);
    d.setDate(d.getDate() + i * 7);
    lista.push(isoLocal(d));
  }
  return lista;
};

const migrarSemana = (r) => {
  if (!r) return r;
  if (esISO(r.semana)) {
    if (esISO(r.semanaFin)) return r;
    return { ...r, semanaFin: sumarDias(r.semana, 7) };
  }
  const n = Number(r.semana);
  if (!Number.isFinite(n) || n < 1 || n > 53) return r;
  const anio = r.fechaSubida ? new Date(r.fechaSubida).getFullYear() : new Date().getFullYear();
  const inicio = isoLocal(lunesDe(new Date(anio, 0, 1 + (n - 1) * 7)));
  return { ...r, semana: inicio, semanaFin: sumarDias(inicio, 7) };
};

/* ---------------- Lectura de archivos ---------------- */

const toBase64 = (file) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result).split(",")[1]);
    r.onerror = () => rej(new Error("No se pudo leer el archivo " + file.name));
    r.readAsDataURL(file);
  });

async function fileToBlocks(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".pdf")) {
    const b64 = await toBase64(file);
    return [{ type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } }];
  }
  if (/\.(png|jpe?g|gif|webp)$/.test(name)) {
    const mt = name.endsWith(".png") ? "image/png"
      : name.endsWith(".webp") ? "image/webp"
      : name.endsWith(".gif") ? "image/gif" : "image/jpeg";
    const b64 = await toBase64(file);
    return [{ type: "image", source: { type: "base64", media_type: mt, data: b64 } }];
  }
  if (/\.(xlsx|xls|csv)$/.test(name)) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
    let text = "";
    wb.SheetNames.forEach((sn) => {
      text += "--- HOJA: " + sn + " ---\n" + XLSX.utils.sheet_to_csv(wb.Sheets[sn]) + "\n";
    });
    return [{ type: "text", text: "CONTENIDO DEL ARCHIVO (" + file.name + "):\n" + text.slice(0, 60000) }];
  }
  if (name.endsWith(".docx")) {
    const buf = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer: buf });
    return [{ type: "text", text: "CONTENIDO DEL ARCHIVO (" + file.name + "):\n" + String(result.value || "").slice(0, 60000) }];
  }
  if (name.endsWith(".txt")) {
    const text = await file.text();
    return [{ type: "text", text: "CONTENIDO DEL ARCHIVO (" + file.name + "):\n" + text.slice(0, 60000) }];
  }
  if (name.endsWith(".doc")) {
    throw new Error("Los archivos .doc antiguos no son compatibles. Guárdalo como .docx o PDF y vuelve a subirlo.");
  }
  throw new Error("Formato no soportado: " + file.name + ". Usa PDF, PNG, JPG, Excel (.xlsx/.xls/.csv) o Word (.docx).");
}

/* ---------------- API de Claude ---------------- */

async function callClaude(contentBlocks, maxTokens) {
  // La llamada real a la API de Anthropic ocurre en /api/claude (función serverless),
  // que guarda la clave API en el servidor. El navegador nunca ve la clave.
  const response = await fetch("/api/claude", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      max_tokens: maxTokens || 4000,
      messages: [{ role: "user", content: contentBlocks }],
    }),
  });
  if (!response.ok) throw new Error("Error de conexión con Claude (HTTP " + response.status + "). Intenta de nuevo.");
  const data = await response.json();
  if (data.error) throw new Error("Claude devolvió un error: " + (data.error.message || "desconocido"));
  return (data.content || []).map((b) => (b.type === "text" ? b.text : "")).filter(Boolean).join("\n");
}

/* Parseo de JSON con detección de respuestas truncadas */
function parseJSONSafe(raw) {
  let t = String(raw || "").replace(/```json/gi, "").replace(/```/g, "").trim();
  const s = t.indexOf("{");
  if (s > 0) t = t.slice(s);
  try { return { data: JSON.parse(t), truncado: false }; } catch (e) { /* reparar */ }
  const cierres = ["", "]", "]}", "]]}", "\"]]}", "}]", "}]}", "\"]}", "}}"];
  const limite = Math.max(0, t.length - 900);
  for (let i = t.length; i > limite; i--) {
    const c = t[i - 1];
    if (c !== "]" && c !== "}" && c !== "\"" && !(c >= "0" && c <= "9")) continue;
    const base = t.slice(0, i);
    for (const suf of cierres) {
      try { return { data: JSON.parse(base + suf), truncado: true }; } catch (e) { /* siguiente */ }
    }
  }
  throw new Error("No se pudo interpretar la respuesta del análisis. Vuelve a intentarlo.");
}

/* ---------------- Prompts y extracción con doble verificación ---------------- */

const PROMPT_REQUISICION = `Analiza el documento adjunto: es un pedido de abastecimiento o una requisición de un departamento de un buque de expedición en Galápagos.

Extrae TODAS las filas de producto de la lista, SIN EXCEPCIÓN y en el mismo orden del documento. Incluye también las filas cuya columna de pedido diga "NO", "0", "-" o esté vacía: forman parte de la lista y deben salir con cantidad 0. Está PROHIBIDO filtrar u omitir filas de producto.

Responde ÚNICAMENTE con JSON válido, sin markdown ni texto adicional:
{"items":[["DESCRIPCION",cantidad,"UNIDAD"],["OTRA DESCRIPCION",cantidad,"UNIDAD"]],"fin":true}

Reglas:
- Una entrada por CADA fila de producto del documento, aunque su pedido sea "NO" o 0 (esas van con cantidad 0).
- Si hay varias columnas numéricas (por ejemplo "CANT. MAXIMA" y "PEDIDO"), la cantidad es SIEMPRE la de la columna del pedido real (normalmente la última). Nunca uses la cantidad máxima ni el stock.
- "DESCRIPCION": nombre del producto en MAYÚSCULAS, conciso (máximo 6 palabras).
- cantidad: número con punto decimal. "NO", "-" o pedido vacío = 0. Si el documento no tiene columna de pedido y la fila no indica cantidad, usa 1.
- "UNIDAD": la unidad que indica el documento, abreviada: KG, LB, UND, LT, GAL, CAJA, PAQ, FUNDA, JABA, CUBETA, BOTELLA, M, ROLLO (LITRO=LT, LIBRA=LB, KILO=KG, UNIDAD=UND). Si no se indica, usa "UND".
- Los títulos de sección o categoría (VEGETALES, CARNICOS, PESCADO, BEBIDAS…), encabezados, firmas y totales NO son ítems.
- "fin": true si listaste TODAS las filas del documento; false si te quedaste sin espacio (te pediré continuar).
- No inventes ítems.`;

const filaAItem = (row) => {
  const d = String((row && row[0]) || "").toUpperCase().trim();
  if (!d) return null;
  const u = String((row && row[2]) || "UND").toUpperCase().trim();
  const bruto = row && row[1];
  const sinPedido = typeof bruto === "string" && ["NO", "N/A", "-", "--", "0"].includes(bruto.trim().toUpperCase());
  return {
    id: uid(), descripcion: d, cantidad: sinPedido ? 0 : (num(bruto) ?? 1), unidad: u || "UND",
    recibido: false, cantidadFacturada: null, precioUnitario: null, precioTotal: null, factura: null,
  };
};

function promptContinuacion(items) {
  const ult = items.slice(-3).map((t) => [t.descripcion, t.cantidad, t.unidad]);
  return `Ya extraje los primeros ${items.length} ítems de este documento. Los últimos fueron: ${JSON.stringify(ult)}.
Continúa EXACTAMENTE desde la fila siguiente del documento, sin repetir las anteriores.
Responde ÚNICAMENTE con JSON válido: {"items":[["DESCRIPCION",cantidad,"UNIDAD"]],"fin":true}
- "fin": true si con esto ya quedan listadas TODAS las filas de producto; false si aún faltan.
- Mismas reglas: incluye TODAS las filas de producto, también las marcadas "NO" o sin pedido (con cantidad 0); usa la columna del PEDIDO real (no la cantidad máxima); MAYÚSCULAS concisas; unidad del documento abreviada ("UND" si no se indica).
- Si ya no quedan filas, responde {"items":[],"fin":true}.`;
}

function promptVerificacionReq(items) {
  const compacta = items.map((it, i) => [i, it.descripcion, it.cantidad, it.unidad]);
  return `DOBLE VERIFICACIÓN. Esta es la lista de ítems que extraje del documento adjunto, en formato [indice, "descripción", cantidad, "unidad"]:
${JSON.stringify(compacta)}

Compárala línea por línea contra el documento y responde ÚNICAMENTE con JSON válido:
{"correcciones":[[indice,"DESCRIPCION",cantidad,"UNIDAD"]],"faltantes":[["DESCRIPCION",cantidad,"UNIDAD"]],"total_documento":numero}

- "correcciones": solo los índices cuya cantidad, unidad o descripción NO coinciden con el documento, con los valores correctos. La cantidad correcta es la de la columna del PEDIDO real (no la cantidad máxima); las filas cuyo pedido dice "NO" o está vacío deben tener cantidad 0.
- "faltantes": filas de producto del documento que NO están en la lista, INCLUIDAS las marcadas "NO" o sin pedido (esas van con cantidad 0). No cuentes títulos de sección (VEGETALES, CARNICOS, BEBIDAS…).
- "total_documento": cuántas filas de producto tiene el documento en total, contando también las marcadas "NO" y sin contar títulos de sección ni encabezados.
- Si todo está correcto y completo: {"correcciones":[],"faltantes":[],"total_documento":numero}.`;
}

async function extraerItemsVerificado(blocks, onProgreso) {
  const items = [];
  const vistos = new Set();
  const claveDe = (it) => it.descripcion + "|" + it.cantidad + "|" + it.unidad;
  const agregarFilas = (filas) => {
    let agregados = 0;
    (Array.isArray(filas) ? filas : []).forEach((row) => {
      const it = filaAItem(row);
      if (!it) return;
      const k = claveDe(it);
      if (vistos.has(k)) return;
      vistos.add(k);
      items.push(it);
      agregados++;
    });
    return agregados;
  };

  let ronda = 0;
  let continuar = true;
  while (continuar && ronda < 6) {
    ronda++;
    if (onProgreso) onProgreso(ronda === 1 ? "Leyendo el documento…" : "Leyendo más ítems… " + items.length + " detectados");
    const prompt = ronda === 1 ? PROMPT_REQUISICION : promptContinuacion(items);
    const texto = await callClaude([...blocks, { type: "text", text: prompt }], 4000);
    const { data, truncado } = parseJSONSafe(texto);
    const agregados = agregarFilas(data.items);
    continuar = (truncado || data.fin === false) && agregados > 0;
  }
  if (!items.length) return { items, verificacion: null };

  if (onProgreso) onProgreso("Doble verificación de ítems y cantidades…");
  let verificacion = { correcciones: 0, faltantes: 0, totalDoc: null };
  try {
    const texto = await callClaude([...blocks, { type: "text", text: promptVerificacionReq(items) }], 3000);
    const { data } = parseJSONSafe(texto);
    (Array.isArray(data.correcciones) ? data.correcciones : []).forEach((c) => {
      const it = items[c[0]];
      if (!it) return;
      const d = String(c[1] || it.descripcion).toUpperCase().trim();
      it.descripcion = d || it.descripcion;
      it.cantidad = num(c[2]) ?? it.cantidad;
      it.unidad = String(c[3] || it.unidad).toUpperCase().trim() || it.unidad;
      verificacion.correcciones++;
    });
    (Array.isArray(data.faltantes) ? data.faltantes : []).forEach((row) => {
      const it = filaAItem(row);
      if (!it) return;
      const k = claveDe(it);
      if (vistos.has(k)) return;
      vistos.add(k);
      items.push(it);
      verificacion.faltantes++;
    });
    verificacion.totalDoc = num(data.total_documento);
  } catch (e) {
    verificacion = null;
  }
  return { items, verificacion };
}

function promptFactura(compact) {
  const base = `El documento adjunto es una FACTURA de compra para un buque de expedición en Galápagos.

TAREA 1: Extrae el número de factura, el proveedor, la fecha de emisión y el TOTAL de la factura.
`;
  const tarea2 = compact.length
    ? `TAREA 2: Compara cada ítem de la factura con esta lista de ÍTEMS PENDIENTES de requisiciones. Formato: [id, "descripción", cantidad_pedida, "unidad"]:
${JSON.stringify(compact)}
`
    : `TAREA 2: No hay ítems pendientes registrados, así que todos los ítems de la factura van en "x".
`;
  return base + tarea2 + `
Responde ÚNICAMENTE con JSON válido, sin markdown:
{"num":"numero","prov":"proveedor","fecha":"fecha","total":total_factura,"m":[[id,cantidad_facturada,precio_unitario,precio_total]],"x":[["DESCRIPCION",cantidad,"UNIDAD",precio_unitario,precio_total]]}

Reglas:
- "m" (coincidencias): ítems de la factura que corresponden a un pendiente. Usa el id de la lista. Cuenta como coincidencia aunque el nombre varíe levemente (singular/plural, marca, sinónimo). Cada id máximo una vez.
- "x" (extras): ítems de la factura que NO corresponden a ningún pendiente. Descripción concisa en MAYÚSCULAS.
- Precios y total: números con punto decimal, sin símbolo $. Si falta el total de línea, calcúlalo como cantidad × unitario. Usa null si un dato no existe.
- Cada ítem de la factura debe ir en "m" o en "x", sin omitir ninguno.`;
}

function promptVerificacionFactura(lineas) {
  return `DOBLE VERIFICACIÓN. Estas son las líneas que capturé de la factura adjunta, en formato [codigo, "descripción", cantidad, precio_unitario, precio_total]:
${JSON.stringify(lineas)}

Compáralas línea por línea contra la factura y responde ÚNICAMENTE con JSON válido:
{"correcciones":[["codigo",cantidad,precio_unitario,precio_total]],"faltantes":[["DESCRIPCION",cantidad,"UNIDAD",precio_unitario,precio_total]],"total":total_factura}

- "correcciones": solo los códigos cuya cantidad o precios NO coinciden con la factura, con los valores correctos.
- "faltantes": líneas de la factura que NO están en mi lista.
- "total": el total real de la factura (número, sin $). Usa null si no aparece.
- Si todo está correcto: {"correcciones":[],"faltantes":[],"total":total_factura}.`;
}

function promptSugerenciaVinculo(extra, candidatos) {
  return `Una factura de compra de un buque en Galápagos trae esta línea que NO coincidió automáticamente con ninguna requisición:
${JSON.stringify([extra.descripcion, extra.cantidad, extra.unidad])}

Estos son los ítems PENDIENTES de las requisiciones, en formato [indice, "descripción", cantidad, "unidad", "requisición", "departamento"]:
${JSON.stringify(candidatos)}

¿A cuál ítem pendiente corresponde más probablemente esa línea de la factura? Piensa en sinónimos, traducciones inglés/español, nombres comerciales vs. genéricos (ej. "SMELL CLEAN" es un eliminador de olores, "PINOKLIN" es un desinfectante de pino), presentaciones y cantidades parecidas.

Responde ÚNICAMENTE con JSON válido, sin markdown:
{"indice":numero_o_null,"motivo":"explicación muy breve"}
- "indice": el índice del candidato más probable, o null si ninguno es razonable.
- "motivo": máximo 12 palabras, en español.`;
}

/* ---------------- Almacenamiento por embarcación (Supabase) ---------------- */

const TABLA_REQ = "hx_requisiciones";
const TABLA_FACT = "hx_facturas";
const TABLA_CONFIG = "hx_config";

const claveReq = (barcoId, id) => "hx-req:" + barcoId + ":" + id;
const claveFact = (barcoId, id) => "hx-fact:" + barcoId + ":" + id;

/* Interpreta una "clave" con el mismo formato que usaba window.storage
   ("hx-req:barco:id", "hx-fact:barco:id" o "hx-config") y determina a qué
   tabla y fila de Supabase corresponde. */
function _resolverClave(clave) {
  if (clave === "hx-config") return { tabla: TABLA_CONFIG, id: "singleton", barcoId: null };
  const partes = String(clave).split(":");
  const tipo = partes[0];
  const barcoId = partes[1];
  const id = partes.slice(2).join(":");
  return { tabla: tipo === "hx-req" ? TABLA_REQ : TABLA_FACT, id, barcoId };
}

async function persistir(clave, obj) {
  const { tabla, id, barcoId } = _resolverClave(clave);
  try {
    const fila =
      tabla === TABLA_CONFIG
        ? { id, data: obj, updated_at: new Date().toISOString() }
        : { id, barco_id: barcoId, data: obj, updated_at: new Date().toISOString() };
    const { error } = await supabase.from(tabla).upsert(fila);
    if (error) console.error("Error guardando en Supabase (" + tabla + "):", error.message);
  } catch (e) { /* silencioso */ }
}

async function eliminar(clave) {
  const { tabla, id } = _resolverClave(clave);
  try { await supabase.from(tabla).delete().eq("id", id); } catch (e) { /* silencioso */ }
}

async function obtenerConfig() {
  try {
    const { data, error } = await supabase.from(TABLA_CONFIG).select("data").eq("id", "singleton").maybeSingle();
    if (error || !data) return null;
    return data.data;
  } catch (e) { return null; }
}

async function cargarRegistros(tipo, barcoId) {
  const tabla = tipo === "req" ? TABLA_REQ : TABLA_FACT;
  try {
    const { data, error } = await supabase.from(tabla).select("data").eq("barco_id", barcoId);
    if (error || !data) return [];
    return data.map((fila) => fila.data);
  } catch (e) { return []; }
}

/* ============================================================
   Componentes de interfaz (nivel de módulo: identidad estable)
   ============================================================ */

/* Un ítem "exigible" es el que espera compra: pedido > 0, o ya facturado.
   Los marcados "NO" en el documento (cantidad 0) quedan visibles pero no cuentan como pendientes. */
const esExigible = (it) => it.recibido || (num(it.cantidad) ?? 0) > 0;

const estadoDe = (r) => {
  const exigibles = (r.items || []).filter(esExigible);
  const t = exigibles.length;
  const rec = exigibles.filter((i) => i.recibido).length;
  if (t === 0 || rec === 0) return "PENDIENTE";
  if (rec < t) return "PARCIAL";
  return "CONCILIADA";
};

const Etiqueta = ({ children }) => (
  <p className="mb-1 text-xs font-semibold uppercase tracking-widest text-slate-500">{children}</p>
);

const BannerError = ({ msg }) =>
  !msg ? null : (
    <div className="flex items-start gap-2 rounded-lg border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
      <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0" />
      <span>{msg}</span>
    </div>
  );

const Pill = ({ e }) => {
  const estilos = {
    PENDIENTE: "bg-slate-100 text-slate-600",
    PARCIAL: "bg-amber-100 text-amber-700",
    CONCILIADA: "bg-emerald-100 text-emerald-700",
  };
  const punto = { PENDIENTE: "bg-slate-400", PARCIAL: "bg-amber-500", CONCILIADA: "bg-emerald-500" };
  return (
    <span className={"inline-flex items-center gap-1.5 rounded-full px-2 py-0.5 text-xs font-semibold uppercase tracking-wide " + estilos[e]}>
      <span className={"h-1.5 w-1.5 rounded-full " + punto[e]} />
      {e}
    </span>
  );
};

const KPI = ({ Icon, tinte, titulo, valor, sub, tag, tagTinte, barra }) => (
  <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
    <div className="flex items-center gap-2.5">
      <div className={"flex h-9 w-9 items-center justify-center rounded-lg " + tinte}>
        <Icon className="h-5 w-5" />
      </div>
      <p className="text-xs font-semibold uppercase tracking-widest text-slate-500">{titulo}</p>
    </div>
    <p className="mt-2 text-2xl font-bold text-slate-900" style={F_MONO}>{valor}</p>
    {barra != null && (
      <div className="mt-2 h-1.5 rounded-full bg-slate-100">
        <div className="h-1.5 rounded-full bg-teal-500" style={{ width: Math.min(100, barra) + "%" }} />
      </div>
    )}
    <div className="mt-2 flex items-center justify-between gap-2">
      <p className="text-xs text-slate-500">{sub}</p>
      {tag && <p className={"text-xs font-semibold uppercase tracking-wide " + (tagTinte || "text-slate-400")}>{tag}</p>}
    </div>
  </div>
);

const ModalBase = ({ children, onClose, bloqueado }) => (
  <div
    className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/60 p-0 sm:items-center sm:p-4"
    onClick={() => { if (!bloqueado) onClose(); }}
  >
    <div
      className="w-full overflow-y-auto rounded-t-2xl bg-white shadow-xl sm:max-w-lg sm:rounded-2xl"
      style={{ maxHeight: "90vh" }}
      onClick={(e) => e.stopPropagation()}
    >
      {children}
    </div>
  </div>
);

const CabeceraModal = ({ titulo, onClose, bloqueado }) => (
  <div className="sticky top-0 z-10 flex items-center justify-between border-b border-slate-200 bg-white px-4 py-3">
    <h2 className="text-base font-bold text-slate-900">{titulo}</h2>
    {!bloqueado && (
      <button onClick={onClose} className="rounded-lg p-1.5 text-slate-400 hover:bg-slate-100 hover:text-slate-700" aria-label="Cerrar">
        <X className="h-5 w-5" />
      </button>
    )}
  </div>
);

const SelectorBarco = ({ valor, onChange, oscuro, ancho }) => (
  <div className={"relative " + (ancho || "")}>
    <select
      value={valor}
      onChange={(e) => onChange(e.target.value)}
      className={oscuro
        ? "w-full appearance-none rounded-lg border border-slate-700 bg-slate-800 py-2 pl-2 pr-7 text-sm font-semibold text-white focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-500"
        : "w-full appearance-none rounded-lg border border-slate-200 bg-white py-1.5 pl-2 pr-7 text-xs font-semibold text-slate-800 shadow-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"}
      aria-label="Embarcación activa"
    >
      {BARCOS.map((b) => <option key={b.id} value={b.id}>{b.nombre}</option>)}
    </select>
    <ChevronDown className={"pointer-events-none absolute right-2 top-1/2 h-3.5 w-3.5 -translate-y-1/2 " + (oscuro ? "text-slate-400" : "text-slate-500")} />
  </div>
);

const FilaItem = ({ req, it, confirmarBorrar, onConciliar, onEliminarItem }) => {
  const pidiendo = confirmarBorrar === "it:" + it.id;
  return (
    <div className={"flex items-start gap-2.5 border-b border-slate-100 px-3 py-2.5 last:border-b-0 " + (it.recibido ? "bg-emerald-50/40" : "")}>
      <button
        onClick={() => onConciliar(req, it)}
        className="mt-0.5 shrink-0 rounded-full text-slate-300 hover:text-emerald-600 focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
        aria-label={it.recibido ? "Editar conciliación" : "Conciliar ítem"}
        title={it.recibido ? "Editar precio / marcar pendiente" : "Marcar recibido y poner precio"}
      >
        {it.recibido ? <CheckCircle2 className="h-5 w-5 text-emerald-600" /> : <Circle className="h-5 w-5" />}
      </button>
      <div className="min-w-0 flex-1">
        <p className="text-sm font-medium leading-snug text-slate-800">{it.descripcion}</p>
        <p className="mt-0.5 text-xs text-slate-500" style={F_MONO}>
          {!it.recibido && (num(it.cantidad) ?? 0) === 0
            ? <span className="rounded bg-slate-100 px-1 py-0.5 font-semibold text-slate-500">NO PEDIDO · {it.unidad}</span>
            : <>Pedido: {it.cantidad} {it.unidad}</>}
          {it.cantidadFacturada != null && <span> · Fact: {it.cantidadFacturada}</span>}
        </p>
        {it.factura && (
          <span className="mt-1 inline-block rounded border border-cyan-200 bg-cyan-50 px-1.5 py-0.5 text-xs text-cyan-800" style={F_MONO}>
            Fact. {it.factura}
          </span>
        )}
      </div>
      <div className="shrink-0 text-right">
        <p className="text-sm font-semibold text-slate-900" style={F_MONO}>{fmtMoney(it.precioTotal)}</p>
        {it.precioUnitario != null && (
          <p className="text-xs text-slate-400" style={F_MONO}>{fmtMoney(it.precioUnitario)} c/u</p>
        )}
      </div>
      <button
        onClick={() => onEliminarItem(req.id, it.id)}
        className={"mt-0.5 shrink-0 rounded p-1 focus:outline-none focus-visible:ring-2 focus-visible:ring-rose-500 " +
          (pidiendo ? "bg-rose-600 text-white" : "text-slate-300 hover:bg-rose-50 hover:text-rose-600")}
        aria-label="Eliminar ítem"
        title={pidiendo ? "Toca otra vez para eliminar" : "Eliminar ítem"}
      >
        <Trash2 className="h-4 w-4" />
      </button>
    </div>
  );
};

const TarjetaReq = ({ r, abierto, onToggle, confirmarBorrar, onEliminar, onConciliar, onEliminarItem, onNuevoItem }) => {
  const exigibles = r.items.filter(esExigible);
  const total = exigibles.length;
  const rec = exigibles.filter((i) => i.recibido).length;
  const monto = round2(r.items.reduce((s, i) => s + (i.recibido ? i.precioTotal || 0 : 0), 0));
  const pidiendo = confirmarBorrar === "req:" + r.id;
  return (
    <div className="overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm">
      <div className="flex items-center gap-2 p-3">
        <button
          onClick={onToggle}
          className="flex min-w-0 flex-1 items-center gap-3 rounded-lg text-left focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
        >
          {abierto ? <ChevronDown className="h-4 w-4 shrink-0 text-slate-400" /> : <ChevronRight className="h-4 w-4 shrink-0 text-slate-400" />}
          <div className="min-w-0 flex-1">
            <div className="flex flex-wrap items-center gap-x-2 gap-y-1">
              <span className="text-sm font-bold text-slate-900" style={F_MONO}>N° {r.numero}</span>
              <Pill e={estadoDe(r)} />
            </div>
            <p className="mt-0.5 truncate text-xs text-slate-500">
              {nombreSemana(r.semana, r.semanaFin)} · {r.archivoNombre} · {fmtFecha(r.fechaSubida)}
            </p>
          </div>
          <div className="shrink-0 text-right">
            <p className="text-sm text-slate-800" style={F_MONO}>{rec}/{total}</p>
            <p className="text-xs font-semibold text-emerald-700" style={F_MONO}>{fmtMoney(monto)}</p>
          </div>
        </button>
        <button
          onClick={() => onEliminar(r.id)}
          className={"shrink-0 rounded-lg p-2 text-xs focus:outline-none focus-visible:ring-2 focus-visible:ring-rose-500 " +
            (pidiendo ? "bg-rose-600 font-semibold text-white" : "text-slate-400 hover:bg-rose-50 hover:text-rose-600")}
          aria-label="Eliminar requisición"
        >
          {pidiendo ? "¿Eliminar?" : <Trash2 className="h-4 w-4" />}
        </button>
      </div>
      <div className="h-1 bg-slate-100">
        <div className="h-1 bg-emerald-500 transition-all" style={{ width: (total ? (rec / total) * 100 : 0) + "%" }} />
      </div>
      {abierto && (
        <div>
          {r.items.map((it) => (
            <FilaItem key={it.id} req={r} it={it} confirmarBorrar={confirmarBorrar} onConciliar={onConciliar} onEliminarItem={onEliminarItem} />
          ))}
          <button
            onClick={() => onNuevoItem(r.id)}
            className="flex w-full items-center justify-center gap-1.5 border-t border-slate-100 px-3 py-2.5 text-xs font-semibold text-cyan-700 hover:bg-cyan-50"
          >
            <Plus className="h-4 w-4" /> Agregar ítem
          </button>
        </div>
      )}
    </div>
  );
};

const DetalleFactura = ({ f, conAcciones, confirmarBorrar, onEliminarLinea, onAgregarItem, onVincularExtra }) => (
  <div className="space-y-3">
    <div className="flex flex-wrap items-center gap-x-3 gap-y-1 text-xs" style={F_MONO}>
      <span className={"inline-flex items-center gap-1 " + (f.verificada ? "text-emerald-700" : "text-slate-400")}>
        <ShieldCheck className="h-3.5 w-3.5" /> {f.verificada ? "Doble verificación" : "Sin verificación"}
      </span>
      <span className="text-slate-500">Capturado: {fmtMoney(round2((f.totalCoincidencias || 0) + (f.totalExtras || 0)))}</span>
      {f.totalFactura != null && (
        <span className={Math.abs((f.totalFactura || 0) - ((f.totalCoincidencias || 0) + (f.totalExtras || 0))) <= 0.05 ? "text-emerald-700" : "text-amber-700"}>
          Total factura: {fmtMoney(f.totalFactura)}
        </span>
      )}
    </div>
    <div>
      <p className="mb-1 text-xs font-semibold uppercase tracking-widest text-emerald-700">
        Coincidencias marcadas ✓ ({(f.coincidencias || []).length}) — {fmtMoney(f.totalCoincidencias)}
      </p>
      {(f.coincidencias || []).length === 0 ? (
        <p className="text-xs text-slate-500">Ningún ítem de esta factura coincidió con requisiciones pendientes.</p>
      ) : (
        <div className="divide-y divide-emerald-50 overflow-hidden rounded-lg border border-emerald-100">
          {f.coincidencias.map((c, i) => {
            const pidiendo = confirmarBorrar === "lf:" + f.id + ":m:" + i;
            return (
              <div key={i} className="flex items-center gap-2 bg-emerald-50/40 px-2.5 py-1.5">
                <CheckCircle2 className="h-4 w-4 shrink-0 text-emerald-600" />
                <div className="min-w-0 flex-1">
                  <p className="truncate text-xs font-medium text-slate-800">{c.descripcion}</p>
                  <p className="text-xs text-slate-500" style={F_MONO}>
                    {c.cantidad} {c.unidad} · Req. {c.reqNumero} ({depInfo(c.departamento).corto})
                  </p>
                </div>
                <p className="shrink-0 text-xs font-semibold text-slate-900" style={F_MONO}>{fmtMoney(c.precioTotal)}</p>
                {conAcciones && (
                  <button
                    onClick={() => onEliminarLinea(f.id, "m", i)}
                    className={"shrink-0 rounded p-1 " + (pidiendo ? "bg-rose-600 text-white" : "text-slate-300 hover:text-rose-600")}
                    title={pidiendo ? "Toca otra vez: elimina la línea y regresa el ítem a pendiente" : "Eliminar línea"}
                    aria-label="Eliminar línea"
                  >
                    <X className="h-3.5 w-3.5" />
                  </button>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
    <div>
      <p className="mb-1 text-xs font-semibold uppercase tracking-widest text-amber-700">
        Extras fuera de requisición ({(f.extras || []).length}) — {fmtMoney(f.totalExtras)}
      </p>
      {(f.extras || []).length === 0 ? (
        <p className="text-xs text-slate-500">Todos los ítems de la factura estaban en las requisiciones.</p>
      ) : (
        <div className="divide-y divide-amber-50 overflow-hidden rounded-lg border border-amber-100">
          {f.extras.map((x, i) => {
            const pidiendo = confirmarBorrar === "lf:" + f.id + ":x:" + i;
            return (
              <div key={i} className="flex items-center gap-2 bg-amber-50/40 px-2.5 py-1.5">
                <AlertTriangle className="h-4 w-4 shrink-0 text-amber-500" />
                <div className="min-w-0 flex-1">
                  <p className="truncate text-xs font-medium text-slate-800">{x.descripcion}</p>
                  <p className="text-xs text-slate-500" style={F_MONO}>{x.cantidad} {x.unidad}</p>
                </div>
                <p className="shrink-0 text-xs font-semibold text-slate-900" style={F_MONO}>{fmtMoney(x.precioTotal)}</p>
                {conAcciones && onVincularExtra && (
                  <button
                    onClick={() => onVincularExtra(f.id, i)}
                    className="shrink-0 rounded-md border border-cyan-200 bg-cyan-50 px-1.5 py-1 text-xs font-semibold text-cyan-800 hover:bg-cyan-100"
                    title="Agregar este ítem a una requisición (con sugerencia de Claude)"
                  >
                    + Req.
                  </button>
                )}
                {conAcciones && (
                  <button
                    onClick={() => onEliminarLinea(f.id, "x", i)}
                    className={"shrink-0 rounded p-1 " + (pidiendo ? "bg-rose-600 text-white" : "text-slate-300 hover:text-rose-600")}
                    title={pidiendo ? "Toca otra vez para eliminar" : "Eliminar línea"}
                    aria-label="Eliminar línea"
                  >
                    <X className="h-3.5 w-3.5" />
                  </button>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
    {conAcciones && (
      <button
        onClick={() => onAgregarItem(f.id)}
        className="flex w-full items-center justify-center gap-1.5 rounded-lg border border-dashed border-slate-300 px-3 py-2 text-xs font-semibold text-cyan-700 hover:border-cyan-500 hover:bg-cyan-50"
      >
        <Plus className="h-4 w-4" /> Agregar ítem manual
      </button>
    )}
  </div>
);

const TarjetaFact = ({ f, abierto, onToggle, confirmarBorrar, onEliminar, onEliminarLinea, onAgregarItem, onVincularExtra }) => {
  const pidiendo = confirmarBorrar === "fact:" + f.id;
  return (
    <div className="overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm">
      <div className="flex items-center gap-2 p-3">
        <button
          onClick={onToggle}
          className="flex min-w-0 flex-1 items-center gap-3 rounded-lg text-left focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
        >
          {abierto ? <ChevronDown className="h-4 w-4 shrink-0 text-slate-400" /> : <ChevronRight className="h-4 w-4 shrink-0 text-slate-400" />}
          <Receipt className="h-4 w-4 shrink-0 text-cyan-700" />
          <div className="min-w-0 flex-1">
            <p className="truncate text-sm font-bold text-slate-900" style={F_MONO}>Fact. {f.numero}</p>
            <p className="mt-0.5 truncate text-xs text-slate-500">{f.proveedor} · {f.fecha} · {f.archivoNombre}</p>
          </div>
          <div className="shrink-0 text-right">
            <p className="text-xs text-emerald-700" style={F_MONO}>✓ {(f.coincidencias || []).length}</p>
            <p className="text-xs text-amber-700" style={F_MONO}>+ {(f.extras || []).length}</p>
          </div>
        </button>
        <button
          onClick={() => onEliminar(f.id)}
          className={"shrink-0 rounded-lg p-2 text-xs focus:outline-none focus-visible:ring-2 focus-visible:ring-rose-500 " +
            (pidiendo ? "bg-rose-600 font-semibold text-white" : "text-slate-400 hover:bg-rose-50 hover:text-rose-600")}
          aria-label="Eliminar factura"
        >
          {pidiendo ? "¿Eliminar?" : <Trash2 className="h-4 w-4" />}
        </button>
      </div>
      {abierto && (
        <div className="border-t border-slate-100 px-3 py-3">
          <DetalleFactura f={f} conAcciones confirmarBorrar={confirmarBorrar} onEliminarLinea={onEliminarLinea} onAgregarItem={onAgregarItem} onVincularExtra={onVincularExtra} />
        </div>
      )}
    </div>
  );
};

/* ============================================================ */

export default function ExplorerFleetControl() {
  const [barco, setBarco] = useState("humboldt");
  const [vista, setVista] = useState("resumen");
  const [cargando, setCargando] = useState(true);
  const [reqs, setReqs] = useState([]);
  const [facts, setFacts] = useState([]);

  const [expandidos, setExpandidos] = useState({});
  const [expandidosF, setExpandidosF] = useState({});
  const [filtroSemana, setFiltroSemana] = useState("todas");
  const [filtroDep, setFiltroDep] = useState("todos");
  const [confirmarBorrar, setConfirmarBorrar] = useState(null);

  const [catModo, setCatModo] = useState("prov");
  const [catBusqueda, setCatBusqueda] = useState("");
  const [catAbierto, setCatAbierto] = useState(null);

  const [modalReq, setModalReq] = useState(null);
  const [modalFact, setModalFact] = useState(null);
  const [modalCon, setModalCon] = useState(null);
  const [modalNuevo, setModalNuevo] = useState(null);
  const [modalItemFact, setModalItemFact] = useState(null);
  const [modalVincular, setModalVincular] = useState(null);
  const [modalInstalar, setModalInstalar] = useState(false);

  const inputReqRef = useRef(null);
  const inputFactRef = useRef(null);

  const barcoInfo = BARCOS.find((b) => b.id === barco) || BARCOS[0];
  const kReq = (id) => claveReq(barco, id);
  const kFact = (id) => claveFact(barco, id);

  /* ------------ Carga de datos ------------ */

  const cargar = async (barcoId) => {
    setCargando(true);
    const [r0, f0] = await Promise.all([cargarRegistros("req", barcoId), cargarRegistros("fact", barcoId)]);
    const r = r0.map(migrarSemana);
    r.forEach((reg, i) => { if (reg !== r0[i]) persistir(claveReq(barcoId, reg.id), reg); });
    r.sort((a, b) => (b.fechaSubida || "").localeCompare(a.fechaSubida || ""));
    f0.sort((a, b) => (b.fechaSubida || "").localeCompare(a.fechaSubida || ""));
    setReqs(r);
    setFacts(f0);
    setCargando(false);
  };

  useEffect(() => {
    (async () => {
      let b = "humboldt";
      try {
        const obj = await obtenerConfig();
        if (obj && BARCO_IDS.includes(obj.barco)) b = obj.barco;
      } catch (e) { /* primera vez */ }
      setBarco(b);
      await cargar(b);
    })();
  }, []);

  useEffect(() => {
    if (!confirmarBorrar) return;
    const t = setTimeout(() => setConfirmarBorrar(null), 3500);
    return () => clearTimeout(t);
  }, [confirmarBorrar]);

  const cambiarBarco = async (id) => {
    if (id === barco) return;
    setBarco(id);
    setModalReq(null); setModalFact(null); setModalCon(null); setModalNuevo(null); setModalItemFact(null); setModalVincular(null);
    setExpandidos({}); setExpandidosF({});
    setFiltroSemana("todas"); setFiltroDep("todos");
    setCatBusqueda(""); setCatAbierto(null);
    persistir("hx-config", { barco: id });
    await cargar(id);
  };

  /* ------------ Métricas ------------ */

  const stats = useMemo(() => {
    let pend = 0, rec = 0;
    reqs.forEach((r) => (r.items || []).forEach((it) => {
      if (it.recibido) rec++;
      else if (esExigible(it)) pend++;
    }));
    const abiertas = reqs.filter((r) => (r.items || []).some((i) => !i.recibido && esExigible(i))).length;
    const compras = facts.reduce((s, f) => s + (f.totalCoincidencias || 0) + (f.totalExtras || 0), 0);
    const pct = rec + pend > 0 ? Math.round((rec / (rec + pend)) * 100) : 0;
    return { total: reqs.length, abiertas, pend, rec, pct, compras: round2(compras) };
  }, [reqs, facts]);

  const extrasTotal = useMemo(() => facts.reduce((s, f) => s + ((f.extras || []).length), 0), [facts]);

  const semanasDisponibles = useMemo(() => {
    const mapa = new Map();
    reqs.forEach((r) => {
      const clave = claveSemana(r);
      if (!mapa.has(clave)) mapa.set(clave, { clave, inicio: r.semana, fin: r.semanaFin });
    });
    return Array.from(mapa.values()).sort(
      (a, b) => String(a.inicio).localeCompare(String(b.inicio)) || String(a.fin || "").localeCompare(String(b.fin || ""))
    );
  }, [reqs]);

  const grupos = useMemo(() => {
    const base = reqs
      .filter((r) => filtroSemana === "todas" || claveSemana(r) === filtroSemana)
      .filter((r) => filtroDep === "todos" || r.departamento === filtroDep);
    return DEPARTAMENTOS.map((d) => ({
      dep: d,
      lista: base
        .filter((r) => r.departamento === d.id)
        .sort((a, b) => String(b.semana).localeCompare(String(a.semana)) || (b.fechaSubida || "").localeCompare(a.fechaSubida || "")),
    })).filter((g) => g.lista.length > 0);
  }, [reqs, filtroSemana, filtroDep]);

  const cargaDeps = useMemo(() => {
    const arr = DEPARTAMENTOS.map((d) => ({
      dep: d,
      items: reqs.filter((r) => r.departamento === d.id).reduce((s, r) => s + (r.items || []).filter(esExigible).length, 0),
    })).filter((a) => a.items > 0);
    const max = Math.max(1, ...arr.map((a) => a.items), 1);
    return { arr, max };
  }, [reqs]);

  /* Base de datos de compras: cada línea de factura como movimiento */
  const movimientos = useMemo(() => {
    const lista = [];
    facts.forEach((f) => {
      const base = { factId: f.id, factNumero: f.numero, proveedor: f.proveedor || "—", fecha: f.fecha, fechaSubida: f.fechaSubida };
      (f.coincidencias || []).forEach((c) =>
        lista.push({ ...base, descripcion: c.descripcion, cantidad: c.cantidad, unidad: c.unidad, precioUnitario: c.precioUnitario, precioTotal: c.precioTotal, reqNumero: c.reqNumero })
      );
      (f.extras || []).forEach((x) =>
        lista.push({ ...base, descripcion: x.descripcion, cantidad: x.cantidad, unidad: x.unidad, precioUnitario: x.precioUnitario, precioTotal: x.precioTotal, reqNumero: null })
      );
    });
    return lista.sort((a, b) => (b.fechaSubida || "").localeCompare(a.fechaSubida || ""));
  }, [facts]);

  const proveedoresDB = useMemo(() => {
    const m = new Map();
    facts.forEach((f) => {
      const k = f.proveedor || "—";
      const o = m.get(k) || { proveedor: k, n: 0, total: 0, ultima: "", facturas: [] };
      o.n++;
      o.total += (f.totalCoincidencias || 0) + (f.totalExtras || 0);
      if ((f.fechaSubida || "") > o.ultima) o.ultima = f.fechaSubida || "";
      o.facturas.push(f);
      m.set(k, o);
    });
    const arr = Array.from(m.values());
    arr.forEach((o) => o.facturas.sort((a, b) => (b.fechaSubida || "").localeCompare(a.fechaSubida || "")));
    return arr.sort((a, b) => b.total - a.total);
  }, [facts]);

  const itemsDB = useMemo(() => {
    const m = new Map();
    movimientos.forEach((mv) => {
      const k = mv.descripcion;
      if (!k) return;
      const o = m.get(k) || { descripcion: k, unidad: mv.unidad || "UND", compras: 0, cant: 0, gasto: 0, ultima: "", hist: [] };
      o.compras++;
      o.cant += mv.cantidad || 0;
      o.gasto += mv.precioTotal || 0;
      if ((mv.fechaSubida || "") > o.ultima) o.ultima = mv.fechaSubida || "";
      o.hist.push(mv);
      m.set(k, o);
    });
    return Array.from(m.values()).sort((a, b) => b.gasto - a.gasto);
  }, [movimientos]);

  const pendientesGlobal = useMemo(() => {
    const lista = [];
    reqs.forEach((r) => (r.items || []).forEach((it) => { if (!it.recibido) lista.push({ req: r, it }); }));
    return lista;
  }, [reqs]);

  /* ------------ Flujo: subir pedido(s) / requisición(es) ------------ */

  const abrirModalReq = () => {
    const base = semanaActual();
    setModalReq({
      paso: "form", archivos: [], modoFecha: "semana", semana: base,
      fechaInicio: base, fechaFin: sumarDias(base, 7),
      resultados: [], error: null, progreso: "",
    });
  };

  const analizarRequisiciones = async () => {
    const m = modalReq;
    if (!m.archivos.length) return setModalReq({ ...m, error: "Selecciona al menos un archivo de pedido o requisición." });
    const sinNumero = m.archivos.find((a) => !a.numero.trim());
    if (sinNumero) return setModalReq({ ...m, error: "Falta el número de requisición de \"" + sinNumero.archivo.name + "\"." });
    if (m.modoFecha === "personalizado") {
      if (!esISO(m.fechaInicio) || !esISO(m.fechaFin)) return setModalReq({ ...m, error: "Selecciona las fechas de inicio y fin del crucero." });
      if (m.fechaFin <= m.fechaInicio) return setModalReq({ ...m, error: "La fecha de fin debe ser posterior a la de inicio." });
    }
    setModalReq({ ...m, paso: "procesando", error: null, progreso: "" });
    const resultados = [];
    for (let i = 0; i < m.archivos.length; i++) {
      const af = m.archivos[i];
      const etiqueta = m.archivos.length > 1
        ? "Requisición " + (i + 1) + " de " + m.archivos.length + " (N° " + af.numero.trim() + "): "
        : "";
      try {
        const blocks = await fileToBlocks(af.archivo);
        const { items, verificacion } = await extraerItemsVerificado(blocks, (msg) =>
          setModalReq((prev) => (prev ? { ...prev, progreso: etiqueta + msg } : prev))
        );
        if (!items.length) throw new Error("No se detectaron ítems en el documento.");
        resultados.push({
          fileId: af.id, archivoNombre: af.archivo.name, numero: af.numero.trim(), departamento: af.departamento,
          items, verificacion, error: null, nDesc: "", nCant: "", nUnidad: "",
        });
      } catch (e) {
        resultados.push({
          fileId: af.id, archivoNombre: af.archivo.name, numero: af.numero.trim(), departamento: af.departamento,
          items: [], verificacion: null, error: e.message || "No se pudo leer este documento.",
          nDesc: "", nCant: "", nUnidad: "",
        });
      }
    }
    setModalReq((prev) => (prev ? { ...prev, paso: "preview", resultados, error: null, progreso: "" } : prev));
  };

  const quitarArchivoReq = (id) =>
    setModalReq((prev) => ({ ...prev, archivos: prev.archivos.filter((a) => a.id !== id) }));

  const actualizarArchivoReq = (id, cambios) =>
    setModalReq((prev) => ({ ...prev, archivos: prev.archivos.map((a) => (a.id === id ? { ...a, ...cambios } : a)) }));

  const actualizarResultadoReq = (fileId, cambios) =>
    setModalReq((prev) => ({ ...prev, resultados: prev.resultados.map((r) => (r.fileId === fileId ? { ...r, ...cambios } : r)) }));

  const quitarItemResultadoReq = (fileId, itemId) =>
    setModalReq((prev) => ({
      ...prev,
      resultados: prev.resultados.map((r) => (r.fileId === fileId ? { ...r, items: r.items.filter((it) => it.id !== itemId) } : r)),
    }));

  const agregarItemResultadoReq = (fileId) => {
    const r = (modalReq.resultados || []).find((x) => x.fileId === fileId);
    if (!r) return;
    const d = (r.nDesc || "").trim();
    if (!d) return actualizarResultadoReq(fileId, { error: "Escribe la descripción del ítem a agregar." });
    const it = {
      id: uid(), descripcion: d.toUpperCase(), cantidad: num(r.nCant) ?? 1,
      unidad: ((r.nUnidad || "UND").toUpperCase().trim()) || "UND",
      recibido: false, cantidadFacturada: null, precioUnitario: null, precioTotal: null, factura: null,
    };
    actualizarResultadoReq(fileId, { items: [...r.items, it], nDesc: "", nCant: "", nUnidad: "", error: null });
  };

  const guardarRequisiciones = async () => {
    const m = modalReq;
    const personalizado = m.modoFecha === "personalizado";
    const semanaIni = personalizado ? m.fechaInicio : m.semana;
    const semanaFin = personalizado ? m.fechaFin : sumarDias(m.semana, 7);
    const validos = m.resultados.filter((r) => r.items.length > 0);
    if (!validos.length) return;
    const nuevas = validos.map((r) => ({
      id: uid(), numero: r.numero.trim(), semana: semanaIni, semanaFin,
      departamento: r.departamento, archivoNombre: r.archivoNombre,
      fechaSubida: new Date().toISOString(), items: r.items,
    }));
    for (const n of nuevas) await persistir(kReq(n.id), n);
    setReqs((prev) => [...nuevas, ...prev]);
    setExpandidos((prev) => {
      const nx = { ...prev };
      nuevas.forEach((n) => { nx[n.id] = true; });
      return nx;
    });
    setVista("pedidos");
    setModalReq(null);
  };

  /* ------------ Flujo: subir factura(s) con doble verificación ------------ */

  const abrirModalFact = () =>
    setModalFact({ paso: "form", archivos: [], objetivo: "todas", resultados: [], error: null, progreso: "" });

  const procesarFacturas = async () => {
    const m = modalFact;
    if (!m.archivos.length) return setModalFact({ ...m, error: "Selecciona al menos un archivo de factura." });
    setModalFact((prev) => ({ ...prev, paso: "procesando", error: null, progreso: "" }));
    try {
      const working = reqs.map((r) => ({ ...r, items: r.items.map((it) => ({ ...it })) }));
      const tocadas = new Set();
      const nuevasFacturas = [];

      for (let i = 0; i < m.archivos.length; i++) {
        const file = m.archivos[i];
        const etiqueta = m.archivos.length > 1 ? "Factura " + (i + 1) + " de " + m.archivos.length + ": " : "";
        setModalFact((prev) => (prev ? { ...prev, progreso: etiqueta + "leyendo " + file.name + "…" } : prev));

        const fuente = m.objetivo === "todas" ? working : working.filter((r) => r.id === m.objetivo);
        const refs = [];
        fuente.forEach((r) => r.items.forEach((it) => { if (!it.recibido) refs.push({ req: r, it }); }));
        const compact = refs.slice(0, 600).map((p, idx) => [idx, p.it.descripcion, p.it.cantidad, p.it.unidad]);

        const blocks = await fileToBlocks(file);
        const texto = await callClaude([...blocks, { type: "text", text: promptFactura(compact) }], 4000);
        const { data: parsed } = parseJSONSafe(texto);

        const coincidencias = [];
        const refsUsados = [];
        (Array.isArray(parsed.m) ? parsed.m : []).forEach((row) => {
          const ref = refs[row[0]];
          if (!ref || ref.it.recibido) return;
          const cant = num(row[1]) ?? ref.it.cantidad;
          let pu = num(row[2]);
          let pt = num(row[3]);
          if (pt == null && pu != null) pt = round2(cant * pu);
          if (pu == null && pt != null && cant) pu = round2(pt / cant);
          ref.it.recibido = true;
          ref.it.cantidadFacturada = cant;
          ref.it.precioUnitario = pu;
          ref.it.precioTotal = pt;
          ref.it.factura = parsed.num ? String(parsed.num) : file.name;
          tocadas.add(ref.req.id);
          refsUsados.push(ref);
          coincidencias.push({
            reqId: ref.req.id, itemId: ref.it.id, reqNumero: ref.req.numero, departamento: ref.req.departamento,
            descripcion: ref.it.descripcion, cantidad: cant, unidad: ref.it.unidad, precioUnitario: pu, precioTotal: pt,
          });
        });

        const extras = (Array.isArray(parsed.x) ? parsed.x : []).map((row) => {
          const cant = num(row[1]) ?? 1;
          const pu = num(row[3]);
          let pt = num(row[4]);
          if (pt == null && pu != null) pt = round2(cant * pu);
          return {
            descripcion: String(row[0] || "").toUpperCase(), cantidad: cant,
            unidad: String(row[2] || "UND").toUpperCase(), precioUnitario: pu, precioTotal: pt,
          };
        }).filter((x) => x.descripcion);

        /* --- Doble verificación de la factura --- */
        let totalFactura = num(parsed.total);
        let verificada = false;
        try {
          setModalFact((prev) => (prev ? { ...prev, progreso: etiqueta + "doble verificación de " + file.name + "…" } : prev));
          const lineas = [
            ...coincidencias.map((c, idx) => ["M" + idx, c.descripcion, c.cantidad, c.precioUnitario, c.precioTotal]),
            ...extras.map((x, idx) => ["X" + idx, x.descripcion, x.cantidad, x.precioUnitario, x.precioTotal]),
          ];
          const t2 = await callClaude([...blocks, { type: "text", text: promptVerificacionFactura(lineas) }], 3000);
          const { data: v } = parseJSONSafe(t2);
          (Array.isArray(v.correcciones) ? v.correcciones : []).forEach((c) => {
            const cod = String(c[0] || "");
            const idx = parseInt(cod.slice(1), 10);
            const cant = num(c[1]);
            let pu = num(c[2]);
            let pt = num(c[3]);
            if (cod.startsWith("M") && coincidencias[idx]) {
              const co = coincidencias[idx];
              if (cant != null) co.cantidad = cant;
              if (pu != null) co.precioUnitario = pu;
              if (pt == null && pu != null) pt = round2(co.cantidad * pu);
              if (pt != null) co.precioTotal = pt;
              const ref = refsUsados[idx];
              if (ref) {
                ref.it.cantidadFacturada = co.cantidad;
                ref.it.precioUnitario = co.precioUnitario;
                ref.it.precioTotal = co.precioTotal;
              }
            } else if (cod.startsWith("X") && extras[idx]) {
              const ex = extras[idx];
              if (cant != null) ex.cantidad = cant;
              if (pu != null) ex.precioUnitario = pu;
              if (pt == null && pu != null) pt = round2(ex.cantidad * pu);
              if (pt != null) ex.precioTotal = pt;
            }
          });
          (Array.isArray(v.faltantes) ? v.faltantes : []).forEach((row) => {
            const cant = num(row[1]) ?? 1;
            const pu = num(row[3]);
            let pt = num(row[4]);
            if (pt == null && pu != null) pt = round2(cant * pu);
            const d = String(row[0] || "").toUpperCase();
            if (d) extras.push({ descripcion: d, cantidad: cant, unidad: String(row[2] || "UND").toUpperCase(), precioUnitario: pu, precioTotal: pt });
          });
          if (num(v.total) != null) totalFactura = num(v.total);
          verificada = true;
        } catch (e) { /* si la verificación falla, se conserva la primera pasada */ }

        nuevasFacturas.push({
          id: uid(),
          numero: parsed.num ? String(parsed.num) : file.name,
          proveedor: parsed.prov ? String(parsed.prov) : "—",
          fecha: parsed.fecha ? String(parsed.fecha) : "—",
          archivoNombre: file.name,
          fechaSubida: new Date().toISOString(),
          coincidencias, extras, verificada, totalFactura,
          totalCoincidencias: round2(coincidencias.reduce((s, c) => s + (c.precioTotal || 0), 0)),
          totalExtras: round2(extras.reduce((s, c) => s + (c.precioTotal || 0), 0)),
        });
      }

      for (const f of nuevasFacturas) await persistir(kFact(f.id), f);
      for (const idReq of tocadas) {
        const r = working.find((w) => w.id === idReq);
        if (r) await persistir(kReq(r.id), r);
      }
      setReqs(working);
      setFacts((prev) => [...nuevasFacturas, ...prev]);
      setModalFact((prev) => ({ ...prev, paso: "resultados", resultados: nuevasFacturas, progreso: "" }));
    } catch (e) {
      setModalFact((prev) => ({ ...prev, paso: "form", error: e.message, progreso: "" }));
    }
  };

  /* ------------ Conciliación y ediciones manuales ------------ */

  const actualizarItem = async (reqId, itemId, cambios) => {
    let regFinal = null;
    const actualizadas = reqs.map((r) => {
      if (r.id !== reqId) return r;
      regFinal = { ...r, items: r.items.map((it) => (it.id !== itemId ? it : { ...it, ...cambios })) };
      return regFinal;
    });
    setReqs(actualizadas);
    if (regFinal) await persistir(kReq(reqId), regFinal);
  };

  const abrirConciliar = (r, it) =>
    setModalCon({
      reqId: r.id, itemId: it.id, descripcion: it.descripcion, unidad: it.unidad,
      recibido: it.recibido, pedido: it.cantidad,
      cant: it.cantidadFacturada != null ? String(it.cantidadFacturada) : String(it.cantidad),
      precio: it.precioUnitario != null ? String(it.precioUnitario) : "",
      fact: it.factura || "",
    });

  const guardarConciliacion = async () => {
    const m = modalCon;
    const cant = num(m.cant) ?? m.pedido;
    const pu = num(m.precio);
    const pt = pu != null ? round2(cant * pu) : null;
    await actualizarItem(m.reqId, m.itemId, {
      recibido: true, cantidadFacturada: cant, precioUnitario: pu, precioTotal: pt,
      factura: m.fact.trim() || null,
    });
    setModalCon(null);
  };

  const marcarPendiente = async () => {
    const m = modalCon;
    await actualizarItem(m.reqId, m.itemId, {
      recibido: false, cantidadFacturada: null, precioUnitario: null, precioTotal: null, factura: null,
    });
    setModalCon(null);
  };

  const abrirNuevoItem = (reqId) => setModalNuevo({ reqId, desc: "", cant: "", unidad: "" });

  const guardarNuevoItem = async () => {
    const m = modalNuevo;
    const d = (m.desc || "").trim();
    if (!d) return setModalNuevo({ ...m, error: "Escribe la descripción del ítem." });
    const it = {
      id: uid(), descripcion: d.toUpperCase(), cantidad: num(m.cant) ?? 1,
      unidad: ((m.unidad || "UND").toUpperCase().trim()) || "UND",
      recibido: false, cantidadFacturada: null, precioUnitario: null, precioTotal: null, factura: null,
    };
    let regFinal = null;
    const actualizadas = reqs.map((r) => {
      if (r.id !== m.reqId) return r;
      regFinal = { ...r, items: [...r.items, it] };
      return regFinal;
    });
    setReqs(actualizadas);
    if (regFinal) await persistir(kReq(m.reqId), regFinal);
    setModalNuevo(null);
  };

  const recalcularTotalesFactura = (f) => ({
    ...f,
    totalCoincidencias: round2((f.coincidencias || []).reduce((s, c) => s + (c.precioTotal || 0), 0)),
    totalExtras: round2((f.extras || []).reduce((s, c) => s + (c.precioTotal || 0), 0)),
  });

  const reemplazarFactura = async (nuevaF) => {
    setFacts((prev) => prev.map((f) => (f.id === nuevaF.id ? nuevaF : f)));
    setModalFact((prev) =>
      prev && prev.paso === "resultados"
        ? { ...prev, resultados: prev.resultados.map((f) => (f.id === nuevaF.id ? nuevaF : f)) }
        : prev
    );
    await persistir(kFact(nuevaF.id), nuevaF);
  };

  const abrirItemFactura = (factId) =>
    setModalItemFact({ factId, vinculo: "extra", desc: "", cant: "", unidad: "", precio: "", error: null });

  const guardarItemFactura = async () => {
    const m = modalItemFact;
    const fact = facts.find((f) => f.id === m.factId);
    if (!fact) return setModalItemFact(null);
    const pu = num(m.precio);

    if (m.vinculo === "extra") {
      const d = (m.desc || "").trim();
      if (!d) return setModalItemFact({ ...m, error: "Escribe la descripción del ítem." });
      const cant = num(m.cant) ?? 1;
      const pt = pu != null ? round2(cant * pu) : null;
      const nuevaF = recalcularTotalesFactura({
        ...fact,
        extras: [...(fact.extras || []), {
          descripcion: d.toUpperCase(), cantidad: cant,
          unidad: ((m.unidad || "UND").toUpperCase().trim()) || "UND", precioUnitario: pu, precioTotal: pt,
        }],
      });
      await reemplazarFactura(nuevaF);
      setModalItemFact(null);
      return;
    }

    const [reqId, itemId] = String(m.vinculo).split("::");
    const req = reqs.find((r) => r.id === reqId);
    const it = req && req.items.find((i) => i.id === itemId);
    if (!req || !it) return setModalItemFact({ ...m, error: "El ítem seleccionado ya no está disponible." });
    const cant = num(m.cant) ?? it.cantidad;
    const pt = pu != null ? round2(cant * pu) : null;
    await actualizarItem(reqId, itemId, {
      recibido: true, cantidadFacturada: cant, precioUnitario: pu, precioTotal: pt, factura: fact.numero,
    });
    const nuevaF = recalcularTotalesFactura({
      ...fact,
      coincidencias: [...(fact.coincidencias || []), {
        reqId, itemId, reqNumero: req.numero, departamento: req.departamento,
        descripcion: it.descripcion, cantidad: cant, unidad: it.unidad, precioUnitario: pu, precioTotal: pt,
      }],
    });
    await reemplazarFactura(nuevaF);
    setModalItemFact(null);
  };

  /* ------------ Vincular un extra de factura con un ítem pendiente ------------ */

  const abrirVincularExtra = (factId, extraIdx) => {
    const fact = facts.find((f) => f.id === factId);
    const extra = fact && (fact.extras || [])[extraIdx];
    if (!fact || !extra) return;
    const candidatos = pendientesGlobal.slice(0, 400);
    setModalVincular({
      factId, extraIdx, extra, seleccion: "", error: null,
      cargandoSugerencia: candidatos.length > 0, sugerencia: null,
    });
    if (!candidatos.length) return;
    const compact = candidatos.map(({ req, it }, idx) =>
      [idx, it.descripcion, it.cantidad, it.unidad, req.numero, depInfo(req.departamento).corto]);
    (async () => {
      let sug = null;
      try {
        const texto = await callClaude([{ type: "text", text: promptSugerenciaVinculo(extra, compact) }], 300);
        const { data } = parseJSONSafe(texto);
        const idx = num(data.indice);
        const cand = idx != null ? candidatos[idx] : null;
        if (cand && !cand.it.recibido) sug = { reqId: cand.req.id, itemId: cand.it.id, motivo: String(data.motivo || "").trim() };
      } catch (e) { /* sin sugerencia: se elige manual */ }
      setModalVincular((prev) => {
        if (!prev || prev.factId !== factId || prev.extraIdx !== extraIdx) return prev;
        return {
          ...prev, cargandoSugerencia: false, sugerencia: sug,
          seleccion: prev.seleccion || (sug ? sug.reqId + "::" + sug.itemId : ""),
        };
      });
    })();
  };

  const confirmarVinculoExtra = async () => {
    const m = modalVincular;
    if (!m) return;
    if (!m.seleccion) return setModalVincular({ ...m, error: "Elige el ítem pendiente al que corresponde." });
    const fact = facts.find((f) => f.id === m.factId);
    const extra = fact && (fact.extras || [])[m.extraIdx];
    if (!fact || !extra || extra.descripcion !== m.extra.descripcion) {
      return setModalVincular({ ...m, error: "La línea de la factura ya no está disponible." });
    }
    const [reqId, itemId] = String(m.seleccion).split("::");
    const req = reqs.find((r) => r.id === reqId);
    const it = req && req.items.find((i) => i.id === itemId);
    if (!req || !it || it.recibido) {
      return setModalVincular({ ...m, error: "El ítem seleccionado ya no está pendiente. Elige otro." });
    }
    const cant = extra.cantidad != null ? extra.cantidad : it.cantidad;
    let pu = extra.precioUnitario;
    let pt = extra.precioTotal;
    if (pt == null && pu != null) pt = round2(cant * pu);
    if (pu == null && pt != null && cant) pu = round2(pt / cant);
    await actualizarItem(reqId, itemId, {
      recibido: true, cantidadFacturada: cant, precioUnitario: pu, precioTotal: pt, factura: fact.numero,
    });
    const nuevaF = recalcularTotalesFactura({
      ...fact,
      extras: (fact.extras || []).filter((_, i) => i !== m.extraIdx),
      coincidencias: [...(fact.coincidencias || []), {
        reqId, itemId, reqNumero: req.numero, departamento: req.departamento,
        descripcion: it.descripcion, cantidad: cant, unidad: it.unidad, precioUnitario: pu, precioTotal: pt,
      }],
    });
    await reemplazarFactura(nuevaF);
    setModalVincular(null);
  };

  /* ------------ Eliminaciones (confirmación de dos toques) ------------ */

  const pedirConfirmacion = (clave) => {
    if (confirmarBorrar !== clave) {
      setConfirmarBorrar(clave);
      return false;
    }
    setConfirmarBorrar(null);
    return true;
  };

  const eliminarReq = async (reqId) => {
    if (!pedirConfirmacion("req:" + reqId)) return;
    setReqs((prev) => prev.filter((r) => r.id !== reqId));
    try { await eliminar(kReq(reqId)); } catch (e) { /* silencioso */ }
  };

  const eliminarFact = async (factId) => {
    if (!pedirConfirmacion("fact:" + factId)) return;
    setFacts((prev) => prev.filter((f) => f.id !== factId));
    try { await eliminar(kFact(factId)); } catch (e) { /* silencioso */ }
  };

  const eliminarItemReq = async (reqId, itemId) => {
    if (!pedirConfirmacion("it:" + itemId)) return;
    let regFinal = null;
    const actualizadas = reqs.map((r) => {
      if (r.id !== reqId) return r;
      regFinal = { ...r, items: r.items.filter((it) => it.id !== itemId) };
      return regFinal;
    });
    setReqs(actualizadas);
    if (regFinal) await persistir(kReq(reqId), regFinal);
  };

  const eliminarLineaFactura = async (factId, tipo, idx) => {
    if (!pedirConfirmacion("lf:" + factId + ":" + tipo + ":" + idx)) return;
    const fact = facts.find((f) => f.id === factId);
    if (!fact) return;
    if (tipo === "x") {
      const nuevaF = recalcularTotalesFactura({ ...fact, extras: (fact.extras || []).filter((_, i) => i !== idx) });
      await reemplazarFactura(nuevaF);
      return;
    }
    const c = (fact.coincidencias || [])[idx];
    const nuevaF = recalcularTotalesFactura({ ...fact, coincidencias: (fact.coincidencias || []).filter((_, i) => i !== idx) });
    await reemplazarFactura(nuevaF);
    if (!c) return;
    let req = c.reqId ? reqs.find((r) => r.id === c.reqId) : null;
    let it = req && c.itemId ? req.items.find((i) => i.id === c.itemId) : null;
    if (!it) {
      req = reqs.find((r) => r.numero === c.reqNumero);
      it = req && req.items.find((i) => i.descripcion === c.descripcion && i.recibido);
    }
    if (req && it) {
      await actualizarItem(req.id, it.id, {
        recibido: false, cantidadFacturada: null, precioUnitario: null, precioTotal: null, factura: null,
      });
    }
  };

  const irARequisicion = (r) => {
    setVista("pedidos");
    setFiltroSemana("todas");
    setFiltroDep("todos");
    setExpandidos({ [r.id]: true });
  };

  /* ============================ Vistas ============================ */

  const renderResumen = () => {
    const recientes = reqs.slice(0, 6);
    return (
      <div className="space-y-4">
        <div className="rounded-2xl border border-slate-200 border-l-4 border-l-cyan-600 bg-white p-5 shadow-sm">
          <p className="text-xs font-bold uppercase tracking-widest text-cyan-700">
            Control semanal · Semana {numeroSemana(semanaActual())} · {rangoCorto(semanaActual())}
          </p>
          <div className="mt-1 flex flex-wrap items-end justify-between gap-3">
            <div>
              <h2 className="text-xl font-bold text-slate-900 sm:text-2xl">{barcoInfo.nombre}: todo el abastecimiento, en una sola vista.</h2>
              <p className="mt-1 text-sm text-slate-500">
                Sube documentos, revisa los ítems extraídos y concilia cada compra con su requisición.
              </p>
            </div>
            <div className="flex flex-wrap gap-2">
              <button
                onClick={abrirModalFact}
                className="flex items-center gap-2 rounded-xl border border-slate-300 bg-white px-4 py-2.5 text-sm font-semibold text-slate-800 shadow-sm hover:bg-slate-50"
              >
                <Receipt className="h-4 w-4 text-cyan-700" /> Subir factura(s)
              </button>
              <button
                onClick={abrirModalReq}
                className="flex items-center gap-2 rounded-xl bg-slate-900 px-4 py-2.5 text-sm font-semibold text-white shadow-sm hover:bg-slate-800"
              >
                <Upload className="h-4 w-4 text-cyan-300" /> Subir pedido(s) / requisición(es)
              </button>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 xl:grid-cols-4">
          <KPI Icon={ClipboardList} tinte="bg-sky-50 text-sky-700" titulo="Requisiciones abiertas" valor={stats.abiertas}
            sub={stats.total + " registradas en total"} tag="Con pendientes" />
          <KPI Icon={AlertTriangle} tinte="bg-amber-50 text-amber-700" titulo="Pendientes de factura" valor={stats.pend}
            sub="Ítems sin respaldo de compra" tag={stats.pend > 0 ? "Requiere atención" : "Al día"} tagTinte={stats.pend > 0 ? "text-amber-600" : "text-emerald-600"} />
          <KPI Icon={CheckCircle2} tinte="bg-teal-50 text-teal-700" titulo="Conciliación de ítems" valor={stats.pct + "%"}
            sub={stats.rec + " de " + (stats.rec + stats.pend) + " comprobados"} barra={stats.pct} />
          <KPI Icon={Receipt} tinte="bg-emerald-50 text-emerald-700" titulo="Compras registradas" valor={fmtMoney(stats.compras)}
            sub={facts.length + " factura" + (facts.length !== 1 ? "s" : "") + " procesada" + (facts.length !== 1 ? "s" : "")} tag="Acumulado" />
        </div>

        <div className="grid grid-cols-1 gap-4 xl:grid-cols-3">
          <div className="rounded-xl border border-slate-200 bg-white shadow-sm xl:col-span-2">
            <div className="flex items-center justify-between px-4 pt-4">
              <div>
                <h3 className="text-sm font-bold text-slate-900">Actividad reciente</h3>
                <p className="text-xs text-slate-500">Últimos pedidos y requisiciones</p>
              </div>
              <button onClick={() => setVista("pedidos")} className="flex items-center gap-1 text-xs font-semibold text-cyan-700 hover:text-cyan-900">
                Ver todos <ArrowUpRight className="h-3.5 w-3.5" />
              </button>
            </div>
            {recientes.length === 0 ? (
              <p className="px-4 py-8 text-center text-sm text-slate-500">
                Aún no hay registros en {barcoInfo.nombre}. Sube el primer pedido o requisición.
              </p>
            ) : (
              <>
                <div className="mt-3 hidden overflow-x-auto md:block">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b border-slate-100 text-left text-xs uppercase tracking-widest text-slate-400">
                        <th className="px-4 py-2 font-semibold">Requisición</th>
                        <th className="px-2 py-2 font-semibold">Departamento</th>
                        <th className="px-2 py-2 font-semibold">Semana</th>
                        <th className="px-2 py-2 font-semibold">Ítems</th>
                        <th className="px-2 py-2 font-semibold">Estado</th>
                        <th className="px-2 py-2 font-semibold">Registro</th>
                        <th className="px-2 py-2" />
                      </tr>
                    </thead>
                    <tbody>
                      {recientes.map((r) => {
                        const d = depInfo(r.departamento);
                        return (
                          <tr key={r.id} onClick={() => irARequisicion(r)} className="cursor-pointer border-b border-slate-50 last:border-b-0 hover:bg-slate-50">
                            <td className="px-4 py-2.5">
                              <p className="font-semibold text-slate-900" style={F_MONO}>{r.numero}</p>
                              <p className="truncate text-xs text-slate-400">{r.archivoNombre}</p>
                            </td>
                            <td className="px-2 py-2.5">
                              <span className="inline-flex items-center gap-1.5 text-xs text-slate-700">
                                <span className={"h-2 w-2 rounded-full " + d.dot} /> {d.corto}
                              </span>
                            </td>
                            <td className="px-2 py-2.5 text-xs text-slate-600" style={F_MONO}>
                              Sem {numeroSemana(r.semana) || "—"} · {rangoCorto(r.semana, r.semanaFin)}
                            </td>
                            <td className="px-2 py-2.5 text-xs text-slate-700" style={F_MONO}>{r.items.length}</td>
                            <td className="px-2 py-2.5"><Pill e={estadoDe(r)} /></td>
                            <td className="px-2 py-2.5 text-xs text-slate-500" style={F_MONO}>{fmtFechaHora(r.fechaSubida)}</td>
                            <td className="px-2 py-2.5 text-slate-300"><ChevronRight className="h-4 w-4" /></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
                <div className="mt-2 divide-y divide-slate-100 md:hidden">
                  {recientes.map((r) => {
                    const d = depInfo(r.departamento);
                    return (
                      <button key={r.id} onClick={() => irARequisicion(r)} className="flex w-full items-center gap-3 px-4 py-3 text-left hover:bg-slate-50">
                        <span className={"h-2 w-2 shrink-0 rounded-full " + d.dot} />
                        <div className="min-w-0 flex-1">
                          <p className="text-sm font-semibold text-slate-900" style={F_MONO}>{r.numero}</p>
                          <p className="truncate text-xs text-slate-500">{d.corto} · {rangoCorto(r.semana, r.semanaFin)} · {r.items.length} ítems</p>
                        </div>
                        <Pill e={estadoDe(r)} />
                      </button>
                    );
                  })}
                </div>
              </>
            )}
          </div>

          <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
            <h3 className="text-sm font-bold text-slate-900">Carga por departamento</h3>
            <p className="text-xs text-slate-500">Ítems solicitados en {barcoInfo.nombre}</p>
            {cargaDeps.arr.length === 0 ? (
              <p className="py-8 text-center text-sm text-slate-500">Sin datos todavía.</p>
            ) : (
              <div className="mt-4 space-y-3.5">
                {cargaDeps.arr.map(({ dep, items }) => (
                  <div key={dep.id}>
                    <div className="flex items-center justify-between">
                      <span className="inline-flex items-center gap-1.5 text-xs font-semibold text-slate-700">
                        <span className={"h-2 w-2 rounded-full " + dep.dot} /> {dep.corto}
                      </span>
                      <span className="text-xs text-slate-500" style={F_MONO}>{items} ítems</span>
                    </div>
                    <div className="mt-1 h-1.5 rounded-full bg-slate-100">
                      <div className={"h-1.5 rounded-full " + dep.barra} style={{ width: Math.max(6, (items / cargaDeps.max) * 100) + "%" }} />
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        {extrasTotal > 0 && (
          <div className="flex flex-wrap items-center gap-3 rounded-xl border border-amber-200 bg-amber-50 px-4 py-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-amber-100 text-amber-700">
              <AlertTriangle className="h-5 w-5" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-sm font-semibold text-slate-900">{extrasTotal} ítem{extrasTotal !== 1 ? "s" : ""} de facturas sin requisición vinculada</p>
              <p className="text-xs text-slate-600">Compras que no aparecen en ningún pedido registrado.</p>
            </div>
            <button onClick={() => setVista("facturas")} className="flex items-center gap-1 text-xs font-semibold text-amber-800 hover:text-amber-900">
              Revisar en Facturas <ChevronRight className="h-3.5 w-3.5" />
            </button>
          </div>
        )}
      </div>
    );
  };

  const renderPedidos = () => (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center gap-2">
        <button onClick={abrirModalReq} className="flex items-center gap-2 rounded-xl bg-slate-900 px-4 py-2.5 text-sm font-semibold text-white shadow-sm hover:bg-slate-800">
          <Upload className="h-4 w-4 text-cyan-300" /> Subir pedido(s) / requisición(es)
        </button>
        <div className="flex flex-wrap items-center gap-2">
          {semanasDisponibles.length > 0 && (
            <div className="flex items-center gap-1.5">
              <CalendarDays className="h-4 w-4 text-slate-500" />
              <select
                value={filtroSemana}
                onChange={(e) => setFiltroSemana(e.target.value)}
                className="rounded-lg border border-slate-200 bg-white px-2 py-2 text-sm shadow-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
              >
                <option value="todas">Todas las semanas</option>
                {semanasDisponibles.map((s) => <option key={s.clave} value={s.clave}>{nombreSemana(s.inicio, s.fin)}</option>)}
              </select>
            </div>
          )}
          <select
            value={filtroDep}
            onChange={(e) => setFiltroDep(e.target.value)}
            className="rounded-lg border border-slate-200 bg-white px-2 py-2 text-sm shadow-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
          >
            <option value="todos">Todos los departamentos</option>
            {DEPARTAMENTOS.map((d) => <option key={d.id} value={d.id}>{d.corto}</option>)}
          </select>
        </div>
      </div>

      {grupos.length === 0 ? (
        <div className="rounded-xl border border-dashed border-slate-300 bg-white px-6 py-12 text-center">
          <FileText className="mx-auto h-10 w-10 text-slate-300" />
          <p className="mt-3 font-semibold text-slate-800">Sin requisiciones para mostrar</p>
          <p className="mt-1 text-sm text-slate-500">
            Sube un pedido o requisición de {barcoInfo.nombre}: Claude leerá el documento completo, armará el listado y lo verificará dos veces.
          </p>
        </div>
      ) : (
        <div className="space-y-6">
          {grupos.map(({ dep, lista }) => {
            const pend = lista.reduce((s, r) => s + r.items.filter((i) => !i.recibido && esExigible(i)).length, 0);
            return (
              <section key={dep.id}>
                <div className="mb-2 flex items-center gap-2">
                  <div className="flex h-7 w-7 items-center justify-center rounded bg-slate-900">
                    <dep.Icon className="h-4 w-4 text-cyan-300" />
                  </div>
                  <h2 className="text-sm font-bold uppercase tracking-widest text-slate-900">{dep.nombre}</h2>
                  <div className="h-px flex-1 bg-slate-200" />
                  <span className="text-xs text-slate-500" style={F_MONO}>{lista.length} req · {pend} pend.</span>
                </div>
                <div className="space-y-2">
                  {lista.map((r) => (
                    <TarjetaReq
                      key={r.id} r={r} abierto={!!expandidos[r.id]}
                      onToggle={() => setExpandidos((p) => ({ ...p, [r.id]: !p[r.id] }))}
                      confirmarBorrar={confirmarBorrar} onEliminar={eliminarReq}
                      onConciliar={abrirConciliar} onEliminarItem={eliminarItemReq} onNuevoItem={abrirNuevoItem}
                    />
                  ))}
                </div>
              </section>
            );
          })}
        </div>
      )}
    </div>
  );

  const renderFacturas = () => (
    <div className="space-y-4">
      <button onClick={abrirModalFact} className="flex items-center gap-2 rounded-xl bg-slate-900 px-4 py-2.5 text-sm font-semibold text-white shadow-sm hover:bg-slate-800">
        <Receipt className="h-4 w-4 text-cyan-300" /> Subir factura(s)
      </button>
      {facts.length === 0 ? (
        <div className="rounded-xl border border-dashed border-slate-300 bg-white px-6 py-12 text-center">
          <Receipt className="mx-auto h-10 w-10 text-slate-300" />
          <p className="mt-3 font-semibold text-slate-800">Sin facturas procesadas</p>
          <p className="mt-1 text-sm text-slate-500">
            Sube una o varias facturas: Claude marcará ✓ los ítems pedidos, añadirá precio unitario y total, y verificará todo dos veces.
          </p>
        </div>
      ) : (
        <div className="space-y-2">
          {facts.map((f) => (
            <TarjetaFact
              key={f.id} f={f} abierto={!!expandidosF[f.id]}
              onToggle={() => setExpandidosF((p) => ({ ...p, [f.id]: !p[f.id] }))}
              confirmarBorrar={confirmarBorrar} onEliminar={eliminarFact}
              onEliminarLinea={eliminarLineaFactura} onAgregarItem={abrirItemFactura}
              onVincularExtra={abrirVincularExtra}
            />
          ))}
        </div>
      )}
    </div>
  );

  const renderDepartamentos = () => (
    <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 xl:grid-cols-3">
      {DEPARTAMENTOS.map((d) => {
        const lista = reqs.filter((r) => r.departamento === d.id);
        let items = 0, pend = 0, total = 0;
        lista.forEach((r) => r.items.forEach((it) => {
          if (!esExigible(it)) return;
          items++;
          if (!it.recibido) pend++;
          else total += it.precioTotal || 0;
        }));
        return (
          <div key={d.id} className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
            <div className="flex items-center gap-2.5">
              <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-slate-900">
                <d.Icon className="h-5 w-5 text-cyan-300" />
              </div>
              <div className="min-w-0 flex-1">
                <p className="truncate text-sm font-bold text-slate-900">{d.corto}</p>
                <p className="text-xs text-slate-500" style={F_MONO}>{lista.length} requisiciones</p>
              </div>
              <span className={"h-2.5 w-2.5 rounded-full " + d.dot} />
            </div>
            <div className="mt-3 grid grid-cols-3 gap-2 text-center">
              <div className="rounded-lg bg-slate-50 px-1 py-2">
                <p className="text-sm font-bold text-slate-900" style={F_MONO}>{items}</p>
                <p className="text-xs text-slate-500">Ítems</p>
              </div>
              <div className="rounded-lg bg-amber-50 px-1 py-2">
                <p className="text-sm font-bold text-amber-700" style={F_MONO}>{pend}</p>
                <p className="text-xs text-slate-500">Pend.</p>
              </div>
              <div className="rounded-lg bg-emerald-50 px-1 py-2">
                <p className="text-sm font-bold text-emerald-700" style={F_MONO}>{fmtMoney(round2(total))}</p>
                <p className="text-xs text-slate-500">Facturado</p>
              </div>
            </div>
            <button
              onClick={() => { setFiltroDep(d.id); setFiltroSemana("todas"); setVista("pedidos"); }}
              className="mt-3 flex w-full items-center justify-center gap-1 rounded-lg border border-slate-200 px-3 py-2 text-xs font-semibold text-slate-700 hover:bg-slate-50"
            >
              Ver pedidos <ChevronRight className="h-3.5 w-3.5" />
            </button>
          </div>
        );
      })}
    </div>
  );

  const renderCatalogo = () => {
    const q = normaliza(catBusqueda.trim());
    const provs = proveedoresDB.filter((p) => !q || normaliza(p.proveedor).includes(q));
    const its = itemsDB.filter((x) => !q || normaliza(x.descripcion).includes(q));
    return (
      <div className="space-y-4">
        <div className="flex flex-wrap items-center gap-2">
          <div className="flex overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm">
            <button
              onClick={() => { setCatModo("prov"); setCatAbierto(null); }}
              className={"px-4 py-2 text-sm font-semibold " + (catModo === "prov" ? "bg-slate-900 text-white" : "text-slate-600 hover:bg-slate-50")}
            >
              Proveedores
            </button>
            <button
              onClick={() => { setCatModo("item"); setCatAbierto(null); }}
              className={"px-4 py-2 text-sm font-semibold " + (catModo === "item" ? "bg-slate-900 text-white" : "text-slate-600 hover:bg-slate-50")}
            >
              Ítems
            </button>
          </div>
          <div className="relative min-w-0 flex-1">
            <Search className="pointer-events-none absolute left-2.5 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              value={catBusqueda}
              onChange={(e) => setCatBusqueda(e.target.value)}
              placeholder={catModo === "prov" ? "Buscar proveedor…" : "Buscar ítem…"}
              className="w-full rounded-xl border border-slate-200 bg-white py-2 pl-8 pr-3 text-sm shadow-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
            />
          </div>
        </div>

        {facts.length === 0 ? (
          <div className="rounded-xl border border-dashed border-slate-300 bg-white px-6 py-12 text-center">
            <Tags className="mx-auto h-10 w-10 text-slate-300" />
            <p className="mt-3 font-semibold text-slate-800">La base de datos se arma sola</p>
            <p className="mt-1 text-sm text-slate-500">
              Cada factura que subas alimenta el historial de proveedores e ítems de {barcoInfo.nombre}.
            </p>
          </div>
        ) : catModo === "prov" ? (
          provs.length === 0 ? (
            <p className="py-8 text-center text-sm text-slate-500">Ningún proveedor coincide con la búsqueda.</p>
          ) : (
            <div className="space-y-2">
              {provs.map((p) => {
                const abierto = catAbierto === "p:" + p.proveedor;
                return (
                  <div key={p.proveedor} className="overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm">
                    <button
                      onClick={() => setCatAbierto(abierto ? null : "p:" + p.proveedor)}
                      className="flex w-full items-center gap-3 p-3 text-left focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                    >
                      {abierto ? <ChevronDown className="h-4 w-4 shrink-0 text-slate-400" /> : <ChevronRight className="h-4 w-4 shrink-0 text-slate-400" />}
                      <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg bg-cyan-50 text-cyan-700">
                        <Tags className="h-4 w-4" />
                      </div>
                      <div className="min-w-0 flex-1">
                        <p className="truncate text-sm font-bold text-slate-900">{p.proveedor}</p>
                        <p className="text-xs text-slate-500" style={F_MONO}>
                          {p.n} factura{p.n !== 1 ? "s" : ""} · Última: {p.ultima ? fmtFecha(p.ultima) : "—"}
                        </p>
                      </div>
                      <p className="shrink-0 text-sm font-bold text-slate-900" style={F_MONO}>{fmtMoney(round2(p.total))}</p>
                    </button>
                    {abierto && (
                      <div className="border-t border-slate-100">
                        {p.facturas.map((f) => (
                          <div key={f.id} className="flex items-center gap-2 border-b border-slate-50 px-3 py-2 last:border-b-0">
                            <Receipt className="h-4 w-4 shrink-0 text-slate-400" />
                            <div className="min-w-0 flex-1">
                              <p className="truncate text-xs font-semibold text-slate-800" style={F_MONO}>Fact. {f.numero}</p>
                              <p className="text-xs text-slate-500" style={F_MONO}>
                                {f.fecha} · registrada {fmtFecha(f.fechaSubida)} · {(f.coincidencias || []).length + (f.extras || []).length} líneas
                              </p>
                            </div>
                            <p className="shrink-0 text-xs font-semibold text-slate-900" style={F_MONO}>
                              {fmtMoney(round2((f.totalCoincidencias || 0) + (f.totalExtras || 0)))}
                            </p>
                          </div>
                        ))}
                        <button
                          onClick={() => { setVista("facturas"); setExpandidosF(Object.fromEntries(p.facturas.map((f) => [f.id, true]))); }}
                          className="flex w-full items-center justify-center gap-1 px-3 py-2 text-xs font-semibold text-cyan-700 hover:bg-cyan-50"
                        >
                          Abrir en Facturas <ChevronRight className="h-3.5 w-3.5" />
                        </button>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )
        ) : its.length === 0 ? (
          <p className="py-8 text-center text-sm text-slate-500">Ningún ítem coincide con la búsqueda.</p>
        ) : (
          <div className="space-y-2">
            {its.map((x) => {
              const abierto = catAbierto === "i:" + x.descripcion;
              return (
                <div key={x.descripcion} className="overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm">
                  <button
                    onClick={() => setCatAbierto(abierto ? null : "i:" + x.descripcion)}
                    className="flex w-full items-center gap-3 p-3 text-left focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                  >
                    {abierto ? <ChevronDown className="h-4 w-4 shrink-0 text-slate-400" /> : <ChevronRight className="h-4 w-4 shrink-0 text-slate-400" />}
                    <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg bg-teal-50 text-teal-700">
                      <Package className="h-4 w-4" />
                    </div>
                    <div className="min-w-0 flex-1">
                      <p className="truncate text-sm font-bold text-slate-900">{x.descripcion}</p>
                      <p className="text-xs text-slate-500" style={F_MONO}>
                        {round2(x.cant)} {x.unidad} en {x.compras} compra{x.compras !== 1 ? "s" : ""} · Última: {x.ultima ? fmtFecha(x.ultima) : "—"}
                      </p>
                    </div>
                    <p className="shrink-0 text-sm font-bold text-slate-900" style={F_MONO}>{fmtMoney(round2(x.gasto))}</p>
                  </button>
                  {abierto && (
                    <div className="border-t border-slate-100">
                      {x.hist.map((mv, i) => (
                        <div key={i} className="flex items-center gap-2 border-b border-slate-50 px-3 py-2 last:border-b-0">
                          <CalendarDays className="h-4 w-4 shrink-0 text-slate-400" />
                          <div className="min-w-0 flex-1">
                            <p className="truncate text-xs font-semibold text-slate-800">{mv.proveedor}</p>
                            <p className="truncate text-xs text-slate-500" style={F_MONO}>
                              Fact. {mv.factNumero} · {mv.fecha !== "—" ? mv.fecha : fmtFecha(mv.fechaSubida)}
                              {mv.reqNumero ? " · Req. " + mv.reqNumero : " · extra"}
                            </p>
                          </div>
                          <div className="shrink-0 text-right">
                            <p className="text-xs font-semibold text-slate-900" style={F_MONO}>{fmtMoney(mv.precioTotal)}</p>
                            <p className="text-xs text-slate-400" style={F_MONO}>
                              {mv.cantidad} {mv.unidad}{mv.precioUnitario != null ? " × " + fmtMoney(mv.precioUnitario) : ""}
                            </p>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  };

  /* ------------ Navegación ------------ */

  const NAV = [
    { id: "resumen", label: "Resumen operativo", corto: "Resumen", Icon: LayoutDashboard },
    { id: "pedidos", label: "Pedidos y requisiciones", corto: "Pedidos", Icon: ClipboardList, badge: stats.abiertas },
    { id: "facturas", label: "Facturas", corto: "Facturas", Icon: Receipt, badge: facts.length },
    { id: "departamentos", label: "Departamentos", corto: "Deptos", Icon: Boxes },
    { id: "catalogo", label: "Proveedores e ítems", corto: "Catálogo", Icon: Tags },
  ];

  const TITULOS = {
    resumen: "Resumen operativo", pedidos: "Pedidos y requisiciones", facturas: "Facturas",
    departamentos: "Departamentos", catalogo: "Proveedores e ítems",
  };
  const MIGAS = {
    resumen: "Abastecimiento", pedidos: "Pedidos", facturas: "Facturas",
    departamentos: "Departamentos", catalogo: "Base de datos",
  };

  const hoy = new Date();
  const chipFecha = hoy.getDate() + " " + MES_CORTO[hoy.getMonth()].toUpperCase() + " " + hoy.getFullYear();

  /* ============================ Render ============================ */

  return (
    <div className="min-h-screen bg-slate-100 text-slate-900" style={F_BASE}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
        @media (prefers-reduced-motion: reduce) { * { animation-duration: 0.01ms !important; transition-duration: 0.01ms !important; } }
      `}</style>

      <div className="flex">
        {/* ---------- Barra lateral (escritorio) ---------- */}
        <aside className="sticky top-0 hidden h-screen w-64 shrink-0 flex-col overflow-y-auto bg-slate-900 text-slate-300 lg:flex">
          <div className="flex items-center gap-2.5 px-4 py-4">
            <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-cyan-600/20 ring-1 ring-cyan-500/40">
              <Anchor className="h-5 w-5 text-cyan-300" />
            </div>
            <div className="min-w-0">
              <p className="truncate text-sm font-extrabold uppercase tracking-widest text-white">{barcoInfo.nombre}</p>
              <p className="text-xs uppercase tracking-widest text-slate-400">Control</p>
            </div>
          </div>

          <div className="mx-3 rounded-xl bg-slate-800/60 p-3 ring-1 ring-slate-700/60">
            <p className="text-xs font-semibold uppercase tracking-widest text-slate-400">Embarcación activa</p>
            <div className="mt-1.5"><SelectorBarco valor={barco} onChange={cambiarBarco} oscuro /></div>
            <p className="mt-1.5 flex items-center gap-1.5 text-xs text-slate-400">
              <span className="h-1.5 w-1.5 rounded-full bg-emerald-400" /> Operativo · Galápagos
            </p>
          </div>

          <p className="mt-5 px-4 text-xs font-semibold uppercase tracking-widest text-slate-500">Operaciones</p>
          <nav className="mt-1 space-y-0.5 px-2">
            {NAV.filter((n) => n.id !== "catalogo").map((n) => (
              <button
                key={n.id}
                onClick={() => setVista(n.id)}
                className={"flex w-full items-center gap-2.5 rounded-lg px-2.5 py-2 text-sm font-medium " +
                  (vista === n.id ? "bg-sky-900/70 text-white" : "text-slate-300 hover:bg-slate-800/70 hover:text-white")}
              >
                <n.Icon className="h-4 w-4" />
                <span className="flex-1 text-left">{n.label}</span>
                {n.badge > 0 && <span className="rounded-full bg-slate-700 px-1.5 py-0.5 text-xs" style={F_MONO}>{n.badge}</span>}
              </button>
            ))}
          </nav>

          <p className="mt-5 px-4 text-xs font-semibold uppercase tracking-widest text-slate-500">Configuración</p>
          <nav className="mt-1 space-y-0.5 px-2">
            <button
              onClick={() => setVista("catalogo")}
              className={"flex w-full items-center gap-2.5 rounded-lg px-2.5 py-2 text-sm font-medium " +
                (vista === "catalogo" ? "bg-sky-900/70 text-white" : "text-slate-300 hover:bg-slate-800/70 hover:text-white")}
            >
              <Tags className="h-4 w-4" />
              <span className="flex-1 text-left">Proveedores e ítems</span>
            </button>
            <button
              onClick={() => setModalInstalar(true)}
              className="flex w-full items-center gap-2.5 rounded-lg px-2.5 py-2 text-sm font-medium text-slate-300 hover:bg-slate-800/70 hover:text-white"
            >
              <Download className="h-4 w-4" />
              <span className="flex-1 text-left">Instalar aplicación</span>
            </button>
          </nav>

          <div className="mt-auto flex items-center gap-2.5 border-t border-slate-800 px-4 py-3.5">
            <div className="flex h-8 w-8 items-center justify-center rounded-full bg-cyan-700 text-xs font-bold text-white">AS</div>
            <div className="min-w-0">
              <p className="truncate text-sm font-semibold text-white">Antoine Saman</p>
              <p className="text-xs text-slate-400">Administración</p>
            </div>
          </div>
        </aside>

        {/* ---------- Contenido ---------- */}
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-2 bg-slate-900 px-4 py-3 lg:hidden">
            <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-cyan-600/20 ring-1 ring-cyan-500/40">
              <Anchor className="h-4 w-4 text-cyan-300" />
            </div>
            <p className="min-w-0 flex-1 truncate text-sm font-extrabold uppercase tracking-widest text-white">{barcoInfo.nombre}</p>
            <button
              onClick={() => setModalInstalar(true)}
              className="rounded-lg border border-slate-700 bg-slate-800 p-2 text-slate-300 hover:text-white"
              aria-label="Instalar aplicación"
            >
              <Download className="h-4 w-4" />
            </button>
            <SelectorBarco valor={barco} onChange={cambiarBarco} oscuro ancho="w-40" />
          </div>

          <main className="mx-auto max-w-6xl px-4 pb-24 pt-5 sm:px-6 lg:px-8 lg:pb-10">
            <div className="mb-4 flex flex-wrap items-center justify-between gap-2">
              <div>
                <p className="text-xs font-semibold uppercase tracking-widest text-slate-400">
                  Operaciones / {MIGAS[vista]}
                </p>
                <h1 className="text-xl font-bold text-slate-900 sm:text-2xl">{TITULOS[vista]}</h1>
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => cargar(barco)}
                  className="flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-600 shadow-sm hover:bg-slate-50"
                  title="Recargar datos"
                >
                  <RefreshCw className={"h-3.5 w-3.5 " + (cargando ? "animate-spin" : "")} />
                  <span className="hidden sm:inline">Datos actualizados</span>
                </button>
                <span className="flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-semibold text-slate-700 shadow-sm" style={F_MONO}>
                  <CalendarDays className="h-3.5 w-3.5 text-slate-500" /> {chipFecha}
                </span>
              </div>
            </div>

            {cargando ? (
              <div className="flex items-center justify-center gap-2 py-20 text-slate-500">
                <Loader2 className="h-5 w-5 animate-spin" /> Cargando {barcoInfo.nombre}…
              </div>
            ) : vista === "resumen" ? renderResumen()
              : vista === "pedidos" ? renderPedidos()
              : vista === "facturas" ? renderFacturas()
              : vista === "departamentos" ? renderDepartamentos()
              : renderCatalogo()}

            <p className="pt-8 text-center text-xs text-slate-400">
              {barcoInfo.nombre} · Los registros se guardan automáticamente por embarcación · PDF, PNG, JPG, Excel, Word
            </p>
          </main>
        </div>
      </div>

      {/* ---------- Navegación inferior (móvil) ---------- */}
      <nav className="fixed bottom-0 left-0 right-0 z-40 grid grid-cols-5 border-t border-slate-200 bg-white py-1 lg:hidden">
        {NAV.map((n) => (
          <button key={n.id} onClick={() => setVista(n.id)} className="flex flex-col items-center gap-0.5 py-1.5">
            <n.Icon className={"h-5 w-5 " + (vista === n.id ? "text-cyan-700" : "text-slate-400")} />
            <span className={"text-xs " + (vista === n.id ? "font-semibold text-cyan-700" : "text-slate-400")}>{n.corto}</span>
          </button>
        ))}
      </nav>

      {/* ---------- Modal: pedido(s) / requisición(es) ---------- */}
      {modalReq && (
        <ModalBase onClose={() => setModalReq(null)} bloqueado={modalReq.paso === "procesando"}>
          <CabeceraModal titulo="Subir Pedido(s) / Requisición(es)" onClose={() => setModalReq(null)} bloqueado={modalReq.paso === "procesando"} />
          <div className="space-y-4 p-4">
            {modalReq.paso === "form" && (
              <>
                <BannerError msg={modalReq.error} />
                <div>
                  <Etiqueta>Archivos de pedido / requisición</Etiqueta>
                  <input
                    ref={inputReqRef}
                    type="file"
                    multiple
                    accept=".pdf,.png,.jpg,.jpeg,.webp,.gif,.xlsx,.xls,.csv,.docx,.txt"
                    className="hidden"
                    onChange={(e) => {
                      const fs = Array.from(e.target.files || []);
                      if (fs.length) {
                        setModalReq((prev) => ({
                          ...prev,
                          archivos: [...prev.archivos, ...fs.map((f) => ({ id: uid(), archivo: f, numero: "", departamento: "ABASTECIMIENTO" }))],
                          error: null,
                        }));
                      }
                      e.target.value = "";
                    }}
                  />
                  <button
                    onClick={() => inputReqRef.current && inputReqRef.current.click()}
                    className="flex w-full items-center gap-2 rounded-lg border-2 border-dashed border-slate-300 bg-slate-50 px-3 py-3 text-left text-sm text-slate-700 hover:border-cyan-600 hover:bg-cyan-50 focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                  >
                    <Upload className="h-4 w-4 shrink-0 text-cyan-700" />
                    <span>Elegir uno o varios documentos (PDF, imagen, Excel o Word)</span>
                  </button>
                  {modalReq.archivos.length > 0 && (
                    <div className="mt-2 space-y-2">
                      {modalReq.archivos.map((af) => (
                        <div key={af.id} className="space-y-2 rounded-lg border border-slate-200 bg-slate-50 p-2.5">
                          <div className="flex items-center gap-2">
                            <FileText className="h-4 w-4 shrink-0 text-cyan-700" />
                            <span className="flex-1 truncate text-xs text-slate-700">{af.archivo.name}</span>
                            <button
                              onClick={() => quitarArchivoReq(af.id)}
                              className="rounded p-0.5 text-slate-400 hover:text-rose-600"
                              aria-label="Quitar archivo"
                            >
                              <X className="h-4 w-4" />
                            </button>
                          </div>
                          <div className="grid grid-cols-2 gap-2">
                            <input
                              type="text"
                              value={af.numero}
                              onChange={(e) => actualizarArchivoReq(af.id, { numero: e.target.value })}
                              placeholder="N° de requisición"
                              className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-xs focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                              style={F_MONO}
                            />
                            <select
                              value={af.departamento}
                              onChange={(e) => actualizarArchivoReq(af.id, { departamento: e.target.value })}
                              className="w-full rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                            >
                              {DEPARTAMENTOS.map((d) => <option key={d.id} value={d.id}>{d.corto}</option>)}
                            </select>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                  <p className="mt-1.5 text-xs text-slate-500">
                    Cada archivo se procesa por separado y se guarda como su propia requisición, con el número y departamento que le pongas.
                  </p>
                </div>
                <div>
                  <Etiqueta>Semana de crucero (aplica a todo el grupo)</Etiqueta>
                  {modalReq.modoFecha === "semana" ? (
                    <select
                      value={modalReq.semana}
                      onChange={(e) => setModalReq((prev) => ({ ...prev, semana: e.target.value }))}
                      className="w-full rounded-lg border border-slate-300 bg-white px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                    >
                      {generarSemanas().map((s) => <option key={s} value={s}>{nombreSemana(s)}</option>)}
                    </select>
                  ) : (
                    <div className="grid grid-cols-2 gap-2">
                      <div>
                        <p className="mb-0.5 text-xs text-slate-500">Inicio del crucero</p>
                        <input
                          type="date"
                          value={modalReq.fechaInicio}
                          onChange={(e) => setModalReq((prev) => ({ ...prev, fechaInicio: e.target.value, error: null }))}
                          className="w-full rounded-lg border border-slate-300 bg-white px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                          style={F_MONO}
                        />
                      </div>
                      <div>
                        <p className="mb-0.5 text-xs text-slate-500">Fin del crucero</p>
                        <input
                          type="date"
                          value={modalReq.fechaFin}
                          onChange={(e) => setModalReq((prev) => ({ ...prev, fechaFin: e.target.value, error: null }))}
                          className="w-full rounded-lg border border-slate-300 bg-white px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                          style={F_MONO}
                        />
                      </div>
                    </div>
                  )}
                  {modalReq.modoFecha === "personalizado" && esISO(modalReq.fechaInicio) && esISO(modalReq.fechaFin) && (
                    modalReq.fechaFin > modalReq.fechaInicio ? (
                      <p className="mt-1 text-xs font-medium text-cyan-800" style={F_MONO}>
                        {nombreSemana(modalReq.fechaInicio, modalReq.fechaFin)} · {Math.round((parseISO(modalReq.fechaFin) - parseISO(modalReq.fechaInicio)) / 86400000) + 1} días
                      </p>
                    ) : (
                      <p className="mt-1 text-xs text-rose-600">La fecha de fin debe ser posterior a la de inicio.</p>
                    )
                  )}
                  <button
                    onClick={() =>
                      setModalReq((prev) =>
                        prev.modoFecha === "semana"
                          ? { ...prev, modoFecha: "personalizado", fechaInicio: prev.semana, fechaFin: sumarDias(prev.semana, 7), error: null }
                          : { ...prev, modoFecha: "semana", error: null }
                      )
                    }
                    className="mt-1.5 rounded text-xs font-semibold text-cyan-700 underline underline-offset-2 hover:text-cyan-900 focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                  >
                    {modalReq.modoFecha === "semana"
                      ? "Cambiar fechas manualmente (crucero de 10 días u otro rango)"
                      : "Volver a semanas estándar (lunes a lunes)"}
                  </button>
                </div>
                <p className="text-xs text-slate-500">
                  Claude lee cada documento por completo, fila por fila, y hace una doble verificación para asegurar que no falte ningún ítem de la lista.
                </p>
                <button
                  onClick={analizarRequisiciones}
                  className="w-full rounded-xl bg-slate-900 px-4 py-3 font-semibold text-white hover:bg-slate-800 focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                >
                  Analizar {modalReq.archivos.length > 1 ? modalReq.archivos.length + " documentos" : "documento"}
                </button>
              </>
            )}

            {modalReq.paso === "procesando" && (
              <div className="flex flex-col items-center gap-3 py-10 text-center">
                <Loader2 className="h-8 w-8 animate-spin text-cyan-700" />
                <p className="text-sm font-medium text-slate-800">Claude está trabajando…</p>
                <p className="text-xs text-slate-500">{modalReq.progreso || "Extrayendo ítems, cantidades y unidades."}</p>
              </div>
            )}

            {modalReq.paso === "preview" && (
              <>
                <BannerError msg={modalReq.error} />
                <p className="text-sm font-semibold text-slate-800">
                  {modalReq.resultados.length} requisición{modalReq.resultados.length !== 1 ? "es" : ""} procesada{modalReq.resultados.length !== 1 ? "s" : ""}
                </p>
                <div className="space-y-4">
                  {modalReq.resultados.map((r) => (
                    <div key={r.fileId} className="space-y-2.5 rounded-xl border border-slate-200 p-3">
                      <div className="flex items-center justify-between gap-2 rounded-lg bg-slate-900 px-3 py-2">
                        <div className="min-w-0">
                          <p className="truncate text-sm font-bold text-white" style={F_MONO}>N° {r.numero}</p>
                          <p className="truncate text-xs text-slate-300">{depInfo(r.departamento).corto} · {r.archivoNombre}</p>
                        </div>
                        <div className="flex shrink-0 items-center gap-1.5">
                          <span className="rounded bg-cyan-500/20 px-1.5 py-0.5 text-xs font-semibold text-cyan-200" style={F_MONO}>
                            {r.items.length} ítems
                          </span>
                          <button
                            onClick={() => setModalReq((prev) => ({ ...prev, resultados: prev.resultados.filter((x) => x.fileId !== r.fileId) }))}
                            className="rounded p-1 text-slate-400 hover:bg-rose-500/20 hover:text-rose-300"
                            aria-label="Quitar esta requisición del lote"
                            title="Quitar del lote"
                          >
                            <X className="h-3.5 w-3.5" />
                          </button>
                        </div>
                      </div>
                      <BannerError msg={r.error} />
                      {r.verificacion && (
                        <div className={"flex items-start gap-2 rounded-lg px-3 py-2 text-xs " +
                          (r.verificacion.totalDoc != null && r.items.length < r.verificacion.totalDoc
                            ? "border border-amber-200 bg-amber-50 text-amber-800"
                            : "border border-emerald-200 bg-emerald-50 text-emerald-800")}>
                          <ShieldCheck className="mt-0.5 h-4 w-4 shrink-0" />
                          <span>
                            Doble verificación: {r.verificacion.correcciones} correcciones · {r.verificacion.faltantes} ítems recuperados
                            {r.verificacion.totalDoc != null && (
                              <> · {r.items.length}/{r.verificacion.totalDoc} según el documento
                                {r.items.length < r.verificacion.totalDoc && ". Revisa y agrega manualmente los que falten."}
                              </>
                            )}
                          </span>
                        </div>
                      )}
                      <div className="max-h-56 divide-y divide-slate-100 overflow-y-auto rounded-lg border border-slate-200">
                        {r.items.length === 0 && (
                          <p className="px-2.5 py-3 text-center text-xs text-slate-400">Sin ítems. Agrégalos manualmente abajo.</p>
                        )}
                        {r.items.map((it) => (
                          <div key={it.id} className="flex items-center gap-2 px-2.5 py-1.5">
                            <div className="min-w-0 flex-1">
                              <p className="truncate text-xs font-medium text-slate-800">{it.descripcion}</p>
                              <p className="text-xs text-slate-500" style={F_MONO}>
                                {(num(it.cantidad) ?? 0) === 0
                                  ? <span className="rounded bg-slate-100 px-1 py-0.5 font-semibold text-slate-500">NO PEDIDO · {it.unidad}</span>
                                  : <>{it.cantidad} {it.unidad}</>}
                              </p>
                            </div>
                            <button
                              onClick={() => quitarItemResultadoReq(r.fileId, it.id)}
                              className="rounded p-1 text-slate-300 hover:bg-rose-50 hover:text-rose-600"
                              aria-label="Quitar ítem"
                            >
                              <X className="h-4 w-4" />
                            </button>
                          </div>
                        ))}
                      </div>
                      <div className="space-y-2 rounded-lg border border-dashed border-slate-300 bg-slate-50 p-2.5">
                        <p className="text-xs font-semibold uppercase tracking-widest text-slate-500">Agregar ítem manual</p>
                        <input
                          type="text"
                          value={r.nDesc}
                          onChange={(e) => actualizarResultadoReq(r.fileId, { nDesc: e.target.value })}
                          placeholder="Descripción del ítem"
                          className="w-full rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                        />
                        <div className="flex gap-2">
                          <input
                            type="number"
                            value={r.nCant}
                            onChange={(e) => actualizarResultadoReq(r.fileId, { nCant: e.target.value })}
                            placeholder="Cantidad"
                            className="w-full min-w-0 flex-1 rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                            style={F_MONO}
                          />
                          <input
                            type="text"
                            value={r.nUnidad}
                            onChange={(e) => actualizarResultadoReq(r.fileId, { nUnidad: e.target.value })}
                            placeholder="Unidad"
                            className="w-full min-w-0 flex-1 rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                            style={F_MONO}
                          />
                          <button
                            onClick={() => agregarItemResultadoReq(r.fileId)}
                            className="flex shrink-0 items-center gap-1 rounded-lg bg-cyan-700 px-3 py-2 text-sm font-semibold text-white hover:bg-cyan-800"
                          >
                            <Plus className="h-4 w-4" /> Agregar
                          </button>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <button
                    onClick={() => setModalReq((prev) => ({ ...prev, paso: "form" }))}
                    className="rounded-xl border border-slate-300 px-4 py-2.5 text-sm font-semibold text-slate-700 hover:bg-slate-50"
                  >
                    Volver
                  </button>
                  <button
                    onClick={guardarRequisiciones}
                    disabled={modalReq.resultados.every((r) => r.items.length === 0)}
                    className="rounded-xl bg-emerald-600 px-4 py-2.5 text-sm font-semibold text-white hover:bg-emerald-700 disabled:opacity-40"
                  >
                    Guardar {modalReq.resultados.filter((r) => r.items.length > 0).length > 1
                      ? modalReq.resultados.filter((r) => r.items.length > 0).length + " requisiciones"
                      : "requisición"}
                  </button>
                </div>
                {modalReq.resultados.some((r) => r.items.length === 0) && (
                  <p className="text-xs text-amber-700">
                    N° {modalReq.resultados.filter((r) => r.items.length === 0).map((r) => r.numero).join(", ")} no se {modalReq.resultados.filter((r) => r.items.length === 0).length > 1 ? "guardarán" : "guardará"} por no tener ítems. Agrega al menos uno manualmente o vuelve atrás para quitarla del lote.
                  </p>
                )}
              </>
            )}
          </div>
        </ModalBase>
      )}

      {/* ---------- Modal: facturas ---------- */}
      {modalFact && (
        <ModalBase onClose={() => setModalFact(null)} bloqueado={modalFact.paso === "procesando"}>
          <CabeceraModal titulo="Subir Factura(s)" onClose={() => setModalFact(null)} bloqueado={modalFact.paso === "procesando"} />
          <div className="space-y-4 p-4">
            {modalFact.paso === "form" && (
              <>
                <BannerError msg={modalFact.error} />
                <div>
                  <Etiqueta>Archivos de factura</Etiqueta>
                  <input
                    ref={inputFactRef}
                    type="file"
                    multiple
                    accept=".pdf,.png,.jpg,.jpeg,.webp,.gif,.xlsx,.xls,.csv,.docx,.txt"
                    className="hidden"
                    onChange={(e) => {
                      const fs = Array.from(e.target.files || []);
                      if (fs.length) setModalFact((prev) => ({ ...prev, archivos: [...prev.archivos, ...fs], error: null }));
                      e.target.value = "";
                    }}
                  />
                  <button
                    onClick={() => inputFactRef.current && inputFactRef.current.click()}
                    className="flex w-full items-center gap-2 rounded-lg border-2 border-dashed border-slate-300 bg-slate-50 px-3 py-3 text-left text-sm text-slate-700 hover:border-cyan-600 hover:bg-cyan-50 focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                  >
                    <Upload className="h-4 w-4 shrink-0 text-cyan-700" />
                    <span>Elegir una o varias facturas (PDF, imagen, Excel o Word)</span>
                  </button>
                  {modalFact.archivos.length > 0 && (
                    <div className="mt-2 space-y-1">
                      {modalFact.archivos.map((f, i) => (
                        <div key={i} className="flex items-center gap-2 rounded-lg border border-slate-200 bg-slate-50 px-2.5 py-1.5">
                          <FileText className="h-4 w-4 shrink-0 text-cyan-700" />
                          <span className="flex-1 truncate text-xs text-slate-700">{f.name}</span>
                          <button
                            onClick={() => setModalFact((prev) => ({ ...prev, archivos: prev.archivos.filter((_, j) => j !== i) }))}
                            className="rounded p-0.5 text-slate-400 hover:text-rose-600"
                            aria-label="Quitar archivo"
                          >
                            <X className="h-4 w-4" />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
                <div>
                  <Etiqueta>Comparar contra</Etiqueta>
                  <select
                    value={modalFact.objetivo}
                    onChange={(e) => setModalFact((prev) => ({ ...prev, objetivo: e.target.value }))}
                    className="w-full rounded-lg border border-slate-300 bg-white px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                  >
                    <option value="todas">Todas las requisiciones con ítems pendientes</option>
                    {reqs.filter((r) => r.items.some((i) => !i.recibido)).map((r) => (
                      <option key={r.id} value={r.id}>
                        Req. {r.numero} · {depInfo(r.departamento).corto} · {nombreSemana(r.semana, r.semanaFin)}
                      </option>
                    ))}
                  </select>
                </div>
                <p className="text-xs text-slate-500">
                  Claude ubicará cada ítem de la factura, marcará ✓ los que coincidan con lo pedido (con precio unitario y total) y hará una doble verificación de cantidades y precios contra el documento.
                </p>
                <button
                  onClick={procesarFacturas}
                  className="w-full rounded-xl bg-slate-900 px-4 py-3 font-semibold text-white hover:bg-slate-800 focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                >
                  Procesar {modalFact.archivos.length > 1 ? modalFact.archivos.length + " facturas" : "factura"}
                </button>
              </>
            )}

            {modalFact.paso === "procesando" && (
              <div className="flex flex-col items-center gap-3 py-10 text-center">
                <Loader2 className="h-8 w-8 animate-spin text-cyan-700" />
                <p className="text-sm font-medium text-slate-800">Claude está cruzando facturas con requisiciones…</p>
                {modalFact.progreso && <p className="text-xs text-slate-500">{modalFact.progreso}</p>}
              </div>
            )}

            {modalFact.paso === "resultados" && (
              <>
                <p className="text-sm font-semibold text-slate-800">
                  {modalFact.resultados.length} factura{modalFact.resultados.length !== 1 ? "s" : ""} procesada{modalFact.resultados.length !== 1 ? "s" : ""}
                </p>
                <div className="space-y-5">
                  {modalFact.resultados.map((f) => (
                    <div key={f.id} className="space-y-2">
                      <div className="rounded-lg bg-slate-900 px-3 py-2 text-white">
                        <p className="text-sm font-bold" style={F_MONO}>Fact. {f.numero}</p>
                        <p className="text-xs text-slate-300">{f.proveedor} · {f.fecha}</p>
                      </div>
                      <DetalleFactura f={f} conAcciones confirmarBorrar={confirmarBorrar} onEliminarLinea={eliminarLineaFactura} onAgregarItem={abrirItemFactura} onVincularExtra={abrirVincularExtra} />
                    </div>
                  ))}
                </div>
                <button
                  onClick={() => setModalFact(null)}
                  className="w-full rounded-xl bg-slate-900 px-4 py-3 font-semibold text-white hover:bg-slate-800"
                >
                  Listo
                </button>
              </>
            )}
          </div>
        </ModalBase>
      )}

      {/* ---------- Modal: conciliar / editar ítem ---------- */}
      {modalCon && (
        <ModalBase onClose={() => setModalCon(null)}>
          <CabeceraModal titulo="Conciliar ítem" onClose={() => setModalCon(null)} />
          <div className="space-y-4 p-4">
            <div className="rounded-lg bg-slate-50 px-3 py-2">
              <p className="text-sm font-semibold text-slate-800">{modalCon.descripcion}</p>
              <p className="text-xs text-slate-500" style={F_MONO}>Pedido: {modalCon.pedido} {modalCon.unidad}</p>
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <Etiqueta>Cantidad facturada</Etiqueta>
                <input
                  type="number"
                  value={modalCon.cant}
                  onChange={(e) => setModalCon((prev) => ({ ...prev, cant: e.target.value }))}
                  className="w-full rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                  style={F_MONO}
                />
              </div>
              <div>
                <Etiqueta>Precio unitario ($)</Etiqueta>
                <input
                  type="number"
                  value={modalCon.precio}
                  onChange={(e) => setModalCon((prev) => ({ ...prev, precio: e.target.value }))}
                  placeholder="Opcional"
                  className="w-full rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                  style={F_MONO}
                />
              </div>
            </div>
            <div>
              <Etiqueta>N° de factura (opcional)</Etiqueta>
              <input
                type="text"
                value={modalCon.fact}
                onChange={(e) => setModalCon((prev) => ({ ...prev, fact: e.target.value }))}
                placeholder="Ej. 001-002-000123456"
                className="w-full rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                style={F_MONO}
              />
            </div>
            {num(modalCon.precio) != null && num(modalCon.cant) != null && (
              <p className="text-xs text-slate-600" style={F_MONO}>
                Total: {fmtMoney(round2((num(modalCon.cant) || 0) * (num(modalCon.precio) || 0)))}
              </p>
            )}
            <div className="grid grid-cols-2 gap-2">
              {modalCon.recibido ? (
                <button onClick={marcarPendiente} className="rounded-xl border border-amber-300 bg-amber-50 px-4 py-2.5 text-sm font-semibold text-amber-800 hover:bg-amber-100">
                  Marcar pendiente
                </button>
              ) : (
                <button onClick={() => setModalCon(null)} className="rounded-xl border border-slate-300 px-4 py-2.5 text-sm font-semibold text-slate-700 hover:bg-slate-50">
                  Cancelar
                </button>
              )}
              <button onClick={guardarConciliacion} className="rounded-xl bg-emerald-600 px-4 py-2.5 text-sm font-semibold text-white hover:bg-emerald-700">
                Guardar ✓
              </button>
            </div>
          </div>
        </ModalBase>
      )}

      {/* ---------- Modal: agregar ítem a requisición ---------- */}
      {modalNuevo && (
        <ModalBase onClose={() => setModalNuevo(null)}>
          <CabeceraModal titulo="Agregar ítem a la requisición" onClose={() => setModalNuevo(null)} />
          <div className="space-y-4 p-4">
            <BannerError msg={modalNuevo.error} />
            <div>
              <Etiqueta>Descripción</Etiqueta>
              <input
                type="text"
                value={modalNuevo.desc}
                onChange={(e) => setModalNuevo((prev) => ({ ...prev, desc: e.target.value, error: null }))}
                placeholder="Ej. ACEITE DE OLIVA"
                className="w-full rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
              />
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <Etiqueta>Cantidad</Etiqueta>
                <input
                  type="number"
                  value={modalNuevo.cant}
                  onChange={(e) => setModalNuevo((prev) => ({ ...prev, cant: e.target.value }))}
                  placeholder="1"
                  className="w-full rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                  style={F_MONO}
                />
              </div>
              <div>
                <Etiqueta>Unidad</Etiqueta>
                <input
                  type="text"
                  value={modalNuevo.unidad}
                  onChange={(e) => setModalNuevo((prev) => ({ ...prev, unidad: e.target.value }))}
                  placeholder="UND, KG, LT…"
                  className="w-full rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                  style={F_MONO}
                />
              </div>
            </div>
            <button onClick={guardarNuevoItem} className="w-full rounded-xl bg-slate-900 px-4 py-3 font-semibold text-white hover:bg-slate-800">
              Agregar ítem
            </button>
          </div>
        </ModalBase>
      )}

      {/* ---------- Modal: agregar ítem manual a factura ---------- */}
      {modalItemFact && (
        <ModalBase onClose={() => setModalItemFact(null)}>
          <CabeceraModal titulo="Agregar ítem a la factura" onClose={() => setModalItemFact(null)} />
          <div className="space-y-4 p-4">
            <BannerError msg={modalItemFact.error} />
            <div>
              <Etiqueta>Vincular con ítem pendiente</Etiqueta>
              <select
                value={modalItemFact.vinculo}
                onChange={(e) => {
                  const v = e.target.value;
                  let cant = modalItemFact.cant;
                  if (v !== "extra") {
                    const [rid, iid] = v.split("::");
                    const req = reqs.find((r) => r.id === rid);
                    const it = req && req.items.find((i) => i.id === iid);
                    if (it) cant = String(it.cantidad);
                  }
                  setModalItemFact((prev) => ({ ...prev, vinculo: v, cant, error: null }));
                }}
                className="w-full rounded-lg border border-slate-300 bg-white px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
              >
                <option value="extra">Ítem extra (no estaba en requisiciones)</option>
                {pendientesGlobal.map(({ req, it }) => (
                  <option key={it.id} value={req.id + "::" + it.id}>
                    {it.descripcion} · {it.cantidad} {it.unidad} · Req. {req.numero}
                  </option>
                ))}
              </select>
              <p className="mt-1 text-xs text-slate-500">
                Si lo vinculas, el ítem se marcará ✓ en su requisición con el precio que ingreses.
              </p>
            </div>
            {modalItemFact.vinculo === "extra" && (
              <>
                <div>
                  <Etiqueta>Descripción</Etiqueta>
                  <input
                    type="text"
                    value={modalItemFact.desc}
                    onChange={(e) => setModalItemFact((prev) => ({ ...prev, desc: e.target.value, error: null }))}
                    placeholder="Ej. HIELO EN FUNDA"
                    className="w-full rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                  />
                </div>
                <div>
                  <Etiqueta>Unidad</Etiqueta>
                  <input
                    type="text"
                    value={modalItemFact.unidad}
                    onChange={(e) => setModalItemFact((prev) => ({ ...prev, unidad: e.target.value }))}
                    placeholder="UND, KG, LT…"
                    className="w-full rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                    style={F_MONO}
                  />
                </div>
              </>
            )}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <Etiqueta>Cantidad</Etiqueta>
                <input
                  type="number"
                  value={modalItemFact.cant}
                  onChange={(e) => setModalItemFact((prev) => ({ ...prev, cant: e.target.value }))}
                  placeholder="1"
                  className="w-full rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                  style={F_MONO}
                />
              </div>
              <div>
                <Etiqueta>Precio unitario ($)</Etiqueta>
                <input
                  type="number"
                  value={modalItemFact.precio}
                  onChange={(e) => setModalItemFact((prev) => ({ ...prev, precio: e.target.value }))}
                  placeholder="Opcional"
                  className="w-full rounded-lg border border-slate-300 px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
                  style={F_MONO}
                />
              </div>
            </div>
            <button onClick={guardarItemFactura} className="w-full rounded-xl bg-slate-900 px-4 py-3 font-semibold text-white hover:bg-slate-800">
              Agregar a la factura
            </button>
          </div>
        </ModalBase>
      )}

      {/* ---------- Modal: agregar extra de factura a una requisición ---------- */}
      {modalVincular && (
        <ModalBase onClose={() => setModalVincular(null)}>
          <CabeceraModal titulo="Agregar a requisición" onClose={() => setModalVincular(null)} />
          <div className="space-y-4 p-4">
            <BannerError msg={modalVincular.error} />
            <div className="rounded-lg border border-amber-200 bg-amber-50 px-3 py-2">
              <p className="text-xs font-semibold uppercase tracking-widest text-amber-700">Ítem de la factura</p>
              <p className="mt-0.5 text-sm font-medium text-slate-900">{modalVincular.extra.descripcion}</p>
              <p className="text-xs text-slate-600" style={F_MONO}>
                {modalVincular.extra.cantidad} {modalVincular.extra.unidad}
                {modalVincular.extra.precioTotal != null && <> · {fmtMoney(modalVincular.extra.precioTotal)}</>}
              </p>
            </div>

            {modalVincular.cargandoSugerencia ? (
              <div className="flex items-center gap-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2.5 text-sm text-slate-600">
                <Loader2 className="h-4 w-4 animate-spin text-cyan-700" />
                Analizando a qué ítem pendiente corresponde…
              </div>
            ) : modalVincular.sugerencia ? (() => {
              const req = reqs.find((r) => r.id === modalVincular.sugerencia.reqId);
              const it = req && req.items.find((i) => i.id === modalVincular.sugerencia.itemId);
              if (!req || !it) return null;
              const valor = req.id + "::" + it.id;
              const activa = modalVincular.seleccion === valor;
              return (
                <button
                  onClick={() => setModalVincular((prev) => ({ ...prev, seleccion: valor, error: null }))}
                  className={"w-full rounded-lg border px-3 py-2.5 text-left " +
                    (activa ? "border-cyan-500 bg-cyan-50 ring-1 ring-cyan-500" : "border-slate-200 bg-white hover:border-cyan-300")}
                >
                  <p className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-widest text-cyan-700">
                    <ShieldCheck className="h-3.5 w-3.5" /> Sugerencia de Claude {activa && "· seleccionada"}
                  </p>
                  <p className="mt-1 text-sm font-medium text-slate-900">{it.descripcion}</p>
                  <p className="text-xs text-slate-500" style={F_MONO}>
                    {(num(it.cantidad) ?? 0) === 0 ? "NO PEDIDO" : it.cantidad + " " + it.unidad} · Req. {req.numero} ({depInfo(req.departamento).corto})
                  </p>
                  {modalVincular.sugerencia.motivo && (
                    <p className="mt-1 text-xs text-slate-500">{modalVincular.sugerencia.motivo}</p>
                  )}
                </button>
              );
            })() : (
              <p className="text-xs text-slate-500">Sin una sugerencia clara para este ítem. Elígelo manualmente en la lista.</p>
            )}

            <div>
              <Etiqueta>Ítem pendiente de requisición</Etiqueta>
              <select
                value={modalVincular.seleccion}
                onChange={(e) => setModalVincular((prev) => ({ ...prev, seleccion: e.target.value, error: null }))}
                className="w-full rounded-lg border border-slate-300 bg-white px-2 py-2 text-sm focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-600"
              >
                <option value="">Elegir ítem pendiente…</option>
                {pendientesGlobal.map(({ req, it }) => (
                  <option key={it.id} value={req.id + "::" + it.id}>
                    {it.descripcion} · {(num(it.cantidad) ?? 0) === 0 ? "NO PEDIDO" : it.cantidad + " " + it.unidad} · Req. {req.numero}
                  </option>
                ))}
              </select>
              <p className="mt-1 text-xs text-slate-500">
                El ítem quedará ✓ en su requisición con la cantidad y el precio de esta factura, y la línea pasará de extras a coincidencias.
              </p>
            </div>

            <button
              onClick={confirmarVinculoExtra}
              disabled={!modalVincular.seleccion}
              className="w-full rounded-xl bg-slate-900 px-4 py-3 font-semibold text-white hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-300"
            >
              Agregar a la requisición
            </button>
          </div>
        </ModalBase>
      )}

      {/* ---------- Modal: instalar aplicación ---------- */}
      {modalInstalar && (
        <ModalBase onClose={() => setModalInstalar(false)}>
          <CabeceraModal titulo="Instalar aplicación" onClose={() => setModalInstalar(false)} />
          <div className="space-y-4 p-4">
            <div className="flex items-start gap-3 rounded-lg border border-cyan-200 bg-cyan-50 px-3 py-2.5">
              <Smartphone className="mt-0.5 h-5 w-5 shrink-0 text-cyan-700" />
              <p className="text-sm text-slate-700">
                Publica este artefacto en Claude (botón <span className="font-semibold">Compartir → Publicar</span>) y abre el enlace en tu celular o computadora. Luego agrégalo como app:
              </p>
            </div>
            <div className="space-y-3">
              <div className="rounded-lg border border-slate-200 p-3">
                <p className="text-sm font-bold text-slate-900">iPhone / iPad (Safari)</p>
                <p className="mt-1 text-sm text-slate-600">
                  Abre el enlace → botón <span className="font-semibold">Compartir</span> (cuadro con flecha) → <span className="font-semibold">Añadir a pantalla de inicio</span>. Queda con su propio ícono, como una app.
                </p>
              </div>
              <div className="rounded-lg border border-slate-200 p-3">
                <p className="text-sm font-bold text-slate-900">Android (Chrome)</p>
                <p className="mt-1 text-sm text-slate-600">
                  Abre el enlace → menú <span className="font-semibold">⋮</span> → <span className="font-semibold">Añadir a pantalla principal</span> (o "Instalar app").
                </p>
              </div>
              <div className="rounded-lg border border-slate-200 p-3">
                <p className="text-sm font-bold text-slate-900">Computadora (Chrome / Edge)</p>
                <p className="mt-1 text-sm text-slate-600">
                  Abre el enlace y usa el ícono de <span className="font-semibold">instalar</span> en la barra de direcciones, o guarda el enlace en marcadores.
                </p>
              </div>
            </div>
            <p className="text-xs text-slate-500">
              Los datos quedan guardados con tu cuenta de Claude, separados por embarcación, así que verás lo mismo en el celular y en la computadora al iniciar sesión.
            </p>
            <button onClick={() => setModalInstalar(false)} className="w-full rounded-xl bg-slate-900 px-4 py-3 font-semibold text-white hover:bg-slate-800">
              Entendido
            </button>
          </div>
        </ModalBase>
      )}
    </div>
  );
}
