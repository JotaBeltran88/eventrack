import React, { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import * as XLSXStyle from "xlsx-js-style"; // solo para la plantilla con bordes
import { loadState, saveState, saveWithMerge, subscribeState, login, setCodes } from "./storage";

// ── Fusión a 3 bandas (base, local, servidor) para guardado concurrente ──
// Permite que dos personas editen barras distintas a la vez sin pisarse.
function deepEqual(a, b) { return JSON.stringify(a) === JSON.stringify(b); }
function isObj(x) { return x && typeof x === "object" && !Array.isArray(x); }
function tieneId(x) { return isObj(x) && "id" in x; }
// Serialización con claves ordenadas: estable frente al reordenado de jsonb de Supabase.
function sjson(x) {
  if (x === null || typeof x !== "object") return JSON.stringify(x);
  if (Array.isArray(x)) return "[" + x.map(sjson).join(",") + "]";
  return "{" + Object.keys(x).sort().map((k) => JSON.stringify(k) + ":" + sjson(x[k])).join(",") + "}";
}
function merge3(base, mine, theirs) {
  if (mine === undefined) {
    if (base !== undefined && deepEqual(base, theirs)) return undefined; // yo borré, ellos no tocaron
    return theirs;
  }
  if (theirs === undefined) {
    if (base !== undefined && deepEqual(base, mine)) return undefined; // ellos borraron, yo no toqué
    return mine;
  }
  if (isObj(mine) && isObj(theirs)) {
    const out = {};
    const keys = new Set([...Object.keys(mine), ...Object.keys(theirs), ...(isObj(base) ? Object.keys(base) : [])]);
    for (const k of keys) { const v = merge3(isObj(base) ? base[k] : undefined, mine[k], theirs[k]); if (v !== undefined) out[k] = v; }
    return out;
  }
  if (Array.isArray(mine) && Array.isArray(theirs)) {
    const baseArr = Array.isArray(base) ? base : [];
    if (mine.every(tieneId) && theirs.every(tieneId)) {
      const byId = (a) => { const m = {}; a.forEach((x) => (m[x.id] = x)); return m; };
      const mMine = byId(mine), mTheirs = byId(theirs), mBase = byId(baseArr);
      const res = []; const seen = new Set();
      theirs.forEach((t) => { seen.add(t.id); const m = merge3(mBase[t.id], mMine[t.id], mTheirs[t.id]); if (m !== undefined) res.push(m); });
      mine.forEach((mn) => { if (!seen.has(mn.id)) { const m = merge3(mBase[mn.id], mn, undefined); if (m !== undefined) res.push(m); } });
      return res;
    }
    if (!deepEqual(mine, baseArr)) return mine;
    return theirs;
  }
  if (deepEqual(mine, base)) return theirs; // yo no cambié → lo de ellos
  if (deepEqual(theirs, base)) return mine;  // ellos no cambiaron → lo mío
  return mine; // conflicto en el mismo campo → gana quien guarda
}
function mergeEstado(base, mine, theirs) {
  const r = merge3(base || { eventos: [] }, mine || { eventos: [] }, theirs || { eventos: [] });
  return r && Array.isArray(r.eventos) ? r : { eventos: (mine && mine.eventos) || [] };
}

// Compacta: elimina las celdas de conteo a cero (se asumen 0 por defecto al leer).
function podarEvento(ev) {
  if (!ev || !Array.isArray(ev.jornadas)) return ev;
  return {
    ...ev,
    jornadas: ev.jornadas.map((j) => {
      if (!j || !j.conteo) return j;
      const conteo = {};
      for (const u of Object.keys(j.conteo)) {
        const cu = {};
        for (const pid of Object.keys(j.conteo[u] || {})) {
          const c = j.conteo[u][pid];
          if (c && ((c.inicial || 0) !== 0 || (c.final || 0) !== 0)) cu[pid] = c;
        }
        if (Object.keys(cu).length) conteo[u] = cu;
      }
      return { ...j, conteo };
    }),
  };
}
function podarEstado(estado) {
  if (!estado || !Array.isArray(estado.eventos)) return estado;
  return { ...estado, eventos: estado.eventos.map(podarEvento) };
}

// ─────────────────────────────────────────────────────────────
//  EVENTRACK · Jota Beltrán
//  App multi-evento de inventario y control de stock, con JORNADAS.
//  Configuración común al evento:
//    · Ubicaciones · Referencias de producto
//  Por jornada:
//    · Conteo de inventario (inicial/final → consumo)
//  Más acumulado del evento completo (suma de jornadas).
//  Datos compartidos persistentes en la nube (Supabase, ver src/storage.js).
// ─────────────────────────────────────────────────────────────

// Tema "SaaS Claro": fondo claro, acento índigo. Se mantienen los nombres
// semánticos (gold = acento, cream = texto principal) para no tocar los estilos.
const COLORS = {
  bg: "#f7f8fa", panel: "#ffffff", panel2: "#f1f2f6", line: "#e6e8ee",
  gold: "#4f46e5", goldDim: "#8b90a8", cream: "#1f2430", dim: "#6b7280",
  green: "#1a9d5f", red: "#dc2626", amber: "#d97706", amberBg: "#fff7ed", amberLine: "#fed7aa",
};

// Estado de completitud de una jornada: cuántas ubicaciones están confirmadas.
function jornadaEstado(evento, j) {
  const total = evento.ubicaciones.length;
  const conf = evento.ubicaciones.filter((u) => j.confirmado && j.confirmado[u]).length;
  return { total, conf, completo: total > 0 && conf === total, empezado: conf > 0 };
}

// Jornada por defecto al abrir el evento: la del "día de trabajo" actual.
// El día activo NO cambia hasta las 10:00 del día siguiente (antes de esa hora,
// cuenta como el día anterior). Ej.: domingo 9:35 → se considera el sábado.
function jornadaPorDefecto(jornadas) {
  if (!jornadas || jornadas.length === 0) return null;
  const now = new Date();
  const ref = new Date(now);
  if (now.getHours() < 10) ref.setDate(ref.getDate() - 1);
  const hoy = `${ref.getFullYear()}-${String(ref.getMonth() + 1).padStart(2, "0")}-${String(ref.getDate()).padStart(2, "0")}`;
  const exacto = jornadas.find((j) => j.fecha === hoy);
  if (exacto) return exacto.id;
  // Si no hay jornada para hoy, la más reciente cuya fecha ya haya pasado.
  const pasadas = jornadas.filter((j) => j.fecha && j.fecha <= hoy).sort((a, b) => b.fecha.localeCompare(a.fecha));
  if (pasadas.length) return pasadas[0].id;
  return jornadas[0].id; // todas futuras o sin fecha → la primera
}

// ¿Se ha EMPEZADO a contar esta ubicación? No cuenta el Inicial precargado
// (stock de partida): solo un Final introducido o un movimiento registrado.
function ubicacionEmpezada(j, u) {
  const cont = j.conteo && j.conteo[u];
  if (cont) {
    for (const pid in cont) {
      if (((cont[pid] || {}).final || 0) !== 0) return true;
    }
  }
  for (const m of (j.movimientos || [])) {
    if (m.ubic === u || m.destino === u) return true;
  }
  return false;
}

// Ubicaciones SIN confirmar donde alguien ha empezado: hay un nombre en
// "Realizado por", o un Final/movimiento. El Inicial precargado NO cuenta.
function jornadaPendienteConfirmar(evento, j) {
  return evento.ubicaciones.filter((u) => {
    if (j.confirmado && j.confirmado[u]) return false;
    const tieneNombre = !!(j.realizadoPor && j.realizadoPor[u] && String(j.realizadoPor[u]).trim());
    return tieneNombre || ubicacionEmpezada(j, u);
  });
}

// Valores sospechosos de una jornada: celdas donde el Final supera lo disponible
// (Final > inicial + entradas − salidas) ⇒ imposible, casi siempre un error de tecleo.
function jornadaRevisar(evento, j) {
  const issues = [];
  for (const u of evento.ubicaciones) {
    for (const p of evento.productos) {
      const t = celdaTotales(j, u, p.id);
      if (t.fin > t.ini) issues.push({ ubic: u, pid: p.id, prod: p.nombre, fin: t.fin, disp: t.ini });
    }
  }
  return issues;
}
// ¿Esta celda concreta tiene un valor sospechoso? (Final > disponible)
function celdaSospechosa(jornada, ubic, pid) {
  const t = celdaTotales(jornada, ubic, pid);
  return t.fin > t.ini;
}

// Banners de aviso (días sin terminar + valores a revisar). onIr(jornadaId)
// se llama al tocar un día. Reutilizado en Conteo y Resumen.
function AvisosJornadas({ evento, jornadaActivaId, onIr }) {
  const pendientes = [];
  evento.jornadas.forEach((j) => jornadaPendienteConfirmar(evento, j).forEach((u) => pendientes.push({ j, ubic: u })));
  const conIssues = evento.jornadas.map((j) => ({ j, n: jornadaRevisar(evento, j).length })).filter((x) => x.n > 0);
  if (pendientes.length === 0 && conIssues.length === 0) return null;
  const totalIssues = conIssues.reduce((s, x) => s + x.n, 0);
  return (
    <>
      {pendientes.length > 0 && (
        <div style={styles.alertBox}>
          <div style={{ fontWeight: 700, color: COLORS.amber, marginBottom: 8 }}>
            ⚠ {pendientes.length === 1 ? "1 ubicación con datos sin confirmar" : `${pendientes.length} ubicaciones con datos sin confirmar`}
          </div>
          <div style={{ fontSize: 12.5, color: COLORS.dim, marginBottom: 9 }}>
            Tienen información introducida pero sin confirmar. Toca una para ir a revisarla y confirmarla.
          </div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 7 }}>
            {pendientes.map(({ j, ubic }) => {
              const activa = j.id === jornadaActivaId;
              const quien = (j.realizadoPor && j.realizadoPor[ubic]) || "";
              return (
                <button key={j.id + "·" + ubic} onClick={() => onIr(j.id, ubic)} style={{ ...styles.alertChip, ...(activa ? styles.alertChipActive : {}) }}>
                  {fechaLabel(j.fecha)} · {ubic}{quien ? ` · ${quien}` : ""}{j.editable === false ? " 🔒" : ""}
                </button>
              );
            })}
          </div>
        </div>
      )}
      {conIssues.length > 0 && (
        <div style={styles.errorBox}>
          <div style={{ fontWeight: 700, color: COLORS.red, marginBottom: 8 }}>
            ⛔ {totalIssues === 1 ? "1 valor a revisar" : `${totalIssues} valores a revisar`}
          </div>
          <div style={{ fontSize: 12.5, color: COLORS.dim, marginBottom: 9 }}>
            Hay un Final mayor que el stock disponible (imposible) — probablemente un número mal escrito. Toca el día para ir a corregirlo; la celda aparece marcada en rojo.
          </div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 7 }}>
            {conIssues.map(({ j, n }) => {
              const activa = j.id === jornadaActivaId;
              return (
                <button key={j.id} onClick={() => onIr(j.id)} style={{ ...styles.errorChip, ...(activa ? styles.errorChipActive : {}) }}>
                  {fechaLabel(j.fecha)} · {n}{j.editable === false ? " 🔒" : ""}
                </button>
              );
            })}
          </div>
        </div>
      )}
    </>
  );
}

function nuevoEvento(nombre, fecha) {
  return {
    id: "ev" + Date.now().toString(36),
    nombre: nombre || "Evento sin nombre",
    tipo: "single",
    fecha: fecha || "",
    fechaFin: fecha || "",
    ubicaciones: [],
    productos: [],
    jornadas: [],
  };
}

// Una jornada vacía (sin conteo); el conteo se rellena al añadir ubicaciones/productos.
function jornadaVacia(fecha) {
  return { id: "j" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), fecha, conteo: {} };
}

function hoyISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

// Última fecha del evento (la jornada más tardía, o su fecha fin/inicio).
function fechaFinEvento(ev) {
  const fechas = (ev.jornadas || []).map((j) => j.fecha).filter(Boolean).sort();
  if (fechas.length) return fechas[fechas.length - 1];
  return ev.fechaFin || ev.fecha || "";
}

// Un evento es "pasado" si su última fecha ya quedó atrás. Sin fecha => en curso.
function esEventoPasado(ev) {
  const fin = fechaFinEvento(ev);
  if (!fin) return false;
  return fin < hoyISO();
}

function fechaLabel(iso) {
  if (!iso) return "Sin fecha";
  const [y, m, d] = iso.split("-");
  if (!y || !m || !d) return iso;
  const dias = ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"];
  const meses = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  const dt = new Date(Number(y), Number(m) - 1, Number(d));
  return `${dias[dt.getDay()]} ${Number(d)} ${meses[Number(m) - 1]}`;
}

// ── Lectura de la plantilla Excel ──
const normTxt = (s) => s.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

function leerHojaWB(wb, nombres) {
  const target = nombres.map(normTxt);
  const hoja = wb.SheetNames.find((sn) => target.includes(normTxt(sn)));
  if (!hoja) return [];
  return XLSX.utils.sheet_to_json(wb.Sheets[hoja], { defval: "" });
}

function campoFila(fila, alias) {
  const claves = Object.keys(fila);
  for (const a of alias) {
    const k = claves.find((c) => normTxt(c) === normTxt(a));
    if (k != null && fila[k] !== "") return fila[k].toString().trim();
  }
  return "";
}

// Lee un workbook y devuelve { ubicaciones:[str], productos:[{nombre,categoria,unidad}] }
function parsearPlantilla(wb) {
  const ubicaciones = [];
  leerHojaWB(wb, ["Ubicaciones", "Ubicacion"]).forEach((r) => {
    const v = campoFila(r, ["Ubicación", "Ubicacion", "Nombre"]);
    if (v && !ubicaciones.includes(v)) ubicaciones.push(v);
  });
  const productos = [];
  leerHojaWB(wb, ["Productos", "Producto"]).forEach((r) => {
    const nom = campoFila(r, ["Producto", "Nombre", "Referencia"]);
    if (!nom) return;
    productos.push({
      nombre: nom,
      categoria: campoFila(r, ["Categoría", "Categoria"]) || "Otros",
      unidad: campoFila(r, ["Unidad", "Ud"]) || "ud",
    });
  });
  return { ubicaciones, productos };
}

// Genera y descarga la plantilla en blanco desde el navegador
function descargarPlantillaBlanco() {
  const wb = XLSX.utils.book_new();
  const info = XLSX.utils.aoa_to_sheet([
    ["EVENTRACK · Plantilla de importación"],
    ["Rellena las hojas Ubicaciones y Productos."],
    ["No cambies los nombres de las hojas ni los encabezados de la fila 1."],
    ["En Eventrack pulsa «Importar desde Excel» al crear un evento o dentro de Configuración."],
  ]);
  info["!cols"] = [{ wch: 80 }];
  XLSX.utils.book_append_sheet(wb, info, "Instrucciones");
  const u = XLSX.utils.aoa_to_sheet([["Ubicación"]]); u["!cols"] = [{ wch: 30 }];
  XLSX.utils.book_append_sheet(wb, u, "Ubicaciones");
  const p = XLSX.utils.aoa_to_sheet([["Producto", "Categoría", "Unidad"]]); p["!cols"] = [{ wch: 32 }, { wch: 20 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(wb, p, "Productos");
  XLSX.writeFile(wb, "Eventrack_Plantilla_Importacion.xlsx");
}

export default function App() {
  const [loaded, setLoaded] = useState(false);
  const [saving, setSaving] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [eventos, setEventos] = useState([]);
  const [eventoActivoId, setEventoActivoId] = useState(null);
  const [session, setSession] = useState(() => {
    try { return JSON.parse(localStorage.getItem("eventrack-session") || "null"); }
    catch { return null; }
  });
  const [showAcceso, setShowAcceso] = useState(false);
  const [showPapelera, setShowPapelera] = useState(false);

  const role = session?.role || null;
  const handleLogin = (r) => {
    const s = { role: r };
    localStorage.setItem("eventrack-session", JSON.stringify(s));
    setSession(s);
  };
  const cerrarSesion = () => {
    localStorage.removeItem("eventrack-session");
    setSession(null);
    setEventoActivoId(null);
    setShowAcceso(false);
    setShowPapelera(false);
  };

  const lastWritten = useRef("");                 // firma estable de lo último escrito (anti-eco)
  const baseRef = useRef({ eventos: [] });        // estado del servidor que conocemos (ancestro)
  const baseJsonRef = useRef(sjson({ eventos: [] }));
  const eventosRef = useRef([]);                  // valor local actual (para leer en callbacks)
  useEffect(() => { eventosRef.current = eventos; }, [eventos]);

  const fijarBase = (obj, json) => { baseRef.current = obj; baseJsonRef.current = json || sjson(obj); };

  // Carga inicial + suscripción en tiempo real (fusiona lo entrante con lo local sin guardar).
  useEffect(() => {
    (async () => {
      try {
        const data = await loadState();
        if (data && Array.isArray(data.eventos)) {
          const json = sjson(data);
          fijarBase(data, json);
          lastWritten.current = json;
          setEventos(data.eventos);
        }
      } catch (e) { /* primera vez o sin conexión */ }
      setLoaded(true);
    })();

    const unsub = subscribeState((data) => {
      if (!data || !Array.isArray(data.eventos)) return;
      const json = sjson(data);
      if (json === lastWritten.current) return; // nuestro propio eco
      const merged = mergeEstado(baseRef.current, { eventos: eventosRef.current }, data);
      fijarBase(data, json);                     // su estado es el nuevo ancestro común
      setEventos(merged.eventos);                // mantiene mis cambios sin guardar + los suyos
    });
    return unsub;
  }, []);

  const save = useCallback(async () => {
    setSaving(true);
    try {
      const merged = await saveWithMerge((server) =>
        podarEstado(mergeEstado(baseRef.current, { eventos: eventosRef.current }, server && Array.isArray(server.eventos) ? server : { eventos: [] }))
      );
      if (merged) {
        const json = sjson(merged);
        fijarBase(merged, json);
        lastWritten.current = json;
        setEventos(merged.eventos);
        setDirty(false);
      }
    } catch (e) { console.error("Error guardando", e); }
    setSaving(false);
  }, []);

  const guardarAhora = save;

  // Guardado automático (debounced) cuando lo local difiere del servidor conocido.
  useEffect(() => {
    if (!loaded) return;
    const json = sjson({ eventos });
    if (json === baseJsonRef.current) { setDirty(false); return; } // local == servidor
    setDirty(true);
    const t = setTimeout(() => { save(); }, 600);
    return () => clearTimeout(t);
  }, [eventos, loaded, save]);

  // Aviso al cerrar/recargar si quedan cambios sin guardar.
  useEffect(() => {
    const handler = (e) => { if (dirty || saving) { e.preventDefault(); e.returnValue = ""; } };
    window.addEventListener("beforeunload", handler);
    return () => window.removeEventListener("beforeunload", handler);
  }, [dirty, saving]);

  const updateEvento = (id, updater) =>
    setEventos((prev) => prev.map((ev) => (ev.id === id ? updater(ev) : ev)));
  const addEvento = (nombre, fecha) => {
    const ev = nuevoEvento(nombre, fecha);
    setEventos((prev) => [...prev, ev]);
    return ev.id;
  };
  const addEventoCompleto = (ev) => {
    setEventos((prev) => [...prev, ev]);
    return ev.id;
  };
  // Eliminar = mover a la papelera (borrado lógico, recuperable).
  const removeEvento = (id) => {
    updateEvento(id, (ev) => ({ ...ev, borrado: true }));
    if (eventoActivoId === id) setEventoActivoId(null);
  };
  const restaurarEvento = (id) => updateEvento(id, (ev) => ({ ...ev, borrado: false }));
  const eliminarDefinitivo = (id) =>
    setEventos((prev) => prev.filter((ev) => ev.id !== id));

  if (!session) {
    return <Login onLogin={handleLogin} />;
  }

  if (!loaded) {
    return (
      <div style={{ ...styles.app, justifyContent: "center", alignItems: "center", minHeight: "60vh" }}>
        <style>{globalCSS}</style>
        <div style={{ color: COLORS.gold, fontFamily: "'Fraunces', serif", fontSize: 18 }}>Cargando Eventrack…</div>
      </div>
    );
  }

  const eventoActivo = eventos.find((ev) => ev.id === eventoActivoId);

  return (
    <div style={styles.app}>
      <style>{globalCSS}</style>
      <header style={styles.header}>
        <div>
          {eventoActivo && (
            <button onClick={() => { setEventoActivoId(null); setShowAcceso(false); setShowPapelera(false); }} style={styles.volverBtn}>‹ Volver a eventos</button>
          )}
          <div style={{ cursor: "pointer", display: "flex", alignItems: "center", gap: 10 }} onClick={() => { setEventoActivoId(null); setShowAcceso(false); setShowPapelera(false); }}>
            <img src="/icon.png" alt="Eventrack" style={{ height: 44, width: 44, objectFit: "contain" }} />
            <h1 style={styles.title}>Eventrack</h1>
          </div>
        </div>
        <div style={styles.headerRight}>
          {saving
            ? <span style={styles.saveTag}>Guardando…</span>
            : dirty
              ? <button onClick={guardarAhora} style={styles.saveDirtyBtn}>● Guardar cambios</button>
              : <span style={styles.saveTag}>Guardado ✓</span>}
          <span style={styles.roleTag}>{role === "admin" ? "Admin" : role === "inventario" ? "Gestor" : "Contador"}</span>
          {role === "admin" && (
            <button onClick={() => { setEventoActivoId(null); setShowPapelera(false); setShowAcceso(true); }} style={styles.linkBtn}>⚙ Acceso</button>
          )}
          {role === "admin" && (
            <button onClick={() => { setEventoActivoId(null); setShowAcceso(false); setShowPapelera(true); }} style={styles.linkBtn}>🗑 Papelera</button>
          )}
          <button onClick={cerrarSesion} style={styles.linkBtn}>Cerrar sesión</button>
        </div>
      </header>

      {showAcceso ? (
        <AccessView onClose={() => setShowAcceso(false)} />
      ) : showPapelera ? (
        <PapeleraView eventos={eventos} onClose={() => setShowPapelera(false)} onRestore={restaurarEvento} onHardDelete={eliminarDefinitivo} />
      ) : !eventoActivo ? (
        <EventosList eventos={eventos} role={role} onOpen={setEventoActivoId} onAddCompleto={addEventoCompleto} onRemove={removeEvento} updateEvento={updateEvento} />
      ) : (
        <EventoDetalle evento={eventoActivo} role={role} updateEvento={updateEvento} onGuardar={guardarAhora} />
      )}
    </div>
  );
}

function EventosList({ eventos, role, onOpen, onAddCompleto, onRemove, updateEvento }) {
  const esAdmin = role === "admin";
  const [tipo, setTipo] = useState("single");
  const [nombre, setNombre] = useState("");
  const [fecha, setFecha] = useState("");
  const [desde, setDesde] = useState("");
  const [hasta, setHasta] = useState("");
  const [editId, setEditId] = useState(null);
  const [editNombre, setEditNombre] = useState("");
  const [editFecha, setEditFecha] = useState("");
  const [pasadosAbierto, setPasadosAbierto] = useState(false);
  const [creando, setCreando] = useState(false);
  const empezarEdicion = (ev) => { setEditId(ev.id); setEditNombre(ev.nombre); setEditFecha(ev.fecha || ""); };
  const guardarEdicion = () => {
    updateEvento(editId, (ev) => ({ ...ev, nombre: editNombre.trim() || ev.nombre, fecha: editFecha }));
    setEditId(null);
  };

  const jornadasDeRango = (d1, d2) => {
    const out = [];
    if (!d1 || !d2) return out;
    const start = new Date(d1 + "T00:00:00");
    const end = new Date(d2 + "T00:00:00");
    if (end < start) return out;
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) out.push(jornadaVacia(d.toISOString().slice(0, 10)));
    return out;
  };

  const crear = () => {
    if (!nombre.trim()) return;
    const t = Date.now().toString(36);
    const single = tipo === "single";
    const jornadas = single ? (fecha ? [jornadaVacia(fecha)] : []) : jornadasDeRango(desde, hasta);
    const ev = {
      id: "ev" + t, nombre: nombre.trim(), tipo,
      fecha: (single ? fecha : desde) || "",
      fechaFin: (single ? fecha : hasta) || "",
      ubicaciones: [], productos: [], jornadas,
    };
    const id = onAddCompleto(ev);
    setNombre(""); setFecha(""); setDesde(""); setHasta(""); setTipo("single");
    onOpen(id);
  };

  const metaLinea = (ev) => {
    const nj = ev.jornadas.length;
    const fechaTxt = ev.tipo === "multi" && ev.fecha && ev.fechaFin
      ? `${fechaLabel(ev.fecha)} – ${fechaLabel(ev.fechaFin)}`
      : (ev.fecha ? fechaLabel(ev.fecha) : "Sin fecha");
    return `${fechaTxt} · ${nj} ${nj === 1 ? "jornada" : "jornadas"} · ${ev.ubicaciones.length} ubic. · ${ev.productos.length} ref.`;
  };

  const tarjeta = (ev) => {
    if (editId === ev.id) {
      return (
        <div key={ev.id} style={{ ...styles.eventCard, flexDirection: "column", alignItems: "stretch", gap: 8 }}>
          <input value={editNombre} onChange={(e) => setEditNombre(e.target.value)} placeholder="Nombre del evento" style={styles.textInput} />
          <input type="date" value={editFecha} onChange={(e) => setEditFecha(e.target.value)} style={styles.textInput} />
          <div style={styles.formRow}>
            <button onClick={guardarEdicion} style={styles.addBtn}>Guardar</button>
            <button onClick={() => setEditId(null)} style={styles.smallBtn}>Cancelar</button>
            <button onClick={() => { if (confirm(`¿Mover "${ev.nombre}" a la papelera?`)) { onRemove(ev.id); setEditId(null); } }} style={styles.deleteBtn}>Eliminar</button>
          </div>
        </div>
      );
    }
    return (
      <div key={ev.id} style={styles.eventCard} onClick={() => onOpen(ev.id)}>
        <div style={{ flex: 1 }}>
          <div style={styles.eventName}>{ev.nombre}</div>
          <div style={styles.eventMeta}>{metaLinea(ev)}</div>
        </div>
        {esAdmin && <button onClick={(e) => { e.stopPropagation(); empezarEdicion(ev); }} style={styles.editBtn} title="Editar">✎</button>}
        {esAdmin && <button onClick={(e) => { e.stopPropagation(); if (confirm(`¿Borrar "${ev.nombre}"?`)) onRemove(ev.id); }} style={styles.xBtn}>×</button>}
        <span style={styles.chevron}>›</span>
      </div>
    );
  };

  // En curso: por fecha de inicio ascendente (sin fecha al final).
  const porInicioAsc = (a, b) => (a.fecha || "9999-99-99").localeCompare(b.fecha || "9999-99-99");
  // Pasados: por fecha de fin descendente (el más reciente primero).
  const porFinDesc = (a, b) => (fechaFinEvento(b) || "").localeCompare(fechaFinEvento(a) || "");
  const enCurso = eventos.filter((ev) => !ev.borrado && !esEventoPasado(ev)).sort(porInicioAsc);
  const pasados = eventos.filter((ev) => !ev.borrado && esEventoPasado(ev)).sort(porFinDesc);

  // Página aparte para crear un evento nuevo.
  if (esAdmin && creando) {
    return (
      <div>
        <button onClick={() => setCreando(false)} style={styles.volverBtn}>‹ Volver</button>
        <div style={{ ...styles.sectionTitle, marginTop: 14 }}>Nuevo evento</div>
        <div style={styles.formCard}>
          <input value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Nombre del evento" style={styles.textInput} />
          <div style={styles.formRow}>
            <button onClick={() => setTipo("single")} style={{ ...styles.chip, ...(tipo === "single" ? styles.chipActive : {}) }}>Un solo día</button>
            <button onClick={() => setTipo("multi")} style={{ ...styles.chip, ...(tipo === "multi" ? styles.chipActive : {}) }}>Varios días</button>
          </div>
          {tipo === "single" ? (
            <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} style={styles.textInput} />
          ) : (
            <div style={styles.formRow}>
              <input type="date" value={desde} onChange={(e) => setDesde(e.target.value)} placeholder="Desde" style={styles.textInput} />
              <input type="date" value={hasta} onChange={(e) => setHasta(e.target.value)} placeholder="Hasta" style={styles.textInput} />
            </div>
          )}
          <div style={styles.formRow}>
            <button onClick={crear} style={styles.addBtn}>+ Crear evento</button>
          </div>
          <div style={styles.dimText}>
            Elige si el evento es de un solo día o de varios. En "varios días" se crean automáticamente las jornadas del rango. Las ubicaciones y productos se añaden en el siguiente paso, al configurar el evento.
          </div>
        </div>
      </div>
    );
  }

  return (
    <div>
      {esAdmin && (
        <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 22 }}>
          <button onClick={() => setCreando(true)} style={styles.addBtn}>+ Nuevo evento</button>
        </div>
      )}

      <div style={styles.sectionTitle}>Eventos en curso</div>
      {enCurso.length === 0 && <div style={styles.empty}>{esAdmin ? "No hay eventos en curso. Crea uno con «+ Nuevo evento»." : "No hay eventos en curso."}</div>}
      <div style={{ display: "grid", gap: 10, marginBottom: 28 }}>
        {enCurso.map(tarjeta)}
      </div>

      <div onClick={() => setPasadosAbierto((v) => !v)} style={styles.collapseHeader}>
        <span style={{ ...styles.sectionTitle, marginBottom: 0 }}>Eventos pasados</span>
        <span style={styles.collapseMeta}>{pasados.length} {pasadosAbierto ? "▲" : "▼"}</span>
      </div>
      {pasadosAbierto && (
        pasados.length === 0
          ? <div style={styles.empty}>No hay eventos pasados.</div>
          : <div style={{ display: "grid", gap: 10, marginBottom: 28 }}>{pasados.map(tarjeta)}</div>
      )}
    </div>
  );
}

function EventoDetalle({ evento, role, updateEvento, onGuardar }) {
  const esAdmin = role === "admin";
  const [modo, setModo] = useState(esAdmin && !evento.configurado ? "config" : "operar");
  const [tab, setTab] = useState("conteo");
  const [jornadaActivaId, setJornadaActivaId] = useState(() => jornadaPorDefecto(evento.jornadas));
  const [ubicActiva, setUbicActiva] = useState(evento.ubicaciones[0] || "");

  const upd = (updater) => updateEvento(evento.id, updater);

  useEffect(() => {
    if (!evento.ubicaciones.includes(ubicActiva)) setUbicActiva(evento.ubicaciones[0] || "");
  }, [evento.ubicaciones, ubicActiva]);

  useEffect(() => {
    if (jornadaActivaId && !evento.jornadas.find((j) => j.id === jornadaActivaId)) {
      setJornadaActivaId(jornadaPorDefecto(evento.jornadas));
    }
  }, [evento.jornadas, jornadaActivaId]);

  const jornadaActiva = evento.jornadas.find((j) => j.id === jornadaActivaId) || null;

  // ── Fase 1: configuración del evento (solo admin) ──
  if (modo === "config") {
    return (
      <div>
        <div style={styles.eventHeader}>
          <div>
            <div style={styles.eventTitle}>{evento.nombre}</div>
            <div style={styles.eventMeta}>Configuración del evento</div>
          </div>
        </div>
        <ConfigView evento={evento} upd={upd} />
        <JornadasView evento={evento} upd={upd} jornadaActivaId={jornadaActivaId} setJornadaActivaId={setJornadaActivaId} />
        <button
          onClick={() => { upd((ev) => ({ ...ev, configurado: true })); setTab("conteo"); setModo("operar"); }}
          style={styles.continueBtn}
        >
          Guardar configuración y continuar →
        </button>
      </div>
    );
  }

  // ── Fase 2: operar el evento (conteo y resumen) ──
  return (
    <div>
      <div style={styles.eventHeader}>
        <div>
          <div style={styles.eventTitle}>{evento.nombre}</div>
          <div style={styles.eventMeta}>{evento.jornadas.length} jornadas · {evento.ubicaciones.length} ubicaciones · {evento.productos.length} referencias</div>
        </div>
        {esAdmin && <ExportButton evento={evento} />}
      </div>

      <nav style={styles.tabs}>
        {(esAdmin ? [["conteo", "Conteo"], ["resumen", "Resumen"]] : [["conteo", "Conteo"]]).map(([k, label]) => (
          <button key={k} onClick={() => setTab(k)} style={{ ...styles.tab, ...(tab === k ? styles.tabActive : {}) }}>{label}</button>
        ))}
        {esAdmin && (
          <button onClick={() => setModo("config")} style={styles.tab}>⚙ Configuración</button>
        )}
      </nav>

      {tab === "conteo" && (
        <ConteoView evento={evento} role={role} upd={upd} jornada={jornadaActiva} jornadaActivaId={jornadaActivaId} setJornadaActivaId={setJornadaActivaId} ubicActiva={ubicActiva} setUbicActiva={setUbicActiva} onGuardar={onGuardar} />
      )}
      {tab === "resumen" && <ResumenView evento={evento} setTab={setTab} setJornadaActivaId={setJornadaActivaId} setUbicActiva={setUbicActiva} />}
    </div>
  );
}

// Jornada sin conteo: las celdas se crean solo al ponerles valor (almacenamiento disperso).
function emptyJornada(fecha) {
  return { id: "j" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), fecha, conteo: {} };
}

function JornadasView({ evento, upd, jornadaActivaId, setJornadaActivaId }) {
  const esSingle = evento.tipo === "single";
  const [fecha, setFecha] = useState("");
  const [desde, setDesde] = useState("");
  const [hasta, setHasta] = useState("");

  const addJornada = (f) => {
    if (!f) return;
    if (evento.jornadas.some((j) => j.fecha === f)) return;
    const j = emptyJornada(f, evento);
    upd((ev) => ({ ...ev, jornadas: [...ev.jornadas, j].sort((a, b) => (a.fecha || "").localeCompare(b.fecha || "")) }));
    setJornadaActivaId(j.id);
  };

  const addRango = () => {
    if (!desde || !hasta) return;
    const start = new Date(desde + "T00:00:00");
    const end = new Date(hasta + "T00:00:00");
    if (end < start) return;
    const nuevas = [];
    const existentes = new Set(evento.jornadas.map((j) => j.fecha));
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      const iso = d.toISOString().slice(0, 10);
      if (!existentes.has(iso)) { nuevas.push(emptyJornada(iso, evento)); existentes.add(iso); }
    }
    if (nuevas.length) {
      upd((ev) => ({ ...ev, jornadas: [...ev.jornadas, ...nuevas].sort((a, b) => (a.fecha || "").localeCompare(b.fecha || "")) }));
    }
    setDesde(""); setHasta("");
  };

  const removeJornada = (id) => upd((ev) => ({ ...ev, jornadas: ev.jornadas.filter((j) => j.id !== id) }));
  const toggleEditable = (id) => upd((ev) => ({
    ...ev,
    jornadas: ev.jornadas.map((j) => (j.id === id ? { ...j, editable: j.editable === false ? true : false } : j)),
  }));

  return (
    <div>
      <div style={{ ...styles.sectionTitle, marginTop: 34 }}>Jornadas {esSingle ? "(día único)" : "del evento"}</div>
      {!esSingle && <div style={styles.dimText}>Marca qué jornadas puede editar el Contador. Las de «Solo lectura» solo podrá verlas.</div>}
      {evento.jornadas.length === 0 && <div style={styles.empty}>{esSingle ? "Añade la fecha del día abajo." : "Sin jornadas. Añade noches sueltas o genera un rango de fechas abajo."}</div>}

      <div style={{ display: "grid", gap: 8, marginBottom: 24 }}>
        {evento.jornadas.map((j) => (
          <div key={j.id} style={{ ...styles.jornadaRow, ...(j.id === jornadaActivaId ? styles.jornadaRowActive : {}) }} onClick={() => setJornadaActivaId(j.id)}>
            <span style={{ flex: 1, color: COLORS.cream, fontWeight: 500 }}>{fechaLabel(j.fecha)}</span>
            {!esSingle && (
              <button
                onClick={(e) => { e.stopPropagation(); toggleEditable(j.id); }}
                style={{ ...styles.toggleBtn, ...(j.editable === false ? styles.toggleLocked : styles.toggleOpen) }}
              >
                {j.editable === false ? "🔒 Solo lectura" : "🔓 Editable"}
              </button>
            )}
            <button onClick={(e) => { e.stopPropagation(); removeJornada(j.id); }} style={styles.xBtnSm}>×</button>
          </div>
        ))}
      </div>

      {(!esSingle || evento.jornadas.length === 0) && (
        <div style={styles.formCard}>
          <div style={styles.formCardTitle}>Añadir una jornada</div>
          <div style={styles.formRow}>
            <input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} style={styles.textInput} />
            <button onClick={() => { addJornada(fecha); setFecha(""); }} style={styles.addBtn}>+ Añadir</button>
          </div>
        </div>
      )}

      {!esSingle && (
        <div style={styles.formCard}>
          <div style={styles.formCardTitle}>Generar rango de fechas</div>
          <div style={styles.formRow}>
            <input type="date" value={desde} onChange={(e) => setDesde(e.target.value)} style={styles.textInput} />
            <input type="date" value={hasta} onChange={(e) => setHasta(e.target.value)} style={styles.textInput} />
          </div>
          <button onClick={addRango} style={styles.addBtn}>+ Generar todas las jornadas</button>
          <div style={styles.dimText}>Crea una jornada por cada día entre las dos fechas (incluidas).</div>
        </div>
      )}
    </div>
  );
}

function ConfigView({ evento, upd }) {
  const [nuevaUbic, setNuevaUbic] = useState("");
  const [pNombre, setPNombre] = useState(""); const [pCat, setPCat] = useState(""); const [pUnidad, setPUnidad] = useState("");
  const [importMsg, setImportMsg] = useState(null);
  const fileRef = React.useRef(null);

  const onFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportMsg("Leyendo archivo…");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const datos = parsearPlantilla(wb);

      upd((ev) => {
        const t = Date.now().toString(36);
        const ubicNuevas = datos.ubicaciones.filter((u) => !ev.ubicaciones.includes(u));
        const ubicaciones = [...ev.ubicaciones, ...ubicNuevas];

        const existeProd = (nom) => ev.productos.some((p) => normTxt(p.nombre) === normTxt(nom));
        const prodNuevos = datos.productos
          .filter((p) => !existeProd(p.nombre))
          .map((p, i) => ({ id: "p" + t + i, ...p }));
        const productos = [...ev.productos, ...prodNuevos];

        return { ...ev, ubicaciones, productos };
      });

      setImportMsg(`Importado: +${datos.ubicaciones.length} ubic., +${datos.productos.length} prod. (se omiten los repetidos).`);
      e.target.value = "";
    } catch (err) {
      console.error(err);
      setImportMsg("No se pudo leer el archivo. Asegúrate de que es un .xlsx válido.");
      e.target.value = "";
    }
  };

  const addUbic = () => {
    const n = nuevaUbic.trim();
    if (!n || evento.ubicaciones.includes(n)) return;
    upd((ev) => ({ ...ev, ubicaciones: [...ev.ubicaciones, n] }));
    setNuevaUbic("");
  };
  const removeUbic = (n) => upd((ev) => ({
    ...ev,
    ubicaciones: ev.ubicaciones.filter((u) => u !== n),
    jornadas: ev.jornadas.map((j) => { const c = { ...j.conteo }; delete c[n]; return { ...j, conteo: c }; }),
  }));

  const addProd = () => {
    const n = pNombre.trim();
    if (!n) return;
    const id = "p" + Date.now().toString(36);
    const prod = { id, nombre: n, categoria: pCat.trim() || "Otros", unidad: pUnidad.trim() || "ud" };
    upd((ev) => ({ ...ev, productos: [...ev.productos, prod] }));
    setPNombre(""); setPCat(""); setPUnidad("");
  };
  const removeProd = (pid) => upd((ev) => ({
    ...ev,
    productos: ev.productos.filter((p) => p.id !== pid),
    jornadas: ev.jornadas.map((j) => {
      const c = {};
      Object.keys(j.conteo).forEach((u) => { c[u] = { ...j.conteo[u] }; delete c[u][pid]; });
      return { ...j, conteo: c };
    }),
  }));

  return (
    <div>
      <div style={styles.dimText}>Esta configuración es común a todas las jornadas del evento.</div>

      <div style={{ ...styles.formCard, marginTop: 0 }}>
        <div style={styles.formCardTitle}>Importar desde Excel</div>
        <div style={styles.dimText}>Añade ubicaciones y productos desde la plantilla. No duplica lo que ya exista en el evento.</div>
        <div style={styles.formRow}>
          <button onClick={() => { setImportMsg(null); fileRef.current?.click(); }} style={styles.importBtn}>↑ Importar Excel</button>
          <button onClick={descargarPlantillaBlanco} style={styles.smallBtn}>↓ Descargar plantilla</button>
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={onFile} style={{ display: "none" }} />
        {importMsg && <div style={{ color: COLORS.green, fontSize: 13 }}>{importMsg}</div>}
      </div>

      <div style={styles.sectionTitle}>Ubicaciones</div>
      <div style={styles.chipWrap}>
        {evento.ubicaciones.map((u) => (
          <span key={u} style={styles.chipEdit}>{u}<button onClick={() => removeUbic(u)} style={styles.xBtnSm}>×</button></span>
        ))}
        {evento.ubicaciones.length === 0 && <span style={styles.dimText}>Sin ubicaciones todavía</span>}
      </div>
      <div style={styles.formRow}>
        <input value={nuevaUbic} onChange={(e) => setNuevaUbic(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addUbic()} placeholder="Nueva ubicación" style={styles.textInput} />
        <button onClick={addUbic} style={styles.addBtn}>+ Añadir</button>
      </div>

      <div style={{ ...styles.sectionTitle, marginTop: 34 }}>Referencias de producto</div>
      {evento.productos.map((p) => (
        <div key={p.id} style={styles.listRow}>
          <span style={{ flex: 1, color: COLORS.cream }}>{p.nombre}</span>
          <span style={styles.metaTag}>{p.categoria} · {p.unidad}</span>
          <button onClick={() => removeProd(p.id)} style={styles.xBtnSm}>×</button>
        </div>
      ))}
      {evento.productos.length === 0 && <div style={styles.dimText}>Sin referencias todavía</div>}
      <div style={styles.formCard}>
        <input value={pNombre} onChange={(e) => setPNombre(e.target.value)} placeholder="Nombre del producto" style={styles.textInput} />
        <div style={styles.formRow}>
          <input value={pCat} onChange={(e) => setPCat(e.target.value)} placeholder="Categoría" style={styles.textInput} />
          <input value={pUnidad} onChange={(e) => setPUnidad(e.target.value)} placeholder="Unidad" style={styles.textInput} />
        </div>
        <button onClick={addProd} style={styles.addBtn}>+ Añadir producto</button>
      </div>
    </div>
  );
}

function JornadaSelector({ evento, jornadaActivaId, setJornadaActivaId }) {
  if (evento.jornadas.length === 0) return null;
  // Varios días → desplegable (evita un muro de chips). Un día → etiqueta.
  if (evento.tipo === "multi") {
    return (
      <div style={{ marginBottom: 18 }}>
        <label style={{ ...styles.fieldLabel, display: "block", marginBottom: 6 }}>Jornada</label>
        <select value={jornadaActivaId || ""} onChange={(e) => setJornadaActivaId(e.target.value)} style={styles.select}>
          {evento.jornadas.map((j) => {
            const { total, conf, completo } = jornadaEstado(evento, j);
            const pend = jornadaPendienteConfirmar(evento, j).length;
            const marca = completo ? " ✓" : (pend > 0 ? " ⚠ sin confirmar" : "");
            return (
              <option key={j.id} value={j.id}>{fechaLabel(j.fecha)} · {conf}/{total}{marca}{j.editable === false ? " · 🔒" : ""}</option>
            );
          })}
        </select>
      </div>
    );
  }
  return (
    <div style={styles.chipWrap}>
      {evento.jornadas.map((j) => (
        <button key={j.id} onClick={() => setJornadaActivaId(j.id)} style={{ ...styles.chip, ...(j.id === jornadaActivaId ? styles.chipActive : {}) }}>
          {fechaLabel(j.fecha)}
        </button>
      ))}
    </div>
  );
}

function ConteoView({ evento, role, upd, jornada, jornadaActivaId, setJornadaActivaId, ubicActiva, setUbicActiva, onGuardar }) {
  const [catActiva, setCatActiva] = useState("");
  const [movTipo, setMovTipo] = useState("entrada");
  const [movProd, setMovProd] = useState("");
  const [movCant, setMovCant] = useState("");
  const [movDest, setMovDest] = useState("");
  const [movCom, setMovCom] = useState("");
  const [importInvMsg, setImportInvMsg] = useState(null);
  const invFileRef = React.useRef(null);
  const [mostrarPapel, setMostrarPapel] = useState(false);
  const [mostrarMov, setMostrarMov] = useState(false);
  if (evento.ubicaciones.length === 0 || evento.productos.length === 0)
    return <div style={styles.empty}>Necesitas ubicaciones y referencias (Configuración) para el conteo.</div>;
  if (evento.jornadas.length === 0)
    return <div style={styles.empty}>Crea al menos una jornada para registrar el conteo.</div>;
  if (!jornada) return (<div><JornadaSelector evento={evento} jornadaActivaId={jornadaActivaId} setJornadaActivaId={setJornadaActivaId} /><div style={styles.empty}>Selecciona una jornada.</div></div>);

  const getCell = (ubic, pid) => jornada.conteo?.[ubic]?.[pid] || { inicial: 0, final: 0 };
  const esAdminC = role === "admin";
  const jornadaEditable = esAdminC || jornada.editable !== false;
  const puedeContar = jornadaEditable;                                          // los 3 roles cuentan el Final
  const puedeInicial = jornadaEditable && (esAdminC || role === "inventario");  // solo Admin e Inventario tocan el Inicial / movimientos
  const puedeOperar = puedeContar;                                             // puede registrar algo en esta jornada

  const confirmadoUbic = (u) => !!(jornada.confirmado && jornada.confirmado[u]);
  const confirmado = confirmadoUbic(ubicActiva);
  const confirmadasCount = evento.ubicaciones.filter(confirmadoUbic).length;
  const setConfirmado = (val) => {
    upd((ev) => ({
      ...ev,
      jornadas: ev.jornadas.map((j) => {
        if (j.id !== jornada.id) return j;
        const conf = { ...(j.confirmado || {}) };
        const rh = { ...(j.realizadoHora || {}) };
        if (val) { conf[ubicActiva] = true; rh[ubicActiva] = horaActual(); }   // la hora se marca al confirmar
        else { delete conf[ubicActiva]; delete rh[ubicActiva]; }                // al reabrir se borra
        return { ...j, confirmado: conf, realizadoHora: rh };
      }),
    }));
  };

  const realizadoPor = (jornada.realizadoPor && jornada.realizadoPor[ubicActiva]) || "";
  const realizadoHora = (jornada.realizadoHora && jornada.realizadoHora[ubicActiva]) || "";
  const horaActual = () => {
    const d = new Date();
    return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
  };
  const setRealizadoPor = (nombre) => {
    upd((ev) => ({
      ...ev,
      jornadas: ev.jornadas.map((j) => {
        if (j.id !== jornada.id) return j;
        const rp = { ...(j.realizadoPor || {}) };
        rp[ubicActiva] = nombre;
        return { ...j, realizadoPor: rp };
      }),
    }));
  };
  const puedeAvanzar = !puedeOperar || !!realizadoPor.trim();

  const limpiarInventario = () => {
    if (!puedeContar) return;
    if (!confirm(`¿Limpiar el Final de ${ubicActiva} en ${fechaLabel(jornada.fecha)}?\nSe pondrá a cero el Final de todos sus productos (el Inicial se mantiene) y se quitará la confirmación.`)) return;
    upd((ev) => ({
      ...ev,
      jornadas: ev.jornadas.map((j) => {
        if (j.id !== jornada.id) return j;
        const conteo = { ...j.conteo };
        const cu = { ...(conteo[ubicActiva] || {}) };
        Object.keys(cu).forEach((pid) => { cu[pid] = { ...cu[pid], final: 0 }; });
        conteo[ubicActiva] = cu;
        const conf = { ...(j.confirmado || {}) }; delete conf[ubicActiva];
        const rh = { ...(j.realizadoHora || {}) }; delete rh[ubicActiva];
        return { ...j, conteo, confirmado: conf, realizadoHora: rh };
      }),
    }));
  };

  const prodNombre = (pid) => { const p = evento.productos.find((x) => x.id === pid); return p ? p.nombre : pid; };
  const movimientosUbic = (jornada.movimientos || []).filter((m) => m.ubic === ubicActiva || (m.tipo === "traspaso" && m.destino === ubicActiva));
  const addMovimiento = () => {
    if (!puedeInicial) return;
    const cant = Number(movCant);
    if (!movProd || !cant || cant <= 0) return;
    if (movTipo === "traspaso" && (!movDest || movDest === ubicActiva)) return;
    const mov = {
      id: "m" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5),
      tipo: movTipo, productoId: movProd, ubic: ubicActiva,
      destino: movTipo === "traspaso" ? movDest : "",
      cantidad: cant, comentario: movCom.trim(), hora: horaActual(),
    };
    upd((ev) => ({ ...ev, jornadas: ev.jornadas.map((j) => (j.id === jornada.id ? { ...j, movimientos: [...(j.movimientos || []), mov] } : j)) }));
    setMovCant(""); setMovCom(""); setMovDest("");
  };
  const removeMovimiento = (id) => upd((ev) => ({ ...ev, jornadas: ev.jornadas.map((j) => (j.id === jornada.id ? { ...j, movimientos: (j.movimientos || []).filter((m) => m.id !== id) } : j)) }));

  const importarInventario = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportInvMsg("Leyendo archivo…");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ubicByNorm = {}; evento.ubicaciones.forEach((u) => { ubicByNorm[normTxt(u)] = u; });
      const prodByNorm = {}; evento.productos.forEach((p) => { prodByNorm[normTxt(p.nombre)] = p.id; });
      const cambios = {};
      let aplicadas = 0;
      wb.SheetNames.forEach((sheetName) => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "" });
        // La ubicación de esta hoja: por nombre de hoja, o por la fila "Ubicación: X".
        let ubic = ubicByNorm[normTxt(sheetName)];
        if (!ubic) {
          for (const r of rows) {
            const c0 = String(r[0] || "");
            if (normTxt(c0).startsWith("ubicacion")) {
              const val = (c0.includes(":") ? c0.split(":").slice(1).join(":") : String(r[1] || "")).trim();
              if (ubicByNorm[normTxt(val)]) { ubic = ubicByNorm[normTxt(val)]; break; }
            }
          }
        }
        if (!ubic) return;
        rows.forEach((r) => {
          const nombre = String(r[0] || "").trim();
          if (!nombre) return;
          const pid = prodByNorm[normTxt(nombre)];
          if (!pid) return; // familias, cabeceras y "Ubicación:" se ignoran solas
          const iniRaw = r[2], finRaw = r[3];
          const set = {};
          if (iniRaw !== "" && iniRaw != null) set.inicial = Math.max(0, Number(iniRaw) || 0);
          if (finRaw !== "" && finRaw != null) set.final = Math.max(0, Number(finRaw) || 0);
          if (!Object.keys(set).length) return;
          cambios[ubic] = cambios[ubic] || {};
          cambios[ubic][pid] = { ...(cambios[ubic][pid] || {}), ...set };
          aplicadas++;
        });
      });
      if (aplicadas === 0) {
        setImportInvMsg("No se encontraron datos válidos. Usa la plantilla descargada de este mismo evento.");
        e.target.value = ""; return;
      }
      upd((ev) => ({
        ...ev,
        jornadas: ev.jornadas.map((j) => {
          if (j.id !== jornada.id) return j;
          const conteo = { ...j.conteo };
          Object.keys(cambios).forEach((u) => {
            const cu = { ...(conteo[u] || {}) };
            Object.keys(cambios[u]).forEach((pid) => { cu[pid] = { ...(cu[pid] || { inicial: 0, final: 0 }), ...cambios[u][pid] }; });
            conteo[u] = cu;
          });
          return { ...j, conteo };
        }),
      }));
      setImportInvMsg(`Importado en ${fechaLabel(jornada.fecha)}: ${aplicadas} líneas.`);
      e.target.value = "";
    } catch (err) {
      console.error(err);
      setImportInvMsg("No se pudo leer el archivo. Asegúrate de que es un .xlsx válido.");
      e.target.value = "";
    }
  };

  const setValor = (pid, campo, valor) => {
    if (campo === "inicial" && !puedeInicial) return;
    if (campo === "final" && !puedeContar) return;
    const v = valor === "" ? 0 : Math.max(0, Number(valor));
    upd((ev) => {
      // Localiza la jornada siguiente por fecha: su Inicial heredará este Final.
      const orden = [...ev.jornadas].sort((a, b) => (a.fecha || "").localeCompare(b.fecha || ""));
      const idx = orden.findIndex((j) => j.id === jornada.id);
      const siguienteId = campo === "final" && idx >= 0 && idx < orden.length - 1 ? orden[idx + 1].id : null;

      const ponerCampo = (j, campoSet, valorSet) => {
        const conteo = { ...j.conteo };
        const ubic = { ...(conteo[ubicActiva] || {}) };
        ubic[pid] = { ...(ubic[pid] || { inicial: 0, final: 0 }), [campoSet]: valorSet };
        conteo[ubicActiva] = ubic;
        return { ...j, conteo };
      };

      return {
        ...ev,
        jornadas: ev.jornadas.map((j) => {
          if (j.id === jornada.id) return ponerCampo(j, campo, v);
          if (siguienteId && j.id === siguienteId) return ponerCampo(j, "inicial", v); // final de hoy = inicial de mañana
          return j;
        }),
      };
    });
  };

  const categorias = [...new Set(evento.productos.map((p) => p.categoria))];
  const catSel = categorias.includes(catActiva) ? catActiva : "";

  return (
    <div>
      <JornadaSelector evento={evento} jornadaActivaId={jornadaActivaId} setJornadaActivaId={setJornadaActivaId} />

      <AvisosJornadas evento={evento} jornadaActivaId={jornadaActivaId} onIr={(id, u) => { setJornadaActivaId(id); if (u) setUbicActiva(u); }} />

      <div style={{ marginBottom: 16 }}>
        <button onClick={() => setMostrarPapel((v) => !v)} style={styles.linkBtn}>{mostrarPapel ? "▲" : "▼"} Plantilla / importar</button>
        {mostrarPapel && (
          <div style={{ ...styles.formCard, marginTop: 10 }}>
            <div style={styles.dimText}>Descarga la plantilla en blanco para contar a mano. Cuando la tengas rellena, impórtala para volcar el Inicial y el Final en la jornada seleccionada ({fechaLabel(jornada.fecha)}).</div>
            <div style={styles.formRow}>
              <button onClick={() => descargarPlantillaInventario(evento)} style={styles.smallBtn}>↓ Descargar plantilla</button>
              {puedeInicial && <button onClick={() => { setImportInvMsg(null); invFileRef.current?.click(); }} style={styles.importBtn}>↑ Importar inventario</button>}
            </div>
            <input ref={invFileRef} type="file" accept=".xlsx,.xls" onChange={importarInventario} style={{ display: "none" }} />
            {importInvMsg && <div style={{ color: COLORS.green, fontSize: 13 }}>{importInvMsg}</div>}
          </div>
        )}
      </div>

      <div style={styles.chipWrap}>
        {evento.ubicaciones.map((u) => (
          <button key={u} onClick={() => setUbicActiva(u)} style={{ ...styles.chip, ...(u === ubicActiva ? styles.chipActive : {}), ...(confirmadoUbic(u) && u !== ubicActiva ? styles.chipDone : {}) }}>{u}{confirmadoUbic(u) ? " ✓" : ""}</button>
        ))}
      </div>

      <div style={{ marginBottom: 14 }}>
        <label style={{ ...styles.fieldLabel, display: "block", marginBottom: 6 }}>Realizado por *</label>
        <input type="text" value={realizadoPor} placeholder="Nombre de quien cuenta" disabled={!puedeOperar} onChange={(e) => setRealizadoPor(e.target.value)} style={{ ...styles.textInput, width: "100%", ...(puedeOperar ? {} : styles.inputDisabled) }} />
      </div>

      {!puedeAvanzar ? (
        <div style={{ ...styles.empty, color: COLORS.gold }}>Indica quién realiza el inventario (obligatorio) para empezar a contar.</div>
      ) : (
      <>
      {!jornadaEditable && (
        <div style={{ ...styles.empty, color: COLORS.gold, marginBottom: 16 }}>🔒 Esta jornada está en solo lectura. Un administrador debe habilitarla para poder modificarla.</div>
      )}

      <div style={{ marginBottom: 18 }}>
        <label style={{ ...styles.fieldLabel, display: "block", marginBottom: 6 }}>Familia de productos</label>
        <select value={catSel} onChange={(e) => setCatActiva(e.target.value)} style={styles.select}>
          <option value="">— Selecciona una familia —</option>
          {categorias.map((cat) => (
            <option key={cat} value={cat}>{cat}</option>
          ))}
        </select>
      </div>

      {!catSel ? (
        <div style={{ ...styles.dimText, marginBottom: 20 }}>Elige una familia para ver sus referencias.</div>
      ) : (
        <div style={{ marginBottom: 26 }}>
          <div style={styles.dimText}>El Inicial es el stock con el que empezó la ubicación. Si hay entradas, salidas o traspasos, se muestran debajo del producto con el disponible resultante. Consumo = Disponible − Final.</div>
          <div style={styles.tableHead}>
            <span style={{ flex: 2 }}>Producto</span>
            <span style={styles.colNum}>Inicial</span><span style={styles.colNum}>Final</span><span style={styles.colNum}>Consumo</span>
          </div>
          {evento.productos.filter((p) => p.categoria === catSel).map((p) => {
            const c = getCell(ubicActiva, p.id);
            const t = celdaTotales(jornada, ubicActiva, p.id);
            const sospechosa = t.fin > t.ini;
            return (
              <div key={p.id} style={styles.row}>
                <span style={{ flex: 2 }}>
                  <span style={{ color: COLORS.cream }}>{p.nombre}</span><span style={styles.unidad}> · {p.unidad}</span>
                  {(t.ent > 0 || t.sal > 0) && <div style={styles.movHint}>{t.ent > 0 ? `▲ +${t.ent} ` : ""}{t.sal > 0 ? `▼ −${t.sal} ` : ""}· disponible {t.ini}</div>}
                  {sospechosa && <div style={{ ...styles.movHint, color: COLORS.red, fontWeight: 600 }}>⛔ Final ({t.fin}) mayor que el disponible ({t.ini}) — revisa el número</div>}
                </span>
                <input type="number" min="0" value={c.inicial || ""} placeholder="0" disabled={!puedeInicial} onChange={(e) => setValor(p.id, "inicial", e.target.value)} style={{ ...styles.input, ...(puedeInicial ? {} : styles.inputDisabled) }} />
                <input type="number" min="0" value={c.final || ""} placeholder="0" disabled={!puedeContar} onChange={(e) => setValor(p.id, "final", e.target.value)} style={{ ...styles.input, ...(puedeContar ? {} : styles.inputDisabled), ...(sospechosa ? styles.inputError : {}) }} />
                <span style={{ ...styles.colNum, color: t.con > 0 ? COLORS.gold : COLORS.dim, fontWeight: 600 }}>{t.con}</span>
              </div>
            );
          })}
        </div>
      )}

      <div style={styles.formCard}>
        <div onClick={() => setMostrarMov((v) => !v)} style={{ ...styles.collapseHeader, marginBottom: mostrarMov ? 10 : 0 }}>
          <span style={styles.formCardTitle}>Movimientos · {ubicActiva}</span>
          <span style={styles.collapseMeta}>{movimientosUbic.length > 0 ? `${movimientosUbic.length} ` : ""}{mostrarMov ? "▲" : "▼"}</span>
        </div>
        {mostrarMov && (
        <>
        <div style={styles.dimText}>Entradas (reposición), salidas (mermas o cesiones) y traspasos a otra ubicación. Cada uno con su comentario.</div>
        {movimientosUbic.length === 0 && <div style={{ color: COLORS.dim, fontSize: 13 }}>Sin movimientos en esta ubicación.</div>}
        {movimientosUbic.map((m) => {
          const recibido = m.tipo === "traspaso" && m.destino === ubicActiva;
          const etiqueta = m.tipo === "entrada" ? "🔵 Entrada"
            : m.tipo === "salida" ? "🔴 Salida"
            : recibido ? `🔁 Recibido de ${m.ubic}` : `🔁 Traspaso → ${m.destino}`;
          const signo = (m.tipo === "entrada" || recibido) ? "+" : "−";
          return (
            <div key={m.id} style={styles.movRow}>
              <div style={{ flex: 1 }}>
                <div style={{ color: COLORS.cream, fontSize: 13 }}>{etiqueta} · {prodNombre(m.productoId)} · {signo}{m.cantidad}</div>
                {(m.comentario || m.hora) && <div style={{ color: COLORS.dim, fontSize: 12 }}>{m.hora ? m.hora + " · " : ""}{m.comentario}</div>}
              </div>
              {puedeInicial && !recibido && <button onClick={() => removeMovimiento(m.id)} style={styles.xBtnSm}>×</button>}
            </div>
          );
        })}
        {puedeInicial && (
          <div style={{ borderTop: `1px solid ${COLORS.line}`, marginTop: 10, paddingTop: 12, display: "flex", flexDirection: "column", gap: 8 }}>
            <select value={movTipo} onChange={(e) => setMovTipo(e.target.value)} style={styles.select}>
              <option value="entrada">Entrada (reposición)</option>
              <option value="salida">Salida (merma / cesión)</option>
              <option value="traspaso">Traspaso a otra ubicación</option>
            </select>
            <select value={movProd} onChange={(e) => setMovProd(e.target.value)} style={styles.select}>
              <option value="">— Producto —</option>
              {evento.productos.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
            </select>
            <div style={styles.formRow}>
              <input type="number" min="0" value={movCant} placeholder="Cantidad" onChange={(e) => setMovCant(e.target.value)} style={styles.textInput} />
              {movTipo === "traspaso" && (
                <select value={movDest} onChange={(e) => setMovDest(e.target.value)} style={styles.textInput}>
                  <option value="">— Destino —</option>
                  {evento.ubicaciones.filter((u) => u !== ubicActiva).map((u) => <option key={u} value={u}>{u}</option>)}
                </select>
              )}
            </div>
            <input type="text" value={movCom} placeholder="Comentario (ej. pedido de camerinos)" onChange={(e) => setMovCom(e.target.value)} style={styles.textInput} />
            <button onClick={addMovimiento} style={styles.addBtn}>+ Añadir movimiento</button>
          </div>
        )}
        </>
        )}
      </div>

      <div style={styles.confirmBar}>
        <span style={{ fontSize: 13 }}>
          {confirmado
            ? <span style={{ color: COLORS.green, fontWeight: 600 }}>✓ {ubicActiva} confirmada{realizadoHora ? ` · ${realizadoHora}` : ""}</span>
            : <span style={{ color: COLORS.dim }}>{ubicActiva}: pendiente de confirmar</span>}
          <span style={{ color: COLORS.dim }}>{"  ·  "}{confirmadasCount}/{evento.ubicaciones.length} ubicaciones confirmadas</span>
        </span>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", justifyContent: "flex-end" }}>
          {puedeContar && <button onClick={limpiarInventario} style={styles.linkBtn}>🧹 Limpiar Final</button>}
          {puedeContar && (confirmado
            ? <button onClick={() => setConfirmado(false)} style={styles.smallBtn}>Reabrir</button>
            : <button onClick={() => setConfirmado(true)} style={styles.addBtn}>✓ Confirmar</button>
          )}
        </div>
      </div>
      </>
      )}

      <button onClick={() => descargarJornadaExcel(evento, jornada)} style={{ ...styles.smallBtn, width: "100%", marginTop: 18 }}>↓ Descargar inventario del día (Excel)</button>
    </div>
  );
}

function ResumenView({ evento, setTab, setJornadaActivaId, setUbicActiva }) {
  if (evento.jornadas.length === 0) return <div style={styles.empty}>Sin jornadas todavía.</div>;
  if (evento.productos.length === 0) return <div style={styles.empty}>Sin referencias todavía.</div>;
  const irAJornada = (id, u) => { if (setJornadaActivaId) setJornadaActivaId(id); if (u && setUbicActiva) setUbicActiva(u); if (setTab) setTab("conteo"); };

  // Totales por referencia, sumando todas las jornadas y ubicaciones.
  const totalRef = (pid) => {
    let ini = 0, fin = 0, ent = 0, sal = 0, con = 0;
    evento.jornadas.forEach((j) => evento.ubicaciones.forEach((u) => {
      const t = celdaTotales(j, u, pid);
      ini += t.base; fin += t.fin; ent += t.ent; sal += t.sal; con += t.con;
    }));
    return { ini, fin, ent, sal, con };
  };

  return (
    <div>
      <AvisosJornadas evento={evento} jornadaActivaId={null} onIr={irAJornada} />
      <div style={styles.dimText}>Total por referencia, sumando todas las jornadas y ubicaciones del evento.</div>
      <button onClick={() => descargarResumenExcel(evento)} style={{ ...styles.exportBtn, marginBottom: 18 }}>↓ Descargar resumen (Excel)</button>
      <div style={{ overflowX: "auto" }}>
        <table style={styles.matrix}>
          <thead><tr>
            <th style={{ ...styles.th, textAlign: "left" }}>Producto</th>
            <th style={styles.th}>Unidad</th>
            <th style={styles.th}>Inicial</th>
            <th style={styles.th}>Entradas</th>
            <th style={styles.th}>Salidas</th>
            <th style={styles.th}>Final</th>
            <th style={{ ...styles.th, color: COLORS.gold }}>Consumo</th>
          </tr></thead>
          <tbody>
            {evento.productos.map((p) => {
              const t = totalRef(p.id);
              return (
                <tr key={p.id}>
                  <td style={{ ...styles.td, textAlign: "left", color: COLORS.cream }}>{p.nombre}</td>
                  <td style={{ ...styles.td, color: COLORS.dim }}>{p.unidad}</td>
                  <td style={{ ...styles.td, color: COLORS.cream }}>{t.ini}</td>
                  <td style={{ ...styles.td, color: COLORS.cream }}>{t.ent}</td>
                  <td style={{ ...styles.td, color: COLORS.cream }}>{t.sal}</td>
                  <td style={{ ...styles.td, color: COLORS.cream }}>{t.fin}</td>
                  <td style={{ ...styles.td, color: COLORS.gold, fontWeight: 600 }}>{t.con}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function nombreSeguro(nombre) {
  return (nombre || "evento").replace(/[^\w\sáéíóúñ-]/gi, "").replace(/\s+/g, "_").slice(0, 40) || "evento";
}

// Totales de una celda (ubicación + producto) en una jornada, con movimientos.
// El stock disponible (ini) = arrastre (base) + entradas − salidas.
// Consumo = disponible − Final.
function celdaTotales(jornada, ubic, pid) {
  const c = (jornada.conteo && jornada.conteo[ubic] && jornada.conteo[ubic][pid]) || { inicial: 0, final: 0 };
  const base = c.inicial || 0, fin = c.final || 0;
  let ent = 0, sal = 0;
  for (const m of (jornada.movimientos || [])) {
    if (m.productoId !== pid) continue;
    const cant = Number(m.cantidad) || 0;
    if (m.tipo === "entrada" && m.ubic === ubic) ent += cant;
    else if (m.tipo === "traspaso" && m.destino === ubic) ent += cant;
    else if (m.tipo === "salida" && m.ubic === ubic) sal += cant;
    else if (m.tipo === "traspaso" && m.ubic === ubic) sal += cant;
  }
  const ini = base + ent - sal; // stock disponible
  return { base, ini, fin, ent, sal, con: Math.max(0, ini - fin) };
}

// Construye la hoja de inventario de UNA jornada en formato cruzado:
// una fila por producto, y por cada ubicación (más un grupo Total) las
// columnas Inicial / Final / Consumo, con cabecera de dos niveles.
function hojaInventarioJornada(evento, jornada) {
  const ubic = evento.ubicaciones;
  const grupos = ["Total", ...ubic];

  const fila1 = ["", "", ""];
  grupos.forEach((g) => fila1.push(g, "", ""));
  const fila2 = ["Categoría", "Producto", "Unidad"];
  grupos.forEach(() => fila2.push("Inicial", "Final", "Consumo"));

  const filas = [fila1, fila2];
  evento.productos.forEach((p) => {
    const row = [p.categoria, p.nombre, p.unidad];
    let tIni = 0, tFin = 0, tCon = 0;
    ubic.forEach((u) => { const t = celdaTotales(jornada, u, p.id); tIni += t.ini; tFin += t.fin; tCon += t.con; });
    row.push(tIni, tFin, tCon);
    ubic.forEach((u) => { const t = celdaTotales(jornada, u, p.id); row.push(t.ini, t.fin, t.con); });
    filas.push(row);
  });

  const ws = XLSX.utils.aoa_to_sheet(filas);
  // Combinar cada etiqueta de grupo (Total, Barra 1…) sobre sus 3 columnas.
  ws["!merges"] = grupos.map((g, gi) => ({ s: { r: 0, c: 3 + gi * 3 }, e: { r: 0, c: 3 + gi * 3 + 2 } }));
  ws["!cols"] = [{ wch: 14 }, { wch: 30 }, { wch: 10 }, ...grupos.flatMap(() => [{ wch: 9 }, { wch: 9 }, { wch: 10 }])];
  return ws;
}

// Descarga el inventario de UNA jornada en el formato cruzado.
function descargarJornadaExcel(evento, jornada) {
  if (!jornada) { alert("Selecciona una jornada para descargar su inventario."); return; }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, hojaInventarioJornada(evento, jornada), "Inventario");
  const resp = evento.ubicaciones.map((u) => ({
    "Ubicación": u,
    "Realizado por": (jornada.realizadoPor && jornada.realizadoPor[u]) || "",
    "Hora": (jornada.realizadoHora && jornada.realizadoHora[u]) || "",
    "Confirmado": (jornada.confirmado && jornada.confirmado[u]) ? "Sí" : "No",
  }));
  const wsR = XLSX.utils.json_to_sheet(resp);
  wsR["!cols"] = [{ wch: 20 }, { wch: 28 }, { wch: 10 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, wsR, "Responsables");

  const movs = (jornada.movimientos || []).map((m) => ({
    Tipo: m.tipo === "entrada" ? "Entrada" : m.tipo === "salida" ? "Salida" : "Traspaso",
    Producto: (evento.productos.find((p) => p.id === m.productoId) || {}).nombre || m.productoId,
    Cantidad: m.cantidad, Origen: m.ubic, Destino: m.destino || "",
    Comentario: m.comentario || "", Hora: m.hora || "",
  }));
  if (movs.length) {
    const wsM = XLSX.utils.json_to_sheet(movs);
    wsM["!cols"] = [{ wch: 10 }, { wch: 26 }, { wch: 10 }, { wch: 16 }, { wch: 16 }, { wch: 34 }, { wch: 8 }];
    XLSX.utils.book_append_sheet(wb, wsM, "Movimientos");
  }
  XLSX.writeFile(wb, `Eventrack_${nombreSeguro(evento.nombre)}_${jornada.fecha || "sinfecha"}.xlsx`);
}

// Plantilla de inventario para papel: una HOJA por ubicación, con la barra
// como cabecera y los productos agrupados por familia (familia como título).
function descargarPlantillaInventario(evento) {
  if (!evento.ubicaciones.length || !evento.productos.length) {
    alert("Configura ubicaciones y productos antes de descargar la plantilla.");
    return;
  }
  const wb = XLSXStyle.utils.book_new();
  const cats = [...new Set(evento.productos.map((p) => p.categoria))].sort((a, b) => String(a).localeCompare(String(b)));
  const usados = {};
  // Estilos reutilizables
  const thin = { style: "thin", color: { rgb: "000000" } };
  const border = { top: thin, bottom: thin, left: thin, right: thin };
  evento.ubicaciones.forEach((u) => {
    const aoa = []; const merges = []; const meta = []; // meta: tipo de fila para estilar
    aoa.push([`Ubicación: ${u}`]); meta.push("title"); merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } });
    aoa.push([]); meta.push(null);
    cats.forEach((cat) => {
      const prods = evento.productos.filter((p) => p.categoria === cat).sort((a, b) => a.nombre.localeCompare(b.nombre));
      if (!prods.length) return;
      merges.push({ s: { r: aoa.length, c: 0 }, e: { r: aoa.length, c: 3 } });
      aoa.push([cat]); meta.push("cat");
      aoa.push(["Producto", "Unidad", "Inicial", "Final"]); meta.push("head");
      prods.forEach((p) => { aoa.push([p.nombre, p.unidad, "", ""]); meta.push("prod"); });
      aoa.push([]); meta.push(null);
    });
    const ws = XLSXStyle.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 42 }, { wch: 14 }, { wch: 12 }, { wch: 12 }];
    ws["!merges"] = merges;
    // Aplica bordes y resaltados según el tipo de cada fila
    const rows = [];
    meta.forEach((t, r) => {
      if (!t) { rows.push({ hpt: 6 }); return; } // separador fino entre familias
      rows.push({ hpt: t === "prod" ? 22 : 18 });
      for (let c = 0; c <= 3; c++) {
        const ref = XLSXStyle.utils.encode_cell({ r, c });
        if (!ws[ref]) ws[ref] = { t: "s", v: "" };
        const al = { vertical: "center" };
        if (t === "title") ws[ref].s = { border, font: { bold: true, sz: 13 }, alignment: al };
        else if (t === "cat") ws[ref].s = { border, font: { bold: true, sz: 11 }, fill: { fgColor: { rgb: "F1F2F6" } }, alignment: al };
        else if (t === "head") ws[ref].s = { border, font: { bold: true }, fill: { fgColor: { rgb: "EEF0FF" } }, alignment: { ...al, horizontal: c >= 2 ? "center" : "left" } };
        else ws[ref].s = { border, alignment: { ...al, horizontal: c >= 2 ? "center" : "left" } };
      }
    });
    ws["!rows"] = rows;
    // Nombre de hoja válido (sin caracteres prohibidos, máx 31, único)
    let name = String(u).replace(/[:\\/?*\[\]]/g, " ").trim().slice(0, 28) || "Barra";
    const base = name; let n = 2;
    while (usados[name.toLowerCase()]) { name = (base.slice(0, 25) + " " + n); n++; }
    usados[name.toLowerCase()] = true;
    XLSXStyle.utils.book_append_sheet(wb, ws, name);
  });
  XLSXStyle.writeFile(wb, `Eventrack_${nombreSeguro(evento.nombre)}_plantilla_inventario.xlsx`);
}

// Descarga el resumen total por referencia (suma de todas las jornadas y ubicaciones).
function descargarResumenExcel(evento) {
  const wb = XLSX.utils.book_new();
  const filas = evento.productos.map((p) => {
    let ini = 0, fin = 0, ent = 0, sal = 0, con = 0;
    evento.jornadas.forEach((j) => evento.ubicaciones.forEach((u) => {
      const t = celdaTotales(j, u, p.id);
      ini += t.base; fin += t.fin; ent += t.ent; sal += t.sal; con += t.con;
    }));
    return { Producto: p.nombre, "Categoría": p.categoria, Unidad: p.unidad, "Inicial total": ini, "Entradas": ent, "Salidas": sal, "Final total": fin, "Consumo total": con };
  });
  const ws = XLSX.utils.json_to_sheet(filas);
  ws["!cols"] = [{ wch: 26 }, { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, ws, "Total por referencia");
  XLSX.writeFile(wb, `Eventrack_${nombreSeguro(evento.nombre)}_resumen.xlsx`);
}

function ExportButton({ evento }) {
  const exportar = () => {
    const wb = XLSX.utils.book_new();
    const cons = (j, u, pid) => celdaTotales(j, u, pid).con;

    if (evento.jornadas.length) {
      const filas = evento.jornadas.map((j) => {
        const consumo = evento.ubicaciones.reduce((a, u) => a + evento.productos.reduce((aa, p) => aa + cons(j, u, p.id), 0), 0);
        return { Jornada: j.fecha || "(sin fecha)", "Consumo total": consumo };
      });
      const ws = XLSX.utils.json_to_sheet(filas);
      ws["!cols"] = [{ wch: 14 }, { wch: 14 }];
      XLSX.utils.book_append_sheet(wb, ws, "Resumen jornadas");
    }

    const detalle = [];
    evento.jornadas.forEach((j) => evento.ubicaciones.forEach((u) => evento.productos.forEach((p) => {
      const t = celdaTotales(j, u, p.id);
      detalle.push({ Jornada: j.fecha, "Ubicación": u, "Categoría": p.categoria, Producto: p.nombre, Unidad: p.unidad, Inicial: t.base, Entradas: t.ent, Salidas: t.sal, Final: t.fin, Consumo: t.con });
    })));
    if (detalle.length) {
      const ws = XLSX.utils.json_to_sheet(detalle);
      ws["!cols"] = [{ wch: 12 }, { wch: 18 }, { wch: 14 }, { wch: 26 }, { wch: 10 }, { wch: 9 }, { wch: 9 }, { wch: 9 }, { wch: 9 }, { wch: 10 }];
      XLSX.utils.book_append_sheet(wb, ws, "Detalle inventario");
    }

    if (evento.productos.length) {
      const acc = evento.productos.map((p) => {
        let ini = 0, fin = 0, ent = 0, sal = 0, con = 0;
        evento.jornadas.forEach((j) => evento.ubicaciones.forEach((u) => {
          const t = celdaTotales(j, u, p.id);
          ini += t.base; fin += t.fin; ent += t.ent; sal += t.sal; con += t.con;
        }));
        return { Producto: p.nombre, "Categoría": p.categoria, Unidad: p.unidad, "Inicial total": ini, "Entradas": ent, "Salidas": sal, "Final total": fin, "Consumo total": con };
      });
      const ws = XLSX.utils.json_to_sheet(acc);
      ws["!cols"] = [{ wch: 26 }, { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 14 }];
      XLSX.utils.book_append_sheet(wb, ws, "Total por referencia");
    }

    const movsAll = [];
    evento.jornadas.forEach((j) => (j.movimientos || []).forEach((m) => {
      movsAll.push({
        Jornada: j.fecha, Tipo: m.tipo === "entrada" ? "Entrada" : m.tipo === "salida" ? "Salida" : "Traspaso",
        Producto: (evento.productos.find((p) => p.id === m.productoId) || {}).nombre || m.productoId,
        Cantidad: m.cantidad, Origen: m.ubic, Destino: m.destino || "", Comentario: m.comentario || "", Hora: m.hora || "",
      });
    }));
    if (movsAll.length) {
      const ws = XLSX.utils.json_to_sheet(movsAll);
      ws["!cols"] = [{ wch: 12 }, { wch: 10 }, { wch: 26 }, { wch: 10 }, { wch: 16 }, { wch: 16 }, { wch: 34 }, { wch: 8 }];
      XLSX.utils.book_append_sheet(wb, ws, "Movimientos");
    }

    if (wb.SheetNames.length === 0) { alert("No hay datos que exportar todavía."); return; }
    const safe = evento.nombre.replace(/[^\w\sáéíóúñ-]/gi, "").replace(/\s+/g, "_").slice(0, 40);
    XLSX.writeFile(wb, `Eventrack_${safe || "evento"}.xlsx`);
  };

  return <button onClick={exportar} style={styles.exportBtn}>↓ Exportar Excel</button>;
}

function Login({ onLogin }) {
  const [code, setCode] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  const entrar = async () => {
    const c = code.trim();
    if (!c) return;
    setBusy(true); setErr("");
    try {
      const r = await login(c);
      if (r === "admin" || r === "inventario" || r === "contador") { onLogin(r); return; }
      setErr("Código incorrecto. Inténtalo de nuevo.");
    } catch (e) {
      setErr("No se pudo conectar. Revisa tu conexión.");
    }
    setBusy(false);
  };

  return (
    <div style={{ ...styles.app, display: "flex", flexDirection: "column", justifyContent: "center", alignItems: "center", minHeight: "100vh" }}>
      <style>{globalCSS}</style>
      <img src="/icon.png" alt="Eventrack" style={{ width: 96, height: 96, objectFit: "contain", marginBottom: 8 }} />
      <h1 style={{ ...styles.title, fontSize: 40, marginBottom: 6 }}>Eventrack</h1>
      <p style={{ color: COLORS.dim, marginBottom: 22, fontSize: 14 }}>Introduce tu código de acceso</p>
      <div style={{ ...styles.formCard, width: "100%", maxWidth: 320, marginTop: 0 }}>
        <input
          type="password"
          value={code}
          placeholder="Código de acceso"
          onChange={(e) => setCode(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && entrar()}
          style={styles.textInput}
          autoFocus
        />
        <button onClick={entrar} disabled={busy} style={styles.addBtn}>{busy ? "Entrando…" : "Entrar"}</button>
        {err && <div style={{ color: COLORS.red, fontSize: 13 }}>{err}</div>}
      </div>
    </div>
  );
}

function PapeleraView({ eventos, onClose, onRestore, onHardDelete }) {
  const borrados = eventos.filter((ev) => ev.borrado);
  return (
    <div>
      <button onClick={onClose} style={styles.volverBtn}>‹ Volver</button>
      <div style={{ ...styles.sectionTitle, marginTop: 14 }}>Papelera</div>
      <div style={styles.dimText}>Eventos eliminados. Puedes restaurarlos o borrarlos definitivamente (esto último no se puede deshacer).</div>
      {borrados.length === 0 && <div style={styles.empty}>La papelera está vacía.</div>}
      <div style={{ display: "grid", gap: 10, marginTop: 8 }}>
        {borrados.map((ev) => (
          <div key={ev.id} style={styles.eventCard}>
            <div style={{ flex: 1 }}>
              <div style={styles.eventName}>{ev.nombre}</div>
              <div style={styles.eventMeta}>{ev.fecha ? fechaLabel(ev.fecha) : "Sin fecha"} · {ev.jornadas.length} {ev.jornadas.length === 1 ? "jornada" : "jornadas"}</div>
            </div>
            <button onClick={() => onRestore(ev.id)} style={styles.smallBtn}>↩ Restaurar</button>
            <button onClick={() => { if (confirm(`¿Eliminar definitivamente "${ev.nombre}"? Esto NO se puede deshacer.`)) onHardDelete(ev.id); }} style={styles.deleteBtn}>Eliminar def.</button>
          </div>
        ))}
      </div>
    </div>
  );
}

function AccessView({ onClose }) {
  const [actual, setActual] = useState("");
  const [nuevoAdmin, setNuevoAdmin] = useState("");
  const [nuevoContador, setNuevoContador] = useState("");
  const [nuevoInventario, setNuevoInventario] = useState("");
  const [msg, setMsg] = useState(null);
  const [busy, setBusy] = useState(false);

  const guardar = async () => {
    if (!actual || !nuevoAdmin || !nuevoContador || !nuevoInventario) {
      setMsg({ ok: false, t: "Rellena los cuatro campos." });
      return;
    }
    setBusy(true); setMsg(null);
    try {
      const ok = await setCodes(actual, nuevoAdmin.trim(), nuevoContador.trim(), nuevoInventario.trim());
      if (ok) {
        setMsg({ ok: true, t: "Códigos actualizados correctamente." });
        setActual(""); setNuevoAdmin(""); setNuevoContador(""); setNuevoInventario("");
      } else {
        setMsg({ ok: false, t: "El código de admin actual no es correcto." });
      }
    } catch (e) {
      setMsg({ ok: false, t: "No se pudo guardar." });
    }
    setBusy(false);
  };

  return (
    <div>
      <button onClick={onClose} style={styles.volverBtn}>‹ Volver</button>
      <div style={{ ...styles.sectionTitle, marginTop: 14 }}>Gestión de acceso</div>
      <div style={styles.dimText}>Cambia los códigos de Admin, Gestor y Contador. Necesitas el código de admin actual para confirmar el cambio.</div>
      <div style={styles.formCard}>
        <label style={styles.fieldLabel}>Código de admin actual</label>
        <input type="password" value={actual} onChange={(e) => setActual(e.target.value)} placeholder="Código actual" style={styles.textInput} />
        <label style={styles.fieldLabel}>Nuevo código de Admin</label>
        <input type="text" value={nuevoAdmin} onChange={(e) => setNuevoAdmin(e.target.value)} placeholder="Nuevo código de admin" style={styles.textInput} />
        <label style={styles.fieldLabel}>Nuevo código de Gestor (puede editar el inicial)</label>
        <input type="text" value={nuevoInventario} onChange={(e) => setNuevoInventario(e.target.value)} placeholder="Nuevo código de gestor" style={styles.textInput} />
        <label style={styles.fieldLabel}>Nuevo código de Contador (solo cuenta el final)</label>
        <input type="text" value={nuevoContador} onChange={(e) => setNuevoContador(e.target.value)} placeholder="Nuevo código de contador" style={styles.textInput} />
        <button onClick={guardar} disabled={busy} style={styles.addBtn}>{busy ? "Guardando…" : "Guardar códigos"}</button>
        {msg && <div style={{ color: msg.ok ? COLORS.green : COLORS.red, fontSize: 13 }}>{msg.t}</div>}
      </div>
    </div>
  );
}

const globalCSS = `
  @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,900&family=Outfit:wght@300;400;500;600&display=swap');
  * { box-sizing: border-box; }
  input[type=number]::-webkit-inner-spin-button { opacity: 0.3; }
  input:focus { outline: none; border-color: ${COLORS.gold} !important; }
  button { cursor: pointer; font-family: 'Outfit', sans-serif; }
  ::placeholder { color: ${COLORS.goldDim}; }
`;

const styles = {
  app: { fontFamily: "'Outfit', sans-serif", background: COLORS.bg, backgroundImage: `radial-gradient(circle at 100% 0%, ${COLORS.panel2} 0%, transparent 45%)`, color: COLORS.cream, minHeight: "100vh", padding: "28px 22px 60px", maxWidth: 880, margin: "0 auto" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "flex-end", borderBottom: `1px solid ${COLORS.line}`, paddingBottom: 18, marginBottom: 22 },
  kicker: { fontSize: 11, letterSpacing: "0.22em", color: COLORS.goldDim, fontWeight: 500, marginBottom: 6 },
  title: { fontFamily: "'Fraunces', serif", fontSize: 34, fontWeight: 900, margin: 0, color: COLORS.cream, letterSpacing: "-0.01em" },
  saveTag: { fontSize: 12, color: COLORS.goldDim, fontWeight: 500, whiteSpace: "nowrap" },
  saveDirtyBtn: { background: COLORS.gold, border: "none", color: COLORS.bg, fontSize: 12, fontWeight: 700, padding: "5px 12px", borderRadius: 14, whiteSpace: "nowrap" },
  headerRight: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" },
  roleTag: { fontSize: 11, fontWeight: 600, color: COLORS.bg, background: COLORS.gold, padding: "3px 10px", borderRadius: 12, letterSpacing: "0.04em", whiteSpace: "nowrap" },
  linkBtn: { background: "transparent", border: `1px solid ${COLORS.line}`, color: COLORS.dim, fontSize: 12, fontWeight: 500, padding: "5px 10px", borderRadius: 14, whiteSpace: "nowrap" },
  fieldLabel: { fontSize: 12, color: COLORS.dim, fontWeight: 500, marginTop: 4 },
  volverBtn: { background: "transparent", border: `1px solid ${COLORS.line}`, color: COLORS.gold, fontSize: 13, fontWeight: 500, padding: "5px 12px", borderRadius: 18, marginBottom: 8 },
  sectionTitle: { fontFamily: "'Fraunces', serif", fontSize: 22, fontWeight: 600, color: COLORS.cream, marginBottom: 16 },
  collapseHeader: { display: "flex", alignItems: "center", justifyContent: "space-between", cursor: "pointer", marginBottom: 16 },
  collapseMeta: { color: COLORS.gold, fontSize: 14, fontWeight: 600 },
  empty: { background: COLORS.panel, border: `1px solid ${COLORS.line}`, borderRadius: 12, padding: 24, color: COLORS.dim, textAlign: "center", fontSize: 14 },
  dimText: { color: COLORS.dim, fontSize: 13, fontStyle: "italic", marginBottom: 8, display: "block" },
  eventCard: { display: "flex", alignItems: "center", gap: 10, background: COLORS.panel, border: `1px solid ${COLORS.line}`, borderRadius: 12, padding: "16px 18px", cursor: "pointer" },
  eventName: { fontFamily: "'Fraunces', serif", fontSize: 18, fontWeight: 600, color: COLORS.cream },
  eventMeta: { fontSize: 12.5, color: COLORS.dim, marginTop: 3 },
  chevron: { color: COLORS.gold, fontSize: 22, fontWeight: 300 },
  eventHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, marginBottom: 20 },
  eventTitle: { fontFamily: "'Fraunces', serif", fontSize: 26, fontWeight: 900, color: COLORS.cream },
  tabs: { display: "flex", gap: 4, marginBottom: 26, flexWrap: "wrap" },
  tab: { background: "transparent", border: "none", color: COLORS.dim, fontSize: 14, fontWeight: 500, padding: "8px 14px", borderRadius: 8 },
  tabActive: { background: COLORS.panel, color: COLORS.gold },
  chipWrap: { display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 18, alignItems: "center" },
  chip: { background: COLORS.panel, border: `1px solid ${COLORS.line}`, color: COLORS.dim, padding: "7px 15px", borderRadius: 20, fontSize: 13, fontWeight: 500 },
  chipActive: { background: COLORS.gold, color: COLORS.bg, borderColor: COLORS.gold },
  chipDone: { borderColor: COLORS.green, color: COLORS.green },
  alertBox: { background: COLORS.amberBg, border: `1px solid ${COLORS.amberLine}`, borderRadius: 10, padding: "12px 14px", marginBottom: 18 },
  alertChip: { background: COLORS.panel, border: `1px solid ${COLORS.amberLine}`, color: COLORS.amber, padding: "7px 13px", borderRadius: 20, fontSize: 13, fontWeight: 600 },
  alertChipActive: { background: COLORS.amber, color: "#ffffff", borderColor: COLORS.amber },
  errorBox: { background: "#fef2f2", border: `1px solid #fecaca`, borderRadius: 10, padding: "12px 14px", marginBottom: 18 },
  errorChip: { background: COLORS.panel, border: `1px solid #fecaca`, color: COLORS.red, padding: "7px 13px", borderRadius: 20, fontSize: 13, fontWeight: 600 },
  errorChipActive: { background: COLORS.red, color: "#ffffff", borderColor: COLORS.red },
  inputError: { borderColor: COLORS.red, background: "#fef2f2" },
  confirmBar: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap", background: COLORS.panel, border: `1px solid ${COLORS.line}`, borderRadius: 10, padding: "10px 14px", marginBottom: 18 },
  chipEdit: { background: COLORS.panel, border: `1px solid ${COLORS.line}`, color: COLORS.cream, padding: "6px 8px 6px 14px", borderRadius: 20, fontSize: 13, display: "inline-flex", alignItems: "center", gap: 4 },
  jornadaRow: { display: "flex", alignItems: "center", gap: 10, background: COLORS.panel, border: `1px solid ${COLORS.line}`, borderRadius: 10, padding: "12px 14px", cursor: "pointer" },
  jornadaRowActive: { borderColor: COLORS.gold, background: COLORS.panel2 },
  summaryBar: { background: COLORS.panel, border: `1px solid ${COLORS.line}`, borderRadius: 12, padding: "16px 20px", marginBottom: 20, display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap" },
  bigNum: { fontFamily: "'Fraunces', serif", fontSize: 30, fontWeight: 900, color: COLORS.gold, marginLeft: 4, marginRight: 4 },
  catTitle: { fontFamily: "'Fraunces', serif", fontSize: 18, fontWeight: 600, color: COLORS.gold, marginBottom: 12, paddingBottom: 6, borderBottom: `1px solid ${COLORS.line}` },
  tableHead: { display: "flex", alignItems: "center", fontSize: 11, letterSpacing: "0.08em", color: COLORS.dim, textTransform: "uppercase", padding: "0 4px 8px" },
  colNum: { width: 78, textAlign: "center" },
  row: { display: "flex", alignItems: "center", padding: "8px 4px", borderBottom: `1px solid ${COLORS.panel2}` },
  unidad: { color: COLORS.dim, fontSize: 12 },
  input: { width: 78, textAlign: "center", background: COLORS.panel2, border: `1px solid ${COLORS.line}`, borderRadius: 7, color: COLORS.cream, fontSize: 14, padding: "7px 4px", fontFamily: "'Outfit', sans-serif" },
  inputDisabled: { opacity: 0.5, cursor: "not-allowed" },
  movRow: { display: "flex", alignItems: "center", gap: 8, padding: "8px 0", borderBottom: `1px solid ${COLORS.panel2}` },
  movHint: { fontSize: 11, color: COLORS.gold, marginTop: 2, fontWeight: 600 },
  toggleBtn: { border: "1px solid", fontSize: 11, fontWeight: 600, padding: "4px 9px", borderRadius: 14, whiteSpace: "nowrap", background: "transparent" },
  toggleOpen: { borderColor: COLORS.green, color: COLORS.green },
  toggleLocked: { borderColor: COLORS.goldDim, color: COLORS.dim },
  barTrack: { height: 10, background: COLORS.panel2, borderRadius: 6, overflow: "hidden" },
  barFill: { height: "100%", background: `linear-gradient(90deg, ${COLORS.goldDim}, ${COLORS.gold})`, borderRadius: 6, transition: "width .4s" },
  matrix: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  th: { padding: "10px 8px", textAlign: "center", color: COLORS.dim, fontWeight: 500, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", borderBottom: `1px solid ${COLORS.line}` },
  td: { padding: "9px 8px", textAlign: "center", borderBottom: `1px solid ${COLORS.panel2}` },
  listRow: { display: "flex", alignItems: "center", gap: 10, padding: "10px 4px", borderBottom: `1px solid ${COLORS.panel2}` },
  metaTag: { color: COLORS.dim, fontSize: 12.5 },
  formRow: { display: "flex", gap: 8, flexWrap: "wrap" },
  formCard: { background: COLORS.panel, border: `1px solid ${COLORS.line}`, borderRadius: 12, padding: 16, marginTop: 14, marginBottom: 14, display: "flex", flexDirection: "column", gap: 10 },
  formCardTitle: { fontFamily: "'Fraunces', serif", fontSize: 16, fontWeight: 600, color: COLORS.gold },
  textInput: { flex: 1, minWidth: 120, background: COLORS.panel2, border: `1px solid ${COLORS.line}`, borderRadius: 7, color: COLORS.cream, fontSize: 14, padding: "9px 12px", fontFamily: "'Outfit', sans-serif" },
  select: { width: "100%", background: COLORS.panel2, border: `1px solid ${COLORS.line}`, borderRadius: 8, color: COLORS.cream, fontSize: 15, padding: "11px 12px", fontFamily: "'Outfit', sans-serif" },
  totalBar: { background: COLORS.panel, border: `1px solid ${COLORS.line}`, borderRadius: 12, padding: "14px 18px", marginBottom: 22, display: "flex", flexWrap: "wrap", gap: 22, alignItems: "center" },
  totalItem: { display: "flex", flexDirection: "column", gap: 2 },
  totalLabel: { fontSize: 11, letterSpacing: "0.08em", textTransform: "uppercase", color: COLORS.dim },
  totalNum: { fontFamily: "'Fraunces', serif", fontSize: 24, fontWeight: 900, color: COLORS.cream },
  addBtn: { background: COLORS.gold, border: "none", color: COLORS.bg, fontWeight: 600, fontSize: 14, padding: "9px 18px", borderRadius: 7, whiteSpace: "nowrap" },
  smallBtn: { background: COLORS.panel, border: `1px solid ${COLORS.line}`, color: COLORS.cream, fontSize: 13, padding: "7px 14px", borderRadius: 7 },
  exportBtn: { background: "transparent", border: `1px solid ${COLORS.gold}`, color: COLORS.gold, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 8, whiteSpace: "nowrap" },
  importBtn: { background: "transparent", border: `1px solid ${COLORS.goldDim}`, color: COLORS.cream, fontSize: 14, fontWeight: 600, padding: "9px 18px", borderRadius: 7, whiteSpace: "nowrap" },
  editBtn: { background: "transparent", border: `1px solid ${COLORS.line}`, color: COLORS.gold, fontSize: 14, lineHeight: 1, padding: "6px 9px", borderRadius: 8 },
  deleteBtn: { background: "transparent", border: `1px solid ${COLORS.red}`, color: COLORS.red, fontSize: 13, fontWeight: 600, padding: "8px 14px", borderRadius: 7, whiteSpace: "nowrap" },
  continueBtn: { background: COLORS.gold, border: "none", color: COLORS.bg, fontWeight: 700, fontSize: 15, padding: "14px 18px", borderRadius: 9, width: "100%", marginTop: 22 },
  xBtn: { background: "transparent", border: "none", color: COLORS.red, fontSize: 22, lineHeight: 1, padding: "0 6px" },
  xBtnSm: { background: "transparent", border: "none", color: COLORS.red, fontSize: 18, lineHeight: 1, padding: "0 4px" },
};
