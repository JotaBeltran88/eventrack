import React, { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { loadState, saveState, subscribeState, login, setCodes } from "./storage";

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

const COLORS = {
  bg: "#0e0b08", panel: "#1a1410", panel2: "#221a13", line: "#3a2c1f",
  gold: "#d9a441", goldDim: "#8a6a2c", cream: "#f2e8d8", dim: "#9b8a73",
  green: "#7fae6b", red: "#cc6a55",
};

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

  // Guarda la última versión sincronizada para no reescribir/recargar en bucle.
  const lastSynced = useRef("");

  // Carga inicial desde Supabase + suscripción en tiempo real.
  useEffect(() => {
    (async () => {
      try {
        const data = await loadState();
        if (data && Array.isArray(data.eventos)) {
          lastSynced.current = JSON.stringify(data);
          setEventos(data.eventos);
        }
      } catch (e) { /* primera vez o sin conexión */ }
      setLoaded(true);
    })();

    const unsub = subscribeState((data) => {
      if (!data || !Array.isArray(data.eventos)) return;
      const json = JSON.stringify(data);
      if (json === lastSynced.current) return; // es nuestro propio cambio, lo ignoramos
      lastSynced.current = json;
      setEventos(data.eventos);
    });
    return unsub;
  }, []);

  const save = useCallback(async (next) => {
    setSaving(true);
    try { await saveState(next); }
    catch (e) { console.error("Error guardando", e); }
    setSaving(false);
  }, []);

  // Guardado manual inmediato (botón "Guardar").
  const guardarAhora = useCallback(() => save({ eventos }), [save, eventos]);

  // Guardado automático (debounced) cuando cambian los datos localmente.
  useEffect(() => {
    if (!loaded) return;
    const payload = { eventos };
    const json = JSON.stringify(payload);
    if (json === lastSynced.current) return; // nada nuevo que guardar
    const t = setTimeout(() => {
      lastSynced.current = json;
      save(payload);
    }, 600);
    return () => clearTimeout(t);
  }, [eventos, loaded, save]);

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
          <div style={{ cursor: "pointer" }} onClick={() => { setEventoActivoId(null); setShowAcceso(false); setShowPapelera(false); }}>
            <div style={styles.kicker}>J.B.APP</div>
            <h1 style={styles.title}>Event<span style={{ color: COLORS.gold }}>rack</span></h1>
          </div>
        </div>
        <div style={styles.headerRight}>
          <span style={styles.saveTag}>{saving ? "Guardando…" : "Guardado ✓"}</span>
          <span style={styles.roleTag}>{role === "admin" ? "Admin" : "Contador"}</span>
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

  return (
    <div>
      {esAdmin && (
        <div style={styles.formCard}>
          <div style={styles.formCardTitle}>Nuevo evento</div>
          <input value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Nombre del evento (ej. Noches del Botánico)" style={styles.textInput} />
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
      )}

      <div style={{ ...styles.sectionTitle, marginTop: esAdmin ? 30 : 0 }}>Eventos en curso</div>
      {enCurso.length === 0 && <div style={styles.empty}>{esAdmin ? "No hay eventos en curso. Crea uno arriba." : "No hay eventos en curso."}</div>}
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
  const [jornadaActivaId, setJornadaActivaId] = useState(evento.jornadas[0]?.id || null);
  const [ubicActiva, setUbicActiva] = useState(evento.ubicaciones[0] || "");

  const upd = (updater) => updateEvento(evento.id, updater);

  useEffect(() => {
    if (!evento.ubicaciones.includes(ubicActiva)) setUbicActiva(evento.ubicaciones[0] || "");
  }, [evento.ubicaciones, ubicActiva]);

  useEffect(() => {
    if (jornadaActivaId && !evento.jornadas.find((j) => j.id === jornadaActivaId)) {
      setJornadaActivaId(evento.jornadas[0]?.id || null);
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
        {[["conteo", "Conteo"], ["resumen", "Resumen"]].map(([k, label]) => (
          <button key={k} onClick={() => setTab(k)} style={{ ...styles.tab, ...(tab === k ? styles.tabActive : {}) }}>{label}</button>
        ))}
        {esAdmin && (
          <button onClick={() => setModo("config")} style={styles.tab}>⚙ Configuración</button>
        )}
      </nav>

      {tab === "conteo" && (
        <ConteoView evento={evento} upd={upd} jornada={jornadaActiva} jornadaActivaId={jornadaActivaId} setJornadaActivaId={setJornadaActivaId} ubicActiva={ubicActiva} setUbicActiva={setUbicActiva} onGuardar={onGuardar} />
      )}
      {tab === "resumen" && <ResumenView evento={evento} />}
    </div>
  );
}

function emptyJornada(fecha, evento) {
  const conteo = {};
  evento.ubicaciones.forEach((u) => {
    conteo[u] = Object.fromEntries(evento.productos.map((p) => [p.id, { inicial: 0, final: 0 }]));
  });
  return { id: "j" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), fecha, conteo };
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

  return (
    <div>
      <div style={{ ...styles.sectionTitle, marginTop: 34 }}>Jornadas {esSingle ? "(día único)" : "del evento"}</div>
      {evento.jornadas.length === 0 && <div style={styles.empty}>{esSingle ? "Añade la fecha del día abajo." : "Sin jornadas. Añade noches sueltas o genera un rango de fechas abajo."}</div>}

      <div style={{ display: "grid", gap: 8, marginBottom: 24 }}>
        {evento.jornadas.map((j) => (
          <div key={j.id} style={{ ...styles.jornadaRow, ...(j.id === jornadaActivaId ? styles.jornadaRowActive : {}) }} onClick={() => setJornadaActivaId(j.id)}>
            <span style={{ flex: 1, color: COLORS.cream, fontWeight: 500 }}>{fechaLabel(j.fecha)}</span>
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

        const jornadas = ev.jornadas.map((j) => {
          const conteo = { ...j.conteo };
          ubicNuevas.forEach((u) => {
            conteo[u] = Object.fromEntries(productos.map((p) => [p.id, { inicial: 0, final: 0 }]));
          });
          ev.ubicaciones.forEach((u) => {
            const base = { ...(conteo[u] || {}) };
            prodNuevos.forEach((p) => { base[p.id] = base[p.id] || { inicial: 0, final: 0 }; });
            conteo[u] = base;
          });
          return { ...j, conteo };
        });

        return { ...ev, ubicaciones, productos, jornadas };
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
    upd((ev) => ({
      ...ev,
      ubicaciones: [...ev.ubicaciones, n],
      jornadas: ev.jornadas.map((j) => ({
        ...j,
        conteo: { ...j.conteo, [n]: Object.fromEntries(ev.productos.map((p) => [p.id, { inicial: 0, final: 0 }])) },
      })),
    }));
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
    upd((ev) => ({
      ...ev,
      productos: [...ev.productos, prod],
      jornadas: ev.jornadas.map((j) => {
        const c = { ...j.conteo };
        ev.ubicaciones.forEach((u) => { c[u] = { ...c[u], [id]: { inicial: 0, final: 0 } }; });
        return { ...j, conteo: c };
      }),
    }));
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
          {evento.jornadas.map((j) => (
            <option key={j.id} value={j.id}>{fechaLabel(j.fecha)}</option>
          ))}
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

function ConteoView({ evento, upd, jornada, jornadaActivaId, setJornadaActivaId, ubicActiva, setUbicActiva, onGuardar }) {
  const [catActiva, setCatActiva] = useState("");
  const [guardado, setGuardado] = useState(false);
  if (evento.ubicaciones.length === 0 || evento.productos.length === 0)
    return <div style={styles.empty}>Necesitas ubicaciones y referencias (Configuración) para el conteo.</div>;
  if (evento.jornadas.length === 0)
    return <div style={styles.empty}>Crea al menos una jornada para registrar el conteo.</div>;
  if (!jornada) return (<div><JornadaSelector evento={evento} jornadaActivaId={jornadaActivaId} setJornadaActivaId={setJornadaActivaId} /><div style={styles.empty}>Selecciona una jornada.</div></div>);

  const getCell = (ubic, pid) => jornada.conteo?.[ubic]?.[pid] || { inicial: 0, final: 0 };
  const setValor = (pid, campo, valor) => {
    const v = valor === "" ? 0 : Math.max(0, Number(valor));
    upd((ev) => ({
      ...ev,
      jornadas: ev.jornadas.map((j) => {
        if (j.id !== jornada.id) return j;
        const conteo = { ...j.conteo };
        const ubic = { ...(conteo[ubicActiva] || {}) };
        ubic[pid] = { ...(ubic[pid] || { inicial: 0, final: 0 }), [campo]: v };
        conteo[ubicActiva] = ubic;
        return { ...j, conteo };
      }),
    }));
  };

  const categorias = [...new Set(evento.productos.map((p) => p.categoria))];
  const catSel = categorias.includes(catActiva) ? catActiva : "";

  const guardar = async () => {
    if (onGuardar) await onGuardar();
    setGuardado(true);
    setTimeout(() => setGuardado(false), 2500);
  };

  return (
    <div>
      <JornadaSelector evento={evento} jornadaActivaId={jornadaActivaId} setJornadaActivaId={setJornadaActivaId} />
      <div style={styles.chipWrap}>
        {evento.ubicaciones.map((u) => (
          <button key={u} onClick={() => setUbicActiva(u)} style={{ ...styles.chip, ...(u === ubicActiva ? styles.chipActive : {}) }}>{u}</button>
        ))}
      </div>

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
        <div style={styles.empty}>Elige una familia de productos para ver sus referencias.</div>
      ) : (
        <div style={{ marginBottom: 26 }}>
          <div style={styles.tableHead}>
            <span style={{ flex: 2 }}>Producto</span>
            <span style={styles.colNum}>Inicial</span><span style={styles.colNum}>Final</span><span style={styles.colNum}>Consumo</span>
          </div>
          {evento.productos.filter((p) => p.categoria === catSel).map((p) => {
            const c = getCell(ubicActiva, p.id);
            const consumo = Math.max(0, c.inicial - c.final);
            return (
              <div key={p.id} style={styles.row}>
                <span style={{ flex: 2 }}><span style={{ color: COLORS.cream }}>{p.nombre}</span><span style={styles.unidad}> · {p.unidad}</span></span>
                <input type="number" min="0" value={c.inicial || ""} placeholder="0" onChange={(e) => setValor(p.id, "inicial", e.target.value)} style={styles.input} />
                <input type="number" min="0" value={c.final || ""} placeholder="0" onChange={(e) => setValor(p.id, "final", e.target.value)} style={styles.input} />
                <span style={{ ...styles.colNum, color: consumo > 0 ? COLORS.gold : COLORS.dim, fontWeight: 600 }}>{consumo}</span>
              </div>
            );
          })}
        </div>
      )}

      <button onClick={guardar} style={styles.continueBtn}>Guardar</button>
      {guardado && <div style={{ color: COLORS.green, fontSize: 13, textAlign: "center", marginTop: 8 }}>Guardado ✓</div>}
    </div>
  );
}

function ResumenView({ evento }) {
  if (evento.jornadas.length === 0) return <div style={styles.empty}>Sin jornadas todavía.</div>;
  if (evento.productos.length === 0) return <div style={styles.empty}>Sin referencias todavía.</div>;

  // Totales por referencia, sumando todas las jornadas y ubicaciones.
  const totalRef = (pid) => {
    let ini = 0, fin = 0, con = 0;
    evento.jornadas.forEach((j) => evento.ubicaciones.forEach((u) => {
      const c = j.conteo?.[u]?.[pid] || { inicial: 0, final: 0 };
      ini += c.inicial || 0; fin += c.final || 0; con += Math.max(0, (c.inicial || 0) - (c.final || 0));
    }));
    return { ini, fin, con };
  };

  return (
    <div>
      <div style={styles.dimText}>Total por referencia, sumando todas las jornadas y ubicaciones del evento. Pulsa «Exportar Excel» (arriba) para descargar y tratar estos datos.</div>
      <div style={{ overflowX: "auto" }}>
        <table style={styles.matrix}>
          <thead><tr>
            <th style={{ ...styles.th, textAlign: "left" }}>Producto</th>
            <th style={styles.th}>Unidad</th>
            <th style={styles.th}>Inicial</th>
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

function ExportButton({ evento }) {
  const exportar = () => {
    const wb = XLSX.utils.book_new();
    const cons = (j, u, pid) => { const c = j.conteo?.[u]?.[pid] || { inicial: 0, final: 0 }; return Math.max(0, c.inicial - c.final); };

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
      const c = j.conteo?.[u]?.[p.id] || { inicial: 0, final: 0 };
      detalle.push({ Jornada: j.fecha, "Ubicación": u, "Categoría": p.categoria, Producto: p.nombre, Unidad: p.unidad, Inicial: c.inicial, Final: c.final, Consumo: Math.max(0, c.inicial - c.final) });
    })));
    if (detalle.length) {
      const ws = XLSX.utils.json_to_sheet(detalle);
      ws["!cols"] = [{ wch: 12 }, { wch: 18 }, { wch: 14 }, { wch: 26 }, { wch: 10 }, { wch: 9 }, { wch: 9 }, { wch: 10 }];
      XLSX.utils.book_append_sheet(wb, ws, "Detalle inventario");
    }

    if (evento.productos.length) {
      const acc = evento.productos.map((p) => {
        let ini = 0, fin = 0, con = 0;
        evento.jornadas.forEach((j) => evento.ubicaciones.forEach((u) => {
          const c = j.conteo?.[u]?.[p.id] || { inicial: 0, final: 0 };
          ini += c.inicial || 0; fin += c.final || 0; con += Math.max(0, (c.inicial || 0) - (c.final || 0));
        }));
        return { Producto: p.nombre, "Categoría": p.categoria, Unidad: p.unidad, "Inicial total": ini, "Final total": fin, "Consumo total": con };
      });
      const ws = XLSX.utils.json_to_sheet(acc);
      ws["!cols"] = [{ wch: 26 }, { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 14 }];
      XLSX.utils.book_append_sheet(wb, ws, "Total por referencia");
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
      if (r === "admin" || r === "contador") { onLogin(r); return; }
      setErr("Código incorrecto. Inténtalo de nuevo.");
    } catch (e) {
      setErr("No se pudo conectar. Revisa tu conexión.");
    }
    setBusy(false);
  };

  return (
    <div style={{ ...styles.app, display: "flex", flexDirection: "column", justifyContent: "center", alignItems: "center", minHeight: "100vh" }}>
      <style>{globalCSS}</style>
      <div style={styles.kicker}>J.B.APP</div>
      <h1 style={{ ...styles.title, fontSize: 40, marginBottom: 6 }}>Event<span style={{ color: COLORS.gold }}>rack</span></h1>
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
  const [msg, setMsg] = useState(null);
  const [busy, setBusy] = useState(false);

  const guardar = async () => {
    if (!actual || !nuevoAdmin || !nuevoContador) {
      setMsg({ ok: false, t: "Rellena los tres campos." });
      return;
    }
    setBusy(true); setMsg(null);
    try {
      const ok = await setCodes(actual, nuevoAdmin.trim(), nuevoContador.trim());
      if (ok) {
        setMsg({ ok: true, t: "Códigos actualizados correctamente." });
        setActual(""); setNuevoAdmin(""); setNuevoContador("");
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
      <div style={styles.dimText}>Cambia los códigos de Admin y Contador. Necesitas el código de admin actual para confirmar el cambio.</div>
      <div style={styles.formCard}>
        <label style={styles.fieldLabel}>Código de admin actual</label>
        <input type="password" value={actual} onChange={(e) => setActual(e.target.value)} placeholder="Código actual" style={styles.textInput} />
        <label style={styles.fieldLabel}>Nuevo código de Admin</label>
        <input type="text" value={nuevoAdmin} onChange={(e) => setNuevoAdmin(e.target.value)} placeholder="Nuevo código de admin" style={styles.textInput} />
        <label style={styles.fieldLabel}>Nuevo código de Contador</label>
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
