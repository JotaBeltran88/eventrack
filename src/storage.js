// Capa de almacenamiento compartido en la nube (Supabase).
// Reemplaza al window.storage del entorno de Claude.
// Toda la app guarda un único "estado" (la lista de eventos) en una fila.
import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !anonKey) {
  // Aviso claro en consola si faltan las variables de entorno al desplegar.
  console.error(
    "Faltan VITE_SUPABASE_URL o VITE_SUPABASE_ANON_KEY. " +
      "Configúralas en Vercel (Project Settings → Environment Variables)."
  );
}

export const supabase = createClient(url || "", anonKey || "");

// Identificador de la fila única donde vive todo el estado de la app.
const ROW_ID = "eventrack-v3";
const TABLE = "eventrack_state";

// Lee el estado guardado. Devuelve el objeto { eventos: [...] } o null si no hay nada.
export async function loadState() {
  const { data, error } = await supabase
    .from(TABLE)
    .select("data")
    .eq("id", ROW_ID)
    .maybeSingle();
  if (error) {
    console.error("Error cargando estado:", error.message);
    return null;
  }
  return data?.data ?? null;
}

// Guarda (upsert) el estado completo.
export async function saveState(obj) {
  const { error } = await supabase
    .from(TABLE)
    .upsert({ id: ROW_ID, data: obj, updated_at: new Date().toISOString() });
  if (error) console.error("Error guardando estado:", error.message);
}

// Lee la fila con su marca de versión (updated_at), para control de concurrencia.
export async function loadStateRow() {
  const { data, error } = await supabase
    .from(TABLE)
    .select("data, updated_at")
    .eq("id", ROW_ID)
    .maybeSingle();
  if (error) { console.error("Error cargando fila:", error.message); return null; }
  return data; // { data, updated_at } | null
}

// Guarda fusionando con lo último del servidor, con control de versión (CAS).
// mergeFn(serverData) debe devolver el estado fusionado a escribir.
// Devuelve el estado fusionado guardado, o null si no se pudo.
export async function saveWithMerge(mergeFn, maxRetries = 5) {
  for (let i = 0; i < maxRetries; i++) {
    const row = await loadStateRow();
    const server = row && row.data ? row.data : { eventos: [] };
    const merged = mergeFn(server);
    const nowIso = new Date().toISOString();
    if (!row) {
      const { error } = await supabase.from(TABLE).insert({ id: ROW_ID, data: merged, updated_at: nowIso });
      if (!error) return merged;
      continue; // alguien insertó a la vez → reintentar como update
    }
    const { data: upd, error } = await supabase
      .from(TABLE)
      .update({ data: merged, updated_at: nowIso })
      .eq("id", ROW_ID)
      .eq("updated_at", row.updated_at)
      .select("id");
    if (error) { console.error("Error guardando estado:", error.message); return null; }
    if (upd && upd.length > 0) return merged; // nadie escribió en medio → éxito
    // updated_at cambió (otro guardó) → reintentar con el estado más nuevo
  }
  console.warn("saveWithMerge: demasiados reintentos de concurrencia");
  return null;
}

// Se suscribe a cambios en tiempo real. Llama a cb(data) cuando otro
// dispositivo guarda. Devuelve una función para cancelar la suscripción.
export function subscribeState(cb) {
  const channel = supabase
    .channel("eventrack_state_changes")
    .on(
      "postgres_changes",
      { event: "*", schema: "public", table: TABLE, filter: `id=eq.${ROW_ID}` },
      (payload) => {
        if (payload.new && payload.new.data) cb(payload.new.data);
      }
    )
    .subscribe();
  return () => supabase.removeChannel(channel);
}

// ── Acceso por código (verificación en el servidor) ──

// Comprueba un código. Devuelve "admin", "contador" o "" si no coincide.
export async function login(code) {
  const { data, error } = await supabase.rpc("eventrack_login", { code });
  if (error) {
    console.error("Error de login:", error.message);
    return "";
  }
  return data || "";
}

// Cambia los códigos. Requiere el código de admin actual. Devuelve true si funcionó.
export async function setCodes(currentAdmin, newAdmin, newContador, newInventario) {
  const { data, error } = await supabase.rpc("eventrack_set_codes", {
    current_admin: currentAdmin,
    new_admin: newAdmin,
    new_contador: newContador,
    new_inventario: newInventario,
  });
  if (error) {
    console.error("Error cambiando códigos:", error.message);
    return false;
  }
  return data === true;
}
