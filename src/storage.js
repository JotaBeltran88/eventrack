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
