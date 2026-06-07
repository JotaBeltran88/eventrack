# Guía de despliegue de Eventrack

Objetivo: tener Eventrack en una URL propia (tipo `eventrack.vercel.app`) que puedas
abrir en el iPhone y añadir a la pantalla de inicio, con datos compartidos entre móviles.

Son **3 servicios gratuitos**: Supabase (base de datos), GitHub (código) y Vercel (web).
El código y GitHub ya te los deja montados Claude Code; tú haces Supabase y Vercel.

---

## Parte A · Supabase (base de datos compartida)

1. Entra en **https://supabase.com** → *Start your project* → inicia sesión (con GitHub es lo más rápido).
2. **New project**. Ponle nombre (`eventrack`), una contraseña de base de datos (guárdala) y región Europe (West). Crear. Espera ~2 min a que esté listo.
3. En la barra lateral abre **SQL Editor** → *New query*.
4. Abre el archivo **`supabase.sql`** de este proyecto, copia TODO su contenido, pégalo y pulsa **Run**. Debe decir *Success*.
5. Ahora copia tus **dos claves de conexión**. Esta es la parte que cambió de sitio:
   - Arriba a la derecha del proyecto pulsa el botón **Connect**.
   - En el desplegable de **Framework** elige **Vite** (si no está, no pasa nada).
   - Baja hasta el recuadro de variables de entorno. Verás dos líneas:
     ```
     VITE_SUPABASE_URL=https://xxxxxxxx.supabase.co
     VITE_SUPABASE_ANON_KEY=eyJhbGc...   (una cadena larguísima)
     ```
   - Copia esos **dos valores** (lo de la derecha del `=`). Los necesitarás en Vercel.
   - Alternativa si no ves Connect: barra lateral → **Settings** → la URL está en
     *Data API* y la clave pública en *API Keys* (vale tanto `anon`/`public` como la
     nueva `sb_publishable_...`). **Nunca uses la `service_role`/`secret`.**

---

## Parte B · GitHub (código)

Claude Code ya creó el repositorio y subió el proyecto. La URL es la que te indicó
en el chat (algo como `https://github.com/JotaBeltran88/eventrack`).
Si tuvieras que volver a subirlo a mano, sería: crear repo vacío en GitHub y
`git push`.

---

## Parte C · Vercel (publicar la web)

1. Entra en **https://vercel.com** → *Sign up* / *Log in* **con GitHub**.
2. **Add New… → Project**. Te listará tus repos de GitHub. Elige **eventrack** → *Import*.
3. Vercel detecta solo que es **Vite**. No cambies nada del build.
4. Despliega las **Environment Variables** (importante) y añade las dos:
   | Name | Value |
   |------|-------|
   | `VITE_SUPABASE_URL` | el valor que acaba en `.supabase.co` |
   | `VITE_SUPABASE_ANON_KEY` | la cadena larga (`eyJ...` o `sb_publishable_...`) |
5. Pulsa **Deploy**. En ~1 minuto tendrás una URL tipo `https://eventrack-xxxx.vercel.app`.

> Si el build fallara, Vercel muestra el error en rojo: cópialo y pásamelo, lo arreglamos.

---

## Parte D · Añadir al iPhone como app

1. Abre la URL de Vercel en **Safari** (en iOS tiene que ser Safari).
2. Toca el botón **Compartir** (cuadrado con flecha hacia arriba).
3. **Añadir a pantalla de inicio** → nombre *Eventrack* → *Añadir*.
4. Aparece el icono y se abre a pantalla completa, como una app.

---

## Notas

- **Datos compartidos:** todos los que abran el enlace ven y editan el mismo
  inventario, en tiempo real. La regla es "última escritura gana": que cada
  ubicación/jornada la lleve una sola persona a la vez.
- **Actualizar la app:** cuando cambiemos el código y se suba a GitHub, Vercel
  vuelve a desplegar solo. El enlace y el icono del iPhone ya muestran la versión nueva.
- **Copia de seguridad:** al cerrar cada jornada, usa *Exportar Excel* para tener
  una copia en tu mano que no dependa de la nube.
