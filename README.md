# Eventrack

App de inventario y control de stock para eventos (multi-evento, con jornadas por día,
conteo inicial/final → consumo, resumen acumulado e importación/exportación Excel).

Hecha con **Vite + React** y datos compartidos en la nube con **Supabase**.

## Puesta en marcha en local

```bash
npm install
cp .env.example .env.local   # y rellena tus claves de Supabase
npm run dev
```

## Despliegue

Ver la guía paso a paso en **[DEPLOY.md](DEPLOY.md)**: crear la base de datos en
Supabase, subir a GitHub y desplegar en Vercel.

## Estructura

- `src/App.jsx` — la app completa (UI + lógica).
- `src/storage.js` — capa de guardado contra Supabase (carga, guardado y tiempo real).
- `supabase.sql` — SQL para crear la tabla y las políticas en Supabase.
