# Explorer Fleet Control

App de abastecimiento y requisiciones para la flota (Humboldt Explorer, Tiburón Explorer, Grand Majestic).

Migrada desde un artifact de Claude a una app independiente:
- **Datos**: se guardan en Supabase (antes usaban `window.storage`, exclusivo de Claude).
- **Sugerencia de Claude / extracción de documentos**: pasa por una función serverless
  (`/api/claude.js`) que guarda la clave API de Anthropic en el servidor, nunca en el navegador.

## Desplegar en Vercel

1. Sube esta carpeta a un repositorio de GitHub.
2. En vercel.com -> "Add New Project" -> importa el repositorio. Vercel detecta Vite automáticamente.
3. Antes de darle "Deploy", ve a **Environment Variables** y agrega:

   | Nombre | Valor |
   |---|---|
   | `VITE_SUPABASE_URL` | la URL de tu proyecto Supabase |
   | `VITE_SUPABASE_ANON_KEY` | la clave publicable ("anon key") de Supabase |
   | `ANTHROPIC_API_KEY` | tu clave de console.anthropic.com (nunca la de VITE_, esta es solo de servidor) |

4. Click "Deploy". En 1-2 minutos tendrás tu URL pública (`algo.vercel.app`).

## Base de datos

Antes de usar la app, corre el contenido de `supabase_schema.sql` en:
Supabase Dashboard -> tu proyecto -> SQL Editor -> New query -> pega todo -> Run.

Esto crea las tablas `hx_requisiciones`, `hx_facturas`, `hx_config` con seguridad a nivel de fila (RLS).

## Desarrollo local

```bash
npm install
cp .env.example .env.local   # y completa con tus valores reales
npm run dev
```

Nota: en local, `/api/claude.js` no corre a menos que uses `vercel dev` (la CLI de Vercel) en vez de `npm run dev`, porque las funciones serverless son propias de Vercel.
