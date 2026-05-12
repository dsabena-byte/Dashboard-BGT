# Deploy en Vercel

El dashboard es estatico (HTML + JSON), no requiere build step.

## Importar el repo (una sola vez)

1. https://vercel.com/new
2. **Import Git Repository** -> elegi `dsabena-byte/Dashboard-BGT`.
3. **Configure Project**:
   - **Framework Preset**: `Other` (es estatico puro).
   - **Build Command**: vacio.
   - **Output Directory**: vacio (default).
   - **Install Command**: vacio.
   - **Root Directory**: vacio.
4. **Deploy**. En 10-15 segundos esta arriba en `https://<nombre>.vercel.app`.

## Como se actualiza

La GitHub Action que sincroniza la planilla con SharePoint commitea
`data.json` a `main` cada 6 horas (o cuando lo corres manualmente). Cada commit
en `main` dispara un nuevo deploy en Vercel **automaticamente** — no hay que
tocar nada.

Tiempo total desde que actualizas la planilla en SharePoint hasta verlo en el
dashboard: maximo 6 horas + ~15 segundos del deploy.

## Cache

El `vercel.json` configura `Cache-Control: max-age=0, must-revalidate` para
`/data.json` e `/index.html`. Esto asegura que los usuarios siempre vean la
ultima version sincronizada (sin esperar a que expire el CDN).

## Forzar actualizacion sin esperar el cron

1. **Actions** -> **Sync presupuestos desde SharePoint** -> **Run workflow**.
2. Si hay cambios, se commitea `data.json` y Vercel redespliega solo.
