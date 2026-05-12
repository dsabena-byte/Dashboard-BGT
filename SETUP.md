# Integracion SharePoint -> Dashboard

El dashboard (`budget_dashboard.html`) lee los datos de `data.json`. Ese archivo
se actualiza automaticamente desde una planilla de Excel publicada en
SharePoint mediante una GitHub Action programada.

## Arquitectura

```
SharePoint (Excel)
      |
      | "anyone with the link" (lectura)
      v
GitHub Action (cron cada 6h)
      |
      | scripts/sync_sharepoint.py
      v
data.json (commit + push)
      |
      v
budget_dashboard.html (fetch en runtime)
```

No requiere Azure AD, App Registration ni intervencion de IT.

---

## Configuracion paso a paso

### 1. Compartir la planilla desde SharePoint

1. Abri la planilla `.xlsx` en SharePoint Online.
2. Click en **Share** (Compartir) arriba a la derecha.
3. En "Link settings" (Configuracion del vinculo) eligi:
   - **Anyone with the link** (Cualquier persona con el vinculo)
   - Permiso: **Can view** (Puede ver)
4. Copia el link generado. Va a verse como:

   ```
   https://<tenant>.sharepoint.com/:x:/s/<site>/Ed_xxxxxxxxxxx?e=AbCdEf
   ```

> **Si tu tenant bloquea "anyone with the link"** (lo gestionan los admins de
> IT), esta opcion no va a funcionar. En ese caso hay que recurrir al fallback
> con script local (ver seccion "Fallback").

### 2. Cargar el secret en GitHub

En el repo de GitHub:

1. **Settings** -> **Secrets and variables** -> **Actions** -> **New repository secret**.
2. Nombre: `SHAREPOINT_URL`
3. Valor: el link completo del paso anterior.
4. Guardar.

(Opcional) Si la planilla tiene varias hojas y queres leer una en particular:

1. **Settings** -> **Secrets and variables** -> **Actions** -> tab **Variables** -> **New repository variable**.
2. Nombre: `SHAREPOINT_SHEET`
3. Valor: el nombre exacto de la hoja (ej. `Datos`).

Si no se setea, el script lee la primera hoja del workbook.

### 3. Formato esperado de la planilla

La primera fila debe ser encabezados. El script reconoce estas columnas
(case-insensitive, acepta sinonimos):

| Columna requerida | Alias aceptados                              |
|-------------------|----------------------------------------------|
| `presupuesto`     | `tipo`, `tipo de presupuesto`                |
| `cuenta`          | `categoria`, `categoria contable`            |
| `anio`            | `ano`, `año`, `year`                         |
| `mes`             | `month` (valores ENERO..DICIEMBRE en mayus)  |
| `concepto`        | `subcuenta`, `detalle`                       |
| `ars`             | `importe ars`, `monto ars`, `pesos`          |
| `usd`             | `importe usd`, `monto usd`, `dolares`        |

Filas con `mes` invalido o `presupuesto` vacio se descartan.

### 4. Probar manualmente

En GitHub: **Actions** -> **Sync presupuestos desde SharePoint** ->
**Run workflow**. Si todo esta bien:
- En el job vas a ver: `Parseadas N filas validas.`
- Si hubo cambios respecto al `data.json` actual, se commitea automaticamente.

### 5. Schedule

Por defecto corre cada 6 horas (`cron: "0 */6 * * *"` UTC).
Para cambiar la frecuencia, editar `.github/workflows/sync-sharepoint.yml`.

---

## Probar localmente

```bash
pip install -r scripts/requirements.txt
export SHAREPOINT_URL="https://...sharepoint.com/:x:/s/..."
python scripts/sync_sharepoint.py
```

Genera/actualiza `data.json` en la raiz del repo.

---

## Fallback: script local + Task Scheduler (Windows)

Si el tenant no permite links anonimos, podes correr el sync desde tu PC:

1. Sincroniza la carpeta de SharePoint con OneDrive Desktop.
2. Modifica `scripts/sync_sharepoint.py` para leer un path local:
   ```python
   # Reemplaza download_excel() por:
   blob = Path(os.environ["LOCAL_XLSX"]).read_bytes()
   ```
3. Configura Task Scheduler para correr cada N horas:
   ```
   python C:\ruta\al\repo\scripts\sync_sharepoint.py
   git -C C:\ruta\al\repo add data.json
   git -C C:\ruta\al\repo commit -m "sync"
   git -C C:\ruta\al\repo push
   ```

---

## Troubleshooting

**El job falla con "La URL devolvio HTML en vez del archivo".**
El link no permite descarga anonima. Verifica:
- Que el link sea "anyone with the link" (no "people in your organization").
- Que el admin del tenant haya habilitado links externos para esa biblioteca.

**El job falla con "Faltan columnas en la planilla".**
Revisa que los encabezados de la primera fila coincidan con alguno de los
alias en la tabla de arriba. Si las columnas tienen otro nombre, agregalo a
`COLUMN_ALIASES` en `scripts/sync_sharepoint.py`.

**El dashboard muestra "Error cargando datos".**
Abri DevTools -> Network y mira la respuesta de `data.json`. Si es 404,
verifica que el archivo este commiteado en la branch que esta desplegada.
