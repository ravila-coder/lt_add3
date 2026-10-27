# Portfolio Analysis

Convierte una cinta de créditos (CSV o Excel), y opcionalmente los archivos de pagos esperados y recibidos, en un libro de Excel listo para comité con todas las métricas como fórmulas vivas.

**Los archivos nunca salen de tu computadora.** El análisis corre dentro del navegador (Python en WebAssembly con Pyodide). Vercel solo sirve la página; no hay servidor que reciba la cinta.

## Publicar en GitHub y Vercel

Todo el programa vive en **un solo archivo: `public/index.html`**. Es lo único que Vercel necesita servir.

1. En GitHub, en tu repositorio, sube `vercel.json` a la raíz y `index.html` dentro de la carpeta `public/` (si ya existe, reemplázalo).
2. En [vercel.com](https://vercel.com): *Add New... > Project*, elige el repositorio y *Deploy*, sin cambiar opciones. Si el proyecto ya existe, Vercel vuelve a publicar solo al subir el archivo.
3. Abre la URL que te da Vercel.

Si quieres que solo tu equipo pueda abrir la página: *Project > Settings > Deployment Protection*.

## Cómo se usa

1. La primera vez la página tarda unos 20 segundos en cargar el motor; después queda en caché.
2. *Scope*: *Tape analysis* para la cinta sola, o *Tape analysis + COC* para agregar cosechas y MOIC.
3. *Files*: la cinta es obligatoria (CSV, TXT o XLSX). En COC agrega pagos esperados y pagos recibidos.
4. *Columns*: revisa el tipo detectado de cada columna. Puedes crear columnas calculadas con SQL; los nombres con espacios o acentos van entre comillas dobles.
5. *Field mapping*: Saldo Insoluto y Días de Atraso son obligatorios. En COC además ID de crédito, monto originado y fecha de originación.
6. Arrastra columnas a Resumen (hasta 6), PAR, Concentración e Histogramas.
7. *Generate* y luego *Download*.

Una cinta de unos 16,000 créditos con 100,000 pagos tarda alrededor de 45 segundos en generar. Usa Chrome o Edge en computadora.

## Qué contiene el libro

- **Portfolio**: resumen con promedios ponderados por saldo y buckets de atraso.
- **MOIC**: MOIC por cosecha y su distribución.
- **COC**: triángulo de cosechas por meses en libros (MoB) y cobranza esperada contra recibida por mes.
- **PAR**: PAR 7, 15 y 30 por saldo y por número de créditos, por atributo.
- **Concentration**: 20 principales exposiciones por atributo con participación acumulada.
- **Histograms**: distribución de cada columna numérica elegida.
- **Auditoría de Datos**: nulos, llenado, integridad, riesgos inmediatos y optimizaciones.
- Al final, las tablas de datos de referencia.

Requiere Excel 365 (la auditoría usa `UNIQUE` y `FILTER`).

## Reglas

- Días de atraso en blanco cuentan como al corriente (0 días).
- Del archivo de pagos esperados se eliminan duplicados exactos (opción activable). Los pagos recibidos nunca se deduplican.
- Formatos mexicanos: separador de miles, punto decimal y fechas DD/MM/AAAA.

## Usarlo en tu computadora sin Vercel

Doble clic en `public/index.html` (o en `Iniciar.bat`). Funciona igual que en Vercel; solo necesita internet la primera vez para cargar el motor.

## Estructura

```
public/index.html   la aplicación completa (generada, es lo que se publica)
vercel.json         publica la carpeta public/
source/             el código fuente; no se publica
  index.html        la interfaz
  worker.js         carga Python (Pyodide) y ejecuta el análisis en segundo plano
  py/app.py         toda la lógica: lectura, tipos, SQL y armado del Excel
  py/bridge.py      conecta la página con app.py
  py/template.xlsx  plantilla del tear sheet (su formato es parte del resultado)
  py/Theme1.thmx    tema de colores que se incrusta en cada libro
  wheels/           librerías de Python (Flask, openpyxl) fijadas por versión
  build.py          empaqueta todo lo anterior dentro de public/index.html
```

Para cambiar el análisis: edita `source/py/app.py`, corre `python source/build.py` y sube el `public/index.html` resultante. Prueba rápida: `python source/py/app.py --selftest`.

**No subas cintas ni archivos de pagos al repositorio**: tienen nombres de acreditados. El `.gitignore` ya bloquea `.csv` y `.xlsx` en la raíz.

## Solución de problemas

- *"No se pudo iniciar el motor"*: la red bloquea `cdn.jsdelivr.net`. Prueba otra red o pide a TI que lo permita.
- La página se queda en *Generating...* con archivos muy grandes: espera; cierra otras pestañas pesadas si la computadora tiene poca memoria.
