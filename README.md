# Dashboard Convenio Atenciones Salud POL

Panel de seguimiento del convenio (Unidad de Seguros Publicos y Privados).

## Archivos
- `index.html`  -> el dashboard. Publicalo con GitHub Pages (Settings > Pages > selecciona la rama/carpeta).
- `data.json`   -> los datos que alimentan el dashboard (totales, series mensuales/anuales, transferencias y el listado completo de atenciones).
- `logo.png`    -> logo de la Unidad, usado en el encabezado.
- `admin.html`  -> herramienta para actualizar `data.json` cuando cambie el Excel (mas atenciones, mas transferencias).

## Como actualizar los datos cuando el Excel crezca
1. Abre `admin.html` en tu navegador (doble clic, no necesita internet salvo para cargar la libreria de lectura de Excel).
2. Arrastra o selecciona el Excel actualizado (debe tener las hojas "CONSOLIDADO" y "Transferencias" con las mismas columnas).
3. Presiona "Generar data.json" y revisa el resumen que aparece (totales, cantidad de atenciones, cobertura).
4. Descarga el archivo `data.json` generado.
5. Reemplaza el `data.json` de este repositorio por el nuevo archivo y sube el cambio (commit + push).
6. El dashboard en GitHub Pages se actualiza solo, sin tocar `index.html`.

## Nota
El listado completo de atenciones (con nombre y DNI del paciente) esta disponible para descarga en CSV desde el propio dashboard, boton "Descargar lista (CSV)". Es informacion personal de salud: compartir el enlace del dashboard solo con quienes deban verla.
