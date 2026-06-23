# Control Etiquetas/Remitos — Despacho RDC Troncal Rosario

Aplicación web (PWA) para **controlar el despacho** comparando dos archivos
Excel y detectar anomalías antes de imprimir los remitos.

- **Archivo A — Etiquetas a Controlar** (`TU*.xlsx`): debe contener la columna
  `Handling Unit`.
- **Archivo B — Impresos / Remitos** (`impresos*.xlsx`): debe contener las
  columnas `Handling Unit` y `Remito`.

La app compara ambos y genera un reporte con:

1. Códigos presentes en A pero **no** en B.
2. Códigos presentes en B pero **no** en A.
3. Handling Units que aparecen en **más de un Remito**.
4. Códigos duplicados dentro del archivo A.

El resultado se puede **exportar a PDF**.

Todo el procesamiento ocurre **en el navegador**; los archivos no se suben a
ningún servidor.

## Uso

1. Abrí la app (ver *Despliegue*).
2. **📥 Importar Etiquetas a Controlar** y elegí el `.xlsx` de etiquetas.
3. **📥 Importar Impresos** y elegí el `.xlsx` de remitos.
4. **🔍 Comparar**.
5. **📄 Exportar Resultado** para descargar el PDF.
6. **🗑️ Limpiar Datos** para reiniciar.

## Estructura del proyecto

```
.
├── index.html              # Estructura de la página
├── css/
│   └── styles.css          # Estilos (tema azul oscuro / rojo / blanco)
├── js/
│   └── app.js              # Lógica: importar Excel, comparar, exportar PDF
├── libs/                   # Librerías de terceros (vendored, sin CDN)
│   ├── xlsx.full.min.js    # SheetJS — lectura de Excel
│   └── jspdf.umd.min.js    # jsPDF — generación de PDF
├── icons/                  # Íconos de la PWA (192, 512, 512 maskable)
├── manifest.webmanifest    # Configuración PWA
├── service-worker.js       # Caché offline
└── LICENSE
```

## Desarrollo / ejecución local

Al usar un Service Worker, la app debe servirse por **HTTP** (no abrir el
`index.html` con `file://`). Desde la raíz del proyecto:

```bash
# Con Python
python -m http.server 8000

# o con Node
npx serve .
```

Luego abrí <http://localhost:8000>.

## Despliegue (GitHub Pages)

Los paths del `manifest.webmanifest` y del `service-worker.js` son **relativos**,
por lo que funciona tanto en la raíz de un dominio como en un subdirectorio de
GitHub Pages (`https://usuario.github.io/<repo>/`).

## Tecnologías

- HTML + CSS + JavaScript (sin framework ni build step).
- [SheetJS (xlsx)](https://sheetjs.com/) para leer Excel.
- [jsPDF](https://github.com/parallax/jsPDF) para generar PDF.
- PWA: `manifest.webmanifest` + `service-worker.js` (funciona offline).

## Notas

- Los íconos actuales (`icons/`) son **placeholders** con la marca RDC. Podés
  reemplazarlos por los oficiales manteniendo los mismos nombres y tamaños.
- Cada vez que cambies archivos cacheados, subí la versión `CACHE_NAME` en
  `service-worker.js` para forzar la actualización en los clientes.
