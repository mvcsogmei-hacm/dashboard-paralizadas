# Tablero de Obras Paralizadas y Reactivadas — MVCS · OGMEI

Dashboard ejecutivo interactivo (HTML + [Apache ECharts](https://echarts.apache.org/)) con dos páginas
— **Paralizadas** y **Reactivadas** — conectado a los datos del Excel institucional, con filtrado
cruzado estilo Power BI, drill-down en mapa y series de tiempo, y tema claro/oscuro.

## Estructura

```
index.html                  Aplicación completa (portada + 2 dashboards)
js/
  echarts.min.js            Librería de gráficos
  peru-geo.js               Mapa del Perú (GeoJSON precargado)
Data/
  data.xlsx                 Fuente de datos (editable)
  data.js                   Datos generados para el dashboard (NO editar a mano)
  build-data.py             Conversor Excel → data.js
  actualizar-datos.bat      Doble clic para regenerar data.js tras editar el Excel
  peru.geojson              Fuente original del mapa
img/
  cover.jpg                 Fondo de la portada
docs/                       Material de diseño (propuestas, guías, capturas)
```

## Actualizar los datos

1. Edita `Data/data.xlsx` (misma estructura de columnas).
2. Doble clic en `Data/actualizar-datos.bat` (requiere Python + `openpyxl`).
3. Sube el `Data/data.js` regenerado al repositorio.

## Publicar en GitHub Pages

1. Sube todo el contenido de esta carpeta a la raíz del repositorio.
2. En GitHub: **Settings → Pages → Source: Deploy from a branch → main / (root)**.
3. La página queda en `https://<usuario>.github.io/<repo>/`.

> El archivo `.nojekyll` evita que GitHub Pages procese el sitio con Jekyll.
> Importante: GitHub Pages distingue mayúsculas/minúsculas — mantener los nombres
> de carpetas tal cual (`Data/`, `js/`, `img/`).
