# Mapa Interactivo de Muestras Geológicas — UNINORTE

Mapa web interactivo desarrollado como parte de la tesis de pregrado en Geología de la Universidad del Norte (Barranquilla, Colombia).

## ¿De qué trata?

Visualización georreferenciada de **51 muestras de roca** recolectadas en ríos de América del Sur y Central, asociadas a plantas acuáticas de la familia *Podostemaceae*. Cada muestra fue analizada petrográficamente y en algunos casos mediante difracción de rayos X (DRX) y geoquímica de roca total.

## ¿Qué puedes ver en el mapa?

- **Ubicación** de cada muestra en el mapa (coordenadas originales o geocodificadas automáticamente)
- **Ficha completa** de cada muestra: tipo de roca, clasificación, especie vegetal asociada, autor, fecha y observaciones petrográficas
- **Fotos de muestra de mano** para 42 de las 51 muestras
- **Difractogramas DRX** (EVA) para 25 muestras
- **Láminas delgadas** en nícoles cruzados (NC) y paralelos (NP) para NYB25 y NYB37
- **Diagramas geoquímicos** por tipo de roca (TAS, AFM, QAPF, Harker, CIA, Winchester-Floyd, ACF, Zr-Ti, entre otros)

## Filtros disponibles

- Por país, tipo de roca, autor o especie vegetal
- Solo muestras con análisis DRX
- Solo muestras con análisis geoquímico
- Solo muestras con análisis petrográfico

## Tipos de roca

| Color | Tipo |
|-------|------|
| 🔴 Rojo | Ígnea |
| 🟡 Amarillo | Sedimentaria |
| 🟢 Verde | Metamórfica |
| ⚫ Gris | Indeterminada |

Los marcadores con borde discontinuo indican ubicación aproximada (geocodificada a partir del nombre del lugar).

## Ver el mapa

👉 **https://loresofi.github.io/tesis-mapa/mapa_muestras.html**

## Tecnologías

- [Leaflet.js](https://leafletjs.com/) — mapa interactivo
- [Leaflet.markercluster](https://github.com/Leaflet/Leaflet.markercluster) — agrupación de marcadores
- Python + openpyxl + matplotlib + pyrolite — procesamiento de datos y diagramas
- Google Drive — almacenamiento de imágenes
- GitHub Pages — publicación web

---

*Tesis de pregrado — Geología, Universidad del Norte, 2026*
