# Script de Presupuestos para Módulos de Carpintería

Este script automatiza el cálculo de materiales y costos para la fabricación de módulos de carpintería (bajomesadas, alacenas, etc.) usando Google Apps Script sobre Google Sheets.

## ¿Qué hace el script?

- Lee datos de módulos y configuraciones desde varias hojas de cálculo.
- Calcula automáticamente:
  - Cantidad de placas necesarias (por espesor y visibilidad)
  - Metros lineales de canto (blanco y color)
  - Herrajes requeridos y sus costos
  - Costos totales por módulo y combinaciones de materiales (todo blanco, solo frentes color, todo color, etc.)
- Genera hojas de resumen con:
  - Detalle de materiales y costos por módulo
  - Inventario global de materiales
  - Resúmenes generales de costos
  - Precios unitarios de referencia

## ¿Cómo se usa?

1. Vincula el script a una hoja de cálculo de Google Sheets con las siguientes hojas:
   - `Bajomesada`, `Alacena` (entrada de módulos)
   - `Config`, `Config2` (configuración de piezas)
   - `PreciosPlacas`, `PreciosCanto`, `Herrajes`, `ConfigHerrajes` (precios y herrajes)
2. Ejecuta la función principal `generarSalidaCompleta` desde el editor de Apps Script.
3. Revisa las hojas generadas automáticamente para ver los resultados y resúmenes.

## Personalización

- Puedes modificar las hojas de entrada y configuración para adaptarlas a tus necesidades.
- Los precios y descripciones se actualizan desde las hojas de precios.

## Requisitos

- Google Sheets
- Google Apps Script

## Funciones principales del script

- **generarSalidaCompleta**: Función principal que coordina la lectura de datos, el cálculo de materiales y costos, y la generación de hojas de resumen. Ejecuta todo el flujo de trabajo.
- **cargarPrecios**: Lee los precios de placas y cantos desde las hojas correspondientes y los estructura para su uso en los cálculos.
- **cargarHerrajes / cargarConfigHerrajes**: Obtienen la información de herrajes y su configuración por pieza desde las hojas respectivas.
- **calcularHerrajesPorPieza**: Determina la cantidad y costo de herrajes necesarios para cada pieza de módulo según la configuración.
- **calcularMetrosLinealesCanto**: Calcula los metros lineales de canto requeridos para cada pieza, diferenciando entre blanco y color.
- **combinarDatosDeHojas**: Fusiona datos de varias hojas de entrada o configuración para procesarlos de forma unificada.
- **generarHojaSalidaFinal**: Genera la hoja de resumen global con todos los costos, inventarios y precios unitarios.
- **escribirResumenEnHojasEntrada**: Escribe los resultados de los cálculos (metros de canto y costo total) en las hojas de entrada de módulos.

Estas funciones permiten automatizar el proceso de presupuestado, asegurando precisión y rapidez en la generación de resultados.

---

_Desarrollado para automatizar y optimizar la gestión de presupuestos en carpintería modular._
