/**
 * Función principal para generar la salida completa del presupuesto.
 */
function generarSalidaCompleta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Validate spreadsheet object
  if (!ss) {
    Logger.log('Error: No se pudo acceder a la hoja de cálculo activa.');
    throw new Error(
      'No se pudo acceder a la hoja de cálculo activa. ' +
        'Asegúrese de que el script esté vinculado a una hoja de cálculo.'
    );
  }

  const hojasEntrada = ['Bajomesada', 'Alacena'];
  const hojasConfig = ['Config', 'Config2'];

  // Ajustado el número de columnas para incluir 'Visible'
  const mergedEntradaData = combinarDatosDeHojas(ss, hojasEntrada, 2, 1, 8);
  const mergedConfigData = combinarDatosDeHojas(ss, hojasConfig, 2, 1, 12); // Ahora 12 columnas

  // Cargar datos de precios con información de color
  const herrajesData = cargarHerrajes(ss);
  const configHerrajesData = cargarConfigHerrajes(ss);
  const preciosData = cargarPrecios(ss);

  // --- Mover aquí la función getPrecioCanto para que tenga acceso a preciosData ---
  function getPrecioCanto(grosor, color) {
    return (
      preciosData.canto.find(
        (c) =>
          c.grosor.replace(',', '.').replace('mm', '').trim() ===
            grosor.toString().replace(',', '.').replace('mm', '').trim() &&
          c.color.toLowerCase() === color.toLowerCase()
      )?.precio || 0
    );
  }

  const salidaPorEspesor = {};
  const resumenPorModulo = {}; // Para la hoja de salida final

  // Acumuladores globales para el resumen final de materiales (cantidades)
  const totalPlacasGlobalM2 = {
    ocultas: {}, // m2 por espesor
    visibles: {}, // m2 por espesor
  };
  const totalMetrosLinealesCantoGlobal = {
    blanco: 0,
    color: 0,
  };
  const totalHerrajesGlobal = {}; // Cantidad total por código de herraje

  mergedEntradaData.forEach(
    (
      [
        modulo,
        altura,
        ancho,
        profundidad,
        cantidadModulos,
        puerta,
        estante,
        cajon,
        divisor,
      ],
      index
    ) => {
      if (!modulo || !altura || !ancho || !profundidad || !cantidadModulos)
        return;

      const moduloSinEspacios = modulo.toString().replace(/\s/g, '-');
      const moduloID = `${moduloSinEspacios}-${altura}x${ancho}x${profundidad}`;

      const entradaVariables = {
        estante: parseFloat(estante) || 0,
        cajon: parseFloat(cajon) || 0,
        puerta: parseFloat(puerta) || 0,
        divisor: parseFloat(divisor) || 0,
      };

      // Inicializar resumen para este módulo
      if (!resumenPorModulo[moduloID]) {
        resumenPorModulo[moduloID] = {
          modulo: modulo,
          dimensiones: `${altura}x${ancho}x${profundidad}`,
          cantidadModulos: parseFloat(cantidadModulos),
          placasPorEspesor: {
            ocultas: {}, // m2 por espesor
            visibles: {}, // m2 por espesor
          },
          metrosLinealesCanto: {
            blanco: 0,
            color: 0,
          },
          herrajes: [],
          costosCombinados: {}, // Se calculará al final por módulo
          costoTotal: 0, // Por defecto para la hoja de entrada
          filaOriginal: index + 2,
        };
      }

      // Procesar cada pieza del módulo
      mergedConfigData.forEach((configRow) => {
        const [
          confModulo,
          parte,
          cantidadBase,
          largoRaw,
          anchoRaw,
          posicion,
          offsetLargo,
          offsetAncho,
          espesor,
          tipoVariable,
          configuracionCanto,
          visibleRaw, // Nueva columna 'Visible'
        ] = configRow;

        if (!confModulo || !parte) return;

        const confModuloSinEspacios = confModulo.toString().replace(/\s/g, '-');
        const parteSinEspacios = parte.toString().replace(/\s/g, '-');

        if (confModuloSinEspacios !== moduloSinEspacios) return;

        let finalCantidad = parseFloat(cantidadBase) || 0;
        if (
          tipoVariable &&
          entradaVariables[tipoVariable.toString().toLowerCase()]
        ) {
          finalCantidad =
            entradaVariables[tipoVariable.toString().toLowerCase()];
        }

        const { largo, anchoCalculado } = calcularDimensiones(
          altura,
          ancho,
          profundidad,
          largoRaw,
          anchoRaw,
          posicion,
          parseFloat(offsetLargo) || 0,
          parseFloat(offsetAncho) || 0
        );

        const totalCantidad =
          finalCantidad * (parseFloat(cantidadModulos) || 0);

        // Calcular metros cuadrados de esta pieza
        const m2Pieza =
          (largo / 1000) * (anchoCalculado / 1000) * totalCantidad;

        // Determinar si la pieza es visible
        const esVisible =
          visibleRaw && visibleRaw.toString().toLowerCase() === 'sí';

        // Acumular por espesor y visibilidad en el resumen del módulo
        const espesorKey = espesor ? espesor.toString() : 'Sin_Espesor';
        const tipoPlacaKey = esVisible ? 'visibles' : 'ocultas';

        if (
          !resumenPorModulo[moduloID].placasPorEspesor[tipoPlacaKey][espesorKey]
        ) {
          resumenPorModulo[moduloID].placasPorEspesor[tipoPlacaKey][
            espesorKey
          ] = {
            m2Total: 0,
          };
        }
        resumenPorModulo[moduloID].placasPorEspesor[tipoPlacaKey][
          espesorKey
        ].m2Total += m2Pieza;

        // Acumular m2 en el total global de placas (separado por si es oculta o visible)
        if (!totalPlacasGlobalM2[tipoPlacaKey][espesorKey]) {
          totalPlacasGlobalM2[tipoPlacaKey][espesorKey] = 0;
        }
        totalPlacasGlobalM2[tipoPlacaKey][espesorKey] += m2Pieza;

        // Calcular metros lineales de canto
        if (configuracionCanto) {
          const metrosCantoLineal = calcularMetrosLinealesCanto(
            configuracionCanto.toString(),
            largo,
            anchoCalculado,
            totalCantidad
          );
          // Si la pieza es visible, su canto se considera 'color'. Si es oculta, 'blanco'.
          const cantoTipo = esVisible ? 'color' : 'blanco';
          resumenPorModulo[moduloID].metrosLinealesCanto[cantoTipo] +=
            metrosCantoLineal;

          // Acumular en el total global de cantos
          totalMetrosLinealesCantoGlobal[cantoTipo] += metrosCantoLineal;
        }

        // Calcular herrajes para esta pieza
        const herrajesInfo = calcularHerrajesPorPieza(
          confModuloSinEspacios,
          parteSinEspacios,
          totalCantidad,
          entradaVariables,
          configHerrajesData,
          herrajesData
        );

        // Agregar herrajes al resumen del módulo y al total global
        herrajesInfo.forEach((herraje) => {
          // Para el módulo
          const herrajeExistenteModulo = resumenPorModulo[
            moduloID
          ].herrajes.find((h) => h.codigo === herraje.codigo);
          if (herrajeExistenteModulo) {
            herrajeExistenteModulo.cantidad += herraje.cantidad;
            herrajeExistenteModulo.costoTotal += herraje.costoTotal;
          } else {
            resumenPorModulo[moduloID].herrajes.push({ ...herraje });
          }

          // Para el total global de herrajes
          if (!totalHerrajesGlobal[herraje.codigo]) {
            totalHerrajesGlobal[herraje.codigo] = {
              cantidad: 0,
              costoTotal: 0,
              descripcion: herraje.descripcion,
              precioUnitario: herraje.precioUnitario,
            };
          }
          totalHerrajesGlobal[herraje.codigo].cantidad += herraje.cantidad;
          totalHerrajesGlobal[herraje.codigo].costoTotal += herraje.costoTotal;
        });

        // Para las hojas de salida por espesor (mantener funcionalidad existente)
        const id = `${moduloID}-${parteSinEspacios}`.toLowerCase();
        const espesorKeyOutput = espesor
          ? `Placa_${espesor.toString()}mm`
          : 'Placa_Sin_Espesor';

        if (!salidaPorEspesor[espesorKeyOutput]) {
          salidaPorEspesor[espesorKeyOutput] = [
            ['Cantidad', 'Largo (mm)', 'Ancho (mm)', 'Detalle', 'm²'],
          ];
        }

        salidaPorEspesor[espesorKeyOutput].push([
          totalCantidad,
          largo,
          anchoCalculado,
          id,
          m2Pieza.toFixed(3).replace('.', ','),
        ]);
      });
    }
  );

  // Calcular costos combinados por módulo
  Object.keys(resumenPorModulo).forEach((moduloID) => {
    const modulo = resumenPorModulo[moduloID];

    modulo.costosCombinados = {
      todoBlanco: {
        placas: 0,
        canto: 0,
        herrajes: modulo.herrajes.reduce((sum, h) => sum + h.costoTotal, 0),
        total: 0,
      },
      soloOcultasBlancas: {
        placas: 0,
        canto: 0,
        herrajes: modulo.herrajes.reduce((sum, h) => sum + h.costoTotal, 0),
        total: 0,
      },
      todoColor: {
        placas: 0,
        canto: 0,
        herrajes: modulo.herrajes.reduce((sum, h) => sum + h.costoTotal, 0),
        total: 0,
      },
    };

    // --- NUEVO: Calcular costos de canto para 0,45mm y 2mm (por color) ---
    // Definir combinaciones
    modulo.costosCombinados = {
      // 0,45mm
      todoBlanco_045: { placas: 0, canto: 0, herrajes: 0, total: 0 },
      soloOcultasBlancas_045: { placas: 0, canto: 0, herrajes: 0, total: 0 },
      todoColor_045: { placas: 0, canto: 0, herrajes: 0, total: 0 },
      // 2mm en visibles
      todoBlanco_2: { placas: 0, canto: 0, herrajes: 0, total: 0 },
      soloOcultasBlancas_2: { placas: 0, canto: 0, herrajes: 0, total: 0 },
      todoColor_2: { placas: 0, canto: 0, herrajes: 0, total: 0 },
    };

    // --- Asegúrate de inicializar las propiedades .placas para cada combinación ---
    // Copiar el cálculo de placas para cada combinación
    // Todo Blanco: Todas las placas con precio de blanco
    Object.keys(modulo.placasPorEspesor.ocultas).forEach((espesor) => {
      const placaInfo = modulo.placasPorEspesor.ocultas[espesor];
      const precioPlacaBlanco = preciosData.placas?.[espesor]?.find(
        (p) => p.color.toLowerCase() === 'blanco'
      )?.precio;
      if (precioPlacaBlanco !== undefined) {
        placaInfo.placasNecesarias = Math.ceil(placaInfo.m2Total / 4.76);
        modulo.costosCombinados.todoBlanco_045.placas +=
          placaInfo.placasNecesarias * precioPlacaBlanco;
        modulo.costosCombinados.todoBlanco_2.placas +=
          placaInfo.placasNecesarias * precioPlacaBlanco;
      }
    });
    Object.keys(modulo.placasPorEspesor.visibles).forEach((espesor) => {
      const placaInfo = modulo.placasPorEspesor.visibles[espesor];
      const precioPlacaBlanco = preciosData.placas?.[espesor]?.find(
        (p) => p.color.toLowerCase() === 'blanco'
      )?.precio;
      if (precioPlacaBlanco !== undefined) {
        placaInfo.placasNecesarias = Math.ceil(placaInfo.m2Total / 4.76);
        modulo.costosCombinados.todoBlanco_045.placas +=
          placaInfo.placasNecesarias * precioPlacaBlanco;
        modulo.costosCombinados.todoBlanco_2.placas +=
          placaInfo.placasNecesarias * precioPlacaBlanco;
      }
    });

    // Solo Ocultas Blancas (Visibles con color): Ocultas con precio de blanco, visibles con precio de color
    Object.keys(modulo.placasPorEspesor.ocultas).forEach((espesor) => {
      const placaInfo = modulo.placasPorEspesor.ocultas[espesor];
      const precioPlacaBlanco = preciosData.placas?.[espesor]?.find(
        (p) => p.color.toLowerCase() === 'blanco'
      )?.precio;
      if (precioPlacaBlanco !== undefined) {
        placaInfo.placasNecesarias = Math.ceil(placaInfo.m2Total / 4.76);
        modulo.costosCombinados.soloOcultasBlancas_045.placas +=
          placaInfo.placasNecesarias * precioPlacaBlanco;
        modulo.costosCombinados.soloOcultasBlancas_2.placas +=
          placaInfo.placasNecesarias * precioPlacaBlanco;
      }
    });
    Object.keys(modulo.placasPorEspesor.visibles).forEach((espesor) => {
      const placaInfo = modulo.placasPorEspesor.visibles[espesor];
      const precioPlacaColor = preciosData.placas?.[espesor]?.find(
        (p) => p.color.toLowerCase() === 'color'
      )?.precio;
      if (precioPlacaColor !== undefined) {
        placaInfo.placasNecesarias = Math.ceil(placaInfo.m2Total / 4.76);
        modulo.costosCombinados.soloOcultasBlancas_045.placas +=
          placaInfo.placasNecesarias * precioPlacaColor;
        modulo.costosCombinados.soloOcultasBlancas_2.placas +=
          placaInfo.placasNecesarias * precioPlacaColor;
      } else {
        // Fallback a blanco si no hay precio de color
        const precioPlacaBlanco = preciosData.placas?.[espesor]?.find(
          (p) => p.color.toLowerCase() === 'blanco'
        )?.precio;
        if (precioPlacaBlanco !== undefined) {
          placaInfo.placasNecesarias = Math.ceil(placaInfo.m2Total / 4.76);
          modulo.costosCombinados.soloOcultasBlancas_045.placas +=
            placaInfo.placasNecesarias * precioPlacaBlanco;
          modulo.costosCombinados.soloOcultasBlancas_2.placas +=
            placaInfo.placasNecesarias * precioPlacaBlanco;
        }
      }
    });

    // Todo Color: Todas las placas con precio de color (con fallback a blanco si no existe)
    Object.keys(modulo.placasPorEspesor.ocultas).forEach((espesor) => {
      const placaInfo = modulo.placasPorEspesor.ocultas[espesor];
      const precioPlacaColor = preciosData.placas?.[espesor]?.find(
        (p) => p.color.toLowerCase() === 'color'
      )?.precio;
      if (precioPlacaColor !== undefined) {
        placaInfo.placasNecesarias = Math.ceil(placaInfo.m2Total / 4.76);
        modulo.costosCombinados.todoColor_045.placas +=
          placaInfo.placasNecesarias * precioPlacaColor;
        modulo.costosCombinados.todoColor_2.placas +=
          placaInfo.placasNecesarias * precioPlacaColor;
      } else {
        const precioPlacaBlanco = preciosData.placas?.[espesor]?.find(
          (p) => p.color.toLowerCase() === 'blanco'
        )?.precio;
        if (precioPlacaBlanco !== undefined) {
          placaInfo.placasNecesarias = Math.ceil(placaInfo.m2Total / 4.76);
          modulo.costosCombinados.todoColor_045.placas +=
            placaInfo.placasNecesarias * precioPlacaBlanco;
          modulo.costosCombinados.todoColor_2.placas +=
            placaInfo.placasNecesarias * precioPlacaBlanco;
        }
      }
    });
    Object.keys(modulo.placasPorEspesor.visibles).forEach((espesor) => {
      const placaInfo = modulo.placasPorEspesor.visibles[espesor];
      const precioPlacaColor = preciosData.placas?.[espesor]?.find(
        (p) => p.color.toLowerCase() === 'color'
      )?.precio;
      if (precioPlacaColor !== undefined) {
        placaInfo.placasNecesarias = Math.ceil(placaInfo.m2Total / 4.76);
        modulo.costosCombinados.todoColor_045.placas +=
          placaInfo.placasNecesarias * precioPlacaColor;
        modulo.costosCombinados.todoColor_2.placas +=
          placaInfo.placasNecesarias * precioPlacaColor;
      } else {
        const precioPlacaBlanco = preciosData.placas?.[espesor]?.find(
          (p) => p.color.toLowerCase() === 'blanco'
        )?.precio;
        if (precioPlacaBlanco !== undefined) {
          placaInfo.placasNecesarias = Math.ceil(placaInfo.m2Total / 4.76);
          modulo.costosCombinados.todoColor_045.placas +=
            placaInfo.placasNecesarias * precioPlacaBlanco;
          modulo.costosCombinados.todoColor_2.placas +=
            placaInfo.placasNecesarias * precioPlacaBlanco;
        }
      }
    });

    // --- Calcular costos de canto para 0,45mm y 2mm (por color) ---
    // 0,45mm
    const precioCantoBlanco_045 = getPrecioCanto('0.45', 'blanco');
    const precioCantoColor_045 = getPrecioCanto('0.45', 'color');
    // 2mm
    const precioCantoBlanco_2 = getPrecioCanto('2', 'blanco');
    const precioCantoColor_2 = getPrecioCanto('2', 'color');

    // 0,45mm: todos los cantos usan 0,45mm
    modulo.costosCombinados.todoBlanco_045.canto =
      (modulo.metrosLinealesCanto.blanco + modulo.metrosLinealesCanto.color) *
      (precioCantoBlanco_045 || 0);
    modulo.costosCombinados.soloOcultasBlancas_045.canto =
      modulo.metrosLinealesCanto.blanco * (precioCantoBlanco_045 || 0) +
      modulo.metrosLinealesCanto.color * (precioCantoColor_045 || 0);
    modulo.costosCombinados.todoColor_045.canto =
      (modulo.metrosLinealesCanto.blanco + modulo.metrosLinealesCanto.color) *
      (precioCantoColor_045 || precioCantoBlanco_045 || 0);

    // 2mm: piezas visibles usan 2mm, ocultas usan 0,45mm
    // Asumimos metrosLinealesCanto.color = visibles, blanco = ocultas
    modulo.costosCombinados.todoBlanco_2.canto =
      modulo.metrosLinealesCanto.blanco * (precioCantoBlanco_045 || 0) +
      modulo.metrosLinealesCanto.color * (precioCantoBlanco_2 || 0);
    modulo.costosCombinados.soloOcultasBlancas_2.canto =
      modulo.metrosLinealesCanto.blanco * (precioCantoBlanco_045 || 0) +
      modulo.metrosLinealesCanto.color * (precioCantoColor_2 || 0);
    modulo.costosCombinados.todoColor_2.canto =
      modulo.metrosLinealesCanto.blanco *
        (precioCantoColor_045 || precioCantoBlanco_045 || 0) +
      modulo.metrosLinealesCanto.color *
        (precioCantoColor_2 || precioCantoBlanco_2 || 0);

    // --- Calcular Totales Combinados para este módulo ---
    Object.keys(modulo.costosCombinados).forEach((combinacion) => {
      modulo.costosCombinados[combinacion].total =
        modulo.costosCombinados[combinacion].placas +
        modulo.costosCombinados[combinacion].canto +
        modulo.costosCombinados[combinacion].herrajes;
    });

    // Por defecto para la columna P de la hoja de entrada
    modulo.costoTotal = modulo.costosCombinados.todoBlanco_045.total;
  });

  // --- Calcular el resumen global final de costos (placas y cantos) para las 6 combinaciones ---
  const resumenGlobalCostos = {
    placas: {
      todoBlanco: 0,
      soloOcultasBlancas: 0,
      todoColor: 0,
    },
    canto: {
      todoBlanco_045: 0,
      soloOcultasBlancas_045: 0,
      todoColor_045: 0,
      todoBlanco_2: 0,
      soloOcultasBlancas_2: 0,
      todoColor_2: 0,
    },
    herrajes: {
      total: Object.values(totalHerrajesGlobal).reduce(
        (sum, h) => sum + h.costoTotal,
        0
      ),
      detalle: Object.values(totalHerrajesGlobal), // Para la tabla de herrajes totales
    },
    // Inventario de materiales (cantidades sin costo aún)
    inventario: {
      placas: { ocultas: {}, visibles: {} },
      canto: {
        blanco: totalMetrosLinealesCantoGlobal.blanco,
        color: totalMetrosLinealesCantoGlobal.color,
      },
    },
  };

  // Canto Global: Calcular los totales de canto para cada combinación global
  // 0,45mm
  resumenGlobalCostos.canto.todoBlanco_045 =
    (totalMetrosLinealesCantoGlobal.blanco +
      totalMetrosLinealesCantoGlobal.color) *
    getPrecioCanto('0.45', 'blanco');
  resumenGlobalCostos.canto.soloOcultasBlancas_045 =
    totalMetrosLinealesCantoGlobal.blanco * getPrecioCanto('0.45', 'blanco') +
    totalMetrosLinealesCantoGlobal.color * getPrecioCanto('0.45', 'color');
  resumenGlobalCostos.canto.todoColor_045 =
    (totalMetrosLinealesCantoGlobal.blanco +
      totalMetrosLinealesCantoGlobal.color) *
    (getPrecioCanto('0.45', 'color') || getPrecioCanto('0.45', 'blanco') || 0);

  // 2mm en visibles
  resumenGlobalCostos.canto.todoBlanco_2 =
    totalMetrosLinealesCantoGlobal.blanco * getPrecioCanto('0.45', 'blanco') +
    totalMetrosLinealesCantoGlobal.color * getPrecioCanto('2', 'blanco');
  resumenGlobalCostos.canto.soloOcultasBlancas_2 =
    totalMetrosLinealesCantoGlobal.blanco * getPrecioCanto('0.45', 'blanco') +
    totalMetrosLinealesCantoGlobal.color * getPrecioCanto('2', 'color');
  resumenGlobalCostos.canto.todoColor_2 =
    totalMetrosLinealesCantoGlobal.blanco *
      (getPrecioCanto('0.45', 'color') ||
        getPrecioCanto('0.45', 'blanco') ||
        0) +
    totalMetrosLinealesCantoGlobal.color *
      (getPrecioCanto('2', 'color') || getPrecioCanto('2', 'blanco') || 0);

  // Placas Global y detalle: Calcular los totales de placa para cada combinación global
  // Procesar placas "ocultas"
  Object.keys(totalPlacasGlobalM2.ocultas).forEach((espesor) => {
    const m2 = totalPlacasGlobalM2.ocultas[espesor];
    const placasNecesarias = Math.ceil(m2 / 4.76);
    const precioBlanco = preciosData.placas[espesor]?.find(
      (p) => p.color.toLowerCase() === 'blanco'
    )?.precio;
    const precioColor = preciosData.placas[espesor]?.find(
      (p) => p.color.toLowerCase() === 'color'
    )?.precio;

    // Guardar para el inventario
    resumenGlobalCostos.inventario.placas.ocultas[espesor] = {
      m2Total: m2,
      placasNecesarias: placasNecesarias,
    };

    // Acumular para las combinaciones globales de placas
    resumenGlobalCostos.placas.todoBlanco +=
      placasNecesarias * (precioBlanco || 0);
    resumenGlobalCostos.placas.soloOcultasBlancas +=
      placasNecesarias * (precioBlanco || 0); // Las ocultas siempre son blancas aquí
    resumenGlobalCostos.placas.todoColor +=
      placasNecesarias * (precioColor || precioBlanco || 0); // Ocultas intentan ser color, sino blanco
  });

  // Procesar placas "visibles"
  Object.keys(totalPlacasGlobalM2.visibles).forEach((espesor) => {
    const m2 = totalPlacasGlobalM2.visibles[espesor];
    const placasNecesarias = Math.ceil(m2 / 4.76);
    const precioBlanco = preciosData.placas[espesor]?.find(
      (p) => p.color.toLowerCase() === 'blanco'
    )?.precio;
    const precioColor = preciosData.placas[espesor]?.find(
      (p) => p.color.toLowerCase() === 'color'
    )?.precio;

    // Guardar para el inventario
    resumenGlobalCostos.inventario.placas.visibles[espesor] = {
      m2Total: m2,
      placasNecesarias: placasNecesarias,
    };

    // Acumular para las combinaciones globales de placas
    resumenGlobalCostos.placas.todoBlanco +=
      placasNecesarias * (precioBlanco || 0); // Las visibles también van con blanco aquí
    resumenGlobalCostos.placas.soloOcultasBlancas +=
      placasNecesarias * (precioColor || precioBlanco || 0); // Las visibles van con color aquí
    resumenGlobalCostos.placas.todoColor +=
      placasNecesarias * (precioColor || precioBlanco || 0); // Las visibles también van con color aquí
  });

  // Generar hojas de salida
  generarHojaSalidaFinal(
    ss,
    resumenPorModulo,
    resumenGlobalCostos,
    preciosData
  );
  escribirResumenEnHojasEntrada(ss, hojasEntrada, resumenPorModulo);

  // Generar hojas por espesor (mantener funcionalidad existente)
  for (const espesor in salidaPorEspesor) {
    let sheet = ss.getSheetByName(espesor);
    if (!sheet) {
      sheet = ss.insertSheet(espesor);
    } else {
      sheet.clearContents();
    }

    const data = salidaPorEspesor[espesor];
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    sheet
      .getRange(1, 1, 1, data[0].length)
      .setFontWeight('bold')
      .setBackground('#E8F0FE');
  }
}

/**
 * Carga los precios desde la hoja de precios
 */
function cargarPrecios(ss) {
  const precios = {
    placas: {},
    canto: [], // array de {grosor, precio, color}
  };

  // Cargar precios de placas
  const hojaPlacas = ss.getSheetByName('PreciosPlacas');
  if (hojaPlacas) {
    // Ahora incluye la columna de color
    const datosPlacas = hojaPlacas
      .getRange(2, 1, hojaPlacas.getLastRow() - 1, 4)
      .getValues();
    datosPlacas.forEach(([espesor, precio, descripcion, color]) => {
      if (espesor) {
        const espesorStr = espesor.toString();
        if (!precios.placas[espesorStr]) {
          precios.placas[espesorStr] = [];
        }
        precios.placas[espesorStr].push({
          precio: parseFloat(precio) || 0,
          descripcion: descripcion ? descripcion.toString() : 'Sin descripción',
          color: color ? color.toString().toLowerCase() : 'blanco', // Valor por defecto
        });
      }
    });
  }

  // Cargar precio de canto (ahora incluye grosor)
  const hojaCanto = ss.getSheetByName('PreciosCanto');
  if (hojaCanto) {
    const datosCanto = hojaCanto
      .getRange(2, 1, hojaCanto.getLastRow() - 1, 3)
      .getValues();
    datosCanto.forEach(([grosor, precio, color]) => {
      if (precio) {
        precios.canto.push({
          grosor: grosor ? grosor.toString() : '0,45', // por defecto
          precio: parseFloat(precio) || 0,
          color: color ? color.toString().toLowerCase() : 'blanco',
        });
      }
    });
  }

  return precios;
}

/**
 * Genera la hoja de salida final con resumen completo
 * Ahora recibe resumenGlobalCostos
 */
function generarHojaSalidaFinal(
  ss,
  resumenPorModulo,
  resumenGlobalCostos,
  preciosData
) {
  // Validate spreadsheet object
  if (!ss) {
    Logger.log('Error: No se pudo acceder a la hoja de cálculo activa.');
    throw new Error(
      'No se pudo acceder a la hoja de cálculo activa. ' +
        'Asegúrese de que el script esté vinculado a una hoja de cálculo.'
    );
  }

  let hojaSalida = ss.getSheetByName('Resumen_Costos');
  if (!hojaSalida) {
    hojaSalida = ss.insertSheet('Resumen_Costos');
    Logger.log('La hoja "Resumen_Costos" no existía y fue creada.');
  } else {
    hojaSalida.clearContents();
  }

  const datos = [];
  let filaActualOffset = 0;

  // Encabezados principales
  datos.push(['RESUMEN DE COSTOS POR MÓDULO', '', '', '', '', '', '', '']);
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset += 2;

  Object.keys(resumenPorModulo).forEach((moduloID) => {
    const modulo = resumenPorModulo[moduloID];

    // Información del módulo
    datos.push([
      `MÓDULO: ${modulo.modulo} (${modulo.dimensiones}) - Cantidad: ${modulo.cantidadModulos}`,
      '',
      '',
      '',
      '',
      '',
      '',
      '',
    ]);
    datos.push(['', '', '', '', '', '', '', '']);
    filaActualOffset += 2;

    // Resumen de costos por combinación (por módulo)
    datos.push([
      'RESUMEN COMBINACIONES POR MÓDULO',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
    ]);
    datos.push([
      'Combinación',
      'Costo Placas',
      'Costo Cantos',
      'Costo Herrajes',
      'COSTO TOTAL FINAL',
    ]);
    filaActualOffset += 2;

    Object.keys(modulo.costosCombinados).forEach((combinacionKey) => {
      let nombreCombinacion = '';
      switch (combinacionKey) {
        case 'todoBlanco':
          nombreCombinacion = 'Todo Blanco';
          break;
        case 'soloOcultasBlancas':
          nombreCombinacion = 'Solo Ocultas Blancas (Visibles Color)';
          break;
        case 'todoColor':
          nombreCombinacion = 'Todo Color';
          break;
        default:
          nombreCombinacion = combinacionKey;
      }

      const costos = modulo.costosCombinados[combinacionKey];
      datos.push([
        nombreCombinacion,
        `$${costos.placas.toFixed(2).replace('.', ',')}`,
        `$${costos.canto.toFixed(2).replace('.', ',')}`,
        `$${costos.herrajes.toFixed(2).replace('.', ',')}`,
        `$${costos.total.toFixed(2).replace('.', ',')}`,
      ]);
      filaActualOffset++;
    });

    datos.push(['', '', '', '', '', '', '', '']);
    filaActualOffset++;

    // Detalle de placas (separado por visibles y ocultas) para el módulo actual
    datos.push(['DETALLE DE PLACAS POR MÓDULO', '', '', '', '', '', '', '']);
    datos.push(['TIPO', 'Espesor', 'm² Necesarios', 'Placas Enteras']);
    filaActualOffset += 2;

    // Placas Ocultas
    if (Object.keys(modulo.placasPorEspesor.ocultas).length > 0) {
      datos.push(['PIEZAS OCULTAS', '', '', '', '', '', '', '']);
      filaActualOffset++;
      Object.keys(modulo.placasPorEspesor.ocultas).forEach((espesor) => {
        const placa = modulo.placasPorEspesor.ocultas[espesor];
        datos.push([
          '',
          `${espesor}mm`,
          placa.m2Total.toFixed(2).replace('.', ','),
          placa.placasNecesarias,
        ]);
        filaActualOffset++;
      });
    }

    // Placas Visibles
    if (Object.keys(modulo.placasPorEspesor.visibles).length > 0) {
      datos.push(['PIEZAS VISIBLES', '', '', '', '', '', '', '']);
      filaActualOffset++;
      Object.keys(modulo.placasPorEspesor.visibles).forEach((espesor) => {
        const placa = modulo.placasPorEspesor.visibles[espesor];
        datos.push([
          '',
          `${espesor}mm`,
          placa.m2Total.toFixed(2).replace('.', ','),
          placa.placasNecesarias,
        ]);
        filaActualOffset++;
      });
    }
    datos.push(['', '', '', '', '', '', '', '']);
    filaActualOffset++;

    // Detalle de cantos por módulo
    datos.push(['DETALLE DE CANTOS POR MÓDULO', '', '', '', '', '', '', '']);
    datos.push(['Tipo Canto', 'Metros Lineales']);
    filaActualOffset += 2;

    datos.push([
      'Blanco',
      modulo.metrosLinealesCanto.blanco.toFixed(2).replace('.', ','),
    ]);
    filaActualOffset++;
    datos.push([
      'Color',
      modulo.metrosLinealesCanto.color.toFixed(2).replace('.', ','),
    ]);
    filaActualOffset++;
    datos.push(['', '', '', '', '', '', '', '']);
    filaActualOffset++;

    // Detalle de herrajes por módulo
    if (modulo.herrajes.length > 0) {
      datos.push([
        'DETALLE DE HERRAJES POR MÓDULO',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
      ]);
      datos.push(['Código', 'Descripción', 'Cantidad', 'Precio Unitario']);
      filaActualOffset += 2;

      modulo.herrajes.forEach((herraje) => {
        datos.push([
          herraje.codigo,
          herraje.descripcion,
          herraje.cantidad,
          `$${herraje.precioUnitario.toFixed(2).replace('.', ',')}`,
        ]);
        filaActualOffset++;
      });
    } else {
      datos.push([
        'HERRAJES: Sin herrajes configurados',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
      ]);
      filaActualOffset++;
    }

    datos.push(['', '', '', '', '', '', '', '']);
    filaActualOffset++;
    datos.push(['', '', '', '', '', '', '', '']);
    filaActualOffset++;
  }); // Fin de la iteración por módulo

  // --- NUEVA SECCIÓN: INVENTARIO DE MATERIALES TOTALES ---
  datos.push(['INVENTARIO DE MATERIALES TOTALES', '', '', '', '', '', '', '']);
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset += 2;

  datos.push(['PLACAS NECESARIAS TOTALES', '', '', '', '', '', '', '']);
  datos.push([
    'Espesor',
    'Tipo Uso (Diseño)', // Oculta/Visible
    'm² Total',
    'Placas Enteras Necesarias',
  ]);
  filaActualOffset += 2;

  // Placas Ocultas Globales
  Object.keys(resumenGlobalCostos.inventario.placas.ocultas).forEach(
    (espesor) => {
      const placaInfo = resumenGlobalCostos.inventario.placas.ocultas[espesor];
      datos.push([
        `${espesor}mm`,
        'Oculta',
        placaInfo.m2Total.toFixed(2).replace('.', ','),
        placaInfo.placasNecesarias,
      ]);
      filaActualOffset++;
    }
  );

  // Placas Visibles Globales
  Object.keys(resumenGlobalCostos.inventario.placas.visibles).forEach(
    (espesor) => {
      const placaInfo = resumenGlobalCostos.inventario.placas.visibles[espesor];
      datos.push([
        `${espesor}mm`,
        'Visible',
        placaInfo.m2Total.toFixed(2).replace('.', ','),
        placaInfo.placasNecesarias,
      ]);
      filaActualOffset++;
    }
  );

  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset++;

  datos.push(['CANTOS NECESARIOS TOTALES', '', '', '', '', '', '', '']);
  datos.push(['Tipo Uso (Diseño)', 'Metros Lineales Total']);
  filaActualOffset += 2;

  datos.push([
    'Blanco (Oculto)',
    resumenGlobalCostos.inventario.canto.blanco.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Color (Visible)',
    resumenGlobalCostos.inventario.canto.color.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;

  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset++;

  // --- NUEVA SECCIÓN: PRECIOS UNITARIOS DE REFERENCIA ---
  datos.push(['PRECIOS UNITARIOS DE REFERENCIA', '', '', '', '', '', '', '']);
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset += 2;

  datos.push(['PRECIOS DE PLACAS POR UNIDAD', '', '', '', '', '', '', '']);
  datos.push(['Espesor', 'Descripción', 'Color', 'Precio ($)']);
  filaActualOffset += 2;

  Object.keys(preciosData.placas).forEach((espesor) => {
    preciosData.placas[espesor].forEach((placaPrecio) => {
      datos.push([
        `${espesor}mm`,
        placaPrecio.descripcion,
        placaPrecio.color,
        placaPrecio.precio.toFixed(2).replace('.', ','),
      ]);
      filaActualOffset++;
    });
  });

  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset++;

  datos.push(['PRECIOS DE CANTOS POR METRO', '', '', '', '', '', '', '']);
  datos.push(['Tipo', 'Color', 'Precio ($)']);
  filaActualOffset += 2;

  preciosData.canto.forEach((cantoPrecio) => {
    datos.push([
      'General', // O el tipo si lo tuvieras en PreciosCanto
      cantoPrecio.color,
      cantoPrecio.precio.toFixed(2).replace('.', ','),
    ]);
    filaActualOffset++;
  });

  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset++;

  // --- TRES RESÚMENES GENERALES DE COSTOS ---

  // RESUMEN GENERAL: Todo Blanco
  datos.push(['RESUMEN GENERAL: TODO BLANCO', '', '', '', '', '', '', '']);
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset += 2;
  datos.push(['Componente', 'Costo Total ($)']);
  filaActualOffset++;
  datos.push([
    'Placas',
    resumenGlobalCostos.placas.todoBlanco.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Cantos',
    resumenGlobalCostos.canto.todoBlanco_045.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Herrajes',
    resumenGlobalCostos.herrajes.total.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'COSTO TOTAL GENERAL (TODO BLANCO)',
    (
      resumenGlobalCostos.placas.todoBlanco +
      resumenGlobalCostos.canto.todoBlanco_045 +
      resumenGlobalCostos.herrajes.total
    )
      .toFixed(2)
      .replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset++;

  // RESUMEN GENERAL: Solo Ocultas Blancas (Visibles Color)
  datos.push([
    'RESUMEN GENERAL: SOLO OCULTAS BLANCAS (VISIBLES COLOR)',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
  ]);
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset += 2;
  datos.push(['Componente', 'Costo Total ($)']);
  filaActualOffset++;
  datos.push([
    'Placas',
    resumenGlobalCostos.placas.soloOcultasBlancas.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Cantos',
    resumenGlobalCostos.canto.soloOcultasBlancas_045
      .toFixed(2)
      .replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Herrajes',
    resumenGlobalCostos.herrajes.total.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'COSTO TOTAL GENERAL (SOLO OCULTAS BLANCAS)',
    (
      resumenGlobalCostos.placas.soloOcultasBlancas +
      resumenGlobalCostos.canto.soloOcultasBlancas_045 +
      resumenGlobalCostos.herrajes.total
    )
      .toFixed(2)
      .replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset++;

  // RESUMEN GENERAL: Todo Color
  datos.push(['RESUMEN GENERAL: TODO COLOR', '', '', '', '', '', '', '']);
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset += 2;
  datos.push(['Componente', 'Costo Total ($)']);
  filaActualOffset++;
  datos.push([
    'Placas',
    resumenGlobalCostos.placas.todoColor.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Cantos',
    resumenGlobalCostos.canto.todoColor_045.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Herrajes',
    resumenGlobalCostos.herrajes.total.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'COSTO TOTAL GENERAL (TODO COLOR)',
    (
      resumenGlobalCostos.placas.todoColor +
      resumenGlobalCostos.canto.todoColor_045 +
      resumenGlobalCostos.herrajes.total
    )
      .toFixed(2)
      .replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset++;

  // RESUMEN GENERAL: Todo Blanco (placa, canto 2mm en visibles)
  datos.push([
    'RESUMEN GENERAL: TODO BLANCO (placa, canto 2mm en visibles)',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
  ]);
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset += 2;
  datos.push(['Componente', 'Costo Total ($)']);
  filaActualOffset++;
  datos.push([
    'Placas',
    resumenGlobalCostos.placas.todoBlanco.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Cantos',
    resumenGlobalCostos.canto.todoBlanco_2.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Herrajes',
    resumenGlobalCostos.herrajes.total.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'COSTO TOTAL GENERAL',
    (
      resumenGlobalCostos.placas.todoBlanco +
      resumenGlobalCostos.canto.todoBlanco_2 +
      resumenGlobalCostos.herrajes.total
    )
      .toFixed(2)
      .replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset++;

  // RESUMEN GENERAL: Solo Ocultas Blancas (placa, canto 2mm en visibles)
  datos.push([
    'RESUMEN GENERAL: SOLO OCULTAS BLANCAS (placa, canto 2mm en visibles)',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
  ]);
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset += 2;
  datos.push(['Componente', 'Costo Total ($)']);
  filaActualOffset++;
  datos.push([
    'Placas',
    resumenGlobalCostos.placas.soloOcultasBlancas.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Cantos',
    resumenGlobalCostos.canto.soloOcultasBlancas_2.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Herrajes',
    resumenGlobalCostos.herrajes.total.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'COSTO TOTAL GENERAL',
    (
      resumenGlobalCostos.placas.soloOcultasBlancas +
      resumenGlobalCostos.canto.soloOcultasBlancas_2 +
      resumenGlobalCostos.herrajes.total
    )
      .toFixed(2)
      .replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset++;

  // RESUMEN GENERAL: Todo Color (placa, canto 2mm en visibles)
  datos.push([
    'RESUMEN GENERAL: TODO COLOR (placa, canto 2mm en visibles)',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
  ]);
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset += 2;
  datos.push(['Componente', 'Costo Total ($)']);
  filaActualOffset++;
  datos.push([
    'Placas',
    resumenGlobalCostos.placas.todoColor.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Cantos',
    resumenGlobalCostos.canto.todoColor_2.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'Herrajes',
    resumenGlobalCostos.herrajes.total.toFixed(2).replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push([
    'COSTO TOTAL GENERAL',
    (
      resumenGlobalCostos.placas.todoColor +
      resumenGlobalCostos.canto.todoColor_2 +
      resumenGlobalCostos.herrajes.total
    )
      .toFixed(2)
      .replace('.', ','),
  ]);
  filaActualOffset++;
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset++;

  // --- DETALLE DE HERRAJES TOTALES (ÚNICA SECCIÓN AL FINAL) ---
  datos.push(['DETALLE DE HERRAJES TOTALES', '', '', '', '', '', '', '']);
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset += 2;

  datos.push([
    'Código',
    'Descripción',
    'Cantidad Total',
    'Precio Unitario ($)',
    'Costo Total ($)',
  ]);
  filaActualOffset += 2;

  resumenGlobalCostos.herrajes.detalle.forEach((herraje) => {
    datos.push([
      herraje.codigo,
      herraje.descripcion,
      herraje.cantidad,
      herraje.precioUnitario.toFixed(2).replace('.', ','),
      herraje.costoTotal.toFixed(2).replace('.', ','),
    ]);
    filaActualOffset++;
  });
  datos.push(['', '', '', '', '', '', '', '']);
  filaActualOffset++;

  // Escribir datos en la hoja
  if (datos.length > 0) {
    const range = hojaSalida.getRange(1, 1, datos.length, 8);
    const values = [];
    for (let r = 0; r < datos.length; r++) {
      const rowValues = [];
      for (let c = 0; c < 8; c++) {
        rowValues.push(datos[r][c] || ''); // Asegura que no haya 'undefined'
      }
      values.push(rowValues);
    }
    range.setValues(values);

    // Formatear encabezados y secciones
    hojaSalida
      .getRange(1, 1, 1, 8)
      .setFontWeight('bold')
      .setFontSize(14)
      .setBackground('#4285F4')
      .setFontColor('white');

    for (let i = 0; i < datos.length; i++) {
      const fila = datos[i];
      if (fila[0] && fila[0].toString().includes('MÓDULO:')) {
        hojaSalida
          .getRange(i + 1, 1, 1, 8)
          .setFontWeight('bold')
          .setBackground('#FFF3E0');
      }
      if (
        fila[0] &&
        (fila[0].toString().includes('RESUMEN COMBINACIONES POR MÓDULO') ||
          fila[0].toString().includes('DETALLE DE PLACAS POR MÓDULO') ||
          fila[0].toString().includes('DETALLE DE CANTOS POR MÓDULO') ||
          fila[0].toString().includes('DETALLE DE HERRAJES POR MÓDULO'))
      ) {
        hojaSalida
          .getRange(i + 1, 1, 1, 8)
          .setFontWeight('bold')
          .setBackground('#D1E8F7');
        if (i + 2 <= datos.length) {
          const numCols =
            fila[0].toString().includes('RESUMEN COMBINACIONES POR MÓDULO') ||
            fila[0].toString().includes('DETALLE DE HERRAJES POR MÓDULO')
              ? 5
              : fila[0].toString().includes('DETALLE DE PLACAS POR MÓDULO')
              ? 4
              : 2; // Para cantos
          hojaSalida
            .getRange(i + 2, 1, 1, numCols)
            .setFontWeight('bold')
            .setBackground('#E8F0FE');
        }
      }
      if (
        fila[0] &&
        fila[0].toString().includes('INVENTARIO DE MATERIALES TOTALES')
      ) {
        hojaSalida
          .getRange(i + 1, 1, 1, 8)
          .setFontWeight('bold')
          .setFontSize(12)
          .setBackground('#A9D6E5');
        // Formatear los encabezados de las tablas de inventario
        if (i + 3 <= datos.length) {
          // Placas Necesarias Totales
          hojaSalida
            .getRange(i + 3, 1, 1, 4)
            .setFontWeight('bold')
            .setBackground('#E8F0FE');
        }
        if (i + 6 <= datos.length) {
          // Cantos Necesarios Totales (aproximado)
          hojaSalida
            .getRange(i + 6, 1, 1, 2)
            .setFontWeight('bold')
            .setBackground('#E8F0FE');
        }
      }
      if (
        fila[0] &&
        fila[0].toString().includes('PRECIOS UNITARIOS DE REFERENCIA')
      ) {
        hojaSalida
          .getRange(i + 1, 1, 1, 8)
          .setFontWeight('bold')
          .setFontSize(12)
          .setBackground('#A9D6E5');
        // Formatear los encabezados de las tablas de precios
        if (i + 3 <= datos.length) {
          // Precios de Placas
          hojaSalida
            .getRange(i + 3, 1, 1, 4)
            .setFontWeight('bold')
            .setBackground('#E8F0FE');
        }
        if (i + 6 <= datos.length) {
          // Precios de Cantos (aproximado)
          hojaSalida
            .getRange(i + 6, 1, 1, 3)
            .setFontWeight('bold')
            .setBackground('#E8F0FE');
        }
      }
      if (fila[0] && fila[0].toString().includes('RESUMEN GENERAL:')) {
        hojaSalida
          .getRange(i + 1, 1, 1, 8)
          .setFontWeight('bold')
          .setFontSize(12)
          .setBackground('#89CFF0');
        hojaSalida
          .getRange(i + 3, 1, 1, 2)
          .setFontWeight('bold')
          .setBackground('#E8F0FE');
        hojaSalida
          .getRange(i + 7, 1, 1, 2)
          .setFontWeight('bold')
          .setBackground('#A0D9F5'); // Total final
      }
      if (
        fila[0] &&
        fila[0].toString().includes('DETALLE DE HERRAJES TOTALES')
      ) {
        hojaSalida
          .getRange(i + 1, 1, 1, 8)
          .setFontWeight('bold')
          .setFontSize(12)
          .setBackground('#CDE7F5');
        hojaSalida
          .getRange(i + 3, 1, 1, 5)
          .setFontWeight('bold')
          .setBackground('#E8F0FE');
      }
      if (
        fila[0] &&
        (fila[0].toString().includes('PIEZAS OCULTAS') ||
          fila[0].toString().includes('PIEZAS VISIBLES'))
      ) {
        hojaSalida.getRange(i + 1, 1).setFontWeight('bold');
      }
    }

    // Ajustar anchos de columna para mejor visualización
    hojaSalida.autoResizeColumns(1, 8);
  }
}

/**
 * Carga los datos de herrajes (código, precio y descripción)
 */
function cargarHerrajes(ss) {
  const hojaHerrajes = ss.getSheetByName('Herrajes');
  if (!hojaHerrajes) {
    Logger.log('Hoja "Herrajes" no encontrada');
    return {};
  }

  const datos = hojaHerrajes
    .getRange(2, 1, hojaHerrajes.getLastRow() - 1, 3)
    .getValues(); // Include description column
  const herrajes = {};

  datos.forEach(([codigo, precio, descripcion]) => {
    if (codigo) {
      herrajes[codigo.toString()] = {
        precio: parseFloat(precio) || 0,
        descripcion: descripcion ? descripcion.toString() : 'Sin descripción',
      };
    }
  });

  return herrajes;
}

/**
 * Carga la configuración de herrajes por pieza
 */
function cargarConfigHerrajes(ss) {
  const hojaConfig = ss.getSheetByName('ConfigHerrajes');
  if (!hojaConfig) {
    Logger.log('Hoja "ConfigHerrajes" no encontrada');
    return [];
  }

  const numFilas = hojaConfig.getLastRow() - 1;
  if (numFilas <= 0) return [];

  return hojaConfig.getRange(2, 1, numFilas, 5).getValues();
}

/**
 * Calcula los herrajes necesarios para una pieza específica
 */
function calcularHerrajesPorPieza(
  modulo,
  parte,
  cantidadPiezas,
  variables,
  configHerrajes,
  herrajes
) {
  const herrajesCalculados = [];

  configHerrajes.forEach(
    ([
      confModulo,
      confParte,
      codigoHerraje,
      cantidadPorPieza,
      tipoVariable,
    ]) => {
      if (!confModulo || !confParte || !codigoHerraje) return;

      const confModuloLimpio = confModulo.toString().replace(/\s/g, '-');
      const confParteLimpia = confParte.toString().replace(/\s/g, '-');

      if (confModuloLimpio !== modulo || confParteLimpia !== parte) return;

      let cantidadFinal = parseFloat(cantidadPorPieza) || 0;

      // Si depende de una variable
      if (tipoVariable && variables[tipoVariable.toString().toLowerCase()]) {
        cantidadFinal = variables[tipoVariable.toString().toLowerCase()];
      }

      const herraje = herrajes[codigoHerraje.toString()];
      if (herraje) {
        const cantidadTotal = cantidadFinal * cantidadPiezas;
        const costoTotal = herraje.precio * cantidadTotal;

        herrajesCalculados.push({
          codigo: codigoHerraje.toString(),
          descripcion: herraje.descripcion,
          cantidad: cantidadTotal,
          precioUnitario: herraje.precio,
          costoTotal: costoTotal,
        });
      }
    }
  );

  return herrajesCalculados;
}

/**
 * Escribe el resumen en las hojas de entrada
 */
function escribirResumenEnHojasEntrada(ss, hojasEntrada, resumenPorModulo) {
  hojasEntrada.forEach((nombreHoja) => {
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) return;

    const numFilas = hoja.getLastRow();
    if (numFilas < 2) return;

    // Ajustado el rango para leer las columnas existentes
    const datos = hoja.getRange(2, 1, numFilas - 1, 8).getValues();

    datos.forEach((fila, index) => {
      const [modulo, altura, ancho, profundidad, cantidadModulos] = fila;
      if (!modulo || !altura || !ancho || !profundidad || !cantidadModulos)
        return;

      const moduloSinEspacios = modulo.toString().replace(/\s/g, '-');
      const moduloID = `${moduloSinEspacios}-${altura}x${ancho}x${profundidad}`;

      if (resumenPorModulo[moduloID]) {
        const filaActual = index + 2;
        const resumen = resumenPorModulo[moduloID];

        // Columna O: Metros lineales de canto (total blanco + color)
        const totalMetrosCanto =
          resumen.metrosLinealesCanto.blanco +
          resumen.metrosLinealesCanto.color;
        hoja
          .getRange(filaActual, 15)
          .setValue(totalMetrosCanto.toFixed(2).replace('.', ','));

        // Columna P: Costo total del módulo (usando "Todo Blanco" por defecto para esta columna)
        const costoFormateado = resumen.costoTotal.toFixed(2).replace('.', ',');
        hoja.getRange(filaActual, 16).setValue(costoFormateado);
      }
    });
  });
}

// Mantener las funciones de cálculo existentes
function calcularMetrosLinealesCanto(
  configuracionCanto,
  largo,
  ancho,
  cantidad
) {
  let metrosLineales = 0;
  const config = configuracionCanto.toUpperCase();

  const largoM = largo / 1000;
  const anchoM = ancho / 1000;

  const cantidadL = (config.match(/L/g) || []).length;
  const cantidadA = (config.match(/A/g) || []).length;

  metrosLineales = (cantidadL * largoM + cantidadA * anchoM) * cantidad;

  return metrosLineales;
}

function combinarDatosDeHojas(
  ss,
  nombresHojas,
  filaInicio,
  columnaInicio,
  numColumnas
) {
  const datosCombinados = [];
  nombresHojas.forEach((nombreHoja) => {
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) return;
    const numFilas = hoja.getLastRow() - filaInicio + 1;
    if (numFilas > 0) {
      const datos = hoja
        .getRange(filaInicio, columnaInicio, numFilas, numColumnas)
        .getValues();
      datosCombinados.push(...datos);
    }
  });
  return datosCombinados;
}

function calcularDimensiones(
  altura,
  ancho,
  profundidad,
  largoRaw,
  anchoRaw,
  posicion,
  offsetLargo,
  offsetAncho
) {
  let largo, anchoCalculado;

  const alturaMs = altura * 10;
  const anchoMm = ancho * 10;
  const profundidadMm = profundidad * 10;

  largo = calcularDimensionPorPosicion(
    alturaMs,
    anchoMm,
    profundidadMm,
    posicion,
    'largo',
    largoRaw
  );
  largo += offsetLargo * 10;

  anchoCalculado = calcularDimensionPorPosicion(
    alturaMs,
    anchoMm,
    profundidadMm,
    posicion,
    'ancho',
    anchoRaw
  );
  anchoCalculado += offsetAncho * 10;

  return { largo, anchoCalculado };
}

function calcularDimensionPorPosicion(
  alturaMm,
  anchoMm,
  profundidadMm,
  posicion,
  dimension,
  rawValor
) {
  const valorNumerico = parseFloat(rawValor);
  if (!isNaN(valorNumerico)) {
    return valorNumerico * 10;
  }

  if (!posicion) {
    Logger.log('Posición no especificada. Usando valores por defecto.');
    return dimension === 'largo' ? anchoMm : alturaMm;
  }

  const pos = posicion.toString().toLowerCase();

  switch (pos) {
    case 'xy':
      return dimension === 'largo' ? anchoMm : alturaMm;
    case 'xz':
      return dimension === 'largo' ? anchoMm : profundidadMm;
    case 'yz':
      return dimension === 'largo' ? profundidadMm : alturaMm;
    default:
      Logger.log(`Posición no reconocida: ${pos}. Usando XY por defecto.`);
      return dimension === 'largo' ? anchoMm : alturaMm;
  }
}
