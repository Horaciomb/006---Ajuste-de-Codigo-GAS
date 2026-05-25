// ============================================================
// ACTUALIZAR YAPE - v2.2 (llave única por CODIGO)
// ============================================================

function actualizarYAPE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDestino = ss.getSheetByName('CONTRATOS YAPE');

  var TOTAL_COLUMNAS_YAPE = 20;
  var COL_FECHA_INGRESO = 14; // N (1-based)

  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert('Error: No se encontró la hoja "CONTRATOS YAPE"');
    return;
  }

  var encabezado = hojaDestino.getRange(1, 1, 1, hojaDestino.getLastColumn()).getValues()[0];
  if (encabezado[0] !== 'EMPRESA' || encabezado[2] !== 'C.I.' || encabezado[3] !== 'CELULAR') {
    SpreadsheetApp.getUi().alert('Error: La hoja CONTRATOS YAPE no está estandarizada.');
    return;
  }

  var datosOrigen;
  try {
    ss.toast('Descargando datos de YAPE...', 'Procesando', -1);
    datosOrigen = descargarCSV(CONFIG.YAPE.url);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error al descargar datos de YAPE:\n' + e.message);
    return;
  }

  if (datosOrigen.length < 2) {
    SpreadsheetApp.getUi().alert('No se encontraron datos en la fuente YAPE');
    return;
  }

  datosOrigen.shift(); // quitar encabezado

  // ── 1. Leer destino: índice por CODIGO (col B = índice 1 del array) ──
  var ultimaFilaDestino = hojaDestino.getLastRow();
  var mapaDestino = {};      // codigo -> { filaReal, datos[] }
  var codigosDestino = new Set();

  if (ultimaFilaDestino > 1) {
    var datosDestino = hojaDestino
      .getRange(2, 1, ultimaFilaDestino - 1, TOTAL_COLUMNAS_YAPE)
      .getValues();

    datosDestino.forEach(function(fila, index) {
      var codigo = String(fila[1] || '').trim(); // col B
      if (codigo !== '') {
        codigosDestino.add(codigo);
        mapaDestino[codigo] = {
          filaReal: index + 2,
          datos: fila
        };
      }
    });
  }

  // ── 2. Procesar origen ──
  var registrosNuevos = [];
  var registrosActualizados = 0;
  var registrosNuevosAgregados = 0;
  var estadosActualizados = 0;

  // Acumular cambios en lote para no hacer getRange celda por celda
  var loteActualizaciones = []; // { filaReal, nuevaFila[] }

  datosOrigen.forEach(function(filaOrigen) {
    var codigo       = String(filaOrigen[MAPEO_YAPE.codigo]   || '').trim();
    var estadoRaw    = String(filaOrigen[MAPEO_YAPE.estado]   || '').trim();
    var estadoOrigen = estadoRaw.toUpperCase();
    var fechaSalida  = String(filaOrigen[MAPEO_YAPE.fechaSalida] || '').trim();
    var ciudad       = filaOrigen[MAPEO_YAPE.ciudad] || '';

    if (codigo === '') return; // sin clave, ignorar

    var estadoFinal = (estadoOrigen === 'INACTIVO' && fechaSalida !== '')
      ? 'INACTIVO'
      : 'ACTIVO';

    if (codigosDestino.has(codigo)) {
      // ── REGISTRO EXISTENTE: actualizar columnas automáticas, conservar manuales ──
      var destino  = mapaDestino[codigo];
      var filaActual = destino.datos;

      var nuevaFila = [
        filaOrigen[MAPEO_YAPE.empresa]        || '',                          // A  col 0
        codigo,                                                                // B  col 1
        filaOrigen[MAPEO_YAPE.ci]             || '',                          // C  col 2
        filaOrigen[MAPEO_YAPE.celular]        || '',                          // D  col 3  ← se actualiza
        aNombrePropio(filaOrigen[MAPEO_YAPE.nombreCompleto]),                 // E  col 4
        aNombrePropio(filaOrigen[MAPEO_YAPE.apellidoPaterno]),                // F  col 5
        aNombrePropio(filaOrigen[MAPEO_YAPE.apellidoMaterno]),                // G  col 6
        aNombrePropio(filaOrigen[MAPEO_YAPE.nombres]),                        // H  col 7
        filaOrigen[MAPEO_YAPE.cargo]          || '',                          // I  col 8
        ciudad,                                                                // J  col 9
        obtenerDepartamento(ciudad),                                           // K  col 10
        aNombrePropio(filaOrigen[MAPEO_YAPE.supervisor]),                     // L  col 11
        filaActual[12],  // M  FECHA DE CONTRATO  → manual, se conserva
        parseFechaYAPE(filaOrigen[MAPEO_YAPE.fechaIngreso]),                  // N  col 13
        estadoFinal,                                                           // O  col 14
        filaActual[15],  // P  FIRMA              → manual
        filaActual[16],  // Q  ESTADO DE CONTRATO → manual
        filaActual[17],  // R  OBS                → manual
        filaActual[18],  // S  OBS SUPERVISOR     → manual
        filaActual[19]   // T  OBS SUBGERENTE     → manual
      ];

      // Detectar si el estado cambió a INACTIVO
      var estadoActual = String(filaActual[14] || '').toUpperCase().trim();
      if (estadoActual !== 'INACTIVO' && estadoFinal === 'INACTIVO') {
        estadosActualizados++;
      }

      loteActualizaciones.push({ filaReal: destino.filaReal, nuevaFila: nuevaFila });
      registrosActualizados++;

    } else {
      // ── REGISTRO NUEVO: solo insertar si está ACTIVO ──
      if (estadoOrigen !== 'ACTIVO') return;

      registrosNuevos.push([
        filaOrigen[MAPEO_YAPE.empresa]        || '',
        codigo,
        filaOrigen[MAPEO_YAPE.ci]            || '',
        filaOrigen[MAPEO_YAPE.celular]        || '',
        aNombrePropio(filaOrigen[MAPEO_YAPE.nombreCompleto]),
        aNombrePropio(filaOrigen[MAPEO_YAPE.apellidoPaterno]),
        aNombrePropio(filaOrigen[MAPEO_YAPE.apellidoMaterno]),
        aNombrePropio(filaOrigen[MAPEO_YAPE.nombres]),
        filaOrigen[MAPEO_YAPE.cargo]          || '',
        ciudad,
        obtenerDepartamento(ciudad),
        aNombrePropio(filaOrigen[MAPEO_YAPE.supervisor]),
        '',   // M  FECHA DE CONTRATO → vacío, manual
        parseFechaYAPE(filaOrigen[MAPEO_YAPE.fechaIngreso]),
        'ACTIVO',
        '',   // P  FIRMA
        '',   // Q  ESTADO DE CONTRATO
        '',   // R  OBS
        '',   // S  OBS SUPERVISOR
        ''    // T  OBS SUBGERENTE
      ]);

      codigosDestino.add(codigo);
      registrosNuevosAgregados++;
    }
  });

  // ── 3. Escribir actualizaciones en lote ──
  loteActualizaciones.forEach(function(item) {
    hojaDestino
      .getRange(item.filaReal, 1, 1, TOTAL_COLUMNAS_YAPE)
      .setValues([item.nuevaFila]);

    // Mantener columna N como texto para no perder formato de fecha
    hojaDestino
      .getRange(item.filaReal, COL_FECHA_INGRESO, 1, 1)
      .setNumberFormat('@');
  });

  // ── 4. Escribir nuevos en lote ──
  if (registrosNuevos.length > 0) {
    var filaInicio = hojaDestino.getLastRow() + 1;
    hojaDestino
      .getRange(filaInicio, 1, registrosNuevos.length, TOTAL_COLUMNAS_YAPE)
      .setValues(registrosNuevos);
    hojaDestino
      .getRange(filaInicio, COL_FECHA_INGRESO, registrosNuevos.length, 1)
      .setNumberFormat('@');
  }

  ss.toast('', '', 1);

  SpreadsheetApp.getUi().alert(
    'Actualización YAPE completada.\n\n' +
    'Registros nuevos agregados: '          + registrosNuevosAgregados + '\n' +
    'Registros existentes actualizados: '   + registrosActualizados    + '\n' +
    'Estados actualizados a INACTIVO: '     + estadosActualizados      + '\n\n' +
    'Columnas manuales conservadas: M, P, Q, R, S y T'
  );
}