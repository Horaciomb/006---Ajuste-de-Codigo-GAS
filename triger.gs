// ============================================================
// TRIGGERS - Versión 2.4
//
// Funciones invocadas por disparadores automáticos de Google Sheets.
// No usan UI (sin alert/toast) — los resultados se escriben en Logger.
// Misma lógica que original.gs v2.4:
//   - Llave: CODIGO + C.I. (única en fuente, soporta duplicados en sheet)
//   - Actualiza TODOS los campos automáticos (A–D, E–L, N, O)
//   - Conserva columnas manuales: M, P, Q, R, S, T
//   - Estado refleja directamente la fuente (sin condición de fechaSalida)
//   - Escritura en bloque (un solo setValues por hoja)
// ============================================================

// ── Entrada principal del trigger (ejecuta las tres hojas) ──

function triggerActualizarTodoContratos() {
  var inicio = new Date();
  var resultados = [];
  var errores = [];

  try {
    resultados.push(actualizarYAPE_SinUI_());
  } catch (e) {
    errores.push('YAPE ERROR: ' + e.message);
  }

  try {
    resultados.push(actualizarBNB_SinUI_());
  } catch (e) {
    errores.push('BNB ERROR: ' + e.message);
  }

  try {
    resultados.push(actualizarZAS_SinUI_());
  } catch (e) {
    errores.push('ZAS ERROR: ' + e.message);
  }

  var fin      = new Date();
  var duracion = ((fin - inicio) / 1000).toFixed(2);

  var mensajeFinal =
    '=== ACTUALIZACIÓN CONTRATOS ===\n\n' +
    resultados.join('\n') +
    (errores.length ? '\n\n=== ERRORES ===\n' + errores.join('\n') : '') +
    '\n\nDuración: ' + duracion + ' seg';

  Logger.log(mensajeFinal);
  return mensajeFinal;
}

// ── Entradas individuales por hoja ──

function triggerActualizarYAPE() { return actualizarYAPE_SinUI_(); }
function triggerActualizarBNB()  { return actualizarBNB_SinUI_();  }
function triggerActualizarZAS()  { return actualizarZAS_SinUI_();  }

// ============================================================
// YAPE - SinUI v2.4
// ============================================================

function actualizarYAPE_SinUI_() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDestino = ss.getSheetByName('CONTRATOS YAPE');
  var TOTAL_COLUMNAS   = 20;
  var COL_FECHA_INGRESO = 14; // columna N, 1-based

  if (!hojaDestino) {
    Logger.log('YAPE: No se encontró la hoja CONTRATOS YAPE');
    return 'YAPE: hoja no encontrada';
  }

  var datosOrigen;
  try {
    datosOrigen = descargarCSV(CONFIG.YAPE.url);
  } catch (e) {
    Logger.log('YAPE error descarga: ' + e.message);
    return 'YAPE: error descarga - ' + e.message;
  }

  if (datosOrigen.length < 2) return 'YAPE: sin datos en origen';
  datosOrigen.shift();

  // ── 1. Leer destino completo en memoria ──
  var ultimaFila    = hojaDestino.getLastRow();
  var datosDestino  = [];
  var mapaDestino   = {}; // 'CODIGO|CI' → [idx0, idx1, ...]
  var clavesDestino = new Set();

  if (ultimaFila > 1) {
    datosDestino = hojaDestino
      .getRange(2, 1, ultimaFila - 1, TOTAL_COLUMNAS)
      .getValues();

    datosDestino.forEach(function(fila, idx) {
      var codigo = String(fila[1] || '').trim();
      var ci     = String(fila[2] || '').trim();
      if (codigo === '' || ci === '') return;
      var key = codigo + '|' + ci;
      clavesDestino.add(key);
      if (!mapaDestino[key]) mapaDestino[key] = [];
      mapaDestino[key].push(idx);
    });
  }

  // ── 2. Procesar fuente ──
  var registrosNuevos          = [];
  var registrosActualizados    = 0;
  var registrosNuevosAgregados = 0;
  var estadosActualizados      = 0;

  datosOrigen.forEach(function(filaOrigen) {
    var codigo       = String(filaOrigen[MAPEO_YAPE.codigo] || '').trim();
    var ci           = String(filaOrigen[MAPEO_YAPE.ci]     || '').trim();
    var estadoOrigen = String(filaOrigen[MAPEO_YAPE.estado] || '').trim().toUpperCase();
    var ciudad       = filaOrigen[MAPEO_YAPE.ciudad]         || '';

    if (codigo === '' || ci === '') return;

    var key = codigo + '|' + ci;

    // Estado espejo directo de la fuente
    var estadoFinal = estadoOrigen !== '' ? estadoOrigen : 'ACTIVO';

    if (clavesDestino.has(key)) {
      mapaDestino[key].forEach(function(idx) {
        var fa = datosDestino[idx]; // fila actual

        var estadoActual = String(fa[14] || '').toUpperCase().trim();
        if (estadoActual !== 'INACTIVO' && estadoFinal === 'INACTIVO') estadosActualizados++;

        datosDestino[idx] = [
          filaOrigen[MAPEO_YAPE.empresa]                    || '',  // A [0]
          codigo,                                                    // B [1]
          ci,                                                        // C [2]
          filaOrigen[MAPEO_YAPE.celular]                    || '',  // D [3]
          aNombrePropio(filaOrigen[MAPEO_YAPE.nombreCompleto]),     // E [4]
          aNombrePropio(filaOrigen[MAPEO_YAPE.apellidoPaterno]),    // F [5]
          aNombrePropio(filaOrigen[MAPEO_YAPE.apellidoMaterno]),    // G [6]
          aNombrePropio(filaOrigen[MAPEO_YAPE.nombres]),            // H [7]
          filaOrigen[MAPEO_YAPE.cargo]                      || '',  // I [8]
          ciudad,                                                    // J [9]
          obtenerDepartamento(ciudad),                               // K [10]
          aNombrePropio(filaOrigen[MAPEO_YAPE.supervisor]),         // L [11]
          fa[12],         // M  FECHA DE CONTRATO  ← manual
          parseFechaYAPE(filaOrigen[MAPEO_YAPE.fechaIngreso]),      // N [13]
          estadoFinal,                                               // O [14]
          fa[15],         // P  FIRMA              ← manual
          fa[16],         // Q  ESTADO DE CONTRATO ← manual
          fa[17],         // R  OBS                ← manual
          fa[18],         // S  OBS SUPERVISOR     ← manual
          fa[19]          // T  OBS SUBGERENTE     ← manual
        ];
        registrosActualizados++;
      });

    } else {
      if (estadoOrigen !== 'ACTIVO') return;

      registrosNuevos.push([
        filaOrigen[MAPEO_YAPE.empresa]                    || '',
        codigo, ci,
        filaOrigen[MAPEO_YAPE.celular]                    || '',
        aNombrePropio(filaOrigen[MAPEO_YAPE.nombreCompleto]),
        aNombrePropio(filaOrigen[MAPEO_YAPE.apellidoPaterno]),
        aNombrePropio(filaOrigen[MAPEO_YAPE.apellidoMaterno]),
        aNombrePropio(filaOrigen[MAPEO_YAPE.nombres]),
        filaOrigen[MAPEO_YAPE.cargo]                      || '',
        ciudad,
        obtenerDepartamento(ciudad),
        aNombrePropio(filaOrigen[MAPEO_YAPE.supervisor]),
        '',                                                          // M manual
        parseFechaYAPE(filaOrigen[MAPEO_YAPE.fechaIngreso]),
        'ACTIVO',
        '', '', '', '', ''                                           // P-T manuales
      ]);
      clavesDestino.add(key);
      registrosNuevosAgregados++;
    }
  });

  // ── 3. Escribir existentes en bloque ──
  if (datosDestino.length > 0) {
    hojaDestino.getRange(2, 1, datosDestino.length, TOTAL_COLUMNAS).setValues(datosDestino);
    hojaDestino.getRange(2, COL_FECHA_INGRESO, datosDestino.length, 1).setNumberFormat('@');
  }

  // ── 4. Agregar nuevos en bloque ──
  if (registrosNuevos.length > 0) {
    var filaInicio = hojaDestino.getLastRow() + 1;
    hojaDestino.getRange(filaInicio, 1, registrosNuevos.length, TOTAL_COLUMNAS).setValues(registrosNuevos);
    hojaDestino.getRange(filaInicio, COL_FECHA_INGRESO, registrosNuevos.length, 1).setNumberFormat('@');
  }

  return 'YAPE → nuevos: ' + registrosNuevosAgregados +
         ' | actualizados: ' + registrosActualizados +
         ' | a INACTIVO: ' + estadosActualizados;
}

// ============================================================
// BNB - SinUI v2.4
// ============================================================

function actualizarBNB_SinUI_() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDestino = ss.getSheetByName('CONTRATOS BNB');
  var TOTAL_COLUMNAS    = 20;
  var COL_FECHA_INGRESO = 14; // columna N, 1-based

  if (!hojaDestino) {
    Logger.log('BNB: No se encontró la hoja CONTRATOS BNB');
    return 'BNB: hoja no encontrada';
  }

  var datosOrigen;
  try {
    datosOrigen = descargarCSV(CONFIG.BNB.url);
  } catch (e) {
    Logger.log('BNB error descarga: ' + e.message);
    return 'BNB: error descarga - ' + e.message;
  }

  if (datosOrigen.length < 2) return 'BNB: sin datos en origen';
  datosOrigen.shift();

  // ── 1. Leer destino completo en memoria ──
  var ultimaFila    = hojaDestino.getLastRow();
  var datosDestino  = [];
  var mapaDestino   = {}; // 'CODIGO|CI' → [idx0, idx1, ...]
  var clavesDestino = new Set();

  if (ultimaFila > 1) {
    datosDestino = hojaDestino
      .getRange(2, 1, ultimaFila - 1, TOTAL_COLUMNAS)
      .getValues();

    datosDestino.forEach(function(fila, idx) {
      var codigo = String(fila[1] || '').trim();
      var ci     = String(fila[2] || '').trim();
      if (codigo === '' || ci === '') return;
      var key = codigo + '|' + ci;
      clavesDestino.add(key);
      if (!mapaDestino[key]) mapaDestino[key] = [];
      mapaDestino[key].push(idx);
    });
  }

  // ── 2. Procesar fuente ──
  var registrosNuevos          = [];
  var registrosActualizados    = 0;
  var registrosNuevosAgregados = 0;
  var estadosActualizados      = 0;

  datosOrigen.forEach(function(filaOrigen) {
    var codigo       = String(filaOrigen[MAPEO_BNB.codigo] || '').trim();
    var ci           = String(filaOrigen[MAPEO_BNB.ci]     || '').trim();
    var estadoOrigen = String(filaOrigen[MAPEO_BNB.estado] || '').trim().toUpperCase();
    var ciudad       = filaOrigen[MAPEO_BNB.ciudad]         || '';

    if (codigo === '' || ci === '') return;

    var key = codigo + '|' + ci;

    // Estado espejo directo de la fuente
    var estadoFinal = estadoOrigen !== '' ? estadoOrigen : 'ACTIVO';

    if (clavesDestino.has(key)) {
      mapaDestino[key].forEach(function(idx) {
        var fa = datosDestino[idx];

        var estadoActual = String(fa[14] || '').toUpperCase().trim();
        if (estadoActual !== 'INACTIVO' && estadoFinal === 'INACTIVO') estadosActualizados++;

        datosDestino[idx] = [
          filaOrigen[MAPEO_BNB.empresa]                    || '',  // A [0]
          codigo,                                                    // B [1]
          ci,                                                        // C [2]
          filaOrigen[MAPEO_BNB.celular]                    || '',  // D [3]
          aNombrePropio(filaOrigen[MAPEO_BNB.nombreCompleto]),     // E [4]
          aNombrePropio(filaOrigen[MAPEO_BNB.apellidoPaterno]),    // F [5]
          aNombrePropio(filaOrigen[MAPEO_BNB.apellidoMaterno]),    // G [6]
          aNombrePropio(filaOrigen[MAPEO_BNB.nombres]),            // H [7]
          filaOrigen[MAPEO_BNB.cargo]                      || '',  // I [8]
          ciudad,                                                    // J [9]
          obtenerDepartamento(ciudad),                               // K [10]
          aNombrePropio(filaOrigen[MAPEO_BNB.supervisor]),         // L [11]
          fa[12],         // M  FECHA DE CONTRATO  ← manual
          parseFechaYAPE(filaOrigen[MAPEO_BNB.fechaIngreso]),      // N [13]
          estadoFinal,                                               // O [14]
          fa[15],         // P  FIRMA              ← manual
          fa[16],         // Q  ESTADO DE CONTRATO ← manual
          fa[17],         // R  OBS                ← manual
          fa[18],         // S  OBS SUPERVISOR     ← manual
          fa[19]          // T  OBS SUBGERENTE     ← manual
        ];
        registrosActualizados++;
      });

    } else {
      if (estadoOrigen !== 'ACTIVO') return;

      registrosNuevos.push([
        filaOrigen[MAPEO_BNB.empresa]                    || '',
        codigo, ci,
        filaOrigen[MAPEO_BNB.celular]                    || '',
        aNombrePropio(filaOrigen[MAPEO_BNB.nombreCompleto]),
        aNombrePropio(filaOrigen[MAPEO_BNB.apellidoPaterno]),
        aNombrePropio(filaOrigen[MAPEO_BNB.apellidoMaterno]),
        aNombrePropio(filaOrigen[MAPEO_BNB.nombres]),
        filaOrigen[MAPEO_BNB.cargo]                      || '',
        ciudad,
        obtenerDepartamento(ciudad),
        aNombrePropio(filaOrigen[MAPEO_BNB.supervisor]),
        '',                                                          // M manual
        parseFechaYAPE(filaOrigen[MAPEO_BNB.fechaIngreso]),
        'ACTIVO',
        '', '', '', '', ''                                           // P-T manuales
      ]);
      clavesDestino.add(key);
      registrosNuevosAgregados++;
    }
  });

  // ── 3. Escribir existentes en bloque ──
  if (datosDestino.length > 0) {
    hojaDestino.getRange(2, 1, datosDestino.length, TOTAL_COLUMNAS).setValues(datosDestino);
    hojaDestino.getRange(2, COL_FECHA_INGRESO, datosDestino.length, 1).setNumberFormat('@');
  }

  // ── 4. Agregar nuevos en bloque ──
  if (registrosNuevos.length > 0) {
    var filaInicio = hojaDestino.getLastRow() + 1;
    hojaDestino.getRange(filaInicio, 1, registrosNuevos.length, TOTAL_COLUMNAS).setValues(registrosNuevos);
    hojaDestino.getRange(filaInicio, COL_FECHA_INGRESO, registrosNuevos.length, 1).setNumberFormat('@');
  }

  return 'BNB → nuevos: ' + registrosNuevosAgregados +
         ' | actualizados: ' + registrosActualizados +
         ' | a INACTIVO: ' + estadosActualizados;
}

// ============================================================
// ZAS - SinUI v2.4
// ============================================================

function actualizarZAS_SinUI_() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDestino = ss.getSheetByName('CONTRATOS ZAS');
  var TOTAL_COLUMNAS    = 20;
  var COL_FECHA_INGRESO = 14; // columna N, 1-based

  if (!hojaDestino) {
    Logger.log('ZAS: No se encontró la hoja CONTRATOS ZAS');
    return 'ZAS: hoja no encontrada';
  }

  var datosOrigen;
  var usuariosExtraterritoriales = obtenerUsuariosExtraterritorialesZAS();
  var omitidosExtraterritorial   = 0;

  try {
    datosOrigen = descargarCSV(CONFIG.ZAS.url);
  } catch (e) {
    Logger.log('ZAS error descarga: ' + e.message);
    return 'ZAS: error descarga - ' + e.message;
  }

  if (datosOrigen.length < 2) return 'ZAS: sin datos en origen';
  datosOrigen.shift();

  // ── 1. Leer destino completo en memoria ──
  var ultimaFila    = hojaDestino.getLastRow();
  var datosDestino  = [];
  var mapaDestino   = {}; // 'CODIGO|CI' → [idx0, idx1, ...]
  var clavesDestino = new Set();

  if (ultimaFila > 1) {
    datosDestino = hojaDestino
      .getRange(2, 1, ultimaFila - 1, TOTAL_COLUMNAS)
      .getValues();

    datosDestino.forEach(function(fila, idx) {
      var codigo = String(fila[1] || '').trim();
      var ci     = String(fila[2] || '').trim();
      if (codigo === '') return;
      var key = codigo + '|' + ci;
      clavesDestino.add(key);
      if (!mapaDestino[key]) mapaDestino[key] = [];
      mapaDestino[key].push(idx);
    });
  }

  // ── 2. Procesar fuente ──
  var registrosNuevos          = [];
  var registrosActualizados    = 0;
  var registrosNuevosAgregados = 0;
  var estadosActualizados      = 0;

  datosOrigen.forEach(function(filaOrigen) {
    var codigo       = String(filaOrigen[MAPEO_ZAS.codigo]  || '').trim();
    var ci           = String(filaOrigen[MAPEO_ZAS.ci]      || '').trim();
    var estadoOrigen = String(filaOrigen[MAPEO_ZAS.estado]  || '').trim().toUpperCase();
    var usuario      = String(filaOrigen[MAPEO_ZAS.usuario] || '').trim().toUpperCase();

    var ciudad  = filaOrigen[MAPEO_ZAS.ciudad_txt] || filaOrigen[MAPEO_ZAS.ciudad] || '';
    var dptoRaw = filaOrigen[MAPEO_ZAS.dpto] || '';
    var dpto    = estaVacio(dptoRaw)
      ? obtenerDepartamento(ciudad)
      : (obtenerDepartamento(dptoRaw) || dptoRaw);

    if (codigo === '') return;

    var key = codigo + '|' + ci;

    // Estado espejo directo de la fuente
    var estadoFinal = estadoOrigen !== '' ? estadoOrigen : 'ACTIVO';

    if (clavesDestino.has(key)) {
      mapaDestino[key].forEach(function(idx) {
        var fa = datosDestino[idx];

        var estadoActual = String(fa[14] || '').toUpperCase().trim();
        if (estadoActual !== 'INACTIVO' && estadoFinal === 'INACTIVO') estadosActualizados++;

        datosDestino[idx] = [
          filaOrigen[MAPEO_ZAS.empresa] || 'ZAS',                    // A [0]
          codigo,                                                      // B [1]
          ci,                                                          // C [2]
          filaOrigen[MAPEO_ZAS.celular]                      || '',  // D [3]
          aNombrePropio(filaOrigen[MAPEO_ZAS.nombreCompleto]),       // E [4]
          aNombrePropio(filaOrigen[MAPEO_ZAS.apellidoPaterno]),      // F [5]
          aNombrePropio(filaOrigen[MAPEO_ZAS.apellidoMaterno]),      // G [6]
          aNombrePropio(filaOrigen[MAPEO_ZAS.nombres]),              // H [7]
          filaOrigen[MAPEO_ZAS.cargo]                        || '',  // I [8]
          ciudad,                                                      // J [9]
          dpto,                                                        // K [10]
          aNombrePropio(filaOrigen[MAPEO_ZAS.supervisor]),           // L [11]
          fa[12],         // M  FECHA DE CONTRATO  ← manual
          parseFechaZASComoDate(filaOrigen[MAPEO_ZAS.fechaIngreso]), // N [13]
          estadoFinal,                                                 // O [14]
          fa[15],         // P  FIRMA              ← manual
          fa[16],         // Q  ESTADO DE CONTRATO ← manual
          fa[17],         // R  OBS                ← manual
          fa[18],         // S  OBS SUPERVISOR     ← manual
          fa[19]          // T  OBS SUBGERENTE     ← manual
        ];
        registrosActualizados++;
      });

    } else {
      // Solo insertar ACTIVOS; omitir extraterritoriales
      if (estadoOrigen !== 'ACTIVO') return;
      if (usuario !== '' && usuariosExtraterritoriales.has(usuario)) {
        omitidosExtraterritorial++;
        return;
      }

      registrosNuevos.push([
        filaOrigen[MAPEO_ZAS.empresa] || 'ZAS',
        codigo, ci,
        filaOrigen[MAPEO_ZAS.celular]                      || '',
        aNombrePropio(filaOrigen[MAPEO_ZAS.nombreCompleto]),
        aNombrePropio(filaOrigen[MAPEO_ZAS.apellidoPaterno]),
        aNombrePropio(filaOrigen[MAPEO_ZAS.apellidoMaterno]),
        aNombrePropio(filaOrigen[MAPEO_ZAS.nombres]),
        filaOrigen[MAPEO_ZAS.cargo]                        || '',
        ciudad, dpto,
        aNombrePropio(filaOrigen[MAPEO_ZAS.supervisor]),
        '',                                                            // M manual
        parseFechaZASComoDate(filaOrigen[MAPEO_ZAS.fechaIngreso]),
        'ACTIVO',
        '', '', '', '', ''                                             // P-T manuales
      ]);
      clavesDestino.add(key);
      registrosNuevosAgregados++;
    }
  });

  // ── 3. Escribir existentes en bloque ──
  if (datosDestino.length > 0) {
    hojaDestino.getRange(2, 1, datosDestino.length, TOTAL_COLUMNAS).setValues(datosDestino);
    hojaDestino.getRange(2, COL_FECHA_INGRESO, datosDestino.length, 1).setNumberFormat('dd/mm/yyyy');
  }

  // ── 4. Agregar nuevos en bloque ──
  if (registrosNuevos.length > 0) {
    var filaInicio = hojaDestino.getLastRow() + 1;
    hojaDestino.getRange(filaInicio, 1, registrosNuevos.length, TOTAL_COLUMNAS).setValues(registrosNuevos);
    hojaDestino.getRange(filaInicio, COL_FECHA_INGRESO, registrosNuevos.length, 1).setNumberFormat('dd/mm/yyyy');
  }

  return 'ZAS → nuevos: ' + registrosNuevosAgregados +
         ' | actualizados: ' + registrosActualizados +
         ' | a INACTIVO: ' + estadosActualizados +
         ' | omitidos extraterritorial: ' + omitidosExtraterritorial;
}
