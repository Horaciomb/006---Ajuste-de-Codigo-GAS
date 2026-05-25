// ============================================================
// ACTUALIZAR YAPE - v2.3
//
// Cambios respecto a versiones anteriores:
//   v2.1 usaba llave CODIGO+CELULAR → falla cuando el celular cambia
//   v2.2 usaba solo CODIGO          → CODIGO no es único (529 duplicados)
//   v2.3 usa llave CODIGO+CI        → única en fuente (0 duplicados)
//
// Qué hace esta función:
//   - Descarga el CSV de YAPE
//   - Para cada persona en fuente:
//       · Si ya existe en contratos (por CODIGO+CI): actualiza todos los
//         campos automáticos (A-D, E-L, N, O) y conserva los manuales (M, P-T)
//       · Si NO existe y está ACTIVO: la inserta con manuales en blanco
//   - Escribe existentes en un solo bloque (no celda a celda)
//   - Soporta filas duplicadas en el sheet de contratos: actualiza todas
// ============================================================

function actualizarYAPE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDestino = ss.getSheetByName('CONTRATOS YAPE');

  var TOTAL_COLUMNAS_YAPE = 20;
  var COL_FECHA_INGRESO   = 14; // columna N, 1-based

  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert('Error: No se encontró la hoja "CONTRATOS YAPE"');
    return;
  }

  // Validar que la hoja tiene la estructura esperada
  var encabezado = hojaDestino
    .getRange(1, 1, 1, hojaDestino.getLastColumn())
    .getValues()[0];
  if (encabezado[0] !== 'EMPRESA' || encabezado[2] !== 'C.I.' || encabezado[3] !== 'CELULAR') {
    SpreadsheetApp.getUi().alert('Error: La hoja CONTRATOS YAPE no está estandarizada.');
    return;
  }

  // ── Descargar fuente YAPE ──
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

  datosOrigen.shift(); // quitar fila de encabezado

  // ── 1. Leer destino completo en memoria ──
  //
  // mapaDestino: 'CODIGO|CI' → [idx0, idx1, ...]
  // Se guardan TODOS los índices porque el sheet puede tener duplicados
  // (mismo CODIGO+CI en dos filas). Se actualizan todos.
  //
  var ultimaFila    = hojaDestino.getLastRow();
  var datosDestino  = [];
  var mapaDestino   = {};
  var clavesDestino = new Set();

  if (ultimaFila > 1) {
    datosDestino = hojaDestino
      .getRange(2, 1, ultimaFila - 1, TOTAL_COLUMNAS_YAPE)
      .getValues();

    datosDestino.forEach(function(fila, idx) {
      var codigo = String(fila[1] || '').trim(); // col B
      var ci     = String(fila[2] || '').trim(); // col C
      if (codigo === '' || ci === '') return;
      var key = codigo + '|' + ci;
      clavesDestino.add(key);
      if (!mapaDestino[key]) mapaDestino[key] = [];
      mapaDestino[key].push(idx);
    });
  }

  // ── 2. Procesar cada fila de la fuente ──
  var registrosNuevos          = [];
  var registrosActualizados    = 0;
  var registrosNuevosAgregados = 0;
  var estadosActualizados      = 0;

  datosOrigen.forEach(function(filaOrigen) {
    var codigo       = String(filaOrigen[MAPEO_YAPE.codigo]      || '').trim();
    var ci           = String(filaOrigen[MAPEO_YAPE.ci]          || '').trim();
    var estadoRaw    = String(filaOrigen[MAPEO_YAPE.estado]      || '').trim();
    var estadoOrigen = estadoRaw.toUpperCase();
    var fechaSalida  = String(filaOrigen[MAPEO_YAPE.fechaSalida] || '').trim();
    var ciudad       = filaOrigen[MAPEO_YAPE.ciudad]              || '';

    if (codigo === '' || ci === '') return;

    var key = codigo + '|' + ci;

    // Estado final para columna O:
    // Solo marca INACTIVO si la fuente lo dice Y tiene fecha de salida,
    // para evitar desactivaciones accidentales por datos incompletos.
    var estadoFinal = (estadoOrigen === 'INACTIVO' && fechaSalida !== '')
      ? 'INACTIVO'
      : 'ACTIVO';

    if (clavesDestino.has(key)) {
      // ── Registro existente: actualizar columnas automáticas en memoria ──
      // Si hay duplicados en el sheet, se actualizan TODAS las filas con esa llave.
      mapaDestino[key].forEach(function(idx) {
        var filaActual = datosDestino[idx];

        // Contar cuántos pasan a INACTIVO
        var estadoActual = String(filaActual[14] || '').toUpperCase().trim();
        if (estadoActual !== 'INACTIVO' && estadoFinal === 'INACTIVO') {
          estadosActualizados++;
        }

        // Columnas automáticas: A B C D E F G H I J K L N O
        // Columnas manuales que NO se tocan: M(12) P(15) Q(16) R(17) S(18) T(19)
        datosDestino[idx] = [
          filaOrigen[MAPEO_YAPE.empresa]                    || '',  // A [0]
          codigo,                                                    // B [1]
          ci,                                                        // C [2]
          filaOrigen[MAPEO_YAPE.celular]                    || '',  // D [3]  ← se actualiza desde fuente
          aNombrePropio(filaOrigen[MAPEO_YAPE.nombreCompleto]),     // E [4]
          aNombrePropio(filaOrigen[MAPEO_YAPE.apellidoPaterno]),    // F [5]
          aNombrePropio(filaOrigen[MAPEO_YAPE.apellidoMaterno]),    // G [6]
          aNombrePropio(filaOrigen[MAPEO_YAPE.nombres]),            // H [7]
          filaOrigen[MAPEO_YAPE.cargo]                      || '',  // I [8]
          ciudad,                                                    // J [9]
          obtenerDepartamento(ciudad),                               // K [10]
          aNombrePropio(filaOrigen[MAPEO_YAPE.supervisor]),         // L [11]
          filaActual[12],  // M  FECHA DE CONTRATO  ← manual, NO tocar
          parseFechaYAPE(filaOrigen[MAPEO_YAPE.fechaIngreso]),      // N [13]
          estadoFinal,                                               // O [14]
          filaActual[15],  // P  FIRMA              ← manual
          filaActual[16],  // Q  ESTADO DE CONTRATO ← manual
          filaActual[17],  // R  OBS                ← manual
          filaActual[18],  // S  OBS SUPERVISOR     ← manual
          filaActual[19]   // T  OBS SUBGERENTE     ← manual
        ];

        registrosActualizados++;
      });

    } else {
      // ── Registro nuevo: insertar solo si está ACTIVO en la fuente ──
      // Los INACTIVOS que nunca estuvieron en contratos no se insertan
      // (ya salieron antes de ser registrados).
      if (estadoOrigen !== 'ACTIVO') return;

      registrosNuevos.push([
        filaOrigen[MAPEO_YAPE.empresa]                    || '',  // A
        codigo,                                                    // B
        ci,                                                        // C
        filaOrigen[MAPEO_YAPE.celular]                    || '',  // D
        aNombrePropio(filaOrigen[MAPEO_YAPE.nombreCompleto]),     // E
        aNombrePropio(filaOrigen[MAPEO_YAPE.apellidoPaterno]),    // F
        aNombrePropio(filaOrigen[MAPEO_YAPE.apellidoMaterno]),    // G
        aNombrePropio(filaOrigen[MAPEO_YAPE.nombres]),            // H
        filaOrigen[MAPEO_YAPE.cargo]                      || '',  // I
        ciudad,                                                    // J
        obtenerDepartamento(ciudad),                               // K
        aNombrePropio(filaOrigen[MAPEO_YAPE.supervisor]),         // L
        '',                                                        // M  FECHA DE CONTRATO → vacío, manual
        parseFechaYAPE(filaOrigen[MAPEO_YAPE.fechaIngreso]),      // N
        'ACTIVO',                                                  // O
        '',                                                        // P  FIRMA
        '',                                                        // Q  ESTADO DE CONTRATO
        '',                                                        // R  OBS
        '',                                                        // S  OBS SUPERVISOR
        ''                                                         // T  OBS SUBGERENTE
      ]);

      clavesDestino.add(key);
      registrosNuevosAgregados++;
    }
  });

  // ── 3. Escribir todos los registros existentes en un único bloque ──
  // Una sola llamada setValues en vez de N llamadas individuales.
  if (datosDestino.length > 0) {
    hojaDestino
      .getRange(2, 1, datosDestino.length, TOTAL_COLUMNAS_YAPE)
      .setValues(datosDestino);

    // Forzar texto en columna N (FECHA DE INGRESO) para que no se convierta a número
    hojaDestino
      .getRange(2, COL_FECHA_INGRESO, datosDestino.length, 1)
      .setNumberFormat('@');
  }

  // ── 4. Agregar registros nuevos en un único bloque ──
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
