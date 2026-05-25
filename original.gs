/**
 * CONTROL DE CONTRATOS POR PROYECTOS - Apps Script
 * Versión: 2.4 (estado refleja directamente la fuente, sin condición de fechaSalida)
 * Fecha: Mayo 2026
 */

// ============================================================
// CONFIGURACIÓN - URLs de datos remotos
// ============================================================

var CONFIG = {
  YAPE: {
    url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRMfgLqKN9saaKvdMMQ_cu2ZtBjaE-VgQjVKKIe0k3MpDZLPVDINb_EilHMzQSeiw/pub?output=csv',
    clave: 'CELULAR',
    indiceClave: 12,
    nombreHoja: 'Anexar1'
  },
  BNB: {
    url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSMPGglYuRuNVPuls71_XeSjxVtejUZ7FQocgAeoMWSVHLANr98MWTE3gyCKvOEzA/pub?output=csv',
    clave: 'CI',
    indiceClave: 2,
    nombreHoja: null
  },
  ZAS: {
    url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQYCuKQlYTn7UWL9qgUEXjnvyD9BHQMor6nWj4uBBXay6GA0jvV2iTe_Tg54n2WWQ/pub?output=csv&gid=1900067435',
    urlExtraterritorial: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQYCuKQlYTn7UWL9qgUEXjnvyD9BHQMor6nWj4uBBXay6GA0jvV2iTe_Tg54n2WWQ/pub?gid=232295204&single=true&output=csv',
    clave: 'CODIGO',
    indiceClave: 1,
    nombreHoja: null
  }
};

// Mapeo de índices origen -> campos destino
// YAPE: EMPRESA(0), CODIGO(1), C.I.(2), NOMBRE(3), AP_PAT(4), AP_MAT(5), NOMBRES(6),
//       FECHA_ING(7), CARGO(8), CIUDAD(9), SUPERVISOR(10), ESTADO(11), CELULAR(12)
var MAPEO_YAPE = {
  empresa: 0, codigo: 1, ci: 2, nombreCompleto: 3,
  apellidoPaterno: 4, apellidoMaterno: 5, nombres: 6,
  fechaIngreso: 7, cargo: 8, ciudad: 9, supervisor: 10,
  estado: 11, celular: 12, fechaSalida: 13
};

// BNB: misma estructura que YAPE
var MAPEO_BNB = {
  empresa: 0, codigo: 1, ci: 2, nombreCompleto: 3,
  apellidoPaterno: 4, apellidoMaterno: 5, nombres: 6,
  fechaIngreso: 7, cargo: 8, ciudad: 9, supervisor: 10,
  estado: 11, celular: 12, fechaSalida: 13
};

// ZAS: similar pero con Dpto en índice 20
var MAPEO_ZAS = {
  empresa: 0, codigo: 1, ci: 2, nombreCompleto: 3,
  apellidoPaterno: 4, apellidoMaterno: 5, nombres: 6,
  fechaIngreso: 7, cargo: 8, ciudad: 9, supervisor: 10,
  estado: 12, celular: 13, fechaSalida: 14, dpto: 21, ciudad_txt: 22, usuario: 24
};

var ESTRUCTURA_ESTANDAR = [
  'EMPRESA', 'CODIGO', 'C.I.', 'CELULAR', 'NOMBRE COMPLETO',
  'APELLIDO PATERNO', 'APELLIDO MATERNO', 'NOMBRES', 'CARGO',
  'CIUDAD', 'DPTO', 'SUPERVISOR', 'FECHA DE CONTRATO',
  'FECHA DE INGRESO', 'ESTADO', 'FIRMA', 'ESTADO DE CONTRATO',
  'OBS', 'OBS SUPERVISOR', 'OBS SUBGERENTE'
];

var MAPA_DEPARTAMENTO = {
  'LA PAZ': 'La Paz', 'EL ALTO': 'La Paz', 'SANTA CRUZ': 'Santa Cruz',
  'COCHABAMBA': 'Cochabamba', 'BENI': 'Beni', 'PANDO': 'Pando',
  'SUCRE': 'Chuquisaca', 'ORURO': 'Oruro', 'POTOSI': 'Potosí',
  'POTOSÍ': 'Potosí', 'TARIJA': 'Tarija', 'NORTE INTEGRADO': 'Santa Cruz',
  'TRINIDAD': 'Beni', 'COBIJA': 'Pando',
  // Códigos cortos usados en ZAS
  'LP': 'La Paz', 'SC': 'Santa Cruz', 'CB': 'Cochabamba',
  'OR': 'Oruro', 'PT': 'Potosí', 'TJ': 'Tarija',
  'CH': 'Chuquisaca', 'BN': 'Beni', 'PD': 'Pando'
};

// ============================================================
// MENÚ
// ============================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Procesar')
    .addItem('Actualizar YAPE', 'actualizarYAPE')
    .addItem('Actualizar BNB', 'actualizarBNB')
    .addItem('Actualizar ZAS', 'actualizarZAS')
    .addSeparator()
    .addItem('Normalizar Estructura', 'normalizarEstructura')
    .addItem('Llenar Departamento', 'llenarDepartamento')
    .addSeparator()
    .addItem('Verificar URLs', 'verificarURLs')
    .addToUi();
}

// ============================================================
// NORMALIZAR ESTRUCTURA (20 columnas estándar)
// ============================================================

var ENCABEZADOS_ESTANDAR = [
  'EMPRESA',
  'CODIGO',
  'C.I.',
  'CELULAR',
  'NOMBRE COMPLETO',
  'APELLIDO PATERNO',
  'APELLIDO MATERNO',
  'NOMBRES',
  'CARGO',
  'CIUDAD',
  'DPTO',
  'SUPERVISOR',
  'FECHA DE CONTRATO',
  'FECHA DE INGRESO',
  'ESTADO',
  'FIRMA',
  'ESTADO DE CONTRATO',
  'OBS',
  'OBS SUPERVISOR',
  'OBS SUBGERENTE'
];


function normalizarEstructura() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var resultados = [];

  var hojas = ['CONTRATOS YAPE', 'CONTRATOS BNB', 'CONTRATOS ZAS'];

  hojas.forEach(function(nombreHoja) {
    var hoja = ss.getSheetByName(nombreHoja);

    if (!hoja) {
      resultados.push(nombreHoja + ': No encontrada');
      return;
    }

    var resultado = normalizarHoja(hoja, ENCABEZADOS_ESTANDAR);
    resultados.push(nombreHoja + ': ' + resultado);
  });

  ui.alert('Normalización completada:\n\n' + resultados.join('\n'));
}

function normalizarHoja(hoja, encabezadosEstandar) {
  var TOTAL_COLUMNAS = encabezadosEstandar.length; // 20
  var ultimaFila = hoja.getLastRow();
  var ultimaColumna = hoja.getLastColumn();
  var cambios = [];

  if (ultimaColumna === 0 || ultimaFila === 0) {
    hoja.getRange(1, 1, 1, TOTAL_COLUMNAS).setValues([encabezadosEstandar]);
    formatearEncabezados(hoja, TOTAL_COLUMNAS);
    return 'Encabezados creados';
  }

  var encabezadosActuales = hoja.getRange(1, 1, 1, ultimaColumna).getValues()[0];
  var mapaColumnas = {};

  for (var i = 0; i < encabezadosActuales.length; i++) {
    var nombreNorm = normalizarNombreColumna(String(encabezadosActuales[i]));

    if (nombreNorm) {
      var datos = [];

      if (ultimaFila > 1) {
        datos = hoja.getRange(2, i + 1, ultimaFila - 1, 1).getValues();
      }

      mapaColumnas[nombreNorm] = datos;
    }
  }

  var nuevasDatos = [];

  if (ultimaFila > 1) {
    for (var fila = 0; fila < ultimaFila - 1; fila++) {
      var nuevaFila = [];

      for (var col = 0; col < TOTAL_COLUMNAS; col++) {
        var nombreEstandar = normalizarNombreColumna(encabezadosEstandar[col]);

        if (mapaColumnas[nombreEstandar] && mapaColumnas[nombreEstandar][fila]) {
          nuevaFila.push(mapaColumnas[nombreEstandar][fila][0]);
        } else {
          nuevaFila.push('');
        }
      }

      nuevasDatos.push(nuevaFila);
    }
  }

  if (ultimaColumna < TOTAL_COLUMNAS) {
    hoja.insertColumnsAfter(ultimaColumna, TOTAL_COLUMNAS - ultimaColumna);
    cambios.push((TOTAL_COLUMNAS - ultimaColumna) + ' columnas agregadas');
  }

  if (ultimaColumna > TOTAL_COLUMNAS) {
    cambios.push((ultimaColumna - TOTAL_COLUMNAS) + ' columnas extra detectadas, NO eliminadas');
  }

  hoja.getRange(1, 1, 1, TOTAL_COLUMNAS).setValues([encabezadosEstandar]);

  if (nuevasDatos.length > 0) {
    hoja.getRange(2, 1, nuevasDatos.length, TOTAL_COLUMNAS).setValues(nuevasDatos);
  }

  formatearEncabezados(hoja, TOTAL_COLUMNAS);

  if (cambios.length === 0) {
    cambios.push('OK');
  }

  return cambios.join(', ');
}

function normalizarNombreColumna(nombre) {
  if (!nombre) return '';

  var norm = String(nombre).toUpperCase().trim();

  var mapeoVariaciones = {
    'ESTADO (ACTIVO/INACTIVO)': 'ESTADO',
    'ESTADO(ACTIVO/INACTIVO)': 'ESTADO',
    'ESTADO ACTIVO/INACTIVO': 'ESTADO',
    'OBS SUPERVIS': 'OBS SUPERVISOR',
    'OBS SUBGEREN': 'OBS SUBGERENTE',
    'NOMBRES Y APELLIDOS': 'NOMBRE COMPLETO',
    'NOMBRE': 'NOMBRE COMPLETO',
    'CI': 'C.I.',
    'C.I': 'C.I.',
    'DEPARTAMENTO': 'DPTO',
    'FECHA CONTRATO': 'FECHA DE CONTRATO',
    'FECHA_CONTRATO': 'FECHA DE CONTRATO',
    'FECHA DE CONTRATO': 'FECHA DE CONTRATO',
    'FECHA INGRESO': 'FECHA DE INGRESO',
    'FECHA_INGRESO': 'FECHA DE INGRESO',
    'FECHA DE INGRESO': 'FECHA DE INGRESO'
  };

  return mapeoVariaciones[norm] || norm;
}

function formatearEncabezados(hoja, totalColumnas) {
  var rangoEncabezados = hoja.getRange(1, 1, 1, totalColumnas);

  rangoEncabezados.setFontWeight('bold');
  rangoEncabezados.setBackground('#D9EAD3');
  rangoEncabezados.setHorizontalAlignment('center');

  hoja.setFrozenRows(1);
}

// ============================================================
// FUNCIONES UTILITARIAS
// ============================================================

function parseFechaYAPE(valor) {
  if (valor === null || valor === undefined || valor === '') return '';

  var texto = String(valor).trim();
  var partes = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (!partes) {
    return texto;
  }

  var p1 = parseInt(partes[1], 10);
  var p2 = parseInt(partes[2], 10);
  var anio = partes[3];

  var dia, mes;

  if (p1 <= 12) {
    // Fuente YAPE/BNB usa formato M/D/YYYY
    mes = p1;
    dia = p2;
  } else {
    dia = p1;
    mes = p2;
  }

  var diaStr = ('0' + dia).slice(-2);
  var mesStr = ('0' + mes).slice(-2);

  return diaStr + '/' + mesStr + '/' + anio;
}

function aNombrePropio(texto) {
  if (!texto || texto === '') return '';
  return String(texto).toLowerCase().replace(/(?:^|\s)\S/g, function(letra) {
    return letra.toUpperCase();
  });
}

function obtenerDepartamento(ciudad) {
  if (!ciudad || ciudad === '') return '';
  var ciudadNormalizada = String(ciudad).toUpperCase().trim();
  if (MAPA_DEPARTAMENTO[ciudadNormalizada]) {
    return MAPA_DEPARTAMENTO[ciudadNormalizada];
  }
  for (var key in MAPA_DEPARTAMENTO) {
    if (ciudadNormalizada.indexOf(key) !== -1 || key.indexOf(ciudadNormalizada) !== -1) {
      return MAPA_DEPARTAMENTO[key];
    }
  }
  return '';
}

function estaVacio(valor) {
  return valor === '' || valor === null || valor === undefined;
}

function descargarCSV(url) {
  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('HTTP ' + response.getResponseCode());
    }

    return Utilities.parseCsv(response.getContentText('UTF-8'));
  } catch (e) {
    throw new Error('Error al descargar: ' + e.message);
  }
}

function obtenerUsuariosExtraterritorialesZAS() {
  var datos = descargarCSV(CONFIG.ZAS.urlExtraterritorial);
  var usuarios = new Set();

  if (datos.length < 2) return usuarios;

  datos.shift();

  datos.forEach(function(fila) {
    var usuarioTemporal = String(fila[0] || '').trim().toUpperCase();

    if (usuarioTemporal !== '') {
      usuarios.add(usuarioTemporal);
    }
  });

  return usuarios;
}

function descargarXLSX(url, nombreHoja) {
  var archivoTemporal = null;
  var archivoConvertido = null;

  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('HTTP ' + response.getResponseCode());
    }

    var blob = response.getBlob().setName('temp_descarga.xlsx');
    archivoTemporal = DriveApp.createFile(blob);
    var fileId = archivoTemporal.getId();

    var recurso = {
      title: 'TEMP_CONVERTIDO_' + new Date().getTime(),
      mimeType: MimeType.GOOGLE_SHEETS
    };

    var convertido = Drive.Files.copy(recurso, fileId);
    var convertidoId = convertido.id;
    archivoConvertido = DriveApp.getFileById(convertidoId);

    var tempSS = SpreadsheetApp.openById(convertidoId);

    var hoja = null;
    if (nombreHoja) {
      hoja = tempSS.getSheetByName(nombreHoja);

      if (!hoja) {
        var hojas = tempSS.getSheets();
        for (var i = 0; i < hojas.length; i++) {
          var nombreActual = hojas[i].getName().toLowerCase().trim();
          var nombreBuscado = nombreHoja.toLowerCase().trim();

          if (nombreActual === nombreBuscado || nombreActual.indexOf(nombreBuscado) !== -1) {
            hoja = hojas[i];
            break;
          }
        }
      }
    }

    if (!hoja) {
      hoja = tempSS.getSheets()[0];
    }

    return hoja.getDataRange().getValues();

  } catch (e) {
    throw new Error('Error al procesar XLSX: ' + e.message);
  } finally {
    try { if (archivoTemporal)   archivoTemporal.setTrashed(true);  } catch (e1) {}
    try { if (archivoConvertido) archivoConvertido.setTrashed(true); } catch (e2) {}
  }
}

function verificarURLs() {
  var resultados = [];

  try {
    var datosYAPE;
    if (CONFIG.YAPE.url.includes('.xlsx') || CONFIG.YAPE.url.includes('output=xlsx')) {
      datosYAPE = descargarXLSX(CONFIG.YAPE.url, CONFIG.YAPE.nombreHoja);
    } else {
      datosYAPE = descargarCSV(CONFIG.YAPE.url);
    }
    resultados.push('YAPE: OK (' + (datosYAPE.length - 1) + ' registros)');
  } catch (e) {
    resultados.push('YAPE: ERROR - ' + e.message);
  }

  try {
    var datosBNB = descargarCSV(CONFIG.BNB.url);
    resultados.push('BNB: OK (' + (datosBNB.length - 1) + ' registros)');
  } catch (e) {
    resultados.push('BNB: ERROR - ' + e.message);
  }

  try {
    var datosZAS = descargarCSV(CONFIG.ZAS.url);
    resultados.push('ZAS: OK (' + (datosZAS.length - 1) + ' registros)');
  } catch (e) {
    resultados.push('ZAS: ERROR - ' + e.message);
  }

  SpreadsheetApp.getUi().alert('Verificación de URLs:\n\n' + resultados.join('\n'));
}

function parseFechaZASComoDate(valor) {
  if (!valor) return '';

  var texto = String(valor).trim().toLowerCase();

  var meses = {
    'ene': 0, 'feb': 1, 'mar': 2, 'abr': 3, 'may': 4, 'jun': 5,
    'jul': 6, 'ago': 7, 'sep': 8, 'sept': 8, 'set': 8, 'oct': 9,
    'nov': 10, 'dic': 11
  };

  // Formato dd/mm/yyyy
  var slash = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slash) {
    var dia = parseInt(slash[1], 10);
    var mes = parseInt(slash[2], 10) - 1;
    var anio = parseInt(slash[3], 10);
    return new Date(anio, mes, dia);
  }

  // Formato dd-mmm-yyyy, ejemplo: 5-sept-2025
  var textoMes = texto.match(/^(\d{1,2})-([a-z]+)-(\d{4})$/);
  if (textoMes) {
    var dia2 = parseInt(textoMes[1], 10);
    var mesTxt = textoMes[2];
    var anio2 = parseInt(textoMes[3], 10);

    if (meses[mesTxt] !== undefined) {
      return new Date(anio2, meses[mesTxt], dia2);
    }
  }

  return valor;
}

// ============================================================
// ACTUALIZAR YAPE - v2.3
//
// Llave: CODIGO + C.I. (única en fuente, 0 duplicados)
// Actualiza TODOS los campos automáticos (A-D, E-L, N, O)
// Conserva columnas manuales (M, P, Q, R, S, T)
// Escribe existentes en un único bloque (una sola llamada setValues)
// Soporta filas duplicadas en el sheet: actualiza todas
// ============================================================

function actualizarYAPE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDestino = ss.getSheetByName('CONTRATOS YAPE');

  var TOTAL_COLUMNAS = 20;
  var COL_FECHA_INGRESO = 14; // columna N, 1-based

  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert('Error: No se encontró la hoja "CONTRATOS YAPE"');
    return;
  }

  var encabezado = hojaDestino
    .getRange(1, 1, 1, hojaDestino.getLastColumn())
    .getValues()[0];
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

  datosOrigen.shift();

  // ── 1. Leer destino completo en memoria ──
  var ultimaFila    = hojaDestino.getLastRow();
  var datosDestino  = [];
  var mapaDestino   = {}; // 'CODIGO|CI' → [idx0, idx1, ...] — soporta duplicados
  var clavesDestino = new Set();

  if (ultimaFila > 1) {
    datosDestino = hojaDestino
      .getRange(2, 1, ultimaFila - 1, TOTAL_COLUMNAS)
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

  // ── 2. Procesar fuente ──
  var registrosNuevos          = [];
  var registrosActualizados    = 0;
  var registrosNuevosAgregados = 0;
  var estadosActualizados      = 0;

  datosOrigen.forEach(function(filaOrigen) {
    var codigo       = String(filaOrigen[MAPEO_YAPE.codigo]      || '').trim();
    var ci           = String(filaOrigen[MAPEO_YAPE.ci]          || '').trim();
    var estadoOrigen = String(filaOrigen[MAPEO_YAPE.estado]      || '').trim().toUpperCase();
    var ciudad       = filaOrigen[MAPEO_YAPE.ciudad]              || '';

    if (codigo === '' || ci === '') return;

    var key = codigo + '|' + ci;

    // El estado refleja directamente la fuente — si la fuente dice INACTIVO, el sheet dice INACTIVO
    var estadoFinal = estadoOrigen !== '' ? estadoOrigen : 'ACTIVO';

    if (clavesDestino.has(key)) {
      // Actualizar todas las filas con esta llave (cubre duplicados en el sheet)
      mapaDestino[key].forEach(function(idx) {
        var filaActual = datosDestino[idx];

        var estadoActual = String(filaActual[14] || '').toUpperCase().trim();
        if (estadoActual !== 'INACTIVO' && estadoFinal === 'INACTIVO') estadosActualizados++;

        // Automáticos: A B C D E F G H I J K L N O
        // Manuales intactos: M[12] P[15] Q[16] R[17] S[18] T[19]
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
          filaActual[12],  // M  FECHA DE CONTRATO  ← manual
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
      // Insertar solo si está ACTIVO en la fuente
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
        '',                                                        // M  manual, vacío
        parseFechaYAPE(filaOrigen[MAPEO_YAPE.fechaIngreso]),      // N
        'ACTIVO',                                                  // O
        '',                                                        // P  manual
        '',                                                        // Q  manual
        '',                                                        // R  manual
        '',                                                        // S  manual
        ''                                                         // T  manual
      ]);

      clavesDestino.add(key);
      registrosNuevosAgregados++;
    }
  });

  // ── 3. Escribir existentes en un solo bloque ──
  if (datosDestino.length > 0) {
    hojaDestino.getRange(2, 1, datosDestino.length, TOTAL_COLUMNAS).setValues(datosDestino);
    hojaDestino.getRange(2, COL_FECHA_INGRESO, datosDestino.length, 1).setNumberFormat('@');
  }

  // ── 4. Agregar nuevos en un solo bloque ──
  if (registrosNuevos.length > 0) {
    var filaInicio = hojaDestino.getLastRow() + 1;
    hojaDestino.getRange(filaInicio, 1, registrosNuevos.length, TOTAL_COLUMNAS).setValues(registrosNuevos);
    hojaDestino.getRange(filaInicio, COL_FECHA_INGRESO, registrosNuevos.length, 1).setNumberFormat('@');
  }

  ss.toast('', '', 1);

  SpreadsheetApp.getUi().alert(
    'Actualización YAPE completada.\n\n' +
    'Registros nuevos agregados: '        + registrosNuevosAgregados + '\n' +
    'Registros existentes actualizados: ' + registrosActualizados    + '\n' +
    'Estados actualizados a INACTIVO: '   + estadosActualizados      + '\n\n' +
    'Columnas manuales conservadas: M, P, Q, R, S y T'
  );
}

// ============================================================
// ACTUALIZAR BNB - v2.3
//
// Llave: CODIGO + C.I.
// Misma lógica que YAPE: actualiza todos los campos automáticos,
// conserva manuales (M, P, Q, R, S, T), escribe en bloque.
// ============================================================

function actualizarBNB() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDestino = ss.getSheetByName('CONTRATOS BNB');

  var TOTAL_COLUMNAS = 20;
  var COL_FECHA_INGRESO = 14; // columna N, 1-based

  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert('Error: No se encontró la hoja "CONTRATOS BNB"');
    return;
  }

  var encabezado = hojaDestino
    .getRange(1, 1, 1, hojaDestino.getLastColumn())
    .getValues()[0];
  if (encabezado[0] !== 'EMPRESA' || encabezado[2] !== 'C.I.' || encabezado[3] !== 'CELULAR') {
    SpreadsheetApp.getUi().alert('Error: La hoja CONTRATOS BNB no está estandarizada.');
    return;
  }

  var datosOrigen;
  try {
    ss.toast('Descargando datos de BNB...', 'Procesando', -1);
    datosOrigen = descargarCSV(CONFIG.BNB.url);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error al descargar datos de BNB:\n' + e.message);
    return;
  }

  if (datosOrigen.length < 2) {
    SpreadsheetApp.getUi().alert('No se encontraron datos en la fuente BNB');
    return;
  }

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
      var codigo = String(fila[1] || '').trim(); // col B
      var ci     = String(fila[2] || '').trim(); // col C
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
    var codigo       = String(filaOrigen[MAPEO_BNB.codigo]      || '').trim();
    var ci           = String(filaOrigen[MAPEO_BNB.ci]          || '').trim();
    var estadoOrigen = String(filaOrigen[MAPEO_BNB.estado]      || '').trim().toUpperCase();
    var ciudad       = filaOrigen[MAPEO_BNB.ciudad]              || '';

    if (codigo === '' || ci === '') return;

    var key = codigo + '|' + ci;

    // El estado refleja directamente la fuente
    var estadoFinal = estadoOrigen !== '' ? estadoOrigen : 'ACTIVO';

    if (clavesDestino.has(key)) {
      mapaDestino[key].forEach(function(idx) {
        var filaActual = datosDestino[idx];

        var estadoActual = String(filaActual[14] || '').toUpperCase().trim();
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
          filaActual[12],  // M  FECHA DE CONTRATO  ← manual
          parseFechaYAPE(filaOrigen[MAPEO_BNB.fechaIngreso]),      // N [13]
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
      if (estadoOrigen !== 'ACTIVO') return;

      registrosNuevos.push([
        filaOrigen[MAPEO_BNB.empresa]                    || '',  // A
        codigo,                                                    // B
        ci,                                                        // C
        filaOrigen[MAPEO_BNB.celular]                    || '',  // D
        aNombrePropio(filaOrigen[MAPEO_BNB.nombreCompleto]),     // E
        aNombrePropio(filaOrigen[MAPEO_BNB.apellidoPaterno]),    // F
        aNombrePropio(filaOrigen[MAPEO_BNB.apellidoMaterno]),    // G
        aNombrePropio(filaOrigen[MAPEO_BNB.nombres]),            // H
        filaOrigen[MAPEO_BNB.cargo]                      || '',  // I
        ciudad,                                                    // J
        obtenerDepartamento(ciudad),                               // K
        aNombrePropio(filaOrigen[MAPEO_BNB.supervisor]),         // L
        '',                                                        // M  manual, vacío
        parseFechaYAPE(filaOrigen[MAPEO_BNB.fechaIngreso]),      // N
        'ACTIVO',                                                  // O
        '',                                                        // P  manual
        '',                                                        // Q  manual
        '',                                                        // R  manual
        '',                                                        // S  manual
        ''                                                         // T  manual
      ]);

      clavesDestino.add(key);
      registrosNuevosAgregados++;
    }
  });

  // ── 3. Escribir existentes en un solo bloque ──
  if (datosDestino.length > 0) {
    hojaDestino.getRange(2, 1, datosDestino.length, TOTAL_COLUMNAS).setValues(datosDestino);
    hojaDestino.getRange(2, COL_FECHA_INGRESO, datosDestino.length, 1).setNumberFormat('@');
  }

  // ── 4. Agregar nuevos en un solo bloque ──
  if (registrosNuevos.length > 0) {
    var filaInicio = hojaDestino.getLastRow() + 1;
    hojaDestino.getRange(filaInicio, 1, registrosNuevos.length, TOTAL_COLUMNAS).setValues(registrosNuevos);
    hojaDestino.getRange(filaInicio, COL_FECHA_INGRESO, registrosNuevos.length, 1).setNumberFormat('@');
  }

  ss.toast('', '', 1);

  SpreadsheetApp.getUi().alert(
    'Actualización BNB completada.\n\n' +
    'Registros nuevos agregados: '        + registrosNuevosAgregados + '\n' +
    'Registros existentes actualizados: ' + registrosActualizados    + '\n' +
    'Estados actualizados a INACTIVO: '   + estadosActualizados      + '\n\n' +
    'Columnas manuales conservadas: M, P, Q, R, S y T'
  );
}

// ============================================================
// ACTUALIZAR ZAS - v2.3
//
// Llave: CODIGO + C.I.
// Particularidades de ZAS:
//   - parseFechaZASComoDate (retorna Date, formato dd/mm/yyyy)
//   - ciudad viene de ciudad_txt (col 22) con fallback a ciudad (col 9)
//   - dpto viene directo de la fuente (col 21) o se deriva de ciudad
//   - Usuarios extraterritoriales se omiten en inserción
// Conserva columnas manuales (M, P, Q, R, S, T), escribe en bloque.
// ============================================================

function actualizarZAS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDestino = ss.getSheetByName('CONTRATOS ZAS');

  var TOTAL_COLUMNAS = 20;
  var COL_FECHA_INGRESO = 14; // columna N, 1-based

  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert('Error: No se encontró la hoja "CONTRATOS ZAS"');
    return;
  }

  var encabezado = hojaDestino
    .getRange(1, 1, 1, hojaDestino.getLastColumn())
    .getValues()[0];
  if (encabezado[0] !== 'EMPRESA' || encabezado[2] !== 'C.I.' || encabezado[3] !== 'CELULAR') {
    SpreadsheetApp.getUi().alert('Error: La hoja CONTRATOS ZAS no está estandarizada.');
    return;
  }

  var datosOrigen;
  var usuariosExtraterritoriales = obtenerUsuariosExtraterritorialesZAS();
  var omitidosExtraterritorial   = 0;

  try {
    ss.toast('Descargando datos de ZAS...', 'Procesando', -1);
    datosOrigen = descargarCSV(CONFIG.ZAS.url);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error al descargar datos de ZAS:\n' + e.message);
    return;
  }

  if (datosOrigen.length < 2) {
    SpreadsheetApp.getUi().alert('No se encontraron datos en la fuente ZAS');
    return;
  }

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
      var codigo = String(fila[1] || '').trim(); // col B
      var ci     = String(fila[2] || '').trim(); // col C
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
    var codigo       = String(filaOrigen[MAPEO_ZAS.codigo]      || '').trim();
    var ci           = String(filaOrigen[MAPEO_ZAS.ci]          || '').trim();
    var estadoOrigen = String(filaOrigen[MAPEO_ZAS.estado]      || '').trim().toUpperCase();
    var usuario      = String(filaOrigen[MAPEO_ZAS.usuario]     || '').trim().toUpperCase();

    // Ciudad: usar ciudad_txt (más descriptivo) con fallback a ciudad
    var ciudad = filaOrigen[MAPEO_ZAS.ciudad_txt] || filaOrigen[MAPEO_ZAS.ciudad] || '';

    // Departamento: usar el de la fuente si existe, sino derivar de ciudad
    var dptoRaw = filaOrigen[MAPEO_ZAS.dpto] || '';
    var dpto    = estaVacio(dptoRaw)
      ? obtenerDepartamento(ciudad)
      : (obtenerDepartamento(dptoRaw) || dptoRaw);

    if (codigo === '') return;

    var key = codigo + '|' + ci;

    // El estado refleja directamente la fuente
    var estadoFinal = estadoOrigen !== '' ? estadoOrigen : 'ACTIVO';

    if (clavesDestino.has(key)) {
      mapaDestino[key].forEach(function(idx) {
        var filaActual = datosDestino[idx];

        var estadoActual = String(filaActual[14] || '').toUpperCase().trim();
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
          filaActual[12],  // M  FECHA DE CONTRATO  ← manual
          parseFechaZASComoDate(filaOrigen[MAPEO_ZAS.fechaIngreso]), // N [13]
          estadoFinal,                                                 // O [14]
          filaActual[15],  // P  FIRMA              ← manual
          filaActual[16],  // Q  ESTADO DE CONTRATO ← manual
          filaActual[17],  // R  OBS                ← manual
          filaActual[18],  // S  OBS SUPERVISOR     ← manual
          filaActual[19]   // T  OBS SUBGERENTE     ← manual
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
        filaOrigen[MAPEO_ZAS.empresa] || 'ZAS',                    // A
        codigo,                                                      // B
        ci,                                                          // C
        filaOrigen[MAPEO_ZAS.celular]                      || '',  // D
        aNombrePropio(filaOrigen[MAPEO_ZAS.nombreCompleto]),       // E
        aNombrePropio(filaOrigen[MAPEO_ZAS.apellidoPaterno]),      // F
        aNombrePropio(filaOrigen[MAPEO_ZAS.apellidoMaterno]),      // G
        aNombrePropio(filaOrigen[MAPEO_ZAS.nombres]),              // H
        filaOrigen[MAPEO_ZAS.cargo]                        || '',  // I
        ciudad,                                                      // J
        dpto,                                                        // K
        aNombrePropio(filaOrigen[MAPEO_ZAS.supervisor]),           // L
        '',                                                          // M  manual, vacío
        parseFechaZASComoDate(filaOrigen[MAPEO_ZAS.fechaIngreso]), // N
        'ACTIVO',                                                    // O
        '',                                                          // P  manual
        '',                                                          // Q  manual
        '',                                                          // R  manual
        '',                                                          // S  manual
        ''                                                           // T  manual
      ]);

      clavesDestino.add(key);
      registrosNuevosAgregados++;
    }
  });

  // ── 3. Escribir existentes en un solo bloque ──
  if (datosDestino.length > 0) {
    hojaDestino.getRange(2, 1, datosDestino.length, TOTAL_COLUMNAS).setValues(datosDestino);
    hojaDestino.getRange(2, COL_FECHA_INGRESO, datosDestino.length, 1).setNumberFormat('dd/mm/yyyy');
  }

  // ── 4. Agregar nuevos en un solo bloque ──
  if (registrosNuevos.length > 0) {
    var filaInicio = hojaDestino.getLastRow() + 1;
    hojaDestino.getRange(filaInicio, 1, registrosNuevos.length, TOTAL_COLUMNAS).setValues(registrosNuevos);
    hojaDestino.getRange(filaInicio, COL_FECHA_INGRESO, registrosNuevos.length, 1).setNumberFormat('dd/mm/yyyy');
  }

  ss.toast('', '', 1);

  SpreadsheetApp.getUi().alert(
    'Actualización ZAS completada.\n\n' +
    'Registros nuevos agregados: '                       + registrosNuevosAgregados + '\n' +
    'Registros existentes actualizados: '                + registrosActualizados    + '\n' +
    'Estados actualizados a INACTIVO: '                  + estadosActualizados      + '\n' +
    'Registros omitidos (Producción Extraterritorial): ' + omitidosExtraterritorial + '\n\n' +
    'Columnas manuales conservadas: M, P, Q, R, S y T'
  );
}

// ============================================================
// LLENAR DEPARTAMENTO
// ============================================================

function llenarDepartamento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resultados = [];
  var COL_CIUDAD = 10, COL_DPTO = 11;

  ['CONTRATOS YAPE', 'CONTRATOS BNB', 'CONTRATOS ZAS'].forEach(function(nombreHoja) {
    var hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) { resultados.push(nombreHoja + ': No encontrada'); return; }

    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) { resultados.push(nombreHoja + ': Sin datos'); return; }

    var datosCiudad = hoja.getRange(2, COL_CIUDAD, ultimaFila - 1, 1).getValues();
    var datosDpto   = hoja.getRange(2, COL_DPTO,   ultimaFila - 1, 1).getValues();
    var actualizados = 0;

    datosCiudad.forEach(function(fila, index) {
      if (estaVacio(datosDpto[index][0]) && fila[0]) {
        var dpto = obtenerDepartamento(fila[0]);
        if (dpto !== '') {
          hoja.getRange(index + 2, COL_DPTO).setValue(dpto);
          actualizados++;
        }
      }
    });

    resultados.push(nombreHoja + ': ' + actualizados + ' actualizados');
  });

  SpreadsheetApp.getUi().alert('Llenar Departamento completado:\n\n' + resultados.join('\n'));
}
