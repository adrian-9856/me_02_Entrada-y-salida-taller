// ==================== CONFIGURACIÓN ====================
var NOMBRE_EMPRESA = "mi_eelo";
var NIT_EMPRESA = "NIT: 12345678-9";
var DIRECCION = "Creamos Guatemala";
var HORAS_JORNADA_NORMAL = 7; // Horas de trabajo por día

// URL de tu exportación de KoboToolbox
var URL_KOBO = "https://kf.kobotoolbox.org/api/v2/assets/agi395bJj6ojXJzPPDT9n6/export-settings/es4oUjEmPvovgLd6Y5yrQ4K/data.csv";

// Nombres de los días en español (0=Domingo, 1=Lunes, ..., 6=Sábado)
var DIAS_SEMANA = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];

// ==================== MENÚ Y TRIGGERS ====================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⏱️ Control de Horas')
    .addItem('🔄 Actualizar desde Kobo', 'importarCSVdesdeKobo')
    .addSeparator()
    .addItem('📅 Reporte por Día', 'generarReportePorDia')
    .addItem('📆 Reporte por Semana', 'generarReportePorSemana')
    .addItem('🗓️ Reporte por Mes', 'generarReportePorMes')
    .addItem('📊 Reporte por Rango', 'generarReportePorRango')
    .addItem('📋 Reporte Completo', 'generarReporteTodo')
    .addSeparator()
    .addItem('📚 Configurar Días de Estudio', 'crearHojaDiasEstudio')
    .addItem('⚙️ Configurar actualización automática', 'configurarActualizacionAutomatica')
    .addToUi();
}

// ==================== IMPORTACIÓN DESDE KOBO ====================
function importarCSVdesdeKobo() {
  try {
    var response = UrlFetchApp.fetch(URL_KOBO);
    var csv = response.getContentText();
    var datos = Utilities.parseCsv(csv, ";");

    if (datos.length === 0) return;

    // Encontrar cuántas columnas válidas hay (encabezados no vacíos)
    var encabezados = datos[0];
    var ultimaColumnaValida = 0;
    for (var c = 0; c < encabezados.length; c++) {
      if (String(encabezados[c]).trim() !== '') {
        ultimaColumnaValida = c + 1;
      }
    }

    // Si no hay columnas válidas, no importar
    if (ultimaColumnaValida === 0) {
      SpreadsheetApp.getUi().alert('Error: El CSV no tiene encabezados válidos.');
      return;
    }

    // Recortar todas las filas para solo incluir columnas con encabezado
    var datosLimpios = [];
    for (var f = 0; f < datos.length; f++) {
      datosLimpios.push(datos[f].slice(0, ultimaColumnaValida));
    }

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = spreadsheet.getSheetByName("DatosKobo");

    if (!hoja) {
      hoja = spreadsheet.insertSheet("DatosKobo");
    }

    hoja.clearContents();

    hoja.getRange(1, 1, datosLimpios.length, datosLimpios[0].length).setValues(datosLimpios);

    // Formatear encabezados
    hoja.getRange(1, 1, 1, datosLimpios[0].length)
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('#ffffff');

    hoja.setFrozenRows(1);

    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Datos actualizados (' + (datosLimpios.length - 1) + ' registros, ' + ultimaColumnaValida + ' columnas)', 'Importación Exitosa', 3);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error al importar datos: ' + e.message);
  }
}

// Configurar triggers automáticos (cada hora + al abrir)
function configurarActualizacionAutomatica() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var handler = triggers[i].getHandlerFunction();
    if (handler === 'importarCSVdesdeKobo' || handler === 'importarAlAbrir') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Trigger cada hora
  ScriptApp.newTrigger('importarCSVdesdeKobo')
    .timeBased()
    .everyHours(1)
    .create();

  // Trigger instalable al abrir (tiene permisos completos, a diferencia del onOpen simple)
  ScriptApp.newTrigger('importarAlAbrir')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();

  SpreadsheetApp.getUi().alert('✅ Actualización automática configurada:\n- Cada hora\n- Al abrir la hoja de cálculo');
}

// Función llamada por el trigger instalable onOpen (tiene permisos completos)
function importarAlAbrir() {
  importarCSVdesdeKobo();
}

// ==================== DÍAS DE ESTUDIO ====================
// Crea o abre la hoja "DiasEstudio" donde se configuran los días que cada participante estudia
function crearHojaDiasEstudio() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("DiasEstudio");

  if (!hoja) {
    hoja = spreadsheet.insertSheet("DiasEstudio");

    // Encabezados
    var encabezados = ['Participante', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo', 'Fecha Inicio', 'Fecha Fin'];
    hoja.getRange(1, 1, 1, 10).setValues([encabezados]);
    hoja.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#7b1fa2').setFontColor('#ffffff').setHorizontalAlignment('center');
    hoja.setFrozenRows(1);

    // Llenar con participantes existentes si hay datos de Kobo
    var hojaKobo = spreadsheet.getSheetByName("DatosKobo");
    if (hojaKobo) {
      var datosKobo = hojaKobo.getDataRange().getValues();
      var encabezadosKobo = datosKobo[0];
      var colPart = -1;
      for (var i = 0; i < encabezadosKobo.length; i++) {
        if (String(encabezadosKobo[i]).trim() === 'Participante') { colPart = i; break; }
      }
      if (colPart !== -1) {
        var participantes = {};
        for (var f = 1; f < datosKobo.length; f++) {
          var nombre = String(datosKobo[f][colPart] || '').trim();
          if (nombre) participantes[nombre] = true;
        }
        var lista = Object.keys(participantes).sort();
        for (var p = 0; p < lista.length; p++) {
          hoja.getRange(p + 2, 1).setValue(lista[p]);
        }
      }
    }

    // Validación: solo permitir "X" o vacío en las columnas de días
    var regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(['X', ''], true)
      .setAllowInvalid(false)
      .setHelpText('Escribe X si ese día es de estudio, déjalo vacío si no')
      .build();
    hoja.getRange(2, 2, 50, 7).setDataValidation(regla);

    // Formato de fecha en columnas I y J
    hoja.getRange(2, 9, 50, 2).setNumberFormat('dd/MM/yyyy');

    // Ancho de columnas
    hoja.setColumnWidth(1, 200);
    for (var c = 2; c <= 8; c++) {
      hoja.setColumnWidth(c, 100);
    }
    hoja.setColumnWidth(9, 120);
    hoja.setColumnWidth(10, 120);

    hoja.getRange(2, 2, 50, 7).setHorizontalAlignment('center');

    SpreadsheetApp.getUi().alert(
      '📚 HOJA DE DÍAS DE ESTUDIO CREADA\n\n' +
      'Instrucciones:\n' +
      '1. En "Participante" escribe el nombre exacto como aparece en Kobo\n' +
      '2. Marca con "X" los días que esa persona tiene clase\n' +
      '3. En "Fecha Inicio" pon cuándo empiezan las clases (ej: 01/01/2026)\n' +
      '4. En "Fecha Fin" pon cuándo terminan las clases (ej: 30/11/2026)\n' +
      '5. Si no pones fechas, aplica siempre\n\n' +
      'Los días marcados = 0 horas trabajadas, 0 pago'
    );
  }

  hoja.activate();
}

// Lee la hoja DiasEstudio y devuelve un mapa: { "participante": [0,1,0,1,0,0,0] } (Lun-Dom)
function obtenerDiasEstudio() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("DiasEstudio");
  var mapa = {};

  if (!hoja) return mapa;

  var datos = hoja.getDataRange().getValues();
  // Columnas: A=Participante, B-H=Lun-Dom, I=Fecha Inicio, J=Fecha Fin
  for (var f = 1; f < datos.length; f++) {
    var participante = String(datos[f][0] || '').trim();
    if (!participante) continue;

    // Columnas 1-7 corresponden a Lunes(1), Martes(2), Miércoles(3), Jueves(4), Viernes(5), Sábado(6), Domingo(0)
    // Mapeamos a índice JS de día de semana: 0=Domingo, 1=Lunes, ..., 6=Sábado
    var dias = {};
    dias[1] = String(datos[f][1] || '').trim().toUpperCase() === 'X'; // Lunes
    dias[2] = String(datos[f][2] || '').trim().toUpperCase() === 'X'; // Martes
    dias[3] = String(datos[f][3] || '').trim().toUpperCase() === 'X'; // Miércoles
    dias[4] = String(datos[f][4] || '').trim().toUpperCase() === 'X'; // Jueves
    dias[5] = String(datos[f][5] || '').trim().toUpperCase() === 'X'; // Viernes
    dias[6] = String(datos[f][6] || '').trim().toUpperCase() === 'X'; // Sábado
    dias[0] = String(datos[f][7] || '').trim().toUpperCase() === 'X'; // Domingo

    // Fechas de vigencia (columnas I y J) — opcionales
    var fechaInicio = datos[f][8] ? new Date(datos[f][8]) : null;
    var fechaFin = datos[f][9] ? new Date(datos[f][9]) : null;

    // Si hay fecha fin, ajustar al final del día
    if (fechaFin) {
      fechaFin.setHours(23, 59, 59, 999);
    }
    // Si hay fecha inicio, ajustar al inicio del día
    if (fechaInicio) {
      fechaInicio.setHours(0, 0, 0, 0);
    }

    mapa[participante] = {
      dias: dias,
      fechaInicio: fechaInicio,
      fechaFin: fechaFin
    };
  }

  return mapa;
}

// Verifica si una fecha es día de estudio para un participante
function esDiaDeEstudio(participante, fecha, diasEstudioMapa) {
  if (!diasEstudioMapa[participante]) return false;

  var config = diasEstudioMapa[participante];

  // Verificar si la fecha está dentro del rango de vigencia
  if (config.fechaInicio && fecha < config.fechaInicio) return false;
  if (config.fechaFin && fecha > config.fechaFin) return false;

  var diaSemana = fecha.getDay(); // 0=Domingo, 1=Lunes, ...
  return config.dias[diaSemana] === true;
}

// ==================== FUNCIONES DE REPORTES ====================
function generarReportePorDia() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.prompt('Reporte por Día', 'Ingresa la fecha (dd/mm/yyyy):', ui.ButtonSet.OK_CANCEL);

  if (respuesta.getSelectedButton() == ui.Button.OK) {
    try {
      var partes = respuesta.getResponseText().split('/');
      var fecha = new Date(partes[2], partes[1] - 1, partes[0]);
      generarReporte('dia', fecha, null);
    } catch (e) {
      ui.alert('Error: Formato incorrecto. Usa dd/mm/yyyy');
    }
  }
}

function generarReportePorSemana() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.prompt('Reporte por Semana', 'Ingresa la fecha de inicio (dd/mm/yyyy):', ui.ButtonSet.OK_CANCEL);

  if (respuesta.getSelectedButton() == ui.Button.OK) {
    try {
      var partes = respuesta.getResponseText().split('/');
      var fecha = new Date(partes[2], partes[1] - 1, partes[0]);
      generarReporte('semana', fecha, null);
    } catch (e) {
      ui.alert('Error: Formato incorrecto. Usa dd/mm/yyyy');
    }
  }
}

function generarReportePorMes() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.prompt('Reporte por Mes', 'Ingresa mes y año (mm/yyyy):', ui.ButtonSet.OK_CANCEL);

  if (respuesta.getSelectedButton() == ui.Button.OK) {
    try {
      var partes = respuesta.getResponseText().split('/');
      var fecha = new Date(partes[1], partes[0] - 1, 1);
      generarReporte('mes', fecha, null);
    } catch (e) {
      ui.alert('Error: Formato incorrecto. Usa mm/yyyy');
    }
  }
}

function generarReportePorRango() {
  var ui = SpreadsheetApp.getUi();

  var respuestaInicio = ui.prompt('Reporte por Rango - Paso 1', 'Ingresa FECHA INICIO (dd/mm/yyyy):', ui.ButtonSet.OK_CANCEL);
  if (respuestaInicio.getSelectedButton() != ui.Button.OK) return;

  var respuestaFin = ui.prompt('Reporte por Rango - Paso 2', 'Ingresa FECHA FIN (dd/mm/yyyy):', ui.ButtonSet.OK_CANCEL);
  if (respuestaFin.getSelectedButton() != ui.Button.OK) return;

  try {
    var partesInicio = respuestaInicio.getResponseText().split('/');
    var fechaInicio = new Date(partesInicio[2], partesInicio[1] - 1, partesInicio[0]);
    fechaInicio.setHours(0, 0, 0, 0);

    var partesFin = respuestaFin.getResponseText().split('/');
    var fechaFin = new Date(partesFin[2], partesFin[1] - 1, partesFin[0]);
    fechaFin.setHours(23, 59, 59, 999);

    if (fechaInicio > fechaFin) {
      ui.alert('Error: La fecha de inicio debe ser menor o igual a la de fin');
      return;
    }

    generarReporte('rango', fechaInicio, fechaFin);
  } catch (e) {
    ui.alert('Error: Formato incorrecto. Usa dd/mm/yyyy');
  }
}

function generarReporteTodo() {
  generarReporte('todo', null, null);
}

// ==================== FUNCIÓN PRINCIPAL - MEJORADA ====================
function generarReporte(tipo, fechaInicio, fechaFin) {
  var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DatosKobo");

  if (!hojaActual) {
    SpreadsheetApp.getUi().alert('ERROR\n\nNo existe la hoja "DatosKobo".\nPrimero importa los datos desde KoboToolbox.');
    return;
  }

  var datos = hojaActual.getDataRange().getValues();
  var encabezados = datos[0];

  // Buscar columnas
  var colStart = -1;
  var colEnd = -1;
  var colParticipante = -1;
  var colIngreso = -1;
  var colEgreso = -1;
  var colTerapia = -1;
  var colComputacion = -1;
  var colPermiso = -1;

  for (var i = 0; i < encabezados.length; i++) {
    var nombreColumna = String(encabezados[i]).trim();

    if (nombreColumna === 'start') {
      colStart = i;
    } else if (nombreColumna === 'end') {
      colEnd = i;
    } else if (nombreColumna === 'Participante') {
      colParticipante = i;
    } else if (nombreColumna === 'Ingreso / Egreso/Ingreso') {
      colIngreso = i;
    } else if (nombreColumna === 'Ingreso / Egreso/Egreso') {
      colEgreso = i;
    } else if (nombreColumna === 'Ingreso / Egreso/Terapia (se paga el 100%)') {
      colTerapia = i;
    } else if (nombreColumna === 'Ingreso / Egreso/Computación (se paga el 50%)') {
      colComputacion = i;
    } else if (nombreColumna === 'Ingreso / Egreso/Permiso') {
      colPermiso = i;
    }
  }

  if (colStart === -1 || colParticipante === -1) {
    SpreadsheetApp.getUi().alert('ERROR\n\nNo se encontraron las columnas esperadas (start, Participante).');
    return;
  }

  // Cargar días de estudio
  var diasEstudioMapa = obtenerDiasEstudio();

  // Crear hoja de reporte
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var periodo = obtenerTextoPeriodo(tipo, fechaInicio, fechaFin);
  var nombreHojaReporte = 'Reporte_' + new Date().getTime();
  var hoja = spreadsheet.insertSheet(nombreHojaReporte);

  // Encabezados del reporte
  var encabezadosReporte = ['Fecha', 'Día', 'Entrada', 'Salida', 'Tipo', 'Horas Trabajadas', 'Porcentaje', 'Horas a Pagar'];
  hoja.getRange(1, 1, 1, 8).setValues([encabezadosReporte]);
  hoja.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1f54a8').setFontColor('#ffffff').setHorizontalAlignment('center');
  hoja.setFrozenRows(1);

  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  var filaActual = 2;

  // Título del período
  hoja.getRange(filaActual, 1, 1, 8).merge().setValue('Período: ' + periodo);
  hoja.getRange(filaActual, 1).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setBackground('#e8eaf6');
  filaActual++;

  // Agrupar datos por empleado
  var empleadosDatos = {};

  for (var fila = 1; fila < datos.length; fila++) {
    var empleadoId = String(datos[fila][colParticipante] || '').trim();

    if (!empleadoId) continue;

    if (!empleadosDatos[empleadoId]) {
      empleadosDatos[empleadoId] = [];
    }

    empleadosDatos[empleadoId].push(fila);
  }

  var listaEmpleados = Object.keys(empleadosDatos);
  var totalGeneralHoras = 0;
  var totalGeneralPagar = 0;
  var totalRegistros = 0;
  var totalDiasEstudioGeneral = 0;

  // Procesar cada empleado
  for (var emp = 0; emp < listaEmpleados.length; emp++) {
    var empleadoId = listaEmpleados[emp];
    var filasEmpleado = empleadosDatos[empleadoId];

    // Encabezado del empleado
    filaActual++;
    var textoEstudio = '';
    if (diasEstudioMapa[empleadoId]) {
      var configEstudio = diasEstudioMapa[empleadoId];
      var diasTexto = [];
      for (var d = 1; d <= 6; d++) { if (configEstudio.dias[d]) diasTexto.push(DIAS_SEMANA[d]); }
      if (configEstudio.dias[0]) diasTexto.push(DIAS_SEMANA[0]);
      if (diasTexto.length > 0) {
        textoEstudio = ' (Estudia: ' + diasTexto.join(', ');
        if (configEstudio.fechaInicio || configEstudio.fechaFin) {
          var desde = configEstudio.fechaInicio ? Utilities.formatDate(configEstudio.fechaInicio, Session.getScriptTimeZone(), 'dd/MM/yyyy') : 'inicio';
          var hasta = configEstudio.fechaFin ? Utilities.formatDate(configEstudio.fechaFin, Session.getScriptTimeZone(), 'dd/MM/yyyy') : 'indefinido';
          textoEstudio += ' | Vigencia: ' + desde + ' - ' + hasta;
        }
        textoEstudio += ')';
      }
    }
    hoja.getRange(filaActual, 1, 1, 8).merge().setValue('👤 ' + empleadoId.toUpperCase() + textoEstudio);
    hoja.getRange(filaActual, 1).setFontWeight('bold').setFontSize(11).setBackground('#e3f2fd').setHorizontalAlignment('left');
    filaActual++;

    var totalHorasEmpleado = 0;
    var totalPagarEmpleado = 0;
    var totalDiasEstudioEmpleado = 0;

    // Procesar filas del empleado - BUSCAR INGRESOS Y EGRESOS
    for (var i = 0; i < filasEmpleado.length; i++) {
      var fila = filasEmpleado[i];

      var tieneIngreso = colIngreso !== -1 && datos[fila][colIngreso];
      var tieneEgreso = colEgreso !== -1 && datos[fila][colEgreso];
      var tieneTerapia = colTerapia !== -1 && datos[fila][colTerapia];
      var tieneComputacion = colComputacion !== -1 && datos[fila][colComputacion];
      var tienePermiso = colPermiso !== -1 && datos[fila][colPermiso];

      // Si esta fila es un egreso, buscar su ingreso correspondiente
      if (tieneEgreso && colEnd !== -1) {
        var fechaEnd = new Date(datos[fila][colEnd]);

        // Buscar el ingreso del mismo día
        var ingresoCorrespondiente = null;
        for (var j = 0; j < filasEmpleado.length; j++) {
          var filaIngreso = filasEmpleado[j];
          var esIngreso = colIngreso !== -1 && datos[filaIngreso][colIngreso];

          if (esIngreso && colStart !== -1) {
            var fechaStart = new Date(datos[filaIngreso][colStart]);

            // Verificar que sea AM
            var horaEntrada = fechaStart.getHours();
            if (horaEntrada >= 12) {
              // Es PM, ignorar
              continue;
            }

            if (esMismaFecha(fechaStart, fechaEnd)) {
              ingresoCorrespondiente = filaIngreso;
              break;
            }
          }
        }

        if (!ingresoCorrespondiente) continue;

        var fechaStart = new Date(datos[ingresoCorrespondiente][colStart]);
        var horasTrabajadas = calcularHorasTrabajadas(fechaStart, fechaEnd);

        // ===== REGLA: SOLO SI LA ENTRADA ES AM =====
        var horaEntrada = fechaStart.getHours();
        if (horaEntrada >= 12) {
          continue; // Ignorar si es PM
        }

        // Validar rango
        var enRango = validarEnRango(tipo, fechaStart, fechaInicio, fechaFin);
        if (!enRango) continue;

        // Determinar tipo de registro
        var tipoFila = 'Normal';
        var porcentajeCalc = 100;

        // Verificar si es día de estudio — no se pagan horas, no cuenta como trabajo
        if (esDiaDeEstudio(empleadoId, fechaStart, diasEstudioMapa)) {
          tipoFila = 'Día de Estudio';
          porcentajeCalc = 0;
          horasTrabajadas = 0; // No trabajó, solo marcó entrada/salida
          totalDiasEstudioEmpleado++;
        } else if (tieneTerapia) {
          tipoFila = 'Terapia';
          porcentajeCalc = 100;
        } else if (tieneComputacion) {
          tipoFila = 'Computación';
          porcentajeCalc = 50;
        } else if (tienePermiso) {
          tipoFila = 'Permiso';
          porcentajeCalc = 0;
        }

        var horasPagar = horasTrabajadas * (porcentajeCalc / 100);

        var textoFechaStart = Utilities.formatDate(fechaStart, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        var textoFechaEnd = Utilities.formatDate(fechaEnd, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

        var textoHoraEntrada = textoFechaStart.split(' ')[1] || '';
        var textoHoraSalida = textoFechaEnd.split(' ')[1] || '';

        var horaEntradaNum = parseInt(textoHoraEntrada.split(':')[0]);
        var horaSalidaNum = parseInt(textoHoraSalida.split(':')[0]);

        textoHoraEntrada += (horaEntradaNum < 12 ? ' AM' : ' PM');
        textoHoraSalida += (horaSalidaNum < 12 ? ' AM' : ' PM');

        var fecha = textoFechaStart.split(' ')[0];
        var diaSemanaTexto = DIAS_SEMANA[fechaStart.getDay()];

        hoja.getRange(filaActual, 1, 1, 8).setValues([[
          fecha,
          diaSemanaTexto,
          textoHoraEntrada,
          textoHoraSalida,
          tipoFila,
          horasTrabajadas.toFixed(2),
          porcentajeCalc + '%',
          horasPagar.toFixed(2)
        ]]);

        hoja.getRange(filaActual, 1, 1, 8).setFontSize(9).setHorizontalAlignment('center');
        hoja.getRange(filaActual, 5).setHorizontalAlignment('left');

        if (tipoFila === 'Día de Estudio') {
          hoja.getRange(filaActual, 1, 1, 8).setBackground('#e1bee7'); // Morado claro
        } else if (porcentajeCalc === 0) {
          hoja.getRange(filaActual, 1, 1, 8).setBackground('#ffebee');
        } else if (porcentajeCalc === 50) {
          hoja.getRange(filaActual, 1, 1, 8).setBackground('#fff9c4');
        }

        totalHorasEmpleado += horasTrabajadas;
        totalPagarEmpleado += horasPagar;
        filaActual++;
        totalRegistros++;
      }
    }

    // Procesar ingresos sin egreso
    for (var i = 0; i < filasEmpleado.length; i++) {
      var fila = filasEmpleado[i];
      var tieneIngreso = colIngreso !== -1 && datos[fila][colIngreso];

      if (tieneIngreso && colStart !== -1) {
        var fechaStart = new Date(datos[fila][colStart]);

        // ===== REGLA: SOLO SI LA ENTRADA ES AM =====
        var horaEntrada = fechaStart.getHours();
        if (horaEntrada >= 12) {
          continue; // Ignorar si es PM
        }

        // Verificar que no tenga egreso
        var tieneEgreso = false;
        for (var j = 0; j < filasEmpleado.length; j++) {
          var filaEgreso = filasEmpleado[j];
          var esEgreso = colEgreso !== -1 && datos[filaEgreso][colEgreso];

          if (esEgreso && colEnd !== -1) {
            var fechaEnd = new Date(datos[filaEgreso][colEnd]);
            if (esMismaFecha(fechaStart, fechaEnd)) {
              tieneEgreso = true;
              break;
            }
          }
        }

        if (tieneEgreso) continue; // Ya fue procesado arriba

        // Validar rango
        var enRango = validarEnRango(tipo, fechaStart, fechaInicio, fechaFin);
        if (!enRango) continue;

        var horasTrabajadas = HORAS_JORNADA_NORMAL;
        var porcentajeCalc = 100;
        var tipoFila = 'Normal';

        // Verificar si es día de estudio — no se pagan horas, no cuenta como trabajo
        if (esDiaDeEstudio(empleadoId, fechaStart, diasEstudioMapa)) {
          tipoFila = 'Día de Estudio';
          porcentajeCalc = 0;
          horasTrabajadas = 0; // No trabajó, solo marcó entrada
          totalDiasEstudioEmpleado++;
        }

        var horasPagar = horasTrabajadas * (porcentajeCalc / 100);

        var textoFechaStart = Utilities.formatDate(fechaStart, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        var textoHoraEntrada = textoFechaStart.split(' ')[1] || '';
        var horaEntradaNum = parseInt(textoHoraEntrada.split(':')[0]);
        textoHoraEntrada += (horaEntradaNum < 12 ? ' AM' : ' PM');

        var salidaEstimada = new Date(fechaStart.getTime() + (HORAS_JORNADA_NORMAL * 60 * 60 * 1000));
        var textoHoraSalida = Utilities.formatDate(salidaEstimada, Session.getScriptTimeZone(), 'HH:mm') + '*';
        var horaSalidaNum = parseInt(textoHoraSalida.split(':')[0]);
        textoHoraSalida = textoHoraSalida.replace('*', '') + (horaSalidaNum < 12 ? ' AM' : ' PM') + '*';

        var fecha = textoFechaStart.split(' ')[0];
        var diaSemanaTexto = DIAS_SEMANA[fechaStart.getDay()];

        hoja.getRange(filaActual, 1, 1, 8).setValues([[
          fecha,
          diaSemanaTexto,
          textoHoraEntrada,
          textoHoraSalida,
          tipoFila,
          horasTrabajadas.toFixed(2),
          porcentajeCalc + '%',
          horasPagar.toFixed(2)
        ]]);

        if (tipoFila === 'Día de Estudio') {
          hoja.getRange(filaActual, 1, 1, 8).setBackground('#e1bee7').setFontStyle('italic').setFontSize(9);
        } else {
          hoja.getRange(filaActual, 1, 1, 8).setBackground('#e8f5e9').setFontStyle('italic').setFontSize(9);
        }
        hoja.getRange(filaActual, 1, 1, 8).setHorizontalAlignment('center');
        hoja.getRange(filaActual, 5).setHorizontalAlignment('left');

        totalHorasEmpleado += horasTrabajadas;
        totalPagarEmpleado += horasPagar;
        filaActual++;
        totalRegistros++;
      }
    }

    // Subtotal empleado
    hoja.getRange(filaActual, 1, 1, 5).merge();
    hoja.getRange(filaActual, 1).setValue('SUBTOTAL ' + empleadoId.toUpperCase());
    hoja.getRange(filaActual, 1).setFontWeight('bold').setHorizontalAlignment('right').setBackground('#e8e8e8');
    hoja.getRange(filaActual, 6).setValue(totalHorasEmpleado.toFixed(2)).setFontWeight('bold').setBackground('#e8e8e8').setHorizontalAlignment('center');
    hoja.getRange(filaActual, 8).setValue(totalPagarEmpleado.toFixed(2)).setFontWeight('bold').setBackground('#e8e8e8').setHorizontalAlignment('center');
    filaActual++;

    if (totalDiasEstudioEmpleado > 0) {
      hoja.getRange(filaActual, 1, 1, 5).merge();
      hoja.getRange(filaActual, 1).setValue('📚 Días de estudio: ' + totalDiasEstudioEmpleado + ' (0 horas pagadas)');
      hoja.getRange(filaActual, 1).setFontStyle('italic').setHorizontalAlignment('right').setBackground('#e1bee7').setFontSize(8);
      filaActual++;
    }

    filaActual++;

    totalGeneralHoras += totalHorasEmpleado;
    totalGeneralPagar += totalPagarEmpleado;
    totalDiasEstudioGeneral += totalDiasEstudioEmpleado;
  }

  // RESUMEN FINAL
  filaActual++;
  hoja.getRange(filaActual, 1, 1, 8).merge();
  hoja.getRange(filaActual, 1).setValue('RESUMEN GENERAL');
  hoja.getRange(filaActual, 1).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  filaActual++;

  hoja.getRange(filaActual, 1).setValue('Total Empleados:').setFontWeight('bold');
  hoja.getRange(filaActual, 2).setValue(listaEmpleados.length);
  hoja.getRange(filaActual, 3).setValue('Total Registros:').setFontWeight('bold');
  hoja.getRange(filaActual, 4).setValue(totalRegistros);
  hoja.getRange(filaActual, 5).setValue('Días Estudio:').setFontWeight('bold');
  hoja.getRange(filaActual, 6).setValue(totalDiasEstudioGeneral);
  filaActual++;

  hoja.getRange(filaActual, 1, 1, 2).merge().setValue('TOTAL HORAS LABORADAS:');
  hoja.getRange(filaActual, 1).setFontWeight('bold').setBackground('#d0d0d0');
  hoja.getRange(filaActual, 3).setValue(totalGeneralHoras.toFixed(2) + ' hrs').setFontWeight('bold').setBackground('#d0d0d0').setHorizontalAlignment('center');
  filaActual++;

  hoja.getRange(filaActual, 1, 1, 2).merge().setValue('TOTAL HORAS A REMUNERAR:');
  hoja.getRange(filaActual, 1).setFontWeight('bold').setBackground('#c0c0c0');
  hoja.getRange(filaActual, 3).setValue(totalGeneralPagar.toFixed(2) + ' hrs').setFontWeight('bold').setBackground('#c0c0c0').setHorizontalAlignment('center');
  filaActual += 2;

  // Leyenda
  hoja.getRange(filaActual, 1, 1, 8).merge().setValue('* Salida estimada (se asumieron ' + HORAS_JORNADA_NORMAL + ' horas de trabajo)');
  hoja.getRange(filaActual, 1).setFontSize(7).setFontStyle('italic').setFontColor('#666666');
  filaActual++;

  hoja.getRange(filaActual, 1, 1, 8).merge().setValue('📚 Morado = Día de Estudio (0% pago) | 🟡 Amarillo = Computación (50%) | 🔴 Rojo = Permiso (0%)');
  hoja.getRange(filaActual, 1).setFontSize(7).setFontStyle('italic').setFontColor('#666666');
  filaActual += 2;

  // NOTA IMPORTANTE
  hoja.getRange(filaActual, 1, 1, 8).merge().setValue('⚠️ NOTA: Solo se cuentan como días válidos aquellos con entrada AM (antes de las 12:00). Las entradas PM se ignoran.');
  hoja.getRange(filaActual, 1).setFontSize(8).setFontStyle('italic').setFontColor('#d32f2f').setBackground('#ffebee');
  filaActual += 2;

  // FIRMAS
  hoja.getRange(filaActual, 1, 1, 4).merge().setValue('_________________').setHorizontalAlignment('center');
  hoja.getRange(filaActual, 5, 1, 4).merge().setValue('_________________').setHorizontalAlignment('center');
  filaActual++;
  hoja.getRange(filaActual, 1, 1, 4).merge().setValue('Elaborado por').setHorizontalAlignment('center').setFontSize(8).setFontWeight('bold');
  hoja.getRange(filaActual, 5, 1, 4).merge().setValue('Vo.Bo. Recursos Humanos').setHorizontalAlignment('center').setFontSize(8).setFontWeight('bold');
  filaActual += 2;

  hoja.getRange(filaActual, 1, 1, 8).merge().setValue('Documento generado el ' + ahora + ' • ' + NOMBRE_EMPRESA);
  hoja.getRange(filaActual, 1).setFontSize(7).setFontStyle('italic').setFontColor('#666666').setHorizontalAlignment('center');

  // Ajustar columnas
  hoja.setColumnWidth(1, 90);
  hoja.setColumnWidth(2, 90);
  hoja.setColumnWidth(3, 100);
  hoja.setColumnWidth(4, 100);
  hoja.setColumnWidth(5, 120);
  hoja.setColumnWidth(6, 100);
  hoja.setColumnWidth(7, 80);
  hoja.setColumnWidth(8, 100);

  hoja.activate();

  SpreadsheetApp.getUi().alert(
    '✅ PLANILLA GENERADA\n\n' +
    'Empleados: ' + listaEmpleados.length + '\n' +
    'Registros (solo AM): ' + totalRegistros + '\n' +
    'Horas trabajadas: ' + totalGeneralHoras.toFixed(2) + ' hrs\n' +
    'Horas a pagar: ' + totalGeneralPagar.toFixed(2) + ' hrs\n\n' +
    '⚠️ Solo se contaron entradas AM (antes de 12:00)'
  );
}

// ==================== FUNCIONES AUXILIARES ====================
function calcularHorasTrabajadas(fechaStart, fechaEnd) {
  var diferencia = fechaEnd - fechaStart;
  var horas = diferencia / (1000 * 60 * 60);
  return Math.max(0, horas);
}

function esMismaFecha(fecha1, fecha2) {
  return fecha1.getFullYear() === fecha2.getFullYear() &&
         fecha1.getMonth() === fecha2.getMonth() &&
         fecha1.getDate() === fecha2.getDate();
}

function estaEnSemana(fecha, inicioSemana) {
  var finSemana = new Date(inicioSemana);
  finSemana.setDate(finSemana.getDate() + 6);
  return fecha >= inicioSemana && fecha <= finSemana;
}

function estaEnMes(fecha, mesFecha) {
  return fecha.getFullYear() === mesFecha.getFullYear() &&
         fecha.getMonth() === mesFecha.getMonth();
}

function validarEnRango(tipo, fechaRegistro, fechaInicio, fechaFin) {
  if (tipo === 'dia' && fechaInicio) {
    return esMismaFecha(fechaRegistro, fechaInicio);
  } else if (tipo === 'semana' && fechaInicio) {
    return estaEnSemana(fechaRegistro, fechaInicio);
  } else if (tipo === 'mes' && fechaInicio) {
    return estaEnMes(fechaRegistro, fechaInicio);
  } else if (tipo === 'rango' && fechaInicio && fechaFin) {
    var fechaReg = new Date(fechaRegistro);
    fechaReg.setHours(0, 0, 0, 0);
    return fechaReg >= fechaInicio && fechaReg <= fechaFin;
  } else if (tipo === 'todo') {
    return true;
  }
  return true;
}

function obtenerTextoPeriodo(tipo, fechaInicio, fechaFin) {
  if (tipo === 'dia') {
    return Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } else if (tipo === 'semana') {
    var finSemana = new Date(fechaInicio);
    finSemana.setDate(finSemana.getDate() + 6);
    return Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
           ' al ' + Utilities.formatDate(finSemana, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } else if (tipo === 'mes') {
    return Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), 'MMMM yyyy').toUpperCase();
  } else if (tipo === 'rango') {
    return Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
           ' al ' + Utilities.formatDate(fechaFin, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } else {
    return 'TODOS LOS REGISTROS';
  }
}
