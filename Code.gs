// ==================== CONFIGURACIÓN ====================
var NOMBRE_EMPRESA = "mi_eelo";
var NIT_EMPRESA = "NIT: 12345678-9";
var DIRECCION = "Creamos Guatemala";
var HORAS_JORNADA_NORMAL = 7; // Horas de trabajo por día

// URL de tu exportación de KoboToolbox
var URL_KOBO = "https://kf.kobotoolbox.org/api/v2/assets/agi395bJj6ojXJzPPDT9n6/export-settings/es4oUjEmPvovgLd6Y5yrQ4K/data.csv";

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
    .addToUi();

  // Actualizar automáticamente al abrir
  importarCSVdesdeKobo();
}

// ==================== IMPORTACIÓN DESDE KOBO ====================
function importarCSVdesdeKobo() {
  try {
    var response = UrlFetchApp.fetch(URL_KOBO);
    var csv = response.getContentText();
    var datos = Utilities.parseCsv(csv, ";");

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = spreadsheet.getSheetByName("DatosKobo");

    if (!hoja) {
      hoja = spreadsheet.insertSheet("DatosKobo");
    }

    hoja.clearContents();

    if (datos.length > 0) {
      hoja.getRange(1, 1, datos.length, datos[0].length).setValues(datos);

      // Formatear encabezados
      hoja.getRange(1, 1, 1, datos[0].length)
        .setFontWeight('bold')
        .setBackground('#4a86e8')
        .setFontColor('#ffffff');

      hoja.setFrozenRows(1);

      SpreadsheetApp.getActiveSpreadsheet().toast('✅ Datos actualizados desde KoboToolbox', 'Importación Exitosa', 3);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error al importar datos: ' + e.message);
  }
}

// Configurar trigger automático cada hora
function configurarActualizacionAutomatica() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'importarCSVdesdeKobo') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('importarCSVdesdeKobo')
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp.getUi().alert('✅ Actualización automática configurada cada hora');
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

  // Crear hoja de reporte
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var nombreHojaReporte = 'Reporte_' + new Date().getTime();
  var hoja = spreadsheet.insertSheet(nombreHojaReporte);

  // Encabezados del reporte
  var encabezadosReporte = ['Fecha', 'Entrada', 'Salida', 'Tipo', 'Horas Trabajadas', 'Porcentaje', 'Horas a Pagar'];
  hoja.getRange(1, 1, 1, 7).setValues([encabezadosReporte]);
  hoja.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1f54a8').setFontColor('#ffffff').setHorizontalAlignment('center');
  hoja.setFrozenRows(1);

  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  var filaActual = 2;

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

  // Procesar cada empleado
  for (var emp = 0; emp < listaEmpleados.length; emp++) {
    var empleadoId = listaEmpleados[emp];
    var filasEmpleado = empleadosDatos[empleadoId];

    // Encabezado del empleado
    filaActual++;
    hoja.getRange(filaActual, 1, 1, 7).merge().setValue('👤 ' + empleadoId.toUpperCase());
    hoja.getRange(filaActual, 1).setFontWeight('bold').setFontSize(11).setBackground('#e3f2fd').setHorizontalAlignment('left');
    filaActual++;

    var totalHorasEmpleado = 0;
    var totalPagarEmpleado = 0;

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

        var tipoFila = 'Normal';
        var porcentajeCalc = 100;

        if (tieneIngreso) {
          tipoFila = 'Ingreso';
          porcentajeCalc = 100;
        } else if (tieneEgreso) {
          tipoFila = 'Egreso';
          porcentajeCalc = 100;
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

        hoja.getRange(filaActual, 1, 1, 7).setValues([[
          fecha,
          textoHoraEntrada,
          textoHoraSalida,
          tipoFila,
          horasTrabajadas.toFixed(2),
          porcentajeCalc + '%',
          horasPagar.toFixed(2)
        ]]);

        hoja.getRange(filaActual, 1, 1, 7).setFontSize(9).setHorizontalAlignment('center');
        hoja.getRange(filaActual, 4).setHorizontalAlignment('left');

        if (porcentajeCalc === 0) {
          hoja.getRange(filaActual, 1, 1, 7).setBackground('#ffebee');
        } else if (porcentajeCalc === 50) {
          hoja.getRange(filaActual, 1, 1, 7).setBackground('#fff9c4');
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
        var horasPagar = horasTrabajadas;

        var textoFechaStart = Utilities.formatDate(fechaStart, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        var textoHoraEntrada = textoFechaStart.split(' ')[1] || '';
        var horaEntradaNum = parseInt(textoHoraEntrada.split(':')[0]);
        textoHoraEntrada += (horaEntradaNum < 12 ? ' AM' : ' PM');

        var salidaEstimada = new Date(fechaStart.getTime() + (HORAS_JORNADA_NORMAL * 60 * 60 * 1000));
        var textoHoraSalida = Utilities.formatDate(salidaEstimada, Session.getScriptTimeZone(), 'HH:mm') + '*';
        var horaSalidaNum = parseInt(textoHoraSalida.split(':')[0]);
        textoHoraSalida = textoHoraSalida.replace('*', '') + (horaSalidaNum < 12 ? ' AM' : ' PM') + '*';

        var fecha = textoFechaStart.split(' ')[0];

        hoja.getRange(filaActual, 1, 1, 7).setValues([[
          fecha,
          textoHoraEntrada,
          textoHoraSalida,
          'Normal',
          horasTrabajadas.toFixed(2),
          '100%',
          horasPagar.toFixed(2)
        ]]);

        hoja.getRange(filaActual, 1, 1, 7).setBackground('#e8f5e9').setFontStyle('italic').setFontSize(9);
        hoja.getRange(filaActual, 1, 1, 7).setHorizontalAlignment('center');
        hoja.getRange(filaActual, 4).setHorizontalAlignment('left');

        totalHorasEmpleado += horasTrabajadas;
        totalPagarEmpleado += horasPagar;
        filaActual++;
        totalRegistros++;
      }
    }

    // Subtotal empleado
    hoja.getRange(filaActual, 1, 1, 4).merge();
    hoja.getRange(filaActual, 1).setValue('SUBTOTAL ' + empleadoId.toUpperCase());
    hoja.getRange(filaActual, 1).setFontWeight('bold').setHorizontalAlignment('right').setBackground('#e8e8e8');
    hoja.getRange(filaActual, 5).setValue(totalHorasEmpleado.toFixed(2)).setFontWeight('bold').setBackground('#e8e8e8').setHorizontalAlignment('center');
    hoja.getRange(filaActual, 7).setValue(totalPagarEmpleado.toFixed(2)).setFontWeight('bold').setBackground('#e8e8e8').setHorizontalAlignment('center');
    filaActual += 2;

    totalGeneralHoras += totalHorasEmpleado;
    totalGeneralPagar += totalPagarEmpleado;
  }

  // RESUMEN FINAL
  filaActual++;
  hoja.getRange(filaActual, 1, 1, 7).merge();
  hoja.getRange(filaActual, 1).setValue('RESUMEN GENERAL');
  hoja.getRange(filaActual, 1).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  filaActual++;

  hoja.getRange(filaActual, 1).setValue('Total Empleados:').setFontWeight('bold');
  hoja.getRange(filaActual, 2).setValue(listaEmpleados.length);
  hoja.getRange(filaActual, 3).setValue('Total Registros:').setFontWeight('bold');
  hoja.getRange(filaActual, 4).setValue(totalRegistros);
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
  hoja.getRange(filaActual, 1, 1, 7).merge().setValue('* Salida estimada (se asumieron ' + HORAS_JORNADA_NORMAL + ' horas de trabajo)');
  hoja.getRange(filaActual, 1).setFontSize(7).setFontStyle('italic').setFontColor('#666666');
  filaActual += 2;

  // NOTA IMPORTANTE
  hoja.getRange(filaActual, 1, 1, 7).merge().setValue('⚠️ NOTA: Solo se cuentan como días válidos aquellos con entrada AM (antes de las 12:00). Las entradas PM se ignoran.');
  hoja.getRange(filaActual, 1).setFontSize(8).setFontStyle('italic').setFontColor('#d32f2f').setBackground('#ffebee');
  filaActual += 2;

  // FIRMAS
  hoja.getRange(filaActual, 1, 1, 3).merge().setValue('_________________').setHorizontalAlignment('center');
  hoja.getRange(filaActual, 5, 1, 3).merge().setValue('_________________').setHorizontalAlignment('center');
  filaActual++;
  hoja.getRange(filaActual, 1, 1, 3).merge().setValue('Elaborado por').setHorizontalAlignment('center').setFontSize(8).setFontWeight('bold');
  hoja.getRange(filaActual, 5, 1, 3).merge().setValue('Vo.Bo. Recursos Humanos').setHorizontalAlignment('center').setFontSize(8).setFontWeight('bold');
  filaActual += 2;

  hoja.getRange(filaActual, 1, 1, 7).merge().setValue('Documento generado el ' + ahora + ' • ' + NOMBRE_EMPRESA);
  hoja.getRange(filaActual, 1).setFontSize(7).setFontStyle('italic').setFontColor('#666666').setHorizontalAlignment('center');

  // Ajustar columnas
  hoja.setColumnWidth(1, 90);
  hoja.setColumnWidth(2, 130);
  hoja.setColumnWidth(3, 130);
  hoja.setColumnWidth(4, 120);
  hoja.setColumnWidth(5, 100);
  hoja.setColumnWidth(6, 80);
  hoja.setColumnWidth(7, 100);

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
