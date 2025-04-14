function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Menú Personalizado')
    .addItem('Limpiar Datos Generales', 'limpiarDatosGeneralesConConfirmacion')
    .addItem('Limpiar Todas las Hojas', 'limpiarTodasLasHojasConConfirmacion')
    .addItem('Enviar Correos', 'enviarCorreosConConfirmacion')
    .addItem('Crear Carpetas de Transporte', 'crearCarpetasDeTransporteConConfirmacion')
    .addToUi();
}

function crearCarpetasDeTransporteConConfirmacion() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirmación', '¿Está seguro de que quiere crear las carpetas de transporte?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    crearCarpetasDeTransporte();
    ui.alert('Carpetas de Transporte creadas.');
  } else {
    ui.alert('Acción cancelada.');
  }
}

function crearCarpetasDeTransporte() {
  var hojaTransporte = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DatosTransporte');
  var datosTransporte = hojaTransporte.getRange('A2:J' + hojaTransporte.getLastRow()).getValues();
  var carpetaPrincipal = DriveApp.getFoldersByName('Carpeta de transporte eulen').next();
  
  datosTransporte.forEach(function(fila) {
    var placa = fila[3];
    if (placa) {
      var folderIterator = carpetaPrincipal.getFoldersByName(placa);
      if (!folderIterator.hasNext()) {
        carpetaPrincipal.createFolder(placa);
      }
    }
  });
}

function enviarCorreosConConfirmacion() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirmación', '¿Está seguro de que quiere enviar los correos?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    enviarCorreos();
  } else {
    ui.alert('Acción cancelada.');
  }
}

function enviarCorreos() {
  var hojaGenerales = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DatosGenerales');
  var hojaTransporte = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DatosTransporte');
  var datosGenerales = hojaGenerales.getRange('A2:K' + hojaGenerales.getLastRow()).getValues();
  var datosTransporte = hojaTransporte.getRange('A2:J' + hojaTransporte.getLastRow()).getValues();

  var carpetaPrincipal = DriveApp.getFoldersByName('Carpeta de transporte eulen').next();

  var correosCc = ['EXAMPLE@DSDS.COM,']; 

  var descripcionesMap = {};
  datosGenerales.forEach(function(fila) {
    var descripcion = fila[4];
    if (!descripcionesMap[descripcion]) {
      descripcionesMap[descripcion] = [];
    }
    descripcionesMap[descripcion].push(fila);
  });

  var archivosEncontrados = false;

  for (var descripcion in descripcionesMap) {
    var filas = descripcionesMap[descripcion];
    var correos = filas[0][8].split(',').map(correo => correo.trim()); 
    var estadoActual = filas[0][10];

    if (estadoActual === "Enviado") {
      continue;
    }

    var estado = "Enviado";

    var datosTransporteFiltrados = datosTransporte.filter(function(filaTransporte) {
      return filaTransporte[3] == filas[0][9];
    })[0];

    if (datosTransporteFiltrados) {
      var tn = datosTransporteFiltrados[0];
      var tr = datosTransporteFiltrados[1];
      var transporte = datosTransporteFiltrados[2];
      var placa = datosTransporteFiltrados[3];
      var chofer = datosTransporteFiltrados[4];
      var licencia = datosTransporteFiltrados[5];
      var auxiliar = datosTransporteFiltrados[6];
      var dni = datosTransporteFiltrados[7];
      var telefono = datosTransporteFiltrados[8];
      var obsevaciones = datosTransporteFiltrados[9];

      try {
        var folderIterator = carpetaPrincipal.getFoldersByName(placa);
        var folder;
        if (folderIterator.hasNext()) {
          folder = folderIterator.next();
        } else {
          estado = 'No se encontró la carpeta para la placa: ' + placa;
          filas.forEach(function(fila, index) {
            hojaGenerales.getRange('K' + (datosGenerales.indexOf(fila) + 2)).setValue(estado);
          });
          continue;
        }

        var files = [];
        var fileIterator = folder.getFiles();
        while (fileIterator.hasNext()) {
          files.push(fileIterator.next()); 
          archivosEncontrados = true; 
        }

        if (files.length === 0) {
          estado = 'No se encontraron archivos en la carpeta para la placa: ' + placa;
          filas.forEach(function(fila, index) {
            hojaGenerales.getRange('K' + (datosGenerales.indexOf(fila) + 2)).setValue(estado);
          });
          continue;
        }
      } catch (e) {
        estado = 'Error al buscar la carpeta para la placa: ' + placa;
        filas.forEach(function(fila, index) {
          hojaGenerales.getRange('K' + (datosGenerales.indexOf(fila) + 2)).setValue(estado);
        });
        continue;
      }

      var mensajeHTML = `
        <p>Estimados</p>
        <p>Favor de dar facilidades de ingreso para la entrega de pedidos EULEN adjunto los detalles.</p>
        <table border="1" cellpadding="5" cellspacing="0" style="margin-bottom: 20px;">
          <tr>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Fecha de Despacho</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Fecha de Entrega</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Pedido</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Sector de Requerimientos</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Descripción del Sector</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Dirección</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Distrito</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Departamento</th>
          </tr>`;
          
      filas.forEach(function(fila) {
        mensajeHTML += `
          <tr>
            <td>${formatearFecha(fila[0])}</td>
            <td>${formatearFecha(fila[1])}</td>
            <td>${fila[2]}</td>
            <td>${fila[3]}</td>
            <td>${fila[4]}</td>
            <td>${fila[5]}</td>
            <td>${fila[6]}</td>
            <td>${fila[7]}</td>
          </tr>`;
      });

      mensajeHTML += `
        </table>
        <table border="1" cellpadding="5" cellspacing="0">
          <tr>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">TN</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">TR</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Transporte</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Placa</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Chofer</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Licencia</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Auxiliar</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">DNI</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">Auxiliar 2</th>
            <th style="background-color: #7F0F0F; color: #FFFFFF;">DNI.</th>
          </tr>
          <tr>
            <td>${tn}</td>
            <td>${tr}</td>
            <td>${transporte}</td>
            <td>${placa}</td>
            <td>${chofer}</td>
            <td>${licencia}</td>
            <td>${auxiliar}</td>
            <td>${dni}</td>
            <td>${telefono}</td>
            <td>${obsevaciones}</td>
          </tr>
        </table>
        <p>Saludos cordiales,<br>Anthony Jesus Marcelo</p>
      `;
      
      var asunto = 'PERMISO DE INGRESO - ' + descripcion;
      
      try {
        MailApp.sendEmail({
          to: correos.join(','),
          subject: asunto,
          htmlBody: mensajeHTML,
          attachments: files,
          cc: correosCc.join(',')
        });
      } catch (e) {
        Logger.log('Error al enviar el correo: ' + e.toString());
        estado = 'Error al enviar el correo: ' + e.toString();
      }
    } else {
      estado = 'No se encontraron datos de transporte para la placa: ' + placa;
    }

    filas.forEach(function(fila, index) {
      hojaGenerales.getRange('K' + (datosGenerales.indexOf(fila) + 2)).setValue(estado);
    });
  }

  if (!archivosEncontrados) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('No se ha encontrado ningún archivo en la carpeta.');
  }
}

function formatearFecha(fecha) {
  if (fecha instanceof Date) {
    var formato = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
    return formato;
  }
  return fecha;
}

function limpiarDatosGeneralesConConfirmacion() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirmación', '¿Está seguro de que quiere limpiar los datos generales?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    limpiarDatosGenerales();
    ui.alert('Datos Generales limpiados.');
  } else {
    ui.alert('Acción cancelada.');
  }
}

function limpiarTodasLasHojasConConfirmacion() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirmación', '¿Está seguro de que quiere limpiar todas las hojas?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    limpiarDatosGenerales();
    limpiarDatosTransporte();
    ui.alert('Todas las hojas limpiadas.');
  } else {
    ui.alert('Acción cancelada.');
  }
}

function limpiarDatosGenerales() {
  var hojaGenerales = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DatosGenerales');
  hojaGenerales.getRange('A2:K' + hojaGenerales.getLastRow()).clearContent();
}

function limpiarDatosTransporte() {
  var hojaTransporte = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DatosTransporte');
  hojaTransporte.getRange('A2:J' + hojaTransporte.getLastRow()).clearContent();
}
