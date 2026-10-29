function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Visor de Operaciones - Valdishopper')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

const FOLDER_ID = '1jmITWnSqY7R0D7bR9liUcvvnGF82vQD0';
const SHEET_ID = '11hnF2NI_Lzh9xza8wuEzm7dQ_xIirmwSvnLGmOoTs1s';

// ==========================================
// MÓDULO 1: GESTIÓN DE IMÁGENES Y CARRUSEL
// ==========================================
function getImagesFromFolder() {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var files = folder.getFiles();
  var images = [];

  while (files.hasNext()) {
    var file = files.next();
    var mimeType = file.getMimeType();
    var name = file.getName();

    if (mimeType.indexOf('image') !== -1 && !name.startsWith('Pantallazo')) {
      images.push({
        name: name,
        id: file.getId()
      });
    }
  }

  images.sort(function(a, b) {
    var regex = /(\d+)/g;
    var aMatch = a.name.match(regex);
    var bMatch = b.name.match(regex);

    if (aMatch && bMatch) {
        return parseInt(aMatch[0]) - parseInt(bMatch[0]);
    }
    return a.name.localeCompare(b.name);
  });

  return images;
}

// Devuelve la imagen como data URL (base64). Así el navegador no depende de
// que el archivo esté compartido públicamente ni de las cookies de Google,
// que es lo que hacía fallar el enlace drive.google.com/thumbnail.
// Primero pide a Drive una versión reducida (máx. 1600 px) para que el envío
// sea liviano; si no se puede, manda el archivo original.
function getImageData(id) {
  try {
    var headers = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
    var meta = JSON.parse(UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + id + '?fields=thumbnailLink',
      { headers: headers }).getContentText());

    if (meta.thumbnailLink) {
      var url = meta.thumbnailLink.replace(/=s\d+$/, '=s1600');
      var resp = UrlFetchApp.fetch(url, { headers: headers, muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        var thumb = resp.getBlob();
        return 'data:' + (thumb.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(thumb.getBytes());
      }
    }
  } catch (e) {
    console.warn('Miniatura no disponible para ' + id + ': ' + e);
  }

  var blob = DriveApp.getFileById(id).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

// Ejecuta esta función desde el editor (botón ▶) para revisar el carrusel:
// muestra en el registro cuántas imágenes ve y si la primera se puede leer.
function probarCarrusel() {
  var images = getImagesFromFolder();
  console.log('Imágenes encontradas: ' + images.length + ' -> ' + images.map(function(i) { return i.name; }).join(', '));
  if (images.length) {
    var data = getImageData(images[0].id);
    console.log('Primera imagen OK, tamaño enviado: ' + Math.round(data.length / 1024) + ' KB');
  }
}

function uploadImage(base64Data, filename) {
  try {
    var folder = DriveApp.getFolderById(FOLDER_ID);

    var splitBase = base64Data.split(',');
    var contentType = splitBase[0].split(';')[0].replace('data:', '');
    var byteCharacters = Utilities.base64Decode(splitBase[1]);

    var blob = Utilities.newBlob(byteCharacters, contentType, filename);
    var file = folder.createFile(blob);

    return { success: true, name: file.getName() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// MÓDULO 2: DATOS DE OPERACIÓN (ASIGNACIONES)
// ==========================================
function getData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const unassigned = [];
  const assigned = [];

  for (let i = 1; i < data.length; i++) {
    const nombre = data[i][0];
    const sg = data[i][1];
    let hora = data[i][2];

    if (nombre === "") continue;

    if (hora instanceof Date) {
      hora = Utilities.formatDate(hora, Session.getScriptTimeZone(), "HH:mm");
    } else if (!hora) {
      hora = "--:--";
    }

    if (sg === "") {
      unassigned.push({ nombre: String(nombre) });
    } else {
      assigned.push({ nombre: String(nombre), sg: String(sg), hora: String(hora) });
    }
  }

  // Todos los asignados, el más reciente primero
  const recentAssigned = assigned.reverse();

  return {
    unassigned: unassigned,
    recentAssigned: recentAssigned
  };
}
