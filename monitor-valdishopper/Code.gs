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
function getImageData(id) {
  var blob = DriveApp.getFileById(id).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
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
