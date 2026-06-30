/**
 * =============================================================================
 *  SMART ROUTER V26.3 - FINAL STABLE (AUTO-SAVE ADDED)
 * =============================================================================
 *
 *  ¿QUÉ ES ESTE ARCHIVO?
 *  ---------------------
 *  Este es el "cerebro" del sistema (el código que corre en los servidores de
 *  Google, NO en el navegador del usuario). Está escrito en Google Apps Script,
 *  que es básicamente JavaScript con permisos especiales para usar Google Drive
 *  y Google Sheets.
 *
 *  ¿PARA QUÉ SIRVE EL SISTEMA?
 *  ---------------------------
 *  Es una herramienta de optimización de rutas de despacho. El usuario sube un
 *  archivo Excel con los pedidos del día (clientes, direcciones, coordenadas) y
 *  el sistema los agrupa en rutas de reparto para los distintos móviles/camiones.
 *
 *  RESPONSABILIDADES PRINCIPALES DE ESTE ARCHIVO:
 *  - Recibir y leer el archivo Excel que sube el usuario.
 *  - Extraer los pedidos según la sucursal y el horario elegido.
 *  - Generar el archivo Excel final con las rutas ya asignadas.
 *  - Guardar automáticamente una copia en Google Drive.
 *
 *  CARACTERÍSTICAS DESTACADAS DE ESTA VERSIÓN:
 *  - Diseño: Valdishopper Dark/Pink (la estética visual de la interfaz).
 *  - Lógica Bultos: extracción del número de bultos mediante "Regex".
 *  - Lógica Picking: determina si un pedido está armado usando "FECHA DE PICKING".
 *  - Excel: corrige la suma de pendientes y separa la comuna de la dirección.
 *  - Feature: auto-guardado del resultado en Google Drive.
 * =============================================================================
 */


/* -----------------------------------------------------------------------------
 *  CONSTANTES GLOBALES (valores fijos que usa todo el programa)
 * -------------------------------------------------------------------------- */

// Identificador único de la carpeta de Google Drive donde se leen y guardan los
// archivos Excel. Google le asigna este código a cada carpeta.
const FOLDER_ID = '1H0gLR58Dcs_xJ0zjXtBlFJ55e2W5XrIi';

// Lista de los bloques horarios válidos de despacho. Solo se consideran los
// pedidos cuya hora de entrega coincida con alguno de estos horarios.
const HORARIOS_FIJOS = ["10:15", "11:15", "12:15", "13:15", "14:15", "15:15", "16:15", "17:15", "18:15", "19:15"];

// Paleta de colores suaves (pasteles). Se usa para pintar cada ruta de un color
// distinto en el Excel final, y así diferenciarlas visualmente de un vistazo.
const COLORES_PASTELES = ["#f4cccc", "#cfe2f3", "#d9ead3", "#fff2cc", "#fce5cd", "#d9d2e9", "#d0e0e3", "#ead1dc"];

// Color gris para la categoría especial "LAT" (pedidos laterales / aparte) en el Excel.
const COLOR_LAT_EXCEL = "#e0e0e0";


/**
 * doGet()
 * -------
 * Función especial de Google Apps Script: se ejecuta automáticamente cuando
 * alguien abre la URL de la aplicación web en su navegador.
 *
 * Su única tarea es "servir" (entregar) la página web (el archivo Index.html),
 * que es la interfaz visual con la que interactúa el usuario.
 *
 * El método setXFrameOptionsMode(...ALLOWALL) permite que esta página pueda
 * mostrarse incrustada dentro de otra página web (por ejemplo en un iframe).
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('VS RUTAS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * cleanTxt(t)
 * -----------
 * Función auxiliar para "limpiar" y normalizar texto. Sirve para comparar
 * textos sin que las tildes, mayúsculas o espacios causen problemas.
 *
 * Paso a paso, sobre el texto recibido:
 *   1. Si no llega nada, devuelve un texto vacío "".
 *   2. Lo convierte todo a MAYÚSCULAS.
 *   3. Le quita las tildes y acentos (ej: "DIRECCIÓN" -> "DIRECCION").
 *   4. Elimina los espacios sobrantes al inicio y al final.
 *
 * Ejemplo: "  Dirección  "  ->  "DIRECCION"
 */
function cleanTxt(t) {
  if (!t) return "";
  return t.toString().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}


/* =============================================================================
 *  PROCESAR SUBIDA DEL ARCHIVO
 * ========================================================================== */

/**
 * processUpload(fileData, fileName)
 * ---------------------------------
 * Se ejecuta cuando el usuario sube un archivo Excel desde la interfaz.
 *
 * ¿Qué hace en resumen?
 *   - Recibe el Excel, lo guarda en Drive y lo lee.
 *   - Cuenta cuántos pedidos hay por cada sucursal (LOCAL) y por cada horario.
 *   - Devuelve ese conteo para llenar los menús desplegables de la interfaz.
 *
 * Parámetros:
 *   - fileData: el contenido del archivo codificado en texto (formato Base64).
 *   - fileName: el nombre original del archivo.
 *
 * Devuelve un objeto con:
 *   - mapaConteos: cantidad de pedidos por sucursal y horario.
 *   - locales: la lista de sucursales encontradas.
 *   - archivoUsado: el nombre del archivo procesado.
 * (Si algo falla, devuelve un objeto con la propiedad "error".)
 */
function processUpload(fileData, fileName) {
  try {
    // 1. Apunta a la carpeta de Drive donde se trabajará.
    const folder = DriveApp.getFolderById(FOLDER_ID);

    // 2. Convierte el texto Base64 recibido de vuelta a un archivo Excel real (un "blob" = bloque de datos binarios).
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData.split(',')[1]), MimeType.MICROSOFT_EXCEL, fileName);

    // 3. Guarda ese Excel en la carpeta de Drive.
    const newFile = folder.createFile(blob);

    // 4. Crea una COPIA TEMPORAL convertida a Google Sheets, porque Apps Script
    //    no puede leer celdas directamente de un .xlsx, pero sí de un Google Sheet.
    const tempFile = Drive.Files.create({name: "TEMP_UPLOAD", parents: [FOLDER_ID], mimeType: 'application/vnd.google-apps.spreadsheet'}, newFile.getBlob());

    // 5. Lee TODAS las celdas de la primera hoja como una matriz (filas y columnas).
    const data = SpreadsheetApp.openById(tempFile.id).getSheets()[0].getDataRange().getValues();

    // 6. Borra la copia temporal: ya tenemos los datos en memoria, no la necesitamos.
    Drive.Files.remove(tempFile.id);

    // 7. Toma la PRIMERA fila (los títulos/encabezados de columna) y los limpia
    //    con cleanTxt para poder buscarlos sin problemas de tildes o mayúsculas.
    const headers = data[0].map(cleanTxt);

    // 8. Averigua en qué número de columna están los datos que nos interesan.
    //    indexOf devuelve la posición; si la columna no existe, devuelve -1.
    const idxLocal = headers.indexOf("LOCAL");                // Columna de la sucursal.
    const idxMin = headers.indexOf("TIEMPO MIN ENTREGA");     // Columna de la hora de entrega.

    // Si no existe la columna "LOCAL", no podemos continuar: lanzamos un error.
    if (idxLocal === -1) throw new Error("No se encontró columna LOCAL");

    // 9. Recorremos fila por fila para contar pedidos por sucursal y horario.
    //    mapaConteos terminará con una estructura tipo: { "L58": { "10:15": 4, "11:15": 7 } }
    let mapaConteos = {};
    for (let i = 1; i < data.length; i++) { // Empieza en 1 para saltarse la fila de títulos.

      // Lee el código de la sucursal y le quita el ".0" que a veces agrega Excel.
      let lRaw = data[i][idxLocal] ? data[i][idxLocal].toString().replace('.0','').trim() : "";
      if (!lRaw) continue; // Si la fila no tiene sucursal, la saltamos.

      // Normaliza el formato del código: si no empieza con "L", se la agregamos (ej: "58" -> "L58").
      let lId = lRaw.startsWith('L') ? lRaw : 'L' + lRaw;

      // Extrae la hora de entrega. Puede venir como fecha real o como texto, así que manejamos ambos casos.
      let valHora = data[i][idxMin], rowTime = "";
      if (valHora instanceof Date) {
        // Si es una fecha, la formateamos para quedarnos solo con "HH:mm" (hora:minutos).
        rowTime = Utilities.formatDate(valHora, "GMT", "HH:mm");
      } else {
        // Si es texto, recortamos para quedarnos con los primeros 5 caracteres de la hora.
        let tStr = valHora.toString().trim();
        rowTime = tStr.includes(" ") ? tStr.split(" ")[1].substring(0, 5) : tStr.substring(0, 5);
      }

      // Solo contamos el pedido si su horario es uno de los bloques válidos.
      if (HORARIOS_FIJOS.includes(rowTime)) {
        if (!mapaConteos[lId]) mapaConteos[lId] = {};                       // Si es la primera vez que vemos esta sucursal, la inicializamos.
        mapaConteos[lId][rowTime] = (mapaConteos[lId][rowTime] || 0) + 1;   // Sumamos 1 al conteo de ese horario.
      }
    }

    // 10. Obtenemos la lista de sucursales encontradas y la ordenamos alfabéticamente.
    const locs = Object.keys(mapaConteos).sort();

    // 11. Devolvemos los resultados a la interfaz.
    return { mapaConteos: mapaConteos, locales: locs, archivoUsado: fileName };

  } catch(e) {
    // Si ocurre cualquier error en el proceso, lo devolvemos como texto para mostrarlo en pantalla.
    return { error: e.toString() };
  }
}


/**
 * obtenerUltimoMaestro(folder)
 * ----------------------------
 * Busca dentro de la carpeta de Drive el archivo Excel "maestro" más reciente
 * (el último que se subió o modificó).
 *
 * Reglas de selección:
 *   - Solo considera archivos .xlsx o .xls.
 *   - IGNORA los archivos cuyo nombre empieza con "DESPACHO", porque esos son
 *     resultados generados por el propio sistema, no archivos de origen.
 *
 * Devuelve el archivo más recientemente actualizado que cumpla las reglas.
 */
function obtenerUltimoMaestro(folder) {
  const files = folder.getFiles();             // Obtiene un "iterador" para recorrer todos los archivos de la carpeta.
  let latestFile = null, latestTime = 0;       // Guardará el archivo ganador y su fecha (en milisegundos).

  while (files.hasNext()) {                     // Mientras queden archivos por revisar...
    let f = files.next();                       // Toma el siguiente archivo.
    let name = f.getName().toUpperCase();       // Su nombre en mayúsculas, para comparar sin importar mayúsculas/minúsculas.

    // Nos quedamos solo con Excel que NO sean resultados "DESPACHO".
    if (!name.startsWith("DESPACHO") && (name.endsWith(".XLSX") || name.endsWith(".XLS"))) {
      let updateTime = f.getLastUpdated().getTime();   // Fecha de última modificación en milisegundos.
      // Si este archivo es más nuevo que el mejor encontrado hasta ahora, pasa a ser el ganador.
      if (updateTime > latestTime) { latestTime = updateTime; latestFile = f; }
    }
  }
  return latestFile;
}


/* =============================================================================
 *  OBTENER PEDIDOS
 * ========================================================================== */

/**
 * getOrders(localId, horasSeleccionadas)
 * --------------------------------------
 * Devuelve la lista detallada de pedidos para una sucursal y unos horarios
 * específicos. Esta lista es la que luego se pinta como puntos en el mapa.
 *
 * Parámetros:
 *   - localId: la sucursal elegida (ej: "L58").
 *   - horasSeleccionadas: uno o varios bloques horarios elegidos por el usuario.
 *
 * Para cada pedido devuelve, entre otras cosas: número de orden, cliente,
 * dirección, comuna, coordenadas (latitud/longitud), cantidad de bultos y si
 * está armado o no.
 */
function getOrders(localId, horasSeleccionadas) {
  // 1. Localiza la carpeta y el último Excel maestro disponible.
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const excelFile = obtenerUltimoMaestro(folder);

  // 2. Igual que antes: crea una copia temporal en formato Google Sheets para poder leerla.
  const tempFile = Drive.Files.create({name: "T_Q", parents: [FOLDER_ID], mimeType: 'application/vnd.google-apps.spreadsheet'}, excelFile.getBlob());
  const data = SpreadsheetApp.openById(tempFile.id).getSheets()[0].getDataRange().getValues();
  Drive.Files.remove(tempFile.id); // Borramos la copia temporal.

  // 3. Limpiamos los encabezados para poder ubicar las columnas.
  const headers = data[0].map(cleanTxt);

  // 4. Ubicamos la posición de cada columna que necesitamos.
  const idxLocal = headers.indexOf("LOCAL");                 // Sucursal.
  const idxOrden = headers.indexOf("ORDEN");                 // Número de orden / pedido.
  const idxCliente = headers.indexOf("NOMBRE CLIENTE");      // Nombre del cliente.
  const idxDirFull = headers.indexOf("DIRECCION CLIENTE");   // Dirección completa (calle + comuna).
  const idxMin = headers.indexOf("TIEMPO MIN ENTREGA");      // Hora de entrega.
  const idxLat = headers.indexOf("LATITUD DIRECCION");       // Latitud (coordenada geográfica norte-sur).
  const idxLon = headers.indexOf("LONGITUD DIRECCION");      // Longitud (coordenada geográfica este-oeste).
  const idxBultos = headers.indexOf("DETALLE BULTOS");       // Texto con el detalle de bultos.
  const idxFechaPick = headers.indexOf("FECHA DE PICKING");  // Fecha en que se armó (picking) el pedido.

  // 5. Preparamos los criterios de filtrado.
  //    Al código de sucursal le quitamos la "L" para compararlo con el del Excel (que viene sin "L").
  const target = localId.replace('L', '').trim();
  //    Nos aseguramos de que los horarios sean siempre una lista, aunque venga uno solo.
  const targetHoras = Array.isArray(horasSeleccionadas) ? horasSeleccionadas : [horasSeleccionadas];

  // 6. Procesamos las filas en tres pasos encadenados: filter -> map -> filter.
  const orders = data.slice(1)  // .slice(1) descarta la fila de títulos y deja solo los datos.

    // --- PASO A: FILTRAR ---
    // Nos quedamos únicamente con las filas de la sucursal y horarios solicitados.
    .filter(row => {
      let rLoc = row[idxLocal] ? row[idxLocal].toString().replace('.0','').trim() : ""; // Sucursal de la fila.

      // Extrae la hora de la fila (mismo manejo de fecha/texto que en processUpload).
      let vH = row[idxMin], rTime = "";
      if (vH instanceof Date) rTime = Utilities.formatDate(vH, "GMT", "HH:mm");
      else { let tStr = vH.toString().trim(); rTime = tStr.includes(" ") ? tStr.split(" ")[1].substring(0, 5) : tStr.substring(0, 5); }

      // La fila pasa el filtro si coincide la sucursal Y la hora está entre las elegidas.
      return rLoc == target && targetHoras.includes(rTime);
    })

    // --- PASO B: TRANSFORMAR (map) ---
    // Convertimos cada fila cruda del Excel en un objeto ordenado y fácil de usar.
    .map(row => {

      // --- LÓGICA DE DIRECCIÓN Y COMUNA ---
      // La dirección viene completa en un solo campo, ej: "Pje. Uno 2457, Hualpén".
      // La separamos por la coma para obtener calle y comuna por separado.
      let fullText = (row[idxDirFull] || "").toString();
      let parts = fullText.split(",").map(p => p.trim());

      let calleLimpia = parts[0] || "";                       // Primer trozo = la calle (ej: "Pje. Uno 2457").
      let comunaLimpia = (parts.length > 1) ? parts[1] : "";  // Segundo trozo = la comuna (ej: "Hualpén").

      // Número de orden: si es numérico, lo convertimos a número de verdad.
      let sgVal = row[idxOrden];
      if (!isNaN(parseFloat(sgVal))) sgVal = parseFloat(sgVal);

      // --- LÓGICA DE BULTOS (usando Regex) ---
      // El número de bultos está escondido dentro de un texto, después de la palabra "Totales".
      // Usamos una "expresión regular" (Regex) para extraer ese número.
      let bultosVal = 0;
      if (idxBultos !== -1 && row[idxBultos]) {
          let txt = row[idxBultos].toString();
          let match = txt.match(/Totales\s*(\d+)/i); // Busca "Totales" seguido de un número.
          if (match && match[1]) bultosVal = parseInt(match[1], 10); // Guarda el número encontrado.
      }

      // --- LÓGICA "SIN ARMAR" ---
      // Determina si el pedido todavía NO se ha armado en bodega.
      //   - sinArmar = 1  -> el pedido NO está armado (pendiente).
      //   - sinArmar = 0  -> el pedido SÍ está armado (tiene fecha de picking).
      // La regla: si la columna "FECHA DE PICKING" tiene un valor, está armado.
      let sinArmarVal = 1; // Por defecto, asumimos que está sin armar.
      if (idxFechaPick !== -1) {
          let valF = row[idxFechaPick];
          if (valF && valF.toString().trim() !== "") {
              sinArmarVal = 0; // Tiene fecha de picking => está armado.
          }
      }

      // Devolvemos el pedido ya ordenado como un objeto con nombres claros.
      return {
        sg_original: sgVal,                                        // Número de orden original.
        cliente: row[idxCliente],                                  // Nombre del cliente.
        direccion: calleLimpia,                                    // Calle.
        comuna: comunaLimpia,                                      // Comuna.
        lat: parseFloat(row[idxLat].toString().replace(',','.')),  // Latitud (cambiamos coma por punto para que sea un número válido).
        lng: parseFloat(row[idxLon].toString().replace(',','.')),  // Longitud (igual que arriba).
        bultos: bultosVal,                                         // Cantidad de bultos.
        sinArmar: sinArmarVal,                                     // 1 = pendiente, 0 = armado.
        ruta: 1                                                    // Ruta inicial (luego se reasigna al optimizar).
      };
    })

    // --- PASO C: FILTRAR DE NUEVO ---
    // Descartamos pedidos sin coordenadas válidas (no se pueden ubicar en el mapa).
    // isNaN(o.lat) significa "la latitud no es un número"; con "!" nos quedamos con los que SÍ lo son.
    .filter(o => !isNaN(o.lat));

  return { orders };
}


/* =============================================================================
 *  GUARDAR Y EXPORTAR EL RESULTADO FINAL
 * ========================================================================== */

/**
 * saveFinalRouting(jsonData, local, horas)
 * ----------------------------------------
 * Genera el archivo Excel FINAL con todas las rutas ya asignadas, lo guarda
 * automáticamente en Drive y lo devuelve para que el usuario lo descargue.
 *
 * Parámetros:
 *   - jsonData: la lista de pedidos con su ruta ya asignada (viene del navegador).
 *   - local: la sucursal.
 *   - horas: el/los horario(s) procesados.
 *
 * Devuelve:
 *   - En caso de éxito: { status, fileName, base64 } (el archivo listo para descargar).
 *   - En caso de error: { status: "error", message }.
 */
function saveFinalRouting(jsonData, local, horas) {
  try {

    // -------------------------------------------------------------------------
    // 1. ORDENAR los pedidos.
    //    Primero por ruta (1, 2, 3...) y dejando los pedidos "LAT" siempre al final.
    // -------------------------------------------------------------------------
    jsonData.sort((a, b) => {
      if (a.ruta === "LAT" && b.ruta !== "LAT") return 1;   // "a" es LAT => va después.
      if (a.ruta !== "LAT" && b.ruta === "LAT") return -1;  // "b" es LAT => "a" va antes.
      if (a.ruta === "LAT" && b.ruta === "LAT") return 0;   // Ambos LAT => da igual el orden.
      return a.ruta - b.ruta;                               // Ambos normales => ordena por número de ruta.
    });

    // -------------------------------------------------------------------------
    // 2. CALCULAR TOTALES por ruta: cuántos bultos y cuántos pedidos sin armar.
    // -------------------------------------------------------------------------
    let sumasBultos = {};    // Total de bultos por ruta.
    let sumasSinArmar = {};  // Total de pedidos pendientes (sin armar) por ruta.

    jsonData.forEach(o => {
      // Si es la primera vez que vemos esta ruta, inicializamos sus contadores en 0.
      if (sumasBultos[o.ruta] === undefined) {
        sumasBultos[o.ruta] = 0;
        sumasSinArmar[o.ruta] = 0;
      }
      sumasBultos[o.ruta] += (o.bultos || 0);     // Acumula bultos.
      sumasSinArmar[o.ruta] += (o.sinArmar || 0); // Acumula pendientes.
    });

    // -------------------------------------------------------------------------
    // 3. CREAR la hoja de cálculo temporal donde armaremos el reporte.
    // -------------------------------------------------------------------------
    // Construye el nombre del archivo: si hay varios horarios usa "MULTI",
    // si es uno solo reemplaza los ":" por "h" (ej: "10:15" -> "10h15").
    let horaStr = Array.isArray(horas) ? "MULTI" : horas.replace(':','h');
    const fileName = `DESPACHO_${local}_${horaStr}`;

    const ss = SpreadsheetApp.create("TEMP_EXPORT"); // Crea una planilla nueva, temporal.
    const sheet = ss.getSheets()[0];                 // Toma su primera hoja.

    // Escribe la fila de títulos (encabezados) del reporte.
    sheet.appendRow(["Orden", "SG", "Nombre cliente", "Dirección", "Comuna", "Ruta", "Total Bultos Ruta"]);

    let rows = [];            // Aquí acumulamos todas las filas de datos antes de escribirlas.
    let routeCounters = {};   // Contador para numerar los pedidos dentro de cada ruta (1, 2, 3...).

    jsonData.forEach((o) => {
      // Numera correlativamente cada pedido DENTRO de su propia ruta.
      if (!routeCounters[o.ruta]) routeCounters[o.ruta] = 1;
      let correlativo = routeCounters[o.ruta];
      routeCounters[o.ruta]++;

      // Nombre legible de la ruta: "LAT" o "RUTA 1", "RUTA 2", etc.
      let nombreRuta = (o.ruta === "LAT") ? "LAT" : "RUTA " + o.ruta;

      // Recupera los totales calculados en el paso 2.
      let totalRuta = sumasBultos[o.ruta] || 0;
      let totalSinArmar = sumasSinArmar[o.ruta] || 0;

      // Arma el texto resumen de la celda de totales (ej: "12 bultos" + estado).
      let textoCelda = `${totalRuta} bultos`;
      if (totalSinArmar > 0) {
          textoCelda += `\n(${totalSinArmar} SG sin armar)`; // Hay pendientes: lo avisamos.
      } else {
          textoCelda += `\n(Todos OK)`;                       // Todo armado: mensaje positivo.
      }

      // Agrega la fila completa de este pedido a la lista.
      rows.push([correlativo, o.sg_original, o.cliente, o.direccion, o.comuna, nombreRuta, textoCelda]);
    });

    // Escribe TODAS las filas de datos en la hoja de una sola vez (más eficiente).
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
    // Fuerza la columna "SG" (la 2) a mostrarse como número entero sin decimales.
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat("0");

    // -------------------------------------------------------------------------
    // 4. FORMATEO Y FUSIÓN VISUAL.
    //    Recorre las filas y, por cada bloque de una misma ruta, le aplica un
    //    color de fondo y fusiona (combina) las celdas de "Ruta" y "Total Bultos".
    // -------------------------------------------------------------------------
    let startRow = 2; // Fila donde empieza el bloque de ruta que estamos formateando.
    for (let i = 0; i < rows.length; i++) {
      let currentR = rows[i][5];                                  // Nombre de ruta de la fila actual.
      let nextR = (i < rows.length - 1) ? rows[i+1][5] : null;    // Nombre de ruta de la fila siguiente.

      // Cuando la ruta cambia (o llegamos al final), cerramos y formateamos el bloque.
      if (currentR !== nextR) {
        let num = (i + 2) - startRow + 1;                         // Cantidad de filas que tiene este bloque.
        let fullR = sheet.getRange(startRow, 1, num, 7);          // El rango completo del bloque.

        // Elige el color de fondo del bloque.
        let color = "#ffffff";
        if (currentR === "LAT") {
          color = COLOR_LAT_EXCEL;                                // Gris para los LAT.
        } else {
          // Para rutas normales, toma un color de la paleta según el número de ruta.
          let rutaNum = parseInt(currentR.replace("RUTA ",""));
          color = COLORES_PASTELES[(rutaNum-1) % COLORES_PASTELES.length];
        }
        fullR.setBackground(color);

        // Si el bloque tiene más de una fila, fusiona las celdas de Ruta y Bultos
        // para que el dato aparezca una sola vez, centrado en todo el bloque.
        if (num > 1) {
            sheet.getRange(startRow, 6, num, 1).merge(); // Fusiona la columna "Ruta".
            sheet.getRange(startRow, 7, num, 1).merge(); // Fusiona la columna "Total Bultos".
        }

        // Centra el contenido de las columnas fusionadas y resalta el total de bultos.
        sheet.getRange(startRow, 6, num, 2).setVerticalAlignment("middle").setHorizontalAlignment("center");
        sheet.getRange(startRow, 7, num, 1).setFontWeight("bold").setFontSize(11).setWrap(true);

        startRow = i + 3; // El siguiente bloque empieza después de este.
      }
    }

    // Da formato a la fila de títulos: negrita, fondo oscuro y letra blanca.
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#333").setFontColor("white");

    // Define el ancho de cada columna para que el reporte se vea ordenado.
    sheet.setColumnWidth(1, 80);   // Orden
    sheet.setColumnWidth(2, 150);  // SG
    sheet.setColumnWidth(3, 180);  // Nombre cliente
    sheet.setColumnWidth(4, 250);  // Dirección
    sheet.setColumnWidth(5, 120);  // Comuna
    sheet.setColumnWidth(6, 70);   // Ruta
    sheet.setColumnWidth(7, 180);  // Total Bultos Ruta

    SpreadsheetApp.flush(); // Obliga a Google a aplicar TODOS los cambios pendientes ahora mismo.
    Utilities.sleep(1000);  // Espera 1 segundo para asegurar que todo quedó guardado antes de exportar.

    // -------------------------------------------------------------------------
    // 5. GENERAR el archivo .xlsx descargable y GUARDARLO en Drive.
    // -------------------------------------------------------------------------
    // Pide a Google que exporte la planilla en formato Excel (.xlsx) mediante una URL especial.
    const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx";
    // Hace la descarga, incluyendo el "token" de autorización para tener permiso de exportar.
    const res = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } });
    const blob = res.getBlob(); // El archivo Excel ya en formato binario.

    // -- INICIO AUTO-GUARDADO --
    // Guarda automáticamente una copia del Excel final en la carpeta de Drive,
    // con el nombre "DESPACHO_SUCURSAL_HORA.xlsx".
    try {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      folder.createFile(blob).setName(fileName + ".xlsx");
    } catch (saveError) {
      // Si falla el guardado en Drive, lo registramos pero NO detenemos el proceso:
      // el usuario igual podrá descargar su archivo.
      console.log("Error guardando en Drive: " + saveError);
    }
    // -- FIN AUTO-GUARDADO --

    // Manda la planilla temporal a la papelera (ya tenemos el archivo final).
    DriveApp.getFileById(ss.getId()).setTrashed(true);

    // Devuelve el resultado al navegador, con el archivo codificado en Base64 para descargarlo.
    return { status: "success", fileName: fileName + ".xlsx", base64: Utilities.base64Encode(blob.getBytes()) };

  } catch(e) {
    // Si algo falla en todo el proceso, devolvemos el error como texto.
    return { status: "error", message: e.toString() };
  }
}
