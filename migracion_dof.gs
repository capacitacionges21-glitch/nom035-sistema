function migrarRespuestasAlDOF() {

  var MAP = {
    '1':'1','2':'2','3':'3','4':'4','5':'5',
    '6':null,'7':null,'8':null,'9':null,'10':null,
    '11':'6','12':'7','13':'8',
    '14':'9','15':'10','16':'11','17':'12',
    '18':'13','19':'14','20':'15','21':'16',
    '22':'17','23':'18','24':'19','25':'20','26':'21','27':'22',
    '28':'23','29':'24','30':'25','31':'26','32':'27','33':'28',
    '34':'29','35':'30',
    '36':null,
    '37':'31','38':'32','39':'33','40':'34',
    '41':null,'42':null,
    '43':'35','44':'36','45':'37','46':'38','47':'39',
    '48':'40','49':'41','50':'42','51':'43',
    '52':null,
    '53':'44','54':'45','55':'46',
    '56':null,'57':null,
    '58':'47','59':'48','60':'49','61':'50','62':'51',
    '63':null,'64':null,
    '65':'52',
    '66':'53','67':'54','68':'55','69':'56',
    '70':null,'71':null,'72':null,'73':null,'74':null,'75':null,'76':null,
    '77':'57','78':'58','79':'59','80':'60',
    '81':'61','82':'62','83':'63','84':'64',
    '85':'65','86':'66','87':'67','88':'68','89':'69',
    '90':'70','91':'71','92':'72','93':'73','94':'74',
    '95':'75','96':'76','97':'77','98':'78','99':'79','100':'80',
    '101':'81','102':'82','103':'83','104':'84','105':'85',
    '106':'86','107':'87','108':'88','109':'89','110':'90',
    '111':'91','112':'92','113':'93','114':'94'
  };

  var SHEET_NAME = 'Respuestas';

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('Hoja "' + SHEET_NAME + '" no encontrada.'); return; }

  // Respaldo
  var ts     = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyyMMdd_HHmm');
  var backup = sheet.copyTo(ss);
  backup.setName(SHEET_NAME + '_backup_' + ts);
  Logger.log('Respaldo: ' + backup.getName());

  // Encabezados
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = {};
  for (var i = 0; i < headers.length; i++) { col[String(headers[i]).trim()] = i + 1; }

  var colResp  = col['respuestas'];
  var colTexto = col['respuestasTexto'];
  var colPts   = col['puntajeTotal'];
  var colRg    = col['riesgoGlobal'];

  if (!colResp) { SpreadsheetApp.getUi().alert('Columna "respuestas" no encontrada.'); return; }

  var lastRow = sheet.getLastRow();
  var migradas = 0, omitidas = 0, errores = 0;

  for (var row = 2; row <= lastRow; row++) {
    var rawResp = sheet.getRange(row, colResp).getValue();
    if (!rawResp) { omitidas++; continue; }

    var resp;
    try { resp = JSON.parse(rawResp); } catch(e) { errores++; continue; }

    // Detección: key '65' numérica 0-4 = sistema viejo; 'SI'/'NO' = ya migrado
    var v65 = resp['65'];
    if (v65 === undefined || v65 === 'SI' || v65 === 'NO') { omitidas++; continue; }
    var n65 = parseInt(v65, 10);
    if (isNaN(n65) || n65 < 0 || n65 > 4) { omitidas++; continue; }

    // Migrar respuestas
    var newResp = {};
    for (var oldKey in resp) {
      if (!resp.hasOwnProperty(oldKey)) continue;
      var newKey = MAP[oldKey];
      if (newKey !== undefined) {
        if (newKey !== null) newResp[newKey] = resp[oldKey];
      } else if (isNaN(parseInt(oldKey, 10))) {
        newResp[oldKey] = resp[oldKey]; // conservar claves no numéricas
      }
    }

    // Puntaje total
    var pts = 0;
    for (var k in newResp) {
      if (!newResp.hasOwnProperty(k)) continue;
      if (!isNaN(parseInt(k, 10))) {
        var n = parseInt(newResp[k], 10);
        if (!isNaN(n)) pts += n;
      }
    }

    // Riesgo global (umbrales oficiales NOM-035)
    var rg = pts >= 140 ? 'Muy alto' : pts >= 99 ? 'Alto' : pts >= 75 ? 'Medio' : pts >= 50 ? 'Bajo' : 'Nulo';

    // Guardar
    sheet.getRange(row, colResp).setValue(JSON.stringify(newResp));

    if (colTexto) {
      var rawTxt = sheet.getRange(row, colTexto).getValue();
      if (rawTxt) {
        try {
          var oldTxt = JSON.parse(rawTxt);
          var newTxt = {};
          for (var tk in oldTxt) {
            if (!oldTxt.hasOwnProperty(tk)) continue;
            var nk = MAP[tk];
            if (nk !== undefined) { if (nk !== null) newTxt[nk] = oldTxt[tk]; }
            else if (isNaN(parseInt(tk, 10))) newTxt[tk] = oldTxt[tk];
          }
          sheet.getRange(row, colTexto).setValue(JSON.stringify(newTxt));
        } catch(e) {}
      }
    }

    if (colPts) sheet.getRange(row, colPts).setValue(pts);
    if (colRg)  sheet.getRange(row, colRg).setValue(rg);

    migradas++;
    Logger.log('Fila ' + row + ' OK — puntaje: ' + pts + ' (' + rg + ')');
  }

  var msg = 'Migracion completada\n\n'
    + 'Filas migradas: ' + migradas + '\n'
    + 'Filas omitidas: ' + omitidas + '\n'
    + 'Errores: ' + errores + '\n\n'
    + 'Respaldo: ' + backup.getName();

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}
