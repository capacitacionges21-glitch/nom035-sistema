/**
 * Migración NOM-035 — Renumeración al cuestionario oficial DOF
 * ============================================================
 * Convierte las respuestas almacenadas en Google Sheets del sistema
 * original (preguntas 1-114) al cuestionario oficial Guía II DOF
 * (preguntas 1-94 en esquema de sistema).
 *
 * INSTRUCCIONES:
 *   1. Abre el editor de Apps Script del spreadsheet NOM-035.
 *   2. Pega este archivo completo como un nuevo script.
 *   3. Ejecuta la función  migrarRespuestasAlDOF()  UNA SOLA VEZ.
 *   4. Revisa el log y el mensaje de confirmación.
 *   5. Verifica la hoja de respaldo creada antes de eliminarla.
 *
 * SEGURIDAD:
 *   - Crea un respaldo automático con timestamp antes de modificar datos.
 *   - Detecta filas ya migradas (key '65' = 'SI'/'NO') y las omite.
 *   - Recalcula puntajeTotal y riesgoGlobal con los umbrales oficiales DOF.
 */

// ── Mapa de migración: viejo ID → nuevo ID  (null = eliminar) ──────────────
const MIGRATION_MAP = {
  '1':'1',  '2':'2',  '3':'3',  '4':'4',  '5':'5',
  '6':null, '7':null, '8':null, '9':null, '10':null,   // condiciones deficientes (no DOF)
  '11':'6', '12':'7', '13':'8',
  '14':'9', '15':'10','16':'11','17':'12',
  '18':'13','19':'14','20':'15','21':'16',
  '22':'17','23':'18','24':'19','25':'20','26':'21','27':'22',
  '28':'23','29':'24','30':'25','31':'26','32':'27','33':'28',
  '34':'29','35':'30',
  '36':null,                                            // inducción (no DOF)
  '37':'31','38':'32','39':'33','40':'34',
  '41':null,'42':null,                                  // capacitaciones (no DOF)
  '43':'35','44':'36','45':'37','46':'38','47':'39',
  '48':'40','49':'41','50':'42','51':'43',
  '52':null,                                            // jefe soluciona problemas (no DOF)
  '53':'44','54':'45','55':'46',
  '56':null,'57':null,                                  // compañeros teamwork/help (no DOF)
  '58':'47','59':'48','60':'49','61':'50','62':'51',
  '63':null,'64':null,                                  // nómina + reconocimiento extra (no DOF)
  '65':'52',                                            // movilidad → nuevo texto "pueden crecer"
  '66':'53','67':'54','68':'55','69':'56',
  '70':null,'71':null,'72':null,'73':null,
  '74':null,'75':null,'76':null,                        // pertenencia extras (no DOF)
  '77':'57','78':'58','79':'59','80':'60',
  '81':'61','82':'62','83':'63','84':'64',
  '85':'65','86':'66','87':'67','88':'68','89':'69',    // clientes: filtro→65, preguntas→66-69
  '90':'70','91':'71','92':'72','93':'73','94':'74',    // supervisores: filtro→70, preguntas→71-74
  '95':'75','96':'76','97':'77','98':'78','99':'79','100':'80', // trauma filtro →75-80
  '101':'81','102':'82','103':'83','104':'84','105':'85',
  '106':'86','107':'87','108':'88','109':'89','110':'90',
  '111':'91','112':'92','113':'93','114':'94'           // trauma eval →81-94
};

// ── IDs de la hoja y de la pestaña ─────────────────────────────────────────
const SS_ID     = '1AUtcH5u3LbgCCO7jzmcFgC5xch_R1rjq2__oe7ZFZLY';
const SHEET_RESP = 'Respuestas';

// ── Función principal ───────────────────────────────────────────────────────
function migrarRespuestasAlDOF() {
  const ss    = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(SHEET_RESP);

  if (!sheet) {
    Browser.msgBox('❌ Hoja "' + SHEET_RESP + '" no encontrada.');
    return;
  }

  // 1. Crear respaldo con timestamp
  const tz      = 'America/Mexico_City';
  const ts      = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmm');
  const backup  = sheet.copyTo(ss);
  backup.setName(SHEET_RESP + '_backup_' + ts);
  Logger.log('✅ Respaldo creado: ' + backup.getName());

  // 2. Leer encabezados (fila 1)
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = {};  // nombre → índice 1-based
  headers.forEach(function(h, i) { col[String(h).trim()] = i + 1; });

  const colResp  = col['respuestas'];
  const colTexto = col['respuestasTexto'];
  const colPts   = col['puntajeTotal'];
  const colRg    = col['riesgoGlobal'];

  if (!colResp) {
    Browser.msgBox('❌ Columna "respuestas" no encontrada.\nVerifica el encabezado en fila 1.');
    return;
  }

  // 3. Recorrer filas de datos
  const lastRow = sheet.getLastRow();
  var migradas = 0, omitidas = 0, errores = 0;

  for (var row = 2; row <= lastRow; row++) {
    var rawResp = sheet.getRange(row, colResp).getValue();
    if (!rawResp) { omitidas++; continue; }

    // Parsear JSON
    var resp;
    try {
      resp = (typeof rawResp === 'string') ? JSON.parse(rawResp) : rawResp;
    } catch (e) {
      Logger.log('Fila ' + row + ': JSON inválido — ' + e);
      errores++; continue;
    }

    // ── Detección de sistema ──
    // En sistema VIEJO: key '65' = pregunta "movilidad" con valor numérico 0-4
    // En sistema NUEVO: key '65' = filtroClientes con valor 'SI' o 'NO'
    var val65 = resp['65'];
    var num65 = parseInt(val65, 10);

    if (val65 === undefined) {
      // Sin key '65': respuesta incompleta o sistema desconocido → omitir
      Logger.log('Fila ' + row + ': sin key 65, omitida.');
      omitidas++; continue;
    }
    if (val65 === 'SI' || val65 === 'NO' || isNaN(num65) || num65 < 0 || num65 > 4) {
      // Sistema NUEVO o ya migrada → omitir
      omitidas++; continue;
    }

    // ── Migrar respuestas ──
    var newResp = {};
    for (var oldKey in resp) {
      if (!resp.hasOwnProperty(oldKey)) continue;
      var newKey = MIGRATION_MAP[oldKey];
      if (newKey !== undefined) {
        if (newKey !== null) newResp[newKey] = resp[oldKey];
        // null → eliminar, no copiar
      } else {
        // Clave no numérica (eventoFecha, comentarioFinal, etc.) → conservar
        if (isNaN(parseInt(oldKey, 10))) {
          newResp[oldKey] = resp[oldKey];
        }
        // Clave numérica desconocida → eliminar
      }
    }

    // ── Recalcular puntajeTotal (suma de valores numéricos) ──
    var newPts = 0;
    for (var k in newResp) {
      if (!newResp.hasOwnProperty(k)) continue;
      if (!isNaN(parseInt(k, 10))) {
        var n = parseInt(newResp[k], 10);
        if (!isNaN(n)) newPts += n;
      }
    }

    // ── Calcular riesgoGlobal con umbrales DOF oficiales ──
    var riesgo;
    if      (newPts >= 140) riesgo = 'Muy alto';
    else if (newPts >=  99) riesgo = 'Alto';
    else if (newPts >=  75) riesgo = 'Medio';
    else if (newPts >=  50) riesgo = 'Bajo';
    else                    riesgo = 'Nulo';

    // ── Escribir de regreso ──
    sheet.getRange(row, colResp).setValue(JSON.stringify(newResp));

    if (colTexto) {
      var rawTxt = sheet.getRange(row, colTexto).getValue();
      if (rawTxt) {
        try {
          var oldTxt = (typeof rawTxt === 'string') ? JSON.parse(rawTxt) : rawTxt;
          var newTxt = {};
          for (var tk in oldTxt) {
            if (!oldTxt.hasOwnProperty(tk)) continue;
            var nk = MIGRATION_MAP[tk];
            if (nk !== undefined) {
              if (nk !== null) newTxt[nk] = oldTxt[tk];
            } else if (isNaN(parseInt(tk, 10))) {
              newTxt[tk] = oldTxt[tk];
            }
          }
          sheet.getRange(row, colTexto).setValue(JSON.stringify(newTxt));
        } catch (e) { /* ignora errores en texto */ }
      }
    }

    if (colPts) sheet.getRange(row, colPts).setValue(newPts);
    if (colRg)  sheet.getRange(row, colRg).setValue(riesgo);

    migradas++;
    Logger.log('Fila ' + row + ' migrada — puntaje: ' + newPts + ' → ' + riesgo);
  }

  // 4. Resultado final
  var msg = '🏁 Migración completada\n\n'
    + '✅ Filas migradas:  ' + migradas + '\n'
    + '⏭️  Filas omitidas:  ' + omitidas + '\n'
    + '❌ Errores:          ' + errores + '\n\n'
    + '💾 Respaldo guardado en:\n' + backup.getName();

  Logger.log(msg);
  Browser.msgBox(msg);
}
