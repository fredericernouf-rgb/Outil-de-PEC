/****************************************************
 * PURGE (Script 3) - VERSION =IMAGE() FORMULES
 *
 * ⚠️ DÉPENDANCES (projet unique = global partagé) :
 *   - CONFIG, logStructured_, sanitizeName_,
 *     getOrCreateClientsIndexSheet_, loadClientsIndex_,
 *     loadDeltaTodoSet_         → Script 1 (Rapport.gs)
 *   - getSheetByNameLoose_, toNumberRobust_, findFileByExactName_,
 *     buildPerimetreIndexFromSheet_,
 *     getNonDebuteeSiteNameSetFromPerimetre_ → Script 2 (Export.gs)
 *
 * REWORK :
 * ✅ Compteurs DETAIL : images lues/réécrites via getFormulas/setFormulas
 * ✅ Travaux DETAIL    : images lues/réécrites via getFormulas/setFormulas
 *    (compatible avec Script 2 qui écrit désormais en =IMAGE("url"))
 ****************************************************/

/****************************************************
 * FONCTION PRINCIPALE
 ****************************************************/
function purgeRapportClientNonDebute() {
  var startTime = Date.now();
  var executionId = Utilities.getUuid();

  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === 'purgeRapportClientNonDebute') {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }

  logStructured_('INFO', 'SCRIPT3', '🚀 Démarrage purge sites non débutés', { executionId: executionId });
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🚀 SCRIPT 3 - PURGE SITES NON DÉBUTÉS');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🆔 Execution ID : ' + executionId);
  Logger.log('');

  var ssBase = SpreadsheetApp.openById(CONFIG.spreadsheets.base);
  var shPEC = ssBase.getSheetByName(CONFIG.sheets.pec);
  var dataPEC = shPEC.getDataRange().getValues();

  var shSites = ssBase.getSheetByName(CONFIG.sheets.sites);
  var dataSites = shSites.getDataRange().getValues();

  var agences = CONFIG.agences;

  var indexCodesEntites = {};
  for (var k = 1; k < dataSites.length; k++) {
    indexCodesEntites[String(dataSites[k][2]).trim()] = dataSites[k][5];
  }

  var shIndex = getOrCreateClientsIndexSheet_(ssBase);
  var clientsIndexMap = loadClientsIndex_(shIndex);

  var clientsAvecNonDebute = {};
  for (var i = 1; i < dataPEC.length; i++) {
    var nomClient  = dataPEC[i][0];
    var codeClient = dataPEC[i][1];
    var etat       = String(dataPEC[i][4] || '').trim();

    if (!codeClient || !nomClient) continue;
    if (etat !== 'Prise en charge non débutée') continue;

    var key = String(codeClient).trim();
    if (!clientsAvecNonDebute[key]) {
      clientsAvecNonDebute[key] = {
        nomClient: String(nomClient).trim(),
        codeClient: key
      };
    }
  }

  logStructured_('INFO', 'SCRIPT3', '📋 Clients avec sites non débutés', {
    executionId: executionId,
    count: Object.keys(clientsAvecNonDebute).length
  });
  Logger.log('📋 Clients avec site(s) non débuté(s) : ' + Object.keys(clientsAvecNonDebute).length);

  var deltaSet = loadDeltaTodoSet_();
  if (deltaSet !== null) {
    logStructured_('INFO', 'SCRIPT3', '🎯 DELTA actif', {
      executionId: executionId,
      count: Object.keys(deltaSet).length
    });
    Logger.log('🎯 DELTA actif : ' + Object.keys(deltaSet).length + ' client(s)');
  } else {
    logStructured_('INFO', 'SCRIPT3', '📋 DELTA inactif', { executionId: executionId });
    Logger.log('📋 DELTA inactif : traitement complet');
  }
  Logger.log('');

  var clientsTraites = {};
  var purgeCount = 0;
  var errorCount = 0;
  var skippedCount = 0;

  for (var codeKeyPurge in clientsAvecNonDebute) {

    var infoClient = clientsAvecNonDebute[codeKeyPurge];

    if (deltaSet !== null && !deltaSet[codeKeyPurge]) {
      skippedCount++;
      continue;
    }

    var codeEntitePurge = indexCodesEntites[codeKeyPurge];
    if (!codeEntitePurge) {
      Logger.log('⚠️  WARN   | ' + codeKeyPurge + ' | codeEntite manquant');
      logStructured_('WARN', 'SCRIPT3', 'codeEntite manquant', {
        executionId: executionId,
        codeClient: codeKeyPurge
      });
      errorCount++;
      continue;
    }

    var agenceNomPurge = '';
    for (var clePurge in agences) {
      if (String(codeEntitePurge).indexOf(clePurge) !== -1) {
        agenceNomPurge = clePurge;
        break;
      }
    }
    if (!agenceNomPurge) {
      Logger.log('⚠️  WARN   | ' + codeKeyPurge + ' | Agence introuvable');
      logStructured_('WARN', 'SCRIPT3', 'Agence introuvable', {
        executionId: executionId,
        codeClient: codeKeyPurge
      });
      errorCount++;
      continue;
    }

    var idxKeyPurge = agenceNomPurge + '|' + codeKeyPurge;

    if (clientsTraites[idxKeyPurge]) continue;
    clientsTraites[idxKeyPurge] = true;

    var folderIdPurge = clientsIndexMap[idxKeyPurge] && clientsIndexMap[idxKeyPurge].folderId;
    if (!folderIdPurge) {
      Logger.log('⚠️  WARN   | ' + agenceNomPurge + ' | ' + codeKeyPurge + ' | ClientsIndex manquant');
      logStructured_('WARN', 'SCRIPT3', 'ClientsIndex manquant', {
        executionId: executionId,
        key: idxKeyPurge
      });
      errorCount++;
      continue;
    }

    var reportSSPurge;
    var reportFileIdPurge = clientsIndexMap[idxKeyPurge] && clientsIndexMap[idxKeyPurge].reportFileId;

    try {
      if (reportFileIdPurge) {
        reportSSPurge = SpreadsheetApp.openById(reportFileIdPurge);
      } else {
        var clientFolderPurge = DriveApp.getFolderById(folderIdPurge);
        var reportFilePurge = findFileByExactName_(clientFolderPurge, sanitizeName_(infoClient.nomClient));
        if (!reportFilePurge) {
          Logger.log('⚠️  WARN   | ' + agenceNomPurge + ' | ' + codeKeyPurge + ' | ' + infoClient.nomClient + ' | Fichier introuvable');
          logStructured_('WARN', 'SCRIPT3', 'Fichier introuvable', {
            executionId: executionId,
            nomClient: infoClient.nomClient
          });
          errorCount++;
          continue;
        }
        reportSSPurge = SpreadsheetApp.openById(reportFilePurge.getId());
      }
    } catch(e) {
      Logger.log('❌ ERREUR | ' + agenceNomPurge + ' | ' + codeKeyPurge + ' | ' + infoClient.nomClient + ' | ' + e.toString());
      logStructured_('ERROR', 'SCRIPT3', 'Erreur ouverture fichier', {
        executionId: executionId,
        key: idxKeyPurge,
        error: e.toString()
      });
      errorCount++;
      continue;
    }

    var nonDebuteeSetPurge = getNonDebuteeSiteNameSetFromPerimetre_(reportSSPurge);

    if (Object.keys(nonDebuteeSetPurge).length > 0) {
      try {
        purgeSitesNonDebutee_(reportSSPurge, nonDebuteeSetPurge);
        purgeCount++;

        Logger.log('🧹 PURGÉ  | ' + agenceNomPurge + ' | ' + codeKeyPurge + ' | ' + infoClient.nomClient + ' | ' + Object.keys(nonDebuteeSetPurge).length + ' site(s)');
        logStructured_('INFO', 'SCRIPT3', '🧹 Client purgé', {
          executionId: executionId,
          codeClient: codeKeyPurge,
          nomClient: infoClient.nomClient,
          agence: agenceNomPurge,
          nbSites: Object.keys(nonDebuteeSetPurge).length
        });
      } catch(e) {
        errorCount++;
        Logger.log('❌ ERREUR | ' + agenceNomPurge + ' | ' + codeKeyPurge + ' | ' + infoClient.nomClient + ' | ' + e.toString());
        logStructured_('ERROR', 'SCRIPT3', '❌ Erreur purge', {
          executionId: executionId,
          codeClient: codeKeyPurge,
          error: e.toString()
        });
      }
    } else {
      Logger.log('⏭️  SKIP   | ' + agenceNomPurge + ' | ' + codeKeyPurge + ' | ' + infoClient.nomClient + ' | Aucun site non débuté');
    }
  }

  var duration = ((Date.now() - startTime) / 1000).toFixed(1) + 's';

  Logger.log('');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('✅ SCRIPT 3 TERMINÉ');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('📊 Résumé :');
  Logger.log('   • Purgés  : ' + purgeCount);
  Logger.log('   • Ignorés : ' + skippedCount + ' (DELTA)');
  Logger.log('   • Erreurs : ' + errorCount);
  Logger.log('   • Durée   : ' + duration);
  Logger.log('═══════════════════════════════════════════════════');

  logStructured_('INFO', 'SCRIPT3', '✅ Terminé', {
    executionId: executionId,
    purged: purgeCount,
    skipped: skippedCount,
    errors: errorCount,
    duration: duration
  });
}

/****************************************************
 * FONCTION PRINCIPALE DE PURGE
 ****************************************************/
function purgeSitesNonDebutee_(reportSS, nonDebuteeSet) {
  if (!nonDebuteeSet || Object.keys(nonDebuteeSet).length === 0) return;

  var perimetreIdx = buildPerimetreIndexFromSheet_(reportSS);
  var codeByName = perimetreIdx.codeByName;

  function norm_(v) {
    return String(v || '')
      .replace(/[\u00A0\u202F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function isNonDebute_(nomSite) {
    var n = norm_(nomSite);
    var code = codeByName[n];
    if (!code) return false;
    return !!nonDebuteeSet[code];
  }

  function isRowEmpty_(row) {
    for (var k = 0; k < row.length; k++) {
      var v = row[k];
      if (v !== '' && v != null) return false;
    }
    return true;
  }

  function withRetry_(fn, label) {
    for (var a = 1; a <= 4; a++) {
      try { return fn(); }
      catch (e) {
        Logger.log('Sheets retry ' + a + '/4 ' + (label || '') + ' : ' + e);
        Utilities.sleep(400 * a);
        if (a === 4) throw e;
      }
    }
  }

  // ── 1) NE PAS TOUCHER "1.Périmètre de la prise en charge" ─

  // ── 2) Photothèque ────────────────────────────────────────
  // (inchangé : reconstruit déjà =IMAGE() depuis l'URL stockée en colonne K)
  (function purgePhototheque_() {
    var sh = getSheetByNameLoose_(reportSS, '2. Bibliothèque photographique');
    if (!sh) return;

    var startRow = 11;
    var last = sh.getLastRow();
    if (last < startRow) return;

    var n = last - startRow + 1;

    var data = withRetry_(function () {
      return sh.getRange(startRow, 2, n, 10).getValues();
    }, 'Photo getValues');

    var kept = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (isRowEmpty_(row)) continue;

      var nomSite = norm_(row[1]);
      if (nomSite && isNonDebute_(nomSite)) continue;

      row[1] = nomSite;
      kept.push(row);
    }

    kept.sort(function (a, b) {
      return norm_(a[1]).localeCompare(norm_(b[1]));
    });
    for (var r = 0; r < kept.length; r++) kept[r][0] = r + 1;

    withRetry_(function () {
      sh.getRange(startRow, 2, n, 10).clearContent();
    }, 'Photo clear');

    if (!kept.length) return;

    var colB  = kept.map(function(x) { return [x[0] || '']; });
    var colC  = kept.map(function(x) { return [x[1] || '']; });
    var colE  = kept.map(function(x) { return [x[3] || '']; });
    var colJ  = kept.map(function(x) { return [x[8] || '']; });
    var colK  = kept.map(function(x) { return [x[9] || '']; });

    var colGf = kept.map(function(x) {
      var url = String(x[9] || '').trim();
      if (!url) return [''];
      return ['=IMAGE("' + url + '")'];
    });

    withRetry_(function () { sh.getRange(startRow, 2,  kept.length, 1).setValues(colB);    }, 'Photo set B');
    withRetry_(function () { sh.getRange(startRow, 3,  kept.length, 1).setValues(colC);    }, 'Photo set C');
    withRetry_(function () { sh.getRange(startRow, 5,  kept.length, 1).setValues(colE);    }, 'Photo set E');
    withRetry_(function () { sh.getRange(startRow, 10, kept.length, 1).setValues(colJ);    }, 'Photo set J');
    withRetry_(function () {
      sh.getRange(startRow, 11, kept.length, 1)
        .setValues(colK)
        .setFontColor('#FFFFFF');
    }, 'Photo set K url');
    withRetry_(function () { sh.getRange(startRow, 7, kept.length, 1).setFormulas(colGf); }, 'Photo set G formulas');
  })();

  // ── 3) Compteurs DETAIL ───────────────────────────────────
  // ✅ REWORK : lit/réécrit les images via getFormulas/setFormulas (=IMAGE)
  (function purgeCompteursDetail_() {
    var sh = getSheetByNameLoose_(reportSS, '4. Inventaire des compteurs');
    if (!sh) {
      Logger.log('CompteursDetail: feuille introuvable, skip.');
      return;
    }

    var startRow = 17;
    var last = sh.getLastRow();
    if (last < startRow) {
      Logger.log('CompteursDetail: aucune donnée ligne 17+, skip.');
      return;
    }

    var n = last - startRow + 1;

    // Texte A..L + Formules images J..K
    var data, formsJK;
    try {
      data = withRetry_(function () {
        return sh.getRange(startRow, 1, n, 12).getValues();   // A..L
      }, 'CompteursDetail getValues');
      formsJK = withRetry_(function () {
        return sh.getRange(startRow, 10, n, 2).getFormulas();  // J..K
      }, 'CompteursDetail getFormulas');
    } catch(e) {
      Logger.log('CompteursDetail ERREUR lecture: ' + e);
      return;
    }

    // Filtrage parallèle (données + formules alignées)
    var keptData = [];
    var keptForm = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var nomSite = norm_(row[1]);  // col B
      if (nomSite && isNonDebute_(nomSite)) continue;
      keptData.push(row);
      keptForm.push(formsJK[i]);    // [formuleJ, formuleK]
    }

    // Renumérotation col A
    var counter = 1;
    for (var r = 0; r < keptData.length; r++) {
      if (keptData[r][0] !== '' && keptData[r][0] != null) {
        keptData[r][0] = counter++;
      }
    }

    try {
      withRetry_(function () {
        sh.getRange(startRow, 1, n, 12).clearContent();
      }, 'CompteursDetail clear');
    } catch(e) {
      Logger.log('CompteursDetail ERREUR clear: ' + e);
      return;
    }

    if (keptData.length === 0) return;

    try {
      // Texte A..I (un seul bloc)
      var valuesAI = keptData.map(function(row) {
        return [row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8]];
      });
      withRetry_(function() {
        sh.getRange(startRow, 1, keptData.length, 9).setValues(valuesAI);
      }, 'CompteursDetail set A..I');

      // Col L = état
      var valuesL = keptData.map(function(row) {
        return [row[11] !== undefined ? row[11] : ''];
      });
      withRetry_(function() {
        sh.getRange(startRow, 12, keptData.length, 1).setValues(valuesL);
      }, 'CompteursDetail set L');

      // ✅ Images J..K : réécriture batch des formules =IMAGE()
      withRetry_(function() {
        sh.getRange(startRow, 10, keptForm.length, 2).setFormulas(keptForm);
      }, 'CompteursDetail set img JK');

    } catch(e) {
      Logger.log('CompteursDetail ERREUR écriture: ' + e);
    }
  })();

  // ── 4) Travaux DETAIL ─────────────────────────────────────
  // ✅ REWORK : lit/réécrit les images via getFormulas/setFormulas (=IMAGE)
  (function purgeTravauxDetail_() {
    var sh = getSheetByNameLoose_(reportSS, "5. Proposition de travaux d'amélioration");
    if (!sh) return;

    var startRow = 12;
    var last = sh.getLastRow();
    if (last < startRow) return;

    var n = last - startRow + 1;

    // Texte B..I + Formules images G..H
    var data = withRetry_(function () {
      return sh.getRange(startRow, 2, n, 8).getValues();    // B..I
    }, 'Travaux getValues');
    var formsGH = withRetry_(function () {
      return sh.getRange(startRow, 7, n, 2).getFormulas();   // G..H
    }, 'Travaux getFormulas');

    // Bundle pour garder les formules alignées pendant le tri
    var bundle = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (isRowEmpty_(row)) continue;

      var nomSite = norm_(row[1]);  // col C (index 1 depuis B)
      if (nomSite && isNonDebute_(nomSite)) continue;

      row[1] = nomSite;
      bundle.push({ row: row, form: formsGH[i] });  // form = [G, H]
    }

    // Tri par nomSite puis numéro
    bundle.sort(function (a, b) {
      var c = norm_(a.row[1]).localeCompare(norm_(b.row[1]));
      if (c !== 0) return c;
      return toNumberRobust_(a.row[0]) - toNumberRobust_(b.row[0]);
    });
    for (var r = 0; r < bundle.length; r++) bundle[r].row[0] = r + 1;

    withRetry_(function () {
      sh.getRange(startRow, 2, n, 8).clearContent();
    }, 'Travaux clear');

    if (!bundle.length) return;

    // Texte B, C, D, E, F (index 0..4) et I (index 7)
    var colB = bundle.map(function(x) { return [x.row[0] || '']; });
    var colC = bundle.map(function(x) { return [x.row[1] || '']; });
    var colD = bundle.map(function(x) { return [x.row[2] || '']; });
    var colE = bundle.map(function(x) { return [x.row[3] || '']; });
    var colF = bundle.map(function(x) { return [x.row[4] || '']; });
    var colI = bundle.map(function(x) { return [x.row[7] || '']; });

    withRetry_(function(){ sh.getRange(startRow, 2, bundle.length, 1).setValues(colB); }, 'Travaux set B');
    withRetry_(function(){ sh.getRange(startRow, 3, bundle.length, 1).setValues(colC); }, 'Travaux set C');
    withRetry_(function(){ sh.getRange(startRow, 4, bundle.length, 1).setValues(colD); }, 'Travaux set D');
    withRetry_(function(){ sh.getRange(startRow, 5, bundle.length, 1).setValues(colE); }, 'Travaux set E');
    withRetry_(function(){ sh.getRange(startRow, 6, bundle.length, 1).setValues(colF); }, 'Travaux set F');
    withRetry_(function(){ sh.getRange(startRow, 9, bundle.length, 1).setValues(colI); }, 'Travaux set I');

    // ✅ Images G..H : réécriture batch des formules =IMAGE()
    var formGH = bundle.map(function(x) { return x.form; });  // [G, H]
    withRetry_(function(){
      sh.getRange(startRow, 7, bundle.length, 2).setFormulas(formGH);
    }, 'Travaux set img GH');
  })();

  // ── 5) Inventaire équipements par site (feuilles 3.x) ────
  // (inchangé : pour un site non débuté on efface tout, images comprises)
  (function purgeInvParSite_() {
    var prefix = '3. Inventaire des équipements par site ';
    var startRow = 19;
    var colStart = 1;
    var numCols = 14;

    function clearAllRowGroups_(sh, fromRow, toRow) {
      for (var row = fromRow; row <= toRow; row++) {
        for (var depth = 8; depth >= 1; depth--) {
          try {
            var grp = sh.getRowGroup(row, depth);
            if (grp) {
              grp.remove();
            }
          } catch(e) {}
        }
      }
    }

    function breakApartFromRow_(sh, fromRow, toRow) {
      var maxCols = sh.getMaxColumns();
      for (var r = fromRow; r <= toRow; r++) {
        try {
          sh.getRange(r, 1, 1, maxCols).breakApart();
        } catch(e) {}
      }
    }

    for (var idx = 1; ; idx++) {
      var sh = getSheetByNameLoose_(reportSS, prefix + idx);
      if (!sh) break;

      var codeSiteD7 = String(sh.getRange('D7').getValue() || '').trim();

      if (!codeSiteD7) {
        var nomSiteA4 = norm_((function(sheet) {
          return withRetry_(function() {
            return sheet.getRange('A4').getValue();
          }, 'InvParSite A4 ' + idx);
        })(sh));

        if (!nomSiteA4) continue;

        codeSiteD7 = codeByName[nomSiteA4] || '';

        if (!codeSiteD7) {
          Logger.log('purgeInvParSite: D7 vide et A4 introuvable dans index, feuille ' + idx + ' → skip');
          continue;
        }
        Logger.log('purgeInvParSite: D7 vide, fallback A4 → codeSite="' + codeSiteD7 + '"');
      }

      var estNonDebute = !!nonDebuteeSet[codeSiteD7];

      if (estNonDebute) {

        try {
          sh.getRange('A6').setValue('Prise en charge non débutée / en cours de reprise');
        } catch(e) {}

        var last = sh.getLastRow();
        if (last >= startRow) {
          var n = last - startRow + 1;

          (function(sheet, nRows, lastRow) {

            try {
              clearAllRowGroups_(sheet, startRow, lastRow);
            } catch(e) {
              Logger.log('WARN clearAllRowGroups_ : ' + e);
            }

            breakApartFromRow_(sheet, startRow, lastRow);

            try {
              withRetry_(function() {
                sheet.getRange(startRow, colStart, nRows, numCols).clearContent();
              }, 'InvParSite clear');
            } catch(e) {
              Logger.log('WARN clearContent InvParSite : ' + e);
            }

            try { sheet.getRange(startRow, colStart, nRows, numCols).setBackground('#FFFFFF'); } catch(e) {}
            try { sheet.getRange(startRow, colStart, nRows, numCols).setFontColor('#000000'); } catch(e) {}
            try { sheet.getRange(startRow, colStart, nRows, numCols).setFontWeight('normal'); } catch(e) {}
            try { sheet.getRange(startRow, colStart, nRows, numCols).setHorizontalAlignment('left'); } catch(e) {}
            try { sheet.setRowHeightsForced(startRow, nRows, 250); } catch(e) {}

            SpreadsheetApp.flush();

          })(sh, n, last);
        }

      } else {
        try {
          sh.getRange('A6').clearContent();
        } catch(e) {}
      }
    }
  })();
}
