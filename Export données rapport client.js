function exportRapportClient_AllInOne() {
  var startTime = Date.now();

  // ✅ MODIF #6 : Sécurité anti-boucle infinie
  var propsSafe = PropertiesService.getScriptProperties();
  var ck0 = loadCheckpoint_();
  var lastRowSeen = parseInt(propsSafe.getProperty('exportRC_lastRowSeen') || '-1', 10);
  var stuckCount = parseInt(propsSafe.getProperty('exportRC_stuckCount') || '0', 10);

  if (ck0.lastRow === lastRowSeen) {
    stuckCount++;
    propsSafe.setProperty('exportRC_stuckCount', String(stuckCount));
    if (stuckCount >= 3) {
      var skipRow = ck0.lastRow + 1;
      logStructured_('ERROR', 'SCRIPT2', '🚨 Client bloquant détecté - saut forcé', {
        row: ck0.lastRow, nouveauRow: skipRow, tentatives: stuckCount
      });
      Logger.log('🚨 BLOCAGE row ' + ck0.lastRow + ' après ' + stuckCount + ' tentatives → saut forcé');
      saveCheckpoint_(skipRow, '', '');
      propsSafe.setProperty('exportRC_stuckCount', '0');
      propsSafe.setProperty('exportRC_lastRowSeen', String(skipRow));
    }
  } else {
    propsSafe.setProperty('exportRC_stuckCount', '0');
    propsSafe.setProperty('exportRC_lastRowSeen', String(ck0.lastRow));
  }

  function normKey_(v) {
    return String(v || '')
      .replace(/[\u00A0\u202F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === 'exportRapportClient_AllInOne') {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }

  logStructured_('INFO', 'SCRIPT2', '🚀 Démarrage export données', {});
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🚀 SCRIPT 2 - EXPORT DONNÉES (ALL IN ONE)');
  Logger.log('═══════════════════════════════════════════════════');

  var ssBase = SpreadsheetApp.openById(CONFIG.spreadsheets.base);
  var shPEC = ssBase.getSheetByName(CONFIG.sheets.pec);
  var dataPEC = shPEC.getDataRange().getValues();

  var todoSet = {};
  var deltaRaw = PropertiesService.getScriptProperties().getProperty('delta_todoClients');
  var todoKeys = deltaRaw ? JSON.parse(deltaRaw) : [];
  todoKeys.forEach(function(k){ todoSet[normKey_(k)] = true; });
  var isBootstrap = false;

  logStructured_('INFO', 'SCRIPT2', '🎯 DELTA chargé', { count: todoKeys.length });
  Logger.log('🎯 DELTA : ' + todoKeys.length + ' client(s) à traiter');
  Logger.log('');

  var shSites = ssBase.getSheetByName(CONFIG.sheets.sites);
  var dataSites = shSites.getDataRange().getValues();

  var agences = CONFIG.agences;

  var indexCodesEntites = {};
  for (var k = 1; k < dataSites.length; k++) {
    var keySite = normKey_(dataSites[k][2]);
    if (keySite) indexCodesEntites[keySite] = dataSites[k][5];
  }

  var shIndex = getOrCreateClientsIndexSheet_(ssBase);
  var clientsIndexMap = loadClientsIndex_(shIndex);

  var finalizedSitesByClient = buildFinalizedSitesByClient_(dataPEC);
  var clientsAvecNonDebute   = buildClientsAvecNonDebute_(dataPEC);

  var ssEquip = SpreadsheetApp.openById(CONFIG.spreadsheets.equipement);

  var shPhoto   = ssEquip.getSheetByName('Photothèque');
  var dataPhoto = shPhoto.getDataRange().getValues();
  var photosByClient = buildPhotosByClient_(dataPhoto);

  var shEquip   = ssEquip.getSheetByName('Equipement');
  var dataEquip = shEquip.getDataRange().getValues();
  var ecartInventaireByClientSite = buildEcartInventaireByClientSite_(dataEquip);
  var equipAgg                    = buildEquipAgg_(dataEquip);
  var equipRowsByClientSite       = buildEquipRowsByClientSite_(dataEquip);
  var compteurAgg                 = buildCompteurAgg_(dataEquip);
  var compteurRowsByClient        = buildCompteurRowsByClient_(dataEquip, shEquip);

  // ✅ MODIF #1 : Index équipements par code (lookup photo/nom/marque/modèle pour blocs CVC)
  var equipByCode = buildEquipByCode_(dataEquip);

  var shTravaux   = ssEquip.getSheetByName('Proposition_de_travaux');
  var dataTravaux = shTravaux.getDataRange().getValues();
  var travauxAgg          = buildTravauxAgg_(dataTravaux);
  var travauxRowsByClient = buildTravauxRowsByClient_(dataTravaux, shTravaux);

  var infoSiteByCodeSite = buildInfoSiteByCodeSite_(dataPEC);

  // ═══════════════════════════════════════════════════════════
  // ✅ MODIF A : Chargement CVC + index des locaux par site
  // ═══════════════════════════════════════════════════════════
  var CVC_SS_ID = '1hXdZY75SXTcGH6HPTLtMSio50bcDcFnPulQFYmgbYJg';
  var cvcLocauxBySite = {};
  try {
    var ssCVC   = SpreadsheetApp.openById(CVC_SS_ID);
    var shCVC   = ssCVC.getSheetByName('Ensemble technique CVC');
    var dataCVC = shCVC.getDataRange().getValues();
    cvcLocauxBySite = buildCvcLocauxBySite_(dataCVC);
  } catch(e) {
    logStructured_('ERROR', 'SCRIPT2', '❌ Erreur chargement CVC', { error: e.toString() });
    Logger.log('❌ Erreur chargement CVC: ' + e.toString());
  }
  // ═══════════════════════════════════════════════════════════

  var ck              = loadCheckpoint_();
  var lastIdx         = ck.lastRow;
  var resumeClientKey = ck.currentKey;
  var resumeStep      = ck.step;
  var props           = PropertiesService.getScriptProperties();

  logStructured_('INFO', 'SCRIPT2', '🔎 Checkpoint au démarrage', {
    lastRow: lastIdx, currentKey: resumeClientKey, step: resumeStep
  });

  // ✅ MODIF #1 : 4 min au lieu de 5 (marge pour scheduleResume_)
  var maxExecutionTime = 4 * 60 * 1000;

  var clientsTraites = {};
  var updated        = 0;
  var purgeOnlyCount = 0;
  var errorCount     = 0;

  // =========================================================
  // Purge des clients qui n'ont QUE des sites non débutés
  // =========================================================
  for (var codeKeyPurge in clientsAvecNonDebute) {
    if (finalizedSitesByClient[codeKeyPurge]) continue;
    if (!isBootstrap && !todoSet[codeKeyPurge]) continue;

    var infoClient      = clientsAvecNonDebute[codeKeyPurge];
    var codeEntitePurge = indexCodesEntites[codeKeyPurge];
    if (!codeEntitePurge) {
      logStructured_('WARN', 'SCRIPT2', '⏭️ Purge sautée: codeEntite manquant', { codeClient: codeKeyPurge });
      continue;
    }

    var agenceNomPurge = '';
    for (var clePurge in agences) {
      if (String(codeEntitePurge).indexOf(clePurge) !== -1) {
        agenceNomPurge = clePurge;
        break;
      }
    }
    if (!agenceNomPurge) continue;

    var idxKeyPurge   = agenceNomPurge + '|' + codeKeyPurge;
    var folderIdPurge = clientsIndexMap[idxKeyPurge] && clientsIndexMap[idxKeyPurge].folderId;
    if (!folderIdPurge) {
      logStructured_('WARN', 'SCRIPT2', 'ClientsIndex manquant (purge only)', { key: idxKeyPurge });
      continue;
    }

    var reportSSPurge;
    var reportFileIdPurge = clientsIndexMap[idxKeyPurge] && clientsIndexMap[idxKeyPurge].reportFileId;

    if (reportFileIdPurge) {
      try {
        reportSSPurge = SpreadsheetApp.openById(reportFileIdPurge);
      } catch(e) {
        logStructured_('ERROR', 'SCRIPT2', 'Impossible ouvrir fichier purge', { id: reportFileIdPurge, error: e.toString() });
        continue;
      }
    } else {
      try {
        var clientFolderPurge = DriveApp.getFolderById(folderIdPurge);
        var reportFilePurge   = findFileByExactName_(
          clientFolderPurge,
          sanitizeName_(infoClient.nomClient)
        );
        if (!reportFilePurge) {
          logStructured_('WARN', 'SCRIPT2', 'Fichier introuvable pour purge', { nomClient: infoClient.nomClient });
          continue;
        }
        reportSSPurge = SpreadsheetApp.openById(reportFilePurge.getId());
      } catch(e) {
        logStructured_('ERROR', 'SCRIPT2', 'Erreur Drive purge', { key: idxKeyPurge, error: e.toString() });
        continue;
      }
    }

    var nonDebuteeSetPurge = getNonDebuteeSiteNameSetFromPerimetre_(reportSSPurge);
    if (Object.keys(nonDebuteeSetPurge).length > 0) {
      Logger.log('🧹 PURGE ONLY | ' + agenceNomPurge + ' | ' + codeKeyPurge + ' | ' + infoClient.nomClient + ' | ' + Object.keys(nonDebuteeSetPurge).length + ' site(s) non débuté(s)');
      try {
        purgeSitesNonDebutee_(reportSSPurge, nonDebuteeSetPurge, [], []);
      } catch(e) {
        logStructured_('WARN', 'SCRIPT2', 'Erreur purgeSitesNonDebutee (purge only)', { codeClient: codeKeyPurge, error: e.toString() });
      }
      purgeOnlyCount++;
    }
  }

  // =========================================================
  // Boucle principale — clients finalisés
  // =========================================================
  for (var r = lastIdx; r < dataPEC.length; r++) {

    // ✅ MODIF #2 : Timeout périodique + scheduleResume_
    if ((r - lastIdx) % 4 === 0) {
      if (Date.now() - startTime > maxExecutionTime) {
        props.setProperty('exportRC_all_lastRow', String(r));
        scheduleResume_();
        logStructured_('WARN', 'SCRIPT2', '⏱️ TIMEOUT - reprise programmée', { row: r });
        Logger.log('⏱️ TIMEOUT: reprise sauvegardée à la ligne ' + r + ' (reprise dans 1 min)');
        return;
      }
    }

    var nomClient  = dataPEC[r][0];
    var codeClient = dataPEC[r][1];
    var etatSite   = dataPEC[r][4];

    if (!nomClient || !codeClient) continue;

    var codeKey = normKey_(codeClient);

    if (!isBootstrap && !todoSet[codeKey]) continue;
    if (etatSite !== 'Prise en charge finalisée') continue;

    var codeEntite = indexCodesEntites[codeKey];
    if (!codeEntite) {
      logStructured_('WARN', 'SCRIPT2', '⏭️ Sauté: codeEntite manquant', { codeClient: codeKey, nomClient: nomClient });
      continue;
    }

    var agenceNom = '';
    for (var cle in agences) {
      if (String(codeEntite).indexOf(cle) !== -1) {
        agenceNom = cle;
        break;
      }
    }
    if (!agenceNom) {
      logStructured_('WARN', 'SCRIPT2', '⏭️ Sauté: agence introuvable', { codeClient: codeKey, codeEntite: codeEntite });
      continue;
    }

    var clientUniqueKey = agenceNom + '|' + codeKey;
    if (clientsTraites[clientUniqueKey]) continue;
    clientsTraites[clientUniqueKey] = true;

    var finalized = finalizedSitesByClient[codeKey];
    if (!finalized || !finalized.list || finalized.list.length === 0) {
      logStructured_('WARN', 'SCRIPT2', '⏭️ Sauté: aucun site finalisé', { codeClient: codeKey });
      continue;
    }

    // ✅ MODIF #3 : Vérifier le timeout AVANT de démarrer un client (potentiellement long)
    if (Date.now() - startTime > maxExecutionTime) {
      props.setProperty('exportRC_all_lastRow', String(r));
      scheduleResume_();
      logStructured_('WARN', 'SCRIPT2', '⏱️ TIMEOUT avant client - reprise programmée', {
        row: r, codeClient: codeKey, nbSites: finalized.list.length
      });
      Logger.log('⏱️ TIMEOUT avant client ' + codeKey + ' (reprise dans 1 min)');
      return;
    }

    var idxKey = agenceNom + '|' + codeKey;
    if (resumeClientKey && idxKey !== resumeClientKey) {
      logStructured_('WARN', 'SCRIPT2', '⏭️ Sauté: bloqué par checkpoint', { codeClient: codeKey, resumeKey: resumeClientKey });
      continue;
    }

    var folderId = clientsIndexMap[idxKey] && clientsIndexMap[idxKey].folderId;
    if (!folderId) {
      logStructured_('WARN', 'SCRIPT2', '⏭️ Sauté: ClientsIndex manquant (folderId)', { key: idxKey });
      continue;
    }

    var reportSS;
    var reportFileId = clientsIndexMap[idxKey] && clientsIndexMap[idxKey].reportFileId;

    try {
      if (reportFileId) {
        reportSS = SpreadsheetApp.openById(reportFileId);
      } else {
        var clientFolder;
        try {
          clientFolder = DriveApp.getFolderById(folderId);
        } catch(e) {
          logStructured_('ERROR', 'SCRIPT2', 'FolderId invalide', { key: idxKey, folderId: folderId });
          continue;
        }
        var fileName   = sanitizeName_(nomClient);
        var reportFile = findFileByExactName_(clientFolder, fileName);
        if (!reportFile) {
          logStructured_('WARN', 'SCRIPT2', 'Rapport introuvable', { fileName: fileName, codeClient: codeKey });
          continue;
        }
        reportSS = SpreadsheetApp.openById(reportFile.getId());
      }

      var nonDebuteeSet = getNonDebuteeSiteNameSetFromPerimetre_(reportSS);

      saveCheckpoint_(r, idxKey, resumeStep || 'start');

      // ── Photothèque ──
      var shRCP = getSheetByNameLoose_(reportSS, '2. Bibliothèque photographique');
      if (shRCP) {
        var photosClient = photosByClient[codeKey] || [];
        writePhotoLibrary_(shRCP, photosClient);
      }
      saveCheckpoint_(r, idxKey, 'afterPhotos');
      resumeStep = 'afterPhotos';

      // ── Inventaire général ──
      var shRCI = getSheetByNameLoose_(reportSS, '3. Inventaire des équipements général');
      if (shRCI) {
        writeInventaireGeneral_(
          shRCI, codeKey, finalized.list, equipAgg,
          finalized.codeSiteSet, ecartInventaireByClientSite
        );
        var etats = countEquipEtatForClient_(dataEquip, codeKey);
        shRCI.getRange('F8:J8').setValues([[
          etats.bon, etats.moyen, etats.mauvais, etats.hsProv, etats.hsDef
        ]]);
        var pres = sumQteByPresenceForClient_(dataEquip, codeKey);
        shRCI.getRange('B9:E9').setValues([[
          pres.present, pres.aAjouter, pres.aSupprimer, pres.sansMention
        ]]);
      }
      saveCheckpoint_(r, idxKey, 'afterInvGeneral');
      resumeStep = 'afterInvGeneral';

      // ── Inventaire par site ──
      var eligibleSites = getEligibleSitesFromPerimetre_(reportSS);
      var invInfo       = ensureInventaireParSiteSheets_(reportSS, eligibleSites);
      writePerimetreLinks_(reportSS, invInfo.basePrefix);
      fillInventaireParSiteEquipements_(
        reportSS, codeKey, invInfo.orderedSitesForSheets,
        equipRowsByClientSite, nonDebuteeSet, infoSiteByCodeSite
      );
      saveCheckpoint_(r, idxKey, 'afterInvSite');
      resumeStep = 'afterInvSite';

      // ── Compteurs général ──
      var shCompteursGen = getSheetByNameLoose_(reportSS, '4. Inventaire des compteurs général');
      if (shCompteursGen) {
        writeCompteursGeneral_(
          shCompteursGen, codeKey, finalized.list,
          compteurAgg, finalized.codeSiteSet
        );
      }
      saveCheckpoint_(r, idxKey, 'afterCompteurs');
      resumeStep = 'afterCompteurs';

      // ── Travaux général ──
      var shTravGen = getSheetByNameLoose_(reportSS, "5. Proposition de travaux d'amélioration général");
      if (shTravGen) {
        writeTravauxGeneral_(
          shTravGen, codeKey, finalized.list,
          travauxAgg, finalized.codeSiteSet
        );
      }
      saveCheckpoint_(r, idxKey, 'afterTravaux');
      resumeStep = 'afterTravaux';

      // ═══════════════════════════════════════════════════════
      // ✅ MODIF B : Création feuilles CVC + écriture des 15 blocs
      // ═══════════════════════════════════════════════════════
      var orderedLocaux = buildOrderedCvcLocauxForClient_(cvcLocauxBySite, finalized);
      var cvcInfo = ensureEtatReglementaireSheets_(reportSS, orderedLocaux);

      // BLOC 1 : Fiche d'identification du site
      writeEtatReglementaireFicheIdentif_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
      // BLOC 2 : Conformité Sécurité, Santé au travail & Accès
      writeEtatReglementaireConformiteSecurite_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
      // BLOC 3 : Conformité Installation Électrique
      writeEtatReglementaireConformiteElectrique_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix, equipByCode);
      // BLOC 4 : Conformité Hydraulique & Traitement
      writeEtatReglementaireConformiteHydraulique_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix, equipByCode);
      // BLOC 5 : Conformité Évacuation & Rétention
      writeEtatReglementaireConformiteEvacuation_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
      // BLOC 6 : Conformité Calorifugeage & Isolation
      writeEtatReglementaireConformiteCalorifugeage_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
      // BLOC 7 : Conformité Fumisterie & Conduits
      writeEtatReglementaireConformiteFumisterie_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
      // BLOC 8 : Conformité Sécurité Incendie en Chaufferie
      writeEtatReglementaireConformiteIncendie_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
      // BLOC 9 : Conformité Installation Gaz
      writeEtatReglementaireConformiteGaz_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
      // BLOC 10 : Conformité Installation Fioul
      writeEtatReglementaireConformiteFioul_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
      // BLOC 11 : Conformité Froid & Climatisation
      writeEtatReglementaireConformiteFroid_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix, equipByCode);
      // BLOC 12 : Conformité ECS & Légionellose
      writeEtatReglementaireConformiteECS_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix, equipByCode);
      // BLOC 13 : Conformité Régulation & BACS
      writeEtatReglementaireConformiteRegulation_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
      // BLOC 14 : Conformité Ventilations Chaufferie
      writeEtatReglementaireConformiteVentilation_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
      // BLOC 15 : Conformité Chaudière Bois & Biomasse
      writeEtatReglementaireConformiteBiomasse_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);

      Logger.log('   📋 CVC: ' + orderedLocaux.length + ' feuille(s) + 15 blocs écrits');
      saveCheckpoint_(r, idxKey, 'afterCvc');
      resumeStep = 'afterCvc';
      // ═══════════════════════════════════════════════════════

      // ── Purge sites non débutés d'abord ──
      try {
        purgeSitesNonDebutee_(
          reportSS,
          nonDebuteeSet,
          compteurRowsByClient[codeKey] || [],
          travauxRowsByClient[codeKey]  || []
        );
      } catch(e) {
        logStructured_('WARN', 'SCRIPT2', 'Erreur purgeSitesNonDebutee', { codeClient: codeKey, error: e.toString() });
      }

      // ── Puis réécriture FINALE des images ──
      writeCompteursDetail_(reportSS, codeKey, compteurRowsByClient, finalized.codeSiteSet);
      writeTravauxDetail_(reportSS, codeKey, travauxRowsByClient, finalized.codeSiteSet);

      saveCheckpoint_(r, idxKey, 'done');
      resumeStep      = 'done';
      updated++;

      Logger.log('✅ TRAITÉ | ' + agenceNom + ' | ' + codeKey + ' | ' + nomClient + ' | ' + finalized.list.length + ' site(s)');
      logStructured_('INFO', 'SCRIPT2', '✅ Client traité', {
        codeClient: codeKey,
        nomClient: nomClient,
        agence: agenceNom,
        nbSites: finalized.list.length
      });

      resumeClientKey = '';
      resumeStep      = '';
      saveCheckpoint_(r + 1, '', '');

    } catch(e) {
      errorCount++;
      Logger.log('❌ ERREUR | ' + agenceNom + ' | ' + codeKey + ' | ' + nomClient + ' | ' + e.toString());
      logStructured_('ERROR', 'SCRIPT2', '❌ Erreur traitement client', {
        codeClient: codeKey,
        nomClient: nomClient,
        agence: agenceNom,
        error: e.toString(),
        stack: e.stack || ''
      });
    }
  }

  // ✅ MODIF #4 : Nettoyage COMPLET du checkpoint (currentKey + step + anti-boucle)
  clearCheckpoint_();
  props.deleteProperty('delta_todoClients');

  var duration = ((Date.now() - startTime) / 1000).toFixed(1) + 's';

  Logger.log('');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('✅ SCRIPT 2 TERMINÉ');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('📊 Résumé :');
  Logger.log('   • Traités     : ' + updated);
  Logger.log('   • Purge only  : ' + purgeOnlyCount);
  Logger.log('   • Erreurs     : ' + errorCount);
  Logger.log('   • Durée       : ' + duration);
  Logger.log('═══════════════════════════════════════════════════');

  logStructured_('INFO', 'SCRIPT2', '✅ Terminé', {
    updated: updated,
    purgeOnly: purgeOnlyCount,
    errors: errorCount,
    duration: duration
  });
}




/****************************************************
 * ═══════════════════════════════════════════════
 * EXPORT CLIENT SPÉCIFIQUE (TEST MANUEL)
 * ═══════════════════════════════════════════════
 ****************************************************/
function exportClientSpecifique() {
  var CLIENT_CIBLE = '1000164081';  // ✅ client de test CVC
  var startTime = Date.now();

  function normKey_(v) {
    return String(v || '')
      .replace(/[\u00A0\u202F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🧪 EXPORT CLIENT SPÉCIFIQUE (TEST) : ' + CLIENT_CIBLE);
  Logger.log('═══════════════════════════════════════════════════');

  var ssBase = SpreadsheetApp.openById(CONFIG.spreadsheets.base);
  var shPEC   = ssBase.getSheetByName(CONFIG.sheets.pec);
  var dataPEC = shPEC.getDataRange().getValues();

  var shSites   = ssBase.getSheetByName(CONFIG.sheets.sites);
  var dataSites = shSites.getDataRange().getValues();

  var agences = CONFIG.agences;

  var indexCodesEntites = {};
  for (var k = 1; k < dataSites.length; k++) {
    var keySite = normKey_(dataSites[k][2]);
    if (keySite) indexCodesEntites[keySite] = dataSites[k][5];
  }

  var shIndex         = getOrCreateClientsIndexSheet_(ssBase);
  var clientsIndexMap = loadClientsIndex_(shIndex);

  var finalizedSitesByClient = buildFinalizedSitesByClient_(dataPEC);
  var clientsAvecNonDebute   = buildClientsAvecNonDebute_(dataPEC);

  var ssEquip = SpreadsheetApp.openById(CONFIG.spreadsheets.equipement);

  var shPhoto    = ssEquip.getSheetByName('Photothèque');
  var dataPhoto  = shPhoto.getDataRange().getValues();
  var photosByClient = buildPhotosByClient_(dataPhoto);

  var shEquip   = ssEquip.getSheetByName('Equipement');
  var dataEquip = shEquip.getDataRange().getValues();
  var ecartInventaireByClientSite = buildEcartInventaireByClientSite_(dataEquip);
  var equipAgg                    = buildEquipAgg_(dataEquip);
  var equipRowsByClientSite       = buildEquipRowsByClientSite_(dataEquip);
  var compteurAgg                 = buildCompteurAgg_(dataEquip);
  var compteurRowsByClient        = buildCompteurRowsByClient_(dataEquip, shEquip);

  var shTravaux   = ssEquip.getSheetByName('Proposition_de_travaux');
  var dataTravaux = shTravaux.getDataRange().getValues();
  var travauxAgg          = buildTravauxAgg_(dataTravaux);
  var travauxRowsByClient = buildTravauxRowsByClient_(dataTravaux, shTravaux);

  var infoSiteByCodeSite = buildInfoSiteByCodeSite_(dataPEC);

  // ═══════════════════════════════════════════════════════════
  // ✅ MODIF A : Chargement CVC + index des locaux par site
  // ═══════════════════════════════════════════════════════════
  var CVC_SS_ID = '1hXdZY75SXTcGH6HPTLtMSio50bcDcFnPulQFYmgbYJg';
  var cvcLocauxBySite = {};
  try {
    var ssCVC   = SpreadsheetApp.openById(CVC_SS_ID);
    var shCVC   = ssCVC.getSheetByName('Ensemble technique CVC');
    var dataCVC = shCVC.getDataRange().getValues();
    cvcLocauxBySite = buildCvcLocauxBySite_(dataCVC);
  } catch(e) {
    Logger.log('❌ Erreur chargement CVC: ' + e.toString());
  }
  // ═══════════════════════════════════════════════════════════

  // ── Recherche du client cible ──
  var codeKey   = normKey_(CLIENT_CIBLE);
  var nomClient = '';

  for (var i = 1; i < dataPEC.length; i++) {
    if (normKey_(dataPEC[i][1]) === codeKey) {
      nomClient = String(dataPEC[i][0]).trim();
      break;
    }
  }

  if (!nomClient) {
    Logger.log('❌ Client introuvable dans PEC: ' + codeKey);
    return;
  }
  Logger.log('✅ Client trouvé dans PEC: ' + nomClient);

  // ── Recherche agence ──
  var codeEntite = '';
  for (var ks = 1; ks < dataSites.length; ks++) {
    if (normKey_(dataSites[ks][2]) === codeKey) {
      var ce = String(dataSites[ks][5] || '').trim();
      if (ce) {
        codeEntite = ce;
        break;
      }
    }
  }

  var agenceNom = '';
  if (codeEntite) {
    for (var cle in agences) {
      if (String(codeEntite).indexOf(cle) !== -1) {
        agenceNom = cle;
        break;
      }
    }
  }

  if (!agenceNom) {
    for (var idxSearch in clientsIndexMap) {
      if (idxSearch.indexOf('|' + codeKey) !== -1) {
        agenceNom = idxSearch.split('|')[0];
        break;
      }
    }
  }

  if (!agenceNom) {
    Logger.log('❌ Impossible de trouver agence pour: ' + codeKey);
    return;
  }
  Logger.log('✅ agenceNom: ' + agenceNom);

  // ── Vérification sites finalisés ──
  var finalized = finalizedSitesByClient[codeKey];
  if (!finalized || !finalized.list || finalized.list.length === 0) {
    Logger.log('❌ Aucun site finalisé pour: ' + codeKey);
    return;
  }
  Logger.log('✅ Sites finalisés: ' + finalized.list.length);

  // ── Ouverture du fichier rapport ──
  var idxKey   = agenceNom + '|' + codeKey;
  var folderId = clientsIndexMap[idxKey] && clientsIndexMap[idxKey].folderId;
  if (!folderId) {
    Logger.log('❌ ClientsIndex manquant pour: ' + idxKey);
    return;
  }

  var reportSS;
  var reportFileId = clientsIndexMap[idxKey] && clientsIndexMap[idxKey].reportFileId;

  if (reportFileId) {
    reportSS = SpreadsheetApp.openById(reportFileId);
  } else {
    Logger.log('❌ reportFileId manquant pour: ' + idxKey);
    return;
  }
  Logger.log('✅ Fichier rapport ouvert: ' + reportSS.getName());

  var nonDebuteeSet = getNonDebuteeSiteNameSetFromPerimetre_(reportSS);
  Logger.log('✅ Sites non débutés: ' + Object.keys(nonDebuteeSet).length);

  // ── Photothèque ──
  var shRCP = getSheetByNameLoose_(reportSS, '2. Bibliothèque photographique');
  if (shRCP) {
    var photosClient = photosByClient[codeKey] || [];
    Logger.log('✅ Photos: ' + photosClient.length);
    writePhotoLibrary_(shRCP, photosClient);
  }

  // ── Inventaire général ──
  var shRCI = getSheetByNameLoose_(reportSS, '3. Inventaire des équipements général');
  if (shRCI) {
    writeInventaireGeneral_(
      shRCI, codeKey, finalized.list, equipAgg,
      finalized.codeSiteSet, ecartInventaireByClientSite
    );
    var etats = countEquipEtatForClient_(dataEquip, codeKey);
    shRCI.getRange('F8:J8').setValues([[
      etats.bon, etats.moyen, etats.mauvais, etats.hsProv, etats.hsDef
    ]]);
    var pres = sumQteByPresenceForClient_(dataEquip, codeKey);
    shRCI.getRange('B9:E9').setValues([[
      pres.present, pres.aAjouter, pres.aSupprimer, pres.sansMention
    ]]);
  }

  // ── Inventaire par site ──
  var eligibleSites = getEligibleSitesFromPerimetre_(reportSS);
  var invInfo       = ensureInventaireParSiteSheets_(reportSS, eligibleSites);
  writePerimetreLinks_(reportSS, invInfo.basePrefix);
  fillInventaireParSiteEquipements_(
    reportSS, codeKey, invInfo.orderedSitesForSheets,
    equipRowsByClientSite, nonDebuteeSet, infoSiteByCodeSite
  );

  // ── Compteurs général ──
  var shCompteursGen = getSheetByNameLoose_(reportSS, '4. Inventaire des compteurs général');
  if (shCompteursGen) {
    writeCompteursGeneral_(
      shCompteursGen, codeKey, finalized.list,
      compteurAgg, finalized.codeSiteSet
    );
  }

  // ── Travaux général ──
  var shTravGen = getSheetByNameLoose_(reportSS, "5. Proposition de travaux d'amélioration général");
  if (shTravGen) {
    writeTravauxGeneral_(
      shTravGen, codeKey, finalized.list,
      travauxAgg, finalized.codeSiteSet
    );
  }

  // ═══════════════════════════════════════════════════════════
  // ✅ MODIF B : Feuilles CVC (Etat réglementaire) + BLOCS 1-2-3
  // ═══════════════════════════════════════════════════════════
  var orderedLocaux = buildOrderedCvcLocauxForClient_(cvcLocauxBySite, finalized);
  Logger.log('📋 CVC locaux à afficher: ' + orderedLocaux.length);
  for (var lc = 0; lc < orderedLocaux.length; lc++) {
    var L = orderedLocaux[lc];
    Logger.log('   Feuille ' + (lc+1) + ' | site=' + L.codeSite +
               ' | local=' + L.designation +
               ' | idCvc=' + L.idCvc + ' | audits=' + L.auditRows.length);
  }
  var cvcInfo = ensureEtatReglementaireSheets_(reportSS, orderedLocaux);

  // ✅ Index équipements par code (lookup photo/nom pour les blocs CVC)
  var equipByCode = buildEquipByCode_(dataEquip);

  // ✅ BLOC 1 : Fiche d'identification du site
  writeEtatReglementaireFicheIdentif_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);

  // ✅ BLOC 2 : Conformité Sécurité, Santé au travail & Accès
  writeEtatReglementaireConformiteSecurite_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);

  // ✅ BLOC 3 : Conformité Installation Électrique
  writeEtatReglementaireConformiteElectrique_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix, equipByCode);
  // ✅ BLOC 4 : Conformité Hydraulique & Traitement
  writeEtatReglementaireConformiteHydraulique_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix, equipByCode);
  // ✅ BLOC 5 : Conformité Évacuation & Rétention
  writeEtatReglementaireConformiteEvacuation_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
  // ✅ BLOC 6 : Conformité Calorifugeage & Isolation
  writeEtatReglementaireConformiteCalorifugeage_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
  // ✅ BLOC 7 : Conformité Fumisterie & Conduits
  writeEtatReglementaireConformiteFumisterie_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
  // ✅ BLOC 8 : Conformité Sécurité Incendie en Chaufferie
  writeEtatReglementaireConformiteIncendie_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
  // ✅ BLOC 9 : Conformité Installation Gaz
  writeEtatReglementaireConformiteGaz_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
  // ✅ BLOC 10 : Conformité Installation Fioul
  writeEtatReglementaireConformiteFioul_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
  // ✅ BLOC 11 : Conformité Froid & Climatisation
  writeEtatReglementaireConformiteFroid_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix, equipByCode);
  // ✅ BLOC 12 : Conformité ECS & Légionellose
  writeEtatReglementaireConformiteECS_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix, equipByCode);
  // ✅ BLOC 13 : Conformité Régulation & BACS
  writeEtatReglementaireConformiteRegulation_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
  // ✅ BLOC 14 : Conformité Ventilations Chaufferie
  writeEtatReglementaireConformiteVentilation_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);
  // ✅ BLOC 15 : Conformité Chaudière Bois & Biomasse
  writeEtatReglementaireConformiteBiomasse_(reportSS, cvcInfo.orderedLocauxForSheets, cvcInfo.basePrefix);

  Logger.log('✅ CVC: ' + orderedLocaux.length + ' feuille(s) + blocs 1-3 écrits');
  // ═══════════════════════════════════════════════════════════

  // ── Purge sites non débutés d'abord ──
  try {
    purgeSitesNonDebutee_(
      reportSS,
      nonDebuteeSet,
      compteurRowsByClient[codeKey] || [],
      travauxRowsByClient[codeKey]  || []
    );
  } catch(e) {
    Logger.log('⚠️ WARN purgeSitesNonDebutee_: ' + e);
  }

  // ── Puis réécriture FINALE des images ──
  writeCompteursDetail_(reportSS, codeKey, compteurRowsByClient, finalized.codeSiteSet);
  writeTravauxDetail_(reportSS, codeKey, travauxRowsByClient, finalized.codeSiteSet);

  var duration = ((Date.now() - startTime) / 1000).toFixed(1) + 's';
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('✅ Export terminé pour: ' + nomClient + ' (' + codeKey + ')');
  Logger.log('⏱️  Durée: ' + duration);
  Logger.log('═══════════════════════════════════════════════════');
}


/**
 * Construit l'index des LOCAUX CVC regroupés par site.
 * Un "local" = une ligne de création (I vide + FX = "Fiche d'identification du site")
 * Ses "audits" = les lignes dont I pointe vers l'ID CVC du local.
 *
 * @return {Object} { codeSite: [ {idCvc, codeClient, codeSite, nomSite, designation, ficheRow, auditRows} ] }
 */
function buildCvcLocauxBySite_(dataCVC) {
  var CVC_IDCVC      = 0;   // A
  var CVC_CODECLIENT = 2;   // C
  var CVC_CODESITE   = 3;   // D
  var CVC_NOMSITE    = 4;   // E
  var CVC_DESIGNATION= 6;   // G
  var CVC_REFINSTALL = 8;   // I
  var CVC_TYPEAUDIT  = 179; // FX

  var FICHE_IDENT = 'fiche d\'identification du site';

  function norm_(v) {
    return String(v || '')
      .replace(/[\u00A0\u202F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // ── ÉTAPE 1 : Identifier les LOCAUX (lignes de création) ──
  var locauxByIdCvc = {};  // idCvc → objet local

  for (var i = 1; i < dataCVC.length; i++) {
    var idCvc      = norm_(dataCVC[i][CVC_IDCVC]);
    var refInstall = norm_(dataCVC[i][CVC_REFINSTALL]);
    var typeAudit  = norm_(dataCVC[i][CVC_TYPEAUDIT]).toLowerCase();

    if (!idCvc) continue;

    // Création d'un local : I vide ET FX = "Fiche d'identification du site"
    if (refInstall === '' && typeAudit === FICHE_IDENT) {
      locauxByIdCvc[idCvc] = {
        idCvc:       idCvc,
        codeClient:  norm_(dataCVC[i][CVC_CODECLIENT]),
        codeSite:    norm_(dataCVC[i][CVC_CODESITE]),
        nomSite:     norm_(dataCVC[i][CVC_NOMSITE]),
        designation: norm_(dataCVC[i][CVC_DESIGNATION]),
        ficheRow:    dataCVC[i],   // la ligne complète de création
        auditRows:   []            // rempli en étape 2
      };
    }
  }

  // ── ÉTAPE 2 : Rattacher les AUDITS à leur local ──
  var orphelins = 0;
  for (var j = 1; j < dataCVC.length; j++) {
    var refInstallA = norm_(dataCVC[j][CVC_REFINSTALL]);
    if (refInstallA === '') continue;  // ce n'est pas un audit rattaché

    if (locauxByIdCvc[refInstallA]) {
      locauxByIdCvc[refInstallA].auditRows.push(dataCVC[j]);
    } else {
      orphelins++;  // audit pointant vers un local inexistant
    }
  }

  if (orphelins > 0) {
    Logger.log('⚠️ buildCvcLocauxBySite_: ' + orphelins + ' audit(s) orphelin(s) (refInstall sans local)');
  }

  // ── ÉTAPE 3 : Grouper par site ──
  var bySite = {};  // codeSite → [local, local, ...]
  for (var id in locauxByIdCvc) {
    var loc = locauxByIdCvc[id];
    var cs  = loc.codeSite;
    if (!cs) continue;
    if (!bySite[cs]) bySite[cs] = [];
    bySite[cs].push(loc);
  }

  // Tri des locaux dans chaque site (par désignation puis idCvc)
  for (var siteKey in bySite) {
    bySite[siteKey].sort(function(a, b) {
      var c = (a.designation || '').localeCompare(b.designation || '');
      if (c !== 0) return c;
      return (a.idCvc || '').localeCompare(b.idCvc || '');
    });
  }

  Logger.log('buildCvcLocauxBySite_: ' + Object.keys(locauxByIdCvc).length +
             ' local(aux) sur ' + Object.keys(bySite).length + ' site(s)');

  return bySite;
}
/**
 * Produit la liste ordonnée des locaux à afficher pour un client,
 * en ne gardant QUE les sites finalisés, dans l'ordre du périmètre.
 *
 * @param {Object} cvcLocauxBySite - résultat de buildCvcLocauxBySite_
 * @param {Object} finalized       - finalizedSitesByClient[codeKey] (list + codeSiteSet)
 * @return {Array} liste ordonnée de locaux
 */
function buildOrderedCvcLocauxForClient_(cvcLocauxBySite, finalized) {
  var out = [];
  if (!finalized || !finalized.list) return out;

  // Parcours des sites finalisés dans leur ordre (déjà trié par nomSite)
  for (var s = 0; s < finalized.list.length; s++) {
    var codeSite = String(finalized.list[s].codeSite || '').trim();
    var locaux   = cvcLocauxBySite[codeSite];
    if (!locaux || locaux.length === 0) continue;

    // Les locaux sont déjà triés dans buildCvcLocauxBySite_
    for (var l = 0; l < locaux.length; l++) {
      out.push(locaux[l]);
    }
  }

  return out;
}

/**
 * Crée/réutilise/supprime les feuilles "6. Etat réglementaire - Récapitulatif N"
 * pour chaque local des sites finalisés du client.
 *
 * @param {Spreadsheet} reportSS - le fichier rapport du client
 * @param {Array} orderedLocaux - liste ordonnée des locaux à afficher
 *                                 [ {idCvc, codeSite, nomSite, designation, ...} ]
 * @return {Object} { basePrefix, orderedLocauxForSheets }
 */
function ensureEtatReglementaireSheets_(reportSS, orderedLocaux) {
  var templateName = '6. Etat réglementaire - Récapitulatif';
  var basePrefix   = '6. Etat réglementaire - Récapitulatif ';  // avec espace final

  var templateSheet = getSheetByNameLoose_(reportSS, templateName);
  if (!templateSheet) {
    Logger.log('Template CVC introuvable: ' + templateName);
    return { basePrefix: basePrefix, orderedLocauxForSheets: [] };
  }

  var templatePos = templateSheet.getIndex();
  var N = orderedLocaux.length;

  // ── Crée/réutilise les feuilles pour i = 1..N ──
  for (var i = 0; i < N; i++) {
    var idx        = i + 1;
    var targetName = basePrefix + idx;
    var loc        = orderedLocaux[i];

    var sh = getSheetByNameLoose_(reportSS, targetName);
    if (!sh) {
      sh = templateSheet.copyTo(reportSS);
      sh.setName(targetName);
      reportSS.setActiveSheet(sh);
      reportSS.moveActiveSheet(templatePos + idx);
    }

    // ── Cellules de repère ──
    sh.getRange('A4').setValue(loc.nomSite || '');          // nom du site

    sh.getRange('A7')                                       // code site (purge)
      .setValue(loc.codeSite || '')
      .setFontColor('#FFFFFF');

    sh.getRange('A8')                                       // ID CVC local (identité)
      .setValue(loc.idCvc || '')
      .setFontColor('#FFFFFF');
  }

  // ── Supprime les feuilles en trop (j > N) ──
  var j = N + 1;
  while (true) {
    var extra = getSheetByNameLoose_(reportSS, basePrefix + j);
    if (!extra) break;
    Logger.log('Suppression feuille CVC en trop: ' + basePrefix + j);
    reportSS.deleteSheet(extra);
    j++;
  }

  // ── Masque le template ──
  try {
    templateSheet.hideSheet();
  } catch (e) {
    Logger.log('Impossible de masquer template CVC: ' + e);
  }

  return { basePrefix: basePrefix, orderedLocauxForSheets: orderedLocaux };
}
/**
 * BLOC 1 — Écrit la fiche d'identification du site dans chaque feuille CVC.
 * Source : loc.ficheRow (ligne où FX = "Fiche d'identification du site" + I vide).
 */
function writeEtatReglementaireFicheIdentif_(reportSS, orderedLocaux, basePrefix) {
  // ── Index colonnes CVC (0-based) ──
  var C_F  = 5;    // F  → D13
  var C_H  = 7;    // H  → D14
  var C_G  = 6;    // G  → D15 (nom local)
  var C_L  = 11;   // L  → D16
  var C_O  = 14;   // O  → D17
  var C_P  = 15;   // P  → D18 (TRUE/FALSE → Oui/Non)
  var C_N  = 13;   // N  → D19
  var C_M  = 12;   // M  → D20
  var C_AK = 36;   // AK → D21 (TRUE/FALSE → Oui/Non)
  var C_FZ = 181;  // FZ → G13 (URL photo)

  function val_(row, idx) {
    return String(row[idx] || '').trim();
  }

  // ✅ Conversion TRUE/FALSE → Oui/Non
  function ouiNon_(v) {
    if (v === true || String(v).toUpperCase() === 'TRUE')  return 'Oui';
    if (v === false || String(v).toUpperCase() === 'FALSE') return 'Non';
    return String(v || '').trim();  // sinon valeur brute (sécurité)
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var fiche = orderedLocaux[i].ficheRow;
    if (!fiche) continue;

    // ── D13:D21 en UN SEUL appel (9 cellules contiguës) ──
    sh.getRange('D13:D21').setValues([
      [val_(fiche, C_F)],          // D13
      [val_(fiche, C_H)],          // D14
      [val_(fiche, C_G)],          // D15 (nom local)
      [val_(fiche, C_L)],          // D16
      [val_(fiche, C_O)],          // D17
      [ouiNon_(fiche[C_P])],       // D18 ✅ Oui/Non
      [val_(fiche, C_N)],          // D19
      [val_(fiche, C_M)],          // D20
      [ouiNon_(fiche[C_AK])]       // D21 ✅ Oui/Non
    ]);

    // ── Photo en G13 via =IMAGE("url") ──
    var photoUrl = val_(fiche, C_FZ);
    if (photoUrl) {
      sh.getRange('G13').setFormula('=IMAGE("' + photoUrl + '")');
    } else {
      sh.getRange('G13').clearContent();
    }
  }
}
/**
 * BLOC 2 — Conformité Sécurité, Santé au travail & Accès.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Sécurité, Santé au travail & Accès".
 * Préserve les titres en D32, D38, D44 (écriture en 4 blocs séparés).
 */
function writeEtatReglementaireConformiteSecurite_(reportSS, orderedLocaux, basePrefix) {
  var C_FX = 179;  // FX = type d'audit
  var TYPE_CIBLE = 'conformité sécurité, santé au travail & accès';

  // Index colonnes CVC (0-based)
  var C_AI = 34, C_AF = 31, C_CS = 96, C_AG = 32, C_AH = 33, C_AJ = 35;
  var C_Q = 16, C_R = 17, C_S = 18, C_T = 19, C_U = 20, C_V = 21;
  var C_W = 22, C_X = 23, C_Y = 24, C_Z = 25, C_AA = 26, C_AB = 27;
  var C_AC = 28, C_AD = 29, C_AE = 30;

  function val_(row, idx) {
    return String(row[idx] || '').trim();
  }

  function normFx_(v) {
    return String(v || '')
      .replace(/[\u00A0\u202F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];

    // ── Chercher l'audit "Conformité Sécurité..." ──
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) {
        audit = loc.auditRows[a];
        break;
      }
    }

    if (!audit) {
      Logger.log('   ⏭️ Bloc2: pas d\'audit "Conformité Sécurité" pour local ' + loc.idCvc);
      continue;
    }

    // ── Bloc 1 : D27:D31 (AI, AF, CS, AG, AH) ──
    sh.getRange('D27:D31').setValues([
      [val_(audit, C_AI)],   // D27
      [val_(audit, C_AF)],   // D28
      [val_(audit, C_CS)],   // D29
      [val_(audit, C_AG)],   // D30
      [val_(audit, C_AH)]    // D31
    ]);
    // D32 = titre → NON touchée

    // ── Bloc 2 : D33:D37 (AJ, Q, R, S, T) ──
    sh.getRange('D33:D37').setValues([
      [val_(audit, C_AJ)],   // D33
      [val_(audit, C_Q)],    // D34
      [val_(audit, C_R)],    // D35
      [val_(audit, C_S)],    // D36
      [val_(audit, C_T)]     // D37
    ]);
    // D38 = titre → NON touchée

    // ── Bloc 3 : D39:D43 (U, V, W, X, Y) ──
    sh.getRange('D39:D43').setValues([
      [val_(audit, C_U)],    // D39
      [val_(audit, C_V)],    // D40
      [val_(audit, C_W)],    // D41
      [val_(audit, C_X)],    // D42
      [val_(audit, C_Y)]     // D43
    ]);
    // D44 = titre → NON touchée

    // ── Bloc 4 : D45:D50 (Z, AA, AB, AC, AD, AE) ──
    sh.getRange('D45:D50').setValues([
      [val_(audit, C_Z)],    // D45
      [val_(audit, C_AA)],   // D46
      [val_(audit, C_AB)],   // D47
      [val_(audit, C_AC)],   // D48
      [val_(audit, C_AD)],   // D49
      [val_(audit, C_AE)]    // D50
    ]);
  }
}
/**
 * Index des équipements par CODE équipement (colonne A de la table Equipement).
 * Retourne { codeEquip: { nom (col H), photo (col AA) } }
 */
function buildEquipByCode_(dataEquip) {
  var EQ_CODE     = 0;   // A  = code équipement
  var EQ_NOM      = 7;   // H  = nom
  var EQ_AG       = 32;  // AG = marque
  var EQ_AH       = 33;  // AH = modèle
  var EQ_AK       = 36;  // AK = numéro de série
  var EQ_PHOTO_AA = 26;  // AA = url photo

  var map = {};
  for (var i = 1; i < dataEquip.length; i++) {
    var code = String(dataEquip[i][EQ_CODE] || '').trim();
    if (!code) continue;
    if (!map[code]) {
      map[code] = {
        nom:    String(dataEquip[i][EQ_NOM]      || '').trim(),
        ag:     String(dataEquip[i][EQ_AG]       || '').trim(),
        ah:     String(dataEquip[i][EQ_AH]       || '').trim(),
        serie:  String(dataEquip[i][EQ_AK]       || '').trim(),
        photo:  String(dataEquip[i][EQ_PHOTO_AA] || '').trim()
      };
    }
  }
  return map;
}



/**
 * BLOC 3 — Conformité Installation Électrique.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Installation Électrique".
 * Préserve les titres en D61, D66, D69.
 * Partie spéciale : colonne AN (code équip) → lookup table Equipement → photo G56 + nom H67.
 */
function writeEtatReglementaireConformiteElectrique_(reportSS, orderedLocaux, basePrefix, equipByCode) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité installation électrique';
  var NOM_EQUIP_CELL = 'H67';

  var C_AL = 37, C_AS = 44, C_AO = 40, C_AR = 43, C_AZ = 51;
  var C_AM = 38, C_AT = 45, C_AU = 46, C_BB = 53;
  var C_AV = 47, C_AW = 48;
  var C_CU = 98, C_CT = 97;
  var C_AN = 39;
  var C_DX = 127;
  var C_DY = 128;
  var C_GA = 182;   // ⚠️ AJOUT : nouvelle ligne D68

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function ouiNon_(v) {
    if (v === true || String(v).toUpperCase() === 'TRUE')  return 'Oui';
    if (v === false || String(v).toUpperCase() === 'FALSE') return 'Non';
    return String(v || '').trim();
  }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc3: pas d\'audit "Conformité Installation Électrique" pour local ' + loc.idCvc);
      continue;
    }

    // ── Bloc 1 : D56:D60 (AL, AS, AO, AR, AZ) ──
    sh.getRange('D56:D60').setValues([
      [val_(audit, C_AL)], [val_(audit, C_AS)], [val_(audit, C_AO)],
      [val_(audit, C_AR)], [val_(audit, C_AZ)]
    ]);
    // D61 = titre → NON touchée

    // ── Bloc 2 : D62:D65 (AM, AT, AU→Oui/Non, BB→Oui/Non) ──
    sh.getRange('D62:D65').setValues([
      [val_(audit, C_AM)], [val_(audit, C_AT)],
      [ouiNon_(audit[C_AU])], [ouiNon_(audit[C_BB])]
    ]);
    // D66 = titre → NON touchée

    // ── Bloc 3 : ⚠️ MODIF ajout ligne D68 ──
    // Avant : D67:D68 = [AV, AW]
    // Après : D67 = AV | D68 = GA (nouveau) | D69 = AW
    sh.getRange('D67').setValue(ouiNon_(audit[C_AV]));   // D67 (inchangé)
    sh.getRange('D68').setValue(val_(audit, C_GA));      // D68 ⚠️ NOUVEAU (GA)
    sh.getRange('D69').setValue(val_(audit, C_AW));      // D69 (ex-D68)
    // D70 = titre (ex-D69) → NON touchée

    // ── Bloc 4 : D71:D74 (CU, CT, DX, DY) — ex-D70:D73 ──
    sh.getRange('D71:D74').setValues([
      [val_(audit, C_CU)],   // D71
      [val_(audit, C_CT)],   // D72
      [val_(audit, C_DX)],   // D73
      [val_(audit, C_DY)]    // D74
    ]);

    // ── Partie spéciale : photo (G56) + nom (H67) ──
    var codeEquip = val_(audit, C_AN);
    if (codeEquip && equipByCode && equipByCode[codeEquip]) {
      var eq = equipByCode[codeEquip];
      if (eq.photo) sh.getRange('G56').setFormula('=IMAGE("' + eq.photo + '")');
      else sh.getRange('G56').clearContent();
      sh.getRange(NOM_EQUIP_CELL).setValue(eq.nom || '');
    } else {
      sh.getRange('G56').clearContent();
      sh.getRange(NOM_EQUIP_CELL).clearContent();
      if (codeEquip) Logger.log('   ⚠️ Bloc3: codeEquip "' + codeEquip + '" introuvable dans table Equipement');
    }
  }
}


/**
 * BLOC 4 — Conformité Hydraulique & Traitement.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Hydraulique & Traitement".
 * Préserve les titres (sauts) en D82, D85, D89, D92, D96, D104, D108.
 * 4 lookups équipement (clés DB→G, BH→F, BL→E, BJ→D) : image109 + nom120 + AG121 + AH122.
 */
function writeEtatReglementaireConformiteHydraulique_(reportSS, orderedLocaux, basePrefix, equipByCode) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité hydraulique & traitement';

  var C_J = 9, C_K = 10;
  var C_BG = 58, C_BH = 59, C_BJ = 61, C_BK = 62, C_BL = 63;
  var C_BM = 64, C_BN = 65, C_BO = 66, C_BP = 67, C_BQ = 68;
  var C_BR = 69, C_BS = 70, C_BT = 71, C_BU = 72;
  var C_CY = 102, C_CZ = 103;
  var C_DA = 104, C_DB = 105, C_DC = 106, C_DD = 107, C_DE = 108, C_DF = 109;
  var C_GB = 183;   // ⚠️ AJOUT : nouvelle ligne D87

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function ouiNon_(v) {
    if (v === true || String(v).toUpperCase() === 'TRUE')  return 'Oui';
    if (v === false || String(v).toUpperCase() === 'FALSE') return 'Non';
    return String(v || '').trim();
  }
  function presence_(v) { return String(v || '').trim() !== '' ? 'Oui' : 'Non'; }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }
  function firstKey_(v) {
    var s = String(v || '').trim();
    if (!s) return '';
    return s.split(/[,;\n]/)[0].trim();
  }

  // ⚠️ MODIF +2 : lookups 109→111, 120→122, 121→123, 122→124
  function fillEquipLookup_(sh, audit, keyColIdx, destCol) {
    var key = firstKey_(audit[keyColIdx]);
    if (key && equipByCode && equipByCode[key]) {
      var eq = equipByCode[key];
      if (eq.photo) sh.getRange(destCol + '111').setFormula('=IMAGE("' + eq.photo + '")');
      else sh.getRange(destCol + '111').clearContent();
      sh.getRange(destCol + '122').setValue(eq.nom || '');
      sh.getRange(destCol + '123').setValue(eq.ag  || '');
      sh.getRange(destCol + '124').setValue(eq.ah  || '');
    } else {
      sh.getRange(destCol + '111').clearContent();
      sh.getRange(destCol + '122').clearContent();
      sh.getRange(destCol + '123').clearContent();
      sh.getRange(destCol + '124').clearContent();
      if (key) Logger.log('   ⚠️ Bloc4: clé "' + key + '" (col ' + destCol + ') introuvable dans Equipement');
    }
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc4: pas d\'audit "Conformité Hydraulique & Traitement" pour local ' + loc.idCvc);
      continue;
    }

    // ── D80:D81 (DE→Oui/Non, DF) — saut D82 ── (inchangé)
    sh.getRange('D80:D81').setValues([
      [ouiNon_(audit[C_DE])], [val_(audit, C_DF)]
    ]);
    // ── D83:D84 (K, J) — saut D85 ── (inchangé)
    sh.getRange('D83:D84').setValues([
      [val_(audit, C_K)], [val_(audit, C_J)]
    ]);
    // ── D86:D87 ⚠️ NOUVELLES LIGNES (J, GB) ──
    sh.getRange('D86:D87').setValues([
      [val_(audit, C_J)],    // D86 (J)
      [val_(audit, C_GB)]    // D87 (GB)
    ]);
    // ── D88:D90 (BK, CY, CZ) — saut D91 ──
    sh.getRange('D88:D90').setValues([
      [val_(audit, C_BK)], [val_(audit, C_CY)], [val_(audit, C_CZ)]
    ]);
    // ── D92:D93 (DA, BO) — saut D94 ──
    sh.getRange('D92:D93').setValues([
      [val_(audit, C_DA)], [val_(audit, C_BO)]
    ]);
    // ── D95:D97 (BM, BN, BG) — saut D98 ──
    sh.getRange('D95:D97').setValues([
      [val_(audit, C_BM)], [val_(audit, C_BN)], [val_(audit, C_BG)]
    ]);
    // ── D99:D105 (BH→présence, BP, BQ, BR, BS, BT, BU) — saut D106 ──
    sh.getRange('D99:D105').setValues([
      [presence_(audit[C_BH])],
      [val_(audit, C_BP)], [val_(audit, C_BQ)], [val_(audit, C_BR)],
      [val_(audit, C_BS)], [val_(audit, C_BT)], [val_(audit, C_BU)]
    ]);
    // ── D107:D109 (DB→présence, DC, DD) ──
    sh.getRange('D107:D109').setValues([
      [presence_(audit[C_DB])], [val_(audit, C_DC)], [val_(audit, C_DD)]
    ]);

    // ── Lookups équipement (111/122/123/124) ──
    fillEquipLookup_(sh, audit, C_DB, 'G');
    fillEquipLookup_(sh, audit, C_BH, 'F');
    fillEquipLookup_(sh, audit, C_BL, 'E');
    fillEquipLookup_(sh, audit, C_BJ, 'D');
  }
}



/**
 * BLOC 5 — Conformité Évacuation & Rétention.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Évacuation & Rétention".
 * Préserve les titres (sauts) en D131, D134.
 */
function writeEtatReglementaireConformiteEvacuation_(reportSS, orderedLocaux, basePrefix) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité évacuation & rétention';

  var C_BV = 73, C_BW = 74, C_BX = 75;
  var C_BY = 76, C_BZ = 77;
  var C_CA = 78, C_CB = 79;

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc5: pas d\'audit "Conformité Évacuation & Rétention" pour local ' + loc.idCvc);
      continue;
    }

    // ⚠️ MODIF +2
    // ── D130:D132 (BV, BW, BX) — saut D133 ──
    sh.getRange('D130:D132').setValues([
      [val_(audit, C_BV)], [val_(audit, C_BW)], [val_(audit, C_BX)]
    ]);
    // ── D134:D135 (BY, BZ) — saut D136 ──
    sh.getRange('D134:D135').setValues([
      [val_(audit, C_BY)], [val_(audit, C_BZ)]
    ]);
    // ── D137:D138 (CA, CB) ──
    sh.getRange('D137:D138').setValues([
      [val_(audit, C_CA)], [val_(audit, C_CB)]
    ]);
  }
}


/**
 * BLOC 6 — Conformité Calorifugeage & Isolation.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Calorifugeage & Isolation".
 */
function writeEtatReglementaireConformiteCalorifugeage_(reportSS, orderedLocaux, basePrefix) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité calorifugeage & isolation';

  var C_FK = 166, C_CD = 81, C_CE = 82, C_CF = 83;

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc6: pas d\'audit "Conformité Calorifugeage & Isolation" pour local ' + loc.idCvc);
      continue;
    }

    // ⚠️ MODIF +2 : D144:D147 (FK, CD, CE, CF)
    sh.getRange('D144:D147').setValues([
      [val_(audit, C_FK)], [val_(audit, C_CD)],
      [val_(audit, C_CE)], [val_(audit, C_CF)]
    ]);
  }
}




/**
 * BLOC 7 — Conformité Fumisterie & Conduits.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Fumisterie & Conduits".
 * Préserve le titre (saut) en D154.
 */
function writeEtatReglementaireConformiteFumisterie_(reportSS, orderedLocaux, basePrefix) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité fumisterie & conduits';

  var C_CG = 84, C_CH = 85, C_CI = 86;
  var C_CJ = 87, C_CK = 88, C_CL = 89;

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc7: pas d\'audit "Conformité Fumisterie & Conduits" pour local ' + loc.idCvc);
      continue;
    }

    // ⚠️ MODIF +2
    // ── D153:D155 (CG, CH, CI) — saut D156 ──
    sh.getRange('D153:D155').setValues([
      [val_(audit, C_CG)], [val_(audit, C_CH)], [val_(audit, C_CI)]
    ]);
    // ── D157:D159 (CL, CK, CJ) ──
    sh.getRange('D157:D159').setValues([
      [val_(audit, C_CL)], [val_(audit, C_CK)], [val_(audit, C_CJ)]
    ]);
  }
}


/**
 * BLOC 8 — Conformité Sécurité Incendie en Chaufferie.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Sécurité Incendie en Chaufferie".
 * Préserve les titres (sauts) en D167, D169, D173, D175.
 */
function writeEtatReglementaireConformiteIncendie_(reportSS, orderedLocaux, basePrefix) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité sécurité incendie en chaufferie';

  var C_CN = 91, C_CO = 92, C_CP = 93, C_CQ = 94;
  var C_CR = 95;
  var C_CV = 99, C_CW = 100, C_EL = 141;
  var C_CX = 101;
  var C_CM = 90;
  var C_BA = 52;
  var C_GC = 184;   // ⚠️ AJOUT : nouvelle ligne D172

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc8: pas d\'audit "Conformité Sécurité Incendie en Chaufferie" pour local ' + loc.idCvc);
      continue;
    }

    // ── D165:D168 (CN, CO, CP, CQ) — saut D169 ── (inchangé)
    sh.getRange('D165:D168').setValues([
      [val_(audit, C_CN)], [val_(audit, C_CO)],
      [val_(audit, C_CP)], [val_(audit, C_CQ)]
    ]);
    // ── D170 (CR) — saut D171 ── (inchangé)
    sh.getRange('D170').setValue(val_(audit, C_CR));
    // ── D172 ⚠️ NOUVELLE LIGNE (GC) ──
    sh.getRange('D172').setValue(val_(audit, C_GC));
    // ── D173:D175 (CV, CW, EL) — saut D176 ── (ex-D172:D174)
    sh.getRange('D173:D175').setValues([
      [val_(audit, C_CV)], [val_(audit, C_CW)], [val_(audit, C_EL)]
    ]);
    // ── D177 (CX) — saut D178 ── (ex-D176)
    sh.getRange('D177').setValue(val_(audit, C_CX));
    // ── D179 (CM) ── (ex-D178)
    sh.getRange('D179').setValue(val_(audit, C_CM));
    // ── D180 (BA) — donnée déplacée depuis l'électrique ── (ex-D179)
    sh.getRange('D180').setValue(val_(audit, C_BA));
  }
}




/**
 * BLOC 9 — Conformité Installation Gaz.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Installation Gaz".
 * Préserve les titres (sauts) en D185, D191, D194, D197, D202.
 */
function writeEtatReglementaireConformiteGaz_(reportSS, orderedLocaux, basePrefix) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité installation gaz';

  var C_DN = 117, C_DO = 118, C_EJ = 139;
  var C_DP = 119, C_DQ = 120, C_EH = 137, C_EI = 138, C_EK = 140;
  var C_DZ = 129, C_EA = 130;
  var C_EB = 131, C_EC = 132;
  var C_DR = 121, C_DS = 122, C_DT = 123, C_DU = 124;
  var C_DV = 125, C_DW = 126;

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc9: pas d\'audit "Conformité Installation Gaz" pour local ' + loc.idCvc);
      continue;
    }

    // ⚠️ MODIF +1
    // ── D186:D188 (DN, DO, EJ) — saut D189 ──
    sh.getRange('D186:D188').setValues([
      [val_(audit, C_DN)], [val_(audit, C_DO)], [val_(audit, C_EJ)]
    ]);
    // ── D190:D194 (DP, DQ, EH, EI, EK) — saut D195 ──
    sh.getRange('D190:D194').setValues([
      [val_(audit, C_DP)], [val_(audit, C_DQ)], [val_(audit, C_EH)],
      [val_(audit, C_EI)], [val_(audit, C_EK)]
    ]);
    // ── D196:D197 (DZ, EA) — saut D198 ──
    sh.getRange('D196:D197').setValues([
      [val_(audit, C_DZ)], [val_(audit, C_EA)]
    ]);
    // ── D199:D200 (EB, EC) — saut D201 ──
    sh.getRange('D199:D200').setValues([
      [val_(audit, C_EB)], [val_(audit, C_EC)]
    ]);
    // ── D202:D205 (DR, DS, DT, DU) — saut D206 ──
    sh.getRange('D202:D205').setValues([
      [val_(audit, C_DR)], [val_(audit, C_DS)],
      [val_(audit, C_DT)], [val_(audit, C_DU)]
    ]);
    // ── D207:D208 (DV, DW) ──
    sh.getRange('D207:D208').setValues([
      [val_(audit, C_DV)], [val_(audit, C_DW)]
    ]);
  }
}



/**
 * BLOC 10 — Conformité Installation Fioul.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Installation Fioul".
 * Préserve les titres (sauts) en D215, D220, D223.
 */
function writeEtatReglementaireConformiteFioul_(reportSS, orderedLocaux, basePrefix) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité installation fioul';

  var C_EN = 143, C_EU = 150, C_ET = 149, C_EV = 151, C_ER = 147;
  var C_EO = 144, C_EP = 145, C_ES = 148, C_FA = 156;
  var C_EM = 142, C_EQ = 146;
  var C_EW = 152, C_EX = 153, C_EY = 154, C_EZ = 155;

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc10: pas d\'audit "Conformité Installation Fioul" pour local ' + loc.idCvc);
      continue;
    }

    // ⚠️ MODIF +1
    // ── D214:D218 (EN, EU, ET, EV, ER) — saut D219 ──
    sh.getRange('D214:D218').setValues([
      [val_(audit, C_EN)], [val_(audit, C_EU)], [val_(audit, C_ET)],
      [val_(audit, C_EV)], [val_(audit, C_ER)]
    ]);
    // ── D220:D223 (EO, EP, ES, FA) — saut D224 ──
    sh.getRange('D220:D223').setValues([
      [val_(audit, C_EO)], [val_(audit, C_EP)],
      [val_(audit, C_ES)], [val_(audit, C_FA)]
    ]);
    // ── D225:D226 (EM, EQ) — saut D227 ──
    sh.getRange('D225:D226').setValues([
      [val_(audit, C_EM)], [val_(audit, C_EQ)]
    ]);
    // ── D228:D231 (EW, EX, EY, EZ) ──
    sh.getRange('D228:D231').setValues([
      [val_(audit, C_EW)], [val_(audit, C_EX)],
      [val_(audit, C_EY)], [val_(audit, C_EZ)]
    ]);
  }
}


/**
 * BLOC 11 — Conformité Froid & Climatisation.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Froid & Climatisation".
 * Préserve les titres (sauts) en D235, D237.
 * Lookup équipement (clé BC → colonne G) : image232 + nom242 + marque243 + modèle244 + série245.
 */
function writeEtatReglementaireConformiteFroid_(reportSS, orderedLocaux, basePrefix, equipByCode) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité froid & climatisation';

  var C_BI = 60, C_BF = 57, C_BD = 55, C_BE = 56;
  var C_BC = 54;

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }
  function firstKey_(v) {
    var s = String(v || '').trim();
    if (!s) return '';
    return s.split(/[,;\n]/)[0].trim();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc11: pas d\'audit "Conformité Froid & Climatisation" pour local ' + loc.idCvc);
      continue;
    }

    // ⚠️ MODIF +1
    // ── D237:D238 (BI, BF) — saut D239 ──
    sh.getRange('D237:D238').setValues([
      [val_(audit, C_BI)], [val_(audit, C_BF)]
    ]);
    // ── D240 (BD) — saut D241 ──
    sh.getRange('D240').setValue(val_(audit, C_BD));
    // ── D242 (BE) ──
    sh.getRange('D242').setValue(val_(audit, C_BE));

    // ── Lookup équipement (clé BC → colonne G) +1 ──
    var key = firstKey_(audit[C_BC]);
    if (key && equipByCode && equipByCode[key]) {
      var eq = equipByCode[key];
      if (eq.photo) sh.getRange('G236').setFormula('=IMAGE("' + eq.photo + '")');
      else sh.getRange('G236').clearContent();
      sh.getRange('G246').setValue(eq.nom   || '');
      sh.getRange('G247').setValue(eq.ag    || '');
      sh.getRange('G248').setValue(eq.ah    || '');
      sh.getRange('G249').setValue(eq.serie || '');
    } else {
      sh.getRange('G236').clearContent();
      sh.getRange('G246').clearContent();
      sh.getRange('G247').clearContent();
      sh.getRange('G248').clearContent();
      sh.getRange('G249').clearContent();
      if (key) Logger.log('   ⚠️ Bloc11: clé "' + key + '" (BC) introuvable dans Equipement');
    }
  }
}


/**
 * BLOC 12 — Conformité ECS & Légionellose.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité ECS & Légionellose".
 * Préserve les titres (sauts) en D256, D260, D266.
 * Lookup équipement (clé ED) : image G250 + nom H262 + marque H263 + modèle H264.
 */
function writeEtatReglementaireConformiteECS_(reportSS, orderedLocaux, basePrefix, equipByCode) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité ecs & légionellose';

  var C_EE = 134, C_FB = 157, C_FC = 158, C_EF = 135, C_EG = 136;
  var C_FJ = 165, C_FN = 169, C_FI = 164;
  var C_FF = 161, C_FG = 162, C_FD = 159, C_FE = 160, C_EA = 130;
  var C_FP = 171, C_FQ = 172, C_FO = 170, C_FL = 167, C_FM = 168;
  var C_ED = 133;

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }
  function firstKey_(v) {
    var s = String(v || '').trim();
    if (!s) return '';
    return s.split(/[,;\n]/)[0].trim();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc12: pas d\'audit "Conformité ECS & Légionellose" pour local ' + loc.idCvc);
      continue;
    }

    // ⚠️ MODIF +1
    // ── D255:D259 (EE, FB, FC, EF, EG) — saut D260 ──
    sh.getRange('D255:D259').setValues([
      [val_(audit, C_EE)], [val_(audit, C_FB)], [val_(audit, C_FC)],
      [val_(audit, C_EF)], [val_(audit, C_EG)]
    ]);
    // ── D261:D263 (FJ, FN, FI) — saut D264 ──
    sh.getRange('D261:D263').setValues([
      [val_(audit, C_FJ)], [val_(audit, C_FN)], [val_(audit, C_FI)]
    ]);
    // ── D265:D269 (FF, FG, FD, FE, EA) — saut D270 ──
    sh.getRange('D265:D269').setValues([
      [val_(audit, C_FF)], [val_(audit, C_FG)], [val_(audit, C_FD)],
      [val_(audit, C_FE)], [val_(audit, C_EA)]
    ]);
    // ── D271:D275 (FP, FQ, FO, FL, FM) ──
    sh.getRange('D271:D275').setValues([
      [val_(audit, C_FP)], [val_(audit, C_FQ)], [val_(audit, C_FO)],
      [val_(audit, C_FL)], [val_(audit, C_FM)]
    ]);

    // ── Lookup équipement (clé ED) +1 ──
    var key = firstKey_(audit[C_ED]);
    if (key && equipByCode && equipByCode[key]) {
      var eq = equipByCode[key];
      if (eq.photo) sh.getRange('G254').setFormula('=IMAGE("' + eq.photo + '")');
      else sh.getRange('G254').clearContent();
      sh.getRange('H266').setValue(eq.nom || '');
      sh.getRange('H267').setValue(eq.ag  || '');
      sh.getRange('H268').setValue(eq.ah  || '');
    } else {
      sh.getRange('G254').clearContent();
      sh.getRange('H266').clearContent();
      sh.getRange('H267').clearContent();
      sh.getRange('H268').clearContent();
      if (key) Logger.log('   ⚠️ Bloc12: clé "' + key + '" (ED) introuvable dans Equipement');
    }
  }
}



/**
 * BLOC 13 — Conformité Régulation & BACS.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Régulation & BACS".
 * D277:D281 contigus (aucun saut).
 */
function writeEtatReglementaireConformiteRegulation_(reportSS, orderedLocaux, basePrefix) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité régulation & bacs';

  var C_DH = 111, C_DI = 112, C_AX = 49, C_AY = 50, C_DG = 110;

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc13: pas d\'audit "Conformité Régulation & BACS" pour local ' + loc.idCvc);
      continue;
    }

    // ⚠️ MODIF +1 : D281:D285
    sh.getRange('D281:D285').setValues([
      [val_(audit, C_DH)],   // D281
      [val_(audit, C_DI)],   // D282
      [val_(audit, C_AX)],   // D283
      [val_(audit, C_AY)],   // D284
      [val_(audit, C_DG)]    // D285
    ]);
  }
}


/**
 * BLOC 14 — Conformité Ventilations Chaufferie.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Ventilations Chaufferie".
 * D287:D290 contigus (aucun saut).
 */
function writeEtatReglementaireConformiteVentilation_(reportSS, orderedLocaux, basePrefix) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité ventilations chaufferie';

  var C_DJ = 113, C_DK = 114, C_DL = 115, C_DM = 116;

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc14: pas d\'audit "Conformité Ventilations Chaufferie" pour local ' + loc.idCvc);
      continue;
    }

    // ⚠️ MODIF +1 : D291:D294
    sh.getRange('D291:D294').setValues([
      [val_(audit, C_DJ)],   // D291
      [val_(audit, C_DK)],   // D292
      [val_(audit, C_DL)],   // D293
      [val_(audit, C_DM)]    // D294
    ]);
  }
}



/**
 * BLOC 15 — Conformité Chaudière Bois & Biomasse.
 * Cherche dans loc.auditRows la ligne dont FX = "Conformité Chaudière Bois & Biomasse".
 * Préserve le titre (saut) en D299.
 */
function writeEtatReglementaireConformiteBiomasse_(reportSS, orderedLocaux, basePrefix) {
  var C_FX = 179;
  var TYPE_CIBLE = 'conformité chaudière bois & biomasse';

  var C_FR = 173, C_FS = 174, C_FT = 175, C_FU = 176, C_FV = 177, C_FW = 178;

  function val_(row, idx) { return String(row[idx] || '').trim(); }
  function normFx_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  for (var i = 0; i < orderedLocaux.length; i++) {
    var idx = i + 1;
    var sh  = getSheetByNameLoose_(reportSS, basePrefix + idx);
    if (!sh) continue;

    var loc = orderedLocaux[i];
    var audit = null;
    for (var a = 0; a < loc.auditRows.length; a++) {
      if (normFx_(loc.auditRows[a][C_FX]) === TYPE_CIBLE) { audit = loc.auditRows[a]; break; }
    }
    if (!audit) {
      Logger.log('   ⏭️ Bloc15: pas d\'audit "Conformité Chaudière Bois & Biomasse" pour local ' + loc.idCvc);
      continue;
    }

    // ⚠️ MODIF +1
    // ── D300:D302 (FW, FV, FR) — saut D303 ──
    sh.getRange('D300:D302').setValues([
      [val_(audit, C_FW)], [val_(audit, C_FV)], [val_(audit, C_FR)]
    ]);
    // ── D304:D306 (FS, FT, FU) ──
    sh.getRange('D304:D306').setValues([
      [val_(audit, C_FS)], [val_(audit, C_FT)], [val_(audit, C_FU)]
    ]);
  }
}



/****************************************************
 * ═══════════════════════════════════════════════
 * FONCTIONS BUILD (construction des index)
 * ═══════════════════════════════════════════════
 ****************************************************/

function buildClientsAvecNonDebute_(dataPEC) {
  var map = {};

  for (var i = 1; i < dataPEC.length; i++) {
    var nomClient  = dataPEC[i][0];
    var codeClient = dataPEC[i][1];
    var etat       = String(dataPEC[i][4] || '').trim();

    if (!codeClient || !nomClient) continue;
    if (etat !== 'Prise en charge non débutée') continue;

    var key = String(codeClient).trim();
    if (!map[key]) {
      map[key] = { nomClient: String(nomClient).trim(), codeClient: key };
    }
  }

  return map;
}

function getAllSitesFromPerimetre_(reportSS) {
  var sh = getSheetByNameLoose_(reportSS, '1.Périmètre de la prise en charge');
  if (!sh) return [];

  var startRow = 11;
  var last = sh.getLastRow();
  if (last < startRow) return [];

  var n = last - startRow + 1;
  var vals = sh.getRange(startRow, 1, n, 3).getValues();

  var out = [];
  var seen = {};

  for (var i = 0; i < vals.length; i++) {
    var codeSite = String(vals[i][0] || '').trim();
    var nomSite  = String(vals[i][2] || '').trim();
    if (!codeSite || !nomSite) continue;

    var key = codeSite;
    if (seen[key]) continue;
    seen[key] = true;

    out.push({ codeSite: codeSite, nomSite: nomSite });
  }

  out.sort(function(a,b){ return (a.nomSite||'').localeCompare(b.nomSite||''); });
  return out;
}

function buildFinalizedSitesByClient_(dataPEC) {
  var map = {};

  for (var i = 1; i < dataPEC.length; i++) {
    var codeClient = dataPEC[i][1];
    var codeSite = dataPEC[i][2];
    var nomSite = dataPEC[i][3];
    var etat = dataPEC[i][4];

    if (etat !== 'Prise en charge finalisée') continue;
    if (!codeClient || !codeSite || !nomSite) continue;

    var cKey = String(codeClient).trim();
    var sKey = String(codeSite).trim();
    var nKey = String(nomSite).trim();

    if (!map[cKey]) {
      map[cKey] = { list: [], codeSiteSet: {}, siteNameSet: {} };
    }

    if (!map[cKey].codeSiteSet[sKey]) {
      map[cKey].codeSiteSet[sKey] = true;
      map[cKey].list.push({ codeSite: sKey, nomSite: nKey });
    }

    map[cKey].siteNameSet[nKey] = true;
  }

  for (var c in map) {
    map[c].list.sort(function(a, b) {
      return (a.nomSite || '').localeCompare(b.nomSite || '');
    });
  }

  return map;
}

function buildPhotosByClient_(dataPhoto) {
  var photosByClient = {};

  for (var i = 1; i < dataPhoto.length; i++) {
    var codeClient = dataPhoto[i][2];
    if (!codeClient) continue;
    var key = String(codeClient).trim();

    var siteName = String(dataPhoto[i][4] || '').trim();
    var qhseVal = dataPhoto[i][5];
    var url = String(dataPhoto[i][7] || '').trim();
    var title = String(dataPhoto[i][8] || '').trim();

    var qhseText = '';
    if (qhseVal === true || String(qhseVal).toUpperCase() === 'TRUE') qhseText = 'Oui';
    else if (qhseVal === false || String(qhseVal).toUpperCase() === 'FALSE') qhseText = 'Non';
    else qhseText = String(qhseVal || '').trim();

    if (!photosByClient[key]) photosByClient[key] = [];
    photosByClient[key].push({ siteName: siteName, title: title, url: url, qhse: qhseText });
  }

  return photosByClient;
}

function buildInfoSiteByCodeSite_(dataPEC) {
  var map = {};

  var colCodeSite = -1;
  var colInfoSite = -1;

  for (var c = 0; c < dataPEC[0].length; c++) {
    var header = String(dataPEC[0][c] || '').trim().toLowerCase();

    if ((header.indexOf('code') !== -1 && header.indexOf('site') !== -1) && colCodeSite === -1) {
      colCodeSite = c;
    }

    if ((header.indexOf('info') !== -1 && header.indexOf('site') !== -1) && colInfoSite === -1) {
      colInfoSite = c;
    }
  }

  if (colCodeSite === -1) colCodeSite = 2;
  if (colInfoSite === -1) colInfoSite = 6;

  for (var i = 1; i < dataPEC.length; i++) {
    var codeSite = String(dataPEC[i][colCodeSite] || '').trim();
    var infoSite = String(dataPEC[i][colInfoSite] || '').trim();

    if (!codeSite) continue;
    if (!map[codeSite]) {
      map[codeSite] = infoSite;
    }
  }

  Logger.log('buildInfoSiteByCodeSite_: ' + Object.keys(map).length + ' sites indexés');
  return map;
}

function buildEquipAgg_(dataEquip) {
  var EQ_CODECLIENT = 1;
  var EQ_CODESITE = 3;
  var EQ_QTE = 8;
  var EQ_NONCONFORME = 28;
  var EQ_ETAT = 41;

  var equipAgg = {};

  for (var i = 1; i < dataEquip.length; i++) {
    var codeClient = dataEquip[i][EQ_CODECLIENT];
    var codeSite = dataEquip[i][EQ_CODESITE];
    if (!codeClient || !codeSite) continue;

    var key = String(codeClient).trim() + '||' + String(codeSite).trim();
    if (!equipAgg[key]) {
      equipAgg[key] = { totalQty: 0, nonConforme: 0, mauvaisEtat: 0, hsProv: 0 };
    }

    equipAgg[key].totalQty += toNumberRobust_(dataEquip[i][EQ_QTE]);

    if (String(dataEquip[i][EQ_NONCONFORME] || '').trim() === 'Non conforme') {
      equipAgg[key].nonConforme++;
    }

    var etat = String(dataEquip[i][EQ_ETAT] || '').trim();
    if (etat === 'En service - Mauvais état') equipAgg[key].mauvaisEtat++;
    if (etat === 'Hors service - Arrêt provisoire') equipAgg[key].hsProv++;
  }

  return equipAgg;
}

function buildEcartInventaireByClientSite_(dataEquip) {
  var EQ_CODECLIENT = 1;
  var EQ_CODESITE   = 3;
  var EQ_PRESENCE   = 5;
  var EQ_QTE        = 8;

  function norm_(v) {
    return String(v || '')
      .replace(/[\u00A0\u202F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  }

  var map = {};

  for (var i = 1; i < dataEquip.length; i++) {
    var codeClient = dataEquip[i][EQ_CODECLIENT];
    var codeSite   = dataEquip[i][EQ_CODESITE];
    if (!codeClient || !codeSite) continue;

    var presenceN = norm_(dataEquip[i][EQ_PRESENCE]);
    if (!presenceN) continue;

    var qte = toNumberRobust_(dataEquip[i][EQ_QTE]);

    var delta = 0;
    if (presenceN.indexOf('ajouter') !== -1) {
      delta = qte;
    }
    else if (presenceN.indexOf('supprimer') !== -1) {
      delta = -qte;
    } else {
      continue;
    }

    var key = String(codeClient).trim() + '||' + String(codeSite).trim();
    map[key] = (map[key] || 0) + delta;
  }

  return map;
}

function buildEquipRowsByClientSite_(dataEquip) {
  var EQ_CODECLIENT = 1;
  var EQ_CODESITE = 3;
  var EQ_BATIMENT = 6;
  var EQ_INVENTAIRE = 46;   // AU (commentaire)
  var EQ_PRESENCE = 5;
  var EQ_NOMEQUIP = 7;
  var EQ_QTE = 8;
  var EQ_LOCAL = 10;
  var EQ_PHOTO1 = 26;
  var EQ_PHOTO2 = 27;
  var EQ_CONFORMITE = 28;
  var EQ_MARQUE = 32;       // AG
  var EQ_MODELE = 33;       // AH
  var EQ_ETAT = 41;         // AP
  var EQ_TYPOLOGIE = 16;

  // ✅ NOUVEAUX CHAMPS
  var EQ_DATE_MES  = 31;    // AF = date de mise en service
  var EQ_PUISSANCE = 34;    // AI = puissance
  var EQ_NUM_SERIE = 36;    // AK = numéro de série

  // 🆕 COLONNES POUR LE COMMENTAIRE ENRICHI (colonne I)
  var EQ_COMMENT_AS = 44;   // AS
  var EQ_COMMENT_AM = 38;   // AM

  var map = {};

  for (var i = 1; i < dataEquip.length; i++) {
    var codeClient = dataEquip[i][EQ_CODECLIENT];
    var codeSite = dataEquip[i][EQ_CODESITE];
    if (!codeClient || !codeSite) continue;

    var key = String(codeClient).trim() + '||' + String(codeSite).trim();
    if (!map[key]) map[key] = [];

    map[key].push({
      nomEquip: String(dataEquip[i][EQ_NOMEQUIP] || '').trim(),
      presence: String(dataEquip[i][EQ_PRESENCE] || '').trim(),
      batiment: String(dataEquip[i][EQ_BATIMENT] || '').trim(),
      local: String(dataEquip[i][EQ_LOCAL] || '').trim(),
      marque: String(dataEquip[i][EQ_MARQUE] || '').trim(),
      modele: String(dataEquip[i][EQ_MODELE] || '').trim(),
      inventaire: String(dataEquip[i][EQ_INVENTAIRE] || '').trim(),
      qte: dataEquip[i][EQ_QTE],
      etat: String(dataEquip[i][EQ_ETAT] || '').trim(),
      conformite: String(dataEquip[i][EQ_CONFORMITE] || '').trim(),
      photo1: String(dataEquip[i][EQ_PHOTO1] || '').trim(),
      photo2: String(dataEquip[i][EQ_PHOTO2] || '').trim(),
      typologie: String(dataEquip[i][EQ_TYPOLOGIE] || '').trim(),

      // ✅ NOUVEAUX CHAMPS
      numeroSerie: String(dataEquip[i][EQ_NUM_SERIE] || '').trim(),
      puissance:   String(dataEquip[i][EQ_PUISSANCE] || '').trim(),
      dateMES:     dataEquip[i][EQ_DATE_MES],  // gardé brut (date)

      // 🆕 CHAMPS COMMENTAIRE ENRICHI
      commentAS: String(dataEquip[i][EQ_COMMENT_AS] || '').trim(),
      commentAM: String(dataEquip[i][EQ_COMMENT_AM] || '').trim()
    });
  }

  return map;
}


function buildCompteurAgg_(dataEquip) {
  var EQ_CODECLIENT = 1;
  var EQ_CODESITE = 3;
  var EQ_TYPE = 11;
  var EQ_ETAT = 41;

  var TYPE_TOKEN = 'COMPTR';
  var HS_TOKEN = 'Hors service';

  var map = {};

  for (var i = 1; i < dataEquip.length; i++) {
    var codeClient = dataEquip[i][EQ_CODECLIENT];
    var codeSite = dataEquip[i][EQ_CODESITE];
    if (!codeClient || !codeSite) continue;

    var type = String(dataEquip[i][EQ_TYPE] || '').trim();
    if (type.indexOf(TYPE_TOKEN) === -1) continue;

    var key = String(codeClient).trim() + '||' + String(codeSite).trim();
    if (!map[key]) map[key] = { totalCompteurs: 0, horsService: 0 };

    map[key].totalCompteurs++;

    var etat = String(dataEquip[i][EQ_ETAT] || '').trim();
    if (etat.indexOf(HS_TOKEN) !== -1) {
      map[key].horsService++;
    }
  }

  return map;
}

function buildCompteurRowsByClient_(dataEquip, shEquip) {
  var EQ_CODECLIENT      = 1;
  var EQ_CODESITE        = 3;
  var EQ_NOMSITE         = 4;
  var EQ_TYPE            = 11;
  var EQ_NOMEQUIP        = 7;
  var EQ_LOCAL           = 10;
  var EQ_TYPE_COMPTEUR_S = 18;
  var EQ_ETAT_AP         = 41;
  var EQ_INDEX           = 20;
  var EQ_UNITE           = 21;
  var EQ_PDL             = 23;
  var EQ_PHOTO1          = 26;
  var EQ_PHOTO2          = 27;
  var EQ_DATE_REL        = 47;
  var TYPE_TOKEN         = 'COMPTR';

  var map = {};

  for (var i = 1; i < dataEquip.length; i++) {
    var codeClient = dataEquip[i][EQ_CODECLIENT];
    if (!codeClient) continue;

    var type = String(dataEquip[i][EQ_TYPE] || '').trim();
    if (type.indexOf(TYPE_TOKEN) === -1) continue;

    var key      = String(codeClient).trim();
    var codeSite = String(dataEquip[i][EQ_CODESITE] || '').trim();

    var photo1 = String(dataEquip[i][EQ_PHOTO1] || '').trim();
    var photo2 = String(dataEquip[i][EQ_PHOTO2] || '').trim();

    if (!map[key]) map[key] = [];

    map[key].push({
      codeSite:      codeSite,
      nomSite:       String(dataEquip[i][EQ_NOMSITE]         || '').trim(),
      nomEquip:      String(dataEquip[i][EQ_NOMEQUIP]        || '').trim(),
      local:         String(dataEquip[i][EQ_LOCAL]           || '').trim(),
      typeCompteurS: String(dataEquip[i][EQ_TYPE_COMPTEUR_S] || '').trim(),
      etatEquip:     String(dataEquip[i][EQ_ETAT_AP]         || '').trim(),
      indexCompteur: dataEquip[i][EQ_INDEX],
      unite:         String(dataEquip[i][EQ_UNITE]           || '').trim(),
      dateReleve:    dataEquip[i][EQ_DATE_REL],
      pdl:           String(dataEquip[i][EQ_PDL]             || '').trim(),
      photo1:        photo1,
      photo2:        photo2
    });
  }

  Logger.log('buildCompteurRowsByClient_: ' + Object.keys(map).length + ' client(s)');
  return map;
}

function buildTravauxAgg_(dataTravaux) {
  var PROP_CODECLIENT = 2;
  var PROP_CODESITE = 3;
  var PROP_NOMSITE = 4;
  var PROP_MOTIF = 5;
  var PROP_MONTANT = 18;

  var map = {};

  for (var i = 1; i < dataTravaux.length; i++) {
    var codeClient = dataTravaux[i][PROP_CODECLIENT];
    var codeSite = dataTravaux[i][PROP_CODESITE];
    if (!codeClient || !codeSite) continue;

    var key = String(codeClient).trim() + '||' + String(codeSite).trim();
    if (!map[key]) {
      map[key] = {
       nomSite: String(dataTravaux[i][PROP_NOMSITE] || '').trim(),
       total: 0,
       reg: 0,
       prevAcc: 0,
       amelio: 0,
       remise: 0,
       montant: 0
     };
    }

    map[key].total++;

    var motif = String(dataTravaux[i][PROP_MOTIF] || '');
    if (motif.indexOf('Réglementaire') !== -1) map[key].reg++;
    if (motif.indexOf('Remise en état') !== -1) map[key].remise++;
    if (motif.indexOf('Prévention accident') !== -1) map[key].prevAcc++;
    if (motif.indexOf('Amélioration') !== -1) map[key].amelio++;

    map[key].montant += toNumberRobust_(dataTravaux[i][PROP_MONTANT]);
  }

  return map;
}

function buildTravauxRowsByClient_(dataTravaux, shTravaux) {
  var T_CODECLIENT  = 2;
  var T_CODESITE    = 3;
  var T_NOMSITE     = 4;
  var T_MOTIF       = 5;
  var T_DESIGNATION = 6;
  var T_EQUIPEMENT  = 7;
  var T_IMG_K       = 10;
  var T_IMG_L       = 11;
  var T_MONTANT     = 18;

  var map = {};

  for (var i = 1; i < dataTravaux.length; i++) {
    var codeClient = dataTravaux[i][T_CODECLIENT];
    if (!codeClient) continue;

    var key = String(codeClient).trim();

    var imageUrlK = String(dataTravaux[i][T_IMG_K] || '').trim();
    var imageUrlL = String(dataTravaux[i][T_IMG_L] || '').trim();

    if (!map[key]) map[key] = [];

    map[key].push({
      codeSite:           String(dataTravaux[i][T_CODESITE]    || '').trim(),
      nomSite:            String(dataTravaux[i][T_NOMSITE]      || '').trim(),
      equipement:         String(dataTravaux[i][T_EQUIPEMENT]   || '').trim(),
      motif:              String(dataTravaux[i][T_MOTIF]        || '').trim(),
      designationTravaux: String(dataTravaux[i][T_DESIGNATION]  || '').trim(),
      imageUrlK:          imageUrlK,
      imageUrlL:          imageUrlL,
      montant:            dataTravaux[i][T_MONTANT]
    });
  }

  Logger.log('buildTravauxRowsByClient_: ' + Object.keys(map).length + ' client(s)');
  return map;
}

/****************************************************
 * ═══════════════════════════════════════════════
 * FONCTIONS WRITE (écriture dans les rapports)
 * ═══════════════════════════════════════════════
 ****************************************************/

function writePhotoLibrary_(sh, rows) {
  var startRow = 11;

  var last = sh.getLastRow();
  if (last >= startRow) {
    sh.getRange(startRow, 2, last - startRow + 1, 10).clearContent();
  }
  if (!rows || rows.length === 0) return;

  var neededLastRow = startRow - 1 + rows.length;
  var maxRows = sh.getMaxRows();

  if (maxRows < neededLastRow) {
    sh.insertRowsAfter(maxRows, neededLastRow - maxRows);
    for (var r = maxRows + 1; r <= neededLastRow; r++) {
      safeMerge_(sh, r, 3, 2);
      safeMerge_(sh, r, 5, 2);
      safeMerge_(sh, r, 7, 3);
    }
  }

  var colB = [], colC = [], colE = [], colJ = [], colK = [];
  for (var i = 0; i < rows.length; i++) {
    colB.push([i + 1]);
    colC.push([rows[i].siteName || '']);
    colE.push([rows[i].title   || '']);
    colJ.push([rows[i].qhse   || '']);
    colK.push([rows[i].url    || '']);
  }

  sh.getRange(startRow, 2,  rows.length, 1).setValues(colB);
  sh.getRange(startRow, 3,  rows.length, 1).setValues(colC);
  sh.getRange(startRow, 5,  rows.length, 1).setValues(colE);
  sh.getRange(startRow, 10, rows.length, 1).setValues(colJ);

  sh.getRange(startRow, 11, rows.length, 1)
    .setValues(colK)
    .setFontColor('#FFFFFF');

  var formulasG = rows.map(function(r) {
    var url = (r && r.url) ? String(r.url).trim() : '';
    return [url ? '=IMAGE("' + url + '")' : ''];
  });
  sh.getRange(startRow, 7, rows.length, 1).setFormulas(formulasG);
}

function writeInventaireGeneral_(sh, codeClient, sitesFinalises, equipAgg, finalizedCodeSiteSet, ecartInventaireByClientSite) {
  var startRow = 15;
  var last = sh.getLastRow();

  if (last >= startRow) sh.getRange(startRow, 2, last - startRow + 1, 9).clearContent();
  if (!sitesFinalises || sitesFinalises.length === 0) return;

  var neededLastRow = startRow - 1 + sitesFinalises.length;
  var maxRows = sh.getMaxRows();
  if (maxRows < neededLastRow) {
    sh.insertRowsAfter(maxRows, neededLastRow - maxRows);
    for (var r = maxRows + 1; r <= neededLastRow; r++) safeMerge_(sh, r, 3, 3);
  }

  var colB = [], colC = [], colF = [], colG = [], colH = [], colI = [], colJ = [];

  for (var i = 0; i < sitesFinalises.length; i++) {
    var codeSite = String(sitesFinalises[i].codeSite || '').trim();
    var nomSite = sitesFinalises[i].nomSite;

    var isFinal = finalizedCodeSiteSet && finalizedCodeSiteSet[codeSite];
    var agg = { totalQty: 0, nonConforme: 0, mauvaisEtat: 0, hsProv: 0 };

    if (isFinal) {
      var key = String(codeClient).trim() + '||' + codeSite;
      agg = equipAgg[key] || agg;
    }

    var eKey = String(codeClient).trim() + '||' + codeSite;
    var ecart = (ecartInventaireByClientSite && ecartInventaireByClientSite[eKey]) ? ecartInventaireByClientSite[eKey] : 0;

    colB.push([i + 1]);
    colC.push([nomSite || '']);
    colF.push([agg.totalQty]);
    colG.push([ecart]);
    colH.push([agg.nonConforme]);
    colI.push([agg.mauvaisEtat]);
    colJ.push([agg.hsProv]);
  }

  sh.getRange(startRow, 2, sitesFinalises.length, 1).setValues(colB);
  sh.getRange(startRow, 3, sitesFinalises.length, 1).setValues(colC);
  sh.getRange(startRow, 6, sitesFinalises.length, 1).setValues(colF);
  sh.getRange(startRow, 7, sitesFinalises.length, 1).setValues(colG);
  sh.getRange(startRow, 8, sitesFinalises.length, 1).setValues(colH);
  sh.getRange(startRow, 9, sitesFinalises.length, 1).setValues(colI);
  sh.getRange(startRow, 10, sitesFinalises.length, 1).setValues(colJ);
}

function ensureInventaireParSiteSheets_(reportSS, finalizedSitesList) {
  var templateName = '3. Inventaire des équipements par site';
  var basePrefix = '3. Inventaire des équipements par site ';

  var templateSheet = getSheetByNameLoose_(reportSS, templateName);
  if (!templateSheet) {
    Logger.log('Template introuvable: ' + templateName);
    return { basePrefix: basePrefix, orderedSitesForSheets: [] };
  }

  var nameBySiteCode = {};
  for (var i = 0; i < finalizedSitesList.length; i++) {
    var c = String(finalizedSitesList[i].codeSite || '').trim();
    var n = String(finalizedSitesList[i].nomSite  || '').trim();
    if (c && n && !nameBySiteCode[c]) nameBySiteCode[c] = n;
  }

  var perimetreIdx = buildPerimetreIndexFromSheet_(reportSS);

  var orderedSites = [];
  var shPer = getSheetByNameLoose_(reportSS, '1.Périmètre de la prise en charge');

  if (shPer) {
    var startRow = 11;
    var last = shPer.getLastRow();
    if (last >= startRow) {
      var nRows = last - startRow + 1;
      var vals = shPer.getRange(startRow, 1, nRows, 3).getValues();

      var seen = {};
      for (var r = 0; r < vals.length; r++) {
        var codeSite = String(vals[r][0] || '').trim();
        if (!codeSite) continue;
        if (!nameBySiteCode[codeSite]) continue;
        if (seen[codeSite]) continue;
        seen[codeSite] = true;
        orderedSites.push({ codeSite: codeSite, nomSite: nameBySiteCode[codeSite] });
      }
    }
  }

  if (orderedSites.length === 0) {
    for (var z = 0; z < finalizedSitesList.length; z++) {
      orderedSites.push({
        codeSite: String(finalizedSitesList[z].codeSite || '').trim(),
        nomSite:  String(finalizedSitesList[z].nomSite  || '').trim()
      });
    }
  }

  var templatePos = templateSheet.getIndex();

  for (var i2 = 0; i2 < orderedSites.length; i2++) {
    var idx = i2 + 1;
    var targetName = basePrefix + idx;

    var sh = getSheetByNameLoose_(reportSS, targetName);
    if (!sh) {
      sh = templateSheet.copyTo(reportSS);
      sh.setName(targetName);
      reportSS.setActiveSheet(sh);
      reportSS.moveActiveSheet(templatePos + i2 + 1);
    }

    sh.getRange('A4').setValue(orderedSites[i2].nomSite);
    sh.getRange('D7')
      .setValue(orderedSites[i2].codeSite)
      .setFontColor('#FFFFFF');
  }

  var j = orderedSites.length + 1;
  while (true) {
    var extra = getSheetByNameLoose_(reportSS, basePrefix + j);
    if (!extra) break;

    var existingD7 = String(extra.getRange('D7').getValue() || '').trim();

    if (!existingD7) {
      var nomA4 = String(extra.getRange('A4').getValue() || '').trim();
      var codeFromA4 = nomA4 ? (perimetreIdx.codeByName[nomA4] || '') : '';

      if (codeFromA4) {
        extra.getRange('D7')
          .setValue(codeFromA4)
          .setFontColor('#FFFFFF');
        Logger.log('D7 renseigné pour feuille extra ' + j + ': ' + codeFromA4);
        j++;
      } else {
        Logger.log('Suppression feuille en trop: ' + basePrefix + j);
        reportSS.deleteSheet(extra);
      }
    } else {
      j++;
    }
  }

  try {
    templateSheet.hideSheet();
  } catch (e) {
    Logger.log('Impossible de masquer le calque "' + templateName + '" : ' + e);
  }

  return { basePrefix: basePrefix, orderedSitesForSheets: orderedSites };
}

function writePerimetreLinks_(reportSS, basePrefix) {
  var sh = getSheetByNameLoose_(reportSS, '1.Périmètre de la prise en charge');
  if (!sh) return;

  var startRow = 11;
  var colSiteName = 3;
  var colLink = 11;

  var last = sh.getLastRow();
  if (last < startRow) return;

  var n = last - startRow + 1;
  var names = sh.getRange(startRow, colSiteName, n, 1).getValues();

  sh.getRange(startRow, colLink, n, 1).clearContent();

  var sheetIdx = 0;
  for (var i = 0; i < names.length; i++) {
    var nomSite = String(names[i][0] || '').trim();
    if (!nomSite) continue;

    sheetIdx++;
    var target = getSheetByNameLoose_(reportSS, basePrefix + sheetIdx);
    if (!target) continue;

    var gid = target.getSheetId();
    sh.getRange(startRow + i, colLink).setFormula('=HYPERLINK("#gid=' + gid + '";"Ouvrir")');
  }
}

function fillInventaireParSiteEquipements_(reportSS, codeClient, orderedSitesForSheets, equipRowsByClientSite, nonDebuteeSet, infoSiteByCodeSite) {
  var templatePrefix = '3. Inventaire des équipements par site ';
  var globalCounter = 1;

  // ✅ Nb de colonnes texte (A..L) et total (A..N avec images M/N)
  var NUM_TEXT_COLS = 12;   // A..L
  var NUM_ALL_COLS  = 14;   // A..N
  var IMG_START_COL = 13;   // M

  var TYPO_ORDER = [
    'Production de chaleur',
    'Circuit chauffage',
    "Production d'eau chaude sanitaire",
    "Circuit d'eau chaude sanitaire",
    'Expansion',
    "Traitement d'eaux",
    'Production de froid',
    'Circuit froid',
    'Production réversible (Chaud/Froid)',
    'Circuit réversible (Chaud/Froid)',
    'Émetteur (radiants, bouches de soufflage, clim intérieur, ventilo-convecteur...)',
    'Relevage',
    'Electricité',
    'Régulation',
    "Traitement d'air",
    'Matériels annexes',
    'Equipement sans rattachement'
  ];

  function getTypoRank_(typo) {
    var t = String(typo || '').trim().toLowerCase();
    for (var idx = 0; idx < TYPO_ORDER.length; idx++) {
      if (TYPO_ORDER[idx].toLowerCase() === t) return idx;
    }
    return TYPO_ORDER.length - 2;
  }

  function withRetryLocal_(fn, label) {
    for (var a = 1; a <= 4; a++) {
      try {
        return fn();
      } catch (e) {
        Logger.log('Sheets retry ' + a + '/4 - ' + (label || '') + ' : ' + e);
        Utilities.sleep(400 * a);
        if (a === 4) throw e;
      }
    }
  }

  function clearAllRowGroups_(sh, fromRow, toRow) {
    try {
      for (var row = fromRow; row <= toRow; row++) {
        for (var depth = 5; depth >= 1; depth--) {
          try {
            var grp = sh.getRowGroup(row, depth);
            if (grp) grp.remove();
          } catch(e) {}
        }
      }
    } catch(e) {
      Logger.log('clearAllRowGroups_ warn: ' + e);
    }
  }

  function cleanSheet_(sh, startRow) {
    var last = sh.getLastRow();
    if (last < startRow) return;
    var n = last - startRow + 1;
    clearAllRowGroups_(sh, startRow, last);
    try { sh.getRange(startRow, 1, n, NUM_ALL_COLS).breakApart(); } catch(e) {}
    try { sh.getRange(startRow, 1, n, NUM_ALL_COLS).clearContent(); } catch(e) {}
    try { sh.getRange(startRow, 1, n, NUM_ALL_COLS).setBackground('#FFFFFF'); } catch(e) {}
    try { sh.getRange(startRow, 1, n, NUM_ALL_COLS).setFontColor('#000000'); } catch(e) {}
    try { sh.getRange(startRow, 1, n, NUM_ALL_COLS).setFontWeight('normal'); } catch(e) {}
    try { sh.getRange(startRow, 1, n, NUM_ALL_COLS).setHorizontalAlignment('left'); } catch(e) {}
    try { sh.setRowHeightsForced(startRow, n, 250); } catch(e) {}
  }

  // 🆕 Construit le commentaire enrichi (colonne I) : AU | AS | Type de fluide : AM
  function buildCommentaire_(eq) {
    var parts = [];
    if (eq.inventaire) parts.push(eq.inventaire);                     // AU (commentaire libre) — toujours en premier
    if (eq.commentAS)  parts.push(eq.commentAS);                      // AS (valeur brute)
    if (eq.commentAM)  parts.push('Type de fluide : ' + eq.commentAM); // AM
    return parts.join(' | ');
  }

  // ✅ Helper : construit la ligne texte A..L (12 colonnes) pour un équipement
  function buildRowValues_(eq) {
    var bat = String(eq.batiment || '').trim();
    var loc = String(eq.local || '').trim();
    var batLoc = (bat && loc) ? (bat + ' - ' + loc) : (bat || loc);
    var marque = String(eq.marque || '').trim();
    var modele = String(eq.modele || '').trim();
    var marqueModele = (marque && modele) ? (marque + ' - ' + modele) : (marque || modele);

    return [
      globalCounter++,            // A : numéro
      eq.nomEquip    || '',       // B : nom équipement
      eq.presence    || '',       // C : présence
      batLoc,                     // D : bâtiment - local
      marqueModele,               // E : marque - modèle
      eq.numeroSerie || '',       // F : 🆕 numéro de série
      eq.puissance   || '',       // G : 🆕 puissance
      eq.dateMES     || '',       // H : 🆕 date mise en service
      buildCommentaire_(eq),      // I : 🆕 commentaire enrichi (AU + AS + AM)
      eq.qte         || '',       // J : nombre
      eq.etat        || '',       // K : état
      eq.conformite  || ''        // L : conformité
    ];
  }

  // ✅ Helper : construit les formules images M..N (2 colonnes)
  function buildRowFormulas_(eq) {
    var p1 = eq.photo1 ? String(eq.photo1).trim() : '';
    var p2 = eq.photo2 ? String(eq.photo2).trim() : '';
    return [
      p1 ? '=IMAGE("' + p1 + '")' : '',   // M : photo 1
      p2 ? '=IMAGE("' + p2 + '")' : ''    // N : photo 2
    ];
  }

  var startRow = 19;

  for (var idx = 1; ; idx++) {
    var sh = getSheetByNameLoose_(reportSS, templatePrefix + idx);
    if (!sh) break;

    var codeSite = String(sh.getRange('D7').getValue() || '').trim();

    if (!codeSite) {
      Logger.log('InvParSite: D7 vide pour feuille ' + idx + ' → skip');
      continue;
    }

    var infoSite = (infoSiteByCodeSite && infoSiteByCodeSite[codeSite]) 
                    ? infoSiteByCodeSite[codeSite] 
                    : '';
    
    try {
      sh.getRange('E14').setValue(infoSite);
    } catch(e) {
      Logger.log('WARN: Impossible écrire E14 feuille ' + idx + ' : ' + e);
    }

    if (nonDebuteeSet && nonDebuteeSet[codeSite]) {
      withRetryLocal_(function() {
        cleanSheet_(sh, startRow);
      }, 'InvParSite clean non débuté ' + idx);
      SpreadsheetApp.flush();
      continue;
    }

    var key = String(codeClient).trim() + '||' + codeSite;
    var rows = equipRowsByClientSite[key] || [];

    withRetryLocal_(function() {
      cleanSheet_(sh, startRow);
    }, 'InvParSite clean ' + idx);

    if (rows.length === 0) {
      SpreadsheetApp.flush();
      continue;
    }

    var byTypo = {};
    var typoList = [];

    for (var r = 0; r < rows.length; r++) {
      var typo = String(rows[r].typologie || '').trim();
      if (!typo) typo = 'Equipement sans rattachement';
      if (!byTypo[typo]) {
        byTypo[typo] = [];
        typoList.push(typo);
      }
      byTypo[typo].push(rows[r]);
    }

    typoList.sort(function(a, b) {
      return getTypoRank_(a) - getTypoRank_(b);
    });

    var currentRow = startRow;
    var hasTypologies = typoList.length > 1 ||
      (typoList.length === 1 && typoList[0] !== 'Equipement sans rattachement');

    // ── CAS SANS TYPOLOGIE ──────────────────────────────────
    if (!hasTypologies) {
      var neededLastRow = currentRow - 1 + rows.length;
      var maxRows = sh.getMaxRows();
      if (maxRows < neededLastRow) sh.insertRowsAfter(maxRows, neededLastRow - maxRows);

      var values = new Array(rows.length);
      var formulasMN = new Array(rows.length);

      for (var r = 0; r < rows.length; r++) {
        values[r]     = buildRowValues_(rows[r]);
        formulasMN[r] = buildRowFormulas_(rows[r]);
      }

      // ✅ Texte A..L (12 colonnes)
      withRetryLocal_(function() {
        sh.getRange(currentRow, 1, rows.length, NUM_TEXT_COLS).setValues(values);
      }, 'InvParSite setValues ' + idx);

      // ✅ Images M..N (2 colonnes)
      withRetryLocal_(function() {
        sh.getRange(currentRow, IMG_START_COL, rows.length, 2).setFormulas(formulasMN);
      }, 'InvParSite setFormulas ' + idx);

      // ✅ Noir sur fond blanc sur A..N
      sh.getRange(currentRow, 1, rows.length, NUM_ALL_COLS)
        .setFontColor('#000000')
        .setBackground('#FFFFFF');

      SpreadsheetApp.flush();
      continue;
    }

    // ── CAS AVEC TYPOLOGIES ─────────────────────────────────
    var totalNeeded = 0;
    for (var t = 0; t < typoList.length; t++) {
      totalNeeded += 1 + byTypo[typoList[t]].length;
    }

    var neededLastRow = currentRow - 1 + totalNeeded;
    var maxRows = sh.getMaxRows();
    if (maxRows < neededLastRow) sh.insertRowsAfter(maxRows, neededLastRow - maxRows);

    for (var t = 0; t < typoList.length; t++) {
      var typoName      = typoList[t];
      var typoEquips    = byTypo[typoName];
      var titreRow      = currentRow;
      var firstEquipRow = currentRow + 1;
      var lastEquipRow  = currentRow + typoEquips.length;

      // ✅ Titre typologie : fusion sur TOUTE la largeur A..N (14 colonnes)
      (function(tRow, tName) {
        withRetryLocal_(function() {
          var titreRange = sh.getRange(tRow, 1, 1, NUM_ALL_COLS);
          titreRange.merge();
          titreRange.setValue(tName);
          titreRange.setFontWeight('bold');
          titreRange.setBackground('#CFE2F3');
          titreRange.setFontColor('#1A3A5C');
          titreRange.setHorizontalAlignment('center');
          titreRange.setVerticalAlignment('middle');
          sh.setRowHeight(tRow, 75);
        }, 'InvParSite titre row ' + tRow);
      })(titreRow, typoName);

      currentRow++;

      var values     = new Array(typoEquips.length);
      var formulasMN = new Array(typoEquips.length);

      for (var r = 0; r < typoEquips.length; r++) {
        values[r]     = buildRowValues_(typoEquips[r]);
        formulasMN[r] = buildRowFormulas_(typoEquips[r]);
      }

      (function(fRow, equips, vals, fMN) {
        // ✅ Texte A..L
        withRetryLocal_(function() {
          sh.getRange(fRow, 1, equips.length, NUM_TEXT_COLS).setValues(vals);
        }, 'InvParSite equips row ' + fRow);
        // ✅ Images M..N
        withRetryLocal_(function() {
          sh.getRange(fRow, IMG_START_COL, equips.length, 2).setFormulas(fMN);
        }, 'InvParSite photos row ' + fRow);
        // ✅ Noir sur fond blanc sur A..N
        sh.getRange(fRow, 1, equips.length, NUM_ALL_COLS)
          .setFontColor('#000000')
          .setBackground('#FFFFFF');
      })(currentRow, typoEquips, values, formulasMN);

      (function(fRow, lRow) {
        withRetryLocal_(function() {
          sh.getRange(fRow, 1, lRow - fRow + 1, 1).shiftRowGroupDepth(1);
        }, 'InvParSite group ' + fRow + '-' + lRow);
        withRetryLocal_(function() {
          try {
            var grp = sh.getRowGroup(fRow, 1);
            if (grp) grp.expand();
          } catch(e) {
            Logger.log('Expand warn row ' + fRow + ': ' + e);
          }
        }, 'InvParSite expand ' + fRow);
      })(firstEquipRow, lastEquipRow);

      currentRow += typoEquips.length;
    }

    SpreadsheetApp.flush();
  }

  SpreadsheetApp.flush();
}



function writeCompteursGeneral_(sh, codeClient, sitesFinalises, compteurAgg, finalizedCodeSiteSet) {
  var startRow = 15;
  var last = sh.getLastRow();
  if (last >= startRow) sh.getRange(startRow, 2, last - startRow + 1, 8).clearContent();
  if (!sitesFinalises || sitesFinalises.length === 0) return;

  var neededLastRow = startRow - 1 + sitesFinalises.length;
  var maxRows = sh.getMaxRows();
  if (maxRows < neededLastRow) {
    sh.insertRowsAfter(maxRows, neededLastRow - maxRows);
    for (var r = maxRows + 1; r <= neededLastRow; r++) safeMerge_(sh, r, 3, 5);
  }

  var colB = [], colC = [], colH = [], colI = [];
  for (var i = 0; i < sitesFinalises.length; i++) {
    var codeSite = String(sitesFinalises[i].codeSite || '').trim();
    var nomSite = sitesFinalises[i].nomSite;

    var isFinal = finalizedCodeSiteSet && finalizedCodeSiteSet[codeSite];
    var agg = { totalCompteurs: 0, horsService: 0 };
    if (isFinal) {
      var key = String(codeClient).trim() + '||' + codeSite;
      agg = compteurAgg[key] || agg;
    }

    colB.push([i + 1]);
    colC.push([nomSite || '']);
    colH.push([agg.totalCompteurs]);
    colI.push([agg.horsService]);
  }

  sh.getRange(startRow, 2, sitesFinalises.length, 1).setValues(colB);
  sh.getRange(startRow, 3, sitesFinalises.length, 1).setValues(colC);
  sh.getRange(startRow, 8, sitesFinalises.length, 1).setValues(colH);
  sh.getRange(startRow, 9, sitesFinalises.length, 1).setValues(colI);
}

function writeCompteursDetail_(reportSS, codeClient, compteurRowsByClient, finalizedCodeSiteSet) {
  var sh = getSheetByNameLoose_(reportSS, '4. Inventaire des compteurs');
  if (!sh) return;

  var cKey    = String(codeClient).trim();
  var rowsAll = compteurRowsByClient[cKey] || [];

  var rows = [];
  for (var i = 0; i < rowsAll.length; i++) {
    var cs = String(rowsAll[i].codeSite || '').trim();
    if (!finalizedCodeSiteSet || finalizedCodeSiteSet[cs]) rows.push(rowsAll[i]);
  }

  var startRow = 17;
  var last     = sh.getLastRow();
  if (last >= startRow) {
    var maxCol      = sh.getMaxColumns();
    var colsToClean = Math.min(15, maxCol);
    sh.getRange(startRow, 1, last - startRow + 1, colsToClean).clearContent();
  }

  if (rows.length === 0) return;

  var neededLastRow = startRow - 1 + rows.length;
  var maxRows       = sh.getMaxRows();
  if (maxRows < neededLastRow) sh.insertRowsAfter(maxRows, neededLastRow - maxRows);

  // ── Données texte A..I (1 seul appel) ──
  var values = [];
  for (var r = 0; r < rows.length; r++) {
    values.push([
      r + 1,
      rows[r].nomSite       || '',
      rows[r].nomEquip      || '',
      rows[r].local         || '',
      rows[r].typeCompteurS || '',
      rows[r].indexCompteur || '',
      rows[r].unite         || '',
      rows[r].dateReleve    || '',
      rows[r].pdl           || ''
    ]);
  }
  sh.getRange(startRow, 1, rows.length, 9).setValues(values);

  // ── Colonne L = etat (1 seul appel) ──
  sh.getRange(startRow, 12, rows.length, 1)
    .setValues(rows.map(function(rr) { return [rr.etatEquip || '']; }));

  // ✅ OPTIMISATION : Images J et K via =IMAGE() en UN SEUL setFormulas
  var formulasJK = rows.map(function(rr) {
    var p1 = String(rr.photo1 || '').trim();
    var p2 = String(rr.photo2 || '').trim();
    return [
      p1 ? '=IMAGE("' + p1 + '")' : '',
      p2 ? '=IMAGE("' + p2 + '")' : ''
    ];
  });
  sh.getRange(startRow, 10, rows.length, 2).setFormulas(formulasJK);
}


function writeTravauxGeneral_(sh, codeClient, sitesFinalises, travauxAgg, finalizedCodeSiteSet) {
  var startRow = 15;
  var last = sh.getLastRow();

  if (last >= startRow) sh.getRange(startRow, 2, last - startRow + 1, 10).clearContent();
  if (!sitesFinalises || sitesFinalises.length === 0) return;

  var neededLastRow = startRow - 1 + sitesFinalises.length;
  var maxRows = sh.getMaxRows();
  if (maxRows < neededLastRow) {
    sh.insertRowsAfter(maxRows, neededLastRow - maxRows);
    for (var r = maxRows + 1; r <= neededLastRow; r++) safeMerge_(sh, r, 3, 3);
  }

  var colB = [], colC = [], colF = [], colG = [], colH = [], colI = [], colJ = [], colK = [];

  for (var i = 0; i < sitesFinalises.length; i++) {
    var codeSite = String(sitesFinalises[i].codeSite || '').trim();
    var nomSite = sitesFinalises[i].nomSite;

    var isFinal = finalizedCodeSiteSet && finalizedCodeSiteSet[codeSite];
    var agg = { nomSite: nomSite, total: 0, reg: 0, prevAcc: 0, amelio: 0, remise: 0, montant: 0 };

    if (isFinal) {
      var key = String(codeClient).trim() + '||' + codeSite;
      agg = travauxAgg[key] || agg;
      if (!agg.nomSite) agg.nomSite = nomSite;
    }

    colB.push([i + 1]);
    colC.push([agg.nomSite || nomSite || '']);
    colF.push([agg.total]);
    colG.push([agg.reg]);
    colH.push([agg.prevAcc]);
    colI.push([agg.amelio]);
    colJ.push([agg.remise]);
    colK.push([agg.montant]);
  }

  sh.getRange(startRow, 2, sitesFinalises.length, 1).setValues(colB);
  sh.getRange(startRow, 3, sitesFinalises.length, 1).setValues(colC);
  sh.getRange(startRow, 6, sitesFinalises.length, 1).setValues(colF);
  sh.getRange(startRow, 7, sitesFinalises.length, 1).setValues(colG);
  sh.getRange(startRow, 8, sitesFinalises.length, 1).setValues(colH);
  sh.getRange(startRow, 9, sitesFinalises.length, 1).setValues(colI);
  sh.getRange(startRow, 10, sitesFinalises.length, 1).setValues(colJ);
  sh.getRange(startRow, 11, sitesFinalises.length, 1).setValues(colK);
}

function writeTravauxDetail_(reportSS, codeClient, travauxRowsByClient, finalizedCodeSiteSet) {
  var sh = getSheetByNameLoose_(reportSS, "5. Proposition de travaux d'amélioration");
  if (!sh) {
    Logger.log('Feuille "5. Proposition de travaux d\'amélioration" introuvable');
    return;
  }

  var rowsAll = travauxRowsByClient[String(codeClient).trim()] || [];
  var rows    = [];
  for (var i = 0; i < rowsAll.length; i++) {
    if (!finalizedCodeSiteSet || finalizedCodeSiteSet[rowsAll[i].codeSite]) rows.push(rowsAll[i]);
  }

  var startRow = 12;
  var last     = sh.getLastRow();
  if (last >= startRow) {
    var maxCol      = sh.getMaxColumns();
    var colsToClean = Math.min(12, maxCol - 1);
    sh.getRange(startRow, 2, last - startRow + 1, colsToClean).clearContent();
  }

  if (rows.length === 0) return;

  var neededLastRow = startRow - 1 + rows.length;
  var maxRows       = sh.getMaxRows();
  if (maxRows < neededLastRow) sh.insertRowsAfter(maxRows, neededLastRow - maxRows);

  // ── Données texte B..F et I (batch) ──
  var outBE = [], outF = [], outI = [];
  for (var r = 0; r < rows.length; r++) {
    outBE.push([r + 1, rows[r].nomSite || '', rows[r].equipement || '', rows[r].motif || '']);
    outF.push([rows[r].designationTravaux || '']);
    outI.push([rows[r].montant || '']);
  }
  sh.getRange(startRow, 2, rows.length, 4).setValues(outBE);
  sh.getRange(startRow, 6, rows.length, 1).setValues(outF);
  sh.getRange(startRow, 9, rows.length, 1).setValues(outI);

  // ✅ OPTIMISATION : Images G et H via =IMAGE() en UN SEUL setFormulas
  var formulasGH = rows.map(function(rr) {
    var pG = String(rr.imageUrlK || '').trim();
    var pH = String(rr.imageUrlL || '').trim();
    return [
      pG ? '=IMAGE("' + pG + '")' : '',
      pH ? '=IMAGE("' + pH + '")' : ''
    ];
  });
  sh.getRange(startRow, 7, rows.length, 2).setFormulas(formulasGH);
}


/****************************************************
 * ═══════════════════════════════════════════════
 * FONCTIONS HELPERS DIVERS
 * ═══════════════════════════════════════════════
 ****************************************************/

function toNumberRobust_(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;

  var s = String(v)
    .replace(/\u00A0|\u202F/g, ' ')
    .replace(/\s+/g, '')
    .replace(/'/g, '')
    .replace(/,/g, '.');

  var parts = s.split('.');
  if (parts.length > 2) {
    var last = parts.pop();
    s = parts.join('') + '.' + last;
  }

  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function getEligibleSitesFromPerimetre_(reportSS) {
  var sh = getSheetByNameLoose_(reportSS, '1.Périmètre de la prise en charge');
  if (!sh) return [];

  var startRow = 11;
  var last = sh.getLastRow();
  if (last < startRow) return [];

  var n = last - startRow + 1;
  var vals = sh.getRange(startRow, 1, n, 9).getValues();

  function norm_(v) {
    return String(v || '').replace(/[\u00A0\u202F]/g, ' ').replace(/\s+/g, ' ').trim();
  }

  var out = [];
  var seen = {};

  for (var i = 0; i < vals.length; i++) {
    var codeSite = norm_(vals[i][0]);
    var nomSite  = norm_(vals[i][2]);
    var statut   = norm_(vals[i][8]);
    if (!codeSite || !nomSite) continue;

    var ok = (statut === 'Prise en charge finalisée' ||
              statut === 'Prise en charge non débutée' ||
              statut === 'Prise en charge pas réalisée');

    if (!ok) continue;

    if (seen[codeSite]) continue;
    seen[codeSite] = true;

    out.push({ codeSite: codeSite, nomSite: nomSite });
  }

  out.sort(function(a,b){ return (a.nomSite||'').localeCompare(b.nomSite||''); });
  return out;
}

function findFileByExactName_(folder, exactName) {
  var targets = {};
  [exactName, exactName + '.xlsx', exactName + '.pdf', exactName + '.docx'].forEach(function(n) {
    targets[String(n)] = true;
  });

  for (var attempt = 1; attempt <= 3; attempt++) {
    try {
      var it = folder.getFiles();
      while (it.hasNext()) {
        var f = it.next();
        if (targets[f.getName()]) return f;
      }
      return null;
    } catch (e) {
      Utilities.sleep(500 * attempt);
      if (attempt === 3) {
        Logger.log('Drive error getFiles() in folder ' + folder.getId() + ' : ' + e);
        throw e;
      }
    }
  }

  return null;
}

function safeMerge_(sh, row, col, numCols) {
  try {
    sh.getRange(row, col, 1, numCols).merge();
  } catch (e) {}
}

function getSheetByNameLoose_(ss, expectedName) {
  var norm = function(s) {
    return String(s || '')
      .replace(/[\u00A0\u202F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  };

  var target = norm(expectedName);
  var sheets = ss.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    if (norm(sheets[i].getName()) === target) return sheets[i];
  }

  return null;
}

function countEquipEtatForClient_(dataEquip, codeClient) {
  var EQ_CODECLIENT = 1;
  var EQ_ETAT = 41;

  var cKey = String(codeClient).trim();
  var out = { bon: 0, moyen: 0, mauvais: 0, hsProv: 0, hsDef: 0 };

  for (var i = 1; i < dataEquip.length; i++) {
    if (String(dataEquip[i][EQ_CODECLIENT] || '').trim() !== cKey) continue;

    var etat = String(dataEquip[i][EQ_ETAT] || '').trim();
    if (etat === 'En service - Bon état') out.bon++;
    else if (etat === 'En service - État moyen') out.moyen++;
    else if (etat === 'En service - Mauvais état') out.mauvais++;
    else if (etat === 'Hors service - Arrêt provisoire') out.hsProv++;
    else if (etat === 'Hors service - Arrêt définitif') out.hsDef++;
  }

  return out;
}

function sumQteByPresenceForClient_(dataEquip, codeClient) {
  var EQ_CODECLIENT = 1;
  var EQ_PRESENCE   = 5;
  var EQ_QTE        = 8;

  var cKey = String(codeClient).trim();

  var out = {
    present: 0,
    aAjouter: 0,
    aSupprimer: 0,
    sansMention: 0
  };

  function norm_(v) {
    return String(v || '')
      .replace(/[\u00A0\u202F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  for (var i = 1; i < dataEquip.length; i++) {
    if (String(dataEquip[i][EQ_CODECLIENT] || '').trim() !== cKey) continue;

    var presence = norm_(dataEquip[i][EQ_PRESENCE]);
    var qte = toNumberRobust_(dataEquip[i][EQ_QTE]);

    if (presence === 'Équipement présent') out.present += qte;
    else if (presence === 'À ajouter') out.aAjouter += qte;
    else if (presence === 'À supprimer (introuvable)') out.aSupprimer += qte;
    else out.sansMention += qte;
  }

  return out;
}

/****************************************************
 * ═══════════════════════════════════════════════
 * CHECKPOINT & PEC STATE
 * ═══════════════════════════════════════════════
 ****************************************************/

function saveCheckpoint_(r, clientKey, step) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('exportRC_all_lastRow', String(r));
  props.setProperty('exportRC_all_currentKey', String(clientKey || ''));
  props.setProperty('exportRC_all_step', String(step || ''));
}

function loadCheckpoint_() {
  var props = PropertiesService.getScriptProperties();
  return {
    lastRow: parseInt(props.getProperty('exportRC_all_lastRow') || '1', 10),
    currentKey: props.getProperty('exportRC_all_currentKey') || '',
    step: props.getProperty('exportRC_all_step') || ''
  };
}

function clearCheckpoint_() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('exportRC_all_lastRow');
  props.deleteProperty('exportRC_all_currentKey');
  props.deleteProperty('exportRC_all_step');
  // ✅ AJOUT : nettoyage des propriétés anti-boucle
  props.deleteProperty('exportRC_lastRowSeen');
  props.deleteProperty('exportRC_stuckCount');
}


function resetExportRapportClientAllProgress() {
  clearCheckpoint_();
  Logger.log('Progression All-in-one réinitialisée');
}

function getOrCreatePecStateSheet_(ssBase) {
  var sh = ssBase.getSheetByName(CONFIG.sheets.pecState);
  if (!sh) {
    sh = ssBase.insertSheet(CONFIG.sheets.pecState);
    sh.getRange(1, 1, 1, 4).setValues([['codeClient', 'codeSite', 'etat', 'updatedAt']]);
  }
  return sh;
}
function resetPecState() {
  var ssBase = SpreadsheetApp.openById(CONFIG.spreadsheets.base);
  var sh = ssBase.getSheetByName(CONFIG.sheets.pecState);
  if (sh) {
    sh.clearContents();
    sh.getRange(1, 1, 1, 4).setValues([['codeClient', 'codeSite', 'etat', 'updatedAt']]);
  }
  Logger.log('✅ PEC_State réinitialisé');
}

function loadPecStateMap_(sh) {
  var values = sh.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var codeClient = values[i][0];
    var codeSite = values[i][1];
    var etat = values[i][2];
    if (!codeClient || !codeSite) continue;
    map[String(codeClient).trim() + '||' + String(codeSite).trim()] = String(etat || '').trim();
  }
  return map;
}
function computeAndSaveDelta_() {
  var ssBase = SpreadsheetApp.openById(CONFIG.spreadsheets.base);
  var shPEC = ssBase.getSheetByName(CONFIG.sheets.pec);
  var dataPEC = shPEC.getDataRange().getValues();

  var shPecState = getOrCreatePecStateSheet_(ssBase);
  var oldStateMap = loadPecStateMap_(shPecState);
  var delta = buildTodoClientsFromDelta_(dataPEC, oldStateMap);

  var todoKeys = delta.todo.map(function(x) { 
    return String(x.codeClient).trim(); 
  });

  PropertiesService.getScriptProperties().setProperty(
    'delta_todoClients', 
    JSON.stringify(todoKeys)
  );

  // ✅ AJOUT : sauvegarde le snapshot maintenant (au lieu d'attendre la fin du Script 2)
  savePecStateMap_(shPecState, delta.newStateMap);

  Logger.log('DELTA calculé : ' + todoKeys.length + ' client(s) impacté(s)');
  Logger.log('Clients : ' + todoKeys.join(', '));

  return todoKeys;
}


function buildTodoClientsFromDelta_(dataPEC, oldStateMap) {
  var todoByClient = {};
  var newStateMap = {};

  var DEBUG_SITE = 'SIT100453'; // ⚠️ remplace par le code du site que tu testes

  for (var i = 1; i < dataPEC.length; i++) {
    var nomClient = dataPEC[i][0];
    var codeClient = dataPEC[i][1];
    var codeSite = dataPEC[i][2];
    var etat = dataPEC[i][4];

    if (!codeClient || !codeSite) continue;

    var c = String(codeClient).trim();
    var s = String(codeSite).trim();
    var e = String(etat || '').trim();
    var key = c + '||' + s;

    newStateMap[key] = e;

    var old = oldStateMap[key] || '';

    // 🔎 DEBUG ciblé (supprime après diagnostic)
    if (s === DEBUG_SITE) {
      Logger.log('🔎 Site ' + s + ' | OLD="' + old + '" | NEW="' + e + '" | Changement=' + (old !== e));
    }

    if (old !== e) {
      todoByClient[c] = { nomClient: nomClient, codeClient: c };
    }
  }

  var todo = Object.keys(todoByClient).map(function(c) { return todoByClient[c]; });
  return { todo: todo, newStateMap: newStateMap };
}



function savePecStateMap_(sh, newStateMap) {
  var keys = Object.keys(newStateMap);
  var out = new Array(keys.length);
  var now = new Date();

  for (var i = 0; i < keys.length; i++) {
    var parts = keys[i].split('||');
    out[i] = [parts[0], parts[1], newStateMap[keys[i]], now];
  }

  sh.clearContents();
  sh.getRange(1, 1, 1, 4).setValues([['codeClient', 'codeSite', 'etat', 'updatedAt']]);
  if (out.length) sh.getRange(2, 1, out.length, 4).setValues(out);
}

function scheduleResume_() {
  // Supprimer les anciens triggers de reprise (évite les doublons)
  var triggers = ScriptApp.getProjectTriggers();
  var deleted = 0;
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === 'exportRapportClient_AllInOne') {
      ScriptApp.deleteTrigger(triggers[t]);
      deleted++;
    }
  }

  // Programmer la reprise dans 1 minute
  ScriptApp.newTrigger('exportRapportClient_AllInOne')
    .timeBased()
    .after(1 * 60 * 1000)
    .create();

  var ck = loadCheckpoint_();
  logStructured_('INFO', 'SCRIPT2', '⏰ Reprise programmée', {
    dans: '1min',
    reprendAuRow: ck.lastRow,
    anciensTriggersSupprimes: deleted
  });
  Logger.log('✅ Reprise programmée dans 1 minute (reprise au row ' + ck.lastRow + ')');
}


/****************************************************
 * ═══════════════════════════════════════════════
 * INDEX PÉRIMÈTRE & PURGE
 * ═══════════════════════════════════════════════
 ****************************************************/

function buildPerimetreIndexFromSheet_(reportSS) {
  var result = {
    nonDebuteeCodeSet: {},
    codeByName:        {},
    nameByCode:        {}
  };

  var sh = getSheetByNameLoose_(reportSS, '1.Périmètre de la prise en charge');
  if (!sh) return result;

  var startRow = 11;
  var last = sh.getLastRow();
  if (last < startRow) return result;

  var n = last - startRow + 1;
  var vals = sh.getRange(startRow, 1, n, 9).getValues();

  function norm_(v) {
    return String(v || '')
      .replace(/[\u00A0\u202F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  for (var i = 0; i < vals.length; i++) {
    var codeSite = norm_(vals[i][0]);
    var nomSite  = norm_(vals[i][2]);
    var statut   = norm_(vals[i][8]);

    if (!codeSite) continue;

    if (nomSite) {
      if (!result.codeByName[nomSite]) result.codeByName[nomSite] = codeSite;
      result.nameByCode[codeSite] = nomSite;
    }

    if (statut === 'Prise en charge non débutée') {
      result.nonDebuteeCodeSet[codeSite] = true;
    }
  }

  return result;
}

function getNonDebuteeSiteNameSetFromPerimetre_(reportSS) {
  return buildPerimetreIndexFromSheet_(reportSS).nonDebuteeCodeSet;
}

function backfillReportFileIdInClientsIndex() {
  var ssBase = SpreadsheetApp.openById(CONFIG.spreadsheets.base);
  var sh = ssBase.getSheetByName(CONFIG.sheets.clientsIndex);
  var values = sh.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) {
    var codeClient = values[i][0];
    var agence = values[i][1];
    var folderId = values[i][2];
    var reportFileId = values[i][4];

    if (!codeClient || !agence || !folderId) continue;
    if (reportFileId) continue;

    var folder;
    try { folder = DriveApp.getFolderById(folderId); }
    catch(e){ Logger.log('Folder invalide ligne ' + (i+1)); continue; }

    var it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    if (!it.hasNext()) { Logger.log('Aucun Google Sheet dans folder ' + folderId); continue; }

    var f = it.next();
    sh.getRange(i+1, 5).setValue(f.getId());
    Logger.log('OK ' + agence + '|' + codeClient + ' => ' + f.getName());
  }
}
