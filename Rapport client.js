/****************************************************
 * RAPPORT CLIENT (1 fichier par client)
 * VERSION OPTIMISÉE avec P0 + P1 + CORRECTIFS
 *
 * Onglet "1.Périmètre de la prise en charge" (à partir de ligne 11):
 *   A = codeSite (police blanche)
 *   B = numéro
 *   C:G (fusion) = nomSite
 *   H = adresse = (J) + ", " + (M) + " " + (N) depuis la feuille Sites (via codeSite Sites!A)
 *   I = état (PEC!E)
 *   J = mail / réalisateur (PEC!J)
 *
 * CRÉATION :
 * - Crée 1 copie du template par client (finalisé OU non débuté)
 * - Remplit Page de garde (B43 nom client, B50 dernière date finalisée)
 *
 * MISE À JOUR :
 * - Réécrit complètement l'onglet périmètre avec :
 *    * Prise en charge finalisée
 *    * Prise en charge non débutée
 * - Met à jour H/I/J
 *
 * CORRECTIFS APPLIQUÉS :
 * ✅ CONFIG unique et complet (avec pecState + templates)
 * ✅ logStructured_ sans création de feuille (Logger.log uniquement)
 ****************************************************/

/****************************************************
 * ═══════════════════════════════════════════════
 * CONFIGURATION CENTRALISÉE (UNIQUE DANS LE PROJET)
 * ═══════════════════════════════════════════════
 ****************************************************/
var CONFIG = {
  spreadsheets: {
    base: '1mIfo-H14tmW_Jz6EFJPEdcFiOVXzhzAkf5JgUCkeVqI',
    equipement: '1YPEO7-jRw5adU9OlBmeVH9DEZ6uVvYPmSJn9ZICa4pw'
  },
  templates: {
    rapportClient: '1wgd6eeTrLSPRbNwZqtZC9CYm1GiWUEpSQ-XxkcClEbQ'
  },
  agences: {
    'PRO':    '1dEDZOLDPNjj17K1wUlyJ1VW4pkCDQR-D',
    'ALM':    '1TbSChiIKf6aDVOB5KXwaIUe1xGfE62Gn',
    'NAQ':    '1J5QqdZJv4XpdBy3rl71ge2DlBO2_O72y',
    'VAR':    '1nYlC01I_palN9xWWjIasSs2nPhFDluGq',
    'AUR':    '1ABo0FjOybixCbPFNt076qGT47pwbUdkr',
    'OCC':    '1lRlFB9BlmcUm7UR98Pc-VlcH8HLcQZxy',
    'LGR':    '19M36fkpNm-kYQRsRODTBDr34zTYO5vag',
    'IDFBS':  '13IoNu6gPnNPqF81o3Pubc1pbeX2flN_j',
    'IDFCO':  '13IoNu6gPnNPqF81o3Pubc1pbeX2flN_j',
    'IDFCP':  '13IoNu6gPnNPqF81o3Pubc1pbeX2flN_j',
    'IDFTE':  '13IoNu6gPnNPqF81o3Pubc1pbeX2flN_j',
    'GESTEN': '13IoNu6gPnNPqF81o3Pubc1pbeX2flN_j'
  },
  sheets: {
    pec: 'Tableetatpriseencharge',
    sites: 'Sites',
    clientsIndex: 'ClientsIndex',
    pecState: 'PEC_State'  // ✅ AJOUTÉ — corrige les feuilles parasites + DELTA
  }
};

/****************************************************
 * ═══════════════════════════════════════════════
 * LOGS STRUCTURÉS (sans création de feuille)
 * ═══════════════════════════════════════════════
 ****************************************************/
function logStructured_(level, category, message, data) {
  // ✅ Écrit uniquement dans les logs Apps Script (Exécutions > Afficher les journaux)
  var logLine = '[' + level + '][' + category + '] ' + message;
  if (data) {
    logLine += ' | ' + JSON.stringify(data);
  }
  Logger.log(logLine);
}

/****************************************************
 * ═══════════════════════════════════════════════
 * FONCTION PRINCIPALE
 * ═══════════════════════════════════════════════
 ****************************************************/
function myFunction() {
  createRapportClientFiles();
}

/***************
 * CACHE LAZY
 ***************/
var DRIVE_CACHE = {
  agences: {},
  clientFiles: {}
};

function sanitizeName_(s) {
  return String(s || '').replace(/[/:?"<>|\[\]]/g, '').trim();
}

/****************************************************
 * DRIVE helpers
 ****************************************************/
function getAgenceCache_(agenceNom, agenceFolderId) {
  if (DRIVE_CACHE.agences[agenceNom]) return DRIVE_CACHE.agences[agenceNom];

  var agenceFolder = DriveApp.getFolderById(agenceFolderId);
  var clientsByName = {};

  var it = agenceFolder.getFolders();
  while (it.hasNext()) {
    var f = it.next();
    clientsByName[f.getName()] = f;
  }

  DRIVE_CACHE.agences[agenceNom] = { folder: agenceFolder, clientsByName: clientsByName };
  return DRIVE_CACHE.agences[agenceNom];
}

function getOrCreateClientFolder_(agenceCache, nomClient) {
  var existing = agenceCache.clientsByName[nomClient];
  if (existing) return existing;

  var created = agenceCache.folder.createFolder(nomClient);
  agenceCache.clientsByName[nomClient] = created;
  return created;
}

function getClientFilesIndex_(clientFolder) {
  var id = clientFolder.getId();
  if (DRIVE_CACHE.clientFiles[id]) return DRIVE_CACHE.clientFiles[id];

  var filesIndex = {};
  var it = clientFolder.getFiles();
  while (it.hasNext()) {
    var file = it.next();
    filesIndex[file.getName()] = { exists: true, id: file.getId() };
  }

  DRIVE_CACHE.clientFiles[id] = filesIndex;
  return filesIndex;
}

function markFileInClientIndex_(clientFolder, fileName, fileId) {
  var idx = getClientFilesIndex_(clientFolder);
  idx[fileName] = { exists: true, id: fileId };
}

/****************************************************
 * ClientsIndex helpers
 ****************************************************/
function getOrCreateClientsIndexSheet_(ss) {
  var sh = ss.getSheetByName(CONFIG.sheets.clientsIndex);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.sheets.clientsIndex);
    sh.getRange(1, 1, 1, 5).setValues([
      ["codeClient", "agence", "folderId", "folderName", "reportFileId"]
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function loadClientsIndex_(indexSheet) {
  var values = indexSheet.getDataRange().getValues();
  var map = {};

  for (var i = 1; i < values.length; i++) {
    var row = values[i] || [];
    var codeClient   = row[0];
    var agence       = row[1];
    var folderId     = row[2];
    var folderName   = row[3];
    var reportFileId = row.length > 4 ? row[4] : '';

    if (!codeClient || !agence || !folderId) continue;

    map[String(agence) + '|' + String(codeClient)] = {
      folderId:     String(folderId),
      folderName:   String(folderName   || ''),
      reportFileId: String(reportFileId || '')
    };
  }
  return map;
}

function upsertClientsIndexRow_(indexSheet, agence, codeClient, folder) {
  var data      = indexSheet.getDataRange().getValues();
  var keyAgence = String(agence);
  var keyCode   = String(codeClient);
  var folderId  = folder.getId();
  var folderName = folder.getName();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === keyCode && String(data[i][1]) === keyAgence) {
      indexSheet.getRange(i + 1, 3, 1, 2).setValues([[folderId, folderName]]);
      return;
    }
  }
  indexSheet.appendRow([keyCode, keyAgence, folderId, folderName]);
}

function resolveClientFolder_(agenceCache, clientsIndexMap, indexSheet,
                              agenceNom, codeClient, nomClient) {
  var mapKey = String(agenceNom) + '|' + String(codeClient);
  var entry  = clientsIndexMap[mapKey];

  if (entry && entry.folderId) {
    try {
      var byId = DriveApp.getFolderById(entry.folderId);
      upsertClientsIndexRow_(indexSheet, agenceNom, codeClient, byId);
      return byId;
    } catch (e) {
      logStructured_('WARN', 'SCRIPT1', 'Folder invalide dans index', {
        agence: agenceNom,
        codeClient: codeClient,
        folderId: entry.folderId
      });
    }
  }

  var folder = agenceCache.clientsByName[nomClient];
  if (!folder) folder = getOrCreateClientFolder_(agenceCache, nomClient);

  upsertClientsIndexRow_(indexSheet, agenceNom, codeClient, folder);
  clientsIndexMap[mapKey] = { folderId: folder.getId(), folderName: folder.getName() };

  return folder;
}

/****************************************************
 * INDEXES depuis "Sites"
 ****************************************************/
function buildClientNameByCodeFromSites_(dataSites) {
  var map = {};
  for (var i = 1; i < dataSites.length; i++) {
    var codeClient = dataSites[i][2]; // C
    if (!codeClient) continue;
    var key = String(codeClient).trim();

    var nomClientBB = dataSites[i][53]; // BB
    if (nomClientBB && !map[key]) map[key] = String(nomClientBB).trim();
  }
  return map;
}

function buildAdresseByCodeSiteFromSites_(dataSites) {
  var map = {};
  for (var i = 1; i < dataSites.length; i++) {
    var codeSite = dataSites[i][0]; // A
    if (!codeSite) continue;

    var cs   = String(codeSite).trim();
    var adrJ = String(dataSites[i][9]  || '').trim(); // J
    var adrM = String(dataSites[i][12] || '').trim(); // M
    var adrN = String(dataSites[i][13] || '').trim(); // N
    var adr  = (adrJ + ', ' + adrM + ' ' + adrN).trim();

    if (!map[cs]) map[cs] = adr;
    else if (adr && adr.length > map[cs].length) map[cs] = adr;
  }
  return map;
}

/****************************************************
 * PEC — récupère tous les sites (finalisé + non débuté)
 * et la dernière date uniquement sur les finalisées
 ****************************************************/
function buildClientSitesAndLastDate_(dataPEC) {
  var sitesByClient    = {};
  var lastDateByClient = {};

  var etatsAutorises = {
    'Prise en charge finalisée':   true,
    'Prise en charge non débutée': true
  };

  for (var i = 1; i < dataPEC.length; i++) {
    var codeClient = dataPEC[i][1];
    var codeSite   = dataPEC[i][2];
    var nomSite    = dataPEC[i][3];
    var etat       = String(dataPEC[i][4] || '').trim();
    var dateVal    = dataPEC[i][8];
    var mail       = dataPEC[i][9];

    if (!etatsAutorises[etat]) continue;
    if (!codeClient || !nomSite || !codeSite) continue;

    var key = String(codeClient).trim();
    var cs  = String(codeSite).trim();
    var ns  = String(nomSite).trim();

    if (!sitesByClient[key]) sitesByClient[key] = [];

    var found = false;
    for (var x = 0; x < sitesByClient[key].length; x++) {
      if (sitesByClient[key][x].codeSite === cs) {
        sitesByClient[key][x].nomSite = ns;
        sitesByClient[key][x].etat   = etat;
        sitesByClient[key][x].mail   = String(mail || '').trim();
        found = true;
        break;
      }
    }

    if (!found) {
      sitesByClient[key].push({
        codeSite: cs,
        nomSite:  ns,
        etat:     etat,
        mail:     String(mail || '').trim()
      });
    }

    if (etat === 'Prise en charge finalisée' && dateVal) {
      var d = null;
      if (Object.prototype.toString.call(dateVal) === '[object Date]' &&
          !isNaN(dateVal.getTime())) {
        d = dateVal;
      } else {
        var parsed = new Date(dateVal);
        if (!isNaN(parsed.getTime())) d = parsed;
      }
      if (d) {
        if (!lastDateByClient[key] ||
            d.getTime() > lastDateByClient[key].getTime()) {
          lastDateByClient[key] = d;
        }
      }
    }
  }

  for (var k in sitesByClient) {
    sitesByClient[k].sort(function(a, b) {
      return a.nomSite.localeCompare(b.nomSite);
    });
  }

  return { sitesByClient: sitesByClient, lastDateByClient: lastDateByClient };
}

/****************************************************
 * Index codeSite -> {etat, mail} depuis PEC
 ****************************************************/
function buildInfoBySiteCode_(dataPEC) {
  var map = {};
  for (var i = 1; i < dataPEC.length; i++) {
    var codeSite = dataPEC[i][2];
    if (!codeSite) continue;
    map[String(codeSite).trim()] = {
      etat: String(dataPEC[i][4] || '').trim(),
      mail: String(dataPEC[i][9] || '').trim()
    };
  }
  return map;
}

/****************************************************
 * Écritures dans le fichier rapport
 ****************************************************/
function writeClientNameToReport_(reportSS, clientName) {
  var sh = reportSS.getSheetByName('Page de garde');
  if (!sh) return;
  sh.getRange('B43').setValue(clientName);
}

function writeLastPecDateToReport_(reportSS, lastDate) {
  var sh = reportSS.getSheetByName('Page de garde');
  if (!sh) return;
  if (!lastDate) return;
  sh.getRange('B50').setValue(lastDate);
}

/****************************************************
 * ✅ P0 - Réécrit complètement le périmètre (OPTIMISÉ)
 * Réduit de 5 setValues() à 1 seul
 ****************************************************/
function syncSitesToReport_(reportSS, sites) {
  var sh = reportSS.getSheetByName('1.Périmètre de la prise en charge');
  if (!sh) return;

  var startRow = 11;
  var lastRow  = sh.getLastRow();

  if (lastRow >= startRow) {
    sh.getRange(startRow, 1, lastRow - startRow + 1, 10).clearContent();
  }

  if (!sites || sites.length === 0) return;

  var needed       = sites.length;
  var maxRows      = sh.getMaxRows();
  var neededLastRow = startRow - 1 + needed;

  if (maxRows < neededLastRow) {
    sh.insertRowsAfter(maxRows, neededLastRow - maxRows);
    for (var r = maxRows + 1; r <= neededLastRow; r++) {
      try { sh.getRange(r, 3, 1, 5).merge(); } catch (e) {}
    }
  }

  // ✅ OPTIMISATION : 1 seul setValues() au lieu de 5
  var allData = [];
  for (var i = 0; i < sites.length; i++) {
    allData.push([
      sites[i].codeSite || '',  // A (blanc)
      i + 1,                    // B
      sites[i].nomSite || '',   // C
      '', '', '', '',           // D, E, F, G (vides)
      '',                       // H (adresse, remplie après)
      sites[i].etat || '',      // I
      sites[i].mail || ''       // J
    ]);
  }

  sh.getRange(startRow, 1, needed, 10).setValues(allData);
  sh.getRange(startRow, 1, needed, 1).setFontColor('#FFFFFF'); // A en blanc
}

/****************************************************
 * Update H (adresse) via codeSite
 ****************************************************/
function updateAdresseInReportByCodeSite_(reportSS, adresseByCodeSite) {
  var sh = reportSS.getSheetByName('1.Périmètre de la prise en charge');
  if (!sh) return;

  var startRow = 11;
  var lastRow  = sh.getLastRow();
  if (lastRow < startRow) return;

  var nbRows     = lastRow - startRow + 1;
  var colAValues = sh.getRange(startRow, 1, nbRows, 1).getValues();
  var newColH    = [];

  for (var i = 0; i < colAValues.length; i++) {
    var codeSite = String(colAValues[i][0] || '').trim();
    newColH.push([adresseByCodeSite[codeSite] || '']);
  }

  sh.getRange(startRow, 8, nbRows, 1).setValues(newColH);
}

/****************************************************
 * Update I (état) et J (mail) via codeSite en A
 ****************************************************/
function updateInfoInReport_(reportSS, infoBySiteCode) {
  var sh = reportSS.getSheetByName('1.Périmètre de la prise en charge');
  if (!sh) return;

  var startRow = 11;
  var lastRow  = sh.getLastRow();
  if (lastRow < startRow) return;

  var nbRows     = lastRow - startRow + 1;
  var colAValues = sh.getRange(startRow, 1, nbRows, 1).getValues();
  var newColI    = [];
  var newColJ    = [];

  for (var i = 0; i < colAValues.length; i++) {
    var codeSite = String(colAValues[i][0] || '').trim();
    var info     = codeSite ? infoBySiteCode[codeSite] : null;
    newColI.push([info ? (info.etat || '') : '']);
    newColJ.push([info ? (info.mail || '') : '']);
  }

  sh.getRange(startRow, 9,  nbRows, 1).setValues(newColI);
  sh.getRange(startRow, 10, nbRows, 1).setValues(newColJ);
}

/****************************************************
 * DELTA helper — charge le todoSet depuis ScriptProperties
 * Retourne null si pas de DELTA (= traitement complet)
 ****************************************************/
function loadDeltaTodoSet_() {
  var raw = PropertiesService.getScriptProperties()
              .getProperty('delta_todoClients');
  if (!raw) return null;

  var keys = JSON.parse(raw);
  var set  = {};
  keys.forEach(function(k) { set[k] = true; });
  return set;
}

/****************************************************
 * ✅ SCRIPT PRINCIPAL (avec logs détaillés par client)
 ****************************************************/
function createRapportClientFiles() {
  var startTime = Date.now();
  
  logStructured_('INFO', 'SCRIPT1', '🚀 Démarrage création fichiers rapports', {});
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('🚀 SCRIPT 1 - CRÉATION FICHIERS RAPPORTS');
  Logger.log('═══════════════════════════════════════════════════');
  
  PropertiesService.getScriptProperties().deleteProperty('script1_done');

  var ss = SpreadsheetApp.openById(CONFIG.spreadsheets.base);

  var sheetPEC = ss.getSheetByName(CONFIG.sheets.pec);
  var dataPEC  = sheetPEC.getDataRange().getValues();

  var sheetSites = ss.getSheetByName(CONFIG.sheets.sites);
  var dataSites  = sheetSites.getDataRange().getValues();

  var agences = CONFIG.agences;
  var modeleId = CONFIG.templates.rapportClient;

  // codeClient -> codeEntite
  var indexCodesEntites = {};
  for (var k = 1; k < dataSites.length; k++) {
    var ck = String(dataSites[k][2]).trim();
    var ce = dataSites[k][5];
    if (ck && (!indexCodesEntites[ck] || !String(indexCodesEntites[ck]).trim())) {
      indexCodesEntites[ck] = ce;
    }
  }

  var clientNameByCode      = buildClientNameByCodeFromSites_(dataSites);
  var adresseByCodeSite     = buildAdresseByCodeSiteFromSites_(dataSites);
  var clientSitesData       = buildClientSitesAndLastDate_(dataPEC);
  var clientSitesByCode     = clientSitesData.sitesByClient;
  var lastFinalizedDateByClient = clientSitesData.lastDateByClient;
  var infoBySiteCode        = buildInfoBySiteCode_(dataPEC);

  var indexSheet     = getOrCreateClientsIndexSheet_(ss);
  var clientsIndexMap = loadClientsIndex_(indexSheet);

  var clientsTraites = {};
  var createdCount   = 0;
  var processedCount = 0;
  var errorCount     = 0;
  var skippedCount   = 0;

  // ✅ Charge le DELTA
  var deltaSet = loadDeltaTodoSet_();
  if (deltaSet !== null) {
    var deltaCount = Object.keys(deltaSet).length;
    logStructured_('INFO', 'SCRIPT1', '🎯 DELTA actif', { count: deltaCount });
    Logger.log('🎯 DELTA actif : ' + deltaCount + ' client(s) à traiter');
  } else {
    logStructured_('INFO', 'SCRIPT1', '📋 DELTA inactif : traitement complet', {});
    Logger.log('📋 DELTA inactif : traitement complet');
  }
  Logger.log('');

  for (var i = 1; i < dataPEC.length; i++) {
    var nomClient  = dataPEC[i][0];
    var codeClient = dataPEC[i][1];
    var etatSite   = String(dataPEC[i][4] || '').trim();

    var etatsAutorises = {
      'Prise en charge finalisée':   true,
      'Prise en charge non débutée': true
    };
    if (!etatsAutorises[etatSite]) continue;
    if (!nomClient || !codeClient) continue;

    var codeKey = String(codeClient).trim();

    // ✅ Filtre DELTA
    if (deltaSet !== null && !deltaSet[codeKey]) {
      skippedCount++;
      continue;
    }

    var clientSites = clientSitesByCode[codeKey] || [];
    if (clientSites.length === 0) continue;

    nomClient = sanitizeName_(nomClient);

    var codeEntite = indexCodesEntites[codeKey];
    if (!codeEntite) continue;

    var agenceNom = '';
    for (var cle in agences) {
      if (String(codeEntite).indexOf(cle) !== -1) {
        agenceNom = cle;
        break;
      }
    }
    if (!agenceNom) continue;

    var treatedKey = agenceNom + '|' + codeKey;
    if (clientsTraites[treatedKey]) continue;

    try {
      var agenceCache  = getAgenceCache_(agenceNom, agences[agenceNom]);
      var clientFolder = resolveClientFolder_(agenceCache, clientsIndexMap,
                         indexSheet, agenceNom, codeClient, nomClient);

      var fichiersIndex    = getClientFilesIndex_(clientFolder);
      var fileBaseName     = nomClient;
      var existingFileId   = null;
      var formatsAVerifier = [
        fileBaseName,
        fileBaseName + '.xlsx',
        fileBaseName + '.pdf',
        fileBaseName + '.docx'
      ];

      for (var f = 0; f < formatsAVerifier.length; f++) {
        if (fichiersIndex[formatsAVerifier[f]] &&
            fichiersIndex[formatsAVerifier[f]].id) {
          existingFileId = fichiersIndex[formatsAVerifier[f]].id;
          break;
        }
      }

      var reportFileId = null;
      var isNew        = false;

      if (!existingFileId) {
        var newFile = DriveApp.getFileById(modeleId).makeCopy(fileBaseName, clientFolder);
        reportFileId = newFile.getId();
        markFileInClientIndex_(clientFolder, fileBaseName, reportFileId);
        upsertClientsIndexRowWithReportId_(indexSheet, agenceNom, codeClient,
                                           clientFolder, reportFileId);
        createdCount++;
        isNew = true;
        
        Logger.log('✨ CRÉÉ   | ' + agenceNom + ' | ' + codeKey + ' | ' + nomClient + ' | ' + clientSites.length + ' site(s)');
        logStructured_('INFO', 'SCRIPT1', '✨ Fichier créé', {
          codeClient: codeKey,
          nomClient: nomClient,
          agence: agenceNom,
          nbSites: clientSites.length
        });
      } else {
        reportFileId = existingFileId;
        
        Logger.log('🔄 MAJ    | ' + agenceNom + ' | ' + codeKey + ' | ' + nomClient + ' | ' + clientSites.length + ' site(s)');
        logStructured_('INFO', 'SCRIPT1', '🔄 Fichier mis à jour', {
          codeClient: codeKey,
          nomClient: nomClient,
          agence: agenceNom,
          nbSites: clientSites.length
        });
      }

      var reportSS     = SpreadsheetApp.openById(reportFileId);
      var nameFromSites = clientNameByCode[codeKey];

      writeClientNameToReport_(reportSS, sanitizeName_(nameFromSites || nomClient));

      var lastDate = lastFinalizedDateByClient[codeKey] || null;
      writeLastPecDateToReport_(reportSS, lastDate);

      syncSitesToReport_(reportSS, clientSites);
      updateAdresseInReportByCodeSite_(reportSS, adresseByCodeSite);
      updateInfoInReport_(reportSS, infoBySiteCode);

      processedCount++;
      clientsTraites[treatedKey] = true;
      
    } catch(e) {
      errorCount++;
      
      Logger.log('❌ ERREUR | ' + agenceNom + ' | ' + codeKey + ' | ' + nomClient + ' | ' + e.toString());
      logStructured_('ERROR', 'SCRIPT1', '❌ Erreur traitement client', {
        codeClient: codeKey,
        nomClient: nomClient,
        agence: agenceNom,
        error: e.toString(),
        stack: e.stack || ''
      });
    }
  }

  var duration = ((Date.now() - startTime) / 1000).toFixed(1) + 's';
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('✅ SCRIPT 1 TERMINÉ');
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('📊 Résumé :');
  Logger.log('   • Créés       : ' + createdCount);
  Logger.log('   • Mis à jour  : ' + (processedCount - createdCount));
  Logger.log('   • Total traité: ' + processedCount);
  Logger.log('   • Ignorés     : ' + skippedCount + ' (DELTA)');
  Logger.log('   • Erreurs     : ' + errorCount);
  Logger.log('   • Durée       : ' + duration);
  Logger.log('═══════════════════════════════════════════════════');
  
  logStructured_('INFO', 'SCRIPT1', '✅ Terminé', {
    created: createdCount,
    updated: processedCount - createdCount,
    processed: processedCount,
    skipped: skippedCount,
    errors: errorCount,
    duration: duration
  });
}

/****************************************************
 * Upsert ClientsIndex avec reportFileId (colonne E)
 ****************************************************/
function upsertClientsIndexRowWithReportId_(indexSheet, agence, codeClient,
                                            folder, reportFileId) {
  var data = indexSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(codeClient) &&
        String(data[i][1]) === String(agence)) {
      indexSheet.getRange(i + 1, 3, 1, 3).setValues([
        [folder.getId(), folder.getName(), reportFileId]
      ]);
      return;
    }
  }
  indexSheet.appendRow([String(codeClient), String(agence),
                        folder.getId(), folder.getName(), reportFileId]);
}
