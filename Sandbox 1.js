/****************************************************
 * CACHE LAZY (en mémoire, valable uniquement pendant l’exécution)
 ****************************************************/
var DRIVE_CACHE = {
  agences: {
    // agenceNom: { folder: Folder, clientsByName: { nomClient: Folder } }
  },
  clientFiles: {
    // folderId: { fileName: true }
  }
};

function getAgenceCache_(agenceNom, agenceFolderId) {
  if (DRIVE_CACHE.agences[agenceNom]) return DRIVE_CACHE.agences[agenceNom];

  var agenceFolder = DriveApp.getFolderById(agenceFolderId);
  var clientsByName = {};

  // On indexe les dossiers clients de l’agence une seule fois, quand on en a besoin
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
  // On indexe les fichiers du client une seule fois, quand on en a besoin
  var it = clientFolder.getFiles();
  while (it.hasNext()) {
    var file = it.next();
    filesIndex[file.getName()] = true;
  }

  DRIVE_CACHE.clientFiles[id] = filesIndex;
  return filesIndex;
}

function markFileInClientIndex_(clientFolder, fileName) {
  var idx = getClientFilesIndex_(clientFolder);
  idx[fileName] = true;
}

/****************************************************
 * SCRIPT PRINCIPAL (base identique, mais sans pré-scan global Drive)
 ****************************************************/
function createFoldersAndSiteFiles() {
  var ss = SpreadsheetApp.openById('1mIfo-H14tmW_Jz6EFJPEdcFiOVXzhzAkf5JgUCkeVqI');
  var sheet = ss.getSheetByName('Tableetatpriseencharge');
  var data = sheet.getDataRange().getValues();
  var sheetSites = ss.getSheetByName('Sites');
  var dataSites = sheetSites.getDataRange().getValues();

  var agences = {
    'PRO': '1dEDZOLDPNjj17K1wUlyJ1VW4pkCDQR-D',
    'ALM': '1TbSChiIKf6aDVOB5KXwaIUe1xGfE62Gn',
    'NAQ': '1J5QqdZJv4XpdBy3rl71ge2DlBO2_O72y',
    'VAR': '1nYlC01I_palN9xWWjIasSs2nPhFDluGq',
    'AUR': '1ABo0FjOybixCbPFNt076qGT47pwbUdkr',
    'OCC': '1lRlFB9BlmcUm7UR98Pc-VlcH8HLcQZxy',
    'LGR': '19M36fkpNm-kYQRsRODTBDr34zTYO5vag',
    'IDFBS': '13IoNu6gPnNPqF81o3Pubc1pbeX2flN_j',
    'IDFCO': '13IoNu6gPnNPqF81o3Pubc1pbeX2flN_j',
    'IDFCP': '13IoNu6gPnNPqF81o3Pubc1pbeX2flN_j',
    'IDFTE': '13IoNu6gPnNPqF81o3Pubc1pbeX2flN_j'
  };

  var modeleId = '1676EJY-Yv5wQMsbWQPc8OYFZpxq13A_Rg6CBg71lAak';

  // 1) Créer un index des codes entités (inchangé)
  var indexCodesEntites = {};
  for (var k = 1; k < dataSites.length; k++) {
    indexCodesEntites[String(dataSites[k][2])] = dataSites[k][5];
  }

  // 2) Traitement (optimisé via cache lazy)
  var startTime = new Date().getTime();
  var maxExecutionTime = 5 * 60 * 1000; // 5 minutes
  var sitesTraites = 0;

  for (var i = 1; i < data.length; i++) {
    // Vérifier le temps d'exécution
    if (i % 10 === 0) {
      var currentTime = new Date().getTime();
      if (currentTime - startTime > maxExecutionTime) {
        Logger.log('Timeout atteint. Sites traités: ' + sitesTraites + '/' + (data.length - 1));
        break;
      }
    }

    var nomClient = data[i][0];
    var codeClient = data[i][1];
    var codeSite = data[i][2];
    var nomSite = data[i][3];
    var etatSite = data[i][4];

    if (etatSite !== 'Prise en charge finalisée') continue;
    if (!nomClient || !codeClient || !codeSite || !nomSite) continue;

    nomClient = String(nomClient).replace(/[/:?"<>|\[\]]/g, '').trim();
    nomSite = String(nomSite).replace(/[/:?"<>|\[\]]/g, '').trim();

    // Utiliser l'index
    var codeEntite = indexCodesEntites[String(codeClient)];
    if (!codeEntite) {
      Logger.log('Pas de code_entite pour codeClient: ' + codeClient);
      continue;
    }

    // Trouver l’agence
    var agenceNom = '';
    for (var cle in agences) {
      if (String(codeEntite).indexOf(cle) !== -1) {
        agenceNom = cle;
        break;
      }
    }
    if (!agenceNom) {
      Logger.log('Pas d\'agence pour codeEntite: ' + codeEntite);
      continue;
    }

    // === Lazy: charger l’agence seulement quand nécessaire
    var agenceCache = getAgenceCache_(agenceNom, agences[agenceNom]);

    // === Lazy: récupérer/créer le dossier client
    var clientFolder = agenceCache.clientsByName[nomClient];
    if (!clientFolder) {
      clientFolder = getOrCreateClientFolder_(agenceCache, nomClient);
      Logger.log('Nouveau dossier client créé: ' + nomClient);
    }

    // === Lazy: indexer les fichiers du client seulement quand nécessaire
    var fichiersIndex = getClientFilesIndex_(clientFolder);

    // Vérifier si le fichier existe déjà
    var formatsAVerifier = [nomSite + '.xlsx', nomSite + '.docx', nomSite + '.pdf', nomSite];
    var fichierExiste = false;
    for (var f = 0; f < formatsAVerifier.length; f++) {
      if (fichiersIndex[formatsAVerifier[f]]) {
        fichierExiste = true;
        Logger.log('Fichier existant détecté: ' + formatsAVerifier[f] + ' - CRÉATION ANNULÉE');
        break;
      }
    }
    if (fichierExiste) continue;

    // Créer le fichier site
    try {
      var modeleFile = DriveApp.getFileById(modeleId);
      modeleFile.makeCopy(nomSite, clientFolder);

      // Mettre à jour le cache
      markFileInClientIndex_(clientFolder, nomSite);

      Logger.log('NOUVEAU FICHIER CRÉÉ: ' + nomSite + ' pour client: ' + nomClient);
      sitesTraites++;

    } catch (error) {
      Logger.log('Erreur lors de la création du fichier ' + nomSite + ': ' + error.toString());
    }
  }

  Logger.log('Script terminé. Total sites traités: ' + sitesTraites);
}
