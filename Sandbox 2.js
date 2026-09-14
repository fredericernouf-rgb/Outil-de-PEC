function exportAllByClientAndSite() {
  // ===== PARTIE 1: ATTENTE ET VÉRIFICATION DES FICHIERS =====
  Logger.log('Début du script - Attente de la création des fichiers...');
  
  // Attendre 2-3 minutes que la création des fichiers soit terminée
  Utilities.sleep(1); // 1 minute
  
  // Vérifier si les fichiers existent avant de continuer
  var maxTentatives = 10;
  var tentative = 0;
  var fichiersPresents = false;
  
  while (!fichiersPresents && tentative < maxTentatives) {
    fichiersPresents = verifierPresenceFichiers();
    if (!fichiersPresents) {
      Logger.log('Tentative ' + (tentative + 1) + ': Fichiers pas encore créés, attente...');
      Utilities.sleep(30000); // Attendre 30 secondes
      tentative++;
    }
  }
  
  if (!fichiersPresents) {
    Logger.log('ERREUR: Fichiers toujours pas créés après ' + maxTentatives + ' tentatives');
    return;
  }
  
  Logger.log('Fichiers détectés, début du traitement d\'export...');
  
  // ===== PARTIE 2: INITIALISATION DES DONNÉES =====
  var folderRoot = DriveApp.getFolderById('13gsSiHdPyeaJmdAFivDNpGgKLChjGgED');
  var ssBase = SpreadsheetApp.openById('1mIfo-H14tmW_Jz6EFJPEdcFiOVXzhzAkf5JgUCkeVqI');
  var sheetSites = ssBase.getSheetByName('Tableetatpriseencharge');
  var dataSites = sheetSites.getDataRange().getValues();

  var ssEquip = SpreadsheetApp.openById('1YPEO7-jRw5adU9OlBmeVH9DEZ6uVvYPmSJn9ZICa4pw');
  var sheetEquip = ssEquip.getSheetByName('Equipement');
  var dataEquip = sheetEquip.getDataRange().getValues();
  var sheetPhoto = ssEquip.getSheetByName('Photothèque');
  var dataPhoto = sheetPhoto.getDataRange().getValues();
  var sheetTravaux = ssEquip.getSheetByName('Proposition_de_travaux');
  var dataTravaux = sheetTravaux.getDataRange().getValues();

  // CORRECTION: Données CVC avec le bon ID de fichier
  var ssCVC = SpreadsheetApp.openById('1hXdZY75SXTcGH6HPTLtMSio50bcDcFnPulQFYmgbYJg');
  var sheetCVC = ssCVC.getSheetByName('Ensemble technique CVC');
  var dataCVC = sheetCVC.getDataRange().getValues();

  // Gestion du timeout
  var startTime = new Date().getTime();
  var maxExecutionTime = 5 * 60 * 1000; // 5 minutes
  var fichierTraites = 0;

  // ===== PARTIE 3: FILTRAGE ET TRI DES SITES =====
  var sitesATraiter = [];
  for (var i = 1; i < dataSites.length; i++) {
    if (dataSites[i][4] === "Prise en charge finalisée") {
      sitesATraiter.push({
        index: i,
        codeSite: String(dataSites[i][2]).trim(),
        timestamp: dataSites[i][8],
        nomClient: dataSites[i][0],
        nomSite: dataSites[i][3],
        codeClient: String(dataSites[i][1]).trim(),
        adresseSite: dataSites[i][10],
        prisEnChargePar: dataSites[i][9],
        urlPhotoSite: dataSites[i][7],
        colonneG: dataSites[i][6]
      });
    }
  }

  // Trier par timestamp décroissant (plus récent en premier)
  sitesATraiter.sort(function(a, b) {
    return new Date(b.timestamp) - new Date(a.timestamp);
  });

  Logger.log('Sites à traiter: ' + sitesATraiter.length + ' (statut "Prise en charge finalisée")');

  // ===== PARTIE 4: GESTION DE LA PROGRESSION =====
  var properties = PropertiesService.getScriptProperties();
  var dernierSiteTraite = parseInt(properties.getProperty('dernierSiteTraite') || '0');
  
  if (dernierSiteTraite > 0) {
    Logger.log('Reprise du traitement à partir du site ' + (dernierSiteTraite + 1));
  }

  // ===== PARTIE 5: CRÉATION DES INDEX =====
  // Index pour la page de garde
  var indexSites = {};
  for (var s = 0; s < sitesATraiter.length; s++) {
    var site = sitesATraiter[s];
    var key = site.nomClient + '||' + site.nomSite;
    indexSites[key] = {
      nomClient: site.nomClient,
      nomSite: site.nomSite,
      adresseSite: site.adresseSite,
      prisEnChargePar: site.prisEnChargePar,
      colonneG: site.colonneG,
      lastUpdate: site.timestamp,
      urlPhotoSite: site.urlPhotoSite,
      codeSite: site.codeSite
    };
  }

  // Index équipements par codeClient||codeSite
  var equipementsParClientSite = {};
  var compteursParClientSite = {};
  for (var i = 1; i < dataEquip.length; i++) {
    var codeClient = String(dataEquip[i][1]).trim();
    var codeSite = String(dataEquip[i][3]).trim();
    var typeEquip = dataEquip[i][11];
    if (!codeClient || !codeSite) continue;
    var key = codeClient + '||' + codeSite;
    if (!equipementsParClientSite[key]) equipementsParClientSite[key] = [];
    equipementsParClientSite[key].push(i);
    if (typeEquip === "Compteur & sous compteur") {
      if (!compteursParClientSite[key]) compteursParClientSite[key] = [];
      compteursParClientSite[key].push(i);
    }
  }

  // Index photos par codeClient||codeSite
  var photosParClientSite = {};
  for (var i = 1; i < dataPhoto.length; i++) {
    var codeClient = String(dataPhoto[i][2]).trim();
    var codeSite = String(dataPhoto[i][3]).trim();
    if (!codeClient || !codeSite) continue;
    var key = codeClient + '||' + codeSite;
    if (!photosParClientSite[key]) photosParClientSite[key] = [];
    photosParClientSite[key].push(i);
  }

  // Index propositions de travaux par codeClient||codeSite
  var travauxParClientSite = {};
  for (var i = 1; i < dataTravaux.length; i++) {
    var codeClient = String(dataTravaux[i][2]).trim();
    var codeSite = String(dataTravaux[i][3]).trim();
    if (!codeClient || !codeSite) continue;
    var key = codeClient + '||' + codeSite;
    if (!travauxParClientSite[key]) travauxParClientSite[key] = [];
    travauxParClientSite[key].push(i);
  }

  // CORRECTION: Index CVC par codeClient||codeSite
  var cvcParClientSite = {};
  for (var i = 1; i < dataCVC.length; i++) {
    var codeClient = String(dataCVC[i][2]).trim(); // Colonne C (index 2)
    var codeSite = String(dataCVC[i][3]).trim();   // Colonne D (index 3)
    Logger.log('Ligne CVC ' + i + ': Client=' + codeClient + ', Site=' + codeSite);
    if (!codeClient || !codeSite) continue;
    var key = codeClient + '||' + codeSite;
    if (!cvcParClientSite[key]) cvcParClientSite[key] = [];
    cvcParClientSite[key].push(i);
  }
  Logger.log('Index CVC créé: ' + Object.keys(cvcParClientSite).length + ' clés');

  // ===== PARTIE 6: INDEX DES FICHIERS =====
  var indexFichiers = {};
  var agencesFolders = folderRoot.getFolders();
  while (agencesFolders.hasNext()) {
    var agenceFolder = agencesFolders.next();
    var clientsFolders = agenceFolder.getFolders();
    while (clientsFolders.hasNext()) {
      var clientFolder = clientsFolders.next();
      var nomClient = clientFolder.getName();
      var files = clientFolder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        var nomSite = file.getName();
        var key = nomClient + '||' + nomSite;
        indexFichiers[key] = file;
      }
    }
  }

  Logger.log('Début du traitement des fichiers prioritaires...');

  // ===== PARTIE 7: TRAITEMENT DES SITES =====
  for (var s = dernierSiteTraite; s < sitesATraiter.length; s++) {
    // Vérification du timeout toutes les 5 fichiers
    if (fichierTraites % 5 === 0) {
      var currentTime = new Date().getTime();
      if (currentTime - startTime > maxExecutionTime) {
        Logger.log('TIMEOUT ATTEINT - Sites traités: ' + fichierTraites + '/' + sitesATraiter.length);
        properties.setProperty('dernierSiteTraite', s.toString());
        Logger.log('Progression sauvegardée. Relancer le script pour continuer.');
        return;
      }
    }

    var site = sitesATraiter[s];
    var key = site.nomClient + '||' + site.nomSite;
    var file = indexFichiers[key];
    
    if (!file) {
      Logger.log('Fichier non trouvé pour: ' + site.nomClient + ' - ' + site.nomSite + ' (code: ' + site.codeSite + ')');
      continue;
    }

    Logger.log('Traitement prioritaire: ' + site.nomClient + ' - ' + site.nomSite + ' (modifié: ' + site.timestamp + ')');

    var keyEquip = site.codeClient + '||' + site.codeSite;
    var ssSite = SpreadsheetApp.openById(file.getId());

    // ===== MISE À JOUR PAGE DE GARDE =====
    var sheetPageDeGarde = ssSite.getSheetByName('Page de garde');
    if (sheetPageDeGarde && indexSites[key]) {
      sheetPageDeGarde.getRange('C1').setValue(indexSites[key].nomClient);
      sheetPageDeGarde.getRange('C2').setValue(indexSites[key].nomSite);
      sheetPageDeGarde.getRange('C6').setValue(indexSites[key].adresseSite);
      sheetPageDeGarde.getRange('C12').setValue(indexSites[key].colonneG);
      sheetPageDeGarde.getRange('C9').setValue(indexSites[key].prisEnChargePar);
      sheetPageDeGarde.getRange('C10').setValue(indexSites[key].lastUpdate);
      if (indexSites[key].urlPhotoSite) {
        sheetPageDeGarde.getRange('B18').setFormula('=image("' + indexSites[key].urlPhotoSite + '")');
      }
    }

    // ===== MISE À JOUR INVENTAIRE =====
    var sheetInventaire = ssSite.getSheetByName('Inventaire');
    if (sheetInventaire && equipementsParClientSite[keyEquip]) {
      var lastRow = sheetInventaire.getLastRow();
      if (lastRow >= 12) {
        sheetInventaire.getRange(12, 1, lastRow - 11, 11).clearContent();
      }
      
      var nbEquip = equipementsParClientSite[keyEquip].length;
      var neededRows = 11 + nbEquip;
      var totalRows = sheetInventaire.getMaxRows();
      if (totalRows < neededRows) {
        sheetInventaire.insertRowsAfter(totalRows, neededRows - totalRows);
      }
      for (var k = 0; k < nbEquip; k++) {
        var idx = equipementsParClientSite[keyEquip][k];
        sheetInventaire.getRange(12 + k, 1).setValue(k + 1);
        sheetInventaire.getRange(12 + k, 2).setValue(dataEquip[idx][7]);
        sheetInventaire.getRange(12 + k, 3).setValue(dataEquip[idx][5]);
        sheetInventaire.getRange(12 + k, 4).setValue(dataEquip[idx][10]);
        sheetInventaire.getRange(12 + k, 5).setValue(dataEquip[idx][32]);
        sheetInventaire.getRange(12 + k, 6).setValue(dataEquip[idx][33]);
        sheetInventaire.getRange(12 + k, 7).setValue(dataEquip[idx][8]);
        sheetInventaire.getRange(12 + k, 8).setValue(dataEquip[idx][41]);
        sheetInventaire.getRange(12 + k, 9).setValue(dataEquip[idx][28]);
        sheetInventaire.getRange(12 + k, 10).setFormula(dataEquip[idx][26] ? '=image("' + dataEquip[idx][26] + '")' : '');
        sheetInventaire.getRange(12 + k, 11).setFormula(dataEquip[idx][27] ? '=image("' + dataEquip[idx][27] + '")' : '');
      }
    }

    // ===== MISE À JOUR RÉPERTOIRE COMPTEUR =====
    var sheetRepCompteur = ssSite.getSheetByName('Répertoire compteur');
    if (sheetRepCompteur && compteursParClientSite[keyEquip]) {
      var lastRow = sheetRepCompteur.getLastRow();
      if (lastRow >= 12) {
        sheetRepCompteur.getRange(12, 1, lastRow - 11, 10).clearContent();
      }
      
      var nbCompteurs = compteursParClientSite[keyEquip].length;
      var neededRows = 11 + nbCompteurs;
      var totalRows = sheetRepCompteur.getMaxRows();
      if (totalRows < neededRows) {
        sheetRepCompteur.insertRowsAfter(totalRows, neededRows - totalRows);
      }
      for (var k = 0; k < nbCompteurs; k++) {
        var idx = compteursParClientSite[keyEquip][k];
        sheetRepCompteur.getRange(12 + k, 1).setValue(k + 1);
        sheetRepCompteur.getRange(12 + k, 2).setValue(dataEquip[idx][7]);
        sheetRepCompteur.getRange(12 + k, 3).setValue(dataEquip[idx][10]);
        sheetRepCompteur.getRange(12 + k, 4).setValue(dataEquip[idx][17]);
        sheetRepCompteur.getRange(12 + k, 5).setValue(dataEquip[idx][20]);
        sheetRepCompteur.getRange(12 + k, 6).setValue(dataEquip[idx][21]);
        sheetRepCompteur.getRange(12 + k, 7).setValue(dataEquip[idx][47]);
        sheetRepCompteur.getRange(12 + k, 8).setValue(dataEquip[idx][23]);
        sheetRepCompteur.getRange(12 + k, 9).setFormula(dataEquip[idx][26] ? '=image("' + dataEquip[idx][26] + '")' : '');
        sheetRepCompteur.getRange(12 + k, 10).setFormula(dataEquip[idx][27] ? '=image("' + dataEquip[idx][27] + '")' : '');
      }
    }

    // ===== MISE À JOUR PHOTOTHÈQUE =====
    var sheetPhototheque = ssSite.getSheetByName('Photohèque');
    if (sheetPhototheque && photosParClientSite[keyEquip]) {
      var lastRow = sheetPhototheque.getLastRow();
      if (lastRow >= 9) {
        sheetPhototheque.getRange(9, 1, lastRow - 8, 5).clearContent();
      }
      
      var nbPhotos = photosParClientSite[keyEquip].length;
      var neededRows = 8 + nbPhotos;
      var totalRows = sheetPhototheque.getMaxRows();
      if (totalRows < neededRows) {
        sheetPhototheque.insertRowsAfter(totalRows, neededRows - totalRows);
      }
      for (var k = 0; k < nbPhotos; k++) {
        var idx = photosParClientSite[keyEquip][k];
        sheetPhototheque.getRange(9 + k, 1).setValue(k + 1);
        sheetPhototheque.getRange(9 + k, 2).setValue(dataPhoto[idx][8]);
        sheetPhototheque.getRange(9 + k, 3).setValue(dataPhoto[idx][9]);
        sheetPhototheque.getRange(9 + k, 4).setValue(dataPhoto[idx][5]);
        sheetPhototheque.getRange(9 + k, 5).setFormula(dataPhoto[idx][7] ? '=image("' + dataPhoto[idx][7] + '")' : '');
      }
    }

    // ===== MISE À JOUR PROPOSITIONS DE TRAVAUX - P5 =====
    var sheetP5 = ssSite.getSheetByName('Propositions de travaux - P5');
    if (sheetP5 && travauxParClientSite[keyEquip]) {
      var lastRow = sheetP5.getLastRow();
      if (lastRow >= 12) {
        sheetP5.getRange(12, 1, lastRow - 11, 7).clearContent();
      }
      
      var nbTravaux = travauxParClientSite[keyEquip].length;
      var neededRows = 11 + nbTravaux;
      var totalRows = sheetP5.getMaxRows();
      if (totalRows < neededRows) {
        sheetP5.insertRowsAfter(totalRows, neededRows - totalRows);
      }
      for (var k = 0; k < nbTravaux; k++) {
        var idx = travauxParClientSite[keyEquip][k];
        sheetP5.getRange(12 + k, 1).setValue(k + 1);
        sheetP5.getRange(12 + k, 2).setValue(dataTravaux[idx][7]);
        sheetP5.getRange(12 + k, 3).setValue(dataTravaux[idx][5]);
        sheetP5.getRange(12 + k, 4).setValue(dataTravaux[idx][6]);
        sheetP5.getRange(12 + k, 5).setFormula(dataTravaux[idx][10] ? '=image("' + dataTravaux[idx][10] + '")' : '');
        sheetP5.getRange(12 + k, 6).setFormula(dataTravaux[idx][11] ? '=image("' + dataTravaux[idx][11] + '")' : '');
        sheetP5.getRange(12 + k, 7).setValue(dataTravaux[idx][18]);
      }
    }

    // ===== MISE À JOUR ETAT RÉGLEMENTAIRE - RÉCAPITULATIF =====
    var sheetEtatReglem = ssSite.getSheetByName('Etat réglementaire - Récapitulatif');
    if (sheetEtatReglem && cvcParClientSite[keyEquip]) {
      Logger.log('Traitement CVC pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');
      
      var lastRow = sheetEtatReglem.getLastRow();
      if (lastRow >= 12) {
        sheetEtatReglem.getRange(12, 1, lastRow - 11, 8).clearContent();
      }
      
      var nbCVC = cvcParClientSite[keyEquip].length;
      var neededRows = 11 + nbCVC;
      var totalRows = sheetEtatReglem.getMaxRows();
      if (totalRows < neededRows) {
        sheetEtatReglem.insertRowsAfter(totalRows, neededRows - totalRows);
      }
      for (var k = 0; k < nbCVC; k++) {
        var idx = cvcParClientSite[keyEquip][k];
        sheetEtatReglem.getRange(12 + k, 1).setValue(k + 1); // A12: Numérotation
        sheetEtatReglem.getRange(12 + k, 2).setValue(dataCVC[idx][5]); // B12: Type de local (colonne F)
        sheetEtatReglem.getRange(12 + k, 3).setValue(dataCVC[idx][6]); // C12: Désignation de l'installation (colonne G)
        sheetEtatReglem.getRange(12 + k, 4).setValue(dataCVC[idx][11]); // D12: Situation de l'installation (colonne L)
        sheetEtatReglem.getRange(12 + k, 5).setValue(dataCVC[idx][8]); // E12: Type d'installation (colonne I)
        sheetEtatReglem.getRange(12 + k, 6).setValue(dataCVC[idx][13]); // F12: Energie alimentante (colonne N)
        sheetEtatReglem.getRange(12 + k, 7).setValue(dataCVC[idx][12]); // G12: Puissance total (kW) (colonne M)
        sheetEtatReglem.getRange(12 + k, 8).setValue(dataCVC[idx][15]); // H12: Production d'ECS (colonne P)
        sheetEtatReglem.getRange(12 + k, 9).setValue(dataCVC[idx][14]); // I12: Schéma de principe chaufferie (colonne O)
      }
    }
// ===== NOUVELLE PARTIE: MISE À JOUR ETAT RÉGLEMENTAIRE - QHSE RISQUE HAUTEUR =====
var sheetQHSERisque = ssSite.getSheetByName('Etat réglementaire - QHSE Risque hauteur');
if (sheetQHSERisque && cvcParClientSite[keyEquip]) {
  var lastRow = sheetQHSERisque.getLastRow();
  if (lastRow >= 12) {
    sheetQHSERisque.getRange(12, 1, lastRow - 11, 7).clearContent();
  }
  
  var nbCVC = cvcParClientSite[keyEquip].length;
  var neededRows = 11 + nbCVC;
  var totalRows = sheetQHSERisque.getMaxRows();
  if (totalRows < neededRows) {
    sheetQHSERisque.insertRowsAfter(totalRows, neededRows - totalRows);
  }
  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];
    sheetQHSERisque.getRange(12 + k, 1).setValue(k + 1); // A12: Numérotation
    sheetQHSERisque.getRange(12 + k, 2).setValue(dataCVC[idx][5]); // B12: Type de local (colonne F)
    sheetQHSERisque.getRange(12 + k, 3).setValue(dataCVC[idx][6]); // C12: Désignation de l'installation (colonne G)
    sheetQHSERisque.getRange(12 + k, 4).setValue(dataCVC[idx][16]); // D12: Risque de chute en hauteur (colonne Q)
    sheetQHSERisque.getRange(12 + k, 5).setValue(dataCVC[idx][17]); // E12: Garde de corps présent ? (colonne R)
    sheetQHSERisque.getRange(12 + k, 6).setValue(dataCVC[idx][18]); // F12: Ligne de vie présente ? (colonne S)
    sheetQHSERisque.getRange(12 + k, 7).setValue(dataCVC[idx][19]); // G12: Point d'encrage présent ? (colonne T)
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - QHSE ESPACE CONFINÉ =====
var sheetQHSEEspaceConfine = ssSite.getSheetByName('Etat réglementaire - QHSE Espace confiné');
if (sheetQHSEEspaceConfine && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement QHSE Espace confiné pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');
  
  var lastRow = sheetQHSEEspaceConfine.getLastRow();
  if (lastRow >= 12) {
    sheetQHSEEspaceConfine.getRange(12, 1, lastRow - 11, 9).clearContent();
  }
  
  var nbCVC = cvcParClientSite[keyEquip].length;
  var neededRows = 11 + nbCVC;
  var totalRows = sheetQHSEEspaceConfine.getMaxRows();
  if (totalRows < neededRows) {
    sheetQHSEEspaceConfine.insertRowsAfter(totalRows, neededRows - totalRows);
  }
  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];
    sheetQHSEEspaceConfine.getRange(12 + k, 1).setValue(k + 1); // A12: Numérotation
    sheetQHSEEspaceConfine.getRange(12 + k, 2).setValue(dataCVC[idx][5]); // B12: Type de local (colonne F)
    sheetQHSEEspaceConfine.getRange(12 + k, 3).setValue(dataCVC[idx][6]); // C12: Désignation de l'installation (colonne G)
    sheetQHSEEspaceConfine.getRange(12 + k, 4).setValue(dataCVC[idx][25]); // D12: Risque espace confiné (colonne Z)
    sheetQHSEEspaceConfine.getRange(12 + k, 5).setValue(dataCVC[idx][26]); // E12: Risque espace confiné (colonne AA)
    sheetQHSEEspaceConfine.getRange(12 + k, 6).setValue(dataCVC[idx][27]); // F12: Environnement de l'espace confiné (colonne AB)
    sheetQHSEEspaceConfine.getRange(12 + k, 7).setValue(dataCVC[idx][28]); // G12: Accès à l'espace confiné (colonne AC)
    sheetQHSEEspaceConfine.getRange(12 + k, 8).setValue(dataCVC[idx][29]); // H12: Profondeur de l'espace confiné (colonne AD)
    sheetQHSEEspaceConfine.getRange(12 + k, 9).setValue(dataCVC[idx][30]); // I12: Eclairage de l'espace confiné (colonne AE)
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - QHSE RISQUE EN COMBLE =====
var sheetQHSERisqueComble = ssSite.getSheetByName('Etat réglementaire - QHSE Risque en comble');
if (sheetQHSERisqueComble && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement QHSE Risque en comble pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');
  
  var lastRow = sheetQHSERisqueComble.getLastRow();
  if (lastRow >= 12) {
    sheetQHSERisqueComble.getRange(12, 1, lastRow - 11, 8).clearContent();
  }
  
  var nbCVC = cvcParClientSite[keyEquip].length;
  var neededRows = 11 + nbCVC;
  var totalRows = sheetQHSERisqueComble.getMaxRows();
  if (totalRows < neededRows) {
    sheetQHSERisqueComble.insertRowsAfter(totalRows, neededRows - totalRows);
  }
  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];
    sheetQHSERisqueComble.getRange(12 + k, 1).setValue(k + 1); // A12: Numérotation
    sheetQHSERisqueComble.getRange(12 + k, 2).setValue(dataCVC[idx][5]); // B12: Type de local (colonne F)
    sheetQHSERisqueComble.getRange(12 + k, 3).setValue(dataCVC[idx][6]); // C12: Désignation de l'installation (colonne G)
    sheetQHSERisqueComble.getRange(12 + k, 4).setValue(dataCVC[idx][20]); // D12: Risque en comble (colonne U)
    sheetQHSERisqueComble.getRange(12 + k, 5).setValue(dataCVC[idx][21]); // E12: Accès aux combles sécurisé ? (colonne V)
    sheetQHSERisqueComble.getRange(12 + k, 6).setValue(dataCVC[idx][22]); // F12: Éclairage suffisant dans les combles ? (colonne W)
    sheetQHSERisqueComble.getRange(12 + k, 7).setValue(dataCVC[idx][23]); // G12: Plancher dégagé et bien visible ? (colonne X)
    sheetQHSERisqueComble.getRange(12 + k, 8).setValue(dataCVC[idx][24]); // H12: Plancher solide ? (colonne Y)
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - QHSE RISQUE DIVERS =====
var sheetQHSERisqueDivers = ssSite.getSheetByName('Etat réglementaire - QHSE Risque divers');
if (sheetQHSERisqueDivers && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement QHSE Risque divers pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');
  
  var lastRow = sheetQHSERisqueDivers.getLastRow();
  if (lastRow >= 12) {
    sheetQHSERisqueDivers.getRange(12, 1, lastRow - 11, 9).clearContent();
  }
  
  var nbCVC = cvcParClientSite[keyEquip].length;
  var neededRows = 11 + nbCVC;
  var totalRows = sheetQHSERisqueDivers.getMaxRows();
  if (totalRows < neededRows) {
    sheetQHSERisqueDivers.insertRowsAfter(totalRows, neededRows - totalRows);
  }
  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];
    sheetQHSERisqueDivers.getRange(12 + k, 1).setValue(k + 1); // A12: Numérotation
    sheetQHSERisqueDivers.getRange(12 + k, 2).setValue(dataCVC[idx][5]); // B12: Type de local (colonne F)
    sheetQHSERisqueDivers.getRange(12 + k, 3).setValue(dataCVC[idx][6]); // C12: Désignation de l'installation (colonne G)
    sheetQHSERisqueDivers.getRange(12 + k, 4).setValue(dataCVC[idx][31]); // D12: Encombrement du local (colonne AF)
    sheetQHSERisqueDivers.getRange(12 + k, 5).setValue(dataCVC[idx][32]); // E12: Risque chute de plain pied (colonne AG)
    sheetQHSERisqueDivers.getRange(12 + k, 6).setValue(dataCVC[idx][33]); // F12: Risque de heurt (colonne AH)
    sheetQHSERisqueDivers.getRange(12 + k, 7).setValue(dataCVC[idx][34]); // G12: Accès au local sécurisé ? (colonne AI)
    sheetQHSERisqueDivers.getRange(12 + k, 8).setValue(dataCVC[idx][35]); // H12: Accès en hauteur sécurisé ? (colonne AJ)
    sheetQHSERisqueDivers.getRange(12 + k, 9).setValue(dataCVC[idx][36]); // I12: Bâtiment construit avant 1997 ? (colonne AK)
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - SÉCURITÉ TECHNIQUE =====
var sheetSecuriteTechnique = ssSite.getSheetByName('Etat réglementaire - Sécurité technique');
if (sheetSecuriteTechnique && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement Sécurité technique pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');
  
  var lastRow = sheetSecuriteTechnique.getLastRow();
  if (lastRow >= 12) {
    sheetSecuriteTechnique.getRange(12, 1, lastRow - 11, 16).clearContent();
  }
  
  var nbCVC = cvcParClientSite[keyEquip].length;
  var neededRows = 11 + nbCVC;
  var totalRows = sheetSecuriteTechnique.getMaxRows();
  if (totalRows < neededRows) {
    sheetSecuriteTechnique.insertRowsAfter(totalRows, neededRows - totalRows);
  }
  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];
    sheetSecuriteTechnique.getRange(12 + k, 1).setValue(k + 1); // A12: Numérotation
    sheetSecuriteTechnique.getRange(12 + k, 2).setValue(dataCVC[idx][5]); // B12: Type de local (colonne F)
    sheetSecuriteTechnique.getRange(12 + k, 3).setValue(dataCVC[idx][6]); // C12: Désignation de l'installation (colonne G)
    sheetSecuriteTechnique.getRange(12 + k, 4).setValue(dataCVC[idx][89]); // D12: Portes coupe feu ou SAS ? (colonne CL)
    sheetSecuriteTechnique.getRange(12 + k, 5).setValue(dataCVC[idx][90]); // E12: Equipement des portes avec serrures anti-panique (colonne CM)
    sheetSecuriteTechnique.getRange(12 + k, 6).setValue(dataCVC[idx][91]); // F12: Ferme porte automatique (colonne CN)
    sheetSecuriteTechnique.getRange(12 + k, 7).setValue(dataCVC[idx][96]); // G12: Contrôle périodique extincteurs (colonne CS)
    sheetSecuriteTechnique.getRange(12 + k, 8).setValue(dataCVC[idx][97]); // H12: Date contrôle périodique extincteurs (colonne CT)
    sheetSecuriteTechnique.getRange(12 + k, 9).setValue(dataCVC[idx][95]); // I12: Blocs de secours (colonne CR)
    sheetSecuriteTechnique.getRange(12 + k, 10).setValue(dataCVC[idx][102]); // J12: Pressostat de Sécurité (colonne CY)
    sheetSecuriteTechnique.getRange(12 + k, 11).setValue(dataCVC[idx][103]); // K12: Soupapes de sécurité non isolables (colonne CZ)
    sheetSecuriteTechnique.getRange(12 + k, 12).setValue(dataCVC[idx][87]); // L12: Dispositif d'arrêt d'urgence (colonne CJ)
    sheetSecuriteTechnique.getRange(12 + k, 13).setValue(dataCVC[idx][88]); // M12: Vanne de barrage sous coffret (colonne CK)
    sheetSecuriteTechnique.getRange(12 + k, 14).setValue(dataCVC[idx][104]); // N12: Conformité au décret BACS (colonne DA)
    sheetSecuriteTechnique.getRange(12 + k, 15).setValue(dataCVC[idx][106]); // O12: Dispositif de régulation T°C extérieur (colonne DC)
    sheetSecuriteTechnique.getRange(12 + k, 16).setValue(dataCVC[idx][105]); // P12: Dispositif de régulation > 30kW (colonne DB)
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - SÉCURITÉ TECHNIQUE 2 =====
var sheetSecuriteTechnique2 = ssSite.getSheetByName('Etat réglementaire - Sécurité technique 2');
if (sheetSecuriteTechnique2 && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement Sécurité technique 2 pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');
  
  var lastRow = sheetSecuriteTechnique2.getLastRow();
  if (lastRow >= 12) {
    sheetSecuriteTechnique2.getRange(12, 1, lastRow - 11, 9).clearContent();
  }
  
  var nbCVC = cvcParClientSite[keyEquip].length;
  var neededRows = 11 + nbCVC;
  var totalRows = sheetSecuriteTechnique2.getMaxRows();
  if (totalRows < neededRows) {
    sheetSecuriteTechnique2.insertRowsAfter(totalRows, neededRows - totalRows);
  }
  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];
    sheetSecuriteTechnique2.getRange(12 + k, 1).setValue(k + 1); // A12: Numérotation
    sheetSecuriteTechnique2.getRange(12 + k, 2).setValue(dataCVC[idx][5]); // B12: Type de local (colonne F)
    sheetSecuriteTechnique2.getRange(12 + k, 3).setValue(dataCVC[idx][6]); // C12: Désignation de l'installation (colonne G)
    sheetSecuriteTechnique2.getRange(12 + k, 4).setValue(dataCVC[idx][94]); // D12: Eclairage anti-déflagrant (colonne CQ)
    sheetSecuriteTechnique2.getRange(12 + k, 5).setValue(dataCVC[idx][93]); // E12: Eclairage normal suffisant (colonne CP)
    sheetSecuriteTechnique2.getRange(12 + k, 6).setValue(dataCVC[idx][107]); // F12: Etat des ventilations hautes (colonne DD)
    sheetSecuriteTechnique2.getRange(12 + k, 7).setValue(dataCVC[idx][108]); // G12: Dimension des ventilations hautes (colonne DE)
    sheetSecuriteTechnique2.getRange(12 + k, 8).setValue(dataCVC[idx][109]); // H12: Etat des ventilations basses (colonne DF)
    sheetSecuriteTechnique2.getRange(12 + k, 9).setValue(dataCVC[idx][110]); // I12: Dimension des ventilations basses (colonne DG)
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - HYDRAULIQUE =====
var sheetHydraulique = ssSite.getSheetByName('Etat réglementaire - Hydraulique');
if (sheetHydraulique && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement Hydraulique pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');
  
  var lastRow = sheetHydraulique.getLastRow();
  if (lastRow >= 12) {
    sheetHydraulique.getRange(12, 1, lastRow - 11, 16).clearContent();
  }
  
  var nbCVC = cvcParClientSite[keyEquip].length;
  var neededRows = 11 + nbCVC;
  var totalRows = sheetHydraulique.getMaxRows();
  if (totalRows < neededRows) {
    sheetHydraulique.insertRowsAfter(totalRows, neededRows - totalRows);
  }
  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];
    sheetHydraulique.getRange(12 + k, 1).setValue(k + 1); // A12: Numérotation
    sheetHydraulique.getRange(12 + k, 2).setValue(dataCVC[idx][5]); // B12: Type de local (colonne F)
    sheetHydraulique.getRange(12 + k, 3).setValue(dataCVC[idx][6]); // C12: Désignation de l'installation (colonne G)
    sheetHydraulique.getRange(12 + k, 4).setValue(dataCVC[idx][60]); // D12: Disconnecteur de remplissage (colonne BI)
    sheetHydraulique.getRange(12 + k, 5).setValue(dataCVC[idx][99]); // E12: Disconnecteur contrôle (colonne CV)
    sheetHydraulique.getRange(12 + k, 6).setValue(dataCVC[idx][100]); // F12: Date de contrôle disconnecteur (colonne CW)
    sheetHydraulique.getRange(12 + k, 7).setValue(dataCVC[idx][61]); // G12: Type d'expansion (colonne BJ)
    sheetHydraulique.getRange(12 + k, 8).setValue(dataCVC[idx][62]); // H12: Modèle d'expansion (colonne BK)
    sheetHydraulique.getRange(12 + k, 9).setValue(dataCVC[idx][63]); // I12: Remplissage (colonne BL)
    sheetHydraulique.getRange(12 + k, 10).setValue(dataCVC[idx][101]); // J12: Compteur d'appoint d'eau (colonne CX)
    sheetHydraulique.getRange(12 + k, 11).setValue(dataCVC[idx][64]); // K12: Pot à boues tangentielle (colonne BM)
    sheetHydraulique.getRange(12 + k, 12).setValue(dataCVC[idx][65]); // L12: Clarificateur magnétique (colonne BN)
    sheetHydraulique.getRange(12 + k, 13).setValue(dataCVC[idx][66]); // M12: Raccordement hydraulique clarificateur (colonne BO)
    sheetHydraulique.getRange(12 + k, 14).setValue(dataCVC[idx][67]); // N12: Pompe irrigation clarificateur (colonne BP)
    sheetHydraulique.getRange(12 + k, 15).setValue(dataCVC[idx][68]); // O12: Filtre à tamis (colonne BQ)
    sheetHydraulique.getRange(12 + k, 16).setValue(dataCVC[idx][69]); // P12: Système injection désembouant (colonne BR)
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - HYDRAULIQUE 2 =====
var sheetHydraulique2 = ssSite.getSheetByName('Etat réglementaire - Hydraulique 2');
if (sheetHydraulique2 && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement Hydraulique 2 pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');

  // Nettoyage de l’ancienne zone (A:I à partir de la ligne 12)
  var lastRow = sheetHydraulique2.getLastRow();
  if (lastRow >= 12) {
    sheetHydraulique2.getRange(12, 1, lastRow - 11, 9).clearContent(); // 9 colonnes = A..I
  }

  // Filtrer les lignes CVC à écrire: seulement si BS/BT/BU/BV/BW/CO a au moins une valeur
  var rowsToWrite = [];
  var nbCVC = cvcParClientSite[keyEquip].length;

  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];

    // Colonnes demandées (index = lettreExcel-1)
    // F=5, G=6
    // BS=70, BT=71, BU=72, BV=73, BW=74
    // CO=92
    var bs = dataCVC[idx][70];
    var bt = dataCVC[idx][71];
    var bu = dataCVC[idx][72];
    var bv = dataCVC[idx][73];
    var bw = dataCVC[idx][74];
    var co = dataCVC[idx][92];

    // condition: on écrit la ligne si au moins une des 6 colonnes est renseignée
    var hasAny = (bs !== '' && bs != null) ||
                 (bt !== '' && bt != null) ||
                 (bu !== '' && bu != null) ||
                 (bv !== '' && bv != null) ||
                 (bw !== '' && bw != null) ||
                 (co !== '' && co != null);

    if (!hasAny) continue;

    rowsToWrite.push([
      rowsToWrite.length + 1,    // A: numérotation 1..n (après filtrage)
      dataCVC[idx][5],           // B: colonne F
      dataCVC[idx][6],           // C: colonne G
      bs,                        // D: colonne BS
      bt,                        // E: colonne BT
      bu,                        // F: colonne BU
      bv,                        // G: colonne BV
      bw,                        // H: colonne BW
      co                         // I: colonne CO
    ]);
  }

  // Ajuster le nombre de lignes si besoin
  var neededRows = 11 + rowsToWrite.length;
  var totalRows = sheetHydraulique2.getMaxRows();
  if (totalRows < neededRows) {
    sheetHydraulique2.insertRowsAfter(totalRows, neededRows - totalRows);
  }

  // Écriture en une fois (plus rapide + moins de risques de timeout)
  if (rowsToWrite.length > 0) {
    sheetHydraulique2.getRange(12, 1, rowsToWrite.length, 9).setValues(rowsToWrite);
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - FROID =====
var sheetFroid = ssSite.getSheetByName('Etat réglementaire - Froid');
if (sheetFroid && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement Froid pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');

  // Nettoyage de l’ancienne zone (A:I à partir de la ligne 12)
  var lastRow = sheetFroid.getLastRow();
  if (lastRow >= 12) {
    sheetFroid.getRange(12, 1, lastRow - 11, 9).clearContent(); // 9 colonnes = A..I
  }

  // Filtrer les lignes CVC à écrire: seulement si BC/BD/BE/BF/BG/BH a au moins une valeur
  var rowsToWrite = [];
  var nbCVC = cvcParClientSite[keyEquip].length;

  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];

    // Colonnes demandées (index = lettreExcel-1)
    // F=5, G=6
    // BC=54, BD=55, BE=56, BF=57, BG=58, BH=59
    var bc = dataCVC[idx][54];
    var bd = dataCVC[idx][55];
    var be = dataCVC[idx][56];
    var bf = dataCVC[idx][57];
    var bg = dataCVC[idx][58];
    var bh = dataCVC[idx][59];

    // condition: on écrit la ligne si au moins une des 6 colonnes est renseignée
    var hasAny = (bc !== '' && bc != null) ||
                 (bd !== '' && bd != null) ||
                 (be !== '' && be != null) ||
                 (bf !== '' && bf != null) ||
                 (bg !== '' && bg != null) ||
                 (bh !== '' && bh != null);

    if (!hasAny) continue;

    rowsToWrite.push([
      rowsToWrite.length + 1, // A: numérotation 1..n (après filtrage)
      dataCVC[idx][5],        // B: colonne F
      dataCVC[idx][6],        // C: colonne G
      bc,                     // D: colonne BC
      bd,                     // E: colonne BD
      be,                     // F: colonne BE
      bf,                     // G: colonne BF
      bg,                     // H: colonne BG
      bh                      // I: colonne BH
    ]);
  }

  // Ajuster le nombre de lignes si besoin
  var neededRows = 11 + rowsToWrite.length;
  var totalRows = sheetFroid.getMaxRows();
  if (totalRows < neededRows) {
    sheetFroid.insertRowsAfter(totalRows, neededRows - totalRows);
  }

  // Écriture en une fois
  if (rowsToWrite.length > 0) {
    sheetFroid.getRange(12, 1, rowsToWrite.length, 9).setValues(rowsToWrite);
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - INSTALLATION GAZ =====
var sheetGaz = ssSite.getSheetByName('Etat réglementaire - Installation gaz');
if (sheetGaz && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement Installation gaz pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');

  // Nettoyage de l’ancienne zone (A:I à partir de la ligne 12)
  var lastRow = sheetGaz.getLastRow();
  if (lastRow >= 12) {
    sheetGaz.getRange(12, 1, lastRow - 11, 9).clearContent(); // 9 colonnes = A..I
  }

  // Filtrer les lignes CVC à écrire: seulement si DH/DI/DJ/DK/DT/DW a au moins une valeur
  var rowsToWrite = [];
  var nbCVC = cvcParClientSite[keyEquip].length;

  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];

    // Colonnes demandées (index = lettreExcel-1)
    // F=5, G=6
    // DH=111, DI=112, DJ=113, DK=114, DT=123, DW=126
    var dh = dataCVC[idx][111];
    var di = dataCVC[idx][112];
    var dj = dataCVC[idx][113];
    var dk = dataCVC[idx][114];
    var dt = dataCVC[idx][123];
    var dw = dataCVC[idx][126];

    // condition: on écrit la ligne si au moins une des 6 colonnes est renseignée
    var hasAny = (dh !== '' && dh != null) ||
                 (di !== '' && di != null) ||
                 (dj !== '' && dj != null) ||
                 (dk !== '' && dk != null) ||
                 (dt !== '' && dt != null) ||
                 (dw !== '' && dw != null);

    if (!hasAny) continue;

    rowsToWrite.push([
      rowsToWrite.length + 1, // A: numérotation 1..n (après filtrage)
      dataCVC[idx][5],        // B: colonne F
      dataCVC[idx][6],        // C: colonne G
      dh,                     // D: colonne DH
      di,                     // E: colonne DI
      dj,                     // F: colonne DJ
      dk,                     // G: colonne DK
      dt,                     // H: colonne DT
      dw                      // I: colonne DW
    ]);
  }

  // Ajuster le nombre de lignes si besoin
  var neededRows = 11 + rowsToWrite.length;
  var totalRows = sheetGaz.getMaxRows();
  if (totalRows < neededRows) {
    sheetGaz.insertRowsAfter(totalRows, neededRows - totalRows);
  }

  // Écriture en une fois
  if (rowsToWrite.length > 0) {
    sheetGaz.getRange(12, 1, rowsToWrite.length, 9).setValues(rowsToWrite);
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - INSTALLATION FOD =====
var sheetFOD = ssSite.getSheetByName('Etat réglementaire - Installation FOD');
if (sheetFOD && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement Installation FOD pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');

  // Nettoyage de l’ancienne zone (A:P à partir de la ligne 12)
  var lastRow = sheetFOD.getLastRow();
  if (lastRow >= 12) {
    sheetFOD.getRange(12, 1, lastRow - 11, 16).clearContent(); // 16 colonnes = A..P
  }

  // Filtrer les lignes CVC à écrire: seulement si une des colonnes EC..EO a au moins une valeur
  var rowsToWrite = [];
  var nbCVC = cvcParClientSite[keyEquip].length;

  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];

    // Colonnes demandées (index = lettreExcel-1)
    // F=5, G=6
    // EC=132, ED=133, EE=134, EF=135, EG=136, EH=137, EI=138, EJ=139,
    // EK=140, EL=141, EM=142, EN=143, EO=144
    var ec = dataCVC[idx][132];
    var ed = dataCVC[idx][133];
    var ee = dataCVC[idx][134];
    var ef = dataCVC[idx][135];
    var eg = dataCVC[idx][136];
    var eh = dataCVC[idx][137];
    var ei = dataCVC[idx][138];
    var ej = dataCVC[idx][139];
    var ek = dataCVC[idx][140];
    var el = dataCVC[idx][141];
    var em = dataCVC[idx][142];
    var en = dataCVC[idx][143];
    var eo = dataCVC[idx][144];

    // condition: on écrit la ligne si au moins une des colonnes EC..EO est renseignée
    var hasAny = (ec !== '' && ec != null) ||
                 (ed !== '' && ed != null) ||
                 (ee !== '' && ee != null) ||
                 (ef !== '' && ef != null) ||
                 (eg !== '' && eg != null) ||
                 (eh !== '' && eh != null) ||
                 (ei !== '' && ei != null) ||
                 (ej !== '' && ej != null) ||
                 (ek !== '' && ek != null) ||
                 (el !== '' && el != null) ||
                 (em !== '' && em != null) ||
                 (en !== '' && en != null) ||
                 (eo !== '' && eo != null);

    if (!hasAny) continue;

    rowsToWrite.push([
      rowsToWrite.length + 1, // A: numérotation 1..n (après filtrage)
      dataCVC[idx][5],        // B: colonne F
      dataCVC[idx][6],        // C: colonne G
      ec,                     // D: colonne EC
      ed,                     // E: colonne ED
      ee,                     // F: colonne EE
      ef,                     // G: colonne EF
      eg,                     // H: colonne EG
      eh,                     // I: colonne EH
      ei,                     // J: colonne EI
      ej,                     // K: colonne EJ
      ek,                     // L: colonne EK
      el,                     // M: colonne EL
      em,                     // N: colonne EM
      en,                     // O: colonne EN
      eo                      // P: colonne EO
    ]);
  }

  // Ajuster le nombre de lignes si besoin
  var neededRows = 11 + rowsToWrite.length;
  var totalRows = sheetFOD.getMaxRows();
  if (totalRows < neededRows) {
    sheetFOD.insertRowsAfter(totalRows, neededRows - totalRows);
  }

  // Écriture en une fois
  if (rowsToWrite.length > 0) {
    sheetFOD.getRange(12, 1, rowsToWrite.length, 16).setValues(rowsToWrite);
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - INSTALLATION ELECTRIQUE =====
var sheetElec = ssSite.getSheetByName('Etat réglementaire - Installation Electrique');
if (sheetElec && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement Installation Electrique pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');

  // Nettoyage de l’ancienne zone (A:Q à partir de la ligne 12)
  var lastRow = sheetElec.getLastRow();
  if (lastRow >= 12) {
    sheetElec.getRange(12, 1, lastRow - 11, 17).clearContent(); // 17 colonnes = A..Q
  }

  // Filtrer les lignes CVC à écrire: seulement si une des colonnes AL..AW + AZ + BB a au moins une valeur
  var rowsToWrite = [];
  var nbCVC = cvcParClientSite[keyEquip].length;

  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];

    // Colonnes demandées (index = lettreExcel-1)
    // F=5, G=6
    // AL=37, AM=38, AN=39, AO=40, AP=41, AQ=42, AR=43, AS=44, AT=45, AU=46, AV=47, AW=48
    // AZ=51, BB=53
    var al = dataCVC[idx][37];
    var am = dataCVC[idx][38];
    var an = dataCVC[idx][39];
    var ao = dataCVC[idx][40];
    var ap = dataCVC[idx][41];
    var aq = dataCVC[idx][42];
    var ar = dataCVC[idx][43];
    var as = dataCVC[idx][44];
    var at = dataCVC[idx][45];
    var au = dataCVC[idx][46];
    var av = dataCVC[idx][47];
    var aw = dataCVC[idx][48];
    var az = dataCVC[idx][51];
    var bb = dataCVC[idx][53];

    // condition: on écrit la ligne si au moins une des colonnes est renseignée
    var hasAny = (al !== '' && al != null) ||
                 (am !== '' && am != null) ||
                 (an !== '' && an != null) ||
                 (ao !== '' && ao != null) ||
                 (ap !== '' && ap != null) ||
                 (aq !== '' && aq != null) ||
                 (ar !== '' && ar != null) ||
                 (as !== '' && as != null) ||
                 (at !== '' && at != null) ||
                 (au !== '' && au != null) ||
                 (av !== '' && av != null) ||
                 (aw !== '' && aw != null) ||
                 (az !== '' && az != null) ||
                 (bb !== '' && bb != null);

    if (!hasAny) continue;

    rowsToWrite.push([
      rowsToWrite.length + 1, // A: numérotation 1..n (après filtrage)
      dataCVC[idx][5],        // B: colonne F
      dataCVC[idx][6],        // C: colonne G
      al,                     // D: colonne AL
      am,                     // E: colonne AM
      an,                     // F: colonne AN
      ao,                     // G: colonne AO
      ap,                     // H: colonne AP
      aq,                     // I: colonne AQ
      ar,                     // J: colonne AR
      as,                     // K: colonne AS
      at,                     // L: colonne AT
      au,                     // M: colonne AU
      av,                     // N: colonne AV
      aw,                     // O: colonne AW
      az,                     // P: colonne AZ
      bb                      // Q: colonne BB
    ]);
  }

  // Ajuster le nombre de lignes si besoin
  var neededRows = 11 + rowsToWrite.length;
  var totalRows = sheetElec.getMaxRows();
  if (totalRows < neededRows) {
    sheetElec.insertRowsAfter(totalRows, neededRows - totalRows);
  }

  // Écriture en une fois
  if (rowsToWrite.length > 0) {
    sheetElec.getRange(12, 1, rowsToWrite.length, 17).setValues(rowsToWrite);
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - FUMISTERIE =====
var sheetFumisterie = ssSite.getSheetByName('Etat réglementaire - Fumisterie');
if (sheetFumisterie && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement Fumisterie pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');

  // Nettoyage de l’ancienne zone (A:N à partir de la ligne 12)
  var lastRow = sheetFumisterie.getLastRow();
  if (lastRow >= 12) {
    sheetFumisterie.getRange(12, 1, lastRow - 11, 14).clearContent(); // 14 colonnes = A..N
  }

  // Filtrer les lignes CVC à écrire: seulement si une des colonnes CD/CE/CF/CG/CH/CI/DL/DM/DN/DO/EO a au moins une valeur
  var rowsToWrite = [];
  var nbCVC = cvcParClientSite[keyEquip].length;

  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];

    // Colonnes demandées (index = lettreExcel-1)
    // F=5, G=6
    // CD=81, CE=82, CF=83, CG=84, CH=85, CI=86
    // DL=115, DM=116, DN=117, DO=118
    // EO=144
    var cd = dataCVC[idx][81];
    var ce = dataCVC[idx][82];
    var cf = dataCVC[idx][83];
    var cg = dataCVC[idx][84];
    var ch = dataCVC[idx][85];
    var ci = dataCVC[idx][86];

    var dl = dataCVC[idx][115];
    var dm = dataCVC[idx][116];
    var dn = dataCVC[idx][117];
    var dO = dataCVC[idx][118];

    var eo = dataCVC[idx][144];

    // condition: on écrit la ligne si au moins une des colonnes est renseignée
    var hasAny = (cd !== '' && cd != null) ||
                 (ce !== '' && ce != null) ||
                 (cf !== '' && cf != null) ||
                 (cg !== '' && cg != null) ||
                 (ch !== '' && ch != null) ||
                 (ci !== '' && ci != null) ||
                 (dl !== '' && dl != null) ||
                 (dm !== '' && dm != null) ||
                 (dn !== '' && dn != null) ||
                 (dO !== '' && dO != null) ||
                 (eo !== '' && eo != null);

    if (!hasAny) continue;

    rowsToWrite.push([
      rowsToWrite.length + 1, // A: numérotation 1..n (après filtrage)
      dataCVC[idx][5],        // B: colonne F
      dataCVC[idx][6],        // C: colonne G
      cd,                     // D: colonne CD
      ce,                     // E: colonne CE
      cf,                     // F: colonne CF
      cg,                     // G: colonne CG
      ch,                     // H: colonne CH
      ci,                     // I: colonne CI
      dl,                     // J: colonne DL
      dm,                     // K: colonne DM
      dn,                     // L: colonne DN
      dO,                     // M: colonne DO
      eo                      // N: colonne EO
    ]);
  }

  // Ajuster le nombre de lignes si besoin
  var neededRows = 11 + rowsToWrite.length;
  var totalRows = sheetFumisterie.getMaxRows();
  if (totalRows < neededRows) {
    sheetFumisterie.insertRowsAfter(totalRows, neededRows - totalRows);
  }

  // Écriture en une fois
  if (rowsToWrite.length > 0) {
    sheetFumisterie.getRange(12, 1, rowsToWrite.length, 14).setValues(rowsToWrite);
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - CONTRÔLES PÉRIODIQUES =====
var sheetControles = ssSite.getSheetByName('Etat réglementaire - Contrôles périodiques');
if (sheetControles && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement Contrôles périodiques pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');

  // Nettoyage de l’ancienne zone (A:G à partir de la ligne 12)
  var lastRow = sheetControles.getLastRow();
  if (lastRow >= 12) {
    sheetControles.getRange(12, 1, lastRow - 11, 7).clearContent(); // 7 colonnes = A..G
  }

  // Filtrer les lignes CVC à écrire: seulement si DP/DQ/DR/DS a au moins une valeur
  var rowsToWrite = [];
  var nbCVC = cvcParClientSite[keyEquip].length;

  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];

    // Colonnes demandées (index = lettreExcel-1)
    // F=5, G=6
    // DP=119, DQ=120, DR=121, DS=122
    var dp = dataCVC[idx][119];
    var dq = dataCVC[idx][120];
    var dr = dataCVC[idx][121];
    var ds = dataCVC[idx][122];

    // condition: on écrit la ligne si au moins une des 4 colonnes est renseignée
    var hasAny = (dp !== '' && dp != null) ||
                 (dq !== '' && dq != null) ||
                 (dr !== '' && dr != null) ||
                 (ds !== '' && ds != null);

    if (!hasAny) continue;

    rowsToWrite.push([
      rowsToWrite.length + 1, // A: numérotation 1..n (après filtrage)
      dataCVC[idx][5],        // B: colonne F
      dataCVC[idx][6],        // C: colonne G
      dp,                     // D: colonne DP
      dq,                     // E: colonne DQ
      dr,                     // F: colonne DR
      ds                      // G: colonne DS
    ]);
  }

  // Ajuster le nombre de lignes si besoin
  var neededRows = 11 + rowsToWrite.length;
  var totalRows = sheetControles.getMaxRows();
  if (totalRows < neededRows) {
    sheetControles.insertRowsAfter(totalRows, neededRows - totalRows);
  }

  // Écriture en une fois
  if (rowsToWrite.length > 0) {
    sheetControles.getRange(12, 1, rowsToWrite.length, 7).setValues(rowsToWrite);
  }
}
// ===== MISE À JOUR ETAT RÉGLEMENTAIRE - ECS =====
var sheetECS = ssSite.getSheetByName('Etat réglementaire - ECS');
if (sheetECS && cvcParClientSite[keyEquip]) {
  Logger.log('Traitement ECS pour: ' + keyEquip + ' - ' + cvcParClientSite[keyEquip].length + ' éléments');

  // Nettoyage de l’ancienne zone (A:S à partir de la ligne 12)
  var lastRow = sheetECS.getLastRow();
  if (lastRow >= 12) {
    sheetECS.getRange(12, 1, lastRow - 11, 19).clearContent(); // 19 colonnes = A..S
  }

  // Filtrer les lignes CVC à écrire: seulement si une des colonnes demandées a au moins une valeur
  var rowsToWrite = [];
  var nbCVC = cvcParClientSite[keyEquip].length;

  for (var k = 0; k < nbCVC; k++) {
    var idx = cvcParClientSite[keyEquip][k];

    // Colonnes demandées (index = lettreExcel-1)
    // F=5, G=6
    // DX=127
    // ES=148, ET=149, EU=150, EV=151, EW=152, EX=153, EY=154, EZ=155
    // FA=156, FB=157, FC=158, FD=159, FE=160, FF=161, FG=162
    var dx = dataCVC[idx][127];

    var es = dataCVC[idx][148];
    var et = dataCVC[idx][149];
    var eu = dataCVC[idx][150];
    var ev = dataCVC[idx][151];
    var ew = dataCVC[idx][152];
    var ex = dataCVC[idx][153];
    var ey = dataCVC[idx][154];
    var ez = dataCVC[idx][155];

    var fa = dataCVC[idx][156];
    var fb = dataCVC[idx][157];
    var fc = dataCVC[idx][158];
    var fd = dataCVC[idx][159];
    var fe = dataCVC[idx][160];
    var ff = dataCVC[idx][161];
    var fg = dataCVC[idx][162];

    // condition: on écrit la ligne si au moins une des colonnes est renseignée
    var hasAny = (dx !== '' && dx != null) ||
                 (fa !== '' && fa != null) ||
                 (fb !== '' && fb != null) ||
                 (ff !== '' && ff != null) ||
                 (fg !== '' && fg != null) ||
                 (fe !== '' && fe != null) ||
                 (es !== '' && es != null) ||
                 (et !== '' && et != null) ||
                 (eu !== '' && eu != null) ||
                 (ev !== '' && ev != null) ||
                 (ew !== '' && ew != null) ||
                 (ex !== '' && ex != null) ||
                 (ey !== '' && ey != null) ||
                 (ez !== '' && ez != null) ||
                 (fc !== '' && fc != null) ||
                 (fd !== '' && fd != null);

    if (!hasAny) continue;

    rowsToWrite.push([
      rowsToWrite.length + 1, // A: numérotation 1..n (après filtrage)
      dataCVC[idx][5],        // B: colonne F
      dataCVC[idx][6],        // C: colonne G
      dx,                     // D: colonne DX
      fa,                     // E: colonne FA
      fb,                     // F: colonne FB
      ff,                     // G: colonne FF
      fg,                     // H: colonne FG
      fe,                     // I: colonne FE
      es,                     // J: colonne ES
      et,                     // K: colonne ET
      eu,                     // L: colonne EU
      ev,                     // M: colonne EV
      ew,                     // N: colonne EW
      ex,                     // O: colonne EX
      ey,                     // P: colonne EY
      ez,                     // Q: colonne EZ
      fc,                     // R: colonne FC
      fd                      // S: colonne FD
    ]);
  }

  // Ajuster le nombre de lignes si besoin
  var neededRows = 11 + rowsToWrite.length;
  var totalRows = sheetECS.getMaxRows();
  if (totalRows < neededRows) {
    sheetECS.insertRowsAfter(totalRows, neededRows - totalRows);
  }

  // Écriture en une fois
  if (rowsToWrite.length > 0) {
    sheetECS.getRange(12, 1, rowsToWrite.length, 19).setValues(rowsToWrite);
  }
}

    fichierTraites++;
    
    // Log de progression toutes les 10 fichiers
    if (fichierTraites % 10 === 0) {
      Logger.log('Progression: ' + fichierTraites + ' fichiers traités sur ' + sitesATraiter.length);
    }
  }
  
  // ===== FINALISATION =====
  properties.deleteProperty('dernierSiteTraite');
  Logger.log('Script terminé avec succès. Total fichiers traités: ' + fichierTraites + '/' + sitesATraiter.length);
}

// ===== FONCTIONS UTILITAIRES =====
function verifierPresenceFichiers() {
  try {
    var folderRoot = DriveApp.getFolderById('13gsSiHdPyeaJmdAFivDNpGgKLChjGgED');
    var agencesFolders = folderRoot.getFolders();
    var compteurFichiers = 0;
    
    while (agencesFolders.hasNext() && compteurFichiers < 5) {
      var agenceFolder = agencesFolders.next();
      var clientsFolders = agenceFolder.getFolders();
      while (clientsFolders.hasNext() && compteurFichiers < 5) {
        var clientFolder = clientsFolders.next();
        var files = clientFolder.getFiles();
        if (files.hasNext()) {
          compteurFichiers++;
        }
      }
    }
    
    return compteurFichiers > 0;
  } catch (e) {
    return false;
  }
}

function resetProgression() {
  PropertiesService.getScriptProperties().deleteProperty('dernierSiteTraite');
  Logger.log('Progression réinitialisée');
}
