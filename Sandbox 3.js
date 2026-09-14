/*************************************************************
 *  GÉNÉRATEUR DE SCHÉMA DE COMPTEURS (parent -> divisionnaires)
 *  Sortie : Google Sheet natif déposé dans un dossier Drive
 *************************************************************/

// === PARAMÈTRES À ADAPTER ===
const SOURCE_ID   = '1YPEO7-jRw5adU9OlBmeVH9DEZ6uVvYPmSJn9ZICa4pw'; // Sheet source (base de données)
const FOLDER_ID   = '16Q4yZo8Ws-aeTb2EvmmUCWTto4Hv6Emz';             // dossier de sortie
const SOURCE_NAME = 'Equipement';   // nom de l'onglet contenant tes données
                                    // (mets null pour utiliser le 1er onglet)

// Dimensions de la grille (px)
const CELL_W   = 26;   // largeur d'une cellule-unité
const CELL_H   = 20;   // hauteur d'une cellule-unité
const BOX_COLS = 8;    // largeur d'une boîte (en cellules)
const BOX_ROWS = 6;    // hauteur d'une boîte
const GAP_COLS = 2;    // espace horizontal entre 2 boîtes enfants

// Couleurs
const COUL_PARENT = '#1565C0';
const COUL_ENFANT = '#E3F2FD';
const COUL_TRAIT  = '#37474F';


/**
 * OUTIL : liste les onglets du fichier source (à lancer en cas de doute)
 */
function listerOnglets() {
  const ss = SpreadsheetApp.openById(SOURCE_ID);
  const noms = ss.getSheets().map(s => '«' + s.getName() + '»');
  Logger.log('Onglets trouvés : ' + noms.join(', '));
}


/**
 * FONCTION PRINCIPALE : à lancer
 */
function genererSchemaCompteurs() {
  const ss = SpreadsheetApp.openById(SOURCE_ID);
  const source = SOURCE_NAME ? ss.getSheetByName(SOURCE_NAME) : ss.getSheets()[0];
  if (!source) throw new Error('Onglet source introuvable : ' + SOURCE_NAME);

  const data = source.getDataRange().getValues();
  const headers = data.shift();

  // Repérage des colonnes (par nom d'en-tête)
  const c = {
    id:      idx(headers, 'ID'),
    codeSite:idx(headers, 'code_site'),
    site:    idx(headers, 'Nom site'),
    client:  idx(headers, 'nom_client'),
    designation: idx(headers, 'DESIGNATION'),
    nom:     idx(headers, 'Nom équipement'),
    type:    idx(headers, 'Type de compteur'),
    index:   idx(headers, 'Index'),
    unite:   idx(headers, 'Unité'),
    etage:   idx(headers, 'Localisation (étage)'),
    local:   idx(headers, 'Localisation (local)')
  };

  // On ne garde que les compteurs
  const compteurs = data.filter(r =>
    String(r[c.designation]).toUpperCase().indexOf('COMPTR') > -1 ||
    String(r[c.type]).trim() !== ''
  );
  if (!compteurs.length) throw new Error('Aucun compteur trouvé.');

  // Regroupement par site
  const parSite = {};
  compteurs.forEach(r => {
    const cle = r[c.codeSite] + ' | ' + r[c.site];
    (parSite[cle] = parSite[cle] || []).push(r);
  });

  // Création du fichier de sortie
  const nomFichier = 'Schéma compteurs - ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const outSs = SpreadsheetApp.create(nomFichier);

  // On dessine un onglet par site
  let premier = true;
  Object.keys(parSite).forEach(cle => {
    const lignes = parSite[cle];
    const nomSite = String(lignes[0][c.site]).substring(0, 90) || 'Site';
    const sheet = premier
      ? outSs.getActiveSheet().setName(nettoyerNom(nomSite))
      : outSs.insertSheet(nettoyerNom(nomSite));
    premier = false;
    dessinerSite(sheet, lignes, c);
  });

  // Déplacement du fichier dans le bon dossier Drive
  const file = DriveApp.getFileById(outSs.getId());
  DriveApp.getFolderById(FOLDER_ID).addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  // Log (fonctionne même en script standalone)
  Logger.log('✅ Schéma généré !');
  Logger.log('Fichier : ' + nomFichier);
  Logger.log('URL : ' + outSs.getUrl());
}


/**
 * Dessine le schéma d'un site sur une feuille
 */
function dessinerSite(sheet, lignes, c) {
  // Grille
  sheet.clear();
  const totalCols = Math.max(30, lignes.length * (BOX_COLS + GAP_COLS) + 4);
  const totalRows = 40;

  // On crée les colonnes/lignes manquantes AVANT de dessiner
  const maxCols = sheet.getMaxColumns();
  if (totalCols > maxCols) sheet.insertColumnsAfter(maxCols, totalCols - maxCols);
  const maxRows = sheet.getMaxRows();
  if (totalRows > maxRows) sheet.insertRowsAfter(maxRows, totalRows - maxRows);

  for (let col = 1; col <= totalCols; col++) sheet.setColumnWidth(col, CELL_W);
  for (let row = 1; row <= totalRows; row++) sheet.setRowHeight(row, CELL_H);
  sheet.setHiddenGridlines(true);

  // Titre
  const titre = sheet.getRange(1, 1, 1, totalCols).merge();
  titre.setValue('SCHÉMA COMPTEURS — ' + lignes[0][c.site]);
  titre.setFontWeight('bold').setFontSize(12)
       .setHorizontalAlignment('left').setBackground('#ECEFF1');

  // Parent = 1re ligne, enfants = le reste
  const parent = lignes[0];
  const enfants = lignes.slice(1);

  // Positions verticales
  const parentRowStart = 3;
  const parentRowEnd   = parentRowStart + BOX_ROWS - 1;   // 8
  const busRow         = parentRowEnd + 2;                // 10
  const childRowStart  = busRow + 2;                      // 12

  // Positions horizontales des enfants
  const step = BOX_COLS + GAP_COLS;
  const nbEnfants = enfants.length;
  const largeurBloc = nbEnfants > 0
    ? nbEnfants * BOX_COLS + (nbEnfants - 1) * GAP_COLS
    : BOX_COLS;
  const debutBloc = 2;

  // Boîte parent (centrée au-dessus du bloc enfants)
  const parentStartCol = Math.max(2, Math.round(debutBloc + largeurBloc / 2 - BOX_COLS / 2));
  dessinerBoite(sheet, parentRowStart, parentStartCol, texteBoite(parent, c, true),
                COUL_PARENT, '#FFFFFF');

  if (nbEnfants === 0) return;

  const parentBoundary = parentStartCol + Math.round(BOX_COLS / 2);

  // Boîtes enfants + connecteurs
  const boundaries = [];
  enfants.forEach((e, i) => {
    const colStart = debutBloc + i * step;
    dessinerBoite(sheet, childRowStart, colStart, texteBoite(e, c, false),
                  COUL_ENFANT, '#0D47A1');
    const b = colStart + Math.round(BOX_COLS / 2);
    boundaries.push(b);
    // trait vertical du bus -> enfant
    ligneVerticale(sheet, b, busRow, childRowStart - 1);
  });

  // Trait vertical parent -> bus
  ligneVerticale(sheet, parentBoundary, parentRowEnd + 1, busRow);

  // Trait horizontal (bus)
  const minB = Math.min(parentBoundary, ...boundaries);
  const maxB = Math.max(parentBoundary, ...boundaries);
  ligneHorizontale(sheet, busRow, minB, maxB);
}


/* ---------- Fonctions utilitaires ---------- */

function texteBoite(r, c, estParent) {
  const nom   = r[c.nom]   || '(sans nom)';
  const type  = r[c.type]  || '—';
  const index = (r[c.index] !== '' ? r[c.index] : '—');
  const unite = r[c.unite] || '';
  const loc   = [r[c.etage], r[c.local]].filter(String).join(' - ') || '—';
  return (estParent ? '★ GÉNÉRAL\n' : '') +
    nom + '\n' +
    'Type : ' + type + '\n' +
    'Index : ' + index + (unite ? ' ' + unite : '') + '\n' +
    '📍 ' + loc;
}

function dessinerBoite(sheet, r, col, texte, fond, couleurTexte) {
  const rng = sheet.getRange(r, col, BOX_ROWS, BOX_COLS).merge();
  rng.setValue(texte)
     .setBackground(fond)
     .setFontColor(couleurTexte)
     .setFontSize(9)
     .setWrap(true)
     .setVerticalAlignment('middle')
     .setHorizontalAlignment('center')
     .setBorder(true, true, true, true, false, false,
                COUL_TRAIT, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function ligneVerticale(sheet, col, rDebut, rFin) {
  sheet.getRange(rDebut, col, rFin - rDebut + 1, 1)
       .setBorder(null, true, null, null, null, null,
                  COUL_TRAIT, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function ligneHorizontale(sheet, row, cDebut, cFin) {
  sheet.getRange(row, cDebut, 1, cFin - cDebut + 1)
       .setBorder(null, null, true, null, null, null,
                  COUL_TRAIT, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function idx(headers, nom) {
  const i = headers.indexOf(nom);
  if (i === -1) Logger.log('⚠️ Colonne introuvable : ' + nom);
  return i;
}

function nettoyerNom(nom) {
  return String(nom).replace(/[\\/?*\[\]:]/g, ' ').substring(0, 95).trim() || 'Site';
}
