/**
 * Dictionnaire de Données DSN - Implémentation Google Sheets
 * 
 * Ce script est destiné à être utilisé directement dans Google Sheets via Apps Script.
 * Il crée un dictionnaire de données DSN à partir des onglets importés.
 * 
 * Instructions d'utilisation:
 * 1. Importez vos fichiers Excel dans Google Sheets (en créant un onglet par feuille Excel)
 * 2. Copiez ce code dans l'éditeur de script (Extensions > Apps Script)
 * 3. Exécutez la fonction onOpen() pour ajouter le menu, puis utilisez le menu pour générer le dictionnaire
 */

/**
 * Crée un menu personnalisé lors de l'ouverture du document
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('DSN Dictionary')
    .addItem('Importer les fichiers XLSX', 'showImportDialog')
    .addItem('Générer le dictionnaire', 'generateDictionary')
    .addToUi();
}

/**
 * Affiche une boîte de dialogue pour l'importation
 */
function showImportDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Importation des fichiers XLSX',
    'Pour importer les fichiers DSN, veuillez:\n\n' +
    '1. Télécharger le fichier XLSX dans Google Drive\n' +
    '2. Ouvrir le fichier avec Google Sheets\n' +
    '3. Copier chaque onglet dans ce document\n\n' +
    'Les onglets suivants sont nécessaires:\n' +
    '- "1 - Tableau des usages" (du fichier dsntableaudesusagesCT2025.1.xlsx)\n' +
    '- "Data Types", "Fields", "Blocks" (du fichier dsndatatypesCT2025.xlsx)\n\n' +
    'Avez-vous déjà importé ces onglets?',
    ui.ButtonSet.YES_NO
  );
  
  if (result == ui.Button.YES) {
    checkRequiredSheets();
  } else {
    ui.alert('Veuillez importer les onglets nécessaires avant de générer le dictionnaire.');
  }
}

/**
 * Vérifie que les onglets requis sont présents
 */
function checkRequiredSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [
    '1 - Tableau des usages',
    'Data Types',
    'Fields',
    'Blocks'
  ];
  
  const missingSheets = [];
  
  requiredSheets.forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      missingSheets.push(sheetName);
    }
  });
  
  const ui = SpreadsheetApp.getUi();
  
  if (missingSheets.length > 0) {
    ui.alert(
      'Onglets manquants',
      'Les onglets suivants sont manquants:\n- ' + missingSheets.join('\n- ') +
      '\n\nVeuillez les importer avant de générer le dictionnaire.',
      ui.ButtonSet.OK
    );
    return false;
  }
  
  ui.alert('Tous les onglets nécessaires sont présents. Vous pouvez générer le dictionnaire.');
  return true;
}

/**
 * Fonction principale pour générer le dictionnaire de données
 */
function generateDictionary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Vérifier que tous les onglets nécessaires sont présents
  if (!checkRequiredSheets()) {
    return;
  }
  
  // Création ou récupération des onglets nécessaires
  let dictionarySheet = ss.getSheetByName('Dictionnaire');
  if (!dictionarySheet) {
    dictionarySheet = ss.insertSheet('Dictionnaire');
  } else {
    dictionarySheet.clear();
  }
  
  let nomenclaturesSheet = ss.getSheetByName('Nomenclatures');
  if (!nomenclaturesSheet) {
    nomenclaturesSheet = ss.insertSheet('Nomenclatures');
  } else {
    nomenclaturesSheet.clear();
  }
  
  // Récupération des données des onglets source
  const fieldsData = getFieldsData();
  const dataTypes = getDataTypes();
  const blocksData = getBlocksData();
  const usageData = getUsageData();
  const nomenclatures = getNomenclatures(dataTypes);
  
  // Préparation et remplissage du dictionnaire
  populateDictionarySheet(dictionarySheet, fieldsData, dataTypes, blocksData, usageData, nomenclatures);
  
  // Préparation et remplissage des nomenclatures
  populateNomenclaturesSheet(nomenclaturesSheet, nomenclatures);
  
  // Formatage des onglets
  formatDictionarySheet(dictionarySheet);
  formatNomenclaturesSheet(nomenclaturesSheet);
  
  ui.alert('Dictionnaire de données DSN généré avec succès!');
}

/**
 * Récupère les données des rubriques (fields)
 */
function getFieldsData() {
  const fieldsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Fields');
  if (!fieldsSheet) {
    throw new Error('Onglet "Fields" introuvable');
  }
  
  const data = fieldsSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Conversion en tableau d'objets
  return data.slice(1).map(row => {
    const field = {};
    headers.forEach((header, index) => {
      field[header] = row[index];
    });
    return field;
  });
}

/**
 * Récupère les données des types de données
 */
function getDataTypes() {
  const dataTypesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Types');
  if (!dataTypesSheet) {
    throw new Error('Onglet "Data Types" introuvable');
  }
  
  const data = dataTypesSheet.getDataRange().getValues();
  const headers = data[0];
  
  return data.slice(1).map(row => {
    const type = {};
    headers.forEach((header, index) => {
      type[header] = row[index];
    });
    return type;
  });
}

/**
 * Récupère les données des blocs
 */
function getBlocksData() {
  const blocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Blocks');
  if (!blocksSheet) {
    throw new Error('Onglet "Blocks" introuvable');
  }
  
  const data = blocksSheet.getDataRange().getValues();
  const headers = data[0];
  
  return data.slice(1).map(row => {
    const block = {};
    headers.forEach((header, index) => {
      block[header] = row[index];
    });
    return block;
  });
}

/**
 * Récupère les données d'usage (obligatoire, conditionnel, etc.)
 */
function getUsageData() {
  const usageSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('1 - Tableau des usages');
  if (!usageSheet) {
    throw new Error('Onglet "1 - Tableau des usages" introuvable');
  }
  
  const data = usageSheet.getDataRange().getValues();
  
  // Le format est spécifique au tableau des usages
  const declarationTypes = data[1].slice(4); // Types de déclaration (en-têtes)
  
  const usageMap = {};
  for (let i = 3; i < data.length; i++) {
    const row = data[i];
    if (row[2]) { // Si la ligne contient une rubrique
      const rubrique = row[2];
      const obligations = row.slice(4).map(v => v || '');
      
      usageMap[rubrique] = {
        isObligatory: obligations.some(o => o === 'O'),
        isConditional: obligations.some(o => o === 'C'),
        details: {}
      };
      
      // Associer chaque obligation à son type de déclaration
      for (let j = 0; j < obligations.length && j < declarationTypes.length; j++) {
        if (declarationTypes[j]) {
          usageMap[rubrique].details[declarationTypes[j]] = obligations[j];
        }
      }
    }
  }
  
  return usageMap;
}

/**
 * Extrait les nomenclatures des types de données
 */
function getNomenclatures(dataTypes) {
  const nomenclatures = {};
  
  dataTypes.forEach(type => {
    if (type.Values && typeof type.Values === 'string' && type.Values.includes('=')) {
      const values = type.Values.split(';').map(val => {
        const parts = val.split('=');
        return {
          code: parts[0],
          libelle: parts[1] || '',
          commentaire: ''
        };
      });
      
      nomenclatures[type.Id] = {
        id: type.Id,
        name: type.Name || type.Id,
        values: values
      };
    }
  });
  
  return nomenclatures;
}

/**
 * Remplit l'onglet du dictionnaire avec les données
 */
function populateDictionarySheet(sheet, fields, dataTypes, blocks, usageMap, nomenclatures) {
  // En-têtes du dictionnaire
  const headers = [
    'Bloc', 'Rubrique', 'Nom du champ', 'Type', 'Longueur',
    'Obligatoire', 'Nomenclature', 'Description', 'Commentaire'
  ];
  
  // Ajouter les en-têtes
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Créer un mapping des blocs pour accéder facilement à leurs descriptions
  const blockMap = {};
  blocks.forEach(block => {
    blockMap[block.Id] = block.Name || '';
  });
  
  // Préparer les données
  const data = fields.map(field => {
    // Trouver le type de données
    const dataType = dataTypes.find(type => type.Id === field['DataType Id']);
    
    // Construire l'identifiant complet de la rubrique
    const fieldId = `${field['Block Id']}.${field.Id}`;
    
    // Récupérer les informations d'usage
    const usage = usageMap[fieldId] || { isObligatory: false, isConditional: false };
    
    // Déterminer si la rubrique a une nomenclature
    const hasNomenclature = dataType && nomenclatures[dataType.Id];
    
    return [
      field['Block Id'] + ' - ' + (blockMap[field['Block Id']] || ''),
      fieldId,
      field.Name,
      dataType ? dataType.Nature : '',
      dataType ? `${dataType['Lg Min'] || ''}-${dataType['Lg Max'] || ''}` : '',
      usage.isObligatory ? 'Oui' : (usage.isConditional ? 'Conditionnel' : 'Non'),
      hasNomenclature ? dataType.Id : '',
      field.Description || '',
      field.Comment || ''
    ];
  });
  
  // Ajouter les données
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
}

/**
 * Remplit l'onglet des nomenclatures avec les codes et leur signification
 */
function populateNomenclaturesSheet(sheet, nomenclatures) {
  // En-têtes pour les nomenclatures
  const headers = ['Nomenclature', 'Code', 'Libellé', 'Commentaire'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Préparer les données de toutes les nomenclatures
  let allValues = [];
  Object.values(nomenclatures).forEach(nomenclature => {
    const values = nomenclature.values.map(value => [
      nomenclature.id,
      value.code,
      value.libelle,
      value.commentaire || ''
    ]);
    allValues = allValues.concat(values);
  });
  
  // Ajouter les données
  if (allValues.length > 0) {
    sheet.getRange(2, 1, allValues.length, headers.length).setValues(allValues);
  }
}

/**
 * Formate l'onglet du dictionnaire pour une meilleure lisibilité
 */
function formatDictionarySheet(sheet) {
  // Figer la première ligne
  sheet.setFrozenRows(1);
  
  // Mettre en forme les en-têtes
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  // Ajuster la largeur des colonnes
  sheet.setColumnWidth(1, 200); // Bloc
  sheet.setColumnWidth(2, 150); // Rubrique
  sheet.setColumnWidth(3, 250); // Nom du champ
  sheet.setColumnWidth(4, 100); // Type
  sheet.setColumnWidth(5, 100); // Longueur
  sheet.setColumnWidth(6, 100); // Obligatoire
  sheet.setColumnWidth(7, 150); // Nomenclature
  sheet.setColumnWidth(8, 400); // Description
  sheet.setColumnWidth(9, 200); // Commentaire
  
  // Créer un filtre sur les en-têtes
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).createFilter();
  
  // Mettre en forme les données
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  dataRange.setVerticalAlignment('top');
  dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // Mettre en surbrillance les champs obligatoires
  const obligatoireColumn = 6;
  const obligatoireRange = sheet.getRange(2, obligatoireColumn, sheet.getLastRow() - 1, 1);
  
  // Appliquer des formats conditionnels
  const obligatoireRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Oui')
    .setBackground('#e6f4ea')
    .setRanges([obligatoireRange])
    .build();
  
  const conditionnelRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Conditionnel')
    .setBackground('#fff0e6')
    .setRanges([obligatoireRange])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(obligatoireRule, conditionnelRule);
  sheet.setConditionalFormatRules(rules);
}

/**
 * Formate l'onglet des nomenclatures pour une meilleure lisibilité
 */
function formatNomenclaturesSheet(sheet) {
  // Figer la première ligne
  sheet.setFrozenRows(1);
  
  // Mettre en forme les en-têtes
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  // Ajuster la largeur des colonnes
  sheet.setColumnWidth(1, 200); // Nomenclature
  sheet.setColumnWidth(2, 100); // Code
  sheet.setColumnWidth(3, 350); // Libellé
  sheet.setColumnWidth(4, 250); // Commentaire
  
  // Créer un filtre sur les en-têtes
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).createFilter();
  
  // Mettre en forme l'alternance de couleurs
  // Utiliser la mise en forme conditionnelle pour alterner les couleurs selon la nomenclature
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  dataRange.setVerticalAlignment('top');
  dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // Créer une mise en forme alternée
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
}

/**
 * Crée une feuille avec les instructions d'utilisation
 */
function createInstructionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let instructionsSheet = ss.getSheetByName('Instructions');
  
  if (!instructionsSheet) {
    instructionsSheet = ss.insertSheet('Instructions');
  } else {
    instructionsSheet.clear();
  }
  
  const instructions = [
    ['Dictionnaire de Données DSN - Instructions d\'utilisation'],
    [''],
    ['Ce fichier contient un dictionnaire complet des données de la Déclaration Sociale Nominative (DSN).'],
    [''],
    ['Onglets disponibles:'],
    ['1. Dictionnaire - Contient toutes les rubriques DSN avec leurs caractéristiques'],
    ['2. Nomenclatures - Contient les listes de codes et leur signification'],
    [''],
    ['Utilisation du dictionnaire:'],
    ['- Le champ "Obligatoire" indique si la rubrique est obligatoire (Oui), conditionnelle (Conditionnel) ou facultative (Non)'],
    ['- Le champ "Nomenclature" fait référence à une liste de codes disponible dans l\'onglet "Nomenclatures"'],
    ['- Utilisez les filtres (en-tête des colonnes) pour rechercher rapidement des rubriques spécifiques'],
    [''],
    ['Utilisation pour le parsing de DSN:'],
    ['- Pour chaque bloc/rubrique dans le fichier DSN, vous pouvez retrouver ses caractéristiques dans le dictionnaire'],
    ['- Pour les rubriques associées à une nomenclature, vous pouvez traduire les codes en libellés plus compréhensibles'],
    [''],
    ['Menu personnalisé:'],
    ['- Utilisez le menu "DSN Dictionary" pour mettre à jour le dictionnaire ou importer de nouveaux fichiers']
  ];
  
  instructionsSheet.getRange(1, 1, instructions.length, 1).setValues(instructions);
  
  // Mise en forme
  instructionsSheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  instructionsSheet.getRange(5, 1).setFontWeight('bold');
  instructionsSheet.getRange(9, 1).setFontWeight('bold');
  instructionsSheet.getRange(14, 1).setFontWeight('bold');
  instructionsSheet.getRange(18, 1).setFontWeight('bold');
  
  instructionsSheet.setColumnWidth(1, 600);
}

/**
 * Point d'entrée principal
 */
function main() {
  createInstructionsSheet();
  generateDictionary();
}
