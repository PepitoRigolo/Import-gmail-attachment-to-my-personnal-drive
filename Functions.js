
/**
 * Représente la configuration d'un libellé
 * @typedef {Object} LibelleConfig
 * @property {string} libelle
 * @property {string} folderPath
 */

/**
 * Récupère les objects LibelleConfig depuis le sheet de configuration
 * @returns {LibelleConfig[]}
 */
function getLibelleConfigs() {
    let spreadSheet = getOrCreateSpreadSheet(config.configSheetPath, config.configSheetHeader);

    let sheet = spreadSheet.getSheets()[0];

    let values = sheet.getDataRange().getValues();

    if (values.length == 0 || values.length == 1 && values[0].length == 1 && values[0][0] == "") sheet.appendRow(config.configSheetHeader);

    return values
        .filter((row, i) => {
            return i > 0 && row.length > 0 && row[0] != "";
        })
        .map(row => {
            let libelle = row[0].trim();
            let folderPath = (row.length>1 && row[1] != "" ? row[1] : `/${row[0]}`);
            if (!folderPath.startsWith('/')) folderPath = '/' + folderPath;
            folderPath = folderPath.trim();
            return {
                libelle: libelle,
                folderPath: folderPath
            }
        });
}
/**
 * Ajoute une ligne dans le sheet d'historique
 * @param {string} idMessage 
 * @param {string} expediteur 
 * @param {string} nomPieceJointe 
 * @param {string} cheminSauvegarde 
 */
function appendHistory(idMessage, libelle, expediteur, nomPieceJointe, cheminSauvegarde)
{
    let spreadSheet = getOrCreateSpreadSheet(config.historiqueSheetPath, config.historiqueSheetHeader);

    let sheet = spreadSheet.getSheets()[0];

    let time = Utilities.formatDate(new Date(), "Europe/Paris", "yyyy-MM-dd'T'HH:mm:ss'Z'");

    sheet.appendRow([time, idMessage, libelle, expediteur, nomPieceJointe, cheminSauvegarde]);
}

/**
 * Retourne l'historique de tous les messages traités par le script pour cet utilisateur
 * @returns {string[]}
 */
function getHistory()
{
    let spreadSheet = getOrCreateSpreadSheet(config.historiqueSheetPath, config.historiqueSheetHeader);

    let sheet = spreadSheet.getSheets()[0];

    let values = sheet.getDataRange().getValues();

    if (values.length == 0 || values.length == 1 && values[0].length == 1 && values[0][0] == "") sheet.appendRow(config.historiqueSheetHeader);

    return values
        .filter((row, i) => {
            return i > 0 && row.length > 1 && row[1] != "";
        })
        .map(row => {
            return row[1];
        });
}

/**
 * Récupère ou créé un dossier à partir de son chemin
 * @param {string} path - arborescence du dossier séparé par des '/'
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getFolderFromPath(path)
{
    let strings = path.split("/");

    let currentFolder = DriveApp.getRootFolder();

    strings.forEach((string, index) => {
        if (string == "") return;
        currentFolder = getOrCreateFolder(string, currentFolder);
    });

    return currentFolder;
}

/**
 * Récupère ou créé un Spreadsheet sur le drive de l'utilisateur
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getOrCreateSpreadSheet(path, firstRow) {
    let strings = path.split("/");

    let currentFolder = DriveApp.getRootFolder();

    let spreadsheet = null;

    strings.forEach((string, index) => {
        if (string == "") return;
        if (index < strings.length - 1) {
            currentFolder = getOrCreateFolder(string, currentFolder);
        }
        else {
            let files = currentFolder.getFilesByName(string);

            if (files.hasNext()) spreadsheet = SpreadsheetApp.open(files.next());
            else spreadsheet = createSpreadSheet(currentFolder, string, firstRow);
        }
    });

    return spreadsheet;
}


/**
 * Récupère ou créé un dossier
 * @param {string} folderName - nom du dossier à trouver ou à créé
 * @param {GoogleAppsScript.Drive.Folder} parentFolder - dossier parent
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getOrCreateFolder(folderName, parentFolder) {
    let folders = parentFolder.getFoldersByName(folderName);

    if (folders.hasNext()) {
        return folders.next();
    }
    else {
        return parentFolder.createFolder(folderName);
    }
}

/**
 * Créé et initialise un spreadsheet
 * @param {GoogleAppsScript.Drive.Folder} parentFolder - dossier parent où créer le spreadsheet
 * @param {string} spreadSheetName - nom du spreadsheet à créer
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function createSpreadSheet(parentFolder, spreadSheetName, firstRow)
{
    let spreadSheet = SpreadsheetApp.create(spreadSheetName);

    DriveApp.getFileById(spreadSheet.getId()).moveTo(parentFolder);

    spreadSheet.getSheets()[0].appendRow(firstRow);

    return spreadSheet;
}
