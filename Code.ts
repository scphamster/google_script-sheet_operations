import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;
import {getValueFromCell} from "./src/Helpers";
import {SheetManager} from "./src/SheetManager";

function findSheet(spreadsheetName: string): Spreadsheet {
    let fileType = 'application/vnd.google-apps.spreadsheet';
    let query = 'title = "' + spreadsheetName + '" and mimeType = "' + fileType + '" and trashed = false';

    let foundFiles = DriveApp.searchFiles(query);

    if (!foundFiles.hasNext()) {
        throw Error("File with given name was not found!");
    }

    return SpreadsheetApp.openById(foundFiles.next().getId());
}

function findStringInSheet(spreadsheet: Spreadsheet, sheetName: string, needle: string) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
        Logger.log("sheet was found");
        var rangeData = sheet.getDataRange();
        var data = rangeData.getValues();

        for (var i = 0; i < data.length; i++) {
            for (var j = 0; j < data[i].length; j++) {
                if (data[i][j] == needle) {
                    Logger.log("Found: row=" + (i + 1) + " column=" + (j + 1));  // "+1" is added because spreadsheet rows & columns are 1-indexed
                }
            }
        }
    }
}

function main() {
    const sourceFileName = 'src';
    let sourceSprSheet = findSheet(sourceFileName);

    Logger.log("file was found! " + sourceSprSheet);

    findStringInSheet(sourceSprSheet, "sheet1", "somestring");

    const sheetName = "sheet1";

    let sheetManager = new SheetManager(sourceSprSheet.getSheetByName(sheetName));

    const cellOfInterest = "Final delivery date";
    let [rowId, columnId] = sheetManager.findCell(cellOfInterest);

    if (rowId != -1 && columnId != -1) {
        Logger.log(`Found a cell at ${rowId}:${columnId}`);
    }
    // getValueForRowAndColumnInSheet()
}