import Sheet = GoogleAppsScript.Spreadsheet.Sheet;

export class SheetManager {
    constructor(sheet: Sheet) {
        this.sheet = sheet;
    }

    public findCell(value: string): [number, number] {
        let rowIndex = -1;
        let columnIndex = -1;

        let rows = this.sheet.getDataRange().getValues();

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];

            for (let j = 0; j < row.length; j++) {
                const cell = row[j];

                if (cell === value) {
                    rowIndex = i;
                    columnIndex = j;
                    break;
                }
            }

            if (rowIndex !== -1) {
                break;
            }
        }

        return [rowIndex, columnIndex];
    }

    private sheet: Sheet;
}