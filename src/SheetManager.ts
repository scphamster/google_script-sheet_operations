import Sheet = GoogleAppsScript.Spreadsheet.Sheet;


export interface Cell {
    value: string | number | undefined;
    row: number;
    column: number;
}

export class SheetManager {
    constructor(sheet: Sheet) {
        this.sheet = sheet;
    }

    public findCellWithValue(value: string): Cell | undefined {
        let rows = this.sheet.getDataRange().getValues();

        for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
            const row = rows[rowIdx];

            for (let colIdx = 0; colIdx < row.length; colIdx++) {
                const cell = row[colIdx];

                if (typeof value == typeof cell && cell === value) {
                    return {value, row: rowIdx, column: colIdx};
                }
            }
        }

        return null;
    }

    private sheet: Sheet;
}