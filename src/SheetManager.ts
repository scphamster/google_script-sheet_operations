import Sheet = GoogleAppsScript.Spreadsheet.Sheet;


export interface Cell {
    value: string | number | undefined;
    position: CellId;
}

export interface CellId {
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
                    return {value, position: {row: rowIdx, column: colIdx}};
                }
            }
        }

        return null;
    }

    public findCellWithValueInRow(value: any, rowIdx: number): Cell | undefined {
        let rows = this.sheet.getDataRange().getValues();
        if (rows.length <= rowIdx) return null;

        let row = rows[rowIdx];

        for (let colIdx = 0; colIdx < rows.length; colIdx++) {
            let cell = row[colIdx];

            if (cell == value) return {value: value, position: {row: rowIdx, column:colIdx}};
        }

        return null;
    }

    public findCellWithValueInColumn(value: any, column: number): Cell | undefined {
        let rows = this.sheet.getDataRange().getValues();

        for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
            let row = rows[rowIdx];

            if (row.length <= column) return null;

            if (row[column] == value) return {value, position: {row: rowIdx, column}};
        }

        return null;
    }

    public getCellValue(pos: CellId) {
        return this.sheet.getDataRange().getCell(pos.row, pos.column).getValue();
    }

    private sheet: Sheet;
}