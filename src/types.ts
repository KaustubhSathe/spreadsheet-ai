export interface Cell {
    value: string;
    formula: string;
    computed: string;
}
  
export interface SpreadsheetData {
    [key: string]: Cell;
}

export interface Sheet {
    id: string;
    name: string;
    data: SpreadsheetData;
}