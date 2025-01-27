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
    data: SpreadsheetData;
    created_at: string;
    updated_at: string;
    deleted_at: string | null;
}

export interface Spreadsheet {
    id: string;
    title: string;
    user_id: string;
    sheets: Sheet[];
    created_at: string;
    updated_at: string;
    deleted_at: string | null;
}