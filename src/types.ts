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
    sheet_number: number;
    data: SpreadsheetData;
    created_at: string;
    updated_at: string;
    deleted_at: string | null;
}

export interface SpreadSheet {
    id: string;
    title: string;
    user_id: string;
    sheets: Sheet[];
    created_at: string;
    updated_at: string;
    deleted_at: string | null;
}