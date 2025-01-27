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
    title: string;
    data: SpreadsheetData;
    created_at: string;
    user_id: string;
    updated_at: string;
    deleted_at: string | null;
}