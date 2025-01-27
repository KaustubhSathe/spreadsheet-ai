-- Add sheet_number column
ALTER TABLE sheets ADD COLUMN sheet_number INTEGER NOT NULL DEFAULT 1;

-- Add constraint to ensure sheet_number is unique within a spreadsheet
ALTER TABLE sheets 
ADD CONSTRAINT unique_sheet_number_per_spreadsheet 
UNIQUE (spreadsheet_id, sheet_number);

-- Update existing sheets to have sequential numbers
WITH numbered_sheets AS (
  SELECT id, spreadsheet_id, 
         ROW_NUMBER() OVER (PARTITION BY spreadsheet_id ORDER BY created_at) as rnum
  FROM sheets
  WHERE deleted_at IS NULL
)
UPDATE sheets
SET sheet_number = numbered_sheets.rnum
FROM numbered_sheets
WHERE sheets.id = numbered_sheets.id; 