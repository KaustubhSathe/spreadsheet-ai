-- Create spreadsheets table
CREATE TABLE spreadsheets (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  title TEXT NOT NULL DEFAULT 'Untitled spreadsheet',
  user_id UUID NOT NULL REFERENCES auth.users(id),
  created_at TIMESTAMP WITH TIME ZONE DEFAULT TIMEZONE('utc'::text, NOW()) NOT NULL,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT TIMEZONE('utc'::text, NOW()) NOT NULL,
  deleted_at TIMESTAMP WITH TIME ZONE
);

-- Create sheets table
CREATE TABLE sheets (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  spreadsheet_id UUID NOT NULL REFERENCES spreadsheets(id) ON DELETE CASCADE,
  data JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT TIMEZONE('utc'::text, NOW()) NOT NULL,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT TIMEZONE('utc'::text, NOW()) NOT NULL,
  deleted_at TIMESTAMP WITH TIME ZONE
);

-- Add RLS policies
ALTER TABLE spreadsheets ENABLE ROW LEVEL SECURITY;
ALTER TABLE sheets ENABLE ROW LEVEL SECURITY;

-- Policies for spreadsheets
CREATE POLICY "Users can view their own spreadsheets" ON spreadsheets
  FOR SELECT USING (auth.uid() = user_id);

CREATE POLICY "Users can create their own spreadsheets" ON spreadsheets
  FOR INSERT WITH CHECK (auth.uid() = user_id);

CREATE POLICY "Users can update their own spreadsheets" ON spreadsheets
  FOR UPDATE USING (auth.uid() = user_id);

-- Policies for sheets
CREATE POLICY "Users can view sheets of their spreadsheets" ON sheets
  FOR SELECT USING (
    EXISTS (
      SELECT 1 FROM spreadsheets 
      WHERE spreadsheets.id = sheets.spreadsheet_id 
      AND spreadsheets.user_id = auth.uid()
    )
  );

CREATE POLICY "Users can create sheets in their spreadsheets" ON sheets
  FOR INSERT WITH CHECK (
    EXISTS (
      SELECT 1 FROM spreadsheets 
      WHERE spreadsheets.id = sheets.spreadsheet_id 
      AND spreadsheets.user_id = auth.uid()
    )
  );

CREATE POLICY "Users can update sheets in their spreadsheets" ON sheets
  FOR UPDATE USING (
    EXISTS (
      SELECT 1 FROM spreadsheets 
      WHERE spreadsheets.id = sheets.spreadsheet_id 
      AND spreadsheets.user_id = auth.uid()
    )
  ); 