import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2.21.0";
import * as jose from "https://deno.land/x/jose@v4.13.1/index.ts";

const supabaseUrl = 'https://lmrpcvsshrhlqeupdldg.supabase.co';
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImxtcnBjdnNzaHJobHFldXBkbGRnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Mzc5MDEzMjMsImV4cCI6MjA1MzQ3NzMyM30.MVdGDbhCqZW0pGc9SVsuZtmYFsgovRe_pim4eDyVaFY';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

serve(async (req: Request) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    const authHeader = req.headers.get('Authorization');
    if (!authHeader?.startsWith('Bearer ')) {
      throw new Error('Invalid authorization header');
    }

    const token = authHeader.replace('Bearer ', '');
    const decoded = jose.decodeJwt(token);
    const userId = decoded.sub;
    
    if (!userId) {
      throw new Error('Invalid token - no user ID');
    }

    const body = await req.json();
    const { spreadsheetId } = body;
    
    if (!spreadsheetId) {
      throw new Error('Spreadsheet ID is required');
    }

    const supabaseClient = createClient(supabaseUrl, supabaseAnonKey, {
      global: {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    });

    // Verify spreadsheet exists and belongs to user
    const { data: spreadsheet, error: verifyError } = await supabaseClient
      .from('spreadsheets')
      .select('id')
      .eq('id', spreadsheetId)
      .eq('user_id', userId)
      .single();

    if (verifyError || !spreadsheet) {
      throw new Error('Spreadsheet not found or access denied');
    }

    // Inside the create-sheet function, get the next sheet number
    const { data: lastSheet, error: countError } = await supabaseClient
      .from('sheets')
      .select('sheet_number')
      .eq('spreadsheet_id', spreadsheetId)
      .order('sheet_number', { ascending: false })
      .limit(1)
      .single();

    const nextSheetNumber = (lastSheet?.sheet_number || 0) + 1;

    // Create new sheet with sheet_number
    const { data: sheet, error } = await supabaseClient
      .from('sheets')
      .insert([{
        spreadsheet_id: spreadsheetId,
        data: {},
        sheet_number: nextSheetNumber
      }])
      .select()
      .single();

    if (error) throw error;

    return new Response(JSON.stringify(sheet), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 200,
    });
  } catch (error) {
    return new Response(JSON.stringify({ error: error.message }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 401,
    });
  }
}); 