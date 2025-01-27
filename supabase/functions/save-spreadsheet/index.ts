import { serve } from "https://deno.land/std@0.177.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2.7.1";
import * as jose from "https://deno.land/x/jose@v4.9.1/index.ts";

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

    const { spreadsheetId, title, sheetId, data } = await req.json();
    
    if (!spreadsheetId || !sheetId) {
      throw new Error('Spreadsheet ID and Sheet ID are required');
    }

    const supabaseClient = createClient(supabaseUrl, supabaseAnonKey, {
      global: {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    });

    // Update spreadsheet title if provided
    if (title) {
      const { error: spreadsheetError } = await supabaseClient
        .from('spreadsheets')
        .update({
          title,
          updated_at: new Date().toISOString()
        })
        .eq('id', spreadsheetId)
        .eq('user_id', userId);

      if (spreadsheetError) throw spreadsheetError;
    }

    // Update sheet data
    const { error: sheetError } = await supabaseClient
      .from('sheets')
      .update({
        data,
        updated_at: new Date().toISOString()
      })
      .eq('id', sheetId);

    if (sheetError) throw sheetError;

    return new Response(JSON.stringify({ success: true }), {
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