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
  // Handle CORS preflight requests
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    const authHeader = req.headers.get('Authorization');
    if (!authHeader?.startsWith('Bearer ')) {
      throw new Error('Invalid authorization header');
    }

    // Extract the JWT token
    const token = authHeader.replace('Bearer ', '');
    
    // Decode the JWT to get the user ID
    const decoded = jose.decodeJwt(token);
    const userId = decoded.sub;
    
    if (!userId) {
      throw new Error('Invalid token - no user ID');
    }

    // Create spreadsheet directly with the user ID from JWT
    const supabaseClient = createClient(supabaseUrl, supabaseAnonKey, {
      global: {
        headers: {
          Authorization: `Bearer ${token}`  // Pass the JWT token to authenticate the request
        }
      }
    });
    const { data, error } = await supabaseClient
      .from('spreadsheets')
      .insert([
        {
          title: 'Untitled spreadsheet',
          user_id: userId,
          data: {},
        },
      ])
      .select()
      .single();

    if (error) {
      console.error('Database error:', error);
      throw error;
    }

    return new Response(JSON.stringify(data), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 200,
    })
  } catch (error) {
    console.error('Error:', error.message);
    return new Response(JSON.stringify({ error: error.message }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 401,
    })
  }
})