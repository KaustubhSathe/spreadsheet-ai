import { createClient } from '@supabase/supabase-js';

export const supabaseUrl = 'https://lmrpcvsshrhlqeupdldg.supabase.co';
export const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImxtcnBjdnNzaHJobHFldXBkbGRnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Mzc5MDEzMjMsImV4cCI6MjA1MzQ3NzMyM30.MVdGDbhCqZW0pGc9SVsuZtmYFsgovRe_pim4eDyVaFY';

export const supabase = createClient(supabaseUrl, supabaseAnonKey, {
  auth: {
    autoRefreshToken: true,
    persistSession: true
  }
}); 