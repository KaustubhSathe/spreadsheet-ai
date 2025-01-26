declare module "https://deno.land/std@0.177.0/http/server.ts" {
  export function serve(handler: (req: Request) => Response | Promise<Response>): void;
}

declare module "https://esm.sh/@supabase/supabase-js@2.7.1" {
  export * from '@supabase/supabase-js';
}

declare module "https://deno.land/x/jose@v4.9.1/index.ts" {
  export function decodeJwt(token: string): { sub?: string; [key: string]: any };
} 