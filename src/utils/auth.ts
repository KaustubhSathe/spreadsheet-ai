import { supabase } from '../supabase';

export async function checkAuth() {
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) {
    window.location.href = '/';
    return false;
  }
  return true;
}

export async function handleLogout() {
  const { error } = await supabase.auth.signOut();
  if (!error) {
    window.location.href = '/';
  }
} 