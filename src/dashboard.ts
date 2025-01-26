import { supabase, supabaseAnonKey } from './supabase';

interface Spreadsheet {
  id: string;
  title: string;
  created_at: string;
  user_id: string;
}

async function checkAuth() {
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) {
    window.location.href = '/';
    return false;
  }
  return true;
}

async function loadSpreadsheets() {
  const { data: spreadsheets, error } = await supabase
    .from('spreadsheets')
    .select('*')
    .order('created_at', { ascending: false });

  if (error) {
    console.error('Error loading spreadsheets:', error);
    return;
  }

  const list = document.getElementById('spreadsheets-list');
  if (!list) return;

  list.innerHTML = spreadsheets.map((sheet: Spreadsheet) => `
    <div class="spreadsheet-card" onclick="window.location.href='/app?id=${sheet.id}'">
      <div class="spreadsheet-title">${sheet.title}</div>
      <div class="spreadsheet-date">
        ${new Date(sheet.created_at).toLocaleDateString()}
      </div>
    </div>
  `).join('');
}

// Initialize
document.addEventListener('DOMContentLoaded', async () => {
  if (await checkAuth()) {
    loadSpreadsheets();
    
    // Add new spreadsheet handler
    const newButton = document.querySelector('.new-sheet-btn');
    if (newButton) {
      newButton.addEventListener('click', async () => {
        try {
          // Add loading state
          newButton.classList.add('loading');
          
          const { data: { session }, error: sessionError } = await supabase.auth.getSession();
          if (sessionError || !session) {
            throw new Error('No active session');
          }
          
          console.log('Session:', session); // Debug log
          
          const { data, error } = await supabase.functions.invoke('create-spreadsheet', {
            method: 'POST',
            headers: {
              Authorization: `Bearer ${session.access_token}`,
              'apikey': supabaseAnonKey,
            }
          });

          if (error) throw error;

          window.location.href = `/app?id=${data.id}`;
        } catch (error) {
          console.error('Error:', error);
          if (error.message === 'No active session') {
            window.location.href = '/';
          } else {
            alert('Failed to create spreadsheet. Please try again.');
          }
        } finally {
          newButton.classList.remove('loading');
        }
      });
    }

    // Add logout handler
    const logoutButton = document.querySelector('.logout-button');
    if (logoutButton) {
      logoutButton.addEventListener('click', async () => {
        await supabase.auth.signOut();
        window.location.href = '/';
      });
    }
  }
}); 