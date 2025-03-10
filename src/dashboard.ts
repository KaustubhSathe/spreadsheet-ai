import { supabase, supabaseAnonKey } from './supabase';
import { SpreadSheet } from './types';
import { checkAuth, handleLogout } from './utils/auth';

interface Spreadsheet {
  id: string;
  title: string;
  created_at: string;
  user_id: string;
}

async function loadSpreadsheets() {
  try {
    const { data: { session }, error: sessionError } = await supabase.auth.getSession();
    if (sessionError || !session) {
      throw new Error('No active session');
    }

    const { data: spreadsheets, error } = await supabase.functions.invoke('get-spreadsheet', {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${session.access_token}`,
        'apikey': supabaseAnonKey,
      }
    }) as { data: SpreadSheet[], error: any };

    console.log(spreadsheets);

    if (error) throw error;

    const list = document.getElementById('spreadsheets-list');
    if (!list) return;

    list.innerHTML = spreadsheets.map((spreadsheet: SpreadSheet) => `
      <div class="spreadsheet-card">
        <div class="card-content" onclick="window.location.href='/app?id=${spreadsheet.id}'">
          <div class="spreadsheet-title">${spreadsheet.title}</div>
          <div class="spreadsheet-date">
            ${new Date(spreadsheet.created_at).toLocaleDateString()}
          </div>
          <button class="delete-btn" data-sheet-id="${spreadsheet.id}" onclick="event.stopPropagation()">
            <span class="material-icons">delete</span>
          </button>
        </div>
      </div>
    `).join('');

    // Add delete handlers
    const deleteButtons = list.querySelectorAll('.delete-btn');
    deleteButtons.forEach(button => {
      button.addEventListener('click', async (e) => {
        e.preventDefault();
        e.stopPropagation();
        
        const sheetId = button.getAttribute('data-sheet-id');
        if (!sheetId) return;

        if (confirm('Are you sure you want to delete this spreadsheet?')) {
          try {
            const { data, error } = await supabase.functions.invoke('delete-spreadsheet', {
              method: 'POST',
              headers: {
                Authorization: `Bearer ${session.access_token}`,
                'apikey': supabaseAnonKey,
              },
              body: { id: sheetId }
            });

            if (error) throw error;

            if (data?.success) {
              const card = button.closest('.spreadsheet-card');
              if (card) {
                card.remove();
              } else {
                await loadSpreadsheets();
              }
            } else {
              throw new Error('Delete operation failed');
            }
          } catch (error) {
            alert('Failed to delete spreadsheet. Please try again.');
            await loadSpreadsheets();
          }
        }
      });
    });
  } catch (error) {
    window.location.href = '/';
  }
}

function setupThemeToggle() {
  const themeToggle = document.querySelector('.theme-toggle-button');
  if (!themeToggle) return;

  // Initialize theme from localStorage
  const isDarkTheme = localStorage.getItem('darkTheme') === 'true';
  document.documentElement.classList.toggle('dark-theme', isDarkTheme);
  themeToggle.querySelector('.theme-check')?.classList.toggle('active', isDarkTheme);

  themeToggle.addEventListener('click', () => {
    document.documentElement.classList.toggle('dark-theme');
    const isNowDark = document.documentElement.classList.contains('dark-theme');
    localStorage.setItem('darkTheme', isNowDark.toString());
    themeToggle.querySelector('.theme-check')?.classList.toggle('active', isNowDark);
  });
}

// Initialize
document.addEventListener('DOMContentLoaded', async () => {
  if (await checkAuth()) {
    setupThemeToggle();
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
      logoutButton.addEventListener('click', handleLogout);
    }
  }
}); 