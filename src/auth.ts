import { supabase } from './supabase';

async function handleGitHubLogin() {
  try {
    const loginButton = document.getElementById('github-login-btn') as HTMLButtonElement;
    if (loginButton) {
      loginButton.disabled = true;
      loginButton.textContent = 'Signing in...';
    }

    const { data, error } = await supabase.auth.signInWithOAuth({
      provider: 'github',
      options: {
        redirectTo: `${window.location.origin}/dashboard`
      }
    });

    if (error) throw error;
  } catch (error) {
    console.error('Error logging in:', error);
    alert('Error logging in. Please try again.');
    
    // Reset button state
    const loginButton = document.getElementById('github-login-btn') as HTMLButtonElement;
    if (loginButton) {
      loginButton.disabled = false;
      loginButton.innerHTML = `
        <svg height="20" viewBox="0 0 16 16" width="20" class="github-icon">
          <path fill="currentColor" d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"></path>
        </svg>
        Sign in with GitHub
      `;
    }
  }
}

// Initialize login button handler
document.addEventListener('DOMContentLoaded', () => {
  const loginButton = document.getElementById('github-login-btn');
  if (loginButton) {
    loginButton.addEventListener('click', handleGitHubLogin);
  }
});

// Update the checkUser function
async function checkUser() {
  const { data: { user } } = await supabase.auth.getUser();

  if (user) {
    // If user is logged in, show profile
    const userContainer = document.getElementById('user-container');
    if (userContainer) {
      userContainer.innerHTML = `
        <div class="user-profile">
          <img src="${user.user_metadata.avatar_url}" alt="Profile" />
          <span>${user.user_metadata.name}</span>
        </div>
        <button id="logout-button" class="logout-button">Sign Out</button>
      `;
      
      // Add logout handler
      document.getElementById('logout-button')?.addEventListener('click', async () => {
        await supabase.auth.signOut();
        window.location.href = '/';
      });

      // Redirect to dashboard if on landing page
      if (window.location.pathname === '/' || window.location.pathname === '/index.html') {
        window.location.href = '/dashboard';
      }
    }
  } else {
    // If not logged in and on protected pages, redirect to landing
    if (window.location.pathname === '/app' || 
        window.location.pathname === '/app.html' || 
        window.location.pathname === '/dashboard' || 
        window.location.pathname === '/dashboard.html') {
      window.location.href = '/';
      return;
    }

    // Clear the user container when not logged in
    const userContainer = document.getElementById('user-container');
    if (userContainer) {
      userContainer.innerHTML = '';
    }
  }
}

// Check user status on page load
checkUser();

// Listen for auth state changes
supabase.auth.onAuthStateChange((event, session) => {
  if (event === 'SIGNED_IN') {
    window.location.href = '/dashboard';
  } else if (event === 'SIGNED_OUT') {
    window.location.href = '/';
  }
}); 