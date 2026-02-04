// ========================================
// 簡易パスワード認証
// ========================================
const AUTH_PASSWORD = '19627004';
const AUTH_KEY = 'arai_sales_pro_auth';

// 認証済みかチェック
function checkAuth() {
    return sessionStorage.getItem(AUTH_KEY) === 'authenticated';
}

// 認証画面を表示
function showAuthOverlay() {
    if (checkAuth()) return;

    const overlay = document.createElement('div');
    overlay.className = 'auth-overlay';
    overlay.id = 'authOverlay';
    overlay.innerHTML = `
    <div class="auth-box">
      <div class="auth-logo">ARAI <span>SALES PRO</span></div>
      <p class="auth-subtitle">パスワードを入力してください</p>
      <input type="password" class="auth-input" id="authPassword" placeholder="パスワード" inputmode="numeric" pattern="[0-9]*" maxlength="10" autocomplete="off">
      <button type="button" class="auth-btn" id="authSubmit">ログイン</button>
      <p class="auth-error" id="authError">パスワードが正しくありません</p>
    </div>
  `;
    document.body.appendChild(overlay);

    const input = document.getElementById('authPassword');
    const btn = document.getElementById('authSubmit');
    const error = document.getElementById('authError');

    // フォーカスを入力フィールドに
    setTimeout(() => input.focus(), 100);

    // ログインボタンクリック
    btn.addEventListener('click', attemptLogin);

    // Enterキーでログイン
    input.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') attemptLogin();
    });

    function attemptLogin() {
        if (input.value === AUTH_PASSWORD) {
            sessionStorage.setItem(AUTH_KEY, 'authenticated');
            overlay.remove();
        } else {
            error.classList.add('show');
            input.value = '';
            input.focus();
            setTimeout(() => error.classList.remove('show'), 2000);
        }
    }
}

// ページ読み込み時に認証チェック
document.addEventListener('DOMContentLoaded', showAuthOverlay);
