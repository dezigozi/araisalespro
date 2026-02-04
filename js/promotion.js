// ========================================
// 設定
// ========================================
const GAS_URL = 'https://script.google.com/macros/s/AKfycbwrd5rm_eKQGJgON83RW8qg5H0SkMkqk6Zmrwh-lM62cqG6he9Ugq-7vmN0wXaaj-a3Nw/exec';





// ========================================
// 状態管理
// ========================================
let materials = [];
let currentFilter = 'all';

// ========================================
// 初期化
// ========================================
document.addEventListener('DOMContentLoaded', async () => {
  setupModal();
  await loadMaterials();
});



// ========================================
// 資料一覧を取得
// ========================================
async function loadMaterials() {
  const result = await fetchAPI('getMaterials');
  if (result?.success) {
    materials = result.data;
  }

  // カテゴリボタンを動的生成
  generateCategoryButtons();

  renderMaterials();
}

// カテゴリボタンを動的生成
function generateCategoryButtons() {
  const container = document.getElementById('categoryFilter');

  // 既存のカテゴリボタン（「すべて」以外）を削除
  container.querySelectorAll('.filter-btn:not([data-category="all"])').forEach(btn => btn.remove());

  // 資料からユニークなカテゴリを抽出
  const categories = [...new Set(materials.map(m => m.category).filter(c => c))];

  // カテゴリボタンを追加
  categories.forEach(cat => {
    const btn = document.createElement('button');
    btn.className = 'filter-btn';
    btn.dataset.category = cat;
    btn.textContent = cat;
    container.appendChild(btn);
  });

  // イベントリスナーを設定
  setupCategoryFilter();
}

function renderMaterials() {
  const list = document.getElementById('materialsList');
  const filtered = currentFilter === 'all'
    ? materials
    : materials.filter(m => m.category === currentFilter);

  if (filtered.length === 0) {
    list.innerHTML = '<div class="empty-state">資料がありません</div>';
    return;
  }

  list.innerHTML = filtered.map(m => `
    <div class="material-card" data-id="${m.id}">
      <div class="material-header">
        <div class="material-info">
          <div class="material-title">${m.title}</div>
          <span class="material-category">${m.category}</span>
        </div>
      </div>
      <div class="material-actions">
        ${m.pdfLink ? `<button class="action-btn preview-btn" data-pdf="${m.pdfLink}" data-title="${m.title}">
          👁️ プレビュー
        </button>` : ''}
        ${m.youtubeLink ? `<button class="action-btn youtube" onclick="window.open('${m.youtubeLink}', '_blank')">
          📺 YouTube
        </button>` : ''}
        ${m.link ? `<button class="action-btn link-btn" onclick="window.open('${m.link}', '_blank')">
          🔗 リンク
        </button>` : ''}
      </div>
    </div>
  `).join('');

  // プレビューボタン
  list.querySelectorAll('.preview-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      openPdfModal(btn.dataset.pdf, btn.dataset.title);
    });
  });
}

// ========================================
// カテゴリフィルター
// ========================================
function setupCategoryFilter() {
  document.querySelectorAll('.filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentFilter = btn.dataset.category;
      renderMaterials();
    });
  });
}



// ========================================
// PDFモーダル
// ========================================
function setupModal() {
  const modal = document.getElementById('pdfModal');
  const closeBtn = document.getElementById('closeModal');

  closeBtn.addEventListener('click', closeModal);
  modal.addEventListener('click', (e) => {
    if (e.target === modal) closeModal();
  });
}

function openPdfModal(pdfLink, title) {
  const modal = document.getElementById('pdfModal');
  const frame = document.getElementById('pdfFrame');
  const titleEl = document.getElementById('modalTitle');

  titleEl.textContent = title;
  frame.src = pdfLink;
  modal.classList.add('show');
}

function closeModal() {
  const modal = document.getElementById('pdfModal');
  const frame = document.getElementById('pdfFrame');
  modal.classList.remove('show');
  frame.src = '';
}



// ========================================
// API通信
// ========================================
async function fetchAPI(action, params = {}) {
  const url = new URL(GAS_URL);
  url.searchParams.append('action', action);
  Object.keys(params).forEach(key => url.searchParams.append(key, params[key]));

  try {
    const res = await fetch(url);
    return await res.json();
  } catch (e) {
    console.error('API Error:', e);
    showToast('通信エラー', true);
    return null;
  }
}


// ========================================
// Toast通知
// ========================================
function showToast(msg, isError = false) {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.className = 'toast' + (isError ? ' error' : '');
  toast.classList.add('show');

  setTimeout(() => {
    toast.classList.remove('show');
  }, 3000);
}
