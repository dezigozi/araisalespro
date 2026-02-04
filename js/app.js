// ========================================
// 設定
// ========================================
const GAS_URL = 'https://script.google.com/macros/s/AKfycbwrd5rm_eKQGJgON83RW8qg5H0SkMkqk6Zmrwh-lM62cqG6he9Ugq-7vmN0wXaaj-a3Nw/exec';
const CACHE_KEY = 'sfa_master_cache';
const CACHE_DURATION = 24 * 60 * 60 * 1000; // 24時間

// ========================================
// キャッシュ管理
// ========================================
const CacheManager = {
  // キャッシュからデータ取得
  get() {
    try {
      const cached = localStorage.getItem(CACHE_KEY);
      if (!cached) return null;

      const { data, timestamp } = JSON.parse(cached);
      const now = Date.now();

      // キャッシュ期限チェック（24時間）
      if (now - timestamp > CACHE_DURATION) {
        console.log('⏰ キャッシュ期限切れ');
        this.clear();
        return null;
      }

      console.log('✅ キャッシュから読み込み');
      return data;
    } catch (e) {
      console.error('キャッシュ読み込みエラー:', e);
      return null;
    }
  },

  // キャッシュにデータ保存
  set(data) {
    try {
      const cacheData = {
        data,
        timestamp: Date.now()
      };
      localStorage.setItem(CACHE_KEY, JSON.stringify(cacheData));
      console.log('💾 キャッシュに保存');
    } catch (e) {
      console.error('キャッシュ保存エラー:', e);
    }
  },

  // キャッシュクリア
  clear() {
    localStorage.removeItem(CACHE_KEY);
    console.log('🗑️ キャッシュをクリア');
  },

  // キャッシュ情報取得
  getInfo() {
    try {
      const cached = localStorage.getItem(CACHE_KEY);
      if (!cached) return null;

      const { timestamp } = JSON.parse(cached);
      const age = Date.now() - timestamp;
      const ageMinutes = Math.floor(age / 60000);

      return {
        ageMinutes,
        ageText: ageMinutes < 60 ? `${ageMinutes}分前` : `${Math.floor(ageMinutes / 60)}時間前`
      };
    } catch (e) {
      return null;
    }
  }
};

// ========================================
// マスタデータ（メモリ内キャッシュ）
// ========================================
let masterData = {
  customers: [],      // 顧客一覧（会社名リスト）
  departments: {},    // { 会社名: [部署リスト] }
  contacts: {}        // { "会社名_部署名": [担当者リスト] }
};

// ========================================
// 状態管理
// ========================================
let state = {
  type: '訪問',
  met: '○',
  reaction: '',
  salesRep: ''
};

// ========================================
// 初期化
// ========================================
document.addEventListener('DOMContentLoaded', async () => {
  initDateTime();
  initSalesRep();
  setupEventListeners();
  setupTabs();
  setupEditActivityModal();

  // マスタデータ読み込み（キャッシュ優先）
  await initMasterData();

  // 提案商品を読み込み
  await loadProposalProducts();

  // URLパラメータからアクションリストの担当者情報を取得して自動選択
  await handleUrlParameters();

  loadActivities();
});

// ========================================
// URLパラメータ処理（アクションリストからの遷移対応）
// ========================================
async function handleUrlParameters() {
  const params = new URLSearchParams(window.location.search);
  const company = params.get('company');
  const department = params.get('department');
  const contact = params.get('contact');
  const fromActionList = params.get('fromActionList');

  if (fromActionList && company && department && contact) {
    // 会社を選択
    const companySelect = document.getElementById('companySelect');
    if (companySelect) {
      companySelect.value = company;
      await loadDepartments(company);
    }

    // 部署を選択
    const deptSelect = document.getElementById('departmentSelect');
    if (deptSelect) {
      deptSelect.value = department;
      await loadContacts(company, department);
    }

    // 担当者をチェック（少し遅延させてDOMが更新されるのを待つ）
    setTimeout(() => {
      const contactList = document.getElementById('contactCheckboxList');
      if (contactList) {
        const checkbox = contactList.querySelector(`input[value="${contact}"]`);
        if (checkbox) {
          checkbox.checked = true;
          checkbox.closest('.contact-checkbox-item')?.classList.add('checked');
        }
      }
    }, 100);

    // URLパラメータをクリア（ブラウザ履歴置換）
    window.history.replaceState({}, document.title, window.location.pathname);

    showToast(`📋 ${contact} さんの活動を記録してください`);
  }
}

// 営業担当者の初期化（保存済みの選択を復元）
function initSalesRep() {
  const savedRep = localStorage.getItem('selectedSalesRep');
  const select = document.getElementById('salesRepSelect');

  if (savedRep && select) {
    select.value = savedRep;
    state.salesRep = savedRep;
  }
}

// マスタデータ初期化
async function initMasterData() {
  // キャッシュチェック
  const cached = CacheManager.get();

  if (cached) {
    // キャッシュあり → 即座に使用
    masterData = cached;
    populateCompanySelect();
    updateCacheStatus();
    showToast('⚡ キャッシュから読み込みました');
  } else {
    // キャッシュなし → APIから取得
    await loadAllMasterData();
  }
}

// 全マスタデータを一括取得
async function loadAllMasterData() {
  showToast('📊 マスタデータ読み込み中...');

  const result = await fetchAPI('getAllMasterData');

  if (result?.success && result.data) {
    masterData = result.data;
    CacheManager.set(masterData);
    populateCompanySelect();
    updateCacheStatus();
    showToast('✅ マスタデータを取得しました');
  } else {
    // フォールバック: 従来の方法で会社リストのみ取得
    console.log('getAllMasterData非対応、フォールバック実行');
    await loadCompaniesLegacy();
  }
}

// 会社プルダウンを設定
function populateCompanySelect() {
  const select = document.getElementById('companySelect');
  select.innerHTML = '<option value="">選択してください</option>';

  masterData.customers.forEach(c => {
    const opt = document.createElement('option');
    opt.value = c;
    opt.textContent = c;
    select.appendChild(opt);
  });
}

// キャッシュ状態を表示
function updateCacheStatus() {
  const info = CacheManager.getInfo();
  if (info) {
    console.log(`📋 キャッシュ: ${info.ageText}に更新`);
  }
}

// 現在日付をセット
function initDateTime() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  document.getElementById('visitDate').value = `${year}-${month}-${day}`;
}

// ========================================
// タブ切り替え
// ========================================
function setupTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      // タブボタン切り替え
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      // コンテンツ切り替え
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');

      // 履歴タブの場合は再読み込み
      if (btn.dataset.tab === 'list') {
        loadActivities();
      }
    });
  });
}

// ========================================
// イベントリスナー設定
// ========================================
function setupEventListeners() {
  // アプローチ手段
  document.querySelectorAll('.approach-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.approach-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.type = btn.dataset.type;
    });
  });

  // 会えた？
  document.querySelectorAll('.met-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.met-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.met = btn.dataset.met;
    });
  });

  // 反応
  document.querySelectorAll('.reaction-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.reaction-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.reaction = btn.dataset.reaction;
    });
  });

  // 会社選択 - キャッシュから即座に表示
  document.getElementById('companySelect').addEventListener('change', (e) => {
    if (e.target.value) {
      loadDepartments(e.target.value);
    } else {
      resetDepartmentSelect();
      resetContactSelect();
    }
  });

  // 部署選択 - キャッシュから即座に表示
  document.getElementById('departmentSelect').addEventListener('change', (e) => {
    const company = document.getElementById('companySelect').value;
    if (company && e.target.value) {
      loadContacts(company, e.target.value);
    } else {
      resetContactSelect();
    }
  });

  // 担当者検索
  const contactSearch = document.getElementById('contactSearch');
  if (contactSearch) {
    contactSearch.addEventListener('input', (e) => {
      filterContacts(e.target.value);
    });
  }

  // フォーム送信
  document.getElementById('activityForm').addEventListener('submit', handleSubmit);

  // 営業担当者選択
  const salesRepSelect = document.getElementById('salesRepSelect');
  if (salesRepSelect) {
    salesRepSelect.addEventListener('change', (e) => {
      state.salesRep = e.target.value;
      // 選択を保存
      if (e.target.value) {
        localStorage.setItem('selectedSalesRep', e.target.value);
      } else {
        localStorage.removeItem('selectedSalesRep');
      }
    });
  }

  // マスタ更新ボタン
  const refreshBtn = document.getElementById('refreshMasterBtn');
  if (refreshBtn) {
    refreshBtn.addEventListener('click', refreshMasterData);
  }

  // 「その他」チェックボックス
  const proposalOther = document.getElementById('proposalOther');
  const proposalOtherText = document.getElementById('proposalOtherText');
  if (proposalOther && proposalOtherText) {
    proposalOther.addEventListener('change', () => {
      proposalOtherText.disabled = !proposalOther.checked;
      if (!proposalOther.checked) {
        proposalOtherText.value = '';
      }
    });
  }

  // 新規担当者登録モーダル
  setupAddContactModal();

  // 履歴フィルター
  const historyFilter = document.getElementById('historyFilterSelect');
  if (historyFilter) {
    historyFilter.addEventListener('change', () => {
      loadActivities();
    });
  }
}

// ========================================
// マスタデータ更新（手動）
// ========================================
async function refreshMasterData() {
  const btn = document.getElementById('refreshMasterBtn');
  if (btn) {
    btn.disabled = true;
    btn.textContent = '更新中...';
  }

  CacheManager.clear();
  await loadAllMasterData();

  // プルダウンをリセット
  resetDepartmentSelect();
  resetContactSelect();

  if (btn) {
    btn.disabled = false;
    btn.textContent = '🔄 マスタ更新';
  }
}

// ========================================
// API呼び出し
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

async function postAPI(data) {
  try {
    // GAS Web App has CORS limitations for POST requests
    // Use no-cors mode and verify by reloading data
    await fetch(GAS_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(data)
    });
    // Cannot read response with no-cors, return success
    return { success: true };
  } catch (e) {
    console.error('API Error:', e);
    showToast('通信エラー', true);
    return null;
  }
}

// ========================================
// 提案商品マスタ取得・描画
// ========================================
async function loadProposalProducts() {
  const container = document.getElementById('proposalProducts');
  if (!container) return;

  // 現在の年月を取得
  const now = new Date();
  const yearMonth = `${now.getFullYear()}年${now.getMonth() + 1}月`;

  const result = await fetchAPI('getProposalProducts', { yearMonth });

  if (result?.success && result.data && result.data.length > 0) {
    container.innerHTML = result.data.map((product, index) => `
      <div class="proposal-item" id="proposal-item-${index}">
        <input type="checkbox" id="proposal-${index}" value="${product}">
        <label for="proposal-${index}">${product}</label>
      </div>
    `).join('');

    // チェックボックスのイベントリスナー
    container.querySelectorAll('.proposal-item').forEach(item => {
      const checkbox = item.querySelector('input[type="checkbox"]');
      item.addEventListener('click', (e) => {
        if (e.target.tagName !== 'INPUT') {
          checkbox.checked = !checkbox.checked;
        }
        item.classList.toggle('checked', checkbox.checked);
      });
    });
  } else {
    container.innerHTML = '<div class="no-proposals">この月の重点商品は設定されていません</div>';
  }
}

// 選択された提案商品を取得
function getSelectedProposals() {
  const proposals = [];

  // チェックされた商品を取得
  document.querySelectorAll('#proposalProducts input[type="checkbox"]:checked').forEach(cb => {
    proposals.push(cb.value);
  });

  // 「その他」がチェックされている場合
  const proposalOther = document.getElementById('proposalOther');
  const proposalOtherText = document.getElementById('proposalOtherText');
  if (proposalOther?.checked && proposalOtherText?.value.trim()) {
    proposals.push(proposalOtherText.value.trim());
  }

  return proposals;
}

// 提案商品をリセット
function resetProposalProducts() {
  document.querySelectorAll('#proposalProducts input[type="checkbox"]').forEach(cb => {
    cb.checked = false;
    cb.closest('.proposal-item')?.classList.remove('checked');
  });

  const proposalOther = document.getElementById('proposalOther');
  const proposalOtherText = document.getElementById('proposalOtherText');
  if (proposalOther) proposalOther.checked = false;
  if (proposalOtherText) {
    proposalOtherText.value = '';
    proposalOtherText.disabled = true;
  }
}

// ========================================
// データ読み込み（キャッシュ使用版）
// ========================================

// 部署読み込み - キャッシュから即座に表示 ⚡
function loadDepartments(company) {
  const select = document.getElementById('departmentSelect');
  select.innerHTML = '<option value="">選択してください</option>';

  const departments = masterData.departments[company] || [];

  if (departments.length > 0) {
    // キャッシュにあれば即座に表示
    departments.forEach(d => {
      const opt = document.createElement('option');
      opt.value = d;
      opt.textContent = d;
      select.appendChild(opt);
    });
    select.disabled = false;
    resetContactSelect();
  } else {
    // キャッシュにない場合はAPIから取得（フォールバック）
    loadDepartmentsFromAPI(company);
  }
}

// 担当者読み込み - チェックボックスリストに変更 ⚡
function loadContacts(company, department) {
  const listContainer = document.getElementById('contactCheckboxList');
  const searchInput = document.getElementById('contactSearch');
  const addBtn = document.getElementById('addContactBtn');

  const key = `${company}_${department}`;
  const contacts = masterData.contacts[key] || [];

  if (contacts.length > 0) {
    // チェックボックスリストを生成
    listContainer.innerHTML = contacts.map((c, idx) => `
      <div class="contact-checkbox-item" data-name="${c}">
        <input type="checkbox" id="contact-${idx}" value="${c}">
        <label for="contact-${idx}">${c}</label>
      </div>
    `).join('');

    // 各アイテムにクリックイベントを追加
    listContainer.querySelectorAll('.contact-checkbox-item').forEach(item => {
      item.addEventListener('click', (e) => {
        if (e.target.tagName !== 'INPUT') {
          const checkbox = item.querySelector('input[type="checkbox"]');
          checkbox.checked = !checkbox.checked;
        }
        item.classList.toggle('checked', item.querySelector('input').checked);
      });
    });

    listContainer.classList.remove('disabled');
    searchInput.disabled = false;
    searchInput.value = '';
    if (addBtn) addBtn.disabled = false;
  } else {
    // キャッシュにない場合はAPIから取得（フォールバック）
    loadContactsFromAPI(company, department);
  }
}

// 担当者の検索フィルタリング（チェック済みは常に表示）
function filterContacts(query) {
  const listContainer = document.getElementById('contactCheckboxList');
  const items = listContainer.querySelectorAll('.contact-checkbox-item');
  const searchTerm = query.toLowerCase().trim();

  items.forEach(item => {
    const name = item.dataset.name.toLowerCase();
    const isChecked = item.querySelector('input[type="checkbox"]').checked;

    // チェック済みの人は常に表示、またはマッチする場合は表示
    if (isChecked || searchTerm === '' || name.includes(searchTerm)) {
      item.classList.remove('hidden');
    } else {
      item.classList.add('hidden');
    }
  });

  // チェック済みのアイテムを上部に移動
  sortContactsByChecked();
}

// チェック済みのアイテムを上部に並べ替え
function sortContactsByChecked() {
  const listContainer = document.getElementById('contactCheckboxList');
  const items = Array.from(listContainer.querySelectorAll('.contact-checkbox-item'));

  // チェック済みのものを先に、未チェックを後にソート
  items.sort((a, b) => {
    const aChecked = a.querySelector('input[type="checkbox"]').checked ? 1 : 0;
    const bChecked = b.querySelector('input[type="checkbox"]').checked ? 1 : 0;
    return bChecked - aChecked; // チェック済みが上
  });

  // 並べ替えた順序でDOMに追加し直す
  items.forEach(item => listContainer.appendChild(item));
}

// 選択された担当者を取得
function getSelectedContacts() {
  const contacts = [];
  document.querySelectorAll('#contactCheckboxList input[type="checkbox"]:checked').forEach(cb => {
    contacts.push(cb.value);
  });
  return contacts;
}

// ========================================
// フォールバック用API呼び出し
// ========================================
async function loadCompaniesLegacy() {
  const result = await fetchAPI('getCustomers');
  if (result?.success) {
    masterData.customers = result.data;
    populateCompanySelect();
  }
}

async function loadDepartmentsFromAPI(company) {
  const select = document.getElementById('departmentSelect');
  select.innerHTML = '<option value="">読み込み中...</option>';

  const result = await fetchAPI('getDepartments', { company });
  if (result?.success) {
    // キャッシュに追加
    masterData.departments[company] = result.data;
    CacheManager.set(masterData);

    select.innerHTML = '<option value="">選択してください</option>';
    result.data.forEach(d => {
      const opt = document.createElement('option');
      opt.value = d;
      opt.textContent = d;
      select.appendChild(opt);
    });
    select.disabled = false;
    resetContactSelect();
  }
}

async function loadContactsFromAPI(company, department) {
  const select = document.getElementById('contactSelect');
  const addBtn = document.getElementById('addContactBtn');
  select.innerHTML = '<option value="">読み込み中...</option>';

  const result = await fetchAPI('getContactsByDept', { company, department });
  if (result?.success) {
    // キャッシュに追加
    const key = `${company}_${department}`;
    masterData.contacts[key] = result.data;
    CacheManager.set(masterData);

    select.innerHTML = '<option value="">選択してください</option>';
    result.data.forEach(c => {
      const opt = document.createElement('option');
      opt.value = c;
      opt.textContent = c;
      select.appendChild(opt);
    });
    select.disabled = false;
    if (addBtn) addBtn.disabled = false;
  }
}

// ========================================
// 活動履歴読み込み（編集・削除ボタン付き）
// ========================================
async function loadActivities() {
  const listEl = document.getElementById('activityList');
  listEl.innerHTML = '<div class="loading">読み込み中</div>';

  const result = await fetchAPI('getActivities');

  if (result?.success) {
    let activities = result.data;

    // 履歴フィルターが選択されている場合は適用
    const historyFilter = document.getElementById('historyFilterSelect');
    const filterValue = historyFilter ? historyFilter.value : '';
    if (filterValue) {
      activities = activities.filter(a => a.salesRep === filterValue);
    }

    if (activities.length === 0) {
      listEl.innerHTML = '<div class="empty-state">まだ活動記録がありません</div>';
      return;
    }

    listEl.innerHTML = activities.map(a => {
      const typeClass = a.type === '訪問' ? 'visit' : a.type === '電話' ? 'phone' : 'email';
      const reactionClass = a.reaction === '好反応' ? 'good' : a.reaction === '普通' ? 'normal' : 'bad';

      const date = new Date(a.datetime);
      const dateStr = `${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getDate().toString().padStart(2, '0')}`;

      return `
        <div class="activity-item" data-id="${a.id}" data-activity='${JSON.stringify(a).replace(/'/g, "&apos;")}'>
          <div class="activity-header">
            <span class="activity-date">${dateStr}</span>
            <span class="activity-type ${typeClass}">${a.type}</span>
            <span class="activity-reaction ${reactionClass}">${a.reaction}</span>
            ${a.salesRep ? `<span class="activity-sales-rep">👤 ${a.salesRep}</span>` : ''}
          </div>
          <div class="activity-company">${a.company}</div>
          <div class="activity-contact">${a.contact}様（${a.department}）</div>
          ${a.note ? `<div class="activity-note">${a.note}</div>` : ''}
          <div class="activity-actions">
            <button type="button" class="action-btn edit-btn" data-id="${a.id}" title="編集">
              ✏️ 編集
            </button>
            <button type="button" class="action-btn delete-btn" data-id="${a.id}" title="削除">
              🗑️ 削除
            </button>
          </div>
        </div>
      `;
    }).join('');

    // 編集ボタンのイベント
    listEl.querySelectorAll('.edit-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const item = btn.closest('.activity-item');
        const activityData = JSON.parse(item.dataset.activity.replace(/&apos;/g, "'"));
        openEditModal(activityData);
      });
    });

    // 削除ボタンのイベント
    listEl.querySelectorAll('.delete-btn').forEach(btn => {
      btn.addEventListener('click', async (e) => {
        e.stopPropagation();
        const id = btn.dataset.id;
        if (confirm('この活動記録を削除しますか？')) {
          await deleteActivityRecord(id);
        }
      });
    });
  } else {
    listEl.innerHTML = '<div class="empty-state">読み込み失敗</div>';
  }
}

// ========================================
// 活動記録削除処理
// ========================================
async function deleteActivityRecord(id) {
  showToast('削除中...');

  const result = await postAPI({
    action: 'deleteActivity',
    id: id
  });

  if (result) {
    showToast('活動記録を削除しました');
    loadActivities();
  } else {
    showToast('削除に失敗しました', true);
  }
}

// ========================================
// 編集モーダルを開く
// ========================================
async function openEditModal(activity) {
  const modal = document.getElementById('editActivityModal');
  if (!modal) return;

  // フォームに値をセット
  document.getElementById('editActivityId').value = activity.id;
  document.getElementById('editVisitDate').value = activity.datetime.split('T')[0];
  document.getElementById('editCompany').value = activity.company;
  document.getElementById('editDepartment').value = activity.department;
  document.getElementById('editNoteText').value = activity.note || '';

  // 担当者リストを動的に生成（ラジオボタン＋検索）
  const listContainer = document.getElementById('editContactList');
  const searchInput = document.getElementById('editContactSearch');
  const key = `${activity.company}_${activity.department}`;
  let contacts = masterData.contacts[key] || [];

  // 現在の担当者がリストにない場合は追加
  if (activity.contact && !contacts.includes(activity.contact)) {
    contacts = [activity.contact, ...contacts];
  }

  if (contacts.length > 0) {
    listContainer.innerHTML = contacts.map((c, idx) => {
      const isSelected = c === activity.contact;
      return `
        <div class="contact-radio-item ${isSelected ? 'selected' : ''}" data-name="${c}">
          <input type="radio" name="editContactRadio" id="editContact-${idx}" value="${c}" ${isSelected ? 'checked' : ''}>
          <label for="editContact-${idx}">${c}</label>
        </div>
      `;
    }).join('');

    // ラジオボタンのクリックイベント
    listContainer.querySelectorAll('.contact-radio-item').forEach(item => {
      item.addEventListener('click', (e) => {
        if (e.target.tagName !== 'INPUT') {
          item.querySelector('input[type="radio"]').checked = true;
        }
        listContainer.querySelectorAll('.contact-radio-item').forEach(i => i.classList.remove('selected'));
        item.classList.add('selected');
      });
    });

    // 検索フィルタ
    searchInput.value = '';
    searchInput.oninput = (e) => {
      const query = e.target.value.toLowerCase().trim();
      listContainer.querySelectorAll('.contact-radio-item').forEach(item => {
        const name = item.dataset.name.toLowerCase();
        const isSelected = item.querySelector('input[type="radio"]').checked;
        // 選択済みは常に表示、またはマッチする場合は表示
        if (isSelected || query === '' || name.includes(query)) {
          item.classList.remove('hidden');
        } else {
          item.classList.add('hidden');
        }
      });
    };
  } else {
    listContainer.innerHTML = '<div class="contact-placeholder">担当者がいません</div>';
  }

  // 当社担当を選択
  document.getElementById('editSalesRep').value = activity.salesRep || '';

  // アプローチ種類
  document.querySelectorAll('#editActivityModal .edit-approach-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.type === activity.type);
  });

  // 会えた？
  document.querySelectorAll('#editActivityModal .edit-met-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.met === activity.met);
  });

  // 反応
  document.querySelectorAll('#editActivityModal .edit-reaction-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.reaction === activity.reaction);
  });

  // 提案商品を動的に生成
  await loadEditProposalProducts(activity.proposals || []);

  modal.style.display = 'flex';
}

// ========================================
// 編集モーダル用提案商品読み込み
// ========================================
async function loadEditProposalProducts(selectedProposals) {
  const container = document.getElementById('editProposalProducts');
  if (!container) return;

  // 現在の年月を取得
  const now = new Date();
  const yearMonth = `${now.getFullYear()}年${now.getMonth() + 1}月`;

  const result = await fetchAPI('getProposalProducts', { yearMonth });

  if (result?.success && result.data && result.data.length > 0) {
    container.innerHTML = result.data.map((product, index) => {
      const isChecked = selectedProposals.includes(product);
      return `
        <div class="proposal-item ${isChecked ? 'checked' : ''}" id="edit-proposal-item-${index}">
          <input type="checkbox" id="edit-proposal-${index}" value="${product}" ${isChecked ? 'checked' : ''}>
          <label for="edit-proposal-${index}">${product}</label>
        </div>
      `;
    }).join('');

    // チェックボックスのイベントリスナー
    container.querySelectorAll('.proposal-item').forEach(item => {
      const checkbox = item.querySelector('input[type="checkbox"]');
      item.addEventListener('click', (e) => {
        if (e.target.tagName !== 'INPUT') {
          checkbox.checked = !checkbox.checked;
        }
        item.classList.toggle('checked', checkbox.checked);
      });
    });
  } else {
    container.innerHTML = '<div class="no-proposals">この月の重点商品は設定されていません</div>';
  }
}

// ========================================
// 編集モーダルセットアップ
// ========================================
function setupEditActivityModal() {
  const modal = document.getElementById('editActivityModal');
  if (!modal) return;

  const closeBtn = document.getElementById('closeEditModal');
  const cancelBtn = document.getElementById('cancelEditActivity');
  const submitBtn = document.getElementById('submitEditActivity');

  // 閉じる
  const closeModal = () => {
    modal.style.display = 'none';
  };

  closeBtn?.addEventListener('click', closeModal);
  cancelBtn?.addEventListener('click', closeModal);
  modal.addEventListener('click', (e) => {
    if (e.target === modal) closeModal();
  });

  // アプローチ種類ボタン
  modal.querySelectorAll('.edit-approach-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      modal.querySelectorAll('.edit-approach-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
    });
  });

  // 会えた？ボタン
  modal.querySelectorAll('.edit-met-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      modal.querySelectorAll('.edit-met-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
    });
  });

  // 反応ボタン
  modal.querySelectorAll('.edit-reaction-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      modal.querySelectorAll('.edit-reaction-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
    });
  });

  // 送信
  submitBtn?.addEventListener('click', async () => {
    const id = document.getElementById('editActivityId').value;
    const datetime = document.getElementById('editVisitDate').value;
    const company = document.getElementById('editCompany').value;
    const department = document.getElementById('editDepartment').value;
    const note = document.getElementById('editNoteText').value;
    const salesRep = document.getElementById('editSalesRep').value;

    // ラジオボタンから選択された担当者を取得
    const selectedContactRadio = modal.querySelector('input[name="editContactRadio"]:checked');
    const contact = selectedContactRadio ? selectedContactRadio.value : '';

    const type = modal.querySelector('.edit-approach-btn.active')?.dataset.type || '訪問';
    const met = modal.querySelector('.edit-met-btn.active')?.dataset.met || '○';
    const reaction = modal.querySelector('.edit-reaction-btn.active')?.dataset.reaction || '';

    if (!reaction) {
      showToast('反応を選択してください', true);
      return;
    }

    submitBtn.disabled = true;
    submitBtn.textContent = '更新中...';

    // 選択された提案商品を取得
    const proposals = [];
    document.querySelectorAll('#editProposalProducts input[type="checkbox"]:checked').forEach(cb => {
      proposals.push(cb.value);
    });

    try {
      const result = await postAPI({
        action: 'updateActivity',
        id,
        datetime,
        type,
        company,
        department,
        contact,
        reaction,
        met,
        note,
        salesRep,
        proposals
      });

      if (result) {
        showToast('活動記録を更新しました');
        closeModal();
        loadActivities();
      } else {
        throw new Error('更新に失敗しました');
      }
    } catch (error) {
      showToast(error.message || '更新に失敗しました', true);
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = '更新する';
    }
  });
}

// ========================================
// セレクトボックスリセット
// ========================================
function resetDepartmentSelect() {
  const s = document.getElementById('departmentSelect');
  s.innerHTML = '<option value="">会社を選択してください</option>';
  s.disabled = true;
}

function resetContactSelect() {
  const listContainer = document.getElementById('contactCheckboxList');
  const searchInput = document.getElementById('contactSearch');
  const addBtn = document.getElementById('addContactBtn');

  listContainer.innerHTML = '<div class="contact-placeholder">部署を選択してください</div>';
  listContainer.classList.add('disabled');
  searchInput.disabled = true;
  searchInput.value = '';
  if (addBtn) addBtn.disabled = true;
}

// ========================================
// フォーム送信
// ========================================
async function handleSubmit(e) {
  e.preventDefault();

  if (!state.salesRep) {
    showToast('営業担当者を選択してください', true);
    return;
  }

  // 担当者の選択チェック
  const selectedContacts = getSelectedContacts();
  if (selectedContacts.length === 0) {
    showToast('担当者を選択してください', true);
    return;
  }

  if (!state.reaction) {
    showToast('反応を選択してください', true);
    return;
  }

  const btn = document.getElementById('submitBtn');
  btn.disabled = true;
  btn.textContent = '送信中...';

  const data = {
    action: 'addActivity',
    datetime: document.getElementById('visitDate').value,
    type: state.type,
    salesRep: state.salesRep,
    company: document.getElementById('companySelect').value,
    department: document.getElementById('departmentSelect').value,
    contacts: selectedContacts, // 複数の担当者を送信
    reaction: state.reaction,
    met: state.met,
    note: document.getElementById('noteText').value,
    proposals: getSelectedProposals() // 提案商品を追加
  };

  const result = await postAPI(data);

  if (result) {
    showToast('活動を記録しました！');

    // フォームリセット
    document.getElementById('activityForm').reset();
    initDateTime();
    state.reaction = '';
    document.querySelectorAll('.reaction-btn').forEach(b => b.classList.remove('active'));
    resetDepartmentSelect();
    resetContactSelect();
    resetProposalProducts(); // 提案商品もリセット
  }

  btn.disabled = false;
  btn.textContent = '活動を記録';
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

// ========================================
// 新規担当者登録モーダル
// ========================================
function setupAddContactModal() {
  const modal = document.getElementById('addContactModal');
  const openBtn = document.getElementById('addContactBtn');
  const closeBtn = document.getElementById('closeAddContactModal');
  const cancelBtn = document.getElementById('cancelAddContact');
  const submitBtn = document.getElementById('submitAddContact');
  const nameInput = document.getElementById('newContactName');
  const modalCompany = document.getElementById('modalCompany');
  const modalDepartment = document.getElementById('modalDepartment');

  if (!modal || !openBtn) return;

  // モーダルを開く
  openBtn.addEventListener('click', () => {
    const company = document.getElementById('companySelect').value;
    const department = document.getElementById('departmentSelect').value;

    if (!company || !department) {
      showToast('会社と部署を選択してください', true);
      return;
    }

    modalCompany.textContent = company;
    modalDepartment.textContent = department;
    nameInput.value = '';
    modal.style.display = 'flex';
    nameInput.focus();
  });

  // モーダルを閉じる
  const closeModal = () => {
    modal.style.display = 'none';
    nameInput.value = '';
  };

  closeBtn.addEventListener('click', closeModal);
  cancelBtn.addEventListener('click', closeModal);

  // 背景クリックで閉じる
  modal.addEventListener('click', (e) => {
    if (e.target === modal) closeModal();
  });

  // 登録処理
  submitBtn.addEventListener('click', async () => {
    const contactName = nameInput.value.trim();
    const company = document.getElementById('companySelect').value;
    const department = document.getElementById('departmentSelect').value;

    if (!contactName) {
      showToast('担当者名を入力してください', true);
      nameInput.focus();
      return;
    }

    submitBtn.disabled = true;
    submitBtn.textContent = '登録中...';

    try {
      const result = await postAPI({
        action: 'addContact',
        company: company,
        department: department,
        contactName: contactName
      });

      if (result && result.success !== false) {
        showToast(`★ ${contactName} を登録しました`);

        // キャッシュを更新
        const key = `${company}_${department}`;
        if (!masterData.contacts[key]) {
          masterData.contacts[key] = [];
        }
        masterData.contacts[key].push(contactName);
        CacheManager.set(masterData);

        // プルダウンに追加して選択状態に
        const select = document.getElementById('contactSelect');
        const opt = document.createElement('option');
        opt.value = contactName;
        opt.textContent = contactName;
        select.appendChild(opt);
        select.value = contactName;

        closeModal();
      } else {
        throw new Error(result?.error || '登録に失敗しました');
      }
    } catch (error) {
      showToast(error.message || '通信エラー', true);
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = '登録する';
    }
  });

  // Enterキーで送信
  nameInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      submitBtn.click();
    }
  });
}
