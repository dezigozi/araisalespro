// ========================================
// 売上分析ページ JavaScript
// ========================================

// GAS Web App URL（js/gas-config.js）
const API_URL = GAS_WEBAPP_URL;

// キャッシュキー
// キャッシュキー
const CACHE_KEY_ESTIMATE = 'salesAnalysisData_estimate';
const CACHE_KEY_ORDER = 'salesAnalysisData_order';
const CACHE_TIMESTAMP_KEY_ESTIMATE = 'salesAnalysisCacheTime_estimate';
const CACHE_TIMESTAMP_KEY_ORDER = 'salesAnalysisCacheTime_order';

const PHONE_CACHE_KEY = 'customerPhonesData';

// IndexedDB設定（大容量データ対応）
const IDB_NAME = 'SalesAnalysisDB';
const IDB_VERSION = 1;
const IDB_STORE = 'cache';

// ========================================
// IndexedDB ヘルパー関数
// ========================================

function openIndexedDB() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(IDB_NAME, IDB_VERSION);

        request.onerror = () => reject(request.error);
        request.onsuccess = () => resolve(request.result);

        request.onupgradeneeded = (event) => {
            const db = event.target.result;
            if (!db.objectStoreNames.contains(IDB_STORE)) {
                db.createObjectStore(IDB_STORE, { keyPath: 'key' });
            }
        };
    });
}

async function getFromIndexedDB(key) {
    try {
        const db = await openIndexedDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(IDB_STORE, 'readonly');
            const store = tx.objectStore(IDB_STORE);
            const request = store.get(key);

            request.onerror = () => reject(request.error);
            request.onsuccess = () => resolve(request.result?.value);

            tx.oncomplete = () => db.close();
        });
    } catch (e) {
        console.warn('IndexedDB読み込みエラー:', e);
        return null;
    }
}

async function setToIndexedDB(key, value) {
    try {
        const db = await openIndexedDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(IDB_STORE, 'readwrite');
            const store = tx.objectStore(IDB_STORE);
            const request = store.put({ key, value });

            request.onerror = () => reject(request.error);
            request.onsuccess = () => resolve(true);

            tx.oncomplete = () => db.close();
        });
    } catch (e) {
        console.warn('IndexedDB保存エラー:', e);
        return false;
    }
}

async function deleteFromIndexedDB(key) {
    try {
        const db = await openIndexedDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(IDB_STORE, 'readwrite');
            const store = tx.objectStore(IDB_STORE);
            const request = store.delete(key);

            request.onerror = () => reject(request.error);
            request.onsuccess = () => resolve(true);

            tx.oncomplete = () => db.close();
        });
    } catch (e) {
        console.warn('IndexedDB削除エラー:', e);
        return false;
    }
}

// グローバル状態
let allData = [];
let customerPhones = {}; // 部署→電話番号のマッピング
let branchOrder = []; // 部署の順序（担当者マスタの並び順）
let repSummaryData = []; // 第1階層: 担当者別
let clientSummaryData = []; // 第2階層: 顧客別
let productSummaryData = []; // 第3階層: 品番別
let currentView = 'rep'; // 'rep' | 'client' | 'product'
let currentRepFilter = null;
let currentClientFilter = null;
let currentSort = { field: 'totalAmount', direction: 'desc' };
let selectedClientName = null; // 顧客名横断検索で選択された顧客
let selectedRepName = null; // 担当者名横断検索で選択された担当者
let currentDataMode = 'estimate'; // 'estimate' (見積) | 'order' (受注)

// モード切り替え
async function switchDataMode(mode) {
    if (currentDataMode === mode) return;

    currentDataMode = mode;
    allData = []; // データリセット

    // UI更新
    const btnEstimate = document.getElementById('modeEstimate');
    const btnOrder = document.getElementById('modeOrder');

    const activeClass = ['bg-blue-600', 'text-white', 'shadow-sm', 'transform', 'scale-105', 'font-bold'];
    const inactiveClass = ['text-gray-400', 'hover:text-white', 'font-medium'];

    // クラスをリセットして適用
    if (mode === 'estimate') {
        btnEstimate.className = `px-8 py-2 rounded-md text-sm transition-all duration-200 ${activeClass.join(' ')}`;
        btnOrder.className = `px-8 py-2 rounded-md text-sm transition-all duration-200 ${inactiveClass.join(' ')}`;
    } else {
        btnOrder.className = `px-8 py-2 rounded-md text-sm transition-all duration-200 ${activeClass.join(' ')}`;
        btnEstimate.className = `px-8 py-2 rounded-md text-sm transition-all duration-200 ${inactiveClass.join(' ')}`;
    }

    // 画面クリア
    clearFilters();
    populateFilters(); // 空で再描画

    // キャッシュ確認と再描画
    await checkCacheStatus();

    showToast(`${mode === 'estimate' ? '見積' : '受注'}データに切り替えました`, 'info');
}

// 現在のキャッシュキーを取得
function getCurrentCacheKey() {
    return currentDataMode === 'estimate' ? CACHE_KEY_ESTIMATE : CACHE_KEY_ORDER;
}

function getCurrentTimestampKey() {
    return currentDataMode === 'estimate' ? CACHE_TIMESTAMP_KEY_ESTIMATE : CACHE_TIMESTAMP_KEY_ORDER;
}

// ========================================
// 初期化
// ========================================
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
    checkCacheStatus();
});

function initializeEventListeners() {
    // フィルタ関連
    document.getElementById('abbrFilter').addEventListener('change', onAbbrChange);
    document.getElementById('loadDataBtn').addEventListener('click', loadAllData);
    document.getElementById('searchBtn').addEventListener('click', executeSearch);
    document.getElementById('clearBtn').addEventListener('click', clearFilters);

    // モード切り替え
    document.getElementById('modeEstimate').addEventListener('click', () => switchDataMode('estimate'));
    document.getElementById('modeOrder').addEventListener('click', () => switchDataMode('order'));

    // 顧客名検索オートコンプリート
    const clientSearchInput = document.getElementById('clientNameSearch');
    if (clientSearchInput) {
        clientSearchInput.addEventListener('input', onClientNameInput);
        clientSearchInput.addEventListener('focus', onClientNameInput);
        clientSearchInput.addEventListener('blur', () => {
            setTimeout(() => hideSuggestions(), 150);
        });
    }

    // 担当者名検索オートコンプリート
    const repSearchInput = document.getElementById('repNameSearch');
    if (repSearchInput) {
        repSearchInput.addEventListener('input', onRepNameInput);
        repSearchInput.addEventListener('focus', onRepNameInput);
        repSearchInput.addEventListener('blur', () => {
            setTimeout(() => hideRepSuggestions(), 150);
        });
    }

    // 詳細画面
    document.getElementById('backToSummaryBtn').addEventListener('click', handleBackButton);

    // PDF出力
    document.getElementById('exportSummaryPdfBtn').addEventListener('click', () => exportToPdf('summary'));
    document.getElementById('exportDetailPdfBtn').addEventListener('click', () => exportToPdf('detail'));

    // テーブルヘッダークリックでソート
    document.querySelectorAll('#summaryTable th[data-sort]').forEach(th => {
        th.addEventListener('click', () => sortTable(th.dataset.sort));
    });
}

// ========================================
// データ処理ユーティリティ
// ========================================

/**
 * 半角カタカナ→全角カタカナ変換
 */
function hankakuToZenkakuKana(str) {
    if (!str) return '';

    const kanaMap = {
        'ｱ': 'ア', 'ｲ': 'イ', 'ｳ': 'ウ', 'ｴ': 'エ', 'ｵ': 'オ',
        'ｶ': 'カ', 'ｷ': 'キ', 'ｸ': 'ク', 'ｹ': 'ケ', 'ｺ': 'コ',
        'ｻ': 'サ', 'ｼ': 'シ', 'ｽ': 'ス', 'ｾ': 'セ', 'ｿ': 'ソ',
        'ﾀ': 'タ', 'ﾁ': 'チ', 'ﾂ': 'ツ', 'ﾃ': 'テ', 'ﾄ': 'ト',
        'ﾅ': 'ナ', 'ﾆ': 'ニ', 'ﾇ': 'ヌ', 'ﾈ': 'ネ', 'ﾉ': 'ノ',
        'ﾊ': 'ハ', 'ﾋ': 'ヒ', 'ﾌ': 'フ', 'ﾍ': 'ヘ', 'ﾎ': 'ホ',
        'ﾏ': 'マ', 'ﾐ': 'ミ', 'ﾑ': 'ム', 'ﾒ': 'メ', 'ﾓ': 'モ',
        'ﾔ': 'ヤ', 'ﾕ': 'ユ', 'ﾖ': 'ヨ',
        'ﾗ': 'ラ', 'ﾘ': 'リ', 'ﾙ': 'ル', 'ﾚ': 'レ', 'ﾛ': 'ロ',
        'ﾜ': 'ワ', 'ｦ': 'ヲ', 'ﾝ': 'ン',
        'ｧ': 'ァ', 'ｨ': 'ィ', 'ｩ': 'ゥ', 'ｪ': 'ェ', 'ｫ': 'ォ',
        'ｬ': 'ャ', 'ｭ': 'ュ', 'ｮ': 'ョ',
        'ｯ': 'ッ', 'ｰ': 'ー',
        'ﾞ': '゛', 'ﾟ': '゜'
    };

    // 濁点・半濁点の結合
    const dakutenMap = {
        'カ': 'ガ', 'キ': 'ギ', 'ク': 'グ', 'ケ': 'ゲ', 'コ': 'ゴ',
        'サ': 'ザ', 'シ': 'ジ', 'ス': 'ズ', 'セ': 'ゼ', 'ソ': 'ゾ',
        'タ': 'ダ', 'チ': 'ヂ', 'ツ': 'ヅ', 'テ': 'デ', 'ト': 'ド',
        'ハ': 'バ', 'ヒ': 'ビ', 'フ': 'ブ', 'ヘ': 'ベ', 'ホ': 'ボ',
        'ウ': 'ヴ'
    };

    const handakutenMap = {
        'ハ': 'パ', 'ヒ': 'ピ', 'フ': 'プ', 'ヘ': 'ペ', 'ホ': 'ポ'
    };

    let result = '';
    for (let i = 0; i < str.length; i++) {
        let char = str[i];

        // 半角カタカナを全角に変換
        if (kanaMap[char]) {
            char = kanaMap[char];
        }

        // 次の文字が濁点・半濁点の場合、結合
        if (i + 1 < str.length) {
            const nextChar = str[i + 1];
            if (nextChar === 'ﾞ' && dakutenMap[char]) {
                char = dakutenMap[char];
                i++; // 濁点をスキップ
            } else if (nextChar === 'ﾟ' && handakutenMap[char]) {
                char = handakutenMap[char];
                i++; // 半濁点をスキップ
            }
        }

        result += char;
    }

    return result;
}

/**
 * 顧客名の名寄せ（正規化）
 */
function normalizeCustomerName(name) {
    if (!name) return '';

    // 半角カタカナ→全角カタカナ（改善版）
    let normalized = hankakuToZenkakuKana(name);

    // 法人格を削除
    const corpTypes = [
        '株式会社', '㈱', '（株）', '(株)',
        '有限会社', '㈲', '（有）', '(有)',
        '合同会社', '合資会社', '合名会社',
        '一般社団法人', '一般財団法人',
        '公益社団法人', '公益財団法人'
    ];
    corpTypes.forEach(type => {
        normalized = normalized.replace(new RegExp(type, 'g'), '');
    });

    // スペース除去（全角・半角両方）
    normalized = normalized.replace(/[\s　]+/g, '').trim();

    return normalized;
}

/**
 * 担当者名から名字を抽出
 */
function extractLastName(fullName) {
    if (!fullName) return '';
    // 全角・半角スペースで分割し、最初の要素（名字）を返す
    const parts = fullName.split(/[\s　]+/);
    return parts[0] || fullName;
}

/**
 * 金額フォーマット
 */
function formatCurrency(amount) {
    if (typeof amount !== 'number') return '¥0';
    const isNegative = amount < 0;
    const formatted = Math.abs(amount).toLocaleString('ja-JP');
    return isNegative ? `-¥${formatted}` : `¥${formatted}`;
}

// ========================================
// キャッシュ管理
// ========================================

async function checkCacheStatus() {
    const statusEl = document.getElementById('cacheStatus');
    const textEl = document.getElementById('cacheText');

    // IndexedDBからキャッシュを読み込み
    try {
        const cacheKey = getCurrentCacheKey();
        const timeKey = getCurrentTimestampKey();

        const cachedData = await getFromIndexedDB(cacheKey);
        const cacheTime = await getFromIndexedDB(timeKey);
        const phoneData = await getFromIndexedDB(PHONE_CACHE_KEY);

        // 電話番号データの読み込み
        if (phoneData) {
            if (phoneData.phones) {
                customerPhones = phoneData.phones;
                branchOrder = phoneData.branchOrder || [];
            } else {
                customerPhones = phoneData;
                branchOrder = [];
            }
        }

        if (cachedData && cacheTime) {
            allData = cachedData;
            const date = new Date(cacheTime);
            textEl.innerHTML = `<span class="cache-badge fresh">✓ キャッシュあり: ${date.toLocaleString('ja-JP')} (${allData.length}件)</span>`;
            statusEl.classList.remove('hidden');

            // フィルタ選択肢を生成
            populateFilters();

            // バックグラウンドで更新チェック
            checkForUpdates();
        } else {
            textEl.innerHTML = '<span class="cache-badge none">キャッシュなし - 「一括読み込み」を実行してください</span>';
            statusEl.classList.remove('hidden');
        }
    } catch (e) {
        console.warn('キャッシュ読み込みエラー:', e);
        textEl.innerHTML = '<span class="cache-badge none">キャッシュなし - 「一括読み込み」を実行してください</span>';
        statusEl.classList.remove('hidden');
    }
}

async function checkForUpdates() {
    if (!API_URL) return;

    try {
        const response = await fetch(`${API_URL}?action=getSheetLastModified`);
        const result = await response.json();

        if (result.success) {
            const serverTime = new Date(result.data.lastModified).getTime();
            const timeKey = getCurrentTimestampKey();
            const cacheTime = await getFromIndexedDB(timeKey) || 0;

            if (serverTime > cacheTime) {
                const textEl = document.getElementById('cacheText');
                textEl.innerHTML = '<span class="cache-badge stale">⚠ 新しいデータがあります - 「一括読み込み」を実行してください</span>';
            }
        }
    } catch (e) {
        console.log('Update check failed:', e);
    }
}

// ========================================
// データ取得
// ========================================

async function loadAllData() {
    if (!API_URL) {
        showToast('GAS URLが設定されていません', 'error');
        return;
    }

    const btn = document.getElementById('loadDataBtn');
    const originalText = btn.innerHTML;
    btn.innerHTML = '<span class="loading-spinner"></span> 読み込み中... 0%';
    btn.disabled = true;

    try {
        let salesData = [];
        let phoneData = null;

        // 電話番号データは並行して取得（軽量なので一括でOK）
        const phonePromise = fetch(`${API_URL}?action=getCustomerPhones`)
            .then(res => res.json());

        if (currentDataMode === 'estimate') {
            // 見積データは従来通り一括取得（データ量が少ない場合）
            // ※必要に応じてこちらもページネーション化可
            const response = await fetch(`${API_URL}?action=getSalesAnalysisData`);
            const result = await response.json();

            if (!result.success) throw new Error(result.error || 'データ取得に失敗しました');
            salesData = result.data;

            // 電話番号データの完了を待つ
            const phoneResult = await phonePromise;
            if (phoneResult.success) phoneData = phoneResult.data;

        } else {
            // 受注データ（9万行超）は分割取得
            const CHUNK_SIZE = 3000;
            let offset = 0;
            let total = 0;
            let hasMore = true;

            // 電話番号データを先に待機（必須ではないがエラーチェック用）
            // const phoneResult = await phonePromise; 
            // 並行処理を阻害しないよう、ここでは待たない

            while (hasMore) {
                const response = await fetch(`${API_URL}?action=getOrderAnalysisData&offset=${offset}&limit=${CHUNK_SIZE}`);
                const result = await response.json();

                if (!result.success) throw new Error(result.error || 'データ取得に失敗しました');

                const chunkData = result.data;
                salesData = salesData.concat(chunkData);
                total = result.total || 0;

                // 進捗表示
                const progress = total > 0 ? Math.round((salesData.length / total) * 100) : 0;
                btn.innerHTML = `<span class="loading-spinner"></span> 読み込み中... ${progress}%`;

                if (chunkData.length < CHUNK_SIZE || salesData.length >= total) {
                    hasMore = false;
                } else {
                    offset += CHUNK_SIZE;
                }
            }

            // 電話番号データの完了を待つ
            const phoneResult = await phonePromise;
            if (phoneResult.success) phoneData = phoneResult.data;
        }

        allData = salesData;
        const now = Date.now();

        // IndexedDBにキャッシュ保存
        const cacheKey = getCurrentCacheKey();
        const timeKey = getCurrentTimestampKey();

        const saveSuccess = await setToIndexedDB(cacheKey, allData);
        await setToIndexedDB(timeKey, now);

        populateFilters();
        showToast(`${allData.length}件のデータを読み込みました`, 'success');

        // キャッシュ状態を表示更新
        const textEl = document.getElementById('cacheText');
        if (saveSuccess) {
            const date = new Date(now);
            textEl.innerHTML = `<span class="cache-badge fresh">✓ キャッシュあり: ${date.toLocaleString('ja-JP')} (${allData.length}件)</span>`;
        } else {
            textEl.innerHTML = `<span class="cache-badge stale">⚡ メモリ上で動作中（${allData.length}件）- キャッシュ保存に失敗</span>`;
        }

        if (phoneData) {
            // 新形式: { phones, branchOrder }
            if (phoneData.phones) {
                customerPhones = phoneData.phones;
                branchOrder = phoneData.branchOrder || [];
            } else {
                // 後方互換: 旧形式
                customerPhones = phoneData;
                branchOrder = [];
            }
            await setToIndexedDB(PHONE_CACHE_KEY, phoneData);
        }

    } catch (error) {
        showToast(error.message, 'error');
        console.error(error);
    } finally {
        btn.innerHTML = originalText;
        btn.disabled = false;
    }
}

// ========================================
// フィルタ処理
// ========================================

function populateFilters() {
    if (!allData.length) return;

    // 年月の重複を除いたリスト（降順ソート：新しい順）
    const yearMonths = [...new Set(allData.map(d => d.registYearMonth).filter(Boolean))].sort().reverse();

    // 開始年月プルダウン
    const startYMSelect = document.getElementById('startYearMonth');
    let startYMOpts = '<option value="">すべて</option>';
    yearMonths.forEach(ym => {
        const displayYM = ym.replace('/', '年') + '月';
        startYMOpts += `<option value="${ym}">${displayYM}</option>`;
    });
    startYMSelect.innerHTML = startYMOpts;

    // 終了年月プルダウン（デフォルトで最新を選択）
    const endYMSelect = document.getElementById('endYearMonth');
    let endYMOpts = '<option value="">すべて</option>';
    yearMonths.forEach(ym => {
        const displayYM = ym.replace('/', '年') + '月';
        endYMOpts += `<option value="${ym}">${displayYM}</option>`;
    });
    endYMSelect.innerHTML = endYMOpts;
    // 最新の年月をデフォルト選択
    if (yearMonths.length > 0) {
        endYMSelect.value = yearMonths[0];
    }

    // 略称の重複を除いたリスト
    const abbrs = [...new Set(allData.map(d => d.abbr).filter(Boolean))].sort();
    const abbrSelect = document.getElementById('abbrFilter');
    let abbrOpts = '<option value="">すべて</option>';
    abbrs.forEach(abbr => {
        abbrOpts += `<option value="${abbr}">${abbr}</option>`;
    });
    abbrSelect.innerHTML = abbrOpts;

    // 部店: 担当者マスタの順序を使用（データに存在するもののみ）
    const dataBranches = new Set(allData.map(d => d.branch).filter(Boolean));
    const orderedBranches = branchOrder.filter(b => dataBranches.has(b));
    // branchOrderにないけどデータにある部店も追加
    dataBranches.forEach(b => {
        if (!orderedBranches.includes(b)) {
            orderedBranches.push(b);
        }
    });
    const branchSelect = document.getElementById('branchFilter');
    let branchOpts = '<option value="">すべて</option>';
    orderedBranches.forEach(branch => {
        branchOpts += `<option value="${branch}">${branch}</option>`;
    });
    branchSelect.innerHTML = branchOpts;
}

function onAbbrChange() {
    const abbr = document.getElementById('abbrFilter').value;
    const branchSelect = document.getElementById('branchFilter');

    // データに存在する部店を取得
    let dataBranches;
    if (!abbr) {
        // 略称未選択時は全部店
        dataBranches = new Set(allData.map(d => d.branch).filter(Boolean));
    } else {
        // 選択された略称に紐づく部店のみ
        dataBranches = new Set(
            allData.filter(d => d.abbr === abbr).map(d => d.branch).filter(Boolean)
        );
    }

    // 担当者マスタの順序を使用
    const orderedBranches = branchOrder.filter(b => dataBranches.has(b));
    // branchOrderにないけどデータにある部店も追加
    dataBranches.forEach(b => {
        if (!orderedBranches.includes(b)) {
            orderedBranches.push(b);
        }
    });

    let branchOpts = '<option value="">すべて</option>';
    orderedBranches.forEach(branch => {
        branchOpts += `<option value="${branch}">${branch}</option>`;
    });
    branchSelect.innerHTML = branchOpts;
}

function clearFilters() {
    document.getElementById('startYearMonth').value = '';
    document.getElementById('endYearMonth').value = '';
    document.getElementById('abbrFilter').value = '';
    document.getElementById('hqOnlyFilter').checked = false;

    // 部店を全件に戻す（担当者マスタの順序）
    const dataBranches = new Set(allData.map(d => d.branch).filter(Boolean));
    const orderedBranches = branchOrder.filter(b => dataBranches.has(b));
    dataBranches.forEach(b => {
        if (!orderedBranches.includes(b)) {
            orderedBranches.push(b);
        }
    });
    const branchSelect = document.getElementById('branchFilter');
    let branchOpts = '<option value="">すべて</option>';
    orderedBranches.forEach(branch => {
        branchOpts += `<option value="${branch}">${branch}</option>`;
    });
    branchSelect.innerHTML = branchOpts;

    // 顧客名検索をクリア
    const clientSearchInput = document.getElementById('clientNameSearch');
    if (clientSearchInput) {
        clientSearchInput.value = '';
    }
    selectedClientName = null;

    // 担当者名検索をクリア
    const repSearchInput = document.getElementById('repNameSearch');
    if (repSearchInput) {
        repSearchInput.value = '';
    }
    selectedRepName = null;

    // ヘッダーの顧客名をクリア
    updateHeaderCustomerName('');

    // 状態リセット
    currentView = 'rep';
    currentRepFilter = null;
    currentClientFilter = null;
    repSummaryData = [];
    clientSummaryData = [];
    productSummaryData = [];

    // サマリーをクリア
    document.getElementById('summaryTableBody').innerHTML =
        '<tr><td colspan="4" class="px-4 py-8 text-center text-gray-500">🔍 検索条件を指定して「検索」ボタンを押してください</td></tr>';
    document.getElementById('summaryCards').innerHTML =
        '<div class="text-center text-gray-500 py-8">🔍 検索条件を指定して「検索」ボタンを押してください</div>';

    // 詳細画面を隠す
    document.getElementById('detailSection').classList.add('hidden');
    document.getElementById('summarySection').classList.remove('hidden');
}

// ========================================
// 顧客名オートコンプリート
// ========================================

function onClientNameInput(e) {
    const query = e.target.value.trim();

    if (query.length < 1) {
        hideSuggestions();
        selectedClientName = null;
        return;
    }

    const suggestions = searchClients(query);
    showSuggestions(suggestions);
}

function searchClients(query) {
    if (!allData.length) return [];

    const queryLower = query.toLowerCase();
    const queryNormalized = normalizeCustomerName(query);

    // 顧客名でグループ化して候補を作成
    const clientMap = {};

    allData.forEach(item => {
        if (!item.clientName) return;

        const clientName = item.clientName;
        const clientNameLower = clientName.toLowerCase();
        const clientNormalized = normalizeCustomerName(clientName);

        // 曖昧検索：部分一致
        if (clientNameLower.includes(queryLower) || clientNormalized.includes(queryNormalized)) {
            if (!clientMap[clientNormalized]) {
                clientMap[clientNormalized] = {
                    clientName: clientName,
                    normalizedName: clientNormalized,
                    customers: new Set(),
                    count: 0
                };
            }
            clientMap[clientNormalized].customers.add(item.customerName);
            clientMap[clientNormalized].count++;
        }
    });

    // 配列に変換して件数の多い順にソート（最大10件）
    return Object.values(clientMap)
        .sort((a, b) => b.count - a.count)
        .slice(0, 10);
}

function showSuggestions(suggestions) {
    const container = document.getElementById('clientSuggestions');
    if (!container) return;

    if (suggestions.length === 0) {
        container.innerHTML = '<div class="client-suggestion-item"><span class="text-gray-400">該当なし</span></div>';
        container.classList.remove('hidden');
        return;
    }

    container.innerHTML = suggestions.map((s, index) => {
        const customerList = Array.from(s.customers).slice(0, 3).join(', ');
        const moreCount = s.customers.size > 3 ? ` 他${s.customers.size - 3}件` : '';
        return `
            <div class="client-suggestion-item" data-index="${index}" data-name="${s.clientName}" data-normalized="${s.normalizedName}">
                <div class="client-suggestion-name">${s.clientName}<span class="client-suggestion-count">(${s.count}件)</span></div>
                <div class="client-suggestion-info">得意先: ${customerList}${moreCount}</div>
            </div>
        `;
    }).join('');

    container.classList.remove('hidden');

    // クリックイベントを追加
    container.querySelectorAll('.client-suggestion-item').forEach(item => {
        item.addEventListener('click', () => selectClient(item));
    });
}

function selectClient(item) {
    const clientName = item.dataset.name;
    const normalizedName = item.dataset.normalized;

    if (!clientName) return;

    // 入力欄に選択した顧客名をセット
    const input = document.getElementById('clientNameSearch');
    if (input) {
        input.value = clientName;
    }

    selectedClientName = normalizedName;
    hideSuggestions();

    // ヘッダーに顧客名を表示
    updateHeaderCustomerName(clientName);

    showToast(`「${clientName}」を選択しました`, 'info');
}

// ヘッダーの顧客名を更新
function updateHeaderCustomerName(name) {
    const headerEl = document.getElementById('headerCustomerName');
    if (headerEl) {
        headerEl.textContent = name ? `📍 ${name}` : '';
    }
}

function hideSuggestions() {
    const container = document.getElementById('clientSuggestions');
    if (container) {
        container.classList.add('hidden');
    }
}

// ========================================
// 担当者名オートコンプリート（部署横断）
// ========================================

function onRepNameInput(e) {
    const query = e.target.value.trim();

    if (query.length < 1) {
        hideRepSuggestions();
        selectedRepName = null;
        return;
    }

    const suggestions = searchReps(query);
    showRepSuggestions(suggestions);
}

function searchReps(query) {
    if (!allData.length) return [];

    const queryLower = query.toLowerCase();
    const abbr = document.getElementById('abbrFilter').value;

    // 略称でフィルタリング（選択されている場合）
    let filteredData = allData;
    if (abbr) {
        filteredData = allData.filter(d => d.abbr === abbr);
    }

    // 担当者でグループ化して候補を作成
    const repMap = {};

    filteredData.forEach(item => {
        if (!item.repName) return;

        const repName = item.repName;
        const repLastName = extractLastName(repName);
        const repNameLower = repName.toLowerCase();

        // 曖昧検索：部分一致
        if (repNameLower.includes(queryLower) || repLastName.toLowerCase().includes(queryLower)) {
            if (!repMap[repLastName]) {
                repMap[repLastName] = {
                    repName: repName,
                    repLastName: repLastName,
                    branches: new Set(),
                    count: 0
                };
            }
            repMap[repLastName].branches.add(item.branch || '不明');
            repMap[repLastName].count++;
        }
    });

    // 配列に変換して件数の多い順にソート（最大10件）
    return Object.values(repMap)
        .sort((a, b) => b.count - a.count)
        .slice(0, 10);
}

function showRepSuggestions(suggestions) {
    const container = document.getElementById('repSuggestions');
    if (!container) return;

    if (suggestions.length === 0) {
        container.innerHTML = '<div class="client-suggestion-item"><span class="text-gray-400">該当なし</span></div>';
        container.classList.remove('hidden');
        return;
    }

    container.innerHTML = suggestions.map((s, index) => {
        const branchList = Array.from(s.branches).slice(0, 3).join(', ');
        const moreCount = s.branches.size > 3 ? ` 他${s.branches.size - 3}件` : '';
        return `
            <div class="client-suggestion-item" data-index="${index}" data-name="${s.repLastName}" data-fullname="${s.repName}">
                <div class="client-suggestion-name">${s.repLastName}<span class="client-suggestion-count">(${s.count}件)</span></div>
                <div class="client-suggestion-info">部署: ${branchList}${moreCount}</div>
            </div>
        `;
    }).join('');

    container.classList.remove('hidden');

    // クリックイベントを追加
    container.querySelectorAll('.client-suggestion-item').forEach(item => {
        item.addEventListener('click', () => selectRep(item));
    });
}

function selectRep(item) {
    const repName = item.dataset.name;
    const fullName = item.dataset.fullname;

    if (!repName) return;

    // 入力欄に選択した担当者名をセット
    const input = document.getElementById('repNameSearch');
    if (input) {
        input.value = repName;
    }

    selectedRepName = repName;
    hideRepSuggestions();

    showToast(`「${repName}」を選択しました`, 'info');
}

function hideRepSuggestions() {
    const container = document.getElementById('repSuggestions');
    if (container) {
        container.classList.add('hidden');
    }
}

// ========================================
// 検索・集計処理
// ========================================

function executeSearch() {
    if (!allData.length) {
        showToast('先にデータを読み込んでください', 'error');
        return;
    }

    // 年月フィルター取得（selectプルダウンはYYYY/MM形式）
    const startYM = document.getElementById('startYearMonth').value; // "2025/09" など
    const endYM = document.getElementById('endYearMonth').value;
    const abbr = document.getElementById('abbrFilter').value;
    const branch = document.getElementById('branchFilter').value;
    const hqOnly = document.getElementById('hqOnlyFilter').checked;

    // フィルタリング
    let filtered = allData;

    // 年月範囲フィルタ
    if (startYM || endYM) {
        filtered = filtered.filter(d => {
            if (!d.registYearMonth) return false;

            // 開始年月チェック
            if (startYM && d.registYearMonth < startYM) {
                return false;
            }
            // 終了年月チェック
            if (endYM && d.registYearMonth > endYM) {
                return false;
            }
            return true;
        });
    }

    if (abbr) {
        filtered = filtered.filter(d => d.abbr === abbr);
    }
    if (branch) {
        filtered = filtered.filter(d => d.branch === branch);
    }
    if (hqOnly) {
        filtered = filtered.filter(d => d.hqFlag === true);
    }

    // 顧客名フィルタ（横断検索）
    if (selectedClientName) {
        filtered = filtered.filter(d => {
            const normalized = normalizeCustomerName(d.clientName);
            return normalized === selectedClientName;
        });
    }

    // 担当者名フィルタ（横断検索）
    if (selectedRepName) {
        filtered = filtered.filter(d => {
            const repLastName = extractLastName(d.repName);
            return repLastName === selectedRepName;
        });
    }

    if (filtered.length === 0) {
        showToast('該当するデータがありません', 'info');
        return;
    }

    // 状態リセット
    currentView = 'rep';
    currentRepFilter = null;
    currentClientFilter = null;

    // 第1階層: 担当者別に集計
    const aggregated = filtered.reduce((acc, item) => {
        const repLastName = extractLastName(item.repName);
        const key = repLastName;

        if (!acc[key]) {
            acc[key] = {
                repLastName: repLastName,
                repFullName: item.repName || '',
                customerName: item.customerName || '',
                clientName: item.clientName || '',
                abbr: item.abbr || '',  // 略称
                branch: item.branch || '',  // TEL紐づけ用（部店）
                totalAmount: 0,
                items: []
            };
        }

        acc[key].totalAmount += item.unitPrice;
        acc[key].items.push(item);

        return acc;
    }, {});

    // 配列に変換
    repSummaryData = Object.values(aggregated);

    // デフォルトソート（合計/単価の降順）
    currentSort = { field: 'totalAmount', direction: 'desc' };
    sortAndRenderRepTable();

    showToast(`${repSummaryData.length}名の担当者`, 'success');
}

// ========================================
// 第1階層: 担当者別テーブル表示
// ========================================

function sortAndRenderRepTable() {
    repSummaryData.sort((a, b) => {
        let valA = a[currentSort.field];
        let valB = b[currentSort.field];

        if (typeof valA === 'string') {
            valA = valA.toLowerCase();
            valB = valB.toLowerCase();
        }

        if (currentSort.direction === 'asc') {
            return valA > valB ? 1 : valA < valB ? -1 : 0;
        } else {
            return valA < valB ? 1 : valA > valB ? -1 : 0;
        }
    });

    renderRepSummaryTable();
}

function renderRepSummaryTable() {
    const tbody = document.getElementById('summaryTableBody');
    const cards = document.getElementById('summaryCards');

    // テーブルヘッダーを担当者用に更新（略称・部署・担当者・合計/単価・TEL）
    const thead = document.querySelector('#summaryTable thead tr');
    thead.innerHTML = `
        <th class="px-2 py-3 rounded-tl-lg text-sm">略称</th>
        <th class="px-2 py-3 text-sm">部署</th>
        <th class="px-3 py-3 cursor-pointer hover:bg-gray-600" data-sort="repLastName">担当者 <span class="sort-icon">⇅</span></th>
        <th class="px-3 py-3 cursor-pointer hover:bg-gray-600 text-right" data-sort="totalAmount">合計/単価 <span class="sort-icon">⇅</span></th>
        <th class="px-2 py-3 rounded-tr-lg text-sm">TEL</th>
    `;

    // ソートイベントを再登録
    thead.querySelectorAll('th[data-sort]').forEach(th => {
        th.addEventListener('click', () => sortTable(th.dataset.sort));
    });

    if (repSummaryData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" class="px-4 py-8 text-center text-gray-500">データがありません</td></tr>';
        cards.innerHTML = '<div class="text-center text-gray-500 py-8">データがありません</div>';
        return;
    }

    // テーブル表示（デスクトップ）: 略称・部署・担当者・合計/単価・TEL
    tbody.innerHTML = repSummaryData.map((row, index) => {
        const phone = customerPhones[row.branch] || '-';
        return `
            <tr class="clickable-row hover:bg-gray-700/50 transition-colors" data-index="${index}">
                <td class="px-2 py-3 text-gray-300 text-xs">${row.abbr || '-'}</td>
                <td class="px-2 py-3 text-gray-400 text-xs">${row.branch || '-'}</td>
                <td class="px-3 py-3 font-medium text-blue-400">${row.repLastName || '-'}</td>
                <td class="px-3 py-3 text-right amount ${row.totalAmount < 0 ? 'negative' : ''}">${formatCurrency(row.totalAmount)}</td>
                <td class="px-2 py-3 text-gray-400 text-xs">${phone}</td>
            </tr>
        `;
    }).join('');

    // カード表示（スマホ）: 略称・部署・担当者・合計/単価・TEL
    cards.innerHTML = repSummaryData.map((row, index) => {
        const phone = customerPhones[row.branch] || '-';
        return `
            <div class="summary-card clickable" data-index="${index}">
                <div class="card-row">
                    <span class="card-label">略称</span>
                    <span class="card-value text-gray-300 text-sm">${row.abbr || '-'}</span>
                </div>
                <div class="card-row">
                    <span class="card-label">部署</span>
                    <span class="card-value text-gray-400 text-sm">${row.branch || '-'}</span>
                </div>
                <div class="card-row">
                    <span class="card-label">担当者</span>
                    <span class="card-value client-name">${row.repLastName || '-'}</span>
                </div>
                <div class="card-row">
                    <span class="card-label">合計/単価</span>
                    <span class="card-value highlight ${row.totalAmount < 0 ? 'negative' : ''}">${formatCurrency(row.totalAmount)}</span>
                </div>
                <div class="card-row">
                    <span class="card-label">TEL</span>
                    <span class="card-value">${phone}</span>
                </div>
            </div>
        `;
    }).join('');

    // クリックイベント追加（担当者→顧客へドリルダウン）
    tbody.querySelectorAll('.clickable-row').forEach(row => {
        row.addEventListener('click', () => showClientsByRep(parseInt(row.dataset.index)));
    });
    cards.querySelectorAll('.summary-card').forEach(card => {
        card.addEventListener('click', () => showClientsByRep(parseInt(card.dataset.index)));
    });
}

// ========================================
// 第2階層: 顧客別表示
// ========================================

async function showClientsByRep(index) {
    const repData = repSummaryData[index];
    if (!repData) return;

    // 詳細データ未取得（サマリのみ）の場合は取得する
    // items配列の中に isSummary=true のものがあれば未取得とみなす
    const hasSummaryOnly = repData.items.some(item => item.isSummary);

    if (hasSummaryOnly) {
        try {
            showToast(`${repData.repLastName}の詳細データを取得中...`, 'info');

            // 担当者名（フルネーム）で検索
            // repFullNameがない場合はrepLastNameを使う等のフォールバック
            const targetName = repData.repFullName || repData.repLastName;
            const response = await fetch(`${API_URL}?action=getOrderDetailsByRep&repName=${encodeURIComponent(targetName)}`);
            const result = await response.json();

            if (!result.success) throw new Error(result.error || '詳細データの取得に失敗しました');

            // 既存のサマリ行を削除し、詳細データを追加
            // 全データから、この担当者のサマリ行を除外（名前が一致、かつisSummaryフラグあり）
            const targetRepName = repData.repFullName; /* executeSearchでrepFullNameはセットされているはず */
            allData = allData.filter(d => !(d.repName === targetRepName && d.isSummary));

            // 詳細データを追加
            allData = allData.concat(result.data);

            // IndexedDBキャッシュも更新しておく（次回のために）
            const cacheKey = getCurrentCacheKey();
            await setToIndexedDB(cacheKey, allData);

            // 検索・集計を再実行（これで repSummaryData が更新される）
            executeSearch();

            // 更新された repSummaryData から同じ担当者を探す
            // 並び順が変わる可能性があるので名前で検索
            const newIndex = repSummaryData.findIndex(d => d.repLastName === repData.repLastName);
            if (newIndex !== -1) {
                showClientsByRep(newIndex); // 再帰呼び出し
            } else {
                showToast('データの更新後に担当者が見つかりませんでした', 'error');
            }
            return;

        } catch (error) {
            showToast(error.message, 'error');
            return;
        }
    }

    currentView = 'client';
    currentRepFilter = repData.repLastName;

    document.getElementById('summarySection').classList.add('hidden');
    document.getElementById('detailSection').classList.remove('hidden');

    // タイトル更新
    document.querySelector('#detailTitle .text-blue-400').textContent = repData.repLastName;

    // 顧客別に集計
    const aggregated = repData.items.reduce((acc, item) => {
        const clientNormalized = normalizeCustomerName(item.clientName);
        const key = clientNormalized;

        if (!acc[key]) {
            acc[key] = {
                clientName: item.clientName,
                clientNormalized: clientNormalized,
                totalAmount: 0,
                items: []
            };
        }

        acc[key].totalAmount += item.unitPrice;
        acc[key].items.push(item);

        return acc;
    }, {});

    clientSummaryData = Object.values(aggregated).sort((a, b) => b.totalAmount - a.totalAmount);

    renderClientTable();
}

function renderClientTable() {
    // テーブルヘッダーを顧客用に更新
    const thead = document.querySelector('#detailTable thead tr');
    thead.innerHTML = `
        <th class="px-4 py-3 rounded-tl-lg">顧客名</th>
        <th class="px-4 py-3 rounded-tr-lg text-right">合計/単価</th>
    `;

    const tbody = document.getElementById('detailTableBody');
    const cards = document.getElementById('detailCards');

    tbody.innerHTML = clientSummaryData.map((row, index) => `
        <tr class="clickable-row hover:bg-gray-700/50 transition-colors" data-index="${index}">
            <td class="px-4 py-3 text-blue-400 font-medium">${row.clientName || '-'}</td>
            <td class="px-4 py-3 text-right amount ${row.totalAmount < 0 ? 'negative' : ''}">${formatCurrency(row.totalAmount)}</td>
        </tr>
    `).join('');

    cards.innerHTML = clientSummaryData.map((row, index) => `
        <div class="detail-card clickable" data-index="${index}">
            <div class="card-row">
                <span class="card-label">顧客名</span>
                <span class="card-value client-name">${row.clientName || '-'}</span>
            </div>
            <div class="card-row">
                <span class="card-label">合計/単価</span>
                <span class="card-value highlight ${row.totalAmount < 0 ? 'negative' : ''}">${formatCurrency(row.totalAmount)}</span>
            </div>
        </div>
    `).join('');

    // クリックイベント追加（顧客→品番へドリルダウン）
    tbody.querySelectorAll('.clickable-row').forEach(row => {
        row.addEventListener('click', () => showProductsByClient(parseInt(row.dataset.index)));
    });
    cards.querySelectorAll('.detail-card').forEach(card => {
        card.addEventListener('click', () => showProductsByClient(parseInt(card.dataset.index)));
    });
}

// ========================================
// 第3階層: 品番別表示
// ========================================

function showProductsByClient(index) {
    const clientData = clientSummaryData[index];
    if (!clientData) return;

    currentView = 'product';
    currentClientFilter = clientData.clientNormalized;

    // タイトル更新
    document.querySelector('#detailTitle .text-blue-400').textContent =
        `${currentRepFilter} > ${clientData.clientName}`;

    // 品番別に集計
    const aggregated = clientData.items.reduce((acc, item) => {
        const key = item.productCode || 'その他';

        if (!acc[key]) {
            acc[key] = {
                productCode: item.productCode,
                productName: item.productName,
                totalQuantity: 0,
                totalAmount: 0
            };
        }

        acc[key].totalQuantity += item.quantity;
        acc[key].totalAmount += item.unitPrice;

        return acc;
    }, {});

    productSummaryData = Object.values(aggregated).sort((a, b) => b.totalAmount - a.totalAmount);

    renderProductTable();
}

function renderProductTable() {
    // テーブルヘッダーを品番用に更新
    const thead = document.querySelector('#detailTable thead tr');
    thead.innerHTML = `
        <th class="px-4 py-3 rounded-tl-lg">品番</th>
        <th class="px-4 py-3">商品名</th>
        <th class="px-4 py-3 text-right">受注数量</th>
        <th class="px-4 py-3 rounded-tr-lg text-right">合計/単価</th>
    `;

    const tbody = document.getElementById('detailTableBody');
    const cards = document.getElementById('detailCards');

    tbody.innerHTML = productSummaryData.map(row => `
        <tr class="hover:bg-gray-700/50 transition-colors">
            <td class="px-4 py-3 font-mono text-sm">${row.productCode || '-'}</td>
            <td class="px-4 py-3">${hankakuToZenkakuKana(row.productName) || '-'}</td>
            <td class="px-4 py-3 text-right">${row.totalQuantity}</td>
            <td class="px-4 py-3 text-right amount ${row.totalAmount < 0 ? 'negative' : ''}">${formatCurrency(row.totalAmount)}</td>
        </tr>
    `).join('');

    cards.innerHTML = productSummaryData.map(row => `
        <div class="detail-card">
            <div class="card-row">
                <span class="card-label">品番</span>
                <span class="card-value font-mono">${row.productCode || '-'}</span>
            </div>
            <div class="card-row">
                <span class="card-label">商品名</span>
                <span class="card-value">${hankakuToZenkakuKana(row.productName) || '-'}</span>
            </div>
            <div class="card-row">
                <span class="card-label">数量</span>
                <span class="card-value">${row.totalQuantity}</span>
            </div>
            <div class="card-row">
                <span class="card-label">合計</span>
                <span class="card-value highlight ${row.totalAmount < 0 ? 'negative' : ''}">${formatCurrency(row.totalAmount)}</span>
            </div>
        </div>
    `).join('');
}

// ========================================
// 戻るボタン処理
// ========================================

function handleBackButton() {
    if (currentView === 'product') {
        // 品番→顧客に戻る
        currentView = 'client';
        document.querySelector('#detailTitle .text-blue-400').textContent = currentRepFilter;
        renderClientTable();
    } else if (currentView === 'client') {
        // 顧客→担当者に戻る
        currentView = 'rep';
        currentRepFilter = null;
        document.getElementById('detailSection').classList.add('hidden');
        document.getElementById('summarySection').classList.remove('hidden');
    } else {
        document.getElementById('detailSection').classList.add('hidden');
        document.getElementById('summarySection').classList.remove('hidden');
    }
}

// ========================================
// ソート処理
// ========================================

function sortTable(field) {
    if (currentSort.field === field) {
        currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
    } else {
        currentSort.field = field;
        currentSort.direction = 'desc';
    }

    // ソートアイコン更新
    document.querySelectorAll('#summaryTable th[data-sort]').forEach(th => {
        th.classList.remove('sorted-asc', 'sorted-desc');
        if (th.dataset.sort === field) {
            th.classList.add(currentSort.direction === 'asc' ? 'sorted-asc' : 'sorted-desc');
        }
    });

    // ソート実行
    if (currentView === 'rep') {
        sortAndRenderRepTable();
    }
}

// ========================================
// PDF出力
// ========================================

async function exportToPdf(type) {
    const { jsPDF } = window.jspdf;

    // 対象のテーブルを特定
    const tableId = type === 'summary' ? 'summaryTable' : 'detailTable';
    const table = document.getElementById(tableId);

    if (!table) {
        showToast('出力対象が見つかりません', 'error');
        return;
    }

    showToast('PDF生成中...', 'info');

    try {
        // html2canvasでキャプチャ
        const canvas = await html2canvas(table, {
            scale: 2,
            backgroundColor: '#1f2937',
            logging: false
        });

        const imgData = canvas.toDataURL('image/png');
        const pdf = new jsPDF({
            orientation: canvas.width > canvas.height ? 'l' : 'p',
            unit: 'mm',
            format: 'a4'
        });

        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const imgWidth = pageWidth - 20;
        const imgHeight = (canvas.height * imgWidth) / canvas.width;

        pdf.addImage(imgData, 'PNG', 10, 10, imgWidth, Math.min(imgHeight, pageHeight - 20));

        const fileName = type === 'summary' ? '売上分析_サマリー.pdf' : '売上分析_詳細.pdf';
        pdf.save(fileName);

        showToast('PDFを出力しました', 'success');
    } catch (error) {
        console.error('PDF export error:', error);
        showToast('PDF出力に失敗しました', 'error');
    }
}

// ========================================
// ユーティリティ
// ========================================

function showToast(message, type = 'info') {
    const toast = document.getElementById('toast');
    toast.textContent = message;
    toast.className = `fixed bottom-4 right-4 px-6 py-3 rounded-lg shadow-lg z-50 toast show ${type}`;

    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}
