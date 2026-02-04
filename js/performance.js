// ========================================
// 過去実績ページ
// ========================================

// GAS WebアプリURL
const GAS_URL = 'https://script.google.com/macros/s/AKfycbwrd5rm_eKQGJgON83RW8qg5H0SkMkqk6Zmrwh-lM62cqG6he9Ugq-7vmN0wXaaj-a3Nw/exec';

// キャッシュ設定
const PERFORMANCE_CACHE_KEY = 'sfa_performance_cache';
const PERFORMANCE_RAW_CACHE_KEY = 'sfa_performance_raw_cache';
const CACHE_DURATION = 24 * 60 * 60 * 1000; // 24時間

// データ格納
let performanceData = [];
let rawPerformanceData = []; // 詳細分析用の生データ
let filteredData = [];
let currentSort = { field: 'orderYearMonth', direction: 'desc' };
let dataLoaded = false; // データが読み込まれたかどうか

// ========================================
// キャッシュ管理
// ========================================
const PerformanceCacheManager = {
    get(key) {
        try {
            const cached = localStorage.getItem(key);
            if (!cached) return null;
            const { data, timestamp } = JSON.parse(cached);
            if (Date.now() - timestamp > CACHE_DURATION) {
                this.clear(key);
                return null;
            }
            return data;
        } catch (e) {
            return null;
        }
    },
    set(key, data) {
        try {
            localStorage.setItem(key, JSON.stringify({ data, timestamp: Date.now() }));
        } catch (e) { }
    },
    clear(key) {
        localStorage.removeItem(key);
    },
    clearAll() {
        localStorage.removeItem(PERFORMANCE_CACHE_KEY);
        localStorage.removeItem(PERFORMANCE_RAW_CACHE_KEY);
    }
};

// ========================================
// 初期化
// ========================================
document.addEventListener('DOMContentLoaded', async () => {
    if (!checkAuth()) return;

    // イベントリスナー設定
    setupEventListeners();

    // 初期状態ではデータを取得しない（検索ボタン押下時に取得）
    // テーブルは「検索してください」のメッセージを表示
});

// ========================================
// イベントリスナー設定
// ========================================
function setupEventListeners() {
    // 検索ボタン
    document.getElementById('searchBtn').addEventListener('click', handleSearch);

    // クリアボタン
    document.getElementById('clearBtn').addEventListener('click', clearFilters);

    // 得意先名変更時に担当者フィルタを更新
    document.getElementById('customerFilter').addEventListener('change', updateRepFilter);

    // ソートヘッダー
    document.querySelectorAll('.sortable').forEach(th => {
        th.addEventListener('click', () => {
            const field = th.dataset.sort;
            if (currentSort.field === field) {
                currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
            } else {
                currentSort.field = field;
                currentSort.direction = 'asc';
            }
            updateSortIndicators();
            sortAndRender();
        });
    });

    // PDF出力
    document.getElementById('exportPdfBtn').addEventListener('click', exportPdf);

    // CSV出力
    document.getElementById('exportCsvBtn').addEventListener('click', exportCsv);

    // モーダル閉じる
    document.getElementById('closeModal').addEventListener('click', closeModal);
    document.getElementById('detailModal').addEventListener('click', (e) => {
        if (e.target.id === 'detailModal') closeModal();
    });

    // モーダルタブ切替
    document.querySelectorAll('.modal-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            document.querySelectorAll('.modal-tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.modal-tab-content').forEach(c => c.classList.remove('active'));
            tab.classList.add('active');
            document.getElementById('tab-' + tab.dataset.tab).classList.add('active');
        });
    });

    // モーダルPDF表示
    document.getElementById('modalPdfBtn').addEventListener('click', exportModalPdf);
}

// ========================================
// 検索ボタン処理
// ========================================
async function handleSearch() {
    // データが未読み込みの場合は取得
    if (!dataLoaded) {
        await loadPerformanceData();
    } else {
        // データがある場合はフィルタのみ適用
        applyFilters();
    }
}

// ========================================
// データ取得
// ========================================
async function loadPerformanceData() {
    const tbody = document.getElementById('performanceTableBody');

    // キャッシュをチェック
    const cachedData = PerformanceCacheManager.get(PERFORMANCE_CACHE_KEY);
    const cachedRawData = PerformanceCacheManager.get(PERFORMANCE_RAW_CACHE_KEY);

    if (cachedData) {
        // キャッシュからデータを復元
        performanceData = cachedData;
        rawPerformanceData = cachedRawData || [];
        dataLoaded = true;
        populateFilters();
        applyFilters();
        showToast('⚡ キャッシュから読み込みました');
        return;
    }

    // キャッシュがない場合はAPIから取得
    tbody.innerHTML = '<tr><td colspan="6" class="loading">データを取得中...</td></tr>';
    showToast('📊 データを取得中...');

    try {
        const response = await fetch(`${GAS_URL}?action=getPerformanceData`);
        const result = await response.json();

        if (result.success) {
            performanceData = result.data;
            PerformanceCacheManager.set(PERFORMANCE_CACHE_KEY, performanceData);
            dataLoaded = true;
            populateFilters();
            applyFilters();
            showToast('✅ データを取得しました');
        } else {
            showToast('データ取得に失敗しました: ' + result.error, true);
            tbody.innerHTML = '<tr><td colspan="6" class="empty-row">データの取得に失敗しました</td></tr>';
        }

        // 詳細分析用の生データも取得
        const rawResponse = await fetch(`${GAS_URL}?action=getPerformanceRawData`);
        const rawResult = await rawResponse.json();
        if (rawResult.success) {
            rawPerformanceData = rawResult.data;
            PerformanceCacheManager.set(PERFORMANCE_RAW_CACHE_KEY, rawPerformanceData);
        }
    } catch (error) {
        console.error('Error loading performance data:', error);
        showToast('データ取得に失敗しました', true);
        tbody.innerHTML = '<tr><td colspan="6" class="empty-row">データの取得に失敗しました</td></tr>';
    }
}

// ========================================
// フィルタ選択肢を設定
// ========================================
function populateFilters() {
    const customerFilter = document.getElementById('customerFilter');

    // 得意先名の一覧
    const customers = [...new Set(performanceData.map(d => d.customerName))].filter(c => c).sort();
    customerFilter.innerHTML = '<option value="">すべて</option>';
    customers.forEach(c => {
        customerFilter.innerHTML += `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`;
    });

    // 担当者フィルタも初期化
    updateRepFilter();
}

// ========================================
// 得意先に紐づく担当者のみ表示
// ========================================
function updateRepFilter() {
    const customerFilter = document.getElementById('customerFilter');
    const repFilter = document.getElementById('repFilter');
    const selectedCustomer = customerFilter.value;

    // 選択された得意先に基づいて担当者をフィルタ
    let filteredData = performanceData;
    if (selectedCustomer) {
        filteredData = performanceData.filter(d => d.customerName === selectedCustomer);
    }

    const reps = [...new Set(filteredData.map(d => d.customerRep))].filter(r => r).sort();
    repFilter.innerHTML = '<option value="">すべて</option>';
    reps.forEach(r => {
        repFilter.innerHTML += `<option value="${escapeHtml(r)}">${escapeHtml(r)}</option>`;
    });
}

// ========================================
// フィルタ適用
// ========================================
function applyFilters() {
    const customerFilter = document.getElementById('customerFilter').value;
    const repFilter = document.getElementById('repFilter').value;

    filteredData = performanceData.filter(d => {
        if (customerFilter && d.customerName !== customerFilter) return false;
        if (repFilter && d.customerRep !== repFilter) return false;
        return true;
    });

    sortAndRender();
}

// ========================================
// フィルタクリア
// ========================================
function clearFilters() {
    document.getElementById('customerFilter').value = '';
    document.getElementById('repFilter').value = '';
    // キャッシュをクリア
    PerformanceCacheManager.clearAll();
    dataLoaded = false;
    performanceData = [];
    rawPerformanceData = [];
    filteredData = [];
    // テーブルを初期状態に戻す
    const tbody = document.getElementById('performanceTableBody');
    const tfoot = document.getElementById('performanceTableFoot');
    tbody.innerHTML = '<tr><td colspan="6" class="empty-row">🔍 検索条件を指定して「検索」ボタンを押してください</td></tr>';
    tfoot.innerHTML = '';
    // 担当者フィルタもリセット
    document.getElementById('repFilter').innerHTML = '<option value="">すべて</option>';
    showToast('🔄 キャッシュをクリアしました');
}

// ========================================
// ソートして描画
// ========================================
function sortAndRender() {
    const { field, direction } = currentSort;

    filteredData.sort((a, b) => {
        let valA = a[field];
        let valB = b[field];

        // 数値フィールドの場合
        if (field === 'orderCount' || field === 'salesAmount') {
            valA = Number(valA) || 0;
            valB = Number(valB) || 0;
        } else {
            valA = String(valA || '');
            valB = String(valB || '');
        }

        if (valA < valB) return direction === 'asc' ? -1 : 1;
        if (valA > valB) return direction === 'asc' ? 1 : -1;
        return 0;
    });

    renderTable();
}

// ========================================
// ソートインジケータ更新
// ========================================
function updateSortIndicators() {
    document.querySelectorAll('.sortable').forEach(th => {
        th.classList.remove('asc', 'desc');
        if (th.dataset.sort === currentSort.field) {
            th.classList.add(currentSort.direction);
        }
    });
}

// ========================================
// テーブル描画
// ========================================
function renderTable() {
    const tbody = document.getElementById('performanceTableBody');
    const tfoot = document.getElementById('performanceTableFoot');

    if (filteredData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-row">データがありません</td></tr>';
        tfoot.innerHTML = '';
        return;
    }

    // 集計
    let totalCount = 0;
    let totalAmount = 0;

    let html = '';
    filteredData.forEach(row => {
        totalCount += row.orderCount || 0;
        totalAmount += row.salesAmount || 0;

        html += `
            <tr>
                <td>${escapeHtml(row.orderYearMonth)}</td>
                <td>${escapeHtml(row.customerName)}</td>
                <td><span class="clickable-rep" data-rep="${escapeHtml(row.customerRep)}">${escapeHtml(row.customerRep)}</span></td>
                <td>${escapeHtml(row.clientName)}</td>
                <td class="number">${formatNumber(row.orderCount)}</td>
                <td class="number">¥${formatNumber(row.salesAmount)}</td>
            </tr>
        `;
    });

    tbody.innerHTML = html;

    // フッター（合計）
    tfoot.innerHTML = `
        <tr>
            <td colspan="4" style="text-align: right; font-weight: 700;">合計</td>
            <td class="number">${formatNumber(totalCount)}</td>
            <td class="number">¥${formatNumber(totalAmount)}</td>
        </tr>
    `;

    // 担当者名クリックイベント
    document.querySelectorAll('.clickable-rep').forEach(span => {
        span.addEventListener('click', () => {
            openDetailModal(span.dataset.rep);
        });
    });
}

// ========================================
// PDF出力（印刷ダイアログ）
// ========================================
function exportPdf() {
    window.print();
}

// ========================================
// CSV出力
// ========================================
function exportCsv() {
    if (filteredData.length === 0) {
        showToast('出力するデータがありません', true);
        return;
    }

    // BOM付きUTF-8でCSV作成
    const BOM = '\uFEFF';
    const headers = ['受注年月', '得意先名', '得意先担当者名', '顧客名_漢字', '受注件数', '売上金額'];

    let csv = BOM + headers.join(',') + '\n';

    filteredData.forEach(row => {
        const values = [
            row.orderYearMonth,
            row.customerName,
            row.customerRep,
            row.clientName,
            row.orderCount,
            row.salesAmount
        ].map(v => `"${String(v || '').replace(/"/g, '""')}"`);
        csv += values.join(',') + '\n';
    });

    // ダウンロード
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `過去実績_${formatDate(new Date())}.csv`;
    link.click();
    URL.revokeObjectURL(url);

    showToast('CSVファイルをダウンロードしました');
}

// ========================================
// ユーティリティ関数
// ========================================
function escapeHtml(text) {
    if (text == null) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function formatNumber(num) {
    if (num == null || isNaN(num)) return '0';
    return Number(num).toLocaleString('ja-JP');
}

function formatDate(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}${m}${d}`;
}

function showToast(message, isError = false) {
    const toast = document.getElementById('toast');
    toast.textContent = message;
    toast.className = 'toast' + (isError ? ' error' : '');
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 3000);
}

// ========================================
// モーダル関連
// ========================================
function openDetailModal(repName) {
    const modal = document.getElementById('detailModal');
    const modalTitle = document.getElementById('modalTitle');

    modalTitle.textContent = `${repName} の詳細分析`;
    modal.classList.add('show');

    // タブをリセット
    document.querySelectorAll('.modal-tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.modal-tab-content').forEach(c => c.classList.remove('active'));
    document.querySelector('.modal-tab[data-tab="navi"]').classList.add('active');
    document.getElementById('tab-navi').classList.add('active');

    // 詳細データを描画
    renderDetailTables(repName);
}

function closeModal() {
    document.getElementById('detailModal').classList.remove('show');
}

function renderDetailTables(repName) {
    const naviBody = document.getElementById('naviTableBody');
    const dorareBody = document.getElementById('dorareTableBody');
    const vehicleBody = document.getElementById('vehicleTableBody');

    // データがない場合
    if (!rawPerformanceData || rawPerformanceData.length === 0) {
        const noData = '<tr><td colspan="5" class="empty-row">データがありません（GASデプロイ後に表示されます）</td></tr>';
        naviBody.innerHTML = noData;
        dorareBody.innerHTML = noData;
        vehicleBody.innerHTML = '<tr><td colspan="4" class="empty-row">データがありません（GASデプロイ後に表示されます）</td></tr>';
        return;
    }

    // 担当者でフィルタ
    const repData = rawPerformanceData.filter(d => d.customerRep === repName);

    // ナビ: 商品大分類=2, 商品中分類=S, 商品小分類=C, メーカーCD≠9080
    const naviData = repData.filter(d =>
        d.productMajor == 2 &&
        d.productMiddle === 'S' &&
        d.productMinor === 'C' &&
        d.makerCode != '9080'
    );

    // ドラレコ: 商品大分類=2, 商品中分類=S, 商品小分類=Y
    const dorareData = repData.filter(d =>
        d.productMajor == 2 &&
        d.productMiddle === 'S' &&
        d.productMinor === 'Y'
    );

    // ナビ集計（顧客_漢字, 車種名, 商品名でグループ化）
    const naviAggregated = aggregateByProduct(naviData, repName);
    renderProductTable(naviBody, naviAggregated);

    // ドラレコ集計
    const dorareAggregated = aggregateByProduct(dorareData, repName);
    renderProductTable(dorareBody, dorareAggregated);

    // 車種集計（顧客_漢字, 車種名でグループ化）
    const vehicleAggregated = aggregateByVehicle(repData, repName);
    renderVehicleTable(vehicleBody, vehicleAggregated);
}

function aggregateByProduct(data, repName) {
    const aggregated = {};

    data.forEach(row => {
        const key = `${row.clientName}_${row.vehicleName}_${row.productCode}`;
        if (!aggregated[key]) {
            aggregated[key] = {
                customerRep: repName,
                clientName: row.clientName || '',
                vehicleName: row.vehicleName || '',
                productCode: row.productCode || '',
                orderCount: 0,
                processedOrders: new Set()
            };
        }
        // 代表受注伝票番号でカウント
        if (row.orderNo && !aggregated[key].processedOrders.has(row.orderNo)) {
            aggregated[key].orderCount += 1;
            aggregated[key].processedOrders.add(row.orderNo);
        }
    });

    // 受注件数降順でソート
    return Object.values(aggregated)
        .map(item => ({
            customerRep: item.customerRep,
            clientName: item.clientName,
            vehicleName: item.vehicleName,
            productCode: item.productCode,
            orderCount: item.orderCount
        }))
        .sort((a, b) => b.orderCount - a.orderCount);
}

function aggregateByVehicle(data, repName) {
    const aggregated = {};

    data.forEach(row => {
        const key = `${row.clientName}_${row.vehicleName}`;
        if (!aggregated[key]) {
            aggregated[key] = {
                customerRep: repName,
                clientName: row.clientName || '',
                vehicleName: row.vehicleName || '',
                orderCount: 0,
                processedOrders: new Set()
            };
        }
        if (row.orderNo && !aggregated[key].processedOrders.has(row.orderNo)) {
            aggregated[key].orderCount += 1;
            aggregated[key].processedOrders.add(row.orderNo);
        }
    });

    return Object.values(aggregated)
        .map(item => ({
            customerRep: item.customerRep,
            clientName: item.clientName,
            vehicleName: item.vehicleName,
            orderCount: item.orderCount
        }))
        .sort((a, b) => b.orderCount - a.orderCount);
}

function renderProductTable(tbody, data) {
    if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" class="empty-row">該当データがありません</td></tr>';
        return;
    }

    let html = '';
    data.forEach(row => {
        html += `
            <tr>
                <td>${escapeHtml(row.customerRep)}</td>
                <td>${escapeHtml(row.clientName)}</td>
                <td>${escapeHtml(row.vehicleName)}</td>
                <td>${escapeHtml(row.productCode)}</td>
                <td class="number">${formatNumber(row.orderCount)}</td>
            </tr>
        `;
    });
    tbody.innerHTML = html;
}

function renderVehicleTable(tbody, data) {
    if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" class="empty-row">該当データがありません</td></tr>';
        return;
    }

    let html = '';
    data.forEach(row => {
        html += `
            <tr>
                <td>${escapeHtml(row.customerRep)}</td>
                <td>${escapeHtml(row.clientName)}</td>
                <td>${escapeHtml(row.vehicleName)}</td>
                <td class="number">${formatNumber(row.orderCount)}</td>
            </tr>
        `;
    });
    tbody.innerHTML = html;
}

// ========================================
// モーダルPDF表示
// ========================================
function exportModalPdf() {
    // モーダルを印刷用に一時的にフルスクリーンに
    const modal = document.getElementById('detailModal');
    const modalContent = modal.querySelector('.modal-content');

    modal.classList.add('print-mode');

    window.print();

    modal.classList.remove('print-mode');
}
