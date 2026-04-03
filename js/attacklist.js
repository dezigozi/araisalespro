// ========================================
// 設定
// ========================================
const GAS_URL = GAS_WEBAPP_URL;

// ========================================
// 状態管理
// ========================================
let actionList = [];
let filteredActionList = [];
let currentYear = new Date().getFullYear();
let currentMonth = new Date().getMonth() + 1;

// ========================================
// 初期化
// ========================================
document.addEventListener('DOMContentLoaded', async () => {
    setupEventListeners();
    updateMonthDisplay();
    await loadActionList();
});

// ========================================
// イベントリスナー設定
// ========================================
function setupEventListeners() {
    // 年月ナビゲーション
    document.getElementById('prevMonth').addEventListener('click', async () => {
        currentMonth--;
        if (currentMonth < 1) {
            currentMonth = 12;
            currentYear--;
        }
        updateMonthDisplay();
        await loadActionList();
    });

    document.getElementById('nextMonth').addEventListener('click', async () => {
        currentMonth++;
        if (currentMonth > 12) {
            currentMonth = 1;
            currentYear++;
        }
        updateMonthDisplay();
        await loadActionList();
    });

    // フィルター
    document.getElementById('salesRepFilter').addEventListener('change', applyFilters);
    document.getElementById('statusFilter').addEventListener('change', applyFilters);
    document.getElementById('sortBy').addEventListener('change', applyFilters);
}

// ========================================
// 月表示更新
// ========================================
function updateMonthDisplay() {
    const yearMonth = `${currentYear}年${currentMonth}月`;
    document.getElementById('targetMonth').textContent = yearMonth;
}

// ========================================
// 年月文字列取得
// ========================================
function getCurrentYearMonth() {
    return `${currentYear}年${currentMonth}月`;
}

// ========================================
// アクションリスト読み込み
// ========================================
async function loadActionList() {
    const tbody = document.getElementById('actionListTableBody');
    tbody.innerHTML = '<tr><td colspan="9" class="loading">読み込み中...</td></tr>';

    try {
        const yearMonth = getCurrentYearMonth();
        const result = await fetchAPI('getActionList', { yearMonth: yearMonth });

        if (result?.success) {
            actionList = result.data || [];
            populateSalesRepFilter();
            applyFilters();
            updateProgressCount();
            toggleEmptyGuide();
        } else {
            tbody.innerHTML = '<tr><td colspan="9" class="empty-state">読み込みに失敗しました</td></tr>';
        }
    } catch (error) {
        console.error('Error loading action list:', error);
        tbody.innerHTML = '<tr><td colspan="9" class="empty-state">読み込みに失敗しました</td></tr>';
    }
}

// ========================================
// 営業担当フィルター生成
// ========================================
function populateSalesRepFilter() {
    const select = document.getElementById('salesRepFilter');
    const currentValue = select.value;

    // 既存のオプションをクリア（最初の「全員」以外）
    while (select.options.length > 1) {
        select.remove(1);
    }

    // アクションリストから営業担当を抽出
    const reps = [...new Set(actionList.map(a => a.salesRep).filter(r => r))];
    reps.sort().forEach(rep => {
        const option = document.createElement('option');
        option.value = rep;
        option.textContent = rep;
        select.appendChild(option);
    });

    // 以前の選択値を復元
    if (currentValue && reps.includes(currentValue)) {
        select.value = currentValue;
    }
}

// ========================================
// フィルター・ソート適用
// ========================================
function applyFilters() {
    const salesRepFilter = document.getElementById('salesRepFilter').value;
    const statusFilter = document.getElementById('statusFilter').value;
    const sortBy = document.getElementById('sortBy').value;

    // フィルター適用
    filteredActionList = actionList.filter(action => {
        // 営業担当フィルター
        if (salesRepFilter && action.salesRep !== salesRepFilter) {
            return false;
        }
        // ステータスフィルター
        if (statusFilter && action.status !== statusFilter) {
            return false;
        }
        return true;
    });

    // ソート適用
    filteredActionList.sort((a, b) => {
        switch (sortBy) {
            case 'daysSince':
                // 経過日数が多い順（nullは最後）
                if (a.daysSince === null && b.daysSince === null) return 0;
                if (a.daysSince === null) return -1;
                if (b.daysSince === null) return 1;
                return b.daysSince - a.daysSince;
            case 'daysSinceAsc':
                // 経過日数が少ない順（nullは最後）
                if (a.daysSince === null && b.daysSince === null) return 0;
                if (a.daysSince === null) return 1;
                if (b.daysSince === null) return -1;
                return a.daysSince - b.daysSince;
            case 'company':
                return (a.company || '').localeCompare(b.company || '');
            case 'status':
                const statusOrder = { 'pending': 0, 'in_progress': 1, 'completed': 2, 'skip': 3 };
                return (statusOrder[a.status] || 0) - (statusOrder[b.status] || 0);
            default:
                return 0;
        }
    });

    renderActionList();
}

// ========================================
// アクションリスト描画
// ========================================
function renderActionList() {
    const tbody = document.getElementById('actionListTableBody');

    if (!filteredActionList || filteredActionList.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" class="empty-state">該当するデータがありません</td></tr>';
        return;
    }

    tbody.innerHTML = filteredActionList.map(action => {
        const lastDate = action.lastVisitDate || '未訪問';
        const daysText = action.daysSince !== null ? `${action.daysSince}日` : '-';
        const visitStatus = action.visitStatus || 'distant';
        const telLink = action.tel ? `<a href="tel:${action.tel}" class="tel-link">📞 ${action.tel}</a>` : 'ー';

        // ステータスバッジ
        const statusLabels = {
            'pending': '未着手',
            'in_progress': '進行中',
            'completed': '完了',
            'skip': 'スキップ'
        };
        const statusLabel = statusLabels[action.status] || '未着手';
        const statusClass = action.status || 'pending';

        // 活動記録ページへのリンク
        const params = new URLSearchParams({
            company: action.company,
            department: action.department || '',
            contact: action.contactName || '',
            fromActionList: 'true'
        });

        return `
            <tr data-id="${action.id}" data-year-month="${action.yearMonth}">
                <td class="col-company">${action.company}</td>
                <td class="col-dept">${action.department || '-'}</td>
                <td class="col-contact">${action.contactName || '-'}</td>
                <td class="col-lastvisit">${lastDate}</td>
                <td class="col-days"><span class="days-cell ${visitStatus}">${daysText}</span></td>
                <td class="col-rep">${action.salesRep || '-'}</td>
                <td class="col-tel">${telLink}</td>
                <td class="col-status">
                    <select class="status-select ${statusClass}" onchange="updateStatus('${action.yearMonth}', '${action.id}', this.value)">
                        <option value="pending" ${action.status === 'pending' ? 'selected' : ''}>未着手</option>
                        <option value="in_progress" ${action.status === 'in_progress' ? 'selected' : ''}>進行中</option>
                        <option value="completed" ${action.status === 'completed' ? 'selected' : ''}>完了</option>
                        <option value="skip" ${action.status === 'skip' ? 'selected' : ''}>スキップ</option>
                    </select>
                </td>
                <td class="col-action">
                    <a href="index.html?${params.toString()}" class="activity-link">活動記録 →</a>
                </td>
            </tr>
        `;
    }).join('');
}

// ========================================
// ステータス更新
// ========================================
async function updateStatus(yearMonth, contactId, newStatus) {
    try {
        // ローカルのデータを即時更新
        const action = actionList.find(a => String(a.id) === String(contactId) && a.yearMonth === yearMonth);
        if (action) {
            action.status = newStatus;
        }

        // UIを更新
        const select = document.querySelector(`tr[data-id="${contactId}"][data-year-month="${yearMonth}"] .status-select`);
        if (select) {
            select.className = `status-select ${newStatus}`;
        }

        updateProgressCount();

        // GASに送信（no-corsモード）
        await fetch(GAS_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                action: 'updateActionStatus',
                yearMonth: yearMonth,
                contactId: contactId,
                status: newStatus
            })
        });

        showToast('ステータスを更新しました');
    } catch (error) {
        console.error('Error updating status:', error);
        showToast('ステータスの更新に失敗しました', true);
    }
}

// ========================================
// 進捗カウント更新
// ========================================
function updateProgressCount() {
    const total = actionList.length;
    const completed = actionList.filter(a => a.status === 'completed').length;
    document.getElementById('progressCount').textContent = `${completed}/${total}`;
}

// ========================================
// 空の場合のガイド表示切替
// ========================================
function toggleEmptyGuide() {
    const guide = document.getElementById('emptyGuide');
    const table = document.querySelector('.table-container');

    if (actionList.length === 0) {
        guide.classList.remove('hidden');
        table.classList.add('hidden');
    } else {
        guide.classList.add('hidden');
        table.classList.remove('hidden');
    }
}

// ========================================
// API通信（GET）
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
