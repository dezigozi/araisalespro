// ========================================
// 設定
// ========================================
const GAS_URL = 'https://script.google.com/macros/s/AKfycbwrd5rm_eKQGJgON83RW8qg5H0SkMkqk6Zmrwh-lM62cqG6he9Ugq-7vmN0wXaaj-a3Nw/exec';

// ========================================
// 状態管理
// ========================================
let currentYear = new Date().getFullYear();
let currentMonth = new Date().getMonth() + 1;
let activities = [];
let currentGoals = {};
let currentSummaryData = []; // CSV出力用
let currentVisitsData = [];  // CSV出力用
let selectedSalesRep = '';   // 担当者フィルター
let selectedResult = '';      // 結果フィルター

// ========================================
// 初期化
// ========================================
document.addEventListener('DOMContentLoaded', async () => {
    setupEventListeners();
    updateMonthDisplay();
    await loadDashboardData();
});

// ========================================
// イベントリスナー
// ========================================
function setupEventListeners() {
    document.getElementById('prevMonth').addEventListener('click', async () => {
        currentMonth--;
        if (currentMonth < 1) {
            currentMonth = 12;
            currentYear--;
        }
        updateMonthDisplay();
        await loadDashboardData();
    });

    document.getElementById('nextMonth').addEventListener('click', async () => {
        currentMonth++;
        if (currentMonth > 12) {
            currentMonth = 1;
            currentYear++;
        }
        updateMonthDisplay();
        await loadDashboardData();
    });

    // CSV出力ボタン
    document.getElementById('exportSummaryCsv').addEventListener('click', () => {
        exportSummaryCsv();
    });

    document.getElementById('exportVisitsCsv').addEventListener('click', () => {
        exportVisitsCsv();
    });

    // 担当者フィルター
    const salesRepFilter = document.getElementById('dashboardSalesRepSelect');
    if (salesRepFilter) {
        salesRepFilter.addEventListener('change', (e) => {
            selectedSalesRep = e.target.value;
            renderDashboard();
        });
    }

    // 結果フィルター
    const resultFilter = document.getElementById('dashboardResultSelect');
    if (resultFilter) {
        resultFilter.addEventListener('change', (e) => {
            selectedResult = e.target.value;
            renderDashboard();
        });
    }
}

// ========================================
// 月表示更新
// ========================================
function updateMonthDisplay() {
    document.getElementById('currentMonth').textContent = `${currentYear}年${currentMonth}月`;
}

// ========================================
// ダッシュボードデータ一括読み込み
// ========================================
async function loadDashboardData() {
    const [goalsResult, activitiesResult] = await Promise.all([
        loadGoals(),
        fetchAPI('getActivities')
    ]);

    if (goalsResult?.success) {
        currentGoals = goalsResult.data;
    } else {
        currentGoals = {};
    }

    if (activitiesResult?.success) {
        activities = activitiesResult.data;
    } else {
        activities = [];
        document.getElementById('summaryTableBody').innerHTML = '<tr><td colspan="10">読み込みエラー</td></tr>';
        document.getElementById('visitsTableBody').innerHTML = '<tr><td colspan="4">読み込みエラー</td></tr>';
        return;
    }

    renderDashboard();
}

// ========================================
// 目標データ取得
// ========================================
async function loadGoals() {
    const yearMonth = `${currentYear}年${currentMonth}月`;
    return await fetchAPI('getGoals', { yearMonth });
}

// ========================================
// ダッシュボード描画
// ========================================
function renderDashboard() {
    // 選択月の全活動をフィルタ
    let monthActivities = activities.filter(a => {
        const date = new Date(a.datetime);
        return date.getFullYear() === currentYear &&
            (date.getMonth() + 1) === currentMonth;
    });

    // 担当者フィルターが選択されている場合は適用
    if (selectedSalesRep) {
        monthActivities = monthActivities.filter(a => a.salesRep === selectedSalesRep);
    }

    // 会えた人のみ（実績用）
    const metActivities = monthActivities.filter(a => a.met === '○');

    // 会社別に集計
    const companyStats = {};

    // 目標がある会社を初期化
    Object.keys(currentGoals).forEach(company => {
        companyStats[company] = {
            company: company,
            target: currentGoals[company],
            count: 0,        // 会えた数（実績）
            totalAttack: 0,  // 総アタック数
            hitCount: 0,     // ○の数
            triangleCount: 0, // △の数
            missCount: 0,    // ×の数
            visits: []
        };
    });

    // 全アタックを集計
    monthActivities.forEach(a => {
        let matchedCompany = null;
        Object.keys(currentGoals).forEach(key => {
            if (a.company.includes(key) || key.includes(a.company)) {
                matchedCompany = key;
            }
        });

        if (!matchedCompany) {
            matchedCompany = a.company;
            if (!companyStats[matchedCompany]) {
                companyStats[matchedCompany] = {
                    company: matchedCompany,
                    target: 0,
                    count: 0,
                    totalAttack: 0,
                    hitCount: 0,
                    triangleCount: 0,
                    missCount: 0,
                    visits: []
                };
            }
        }

        // 総アタック数カウント
        companyStats[matchedCompany].totalAttack++;

        // 会えた結果別にカウント
        if (a.met === '○') {
            companyStats[matchedCompany].hitCount++;
            companyStats[matchedCompany].count++;
            companyStats[matchedCompany].visits.push(a);
        } else if (a.met === '△') {
            companyStats[matchedCompany].triangleCount++;
            companyStats[matchedCompany].visits.push(a);
        } else if (a.met === '×') {
            companyStats[matchedCompany].missCount++;
            companyStats[matchedCompany].visits.push(a);
        }
    });

    renderSummaryTable(companyStats);
    renderVisitsTable(companyStats);
}

// ========================================
// サマリーテーブル描画
// ========================================
function renderSummaryTable(companyStats) {
    const tbody = document.getElementById('summaryTableBody');
    const companies = Object.values(companyStats);

    if (companies.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" style="text-align:center; color: var(--text-secondary);">この月の目標データがありません</td></tr>';
        currentSummaryData = [];
        return;
    }

    let totalTarget = 0;
    let totalCount = 0;
    let totalAttack = 0;
    let totalHit = 0;
    let totalTriangle = 0;
    let totalMiss = 0;

    // CSV用データ初期化
    currentSummaryData = [['会社名', '目標', '実績', '進捗率', '残数', '総アタック', '○', '△', '×', 'ヒット率']];

    let html = companies.map(stat => {
        const progress = stat.target > 0 ? Math.round((stat.count / stat.target) * 100) : 0;
        const remaining = Math.max(0, stat.target - stat.count);
        const hitRate = stat.totalAttack > 0 ? Math.round((stat.hitCount / stat.totalAttack) * 100) : 0;

        totalTarget += stat.target;
        totalCount += stat.count;
        totalAttack += stat.totalAttack;
        totalHit += stat.hitCount;
        totalTriangle += stat.triangleCount;
        totalMiss += stat.missCount;

        // CSV用データ追加
        currentSummaryData.push([
            stat.company,
            stat.target,
            stat.count,
            `${progress}%`,
            remaining,
            "", // separator column
            stat.totalAttack,
            stat.hitCount,
            stat.triangleCount,
            stat.missCount,
            `${hitRate}%`
        ]);

        return `
      <tr>
        <td>${stat.company}</td>
        <td class="number">${stat.target}</td>
        <td class="number">${stat.count}</td>
        <td class="number">${progress}%</td>
        <td class="number">${remaining}</td>
        <td class="separator"></td>
        <td class="number">${stat.totalAttack}</td>
        <td class="number">${stat.hitCount}</td>
        <td class="number">${stat.triangleCount}</td>
        <td class="number">${stat.missCount}</td>
        <td class="number">${hitRate}%</td>
      </tr>
    `;
    }).join('');

    // 合計行
    const totalProgress = totalTarget > 0 ? Math.round((totalCount / totalTarget) * 100) : 0;
    const totalRemaining = Math.max(0, totalTarget - totalCount);
    const totalHitRate = totalAttack > 0 ? Math.round((totalHit / totalAttack) * 100) : 0;

    // CSV用合計追加
    currentSummaryData.push([
        '計',
        totalTarget,
        totalCount,
        `${totalProgress}%`,
        totalRemaining,
        "", // separator column
        totalAttack,
        totalHit,
        totalTriangle,
        totalMiss,
        `${totalHitRate}%`
    ]);

    html += `
    <tr class="total-row">
      <td>計</td>
      <td class="number">${totalTarget}</td>
      <td class="number">${totalCount}</td>
      <td class="number">${totalProgress}%</td>
      <td class="number">${totalRemaining}</td>
      <td class="separator"></td>
      <td class="number">${totalAttack}</td>
      <td class="number">${totalHit}</td>
      <td class="number">${totalTriangle}</td>
      <td class="number">${totalMiss}</td>
      <td class="number">${totalHitRate}%</td>
    </tr>
  `;

    tbody.innerHTML = html;
}

// ========================================
// 訪問者テーブル描画
// ========================================
function renderVisitsTable(companyStats) {
    const tbody = document.getElementById('visitsTableBody');
    const allVisits = [];

    Object.values(companyStats).forEach(stat => {
        stat.visits.forEach(v => {
            allVisits.push({
                company: stat.company,
                department: v.department || '',
                contact: v.contact,
                datetime: v.datetime,
                met: v.met,
                salesRep: v.salesRep || ''
            });
        });
    });

    // CSV用データ初期化
    currentVisitsData = [['会社名', '部署名', '氏名', '月日', '結果', '担当者']];

    if (allVisits.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; color: var(--text-secondary);">訪問者データなし</td></tr>';
        return;
    }

    // 結果フィルター適用
    let filteredVisits = allVisits;
    if (selectedResult) {
        filteredVisits = allVisits.filter(v => v.met === selectedResult);
    }

    // 日付でソート（新しい順）
    filteredVisits.sort((a, b) => new Date(b.datetime) - new Date(a.datetime));

    if (filteredVisits.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; color: var(--text-secondary);">該当する訪問者データがありません</td></tr>';
        return;
    }

    tbody.innerHTML = filteredVisits.map(v => {
        const date = new Date(v.datetime);
        const dateStr = `${date.getMonth() + 1}月${date.getDate()}日`;

        // 結果バッジのクラス
        let badgeClass = 'good';
        if (v.met === '△') badgeClass = 'triangle';
        else if (v.met === '×') badgeClass = 'bad';

        // CSV用データ追加（フィルター適用後のデータのみ）
        currentVisitsData.push([v.company, v.department, v.contact, dateStr, v.met, v.salesRep]);

        return `
      <tr>
        <td>${v.company}</td>
        <td>${v.department}</td>
        <td>${v.contact}</td>
        <td>${dateStr}</td>
        <td><span class="result-badge ${badgeClass}">${v.met}</span></td>
        <td>${v.salesRep}</td>
      </tr>
    `;
    }).join('');
}

// ========================================
// CSV出力（サマリー）
// ========================================
function exportSummaryCsv() {
    if (currentSummaryData.length <= 1) {
        showToast('出力するデータがありません', true);
        return;
    }

    const filename = `サマリー_${currentYear}年${currentMonth}月.csv`;
    downloadCsv(filename, currentSummaryData);
    showToast('サマリーCSVをダウンロードしました');
}

// ========================================
// CSV出力（訪問者）
// ========================================
function exportVisitsCsv() {
    if (currentVisitsData.length <= 1) {
        showToast('出力するデータがありません', true);
        return;
    }

    const filename = `訪問者一覧_${currentYear}年${currentMonth}月.csv`;
    downloadCsv(filename, currentVisitsData);
    showToast('訪問者CSVをダウンロードしました');
}

// ========================================
// CSVダウンロード共通関数
// ========================================
function downloadCsv(filename, rows) {
    const bom = '\uFEFF'; // UTF-8 BOM for Excel
    const csv = rows.map(row =>
        row.map(cell => {
            // カンマや改行を含む場合はダブルクォートで囲む
            const str = String(cell);
            if (str.includes(',') || str.includes('\n') || str.includes('"')) {
                return `"${str.replace(/"/g, '""')}"`;
            }
            return str;
        }).join(',')
    ).join('\n');

    const blob = new Blob([bom + csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
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
