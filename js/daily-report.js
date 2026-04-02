// ========================================
// 営業日報 機能
// ========================================

(function () {
  'use strict';

  // ページ読み込み時に初期化
  document.addEventListener('DOMContentLoaded', function () {
    initDailyReport();
  });

  function initDailyReport() {
    // 今日の日付をセット
    var reportDateInput = document.getElementById('reportDate');
    if (reportDateInput) {
      var now = new Date();
      var y = now.getFullYear();
      var m = String(now.getMonth() + 1).padStart(2, '0');
      var d = String(now.getDate()).padStart(2, '0');
      reportDateInput.value = y + '-' + m + '-' + d;

      // 日付変更時に提出済みバッジを更新
      reportDateInput.addEventListener('change', checkSubmittedBadge);
    }

    // フォーム送信
    var form = document.getElementById('dailyReportForm');
    if (form) {
      form.addEventListener('submit', handleDailyReportSubmit);
    }

    // 過去の日報トグル
    var toggleBtn = document.getElementById('togglePastReports');
    if (toggleBtn) {
      toggleBtn.addEventListener('click', togglePastReportsList);
    }

    // 担当者変更時にバッジも更新
    var salesRepSelect = document.getElementById('salesRepSelect');
    if (salesRepSelect) {
      salesRepSelect.addEventListener('change', checkSubmittedBadge);
    }

    // 初期バッジチェック
    checkSubmittedBadge();
  }

  // ----------------------------------------
  // 日報フォーム送信
  // ----------------------------------------
  async function handleDailyReportSubmit(e) {
    e.preventDefault();

    var salesRep = document.getElementById('salesRepSelect') &&
                   document.getElementById('salesRepSelect').value;
    if (!salesRep) {
      showToast('営業担当者を選択してください', true);
      return;
    }

    var reportDate = document.getElementById('reportDate').value;
    var amContent = document.getElementById('amContent').value.trim();
    var pmContent = document.getElementById('pmContent').value.trim();

    if (!amContent && !pmContent) {
      showToast('AM・PMどちらかの業務内容を入力してください', true);
      return;
    }

    var btn = document.getElementById('submitDailyReportBtn');
    if (btn) {
      btn.disabled = true;
      btn.textContent = '送信中...';
    }

    try {
      await postAPI({
        action: 'saveDailyReport',
        salesRep: salesRep,
        date: reportDate,
        amContent: amContent,
        pmContent: pmContent
      });

      // 提出済みフラグをlocalStorageに保存
      localStorage.setItem('drep_' + salesRep + '_' + reportDate, '1');

      showToast('✅ 日報を提出しました！');

      // テキストエリアをクリア
      document.getElementById('amContent').value = '';
      document.getElementById('pmContent').value = '';

      // バッジ更新
      checkSubmittedBadge();

      // 過去の日報が表示中なら再読み込み
      var pastList = document.getElementById('pastReportsList');
      if (pastList && pastList.style.display !== 'none') {
        pastReportsLoaded = false;
        await loadPastReports();
        pastReportsLoaded = true;
      }
    } catch (err) {
      showToast('送信失敗: ' + (err.message || '不明なエラー'), true);
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.textContent = '📝 日報を提出する';
      }
    }
  }

  // ----------------------------------------
  // 提出済みバッジ表示/非表示
  // ----------------------------------------
  function checkSubmittedBadge() {
    var salesRep = document.getElementById('salesRepSelect') &&
                   document.getElementById('salesRepSelect').value;
    var reportDate = document.getElementById('reportDate') &&
                     document.getElementById('reportDate').value;
    var badge = document.getElementById('reportSubmittedBadge');
    if (!badge) return;

    if (salesRep && reportDate) {
      var submitted = localStorage.getItem('drep_' + salesRep + '_' + reportDate);
      badge.style.display = submitted ? 'inline-flex' : 'none';
    } else {
      badge.style.display = 'none';
    }
  }

  // ----------------------------------------
  // 過去の日報 開閉トグル
  // ----------------------------------------
  var pastReportsLoaded = false;

  async function togglePastReportsList() {
    var list = document.getElementById('pastReportsList');
    var btn = document.getElementById('togglePastReports');
    if (!list || !btn) return;

    var isHidden = list.style.display === 'none';

    if (isHidden) {
      list.style.display = 'flex';
      btn.textContent = '📅 過去の日報を閉じる ▲';
      if (!pastReportsLoaded) {
        await loadPastReports();
        pastReportsLoaded = true;
      }
    } else {
      list.style.display = 'none';
      btn.textContent = '📅 過去の日報を見る ▼';
      pastReportsLoaded = false; // 次回開いたとき再取得
    }
  }

  // ----------------------------------------
  // 過去の日報を読み込んで描画
  // ----------------------------------------
  async function loadPastReports() {
    var list = document.getElementById('pastReportsList');
    if (!list) return;

    list.innerHTML = '<div style="color:var(--text-secondary);text-align:center;padding:1rem;">読み込み中...</div>';

    var salesRep = (document.getElementById('salesRepSelect') &&
                    document.getElementById('salesRepSelect').value) || '';

    var result = null;
    try {
      result = await fetchAPI('getDailyReports', { salesRep: salesRep, limit: 30 });
    } catch (e) {
      list.innerHTML = '<div style="color:var(--accent-red);text-align:center;padding:1rem;">読み込みに失敗しました</div>';
      return;
    }

    if (!result || !result.success) {
      list.innerHTML = '<div style="color:var(--accent-red);text-align:center;padding:1rem;">読み込みに失敗しました</div>';
      return;
    }

    if (!result.data || result.data.length === 0) {
      list.innerHTML = '<div style="color:var(--text-secondary);text-align:center;padding:1rem;">過去の日報がありません</div>';
      return;
    }

    var dayNames = ['日', '月', '火', '水', '木', '金', '土'];

    list.innerHTML = result.data.map(function (report) {
      var dateObj = new Date(report.date + 'T00:00:00');
      var dateStr = dateObj.getFullYear() + '年' +
                    (dateObj.getMonth() + 1) + '月' +
                    dateObj.getDate() + '日(' +
                    dayNames[dateObj.getDay()] + ')';

      var amHtml = report.amContent
        ? '<div class="past-report-section">' +
          '<div class="past-report-section-label am">☀️ AM 業務内容</div>' +
          '<div class="past-report-content">' + escapeHtml(report.amContent) + '</div>' +
          '</div>'
        : '';

      var pmHtml = report.pmContent
        ? '<div class="past-report-section">' +
          '<div class="past-report-section-label pm">🌙 PM 業務内容</div>' +
          '<div class="past-report-content">' + escapeHtml(report.pmContent) + '</div>' +
          '</div>'
        : '';

      var repBadge = report.salesRep
        ? '<span class="past-report-rep">' + escapeHtml(report.salesRep) + '</span>'
        : '';

      return '<div class="past-report-item">' +
               '<div class="past-report-date-row">' +
                 '<span class="past-report-date">📅 ' + dateStr + '</span>' +
                 repBadge +
               '</div>' +
               amHtml +
               pmHtml +
             '</div>';
    }).join('');
  }

  // ----------------------------------------
  // HTMLエスケープ
  // ----------------------------------------
  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

})();
