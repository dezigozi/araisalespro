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

      // 日付変更時に提出済みバッジを更新し、既存データがあれば読み込む
      reportDateInput.addEventListener('change', function() {
        checkSubmittedBadge();
        loadExistingReport();
      });
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

    // 担当者変更時にバッジも更新し、既存データがあれば読み込む
    var salesRepSelect = document.getElementById('salesRepSelect');
    if (salesRepSelect) {
      salesRepSelect.addEventListener('change', function() {
        checkSubmittedBadge();
        loadExistingReport();
      });
    }

    // 初期バッジチェックと既存データ読み込み
    checkSubmittedBadge();
    loadExistingReport();

    // 日報検索ボタン
    var searchBtn = document.getElementById('reportSearchBtn');
    if (searchBtn) {
      searchBtn.addEventListener('click', handleReportSearch);
      
      // Enterキーで検索
      var searchInput = document.getElementById('reportSearchInput');
      if (searchInput) {
        searchInput.addEventListener('keydown', function(e) {
          if (e.key === 'Enter') {
            e.preventDefault();
            handleReportSearch();
          }
        });
      }
    }
  }

  // ----------------------------------------
  // 既存の日報データをフォームに読み込む
  // ----------------------------------------
  async function loadExistingReport() {
    var salesRep = document.getElementById('salesRepSelect') && document.getElementById('salesRepSelect').value;
    var reportDate = document.getElementById('reportDate') && document.getElementById('reportDate').value;
    var amField = document.getElementById('amContent');
    var pmField = document.getElementById('pmContent');
    var submitBtn = document.getElementById('submitDailyReportBtn');

    if (!amField || !pmField) return;

    if (!salesRep || !reportDate) {
      amField.value = '';
      pmField.value = '';
      return;
    }

    if (submitBtn) {
      submitBtn.disabled = true;
      submitBtn.textContent = 'データ取得中...';
    }

    amField.value = '';
    pmField.value = '';

    try {
      var res = await fetchAPI('getDailyReports', { salesRep: salesRep, limit: 100 });
      if (res && res.success && res.data) {
        var found = res.data.find(function(r) { return r.date === reportDate; });
        if (found) {
          amField.value = found.amContent || '';
          pmField.value = found.pmContent || '';
        }
      }
    } catch(err) {
      console.error(err);
    } finally {
      if (submitBtn) {
        submitBtn.disabled = false;
        submitBtn.textContent = '📝 日報を提出する';
      }
    }
  }

  // ----------------------------------------
  // POST が CORS で応答を読めない場合、GET で保存済みか照合する
  // ----------------------------------------
  function normDailyText(s) {
    return String(s || '').trim();
  }

  async function verifyDailyReportSaved(salesRep, reportDate, amContent, pmContent) {
    var waitMs = [0, 450, 900, 1600];
    var i;
    for (i = 0; i < waitMs.length; i++) {
      if (waitMs[i] > 0) {
        await new Promise(function (resolve) { setTimeout(resolve, waitMs[i]); });
      }
      var res = await fetchAPI('getDailyReports', { salesRep: salesRep, limit: 60 });
      if (!res || !res.success || !res.data) continue;
      var j;
      for (j = 0; j < res.data.length; j++) {
        var r = res.data[j];
        if (r.date !== reportDate || r.salesRep !== salesRep) continue;
        if (normDailyText(r.amContent) !== normDailyText(amContent)) continue;
        if (normDailyText(r.pmContent) !== normDailyText(pmContent)) continue;
        return true;
      }
    }
    return false;
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
      var result = await postAPI({
        action: 'saveDailyReport',
        salesRep: salesRep,
        date: reportDate,
        amContent: amContent,
        pmContent: pmContent
      });

      if (!result || !result.success) {
        showToast('送信失敗: ' + ((result && result.error) || '不明なエラー'), true);
        return;
      }

      var confirmed = !result.unverified;
      if (result.unverified) {
        if (btn) btn.textContent = '保存確認中...';
        confirmed = await verifyDailyReportSaved(salesRep, reportDate, amContent, pmContent);
      }

      if (confirmed) {
        localStorage.setItem('drep_' + salesRep + '_' + reportDate, '1');
        showToast('✅ 日報を提出しました！');
        // 保存した内容をフォームに残し、続けて編集が行えるようにする
        checkSubmittedBadge();
        var pastList = document.getElementById('pastReportsList');
        if (pastList && pastList.style.display !== 'none') {
          pastReportsLoaded = false;
          await loadPastReports();
          pastReportsLoaded = true;
        }
      } else {
        showToast(
          '送信リクエストは出しましたが、保存を確認できませんでした。' +
            '「過去の日報」またはスプレッドシートで確認し、未反映なら再度お試しください。',
          true
        );
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

  // ----------------------------------------
  // 日報検索
  // ----------------------------------------
  async function handleReportSearch() {
    var query = document.getElementById('reportSearchInput').value.trim();
    var resultList = document.getElementById('reportSearchContent');
    var loading = document.getElementById('reportSearchLoading');
    var container = document.getElementById('reportSearchResult');

    if (!query) {
      showToast('検索キーワードを入力してください', true);
      return;
    }

    container.style.display = 'flex';
    container.style.flexDirection = 'column';
    container.style.gap = '0.75rem';
    container.style.marginTop = '1rem';
    loading.style.display = 'block';
    resultList.innerHTML = '';
    
    var btn = document.getElementById('reportSearchBtn');
    if (btn) btn.disabled = true;

    try {
      var result = await fetchAPI('searchDailyReports', { query: query });
      
      loading.style.display = 'none';

      if (!result || !result.success) {
        resultList.innerHTML = '<div style="color:var(--accent-red);text-align:center;padding:1rem;">検索に失敗しました</div>';
        return;
      }

      if (!result.data || result.data.length === 0) {
        resultList.innerHTML = '<div style="color:var(--text-secondary);text-align:center;padding:1rem;">一致する日報が見つかりません</div>';
        return;
      }

      var dayNames = ['日', '月', '火', '水', '木', '金', '土'];

      resultList.innerHTML = result.data.map(function (report) {
        var dateObj = new Date(report.date + 'T00:00:00');
        var dateStr = dateObj.getFullYear() + '年' +
                      (dateObj.getMonth() + 1) + '月' +
                      dateObj.getDate() + '日(' +
                      dayNames[dateObj.getDay()] + ')';

        var amHtml = report.amContent
          ? '<div style="margin-top:0.75rem;font-size:0.85rem;"><strong style="color:var(--accent-red);">☀️ AM</strong><div style="color:var(--text-secondary);white-space:pre-wrap;margin-top:0.25rem;line-height:1.5;">' + escapeHtml(report.amContent) + '</div></div>'
          : '';

        var pmHtml = report.pmContent
          ? '<div style="margin-top:0.75rem;font-size:0.85rem;"><strong style="color:var(--accent-blue);">🌙 PM</strong><div style="color:var(--text-secondary);white-space:pre-wrap;margin-top:0.25rem;line-height:1.5;">' + escapeHtml(report.pmContent) + '</div></div>'
          : '';

        var repBadge = report.salesRep
          ? '<div class="activity-sales-rep" style="margin-left: auto;">' + escapeHtml(report.salesRep) + '</div>'
          : '';

        return '<div class="activity-item">' +
                 '<div class="activity-header">' +
                   '<span class="activity-date">📅 ' + dateStr + '</span>' +
                   repBadge +
                 '</div>' +
                 amHtml + pmHtml +
               '</div>';
      }).join('');
    } catch (e) {
      loading.style.display = 'none';
      resultList.innerHTML = '<div style="color:var(--accent-red);text-align:center;padding:1rem;">エラーが発生しました</div>';
    } finally {
      if (btn) btn.disabled = false;
    }
  }

})();
