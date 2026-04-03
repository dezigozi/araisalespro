// ========================================
// 訪問予定カレンダー
// ========================================
const GAS_URL_CALENDAR = GAS_WEBAPP_URL;

// 状態管理
let calendarState = {
  year: new Date().getFullYear(),
  month: new Date().getMonth() + 1,
  schedules: [],
  selectedCell: null,  // 単一選択に変更
  salesReps: ['高野', '青木', '土岐', '中村'],
  dragData: null  // ドラッグ中のデータ
};

// ========================================
// 初期化
// ========================================
function initVisitCalendar() {
  const calendarSection = document.getElementById('visitCalendarSection');
  if (!calendarSection) return;

  setupCalendarEvents();
  loadVisitSchedule();
}

// ========================================
// イベント設定
// ========================================
function setupCalendarEvents() {
  // 月ナビゲーション
  document.getElementById('prevMonthBtn')?.addEventListener('click', () => {
    calendarState.month--;
    if (calendarState.month < 1) {
      calendarState.month = 12;
      calendarState.year--;
    }
    loadVisitSchedule();
  });

  document.getElementById('nextMonthBtn')?.addEventListener('click', () => {
    calendarState.month++;
    if (calendarState.month > 12) {
      calendarState.month = 1;
      calendarState.year++;
    }
    loadVisitSchedule();
  });

  // 保存ボタン
  document.getElementById('saveScheduleBtn')?.addEventListener('click', saveSchedule);

  // 削除ボタン
  document.getElementById('deleteScheduleBtn')?.addEventListener('click', deleteSchedule);

  // 時間帯ラジオボタン変更時に即座に適用
  document.querySelectorAll('input[name="timeSlot"]').forEach(radio => {
    radio.addEventListener('change', applyTimeSlotToSelected);
  });
}

// ========================================
// 訪問予定取得
// ========================================
async function loadVisitSchedule() {
  updateCalendarMonthDisplay();

  const matrix = document.getElementById('calendarMatrix');
  if (!matrix) return;

  matrix.innerHTML = '<div class="loading">読み込み中...</div>';

  const result = await fetchCalendarAPI('getVisitSchedule', {
    year: calendarState.year,
    month: calendarState.month
  });

  if (result?.success) {
    calendarState.schedules = result.data || [];
    renderCalendarMatrix();
  } else {
    matrix.innerHTML = '<div class="empty-state">データの読み込みに失敗しました</div>';
  }
}

// ========================================
// 月表示更新
// ========================================
function updateCalendarMonthDisplay() {
  // dashboard.htmlでは calendarMonth, index.htmlでは currentMonth
  const display = document.getElementById('calendarMonth') || document.getElementById('currentMonth');
  if (display) {
    display.textContent = `${calendarState.year}年${calendarState.month}月`;
  }
}

// ========================================
// カレンダーマトリクス描画
// ========================================
function renderCalendarMatrix() {
  const matrix = document.getElementById('calendarMatrix');
  if (!matrix) return;

  // 選択状態リセット
  calendarState.selectedCell = null;
  updateSaveButtonState();

  // 月の日数を計算
  const daysInMonth = new Date(calendarState.year, calendarState.month, 0).getDate();

  // 曜日名
  const dayNames = ['日', '月', '火', '水', '木', '金', '土'];

  // テーブル生成
  let html = '<table class="calendar-table">';

  // ヘッダー行（曜日 + 日付）
  html += '<thead><tr><th class="rep-header">担当</th>';
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(calendarState.year, calendarState.month - 1, d);
    const dayOfWeek = date.getDay();
    const dayClass = dayOfWeek === 0 ? 'sunday' : dayOfWeek === 6 ? 'saturday' : '';
    html += `<th class="date-header ${dayClass}">
      <span class="day-name">${dayNames[dayOfWeek]}</span>
      <span class="day-num">${d}</span>
    </th>`;
  }
  html += '</tr></thead>';

  // 担当者ごとの行
  html += '<tbody>';
  calendarState.salesReps.forEach(rep => {
    html += `<tr><td class="rep-name">${rep}</td>`;

    for (let d = 1; d <= daysInMonth; d++) {
      const dateStr = `${calendarState.year}-${String(calendarState.month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const schedule = calendarState.schedules.find(
        s => s.担当 === rep && s.訪問予定日 === dateStr
      );

      const timeSlot = schedule?.訪問時間 || '';
      const destination = schedule?.訪問先 || '';
      const timeClass = getTimeSlotClass(timeSlot);

      const date = new Date(calendarState.year, calendarState.month - 1, d);
      const dayOfWeek = date.getDay();
      const weekendClass = dayOfWeek === 0 ? 'sunday-cell' : dayOfWeek === 6 ? 'saturday-cell' : '';
      
      // 予定があればドラッグ可能
      const hasSchedule = !!timeSlot;
      const draggableAttr = hasSchedule ? 'draggable="true"' : '';
      const draggableClass = hasSchedule ? 'draggable' : '';

      html += `<td class="schedule-cell ${timeClass} ${weekendClass} ${draggableClass}" 
                   data-rep="${rep}" 
                   data-date="${dateStr}"
                   data-timeslot="${timeSlot}"
                   data-destination="${destination}"
                   ${draggableAttr}>
        <div class="cell-content">
          <input type="text" 
                 class="destination-input" 
                 value="${destination}" 
                 maxlength="10" 
                 placeholder=""
                 data-rep="${rep}"
                 data-date="${dateStr}">
          ${timeSlot ? `<span class="time-badge ${timeClass}">${timeSlot}</span>` : ''}
        </div>
      </td>`;
    }

    html += '</tr>';
  });
  html += '</tbody></table>';

  matrix.innerHTML = html;

  // セルクリックイベント（単一選択）
  matrix.querySelectorAll('.schedule-cell').forEach(cell => {
    cell.addEventListener('click', (e) => {
      if (e.target.classList.contains('destination-input')) return;
      handleCellClick(cell);
    });
    
    // ドラッグ&ドロップイベント
    cell.addEventListener('dragstart', handleDragStart);
    cell.addEventListener('dragover', handleDragOver);
    cell.addEventListener('dragenter', handleDragEnter);
    cell.addEventListener('dragleave', handleDragLeave);
    cell.addEventListener('drop', handleDrop);
    cell.addEventListener('dragend', handleDragEnd);
  });

  // 訪問先入力イベント
  matrix.querySelectorAll('.destination-input').forEach(input => {
    input.addEventListener('change', (e) => {
      handleDestinationChange(e.target);
    });
    input.addEventListener('focus', (e) => {
      const cell = e.target.closest('.schedule-cell');
      if (cell && !cell.classList.contains('selected')) {
        handleCellClick(cell);
      }
    });
    // 入力欄のドラッグを無効化
    input.addEventListener('dragstart', (e) => e.preventDefault());
  });
}

// ========================================
// 時間帯のクラス取得
// ========================================
function getTimeSlotClass(timeSlot) {
  switch (timeSlot) {
    case '終日': return 'slot-allday';
    case 'AM': return 'slot-am';
    case 'PM': return 'slot-pm';
    default: return '';
  }
}

// ========================================
// セルクリック処理（単一選択）
// ========================================
function handleCellClick(cell) {
  // 以前の選択を解除
  if (calendarState.selectedCell && calendarState.selectedCell !== cell) {
    calendarState.selectedCell.classList.remove('selected');
  }

  // 新しいセルを選択
  const wasSelected = cell.classList.contains('selected');

  if (wasSelected) {
    cell.classList.remove('selected');
    calendarState.selectedCell = null;
  } else {
    cell.classList.add('selected');
    calendarState.selectedCell = cell;

    // 現在の時間帯をラジオボタンに反映
    const currentSlot = cell.dataset.timeslot;
    if (currentSlot) {
      const radio = document.querySelector(`input[name="timeSlot"][value="${currentSlot}"]`);
      if (radio) radio.checked = true;
    }
  }

  updateSaveButtonState();
}

// ========================================
// 時間帯選択を選択中のセルに適用
// ========================================
function applyTimeSlotToSelected() {
  if (!calendarState.selectedCell) return;

  const selectedSlot = document.querySelector('input[name="timeSlot"]:checked')?.value;
  if (!selectedSlot) return;

  const cell = calendarState.selectedCell;

  // クラスを更新
  cell.classList.remove('slot-allday', 'slot-am', 'slot-pm');
  cell.classList.add(getTimeSlotClass(selectedSlot));
  cell.dataset.timeslot = selectedSlot;

  // バッジを更新
  let badge = cell.querySelector('.time-badge');
  if (!badge) {
    badge = document.createElement('span');
    badge.className = 'time-badge';
    cell.querySelector('.cell-content').appendChild(badge);
  }
  badge.textContent = selectedSlot;
  badge.className = `time-badge ${getTimeSlotClass(selectedSlot)}`;

  updateSaveButtonState();
}

// ========================================
// 訪問先変更処理
// ========================================
function handleDestinationChange(input) {
  updateSaveButtonState();
}

// ========================================
// 保存ボタン状態更新
// ========================================
function updateSaveButtonState() {
  const saveBtn = document.getElementById('saveScheduleBtn');
  const deleteBtn = document.getElementById('deleteScheduleBtn');
  const hasSelection = !!calendarState.selectedCell;

  if (saveBtn) {
    // 選択中のセルがあれば保存可能
    saveBtn.disabled = !hasSelection;
  }

  if (deleteBtn) {
    // 選択中のセルに予定が登録されていれば削除可能
    const hasSchedule = hasSelection && calendarState.selectedCell.dataset.timeslot;
    deleteBtn.disabled = !hasSchedule;
  }
}

// ========================================
// 選択中のセルを保存
// ========================================
async function saveSchedule() {
  if (!calendarState.selectedCell) return;

  const btn = document.getElementById('saveScheduleBtn');
  if (btn) {
    btn.disabled = true;
    btn.textContent = '保存中...';
  }

  const cell = calendarState.selectedCell;
  const rep = cell.dataset.rep;
  const date = cell.dataset.date;
  const timeSlot = cell.dataset.timeslot || '';
  const input = cell.querySelector('.destination-input');
  const destination = input?.value || '';

  try {
    await postCalendarAPI({
      action: 'updateVisitSchedule',
      担当: rep,
      訪問予定日: date,
      訪問時間: timeSlot,
      訪問先: destination
    });

    // 選択解除
    cell.classList.remove('selected');
    calendarState.selectedCell = null;

    showCalendarToast('保存しました');

    // リロード
    await loadVisitSchedule();
  } catch (e) {
    console.error('保存エラー:', e);
    showCalendarToast('保存に失敗しました', true);
  }

  if (btn) {
    btn.disabled = true;
    btn.textContent = '💾 保存';
  }
}

// ========================================
// 選択中のセルを削除
// ========================================
async function deleteSchedule() {
  if (!calendarState.selectedCell) return;

  const cell = calendarState.selectedCell;
  const timeSlot = cell.dataset.timeslot;

  // 予定が登録されていない場合は何もしない
  if (!timeSlot) return;

  // 確認ダイアログ
  if (!confirm('この予定を削除しますか？')) return;

  const btn = document.getElementById('deleteScheduleBtn');
  if (btn) {
    btn.disabled = true;
    btn.textContent = '削除中...';
  }

  const rep = cell.dataset.rep;
  const date = cell.dataset.date;

  try {
    // GETリクエストで削除（POSTよりCORS対応が良い）
    const result = await fetchCalendarAPI('deleteVisitSchedule', {
      rep: rep,
      date: date
    });

    if (result?.success) {
      // 選択解除
      cell.classList.remove('selected');
      calendarState.selectedCell = null;

      showCalendarToast('削除しました');

      // リロード
      await loadVisitSchedule();
    } else {
      showCalendarToast(result?.error || '削除に失敗しました', true);
    }
  } catch (e) {
    console.error('削除エラー:', e);
    showCalendarToast('削除に失敗しました', true);
  }

  if (btn) {
    btn.disabled = true;
    btn.textContent = '🗑️ 削除';
  }
}

// ========================================
// API通信
// ========================================
async function fetchCalendarAPI(action, params = {}) {
  const url = new URL(GAS_URL_CALENDAR);
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

async function postCalendarAPI(data) {
  // GAS Web App has CORS limitations for POST requests
  // Use no-cors mode and verify by reloading data
  await fetch(GAS_URL_CALENDAR, {
    method: 'POST',
    mode: 'no-cors',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify(data)
  });
  // Cannot read response with no-cors, return success and verify via reload
  return { success: true };
}

// ========================================
// ドラッグ&ドロップ処理
// ========================================
function handleDragStart(e) {
  const cell = e.target.closest('.schedule-cell');
  if (!cell || !cell.dataset.timeslot) {
    e.preventDefault();
    return;
  }
  
  // ドラッグデータを保存
  calendarState.dragData = {
    rep: cell.dataset.rep,
    date: cell.dataset.date,
    timeslot: cell.dataset.timeslot,
    destination: cell.dataset.destination || cell.querySelector('.destination-input')?.value || ''
  };
  
  cell.classList.add('dragging');
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/plain', JSON.stringify(calendarState.dragData));
}

function handleDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
}

function handleDragEnter(e) {
  e.preventDefault();
  const cell = e.target.closest('.schedule-cell');
  if (cell && calendarState.dragData) {
    // 同じ担当者の行のみドロップ可能
    if (cell.dataset.rep === calendarState.dragData.rep) {
      cell.classList.add('drag-over');
    } else {
      cell.classList.add('drag-invalid');
    }
  }
}

function handleDragLeave(e) {
  const cell = e.target.closest('.schedule-cell');
  if (cell) {
    cell.classList.remove('drag-over', 'drag-invalid');
  }
}

function handleDrop(e) {
  e.preventDefault();
  const targetCell = e.target.closest('.schedule-cell');
  if (!targetCell || !calendarState.dragData) return;
  
  targetCell.classList.remove('drag-over', 'drag-invalid');
  
  const sourceData = calendarState.dragData;
  const targetRep = targetCell.dataset.rep;
  const targetDate = targetCell.dataset.date;
  
  // 同じ担当者の行のみ移動可能
  if (targetRep !== sourceData.rep) {
    showCalendarToast('同じ担当者の行にのみ移動できます', true);
    return;
  }
  
  // 同じ日付なら何もしない
  if (targetDate === sourceData.date) {
    return;
  }
  
  // 移動先に既に予定がある場合は確認
  if (targetCell.dataset.timeslot) {
    if (!confirm('移動先に既存の予定があります。上書きしますか？')) {
      return;
    }
  }
  
  // 予定を移動
  moveSchedule(sourceData, targetDate);
}

function handleDragEnd(e) {
  // すべてのドラッグ関連クラスをリセット
  document.querySelectorAll('.schedule-cell').forEach(cell => {
    cell.classList.remove('dragging', 'drag-over', 'drag-invalid');
  });
  calendarState.dragData = null;
}

// 予定を移動する
async function moveSchedule(sourceData, targetDate) {
  showCalendarToast('移動中...', false);
  
  try {
    // 1. 移動先に新しい予定を作成
    await postCalendarAPI({
      action: 'updateVisitSchedule',
      担当: sourceData.rep,
      訪問予定日: targetDate,
      訪問時間: sourceData.timeslot,
      訪問先: sourceData.destination
    });
    
    // 2. 元の予定を削除
    await fetchCalendarAPI('deleteVisitSchedule', {
      rep: sourceData.rep,
      date: sourceData.date
    });
    
    showCalendarToast('移動しました');
    
    // リロード
    await loadVisitSchedule();
  } catch (e) {
    console.error('移動エラー:', e);
    showCalendarToast('移動に失敗しました', true);
    await loadVisitSchedule();
  }
}

// ========================================
// Toast通知
// ========================================
function showCalendarToast(msg, isError = false) {
  const toast = document.getElementById('toast');
  if (!toast) return;

  toast.textContent = msg;
  toast.className = 'toast' + (isError ? ' error' : '');
  toast.classList.add('show');

  setTimeout(() => {
    toast.classList.remove('show');
  }, 3000);
}

// ========================================
// 初期化実行
// ========================================
document.addEventListener('DOMContentLoaded', () => {
  initVisitCalendar();
});
