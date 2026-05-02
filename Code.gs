// ========================================
// 運用方針（Apps Script の置き場所）
// ========================================
// ・このプロジェクトは「ARAI_SALES_PRO」スプレッドシートにのみコンテナバインドすること。
//   （拡張機能 → Apps Script は ARAI のブックから開く）
// ・ウェブアプリのデプロイはこの 1 か所だけ。フロントは js/gas-config.js の GAS_WEBAPP_URL をこの /exec に合わせる。
// ・「特販部_営業訪問ログ」には GAS を置かない（過去に複製している場合は削除してよい）。
//   活動記録・営業日報は下記 ACTIVITY_LOG_SPREADSHEET_ID で openById して書き込む。
// ========================================
// 設定
// ========================================
const SPREADSHEET_ID = '1FcBvdskysuy59z8JhNSgrkN0WWdl7hM5sBtZpBn03ws'; // ARAI_SALES_PRO メインスプレッドシート


// 活動記録・営業日報用スプレッドシートID（同一ブック「特販部_営業訪問ログ」内のシート「活動記録」「営業日報」）
const ACTIVITY_LOG_SPREADSHEET_ID = '1l2K-ODGJGmE1zqYlUVDck_cVmoefHkoebxHGI53Ekzw';

// 営業日報も上記と同じブックのシート「営業日報」へ保存（別ブックに分ける場合のみここを別IDに変更）
const DAILY_REPORT_SPREADSHEET_ID = ACTIVITY_LOG_SPREADSHEET_ID;

function getDailyReportSpreadsheet() {
    return SpreadsheetApp.openById(DAILY_REPORT_SPREADSHEET_ID);
}

// ========================================
// 【初回のみ手動実行】営業日報シート初期設定
// GASエディタで setupDailyReportSheet を選択して▶実行
// ========================================
function setupDailyReportSheet() {
    var ss = getDailyReportSpreadsheet();
    var sheetName = '営業日報';
    var sheet = ss.getSheetByName(sheetName);

    // なければ作成
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        Logger.log('✅ シート「営業日報」を新規作成しました');
    } else {
        Logger.log('ℹ️ シート「営業日報」は既に存在します。書式を更新します');
    }

    // ヘッダー行が空なら設定
    var firstCell = sheet.getRange('A1').getValue();
    if (!firstCell) {
        sheet.getRange('A1:D1').setValues([['年月日', '営業担当者', '作業時間', '項目']]);
        Logger.log('✅ ヘッダー行を設定しました');
    }

    // ---- 書式設定 ----

    // ヘッダー行: 背景色・文字色・太字
    var headerRange = sheet.getRange('A1:D1');
    headerRange.setBackground('#1a237e');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    headerRange.setHorizontalAlignment('center');

    // 列幅
    sheet.setColumnWidth(1, 120);  // A: 年月日
    sheet.setColumnWidth(2, 110);  // B: 営業担当者
    sheet.setColumnWidth(3, 80);   // C: 作業時間
    sheet.setColumnWidth(4, 600);  // D: 項目

    // 1行目を固定
    sheet.setFrozenRows(1);

    // C列（作業時間）に入力規則 AM/PM
    var timeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['AM', 'PM'], true)
        .setAllowInvalid(false)
        .build();
    sheet.getRange('C2:C1000').setDataValidation(timeRule);

    // A列（日付）の書式
    sheet.getRange('A2:A1000').setNumberFormat('yyyy/MM/dd');

    // D列（項目）折り返し
    sheet.getRange('D2:D1000').setWrap(true);

    // グリッド線を見やすく（縞模様は不要なので省略）
    Logger.log('✅ 営業日報シートの初期設定が完了しました。スプレッドシートを確認してください。');
    Logger.log('📋 URL: https://docs.google.com/spreadsheets/d/' + DAILY_REPORT_SPREADSHEET_ID);
}


// 活動記録用スプレッドシートを取得するヘルパー関数
function getActivityLogSpreadsheet() {
    return SpreadsheetApp.openById(ACTIVITY_LOG_SPREADSHEET_ID);
}

// ========================================
// Webアプリのエントリーポイント（GET）
// ========================================
function doGet(e) {
    const action = e.parameter.action;
    let result;

    switch (action) {
        case 'getCustomers':
            result = getCustomers();
            break;
        case 'getDepartments':
            result = getDepartments(e.parameter.company);
            break;
        case 'getContactsByDept':
            result = getContactsByDept(e.parameter.company, e.parameter.department);
            break;
        case 'getAllMasterData':
            result = getAllMasterData();
            break;
        case 'getActivities':
            result = getActivities();
            break;
        case 'getMaterials':
            result = getMaterials();
            break;
        case 'getOurReps':
            result = getOurReps();
            break;
        case 'getGoals':
            result = getGoals(e.parameter.yearMonth);
            break;
        case 'getProposalProducts':
            result = getProposalProducts(e.parameter.yearMonth);
            break;
        case 'getVisitSchedule':
            result = getVisitSchedule(e.parameter.year, e.parameter.month);
            break;
        case 'getPerformanceData':
            result = getPerformanceData();
            break;
        case 'getPerformanceRawData':
            result = getPerformanceRawData();
            break;
        case 'getSalesAnalysisData':
            result = getSalesAnalysisData();
            break;
        case 'getOrderAnalysisData':
            result = getOrderAnalysisData(e.parameter.offset, e.parameter.limit);
            break;
        case 'getSheetLastModified':
            result = getSheetLastModified();
            break;
        case 'getCustomerPhones':
            result = getCustomerPhones();
            break;
        case 'getContactsWithLastActivity':
            result = getContactsWithLastActivity();
            break;
        case 'getActionList':
            result = getActionList(e.parameter.yearMonth);
            break;
        case 'deleteVisitSchedule':
            result = deleteVisitSchedule({
                担当: e.parameter.rep,
                訪問予定日: e.parameter.date
            });
            break;
        case 'getDailyReports':
            result = getDailyReports(e.parameter.salesRep, e.parameter.limit);
            break;
        case 'searchDailyReports':
            result = searchDailyReports(e.parameter.query);
            break;
        default:
            result = { success: false, error: 'Unknown action' };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
}

// ========================================
// Webアプリのエントリーポイント（POST）
// ========================================
function doPost(e) {
    const data = JSON.parse(e.postData.contents);
    let result;

    switch (data.action) {
        case 'addActivity':
            result = addActivity(data);
            break;
        case 'sendPromotionEmail':
            result = sendPromotionEmail(data);
            break;
        case 'updateVisitSchedule':
            result = updateVisitSchedule(data);
            break;
        case 'addContact':
            result = addContact(data);
            break;
        case 'updateActivity':
            result = updateActivity(data);
            break;
        case 'deleteActivity':
            result = deleteActivity(data);
            break;
        case 'deleteVisitSchedule':
            result = deleteVisitSchedule(data);
            break;
        case 'saveActionList':
            result = saveActionList(data);
            break;
        case 'removeFromActionList':
            result = removeFromActionList(data);
            break;
        case 'saveDailyReport':
            result = saveDailyReport(data);
            break;
        default:
            result = { success: false, error: 'Unknown action' };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
}

// ========================================
// 顧客一覧取得（修正済み：B列=会社名）
// ========================================
function getCustomers() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('担当者マスタ');
    const data = sheet.getDataRange().getValues();

    // A列: 会社名
    const customers = [...new Set(data.slice(1).map(row => row[0]).filter(c => c))];
    return { success: true, data: customers };
}

// ========================================
// 部署一覧取得（修正済み：B列=会社名, C列=部署）
// ========================================
function getDepartments(company) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('担当者マスタ');
    const data = sheet.getDataRange().getValues();

    const departments = [...new Set(
        data.slice(1)
            .filter(row => row[0] === company) // A列: 会社名
            .map(row => row[1]) // B列: 部署
            .filter(d => d)
    )];

    return { success: true, data: departments };
}

// ========================================
// 担当者一覧取得（修正済み：D列=担当者名）
// ========================================
function getContactsByDept(company, department) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('担当者マスタ');
    const data = sheet.getDataRange().getValues();

    const contacts = data.slice(1)
        .filter(row => row[0] === company && row[1] === department) // A列: 会社名, B列: 部署
        .map(row => row[3]) // D列: 担当者名
        .filter(c => c);

    return { success: true, data: contacts };
}

// ========================================
// 全マスタデータ一括取得（修正済み：F列=メール）
// ========================================
function getAllMasterData() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('担当者マスタ');
    const data = sheet.getDataRange().getValues();

    const customers = [];
    const departments = {};
    const contacts = {};

    data.slice(1).forEach(row => {
        const company = row[0];  // A列: 会社名
        const dept = row[1];     // B列: 部署
        const contact = row[3];  // D列: 担当者名

        if (!company) return;

        if (!customers.includes(company)) {
            customers.push(company);
        }

        if (!departments[company]) {
            departments[company] = [];
        }
        if (dept && !departments[company].includes(dept)) {
            departments[company].push(dept);
        }

        const key = `${company}_${dept}`;
        if (!contacts[key]) {
            contacts[key] = [];
        }
        if (contact && !contacts[key].includes(contact)) {
            contacts[key].push(contact);
        }
    });

    return {
        success: true,
        // contactEmailsは空オブジェクトを返してフロント側のエラーを防ぐ
        data: { customers, departments, contacts, contactEmails: {} }
    };
}

// ========================================
// 活動履歴取得（別スプレッドシート対応版）
// ========================================
function getActivities() {
    const ss = getActivityLogSpreadsheet();
    const sheet = ss.getSheetByName('活動記録');

    if (!sheet) {
        return { success: true, data: [] };
    }

    const data = sheet.getDataRange().getValues();
    const activities = data.slice(1).map(row => ({
        id: row[0],
        salesRep: row[1] || '',
        datetime: row[2],
        type: row[3],
        company: row[4],
        department: row[5],
        contact: row[6],
        reaction: row[7],
        met: row[8],
        note: row[9],
        proposals: [row[10], row[11], row[12], row[13], row[14]].filter(p => p)
    })).reverse();

    return { success: true, data: activities };
}

// ========================================
// 活動記録追加（別スプレッドシート対応版）
// ========================================
function addActivity(data) {
    const ss = getActivityLogSpreadsheet();
    let sheet = ss.getSheetByName('活動記録');

    // シートがなければ作成
    if (!sheet) {
        sheet = ss.insertSheet('活動記録');
        sheet.appendRow(['ID', '当社担当', '日時', '種類', '会社名', '部署', '先方担当', '反応', '会えた', '内容',
                         '提案商品1', '提案商品2', '提案商品3', '提案商品4', '提案商品5']);
    }

    // 担当者リストを取得（後方互換：contactがあればそれを使用、なければcontacts配列）
    let contactList = [];
    if (data.contacts && Array.isArray(data.contacts) && data.contacts.length > 0) {
        contactList = data.contacts;
    } else if (data.contact) {
        contactList = [data.contact];
    }

    if (contactList.length === 0) {
        return { success: false, error: '担当者が指定されていません' };
    }

    // 提案商品を分割（最大5つ）
    const proposals = data.proposals || [];

    // 各担当者ごとに1行登録
    contactList.forEach(contact => {
        // 次のIDを取得
        const lastRow = sheet.getLastRow();
        const nextId = lastRow > 1 ? lastRow : 1;

        sheet.appendRow([
            nextId,
            data.salesRep || '',
            data.datetime,
            data.type,
            data.company,
            data.department,
            contact, // 各担当者ごとに登録
            data.reaction,
            data.met,
            data.note,
            proposals[0] || '',
            proposals[1] || '',
            proposals[2] || '',
            proposals[3] || '',
            proposals[4] || ''
        ]);
    });

    return { success: true, count: contactList.length };
}

// ========================================
// 訪問予定取得（年月指定）
// ========================================
function getVisitSchedule(year, month) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('訪問日マスタ');

    if (!sheet) {
        return { success: false, error: '訪問日マスタが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const schedules = [];

    // ヘッダー: A列=担当, B列=訪問予定日, C列=訪問時間, D列=訪問先
    for (let i = 1; i < data.length; i++) {
        const rep = data[i][0];
        let visitDate = data[i][1];
        const timeSlot = data[i][2];
        const destination = data[i][3];

        if (!rep || !visitDate) continue;

        // 日付処理
        if (visitDate instanceof Date) {
            const y = visitDate.getFullYear();
            const m = visitDate.getMonth() + 1;
            
            // 指定年月のデータのみ
            if (String(y) === String(year) && String(m) === String(month)) {
                schedules.push({
                    rowIndex: i + 1, // スプレッドシートの行番号（1始まり）
                    担当: rep,
                    訪問予定日: Utilities.formatDate(visitDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
                    訪問時間: timeSlot || '',
                    訪問先: destination || ''
                });
            }
        }
    }

    return { success: true, data: schedules };
}

// ========================================
// 訪問予定更新/追加
// ========================================
function updateVisitSchedule(data) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('訪問日マスタ');

    // シートがなければ作成
    if (!sheet) {
        sheet = ss.insertSheet('訪問日マスタ');
        sheet.appendRow(['担当', '訪問予定日', '訪問時間', '訪問先', '訪問日']);
    }

    const { 担当, 訪問予定日, 訪問時間, 訪問先 } = data;
    
    if (!担当 || !訪問予定日) {
        return { success: false, error: '担当と訪問予定日は必須です' };
    }

    // 既存データを検索
    const allData = sheet.getDataRange().getValues();
    let foundRow = -1;

    for (let i = 1; i < allData.length; i++) {
        const rowRep = allData[i][0];
        let rowDate = allData[i][1];

        if (rowDate instanceof Date) {
            rowDate = Utilities.formatDate(rowDate, 'Asia/Tokyo', 'yyyy-MM-dd');
        }

        if (rowRep === 担当 && rowDate === 訪問予定日) {
            foundRow = i + 1;
            break;
        }
    }

    // 日付オブジェクトを作成（時間はリセット）
    const dateParts = 訪問予定日.split('-');
    const dateObj = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));

    if (foundRow > 0) {
        // 既存レコードを更新
        sheet.getRange(foundRow, 2).setValue(dateObj); // 日付も更新（時間なし）
        sheet.getRange(foundRow, 3).setValue(訪問時間);
        sheet.getRange(foundRow, 4).setValue(訪問先);
    } else {
        // 新規追加
        sheet.appendRow([担当, dateObj, 訪問時間, 訪問先, '']);
    }

    return { success: true };
}

// ========================================
// 訪問予定削除
// ========================================
function deleteVisitSchedule(data) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('訪問日マスタ');

    if (!sheet) {
        return { success: false, error: '訪問日マスタが見つかりません' };
    }

    const { 担当, 訪問予定日 } = data;
    
    if (!担当 || !訪問予定日) {
        return { success: false, error: '担当と訪問予定日は必須です' };
    }

    // 既存データを検索
    const allData = sheet.getDataRange().getValues();
    let foundRow = -1;

    for (let i = 1; i < allData.length; i++) {
        const rowRep = allData[i][0];
        let rowDate = allData[i][1];

        if (rowDate instanceof Date) {
            rowDate = Utilities.formatDate(rowDate, 'Asia/Tokyo', 'yyyy-MM-dd');
        }

        if (rowRep === 担当 && rowDate === 訪問予定日) {
            foundRow = i + 1;
            break;
        }
    }

    if (foundRow > 0) {
        // 行を削除
        sheet.deleteRow(foundRow);
        return { success: true };
    } else {
        return { success: false, error: '該当する予定が見つかりません' };
    }
}

// ========================================
// 目標マスタから指定年月の目標を取得（日付型対応版）
// ========================================
function getGoals(yearMonth) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('目標マスタ');

    if (!sheet) {
        return { success: false, error: '目標マスタシートが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const header = data[0];

    // 年月の列を探す
    let targetColIndex = -1;
    for (let i = 1; i < header.length; i++) {
        let headerValue = header[i];
        
        // 日付型の場合は「YYYY年M月」形式に変換
        if (headerValue instanceof Date) {
            headerValue = headerValue.getFullYear() + '年' + (headerValue.getMonth() + 1) + '月';
        }
        
        if (headerValue === yearMonth) {
            targetColIndex = i;
            break;
        }
    }

    if (targetColIndex === -1) {
        return { success: true, data: {} };
    }

    const goals = {};
    for (let i = 1; i < data.length; i++) {
        const company = data[i][0];
        const target = data[i][targetColIndex];
        if (company && company !== '合計' && target) {
            goals[company] = parseInt(target) || 0;
        }
    }

    return { success: true, data: goals };
}

// ========================================
// 提案商品マスタから指定年月の商品を取得
// ========================================
function getProposalProducts(yearMonth) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('提案商品マスタ');

    if (!sheet) {
        return { success: false, error: '提案商品マスタが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const products = [];

    for (let i = 1; i < data.length; i++) {
        let rowYearMonth = data[i][0];
        
        // 日付型の場合は変換
        if (rowYearMonth instanceof Date) {
            rowYearMonth = rowYearMonth.getFullYear() + '年' + (rowYearMonth.getMonth() + 1) + '月';
        }
        
        if (rowYearMonth === yearMonth && data[i][1]) {
            products.push(data[i][1]);
        }
    }

    return { success: true, data: products };
}

// ========================================
// 販促資料一覧を取得
// ========================================
function getMaterials() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('販促シート');

        if (!sheet) {
            return { success: false, error: '販促シートが見つかりません' };
        }

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) {
            return { success: true, data: [] };
        }

        // A列:ID, B列:資料名, C列:YouTube, D列:PDF, E列:リンク, F列:カテゴリ
        const range = sheet.getRange(2, 1, lastRow - 1, 6);
        const values = range.getValues();

        const materials = values
            .filter(row => row[0] && row[1]) // IDと資料名があるもの
            .map(row => ({
                id: String(row[0]),
                title: row[1],
                youtubeLink: row[2] || '',
                pdfLink: row[3] || '',
                link: row[4] || '',      // E列: リンク
                category: row[5] || ''
            }));

        return { success: true, data: materials };

    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ========================================
// 当社担当一覧取得
// ========================================
function getOurReps() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('当社担当マスタ');

    if (!sheet) {
        return { success: false, error: '当社担当マスタが見つかりません' };
    }

    const data = sheet.getDataRange().getValues();
    const reps = [];

    for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
            reps.push({
                name: data[i][0],
                email: data[i][1] || ''
            });
        }
    }

    return { success: true, data: reps };
}

// ========================================
// 販促メール送信
// ========================================
function sendPromotionEmail(payload) {
    try {
        const { materials, recipient, ourRep } = payload;

        let materialList = '';
        materials.forEach(m => {
            materialList += `━━━━━━━━━━━━━━━━━━━━━━━━\n`;
            materialList += `📄 ${m.title}\n`;
            if (m.youtubeLink) {
                materialList += `  🎬 YouTube: ${m.youtubeLink}\n`;
            }
            if (m.pdfLink) {
                materialList += `  📎 PDF: ${m.pdfLink}\n`;
            }
            materialList += `\n`;
        });

        const body = `${recipient.company}
${recipient.department}
${recipient.contact} 様

いつもお世話になっております。
アライ電機産業 特販部の${ourRep.name}です。

この度、以下の資料をご案内させていただきます。

${materialList}
ご確認いただき、ご不明点等ございましたら
お気軽にお問い合わせください。

今後ともよろしくお願いいたします。

────────────────────────────
${ourRep.name}
アライ電機産業株式会社 特販部
Email: ${ourRep.email}
TEL: 048-458-7004
FAX: 048-458-7315
────────────────────────────`;

        MailApp.sendEmail({
            to: recipient.email,
            cc: ourRep.email,
            replyTo: ourRep.email,
            subject: '【アライ電機産業】カーナビ・ETC関連資料のご案内',
            body: body,
            name: ourRep.name + ' / アライ電機産業 特販部'
        });

        logEmailSent(recipient, ourRep, materials);

        return { success: true, message: 'メール送信完了' };

    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// メール送信ログ記録
// ========================================
function logEmailSent(recipient, ourRep, materials) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('送信ログ');

    if (!sheet) {
        sheet = ss.insertSheet('送信ログ');
        sheet.appendRow(['送信日時', '当社担当', '送付先会社', '送付先部署', '送付先担当者', '送付先メール', '送付資料']);
    }

    const materialTitles = materials.map(m => m.title).join(', ');

    sheet.appendRow([
        new Date(),
        ourRep.name,
        recipient.company,
        recipient.department,
        recipient.contact,
        recipient.email,
        materialTitles
    ]);
}

// ========================================
// テスト用：メール送信権限確認
// ========================================
function testSendEmail() {
    MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: '【テスト】特販部 SALES PRO メール送信テスト',
        body: 'このメールが届いていれば、メール送信機能は正常に動作しています。'
    });
    Logger.log('テストメール送信完了');
}

// ========================================
// 過去実績データ取得
// ========================================
function getPerformanceData() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('過去実績マスタ');

        if (!sheet) {
            return { success: false, error: '過去実績マスタシートが見つかりません' };
        }

        const data = sheet.getDataRange().getValues();
        const header = data[0];

        // 列インデックスを取得（0始まり）
        // D: 代表受注伝票番号 (index 3)
        // J: 得意先名 (index 9)
        // K: 得意先担当者名 (index 10)
        // L: 顧客名_漢字 (index 11)
        // Z: 税抜合計 (index 25)
        // AB: 登録日 (index 27)
        const COL_ORDER_NO = 3;      // D列: 代表受注伝票番号
        const COL_CUSTOMER = 9;      // J列: 得意先名
        const COL_REP = 10;          // K列: 得意先担当者名
        const COL_CLIENT = 11;       // L列: 顧客名_漢字
        const COL_AMOUNT = 25;       // Z列: 税抜合計
        const COL_DATE = 27;         // AB列: 登録日

        // データを集計（受注年月、得意先名、得意先担当者名、顧客名_漢字）でグループ化
        const aggregated = {};
        const processedOrders = new Set(); // 代表受注伝票番号の重複チェック用

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const orderNo = row[COL_ORDER_NO];
            const customerName = row[COL_CUSTOMER] || '';
            const customerRep = row[COL_REP] || '';
            const clientName = row[COL_CLIENT] || '';
            const amount = Number(row[COL_AMOUNT]) || 0;
            let regDate = row[COL_DATE];

            // 登録日から年月を抽出
            let yearMonth = '';
            if (regDate instanceof Date) {
                const y = regDate.getFullYear();
                const m = regDate.getMonth() + 1;
                yearMonth = y + '/' + (m < 10 ? '0' + m : m);
            } else if (regDate) {
                // 文字列の場合（例: 2025/10/1）
                const parts = String(regDate).split(/[\/-]/);
                if (parts.length >= 2) {
                    yearMonth = parts[0] + '/' + (parts[1].length === 1 ? '0' + parts[1] : parts[1]);
                }
            }

            if (!yearMonth || !customerName) continue;

            // 集計キー
            const key = `${yearMonth}_${customerName}_${customerRep}_${clientName}`;

            if (!aggregated[key]) {
                aggregated[key] = {
                    orderYearMonth: yearMonth,
                    customerName: customerName,
                    customerRep: customerRep,
                    clientName: clientName,
                    orderCount: 0,
                    salesAmount: 0,
                    processedOrderNos: new Set()
                };
            }

            // 売上金額を加算
            aggregated[key].salesAmount += amount;

            // 代表受注伝票番号ごとに1件カウント（同一キー内での重複を除外）
            if (orderNo && !aggregated[key].processedOrderNos.has(orderNo)) {
                aggregated[key].orderCount += 1;
                aggregated[key].processedOrderNos.add(orderNo);
            }
        }

        // 結果を配列に変換（processedOrderNosは除外）
        const result = Object.values(aggregated).map(item => ({
            orderYearMonth: item.orderYearMonth,
            customerName: item.customerName,
            customerRep: item.customerRep,
            clientName: item.clientName,
            orderCount: item.orderCount,
            salesAmount: item.salesAmount
        }));

        // 受注年月の降順でソート
        result.sort((a, b) => b.orderYearMonth.localeCompare(a.orderYearMonth));

        return { success: true, data: result };

    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// 過去実績生データ取得（詳細分析用）
// ========================================
function getPerformanceRawData() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('過去実績マスタ');

        if (!sheet) {
            return { success: false, error: '過去実績マスタシートが見つかりません' };
        }

        const data = sheet.getDataRange().getValues();

        // 列インデックス（0始まり）
        // D: 代表受注伝票番号 (index 3)
        // K: 得意先担当者名 (index 10)
        // L: 顧客名_漢字 (index 11)
        // N: 車種名 (index 13)
        // Q: 商品大分類 (index 16)
        // R: 商品中分類 (index 17)
        // S: 商品小分類 (index 18)
        // T: メーカーCD (index 19)
        // U: 品番 (index 20)
        const COL_ORDER_NO = 3;
        const COL_REP = 10;
        const COL_CLIENT = 11;
        const COL_VEHICLE = 13;
        const COL_PRODUCT_MAJOR = 16;
        const COL_PRODUCT_MIDDLE = 17;
        const COL_PRODUCT_MINOR = 18;
        const COL_MAKER = 19;
        const COL_PRODUCT_CODE = 20;

        const result = [];

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            result.push({
                orderNo: row[COL_ORDER_NO],
                customerRep: row[COL_REP] || '',
                clientName: row[COL_CLIENT] || '',
                vehicleName: row[COL_VEHICLE] || '',
                productMajor: row[COL_PRODUCT_MAJOR],
                productMiddle: row[COL_PRODUCT_MIDDLE] || '',
                productMinor: row[COL_PRODUCT_MINOR] || '',
                makerCode: String(row[COL_MAKER] || ''),
                productCode: row[COL_PRODUCT_CODE] || ''
            });
        }

        return { success: true, data: result };

    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// 売上分析用データ取得（必要列のみ取得で高速化）
// ========================================
function getSalesAnalysisData() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('見積実績マスタ');

        if (!sheet) {
            return { success: false, error: '見積実績マスタシートが見つかりません' };
        }

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) {
            return { success: true, data: [] };
        }

        const numRows = lastRow - 1; // ヘッダー除く

        // 必要な列のみを個別に取得（全列取得を避けて高速化）
        // J列(10): 得意先名, K列(11): 担当者名, L列(12): 顧客名
        const rangeJL = sheet.getRange(2, 10, numRows, 3).getValues();
        
        // U列(21): 品番, V列(22): 商品名, W列(23): 数量, X列(24): 単価
        const rangeUX = sheet.getRange(2, 21, numRows, 4).getValues();
        
        // Z列(26): 税抜合計
        const rangeZ = sheet.getRange(2, 26, numRows, 1).getValues();
        
        // AA列(27): 登録日
        const rangeAA = sheet.getRange(2, 27, numRows, 1).getValues();
        
        // AD列(30): 略称, AE列(31): 部店, AF列(32): 本社
        const rangeADAF = sheet.getRange(2, 30, numRows, 3).getValues();

        const result = [];

        for (let i = 0; i < numRows; i++) {
            // 登録日から年月を抽出
            let registYearMonth = '';
            const registDate = rangeAA[i][0];
            if (registDate instanceof Date) {
                const y = registDate.getFullYear();
                const m = registDate.getMonth() + 1;
                registYearMonth = `${y}/${m < 10 ? '0' + m : m}`;
            } else if (registDate) {
                const dateStr = String(registDate);
                const parts = dateStr.split(/[\/\-]/);
                if (parts.length >= 2) {
                    const y = parts[0];
                    const m = parts[1].padStart(2, '0');
                    registYearMonth = `${y}/${m}`;
                }
            }
            
            result.push({
                customerName: rangeJL[i][0] || '',
                repName: rangeJL[i][1] || '',
                clientName: rangeJL[i][2] || '',
                productCode: rangeUX[i][0] || '',
                productName: rangeUX[i][1] || '',
                quantity: Number(rangeUX[i][2]) || 0,
                unitPrice: Number(rangeUX[i][3]) || 0,
                total: Number(rangeZ[i][0]) || 0,
                registYearMonth: registYearMonth,
                abbr: rangeADAF[i][0] || '',
                branch: rangeADAF[i][1] || '',
                hqFlag: rangeADAF[i][2] === '●'
            });
        }

        return { success: true, data: result };

    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// シート最終更新日時取得（キャッシュ判定用）
// ========================================
function getSheetLastModified() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const file = DriveApp.getFileById(ss.getId());
        const lastUpdated = file.getLastUpdated();
        
        return { 
            success: true, 
            data: { lastModified: lastUpdated.toISOString() } 
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// 顧客TEL情報取得（顧客マスタのC列:部署→D列:TEL）
// 部署の順序リストも返す（スプレッドシートの並び順）
// ========================================
function getCustomerPhones() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('顧客マスタ');

        if (!sheet) {
            return { success: false, error: '顧客マスタが見つかりません' };
        }

        const data = sheet.getDataRange().getValues();
        const phones = {};
        const branchOrder = []; // 順序付きリスト

        // B列: 会社名, C列: 部署
        // 電話番号の参照は廃止（列が削除されるため）
        for (let i = 1; i < data.length; i++) {
            const department = data[i][2] || ''; // C列 (index 2): 部署
            
            if (department) {
                // 重複を避けて順序を維持
                if (!branchOrder.includes(department)) {
                    branchOrder.push(department);
                }
            }
        }

        return { success: true, data: { phones, branchOrder } };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// 新規担当者登録（Webから登録 = ★マーク付与）
// ========================================
function addContact(data) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('担当者マスタ');

        if (!sheet) {
            return { success: false, error: '担当者マスタが見つかりません' };
        }

        const { company, department, contactName } = data;

        if (!company || !department || !contactName) {
            return { success: false, error: '会社名、部署、担当者名は必須です' };
        }

        // 新規行を追加（新しい列構成: A:会社名, B:部署, C:役職, D:担当者名, E:営業担当）
        sheet.appendRow([
            company,    // A列: 会社名
            department, // B列: 部署
            '',         // C列: 役職
            contactName,// D列: 担当者名
            ''          // E列: 営業担当
        ]);

        return { 
            success: true, 
            message: `${contactName} を登録しました（★Web登録）` 
        };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// 活動記録更新
// ========================================
function updateActivity(data) {
    try {
        const ss = getActivityLogSpreadsheet();
        const sheet = ss.getSheetByName('活動記録');

        if (!sheet) {
            return { success: false, error: '活動記録シートが見つかりません' };
        }

        const { id, datetime, type, company, department, contact, reaction, met, note, proposals, salesRep } = data;

        if (!id) {
            return { success: false, error: 'IDが指定されていません' };
        }

        // IDで行を検索
        const allData = sheet.getDataRange().getValues();
        let targetRow = -1;

        for (let i = 1; i < allData.length; i++) {
            if (String(allData[i][0]) === String(id)) {
                targetRow = i + 1; // スプレッドシートの行番号（1始まり）
                break;
            }
        }

        if (targetRow === -1) {
            return { success: false, error: '指定されたIDの活動記録が見つかりません' };
        }

        // 提案商品を分割（最大5つ）
        const proposalArray = proposals || [];

        // 行を更新（A列:ID, B列:当社担当, C列:日時, D列:種類, E列:会社名, F列:部署, G列:先方担当, H列:反応, I列:会えた, J列:内容, K-O列:提案商品）
        sheet.getRange(targetRow, 2, 1, 14).setValues([[
            salesRep || '',
            datetime,
            type,
            company,
            department,
            contact,
            reaction,
            met,
            note,
            proposalArray[0] || '',
            proposalArray[1] || '',
            proposalArray[2] || '',
            proposalArray[3] || '',
            proposalArray[4] || ''
        ]]);

        return { success: true, message: '活動記録を更新しました' };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// 活動記録削除
// ========================================
function deleteActivity(data) {
    try {
        const ss = getActivityLogSpreadsheet();
        const sheet = ss.getSheetByName('活動記録');

        if (!sheet) {
            return { success: false, error: '活動記録シートが見つかりません' };
        }

        const { id } = data;

        if (!id) {
            return { success: false, error: 'IDが指定されていません' };
        }

        // IDで行を検索
        const allData = sheet.getDataRange().getValues();
        let targetRow = -1;

        for (let i = 1; i < allData.length; i++) {
            if (String(allData[i][0]) === String(id)) {
                targetRow = i + 1;
                break;
            }
        }

        if (targetRow === -1) {
            return { success: false, error: '指定されたIDの活動記録が見つかりません' };
        }

        // 行を削除
        sheet.deleteRow(targetRow);

        return { success: true, message: '活動記録を削除しました' };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// 担当者と最終活動日を取得（アタックリスト用）
// ========================================
function getContactsWithLastActivity() {
    try {
        const masterSS = SpreadsheetApp.openById(SPREADSHEET_ID);
        const activitySS = getActivityLogSpreadsheet();
        const contactSheet = masterSS.getSheetByName('担当者マスタ');
        const activitySheet = activitySS.getSheetByName('活動記録');

        if (!contactSheet) {
            return { success: false, error: '担当者マスタが見つかりません' };
        }

        const contactData = contactSheet.getDataRange().getValues();
        const activityData = activitySheet ? activitySheet.getDataRange().getValues() : [];

        // 活動記録から担当者ごとの最終「会えた」日を抽出
        const lastMetDates = {};
        for (let i = 1; i < activityData.length; i++) {
            const contact = activityData[i][6];  // G列: 先方担当
            const met = activityData[i][8];      // I列: 会えた
            const datetime = activityData[i][2]; // C列: 日時
            const company = activityData[i][4];  // E列: 会社名
            const department = activityData[i][5]; // F列: 部署

            if (contact && met === '○' && datetime) {
                const key = `${company}_${department}_${contact}`;
                const date = new Date(datetime);
                if (!lastMetDates[key] || date > lastMetDates[key]) {
                    lastMetDates[key] = date;
                }
            }
        }

        const today = new Date();
        const contacts = [];
        const seenContacts = new Set();

        for (let i = 1; i < contactData.length; i++) {
            const company = contactData[i][0];     // A列: 会社名
            const department = contactData[i][1];  // B列: 部署
            const title = contactData[i][2];       // C列: 役職
            const contactName = contactData[i][3]; // D列: 担当者名

            if (!company || !contactName) continue;

            // A列のIDは使わず、自動生成
            const id = `${company}_${department}_${contactName}`;
            
            const key = `${company}_${department}_${contactName}`;
            if (seenContacts.has(key)) continue;
            seenContacts.add(key);

            const lastMet = lastMetDates[key];
            let daysSince = null;
            let status = 'distant'; // デフォルト: 疎遠（赤）

            if (lastMet) {
                daysSince = Math.floor((today - lastMet) / (1000 * 60 * 60 * 24));
                if (daysSince <= 30) {
                    status = 'good';     // 緑: 30日以内
                } else if (daysSince <= 60) {
                    status = 'warning';  // 黄: 31-60日
                } else {
                    status = 'distant';  // 赤: 61日以上
                }
            }

            contacts.push({
                id: id,
                company: company,
                department: department,
                contactName: contactName,
                lastActivityDate: lastMet ? Utilities.formatDate(lastMet, 'Asia/Tokyo', 'yyyy-MM-dd') : null,
                daysSince: daysSince,
                status: status
            });
        }

        // 部署順、ステータス順でソート
        contacts.sort((a, b) => {
            if (a.company !== b.company) return a.company.localeCompare(b.company);
            if (a.department !== b.department) return a.department.localeCompare(b.department);
            // ステータス順: distant > warning > good
            const statusOrder = { 'distant': 0, 'warning': 1, 'good': 2 };
            return statusOrder[a.status] - statusOrder[b.status];
        });

        return { success: true, data: contacts };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// アクションリスト取得
// ========================================
function getActionList(yearMonth) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('アクションリスト');

        if (!sheet) {
            return { success: true, data: [] };
        }

        const data = sheet.getDataRange().getValues();
        const actions = [];

        for (let i = 1; i < data.length; i++) {
            const rowYearMonth = data[i][0]; // A列: 年月
            if (String(rowYearMonth) === String(yearMonth) || !yearMonth) {
                actions.push({
                    id: data[i][1],           // B列: 担当者ID
                    company: data[i][2],      // C列: 会社名
                    department: data[i][3],   // D列: 部署
                    contactName: data[i][4],  // E列: 担当者名
                    createdAt: data[i][5],    // F列: 登録日時
                    status: data[i][6]        // G列: ステータス
                });
            }
        }

        return { success: true, data: actions };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// アクションリスト保存
// ========================================
function saveActionList(data) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let sheet = ss.getSheetByName('アクションリスト');

        if (!sheet) {
            sheet = ss.insertSheet('アクションリスト');
            sheet.appendRow(['年月', '担当者ID', '会社名', '部署', '担当者名', '登録日時', 'ステータス']);
        }

        const { yearMonth, contacts } = data;
        const now = new Date();

        // 既存の同月データを削除
        const existingData = sheet.getDataRange().getValues();
        for (let i = existingData.length - 1; i >= 1; i--) {
            if (String(existingData[i][0]) === String(yearMonth)) {
                sheet.deleteRow(i + 1);
            }
        }

        // 新しいデータを追加
        contacts.forEach(contact => {
            sheet.appendRow([
                yearMonth,
                contact.id,
                contact.company,
                contact.department,
                contact.contactName,
                now,
                'pending'
            ]);
        });

        return { success: true, count: contacts.length };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// アクションリストから削除
// ========================================
function removeFromActionList(data) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('アクションリスト');

        if (!sheet) {
            return { success: false, error: 'アクションリストが見つかりません' };
        }

        const { id, yearMonth } = data;
        const existingData = sheet.getDataRange().getValues();

        for (let i = existingData.length - 1; i >= 1; i--) {
            if (String(existingData[i][0]) === String(yearMonth) && 
                String(existingData[i][1]) === String(id)) {
                sheet.deleteRow(i + 1);
                return { success: true };
            }
        }

        return { success: false, error: '該当するアクションが見つかりません' };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// 定期バックアップ（Excelメール送信）
// 毎日8時と18時に takano@araidenki.com へ送信
// ========================================
const BACKUP_EMAIL_RECIPIENT = 'takano@araidenki.com';

function dailyBackupToExcel() {
    try {
        const recipient = BACKUP_EMAIL_RECIPIENT;

        // 活動ログとARAI_SALES_PROの両方のIDを指定
        const targetSsIds = [ACTIVITY_LOG_SPREADSHEET_ID, SPREADSHEET_ID];
        
        const now = new Date();
        const todayStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyyMMdd");
        const timeStr = Utilities.formatDate(now, "Asia/Tokyo", "HH:mm");
        
        const blobs = [];
        const ssNames = [];

        const options = {
            headers: {
                Authorization: "Bearer " + ScriptApp.getOAuthToken()
            },
            muteHttpExceptions: true
        };

        // 各スプレッドシートをExcelとしてダウンロード
        for (const targetId of targetSsIds) {
            const targetSs = SpreadsheetApp.openById(targetId);
            const ssName = targetSs.getName();
            ssNames.push(ssName);
            console.log(`バックアップ開始: ${ssName}`);

            const url = "https://docs.google.com/spreadsheets/d/" + targetId + "/export?format=xlsx";
            const response = UrlFetchApp.fetch(url, options);
            
            if (response.getResponseCode() !== 200) {
                throw new Error(`${ssName} のダウンロードに失敗しました (Status: ${response.getResponseCode()})`);
            }
            
            const blob = response.getBlob().setName(`${ssName}_${todayStr}_${timeStr.replace(':', '')}.xlsx`);
            blobs.push(blob);
        }
        
        const subject = `【営業活動レポート＆データバックアップ】${Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm")}`;
        const body = `定期レポートおよびバックアップデータです。

対象データ:
${ssNames.map(name => "・" + name).join('\n')}

送信時刻: ${Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm")}

本メールはGoogle Apps Scriptによる自動送信です。
毎日 8:00 と 18:00 に送信されます。`;
        
        MailApp.sendEmail({
            to: recipient,
            subject: subject,
            body: body,
            attachments: blobs
        });
        
        console.log(`バックアップメールを送信しました: ${recipient} (${timeStr})`);
        
    } catch (e) {
        console.error("バックアップ処理エラー:", e);
    }
}

// ========================================
// トリガー設定（8時と18時に自動実行）
// GASエディタで一度だけ手動実行してください
// ========================================
function setupDailyBackupTriggers() {
    // 既存のトリガーを削除（重複防止）
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
        if (trigger.getHandlerFunction() === 'dailyBackupToExcel') {
            ScriptApp.deleteTrigger(trigger);
            console.log('既存のトリガーを削除しました');
        }
    });

    // 毎日8時に実行
    ScriptApp.newTrigger('dailyBackupToExcel')
        .timeBased()
        .atHour(8)
        .everyDays(1)
        .inTimezone('Asia/Tokyo')
        .create();
    console.log('8時のトリガーを設定しました');

    // 毎日18時に実行
    ScriptApp.newTrigger('dailyBackupToExcel')
        .timeBased()
        .atHour(18)
        .everyDays(1)
        .inTimezone('Asia/Tokyo')
        .create();
    console.log('18時のトリガーを設定しました');

    SpreadsheetApp.getUi().alert('トリガーを設定しました！\n\n毎日 8:00 と 18:00 に\ntakano@araidenki.com へ\n営業活動レポートが送信されます。');
}

// ========================================
// テスト用：即座にバックアップメールを送信
// ========================================
function testBackupEmail() {
    dailyBackupToExcel();
    SpreadsheetApp.getUi().alert(`テストメールを ${BACKUP_EMAIL_RECIPIENT} に送信しました。\n受信を確認してください。`);
}

// ========================================
// 受注分析用データ取得（チャンク取得 ＆ 必要列のみ取得で高速化）
// ========================================
function getOrderAnalysisData(offsetStr, limitStr) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('受注実績マスタ'); // 受注データシート名にあわせる

        if (!sheet) {
            return { success: false, error: '受注実績マスタシートが見つかりません' };
        }

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) {
            return { success: true, data: [], total: 0 };
        }

        const numRows = lastRow - 1; // ヘッダー除く
        const totalRows = numRows;

        // パラメータを数値化（指定がない場合は全件取得の挙動）
        const offset = offsetStr ? parseInt(offsetStr, 10) : 0;
        const limit = limitStr ? parseInt(limitStr, 10) : totalRows;

        // オフセットがデータ範囲を超えている場合
        if (offset >= totalRows) {
            return { success: true, data: [], total: totalRows };
        }

        // このチャンクで取得する行数を計算
        const fetchRows = Math.min(limit, totalRows - offset);
        
        // シートから取得する行はヘッダー分(+1)とoffset分をずらす
        const startRow = 2 + offset;

        // ※ 実際のシートの列構成に合わせて調整が必要です。
        // ここでは見積実績マスタとあわせた例としています（J,K,L,U,V,W,X,Z,AA,AD,AE,AF列）
        // 受注実績マスタの列構成が異なる場合は、列インデックス(getRangeの第2引数)を変更してください。

        // J列(10): 得意先名, K列(11): 担当者名, L列(12): 顧客名
        const rangeJL = sheet.getRange(startRow, 10, fetchRows, 3).getValues();
        
        // U列(21): 品番, V列(22): 商品名, W列(23): 数量, X列(24): 単価
        const rangeUX = sheet.getRange(startRow, 21, fetchRows, 4).getValues();
        
        // Z列(26): 税抜合計
        const rangeZ = sheet.getRange(startRow, 26, fetchRows, 1).getValues();
        
        // AA列(27): 登録日
        const rangeAA = sheet.getRange(startRow, 27, fetchRows, 1).getValues();
        
        // AD列(30): 略称, AE列(31): 部店, AF列(32): 本社
        const rangeADAF = sheet.getRange(startRow, 30, fetchRows, 3).getValues();

        const result = [];

        for (let i = 0; i < fetchRows; i++) {
            // 登録日から年月を抽出
            let registYearMonth = '';
            const registDate = rangeAA[i][0];
            if (registDate instanceof Date) {
                const y = registDate.getFullYear();
                const m = registDate.getMonth() + 1;
                registYearMonth = `${y}/${m < 10 ? '0' + m : m}`;
            } else if (registDate) {
                const dateStr = String(registDate);
                const parts = dateStr.split(/[\/\-]/);
                if (parts.length >= 2) {
                    const y = parts[0];
                    const m = parts[1].padStart(2, '0');
                    registYearMonth = `${y}/${m}`;
                }
            }
            
            result.push({
                customerName: rangeJL[i][0] || '',
                repName: rangeJL[i][1] || '',
                clientName: rangeJL[i][2] || '',
                productCode: rangeUX[i][0] || '',
                productName: rangeUX[i][1] || '',
                quantity: Number(rangeUX[i][2]) || 0,
                unitPrice: Number(rangeUX[i][3]) || 0,
                total: Number(rangeZ[i][0]) || 0,
                registYearMonth: registYearMonth,
                abbr: rangeADAF[i][0] || '',
                branch: rangeADAF[i][1] || '',
                hqFlag: rangeADAF[i][2] === '●'
            });
        }

        return { 
            success: true, 
            data: result,
            total: totalRows
        };

    } catch (error) {
        return { success: false, error: error.message };
    }
}

// ========================================
// 営業日報取得
// ========================================
function getDailyReports(salesRep, limit) {
    try {
        var ss = getDailyReportSpreadsheet();
        var sheet = ss.getSheetByName('営業日報');

        if (!sheet || sheet.getLastRow() < 2) {
            return { success: true, data: [] };
        }

        var data = sheet.getDataRange().getValues();
        var reportsMap = {};

        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var date = row[0];
            var rep = row[1];
            var timeSlot = row[2];
            var content = row[3];

            if (!date || !rep) continue;
            if (salesRep && rep !== salesRep) continue;

            var dateStr;
            if (date instanceof Date) {
                dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
            } else {
                dateStr = String(date).substring(0, 10);
            }

            var key = dateStr + '_' + rep;
            if (!reportsMap[key]) {
                reportsMap[key] = { date: dateStr, salesRep: rep, amContent: '', pmContent: '' };
            }

            if (timeSlot === 'AM') {
                reportsMap[key].amContent = content ? String(content) : '';
            } else if (timeSlot === 'PM') {
                reportsMap[key].pmContent = content ? String(content) : '';
            }
        }

        var reports = Object.values(reportsMap).sort(function(a, b) {
            return b.date.localeCompare(a.date);
        });

        var maxRows = limit ? parseInt(limit) : 30;
        if (reports.length > maxRows) {
            reports = reports.slice(0, maxRows);
        }

        return { success: true, data: reports };

    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ========================================
// 営業日報保存（同日・同担当者は上書き）
// ========================================
function saveDailyReport(data) {
    try {
        var ss = getDailyReportSpreadsheet();
        var sheet = ss.getSheetByName('営業日報');

        if (!sheet) {
            sheet = ss.insertSheet('営業日報');
            sheet.appendRow(['年月日', '営業担当者', '作業時間', '項目']);
        }

        var salesRep = data.salesRep;
        var date = data.date;
        var amContent = data.amContent || '';
        var pmContent = data.pmContent || '';
        var overwrite = !!data.overwrite;

        if (!salesRep || !date) {
            return { success: false, error: '担当者と日付は必須です' };
        }

        var dateParts = date.split('-');
        var dateObj = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));

        // 既存の同日・同担当者データを探して削除対象にする
        var allData = sheet.getDataRange().getValues();
        var rowsToDelete = [];
        var existingAm = '';
        var existingPm = '';

        for (var i = 1; i < allData.length; i++) {
            var rowDate = allData[i][0];
            var rowRep = allData[i][1];
            var rowTime = allData[i][2];
            var rowContent = allData[i][3];

            if (rowDate instanceof Date) {
                rowDate = Utilities.formatDate(rowDate, 'Asia/Tokyo', 'yyyy-MM-dd');
            } else {
                rowDate = String(rowDate).substring(0, 10);
            }

            if (rowDate === date && rowRep === salesRep) {
                rowsToDelete.push(i + 1);
                if (rowTime === 'AM' && rowContent) {
                    existingAm = String(rowContent);
                } else if (rowTime === 'PM' && rowContent) {
                    existingPm = String(rowContent);
                }
            }
        }

        // 編集モード（overwrite）でなければ、送信内容が空の場合に既存データを保持し、入力がある場合は「追記」する（マージ処理）
        if (!overwrite) {
            if (amContent.trim() && existingAm) {
                amContent = existingAm + '\n\n' + amContent;
            } else if (!amContent.trim() && existingAm) {
                amContent = existingAm;
            }

            if (pmContent.trim() && existingPm) {
                pmContent = existingPm + '\n\n' + pmContent;
            } else if (!pmContent.trim() && existingPm) {
                pmContent = existingPm;
            }
        }

        for (var j = rowsToDelete.length - 1; j >= 0; j--) {
            sheet.deleteRow(rowsToDelete[j]);
        }

        if (amContent) {
            sheet.appendRow([dateObj, salesRep, 'AM', amContent]);
        }
        if (pmContent) {
            sheet.appendRow([dateObj, salesRep, 'PM', pmContent]);
        }

        return { success: true };

    } catch (e) {
        return { success: false, error: e.message };
    }
}

// ========================================
// 営業日報検索（複数キーワードAND検索）
// ========================================
function searchDailyReports(query) {
    try {
        if (!query) return { success: true, data: [] };
        
        // 全角スペースを半角に変換し、スペース区切りで配列にする (NFKCで全角半角カナ等を正規化)
        var keywords = String(query).normalize('NFKC').replace(/　/g, ' ').split(' ').filter(function(k) { return k.trim() !== ''; });
        if (keywords.length === 0) return { success: true, data: [] };
        
        var ss = getDailyReportSpreadsheet();
        var sheet = ss.getSheetByName('営業日報');
        if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };
        
        var data = sheet.getDataRange().getValues();
        var reportsMap = {};
        
        // getDailyReportsと同様に、同日・同担当者でAM/PMをまとめる
        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var date = row[0];
            var rep = row[1];
            var timeSlot = row[2];
            var content = row[3];
            
            if (!date || !rep) continue;
            
            var dateStr;
            if (date instanceof Date) {
                dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
            } else {
                dateStr = String(date).substring(0, 10);
            }
            
            var key = dateStr + '_' + rep;
            if (!reportsMap[key]) {
                reportsMap[key] = { date: dateStr, salesRep: rep, amContent: '', pmContent: '' };
            }
            
            if (timeSlot === 'AM') {
                reportsMap[key].amContent = content ? String(content) : '';
            } else if (timeSlot === 'PM') {
                reportsMap[key].pmContent = content ? String(content) : '';
            }
        }
        
        // 検索処理
        var allReports = Object.values(reportsMap);
        var matchedReports = [];
        
        for (var j = 0; j < allReports.length; j++) {
            var r = allReports[j];
            // 検索対象テキスト：担当名、AM内容、PM内容の結合 (NFKCで全角半角カナ等を正規化)
            var targetText = (r.salesRep + ' ' + (r.amContent || '') + ' ' + (r.pmContent || '')).normalize('NFKC').toLowerCase();
            var isMatch = true;
            
            // AND検索
            for (var k = 0; k < keywords.length; k++) {
                if (targetText.indexOf(keywords[k].toLowerCase()) === -1) {
                    isMatch = false;
                    break;
                }
            }
            if (isMatch) {
                matchedReports.push(r);
            }
        }
        
        // 降順ソート
        matchedReports.sort(function(a, b) {
            return b.date.localeCompare(a.date);
        });
        
        // 最大50件に制限
        if (matchedReports.length > 50) {
            matchedReports = matchedReports.slice(0, 50);
        }
        
        return { success: true, data: matchedReports };
    } catch (e) {
        return { success: false, error: e.message };
    }
}
