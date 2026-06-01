// =================================================================
//  基本設定區 (Global Constants)
// =================================================================

// --- SPREADSHEET & SHEET NAMES ---
const SPREADSHEET_ID                     = '1Wh8ZLzAcw5TrEIQS2U6p8DT9-g2-7W1tn5lqfLTohYw'; 
const REGISTER_SHEET_NAME                = '課程登錄紀錄';
const COURSE_LIST_SHEET_NAME             = '課程列表';
const VERIFICATION_SHEET_NAME            = '身分驗證申請';
const BUILDING_LIST_SHEET_NAME           = '社宅列表';
const VERIFICATION_STATUS_SHEET_NAME     = '通過驗證住戶資料';
const CONFIRMED_POINTS_SHEET_NAME        = '確認集點的資料'; 
const REWARD_LIST_SHEET_NAME             = '獎品清單';
const REDEMPTION_RECORD_SHEET_NAME       = '兌換紀錄';

// --- COLUMN INDEXES ---
const C_NAME_COLUMN                      = 1; 
const C_PHONE_COLUMN                     = 2; 
const C_COURSE_COLUMN                    = 3; 
const C_POINTS_COLUMN                    = 4; 
const VERIFY_NAME_COLUMN                 = 3; 
const VERIFY_PHONE_COLUMN                = 4; 

// =================================================================
//  主要功能函式 (Main Functions)
// =================================================================

/**
 * 顯示網頁的必要函式 (已升級與地圖相同的縮放規格)
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setTitle('生活學校集點系統');
}

/** 【功能二】登錄課程集點 */
function registerCourse(name, phone, course) {
  if (!name || !phone || !course) return '錯誤：所有欄位皆為必填。';
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(REGISTER_SHEET_NAME);
    if (!sheet) return `系統錯誤：找不到工作表。`;
    sheet.appendRow([new Date(), name, phone, course]);
    return `<b>已完成登錄成功！</b><br>姓名：${name}<br>課程：${course}<br>與開課單位核對確認簽到後才會進行點數累積！`;
  } catch (err) {
    return '登錄時發生系統錯誤，請聯繫管理員。';
  }
}

/** 【功能三】讀取課程清單 */
function getCourseList() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COURSE_LIST_SHEET_NAME);
    if (!sheet) return ['錯誤：找不到「課程列表」工作表'];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const values = sheet.getRange('A2:A' + lastRow).getValues();
    return values.flat().filter(String);
  } catch (err) {
    return ['錯誤：讀取課程時發生問題'];
  }
}

/** 【功能四】讀取社宅棟別清單 */
function getBuildingList() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(BUILDING_LIST_SHEET_NAME);
    if (!sheet) return ['錯誤：找不到「社宅列表」工作表'];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const values = sheet.getRange('A2:A' + lastRow).getValues();
    return values.flat().filter(String);
  } catch (err) {
    return ['錯誤：讀取棟別時發生問題'];
  }
}

/** 【功能五】接收身分驗證申請 */
function submitVerificationRequest(data) {
  if (!data.building || !data.unit || !data.name || !data.phone || !data.email) return '錯誤：所有欄位皆為必填。';
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(VERIFICATION_SHEET_NAME);
    if (!sheet) return `系統錯誤：找不到工作表。`;
    sheet.appendRow([new Date(), data.building, data.unit, data.name, data.phone, data.email]);
    return `<b>申請已送出！</b><br>將會在一週內進行人工審核，審核結果可於「驗證進度查詢」分頁進行。感謝您的申請！`;
  } catch (err) {
    return '申請時發生系統錯誤，請聯繫管理員。';
  }
}

/** 【功能六】查詢住戶身分驗證狀態 */
function checkVerificationStatus(name, phone) {
  if (!name || !phone) return '錯誤：姓名和手機為必填項目。';
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(VERIFICATION_STATUS_SHEET_NAME);
    if (!sheet) return `系統錯誤：找不到工作表。`;
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row[VERIFY_NAME_COLUMN - 1] === name && String(row[VERIFY_PHONE_COLUMN - 1]) === phone) {
        return `<b>已通過驗證！</b><br>姓名：${name}<br>電話：${phone}`;
      }
    }
    return '未通過驗證或資料輸入有誤';
  } catch (err) {
    return '查詢時發生系統錯誤，請聯繫管理員。';
  }
}

/** 【功能七】讀取獎品清單 */
function getRewardList() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(REWARD_LIST_SHEET_NAME);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    return values.map(row => ({ name: row[0], cost: Number(row[1]) })).filter(item => item.name && !isNaN(item.cost));
  } catch (err) {
    return [];
  }
}

/** 【功能八】取得使用者點數 */
function getUserPointsForRedeem(name, phone) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CONFIRMED_POINTS_SHEET_NAME);
    if (!sheet) return { success: false, message: '系統錯誤' };
    const values = sheet.getDataRange().getValues();
    let totalPoints = 0;
    const userRecords = [];
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][C_PHONE_COLUMN - 1]) === phone) {
        const points = Number(values[i][C_POINTS_COLUMN - 1]);
        if (!isNaN(points)) totalPoints += points;
        userRecords.push(values[i]);
      }
    }
    if (userRecords.length === 0) {
       return { success: false, message: `很抱歉，找不到手機號碼為「${phone}」的紀錄，請確認是否已完成「住戶資料驗證」。` };
    }
    let responseHtml = `您好，${name}！<br>此電話號碼(${phone})總累積點數為： <b>${totalPoints}</b> 點。<br><br><b><u>集點明細：</u></b><br>`;
    userRecords.forEach(function(record) {
      responseHtml += `&#8226; ${maskName(record[C_NAME_COLUMN - 1])} ｜ ${record[C_COURSE_COLUMN - 1]} [+${record[C_POINTS_COLUMN - 1]}點]<br>`;
    });
    return { success: true, points: totalPoints, queryResultHtml: responseHtml };
  } catch (err) {
    return { success: false, message: '查詢時發生錯誤' };
  }
}

/** 【功能九】送出兌換申請 */
function submitRedemption(data) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(REDEMPTION_RECORD_SHEET_NAME);
    if (!sheet) return '錯誤：找不到兌換紀錄表';
    const itemsString = data.items.map(i => `${i.name} x${i.quantity}`).join(', ');
    sheet.appendRow([new Date(), data.name, data.phone, data.email, data.address, itemsString, data.totalCost]);
    return '兌換申請成功！我們會盡快確認後為您寄出。';
  } catch (err) {
    return '兌換失敗：' + err.message;
  }
}

function maskName(name) {
  if (!name) return '';
  if (name.length <= 1) return name + '○';
  return name.substring(0, name.length - 1) + '○';
}