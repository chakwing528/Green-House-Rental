// ====== 請替換為你的 Google 試算表網址 ======
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1LEb_G-appATbb8IlO8_WIFemak__zwTQtXuzxy4jMeg/edit';

function getSheet(sheetName) {
  return SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName(sheetName);
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    let result = {};
    
    // 新架構：打開網頁時，一次性把「租借榜」、「歷史」與「全校學生名單」傳給前端
    if (action === 'getInitialData') {
      result = getInitialData();
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const data = payload.data;

    // 接收前端自己產生的時間，確保前後端時間完全一致
    if (action === 'borrowItem') {
      borrowItem(data);
    } else if (action === 'returnItem') {
      returnItem(data);
    }
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ================== 核心邏輯 ==================

function getInitialData() {
  const activeSheet = getSheet('Active');
  const historySheet = getSheet('History');
  const studentSheet = getSheet('StudentList');

  const activeData = activeSheet.getDataRange().getDisplayValues();
  const historyData = historySheet.getDataRange().getDisplayValues();
  const studentData = studentSheet.getDataRange().getDisplayValues();

  // 1. 整理正在租借
  const active = activeData.slice(1).map(row => ({
    classNum: row[0], studentNum: row[1], name: row[2], house: row[3], item: row[4], borrowTime: row[5]
  })).sort((a, b) => new Date(a.borrowTime) - new Date(b.borrowTime));

  // 2. 整理歷史紀錄
  const history = historyData.slice(1).map(row => ({
    classNum: row[0], studentNum: row[1], name: row[2], house: row[3], item: row[4], borrowTime: row[5], returnTime: row[6]
  })).reverse().slice(0, 30);

  // 3. 整理全校學生名單 (轉為物件格式方便前端秒查)
  let students = {};
  for (let i = 1; i < studentData.length; i++) {
    const cNum = studentData[i][0];
    const sNum = studentData[i][1];
    
    if (!cNum || !sNum) continue;
    if (!students[cNum]) students[cNum] = {};

    const engName = studentData[i][2].trim();
    const chiName = studentData[i][3].trim();
    const house = studentData[i][6] ? studentData[i][6].trim() : "";
    const remarks = studentData[i][7] ? studentData[i][7].trim() : "";
    
    let displayName = chiName !== "" ? chiName : engName;
    if (remarks) displayName += ` (${remarks})`;

    students[cNum][sNum] = { name: displayName, house: house };
  }

  return { active: active, history: history, students: students };
}

function borrowItem(studentData) {
  const sheet = getSheet('Active');
  // 直接使用前端傳來的時間
  sheet.appendRow([studentData.classNum, studentData.studentNum, studentData.name, studentData.house, studentData.item, studentData.borrowTime]);
  return true;
}

function returnItem(record) {
  const activeSheet = getSheet('Active');
  const historySheet = getSheet('History');
  const data = activeSheet.getDataRange().getDisplayValues();
  let targetRowIndex = -1;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === record.classNum && data[i][1] === record.studentNum.toString() && data[i][4] === record.item && data[i][5] === record.borrowTime) {
      targetRowIndex = i + 1;
      break;
    }
  }
  
  if (targetRowIndex !== -1) {
    historySheet.appendRow([record.classNum, record.studentNum, record.name, record.house, record.item, record.borrowTime, record.returnTime]);
    activeSheet.deleteRow(targetRowIndex);
    return true;
  } else {
    throw new Error('找不到該筆紀錄');
  }
}
