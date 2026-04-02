const SHEET_ID = "1lTpW--FU6z3_4lTZkEUs9IF5Y8zOkLfLphcacPnax8A";
const FOLDER_ID = "16s8rh0qPPDcGctOzgURWi9DHnoACzUgS";
const ADMIN_PW = "alsk0118**";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    let result;
    if (data.action === "checkAdmin") {
      result = { success: data.pw === ADMIN_PW };
    } else if (data.action === "submit") {
      result = handleSubmit(data);
    } else if (data.action === "getAll") {
      result = handleGetAll();
    } else if (data.action === "updateStatus") {
      result = handleUpdateStatus(data);
    } else if (data.action === "deleteRow") {
      result = handleDeleteRow(data);
    } else if (data.action === "getMembers") {
      result = handleGetMembers();
    } else if (data.action === "saveMembers") {
      result = handleSaveMembers(data);
    } else if (data.action === "clearAllData") {
      result = handleClearAllData();
    } else if (data.action === "getStorageInfo") {
      result = handleGetStorageInfo();
    } else if (data.action === "saveCategories") {
      result = handleSaveCategories(data);
    } else if (data.action === "getCategories") {
      result = handleGetCategories();
    } else if (data.action === "saveDeptCode") {
      result = data.pw === ADMIN_PW ? handleSaveDeptCode(data) : { success: false, error: "권한 없음" };
    } else if (data.action === "getDeptCodes") {
      result = data.pw === ADMIN_PW ? handleGetDeptCodes() : { success: false, error: "권한 없음" };
    } else if (data.action === "deleteDeptCode") {
      result = data.pw === ADMIN_PW ? handleDeleteDeptCode(data) : { success: false, error: "권한 없음" };
    } else if (data.action === "verifyDeptCode") {
      result = handleVerifyDeptCode(data);
    } else {
      result = { success: false, error: "알 수 없는 액션" };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false, error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === "getAll") {
    return ContentService.createTextOutput(JSON.stringify(handleGetAll()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "getMembers") {
    return ContentService.createTextOutput(JSON.stringify(handleGetMembers()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "getCategories") {
    return ContentService.createTextOutput(JSON.stringify(handleGetCategories()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({
    success: true, message: "서버 정상 작동 중"
  })).setMimeType(ContentService.MimeType.JSON);
}

function handleGetStorageInfo() {
  try {
    const quota = DriveApp.getStorageLimit();
    const used = DriveApp.getStorageUsed();
    return { success: true, total: quota, used: used };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function handleSubmit(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("신청내역") || ss.insertSheet("신청내역");

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID","부서","사번","이름","분야","활동항목","건수","마일리지","상태","메모","증빙파일명","증빙파일URL","신청일시","반려사유"]);
    sheet.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#0f2557").setFontColor("white");
  }

  let fileUrl = "";
  let fileName = "";
  if (data.fileData && data.fileName) {
    // 용량 여유 확인 (50MB 미만이면 경고)
    try {
      const remaining = DriveApp.getStorageLimit() - DriveApp.getStorageUsed();
      if (remaining < 50 * 1024 * 1024) {
        return { success: false, error: "STORAGE_FULL" };
      }
    } catch(e) {}
    try {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const base64 = data.fileData.indexOf(",") > -1 ? data.fileData.split(",")[1] : data.fileData;
      const mimeType = data.fileData.indexOf(";") > -1 ? data.fileData.split(";")[0].split(":")[1] : "application/octet-stream";
      const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, data.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
      fileName = data.fileName;
    } catch(err) {
      fileName = data.fileName || "";
    }
  }

  const id = Utilities.getUuid();
  sheet.appendRow([
    id,
    data.dept, data.empNo, data.name,
    data.categoryName, data.itemName,
    data.count, data.points,
    "대기", data.memo || "",
    fileName, fileUrl,
    new Date().toLocaleString("ko-KR"),
    ""
  ]);

  return { success: true, id: id };
}

function handleGetAll() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("신청내역");
  if (!sheet || sheet.getLastRow() <= 1) {
    return { success: true, data: [] };
  }
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const data = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return { success: true, data: data };
}

function handleUpdateStatus(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("신청내역");
  if (!sheet) return { success: false, error: "시트 없음" };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.getRange(i + 1, 9).setValue(data.status);
      sheet.getRange(i + 1, 14).setValue(data.rejectReason || "");
      const color = data.status === "승인" ? "#f0fdf4" : data.status === "반려" ? "#fef2f2" : "white";
      sheet.getRange(i + 1, 1, 1, 14).setBackground(color);
      break;
    }
  }
  return { success: true };
}

function handleDeleteRow(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("신청내역");
  if (!sheet) return { success: false, error: "시트 없음" };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: "해당 ID를 찾을 수 없습니다." };
}

function handleGetMembers() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("명단") || ss.insertSheet("명단");
  const data = sheet.getDataRange().getValues();
  const members = data.slice(1)
    .filter(r => r[0] && r[1] && r[2])
    .map(r => ({ dept: r[0], empNo: String(r[1]), name: r[2] }));
  return { members: members };
}

function handleSaveMembers(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("명단") || ss.insertSheet("명단");
  const members = JSON.parse(data.members);
  sheet.clearContents();
  sheet.appendRow(["부서명", "사번", "이름"]);
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#0f2557").setFontColor("white");
  members.forEach(m => sheet.appendRow([m.dept, m.empNo, m.name]));
  return { ok: true };
}

function testWrite() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("신청내역");
  Logger.log(sheet ? "시트 찾음: " + sheet.getLastRow() + "행" : "시트 없음!");
}

function handleClearAllData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("신청내역");
    if (!sheet) return { success: true };
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return { success: true };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function handleSaveCategories(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName("운영항목");
    if (!sheet) {
      sheet = ss.insertSheet("운영항목");
    }
    sheet.clearContents();
    sheet.appendRow(["JSON"]);
    sheet.getRange(1, 1).setFontWeight("bold").setBackground("#0f2557").setFontColor("white");
    sheet.getRange(2, 1).setValue(data.data);
    return { success: true };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function handleGetCategories() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("운영항목");
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, data: null }; // 없으면 기본값 사용
    }
    const json = sheet.getRange(2, 1).getValue();
    return { success: true, data: json };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

// ── 부서코드 관리 ──────────────────────────────────
function getDeptCodeSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("부서코드");
  if (!sheet) {
    sheet = ss.insertSheet("부서코드");
    sheet.appendRow(["id", "dept", "code", "createdAt"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#0f2557").setFontColor("white");
  }
  return sheet;
}

function handleSaveDeptCode(data) {
  try {
    const sheet = getDeptCodeSheet();
    const id = Utilities.getUuid().replace(/-/g, "").substring(0, 12);
    const code = data.dept.substring(0, 2) + "-" + Utilities.getUuid().replace(/-/g, "").substring(0, 6);
    sheet.appendRow([id, data.dept, code, new Date().toLocaleString("ko-KR")]);
    return { success: true, id: id, code: code };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function handleGetDeptCodes() {
  try {
    const sheet = getDeptCodeSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, codes: [] };
    const codes = data.slice(1).filter(r => r[0]).map(r => ({
      id: r[0], dept: r[1], code: r[2], createdAt: r[3]
    }));
    return { success: true, codes: codes };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function handleDeleteDeptCode(data) {
  try {
    const sheet = getDeptCodeSheet();
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: "코드를 찾을 수 없습니다" };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

function handleVerifyDeptCode(data) {
  try {
    const sheet = getDeptCodeSheet();
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][2]).trim().toLowerCase() === String(data.code).trim().toLowerCase()) {
        return { success: true, dept: rows[i][1] };
      }
    }
    return { success: false, error: "유효하지 않은 코드입니다" };
  } catch(err) {
    return { success: false, error: err.message };
  }
}
