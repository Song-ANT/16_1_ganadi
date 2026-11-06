// =========================================================================================
// [Google Apps Script 최종 완성 코드 - 3대 승인 게이트 (Y, -, N, X) 및 롤백 동기화]
// =========================================================================================

function getTargetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("과제_대시보드");
  if (!sheet) {
    sheet = ss.getSheets()[0];
  }
  return sheet;
}

function doGet(e) {
  const sheet = getTargetSheet();
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0] && !r[1]) continue;
    rows.push({
      rowIndex: i + 1,
      problem: String(r[0] || ''),
      solution: String(r[1] || ''),
      saenggi: String(r[2] || '-'),
      bojeon: String(r[3] || '-'),
      anjeon: String(r[4] || '-'),
      typingCount: Number(r[5]) || 1,
      rollbackCount: Number(r[6]) || 0
    });
  }

  return ContentService.createTextOutput(JSON.stringify({ success: true, rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const sheet = getTargetSheet();

    // 1) 승인 상태 업데이트 처리 (Y, -, N, X 및 롤백 카운트)
    if (contents.action === "updateDept") {
      const colMap = { 'saenggi': 3, 'bojeon': 4, 'anjeon': 5 };
      const col = colMap[contents.dept];
      let row = Number(contents.rowIndex);

      // 문제점 텍스트로 실제 행을 정확하게 자동 탐색
      if (contents.problem) {
        const allData = sheet.getDataRange().getValues();
        for (let i = 0; i < allData.length; i++) {
          if (String(allData[i][0]).trim() === String(contents.problem).trim()) {
            row = i + 1;
            break;
          }
        }
      }
      
      if (col && row) {
        sheet.getRange(row, col).setValue(contents.status);
        if (contents.rollbackCount !== undefined && contents.rollbackCount !== null) {
          sheet.getRange(row, 7).setValue(Number(contents.rollbackCount));
        }
        SpreadsheetApp.flush(); // 즉시 구글 시트 DB에 기록 강제
        return ContentService.createTextOutput(JSON.stringify({ 
          success: true, 
          message: "업데이트 성공", 
          row: row, 
          col: col, 
          status: contents.status,
          rollbackCount: contents.rollbackCount
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 2) 신규 과제 등록 처리 (새 행 추가)
    if (contents.action === "create") {
      sheet.appendRow([
        contents.problem || "",
        contents.solution || "",
        contents.saenggi || "-",
        contents.bojeon || "-",
        contents.anjeon || "-",
        1,
        0
      ]);
      SpreadsheetApp.flush(); // 즉시 구글 시트 DB에 기록 강제
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        message: "등록 성공" 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "알 수 없는 액션" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
