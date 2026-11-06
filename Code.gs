// =========================================================================================
// [Google Apps Script 최종 완성 코드 - 시트 자동 감지 및 실시간 반영]
// 1. 구글 시트 상단 메뉴: [확장 프로그램] -> [Apps Script] 클릭
// 2. 기존 코드를 모두 지우고 아래 코드를 그대로 붙여넣기 (Ctrl+A -> Ctrl+V)
// 3. 우측 상단 [배포] -> [배포 관리] 클릭
// 4. 연필 아이콘(수정) 클릭 -> 버전 드롭다운에서 [새 버전] 선택 -> [배포] 클릭!
// =========================================================================================

function getTargetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 1순위: '과제_대시보드' 시트가 있으면 사용
  let sheet = ss.getSheetByName("과제_대시보드");
  // 2순위: 없으면 현재 데이터가 있는 첫 번째 시트를 자동으로 사용 (시트 이름 무관)
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
      saenggi: String(r[2] || 'N'),
      bojeon: String(r[3] || 'N'),
      anjeon: String(r[4] || 'N'),
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

    // 1) 승인 상태 토글 처리 (생기/보전/안전 버튼 클릭)
    if (contents.action === "updateDept") {
      const colMap = { 'saenggi': 3, 'bojeon': 4, 'anjeon': 5 };
      const col = colMap[contents.dept];
      const row = Number(contents.rowIndex);
      
      if (col && row) {
        sheet.getRange(row, col).setValue(contents.status);
        SpreadsheetApp.flush(); // 즉시 구글 시트 DB에 기록 강제
        return ContentService.createTextOutput(JSON.stringify({ 
          success: true, 
          message: "업데이트 성공", 
          row: row, 
          col: col, 
          status: contents.status 
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 2) 신규 과제 등록 처리 (새 행 추가)
    if (contents.action === "create") {
      sheet.appendRow([
        contents.problem || "",
        contents.solution || "",
        contents.saenggi || "N",
        contents.bojeon || "N",
        contents.anjeon || "N",
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
