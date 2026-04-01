/**
 * 案件登録（新規事業者の自動マスタ登録付き）
 */
function addJob(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('案件管理');
  
  // --- 新規事業者の自動マスタ登録 ---
  const companyName = String(formData.company || "").trim();
  let isNewCompany = false;

  if (companyName) {
    const compSheet = ss.getSheetByName('事業者マスタ');
    if (compSheet) {
      const compData = compSheet.getDataRange().getValues();
      const exists = compData.some(row => String(row[1]).trim() === companyName);
      
      if (!exists) {
        isNewCompany = true;
        let lastIdNum = 0;
        for (let i = 1; i < compData.length; i++) {
          let idVal = String(compData[i][0]);
          let match = idVal.match(/\d+/);
          if (match) {
            let num = parseInt(match[0], 10);
            if (num > lastIdNum) lastIdNum = num;
          }
        }
        const nextCompId = "CO-" + (lastIdNum + 1).toString().padStart(4, '0');
        compSheet.appendRow([nextCompId, companyName, "", "", "", "案件登録により自動追加"]);
      }
    }
  }

  const idVals = sheet.getRange("A1:A" + sheet.getMaxRows()).getValues();
  let lastDataRow = 1;
  let lastIdNum = 0;

  for (let i = 1; i < idVals.length; i++) {
    let idVal = String(idVals[i][0]).trim();
    if (idVal !== "") {
      lastDataRow = i + 1;
      let match = idVal.match(/\d+/);
      if (match) {
        let num = parseInt(match[0], 10);
        if (num > lastIdNum) lastIdNum = num;
      }
    }
  }

  sheet.insertRowAfter(lastDataRow);
  const targetRow = lastDataRow + 1;

  const nextId = "JOB-" + (lastIdNum + 1).toString().padStart(4, '0');
  const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  const fileUrls = Array.isArray(formData.relatedFiles) ? formData.relatedFiles.join('\n') : '';

  const rowData = [
    nextId, '未着手', todayStr, companyName, formData.skill, formData.candidates.join('\n'), 
    formData.interviewDate || '', '', '', formData.memo || ''
  ];
  
  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  convertToSmartChips(sheet, targetRow, 9, fileUrls);
  sheet.getRange(targetRow, 3).setNumberFormat('yyyy/MM/dd');

  // オブジェクトで詳細な情報を返す
  return {
    success: true,
    message: `案件登録が完了しました: ${nextId}`,
    isNewCompany: isNewCompany,
    companyName: companyName
  };
}

// 他の関数は変更なしのため、既存の convertToSmartChips, getJobDetails, updateJob などをそのまま含めてください。