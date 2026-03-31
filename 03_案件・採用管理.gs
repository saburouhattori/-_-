// =========================================
// 案件管理の操作（登録・更新・削除・採用）
// =========================================

function addJob(formData) {
  const sheet = getMasterSheet('案件管理');
  const lastRow = sheet.getLastRow();
  let nextNumber = 1;
  if (lastRow >= 2) {
    const lastId = String(sheet.getRange(lastRow, 1).getValue());
    const match = lastId.match(/\d+/);
    if (match) nextNumber = parseInt(match[0], 10) + 1;
  }
  const nextId = "JOB-" + nextNumber.toString().padStart(4, '0');

  const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  
  // 複数のURLを単なる改行区切りのテキストにする
  const fileUrls = Array.isArray(formData.relatedFiles) ? formData.relatedFiles.join('\n') : '';

  const row = [
    nextId,
    '未着手',
    todayStr,
    formData.company,
    formData.skill,
    formData.candidates.join('\n'),
    formData.interviewDate || '',
    '', // 採用者氏名
    fileUrls, // ★修正：生のURLをそのまま保存。Googleスプレッドシートが自動でスマートチップ化します
    formData.memo || ''
  ];
  
  sheet.appendRow(row);
  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 3).setNumberFormat('yyyy/MM/dd');

  return `案件登録が完了しました: ${nextId}`;
}

function getJobDetails(jobId) {
  const sheet = getMasterSheet('案件管理');
  const data = sheet.getDataRange().getValues();
  const searchId = String(jobId).trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === searchId) {
      
      let ivDate = data[i][6];
      if (ivDate instanceof Date) {
        ivDate = Utilities.formatDate(ivDate, "JST", "yyyy-MM-dd");
      }

      let rDate = data[i][2];
      if (rDate instanceof Date) {
        rDate = Utilities.formatDate(rDate, "JST", "yyyy-MM-dd");
      }

      return {
        row: i + 1,
        id: data[i][0],
        status: data[i][1],
        date: rDate,
        company: data[i][3],
        skill: data[i][4],
        candidates: data[i][5],
        interviewDate: ivDate,
        hireNames: data[i][7],
        relatedFile: String(data[i][8] || ""), // 生のURLテキストとして取得
        memo: data[i][9]
      };
    }
  }
  return null;
}

function updateJob(formData) {
  const sheet = getMasterSheet('案件管理');
  const row = Number(formData.row);
  
  // 複数のURLを単なる改行区切りのテキストにする
  const fileUrls = Array.isArray(formData.relatedFiles) ? formData.relatedFiles.join('\n') : '';
  
  sheet.getRange(row, 2).setValue(formData.status);
  sheet.getRange(row, 4).setValue(formData.company);
  sheet.getRange(row, 5).setValue(formData.skill);
  sheet.getRange(row, 6).setValue(formData.candidates.join('\n'));
  sheet.getRange(row, 7).setValue(formData.interviewDate);
  sheet.getRange(row, 9).setValue(fileUrls); // ★修正：生のURLをそのまま上書き保存
  sheet.getRange(row, 10).setValue(formData.memo);
  
  return "案件情報を更新しました。";
}

function deleteJobRow(jobId) {
  const sheet = getMasterSheet('案件管理');
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === jobId) {
      sheet.deleteRow(i + 1);
      return "案件を削除しました。";
    }
  }
  return "エラー：案件が見つかりませんでした。";
}

function getJobCandidates(jobId) {
  const details = getJobDetails(jobId);
  if (!details || !details.candidates) return [];
  const candDict = getCandidateDict();
  const ids = details.candidates.split(/\r?\n/);
  return ids.map(id => {
    const cleanId = id.split('-').slice(0,2).join('-').trim();
    return { id: cleanId, display: candDict[cleanId] ? `${cleanId} (${candDict[cleanId]})` : cleanId };
  }).filter(c => c.id);
}

function registerHire(jobId, hiredIds) {
  const sheet = getMasterSheet('案件管理');
  const mSheet = getMasterSheet('登録者マスタ');
  const mCol = getMasterColumnMap(mSheet);
  const data = sheet.getDataRange().getValues();
  const mData = mSheet.getDataRange().getValues();
  
  let companyName = "";
  let targetJobRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === jobId) {
      companyName = data[i][3];
      targetJobRow = i + 1;
      break;
    }
  }
  if (!companyName) return "エラー：案件が見つかりません。";

  const hiredNames = [];
  const candDict = getCandidateDict();
  hiredIds.forEach(id => {
    const name = candDict[id] || id;
    hiredNames.push(name);
    
    // 登録者マスタの更新
    for (let j = 1; j < mData.length; j++) {
      if (String(mData[j][0]) === id) {
        if (mCol['ステータス']) mSheet.getRange(j + 1, mCol['ステータス']).setValue('採用');
        if (mCol['採用事業者']) mSheet.getRange(j + 1, mCol['採用事業者']).setValue(companyName);
        break;
      }
    }
  });

  sheet.getRange(targetJobRow, 8).setValue(hiredNames.join(', '));
  sheet.getRange(targetJobRow, 2).setValue('終了');

  return `${hiredIds.length} 名の採用登録を完了しました。案件ステータスを「終了」にしました。`;
}