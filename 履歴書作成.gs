/**
 * 登録者IDをもとに履歴書を作成する（自動追従版）
 */
function rirekisyo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = getMasterSheet("登録者マスタ");
  const sheet = ss.getSheetByName("履歴書");
  
  // 1. 履歴書シートのB2セルからIDを取得
  const adminId = sheet.getRange("B2").getValue();
  if (!adminId) {
    Browser.msgBox("B2セルに登録者ID（SD-xxxx）を入力してください。");
    return;
  }

  // 2. マスタの列番号を名前で取得（ズレ防止）
  const col = getMasterColumnMap(masterSheet);
  const dataRows = masterSheet.getDataRange().getValues();
  let rowData = null;
  let targetRowIndex = -1;

  // 3. IDに一致する行を検索
  for (let i = 1; i < dataRows.length; i++) {
    if (String(dataRows[i][0]).trim().toUpperCase() === String(adminId).trim().toUpperCase()) {
      rowData = dataRows[i];
      targetRowIndex = i + 1;
      break;
    }
  }

  if (!rowData) {
    Browser.msgBox("指定されたID（" + adminId + "）が「登録者マスタ」に見つかりません。");
    return;
  }

  // 補助関数：項目名から値を取得する
  const getVal = (name) => {
    const cIdx = col[name.replace(/\s/g, '')];
    if (!cIdx) return "";
    const val = rowData[cIdx - 1];
    // 日付型の場合はフォーマット
    if (val instanceof Date) return Utilities.formatDate(val, "JST", "yyyy年M月d日");
    return val || "";
  };

  // ---------------------------------------------------------
  // 4. 基本情報の書き込み
  // ---------------------------------------------------------
  sheet.getRange("C4").setValue(getVal('フリガナ'));
  sheet.getRange("C5").setValue(getVal('呼び名'));
  sheet.getRange("C6").setValue(getVal('名前')); // 履歴書上のラベルは氏名
  
  sheet.getRange("C8").setValue(getVal('生年月日'));
  sheet.getRange("G8").setValue(getVal('満年齢') + "歳");
  sheet.getRange("C9").setValue(" " + getVal('性別') + " "); // スペースで調整
  sheet.getRange("K8").setValue(" " + getVal('配偶者') + " ");
  sheet.getRange("K9").setValue(getVal('身長') + "cm");
  sheet.getRange("K10").setValue(getVal('体重') + "kg");

  sheet.getRange("C12").setValue(getVal('現住所'));
  sheet.getRange("C13").setValue(getVal('住所（出身地）'));
  sheet.getRange("K12").setValue(getVal('メールアドレス'));

  // ---------------------------------------------------------
  // 5. 写真の取得（登録者マスタのC列/顔写真列から直接取得）
  // ---------------------------------------------------------
  const photoColIdx = col['顔写真'];
  if (photoColIdx) {
    const photoImage = masterSheet.getRange(targetRowIndex, photoColIdx).getValue();
    if (photoImage) {
      sheet.getRange("D4").setValue(photoImage);
    } else {
      sheet.getRange("D4").clearContent();
    }
  }

  // ---------------------------------------------------------
  // 6. 学歴・職歴の書き込み
  // ---------------------------------------------------------
  // いったんクリア
  sheet.getRange("B17:T20").clearContent();
  sheet.getRange("B22:T24").clearContent();

  // 学歴
  sheet.getRange("B17").setValue(getVal('学歴＞入学年月'));
  sheet.getRange("G17").setValue(getVal('学歴＞学校名') + "　入学");
  sheet.getRange("B18").setValue(getVal('学歴＞卒業/中退年月'));
  sheet.getRange("G18").setValue(" " + getVal('学歴＞状況') + " ");
  sheet.getRange("G19").setValue(getVal('学歴＞補足'));

  // 職歴（3つ分）
  if (getVal('職歴①＞期間')) {
    sheet.getRange("B22").setValue(getVal('職歴①＞期間'));
    sheet.getRange("G22").setValue(getVal('職歴①＞内容'));
  }
  if (getVal('職歴②＞期間')) {
    sheet.getRange("B23").setValue(getVal('職歴②＞期間'));
    sheet.getRange("G23").setValue(getVal('職歴②＞内容'));
  }
  if (getVal('職歴③＞期間')) {
    sheet.getRange("B24").setValue(getVal('職歴③＞期間'));
    sheet.getRange("G24").setValue(getVal('職歴③＞内容'));
  }

  // ---------------------------------------------------------
  // 7. 日本語能力・資格
  // ---------------------------------------------------------
  // JLPT
  let jlpt = getVal('特定技能要件＞JLPTレベル') ? getVal('特定技能要件＞JLPTレベル') + "合格" : "-";
  if (getVal('特定技能要件＞JLPT取得年月')) jlpt += "（" + getVal('特定技能要件＞JLPT取得年月') + "）";
  sheet.getRange("D29").setValue(jlpt);

  // JFT
  let jft = getVal('特定技能要件＞JFTBasicレベル') ? getVal('特定技能要件＞JFTBasicレベル') + "合格" : "-";
  if (getVal('特定技能要件＞JFT取得年月')) jft += "（" + getVal('特定技能要件＞JFT取得年月') + "）";
  sheet.getRange("D30").setValue(jft);

  // 介護技能
  let skill = getVal('特定技能要件＞介護技能評価試験') || "-";
  if (getVal('特定技能要件＞介護技能取得年月')) skill += " （" + getVal('特定技能要件＞介護技能取得年月') + "）";
  sheet.getRange("L29").setValue(skill);

  // 介護日本語
  let lang = getVal('特定技能要件＞介護日本語評価試験') || "-";
  if (getVal('特定技能要件＞介護日本語取得年月')) lang += " （" + getVal('特定技能要件＞介護日本語取得年月') + "）";
  sheet.getRange("L30").setValue(lang);

  // その他試験
  sheet.getRange("D31").setValue(getVal('その他の日本語能力試験') || "-");
  sheet.getRange("H31").setValue(getVal('取得年月') ? "（" + getVal('取得年月') + "）" : "");

  // ---------------------------------------------------------
  // 8. その他
  // ---------------------------------------------------------
  sheet.getRange("B34").setValue(getVal('日本在住の親族について'));
  sheet.getRange("B38").setValue(getVal('コメント'));

  // 完了メッセージ（右下に日付を入れる）
  const today = Utilities.formatDate(new Date(), "JST", "yyyy年M月d日");
  sheet.getRange("O2").setValue(today);

  Browser.msgBox(getVal('名前') + " さんの履歴書を作成しました。");
}