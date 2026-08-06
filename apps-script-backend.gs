/**
 * ドリンクバー台帳 - データ保存用バックエンド
 *
 * 使い方:
 * 1. https://sheets.google.com で新しいスプレッドシートを作成する
 * 2. メニューの「拡張機能」→「Apps Script」を開く
 * 3. デフォルトで入っているコードを全部消して、このファイルの内容を貼り付ける
 * 4. 下の SECRET_TOKEN を、自分だけが知っている好きな文字列に変更する
 *    (例: "kaya-drinkbar-2026" のような、他人に推測されにくいもの)
 * 5. 上部の「デプロイ」→「新しいデプロイ」をクリック
 * 6. 種類の選択で「ウェブアプリ」を選ぶ
 * 7. 「アクセスできるユーザー」を必ず「全員」にする(これをしないとアプリから呼べません)
 * 8. 「デプロイ」をクリックし、表示された「ウェブアプリURL」をコピーしておく
 *    → このURLと、4で決めたSECRET_TOKENを、HTMLファイルの設定欄に入力します
 */

// ここを自分だけの合言葉に変更してください
const SECRET_TOKEN = "CHANGE_ME_TO_YOUR_OWN_SECRET";

const SHEET_NAME = "drinkbar_data";

function doGet(e) {
  try {
    if (!e || e.parameter.token !== SECRET_TOKEN) {
      return jsonResponse({ ok: false, error: "unauthorized" });
    }
    const sheet = getSheet();
    const value = sheet.getRange("A1").getValue();
    return jsonResponse({ ok: true, value: value ? value : null });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    if (params.token !== SECRET_TOKEN) {
      return jsonResponse({ ok: false, error: "unauthorized" });
    }
    const sheet = getSheet();
    sheet.getRange("A1").setValue(params.value);
    sheet.getRange("B1").setValue(new Date());
    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange("A1").setValue("");
    sheet.getRange("C1").setValue("← A列に台帳の全データ(JSON形式)が保存されます。B列は最終更新日時です。");
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
