/**
 * Times. - 工数管理システム
 * フェーズ1: ログイン機能の実装
 */

let SPREADSHEET_ID: string | null = null;

/**
 * Webアプリケーションのエントリポイント
 * ログイン画面またはホーム画面を表示
 */
function doGet(
  e: GoogleAppsScript.Events.DoGet
): GoogleAppsScript.HTML.HtmlOutput {
  const page = e.parameter.page || "login";

  if (page === "home") {
    return HtmlService.createTemplateFromFile("home")
      .evaluate()
      .setTitle("Times. - ホーム")
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Times. - ログイン")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTMLファイルのインクルード用関数
 * @param filename インクルードするファイル名
 */
function include(filename: string): string {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * スプレッドシートの取得または作成
 * @returns Spreadsheet オブジェクト
 */
function getOrCreateSpreadsheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  const properties = PropertiesService.getScriptProperties();
  let spreadsheetId = properties.getProperty("SPREADSHEET_ID");

  if (spreadsheetId) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      console.error("既存のスプレッドシートが見つかりません:", error);
    }
  }

  // 新しいスプレッドシートを作成
  const spreadsheet = SpreadsheetApp.create("Times. データベース");
  spreadsheetId = spreadsheet.getId();
  properties.setProperty("SPREADSHEET_ID", spreadsheetId);

  // ユーザーシートを作成
  const sheet = spreadsheet.getActiveSheet();
  sheet.setName("users");
  sheet.getRange(1, 1, 1, 3).setValues([["社員番号", "名前", "登録日時"]]);

  // ヘッダー行のスタイルを設定
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setBackground("#4a90e2");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  console.log("新しいスプレッドシートを作成しました: " + spreadsheetId);

  return spreadsheet;
}

/**
 * ユーザー情報
 */
interface User {
  employeeNumber: string;
  name: string;
  registeredAt?: Date;
}

/**
 * ユーザー登録
 * @param employeeNumber 社員番号
 * @param name 名前
 * @returns 処理結果
 */
function registerUser(
  employeeNumber: string,
  name: string
): { success: boolean; message: string } {
  if (!employeeNumber || !name) {
    return { success: false, message: "社員番号と名前を入力してください。" };
  }

  const spreadsheet = getOrCreateSpreadsheet();
  const sheet = spreadsheet.getSheetByName("users");

  if (!sheet) {
    return { success: false, message: "ユーザーシートが見つかりません。" };
  }

  // 既存ユーザーのチェック
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][0].toString() === employeeNumber) {
      return {
        success: false,
        message: "この社員番号は既に登録されています。",
      };
    }
  }

  // 新規ユーザーの追加
  const now = new Date();
  sheet.appendRow([employeeNumber, name, now]);

  return {
    success: true,
    message: `ようこそ、${name}さん！登録が完了しました。`,
  };
}

/**
 * ユーザーログイン
 * @param employeeNumber 社員番号
 * @returns ログイン結果とユーザー情報
 */
function loginUser(employeeNumber: string): {
  success: boolean;
  message: string;
  user?: User;
} {
  // try {
  console.log("loginUser 開始:", employeeNumber);

  if (!employeeNumber) {
    console.log("社員番号が空です");
    return { success: false, message: "社員番号を入力してください。" };
  }

  console.log("スプレッドシート取得開始");
  const spreadsheet = getOrCreateSpreadsheet();
  console.log("スプレッドシート取得完了:", spreadsheet.getId());

  const sheet = spreadsheet.getSheetByName("users");
  console.log("シート取得結果:", sheet ? "OK" : "NG");

  if (!sheet) {
    console.log("ユーザーシートが見つかりません");
    return { success: false, message: "ユーザーシートが見つかりません。" };
  }

  // ユーザーの検索
  console.log("データ範囲取得開始");
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  console.log("データ行数:", values.length);
  console.log("全データ:", values);

  for (let i = 1; i < values.length; i++) {
    console.log(`行${i}のデータ:`, values[i]);
    console.log(`比較: "${values[i][0].toString()}" === "${employeeNumber}"`);

    if (values[i][0].toString() === employeeNumber) {
      console.log("ユーザーが見つかりました");
      const user: User = {
        employeeNumber: values[i][0].toString(),
        name: values[i][1].toString(),
      };

      const result = {
        success: true,
        message: `ログインしました。おかえりなさい、${user.name}さん！`,
        user: user,
      };
      console.log("ログイン成功結果:", result);
      return result;
    }
  }

  console.log("ユーザーが見つかりませんでした");
  return {
    success: false,
    message: "社員番号が見つかりません。先に登録してください。",
  };
  // } catch (error) {
  //   console.error('loginUser でエラーが発生:', error);
  //   return {
  //     success: false,
  //     message: 'システムエラーが発生しました: ' + String(error)
  //   };
  // }
}

/**
 * スプレッドシート情報の取得
 * @returns スプレッドシートのURL
 */
function getSpreadsheetInfo(): { url: string | null; id: string | null } {
  try {
    const spreadsheet = getOrCreateSpreadsheet();
    return {
      url: spreadsheet.getUrl(),
      id: spreadsheet.getId(),
    };
  } catch (error) {
    console.error("スプレッドシート情報の取得に失敗:", error);
    return { url: null, id: null };
  }
}

/**
 * 初期設定の確認
 * @returns セットアップ状態
 */
function checkSetup(): { isReady: boolean; message: string } {
  try {
    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = spreadsheet.getSheetByName("users");

    if (sheet) {
      return {
        isReady: true,
        message: "システムの準備ができています。",
      };
    }

    return {
      isReady: false,
      message: "システムの初期設定が必要です。",
    };
  } catch (error) {
    return {
      isReady: false,
      message: "エラーが発生しました: " + error,
    };
  }
}
