/**
 * .Times - 工数管理システム
 * フェーズ1: ログイン機能の実装
 * フェーズ2: 勤怠管理機能の実装
 */

// ユーザー管理用スプレッドシートID
let USER_SPREADSHEET_ID: string | null = null;

// 勤怠管理用スプレッドシートIDをプロパティストアから取得
const ATTENDANCE_SPREADSHEET_ID =
  PropertiesService.getScriptProperties().getProperty(
    "CALENDAR_SPREADSHEET_ID"
  ) || "";

// 案件管理用スプレッドシートIDをプロパティストアから取得
const PROJECTS_SPREADSHEET_ID =
  PropertiesService.getScriptProperties().getProperty(
    "PROJECTS_SPREADSHEET_ID"
  ) || "";

/**
 * Webアプリケーションのエントリポイント
 * シングルページアプリケーションのHTML
 */
function doGet(): GoogleAppsScript.HTML.HtmlOutput {
  return HtmlService.createTemplateFromFile("app")
    .evaluate()
    .setTitle(".Times")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setFaviconUrl(
      "https://cdn-0.emojis.wiki/emoji-pics/facebook/alarm-clock-facebook.png"
    )
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
  let spreadsheetId = properties.getProperty("USER_SPREADSHEET_ID");

  if (spreadsheetId) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      console.error("既存のスプレッドシートが見つかりません:", error);
    }
  }

  // 新しいスプレッドシートを作成
  const spreadsheet = SpreadsheetApp.create(".Times ユーザーデータベース");
  spreadsheetId = spreadsheet.getId();
  properties.setProperty("USER_SPREADSHEET_ID", spreadsheetId);

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
 * 全ユーザー情報を取得
 * @returns ユーザー一覧
 */
function getAllUsers(): {
  success: boolean;
  message: string;
  data?: Array<{
    employeeNumber: string;
    name: string;
  }>;
} {
  try {
    console.log("getAllUsers開始");

    const spreadsheet = getOrCreateSpreadsheet();
    if (!spreadsheet) {
      console.error("スプレッドシートの取得に失敗しました");
      return {
        success: false,
        message: "スプレッドシートにアクセスできません。",
      };
    }

    const sheet = spreadsheet.getSheetByName("users");
    if (!sheet) {
      console.log("ユーザーシートが見つかりません");
      return {
        success: false,
        message: "ユーザーシートが見つかりません。",
      };
    }

    const dataRange = sheet.getDataRange();
    if (!dataRange) {
      console.error("データ範囲の取得に失敗しました");
      return {
        success: false,
        message: "データ範囲を取得できません。",
      };
    }

    const values = dataRange.getValues();
    if (!values || values.length <= 1) {
      console.log("ユーザーデータが存在しません");
      return {
        success: true,
        message: "登録されているユーザーがいません",
        data: [],
      };
    }

    const users: Array<{
      employeeNumber: string;
      name: string;
      overtime: string;
      updated_at: string;
    }> = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      users.push({
        employeeNumber: row[0] ? row[0].toString() : "",
        name: row[1] ? row[1].toString() : "",
        overtime: row[3] ? row[3].toString() : "0",
        updated_at: row[4] ? row[4].toString() : "",
      });
    }

    console.log(`${users.length}件のユーザーを取得`);

    return {
      success: true,
      message: "ユーザー一覧を取得しました",
      data: users,
    };
  } catch (error) {
    console.error("ユーザー一覧取得エラー:", error);
    return {
      success: false,
      message: "ユーザー一覧の取得に失敗しました: " + String(error),
    };
  }
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
 * 社員番号からユーザー名を取得
 * @param employeeNumber 社員番号
 * @returns ユーザー名（見つからない場合は社員番号をそのまま返す）
 */
function getUserNameByEmployeeNumber(employeeNumber: string): string {
  try {
    if (!employeeNumber) {
      return "";
    }

    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = spreadsheet.getSheetByName("users");

    if (!sheet) {
      return employeeNumber;
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    for (let i = 1; i < values.length; i++) {
      if (values[i][0].toString() === employeeNumber) {
        return values[i][1].toString();
      }
    }

    return employeeNumber;
  } catch (error) {
    console.error("ユーザー名取得エラー:", error);
    return employeeNumber;
  }
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

/**
 * 勤怠スプレッドシートの取得
 */
function getAttendanceSpreadsheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  try {
    const spreadsheet = SpreadsheetApp.openById(ATTENDANCE_SPREADSHEET_ID);
    console.log("勤怠スプレッドシート取得成功:", spreadsheet.getName());
    return spreadsheet;
  } catch (error) {
    console.error("勤怠スプレッドシートの取得に失敗:", error);
    console.error("スプレッドシートID:", ATTENDANCE_SPREADSHEET_ID);
    throw new Error(
      "勤怠スプレッドシートにアクセスできません: " + String(error)
    );
  }
}

/**
 * 現在の年月シート名を取得（YYYYMM形式）
 */
function getCurrentSheetName(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  return `${year}${month}`;
}

/**
 * 勤怠シートの取得（作成はしない、既存のみ）
 */
function getAttendanceSheet(
  sheetName: string
): GoogleAppsScript.Spreadsheet.Sheet | null {
  try {
    const spreadsheet = getAttendanceSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (sheet) {
      console.log(`シート "${sheetName}" が見つかりました`);
      const lastRow = sheet.getLastRow();
      console.log(`シート "${sheetName}" の最終行: ${lastRow}`);
    } else {
      console.log(`シート "${sheetName}" が見つかりません`);
    }

    return sheet;
  } catch (error) {
    console.error(`シート "${sheetName}" の取得に失敗:`, error);
    return null;
  }
}

/**
 * 勤怠シートの取得または作成
 */
function getOrCreateAttendanceSheet(
  sheetName: string
): GoogleAppsScript.Spreadsheet.Sheet {
  let sheet = getAttendanceSheet(sheetName);

  if (!sheet) {
    const spreadsheet = getAttendanceSpreadsheet();
    sheet = spreadsheet.insertSheet(sheetName);
    // ヘッダーを設定
    sheet
      .getRange(1, 1, 1, 5)
      .setValues([
        ["Date", "EmployeeNumber", "Action", "Timestamp", "Details"],
      ]);

    // ヘッダースタイルの設定
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setBackground("#4a90e2");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");

    console.log(`新しいシート "${sheetName}" を作成しました`);
  }

  return sheet;
}

/**
 * 打刻アクション
 */
function stampAction(
  employeeNumber: string,
  action: string,
  details: string = "",
  specifiedDate?: string
): { success: boolean; message: string } {
  try {
    const sheetName = getCurrentSheetName();
    const sheet = getOrCreateAttendanceSheet(sheetName);

    const now = new Date();
    const date =
      specifiedDate || Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM-dd");
    const timestamp = Utilities.formatDate(
      now,
      "Asia/Tokyo",
      "yyyy/MM/dd HH:mm:ss"
    );

    // データを追加
    const lastRow = sheet.getLastRow();
    sheet
      .getRange(lastRow + 1, 1, 1, 5)
      .setValues([[date, employeeNumber, action, timestamp, details]]);

    const actionMessages: { [key: string]: string } = {
      clockIn: "出勤を記録しました",
      clockOut: "退勤を記録しました",
      breakStart: "中抜け開始を記録しました",
      breakEnd: "中抜け終了を記録しました",
      halfDay: "半休を登録しました",
      fullDay: "全休を登録しました",
      holidayWork: "休日出勤を登録しました",
    };

    if (action === "clockOut") {
      try {
        console.log(
          `退勤処理完了。社員番号 ${employeeNumber} の過不足時間を更新します`
        );
        const updateResult = updateUserSurplusDeficitOnClockOut(employeeNumber);
        if (updateResult.success) {
          console.log("過不足時間更新成功:", updateResult.message);
        } else {
          console.warn("過不足時間更新失敗:", updateResult.message);
        }
      } catch (updateError) {
        console.error("過不足時間更新エラー:", updateError);
      }
    }

    return {
      success: true,
      message: actionMessages[action] || "打刻を記録しました",
    };
  } catch (error) {
    console.error("打刻エラー:", error);
    return {
      success: false,
      message: "打刻に失敗しました: " + error,
    };
  }
}

/**
 * 指定した年月の勤怠レコードを取得
 */
function getAttendanceRecords(
  employeeNumber: string,
  year: number,
  month: number
): any[] {
  try {
    const sheetName = `${year}${String(month).padStart(2, "0")}`;
    const sheet = getAttendanceSheet(sheetName);

    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const records = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === employeeNumber) {
        records.push({
          rowNumber: i + 1,
          date: data[i][0]
            ? Utilities.formatDate(
                new Date(data[i][0]),
                "Asia/Tokyo",
                "yyyy-MM-dd"
              )
            : "",
          employeeNumber: data[i][1].toString(),
          action: data[i][2].toString(),
          timestamp: data[i][3]
            ? Utilities.formatDate(
                new Date(data[i][3]),
                "Asia/Tokyo",
                "yyyy-MM-dd HH:mm:ss"
              )
            : "",
          details: data[i][4] ? data[i][4].toString() : "",
        });
      }
    }

    return records;
  } catch (error) {
    console.error("勤怠レコード取得エラー:", error);
    return [];
  }
}

/**
 * オープンなレコード（未完了の出勤/中抜け）を取得
 */
function getOpenRecord(employeeNumber: string): any | null {
  try {
    let openRecord = null;
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;

    // 現在の月のレコードを取得し、時刻順にソート
    const records = getAttendanceRecords(employeeNumber, year, month);
    records.sort(
      (a, b) =>
        new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime()
    );

    records.forEach((rec) => {
      if (rec.action === "clockIn" || rec.action === "breakStart") {
        openRecord = rec;
      } else if (
        rec.action === "clockOut" ||
        rec.action === "breakEnd" ||
        rec.action === "holidayWork" ||
        rec.action === "fullDay" ||
        rec.action === "halfDay"
      ) {
        openRecord = null;
      }
    });

    // もしオープンなレコードが見つからなければ、前月もチェック
    if (!openRecord) {
      let prevMonth = month - 1;
      let prevYear = year;
      if (prevMonth < 1) {
        prevMonth = 12;
        prevYear--;
      }

      const prevRecords = getAttendanceRecords(
        employeeNumber,
        prevYear,
        prevMonth
      );
      prevRecords.sort(
        (a, b) =>
          new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime()
      );

      prevRecords.forEach((rec) => {
        if (rec.action === "clockIn" || rec.action === "breakStart") {
          openRecord = rec;
        } else if (rec.action === "clockOut" || rec.action === "breakEnd") {
          openRecord = null;
        }
      });
    }

    console.log(
      `オープンレコード検索結果 - 社員番号: ${employeeNumber}, オープンレコード:`,
      openRecord
    );
    return openRecord;
  } catch (error) {
    console.error("オープンレコード取得エラー:", error);
    return null;
  }
}

/**
 * 現在の勤怠状態を取得（Sample codeベース）
 */
function getCurrentStatus(employeeNumber: string): { status: string } {
  try {
    console.log(`getCurrentStatus開始 - 社員番号: ${employeeNumber}`);

    const openRecord = getOpenRecord(employeeNumber);
    let status = "未出勤";

    if (openRecord) {
      if (openRecord.action === "clockIn") {
        status = "出勤中";
      } else if (openRecord.action === "breakStart") {
        status = "中抜け中";
      }
    } else {
      // オープンなレコードがなければ、もし直近のレコードがあるなら状態を判定
      const now = new Date();
      const year = now.getFullYear();
      const month = now.getMonth() + 1;
      const records = getAttendanceRecords(employeeNumber, year, month);

      if (records.length > 0) {
        const latest = records[records.length - 1];
        const today = new Date().toISOString().split("T")[0];
        console.log(`最新レコード:`, latest);
        console.log(`今日の日付: ${today}, 最新レコードの日付: ${latest.date}`);

        // 今日のレコードのみを考慮
        if (latest.date === today) {
          if (latest.action === "clockOut") {
            status = "退勤済み";
          } else if (latest.action === "breakEnd") {
            status = "出勤中";
          } else if (latest.action === "fullDay") {
            status = "全休";
          } else if (latest.action === "halfDay") {
            status = "半休";
          } else if (latest.action === "holidayWork") {
            status = "休日出勤";
          }
        } else {
          // 今日のレコードがない場合は未出勤
          status = "未出勤";
        }
      }
    }

    console.log(`最終状態判定: status=${status}, openRecord=`, openRecord);
    return { status };
  } catch (error) {
    console.error("状態取得エラー:", error);
    return { status: "エラー" };
  }
}

/**
 * 日次サマリーの取得
 */
function getDailySummary(
  employeeNumber: string,
  year: number,
  month: number
): any[] {
  try {
    const sheetName = `${year}${String(month).padStart(2, "0")}`;
    console.log(
      `日次サマリー取得開始 - シート: ${sheetName}, 社員番号: ${employeeNumber}`
    );

    const sheet = getAttendanceSheet(sheetName);

    if (!sheet) {
      console.log(`シート ${sheetName} が存在しません`);
      return [];
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    console.log(`シートから ${values.length} 行のデータを取得`);

    if (values.length <= 1) {
      console.log("データが存在しません（ヘッダーのみ）");
      return [];
    }

    // 社員のデータをフィルタリング
    const employeeData = values.slice(1).filter((row) => {
      const rowEmployeeNumber = row[1] ? row[1].toString() : "";
      return rowEmployeeNumber === employeeNumber;
    });

    console.log(`社員 ${employeeNumber} のデータ: ${employeeData.length} 件`);

    if (employeeData.length === 0) {
      console.log(`社員 ${employeeNumber} のデータが見つかりません`);
      return [];
    }

    // データを時系列順にソート（Timestamp列でソート）
    employeeData.sort((a, b) => {
      const timeA = a[3] instanceof Date ? a[3] : new Date(a[3]);
      const timeB = b[3] instanceof Date ? b[3] : new Date(b[3]);
      return timeA.getTime() - timeB.getTime();
    });

    console.log("データを時系列順にソートしました");

    // 日付ごとにグループ化してサマリー作成
    const dailyMap = new Map<string, any>();

    employeeData.forEach((row, index) => {
      console.log(`処理中の行 ${index + 2}:`, row);

      const dateValue = row[0];
      let date: string;

      // 日付の処理を改善
      if (dateValue instanceof Date) {
        date = Utilities.formatDate(dateValue, "Asia/Tokyo", "yyyy-MM-dd");
      } else {
        date = dateValue.toString().split(" ")[0];
      }

      const action = row[2] ? row[2].toString() : "";
      const timestampValue = row[3];
      let timestamp: string;

      // タイムスタンプの処理を改善
      if (timestampValue instanceof Date) {
        timestamp = Utilities.formatDate(
          timestampValue,
          "Asia/Tokyo",
          "yyyy/MM/dd HH:mm:ss"
        );
      } else {
        timestamp = timestampValue.toString();
      }

      console.log(
        `日付: ${date}, アクション: ${action}, タイムスタンプ: ${timestamp}`
      );

      if (!dailyMap.has(date)) {
        dailyMap.set(date, {
          date,
          clockIn: "",
          clockOut: "",
          breakTime: "0",
          workTime: "0",
          overtime: "0",
          halfDay: false,
          fullDay: false,
          holidayWork: false,
          breaks: [], // 中抜け記録
        });
      }

      const dayData = dailyMap.get(date);

      switch (action) {
        case "clockIn":
          // タイムスタンプから時刻部分を抽出
          if (timestamp.includes(" ")) {
            dayData.clockIn = timestamp.split(" ")[1];
          } else {
            dayData.clockIn = timestamp;
          }
          break;
        case "clockOut":
          if (timestamp.includes(" ")) {
            dayData.clockOut = timestamp.split(" ")[1];
          } else {
            dayData.clockOut = timestamp;
          }
          break;
        case "breakStart":
          dayData.breaks.push({ start: timestamp, end: null });
          console.log(`${date} - 中抜け開始を追加: ${timestamp}`);
          break;
        case "breakEnd":
          // 最後の未完了の中抜けを終了
          const lastBreak = dayData.breaks.find((b: any) => b.end === null);
          if (lastBreak) {
            lastBreak.end = timestamp;
            console.log(
              `${date} - 中抜け終了を設定: ${timestamp} (開始: ${lastBreak.start})`
            );
          } else {
            console.log(
              `${date} - 警告: 対応する中抜け開始が見つかりません (終了: ${timestamp})`
            );
          }
          break;
        case "halfDay":
          dayData.halfDay = true;
          break;
        case "fullDay":
          dayData.fullDay = true;
          break;
        case "holidayWork":
          dayData.holidayWork = true;
          break;
      }
    });

    // 勤務時間の計算
    dailyMap.forEach((dayData, date) => {
      // 出退勤がある場合、または全休・半休・休日出勤がある場合に計算
      const hasAttendance = dayData.clockIn && dayData.clockOut;
      const hasSpecialStatus =
        dayData.fullDay || dayData.halfDay || dayData.holidayWork;

      if (hasAttendance || hasSpecialStatus) {
        try {
          let workHours = 0;
          let breakHours = 0;

          // 出退勤がある場合は通常の勤務時間を計算
          if (hasAttendance) {
            // 時刻文字列を解析
            const inTimeStr = dayData.clockIn.includes(":")
              ? dayData.clockIn
              : "00:00:00";
            const outTimeStr = dayData.clockOut.includes(":")
              ? dayData.clockOut
              : "00:00:00";

            const inTime = new Date(`2000/01/01 ${inTimeStr}`);
            const outTime = new Date(`2000/01/01 ${outTimeStr}`);

            // 日をまたいだ場合の処理
            if (outTime < inTime) {
              outTime.setDate(outTime.getDate() + 1);
            }

            workHours =
              (outTime.getTime() - inTime.getTime()) / (1000 * 60 * 60);

            // 中抜け時間を計算
            console.log(`${date} - 中抜け記録数: ${dayData.breaks.length}`);

            dayData.breaks.forEach((breakPeriod: any, index: number) => {
              if (breakPeriod.start && breakPeriod.end) {
                try {
                  // タイムスタンプから時刻部分を抽出
                  let breakStartStr = breakPeriod.start;
                  let breakEndStr = breakPeriod.end;

                  // "yyyy/MM/dd HH:mm:ss" 形式の場合、時刻部分のみ抽出
                  if (breakStartStr.includes(" ")) {
                    breakStartStr = breakStartStr.split(" ")[1];
                  }
                  if (breakEndStr.includes(" ")) {
                    breakEndStr = breakEndStr.split(" ")[1];
                  }

                  // HH:mm:ss 形式でない場合はスキップ
                  if (
                    !breakStartStr.includes(":") ||
                    !breakEndStr.includes(":")
                  ) {
                    console.log(
                      `${date} - 中抜け${
                        index + 1
                      }: 時刻形式が不正 (${breakStartStr} - ${breakEndStr})`
                    );
                    return;
                  }

                  const breakStart = new Date(`2000/01/01 ${breakStartStr}`);
                  const breakEnd = new Date(`2000/01/01 ${breakEndStr}`);

                  if (
                    isNaN(breakStart.getTime()) ||
                    isNaN(breakEnd.getTime())
                  ) {
                    console.log(
                      `${date} - 中抜け${
                        index + 1
                      }: 日付解析エラー (${breakStartStr} - ${breakEndStr})`
                    );
                    return;
                  }

                  // 日をまたいだ場合の処理
                  if (breakEnd < breakStart) {
                    breakEnd.setDate(breakEnd.getDate() + 1);
                  }

                  const periodHours =
                    (breakEnd.getTime() - breakStart.getTime()) /
                    (1000 * 60 * 60);
                  breakHours += periodHours;

                  console.log(
                    `${date} - 中抜け${
                      index + 1
                    }: ${breakStartStr} - ${breakEndStr} = ${periodHours.toFixed(
                      2
                    )}h`
                  );
                } catch (e) {
                  console.error(
                    `${date} - 中抜け${index + 1} 計算エラー:`,
                    e,
                    breakPeriod
                  );
                }
              } else {
                console.log(
                  `${date} - 中抜け${index + 1}: 開始または終了時刻が未設定`,
                  breakPeriod
                );
              }
            });

            console.log(`${date} - 総中抜け時間: ${breakHours.toFixed(2)}h`);
          }

          // サンプルコードに合わせた勤務時間計算
          // 基本：(出勤時刻 - 退勤時刻) - 中抜け時間 - 1時間（昼休憩）
          // 休日出勤/全休/半休の場合は昼休憩を引かない
          let actualWorkHours = workHours - breakHours;

          // 昼休憩（1時間）を基本的に差し引く（特別打刻の場合は除く）
          if (
            !dayData.holidayWork &&
            !dayData.fullDay &&
            !dayData.halfDay &&
            hasAttendance
          ) {
            actualWorkHours -= 1; // 1時間の昼休憩を自動減算
          }

          // 半休・全休による時間加算
          let extraCredit = 0;
          if (dayData.halfDay) {
            extraCredit = 4;
          } else if (dayData.fullDay) {
            extraCredit = 8;
          }

          const effectiveWorkHours = actualWorkHours + extraCredit;

          console.log(
            `${date} - 総勤務時間: ${workHours.toFixed(
              2
            )}h, 中抜け: ${breakHours.toFixed(
              2
            )}h, 昼休憩減算後: ${actualWorkHours.toFixed(
              2
            )}h, 追加クレジット: ${extraCredit}h, 実効勤務時間: ${effectiveWorkHours.toFixed(
              2
            )}h`
          );

          // 結果を保存
          dayData.breakTime = breakHours.toFixed(1);
          dayData.workTime = effectiveWorkHours.toFixed(1);

          // 残業時間の計算（サンプルコードロジック）
          const isHolidayFlag = dayData.holidayWork; // 簡易版：休日判定はholidayWorkフラグのみ
          let overtimeHours = 0;

          if (isHolidayFlag) {
            // 休日出勤の場合：全勤務時間が残業
            overtimeHours = effectiveWorkHours;
          } else if (dayData.fullDay) {
            // 全休の場合：純粋な労働時間のみが残業
            overtimeHours = actualWorkHours > 0 ? actualWorkHours : 0;
          } else {
            // 平日（半休含む）：8時間超過分が残業
            overtimeHours = effectiveWorkHours > 8 ? effectiveWorkHours - 8 : 0;
          }

          dayData.overtime = overtimeHours.toFixed(1);

          console.log(
            `${date} - 最終結果: 実労働時間=${dayData.workTime}h, 中抜け=${dayData.breakTime}h, 残業=${dayData.overtime}h`
          );
        } catch (e) {
          console.error(`${date} の時間計算エラー:`, e);
        }
      }
    });

    const result = Array.from(dailyMap.values());
    console.log(`日次サマリー完了: ${result.length} 日分のデータ`);
    return result;
  } catch (error) {
    console.error("日次サマリー取得エラー:", error);
    return [];
  }
}

/**
 * 月次メトリクスの取得
 */
function getMonthlyMetrics(
  employeeNumber: string,
  year: number,
  month: number
): { workingDays: number; surplusDeficit: number; averageOvertime: number } {
  try {
    console.log(
      `月次メトリクス計算開始 - ${year}年${month}月, 社員番号: ${employeeNumber}`
    );

    const summary = getDailySummary(employeeNumber, year, month);
    console.log(`日次サマリー取得完了: ${summary.length} 日分`);

    if (summary.length === 0) {
      console.log("サマリーデータが存在しません");
      return {
        workingDays: 0,
        surplusDeficit: 0,
        averageOvertime: 0,
      };
    }

    // サンプルコードに合わせた出勤日の計算
    // 休日出勤と全休は出勤日数にカウントしない
    const workingDays = summary.filter((day) => {
      // 実際の勤務記録があり、かつ休日出勤・全休でない日をカウント
      const hasActualWork = (day.clockIn && day.clockIn !== "") || day.halfDay;
      return hasActualWork && !day.holidayWork && !day.fullDay;
    }).length;

    console.log(`出勤日数: ${workingDays}`);

    // 詳細な計算検証を追加
    console.log("=== 詳細計算検証 ===");
    let manualTotalWorkHours = 0;
    let manualTotalBreakHours = 0;

    summary.forEach((day) => {
      const workTime = parseFloat(day.workTime || 0);
      const breakTime = parseFloat(day.breakTime || 0);

      manualTotalWorkHours += workTime;
      manualTotalBreakHours += breakTime;

      if (workTime > 0 || breakTime > 0) {
        console.log(
          `${day.date}: 実労働時間=${workTime}h, 中抜け=${breakTime}h, 出勤=${day.clockIn}, 退勤=${day.clockOut}`
        );
      }
    });

    console.log(
      `手動計算 - 総実労働時間: ${manualTotalWorkHours.toFixed(
        2
      )}h, 総中抜け時間: ${manualTotalBreakHours.toFixed(2)}h`
    );

    // サンプルコードロジックに合わせた計算
    let totalWorkHours = 0;
    let totalOvertime = 0;

    summary.forEach((day) => {
      const workTime = parseFloat(day.workTime || 0);
      const overtime = parseFloat(day.overtime || 0);

      // 全休の場合：基本8時間は過不足計算に含めない、実際の打刻分のみ加算
      if (day.fullDay) {
        // 全休の実打刻時間のみ（8時間を超える分）を過不足に反映
        const actualPunchHours = workTime - 8; // 全休加算分を除く
        totalWorkHours += actualPunchHours > 0 ? actualPunchHours : 0;
      } else {
        totalWorkHours += workTime;
      }

      totalOvertime += overtime;
    });

    console.log(`過不足計算用勤務時間: ${totalWorkHours.toFixed(2)}h`);
    console.log(`総残業時間: ${totalOvertime.toFixed(2)}h`);

    // サンプルコードの過不足計算: totalWorkHours - workingDays * 8
    const surplusDeficit = totalWorkHours - workingDays * 8;

    console.log(
      `標準勤務時間: ${workingDays * 8}h, 過不足: ${surplusDeficit.toFixed(2)}h`
    );

    // サンプルコードの平均残業時間: totalOvertime / workingDays
    const averageOvertime = workingDays > 0 ? totalOvertime / workingDays : 0;

    console.log(
      `平均残業時間: ${averageOvertime.toFixed(
        2
      )}h (出勤日数: ${workingDays}日)`
    );

    const result = {
      workingDays,
      surplusDeficit: Math.round(surplusDeficit * 10) / 10,
      averageOvertime: Math.round(averageOvertime * 10) / 10,
    };

    console.log("月次メトリクス計算完了:", result);

    return result;
  } catch (error) {
    console.error("月次メトリクス取得エラー:", error);
    return {
      workingDays: 0,
      surplusDeficit: 0,
      averageOvertime: 0,
    };
  }
}

/**
 * 勤怠記録の取得
 */

/**
 * 打刻時刻の修正
 */
function updatePunchTime(
  sheetName: string,
  rowNumber: number,
  newTime: string,
  comment: string
): { success: boolean; message: string } {
  try {
    const sheet = getOrCreateAttendanceSheet(sheetName);

    // 時刻を更新
    sheet.getRange(rowNumber, 4).setValue(newTime);

    // コメントがある場合のみ詳細欄に追加
    if (comment && comment.trim() !== "") {
      const currentDetails = sheet.getRange(rowNumber, 5).getValue();
      const updatedDetails = currentDetails
        ? `${currentDetails} | 修正: ${comment}`
        : `修正: ${comment}`;
      sheet.getRange(rowNumber, 5).setValue(updatedDetails);
    }

    return {
      success: true,
      message: "時刻を修正しました",
    };
  } catch (error) {
    console.error("時刻修正エラー:", error);
    return {
      success: false,
      message: "時刻の修正に失敗しました: " + error,
    };
  }
}

/**
 * 打刻アクションの修正
 */
function updatePunchAction(
  sheetName: string,
  rowNumber: number,
  newAction: string,
  comment: string
): { success: boolean; message: string } {
  try {
    const sheet = getOrCreateAttendanceSheet(sheetName);

    // アクションを更新
    sheet.getRange(rowNumber, 3).setValue(newAction);

    // コメントがある場合のみ詳細欄に追加
    if (comment && comment.trim() !== "") {
      const currentDetails = sheet.getRange(rowNumber, 5).getValue();
      const updatedDetails = currentDetails
        ? `${currentDetails} | 修正: ${comment}`
        : `修正: ${comment}`;
      sheet.getRange(rowNumber, 5).setValue(updatedDetails);
    }

    return {
      success: true,
      message: "打刻種類を修正しました",
    };
  } catch (error) {
    console.error("アクション修正エラー:", error);
    return {
      success: false,
      message: "打刻種類の修正に失敗しました: " + error,
    };
  }
}

/**
 * 計算検証用のデバッグ関数
 */
function debugCalculations(
  employeeNumber: string,
  year: number,
  month: number
): { success: boolean; message: string; data?: any } {
  try {
    console.log(
      `=== 計算デバッグ開始 - ${year}年${month}月 社員${employeeNumber} ===`
    );

    const summary = getDailySummary(employeeNumber, year, month);
    const metrics = getMonthlyMetrics(employeeNumber, year, month);

    // 手動計算で検証
    let manualWorkHours = 0;
    let manualBreakHours = 0;
    let manualOvertime = 0;

    const detailedDays = summary.map((day) => {
      const workTime = parseFloat(day.workTime || 0);
      const breakTime = parseFloat(day.breakTime || 0);
      const overtime = parseFloat(day.overtime || 0);

      manualWorkHours += workTime;
      manualBreakHours += breakTime;
      manualOvertime += overtime;

      return {
        date: day.date,
        clockIn: day.clockIn,
        clockOut: day.clockOut,
        workTime,
        breakTime,
        overtime,
        rawBreaks: day.breaks || [],
      };
    });

    const workingDays = summary.filter(
      (day) => day.clockIn || day.fullDay || day.halfDay || day.holidayWork
    ).length;

    const standardHours = workingDays * 8;
    const manualSurplusDeficit = manualWorkHours - standardHours;
    const manualAverageOvertime =
      workingDays > 0 ? manualOvertime / workingDays : 0;

    console.log("=== 手動計算結果 ===");
    console.log(`出勤日数: ${workingDays}`);
    console.log(`総実労働時間: ${manualWorkHours.toFixed(2)}h`);
    console.log(`総中抜け時間: ${manualBreakHours.toFixed(2)}h`);
    console.log(`標準時間: ${standardHours}h`);
    console.log(`過不足: ${manualSurplusDeficit.toFixed(2)}h`);
    console.log(`平均残業: ${manualAverageOvertime.toFixed(2)}h`);

    console.log("=== システム計算結果 ===");
    console.log(`出勤日数: ${metrics.workingDays}`);
    console.log(`過不足: ${metrics.surplusDeficit}h`);
    console.log(`平均残業: ${metrics.averageOvertime}h`);

    return {
      success: true,
      message: "計算デバッグ完了",
      data: {
        summary: detailedDays,
        manual: {
          workingDays,
          totalWorkHours: manualWorkHours,
          totalBreakHours: manualBreakHours,
          surplusDeficit: manualSurplusDeficit,
          averageOvertime: manualAverageOvertime,
        },
        system: metrics,
      },
    };
  } catch (error) {
    console.error("計算デバッグエラー:", error);
    return {
      success: false,
      message: "計算デバッグ失敗: " + String(error),
    };
  }
}

/**
 * スプレッドシート接続テスト用関数
 */
function testSpreadsheetConnection(): {
  success: boolean;
  message: string;
  data?: any;
} {
  try {
    console.log("=== スプレッドシート接続テスト開始 ===");

    // スプレッドシートにアクセス
    const spreadsheet = getAttendanceSpreadsheet();
    console.log("スプレッドシート名:", spreadsheet.getName());

    // 全シートの一覧を取得
    const sheets = spreadsheet.getSheets();
    console.log("利用可能なシート数:", sheets.length);

    const sheetNames = sheets.map((sheet) => sheet.getName());
    console.log("シート名一覧:", sheetNames);

    // 現在の年月シートを確認
    const currentSheetName = getCurrentSheetName();
    console.log("現在の年月シート名:", currentSheetName);

    const currentSheet = getAttendanceSheet(currentSheetName);
    if (currentSheet) {
      const lastRow = currentSheet.getLastRow();
      console.log(`シート "${currentSheetName}" の最終行:`, lastRow);

      if (lastRow > 1) {
        // サンプルデータを取得
        const sampleData = currentSheet
          .getRange(1, 1, Math.min(lastRow, 5), 5)
          .getValues();
        console.log("サンプルデータ:", sampleData);

        return {
          success: true,
          message: "スプレッドシート接続成功",
          data: {
            spreadsheetName: spreadsheet.getName(),
            sheetNames,
            currentSheetName,
            lastRow,
            sampleData,
          },
        };
      } else {
        return {
          success: true,
          message: "スプレッドシート接続成功（データなし）",
          data: {
            spreadsheetName: spreadsheet.getName(),
            sheetNames,
            currentSheetName,
            lastRow: 0,
          },
        };
      }
    } else {
      return {
        success: true,
        message: `スプレッドシート接続成功（シート "${currentSheetName}" は存在しません）`,
        data: {
          spreadsheetName: spreadsheet.getName(),
          sheetNames,
          currentSheetName,
          sheetExists: false,
        },
      };
    }
  } catch (error) {
    console.error("スプレッドシート接続テストエラー:", error);
    return {
      success: false,
      message: "スプレッドシート接続失敗: " + String(error),
    };
  }
}

// ====================================
// 工数管理機能
// ====================================

/**
 * 案件管理用スプレッドシートの取得または作成
 * @returns Spreadsheet オブジェクト
 */
function getOrCreateProjectsSpreadsheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  const properties = PropertiesService.getScriptProperties();
  let spreadsheetId = properties.getProperty("PROJECTS_SPREADSHEET_ID");

  if (spreadsheetId) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      console.error("既存の案件スプレッドシートが見つかりません:", error);
    }
  }

  // 新しい案件スプレッドシートを作成
  const spreadsheet = SpreadsheetApp.create("Times - 案件管理");
  spreadsheetId = spreadsheet.getId();

  // スプレッドシートIDを保存
  properties.setProperty("PROJECTS_SPREADSHEET_ID", spreadsheetId);

  // デフォルトシートを削除
  const defaultSheet = spreadsheet.getSheets()[0];

  // projectsシートを作成
  const projectsSheet = spreadsheet.insertSheet("projects");
  projectsSheet
    .getRange(1, 1, 1, 6)
    .setValues([
      ["案件ID", "案件名", "案件概要", "ステータス", "総工数", "更新日"],
    ]);

  // project_assignmentsシートを作成
  const assignmentsSheet = spreadsheet.insertSheet("project_assignments");
  assignmentsSheet.getRange(1, 1, 1, 2).setValues([["社員番号", "案件ID"]]);

  // デフォルトシートを削除
  if (
    defaultSheet.getName() !== "projects" &&
    defaultSheet.getName() !== "project_assignments"
  ) {
    spreadsheet.deleteSheet(defaultSheet);
  }

  console.log("新しい案件スプレッドシートを作成しました。ID:", spreadsheetId);
  return spreadsheet;
}

/**
 * 案件スプレッドシート内のタブを取得または作成
 * @param projectId 案件ID
 * @param projectName 案件名
 * @returns Sheet オブジェクト
 */
function getOrCreateProjectTab(
  projectId: string,
  projectName: string
): GoogleAppsScript.Spreadsheet.Sheet {
  const spreadsheet = getOrCreateProjectsSpreadsheet();
  const tabName = `${projectId}_${projectName}`;

  let sheet = spreadsheet.getSheetByName(tabName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(tabName);
    sheet
      .getRange(1, 1, 1, 4)
      .setValues([["日付", "社員番号", "工数", "メモ"]]);
  }

  return sheet;
}

/**
 * 案件を作成
 */
function createProject(
  name: string,
  description: string,
  employeeNumber: string
): { success: boolean; message: string; data?: any } {
  try {
    const spreadsheet = getOrCreateProjectsSpreadsheet();

    // projectsシートを取得
    const projectsSheet = spreadsheet.getSheetByName("projects");
    if (!projectsSheet) {
      throw new Error("projectsシートが見つかりません");
    }

    // 新しい案件IDを生成
    const lastRow = projectsSheet.getLastRow();
    let newProjectId = "PROJ001";

    if (lastRow > 1) {
      const lastProjectId = projectsSheet.getRange(lastRow, 1).getValue();
      const lastNumber = parseInt(lastProjectId.replace("PROJ", ""));
      newProjectId = `PROJ${(lastNumber + 1).toString().padStart(3, "0")}`;
    }

    // 案件を追加
    projectsSheet.appendRow([newProjectId, name, description, "open"]);

    // 案件担当者を追加
    assignProjectToUser(newProjectId, employeeNumber);

    // 工数記録用タブを作成
    getOrCreateProjectTab(newProjectId, name);

    return {
      success: true,
      message: "案件を作成しました",
      data: { projectId: newProjectId, name, description },
    };
  } catch (error) {
    console.error("案件作成エラー:", error);
    return {
      success: false,
      message: "案件の作成に失敗しました: " + String(error),
    };
  }
}

/**
 * 案件をユーザーに割り当て
 */
function assignProjectToUser(
  projectId: string,
  employeeNumber: string
): { success: boolean; message: string } {
  try {
    // ユーザースプレッドシートを取得
    const userSpreadsheet = getOrCreateSpreadsheet();

    // project_assignmentsシートを取得または作成
    let assignmentsSheet = userSpreadsheet.getSheetByName(
      "project_assignments"
    );
    if (!assignmentsSheet) {
      assignmentsSheet = userSpreadsheet.insertSheet("project_assignments");
      assignmentsSheet.getRange(1, 1, 1, 2).setValues([["社員番号", "案件ID"]]);

      // ヘッダースタイルの設定
      const headerRange = assignmentsSheet.getRange(1, 1, 1, 2);
      headerRange.setBackground("#4a90e2");
      headerRange.setFontColor("#ffffff");
      headerRange.setFontWeight("bold");

      console.log("project_assignmentsシートを作成しました");
    }

    // 既存の割り当てをチェック
    const data = assignmentsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (
        data[i][1] === projectId &&
        data[i][0].toString() === employeeNumber
      ) {
        return {
          success: true,
          message: "既に割り当て済みです",
        };
      }
    }

    // 新しい割り当てを追加
    assignmentsSheet.appendRow([employeeNumber, projectId]);

    console.log(`案件 ${projectId} を社員 ${employeeNumber} に割り当てました`);

    return {
      success: true,
      message: "案件を割り当てました",
    };
  } catch (error) {
    console.error("案件割り当てエラー:", error);
    return {
      success: false,
      message: "案件の割り当てに失敗しました: " + String(error),
    };
  }
}

/**
 * ユーザーに割り当てられた案件一覧を取得
 */
function getUserProjects(
  employeeNumber: string,
  includeClosed: boolean = false
): { success: boolean; message: string; data?: any[] } {
  try {
    console.log(
      `getUserProjects開始 - 社員番号: ${employeeNumber}, includeClosed: ${includeClosed}`
    );

    const spreadsheet = getOrCreateProjectsSpreadsheet();
    console.log("案件スプレッドシート取得成功:", spreadsheet.getName());

    // ユーザースプレッドシートからproject_assignmentsシートの割り当てられた案件IDを取得
    const userSpreadsheet = getOrCreateSpreadsheet();
    const assignmentsSheet = userSpreadsheet.getSheetByName(
      "project_assignments"
    );
    if (!assignmentsSheet) {
      console.log("project_assignmentsシートが見つかりません");
      return {
        success: true,
        message: "案件が見つかりません",
        data: [],
      };
    }

    const assignmentData = assignmentsSheet.getDataRange().getValues();
    console.log("assignmentData:", assignmentData);

    const userProjectIds = assignmentData
      .slice(1)
      .filter((row) => row[0].toString() === employeeNumber)
      .map((row) => row[1]);

    console.log(
      `社員 ${employeeNumber} に割り当てられた案件ID:`,
      userProjectIds
    );

    if (userProjectIds.length === 0) {
      console.log("割り当てられた案件がありません");
      return {
        success: true,
        message: "割り当てられた案件がありません",
        data: [],
      };
    }

    // projectsシートから案件詳細を取得
    const projectsSheet = spreadsheet.getSheetByName("projects");
    if (!projectsSheet) {
      console.log("projectsシートが見つかりません");
      return {
        success: true,
        message: "案件マスタが見つかりません",
        data: [],
      };
    }

    const projectData = projectsSheet.getDataRange().getValues();
    console.log("projectData:", projectData);

    const projects = [];

    for (const projectId of userProjectIds) {
      const projectRow = projectData.find((row) => row[0] === projectId);
      console.log(`案件ID ${projectId} の詳細:`, projectRow);

      if (projectRow) {
        const status = projectRow[3];
        if (includeClosed || status === "open") {
          // 工数サマリーを取得
          const workloadSummary = getProjectWorkloadSummary(
            projectId,
            projectRow[1],
            employeeNumber
          );
          console.log(`案件 ${projectId} の工数サマリー:`, workloadSummary);

          projects.push({
            projectId: projectRow[0],
            name: projectRow[1],
            description: projectRow[2],
            status: status,
            myWorkload: workloadSummary.myWorkload,
            totalWorkload: workloadSummary.totalWorkload,
            workloadDetails: workloadSummary.details,
          });
        }
      }
    }

    // 案件ID順でソート
    projects.sort((a, b) => a.projectId.localeCompare(b.projectId));

    console.log(`案件一覧取得完了: ${projects.length} 件`, projects);

    return {
      success: true,
      message: "案件一覧を取得しました",
      data: projects,
    };
  } catch (error) {
    console.error("案件一覧取得エラー:", error);
    return {
      success: false,
      message: "案件一覧の取得に失敗しました: " + String(error),
    };
  }
}

/**
 * 案件の工数サマリーを取得
 */
function getProjectWorkloadSummary(
  projectId: string,
  projectName: string,
  employeeNumber: string
): { myWorkload: number; totalWorkload: number; details: any[] } {
  try {
    const spreadsheet = getOrCreateProjectsSpreadsheet();
    const workloadSheetName = `${projectId}_${projectName}`;
    const workloadSheet = spreadsheet.getSheetByName(workloadSheetName);

    if (!workloadSheet) {
      return { myWorkload: 0, totalWorkload: 0, details: [] };
    }

    const data = workloadSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { myWorkload: 0, totalWorkload: 0, details: [] };
    }

    let myWorkload = 0;
    let totalWorkload = 0;
    const details = [];

    for (let i = 1; i < data.length; i++) {
      const [date, recordEmployeeNumber, hours, memo] = data[i];
      const workloadHours = parseFloat(hours) || 0;

      totalWorkload += workloadHours;

      if (recordEmployeeNumber.toString() === employeeNumber) {
        myWorkload += workloadHours;
      }

      details.push({
        date: formatDate(date),
        employeeNumber: recordEmployeeNumber,
        employeeName: getUserNameByEmployeeNumber(
          recordEmployeeNumber.toString()
        ),
        hours: workloadHours,
        memo: memo || "",
      });
    }

    // 日付順でソート（降順）
    details.sort(
      (a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()
    );

    return { myWorkload, totalWorkload, details };
  } catch (error) {
    console.error("工数サマリー取得エラー:", error);
    return { myWorkload: 0, totalWorkload: 0, details: [] };
  }
}

/**
 * projectsタブの総工数と更新日を更新
 */
function updateProjectTotalWorkload(
  projectId: string,
  projectName: string
): { success: boolean; message: string } {
  try {
    console.log(`プロジェクト総工数更新開始: ${projectId}`);

    // 案件スプレッドシートを取得
    const spreadsheet = getOrCreateProjectsSpreadsheet();
    const projectsSheet = spreadsheet.getSheetByName("projects");

    if (!projectsSheet) {
      return {
        success: false,
        message: "projectsシートが見つかりません",
      };
    }

    // 総工数を計算
    const workloadSummary = getProjectWorkloadSummary(
      projectId,
      projectName,
      ""
    );
    const totalWorkload = workloadSummary.totalWorkload;
    const now = new Date();

    console.log(`計算された総工数: ${totalWorkload}`);

    // projectsシートで該当案件を検索して更新
    const projectsData = projectsSheet.getDataRange().getValues();
    let projectRowIndex = -1;

    for (let i = 1; i < projectsData.length; i++) {
      if (projectsData[i][0] === projectId) {
        projectRowIndex = i + 1;
        break;
      }
    }

    if (projectRowIndex === -1) {
      return {
        success: false,
        message: "案件が見つかりません",
      };
    }

    // E列に総工数、F列に更新日を設定
    projectsSheet.getRange(projectRowIndex, 5).setValue(totalWorkload); // E列
    projectsSheet.getRange(projectRowIndex, 6).setValue(now); // F列

    console.log(
      `プロジェクト ${projectId} の総工数を ${totalWorkload} に更新しました`
    );

    return {
      success: true,
      message: "総工数を更新しました",
    };
  } catch (error) {
    console.error("総工数更新エラー:", error);
    return {
      success: false,
      message: "総工数の更新に失敗しました: " + String(error),
    };
  }
}

/**
 * 工数を記録
 */
function recordWorkload(
  projectId: string,
  projectName: string,
  employeeNumber: string,
  date: string,
  hours: number,
  memo: string = ""
): { success: boolean; message: string } {
  try {
    // 案件タブを取得または作成
    const workloadSheet = getOrCreateProjectTab(projectId, projectName);

    // 既存の記録をチェック（同日同ユーザーの場合は上書き）
    const data = workloadSheet.getDataRange().getValues();
    let existingRowIndex = -1;

    console.log(
      `工数記録チェック開始 - 日付: ${date}, 社員番号: ${employeeNumber}`
    );
    console.log(`既存データ行数: ${data.length}`);

    for (let i = 1; i < data.length; i++) {
      const existingDate = formatDate(data[i][0]);
      const existingEmployeeNumber = data[i][1].toString();

      console.log(
        `行${i}: 既存日付=${existingDate}, 既存社員番号=${existingEmployeeNumber}`
      );
      console.log(
        `比較: "${existingDate}" === "${date}" && "${existingEmployeeNumber}" === "${employeeNumber}"`
      );

      if (existingDate === date && existingEmployeeNumber === employeeNumber) {
        existingRowIndex = i + 1;
        console.log(`既存記録発見！行番号: ${existingRowIndex}`);
        break;
      }
    }

    console.log(`最終結果 - existingRowIndex: ${existingRowIndex}`);

    if (existingRowIndex > 0) {
      // 既存記録を更新
      console.log(
        `既存記録を更新: 行${existingRowIndex}, 日付=${date}, 社員番号=${employeeNumber}, 工数=${hours}, メモ=${memo}`
      );
      workloadSheet
        .getRange(existingRowIndex, 1, 1, 4)
        .setValues([[new Date(date), employeeNumber, hours, memo]]);
      console.log(`既存記録の更新完了`);
    } else {
      // 新規記録を追加
      console.log(
        `新規記録を追加: 日付=${date}, 社員番号=${employeeNumber}, 工数=${hours}, メモ=${memo}`
      );
      workloadSheet.appendRow([new Date(date), employeeNumber, hours, memo]);
      console.log(`新規記録の追加完了`);
    }

    // 工数記録後に総工数を更新
    console.log("総工数更新処理を開始");
    const updateResult = updateProjectTotalWorkload(projectId, projectName);
    if (!updateResult.success) {
      console.warn("総工数更新に失敗:", updateResult.message);
    } else {
      console.log("総工数更新成功");
    }

    return {
      success: true,
      message: "工数を記録しました",
    };
  } catch (error) {
    console.error("工数記録エラー:", error);
    return {
      success: false,
      message: "工数の記録に失敗しました: " + String(error),
    };
  }
}

/**
 * 案件情報の更新
 */
function updateProject(
  projectId: string,
  name: string,
  description: string
): { success: boolean; message: string } {
  try {
    console.log(
      `案件情報更新開始 - ID: ${projectId}, 名前: ${name}, 概要: ${description}`
    );

    const spreadsheet = getOrCreateProjectsSpreadsheet();
    const projectsSheet = spreadsheet.getSheetByName("projects");

    if (!projectsSheet) {
      return {
        success: false,
        message: "案件マスタが見つかりません",
      };
    }

    const data = projectsSheet.getDataRange().getValues();
    let projectRowIndex = -1;

    // 該当する案件を検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        projectRowIndex = i + 1;
        break;
      }
    }

    if (projectRowIndex === -1) {
      return {
        success: false,
        message: "案件が見つかりません",
      };
    }

    // 案件名と案件概要を更新
    console.log(
      `案件情報を更新: 行${projectRowIndex}, 名前=${name}, 概要=${description}`
    );
    projectsSheet.getRange(projectRowIndex, 2).setValue(name); // 案件名
    projectsSheet.getRange(projectRowIndex, 3).setValue(description); // 案件概要

    // 工数記録タブの名前も更新
    const oldTabName = `${projectId}_${data[projectRowIndex - 1][1]}`;
    const newTabName = `${projectId}_${name}`;

    if (oldTabName !== newTabName) {
      const workloadSheet = spreadsheet.getSheetByName(oldTabName);
      if (workloadSheet) {
        workloadSheet.setName(newTabName);
        console.log(`工数記録タブ名を更新: ${oldTabName} → ${newTabName}`);
      }
    }

    console.log(`案件情報の更新完了`);

    return {
      success: true,
      message: "案件情報を更新しました",
    };
  } catch (error) {
    console.error("案件情報更新エラー:", error);
    return {
      success: false,
      message: "案件情報の更新に失敗しました: " + String(error),
    };
  }
}

/**
 * 案件のステータスを更新
 */
function updateProjectStatus(
  projectId: string,
  status: "open" | "close"
): { success: boolean; message: string } {
  try {
    const spreadsheet = getOrCreateProjectsSpreadsheet();
    const projectsSheet = spreadsheet.getSheetByName("projects");

    if (!projectsSheet) {
      return {
        success: false,
        message: "案件マスタが見つかりません",
      };
    }

    const data = projectsSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        projectsSheet.getRange(i + 1, 4).setValue(status);
        return {
          success: true,
          message: `案件のステータスを${
            status === "open" ? "オープン" : "クローズ"
          }に更新しました`,
        };
      }
    }

    return {
      success: false,
      message: "案件が見つかりません",
    };
  } catch (error) {
    console.error("案件ステータス更新エラー:", error);
    return {
      success: false,
      message: "案件ステータスの更新に失敗しました: " + String(error),
    };
  }
}

/**
 * 現在のユーザー情報を取得（セッションから）
 */
function getCurrentUser(): { employeeNumber: string; name: string } | null {
  try {
    // Session.getActiveUser() を使用してユーザー情報を取得
    const user = Session.getActiveUser();
    if (!user) return null;

    const userEmail = user.getEmail();
    if (!userEmail) return null;

    // キャッシュからユーザー情報を取得
    const cache = CacheService.getDocumentCache();
    if (!cache) return null;

    const userDataString = cache.get(`user_${userEmail}`);

    if (userDataString) {
      return JSON.parse(userDataString);
    }

    return null;
  } catch (error) {
    console.error("現在ユーザー取得エラー:", error);
    return null;
  }
}

/**
 * 案件の割り当てを解除
 */
function unassignProjectFromUser(
  projectId: string,
  employeeNumber: string
): { success: boolean; message: string } {
  try {
    // ユーザースプレッドシートを取得
    const userSpreadsheet = getOrCreateSpreadsheet();

    // project_assignmentsシートを取得
    const assignmentsSheet = userSpreadsheet.getSheetByName(
      "project_assignments"
    );
    if (!assignmentsSheet) {
      return {
        success: false,
        message: "project_assignmentsシートが見つかりません",
      };
    }

    const data = assignmentsSheet.getDataRange().getValues();
    let rowToDelete = -1;

    // 該当する割り当てを検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === projectId && data[i][0] === employeeNumber) {
        rowToDelete = i + 1;
        break;
      }
    }

    if (rowToDelete === -1) {
      return {
        success: false,
        message: "割り当てが見つかりません",
      };
    }

    // 行を削除
    assignmentsSheet.deleteRow(rowToDelete);

    console.log(
      `案件 ${projectId} の社員 ${employeeNumber} への割り当てを解除しました`
    );

    return {
      success: true,
      message: "案件の割り当てを解除しました",
    };
  } catch (error) {
    console.error("案件割り当て解除エラー:", error);
    return {
      success: false,
      message: "案件の割り当て解除に失敗しました: " + String(error),
    };
  }
}

/**
 * 日付フォーマット用ヘルパー関数
 */
function formatDate(date: Date | string): string {
  if (typeof date === "string") {
    // 既に文字列の場合、YYYY-MM-DD形式であることを確認
    if (date.match(/^\d{4}-\d{2}-\d{2}$/)) {
      return date;
    }
    // その他の文字列形式の場合はDateオブジェクトに変換してフォーマット
    const d = new Date(date);
    if (isNaN(d.getTime())) {
      console.warn(`無効な日付文字列: ${date}`);
      return date; // 無効な場合は元の文字列を返す
    }
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  // Dateオブジェクトの場合
  const d = new Date(date);
  if (isNaN(d.getTime())) {
    console.warn(`無効なDateオブジェクト: ${date}`);
    return String(date);
  }

  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

/**
 * 全案件一覧を取得（割り当て状況含む）
 */
function getAllProjects(
  employeeNumber: string,
  includeClosed: boolean = false
): { success: boolean; message: string; data?: any[] } {
  try {
    console.log(
      `getAllProjects開始 - 社員番号: ${employeeNumber}, includeClosed: ${includeClosed}`
    );

    const spreadsheet = getOrCreateProjectsSpreadsheet();
    console.log("案件スプレッドシート取得成功:", spreadsheet.getName());

    // projectsシートから全案件を取得
    const projectsSheet = spreadsheet.getSheetByName("projects");
    if (!projectsSheet) {
      console.log("projectsシートが見つかりません");
      return {
        success: true,
        message: "案件マスタが見つかりません",
        data: [],
      };
    }

    const projectData = projectsSheet.getDataRange().getValues();
    console.log("全案件データ:", projectData);

    // ユーザースプレッドシートからproject_assignmentsシートの割り当て情報を取得
    const userSpreadsheet = getOrCreateSpreadsheet();
    const assignmentsSheet = userSpreadsheet.getSheetByName(
      "project_assignments"
    );
    let assignmentData: any[][] = [];
    if (assignmentsSheet) {
      assignmentData = assignmentsSheet.getDataRange().getValues();
      console.log("割り当てデータ:", assignmentData);
    } else {
      console.log("project_assignmentsシートが存在しません");
    }

    const projects = [];

    // ヘッダー行をスキップして案件データを処理
    for (let i = 1; i < projectData.length; i++) {
      const [projectId, name, description, status] = projectData[i];

      // クローズ案件を除外する場合のフィルタリング
      if (!includeClosed && status === "close") {
        continue;
      }

      // この案件に割り当てられているユーザーを確認
      const isAssignedToUser = assignmentData.some(
        (row) => row[1] === projectId && row[0].toString() === employeeNumber
      );

      // 工数サマリーを取得
      const workloadSummary = getProjectWorkloadSummary(
        projectId,
        name,
        employeeNumber
      );
      console.log(`案件 ${projectId} の工数サマリー:`, workloadSummary);

      projects.push({
        projectId,
        name,
        description,
        status,
        isAssigned: isAssignedToUser,
        myWorkload: workloadSummary.myWorkload,
        totalWorkload: workloadSummary.totalWorkload,
        workloadDetails: workloadSummary.details,
      });
    }

    // 案件ID順でソート
    projects.sort((a, b) => a.projectId.localeCompare(b.projectId));

    console.log(`全案件一覧取得完了: ${projects.length} 件`, projects);

    return {
      success: true,
      message: "全案件一覧を取得しました",
      data: projects,
    };
  } catch (error) {
    console.error("全案件一覧取得エラー:", error);
    return {
      success: false,
      message: "全案件一覧の取得に失敗しました: " + String(error),
    };
  }
}

/**
 * ホットな案件一覧を取得（E列50時間以上、ステータスがopen）
 */
function getHotProjects(): {
  success: boolean;
  message: string;
  data?: Array<{
    projectId: string;
    name: string;
    totalWorkload: number;
    emoji: string;
  }>;
} {
  try {
    console.log("ホットな案件取得開始");

    const spreadsheet = getOrCreateProjectsSpreadsheet();
    const projectsSheet = spreadsheet.getSheetByName("projects");

    if (!projectsSheet) {
      return {
        success: false,
        message: "projectsシートが見つかりません",
      };
    }

    const data = projectsSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        success: true,
        message: "案件データがありません",
        data: [],
      };
    }

    const hotProjects = [];

    // ヘッダー行をスキップして処理
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const projectId = row[0] ? row[0].toString() : "";
      const name = row[1] ? row[1].toString() : "";
      const status = row[3] ? row[3].toString() : "";
      const totalWorkload = row[4] ? parseFloat(row[4]) || 0 : 0;

      // ステータスがcloseの場合はスキップ
      if (status === "close") {
        continue;
      }

      // 総工数が50時間以上の場合のみ表示
      if (totalWorkload >= 50) {
        let emoji = "";
        if (totalWorkload >= 100) {
          emoji = "☠️";
        } else if (totalWorkload >= 50) {
          emoji = "🔥";
        }

        hotProjects.push({
          projectId,
          name,
          totalWorkload,
          emoji,
        });
      }
    }

    // 総工数の降順でソート
    hotProjects.sort((a, b) => b.totalWorkload - a.totalWorkload);

    console.log(`ホットな案件: ${hotProjects.length}件取得`);

    return {
      success: true,
      message: "ホットな案件を取得しました",
      data: hotProjects,
    };
  } catch (error) {
    console.error("ホットな案件取得エラー:", error);
    return {
      success: false,
      message: "ホットな案件の取得に失敗しました: " + String(error),
    };
  }
}

/**
 * 案件スプレッドシート接続テスト用関数
 */
function testProjectsSpreadsheetConnection(): {
  success: boolean;
  message: string;
  data?: any;
} {
  try {
    console.log("=== 案件スプレッドシート接続テスト開始 ===");

    // 案件スプレッドシートにアクセス
    const spreadsheet = getOrCreateProjectsSpreadsheet();
    console.log("案件スプレッドシート名:", spreadsheet.getName());

    // 全シートの一覧を取得
    const sheets = spreadsheet.getSheets();
    console.log("利用可能なシート数:", sheets.length);

    const sheetNames = sheets.map((sheet) => sheet.getName());
    console.log("シート名一覧:", sheetNames);

    // projectsシートの確認
    const projectsSheet = spreadsheet.getSheetByName("projects");
    let projectsData = null;
    if (projectsSheet) {
      const lastRow = projectsSheet.getLastRow();
      console.log(`projectsシートの最終行:`, lastRow);

      if (lastRow > 0) {
        projectsData = projectsSheet
          .getRange(1, 1, Math.min(lastRow, 5), 4)
          .getValues();
        console.log("projectsシートサンプルデータ:", projectsData);
      }
    }

    // project_assignmentsシートの確認
    const assignmentsSheet = spreadsheet.getSheetByName("project_assignments");
    let assignmentsData = null;
    if (assignmentsSheet) {
      const lastRow = assignmentsSheet.getLastRow();
      console.log(`project_assignmentsシートの最終行:`, lastRow);

      if (lastRow > 0) {
        assignmentsData = assignmentsSheet
          .getRange(1, 1, Math.min(lastRow, 5), 2)
          .getValues();
        console.log(
          "project_assignmentsシートサンプルデータ:",
          assignmentsData
        );
      }
    }

    return {
      success: true,
      message: "案件スプレッドシート接続成功",
      data: {
        spreadsheetName: spreadsheet.getName(),
        spreadsheetId: spreadsheet.getId(),
        sheetNames,
        projectsData,
        assignmentsData,
      },
    };
  } catch (error) {
    console.error("案件スプレッドシート接続テストエラー:", error);
    return {
      success: false,
      message: "案件スプレッドシート接続失敗: " + String(error),
    };
  }
}

/**
 * 全ユーザーの過不足時間を計算してUSER_SPREADSHEET_IDのusersタブを更新
 */
function updateAllUsersSurplusDeficit(): {
  success: boolean;
  message: string;
  data?: any;
} {
  try {
    console.log("=== 全ユーザーの過不足時間更新開始 ===");

    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = spreadsheet.getSheetByName("users");

    if (!sheet) {
      return {
        success: false,
        message: "usersシートが見つかりません",
      };
    }

    // D列とE列が存在しない場合はヘッダーを追加
    const headerRange = sheet.getRange(1, 1, 1, 5);
    const headers = headerRange.getValues()[0];

    if (!headers[3]) {
      sheet.getRange(1, 4).setValue("過不足時間");
    }
    if (!headers[4]) {
      sheet.getRange(1, 5).setValue("更新日時");
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length <= 1) {
      return {
        success: true,
        message: "更新対象のユーザーがいません",
        data: { updatedCount: 0 },
      };
    }

    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;

    let updatedCount = 0;
    const results = [];

    for (let i = 1; i < values.length; i++) {
      const employeeNumber = values[i][0] ? values[i][0].toString() : "";

      if (!employeeNumber) {
        continue;
      }

      try {
        // 過不足時間を計算
        const metrics = getMonthlyMetrics(employeeNumber, year, month);
        const surplusDeficit = metrics.surplusDeficit;

        // D列に過不足時間、E列に更新日時を設定
        sheet.getRange(i + 1, 4).setValue(surplusDeficit);
        sheet.getRange(i + 1, 5).setValue(now);

        updatedCount++;

        results.push({
          employeeNumber,
          name: values[i][1] ? values[i][1].toString() : "",
          surplusDeficit,
        });

        console.log(
          `社員番号: ${employeeNumber}, 過不足時間: ${surplusDeficit}h を更新`
        );
      } catch (error) {
        console.error(
          `社員番号 ${employeeNumber} の過不足時間計算エラー:`,
          error
        );
      }
    }

    console.log(`=== ${updatedCount}名の過不足時間を更新しました ===`);

    return {
      success: true,
      message: `${updatedCount}名の過不足時間を更新しました`,
      data: {
        updatedCount,
        results,
      },
    };
  } catch (error) {
    console.error("過不足時間更新エラー:", error);
    return {
      success: false,
      message: "過不足時間の更新に失敗しました: " + String(error),
    };
  }
}

/**
 * 特定のユーザーの過不足時間を計算してUSER_SPREADSHEET_IDのusersタブを更新
 */
function updateUserSurplusDeficit(employeeNumber: string): {
  success: boolean;
  message: string;
  data?: any;
} {
  try {
    console.log(`=== 社員番号 ${employeeNumber} の過不足時間更新開始 ===`);

    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = spreadsheet.getSheetByName("users");

    if (!sheet) {
      return {
        success: false,
        message: "usersシートが見つかりません",
      };
    }

    // D列とE列が存在しない場合はヘッダーを追加
    const headerRange = sheet.getRange(1, 1, 1, 5);
    const headers = headerRange.getValues()[0];

    if (!headers[3]) {
      sheet.getRange(1, 4).setValue("過不足時間");
    }
    if (!headers[4]) {
      sheet.getRange(1, 5).setValue("更新日時");
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    let rowIndex = -1;

    for (let i = 1; i < values.length; i++) {
      if (values[i][0].toString() === employeeNumber) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return {
        success: false,
        message: `社員番号 ${employeeNumber} が見つかりません`,
      };
    }

    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;

    // 過不足時間を計算
    const metrics = getMonthlyMetrics(employeeNumber, year, month);
    const surplusDeficit = metrics.surplusDeficit;

    // D列に過不足時間、E列に更新日時を設定
    sheet.getRange(rowIndex, 4).setValue(surplusDeficit);
    sheet.getRange(rowIndex, 5).setValue(now);

    console.log(
      `社員番号: ${employeeNumber}, 過不足時間: ${surplusDeficit}h を更新`
    );

    return {
      success: true,
      message: `過不足時間を更新しました: ${surplusDeficit}h`,
      data: {
        employeeNumber,
        surplusDeficit,
        updatedAt: now,
      },
    };
  } catch (error) {
    console.error("過不足時間更新エラー:", error);
    return {
      success: false,
      message: "過不足時間の更新に失敗しました: " + String(error),
    };
  }
}

/**
 * 退勤時に特定のユーザーの過不足時間を計算してUSER_SPREADSHEET_IDのusersタブを更新
 */
function updateUserSurplusDeficitOnClockOut(employeeNumber: string): {
  success: boolean;
  message: string;
  data?: any;
} {
  try {
    console.log(
      `=== 退勤時の社員番号 ${employeeNumber} の過不足時間更新開始 ===`
    );

    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = spreadsheet.getSheetByName("users");

    if (!sheet) {
      return {
        success: false,
        message: "usersシートが見つかりません",
      };
    }

    // D列とE列が存在しない場合はヘッダーを追加
    const headerRange = sheet.getRange(1, 1, 1, 5);
    const headers = headerRange.getValues()[0];

    if (!headers[3]) {
      sheet.getRange(1, 4).setValue("過不足時間");
    }
    if (!headers[4]) {
      sheet.getRange(1, 5).setValue("更新日時");
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    let rowIndex = -1;

    for (let i = 1; i < values.length; i++) {
      if (values[i][0].toString() === employeeNumber) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return {
        success: false,
        message: `社員番号 ${employeeNumber} が見つかりません`,
      };
    }

    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;

    // 過不足時間を計算
    const metrics = getMonthlyMetrics(employeeNumber, year, month);
    const surplusDeficit = metrics.surplusDeficit;

    // D列に過不足時間、E列に更新日時を設定
    sheet.getRange(rowIndex, 4).setValue(surplusDeficit);
    sheet.getRange(rowIndex, 5).setValue(now);

    console.log(
      `退勤時更新: 社員番号: ${employeeNumber}, 過不足時間: ${surplusDeficit}h`
    );

    return {
      success: true,
      message: `過不足時間を更新しました: ${surplusDeficit}h`,
      data: {
        employeeNumber,
        surplusDeficit,
        updatedAt: now,
      },
    };
  } catch (error) {
    console.error("退勤時の過不足時間更新エラー:", error);
    return {
      success: false,
      message: "過不足時間の更新に失敗しました: " + String(error),
    };
  }
}
