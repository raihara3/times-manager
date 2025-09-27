/**
 * Times. - 工数管理システム
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

/**
 * Webアプリケーションのエントリポイント
 * シングルページアプリケーションのHTML
 */
function doGet(): GoogleAppsScript.HTML.HtmlOutput {
  return HtmlService.createTemplateFromFile("app")
    .evaluate()
    .setTitle("Times.")
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
  let spreadsheetId = properties.getProperty("USER_SPREADSHEET_ID");

  if (spreadsheetId) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      console.error("既存のスプレッドシートが見つかりません:", error);
    }
  }

  // 新しいスプレッドシートを作成
  const spreadsheet = SpreadsheetApp.create("Times. ユーザーデータベース");
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
        console.log(`最新レコード:`, latest);

        if (latest.action === "clockOut") {
          status = "退勤済み";
        } else if (latest.action === "breakEnd") {
          status = "出勤中";
        } else if (
          latest.action === "holidayWork" ||
          latest.action === "fullDay" ||
          latest.action === "halfDay"
        ) {
          status = "出勤中";
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
      if (dayData.clockIn && dayData.clockOut) {
        try {
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

          let workHours =
            (outTime.getTime() - inTime.getTime()) / (1000 * 60 * 60);

          // 中抜け時間を計算
          let breakHours = 0;
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

                if (isNaN(breakStart.getTime()) || isNaN(breakEnd.getTime())) {
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

          // サンプルコードに合わせた勤務時間計算
          // 基本：(出勤時刻 - 退勤時刻) - 中抜け時間 - 1時間（昼休憩）
          // 休日出勤/全休/半休の場合は昼休憩を引かない
          let actualWorkHours = workHours - breakHours;

          // 昼休憩（1時間）を基本的に差し引く
          if (!dayData.holidayWork && !dayData.fullDay && !dayData.halfDay) {
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
          } else {
            // 平日：8時間超過分が残業
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
    // 休日出勤は出勤日数にカウントしない
    const workingDays = summary.filter((day) => {
      // 何らかの勤務記録があり、かつ休日出勤でない日をカウント
      const hasWork =
        (day.clockIn && day.clockIn !== "") || day.fullDay || day.halfDay;
      return hasWork && !day.holidayWork;
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
      totalWorkHours += parseFloat(day.workTime || 0);
      totalOvertime += parseFloat(day.overtime || 0);
    });

    console.log(`総勤務時間: ${totalWorkHours.toFixed(2)}h`);
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
    if (comment && comment.trim() !== '') {
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
    if (comment && comment.trim() !== '') {
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
