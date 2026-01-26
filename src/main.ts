"use strict";

interface AttendanceRecord {
  rowNumber: number;
  date: string;
  employeeNumber: string;
  action: string;
  timestamp: string;
  details: string;
}

interface BreakPeriod {
  start: string;
  end: string | null;
}

interface DailySummary {
  date: string;
  clockIn: string;
  clockOut: string;
  breakTime: string;
  workTime: string;
  overtime: string;
  requestOvertime: string;
  halfDay: boolean;
  fullDay: boolean;
  holidayWork: boolean;
  breaks: BreakPeriod[];
}

let USER_SPREADSHEET_ID = null;
const ATTENDANCE_SPREADSHEET_ID =
  PropertiesService.getScriptProperties().getProperty(
    "CALENDAR_SPREADSHEET_ID",
  ) || "";
const PROJECTS_SPREADSHEET_ID =
  PropertiesService.getScriptProperties().getProperty(
    "PROJECTS_SPREADSHEET_ID",
  ) || "";
function doGet() {
  return HtmlService.createTemplateFromFile("app")
    .evaluate()
    .setTitle(".Times")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setFaviconUrl("https://twemoji.maxcdn.com/v/14.0.2/72x72/1f392.png")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(filename: string) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function getOrCreateSpreadsheet() {
  const properties = PropertiesService.getScriptProperties();
  let spreadsheetId = properties.getProperty("USER_SPREADSHEET_ID");
  if (spreadsheetId) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      console.error("既存のスプレッドシートが見つかりません:", error);
    }
  }
  const spreadsheet = SpreadsheetApp.create(".Times ユーザーデータベース");
  spreadsheetId = spreadsheet.getId();
  properties.setProperty("USER_SPREADSHEET_ID", spreadsheetId);
  const sheet = spreadsheet.getActiveSheet();
  sheet.setName("users");
  sheet.getRange(1, 1, 1, 3).setValues([["社員番号", "名前", "登録日時"]]);
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setBackground("#4a90e2");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  console.log("新しいスプレッドシートを作成しました: " + spreadsheetId);
  return spreadsheet;
}
function getAllUsers() {
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
    const users = [];
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
function getGuestUsers() {
  try {
    console.log("getGuestUsers開始");
    const allUsersResult = getAllUsers();
    if (!allUsersResult.success || !allUsersResult.data) {
      return {
        success: false,
        message: "ユーザー一覧の取得に失敗しました",
      };
    }
    const guestUsers = allUsersResult.data.filter((user) =>
      user.name.startsWith("ゲスト"),
    );
    console.log(`${guestUsers.length}件のゲストユーザーを取得`);
    return {
      success: true,
      message: "ゲストユーザー一覧を取得しました",
      data: guestUsers,
    };
  } catch (error) {
    console.error("ゲストユーザー一覧取得エラー:", error);
    return {
      success: false,
      message: "ゲストユーザー一覧の取得に失敗しました: " + String(error),
    };
  }
}
function registerUser(employeeNumber: string, name: string) {
  if (!employeeNumber || !name) {
    return { success: false, message: "社員番号と名前を入力してください。" };
  }
  const spreadsheet = getOrCreateSpreadsheet();
  const sheet = spreadsheet.getSheetByName("users");
  if (!sheet) {
    return { success: false, message: "ユーザーシートが見つかりません。" };
  }
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
  const now = new Date();
  sheet.appendRow([employeeNumber, name, now]);
  return {
    success: true,
    message: `ようこそ、${name}さん！登録が完了しました。`,
  };
}
function loginUser(employeeNumber: string) {
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
      const user = {
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
}
function getUserNameByEmployeeNumber(employeeNumber: string) {
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
function getSpreadsheetInfo() {
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
function checkSetup() {
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
function getAttendanceSpreadsheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(ATTENDANCE_SPREADSHEET_ID);
    console.log("勤怠スプレッドシート取得成功:", spreadsheet.getName());
    return spreadsheet;
  } catch (error) {
    console.error("勤怠スプレッドシートの取得に失敗:", error);
    console.error("スプレッドシートID:", ATTENDANCE_SPREADSHEET_ID);
    throw new Error(
      "勤怠スプレッドシートにアクセスできません: " + String(error),
    );
  }
}
function getCurrentSheetName() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  return `${year}${month}`;
}
function getAttendanceSheet(sheetName: string) {
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
function getOrCreateAttendanceSheet(sheetName: string) {
  let sheet = getAttendanceSheet(sheetName);
  if (!sheet) {
    const spreadsheet = getAttendanceSpreadsheet();
    sheet = spreadsheet.insertSheet(sheetName);
    sheet
      .getRange(1, 1, 1, 5)
      .setValues([
        ["Date", "EmployeeNumber", "Action", "Timestamp", "Details"],
      ]);
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setBackground("#4a90e2");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    console.log(`新しいシート "${sheetName}" を作成しました`);
  }
  return sheet;
}
function stampAction(
  employeeNumber: string,
  action: string,
  details = "",
  specifiedDate?: string
) {
  try {
    const sheetName = getCurrentSheetName();
    const sheet = getOrCreateAttendanceSheet(sheetName);
    const now = new Date();
    let date;
    if (specifiedDate) {
      date = specifiedDate;
    } else if (action === "clockIn") {
      date = Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM-dd");
    } else {
      const currentHour = now.getHours();
      if (currentHour < 5) {
        const yesterday = new Date(now);
        yesterday.setDate(yesterday.getDate() - 1);
        date = Utilities.formatDate(yesterday, "Asia/Tokyo", "yyyy-MM-dd");
      } else {
        date = Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM-dd");
      }
    }
    const timestamp = Utilities.formatDate(
      now,
      "Asia/Tokyo",
      "yyyy/MM/dd HH:mm:ss",
    );
    const lastRow = sheet.getLastRow();
    sheet
      .getRange(lastRow + 1, 1, 1, 5)
      .setValues([[date, employeeNumber, action, timestamp, details]]);
    const actionMessages: Record<string, string> = {
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
          `退勤処理完了。社員番号 ${employeeNumber} の過不足時間を更新します`,
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
function getAttendanceRecords(
  employeeNumber: string,
  year: number,
  month: number
): AttendanceRecord[] {
  try {
    const sheetName = `${year}${String(month).padStart(2, "0")}`;
    const sheet = getAttendanceSheet(sheetName);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const records: AttendanceRecord[] = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString() === employeeNumber) {
        records.push({
          rowNumber: i + 1,
          date: data[i][0]
            ? Utilities.formatDate(
                new Date(data[i][0]),
                "Asia/Tokyo",
                "yyyy-MM-dd",
              )
            : "",
          employeeNumber: data[i][1].toString(),
          action: data[i][2].toString(),
          timestamp: data[i][3]
            ? Utilities.formatDate(
                new Date(data[i][3]),
                "Asia/Tokyo",
                "yyyy-MM-dd HH:mm:ss",
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
function getOpenRecord(employeeNumber: string): AttendanceRecord | null {
  try {
    let openRecord: AttendanceRecord | null = null;
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;
    const records = getAttendanceRecords(employeeNumber, year, month);
    records.sort(
      (a, b) =>
        new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime(),
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
        prevMonth,
      );
      prevRecords.sort(
        (a, b) =>
          new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime(),
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
      openRecord,
    );
    return openRecord;
  } catch (error) {
    console.error("オープンレコード取得エラー:", error);
    return null;
  }
}
function getCurrentStatus(employeeNumber: string) {
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
      const now = new Date();
      const year = now.getFullYear();
      const month = now.getMonth() + 1;
      const records = getAttendanceRecords(employeeNumber, year, month);
      if (records.length > 0) {
        const latest = records[records.length - 1];
        const today = new Date().toISOString().split("T")[0];
        console.log(`最新レコード:`, latest);
        console.log(`今日の日付: ${today}, 最新レコードの日付: ${latest.date}`);
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
function getDailySummary(employeeNumber: string, year: number, month: number) {
  try {
    const sheetName = `${year}${String(month).padStart(2, "0")}`;
    console.log(
      `日次サマリー取得開始 - シート: ${sheetName}, 社員番号: ${employeeNumber}`,
    );
    const sheet = getAttendanceSheet(sheetName);
    if (!sheet) {
      console.log(`シート ${sheetName} が存在しません`);
      return [];
    }
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length <= 1) {
      return [];
    }

    const employeeData = values.slice(1).filter((row) => {
      const rowEmployeeNumber = row[1] ? row[1].toString() : "";
      return rowEmployeeNumber === employeeNumber;
    });

    employeeData.sort((a, b) => {
      const timeA = a[3] instanceof Date ? a[3] : new Date(a[3]);
      const timeB = b[3] instanceof Date ? b[3] : new Date(b[3]);
      return timeA.getTime() - timeB.getTime();
    });

    const dailyMap = new Map<string, DailySummary>();
    employeeData.forEach((row) => {
      const dateValue = row[0];
      let date =
        dateValue instanceof Date
          ? Utilities.formatDate(dateValue, "Asia/Tokyo", "yyyy-MM-dd")
          : dateValue.toString().split(" ")[0];

      const action = row[2] ? row[2].toString() : "";
      const timestampValue = row[3];
      let timestamp =
        timestampValue instanceof Date
          ? Utilities.formatDate(
              timestampValue,
              "Asia/Tokyo",
              "yyyy/MM/dd HH:mm:ss",
            )
          : timestampValue.toString();

      if (!dailyMap.has(date)) {
        dailyMap.set(date, {
          date,
          clockIn: "",
          clockOut: "",
          breakTime: "0",
          workTime: "0",
          overtime: "0",
          requestOvertime: "0",
          halfDay: false,
          fullDay: false,
          holidayWork: false,
          breaks: [],
        });
      }

      const dayData = dailyMap.get(date)!;
      const timeOnly = timestamp.includes(" ")
        ? timestamp.split(" ")[1]
        : timestamp;

      switch (action) {
        case "clockIn":
          dayData.clockIn = timeOnly;
          break;
        case "clockOut":
          dayData.clockOut = timeOnly;
          break;
        case "breakStart":
          dayData.breaks.push({ start: timestamp, end: null });
          break;
        case "breakEnd":
          const lastBreak = dayData.breaks.find((b) => b.end === null);
          if (lastBreak) lastBreak.end = timestamp;
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

    const sortedDays = Array.from(dailyMap.values()).sort((a, b) =>
      a.date.localeCompare(b.date),
    );

    // 累積の不足時間を保持する変数（正の数で保持）
    let totalDeficit = 0;

    sortedDays.forEach((dayData) => {
      const hasAttendance = dayData.clockIn && dayData.clockOut;
      const hasSpecialStatus =
        dayData.fullDay || dayData.halfDay || dayData.holidayWork;

      if (hasAttendance || hasSpecialStatus) {
        try {
          let workHours = 0;
          let breakHours = 0;

          if (hasAttendance) {
            const inTime = new Date(`2000/01/01 ${dayData.clockIn}`);
            const outTime = new Date(`2000/01/01 ${dayData.clockOut}`);
            if (outTime < inTime) outTime.setDate(outTime.getDate() + 1);
            workHours =
              (outTime.getTime() - inTime.getTime()) / (1000 * 60 * 60);

            dayData.breaks.forEach((breakPeriod) => {
              if (breakPeriod.start && breakPeriod.end) {
                let bS = breakPeriod.start.includes(" ")
                  ? breakPeriod.start.split(" ")[1]
                  : breakPeriod.start;
                let bE = breakPeriod.end.includes(" ")
                  ? breakPeriod.end.split(" ")[1]
                  : breakPeriod.end;
                const breakStart = new Date(`2000/01/01 ${bS}`);
                const breakEnd = new Date(`2000/01/01 ${bE}`);
                if (breakEnd < breakStart)
                  breakEnd.setDate(breakEnd.getDate() + 1);
                breakHours +=
                  (breakEnd.getTime() - breakStart.getTime()) /
                  (1000 * 60 * 60);
              }
            });
          }

          let actualWorkHours = workHours - breakHours;
          if (
            !dayData.holidayWork &&
            !dayData.fullDay &&
            !dayData.halfDay &&
            hasAttendance
          ) {
            actualWorkHours -= 1;
          }

          let extraCredit = dayData.halfDay ? 4 : dayData.fullDay ? 8 : 0;
          const effectiveWorkHours = actualWorkHours + extraCredit;

          dayData.breakTime = breakHours.toFixed(1);
          dayData.workTime = effectiveWorkHours.toFixed(1);

          // --- requestOvertime ロジックの再定義 ---
          // 1. その日の素の残業時間を計算
          let rawOvertime = 0;
          if (dayData.holidayWork) {
            rawOvertime = effectiveWorkHours;
          } else if (dayData.fullDay) {
            rawOvertime = actualWorkHours > 0 ? actualWorkHours : 0;
          } else {
            rawOvertime = effectiveWorkHours > 8 ? effectiveWorkHours - 8 : 0;
          }

          // 純粋に8時間に満たない分を不足とする
          let dailyDeficit = 0;
          if (!dayData.holidayWork && effectiveWorkHours < 8) {
            dailyDeficit = 8 - effectiveWorkHours;
          }

          if (dailyDeficit > 0) {
            // 不足がある日は requestOvertime は 0。不足を累積。
            dayData.requestOvertime = "0";
            totalDeficit += dailyDeficit;
          } else {
            // 残業がある場合、過去の不足分を差し引く
            let adjustedOvertime = rawOvertime;
            if (totalDeficit > 0) {
              const payOff = Math.min(adjustedOvertime, totalDeficit);
              adjustedOvertime -= payOff;
              totalDeficit -= payOff;
            }
            dayData.requestOvertime = adjustedOvertime.toFixed(1);
          }

          // 画面表示用の残業時間はそのまま
          dayData.overtime = rawOvertime.toFixed(1);
        } catch (e) {
          console.error(`${dayData.date} の時間計算エラー:`, e);
        }
      }
    });

    console.log(`日次サマリー完了: ${sortedDays.length} 日分のデータ`);
    return sortedDays;
  } catch (error) {
    console.error("日次サマリー取得エラー:", error);
    return [];
  }
}
function getMonthlyMetrics(
  employeeNumber: string,
  year: number,
  month: number
) {
  try {
    console.log(
      `月次メトリクス計算開始 - ${year}年${month}月, 社員番号: ${employeeNumber}`,
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
    const workingDays = summary.filter((day) => {
      const hasActualWork = (day.clockIn && day.clockIn !== "") || day.halfDay;
      return hasActualWork && !day.holidayWork && !day.fullDay;
    }).length;
    console.log(`出勤日数: ${workingDays}`);
    console.log("=== 詳細計算検証 ===");
    let manualTotalWorkHours = 0;
    let manualTotalBreakHours = 0;
    summary.forEach((day) => {
      const workTime = parseFloat(day.workTime || "0");
      const breakTime = parseFloat(day.breakTime || "0");
      manualTotalWorkHours += workTime;
      manualTotalBreakHours += breakTime;
      if (workTime > 0 || breakTime > 0) {
        console.log(
          `${day.date}: 実労働時間=${workTime}h, 中抜け=${breakTime}h, 出勤=${day.clockIn}, 退勤=${day.clockOut}`,
        );
      }
    });
    console.log(
      `手動計算 - 総実労働時間: ${manualTotalWorkHours.toFixed(2)}h, 総中抜け時間: ${manualTotalBreakHours.toFixed(2)}h`,
    );
    let totalWorkHours = 0;
    let totalOvertime = 0;
    summary.forEach((day) => {
      const workTime = parseFloat(day.workTime || "0");
      const overtime = parseFloat(day.overtime || "0");
      if (day.fullDay) {
        const actualPunchHours = workTime - 8;
        totalWorkHours += actualPunchHours > 0 ? actualPunchHours : 0;
      } else {
        totalWorkHours += workTime;
      }
      totalOvertime += overtime;
    });
    console.log(`過不足計算用勤務時間: ${totalWorkHours.toFixed(2)}h`);
    console.log(`総残業時間: ${totalOvertime.toFixed(2)}h`);
    const surplusDeficit = totalWorkHours - workingDays * 8;
    console.log(
      `標準勤務時間: ${workingDays * 8}h, 過不足: ${surplusDeficit.toFixed(2)}h`,
    );
    const averageOvertime = workingDays > 0 ? totalOvertime / workingDays : 0;
    console.log(
      `平均残業時間: ${averageOvertime.toFixed(2)}h (出勤日数: ${workingDays}日)`,
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
function updatePunchTime(
  sheetName: string,
  rowNumber: number,
  newTime: string,
  comment: string
) {
  try {
    const sheet = getOrCreateAttendanceSheet(sheetName);
    sheet.getRange(rowNumber, 4).setValue(newTime);
    const newTimeTmp = new Date(newTime);
    const newTimeDay = `${newTimeTmp.getFullYear()}-${newTimeTmp.getMonth() + 1}-${newTimeTmp.getDate()}`;
    sheet.getRange(rowNumber, 1).setValue(newTimeDay);
    if (comment && comment.trim() !== "") {
      const currentDetails = sheet.getRange(rowNumber, 5).getValue();
      const updatedDetails = currentDetails
        ? `${currentDetails} \n更新: ${comment}`
        : `更新: ${comment}`;
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
function updatePunchAction(
  sheetName: string,
  rowNumber: number,
  newAction: string,
  comment: string
) {
  try {
    const sheet = getOrCreateAttendanceSheet(sheetName);
    sheet.getRange(rowNumber, 3).setValue(newAction);
    if (comment && comment.trim() !== "") {
      const currentDetails = sheet.getRange(rowNumber, 5).getValue();
      const updatedDetails = currentDetails
        ? `${currentDetails} \n更新: ${comment}`
        : `更新: ${comment}`;
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
function debugCalculations(
  employeeNumber: string,
  year: number,
  month: number
) {
  try {
    console.log(
      `=== 計算デバッグ開始 - ${year}年${month}月 社員${employeeNumber} ===`,
    );
    const summary = getDailySummary(employeeNumber, year, month);
    const metrics = getMonthlyMetrics(employeeNumber, year, month);
    let manualWorkHours = 0;
    let manualBreakHours = 0;
    let manualOvertime = 0;
    const detailedDays = summary.map((day) => {
      const workTime = parseFloat(day.workTime || "0");
      const breakTime = parseFloat(day.breakTime || "0");
      const overtime = parseFloat(day.overtime || "0");
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
      (day) => day.clockIn || day.fullDay || day.halfDay || day.holidayWork,
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
function testSpreadsheetConnection() {
  try {
    console.log("=== スプレッドシート接続テスト開始 ===");
    const spreadsheet = getAttendanceSpreadsheet();
    console.log("スプレッドシート名:", spreadsheet.getName());
    const sheets = spreadsheet.getSheets();
    console.log("利用可能なシート数:", sheets.length);
    const sheetNames = sheets.map((sheet) => sheet.getName());
    console.log("シート名一覧:", sheetNames);
    const currentSheetName = getCurrentSheetName();
    console.log("現在の年月シート名:", currentSheetName);
    const currentSheet = getAttendanceSheet(currentSheetName);
    if (currentSheet) {
      const lastRow = currentSheet.getLastRow();
      console.log(`シート "${currentSheetName}" の最終行:`, lastRow);
      if (lastRow > 1) {
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
function getOrCreateProjectsSpreadsheet() {
  const properties = PropertiesService.getScriptProperties();
  let spreadsheetId = properties.getProperty("PROJECTS_SPREADSHEET_ID");
  if (spreadsheetId) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      console.error("既存の案件スプレッドシートが見つかりません:", error);
      console.error("保存されていたID:", spreadsheetId);
      throw new Error(
        `案件スプレッドシート（ID: ${spreadsheetId}）にアクセスできません。スクリプトプロパティを確認してください。`,
      );
    }
  }
  console.warn("⚠️ 警告: 新しい案件スプレッドシートを作成します");
  const spreadsheet = SpreadsheetApp.create("Times - 案件管理");
  spreadsheetId = spreadsheet.getId();
  console.log("新しいスプレッドシートID:", spreadsheetId);
  properties.setProperty("PROJECTS_SPREADSHEET_ID", spreadsheetId);
  const defaultSheet = spreadsheet.getSheets()[0];
  const projectsSheet = spreadsheet.insertSheet("projects");
  projectsSheet
    .getRange(1, 1, 1, 7)
    .setValues([
      [
        "案件ID",
        "案件名",
        "案件概要",
        "ステータス",
        "総工数",
        "更新日",
        "予算",
      ],
    ]);
  const assignmentsSheet = spreadsheet.insertSheet("project_assignments");
  assignmentsSheet.getRange(1, 1, 1, 2).setValues([["社員番号", "案件ID"]]);
  if (
    defaultSheet.getName() !== "projects" &&
    defaultSheet.getName() !== "project_assignments"
  ) {
    spreadsheet.deleteSheet(defaultSheet);
  }
  console.log("新しい案件スプレッドシートを作成しました。ID:", spreadsheetId);
  return spreadsheet;
}
function getOrCreateProjectTab(projectId: string, projectName: string) {
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
function createProject(
  name: string,
  description: string,
  employeeNumber: string,
  budget?: number
) {
  try {
    const spreadsheet = getOrCreateProjectsSpreadsheet();
    const projectsSheet = spreadsheet.getSheetByName("projects");
    if (!projectsSheet) {
      throw new Error("projectsシートが見つかりません");
    }
    const lastRow = projectsSheet.getLastRow();
    let newProjectId = "PROJ001";
    if (lastRow > 1) {
      const lastProjectId = projectsSheet.getRange(lastRow, 1).getValue();
      const lastNumber = parseInt(lastProjectId.replace("PROJ", ""));
      newProjectId = `PROJ${(lastNumber + 1).toString().padStart(3, "0")}`;
    }
    projectsSheet.appendRow([
      newProjectId,
      name,
      description,
      "open",
      "",
      "",
      budget !== undefined && budget !== null ? budget : "",
    ]);
    assignProjectToUser(newProjectId, employeeNumber);
    getOrCreateProjectTab(newProjectId, name);
    return {
      success: true,
      message: "案件を作成しました",
      data: { projectId: newProjectId, name, description, budget },
    };
  } catch (error) {
    console.error("案件作成エラー:", error);
    return {
      success: false,
      message: "案件の作成に失敗しました: " + String(error),
    };
  }
}
function assignProjectToUser(projectId: string, employeeNumber: string) {
  try {
    const userSpreadsheet = getOrCreateSpreadsheet();
    let assignmentsSheet = userSpreadsheet.getSheetByName(
      "project_assignments",
    );
    if (!assignmentsSheet) {
      assignmentsSheet = userSpreadsheet.insertSheet("project_assignments");
      assignmentsSheet.getRange(1, 1, 1, 2).setValues([["社員番号", "案件ID"]]);
      const headerRange = assignmentsSheet.getRange(1, 1, 1, 2);
      headerRange.setBackground("#4a90e2");
      headerRange.setFontColor("#ffffff");
      headerRange.setFontWeight("bold");
      console.log("project_assignmentsシートを作成しました");
    }
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
function getUserProjects(employeeNumber: string, includeClosed = false) {
  try {
    console.log(
      `getUserProjects開始 - 社員番号: ${employeeNumber}, includeClosed: ${includeClosed}`,
    );
    const spreadsheet = getOrCreateProjectsSpreadsheet();
    console.log("案件スプレッドシート取得成功:", spreadsheet.getName());
    const userSpreadsheet = getOrCreateSpreadsheet();
    const assignmentsSheet = userSpreadsheet.getSheetByName(
      "project_assignments",
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
      userProjectIds,
    );
    if (userProjectIds.length === 0) {
      console.log("割り当てられた案件がありません");
      return {
        success: true,
        message: "割り当てられた案件がありません",
        data: [],
      };
    }
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
          projects.push({
            projectId: projectRow[0],
            name: projectRow[1],
            description: projectRow[2],
            status: status,
            totalWorkload: projectRow[4] || 0,
            budget:
              projectRow[6] !== undefined && projectRow[6] !== ""
                ? projectRow[6]
                : null,
          });
        }
      }
    }
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
function getProjectWorkloadSummary(
  projectId: string,
  projectName: string,
  employeeNumber: string
) {
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
          recordEmployeeNumber.toString(),
        ),
        hours: workloadHours,
        memo: memo || "",
      });
    }
    details.sort(
      (a, b) => new Date(b.date).getTime() - new Date(a.date).getTime(),
    );
    return { myWorkload, totalWorkload, details };
  } catch (error) {
    console.error("工数サマリー取得エラー:", error);
    return { myWorkload: 0, totalWorkload: 0, details: [] };
  }
}
function updateProjectTotalWorkload(projectId: string, projectName: string) {
  try {
    console.log(`プロジェクト総工数更新開始: ${projectId}`);
    const spreadsheet = getOrCreateProjectsSpreadsheet();
    const projectsSheet = spreadsheet.getSheetByName("projects");
    if (!projectsSheet) {
      return {
        success: false,
        message: "projectsシートが見つかりません",
      };
    }
    const workloadSummary = getProjectWorkloadSummary(
      projectId,
      projectName,
      "",
    );
    const totalWorkload = workloadSummary.totalWorkload;
    const now = new Date();
    console.log(`計算された総工数: ${totalWorkload}`);
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
    projectsSheet.getRange(projectRowIndex, 5).setValue(totalWorkload);
    projectsSheet.getRange(projectRowIndex, 6).setValue(now);
    console.log(
      `プロジェクト ${projectId} の総工数を ${totalWorkload} に更新しました`,
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
function recordWorkload(
  projectId: string,
  projectName: string,
  employeeNumber: string,
  date: string,
  hours: number,
  memo = ""
) {
  try {
    const workloadSheet = getOrCreateProjectTab(projectId, projectName);
    const data = workloadSheet.getDataRange().getValues();
    let existingRowIndex = -1;
    console.log(
      `工数記録チェック開始 - 日付: ${date}, 社員番号: ${employeeNumber}`,
    );
    console.log(`既存データ行数: ${data.length}`);
    for (let i = 1; i < data.length; i++) {
      const existingDate = formatDate(data[i][0]);
      const existingEmployeeNumber = data[i][1].toString();
      console.log(
        `行${i}: 既存日付=${existingDate}, 既存社員番号=${existingEmployeeNumber}`,
      );
      console.log(
        `比較: "${existingDate}" === "${date}" && "${existingEmployeeNumber}" === "${employeeNumber}"`,
      );
      if (existingDate === date && existingEmployeeNumber === employeeNumber) {
        existingRowIndex = i + 1;
        console.log(`既存記録発見！行番号: ${existingRowIndex}`);
        break;
      }
    }
    console.log(`最終結果 - existingRowIndex: ${existingRowIndex}`);
    if (existingRowIndex > 0) {
      console.log(
        `既存記録を更新: 行${existingRowIndex}, 日付=${date}, 社員番号=${employeeNumber}, 工数=${hours}, メモ=${memo}`,
      );
      workloadSheet
        .getRange(existingRowIndex, 1, 1, 4)
        .setValues([[new Date(date), employeeNumber, hours, memo]]);
      console.log(`既存記録の更新完了`);
    } else {
      console.log(
        `新規記録を追加: 日付=${date}, 社員番号=${employeeNumber}, 工数=${hours}, メモ=${memo}`,
      );
      workloadSheet.appendRow([new Date(date), employeeNumber, hours, memo]);
      console.log(`新規記録の追加完了`);
    }
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
function updateProject(
  projectId: string,
  name: string,
  description: string,
  budget?: number
) {
  try {
    console.log(
      `案件情報更新開始 - ID: ${projectId}, 名前: ${name}, 概要: ${description}, 予算: ${budget}`,
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
    console.log(
      `案件情報を更新: 行${projectRowIndex}, 名前=${name}, 概要=${description}, 予算=${budget}`,
    );
    projectsSheet.getRange(projectRowIndex, 2).setValue(name);
    projectsSheet.getRange(projectRowIndex, 3).setValue(description);
    const budgetValue = budget !== undefined && budget !== null ? budget : "";
    projectsSheet.getRange(projectRowIndex, 7).setValue(budgetValue);
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
function updateProjectStatus(projectId: string, status: string) {
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
          message: `案件のステータスを${status === "open" ? "オープン" : "クローズ"}に更新しました`,
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
function getCurrentUser() {
  try {
    const user = Session.getActiveUser();
    if (!user) return null;
    const userEmail = user.getEmail();
    if (!userEmail) return null;
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
function unassignProjectFromUser(projectId: string, employeeNumber: string) {
  try {
    const userSpreadsheet = getOrCreateSpreadsheet();
    const assignmentsSheet = userSpreadsheet.getSheetByName(
      "project_assignments",
    );
    if (!assignmentsSheet) {
      return {
        success: false,
        message: "project_assignmentsシートが見つかりません",
      };
    }
    const data = assignmentsSheet.getDataRange().getValues();
    let rowToDelete = -1;
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
    assignmentsSheet.deleteRow(rowToDelete);
    console.log(
      `案件 ${projectId} の社員 ${employeeNumber} への割り当てを解除しました`,
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
function formatDate(date: string | Date) {
  if (typeof date === "string") {
    if (date.match(/^\d{4}-\d{2}-\d{2}$/)) {
      return date;
    }
    const d = new Date(date);
    if (isNaN(d.getTime())) {
      console.warn(`無効な日付文字列: ${date}`);
      return date;
    }
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
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
function getAllProjects(employeeNumber: string, includeClosed = false) {
  try {
    console.log(
      `getAllProjects開始 - 社員番号: ${employeeNumber}, includeClosed: ${includeClosed}`,
    );
    const spreadsheet = getOrCreateProjectsSpreadsheet();
    console.log("案件スプレッドシート取得成功:", spreadsheet.getName());
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
    const userSpreadsheet = getOrCreateSpreadsheet();
    const assignmentsSheet = userSpreadsheet.getSheetByName(
      "project_assignments",
    );
    let assignmentData: unknown[][] = [];
    if (assignmentsSheet) {
      assignmentData = assignmentsSheet.getDataRange().getValues();
      console.log("割り当てデータ:", assignmentData);
    } else {
      console.log("project_assignmentsシートが存在しません");
    }
    const projects = [];
    for (let i = 1; i < projectData.length; i++) {
      const projectRow = projectData[i];
      const [projectId, name, description, status] = projectRow;
      if (!includeClosed && status === "close") {
        continue;
      }
      const isAssignedToUser = assignmentData.some(
        (row) =>
          row[1] === projectId && String(row[0]) === employeeNumber,
      );
      projects.push({
        projectId,
        name,
        description,
        status,
        isAssigned: isAssignedToUser,
        totalWorkload: projectRow[4] || 0,
        budget:
          projectRow[6] !== undefined && projectRow[6] !== ""
            ? projectRow[6]
            : null,
      });
    }
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
function getHotProjects() {
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
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const projectId = row[0] ? row[0].toString() : "";
      const name = row[1] ? row[1].toString() : "";
      const status = row[3] ? row[3].toString() : "";
      const totalWorkload = row[4] ? parseFloat(row[4]) || 0 : 0;
      if (status === "close") {
        continue;
      }
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
function testProjectsSpreadsheetConnection() {
  try {
    console.log("=== 案件スプレッドシート接続テスト開始 ===");
    const spreadsheet = getOrCreateProjectsSpreadsheet();
    console.log("案件スプレッドシート名:", spreadsheet.getName());
    const sheets = spreadsheet.getSheets();
    console.log("利用可能なシート数:", sheets.length);
    const sheetNames = sheets.map((sheet) => sheet.getName());
    console.log("シート名一覧:", sheetNames);
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
          assignmentsData,
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
function updateAllUsersSurplusDeficit() {
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
        const metrics = getMonthlyMetrics(employeeNumber, year, month);
        const surplusDeficit = metrics.surplusDeficit;
        sheet.getRange(i + 1, 4).setValue(surplusDeficit);
        sheet.getRange(i + 1, 5).setValue(now);
        updatedCount++;
        results.push({
          employeeNumber,
          name: values[i][1] ? values[i][1].toString() : "",
          surplusDeficit,
        });
        console.log(
          `社員番号: ${employeeNumber}, 過不足時間: ${surplusDeficit}h を更新`,
        );
      } catch (error) {
        console.error(
          `社員番号 ${employeeNumber} の過不足時間計算エラー:`,
          error,
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
function updateUserSurplusDeficit(employeeNumber: string) {
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
    const metrics = getMonthlyMetrics(employeeNumber, year, month);
    const surplusDeficit = metrics.surplusDeficit;
    sheet.getRange(rowIndex, 4).setValue(surplusDeficit);
    sheet.getRange(rowIndex, 5).setValue(now);
    console.log(
      `社員番号: ${employeeNumber}, 過不足時間: ${surplusDeficit}h を更新`,
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
function updateUserSurplusDeficitOnClockOut(employeeNumber: string) {
  try {
    console.log(
      `=== 退勤時の社員番号 ${employeeNumber} の過不足時間更新開始 ===`,
    );
    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = spreadsheet.getSheetByName("users");
    if (!sheet) {
      return {
        success: false,
        message: "usersシートが見つかりません",
      };
    }
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
    const metrics = getMonthlyMetrics(employeeNumber, year, month);
    const surplusDeficit = metrics.surplusDeficit;
    sheet.getRange(rowIndex, 4).setValue(surplusDeficit);
    sheet.getRange(rowIndex, 5).setValue(now);
    console.log(
      `退勤時更新: 社員番号: ${employeeNumber}, 過不足時間: ${surplusDeficit}h`,
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
//# sourceMappingURL=main.js.map
