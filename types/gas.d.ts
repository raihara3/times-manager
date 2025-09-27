declare namespace GoogleAppsScript {
  namespace Base {
    interface User {
      getEmail(): string;
      getName(): string;
    }
  }

  namespace HTML {
    interface HtmlTemplate {
      evaluate(): HtmlOutput;
    }

    interface HtmlOutput {
      setXFrameOptionsMode(mode: XFrameOptionsMode): HtmlOutput;
      setTitle(title: string): HtmlOutput;
    }

    enum XFrameOptionsMode {
      ALLOWALL,
      DEFAULT,
      DENY
    }
  }

  namespace Spreadsheet {
    interface Spreadsheet {
      getSheetByName(name: string): Sheet | null;
      insertSheet(name?: string): Sheet;
      getSheets(): Sheet[];
    }

    interface Sheet {
      appendRow(rowContents: unknown[]): Sheet;
      getDataRange(): Range;
      getRange(row: number, column: number, numRows?: number, numColumns?: number): Range;
      getLastRow(): number;
      getLastColumn(): number;
      insertRowBefore(beforePosition: number): Sheet;
      deleteRow(rowPosition: number): Sheet;
      clear(): Sheet;
      clearContents(): Sheet;
      setName(name: string): Sheet;
      getName(): string;
    }

    interface Range {
      getValues(): unknown[][];
      setValues(values: unknown[][]): Range;
      getValue(): unknown;
      setValue(value: unknown): Range;
    }
  }

  namespace Utilities {
    function formatDate(date: Date, timeZone: string, format: string): string;
    function parseDate(date: string): Date;
  }
}

declare const SpreadsheetApp: {
  getActiveSpreadsheet(): GoogleAppsScript.Spreadsheet.Spreadsheet;
  openById(id: string): GoogleAppsScript.Spreadsheet.Spreadsheet;
  create(name: string): GoogleAppsScript.Spreadsheet.Spreadsheet;
};

declare const HtmlService: {
  createTemplateFromFile(filename: string): GoogleAppsScript.HTML.HtmlTemplate;
  createHtmlOutputFromFile(filename: string): GoogleAppsScript.HTML.HtmlOutput;
  XFrameOptionsMode: typeof GoogleAppsScript.HTML.XFrameOptionsMode;
};

declare const Session: {
  getActiveUser(): GoogleAppsScript.Base.User;
  getEffectiveUser(): GoogleAppsScript.Base.User;
};

declare const Utilities: typeof GoogleAppsScript.Utilities;