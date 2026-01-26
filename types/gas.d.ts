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
      setFaviconUrl(iconUrl: string): HtmlOutput;
      addMetaTag(name: string, content: string): HtmlOutput;
      getContent(): string;
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
      getId(): string;
      getUrl(): string;
      getName(): string;
      getActiveSheet(): Sheet;
      deleteSheet(sheet: Sheet): void;
    }

    interface Sheet {
      appendRow(rowContents: unknown[]): Sheet;
      getDataRange(): Range;
      getRange(
        row: number,
        column: number,
        numRows?: number,
        numColumns?: number
      ): Range;
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
      setBackground(color: string): Range;
      setFontColor(color: string): Range;
      setFontWeight(weight: string): Range;
    }
  }

  namespace Properties {
    interface Properties {
      getProperty(key: string): string | null;
      setProperty(key: string, value: string): Properties;
      deleteProperty(key: string): Properties;
      getProperties(): { [key: string]: string };
    }
  }

  namespace Cache {
    interface Cache {
      get(key: string): string | null;
      put(key: string, value: string, expirationInSeconds?: number): void;
      remove(key: string): void;
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
  createTemplateFromFile(
    filename: string
  ): GoogleAppsScript.HTML.HtmlTemplate;
  createHtmlOutputFromFile(filename: string): GoogleAppsScript.HTML.HtmlOutput;
  XFrameOptionsMode: typeof GoogleAppsScript.HTML.XFrameOptionsMode;
};

declare const Session: {
  getActiveUser(): GoogleAppsScript.Base.User;
  getEffectiveUser(): GoogleAppsScript.Base.User;
};

declare const PropertiesService: {
  getScriptProperties(): GoogleAppsScript.Properties.Properties;
  getUserProperties(): GoogleAppsScript.Properties.Properties;
  getDocumentProperties(): GoogleAppsScript.Properties.Properties;
};

declare const CacheService: {
  getDocumentCache(): GoogleAppsScript.Cache.Cache | null;
  getScriptCache(): GoogleAppsScript.Cache.Cache;
  getUserCache(): GoogleAppsScript.Cache.Cache;
};

declare const Utilities: typeof GoogleAppsScript.Utilities;
