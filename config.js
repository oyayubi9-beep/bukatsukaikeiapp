// ====================================================================
// 設定ファイル
// ここに書く値は、Google Cloud / Google スプレッドシート / Google ドライブ
// の準備が終わったあとに書き換えます。
// README.md の「3. config.js を埋める」を参照してください。
// ====================================================================

const CONFIG = {
  // Google Cloud で発行したOAuthクライアントID（末尾が .apps.googleusercontent.com の文字列）
  GOOGLE_CLIENT_ID: "58597749804-i8nnbj43rlvvhr7c3b81vu264736khtm.apps.googleusercontent.com",

  // 会計システムのGoogleスプレッドシートのID
  // スプレッドシートのURLの /d/ と /edit の間の文字列
  // 例：https://docs.google.com/spreadsheets/d/【ここがID】/edit
  SPREADSHEET_ID: "1noNgJ_hot2bfPhzyFpSZ7BH7nKPqtenzDXc59vsg_gA",

  // 領収書を保存するGoogleドライブのフォルダID（このフォルダの下に月別フォルダを自動作成）
  // フォルダのURLの /folders/ のあとの文字列
  RECEIPT_DRIVE_FOLDER_ID: "1k6I9phH9hMdroKja2MVJ8AunOTftA5OH",

  // 一般会計シート名
  SHEET_LEDGER: "一般会計",
  // 名簿シート名
  SHEET_ROSTER: "名簿",
  // マスタシート名
  SHEET_MASTER: "マスタ",
  // 部費入金管理シート名
  SHEET_DUES: "部費入金管理",

  // データ開始行（ヘッダーの次の行）
  LEDGER_FIRST_ROW: 5,
  ROSTER_FIRST_ROW: 5,
  DUES_FIRST_ROW: 4,

  // 部費入金管理：月の列（4月始まり、A=No, B=氏名, C=世帯ID, D=月会費individual, E=4月...）
  DUES_MONTH_COLUMNS: ["4月","5月","6月","7月","8月","9月","10月","11月","12月","1月","2月","3月"],
  DUES_MONTH_START_COL: "E",
};
