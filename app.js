// ====================================================================
// 部活会計入力アプリ - メインロジック
// Google Identity Services でログインし、ブラウザから直接
// Google Sheets API / Google Drive API を呼び出します（サーバー不要）。
// ====================================================================

const SCOPES = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file";

let accessToken = null;
let tokenClient = null;

let masterData = { categories: [], incomeAccounts: [], expenseAccounts: [], paymentMethods: [] };
let rosterData = []; // [{name, row}]
let specialSheetNames = [];
let sheetIdCache = null;

// --------------------------------------------------------------------
// 認証
// --------------------------------------------------------------------

function initGoogleAuth() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CONFIG.GOOGLE_CLIENT_ID,
    scope: SCOPES,
    callback: async (resp) => {
      if (resp.error) {
        showError("loginError", "ログインに失敗しました。もう一度お試しください。");
        return;
      }
      accessToken = resp.access_token;
      await onSignedIn();
    },
  });
}

function signIn() {
  document.getElementById("loginError").textContent = "";
  tokenClient.requestAccessToken({ prompt: "" });
}

function signOut() {
  if (accessToken) {
    google.accounts.oauth2.revoke(accessToken, () => {});
  }
  accessToken = null;
  ["mainView", "editView", "modeSelectView", "editDetailModal", "receiptReplaceModal"].forEach((id) => {
    document.getElementById(id).classList.add("hidden");
  });
  document.getElementById("loginView").classList.remove("hidden");
  document.getElementById("signoutBtn").classList.add("hidden");
}

async function onSignedIn() {
  document.getElementById("loginView").classList.add("hidden");
  document.getElementById("signoutBtn").classList.remove("hidden");
  toast("読み込み中…");
  try {
    await Promise.all([loadMasterData(), loadRosterData(), loadSpecialSheetNames()]);
    populateAllDropdowns();
    setDefaultDates();
    toast("準備ができました");
    document.getElementById("modeSelectView").classList.remove("hidden");
  } catch (e) {
    console.error(e);
    toast("データの読み込みに失敗しました。再読み込みしてください。", true);
  }
}

// 401対策：トークン切れなら再取得してから1回だけ再試行する
async function authedFetch(url, options = {}) {
  options.headers = Object.assign({}, options.headers, {
    Authorization: "Bearer " + accessToken,
  });
  let res = await fetch(url, options);
  if (res.status === 401) {
    accessToken = await new Promise((resolve) => {
      tokenClient.callback = (resp) => resolve(resp.access_token);
      tokenClient.requestAccessToken({ prompt: "" });
    });
    options.headers.Authorization = "Bearer " + accessToken;
    res = await fetch(url, options);
  }
  return res;
}

// --------------------------------------------------------------------
// Sheets API ヘルパー
// --------------------------------------------------------------------

const SHEETS_BASE = "https://sheets.googleapis.com/v4/spreadsheets";

// アプリが書き込んでよい列の定義。formula列はここに絶対含めない。
const SHEET_COLUMNS = {
  ledger: {
    editable: ["A","B","C","D","E","F","G","H","I","L","M"], // J,K は数式列
    formula:  ["J","K"],
  },
  special: {
    editable: ["A","B","C","D","E","F","G","H","K","L"], // I,J は数式列
    formula:  ["I","J"],
  },
};

async function sheetsGetValues(range) {
  const url = `${SHEETS_BASE}/${CONFIG.SPREADSHEET_ID}/values/${encodeURIComponent(range)}`;
  const res = await authedFetch(url);
  if (!res.ok) throw new Error("sheets get failed: " + (await res.text()));
  const data = await res.json();
  return data.values || [];
}

async function sheetsBatchUpdateValues(valueRanges) {
  const url = `${SHEETS_BASE}/${CONFIG.SPREADSHEET_ID}/values:batchUpdate`;
  const body = {
    valueInputOption: "USER_ENTERED",
    data: valueRanges, // [{range, values: [[...]]}]
  };
  const res = await authedFetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error("sheets batchUpdate failed: " + (await res.text()));
  return res.json();
}

async function getSpreadsheetSheetTitles() {
  const url = `${SHEETS_BASE}/${CONFIG.SPREADSHEET_ID}?fields=sheets.properties.title`;
  const res = await authedFetch(url);
  if (!res.ok) throw new Error("sheets meta failed: " + (await res.text()));
  const data = await res.json();
  return (data.sheets || []).map((s) => s.properties.title);
}

async function getSheetIds() {
  if (sheetIdCache) return sheetIdCache;
  const url = `${SHEETS_BASE}/${CONFIG.SPREADSHEET_ID}?fields=sheets.properties(sheetId,title)`;
  const res = await authedFetch(url);
  if (!res.ok) throw new Error("sheets meta failed: " + (await res.text()));
  const data = await res.json();
  sheetIdCache = {};
  for (const s of (data.sheets || [])) {
    sheetIdCache[s.properties.title] = s.properties.sheetId;
  }
  return sheetIdCache;
}

// 指定列(A列など)を上から読み、最初の空欄の行番号を返す
async function findNextEmptyRow(sheetName, firstRow, columnLetter = "A") {
  const values = await sheetsGetValues(`'${sheetName}'!${columnLetter}${firstRow}:${columnLetter}2000`);
  let row = firstRow;
  for (const r of values) {
    if (!r[0] || String(r[0]).trim() === "") break;
    row++;
  }
  return row;
}

// --------------------------------------------------------------------
// Drive API ヘルパー
// --------------------------------------------------------------------

async function driveFindOrCreateMonthFolder(yyyyMM) {
  const q = encodeURIComponent(
    `name='${yyyyMM}' and '${CONFIG.RECEIPT_DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`
  );
  const searchUrl = `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)`;
  const searchRes = await authedFetch(searchUrl);
  const searchData = await searchRes.json();
  if (searchData.files && searchData.files.length > 0) {
    return searchData.files[0].id;
  }
  const createRes = await authedFetch("https://www.googleapis.com/drive/v3/files", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: yyyyMM,
      mimeType: "application/vnd.google-apps.folder",
      parents: [CONFIG.RECEIPT_DRIVE_FOLDER_ID],
    }),
  });
  const createData = await createRes.json();
  return createData.id;
}

async function driveUploadReceipt(file, folderId, filename) {
  const metadata = { name: filename, parents: [folderId] };
  const boundary = "-------appboundary" + Date.now();
  const delimiter = `\r\n--${boundary}\r\n`;
  const closeDelim = `\r\n--${boundary}--`;

  const fileData = await file.arrayBuffer();
  const base64Data = arrayBufferToBase64(fileData);

  const multipartBody =
    delimiter +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    JSON.stringify(metadata) +
    delimiter +
    `Content-Type: ${file.type || "image/jpeg"}\r\n` +
    "Content-Transfer-Encoding: base64\r\n\r\n" +
    base64Data +
    closeDelim;

  const res = await authedFetch(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink",
    {
      method: "POST",
      headers: { "Content-Type": `multipart/related; boundary=${boundary}` },
      body: multipartBody,
    }
  );
  if (!res.ok) throw new Error("drive upload failed: " + (await res.text()));
  return res.json(); // {id, webViewLink}
}

function arrayBufferToBase64(buffer) {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

// --------------------------------------------------------------------
// マスタ / 名簿 / 特別会計シート一覧 読み込み
// --------------------------------------------------------------------

async function loadMasterData() {
  const rows = await sheetsGetValues(`'${CONFIG.SHEET_MASTER}'!B3:E40`);
  const income = [], expense = [], category = [], payment = [];
  for (const r of rows) {
    if (r[0]) income.push(r[0]);
    if (r[1]) expense.push(r[1]);
    if (r[2]) category.push(r[2]);
    if (r[3]) payment.push(r[3]);
  }
  masterData = { categories: category, incomeAccounts: income, expenseAccounts: expense, paymentMethods: payment };
}

async function loadRosterData() {
  const rows = await sheetsGetValues(`'${CONFIG.SHEET_ROSTER}'!B${CONFIG.ROSTER_FIRST_ROW}:B200`);
  rosterData = [];
  rows.forEach((r, i) => {
    if (r[0] && String(r[0]).trim() !== "") {
      rosterData.push({ name: r[0], row: CONFIG.ROSTER_FIRST_ROW + i });
    }
  });
}

async function loadSpecialSheetNames() {
  const titles = await getSpreadsheetSheetTitles();
  specialSheetNames = titles.filter((t) => t.includes("特別会計"));
}

// --------------------------------------------------------------------
// ドロップダウンの構築
// --------------------------------------------------------------------

function fillSelect(selectEl, options, placeholder) {
  selectEl.innerHTML = "";
  if (placeholder) {
    const o = document.createElement("option");
    o.value = "";
    o.textContent = placeholder;
    o.disabled = true;
    o.selected = true;
    selectEl.appendChild(o);
  }
  options.forEach((opt) => {
    const o = document.createElement("option");
    o.value = opt;
    o.textContent = opt;
    selectEl.appendChild(o);
  });
}

function populateAllDropdowns() {
  // 一般会計タブ
  fillSelect(document.getElementById("l_category"), masterData.categories, "選択してください");
  fillSelect(document.getElementById("l_incomeAccount"), masterData.incomeAccounts, "選択してください");
  fillSelect(document.getElementById("l_expenseAccount"), masterData.expenseAccounts, "選択してください");
  fillSelect(document.getElementById("l_payment"), masterData.paymentMethods.filter((p) => !p.includes("→")), "選択してください");

  // 部費入金タブ
  fillSelect(document.getElementById("d_member"), rosterData.map((r) => r.name), "部員を選択");
  fillSelect(document.getElementById("d_month"), CONFIG.DUES_MONTH_COLUMNS, "対象月を選択");

  // 上位大会タブ
  fillSelect(document.getElementById("s_sheet"), specialSheetNames, "大会を選択");
  fillSelect(document.getElementById("s_incomeAccount"), masterData.incomeAccounts, "選択してください");
  fillSelect(document.getElementById("s_expenseAccount"), masterData.expenseAccounts, "選択してください");
  fillSelect(document.getElementById("s_payment"), masterData.paymentMethods.filter((p) => !p.includes("→")), "選択してください");

  // 修正フォーム
  fillSelect(document.getElementById("e_category"), masterData.categories, "選択してください");
  fillSelect(document.getElementById("e_incomeAccount"), masterData.incomeAccounts, "選択してください");
  fillSelect(document.getElementById("e_expenseAccount"), masterData.expenseAccounts, "選択してください");
  fillSelect(document.getElementById("e_payment"), masterData.paymentMethods.filter((p) => !p.includes("→")), "選択してください");

  // 修正画面 大会選択
  fillSelect(document.getElementById("edit_sheet"), specialSheetNames, "大会を選択");
}

function setDefaultDates() {
  const today = new Date().toISOString().slice(0, 10);
  document.getElementById("l_date").value = today;
  document.getElementById("s_date").value = today;
}

// --------------------------------------------------------------------
// タブ切り替え
// --------------------------------------------------------------------

document.querySelectorAll(".tab-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
  });
});

// 収入／支出／振替 セグメント切り替え（一般会計・特別会計で共通の仕組み）
function wireKindSegment(prefix) {
  const seg = document.getElementById(prefix + "_kind");
  const incomeBlock = document.getElementById(prefix + "_incomeBlock");
  const expenseBlock = document.getElementById(prefix + "_expenseBlock");
  const transferBlock = document.getElementById(prefix + "_transferBlock");
  const payLabel = document.getElementById(prefix + "_payLabel");

  seg.querySelectorAll(".seg-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      seg.querySelectorAll(".seg-btn").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      const kind = btn.dataset.kind;
      incomeBlock.classList.toggle("hidden", kind !== "income");
      expenseBlock.classList.toggle("hidden", kind !== "expense");
      transferBlock.classList.toggle("hidden", kind !== "transfer");
      payLabel.classList.toggle("hidden", kind === "transfer");
    });
  });
}
wireKindSegment("l");
wireKindSegment("s");

// 部費 状態 セグメント
document.getElementById("d_status").querySelectorAll(".seg-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll("#d_status .seg-btn").forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
  });
});

function getActiveKind(prefix) {
  const active = document.querySelector(`#${prefix}_kind .seg-btn.active`);
  return active ? active.dataset.kind : "income";
}
function getActiveDuesStatus() {
  const active = document.querySelector("#d_status .seg-btn.active");
  return active ? active.dataset.status : "○";
}

document.getElementById("l_receipt").addEventListener("change", (e) => {
  const f = e.target.files[0];
  document.getElementById("l_receiptName").textContent = f ? f.name : "";
});

document.getElementById("s_receipt").addEventListener("change", (e) => {
  const f = e.target.files[0];
  document.getElementById("s_receiptName").textContent = f ? f.name : "";
});

// --------------------------------------------------------------------
// フォーム送信: 一般会計
// --------------------------------------------------------------------

document.getElementById("ledgerForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const statusEl = document.getElementById("l_status");
  statusEl.textContent = "記録しています…";

  try {
    const date = document.getElementById("l_date").value;
    const memo = document.getElementById("l_memo").value;
    const category = document.getElementById("l_category").value;
    const kind = getActiveKind("l");
    const payment = kind === "transfer"
      ? document.getElementById("l_transferDir").value
      : document.getElementById("l_payment").value;
    const note = document.getElementById("l_note").value;

    let incomeAccount = "", incomeAmount = "", expenseAccount = "", expenseAmount = "", transferAmount = "";
    let amountForFilename = "";
    let accountForFilename = category;

    if (kind === "income") {
      incomeAccount = document.getElementById("l_incomeAccount").value;
      incomeAmount = document.getElementById("l_incomeAmount").value;
      amountForFilename = incomeAmount;
      accountForFilename = incomeAccount || category;
    } else if (kind === "expense") {
      expenseAccount = document.getElementById("l_expenseAccount").value;
      expenseAmount = document.getElementById("l_expenseAmount").value;
      amountForFilename = expenseAmount;
      accountForFilename = expenseAccount || category;
    } else {
      transferAmount = document.getElementById("l_transferAmount").value;
      amountForFilename = transferAmount;
    }

    if (!date || !memo || !category) {
      statusEl.textContent = "日付・摘要・分類は必須です。";
      return;
    }

    const row = await findNextEmptyRow(CONFIG.SHEET_LEDGER, CONFIG.LEDGER_FIRST_ROW, "A");

    // 領収書アップロード（任意）
    let receiptLink = "";
    const file = document.getElementById("l_receipt").files[0];
    if (file) {
      statusEl.textContent = "領収書をアップロード中…";
      const yyyyMM = date.slice(0, 7);
      const folderId = await driveFindOrCreateMonthFolder(yyyyMM);
      const safeAccount = (accountForFilename || "未分類").replace(/[\\/:*?"<>|]/g, "_");
      const filename = `${date}_${safeAccount}_${amountForFilename || 0}.jpg`;
      const uploaded = await driveUploadReceipt(file, folderId, filename);
      receiptLink = uploaded.webViewLink || "";
    }

    statusEl.textContent = "スプレッドシートに記録中…";

    // A:I （J,K は数式なので触らない）
    const rowA_I = [date, memo, category, incomeAccount, incomeAmount, expenseAccount, expenseAmount, transferAmount, payment];
    // L:M
    const rowL_M = [receiptLink, note];

    await sheetsBatchUpdateValues([
      { range: `'${CONFIG.SHEET_LEDGER}'!A${row}:I${row}`, values: [rowA_I] },
      { range: `'${CONFIG.SHEET_LEDGER}'!L${row}:M${row}`, values: [rowL_M] },
    ]);

    statusEl.textContent = "記録しました。";
    toast("一般会計に記録しました");
    document.getElementById("ledgerForm").reset();
    setDefaultDates();
    document.getElementById("l_receiptName").textContent = "";
  } catch (err) {
    console.error(err);
    statusEl.textContent = "エラーが発生しました。もう一度お試しください。";
    toast("記録に失敗しました", true);
  }
});

// --------------------------------------------------------------------
// フォーム送信: 部費入金記録
// --------------------------------------------------------------------

document.getElementById("duesForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const statusEl = document.getElementById("d_status_msg");
  statusEl.textContent = "記録しています…";

  try {
    const memberName = document.getElementById("d_member").value;
    const month = document.getElementById("d_month").value;
    const status = getActiveDuesStatus();

    if (!memberName || !month) {
      statusEl.textContent = "部員と対象月を選択してください。";
      return;
    }

    const member = rosterData.find((r) => r.name === memberName);
    if (!member) {
      statusEl.textContent = "部員データが見つかりませんでした。";
      return;
    }
    // 部費入金管理シートの行 = 名簿の行 - 1（DUES_FIRST_ROW=4が名簿の5行目に対応）
    const duesRow = member.row - 1;
    const monthIndex = CONFIG.DUES_MONTH_COLUMNS.indexOf(month);
    const colLetter = colLetterFromIndex(colIndexFromLetter(CONFIG.DUES_MONTH_START_COL) + monthIndex);

    await sheetsBatchUpdateValues([
      { range: `'${CONFIG.SHEET_DUES}'!${colLetter}${duesRow}`, values: [[status]] },
    ]);

    statusEl.textContent = "記録しました。";
    toast(`${memberName}さん ${month} を記録しました`);
  } catch (err) {
    console.error(err);
    statusEl.textContent = "エラーが発生しました。もう一度お試しください。";
    toast("記録に失敗しました", true);
  }
});

function colIndexFromLetter(letter) {
  let idx = 0;
  for (const ch of letter) idx = idx * 26 + (ch.charCodeAt(0) - 64);
  return idx - 1; // 0-based
}
function colLetterFromIndex(idx0) {
  let idx = idx0 + 1;
  let s = "";
  while (idx > 0) {
    const rem = (idx - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    idx = Math.floor((idx - 1) / 26);
  }
  return s;
}

// --------------------------------------------------------------------
// フォーム送信: 上位大会特別会計
// --------------------------------------------------------------------

document.getElementById("specialForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const statusEl = document.getElementById("s_status");
  statusEl.textContent = "記録しています…";

  try {
    const sheetName = document.getElementById("s_sheet").value;
    const date = document.getElementById("s_date").value;
    const memo = document.getElementById("s_memo").value;
    const kind = getActiveKind("s");
    const payment = kind === "transfer"
      ? document.getElementById("s_transferDir").value
      : document.getElementById("s_payment").value;
    const note = document.getElementById("s_note").value;

    let incomeAccount = "", incomeAmount = "", expenseAccount = "", expenseAmount = "", transferAmount = "";
    let amountForFilename = "", accountForFilename = "";
    if (kind === "income") {
      incomeAccount = document.getElementById("s_incomeAccount").value;
      incomeAmount = document.getElementById("s_incomeAmount").value;
      amountForFilename = incomeAmount;
      accountForFilename = incomeAccount;
    } else if (kind === "expense") {
      expenseAccount = document.getElementById("s_expenseAccount").value;
      expenseAmount = document.getElementById("s_expenseAmount").value;
      amountForFilename = expenseAmount;
      accountForFilename = expenseAccount;
    } else {
      transferAmount = document.getElementById("s_transferAmount").value;
      amountForFilename = transferAmount;
    }

    if (!sheetName || !date || !memo) {
      statusEl.textContent = "大会・日付・摘要は必須です。";
      return;
    }

    const row = await findNextEmptyRow(sheetName, CONFIG.LEDGER_FIRST_ROW, "A");

    // 領収書アップロード（任意）
    let receiptLink = "";
    const file = document.getElementById("s_receipt").files[0];
    if (file) {
      statusEl.textContent = "領収書をアップロード中…";
      const yyyyMM = date.slice(0, 7);
      const folderId = await driveFindOrCreateMonthFolder(yyyyMM);
      const safeSheet = sheetName.replace(/[\\/:*?"<>|]/g, "_");
      const safeAccount = (accountForFilename || "未分類").replace(/[\\/:*?"<>|]/g, "_");
      const filename = `${safeSheet}_${date}_${safeAccount}_${amountForFilename || 0}.jpg`;
      const uploaded = await driveUploadReceipt(file, folderId, filename);
      receiptLink = uploaded.webViewLink || "";
    }

    statusEl.textContent = "スプレッドシートに記録中…";

    // A:H （I,J は数式なので触らない）
    const rowA_H = [date, memo, incomeAccount, incomeAmount, expenseAccount, expenseAmount, transferAmount, payment];
    // K:L（備考・領収書リンク）
    const rowK_L = [note, receiptLink];

    await sheetsBatchUpdateValues([
      { range: `'${sheetName}'!A${row}:H${row}`, values: [rowA_H] },
      { range: `'${sheetName}'!K${row}:L${row}`, values: [rowK_L] },
    ]);

    statusEl.textContent = "記録しました。";
    toast(`${sheetName}に記録しました`);
    document.getElementById("specialForm").reset();
    setDefaultDates();
    document.getElementById("s_receiptName").textContent = "";
  } catch (err) {
    console.error(err);
    statusEl.textContent = "エラーが発生しました。もう一度お試しください。";
    toast("記録に失敗しました", true);
  }
});

// --------------------------------------------------------------------
// 雑多
// --------------------------------------------------------------------

function showError(elId, msg) {
  document.getElementById(elId).textContent = msg;
}

let toastTimer = null;
function toast(msg, isError = false) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.remove("hidden");
  el.classList.toggle("error", isError);
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.add("hidden"), 2600);
}

document.getElementById("signinBtn").addEventListener("click", signIn);
document.getElementById("signoutBtn").addEventListener("click", signOut);

// ====================================================================
// 修正機能
// ====================================================================

let editCurrentType = "ledger";
let editCurrentSheet = "";
let editListItems = []; // [{rowNum, row}]
let editTargetRowNum = null;
let editTargetType = null;
let editTargetSheet = null;
let editTargetReceiptLink = "";

// モード選択
document.getElementById("modeInputBtn").addEventListener("click", () => {
  document.getElementById("modeSelectView").classList.add("hidden");
  document.getElementById("mainView").classList.remove("hidden");
});

document.getElementById("modeEditBtn").addEventListener("click", () => {
  document.getElementById("modeSelectView").classList.add("hidden");
  document.getElementById("editView").classList.remove("hidden");
  if (editCurrentType === "ledger") loadEditListData();
});

document.getElementById("mainBackBtn").addEventListener("click", backToModeSelect);
document.getElementById("editBackBtn").addEventListener("click", backToModeSelect);

function backToModeSelect() {
  ["mainView", "editView"].forEach((id) => {
    document.getElementById(id).classList.add("hidden");
  });
  document.getElementById("modeSelectView").classList.remove("hidden");
}

// 修正タブ切り替え
document.querySelectorAll(".edit-tab-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".edit-tab-btn").forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    editCurrentType = btn.dataset.etab;
    const isSpecial = editCurrentType === "special";
    document.getElementById("editSheetSelectRow").classList.toggle("hidden", !isSpecial);
    document.getElementById("editList").innerHTML = "";
    if (!isSpecial) {
      editCurrentSheet = "";
      loadEditListData();
    } else {
      editCurrentSheet = document.getElementById("edit_sheet").value;
      if (editCurrentSheet) {
        loadEditListData();
      } else {
        document.getElementById("editListStatus").textContent = "大会を選択してください。";
      }
    }
  });
});

document.getElementById("edit_sheet").addEventListener("change", () => {
  editCurrentSheet = document.getElementById("edit_sheet").value;
  if (editCurrentSheet) loadEditListData();
});

// 一覧読み込み
async function loadEditListData() {
  const statusEl = document.getElementById("editListStatus");
  const listEl = document.getElementById("editList");
  statusEl.textContent = "読み込み中…";
  listEl.innerHTML = "";

  if (editCurrentType === "special" && !editCurrentSheet) {
    statusEl.textContent = "大会を選択してください。";
    return;
  }

  try {
    const sheetName = editCurrentType === "ledger" ? CONFIG.SHEET_LEDGER : editCurrentSheet;
    const firstRow = CONFIG.LEDGER_FIRST_ROW;
    const rangeEnd = editCurrentType === "ledger" ? "M" : "L";

    const values = await sheetsGetValues(`'${sheetName}'!A${firstRow}:${rangeEnd}2000`);

    editListItems = [];
    values.forEach((row, i) => {
      if (row[0] && String(row[0]).trim() !== "") {
        editListItems.push({ rowNum: firstRow + i, row });
      }
    });

    const displayItems = [...editListItems].reverse();

    if (displayItems.length === 0) {
      statusEl.textContent = "データがありません。";
      return;
    }

    statusEl.textContent = `${displayItems.length}件`;
    renderEditList(displayItems);
  } catch (err) {
    console.error(err);
    statusEl.textContent = "読み込みに失敗しました。";
    toast("読み込みに失敗しました", true);
  }
}

function renderEditList(items) {
  const listEl = document.getElementById("editList");
  listEl.innerHTML = "";
  const type = editCurrentType;
  const sheet = editCurrentSheet;

  items.forEach(({ rowNum, row }) => {
    let date, memo, amountText, receiptLink;
    if (type === "ledger") {
      date = row[0] || "";
      memo = row[1] || "";
      receiptLink = row[11] || "";
      if (row[4]) amountText = `収入 ¥${fmtAmt(row[4])}`;
      else if (row[6]) amountText = `支出 ¥${fmtAmt(row[6])}`;
      else if (row[7]) amountText = `振替 ¥${fmtAmt(row[7])}`;
      else amountText = "";
    } else {
      date = row[0] || "";
      memo = row[1] || "";
      receiptLink = row[11] || "";
      if (row[3]) amountText = `収入 ¥${fmtAmt(row[3])}`;
      else if (row[5]) amountText = `支出 ¥${fmtAmt(row[5])}`;
      else if (row[6]) amountText = `振替 ¥${fmtAmt(row[6])}`;
      else amountText = "";
    }

    const card = document.createElement("div");
    card.className = "edit-card";
    card.innerHTML = `
      <div class="edit-card-info">
        <div class="edit-card-date">${escHtml(date)}</div>
        <div class="edit-card-memo">${escHtml(memo)}</div>
        ${amountText ? `<div class="edit-card-amount">${escHtml(amountText)}</div>` : ""}
        ${receiptLink ? `<div class="edit-card-receipt">📎 領収書あり</div>` : ""}
      </div>
      <div class="edit-card-btns">
        <button class="edit-card-btn rr-btn">📷 領収書だけ差し替え</button>
        <button class="edit-card-btn detail-btn">✏️ 詳しく修正</button>
        <button class="edit-card-btn delete-btn">🗑 削除</button>
      </div>
    `;

    card.querySelector(".rr-btn").addEventListener("click", () => openReceiptReplace(type, sheet, rowNum));
    card.querySelector(".detail-btn").addEventListener("click", () => openDetailEdit(type, sheet, rowNum));
    card.querySelector(".delete-btn").addEventListener("click", () => deleteRecord(type, sheet, rowNum));

    listEl.appendChild(card);
  });
}

function fmtAmt(val) {
  const n = parseFloat(String(val || "").replace(/,/g, ""));
  return isNaN(n) ? String(val) : n.toLocaleString("ja-JP");
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// 領収書だけ差し替え
function openReceiptReplace(type, sheet, rowNum) {
  const item = editListItems.find((i) => i.rowNum === rowNum);
  editTargetType = type;
  editTargetSheet = sheet;
  editTargetRowNum = rowNum;

  const modal = document.getElementById("receiptReplaceModal");
  modal.dataset.date = item ? (item.row[0] || "") : "";

  document.getElementById("rr_receipt").value = "";
  document.getElementById("rr_receiptName").textContent = "";
  document.getElementById("rr_status").textContent = "";
  modal.classList.remove("hidden");
}

document.getElementById("rr_receipt").addEventListener("change", (e) => {
  const f = e.target.files[0];
  document.getElementById("rr_receiptName").textContent = f ? f.name : "";
});

document.getElementById("rr_uploadBtn").addEventListener("click", async () => {
  const statusEl = document.getElementById("rr_status");
  const file = document.getElementById("rr_receipt").files[0];
  if (!file) {
    statusEl.textContent = "ファイルを選択してください。";
    return;
  }

  statusEl.textContent = "アップロード中…";
  document.getElementById("rr_uploadBtn").disabled = true;

  try {
    const date = document.getElementById("receiptReplaceModal").dataset.date
      || new Date().toISOString().slice(0, 10);
    const yyyyMM = date.slice(0, 7);
    const folderId = await driveFindOrCreateMonthFolder(yyyyMM);
    const filename = `receipt_${date}_row${editTargetRowNum}.jpg`;
    const uploaded = await driveUploadReceipt(file, folderId, filename);
    const receiptLink = uploaded.webViewLink || "";

    const range = editTargetType === "ledger"
      ? `'${CONFIG.SHEET_LEDGER}'!L${editTargetRowNum}`
      : `'${editTargetSheet}'!L${editTargetRowNum}`;

    await sheetsBatchUpdateValues([{ range, values: [[receiptLink]] }]);

    toast("領収書を更新しました");
    document.getElementById("receiptReplaceModal").classList.add("hidden");
    await loadEditListData();
  } catch (err) {
    console.error(err);
    statusEl.textContent = "エラーが発生しました。";
    toast("更新に失敗しました", true);
  } finally {
    document.getElementById("rr_uploadBtn").disabled = false;
  }
});

document.getElementById("rr_cancelBtn").addEventListener("click", () => {
  document.getElementById("receiptReplaceModal").classList.add("hidden");
});

// 詳細修正モーダル
function setEditKind(kind) {
  document.querySelectorAll("#e_kind .seg-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.kind === kind);
  });
  document.getElementById("e_incomeBlock").classList.toggle("hidden", kind !== "income");
  document.getElementById("e_expenseBlock").classList.toggle("hidden", kind !== "expense");
  document.getElementById("e_transferBlock").classList.toggle("hidden", kind !== "transfer");
  document.getElementById("e_payLabel").classList.toggle("hidden", kind === "transfer");
}

function openDetailEdit(type, sheet, rowNum) {
  const item = editListItems.find((i) => i.rowNum === rowNum);
  if (!item) return;
  const row = item.row;

  editTargetType = type;
  editTargetSheet = sheet;
  editTargetRowNum = rowNum;

  document.getElementById("e_date").value = row[0] || "";
  document.getElementById("e_memo").value = row[1] || "";
  document.getElementById("e_receipt").value = "";
  document.getElementById("e_receiptName").textContent = "";
  document.getElementById("e_status").textContent = "";

  if (type === "ledger") {
    // 仕訳入力: A=date B=memo C=category D=収入科目 E=収入金額 F=支出科目 G=支出金額 H=振替金額 I=支払手段 L=領収書 M=備考
    editTargetReceiptLink = row[11] || "";
    document.getElementById("e_categoryLabel").classList.remove("hidden");
    document.getElementById("e_category").value = row[2] || "";

    let kind;
    if (row[4]) {
      kind = "income";
      setEditKind(kind);
      document.getElementById("e_incomeAccount").value = row[3] || "";
      document.getElementById("e_incomeAmount").value = row[4] || "";
      document.getElementById("e_payment").value = row[8] || "";
    } else if (row[6]) {
      kind = "expense";
      setEditKind(kind);
      document.getElementById("e_expenseAccount").value = row[5] || "";
      document.getElementById("e_expenseAmount").value = row[6] || "";
      document.getElementById("e_payment").value = row[8] || "";
    } else {
      kind = "transfer";
      setEditKind(kind);
      document.getElementById("e_transferDir").value = row[8] || "現金→口座";
      document.getElementById("e_transferAmount").value = row[7] || "";
    }
    document.getElementById("e_note").value = row[12] || "";
  } else {
    // 上位大会: A=date B=memo C=収入科目 D=収入金額 E=支出科目 F=支出金額 G=振替金額 H=支払手段 K=備考 L=領収書
    editTargetReceiptLink = row[11] || "";
    document.getElementById("e_categoryLabel").classList.add("hidden");

    let kind;
    if (row[3]) {
      kind = "income";
      setEditKind(kind);
      document.getElementById("e_incomeAccount").value = row[2] || "";
      document.getElementById("e_incomeAmount").value = row[3] || "";
      document.getElementById("e_payment").value = row[7] || "";
    } else if (row[5]) {
      kind = "expense";
      setEditKind(kind);
      document.getElementById("e_expenseAccount").value = row[4] || "";
      document.getElementById("e_expenseAmount").value = row[5] || "";
      document.getElementById("e_payment").value = row[7] || "";
    } else {
      kind = "transfer";
      setEditKind(kind);
      document.getElementById("e_transferDir").value = row[7] || "現金→口座";
      document.getElementById("e_transferAmount").value = row[6] || "";
    }
    document.getElementById("e_note").value = row[10] || "";
  }

  document.getElementById("editDetailModal").classList.remove("hidden");
}

document.getElementById("e_receipt").addEventListener("change", (e) => {
  const f = e.target.files[0];
  document.getElementById("e_receiptName").textContent = f ? f.name : "";
});

document.getElementById("editDetailClose").addEventListener("click", () => {
  document.getElementById("editDetailModal").classList.add("hidden");
});

document.getElementById("editDetailForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const statusEl = document.getElementById("e_status");
  statusEl.textContent = "保存しています…";

  const type = editTargetType;
  const sheet = editTargetSheet;
  const rowNum = editTargetRowNum;

  try {
    const date = document.getElementById("e_date").value;
    const memo = document.getElementById("e_memo").value;
    const kind = getActiveKind("e");
    const payment = kind === "transfer"
      ? document.getElementById("e_transferDir").value
      : document.getElementById("e_payment").value;
    const note = document.getElementById("e_note").value;

    if (!date || !memo) {
      statusEl.textContent = "日付・摘要は必須です。";
      return;
    }

    let incomeAccount = "", incomeAmount = "", expenseAccount = "", expenseAmount = "", transferAmount = "";
    let amountForFilename = "", accountForFilename = "";

    if (kind === "income") {
      incomeAccount = document.getElementById("e_incomeAccount").value;
      incomeAmount = document.getElementById("e_incomeAmount").value;
      amountForFilename = incomeAmount;
      accountForFilename = incomeAccount;
    } else if (kind === "expense") {
      expenseAccount = document.getElementById("e_expenseAccount").value;
      expenseAmount = document.getElementById("e_expenseAmount").value;
      amountForFilename = expenseAmount;
      accountForFilename = expenseAccount;
    } else {
      transferAmount = document.getElementById("e_transferAmount").value;
      amountForFilename = transferAmount;
    }

    let receiptLink = editTargetReceiptLink;
    const file = document.getElementById("e_receipt").files[0];
    if (file) {
      statusEl.textContent = "領収書をアップロード中…";
      const yyyyMM = date.slice(0, 7);
      const folderId = await driveFindOrCreateMonthFolder(yyyyMM);
      const safeAccount = (accountForFilename || "未分類").replace(/[\\/:*?"<>|]/g, "_");
      const safeSheet = type === "special" ? sheet.replace(/[\\/:*?"<>|]/g, "_") + "_" : "";
      const filename = `${safeSheet}${date}_${safeAccount}_${amountForFilename || 0}.jpg`;
      const uploaded = await driveUploadReceipt(file, folderId, filename);
      receiptLink = uploaded.webViewLink || "";
    }

    statusEl.textContent = "スプレッドシートに更新中…";

    if (type === "ledger") {
      const category = document.getElementById("e_category").value;
      await sheetsBatchUpdateValues([
        { range: `'${CONFIG.SHEET_LEDGER}'!A${rowNum}:I${rowNum}`,
          values: [[date, memo, category, incomeAccount, incomeAmount, expenseAccount, expenseAmount, transferAmount, payment]] },
        { range: `'${CONFIG.SHEET_LEDGER}'!L${rowNum}:M${rowNum}`,
          values: [[receiptLink, note]] },
      ]);
    } else {
      await sheetsBatchUpdateValues([
        { range: `'${sheet}'!A${rowNum}:H${rowNum}`,
          values: [[date, memo, incomeAccount, incomeAmount, expenseAccount, expenseAmount, transferAmount, payment]] },
        { range: `'${sheet}'!K${rowNum}:L${rowNum}`,
          values: [[note, receiptLink]] },
      ]);
    }

    statusEl.textContent = "保存しました。";
    toast("修正しました");
    document.getElementById("editDetailModal").classList.add("hidden");
    await loadEditListData();
  } catch (err) {
    console.error(err);
    statusEl.textContent = "エラーが発生しました。もう一度お試しください。";
    toast("保存に失敗しました", true);
  }
});

// SHEET_COLUMNS.editable を連続列ごとにグループ化してbatchUpdate用レンジ配列を返す
function buildClearRanges(sheetRef, rowNum, cols) {
  const sorted = [...cols].sort((a, b) => colIndexFromLetter(a) - colIndexFromLetter(b));
  const groups = [];
  let start = sorted[0], prev = sorted[0];
  for (let i = 1; i < sorted.length; i++) {
    const cur = sorted[i];
    if (colIndexFromLetter(cur) === colIndexFromLetter(prev) + 1) {
      prev = cur;
    } else {
      groups.push([start, prev]);
      start = cur; prev = cur;
    }
  }
  groups.push([start, prev]);
  return groups.map(([s, e]) => ({
    range: `${sheetRef}!${s}${rowNum}:${e}${rowNum}`,
    values: [new Array(colIndexFromLetter(e) - colIndexFromLetter(s) + 1).fill("")],
  }));
}

// 数式列が空になっていたら1つ上の行からCopyPasteで復元する
async function restoreFormulaColumns(sheetName, rowNum, formulaCols) {
  if (formulaCols.length === 0 || rowNum <= CONFIG.LEDGER_FIRST_ROW) return;
  const sheetRef = `'${sheetName}'`;
  const first = formulaCols[0];
  const last = formulaCols[formulaCols.length - 1];
  const cur = await sheetsGetValues(`${sheetRef}!${first}${rowNum}:${last}${rowNum}`);
  const curRow = cur[0] || [];
  const needsRestore = formulaCols.some((_, i) => !curRow[i] || String(curRow[i]).trim() === "");
  if (!needsRestore) return;
  const sheetIds = await getSheetIds();
  const sheetId = sheetIds[sheetName];
  if (sheetId == null) return;
  const startColIdx = colIndexFromLetter(first);
  const endColIdx   = colIndexFromLetter(last) + 1;
  const url = `${SHEETS_BASE}/${CONFIG.SPREADSHEET_ID}:batchUpdate`;
  const res = await authedFetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ requests: [{ copyPaste: {
      source:      { sheetId, startRowIndex: rowNum - 2, endRowIndex: rowNum - 1, startColumnIndex: startColIdx, endColumnIndex: endColIdx },
      destination: { sheetId, startRowIndex: rowNum - 1, endRowIndex: rowNum,     startColumnIndex: startColIdx, endColumnIndex: endColIdx },
      pasteType: "PASTE_FORMULA",
      pasteOrientation: "NORMAL",
    }}]}),
  });
  if (!res.ok) throw new Error("formula restore failed: " + (await res.text()));
}

// 削除
async function deleteRecord(type, sheet, rowNum) {
  if (!confirm("この記録を削除しますか？元に戻せません。")) return;
  try {
    const sheetName = type === "ledger" ? CONFIG.SHEET_LEDGER : sheet;
    const cols = SHEET_COLUMNS[type];
    await sheetsBatchUpdateValues(buildClearRanges(`'${sheetName}'`, rowNum, cols.editable));
    await restoreFormulaColumns(sheetName, rowNum, cols.formula);
    toast("記録を削除しました");
    await loadEditListData();
  } catch (err) {
    console.error(err);
    toast("削除に失敗しました", true);
  }
}

wireKindSegment("e");

window.addEventListener("load", () => {
  initGoogleAuth();
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("service-worker.js").catch(() => {});
  }
});
