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
  document.getElementById("mainView").classList.add("hidden");
  document.getElementById("loginView").classList.remove("hidden");
  document.getElementById("signoutBtn").classList.add("hidden");
}

async function onSignedIn() {
  document.getElementById("loginView").classList.add("hidden");
  document.getElementById("mainView").classList.remove("hidden");
  document.getElementById("signoutBtn").classList.remove("hidden");
  toast("読み込み中…");
  try {
    await Promise.all([loadMasterData(), loadRosterData(), loadSpecialSheetNames()]);
    populateAllDropdowns();
    setDefaultDates();
    toast("準備ができました");
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

window.addEventListener("load", () => {
  initGoogleAuth();
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("service-worker.js").catch(() => {});
  }
});
