/**
 * 社内ダッシュボード - Code.gs
 *
 * SETTINGS / APP_MASTER を読み込み、Index.html にデータを渡して描画
 *
 * 期待シート:
 *  - SETTINGS: A列=項目名, B列=値
 *  - APP_MASTER: ヘッダー行=3, データ開始=4
 *
 * 公開設定（推奨）:
 *  - 実行ユーザー：自分
 *  - アクセス権：ドメイン内全員
 */

// ===== 設定 =====
const SHEET_SETTINGS = "SETTINGS";
const SHEET_APP_MASTER = "APP_MASTER";
const APP_HEADER_ROW = 3;
const APP_DATA_START_ROW = 4;

// SETTINGSのキー（A列）
const SETTINGS_KEYS = {
  DASHBOARD_TITLE: "ダッシュボード表示名",
  PORTAL_URL: "社内ポータルURL",
  DEFAULT_OPEN_MODE: "既定の開き方",
};

// icon_source 定数
const ICON_SOURCE = {
  DRIVE_FILE_ID: "drive_file_id",
  IMAGE_URL: "image_url",
  NONE: "none",
};

// open_mode 定数
const OPEN_MODE = {
  NEW_TAB: "new_tab",
  SAME_TAB: "same_tab",
};

/**
 * Webアプリ入口
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1) SETTINGS取得
    const settings = loadSettings_(ss);

    // 2) APP_MASTER取得
    const apps = loadApps_(ss, settings.defaultOpenMode);

    // 3) 最終表示データ
    const viewModel = {
      portalUrl: settings.portalUrl,
      dashboardTitle: settings.dashboardTitle,
      lastUpdated: formatDateTime_(new Date()),
      apps: apps,
      hasApps: apps.length > 0,
    };

    // 4) HTMLテンプレに渡す
    const tpl = HtmlService.createTemplateFromFile("Index");
    tpl.portalUrl = viewModel.portalUrl;
    tpl.dashboardTitle = viewModel.dashboardTitle;
    tpl.lastUpdated = viewModel.lastUpdated;
    tpl.apps = viewModel.apps;
    tpl.hasApps = viewModel.hasApps;

    return tpl.evaluate()
      .setTitle(viewModel.dashboardTitle)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    // エラー時もページを返す（真っ白にしない）
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:40px;">' +
      '<h1 style="color:#c00;">エラーが発生しました</h1>' +
      '<p>' + escapeHtml_(error.message) + '</p>' +
      '<p>管理者にお問い合わせください。</p>' +
      '</body></html>'
    ).setTitle("エラー");
  }
}

/**
 * SETTINGSシートを読み込み
 */
function loadSettings_(ss) {
  const sh = ss.getSheetByName(SHEET_SETTINGS);

  // シートが無い場合はデフォルト値を返す
  if (!sh) {
    return {
      dashboardTitle: "社内ダッシュボード",
      portalUrl: "",
      defaultOpenMode: OPEN_MODE.NEW_TAB,
      raw: {},
    };
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 1) {
    return {
      dashboardTitle: "社内ダッシュボード",
      portalUrl: "",
      defaultOpenMode: OPEN_MODE.NEW_TAB,
      raw: {},
    };
  }

  const range = sh.getRange(1, 1, Math.min(lastRow, 200), 2).getValues();

  const map = {};
  range.forEach(row => {
    const key = (row[0] || "").toString().trim();
    const val = (row[1] || "").toString().trim();
    if (key) map[key] = val;
  });

  const dashboardTitle = map[SETTINGS_KEYS.DASHBOARD_TITLE] || "社内ダッシュボード";
  const portalUrl = map[SETTINGS_KEYS.PORTAL_URL] || "";
  const defaultOpenMode = map[SETTINGS_KEYS.DEFAULT_OPEN_MODE] || OPEN_MODE.NEW_TAB;

  return { dashboardTitle, portalUrl, defaultOpenMode, raw: map };
}

/**
 * APP_MASTERを読み込み、表示対象だけ返す
 * 条件: enabled=TRUE かつ status=OK
 */
function loadApps_(ss, defaultOpenMode) {
  const sh = ss.getSheetByName(SHEET_APP_MASTER);

  // シートが無い場合は空配列
  if (!sh) {
    return [];
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < APP_DATA_START_ROW || lastCol < 1) {
    return [];
  }

  // ヘッダー取得
  const headers = sh.getRange(APP_HEADER_ROW, 1, 1, lastCol).getValues()[0]
    .map(h => (h || "").toString().trim().toLowerCase());

  const col = indexColumns_(headers);
  if (!col) {
    return [];
  }

  // データ取得
  const dataRowCount = lastRow - APP_DATA_START_ROW + 1;
  if (dataRowCount < 1) {
    return [];
  }

  const values = sh.getRange(APP_DATA_START_ROW, 1, dataRowCount, lastCol).getValues();

  const apps = [];
  values.forEach((row) => {
    const enabled = normalizeBool_(row[col.enabled]);
    const status = (row[col.status] || "").toString().trim();

    // enabled=TRUE かつ status=OK のみ
    if (!enabled) return;
    if (status !== "OK") return;

    const app = {
      sort: toNumber_(row[col.sort], 9999),
      key: (row[col.key] || "").toString().trim(),
      category: (row[col.category] || "").toString().trim(),
      label: (row[col.label] || "").toString().trim(),
      url: (row[col.url] || "").toString().trim(),
      layout: (row[col.layout] || "small").toString().trim().toLowerCase(),
      iconSource: (row[col.icon_source] || "").toString().trim().toLowerCase(),
      iconValue: (row[col.icon_value] || "").toString().trim(),
      openMode: ((row[col.open_mode] || "").toString().trim() || defaultOpenMode),
      note: (row[col.note] || "").toString().trim(),
      iconUrl: "",
    };

    // layout正規化（small/wide/tall/full以外はsmall）
    if (!["small", "wide", "tall", "full"].includes(app.layout)) {
      app.layout = "small";
    }

    // アイコンURL生成
    app.iconUrl = buildIconUrl_(app.iconSource, app.iconValue);

    // 最低限のバリデーション
    if (!app.label || !app.url) return;

    apps.push(app);
  });

  // sort昇順
  apps.sort((a, b) => (a.sort - b.sort));
  return apps;
}

/**
 * ヘッダー配列から列インデックスを返す（0-based）
 */
function indexColumns_(headers) {
  const required = [
    "enabled", "sort", "key", "category", "label", "url", "layout",
    "icon_source", "icon_value", "open_mode", "note", "status"
  ];

  const idx = {};
  let allFound = true;

  required.forEach(name => {
    const i = headers.indexOf(name);
    if (i === -1) {
      allFound = false;
    }
    idx[name] = i >= 0 ? i : 0;
  });

  // 必須列が見つからなくてもエラーにせず、見つかった列だけ使う
  return idx;
}

/**
 * icon_source と icon_value からブラウザで表示できるURLを作る
 */
function buildIconUrl_(iconSource, iconValue) {
  if (!iconSource || iconSource === ICON_SOURCE.NONE || iconSource === "none") {
    return "";
  }

  if (iconSource === ICON_SOURCE.IMAGE_URL || iconSource === "image_url") {
    return iconValue || "";
  }

  if (iconSource === ICON_SOURCE.DRIVE_FILE_ID || iconSource === "drive_file_id") {
    if (!iconValue) return "";
    // サムネイルURL形式（GAS Webアプリ内での表示に適している）
    return "https://drive.google.com/thumbnail?id=" + encodeURIComponent(iconValue) + "&sz=w800";
  }

  return "";
}

/**
 * TRUE/FALSE 等をbool化
 */
function normalizeBool_(v) {
  const s = (v || "").toString().trim().toLowerCase();
  return (s === "true" || s === "1" || s === "yes" || s === "y");
}

/**
 * 数値変換
 */
function toNumber_(v, fallback) {
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

/**
 * 日時フォーマット
 */
function formatDateTime_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
}

/**
 * HTMLエスケープ
 */
function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/**
 * HTMLファイルを分割する場合に利用（任意）
 * <?!= include('style'); ?> のように呼べる
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
