/**
 * 設定定数
 * シート名、セル位置、API URL等を一元管理
 */

const CONFIG = {
  // シート名
  SHEETS: {
    BS: "BS",
    PL: "PL",
    CR: "CR",
    DOCKET: "管理ドケット",
    TAX_CATEGORY: "区分別表",
    PL_TAX_LEDGER: "PL税務検討用元帳",
    BS_TAX_BREAKDOWN: "BS税務検討用内訳",
    FIXED_ASSETS: "固定資産台帳",
    PL_TREND: "PL推移",
    CLIENT_LIST: "クライアント一覧",
    TAX_STATUS: "税務基本ステータス",
    DELIVERABLE: "成果物"
  },
  
  // セル位置
  CELLS: {
    TIMESTAMP: "G15",
    TAX_METHOD: "F15",
    TAX_CATEGORY_TIMESTAMP: "F42",
    TAX_CATEGORY_TAX_TYPE: "E42",
    FIXED_ASSETS_TIMESTAMP: "J3"
  },
  
  // データ開始行
  DATA_START_ROW: 17,
  HEADER_ROW: 16,
  TAX_CATEGORY_DATA_START_ROW: 44,
  TAX_CATEGORY_HEADER_ROW: 43,
  FIXED_ASSETS_DATA_START_ROW: 5,
  
  // API URL
  API: {
    BASE_URL: "https://api.freee.co.jp/api/1/",
    REPORTS_URL: "https://api.freee.co.jp/api/1/reports/",
    COMPANIES: "https://api.freee.co.jp/api/1/companies",
    ACCOUNT_ITEMS: "https://api.freee.co.jp/api/1/account_items",
    TAX_CODES: "https://api.freee.co.jp/api/1/taxes/codes",
    DEALS: "https://api.freee.co.jp/api/1/deals",
    MANUAL_JOURNALS: "https://api.freee.co.jp/api/1/manual_journals",
    EXPENSE_APPLICATIONS: "https://api.freee.co.jp/api/1/expense_applications",
    JOURNALS: "https://api.freee.co.jp/api/1/journals",
    FIXED_ASSETS: "https://api.freee.co.jp/api/1/fixed_assets"
  },
  
  // ページネーション
  PAGINATION: {
    DEFAULT_LIMIT: 100,
    MAX_RETRIES: 30
  },
  
  // 対象シート（oldシート保持用）
  REFRESH_TARGET_SHEETS: ["BS", "PL", "CR", "区分別表", "PL税務検討用元帳", "BS税務検討用内訳"]
};

// BS税務検討用内訳の対象勘定科目
const BS_TAX_BREAKDOWN_CONFIG = {
  PARTNER_ACCOUNTS: ["未払金", "未払費用"],
  ITEM_ACCOUNTS: ["預り金", "長期借入金"]
};

// CSVヘッダーのエイリアス定義
const HEADER_ALIASES = {
  debitAccount: ["借方勘定科目", "借方科目", "借方勘定科目名"],
  creditAccount: ["貸方勘定科目", "貸方科目", "貸方勘定科目名"],
  debitAmount: ["借方金額"],
  creditAmount: ["貸方金額"],
  debitTax: ["借方税区分", "借方税区分名"],
  creditTax: ["貸方税区分", "貸方税区分名"],
  partner: ["取引先", "取引先名"],
  debitPartner: ["借方取引先"],
  creditPartner: ["貸方取引先"],
  item: ["品目", "品目名"],
  debitItem: ["借方品目"],
  creditItem: ["貸方品目"],
  tag: ["メモタグ", "タグ"],
  debitTag: ["借方メモタグ"],
  creditTag: ["貸方メモタグ"],
  description: ["摘要", "備考"],
  debitDescription: ["借方摘要"],
  creditDescription: ["貸方摘要"],
  transactionDate: ["取引日"]
};

// PL罫線対象カテゴリ
const PL_BORDER_TARGETS = ["売上総損益金額", "営業損益金額", "経常損益金額", "税引前当期純損益金額", "当期純損益金額"];

// BS罫線対象カテゴリ
const BS_BORDER_TARGETS = ["資産", "負債", "純資産"];
