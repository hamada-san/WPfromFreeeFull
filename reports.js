/**
 * 試算表・PLを取得してシートに出力（コア処理）
 */
function getTrialBalanceAndPLCore(ss, companyId, startDateStr, endDateStr) {
  // レポートコンテキストを準備
  const ctx = prepareReportContext(companyId, startDateStr, endDateStr);

  // 各シートを更新
  updateBSSheet(ss, ctx);
  updatePLSheet(ss, ctx);
  updateCRSheet(ss, ctx);

  // 仕訳数を更新
  const docketSheet = ss.getSheetByName(CONFIG.SHEETS.DOCKET);
  if (docketSheet) {
    const journalCount = getJournalCountForReport(companyId, ctx.startDateStr, ctx.endDateStr);
    docketSheet.getRange("D15").setValue(journalCount + "仕訳");
  }

  // PL月次推移
  updatePLTrendSheet(ss, ctx);

  // 区分別表は停止中・税務検討用元帳のみ復活
  const accountOrder = buildAccountOrder(ctx.bsBalances, ctx.plBalances);
  const taxCategoryResult = getTaxCategoryReportCore(ss, companyId, ctx.startDateStr, ctx.endDateStr, ctx.taxAccountingMethod, accountOrder, ctx.timestamp);

  if (taxCategoryResult && taxCategoryResult.plLedgerRows) {
    writePLTaxLedgerSheet(ss, taxCategoryResult.plLedgerRows, ctx.timestamp);
  }

  // BS税務検討用内訳
  const accountItemCategoryMap = getAccountItemCategoryMap(companyId);
  try {
    getBSTaxBreakdownCore(ss, companyId, ctx.startDateStr, ctx.endDateStr, accountItemCategoryMap, ctx.timestamp);
  } catch (e) {
    Logger.log("BS税務検討用内訳取得エラー: " + e.message);
  }

  // 固定資産台帳
  try {
    getFixedAssetsCore(ss, companyId, ctx.fiscalYear, ctx.startDateStr);
  } catch (e) {
    Logger.log("固定資産台帳取得エラー: " + e.message);
  }
}

/**
 * API呼び出し用のオプションを取得
 */
function getApiOptions() {
  const service = getService();
  if (!service.hasAccess()) {
    throw new Error("認証されていません。メニューから認証を行ってください。");
  }
  return {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };
}

/**
 * レポートコンテキストを準備（会社情報・年度情報・ラベル等を一括取得）
 */
function prepareReportContext(companyId, startDateStr, endDateStr) {
  const options = getApiOptions();

  // 会社情報を取得（1回のみ）
  const companyUrl = CONFIG.API.COMPANIES + "/" + companyId;
  const companyResponse = UrlFetchApp.fetch(companyUrl, options);
  const companyInfo = JSON.parse(companyResponse.getContentText());
  const fiscalYears = companyInfo.company.fiscal_years;

  if (!fiscalYears || fiscalYears.length === 0) {
    throw new Error("会計年度情報が取得できません。");
  }

  // 対象会計年度を特定
  let targetFiscalYear = null;
  if (startDateStr && endDateStr) {
    targetFiscalYear = fiscalYears.find(fy => fy.start_date === startDateStr && fy.end_date === endDateStr);
    if (!targetFiscalYear) {
      const endDate = new Date(endDateStr);
      targetFiscalYear = fiscalYears.find(fy => {
        const start = new Date(fy.start_date);
        const end = new Date(fy.end_date);
        return start <= endDate && endDate <= end;
      });
    }
    if (!targetFiscalYear) {
      throw new Error("指定した事業年度が見つかりません。");
    }
  } else {
    targetFiscalYear = fiscalYears[fiscalYears.length - 1];
    startDateStr = targetFiscalYear.start_date;
    endDateStr = targetFiscalYear.end_date;
  }

  const fiscalYear = parseInt(startDateStr.substring(0, 4), 10);
  const startMonth = parseInt(startDateStr.substring(5, 7), 10);
  const endMonth = parseInt(endDateStr.substring(5, 7), 10);
  const endYear = parseInt(endDateStr.substring(0, 4), 10);
  const startYear = parseInt(startDateStr.substring(0, 4), 10);

  // 期ラベル生成
  const periodLabels = {
    current: formatPeriodLabel(startYear, startMonth, endYear, endMonth),
    previous: formatPeriodLabel(startYear - 1, startMonth, endYear - 1, endMonth),
    twoYearsAgo: formatPeriodLabel(startYear - 2, startMonth, endYear - 2, endMonth)
  };

  // 税区分
  let taxAccountingMethod = "";
  if (companyInfo.company.tax_at_source_calc_type === 1) {
    taxAccountingMethod = "税抜経理";
  } else if (companyInfo.company.tax_at_source_calc_type === 0) {
    taxAccountingMethod = "税込経理";
  }

  const timestamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm") + "更新";

  // BS・PLデータを事前取得（再利用のため）
  const baseUrl = CONFIG.API.REPORTS_URL;
  const bsUrl = baseUrl + "trial_bs?company_id=" + companyId + "&start_date=" + startDateStr + "&end_date=" + endDateStr;
  const bsRes = UrlFetchApp.fetch(bsUrl, options);
  const bsBalances = JSON.parse(bsRes.getContentText()).trial_bs.balances;

  const params = "?company_id=" + companyId + "&fiscal_year=" + fiscalYear + "&start_month=" + startMonth + "&end_month=" + endMonth;
  const plUrl = baseUrl + "trial_pl_three_years" + params;
  const plRes = UrlFetchApp.fetch(plUrl, options);
  const plBalances = JSON.parse(plRes.getContentText()).trial_pl_three_years.balances;

  return {
    companyId: companyId,
    companyInfo: companyInfo.company,
    startDateStr: startDateStr,
    endDateStr: endDateStr,
    fiscalYear: fiscalYear,
    startMonth: startMonth,
    endMonth: endMonth,
    endYear: endYear,
    startYear: startYear,
    periodLabels: periodLabels,
    taxAccountingMethod: taxAccountingMethod,
    timestamp: timestamp,
    options: options,
    bsBalances: bsBalances,
    plBalances: plBalances
  };
}

/**
 * 期ラベルをフォーマット（YY/MM-YY/MM形式）
 */
function formatPeriodLabel(startY, startM, endY, endM) {
  const sy = String(startY).slice(-2);
  const ey = String(endY).slice(-2);
  const sm = String(startM).padStart(2, '0');
  const em = String(endM).padStart(2, '0');
  return sy + "/" + sm + "-" + ey + "/" + em;
}

/**
 * BSシートを更新
 */
function updateBSSheet(ss, ctx) {
  const bsSheet = ss.getSheetByName(CONFIG.SHEETS.BS);
  if (!bsSheet) return;

  const startRow = CONFIG.DATA_START_ROW;
  const bsLastRow = bsSheet.getLastRow();
  if (bsLastRow >= startRow) {
    bsSheet.getRange(startRow, 2, bsLastRow - startRow + 1, 8).clearContent();
    bsSheet.getRange(startRow, 2, bsLastRow - startRow + 1, 8).setBorder(false, false, false, false, false, false);
    bsSheet.getRange(startRow, 2, bsLastRow - startRow + 1, 8).setBackground(null);
  }

  bsSheet.getRange(CONFIG.CELLS.TIMESTAMP).setValue(ctx.timestamp);

  const bsHeaders = ["分類", "勘定科目", "期首残高", "借方金額", "貸方金額", "期末残高", "", "", "プリペアコメント", "基礎資料又はfreee仕訳リンク", "レビュワーコメント"];
  bsSheet.getRange(CONFIG.HEADER_ROW, 2, 1, 8).setValues([bsHeaders.slice(0, 8)]);

  const bsRows = ctx.bsBalances.map(i => {
    let category = "";
    let accountName = "";
    const hierarchyLevel = i.hierarchy_level || 0;
    const indent = "　".repeat(hierarchyLevel);

    if (i.account_item_name) {
      accountName = i.account_item_name;
    } else if (i.account_category_name) {
      category = indent + i.account_category_name;
    } else if (i.parent_account_category_name) {
      category = "▼" + i.parent_account_category_name;
    }

    return [category, accountName, Number(i.opening_balance) || 0, Number(i.debit_amount) || 0, Number(i.credit_amount) || 0, Number(i.closing_balance) || 0];
  });

  if (bsRows.length > 0) {
    bsSheet.getRange(startRow, 2, bsRows.length, 6).setValues(bsRows);
    bsSheet.getRange(startRow, 7, bsRows.length, 1).setBackground("#f4cccc");
    // 金額列（D〜G列）を3桁区切り・小数点なしに設定
    bsSheet.getRange(startRow, 4, bsRows.length, 4).setNumberFormat("#,##0");

    for (let i = 0; i < bsRows.length; i++) {
      const categoryName = bsRows[i][0].replace(/^▼/, "").trim();
      if (BS_BORDER_TARGETS.includes(categoryName)) {
        bsSheet.getRange(startRow + i, 2, 1, 8).setBorder(true, false, false, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      }
    }
  }
}

/**
 * PLシートを更新
 */
function updatePLSheet(ss, ctx) {
  const plSheet = ss.getSheetByName(CONFIG.SHEETS.PL);
  if (!plSheet) return;

  const startRow = CONFIG.DATA_START_ROW;
  const plLastRow = plSheet.getLastRow();
  if (plLastRow >= startRow) {
    plSheet.getRange(startRow, 2, plLastRow - startRow + 1, 8).clearContent();
    plSheet.getRange(startRow, 2, plLastRow - startRow + 1, 8).setBorder(false, false, false, false, false, false);
    plSheet.getRange(startRow, 2, plLastRow - startRow + 1, 8).setBackground(null);
  }

  plSheet.getRange(CONFIG.CELLS.TIMESTAMP).setValue(ctx.timestamp);
  plSheet.getRange(CONFIG.CELLS.TAX_METHOD).setValue(ctx.taxAccountingMethod);

  const plHeaders = ["分類", "勘定科目", ctx.periodLabels.twoYearsAgo, ctx.periodLabels.previous, ctx.periodLabels.current, "前年差額", "前年比", "", "プリペアコメント", "基礎資料又はfreee仕訳リンク", "レビュワーコメント"];
  plSheet.getRange(CONFIG.HEADER_ROW, 2, 1, 8).setValues([plHeaders.slice(0, 8)]);

  const plRows = ctx.plBalances.map(item => {
    const currClosing = Number(item.closing_balance) || 0;
    const prevClosing = Number(item.last_year_closing_balance) || 0;
    const twoYearsAgoClosing = Number(item.two_years_before_closing_balance) || 0;
    const difference = currClosing - prevClosing;
    const ratio = prevClosing !== 0 ? Math.round((currClosing / prevClosing) * 10000) / 100 : null;

    let category = "";
    let accountName = "";

    if (item.account_item_name) {
      accountName = item.account_item_name;
    } else if (item.account_category_name) {
      category = item.account_category_name;
    } else if (item.parent_account_category_name) {
      category = item.parent_account_category_name;
    }

    return [category, accountName, twoYearsAgoClosing, prevClosing, currClosing, difference, ratio !== null ? `${ratio}%` : "N/A"];
  });

  if (plRows.length > 0) {
    plSheet.getRange(startRow, 2, plRows.length, 7).setValues(plRows);
    plSheet.getRange(startRow, 6, plRows.length, 1).setBackground("#f4cccc");
    // 金額列（D〜G列）を3桁区切り・小数点なしに設定
    plSheet.getRange(startRow, 4, plRows.length, 4).setNumberFormat("#,##0");

    for (let i = 0; i < plRows.length; i++) {
      if (PL_BORDER_TARGETS.includes(plRows[i][0].trim())) {
        plSheet.getRange(startRow + i, 2, 1, 8).setBorder(true, false, false, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      }
    }
  }
}

/**
 * CRシートを更新
 */
function updateCRSheet(ss, ctx) {
  const crSheet = ss.getSheetByName(CONFIG.SHEETS.CR);
  if (!crSheet) return;

  const startRow = CONFIG.DATA_START_ROW;
  const crLastRow = crSheet.getLastRow();
  if (crLastRow >= startRow) {
    crSheet.getRange(startRow, 2, crLastRow - startRow + 1, 8).clearContent();
    crSheet.getRange(startRow, 2, crLastRow - startRow + 1, 8).setBorder(false, false, false, false, false, false);
    crSheet.getRange(startRow, 2, crLastRow - startRow + 1, 8).setBackground(null);
  }

  crSheet.getRange(CONFIG.CELLS.TIMESTAMP).setValue(ctx.timestamp);

  try {
    const params = "?company_id=" + ctx.companyId + "&fiscal_year=" + ctx.fiscalYear + "&start_month=" + ctx.startMonth + "&end_month=" + ctx.endMonth;
    const crUrl = CONFIG.API.REPORTS_URL + "trial_cr_three_years" + params;
    const crRes = UrlFetchApp.fetch(crUrl, ctx.options);

    if (crRes.getResponseCode() === 200) {
      const crData = JSON.parse(crRes.getContentText());
      const crBalances = crData.trial_cr_three_years?.balances;

      if (crBalances && crBalances.length > 0) {
        crSheet.getRange("B16").setValue("製造原価報告書");

        const crHeaders = ["分類", "勘定科目", ctx.periodLabels.twoYearsAgo, ctx.periodLabels.previous, ctx.periodLabels.current, "前年差額", "前年比", "", "プリペアコメント", "基礎資料又はfreee仕訳リンク", "レビュワーコメント"];
        crSheet.getRange(CONFIG.HEADER_ROW, 2, 1, 8).setValues([crHeaders.slice(0, 8)]);

        const crRows = crBalances.map(item => {
          const currClosing = Number(item.closing_balance) || 0;
          const prevClosing = Number(item.last_year_closing_balance) || 0;
          const twoYearsAgoClosing = Number(item.two_years_before_closing_balance) || 0;
          const difference = currClosing - prevClosing;
          const ratio = prevClosing !== 0 ? Math.round((currClosing / prevClosing) * 10000) / 100 : null;

          let category = "";
          let accountName = "";

          if (item.account_item_name) {
            accountName = item.account_item_name;
          } else if (item.account_category_name) {
            category = item.account_category_name;
          } else if (item.parent_account_category_name) {
            category = item.parent_account_category_name;
          }

          return [category, accountName, twoYearsAgoClosing, prevClosing, currClosing, difference, ratio !== null ? `${ratio}%` : "N/A"];
        });

        if (crRows.length > 0) {
          crSheet.getRange(startRow, 2, crRows.length, 7).setValues(crRows);
          crSheet.getRange(startRow, 6, crRows.length, 1).setBackground("#f4cccc");
          // 金額列（D〜G列）を3桁区切り・小数点なしに設定
          crSheet.getRange(startRow, 4, crRows.length, 4).setNumberFormat("#,##0");
        }
      } else {
        crSheet.getRange("B17").setValue("製造原価報告書なし");
      }
    } else {
      crSheet.getRange("B17").setValue("製造原価報告書なし");
    }
  } catch (e) {
    crSheet.getRange("B17").setValue("製造原価報告書なし");
    Logger.log("CR取得エラー: " + e.message);
  }
}

/**
 * PL推移シートを更新（月次推移）
 */
function updatePLTrendSheet(ss, ctx) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PL_TREND);
  if (!sheet) return;

  // 対象月リストを生成（期首〜期末）
  const months = [];
  let y = ctx.startYear;
  let m = ctx.startMonth;
  while (y < ctx.endYear || (y === ctx.endYear && m <= ctx.endMonth)) {
    months.push({ year: y, month: m });
    m++;
    if (m > 12) { m = 1; y++; }
  }

  // 各月の累計PLを取得（期首〜各月末）
  // freee APIは start_month に関わらず期首からの累計を返すため、
  // 当月発生額 = 当月末累計 - 前月末累計 で計算する
  const cumulativeBalances = months.map(({ month }) => {
    const params = "?company_id=" + ctx.companyId +
      "&fiscal_year=" + ctx.fiscalYear +
      "&start_month=" + ctx.startMonth +
      "&end_month=" + month;
    const url = CONFIG.API.REPORTS_URL + "trial_pl" + params;
    Utilities.sleep(500);
    const res = UrlFetchApp.fetch(url, ctx.options);
    if (res.getResponseCode() !== 200) return [];
    return JSON.parse(res.getContentText()).trial_pl?.balances || [];
  });

  // 行の並び順をplBalancesから決定
  const rowKeys = [];
  const rowMeta = {};
  ctx.plBalances.forEach(item => {
    let label = "";
    let isCategory = false;
    if (item.account_item_name) {
      label = item.account_item_name;
    } else if (item.account_category_name) {
      label = item.account_category_name;
      isCategory = true;
    } else if (item.parent_account_category_name) {
      label = item.parent_account_category_name;
      isCategory = true;
    }
    if (label && !rowMeta[label]) {
      rowKeys.push(label);
      rowMeta[label] = { isCategory };
    }
  });

  // 累計マップを構築（label -> monthIndex -> 累計金額）
  const cumulativeByLabel = {};
  cumulativeBalances.forEach((balances, mi) => {
    balances.forEach(item => {
      let label = "";
      if (item.account_item_name) label = item.account_item_name;
      else if (item.account_category_name) label = item.account_category_name;
      else if (item.parent_account_category_name) label = item.parent_account_category_name;
      if (!label) return;
      if (!cumulativeByLabel[label]) cumulativeByLabel[label] = {};
      cumulativeByLabel[label][mi] = Number(item.closing_balance) || 0;
    });
  });

  // 月次発生額 = 当月末累計 - 前月末累計
  const amountByLabelAndMonth = {};
  rowKeys.forEach(label => {
    amountByLabelAndMonth[label] = {};
    months.forEach((_, mi) => {
      const curr = (cumulativeByLabel[label] && cumulativeByLabel[label][mi] !== undefined)
        ? cumulativeByLabel[label][mi] : 0;
      const prev = (mi > 0 && cumulativeByLabel[label] && cumulativeByLabel[label][mi - 1] !== undefined)
        ? cumulativeByLabel[label][mi - 1] : 0;
      amountByLabelAndMonth[label][mi] = curr - prev;
    });
  });

  // シートをクリア
  const headerRow = CONFIG.HEADER_ROW;
  const startRow = CONFIG.DATA_START_ROW;
  const totalCols = 2 + months.length + 1; // 分類+勘定科目+月数+合計列
  const lastRow = sheet.getLastRow();
  if (lastRow >= headerRow) {
    sheet.getRange(headerRow, 2, lastRow - headerRow + 1, totalCols).clearContent();
    sheet.getRange(headerRow, 2, lastRow - headerRow + 1, totalCols).setBorder(false, false, false, false, false, false);
  }

  // タイムスタンプ
  sheet.getRange(CONFIG.CELLS.TIMESTAMP).setValue(ctx.timestamp);

  // ヘッダー行（B16）
  const monthLabels = months.map(({ year, month }) => {
    return String(year).slice(-2) + "/" + String(month).padStart(2, "0");
  });
  const headers = ["分類", "勘定科目"].concat(monthLabels).concat(["合計"]);
  sheet.getRange(headerRow, 2, 1, headers.length).setValues([headers]);

  // データ行（B17〜）
  const dataRows = rowKeys.map(label => {
    const meta = rowMeta[label];
    const category = meta.isCategory ? label : "";
    const accountName = meta.isCategory ? "" : label;
    const amounts = months.map((_, mi) => {
      return (amountByLabelAndMonth[label] && amountByLabelAndMonth[label][mi] !== undefined)
        ? amountByLabelAndMonth[label][mi]
        : 0;
    });
    const total = amounts.reduce((sum, v) => sum + (Number(v) || 0), 0);
    return [category, accountName].concat(amounts).concat([total]);
  });

  if (dataRows.length > 0) {
    sheet.getRange(startRow, 2, dataRows.length, headers.length).setValues(dataRows);
    // 金額列（月データ列＋合計列）を3桁区切り
    if (months.length > 0) {
      sheet.getRange(startRow, 4, dataRows.length, months.length + 1).setNumberFormat("#,##0");
    }
    // PL小計行に罫線
    for (let i = 0; i < dataRows.length; i++) {
      const categoryLabel = (dataRows[i][0] || "").trim();
      if (PL_BORDER_TARGETS.includes(categoryLabel)) {
        sheet.getRange(startRow + i, 2, 1, headers.length).setBorder(true, false, false, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      }
    }
  }
}

/**
 * 勘定科目名 -> 大分類のマップを作成
 */
function getAccountItemCategoryMap(companyId) {
  const service = getService();
  if (!service.hasAccess()) {
    return {};
  }

  const options = getApiOptions();
  const url = CONFIG.API.ACCOUNT_ITEMS + "?company_id=" + companyId;
  const res = UrlFetchApp.fetch(url, options);
  const items = JSON.parse(res.getContentText()).account_items || [];
  const map = {};
  items.forEach(item => {
    const categories = item.categories || [];
    const key = normalizeText(item.name);
    map[key] = categories.length > 0 ? categories[0] : (item.account_category || "");
  });
  return map;
}

/**
 * ヘッダーインデックスマップを構築
 */
function buildHeaderIndexMap(headers) {
  const map = {};
  for (const [key, candidates] of Object.entries(HEADER_ALIASES)) {
    map[key] = findHeaderIndex(headers, candidates);
  }
  return map;
}

/**
 * 仕訳帳CSVを取得
 */
function getJournalsCsvRows(companyId, startDate, endDate) {
  const options = getApiOptions();

  const url = CONFIG.API.JOURNALS +
    "?company_id=" + companyId +
    "&start_date=" + startDate +
    "&end_date=" + endDate +
    "&download_type=generic" +
    "&encoding=sjis";

  const res = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(res.getContentText()).journals;
  if (!data || !data.status_url) {
    throw new Error("仕訳帳のステータスURLが取得できません。");
  }

  const statusUrl = appendQueryParam(data.status_url, "company_id", companyId);
  let downloadUrl = "";
  const maxTries = CONFIG.PAGINATION.MAX_RETRIES;

  for (let i = 0; i < maxTries; i++) {
    const statusRes = UrlFetchApp.fetch(statusUrl, options);
    const statusData = JSON.parse(statusRes.getContentText()).journals;
    if (statusData.status === "uploaded" && statusData.download_url) {
      downloadUrl = appendQueryParam(statusData.download_url, "company_id", companyId);
      break;
    }
    if (statusData.status === "failed") {
      throw new Error("仕訳帳の生成に失敗しました。");
    }
    Utilities.sleep(2000);
  }

  if (!downloadUrl) {
    throw new Error("仕訳帳CSVの生成がタイムアウトしました。");
  }

  const csvRes = UrlFetchApp.fetch(downloadUrl, options);
  let csvText = csvRes.getContentText("Shift_JIS");
  if (csvText.charCodeAt(0) === 0xFEFF) {
    csvText = csvText.slice(1);
  }
  let rows = Utilities.parseCsv(csvText, "\t");
  if (!rows || rows.length === 0) {
    return rows;
  }
  const isHeaderRow = (row) => {
    if (!row) return false;
    const hasDate = row.some(v => normalizeText(v) === "取引日");
    const hasDebit = row.some(v => normalizeText(v) === "借方勘定科目");
    return hasDate && hasDebit;
  };
  const headerIndex = rows.findIndex(isHeaderRow);
  if (headerIndex > 0) {
    rows = rows.slice(headerIndex);
  }
  const headers = rows[0];
  if (rows.length > 1) {
    const first = rows[1];
    const isNumber = v => /^\d+$/.test(String(v || "").trim());
    const isDate = v => /^\d{4}\/\d{1,2}\/\d{1,2}$/.test(String(v || "").trim());
    if (normalizeText(headers[0]) === "取引日" && isNumber(first[0]) && isDate(first[1])) {
      rows[0] = ["行番号"].concat(headers);
      return rows;
    }
    if (normalizeText(headers[0]) === "" && normalizeText(headers[1]) === "取引日" && isNumber(first[0]) && isDate(first[1])) {
      rows[0][0] = "行番号";
      return rows;
    }
  }
  return rows;
}

function appendQueryParam(url, key, value) {
  const separator = url.indexOf("?") >= 0 ? "&" : "?";
  return url + separator + encodeURIComponent(key) + "=" + encodeURIComponent(value);
}

function findHeaderIndex(headers, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const target = candidates[i];
    for (let j = 0; j < headers.length; j++) {
      const normalized = normalizeText(headers[j]);
      if (normalized === target) {
        return j;
      }
      if (normalized.includes(target)) {
        return j;
      }
    }
  }
  return -1;
}

function parseAmount(value) {
  if (value === null || value === undefined) return 0;
  const s = String(value).replace(/,/g, "").trim();
  if (s === "") return 0;
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function normalizeText(value) {
  return String(value || "").replace(/\u3000/g, " ").trim();
}

function matchesTargetAccount(accountName, targets) {
  if (!accountName) return false;
  for (let i = 0; i < targets.length; i++) {
    if (accountName === targets[i]) return true;
    if (accountName.includes(targets[i])) return true;
  }
  return false;
}

function getCategoryForAccount(accountName, accountItemCategoryMap) {
  if (!accountName) return "";
  if (accountItemCategoryMap[accountName]) return accountItemCategoryMap[accountName];
  const keys = Object.keys(accountItemCategoryMap);
  for (let i = 0; i < keys.length; i++) {
    if (accountName === keys[i] || accountName.includes(keys[i])) {
      return accountItemCategoryMap[keys[i]];
    }
  }
  return "";
}

function getNameMapFromEndpoint(resource, companyId) {
  const service = getService();
  if (!service.hasAccess()) {
    return { error: "no_access" };
  }

  const options = getApiOptions();
  const map = {};
  const limit = CONFIG.PAGINATION.DEFAULT_LIMIT;
  let offset = 0;
  let key = "";
  if (resource === "partners") key = "partners";
  if (resource === "items") key = "items";
  if (resource === "tags") key = "tags";
  if (!key) return map;

  while (true) {
    const url = CONFIG.API.BASE_URL + resource +
      "?company_id=" + companyId +
      "&limit=" + limit +
      "&offset=" + offset;
    const res = UrlFetchApp.fetch(url, options);

    if (res.getResponseCode() !== 200) {
      map._error = res.getResponseCode() + ": " + res.getContentText().substring(0, 100);
      return map;
    }

    const data = JSON.parse(res.getContentText());
    const list = data[key] || [];
    list.forEach(item => {
      map[item.id] = item.name || "";
    });
    if (list.length < limit) break;
    offset += limit;
  }
  return map;
}

function getTagNames(tagIds, tagMap) {
  if (!tagIds || tagIds.length === 0) return "";
  const names = tagIds.map(id => tagMap[id]).filter(name => name);
  return names.join("、");
}

function writePLTaxLedgerSheet(ss, rows, timestamp) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PL_TAX_LEDGER);
  if (!sheet) {
    return;
  }
  if (timestamp) {
    sheet.getRange(CONFIG.CELLS.TIMESTAMP).setValue(timestamp);
  }

  const startRow = CONFIG.DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    // B〜K列（10列）をクリア
    sheet.getRange(startRow, 2, lastRow - startRow + 1, 10).clearContent();
    sheet.getRange(startRow, 2, lastRow - startRow + 1, 10).setBorder(false, false, false, false, false, false);
  }

  // ヘッダー: 日付, 勘定科目, 税区分, 取引先タグ, 品目タグ, メモタグ, 摘要, 取引内容, 借方金額, 貸方金額
  const headersRow = ["日付", "勘定科目", "税区分", "取引先タグ", "品目タグ", "メモタグ", "摘要", "取引内容", "借方金額", "貸方金額"];
  sheet.getRange(CONFIG.HEADER_ROW, 2, 1, headersRow.length).setValues([headersRow]);

  if (rows && rows.length > 0) {
    sheet.getRange(startRow, 2, rows.length, 10).setValues(rows);
    // 金額列（J〜K列）を3桁区切り・小数点なしに設定
    sheet.getRange(startRow, 10, rows.length, 2).setNumberFormat("#,##0");

    // 小計行と勘定科目の境界に罫線を引く
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      // 小計行の場合、下に罫線
      if (row[7] === "小計") {
        sheet.getRange(startRow + i, 2, 1, 10).setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      }
    }
  }
}

/**
 * PL税務検討用元帳を出力
 */
function getPLTaxLedgerCore(ss, companyId, startDate, endDate, accountItemCategoryMap, timestamp) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PL_TAX_LEDGER);
  if (!sheet) {
    return;
  }
  if (timestamp) {
    sheet.getRange(CONFIG.CELLS.TIMESTAMP).setValue(timestamp);
  }

  const targetAccounts = ["雑収入", "雑損失", "固定資産売却益", "固定資産売却損"]
    .map(name => normalizeText(name));
  const rows = getJournalsCsvRows(companyId, startDate, endDate);
  if (!rows || rows.length === 0) {
    return;
  }

  const headers = rows[0];
  const headerMap = buildHeaderIndexMap(headers);

  const debitAccountIdx = headerMap.debitAccount;
  const creditAccountIdx = headerMap.creditAccount;
  const debitAmountIdx = headerMap.debitAmount;
  const creditAmountIdx = headerMap.creditAmount;
  const debitTaxIdx = headerMap.debitTax;
  const creditTaxIdx = headerMap.creditTax;
  const partnerIdx = headerMap.partner;
  const debitPartnerIdx = headerMap.debitPartner;
  const creditPartnerIdx = headerMap.creditPartner;
  const itemIdx = headerMap.item;
  const debitItemIdx = headerMap.debitItem;
  const creditItemIdx = headerMap.creditItem;
  const tagIdx = headerMap.tag;
  const debitTagIdx = headerMap.debitTag;
  const creditTagIdx = headerMap.creditTag;
  const descIdx = headerMap.description;
  const debitDescIdx = headerMap.debitDescription;
  const creditDescIdx = headerMap.creditDescription;

  const output = [];

  const getVal = (row, specificIdx, commonIdx) => {
    if (specificIdx >= 0 && row[specificIdx]) return normalizeText(row[specificIdx]);
    if (commonIdx >= 0 && row[commonIdx]) return normalizeText(row[commonIdx]);
    return "";
  };

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const debitAccount = normalizeText(debitAccountIdx >= 0 ? row[debitAccountIdx] : "");
    const creditAccount = normalizeText(creditAccountIdx >= 0 ? row[creditAccountIdx] : "");
    const debitAmount = parseAmount(debitAmountIdx >= 0 ? row[debitAmountIdx] : 0);
    const creditAmount = parseAmount(creditAmountIdx >= 0 ? row[creditAmountIdx] : 0);

    if (debitAccount && matchesTargetAccount(debitAccount, targetAccounts) && debitAmount !== 0) {
      output.push([
        getCategoryForAccount(debitAccount, accountItemCategoryMap),
        debitAccount,
        getVal(row, debitTaxIdx, -1),
        getVal(row, debitPartnerIdx, partnerIdx),
        getVal(row, debitItemIdx, itemIdx),
        getVal(row, debitTagIdx, tagIdx),
        getVal(row, debitDescIdx, descIdx),
        debitAmount,
        0
      ]);
    }
    if (creditAccount && matchesTargetAccount(creditAccount, targetAccounts) && creditAmount !== 0) {
      output.push([
        getCategoryForAccount(creditAccount, accountItemCategoryMap),
        creditAccount,
        getVal(row, creditTaxIdx, -1),
        getVal(row, creditPartnerIdx, partnerIdx),
        getVal(row, creditItemIdx, itemIdx),
        getVal(row, creditTagIdx, tagIdx),
        getVal(row, creditDescIdx, descIdx),
        0,
        creditAmount
      ]);
    }
  }

  const startRow = CONFIG.DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 2, lastRow - startRow + 1, 9).clearContent();
  }

  const headersRow = ["大分類", "勘定科目", "税区分", "取引先タグ", "品目タグ", "メモタグ", "摘要", "借方金額", "貸方金額"];
  sheet.getRange(CONFIG.HEADER_ROW, 2, 1, headersRow.length).setValues([headersRow]);

  if (output.length > 0) {
    sheet.getRange(startRow, 2, output.length, 9).setValues(output);
  }
}

/**
 * BS税務検討用内訳を出力
 */
function getBSTaxBreakdownCore(ss, companyId, startDate, endDate, accountItemCategoryMap, timestamp) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.BS_TAX_BREAKDOWN);
  if (!sheet) {
    return;
  }
  if (timestamp) {
    sheet.getRange(CONFIG.CELLS.TIMESTAMP).setValue(timestamp);
  }

  // G9:G14から対象勘定科目、H9:H14からタグ種類を取得
  const configRange = sheet.getRange("G9:H14").getValues();
  const targetConfigs = configRange
    .map(row => ({
      accountName: String(row[0] || "").trim(),
      tagType: String(row[1] || "").trim()
    }))
    .filter(config => config.accountName !== "");

  if (targetConfigs.length === 0) {
    return;
  }

  const output = [];
  const subtotalRowOffsets = [];
  const pushEntriesWithSubtotal = (accountName, entries) => {
    if (!entries || entries.length === 0) return;
    const normalizedAccountName = normalizeText(accountName);
    const category = accountItemCategoryMap[normalizedAccountName] || "";
    let sumOpening = 0;
    let sumDebit = 0;
    let sumCredit = 0;
    let sumClosing = 0;
    entries.forEach(entry => {
      sumOpening += entry.opening_balance || 0;
      sumDebit += entry.debit_amount || 0;
      sumCredit += entry.credit_amount || 0;
      sumClosing += entry.closing_balance || 0;
      output.push([
        category,
        accountName,
        entry.name || "",
        entry.opening_balance || 0,
        entry.debit_amount || 0,
        entry.credit_amount || 0,
        entry.closing_balance || 0
      ]);
    });
    subtotalRowOffsets.push(output.length + 1);
    output.push([
      category,
      accountName,
      "小計",
      sumOpening,
      sumDebit,
      sumCredit,
      sumClosing
    ]);
    output.push(["", "", "", "", "", "", ""]);
  };

  // 取引先内訳と品目内訳を両方取得
  const partnerBalances = getTrialBSBreakdown(companyId, startDate, endDate, "partner");
  const itemBalances = getTrialBSBreakdown(companyId, startDate, endDate, "item");

  // シートの並び順通りに処理（H列のタグ種類に基づいて内訳を選択）
  targetConfigs.forEach(config => {
    const normalizedName = normalizeText(config.accountName);
    const tagType = config.tagType;

    // H列のタグ種類に基づいて内訳を選択
    if (tagType === "取引先") {
      // 取引先内訳を検索
      const partnerBalance = partnerBalances.find(b => normalizeText(b.account_item_name) === normalizedName);
      if (partnerBalance) {
        const partners = (partnerBalance.partners || []).filter(p => p.closing_balance !== 0);
        if (partners.length > 0) {
          pushEntriesWithSubtotal(config.accountName, partners);
        }
      }
    } else if (tagType === "品目") {
      // 品目内訳を検索
      const itemBalance = itemBalances.find(b => normalizeText(b.account_item_name) === normalizedName);
      if (itemBalance) {
        const items = (itemBalance.items || []).filter(item => item.closing_balance !== 0);
        if (items.length > 0) {
          pushEntriesWithSubtotal(config.accountName, items);
        }
      }
    } else {
      // タグ種類が指定されていない場合は、取引先→品目の順で探す（後方互換性）
      const partnerBalance = partnerBalances.find(b => normalizeText(b.account_item_name) === normalizedName);
      if (partnerBalance) {
        const partners = (partnerBalance.partners || []).filter(p => p.closing_balance !== 0);
        if (partners.length > 0) {
          pushEntriesWithSubtotal(config.accountName, partners);
          return;
        }
      }
      const itemBalance = itemBalances.find(b => normalizeText(b.account_item_name) === normalizedName);
      if (itemBalance) {
        const items = (itemBalance.items || []).filter(item => item.closing_balance !== 0);
        if (items.length > 0) {
          pushEntriesWithSubtotal(config.accountName, items);
        }
      }
    }
  });

  // 最後の空行を削除
  while (output.length > 0 && output[output.length - 1].every(v => v === "")) {
    output.pop();
  }

  const startRow = CONFIG.DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 2, lastRow - startRow + 1, 7).clearContent();
  }

  if (output.length > 0) {
    sheet.getRange(startRow, 2, output.length, 7).setValues(output);
    // 金額列（E〜H列）を3桁区切り・小数点なしに設定
    sheet.getRange(startRow, 5, output.length, 4).setNumberFormat("#,##0");
    subtotalRowOffsets.forEach(offset => {
      if (offset <= output.length) {
        sheet.getRange(CONFIG.HEADER_ROW + offset, 2, 1, 7).setBorder(true, false, false, false, false, false);
      }
    });
  }
}

function getTrialBSBreakdown(companyId, startDate, endDate, breakdownType) {
  const service = getService();
  if (!service.hasAccess()) {
    return [];
  }

  const options = getApiOptions();
  const url = CONFIG.API.REPORTS_URL + "trial_bs" +
    "?company_id=" + companyId +
    "&start_date=" + startDate +
    "&end_date=" + endDate +
    "&account_item_display_type=account_item" +
    "&breakdown_display_type=" + breakdownType;
  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());
  return json.trial_bs?.balances || [];
}

function getJournalCountForReport(companyId, startDate, endDate) {
  const options = getApiOptions();
  let totalCount = 0;

  // 振替伝票
  const mjUrl = CONFIG.API.MANUAL_JOURNALS + "?company_id=" + companyId +
    "&start_issue_date=" + startDate + "&end_issue_date=" + endDate + "&limit=1";
  const mjResponse = UrlFetchApp.fetch(mjUrl, options);
  const mjData = JSON.parse(mjResponse.getContentText());
  if (mjData.meta?.total_count) {
    totalCount += mjData.meta.total_count;
  }

  // 取引
  const dealsUrl = CONFIG.API.DEALS + "?company_id=" + companyId +
    "&start_issue_date=" + startDate + "&end_issue_date=" + endDate + "&limit=1";
  const dealsResponse = UrlFetchApp.fetch(dealsUrl, options);
  const dealsData = JSON.parse(dealsResponse.getContentText());
  if (dealsData.meta?.total_count) {
    totalCount += dealsData.meta.total_count;
  }

  return totalCount;
}

/**
 * BS・PLの勘定科目順序を構築
 */
function buildAccountOrder(bsBalances, plBalances) {
  const order = {};
  let index = 0;

  // BS科目
  bsBalances.forEach(item => {
    const accountName = item.account_item_name;
    if (accountName && !order.hasOwnProperty(accountName)) {
      order[accountName] = index++;
    }
  });

  // PL科目
  plBalances.forEach(item => {
    const accountName = item.account_item_name;
    if (accountName && !order.hasOwnProperty(accountName)) {
      order[accountName] = index++;
    }
  });

  return order;
}

/**
 * 区分別表を取得してシートに出力
 */
function getTaxCategoryReportCore(ss, companyId, startDate, endDate, taxAccountingMethod, accountOrder, timestamp) {
  const service = getService();
  if (!service.hasAccess()) {
    return;
  }

  // 区分別表シートへの書き込みは停止中（数値不一致のため）
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TAX_CATEGORY);

  const options = getApiOptions();

  // 勘定科目一覧を取得
  const accountItemsUrl = CONFIG.API.ACCOUNT_ITEMS + "?company_id=" + companyId;
  const accountItemsResponse = UrlFetchApp.fetch(accountItemsUrl, options);
  const accountItemsData = JSON.parse(accountItemsResponse.getContentText());
  const accountItems = {};
  const accountItemIdToName = {};
  const accountItemIdToCategory = {};
  if (accountItemsData.account_items) {
    accountItemsData.account_items.forEach(item => {
      accountItems[item.id] = item.name;
      accountItemIdToName[item.id] = item.name;
      accountItemIdToCategory[item.id] = (item.categories && item.categories.length > 0)
        ? item.categories[0]
        : (item.account_category || "");
    });
  }

  // 税区分一覧を取得
  const taxCodesUrl = CONFIG.API.TAX_CODES + "?company_id=" + companyId;
  const taxCodesResponse = UrlFetchApp.fetch(taxCodesUrl, options);
  const taxCodesData = JSON.parse(taxCodesResponse.getContentText());
  const taxCodes = {};
  if (taxCodesData.taxes) {
    taxCodesData.taxes.forEach(tax => {
      taxCodes[tax.code] = tax.name_ja || tax.name || String(tax.code);
    });
  }

  // PL税務検討用元帳の対象勘定科目（K7:K14から取得、並び順を維持）
  const plLedgerSheet = ss.getSheetByName(CONFIG.SHEETS.PL_TAX_LEDGER);
  let targetAccountNamesOrdered = [];
  if (plLedgerSheet) {
    const configRange = plLedgerSheet.getRange("K7:K14").getValues();
    targetAccountNamesOrdered = configRange
      .map(row => row[0])
      .filter(v => v && String(v).trim() !== "")
      .map(name => String(name).trim());
  }
  if (targetAccountNamesOrdered.length === 0) {
    targetAccountNamesOrdered = ["雑収入", "雑損失", "固定資産売却益", "固定資産売却損"];
  }
  const targetAccountNames = targetAccountNamesOrdered.map(name => normalizeText(name));
  const targetAccountIds = Object.keys(accountItems).filter(id => {
    const name = normalizeText(accountItems[id]);
    return targetAccountNames.some(target => name === target || name.includes(target));
  }).map(id => parseInt(id, 10));
  const targetAccountIdSet = new Set(targetAccountIds);

  // マスタデータ取得
  let partnerMap = {};
  let itemMap = {};
  let tagMap = {};

  if (targetAccountIds.length > 0) {
    partnerMap = getNameMapFromEndpoint("partners", companyId);
    itemMap = getNameMapFromEndpoint("items", companyId);
    tagMap = getNameMapFromEndpoint("tags", companyId);
  }

  const plLedgerRows = [];

  // 課税方式を取得
  const companyUrl = CONFIG.API.COMPANIES + "/" + companyId;
  const companyResponse = UrlFetchApp.fetch(companyUrl, options);
  const companyData = JSON.parse(companyResponse.getContentText()).company;
  let taxType = "";
  const taxMethod = companyData.tax_method_of_paying_tax;
  if (taxMethod === 0) taxType = "免税事業者";
  else if (taxMethod === 1) taxType = "原則課税";
  else if (taxMethod === 2) taxType = "簡易課税";

  // 取引データを取得
  const taxCategoryData = [];
  let offset = 0;
  const limit = CONFIG.PAGINATION.DEFAULT_LIMIT;

  while (true) {
    const dealsUrl = CONFIG.API.DEALS + "?company_id=" + companyId +
      "&start_issue_date=" + startDate + "&end_issue_date=" + endDate +
      "&limit=" + limit + "&offset=" + offset +
      "&item=full";
    const dealsResponse = UrlFetchApp.fetch(dealsUrl, options);
    const dealsData = JSON.parse(dealsResponse.getContentText());

    if (!dealsData.deals || dealsData.deals.length === 0) {
      break;
    }

    dealsData.deals.forEach(deal => {
      const dealPartnerId = deal.partner_id;
      const dealPartnerName = partnerMap[dealPartnerId] || "";
      const isCredit = deal.type === "income";
      const dealDate = deal.issue_date || "";
      // 預金明細の内容（deal.description）を取得、なければ取引タイプを表示
      const dealContent = deal.description || (deal.type === "income" ? "収入" : (deal.type === "expense" ? "支出" : "取引"));

      if (deal.details) {
        deal.details.forEach(detail => {
          const accountName = accountItems[detail.account_item_id] || "";
          const taxCodeName = taxCodes[detail.tax_code] || "対象外";
          const amount = detail.amount || 0;
          const vat = detail.vat || 0;

          taxCategoryData.push({
            accountName: accountName,
            taxCodeName: taxCodeName,
            amount: isCredit ? -amount : amount,
            vat: isCredit ? -vat : vat
          });

          if (targetAccountIdSet.has(detail.account_item_id)) {
            const debitAmount = isCredit ? 0 : amount;
            const creditAmount = isCredit ? amount : 0;

            let partnerName = "";
            if (detail.partner_id) {
              partnerName = partnerMap[detail.partner_id] || "";
            }
            if (!partnerName && dealPartnerId) {
              partnerName = dealPartnerName;
            }

            let itemName = "";
            if (detail.item_id) {
              itemName = itemMap[detail.item_id] || "";
            }

            let tagIds = detail.tag_ids || [];
            const tagNames = getTagNames(tagIds, tagMap);

            let description = detail.description || "";

            plLedgerRows.push([
              dealDate,
              accountName,
              taxCodeName,
              partnerName,
              itemName,
              tagNames,
              description,
              dealContent,
              debitAmount,
              creditAmount
            ]);
          }
        });
      }
    });

    if (dealsData.deals.length < limit) {
      break;
    }
    offset += limit;
  }

  // 振替伝票データを取得
  offset = 0;
  while (true) {
    const mjUrl = CONFIG.API.MANUAL_JOURNALS + "?company_id=" + companyId +
      "&start_issue_date=" + startDate + "&end_issue_date=" + endDate +
      "&limit=" + limit + "&offset=" + offset;
    const mjResponse = UrlFetchApp.fetch(mjUrl, options);
    const mjData = JSON.parse(mjResponse.getContentText());

    if (!mjData.manual_journals || mjData.manual_journals.length === 0) {
      break;
    }

    mjData.manual_journals.forEach(mj => {
      const mjDate = mj.issue_date || "";
      if (mj.details) {
        mj.details.forEach(detail => {
          const accountName = accountItems[detail.account_item_id] || "";
          const taxCodeName = taxCodes[detail.tax_code] || "対象外";
          const amount = detail.amount || 0;
          const vat = detail.vat || 0;
          const isCredit = detail.entry_side === "credit";

          taxCategoryData.push({
            accountName: accountName,
            taxCodeName: taxCodeName,
            amount: isCredit ? -amount : amount,
            vat: isCredit ? -vat : vat
          });

          if (targetAccountIdSet.has(detail.account_item_id)) {
            const debitAmount = isCredit ? 0 : amount;
            const creditAmount = isCredit ? amount : 0;

            let partnerName = "";
            if (detail.partner_id) {
              partnerName = partnerMap[detail.partner_id] || "";
            }

            let itemName = "";
            if (detail.item_id) {
              itemName = itemMap[detail.item_id] || "";
            }

            let tagIds = detail.tag_ids || [];
            const tagNames = getTagNames(tagIds, tagMap);

            let description = detail.description || "";

            plLedgerRows.push([
              mjDate,
              accountName,
              taxCodeName,
              partnerName,
              itemName,
              tagNames,
              description,
              "振替伝票",
              debitAmount,
              creditAmount
            ]);
          }
        });
      }
    });

    if (mjData.manual_journals.length < limit) {
      break;
    }
    offset += limit;
  }

  // 経費精算データを取得
  offset = 0;
  while (true) {
    const expUrl = CONFIG.API.EXPENSE_APPLICATIONS + "?company_id=" + companyId +
      "&start_issue_date=" + startDate + "&end_issue_date=" + endDate +
      "&limit=" + limit + "&offset=" + offset + "&status=approved";
    const expResponse = UrlFetchApp.fetch(expUrl, options);
    const expData = JSON.parse(expResponse.getContentText());

    if (!expData.expense_applications || expData.expense_applications.length === 0) {
      break;
    }

    expData.expense_applications.forEach(exp => {
      if (exp.expense_application_lines) {
        exp.expense_application_lines.forEach(line => {
          const accountName = accountItems[line.account_item_id] || "";
          const taxCodeName = taxCodes[line.tax_code] || "対象外";
          const amount = line.amount || 0;
          const vat = line.vat || 0;

          taxCategoryData.push({
            accountName: accountName,
            taxCodeName: taxCodeName,
            amount: amount,
            vat: vat
          });
        });
      }
    });

    if (expData.expense_applications.length < limit) {
      break;
    }
    offset += limit;
  }

  // 勘定科目×税区分で集計
  const summary = {};
  taxCategoryData.forEach(item => {
    const key = item.accountName + "|||" + item.taxCodeName;
    if (!summary[key]) {
      summary[key] = {
        accountName: item.accountName,
        taxCodeName: item.taxCodeName,
        totalAmount: 0,
        totalVat: 0
      };
    }
    summary[key].totalAmount += item.amount;
    summary[key].totalVat += item.vat;
  });

  // 出力データを作成
  const outputData = Object.values(summary)
    .filter(item => item.totalAmount !== 0 || item.totalVat !== 0)
    .map(item => {
      const taxExcluded = item.totalAmount - item.totalVat;
      return {
        accountName: item.accountName,
        taxCodeName: item.taxCodeName,
        taxExcluded: taxExcluded,
        vat: item.totalVat,
        amount: item.totalAmount,
        order: accountOrder[item.accountName] !== undefined ? accountOrder[item.accountName] : 999999
      };
    });

  outputData.sort((a, b) => {
    if (a.order !== b.order) {
      return a.order - b.order;
    }
    return a.taxCodeName.localeCompare(b.taxCodeName, 'ja');
  });

  const outputArray = outputData.map(item => [
    item.accountName,
    item.taxCodeName,
    item.taxExcluded,
    item.vat,
    item.amount
  ]);

  // ===== 区分別表シートへの書き込みは停止中（数値不一致のため） =====
  // const startRow = CONFIG.TAX_CATEGORY_DATA_START_ROW;
  // const lastRowBefore = sheet.getLastRow();
  // if (lastRowBefore >= startRow) {
  //   sheet.getRange(startRow, 2, lastRowBefore - startRow + 1, 5).clearContent();
  //   sheet.getRange(startRow, 2, lastRowBefore - startRow + 1, 5).setBorder(false, false, false, false, false, false);
  // }
  // sheet.getRange(CONFIG.CELLS.TAX_CATEGORY_TIMESTAMP).setValue(timestamp);
  // sheet.getRange(CONFIG.CELLS.TAX_CATEGORY_TAX_TYPE).setValue(taxType);
  // const headers = ["勘定科目", "税区分", "税抜金額", "税額", "税込金額"];
  // sheet.getRange(CONFIG.TAX_CATEGORY_HEADER_ROW, 2, 1, headers.length).setValues([headers]);
  // sheet.getRange(CONFIG.TAX_CATEGORY_HEADER_ROW, 2, 1, headers.length).setFontWeight("bold").setBackground("#4285f4").setFontColor("#ffffff");
  // if (outputArray.length > 0) {
  //   sheet.getRange(startRow, 2, outputArray.length, 5).setValues(outputArray);
  //   sheet.getRange(startRow, 4, outputArray.length, 3).setNumberFormat("#,##0");
  //   const dataLastRow = CONFIG.TAX_CATEGORY_HEADER_ROW + outputArray.length;
  //   sheet.getRange(dataLastRow, 2, 1, 12).setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  //   const sheetLastRow = sheet.getLastRow();
  //   if (sheetLastRow > dataLastRow) {
  //     const bValues = sheet.getRange(dataLastRow + 1, 2, sheetLastRow - dataLastRow, 1).getValues();
  //     let deleteCount = 0;
  //     for (let i = bValues.length - 1; i >= 0; i--) {
  //       if (bValues[i][0] === "" || bValues[i][0] === null) { deleteCount++; } else { break; }
  //     }
  //     if (deleteCount > 0) { sheet.deleteRows(sheetLastRow - deleteCount + 1, deleteCount); }
  //   }
  // } else {
  //   sheet.getRange(startRow, 2).setValue("該当データがありません");
  // }
  // sheet.setColumnWidth(2, 200);
  // sheet.setColumnWidth(3, 150);
  // sheet.setColumnWidth(4, 120);
  // sheet.setColumnWidth(5, 100);
  // sheet.setColumnWidth(6, 120);

  // PL税務検討用元帳のソート（K7:K14の並び順に従う）
  // targetAccountNamesOrderedの順番に基づいてソート
  const accountOrderMap = {};
  targetAccountNamesOrdered.forEach((name, idx) => {
    accountOrderMap[normalizeText(name)] = idx;
  });

  plLedgerRows.sort((a, b) => {
    // 勘定科目の順番（K7:K14の順）
    const orderA = accountOrderMap[normalizeText(a[1])] !== undefined ? accountOrderMap[normalizeText(a[1])] : 999;
    const orderB = accountOrderMap[normalizeText(b[1])] !== undefined ? accountOrderMap[normalizeText(b[1])] : 999;
    if (orderA !== orderB) return orderA - orderB;
    // 同じ勘定科目内は日付順
    return String(a[0]).localeCompare(String(b[0]), "ja");
  });

  // 勘定科目ごとに小計を追加し、1行空ける
  const plLedgerRowsWithSubtotals = [];
  let currentAccountName = "";
  let sumDebit = 0;
  let sumCredit = 0;

  plLedgerRows.forEach((row, idx) => {
    const accountName = row[1];
    const normalizedAccountName = normalizeText(accountName);

    // 勘定科目が変わった場合、前の科目の小計を出力
    if (currentAccountName && normalizedAccountName !== normalizeText(currentAccountName)) {
      // 小計行を追加
      plLedgerRowsWithSubtotals.push([
        "", currentAccountName, "", "", "", "", "", "小計", sumDebit, sumCredit
      ]);
      // 空行を追加
      plLedgerRowsWithSubtotals.push(["", "", "", "", "", "", "", "", "", ""]);
      sumDebit = 0;
      sumCredit = 0;
    }

    currentAccountName = accountName;
    sumDebit += row[8] || 0;
    sumCredit += row[9] || 0;
    plLedgerRowsWithSubtotals.push(row);

    // 最後のデータの場合も小計を出力
    if (idx === plLedgerRows.length - 1) {
      plLedgerRowsWithSubtotals.push([
        "", currentAccountName, "", "", "", "", "", "小計", sumDebit, sumCredit
      ]);
    }
  });

  return {
    plLedgerRows: plLedgerRowsWithSubtotals
  };
}

function getFixedAssetsCore(ss, companyId, fiscalYear, targetDateStr) {
  const service = getService();
  if (!service.hasAccess()) {
    return;
  }

  const sheet = ss.getSheetByName(CONFIG.SHEETS.FIXED_ASSETS);
  if (!sheet) {
    return;
  }

  const options = getApiOptions();

  // 固定資産一覧を取得
  const fixedAssetsUrl = CONFIG.API.FIXED_ASSETS + "?company_id=" + companyId +
    (targetDateStr ? "&target_date=" + targetDateStr : "");
  const fixedAssetsResponse = UrlFetchApp.fetch(fixedAssetsUrl, options);

  if (fixedAssetsResponse.getResponseCode() !== 200) {
    return;
  }

  const fixedAssetsData = JSON.parse(fixedAssetsResponse.getContentText());

  if (!fixedAssetsData.fixed_assets || fixedAssetsData.fixed_assets.length === 0) {
    return;
  }

  // 出力データを作成
  const outputData = [];
  fixedAssetsData.fixed_assets.forEach(asset => {
    outputData.push([
      asset.name || "",
      asset.acquisition_date || "",
      asset.acquisition_cost || 0,
      asset.depreciation_method || "",
      asset.useful_life || "",
      asset.depreciation_amount || 0,
      asset.closing_accumulated_depreciation || asset.accumulated_depreciation || 0,
      asset.undepreciated_balance || asset.book_value || 0
    ]);
  });

  // シートをクリアして出力
  const startRow = CONFIG.FIXED_ASSETS_DATA_START_ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 2, lastRow - startRow + 1, 8).clearContent();
  }

  if (outputData.length > 0) {
    sheet.getRange(startRow, 2, outputData.length, 8).setValues(outputData);
    // 金額列に3桁区切りフォーマットを設定（D:取得価額, G:当期償却額, H:期末減価償却累計額, I:期末帳簿価額）
    sheet.getRange(startRow, 4, outputData.length, 1).setNumberFormat("#,##0");
    sheet.getRange(startRow, 7, outputData.length, 3).setNumberFormat("#,##0");
  }

  // タイムスタンプ
  sheet.getRange(CONFIG.CELLS.FIXED_ASSETS_TIMESTAMP).setValue(Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm") + "更新");
}
