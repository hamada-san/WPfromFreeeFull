/**
 * 試算表・PLを取得してシートに出力（コア処理）
 */
function getTrialBalanceAndPLCore(ss, companyId, startDateStr, endDateStr) {
  const service = getService();
  if (!service.hasAccess()) {
    throw new Error("認証されていません。メニューから認証を行ってください。");
  }
  
  const options = {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };
  
  // 会計年度情報を取得
  const companyUrl = "https://api.freee.co.jp/api/1/companies/" + companyId;
  const companyResponse = UrlFetchApp.fetch(companyUrl, options);
  const companyInfo = JSON.parse(companyResponse.getContentText());
  const fiscalYears = companyInfo.company.fiscal_years;
  
  if (!fiscalYears || fiscalYears.length === 0) {
    throw new Error("会計年度情報が取得できません。");
  }
  
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
    // 最新の会計年度を使用
    targetFiscalYear = fiscalYears[fiscalYears.length - 1];
    startDateStr = targetFiscalYear.start_date;
    endDateStr = targetFiscalYear.end_date;
  }
  
  const fiscalYear = parseInt(startDateStr.substring(0, 4), 10);
  const startMonth = parseInt(startDateStr.substring(5, 7), 10);
  const endMonth = parseInt(endDateStr.substring(5, 7), 10);
  const endYear = parseInt(endDateStr.substring(0, 4), 10);
  
  // 期ラベル（YY/MM-YY/MM形式）
  const formatPeriod = (startY, startM, endY, endM) => {
    const sy = String(startY).slice(-2);
    const ey = String(endY).slice(-2);
    const sm = String(startM).padStart(2, '0');
    const em = String(endM).padStart(2, '0');
    return sy + "/" + sm + "-" + ey + "/" + em;
  };
  
  const startYear = parseInt(startDateStr.substring(0, 4), 10);
  const periodLabels = {
    current: formatPeriod(startYear, startMonth, endYear, endMonth),
    previous: formatPeriod(startYear - 1, startMonth, endYear - 1, endMonth),
    twoYearsAgo: formatPeriod(startYear - 2, startMonth, endYear - 2, endMonth)
  };
  
  // 税区分を取得
  let taxAccountingMethod = "";
  if (companyInfo.company.tax_at_source_calc_type === 1) {
    taxAccountingMethod = "税抜経理";
  } else if (companyInfo.company.tax_at_source_calc_type === 0) {
    taxAccountingMethod = "税込経理";
  }
  
  const baseUrl = "https://api.freee.co.jp/api/1/reports/";
  const params = "?company_id=" + companyId + "&fiscal_year=" + fiscalYear + "&start_month=" + startMonth + "&end_month=" + endMonth;
  const timestamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm") + "更新";
  
  // ===== BS（当期のみ）=====
  const bsSheet = ss.getSheetByName("BS");
  if (bsSheet) {
    const bsLastRow = bsSheet.getLastRow();
    if (bsLastRow >= 17) {
      bsSheet.getRange(17, 2, bsLastRow - 16, 8).clearContent();
      bsSheet.getRange(17, 2, bsLastRow - 16, 8).setBorder(false, false, false, false, false, false);
      bsSheet.getRange(17, 2, bsLastRow - 16, 8).setBackground(null);
    }
    
    bsSheet.getRange("G15").setValue(timestamp);
    
    const bsUrl = baseUrl + "trial_bs?company_id=" + companyId + "&start_date=" + startDateStr + "&end_date=" + endDateStr;
    const bsRes = UrlFetchApp.fetch(bsUrl, options);
    const bsBalances = JSON.parse(bsRes.getContentText()).trial_bs.balances;
    
    const bsHeaders = ["分類", "勘定科目", "期首残高", "借方金額", "貸方金額", "期末残高", "", "", "プリペアコメント", "基礎資料又はfreee仕訳リンク", "レビュワーコメント"];
    bsSheet.getRange(16, 2, 1, 8).setValues([bsHeaders.slice(0, 8)]);
    
    const bsRows = bsBalances.map(i => {
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
      bsSheet.getRange(17, 2, bsRows.length, 6).setValues(bsRows);
      bsSheet.getRange(17, 7, bsRows.length, 1).setBackground("#f4cccc");
      
      const bsBorderTargets = ["資産", "負債", "純資産"];
      for (let i = 0; i < bsRows.length; i++) {
        const categoryName = bsRows[i][0].replace(/^▼/, "").trim();
        if (bsBorderTargets.includes(categoryName)) {
          bsSheet.getRange(17 + i, 2, 1, 8).setBorder(true, false, false, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        }
      }
    }
  }
  
  // ===== PL（3期比較）=====
  const plSheet = ss.getSheetByName("PL");
  if (plSheet) {
    const plLastRow = plSheet.getLastRow();
    if (plLastRow >= 17) {
      plSheet.getRange(17, 2, plLastRow - 16, 8).clearContent();
      plSheet.getRange(17, 2, plLastRow - 16, 8).setBorder(false, false, false, false, false, false);
      plSheet.getRange(17, 2, plLastRow - 16, 8).setBackground(null);
    }
    
    plSheet.getRange("G15").setValue(timestamp);
    plSheet.getRange("F15").setValue(taxAccountingMethod);
    
    // 3期比較APIを使用
    const plUrl = baseUrl + "trial_pl_three_years" + params;
    const plRes = UrlFetchApp.fetch(plUrl, options);
    const plBalances = JSON.parse(plRes.getContentText()).trial_pl_three_years.balances;
    
    const plHeaders = ["分類", "勘定科目", periodLabels.twoYearsAgo, periodLabels.previous, periodLabels.current, "前年差額", "前年比", "", "プリペアコメント", "基礎資料又はfreee仕訳リンク", "レビュワーコメント"];
    plSheet.getRange(16, 2, 1, 8).setValues([plHeaders.slice(0, 8)]);
    
    const plRows = plBalances.map(item => {
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
      plSheet.getRange(17, 2, plRows.length, 7).setValues(plRows);
      plSheet.getRange(17, 6, plRows.length, 1).setBackground("#f4cccc");
      
      const plBorderTargets = ["売上総損益金額", "営業損益金額", "経常損益金額", "税引前当期純損益金額", "当期純損益金額"];
      for (let i = 0; i < plRows.length; i++) {
        if (plBorderTargets.includes(plRows[i][0].trim())) {
          plSheet.getRange(17 + i, 2, 1, 8).setBorder(true, false, false, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        }
      }
    }
  }
  
  // ===== CR（3期比較）=====
  const crSheet = ss.getSheetByName("CR");
  if (crSheet) {
    const crLastRow = crSheet.getLastRow();
    if (crLastRow >= 17) {
      crSheet.getRange(17, 2, crLastRow - 16, 8).clearContent();
      crSheet.getRange(17, 2, crLastRow - 16, 8).setBorder(false, false, false, false, false, false);
      crSheet.getRange(17, 2, crLastRow - 16, 8).setBackground(null);
    }
    
    crSheet.getRange("G15").setValue(timestamp);
    
    try {
      // 3期比較APIを使用
      const crUrl = baseUrl + "trial_cr_three_years" + params;
      const crRes = UrlFetchApp.fetch(crUrl, options);
      
      if (crRes.getResponseCode() === 200) {
        const crData = JSON.parse(crRes.getContentText());
        const crBalances = crData.trial_cr_three_years?.balances;
        
        if (crBalances && crBalances.length > 0) {
          crSheet.getRange("B16").setValue("製造原価報告書");
          
          const crHeaders = ["分類", "勘定科目", periodLabels.twoYearsAgo, periodLabels.previous, periodLabels.current, "前年差額", "前年比", "", "プリペアコメント", "基礎資料又はfreee仕訳リンク", "レビュワーコメント"];
          crSheet.getRange(16, 2, 1, 8).setValues([crHeaders.slice(0, 8)]);
          
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
            crSheet.getRange(17, 2, crRows.length, 7).setValues(crRows);
            crSheet.getRange(17, 6, crRows.length, 1).setBackground("#f4cccc");
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
  
  // ===== 仕訳数 =====
  const journalCount = getJournalCountForReport(companyId, startDateStr, endDateStr);
  const docketSheet = ss.getSheetByName("管理ドケット");
  if (docketSheet) {
    docketSheet.getRange("D15").setValue(journalCount + "仕訳");
  }
  
  // ===== 区分別表 =====
  const bsBalancesForOrder = JSON.parse(UrlFetchApp.fetch(baseUrl + "trial_bs?company_id=" + companyId + "&start_date=" + startDateStr + "&end_date=" + endDateStr, options).getContentText()).trial_bs?.balances || [];
  const plBalancesForOrder = JSON.parse(UrlFetchApp.fetch(baseUrl + "trial_pl?company_id=" + companyId + "&start_date=" + startDateStr + "&end_date=" + endDateStr, options).getContentText()).trial_pl?.balances || [];
  const accountOrder = buildAccountOrder(bsBalancesForOrder, plBalancesForOrder);
  
  const taxCategoryResult = getTaxCategoryReportCore(ss, companyId, startDateStr, endDateStr, taxAccountingMethod, accountOrder, timestamp);

  // ===== PL税務検討用元帳 =====
  if (taxCategoryResult && taxCategoryResult.plLedgerRows) {
    writePLTaxLedgerSheet(ss, taxCategoryResult.plLedgerRows, timestamp);
  }

  // ===== BS税務検討用内訳 =====
  const accountItemCategoryMap = getAccountItemCategoryMap(companyId);
  try {
    getBSTaxBreakdownCore(ss, companyId, startDateStr, endDateStr, accountItemCategoryMap, timestamp);
  } catch (e) {
    Logger.log("BS税務検討用内訳取得エラー: " + e.message);
  }
  
  // ===== 固定資産台帳 =====
  // ※freeeプランによりAPIアクセス不可の場合があるためtry-catchで囲む
  try {
    getFixedAssetsCore(ss, companyId, fiscalYear, startDateStr);
  } catch (e) {
    Logger.log("固定資産台帳取得エラー: " + e.message);
  }
}

/**
 * 仕訳数を取得（reports.gs用）
 */
/**
 * 勘定科目名 -> 大分類のマップを作成
 */
function getAccountItemCategoryMap(companyId) {
  const service = getService();
  if (!service.hasAccess()) {
    return {};
  }
  
  const options = {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };
  
  const url = "https://api.freee.co.jp/api/1/account_items?company_id=" + companyId;
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
 * 仕訳帳CSVを取得
 */
function getJournalsCsvRows(companyId, startDate, endDate) {
  const service = getService();
  if (!service.hasAccess()) {
    throw new Error("認証されていません。");
  }
  
  const options = {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };
  
  const baseUrl = "https://api.freee.co.jp/api/1/journals";
  const url = baseUrl +
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
  const maxTries = 30;
  
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
    // 旧CSVで先頭に行番号列が付くケースの補正
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

  const options = {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };

  const map = {};
  const limit = 100; // freee APIの上限は100
  let offset = 0;
  let key = "";
  if (resource === "partners") key = "partners";
  if (resource === "items") key = "items";
  if (resource === "tags") key = "tags";
  if (!key) return map;

  while (true) {
    const url = "https://api.freee.co.jp/api/1/" + resource +
      "?company_id=" + companyId +
      "&limit=" + limit +
      "&offset=" + offset;
    const res = UrlFetchApp.fetch(url, options);

    // エラーチェック
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
  const sheet = ss.getSheetByName("PL税務検討用元帳");
  if (!sheet) {
    return;
  }
  if (timestamp) {
    sheet.getRange("G15").setValue(timestamp);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow >= 17) {
    // データと罫線をクリア（B〜L列）
    sheet.getRange(17, 2, lastRow - 16, 11).clearContent();
    sheet.getRange(17, 2, lastRow - 16, 11).setBorder(false, false, false, false, false, false);
  }

  const headersRow = ["大分類", "勘定科目", "税区分", "取引先タグ", "品目タグ", "メモタグ", "摘要", "借方金額", "貸方金額"];
  sheet.getRange(16, 2, 1, headersRow.length).setValues([headersRow]);

  if (rows && rows.length > 0) {
    sheet.getRange(17, 2, rows.length, 9).setValues(rows);

    // 勘定科目が変わる行の上部に罫線を引く（B〜L列）
    let prevAccountName = "";
    for (let i = 0; i < rows.length; i++) {
      const currentAccountName = rows[i][1]; // 勘定科目は2列目（インデックス1）
      if (i > 0 && currentAccountName !== prevAccountName) {
        // 勘定科目が変わった行の上に罫線
        sheet.getRange(17 + i, 2, 1, 11).setBorder(true, false, false, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      }
      prevAccountName = currentAccountName;
    }

    // 最後のデータの下にも罫線
    sheet.getRange(17 + rows.length - 1, 2, 1, 11).setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  }
}

/**
 * PL税務検討用元帳を出力
 */
function getPLTaxLedgerCore(ss, companyId, startDate, endDate, accountItemCategoryMap, timestamp) {
  const sheet = ss.getSheetByName("PL税務検討用元帳");
  if (!sheet) {
    return;
  }
  if (timestamp) {
    sheet.getRange("G15").setValue(timestamp);
  }
  
  const targetAccounts = ["雑収入", "雑損失", "固定資産売却益", "固定資産売却損"]
    .map(name => normalizeText(name));
  const rows = getJournalsCsvRows(companyId, startDate, endDate);
  if (!rows || rows.length === 0) {
    return;
  }
  
  const headers = rows[0];

  const debitAccountIdx = findHeaderIndex(headers, ["借方勘定科目", "借方科目", "借方勘定科目名"]);
  const creditAccountIdx = findHeaderIndex(headers, ["貸方勘定科目", "貸方科目", "貸方勘定科目名"]);
  const debitAmountIdx = findHeaderIndex(headers, ["借方金額"]);
  const creditAmountIdx = findHeaderIndex(headers, ["貸方金額"]);
  const debitTaxIdx = findHeaderIndex(headers, ["借方税区分", "借方税区分名"]);
  const creditTaxIdx = findHeaderIndex(headers, ["貸方税区分", "貸方税区分名"]);
  // 取引先：freee仕訳帳CSVでは「取引先」列が1つ
  const partnerIdx = findHeaderIndex(headers, ["取引先", "取引先名"]);
  const debitPartnerIdx = findHeaderIndex(headers, ["借方取引先"]);
  const creditPartnerIdx = findHeaderIndex(headers, ["貸方取引先"]);
  // 品目：freee仕訳帳CSVでは「品目」列が1つ
  const itemIdx = findHeaderIndex(headers, ["品目", "品目名"]);
  const debitItemIdx = findHeaderIndex(headers, ["借方品目"]);
  const creditItemIdx = findHeaderIndex(headers, ["貸方品目"]);
  // メモタグ：freee仕訳帳CSVでは「メモタグ」列が1つ
  const tagIdx = findHeaderIndex(headers, ["メモタグ", "タグ"]);
  const debitTagIdx = findHeaderIndex(headers, ["借方メモタグ"]);
  const creditTagIdx = findHeaderIndex(headers, ["貸方メモタグ"]);
  // 摘要：freee仕訳帳CSVでは「摘要」列が1つ
  const descIdx = findHeaderIndex(headers, ["摘要", "備考"]);
  const debitDescIdx = findHeaderIndex(headers, ["借方摘要"]);
  const creditDescIdx = findHeaderIndex(headers, ["貸方摘要"]);
  
  const output = [];

  // ヘルパー関数：借方/貸方別インデックス → 共通インデックスのフォールバック
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
  
  const lastRow = sheet.getLastRow();
  if (lastRow >= 17) {
    sheet.getRange(17, 2, lastRow - 16, 9).clearContent();
  }
  
  const headersRow = ["大分類", "勘定科目", "税区分", "取引先タグ", "品目タグ", "メモタグ", "摘要", "借方金額", "貸方金額"];
  sheet.getRange(16, 2, 1, headersRow.length).setValues([headersRow]);
  
  if (output.length > 0) {
    sheet.getRange(17, 2, output.length, 9).setValues(output);
  }
}

/**
 * BS税務検討用内訳を出力
 */
function getBSTaxBreakdownCore(ss, companyId, startDate, endDate, accountItemCategoryMap, timestamp) {
  const sheet = ss.getSheetByName("BS税務検討用内訳");
  if (!sheet) {
    return;
  }
  if (timestamp) {
    sheet.getRange("G15").setValue(timestamp);
  }
  
  const partnerAccounts = ["未払金", "未払費用"];
  const itemAccounts = ["預り金", "長期借入金"];
  const partnerAccountsNormalized = partnerAccounts.map(name => normalizeText(name));
  const itemAccountsNormalized = itemAccounts.map(name => normalizeText(name));
  const output = [];
  const subtotalRowOffsets = [];
  const pushEntriesWithSubtotal = (accountName, entries) => {
    if (!entries || entries.length === 0) return;
    const category = accountItemCategoryMap[accountName] || "";
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
  
  const partnerBalances = getTrialBSBreakdown(companyId, startDate, endDate, "partner");
  partnerAccounts.forEach((accountName, idx) => {
    const balance = partnerBalances.find(b => normalizeText(b.account_item_name) === partnerAccountsNormalized[idx]);
    if (!balance) return;
    const partners = (balance.partners || []).filter(p => p.closing_balance !== 0);
    pushEntriesWithSubtotal(accountName, partners);
  });
  
  const itemBalances = getTrialBSBreakdown(companyId, startDate, endDate, "item");
  itemAccounts.forEach((accountName, idx) => {
    const balance = itemBalances.find(b => normalizeText(b.account_item_name) === itemAccountsNormalized[idx]);
    if (!balance) return;
    const items = (balance.items || []).filter(item => item.closing_balance !== 0);
    pushEntriesWithSubtotal(accountName, items);
  });
  
  while (output.length > 0 && output[output.length - 1].every(v => v === "")) {
    output.pop();
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow >= 17) {
    sheet.getRange(17, 2, lastRow - 16, 7).clearContent();
  }
  
  if (output.length > 0) {
    sheet.getRange(17, 2, output.length, 7).setValues(output);
    subtotalRowOffsets.forEach(offset => {
      if (offset <= output.length) {
        sheet.getRange(16 + offset, 2, 1, 7).setBorder(true, false, false, false, false, false);
      }
    });
  }
}

function getTrialBSBreakdown(companyId, startDate, endDate, breakdownType) {
  const service = getService();
  if (!service.hasAccess()) {
    return [];
  }
  
  const options = {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };
  
  const url = "https://api.freee.co.jp/api/1/reports/trial_bs" +
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
  const service = getService();
  const options = {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };
  
  let totalCount = 0;
  
  // 振替伝票
  const mjUrl = "https://api.freee.co.jp/api/1/manual_journals?company_id=" + companyId + 
                "&start_issue_date=" + startDate + "&end_issue_date=" + endDate + "&limit=1";
  const mjResponse = UrlFetchApp.fetch(mjUrl, options);
  const mjData = JSON.parse(mjResponse.getContentText());
  if (mjData.meta?.total_count) {
    totalCount += mjData.meta.total_count;
  }
  
  // 取引
  const dealsUrl = "https://api.freee.co.jp/api/1/deals?company_id=" + companyId + 
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
  
  const sheet = ss.getSheetByName("区分別表");
  if (!sheet) {
    return;
  }
  
  const options = {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };
  
  // 勘定科目一覧を取得
  const accountItemsUrl = "https://api.freee.co.jp/api/1/account_items?company_id=" + companyId;
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
  
  // 税区分一覧を取得（日本語名を取得）
  const taxCodesUrl = "https://api.freee.co.jp/api/1/taxes/codes?company_id=" + companyId;
  const taxCodesResponse = UrlFetchApp.fetch(taxCodesUrl, options);
  const taxCodesData = JSON.parse(taxCodesResponse.getContentText());
  const taxCodes = {};
  if (taxCodesData.taxes) {
    taxCodesData.taxes.forEach(tax => {
      taxCodes[tax.code] = tax.name_ja || tax.name || String(tax.code);
    });
  }
  
  // PL税務検討用元帳の対象勘定科目（シートのJ7:J14から読み取り）
  const plLedgerSheet = ss.getSheetByName("PL税務検討用元帳");
  let targetAccountNames = [];
  if (plLedgerSheet) {
    const configRange = plLedgerSheet.getRange("J7:J14").getValues();
    targetAccountNames = configRange
      .map(row => row[0])
      .filter(v => v && String(v).trim() !== "")
      .map(name => normalizeText(String(name)));
  }
  // フォールバック：シートに設定がない場合はデフォルト
  if (targetAccountNames.length === 0) {
    targetAccountNames = ["雑収入", "雑損失", "固定資産売却益", "固定資産売却損"]
      .map(name => normalizeText(name));
  }
  const targetAccountIds = Object.keys(accountItems).filter(id => {
    const name = normalizeText(accountItems[id]);
    return targetAccountNames.some(target => name === target || name.includes(target));
  }).map(id => parseInt(id, 10));
  const targetAccountIdSet = new Set(targetAccountIds);
  
  // マスタデータは対象勘定科目がある場合のみ取得（処理時間短縮）
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
  const companyUrl = "https://api.freee.co.jp/api/1/companies/" + companyId;
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
  const limit = 100;
  
  while (true) {
    const dealsUrl = "https://api.freee.co.jp/api/1/deals?company_id=" + companyId +
                     "&start_issue_date=" + startDate + "&end_issue_date=" + endDate +
                     "&limit=" + limit + "&offset=" + offset +
                     "&item=full";
    const dealsResponse = UrlFetchApp.fetch(dealsUrl, options);
    const dealsData = JSON.parse(dealsResponse.getContentText());
    
    if (!dealsData.deals || dealsData.deals.length === 0) {
      break;
    }
    
    dealsData.deals.forEach(deal => {
      // 取引レベルの情報を取得
      const dealPartnerId = deal.partner_id;
      const dealPartnerName = partnerMap[dealPartnerId] || "";
      const isCredit = deal.type === "income";

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

            // 取引先：明細 → 取引レベル
            let partnerName = "";
            if (detail.partner_id) {
              partnerName = partnerMap[detail.partner_id] || "";
            }
            if (!partnerName && dealPartnerId) {
              partnerName = dealPartnerName;
            }

            // 品目：明細レベル
            let itemName = "";
            if (detail.item_id) {
              itemName = itemMap[detail.item_id] || "";
            }

            // タグ：明細レベル
            let tagIds = detail.tag_ids || [];
            const tagNames = getTagNames(tagIds, tagMap);

            // 摘要：明細の備考
            let description = detail.description || "";

            plLedgerRows.push([
              accountItemIdToCategory[detail.account_item_id] || "",
              accountName,
              taxCodeName,
              partnerName,
              itemName,
              tagNames,
              description,
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
    const mjUrl = "https://api.freee.co.jp/api/1/manual_journals?company_id=" + companyId + 
                  "&start_issue_date=" + startDate + "&end_issue_date=" + endDate + 
                  "&limit=" + limit + "&offset=" + offset;
    const mjResponse = UrlFetchApp.fetch(mjUrl, options);
    const mjData = JSON.parse(mjResponse.getContentText());
    
    if (!mjData.manual_journals || mjData.manual_journals.length === 0) {
      break;
    }
    
    mjData.manual_journals.forEach(mj => {
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

            // 取引先
            let partnerName = "";
            if (detail.partner_id) {
              partnerName = partnerMap[detail.partner_id] || "";
            }

            // 品目
            let itemName = "";
            if (detail.item_id) {
              itemName = itemMap[detail.item_id] || "";
            }

            // タグ：明細レベル
            let tagIds = detail.tag_ids || [];
            const tagNames = getTagNames(tagIds, tagMap);

            // 摘要：明細レベル
            let description = detail.description || "";

            plLedgerRows.push([
              accountItemIdToCategory[detail.account_item_id] || "",
              accountName,
              taxCodeName,
              partnerName,
              itemName,
              tagNames,
              description,
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
    const expUrl = "https://api.freee.co.jp/api/1/expense_applications?company_id=" + companyId + 
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
  
  // 出力データを作成し、BS→PLの順にソート
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
  
  // 試算表の並び順でソート
  outputData.sort((a, b) => {
    if (a.order !== b.order) {
      return a.order - b.order;
    }
    return a.taxCodeName.localeCompare(b.taxCodeName, 'ja');
  });
  
  // 配列に変換
  const outputArray = outputData.map(item => [
    item.accountName,
    item.taxCodeName,
    item.taxExcluded,
    item.vat,
    item.amount
  ]);
  
  // B44以降のデータ部分のみクリア
  const lastRowBefore = sheet.getLastRow();
  if (lastRowBefore >= 44) {
    sheet.getRange(44, 2, lastRowBefore - 43, 5).clearContent();
    sheet.getRange(44, 2, lastRowBefore - 43, 5).setBorder(false, false, false, false, false, false);
  }
  
  // タイムスタンプと課税方式
  sheet.getRange("F42").setValue(timestamp);
  sheet.getRange("E42").setValue(taxType);
  
  // ヘッダー行（43行目）
  const headers = ["勘定科目", "税区分", "税抜金額", "税額", "税込金額"];
  sheet.getRange(43, 2, 1, headers.length).setValues([headers]);
  sheet.getRange(43, 2, 1, headers.length).setFontWeight("bold").setBackground("#4285f4").setFontColor("#ffffff");
  
  // データ出力（B44から）
  if (outputArray.length > 0) {
    sheet.getRange(44, 2, outputArray.length, 5).setValues(outputArray);
    sheet.getRange(44, 4, outputArray.length, 3).setNumberFormat("#,##0");
    
    const dataLastRow = 43 + outputArray.length;
    sheet.getRange(dataLastRow, 2, 1, 12).setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    
    const sheetLastRow = sheet.getLastRow();
    if (sheetLastRow > dataLastRow) {
      const bValues = sheet.getRange(dataLastRow + 1, 2, sheetLastRow - dataLastRow, 1).getValues();
      let deleteCount = 0;
      for (let i = bValues.length - 1; i >= 0; i--) {
        if (bValues[i][0] === "" || bValues[i][0] === null) {
          deleteCount++;
        } else {
          break;
        }
      }
      if (deleteCount > 0) {
        sheet.deleteRows(sheetLastRow - deleteCount + 1, deleteCount);
      }
    }
  } else {
    sheet.getRange(44, 2).setValue("該当データがありません");
  }
  
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 120);
  
  // PL税務検討用元帳のソート
  plLedgerRows.sort((a, b) => {
    const c = String(a[0]).localeCompare(String(b[0]), "ja");
    if (c !== 0) return c;
    return String(a[1]).localeCompare(String(b[1]), "ja");
  });
  
  return {
    plLedgerRows: plLedgerRows
  };
}

function getFixedAssetsCore(ss, companyId, fiscalYear, targetDateStr) {
  const service = getService();
  if (!service.hasAccess()) {
    return;
  }
  
  const sheet = ss.getSheetByName("固定資産台帳");
  if (!sheet) {
    return;
  }
  
  const options = {
    method: "get",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
    muteHttpExceptions: true
  };
  
  // 固定資産一覧を取得
  const fixedAssetsUrl = "https://api.freee.co.jp/api/1/fixed_assets?company_id=" + companyId +
    (targetDateStr ? "&target_date=" + targetDateStr : "");
  const fixedAssetsResponse = UrlFetchApp.fetch(fixedAssetsUrl, options);

  if (fixedAssetsResponse.getResponseCode() !== 200) {
    // プラン制限等でアクセスできない場合は終了
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
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    sheet.getRange(5, 2, lastRow - 4, 8).clearContent();
  }
  
  if (outputData.length > 0) {
    sheet.getRange(5, 2, outputData.length, 8).setValues(outputData);
  }
  
  // タイムスタンプ
  sheet.getRange("J3").setValue(Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm") + "更新");
}
