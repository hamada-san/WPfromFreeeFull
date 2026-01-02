/**
 * 試算表・PLを取得してシートに出力（コア処理）
 */
function getTrialBalanceAndPLCore(ss, companyId) {
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
  
  // 最新の会計年度を使用
  const targetFiscalYear = fiscalYears[fiscalYears.length - 1];
  
  const fiscalYear = parseInt(targetFiscalYear.start_date.substring(0, 4));
  const startMonth = parseInt(targetFiscalYear.start_date.substring(5, 7));
  const endMonth = parseInt(targetFiscalYear.end_date.substring(5, 7));
  const endYear = parseInt(targetFiscalYear.end_date.substring(0, 4));
  
  const startDateStr = targetFiscalYear.start_date;
  const endDateStr = targetFiscalYear.end_date;
  
  // 期ラベル（YY/MM-YY/MM形式）
  const formatPeriod = (startY, startM, endY, endM) => {
    const sy = String(startY).slice(-2);
    const ey = String(endY).slice(-2);
    const sm = String(startM).padStart(2, '0');
    const em = String(endM).padStart(2, '0');
    return sy + "/" + sm + "-" + ey + "/" + em;
  };
  
  const periodLabels = {
    current: formatPeriod(endYear - 1, startMonth, endYear, endMonth),
    previous: formatPeriod(endYear - 2, startMonth, endYear - 1, endMonth),
    twoYearsAgo: formatPeriod(endYear - 3, startMonth, endYear - 2, endMonth)
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
      bsSheet.getRange(17, 2, bsLastRow - 16, 11).clearContent();
      bsSheet.getRange(17, 2, bsLastRow - 16, 11).setBorder(false, false, false, false, false, false);
      bsSheet.getRange(17, 2, bsLastRow - 16, 11).setBackground(null);
    }
    
    bsSheet.getRange("G15").setValue(timestamp);
    
    const bsUrl = baseUrl + "trial_bs?company_id=" + companyId + "&start_date=" + startDateStr + "&end_date=" + endDateStr;
    const bsRes = UrlFetchApp.fetch(bsUrl, options);
    const bsBalances = JSON.parse(bsRes.getContentText()).trial_bs.balances;
    
    const bsHeaders = ["分類", "勘定科目", "期首残高", "借方金額", "貸方金額", "期末残高", "", "", "プリペアコメント", "基礎資料又はfreee仕訳リンク", "レビュワーコメント"];
    bsSheet.getRange(16, 2, 1, bsHeaders.length).setValues([bsHeaders]);
    
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
          bsSheet.getRange(17 + i, 2, 1, 11).setBorder(true, false, false, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        }
      }
    }
  }
  
  // ===== PL（3期比較）=====
  const plSheet = ss.getSheetByName("PL");
  if (plSheet) {
    const plLastRow = plSheet.getLastRow();
    if (plLastRow >= 17) {
      plSheet.getRange(17, 2, plLastRow - 16, 11).clearContent();
      plSheet.getRange(17, 2, plLastRow - 16, 11).setBorder(false, false, false, false, false, false);
      plSheet.getRange(17, 2, plLastRow - 16, 11).setBackground(null);
    }
    
    plSheet.getRange("G15").setValue(timestamp);
    plSheet.getRange("F15").setValue(taxAccountingMethod);
    
    // 3期比較APIを使用
    const plUrl = baseUrl + "trial_pl_three_years" + params;
    const plRes = UrlFetchApp.fetch(plUrl, options);
    const plBalances = JSON.parse(plRes.getContentText()).trial_pl_three_years.balances;
    
    const plHeaders = ["分類", "勘定科目", periodLabels.twoYearsAgo, periodLabels.previous, periodLabels.current, "前年差額", "前年比", "", "プリペアコメント", "基礎資料又はfreee仕訳リンク", "レビュワーコメント"];
    plSheet.getRange(16, 2, 1, plHeaders.length).setValues([plHeaders]);
    
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
          plSheet.getRange(17 + i, 2, 1, 11).setBorder(true, false, false, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        }
      }
    }
  }
  
  // ===== CR（3期比較）=====
  const crSheet = ss.getSheetByName("CR");
  if (crSheet) {
    const crLastRow = crSheet.getLastRow();
    if (crLastRow >= 17) {
      crSheet.getRange(17, 2, crLastRow - 16, 11).clearContent();
      crSheet.getRange(17, 2, crLastRow - 16, 11).setBorder(false, false, false, false, false, false);
      crSheet.getRange(17, 2, crLastRow - 16, 11).setBackground(null);
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
          crSheet.getRange(16, 2, 1, crHeaders.length).setValues([crHeaders]);
          
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
  if (bsSheet) {
    bsSheet.getRange("D15").setValue(journalCount + "仕訳");
  }
  
  // ===== 区分別表 =====
  const bsBalancesForOrder = JSON.parse(UrlFetchApp.fetch(baseUrl + "trial_bs?company_id=" + companyId + "&start_date=" + startDateStr + "&end_date=" + endDateStr, options).getContentText()).trial_bs?.balances || [];
  const plBalancesForOrder = JSON.parse(UrlFetchApp.fetch(baseUrl + "trial_pl?company_id=" + companyId + "&start_date=" + startDateStr + "&end_date=" + endDateStr, options).getContentText()).trial_pl?.balances || [];
  const accountOrder = buildAccountOrder(bsBalancesForOrder, plBalancesForOrder);
  
  getTaxCategoryReportCore(ss, companyId, startDateStr, endDateStr, taxAccountingMethod, accountOrder, timestamp);
  
  // ===== 固定資産台帳 =====
  getFixedAssetsCore(ss, companyId, fiscalYear);
}

/**
 * 仕訳数を取得（reports.gs用）
 */
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
  if (accountItemsData.account_items) {
    accountItemsData.account_items.forEach(item => {
      accountItems[item.id] = item.name;
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
                     "&limit=" + limit + "&offset=" + offset;
    const dealsResponse = UrlFetchApp.fetch(dealsUrl, options);
    const dealsData = JSON.parse(dealsResponse.getContentText());
    
    if (!dealsData.deals || dealsData.deals.length === 0) {
      break;
    }
    
    dealsData.deals.forEach(deal => {
      if (deal.details) {
        deal.details.forEach(detail => {
          const accountName = accountItems[detail.account_item_id] || "";
          const taxCodeName = taxCodes[detail.tax_code] || "対象外";
          const amount = detail.amount || 0;
          const vat = detail.vat || 0;
          const isCredit = deal.type === "income";
          
          taxCategoryData.push({
            accountName: accountName,
            taxCodeName: taxCodeName,
            amount: isCredit ? -amount : amount,
            vat: isCredit ? -vat : vat
          });
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
    
    // データの最終行を特定
    const dataLastRow = 43 + outputArray.length;
    
    // 最終行の下に罫線（B～M列、12列分）
    sheet.getRange(dataLastRow, 2, 1, 12).setBorder(false, false, true, false, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    
    // データ最終行より下で、B列が空欄の行を一括削除
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
}

/**
 * 固定資産台帳を取得してシートに出力
 */
function getFixedAssetsCore(ss, companyId, fiscalYear) {
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
  const fixedAssetsUrl = "https://api.freee.co.jp/api/1/fixed_assets?company_id=" + companyId;
  const fixedAssetsResponse = UrlFetchApp.fetch(fixedAssetsUrl, options);
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
      asset.accumulated_depreciation || 0,
      asset.book_value || 0
    ]);
  });
  
  // シートをクリアして出力
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    sheet.getRange(5, 2, lastRow - 4, 7).clearContent();
  }
  
  if (outputData.length > 0) {
    sheet.getRange(5, 2, outputData.length, 7).setValues(outputData);
  }
  
  // タイムスタンプ
  sheet.getRange("J3").setValue(Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm") + "更新");
}
