/**
 * freeeから事業所リストを取得
 */
function fetchCompaniesFromFreee() {
  const accessToken = getService().getAccessToken();
  const url = 'https://api.freee.co.jp/api/1/companies';
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + accessToken }
  });
  const json = JSON.parse(res.getContentText());
  return json.companies || [];
}

/**
 * 事業所の詳細情報を取得
 */
function getCompanyDetails(companyId, accessToken) {
  const response = UrlFetchApp.fetch(
    `https://api.freee.co.jp/api/1/companies/${companyId}`,
    { method: 'get', headers: { Authorization: 'Bearer ' + accessToken } }
  );
  const company = JSON.parse(response.getContentText()).company;

  if (!company.fiscal_years || company.fiscal_years.length === 0) {
    throw new Error("会計年度情報が取得できませんでした。");
  }

  const today = new Date();
  const fiscalYears = company.fiscal_years
    .filter(y => y.start_date && y.end_date)
    .sort((a, b) => new Date(b.start_date) - new Date(a.start_date));

  let currentFy = null;
  for (const fy of fiscalYears) {
    const start = new Date(fy.start_date);
    const end = new Date(fy.end_date);
    if (start <= today && today <= end) {
      currentFy = fy;
      break;
    }
  }

  if (!currentFy) {
    currentFy = fiscalYears[0];
  }

  const endDate = new Date(currentFy.end_date);
  const year = endDate.getFullYear();
  const month = endDate.getMonth() + 1;

  let address = "";
  if (company.prefecture_code) {
    const prefectures = ["", "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県", "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県", "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県", "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県", "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"];
    address = prefectures[company.prefecture_code] || "";
  }
  address += company.street_name1 || "";
  address += company.street_name2 || "";

  return {
    companyName: company.display_name || company.name,
    companyId: companyId,
    periodLabel: `${year}年${month}月期`,
    startDate: currentFy.start_date,
    endDate: currentFy.end_date,
    address: address,
    headName: company.head_name || "",
    taxAccountingMethod: company.tax_at_source_calc_type === 0 ? "税込経理" : "税抜経理",
    taxType: getTaxTypeFromCompany(company)
  };
}

/**
 * 会社情報から課税方式を取得
 */
function getTaxTypeFromCompany(company) {
  const taxMethod = company.tax_method_of_paying_tax;
  if (taxMethod === 0) return "免税事業者";
  if (taxMethod === 1) return "原則課税";
  if (taxMethod === 2) return "簡易課税";
  return "不明";
}

/**
 * 課税方式を取得
 */
function getTaxType(companyId, accessToken) {
  const response = UrlFetchApp.fetch(
    `https://api.freee.co.jp/api/1/companies/${companyId}`,
    { method: 'get', headers: { Authorization: 'Bearer ' + accessToken } }
  );
  const company = JSON.parse(response.getContentText()).company;
  return getTaxTypeFromCompany(company);
}

/**
 * 仕訳数を取得
 */
function getJournalCount(companyId, accessToken, startDateStr, endDateStr) {
  let totalCount = 0;
  
  const dealsUrl = `https://api.freee.co.jp/api/1/deals?company_id=${companyId}&start_issue_date=${startDateStr}&end_issue_date=${endDateStr}&limit=1`;
  const dealsRes = UrlFetchApp.fetch(dealsUrl, {
    method: "get",
    headers: { "Authorization": `Bearer ${accessToken}` },
    muteHttpExceptions: true
  });
  if (dealsRes.getResponseCode() === 200) {
    const dealsData = JSON.parse(dealsRes.getContentText());
    totalCount += dealsData.meta?.total_count || 0;
  }
  
  const journalsUrl = `https://api.freee.co.jp/api/1/manual_journals?company_id=${companyId}&start_issue_date=${startDateStr}&end_issue_date=${endDateStr}&limit=1`;
  const journalsRes = UrlFetchApp.fetch(journalsUrl, {
    method: "get",
    headers: { "Authorization": `Bearer ${accessToken}` },
    muteHttpExceptions: true
  });
  if (journalsRes.getResponseCode() === 200) {
    const journalsData = JSON.parse(journalsRes.getContentText());
    totalCount += journalsData.meta?.total_count || 0;
  }
  
  return totalCount;
}

/**
 * 期間ラベルを取得
 */
function getPeriodLabels(companyId, accessToken, endDateStr) {
  companyId = parseInt(companyId, 10);
  if (isNaN(companyId)) throw new Error("companyIdが数値に変換できません。");
  
  const endDate = new Date(endDateStr);
  
  const response = UrlFetchApp.fetch(`https://api.freee.co.jp/api/1/companies/${companyId}`, { method: 'get', headers: { Authorization: 'Bearer ' + accessToken } });
  const company = JSON.parse(response.getContentText()).company;
  
  if (!company.fiscal_years || company.fiscal_years.length === 0) {
    throw new Error("会計年度情報が取得できませんでした。");
  }
  
  const fiscalYears = company.fiscal_years.filter(y => y.start_date && y.end_date).sort((a, b) => new Date(a.start_date) - new Date(b.start_date));
  
  const idx = fiscalYears.findIndex(fy => {
    const start = new Date(fy.start_date);
    const end = new Date(fy.end_date);
    return start <= endDate && endDate <= end;
  });
  
  if (idx < 0) throw new Error("指定した終了日を含む会計年度が見つかりません。");
  
  const current = fiscalYears[idx];
  const previous = fiscalYears[idx - 1];
  const twoYearsAgo = fiscalYears[idx - 2];
  
  if (!previous) throw new Error("前期の会計年度情報が存在しません。");
  
  const fmt = (y, m) => `${String(y).slice(-2)}/${String(m).padStart(2, '0')}`;
  
  const startY = new Date(current.start_date).getFullYear();
  const startM = new Date(current.start_date).getMonth() + 1;
  const endY = endDate.getFullYear();
  const endM = endDate.getMonth() + 1;
  
  return {
    startDateStr: current.start_date,
    prevStartDateStr: previous.start_date,
    prevEndDateStr: previous.end_date,
    twoYearsAgoStartDateStr: twoYearsAgo ? twoYearsAgo.start_date : null,
    twoYearsAgoEndDateStr: twoYearsAgo ? twoYearsAgo.end_date : null,
    thisPeriod: `${fmt(startY, startM)} - ${fmt(endY, endM)}`,
    lastPeriod: `${fmt(startY - 1, startM)} - ${fmt(endM >= startM ? endY - 1 : endY, endM)}`,
    twoYearsAgoPeriod: twoYearsAgo ? `${fmt(startY - 2, startM)} - ${fmt(endM >= startM ? endY - 2 : endY - 1, endM)}` : null
  };
}

