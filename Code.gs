/**
 * 廣告行銷中心 - 後端邏輯
 * Ver 1.8 (支援跨年份動態讀取)
 */

// --- 設定區 ---
const API_KEY = '在此貼上你的API_KEY';  // 請記得填入你的 Google AI Studio API Key
const EXT_MOM_ID = '1_Lu2Hwgdl_ASWG78NfvTTLAPj6zCloFIAWmEewQfgZo';
const EXT_COST_RATIO_ID = '1Fptzp585jNOo42GGVhWnWs4Hjw2f2VZmkBhM3DgNIW0';

/**
 * 網頁進入點
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('廣告行銷中心')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 取得當前試算表中所有包含「廣告編排」的分頁名稱
 */
function getAvailableSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const validSheets = [];
    sheets.forEach(s => {
      if (s.getName().includes("廣告編排")) validSheets.push(s.getName());
    });
    // 建議：可以在這裡做個排序，讓最新的月份排在前面
    validSheets.sort((a, b) => b.localeCompare(a));
    return { sheets: validSheets };
  } catch (e) {
    return { error: "後端讀取錯誤: " + e.toString() };
  }
}

/**
 * 預算數值清理工具
 */
function parseBudget(value) {
  if (typeof value === 'number') return Math.round(value);
  if (!value) return 0;
  let str = value.toString();
  let match = str.match(/[\d,]+\.?\d*/); 
  if (!match) return 0;
  let cleanNum = match[0].replace(/,/g, '');
  let result = parseFloat(cleanNum);
  return isNaN(result) ? 0 : Math.round(result);
}

// --- 外部資料讀取區 ---

/**
 * 讀取外部 MOM 表格 (各平台營收)
 * 此表通常是連續的日期，不需改年份名稱
 */
function getExternalPlatformData(year, month) {
  try {
    const extSS = SpreadsheetApp.openById(EXT_MOM_ID);
    const sheet = extSS.getSheetByName("MOM(月)(更新)");
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    const headers = data[0]; 
    let colMap = {};
    headers.forEach((h, i) => colMap[h.toString().trim()] = i);

    let mStr = month < 10 ? "0" + month : "" + month;
    let targetDateStr = `${year}-${mStr}`;

    let result = { 'MOMO':0, 'PChome':0, 'Shopee':0, 'Official_BH':0, 'Official_How':0, 'Found':false };

    for (let i = 1; i < data.length; i++) {
      let rowDate = data[i][0]; 
      let rowDateStr = "";
      if (rowDate instanceof Date) {
        let y = rowDate.getFullYear();
        let m = rowDate.getMonth() + 1;
        rowDateStr = `${y}-${m < 10 ? "0"+m : m}`;
      } else {
        rowDateStr = String(rowDate).substring(0, 7);
      }

      if (rowDateStr === targetDateStr) {
        result['MOMO'] = parseBudget(data[i][colMap['momo']]);
        result['PChome'] = parseBudget(data[i][colMap['P購']]);
        result['Shopee'] = parseBudget(data[i][colMap['蝦皮']]);
        result['Official_BH'] = parseBudget(data[i][colMap['BH官網']]);
        result['Official_How'] = parseBudget(data[i][colMap['好好運動']]);
        result['Found'] = true;
        break;
      }
    }
    return result;
  } catch (e) { return null; }
}

/**
 * 讀取 BCG 表格 (銷量與佔比)
 * ★ V1.8 修正：根據年份動態抓取 "XX年廣告費佔比"
 */
function getBCGData(year, month) {
  try {
    const extSS = SpreadsheetApp.openById(EXT_COST_RATIO_ID);
    
    // 動態年份處理：2026 -> "26"
    let shortYear = (year % 100).toString();
    const sheetName = `${shortYear}年廣告費佔比`;
    
    const sheet = extSS.getSheetByName(sheetName);
    
    // 如果找不到該年份的分頁，回傳空物件，避免報錯
    if (!sheet) {
      console.log(`找不到 BCG 分頁: ${sheetName}`);
      return {};
    }

    const headerRange = sheet.getRange(1, 1, 3, sheet.getLastColumn());
    const headerValues = headerRange.getValues();
    const dateRow = headerValues[0];

    let targetStr = `${year}/${month}`; 
    let prevY = year;
    let prevM = month - 1;
    if (prevM === 0) { prevM = 12; prevY = year - 1; }
    let prevStr = `${prevY}/${prevM}`;

    let targetCol = -1;
    let prevCol = -1;

    for (let c = 0; c < dateRow.length; c++) {
      let cellVal = dateRow[c];
      let cellStr = "";
      if (cellVal instanceof Date) {
        cellStr = `${cellVal.getFullYear()}/${cellVal.getMonth() + 1}`;
      } else {
        cellStr = String(cellVal).trim();
      }

      if (cellStr === targetStr) targetCol = c;
      if (cellStr === prevStr) prevCol = c;
    }

    if (targetCol === -1) return {}; 

    const OFFSET_VOL = 1;   
    const OFFSET_SPEND = 4; 

    const allData = sheet.getDataRange().getValues();
    let salesMap = {};

    for (let i = 2; i < allData.length; i++) {
      let row = allData[i];
      let prodName = String(row[1]).trim(); 
      if (!prodName) continue;

      let vol = parseBudget(row[targetCol + OFFSET_VOL]);
      let spend = parseBudget(row[targetCol + OFFSET_SPEND]);
      
      let prevVol = 0;
      if (prevCol !== -1) {
        prevVol = parseBudget(row[prevCol + OFFSET_VOL]);
      }

      salesMap[prodName.toUpperCase()] = {
        name: prodName,
        vol: vol,
        spend: spend,
        prevVol: prevVol
      };
    }
    return salesMap;

  } catch (e) {
    return {}; 
  }
}

// --- 核心處理區 ---

function getPlatform(str) {
  let s = str.toUpperCase();
  if (s.includes('FB') || s.includes('FACEBOOK') || s.includes('IG') || s.includes('INSTAGRAM')) return 'FB';
  if (s.includes('GOOG') || s.includes('關鍵字') || s.includes('YOUTUBE')) return 'Google';
  return 'Other';
}

/**
 * 讀取本機營收表
 * ★ V1.8 修正：根據年份動態抓取 "XX年每月廣告費用明細"
 */
function getRevenueData(ss, targetMonth, yearFull) {
  // 動態年份處理：2026 -> "26"
  let shortYear = (yearFull % 100).toString();
  const sheetName = `${shortYear}年每月廣告費用明細`;
  
  const sheet = ss.getSheetByName(sheetName);
  // 如果找不到該年份的表，回傳 null
  if (!sheet) {
    console.log(`找不到營收分頁: ${sheetName}`);
    return null;
  }

  const data = sheet.getDataRange().getValues();
  
  let headerRowIdx = 0;
  for(let i=0; i<Math.min(data.length, 10); i++){
    if(data[i].includes('直營門市') && data[i].includes('總業績')) {
      headerRowIdx = i;
      break;
    }
  }

  const headers = data[headerRowIdx];
  let colMap = {};
  headers.forEach((h, i) => colMap[h.toString().trim()] = i);
  
  const idxStore = colMap['直營門市'] !== undefined ? colMap['直營門市'] : 1;
  const idxDist = colMap['達康業績'] !== undefined ? colMap['達康業績'] : 3;
  const idxEcom = colMap['直營電商'] !== undefined ? colMap['直營電商'] : 4;
  const idxTotal = colMap['總業績'] !== undefined ? colMap['總業績'] : 7;

  let result = { store: 0, distributor: 0, ecom: 0, total: 0 };

  for (let i = headerRowIdx + 1; i < data.length; i++) {
    let monthVal = data[i][0]; 
    let m = parseInt(monthVal.toString().match(/\d+/));
    if (m === targetMonth) {
      result.store = parseBudget(data[i][idxStore]);
      result.distributor = parseBudget(data[i][idxDist]);
      result.ecom = parseBudget(data[i][idxEcom]);
      result.total = parseBudget(data[i][idxTotal]);
      break;
    }
  }
  return result;
}

/**
 * 主要資料抓取流程
 */
function getSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { error: "找不到分頁" };
    
    // 1. 處理廣告資料
    const adData = processSheetRows(sheet);
    if (adData.error) return { error: adData.error }; 

    const match = sheetName.match(/(\d+)年(\d+)月/);
    let revenueData = { store: 0, distributor: 0, ecom: 0, total: 0 };
    let externalData = null;
    
    if (match) {
      let yearShort = parseInt(match[1]);
      let month = parseInt(match[2]);
      let yearFull = 2000 + yearShort; // 26 -> 2026

      // 2. 抓營收 (傳入 yearFull 以決定分頁名稱)
      let rev = getRevenueData(ss, month, yearFull);
      if (rev) revenueData = rev;

      // 3. 抓外部MOM
      externalData = getExternalPlatformData(yearFull, month);
      
      // 4. 抓BCG (銷量與花費)
      let salesMap = getBCGData(yearFull, month);
      
      // 5. 整合比對
      adData.allProducts.forEach(prod => {
        let adNameClean = prod.name.split('(')[0].trim().toUpperCase(); 
        
        let totalVol = 0;
        let totalSpend = 0;
        let totalPrevVol = 0;
        let matchCount = 0;

        Object.keys(salesMap).forEach(k => {
          let salesNameClean = k.toUpperCase();
          let longer = adNameClean.length > salesNameClean.length ? adNameClean : salesNameClean;
          let shorter = adNameClean.length > salesNameClean.length ? salesNameClean : adNameClean;

          let escaped = shorter.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          let re = new RegExp('(^|[^A-Z0-9])' + escaped + '([^A-Z0-9]|$)', 'i');

          if (re.test(longer)) {
            totalVol += salesMap[k].vol;
            totalSpend += salesMap[k].spend;
            totalPrevVol += salesMap[k].prevVol;
            matchCount++;
          }
        });
        
        prod.salesVolume = totalVol;
        prod.extSpend = totalSpend;
        prod.salesMoM = totalPrevVol > 0 ? (totalVol - totalPrevVol) / totalPrevVol : 0;
        prod.hasExtData = (matchCount > 0);
      });
    }

    return { sheetName, adData, revenueData, externalData };
  } catch (e) {
    return { error: "處理失敗: " + e.toString() };
  }
}

function processSheetRows(sheet) {
  const data = sheet.getDataRange().getValues();
  
  // 1. 強制指定第 1 列 (Row 1, Index 0) 是標題列
  const headers = data[0];
  const colMap = {};
  
  // 建立欄位地圖
  headers.forEach((h, i) => {
    let cleanH = h.toString().trim();
    colMap[cleanH] = i; 

    if (cleanH === "總預算" || cleanH === "合計" || cleanH === "總金額") {
      colMap['總預算'] = i; 
    }
    
    if (colMap['總預算'] === undefined && cleanH.includes("預算")) {
       if (!cleanH.includes("日") && !cleanH.includes("平日") && !cleanH.includes("假日")) {
          colMap['總預算'] = i;
       }
    }
    
    if (cleanH.includes("商品") || cleanH.includes("品項") || cleanH.includes("品名")) colMap['商品'] = i;
    if (cleanH.includes("目的地") || cleanH.includes("平台") || cleanH.includes("通路")) colMap['目的地'] = i;
    if (cleanH.includes("實際") && cleanH.includes("花費")) colMap['實際花費'] = i;
    if (cleanH.includes("廣告目標")) colMap['廣告目標'] = i;
    if (cleanH.includes("廣告帳號")) colMap['廣告帳號'] = i;
  });

  let stats = { 
    plannedBudget: 0, actualSpend: 0, 
    spend_ecom: 0, spend_brand: 0, 
    destinations: {}, objectives: {}, 
    products: {},
    platforms: { 'FB': 0, 'Google': 0, 'Other': 0 }
  };

  let hasActualCol = (colMap['實際花費'] !== undefined);

  for (let i = 2; i < data.length; i++) {
    let row = data[i];
    let rowStr = row.join("").toLowerCase();
    if (rowStr.includes("script.google.com") || rowStr.includes("連結") || rowStr.includes("app") || rowStr.includes("http")) {
      continue; 
    }

    let budget = parseBudget(row[colMap['總預算']]);
    if (budget === 0) continue; 

    stats.plannedBudget += budget;
    if (hasActualCol) stats.actualSpend += parseBudget(row[colMap['實際花費']]);

    let rawProd = String(row[colMap['商品']] || "未命名");
    let prods = rawProd.split('\n').map(s=>s.trim()).filter(s=>s && !s.includes('旗艦館') && !s.includes('輪播') && !s.includes('廣告'));
    if (prods.length === 0) prods = [rawProd.split('\n')[0]];
    let splitBudget = Math.round(budget / prods.length);
    
    let dest = String(row[colMap['目的地']] || "其他").trim();
    let obj = String(row[colMap['廣告目標']] || "未設定").trim();
    let checkStr = String(row[colMap['廣告帳號']] || "") + obj + dest;
    let platform = getPlatform(checkStr);
    stats.platforms[platform] += budget;

    let destKey = "其他";
    if (dest.includes('商用')) { destKey = 'BH商用'; stats.spend_brand += budget; }
    else if (dest.includes('i-BH')) { destKey = 'i-BH'; stats.spend_brand += budget; }
    else if (dest.includes('好好')) { destKey = '好好運動'; stats.spend_ecom += budget; }
    else if (dest.includes('MOMO')||dest.includes('momo')) { destKey = 'MOMO'; stats.spend_ecom += budget; }
    else if (dest.includes('蝦皮')) { destKey = '蝦皮'; stats.spend_ecom += budget; }
    else if (dest.includes('PC')||dest.includes('PChome')) { destKey = 'PChome'; stats.spend_ecom += budget; }
    else { stats.spend_ecom += budget; }

    stats.destinations[destKey] = (stats.destinations[destKey] || 0) + budget;
    stats.objectives[obj] = (stats.objectives[obj] || 0) + budget;

    prods.forEach(p => {
      let name = p.split('(')[0].trim();
      if (!stats.products[name]) stats.products[name] = { 
        total: 0, 
        breakdown: {}, 
        objectiveBreakdown: {},
        platformBreakdown: {'FB':0, 'Google':0, 'Other':0} 
      };
      stats.products[name].total += splitBudget;
      stats.products[name].breakdown[destKey] = (stats.products[name].breakdown[destKey] || 0) + splitBudget;
      stats.products[name].objectiveBreakdown[obj] = (stats.products[name].objectiveBreakdown[obj] || 0) + splitBudget;
      stats.products[name].platformBreakdown[platform] += splitBudget;
    });
  }

  stats.allProducts = Object.entries(stats.products)
    .map(([k,v]) => ({
      name:k, total:v.total, 
      breakdown:v.breakdown, 
      objectiveBreakdown:v.objectiveBreakdown,
      platformBreakdown:v.platformBreakdown
    }))
    .sort((a,b) => b.total - a.total);
    
  return stats;
}

function askGeminiStrategy(sheetName) {
  try {
    const result = getSheetData(sheetName);
    if (result.error) return "數據讀取失敗|||請檢查後端";
    
    const ad = result.adData;
    const rev = result.revenueData;
    
    let topProdStr = ad.allProducts.slice(0, 5).map(p => {
       let vol = p.salesVolume ? `${p.salesVolume}台` : "N/A";
       let pct = Math.round((p.total / ad.plannedBudget)*100);
       return `- ${p.name}: 預算$${p.total} (${pct}%) | 銷量:${vol}`;
    }).join('\n');

    let fbTotal = ad.platforms['FB'];
    let googTotal = ad.platforms['Google'];
    let fbPct = Math.round((fbTotal/ad.plannedBudget)*100);
    let googPct = Math.round((googTotal/ad.plannedBudget)*100);

    const roasTotal = rev.total > 0 ? (rev.total / ad.plannedBudget).toFixed(1) : "0.0";
    const budgetUsage = Math.round((ad.actualSpend / ad.plannedBudget) * 100);

    const prompt = `
    你是一位行銷數據分析師。請根據以下數據輸出兩段報告，用 "|||" 分隔。
    使用 HTML <b> 標籤標記重點。

    【數據】
    [總覽] 預算:$${ad.plannedBudget}, 花費:$${ad.actualSpend} (${budgetUsage}%), 業績:$${rev.total}, ROAS:${roasTotal}
    [平台] FB:$${fbTotal} (${fbPct}%), Google:$${googTotal} (${googPct}%)
    [主力產品與銷量]
    ${topProdStr}

    【第一段：投放架構摘要】
    條列本月重點。
    <ul>
    <li><b>本月重點</b>：前五大產品佔比 <b>XX%</b>。重心為 <b>產品A</b> (預算$XX / 銷量XX台)。</li>
    <li><b>平台配置</b>：FB <b>$金額 (XX%)</b> / Google <b>$金額 (XX%)</b>。</li>
    <li><b>細項</b>：<br>${topProdStr.replace(/\n/g, '<br>')}</li>
    </ul>

    【第二段：成效診斷】
    針對 ROAS 與 銷量 進行評價。
    <ul>
    <li><b>預算執行</b>：${budgetUsage}%。</li>
    <li><b>營收效率</b>：全站 ROAS <b>${roasTotal}</b>。</li>
    <li><b>建議</b>：針對高花費低銷量的產品給予建議。</li>
    </ul>
    `;

    const modelVersion = "gemini-2.5-flash"; 
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelVersion}:generateContent?key=${API_KEY}`;
    
    const payload = { "contents": [{ "parts": [{"text": prompt}] }] };
    const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (!json.candidates) return "AI 無回應|||稍後再試";
    return json.candidates[0].content.parts[0].text.replace(/```html/g, "").replace(/```/g, "");

  } catch (e) {
    return "系統錯誤|||" + e.toString();
  }
}
