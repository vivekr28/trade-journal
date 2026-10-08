// ============================================================
// GOOGLE APPS SCRIPT — Capital & Trade Tracker Backend v4
// ============================================================
// Supports partial exits via Trade_Legs.
// Safe to re-run setupSheet() anytime — non-destructive.
// ============================================================

// ==================== MIGRATION ====================
// Run once to convert old P&L_% values (stored as *100) to decimals.
// After running, the column stores true decimals (e.g. -0.0524 = -5.24%).
function migratePnlPctToDecimal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Trades');
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var col = 27; // AA = P&L_%
  var range = sheet.getRange(2, col, lastRow - 1, 1);
  var vals = range.getValues();
  var changed = 0;
  for (var i = 0; i < vals.length; i++) {
    var v = Number(vals[i][0]);
    if (v === 0 || isNaN(v)) continue;
    // Old format: values like -5.24 meaning -5.24%. Convert to -0.0524.
    // New decimal format: values like -0.0524. These are |v| < 1 for trades under ±100%.
    // Heuristic: if |v| >= 1, it's old format (very unlikely a single trade has ±100%+ P&L)
    if (Math.abs(v) >= 1) {
      vals[i][0] = v / 100;
      changed++;
    }
  }
  if (changed > 0) {
    range.setValues(vals);
    range.setNumberFormat('0.00%');
  }
  Logger.log('Migrated ' + changed + ' rows to decimal P&L_%');
}

// Run once to backfill ROCE_% for closed trades that don't have it.
// Uses current capital — not exact close-time capital, but better than 0.
function migrateBackfillRoce() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Trades');
  var summary = ss.getSheetByName('Summary');
  if (!sheet || !summary) return;
  var capital = summary.getRange('B2').getValue() || 0;
  if (capital <= 0) { Logger.log('Capital is 0, cannot compute ROCE'); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  // Col W(23)=Status, Col Z(26)=P&L, Col AJ(36)=ROCE_%
  var statuses = sheet.getRange(2, 23, lastRow - 1, 1).getValues();
  var pnls = sheet.getRange(2, 26, lastRow - 1, 1).getValues();
  var roces = sheet.getRange(2, 36, lastRow - 1, 1).getValues();
  var changed = 0;
  for (var i = 0; i < statuses.length; i++) {
    if (statuses[i][0] === 'Closed' && (!roces[i][0] || roces[i][0] === 0)) {
      var pnl = Number(pnls[i][0]) || 0;
      roces[i][0] = pnl / capital;
      changed++;
    }
  }
  if (changed > 0) {
    sheet.getRange(2, 36, lastRow - 1, 1).setValues(roces);
    sheet.getRange(2, 36, lastRow - 1, 1).setNumberFormat('0.00%');
  }
  Logger.log('Backfilled ROCE_% for ' + changed + ' closed trades using capital=' + capital);
}

// Run once to backfill Capital column (AK=37) for existing trades that don't have it.
// Sets ₹1,00,000 for all existing trades without a capital value.
function migrateBackfillCapital() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Trades');
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  // Ensure column AK exists
  var maxCol = sheet.getMaxColumns();
  if (maxCol < 37) {
    sheet.insertColumnsAfter(maxCol, 37 - maxCol);
  }
  // Add header if missing
  var header = sheet.getRange(1, 37).getValue();
  if (!header || header === '') {
    sheet.getRange(1, 37).setValue('Capital');
    sheet.getRange(1, 37).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
  }
  var capitals = sheet.getRange(2, 37, lastRow - 1, 1).getValues();
  var changed = 0;
  for (var i = 0; i < capitals.length; i++) {
    if (!capitals[i][0] || capitals[i][0] === 0 || capitals[i][0] === '') {
      capitals[i][0] = 100000;
      changed++;
    }
  }
  if (changed > 0) {
    sheet.getRange(2, 37, lastRow - 1, 1).setValues(capitals);
    sheet.getRange(2, 37, lastRow - 1, 1).setNumberFormat('₹#,##0.00');
  }
  Logger.log('Backfilled Capital for ' + changed + ' trades with ₹1,00,000');
}

// ==================== HELPERS ====================
// Google Sheets getValues() returns Date objects for date/time cells.
// JSON.stringify turns them into ISO strings like "2026-08-09T18:30:00.000Z".
// These helpers format them as local strings before returning to frontend.
function fmtDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}
function fmtTime_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  return String(v);
}

// ==================== SETUP (NON-DESTRUCTIVE) ====================
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];
  function ensureSheet(name, headers, widths) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) { sheet = ss.insertSheet(name); log.push('Created: ' + name); }
    var maxCol = sheet.getMaxColumns();
    if (maxCol < headers.length) {
      sheet.insertColumnsAfter(maxCol, headers.length - maxCol);
      log.push('Expanded columns: ' + name + ' (' + maxCol + ' → ' + headers.length + ')');
    }
    const first = sheet.getRange(1,1).getValue();
    if (!first || first === '') {
      sheet.getRange(1,1,1,headers.length).setValues([headers]);
      sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      log.push('Headers added: ' + name);
    } else {
      var currentHeaders = sheet.getRange(1,1,1,maxCol).getValues()[0];
      if (currentHeaders.length < headers.length) {
        var newCols = headers.slice(currentHeaders.length);
        sheet.getRange(1, currentHeaders.length+1, 1, newCols.length).setValues([newCols]);
        sheet.getRange(1, currentHeaders.length+1, 1, newCols.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
        log.push('New columns added to: ' + name + ' (+' + newCols.length + ')');
      }
    }
    if (widths) { for (var i=0;i<widths.length;i++) sheet.setColumnWidth(i+1,widths[i]); }
    return sheet;
  }

  // Capital_Transactions
  ensureSheet('Capital_Transactions',
    ['Txn_ID','Date','Time','Type','Amount','Running_Balance','Timestamp'],
    [120,120,100,100,130,150,180]);

  // Config
  let cfgSheet = ss.getSheetByName('Config');
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet('Config');
    cfgSheet.getRange(1,1,1,3).setValues([['Key','Value','Description']]);
    cfgSheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    cfgSheet.getRange(2,1,23,3).setValues([
      ['Streak_Threshold', 7, 'Consecutive losses before size-cut warning'],
      ['MIS_Brok_Buy', 0.03, 'MIS Brokerage % on Buy'],
      ['MIS_Brok_Buy_Cap', 20, 'MIS Brokerage cap ₹ on Buy (0=no cap)'],
      ['MIS_Brok_Sell', 0.03, 'MIS Brokerage % on Sell'],
      ['MIS_Brok_Sell_Cap', 20, 'MIS Brokerage cap ₹ on Sell (0=no cap)'],
      ['MIS_STT_Buy', 0, 'MIS STT % on Buy (0=N/A)'],
      ['MIS_STT_Buy_Cap', 0, 'MIS STT cap ₹ on Buy (0=no cap)'],
      ['MIS_STT_Sell', 0.025, 'MIS STT % on Sell'],
      ['MIS_STT_Sell_Cap', 0, 'MIS STT cap ₹ on Sell (0=no cap)'],
      ['CNC_Brok_Buy', 0, 'CNC Brokerage % on Buy (0=N/A)'],
      ['CNC_Brok_Buy_Cap', 0, 'CNC Brokerage cap ₹ on Buy (0=no cap)'],
      ['CNC_Brok_Sell', 0, 'CNC Brokerage % on Sell (0=N/A)'],
      ['CNC_Brok_Sell_Cap', 0, 'CNC Brokerage cap ₹ on Sell (0=no cap)'],
      ['CNC_STT_Buy', 0.1, 'CNC STT % on Buy'],
      ['CNC_STT_Buy_Cap', 0, 'CNC STT cap ₹ on Buy (0=no cap)'],
      ['CNC_STT_Sell', 0.1, 'CNC STT % on Sell'],
      ['CNC_STT_Sell_Cap', 0, 'CNC STT cap ₹ on Sell (0=no cap)'],
      ['Txn_Charge_Pct', 0.00307, 'Exchange transaction charge % (NSE equity, both sides)'],
      ['SEBI_Per_Crore', 10, 'SEBI turnover fee ₹ per crore (both sides)'],
      ['GST_Pct', 18, 'GST % on brokerage + SEBI + exchange charges'],
      ['MIS_Stamp_Buy', 0.003, 'MIS stamp duty % (buy side only)'],
      ['CNC_Stamp_Buy', 0.015, 'CNC stamp duty % (buy side only)'],
      ['CNC_DP_Sell', 15.34, 'CNC DP charge ₹ flat per sell leg']
    ]);
    cfgSheet.setColumnWidth(1,180); cfgSheet.setColumnWidth(2,100); cfgSheet.setColumnWidth(3,300);
    log.push('Created: Config with defaults');
  }

  // Trades (must be created BEFORE Summary so formulas resolve)
  // A-P: Trade info + setup, Q-T: Calculated, U: Timestamp
  // V: Remaining_Qty, W: Status, X: Exit_Date, Y: Avg_Exit_Price, Z: P&L, AA: P&L_%, AB: Result
  // AC: LF_Rise, AD: Entry_Chart_URL, AE: Buy_Charges, AF: Net_Buy_Price, AG: Current_SL, AH: Entry_Notes
  // AI: Market_Env, AJ: ROCE_%, AK: Capital, AL: Entry_Chart_URL_2, AM: Entry_Chart_URL_3
  ensureSheet('Trades',
    ['Trade_ID','Type','Contract','Symbol','Lots','Lot_Size','Buy_Date','Buy_Time','Buy_Price','Initial_SL',
     'Setup','RVol','PreMkt_AD_Ratio','Cause_%_Move','Consol_%_Move','MA_Undercut',
     'Invested_Amount','Position_Size_%','SL_%','Risk_of_Capital_%','Timestamp',
     'Remaining_Qty','Status',
     'Exit_Date','Avg_Exit_Price','P&L','P&L_%','Result',
     'LF_Rise','Entry_Chart_URL','Buy_Charges','Net_Buy_Price','Current_SL','Entry_Notes','Market_Env','ROCE_%','Capital',
     'Entry_Chart_URL_2','Entry_Chart_URL_3'],
    [130,50,60,100,50,70,100,80,100,100,80,60,80,80,80,80,130,90,70,100,160,
     90,70,100,110,100,70,60,80,250,110,110,100,300,120,130,250,250]);

  // Trade_Legs — col L: Exit_Charges, col M: Net_Exit_Price, col N: Exit_Notes
  ensureSheet('Trade_Legs',
    ['Leg_ID','Trade_ID','Date','Time','Action','Price','Quantity','Amount','Running_Remaining','Timestamp','Exit_Chart_URL','Exit_Charges','Net_Exit_Price','Exit_Notes'],
    [120,130,100,80,80,100,80,120,120,180,250,110,110,300]);

  // Summary (formulas only — always safe to refresh; placed AFTER Trades so refs resolve)
  let summarySheet = ss.getSheetByName('Summary');
  if (!summarySheet) { summarySheet = ss.insertSheet('Summary'); log.push('Created: Summary'); }
  summarySheet.clear();
  const sd = [
    ['Field','Value','Notes'],
    ['Current_Capital','=SUMPRODUCT((Capital_Transactions!D2:D5000="Added")*Capital_Transactions!E2:E5000)-SUMPRODUCT((Capital_Transactions!D2:D5000="Withdrawn")*Capital_Transactions!E2:E5000)','Auto-calculated'],
    ['Total_Added','=SUMPRODUCT((Capital_Transactions!D2:D5000="Added")*Capital_Transactions!E2:E5000)',''],
    ['Total_Withdrawn','=SUMPRODUCT((Capital_Transactions!D2:D5000="Withdrawn")*Capital_Transactions!E2:E5000)',''],
    ['Transaction_Count','=COUNTA(Capital_Transactions!A2:A5000)',''],
    ['Last_Txn_Date','=IF(COUNTA(Capital_Transactions!G2:G5000)>0,MAX(Capital_Transactions!G2:G5000),"None")',''],
    ['','',''],
    ['--- TRADE SUMMARY ---','',''],
    ['Total_Invested','=SUMPRODUCT((Trades!W2:W5000="Open")*Trades!Q2:Q5000)','Open trades invested amount'],
    ['Invested_%','=IF(B2>0,B9/B2*100,0)','% of capital deployed'],
    ['Open_Positions','=COUNTIF(Trades!W2:W5000,"Open")','All open trades count'],
    ['Unique_Open','=IFERROR(SUMPRODUCT((Trades!W2:W5000="Open")/COUNTIFS(Trades!D2:D5000,Trades!D2:D5000,Trades!W2:W5000,"Open",Trades!D2:D5000,"<>")),0)','Unique symbols open'],
    ['Total_Trades','=COUNTA(Trades!A2:A5000)','All-time'],
    ['','',''],
    ['--- PERFORMANCE ---','',''],
    ['Realised_ROCE','=IF(B2>0,SUMPRODUCT((Trades!W2:W5000="Closed")*(YEAR(Trades!X2:X5000)=YEAR(TODAY()))*(MONTH(Trades!X2:X5000)=MONTH(TODAY()))*Trades!Z2:Z5000)/B2*100,0)','This month closed P&L / Capital'],
    ['Unrealised_ROCE','=0','Placeholder — needs live prices'],
    ['','',''],
    ['--- STREAK ---','',''],
    ['Losing_Streak','=0','Calculated by script from recent exits'],
    ['Consecutive_Wins','=0','Calculated by script from recent exits'],
  ];
  summarySheet.getRange(1,1,sd.length,3).setValues(sd);
  summarySheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
  summarySheet.setColumnWidth(1,180); summarySheet.setColumnWidth(2,300); summarySheet.setColumnWidth(3,250);
  summarySheet.getRange('B2:B4').setNumberFormat('₹#,##0.00');
  summarySheet.getRange('B9').setNumberFormat('₹#,##0.00');
  log.push('Refreshed: Summary formulas');

  // Remove default Sheet1
  try { var ds=ss.getSheetByName('Sheet1'); if(ds&&ss.getSheets().length>1) ss.deleteSheet(ds); } catch(e){}
  SpreadsheetApp.getUi().alert('✅ Setup complete!\n\n'+log.join('\n')+'\n\nExisting data preserved.');
}

// ==================== WEB APP ====================
// Actions that read-modify-write shared rows (Remaining_Qty, Status, Trade_ID allocation, etc.)
// must be serialized across concurrent requests — see doPost's script lock below.
var MUTATING_ACTIONS = ['updateCapital','addNewTrade','addExitLeg','updateConfig','updateTradeSL','updateTradeFields'];

function doPost(e) {
  var lock = null;
  try {
    const data = JSON.parse(e.postData.contents);
    let result;

    if (MUTATING_ACTIONS.indexOf(data.action) !== -1) {
      lock = LockService.getScriptLock();
      if (!lock.tryLock(30000)) {
        return ContentService.createTextOutput(JSON.stringify({error:'Server busy, please retry'})).setMimeType(ContentService.MimeType.JSON);
      }
    }

    switch(data.action) {
      case 'getCapital': result=getCapital(); break;
      case 'updateCapital': result=updateCapital(data.amount,data.status,data.date); break;
      case 'getTransactions': result=getTransactions(); break;
      case 'addNewTrade': result=addNewTrade(data); break;
      case 'addExitLeg': result=addExitLeg(data); break;
      case 'getExitLegs': result=getExitLegs(data.tradeId); break;
      case 'getOpenTrades': result=getOpenTrades(); break;
      case 'getAllTrades': result=getAllTrades(); break;
      case 'getRibbonData': result=getRibbonData(); break;
      case 'updateConfig': result=updateConfig(data.key,data.value); break;
      case 'getInvested': result=getInvested(); break;
      case 'uploadChart': result=uploadChart(data.imageData,data.fileName,data.mimeType); break;
      case 'updateTradeSL': result=updateTradeSL(data.tradeId,data.newSL); break;
      case 'updateTradeFields': result=updateTradeFields(data); break;
      default: result={error:'Unknown action: '+data.action};
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error:err.message})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (lock) lock.releaseLock();
  }
}
function doGet(e) {
  const a = (e.parameter||{}).action;
  let r;
  switch(a) {
    case 'getRibbonData': r=getRibbonData(); break;
    case 'getCapital': r=getCapital(); break;
    default: r={status:'ok',version:'v4'};
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

// ==================== CONFIG ====================
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  if (!sheet) return {Streak_Threshold:7, Intraday_Charges:0.04, Delivery_Charges:0.10};
  const last = sheet.getLastRow();
  if (last < 2) return {Streak_Threshold:7, Intraday_Charges:0.04, Delivery_Charges:0.10};
  const data = sheet.getRange(2,1,last-1,2).getValues();
  const cfg = {};
  data.forEach(function(row) { if(row[0]) cfg[row[0]] = row[1]; });
  return cfg;
}
function updateConfig(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  if (!sheet) return {error:'Config sheet not found'};
  const last = sheet.getLastRow();
  for (var i=2; i<=last; i++) {
    if (sheet.getRange(i,1).getValue() === key) {
      sheet.getRange(i,2).setValue(value);
      return {success:true};
    }
  }
  sheet.getRange(last+1,1,1,2).setValues([[key, value]]);
  return {success:true};
}

// ==================== RIBBON DATA ====================
function getRibbonData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName('Summary');
  const config = getConfig();
  const capital = summary ? (summary.getRange('B2').getValue()||0) : 0;
  const invested = summary ? (summary.getRange('B9').getValue()||0) : 0;
  const investedPct = summary ? (summary.getRange('B10').getValue()||0) : 0;
  const openPositions = summary ? (summary.getRange('B11').getValue()||0) : 0;
  const streaks = calcStreaks();

  // Calculate realisedROCE, uniqueOpen, unrealisedROCE from Trades data directly
  // (avoids broken Summary sheet formulas)
  var realisedROCE = 0;
  var uniqueOpen = 0;
  var unrealisedROCE = 0;
  var totalRealisedPnl = 0;
  var totalPartialRealisedPnl = 0;
  var tradesSheet = ss.getSheetByName('Trades');
  var legsSheet = ss.getSheetByName('Trade_Legs');
  if (tradesSheet) {
    var last = tradesSheet.getLastRow();
    if (last >= 2) {
      var data = tradesSheet.getRange(2,1,last-1,39).getValues();

      // Build tradeId -> [{qty, netExitPrice}] from already-booked exit legs, so a
      // partial exit on a still-Open trade counts toward realised P&L immediately
      // instead of waiting for the trade to fully close.
      var partialLegsByTrade = {};
      if (legsSheet) {
        var legsLast = legsSheet.getLastRow();
        if (legsLast >= 2) {
          var legsData = legsSheet.getRange(2,1,legsLast-1,14).getValues();
          for (var li = 0; li < legsData.length; li++) {
            var lg = legsData[li];
            if (lg[4] !== 'Sell' || !lg[1]) continue; // E: Side, B: Trade_Id
            var tId = lg[1];
            if (!partialLegsByTrade[tId]) partialLegsByTrade[tId] = [];
            partialLegsByTrade[tId].push({ qty: lg[6], netExitPrice: lg[12] }); // G: Quantity, M: Net_Exit_Price
          }
        }
      }

      var openSymbols = {};
      var totalUnrealisedPnl = 0;
      totalRealisedPnl = 0;
      totalPartialRealisedPnl = 0;
      for (var i = 0; i < data.length; i++) {
        var r = data[i];
        if (!r[0]) continue; // skip empty rows
        var status = r[22]; // W: Status
        if (status === 'Closed') {
          totalRealisedPnl += (r[25] || 0); // Z: P&L
        }
        if (status === 'Open' && r[3]) {
          openSymbols[r[3]] = true;
          var netBuyPrice = r[31] || r[8];  // AF: Net_Buy_Price, fallback to I: Buy_Price
          var currentSL = r[32] || r[9];    // AG: Current_SL, fallback to J: Initial_SL
          var remainingQty = r[21];          // V: Remaining_Qty
          var tradeType = r[1];              // B: Type (L or S)
          if (tradeType === 'L') {
            totalUnrealisedPnl += (currentSL - netBuyPrice) * remainingQty;
          } else {
            totalUnrealisedPnl += (netBuyPrice - currentSL) * remainingQty;
          }
          // Realised P&L already booked on partial exit legs of this still-open trade
          var pLegs = partialLegsByTrade[r[0]];
          if (pLegs) {
            for (var pj = 0; pj < pLegs.length; pj++) {
              var pl = pLegs[pj];
              var legPnl = tradeType === 'L'
                ? (pl.netExitPrice - netBuyPrice) * pl.qty
                : (netBuyPrice - pl.netExitPrice) * pl.qty;
              totalPartialRealisedPnl += legPnl;
            }
          }
        }
      }
      totalRealisedPnl += totalPartialRealisedPnl;
      realisedROCE = capital > 0 ? (totalRealisedPnl / capital * 100) : 0;
      uniqueOpen = Object.keys(openSymbols).length;
      unrealisedROCE = capital > 0 ? (totalUnrealisedPnl / capital * 100) : 0;
    }
  }

  return {
    capital: capital,
    invested: invested,
    investedPct: investedPct,
    openPositions: openPositions,
    uniqueOpen: uniqueOpen,
    realisedROCE: realisedROCE,
    unrealisedROCE: unrealisedROCE,
    totalRealisedPnl: totalRealisedPnl,
    totalPartialRealisedPnl: totalPartialRealisedPnl,
    runningCapital: capital + totalRealisedPnl,
    losingStreak: streaks.losing,
    consecWins: streaks.wins,
    config: config
  };
}
function calcStreaks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trades = ss.getSheetByName('Trades');
  if (!trades) return {losing:0, wins:0};
  const last = trades.getLastRow();
  if (last < 2) return {losing:0, wins:0};

  // Build map of Trade_ID -> last exit datetime from Trade_Legs
  const legs = ss.getSheetByName('Trade_Legs');
  var lastExitMap = {};
  if (legs && legs.getLastRow() > 1) {
    var legData = legs.getRange(2, 1, legs.getLastRow()-1, 4).getValues();
    for (var i = 0; i < legData.length; i++) {
      var tid = legData[i][1];  // Trade_ID
      var ld = legData[i][2];   // Date
      var lt = legData[i][3];   // Time (HH:mm)
      if (!tid || !ld) continue;
      var dt = new Date(ld);
      if (lt) {
        if (lt instanceof Date) {
          dt.setHours(lt.getHours(), lt.getMinutes());
        } else {
          var parts = String(lt).split(':');
          dt.setHours(parseInt(parts[0])||0, parseInt(parts[1])||0);
        }
      }
      var ts = dt.getTime();
      if (!lastExitMap[tid] || ts > lastExitMap[tid]) lastExitMap[tid] = ts;
    }
  }

  // W=23 (Status), AB=28 (Result: Win/Loss), X=24 (Exit_Date), A=0 (Trade_ID)
  const data = trades.getRange(2,1,last-1,39).getValues();
  const closed = data
    .filter(function(r) { return r[22]==='Closed' && (r[27]==='Win'||r[27]==='Loss'); })
    .sort(function(a,b) {
      var aTs = lastExitMap[a[0]] || new Date(a[23]).getTime();
      var bTs = lastExitMap[b[0]] || new Date(b[23]).getTime();
      return bTs - aTs;
    });
  if (!closed.length) return {losing:0, wins:0};
  var losing=0, wins=0;
  var firstResult = closed[0][27];
  if (firstResult === 'Loss') {
    for (var i=0; i<closed.length; i++) {
      if (closed[i][27]==='Loss') losing++; else break;
    }
  } else {
    for (var i=0; i<closed.length; i++) {
      if (closed[i][27]==='Win') wins++; else break;
    }
  }
  const summary = ss.getSheetByName('Summary');
  if (summary) {
    summary.getRange('B20').setValue(losing);
    summary.getRange('B21').setValue(wins);
  }
  return {losing:losing, wins:wins};
}

// ==================== CHARGE CALCULATION (Zerodha) ====================
// contract: 'CNC' or 'MIS', side: 'Buy' or 'Sell'
// Returns {brokerage, stt, txn, sebi, gst, stamp, dp, totalCharges, netPrice}
// For Buy: netPrice = rawPrice + charges/qty (cost goes up)
// For Sell: netPrice = rawPrice - charges/qty (proceeds go down)
// Brokerage/STT rates come from Config; the remaining statutory charges fall back to
// ZERODHA_CHARGE_DEFAULTS when a key is missing from the Config sheet.
// KEEP IN SYNC with calcCharges() in trading-journal.html.
var ZERODHA_CHARGE_DEFAULTS = {
  Txn_Charge_Pct: 0.00307,  // NSE exchange transaction charge, % of value, both sides
  SEBI_Per_Crore: 10,       // SEBI turnover fee, ₹ per crore, both sides
  GST_Pct: 18,              // GST on brokerage + SEBI + exchange charges
  MIS_Stamp_Buy: 0.003,     // Stamp duty %, buy side only
  CNC_Stamp_Buy: 0.015,
  CNC_DP_Sell: 15.34        // Flat DP charge ₹ on delivery sells
};
function calcCharges(rawPrice, qty, contract, side) {
  var cfg = getConfig();
  var tradeValue = rawPrice * qty;

  // Pick rate & cap keys based on contract type + side
  var prefix = (contract === 'CNC') ? 'CNC' : 'MIS';
  var brokRate = (cfg[prefix + '_Brok_' + side] || 0) / 100;
  var brokCap  = cfg[prefix + '_Brok_' + side + '_Cap'] || 0;
  var sttRate  = (cfg[prefix + '_STT_' + side] || 0) / 100;
  var sttCap   = cfg[prefix + '_STT_' + side + '_Cap'] || 0;

  var other = function(key) {
    var v = cfg[key];
    return (v === undefined || v === '' || v === null || isNaN(Number(v))) ? ZERODHA_CHARGE_DEFAULTS[key] : Number(v);
  };

  var brokerage = tradeValue * brokRate;
  if (brokCap > 0) brokerage = Math.min(brokerage, brokCap);
  var stt = tradeValue * sttRate;
  if (sttCap > 0) stt = Math.min(stt, sttCap);

  var txn   = tradeValue * other('Txn_Charge_Pct') / 100;
  var sebi  = tradeValue * other('SEBI_Per_Crore') / 10000000;
  var gst   = (brokerage + txn + sebi) * other('GST_Pct') / 100;
  var stamp = (side === 'Buy') ? tradeValue * other(prefix + '_Stamp_Buy') / 100 : 0;
  var dp    = (side === 'Sell' && prefix === 'CNC') ? other('CNC_DP_Sell') : 0;

  var totalCharges = brokerage + stt + txn + sebi + gst + stamp + dp;
  var netPrice;
  if (side === 'Buy') {
    netPrice = (tradeValue + totalCharges) / qty;
  } else {
    netPrice = (tradeValue - totalCharges) / qty;
  }
  return {brokerage: brokerage, stt: stt, txn: txn, sebi: sebi, gst: gst, stamp: stamp, dp: dp,
          totalCharges: totalCharges, netPrice: netPrice};
}

// ==================== CAPITAL ====================
function getCapital() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName('Summary');
  return { capital: summary?(summary.getRange('B2').getValue()||0):0 };
}
function getInvested() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName('Summary');
  return {
    investedAmount: summary?(summary.getRange('B9').getValue()||0):0,
    investedPct: summary?(summary.getRange('B10').getValue()||0):0
  };
}
function updateCapital(amount, status, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txn = ss.getSheetByName('Capital_Transactions');
  if (!txn) return {error:'Sheet not found. Run setupSheet().'};
  amount = parseFloat(amount);
  if (isNaN(amount)||amount<=0) return {error:'Invalid amount'};
  if (status!=='Added'&&status!=='Withdrawn') return {error:'Invalid status'};
  if (status==='Withdrawn') {
    const cap = getCapital().capital;
    if (amount>cap) return {error:'Insufficient capital. Current: ₹'+cap};
  }
  const now = new Date();
  const txnId = 'CAP-'+Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyyMMdd-HHmmss');
  const timeStr = Utilities.formatDate(now,Session.getScriptTimeZone(),'HH:mm:ss');
  const txnDate = date||Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const nr = txn.getLastRow()+1;
  const bf = nr===2 ? '=IF(D2="Added",E2,-E2)' : '=F'+(nr-1)+'+IF(D'+nr+'="Added",E'+nr+',-E'+nr+')';
  txn.getRange(nr,1,1,7).setValues([[txnId,txnDate,timeStr,status,amount,bf,now]]);
  txn.getRange(nr,5).setNumberFormat('₹#,##0.00');
  txn.getRange(nr,6).setNumberFormat('₹#,##0.00');
  txn.getRange(nr,7).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  txn.getRange(nr,4).setFontColor(status==='Withdrawn'?'#cc0000':'#008800');
  return {success:true, txnId:txnId};
}

// ==================== TRANSACTIONS ====================
function getTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txn = ss.getSheetByName('Capital_Transactions');
  if (!txn) return {transactions:[]};
  const last = txn.getLastRow();
  if (last<2) return {transactions:[]};
  const data = txn.getRange(2,1,last-1,7).getValues();
  return {
    transactions: data.filter(function(r){return r[0]!=='';}).map(function(r){
      return {id:r[0],date:r[1],time:r[2],type:r[3],amount:r[4],balance:r[5],timestamp:r[6]};
    })
  };
}

// ==================== NEW TRADE ====================
function addNewTrade(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trades = ss.getSheetByName('Trades');
  if (!trades) return {error:'Trades sheet not found. Run setupSheet().'};
  if (!data.symbol||!data.symbol.trim()) return {error:'Symbol required'};
  if (!data.buyPrice||data.buyPrice<=0) return {error:'Invalid buy price'};
  if (!data.initialSL||data.initialSL<=0) return {error:'Invalid stop loss'};
  if (!data.lotSize||data.lotSize<=0) return {error:'Invalid lot size'};
  const now = new Date();
  const tradeId = 'T-'+Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyyMMdd-HHmmss');
  const nr = trades.getLastRow()+1;
  const totalQty = (data.lots||1) * data.lotSize;
  // Columns A-AG (33 cols)
  // V=Remaining_Qty, W=Status, X-AB=Exit fields, AC=LF_Rise, AD=Entry_Chart_URL
  // AE=Buy_Charges, AF=Net_Buy_Price, AG=Current_SL
  const entryChartUrl = data.entryChartUrl || '';
  const entryChartUrl2 = data.entryChartUrl2 || '';
  const entryChartUrl3 = data.entryChartUrl3 || '';
  const contract = data.contract || 'MIS';
  const buyCharges = calcCharges(data.buyPrice, totalQty, contract, 'Buy');
  // Snapshot current capital at trade creation time
  const summarySheet = ss.getSheetByName('Summary');
  const tradeCapital = summarySheet ? (summarySheet.getRange('B2').getValue() || 0) : 0;
  const row = [
    tradeId, data.type||'L', contract, data.symbol.toUpperCase(),
    data.lots||1, data.lotSize, data.buyDate, data.buyTime,
    data.buyPrice, data.initialSL, data.setup||'',
    data.rvol||0, data.adRatio||0, data.causeMove||0, data.consolMove||0, data.maUndercut||'',
    '=E'+nr+'*F'+nr+'*I'+nr,                                                           // Q: Invested
    '=IF(AK'+nr+'>0,Q'+nr+'/AK'+nr+'*100,0)',                                          // R: Pos Size % (uses trade's own capital)
    '=ABS(I'+nr+'-J'+nr+')/I'+nr+'*100',                                                // S: SL %
    '=IF(AK'+nr+'>0,(E'+nr+'*F'+nr+'*ABS(I'+nr+'-J'+nr+'))/AK'+nr+'*100,0)',           // T: Risk % (uses trade's own capital)
    now,                                                                                 // U: Timestamp
    totalQty,                                                                            // V: Remaining_Qty
    'Open',                                                                              // W: Status
    '','','','','',                                                                      // X-AB: Exit fields
    data.lfRise||'',                                                                     // AC: LF_Rise
    entryChartUrl,                                                                       // AD: Entry_Chart_URL
    buyCharges.totalCharges,                                                             // AE: Buy_Charges
    buyCharges.netPrice,                                                                 // AF: Net_Buy_Price
    data.initialSL,                                                                      // AG: Current_SL (starts as Initial_SL)
    data.entryNotes||'',                                                                   // AH: Entry_Notes
    data.marketEnv||'',                                                                    // AI: Market_Env
    '',                                                                                    // AJ: ROCE_% (empty for open trades)
    tradeCapital,                                                                          // AK: Capital (snapshotted at trade creation)
    entryChartUrl2,                                                                        // AL: Entry_Chart_URL_2
    entryChartUrl3                                                                         // AM: Entry_Chart_URL_3
  ];
  trades.getRange(nr,1,1,row.length).setValues([row]);
  // Formatting
  trades.getRange(nr,9).setNumberFormat('₹#,##0.00');
  trades.getRange(nr,10).setNumberFormat('₹#,##0.00');
  trades.getRange(nr,17).setNumberFormat('₹#,##0.00');
  trades.getRange(nr,18).setNumberFormat('0.00"%"');
  trades.getRange(nr,19).setNumberFormat('0.00"%"');
  trades.getRange(nr,20).setNumberFormat('0.00"%"');
  trades.getRange(nr,21).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  trades.getRange(nr,2).setFontColor(data.type==='L'?'#008800':'#cc0000');
  trades.getRange(nr,23).setFontColor('#008800');
  trades.getRange(nr,31).setNumberFormat('₹#,##0.00'); // AE: Buy_Charges
  trades.getRange(nr,32).setNumberFormat('₹#,##0.00'); // AF: Net_Buy_Price
  trades.getRange(nr,33).setNumberFormat('₹#,##0.00'); // AG: Current_SL
  trades.getRange(nr,37).setNumberFormat('₹#,##0.00'); // AK: Capital

  // Also record entry as a Buy leg in Trade_Legs
  const legs = ss.getSheetByName('Trade_Legs');
  if (legs) {
    const lr = legs.getLastRow()+1;
    const legId = tradeId + '-E1';
    const legTime = data.buyTime || Utilities.formatDate(now,Session.getScriptTimeZone(),'HH:mm');
    legs.getRange(lr,1,1,10).setValues([
      [legId, tradeId, data.buyDate, legTime, 'Buy', data.buyPrice, totalQty, totalQty*data.buyPrice, totalQty, now]
    ]);
    legs.getRange(lr,6).setNumberFormat('₹#,##0.00');
    legs.getRange(lr,8).setNumberFormat('₹#,##0.00');
    legs.getRange(lr,10).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  }

  return {success:true, tradeId:tradeId};
}

// ==================== EXIT LEG (PARTIAL / FULL) ====================
function addExitLeg(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trades = ss.getSheetByName('Trades');
  const legs = ss.getSheetByName('Trade_Legs');
  if (!trades || !legs) return {error:'Sheets not found. Run setupSheet().'};
  if (!data.tradeId) return {error:'Trade ID required'};
  if (!data.exitPrice || data.exitPrice <= 0) return {error:'Invalid exit price'};
  if (!data.exitQty || data.exitQty <= 0) return {error:'Invalid exit quantity'};

  // Find the trade row
  const last = trades.getLastRow();
  if (last < 2) return {error:'No trades found'};
  const allData = trades.getRange(2,1,last-1,39).getValues();
  var tradeRow = -1;
  for (var i=0; i<allData.length; i++) {
    if (allData[i][0] === data.tradeId) { tradeRow = i+2; break; }
  }
  if (tradeRow === -1) return {error:'Trade not found: '+data.tradeId};

  var row = allData[tradeRow-2];
  var remainingQty = row[21]; // V=22nd col, 0-indexed = 21
  var status = row[22];       // W
  var tradeType = row[1];     // B (L or S)
  var buyPrice = row[8];      // I
  var contract = row[2];      // C (CNC or MIS)
  var netBuyPrice = row[31];  // AF (Net_Buy_Price)

  if (status === 'Closed') return {error:'Trade already fully closed'};
  if (data.exitQty > remainingQty) return {error:'Exit qty ('+data.exitQty+') exceeds remaining ('+remainingQty+')'};

  // Update contract type if changed (e.g. CNC trade exited same day as MIS)
  if (data.contract && data.contract !== contract) {
    contract = data.contract;
    trades.getRange(tradeRow, 3).setValue(contract); // C = Contract column
  }

  var exitPrice = parseFloat(data.exitPrice);
  var exitQty = parseInt(data.exitQty);
  var exitDate = data.exitDate || Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  var exitTime = data.exitTime || Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'HH:mm');
  var now = new Date();

  // 1. Record the exit leg
  var legCount = 0;
  var legLast = legs.getLastRow();
  if (legLast >= 2) {
    var legData = legs.getRange(2,2,legLast-1,1).getValues();
    for (var j=0; j<legData.length; j++) {
      if (legData[j][0] === data.tradeId) legCount++;
    }
  }
  var legId = data.tradeId + '-X' + (legCount); // X for exit legs (E was entry)
  var newRemaining = remainingQty - exitQty;
  var legAmount = exitQty * exitPrice;
  var exitChartUrl = data.exitChartUrl || '';
  var sellCharges = calcCharges(exitPrice, exitQty, contract, 'Sell');
  var lr = legs.getLastRow()+1;
  var exitNotes = data.exitNotes || '';
  legs.getRange(lr,1,1,14).setValues([
    [legId, data.tradeId, exitDate, exitTime, 'Sell', exitPrice, exitQty, legAmount, newRemaining, now, exitChartUrl, sellCharges.totalCharges, sellCharges.netPrice, exitNotes]
  ]);
  legs.getRange(lr,6).setNumberFormat('₹#,##0.00');
  legs.getRange(lr,8).setNumberFormat('₹#,##0.00');
  legs.getRange(lr,10).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  legs.getRange(lr,12).setNumberFormat('₹#,##0.00');
  legs.getRange(lr,13).setNumberFormat('₹#,##0.00');

  // 2. Update Remaining_Qty in Trades (col V = 22)
  trades.getRange(tradeRow, 22).setValue(newRemaining);

  // 3. If fully exited, compute final P&L and close
  if (newRemaining === 0) {
    // Get all exit legs for this trade to compute weighted avg exit price
    var allLegs = legs.getRange(2,1,legs.getLastRow()-1,14).getValues();
    var totalNetExitAmount = 0;
    var totalExitQty = 0;
    var totalRawExitAmount = 0;
    var lastExitDate = exitDate;
    for (var k=0; k<allLegs.length; k++) {
      if (allLegs[k][1] === data.tradeId && allLegs[k][4] === 'Sell') {
        totalRawExitAmount += allLegs[k][7];      // Raw Amount
        totalExitQty += allLegs[k][6];             // Quantity
        totalNetExitAmount += allLegs[k][12] * allLegs[k][6]; // Net_Exit_Price × Qty
        var ld = allLegs[k][2];
        if (ld && new Date(ld) > new Date(lastExitDate)) lastExitDate = ld;
      }
    }
    var avgExitPrice = totalExitQty > 0 ? totalRawExitAmount / totalExitQty : 0;
    var totalQty = row[4] * row[5]; // lots * lotSize
    // P&L uses net prices (after charges)
    var totalNetBuyAmount = netBuyPrice * totalQty;
    var pnl = tradeType === 'L'
      ? totalNetExitAmount - totalNetBuyAmount
      : totalNetBuyAmount - totalNetExitAmount;
    var totalBuyAmount = totalQty * buyPrice;
    var pnlPct = totalBuyAmount > 0 ? (pnl / totalBuyAmount) : 0;
    var result = pnl >= 0 ? 'Win' : 'Loss';

    // ROCE_% = P&L as % of capital at trade creation time (snapshotted in AK)
    var tradeCapital = row[36] || 0;
    if (tradeCapital <= 0) {
      // Fallback for older trades without snapshotted capital
      var summarySheet = ss.getSheetByName('Summary');
      tradeCapital = summarySheet ? (summarySheet.getRange('B2').getValue() || 0) : 0;
    }
    var rocePct = tradeCapital > 0 ? (pnl / tradeCapital) : 0;

    // W=Status, X=Exit_Date, Y=Avg_Exit_Price, Z=P&L, AA=P&L_%, AB=Result, AJ=ROCE_%
    trades.getRange(tradeRow, 23).setValue('Closed');
    trades.getRange(tradeRow, 24).setValue(lastExitDate);
    trades.getRange(tradeRow, 25).setValue(avgExitPrice);
    trades.getRange(tradeRow, 26).setValue(pnl);
    trades.getRange(tradeRow, 27).setValue(pnlPct);
    trades.getRange(tradeRow, 28).setValue(result);
    trades.getRange(tradeRow, 36).setValue(rocePct);

    // Formatting
    trades.getRange(tradeRow, 23).setFontColor(result==='Win'?'#008800':'#cc0000');
    trades.getRange(tradeRow, 25).setNumberFormat('₹#,##0.00');
    trades.getRange(tradeRow, 26).setNumberFormat('₹#,##0.00');
    trades.getRange(tradeRow, 27).setNumberFormat('0.00%');
    trades.getRange(tradeRow, 36).setNumberFormat('0.00%');

    return {
      success: true,
      legId: legId,
      status: 'Closed',
      remainingQty: 0,
      avgExitPrice: avgExitPrice,
      pnl: pnl,
      pnlPct: pnlPct,
      rocePct: rocePct,
      result: result
    };
  }

  return {
    success: true,
    legId: legId,
    status: 'Partial',
    remainingQty: newRemaining,
    exitedQty: exitQty,
    exitPrice: exitPrice
  };
}

// ==================== GET EXIT LEGS FOR A TRADE ====================
function getExitLegs(tradeId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const legs = ss.getSheetByName('Trade_Legs');
  if (!legs) return {legs:[]};
  const last = legs.getLastRow();
  if (last < 2) return {legs:[]};
  const data = legs.getRange(2,1,last-1,14).getValues();
  const result = data
    .filter(function(r) { return r[1] === tradeId && r[4] === 'Sell'; })
    .map(function(r) {
      return { legId:r[0], date:r[2], time:r[3], price:r[5], qty:r[6], amount:r[7], remaining:r[8], chartUrl:r[10]||'', exitCharges:r[11]||0, netExitPrice:r[12]||0, exitNotes:r[13]||'' };
    });
  return {legs: result};
}

// ==================== OPEN TRADES ====================
function getOpenTrades() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trades = ss.getSheetByName('Trades');
  if (!trades) return {trades:[]};
  const last = trades.getLastRow();
  if (last<2) return {trades:[]};
  // 32 columns (A-AF)
  const data = trades.getRange(2,1,last-1,39).getValues();
  const open = data.filter(function(r){ return r[22]==='Open'; }).map(function(r){
    var totalQty = r[4]*r[5];
    var remainingQty = r[21];
    var exitedQty = totalQty - remainingQty;
    return {
      id:r[0], type:r[1], contract:r[2], symbol:r[3], lots:r[4], lotSize:r[5],
      totalQty: totalQty,
      remainingQty: remainingQty,
      exitedQty: exitedQty,
      buyDate:fmtDate_(r[6]), buyTime:fmtTime_(r[7]), buyPrice:r[8], initialSL:r[9], setup:r[10],
      rvol:r[11]||0, adRatio:r[12]||0, causeMove:r[13]||0, consolMove:r[14]||0, maUndercut:r[15]||'',
      invested:r[16], posPct:r[17], slPct: typeof r[18]==='number'?r[18].toFixed(2):r[18],
      riskPct: typeof r[19]==='number'?r[19].toFixed(2):r[19],
      lfRise:r[28]||'', entryChartUrl:r[29]||'',
      buyCharges:r[30]||0, netBuyPrice:r[31]||0,
      currentSL:r[32]||r[9], initialSL:r[9], entryNotes:r[33]||'',
      marketEnv:r[34]||'', rocePct:r[35]||0,
      entryChartUrl2:r[37]||'', entryChartUrl3:r[38]||''
    };
  });
  return {trades:open};
}

// ==================== ALL TRADES (Open + Closed) ====================
function getAllTrades() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trades = ss.getSheetByName('Trades');
  if (!trades) return {trades:[]};
  var last = trades.getLastRow();
  if (last<2) return {trades:[]};
  var data = trades.getRange(2,1,last-1,39).getValues();
  var result = data.filter(function(r){ return r[0]!==''; }).map(function(r){
    var totalQty = r[4]*r[5];
    var remainingQty = r[21]||0;
    var exitedQty = totalQty - remainingQty;
    var status = r[22]||'Open';
    return {
      id:r[0], type:r[1], contract:r[2], symbol:r[3], lots:r[4], lotSize:r[5],
      totalQty: totalQty, remainingQty: remainingQty, exitedQty: exitedQty,
      buyDate:fmtDate_(r[6]), buyTime:fmtTime_(r[7]), buyPrice:r[8], initialSL:r[9], setup:r[10],
      rvol:r[11]||0, adRatio:r[12]||0, causeMove:r[13]||0, consolMove:r[14]||0, maUndercut:r[15]||'',
      invested:r[16], posPct:r[17],
      slPct: typeof r[18]==='number'?r[18].toFixed(2):r[18],
      riskPct: typeof r[19]==='number'?r[19].toFixed(2):r[19],
      status: status,
      exitDate:fmtDate_(r[23]), avgExitPrice:r[24]||0, pnl:r[25]||0,
      pnlPct: r[26]||0,
      result:r[27]||'',
      lfRise:r[28]||'', entryChartUrl:r[29]||'',
      buyCharges:r[30]||0, netBuyPrice:r[31]||0,
      currentSL:r[32]||r[9], initialSL:r[9], entryNotes:r[33]||'',
      marketEnv:r[34]||'', rocePct:r[35]||0,
      capital:r[36]||0,
      entryChartUrl2:r[37]||'', entryChartUrl3:r[38]||''
    };
  });
  result.sort(function(a,b){ var da=String(b.buyDate||''), db=String(a.buyDate||''); return da<db?-1:da>db?1:0; });
  return {trades:result};
}

// ==================== UPDATE TRADE SL (Trailing) ====================
function updateTradeSL(tradeId, newSL) {
  if (!tradeId) return {error:'Trade ID required'};
  newSL = parseFloat(newSL);
  if (isNaN(newSL) || newSL <= 0) return {error:'Invalid SL price'};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trades = ss.getSheetByName('Trades');
  if (!trades) return {error:'Trades sheet not found'};
  const last = trades.getLastRow();
  if (last < 2) return {error:'No trades found'};
  const data = trades.getRange(2,1,last-1,39).getValues();
  for (var i=0; i<data.length; i++) {
    if (data[i][0] === tradeId) {
      if (data[i][22] !== 'Open') return {error:'Trade is not open'};
      var row = i+2;
      trades.getRange(row, 33).setValue(newSL); // AG: Current_SL
      trades.getRange(row, 33).setNumberFormat('₹#,##0.00');
      return {success:true, tradeId:tradeId, newSL:newSL};
    }
  }
  return {error:'Trade not found: '+tradeId};
}

// ==================== UPDATE TRADE FIELDS (Market Context, Deep Dive, Chart, Notes) ====================
function updateTradeFields(data) {
  if (!data.tradeId) return {error:'Trade ID required'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trades = ss.getSheetByName('Trades');
  if (!trades) return {error:'Trades sheet not found'};
  var last = trades.getLastRow();
  if (last < 2) return {error:'No trades found'};
  var rows = trades.getRange(2,1,last-1,1).getValues();
  for (var i=0; i<rows.length; i++) {
    if (rows[i][0] === data.tradeId) {
      var row = i + 2;
      // Column map: E=5:Lots, F=6:Lot_Size, I=9:Buy_Price
      // L=12:RVol, M=13:PreMkt_AD, N=14:Cause, O=15:Consol, P=16:MA_Undercut
      // Q=17:Invested(formula), R=18:Pos_Size_%, S=19:SL_%, T=20:Risk_%
      // V=22:Remaining_Qty, AE=31:Buy_Charges, AF=32:Net_Buy_Price
      // AC=29:LF_Rise, AD=30:Entry_Chart_URL, AH=34:Entry_Notes

      // ---- Trade Parameters (Entry Price, Lots, Lot Size) ----
      var priceChanged = false;
      var oldRow = trades.getRange(row,1,1,39).getValues()[0];
      var oldBuyPrice = oldRow[8];   // I (0-indexed col 8)
      var oldLots = oldRow[4];       // E
      var oldLotSize = oldRow[5];    // F
      var contract = oldRow[2];      // C

      if (data.entryPrice !== undefined && data.entryPrice > 0) {
        trades.getRange(row,9).setValue(data.entryPrice);    // I: Buy_Price
        trades.getRange(row,9).setNumberFormat('₹#,##0.00');
        priceChanged = priceChanged || (data.entryPrice !== oldBuyPrice);
      }
      if (data.lots !== undefined && data.lots > 0) {
        trades.getRange(row,5).setValue(data.lots);          // E: Lots
        priceChanged = priceChanged || (data.lots !== oldLots);
      }
      if (data.lotSize !== undefined && data.lotSize > 0) {
        trades.getRange(row,6).setValue(data.lotSize);       // F: Lot_Size
        priceChanged = priceChanged || (data.lotSize !== oldLotSize);
      }

      // Recalculate dependent fields when price/qty changed
      if (priceChanged) {
        var newBuyPrice = (data.entryPrice !== undefined && data.entryPrice > 0) ? data.entryPrice : oldBuyPrice;
        var newLots = (data.lots !== undefined && data.lots > 0) ? data.lots : oldLots;
        var newLotSize = (data.lotSize !== undefined && data.lotSize > 0) ? data.lotSize : oldLotSize;
        var newTotalQty = newLots * newLotSize;

        // Q: Invested formula (=E*F*I)
        trades.getRange(row,17).setFormula('=E'+row+'*F'+row+'*I'+row);
        trades.getRange(row,17).setNumberFormat('₹#,##0.00');

        // R: Position Size % (=IF(AK>0,Q/AK*100,0))
        trades.getRange(row,18).setFormula('=IF(AK'+row+'>0,Q'+row+'/AK'+row+'*100,0)');
        trades.getRange(row,18).setNumberFormat('0.00"%"');

        // S: SL % (=ABS(I-J)/I*100)
        trades.getRange(row,19).setFormula('=ABS(I'+row+'-J'+row+')/I'+row+'*100');
        trades.getRange(row,19).setNumberFormat('0.00"%"');

        // T: Risk % (=IF(AK>0,(E*F*ABS(I-J))/AK*100,0))
        trades.getRange(row,20).setFormula('=IF(AK'+row+'>0,(E'+row+'*F'+row+'*ABS(I'+row+'-J'+row+'))/AK'+row+'*100,0)');
        trades.getRange(row,20).setNumberFormat('0.00"%"');

        // AE: Buy_Charges & AF: Net_Buy_Price (recalculate)
        var buyCharges = calcCharges(newBuyPrice, newTotalQty, contract, 'Buy');
        trades.getRange(row,31).setValue(buyCharges.totalCharges);   // AE
        trades.getRange(row,31).setNumberFormat('₹#,##0.00');
        trades.getRange(row,32).setValue(buyCharges.netPrice);       // AF
        trades.getRange(row,32).setNumberFormat('₹#,##0.00');

        // V: Update Remaining_Qty only if total qty changed (adjust by difference)
        var oldTotalQty = oldLots * oldLotSize;
        var oldRemaining = oldRow[21]; // V (0-indexed 21)
        var oldExitedQty = oldTotalQty - oldRemaining;
        var newRemaining = newTotalQty - oldExitedQty;
        if (newRemaining < 0) newRemaining = 0;
        trades.getRange(row,22).setValue(newRemaining);

        // Update the Buy leg in Trade_Legs
        var legsSheet = ss.getSheetByName('Trade_Legs');
        if (legsSheet) {
          var legsLast = legsSheet.getLastRow();
          if (legsLast >= 2) {
            var legsData = legsSheet.getRange(2,1,legsLast-1,9).getValues();
            for (var li=0; li<legsData.length; li++) {
              if (legsData[li][1] === data.tradeId && legsData[li][4] === 'Buy') {
                var legRow = li + 2;
                legsSheet.getRange(legRow,6).setValue(newBuyPrice);           // Price
                legsSheet.getRange(legRow,7).setValue(newTotalQty);           // Quantity
                legsSheet.getRange(legRow,8).setValue(newTotalQty * newBuyPrice); // Amount
                legsSheet.getRange(legRow,9).setValue(newRemaining);          // Running_Remaining
                legsSheet.getRange(legRow,6).setNumberFormat('₹#,##0.00');
                legsSheet.getRange(legRow,8).setNumberFormat('₹#,##0.00');
                break;
              }
            }
          }
        }

        // ---- Keep Status (col W) in sync with the Remaining_Qty this edit just produced ----
        var oldStatus = oldRow[22]; // W (0-indexed 22)
        var wasAlreadyClosed = oldStatus === 'Closed';
        var nowClosed = newRemaining === 0;

        if (nowClosed) {
          // Fully exited — either already was Closed (recompute P&L for the new price/qty),
          // or this edit just closed it (lots/lotSize reduced to match already-exited qty).
          var netBuyPrice = buyCharges.netPrice;
          var tradeType = oldRow[1]; // B: Type (L or S)
          var legsSheet2 = ss.getSheetByName('Trade_Legs');
          if (legsSheet2) {
            var ll = legsSheet2.getLastRow();
            if (ll >= 2) {
              var allLegs = legsSheet2.getRange(2,1,ll-1,14).getValues();
              var totalNetExitAmount = 0;
              var totalExitQty = 0;
              var totalRawExitAmount = 0;
              for (var kk=0; kk<allLegs.length; kk++) {
                if (allLegs[kk][1] === data.tradeId && allLegs[kk][4] === 'Sell') {
                  totalRawExitAmount += allLegs[kk][7];
                  totalExitQty += allLegs[kk][6];
                  totalNetExitAmount += allLegs[kk][12] * allLegs[kk][6]; // Net_Exit_Price × Qty
                }
              }
              var avgExitPrice = totalExitQty > 0 ? totalRawExitAmount / totalExitQty : 0;
              var totalNetBuyAmount = netBuyPrice * newTotalQty;
              var pnl = tradeType === 'L'
                ? totalNetExitAmount - totalNetBuyAmount
                : totalNetBuyAmount - totalNetExitAmount;
              var totalBuyAmount = newTotalQty * newBuyPrice;
              var pnlPct = totalBuyAmount > 0 ? (pnl / totalBuyAmount) : 0;
              var result = pnl >= 0 ? 'Win' : 'Loss';
              var tradeCapital = oldRow[36] || 0;
              var rocePct = tradeCapital > 0 ? (pnl / tradeCapital) : 0;

              trades.getRange(row,23).setValue('Closed');      // W: Status
              if (!wasAlreadyClosed) {
                // Trade was Open and never got an Exit_Date — derive it from the latest sell leg now.
                var lastExitDate = '';
                for (var md=0; md<allLegs.length; md++) {
                  if (allLegs[md][1] === data.tradeId && allLegs[md][4] === 'Sell') {
                    var legDate = allLegs[md][2];
                    if (legDate && (!lastExitDate || new Date(legDate) > new Date(lastExitDate))) lastExitDate = legDate;
                  }
                }
                trades.getRange(row,24).setValue(lastExitDate); // X: Exit_Date
              }
              trades.getRange(row,25).setValue(avgExitPrice);  // Y
              trades.getRange(row,26).setValue(pnl);           // Z
              trades.getRange(row,27).setValue(pnlPct);        // AA
              trades.getRange(row,28).setValue(result);         // AB
              trades.getRange(row,36).setValue(rocePct);        // AJ
              trades.getRange(row,25).setNumberFormat('₹#,##0.00');
              trades.getRange(row,26).setNumberFormat('₹#,##0.00');
              trades.getRange(row,28).setFontColor(result==='Win'?'#008800':'#cc0000');
              trades.getRange(row,23).setFontColor(result==='Win'?'#008800':'#cc0000');
            }
          }
        } else if (wasAlreadyClosed) {
          // Was Closed but this edit reopened real exposure (Remaining_Qty > 0 again) —
          // reset Status to Open and clear the now-stale closing summary so a partially-open
          // trade doesn't keep showing the final P&L/Result from when it was fully exited.
          trades.getRange(row,23).setValue('Open');  // W: Status
          trades.getRange(row,24).setValue('');      // X: Exit_Date
          trades.getRange(row,25).setValue('');      // Y: Avg_Exit_Price
          trades.getRange(row,26).setValue('');      // Z: P&L
          trades.getRange(row,27).setValue('');      // AA: P&L_%
          trades.getRange(row,28).setValue('');      // AB: Result
          trades.getRange(row,36).setValue('');      // AJ: ROCE_%
          trades.getRange(row,23).setFontColor('#008800');
        }
      }

      // ---- Market Context & Notes (existing fields) ----
      if (data.rvol !== undefined)       trades.getRange(row,12).setValue(parseFloat(data.rvol)||0);
      if (data.adRatio !== undefined)    trades.getRange(row,13).setValue(parseFloat(data.adRatio)||0);
      if (data.causeMove !== undefined)  trades.getRange(row,14).setValue(parseFloat(data.causeMove)||0);
      if (data.consolMove !== undefined) trades.getRange(row,15).setValue(parseFloat(data.consolMove)||0);
      if (data.maUndercut !== undefined) trades.getRange(row,16).setValue(data.maUndercut);
      if (data.lfRise !== undefined)     trades.getRange(row,29).setValue(parseFloat(data.lfRise)||0);
      if (data.entryChartUrl !== undefined) trades.getRange(row,30).setValue(data.entryChartUrl);
      if (data.entryChartUrl2 !== undefined) trades.getRange(row,38).setValue(data.entryChartUrl2);  // AL
      if (data.entryChartUrl3 !== undefined) trades.getRange(row,39).setValue(data.entryChartUrl3);  // AM
      if (data.entryNotes !== undefined) trades.getRange(row,34).setValue(data.entryNotes);
      if (data.marketEnv !== undefined) trades.getRange(row,35).setValue(data.marketEnv);

      // ---- Exit Legs: update price, qty, notes, chart in Trade_Legs ----
      if (data.exitLegs && data.exitLegs.length) {
        var legsSheet = ss.getSheetByName('Trade_Legs');
        if (legsSheet) {
          var ll = legsSheet.getLastRow();
          if (ll >= 2) {
            var allLegs = legsSheet.getRange(2,1,ll-1,14).getValues();
            // Collect sell-leg row indices for this trade (in order)
            var sellLegRows = [];
            for (var si=0; si<allLegs.length; si++) {
              if (allLegs[si][1] === data.tradeId && allLegs[si][4] === 'Sell') {
                sellLegRows.push(si + 2); // sheet row
              }
            }
            var exitLegsChanged = false;
            for (var ei=0; ei<data.exitLegs.length; ei++) {
              var legUp = data.exitLegs[ei];
              var legIdx = legUp.legIndex;
              if (legIdx >= 0 && legIdx < sellLegRows.length) {
                var legRow = sellLegRows[legIdx];
                // Update exit price (col F=6) and qty (col G=7) if provided
                if (legUp.exitPrice && legUp.exitPrice > 0) {
                  var oldLegPrice = allLegs[legRow-2][5];
                  if (legUp.exitPrice !== oldLegPrice) {
                    legsSheet.getRange(legRow,6).setValue(legUp.exitPrice);
                    legsSheet.getRange(legRow,6).setNumberFormat('₹#,##0.00');
                    legsSheet.getRange(legRow,8).setValue(legUp.exitPrice * (legUp.exitQty || allLegs[legRow-2][6])); // Amount
                    legsSheet.getRange(legRow,8).setNumberFormat('₹#,##0.00');
                    // Recalculate exit charges & net exit price
                    var legContract = trades.getRange(row,3).getValue() || 'MIS';
                    var legQty = legUp.exitQty && legUp.exitQty > 0 ? legUp.exitQty : allLegs[legRow-2][6];
                    var exitCharges = calcCharges(legUp.exitPrice, legQty, legContract, 'Sell');
                    legsSheet.getRange(legRow,12).setValue(exitCharges.totalCharges);
                    legsSheet.getRange(legRow,12).setNumberFormat('₹#,##0.00');
                    legsSheet.getRange(legRow,13).setValue(exitCharges.netPrice);
                    legsSheet.getRange(legRow,13).setNumberFormat('₹#,##0.00');
                    exitLegsChanged = true;
                  }
                }
                if (legUp.exitQty && legUp.exitQty > 0) {
                  var oldLegQty = allLegs[legRow-2][6];
                  if (legUp.exitQty !== oldLegQty) {
                    legsSheet.getRange(legRow,7).setValue(legUp.exitQty);
                    var lprice = legUp.exitPrice && legUp.exitPrice > 0 ? legUp.exitPrice : allLegs[legRow-2][5];
                    legsSheet.getRange(legRow,8).setValue(lprice * legUp.exitQty);
                    legsSheet.getRange(legRow,8).setNumberFormat('₹#,##0.00');
                    // Recalculate exit charges & net exit price for new qty
                    var legContract2 = trades.getRange(row,3).getValue() || 'MIS';
                    var exitCharges2 = calcCharges(lprice, legUp.exitQty, legContract2, 'Sell');
                    legsSheet.getRange(legRow,12).setValue(exitCharges2.totalCharges);
                    legsSheet.getRange(legRow,12).setNumberFormat('₹#,##0.00');
                    legsSheet.getRange(legRow,13).setValue(exitCharges2.netPrice);
                    legsSheet.getRange(legRow,13).setNumberFormat('₹#,##0.00');
                    exitLegsChanged = true;
                  }
                }
                // Update exit notes (col N=14) and chart URL (col K=11)
                if (legUp.exitNotes !== undefined) legsSheet.getRange(legRow,14).setValue(legUp.exitNotes);
                if (legUp.chartUrl !== undefined) legsSheet.getRange(legRow,11).setValue(legUp.chartUrl);
              }
            }

            // Recalculate Running_Remaining for all legs of this trade
            if (exitLegsChanged) {
              var updatedLegs = legsSheet.getRange(2,1,legsSheet.getLastRow()-1,14).getValues();
              var totalQtyNow = trades.getRange(row,5).getValue() * trades.getRange(row,6).getValue();
              var runningRem = totalQtyNow;
              for (var ri=0; ri<updatedLegs.length; ri++) {
                if (updatedLegs[ri][1] === data.tradeId) {
                  if (updatedLegs[ri][4] === 'Sell') {
                    runningRem -= updatedLegs[ri][6];
                  }
                  legsSheet.getRange(ri+2,9).setValue(runningRem);
                }
              }
              // Update Remaining_Qty on the trade
              trades.getRange(row,22).setValue(runningRem);

              // ---- Keep Status in sync with the Remaining_Qty this edit just produced ----
              var tradeStatus = trades.getRange(row,23).getValue();
              var wasClosedBeforeLegEdit = tradeStatus === 'Closed';
              var nowClosedAfterLegEdit = runningRem === 0;

              if (nowClosedAfterLegEdit) {
                var netBuyPriceNow = trades.getRange(row,32).getValue(); // AF
                var tradeTypeNow = trades.getRange(row,2).getValue();    // B
                var freshLegs = legsSheet.getRange(2,1,legsSheet.getLastRow()-1,14).getValues();
                var totNetExitAmt = 0, totExitQ = 0, totRawExitAmt = 0;
                for (var fl=0; fl<freshLegs.length; fl++) {
                  if (freshLegs[fl][1] === data.tradeId && freshLegs[fl][4] === 'Sell') {
                    totRawExitAmt += freshLegs[fl][7];
                    totExitQ += freshLegs[fl][6];
                    totNetExitAmt += freshLegs[fl][12] * freshLegs[fl][6];
                  }
                }
                var avgExit = totExitQ > 0 ? totRawExitAmt / totExitQ : 0;
                var buyPriceNow = trades.getRange(row,9).getValue();
                var totalQtyTrade = trades.getRange(row,5).getValue() * trades.getRange(row,6).getValue();
                var totNetBuyAmt = netBuyPriceNow * totalQtyTrade;
                var newPnl = tradeTypeNow === 'L' ? totNetExitAmt - totNetBuyAmt : totNetBuyAmt - totNetExitAmt;
                var totBuyAmt = totalQtyTrade * buyPriceNow;
                var newPnlPct = totBuyAmt > 0 ? (newPnl / totBuyAmt) : 0;
                var newResult = newPnl >= 0 ? 'Win' : 'Loss';
                var capNow = trades.getRange(row,37).getValue() || 0;
                var newRoce = capNow > 0 ? (newPnl / capNow) : 0;

                trades.getRange(row,23).setValue('Closed'); // W: Status
                if (!wasClosedBeforeLegEdit) {
                  // Trade was Open and never got an Exit_Date — derive it from the latest sell leg now.
                  var lastExitDateLeg = '';
                  for (var md2=0; md2<freshLegs.length; md2++) {
                    if (freshLegs[md2][1] === data.tradeId && freshLegs[md2][4] === 'Sell') {
                      var legDate2 = freshLegs[md2][2];
                      if (legDate2 && (!lastExitDateLeg || new Date(legDate2) > new Date(lastExitDateLeg))) lastExitDateLeg = legDate2;
                    }
                  }
                  trades.getRange(row,24).setValue(lastExitDateLeg); // X: Exit_Date
                }
                trades.getRange(row,25).setValue(avgExit);
                trades.getRange(row,26).setValue(newPnl);
                trades.getRange(row,27).setValue(newPnlPct);
                trades.getRange(row,28).setValue(newResult);
                trades.getRange(row,36).setValue(newRoce);
                trades.getRange(row,25).setNumberFormat('₹#,##0.00');
                trades.getRange(row,26).setNumberFormat('₹#,##0.00');
                trades.getRange(row,28).setFontColor(newResult==='Win'?'#008800':'#cc0000');
                trades.getRange(row,23).setFontColor(newResult==='Win'?'#008800':'#cc0000');
              } else if (wasClosedBeforeLegEdit) {
                // Was Closed but this edit reopened real exposure (Remaining_Qty > 0 again) —
                // reset Status to Open and clear the now-stale closing summary.
                trades.getRange(row,23).setValue('Open'); // W: Status
                trades.getRange(row,24).setValue('');     // X: Exit_Date
                trades.getRange(row,25).setValue('');     // Y: Avg_Exit_Price
                trades.getRange(row,26).setValue('');     // Z: P&L
                trades.getRange(row,27).setValue('');     // AA: P&L_%
                trades.getRange(row,28).setValue('');     // AB: Result
                trades.getRange(row,36).setValue('');     // AJ: ROCE_%
                trades.getRange(row,23).setFontColor('#008800');
              }
            }
          }
        }
      }

      return {success:true, tradeId:data.tradeId};
    }
  }
  return {error:'Trade not found: '+data.tradeId};
}

// ==================== CHART IMAGE UPLOAD TO GOOGLE DRIVE ====================
function getOrCreateChartFolder() {
  var folders = DriveApp.getFoldersByName('TradeCharts');
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder('TradeCharts');
}

function uploadChart(imageData, fileName, mimeType) {
  if (!imageData) return {error:'No image data provided'};
  try {
    var folder = getOrCreateChartFolder();
    var decoded = Utilities.base64Decode(imageData);
    var blob = Utilities.newBlob(decoded, mimeType || 'image/png', fileName || 'chart.png');
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();
    var viewUrl = 'https://drive.google.com/uc?id=' + fileId;
    return {success:true, url:viewUrl, fileId:fileId, driveUrl:file.getUrl()};
  } catch(e) {
    return {error:'Upload failed: '+e.message};
  }
}
