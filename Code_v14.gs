// ============================================================
// GREEN SALON — BILLING SYSTEM v14
// Owner: Harsha | Built: v14 complete rewrite for stability
// ============================================================
// SETUP STEPS (do once):
// 1. Extensions → Apps Script → clear all → paste this file
// 2. Fill MASTER_SHEET_ID and BRANCH_SHEETS below
// 3. Run firstTimeSetup() → approve all permissions
// 4. Run setupTriggers()
// 5. Deploy → New Deployment → Web App → Execute as: Me → Anyone
// 6. Copy /exec URL → paste into both HTML files as API_URL
// ============================================================

const MASTER_SHEET_ID = "1YUfe5XL6yFirq3CNfjHFHcqORMMcYAW9gI7pPnJyXuY";
const OWNER_PASSWORD  = "harsha@greensalon2026";
const BRANCH_SHEETS   = {
  "branch1": "1fRTEOMjhjqZ0P3pVfagC3hNdFeE3rL1LD_T98YraWGQ",
  "branch2": "14sPzYtF13bYldvYIC0xtGy0CbjKyWdpEVzCLQrKV5tA",
  "branch3": "1nQ2svtVxaKhCGGDltKS0u2iGooGlSkVboNj1OnZk9rc",
};

const C_DARK  = "#1a5c38";
const C_MED   = "#2d8653";
const C_WHITE = "#ffffff";
const C_ALT   = "#e8f5ee";
const C_RAW   = "#0f172a";

// ── ROUTER ────────────────────────────────────────────────────
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case "ownerLogin":       return R(ownerLogin(d));
      case "getBranches":      return R(getBranches());
      case "addBranch":        return R(addBranch(d));
      case "removeBranch":     return R(removeBranch(d));
      case "recoverBranch":    return R(recoverBranch(d));
      case "renameBranch":     return R(renameBranch(d));
      case "getStaffAdmin":    return R(getStaffAdmin(d));
      case "addStaff":         return R(addStaff(d));
      case "removeStaff":      return R(removeStaff(d));
      case "renameStaff":      return R(renameStaff(d));
      case "updateStaffComm":  return R(updateStaffComm(d));
      case "updateServices":   return R(updateServices(d));
      case "updateProducts":   return R(updateProducts(d));
      case "submitEntry":      return R(submitEntry(d));
      case "submitProduct":    return R(submitProduct(d));
      case "submitExpense":    return R(submitExpense(d));
      case "getMyEntries":     return R(getMyEntries(d));
      case "getTodayAll":      return R(getTodayAll(d));
      case "getBranchSummary": return R(getBranchSummary(d));
      case "deleteEntry":      return R(deleteEntry(d));
      case "deleteProduct":    return R(deleteProduct(d));
      case "deleteExpense":    return R(deleteExpense(d));
      case "getMonthSummary":  return R(getMonthSummary(d));
      case "setReportEmails":  return R(setReportEmails(d));
      case "getReportEmails":  return R(getReportEmails(d));
      case "sendManualReport": return R(sendManualReport(d));
      case "getLastUpdate":    return R(getLastUpdate(d));
      case "fixBranch":        return R(fixBranch(d));
      case "masterFix":        return R(masterFix());
      default: return E("Unknown action: " + d.action);
    }
  } catch (ex) { return E(ex.message + " | " + ex.stack); }
}

function doGet(e) {
  try {
    const a = e.parameter.action, bid = e.parameter.branchId;
    switch (a) {
      case "getStaff":      return R(getStaff(bid));
      case "getServices":   return R(getServices(bid));
      case "getProducts":   return R(getProducts(bid));
      case "getBranches":   return R(getBranches());
      case "getLastUpdate": return R(getLastUpdate({branchId:bid}));
      default: return E("Unknown action: " + a);
    }
  } catch (ex) { return E(ex.message); }
}

function R(d) { return ContentService.createTextOutput(JSON.stringify({success:true,...d})).setMimeType(ContentService.MimeType.JSON); }
function E(m) { return ContentService.createTextOutput(JSON.stringify({success:false,error:m})).setMimeType(ContentService.MimeType.JSON); }

// ── HELPERS ───────────────────────────────────────────────────
function masterSS()  { return SpreadsheetApp.openById(MASTER_SHEET_ID); }
function masterTab(name) { const ss=masterSS(); return ss.getSheetByName(name)||ss.insertSheet(name); }

function branchSS(branchId) {
  // Try hardcoded map first
  const sid = BRANCH_SHEETS[branchId];
  if (sid && !sid.includes("_SHEET_ID_HERE")) return SpreadsheetApp.openById(sid);
  // Fall back to Branches tab for dynamically added branches
  const rows = masterTab("Branches").getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === branchId && rows[i][4] !== false && rows[i][4] !== "FALSE")
      return SpreadsheetApp.openById(rows[i][3]);
  }
  throw new Error("Branch not found: " + branchId);
}

function branchTab(branchId, tabName) {
  const ss = branchSS(branchId);
  return ss.getSheetByName(tabName) || ss.insertSheet(tabName);
}

function safeGetTab(branchId, tabName) {
  try { return branchSS(branchId).getSheetByName(tabName); } catch(ex) { return null; }
}

function todayStr()  { return Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy"); }
function nowIST()    { return Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MMM-yyyy HH:mm:ss"); }
function monthName() {
  const ist = new Date(new Date().toLocaleString("en-US", {timeZone:"Asia/Kolkata"}));
  return ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][ist.getMonth()]+" "+ist.getFullYear();
}
function hdrStyle(rng, bg, fg) { rng.setBackground(bg).setFontColor(fg).setFontWeight("bold").setHorizontalAlignment("center"); }

function branchDisplayName(branchId) {
  try { const rows=masterTab("Branches").getDataRange().getValues(); for(let i=1;i<rows.length;i++){if(rows[i][0]===branchId)return rows[i][1];} } catch(ex) {}
  return "Green Salon";
}

function activeStaffNames(ss) {
  const st = ss.getSheetByName("Staff"); if (!st||st.getLastRow()<2) return [];
  return st.getDataRange().getValues().slice(1).filter(r=>r[6]!==false&&r[6]!=="FALSE").map(r=>String(r[1]));
}

// ── LAST UPDATE POLLING ───────────────────────────────────────
// Staff app polls this every 20s — if changed, reloads services/products/staff
function touchLastUpdate(branchId) {
  try {
    const ss = branchSS(branchId);
    let sh = ss.getSheetByName("_meta") || ss.insertSheet("_meta");
    const ts = nowIST();
    if (sh.getLastRow() === 0) { sh.appendRow(["lastUpdate", ts]); return; }
    const rows = sh.getDataRange().getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === "lastUpdate") { sh.getRange(i+1,2).setValue(ts); return; }
    }
    sh.appendRow(["lastUpdate", ts]);
  } catch(ex) { Logger.log("touchLastUpdate err: " + ex.message); }
}

function getLastUpdate(d) {
  try {
    const ss = branchSS(d.branchId);
    const sh = ss.getSheetByName("_meta"); if (!sh||sh.getLastRow()===0) return {lastUpdate:""};
    const rows = sh.getDataRange().getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === "lastUpdate") return {lastUpdate: String(rows[i][1]||"")};
    }
  } catch(ex) {}
  return {lastUpdate:""};
}

// ── FIRST TIME SETUP ──────────────────────────────────────────
function firstTimeSetup() {
  // Master sheet tabs
  const bsh = masterTab("Branches");
  if (bsh.getLastRow()===0) { bsh.appendRow(["BranchID","Name","Location","SheetID","Active","CreatedAt","DeletedAt"]); hdrStyle(bsh.getRange(1,1,1,7),C_RAW,C_WHITE); }
  const set = masterTab("Settings");
  if (set.getLastRow()===0) { set.appendRow(["BranchID","Email1","Email2","Email3","UpdatedAt"]); hdrStyle(set.getRange(1,1,1,5),C_RAW,C_WHITE); }

  const defs = [
    {id:"branch1",name:"Branch 1",loc:"JC Nagar"},
    {id:"branch2",name:"Branch 2",loc:"Koramangala"},
    {id:"branch3",name:"Branch 3",loc:"Indiranagar"},
  ];
  const existing = bsh.getDataRange().getValues().map(r=>r[0]);
  defs.forEach(b => {
    const sid = BRANCH_SHEETS[b.id];
    if (sid.includes("_SHEET_ID_HERE")) { Logger.log("⚠️ "+b.id+" sheet ID not set — skipping"); return; }
    if (!existing.includes(b.id)) bsh.appendRow([b.id,b.name,b.loc,sid,true,nowIST(),""]);
    try { initBranch(sid, b.name); } catch(ex) { Logger.log("❌ "+b.name+": "+ex.message); }
  });
  Logger.log("✅ Setup done. Run setupTriggers(), then Deploy as Web App.");
}

function initBranch(sheetId, branchName) {
  const ss = SpreadsheetApp.openById(sheetId);
  ensureTab(ss,"Staff",["ID","Name","PIN","Commission%","HasCommission","PhotoURL","Active"],7,C_RAW,()=>{
    const st=ss.getSheetByName("Staff");
    st.appendRow(["S001","Staff 1","1111",40,true,"",true]);
    st.appendRow(["S002","Staff 2","2222",40,true,"",true]);
    st.appendRow(["S003","Staff 3","3333",35,true,"",true]);
  });
  ensureTab(ss,"Services",["ServiceName","Price","Active"],3,C_RAW,()=>{
    const sv=ss.getSheetByName("Services");
    [["Haircut",150],["Shave",80],["Facial",300],["Hair Colour",500],["Head Massage",100],["Beard Trim",60],["Threading",40],["Waxing",200]].forEach(r=>sv.appendRow([r[0],r[1],true]));
  });
  ensureTab(ss,"Products",["ProductName","Price","Active"],3,C_RAW,()=>{
    const pd=ss.getSheetByName("Products");
    [["Shampoo",200],["Hair Oil",150],["Conditioner",180],["Hair Serum",250]].forEach(r=>pd.appendRow([r[0],r[1],true]));
  });
  ensureTab(ss,"Entries",["RowID","Timestamp","Date","StaffID","StaffName","Service","Amount","Tip","Payment","CommApplies","Flagged"],11,C_RAW);
  ensureTab(ss,"ProductSales",["RowID","Timestamp","Date","StaffID","StaffName","Product","Amount","Payment","Flagged"],9,C_RAW);
  ensureTab(ss,"Expenses",["RowID","Timestamp","Date","StaffID","StaffName","Description","Amount","Payment","Flagged"],9,C_RAW);
  buildDailyTab(ss, branchName);
  buildMonthlyTab(ss, branchName, monthName());
  Logger.log("✅ "+branchName+" initialized");
}

// ensureTab — creates tab if missing, adds header if empty, never deletes data
function ensureTab(ss, name, header, cols, bg, onNew) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name); sh.appendRow(header);
    hdrStyle(sh.getRange(1,1,1,cols),bg,C_WHITE);
    if (onNew) onNew();
    return sh;
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(header); hdrStyle(sh.getRange(1,1,1,cols),bg,C_WHITE);
    if (onNew) onNew();
    return sh;
  }
  // Repair missing columns without deleting data
  const ec = sh.getLastColumn();
  if (ec < cols) {
    for (let c = ec+1; c <= cols; c++) sh.getRange(1,c).setValue(header[c-1]);
    hdrStyle(sh.getRange(1,1,1,cols),bg,C_WHITE);
  }
  return sh;
}

// ── FIX / RECOVERY ────────────────────────────────────────────
function fixBranch(d) {
  const bid = d.branchId, ss = branchSS(bid), results = [];
  // Core data tabs
  const tabs = [
    {name:"Staff",        header:["ID","Name","PIN","Commission%","HasCommission","PhotoURL","Active"],       cols:7},
    {name:"Services",     header:["ServiceName","Price","Active"],                                            cols:3},
    {name:"Products",     header:["ProductName","Price","Active"],                                            cols:3},
    {name:"Entries",      header:["RowID","Timestamp","Date","StaffID","StaffName","Service","Amount","Tip","Payment","CommApplies","Flagged"], cols:11},
    {name:"ProductSales", header:["RowID","Timestamp","Date","StaffID","StaffName","Product","Amount","Payment","Flagged"], cols:9},
    {name:"Expenses",     header:["RowID","Timestamp","Date","StaffID","StaffName","Description","Amount","Payment","Flagged"], cols:9},
  ];
  tabs.forEach(t => {
    const before = ss.getSheetByName(t.name) ? (ss.getSheetByName(t.name).getLastRow()===0?"EMPTY":"EXISTS") : "MISSING";
    ensureTab(ss, t.name, t.header, t.cols, C_RAW);
    const after = ss.getSheetByName(t.name) ? "OK" : "FAIL";
    results.push(t.name+": "+before+" → "+after);
  });

  // Daily tab
  const staffNames = activeStaffNames(ss);
  const daily = ss.getSheetByName("Daily");
  if (!daily || daily.getLastRow()===0) {
    if (!daily) ss.insertSheet("Daily");
    const nd = ss.getSheetByName("Daily"); if(nd.getLastRow()>0) nd.deleteRows(1,nd.getLastRow());
    buildDailyHeader(nd, staffNames); results.push("Daily: REBUILT");
  } else {
    const map = dailyColMap(ss); let added = [];
    staffNames.forEach(n => { if (!map[n]) { ensureStaffInDaily(ss,n); added.push(n); } });
    results.push("Daily: OK" + (added.length ? " (added: "+added.join(",")+")":""));
  }

  // Monthly tab
  const tab = monthName();
  if (!ss.getSheetByName(tab)) { buildMonthlyTab(ss, branchDisplayName(bid), tab); results.push("Monthly "+tab+": CREATED"); }
  else results.push("Monthly "+tab+": OK");

  touchLastUpdate(bid);
  return {fixed:true, details:results};
}

function masterFix() {
  const rows = masterTab("Branches").getDataRange().getValues(); const all = {};
  rows.slice(1).forEach(row => {
    if (row[4]===false||row[4]==="FALSE") return;
    try { all[String(row[1])] = fixBranch({branchId:row[0]}).details; }
    catch(ex) { all[String(row[1])] = ["ERROR: "+ex.message]; }
  });
  return {fixed:true, results:all};
}

// ── TRIGGERS ──────────────────────────────────────────────────
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (["midnightReset","checkMonthEnd","sendDailyReport","sendMonthlyReport"].includes(t.getHandlerFunction()))
      ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("midnightReset").timeBased().everyDays(1).atHour(23).nearMinute(55).inTimezone("Asia/Kolkata").create();
  ScriptApp.newTrigger("checkMonthEnd").timeBased().everyDays(1).atHour(0).nearMinute(5).inTimezone("Asia/Kolkata").create();
  ScriptApp.newTrigger("sendDailyReport").timeBased().everyDays(1).atHour(23).nearMinute(0).inTimezone("Asia/Kolkata").create();
  ScriptApp.newTrigger("sendMonthlyReport").timeBased().everyDays(1).atHour(22).nearMinute(55).inTimezone("Asia/Kolkata").create();
  Logger.log("✅ 4 triggers set.");
}

function midnightReset() {
  masterTab("Branches").getDataRange().getValues().slice(1).forEach(row => {
    if (row[4]===false||row[4]==="FALSE") return;
    try {
      const sh = SpreadsheetApp.openById(row[3]).getSheetByName("Daily");
      if (sh && sh.getLastRow()>1) sh.deleteRows(2, sh.getLastRow()-1);
    } catch(ex) { Logger.log("reset err "+row[1]+": "+ex.message); }
  });
}

function checkMonthEnd() {
  const tab = monthName();
  masterTab("Branches").getDataRange().getValues().slice(1).forEach(row => {
    if (row[4]===false||row[4]==="FALSE") return;
    try {
      const ss = SpreadsheetApp.openById(row[3]);
      if (!ss.getSheetByName(tab)) buildMonthlyTab(ss, row[1], tab);
    } catch(ex) { Logger.log("monthEnd err: "+ex.message); }
  });
}

// ── DAILY TAB ─────────────────────────────────────────────────
function buildDailyHeader(sh, staffNames) {
  const h = []; staffNames.forEach(n=>{h.push(n);h.push(n+" Time");});
  h.push("Product"); h.push("Product Time");
  sh.appendRow(h); hdrStyle(sh.getRange(1,1,1,h.length),C_DARK,C_WHITE); sh.setFrozenRows(1);
  for(let c=1;c<=h.length;c++) sh.setColumnWidth(c, String(h[c-1]).endsWith(" Time")?185:78);
}

function buildDailyTab(ss, branchName) {
  let sh = ss.getSheetByName("Daily") || ss.insertSheet("Daily");
  if (sh.getLastRow()>0) return sh;
  buildDailyHeader(sh, activeStaffNames(ss));
  return sh;
}

function dailyColMap(ss) {
  const sh = ss.getSheetByName("Daily"); if (!sh||sh.getLastRow()===0) return {};
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; const m = {};
  h.forEach((v,i) => { if(v && !String(v).endsWith(" Time")) m[String(v)] = {amt:i+1, time:i+2}; });
  return m;
}

function ensureStaffInDaily(ss, name) {
  const sh = ss.getSheetByName("Daily"); if (!sh) return;
  const map = dailyColMap(ss); if (map[name]) return;
  // Insert before Product columns
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const pi = h.indexOf("Product");
  if (pi >= 0) {
    sh.insertColumnsBefore(pi+1, 2);
    sh.getRange(1,pi+1).setValue(name); sh.getRange(1,pi+2).setValue(name+" Time");
    hdrStyle(sh.getRange(1,pi+1,1,2),C_DARK,C_WHITE);
    sh.setColumnWidth(pi+1,78); sh.setColumnWidth(pi+2,185);
  } else {
    const last = sh.getLastColumn();
    sh.getRange(1,last+1).setValue(name); sh.getRange(1,last+2).setValue(name+" Time");
    hdrStyle(sh.getRange(1,last+1,1,2),C_DARK,C_WHITE);
    sh.setColumnWidth(last+1,78); sh.setColumnWidth(last+2,185);
  }
}

function writeDailyEntry(ss, staffName, amount, tip, payment, isProduct) {
  // Auto-fix: rebuild Daily tab if missing or empty
  let sh = ss.getSheetByName("Daily");
  if (!sh || sh.getLastRow()===0) { buildDailyTab(ss,""); sh = ss.getSheetByName("Daily"); }
  if (!isProduct) ensureStaffInDaily(ss, staffName);
  const map = dailyColMap(ss); const key = isProduct ? "Product" : staffName;
  const info = map[key]; if (!info) { Logger.log("writeDailyEntry: no col for "+key); return; }
  const val = amount + (tip>0 ? "+"+tip : "") + (payment==="Cash"?"C":"P");
  const ts  = Utilities.formatDate(new Date(),"Asia/Kolkata","dd-MMM-yyyy hh:mm:ss a");
  const data = sh.getLastRow()>1 ? sh.getRange(2,info.amt,sh.getLastRow()-1,1).getValues() : [];
  let row = 2;
  for (let r=0;r<data.length;r++) { if(!data[r][0]){row=r+2;break;} if(r===data.length-1)row=data.length+2; }
  sh.getRange(row,info.amt).setValue(val).setHorizontalAlignment("center");
  if (info.time) sh.getRange(row,info.time).setValue(ts);
}

// ── MONTHLY TAB ───────────────────────────────────────────────
// ONE ROW PER DAY — cumulative totals per staff, product, expense
function buildMonthlyTab(ss, branchName, tabName) {
  if (ss.getSheetByName(tabName)) return ss.getSheetByName(tabName);
  const sh = ss.insertSheet(tabName);
  const staffNames = activeStaffNames(ss);
  const cols = ["Date",...staffNames,"Extra","Total","Product","Expenses","Commission","Online","Cash","Difference"];
  const nc = cols.length;
  sh.getRange(1,1,1,nc).merge().setValue(branchName).setBackground(C_DARK).setFontColor(C_WHITE).setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1,36);
  sh.getRange(2,1,1,nc).setValues([cols]); hdrStyle(sh.getRange(2,1,1,nc),C_MED,C_WHITE);
  sh.setRowHeight(2,26); sh.setFrozenRows(2); sh.setColumnWidth(1,110);
  for(let c=2;c<=nc;c++) sh.setColumnWidth(c,90);
  return sh;
}

function monthColMap(sh) {
  if (!sh||sh.getLastRow()<2) return {};
  const h = sh.getRange(2,1,1,sh.getLastColumn()).getValues()[0]; const m = {};
  h.forEach((v,i)=>{if(v)m[String(v)]=i+1;}); return m;
}

function ensureStaffInMonthly(sh, staffName, branchName) {
  if (!sh) return;
  const h = sh.getRange(2,1,1,sh.getLastColumn()).getValues()[0];
  if (h.includes(staffName)) return;
  const ei = h.indexOf("Extra"); if (ei<0) return;
  sh.insertColumnBefore(ei+1); sh.getRange(2,ei+1).setValue(staffName);
  hdrStyle(sh.getRange(2,ei+1,1,1),C_MED,C_WHITE); sh.setColumnWidth(ei+1,90);
  // Re-merge title
  const nc = sh.getLastColumn();
  sh.getRange(1,1,1,nc).merge().setValue(branchName).setBackground(C_DARK).setFontColor(C_WHITE).setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");
  for(let r=3;r<=sh.getLastRow();r++){if(String(sh.getRange(r,1).getValue())!=="TOTAL")sh.getRange(r,ei+1).setValue(0);}
}

// updateMonthly — one row per day, cumulative add
function updateMonthly(branchId, entryType, staffName, svcAmt, tipAmt, payment, prodAmt, expAmt) {
  const ss = branchSS(branchId); const bn = branchDisplayName(branchId); const tab = monthName();
  let sh = ss.getSheetByName(tab); if (!sh) sh = buildMonthlyTab(ss,bn,tab);
  if (entryType==="service"&&staffName&&svcAmt>0) ensureStaffInMonthly(sh,staffName,bn);
  const cm = monthColMap(sh); const dt = todayStr(); const nc = sh.getLastColumn();
  const allVals = sh.getLastRow()>=3 ? sh.getRange(3,1,sh.getLastRow()-2,1).getValues() : [];
  let dr=-1, tr=-1;
  allVals.forEach((row,idx)=>{ const v=String(row[0]).trim(),r=idx+3; if(v===dt)dr=r; if(v==="TOTAL")tr=r; });
  if (dr<0) {
    const z=new Array(nc).fill(0); z[0]=dt;
    if (tr>0) { sh.insertRowBefore(tr); sh.getRange(tr,1,1,nc).setValues([z]); dr=tr; }
    else { sh.appendRow(z); dr=sh.getLastRow(); }
    const bg=dr%2===0?"#ffffff":C_ALT;
    sh.getRange(dr,1,1,nc).setBackground(bg).setHorizontalAlignment("center");
    sh.getRange(dr,1).setHorizontalAlignment("left");
  }
  function addVal(k,v){ if(!k||!(Number(v)>0))return; const col=cm[k]; if(!col)return; sh.getRange(dr,col).setValue((Number(sh.getRange(dr,col).getValue())||0)+Number(v)); }
  if (entryType==="service") {
    if(svcAmt>0) addVal(staffName,svcAmt);
    if(tipAmt>0) addVal("Extra",tipAmt);
    addVal(payment==="Cash"?"Cash":"Online",(svcAmt||0)+(tipAmt||0));
  } else if (entryType==="product") {
    if(prodAmt>0){ addVal("Product",prodAmt); addVal(payment==="Cash"?"Cash":"Online",prodAmt); }
  } else if (entryType==="expense") {
    if(expAmt>0) addVal("Expenses",expAmt);
  }
  recalcMonthRow(sh,dr,cm,nc); rebuildTotal(sh,nc);
}

function recalcMonthRow(sh,row,cm,nc) {
  const FIXED = new Set(["Date","Extra","Total","Product","Expenses","Commission","Online","Cash","Difference"]);
  const h = sh.getRange(2,1,1,nc).getValues()[0]; let sum=0;
  h.forEach((v,i)=>{ if(v&&!FIXED.has(String(v))) sum+=Number(sh.getRange(row,i+1).getValue())||0; });
  if(cm["Total"])       sh.getRange(row,cm["Total"]).setValue(sum);
  if(cm["Commission"])  sh.getRange(row,cm["Commission"]).setValue(Math.round(sum*0.40));
  const onl = cm["Online"] ? (Number(sh.getRange(row,cm["Online"]).getValue())||0) : 0;
  const csh = cm["Cash"]   ? (Number(sh.getRange(row,cm["Cash"]).getValue())  ||0) : 0;
  if(cm["Difference"]) sh.getRange(row,cm["Difference"]).setValue(onl+csh-sum);
}

function rebuildTotal(sh,nc) {
  const last=sh.getLastRow(); let tr=-1;
  for(let r=3;r<=last;r++){if(String(sh.getRange(r,1).getValue())==="TOTAL"){tr=r;break;}}
  const sums=new Array(nc).fill(0);
  const endR=tr>0?tr:last+1;
  for(let r=3;r<endR;r++){const v=sh.getRange(r,1,1,nc).getValues()[0];for(let c=1;c<nc;c++)sums[c]+=Number(v[c])||0;}
  const totalRow=["TOTAL",...sums.slice(1)];
  if(tr<0){sh.appendRow(totalRow);tr=sh.getLastRow();}else sh.getRange(tr,1,1,nc).setValues([totalRow]);
  hdrStyle(sh.getRange(tr,1,1,nc),C_DARK,C_WHITE); sh.getRange(tr,1).setHorizontalAlignment("left");
}

// ── AUTH ──────────────────────────────────────────────────────
function ownerLogin(d) { if(d.password!==OWNER_PASSWORD)throw new Error("Wrong password"); return{ownerName:"Harsha"}; }

// ── BRANCHES ──────────────────────────────────────────────────
function getBranches() {
  const rows = masterTab("Branches").getDataRange().getValues();
  const all = rows.slice(1).map(r=>({id:r[0],name:r[1],location:r[2],sheetId:r[3],active:r[4]!==false&&r[4]!=="FALSE",deletedAt:r[6]||""}));
  return {branches:all.filter(b=>b.active), deleted:all.filter(b=>!b.active)};
}
function addBranch(d) {
  if(!d.name||!d.sheetId) throw new Error("Name and SheetID required");
  try{SpreadsheetApp.openById(d.sheetId);}catch(ex){throw new Error("Cannot access Sheet — must be same Gmail account");}
  const id="branch"+Date.now();
  masterTab("Branches").appendRow([id,d.name,d.location||"",d.sheetId,true,nowIST(),""]);
  initBranch(d.sheetId,d.name); return{branchId:id};
}
function removeBranch(d) {
  const sh=masterTab("Branches"),rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){sh.getRange(i+1,5).setValue(false);sh.getRange(i+1,7).setValue(nowIST());return{};}}
  throw new Error("Branch not found");
}
function recoverBranch(d) {
  const sh=masterTab("Branches"),rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){sh.getRange(i+1,5).setValue(true);sh.getRange(i+1,7).setValue("");return{name:rows[i][1]};}}
  throw new Error("Branch not found");
}
function renameBranch(d) {
  if(!d.newName) throw new Error("New name required");
  const sh=masterTab("Branches"),rows=sh.getDataRange().getValues(); let oldName="";
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){oldName=rows[i][1];sh.getRange(i+1,2).setValue(d.newName);break;}}
  try{const ss=branchSS(d.branchId),msh=ss.getSheetByName(monthName());if(msh){const nc=msh.getLastColumn();msh.getRange(1,1,1,nc).merge().setValue(d.newName).setBackground(C_DARK).setFontColor(C_WHITE).setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");}}catch(ex){}
  return{oldName,newName:d.newName};
}

// ── STAFF ─────────────────────────────────────────────────────
function getStaff(bid)    { return{staff:_staffRows(bid)}; }
function getStaffAdmin(d) { return{staff:_staffRows(d.branchId)}; }
function _staffRows(bid) {
  return branchTab(bid,"Staff").getDataRange().getValues().slice(1)
    .filter(r=>r[6]!==false&&r[6]!=="FALSE")
    .map(r=>({id:r[0],name:r[1],pin:r[2],photoUrl:r[5]||"",hasCommission:r[4],commissionPct:Number(r[3])||0}));
}
function addStaff(d) {
  const sh=branchTab(d.branchId,"Staff"); const id="S"+Date.now();
  sh.appendRow([id,d.name,d.pin||"0000",d.commissionPct||0,d.hasCommission!==false,d.photoUrl||"",true]);
  const ss=branchSS(d.branchId); const bn=branchDisplayName(d.branchId);
  ensureStaffInDaily(ss,d.name);
  const msh=ss.getSheetByName(monthName()); if(msh) ensureStaffInMonthly(msh,d.name,bn);
  touchLastUpdate(d.branchId);
  return{staffId:id};
}
function removeStaff(d) {
  const sh=branchTab(d.branchId,"Staff"),rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.staffId){sh.getRange(i+1,7).setValue(false);touchLastUpdate(d.branchId);return{};}}
  throw new Error("Staff not found");
}
function renameStaff(d) {
  if(!d.newName) throw new Error("New name required");
  const sh=branchTab(d.branchId,"Staff"),rows=sh.getDataRange().getValues(); let oldName="";
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.staffId){oldName=rows[i][1];sh.getRange(i+1,2).setValue(d.newName);break;}}
  if(!oldName) throw new Error("Staff not found");
  // Update Daily header
  const ss=branchSS(d.branchId); const daily=ss.getSheetByName("Daily");
  if(daily&&daily.getLastRow()>0){
    const h=daily.getRange(1,1,1,daily.getLastColumn()).getValues()[0];
    h.forEach((v,i)=>{ const s=String(v); if(s===oldName)daily.getRange(1,i+1).setValue(d.newName); else if(s===oldName+" Time")daily.getRange(1,i+1).setValue(d.newName+" Time"); });
  }
  // Update Monthly header
  const msh=ss.getSheetByName(monthName());
  if(msh&&msh.getLastRow()>1){
    const h=msh.getRange(2,1,1,msh.getLastColumn()).getValues()[0];
    h.forEach((v,i)=>{ if(String(v)===oldName) msh.getRange(2,i+1).setValue(d.newName); });
  }
  touchLastUpdate(d.branchId);
  return{oldName,newName:d.newName};
}
function updateStaffComm(d) {
  const sh=branchTab(d.branchId,"Staff"),rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.staffId){sh.getRange(i+1,4).setValue(Number(d.commissionPct)||0);sh.getRange(i+1,5).setValue(d.hasCommission===true||d.hasCommission==="true");return{};}}
  throw new Error("Staff not found");
}

// ── SERVICES & PRODUCTS ───────────────────────────────────────
function getServices(bid){ return{services:branchTab(bid,"Services").getDataRange().getValues().slice(1).filter(r=>r[2]!==false&&r[2]!=="FALSE").map(r=>({name:r[0],price:r[1]}))}; }
function updateServices(d){
  const sh=branchTab(d.branchId,"Services");
  if(sh.getLastRow()>1) sh.deleteRows(2,sh.getLastRow()-1);
  if(d.services&&d.services.length) d.services.forEach(s=>sh.appendRow([s.name,s.price,true]));
  touchLastUpdate(d.branchId); return{};
}
function getProducts(bid){ return{products:branchTab(bid,"Products").getDataRange().getValues().slice(1).filter(r=>r[2]!==false&&r[2]!=="FALSE").map(r=>({name:r[0],price:r[1]}))}; }
function updateProducts(d){
  const sh=branchTab(d.branchId,"Products");
  if(sh.getLastRow()>1) sh.deleteRows(2,sh.getLastRow()-1);
  if(d.products&&d.products.length) d.products.forEach(p=>sh.appendRow([p.name,p.price,true]));
  touchLastUpdate(d.branchId); return{};
}

// ── SUBMIT ENTRIES ────────────────────────────────────────────
function submitEntry(d) {
  const{branchId,staffId,staffName,service,amount,tip,paymentMethod}=d;
  if(!branchId||!staffName||!service||!paymentMethod) throw new Error("Missing required fields");
  const amt=Number(amount)||0, tip2=Number(tip)||0;
  if(amt<=0) throw new Error("Amount must be > 0");
  // Auto-fix structure before writing
  const ss = branchSS(branchId);
  ensureTab(ss,"Entries",["RowID","Timestamp","Date","StaffID","StaffName","Service","Amount","Tip","Payment","CommApplies","Flagged"],11,C_RAW);
  const sRows=branchTab(branchId,"Staff").getDataRange().getValues(); let comm=true;
  for(let i=1;i<sRows.length;i++){if(sRows[i][0]===staffId){comm=sRows[i][4]===true||sRows[i][4]==="TRUE";break;}}
  const rid="E"+Date.now(),ts=nowIST(),dt=todayStr();
  branchTab(branchId,"Entries").appendRow([rid,ts,dt,String(staffId||""),String(staffName),String(service),amt,tip2,String(paymentMethod),comm,false]);
  writeDailyEntry(ss,staffName,amt,tip2,paymentMethod,false);
  updateMonthly(branchId,"service",staffName,amt,tip2,paymentMethod,0,0);
  return{rowId:rid,timestamp:ts};
}

function submitProduct(d) {
  const{branchId,staffId,staffName,product,amount,paymentMethod}=d;
  if(!branchId||!product||!paymentMethod) throw new Error("Missing required fields");
  const amt=Number(amount)||0; if(amt<=0) throw new Error("Amount must be > 0");
  const ss=branchSS(branchId);
  ensureTab(ss,"ProductSales",["RowID","Timestamp","Date","StaffID","StaffName","Product","Amount","Payment","Flagged"],9,C_RAW);
  const rid="P"+Date.now(),ts=nowIST(),dt=todayStr();
  branchTab(branchId,"ProductSales").appendRow([rid,ts,dt,String(staffId||"GLOBAL"),String(staffName||"Branch"),String(product),amt,String(paymentMethod),false]);
  writeDailyEntry(ss,"Product",amt,0,paymentMethod,true);
  updateMonthly(branchId,"product","",0,0,paymentMethod,amt,0);
  return{rowId:rid,timestamp:ts};
}

function submitExpense(d) {
  const{branchId,staffId,staffName,description,amount,paymentMethod}=d;
  if(!branchId||!description||!paymentMethod) throw new Error("Missing required fields");
  const amt=Number(amount)||0; if(amt<=0) throw new Error("Amount must be > 0");
  const ss=branchSS(branchId);
  ensureTab(ss,"Expenses",["RowID","Timestamp","Date","StaffID","StaffName","Description","Amount","Payment","Flagged"],9,C_RAW);
  const rid="X"+Date.now(),ts=nowIST(),dt=todayStr();
  branchTab(branchId,"Expenses").appendRow([rid,ts,dt,String(staffId||"GLOBAL"),String(staffName||"Branch"),String(description),amt,String(paymentMethod),false]);
  updateMonthly(branchId,"expense","",0,0,paymentMethod,0,amt);
  return{rowId:rid,timestamp:ts};
}

// ── GET ENTRIES ───────────────────────────────────────────────
function getMyEntries(d) {
  const{branchId,staffId}=d; const dt=todayStr();
  const sh=safeGetTab(branchId,"Entries"); if(!sh) return{entries:[],totalAmount:0,totalTip:0};
  const out=[]; let ta=0,tt=0;
  sh.getDataRange().getValues().slice(1).forEach(r=>{
    if(String(r[3])===String(staffId)&&String(r[2])===dt&&r[10]!==true&&r[10]!=="TRUE"){
      out.push({rowId:r[0],timestamp:r[1],service:r[5],amount:r[6],tip:r[7],paymentMethod:r[8]});
      ta+=Number(r[6])||0; tt+=Number(r[7])||0;
    }
  });
  return{entries:out,totalAmount:ta,totalTip:tt};
}

function getBranchSummary(d) {
  const bid=d.branchId; const dt=todayStr();
  const sh=safeGetTab(bid,"Entries"); if(!sh) return{totalEntries:0,totalRevenue:0,totalTips:0};
  let te=0,tr=0,tt=0;
  sh.getDataRange().getValues().slice(1).forEach(r=>{
    if(String(r[2])===dt&&r[10]!==true&&r[10]!=="TRUE"){te++;tr+=Number(r[6])||0;tt+=Number(r[7])||0;}
  });
  return{totalEntries:te,totalRevenue:tr,totalTips:tt};
}

function getTodayAll(d) {
  const bid=d.branchId; const dt=todayStr();
  const esh=safeGetTab(bid,"Entries"); const psh=safeGetTab(bid,"ProductSales"); const xsh=safeGetTab(bid,"Expenses");
  const entries=[],sm={};
  if(esh) esh.getDataRange().getValues().slice(1).forEach(r=>{
    if(String(r[2])!==dt) return;
    const fl=r[10]===true||r[10]==="TRUE";
    entries.push({rowId:r[0],timestamp:r[1],staffId:r[3],staffName:r[4],service:r[5],amount:r[6],tip:r[7],paymentMethod:r[8],commissionApplies:r[9],flagged:fl});
    if(!fl){const sn=String(r[4]);if(!sm[sn])sm[sn]={name:sn,totalAmount:0,totalTip:0,entries:0,products:0};sm[sn].totalAmount+=Number(r[6])||0;sm[sn].totalTip+=Number(r[7])||0;sm[sn].entries++;}
  });
  const ps=[];
  if(psh) psh.getDataRange().getValues().slice(1).forEach(r=>{
    if(String(r[2])!==dt) return;
    const fl=r[8]===true||r[8]==="TRUE";
    ps.push({rowId:r[0],timestamp:r[1],staffName:r[4],product:r[5],amount:r[6],paymentMethod:r[7],flagged:fl});
    if(!fl){const sn=String(r[4]);if(!sm[sn])sm[sn]={name:sn,totalAmount:0,totalTip:0,entries:0,products:0};sm[sn].products+=Number(r[6])||0;}
  });
  const xs=[]; let xe=0;
  if(xsh) xsh.getDataRange().getValues().slice(1).forEach(r=>{
    if(String(r[2])!==dt) return;
    const fl=r[8]===true||r[8]==="TRUE";
    xs.push({rowId:r[0],timestamp:r[1],staffName:r[4],description:r[5],amount:r[6],paymentMethod:r[7],flagged:fl});
    if(!fl) xe+=Number(r[6])||0;
  });
  return{entries,staffTotals:Object.values(sm),productSales:ps,expenses:xs,totalExp:xe};
}

function getMonthSummary(d) {
  const bid=d.branchId; const tab=monthName();
  const sh=safeGetTab(bid,tab); if(!sh) return{summary:[],month:tab};
  const all=sh.getDataRange().getValues(); if(all.length<3) return{summary:[],month:tab};
  const headers=all[1];
  return{summary:all.slice(2).filter(r=>r[0]).map(r=>{const o={};headers.forEach((k,i)=>{o[String(k)]=r[i];});return o;}),month:tab};
}

// ── DELETE (flag, never hard delete) ─────────────────────────
function deleteEntry(d)   {const sh=branchTab(d.branchId,"Entries"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.rowId){sh.getRange(i+1,11).setValue(true);return{};}}throw new Error("Not found");}
function deleteProduct(d) {const sh=branchTab(d.branchId,"ProductSales"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.rowId){sh.getRange(i+1,9).setValue(true);return{};}}throw new Error("Not found");}
function deleteExpense(d) {const sh=branchTab(d.branchId,"Expenses"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.rowId){sh.getRange(i+1,9).setValue(true);return{};}}throw new Error("Not found");}

// ── EMAIL REPORTS ─────────────────────────────────────────────
function setReportEmails(d){const sh=masterTab("Settings"),rows=sh.getDataRange().getValues();const emails=d.emails||[];for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){sh.getRange(i+1,2).setValue(emails[0]||"");sh.getRange(i+1,3).setValue(emails[1]||"");sh.getRange(i+1,4).setValue(emails[2]||"");sh.getRange(i+1,5).setValue(nowIST());return{};}}sh.appendRow([d.branchId,emails[0]||"",emails[1]||"",emails[2]||"",nowIST()]);return{};}
function getReportEmails(d){const rows=masterTab("Settings").getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId)return{emails:[rows[i][1]||"",rows[i][2]||"",rows[i][3]||""].filter(Boolean)};}return{emails:[]};}
function getBranchEmails(bid){const rows=masterTab("Settings").getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===bid)return[rows[i][1],rows[i][2],rows[i][3]].filter(Boolean);}return[];}
function getSheetAsCSV(sheetId,name){try{const ss=SpreadsheetApp.openById(sheetId),sh=ss.getSheetByName(name);if(!sh||sh.getLastRow()<1)return"";return sh.getDataRange().getValues().map(row=>row.map(c=>{const s=String(c).replace(/"/g,'""');return s.includes(",")||s.includes('"')||s.includes("\n")?`"${s}"`:s;}).join(",")).join("\n");}catch(ex){return "";}}
function _branchSheetId(bid){const rows=masterTab("Branches").getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===bid)return rows[i][3];}return null;}

function sendDailyReport(){masterTab("Branches").getDataRange().getValues().slice(1).forEach(row=>{if(row[4]===false||row[4]==="FALSE")return;const emails=getBranchEmails(row[0]);if(!emails.length)return;try{_sendDailyCSV(row[0],row[1],row[3],emails);}catch(ex){Logger.log("daily err "+row[1]+": "+ex.message);}});}
function _sendDailyCSV(bid,bname,sid,emails){
  const dt=todayStr(),dtF=dt.replace(/-/g,""),atts=[];
  ["Entries","ProductSales","Expenses","Daily"].forEach(tab=>{const csv=getSheetAsCSV(sid,tab);if(csv)atts.push(Utilities.newBlob(csv,"text/csv",tab+"_"+dtF+".csv"));});
  const body=`Daily report for ${bname}.\nDate: ${dt}\n\nRegards,\nGreen Salon System`;
  emails.forEach(email=>{try{MailApp.sendEmail({to:email,subject:`Green Salon — ${bname} — Daily Report — ${dt}`,body,attachments:atts});}catch(ex){Logger.log("send err "+email);}});
}

function sendMonthlyReport(){const now=new Date(),ist=new Date(now.toLocaleString("en-US",{timeZone:"Asia/Kolkata"})),last=new Date(ist.getFullYear(),ist.getMonth()+1,0).getDate();if(ist.getDate()!==last)return;masterTab("Branches").getDataRange().getValues().slice(1).forEach(row=>{if(row[4]===false||row[4]==="FALSE")return;const emails=getBranchEmails(row[0]);if(!emails.length)return;try{_sendMonthlyCSV(row[0],row[1],row[3],emails);}catch(ex){Logger.log("monthly err: "+ex.message);}});}
function _sendMonthlyCSV(bid,bname,sid,emails){const tab=monthName(),csv=getSheetAsCSV(sid,tab);if(!csv)return;const ist=new Date(new Date().toLocaleString("en-US",{timeZone:"Asia/Kolkata"})),ym=ist.getFullYear()+"-"+String(ist.getMonth()+1).padStart(2,"0"),blob=Utilities.newBlob(csv,"text/csv","Monthly_Report_"+ym+".csv"),body=`Monthly report for ${bname}.\nMonth: ${tab}\n\nRegards,\nGreen Salon System`;emails.forEach(email=>{try{MailApp.sendEmail({to:email,subject:`Green Salon — ${bname} — Monthly Report — ${tab}`,body,attachments:[blob]});}catch(ex){Logger.log("monthly send err "+email);}});}

function sendManualReport(d){
  const emails=getBranchEmails(d.branchId); if(!emails.length) throw new Error("No email addresses set for this branch");
  const bname=branchDisplayName(d.branchId); const sid=_branchSheetId(d.branchId); if(!sid) throw new Error("Branch sheet not found");
  if(d.type==="monthly") _sendMonthlyCSV(d.branchId,bname,sid,emails);
  else _sendDailyCSV(d.branchId,bname,sid,emails);
  return{sent:emails.length,recipients:emails};
}
