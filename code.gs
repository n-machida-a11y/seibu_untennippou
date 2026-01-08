// ===========================================
// かんたん運転日報システム V3.6
// Google Apps Script バックエンドコード
// ===========================================

const DB_SS_ID = '1m-gg-S_WESgZffPBxabUyVLZrr1EqGgZRmVysiDhHF8';
const MASTER_SS_ID = '1Ic4bJeFm8VgQx7WTogNrTiwEUvskw1v0NOTQnD70GGw';

const SHEET_NAMES = {
  LOG_DATA: '記録蓄積用',
  ADMIN_SETTINGS: '管理者設定' 
};

const EXT_SHEET_NAMES = {
  EMPLOYEE: '社員マスタ',
  CARS: ['車両管理_本社', '車両管理_大阪', '車両管理_姫路']
};

const STATUS = {
  WORKING: '作業中',
  SAVED: '保存済',
  SUBMITTED: '提出済',
  APPROVED: '承認済',
  REJECTED: '差戻し'
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('かんたん運転日報 V3.6')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// --- ユーティリティ ---
function getDbSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(DB_SS_ID);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet && sheetName === SHEET_NAMES.ADMIN_SETTINGS) {
      sheet = ss.insertSheet(SHEET_NAMES.ADMIN_SETTINGS);
      sheet.appendRow(['社員番号', '氏名', '担当拠点']);
    }
    return sheet;
  } catch (e) { throw new Error('DB取得失敗: ' + e.message); }
}

function getMasterSheet(sheetName) {
  try { return SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(sheetName); } 
  catch (e) { return null; }
}

function generateUUID() { return Utilities.getUuid(); }

function calculateYearMonth(date) {
  const d = new Date(date);
  if (d.getDate() <= 20) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
  const next = new Date(d); next.setMonth(next.getMonth() + 1);
  return Utilities.formatDate(next, Session.getScriptTimeZone(), 'yyyy-MM');
}

function getMonthPeriod(yearMonth) {
  const [year, month] = yearMonth.split('-').map(Number);
  const s = new Date(year, month - 2, 21);
  const e = new Date(year, month - 1, 20);
  return {
    start: Utilities.formatDate(s, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    end: Utilities.formatDate(e, Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
}

function getCurrentYearMonth() { return calculateYearMonth(new Date()); }

// --- マスタ取得 ---
function fetchAllEmployees() {
  const sheet = getMasterSheet(EXT_SHEET_NAMES.EMPLOYEE);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const employees = {};
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0]).trim();
    const id = String(data[i][1]).trim();
    if (name && id) employees[name] = { name: name, userId: id, email: data[i][2] };
  }
  return employees;
}

function fetchAllCars() {
  let allCars = [];
  EXT_SHEET_NAMES.CARS.forEach(sheetName => {
    const sheet = getMasterSheet(sheetName);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        // H列(index 7)を登録番号として取得
        if (data[i][7]) {
          allCars.push({
            owner: String(data[i][1]), 
            userId: String(data[i][2]),
            number: String(data[i][7]).trim(), // H列: 登録番号
            // name: String(data[i][4]) + ' ' + String(data[i][5]) 
          });
        }
      }
    }
  });
  return allCars;
}

function fetchAdminSettings() {
  const sheet = getDbSheet(SHEET_NAMES.ADMIN_SETTINGS);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const admins = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    const base = String(data[i][2]).trim(); 
    if (id) admins[id] = { base: base };
  }
  return admins;
}

// --- ログイン・認証 ---
function login(userName) {
  try {
    const employees = fetchAllEmployees();
    const user = employees[userName];
    if (!user) return { success: false, message: '社員マスタ未登録' };
    
    const admins = fetchAdminSettings();
    const adminInfo = admins[user.userId];
    const role = adminInfo ? '管理者' : '使用者';
    const adminBase = adminInfo ? adminInfo.base : ''; 
    
    const cars = fetchAllCars();
    const assigned = cars.find(c => c.userId === user.userId);
    
    return {
      success: true,
      user: {
        userId: user.userId, name: user.name,
        carNumber: assigned ? assigned.number : '',
        role: role, adminBase: adminBase
      }
    };
  } catch (e) { return { success: false, message: e.message }; }
}

function getUserList() {
  try {
    const emps = fetchAllEmployees();
    const list = Object.values(emps).map(e => ({ userId: e.userId, name: e.name }))
      .sort((a,b) => a.userId.localeCompare(b.userId));
    return { success: true, users: list };
  } catch (e) { return { success: false, message: e.message }; }
}

function getCarList() {
  try {
    const cars = fetchAllCars();
    return { success: true, cars: cars.map(c => ({ carNumber: c.number, carType: '' })) };
  } catch (e) { return { success: false, message: e.message }; }
}

// --- 履歴取得（車両ごとにグループ化 + 追加車両対応） ---
function getHistory(userId, yearMonth, extraCarNumbers = []) {
  try {
    const sheet = getDbSheet(SHEET_NAMES.LOG_DATA);
    if (!sheet) return { success: true, carHistories: [] };

    const data = sheet.getDataRange().getValues();
    const period = getMonthPeriod(yearMonth);
    const startDate = new Date(period.start); startDate.setHours(0,0,0,0);
    const endDate = new Date(period.end); endDate.setHours(23,59,59,999);
    
    const holidays = {};
    try{
      const cal = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
      cal.getEvents(startDate, endDate).forEach(ev => holidays[Utilities.formatDate(ev.getStartTime(),Session.getScriptTimeZone(),'yyyy-MM-dd')]=true);
    }catch(err){}

    const allMasterCars = fetchAllCars();
    const myAssignedCar = allMasterCars.find(c => c.userId === String(userId));
    
    const usedCarNumbers = new Set();
    if (myAssignedCar) usedCarNumbers.add(myAssignedCar.number);
    // フロントから要求された追加車両もセット
    if (extraCarNumbers && Array.isArray(extraCarNumbers)) {
      extraCarNumbers.forEach(c => usedCarNumbers.add(c));
    }

    const carRecordsMap = {};
    let monthStatus = STATUS.WORKING;
    let hasSub=false, hasApp=false, hasRej=false;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[2]) !== String(userId)) continue;

      const rowDate = new Date(row[1]);
      if (isNaN(rowDate.getTime())) continue;

      if (rowDate >= startDate && rowDate <= endDate) {
        const dateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const carNum = String(row[4] || '未設定'); 
        
        usedCarNumbers.add(carNum);
        
        if (!carRecordsMap[carNum]) carRecordsMap[carNum] = {};
        
        const st = String(row[18] || '作業中');
        if (st === STATUS.APPROVED) hasApp=true;
        if (st === STATUS.SUBMITTED) hasSub=true;
        if (st === STATUS.REJECTED) hasRej=true;

        carRecordsMap[carNum][dateStr] = {
          id: String(row[0]),
          date: dateStr,
          carNumber: carNum,
          departureMeter: String(row[5] || ''),
          arrivalMeter: String(row[6] || ''),
          distance: String(row[7] || ''),
          destination: String(row[8] || ''),
          alcoholStart: String(row[9] || ''),
          alcoholEnd: String(row[11] || ''),
          dailyCheck: String(row[10] || ''),
          carCleaning: String(row[12] || ''),
          troubleStatus: String(row[13] || ''),
          troubleDetail: String(row[14] || ''),
          refuel: String(row[15] || ''),
          remarks: String(row[16] || ''),
          status: st,
          rejectComment: String(row[22] || ''),
          isSaved: true
        };
      }
    }

    if (hasApp) monthStatus = STATUS.APPROVED;
    else if (hasSub) monthStatus = STATUS.SUBMITTED;
    else if (hasRej) monthStatus = STATUS.REJECTED;

    const carHistories = [];
    
    usedCarNumbers.forEach(carNum => {
        const dailyList = [];
        const currentDate = new Date(startDate);
        
        while (currentDate <= endDate) {
            const dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            const dayOfWeek = currentDate.getDay();
            const isHoliday = !!holidays[dateStr];
            
            const record = (carRecordsMap[carNum] && carRecordsMap[carNum][dateStr]) ? carRecordsMap[carNum][dateStr] : null;
            
            if (record) {
                dailyList.push({ ...record, dayOfWeek: dayOfWeek, isHoliday: isHoliday });
            } else {
                dailyList.push({
                    id: '', date: dateStr, carNumber: carNum,
                    departureMeter: '', 
                    arrivalMeter: '', distance: '', destination: '',
                    alcoholStart: '', alcoholEnd: '', dailyCheck: '', carCleaning: '', troubleStatus: '',
                    refuel: '', remarks: '', status: '', rejectComment: '', isSaved: false,
                    dayOfWeek: dayOfWeek, isHoliday: isHoliday
                });
            }
            currentDate.setDate(currentDate.getDate() + 1);
        }
        
        carHistories.push({
            carNumber: carNum,
            records: dailyList
        });
    });
    
    return { success: true, carHistories: carHistories, monthStatus: monthStatus };
  } catch (e) {
    return { success: false, message: '履歴取得エラー: ' + e.message };
  }
}

function getMonthlyDistance(userId, yearMonth) {
  const r = getHistory(userId, yearMonth);
  if (!r.success) return r;
  let dist=0, days=0;
  
  r.carHistories.forEach(hist => {
      hist.records.forEach(re => {
        if(re.isSaved){
          const d = parseFloat(re.distance);
          if(!isNaN(d) && d>0){ dist+=d; days++; }
        }
      });
  });
  
  return {success:true, totalDistance:Math.round(dist*10)/10, workDays:days, status:r.monthStatus};
}

// --- メーター補完ロジック（直近の記録を探す） ---
function findPreviousArrival(carNumber, targetDateStr, sheetData) {
  let maxMeter = 0;
  const targetDate = new Date(targetDateStr);
  
  // ログデータを走査して、対象日以前の最も新しい帰着メーターを探す
  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    if (String(row[4]) === String(carNumber)) { // 車番一致
      const rowDate = new Date(row[1]);
      if (rowDate < targetDate) { // 日付が前
        const m = parseFloat(row[6]); // G列: 帰着メーター
        if (!isNaN(m)) {
          // 最も日付が新しいものを採用したいが、簡易的に最大値を取る（メーターは増えるものなので）
          // ※厳密には日付ソートすべきだが、ここではメーター戻りがない前提でMAXをとる
          if (m > maxMeter) maxMeter = m;
        }
      }
    }
  }
  return maxMeter;
}

// --- 一括保存（自動計算ロジック強化） ---
function saveHistoryBulk(userId, userName, updates) {
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(10000)) return {success:false, message:'他ユーザー編集中'};
  try {
    const sheet = getDbSheet(SHEET_NAMES.LOG_DATA);
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const idMap = {};
    for(let i=1; i<data.length; i++) idMap[String(data[i][0])] = i+1;
    
    updates.forEach(rec=>{
      let rIdx = -1;
      if(rec.id && idMap[rec.id]){
        rIdx = idMap[rec.id];
        const stat = data[rIdx-1][18];
        if(stat===STATUS.SUBMITTED || stat===STATUS.APPROVED) return;
      } else if(!rec.id && !rec.destination && !rec.arrivalMeter) return;

      const dStr = rec.date;
      const ym = calculateYearMonth(new Date(dStr));
      const arr = parseFloat(rec.arrivalMeter)||0;
      
      // ★出発メーター自動補完ロジック
      let dep = 0;
      if (rec.departureMeter !== undefined && rec.departureMeter !== '') {
          // 入力があればそれを使う
          dep = parseFloat(rec.departureMeter);
      } else if (rIdx > 0) {
          // 既存レコードならDBの値
          dep = parseFloat(data[rIdx-1][5]) || 0; 
      }
      
      // まだ0なら、過去データから検索して補完
      if (dep === 0 && arr > 0 && rec.carNumber) {
          dep = findPreviousArrival(rec.carNumber, dStr, data);
      }
      
      // 距離計算
      let dist = '';
      if (arr > 0) {
          dist = (arr > dep) ? (arr - dep) : 0;
      } else if (rec.distance) {
          dist = parseFloat(rec.distance);
      }

      const vals = {
          date: dStr, userId: userId, userName: userName, carNumber: rec.carNumber,
          dep: dep, arr: arr, dist: dist, dest: rec.destination,
          alcStart: rec.alcoholStart||'未', daily: rec.dailyCheck||'異常なし', 
          alcEnd: rec.alcoholEnd||'未', clean: rec.carCleaning||'未',
          trouble: rec.troubleStatus||'なし', detail: rec.troubleDetail||'',
          refuel: rec.refuel, rem: rec.remarks, ym: ym, now: now
      };

      if(rIdx>0){
        const r = rIdx;
        sheet.getRange(r,2).setValue(vals.date);
        sheet.getRange(r,5).setValue(vals.carNumber);
        // 補完したdepを書き込む
        sheet.getRange(r,6).setValue(vals.dep);
        sheet.getRange(r,7).setValue(vals.arr);
        sheet.getRange(r,8).setValue(vals.dist);
        sheet.getRange(r,9).setValue(vals.dest);
        sheet.getRange(r,10).setValue(vals.alcStart);
        sheet.getRange(r,11).setValue(vals.daily);
        sheet.getRange(r,12).setValue(vals.alcEnd);
        sheet.getRange(r,13).setValue(vals.clean);
        sheet.getRange(r,14).setValue(vals.trouble);
        sheet.getRange(r,15).setValue(vals.detail);
        sheet.getRange(r,16).setValue(vals.refuel);
        sheet.getRange(r,17).setValue(vals.rem);
        sheet.getRange(r,19).setValue(STATUS.SAVED);
        sheet.getRange(r,22).setValue(vals.now);
      } else {
        const uuid = generateUUID();
        const row = [
            uuid, vals.date, vals.userId, vals.userName, vals.carNumber, 
            vals.dep, vals.arr, vals.dist, vals.dest, 
            vals.alcStart, vals.daily, vals.alcEnd, vals.clean, 
            vals.trouble, vals.detail, vals.refuel, vals.rem, 
            vals.ym, STATUS.SAVED, '', '', vals.now, ''
        ];
        sheet.appendRow(row);
      }
    });
    return {success:true, message:'一括保存しました'};
  } catch(e){ return {success:false, message:e.message}; } finally { lock.releaseLock(); }
}

function saveArrival(data) {
  // saveHistoryBulkと同様の補完を入れる
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const sheet = getDbSheet(SHEET_NAMES.LOG_DATA);
    const now = new Date();
    const sData = sheet.getDataRange().getValues();
    
    let dep = parseFloat(data.departureMeter)||0;
    const arr = parseFloat(data.arrivalMeter)||0;
    
    // 補完
    if (dep === 0 && arr > 0) {
        dep = findPreviousArrival(data.carNumber, data.date, sData);
    }
    
    const dist = (arr > dep) ? (arr - dep) : 0;
    
    let rIdx = -1;
    if(data.id){
      for(let i=1; i<sData.length; i++){
        if(String(sData[i][0])===String(data.id)){ rIdx=i+1; break; }
      }
    }
    const ym = calculateYearMonth(new Date(data.date));
    if(rIdx>0){
      const st = sData[rIdx-1][18];
      if(st===STATUS.SUBMITTED || st===STATUS.APPROVED) return {success:false, message:'提出済のため編集不可'};
      const r = rIdx;
      const vals = [
          [data.date, data.userId, data.userName, data.carNumber, dep, arr, dist, data.destination,
           data.alcoholCheckStart, data.dailyCheck, data.alcoholCheckEnd, data.carCleaning,
           data.troubleStatus, data.troubleDetail, data.refuel, data.remarks, ym, STATUS.SAVED, '', '', now, '']
      ];
      sheet.getRange(r, 2, 1, 22).setValues(vals);
    } else {
      const uuid = generateUUID();
      sheet.appendRow([uuid, data.date, data.userId, data.userName, data.carNumber, dep, arr, dist, data.destination, data.alcoholCheckStart, data.dailyCheck, data.alcoholCheckEnd, data.carCleaning, data.troubleStatus, data.troubleDetail, data.refuel, data.remarks, ym, STATUS.SAVED, '', '', now, '']);
    }
    return {success:true, message:'保存しました'};
  } catch(e){ return {success:false, message:e.message}; } finally { lock.releaseLock(); }
}

// --- 他機能 ---
function submitMonthly(userId, yearMonth) {
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const sheet = getDbSheet(SHEET_NAMES.LOG_DATA);
    const data = sheet.getDataRange().getValues();
    const period = getMonthPeriod(yearMonth);
    const s = new Date(period.start); s.setHours(0,0,0,0);
    const e = new Date(period.end); e.setHours(23,59,59,999);
    let c = 0;
    for(let i=1; i<data.length; i++){
      const d = new Date(data[i][1]);
      if(String(data[i][2])===String(userId) && d>=s && d<=e && (data[i][18]===STATUS.SAVED || data[i][18]===STATUS.WORKING || data[i][18]===STATUS.REJECTED)){
        sheet.getRange(i+1, 19).setValue(STATUS.SUBMITTED);
        c++;
      }
    }
    return {success:true, message:c+'件 提出しました'};
  } finally { lock.releaseLock(); }
}

function testGetCurrentYearMonth() {
  const ym = getCurrentYearMonth();
  const p = getMonthPeriod(ym);
  return {yearMonth:ym, period:p};
}

function getLastMeter(carNumber) {
  try {
    const sheet = getDbSheet(SHEET_NAMES.LOG_DATA);
    if(!sheet) return {success:true, lastMeter:0};
    const data = sheet.getDataRange().getValues();
    let max = 0;
    for(let i=1; i<data.length; i++){
      if(String(data[i][4])===String(carNumber)){
        const m = parseFloat(data[i][6]); 
        if(!isNaN(m) && m>max) max=m;
      }
    }
    return {success:true, lastMeter:max};
  } catch(e){ return {success:false, lastMeter:0}; }
}

// 管理者機能系
function getAdminPendingList(adminUserId) {
  try {
    const admins = fetchAdminSettings();
    const adminInfo = admins[adminUserId];
    if (!adminInfo) return { success: false, message: '管理者権限なし' };
    const targetBase = adminInfo.base;
    
    const cars = fetchAllCars();
    const carBaseMap = {};
    cars.forEach(c => carBaseMap[c.number] = c.owner);

    const sheet = getDbSheet(SHEET_NAMES.LOG_DATA);
    const data = sheet.getDataRange().getValues();
    const grouped = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = String(row[18]);
      if (status === STATUS.SUBMITTED || status === STATUS.APPROVED) {
        const carNum = String(row[4]);
        const recordBase = carBaseMap[carNum] || '不明';
        
        if (targetBase && targetBase !== recordBase) continue;

        let ym = row[17];
        const userId = String(row[2]);
        const userName = String(row[3]);
        const key = userId + '_' + ym;

        if (!grouped[key]) {
          grouped[key] = { userId:userId, userName:userName, yearMonth:ym, status:status, totalDistance:0, workDays:0, base:recordBase };
        }
        if (status === STATUS.APPROVED) grouped[key].status = STATUS.APPROVED;
        
        const dist = parseFloat(row[7]);
        if (!isNaN(dist) && dist > 0) {
          grouped[key].totalDistance += dist;
          grouped[key].workDays++;
        }
      }
    }
    let list = Object.values(grouped);
    list.sort((a,b)=>(a.status===STATUS.SUBMITTED && b.status!==STATUS.SUBMITTED)?-1:1);
    return { success: true, list: list };
  } catch (e) { return { success: false, message: e.message }; }
}

function approveMonthly(adminId, adminName, targetUserId, yearMonth) {
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(10000)) return {success:false,message:'Busy'};
  try {
    const sheet = getDbSheet(SHEET_NAMES.LOG_DATA);
    const data = sheet.getDataRange().getValues();
    const period = getMonthPeriod(yearMonth);
    const s = new Date(period.start); s.setHours(0,0,0,0);
    const e = new Date(period.end); e.setHours(23,59,59,999);
    const now = new Date();
    let c = 0;
    for(let i=1; i<data.length; i++){
      const d = new Date(data[i][1]);
      if(String(data[i][2])===String(targetUserId) && d>=s && d<=e && String(data[i][18])===STATUS.SUBMITTED){
        sheet.getRange(i+1, 19).setValue(STATUS.APPROVED);
        sheet.getRange(i+1, 20).setValue(adminName);
        sheet.getRange(i+1, 21).setValue(now);
        c++;
      }
    }
    return {success:true, message:c+'件 承認'};
  } finally { lock.releaseLock(); }
}

function rejectMonthly(adminId, targetUserId, yearMonth, comment) {
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(10000)) return {success:false,message:'Busy'};
  try {
    const sheet = getDbSheet(SHEET_NAMES.LOG_DATA);
    const data = sheet.getDataRange().getValues();
    const period = getMonthPeriod(yearMonth);
    const s = new Date(period.start); s.setHours(0,0,0,0);
    const e = new Date(period.end); e.setHours(23,59,59,999);
    let c = 0;
    for(let i=1; i<data.length; i++){
      const d = new Date(data[i][1]);
      const st = String(data[i][18]);
      if(String(data[i][2])===String(targetUserId) && d>=s && d<=e && (st===STATUS.SUBMITTED||st===STATUS.APPROVED)){
        sheet.getRange(i+1, 19).setValue(STATUS.REJECTED);
        sheet.getRange(i+1, 23).setValue(comment);
        sheet.getRange(i+1, 20).clearContent();
        sheet.getRange(i+1, 21).clearContent();
        c++;
      }
    }
    return {success:true, message:c+'件 差戻し'};
  } finally { lock.releaseLock(); }
}

// プレースホルダー
function saveDeparture() {}
function updateCarLastMeter() {}
function addEmployee() {} function updateEmployee() {} function deleteEmployee() {}
function addVehicle() {} function updateVehicle() {} function deleteVehicle() {}