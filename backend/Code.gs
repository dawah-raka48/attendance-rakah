/**
 * Attendance Rakah - Google Apps Script backend starter
 * Deploy as Web App after setting up the project.
 * Never place passwords or admin tokens in the public GitHub frontend.
 */
const APP = {
  DB_ID_KEY: 'ATTENDANCE_DB_ID',
  TOKEN_TTL_SECONDS: 604800,
  SHEETS: {
    employees: ['code','name','department','deviceId','active','createdAt','updatedAt'],
    attendance: ['id','date','code','shift','time','type','source','createdAt','createdBy'],
    users: ['username','name','role','salt','passwordHash','active','createdAt'],
    settings: ['key','value'],
    audit: ['id','actor','action','entity','entityId','details','createdAt'],
    deviceLogs: ['eventId','deviceUserId','deviceUid','deviceName','timestamp','status','punch','createdAt']
  },
  ROLES: ['admin','hr','viewer']
};

function doGet(e) {
  return json_({ ok: true, service: 'attendance-rakah-api', version: '0.1.0', message: 'API is running. Send POST requests for actions.' });
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(payload.action || '');
    if (action === 'login') return json_(login_(payload));
    if (action === 'bootstrapStatus') return json_(bootstrapStatus_());
    if (action === 'syncDeviceData') return json_({ ok: true, data: syncDeviceData_(payload) });
    if (action === 'changePassword') { const user = authorize_(payload.token); return json_(changePassword_(user, payload.currentPassword, payload.newPassword, payload.token)); }
    const user = authorize_(payload.token);
    let result;
    switch (action) {
      case 'dashboard': result = dashboard_(user); break;
      case 'listEmployees': result = listEmployees_(user); break;
      case 'saveEmployee': requireRole_(user, ['admin','hr']); result = saveEmployee_(user, payload.employee || {}); break;
      case 'deactivateEmployee': requireRole_(user, ['admin','hr']); result = deactivateEmployee_(user, payload.code); break;
      case 'listAttendance': result = listAttendance_(user, payload.filters || {}); break;
      case 'savePunch': requireRole_(user, ['admin','hr']); result = savePunch_(user, payload.punch || {}); break;
      case 'getSettings': result = getSettings_(user); break;
      case 'saveSettings': requireRole_(user, ['admin']); result = saveSettings_(user, payload.settings || {}); break;
      case 'listUsers': requireRole_(user, ['admin']); result = listUsers_(user); break;
      case 'createUser': requireRole_(user, ['admin']); result = createUser_(user, payload.account || {}); break;
      default: throw new Error('UNKNOWN_ACTION');
    }
    return json_({ ok: true, data: result });
  } catch (err) {
    return json_({ ok: false, error: safeError_(err) });
  }
}

/** Run once from the Apps Script editor. Creates sheets and a one-time admin password. */
function setupAttendanceDatabase() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const existingId = props.getProperty(APP.DB_ID_KEY);
  if (existingId) {
    ss = SpreadsheetApp.openById(existingId);
  } else {
    ss = SpreadsheetApp.create('قاعدة بيانات نظام الحضور والانصراف - جمعية الراكة');
    props.setProperty(APP.DB_ID_KEY, ss.getId());
  }
  Object.keys(APP.SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = APP.SHEETS[name];
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    sh.setFrozenRows(1);
  });
  const defaultSettings = {
    morningStart: '09:00', morningEnd: '11:30', morningGrace: '0',
    eveningStart: '16:00', eveningEnd: '21:00', eveningGrace: '0',
    timezone: 'Asia/Riyadh'
  };
  const settings = sheet_('settings');
  const existingKeys = settings.getLastRow() > 1 ? settings.getRange(2,1,settings.getLastRow()-1,1).getValues().flat() : [];
  Object.keys(defaultSettings).forEach(k => {
    if (!existingKeys.includes(k)) settings.appendRow([k, defaultSettings[k]]);
  });
  const users = sheet_('users');
  if (users.getLastRow() < 2) {
    const password = randomPassword_();
    const salt = Utilities.getUuid();
    users.appendRow(['admin','مدير النظام','admin',salt,hash_(password,salt),true,new Date()]);
    Logger.log('INITIAL ADMIN USERNAME: admin');
    Logger.log('INITIAL ADMIN PASSWORD (save it now): ' + password);
  }
  Logger.log('DATABASE READY: ' + ss.getUrl());
  return ss.getUrl();
}

function bootstrapStatus_() {
  const id = PropertiesService.getScriptProperties().getProperty(APP.DB_ID_KEY);
  return { configured: !!id, service: 'attendance-rakah-api' };
}
function db_() {
  const id = PropertiesService.getScriptProperties().getProperty(APP.DB_ID_KEY);
  if (!id) throw new Error('DATABASE_NOT_CONFIGURED');
  return SpreadsheetApp.openById(id);
}
function sheet_(name) {
  const sh = db_().getSheetByName(name);
  if (!sh) throw new Error('MISSING_SHEET_' + name);
  return sh;
}
function rows_(name) {
  const sh = sheet_(name), values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values.shift();
  return values.filter(r => r.some(v => v !== '')).map(r => Object.fromEntries(headers.map((h,i) => [h, r[i]])));
}
function appendObject_(name, obj) {
  const sh = sheet_(name), headers = APP.SHEETS[name];
  sh.appendRow(headers.map(h => obj[h] === undefined ? '' : obj[h]));
}
function updateBy_(name, key, value, patch) {
  const sh=sheet_(name), values=sh.getDataRange().getValues(), headers=values[0], keyCol=headers.indexOf(key);
  if (keyCol < 0) throw new Error('INVALID_KEY');
  for (let i=1;i<values.length;i++) {
    if (String(values[i][keyCol]) === String(value)) {
      Object.keys(patch).forEach(k => { const c=headers.indexOf(k); if(c>=0) sh.getRange(i+1,c+1).setValue(patch[k]); });
      return true;
    }
  }
  return false;
}
function login_(p) {
  const username=String(p.username||'').trim(), password=String(p.password||'');
  if (!username || !password) throw new Error('MISSING_CREDENTIALS');
  const account=rows_('users').find(u => String(u.username)===username && u.active===true);
  if (!account || hash_(password,String(account.salt)) !== String(account.passwordHash)) throw new Error('INVALID_CREDENTIALS');
  const token=Utilities.getUuid()+Utilities.getUuid();
  const session = {username:account.username,name:account.name,role:account.role,expiresAt:Date.now()+APP.TOKEN_TTL_SECONDS*1000};
  PropertiesService.getScriptProperties().setProperty(sessionKey_(token), JSON.stringify(session));
  CacheService.getScriptCache().put('session_'+token, JSON.stringify(session), Math.min(APP.TOKEN_TTL_SECONDS, 21600));
  return { token, user: {username:account.username,name:account.name,role:account.role}, expiresIn:APP.TOKEN_TTL_SECONDS };
}
function sessionKey_(token) {
  const digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(token));
  return 'ATTENDANCE_SESSION_'+Utilities.base64EncodeWebSafe(digest).replace(/=+$/,'');
}
function authorize_(token) {
  if (!token) throw new Error('UNAUTHORIZED');
  const key=sessionKey_(token), props=PropertiesService.getScriptProperties();
  let raw=props.getProperty(key);
  if (!raw) raw=CacheService.getScriptCache().get('session_'+String(token));
  if (!raw) throw new Error('SESSION_EXPIRED');
  const session=JSON.parse(raw);
  if (session.expiresAt && Date.now()>Number(session.expiresAt)) {
    props.deleteProperty(key);
    CacheService.getScriptCache().remove('session_'+String(token));
    throw new Error('SESSION_EXPIRED');
  }
  return session;
}
function requireRole_(user, roles) { if (!roles.includes(user.role)) throw new Error('FORBIDDEN'); }
function listEmployees_(user) {
  return rows_('employees').map(e => ({...e, active: e.active === true || String(e.active)==='TRUE'}));
}
function saveEmployee_(user, e) {
  const code=String(e.code||'').trim(), name=String(e.name||'').trim(), department=String(e.department||'').trim();
  if (!code || !name || !department) throw new Error('EMPLOYEE_FIELDS_REQUIRED');
  const current=rows_('employees').find(x=>String(x.code)===code);
  const now=new Date();
  if (current) {
    updateBy_('employees','code',code,{name,department,deviceId:String(e.deviceId||''),updatedAt:now});
  } else {
    appendObject_('employees',{code,name,department,deviceId:String(e.deviceId||''),active:true,createdAt:now,updatedAt:now});
  }
  audit_(user.username,'saveEmployee','employee',code,{name,department});
  return {code};
}
function deactivateEmployee_(user, code) {
  if (!code || !updateBy_('employees','code',String(code),{active:false,updatedAt:new Date()})) throw new Error('EMPLOYEE_NOT_FOUND');
  audit_(user.username,'deactivateEmployee','employee',String(code),{});
  return {code};
}
function savePunch_(user,p) {
  const code=String(p.code||''), date=String(p.date||''), shift=String(p.shift||''), time=String(p.time||''), type=String(p.type||'in');
  if (!rows_('employees').some(e=>String(e.code)===code && (e.active===true || String(e.active)==='TRUE'))) throw new Error('EMPLOYEE_NOT_ACTIVE');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !/^(morning|evening)$/.test(shift) || !/^([01]\d|2[0-3]):[0-5]\d$/.test(time) || !/^(in|out)$/.test(type)) throw new Error('INVALID_PUNCH');
  const obj={id:Utilities.getUuid(),date,code,shift,time,type,source:String(p.source||'manual'),createdAt:new Date(),createdBy:user.username};
  appendObject_('attendance',obj); audit_(user.username,'savePunch','attendance',obj.id,obj);
  return obj;
}
function listAttendance_(user,f) {
  return rows_('attendance').filter(p=>(!f.from || String(p.date)>=String(f.from))&&(!f.to || String(p.date)<=String(f.to))&&(!f.code || String(p.code)===String(f.code))&&(!f.shift || String(p.shift)===String(f.shift)));
}
function getSettings_(user) { return Object.fromEntries(rows_('settings').map(r=>[r.key,String(r.value)])); }
function saveSettings_(user,s) {
  const allowed=['morningStart','morningEnd','morningGrace','eveningStart','eveningEnd','eveningGrace'];
  allowed.forEach(k=>{if(s[k]!==undefined){const ok=updateBy_('settings','key',k,{value:String(s[k])});if(!ok)appendObject_('settings',{key:k,value:String(s[k])});}});
  audit_(user.username,'saveSettings','settings','work-hours',s);
  return getSettings_(user);
}
function dashboard_(user) {
  const employees=rows_('employees').filter(e=>e.active===true || String(e.active)==='TRUE');
  const today=Utilities.formatDate(new Date(),'Asia/Riyadh','yyyy-MM-dd');
  const punches=rows_('attendance').filter(p=>String(p.date)===today);
  return {date:today,totalEmployees:employees.length,uniquePresent:new Set(punches.map(p=>String(p.code))).size,punches};
}
function listUsers_(user) { return rows_('users').map(u=>({username:u.username,name:u.name,role:u.role,active:u.active})); }
function changePassword_(user,currentPassword,newPassword,token) {
  if (!/^[A-Za-z0-9\u0600-\u06FF\u0660-\u0669]{4,12}$/.test(String(newPassword||''))) throw new Error('PASSWORD_MUST_BE_4_TO_12_LETTERS_OR_NUMBERS');
  const sh=sheet_('users'), values=sh.getDataRange().getValues(), headers=values[0];
  const ucol=headers.indexOf('username'), scol=headers.indexOf('salt'), hcol=headers.indexOf('passwordHash');
  for(let i=1;i<values.length;i++){
    if(String(values[i][ucol])===String(user.username)){
      const oldSalt=String(values[i][scol]), oldHash=String(values[i][hcol]);
      if(hash_(String(currentPassword||''),oldSalt)!==oldHash) throw new Error('CURRENT_PASSWORD_INCORRECT');
      const salt=Utilities.getUuid();
      sh.getRange(i+1,scol+1).setValue(salt); sh.getRange(i+1,hcol+1).setValue(hash_(newPassword,salt));
      CacheService.getScriptCache().remove('session_'+String(token));
      PropertiesService.getScriptProperties().deleteProperty(sessionKey_(token));
      audit_(user.username,'changePassword','user',user.username,{});
      return {changed:true,reauthenticate:true};
    }
  }
  throw new Error('USER_NOT_FOUND');
}
function createUser_(user,a) {
  const username=String(a.username||'').trim(), name=String(a.name||'').trim(), password=String(a.password||''), role=String(a.role||'viewer');
  if(!username||!name||!/^[A-Za-z0-9\u0600-\u06FF\u0660-\u0669]{4,12}$/.test(password)||!APP.ROLES.includes(role)) throw new Error('INVALID_USER_INPUT');
  if(rows_('users').some(u=>String(u.username)===username)) throw new Error('USERNAME_EXISTS');
  const salt=Utilities.getUuid();
  appendObject_('users',{username,name,role,salt,passwordHash:hash_(password,salt),active:true,createdAt:new Date()});
  audit_(user.username,'createUser','user',username,{role});
  return {username,role};
}

/**
 * Device-agent endpoint. Configure DEVICE_SYNC_KEY in Apps Script
 * Project Settings > Script Properties. Stores raw events only;
 * it does not infer in/out or shift from undocumented device status codes.
 */
function syncDeviceData_(p) {
  const expected = PropertiesService.getScriptProperties().getProperty('DEVICE_SYNC_KEY');
  if (!expected || String(p.deviceKey || '') !== expected) throw new Error('DEVICE_SYNC_UNAUTHORIZED');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('SYNC_BUSY');
  try {
    const users = Array.isArray(p.users) ? p.users : [];
    const records = Array.isArray(p.records) ? p.records : [];
    if (records.length > 300) throw new Error('BATCH_TOO_LARGE');
    const employeeSheet = sheet_('employees');
    const existingEmployees = rows_('employees');
    const employeeCodes = new Set(existingEmployees.map(e => String(e.code)));
    let employeesAdded = 0;
    users.forEach(u => {
      const code = String(u.userId || '').trim();
      if (!code || !/^\\d{1,20}$/.test(code) || employeeCodes.has(code)) return;
      appendObject_('employees', {
        code, name: 'موظف ' + code, department: 'غير محدد',
        deviceId: code, active: true, createdAt: new Date(), updatedAt: new Date()
      });
      employeeCodes.add(code);
      employeesAdded++;
    });

    const logRows = rows_('deviceLogs');
    const known = new Set(logRows.map(r => String(r.eventId)));
    let inserted = 0, duplicates = 0;
    records.forEach(r => {
      const code = String(r.userId || '').trim();
      const timestamp = String(r.timestamp || '').trim();
      if (!code || !/^\\d{1,20}$/.test(code) || !/^\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2}$/.test(timestamp)) return;
      const raw = [code, timestamp, String(r.status ?? ''), String(r.punch ?? ''), String(r.deviceUid ?? '')].join('|');
      const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
      const eventId = digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2,'0')).join('');
      if (known.has(eventId)) { duplicates++; return; }
      appendObject_('deviceLogs', {
        eventId, deviceUserId: code, deviceUid: String(r.deviceUid ?? ''),
        deviceName: String(r.deviceName || ''), timestamp,
        status: String(r.status ?? ''), punch: String(r.punch ?? ''),
        createdAt: new Date()
      });
      known.add(eventId);
      inserted++;
    });
    return { employeesAdded, recordsInserted: inserted, duplicates, received: records.length };
  } finally {
    lock.releaseLock();
  }
}

function audit_(actor,action,entity,entityId,details) {
  appendObject_('audit',{id:Utilities.getUuid(),actor,action,entity,entityId,details:JSON.stringify(details),createdAt:new Date()});
}
function hash_(password,salt) {
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,salt+'|'+password,Utilities.Charset.UTF_8);
  return bytes.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('');
}
function randomPassword_() {
  const chars='ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%';
  let out='';for(let i=0;i<20;i++)out+=chars[Math.floor(Math.random()*chars.length)];return out;
}
function safeError_(err) {
  const msg=String(err && err.message || err || 'UNKNOWN_ERROR');
  return msg.replace(/[\r\n]/g,' ').slice(0,160);
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
