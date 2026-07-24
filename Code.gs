/**
 * ============================================================
 *  會議室管理系統 - Google Apps Script 後端（最終版）
 *  部署為 Web App (Execute as: Me, Access: Anyone)
 * ============================================================
 */

const SPREADSHEET_ID   = '1TEtUFfiLkQNDMUbIrda4NeyAu61KpXMVa_LovRfw_XQ';
const SHEET_USERS      = 'Users';
const SHEET_BOOKINGS   = 'Bookings';
const SHEET_SSO_TOKENS = 'SsoTokens'; // Intranet SSO 共用 token 表

/* ══════════════════════════════════════════
   doGet — GET + JSONP，全部操作走此端點
══════════════════════════════════════════ */
function doGet(e) {
  let result;
  try {
    const p      = e.parameter;
    const action = (p.action || '').trim();

    switch (action) {
      /* ── SSO ── */
      case 'ssoLogin':
        result = ssoLogin(p.token);
        break;

      /* ── 查詢 ── */
      case 'getUsers':    result = getUsers();                    break;
      case 'getBookings': result = getBookings();                 break;
      case 'login':       result = login(p.id, p.password);      break;

      /* ── 使用者 CRUD ── */
      case 'addUser':
        result = addUser({ id:p.id, cname:p.cname||'', name:p.name, email:p.email, password:p.password, role:p.role });
        break;
      case 'updateUser':
        result = updateUser({ id:p.id, cname:p.cname||'', name:p.name, email:p.email, role:p.role, password:p.password||'' });
        break;
      case 'deleteUser':
        result = deleteUser({ id: p.id });
        break;

      /* ── 會議 CRUD ── */
      case 'addBooking':
        result = addBooking({
          userId:    p.userId,
          roomIdx:   Number(p.roomIdx),
          date:      p.date,
          startSlot: Number(p.startSlot),
          endSlot:   Number(p.endSlot),
          subject:   p.subject,
          desc:      p.desc      || '',
          attendees: p.attendees || '',
          observers: p.observers || '',
        });
        break;
      case 'updateBooking':
        result = updateBooking({
          id:        Number(p.id),
          roomIdx:   p.roomIdx   !== undefined ? Number(p.roomIdx)   : undefined,
          date:      p.date,
          startSlot: p.startSlot !== undefined ? Number(p.startSlot) : undefined,
          endSlot:   p.endSlot   !== undefined ? Number(p.endSlot)   : undefined,
          subject:   p.subject,
          desc:      p.desc      || '',
          attendees: p.attendees || '',
          observers: p.observers || '',
        });
        break;
      case 'deleteBooking':
        result = deleteBooking({ id: Number(p.id) });
        break;

      /* ── Email ── */
      case 'sendEmail':
        result = sendMeetingEmail({
          email:     p.email,
          userName:  p.userName,
          subject:   p.subject,
          desc:      p.desc      || '',
          roomIdx:   Number(p.roomIdx),
          date:      p.date,
          startTime: p.startTime,
          endTime:   p.endTime,
          isUpdate:  p.isUpdate  || 'false',
          attendees: p.attendees || '',
          observers: p.observers || '',
        });
        break;

      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  const json     = JSON.stringify(result);
  const callback = (e.parameter.callback || '').replace(/[^a-zA-Z0-9_]/g, '');
  const output   = callback ? `${callback}(${json})` : json;
  const mime     = callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(output).setMimeType(mime);
}

/* ══════════════════════════════════════════
   SHEET HELPERS
══════════════════════════════════════════ */
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = createSheet(ss, name);
  return sheet;
}

function createSheet(ss, name) {
  const sheet = ss.insertSheet(name);
  if (name === SHEET_USERS) {
    sheet.appendRow(['id','cname','name','email','password','role','createdAt']);
    sheet.appendRow(['admin','系統管理員','Administrator','admin@company.com',hashPassword('admin123'),'admin',new Date().toISOString()]);
  }
  if (name === SHEET_BOOKINGS) {
    sheet.appendRow(['id','userId','roomIdx','date','startSlot','endSlot','subject','desc','attendees','observers','createdAt']);
    sheet.getRange('D:D').setNumberFormat('@STRING@');
  }
  if (name === SHEET_SSO_TOKENS) {
    sheet.appendRow(['token','userId','expireAt']);
  }
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function toDateStr(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v).slice(0, 10);
}

function hashPassword(pwd) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pwd, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

/* ══════════════════════════════════════════
   SSO — Intranet 單一登入
   Intranet 寫入 SsoTokens sheet（token/userId/expireAt），
   本系統讀取驗證後立即刪除（一次性使用）。
══════════════════════════════════════════ */
function ssoLogin(token) {
  if (!token) return { success: false, message: 'token 不可為空' };
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_SSO_TOKENS);
    if (!sheet) return { success: false, message: 'SsoTokens sheet 不存在，請先執行 setup()' };

    const rows     = sheet.getDataRange().getValues();
    const hdrs     = rows[0];
    const tokCol   = hdrs.indexOf('token');
    const userCol  = hdrs.indexOf('userId');
    const expCol   = hdrs.indexOf('expireAt');

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][tokCol]) !== String(token)) continue;

      const expireAt = new Date(rows[i][expCol]);
      if (new Date() > expireAt) {
        sheet.deleteRow(i + 1);
        return { success: false, message: 'SSO token 已過期，請重新從 Intranet 點入' };
      }

      const userId = String(rows[i][userCol]);
      sheet.deleteRow(i + 1); // 一次性，用完即刪

      const users = sheetToObjects(getSheet(SHEET_USERS));
      const user  = users.find(u => String(u.id) === userId);
      if (!user) return { success: false, message: '找不到對應使用者：' + userId };

      return { success: true, user: {
        id: user.id, cname: user.cname||'', name: user.name, email: user.email, role: user.role
      }};
    }
    return { success: false, message: 'SSO token 無效或已使用' };
  } catch(err) {
    return { success: false, message: 'SSO 驗證失敗：' + err.message };
  }
}

/* ══════════════════════════════════════════
   AUTH
══════════════════════════════════════════ */
function login(id, password) {
  if (!id || !password) return { success: false, message: '請輸入帳號與密碼' };
  const users  = sheetToObjects(getSheet(SHEET_USERS));
  const hashed = hashPassword(password);
  const user   = users.find(u => String(u.id) === String(id) && u.password === hashed);
  if (!user) return { success: false, message: '員工代號或密碼錯誤' };
  return { success: true, user: {
    id: user.id, cname: user.cname||'', name: user.name, email: user.email, role: user.role
  }};
}

/* ══════════════════════════════════════════
   USERS CRUD
══════════════════════════════════════════ */
function getUsers() {
  return sheetToObjects(getSheet(SHEET_USERS))
    .map(u => ({ id: u.id, cname: u.cname||'', name: u.name, email: u.email, role: u.role }));
}

function addUser(data) {
  const sheet = getSheet(SHEET_USERS);
  const users = sheetToObjects(sheet);
  if (users.find(u => String(u.id) === String(data.id)))
    return { success: false, message: '此員工代號已存在' };
  if (!data.password) return { success: false, message: '密碼不可為空' };

  // 依標題順序寫入，cname 欄無論在哪個位置都能正確對應
  const hdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = hdrs.map(h => {
    switch(h) {
      case 'id':        return data.id;
      case 'cname':     return data.cname || '';
      case 'name':      return data.name;
      case 'email':     return data.email;
      case 'password':  return hashPassword(data.password);
      case 'role':      return data.role;
      case 'createdAt': return new Date().toISOString();
      default:          return '';
    }
  });
  sheet.appendRow(newRow);
  return { success: true };
}

function updateUser(data) {
  const sheet = getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  const hdrs  = rows[0];
  const idCol = hdrs.indexOf('id');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(data.id)) {
      const set = (col, val) => {
        const c = hdrs.indexOf(col);
        if (c >= 0 && val !== undefined && val !== null)
          sheet.getRange(i+1, c+1).setValue(val);
      };
      set('cname', data.cname);
      set('name',  data.name);
      set('email', data.email);
      set('role',  data.role);
      if (data.password && data.password !== '')
        sheet.getRange(i+1, hdrs.indexOf('password')+1).setValue(hashPassword(data.password));
      return { success: true };
    }
  }
  return { success: false, message: '使用者不存在' };
}

function deleteUser(data) {
  const sheet = getSheet(SHEET_USERS);
  const rows  = sheet.getDataRange().getValues();
  const idCol = rows[0].indexOf('id');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(data.id)) { sheet.deleteRow(i+1); return { success: true }; }
  }
  return { success: false, message: '使用者不存在' };
}

/* ══════════════════════════════════════════
   BOOKINGS CRUD
══════════════════════════════════════════ */
function getBookings() {
  return sheetToObjects(getSheet(SHEET_BOOKINGS)).map(b => {
    // 安全處理：欄位不存在時為 undefined，需轉成空字串而非 "undefined"
    const safeStr = v => (v === undefined || v === null || v === 'undefined') ? '' : String(v);
    return {
      id:        Number(b.id),
      userId:    safeStr(b.userId),
      roomIdx:   Number(b.roomIdx),
      date:      toDateStr(b.date),
      startSlot: Number(b.startSlot),
      endSlot:   Number(b.endSlot),
      subject:   safeStr(b.subject),
      desc:      safeStr(b.desc),
      attendees: safeStr(b.attendees),  // 欄位不存在時回傳 '' 而非 'undefined'
      observers: safeStr(b.observers),
      createdAt: safeStr(b.createdAt),
    };
  });
}

function addBooking(data) {
  const sheet    = getSheet(SHEET_BOOKINGS);
  const bookings = sheetToObjects(sheet);

  const conflict = bookings.find(b =>
    Number(b.roomIdx) === Number(data.roomIdx) &&
    toDateStr(b.date) === String(data.date) &&
    !(Number(data.endSlot) <= Number(b.startSlot) || Number(data.startSlot) >= Number(b.endSlot))
  );
  if (conflict) return { success: false, conflict: true, message: `時段衝突：${conflict.subject}` };

  const ids   = bookings.map(b => Number(b.id)).filter(n => !isNaN(n));
  const newId = ids.length > 0 ? Math.max(...ids) + 1 : 1;

  // 依標題順序寫入，支援 attendees/observers 欄
  const hdrs  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = hdrs.map(h => {
    switch(h) {
      case 'id':        return newId;
      case 'userId':    return data.userId;
      case 'roomIdx':   return data.roomIdx;
      case 'date':      return data.date;
      case 'startSlot': return data.startSlot;
      case 'endSlot':   return data.endSlot;
      case 'subject':   return data.subject;
      case 'desc':      return data.desc || '';
      case 'attendees': return data.attendees || '';
      case 'observers': return data.observers || '';
      case 'createdAt': return new Date().toISOString();
      default:          return '';
    }
  });
  sheet.appendRow(newRow);
  return { success: true, id: newId };
}

function updateBooking(data) {
  const sheet = getSheet(SHEET_BOOKINGS);
  const rows  = sheet.getDataRange().getValues();
  const hdrs  = rows[0];
  const idCol = hdrs.indexOf('id');

  let targetRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (Number(rows[i][idCol]) === Number(data.id)) { targetRow = i; break; }
  }
  if (targetRow < 0) return { success: false, message: '預定資料不存在' };

  if (data.roomIdx !== undefined && data.date !== undefined &&
      data.startSlot !== undefined && data.endSlot !== undefined) {
    const bookings = sheetToObjects(sheet);
    const conflict = bookings.find(b =>
      Number(b.id)      !== Number(data.id) &&
      Number(b.roomIdx) === Number(data.roomIdx) &&
      toDateStr(b.date) === String(data.date) &&
      !(Number(data.endSlot) <= Number(b.startSlot) || Number(data.startSlot) >= Number(b.endSlot))
    );
    if (conflict) return { success: false, conflict: true, message: `時段衝突：${conflict.subject}` };
  }

  const set = (col, val) => {
    const c = hdrs.indexOf(col);
    if (c >= 0 && val !== undefined && val !== null)
      sheet.getRange(targetRow + 1, c + 1).setValue(val);
  };
  set('roomIdx',   data.roomIdx);
  set('date',      data.date);
  set('startSlot', data.startSlot);
  set('endSlot',   data.endSlot);
  set('subject',   data.subject);
  set('desc',      data.desc);
  set('attendees', data.attendees);
  set('observers', data.observers);
  return { success: true };
}

function deleteBooking(data) {
  const sheet = getSheet(SHEET_BOOKINGS);
  const rows  = sheet.getDataRange().getValues();
  const idCol = rows[0].indexOf('id');
  for (let i = 1; i < rows.length; i++) {
    if (Number(rows[i][idCol]) === Number(data.id)) { sheet.deleteRow(i+1); return { success: true }; }
  }
  return { success: false, message: '預定資料不存在' };
}

/* ══════════════════════════════════════════
   EMAIL + iCalendar (.ics) 行事曆邀請
══════════════════════════════════════════ */
function sendMeetingEmail(data) {
  try {
    const rooms     = ['665 會議室', '663 訓練教室', '663 小會議室'];
    const roomName  = rooms[data.roomIdx] || '';
    const isUpdate  = String(data.isUpdate) === 'true';
    const typeLabel = isUpdate ? '會議變更通知' : '會議預定通知';
    const typeDesc  = isUpdate ? '您的會議已更新為以下內容：' : '您受邀出席以下會議：';

    let attendees = [], observers = [];
    try { attendees = JSON.parse(data.attendees || '[]'); } catch(e) {}
    try { observers = JSON.parse(data.observers || '[]'); } catch(e) {}

    const fmtList = arr => arr.length ? arr.map(p => p.name || p.email).join('、') : '（無）';
    const mailSubject = `[${typeLabel}] ${data.subject} - ${data.date} ${data.startTime}`;

    const textBody =
      `親愛的 ${data.userName}，\n\n${typeDesc}\n\n` +
      `會議主旨：${data.subject}\n地點：${roomName}\n日期：${data.date}\n時間：${data.startTime} - ${data.endTime}\n` +
      (data.desc ? `說明：${data.desc}\n` : '') +
      `與會者：${fmtList(attendees)}\n列席者：${fmtList(observers)}\n\n` +
      `請準時出席。\n\n此信件由會議室管理系統自動發送，請勿回覆。`;

    const headerColor = isUpdate ? '#f5a623' : '#4f7cff';
    const htmlBody =
      `<div style="font-family:Arial,sans-serif;max-width:520px;padding:24px;background:#f4f6fb;border-radius:8px;">` +
      `<h2 style="color:${headerColor};margin-bottom:4px;font-size:18px;">[${typeLabel}]</h2>` +
      `<p style="color:#777;font-size:12px;margin-bottom:16px;">此信件含行事曆邀請附件，請點擊接受以加入行事曆。</p>` +
      `<p style="margin-bottom:16px;">親愛的 <strong>${data.userName}</strong>，${typeDesc}</p>` +
      `<table style="border-collapse:collapse;width:100%;background:#fff;border-radius:6px;overflow:hidden;margin-bottom:16px;">` +
      `<tr style="background:#f0f4ff;"><td style="padding:10px 14px;color:#555;width:90px;">會議主旨</td><td style="padding:10px 14px;font-weight:bold;">${data.subject}</td></tr>` +
      `<tr><td style="padding:10px 14px;color:#555;">地點</td><td style="padding:10px 14px;">${roomName}</td></tr>` +
      `<tr style="background:#f0f4ff;"><td style="padding:10px 14px;color:#555;">日期</td><td style="padding:10px 14px;">${data.date}</td></tr>` +
      `<tr><td style="padding:10px 14px;color:#555;">時間</td><td style="padding:10px 14px;">${data.startTime} - ${data.endTime}</td></tr>` +
      (data.desc ? `<tr style="background:#f0f4ff;"><td style="padding:10px 14px;color:#555;">說明</td><td style="padding:10px 14px;">${data.desc}</td></tr>` : '') +
      `<tr ${data.desc?'':'style="background:#f0f4ff;"'}><td style="padding:10px 14px;color:#555;">與會者</td><td style="padding:10px 14px;">${fmtList(attendees)}</td></tr>` +
      `<tr style="background:#f0f4ff;"><td style="padding:10px 14px;color:#555;">列席者</td><td style="padding:10px 14px;">${fmtList(observers)}</td></tr>` +
      `</table>` +
      `<p style="color:#aaa;font-size:12px;">此信件由會議室管理系統自動發送，請勿回覆。</p></div>`;

    // 建立 iCalendar (.ics) 附件
    const icsContent = buildICS(data, roomName, attendees, observers, isUpdate);
    const icsBlob = Utilities.newBlob(icsContent, 'text/calendar; charset=UTF-8; method=REQUEST', 'meeting.ics');

    // 收集所有收件者（主辦人 + 與會者 + 列席者），去重
    const allEmails = new Set();
    if (data.email) allEmails.add(data.email);
    attendees.forEach(p => { if (p.email) allEmails.add(p.email); });
    observers.forEach(p => { if (p.email) allEmails.add(p.email); });

    allEmails.forEach(addr => {
      GmailApp.sendEmail(addr, mailSubject, textBody, {
        htmlBody: htmlBody, attachments: [icsBlob], name: '會議室管理系統',
      });
    });

    return { success: true, sent: allEmails.size };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function buildICS(data, roomName, attendees, observers, isUpdate) {
  const tz = Session.getScriptTimeZone();
  function toICalDT(dateStr, timeStr) {
    const dt = new Date(`${dateStr}T${timeStr}:00`);
    return Utilities.formatDate(dt, tz, "yyyyMMdd'T'HHmmss");
  }
  const dtStart = toICalDT(data.date, data.startTime);
  const dtEnd   = toICalDT(data.date, data.endTime);
  const dtStamp = Utilities.formatDate(new Date(), tz, "yyyyMMdd'T'HHmmss");
  const uid     = `meeting-${data.date}-${dtStamp}@hansient`;
  const desc    = (data.desc || '').replace(/\n/g, '\\n');

  const attendeeLines = [
    ...attendees.map(p => `ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=${p.name||p.email}:MAILTO:${p.email}`),
    ...observers.map(p => `ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=OPT-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=FALSE;CN=${p.name||p.email}:MAILTO:${p.email}`),
  ].join('\r\n');

  return [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//HansientMeetingSystem//ZH',
    'METHOD:REQUEST',
    'BEGIN:VEVENT',
    `UID:${uid}`,
    `DTSTAMP:${dtStamp}`,
    `DTSTART;TZID=${tz}:${dtStart}`,
    `DTEND;TZID=${tz}:${dtEnd}`,
    `SUMMARY:${data.subject}`,
    `LOCATION:${roomName}`,
    desc ? `DESCRIPTION:${desc}` : '',
    `ORGANIZER;CN=會議室管理系統:MAILTO:${data.email}`,
    attendeeLines,
    `SEQUENCE:${isUpdate ? '1' : '0'}`,
    'STATUS:CONFIRMED',
    'TRANSP:OPAQUE',
    'END:VEVENT',
    'END:VCALENDAR',
  ].filter(Boolean).join('\r\n');
}

/* ══════════════════════════════════════════
   UTILS — date column fix
══════════════════════════════════════════ */
/* ══════════════════════════════════════════
   UTILS — 補齊欄位（手動執行一次）
══════════════════════════════════════════ */
function addMissingColumns() {
  const sheet = getSheet(SHEET_BOOKINGS);
  const hdrs  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let added   = 0;

  // 如果沒有 attendees 欄，在最後加上
  if (!hdrs.includes('attendees')) {
    const col = sheet.getLastColumn() + 1;
    sheet.getRange(1, col).setValue('attendees');
    // 舊資料補空陣列
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, col, sheet.getLastRow()-1, 1).setValue('[]');
    }
    added++;
    Logger.log('✅ 已新增 attendees 欄');
  }
  if (!hdrs.includes('observers')) {
    const col = sheet.getLastColumn() + 1;
    sheet.getRange(1, col).setValue('observers');
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, col, sheet.getLastRow()-1, 1).setValue('[]');
    }
    added++;
    Logger.log('✅ 已新增 observers 欄');
  }

  if (added === 0) Logger.log('ℹ️ attendees 和 observers 欄位已存在，無需新增');
  else Logger.log(`✅ 完成，共新增 ${added} 個欄位`);
}

function fixDateColumn() {
  const sheet   = getSheet(SHEET_BOOKINGS);
  const rows    = sheet.getDataRange().getValues();
  const hdrs    = rows[0];
  const dateCol = hdrs.indexOf('date') + 1;
  if (dateCol < 1) { Logger.log('找不到 date 欄'); return; }
  sheet.getRange(1, dateCol, rows.length, 1).setNumberFormat('@STRING@');
  for (let i = 1; i < rows.length; i++) {
    const val = rows[i][dateCol - 1];
    if (val instanceof Date) {
      sheet.getRange(i+1, dateCol).setValue(
        Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      );
    }
  }
  Logger.log('✅ date 欄格式修正完成，共處理 ' + (rows.length - 1) + ' 筆');
}

/* ══════════════════════════════════════════
   SETUP (run once manually)
══════════════════════════════════════════ */
function setup() {
  getSheet(SHEET_USERS);
  getSheet(SHEET_BOOKINGS);
  getSheet(SHEET_SSO_TOKENS);
  Logger.log('✅ 初始化完成：Users、Bookings、SsoTokens 工作表已建立。');
}
