/**
 * ============================================================
 *  會議室管理系統 - app.js
 *  首頁：公開月曆（不需登入）
 *  點「新增預定」或「登入」後才顯示登入視窗 → 進入完整系統
 * ============================================================
 */

/* ──────────────────────────────────────────
   CONFIG & CONSTANTS
────────────────────────────────────────── */
const CONFIG_KEY  = 'meeting_system_api_url';
let   API_URL     = localStorage.getItem(CONFIG_KEY) || '';

const ROOMS       = ['665 會議室', '663 訓練教室', '663 小會議室'];
const ROOM_COLORS = ['r0', 'r1', 'r2'];
const DAYS_CN     = ['日', '一', '二', '三', '四', '五', '六'];
const SLOT_COUNT  = 22;   // 08:00–19:00（每格 30 分）
const START_HOUR  = 8;

/* ──────────────────────────────────────────
   STATE
────────────────────────────────────────── */
let currentUser        = null;
let allUsers           = [];
let allBookings        = [];
let pubBookings        = [];          // 公開頁用（未登入時）
let currentWeekStart   = getWeekStart(new Date());
let activeRoomFilters  = [0, 1, 2];
let pubRoomFilters     = [0, 1, 2];
let currentMonth       = new Date(); // 公開月曆目前顯示月份
let pendingAction      = null;       // 登入後要執行的動作

/* ──────────────────────────────────────────
   API CLIENT — JSONP
   Apps Script Web App 會 302 redirect 到
   script.googleusercontent.com，瀏覽器的
   CORS 政策會擋住 fetch 的跨域 redirect。
   解法：改用 JSONP（動態插入 <script> tag），
   完全繞過 CORS 限制，Apps Script 原生支援。
────────────────────────────────────────── */
let _jsonpCbIdx = 0;

function apiCall(params) {
  return new Promise((resolve, reject) => {
    if (!API_URL) { reject(new Error('尚未設定 API 網址')); return; }

    const cbName = `_gsCallback${++_jsonpCbIdx}`;
    const timeout = setTimeout(() => {
      delete window[cbName];
      script.remove();
      reject(new Error('連線逾時，請確認網址是否正確'));
    }, 15000);

    window[cbName] = (data) => {
      clearTimeout(timeout);
      delete window[cbName];
      script.remove();
      resolve(data);
    };

    const url = new URL(API_URL);
    url.searchParams.set('callback', cbName);
    Object.entries(params).forEach(([k, v]) => {
      if (v !== undefined && v !== null) url.searchParams.set(k, String(v));
    });

    const script = document.createElement('script');
    script.src = url.toString();
    script.onerror = () => {
      clearTimeout(timeout);
      delete window[cbName];
      script.remove();
      reject(new Error('網路錯誤，無法連線至 Apps Script'));
    };
    document.head.appendChild(script);
  });
}
const apiGet  = p => apiCall(p);
const apiPost = d => apiCall(d);

/* ──────────────────────────────────────────
   HELPERS
────────────────────────────────────────── */
function formatDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function getWeekStart(d) {
  const dd = new Date(d);
  dd.setDate(dd.getDate() - dd.getDay());
  dd.setHours(0,0,0,0);
  return dd;
}
function slotToTime(slot) {
  const m = START_HOUR * 60 + slot * 30;
  return `${String(Math.floor(m/60)).padStart(2,'0')}:${String(m%60).padStart(2,'0')}`;
}
function getNowSlot() {
  const n = new Date();
  return (n.getHours() - START_HOUR) * 2 + (n.getMinutes() >= 30 ? 1 : 0);
}
function getUser(id)    { return allUsers.find(u => u.id === id); }
function initials(name) { return name ? name.slice(0,2) : '?'; }

/* ──────────────────────────────────────────
   SETUP / CONFIG
────────────────────────────────────────── */
function showSetupGuide() {
  document.getElementById('apiUrlInput').value = API_URL;
  document.getElementById('setupModal').classList.remove('hidden');
}
async function saveApiUrl() {
  const url = document.getElementById('apiUrlInput').value.trim();
  if (!url) { showToast('請輸入 API 網址', 'error'); return; }
  API_URL = url;
  localStorage.setItem(CONFIG_KEY, url);
  const btn = document.querySelector('#setupModal .btn-primary');
  btn.innerHTML = '<span class="spinner"></span> 測試連線...'; btn.disabled = true;
  try {
    const r = await apiGet({ action: 'getBookings' });
    if (Array.isArray(r)) {
      showToast('✅ 連線成功！');
      closeModal('setupModal');
      document.getElementById('pubSetupNotice').classList.add('hidden');
      await loadPublicBookings();
    } else {
      showToast('連線失敗：' + (r.error || '未知錯誤'), 'error');
    }
  } catch(e) { showToast('連線失敗：' + e.message, 'error'); }
  finally { btn.innerHTML = '儲存並測試連線'; btn.disabled = false; }
}

/* ──────────────────────────────────────────
   PUBLIC PAGE — 月曆 & 會議室下拉選單
────────────────────────────────────────── */
function onPubRoomSelect() {
  const val = document.getElementById('pubRoomSelect').value;
  pubRoomFilters = val === 'all' ? [0,1,2] : [Number(val)];
  renderMonthCal();
}

function togglePubRoomFilter() {} // 保留空函式以防舊引用
async function loadPublicBookings() {
  if (!API_URL) return;
  document.getElementById('pubLoadingBar').classList.remove('hidden');
  try {
    const r = await apiGet({ action: 'getBookings' });
    pubBookings = Array.isArray(r) ? r.map(b => ({
      ...b, id: Number(b.id), roomIdx: Number(b.roomIdx),
      startSlot: Number(b.startSlot), endSlot: Number(b.endSlot),
    })) : [];

    // Also try to get users for display (may fail if not set up yet)
    try {
      const ur = await apiGet({ action: 'getUsers' });
      if (Array.isArray(ur)) allUsers = ur;
    } catch {}

    renderMonthCal();
  } catch(e) {
    // silently fail on public page
  } finally {
    document.getElementById('pubLoadingBar').classList.add('hidden');
  }
}

function changeMonth(delta) {
  currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + delta, 1);
  renderMonthCal();
}
function goThisMonth() {
  currentMonth = new Date();
  currentMonth.setDate(1);
  renderMonthCal();
}

function renderMonthCal() {
  const year  = currentMonth.getFullYear();
  const month = currentMonth.getMonth();
  document.getElementById('monthLabel').textContent =
    `${year} 年 ${month+1} 月`;

  const today     = formatDate(new Date());
  const firstDay  = new Date(year, month, 1);
  const lastDay   = new Date(year, month+1, 0);
  const startPad  = firstDay.getDay();       // 0=Sun
  const totalCells = Math.ceil((startPad + lastDay.getDate()) / 7) * 7;

  let html = '<div class="mcal-head">';
  DAYS_CN.forEach(d => { html += `<div class="mcal-head-cell">${d}</div>`; });
  html += '</div><div class="mcal-body">';

  for (let i = 0; i < totalCells; i++) {
    const dayNum  = i - startPad + 1;
    const isValid = dayNum >= 1 && dayNum <= lastDay.getDate();
    const d       = isValid ? new Date(year, month, dayNum) : null;
    const ds      = d ? formatDate(d) : '';
    const isToday = ds === today;
    const isPast  = ds && ds < today;

    html += `<div class="mcal-cell${isToday?' today':''}${isPast?' past':''}${!isValid?' empty':''}">`;
    if (isValid) {
      html += `<div class="mcal-day-num${isToday?' today':''}">${dayNum}</div>`;
      // Events for this day
      const dayEvts = pubBookings
        .filter(b => b.date === ds && pubRoomFilters.includes(b.roomIdx))
        .sort((a,b) => a.startSlot - b.startSlot);

      const MAX_SHOW = 3;
      dayEvts.slice(0, MAX_SHOW).forEach(b => {
        const u = allUsers.find(u => u.id === b.userId);
        html += `<div class="mcal-event ${ROOM_COLORS[b.roomIdx]}"
          onclick="showPubEventDetail(${b.id})" title="${b.subject}">
          <span class="mcal-evt-time">${slotToTime(b.startSlot)}</span>
          <span class="mcal-evt-name">${b.subject}</span>
        </div>`;
      });
      if (dayEvts.length > MAX_SHOW) {
        html += `<div class="mcal-more" onclick="showDayDetail('${ds}')">+${dayEvts.length - MAX_SHOW} 則</div>`;
      }
      // Click empty cell to book
      if (!isPast) {
        html += `<div class="mcal-add-btn" onclick="requireLogin('book','${ds}')" title="新增預定">＋</div>`;
      }
    }
    html += '</div>';
  }
  html += '</div>';
  document.getElementById('monthCal').innerHTML = html;
}

function showPubEventDetail(bookingId) {
  const b = pubBookings.find(x => x.id === bookingId);
  if (!b) return;
  const user = allUsers.find(u => u.id === b.userId);
  document.getElementById('pubEventModalBody').innerHTML = `
    <div class="event-detail-row"><span class="event-detail-label">📋 主旨</span><strong>${b.subject}</strong></div>
    <div class="event-detail-row"><span class="event-detail-label">🏢 會議室</span>${ROOMS[b.roomIdx]}</div>
    <div class="event-detail-row"><span class="event-detail-label">📅 日期</span>${b.date}</div>
    <div class="event-detail-row"><span class="event-detail-label">⏰ 時間</span>${slotToTime(b.startSlot)} – ${slotToTime(b.endSlot)}</div>
    <div class="event-detail-row"><span class="event-detail-label">👤 預定人</span>${user?.name || b.userId}</div>
    ${b.desc ? `<div class="event-detail-row"><span class="event-detail-label">📝 說明</span><span style="color:var(--text2)">${b.desc}</span></div>` : ''}
  `;
  document.getElementById('pubEventModal').classList.remove('hidden');
}

/* 顯示某天所有事件的 popup（超過3筆時） */
function showDayDetail(ds) {
  const dayEvts = pubBookings
    .filter(b => b.date === ds && pubRoomFilters.includes(b.roomIdx))
    .sort((a,b) => a.startSlot - b.startSlot);
  const [y,m,d] = ds.split('-');
  document.getElementById('pubEventModalBody').innerHTML =
    `<div style="font-weight:600;margin-bottom:0.75rem;color:var(--text2)">${y} 年 ${Number(m)} 月 ${Number(d)} 日</div>` +
    dayEvts.map(b => {
      const u = allUsers.find(u => u.id === b.userId);
      return `<div class="day-detail-item ${ROOM_COLORS[b.roomIdx]}" onclick="showPubEventDetail(${b.id})">
        <div class="day-detail-time">${slotToTime(b.startSlot)}–${slotToTime(b.endSlot)}</div>
        <div class="day-detail-info">
          <div class="day-detail-subject">${b.subject}</div>
          <div class="day-detail-meta">${ROOMS[b.roomIdx]} · ${u?.name || b.userId}</div>
        </div>
      </div>`;
    }).join('');
  document.getElementById('pubEventModal').classList.remove('hidden');
}

/* ──────────────────────────────────────────
   LOGIN FLOW — 點擊功能才要求登入
────────────────────────────────────────── */
function requireLogin(action, extraArg) {
  if (currentUser) {
    // 已登入，直接執行
    afterLoginAction(action, extraArg);
    return;
  }
  pendingAction = { action, extraArg };
  document.getElementById('loginId').value   = '';
  document.getElementById('loginPwd').value  = '';
  document.getElementById('loginError').classList.add('hidden');
  document.getElementById('loginModal').classList.remove('hidden');
  setTimeout(() => document.getElementById('loginId').focus(), 100);
}

function afterLoginAction(action, extraArg) {
  if (action === 'book') {
    showPage('book');
    if (extraArg) {
      document.getElementById('bookDate').value = extraArg;
      updateAvailableSlots();
    }
  } else if (action === 'login') {
    showPage('calendar');
  }
}

/* ──────────────────────────────────────────
   AUTH
────────────────────────────────────────── */
let authMode = 'login';
function switchAuthTab(mode) {
  authMode = mode;
  document.querySelectorAll('#loginModal .auth-tab').forEach((t,i) => {
    t.classList.toggle('active', (i===0 && mode==='login') || (i===1 && mode==='admin'));
  });
}

async function doLogin() {
  const id  = document.getElementById('loginId').value.trim();
  const pwd = document.getElementById('loginPwd').value;
  const errEl = document.getElementById('loginError');
  const btn   = document.getElementById('loginBtn');
  errEl.classList.add('hidden');

  if (!id || !pwd) { errEl.textContent = '請輸入員工代號和密碼。'; errEl.classList.remove('hidden'); return; }
  if (!API_URL)    { errEl.innerHTML = '尚未設定後端 API，<a href="#" onclick="showSetupGuide()" style="color:var(--accent)">點此設定</a>'; errEl.classList.remove('hidden'); return; }

  btn.innerHTML = '<span class="spinner"></span>'; btn.disabled = true;
  try {
    const result = await apiGet({ action: 'login', id, password: pwd });
    if (!result.success) {
      errEl.textContent = result.message || '登入失敗';
      errEl.classList.remove('hidden');
      return;
    }
    if (authMode === 'admin' && result.user.role !== 'admin') {
      errEl.textContent = '此帳號沒有管理員權限。';
      errEl.classList.remove('hidden');
      return;
    }
    currentUser = result.user;
    closeModal('loginModal');
    updatePubLoginState(); // 更新首頁 header 登入狀態
    // Switch to app screen
    document.getElementById('publicScreen').classList.add('hidden');
    document.getElementById('appScreen').classList.remove('hidden');
    await initApp();
    // Execute pending action
    if (pendingAction) {
      afterLoginAction(pendingAction.action, pendingAction.extraArg);
      pendingAction = null;
    }
  } catch(e) {
    errEl.textContent = '連線失敗：' + e.message;
    errEl.classList.remove('hidden');
  } finally {
    btn.innerHTML = '<span id="loginBtnText">登入</span>'; btn.disabled = false;
  }
}

function doLogout() {
  currentUser = null; allUsers = []; allBookings = [];
  document.getElementById('appScreen').classList.add('hidden');
  document.getElementById('publicScreen').classList.remove('hidden');
  // Reload public calendar with fresh data
  loadPublicBookings();
}

/* ──────────────────────────────────────────
   APP INIT & REFRESH
────────────────────────────────────────── */
async function initApp() {
  updateSidebarUser();
  document.getElementById('adminNav').style.display = currentUser.role === 'admin' ? '' : 'none';
  document.getElementById('bookDate').min   = formatDate(new Date());
  document.getElementById('bookDate').value = formatDate(new Date());
  await refreshData();
}

async function refreshData() {
  const btn    = document.getElementById('refreshBtn');
  const loadEl = document.getElementById('calLoading');
  if (btn)    { btn.innerHTML = '<span class="spinner"></span>'; btn.disabled = true; }
  if (loadEl) loadEl.classList.remove('hidden');
  try {
    const [ur, br] = await Promise.all([
      apiGet({ action: 'getUsers' }),
      apiGet({ action: 'getBookings' }),
    ]);
    allUsers    = Array.isArray(ur) ? ur : [];
    allBookings = Array.isArray(br) ? br.map(b => ({
      ...b, id: Number(b.id), roomIdx: Number(b.roomIdx),
      startSlot: Number(b.startSlot), endSlot: Number(b.endSlot),
    })) : [];
    populateBookingUserSelect();
    renderCalendar();
    const ap = document.querySelector('.page.active');
    if (ap?.id === 'page-mybookings')    renderMyBookings();
    if (ap?.id === 'page-adminBookings') renderAdminBookings();
    if (ap?.id === 'page-adminUsers')    renderAdminUsers();
  } catch(e) { showToast('資料載入失敗：' + e.message, 'error'); }
  finally {
    if (btn)    { btn.innerHTML = '↻'; btn.disabled = false; }
    if (loadEl) loadEl.classList.add('hidden');
  }
}

function updateSidebarUser() {
  const display = currentUser.cname || currentUser.name;
  document.getElementById('userAvatarSidebar').textContent = initials(display);
  document.getElementById('userNameSidebar').textContent   = display;
  document.getElementById('userRoleSidebar').textContent   = currentUser.role === 'admin' ? '系統管理員' : '一般使用者';
}

/* 更新首頁 header 的登入狀態 */
function updatePubLoginState() {
  const loggedIn = !!currentUser;
  document.getElementById('loginTriggerBtn').classList.toggle('hidden', loggedIn);
  document.getElementById('pubUserBadge').classList.toggle('hidden', !loggedIn);
  if (loggedIn) {
    const display = currentUser.cname || currentUser.name;
    document.getElementById('pubUserAvatar').textContent = initials(display);
    document.getElementById('pubUserName').textContent   = display;
  }
}

/* 從首頁進入個人系統 */
function enterApp() {
  if (!currentUser) { requireLogin('login'); return; }
  document.getElementById('publicScreen').classList.add('hidden');
  document.getElementById('appScreen').classList.remove('hidden');
  showPage('calendar');
}

/* 真正登出（清除登入狀態） */
function confirmLogout() {
  if (!confirm('確定要登出？')) return;
  currentUser = null; allUsers = []; allBookings = [];
  document.getElementById('appScreen').classList.add('hidden');
  document.getElementById('publicScreen').classList.remove('hidden');
  updatePubLoginState();
  loadPublicBookings();
}

/* 側邊欄「回首頁」—— 保持登入，僅切換畫面 */
function doLogout() {
  document.getElementById('appScreen').classList.add('hidden');
  document.getElementById('publicScreen').classList.remove('hidden');
  updatePubLoginState();
  loadPublicBookings();
}

function populateBookingUserSelect() {
  const sel = document.getElementById('bookUser');
  const val = sel.value;
  sel.innerHTML = '<option value="">-- 選擇預定人 --</option>';
  allUsers.filter(u => u.role !== 'admin').forEach(u => {
    const display = u.cname ? `${u.cname} (${u.id})` : `${u.name} (${u.id})`;
    sel.innerHTML += `<option value="${u.id}">${display}</option>`;
  });
  if (currentUser.role !== 'admin') { sel.value = currentUser.id; sel.disabled = true; }
  else { sel.disabled = false; sel.value = val || ''; }
}

/* ──────────────────────────────────────────
   NAVIGATION (logged-in app)
────────────────────────────────────────── */
function showPage(pageId) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.getElementById('page-' + pageId).classList.add('active');
  document.querySelectorAll('.nav-item[data-page]').forEach(n => {
    n.classList.toggle('active', n.dataset.page === pageId);
  });
  const titles = {
    calendar:'行事曆總覽', book:'預定會議室', mybookings:'我的會議',
    adminBookings:'所有會議管理', adminUsers:'使用者管理', profile:'個人資料'
  };
  document.getElementById('topbarTitle').textContent = titles[pageId] || pageId;
  closeSidebar();
  if (pageId === 'mybookings')    renderMyBookings();
  if (pageId === 'adminBookings') renderAdminBookings();
  if (pageId === 'adminUsers')    renderAdminUsers();
  if (pageId === 'profile')       loadProfile();
  if (pageId === 'book')          resetBookForm();
}
function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('sidebarOverlay').classList.toggle('open');
}
function closeSidebar() {
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('sidebarOverlay').classList.remove('open');
}

/* ──────────────────────────────────────────
   WEEKLY CALENDAR (app)
────────────────────────────────────────── */
function changeWeek(d) { currentWeekStart.setDate(currentWeekStart.getDate() + d*7); renderCalendar(); }
function goToday()     { currentWeekStart = getWeekStart(new Date()); renderCalendar(); }
function onAppRoomSelect() {
  const val = document.getElementById('appRoomSelect').value;
  activeRoomFilters = val === 'all' ? [0,1,2] : [Number(val)];
  renderCalendar();
}

function toggleRoomFilter() {} // 保留空函式以防舊引用

function renderCalendar() {
  const ws  = new Date(currentWeekStart);
  const we  = new Date(ws); we.setDate(we.getDate()+6);
  const fmt = d => `${d.getMonth()+1}/${d.getDate()}`;
  document.getElementById('weekLabel').textContent = `${ws.getFullYear()} / ${fmt(ws)} – ${fmt(we)}`;

  const today   = formatDate(new Date());
  const nowSlot = getNowSlot();
  const grid    = document.getElementById('calendarGrid');

  let html = '<div class="cal-header"><div class="cal-header-cell"></div>';
  for (let i=0;i<7;i++) {
    const d = new Date(ws); d.setDate(d.getDate()+i);
    const isToday = formatDate(d) === today;
    html += `<div class="cal-header-cell${isToday?' today':''}">
      <div class="day-name">${DAYS_CN[d.getDay()]}</div>
      <div class="day-num">${d.getDate()}</div></div>`;
  }
  html += '</div><div class="cal-body"><div class="cal-time-col">';
  for (let s=0;s<SLOT_COUNT;s++)
    html += `<div class="cal-time-slot">${s%2===0?slotToTime(s):''}</div>`;
  html += '</div><div class="cal-days-grid">';

  for (let i=0;i<7;i++) {
    const d  = new Date(ws); d.setDate(d.getDate()+i);
    const ds = formatDate(d);
    html += '<div class="cal-day-col">';

    for (let s=0;s<SLOT_COUNT;s++) {
      const isPast = ds < today || (ds === today && s < nowSlot);
      html += `<div class="cal-slot${isPast?' past':''}" onclick="calSlotClick('${ds}',${s},${isPast})"></div>`;
    }

    if (ds === today && nowSlot>=0 && nowSlot<SLOT_COUNT) {
      const topPx = nowSlot*40 + (new Date().getMinutes()%30)*(40/30);
      html += `<div class="now-line" style="top:${topPx}px"></div>`;
    }

    // 同一天的事件：計算重疊並分欄顯示
    const dayEvts = allBookings.filter(b => b.date===ds && activeRoomFilters.includes(b.roomIdx));

    // 為每個事件計算「欄位 index」和「總欄數」
    // 策略：對每個事件，找出與它時間重疊的所有事件，分配不同欄
    const columns = assignColumns(dayEvts);

    dayEvts.forEach(b => {
      const u    = getUser(b.userId);
      const col  = columns[b.id];
      const cols = col.total;
      const gap  = 1;                           // 欄間距 1%
      const pct  = (100 - gap * (cols - 1)) / cols;
      const left = col.index * (pct + gap);
      const width = pct;

      html += `<div class="cal-event ${ROOM_COLORS[b.roomIdx]}"
        style="top:${b.startSlot*40}px;height:${(b.endSlot-b.startSlot)*40-3}px;
               left:${left.toFixed(1)}%;width:${width.toFixed(1)}%;"
        onclick="showEventDetail(${b.id});event.stopPropagation();"
        title="${b.subject}&#10;${u?.cname||u?.name||b.userId}">
        <div class="event-name">${b.subject}</div>
        <div class="event-user">${u?.cname||u?.name||b.userId}</div>
      </div>`;
    });
    html += '</div>';
  }
  html += '</div></div>';
  grid.innerHTML = html;
}

/** 計算事件群組與欄位，避免重疊遮蓋
 *  回傳 { [bookingId]: { index, total } }
 */
function assignColumns(events) {
  const result = {};
  // 逐一處理，貪婪分配欄位
  const cols = []; // cols[i] = 此欄目前最後結束的 slot

  events.forEach(ev => {
    // 找第一個可放的欄（結束時間 <= 本事件開始）
    let placed = false;
    for (let c = 0; c < cols.length; c++) {
      if (cols[c] <= ev.startSlot) {
        cols[c] = ev.endSlot;
        result[ev.id] = { index: c, total: 0 }; // total 待補
        placed = true;
        break;
      }
    }
    if (!placed) {
      cols.push(ev.endSlot);
      result[ev.id] = { index: cols.length - 1, total: 0 };
    }
  });

  // 補 total：對每個事件，找出與它有重疊的事件群組最大欄數
  events.forEach(ev => {
    const maxCol = events.reduce((max, other) => {
      const overlap = !(ev.endSlot <= other.startSlot || ev.startSlot >= other.endSlot);
      return overlap ? Math.max(max, result[other.id].index) : max;
    }, 0);
    result[ev.id].total = maxCol + 1;
  });

  return result;
}

function calSlotClick(date, slot, isPast) {
  if (isPast) return;
  showPage('book');
  document.getElementById('bookDate').value = date;
  if (!document.getElementById('bookRoom').value) document.getElementById('bookRoom').value = '0';
  updateAvailableSlots();
  setTimeout(() => {
    const s = document.getElementById('bookStartTime');
    if (s) { s.value = slot; onStartTimeChange(); }
  }, 80);
}

/* ──────────────────────────────────────────
   BOOKING FORM — 起訖時間
────────────────────────────────────────── */
function getTakenRanges(roomIdx, date) {
  return allBookings.filter(b => b.roomIdx===roomIdx && b.date===date);
}
function isSlotTaken(slot, ranges) {
  return ranges.some(b => slot>=b.startSlot && slot<b.endSlot);
}
function buildStartOptions(roomIdx, date) {
  const today   = formatDate(new Date());
  const nowSlot = today===date ? getNowSlot() : -1;
  const ranges  = getTakenRanges(roomIdx, date);
  const sel     = document.getElementById('bookStartTime');
  sel.innerHTML = '<option value="">-- 選擇開始時間 --</option>';
  for (let s=0;s<SLOT_COUNT;s++) {
    if ((s<=nowSlot && date<=today) || isSlotTaken(s,ranges)) continue;
    sel.innerHTML += `<option value="${s}">${slotToTime(s)}</option>`;
  }
}
function buildEndOptions(roomIdx, date, startSlot) {
  const ranges = getTakenRanges(roomIdx, date);
  const sel    = document.getElementById('bookEndTime');
  sel.innerHTML = '<option value="">-- 選擇結束時間 --</option>';
  for (let s=startSlot+1;s<=SLOT_COUNT;s++) {
    if (ranges.some(b => !(s<=b.startSlot || startSlot>=b.endSlot))) break;
    sel.innerHTML += `<option value="${s}">${slotToTime(s)}</option>`;
  }
}
function updateAvailableSlots() {
  const room = document.getElementById('bookRoom').value;
  const date = document.getElementById('bookDate').value;
  const sec  = document.getElementById('slotSection');
  if (!room || !date) { sec.style.display='none'; return; }
  sec.style.display = '';
  const roomIdx = parseInt(room);
  buildStartOptions(roomIdx, date);
  document.getElementById('bookEndTime').innerHTML = '<option value="">-- 先選開始時間 --</option>';
  document.getElementById('timeConflictHint').classList.add('hidden');
  renderBookedList(roomIdx, date);
}
function onStartTimeChange() {
  const room     = document.getElementById('bookRoom').value;
  const date     = document.getElementById('bookDate').value;
  const startVal = document.getElementById('bookStartTime').value;
  if (!room||!date||startVal==='') {
    document.getElementById('bookEndTime').innerHTML='<option value="">-- 先選開始時間 --</option>'; return;
  }
  buildEndOptions(parseInt(room), date, parseInt(startVal));
  document.getElementById('bookEndTime').value='';
  document.getElementById('timeConflictHint').classList.add('hidden');
}
function onEndTimeChange() {
  const room     = document.getElementById('bookRoom').value;
  const date     = document.getElementById('bookDate').value;
  const startVal = document.getElementById('bookStartTime').value;
  const endVal   = document.getElementById('bookEndTime').value;
  const hintEl   = document.getElementById('timeConflictHint');

  if (!room || !date || startVal === '' || endVal === '') {
    hintEl.classList.add('hidden');
    return;
  }

  const roomIdx   = parseInt(room);
  const startSlot = parseInt(startVal);
  const endSlot   = parseInt(endVal);
  const conflicts = getConflicts(roomIdx, date, startSlot, endSlot);

  if (conflicts.length > 0) {
    hintEl.innerHTML = conflicts.map(c => {
      const u = getUser(c.userId);
      const name = u?.cname || u?.name || c.userId;
      return `⚠️ <strong>${ROOMS[roomIdx]}</strong> 在 <strong>${slotToTime(c.startSlot)}–${slotToTime(c.endSlot)}</strong> 已由 <strong>${name}</strong> 預定「${c.subject}」`;
    }).join('<br>');
    hintEl.classList.remove('hidden');
  } else {
    hintEl.classList.add('hidden');
  }
}

/* 取得指定會議室、日期、時段的所有衝突預定 */
function getConflicts(roomIdx, date, startSlot, endSlot) {
  return allBookings.filter(b =>
    b.roomIdx === roomIdx &&
    b.date    === date    &&
    !(endSlot <= b.startSlot || startSlot >= b.endSlot)
  );
}

function renderBookedList(roomIdx, date) {
  const el   = document.getElementById('bookedList');
  const list = allBookings.filter(b=>b.roomIdx===roomIdx && b.date===date)
    .sort((a,b)=>a.startSlot-b.startSlot);
  if (!list.length) { el.classList.add('hidden'); return; }
  el.classList.remove('hidden');
  el.innerHTML = `<div class="booked-list-title">📌 當日已預定時段</div>` +
    list.map(b => {
      const u = getUser(b.userId);
      const name = u?.cname || u?.name || b.userId;
      return `<div class="booked-item">
        <span class="booked-time">${slotToTime(b.startSlot)}–${slotToTime(b.endSlot)}</span>
        <span class="booked-subject">${b.subject}</span>
        <span class="booked-user">${name}</span>
      </div>`;
    }).join('');
}

function resetBookForm() {
  if (currentUser?.role!=='admin') document.getElementById('bookUser').value = currentUser?.id||'';
  document.getElementById('bookRoom').value    = '';
  document.getElementById('bookDate').value    = formatDate(new Date());
  document.getElementById('bookSubject').value = '';
  document.getElementById('bookDesc').value    = '';
  document.getElementById('slotSection').style.display = 'none';
  document.getElementById('bookError').classList.add('hidden');
  document.getElementById('bookSuccess').classList.add('hidden');
}

async function submitBooking() {
  const errEl = document.getElementById('bookError');
  const succEl= document.getElementById('bookSuccess');
  errEl.classList.add('hidden'); succEl.classList.add('hidden');

  const userId   = document.getElementById('bookUser').value;
  const room     = document.getElementById('bookRoom').value;
  const date     = document.getElementById('bookDate').value;
  const startVal = document.getElementById('bookStartTime').value;
  const endVal   = document.getElementById('bookEndTime').value;
  const subject  = document.getElementById('bookSubject').value.trim();
  const desc     = document.getElementById('bookDesc').value.trim();

  if (!userId)        { showAlert(errEl,'請選擇預定人。');   return; }
  if (room==='')      { showAlert(errEl,'請選擇會議室。');   return; }
  if (!date)          { showAlert(errEl,'請選擇日期。');     return; }
  if (startVal==='')  { showAlert(errEl,'請選擇開始時間。'); return; }
  if (endVal==='')    { showAlert(errEl,'請選擇結束時間。'); return; }
  if (!subject)       { showAlert(errEl,'請輸入會議主旨。'); return; }

  const roomIdx   = parseInt(room);
  const startSlot = parseInt(startVal);
  const endSlot   = parseInt(endVal);

  // 前端即時衝突檢查
  const conflicts = getConflicts(roomIdx, date, startSlot, endSlot);
  if (conflicts.length > 0) {
    const lines = conflicts.map(c => {
      const cu   = getUser(c.userId);
      const name = cu?.cname || cu?.name || c.userId;
      return `• ${ROOMS[roomIdx]} ${slotToTime(c.startSlot)}–${slotToTime(c.endSlot)}　已由「${name}」預定：${c.subject}`;
    }).join('<br>');
    showAlert(errEl, `⚠️ 時段衝突，無法預定！<br><br>${lines}<br><br>請重新選擇其他時段或會議室。`);
    return;
  }

  const btn = document.getElementById('bookSubmitBtn');
  btn.innerHTML='<span class="spinner"></span> 預定中...'; btn.disabled=true;
  try {
    const res = await apiPost({ action:'addBooking', userId, roomIdx, date, startSlot, endSlot, subject, desc });
    if (!res.success) {
      if (res.conflict) {
        // 伺服器二次驗證衝突（資料可能在填表期間被他人搶先預定）
        showAlert(errEl, `⚠️ 預定失敗！此時段剛剛已被他人搶先預定。<br>（${res.message}）<br><br>請重新整理後另選時段。`);
        await refreshData(); // 重新載入最新資料
      } else {
        showAlert(errEl, res.message || '預定失敗，請稍後再試。');
      }
      return;
    }
    const user = getUser(userId);
    try {
      await apiPost({ action:'sendEmail', email:user?.email||'', userName:user?.name||userId,
        subject, desc, roomIdx, date, startTime:slotToTime(startSlot), endTime:slotToTime(endSlot) });
      succEl.innerHTML = `✅ 預定成功！<br>📧 通知已發送至 ${user?.email}`;
    } catch { succEl.innerHTML = `✅ 預定成功！（Email 通知發送失敗）`; }
    succEl.classList.remove('hidden');
    await refreshData();
  } catch(e) { showAlert(errEl,'系統錯誤：'+e.message); }
  finally { btn.innerHTML='確認預定'; btn.disabled=false; }
}

/* ──────────────────────────────────────────
   EVENT DETAIL (app)
────────────────────────────────────────── */
function showEventDetail(bookingId) {
  const b = allBookings.find(x => x.id===bookingId); if (!b) return;
  const user    = getUser(b.userId);
  const canEdit = currentUser.role==='admin' || currentUser.id===b.userId;
  document.getElementById('eventModalBody').innerHTML = `
    <div class="event-detail-row"><span class="event-detail-label">📋 主旨</span><strong>${b.subject}</strong></div>
    <div class="event-detail-row"><span class="event-detail-label">🏢 會議室</span>${ROOMS[b.roomIdx]}</div>
    <div class="event-detail-row"><span class="event-detail-label">📅 日期</span>${b.date}</div>
    <div class="event-detail-row"><span class="event-detail-label">⏰ 時間</span>${slotToTime(b.startSlot)} – ${slotToTime(b.endSlot)}</div>
    <div class="event-detail-row"><span class="event-detail-label">👤 預定人</span>${user?.name||b.userId}</div>
    ${b.desc?`<div class="event-detail-row"><span class="event-detail-label">📝 說明</span><span style="color:var(--text2)">${b.desc}</span></div>`:''}
  `;
  document.getElementById('eventModalFooter').innerHTML = canEdit
    ? `<button class="btn btn-secondary btn-sm" onclick="openEditModal(${b.id})">✏️ 編輯</button>
       <button class="btn btn-danger btn-sm" onclick="deleteBooking(${b.id})">🗑️ 刪除</button>
       <button class="btn btn-secondary btn-sm" onclick="closeModal('eventModal')">關閉</button>`
    : `<button class="btn btn-secondary btn-sm" onclick="closeModal('eventModal')">關閉</button>`;
  document.getElementById('eventModal').classList.remove('hidden');
}
function openEditModal(bookingId) {
  const b = allBookings.find(x=>x.id===bookingId); if (!b) return;
  if (currentUser.role!=='admin' && currentUser.id!==b.userId) {
    showToast('您沒有權限編輯此會議。','error'); return;
  }
  closeModal('eventModal');

  // 填入目前的值
  document.getElementById('editBookingId').value = bookingId;
  document.getElementById('editRoom').value      = b.roomIdx;
  document.getElementById('editDate').value      = b.date;
  document.getElementById('editSubject').value   = b.subject;
  document.getElementById('editDesc').value      = b.desc || '';
  document.getElementById('editError').classList.add('hidden');
  document.getElementById('editConflictHint').classList.add('hidden');

  // 設定日期最小值
  document.getElementById('editDate').min = formatDate(new Date());

  // 建立時間選單並預選目前時段
  buildEditStartOptions(b.roomIdx, b.date, b.id);
  setTimeout(() => {
    document.getElementById('editStartTime').value = b.startSlot;
    buildEditEndOptions(b.roomIdx, b.date, b.startSlot, b.id);
    setTimeout(() => {
      document.getElementById('editEndTime').value = b.endSlot;
    }, 50);
  }, 50);

  document.getElementById('editModal').classList.remove('hidden');
}

/* 編輯 modal 的時間選單 helpers
   排除「自己」這筆預定的時段（否則自己的時段會被判定為衝突）*/
function buildEditStartOptions(roomIdx, date, excludeId) {
  const today  = formatDate(new Date());
  const ranges = allBookings.filter(b => b.roomIdx===roomIdx && b.date===date && b.id!==excludeId);
  const nowSlot = today===date ? getNowSlot() : -1;
  const sel = document.getElementById('editStartTime');
  sel.innerHTML = '<option value="">-- 選擇開始時間 --</option>';
  for (let s=0; s<SLOT_COUNT; s++) {
    if (s<=nowSlot && date<=today) continue;
    if (ranges.some(b => s>=b.startSlot && s<b.endSlot)) continue;
    sel.innerHTML += `<option value="${s}">${slotToTime(s)}</option>`;
  }
}

function buildEditEndOptions(roomIdx, date, startSlot, excludeId) {
  const ranges = allBookings.filter(b => b.roomIdx===roomIdx && b.date===date && b.id!==excludeId);
  const sel = document.getElementById('editEndTime');
  sel.innerHTML = '<option value="">-- 選擇結束時間 --</option>';
  for (let s=startSlot+1; s<=SLOT_COUNT; s++) {
    if (ranges.some(b => !(s<=b.startSlot || startSlot>=b.endSlot))) break;
    sel.innerHTML += `<option value="${s}">${slotToTime(s)}</option>`;
  }
}

function onEditRoomOrDateChange() {
  const id      = parseInt(document.getElementById('editBookingId').value);
  const roomIdx = parseInt(document.getElementById('editRoom').value);
  const date    = document.getElementById('editDate').value;
  if (!date) return;
  buildEditStartOptions(roomIdx, date, id);
  document.getElementById('editEndTime').innerHTML = '<option value="">-- 先選開始時間 --</option>';
  document.getElementById('editConflictHint').classList.add('hidden');
}

function onEditStartTimeChange() {
  const id       = parseInt(document.getElementById('editBookingId').value);
  const roomIdx  = parseInt(document.getElementById('editRoom').value);
  const date     = document.getElementById('editDate').value;
  const startVal = document.getElementById('editStartTime').value;
  if (!date || startVal==='') return;
  buildEditEndOptions(roomIdx, date, parseInt(startVal), id);
  document.getElementById('editEndTime').value = '';
  document.getElementById('editConflictHint').classList.add('hidden');
}

function onEditEndTimeChange() {
  const id       = parseInt(document.getElementById('editBookingId').value);
  const roomIdx  = parseInt(document.getElementById('editRoom').value);
  const date     = document.getElementById('editDate').value;
  const startVal = document.getElementById('editStartTime').value;
  const endVal   = document.getElementById('editEndTime').value;
  const hintEl   = document.getElementById('editConflictHint');
  if (!date || startVal==='' || endVal==='') { hintEl.classList.add('hidden'); return; }

  // 衝突檢查（排除自己）
  const conflicts = allBookings.filter(b =>
    b.roomIdx===roomIdx && b.date===date && b.id!==id &&
    !(parseInt(endVal)<=b.startSlot || parseInt(startVal)>=b.endSlot)
  );
  if (conflicts.length > 0) {
    hintEl.innerHTML = conflicts.map(c => {
      const u = getUser(c.userId);
      return `⚠️ <strong>${slotToTime(c.startSlot)}–${slotToTime(c.endSlot)}</strong> 已由 <strong>${u?.cname||u?.name||c.userId}</strong> 預定「${c.subject}」`;
    }).join('<br>');
    hintEl.classList.remove('hidden');
  } else {
    hintEl.classList.add('hidden');
  }
}

async function saveEditBooking() {
  const id       = parseInt(document.getElementById('editBookingId').value);
  const roomIdx  = parseInt(document.getElementById('editRoom').value);
  const date     = document.getElementById('editDate').value;
  const startVal = document.getElementById('editStartTime').value;
  const endVal   = document.getElementById('editEndTime').value;
  const subject  = document.getElementById('editSubject').value.trim();
  const desc     = document.getElementById('editDesc').value.trim();
  const errEl    = document.getElementById('editError');
  errEl.classList.add('hidden');

  if (!date)       { showAlert(errEl,'請選擇日期。');       return; }
  if (startVal==='') { showAlert(errEl,'請選擇開始時間。'); return; }
  if (endVal==='')   { showAlert(errEl,'請選擇結束時間。'); return; }
  if (!subject)    { showAlert(errEl,'請輸入會議主旨。');   return; }

  const startSlot = parseInt(startVal);
  const endSlot   = parseInt(endVal);

  // 衝突檢查（排除自己）
  const conflicts = allBookings.filter(b =>
    b.roomIdx===roomIdx && b.date===date && b.id!==id &&
    !(endSlot<=b.startSlot || startSlot>=b.endSlot)
  );
  if (conflicts.length > 0) {
    const lines = conflicts.map(c => {
      const cu = getUser(c.userId);
      return `• ${slotToTime(c.startSlot)}–${slotToTime(c.endSlot)}　${cu?.cname||cu?.name||c.userId}：${c.subject}`;
    }).join('<br>');
    showAlert(errEl, `⚠️ 時段衝突，無法儲存！<br><br>${lines}<br><br>請重新選擇時段。`);
    return;
  }

  const btn = document.getElementById('editSaveBtn');
  btn.innerHTML='<span class="spinner"></span>'; btn.disabled=true;
  try {
    const res = await apiPost({ action:'updateBooking', id, roomIdx, date, startSlot, endSlot, subject, desc });
    if (!res.success) {
      if (res.conflict) {
        showAlert(errEl, `⚠️ 時段剛被他人搶先預定，請重新選擇。`);
        await refreshData();
      } else {
        showAlert(errEl, res.message);
      }
      return;
    }

    // 發送會議變更通知 email
    const booking = allBookings.find(x => x.id === id);
    const user    = booking ? getUser(booking.userId) : null;
    if (user?.email) {
      try {
        await apiPost({
          action:    'sendEmail',
          email:     user.email,
          userName:  user.cname || user.name || user.id,
          subject:   subject,
          desc:      desc,
          roomIdx:   roomIdx,
          date:      date,
          startTime: slotToTime(startSlot),
          endTime:   slotToTime(endSlot),
          isUpdate:  'true',
        });
      } catch { /* email 失敗不影響主流程 */ }
    }

    closeModal('editModal');
    await refreshData();
    showToast('會議已更新，通知已發送。');
  } catch(e) { showAlert(errEl,'更新失敗：'+e.message); }
  finally { btn.innerHTML='儲存變更'; btn.disabled=false; }
}
async function deleteBooking(bookingId) {
  const b = allBookings.find(x=>x.id===bookingId); if (!b) return;
  if (currentUser.role!=='admin' && currentUser.id!==b.userId) { showToast('您沒有權限刪除此會議。','error'); return; }
  if (!confirm(`確定要刪除「${b.subject}」？`)) return;
  closeModal('eventModal');
  try {
    const res = await apiPost({ action:'deleteBooking', id:bookingId });
    if (!res.success) { showToast(res.message,'error'); return; }
    await refreshData(); showToast('會議已刪除。');
  } catch(e) { showToast('刪除失敗：'+e.message,'error'); }
}

/* ──────────────────────────────────────────
   MY BOOKINGS
────────────────────────────────────────── */
function renderMyBookings() {
  const container = document.getElementById('myBookingsList');
  const today     = formatDate(new Date());
  const myB       = allBookings.filter(b=>b.userId===currentUser.id)
    .sort((a,b)=>a.date.localeCompare(b.date)||a.startSlot-b.startSlot);
  if (!myB.length) {
    container.innerHTML=`<div class="alert alert-info">📭 您目前沒有預定任何會議。
      <br><button class="btn btn-primary" style="margin-top:0.8rem;max-width:200px;" onclick="showPage('book')">立即預定會議室</button></div>`;
    return;
  }
  container.innerHTML = myB.map(b => {
    const isPast = b.date < today;
    return `<div class="booking-card">
      <div class="booking-color-bar" style="background:var(--room${b.roomIdx})"></div>
      <div class="booking-info">
        <div class="booking-subject">${b.subject}</div>
        <div class="booking-meta">
          <span>🏢 ${ROOMS[b.roomIdx]}</span><span>📅 ${b.date}</span>
          <span>⏰ ${slotToTime(b.startSlot)}–${slotToTime(b.endSlot)}</span>
          ${b.desc?`<span style="color:var(--text3)">📝 ${b.desc.slice(0,30)}${b.desc.length>30?'…':''}</span>`:''}
        </div>
      </div>
      <div class="booking-actions">
        ${!isPast
          ?`<button class="btn btn-secondary btn-sm" onclick="openEditModal(${b.id})">編輯</button>
            <button class="btn btn-danger btn-sm" onclick="deleteBooking(${b.id})">刪除</button>`
          :`<span style="font-size:0.75rem;color:var(--text3)">已結束</span>`}
      </div>
    </div>`;
  }).join('');
}

/* ──────────────────────────────────────────
   ADMIN BOOKINGS
────────────────────────────────────────── */
function renderAdminBookings() {
  const tbody  = document.getElementById('adminBookingsTbody');
  const sorted = [...allBookings].sort((a,b)=>a.date.localeCompare(b.date)||a.startSlot-b.startSlot);
  if (!sorted.length) { tbody.innerHTML='<tr><td colspan="6" style="text-align:center;color:var(--text3);padding:2rem;">目前沒有任何會議預定</td></tr>'; return; }
  tbody.innerHTML = sorted.map(b => {
    const user = getUser(b.userId);
    return `<tr>
      <td>${b.date}</td><td class="mono">${slotToTime(b.startSlot)}–${slotToTime(b.endSlot)}</td>
      <td><span class="badge" style="background:rgba(79,124,255,0.12);color:var(--room${b.roomIdx})">${ROOMS[b.roomIdx]}</span></td>
      <td>${b.subject}</td><td>${user?.name||b.userId}</td>
      <td class="table-actions">
        <button class="btn btn-secondary btn-sm" onclick="openEditModal(${b.id})">編輯</button>
        <button class="btn btn-danger btn-sm" onclick="deleteBooking(${b.id})">刪除</button>
      </td></tr>`;
  }).join('');
}

/* ──────────────────────────────────────────
   ADMIN USERS
────────────────────────────────────────── */
function renderAdminUsers() {
  const tbody = document.getElementById('adminUsersTbody');
  tbody.innerHTML = allUsers.map(u=>`<tr>
    <td class="mono">${u.id}</td>
    <td>${u.cname||''}</td>
    <td>${u.name}</td>
    <td><a href="mailto:${u.email}" style="color:var(--accent)">${u.email}</a></td>
    <td><span class="badge ${u.role==='admin'?'badge-admin':'badge-user'}">${u.role==='admin'?'管理員':'使用者'}</span></td>
    <td class="table-actions">
      <button class="btn btn-secondary btn-sm" onclick="openEditUserModal('${u.id}')">編輯</button>
      ${u.id!=='admin'&&u.id!==currentUser.id?`<button class="btn btn-danger btn-sm" onclick="confirmDeleteUser('${u.id}')">刪除</button>`:''}
    </td></tr>`).join('');
}
function openAddUserModal() {
  document.getElementById('userModalTitle').textContent = '👤 新增使用者';
  document.getElementById('userModalMode').value = 'add';
  document.getElementById('userModalOrigId').value = '';
  ['umEmpId','umCname','umName','umEmail','umPassword'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('umEmpId').disabled = false;
  document.getElementById('umRole').value = 'user';
  document.getElementById('userModalError').classList.add('hidden');
  document.getElementById('userModal').classList.remove('hidden');
}
function openEditUserModal(uid) {
  const u = allUsers.find(x=>x.id===uid); if (!u) return;
  document.getElementById('userModalTitle').textContent = '✏️ 編輯使用者';
  document.getElementById('userModalMode').value = 'edit';
  document.getElementById('userModalOrigId').value = uid;
  document.getElementById('umEmpId').value  = u.id;  document.getElementById('umEmpId').disabled = true;
  document.getElementById('umCname').value  = u.cname || '';
  document.getElementById('umName').value   = u.name;
  document.getElementById('umEmail').value  = u.email;
  document.getElementById('umPassword').value = '';
  document.getElementById('umRole').value   = u.role;
  document.getElementById('userModalError').classList.add('hidden');
  document.getElementById('userModal').classList.remove('hidden');
}
async function saveUser() {
  const mode  = document.getElementById('userModalMode').value;
  const id    = document.getElementById('umEmpId').value.trim();
  const cname = document.getElementById('umCname').value.trim();
  const name  = document.getElementById('umName').value.trim();
  const email = document.getElementById('umEmail').value.trim();
  const pwd   = document.getElementById('umPassword').value;
  const role  = document.getElementById('umRole').value;
  const errEl = document.getElementById('userModalError');
  errEl.classList.add('hidden');
  if (!id||!name||!email) { showAlert(errEl,'請填寫所有必填欄位。'); return; }
  if (mode==='add'&&!pwd) { showAlert(errEl,'新增使用者需要設定密碼。'); return; }
  const btn = document.getElementById('saveUserBtn');
  btn.innerHTML='<span class="spinner"></span>'; btn.disabled=true;
  try {
    const action  = mode==='add'?'addUser':'updateUser';
    const payload = mode==='add'
      ? { action, id, cname, name, email, password:pwd, role }
      : { action, id, cname, name, email, role, ...(pwd?{password:pwd}:{}) };
    const res = await apiPost(payload);
    if (!res.success) { showAlert(errEl,res.message); return; }
    closeModal('userModal'); await refreshData();
    showToast(mode==='add'?'使用者已新增。':'使用者資料已更新。');
  } catch(e) { showAlert(errEl,'操作失敗：'+e.message); }
  finally { btn.innerHTML='儲存'; btn.disabled=false; }
}
async function confirmDeleteUser(uid) {
  if (!confirm('確定要刪除此使用者？此操作無法復原。')) return;
  try {
    const res = await apiPost({ action:'deleteUser', id:uid });
    if (!res.success) { showToast(res.message,'error'); return; }
    await refreshData(); showToast('使用者已刪除。');
  } catch(e) { showToast('刪除失敗：'+e.message,'error'); }
}

/* ──────────────────────────────────────────
   PROFILE
────────────────────────────────────────── */
function loadProfile() {
  // 從 allUsers 取最新資料（包含手動在 Sheet 新增的 cname）
  const fresh = allUsers.find(u => u.id === currentUser.id);
  if (fresh) {
    currentUser.cname = fresh.cname || '';
    currentUser.name  = fresh.name;
    currentUser.email = fresh.email;
  }
  const u = currentUser;
  document.getElementById('profileAvatar').textContent = initials(u.cname || u.name);
  document.getElementById('profileId').value    = u.id;
  document.getElementById('profileCname').value = u.cname || '';
  document.getElementById('profileName').value  = u.name;
  document.getElementById('profileEmail').value = u.email;
  ['profileOldPwd','profileNewPwd','profileConfirmPwd'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('profileMsg').className='hidden';
}
async function saveProfile() {
  const cname      = document.getElementById('profileCname').value.trim();
  const name       = document.getElementById('profileName').value.trim();
  const email      = document.getElementById('profileEmail').value.trim();
  const oldPwd     = document.getElementById('profileOldPwd').value;
  const newPwd     = document.getElementById('profileNewPwd').value;
  const confirmPwd = document.getElementById('profileConfirmPwd').value;
  const msgEl      = document.getElementById('profileMsg');
  if (!name||!email) { msgEl.className='alert alert-error'; msgEl.textContent='姓名和 Email 不能為空。'; return; }
  const payload = { action:'updateUser', id:currentUser.id, cname, name, email };
  if (oldPwd||newPwd||confirmPwd) {
    if (!oldPwd) { msgEl.className='alert alert-error'; msgEl.textContent='請輸入現有密碼。'; return; }
    if (!newPwd) { msgEl.className='alert alert-error'; msgEl.textContent='請輸入新密碼。'; return; }
    if (newPwd!==confirmPwd) { msgEl.className='alert alert-error'; msgEl.textContent='新密碼確認不符。'; return; }
    try {
      const lr = await apiGet({ action:'login', id:currentUser.id, password:oldPwd });
      if (!lr.success) { msgEl.className='alert alert-error'; msgEl.textContent='現有密碼錯誤。'; return; }
    } catch(e) { msgEl.className='alert alert-error'; msgEl.textContent='驗證失敗：'+e.message; return; }
    payload.password = newPwd;
  }
  try {
    const res = await apiPost(payload);
    if (!res.success) { msgEl.className='alert alert-error'; msgEl.textContent=res.message; return; }
    currentUser.cname=cname; currentUser.name=name; currentUser.email=email;
    updateSidebarUser();
    document.getElementById('profileAvatar').textContent = initials(cname || name);
    msgEl.className='alert alert-success'; msgEl.textContent='個人資料已成功更新！';
    ['profileOldPwd','profileNewPwd','profileConfirmPwd'].forEach(id=>document.getElementById(id).value='');
    await refreshData();
  } catch(e) { msgEl.className='alert alert-error'; msgEl.textContent='更新失敗：'+e.message; }
}

/* ──────────────────────────────────────────
   MODAL & TOAST
────────────────────────────────────────── */
function closeModal(id) { document.getElementById(id).classList.add('hidden'); }
function showAlert(el, msg) { el.innerHTML=msg; el.classList.remove('hidden'); }
function showToast(msg, type='success') {
  const c = document.getElementById('toastContainer');
  const t = document.createElement('div');
  t.className=`toast ${type}`;
  t.textContent=(type==='success'?'✅ ':'❌ ')+msg;
  c.appendChild(t); setTimeout(()=>t.remove(), 3500);
}

/* ──────────────────────────────────────────
   EVENT LISTENERS
────────────────────────────────────────── */
document.querySelectorAll('.modal-overlay').forEach(o => {
  o.addEventListener('click', function(e) { if (e.target===this) this.classList.add('hidden'); });
});
document.addEventListener('keydown', e => {
  if (e.key==='Enter') {
    if (!document.getElementById('loginModal').classList.contains('hidden')) doLogin();
  }
  if (e.key==='Escape') {
    document.querySelectorAll('.modal-overlay:not(.hidden)').forEach(m=>m.classList.add('hidden'));
  }
});

/* ──────────────────────────────────────────
   INIT — 頁面載入
────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', async () => {
  currentMonth.setDate(1);
  updatePubLoginState();

  // 用 addEventListener 綁定選單，比 HTML onchange attribute 更可靠
  const pubSel = document.getElementById('pubRoomSelect');
  if (pubSel) pubSel.addEventListener('change', onPubRoomSelect);

  const appSel = document.getElementById('appRoomSelect');
  if (appSel) appSel.addEventListener('change', onAppRoomSelect);

  if (!API_URL) {
    document.getElementById('pubSetupNotice').classList.remove('hidden');
    renderMonthCal();
  } else {
    document.getElementById('pubSetupNotice').classList.add('hidden');
    await loadPublicBookings();
  }
});

// 每 2 分鐘自動更新公開月曆
setInterval(() => {
  if (!currentUser && API_URL) loadPublicBookings();
  if (currentUser && !document.getElementById('appScreen').classList.contains('hidden')) refreshData();
}, 120000);
