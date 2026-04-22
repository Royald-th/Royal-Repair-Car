/* ============================================================
   CONFIG — แก้ค่าเหล่านี้ให้ตรงกับของคุณ
   ============================================================ */
const CONFIG = {
  LIFF_ID:    '2009307051-5UV6LT8a',   // LIFF ID จาก LINE Developers
  GAS_URL:    'https://script.google.com/macros/s/AKfycbwEj8iUtGMJdEnuqHQZ1i2lLJnwIOZ-zKjY6m52ObqUQZj3SyG6Yhmwh2js4HMkHnXRuQ/exec',   // URL จาก Deploy > Web App ใน GAS
  API_SECRET: 'UJ&8m@+<KL#S9z1$gB!6', // ต้องตรงกับ API_SECRET ใน Code.gs
};

/* ============================================================
   LAZY SCRIPT LOADER
   ============================================================ */
const _scriptCache = {};
function loadScript(url) {
  if (_scriptCache[url]) return _scriptCache[url];
  _scriptCache[url] = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = url; s.async = true;
    s.onload = resolve; s.onerror = reject;
    document.head.appendChild(s);
  });
  return _scriptCache[url];
}

const LIBS = {
  chartjs:     'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js',
  html2canvas: 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
  jspdf:       'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
  qrcodejs:    'https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js',
};

async function ensureChartJs()     { await loadScript(LIBS.chartjs); }
async function ensurePdfLibs()     { await Promise.all([loadScript(LIBS.html2canvas), loadScript(LIBS.jspdf)]); }
async function ensureQRCode()      { await loadScript(LIBS.qrcodejs); }

/* ============================================================
   STATE
   ============================================================ */
let currentUser   = null;   // { lineUid, name, avatar, role, dept }
let allJobs       = [];
let vehicles      = [];
let users         = [];
let filterStatus  = '';
let editingJobId  = null;
let submitting    = false;
let currentPage   = 1;
const PAGE_SIZE   = 80;       // rows ต่อหน้า
let _totalJobs    = 0;        // จำนวนงานทั้งหมด (จาก server)
let _loadingMore  = false;    // กัน double-fetch

/* ============================================================
   LIFF INIT
   ============================================================ */
async function initLiff() {
  showLoading(true);
  try {
    // ===== เช็ค sessionStorage ก่อน — ถ้ามีข้อมูลเก่าให้ใช้ทันที =====
    const cached = sessionStorage.getItem('_repairUser');
    if (cached) {
      try {
        currentUser = JSON.parse(cached);
        bootApp();
        // refresh user ล่าสุดจาก Sheet แบบ silent (background)
        gasCall('getUser', { lineUid: currentUser.lineUid }).then(res => {
          if (res.status === 'ok' && res.user) {
            currentUser = { ...res.user, avatar: currentUser.avatar };
            sessionStorage.setItem('_repairUser', JSON.stringify(currentUser));
          }
        });
        showLoading(false);
        return; // ออกจากฟังก์ชันได้เลย ไม่ต้อง init LIFF ใหม่
      } catch(e) {
        sessionStorage.removeItem('_repairUser'); // ถ้า parse ไม่ได้ให้ล้างทิ้ง
      }
    }

    // ===== ไม่มี cache — init LIFF ปกติ =====
    await liff.init({ liffId: CONFIG.LIFF_ID });
    if (!liff.isLoggedIn()) { liff.login(); return; }

    const profile    = await liff.getProfile();
    const lineUid    = profile.userId;
    const lineAvatar = profile.pictureUrl || AVATAR_DEFAULT;
    const lineName   = profile.displayName || 'ผู้ใช้งาน';

    const res = await gasCall('getUser', { lineUid });

    if (res.status === 'ok' && res.user) {
      currentUser = { ...res.user, avatar: lineAvatar };
      gasCall('updateUserAvatar', { lineUid, avatar: lineAvatar, name: lineName });
    } else {
      // ยังไม่มีในระบบ — สมัครอัตโนมัติ
      const regRes = await gasCall('registerUser', {
        lineUid, name: lineName, avatar: lineAvatar, role: 'user', dept: '',
      });
      if (regRes.status !== 'ok') {
        showLoading(false);
        showToast('ไม่สามารถลงทะเบียนได้: ' + (regRes.message || ''), 'error');
        return;
      }
      currentUser = { lineUid, name: lineName, avatar: lineAvatar, role: 'user', dept: '' };
      showToast('ลงทะเบียนเข้าสู่ระบบสำเร็จ ยินดีต้อนรับ! 🎉');
    }

    // บันทึก session
    sessionStorage.setItem('_repairUser', JSON.stringify(currentUser));
    bootApp();

  } catch (e) {
    showLoading(false);
    console.error(e);
    showToast('เกิดข้อผิดพลาด: ' + e.message, 'error');
  }
}



/* ============================================================
   BOOT APP
   ============================================================ */
// Keep GAS warm — ping ทุก 4 นาที ป้องกัน cold start
let _warmTimer = null;
function startKeepWarm() {
  if (_warmTimer) return;
  _warmTimer = setInterval(() => {
    gasCall('ping', {}).catch(() => {});
  }, 4 * 60 * 1000); // 4 นาที
}

function bootApp() {
  document.getElementById('login-screen').style.display = 'none';
  document.getElementById('main-app').style.display = 'block';
  startKeepWarm();

  const role = currentUser.role || 'user';

  // Set nav
  document.getElementById('nav-username').textContent = currentUser.name;
  document.getElementById('nav-role').textContent = {admin:'Admin', supervisor:'หัวหน้า', accountant:'ฝ่ายบัญชี', manager:'ผู้บริหาร', user:'User'}[role] || 'User';
  document.getElementById('nav-avatar').src = currentUser.avatar;
  document.getElementById('dash-avatar').src = currentUser.avatar;
  document.getElementById('form-avatar').src = currentUser.avatar;
  document.getElementById('dash-name').textContent = currentUser.name;
  document.getElementById('dash-dept').textContent = currentUser.dept || '-';
  document.getElementById('form-username').textContent = currentUser.name;
  document.getElementById('form-lineuid').textContent = currentUser.dept || '-';

  // Sidebar menus
  if (['admin','supervisor'].includes(role)) {
    document.getElementById('admin-menu').style.display = 'block';
  }
  if (['admin','manager'].includes(role)) {
    document.getElementById('manager-menu').style.display = 'block';
  }
  if (['admin','accountant'].includes(role)) {
    document.getElementById('accountant-menu').style.display = 'block';
  }

  // Admin-only extras
  if (role === 'admin') {
    document.getElementById('admin-note-section').classList.remove('d-none');
    const exportBtn = document.getElementById('btn-export-csv');
    if (exportBtn) exportBtn.style.display = 'inline-flex';
    const archiveBtn = document.getElementById('btn-archive');
    if (archiveBtn) archiveBtn.style.display = 'inline-flex';
  }

  // Manager/Accountant — ซ่อน FAB และ sidebar เมนูหลักทั้งก้อน
  if (['manager','accountant'].includes(role)) {
    const fab = document.getElementById('fab-repair');
    if (fab) { fab.style.display = 'none'; fab.style.setProperty('display','none','important'); }
    const mainMenu = document.getElementById('main-menu');
    if (mainMenu) mainMenu.style.display = 'none';
  }

  // navigate ทันทีก่อน
  if (role === 'manager') {
    navigate('manager-jobs');
  } else if (role === 'accountant') {
    navigate('accountant-jobs');
  } else {
    navigate('dashboard');
  }

  // โหลดข้อมูล — ถ้ามี cache แสดงทันที ถ้าไม่มีค่อย fetch
  loadData();
}

/* ============================================================
   DATA LOADING
   ============================================================ */
async function manualSync() {
  const btn = document.getElementById('btn-sync');
  if (!btn || btn.classList.contains('syncing')) return; // ป้องกันกด 2 ครั้ง
  btn.classList.add('syncing');
  await loadData(true);
  btn.classList.remove('syncing');
  showToast('อัปเดตข้อมูลแล้ว ✅');
}

let _loadingInProgress = false;

/* --- Session Cache (TTL 5 นาที) --- */
const _CACHE_KEY = '_repairData';
const _CACHE_TTL = 5 * 60 * 1000;

function _saveCache() {
  try { sessionStorage.setItem(_CACHE_KEY, JSON.stringify({ ts: Date.now(), allJobs, vehicles, users, totalJobs: _totalJobs, page: currentPage })); } catch(e) {}
}
function _loadCache() {
  try {
    const d = JSON.parse(sessionStorage.getItem(_CACHE_KEY) || 'null');
    if (!d || Date.now() - d.ts > _CACHE_TTL) return false;
    allJobs = d.allJobs || []; vehicles = d.vehicles || []; users = d.users || [];
    _totalJobs = d.totalJobs || allJobs.length; currentPage = d.page || 1;
    return true;
  } catch(e) { return false; }
}

function _renderAll() {
  const role = currentUser?.role;
  const isAdmin = ['admin','supervisor','accountant'].includes(role);
  const isManager = role === 'manager';
  const isAccountant = role === 'accountant';

  populateVehicleSelect();
  populateVehicleHistorySelect();
  renderStats();
  renderRecentJobs();
  renderRepairTable();
  if (['admin','supervisor'].includes(role)) { renderUserList(); renderVehicleList(); }
  // อัปเดต badge รออนุมัติ
  if (isManager) {
    const count = allJobs.filter(j => j.status === 'รอการอนุมัติ').length;
    const b = document.getElementById('badge-approval');
    if (b) { b.textContent = count; b.style.display = count ? 'inline' : 'none'; }
    // refresh หน้า manager ถ้ากำลังอยู่หน้านั้น
    if (document.getElementById('page-manager-jobs')?.classList.contains('active')) renderManagerJobs();
  }
}

async function _fetchFresh(silent = false, appendMode = false) {
  try {
    const isAdmin = ['admin','supervisor','accountant','manager'].includes(currentUser?.role);
    const jobParams = {
      lineUid:  currentUser.lineUid,
      isAdmin,
      page:     currentPage,
      pageSize: PAGE_SIZE,
    };
    const [jobsRes, vehiclesRes, usersRes] = await Promise.all([
      gasCall('getJobs', jobParams),
      gasCall('getVehicles', {}),
      isAdmin ? gasCall('getUsers', { adminUid: currentUser.lineUid }) : Promise.resolve({ users: [] })
    ]);
    if (jobsRes.status === 'error') { if (!silent) showToast('getJobs: ' + jobsRes.message, 'error'); return; }
    if (appendMode) {
      allJobs = allJobs.concat(jobsRes.jobs || []);
    } else {
      allJobs = jobsRes.jobs || [];
    }
    _totalJobs = jobsRes.total  || allJobs.length;
    vehicles   = vehiclesRes.vehicles || [];
    users      = usersRes.users || [];
    _saveCache();
    _renderAll();
    _updateLoadMoreBtn();
  } catch(e) {
    console.error('loadData error:', e);
    if (!silent) showToast('โหลดข้อมูลไม่สำเร็จ: ' + e.message, 'error');
  }
}

async function loadData(silent = false) {
  if (_loadingInProgress) return;
  _loadingInProgress = true;
  currentPage = 1; // reset ไปหน้าแรกเสมอ
  // มี cache — แสดงผลทันที แล้ว refresh background
  if (_loadCache()) {
    _renderAll();
    _fetchFresh(true).finally(() => { _loadingInProgress = false; });
    return;
  }
  if (!silent) showLoading(true);
  await _fetchFresh(silent);
  _loadingInProgress = false;
  if (!silent) showLoading(false);
}

/* โหลดหน้าถัดไป */
async function loadMore() {
  if (_loadingMore) return;
  if (allJobs.length >= _totalJobs) return; // โหลดครบแล้ว
  _loadingMore = true;
  currentPage++;
  const btn = document.getElementById('btn-load-more');
  if (btn) { btn.disabled = true; btn.textContent = 'กำลังโหลด...'; }
  await _fetchFresh(true, true); // appendMode = true
  _loadingMore = false;
}

/* อัปเดตปุ่ม Load More */
function _updateLoadMoreBtn() {
  const btn = document.getElementById('btn-load-more');
  if (!btn) return;
  const remaining = _totalJobs - allJobs.length;
  if (remaining > 0) {
    btn.style.display = 'flex';
    btn.disabled = false;
    btn.textContent = `โหลดเพิ่ม (${remaining} รายการ)`;
  } else {
    btn.style.display = 'none';
  }
}

// ฟังก์ชัน loadData แบบรอให้เสร็จแน่นอน (ใช้ตอน bootApp)
async function loadDataAndWait() {
  _loadingInProgress = false; // reset ก่อน
  if (_loadCache()) {
    _renderAll();
    await _fetchFresh(true);
    return;
  }
  showLoading(true);
  await _fetchFresh(false);
  showLoading(false);
}

/* ============================================================
   RENDER
   ============================================================ */
function renderStats() {
  const _isPriv = ['admin','supervisor','accountant','manager'].includes(currentUser.role);
  const mine = _isPriv ? allJobs : allJobs.filter(j => j.lineUid === currentUser.lineUid);
  const count = s => mine.filter(j => j.status === s).length;
  animateNum('stat-waiting',   count('รอดำเนินการ'));
  animateNum('stat-repairing', count('กำลังซ่อม'));
  animateNum('stat-done',      count('เสร็จสิ้น'));
  animateNum('stat-total',     mine.length);
  if (['admin','supervisor'].includes(currentUser.role)) {
    animateNum('stat-rejected', count('ไม่อนุมัติ'));
    animateNum('stat-revise',   count('ส่งกลับแก้ไข'));
    document.querySelectorAll('.admin-only-stat').forEach(el => el.style.display = '');
  }


  // badge รอดำเนินการ
  const pending = allJobs.filter(j => j.status === 'รอดำเนินการ').length;
  const el = document.getElementById('badge-pending');
  el.textContent = pending;
  el.style.display = pending ? 'inline' : 'none';

  // badge รออนุมัติ + refresh manager-jobs (สำหรับ admin ที่มีทั้ง 2 menu)
  if (['admin','manager'].includes(currentUser?.role)) {
    const approvalCount = allJobs.filter(j => j.status === 'รอการอนุมัติ').length;
    const badge = document.getElementById('badge-approval');
    if (badge) { badge.textContent = approvalCount; badge.style.display = approvalCount ? 'inline' : 'none'; }
    if (document.getElementById('page-manager-jobs')?.classList.contains('active')) renderManagerJobs();
  }
  if (['admin','accountant'].includes(currentUser?.role)) {
    const reviewCount = allJobs.filter(j => j.status === 'รอตรวจสอบ').length;
    const badge = document.getElementById('badge-review');
    if (badge) { badge.textContent = reviewCount; badge.style.display = reviewCount ? 'inline' : 'none'; }
    if (document.getElementById('page-accountant-jobs')?.classList.contains('active')) renderAccountantJobs();
  }
}

function renderRecentJobs() {
  const mine = ['admin','supervisor','accountant','manager'].includes(currentUser.role) ? allJobs : allJobs.filter(j => j.lineUid === currentUser.lineUid);
  const recent = mine.slice(0,5);
  const el = document.getElementById('recent-jobs-list');
  if (!recent.length) { el.innerHTML = '<div class="empty-state"><span class="material-icons">inbox</span>ยังไม่มีรายการ</div>'; return; }
  el.innerHTML = recent.map(jobCard).join('');
}

function renderRepairTable() {
  try {
  let jobs = ['admin','supervisor','accountant','manager'].includes(currentUser.role) ? allJobs : allJobs.filter(j => j.lineUid === currentUser.lineUid);

  // filter status
  if (filterStatus) jobs = jobs.filter(j => j.status === filterStatus);

  // search
  const q = document.getElementById('search-input').value.toLowerCase();
  if (q) jobs = jobs.filter(j => ((j.jobId||'')+(j.plate||'')+(j.userName||'')+(j.detail||'')).toLowerCase().includes(q));

  // date filter
  const df = document.getElementById('filter-date').value;
  if (df) {
    const now = new Date(); const y=now.getFullYear(),m=now.getMonth(),d=now.getDate();
    jobs = jobs.filter(j => {
      const t = new Date(j.createdAt);
      if (df==='today') return t.getFullYear()===y && t.getMonth()===m && t.getDate()===d;
      if (df==='week')  { const s=new Date(now); s.setDate(d-7); return t>=s; }
      if (df==='month') return t.getFullYear()===y && t.getMonth()===m;
      return true;
    });
  }

  // ===== SORT =====
  const repairSort = document.getElementById('repair-sort')?.value || 'newest';
  jobs = [...jobs].sort((a, b) => {
    const ta = new Date(a.createdAt).getTime() || 0;
    const tb = new Date(b.createdAt).getTime() || 0;
    return repairSort === 'newest' ? tb - ta : ta - tb;
  });

  // ===== PAGINATION =====
  const totalPages = Math.max(1, Math.ceil(jobs.length / PAGE_SIZE));
  if (currentPage > totalPages) currentPage = totalPages;
  const paged = jobs.slice((currentPage-1)*PAGE_SIZE, currentPage*PAGE_SIZE);

  // table
  const tbody = document.getElementById('repair-tbody');
  if (!paged.length) { tbody.innerHTML = '<tr><td colspan="8" class="text-center py-4 text-muted">ไม่มีรายการที่ตรงกัน</td></tr>'; }
  else {
    tbody.innerHTML = paged.map(j => `
      <tr>
        <td><span style="font-weight:600;font-size:.82rem;color:var(--primary-dark);">${j.jobId}</span></td>
        <td>${formatDate(j.createdAt)}</td>
        <td>${j.userName || '-'}</td>
        <td><strong>${j.plate}</strong></td>
        <td>${Number(j.mileage).toLocaleString()} กม.</td>
        <td>${j.location}</td>
        <td><span class="badge-status ${statusClass(j.status)}">${j.status}</span></td>
        <td>
          <button class="btn-sm-icon" style="background:var(--info-light);color:var(--info);" title="รายละเอียด" onclick="openDetail('${j.jobId}')"><span class="material-icons">visibility</span></button>
          ${['admin','supervisor'].includes(currentUser.role) ? `
          ${(currentUser.role==='admin' ? j.status!=='ไม่อนุมัติ' : (currentUser.role==='accountant' ? false : ['รอดำเนินการ','ส่งกลับแก้ไข'].includes(j.status))) ? `<button class="btn-sm-icon ms-1" style="background:var(--accent-light);color:var(--accent);" title="แก้ไข" onclick="openEditRepair('${j.jobId}')"><span class="material-icons">edit</span></button>` : ''}
          ${(j.status !== 'เสร็จสิ้น' && (currentUser.role !== 'accountant' || j.status === 'รอตรวจสอบ')) ? `<button class="btn-sm-icon ms-1" style="background:var(--primary-xlight);color:var(--primary);" title="เปลี่ยนสถานะ" onclick="openStatusModal('${j.jobId}')"><span class="material-icons">update</span></button>` : ''}
          <button class="btn-sm-icon ms-1" style="background:var(--danger-light);color:var(--danger);" title="ลบรายการ" onclick="confirmDeleteJob('${j.jobId}','${j.plate}')"><span class="material-icons">delete</span></button>
          ` : currentUser.role === 'manager' ? `
          ${j.status === 'รอการอนุมัติ' ? `
          <button class="btn-sm-icon ms-1" style="background:#E8F5E9;color:#2E7D32;" title="อนุมัติ" onclick="quickDecision('${j.jobId}','อนุมัติ')"><span class="material-icons">check</span></button>
          <button class="btn-sm-icon ms-1" style="background:#FFEBEE;color:#C62828;" title="ไม่อนุมัติ" onclick="quickDecision('${j.jobId}','ไม่อนุมัติ')"><span class="material-icons">close</span></button>
          <button class="btn-sm-icon ms-1" style="background:#FFF8E1;color:#F57F17;" title="ส่งกลับแก้ไข" onclick="quickDecision('${j.jobId}','ส่งกลับแก้ไข')"><span class="material-icons">undo</span></button>` : ''}
          ` : ''}
        </td>
      </tr>`).join('');
  }

  // render pagination bar
  renderPagination(totalPages, jobs.length);

  // mobile cards
  const mc = document.getElementById('repair-cards-mobile');
  if (!paged.length) { mc.innerHTML = '<div class="empty-state"><span class="material-icons">inbox</span>ไม่มีรายการ</div>'; }
  else mc.innerHTML = paged.map(jobCard).join('');
  } catch(e) {
    console.error('renderRepairTable error:', e);
    const tbody2 = document.getElementById('repair-tbody');
    if (tbody2) tbody2.innerHTML = '<tr><td colspan="8" class="text-center py-3 text-danger">⚠️ render error: ' + e.message + '</td></tr>';
  }
}

function renderPagination(totalPages, totalItems) {
  let el = document.getElementById('pagination-bar');
  if (!el) {
    el = document.createElement('div');
    el.id = 'pagination-bar';
    el.style.cssText = 'display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:.5rem;margin-top:.75rem;font-size:.85rem;';
    const wrapper = document.querySelector('#page-repair-list .data-table-wrapper');
    if (wrapper) wrapper.after(el);
  }
  if (totalPages <= 1) { el.innerHTML = `<span style="color:var(--gray-500);">ทั้งหมด ${totalItems} รายการ</span>`; return; }

  let pages = '';
  for (let i=1; i<=totalPages; i++) {
    const active = i===currentPage;
    pages += `<button onclick="goPage(${i})" style="min-width:32px;height:32px;border-radius:6px;border:1.5px solid ${active?'var(--primary)':'var(--gray-300)'};background:${active?'var(--primary)':'#fff'};color:${active?'#fff':'var(--gray-700)'};font-weight:600;cursor:pointer;font-size:.82rem;">${i}</button>`;
  }
  el.innerHTML = `
    <span style="color:var(--gray-500);">ทั้งหมด ${totalItems} รายการ | หน้า ${currentPage}/${totalPages}</span>
    <div style="display:flex;gap:.3rem;align-items:center;">
      <button onclick="goPage(${currentPage-1})" ${currentPage===1?'disabled':''} style="min-width:32px;height:32px;border-radius:6px;border:1.5px solid var(--gray-300);background:#fff;cursor:pointer;">◀</button>
      ${pages}
      <button onclick="goPage(${currentPage+1})" ${currentPage===totalPages?'disabled':''} style="min-width:32px;height:32px;border-radius:6px;border:1.5px solid var(--gray-300);background:#fff;cursor:pointer;">▶</button>
    </div>`;
}

function goPage(p) {
  const totalPages = Math.max(1, Math.ceil(allJobs.length / PAGE_SIZE));
  if (p < 1 || p > totalPages) return;
  currentPage = p;
  renderRepairTable();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function jobCard(j) {
  const role = currentUser?.role || 'user';
  const isSuperAdmin = ['admin','supervisor'].includes(role);
  const isManager    = role === 'manager';

  let adminBtns = '';
  if (isSuperAdmin) {
    adminBtns = `
    <div class="d-flex gap-1 mt-2 justify-content-end" onclick="event.stopPropagation()">
      <button class="btn-sm-icon" style="background:var(--info-light);color:var(--info);" onclick="openDetail('${j.jobId}')" title="รายละเอียด"><span class="material-icons">visibility</span></button>
      ${(currentUser.role==='admin' ? j.status!=='ไม่อนุมัติ' : (currentUser.role==='accountant' ? false : ['รอดำเนินการ','ส่งกลับแก้ไข'].includes(j.status))) ? `<button class="btn-sm-icon" style="background:var(--accent-light);color:var(--accent);" onclick="openEditRepair('${j.jobId}')" title="แก้ไข"><span class="material-icons">edit</span></button>` : ''}
      ${(j.status !== 'เสร็จสิ้น' && (currentUser.role !== 'accountant' || j.status === 'รอตรวจสอบ')) ? `<button class="btn-sm-icon" style="background:var(--primary-xlight);color:var(--primary);" onclick="openStatusModal('${j.jobId}')" title="เปลี่ยนสถานะ"><span class="material-icons">update</span></button>` : ''}
      <button class="btn-sm-icon" style="background:var(--danger-light);color:var(--danger);" onclick="confirmDeleteJob('${j.jobId}','${j.plate}')" title="ลบ"><span class="material-icons">delete</span></button>
    </div>`;
  } else if (isManager) {
    adminBtns = `
    <div class="d-flex gap-1 mt-2 justify-content-end" onclick="event.stopPropagation()">
      <button class="btn-sm-icon" style="background:var(--info-light);color:var(--info);" onclick="openDetail('${j.jobId}')" title="รายละเอียด"><span class="material-icons">visibility</span></button>
      ${j.status === 'รอการอนุมัติ' ? `
      <button class="btn-sm-icon" style="background:#E8F5E9;color:#2E7D32;" onclick="quickDecision('${j.jobId}','อนุมัติ')" title="อนุมัติ"><span class="material-icons">check</span></button>
      <button class="btn-sm-icon" style="background:#FFEBEE;color:#C62828;" onclick="quickDecision('${j.jobId}','ไม่อนุมัติ')" title="ไม่อนุมัติ"><span class="material-icons">close</span></button>
      <button class="btn-sm-icon" style="background:#FFF8E1;color:#F57F17;" onclick="quickDecision('${j.jobId}','ส่งกลับแก้ไข')" title="ส่งกลับแก้ไข"><span class="material-icons">undo</span></button>` : ''}
    </div>`;
  }

  return `<div class="repair-card" onclick="openDetail('${j.jobId}')">
    <div class="d-flex align-items-center justify-content-between mb-1">
      <span class="job-id">${j.jobId}</span>
      <span class="badge-status ${statusClass(j.status)}">${j.status}</span>
    </div>
    <div class="plate mb-1">${j.plate}</div>
    <div style="font-size:.8rem;color:var(--gray-500);display:flex;gap:.75rem;flex-wrap:wrap;margin-bottom:.2rem;">
      <span>👤 ${j.userName || '-'}</span>
      <span>📏 ${Number(j.mileage||0).toLocaleString()} กม.</span>
      <span>📍 ${j.location || '-'}</span>
    </div>
    <div style="font-size:.82rem;color:var(--gray-500);">🕐 ${formatDate(j.createdAt)}</div>
    <div style="font-size:.83rem;color:var(--gray-700);margin-top:.25rem;">${(j.detail||'').substring(0,60)}${(j.detail||'').length>60?'...':''}</div>
    ${adminBtns}
  </div>`;
}

const AVATAR_DEFAULT = 'data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%2244%22%20height%3D%2244%22%20viewBox%3D%220%200%2044%2044%22%3E%3Ccircle%20cx%3D%2222%22%20cy%3D%2222%22%20r%3D%2222%22%20fill%3D%22%23C8E6C9%22%2F%3E%3Ccircle%20cx%3D%2222%22%20cy%3D%2217%22%20r%3D%227%22%20fill%3D%22%232E7D32%22%2F%3E%3Cellipse%20cx%3D%2222%22%20cy%3D%2236%22%20rx%3D%2212%22%20ry%3D%228%22%20fill%3D%22%232E7D32%22%2F%3E%3C%2Fsvg%3E';

function renderUserList() {
  const el = document.getElementById('user-list');
  if (!users.length) { el.innerHTML = '<div class="empty-state"><span class="material-icons">people</span>ไม่มีข้อมูล</div>'; return; }
  const isCallerAdmin = currentUser && currentUser.role === 'admin';
  el.innerHTML = users.map(u => {
    const isTargetAdmin = u.role === 'admin';
    const canEdit = isCallerAdmin || !isTargetAdmin;
    const roleLabel = {admin:'👑 Admin',supervisor:'👔 หัวหน้า',accountant:'📊 ฝ่ายบัญชี',manager:'🏢 ผู้บริหาร',user:'👤 User'}[u.role]||u.role;
    const editBtn = canEdit
      ? `<button class="btn-sm-icon" style="background:var(--accent-light);color:var(--accent);" onclick="openEditUser('${u.lineUid}')"><span class="material-icons" style="font-size:1rem;">edit</span></button>`
      : `<button class="btn-sm-icon" style="background:#f5f5f5;color:#bbb;cursor:not-allowed;" disabled title="ไม่มีสิทธิ์แก้ไข admin"><span class="material-icons" style="font-size:1rem;">lock</span></button>`;
    const delBtn = canEdit
      ? `<button class="btn-sm-icon ms-1" style="background:var(--danger-light);color:var(--danger);" onclick="deleteUser('${u.lineUid}','${u.name}')"><span class="material-icons" style="font-size:1rem;">delete</span></button>`
      : '';
    return `
    <div class="user-row">
      <img src="${u.avatar||AVATAR_DEFAULT}" alt="">
      <div class="flex-grow-1">
        <div style="font-weight:600;font-size:.88rem;">${u.name}</div>
        <div style="font-size:.78rem;color:var(--gray-500);">${u.dept||'-'} • ${roleLabel}</div>
      </div>
      ${editBtn}${delBtn}
    </div>`;
  }).join('');
}

function renderVehicleList() {
  const el = document.getElementById('vehicle-list');
  if (!vehicles.length) { el.innerHTML = '<div class="empty-state"><span class="material-icons">directions_car</span>ไม่มีข้อมูล</div>'; return; }
  el.innerHTML = vehicles.map(v => `
    <div class="user-row">
      <span class="material-icons" style="color:var(--primary);font-size:1.4rem;">directions_car</span>
      <div class="flex-grow-1">
        <div style="font-weight:600;font-size:.88rem;">${v.plate}</div>
        <div style="font-size:.78rem;color:var(--gray-500);">${v.model||'-'}</div>
      </div>
      <button class="btn-sm-icon" style="background:var(--accent-light);color:var(--accent);" onclick="openEditVehicle(this)" data-plate="${v.plate}" data-model="${(v.model||'').replace(/"/g,'&quot;')}"><span class="material-icons">edit</span></button>
      <button class="btn-sm-icon ms-1" style="background:var(--danger-light);color:var(--danger);" onclick="deleteVehicle('${v.plate}')"><span class="material-icons">delete</span></button>
    </div>`).join('');
}

function populateVehicleSelect() {
  const sel = document.getElementById('f-plate');
  const prev = sel.value; // เก็บค่าที่เลือกไว้ก่อน
  sel.innerHTML = '<option value="">-- เลือกเลขทะเบียน --</option>' +
    vehicles.map(v => `<option value="${v.plate}">${v.plate}${v.model?' ('+v.model+')':''}</option>`).join('');
  if (prev) sel.value = prev; // restore ค่าเดิม ถ้ายังมีอยู่ใน list
}


/* ============================================================
   ข้อ 5 — ประวัติสถานะ
   ============================================================ */
async function openStatusLog(jobId) {
  document.getElementById('log-job-id').textContent = jobId;
  document.getElementById('status-log-tbody').innerHTML = '<tr><td colspan="5" class="text-center py-3">กำลังโหลด...</td></tr>';
  document.getElementById('status-log-cards').innerHTML = '<div class="p-3 text-center">กำลังโหลด...</div>';
  showModal('modalStatusLog');
  const res = await gasCall('getStatusLog', { jobId });
  if (res.status !== 'ok') {
    showToast('โหลดประวัติไม่สำเร็จ', 'error');
    const emptyMsg = '<tr><td colspan="5" class="text-center py-4 text-muted">โหลดไม่สำเร็จ</td></tr>';
    document.getElementById('status-log-tbody').innerHTML = emptyMsg;
    document.getElementById('status-log-cards').innerHTML = '<div class="p-3 text-center text-muted">โหลดไม่สำเร็จ</div>';
    return;
  }
  const logs = res.logs || [];
  const statusColor = s => ({'รอดำเนินการ':'#F57F17','อนุมัติ':'#2E7D32','กำลังซ่อม':'#1565C0','เสร็จสิ้น':'#1B5E20','ส่งกลับแก้ไข':'#6A1B9A','ไม่อนุมัติ':'#C62828'})[s] || '#555';

  // หาชื่อ+ตำแหน่งจาก users array
  const getChanger = uid => {
    const u = users.find(x => x.lineUid === uid);
    if (!u) return uid || '-';
    return `${u.name}${u.dept ? '<br><span style="font-size:.75rem;color:var(--gray-400);">' + u.dept + '</span>' : ''}`;
  };

  if (!logs.length) {
    document.getElementById('status-log-tbody').innerHTML = '<tr><td colspan="5" class="text-center py-4 text-muted">ยังไม่มีประวัติ</td></tr>';
    document.getElementById('status-log-cards').innerHTML = '<div class="p-4 text-center text-muted" style="color:var(--gray-400);"><span class="material-icons" style="font-size:2rem;display:block;margin-bottom:.5rem;">history</span>ยังไม่มีประวัติ</div>';
    return;
  }

  document.getElementById('status-log-tbody').innerHTML = logs.map(l => `
    <tr>
      <td style="font-size:.82rem;white-space:nowrap;">${formatDate(l.changedAt)}</td>
      <td><span style="color:${statusColor(l.oldStatus)};font-weight:600;">${l.oldStatus||'-'}</span></td>
      <td><span style="color:${statusColor(l.newStatus)};font-weight:600;">→ ${l.newStatus}</span></td>
      <td style="font-size:.82rem;">${l.note||'-'}</td>
      <td style="font-size:.82rem;line-height:1.3;">${getChanger(l.changedBy)}</td>
    </tr>`).join('');

  document.getElementById('status-log-cards').innerHTML = logs.map(l => `
    <div class="repair-card mb-2">
      <div style="font-size:.78rem;color:var(--gray-500);">${formatDate(l.changedAt)}</div>
      <div style="display:flex;gap:.5rem;align-items:center;margin:.25rem 0;flex-wrap:wrap;">
        <span style="color:${statusColor(l.oldStatus)};font-weight:600;font-size:.85rem;">${l.oldStatus||'-'}</span>
        <span class="material-icons" style="font-size:.9rem;color:var(--gray-400);">arrow_forward</span>
        <span style="color:${statusColor(l.newStatus)};font-weight:600;font-size:.85rem;">${l.newStatus}</span>
      </div>
      ${l.note ? `<div style="font-size:.8rem;color:var(--gray-600);">📝 ${l.note}</div>` : ''}
      <div style="font-size:.78rem;color:var(--gray-500);margin-top:.25rem;">
        👤 ${(users.find(x=>x.lineUid===l.changedBy)?.name)||l.changedBy||'-'}
        ${(users.find(x=>x.lineUid===l.changedBy)?.dept) ? `<span style="color:var(--gray-400);"> · ${users.find(x=>x.lineUid===l.changedBy).dept}</span>` : ''}
      </div>
    </div>`).join('');
}

/* ============================================================
   ข้อ 6 — ประวัติตามทะเบียน
   ============================================================ */
function populateVehicleHistorySelect() {
  const sel = document.getElementById('vh-plate-select');
  if (!sel) return;
  const prev = sel.value;
  sel.innerHTML = '<option value="">-- เลือกทะเบียนรถ --</option>' +
    vehicles.map(v => `<option value="${v.plate}">${v.plate}${v.model?' ('+v.model+')':''}</option>`).join('');
  if (prev) sel.value = prev;
}

async function loadVehicleHistory() {
  const plate = document.getElementById('vh-plate-select').value;
  if (!plate) { showToast('กรุณาเลือกทะเบียนรถ', 'error'); return; }
  showLoading(true);
  const res = await gasCall('getVehicleHistory', { plate });
  showLoading(false);
  if (res.status !== 'ok') { showToast('โหลดประวัติไม่สำเร็จ', 'error'); return; }

  const { jobs, totalCost, doneCount } = res;
  document.getElementById('vh-summary').style.display = 'block';
  document.getElementById('vh-total-jobs').textContent = jobs.length;
  document.getElementById('vh-done-jobs').textContent = doneCount;
  document.getElementById('vh-total-cost').textContent = Number(totalCost).toLocaleString();

  if (!jobs.length) {
    document.getElementById('vh-table-wrap').style.display = 'none';
    document.getElementById('vh-empty').style.display = 'flex';
    return;
  }
  document.getElementById('vh-empty').style.display = 'none';
  document.getElementById('vh-table-wrap').style.display = 'block';

  document.getElementById('vh-tbody').innerHTML = jobs.map(j => `
    <tr>
      <td style="font-weight:600;color:var(--primary);">${j.jobId}</td>
      <td style="font-size:.82rem;">${formatDate(j.createdAt)}</td>
      <td style="font-size:.82rem;max-width:200px;">${(j.detail||'').substring(0,60)}${(j.detail||'').length>60?'...':''}</td>
      <td><span class="badge-status ${statusClass(j.status)}">${j.status}</span></td>
      <td>${j.actualCost ? Number(j.actualCost).toLocaleString()+' ฿' : '-'}</td>
      <td style="font-size:.82rem;">${j.note||'-'}</td>
    </tr>`).join('');

  document.getElementById('vh-cards').innerHTML = jobs.map(j => `
    <div class="repair-card mb-2" onclick="openDetail('${j.jobId}')">
      <div class="d-flex justify-content-between"><span class="job-id">${j.jobId}</span><span class="badge-status ${statusClass(j.status)}">${j.status}</span></div>
      <div style="font-size:.82rem;color:var(--gray-500);">${formatDate(j.createdAt)}</div>
      <div style="font-size:.83rem;">${(j.detail||'').substring(0,60)}${(j.detail||'').length>60?'...':''}</div>
      ${j.actualCost ? `<div style="font-size:.82rem;color:var(--accent);font-weight:600;">💰 ${Number(j.actualCost).toLocaleString()} บาท</div>` : ''}
    </div>`).join('');
}

/* ============================================================
   ข้อ 7 — รายงานรายปี
   ============================================================ */
let _chartYearlyMonth = null;
let _chartYearlyCost  = null;

function initYearlySelectors() {
  const sel = document.getElementById('yearly-year');
  if (!sel || sel.options.length > 1) return;
  const yr = new Date().getFullYear();
  for (let y = yr; y >= yr - 4; y--) {
    const o = document.createElement('option');
    o.value = y; o.textContent = (y + 543) + ' (พ.ศ.)';
    if (y === yr) o.selected = true;
    sel.appendChild(o);
  }
}

async function generateYearlyReport() {
  const year = parseInt(document.getElementById('yearly-year').value);
  showLoading(true);
  const res = await gasCall('getYearlyReport', { year });
  showLoading(false);
  if (res.status !== 'ok') { showToast('โหลดรายงานไม่สำเร็จ', 'error'); return; }
  _yearlyReportData = res;
  renderYearlyReport(res);
}

async function renderYearlyReport(res) {
  await ensureChartJs();
  const { totalJobs, totalDone, totalCost, monthly, topPlates } = res;
  const MONTHS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];

  const APPROVED_DS_F = ['อนุมัติ','กำลังซ่อม','เสร็จสิ้น'];
  const totalApproved = res.totalApproved ?? allJobs.filter(j=>APPROVED_DS_F.includes(j.status)).length;
  const totalRejected = res.totalRejected ?? allJobs.filter(j=>j.status==='ไม่อนุมัติ').length;
  const approvedBudgetY = res.approvedBudget ?? allJobs.filter(j=>APPROVED_DS_F.includes(j.status)).reduce((s,j)=>s+(Number(j.estimate)||0),0);
  const summaryData = [
    { icon:'build_circle', label:'งานซ่อมทั้งหมด',   val:totalJobs,  color:'#1565C0', bg:'#E3F2FD' },
    { icon:'check_circle', label:'เสร็จสิ้น',          val:totalDone,  color:'#1B5E20', bg:'#E8F5E9' },
    { icon:'thumb_up',     label:'อนุมัติ',            val:totalApproved, color:'#2E7D32', bg:'#E8F5E9' },
    { icon:'thumb_down',   label:'ไม่อนุมัติ',         val:totalRejected, color:'#C62828', bg:'#FFEBEE' },
    { icon:'payments',     label:'ยอดอนุมัติ (ประเมิน)', val:Number(approvedBudgetY).toLocaleString()+' ฿', color:'#F57F17', bg:'#FFF8E1' },
    { icon:'receipt',      label:'ค่าใช้จ่ายจริงรวม', val:Number(totalCost).toLocaleString()+' ฿', color:'#E65100', bg:'#FFF3E0' },
    { icon:'percent',      label:'อัตราเสร็จสิ้น',    val:(totalJobs ? Math.round(totalDone/totalJobs*100) : 0)+'%', color:'#6A1B9A', bg:'#F3E5F5' },
  ];
  document.getElementById('yearly-summary-row').innerHTML = summaryData.map(s => `
    <div class="report-summary-card" style="background:${s.bg};border-color:${s.color}20;flex:1;min-width:130px;">
      <span class="material-icons" style="color:${s.color};">${s.icon}</span>
      <div style="font-size:1.4rem;font-weight:700;color:${s.color};">${s.val}</div>
      <div style="font-size:.78rem;color:var(--gray-600);">${s.label}</div>
    </div>`).join('');

  if (_chartYearlyMonth) _chartYearlyMonth.destroy();
  _chartYearlyMonth = new Chart(document.getElementById('chart-yearly-monthly'), {
    type: 'bar',
    data: {
      labels: MONTHS,
      datasets: [
        { label:'งานทั้งหมด', data: monthly.map(m=>m.total),    backgroundColor:'rgba(21,101,192,.7)',  borderRadius:4 },
        { label:'เสร็จสิ้น',  data: monthly.map(m=>m.done),     backgroundColor:'rgba(46,125,50,.7)',   borderRadius:4 },
        { label:'อนุมัติ',    data: monthly.map(m=>m.approved||0), backgroundColor:'rgba(56,142,60,.5)', borderRadius:4 },
        { label:'ไม่อนุมัติ', data: monthly.map(m=>m.rejected||0), backgroundColor:'rgba(198,40,40,.5)', borderRadius:4 },
      ]
    },
    options: { responsive:true, plugins:{ legend:{ position:'top' } }, scales:{ y:{ beginAtZero:true, ticks:{ stepSize:1 } } } }
  });

  if (_chartYearlyCost) _chartYearlyCost.destroy();
  _chartYearlyCost = new Chart(document.getElementById('chart-yearly-cost'), {
    type: 'line',
    data: {
      labels: MONTHS,
      datasets: [{ label:'ค่าใช้จ่าย (บาท)', data: monthly.map(m=>m.cost),
        borderColor:'#E65100', backgroundColor:'rgba(230,81,0,.1)',
        fill:true, tension:.4, pointRadius:4 }]
    },
    options: { responsive:true, plugins:{ legend:{ position:'top' } }, scales:{ y:{ beginAtZero:true } } }
  });

  const maxCount = topPlates[0]?.count || 1;
  document.getElementById('yearly-top-plates').innerHTML = topPlates.length
    ? topPlates.map((p,i) => `
      <div style="display:flex;align-items:center;gap:.75rem;margin-bottom:.6rem;">
        <div style="width:24px;font-weight:700;color:var(--gray-500);font-size:.85rem;">${i+1}</div>
        <div style="font-weight:600;min-width:100px;font-size:.88rem;">${p.plate}</div>
        <div style="flex:1;background:var(--gray-100);border-radius:4px;overflow:hidden;">
          <div style="width:${Math.round(p.count/maxCount*100)}%;background:var(--primary);height:20px;border-radius:4px;"></div>
        </div>
        <div style="min-width:60px;font-size:.82rem;color:var(--gray-600);">${p.count} ครั้ง</div>
        <div style="min-width:80px;font-size:.82rem;color:var(--accent);font-weight:600;">${Number(p.cost).toLocaleString()} ฿</div>
      </div>`).join('')
    : '<div class="empty-state"><span class="material-icons">directions_car</span>ไม่มีข้อมูล</div>';
}

/* ============================================================
   ข้อ 7 — Export Yearly CSV และ PDF
   ============================================================ */
let _yearlyReportData = null; // เก็บข้อมูลล่าสุดที่ generate


function exportYearlyCSV() {
  if (!_yearlyReportData) { showToast('กรุณาสร้างรายงานก่อน', 'error'); return; }
  const { year, monthly, topPlates, totalJobs, totalDone, totalCost } = _yearlyReportData;
  const MONTHS = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  const BOM = '﻿';
  const esc = v => '"' + String(v||'').replace(/"/g,'""') + '"';

  let csv = BOM + `รายงานสรุปรายปี ${year+543}

`;
  csv += `ภาพรวมทั้งปี
`;
  csv += `งานทั้งหมด,${totalJobs}
เสร็จสิ้น,${totalDone}
ค่าใช้จ่ายรวม,${totalCost}

`;
  csv += `สรุปรายเดือน
`;
  csv += `เดือน,งานทั้งหมด,เสร็จสิ้น,ค่าใช้จ่าย
`;
  monthly.forEach((m,i) => { csv += `${MONTHS[i]},${m.total},${m.done},${m.cost}
`; });
  csv += `
รถที่ซ่อมบ่อย
`;
  csv += `อันดับ,ทะเบียน,จำนวนครั้ง,ค่าใช้จ่ายรวม
`;
  topPlates.forEach((p,i) => { csv += `${i+1},${esc(p.plate)},${p.count},${p.cost}
`; });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `รายงานรายปี_${year+543}.csv`;
  a.click();
  showToast('ดาวน์โหลด CSV สำเร็จ ✅');
}

async function printYearlyReport() {
  if (!_yearlyReportData) { showToast('กรุณาสร้างรายงานก่อน', 'error'); return; }
  const { year, monthly, topPlates, totalJobs, totalDone, totalCost } = _yearlyReportData;
  const MONTHS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  const printDate = new Date().toLocaleDateString('th-TH', { year:'numeric', month:'long', day:'numeric' });

  const monthRows = monthly.map((m,i) => `
    <tr>
      <td>${MONTHS[i]}</td>
      <td style="text-align:center;">${m.total}</td>
      <td style="text-align:center;">${m.done}</td>
      <td style="text-align:right;">${Number(m.cost).toLocaleString()} ฿</td>
    </tr>`).join('');

  const plateRows = topPlates.map((p,i) => `
    <tr>
      <td style="text-align:center;">${i+1}</td>
      <td>${p.plate}</td>
      <td style="text-align:center;">${p.count}</td>
      <td style="text-align:right;">${Number(p.cost).toLocaleString()} ฿</td>
    </tr>`).join('');

  const html = `\u003c!DOCTYPE html>\u003chtml lang="th"><head><meta charset="UTF-8">
  <style>
    body { font-family: 'Sarabun', sans-serif; font-size: 13px; color: #222; margin: 20px; }
    h2 { color: #B71C1C; text-align: center; margin-bottom: 4px; }
    .subtitle { text-align:center; color:#555; margin-bottom:16px; font-size:12px; }
    .summary { display:flex; gap:12px; margin-bottom:16px; }
    .summary-card { flex:1; border:1px solid #e0e0e0; border-radius:8px; padding:10px; text-align:center; }
    .summary-card .val { font-size:1.4rem; font-weight:700; color:#B71C1C; }
    .summary-card .lbl { font-size:11px; color:#666; }
    table { width:100%; border-collapse:collapse; margin-bottom:16px; }
    th { background:#B71C1C; color:#fff; padding:6px 10px; font-size:12px; }
    td { padding:5px 10px; border-bottom:1px solid #eee; }
    tr:nth-child(even) td { background:#f9f9f9; }
    h3 { color:#7F0000; font-size:13px; margin:12px 0 6px; border-left:4px solid #B71C1C; padding-left:8px; }
    .ft { display:flex; justify-content:space-between; font-size:11px; color:#888; margin-top:16px; }
  </style></head><body>
  <h2>รายงานสรุปรายปี พ.ศ. ${year+543}</h2>
  <div class="subtitle">พิมพ์: ${printDate}</div>
  <div class="summary">
    <div class="summary-card"><div class="val">${totalJobs}</div><div class="lbl">งานทั้งหมด</div></div>
    <div class="summary-card"><div class="val">${totalDone}</div><div class="lbl">เสร็จสิ้น</div></div>
    <div class="summary-card"><div class="val">${totalJobs ? Math.round(totalDone/totalJobs*100) : 0}%</div><div class="lbl">อัตราเสร็จ</div></div>
    <div class="summary-card"><div class="val" style="font-size:1rem;">${Number(totalCost).toLocaleString()}</div><div class="lbl">ค่าใช้จ่ายรวม (฿)</div></div>
  </div>
  <h3>สรุปรายเดือน</h3>
  <table><thead><tr><th>เดือน</th><th>งานทั้งหมด</th><th>เสร็จสิ้น</th><th>ค่าใช้จ่าย</th></tr></thead>
  <tbody>${monthRows}</tbody></table>
  <h3>รถที่ซ่อมบ่อยที่สุด (Top 10)</h3>
  <table><thead><tr><th>อันดับ</th><th>ทะเบียน</th><th>จำนวนครั้ง</th><th>ค่าใช้จ่ายรวม</th></tr></thead>
  <tbody>${plateRows}</tbody></table>
  <div class="ft"><span>รายงานรายปี ${year+543} | ระบบแจ้งซ่อมรถยนต์</span><span>${printDate}</span></div>
  <script>window.onload=function(){window.print();}<\/script>
  \u003c/body>\u003c/html>`;

  const isMobile = /Android|iPhone|iPad|iPod/i.test(navigator.userAgent);
  if (isMobile) {
    showToast('กำลังสร้าง PDF... รอสักครู่');
    await savePDF(html, `รายงานรายปี_${year+543}.pdf`);
  } else {
    const win = window.open('','_blank');
    if (!win) { showToast('กรุณาอนุญาต Popup', 'error'); return; }
    win.document.write(html);
    win.document.close();
  }
}

/* ============================================================
   QR CODE
   ============================================================ */
let _qrInstance       = null;
let _qrPosterInstance = null;

async function initQRCode() {
  await ensureQRCode();
  const url = window.location.href.split('?')[0];

  document.getElementById('qr-url-display').textContent = url;
  document.getElementById('qr-url-label').textContent   = url;

  // ล้าง QR เก่า
  const div1 = document.getElementById('qr-div');
  const div2 = document.getElementById('qr-poster-div');
  div1.innerHTML = '';
  div2.innerHTML = '';

  // สร้าง QR ใหม่ด้วย qrcodejs (render เป็น <img> + <canvas> ใน div)
  new QRCode(div1, { text: url, width: 220, height: 220, correctLevel: QRCode.CorrectLevel.M });
  new QRCode(div2, { text: url, width: 160, height: 160, correctLevel: QRCode.CorrectLevel.M });
}

function downloadQR() {
  setTimeout(() => {
    const img = document.querySelector('#qr-div img');
    if (!img) { showToast('ยังไม่มี QR Code', 'error'); return; }
    const link = document.createElement('a');
    link.download = 'QR_แจ้งซ่อม.png';
    link.href = img.src;
    link.click();
    showToast('บันทึก QR Code แล้ว ✅');
  }, 300);
}

function printQR() {
  setTimeout(() => {
    const img = document.querySelector('#qr-div img');
    if (!img) { showToast('ยังไม่มี QR Code', 'error'); return; }
    const url  = document.getElementById('qr-url-display').textContent;
    const imgData = img.src;
    const win  = window.open('', '_blank', 'width=400,height=500');
    win.document.write(`\u003c!DOCTYPE html>\u003chtml><head><title>QR แจ้งซ่อม</title>
      <style>
        body{margin:0;display:flex;align-items:center;justify-content:center;min-height:100vh;background:#f5f5f5;}
        .card{background:#fff;border-radius:16px;padding:32px 24px;text-align:center;box-shadow:0 4px 24px rgba(0,0,0,.1);max-width:320px;}
        h2{color:#7F0000;font-family:sans-serif;margin:0 0 16px;}
        img{width:220px;height:220px;border:2px solid #eee;border-radius:8px;padding:8px;}
        p{font-size:12px;color:#666;margin-top:12px;word-break:break-all;}
      </style></head><body>
      <div class="card">
        <h2>🔧 แจ้งซ่อมรถยนต์</h2>
        <img src="${imgData}">
        <p>สแกน QR Code เพื่อแจ้งซ่อม</p>
        <p style="font-size:10px;color:#999;">${url}</p>
      </div>
      <script>window.onload=function(){window.print();}<\/script>
      \u003c/body>\u003c/html>`);
    win.document.close();
  }, 300);
}

function printPoster() {
  setTimeout(() => {
    const img = document.querySelector('#qr-poster-div img');
    if (!img) { showToast('ยังไม่มี QR Code', 'error'); return; }
    const url     = document.getElementById('qr-url-display').textContent;
    const imgData = img.src;
    const win     = window.open('', '_blank', 'width=500,height=700');
    win.document.write(`\u003c!DOCTYPE html>\u003chtml><head><title>โปสเตอร์แจ้งซ่อม</title>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;600;700&display=swap');
        *{margin:0;padding:0;box-sizing:border-box;}
        body{display:flex;align-items:center;justify-content:center;min-height:100vh;background:#f0f0f0;}
        .poster{background:linear-gradient(160deg,#7F0000 0%,#B71C1C 60%,#C62828 100%);
          border-radius:20px;padding:40px 32px;text-align:center;color:#fff;
          font-family:'Prompt',sans-serif;width:360px;box-shadow:0 8px 32px rgba(0,0,0,.25);}
        .logo{font-size:2.5rem;margin-bottom:8px;}
        h1{font-size:1.5rem;font-weight:700;margin-bottom:4px;letter-spacing:.5px;}
        .sub{font-size:.9rem;opacity:.85;margin-bottom:20px;}
        .qr-box{background:#fff;border-radius:14px;padding:16px;display:inline-block;margin-bottom:16px;}
        .qr-box img{width:200px;height:200px;display:block;}
        .cta{font-size:1rem;font-weight:600;margin-bottom:8px;}
        .url{font-size:.6rem;opacity:.6;word-break:break-all;margin-top:8px;}
        .steps{background:rgba(255,255,255,.15);border-radius:10px;padding:12px 16px;margin-top:16px;text-align:left;}
        .steps p{font-size:.78rem;margin-bottom:4px;opacity:.9;}
        .steps p:last-child{margin-bottom:0;}
      </style></head><body>
      <div class="poster">
        <div class="logo">🔧</div>
        <h1>แจ้งซ่อมรถยนต์</h1>
        <div class="sub">ระบบแจ้งซ่อม Royal-D</div>
        <div class="qr-box"><img src="${imgData}"></div>
        <div class="cta">📱 สแกน QR Code เพื่อแจ้งซ่อม</div>
        <div class="steps">
          <p>1️⃣  เปิดกล้องมือถือ / LINE</p>
          <p>2️⃣  ส่อง QR Code</p>
          <p>3️⃣  กรอกข้อมูลแจ้งซ่อม</p>
        </div>
        <div class="url">${url}</div>
      </div>
      <script>window.onload=function(){window.print();}<\/script>
      \u003c/body>\u003c/html>`);
    win.document.close();
  }, 300);
}

/* ============================================================
   NAVIGATION
   ============================================================ */
function navigate(page) {
  document.querySelectorAll('.page-section').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.sidebar-item').forEach(el => el.classList.remove('active'));
  const pEl = document.getElementById('page-' + page);
  if (pEl) pEl.classList.add('active');
  const sEl = document.querySelector(`.sidebar-item[data-page="${page}"]`);
  if (sEl) sEl.classList.add('active');
  if (page === 'repair-list') renderRepairTable();
  if (page === 'admin-users') renderUserList();
  if (page === 'admin-vehicles') renderVehicleList();
  if (page === 'admin-report')   { initReportSelectors(); setTimeout(generateReport, 50); }
  if (page === 'admin-yearly')   { initYearlySelectors(); }
  if (page === 'admin-qr')       { setTimeout(initQRCode, 100); }
  if (page === 'vehicle-history')    { populateVehicleHistorySelect(); }
  if (page === 'manager-jobs')    { _mgrFilter='all'; renderManagerJobs(); }
  if (page === 'accountant-jobs') { _purFilter='all'; renderAccountantJobs(); }

  // FAB — แสดงเฉพาะหน้า dashboard และ repair-list และเฉพาะ non-manager
  const fab = document.getElementById('fab-repair');
  if (fab) {
    const showFab = ['dashboard','repair-list'].includes(page) && !['manager','accountant'].includes(currentUser?.role);
    fab.style.display = showFab ? 'flex' : 'none';
  }

  // close sidebar on mobile
  if (window.innerWidth < 768) document.getElementById('sidebar').style.display='none';
}

/* ============================================================
   MANAGER — Jobs with filter
   ============================================================ */
let _mgrFilter = 'all';
let _purFilter = 'all';

function filterManagerJobs(status) {
  _mgrFilter = status;
  document.querySelectorAll('[data-mgr-status]').forEach(el => {
    el.classList.toggle('active', el.dataset.mgrStatus === status);
  });
  renderManagerJobs();
}

function renderManagerJobs() {
  // ===== stats =====
  // งานที่ manager เคยอนุมัติ = สถานะที่ downstream จาก อนุมัติ
  const APPROVED_DOWNSTREAM = ['อนุมัติ','กำลังซ่อม','เสร็จสิ้น'];
  const pending   = allJobs.filter(j => j.status === 'รอการอนุมัติ').length;
  const approved  = allJobs.filter(j => APPROVED_DOWNSTREAM.includes(j.status)).length;
  const done      = allJobs.filter(j => j.status === 'เสร็จสิ้น').length;
  const totalCost = allJobs
    .filter(j => APPROVED_DOWNSTREAM.includes(j.status))
    .reduce((s, j) => s + (Number(j.estimate) || 0), 0);
  document.getElementById('mgr-stat-pending').textContent  = pending;
  document.getElementById('mgr-stat-approved').textContent = approved;
  document.getElementById('mgr-stat-done').textContent     = done;
  document.getElementById('mgr-stat-cost').textContent     = totalCost.toLocaleString();

  // badge sidebar
  const badge = document.getElementById('badge-approval');
  if (badge) { badge.textContent = pending; badge.style.display = pending ? 'inline' : 'none'; }

  // ===== filter + search + sort =====
  const q       = (document.getElementById('mgr-search')?.value || '').toLowerCase().trim();
  const sortVal = document.getElementById('mgr-sort')?.value || 'oldest';

  // สถานะที่ manager เห็น — รวม downstream ของ อนุมัติ ด้วย
  const MGR_VISIBLE = ['รอการอนุมัติ','อนุมัติ','กำลังซ่อม','เสร็จสิ้น','ส่งกลับแก้ไข','ไม่อนุมัติ'];

  let jobs = allJobs.filter(j => MGR_VISIBLE.includes(j.status));
  if (_mgrFilter !== 'all') jobs = jobs.filter(j => j.status === _mgrFilter);
  if (q) jobs = jobs.filter(j =>
    (j.plate||'').toLowerCase().includes(q) ||
    (j.userName||'').toLowerCase().includes(q) ||
    (j.jobId||'').toLowerCase().includes(q) ||
    (j.detail||'').toLowerCase().includes(q)
  );
  jobs = [...jobs].sort((a, b) => {
    const ta = new Date(a.createdAt).getTime();
    const tb = new Date(b.createdAt).getTime();
    return sortVal === 'oldest' ? ta - tb : tb - ta;
  });

  const tbody = document.getElementById('mgr-jobs-tbody');
  const cards = document.getElementById('mgr-jobs-cards');
  const empty = document.getElementById('mgr-jobs-empty');

  if (!jobs.length) {
    // PC: แสดง empty ใน tbody
    tbody.innerHTML = `<tr><td colspan="8">
      <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:2.5rem;color:var(--gray-400);">
        <span class="material-icons" style="font-size:2.5rem;margin-bottom:.5rem;">inbox</span>
        <span style="font-size:.9rem;">ไม่พบรายการ</span>
      </div></td></tr>`;
    // Mobile: แสดง empty ใน cards
    cards.innerHTML = `<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:2.5rem;color:var(--gray-400);">
      <span class="material-icons" style="font-size:2.5rem;margin-bottom:.5rem;">inbox</span>
      <span style="font-size:.9rem;">ไม่พบรายการ</span>
    </div>`;
    empty.style.display = 'none'; // ซ่อน empty-state นอก table เสมอ
    return;
  }
  empty.style.display = 'none';

  // ปุ่มการดำเนินการ — เฉพาะ รอการอนุมัติ แสดง 3 ปุ่ม
  const actionBtns = j => {
    const viewBtn = `<button class="btn-outline-custom btn-sm" onclick="openDetail('${j.jobId}')" title="ดูรายละเอียด">
      <span class="material-icons" style="font-size:.9rem;">visibility</span>
    </button>`;
    if (j.status !== 'รอการอนุมัติ') return viewBtn;
    return `<div class="d-flex gap-1">
      ${viewBtn}
      <button class="btn-primary-custom btn-sm" style="background:var(--success);padding:.3rem .6rem;" onclick="quickDecision('${j.jobId}','อนุมัติ')" title="อนุมัติ">
        <span class="material-icons" style="font-size:.9rem;">check</span>
      </button>
      <button class="btn-primary-custom btn-sm" style="background:var(--danger);padding:.3rem .6rem;" onclick="quickDecision('${j.jobId}','ไม่อนุมัติ')" title="ไม่อนุมัติ">
        <span class="material-icons" style="font-size:.9rem;">close</span>
      </button>
      <button class="btn-primary-custom btn-sm" style="background:var(--warning);color:#333;padding:.3rem .6rem;" onclick="quickDecision('${j.jobId}','ส่งกลับแก้ไข')" title="ส่งกลับแก้ไข">
        <span class="material-icons" style="font-size:.9rem;">undo</span>
      </button>
    </div>`;
  };

  // ===== Desktop =====
  tbody.innerHTML = jobs.map(j => `
    <tr style="${j.status==='รอการอนุมัติ'?'background:rgba(255,193,7,.05);':''}">
      <td style="padding:.75rem .875rem;white-space:nowrap;">
        <span style="font-weight:700;color:var(--primary-dark);font-size:.82rem;">${j.jobId}</span>
      </td>
      <td style="padding:.75rem .875rem;font-size:.85rem;">${j.userName}</td>
      <td style="padding:.75rem .875rem;">
        <span style="font-weight:600;font-size:.88rem;">${j.plate}</span>
      </td>
      <td style="padding:.75rem .875rem;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:.84rem;color:var(--gray-700);" title="${j.detail}">${j.detail}</td>
      <td style="padding:.75rem .875rem;text-align:right;font-weight:600;font-size:.88rem;white-space:nowrap;">
        ${j.estimate ? Number(j.estimate).toLocaleString()+' ฿' : '<span style="color:var(--gray-400);">-</span>'}
      </td>
      <td style="padding:.75rem .875rem;">
        <span class="badge-status ${statusClass(j.status)}">${j.status}</span>
      </td>
      <td style="padding:.75rem .875rem;white-space:nowrap;font-size:.82rem;color:var(--gray-600);">${formatDate(j.createdAt)}</td>
      <td style="padding:.75rem .875rem;text-align:center;">
        <div style="display:flex;gap:.4rem;justify-content:center;align-items:center;">${actionBtns(j)}</div>
      </td>
    </tr>`).join('');

  // ===== Mobile =====
  const borderColor = s => ({
    'รอการอนุมัติ':'var(--warning)',
    'อนุมัติ':'var(--success)',
    'ไม่อนุมัติ':'var(--danger)',
    'ส่งกลับแก้ไข':'#9C27B0',
    'เสร็จสิ้น':'#1B5E20'
  }[s] || 'var(--gray-300)');

  cards.innerHTML = jobs.map(j => `
    <div style="background:#fff;border-radius:12px;border:1px solid var(--gray-200);border-left:4px solid ${borderColor(j.status)};
                padding:1rem;margin-bottom:.75rem;box-shadow:0 1px 4px rgba(0,0,0,.06);">
      <!-- Row 1: JobID + Badge -->
      <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:.5rem;">
        <span style="font-weight:700;color:var(--primary-dark);font-size:.88rem;">${j.jobId}</span>
        <span class="badge-status ${statusClass(j.status)}">${j.status}</span>
      </div>
      <!-- Row 2: ชื่อ + ทะเบียน -->
      <div style="display:flex;gap:.5rem;align-items:center;margin-bottom:.4rem;flex-wrap:wrap;">
        <span style="font-size:.82rem;color:var(--gray-600);">👤 ${j.userName}</span>
        <span style="color:var(--gray-300);">|</span>
        <span style="font-size:.85rem;font-weight:600;color:var(--gray-800);">🚗 ${j.plate}</span>
      </div>
      <!-- Row 3: อาการ -->
      <div style="font-size:.84rem;color:var(--gray-700);margin-bottom:.5rem;line-height:1.4;
                  overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;">
        ${j.detail}
      </div>
      <!-- Row 4: ประเมิน + วันที่ -->
      <div style="display:flex;gap:1rem;font-size:.8rem;color:var(--gray-500);margin-bottom:.75rem;flex-wrap:wrap;">
        <span>💰 <strong style="color:var(--gray-800);">${j.estimate ? Number(j.estimate).toLocaleString()+' ฿' : 'ไม่ระบุ'}</strong></span>
        <span>📅 ${formatDate(j.createdAt)}</span>
        ${j.note ? `<span style="color:var(--warning-dark);">📝 ${j.note}</span>` : ''}
      </div>
      <!-- Row 5: ปุ่มการดำเนินการ -->
      <div style="display:flex;gap:.5rem;flex-wrap:wrap;border-top:1px solid var(--gray-100);padding-top:.75rem;">
        <button onclick="openDetail('${j.jobId}')"
          style="flex:1;min-width:100px;display:flex;align-items:center;justify-content:center;gap:.3rem;
                 padding:.5rem .75rem;border-radius:8px;border:1.5px solid var(--gray-300);
                 background:#fff;color:var(--gray-700);font-size:.82rem;font-weight:600;cursor:pointer;">
          <span class="material-icons" style="font-size:.95rem;">visibility</span> รายละเอียด
        </button>
        ${j.status==='รอการอนุมัติ' ? `
        <button onclick="quickDecision('${j.jobId}','อนุมัติ')"
          style="flex:1;min-width:80px;display:flex;align-items:center;justify-content:center;gap:.3rem;
                 padding:.5rem .75rem;border-radius:8px;border:none;
                 background:#E8F5E9;color:#2E7D32;font-size:.82rem;font-weight:700;cursor:pointer;">
          <span class="material-icons" style="font-size:.95rem;">check_circle</span> อนุมัติ
        </button>
        <button onclick="quickDecision('${j.jobId}','ไม่อนุมัติ')"
          style="flex:1;min-width:80px;display:flex;align-items:center;justify-content:center;gap:.3rem;
                 padding:.5rem .75rem;border-radius:8px;border:none;
                 background:#FFEBEE;color:#C62828;font-size:.82rem;font-weight:700;cursor:pointer;">
          <span class="material-icons" style="font-size:.95rem;">cancel</span> ไม่อนุมัติ
        </button>
        <button onclick="quickDecision('${j.jobId}','ส่งกลับแก้ไข')"
          style="flex:1;min-width:80px;display:flex;align-items:center;justify-content:center;gap:.3rem;
                 padding:.5rem .75rem;border-radius:8px;border:none;
                 background:#FFF8E1;color:#F57F17;font-size:.82rem;font-weight:700;cursor:pointer;">
          <span class="material-icons" style="font-size:.95rem;">undo</span> ส่งกลับ
        </button>` : ''}
      </div>
    </div>`).join('');
}

// ตัดสินใจด่วน — ถ้าต้องใส่เหตุผลให้เปิด modal ก่อน
function quickDecision(jobId, decision) {
  if (['ไม่อนุมัติ','ส่งกลับแก้ไข'].includes(decision)) {
    const icons = { 'ไม่อนุมัติ':'❌', 'ส่งกลับแก้ไข':'⚠️' };
    document.getElementById('mgr-note-jobid').value    = jobId;
    document.getElementById('mgr-note-decision').value = decision;
    document.getElementById('mgr-note-text').value     = '';
    document.getElementById('mgr-note-title').textContent = `${icons[decision]} ${decision}`;
    const btn = document.getElementById('mgr-note-confirm-btn');
    btn.style.background = decision === 'ไม่อนุมัติ' ? 'var(--danger)' : 'var(--warning)';
    btn.style.color = decision === 'ไม่อนุมัติ' ? '#fff' : '#333';
    closeModal('modalDetail'); // ปิด detail modal ถ้าเปิดอยู่
    showModal('modalManagerNote');
  } else {
    // อนุมัติ — ไม่ต้องใส่เหตุผล ทำเลย
    _submitManagerDecision(jobId, 'อนุมัติ', '');
  }
}

async function confirmManagerDecision() {
  const jobId    = document.getElementById('mgr-note-jobid').value;
  const decision = document.getElementById('mgr-note-decision').value;
  const note     = document.getElementById('mgr-note-text').value.trim();
  closeModal('modalManagerNote');
  // accountant ส่งกลับ หรือ manager ตัดสินใจ
  if (currentUser.role === 'accountant') {
    await _submitAccountantDecision(jobId, decision, note);
  } else {
    await _submitManagerDecision(jobId, decision, note);
  }
}

async function _submitManagerDecision(jobId, decision, note) {
  const icons = { 'อนุมัติ':'✅', 'ไม่อนุมัติ':'❌', 'ส่งกลับแก้ไข':'⚠️' };
  showLoading(true);
  const res = await gasCall('updateStatus', {
    jobId,
    status: decision,
    managerNote: note || '',  // บันทึกลง col M (managerNote) โดยเฉพาะ
    actualCost: '',
    adminUid: currentUser.lineUid,
  });
  showLoading(false);
  if (res.status === 'ok') {
    showToast(`${icons[decision]} ${decision}แล้ว`);
    sessionStorage.removeItem('_repairData');
    await loadData(true);
    renderManagerJobs();
  } else {
    showToast('เกิดข้อผิดพลาด: ' + (res.message||'unknown'), 'error');
  }
}

document.getElementById('sidebar-toggle').addEventListener('click', () => {
  const sb = document.getElementById('sidebar');
  sb.style.display = sb.style.display === 'none' ? 'block' : 'none';
});

/* ============================================================
   REPAIR MODAL
   ============================================================ */
function openRepairModal() {
  editingJobId = null;
  document.getElementById('repair-modal-title').textContent = 'แจ้งซ่อมรถยนต์';
  document.getElementById('repair-form').reset();
  document.getElementById('repair-img-preview').innerHTML = '<span class="material-icons" style="font-size:2rem;color:var(--gray-300);">photo_camera</span><div class="placeholder-text">คลิกเพื่อเลือกรูป (ไม่บังคับ)</div>';
  if (currentUser?.role !== 'admin') document.getElementById('admin-note-section').classList.add('d-none');
  showModal('modalRepair');
  // Warm-up GAS ทันทีที่เปิด modal — ปลุก cold start ล่วงหน้าขณะ user กรอกฟอร์ม
  gasCall('ping', {}).catch(() => {});
}

function openEditRepair(jobId) {
  const job = allJobs.find(j => j.jobId === jobId);
  if (!job) return;
  editingJobId = jobId;
  document.getElementById('repair-modal-title').textContent = 'แก้ไขใบแจ้งซ่อม';
  document.getElementById('f-job-id').value = jobId;
  document.getElementById('f-plate').value = job.plate;
  document.getElementById('f-mileage').value = job.mileage;
  document.getElementById('f-detail').value = job.detail;
  document.getElementById('f-estimate').value = job.estimate || '';
  document.getElementById('f-location').value = job.location;
  document.getElementById('f-note').value = job.note || '';
  if (job.imageUrl) {
    const box = document.getElementById('repair-img-preview');
    const urls = parseUrls(job.imageUrl);
    box.innerHTML = urls.map(u =>
      `<img src="${u}" style="height:80px;width:80px;object-fit:cover;border-radius:6px;border:2px solid var(--primary-xlight);" onerror="this.style.display='none'">`
    ).join('') + `<div style="font-size:.78rem;color:var(--gray-500);width:100%;text-align:center;margin-top:4px;">${urls.length} รูป (เลือกใหม่เพื่อแทนที่)</div>`;
  }
  showModal('modalRepair');
}

async function submitRepair() {
  if (submitting) return;
  const plate    = document.getElementById('f-plate').value.trim();
  const mileage  = document.getElementById('f-mileage').value.trim();
  const detail   = document.getElementById('f-detail').value.trim();
  const estimate = document.getElementById('f-estimate').value.trim();
  const location = document.getElementById('f-location').value.trim();
  const note     = document.getElementById('f-note').value.trim();

  if (!plate || !mileage || !detail) { showToast('กรุณากรอกข้อมูลที่จำเป็นให้ครบ', 'error'); return; }
  if (isNaN(mileage) || Number(mileage) < 0) { showToast('เลขไมล์ต้องเป็นตัวเลข', 'error'); return; }
  if (estimate && (isNaN(estimate) || Number(estimate) < 0)) { showToast('ประเมินค่าใช้จ่ายต้องเป็นตัวเลข', 'error'); return; }

  submitting = true;
  showLoading(true);

  // Upload รูปแจ้งซ่อม (parallel หลายรูป)
  let imageUrls = []; let imageUrl = ''; let viewUrl = '';
  const imgFiles = document.getElementById('f-image').files;
  if (imgFiles.length > 0) {
    try {
      const results = await uploadMultiImages(imgFiles, 'repair', editingJobId || '');
      imageUrls = results.map(r => r.url).filter(Boolean);
      imageUrl  = JSON.stringify(imageUrls);   // เก็บเป็น JSON array
      viewUrl   = results[0]?.viewUrl || '';   // viewUrl รูปแรก (backward compat)
      window._lastUploadedFileId = results[0]?.fileId || '';
    } catch(e) {
      showLoading(false); submitting = false;
      showToast('อัปโหลดรูปไม่สำเร็จ: ' + e.message, 'error'); return;
    }
  }

  const action = editingJobId ? 'updateJob' : 'createJob';
  const payload = {
    lineUid: currentUser.lineUid,
    userName: currentUser.name,
    plate, mileage, detail, estimate, location,
    imageUrl: imageUrl || (editingJobId ? undefined : ''),  // JSON array string หรือ string เดิม
    viewUrl,
    note,
    jobId: editingJobId || '',
    imageFileId: (!editingJobId && imgFiles?.length) ? (window._lastUploadedFileId || '') : ''
  };

  const res = await gasCall(action, payload);
  showLoading(false);
  submitting = false;

  if (res.status === 'ok') {
    showToast(editingJobId ? 'แก้ไขเรียบร้อย' : '✅ ส่งแจ้งซ่อมสำเร็จ');
    closeModal('modalRepair');
    // refresh ข้อมูลใน background โดยไม่แสดง loading ซ้ำ
    loadData(true);
    // ยิง LINE notify แบบ fire-and-forget (ไม่รอ — ไม่ให้ loading หมุนรอ)
    if (!editingJobId) {
      gasCall('notifyNewJob', {
        jobId: res.jobId || '',
        userName: currentUser.name,
        plate, mileage, detail, estimate, location,
        now: new Date().toISOString()
      }).catch(() => {});
    }
  } else {
    showToast('เกิดข้อผิดพลาด: ' + (res.message || ''), 'error');
  }
  submitting = false;
  showLoading(false);
}

/* ============================================================
   DETAIL MODAL
   ============================================================ */
function openDetail(jobId) {
  const j = allJobs.find(x => x.jobId === jobId);
  if (!j) return;
  const isAdmin = ['admin','supervisor','accountant'].includes(currentUser.role);
  const body = document.getElementById('detail-body');
  body.innerHTML = `
    <div class="d-flex justify-content-between align-items-center mb-3">
      <div>
        <div style="font-size:.78rem;color:var(--gray-500);">เลขที่งาน</div>
        <div style="font-weight:700;font-size:1.05rem;color:var(--primary-dark);">${j.jobId}</div>
      </div>
      <span class="badge-status ${statusClass(j.status)}" style="font-size:.85rem;">${j.status}</span>
    </div>
    ${j.note ? `<div class="alert" style="background:var(--warning-light);border:1px solid #FFE082;border-radius:8px;padding:.65rem 1rem;font-size:.85rem;margin-bottom:1rem;">
      <strong>📝 หมายเหตุ:</strong> ${j.note}</div>` : ''}
    ${['ผู้แจ้ง','เลขทะเบียน','เลขไมล์','อาการที่พบ','ประเมินค่าใช้จ่าย','สถานที่ซ่อม','วันที่แจ้ง'].map((label,i) => {
      const vals = [j.userName,j.plate,Number(j.mileage).toLocaleString()+' กม.',j.detail,(j.estimate?Number(j.estimate).toLocaleString()+' บาท':'-'),j.location,formatDate(j.createdAt)];
      return `<div class="detail-row"><div class="detail-label">${label}</div><div class="detail-value">${vals[i]}</div></div>`;
    }).join('')}
    ${j.actualCost ? `<div class="detail-row"><div class="detail-label">ค่าใช้จ่ายจริง</div><div class="detail-value text-success fw-bold">${Number(j.actualCost).toLocaleString()} บาท</div></div>` : ''}
    ${renderThumbRow(parseUrls(j.imageUrl), '📷 รูปประกอบ', j.viewUrl)}
    ${renderThumbRow(parseUrls(j.estimateImageUrl), '💰 ใบประเมินราคา', j.estimateViewUrl)}
    ${renderThumbRow(parseUrls(j.billUrl), '🧾 ใบเสร็จ / บิล', j.billViewUrl)}
  `;
  const footer = document.getElementById('detail-footer');
  footer.innerHTML = '';
  // ปุ่ม Print PDF ให้ทุกคนกดได้
  const pdfBtn = `<button class="btn-outline-custom btn-sm" onclick="printJobPDF('${j.jobId}')"><span class="material-icons" style="font-size:.9rem;">picture_as_pdf</span> PDF</button>`;

  const isManager     = currentUser.role === 'manager';
  const isAccountant  = currentUser.role === 'accountant';
  const isSupervisorAdmin = ['admin','supervisor'].includes(currentUser.role);

  if (isAdmin) {
    const logBtn = `<button class="btn-outline-custom btn-sm" onclick="openStatusLog('${j.jobId}')"><span class="material-icons" style="font-size:.9rem;">history</span> ประวัติ</button>`;
    if (j.status === 'เสร็จสิ้น') {
      footer.innerHTML = pdfBtn + logBtn;
    } else if (isAccountant) {
      // accountant: ประวัติ + เปลี่ยนสถานะ เฉพาะ รอตรวจสอบ (ไม่มีปุ่มแก้ไข)
      const statusBtn = j.status === 'รอตรวจสอบ'
        ? `<button class="btn-primary-custom btn-sm" onclick="closeModal('modalDetail');openStatusModal('${j.jobId}')"><span class="material-icons" style="font-size:.9rem;">update</span> ตรวจสอบ</button>`
        : '';
      footer.innerHTML = pdfBtn + logBtn + statusBtn;
    } else {
      footer.innerHTML = pdfBtn + logBtn + `
        ${(currentUser.role==='admin' ? j.status!=='ไม่อนุมัติ' : (currentUser.role==='accountant' ? false : ['รอดำเนินการ','ส่งกลับแก้ไข'].includes(j.status))) ? `<button class="btn-outline-custom btn-sm" onclick="closeModal('modalDetail');openEditRepair('${j.jobId}')"><span class="material-icons" style="font-size:.9rem;">edit</span> แก้ไข</button>` : ''}
        <button class="btn-primary-custom btn-sm" onclick="closeModal('modalDetail');openStatusModal('${j.jobId}')"><span class="material-icons" style="font-size:.9rem;">update</span> เปลี่ยนสถานะ</button>`;
    }
  } else if (isManager && j.status === 'รอการอนุมัติ') {
    footer.innerHTML = pdfBtn + `
      <button class="btn-primary-custom btn-sm" style="background:var(--success);" onclick="closeModal('modalDetail');quickDecision('${j.jobId}','อนุมัติ')">
        <span class="material-icons" style="font-size:.9rem;">check</span> อนุมัติ
      </button>
      <button class="btn-primary-custom btn-sm" style="background:var(--danger);" onclick="closeModal('modalDetail');quickDecision('${j.jobId}','ไม่อนุมัติ')">
        <span class="material-icons" style="font-size:.9rem;">close</span> ไม่อนุมัติ
      </button>
      <button class="btn-primary-custom btn-sm" style="background:var(--warning);color:#333;" onclick="closeModal('modalDetail');quickDecision('${j.jobId}','ส่งกลับแก้ไข')">
        <span class="material-icons" style="font-size:.9rem;">undo</span> ส่งกลับแก้ไข
      </button>`;
  } else {
    footer.innerHTML = pdfBtn;
  }
  showModal('modalDetail');
}

/* ============================================================
   STATUS MODAL
   ============================================================ */
function openStatusModal(jobId) {
  const j = allJobs.find(x => x.jobId === jobId);
  if (!j) return;
  if (['เสร็จสิ้น','ไม่อนุมัติ'].includes(j.status) && currentUser.role !== 'admin') {
    showToast('งานนี้ปิดแล้ว ไม่สามารถเปลี่ยนสถานะได้', 'error');
    return;
  }
  // accountant เปลี่ยนได้เฉพาะ รอตรวจสอบ
  if (currentUser.role === 'accountant' && j.status !== 'รอตรวจสอบ') {
    showToast('ฝ่ายบัญชีตรวจสอบได้เฉพาะงานที่รอตรวจสอบ', 'error');
    return;
  }

  // render options ตาม role
  const role = currentUser?.role || 'user';
  const roleOptions = {
    supervisor: (function() {
      // รอดำเนินการ / ส่งกลับแก้ไข → ส่งต่อผู้บริหารอย่างเดียว
      if (['รอดำเนินการ','ส่งกลับแก้ไข'].includes(j.status)) {
        return [{ v:'รอตรวจสอบ', l:'📊 ส่งให้ฝ่ายบัญชีตรวจสอบ' }];
      }
      // อนุมัติ → กำลังซ่อม หรือ เสร็จสิ้น
      if (j.status === 'อนุมัติ') {
        return [
          { v:'กำลังซ่อม', l:'🔧 กำลังซ่อม' },
          { v:'เสร็จสิ้น', l:'🏁 เสร็จสิ้น' },
        ];
      }
      // กำลังซ่อม → เสร็จสิ้น
      if (j.status === 'กำลังซ่อม') {
        return [{ v:'เสร็จสิ้น', l:'🏁 เสร็จสิ้น' }];
      }
      return [];
    })(),
    accountant: [
      { v:'รอการอนุมัติ', l:'✅ ผ่านการตรวจสอบ (ส่งผู้บริหาร)' },
      { v:'ส่งกลับแก้ไข', l:'↩️ ส่งกลับให้หัวหน้าแก้ไข' },
    ],
    manager: [
      { v:'รอการอนุมัติ', l:'รอการอนุมัติ' },
      { v:'อนุมัติ',       l:'อนุมัติ ✅' },
      { v:'ไม่อนุมัติ',   l:'ไม่อนุมัติ ❌' },
      { v:'ส่งกลับแก้ไข', l:'ส่งกลับแก้ไข' },
    ],
    admin: [
      { v:'รอดำเนินการ',       l:'รอดำเนินการ' },
      { v:'รอตรวจสอบ',     l:'📊 รอตรวจสอบ (ฝ่ายบัญชี)' },
      { v:'รอการอนุมัติ', l:'รอการอนุมัติ' },
      { v:'อนุมัติ',           l:'อนุมัติ' },
      { v:'ไม่อนุมัติ',       l:'ไม่อนุมัติ' },
      { v:'ส่งกลับแก้ไข',     l:'ส่งกลับแก้ไข' },
      { v:'กำลังซ่อม',        l:'กำลังซ่อม' },
      { v:'เสร็จสิ้น',        l:'เสร็จสิ้น' },
    ],
  };
  const opts = roleOptions[role] || roleOptions['admin'];
  if (!opts || opts.length === 0) {
    showToast('ไม่สามารถเปลี่ยนสถานะได้ในขั้นตอนนี้', 'error');
    return;
  }
  const sel = document.getElementById('new-status');
  sel.innerHTML = opts.map(o => `<option value="${o.v}">${o.l}</option>`).join('');

  document.getElementById('status-job-id').value = jobId;
  document.getElementById('new-status').value = j.status;
  document.getElementById('status-note').value = j.note || '';
  document.getElementById('actual-cost').value = '';
  document.getElementById('bill-preview').innerHTML = '<span class="material-icons" style="font-size:2rem;color:var(--gray-300);">receipt</span><div class="placeholder-text">คลิกเพื่อแนบบิล</div>';
  // reset estimate section
  const estPreview = document.getElementById('estimate-preview');
  if (estPreview) estPreview.innerHTML = '<span class="material-icons" style="font-size:2rem;color:var(--gray-300);">attach_money</span><div class="placeholder-text">คลิกเพื่อแนบใบประเมินราคา (ถ้ามี)</div>';
  const estInput = document.getElementById('estimate-image');
  if (estInput) estInput.value = '';
  toggleCompleteSection();
  showModal('modalStatus');
}

document.getElementById('new-status').addEventListener('change', toggleCompleteSection);
function toggleCompleteSection() {
  const sv = document.getElementById('new-status').value;
  // complete-section: แสดงเมื่อ เสร็จสิ้น
  document.getElementById('complete-section').style.display = sv === 'เสร็จสิ้น' ? 'block' : 'none';
  // estimate-section: แสดงเมื่อ รอตรวจสอบ (supervisor ส่งให้บัญชี)
  const estSec = document.getElementById('estimate-section');
  if (estSec) {
    const show = sv === 'รอตรวจสอบ' && ['supervisor','admin'].includes(currentUser?.role);
    estSec.style.display = show ? 'block' : 'none';
  }
}

async function submitStatusChange() {
  const jobId     = document.getElementById('status-job-id').value;
  const status    = document.getElementById('new-status').value;
  const note      = document.getElementById('status-note').value.trim();
  const cost      = document.getElementById('actual-cost').value.trim();
  const billFiles = document.getElementById('bill-image').files;
  const estInput  = document.getElementById('estimate-image');
  const estFiles  = estInput ? estInput.files : null;

  if (status === 'เสร็จสิ้น') {
    if (!cost || isNaN(cost)) { showToast('กรุณากรอกค่าใช้จ่ายจริง', 'error'); return; }
    if (!billFiles.length) { showToast('กรุณาแนบบิล/ใบเสร็จ', 'error'); return; }
  }

  showLoading(true);

  // Upload บิล/ใบเสร็จ (เสร็จสิ้น)
  let billUrl = ''; let billViewUrl = '';
  if (billFiles.length > 0) {
    try {
      const results = await uploadMultiImages(billFiles, 'bill', jobId);
      billUrl     = JSON.stringify(results.map(r => r.url).filter(Boolean));
      billViewUrl = results[0]?.viewUrl || '';
    } catch(e) {
      showLoading(false); showToast('อัปโหลดบิลไม่สำเร็จ: ' + e.message, 'error'); return;
    }
  }

  // Upload ใบประเมินราคา (supervisor ส่งให้บัญชี)
  let estimateImageUrl = ''; let estimateViewUrl = '';
  if (estFiles && estFiles.length > 0 && status === 'รอตรวจสอบ') {
    try {
      const estResults = await uploadMultiImages(estFiles, 'estimate', jobId);
      estimateImageUrl = JSON.stringify(estResults.map(r => r.url).filter(Boolean));
      estimateViewUrl  = estResults[0]?.viewUrl || '';
    } catch(e) {
      showLoading(false); showToast('อัปโหลดใบประเมินไม่สำเร็จ: ' + e.message, 'error'); return;
    }
  }

  const statusPayload = {
    jobId, status, actualCost: cost, billUrl, billViewUrl,
    estimateImageUrl, estimateViewUrl,
    adminUid: currentUser.lineUid,
  };
  if (currentUser.role === 'accountant') {
    statusPayload.accountantNote = note;
  } else {
    statusPayload.note = note;
  }

  const res = await gasCall('updateStatus', statusPayload);
  showLoading(false);
  if (res.status === 'ok') {
    showToast('อัปเดตสถานะสำเร็จ');
    closeModal('modalStatus');
    loadData();
  } else {
    showToast('เกิดข้อผิดพลาด: ' + (res.message || ''), 'error');
  }
}

/* Upload หลายรูปพร้อมกัน (parallel) — คืน array ของ {url, viewUrl} */
async function uploadMultiImages(files, folderType, jobId = '') {
  if (!files || files.length === 0) return [];
  const tasks = Array.from(files).map(async (file, i) => {
    const compressed = await compressImage(file);
    const res = await gasCall('uploadImage', {
      base64: compressed.b64,
      mimeType: compressed.mime,
      filename: compressed.filename,
      folderType,
      jobId,
    });
    if (res.status !== 'ok') throw new Error(res.message || 'upload failed');
    return { url: res.url || '', viewUrl: res.viewUrl || '', fileId: res.fileId || '' };
  });
  return Promise.all(tasks); // parallel ทั้งหมดพร้อมกัน
}

/* ============================================================
   ACCOUNTANT JOBS PAGE
   ============================================================ */
function filterAccountantJobs(status) {
  _purFilter = status;
  document.querySelectorAll('[data-pur-status]').forEach(el => {
    el.classList.toggle('active', el.dataset.purStatus === status);
  });
  renderAccountantJobs();
}

function renderAccountantJobs() {
  // stats
  const PUR_VISIBLE = ['รอตรวจสอบ','รอการอนุมัติ','ส่งกลับแก้ไข','อนุมัติ','กำลังซ่อม','เสร็จสิ้น','ไม่อนุมัติ'];
  const pending  = allJobs.filter(j => j.status === 'รอตรวจสอบ').length;
  const approved = allJobs.filter(j => ['รอการอนุมัติ','อนุมัติ','กำลังซ่อม','เสร็จสิ้น'].includes(j.status)).length;
  const done     = allJobs.filter(j => j.status === 'เสร็จสิ้น').length;
  document.getElementById('pur-stat-pending')?.textContent  !== undefined && (document.getElementById('pur-stat-pending').textContent  = pending);
  document.getElementById('pur-stat-approved')?.textContent !== undefined && (document.getElementById('pur-stat-approved').textContent = approved);
  document.getElementById('pur-stat-done')?.textContent     !== undefined && (document.getElementById('pur-stat-done').textContent     = done);

  // badge
  const badge = document.getElementById('badge-review');
  if (badge) { badge.textContent = pending; badge.style.display = pending ? 'inline' : 'none'; }

  // filter + search
  const q = (document.getElementById('pur-search')?.value || '').toLowerCase().trim();
  let jobs = allJobs.filter(j => PUR_VISIBLE.includes(j.status));
  if (_purFilter !== 'all') jobs = jobs.filter(j => j.status === _purFilter);
  if (q) jobs = jobs.filter(j =>
    (j.plate||'').toLowerCase().includes(q) ||
    (j.userName||'').toLowerCase().includes(q) ||
    (j.jobId||'').toLowerCase().includes(q) ||
    (j.detail||'').toLowerCase().includes(q)
  );
  jobs = [...jobs].sort((a,b) => {
    const sortVal = document.getElementById('acc-sort')?.value || 'oldest';
    const ta = new Date(a.createdAt).getTime();
    const tb = new Date(b.createdAt).getTime();
    return sortVal === 'oldest' ? ta - tb : tb - ta;
  });

  const tbody = document.getElementById('pur-jobs-tbody');
  const cards = document.getElementById('pur-jobs-cards');
  if (!tbody || !cards) return;

  if (!jobs.length) {
    tbody.innerHTML = `<tr><td colspan="7"><div style="display:flex;flex-direction:column;align-items:center;padding:2.5rem;color:var(--gray-400);">
      <span class="material-icons" style="font-size:2.5rem;">inbox</span><span style="font-size:.9rem;">ไม่พบรายการ</span></div></td></tr>`;
    cards.innerHTML = `<div style="display:flex;flex-direction:column;align-items:center;padding:2rem;color:var(--gray-400);">
      <span class="material-icons" style="font-size:2.5rem;">inbox</span><span>ไม่พบรายการ</span></div>`;
    return;
  }

  const actionBtns = j => {
    const viewBtn = `<button class="btn-outline-custom btn-sm" onclick="openDetail('${j.jobId}')" title="ดูรายละเอียด">
      <span class="material-icons" style="font-size:.9rem;">visibility</span></button>`;
    if (j.status !== 'รอตรวจสอบ') return viewBtn;
    return `<div class="d-flex gap-1">
      ${viewBtn}
      <button class="btn-primary-custom btn-sm" style="background:var(--success);padding:.3rem .6rem;" onclick="quickAccountantDecision('${j.jobId}','รอการอนุมัติ')" title="ผ่านการตรวจสอบ">
        <span class="material-icons" style="font-size:.9rem;">check</span>
      </button>
      <button class="btn-primary-custom btn-sm" style="background:var(--warning);color:#333;padding:.3rem .6rem;" onclick="quickAccountantDecision('${j.jobId}','ส่งกลับแก้ไข')" title="ส่งกลับแก้ไข">
        <span class="material-icons" style="font-size:.9rem;">undo</span>
      </button>
    </div>`;
  };

  // Desktop
  tbody.innerHTML = jobs.map(j => `
    <tr style="${j.status==='รอตรวจสอบ'?'background:rgba(255,193,7,.05);':''}">
      <td style="padding:.75rem .875rem;white-space:nowrap;"><span style="font-weight:700;color:var(--primary-dark);font-size:.82rem;">${j.jobId}</span></td>
      <td style="padding:.75rem .875rem;font-size:.85rem;">${j.userName}</td>
      <td style="padding:.75rem .875rem;"><span style="font-weight:600;font-size:.88rem;">${j.plate}</span></td>
      <td style="padding:.75rem .875rem;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:.82rem;">${j.detail}</td>
      <td style="padding:.75rem .875rem;text-align:right;font-weight:600;font-size:.88rem;white-space:nowrap;">${j.estimate ? Number(j.estimate).toLocaleString()+' ฿' : '-'}</td>
      <td style="padding:.75rem .875rem;"><span class="badge-status ${statusClass(j.status)}">${j.status}</span></td>
      <td style="padding:.75rem .875rem;text-align:center;">${actionBtns(j)}</td>
    </tr>`).join('');

  // Mobile
  cards.innerHTML = jobs.map(j => `
    <div class="repair-card" onclick="openDetail('${j.jobId}')">
      <div class="d-flex justify-content-between align-items-start mb-2">
        <span style="font-weight:700;font-size:.8rem;color:var(--primary-dark);">${j.jobId}</span>
        <span class="badge-status ${statusClass(j.status)}">${j.status}</span>
      </div>
      <div style="font-weight:600;margin-bottom:.25rem;">${j.plate}</div>
      <div style="font-size:.82rem;color:var(--gray-600);margin-bottom:.5rem;">${j.detail}</div>
      <div class="d-flex justify-content-between align-items-center">
        <span style="font-size:.78rem;color:var(--gray-500);">${j.userName} • ${formatDate(j.createdAt)}</span>
        ${j.status === 'รอตรวจสอบ' ? `<div class="d-flex gap-1" onclick="event.stopPropagation()">
          <button class="btn-primary-custom btn-sm" style="background:var(--success);padding:.3rem .6rem;" onclick="quickAccountantDecision('${j.jobId}','รอการอนุมัติ')"><span class="material-icons" style="font-size:.9rem;">check</span></button>
          <button class="btn-primary-custom btn-sm" style="background:var(--warning);color:#333;padding:.3rem .6rem;" onclick="quickAccountantDecision('${j.jobId}','ส่งกลับแก้ไข')"><span class="material-icons" style="font-size:.9rem;">undo</span></button>
        </div>` : ''}
      </div>
    </div>`).join('');
}

function quickAccountantDecision(jobId, decision) {
  if (decision === 'ส่งกลับแก้ไข') {
    document.getElementById('mgr-note-jobid').value    = jobId;
    document.getElementById('mgr-note-decision').value = decision;
    document.getElementById('mgr-note-text').value     = '';
    document.getElementById('mgr-note-title').textContent = '↩️ ส่งกลับให้หัวหน้าแก้ไข';
    const btn = document.getElementById('mgr-note-confirm-btn');
    btn.style.background = 'var(--warning)'; btn.style.color = '#333';
    showModal('modalManagerNote');
  } else {
    _submitAccountantDecision(jobId, decision, '');
  }
}

async function _submitAccountantDecision(jobId, decision, note) {
  showLoading(true);
  const res = await gasCall('updateStatus', {
    jobId, status: decision, note: note || '',
    adminUid: currentUser.lineUid,
  });
  showLoading(false);
  if (res.status === 'ok') {
    showToast(decision === 'รอการอนุมัติ' ? '✅ ส่งให้ผู้บริหารแล้ว' : '↩️ ส่งกลับให้หัวหน้าแล้ว');
    sessionStorage.removeItem('_repairData');
    await loadData(true);
    renderAccountantJobs();
  } else {
    showToast('เกิดข้อผิดพลาด: ' + (res.message || ''), 'error');
  }
}

/* ============================================================
   USER MANAGEMENT
   ============================================================ */
/* ซ่อน/ล็อค option admin สำหรับ supervisor */
function _syncAdminRoleOption(targetUser) {
  const callerRole    = currentUser?.role;
  const isCallerAdmin = callerRole === 'admin';
  const isCallerSupervisor = ['admin','supervisor','manager'].includes(callerRole);
  const isCallerManager = ['admin','manager'].includes(callerRole);
  const sel = document.getElementById('u-role');
  const adminOpt      = sel ? sel.querySelector('option[value="admin"]') : null;
  const accountantOpt = sel ? sel.querySelector('option[value="accountant"]') : null;
  if (!adminOpt) return;

  // admin option: admin เท่านั้น
  adminOpt.disabled     = !isCallerAdmin;
  adminOpt.style.display = isCallerAdmin ? '' : 'none';

  // accountant option: supervisor+ ทำได้
  if (accountantOpt) {
    accountantOpt.disabled     = !isCallerSupervisor;
    accountantOpt.style.display = isCallerSupervisor ? '' : 'none';
  }

  // lock dropdown เฉพาะ admin (supervisor เปลี่ยน role accountant ได้ทุก role)
  if (targetUser) {
    const blocked = (!isCallerAdmin && targetUser.role === 'admin');
    sel.disabled = blocked;
  } else {
    sel.disabled = false;
  }
}

function openAddUserModal() {
  document.getElementById('u-row-id').value='';
  document.getElementById('u-line-id').value='';
  document.getElementById('u-name').value='';
  document.getElementById('u-dept').value='';
  document.getElementById('u-role').value='user';
  document.getElementById('u-line-id').disabled=false;
  _syncAdminRoleOption(null);
  showModal('modalAddUser');
}
function openEditUser(lineUid) {
  const u = users.find(x => x.lineUid===lineUid);
  if (!u) return;
  document.getElementById('u-row-id').value = lineUid;
  document.getElementById('u-line-id').value = u.lineUid;
  document.getElementById('u-name').value = u.name;
  document.getElementById('u-dept').value = u.dept||'';
  document.getElementById('u-role').value = u.role||'user';
  document.getElementById('u-line-id').disabled = true;
  _syncAdminRoleOption(u);
  showModal('modalAddUser');
}
async function saveUser() {
  const lineUid = document.getElementById('u-line-id').value.trim();
  const name    = document.getElementById('u-name').value.trim();
  const dept    = document.getElementById('u-dept').value.trim();
  const role    = document.getElementById('u-role').value;
  if (!lineUid || !name) { showToast('กรุณากรอก LINE UID และชื่อ', 'error'); return; }
  showLoading(true);
  const isEdit = !!document.getElementById('u-row-id').value;
  const res = await gasCall(isEdit ? 'updateUser' : 'addUser', { lineUid, name, dept, role, adminUid: currentUser.lineUid });
  showLoading(false);
  if (res.status==='ok') { showToast('บันทึกสำเร็จ'); closeModal('modalAddUser'); loadData(); }
  else showToast('ผิดพลาด: '+(res.message||''), 'error');
}
async function deleteUser(lineUid, name) {
  if (!confirm(`ลบผู้ใช้ "${name}" ออกจากระบบ?`)) return;
  showLoading(true);
  const res = await gasCall('deleteUser', { lineUid, adminUid: currentUser.lineUid });
  showLoading(false);
  if (res.status==='ok') { showToast('ลบสำเร็จ'); loadData(); }
  else showToast('ผิดพลาด', 'error');
}

/* ============================================================
   VEHICLE MANAGEMENT
   ============================================================ */
let _editingPlate = null; // เก็บ plate เดิมตอนแก้ไข

function openAddVehicleModal() {
  _editingPlate = null;
  document.getElementById('v-modal-title').innerHTML = '<span class="material-icons">directions_car</span> เพิ่มทะเบียนรถ';
  document.getElementById('v-plate').value = '';
  document.getElementById('v-plate').disabled = false;
  document.getElementById('v-model').value = '';
  showModal('modalAddVehicle');
}

function openEditVehicle(el) {
  const plate = el.dataset.plate;
  const model = el.dataset.model || '';
  _editingPlate = plate;
  document.getElementById('v-modal-title').innerHTML = '<span class="material-icons">edit</span> แก้ไขทะเบียนรถ';
  document.getElementById('v-plate').value = plate;
  document.getElementById('v-plate').disabled = true; // ป้องกันแก้ plate (ใช้เป็น key)
  document.getElementById('v-model').value = model;
  showModal('modalAddVehicle');
}
async function saveVehicle() {
  const plate = document.getElementById('v-plate').value.trim();
  const model = document.getElementById('v-model').value.trim();
  if (!plate) { showToast('กรุณากรอกเลขทะเบียน', 'error'); return; }
  showLoading(true);
  let res;
  if (_editingPlate) {
    // แก้ไข — ลบเก่าแล้วเพิ่มใหม่ (plate เดิม model ใหม่)
    res = await gasCall('updateVehicle', { plate: _editingPlate, model, adminUid: currentUser.lineUid });
  } else {
    res = await gasCall('addVehicle', { plate, model, adminUid: currentUser.lineUid });
  }
  showLoading(false);
  if (res.status==='ok') {
    showToast(_editingPlate ? 'แก้ไขทะเบียนสำเร็จ' : 'เพิ่มทะเบียนสำเร็จ');
    closeModal('modalAddVehicle');
    loadData();
  } else {
    showToast('ผิดพลาด: '+(res.message||''), 'error');
  }
}
async function deleteVehicle(plate) {
  if (!confirm(`ลบทะเบียน "${plate}"?`)) return;
  showLoading(true);
  const res = await gasCall('deleteVehicle', { plate, adminUid: currentUser.lineUid });
  showLoading(false);
  if (res.status==='ok') { showToast('ลบสำเร็จ'); loadData(); }
  else showToast('ผิดพลาด', 'error');
}

/* Archive งานเก่า — เรียกจากปุ่ม admin */
async function runArchive() {
  const months = prompt('Archive งานที่ปิดแล้ว (เสร็จสิ้น/ไม่อนุมัติ) เก่ากว่ากี่เดือน?', '12');
  if (!months || isNaN(months)) return;
  if (!confirm(`Archive งานเก่ากว่า ${months} เดือน ?
งานเหล่านี้จะย้ายไป Sheet "Jobs_Archive" และไม่แสดงในหน้าหลัก`)) return;
  showLoading(true);
  const res = await gasCall('archiveOldJobs', { adminUid: currentUser.lineUid, monthsOld: parseInt(months) });
  showLoading(false);
  if (res.status === 'ok') {
    showToast(res.archived > 0 ? `✅ Archive ${res.archived} รายการสำเร็จ` : 'ไม่มีงานที่ต้อง archive');
    if (res.archived > 0) loadData();
  } else {
    showToast('Archive ไม่สำเร็จ: ' + (res.message || ''), 'error');
  }
}

/* ============================================================
   GAS CALL
   ใช้ XMLHttpRequest แทน fetch เพราะ LIFF บน LINE app
   บล็อก fetch ที่มี redirect (GAS redirect ทุก request)
   XHR จัดการ redirect ได้อัตโนมัติและทำงานได้ใน LIFF
   ============================================================ */
function gasCall(action, params = {}) {
  return new Promise((resolve) => {
    try {
      const hasBase64 = !!params.base64;
      const xhr = new XMLHttpRequest();
      xhr.withCredentials = false;
      xhr.timeout = hasBase64 ? 120000 : 30000;

      xhr.onload = function () {
        try {
          // GAS อาจ redirect → responseURL จะเป็น URL ใหม่ แต่ responseText คือ JSON
          const text = xhr.responseText;
          if (!text || text.trim() === '') {
            resolve({ status: 'error', message: 'Empty response from GAS' });
            return;
          }
          // หา JSON ใน response (บางครั้ง GAS แนบ HTML ไว้ด้านหน้า)
          const jsonStart = text.indexOf('{');
          const jsonEnd   = text.lastIndexOf('}');
          if (jsonStart === -1) {
            console.error('No JSON in response:', text.substring(0, 300));
            resolve({ status: 'error', message: 'No JSON response' });
            return;
          }
          const json = JSON.parse(text.substring(jsonStart, jsonEnd + 1));
          resolve(json);
        } catch (e) {
          console.error('GAS parse error:', xhr.responseText.substring(0, 300));
          resolve({ status: 'error', message: 'Parse error: ' + e.message });
        }
      };
      xhr.onerror   = () => resolve({ status: 'error', message: 'Network error — ตรวจสอบ GAS_URL' });
      xhr.ontimeout = () => resolve({ status: 'error', message: 'Request timeout (รูปอาจใหญ่เกินไป)' });

      if (hasBase64) {
        // upload รูป: POST พร้อม payload= urlencoded
        // GAS รับได้ผ่าน e.parameter.payload หรือ e.postData.contents
        const payloadStr = JSON.stringify({ action, ...params, _secret: CONFIG.API_SECRET || '' });
        xhr.open('POST', CONFIG.GAS_URL, true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        xhr.send('payload=' + encodeURIComponent(payloadStr));
      } else {
        // request ปกติ: GET query string
        const url = CONFIG.GAS_URL + '?payload=' + encodeURIComponent(JSON.stringify({ action, ...params, _secret: CONFIG.API_SECRET || '' }));
        xhr.open('GET', url, true);
        xhr.send();
      }
    } catch (e) {
      console.error('gasCall exception:', e);
      resolve({ status: 'error', message: e.message });
    }
  });
}

/* ============================================================
   HELPERS
   ============================================================ */
function showModal(id) {
  const el = document.getElementById(id);
  if (!el) return;
  if (el.classList.contains('modal-overlay')) {
    el.style.display = 'flex';
  } else {
    // Bootstrap modal
    const old = bootstrap.Modal.getInstance(el);
    if (old) old.dispose();
    new bootstrap.Modal(el).show();
  }
}
function closeModal(id) {
  const el = document.getElementById(id);
  if (!el) return;
  if (el.classList.contains('modal-overlay')) {
    el.style.display = 'none';
  } else {
    const inst = bootstrap.Modal.getInstance(el);
    if (inst) inst.hide();
    document.querySelectorAll('.modal-backdrop').forEach(b => b.remove());
    document.body.classList.remove('modal-open');
    document.body.style.overflow = '';
    document.body.style.paddingRight = '';
  }
}
function showLoading(v) { document.getElementById('loading-overlay').classList.toggle('show', v); }

function showToast(msg, type='success') {
  const c = document.getElementById('toast-container');
  const el = document.createElement('div');
  el.className = `toast-msg${type==='error'?' error':type==='warning'?' warning':''}`;
  el.innerHTML = `<span class="material-icons" style="font-size:1.1rem;color:${type==='error'?'var(--danger)':type==='warning'?'var(--warning)':'var(--primary)'};">${type==='error'?'error':type==='warning'?'warning':'check_circle'}</span>${msg}`;
  c.appendChild(el);
  setTimeout(() => el.remove(), 3500);
}

function statusClass(s) {
  return { 'รอดำเนินการ':'status-waiting','รอตรวจสอบ':'status-review','รอการอนุมัติ':'status-pending-mgr','อนุมัติ':'status-approved','ไม่อนุมัติ':'status-rejected','ส่งกลับแก้ไข':'status-revise','กำลังซ่อม':'status-repairing','เสร็จสิ้น':'status-done' }[s] || '';
}

function formatDate(d) {
  if (!d) return '-';
  const t = new Date(d);
  return isNaN(t) ? d : t.toLocaleString('th-TH', { day:'2-digit', month:'short', year:'2-digit', hour:'2-digit', minute:'2-digit' });
}

function animateNum(id, val) {
  const el = document.getElementById(id);
  let start = parseInt(el.textContent)||0; const end=val; const dur=500;
  const step = (ts) => {
    if (!start.startTime) start.startTime=ts;
    const prog = Math.min((ts-start.startTime)/dur,1);
    el.textContent = Math.floor(prog*(end-parseInt(el.textContent||0))+parseInt(el.textContent||0));
    if (prog<1) requestAnimationFrame(step); else el.textContent=end;
  };
  el.textContent=start; requestAnimationFrame(ts=>{ start.startTime=ts; el.textContent=Math.round((ts-start.startTime)/dur*(end-start)+start); requestAnimationFrame(step); });
  el.textContent = val; // fallback
}

function previewImage(input, previewId) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    document.getElementById(previewId).innerHTML = `<img src="${e.target.result}" style="max-height:160px;border-radius:8px;">`;
  };
  reader.readAsDataURL(file);
}

/* Multi-image preview — แสดงทุกรูปเป็น thumbnails */
function previewMultiImage(input, previewId) {
  const files = Array.from(input.files);
  if (!files.length) return;
  const box = document.getElementById(previewId);
  box.innerHTML = '';
  files.forEach((file, idx) => {
    const reader = new FileReader();
    reader.onload = e => {
      const wrap = document.createElement('div');
      wrap.style.cssText = 'position:relative;display:inline-block;';
      wrap.innerHTML = `<img src="${e.target.result}" style="height:80px;width:80px;object-fit:cover;border-radius:6px;border:2px solid var(--primary-xlight);">
        <span style="position:absolute;top:2px;right:2px;background:rgba(0,0,0,.45);color:#fff;border-radius:50%;width:18px;height:18px;font-size:11px;display:flex;align-items:center;justify-content:center;cursor:pointer;" onclick="removePreviewImage(this,'${input.id}',${idx})">&times;</span>`;
      box.appendChild(wrap);
    };
    reader.readAsDataURL(file);
  });
  const hint = document.createElement('div');
  hint.style.cssText = 'font-size:.78rem;color:var(--gray-500);width:100%;text-align:center;margin-top:4px;';
  hint.textContent = `${files.length} รูป — คลิกเพื่อเพิ่ม`;
  box.appendChild(hint);
}

/* ลบรูปออกจาก preview (สร้าง FileList ใหม่) */
function removePreviewImage(btn, inputId, removeIdx) {
  const input = document.getElementById(inputId);
  const dt = new DataTransfer();
  Array.from(input.files).forEach((f, i) => { if (i !== removeIdx) dt.items.add(f); });
  input.files = dt.files;
  previewMultiImage(input, input.id === 'f-image' ? 'repair-img-preview' : 'bill-preview');
}

function viewImage(url) {
  document.getElementById('modal-img-src').src = url;
  showModal('modalImage');
}

/* parse imageUrl / billUrl ที่อาจเป็น JSON array หรือ string เดิม */
function parseUrls(raw) {
  if (!raw) return [];
  try { const a = JSON.parse(raw); return Array.isArray(a) ? a.filter(Boolean) : [raw]; }
  catch { return [raw]; }
}
/* สร้าง thumbnail HTML สำหรับแสดงหลายรูป */
function renderThumbRow(urls, label, viewUrl = '') {
  if (!urls.length) return '';
  const thumbs = urls.map(u =>
    `<img src="${u}" class="bill-thumb" onclick="viewImage('${u}')" onerror="this.style.display='none'" style="height:64px;width:64px;object-fit:cover;border-radius:6px;cursor:pointer;border:1px solid #ddd;">`
  ).join('');
  const driveLink = viewUrl
    ? `<a href="${viewUrl}" target="_blank" class="btn-outline-custom btn-sm" style="font-size:.8rem;"><span class="material-icons" style="font-size:.9rem;">open_in_new</span> Drive</a>`
    : '';
  return `<div class="mt-2"><div class="form-label mb-1">${label}</div><div class="d-flex align-items-center gap-2 flex-wrap">${thumbs}${driveLink}</div></div>`;
}

/* Compress รูปก่อน upload — resize ให้ไม่เกิน maxW/maxH และ quality 0.8 */
function compressImage(file, maxW = 800, maxH = 800, quality = 0.72) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = reject;
    reader.onload = e => {
      const img = new Image();
      img.onerror = reject;
      img.onload = () => {
        let w = img.width, h = img.height;
        if (w > maxW || h > maxH) {
          const ratio = Math.min(maxW / w, maxH / h);
          w = Math.round(w * ratio);
          h = Math.round(h * ratio);
        }
        const canvas = document.createElement('canvas');
        canvas.width = w; canvas.height = h;
        canvas.getContext('2d').drawImage(img, 0, 0, w, h);
        const mime = 'image/jpeg'; // force JPEG — PNG ใหญ่กว่า 3-5x ไม่จำเป็นสำหรับรูปซ่อมรถ
        const dataUrl = canvas.toDataURL(mime, quality);
        const b64 = dataUrl.split(',')[1];
        const kb  = Math.round(b64.length * 0.75 / 1024);
        console.log(`Compressed: ${img.width}x${img.height} → ${w}x${h} | ~${kb} KB`);
        resolve({ b64, mime, filename: file.name });
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

function toBase64(file) {
  return new Promise((res,rej) => {
    const r = new FileReader();
    r.onload = ()=>res(r.result.split(',')[1]);
    r.onerror = rej;
    r.readAsDataURL(file);
  });
}

/* ============================================================
   DELETE JOB (Admin only)
   ============================================================ */
function confirmDeleteJob(jobId, plate) {
  // Modal ยืนยันการลบ
  const existingModal = document.getElementById('modalConfirmDelete');
  if (existingModal) existingModal.remove();

  const modal = document.createElement('div');
  modal.id = 'modalConfirmDelete';
  modal.className = 'modal fade';
  modal.setAttribute('tabindex', '-1');
  modal.innerHTML = `
    <div class="modal-dialog modal-dialog-centered">
      <div class="modal-content">
        <div class="modal-header-custom" style="background:linear-gradient(135deg,#C62828,#B71C1C);">
          <div class="modal-title"><span class="material-icons">warning</span> ยืนยันการลบ</div>
          <button class="btn-close-white" onclick="bootstrap.Modal.getInstance(document.getElementById('modalConfirmDelete')).hide()">
            <span class="material-icons" style="font-size:1rem;">close</span>
          </button>
        </div>
        <div class="modal-body text-center py-4">
          <span class="material-icons" style="font-size:3rem;color:var(--danger);display:block;margin-bottom:.75rem;">delete_forever</span>
          <div style="font-size:1rem;font-weight:600;margin-bottom:.4rem;">ลบรายการแจ้งซ่อม</div>
          <div style="font-size:.88rem;color:var(--gray-600);">เลขที่งาน <strong>${jobId}</strong> (${plate})</div>
          <div style="font-size:.82rem;color:var(--danger);margin-top:.75rem;background:var(--danger-light);border-radius:8px;padding:.5rem .75rem;">
            ⚠️ ไม่สามารถกู้คืนได้หลังจากลบแล้ว
          </div>
        </div>
        <div class="modal-footer justify-content-center gap-2">
          <button class="btn-outline-custom" onclick="bootstrap.Modal.getInstance(document.getElementById('modalConfirmDelete')).hide()">ยกเลิก</button>
          <button class="btn-primary-custom" style="background:linear-gradient(135deg,#C62828,#B71C1C);" onclick="executeDeleteJob('${jobId}')">
            <span class="material-icons" style="font-size:1rem;">delete</span> ยืนยันลบ
          </button>
        </div>
      </div>
    </div>`;

  document.body.appendChild(modal);
  const bsModal = new bootstrap.Modal(modal);
  bsModal.show();
  modal.addEventListener('hidden.bs.modal', () => modal.remove());
}

async function executeDeleteJob(jobId) {
  bootstrap.Modal.getInstance(document.getElementById('modalConfirmDelete'))?.hide();
  showLoading(true);
  const res = await gasCall('deleteJob', { jobId, adminUid: currentUser.lineUid });
  showLoading(false);
  if (res.status === 'ok') {
    showToast('ลบรายการ ' + jobId + ' สำเร็จ');
    loadData();
  } else {
    showToast('ลบไม่สำเร็จ: ' + (res.message || ''), 'error');
  }
}

/* ============================================================
   EXPORT CSV
   ============================================================ */
async function exportCSV() {
  const btn = document.getElementById('btn-export-csv');
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="material-icons" style="font-size:1rem;">hourglass_empty</span> กำลัง Export...'; }
  showLoading(true);

  const isAdmin = ['admin','supervisor','accountant','manager'].includes(currentUser.role);
  const res = await gasCall('exportJobs', {
    lineUid: currentUser.lineUid,
    isAdmin,
    filterStatus,
  });
  showLoading(false);
  if (btn) { btn.disabled = false; btn.innerHTML = '<span class="material-icons" style="font-size:1rem;">download</span> Export CSV'; }

  if (res.status !== 'ok') { showToast('Export ไม่สำเร็จ: ' + res.message, 'error'); return; }

  // สร้าง CSV
  const BOM    = '\uFEFF'; // UTF-8 BOM สำหรับ Excel ภาษาไทย
  const escape = v => '"' + String(v).replace(/"/g, '""') + '"';
  const lines  = [
    res.headers.map(escape).join(','),
    ...res.rows.map(row => row.map(escape).join(','))
  ];
  const csv  = BOM + lines.join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  const date = new Date().toLocaleDateString('th-TH').replace(/\//g,'-');
  a.href     = url;
  a.download = `แจ้งซ่อมรถ_${date}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  showToast(`Export สำเร็จ ${res.total} รายการ ✅`);
}

/* ============================================================
   PRINT / EXPORT PDF — ใบรายงานแจ้งซ่อมรถยนต์ (1 หน้า A4)
   ============================================================ */
/* ============================================================
   PDF HELPER — สร้าง PDF จาก HTML แล้ว download (ใช้ได้บนมือถือ)
   ============================================================ */
async function savePDF(htmlContent, filename) {
  await ensurePdfLibs();
  // สร้าง hidden iframe render HTML
  const iframe = document.createElement('iframe');
  iframe.style.cssText = 'position:fixed;left:-9999px;top:-9999px;width:794px;height:1123px;border:none;';
  document.body.appendChild(iframe);

  const iDoc = iframe.contentDocument || iframe.contentWindow.document;
  iDoc.open(); iDoc.write(htmlContent); iDoc.close();

  // รอ font/image โหลด
  await new Promise(r => setTimeout(r, 1200));

  try {
    const canvas = await html2canvas(iDoc.body, {
      scale: 1.8,
      useCORS: true,
      allowTaint: true,
      backgroundColor: '#ffffff',
      width: 794,
      windowWidth: 794,
    });

    const imgData = canvas.toDataURL('image/jpeg', 0.92);
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ orientation:'portrait', unit:'mm', format:'a4' });
    const pW = pdf.internal.pageSize.getWidth();
    const pH = pdf.internal.pageSize.getHeight();
    // คำนวณ height ตามสัดส่วน
    const imgH = (canvas.height / canvas.width) * pW;
    if (imgH <= pH) {
      pdf.addImage(imgData, 'JPEG', 0, 0, pW, imgH);
    } else {
      // กรณียาวเกิน 1 หน้า — แบ่งหน้า
      let yOffset = 0;
      while (yOffset < imgH) {
        if (yOffset > 0) pdf.addPage();
        pdf.addImage(imgData, 'JPEG', 0, -yOffset, pW, imgH);
        yOffset += pH;
      }
    }
    pdf.save(filename);
    const isIOS = /iPhone|iPad|iPod/i.test(navigator.userAgent);
    if (isIOS) {
      showToast('PDF เปิดใน browser → กด Share → Save to Files 📁');
    } else {
      showToast('PDF บันทึกลง Downloads แล้ว ✅');
    }
  } catch(e) {
    console.error('PDF error:', e);
    showToast('ไม่สามารถสร้าง PDF ได้: ' + e.message, 'error');
  } finally {
    document.body.removeChild(iframe);
  }
}

async function printJobPDF(jobId) {
  const j = allJobs.find(x => x.jobId === jobId);
  if (!j) return;

  const statusColorMap = {
    'รอดำเนินการ': '#F57F17', 'อนุมัติ': '#2E7D32', 'กำลังซ่อม': '#1565C0',
    'เสร็จสิ้น': '#1B5E20', 'ส่งกลับแก้ไข': '#6A1B9A', 'ไม่อนุมัติ': '#B71C1C', 'รอการอนุมัติ': '#E65100',
  };
  const statusBgMap = {
    'รอดำเนินการ': '#FFF8E1', 'อนุมัติ': '#FFEBEE', 'กำลังซ่อม': '#E3F2FD',
    'เสร็จสิ้น': '#E8F5E9', 'ส่งกลับแก้ไข': '#F3E5F5', 'ไม่อนุมัติ': '#FFEBEE', 'รอการอนุมัติ': '#FFF3E0',
  };
  const sc = statusColorMap[j.status] || '#333';
  const sb = statusBgMap[j.status]   || '#f5f5f5';

  const printDate = new Date().toLocaleDateString('th-TH', {
    year: 'numeric', month: 'long', day: 'numeric',
    hour: '2-digit', minute: '2-digit'
  });
  const jobDate = j.createdAt ? new Date(j.createdAt).toLocaleDateString('th-TH', {
    year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit'
  }) : '-';
  const updateDate = j.updatedAt ? new Date(j.updatedAt).toLocaleDateString('th-TH', {
    year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit'
  }) : '-';

  const row = (label, value, highlight) => `
    <tr>
      <td style="padding:8px 14px;font-weight:600;color:#555;background:#fafafa;width:36%;border:1px solid #e0e0e0;font-size:13px;">${label}</td>
      <td style="padding:8px 14px;color:${highlight||'#222'};background:#fff;border:1px solid #e0e0e0;font-size:13px;font-weight:${highlight?'700':'400'};">${value||'-'}</td>
    </tr>`;

  const logoSrc = document.querySelector('#login-logo-wrap img') ? document.querySelector('#login-logo-wrap img').src : '';
  const logoHtml = logoSrc && !logoSrc.includes('placeholder')
    ? `<img src="${logoSrc}" style="height:60px;object-fit:contain;">`
    : `<div style="width:60px;height:60px;background:#B71C1C;border-radius:12px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:28px;">🚗</div>`;

  const html = `\u003c!DOCTYPE html>
\x3Chtml lang="th">
<head>
<meta charset="UTF-8">
<title>ใบรายงานแจ้งซ่อม ${j.jobId}</title>
<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&family=Prompt:wght@600;700&display=swap" rel="stylesheet">
<style>
  /* ===== บังคับ 1 หน้า A4 ===== */
  @page {
    size: A4 portrait;
    margin: 0;
  }
  * { box-sizing:border-box; margin:0; padding:0; }
  html, body { width:210mm; height:297mm; overflow:hidden; }
  body {
    font-family:'Sarabun',sans-serif;
    font-size:11px;
    color:#222;
    background:#fff;
  }

  /* ===== LAYOUT WRAPPER ===== */
  .page {
    width:210mm;
    height:297mm;
    display:flex;
    flex-direction:column;
    background:#fff;
    overflow:hidden;
  }

  /* ===== HEADER ===== */
  .hd {
    background:linear-gradient(135deg,#7F0000,#B71C1C,#C62828);
    color:#fff;
    padding:10px 20px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    flex-shrink:0;
  }
  .hd-left { display:flex; align-items:center; gap:10px; }
  .hd-left img, .hd-left .logo-fallback {
    height:42px; width:42px; object-fit:contain;
    border-radius:8px;
  }
  .hd-title h1 { font-family:'Prompt',sans-serif; font-size:14px; font-weight:700; margin:0; }
  .hd-title p  { font-size:9.5px; opacity:.85; margin:1px 0 0; }
  .hd-right { text-align:right; font-size:9px; opacity:.9; line-height:1.65; }

  /* ===== STATUS BAR ===== */
  .sb {
    background:${sb};
    border-left:5px solid ${sc};
    padding:6px 20px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    flex-shrink:0;
  }
  .sb-id   { font-family:'Prompt',sans-serif; font-size:16px; font-weight:700; color:${sc}; }
  .sb-badge {
    background:${sc}; color:#fff;
    padding:3px 14px; border-radius:20px;
    font-family:'Prompt',sans-serif; font-size:11px; font-weight:700;
  }

  /* ===== BODY ===== */
  .body {
    flex:1;
    padding:10px 20px;
    display:flex;
    flex-direction:column;
    gap:8px;
    overflow:hidden;
  }

  /* 2-column grid สำหรับข้อมูล */
  .grid2 { display:grid; grid-template-columns:1fr 1fr; gap:8px; }
  .grid1 { display:grid; grid-template-columns:1fr; }

  .sec-title {
    font-family:'Prompt',sans-serif;
    font-size:9px; font-weight:700;
    color:#B71C1C;
    text-transform:uppercase;
    letter-spacing:.4px;
    border-bottom:1.5px solid #FFCDD2;
    padding-bottom:3px;
    margin-bottom:5px;
  }
  table { width:100%; border-collapse:collapse; }
  td {
    padding:4px 8px;
    font-size:10.5px;
    border:1px solid #e8e8e8;
    vertical-align:top;
  }
  td.lbl {
    width:38%; font-weight:600; color:#555;
    background:#fafafa;
  }
  td.val { color:#111; }

  /* detail text — ห่อบรรทัดได้ */
  .detail-text {
    background:#F8F9FA;
    border:1px solid #e0e0e0;
    border-radius:5px;
    padding:5px 8px;
    font-size:10.5px;
    line-height:1.5;
    min-height:32px;
  }

  /* note */
  .note-box {
    background:#FFFDE7;
    border-left:3px solid #F9A825;
    border-radius:4px;
    padding:4px 8px;
    font-size:10px;
  }

  /* cost */
  .cost-box {
    background:linear-gradient(135deg,#FFEBEE,#F1F8E9);
    border:1.5px solid #E53935;
    border-radius:6px;
    padding:6px 12px;
    display:flex;
    align-items:center;
    justify-content:space-between;
  }
  .cost-amount { font-family:'Prompt',sans-serif; font-size:20px; font-weight:700; color:#7F0000; }

  /* images */
  .img-row { display:flex; gap:8px; flex-wrap:wrap; }
  .img-box { border:1px solid #ddd; border-radius:5px; overflow:hidden; width:100%; }
  .img-box .img-lbl { padding:4px 10px; font-size:10px; font-weight:700; }
  .img-box img { width:100%; max-height:220px; object-fit:contain; display:block; background:#f9f9f9; }

  /* sign */
  .sign-row { display:flex; gap:8px; margin-top:2px; }
  .sign-box {
    flex:1; border:1px solid #ddd; border-radius:5px;
    padding:6px 8px; text-align:center;
  }
  .sign-line { border-bottom:1px dashed #aaa; height:30px; margin-bottom:4px; }
  .sign-lbl { font-size:9px; color:#666; line-height:1.3; }

  /* ===== FOOTER ===== */
  .ft {
    padding:5px 20px;
    background:#fafafa;
    border-top:1px solid #e0e0e0;
    display:flex;
    justify-content:space-between;
    font-size:9px;
    color:#999;
    flex-shrink:0;
  }

  @media print {
    html, body { width:210mm; height:297mm; }
    .page { page-break-after:avoid; }
  }
</style>
</head>
<body>
<div class="page">

  <!-- HEADER -->
  <div class="hd">
    <div class="hd-left">
      ${logoHtml}
      <div class="hd-title">
        <h1>ใบรายงานแจ้งซ่อมรถยนต์</h1>
        <p>Vehicle Maintenance Report</p>
      </div>
    </div>
    <div class="hd-right">
      <div>พิมพ์: ${printDate}</div>
      <div>โดย: ${currentUser.name}</div>
    </div>
  </div>

  <!-- STATUS BAR -->
  <div class="sb">
    <div class="sb-id">${j.jobId}</div>
    <div class="sb-badge">${j.status}</div>
  </div>

  <!-- BODY -->
  <div class="body">

    <!-- แถว 1: ข้อมูลแจ้งซ่อม + ข้อมูลรถ (2 คอลัมน์) -->
    <div class="grid2">

      <!-- ซ้าย: ข้อมูลแจ้งซ่อม -->
      <div>
        <div class="sec-title">📋 ข้อมูลการแจ้งซ่อม</div>
        <table>
          <tr><td class="lbl">เลขที่งาน</td><td class="val" style="font-weight:700;color:#7F0000;">${j.jobId}</td></tr>
          <tr><td class="lbl">ผู้แจ้งซ่อม</td><td class="val">${j.userName||'-'}</td></tr>
          <tr><td class="lbl">วันที่แจ้ง</td><td class="val">${jobDate}</td></tr>
          <tr><td class="lbl">อัปเดตล่าสุด</td><td class="val">${updateDate}</td></tr>
        </table>
      </div>

      <!-- ขวา: ข้อมูลรถ -->
      <div>
        <div class="sec-title">🚙 ข้อมูลรถยนต์</div>
        <table>
          <tr><td class="lbl">เลขทะเบียน</td><td class="val" style="font-size:13px;font-weight:700;">${j.plate}</td></tr>
          <tr><td class="lbl">เลขไมล์</td><td class="val">${Number(j.mileage).toLocaleString()} กม.</td></tr>
          <tr><td class="lbl">สถานที่ซ่อม</td><td class="val">${j.location||'-'}</td></tr>
          <tr><td class="lbl">ราคาประเมิน</td><td class="val">${j.estimate ? Number(j.estimate).toLocaleString()+' บาท' : '-'}</td></tr>
          ${j.estimateViewUrl ? `<tr><td class="lbl">ใบประเมิน</td><td class="val"><a href="${j.estimateViewUrl}" style="color:#2E7D32;font-size:9px;">📄 ดูใบประเมินราคา</a></td></tr>` : ''}
        </table>
      </div>
    </div>

    <!-- แถว 2: อาการที่พบ + หมายเหตุ -->
    <div class="grid2">
      <div>
        <div class="sec-title">🔧 อาการที่พบ</div>
        <div class="detail-text">${j.detail||'-'}</div>
      </div>
      <div>
        <div class="sec-title">📝 หมายเหตุ</div>
        ${j.note
          ? `<div class="note-box">${j.note}</div>`
          : `<div class="detail-text" style="color:#aaa;font-style:italic;">ไม่มีหมายเหตุ</div>`}
        ${j.actualCost ? `
        <div style="margin-top:6px;">
          <div class="sec-title">💰 ค่าใช้จ่ายจริง</div>
          <div class="cost-box">
            <div>
              <div style="font-size:9px;color:#666;">ยอดรวมค่าซ่อม</div>
              <div class="cost-amount">${Number(j.actualCost).toLocaleString()} บาท</div>
            </div>
          </div>
        </div>` : ''}
      </div>
    </div>

    <!-- แถว 3: รูปภาพแนบ (ถ้ามี) -->
    ${(parseUrls(j.imageUrl).length || parseUrls(j.estimateImageUrl).length || parseUrls(j.billUrl).length) ? `
    <div>
      <div class="sec-title">📎 เอกสารแนบ</div>
      <div class="img-row">
        ${parseUrls(j.imageUrl).map((u,i) => `<div class="img-box"><div class="img-lbl" style="background:#E3F2FD;color:#1565C0;">📷 รูปประกอบ${parseUrls(j.imageUrl).length>1?' ('+(i+1)+')':''}</div><img src="${u}" onerror="this.parentElement.style.display='none'"></div>`).join('')}
        ${parseUrls(j.estimateImageUrl).map((u,i) => `<div class="img-box"><div class="img-lbl" style="background:#E8F5E9;color:#2E7D32;">💰 ใบประเมินราคา${parseUrls(j.estimateImageUrl).length>1?' ('+(i+1)+')':''}</div><img src="${u}" onerror="this.parentElement.style.display='none'"></div>`).join('')}
        ${parseUrls(j.billUrl).map((u,i) => `<div class="img-box"><div class="img-lbl" style="background:#FFF8E1;color:#E65100;">🧾 บิล/ใบเสร็จ${parseUrls(j.billUrl).length>1?' ('+(i+1)+')':''}</div><img src="${u}" onerror="this.parentElement.style.display='none'"></div>`).join('')}
      </div>
    </div>` : ''}

    <!-- ลายเซ็น -->
    <div style="margin-top:auto;">
      <div class="sec-title">✍️ ลายเซ็นรับรอง</div>
      <div class="sign-row">
        <div class="sign-box">
          <div class="sign-line"></div>
          <div class="sign-lbl">ผู้แจ้งซ่อม<br>(${j.userName||''})</div>
        </div>
        <div class="sign-box">
          <div class="sign-line"></div>
          <div class="sign-lbl">ผู้อนุมัติ / หัวหน้างาน</div>
        </div>
        <div class="sign-box">
          <div class="sign-line"></div>
          <div class="sign-lbl">ผู้อนุมัติ / ผู้บริหาร</div>
        </div>
      </div>
    </div>

  </div><!-- /body -->

  <!-- FOOTER -->
  <div class="ft">
    <span>อ้างอิง: ${j.jobId} | ระบบแจ้งซ่อมรถยนต์</span>
    <span>พิมพ์: ${printDate}</span>
  </div>

</div>
<script>window.onload=function(){window.print();}<\/script>
<!-- ===== STATUS LOG MODAL (ข้อ 5) ===== -->
<div class="modal fade" id="modalStatusLog" tabindex="-1">
  <div class="modal-dialog modal-lg modal-dialog-scrollable">
    <div class="modal-content">
      <div class="modal-header-custom">
        <div class="modal-title"><span class="material-icons">history</span> ประวัติการเปลี่ยนสถานะ — <span id="log-job-id"></span></div>
        <button class="btn-close-white" onclick="closeModal('modalStatusLog')"><span class="material-icons" style="font-size:1rem;">close</span></button>
      </div>
      <div class="modal-body p-0">
        <div class="table-responsive d-none d-md-block">
          <table class="table mb-0">
            <thead><tr><th>วันที่</th><th>สถานะเดิม</th><th>สถานะใหม่</th><th>หมายเหตุ</th><th>ผู้เปลี่ยน</th></tr></thead>
            <tbody id="status-log-tbody"></tbody>
          </table>
        </div>
        <div class="d-md-none p-2" id="status-log-cards"></div>
      </div>
    </div>
  </div>
</div>

\u003c/body>
\u003c/html>`;

  // ตรวจว่า desktop หรือมือถือ
  const isMobile = /Android|iPhone|iPad|iPod/i.test(navigator.userAgent);
  if (isMobile) {
    showToast('กำลังสร้าง PDF... รอสักครู่');
    await savePDF(html, `แจ้งซ่อม_${j.jobId}.pdf`);
  } else {
    const win = window.open('', '_blank');
    if (!win) { showToast('กรุณาอนุญาต Popup', 'error'); return; }
    win.document.write(html);
    win.document.close();
  }
}

// refresh ข้อมูลเมื่อกลับมาที่ tab (กรณีเปิด tab อื่นนานๆ)
document.addEventListener('visibilitychange', () => {
  if (!document.hidden && currentUser) {
    loadData(true);
  }
});

/* ============================================================
   MONTHLY REPORT — Admin only
   ============================================================ */
let chartStatus = null;
let chartDaily  = null;

function initReportSelectors() {
  const yearSel  = document.getElementById('report-year');
  const monthSel = document.getElementById('report-month');
  if (!yearSel) return;

  // clear แล้ว rebuild ทุกครั้ง เพื่อให้แน่ใจว่ามีค่า
  yearSel.innerHTML = '';
  monthSel.innerHTML = '';

  const now = new Date();
  for (let y = now.getFullYear(); y >= now.getFullYear()-2; y--) {
    const o = document.createElement('option');
    o.value = y; o.text = 'ปี ' + (y + 543);
    yearSel.appendChild(o);
  }
  yearSel.value = now.getFullYear(); // ตั้ง default ปีปัจจุบัน

  const months = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                  'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  months.forEach((m, i) => {
    const o = document.createElement('option');
    o.value = i+1; o.text = m;
    monthSel.appendChild(o);
  });
  monthSel.value = now.getMonth() + 1; // ตั้ง default เดือนปัจจุบัน
}

async function generateReport() {
  // ถ้า selector ยังไม่มีค่าให้ init ก่อน
  const yearEl  = document.getElementById('report-year');
  const monthEl = document.getElementById('report-month');
  if (!yearEl.value || !monthEl.value) { initReportSelectors(); }

  const year  = parseInt(yearEl.value);
  const month = parseInt(monthEl.value);
  if (isNaN(year) || isNaN(month)) { showToast('กรุณาเลือกปีและเดือน','error'); return; }

  // reset canvas และ empty-state เก่าออกก่อน
  ['chart-status','chart-daily'].forEach(id => {
    const cv = document.getElementById(id);
    if (!cv) return;
    // ลบ empty-state div ที่อาจถูกเพิ่มไว้
    const siblings = cv.parentElement.querySelectorAll('div:not([style*="font-weight"])');
    siblings.forEach(s => { if (s.textContent.includes('ไม่มีข้อมูล')) s.remove(); });
    cv.style.display = '';
  });
  const monthName = ['','มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                     'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'][month];

  // แจ้งเตือนถ้าดูปีก่อนหน้าที่อาจถูก archive
  const thisYear = new Date().getFullYear();
  if (year < thisYear) {
    const archiveNote = document.getElementById('report-archive-note');
    if (archiveNote) archiveNote.style.display = 'block';
  } else {
    const archiveNote = document.getElementById('report-archive-note');
    if (archiveNote) archiveNote.style.display = 'none';
  }

  // กรองเฉพาะเดือนที่เลือก
  const jobs = allJobs.filter(j => {
    if (!j.createdAt) return false;
    const d = new Date(j.createdAt);
    return d.getFullYear() === year && d.getMonth()+1 === month;
  });

  // ===== Summary Cards =====
  const statuses = ['รอดำเนินการ','อนุมัติ','กำลังซ่อม','เสร็จสิ้น','ส่งกลับแก้ไข','ไม่อนุมัติ'];
  const colors   = ['#F57F17','#B71C1C','#1565C0','#7F0000','#6A1B9A','#C62828'];
  const totalCost = jobs.reduce((s,j) => s + (parseFloat(j.actualCost)||0), 0);

  const APPROVED_DS = ['อนุมัติ','กำลังซ่อม','เสร็จสิ้น'];
  const approvedCount  = jobs.filter(j => APPROVED_DS.includes(j.status)).length;
  const rejectedCount  = jobs.filter(j => j.status === 'ไม่อนุมัติ').length;
  const approvedBudget = jobs.filter(j => APPROVED_DS.includes(j.status)).reduce((s,j) => s+(Number(j.estimate)||0), 0);
  const summaryCards = [
    { val: jobs.length,  lbl: 'งานทั้งหมด',    icon: 'list_alt',     c: '#283593', bg: '#E8EAF6' },
    { val: jobs.filter(j=>j.status==='เสร็จสิ้น').length, lbl: 'เสร็จสิ้น', icon: 'check_circle', c: '#1B5E20', bg: '#E8F5E9' },
    { val: approvedCount, lbl: 'อนุมัติ',       icon: 'thumb_up',     c: '#2E7D32', bg: '#E8F5E9' },
    { val: rejectedCount, lbl: 'ไม่อนุมัติ',   icon: 'thumb_down',   c: '#C62828', bg: '#FFEBEE' },
    { val: approvedBudget > 0 ? approvedBudget.toLocaleString()+' ฿' : '-', lbl: 'ยอดอนุมัติ (ประเมิน)', icon: 'payments', c: '#F57F17', bg: '#FFF8E1' },
    { val: totalCost > 0 ? totalCost.toLocaleString()+' ฿' : '-', lbl: 'ค่าใช้จ่ายจริงรวม', icon: 'receipt', c: '#1565C0', bg: '#E3F2FD' },
  ];

  const monthLabel = monthName + ' ' + (year+543);
  // label บรรทัดเดือน (วางก่อน grid)
  const reportLabelEl = document.getElementById('report-month-label');
  if (reportLabelEl) {
    reportLabelEl.innerHTML = `📅 สรุปประจำเดือน <strong style="color:var(--primary-dark);">${monthLabel}</strong>
      ${jobs.length===0 ? '<span style="margin-left:.5rem;color:var(--danger);font-size:.78rem;">⚠️ ไม่มีข้อมูลในเดือนนี้</span>' : ''}`;
  }
  document.getElementById('report-summary-row').innerHTML = summaryCards.map(s => `
    <div class="report-summary-card">
      <div style="width:40px;height:40px;background:${s.bg};border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto .5rem;">
        <span class="material-icons" style="color:${s.c};font-size:1.3rem;">${s.icon}</span>
      </div>
      <div class="rs-val" style="color:${s.c};">${s.val}</div>
      <div class="rs-lbl">${s.lbl}</div>
    </div>`).join('');

  // ===== Chart 1: Status Pie =====
  const statusCounts = statuses.map(s => jobs.filter(j => j.status===s).length);
  const hasStatusData = statusCounts.some(c => c > 0);
  await ensureChartJs();
  if (chartStatus) chartStatus.destroy();
  const ctxStatus = document.getElementById('chart-status');
  if (!hasStatusData) {
    ctxStatus.style.display = 'none';
    ctxStatus.parentElement.innerHTML += '<div style="text-align:center;padding:2rem;color:var(--gray-400);">ไม่มีข้อมูลในเดือนนี้</div>';
  } else {
    ctxStatus.style.display = '';
    chartStatus = new Chart(ctxStatus, {
      type: 'doughnut',
      data: {
        labels: statuses,
        datasets: [{ data: statusCounts, backgroundColor: colors, borderWidth: 2, borderColor: '#fff' }]
      },
      options: {
        plugins: {
          legend: { position:'right', labels:{ font:{size:11}, boxWidth:14, padding:8 } }
        },
        cutout: '60%',
        responsive: true,
        maintainAspectRatio: true,
      }
    });
  }

  // ===== Chart 2: Daily Bar =====
  const daysInMonth = new Date(year, month, 0).getDate();
  const dailyLabels = Array.from({length:daysInMonth},(_,i)=>i+1);
  const dailyCounts = dailyLabels.map(d =>
    jobs.filter(j => { const dt=new Date(j.createdAt); return dt.getDate()===d; }).length
  );
  if (chartDaily) chartDaily.destroy();
  const ctxDaily = document.getElementById('chart-daily');
  if (!jobs.length) {
    ctxDaily.style.display = 'none';
    ctxDaily.parentElement.innerHTML += '<div style="text-align:center;padding:2rem;color:var(--gray-400);">ไม่มีข้อมูลในเดือนนี้</div>';
  } else {
    ctxDaily.style.display = '';
    // highlight วันที่มีงานมากสุด
    const maxDay = Math.max(...dailyCounts);
    const barColors = dailyCounts.map(c => c === maxDay && c > 0 ? 'rgba(46,125,50,1)' : 'rgba(46,125,50,.45)');
    chartDaily = new Chart(ctxDaily, {
      type: 'bar',
      data: {
        labels: dailyLabels,
        datasets: [{
          label: 'จำนวนงาน',
          data: dailyCounts,
          backgroundColor: barColors,
          borderRadius: 4,
          borderSkipped: false,
        }]
      },
      options: {
        plugins: {
          legend:{ display:false },
          tooltip: { callbacks: { title: ctx => 'วันที่ '+ctx[0].label } }
        },
        scales: {
          x: { grid:{display:false}, ticks:{font:{size:10}} },
          y: { beginAtZero:true, ticks:{ stepSize:1, precision:0, font:{size:10} }, grid:{ color:'#f0f0f0' } }
        },
        responsive: true,
        maintainAspectRatio: true,
      }
    });
  }

  // ===== Top Users =====
  const userMap = {};
  jobs.forEach(j => {
    const n = j.userName || 'ไม่ระบุ';
    userMap[n] = (userMap[n]||0)+1;
  });
  const topUsers = Object.entries(userMap).sort((a,b)=>b[1]-a[1]).slice(0,6);
  const maxU = topUsers[0]?.[1] || 1;
  document.getElementById('report-top-users').innerHTML = topUsers.length
    ? topUsers.map(([name,cnt],i) => `
        <div class="report-bar-wrap">
          <div style="width:18px;font-weight:700;color:var(--gray-500);text-align:right;">${i+1}.</div>
          <div style="flex:1;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${name}</div>
          <div style="width:80px;">
            <div class="report-bar" style="width:${Math.round(cnt/maxU*100)}%;background:${i===0?'var(--primary)':'var(--gray-300)'};"></div>
          </div>
          <div style="width:24px;text-align:right;font-weight:700;color:var(--primary);">${cnt}</div>
        </div>`).join('')
    : '<div style="color:var(--gray-400);text-align:center;padding:1rem;">ไม่มีข้อมูล</div>';

  // ===== Cost Summary =====
  const costJobs    = jobs.filter(j=>parseFloat(j.actualCost)>0);
  const maxCost     = Math.max(...costJobs.map(j=>parseFloat(j.actualCost)||0), 1);
  const avgCost     = costJobs.length ? totalCost/costJobs.length : 0;
  document.getElementById('report-cost-summary').innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:.5rem;margin-bottom:.75rem;">
      <div style="background:var(--primary-xlight);border-radius:8px;padding:.6rem .8rem;text-align:center;">
        <div style="font-size:1.2rem;font-weight:700;color:var(--primary);font-family:'Prompt',sans-serif;">${totalCost.toLocaleString()}</div>
        <div style="font-size:.75rem;color:var(--gray-500);">รวมทั้งหมด (บาท)</div>
      </div>
      <div style="background:#E3F2FD;border-radius:8px;padding:.6rem .8rem;text-align:center;">
        <div style="font-size:1.2rem;font-weight:700;color:#1565C0;font-family:'Prompt',sans-serif;">${Math.round(avgCost).toLocaleString()}</div>
        <div style="font-size:.75rem;color:var(--gray-500);">เฉลี่ยต่องาน (บาท)</div>
      </div>
    </div>
    ${costJobs.sort((a,b)=>parseFloat(b.actualCost)-parseFloat(a.actualCost)).slice(0,5).map(j=>`
      <div class="report-bar-wrap">
        <div style="flex:1;min-width:0;font-size:.8rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${j.jobId}</div>
        <div style="width:80px;">
          <div class="report-bar" style="width:${Math.round((parseFloat(j.actualCost)||0)/maxCost*100)}%;background:#1565C0;"></div>
        </div>
        <div style="font-size:.8rem;font-weight:700;color:#1565C0;white-space:nowrap;">${parseFloat(j.actualCost).toLocaleString()} ฿</div>
      </div>`).join('') || '<div style="color:var(--gray-400);text-align:center;padding:.5rem;">ยังไม่มีค่าใช้จ่าย</div>'}`;

  // ===== Detail Table =====
  document.getElementById('report-job-count').textContent = `${jobs.length} รายการ | ${monthName} ${year+543}`;
  document.getElementById('report-tbody').innerHTML = jobs.length
    ? [...jobs].sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt)).map(j=>`
        <tr>
          <td style="font-weight:700;color:var(--primary-dark);font-size:.8rem;">${j.jobId}</td>
          <td style="font-size:.8rem;">${formatDate(j.createdAt)}</td>
          <td style="font-size:.8rem;">${j.userName||'-'}</td>
          <td style="font-weight:600;">${j.plate}</td>
          <td><span class="badge-status ${statusClass(j.status)}" style="font-size:.75rem;">${j.status}</span></td>
          <td style="font-size:.8rem;font-weight:600;color:${j.actualCost?'var(--primary)':'var(--gray-400)'};">${j.actualCost?Number(j.actualCost).toLocaleString()+' ฿':'-'}</td>
        </tr>`).join('')
    : '<tr><td colspan="6" class="text-center py-4 text-muted">ไม่มีข้อมูลในเดือนนี้</td></tr>';
}

function exportMonthlyCSV() {
  const year  = parseInt(document.getElementById('report-year').value);
  const month = parseInt(document.getElementById('report-month').value);
  const monthName = ['','มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                     'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'][month];

  // แจ้งเตือนถ้าดูปีก่อนหน้าที่อาจถูก archive
  const thisYear = new Date().getFullYear();
  if (year < thisYear) {
    const archiveNote = document.getElementById('report-archive-note');
    if (archiveNote) archiveNote.style.display = 'block';
  } else {
    const archiveNote = document.getElementById('report-archive-note');
    if (archiveNote) archiveNote.style.display = 'none';
  }
  const jobs = allJobs.filter(j => {
    if (!j.createdAt) return false;
    const d = new Date(j.createdAt);
    return d.getFullYear()===year && d.getMonth()+1===month;
  });
  if (!jobs.length) { showToast('ไม่มีข้อมูลในเดือนนี้', 'error'); return; }

  const BOM = '\uFEFF';
  const esc = v => '"'+String(v||'').replace(/"/g,'""')+'"';
  const headers = ['เลขที่งาน','วันที่แจ้ง','ผู้แจ้ง','ทะเบียน','ไมล์','อาการ','สถานที่','สถานะ','หมายเหตุ','ค่าใช้จ่าย'];
  const rows = [...jobs]
    .sort((a,b)=>new Date(a.createdAt)-new Date(b.createdAt))
    .map(j=>[j.jobId,formatDate(j.createdAt),j.userName,j.plate,j.mileage,j.detail,j.location,j.status,j.note,j.actualCost||''].map(esc).join(','));

  const csv  = BOM + [headers.map(esc).join(','), ...rows].join('\n');
  const blob = new Blob([csv],{type:'text/csv;charset=utf-8;'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = `รายงาน_${monthName}_${year+543}.csv`; a.click();
  URL.revokeObjectURL(url);
  showToast(`Export CSV ${jobs.length} รายการ ✅`);
}

async function printMonthlyReport() {
  const year   = parseInt(document.getElementById('report-year').value);
  const month  = parseInt(document.getElementById('report-month').value);
  const monthName = ['','มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                     'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'][month];

  // แจ้งเตือนถ้าดูปีก่อนหน้าที่อาจถูก archive
  const thisYear = new Date().getFullYear();
  if (year < thisYear) {
    const archiveNote = document.getElementById('report-archive-note');
    if (archiveNote) archiveNote.style.display = 'block';
  } else {
    const archiveNote = document.getElementById('report-archive-note');
    if (archiveNote) archiveNote.style.display = 'none';
  }
  const jobs = allJobs.filter(j => {
    if (!j.createdAt) return false;
    const d = new Date(j.createdAt);
    return d.getFullYear()===year && d.getMonth()+1===month;
  });
  const totalCost = jobs.reduce((s,j) => s+(parseFloat(j.actualCost)||0), 0);
  const done = jobs.filter(j=>j.status==='เสร็จสิ้น').length;
  const pending = jobs.filter(j=>j.status==='รอดำเนินการ').length;

  // capture chart images
  const statusImg = document.getElementById('chart-status')?.toDataURL('image/png') || '';
  const dailyImg  = document.getElementById('chart-daily')?.toDataURL('image/png')  || '';
  const printDate = new Date().toLocaleDateString('th-TH',{year:'numeric',month:'long',day:'numeric'});

  // top users
  const userMap = {};
  jobs.forEach(j => { const n=j.userName||'ไม่ระบุ'; userMap[n]=(userMap[n]||0)+1; });
  const topUsers = Object.entries(userMap).sort((a,b)=>b[1]-a[1]).slice(0,8);

  const html = `\u003c!DOCTYPE html>\u003chtml lang="th"><head><meta charset="UTF-8">
<title>รายงานประจำเดือน ${monthName} ${year+543}</title>
<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&family=Prompt:wght@600;700&display=swap" rel="stylesheet">
<style>
@page { size:A4 portrait; margin:0; }
*{box-sizing:border-box;margin:0;padding:0;}
html,body{width:210mm;height:297mm;overflow:hidden;font-family:'Sarabun',sans-serif;font-size:11px;color:#222;background:#fff;}
.page{width:210mm;height:297mm;display:flex;flex-direction:column;}
.hd{background:linear-gradient(135deg,#7F0000,#B71C1C);color:#fff;padding:12px 20px;flex-shrink:0;}
.hd h1{font-family:'Prompt',sans-serif;font-size:16px;font-weight:700;}
.hd p{font-size:10px;opacity:.85;margin-top:2px;}
.hd-sub{font-size:11px;opacity:.9;margin-top:4px;}
.body{flex:1;padding:12px 20px;display:flex;flex-direction:column;gap:10px;overflow:hidden;}
.cards{display:flex;gap:8px;}
.card{flex:1;background:#f8f9fa;border:1px solid #e0e0e0;border-radius:8px;padding:8px;text-align:center;}
.card-val{font-family:'Prompt',sans-serif;font-size:18px;font-weight:700;}
.card-lbl{font-size:9px;color:#666;}
.charts{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
.chart-box{border:1px solid #e0e0e0;border-radius:8px;padding:8px;}
.chart-box h3{font-family:'Prompt',sans-serif;font-size:10px;font-weight:700;color:#B71C1C;margin-bottom:6px;}
.chart-box img{width:100%;max-height:130px;object-fit:contain;}
.sec-title{font-family:'Prompt',sans-serif;font-size:10px;font-weight:700;color:#B71C1C;border-bottom:1.5px solid #FFCDD2;padding-bottom:3px;margin-bottom:5px;}
table{width:100%;border-collapse:collapse;}
th{background:#B71C1C;color:#fff;padding:4px 6px;font-size:9.5px;text-align:left;}
td{padding:3.5px 6px;font-size:9.5px;border-bottom:1px solid #f0f0f0;}
.ft{padding:6px 20px;background:#fafafa;border-top:1px solid #e0e0e0;display:flex;justify-content:space-between;font-size:9px;color:#999;flex-shrink:0;}
</style></head><body><div class="page">
<div class="hd">
  <h1>📊 รายงานสรุปการแจ้งซ่อมรถยนต์</h1>
  <div class="hd-sub">ประจำเดือน ${monthName} ${year+543} &nbsp;|&nbsp; พิมพ์: ${printDate} &nbsp;|&nbsp; โดย: ${currentUser.name}</div>
</div>
<div class="body">
  <div class="cards">
    <div class="card"><div class="card-val" style="color:#283593;">${jobs.length}</div><div class="card-lbl">งานทั้งหมด</div></div>
    <div class="card"><div class="card-val" style="color:#1B5E20;">${done}</div><div class="card-lbl">เสร็จสิ้น</div></div>
    <div class="card"><div class="card-val" style="color:#E65100;">${pending}</div><div class="card-lbl">รอดำเนินการ</div></div>
    <div class="card"><div class="card-val" style="color:#B71C1C;">${totalCost.toLocaleString()}</div><div class="card-lbl">ค่าใช้จ่ายรวม (฿)</div></div>
  </div>
  <div class="charts">
    <div class="chart-box"><h3>🍩 สัดส่วนสถานะ</h3>${statusImg?`<img src="${statusImg}">`:''}</div>
    <div class="chart-box"><h3>📊 แจ้งซ่อมรายวัน</h3>${dailyImg?`<img src="${dailyImg}">`:''}</div>
  </div>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;flex:1;min-height:0;">
    <div>
      <div class="sec-title">🏆 ผู้แจ้งซ่อมสูงสุด</div>
      <table><thead><tr><th>#</th><th>ชื่อ</th><th>จำนวน</th></tr></thead><tbody>
      ${topUsers.map(([n,c],i)=>`<tr><td>${i+1}</td><td>${n}</td><td style="font-weight:700;color:#B71C1C;">${c}</td></tr>`).join('')}
      </tbody></table>
    </div>
    <div>
      <div class="sec-title">📋 รายการล่าสุด</div>
      <table><thead><tr><th>เลขงาน</th><th>ทะเบียน</th><th>สถานะ</th><th>ค่าใช้จ่าย</th></tr></thead><tbody>
      ${[...jobs].sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt)).slice(0,10).map(j=>`
        <tr><td style="font-weight:700;">${j.jobId}</td><td>${j.plate}</td><td>${j.status}</td><td>${j.actualCost?Number(j.actualCost).toLocaleString()+' ฿':'-'}</td></tr>`).join('')}
      </tbody></table>
    </div>
  </div>
</div>
<div class="ft"><span>รายงานประจำเดือน ${monthName} ${year+543} | ระบบแจ้งซ่อมรถยนต์</span><span>พิมพ์: ${printDate}</span></div>
</div><script>window.onload=function(){window.print();}<\/script>\u003c/body>\u003c/html>`;

  const isMobile = /Android|iPhone|iPad|iPod/i.test(navigator.userAgent);
  const monthName2 = ['','ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
                      'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'][month];
  if (isMobile) {
    showToast('กำลังสร้าง PDF...');
    await savePDF(html, `รายงาน_${monthName2}_${year+543}.pdf`);
  } else {
    const win = window.open('','_blank');
    if (!win) { showToast('กรุณาอนุญาต Popup', 'error'); return; }
    win.document.write(html);
    win.document.close();
  }
}

function logout() {
  if (confirm('ออกจากระบบ?')) {
    sessionStorage.removeItem('_repairUser');
    sessionStorage.removeItem('_repairData');
    if (_warmTimer) { clearInterval(_warmTimer); _warmTimer = null; }
    currentUser = null;
    try { liff.logout(); } catch(e){}
    document.getElementById('main-app').style.display='none';
    document.getElementById('login-screen').style.display='flex';
  }
}

// Filter pills
document.querySelectorAll('.filter-pill').forEach(p => {
  p.addEventListener('click', () => {
    document.querySelectorAll('.filter-pill').forEach(x=>x.classList.remove('active'));
    p.classList.add('active');
    filterStatus = p.dataset.status;
    currentPage = 1;
    renderRepairTable();
  });
});

// Search & date filter live
document.getElementById('search-input').addEventListener('input', () => { currentPage = 1; renderRepairTable(); });
document.getElementById('filter-date').addEventListener('change', () => { currentPage = 1; renderRepairTable(); });
document.getElementById('repair-sort').addEventListener('change', () => { currentPage = 1; renderRepairTable(); });
