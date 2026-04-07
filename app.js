import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import {
  getAuth,
  signInWithEmailAndPassword,
  signOut,
  onAuthStateChanged
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js';
import {
  getDatabase,
  ref,
  onValue,
  push,
  set,
  update,
  remove,
  get,
  child
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-database.js';

const firebaseConfig = {
  apiKey: 'AIzaSyA9hxi8keOUJG_mhdD4OSN32A1jypXrXEA',
  authDomain: 'grading-dura.firebaseapp.com',
  projectId: 'grading-dura',
  storageBucket: 'grading-dura.firebasestorage.app',
  messagingSenderId: '455000354944',
  appId: '1:455000354944:web:69b96169f6174ec5a8b665',
  measurementId: 'G-9J29KM9NHC',
  databaseURL: 'https://grading-dura-default-rtdb.asia-southeast1.firebasedatabase.app'
};

const ROLE_EMAILS = {
  staff: 'staff@dura.local',
  grading: 'grading@dura.local'
};

const pageMeta = {
  dashboard: ['Dashboard', 'Ringkasan operasional grading, Tenera Dura, supplier, dan sopir.'],
  grading: ['Input Grading', 'Fokus utama pada % kematangan dan total potongan.'],
  td: ['Input Tenera Dura', 'Modul terpisah dari grading.'],
  rekapGrading: ['Rekap Grading', 'Data grading lengkap per transaksi.'],
  rekapTD: ['Rekap Tenera Dura', 'Data Tenera Dura lengkap per transaksi.'],
  rekapData: ['Rekap Data', 'Kesimpulan otomatis berdasarkan tanggal dan filter.'],
  sheetGrading: ['Spreadsheet Grading', 'Tarikan spreadsheet detail satu kolom satu data.'],
  sheetTD: ['Spreadsheet Tenera Dura', 'Spreadsheet detail untuk Tenera Dura.'],
  performance: ['Performance', 'Ranking grading, TD, dan gabungan.'],
  analytics: ['Analytics', 'Penyebab potongan dan insight manajerial.'],
  supplier: ['Supplier', 'Kelola master supplier.']
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getDatabase(app);

const state = {
  suppliers: [],
  grading: [],
  td: [],
  currentRole: 'grading',
  currentUser: null,
  activePage: 'dashboard',
  listenersBound: false,
  realtimeBound: false,
  loading: { suppliers: true, grading: true, td: true },
  waContext: null,
  sheetGradeDraft: null,
  sheetTDDraft: null
};

const defaultSuppliers = [
  { name: 'CV LEMBAH HIJAU PERKASA', status: 'active' },
  { name: 'KOPERASI KARYA MANDIRI', status: 'active' },
  { name: 'TANI RAMPAH JAYA', status: 'active' },
  { name: 'PT PUTRA UTAMA LESTARI', status: 'active' },
  { name: 'PT MANUNGGAL ADI JAYA', status: 'active' }
];

function el(id) { return document.getElementById(id); }
function uid() { return 'id-' + Math.random().toString(36).slice(2, 9) + Date.now().toString(36); }
function num(v) { return Number(v || 0); }
function fixed(v) { return Number(v || 0).toFixed(2); }
function pct(v) { return `${fixed(v)}%`; }
function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
function dt(iso) {
  const date = new Date(iso || Date.now());
  return {
    date: date.toLocaleDateString('id-ID'),
    time: date.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })
  };
}
function dateOnly(iso) {
  const d = new Date(iso || Date.now());
  const tz = d.getTimezoneOffset() * 60000;
  return new Date(d.getTime() - tz).toISOString().slice(0, 10);
}
function normalizeText(text) { return String(text || '').trim().toLowerCase(); }
function metric(label, value) { return `<div class="metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`; }
function stat(title, meta, score = 0) {
  const width = Math.max(8, Math.min(100, Math.abs(Number(score) || 0) * 4));
  return `<div class="stat"><strong>${escapeHtml(title)}</strong><div class="meta">${meta}</div><div class="bar"><div style="width:${width}%"></div></div></div>`;
}
function sortByCreatedAtDesc(a, b) { return new Date(b.createdAt || 0) - new Date(a.createdAt || 0); }
function sortByName(a, b) { return String(a.name || '').localeCompare(String(b.name || ''), 'id'); }
function toArray(snapshotVal) {
  if (!snapshotVal) return [];
  return Object.entries(snapshotVal).map(([id, value]) => ({ id, ...value }));
}
function activeSuppliers() { return state.suppliers.filter((s) => s.status !== 'inactive'); }
function supplierExistsByName(name) { return state.suppliers.some((s) => normalizeText(s.name) === normalizeText(name)); }
function setStatus(message = '', type = 'info') {
  const box = el('appStatus');
  if (!message) {
    box.className = 'alert info hidden';
    box.textContent = '';
    return;
  }
  box.className = `alert ${type}`;
  box.textContent = message;
}
function showLoginError(message = '') {
  const box = el('loginError');
  if (!message) {
    box.classList.add('hidden');
    box.textContent = '';
    return;
  }
  box.classList.remove('hidden');
  box.textContent = message;
}
function isStaffEmail(email) { return normalizeText(email) === normalizeText(ROLE_EMAILS.staff); }
function isGradingEmail(email) { return normalizeText(email) === normalizeText(ROLE_EMAILS.grading); }
function deriveRole(email) { return isStaffEmail(email) ? 'staff' : 'grading'; }

function calculateGrading(data) {
  const totalBunches = num(data.totalBunches);
  const mentah = num(data.mentah);
  const mengkal = num(data.mengkal);
  const overripe = num(data.overripe);
  const busuk = num(data.busuk);
  const kosong = num(data.kosong);
  const partheno = num(data.partheno);
  const tikus = num(data.tikus);
  const totalCategories = mentah + mengkal + overripe + busuk + kosong + partheno + tikus;
  const masak = totalBunches - totalCategories;
  const toPct = (value) => (totalBunches > 0 ? (value / totalBunches) * 100 : 0);
  const percentages = {
    masak: toPct(Math.max(masak, 0)),
    mentah: toPct(mentah),
    mengkal: toPct(mengkal),
    overripe: toPct(overripe),
    busuk: toPct(busuk),
    kosong: toPct(kosong),
    partheno: toPct(partheno),
    tikus: toPct(tikus)
  };
  const deductions = {
    dasar: 3,
    mentah: percentages.mentah * 0.5,
    mengkal: percentages.mengkal * 0.15,
    overripe: percentages.overripe > 5 ? (percentages.overripe - 5) * 0.25 : 0,
    busuk: percentages.busuk,
    kosong: percentages.kosong,
    partheno: percentages.partheno * 0.15,
    tikus: percentages.tikus * 0.15
  };
  const totalDeduction = Object.values(deductions).reduce((sum, value) => sum + value, 0);
  let validation = { type: 'info', message: 'Perhitungan siap disimpan.' };
  if (totalCategories > totalBunches) {
    validation = { type: 'error', message: 'ERROR: Total kategori melebihi Total Janjang.' };
  } else if (masak < 0) {
    validation = { type: 'error', message: 'ERROR: Nilai masak negatif.' };
  } else if (!data.driver || !data.plate || !data.supplier || totalBunches <= 0) {
    validation = { type: 'warning', message: 'Lengkapi field wajib dan pastikan Total Janjang lebih dari 0.' };
  }
  let status = 'BAIK';
  let statusClass = 'ok';
  if (totalDeduction > 15) {
    status = 'BURUK';
    statusClass = 'bad';
  } else if (totalDeduction > 8) {
    status = 'PERLU PERHATIAN';
    statusClass = 'warn';
  }
  return {
    totalBunches,
    mentah,
    mengkal,
    overripe,
    busuk,
    kosong,
    partheno,
    tikus,
    masak,
    percentages,
    deductions,
    totalDeduction,
    status,
    statusClass,
    validation
  };
}

function calculateTD(data) {
  const tenera = num(data.tenera);
  const dura = num(data.dura);
  const total = tenera + dura;
  const pctTenera = total > 0 ? (tenera / total) * 100 : 0;
  const pctDura = total > 0 ? (dura / total) * 100 : 0;
  return {
    tenera,
    dura,
    total,
    pctTenera,
    pctDura,
    dominant: pctTenera === pctDura ? '-' : pctTenera > pctDura ? 'Tenera' : 'Dura'
  };
}

async function ensureDefaultSuppliers() {
  const snap = await get(ref(db, 'suppliers'));
  if (snap.exists()) return;
  const writes = defaultSuppliers.map((item) => {
    const node = push(ref(db, 'suppliers'));
    return set(node, {
      name: item.name,
      status: item.status,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });
  });
  await Promise.all(writes);
}

function bindRealtime() {
  if (state.realtimeBound) return;
  state.realtimeBound = true;
  onValue(ref(db, 'suppliers'), (snapshot) => {
    state.suppliers = toArray(snapshot.val()).sort(sortByName);
    state.loading.suppliers = false;
    renderAll();
  }, (error) => {
    state.loading.suppliers = false;
    setStatus(`Gagal memuat supplier: ${error.message}`, 'error');
    renderAll();
  });

  onValue(ref(db, 'grading'), (snapshot) => {
    state.grading = toArray(snapshot.val()).sort(sortByCreatedAtDesc);
    state.loading.grading = false;
    renderAll();
  }, (error) => {
    state.loading.grading = false;
    setStatus(`Gagal memuat data grading: ${error.message}`, 'error');
    renderAll();
  });

  onValue(ref(db, 'td'), (snapshot) => {
    state.td = toArray(snapshot.val()).sort(sortByCreatedAtDesc);
    state.loading.td = false;
    renderAll();
  }, (error) => {
    state.loading.td = false;
    setStatus(`Gagal memuat data Tenera Dura: ${error.message}`, 'error');
    renderAll();
  });
}

function setupRolePicker() {
  document.querySelectorAll('.role-pick').forEach((button) => {
    button.addEventListener('click', () => {
      document.querySelectorAll('.role-pick').forEach((btn) => btn.classList.remove('active'));
      button.classList.add('active');
      const role = button.dataset.role;
      el('loginEmail').value = ROLE_EMAILS[role] || '';
      showLoginError('');
    });
  });
  el('loginEmail').value = ROLE_EMAILS.staff;
  el('loginEmail').setAttribute('autocomplete', 'username');
  el('loginPassword').setAttribute('autocomplete', 'current-password');
}

async function handleLoginSubmit(event) {
  event.preventDefault();
  showLoginError('');
  const email = el('loginEmail').value.trim();
  const password = el('loginPassword').value;
  if (!email || !password) {
    showLoginError('Email dan password wajib diisi.');
    return;
  }
  try {
    await signInWithEmailAndPassword(auth, email, password);
    el('loginPassword').value = '';
  } catch (error) {
    showLoginError(error.message || 'Login gagal.');
  }
}

function updateUserUI() {
  el('roleLabel').textContent = state.currentRole.toUpperCase();
  if (el('userEmail')) el('userEmail').textContent = state.currentUser?.email || '-';
  document.querySelectorAll('.staff-only,.staff-only-page').forEach((node) => {
    node.classList.toggle('hidden', state.currentRole !== 'staff');
  });
  if (state.currentRole !== 'staff' && ['sheetGrading', 'sheetTD', 'performance', 'analytics', 'supplier'].includes(state.activePage)) {
    switchPage('dashboard');
  }
}

function toggleSidebar(open) {
  const appNode = el('app');
  if (window.innerWidth > 900) return;
  appNode.classList.toggle('sidebar-open', open);
}

function switchPage(page) {
  if (state.currentRole !== 'staff' && ['sheetGrading', 'sheetTD', 'performance', 'analytics', 'supplier'].includes(page)) {
    page = 'dashboard';
  }
  state.activePage = page;
  document.querySelectorAll('.menu-item').forEach((btn) => btn.classList.toggle('active', btn.dataset.page === page));
  document.querySelectorAll('.page').forEach((section) => section.classList.toggle('active', section.id === `page-${page}`));
  const [title, subtitle] = pageMeta[page] || ['Dashboard', ''];
  el('pageTitle').textContent = title;
  el('pageSubtitle').textContent = subtitle;
  el('summaryCards').classList.toggle('hidden', page !== 'dashboard');
  toggleSidebar(false);
}

function fillStatic() {
  const active = activeSuppliers();
  const gradingSelect = el('gradingSupplier');
  gradingSelect.innerHTML = '<option value="">Pilih supplier</option>' + active.map((supplier) => `<option value="${escapeHtml(supplier.name)}">${escapeHtml(supplier.name)}</option>`).join('');

  const filterOptions = '<option value="">Semua Supplier</option>' + state.suppliers.map((supplier) => `<option value="${escapeHtml(supplier.name)}">${escapeHtml(supplier.name)}</option>`).join('');
  el('rekapGradingSupplier').innerHTML = filterOptions;
  el('rekapDataSupplier').innerHTML = filterOptions;
  el('waSupplier').innerHTML = filterOptions;

  const driverSet = [...new Set([...state.grading.map((row) => row.driver), ...state.td.map((row) => row.driver)].filter(Boolean))].sort();
  const options = driverSet.map((name) => `<option value="${escapeHtml(name)}"></option>`).join('');
  el('driverList').innerHTML = options;
  el('tdDriverList').innerHTML = options;

  if (!active.length) {
    gradingSelect.innerHTML = '<option value="">Belum ada supplier aktif</option>';
  }
}

function historyHint(name) {
  const last = [...state.grading]
    .filter((item) => normalizeText(item.driver) === normalizeText(name))
    .sort(sortByCreatedAtDesc)[0];
  if (!last) {
    el('driverHint').textContent = 'Belum ada histori sopir.';
    return;
  }
  el('gradingPlate').value = last.plate || '';
  el('gradingSupplier').value = last.supplier || '';
  el('driverHint').textContent = `Histori terakhir: Plat ${last.plate} | Supplier ${last.supplier}`;
}

function filterDate(rows, start, end) {
  return rows.filter((row) => {
    const date = dateOnly(row.createdAt);
    if (start && date < start) return false;
    if (end && date > end) return false;
    return true;
  });
}

function supplierStats(rows = state.grading) {
  const map = {};
  rows.forEach((row) => {
    if (!map[row.supplier]) {
      map[row.supplier] = {
        name: row.supplier,
        count: 0,
        totalJanjang: 0,
        masakPct: 0,
        totalDed: 0,
        maxDed: -Infinity,
        minDed: Infinity
      };
    }
    const item = map[row.supplier];
    item.count += 1;
    item.totalJanjang += num(row.totalBunches);
    item.masakPct += num(row.percentages?.masak);
    item.totalDed += num(row.totalDeduction);
    item.maxDed = Math.max(item.maxDed, num(row.totalDeduction));
    item.minDed = Math.min(item.minDed, num(row.totalDeduction));
  });
  return Object.values(map).map((item) => ({
    ...item,
    avgMasak: item.count ? item.masakPct / item.count : 0,
    avgDed: item.count ? item.totalDed / item.count : 0,
    maxDed: Number.isFinite(item.maxDed) ? item.maxDed : 0,
    minDed: Number.isFinite(item.minDed) ? item.minDed : 0
  })).sort((a, b) => a.avgDed - b.avgDed);
}

function driverStats(rows = state.grading) {
  const map = {};
  rows.forEach((row) => {
    if (!map[row.driver]) {
      map[row.driver] = { name: row.driver, count: 0, totalJanjang: 0, masakPct: 0, totalDed: 0, suppliers: {} };
    }
    const item = map[row.driver];
    item.count += 1;
    item.totalJanjang += num(row.totalBunches);
    item.masakPct += num(row.percentages?.masak);
    item.totalDed += num(row.totalDeduction);
    item.suppliers[row.supplier] = (item.suppliers[row.supplier] || 0) + 1;
  });
  return Object.values(map).map((item) => ({
    ...item,
    avgMasak: item.count ? item.masakPct / item.count : 0,
    avgDed: item.count ? item.totalDed / item.count : 0,
    topSupplier: Object.entries(item.suppliers).sort((a, b) => b[1] - a[1])[0]?.[0] || '-'
  })).sort((a, b) => b.totalJanjang - a.totalJanjang);
}

function tdDriverStats(rows = state.td) {
  const map = {};
  rows.forEach((row) => {
    if (!map[row.driver]) {
      map[row.driver] = { name: row.driver, count: 0, total: 0, totalTenera: 0, totalDura: 0, plates: {} };
    }
    const item = map[row.driver];
    item.count += 1;
    item.total += num(row.total);
    item.totalTenera += num(row.pctTenera);
    item.totalDura += num(row.pctDura);
    item.plates[row.plate] = (item.plates[row.plate] || 0) + 1;
  });
  return Object.values(map).map((item) => ({
    ...item,
    avgTenera: item.count ? item.totalTenera / item.count : 0,
    avgDura: item.count ? item.totalDura / item.count : 0,
    topPlate: Object.entries(item.plates).sort((a, b) => b[1] - a[1])[0]?.[0] || '-'
  })).sort((a, b) => b.total - a.total);
}

function causeTotals(rows = state.grading) {
  const totals = { mentah: 0, mengkal: 0, overripe: 0, busuk: 0, kosong: 0, partheno: 0, tikus: 0 };
  rows.forEach((row) => {
    Object.keys(totals).forEach((key) => { totals[key] += num(row.deductions?.[key]); });
  });
  return totals;
}

function insights(rows = state.grading) {
  if (!rows.length) return ['Belum ada data grading.'];
  const avgMasak = rows.reduce((sum, row) => sum + num(row.percentages?.masak), 0) / rows.length;
  const avgDed = rows.reduce((sum, row) => sum + num(row.totalDeduction), 0) / rows.length;
  const worst = [...rows].sort((a, b) => num(b.totalDeduction) - num(a.totalDeduction))[0];
  const cause = Object.entries(causeTotals(rows)).sort((a, b) => b[1] - a[1])[0];
  return [
    `Rata-rata kematangan ${pct(avgMasak)}.`,
    `Rata-rata total potongan ${pct(avgDed)}.`,
    worst ? `Potongan tertinggi: ${worst.driver} (${pct(worst.totalDeduction)}).` : 'Belum ada potongan tertinggi.',
    cause ? `Penyebab potongan terbesar: ${cause[0]} (${pct(cause[1])}).` : 'Belum ada penyebab dominan.'
  ];
}

function renderSummaryCards() {
  const totalJanjang = state.grading.reduce((sum, row) => sum + num(row.totalBunches), 0);
  const avgMasak = state.grading.length ? state.grading.reduce((sum, row) => sum + num(row.percentages?.masak), 0) / state.grading.length : 0;
  const avgDed = state.grading.length ? state.grading.reduce((sum, row) => sum + num(row.totalDeduction), 0) / state.grading.length : 0;
  const avgT = state.td.length ? state.td.reduce((sum, row) => sum + num(row.pctTenera), 0) / state.td.length : 0;
  el('summaryCards').innerHTML = `
    <div class="summary-card"><span class="label">Total Janjang</span><span class="value">${totalJanjang}</span><span class="sub">Akumulasi grading</span></div>
    <div class="summary-card"><span class="label">Rata-rata % Masak</span><span class="value">${pct(avgMasak)}</span><span class="sub">Fokus kematangan</span></div>
    <div class="summary-card hot"><span class="label">Rata-rata Potongan</span><span class="value">${pct(avgDed)}</span><span class="sub">Fokus utama UI</span></div>
    <div class="summary-card"><span class="label">Rata-rata % Tenera</span><span class="value">${pct(avgT)}</span><span class="sub">Modul Tenera Dura</span></div>`;
}

function renderDashboard() {
  const gradingRows = state.grading;
  const tdRows = state.td;
  el('dashGrading').innerHTML = [
    metric('Data Grading', gradingRows.length),
    metric('Total Janjang', gradingRows.reduce((sum, row) => sum + num(row.totalBunches), 0)),
    metric('% Masak Rata-rata', pct(gradingRows.length ? gradingRows.reduce((sum, row) => sum + num(row.percentages?.masak), 0) / gradingRows.length : 0)),
    metric('Potongan Rata-rata', pct(gradingRows.length ? gradingRows.reduce((sum, row) => sum + num(row.totalDeduction), 0) / gradingRows.length : 0))
  ].join('');
  el('dashTD').innerHTML = [
    metric('Data TD', tdRows.length),
    metric('Total TD', tdRows.reduce((sum, row) => sum + num(row.total), 0)),
    metric('% Tenera Rata-rata', pct(tdRows.length ? tdRows.reduce((sum, row) => sum + num(row.pctTenera), 0) / tdRows.length : 0)),
    metric('% Dura Rata-rata', pct(tdRows.length ? tdRows.reduce((sum, row) => sum + num(row.pctDura), 0) / tdRows.length : 0))
  ].join('');
  el('dashInsights').innerHTML = insights().map((text) => `<div class="stat">${escapeHtml(text)}</div>`).join('');
  el('dashSuppliers').innerHTML = supplierStats().slice(0, 8).map((row) => stat(row.name, `Transaksi: ${row.count} | Total Janjang: ${row.totalJanjang} | % Masak: ${pct(row.avgMasak)} | Potongan: ${pct(row.avgDed)}`, row.avgDed)).join('') || '<div class="empty-state">Belum ada data supplier.</div>';
  el('dashDrivers').innerHTML = driverStats().slice(0, 8).map((row) => stat(row.name, `Transaksi: ${row.count} | Total Janjang: ${row.totalJanjang} | % Masak: ${pct(row.avgMasak)} | Potongan: ${pct(row.avgDed)}`, row.avgDed)).join('') || '<div class="empty-state">Belum ada data sopir.</div>';
}

function renderGradingLive() {
  const form = el('gradingForm');
  const data = Object.fromEntries(new FormData(form).entries());
  const calc = calculateGrading(data);
  el('gradingTotalDeduction').textContent = pct(calc.totalDeduction);
  const statusNode = el('gradingStatus');
  statusNode.className = `status ${calc.statusClass}`;
  statusNode.textContent = calc.status;
  el('gradingLiveCards').innerHTML = [
    metric('Total Janjang', calc.totalBunches),
    metric('Masak', calc.masak),
    metric('% Masak', pct(calc.percentages.masak)),
    metric('Potongan Dasar', pct(calc.deductions.dasar))
  ].join('');
  const rows = [
    ['Masak', calc.masak, calc.percentages.masak, 0],
    ['Mentah', calc.mentah, calc.percentages.mentah, calc.deductions.mentah],
    ['Mengkal', calc.mengkal, calc.percentages.mengkal, calc.deductions.mengkal],
    ['Overripe', calc.overripe, calc.percentages.overripe, calc.deductions.overripe],
    ['Busuk', calc.busuk, calc.percentages.busuk, calc.deductions.busuk],
    ['Tandan Kosong', calc.kosong, calc.percentages.kosong, calc.deductions.kosong],
    ['Parthenocarpi', calc.partheno, calc.percentages.partheno, calc.deductions.partheno],
    ['Makan Tikus', calc.tikus, calc.percentages.tikus, calc.deductions.tikus]
  ];
  el('gradingBreakdown').innerHTML = rows.map((row) => `<tr><td>${row[0]}</td><td>${row[1]}</td><td>${pct(row[2])}</td><td>${pct(row[3])}</td></tr>`).join('');
  const info = el('gradingValidation');
  info.className = `alert ${calc.validation.type}`;
  info.textContent = calc.validation.message;
}

function renderTDLive() {
  const form = el('tdForm');
  const data = Object.fromEntries(new FormData(form).entries());
  const calc = calculateTD(data);
  el('tdTotal').textContent = String(calc.total);
  el('tdPctTenera').textContent = pct(calc.pctTenera);
  el('tdPctDura').textContent = pct(calc.pctDura);
  el('tdDominant').textContent = calc.dominant;
  el('tdBarTenera').style.width = `${calc.pctTenera}%`;
  el('tdBarDura').style.width = `${calc.pctDura}%`;
  el('tdDonut').style.background = `conic-gradient(var(--primary) ${calc.pctTenera * 3.6}deg,#efc56e 0deg)`;
  el('tdDonutText').textContent = pct(calc.pctTenera);
}

function renderRekapGrading() {
  const q = normalizeText(el('rekapGradingSearch').value);
  const supplier = el('rekapGradingSupplier').value;
  const start = el('rekapGradingStart').value;
  const end = el('rekapGradingEnd').value;
  let rows = filterDate(state.grading, start, end);
  rows = rows.filter((row) => {
    const haystack = `${row.driver} ${row.plate} ${row.supplier}`.toLowerCase();
    return (!q || haystack.includes(q)) && (!supplier || row.supplier === supplier);
  });
  el('rekapGradingTable').innerHTML = rows.map((row) => {
    const info = dt(row.createdAt);
    return `<tr data-detail-type="grading" data-detail-id="${row.id}"><td>${info.date}</td><td>${info.time}</td><td>${escapeHtml(row.driver)}</td><td>${escapeHtml(row.plate)}</td><td>${escapeHtml(row.supplier)}</td><td>${row.totalBunches}</td><td>${pct(row.percentages?.masak)}</td><td>${pct(row.totalDeduction)}</td><td>${row.revised ? 'Revisi' : '-'}</td></tr>`;
  }).join('') || '<tr><td colspan="9">Tidak ada data grading.</td></tr>';
}

function renderRekapTD() {
  const q = normalizeText(el('rekapTDSearch').value);
  const start = el('rekapTDStart').value;
  const end = el('rekapTDEnd').value;
  let rows = filterDate(state.td, start, end);
  rows = rows.filter((row) => !q || `${row.driver} ${row.plate}`.toLowerCase().includes(q));
  el('rekapTDTable').innerHTML = rows.map((row) => {
    const info = dt(row.createdAt);
    return `<tr data-detail-type="td" data-detail-id="${row.id}"><td>${info.date}</td><td>${info.time}</td><td>${escapeHtml(row.driver)}</td><td>${escapeHtml(row.plate)}</td><td>${row.tenera}</td><td>${row.dura}</td><td>${pct(row.pctTenera)}</td><td>${pct(row.pctDura)}</td><td>${row.revised ? 'Revisi' : '-'}</td></tr>`;
  }).join('') || '<tr><td colspan="9">Tidak ada data Tenera Dura.</td></tr>';
}

function getRekapDataFiltered() {
  const start = el('rekapDataStart').value;
  const end = el('rekapDataEnd').value;
  const supplier = el('rekapDataSupplier').value;
  const driver = normalizeText(el('rekapDataDriver').value);
  const grading = filterDate(state.grading, start, end).filter((row) => (!supplier || row.supplier === supplier) && (!driver || normalizeText(row.driver).includes(driver)));
  const td = filterDate(state.td, start, end).filter((row) => !driver || normalizeText(row.driver).includes(driver));
  return { start, end, grading, td };
}

function renderRekapData() {
  const { start, end, grading, td } = getRekapDataFiltered();
  const avgMasak = grading.length ? grading.reduce((sum, row) => sum + num(row.percentages?.masak), 0) / grading.length : 0;
  const avgDed = grading.length ? grading.reduce((sum, row) => sum + num(row.totalDeduction), 0) / grading.length : 0;
  const avgT = td.length ? td.reduce((sum, row) => sum + num(row.pctTenera), 0) / td.length : 0;
  const avgD = td.length ? td.reduce((sum, row) => sum + num(row.pctDura), 0) / td.length : 0;
  const cause = Object.entries(causeTotals(grading)).sort((a, b) => b[1] - a[1])[0]?.[0] || '-';
  const bestSupplier = supplierStats(grading)[0]?.name || '-';
  const topDriver = driverStats(grading)[0]?.name || '-';
  const topTDDriver = tdDriverStats(td)[0]?.name || '-';
  el('rekapDataGradingSummary').innerHTML = [
    stat('Periode', `${start || '-'} s/d ${end || '-'}`),
    stat('Transaksi', `${grading.length} transaksi | Total Janjang: ${grading.reduce((sum, row) => sum + num(row.totalBunches), 0)}`),
    stat('Kematangan', `Rata-rata % Masak: ${pct(avgMasak)}`, avgMasak),
    stat('Potongan', `Rata-rata Potongan: ${pct(avgDed)} | Penyebab: ${cause}`, avgDed),
    stat('Highlight', `Supplier terbaik: ${bestSupplier} | Sopir terbanyak: ${topDriver}`)
  ].join('');
  el('rekapDataTDSummary').innerHTML = [
    stat('Periode', `${start || '-'} s/d ${end || '-'}`),
    stat('Transaksi TD', `${td.length} transaksi | Total TD: ${td.reduce((sum, row) => sum + num(row.total), 0)}`),
    stat('Komposisi', `Rata-rata % Tenera: ${pct(avgT)} | Rata-rata % Dura: ${pct(avgD)}`, avgT),
    stat('Highlight', `Sopir TD terbanyak: ${topTDDriver}`)
  ].join('');
  el('rekapDataSupplierTable').innerHTML = supplierStats(grading).map((row) => `<tr><td>${escapeHtml(row.name)}</td><td>${row.count}</td><td>${row.totalJanjang}</td><td>${pct(row.avgMasak)}</td><td>${pct(row.avgDed)}</td></tr>`).join('') || '<tr><td colspan="5">Tidak ada data supplier.</td></tr>';
  const tdMap = Object.fromEntries(tdDriverStats(td).map((row) => [row.name, row]));
  el('rekapDataDriverTable').innerHTML = driverStats(grading).map((row) => `<tr><td>${escapeHtml(row.name)}</td><td>${row.count}</td><td>${row.totalJanjang}</td><td>${pct(row.avgMasak)}</td><td>${tdMap[row.name]?.count || 0}</td></tr>`).join('') || '<tr><td colspan="5">Tidak ada data sopir.</td></tr>';
}

function renderSheetGrading() {
  const q = normalizeText(el('sheetGradingSearch').value);
  const rows = state.grading.filter((row) => !q || JSON.stringify(row).toLowerCase().includes(q));
  const cols = [
    ['date', 'Tanggal'], ['time', 'Jam'], ['driver', 'Sopir'], ['plate', 'Plat'], ['supplier', 'Supplier'], ['totalBunches', 'Total Janjang'], ['masak', 'Masak'], ['pctMasak', '% Masak'],
    ['mentah', 'Mentah'], ['pctMentah', '% Mentah'], ['mengkal', 'Mengkal'], ['pctMengkal', '% Mengkal'], ['overripe', 'Overripe'], ['pctOverripe', '% Overripe'],
    ['busuk', 'Busuk'], ['pctBusuk', '% Busuk'], ['kosong', 'Tandan Kosong'], ['pctKosong', '% Tandan Kosong'], ['partheno', 'Parthenocarpi'], ['pctPartheno', '% Parthenocarpi'],
    ['tikus', 'Makan Tikus'], ['pctTikus', '% Makan Tikus'], ['dedDasar', 'Potongan Dasar'], ['dedMentah', 'Potongan Mentah'], ['dedMengkal', 'Potongan Mengkal'], ['dedOverripe', 'Potongan Overripe'],
    ['dedBusuk', 'Potongan Busuk'], ['dedKosong', 'Potongan Tandan Kosong'], ['dedPartheno', 'Potongan Parthenocarpi'], ['dedTikus', 'Potongan Makan Tikus'], ['totalDeduction', 'Total Potongan'], ['revised', 'Revisi'], ['revisedAt', 'Waktu Revisi'], ['action', 'Aksi']
  ];
  const editableKeys = ['driver', 'plate', 'supplier', 'totalBunches', 'mentah', 'mengkal', 'overripe', 'busuk', 'kosong', 'partheno', 'tikus'];
  el('sheetGradingTable').innerHTML = `<thead><tr>${cols.map((col) => `<th>${col[1]}</th>`).join('')}</tr></thead><tbody>${rows.map((row) => {
    const info = dt(row.createdAt);
    const values = {
      date: info.date,
      time: info.time,
      driver: row.driver,
      plate: row.plate,
      supplier: row.supplier,
      totalBunches: row.totalBunches,
      masak: row.masak,
      pctMasak: fixed(row.percentages?.masak),
      mentah: row.mentah,
      pctMentah: fixed(row.percentages?.mentah),
      mengkal: row.mengkal,
      pctMengkal: fixed(row.percentages?.mengkal),
      overripe: row.overripe,
      pctOverripe: fixed(row.percentages?.overripe),
      busuk: row.busuk,
      pctBusuk: fixed(row.percentages?.busuk),
      kosong: row.kosong,
      pctKosong: fixed(row.percentages?.kosong),
      partheno: row.partheno,
      pctPartheno: fixed(row.percentages?.partheno),
      tikus: row.tikus,
      pctTikus: fixed(row.percentages?.tikus),
      dedDasar: fixed(row.deductions?.dasar),
      dedMentah: fixed(row.deductions?.mentah),
      dedMengkal: fixed(row.deductions?.mengkal),
      dedOverripe: fixed(row.deductions?.overripe),
      dedBusuk: fixed(row.deductions?.busuk),
      dedKosong: fixed(row.deductions?.kosong),
      dedPartheno: fixed(row.deductions?.partheno),
      dedTikus: fixed(row.deductions?.tikus),
      totalDeduction: fixed(row.totalDeduction),
      revised: row.revised ? 'Ya' : '-',
      revisedAt: row.revisedAt ? `${dt(row.revisedAt).date} ${dt(row.revisedAt).time}` : '-'
    };
    return `<tr data-id="${row.id}">${cols.map(([key]) => {
      if (key === 'action') return `<td><button type="button" class="text-btn danger" data-delete-grading="${row.id}">Hapus</button></td>`;
      const editable = editableKeys.includes(key);
      return `<td ${editable ? `class="editable" contenteditable="true" data-key="${key}"` : ''}>${escapeHtml(values[key])}</td>`;
    }).join('')}</tr>`;
  }).join('')}</tbody>`;
}

function renderSheetTD() {
  const q = normalizeText(el('sheetTDSearch').value);
  const rows = state.td.filter((row) => !q || JSON.stringify(row).toLowerCase().includes(q));
  const cols = [['date', 'Tanggal'], ['time', 'Jam'], ['driver', 'Sopir'], ['plate', 'Plat'], ['tenera', 'Tenera'], ['dura', 'Dura'], ['total', 'Total TD'], ['pctTenera', '% Tenera'], ['pctDura', '% Dura'], ['revised', 'Revisi'], ['revisedAt', 'Waktu Revisi'], ['action', 'Aksi']];
  const editableKeys = ['driver', 'plate', 'tenera', 'dura'];
  el('sheetTDTable').innerHTML = `<thead><tr>${cols.map((col) => `<th>${col[1]}</th>`).join('')}</tr></thead><tbody>${rows.map((row) => {
    const info = dt(row.createdAt);
    const values = {
      date: info.date,
      time: info.time,
      driver: row.driver,
      plate: row.plate,
      tenera: row.tenera,
      dura: row.dura,
      total: row.total,
      pctTenera: fixed(row.pctTenera),
      pctDura: fixed(row.pctDura),
      revised: row.revised ? 'Ya' : '-',
      revisedAt: row.revisedAt ? `${dt(row.revisedAt).date} ${dt(row.revisedAt).time}` : '-'
    };
    return `<tr data-id="${row.id}">${cols.map(([key]) => {
      if (key === 'action') return `<td><button type="button" class="text-btn danger" data-delete-td="${row.id}">Hapus</button></td>`;
      const editable = editableKeys.includes(key);
      return `<td ${editable ? `class="editable" contenteditable="true" data-key="${key}"` : ''}>${escapeHtml(values[key])}</td>`;
    }).join('')}</tr>`;
  }).join('')}</tbody>`;
}

function renderPerformance() {
  const mode = el('performanceMode').value;
  const view = el('performanceView').value;
  const head = el('performanceHead');
  const body = el('performanceBody');
  if (mode === 'grading' && view === 'supplier') {
    const rows = supplierStats();
    head.innerHTML = '<tr><th>Ranking</th><th>Supplier</th><th>Transaksi</th><th>Total Janjang</th><th>% Masak</th><th>Potongan</th><th>Max</th><th>Min</th></tr>';
    body.innerHTML = rows.map((row, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(row.name)}</td><td>${row.count}</td><td>${row.totalJanjang}</td><td>${pct(row.avgMasak)}</td><td>${pct(row.avgDed)}</td><td>${pct(row.maxDed)}</td><td>${pct(row.minDed)}</td></tr>`).join('') || '<tr><td colspan="8">Tidak ada data.</td></tr>';
    return;
  }
  if (mode === 'grading' && view === 'driver') {
    const rows = driverStats();
    head.innerHTML = '<tr><th>Ranking</th><th>Sopir</th><th>Transaksi</th><th>Total Janjang</th><th>% Masak</th><th>Potongan</th><th>Supplier Terbanyak</th></tr>';
    body.innerHTML = rows.map((row, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(row.name)}</td><td>${row.count}</td><td>${row.totalJanjang}</td><td>${pct(row.avgMasak)}</td><td>${pct(row.avgDed)}</td><td>${escapeHtml(row.topSupplier)}</td></tr>`).join('') || '<tr><td colspan="7">Tidak ada data.</td></tr>';
    return;
  }
  if (mode === 'td' && view === 'driver') {
    const rows = tdDriverStats();
    head.innerHTML = '<tr><th>Ranking</th><th>Sopir</th><th>Data TD</th><th>Total TD</th><th>% Tenera</th><th>% Dura</th><th>Plat Terbanyak</th></tr>';
    body.innerHTML = rows.map((row, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(row.name)}</td><td>${row.count}</td><td>${row.total}</td><td>${pct(row.avgTenera)}</td><td>${pct(row.avgDura)}</td><td>${escapeHtml(row.topPlate)}</td></tr>`).join('') || '<tr><td colspan="7">Tidak ada data.</td></tr>';
    return;
  }
  if (mode === 'td' && view === 'supplier') {
    head.innerHTML = '<tr><th>Ranking</th><th>Supplier</th><th>Data TD</th><th>Total TD</th><th>% Tenera</th><th>% Dura</th><th>Keterangan</th></tr>';
    body.innerHTML = supplierStats().map((row, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(row.name)}</td><td>-</td><td>-</td><td>-</td><td>-</td><td>TD tidak terkait langsung supplier</td></tr>`).join('') || '<tr><td colspan="7">Tidak ada data.</td></tr>';
    return;
  }
  if (view === 'supplier') {
    const rows = supplierStats().map((row) => ({ name: row.name, gradingScore: 100 - row.avgDed, tdScore: 0, combined: 100 - row.avgDed })).sort((a, b) => b.combined - a.combined);
    head.innerHTML = '<tr><th>Ranking</th><th>Supplier</th><th>Score Grading</th><th>Score TD</th><th>Score Gabungan</th><th>Ringkasan</th></tr>';
    body.innerHTML = rows.map((row, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(row.name)}</td><td>${fixed(row.gradingScore)}</td><td>-</td><td>${fixed(row.combined)}</td><td>Gabungan didominasi grading</td></tr>`).join('') || '<tr><td colspan="6">Tidak ada data.</td></tr>';
    return;
  }
  const gradeRows = driverStats();
  const tdRows = tdDriverStats();
  const rows = gradeRows.map((gradingRow) => {
    const tdRow = tdRows.find((item) => item.name === gradingRow.name);
    const gradingScore = 100 - gradingRow.avgDed;
    const tdScore = tdRow ? tdRow.avgTenera : 0;
    return { name: gradingRow.name, gradingScore, tdScore, combined: (gradingScore + tdScore) / 2, count: gradingRow.count, tdCount: tdRow?.count || 0 };
  }).sort((a, b) => b.combined - a.combined);
  head.innerHTML = '<tr><th>Ranking</th><th>Sopir</th><th>Score Grading</th><th>Score TD</th><th>Score Gabungan</th><th>Transaksi</th><th>Data TD</th></tr>';
  body.innerHTML = rows.map((row, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(row.name)}</td><td>${fixed(row.gradingScore)}</td><td>${fixed(row.tdScore)}</td><td>${fixed(row.combined)}</td><td>${row.count}</td><td>${row.tdCount}</td></tr>`).join('') || '<tr><td colspan="7">Tidak ada data.</td></tr>';
}

function renderAnalytics() {
  el('analyticsCauses').innerHTML = Object.entries(causeTotals()).sort((a, b) => b[1] - a[1]).map(([key, value]) => stat(key, `Akumulasi potongan: ${pct(value)}`, value)).join('') || '<div class="empty-state">Belum ada data.</div>';
  el('analyticsInsights').innerHTML = insights().map((text) => `<div class="stat">${escapeHtml(text)}</div>`).join('');
}

function renderSuppliers() {
  const wrapper = el('supplierList');
  if (!state.suppliers.length) {
    wrapper.innerHTML = '<div class="empty-state">Belum ada supplier di Firebase.</div>';
    return;
  }
  wrapper.innerHTML = state.suppliers.map((supplier) => `<div class="supplier-item"><div><strong>${escapeHtml(supplier.name)}</strong><br><span class="mini-badge ${supplier.status === 'inactive' ? 'inactive' : ''}">${supplier.status === 'inactive' ? 'Nonaktif' : 'Aktif'}</span></div><div><button type="button" class="text-btn" data-edit-supplier="${supplier.id}">Edit</button><button type="button" class="text-btn" data-toggle-supplier="${supplier.id}">${supplier.status === 'inactive' ? 'Aktifkan' : 'Nonaktifkan'}</button></div></div>`).join('');
}

function renderLoadingStates() {
  if (state.loading.suppliers || state.loading.grading || state.loading.td) {
    const messages = [];
    if (state.loading.suppliers) messages.push('supplier');
    if (state.loading.grading) messages.push('grading');
    if (state.loading.td) messages.push('tenera dura');
    setStatus(`Memuat data ${messages.join(', ')} dari Firebase...`, 'info');
  } else if (!el('appStatus').classList.contains('error')) {
    setStatus('');
  }
}

function renderAll() {
  fillStatic();
  updateUserUI();
  renderLoadingStates();
  renderSummaryCards();
  renderDashboard();
  renderGradingLive();
  renderTDLive();
  renderRekapGrading();
  renderRekapTD();
  renderRekapData();
  renderSheetGrading();
  renderSheetTD();
  renderPerformance();
  renderAnalytics();
  renderSuppliers();
  switchPage(state.activePage);
}

async function handleGradingSubmit(event) {
  event.preventDefault();
  const data = Object.fromEntries(new FormData(event.currentTarget).entries());
  const calc = calculateGrading(data);
  renderGradingLive();
  if (calc.validation.type === 'error') return;
  if (!data.supplier) {
    setStatus('Supplier wajib dipilih sebelum menyimpan grading.', 'warning');
    return;
  }
  try {
    const node = push(ref(db, 'grading'));
    await set(node, {
      ...data,
      ...calc,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      revised: false,
      revisedAt: null,
      createdBy: state.currentUser?.email || '-'
    });
    event.currentTarget.reset();
    event.currentTarget.elements.totalBunches.value = 0;
    event.currentTarget.querySelectorAll('.cat').forEach((input) => { input.value = 0; });
    el('driverHint').textContent = 'Belum ada histori sopir.';
    renderGradingLive();
    setStatus('Data grading berhasil disimpan ke Firebase.', 'info');
  } catch (error) {
    setStatus(`Gagal menyimpan grading: ${error.message}`, 'error');
  }
}

async function handleTDSubmit(event) {
  event.preventDefault();
  const data = Object.fromEntries(new FormData(event.currentTarget).entries());
  const calc = calculateTD(data);
  try {
    const node = push(ref(db, 'td'));
    await set(node, {
      ...data,
      ...calc,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      revised: false,
      revisedAt: null,
      createdBy: state.currentUser?.email || '-'
    });
    event.currentTarget.reset();
    event.currentTarget.elements.tenera.value = 0;
    event.currentTarget.elements.dura.value = 0;
    renderTDLive();
    setStatus('Data Tenera Dura berhasil disimpan ke Firebase.', 'info');
  } catch (error) {
    setStatus(`Gagal menyimpan Tenera Dura: ${error.message}`, 'error');
  }
}

async function handleSupplierSubmit(event) {
  event.preventDefault();
  if (state.currentRole !== 'staff') return;
  const form = event.currentTarget;
  const supplierId = form.elements.supplierId.value;
  const name = form.elements.supplierName.value.trim();
  const status = form.elements.supplierStatus.value;
  if (!name) {
    setStatus('Nama supplier wajib diisi.', 'warning');
    return;
  }
  try {
    if (supplierId) {
      await update(ref(db, `suppliers/${supplierId}`), { name, status, updatedAt: new Date().toISOString(), updatedBy: state.currentUser?.email || '-' });
    } else {
      if (supplierExistsByName(name)) {
        setStatus('Nama supplier sudah ada.', 'warning');
        return;
      }
      const node = push(ref(db, 'suppliers'));
      await set(node, { name, status, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(), createdBy: state.currentUser?.email || '-' });
    }
    form.reset();
    form.elements.supplierId.value = '';
    form.elements.supplierStatus.value = 'active';
    setStatus('Supplier berhasil disimpan ke Firebase.', 'info');
  } catch (error) {
    setStatus(`Gagal menyimpan supplier: ${error.message}`, 'error');
  }
}

async function updateGradingRow(id, patch) {
  const current = state.grading.find((row) => row.id === id);
  if (!current) return;
  const merged = { ...current, ...patch };
  const calc = calculateGrading(merged);
  if (calc.validation.type === 'error') {
    setStatus(calc.validation.message, 'error');
    return;
  }
  await update(ref(db, `grading/${id}`), { ...patch, ...calc, revised: true, revisedAt: new Date().toISOString(), updatedAt: new Date().toISOString(), updatedBy: state.currentUser?.email || '-' });
}

async function updateTDRow(id, patch) {
  const current = state.td.find((row) => row.id === id);
  if (!current) return;
  const merged = { ...current, ...patch };
  const calc = calculateTD(merged);
  await update(ref(db, `td/${id}`), { ...patch, ...calc, revised: true, revisedAt: new Date().toISOString(), updatedAt: new Date().toISOString(), updatedBy: state.currentUser?.email || '-' });
}

function handleSheetFocusOut(event) {
  const cell = event.target.closest('td.editable');
  if (!cell || state.currentRole !== 'staff') return;
  const row = cell.closest('tr');
  const id = row?.dataset.id;
  const key = cell.dataset.key;
  if (!id || !key) return;
  const value = cell.textContent.trim();
  if (row.closest('#sheetGradingTable')) {
    const numericKeys = ['totalBunches', 'mentah', 'mengkal', 'overripe', 'busuk', 'kosong', 'partheno', 'tikus'];
    const patch = { [key]: numericKeys.includes(key) ? num(value) : value };
    updateGradingRow(id, patch).catch((error) => setStatus(`Gagal memperbarui grading: ${error.message}`, 'error'));
    return;
  }
  if (row.closest('#sheetTDTable')) {
    const numericKeys = ['tenera', 'dura'];
    const patch = { [key]: numericKeys.includes(key) ? num(value) : value };
    updateTDRow(id, patch).catch((error) => setStatus(`Gagal memperbarui TD: ${error.message}`, 'error'));
  }
}

async function deleteGrading(id) {
  if (state.currentRole !== 'staff') return;
  if (!window.confirm('Hapus data grading ini?')) return;
  try {
    await remove(ref(db, `grading/${id}`));
    setStatus('Data grading berhasil dihapus.', 'info');
  } catch (error) {
    setStatus(`Gagal menghapus grading: ${error.message}`, 'error');
  }
}

async function deleteTD(id) {
  if (state.currentRole !== 'staff') return;
  if (!window.confirm('Hapus data Tenera Dura ini?')) return;
  try {
    await remove(ref(db, `td/${id}`));
    setStatus('Data Tenera Dura berhasil dihapus.', 'info');
  } catch (error) {
    setStatus(`Gagal menghapus TD: ${error.message}`, 'error');
  }
}

function exportTable(rows, headers, filename) {
  const table = `<table><thead><tr>${headers.map((header) => `<th>${header}</th>`).join('')}</tr></thead><tbody>${rows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  const blob = new Blob([`<html><head><meta charset="utf-8"></head><body>${table}</body></html>`], { type: 'application/vnd.ms-excel' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function exportGrading() {
  const rows = filterDate(state.grading, el('rekapGradingStart').value, el('rekapGradingEnd').value)
    .filter((row) => {
      const q = normalizeText(el('rekapGradingSearch').value);
      const supplier = el('rekapGradingSupplier').value;
      return (!q || `${row.driver} ${row.plate} ${row.supplier}`.toLowerCase().includes(q)) && (!supplier || row.supplier === supplier);
    })
    .map((row) => {
      const info = dt(row.createdAt);
      return [info.date, info.time, row.driver, row.plate, row.supplier, row.totalBunches, row.mentah, row.mengkal, row.overripe, row.busuk, row.kosong, row.partheno, row.tikus, fixed(row.percentages?.masak), fixed(row.totalDeduction), row.revised ? 'Ya' : '-', row.revisedAt ? `${dt(row.revisedAt).date} ${dt(row.revisedAt).time}` : '-'];
    });
  exportTable(rows, ['Tanggal', 'Jam', 'Sopir', 'Plat', 'Supplier', 'Total Janjang', 'Mentah', 'Mengkal', 'Overripe', 'Busuk', 'Kosong', 'Partheno', 'Tikus', '% Masak', 'Total Potongan', 'Revisi', 'Waktu Revisi'], 'rekap_grading.xls');
}

function exportTD() {
  const rows = filterDate(state.td, el('rekapTDStart').value, el('rekapTDEnd').value)
    .filter((row) => !normalizeText(el('rekapTDSearch').value) || `${row.driver} ${row.plate}`.toLowerCase().includes(normalizeText(el('rekapTDSearch').value)))
    .map((row) => {
      const info = dt(row.createdAt);
      return [info.date, info.time, row.driver, row.plate, row.tenera, row.dura, row.total, fixed(row.pctTenera), fixed(row.pctDura), row.revised ? 'Ya' : '-', row.revisedAt ? `${dt(row.revisedAt).date} ${dt(row.revisedAt).time}` : '-'];
    });
  exportTable(rows, ['Tanggal', 'Jam', 'Sopir', 'Plat', 'Tenera', 'Dura', 'Total TD', '% Tenera', '% Dura', 'Revisi', 'Waktu Revisi'], 'rekap_td.xls');
}

function openWaModal(type, title) {
  state.waContext = type;
  el('waModalTitle').textContent = title;
  el('waTypeText').value = title;
  if (type.startsWith('grading')) {
    el('waStart').value = el('rekapGradingStart').value;
    el('waEnd').value = el('rekapGradingEnd').value;
    el('waSupplier').value = el('rekapGradingSupplier').value;
    el('waSupplierWrap').classList.remove('hidden');
  } else if (type.startsWith('td')) {
    el('waStart').value = el('rekapTDStart').value;
    el('waEnd').value = el('rekapTDEnd').value;
    el('waSupplier').value = '';
    el('waSupplierWrap').classList.add('hidden');
  }
  updateWaPreview();
  el('waModal').classList.add('open');
}

function rowsForWA() {
  const start = el('waStart').value;
  const end = el('waEnd').value;
  const supplier = el('waSupplier').value;
  const mode = el('waMode').value;
  if (!state.waContext) return { start, end, supplier, mode, rows: [] };
  if (state.waContext.startsWith('grading')) {
    let rows = filterDate(state.grading, start, end);
    if (supplier) rows = rows.filter((row) => row.supplier === supplier);
    return { start, end, supplier, mode, rows };
  }
  return { start, end, supplier: '', mode, rows: filterDate(state.td, start, end) };
}

function buildWAText() {
  const { start, end, supplier, mode, rows } = rowsForWA();
  const period = `${start || '-'} s/d ${end || '-'}`;
  const filterLabel = supplier || 'Semua Supplier';
  const type = state.waContext || '';
  if (type === 'grading-summary' || type === 'grading-detail') {
    if (mode === 'driver') {
      const grouped = driverStats(rows);
      const details = grouped.map((row, index) => `${index + 1}. ${row.name}\n   Transaksi: ${row.count}\n   Total Janjang: ${row.totalJanjang}\n   % Masak: ${pct(row.avgMasak)}\n   Potongan: ${pct(row.avgDed)}`).join('\n');
      return `🌴 REKAP GRADING\nPeriode: ${period}\nSupplier: ${filterLabel}\nMode: Per Sopir\nTotal Transaksi: ${rows.length}\n\n${details || 'Tidak ada data.'}`;
    }
    if (type === 'grading-summary') {
      const avgMasak = rows.length ? rows.reduce((sum, row) => sum + num(row.percentages?.masak), 0) / rows.length : 0;
      const avgDed = rows.length ? rows.reduce((sum, row) => sum + num(row.totalDeduction), 0) / rows.length : 0;
      return `🌴 RINGKASAN GRADING\nPeriode: ${period}\nSupplier: ${filterLabel}\nMode: Keseluruhan\nTransaksi: ${rows.length}\nTotal Janjang: ${rows.reduce((sum, row) => sum + num(row.totalBunches), 0)}\nRata-rata % Masak: ${pct(avgMasak)}\nRata-rata Potongan: ${pct(avgDed)}`;
    }
    return `🌴 DETAIL GRADING\nPeriode: ${period}\nSupplier: ${filterLabel}\nMode: Keseluruhan\n\n${rows.map((row, index) => `${index + 1}. ${row.driver} | ${row.plate} | ${row.supplier} | Janjang ${row.totalBunches} | % Masak ${pct(row.percentages?.masak)} | Potongan ${pct(row.totalDeduction)}`).join('\n') || 'Tidak ada data.'}`;
  }
  if (mode === 'driver') {
    const grouped = tdDriverStats(rows);
    const details = grouped.map((row, index) => `${index + 1}. ${row.name}\n   Data TD: ${row.count}\n   Total TD: ${row.total}\n   % Tenera: ${pct(row.avgTenera)}\n   % Dura: ${pct(row.avgDura)}`).join('\n');
    return `🌴 REKAP TENERA DURA\nPeriode: ${period}\nMode: Per Sopir\nTotal Transaksi: ${rows.length}\n\n${details || 'Tidak ada data.'}`;
  }
  if (type === 'td-summary') {
    const avgT = rows.length ? rows.reduce((sum, row) => sum + num(row.pctTenera), 0) / rows.length : 0;
    const avgD = rows.length ? rows.reduce((sum, row) => sum + num(row.pctDura), 0) / rows.length : 0;
    return `🌴 RINGKASAN TENERA DURA\nPeriode: ${period}\nMode: Keseluruhan\nTransaksi: ${rows.length}\nTotal TD: ${rows.reduce((sum, row) => sum + num(row.total), 0)}\nRata-rata % Tenera: ${pct(avgT)}\nRata-rata % Dura: ${pct(avgD)}`;
  }
  return `🌴 DETAIL TENERA DURA\nPeriode: ${period}\nMode: Keseluruhan\n\n${rows.map((row, index) => `${index + 1}. ${row.driver} | ${row.plate} | Tenera ${row.tenera} | Dura ${row.dura} | % Tenera ${pct(row.pctTenera)} | % Dura ${pct(row.pctDura)}`).join('\n') || 'Tidak ada data.'}`;
}

function updateWaPreview() { el('waPreview').textContent = buildWAText(); }
function openWhatsApp() { window.open(`https://wa.me/?text=${encodeURIComponent(buildWAText())}`, '_blank'); }

function showDetail(type, id) {
  const modal = el('detailModal');
  const body = el('detailBody');
  if (type === 'grading') {
    const row = state.grading.find((item) => item.id === id);
    if (!row) return;
    const info = dt(row.createdAt);
    body.innerHTML = `
      <div class="block-title">Informasi Umum</div>
      <div class="detail-grid">
        <div class="detail-box"><span>Tanggal</span><strong>${info.date}</strong></div>
        <div class="detail-box"><span>Jam</span><strong>${info.time}</strong></div>
        <div class="detail-box"><span>Supplier</span><strong>${escapeHtml(row.supplier)}</strong></div>
        <div class="detail-box"><span>Sopir</span><strong>${escapeHtml(row.driver)}</strong></div>
        <div class="detail-box"><span>Plat</span><strong>${escapeHtml(row.plate)}</strong></div>
        <div class="detail-box"><span>Total Janjang</span><strong>${row.totalBunches}</strong></div>
        <div class="detail-box"><span>Masak</span><strong>${row.masak}</strong></div>
        <div class="detail-box"><span>% Masak</span><strong>${pct(row.percentages?.masak)}</strong></div>
        <div class="detail-box"><span>Total Potongan</span><strong>${pct(row.totalDeduction)}</strong></div>
      </div>`;
  } else {
    const row = state.td.find((item) => item.id === id);
    if (!row) return;
    const info = dt(row.createdAt);
    body.innerHTML = `
      <div class="block-title">Informasi Umum</div>
      <div class="detail-grid">
        <div class="detail-box"><span>Tanggal</span><strong>${info.date}</strong></div>
        <div class="detail-box"><span>Jam</span><strong>${info.time}</strong></div>
        <div class="detail-box"><span>Sopir</span><strong>${escapeHtml(row.driver)}</strong></div>
        <div class="detail-box"><span>Plat</span><strong>${escapeHtml(row.plate)}</strong></div>
        <div class="detail-box"><span>Tenera</span><strong>${row.tenera}</strong></div>
        <div class="detail-box"><span>Dura</span><strong>${row.dura}</strong></div>
        <div class="detail-box"><span>Total TD</span><strong>${row.total}</strong></div>
        <div class="detail-box"><span>% Tenera</span><strong>${pct(row.pctTenera)}</strong></div>
        <div class="detail-box"><span>% Dura</span><strong>${pct(row.pctDura)}</strong></div>
      </div>`;
  }
  modal.classList.add('open');
}

function bindEvents() {
  if (state.listenersBound) return;
  state.listenersBound = true;
  setupRolePicker();
  el('loginForm').addEventListener('submit', handleLoginSubmit);
  el('logoutBtn').addEventListener('click', () => signOut(auth));
  document.querySelectorAll('.menu-item').forEach((button) => button.addEventListener('click', () => switchPage(button.dataset.page)));
  el('menuToggle').addEventListener('click', () => toggleSidebar(!el('app').classList.contains('sidebar-open')));
  el('mobileOverlay').addEventListener('click', () => toggleSidebar(false));
  window.addEventListener('resize', () => { if (window.innerWidth > 900) toggleSidebar(false); });

  el('gradingDriver').addEventListener('input', (event) => { historyHint(event.target.value.trim()); renderGradingLive(); });
  el('gradingForm').addEventListener('input', renderGradingLive);
  el('tdForm').addEventListener('input', renderTDLive);
  el('gradingForm').addEventListener('submit', handleGradingSubmit);
  el('tdForm').addEventListener('submit', handleTDSubmit);

  el('resetGradingBtn').addEventListener('click', () => {
    const form = el('gradingForm');
    form.reset();
    form.elements.totalBunches.value = 0;
    form.querySelectorAll('.cat').forEach((input) => { input.value = 0; });
    el('driverHint').textContent = 'Belum ada histori sopir.';
    renderGradingLive();
  });
  el('resetTDBtn').addEventListener('click', () => {
    const form = el('tdForm');
    form.reset();
    form.elements.tenera.value = 0;
    form.elements.dura.value = 0;
    renderTDLive();
  });
  el('copyLastGrading').addEventListener('click', () => {
    const last = state.grading[0];
    if (!last) return;
    const form = el('gradingForm');
    form.elements.driver.value = last.driver || '';
    form.elements.plate.value = last.plate || '';
    form.elements.supplier.value = last.supplier || '';
    ['totalBunches', 'mentah', 'mengkal', 'overripe', 'busuk', 'kosong', 'partheno', 'tikus'].forEach((key) => { form.elements[key].value = last[key] ?? 0; });
    historyHint(last.driver);
    renderGradingLive();
  });
  el('copyLastTD').addEventListener('click', () => {
    const last = state.td[0];
    if (!last) return;
    const form = el('tdForm');
    form.elements.driver.value = last.driver || '';
    form.elements.plate.value = last.plate || '';
    form.elements.tenera.value = last.tenera ?? 0;
    form.elements.dura.value = last.dura ?? 0;
    renderTDLive();
  });

  ['rekapGradingSearch'].forEach((id) => el(id).addEventListener('input', renderRekapGrading));
  ['rekapGradingSupplier', 'rekapGradingStart', 'rekapGradingEnd'].forEach((id) => el(id).addEventListener('change', renderRekapGrading));
  ['rekapTDSearch'].forEach((id) => el(id).addEventListener('input', renderRekapTD));
  ['rekapTDStart', 'rekapTDEnd'].forEach((id) => el(id).addEventListener('change', renderRekapTD));
  el('rekapDataRunBtn').addEventListener('click', renderRekapData);
  el('rekapDataResetBtn').addEventListener('click', () => {
    el('rekapDataStart').value = '';
    el('rekapDataEnd').value = '';
    el('rekapDataSupplier').value = '';
    el('rekapDataDriver').value = '';
    renderRekapData();
  });
  el('sheetGradingSearch').addEventListener('input', renderSheetGrading);
  el('sheetTDSearch').addEventListener('input', renderSheetTD);
  el('performanceMode').addEventListener('change', renderPerformance);
  el('performanceView').addEventListener('change', renderPerformance);
  el('supplierForm').addEventListener('submit', handleSupplierSubmit);
  el('resetSupplierBtn').addEventListener('click', () => {
    const form = el('supplierForm');
    form.reset();
    form.elements.supplierId.value = '';
    form.elements.supplierStatus.value = 'active';
  });

  document.addEventListener('click', (event) => {
    const gradingDelete = event.target.closest('[data-delete-grading]');
    if (gradingDelete) { deleteGrading(gradingDelete.dataset.deleteGrading); return; }
    const tdDelete = event.target.closest('[data-delete-td]');
    if (tdDelete) { deleteTD(tdDelete.dataset.deleteTd); return; }
    const editSupplier = event.target.closest('[data-edit-supplier]');
    if (editSupplier) {
      const row = state.suppliers.find((supplier) => supplier.id === editSupplier.dataset.editSupplier);
      if (!row) return;
      const form = el('supplierForm');
      form.elements.supplierId.value = row.id;
      form.elements.supplierName.value = row.name || '';
      form.elements.supplierStatus.value = row.status || 'active';
      switchPage('supplier');
      return;
    }
    const toggleSupplier = event.target.closest('[data-toggle-supplier]');
    if (toggleSupplier) {
      const row = state.suppliers.find((supplier) => supplier.id === toggleSupplier.dataset.toggleSupplier);
      if (!row || state.currentRole !== 'staff') return;
      update(ref(db, `suppliers/${row.id}`), { status: row.status === 'inactive' ? 'active' : 'inactive', updatedAt: new Date().toISOString(), updatedBy: state.currentUser?.email || '-' }).catch((error) => setStatus(`Gagal memperbarui supplier: ${error.message}`, 'error'));
      return;
    }
    const detailRow = event.target.closest('[data-detail-id]');
    if (detailRow && !event.target.closest('button')) {
      showDetail(detailRow.dataset.detailType, detailRow.dataset.detailId);
      return;
    }
  });

  document.addEventListener('focusout', handleSheetFocusOut, true);
  el('closeModalBtn').addEventListener('click', () => el('detailModal').classList.remove('open'));
  el('closeWaModalBtn').addEventListener('click', () => el('waModal').classList.remove('open'));
  el('detailModal').addEventListener('click', (event) => { if (event.target === el('detailModal')) el('detailModal').classList.remove('open'); });
  el('waModal').addEventListener('click', (event) => { if (event.target === el('waModal')) el('waModal').classList.remove('open'); });
  document.addEventListener('keydown', (event) => { if (event.key === 'Escape') { el('detailModal').classList.remove('open'); el('waModal').classList.remove('open'); toggleSidebar(false); } });
  el('confirmWaBtn').addEventListener('click', openWhatsApp);
  ['waStart', 'waEnd', 'waSupplier', 'waMode'].forEach((id) => el(id).addEventListener('input', updateWaPreview));
  ['waStart', 'waEnd', 'waSupplier', 'waMode'].forEach((id) => el(id).addEventListener('change', updateWaPreview));
  el('sendGradingSummaryBtn').addEventListener('click', () => openWaModal('grading-summary', 'Kirim Ringkasan Grading'));
  el('sendGradingDetailBtn').addEventListener('click', () => openWaModal('grading-detail', 'Kirim Detail Grading'));
  el('sendTDSummaryBtn').addEventListener('click', () => openWaModal('td-summary', 'Kirim Ringkasan Tenera Dura'));
  el('sendTDDetailBtn').addEventListener('click', () => openWaModal('td-detail', 'Kirim Detail Tenera Dura'));
  el('waFilteredGradingBtn').addEventListener('click', () => openWaModal('grading-summary', 'Kirim Ringkasan Grading'));
  el('waFilteredTDBtn').addEventListener('click', () => openWaModal('td-summary', 'Kirim Ringkasan Tenera Dura'));
  el('exportGradingBtn').addEventListener('click', exportGrading);
  el('exportTDBtn').addEventListener('click', exportTD);
  el('globalSearch').addEventListener('input', (event) => {
    const value = event.target.value;
    el('rekapGradingSearch').value = value;
    el('rekapTDSearch').value = value;
    el('sheetGradingSearch').value = value;
    el('sheetTDSearch').value = value;
    renderRekapGrading();
    renderRekapTD();
    renderSheetGrading();
    renderSheetTD();
  });
}

bindEvents();

onAuthStateChanged(auth, async (user) => {
  state.currentUser = user;
  if (!user) {
    state.currentRole = 'grading';
    state.loading = { suppliers: false, grading: false, td: false };
    el('app').classList.add('hidden');
    el('loginScreen').classList.remove('hidden');
    toggleSidebar(false);
    showLoginError('');
    return;
  }
  if (!isStaffEmail(user.email) && !isGradingEmail(user.email)) {
    await signOut(auth);
    showLoginError('Email login tidak dikenali sebagai staff atau grading.');
    return;
  }
  state.currentRole = deriveRole(user.email || '');
  state.loading = { suppliers: true, grading: true, td: true };
  el('app').classList.remove('hidden');
  el('loginScreen').classList.add('hidden');
  bindRealtime();
  try {
    await ensureDefaultSuppliers();
  } catch (error) {
    setStatus(`Gagal menyiapkan supplier awal: ${error.message}`, 'error');
  }
  renderAll();
});
