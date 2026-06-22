<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<title>Expense Claim</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

:root {
  --bg: #F0F4F8;
  --surface: #FFFFFF;
  --surface2: #F7FAFC;
  --border: #CBD5E0;
  --focus: #3182CE;
  --accent: #2B6CB0;
  --accent-light: #EBF4FF;
  --text: #1A202C;
  --muted: #718096;
  --success: #276749;
  --danger: #C53030;
  --row-even: #F7FAFC;
}

* { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }

body {
  font-family: 'IBM Plex Sans Thai', sans-serif;
  background: var(--bg);
  color: var(--text);
  min-height: 100vh;
  font-size: 15px;
}

/* Header */
header {
  background: var(--accent);
  color: white;
  padding: 14px 16px;
  position: sticky;
  top: 0;
  z-index: 100;
  box-shadow: 0 2px 8px rgba(0,0,0,0.2);
}
header h1 { font-size: 16px; font-weight: 600; }
header .sub { font-size: 11px; opacity: 0.75; margin-top: 1px; }

/* Tabs */
.tabs {
  display: flex;
  background: white;
  border-bottom: 2px solid var(--border);
  position: sticky;
  top: 56px;
  z-index: 99;
}
.tab {
  flex: 1;
  padding: 12px 8px;
  text-align: center;
  font-size: 13px;
  font-weight: 500;
  color: var(--muted);
  cursor: pointer;
  border-bottom: 3px solid transparent;
  margin-bottom: -2px;
  transition: color 0.15s;
  user-select: none;
}
.tab.active { color: var(--accent); border-bottom-color: var(--accent); }

/* Pages */
.page { display: none; padding: 16px; }
.page.active { display: block; }

/* Card */
.card {
  background: white;
  border-radius: 12px;
  border: 1px solid var(--border);
  margin-bottom: 14px;
  overflow: hidden;
}
.card-title {
  padding: 12px 16px;
  font-size: 13px;
  font-weight: 600;
  color: var(--accent);
  background: var(--accent-light);
  border-bottom: 1px solid var(--border);
}

/* Field */
.field { padding: 12px 16px; border-bottom: 1px solid #F0F4F8; }
.field:last-child { border-bottom: none; }
.field label {
  display: block;
  font-size: 11px;
  font-weight: 600;
  color: var(--muted);
  text-transform: uppercase;
  letter-spacing: 0.7px;
  margin-bottom: 6px;
}
.field input, .field select {
  width: 100%;
  border: 1.5px solid var(--border);
  border-radius: 8px;
  padding: 11px 13px;
  font-family: inherit;
  font-size: 15px;
  color: var(--text);
  background: var(--surface2);
  appearance: none;
  -webkit-appearance: none;
}
.field input:focus, .field select:focus {
  outline: none;
  border-color: var(--focus);
  box-shadow: 0 0 0 3px rgba(49,130,206,0.12);
  background: white;
}
.field input[type="number"] {
  font-family: 'IBM Plex Mono', monospace;
  font-size: 16px;
}

.row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }

/* Buttons */
.btn-primary {
  width: 100%;
  background: var(--accent);
  color: white;
  border: none;
  border-radius: 10px;
  padding: 14px;
  font-family: inherit;
  font-size: 16px;
  font-weight: 600;
  cursor: pointer;
  margin-top: 4px;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  -webkit-appearance: none;
}
.btn-primary:active { opacity: 0.85; transform: scale(0.99); }

.btn-success {
  width: 100%;
  background: var(--success);
  color: white;
  border: none;
  border-radius: 10px;
  padding: 14px;
  font-family: inherit;
  font-size: 16px;
  font-weight: 600;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  -webkit-appearance: none;
}
.btn-success:active { opacity: 0.85; }

/* Row cards in list */
.row-card {
  background: white;
  border-radius: 10px;
  border: 1px solid var(--border);
  margin-bottom: 10px;
  overflow: hidden;
}
.row-card-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 10px 14px;
  background: var(--surface2);
  border-bottom: 1px solid var(--border);
}
.row-card-date {
  font-family: 'IBM Plex Mono', monospace;
  font-size: 12px;
  color: var(--muted);
}
.row-card-body { padding: 10px 14px; }
.row-card-body .customer { font-weight: 600; font-size: 14px; margin-bottom: 2px; }
.row-card-body .site { font-size: 12px; color: var(--muted); margin-bottom: 8px; }
.row-nums {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
}
.num-chip {
  background: var(--accent-light);
  color: var(--accent);
  border-radius: 6px;
  padding: 3px 8px;
  font-size: 11px;
  font-family: 'IBM Plex Mono', monospace;
  font-weight: 500;
}
.num-chip.zero { background: #EDF2F7; color: var(--muted); }

.delete-btn {
  background: none;
  border: 1px solid #FEB2B2;
  color: var(--danger);
  border-radius: 6px;
  padding: 5px 10px;
  font-size: 12px;
  cursor: pointer;
  font-family: inherit;
}

.status-badge {
  display: inline-block;
  padding: 2px 9px;
  border-radius: 10px;
  font-size: 11px;
  font-weight: 500;
}
.s-complete { background: #F0FFF4; color: var(--success); }
.s-pending  { background: #FEFCBF; color: #744210; }
.s-empty    { background: #EDF2F7; color: var(--muted); }

/* Summary */
.sum-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; padding: 14px; }
.sum-item { background: var(--surface2); border-radius: 8px; padding: 12px; border: 1px solid var(--border); }
.sum-item .lbl { font-size: 10px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: 0.6px; margin-bottom: 4px; }
.sum-item .val { font-family: 'IBM Plex Mono', monospace; font-size: 20px; font-weight: 700; color: var(--accent); }
.sum-item .unit { font-size: 10px; color: var(--muted); margin-top: 2px; }

.empty-state { text-align: center; padding: 40px 20px; color: var(--muted); }
.empty-state .icon { font-size: 40px; margin-bottom: 10px; }

/* iOS safe area */
.bottom-safe { padding-bottom: env(safe-area-inset-bottom, 16px); }

/* Download hint */
.hint {
  background: #FEFCBF;
  border: 1px solid #F6E05E;
  border-radius: 8px;
  padding: 10px 14px;
  font-size: 12px;
  color: #744210;
  margin-bottom: 14px;
  line-height: 1.5;
}
</style>
</head>
<body>

<header>
  <h1>📋 Service Expenses Claim & OT</h1>
  <div class="sub" style="font-weight:600;">v3 — date-fix • 2026-06-22</div>
  <div class="sub">รายงานค่าใช้จ่ายการปฏิบัติงาน</div>
</header>

<div class="tabs">
  <div class="tab active" onclick="switchTab('tab-info')">ข้อมูล</div>
  <div class="tab" onclick="switchTab('tab-add')">➕ เพิ่ม</div>
  <div class="tab" onclick="switchTab('tab-list')">รายการ <span id="badge"></span></div>
  <div class="tab" onclick="switchTab('tab-export')">ดาวน์โหลด</div>
</div>

<!-- TAB 1: Info -->
<div id="tab-info" class="page active bottom-safe">
  <div class="card">
    <div class="card-title">👤 ข้อมูลพนักงาน</div>
    <div class="field">
      <label>ชื่อพนักงาน (Name)</label>
      <input type="text" id="empName" value="Supakorn Thammasa" placeholder="ชื่อ-นามสกุล">
    </div>
    <div class="field">
      <label>FSE Code</label>
      <input type="text" id="fseCode" value="SNG-313" placeholder="เช่น SNG-313">
    </div>
    <div class="field">
      <label>ช่วงเวลารายงาน</label>
      <div class="row2">
        <div>
          <div style="font-size:11px;color:var(--muted);margin-bottom:4px;">From</div>
          <input type="date" id="dateFrom" value="2025-04-01">
        </div>
        <div>
          <div style="font-size:11px;color:var(--muted);margin-bottom:4px;">To</div>
          <input type="date" id="dateTo" value="2025-04-30">
        </div>
      </div>
    </div>
  </div>

  <div class="sum-grid" style="padding:0;margin-bottom:14px;">
    <div class="sum-item">
      <div class="lbl">รายการทั้งหมด</div>
      <div class="val" id="s_count">0</div>
      <div class="unit">รายการ</div>
    </div>
    <div class="sum-item">
      <div class="lbl">ค่าใช้จ่ายรวม</div>
      <div class="val" id="s_total">0</div>
      <div class="unit">บาท</div>
    </div>
    <div class="sum-item">
      <div class="lbl">Mileage รวม</div>
      <div class="val" id="s_km">0</div>
      <div class="unit">km → <span id="s_km_cost">0</span> ฿</div>
    </div>
    <div class="sum-item">
      <div class="lbl">OT รวม</div>
      <div class="val" id="s_ot">0</div>
      <div class="unit">ชั่วโมง</div>
    </div>
  </div>
</div>

<!-- TAB 2: Add -->
<div id="tab-add" class="page bottom-safe">
  <div class="card">
    <div class="card-title">📅 วันเวลา</div>
    <div class="field">
      <label>วันที่เริ่ม (Start Date)</label>
      <input type="date" id="f_startDate">
    </div>
    <div class="field">
      <label>เวลาเริ่ม (Start Time)</label>
      <input type="time" id="f_startTime" value="08:00">
    </div>
    <div class="field">
      <label>วันที่จบ (Finish Date)</label>
      <input type="date" id="f_endDate">
    </div>
    <div class="field">
      <label>เวลาจบ (Finish Time)</label>
      <input type="time" id="f_endTime" value="17:00">
    </div>
    <div class="field">
      <label>Status</label>
      <select id="f_status">
        <option value="">-- เลือก --</option>
        <option value="complete">complete</option>
        <option value="pending">pending</option>
      </select>
    </div>
  </div>

  <div class="card">
    <div class="card-title">📝 ข้อมูลงาน</div>
    <div class="field">
      <label>Service Request ID</label>
      <input type="text" id="f_srId" placeholder="เช่น SR-12345">
    </div>
    <div class="field">
      <label>Sale Order No.</label>
      <input type="text" id="f_soNo" placeholder="เช่น 2024-SM054">
    </div>
    <div class="field">
      <label>Customer (ชื่อลูกค้า)</label>
      <input type="text" id="f_customer" placeholder="บริษัท...">
    </div>
    <div class="field">
      <label>Sitename (สถานที่)</label>
      <input type="text" id="f_site" placeholder="ชื่อ Site / ที่อยู่">
    </div>
  </div>

  <div class="card">
    <div class="card-title">💰 ค่าใช้จ่าย</div>
    <div class="field">
      <label>Mileage (กิโลเมตร)</label>
      <input type="number" id="f_mileage" min="0" value="0" inputmode="decimal">
    </div>
    <div class="field">
      <label>Transport / ค่าเดินทาง (฿)</label>
      <input type="number" id="f_transport" min="0" value="0" inputmode="decimal">
    </div>
    <div class="field">
      <label>Misc. / เบ็ดเตล็ด (฿)</label>
      <input type="number" id="f_misc" min="0" value="0" inputmode="decimal">
    </div>
    <div class="field">
      <label>Hotel / ที่พัก (฿)</label>
      <input type="number" id="f_hotel" min="0" value="0" inputmode="decimal">
    </div>
    <div class="field">
      <label>Meal / ค่าอาหาร (฿)</label>
      <input type="number" id="f_meal" min="0" value="0" inputmode="decimal">
    </div>
    <div class="field">
      <label>Other / อื่นๆ (฿)</label>
      <input type="number" id="f_other" min="0" value="0" inputmode="decimal">
    </div>
  </div>

  <div class="card">
    <div class="card-title">⏱ ชั่วโมงทำงาน (Real Work Time & OT)</div>
    <div class="field">
      <label>Working Hours (ชม.)</label>
      <input type="number" id="f_workHrs" min="0" step="0.5" value="0" inputmode="decimal">
    </div>
    <div class="field">
      <label>Traveling Hours (ชม.)</label>
      <input type="number" id="f_travelHrs" min="0" step="0.5" value="0" inputmode="decimal">
    </div>
    <div class="field">
      <label>OT × 1.0 (ชม.)</label>
      <input type="number" id="f_ot10" min="0" step="0.5" value="0" inputmode="decimal">
    </div>
    <div class="field">
      <label>OT × 1.5 (ชม.)</label>
      <input type="number" id="f_ot15" min="0" step="0.5" value="0" inputmode="decimal">
    </div>
    <div class="field">
      <label>OT × 3.0 (ชม.)</label>
      <input type="number" id="f_ot30" min="0" step="0.5" value="0" inputmode="decimal">
    </div>
  </div>

  <button class="btn-primary" onclick="addRow()">✚ เพิ่มรายการนี้</button>
</div>

<!-- TAB 3: List -->
<div id="tab-list" class="page bottom-safe">
  <div id="listContainer">
    <div class="empty-state">
      <div class="icon">📄</div>
      <div>ยังไม่มีรายการ<br>ไปที่แท็บ "➕ เพิ่ม" เพื่อกรอกข้อมูล</div>
    </div>
  </div>
</div>

<!-- TAB 4: Export -->
<div id="tab-export" class="page bottom-safe">
  <div class="hint">
    📱 <strong>สำหรับมือถือ:</strong> กดปุ่มด้านล่าง ไฟล์จะดาวน์โหลดอัตโนมัติ<br>
    iOS → เปิดใน Files app หรือ Share ได้เลย<br>
    Android → ดูใน Downloads folder
  </div>

  <div class="card">
    <div class="card-title">📊 สรุปก่อนดาวน์โหลด</div>
    <div class="sum-grid">
      <div class="sum-item">
        <div class="lbl">รายการ</div>
        <div class="val" id="e_count">0</div>
        <div class="unit">รายการ</div>
      </div>
      <div class="sum-item">
        <div class="lbl">Mileage</div>
        <div class="val" id="e_km">0</div>
        <div class="unit">km</div>
      </div>
      <div class="sum-item">
        <div class="lbl">ค่าใช้จ่ายรวม</div>
        <div class="val" id="e_total">0</div>
        <div class="unit">บาท</div>
      </div>
      <div class="sum-item">
        <div class="lbl">OT รวม</div>
        <div class="val" id="e_ot">0</div>
        <div class="unit">ชั่วโมง</div>
      </div>
    </div>
  </div>

  <button class="btn-success" onclick="exportToExcel()">
    ⬇ ดาวน์โหลดไฟล์ Excel
  </button>
  <div style="text-align:center;font-size:11px;color:var(--muted);margin-top:10px;">
    ชื่อไฟล์: รายงานค่าใช้จ่าย_[ชื่อ]_[วันที่].xlsx
  </div>
</div>

<script>
let rows = [];

function switchTab(id) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  const idx = ['tab-info','tab-add','tab-list','tab-export'].indexOf(id);
  document.querySelectorAll('.tab')[idx].classList.add('active');
  if (id === 'tab-list') renderList();
  if (id === 'tab-info' || id === 'tab-export') updateSummary();
}

function fmtNum(n) {
  const v = parseFloat(n) || 0;
  return v.toLocaleString('th-TH', {minimumFractionDigits: 0, maximumFractionDigits: 2});
}

function addRow() {
  const startDate = document.getElementById('f_startDate').value;
  if (!startDate) { alert('กรุณากรอกวันที่เริ่มต้น'); return; }
  const endDate = document.getElementById('f_endDate').value || startDate;
  const row = {
    startDate, startTime: document.getElementById('f_startTime').value,
    endDate, endTime: document.getElementById('f_endTime').value,
    status: document.getElementById('f_status').value,
    srId: document.getElementById('f_srId').value,
    soNo: document.getElementById('f_soNo').value,
    customer: document.getElementById('f_customer').value,
    site: document.getElementById('f_site').value,
    mileage: parseFloat(document.getElementById('f_mileage').value) || 0,
    transport: parseFloat(document.getElementById('f_transport').value) || 0,
    misc: parseFloat(document.getElementById('f_misc').value) || 0,
    hotel: parseFloat(document.getElementById('f_hotel').value) || 0,
    meal: parseFloat(document.getElementById('f_meal').value) || 0,
    other: parseFloat(document.getElementById('f_other').value) || 0,
    workHrs: parseFloat(document.getElementById('f_workHrs').value) || 0,
    travelHrs: parseFloat(document.getElementById('f_travelHrs').value) || 0,
    ot10: parseFloat(document.getElementById('f_ot10').value) || 0,
    ot15: parseFloat(document.getElementById('f_ot15').value) || 0,
    ot30: parseFloat(document.getElementById('f_ot30').value) || 0,
  };
  rows.push(row);
  updateBadge();
  updateSummary();
  clearForm();
  alert('✅ เพิ่มรายการแล้ว! (' + rows.length + ' รายการ)');
}

function clearForm() {
  const today = new Date().toISOString().split('T')[0];
  document.getElementById('f_startDate').value = today;
  document.getElementById('f_endDate').value = today;
  document.getElementById('f_startTime').value = '08:00';
  document.getElementById('f_endTime').value = '17:00';
  document.getElementById('f_status').value = '';
  ['f_srId','f_soNo','f_customer','f_site'].forEach(id => document.getElementById(id).value = '');
  ['f_mileage','f_transport','f_misc','f_hotel','f_meal','f_other','f_workHrs','f_travelHrs','f_ot10','f_ot15','f_ot30']
    .forEach(id => document.getElementById(id).value = '0');
}

function deleteRow(i) {
  if (confirm('ลบรายการนี้?')) {
    rows.splice(i, 1);
    updateBadge();
    updateSummary();
    renderList();
  }
}

function updateBadge() {
  const b = document.getElementById('badge');
  b.textContent = rows.length > 0 ? rows.length : '';
  b.style.cssText = rows.length > 0 ? 'background:var(--accent);color:white;border-radius:10px;padding:1px 6px;font-size:11px;' : '';
}

function updateSummary() {
  const km = rows.reduce((s,r) => s+r.mileage, 0);
  const kmCost = km * 5.7;
  const expense = rows.reduce((s,r) => s+r.transport+r.misc+r.hotel+r.meal+r.other, 0) + kmCost;
  const ot = rows.reduce((s,r) => s+r.ot10+r.ot15+r.ot30, 0);
  ['s_count','e_count'].forEach(id => { if(document.getElementById(id)) document.getElementById(id).textContent = rows.length; });
  ['s_total','e_total'].forEach(id => { if(document.getElementById(id)) document.getElementById(id).textContent = fmtNum(expense); });
  ['s_km','e_km'].forEach(id => { if(document.getElementById(id)) document.getElementById(id).textContent = fmtNum(km); });
  ['s_ot','e_ot'].forEach(id => { if(document.getElementById(id)) document.getElementById(id).textContent = fmtNum(ot); });
  if(document.getElementById('s_km_cost')) document.getElementById('s_km_cost').textContent = fmtNum(kmCost);
}

function renderList() {
  const c = document.getElementById('listContainer');
  if (rows.length === 0) {
    c.innerHTML = '<div class="empty-state"><div class="icon">📄</div><div>ยังไม่มีรายการ<br>ไปที่แท็บ "➕ เพิ่ม" เพื่อกรอกข้อมูล</div></div>';
    return;
  }
  c.innerHTML = rows.map((r,i) => {
    const sc = r.status==='complete'?'s-complete':r.status?'s-pending':'s-empty';
    const chips = [
      r.mileage>0 ? `<span class="num-chip">${r.mileage} km</span>` : '',
      r.transport>0 ? `<span class="num-chip">ขนส่ง ${fmtNum(r.transport)}฿</span>` : '',
      r.misc>0 ? `<span class="num-chip">เบ็ด ${fmtNum(r.misc)}฿</span>` : '',
      r.hotel>0 ? `<span class="num-chip">โรงแรม ${fmtNum(r.hotel)}฿</span>` : '',
      r.meal>0 ? `<span class="num-chip">อาหาร ${fmtNum(r.meal)}฿</span>` : '',
      r.other>0 ? `<span class="num-chip">อื่นๆ ${fmtNum(r.other)}฿</span>` : '',
      (r.ot10+r.ot15+r.ot30)>0 ? `<span class="num-chip">OT ${fmtNum(r.ot10+r.ot15+r.ot30)} ชม.</span>` : '',
    ].filter(Boolean).join('') || '<span class="num-chip zero">ไม่มีค่าใช้จ่าย</span>';
    return `<div class="row-card">
      <div class="row-card-header">
        <div>
          <div class="row-card-date">${r.startDate} ${r.startTime} → ${r.endDate} ${r.endTime}</div>
          <span class="status-badge ${sc}">${r.status||'—'}</span>
          ${r.soNo ? `<span style="font-size:11px;color:var(--muted);margin-left:6px;">${r.soNo}</span>` : ''}
        </div>
        <button class="delete-btn" onclick="deleteRow(${i})">ลบ</button>
      </div>
      <div class="row-card-body">
        <div class="customer">${r.customer||'—'}</div>
        <div class="site">${r.site||'—'}</div>
        <div class="row-nums">${chips}</div>
      </div>
    </div>`;
  }).join('');
}

function toExcelDate(dateStr, timeStr) {
  if (!dateStr) return '';
  const [y, m, d] = dateStr.split('-').map(Number);
  const [hh, mm] = (timeStr || '00:00').split(':').map(Number);

  // Pure calendar arithmetic — NO JavaScript Date object involved at all.
  // This avoids every possible timezone/DST/historical-offset quirk,
  // because Date objects (even with .UTC) can still get reinterpreted
  // by the runtime or by SheetJS using the system timezone.

  // Days since epoch 1970-01-01 using a proleptic Gregorian calendar calc
  function daysFromCivil(y, m, d) {
    y -= m <= 2 ? 1 : 0;
    const era = Math.floor((y >= 0 ? y : y - 399) / 400);
    const yoe = y - era * 400;
    const mp = (m + 9) % 12;
    const doy = Math.floor((153 * mp + 2) / 5) + d - 1;
    const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy;
    return era * 146097 + doe - 719468;
  }

  const daysSinceEpoch1970 = daysFromCivil(y, m, d);
  // Excel serial day 0 = 1899-12-30. 1970-01-01 = Excel serial 25569.
  const excelDay = daysSinceEpoch1970 + 25569;
  const dayFraction = (hh * 3600 + mm * 60) / 86400;
  return excelDay + dayFraction;
}

function exportToExcel() {
  const empName = document.getElementById('empName').value;
  const fseCode = document.getElementById('fseCode').value;
  const dateFrom = document.getElementById('dateFrom').value;
  const dateTo = document.getElementById('dateTo').value;

  const wb = XLSX.utils.book_new();
  const wsData = [];
  wsData.push(['Service Expenses Claim & OT']);
  wsData.push(['Name: '+empName,'','','','','','','','','','','','','','','','From: '+dateFrom]);
  wsData.push(['FSE Code: '+fseCode,'','','','','','','','','','','','','','','','To: '+dateTo]);
  wsData.push(['Start','Finish','Status','Service Request ID','Sale Order No.','Customer','Sitename','Mileage','Transport','Misc.','Hotel','Meal','Other','Real Work Time','','','OT','','','']);
  wsData.push(['','','','','','','','Total','(Baht)','(Baht)','(Baht)','(Baht)','(Baht)','Working','Traveling','Total','x 1.0','x 1.5','x 3.0','Total']);
  wsData.push(['','','','','','','','(km.)','','','','','','(Hrs.)','(Hrs.)','(Hrs.)','(Hrs.)','(Hrs.)','(Hrs.)','(Hrs.)']);

  rows.forEach(r => {
    wsData.push([
      toExcelDate(r.startDate,r.startTime), toExcelDate(r.endDate,r.endTime),
      r.status, r.srId, r.soNo, r.customer, r.site,
      r.mileage, r.transport, r.misc, r.hotel, r.meal, r.other,
      r.workHrs, r.travelHrs, r.workHrs+r.travelHrs,
      r.ot10, r.ot15, r.ot30, r.ot10+r.ot15+r.ot30
    ]);
  });
  // Ensure at least 26 rows (matches the original template's blank-row layout)
  // but allow unlimited rows beyond that if the user added more.
  const minRows = 26;
  for (let i = rows.length; i < minRows; i++) wsData.push(['','','','','','','',0,0,0,0,0,0,0,0,0,'','','',0]);

  // Data occupies rows 7..er (header rows 1-6, then data). Totals come right after.
  const er = 6 + Math.max(rows.length, minRows);
  wsData.push(['Total','','','','','','',`=SUM(H7:H${er})`,`=SUM(I7:I${er})`,`=SUM(J7:J${er})`,0,0,0,0,0,0,`=SUM(Q7:Q${er})`,`=SUM(R7:R${er})`,`=SUM(S7:S${er})`,`=SUM(T7:T${er})`]);
  wsData.push(['Mileage Cost (Baht)','','','','','','',`=H${er+1}*5.7`,'Total Expense','','','',`=SUM(I${er+1}:M${er+1})`,'Over Time Total','',`=P${er+1}`,'Over Time Total','','',`=T${er+1}`]);

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  const fmt = 'DD/MM/YYYY HH:MM';
  // Force these cells to be explicit numeric cells with a date display format.
  // We do NOT let SheetJS infer or convert anything — c.v is set to the raw
  // serial number we already computed via pure arithmetic (no Date object,
  // no timezone involved anywhere in this pipeline).
  for (let row=7; row<=er; row++) {
    ['A','B'].forEach(col => {
      const addr = col+row;
      const c = ws[addr];
      if (c && typeof c.v === 'number') {
        ws[addr] = { t: 'n', v: c.v, z: fmt };
      }
    });
  }
  ws['!cols'] = [{wch:18},{wch:18},{wch:10},{wch:18},{wch:14},{wch:30},{wch:35},{wch:10},{wch:10},{wch:10},{wch:10},{wch:10},{wch:10},{wch:10},{wch:10},{wch:10},{wch:8},{wch:8},{wch:8},{wch:10}];
  XLSX.utils.book_append_sheet(wb, ws, 'รายงานค่าใช้จ่ายการปฎิบัติงาน__');

  const fname = `รายงานค่าใช้จ่าย_${empName.replace(/\s/g,'_')}_${dateFrom}.xlsx`;

  // Mobile-friendly download: use base64 data URI
  const wbout = XLSX.write(wb, {bookType:'xlsx', type:'base64'});
  const dataURI = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + wbout;
  const a = document.createElement('a');
  a.href = dataURI;
  a.download = fname;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

// Init
const today = new Date().toISOString().split('T')[0];
document.getElementById('f_startDate').value = today;
document.getElementById('f_endDate').value = today;
updateSummary();
</script>
</body>
</html>
