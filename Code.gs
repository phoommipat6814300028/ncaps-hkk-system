/* ═══════════════════════════════════════════════════════════════
   🎥 ระบบจัดการกล้อง NCAPs — เขตรักษาพันธุ์สัตว์ป่าห้วยขาแข้ง
   Code.gs — ฉบับสมบูรณ์ (Backend + Router + เสิร์ฟหน้าเว็บ)
   ═══════════════════════════════════════════════════════════════ */

const SS_ID = '';                  // เว้นว่างไว้ถ้าสคริปต์ผูกกับ Sheet อยู่แล้ว
const TZ    = 'Asia/Bangkok';

const SH = { CAM:'CAMERAS', SIM:'SIMS', LOG:'LOG' };

const HEAD = {
  CAMERAS: ['camera_id','name','station','utm_zone','utm_e','utm_n','lat','lng',
            'signal','sim_number','install_date','removed_date','status',
            'remark','updated_by','updated_at'],
  SIMS:    ['sim_number','carrier','promo','promo_start','promo_end',
            'status','camera_id','note'],
  LOG:     ['timestamp','camera_id','action','note','by']
};

/* ═══════════════════════════════════════════════════════════════
   🌐 ROUTER — เสิร์ฟหน้าเว็บ + API ในตัวเดียว
   ═══════════════════════════════════════════════════════════════ */
function doGet(e) {
  const p = (e && e.parameter) || {};

  // โหมด API
  if (p.action) return json(routeGet(p.action));

  // โหมดหน้าเว็บ
  const page  = (p.page === 'dashboard') ? 'dashboard' : 'index';
  const title = (page === 'dashboard')
    ? 'Dashboard สถิติกล้อง NCAPs'
    : 'ระบบจัดการกล้อง NCAPs — ห้วยขาแข้ง';

  return HtmlService.createTemplateFromFile(page).evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try { return json(routePost(JSON.parse(e.postData.contents))); }
  catch (err) { return json({ ok:false, error:String(err) }); }
}

function routeGet(action) {
  try {
    if (action === 'list') return getAll();
    if (action === 'sims') return { ok:true, sims:readSheet(SH.SIM) };
    if (action === 'log')  return { ok:true, log:readSheet(SH.LOG).slice(-200).reverse() };
    if (action === 'ping') return { ok:true, msg:'API พร้อมใช้งาน', time:now() };
    return { ok:false, error:'ไม่รู้จัก action: ' + action };
  } catch (err) { return { ok:false, error:String(err) }; }
}

function routePost(b) {
  try {
    switch (b.action) {
      case 'saveCamera': return saveCamera(b.data, b.by);
      case 'uninstall':  return uninstall(b);
      case 'restore':    return restore(b);
      case 'deleteCam':  return deleteCam(b.camera_id, b.by);
      case 'saveSim':    return saveSim(b.data);
      case 'deleteSim':  return deleteSim(b.sim_number);
      default: return { ok:false, error:'ไม่รู้จัก action: ' + b.action };
    }
  } catch (err) { return { ok:false, error:String(err) }; }
}

/* สะพานสำหรับ google.script.run (เร็วกว่า fetch มาก) */
function apiGet(action)   { return JSON.stringify(routeGet(action)); }
function apiPost(payload) { return JSON.stringify(routePost(JSON.parse(payload))); }
function getUrl()         { return ScriptApp.getService().getUrl(); }

/* ═══════════════════════════════════════════════════════════════
   📚 ตัวช่วยจัดการชีต
   ═══════════════════════════════════════════════════════════════ */
function book() { return SS_ID ? SpreadsheetApp.openById(SS_ID)
                              : SpreadsheetApp.getActiveSpreadsheet(); }

function sheet(name) {

  const b = book();

  let s = b.getSheetByName(name);

  if (!s) {

    s = b.insertSheet(name);

    // ปรับตรงนี้: ถ้า HEAD[name] มีข้อมูล ให้ค่อยใส่

    if (HEAD[name] && HEAD[name].length > 0) {

      s.appendRow(HEAD[name]);

      s.getRange(1, 1, 1, HEAD[name].length).setFontWeight('bold').setBackground('#d1fae5');

      s.setFrozenRows(1);

    }

  }

  return s;

}

function readSheet(name) {
  const s = sheet(name), v = s.getDataRange().getValues();
  if (v.length < 2) return [];
  const h = v[0].map(x => String(x).trim());
  return v.slice(1)
    .filter(r => r.some(c => String(c).trim() !== ''))
    .map(r => {
      const o = {};
      h.forEach((k,i) => o[k] = norm(k, r[i]));
      return o;
    });
}

function norm(key, v) {
  if (v instanceof Date) {
    const dt = (key === 'timestamp' || key === 'updated_at');
    return Utilities.formatDate(v, TZ, dt ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd');
  }
  return v === null || v === undefined ? '' : v;
}

function findRow(name, key, val) {
  const s = sheet(name), v = s.getDataRange().getValues();
  const c = v[0].map(String).indexOf(key);
  if (c < 0) return -1;
  for (let i = 1; i < v.length; i++)
    if (String(v[i][c]).trim() === String(val).trim()) return i + 1;
  return -1;
}

const now   = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
const today = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
const json  = o => ContentService.createTextOutput(JSON.stringify(o))
                     .setMimeType(ContentService.MimeType.JSON);

function writeLog(camera_id, action, note, by) {
  sheet(SH.LOG).appendRow([now(), camera_id||'', action||'', note||'', by||'ระบบ']);
}

/* แจ้งเตือนแบบปลอดภัย — ใช้ได้ทุก context (แก้ปัญหา getUi() พัง) */
function notify(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* ไม่มี UI ก็ข้ามไป */ }
}

/* ═══════════════════════════════════════════════════════════════
   📖 อ่านข้อมูลทั้งหมด
   ═══════════════════════════════════════════════════════════════ */
function getAll() {
  const all  = readSheet(SH.CAM);
  const sims = readSheet(SH.SIM);
  const st = c => String(c.status || 'ONLINE').toUpperCase();
  return {
    ok: true,
    generated_at: now(),
    all: all,
    online:  all.filter(c => st(c) === 'ONLINE'),
    standby: all.filter(c => st(c) === 'STANDBY'),
    removed: all.filter(c => st(c) === 'REMOVED'),
    sims: sims
  };
}

/* ═══════════════════════════════════════════════════════════════
   💾 บันทึก / แก้ไข กล้อง
   ═══════════════════════════════════════════════════════════════ */
function saveCamera(d, by) {

  if (!d || !d.camera_id) return { ok:false, error:'กรุณาระบุรหัสกล้อง' };

  

  const s = sheet(SH.CAM);

  const r = findRow(SH.CAM, 'camera_id', d.camera_id);

  

  // แปลงพิกัด

  if (d.utm_e && d.utm_n) {

    const g = utmToLatLng(+d.utm_e, +d.utm_n, +(d.utm_zone || 47), true);

    d.lat = g.lat.toFixed(6);

    d.lng = g.lng.toFixed(6);

  }

  

  d.updated_by = by || 'ระบบ';

  d.updated_at = now();

  

  const row = HEAD.CAMERAS.map(k => (d[k] !== undefined && d[k] !== null) ? d[k] : '');

  

  if (r > 0) {

    // ถ้ารหัสซ้ำ ให้เขียนทับแถวเดิม (การวนกล้อง)

    s.getRange(r, 1, 1, HEAD.CAMERAS.length).setValues([row]);

    writeLog(d.camera_id, 'UPDATE', 'อัปเดตข้อมูลกล้อง', d.updated_by);

  } else {

    // รหัสใหม่ ให้เพิ่มแถว

    s.appendRow(row);

    writeLog(d.camera_id, 'CREATE', 'เพิ่มกล้องใหม่', d.updated_by);

  }

  

  if (d.sim_number) linkSim(d.sim_number, d.camera_id);

  return { ok:true, msg:'บันทึกสำเร็จ', camera_id:d.camera_id };

}
/* ⛔ ถอนกล้อง */
function uninstall(b) {
  const r = findRow(SH.CAM, 'camera_id', b.camera_id);
  if (r < 0) return { ok:false, error:'ไม่พบกล้อง ' + b.camera_id };
  const s = sheet(SH.CAM);
  const col = k => HEAD.CAMERAS.indexOf(k) + 1;
  const remark = (b.reason || 'ไม่ระบุเหตุผล') + (b.note ? ' — ' + b.note : '');

  s.getRange(r, col('status')).setValue('REMOVED');
  s.getRange(r, col('removed_date')).setValue(b.removed_date || today());
  s.getRange(r, col('remark')).setValue(remark);
  s.getRange(r, col('updated_by')).setValue(b.by || 'ระบบ');
  s.getRange(r, col('updated_at')).setValue(now());

  const sim = s.getRange(r, col('sim_number')).getValue();
  if (sim) unlinkSim(sim);

  writeLog(b.camera_id, 'UNINSTALL', remark, b.by);
  return { ok:true, msg:'ถอนกล้องเรียบร้อย' };
}

/* ↩ คืนกล้องเข้าคลัง */
function restore(b) {
  const r = findRow(SH.CAM, 'camera_id', b.camera_id);
  if (r < 0) return { ok:false, error:'ไม่พบกล้อง ' + b.camera_id };
  const s = sheet(SH.CAM);
  const col = k => HEAD.CAMERAS.indexOf(k) + 1;

  s.getRange(r, col('status')).setValue(b.to || 'STANDBY');
  s.getRange(r, col('removed_date')).setValue('');
  s.getRange(r, col('updated_by')).setValue(b.by || 'ระบบ');
  s.getRange(r, col('updated_at')).setValue(now());

  writeLog(b.camera_id, 'RESTORE', 'นำกลับเข้าคลัง', b.by);
  return { ok:true, msg:'นำกลับเข้าคลังเรียบร้อย' };
}

/* 🗑 ลบกล้อง */
function deleteCam(camera_id, by) {
  const r = findRow(SH.CAM, 'camera_id', camera_id);
  if (r < 0) return { ok:false, error:'ไม่พบกล้อง ' + camera_id };
  sheet(SH.CAM).deleteRow(r);
  writeLog(camera_id, 'DELETE', 'ลบออกจากระบบ', by);
  return { ok:true, msg:'ลบเรียบร้อย' };
}

/* ═══════════════════════════════════════════════════════════════
   📱 จัดการ SIM
   ═══════════════════════════════════════════════════════════════ */
function saveSim(d) {
  if (!d || !d.sim_number) return { ok:false, error:'กรุณาระบุเบอร์ซิม' };
  const s = sheet(SH.SIM);
  const r = findRow(SH.SIM, 'sim_number', d.sim_number);
  const row = HEAD.SIMS.map(k => d[k] !== undefined ? d[k] : '');
  if (r > 0) s.getRange(r, 1, 1, HEAD.SIMS.length).setValues([row]);
  else       s.appendRow(row);
  return { ok:true, msg:'บันทึก SIM เรียบร้อย' };
}

function deleteSim(sim_number) {
  const r = findRow(SH.SIM, 'sim_number', sim_number);
  if (r < 0) return { ok:false, error:'ไม่พบ SIM ' + sim_number };
  sheet(SH.SIM).deleteRow(r);
  return { ok:true, msg:'ลบ SIM เรียบร้อย' };
}

function linkSim(sim_number, camera_id) {
  const r = findRow(SH.SIM, 'sim_number', sim_number);
  if (r < 0) return;
  const s = sheet(SH.SIM);
  s.getRange(r, HEAD.SIMS.indexOf('status') + 1).setValue('IN_USE');
  s.getRange(r, HEAD.SIMS.indexOf('camera_id') + 1).setValue(camera_id);
}

function unlinkSim(sim_number) {
  const r = findRow(SH.SIM, 'sim_number', sim_number);
  if (r < 0) return;
  const s = sheet(SH.SIM);
  s.getRange(r, HEAD.SIMS.indexOf('status') + 1).setValue('SPARE');
  s.getRange(r, HEAD.SIMS.indexOf('camera_id') + 1).setValue('');
}

/* ═══════════════════════════════════════════════════════════════
   🗺️ แปลงพิกัด UTM → Lat/Lng (WGS84)
   ═══════════════════════════════════════════════════════════════ */
function utmToLatLng(easting, northing, zone, northern) {
  const a = 6378137.0, f = 1/298.257223563;
  const e2 = f*(2-f), e1sq = e2/(1-e2), k0 = 0.9996;
  const x = easting - 500000.0;
  const y = northern ? northing : northing - 10000000.0;

  const M  = y / k0;
  const mu = M / (a*(1 - e2/4 - 3*e2*e2/64 - 5*Math.pow(e2,3)/256));
  const e1 = (1 - Math.sqrt(1-e2)) / (1 + Math.sqrt(1-e2));

  const fp = mu
    + (3*e1/2 - 27*Math.pow(e1,3)/32)      * Math.sin(2*mu)
    + (21*e1*e1/16 - 55*Math.pow(e1,4)/32) * Math.sin(4*mu)
    + (151*Math.pow(e1,3)/96)              * Math.sin(6*mu)
    + (1097*Math.pow(e1,4)/512)            * Math.sin(8*mu);

  const C1 = e1sq*Math.pow(Math.cos(fp),2);
  const T1 = Math.pow(Math.tan(fp),2);
  const R1 = a*(1-e2)/Math.pow(1 - e2*Math.pow(Math.sin(fp),2), 1.5);
  const N1 = a/Math.sqrt(1 - e2*Math.pow(Math.sin(fp),2));
  const D  = x/(N1*k0);

  const lat = fp - (N1*Math.tan(fp)/R1) * (
      D*D/2
    - (5 + 3*T1 + 10*C1 - 4*C1*C1 - 9*e1sq)*Math.pow(D,4)/24
    + (61 + 90*T1 + 298*C1 + 45*T1*T1 - 3*C1*C1 - 252*e1sq)*Math.pow(D,6)/720);

  const lng = (
      D
    - (1 + 2*T1 + C1)*Math.pow(D,3)/6
    + (5 - 2*C1 + 28*T1 - 3*C1*C1 + 8*e1sq + 24*T1*T1)*Math.pow(D,5)/120
  ) / Math.cos(fp);

  const lon0 = (zone*6 - 183) * Math.PI/180;
  return { lat: lat*180/Math.PI, lng: (lon0 + lng)*180/Math.PI };
}

/* คำนวณ lat/lng ใหม่ทั้งชีต */
function recalcLatLng() {
  const s = sheet(SH.CAM), v = s.getDataRange().getValues();
  if (v.length < 2) return;
  const h = v[0].map(String);
  const iE = h.indexOf('utm_e'), iN = h.indexOf('utm_n'), iZ = h.indexOf('utm_zone');
  const iLa = h.indexOf('lat'),  iLo = h.indexOf('lng');
  let n = 0;
  for (let i = 1; i < v.length; i++) {
    const E = +v[i][iE], N = +v[i][iN];
    if (!E || !N) continue;
    const g = utmToLatLng(E, N, +(v[i][iZ] || 47), true);
    s.getRange(i+1, iLa+1).setValue(g.lat.toFixed(6));
    s.getRange(i+1, iLo+1).setValue(g.lng.toFixed(6));
    n++;
  }
  notify('✅ คำนวณพิกัดใหม่แล้ว ' + n + ' รายการ');
}

/* ═══════════════════════════════════════════════════════════════
   🛠️ ติดตั้งครั้งแรก
   ═══════════════════════════════════════════════════════════════ */
function setupSheets() {
  const b = book();
  ['CAMERAS','SIMS','LOG'].forEach(n => sheet(n));

  const cam = sheet(SH.CAM);
  if (cam.getLastRow() < 2) {
    cam.appendRow(['HKK-001','กล้องหน่วยไซเบอร์','หน่วยพิทักษ์ป่าไซเบอร์',47,528400,1690200,
      '','','P5','0812345678','2026-01-15','','ONLINE','','ระบบ',now()]);
    cam.appendRow(['HKK-002','กล้องเขาบันได','หน่วยพิทักษ์ป่าเขาบันได',47,531800,1702500,
      '','','P3','','2026-02-01','','ONLINE','','ระบบ',now()]);
    cam.appendRow(['HKK-003','กล้องสำรอง 01','คลังพัสดุ',47,'','','','','','', '','','STANDBY','','ระบบ',now()]);
  }

  const sim = sheet(SH.SIM);
  if (sim.getLastRow() < 2) {
    sim.appendRow(['0812345678','AIS','เน็ต 30GB/เดือน','2026-01-15','2027-01-14','IN_USE','HKK-001','']);
    sim.appendRow(['0899999999','TRUE','เน็ต 20GB/เดือน','2026-06-01','2026-12-01','SPARE','','สำรอง']);
  }

  recalcLatLng();
  notify('✅ สร้างตารางเรียบร้อยแล้ว');
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('🎥 NCAPs')
      .addItem('🛠️ ติดตั้งตารางครั้งแรก', 'setupSheets')
      .addItem('🩹 ซ่อมชีตให้ตรงกับโค้ด', 'repairSheets')
      .addItem('🗺️ คำนวณพิกัดใหม่', 'recalcLatLng')
      .addToUi();
  } catch (e) {}
}
/* ═══════════════════════════════════════════════════════════════
   🩹 ซ่อมชีตให้ตรงกับโค้ด — รันครั้งเดียวพอ
   ═══════════════════════════════════════════════════════════════ */

// ชื่อคอลัมน์เก่า → ชื่อใหม่
const ALIAS = {
  camera_name: 'name',
  cam_name:    'name',
  location:    'station',
  unit:        'station',
  point:       'station',
  by:          'updated_by',
  user:        'updated_by'
};

function repairSheets() {
  const b = book();
  let report = [];

  /* ---------- ซ่อม CAMERAS ---------- */
  const s = b.getSheetByName(SH.CAM);
  if (s) {
    const v = s.getDataRange().getValues();
    const h = v[0].map(x => String(x).trim());

    const good = [], bad = [];
    for (let i = 1; i < v.length; i++) {
      if (v[i].every(c => String(c).trim() === '')) continue;

      const o = {};
      h.forEach((k, j) => {
        const key = ALIAS[k] || k;
        o[key] = v[i][j] instanceof Date
          ? Utilities.formatDate(v[i][j], TZ, 'yyyy-MM-dd')
          : v[i][j];
      });

      // แถวที่ status ไม่ถูกต้อง = แถวที่ข้อมูลเลื่อน → แยกไปสำรอง
      const st = String(o.status || '').trim().toUpperCase();
      if (['ONLINE','STANDBY','REMOVED'].indexOf(st) < 0) { bad.push(v[i]); continue; }

      o.status = st;
      o.utm_zone = String(o.utm_zone || 47).replace(/[^0-9]/g, '') || 47;  // "47N" → 47
      if (!o.updated_by) o.updated_by = 'ระบบ';
      if (!o.station)    o.station    = '';
      good.push(HEAD.CAMERAS.map(k => o[k] !== undefined ? o[k] : ''));
    }

    // เก็บแถวเสียไว้ก่อน กันข้อมูลหาย
    if (bad.length) {
      const bk = b.getSheetByName('CAMERAS_เสีย') || b.insertSheet('CAMERAS_เสีย');
      bk.clear();
      bk.appendRow(h);
      bk.getRange(2, 1, bad.length, bad[0].length).setValues(bad);
    }

    // เขียนใหม่ทั้งชีตด้วยหัวตารางที่ถูกต้อง
    s.clear();
    s.appendRow(HEAD.CAMERAS);
    s.getRange(1,1,1,HEAD.CAMERAS.length).setFontWeight('bold').setBackground('#065f46').setFontColor('#fff');
    s.setFrozenRows(1);
    if (good.length) s.getRange(2, 1, good.length, HEAD.CAMERAS.length).setValues(good);

    report.push('CAMERAS: กู้ได้ ' + good.length + ' แถว · แยกแถวเสีย ' + bad.length + ' แถว');
  }

  /* ---------- ซ่อม SIMS ---------- */
  const si = b.getSheetByName(SH.SIM);
  if (si) {
    const v = si.getDataRange().getValues();
    const h = v[0].map(x => String(x).trim());
    const rows = [];
    for (let i = 1; i < v.length; i++) {
      if (!String(v[i][0]).trim()) continue;
      const o = {};
      h.forEach((k, j) => o[ALIAS[k] || k] = v[i][j] instanceof Date
        ? Utilities.formatDate(v[i][j], TZ, 'yyyy-MM-dd') : v[i][j]);
      rows.push(HEAD.SIMS.map(k => o[k] !== undefined ? o[k] : ''));
    }
    si.clear();
    si.appendRow(HEAD.SIMS);
    si.getRange(1,1,1,HEAD.SIMS.length).setFontWeight('bold').setBackground('#065f46').setFontColor('#fff');
    si.setFrozenRows(1);
    if (rows.length) si.getRange(2, 1, rows.length, HEAD.SIMS.length).setValues(rows);
    report.push('SIMS: ' + rows.length + ' แถว');
  }

  /* ---------- ตรวจ LOG ---------- */
  sheet(SH.LOG);

  recalcLatLng();
  notify('🩹 ซ่อมเรียบร้อย\n\n' + report.join('\n'));
}
