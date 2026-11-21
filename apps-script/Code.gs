/**
 * HIM Stats — ระบบหลังบ้าน (Google Apps Script)
 * วางไฟล์นี้ใน Extensions > Apps Script ของ Google Sheet กลาง
 * แล้วตั้ง Script Properties (Project Settings > Script properties):
 *   LINE_ACCESS_TOKEN  = Channel access token ของ OA (สำหรับตอบแชท)
 *   CURRENT_ROUND      = รอบที่เปิดให้ส่ง เช่น 2026-Q4
 */
const SH = { REG: 'ทะเบียนคริสตจักร', STAT: 'สถิติ', USER: 'ผู้ใช้', LOG: 'บันทึกการแก้ไข' };
const P = PropertiesService.getScriptProperties();
const XF = [['mem','จน.สมาชิก'],['lgatt','มา LG'],['int','ผู้สนใจ'],['nb','ผู้เชื่อใหม่'],['pwo','เข้าเรียน พวอ.'],['lgl','ผู้นำ LG'],['men','พี่เลี้ยง'],['o2o','1 ต่อ 1'],
  ['a1','เด็ก 3-12'],['a2','Ablaze 13-18'],['a3','Uni 19-22'],['a4','Young Pro 23-35'],['a5','Pro 36-59'],['a6','ครอบครัว'],['a7','สูงวัย 60+'],
  ['c1','✓ LG ทุกสัปดาห์'],['c2','✓ ชั้นดาวิดทุกเดือน'],['c3','✓ ชั้นเรียน พวอ.'],['c4','✓ ทีม WAM'],['cW','✓ Worship'],['cA','✓ Art'],['cM','✓ Media'],['c5','✓ เทศนา'],['c6','✓ ถวายทรัพย์'],['c7','✓ ทีมรับใช้เพียงพอ'],['c8','✓ ศบ ประจำ'],['tgt','เป้า Q4 ยอดนมัสการ'],['tlg','เป้า Q4 LG'],['tnb','เป้า Q4 ผู้เชื่อใหม่'],['tmem','เป้า สมาชิก'],['tlga','เป้า มา LG'],['tint','เป้า ผู้สนใจ'],['tpwo','เป้า พวอ.'],['tlgl','เป้า ผู้นำ LG'],['tmen','เป้า พี่เลี้ยง'],['to2o','เป้า 1 ต่อ 1']];
function regTargets_() {   // เป้าหมายรายภาค 2026 จากชีต v2 แท็บ 6
  try { const c = cache(), k = c.get('rt'); if (k) return JSON.parse(k);
    const v = SpreadsheetApp.openById(P.getProperty('SOURCE_SHEET_ID')).getSheets().find(s => /^6 /.test(s.getName())).getDataRange().getValues();
    const o = {}; v.slice(4).forEach(r => { if (r[0] && r[0] !== 'รวม') o[String(r[0]).replace(/\s/g, '')] = { t111: +r[9] || 0, tlg: +r[10] || 0, tpio: +r[11] || 0, tcen: +r[12] || 0, tch: +r[13] || 0 }; });
    c.put('rt', JSON.stringify(o), 600); return o; } catch (e) { return {}; }
}

const ss = () => SpreadsheetApp.openById('1WRbfK1_dg_rUDTlRiZdYIAFxhi5k2fnPvT9sBr5nppc'); // เปิดด้วย ID เสมอ (ใช้ได้ทั้งจากทริกเกอร์ของชีต v2)

/** รันครั้งเดียวหลังวางโค้ด: สร้างแท็บผู้ใช้/บันทึก และหัวคอลัมน์สถิติ */
const UHEAD = ['username','ชื่อ','บทบาท(leader/admin)','รหัสคริสตจักร(คั่นด้วย ,)','สถานะ(active/disabled)','สร้างเมื่อ','สร้างโดย','passwordHash','salt','lineUserId'];
function setup() {
  const s = ss();
  let u = s.getSheetByName(SH.USER);
  if (!u) u = s.insertSheet(SH.USER);
  if (u.getRange(1, 1).getValue() !== 'username') { u.clear(); u.appendRow(UHEAD); }
  if (!s.getSheetByName(SH.LOG)) s.insertSheet(SH.LOG).appendRow(['เวลา','ผู้ใช้','การกระทำ','รายละเอียด']);
  s.getSheetByName(SH.STAT).getRange(1, 1, 1, 11).setValues([['รหัส','ชื่อ','ภาค','รอบ','จำนวนคน','ผู้มาใหม่','บัพติศมา','Life Group','อธิษฐาน/ข่าว','ส่งโดย','เวลาส่ง']]);
  s.getSheetByName(SH.REG).getRange(1, 16, 1, 2).setValues([['เข้าร่วมโครงการ 111', 'เข้าร่วมโครงการพลัส']]);
  if (!P.getProperty('CURRENT_ROUND')) P.setProperty('CURRENT_ROUND', '2026-Q4');
}

/* ---------- ข้อมูล ---------- */
function rows_(name) { const v = ss().getSheetByName(name).getDataRange().getValues(); v.shift(); return v; }
function churches_() {
  return rows_(SH.REG).filter(r => r[0] && r[11] !== 'ปิด').map(r => ({ code: r[0], reg: r[1], prov: r[2], type: r[3], name: r[4], leader: r[5], shirt: r[6], phone: r[7], email: r[8], reged: r[9], regdate: r[10], status: r[11], note: r[14], p111: !!r[15], plus: !!r[16], addr: r[17] || '', map: r[18] || '', photo: r[19] || '', pfull: r[20] || '', freq: r[21] || '', pstart: r[22] instanceof Date ? Utilities.formatDate(r[22], 'Asia/Bangkok', 'yyyy-MM-dd') : (r[22] || ''), mother: r[23] || '', meet: r[24] || '', stage: r[25] || '', need: r[26] || '', pconf: r[27] || '', en: r[28] || '', enl: r[29] || '' }));
}
function stats_() {
  return rows_(SH.STAT).filter(r => r[0]).map(r => { const x = {}; XF.forEach(([k], j) => { if (r[11 + j] !== '' && r[11 + j] != null) x[k] = r[11 + j]; });
    return { code: r[0], round: String(r[3]).slice(0, 7), n: Number(r[4]), nw: r[5], bp: r[6], lg: r[7], note: r[8], by: r[9], at: r[10], x }; });
}
function users_() {
  return rows_(SH.USER).map((r, i) => ({ row: i + 2, username: String(r[0]).trim(), name: r[1], role: r[2], codes: String(r[3]).split(',').map(x => x.trim()).filter(String), status: r[4], hash: r[7], salt: r[8], lineUserId: r[9] })).filter(u => u.username);
}
function log_(who, act, d) { ss().getSheetByName(SH.LOG).appendRow([new Date(), who, act, JSON.stringify(d)]); }
function scope_(u, C) {   // ขยายสิทธิ์: coord = ทั้งภาค (REG:ชื่อภาค), exec = ทั้งประเทศ
  const regs = u.codes.filter(k => k.startsWith('REG:')).map(k => k.slice(4));
  if (u.role === 'exec') u.codes = C.map(c => c.code);
  else if (u.role === 'coord') u.codes = [...new Set(u.codes.filter(k => !k.startsWith('REG:')).concat(C.filter(c => regs.includes(String(c.reg))).map(c => c.code)))];
  u.regs = regs; return u;
}
function pub_(u) { return { username: u.username, name: u.name, role: u.role, codes: u.codes, status: u.status, line: !!u.lineUserId }; }

/* ---------- รหัสผ่านและ session ---------- */
function hash_(pw, salt) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8)); }
function rand_(n) { const c = 'abcdefghjkmnpqrstuvwxyz23456789'; let s = ''; for (let i = 0; i < n; i++) s += c[Math.floor(Math.random() * c.length)]; return s; }
function setPw_(u, pw) { const salt = Utilities.getUuid(); const sh = ss().getSheetByName(SH.USER); sh.getRange(u.row, 8, 1, 2).setValues([[hash_(pw, salt), salt]]); }
const cache = () => CacheService.getScriptCache();
function session_(token) {
  const name = token && cache().get('s:' + token);
  if (!name) throw new Error('SESSION');
  cache().put('s:' + token, name, 21600);
  const u = users_().find(x => x.username === name);
  if (!u || u.status !== 'active') throw new Error('SESSION');
  return u;
}

/* ---------- API สำหรับเว็บแอป ---------- */
function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  if (body.events) return webhook_(body);           // ข้อความจาก LINE OA
  let out;
  try { out = { ok: true, data: api_(body) }; } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return ContentService.createTextOutput('HIM Stats API ok'); }

function api_(b) {
  const lock = LockService.getScriptLock();
  if (b.action === 'names') return users_().filter(u => u.status === 'active' && u.role !== 'admin').map(u => [u.name, u.username]);
  if (b.action === 'counsel') {
    const c = cache(), day = 'cs:' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd'), n = Number(c.get(day) || 0);
    if (n >= 50) throw new Error('วันนี้มีคำขอจำนวนมาก กรุณาลองใหม่พรุ่งนี้ หรือโทร 1323');
    const cl = (x, k) => String(x == null ? '' : x).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, k);
    const f = { name: b.anon ? 'ไม่ประสงค์ออกนาม' : cl(b.name, 80), church: cl(b.church, 80), topic: cl(b.topic, 1500), help: cl(b.help, 200), contact: cl(b.contact, 120), urg: cl(b.urg, 30) };
    if (!f.topic) throw new Error('กรุณาพิมพ์ประเด็นที่อยากปรึกษา');
    if (!b.anon && !f.name) throw new Error('กรุณาใส่ชื่อ หรือเลือกไม่ประสงค์ออกนาม');
    let sh = ss().getSheetByName('คำขอปรึกษา'); if (!sh) { sh = ss().insertSheet('คำขอปรึกษา'); sh.appendRow(['เวลา', 'ชื่อ', 'คริสตจักร', 'ประเด็นที่อยากปรึกษา', 'ความช่วยเหลือที่อยากได้', 'ช่องทางติดต่อ', 'ความเร่งด่วน', 'สถานะ', 'ผู้ดูแล']); }
    sh.appendRow([new Date(), f.name, f.church, f.topic, f.help, f.contact, f.urg, 'รับเรื่อง', '']); c.put(day, String(n + 1), 86400);
    const g = P.getProperty('NOTIFY_GROUP');
    if (g) push_(g, `🔔 ศูนย์ให้คำปรึกษา · คำขอใหม่${f.urg === 'ด่วน' ? ' 🔴 ด่วน' : ''}\n👤 ${f.name}${f.church ? ' · ' + f.church : ''}\n💬 ประเด็น: ${f.topic}\n🤝 อยากให้ช่วย: ${f.help || '-'}\n📞 ติดต่อกลับ: ${f.contact || '-'}\n⏱ ความเร่งด่วน: ${f.urg || '-'}`);
    return { ok: 1 };
  }
  if (b.action === 'status') return { needsAdmin: !users_().some(u => u.role === 'admin' && u.hash) };
  if (b.action === 'bootstrap') {                     // สร้างแอดมินคนแรก (ใช้ได้ครั้งเดียว)
    if (users_().some(u => u.role === 'admin' && u.hash)) throw new Error('มีแอดมินแล้ว');
    if (!b.username || String(b.password || '').length < 8) throw new Error('รหัสผ่านอย่างน้อย 8 ตัว');
    ss().getSheetByName(SH.USER).appendRow([String(b.username).trim(), b.name || b.username, 'admin', '', 'active', new Date(), 'bootstrap', '', '', '']);
    setPw_(users_().find(u => u.username === String(b.username).trim()), b.password);
    log_(b.username, 'bootstrap', {});
    b.action = 'login';
  }
  if (b.action === 'magic') {
    const k = P.getProperty('EXEC_LINK_KEY'), mk = String(b.k || ''), lm = JSON.parse(P.getProperty('LINK_KEYS') || '{}');
    const un = (k && mk === k) ? 'him_exec' : lm[mk]; if (!mk || !un) throw new Error('ลิงก์ไม่ถูกต้อง');
    const t = users_().find(x => x.username === un && x.status === 'active'); if (!t) throw new Error('ไม่พบบัญชี');
    const token = Utilities.getUuid() + rand_(8); cache().put('s:' + token, t.username, 21600); log_(t.username, 'magicLogin', {}); return { token };
  }
  if (b.action === 'login') {
    const key = 'f:' + String(b.username).toLowerCase(), fails = Number(cache().get(key) || 0);
    if (fails >= 5) throw new Error('ใส่รหัสผิดหลายครั้ง กรุณารอ 10 นาที');
    const q = String(b.username || '').trim().toLowerCase(), all = users_();
    let u = all.find(x => x.username.toLowerCase() === q);
    if (!u) { const byName = all.filter(x => String(x.name).trim().toLowerCase() === q); if (byName.length > 1) throw new Error('ชื่อนี้ซ้ำกันหลายแห่ง กรุณาใช้รหัส เช่น n1-001'); u = byName[0]; }
    if (!u || u.status !== 'active' || !u.hash || hash_(String(b.password || ''), u.salt) !== u.hash) {
      cache().put(key, String(fails + 1), 600); throw new Error('username หรือรหัสผ่านไม่ถูกต้อง');
    }
    const token = Utilities.getUuid() + rand_(8);
    cache().put('s:' + token, u.username, 21600); log_(u.username, 'login', {});
    return { token };
  }
  const u = session_(b.token), round = P.getProperty('CURRENT_ROUND'), C = churches_();
  const admin = u.role === 'admin'; scope_(u, C);
  if (u.role === 'exec' && ['submit', 'updateInfo', 'regReport'].includes(b.action)) throw new Error('บัญชีผู้บริหารดูได้อย่างเดียว');
  if (b.action === 'ask') {
    if (!['admin', 'exec'].includes(u.role) || !P.getProperty('ANTHROPIC_API_KEY')) throw new Error('ใช้ได้เฉพาะผู้บริหาร/แอดมิน');
    return { text: askAI_(String(b.q || '').slice(0, 500), u) };
  }
  if (b.action === 'me') {
    const S = stats_().filter(s => admin || u.codes.includes(s.code));
    return { targets: regTargets_(), user: pub_(u), round, churches: admin ? C : C.filter(c => u.codes.includes(c.code)), stats: S };
  }
  if (b.action === 'logout') { cache().remove('s:' + b.token); return {}; }
  if (b.action === 'password') {
    if (hash_(String(b.old || ''), u.salt) !== u.hash) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
    if (String(b.password || '').length < 8) throw new Error('รหัสผ่านใหม่อย่างน้อย 8 ตัว');
    setPw_(u, b.password); log_(u.username, 'password', {}); return {};
  }
  if (b.action === 'regReports' || b.action === 'regReport') {   // รายงานระดับภาค (R2,R4,R5,R6,R8)
    const RT = { r2: ['รายงานภาค R2 สถาบัน Youth', ['จังหวัด','มหาวิทยาลัย/วิทยาลัย','แผนการประกาศ','แผนการสร้างสาวก']],
      r4: ['รายงานภาค R4 ผลดำเนินงาน', ['เป้าหมาย','ผลลัพธ์','กิจกรรม/การประกาศ','ลิงก์ภาพ']],
      r5: ['รายงานภาค R5 BestPractice-PainPoint', ['ประเภท','ลำดับ','รายละเอียด']],
      r6: ['รายงานภาค R6 แผน 111', ['สิ่งที่ตั้งใจเปิด','จำนวน (แห่ง)','ชื่อผู้นำ','กำหนดเปิด','สิ่งที่ต้องการสนับสนุน','ประเมินผล']],
      r8: ['รายงานภาค R8 ประชุมดาวิด', ['วันที่ประชุม','ย้ำนิมิต','ทบทวนเป้าหมาย','สถิติการเติบโต','โครงการ 111/ประกาศ','สอนพระคัมภีร์','รายงานถวายทรัพย์','สะท้อนปัญหา/แก้ไข','อธิษฐานเผื่อกัน','อื่นๆ','รายละเอียด']] };
    const myRegs = admin ? [...new Set(C.map(c => c.reg))] : [...new Set(C.filter(c => u.codes.includes(c.code)).map(c => c.reg))];
    const tab = t => { const [n, h] = RT[t]; let sh = ss().getSheetByName(n); if (!sh) { sh = ss().insertSheet(n); sh.appendRow(['รอบ','ภาค'].concat(h).concat(['ส่งโดย','เวลาส่ง'])); sh.setFrozenRows(1); } return sh; };
    if (b.action === 'regReports') {
      const out = {}; Object.keys(RT).forEach(t => { out[t] = tab(t).getDataRange().getValues().slice(1).filter(r => String(r[0]) === round && myRegs.includes(r[1])).map(r => r.slice(1, 2 + RT[t][1].length)); });
      return { regs: myRegs, round, data: out };
    }
    if (!RT[b.type] || !myRegs.includes(b.reg)) throw new Error('ไม่มีสิทธิ์ส่งรายงานของภาคนี้');
    lock.waitLock(20000);
    try {
      const sh = tab(b.type), v = sh.getDataRange().getValues(), k = RT[b.type][1].length;
      for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === round && v[i][1] === b.reg) sh.deleteRow(i + 1);
      const rows = (b.rows || []).filter(r => r.some(x => x !== '' && x !== false && x != null)).map(r => [round, b.reg].concat(r.slice(0, k).map(x => x === true ? '✓' : x === false ? '' : x)).concat([u.name || u.username, new Date()]));
      if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    } finally { lock.releaseLock(); }
    log_(u.username, 'regReport', { type: b.type, reg: b.reg, n: (b.rows || []).length });
    gpush_(`📝 รายงานภาค ${b.reg} · ${({ r6: 'แผน 111', r5: 'Best Practice & Pain Point', r4: 'ผลการดำเนินงาน Q3', r2: 'แผนสถาบัน Youth', r8: 'ประชุมดาวิดภาค' })[b.type] || b.type}\n✍️ โดย ${u.name || u.username}`);
    return { saved: true };
  }
  if (['members', 'memberSave', 'memberDel', 'mvConn'].includes(b.action)) {   // ทะเบียนสมาชิกรายคริสตจักร
    const MH = ['id','รหัสคริสตจักร','คำนำหน้า','ชื่อ','สกุล','ชื่อเล่น','วันเกิด','เพศ','ที่อยู่','จังหวัด','อำเภอ','ตำบล','รหัสไปรษณีย์','ทำงานที่','ที่อยู่ที่ทำงาน','email','โทรศัพท์','มือถือ','ผู้ติดต่อฉุกเฉิน','เบอร์ฉุกเฉิน','พี่เลี้ยง','สถานะ','หมายเหตุ','รูป (Drive id)','สถานะฝ่ายวิญญาณ','รับเชื่อ','นิวไลฟ์','บัพติศมา','สมาชิกภาพ','โน้ตฝ่ายวิญญาณ','คู่สมรส','บิดา','มารดา','แก้ไขโดย','เวลาแก้ไข'];
    const MK = ['id','code','pre','fn','ln','nick','dob','sex','addr','prov','dist','sub','zip','work','waddr','email','tel','mob','emer','emtel','mentor','status','note','photo','sp','bel','nl','bap','memb','spnote','spouse','father','mother','mvloc','mvwhat','mvstat','mvserve','mvdate','mvreg','mvconn'];
    let sh = ss().getSheetByName('สมาชิก'); if (!sh) { sh = ss().insertSheet('สมาชิก'); sh.setFrozenRows(1); }
    if (sh.getRange(1, 1, 1, MH.length).getValues()[0].join('|') !== MH.join('|')) sh.getRange(1, 1, 1, MH.length).setValues([MH]);
    const v = sh.getDataRange().getValues();
    const RG = MK.indexOf('mvreg'), inReg = r => (u.regs || []).includes(String(r[RG] || ''));
    if (b.action === 'mvConn') { const i = v.findIndex(r => r[0] === b.id); if (i < 1 || !(u.role === 'admin' || (u.role === 'coord' && (inReg(v[i]) || u.codes.includes(v[i][1]))))) throw new Error('ไม่มีสิทธิ์'); const st = ['ยังไม่ได้ติดต่อ', 'ติดต่อแล้ว', 'เข้าร่วมแล้ว'].includes(b.conn) ? b.conn : 'ยังไม่ได้ติดต่อ'; sh.getRange(i + 1, MK.indexOf('mvconn') + 1).setValue(st); log_(u.username, 'mvConn', { id: b.id, st }); gpush_(`🚚 อัปเดตการเชื่อมต่อสมาชิกย้ายเข้า: ${v[i][MK.indexOf('nick')] || v[i][MK.indexOf('fn')]} → ${st}\n📍 ภาค${v[i][RG]}\n✍️ โดย ${u.name || u.username}`); return { ok: 1 }; }
    if (b.action === 'members') return v.slice(1).filter(r => r[0] && (u.codes.includes(r[1]) || inReg(r))).map(r => { const o = {}; MK.forEach((k, i) => o[k] = r[i] instanceof Date ? Utilities.formatDate(r[i], 'Asia/Bangkok', 'yyyy-MM-dd') : r[i]); return o; });
    if (u.role === 'exec') throw new Error('บัญชีผู้บริหารดูได้อย่างเดียว');
    let prevReg = ''; const m = b.m || {}, isNew = !m.id; if (!u.codes.includes(m.code)) throw new Error('ไม่มีสิทธิ์จัดการสมาชิกของคริสตจักรนี้');
    lock.waitLock(20000);
    try {
      const i = m.id ? v.findIndex(r => r[0] === m.id) : -1; if (i > 0 && !u.codes.includes(v[i][1])) throw new Error('ไม่มีสิทธิ์');
      prevReg = i > 0 ? String(v[i][RG] || '') : '';
      if (b.action === 'memberDel') { if (i > 0) sh.deleteRow(i + 1); log_(u.username, 'memberDel', { id: m.id, code: m.code }); return { ok: true }; }
      if (!String(m.fn || m.nick || '').trim()) throw new Error('กรุณาใส่ชื่อหรือชื่อเล่น');
      if (!m.id) m.id = 'M' + Date.now().toString(36);
      if (m.photo && !(/^data:image\/jpeg;base64,/.test(m.photo) && m.photo.length < 48000)) m.photo = '';
      const row = MK.map(k => k === 'dob' && m.dob ? "'" + m.dob : m[k] === true ? '✓' : m[k] === false ? '' : (m[k] ?? '')).concat([u.name || u.username, new Date()]);
      if (i > 0) sh.getRange(i + 1, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
    } finally { lock.releaseLock(); }
    log_(u.username, 'memberSave', { id: m.id, code: m.code }); if (m.status === 'ย้ายไปต่างจังหวัด' && m.mvreg && m.mvreg !== prevReg) { const c0 = churches_().find(x => x.code === m.code); gpush_(`🚚 สมาชิกย้ายเข้าภาค${m.mvreg}\n👤 ${m.nick || m.fn || ''} (${m.mvstat || 'สมาชิก'}) จาก ${c0 ? c0.name : m.code}\n📍 ${m.mvloc || '-'} · ${m.mvwhat || '-'}${m.mvserve ? '\n⛪ เคยรับใช้: ' + m.mvserve : ''}\n🙏 รบกวนผู้ประสานงานภาค${m.mvreg} ติดต่อและเชื่อมต่อด้วยค่ะ`); return { id: m.id }; } { const cc = churches_().find(x => x.code === m.code); gpush_(`👥 ${isNew ? 'เพิ่ม' : 'แก้ไข'}ข้อมูลสมาชิก: ${m.nick || m.fn || ''}${m.status ? ' (' + m.status + ')' : ''}\n⛪ ${cc ? cc.name : m.code}\n✍️ โดย ${u.name || u.username}`); } return { id: m.id };
  }
  if (b.action === 'linkcode') {                       // รหัสเชื่อม LINE OA
    const code = String(Math.floor(100000 + Math.random() * 900000));
    cache().put('l:' + code, u.username, 600); return { code };
  }
  if (b.action === 'submit') {
    const c = C.find(x => x.code === b.code);
    if (!c || (!admin && !u.codes.includes(b.code))) throw new Error('ไม่มีสิทธิ์ส่งข้อมูลคริสตจักรนี้');
    const n = Number(b.n); if (!(n >= 0) || b.n === '' || b.n == null) throw new Error('จำนวนคนไม่ถูกต้อง');
    lock.waitLock(20000);
    try {
      const sh = ss().getSheetByName(SH.STAT), v = sh.getDataRange().getValues();
      const x = b.x || {}; if (sh.getRange(1, 11 + XF.length).getValue() !== XF[XF.length - 1][1]) sh.getRange(1, 12, 1, XF.length).setValues([XF.map(f => f[1])]);
      const row = [c.code, c.name, c.reg, round, n, b.nw || '', b.bp || '', b.lg || '', b.note || '', u.name || u.username, new Date()].concat(XF.map(([k]) => x[k] === true ? 1 : x[k] === false ? '' : (x[k] ?? '')));
      const i = v.findIndex(r => r[0] === c.code && String(r[3]).slice(0, 7) === round);
      if (i > 0) sh.getRange(i + 1, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
    } finally { lock.releaseLock(); }
    log_(u.username, 'submit', { code: c.code, round, n });
    const xx = b.x || {}; gpush_(`📊 ส่งสถิติรอบ ${round}\n⛪ ${c.name} · ${c.reg}\n👥 มา คจ. ${n} คน${xx.mem ? ' · สมาชิก ' + xx.mem : ''}${b.lg ? ' · LG ' + b.lg : ''}${b.nw ? '\n🌱 ผู้มาใหม่ ' + b.nw : ''}${b.bp ? ' · บัพติศมา ' + b.bp : ''}\n✍️ โดย ${u.name || u.username}`);
    return { saved: true };
  }
  if (b.action === 'pioConfirm') {
    if (!['admin', 'coord'].includes(u.role) || !u.codes.includes(b.code)) throw new Error('เฉพาะผู้ประสานงานภาคยืนยันได้');
    const sh = ss().getSheetByName(SH.REG), v = sh.getDataRange().getValues(), i = v.findIndex(r => r[0] === b.code); if (i < 1) throw new Error('ไม่พบคริสตจักร');
    const t = '✅ ' + (u.name || u.username) + ' · ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'd/M/yyyy');
    sh.getRange(1, 28).setValue('ยืนยันโดยผู้ประสานงาน'); sh.getRange(i + 1, 28).setValue(t); log_(u.username, 'pioConfirm', { code: b.code }); gpush_(`✅ ยืนยันข้อมูลกลุ่มบุกเบิก: ${v[i][4] || b.code}\n✍️ โดย ${u.name || u.username}`); return { pconf: t };
  }
  if (b.action === 'counselList') {
    if (!['admin', 'exec', 'coord'].includes(u.role)) throw new Error('ไม่มีสิทธิ์ดูคำขอปรึกษา');
    const sh = ss().getSheetByName('คำขอปรึกษา'); if (!sh) return [];
    return sh.getDataRange().getValues().slice(1).map((r, i) => ({ row: i + 2, at: r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Bangkok', 'd/M/yyyy HH:mm') : String(r[0]), name: r[1], church: r[2], topic: r[3], help: r[4], contact: r[5], urg: r[6], status: r[7], owner: r[8] })).reverse();
  }
  if (b.action === 'counselSet') {
    if (!['admin', 'coord'].includes(u.role)) throw new Error('ไม่มีสิทธิ์แก้สถานะ');
    const sh = ss().getSheetByName('คำขอปรึกษา'), row = Number(b.row);
    if (!sh || !(row >= 2 && row <= sh.getLastRow())) throw new Error('ไม่พบคำขอ');
    const st = ['รับเรื่อง', 'กำลังดูแล', 'เสร็จสิ้น'].includes(b.status) ? b.status : 'รับเรื่อง';
    sh.getRange(row, 8, 1, 2).setValues([[st, String(b.owner || '').trim().slice(0, 80)]]); log_(u.username, 'counselSet', { row, st }); return { ok: 1 };
  }
  if (b.action === 'updateInfo') {                     // ผู้นำแก้ข้อมูลทั่วไป + การขึ้นทะเบียน ศบ.
    if (!admin && !u.codes.includes(b.code)) throw new Error('ไม่มีสิทธิ์แก้ข้อมูลคริสตจักรนี้');
    const sh = ss().getSheetByName(SH.REG), v = sh.getDataRange().getValues();
    const i = v.findIndex(r => r[0] === b.code); if (i < 1) throw new Error('ไม่พบคริสตจักร');
    const clean = x => String(x == null ? '' : x).trim().slice(0, 120);
    const f = { 3: clean(b.prov), 6: clean(b.leader), 7: clean(b.shirt), 8: clean(b.phone), 9: clean(b.email), 10: b.reged ? 'ขึ้นทะเบียนแล้ว' : '', 11: b.reged ? clean(b.regdate) : '', 16: b.p111 ? '✓' : '', 17: b.plus ? '✓' : '', 18: String(b.addr || '').trim().slice(0, 300), 19: clean(b.map) };
    if (b.pio) { Object.assign(f, { 21: clean(b.pfull), 22: clean(b.freq), 23: clean(b.pstart), 24: clean(b.mother), 25: clean(b.meet), 26: clean(b.stage), 27: String(b.need || '').trim().slice(0, 300), 28: '' }); if (sh.getRange(1, 21).getValue() !== 'ผู้นำบุกเบิก (ชื่อ-สกุล)') sh.getRange(1, 21, 1, 7).setValues([['ผู้นำบุกเบิก (ชื่อ-สกุล)', 'ความถี่ลงพื้นที่', 'วันที่เริ่มบุกเบิก', 'คริสตจักรแม่', 'วัน-เวลานัดพบ', 'ระยะการเติบโต', 'สิ่งที่ต้องการสนับสนุน']]); }
    if (sh.getRange(1, 18).getValue() !== 'ที่อยู่') sh.getRange(1, 18, 1, 3).setValues([['ที่อยู่', 'แผนที่/พิกัด', 'รูป (Drive id)']]);
    if (f[9] && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(f[9])) throw new Error('อีเมลไม่ถูกต้อง');
    if (typeof b.photo === 'string' && /^data:image\/jpeg;base64,/.test(b.photo) && b.photo.length < 48000) f[20] = b.photo;
    const old = {}; lock.waitLock(20000);
    try { Object.keys(f).forEach(k => { old[k] = v[i][k - 1]; sh.getRange(i + 1, Number(k)).setValue(f[k]); }); } finally { lock.releaseLock(); }
    log_(u.username, 'updateInfo', { code: b.code, old, new: f });
    gpush_(`⛪ อัปเดตข้อมูล ${v[i][4] || b.code}${b.pio ? '\n🌱 กลุ่มบุกเบิก · ความถี่: ' + (b.freq || '-') + ' · ระยะ: ' + (b.stage || '-') + '\n⏳ รอผู้ประสานงานยืนยัน' : ''}\n✍️ โดย ${u.name || u.username}`);
    return { saved: true };
  }
  if (!admin) throw new Error('ต้องเป็นแอดมิน');
  if (b.action === 'users') return users_().map(pub_);
  if (b.action === 'createUser') {
    const name = String(b.username || '').trim();
    if (!/^[A-Za-z0-9_.-]{3,30}$/.test(name)) throw new Error('username ใช้ a-z 0-9 . _ - (3–30 ตัว)');
    if (users_().some(x => x.username.toLowerCase() === name.toLowerCase())) throw new Error('username นี้มีแล้ว');
    const role = ['admin', 'exec', 'coord'].includes(b.role) ? b.role : 'leader';
    const codes = String(b.codes || '').split(',').map(x => x.trim()).filter(String).map(k => k.startsWith('REG:') ? k : k.toUpperCase());
    codes.forEach(k => { if (k.startsWith('REG:') ? !C.some(c => String(c.reg) === k.slice(4)) : !C.find(c => c.code === k)) throw new Error('ไม่พบ ' + k); });
    ss().getSheetByName(SH.USER).appendRow([name, b.name || name, role, codes.join(','), 'active', new Date(), u.username, '', '', '']);
    const pw = rand_(10); setPw_(users_().find(x => x.username === name), pw);
    log_(u.username, 'createUser', { name, codes }); return { username: name, password: pw };
  }
  if (b.action === 'setAllPw') {
    if (String(b.password || '').length < 8) throw new Error('รหัสผ่านอย่างน้อย 8 ตัว');
    const L = users_().filter(x => ['coord', 'exec'].concat(b.leaders ? ['leader'] : []).includes(x.role) && x.status !== 'disabled'); L.forEach(t => setPw_(t, String(b.password)));
    log_(u.username, 'setAllPw', { n: L.length, leaders: !!b.leaders }); return { n: L.length };
  }
  if (b.action === 'setUserPw') {
    const t = users_().find(x => x.username === b.username); if (!t) throw new Error('ไม่พบผู้ใช้');
    if (String(b.password || '').length < 8) throw new Error('รหัสผ่านอย่างน้อย 8 ตัว');
    setPw_(t, String(b.password)); log_(u.username, 'setUserPw', { user: t.username }); return { ok: 1 };
  }
  if (b.action === 'resetPassword') {
    const t = users_().find(x => x.username === b.username); if (!t) throw new Error('ไม่พบผู้ใช้');
    const pw = rand_(10); setPw_(t, pw); log_(u.username, 'resetPassword', { user: t.username }); return { username: t.username, password: pw };
  }
  if (b.action === 'setStatus') {
    const t = users_().find(x => x.username === b.username); if (!t) throw new Error('ไม่พบผู้ใช้');
    if (t.username === u.username) throw new Error('ปิดบัญชีตัวเองไม่ได้');
    ss().getSheetByName(SH.USER).getRange(t.row, 5).setValue(b.status === 'active' ? 'active' : 'disabled');
    log_(u.username, 'setStatus', { user: t.username, status: b.status }); return {};
  }
  throw new Error('คำสั่งไม่ถูกต้อง');
}

/* ---------- LINE OA: ตอบคำถาม ---------- */
function webhook_(body) {
  body.events.forEach(ev => {
    const grp = ev.source && ev.source.type !== 'user';
    if (grp && !P.getProperty('NOTIFY_GROUP') && (ev.source.groupId || ev.source.roomId)) P.setProperty('NOTIFY_GROUP', ev.source.groupId || ev.source.roomId);
    if (ev.type === 'join') return reply_(ev.replyToken, 'สวัสดีค่ะ 🙏 เลขางานพันธกิจ H.I.M. ยินดีให้บริการค่ะ\nในกลุ่มนี้ ถ้าต้องการถามข้อมูล พิมพ์ขึ้นต้นด้วย "โรส" หรือแท็ก @ ถึงเลขาได้เลยค่ะ เช่น "โรส สรุปภาพรวม"');
    if (ev.type !== 'message' || ev.message.type !== 'text') return;
    let q = ev.message.text.trim(); const uid = ev.source.userId;
    if (grp) {
      const tagged = ((ev.message.mention || {}).mentionees || []).some(x => x.isSelf);
      if (!tagged && !/^(โรส|เลขา)/.test(q) && !/^(เชื่อม|link)\s*\d{6}$/i.test(q)) return;
      q = q.replace(/@\S+\s*/g, '').replace(/^(โรส|เลขา)\s*/, '').trim() || 'สรุปภาพรวม';
    }
    let text, u = users_().find(x => x.lineUserId === uid && x.status === 'active');
    const m = q.match(/^(เชื่อม|link)\s*(\d{6})$/i);
    if (m) {
      const name = cache().get('l:' + m[2]); const t = name && users_().find(x => x.username === name);
      if (!t) text = 'รหัสไม่ถูกต้องหรือหมดอายุ กรุณาขอรหัสใหม่จากเว็บแอป (เมนู บัญชีของฉัน)';
      else { ss().getSheetByName(SH.USER).getRange(t.row, 10).setValue(uid); cache().remove('l:' + m[2]); log_(t.username, 'linkLine', {}); text = `เชื่อมบัญชี ${t.name || t.username} เรียบร้อยแล้ว 🙏\nถามข้อมูลได้เลย เช่น "สรุปภาพรวม"`; }
    } else if (!u) text = 'สวัสดีค่ะ 🙏 งานพันธกิจ H.I.M.\nหากต้องการถามข้อมูลสถิติ กรุณาเชื่อมบัญชีก่อน: เข้าเว็บแอป → บัญชีของฉัน → "รหัสเชื่อม LINE" แล้วพิมพ์ เชื่อม ตามด้วยรหัส 6 หลักที่นี่';
    else text = (['admin', 'exec'].includes(u.role) && P.getProperty('ANTHROPIC_API_KEY')) ? askAI_(q, u) : answer_(q, u);
    reply_(ev.replyToken, text);
  });
  return ContentService.createTextOutput('ok');
}
function gpush_(text) { try { const g = P.getProperty('NOTIFY_GROUP'); if (g) push_(g, text); } catch (e) {} }
function push_(to, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + P.getProperty('LINE_ACCESS_TOKEN') },
    payload: JSON.stringify({ to, messages: [{ type: 'text', text: text.slice(0, 4900) }] })
  });
}
function reply_(token, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + P.getProperty('LINE_ACCESS_TOKEN') },
    payload: JSON.stringify({ replyToken: token, messages: [{ type: 'text', text: text.slice(0, 4900) }] })
  });
}
function answer_(q, u) {
  const C = churches_(), S = stats_(), round = P.getProperty('CURRENT_ROUND');
  const by = {}; S.forEach(s => (by[s.code] = by[s.code] || {})[s.round] = s.n);
  const rounds = [...new Set(S.map(s => s.round))].sort();
  const last = c => { const r = rounds.filter(x => by[c.code] && by[c.code][x] != null); return r.length ? by[c.code][r[r.length - 1]] : null; };
  const prev = c => { const r = rounds.filter(x => by[c.code] && by[c.code][x] != null); return r.length > 1 ? by[c.code][r[r.length - 2]] : null; };
  scope_(u, C); const pool0 = u.role === 'admin' ? C : C.filter(c => u.codes.includes(c.code));
  const regs = [...new Set(C.map(c => c.reg))];
  const g = regs.find(r => q.replace(/\s/g, '').includes(String(r).replace(/\s/g, '')));
  const pool = g ? pool0.filter(c => c.reg === g) : pool0;
  const sum = a => a.reduce((t, c) => t + (last(c) || 0), 0);
  // จังหวัด + ตัวชี้วัด (LG / ผู้มาใหม่ / บัพติศมา / จำนวนคน / จำนวนแห่ง)
  const provs = [...new Set(C.map(c => String(c.prov || '').trim()).filter(String))];
  const pv = provs.find(p => q.includes(p)) || (q.match(/(?:จ\.?|จังหวัด)\s*([ก-๙]+)/) || [])[1];
  const metric = /lg|life ?group|ไลฟ์|กลุ่มย่อย/i.test(q) ? ['lg', 'Life Group', 'กลุ่ม'] : /ผู้มาใหม่|คนใหม่/.test(q) ? ['nw', 'ผู้มาใหม่', 'คน'] : /บัพติศมา|บัพ/.test(q) ? ['bp', 'บัพติศมา', 'คน'] : null;
  if (pv || metric) {
    const area = pv ? pool0.filter(c => String(c.prov || '').includes(pv) || String(c.name).includes(pv)) : pool;
    const label = pv ? 'จ.' + pv : (g ? 'ภาค ' + g : 'ทั้งหมด');
    if (!area.length) return `ยังไม่พบคริสตจักรใน ${label} ในระบบ\n(ข้อมูลจังหวัดบางแห่งยังไม่ครบ ผู้นำเติมได้ที่เมนู ข้อมูลของฉัน → แก้ไขข้อมูล)`;
    if (metric) {
      const k = metric[0], lastV = c => { const a = S.filter(s => s.code === c.code && s[k] !== '' && s[k] != null).sort((x, y) => x.round < y.round ? 1 : -1); return a.length ? Number(a[0][k]) || 0 : null; };
      const rows = area.map(c => [c, lastV(c)]).filter(r => r[1] != null);
      const tot = rows.reduce((t, r) => t + r[1], 0);
      return `📊 ${metric[1]} ${label}\nรวม ${tot} ${metric[2]} (จาก ${rows.length}/${area.length} แห่งที่รายงาน)\n` + rows.filter(r => r[1] > 0).map(r => `• ${r[0].name}: ${r[1]}`).join('\n') + (rows.length < area.length ? `\n\nยังไม่รายงาน ${area.length - rows.length} แห่ง` : '');
    }
    const t2 = k => area.filter(c => c.type === k).length;
    return `📍 ${label}: ${area.length} แห่ง\nคจ ${t2('คริสตจักร')} · ศูนย์ ${t2('ศูนย์')} · บุกเบิก ${t2('กลุ่มบุกเบิก')}\nยอดล่าสุดรวม ${sum(area)} คน\n` + area.map(c => `• ${c.name} (${last(c) ?? '-'})`).join('\n');
  }
  if (/ยังไม่ส่ง/.test(q)) { const a = pool.filter(c => !(by[c.code] || {})[round]); return `ยังไม่ส่งรอบ ${round}: ${a.length} แห่ง\n` + a.slice(0, 30).map(c => `• ${c.name} (${c.reg})`).join('\n'); }
  if (/ลด/.test(q)) { const a = pool.filter(c => prev(c) >= 15 && last(c) < prev(c) * 0.8); return `ยอดลดเกิน 20%: ${a.length} แห่ง\n` + a.map(c => `• ${c.name} ${prev(c)}→${last(c)}`).join('\n'); }
  if (/XS|ต่ำกว่า|เล็ก/.test(q)) { const a = pool.filter(c => last(c) != null && last(c) < 30); return `กำลังเติบโต (<30 คน): ${a.length} แห่ง\n` + a.slice(0, 30).map(c => `• ${c.name} (${last(c)})`).join('\n'); }
  const c = pool0.find(x => q.includes(String(x.name).replace(/^(คริสตจักร|ศูนย์สร้างสรรค์ชีวิต)/, '').trim()));
  if (c && !/สรุป/.test(q)) return `${c.name} (${c.code})\nภาค ${c.reg} · ${c.type}\nผู้นำ: ${c.leader}\n` + rounds.filter(r => (by[c.code] || {})[r] != null).map(r => `${r}: ${by[c.code][r]} คน`).join('\n');
  const t = k => pool.filter(c => c.type === k).length;
  return `📊 สรุป${g ? 'ภาค ' + g : (u.role === 'admin' ? 'ภาพรวม' : 'คริสตจักรของท่าน')}\nคจ ${t('คริสตจักร')} · ศูนย์ ${t('ศูนย์')} · บุกเบิก ${t('กลุ่มบุกเบิก')}\nยอดล่าสุดรวม ${sum(pool)} คน\nส่งรอบ ${round} แล้ว ${pool.filter(c => (by[c.code] || {})[round] != null).length}/${pool.length}\n\nถามได้: "สรุปภาค อีสาน 4", "ยังไม่ส่ง", "ยอดลด", "XS", หรือชื่อคริสตจักร`;
}

/* ---------- AI เลขา: อ่านทุกแท็บใน Sheet แล้วตอบคำถาม (เฉพาะแอดมิน) ---------- */
function sheetsContext_() {
  const skip = [SH.USER, SH.LOG, 'สมาชิก'];
  let out = '';
  const books = [[ss(), 'ฐานข้อมูลกลาง']];
  try { books.push([SpreadsheetApp.openById(P.getProperty('SOURCE_SHEET_ID') || '1QV6Ft7ED8eK0zE_BcZm2TJDtih9BXP8XZZpu7OTykJw'), 'ชีตต้นฉบับงานพันธกิจ']); } catch (e) {}
  books.forEach(([book, label]) => book.getSheets().forEach(sh => {
    if (skip.includes(sh.getName())) return;
    const v = sh.getDataRange().getDisplayValues().filter(r => r.some(String));
    out += `\n### [${label}] แท็บ: ${sh.getName()} (${v.length - 1} แถว)\n` + v.map(r => r.map(x => String(x).replace(/[\t\n]/g, ' ')).join('\t')).join('\n') + '\n';
  }));
  return out.slice(0, 180000);
}
function askAI_(q, u) {
  const c = cache(), key = 'ctx';
  let ctx = c.get(key);
  if (!ctx) { ctx = sheetsContext_(); try { c.put(key, ctx, 300); } catch (e) {} }
  const sys = 'คุณคือ "เลขางานพันธกิจ H.I.M. Thailand" ตอบคำถามฝ่ายบริหารจากข้อมูลใน Google Sheet ที่ให้มาเท่านั้น ' +
    'บุคลิก: เป็นเลขาผู้หญิงที่อบอุ่น สุภาพ คุยเป็นธรรมชาติเหมือนมนุษย์ ลงท้ายประโยคด้วย "ค่ะ" (หรือ "คะ" เมื่อเป็นคำถาม) ทุกประโยค ไม่ใช้ "ครับ" ' +
    'ตอบภาษาไทย สั้น กระชับ เหมือนพิมพ์แชท LINE จริง (ไม่ใช้ markdown ตาราง หัวข้อ หรือสัญลักษณ์เยอะ) ใส่ตัวเลขที่นับจริงจากข้อมูล ' +
    'ข้อมูลมาจาก 2 แหล่ง: ฐานข้อมูลกลาง (ทะเบียน/สถิติที่ผู้นำกรอก) และชีตต้นฉบับงานพันธกิจ (สรุปรายภาค เป้าหมาย โครงการ ผู้สมัครบุกเบิก ฯลฯ) ให้ค้นทุกแท็บและเชื่อมโยงกัน เช่น จังหวัด→ภาค ' +
    'ห้ามตอบเพียงว่า "ไม่มีข้อมูล" ให้หาข้อมูลที่ใกล้เคียงที่สุดมาตอบเสมอ แล้วบอกอย่างเป็นธรรมชาติตั้งแต่ประโยคแรกว่าตรงกับที่ถามเลย หรือเป็นข้อมูลใกล้เคียง เช่น "ตรงกับที่ถามเลยค่ะ" หรือ "ยังไม่มีข้อมูลตรงๆ ค่ะ แต่ที่ใกล้เคียงที่สุดคือ..." ' +
    'ถ้าใกล้เคียง ให้อธิบายสั้นๆ ว่าต่างจากที่ถามอย่างไร (เช่น เป็นเป้าหมายไม่ใช่ยอดจริง, เป็นระดับภาคไม่ใช่จังหวัด) และระบุแท็บที่ใช้ ห้ามแต่งตัวเลขเอง ' +
    'นิยามของงานพันธกิจ (จากแท็บแรกของชีตต้นฉบับ): ขนาดตามจำนวนผู้มานมัสการ XS = ต่ำกว่า 30 คน, S = 30-50, M = 51-100, L = 101 ขึ้นไป; "คจ มาตรฐาน" = 30 คนขึ้นไป; "ศูนย์ที่มีศักยภาพ" = ศูนย์ที่กำลังเติบโตสู่ คจ มาตรฐาน; "คจ ที่ต่ำกว่ามาตรฐาน" = คจ ที่ยังไม่ถึง 30 คน ให้ใช้นิยามนี้เสมอ ' +
    'เมื่อถูกขอ "รายงาน" หรือ "ข้อมูลทั้งหมด" ของคริสตจักรใด ให้รวบรวมให้ครบทุกแหล่งที่พบชื่อหรือรหัสนั้น เรียงเป็นช่วงสั้นๆ: ข้อมูลทั่วไป (ภาค จังหวัด ประเภท ผู้นำ เบอร์ สถานะ) / ยอดนมัสการล่าสุดและขนาด / สถิติรายไตรมาส / นิมิตเป้าหมาย / โครงการ 111 พลัส ศบ. / หมายเหตุ ถ้าช่องไหนว่างให้บอกว่า "ยังไม่ได้ระบุค่ะ" ห้ามจบประโยคค้างกลางคัน ' +
    `รอบรายงานปัจจุบันคือ ${P.getProperty('CURRENT_ROUND')} ใช้ภาษาที่สุภาพและให้กำลังใจ ใช้คำว่า "กำลังเติบโต" แทน "ต่ำกว่ามาตรฐาน"`;
  const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': P.getProperty('ANTHROPIC_API_KEY'), 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: P.getProperty('AI_MODEL') || 'claude-haiku-5-5', max_tokens: 8000, system: sys,
      messages: [{ role: 'user', content: 'ข้อมูลทั้งหมด (คั่นคอลัมน์ด้วย tab):\n' + ctx + '\n\nคำถาม: ' + q }] })
  });
  if (r.getResponseCode() !== 200) { log_(u.username, 'aiError', { code: r.getResponseCode(), body: r.getContentText().slice(0, 300) }); return answer_(q, u); }
  const j = JSON.parse(r.getContentText());
  log_(u.username, 'askAI', { q, stop: j.stop_reason, in: j.usage && j.usage.input_tokens, out: j.usage && j.usage.output_tokens });
  const txt = (j.content || []).map(x => x.text || '').join('').replace(/\*\*|__|^#+\s*/gm, '').trim();
  return txt || answer_(q, u);
}

/** ซิงก์ทะเบียน: ฐานข้อมูลกลาง → ชีต v2 แท็บ 1-5 (รันอัตโนมัติทุกชั่วโมง) จับคู่ด้วยคอลัมน์ "รหัส" */
function syncV2() {
  const ssv = SpreadsheetApp.openById('1YwI2oUb08zAVavO4o4moHpvjecQI0uTThvnul4wZDVg');
  const T = { 'คริสตจักร': 'คจ' }, C = {}; churches_().forEach(c => C[c.code] = c);
  const want = c => [String(c.reg || '').replace(/\s/g, ''), c.prov, T[c.type] || c.type, c.name, c.leader, '', c.phone, c.email];
  let upd = 0, add = 0;
  ssv.getSheets().filter(s => /^[1-5] /.test(s.getName())).forEach(sh => {
    const main = /^1 /.test(sh.getName());
    const v = sh.getDataRange().getValues(); const h = v.findIndex(r => String(r[0]).trim() === 'ลำดับ'); if (h < 0) return;
    let KC = v[h].indexOf('รหัส'); if (KC < 0) { KC = v[h].length; sh.getRange(h + 1, KC + 1).setValue('รหัส'); }
    let EC = v[h].indexOf('ชื่อคริสตจักร (EN)'); if (EC < 0) { EC = Math.max(v[h].length, KC + 1); sh.getRange(h + 1, EC + 1, 1, 2).setValues([['ชื่อคริสตจักร (EN)', 'ชื่อผู้นำ (EN)']]); }
    const pos = {}; let last = h, maxSeq = 0;
    for (let i = h + 1; i < v.length; i++) { if (v[i][0] === '') continue; last = i; maxSeq = Math.max(maxSeq, +v[i][0] || 0); const k = String(v[i][KC] || '').trim(); if (k) pos[k] = i; }
    Object.keys(pos).forEach(k => { const c = C[k]; if (!c) return; const i = pos[k];
      want(c).forEach((x, j) => { if (x !== '' && x != null && String(v[i][j + 1]).trim() !== String(x).trim()) { sh.getRange(i + 1, j + 2).setValue(x); upd++; } });
      [c.en, c.enl].forEach((x, j) => { if (x && String(v[i][EC + j] || '').trim() !== String(x).trim()) { sh.getRange(i + 1, EC + j + 1).setValue(x); upd++; } }); });
    if (main) Object.keys(C).filter(k => pos[k] == null).forEach(k => {
      sh.insertRowsAfter(last + 1, 1); const r = last + 2, W = sh.getLastColumn();
      sh.getRange(last + 1, 1, 1, W).copyTo(sh.getRange(r, 1, 1, W));
      for (let col = 2; col <= W; col++) { const cell = sh.getRange(r, col); if (!cell.getFormula()) cell.clearContent(); }
      sh.getRange(r, 1).setValue(++maxSeq); sh.getRange(r, KC + 1).setValue(k);
      sh.getRange(r, 2, 1, 8).setValues([want(C[k]).map(x => x == null ? '' : x)]); last++; add++;
    });
    if (main) sh.getRange(h - 1, 1).setValue('ทะเบียนหลักทุกแห่ง (ซิงก์อัตโนมัติจากฐานข้อมูลกลาง ล่าสุด ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'd/M/yyyy HH:mm') + ') · ขนาด XS<30, S 30-50, M 51-100, L 101+ · มาตรฐาน = 30 คนขึ้นไป · ยอดล่าสุด/ขนาด/มาตรฐาน คำนวณอัตโนมัติ');
  });
  Logger.log('syncV2 upd ' + upd + ' add ' + add);
}

/** แก้ในชีต v2 (แท็บ 1-5 คอลัมน์ ภาค..อีเมล) → เขียนกลับฐานข้อมูลกลางทันที แล้วซิงก์แท็บอื่นให้ตรงกัน */
function onV2Edit(e) {
  try {
    const sh = e.range.getSheet(); if (!/^[1-5] /.test(sh.getName())) return;
    const v = sh.getDataRange().getValues(); const h = v.findIndex(r => String(r[0]).trim() === 'ลำดับ'); if (h < 0) return;
    const KC = v[h].indexOf('รหัส'); if (KC < 0) return;
    const F = { 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9 }; // คอลัมน์ v2 (B..I) → คอลัมน์ฐานข้อมูลกลาง
    const CDB = SpreadsheetApp.openById('1WRbfK1_dg_rUDTlRiZdYIAFxhi5k2fnPvT9sBr5nppc'), reg = CDB.getSheetByName(SH.REG), rv = reg.getDataRange().getValues(); let n = 0;
    for (let r = e.range.getRow(); r <= e.range.getLastRow(); r++) {
      const code = String(v[r - 1] && v[r - 1][KC] || '').trim(); if (!code || r - 1 <= h) continue;
      const i = rv.findIndex(x => x[0] === code); if (i < 1) continue;
      for (let c = e.range.getColumn(); c <= e.range.getLastColumn(); c++) {
        if (!F[c]) continue; let x = v[r - 1][c - 1];
        if (c === 2) x = String(x).replace(/([^\d\s])(\d)/, '$1 $2');
        if (c === 4 && x === 'คจ') x = 'คริสตจักร';
        if (String(rv[i][F[c] - 1]).trim() !== String(x).trim()) { reg.getRange(i + 1, F[c]).setValue(x); n++; }
      }
    }
    if (n) { CDB.getSheetByName(SH.LOG).appendRow([new Date(), 'v2-sheet', 'v2Edit', JSON.stringify({ tab: sh.getName(), n })]); syncV2(); }
  } catch (err) { console.error(err); }
}
