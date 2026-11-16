/**
 * HIM Stats — ระบบหลังบ้าน (Google Apps Script)
 * วางไฟล์นี้ใน Extensions > Apps Script ของ Google Sheet กลาง
 * แล้วตั้ง Script Properties (Project Settings > Script properties):
 *   LINE_ACCESS_TOKEN  = Channel access token ของ OA (สำหรับตอบแชท)
 *   CURRENT_ROUND      = รอบที่เปิดให้ส่ง เช่น 2026-Q4
 */
const SH = { REG: 'ทะเบียนคริสตจักร', STAT: 'สถิติ', USER: 'ผู้ใช้', LOG: 'บันทึกการแก้ไข' };
const P = PropertiesService.getScriptProperties();
const ss = () => SpreadsheetApp.getActiveSpreadsheet();

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
  return rows_(SH.REG).filter(r => r[0]).map(r => ({ code: r[0], reg: r[1], prov: r[2], type: r[3], name: r[4], leader: r[5], shirt: r[6], phone: r[7], email: r[8], reged: r[9], regdate: r[10], status: r[11], note: r[14], p111: !!r[15], plus: !!r[16] }));
}
function stats_() {
  return rows_(SH.STAT).filter(r => r[0]).map(r => ({ code: r[0], round: String(r[3]).slice(0, 7), n: Number(r[4]), nw: r[5], bp: r[6], lg: r[7], note: r[8], by: r[9], at: r[10] }));
}
function users_() {
  return rows_(SH.USER).map((r, i) => ({ row: i + 2, username: String(r[0]).trim(), name: r[1], role: r[2], codes: String(r[3]).split(',').map(x => x.trim()).filter(String), status: r[4], hash: r[7], salt: r[8], lineUserId: r[9] })).filter(u => u.username);
}
function log_(who, act, d) { ss().getSheetByName(SH.LOG).appendRow([new Date(), who, act, JSON.stringify(d)]); }
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
  if (b.action === 'status') return { needsAdmin: !users_().some(u => u.role === 'admin' && u.hash) };
  if (b.action === 'bootstrap') {                     // สร้างแอดมินคนแรก (ใช้ได้ครั้งเดียว)
    if (users_().some(u => u.role === 'admin' && u.hash)) throw new Error('มีแอดมินแล้ว');
    if (!b.username || String(b.password || '').length < 8) throw new Error('รหัสผ่านอย่างน้อย 8 ตัว');
    ss().getSheetByName(SH.USER).appendRow([String(b.username).trim(), b.name || b.username, 'admin', '', 'active', new Date(), 'bootstrap', '', '', '']);
    setPw_(users_().find(u => u.username === String(b.username).trim()), b.password);
    log_(b.username, 'bootstrap', {});
    b.action = 'login';
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
  const admin = u.role === 'admin';
  if (b.action === 'me') {
    const S = stats_().filter(s => admin || u.codes.includes(s.code));
    return { user: pub_(u), round, churches: admin ? C : C.filter(c => u.codes.includes(c.code)), stats: S };
  }
  if (b.action === 'logout') { cache().remove('s:' + b.token); return {}; }
  if (b.action === 'password') {
    if (hash_(String(b.old || ''), u.salt) !== u.hash) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
    if (String(b.password || '').length < 8) throw new Error('รหัสผ่านใหม่อย่างน้อย 8 ตัว');
    setPw_(u, b.password); log_(u.username, 'password', {}); return {};
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
      const row = [c.code, c.name, c.reg, round, n, b.nw || '', b.bp || '', b.lg || '', b.note || '', u.name || u.username, new Date()];
      const i = v.findIndex(r => r[0] === c.code && String(r[3]).slice(0, 7) === round);
      if (i > 0) sh.getRange(i + 1, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
    } finally { lock.releaseLock(); }
    log_(u.username, 'submit', { code: c.code, round, n });
    return { saved: true };
  }
  if (b.action === 'updateInfo') {                     // ผู้นำแก้ข้อมูลทั่วไป + การขึ้นทะเบียน ศบ.
    if (!admin && !u.codes.includes(b.code)) throw new Error('ไม่มีสิทธิ์แก้ข้อมูลคริสตจักรนี้');
    const sh = ss().getSheetByName(SH.REG), v = sh.getDataRange().getValues();
    const i = v.findIndex(r => r[0] === b.code); if (i < 1) throw new Error('ไม่พบคริสตจักร');
    const clean = x => String(x == null ? '' : x).trim().slice(0, 120);
    const f = { 3: clean(b.prov), 6: clean(b.leader), 7: clean(b.shirt), 8: clean(b.phone), 9: clean(b.email), 10: b.reged ? 'ขึ้นทะเบียนแล้ว' : '', 11: b.reged ? clean(b.regdate) : '', 16: b.p111 ? '✓' : '', 17: b.plus ? '✓' : '' };
    if (f[9] && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(f[9])) throw new Error('อีเมลไม่ถูกต้อง');
    const old = {}; lock.waitLock(20000);
    try { Object.keys(f).forEach(k => { old[k] = v[i][k - 1]; sh.getRange(i + 1, Number(k)).setValue(f[k]); }); } finally { lock.releaseLock(); }
    log_(u.username, 'updateInfo', { code: b.code, old, new: f });
    return { saved: true };
  }
  if (!admin) throw new Error('ต้องเป็นแอดมิน');
  if (b.action === 'users') return users_().map(pub_);
  if (b.action === 'createUser') {
    const name = String(b.username || '').trim();
    if (!/^[A-Za-z0-9_.-]{3,30}$/.test(name)) throw new Error('username ใช้ a-z 0-9 . _ - (3–30 ตัว)');
    if (users_().some(x => x.username.toLowerCase() === name.toLowerCase())) throw new Error('username นี้มีแล้ว');
    const codes = String(b.codes || '').toUpperCase().split(',').map(x => x.trim()).filter(String);
    codes.forEach(k => { if (!C.find(c => c.code === k)) throw new Error('ไม่พบรหัส ' + k); });
    ss().getSheetByName(SH.USER).appendRow([name, b.name || name, b.role === 'admin' ? 'admin' : 'leader', codes.join(','), 'active', new Date(), u.username, '', '', '']);
    const pw = rand_(10); setPw_(users_().find(x => x.username === name), pw);
    log_(u.username, 'createUser', { name, codes }); return { username: name, password: pw };
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
    if (ev.type !== 'message' || ev.message.type !== 'text') return;
    const q = ev.message.text.trim(), uid = ev.source.userId;
    let text, u = users_().find(x => x.lineUserId === uid && x.status === 'active');
    const m = q.match(/^(เชื่อม|link)\s*(\d{6})$/i);
    if (m) {
      const name = cache().get('l:' + m[2]); const t = name && users_().find(x => x.username === name);
      if (!t) text = 'รหัสไม่ถูกต้องหรือหมดอายุ กรุณาขอรหัสใหม่จากเว็บแอป (เมนู บัญชีของฉัน)';
      else { ss().getSheetByName(SH.USER).getRange(t.row, 10).setValue(uid); cache().remove('l:' + m[2]); log_(t.username, 'linkLine', {}); text = `เชื่อมบัญชี ${t.name || t.username} เรียบร้อยแล้ว 🙏\nถามข้อมูลได้เลย เช่น "สรุปภาพรวม"`; }
    } else if (!u) text = 'สวัสดีครับ 🙏 งานพันธกิจ H.I.M.\nหากต้องการถามข้อมูลสถิติ กรุณาเชื่อมบัญชีก่อน: เข้าเว็บแอป → บัญชีของฉัน → "รหัสเชื่อม LINE" แล้วพิมพ์ เชื่อม ตามด้วยรหัส 6 หลักที่นี่';
    else text = (u.role === 'admin' && P.getProperty('ANTHROPIC_API_KEY')) ? askAI_(q, u) : answer_(q, u);
    reply_(ev.replyToken, text);
  });
  return ContentService.createTextOutput('ok');
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
  const pool0 = u.role === 'admin' ? C : C.filter(c => u.codes.includes(c.code));
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
  const skip = [SH.USER, SH.LOG];
  let out = '';
  ss().getSheets().forEach(sh => {
    if (skip.includes(sh.getName())) return;
    const v = sh.getDataRange().getDisplayValues().filter(r => r.some(String));
    out += `\n### แท็บ: ${sh.getName()} (${v.length - 1} แถว)\n` + v.map(r => r.map(x => String(x).replace(/[\t\n]/g, ' ')).join('\t')).join('\n') + '\n';
  });
  return out.slice(0, 180000);
}
function askAI_(q, u) {
  const c = cache(), key = 'ctx';
  let ctx = c.get(key);
  if (!ctx) { ctx = sheetsContext_(); try { c.put(key, ctx, 300); } catch (e) {} }
  const sys = 'คุณคือ "เลขางานพันธกิจ H.I.M. Thailand" ตอบคำถามฝ่ายบริหารจากข้อมูลใน Google Sheet ที่ให้มาเท่านั้น ' +
    'ตอบภาษาไทย สั้น กระชับ เป็นข้อความธรรมดาสำหรับแชท LINE (ไม่ใช้ markdown ตาราง) ใส่ตัวเลขที่นับจริงจากข้อมูล ' +
    'ถ้าข้อมูลไม่พอหรือช่องว่าง ให้บอกตรงๆ ว่าข้อมูลส่วนไหนยังไม่ครบ ห้ามเดา ' +
    `รอบรายงานปัจจุบันคือ ${P.getProperty('CURRENT_ROUND')} ใช้ภาษาที่สุภาพและให้กำลังใจ ใช้คำว่า "กำลังเติบโต" แทน "ต่ำกว่ามาตรฐาน"`;
  const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': P.getProperty('ANTHROPIC_API_KEY'), 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: P.getProperty('AI_MODEL') || 'claude-haiku-5-5', max_tokens: 1200, system: sys,
      messages: [{ role: 'user', content: 'ข้อมูลทั้งหมด (คั่นคอลัมน์ด้วย tab):\n' + ctx + '\n\nคำถาม: ' + q }] })
  });
  if (r.getResponseCode() !== 200) { log_(u.username, 'aiError', { code: r.getResponseCode(), body: r.getContentText().slice(0, 300) }); return answer_(q, u); }
  const j = JSON.parse(r.getContentText());
  log_(u.username, 'askAI', { q, in: j.usage && j.usage.input_tokens, out: j.usage && j.usage.output_tokens });
  return (j.content || []).map(x => x.text || '').join('').trim() || answer_(q, u);
}
