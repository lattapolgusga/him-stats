/**
 * HIM Stats — ระบบหลังบ้าน (Google Apps Script)
 * วางไฟล์นี้ใน Extensions > Apps Script ของ Google Sheet กลาง
 * แล้วตั้ง Script Properties (Project Settings > Script properties):
 *   LOGIN_CHANNEL_ID   = Channel ID ของ LINE Login "HIM Stats App" (2012642298)
 *   LINE_ACCESS_TOKEN  = Channel access token ของ OA (สำหรับตอบแชท)
 *   CURRENT_ROUND      = รอบที่เปิดให้ส่ง เช่น 2026-Q4
 */
const SH = { REG: 'ทะเบียนคริสตจักร', STAT: 'สถิติ', USER: 'ผู้ใช้', LOG: 'บันทึกการแก้ไข' };
const P = PropertiesService.getScriptProperties();
const ss = () => SpreadsheetApp.getActiveSpreadsheet();

/** รันครั้งเดียวหลังวางโค้ด: สร้างแท็บผู้ใช้/บันทึก และหัวคอลัมน์สถิติเพิ่ม */
function setup() {
  const s = ss();
  if (!s.getSheetByName(SH.USER)) s.insertSheet(SH.USER).appendRow(['lineUserId','ชื่อ LINE','บทบาท(leader/admin)','รหัสคริสตจักร(คั่นด้วย ,)','สถานะ(pending/active)','ขอเมื่อ','อนุมัติโดย']);
  if (!s.getSheetByName(SH.LOG)) s.insertSheet(SH.LOG).appendRow(['เวลา','lineUserId','การกระทำ','รายละเอียด']);
  const st = s.getSheetByName(SH.STAT);
  st.getRange(1, 1, 1, 11).setValues([['รหัส','ชื่อ','ภาค','รอบ','จำนวนคน','ผู้มาใหม่','บัพติศมา','Life Group','อธิษฐาน/ข่าว','ส่งโดย','เวลาส่ง']]);
  if (!P.getProperty('LOGIN_CHANNEL_ID')) P.setProperty('LOGIN_CHANNEL_ID', '2012642298');
  if (!P.getProperty('CURRENT_ROUND')) P.setProperty('CURRENT_ROUND', '2026-Q4');
}

/* ---------- ข้อมูล ---------- */
function rows_(name) { const v = ss().getSheetByName(name).getDataRange().getValues(); v.shift(); return v; }
function churches_() {
  return rows_(SH.REG).filter(r => r[0]).map(r => ({ code: r[0], reg: r[1], prov: r[2], type: r[3], name: r[4], leader: r[5], status: r[11] }));
}
function stats_() {
  return rows_(SH.STAT).filter(r => r[0]).map(r => ({ code: r[0], round: String(r[3]).slice(0, 7), n: Number(r[4]), nw: r[5], bp: r[6], lg: r[7] }));
}
function users_() { return rows_(SH.USER).map((r, i) => ({ row: i + 2, uid: r[0], name: r[1], role: r[2], codes: String(r[3]).split(',').map(x => x.trim()).filter(String), status: r[4] })); }
function userOf_(uid) { return users_().find(u => u.uid === uid); }
function log_(uid, act, d) { ss().getSheetByName(SH.LOG).appendRow([new Date(), uid, act, JSON.stringify(d)]); }

/* ---------- ยืนยันตัวตนด้วย LINE ID token ---------- */
function verify_(idToken) {
  const r = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post', muteHttpExceptions: true,
    payload: { id_token: idToken, client_id: P.getProperty('LOGIN_CHANNEL_ID') }
  });
  if (r.getResponseCode() !== 200) throw new Error('LINE login หมดอายุ กรุณาเปิดใหม่');
  const j = JSON.parse(r.getContentText());
  return { uid: j.sub, name: j.name };
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
  const me = verify_(b.idToken), u = userOf_(me.uid), round = P.getProperty('CURRENT_ROUND');
  const C = churches_();
  if (b.action === 'me') {
    if (!u) return { state: 'new', name: me.name, round };
    if (u.status !== 'active') return { state: 'pending', name: me.name, round };
    const mine = u.role === 'admin' ? [] : C.filter(c => u.codes.includes(c.code));
    const S = stats_().filter(s => u.role === 'admin' || u.codes.includes(s.code));
    return { state: 'active', role: u.role, name: me.name, round, churches: u.role === 'admin' ? C : mine, stats: S };
  }
  if (b.action === 'request') {                       // ผูกบัญชีครั้งแรก
    if (u) throw new Error('ส่งคำขอไปแล้ว รอแอดมินอนุมัติ');
    const code = String(b.code || '').trim().toUpperCase();
    if (!C.find(c => c.code === code)) throw new Error('ไม่พบรหัสคริสตจักรนี้');
    ss().getSheetByName(SH.USER).appendRow([me.uid, me.name, 'leader', code, 'pending', new Date(), '']);
    log_(me.uid, 'request', { code });
    return { state: 'pending' };
  }
  if (!u || u.status !== 'active') throw new Error('บัญชียังไม่ได้รับอนุมัติ');
  if (b.action === 'submit') {
    const c = C.find(x => x.code === b.code);
    if (!c || (u.role !== 'admin' && !u.codes.includes(b.code))) throw new Error('ไม่มีสิทธิ์ส่งข้อมูลคริสตจักรนี้');
    const n = Number(b.n); if (!(n >= 0)) throw new Error('จำนวนคนไม่ถูกต้อง');
    const sh = ss().getSheetByName(SH.STAT), v = sh.getDataRange().getValues();
    const row = [c.code, c.name, c.reg, round, n, b.nw || '', b.bp || '', b.lg || '', b.note || '', me.name, new Date()];
    const i = v.findIndex(r => r[0] === c.code && String(r[3]).slice(0, 7) === round);
    if (i > 0) sh.getRange(i + 1, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
    log_(me.uid, 'submit', { code: c.code, round, n });
    return { saved: true };
  }
  if (b.action === 'pending' && u.role === 'admin') return users_().filter(x => x.status === 'pending');
  if (b.action === 'approve' && u.role === 'admin') {
    const t = users_().find(x => x.uid === b.uid); if (!t) throw new Error('ไม่พบผู้ใช้');
    const us = ss().getSheetByName(SH.USER);
    us.getRange(t.row, 5).setValue('active'); us.getRange(t.row, 7).setValue(me.name);
    log_(me.uid, 'approve', { uid: b.uid });
    return { approved: true };
  }
  throw new Error('คำสั่งไม่ถูกต้อง');
}

/* ---------- LINE OA: ตอบคำถาม ---------- */
function webhook_(body) {
  body.events.forEach(ev => {
    if (ev.type !== 'message' || ev.message.type !== 'text') return;
    const u = userOf_(ev.source.userId);
    let text;
    if (!u || u.status !== 'active') text = 'สวัสดีครับ 🙏 กดเมนู "อัปเดทสถิติ" เพื่อลงทะเบียนคริสตจักรของท่าน';
    else text = answer_(ev.message.text, u);
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
  if (/ยังไม่ส่ง/.test(q)) { const a = pool.filter(c => !(by[c.code] || {})[round]); return `ยังไม่ส่งรอบ ${round}: ${a.length} แห่ง\n` + a.slice(0, 30).map(c => `• ${c.name} (${c.reg})`).join('\n'); }
  if (/ลด/.test(q)) { const a = pool.filter(c => prev(c) >= 15 && last(c) < prev(c) * 0.8); return `ยอดลดเกิน 20%: ${a.length} แห่ง\n` + a.map(c => `• ${c.name} ${prev(c)}→${last(c)}`).join('\n'); }
  if (/XS|ต่ำกว่า|เล็ก/.test(q)) { const a = pool.filter(c => last(c) != null && last(c) < 30); return `กำลังเติบโต (<30 คน): ${a.length} แห่ง\n` + a.slice(0, 30).map(c => `• ${c.name} (${last(c)})`).join('\n'); }
  const c = pool0.find(x => q.includes(String(x.name).replace(/^(คริสตจักร|ศูนย์สร้างสรรค์ชีวิต)/, '').trim()));
  if (c && !/สรุป/.test(q)) return `${c.name} (${c.code})\nภาค ${c.reg} · ${c.type}\nผู้นำ: ${c.leader}\n` + rounds.filter(r => (by[c.code] || {})[r] != null).map(r => `${r}: ${by[c.code][r]} คน`).join('\n');
  const t = k => pool.filter(c => c.type === k).length;
  return `📊 สรุป${g ? 'ภาค ' + g : (u.role === 'admin' ? 'ภาพรวม' : 'คริสตจักรของท่าน')}\nคจ ${t('คริสตจักร')} · ศูนย์ ${t('ศูนย์')} · บุกเบิก ${t('กลุ่มบุกเบิก')}\nยอดล่าสุดรวม ${sum(pool)} คน\nส่งรอบ ${round} แล้ว ${pool.filter(c => (by[c.code] || {})[round] != null).length}/${pool.length}\n\nถามได้: "สรุปภาค อีสาน 4", "ยังไม่ส่ง", "ยอดลด", "XS", หรือชื่อคริสตจักร`;
}
