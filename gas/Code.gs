/**
 * @OnlyCurrentDoc
 * 揪團點餐 — Google Apps Script 後端（含 LINE Bot）
 *
 * 試算表只當資料庫：店家、菜單、選項、成員、開團、訂單全部由網頁的「管理」維護。
 * 平常不用打開，網頁壞掉時才當備援直接看或改。
 *
 * 指令碼屬性（專案設定 → 指令碼屬性）：
 *   LINE_TOKEN        LINE Messaging API 的 Channel access token
 *   SITE_URL          點餐網頁網址（LINE 回覆裡的按鈕會用到）
 *   PUSH_ON_DEADLINE  true = 截止時自動推播統計到群組（會用到每月免費推播額度）
 *   LINE_GROUP_ID     不用填，Bot 在群組收到訊息時會自動記下
 */

const PIN = '2750';
const KEEP_DAYS = 3;          // 網頁只顯示最近幾天的團
const TZ = 'Asia/Taipei';

const TAB = { shops: '店家', menu: '菜單', options: '選項', members: '成員', groups: '開團', orders: '訂單' };
const HEAD = {
  shops:   ['店名', '類型', '電話', '備註', '停用'],
  menu:    ['店名', '分類', '品名', '價格', '第二容量價', '說明', '容量名稱', '優惠'],
  options: ['類型', '選項組', '選項', '加價', '複選', '必選'],
  members: ['姓名'],
  groups:  ['團ID', '店名', '截止時間', '狀態', '備註', '建立時間', '已推播', '買一送一'],
  orders:  ['團ID', '姓名', '內容', '金額', '已付金額', '更新時間', '明細'],
};
const ST_OPEN = '進行中', ST_CLOSED = '已結單';

/* ================= 網頁 API ================= */

function doGet(e) {
  try {
    const a = (e && e.parameter && e.parameter.action) || '';
    if (a === 'bootstrap') return out_(bootstrapJson_());
    return out_(JSON.stringify({ ok: true, msg: '揪團點餐 API 運作中' }));
  } catch (x) {
    return out_(JSON.stringify({ error: x.message, code: x.code || 'ERR' }));
  }
}

function doPost(e) {
  let body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (x) { return out_(JSON.stringify({ error: '資料格式不對', code: 'BAD' })); }

  // LINE webhook
  if (body.events) {
    try { line_(body); } catch (x) { console.error(x); }
    return out_('{"ok":true}');
  }

  const fn = ACTIONS[body.action];
  if (!fn) return out_(JSON.stringify({ error: '未知的動作', code: 'NOACTION' }));
  // 同一時間只讓一個請求改試算表；排不到就回 BUSY，網頁會自動重試直到存進去
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return out_(JSON.stringify({ error: '伺服器忙碌中，稍後自動重試', code: 'BUSY' }));
  try {
    const r = fn(body);
    CacheService.getScriptCache().remove('boot');
    return out_(JSON.stringify(r));
  } catch (x) {
    return out_(JSON.stringify({ error: x.message, code: x.code || 'ERR' }));
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

const ACTIONS = {
  saveOrder(p) {
    const all = readAll_(), g = all.groups.filter(x => x.id === String(p.groupId))[0];
    if (!g) throw err_('NOGROUP', '找不到這一團');
    // 截止前按下的（p.at = 按的時間），因為排隊重試晚一點才到，3 分鐘內還是收；手動結單的不收
    const t = Date.now(), at = Math.min(+p.at || t, t);
    const inTime = t < g.deadline || (at < g.deadline && t < g.deadline + 3 * 60000);
    if (!(g.status === 'open' && inTime)) throw err_('CLOSED', '這一團已截止');
    const name = String(p.name || '').trim().slice(0, 20);
    if (!name) throw err_('NONAME', '請先填名字');
    const lines = cleanLines_(p.lines);
    const sh = sheet_('orders'), v = sh.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < v.length; i++) if (String(v[i][0]) === g.id && String(v[i][1]) === name) { row = i + 1; break; }
    if (!lines.length) {
      if (row > 0) sh.deleteRow(row);
      return { ok: true };
    }
    const paid = row > 0 ? v[row - 1][4] : '';
    const dups = dupNames_(all.menu.filter(m => m.shop === g.shop));
    const data = [g.id, name, linesText_(lines, dups), sumAmt_(lines), paid, new Date(), JSON.stringify(lines)];
    if (row > 0) sh.getRange(row, 1, 1, data.length).setValues([data]);
    else sh.appendRow(data);
    return { ok: true };
  },

  setPaid(p) {
    needPin_(p);
    const sh = sheet_('orders'), v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][0]) === String(p.groupId) && String(v[i][1]) === String(p.name)) {
        sh.getRange(i + 1, 5).setValue(p.paid ? num_(v[i][3]) : '');
        return { ok: true };
      }
    }
    throw err_('NOORDER', '找不到這筆訂單');
  },

  openGroup(p) {
    needPin_(p);
    const shop = String(p.shop || '');
    if (!readShops_().some(s => s.name === shop && !s.disabled)) throw err_('NOSHOP', '找不到「' + shop + '」，或這家店暫停使用中');
    const dl = +p.deadline;
    if (!(dl > Date.now())) throw err_('BADTIME', '截止時間已經過了');
    // 前面加 g：全是數字的 ID 會被試算表當成數字改掉（例如 01234567 → 1234567），就找不到這一團
    const id = 'g' + Utilities.getUuid().replace(/-/g, '').slice(0, 7);
    // 今日買一送一：只收這家店菜單裡有的品項
    const items = readAll_().menu.filter(m => m.shop === shop), seen = {};
    const promos = (Array.isArray(p.promos) ? p.promos : []).filter(x => {
      if (!x || !items.some(m => sameItem_(m, x))) return false;
      const k = (x.cat || '') + '|' + x.item; if (seen[k]) return false; seen[k] = true; return true;
    }).slice(0, 20).map(x => ({ item: String(x.item), cat: String(x.cat || '').slice(0, 20), only: !!x.only, size: String(x.size || '').trim().slice(0, 6) }));
    const sh = sheet_('groups');
    sh.getRange(1, 1, 1, HEAD.groups.length).setValues([HEAD.groups]).setFontWeight('bold').setBackground('#DDEFE3');
    sh.appendRow([id, shop, new Date(dl), ST_OPEN, String(p.note || '').slice(0, 80), new Date(), '', promos.length ? JSON.stringify(promos) : '']);
    return { ok: true, id: id };
  },

  closeGroup(p) {
    needPin_(p);
    const r = groupRow_(p.groupId);
    r.sh.getRange(r.row, 4).setValue(ST_CLOSED);
    return { ok: true };
  },

  // 開錯團用：整團連同訂單刪掉
  deleteGroup(p) {
    needPin_(p);
    const id = String(p.groupId), r = groupRow_(id);
    r.sh.deleteRow(r.row);
    writeRows_('orders', rows_('orders').filter(x => String(x[0]).trim() && String(x[0]) !== id));
    return { ok: true };
  },

  extendGroup(p) {
    needPin_(p);
    const r = groupRow_(p.groupId);
    const dl = Math.max(Date.now(), ms_(r.values[2])) + (Math.min(240, +p.minutes || 15)) * 60000;
    r.sh.getRange(r.row, 3, 1, 2).setValues([[new Date(dl), ST_OPEN]]);
    r.sh.getRange(r.row, 7).setValue('');   // 新的截止時間到了可以再推播一次
    return { ok: true };
  },

  // 新增或修改一家店：基本資料＋整份菜單一起存（orig = 原本的店名，新增時留空）
  saveShop(p) {
    needPin_(p);
    const s = p.shop || {}, name = String(s.name || '').trim().slice(0, 30), orig = p.orig ? String(p.orig) : '';
    if (!name) throw err_('BAD', '請填店名');
    const type = s.type === '飲料' ? '飲料' : '便當';
    const items = cleanItems_(p.items, type);
    if (!items.length) throw err_('BAD', '菜單至少要有一個品項');
    const sh = sheet_('shops'), v = sh.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < v.length; i++) {
      const n = String(v[i][0]).trim();
      if (n === name && n !== orig) throw err_('DUP', '已經有「' + name + '」了');
      if (orig && n === orig) row = i + 1;
    }
    if (orig && row < 0) throw err_('NOSHOP', '找不到「' + orig + '」');
    const data = [name, type, String(s.phone || '').slice(0, 30), String(s.note || '').slice(0, 60), s.disabled ? 'Y' : ''];
    sh.getRange('C2:C').setNumberFormat('@');      // 電話存成文字，開頭的 0 才不會被試算表吃掉
    if (row > 0) sh.getRange(row, 1, 1, data.length).setValues([data]);
    else sh.appendRow(data);

    const key = orig || name;
    const menu = rows_('menu').filter(r => String(r[0]).trim() !== key && String(r[0]).trim());
    items.forEach(it => menu.push([name, it.cat, it.item, it.price, it.priceL || '', it.desc, it.sizes, it.disc]));
    writeRows_('menu', menu);

    if (orig && orig !== name) {         // 改店名：開過的團、專屬選項一起改，才對得上
      const gs = sheet_('groups'), gv = gs.getDataRange().getValues();
      for (let i = 1; i < gv.length; i++) if (String(gv[i][1]).trim() === orig) gs.getRange(i + 1, 2).setValue(name);
      const os = sheet_('options'), ov = os.getDataRange().getValues();
      for (let i = 1; i < ov.length; i++) if (String(ov[i][0]).trim() === orig) os.getRange(i + 1, 1).setValue(name);
    }
    return { ok: true };
  },

  deleteShop(p) {
    needPin_(p);
    const name = String(p.name || '');
    if (readAll_().groups.some(g => g.shop === name && g.status === 'open' && g.deadline > Date.now())) {
      throw err_('BUSY', '這家店還有進行中的團，先結單再刪除');
    }
    writeRows_('shops', rows_('shops').filter(r => String(r[0]).trim() && String(r[0]).trim() !== name));
    writeRows_('menu', rows_('menu').filter(r => String(r[0]).trim() && String(r[0]).trim() !== name));
    writeRows_('options', rows_('options').filter(r => String(r[0]).trim() && String(r[0]).trim() !== name));
    return { ok: true };
  },

  // 整套換掉一組選項。type = 便當／飲料（共用）或店名（這家店專屬；groups 空的 = 改回用共用）
  saveOptions(p) {
    needPin_(p);
    const type = String(p.type || '').trim();
    if (type !== '便當' && type !== '飲料' && !readShops_().some(s => s.name === type)) throw err_('NOSHOP', '找不到「' + type + '」');
    const rows = rows_('options').filter(r => String(r[0]).trim() && String(r[0]).trim() !== type);
    (Array.isArray(p.groups) ? p.groups : []).slice(0, 10).forEach(g => {
      const gname = String(g.name || '').trim().slice(0, 12);
      if (!gname) return;
      (g.items || []).slice(0, 30).forEach(i => {
        const label = String(i.label || '').trim().slice(0, 12);
        if (label) rows.push([type, gname, label, Math.max(0, Math.round(num_(i.add))), g.multi ? 'Y' : '', g.required ? 'Y' : '']);
      });
    });
    writeRows_('options', rows);
    return { ok: true };
  },

  saveMembers(p) {
    needPin_(p);
    const seen = {}, list = [];
    (Array.isArray(p.members) ? p.members : []).forEach(n => {
      n = String(n).trim().slice(0, 20);
      if (n && !seen[n]) { seen[n] = true; list.push([n]); }
    });
    writeRows_('members', list.slice(0, 100));
    return { ok: true };
  },
};

/* ================= 讀資料 ================= */

function bootstrapJson_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('boot');
  if (hit) return hit;
  const s = JSON.stringify(readAll_());
  try { if (s.length < 90000) cache.put('boot', s, 120); } catch (_) { }
  return s;
}

function readAll_() {
  const cutoff = Date.now() - KEEP_DAYS * 864e5;
  const groups = rows_('groups').filter(r => r[0]).map(r => ({
    id: String(r[0]), shop: String(r[1]).trim(), deadline: ms_(r[2]),
    status: String(r[3]).trim() === ST_CLOSED ? 'closed' : 'open',
    note: String(r[4] || ''), createdAt: ms_(r[5]), promos: json_(r[7], []),
  })).filter(g => g.deadline > cutoff);
  const ids = {}; groups.forEach(g => { ids[g.id] = true; });
  const orders = rows_('orders').filter(r => ids[String(r[0])]).map(r => {
    let lines = [];
    try { lines = JSON.parse(r[6] || '[]'); } catch (_) { }
    return { groupId: String(r[0]), name: String(r[1]), total: num_(r[3]), paid: num_(r[4]), updatedAt: ms_(r[5]), lines: lines };
  });
  return {
    shops: readShops_(),
    menu: rows_('menu').filter(r => r[0] && r[2]).map(r => ({
      shop: String(r[0]).trim(), cat: String(r[1] || '其他').trim(), item: String(r[2]).trim(),
      price: num_(r[3]), priceL: num_(r[4]), desc: String(r[5] || ''), sizes: String(r[6] || ''), disc: String(r[7] || ''),
    })),
    options: rows_('options').filter(r => r[0] && r[2]).map(r => ({
      type: String(r[0]).trim(), group: String(r[1] || '客製').trim(), label: String(r[2]).trim(),
      add: num_(r[3]), multi: yes_(r[4]), required: yes_(r[5]),
    })),
    members: rows_('members').map(r => String(r[0]).trim()).filter(String),
    groups: groups,
    orders: orders,
  };
}

function readShops_() {
  return rows_('shops').filter(r => r[0]).map(r => ({
    name: String(r[0]).trim(), type: String(r[1]).trim() === '飲料' ? '飲料' : '便當',
    // 舊資料的電話被試算表當成數字、吃掉開頭的 0；台灣電話都是 0 開頭，補回來
    phone: typeof r[2] === 'number' ? '0' + r[2] : String(r[2] || ''), note: String(r[3] || ''), disabled: yes_(r[4]),
  }));
}

// 整張表（標題列以下）換成 rows
function writeRows_(k, rows) {
  const sh = sheet_(k), n = HEAD[k].length, last = sh.getLastRow();
  sh.getRange(1, 1, 1, n).setValues([HEAD[k]]).setFontWeight('bold').setBackground('#DDEFE3');   // 舊版表格補上新欄位標題
  if (last > 1) sh.getRange(2, 1, last - 1, Math.max(n, sh.getLastColumn())).clearContent();
  if (!rows.length) return;
  const data = rows.map(r => { const x = r.slice(0, n); while (x.length < n) x.push(''); return x; });
  sh.getRange(2, 1, data.length, n).setValues(data);
}

function cleanItems_(items, type) {
  return (Array.isArray(items) ? items : []).filter(it => it && String(it.item || '').trim()).slice(0, 200).map(it => ({
    cat: String(it.cat || '').trim().slice(0, 20) || '其他',
    item: String(it.item).trim().slice(0, 40),
    price: Math.max(0, Math.round(num_(it.price))),
    priceL: Math.max(0, Math.round(num_(it.priceL))),
    desc: String(it.desc || '').trim().slice(0, 40),
    sizes: num_(it.priceL) ? sizes_(it.sizes, type) : '',
    disc: disc_(it.disc),
  }));
}
// 分類優惠：「-5」= 每份減 5 元、「x90」= 打 9 折；其他一律當作沒有優惠
function disc_(s) {
  s = String(s || '').trim();
  let m = s.match(/^-(\d{1,3})$/); if (m && +m[1] > 0) return '-' + (+m[1]);
  m = s.match(/^x(\d{1,2})$/); if (m && +m[1] > 0 && +m[1] < 100) return 'x' + (+m[1]);
  return '';
}
function json_(s, d) { try { return s ? JSON.parse(s) : d; } catch (_) { return d; } }
// 兩種份量的名稱，例如「M/L」「L/瓶」「小/大」
function sizes_(s, type) {
  const p = String(s || '').split('/').map(x => x.trim().slice(0, 4)).filter(String);
  return p.length === 2 && p[0] !== p[1] ? p.join('/') : (type === '飲料' ? 'M/L' : '小/大');
}
// 不同分類可以有同名品項（鍋燒類「拉麵」、泡菜類「拉麵」），用「分類＋品名」分辨
function sameItem_(a, b) { return a.item === b.item && (!a.cat || !b.cat || a.cat === b.cat); }
function dupNames_(menu) {
  const seen = {}, d = {};
  menu.forEach(m => { if (seen[m.item]) d[m.item] = true; seen[m.item] = true; });
  return d;
}
function label_(x, dups) { return x.cat && dups[x.item] ? x.cat + ' ' + x.item : x.item; }

/* ================= LINE Bot ================= */

const HELP_ = '揪團點餐小幫手\n' +
  '・統計：各團的品項數量、每個人點了什麼、付款狀態\n' +
  '・未付：還沒付錢的人\n' +
  '・點餐：點餐網頁連結';

function line_(body) {
  const token = prop_('LINE_TOKEN');
  if (!token) return;
  body.events.forEach(ev => {
    const src = ev.source || {};
    if (src.type === 'group' && src.groupId && prop_('LINE_GROUP_ID') !== src.groupId) {
      PropertiesService.getScriptProperties().setProperty('LINE_GROUP_ID', src.groupId);
    }
    if (ev.type === 'join') return reply_(ev.replyToken, [{ type: 'text', text: HELP_ }]);
    if (ev.type !== 'message' || !ev.message || ev.message.type !== 'text') return;
    const t = String(ev.message.text).trim().replace(/^[\/!！]/, '');
    let msgs = null;
    if (/^(統計|結算|點了什麼)$/.test(t)) msgs = statsMsgs_();
    else if (/^(未付|誰沒付|收錢)$/.test(t)) msgs = unpaidMsgs_();
    else if (/^(點餐|連結|網址)$/.test(t)) msgs = [linkMsg_()];
    else if (/^(說明|指令|help)$/i.test(t)) msgs = [{ type: 'text', text: HELP_ }];
    if (msgs) reply_(ev.replyToken, msgs);
  });
}

// 進行中的團＋今天截止的團
function relevantGroups_(d) {
  const today = fmt_(Date.now(), 'yyyyMMdd');
  return d.groups.filter(g => (g.status === 'open' && g.deadline > Date.now()) || fmt_(g.deadline, 'yyyyMMdd') === today)
    .sort((a, b) => a.deadline - b.deadline).slice(0, 10);
}

function statsMsgs_() {
  const d = lineData_();
  const gs = relevantGroups_(d);
  if (!gs.length) return [{ type: 'text', text: '今天還沒有開團。' }];
  return [flexOf_(gs, d)];
}

function flexOf_(gs, d) {
  const bubbles = gs.map(g => bubble_(g, d));
  if (gs.length > 1) bubbles.unshift(overviewBubble_(gs, d));   // 兩團以上：第一張是今日總覽
  return {
    type: 'flex', altText: '點餐統計：' + gs.map(g => g.shop).join('、'),
    contents: bubbles.length === 1 ? bubbles[0] : { type: 'carousel', contents: bubbles },
  };
}

function unpaidMsgs_() {
  const d = lineData_();
  const gs = relevantGroups_(d);
  if (!gs.length) return [{ type: 'text', text: '今天還沒有開團。' }];
  const parts = gs.map(g => {
    const s = summarize_(g, d);
    const un = s.os.filter(o => o.paid !== o.total);
    if (!s.os.length) return '【' + g.shop + '】還沒有人點';
    if (!un.length) return '【' + g.shop + '】大家都付清了';
    return '【' + g.shop + '】未付 ' + un.length + ' 人，共 ' + money_(s.due) + '\n' +
      un.map(o => o.name + '  ' + money_(o.total - o.paid) + (o.paid ? '（已付 ' + money_(o.paid) + '）' : '')).join('\n');
  });
  return [{ type: 'text', text: parts.join('\n\n').slice(0, 4900) }];
}

function linkMsg_() {
  const url = prop_('SITE_URL');
  return { type: 'text', text: url ? '點餐請到這裡：\n' + url : '管理員還沒設定點餐網址（SITE_URL）。' };
}

function summarize_(g, d) {
  const os = d.orders.filter(o => o.groupId === g.id && o.lines && o.lines.length);
  const menu = d.menu.filter(m => m.shop === g.shop), dups = dupNames_(menu);
  const order = menu.map(m => label_(m, dups));
  const items = {}, customs = [];
  os.forEach(o => o.lines.forEach(l => {
    // 客製點餐另外列，不混進品項合計（要先問店家能不能做）
    if (l.custom) { customs.push({ name: o.name, text: l.custom.text, qty: l.qty, fallback: fallback_(l, dups) }); return; }
    const name = label_(l, dups);       // 同名不同分類的會帶分類，例如「鍋燒類 拉麵」
    const it = items[name] || (items[name] = { item: name, qty: 0, sets: 0, vars: {} });
    it.qty += l.qty * (l.bogo ? 2 : 1);                 // 買一送一的一組是 2 杯
    if (l.bogo) it.sets += l.qty;
    // 買一送一拆成兩杯各自統計（兩杯甜度冰塊可能不同），店家照杯數做
    const cups = l.bogo ? [l.opts || [], cup2_(l)] : [l.opts || []];
    cups.forEach((opts, ci) => {
      const vk = [l.bogo ? '買一送一' : '', l.size].concat(opts).filter(String).join(' ');
      const v = it.vars[vk] || (it.vars[vk] = { label: vk, qty: 0, bogo: !!l.bogo, notes: [] });
      v.qty += l.qty;
      if (l.note && ci === 0) v.notes.push(o.name + '：' + l.note);
    });
  }));
  const idx = n => { const i = order.indexOf(n); return i < 0 ? 999 : i; };
  const list = Object.keys(items).map(k => items[k]).sort((a, b) => idx(a.item) - idx(b.item)).map(it => {
    it.vars = Object.keys(it.vars).map(k => it.vars[k]).sort((a, b) => b.qty - a.qty);
    return it;
  });
  let total = 0, paid = 0, due = 0, qty = 0;
  os.forEach(o => { total += o.total; paid += o.paid; due += Math.max(0, o.total - o.paid); });
  list.forEach(it => { qty += it.qty; });
  return { os: os, list: list, customs: customs, total: total, paid: paid, due: due, qty: qty };
}

// 今日總覽：每家店要付多少、每個人今天總共要付多少（一次收齊）
function overviewBubble_(gs, d) {
  const ink = '#16231B', muted = '#5D6D63', green = '#1E7A4C';
  const txt = (t, o) => Object.assign({ type: 'text', text: String(t || ' '), size: 'sm', color: ink, wrap: true }, o || {});
  const row = (l, r, o) => ({ type: 'box', layout: 'horizontal', spacing: 'sm', contents: [
    txt(l, Object.assign({ flex: 5 }, o && o.l)), txt(r, Object.assign({ flex: 2, align: 'end' }, o && o.r)),
  ] });
  const typeOf = g => (d.shops.filter(x => x.name === g.shop)[0] || { type: '便當' }).type;
  const count = {}; gs.forEach(g => { count[typeOf(g)] = (count[typeOf(g)] || 0) + 1; });
  const tag = g => count[typeOf(g)] > 1 ? g.shop : typeOf(g);
  const shopRows = [], people = {}, order = [];
  let total = 0, paid = 0;
  gs.forEach(g => {
    const s = summarize_(g, d);
    total += s.total; paid += s.paid;
    shopRows.push(row(typeOf(g) + '｜' + g.shop + '　' + s.qty + (typeOf(g) === '飲料' ? ' 杯' : ' 份'), money_(s.total), { r: { weight: 'bold' } }));
    s.os.forEach(o => {
      if (!people[o.name]) { people[o.name] = { name: o.name, parts: [], total: 0, paid: 0 }; order.push(o.name); }
      const p = people[o.name];
      p.parts.push(tag(g) + ' ' + money_(o.total)); p.total += o.total; p.paid += o.paid;
    });
  });
  const personRows = [];
  order.map(n => people[n]).sort((a, b) => (a.paid >= a.total) - (b.paid >= b.total)).slice(0, 35).forEach(p => {
    const ok = p.paid >= p.total;
    personRows.push(row((ok ? '✅ ' : '⬜ ') + p.name, money_(p.total), { l: { weight: 'bold' }, r: { color: ok ? green : ink, weight: 'bold' } }));
    if (p.parts.length > 1) personRows.push(txt('　' + p.parts.join(' ＋ '), { size: 'xxs', color: muted }));
  });
  const url = prop_('SITE_URL');
  const footer = [
    row('合計', money_(total), { l: { weight: 'bold' }, r: { weight: 'bold' } }),
    row('已收', money_(paid), { r: { color: green } }),
    row('未收', money_(total - paid), { r: { color: total - paid ? '#CF4128' : ink, weight: 'bold' } }),
  ];
  if (url) footer.push({ type: 'button', style: 'primary', color: green, height: 'sm', margin: 'md', action: { type: 'uri', label: '看今日總覽', uri: url.replace(/#.*$/, '') + '#today' } });
  return {
    type: 'bubble', size: 'mega',
    header: { type: 'box', layout: 'vertical', paddingAll: '16px', backgroundColor: ink, contents: [
      txt('今日總覽', { size: 'xl', weight: 'bold', color: '#FFFFFF' }),
      txt(fmtDay_(Date.now()) + ' · ' + gs.length + ' 團 · ' + order.length + ' 人', { size: 'xs', color: '#C9D3CC' }),
    ] },
    body: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [txt('每家店要付', { weight: 'bold', size: 'xs', color: muted })]
      .concat(shopRows, [{ type: 'separator', margin: 'lg' }, txt('每個人要付', { weight: 'bold', size: 'xs', color: muted, margin: 'lg' })], personRows.length ? personRows : [txt('還沒有人點', { color: muted })]) },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: footer },
  };
}

function bubble_(g, d) {
  const s = summarize_(g, d);
  const shop = d.shops.filter(x => x.name === g.shop)[0] || { type: '便當' };
  const open = g.status === 'open' && g.deadline > Date.now();
  const drink = shop.type === '飲料';
  const ink = '#16231B', muted = '#5D6D63';
  const txt = (t, o) => Object.assign({ type: 'text', text: String(t || ' '), size: 'sm', color: ink, wrap: true }, o || {});
  const row = (l, r, o) => ({
    type: 'box', layout: 'horizontal', spacing: 'sm', contents: [
      txt(l, Object.assign({ flex: 5 }, o && o.l)), txt(r, Object.assign({ flex: 2, align: 'end' }, o && o.r)),
    ],
  });
  const title = t => txt(t, { weight: 'bold', size: 'xs', color: muted, margin: 'lg' });

  const itemRows = [];
  s.list.slice(0, 25).forEach(it => {
    const simple = it.vars.length === 1 && !it.vars[0].label;
    itemRows.push(row(it.item, '×' + it.qty, { l: { weight: 'bold' }, r: { weight: 'bold' } }));
    if (!simple) it.vars.forEach(v => itemRows.push(row('　' + (v.label || '一般'), '×' + v.qty + (v.bogo ? ' 杯' : ''), { l: { color: muted, size: 'xs' }, r: { color: muted, size: 'xs' } })));
    if (it.sets) itemRows.push(txt('　買一送一共 ' + it.sets + ' 組', { size: 'xxs', color: muted }));
    it.vars.forEach(v => v.notes.forEach(n => itemRows.push(txt('　備註 ' + n, { size: 'xxs', color: muted }))));
  });
  if (s.customs.length) {
    itemRows.push({ type: 'separator', margin: 'lg' }, title('客製點餐（先問店家能不能做）'));
    s.customs.slice(0, 15).forEach(c => {
      itemRows.push(txt(c.name + '：' + c.text + (c.qty > 1 ? ' ×' + c.qty : ''), { weight: 'bold' }));
      itemRows.push(txt('　做不到改：' + c.fallback, { size: 'xs', color: muted }));
    });
  }
  const peopleRows = s.os.slice(0, 35).map(o => {
    const ok = o.paid === o.total;
    return row((ok ? '✅ ' : '⬜ ') + o.name, money_(o.total), { r: { color: ok ? '#1E7A4C' : ink } });
  });

  const footer = [
    row('合計 ' + s.qty + ' 份', money_(s.total), { l: { weight: 'bold' }, r: { weight: 'bold' } }),
    row('已收', money_(s.paid), { r: { color: '#1E7A4C' } }),
    row('未收', money_(s.due), { r: { color: s.due ? '#CF4128' : ink, weight: 'bold' } }),
  ];
  const url = prop_('SITE_URL');
  if (url) footer.push({ type: 'button', style: 'primary', color: '#1E7A4C', height: 'sm', margin: 'md', action: { type: 'uri', label: open ? '去點餐' : '看明細', uri: url } });

  return {
    type: 'bubble', size: 'mega',
    header: {
      type: 'box', layout: 'vertical', paddingAll: '16px', backgroundColor: drink ? '#F2B431' : '#1E7A4C',
      contents: [
        txt(g.shop, { size: 'xl', weight: 'bold', color: drink ? '#2B2105' : '#FFFFFF' }),
        txt(fmtDay_(g.deadline) + ' ' + fmt_(g.deadline, 'HH:mm') + ' 截止 · ' + (open ? '進行中' : '已截止') + ' · ' + s.os.length + ' 人',
          { size: 'xs', color: drink ? '#2B2105' : '#DDEFE3' }),
      ],
    },
    body: {
      type: 'box', layout: 'vertical', spacing: 'sm',
      contents: s.os.length
        ? [title('品項統計')].concat(itemRows, [{ type: 'separator', margin: 'lg' }, title('訂購人')], peopleRows)
        : [txt('還沒有人點', { color: muted })],
    },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: footer },
  };
}

function reply_(token, msgs) {
  lineApi_('reply', { replyToken: token, messages: msgs.slice(0, 5) });
}
function push_(to, msgs) {
  lineApi_('push', { to: to, messages: msgs.slice(0, 5) });
}
function lineApi_(kind, payload) {
  const r = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/' + kind, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + prop_('LINE_TOKEN') },
    payload: JSON.stringify(payload),
  });
  if (r.getResponseCode() !== 200) console.error('LINE ' + kind + ' ' + r.getResponseCode() + ' ' + r.getContentText());
}

/** 每 5 分鐘由觸發條件執行：截止的團推播一次統計（PUSH_ON_DEADLINE=true 才會推） */
function checkDeadlines() {
  if (prop_('PUSH_ON_DEADLINE') !== 'true') return;
  const to = prop_('LINE_GROUP_ID');
  if (!to || !prop_('LINE_TOKEN')) return;
  const d = lineData_(), t = Date.now();
  d.groups.forEach(g => {
    if (g.status !== 'open' || g.pushedAt || g.deadline > t || t - g.deadline > 3600e3) return;   // 只推剛截止一小時內的
    push_(to, [{ type: 'text', text: '【' + g.shop + '】截止囉，統計如下：' }, flexOf_([g], d)]);
    if (DATA_SOURCE === 'firebase') fsPatch_('groups/' + g.id, { pushedAt: t });
  });
}

/* ================= Firebase（Firestore）：網頁的資料都放在這裡，LINE Bot 從這裡讀 ================= */

// 'firebase' = 讀 Firestore（新）；'sheet' = 讀試算表（舊）
const DATA_SOURCE = 'firebase';
const FB = { project: 'jiutuan-order', key: 'AIzaSyDRvhglm1In1NWISq0vvnVjSX1jzXrnFhQ' };
const FS_BASE = 'https://firestore.googleapis.com/v1/projects/' + FB.project + '/databases/(default)/documents';

function lineData_() { return DATA_SOURCE === 'firebase' ? fsAll_(Date.now() - KEEP_DAYS * 864e5) : JSON.parse(bootstrapJson_()); }

// Firestore REST 的值 → 一般的 JS 值
function fsVal_(v) {
  if (!v) return null;
  if ('stringValue' in v) return v.stringValue;
  if ('integerValue' in v) return +v.integerValue;
  if ('doubleValue' in v) return +v.doubleValue;
  if ('booleanValue' in v) return v.booleanValue;
  if ('nullValue' in v) return null;
  if ('timestampValue' in v) return new Date(v.timestampValue).getTime();
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(fsVal_);
  if ('mapValue' in v) { const o = {}, f = v.mapValue.fields || {}; Object.keys(f).forEach(k => { o[k] = fsVal_(f[k]); }); return o; }
  return null;
}
function fsDoc_(d) { const o = fsVal_({ mapValue: { fields: d.fields || {} } }); o._id = d.name.split('/').pop(); return o; }
function fsFetch_(url, opt) {
  const r = UrlFetchApp.fetch(url + (url.indexOf('?') < 0 ? '?' : '&') + 'key=' + FB.key, Object.assign({ muteHttpExceptions: true }, opt || {}));
  if (r.getResponseCode() !== 200) throw new Error('Firestore ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
  return JSON.parse(r.getContentText());
}
function fsList_(col) {
  let out = [], token = '';
  do {
    const j = fsFetch_(FS_BASE + '/' + col + '?pageSize=300' + (token ? '&pageToken=' + encodeURIComponent(token) : ''));
    out = out.concat((j.documents || []).map(fsDoc_)); token = j.nextPageToken || '';
  } while (token);
  return out;
}
function fsQuery_(col, field, op, value) {
  const q = { structuredQuery: { from: [{ collectionId: col }], where: { fieldFilter: { field: { fieldPath: field }, op: op, value: { integerValue: String(Math.round(value)) } } } } };
  const j = fsFetch_(FS_BASE + ':runQuery', { method: 'post', contentType: 'application/json', payload: JSON.stringify(q) });
  return j.filter(x => x.document).map(x => fsDoc_(x.document));
}
function fsPatch_(path, fields) {
  const f = {}, mask = Object.keys(fields).map(k => 'updateMask.fieldPaths=' + k).join('&');
  Object.keys(fields).forEach(k => { f[k] = typeof fields[k] === 'number' ? { integerValue: String(fields[k]) } : { stringValue: String(fields[k]) }; });
  fsFetch_(FS_BASE + '/' + path + '?' + mask, { method: 'patch', contentType: 'application/json', payload: JSON.stringify({ fields: f }) });
}
// 組成跟舊版 readAll_() 一樣的形狀，LINE 的統計程式不用改；since = 0 代表全部（備份用）
function fsAll_(since) {
  const shops = fsList_('shops').sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0));
  const meta = (fsList_('meta').filter(x => x._id === 'app')[0]) || { members: [], options: [] };
  const groups = (since ? fsQuery_('groups', 'deadline', 'GREATER_THAN', since) : fsList_('groups')).map(g => ({
    id: g._id, shop: g.shop, deadline: g.deadline, status: g.status === 'closed' ? 'closed' : 'open',
    note: g.note || '', createdAt: g.createdAt || 0, promos: g.promos || [], pushedAt: g.pushedAt || 0,
  }));
  const orders = (since ? fsQuery_('orders', 'gdl', 'GREATER_THAN', since) : fsList_('orders'))
    .filter(o => o.lines && o.lines.length)
    .map(o => ({ groupId: o.groupId, name: o.name, total: o.total || 0, paid: o.paid || 0, updatedAt: o.updatedAt || 0, lines: o.lines }));
  return {
    shops: shops.map(s => ({ name: s.name, type: s.type, phone: s.phone || '', note: s.note || '', disabled: !!s.disabled })),
    menu: [].concat.apply([], shops.map(s => (s.menu || []).map(m => Object.assign({ shop: s.name }, m)))),
    options: meta.options || [], members: meta.members || [], groups: groups, orders: orders,
  };
}

/** 把 Firebase 的資料備份到試算表（寫到「備份_」開頭的分頁，不動原本的分頁）。可以手動執行，也可以每天自動跑 */
function backupToSheet() {
  const d = fsAll_(0), ss = ss_(), when = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm');
  const dups = {}; d.shops.forEach(s => { dups[s.name] = dupNames_(d.menu.filter(m => m.shop === s.name)); });
  const shopOf = id => (d.groups.filter(g => g.id === id)[0] || {}).shop || '';
  const tabs = {
    '備份_店家': [['店名', '類型', '電話', '備註', '停用']].concat(d.shops.map(s => [s.name, s.type, s.phone, s.note, s.disabled ? 'Y' : ''])),
    '備份_菜單': [['店名', '分類', '品名', '價格', '第二份量價', '說明', '份量名稱', '優惠']].concat(d.menu.map(m => [m.shop, m.cat, m.item, m.price, m.priceL || '', m.desc || '', m.sizes || '', m.disc || ''])),
    '備份_開團': [['團ID', '店名', '截止時間', '狀態', '備註']].concat(d.groups.sort((a, b) => b.deadline - a.deadline).map(g => [g.id, g.shop, new Date(g.deadline), g.status === 'closed' ? ST_CLOSED : ST_OPEN, g.note])),
    '備份_訂單': [['團ID', '店名', '截止時間', '姓名', '內容', '金額', '已付金額']].concat(d.orders.map(o => {
      const g = d.groups.filter(x => x.id === o.groupId)[0] || {};
      return [o.groupId, shopOf(o.groupId), g.deadline ? new Date(g.deadline) : '', o.name, linesText_(o.lines, dups[g.shop] || {}), o.total, o.paid || ''];
    })),
    '備份_成員': [['姓名']].concat(d.members.map(n => [n])),
  };
  Object.keys(tabs).forEach(name => {
    const rows = tabs[name];
    let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name);
    sh.clearContents();
    sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold').setBackground('#DDEFE3');
    sh.setFrozenRows(1);
    sh.getRange(1, rows[0].length + 2).setValue('最後備份：' + when);
  });
}
/** 執行一次：每天晚上 11 點自動備份到試算表 */
function installBackupTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'backupToSheet').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupToSheet').timeBased().everyDays(1).atHour(23).create();
}

/* ================= 安裝 ================= */

/** 第一次使用執行一次：建立工作表並放入範例資料 */
function setup() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);
  Object.keys(TAB).forEach(k => {
    let sh = ss.getSheetByName(TAB[k]);
    if (!sh) sh = ss.insertSheet(TAB[k]);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEAD[k].length).setValues([HEAD[k]]).setFontWeight('bold').setBackground('#DDEFE3');
      sh.setFrozenRows(1);
    }
  });
  const fill = (k, rows) => {
    const sh = ss.getSheetByName(TAB[k]);
    if (sh.getLastRow() > 1 || !rows.length) return;
    sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  };
  fill('shops', [
    ['福記便當', '便當', '02-2750-1234', '11:40 左右送到一樓', ''],
    ['青禾手搖', '飲料', '02-2750-5678', '滿 10 杯外送', ''],
  ]);
  fill('menu', [
    ['福記便當', '招牌便當', '招牌雞腿飯', 110, '', '炸雞腿、三樣配菜'],
    ['福記便當', '招牌便當', '香酥排骨飯', 95, '', ''],
    ['福記便當', '招牌便當', '古早味控肉飯', 100, '', '肥瘦各半'],
    ['福記便當', '招牌便當', '鹽烤鯖魚飯', 105, '', ''],
    ['福記便當', '招牌便當', '蔬食便當', 85, '', ''],
    ['福記便當', '湯品', '貢丸湯', 30, '', ''],
    ['青禾手搖', '原葉茶', '茉莉綠茶', 30, 35, ''],
    ['青禾手搖', '原葉茶', '四季春', 30, 35, ''],
    ['青禾手搖', '奶茶', '珍珠奶茶', 50, 60, ''],
    ['青禾手搖', '奶茶', '紅茶拿鐵', 55, 65, '鮮奶'],
    ['青禾手搖', '果茶', '百香綠茶', 50, 60, ''],
  ]);
  const opt = [];
  ['飯少', '飯多', '不要辣', '不要蔥'].forEach(l => opt.push(['便當', '客製', l, 0, 'Y', '']));
  opt.push(['便當', '加點', '加滷蛋', 15, 'Y', '']);
  ['正常糖', '少糖', '半糖', '微糖', '一分糖', '無糖'].forEach(l => opt.push(['飲料', '甜度', l, 0, '', 'Y']));
  ['正常冰', '少冰', '微冰', '去冰', '常溫', '熱'].forEach(l => opt.push(['飲料', '冰塊', l, 0, '', 'Y']));
  [['珍珠', 10], ['椰果', 10], ['仙草凍', 10], ['布丁', 15]].forEach(a => opt.push(['飲料', '加料', a[0], a[1], 'Y', '']));
  fill('options', opt);
  ss.getSheetByName(TAB.groups).getRange('C:C').setNumberFormat('yyyy/mm/dd hh:mm');
  ss.getSheetByName(TAB.shops).getRange('C2:C').setNumberFormat('@');
  ss.getSheetByName(TAB.groups).getRange('F:G').setNumberFormat('yyyy/mm/dd hh:mm');
  ss.getSheetByName(TAB.orders).getRange('F:F').setNumberFormat('yyyy/mm/dd hh:mm');
  const s1 = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  CacheService.getScriptCache().remove('boot');
}

/** 要用截止自動推播才需要執行一次 */
function installTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'checkDeadlines').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkDeadlines').timeBased().everyMinutes(5).create();
}

/* ================= 小工具 ================= */

function ss_() { return SpreadsheetApp.getActive(); }
function sheet_(k) {
  const sh = ss_().getSheetByName(TAB[k]);
  if (!sh) throw err_('SETUP', '找不到工作表「' + TAB[k] + '」，請先執行 setup()');
  return sh;
}
function rows_(k) { const v = sheet_(k).getDataRange().getValues(); v.shift(); return v; }
function out_(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON); }
function err_(code, msg) { const e = new Error(msg); e.code = code; return e; }
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function needPin_(p) { if (String(p.pin) !== PIN) throw err_('PIN', 'PIN 碼不對'); }
function num_(v) { const n = +v; return isFinite(n) ? n : 0; }
function yes_(v) { return v === true || /^(Y|YES|是|TRUE|V|1)$/i.test(String(v).trim()); }
function ms_(v) { if (v instanceof Date) return v.getTime(); const n = new Date(v).getTime(); return isFinite(n) ? n : 0; }
function fmt_(ms, f) { return Utilities.formatDate(new Date(ms), TZ, f); }
function fmtDay_(ms) { return fmt_(ms, 'M/d') + '（' + '一二三四五六日'.charAt(+fmt_(ms, 'u') - 1) + '）'; }
function money_(n) { return '$' + Math.round(n || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ','); }

function findGroup_(id) {
  const g = readAll_().groups.filter(x => x.id === String(id))[0];
  if (!g) throw err_('NOGROUP', '找不到這一團');
  return g;
}
function groupRow_(id) {
  const sh = sheet_('groups'), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) return { sh: sh, row: i + 1, values: v[i] };
  throw err_('NOGROUP', '找不到這一團');
}

function cleanLines_(lines) {
  return (Array.isArray(lines) ? lines : []).filter(l => l && l.item && +l.qty > 0).slice(0, 30).map(l => {
    const x = {
      item: String(l.item).slice(0, 40),
      cat: String(l.cat || '').slice(0, 20),
      bogo: !!l.bogo,
      size: String(l.size || '').trim().slice(0, 6),
      opts: (Array.isArray(l.opts) ? l.opts : []).map(String).slice(0, 12),
      opts2: l.bogo && Array.isArray(l.opts2) ? l.opts2.map(String).slice(0, 12) : [],
      unit: Math.max(0, Math.round(+l.unit || 0)),
      qty: Math.min(50, Math.max(1, Math.round(+l.qty))),
      note: String(l.note || '').slice(0, 60),
    };
    // 客製點餐：custom.text 是想點的（菜單上沒有），其餘欄位是做不到時改點的備案
    if (l.custom && String(l.custom.text || '').trim()) {
      x.custom = { text: String(l.custom.text).trim().slice(0, 60), price: Math.max(0, Math.round(num_(l.custom.price))) };
      x.fbUnit = Math.max(0, Math.round(num_(l.fbUnit)));
    }
    return x;
  });
}
function sumAmt_(lines) { return lines.reduce((s, l) => s + l.unit * l.qty, 0); }
// 買一送一第 2 杯的選項；沒有另外選就跟第 1 杯一樣
function cup2_(l) { return l.opts2 && l.opts2.length ? l.opts2 : (l.opts || []); }
function variant_(l) {
  const v = [l.bogo ? '買一送一' : '', l.size].concat(l.opts || []).filter(String).join(' ');
  const c2 = cup2_(l).join(' ');
  return l.bogo && c2 !== (l.opts || []).join(' ') ? v + '／第2杯 ' + c2 : v;
}
function fallback_(l, dups) { return label_(l, dups || {}) + (variant_(l) ? ' ' + variant_(l) : ''); }
function linesText_(lines, dups) {
  return lines.map(l => (l.custom ? '客製：' + l.custom.text + '（做不到改：' + fallback_(l, dups) + '）' : fallback_(l, dups)) +
    ' ×' + l.qty + (l.bogo ? ' 組' : '') + (l.note ? '（' + l.note + '）' : '')).join('；');
}
