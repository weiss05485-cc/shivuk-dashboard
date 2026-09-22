/**
 * שיווק — ניהול משלוחים | Backend (Node.js, zero-dependency)
 * מחליף את Google Apps Script. אחסון: data.json. EZcount: ez.env.
 * רץ מאחורי nginx על 127.0.0.1:PORT (ברירת מחדל 3510).
 *
 * קבצים על השרת (ב-/opt/shivuk):
 *   server.js, data.json, ez.env
 * ez.env (KEY=VALUE בכל שורה):
 *   EZ_APIKEY=...
 *   EZ_EMAIL=weiss05485@gmail.com
 *   EZ_DEMO=1        # 1=דמו, 0/ריק=אמיתי
 */
const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3510;
const DIR = __dirname;
const DATA_FILE = path.join(DIR, 'data.json');
const ENV_FILE = path.join(DIR, 'ez.env');
const BACKUP_DIR = path.join(DIR, 'backups');

const SHEETS = ['משלוחים', 'מחירון', 'תשלומים', 'הוצאות', 'ספק_GB', 'ספק_טירול'];

// ---------- data store ----------
function loadDB() {
  try { return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); }
  catch (e) { return {}; }
}
function saveDB(db) {
  const tmp = DATA_FILE + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(db, null, 1));
  fs.renameSync(tmp, DATA_FILE);
  // גיבוי מדי פעם
  try {
    if (!fs.existsSync(BACKUP_DIR)) fs.mkdirSync(BACKUP_DIR);
    const stamp = new Date().toISOString().slice(0, 13).replace(/[:T]/g, '-');
    fs.writeFileSync(path.join(BACKUP_DIR, 'data-' + stamp + '.json'), JSON.stringify(db));
    const files = fs.readdirSync(BACKUP_DIR).filter(f => f.startsWith('data-')).sort();
    while (files.length > 60) fs.unlinkSync(path.join(BACKUP_DIR, files.shift()));
  } catch (e) {}
}
function stripApos(v) { return (typeof v === 'string' && v[0] === "'") ? v.slice(1) : v; }
function cleanRow(r) { const o = {}; for (const k in r) o[k] = stripApos(r[k]); return o; }

// ---------- ez.env ----------
function ezEnv() {
  const env = {};
  try {
    fs.readFileSync(ENV_FILE, 'utf8').split('\n').forEach(line => {
      const i = line.indexOf('='); if (i < 0) return;
      env[line.slice(0, i).trim()] = line.slice(i + 1).trim();
    });
  } catch (e) {}
  return env;
}

// ---------- EZcount ----------
function ezDoc(p) {
  return new Promise(resolve => {
    const env = ezEnv();
    if (!env.EZ_APIKEY || !env.EZ_EMAIL) return resolve({ ok: false, error: 'חסרים פרטי EZcount ב-ez.env' });
    const demo = env.EZ_DEMO === '1';
    const amount = Number(p.amount) || 0;
    const body = {
      api_key: env.EZ_APIKEY,
      developer_email: env.EZ_EMAIL,
      type: Number(p.type) || 320,
      customer_name: p.customer_name || 'לקוח',
      item: [{ details: p.details || 'תשלום עבור סחורה', price: amount, amount: 1, vat_type: 'INC' }],
      payment: [{ payment_type: Number(p.payment_type) || 1, payment_sum: amount }],
      price_total: amount,
      email_to_client: false
    };
    if (p.customer_email) body.customer_email = p.customer_email;
    if (p.customer_taxid) body.customer_business_number = p.customer_taxid;
    if (Number(p.type) === 400) delete body.item;
    const payload = JSON.stringify(body);
    const host = demo ? 'demo.ezcount.co.il' : 'api.ezcount.co.il';
    const req = https.request({ host, path: '/api/createDoc', method: 'POST',
      headers: { 'content-type': 'application/json', 'content-length': Buffer.byteLength(payload) } },
      res => {
        let d = ''; res.on('data', c => d += c);
        res.on('end', () => {
          let j; try { j = JSON.parse(d); } catch (e) { return resolve({ ok: false, error: 'EZcount: ' + d.slice(0, 300) }); }
          if (!j.success) return resolve({ ok: false, error: (j.errMsg || j.error || 'EZcount error'), raw: j });
          resolve({ ok: true, demo, doc_number: j.doc_number || j.docNumber || '',
            doc_url: j.doc_url || j.pdf_link || j.pdf_link_original || j.doc_url_for_email || '', raw: j });
        });
      });
    req.on('error', e => resolve({ ok: false, error: String(e) }));
    req.write(payload); req.end();
  });
}

// ---------- request handling ----------
function parseBody(req) {
  return new Promise(resolve => {
    let b = ''; req.on('data', c => b += c);
    req.on('end', () => {
      const ct = req.headers['content-type'] || '';
      if (ct.indexOf('application/json') >= 0) { try { return resolve(JSON.parse(b)); } catch (e) { return resolve({}); } }
      const o = {}; new URLSearchParams(b).forEach((v, k) => o[k] = v); resolve(o);
    });
  });
}
async function handle(p) {
  const action = p.action || '';
  const db = loadDB();
  if (action === 'ping') return { ok: true, app: 'shivuk', host: 'vultr' };
  if (action === 'getTable') return { ok: true, rows: db[p.sheet] || [] };
  if (action === 'ezDoc') return await ezDoc(p);
  if (action === 'saveRow') {
    const rec = cleanRow(JSON.parse(p.data));
    const arr = db[p.sheet] || (db[p.sheet] = []);
    const i = arr.findIndex(r => String(r.id) === String(rec.id));
    if (i >= 0) arr[i] = rec; else arr.push(rec);
    saveDB(db); return { ok: true };
  }
  if (action === 'saveRows') {
    const rows = JSON.parse(p.data).map(cleanRow);
    if (p.clear === '1' || !db[p.sheet]) db[p.sheet] = [];
    db[p.sheet] = db[p.sheet].concat(rows);
    saveDB(db); return { ok: true, count: rows.length };
  }
  if (action === 'deleteRow') {
    const arr = db[p.sheet] || [];
    db[p.sheet] = arr.filter(r => String(r.id) !== String(p.id));
    saveDB(db); return { ok: true };
  }
  return { ok: false, error: 'unknown action' };
}

http.createServer(async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'content-type');
  if (req.method === 'OPTIONS') { res.writeHead(204); return res.end(); }
  try {
    const url = new URL(req.url, 'http://x');
    let p = {};
    url.searchParams.forEach((v, k) => p[k] = v);
    if (req.method === 'POST') { const b = await parseBody(req); p = Object.assign(p, b); }
    const out = await handle(p);
    res.setHeader('content-type', 'application/json; charset=utf-8');
    res.end(JSON.stringify(out));
  } catch (e) {
    res.setHeader('content-type', 'application/json; charset=utf-8');
    res.end(JSON.stringify({ ok: false, error: String(e) }));
  }
}).listen(PORT, '127.0.0.1', () => console.log('shivuk server on ' + PORT));
