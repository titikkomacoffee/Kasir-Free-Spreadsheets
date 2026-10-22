// KasirKu - Backend (Google Apps Script)
const SS_ID = '1xYhyxbl3_mKt9XKKEWUg0OAAOYZjCpVgsEQGKj88vAg';

const T = {
  Users: ['id','nama','user','pass','role','aktif','cabang'],
  Products: ['id','sku','nama','kategori','tipe','satuan','harga','modal','foto'],
  Stok: ['id','cabang','pid','qty'],
  Transfers: ['id','tgl','dari','ke','pid','nama','tipe','qty','status','pengirim','penerima','tglTerima'],
  Gaji: ['id','user','nama','periode','pokok','kasbon','setoran','absen','fee','bersih','tgl'],
  Trx: ['id','tgl','kasir','cabang','metode','diskon','total','bayar','kembali','items','bukti','status'],
  Debts: ['id','tgl','pelanggan','telp','total','dibayar','jatuhtempo','status','trx'],
  Shifts: ['id','kasir','cabang','buka','tutup','kasAwal','kasAkhir','status'],
  Absen: ['id','user','tipe','waktu','lat','lng','foto'],
  Branches: ['id','nama','alamat'],
  Settings: ['id','value']
};
let SS = SS_ID; // spreadsheet usaha yang sedang aktif
const BASE = ['Trx','Debts','Shifts','Absen','Stok','Transfers'];
const W = { Kasir: BASE, Manager: BASE.concat(['Products']), Admin: Object.keys(T) };

// Fungsi berakhiran _ bersifat privat (tidak muncul di menu Jalankan)
function hash_(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s)));
}

function sheet_(n) {
  const ss = SpreadsheetApp.openById(SS);
  let s = ss.getSheetByName(n);
  if (!s) {
    s = ss.insertSheet(n);
    s.getRange(1, 1, 500, T[n].length).setNumberFormat('@');
    s.appendRow(T[n]);
    s.setFrozenRows(1);
    if (n == 'Users') s.appendRow(['u1','Admin','admin',hash_('admin123'),'Admin','Aktif','Cabang Utama']);
    if (n == 'Branches') s.appendRow(['b1','Cabang Utama','']);
  }
  const hd = s.getRange(1, 1, 1, T[n].length);
  if (hd.getValues()[0].join() != T[n].join()) hd.setValues([T[n]]);
  return s;
}

function rows_(n) {
  const v = sheet_(n).getDataRange().getValues();
  const h = v.shift();
  return v.map(function (r) {
    const o = {};
    h.forEach(function (k, i) { o[k] = r[i]; });
    return o;
  });
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return out_({ ok: 1, info: 'KasirKu API aktif' });
}

// ---- Direktori akun -> spreadsheet usaha (multi usaha) ----
function dir_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let s = ss.getSheetByName('Direktori');
  if (!s) {
    s = ss.insertSheet('Direktori');
    s.getRange(1, 1, 1000, 3).setNumberFormat('@');
    s.appendRow(['user', 'ssid', 'usaha']);
    s.setFrozenRows(1);
  }
  return s;
}
function tenant_(user) {
  const v = dir_().getDataRange().getValues();
  const r = v.find(function (x, j) { return j > 0 && x[0] == user; });
  return r ? r[1] : SS_ID;
}
function taken_(user) {
  const v = dir_().getDataRange().getValues();
  if (v.some(function (x, j) { return j > 0 && x[0] == user; })) return true;
  const old = SS; SS = SS_ID;
  const t = rows_('Users').some(function (x) { return x.user == user; });
  SS = old;
  return t;
}
function dirDel_(user) {
  const s = dir_(), v = s.getDataRange().getValues();
  const i = v.findIndex(function (x, j) { return j > 0 && x[0] == user; });
  if (i > 0) s.deleteRow(i + 1);
}
function register_(q) {
  const user = String(q.user || '').trim(), usaha = String(q.usaha || '').trim();
  if (!user || !usaha || String(q.pass || '').length < 6) return out_({ ok: 0, err: 'Lengkapi data (password minimal 6 karakter)' });
  if (!/^[A-Za-z0-9_.]+$/.test(user)) return out_({ ok: 0, err: 'Username hanya boleh huruf, angka, titik, dan garis bawah' });
  if (taken_(user)) return out_({ ok: 0, err: 'Username sudah dipakai' });
  const ss = SpreadsheetApp.create('KasirKu - ' + usaha);
  SS = ss.getId();
  Object.keys(T).forEach(function (n) { sheet_(n); });
  sheet_('Users').getRange(2, 1, 1, 7).setValues([['u1', q.nama || user, user, hash_(q.pass), 'Admin', 'Aktif', 'Pusat']]);
  sheet_('Branches').getRange(2, 1, 1, 3).setValues([['b1', 'Pusat', '']]);
  sheet_('Settings').appendRow(['toko', usaha]);
  ss.getSheets().forEach(function (sh) { if (!T[sh.getName()] && ss.getSheets().length > 1) ss.deleteSheet(sh); });
  dir_().appendRow([user, SS, usaha]);
  return out_({ ok: 1 });
}

function put_(u, us, t, r) {
  const s = sheet_(t), v = s.getDataRange().getValues(), h = v[0];
  const i = v.findIndex(function (x, j) { return j > 0 && x[0] == r.id; });
  if (t == 'Trx' && i > 0 && u.role == 'Kasir') throw new Error('Hanya Admin/Manager yang boleh mengubah atau membatalkan transaksi');
  if (t == 'Users') {
    const old = us.find(function (x) { return x.id == r.id; });
    r.pass = r.pass ? hash_(r.pass) : (old ? old.pass : hash_('123456'));
  }
  if (t == 'Users' && i <= 0 && taken_(r.user)) throw new Error('Username sudah dipakai');
  const row = h.map(function (k) { return r[k] === undefined ? '' : r[k]; });
  if (i > 0) s.getRange(i + 1, 1, 1, h.length).setValues([row]); else s.appendRow(row);
  if (t == 'Users' && i <= 0) dir_().appendRow([r.user, SS, '']);
}

function inc_(c) {
  const id = c.cabang + '|' + c.pid, s = sheet_('Stok'), v = s.getDataRange().getValues();
  const i = v.findIndex(function (x, j) { return j > 0 && x[0] == id; });
  if (i > 0) s.getRange(i + 1, 4).setValue(String(Number(v[i][3] || 0) + Number(c.n)));
  else s.appendRow([id, c.cabang, c.pid, String(Number(c.n))]);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const q = JSON.parse(e.postData.contents);
    if (q.a == 'register') return register_(q);
    SS = tenant_(q.u);
    const us = rows_('Users');
    const u = us.find(function (x) { return x.user == q.u && x.pass == hash_(q.p) && x.aktif == 'Aktif'; });
    if (!u) return out_({ ok: 0, err: 'Username/password salah atau akun nonaktif' });
    if (q.a == 'login') return out_({ ok: 1, user: { id: u.id, nama: u.nama, user: u.user, role: u.role, cabang: u.cabang } });
    if (q.a == 'list') {
      if (!T[q.t]) return out_({ ok: 0, err: 'Tabel tidak valid' });
      let d = rows_(q.t);
      if (q.t == 'Users') d = d.map(function (x) { delete x.pass; return x; });
      if (q.t == 'Gaji' && u.role != 'Admin') d = d.filter(function (x) { return x.user == u.user; });
      return out_({ ok: 1, d: d });
    }
    const items = q.a == 'batch' ? q.d : [{ t: q.t, d: q.d, del: q.a == 'del' }];
    items.forEach(function (it) { if (!T[it.t] || W[u.role].indexOf(it.t) < 0) throw new Error('Tidak punya akses: ' + it.t); });
    items.forEach(function (it) {
      if (it.inc) inc_(it.inc);
      else if (it.del) {
        const s = sheet_(it.t), v = s.getDataRange().getValues();
        const i = v.findIndex(function (x, j) { return j > 0 && x[0] == it.d.id; });
        if (i > 0) { if (it.t == 'Users') dirDel_(v[i][2]); s.deleteRow(i + 1); }
      } else put_(u, us, it.t, it.d);
    });
    return out_({ ok: 1 });
  } catch (err) {
    return out_({ ok: 0, err: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// Jalankan SEKALI dari editor: memberi izin akses & membuat semua sheet
function setup() {
  dir_();
  Object.keys(T).forEach(function (n) { sheet_(n); });
  Logger.log('Selesai: semua sheet berhasil dibuat');
}
