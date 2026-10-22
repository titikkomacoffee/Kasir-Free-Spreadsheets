// KasirKu - Backend (Google Apps Script)
// GANTI teks di bawah dengan ID Spreadsheet Anda
const SS_ID='1xYhyxbl3_mKt9XKKEWUg0OAAOYZjCpVgsEQGKj88vAg';
const T={Users:['id','nama','user','pass','role','aktif','cabang'],
Products:['id','sku','nama','kategori','harga','modal','stok','foto'],
Trx:['id','tgl','kasir','cabang','metode','diskon','total','bayar','kembali','items','bukti','status'],
Debts:['id','tgl','pelanggan','telp','total','dibayar','jatuhtempo','status','trx'],
Shifts:['id','kasir','cabang','buka','tutup','kasAwal','kasAkhir','status'],
Absen:['id','user','tipe','waktu','lat','lng','foto'],
Branches:['id','nama','alamat'],Settings:['id','value']};
const W={Kasir:['Trx','Debts','Shifts','Absen','Products'],Manager:['Trx','Debts','Shifts','Absen','Products'],Admin:Object.keys(T)};
const H=s=>Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s));

function sh(n){
  const ss=SpreadsheetApp.openById(SS_ID);let s=ss.getSheetByName(n);
  if(!s){s=ss.insertSheet(n);s.getRange(1,1,500,T[n].length).setNumberFormat('@');s.appendRow(T[n]);s.setFrozenRows(1);
    if(n=='Users')s.appendRow(['u1','Admin','admin',H('admin123'),'Admin','Aktif','Cabang Utama']);
    if(n=='Branches')s.appendRow(['b1','Cabang Utama','']);}
  return s;
}
function rows(n){
  const v=sh(n).getDataRange().getValues(),h=v.shift();
  return v.map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i]);return o;});
}
const out=o=>ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
function doGet(){return out({ok:1,info:'KasirKu API aktif'});}

function doPost(e){
  const lock=LockService.getScriptLock();lock.waitLock(25000);
  try{
    const q=JSON.parse(e.postData.contents),us=rows('Users');
    const u=us.find(x=>x.user==q.u&&x.pass==H(q.p)&&x.aktif=='Aktif');
    if(!u)return out({ok:0,err:'Username/password salah atau akun nonaktif'});
    if(q.a=='login')return out({ok:1,user:{id:u.id,nama:u.nama,user:u.user,role:u.role,cabang:u.cabang}});
    if(!T[q.t])return out({ok:0,err:'Tabel tidak valid'});
    if(q.a=='list'){
      let d=rows(q.t);if(q.t=='Users')d=d.map(x=>{delete x.pass;return x;});
      return out({ok:1,d});
    }
    if(W[u.role].indexOf(q.t)<0)return out({ok:0,err:'Tidak punya akses'});
    const s=sh(q.t),v=s.getDataRange().getValues(),h=v[0];
    if(q.a=='save'){
      const r=q.d;
      if(q.t=='Users'){const old=us.find(x=>x.id==r.id);r.pass=r.pass?H(r.pass):(old?old.pass:H('123456'));}
      const row=h.map(k=>r[k]===undefined?'':r[k]);
      const i=v.findIndex((x,j)=>j>0&&x[0]==r.id);
      if(i>0)s.getRange(i+1,1,1,h.length).setValues([row]);else s.appendRow(row);
      return out({ok:1});
    }
    if(q.a=='del'){
      const i=v.findIndex((x,j)=>j>0&&x[0]==q.d.id);
      if(i>0)s.deleteRow(i+1);return out({ok:1});
    }
    return out({ok:0,err:'Aksi tidak dikenal'});
  }catch(err){return out({ok:0,err:String(err)});}
  finally{lock.releaseLock();}
}

// Jalankan fungsi ini SEKALI dari editor: memberi izin akses & membuat semua sheet
function setup(){Object.keys(T).forEach(sh);Logger.log('Selesai: semua sheet dibuat');}
