// ============================================================
//  SISTEM JADWAL SEKOLAH — Google Apps Script
//  Backend API + Admin Panel
//  Setup: Tools > Script Properties > tambah ADMIN_PASSWORD
// ============================================================

// ── KONFIGURASI ──────────────────────────────────────────────
const CONFIG = {
  SHEET_GURU:       'DataGuru',
  SHEET_JADWAL:     'Jadwal',
  SHEET_MAPEL:      'DataMapel',
  SHEET_LOG:        'Log',
  SHEET_WALI:       'WaliKelas',
  SHEET_PIKET:      'PiketGuru',
  SHEET_PENGUMUMAN: 'Pengumuman',
  SHEET_KALENDER:   'KalenderAkademik',
  SHEET_KARYAWAN:   'DataKaryawan',
};

const HARI_ORDER = ['SENIN','SELASA','RABU','KAMIS','JUMAT'];
const JAM_LIST   = [1,2,3,4,5,6,7,8,9,10];

// Jam pelajaran Senin–Kamis (10 JP)
const JAM_LABEL_REGULER = {
  1:'07:15–08:00', 2:'08:00–08:45', 3:'08:45–09:30', 4:'09:30–10:15',
  5:'10:30–11:15', 6:'11:15–12:00', 7:'12:35–13:15', 8:'13:15–13:55',
  9:'13:55–14:35',10:'14:35–15:15'
};
// Jam pelajaran khusus Jumat (8 JP)
const JAM_LABEL_JUMAT = {
  1:'07:15–07:55', 2:'07:55–08:35', 3:'08:35–09:15', 4:'09:15–09:55',
  5:'10:10–10:50', 6:'10:50–11:30', 7:'12:35–13:15', 8:'13:15–13:55'
};
const LITERASI_LABEL = { REGULER:'07:00–07:15', JUMAT:'07:00–07:15' };

function getJamLabel(hari) {
  return hari === 'JUMAT' ? JAM_LABEL_JUMAT : JAM_LABEL_REGULER;
}
// Untuk kompatibilitas kode lama yang masih memakai JAM_LABEL generik
const JAM_LABEL = JAM_LABEL_REGULER;

// ── ENTRY POINTS ─────────────────────────────────────────────

function doGet(e) {
  const page = (e.parameter.page || 'public');
  if (page === 'admin') {
    if (!checkAuth(e.parameter.token)) {
      return HtmlService.createHtmlOutput(loginPage())
        .setTitle('Login Admin')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    return HtmlService.createHtmlOutput(adminPage())
      .setTitle('Admin Panel — Jadwal KBM SMANSABA')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // Public end-user page
  return HtmlService.createHtmlOutput(publicPage())
    .setTitle('Aplikasi Jadwal KBM | Demo Sekolah')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    // Auth check — getPublicData & login bebas akses (tanpa token)
    const publicActions = ['login', 'getPublicData'];
    if (!publicActions.includes(action) && !checkAuthToken(data.token)) {
      return jsonResponse({ ok: false, msg: 'Sesi berakhir. Silakan login ulang.' });
    }

    switch(action) {
      case 'login':         return handleLogin(data);
      case 'getData':       return jsonResponse(getAllData());
      case 'saveGuru':      return handleSaveGuru(data);
      case 'deleteGuru':    return handleDeleteGuru(data);
      case 'saveJadwal':    return handleSaveJadwal(data);
      case 'clearJadwal':   return handleClearJadwal(data);
      case 'saveMapel':     return handleSaveMapel(data);
      case 'saveWali':      return handleSaveWali(data);
      case 'savePiket':     return handleSavePiket(data);
      case 'saveKaryawan':   return handleSaveKaryawan(data);
      case 'deleteKaryawan': return handleDeleteKaryawan(data);
      case 'savePengumuman':   return handleSavePengumuman(data);
      case 'deletePengumuman': return handleDeletePengumuman(data);
      case 'getAllPengumuman':  return jsonResponse({ ok:true, list: getAllPengumumanAdmin() });
      case 'saveKalender':     return handleSaveKalender(data);
      case 'deleteKalender':   return handleDeleteKalender(data);
      case 'getAllKalender':    return jsonResponse({ ok:true, list: getKalenderData() });
      case 'deleteMapel':   return handleDeleteMapel(data);
      case 'getPublicData': return jsonResponse(getPublicData());
      default:              return jsonResponse({ ok: false, msg: 'Aksi tidak dikenal' });
    }
  } catch(err) {
    return jsonResponse({ ok: false, msg: err.toString() });
  }
}

// ── AUTH ─────────────────────────────────────────────────────

function getAdminPassword() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || 'admin123';
}

function makeToken() {
  // Tidak pakai Session.getActiveUser() karena tidak reliable untuk akses publik "Anyone"
  const secret = getAdminPassword() + '|jadwal-sekolah-salt|' + new Date().toDateString();
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('').slice(0, 32);
}

function checkAuth(token) {
  if (!token) return false;
  return token === makeToken();
}

function checkAuthToken(token) {
  return checkAuth(token);
}

function handleLogin(data) {
  if (data.password === getAdminPassword()) {
    const token = makeToken();
    writeLog('LOGIN', 'Admin login berhasil');
    return jsonResponse({ ok: true, token: token });
  }
  writeLog('LOGIN_FAIL', 'Password salah');
  return jsonResponse({ ok: false, msg: 'Password salah' });
}

// ── DATA ACCESS ───────────────────────────────────────────────

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(sheet, name);
  }
  return sheet;
}

function initSheet(sheet, name) {
  if (name === CONFIG.SHEET_GURU) {
    sheet.appendRow(['kode','nama','mapel']);
    sheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#1a7a54').setFontColor('#ffffff');
  } else if (name === CONFIG.SHEET_MAPEL) {
    sheet.appendRow(['nama','kode','kelompok','warna']);
    sheet.getRange(1,1,1,4).setFontWeight('bold').setBackground('#1a7a54').setFontColor('#ffffff');
  } else if (name === CONFIG.SHEET_JADWAL) {
    // Header: hari | jam | KELAS X-1 | KELAS X-2 | ...
    const headers = ['hari','jam'].concat(getKelasList());
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1a7a54').setFontColor('#ffffff');
  } else if (name === CONFIG.SHEET_LOG) {
    sheet.appendRow(['waktu','aksi','keterangan']);
    sheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#333').setFontColor('#ffffff');
  } else if (name === CONFIG.SHEET_WALI) {
    // Header: kelas | kode_guru   (1 kelas = 1 wali kelas)
    sheet.appendRow(['kelas','kode_guru']);
    sheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1a7a54').setFontColor('#ffffff');
    getKelasList().forEach(k => sheet.appendRow([k, '']));
    sheet.setFrozenRows(1);
  } else if (name === CONFIG.SHEET_PIKET) {
    // Struktur baru: hari | data_json  (JSON array of {kode, pin, ket})
    sheet.appendRow(['hari','data_json']);
    sheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1a7a54').setFontColor('#ffffff');
    HARI_ORDER.forEach(h => sheet.appendRow([h, '[]']));
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,80);
    sheet.setColumnWidth(2,600);
  } else if (name === CONFIG.SHEET_PENGUMUMAN) {
    const hdrs = ['id','judul','isi','tipe','aktif','tanggal_mulai','tanggal_selesai','dibuat'];
    sheet.appendRow(hdrs);
    sheet.getRange(1,1,1,hdrs.length).setFontWeight('bold').setBackground('#1a4078').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,80); sheet.setColumnWidth(2,200); sheet.setColumnWidth(3,320);
  } else if (name === CONFIG.SHEET_KALENDER) {
    // id | judul | keterangan | kategori | tgl_mulai | tgl_selesai | warna | dibuat
    const hdrs = ['id','judul','keterangan','kategori','tgl_mulai','tgl_selesai','warna','dibuat'];
    sheet.appendRow(hdrs);
    sheet.getRange(1,1,1,hdrs.length).setFontWeight('bold').setBackground('#1a4078').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(2,220); sheet.setColumnWidth(3,300);
    sheet.setColumnWidth(5,110); sheet.setColumnWidth(6,110);
  } else if (name === CONFIG.SHEET_KARYAWAN) {
    // id | nama | jabatan
    const hdrs = ['id','nama','jabatan'];
    sheet.appendRow(hdrs);
    sheet.getRange(1,1,1,hdrs.length).setFontWeight('bold').setBackground('#1a4078').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,120); sheet.setColumnWidth(2,220); sheet.setColumnWidth(3,200);
  }
}

// ── KARYAWAN ──────────────────────────────────────────────────
// Karyawan hanya untuk piket — tidak masuk jadwal KBM sama sekali

function getKaryawanData() {
  const sheet = getSheet(CONFIG.SHEET_KARYAWAN);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  return rows.slice(1)
    .filter(r => r[0] && r[1])
    .map(r => ({
      id:      String(r[0]),
      nama:    String(r[1]),
      jabatan: String(r[2] || ''),
    }));
}

function handleSaveKaryawan(data) {
  const { id, nama, jabatan } = data;
  if (!nama || !nama.trim()) return jsonResponse({ ok:false, msg:'Nama karyawan wajib diisi' });
  const sheet = getSheet(CONFIG.SHEET_KARYAWAN);
  const rows  = sheet.getDataRange().getValues();
  if (id) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.getRange(i+1, 2, 1, 2).setValues([[nama.trim(), jabatan||'']]);
        writeLog('EDIT_KARYAWAN', 'Update: ' + nama);
        return jsonResponse({ ok:true, msg:'Karyawan berhasil diperbarui' });
      }
    }
  }
  const newId = 'k_' + new Date().getTime();
  sheet.appendRow([newId, nama.trim(), jabatan||'']);
  writeLog('ADD_KARYAWAN', 'Tambah: ' + nama);
  return jsonResponse({ ok:true, msg:'Karyawan berhasil ditambahkan', id: newId });
}

function handleDeleteKaryawan(data) {
  const sheet = getSheet(CONFIG.SHEET_KARYAWAN);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i+1);
      writeLog('DELETE_KARYAWAN', 'Hapus id: ' + data.id);
      return jsonResponse({ ok:true, msg:'Karyawan berhasil dihapus' });
    }
  }
  return jsonResponse({ ok:false, msg:'Karyawan tidak ditemukan' });
}

// Resolver: ambil info nama dari guru (kode angka) atau karyawan (kode "k_xxx")
// Dipakai oleh piket agar bisa menampilkan guru dan karyawan sekaligus
function getPersonInfo(kode, guruData, karyawanData) {
  if (!kode) return null;
  const k = String(kode);
  // Karyawan: id dimulai dengan "k_"
  if (k.startsWith('k_')) {
    const kar = karyawanData.find(x => x.id === k);
    if (kar) return { nama: kar.nama, sub: kar.jabatan || 'Karyawan', isKaryawan: true };
    return { nama: 'Karyawan #'+k, sub: 'Karyawan', isKaryawan: true };
  }
  // Guru: kode angka
  const g = guruData[k];
  if (g) return { nama: g.nama, sub: g.mapel || '', isKaryawan: false };
  return { nama: 'Kode '+k, sub: '', isKaryawan: false };
}

// ── KALENDER AKADEMIK ─────────────────────────────────────────

function getKalenderData() {
  const sheet = getSheet(CONFIG.SHEET_KALENDER);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  return rows.slice(1).filter(r => r[0] && r[1]).map(r => ({
    id:          String(r[0]),
    judul:       String(r[1]||''),
    keterangan:  String(r[2]||''),
    kategori:    String(r[3]||'lainnya'),
    tglMulai:    r[4] ? Utilities.formatDate(new Date(r[4]), 'Asia/Jakarta', 'yyyy-MM-dd') : '',
    tglSelesai:  r[5] ? Utilities.formatDate(new Date(r[5]), 'Asia/Jakarta', 'yyyy-MM-dd') : '',
    warna:       String(r[6]||''),
    dibuat:      r[7] ? Utilities.formatDate(new Date(r[7]), 'Asia/Jakarta', 'dd MMM yyyy HH:mm') : '',
  })).sort((a,b) => a.tglMulai.localeCompare(b.tglMulai));
}

function handleSaveKalender(data) {
  const sheet = getSheet(CONFIG.SHEET_KALENDER);
  const { id, judul, keterangan, kategori, tglMulai, tglSelesai, warna } = data;
  if (!judul || !tglMulai) return jsonResponse({ ok:false, msg:'Judul dan tanggal mulai wajib diisi' });
  const rows = sheet.getDataRange().getValues();
  if (id) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.getRange(i+1,2,1,6).setValues([[judul, keterangan||'', kategori||'lainnya', tglMulai, tglSelesai||'', warna||'']]);
        writeLog('EDIT_KALENDER', 'Update: '+judul);
        return jsonResponse({ ok:true, msg:'Event diperbarui' });
      }
    }
  }
  const newId = 'k'+new Date().getTime();
  sheet.appendRow([newId, judul, keterangan||'', kategori||'lainnya', tglMulai, tglSelesai||'', warna||'', new Date()]);
  writeLog('ADD_KALENDER', 'Tambah: '+judul);
  return jsonResponse({ ok:true, msg:'Event ditambahkan', id:newId });
}

function handleDeleteKalender(data) {
  const sheet = getSheet(CONFIG.SHEET_KALENDER);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i+1);
      writeLog('DELETE_KALENDER', 'Hapus id: '+data.id);
      return jsonResponse({ ok:true, msg:'Event dihapus' });
    }
  }
  return jsonResponse({ ok:false, msg:'Event tidak ditemukan' });
}

// ── PENGUMUMAN ────────────────────────────────────────────────

function getPengumumanData() {
  // Hanya kembalikan yang aktif dan dalam rentang tanggal berlaku (untuk publik)
  const sheet = getSheet(CONFIG.SHEET_PENGUMUMAN);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const today = new Date(); today.setHours(0,0,0,0);
  const list  = [];
  for (let i = 1; i < rows.length; i++) {
    const [id, judul, isi, tipe, aktif, tglMulai, tglSelesai, dibuat] = rows[i];
    if (!id || !judul) continue;
    if (String(aktif).toLowerCase() !== 'ya') continue;
    const mulai   = tglMulai   ? new Date(tglMulai)  : null;
    const selesai = tglSelesai ? new Date(tglSelesai) : null;
    if (mulai)   { mulai.setHours(0,0,0,0);    if (today < mulai)   continue; }
    if (selesai) { selesai.setHours(23,59,59);  if (today > selesai) continue; }
    list.push({
      id: String(id), judul: String(judul), isi: String(isi||''),
      tipe: String(tipe||'info'),
      tglMulai:   tglMulai   ? Utilities.formatDate(new Date(tglMulai),  'Asia/Jakarta','dd MMM yyyy') : '',
      tglSelesai: tglSelesai ? Utilities.formatDate(new Date(tglSelesai),'Asia/Jakarta','dd MMM yyyy') : '',
    });
  }
  return list.reverse();
}

function getAllPengumumanAdmin() {
  // Kembalikan semua (untuk admin), termasuk yang nonaktif
  const sheet = getSheet(CONFIG.SHEET_PENGUMUMAN);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  return rows.slice(1).filter(r=>r[0]).map(r=>({
    id: String(r[0]), judul: String(r[1]||''), isi: String(r[2]||''),
    tipe: String(r[3]||'info'), aktif: String(r[4]||'tidak'),
    tglMulai:   r[5] ? Utilities.formatDate(new Date(r[5]),'Asia/Jakarta','yyyy-MM-dd') : '',
    tglSelesai: r[6] ? Utilities.formatDate(new Date(r[6]),'Asia/Jakarta','yyyy-MM-dd') : '',
    dibuat:     r[7] ? Utilities.formatDate(new Date(r[7]),'Asia/Jakarta','dd MMM yyyy HH:mm') : '',
  })).reverse();
}

function handleSavePengumuman(data) {
  const sheet = getSheet(CONFIG.SHEET_PENGUMUMAN);
  const { id, judul, isi, tipe, aktif, tglMulai, tglSelesai } = data;
  if (!judul) return jsonResponse({ ok:false, msg:'Judul wajib diisi' });
  const rows = sheet.getDataRange().getValues();
  if (id) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.getRange(i+1,2,1,6).setValues([[judul,isi||'',tipe||'info',aktif||'ya',tglMulai||'',tglSelesai||'']]);
        writeLog('EDIT_PENGUMUMAN','Update: '+judul);
        return jsonResponse({ ok:true, msg:'Pengumuman diperbarui' });
      }
    }
  }
  const newId = 'p'+new Date().getTime();
  sheet.appendRow([newId,judul,isi||'',tipe||'info',aktif||'ya',tglMulai||'',tglSelesai||'',new Date()]);
  writeLog('ADD_PENGUMUMAN','Tambah: '+judul);
  return jsonResponse({ ok:true, msg:'Pengumuman ditambahkan', id:newId });
}

function handleDeletePengumuman(data) {
  const sheet = getSheet(CONFIG.SHEET_PENGUMUMAN);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i+1);
      writeLog('DELETE_PENGUMUMAN','Hapus id: '+data.id);
      return jsonResponse({ ok:true, msg:'Pengumuman dihapus' });
    }
  }
  return jsonResponse({ ok:false, msg:'Tidak ditemukan' });
}

function getKelasList() {
  // Baca dari sheet guru sebagai referensi, atau hardcode dari data asli
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const propKelas = PropertiesService.getScriptProperties().getProperty('KELAS_LIST');
  if (propKelas) return JSON.parse(propKelas);
  // Default dari file asli
  const defaultKelas = [];
  ['X','XI','XII'].forEach(t => {
    for (let i = 1; i <= 11; i++) defaultKelas.push(`KELAS ${t} - ${i}`);
  });
  return defaultKelas;
}

function getAllData() {
  const guruData  = getGuruData();
  const mapelData = getMapelData();
  const jadwalData = getJadwalData();
  const conflicts  = detectConflicts(jadwalData, guruData);
  return {
    ok: true,
    guru:      guruData,
    mapel:     mapelData,
    jadwal:    jadwalData,
    kelas:     getKelasList(),
    conflicts: conflicts,
    jam_label:        JAM_LABEL_REGULER,
    jam_label_reguler: JAM_LABEL_REGULER,
    jam_label_jumat:   JAM_LABEL_JUMAT,
    wali:       getWaliKelasData(),
    piket:      getPiketData(),
    karyawan:   getKaryawanData(),
    pengumuman: getPengumumanData(),
    kalender:   getKalenderData(),
  };
}

function getPublicData() {
  return { ok: true, ...getAllData() };
}

// ── GURU ─────────────────────────────────────────────────────

function getGuruData() {
  const sheet = getSheet(CONFIG.SHEET_GURU);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return {};
  const guru = {};
  for (let i = 1; i < rows.length; i++) {
    const [kode, nama, mapel] = rows[i];
    if (kode) guru[String(kode)] = { nama: String(nama), mapel: String(mapel) };
  }
  return guru;
}

function handleSaveGuru(data) {
  const { kode, nama, mapel } = data;
  if (!kode || !nama) return jsonResponse({ ok: false, msg: 'Kode dan nama wajib diisi' });
  const sheet = getSheet(CONFIG.SHEET_GURU);
  const rows  = sheet.getDataRange().getValues();
  // Cari baris existing
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(kode)) {
      sheet.getRange(i+1, 1, 1, 3).setValues([[kode, nama, mapel]]);
      writeLog('EDIT_GURU', `Update guru kode ${kode}: ${nama}`);
      return jsonResponse({ ok: true, msg: 'Guru berhasil diupdate' });
    }
  }
  // Tambah baru
  sheet.appendRow([kode, nama, mapel]);
  writeLog('ADD_GURU', `Tambah guru kode ${kode}: ${nama}`);
  return jsonResponse({ ok: true, msg: 'Guru berhasil ditambahkan' });
}

function handleDeleteGuru(data) {
  const { kode } = data;
  const sheet = getSheet(CONFIG.SHEET_GURU);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(kode)) {
      sheet.deleteRow(i+1);
      writeLog('DELETE_GURU', `Hapus guru kode ${kode}`);
      return jsonResponse({ ok: true, msg: 'Guru berhasil dihapus' });
    }
  }
  return jsonResponse({ ok: false, msg: 'Guru tidak ditemukan' });
}

// ── MAPEL ─────────────────────────────────────────────────────

function getMapelData() {
  const sheet = getSheet(CONFIG.SHEET_MAPEL);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const [nama, kode, kelompok, warna] = rows[i];
    if (nama) list.push({ nama: String(nama), kode: String(kode||''), kelompok: String(kelompok||'Umum'), warna: String(warna||'teal') });
  }
  return list;
}

function handleSaveMapel(data) {
  const { nama, kode, kelompok, warna, editNama } = data;
  if (!nama) return jsonResponse({ ok: false, msg: 'Nama mapel wajib diisi' });
  const sheet = getSheet(CONFIG.SHEET_MAPEL);
  const rows  = sheet.getDataRange().getValues();
  const searchNama = editNama || nama;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === searchNama) {
      sheet.getRange(i+1, 1, 1, 4).setValues([[nama, kode, kelompok, warna]]);
      writeLog('EDIT_MAPEL', `Update mapel: ${nama}`);
      return jsonResponse({ ok: true, msg: 'Mapel berhasil diupdate' });
    }
  }
  sheet.appendRow([nama, kode, kelompok, warna]);
  writeLog('ADD_MAPEL', `Tambah mapel: ${nama}`);
  return jsonResponse({ ok: true, msg: 'Mapel berhasil ditambahkan' });
}

function handleDeleteMapel(data) {
  const { nama } = data;
  const sheet = getSheet(CONFIG.SHEET_MAPEL);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === nama) {
      sheet.deleteRow(i+1);
      writeLog('DELETE_MAPEL', `Hapus mapel: ${nama}`);
      return jsonResponse({ ok: true, msg: 'Mapel berhasil dihapus' });
    }
  }
  return jsonResponse({ ok: false, msg: 'Mapel tidak ditemukan' });
}

// ── WALI KELAS (PEMBIMBING AKADEMIK) ────────────────────────────

function getWaliKelasData() {
  const sheet = getSheet(CONFIG.SHEET_WALI);
  const rows  = sheet.getDataRange().getValues();
  const wali  = {};
  for (let i = 1; i < rows.length; i++) {
    const [kelas, kode] = rows[i];
    if (kelas) wali[String(kelas)] = kode ? String(kode) : '';
  }
  return wali;
}

function handleSaveWali(data) {
  const { kelas, kodeGuru } = data;
  if (!kelas) return jsonResponse({ ok: false, msg: 'Kelas wajib diisi' });
  const sheet = getSheet(CONFIG.SHEET_WALI);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(kelas)) {
      sheet.getRange(i+1, 2).setValue(kodeGuru || '');
      writeLog('SAVE_WALI', `Wali kelas ${kelas} → guru ${kodeGuru}`);
      return jsonResponse({ ok: true, msg: 'Wali kelas berhasil disimpan' });
    }
  }
  sheet.appendRow([kelas, kodeGuru || '']);
  writeLog('ADD_WALI', `Wali kelas ${kelas} → guru ${kodeGuru} (baris baru)`);
  return jsonResponse({ ok: true, msg: 'Wali kelas berhasil disimpan' });
}

// ── PIKET GURU HARIAN ─────────────────────────────────────────

function getPiketData() {
  const sheet = getSheet(CONFIG.SHEET_PIKET);
  const rows  = sheet.getDataRange().getValues();
  const piket = {};
  for (let i = 1; i < rows.length; i++) {
    const hari = String(rows[i][0]).toUpperCase().trim();
    if (!HARI_ORDER.includes(hari)) continue;
    try {
      const raw = String(rows[i][1] || '').trim();
      if (raw.startsWith('[')) {
        // Format baru: JSON array [{kode, pin, ket}]
        const parsed = JSON.parse(raw);
        // Filter slot kosong dan pastikan kode tidak kosong
        piket[hari] = parsed.filter(s => s && s.kode && String(s.kode).trim() !== '');
      } else {
        // Format lama: kolom terpisah (guru1=col2, guru2=col3, dst)
        // Kolom 0=hari, 1=data_json(kosong), 2..=kode guru
        const list = [];
        for (let c = 2; c < rows[i].length && list.length < 15; c++) {
          const v = String(rows[i][c] || '').trim();
          if (v && v !== '' && v !== '[]') list.push({ kode: v, pin: false, ket: '' });
        }
        piket[hari] = list;
      }
    } catch(e) {
      piket[hari] = [];
    }
  }
  return piket;
}

function handleSavePiket(data) {
  // data: { hari, slots: [{kode, pin, ket}, ...] } — simpan seluruh hari sekaligus
  const { hari, slots } = data;
  if (!hari || !Array.isArray(slots)) return jsonResponse({ ok: false, msg: 'Data tidak lengkap' });
  if (slots.length > 15) return jsonResponse({ ok: false, msg: 'Maksimal 15 guru piket per hari' });
  const sheet = getSheet(CONFIG.SHEET_PIKET);
  const rows  = sheet.getDataRange().getValues();
  const json  = JSON.stringify(slots);
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toUpperCase().trim() === hari.toUpperCase()) {
      sheet.getRange(i+1, 2).setValue(json);
      writeLog('SAVE_PIKET', `Piket ${hari} → ${slots.length} guru`);
      return jsonResponse({ ok: true, msg: `Piket ${hari} berhasil disimpan` });
    }
  }
  sheet.appendRow([hari.toUpperCase(), json]);
  writeLog('ADD_PIKET', `Piket ${hari} → ${slots.length} guru (baru)`);
  return jsonResponse({ ok: true, msg: `Piket ${hari} berhasil disimpan` });
}

// ── JADWAL ────────────────────────────────────────────────────

function getJadwalData() {
  const sheet   = getSheet(CONFIG.SHEET_JADWAL);
  const rows    = sheet.getDataRange().getValues();
  if (rows.length < 2) return {};
  const headers = rows[0]; // ['hari','jam','KELAS X - 1',...]
  const jadwal  = {};
  HARI_ORDER.forEach(h => {
    jadwal[h] = {};
    JAM_LIST.forEach(j => { jadwal[h][j] = {}; });
  });
  for (let i = 1; i < rows.length; i++) {
    const hari = String(rows[i][0]).toUpperCase().trim();
    const jam  = parseInt(rows[i][1]);
    if (!HARI_ORDER.includes(hari) || !jam) continue;
    for (let c = 2; c < headers.length; c++) {
      const kelas = headers[c];
      const val   = rows[i][c];
      if (val !== '' && val !== null && val !== undefined) {
        if (!jadwal[hari]) jadwal[hari] = {};
        if (!jadwal[hari][jam]) jadwal[hari][jam] = {};
        jadwal[hari][jam][kelas] = parseInt(val) || val;
      }
    }
  }
  return jadwal;
}

function handleSaveJadwal(data) {
  const { hari, jam, kelas, kodeGuru } = data;
  if (!hari || !jam || !kelas) return jsonResponse({ ok: false, msg: 'Data tidak lengkap' });
  const sheet   = getSheet(CONFIG.SHEET_JADWAL);
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  const kelasCol = headers.indexOf(kelas);
  if (kelasCol < 0) return jsonResponse({ ok: false, msg: `Kelas ${kelas} tidak ditemukan di header` });

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toUpperCase().trim() === hari.toUpperCase() &&
        parseInt(rows[i][1]) === parseInt(jam)) {
      sheet.getRange(i+1, kelasCol+1).setValue(kodeGuru || '');
      writeLog('SAVE_JADWAL', `${hari} jam ${jam} ${kelas} → guru ${kodeGuru}`);
      return jsonResponse({ ok: true, msg: 'Jadwal disimpan' });
    }
  }
  // Baris belum ada, buat baru
  const newRow = new Array(headers.length).fill('');
  newRow[0] = hari.toUpperCase();
  newRow[1] = jam;
  newRow[kelasCol] = kodeGuru || '';
  sheet.appendRow(newRow);
  writeLog('ADD_JADWAL', `${hari} jam ${jam} ${kelas} → guru ${kodeGuru} (baris baru)`);
  return jsonResponse({ ok: true, msg: 'Jadwal disimpan (baris baru)' });
}

function handleClearJadwal(data) {
  const { hari, jam, kelas } = data;
  return handleSaveJadwal({ ...data, kodeGuru: '' });
}

// ── KONFLIK ───────────────────────────────────────────────────

function detectConflicts(jadwal, guru) {
  const conflicts = [];
  HARI_ORDER.forEach(hari => {
    JAM_LIST.forEach(jam => {
      const slot = (jadwal[hari] || {})[jam] || {};
      const guruSeen = {};
      Object.entries(slot).forEach(([kelas, kode]) => {
        if (!kode) return;
        const k = String(kode);
        if (!guruSeen[k]) guruSeen[k] = [];
        guruSeen[k].push(kelas);
      });
      Object.entries(guruSeen).forEach(([kode, kelasList]) => {
        if (kelasList.length > 1) {
          const g = guru[kode] || { nama: `Kode ${kode}`, mapel: '?' };
          conflicts.push({ hari, jam, kode: parseInt(kode), namaGuru: g.nama, mapel: g.mapel, kelas: kelasList });
        }
      });
    });
  });
  return conflicts;
}

// ── UTILS ─────────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function writeLog(aksi, ket) {
  try {
    const sheet = getSheet(CONFIG.SHEET_LOG);
    sheet.appendRow([new Date().toLocaleString('id-ID'), aksi, ket]);
  } catch(e) {}
}

// ── SETUP AWAL (jalankan sekali dari menu) ────────────────────

function setupFromExistingData() {
  // Fungsi ini akan dipanggil sekali untuk mengisi sheet dari data yang sudah ada
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ui = SpreadsheetApp.getUi();

  // Setup sheet Jadwal dengan header
  const kelasList = getKelasList();
  const jadwalSheet = ss.getSheetByName(CONFIG.SHEET_JADWAL) || ss.insertSheet(CONFIG.SHEET_JADWAL);
  jadwalSheet.clearContents();
  const hdr = ['hari','jam'].concat(kelasList);
  jadwalSheet.appendRow(hdr);
  jadwalSheet.getRange(1,1,1,hdr.length).setFontWeight('bold').setBackground('#1a7a54').setFontColor('#ffffff');

  // Isi baris jadwal (hari x jam)
  HARI_ORDER.forEach(hari => {
    JAM_LIST.forEach(jam => {
      const row = [hari, jam].concat(new Array(kelasList.length).fill(''));
      jadwalSheet.appendRow(row);
    });
  });

  // Freeze header
  jadwalSheet.setFrozenRows(1);
  jadwalSheet.setFrozenColumns(2);

  ui.alert('✅ Setup selesai!\n\nSheet Jadwal sudah disiapkan dengan ' + kelasList.length + ' kolom kelas.\n\nSelanjutnya:\n1. Import data guru ke sheet DataGuru\n2. Isi jadwal via admin panel\n3. Deploy sebagai Web App');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Jadwal Sekolah')
    .addItem('Setup Awal (jalankan sekali)', 'setupFromExistingData')
    .addItem('Migrasi Sheet Piket ke Format Baru', 'migrasiPiketKeFormatBaru')
    .addItem('Cek Konflik Sekarang', 'cekKonflikManual')
    .addSeparator()
    .addItem('Buka Admin Panel', 'bukaAdminPanel')
    .addToUi();
}

function migrasiPiketKeFormatBaru() {
  // Konversi sheet PiketGuru dari format lama (banyak kolom) ke format baru (hari | data_json)
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ui    = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('PiketGuru');
  if (!sheet) { ui.alert('Sheet PiketGuru tidak ditemukan!'); return; }

  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];

  // Jika sudah format baru (hanya 2 kolom: hari, data_json)
  if (headers.length <= 2 && String(headers[1]).toLowerCase().includes('json')) {
    ui.alert('Sheet PiketGuru sudah dalam format baru ✅');
    return;
  }

  // Baca semua data lama
  const converted = [];
  for (let i = 1; i < rows.length; i++) {
    const hari = String(rows[i][0]).toUpperCase().trim();
    if (!HARI_ORDER.includes(hari)) continue;
    const slots = [];
    for (let c = 1; c < rows[i].length && slots.length < 15; c++) {
      const v = String(rows[i][c] || '').trim();
      if (v && v !== '' && !v.startsWith('[')) {
        slots.push({ kode: v, pin: false, ket: '' });
      }
    }
    converted.push({ hari, json: JSON.stringify(slots) });
  }

  // Hapus dan buat ulang sheet
  ss.deleteSheet(sheet);
  const newSheet = ss.insertSheet('PiketGuru');
  newSheet.appendRow(['hari','data_json']);
  newSheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#1a7a54').setFontColor('#ffffff');
  HARI_ORDER.forEach(h => {
    const found = converted.find(c => c.hari === h);
    newSheet.appendRow([h, found ? found.json : '[]']);
  });
  newSheet.setFrozenRows(1);
  newSheet.setColumnWidth(1, 80);
  newSheet.setColumnWidth(2, 600);

  ui.alert(`✅ Migrasi selesai!\n\n${converted.length} hari berhasil dikonversi ke format JSON.\n\nData piket lama tetap tersimpan dalam format baru.`);
}

function cekKonflikManual() {
  const jadwal    = getJadwalData();
  const guru      = getGuruData();
  const conflicts = detectConflicts(jadwal, guru);
  const ui = SpreadsheetApp.getUi();
  if (!conflicts.length) {
    ui.alert('✅ Tidak ada konflik jadwal!');
  } else {
    const msg = conflicts.map(c =>
      `⚠ ${c.namaGuru} — ${c.hari} Jam ${c.jam}\n   Mengajar di: ${c.kelas.join(' & ')}`
    ).join('\n\n');
    ui.alert(`Ditemukan ${conflicts.length} konflik:\n\n${msg}`);
  }
}

function bukaAdminPanel() {
  const url = ScriptApp.getService().getUrl() + '?page=admin';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`<script>window.open('${url}','_blank');google.script.host.close();</script>`),
    'Membuka Admin Panel...'
  );
}

// ── HTML PAGES ────────────────────────────────────────────────

function loginPage() {
  return `<!DOCTYPE html><html lang="id"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Login Admin | Jadwal KBM SMANSABA</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f4f8;display:flex;align-items:center;justify-content:center;min-height:100vh}
.card{background:#fff;border-radius:16px;padding:36px 32px;width:100%;max-width:380px;box-shadow:0 4px 24px rgba(0,0,0,.10)}
.logo{text-align:center;margin-bottom:24px}
.logo h1{font-size:22px;color:#1a7a54;font-weight:700;margin-top:8px}
.logo p{font-size:13px;color:#718096;margin-top:4px}
label{font-size:13px;font-weight:600;color:#4a5568;display:block;margin-bottom:5px}
input{width:100%;padding:10px 14px;border:1.5px solid #e2e8f0;border-radius:8px;font-size:14px;margin-bottom:16px;transition:.2s}
input:focus{outline:none;border-color:#1a7a54}
button{width:100%;padding:11px;background:#1a7a54;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;transition:.2s}
button:hover{background:#145e3f}
.err{color:#c0392b;font-size:13px;margin-top:8px;text-align:center;display:none}
</style></head>
<body><div class="card">
<div class="logo">📅<h1>Admin Panel</h1><p>Sistem Jadwal Sekolah</p></div>
<label>Password Admin</label>
<input type="password" id="pw" placeholder="Masukkan password..." onkeydown="if(event.key==='Enter')login()">
<button onclick="login()">Masuk</button>
<div class="err" id="err">Password salah. Coba lagi.</div>
</div>
<script>
const SCRIPT_URL = '${ScriptApp.getService().getUrl()}';
async function login() {
  const btn = document.querySelector('button');
  const errEl = document.getElementById('err');
  errEl.style.display = 'none';
  errEl.textContent = 'Memproses...';
  const pw = document.getElementById('pw').value;
  if (!pw) { errEl.textContent='Password tidak boleh kosong.'; errEl.style.display='block'; return; }
  btn.disabled = true; btn.textContent = 'Memproses...';
  try {
    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'login', password: pw })
    });
    const d = await res.json();
    if (d.ok) {
      sessionStorage.setItem('adminToken', d.token);
      window.location.href = SCRIPT_URL + '?page=admin&token=' + d.token;
    } else {
      errEl.textContent = d.msg || 'Password salah. Coba lagi.';
      errEl.style.display = 'block';
    }
  } catch(err) {
    errEl.textContent = 'Gagal terhubung ke server: ' + err.message;
    errEl.style.display = 'block';
  } finally {
    btn.disabled = false; btn.textContent = 'Masuk';
  }
}
</script></body></html>`;
}

function adminPage() {
  return `<!DOCTYPE html><html lang="id"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Admin Panel | Jadwal KBM SMANSABA</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{--green:#1a7a54;--green-l:#e6f5ee;--red:#c0392b;--red-l:#fdecea;--amber:#b06a00;--amber-l:#fff3e0;--blue:#1d5fa8;--blue-l:#e8f0fb;--gray:#4a5568;--gray-l:#f7f8fa;--border:#e2e8f0}
body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f4f8;color:#1a202c;font-size:14px}
.topbar{background:var(--green);color:#fff;padding:12px 20px;display:flex;align-items:center;gap:10px;position:sticky;top:0;z-index:50;box-shadow:0 2px 8px rgba(0,0,0,.15)}
.topbar h1{font-size:16px;font-weight:700}
.topbar span{font-size:12px;opacity:.8}
.logout{margin-left:auto;background:rgba(255,255,255,.2);border:none;color:#fff;padding:5px 12px;border-radius:6px;cursor:pointer;font-size:12px}
.nav{display:flex;background:#fff;border-bottom:2px solid var(--border);overflow-x:auto}
.nav button{padding:11px 18px;font-size:13px;font-weight:600;color:var(--gray);border:none;background:none;cursor:pointer;white-space:nowrap;border-bottom:3px solid transparent;margin-bottom:-2px}
.nav button.active{color:var(--green);border-bottom-color:var(--green)}
.nav button:hover:not(.active){background:var(--gray-l)}
.content{padding:16px;max-width:1300px;margin:0 auto}
.section{display:none}.section.active{display:block}
.card{background:#fff;border:1px solid var(--border);border-radius:12px;padding:16px;margin-bottom:14px}
.card-title{font-size:14px;font-weight:700;color:var(--green);margin-bottom:12px}
.form-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;margin-bottom:10px}
.fg{display:flex;flex-direction:column;gap:4px}
.fg label{font-size:12px;font-weight:600;color:var(--gray)}
.fg input,.fg select{padding:8px 10px;border:1.5px solid var(--border);border-radius:7px;font-size:13px}
.fg input:focus,.fg select:focus{outline:none;border-color:var(--green)}
.btn{padding:7px 14px;border-radius:7px;font-size:13px;font-weight:600;cursor:pointer;border:none;transition:.15s}
.btn-green{background:var(--green);color:#fff}.btn-green:hover{background:#145e3f}
.btn-red{background:var(--red-l);color:var(--red);border:1px solid #f5b7b1}.btn-red:hover{background:#f5c6c6}
.btn-gray{background:var(--gray-l);color:var(--gray);border:1px solid var(--border)}
.btn-blue{background:var(--blue-l);color:var(--blue);border:1px solid #b3cef0}
table{width:100%;border-collapse:collapse;font-size:13px}
th{background:var(--green);color:#fff;padding:8px 10px;text-align:left;font-weight:600;font-size:12px}
td{padding:8px 10px;border-bottom:1px solid var(--border)}
tr:hover td{background:var(--gray-l)}
.pill{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700}
.pill-green{background:var(--green-l);color:var(--green)}
.pill-red{background:var(--red-l);color:var(--red)}
.pill-amber{background:var(--amber-l);color:var(--amber)}
.pill-gray{background:#f0f0f0;color:#666}
.pill-blue{background:#e8f0fb;color:#1d5fa8}
.pill-purple{background:#f3e8ff;color:#6d28d9}
.toast{position:fixed;bottom:20px;right:20px;background:#1a7a54;color:#fff;padding:12px 20px;border-radius:10px;font-size:13px;font-weight:600;z-index:999;display:none;animation:fadeIn .3s}
@keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
.toast.err{background:var(--red)}
.stats-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;margin-bottom:14px}
.stat{background:#fff;border:1px solid var(--border);border-radius:10px;padding:12px;text-align:center}
.stat .num{font-size:24px;font-weight:700;color:var(--green)}
.stat .num.red{color:var(--red)}
.stat .lbl{font-size:11px;color:var(--gray);margin-top:2px;text-transform:uppercase;letter-spacing:.3px}
.jadwal-wrap{overflow-x:auto}
.jadwal-tbl{border-collapse:collapse;min-width:700px;width:100%;font-size:12px}
.jadwal-tbl th{background:var(--green);color:#fff;padding:7px;text-align:center;white-space:nowrap}
.jadwal-tbl td{border:1px solid #ddd;padding:2px;height:48px;vertical-align:top}
.jadwal-tbl .j-jam{background:#f7faff;font-size:11px;font-weight:700;color:var(--blue);text-align:center;width:40px}
.slot{border-radius:5px;padding:3px 5px;font-size:11px;line-height:1.3;height:44px;overflow:hidden;cursor:pointer;border-left:3px solid}
.slot:hover{filter:brightness(.92)}
.slot-empty{background:#fafafa;border:1px dashed #ccc;height:44px;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:18px;color:#ccc;border-radius:5px}
.slot-empty:hover{background:#f0f4f0;color:var(--green)}
.conflict-slot{outline:2px solid var(--red)!important}
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:200;align-items:center;justify-content:center;padding:16px}
.modal-bg.show{display:flex}
.modal{background:#fff;border-radius:12px;padding:20px;max-width:400px;width:100%}
.modal h3{font-size:15px;font-weight:700;color:var(--green);margin-bottom:14px}
.conflict-item{background:var(--red-l);border:1px solid #f5b7b1;border-radius:8px;padding:12px;margin-bottom:8px}
.conflict-title{font-weight:700;color:var(--red);font-size:13px}
.conflict-detail{font-size:12px;color:var(--gray);margin-top:4px}
.filter-row{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center}
.filter-row select,.filter-row input{padding:7px 10px;border:1px solid var(--border);border-radius:7px;font-size:13px}
.bar-wrap{display:flex;align-items:center;gap:6px}
.bar{height:6px;background:#e2e8f0;border-radius:3px;flex:1;overflow:hidden}
.bar-fill{height:100%;border-radius:3px;background:var(--green)}
.bar-fill.warn{background:var(--amber)}.bar-fill.over{background:var(--red)}
@media(max-width:600px){.content{padding:10px}.form-grid{grid-template-columns:1fr}}
</style></head>
<body>

<div class="topbar">
  <span style="font-size:20px">📅</span>
  <div><h1>Admin Panel</h1><span>Sistem Jadwal Sekolah</span></div>
  <button class="logout" onclick="logout()">Keluar</button>
</div>

<div class="nav">
  <button class="active" onclick="showTab('dashboard')">📊 Dashboard</button>
  <button onclick="showTab('guru')">👤 Data Guru</button>
  <button onclick="showTab('mapel')">📚 Mata Pelajaran</button>
  <button onclick="showTab('jadwal')">🏫 Jadwal</button>
  <button onclick="showTab('wali')">🧑‍🏫 Wali Kelas</button>
  <button onclick="showTab('piket')">🗓️ Piket Harian</button>
  <button onclick="showTab('karyawan')">👷 Data Karyawan</button>
  <button onclick="showTab('pengumuman')">📢 Pengumuman</button>
  <button onclick="showTab('kalender')">📆 Kalender Akademik</button>
  <button onclick="showTab('beban')">📈 Beban Guru</button>
  <button onclick="showTab('konflik')">⚠️ Konflik</button>
</div>

<div class="content">

<!-- DASHBOARD -->
<div class="section active" id="tab-dashboard">
  <div class="stats-row" id="statsRow"></div>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
    <div class="card">
      <div class="card-title">🏆 Top Beban Mengajar</div>
      <div id="topBeban"></div>
    </div>
    <div class="card">
      <div class="card-title">⚠️ Konflik Jadwal</div>
      <div id="dashKonflik"></div>
    </div>
  </div>
</div>

<!-- GURU -->
<div class="section" id="tab-guru">
  <div class="card">
    <div class="card-title">➕ Tambah / Edit Guru</div>
    <div class="form-grid">
      <div class="fg"><label>Kode Guru</label><input id="g-kode" type="number" placeholder="cth: 75"></div>
      <div class="fg"><label>Nama Lengkap</label><input id="g-nama" placeholder="cth: Budi Santoso, S.Pd."></div>
      <div class="fg"><label>Mata Pelajaran</label><input id="g-mapel" placeholder="cth: Matematika"></div>
    </div>
    <div style="display:flex;gap:8px">
      <button class="btn btn-green" onclick="saveGuru()">💾 Simpan Guru</button>
      <button class="btn btn-gray" onclick="clearGuruForm()">✕ Bersihkan</button>
    </div>
  </div>
  <div class="card">
    <div class="card-title">📋 Daftar Guru</div>
    <div class="filter-row">
      <input id="searchGuru" placeholder="🔍 Cari nama / mapel..." oninput="renderGuruTable()" style="flex:1;max-width:280px">
    </div>
    <div style="overflow-x:auto">
      <table><thead><tr><th>Kode</th><th>Nama</th><th>Mata Pelajaran</th><th>JP/Minggu</th><th>Aksi</th></tr></thead>
      <tbody id="guruTable"></tbody></table>
    </div>
  </div>
</div>

<!-- MAPEL -->
<div class="section" id="tab-mapel">
  <div class="card">
    <div class="card-title">➕ Tambah / Edit Mata Pelajaran</div>
    <div class="form-grid">
      <div class="fg"><label>Nama Mapel</label><input id="m-nama" placeholder="cth: Matematika Wajib"></div>
      <div class="fg"><label>Kode</label><input id="m-kode" placeholder="cth: MAT"></div>
      <div class="fg"><label>Kelompok</label>
        <select id="m-kelompok"><option>Umum</option><option>Peminatan IPA</option><option>Peminatan IPS</option><option>Peminatan Bahasa</option><option>Kejuruan</option><option>Mulok</option></select>
      </div>
    </div>
    <div style="display:flex;gap:8px">
      <button class="btn btn-green" onclick="saveMapel()">💾 Simpan Mapel</button>
      <button class="btn btn-gray" onclick="clearMapelForm()">✕ Bersihkan</button>
    </div>
  </div>
  <div class="card">
    <div class="card-title">📋 Daftar Mata Pelajaran</div>
    <div style="overflow-x:auto">
      <table><thead><tr><th>Nama</th><th>Kode</th><th>Kelompok</th><th>Aksi</th></tr></thead>
      <tbody id="mapelTable"></tbody></table>
    </div>
  </div>
</div>

<!-- JADWAL -->
<div class="section" id="tab-jadwal">
  <div class="filter-row">
    <select id="filterKelasJadwal" onchange="renderJadwal()"></select>
    <select id="filterHariJadwal" onchange="renderJadwal()">
      <option value="">Semua Hari</option>
      <option>SENIN</option><option>SELASA</option><option>RABU</option><option>KAMIS</option><option>JUMAT</option>
    </select>
  </div>
  <div class="card">
    <div class="jadwal-wrap" id="jadwalWrap"></div>
  </div>
</div>

<!-- WALI KELAS -->
<div class="section" id="tab-wali">
  <div class="card">
    <div class="card-title">🧑‍🏫 Wali Kelas / Pembimbing Akademik (PA)</div>
    <p style="font-size:12px;color:#718096;margin-bottom:12px">Setiap kelas memiliki 1 wali kelas. Pilih guru berdasarkan kode yang sudah terdaftar di Data Guru.</p>
    <div class="filter-row">
      <input id="searchWali" placeholder="🔍 Cari kelas atau nama guru..." oninput="renderWaliTable()" style="flex:1;max-width:280px">
    </div>
    <div style="overflow-x:auto">
      <table><thead><tr><th>Kelas</th><th>Wali Kelas</th><th>Aksi</th></tr></thead>
      <tbody id="waliTable"></tbody></table>
    </div>
  </div>
</div>

<!-- PIKET GURU -->
<div class="section" id="tab-piket">
  <div class="card">
    <div class="card-title">🗓️ Jadwal Piket Harian</div>
    <p style="font-size:12px;color:#718096;margin-bottom:12px">Setiap hari (Senin–Jumat) terdapat 7 guru piket yang stand by. Pilih guru untuk masing-masing slot.</p>
    <div id="piketGrid"></div>
  </div>
</div>

<!-- KALENDER AKADEMIK -->
<div class="section" id="tab-kalender">
  <div class="card">
    <div class="card-title">📆 Tambah / Edit Event Kalender</div>
    <input type="hidden" id="kId">
    <div class="form-grid">
      <div class="fg" style="grid-column:1/-1"><label>Judul Event</label>
        <input id="kJudul" placeholder="cth: Ujian Akhir Semester Ganjil"></div>
      <div class="fg" style="grid-column:1/-1"><label>Keterangan</label>
        <textarea id="kKet" rows="2" placeholder="Keterangan tambahan (opsional)..." style="padding:8px 10px;border:1.5px solid var(--border);border-radius:7px;font-size:13px;resize:vertical;font-family:inherit;width:100%"></textarea></div>
      <div class="fg"><label>Kategori</label>
        <select id="kKategori">
          <option value="libur-nasional">🇮🇩 Libur Nasional</option>
          <option value="libur-sekolah">🏫 Libur Sekolah</option>
          <option value="ujian">📝 Ujian</option>
          <option value="kegiatan">🎉 Kegiatan Sekolah</option>
          <option value="lainnya">📌 Lainnya</option>
        </select></div>
      <div class="fg"><label>Tanggal Mulai</label><input type="date" id="kMulai"></div>
      <div class="fg"><label>Tanggal Selesai <span style="font-weight:400;color:#999">(kosong = 1 hari)</span></label>
        <input type="date" id="kSelesai"></div>
    </div>
    <div style="display:flex;gap:8px;margin-top:6px">
      <button class="btn btn-green" onclick="saveKalender()">💾 Simpan</button>
      <button class="btn btn-gray" onclick="clearKalenderForm()">✕ Batal / Baru</button>
    </div>
  </div>
  <div class="card">
    <div class="card-title">📋 Daftar Event</div>
    <div class="filter-row">
      <select id="filterKalenderTahun" onchange="loadKalenderAdmin()" style="padding:7px 10px;border:1px solid var(--border);border-radius:7px;font-size:13px"></select>
      <select id="filterKalenderKat" onchange="renderKalenderTable()" style="padding:7px 10px;border:1px solid var(--border);border-radius:7px;font-size:13px">
        <option value="">Semua Kategori</option>
        <option value="libur-nasional">🇮🇩 Libur Nasional</option>
        <option value="libur-sekolah">🏫 Libur Sekolah</option>
        <option value="ujian">📝 Ujian</option>
        <option value="kegiatan">🎉 Kegiatan Sekolah</option>
        <option value="lainnya">📌 Lainnya</option>
      </select>
    </div>
    <div style="overflow-x:auto"><table>
      <thead><tr><th>Judul</th><th>Kategori</th><th>Tanggal</th><th>Keterangan</th><th>Aksi</th></tr></thead>
      <tbody id="kalenderTable"></tbody>
    </table></div>
  </div>
</div>

<!-- KARYAWAN -->
<div class="section" id="tab-karyawan">
  <div class="card">
    <div class="card-title">👷 Tambah / Edit Karyawan</div>
    <p style="font-size:12px;color:#718096;margin-bottom:12px">Karyawan hanya dapat dipilih di jadwal Piket Harian — tidak masuk dalam jadwal KBM.</p>
    <input type="hidden" id="krId">
    <div class="form-grid">
      <div class="fg"><label>Nama Karyawan</label>
        <input id="krNama" placeholder="cth: Ahmad Syaifudin"></div>
      <div class="fg"><label>Jabatan</label>
        <input id="krJabatan" placeholder="cth: Staf TU, Satpam, Penjaga Sekolah"></div>
    </div>
    <div style="display:flex;gap:8px;margin-top:4px">
      <button class="btn btn-green" onclick="saveKaryawan()">💾 Simpan</button>
      <button class="btn btn-gray" onclick="clearKaryawanForm()">✕ Batal / Baru</button>
    </div>
  </div>
  <div class="card">
    <div class="card-title">📋 Daftar Karyawan</div>
    <div style="overflow-x:auto"><table>
      <thead><tr><th>Nama</th><th>Jabatan</th><th>Aksi</th></tr></thead>
      <tbody id="karyawanTable"></tbody>
    </table></div>
  </div>
</div>

<!-- PENGUMUMAN -->
<div class="section" id="tab-pengumuman">
  <div class="card">
    <div class="card-title">📢 Tambah / Edit Pengumuman</div>
    <input type="hidden" id="pId">
    <div class="form-grid">
      <div class="fg" style="grid-column:1/-1"><label>Judul Pengumuman</label>
        <input id="pJudul" placeholder="cth: Libur Hari Raya Idul Fitri"></div>
      <div class="fg" style="grid-column:1/-1"><label>Isi / Keterangan</label>
        <textarea id="pIsi" rows="3" placeholder="Tulis keterangan lengkap di sini..." style="padding:8px 10px;border:1.5px solid var(--border);border-radius:7px;font-size:13px;resize:vertical;font-family:inherit;width:100%"></textarea></div>
      <div class="fg"><label>Tipe</label>
        <select id="pTipe">
          <option value="info">ℹ️ Info</option>
          <option value="penting">⚠️ Penting</option>
          <option value="libur">🏖️ Libur</option>
          <option value="kegiatan">🎉 Kegiatan</option>
        </select></div>
      <div class="fg"><label>Status</label>
        <select id="pAktif">
          <option value="ya">✅ Aktif (tampil di publik)</option>
          <option value="tidak">🚫 Nonaktif (disembunyikan)</option>
        </select></div>
      <div class="fg"><label>Tanggal Mulai</label><input type="date" id="pMulai"></div>
      <div class="fg"><label>Tanggal Selesai <span style="font-weight:400;color:#999">(kosong = tidak ada batas)</span></label><input type="date" id="pSelesai"></div>
    </div>
    <div style="display:flex;gap:8px;margin-top:6px">
      <button class="btn btn-green" onclick="savePengumuman()">💾 Simpan</button>
      <button class="btn btn-gray" onclick="clearPengumumanForm()">✕ Batal / Baru</button>
    </div>
  </div>
  <div class="card">
    <div class="card-title">📋 Daftar Pengumuman</div>
    <div style="overflow-x:auto"><table>
      <thead><tr><th>Judul</th><th>Tipe</th><th>Status</th><th>Berlaku</th><th>Dibuat</th><th>Aksi</th></tr></thead>
      <tbody id="pengumumanTable"></tbody>
    </table></div>
  </div>
</div>

<!-- BEBAN GURU -->
<div class="section" id="tab-beban">
  <div class="card">
    <div class="card-title">📈 Rekap Beban Mengajar Guru</div>
    <div class="filter-row">
      <input id="searchBebanAdmin" placeholder="🔍 Cari nama / mapel..." oninput="renderBebanAdmin()" style="flex:1;max-width:280px">
    </div>
    <div style="overflow-x:auto">
      <table><thead><tr><th>Nama Guru</th><th>Mata Pelajaran</th><th>JP/Minggu</th><th>Kelas</th><th>Status</th></tr></thead>
      <tbody id="bebanTable"></tbody></table>
    </div>
  </div>
</div>

<!-- KONFLIK -->
<div class="section" id="tab-konflik">
  <div class="card">
    <div class="card-title">⚠️ Deteksi Konflik Jadwal</div>
    <div id="konflikList"></div>
  </div>
</div>

</div>

<!-- MODAL EDIT SLOT -->
<div class="modal-bg" id="slotModal" onclick="if(event.target===this)closeModal()">
  <div class="modal">
    <h3 id="modalTitle">Edit Slot Jadwal</h3>
    <div id="modalInfo" style="font-size:12px;color:#718096;margin-bottom:12px"></div>
    <div class="fg" style="margin-bottom:10px"><label>Kode Guru</label>
      <select id="slotGuru"><option value="">— Kosongkan slot —</option></select>
    </div>
    <div style="display:flex;gap:8px">
      <button class="btn btn-green" onclick="saveSlot()">💾 Simpan</button>
      <button class="btn btn-red" onclick="clearSlotModal()">🗑 Kosongkan</button>
      <button class="btn btn-gray" onclick="closeModal()">Batal</button>
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
const SCRIPT_URL = '${ScriptApp.getService().getUrl()}';
const TOKEN = new URLSearchParams(window.location.search).get('token') || sessionStorage.getItem('adminToken') || '';
const HARI  = ['SENIN','SELASA','RABU','KAMIS','JUMAT'];
const JAM   = [1,2,3,4,5,6,7,8,9,10];
const COLORS = ['#e6f5ee','#e8f0fb','#fff3e0','#f3e8ff','#fdecea','#e0f7fa','#f9fbe7','#fce4ec','#ede7f6','#e0f2f1','#fff8e1','#efebe9','#e8eaf6','#f1f8e9','#fbe9e7'];
const BORDERS = ['#2ea874','#4a90d9','#f09b00','#9b59b6','#e74c3c','#00acc1','#9ccc65','#e91e8c','#7e57c2','#26a69a','#ffca28','#8d6e63','#5c6bc0','#7cb342','#ff7043'];

function jamUntukHariAdmin(hari) {
  if (hari === 'JUMAT') return JAM.filter(j => j !== 9 && j !== 10);
  return JAM;
}
function labelUntukHariAdmin(hari) {
  return hari === 'JUMAT' ? (DATA.jam_label_jumat || {}) : (DATA.jam_label_reguler || DATA.jam_label || {});
}

let DATA = { guru:{}, mapel:[], jadwal:{}, kelas:[], conflicts:[], jam_label:{}, jam_label_reguler:{}, jam_label_jumat:{}, wali:{}, piket:{}, karyawan:[] };
let editSlot = null;

async function api(body) {
  try {
    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ ...body, token: TOKEN })
    });
    return await res.json();
  } catch(err) {
    return { ok: false, msg: 'Gagal terhubung ke server: ' + err.message };
  }
}

async function loadData() {
  if (!TOKEN) { showToast('Token tidak ditemukan, silakan login ulang.', true); setTimeout(logout, 1500); return; }
  showToast('Memuat data...', false, true);
  const d = await api({ action:'getData' });
  if (!d.ok) { showToast('Gagal memuat data: ' + d.msg, true); return; }
  DATA = d;
  renderAll();
  hideToast();
}

function renderAll() {
  renderStats();
  renderGuruTable();
  renderMapelTable();
  updateKelasSelect();
  renderJadwal();
  renderKonflik();
  renderTopBeban();
}

// STATS
function renderStats() {
  const totalJP = Object.values(DATA.jadwal).reduce((s,jams)=>
    s+Object.values(jams).reduce((s2,kd)=>s2+Object.values(kd).filter(v=>v).length,0),0);
  document.getElementById('statsRow').innerHTML = \`
    <div class="stat"><div class="num">\${Object.keys(DATA.guru).length}</div><div class="lbl">Guru</div></div>
    <div class="stat"><div class="num">\${DATA.mapel.length}</div><div class="lbl">Mata Pelajaran</div></div>
    <div class="stat"><div class="num">\${DATA.kelas.length}</div><div class="lbl">Kelas</div></div>
    <div class="stat"><div class="num">\${totalJP}</div><div class="lbl">Slot Terisi</div></div>
    <div class="stat"><div class="num \${DATA.conflicts.length?'red':''}">\${DATA.conflicts.length}</div><div class="lbl">Konflik</div></div>
  \`;
}

// GURU TABLE
function renderGuruTable() {
  const search = (document.getElementById('searchGuru')?.value||'').toLowerCase();
  const bebanMap = computeBeban();
  const rows = Object.entries(DATA.guru)
    .filter(([k,v]) => !search || v.nama.toLowerCase().includes(search) || v.mapel.toLowerCase().includes(search))
    .sort((a,b) => parseInt(a[0])-parseInt(b[0]));
  document.getElementById('guruTable').innerHTML = rows.map(([kode,g]) => {
    const jp = bebanMap[kode]||0;
    const sc = jp>40?'pill-red':jp>30?'pill-amber':'pill-green';
    const sl = jp>40?'Overload':jp>30?'Padat':'Normal';
    return \`<tr>
      <td><span class="pill pill-green">\${kode}</span></td>
      <td><strong>\${g.nama}</strong></td>
      <td>\${g.mapel}</td>
      <td><span class="pill \${sc}">\${jp} JP — \${sl}</span></td>
      <td><button class="btn btn-blue btn-sm" onclick="editGuru('\${kode}')">✏️</button>
          <button class="btn btn-red btn-sm" onclick="deleteGuru('\${kode}')">🗑</button></td>
    </tr>\`;
  }).join('') || '<tr><td colspan="5" style="text-align:center;color:#999;padding:20px">Belum ada data guru</td></tr>';
}

function editGuru(kode) {
  const g = DATA.guru[kode];
  if (!g) return;
  document.getElementById('g-kode').value = kode;
  document.getElementById('g-nama').value = g.nama;
  document.getElementById('g-mapel').value = g.mapel;
  document.getElementById('g-kode').disabled = true;
  showTab('guru');
  window.scrollTo(0,0);
}

async function saveGuru() {
  const kode  = document.getElementById('g-kode').value.trim();
  const nama  = document.getElementById('g-nama').value.trim();
  const mapel = document.getElementById('g-mapel').value.trim();
  if (!kode || !nama) { showToast('Kode dan nama wajib diisi', true); return; }
  const d = await api({ action:'saveGuru', kode, nama, mapel });
  showToast(d.msg, !d.ok);
  if (d.ok) { clearGuruForm(); await loadData(); }
}

async function deleteGuru(kode) {
  if (!confirm(\`Hapus guru kode \${kode}?\`)) return;
  const d = await api({ action:'deleteGuru', kode });
  showToast(d.msg, !d.ok);
  if (d.ok) await loadData();
}

function clearGuruForm() {
  ['g-kode','g-nama','g-mapel'].forEach(id => document.getElementById(id).value='');
  document.getElementById('g-kode').disabled = false;
}

// MAPEL
function renderMapelTable() {
  document.getElementById('mapelTable').innerHTML = DATA.mapel.map((m,i) => \`<tr>
    <td><span style="display:inline-block;width:12px;height:12px;background:\${COLORS[i%15]};border:1.5px solid \${BORDERS[i%15]};border-radius:3px;vertical-align:middle;margin-right:6px"></span>\${m.nama}</td>
    <td style="font-family:monospace">\${m.kode||'—'}</td>
    <td>\${m.kelompok}</td>
    <td><button class="btn btn-blue btn-sm" onclick="editMapel('\${m.nama}')">✏️</button>
        <button class="btn btn-red btn-sm" onclick="deleteMapel('\${m.nama}')">🗑</button></td>
  </tr>\`).join('') || '<tr><td colspan="4" style="text-align:center;color:#999;padding:20px">Belum ada mata pelajaran</td></tr>';
}

function editMapel(nama) {
  const m = DATA.mapel.find(x=>x.nama===nama);
  if (!m) return;
  document.getElementById('m-nama').value = m.nama;
  document.getElementById('m-kode').value = m.kode;
  document.getElementById('m-kelompok').value = m.kelompok;
  showTab('mapel');
}

async function saveMapel() {
  const nama = document.getElementById('m-nama').value.trim();
  const kode = document.getElementById('m-kode').value.trim().toUpperCase();
  const kelompok = document.getElementById('m-kelompok').value;
  if (!nama) { showToast('Nama mapel wajib diisi', true); return; }
  const d = await api({ action:'saveMapel', nama, kode, kelompok, warna:'teal' });
  showToast(d.msg, !d.ok);
  if (d.ok) { clearMapelForm(); await loadData(); }
}

async function deleteMapel(nama) {
  if (!confirm(\`Hapus mata pelajaran: \${nama}?\`)) return;
  const d = await api({ action:'deleteMapel', nama });
  showToast(d.msg, !d.ok);
  if (d.ok) await loadData();
}

function clearMapelForm() {
  ['m-nama','m-kode'].forEach(id => document.getElementById(id).value='');
}

// JADWAL
function updateKelasSelect() {
  const sel = document.getElementById('filterKelasJadwal');
  const cur = sel.value;
  const groups = {};
  DATA.kelas.forEach(k => {
    const t = k.split(' - ')[0];
    if (!groups[t]) groups[t]=[];
    groups[t].push(k);
  });
  sel.innerHTML = Object.entries(groups).map(([t,ks])=>
    \`<optgroup label="\${t}">\${ks.map(k=>\`<option value="\${k}" \${k===cur?'selected':''}>\${k}</option>\`).join('')}</optgroup>\`
  ).join('');
  if (!sel.value && DATA.kelas[0]) sel.value = DATA.kelas[0];
}

function getMapelColor(mapelNama) {
  const idx = DATA.mapel.findIndex(m=>m.nama===mapelNama);
  return idx>=0 ? {bg:COLORS[idx%15],border:BORDERS[idx%15]} : {bg:'#f5f5f5',border:'#bbb'};
}

function renderJadwal() {
  const kelasId   = document.getElementById('filterKelasJadwal')?.value;
  const hariFilter = document.getElementById('filterHariJadwal')?.value;
  const hariList  = hariFilter ? [hariFilter] : HARI;
  const conflicts = DATA.conflicts || [];

  const guruSel = document.getElementById('slotGuru');
  if (guruSel) {
    guruSel.innerHTML = '<option value="">— Kosongkan slot —</option>' +
      Object.entries(DATA.guru).sort((a,b)=>parseInt(a[0])-parseInt(b[0]))
        .map(([k,g])=>\`<option value="\${k}">[#\${k}] \${g.nama} — \${g.mapel}</option>\`).join('');
  }

  let html = \`<table class="jadwal-tbl"><thead><tr><th>Hari</th><th>JP</th>\${DATA.kelas.map(k=>\`<th style="min-width:80px">\${k.replace('KELAS ','')}</th>\`).join('')}</tr></thead><tbody>\`;

  hariList.forEach(hari => {
    const jamListHari = jamUntukHariAdmin(hari);
    const labelHari = labelUntukHariAdmin(hari);
    jamListHari.forEach((jam,ji) => {
      const isConflictSlot = (kode) => conflicts.some(c=>c.hari===hari&&c.jam===jam&&c.kode===parseInt(kode));
      const slotData = (DATA.jadwal[hari]||{})[jam]||{};

      html += \`<tr>
        <td class="j-jam" style="font-size:11px;font-weight:700;color:#1d5fa8;text-align:center;\${ji===0?'border-top:2px solid #1a7a54':''}">\${ji===0?hari.substring(0,3):''}</td>
        <td class="j-jam">\${jam}<br><span style="font-size:9px;color:#999">\${labelHari[jam]?.split('–')[0]||''}</span></td>\`;

      DATA.kelas.forEach(kelas => {
        const kode = slotData[kelas];
        const g    = kode ? DATA.guru[String(kode)] : null;
        const isConf = kode && isConflictSlot(kode);
        if (g) {
          const c = getMapelColor(g.mapel);
          html += \`<td onclick="openSlot('\${hari}',\${jam},'\${kelas}',\${kode})">
            <div class="slot \${isConf?'conflict-slot':''}" style="background:\${c.bg};border-left-color:\${c.border}">
              <div style="font-weight:700;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">\${g.mapel.substring(0,10)}</div>
              <div style="opacity:.8;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">#\${kode}</div>
            </div></td>\`;
        } else {
          html += \`<td onclick="openSlot('\${hari}',\${jam},'\${kelas}',null)">
            <div class="slot-empty">+</div></td>\`;
        }
      });
      html += '</tr>';

      // Istirahat
      if (jam===4||jam===6) {
        html += \`<tr><td colspan="\${DATA.kelas.length+2}" style="background:#fff8e1;text-align:center;font-size:11px;font-weight:700;color:#b06a00;padding:4px">— ISTIRAHAT —</td></tr>\`;
      }
    });
  });

  html += '</tbody></table>';
  document.getElementById('jadwalWrap').innerHTML = html;
}

function openSlot(hari, jam, kelas, kode) {
  editSlot = { hari, jam, kelas };
  document.getElementById('modalTitle').textContent = \`Edit Slot — \${kelas}\`;
  document.getElementById('modalInfo').textContent  = \`\${hari}, Jam \${jam} (\${labelUntukHariAdmin(hari)[jam]||''})\`;
  const sel = document.getElementById('slotGuru');
  sel.value = kode ? String(kode) : '';
  document.getElementById('slotModal').classList.add('show');
}

function closeModal() {
  document.getElementById('slotModal').classList.remove('show');
  editSlot = null;
}

async function saveSlot() {
  if (!editSlot) return;
  const kodeGuru = document.getElementById('slotGuru').value;
  const d = await api({ action:'saveJadwal', ...editSlot, kodeGuru: kodeGuru ? parseInt(kodeGuru) : '' });
  showToast(d.msg, !d.ok);
  if (d.ok) { closeModal(); await loadData(); }
}

async function clearSlotModal() {
  if (!editSlot) return;
  const d = await api({ action:'clearJadwal', ...editSlot });
  showToast(d.msg, !d.ok);
  if (d.ok) { closeModal(); await loadData(); }
}

// KONFLIK
function renderKonflik() {
  const list = DATA.conflicts || [];
  document.getElementById('konflikList').innerHTML = list.length
    ? list.map(c=>\`<div class="conflict-item">
        <div class="conflict-title">⚠ \${c.namaGuru} — \${c.mapel||'?'}</div>
        <div class="conflict-detail">Hari <strong>\${c.hari}</strong>, Jam <strong>\${c.jam}</strong> — mengajar di: <strong>\${c.kelas.join(' & ')}</strong></div>
      </div>\`).join('')
    : '<div style="background:#e6f5ee;border-radius:8px;padding:16px;text-align:center;color:#1a7a54;font-weight:700">✅ Tidak ada konflik jadwal!</div>';

  document.getElementById('dashKonflik').innerHTML = document.getElementById('konflikList').innerHTML;
}

function renderBebanAdmin() {
  const search = (document.getElementById('searchBebanAdmin')?.value || '').toLowerCase();
  const bebanMap = computeBeban();
  let rows = Object.entries(DATA.guru).map(([kode,g])=>({kode,nama:g.nama,mapel:g.mapel,jp:bebanMap[kode]||0}));
  if (search) rows = rows.filter(r => r.nama.toLowerCase().includes(search) || r.mapel.toLowerCase().includes(search));
  rows.sort((a,b)=>b.jp-a.jp);
  const maxJP = Math.max(...rows.map(r=>r.jp),1);
  document.getElementById('bebanTable').innerHTML = rows.map(r=>\`<tr>
    <td><strong>\${r.nama}</strong></td>
    <td>\${r.mapel}</td>
    <td><div class="bar-wrap"><div class="bar"><div class="bar-fill \${r.jp>40?'over':r.jp>30?'warn':''}" style="width:\${Math.round(r.jp/maxJP*100)}%"></div></div><strong>\${r.jp}</strong></div></td>
    <td>\${computeKelasCount(r.kode)}</td>
    <td><span class="pill \${r.jp>40?'pill-red':r.jp>30?'pill-amber':'pill-green'}">\${r.jp>40?'Overload':r.jp>30?'Padat':'Normal'}</span></td>
  </tr>\`).join('') || '<tr><td colspan="5" style="text-align:center;color:#999;padding:20px">Tidak ditemukan</td></tr>';
}

function renderTopBeban() {
  const bebanMap = computeBeban();
  const rows = Object.entries(DATA.guru).map(([kode,g])=>({nama:g.nama,jp:bebanMap[kode]||0}))
    .sort((a,b)=>b.jp-a.jp).slice(0,8);
  const maxJP = Math.max(...rows.map(r=>r.jp),1);
  document.getElementById('topBeban').innerHTML = rows.map(r=>\`
    <div style="margin-bottom:8px">
      <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:2px">
        <span>\${r.nama.split(',')[0]}</span><span style="font-weight:700;color:#1a7a54">\${r.jp} JP</span>
      </div>
      <div class="bar"><div class="bar-fill \${r.jp>40?'over':r.jp>30?'warn':''}" style="width:\${Math.round(r.jp/maxJP*100)}%"></div></div>
    </div>\`).join('');
}

function computeBeban() {
  const beban = {};
  Object.values(DATA.jadwal).forEach(jams=>Object.values(jams).forEach(kd=>
    Object.values(kd).forEach(kode=>{ if(kode){const k=String(kode);beban[k]=(beban[k]||0)+1;} })
  ));
  return beban;
}

function computeKelasCount(kode) {
  const kelas = new Set();
  Object.values(DATA.jadwal).forEach(jams=>Object.values(jams).forEach(kd=>{if(kd[DATA.kelas.find(x=>kd[x]&&String(kd[x])===String(kode))])kelas.add('x');}));
  // simpler approach
  let count = 0;
  const seen = new Set();
  Object.values(DATA.jadwal).forEach(jams=>Object.values(jams).forEach(kd=>{
    Object.entries(kd).forEach(([k,v])=>{ if(String(v)===String(kode)&&!seen.has(k)){seen.add(k);count++;} });
  }));
  return count;
}

// KARYAWAN ADMIN
async function loadKaryawanAdmin() {
  const d = await api({ action:'getData' });
  if (!d.ok) return;
  const list = d.karyawan || [];
  const tb = document.getElementById('karyawanTable');
  if (!list.length) {
    tb.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#999;padding:20px">Belum ada karyawan. Tambahkan di atas.</td></tr>';
    return;
  }
  tb.innerHTML = list.map(k => \`<tr>
    <td><strong>\${k.nama}</strong></td>
    <td>\${k.jabatan || '—'}</td>
    <td style="font-size:11px;color:#999;font-family:monospace">\${k.id}</td>
    <td style="white-space:nowrap">
      <button class="btn btn-blue btn-sm" onclick="editKaryawan(\${JSON.stringify(k).replace(/"/g,'&quot;')})">✏️</button>
      <button class="btn btn-red btn-sm" onclick="deleteKaryawan('\${k.id}')">🗑</button>
    </td>
  </tr>\`).join('');
  // Simpan ke DATA agar dropdown piket bisa pakai
  DATA.karyawan = list;
  renderPiketGrid();
}

function editKaryawan(k) {
  document.getElementById('krNama').value    = k.nama;
  document.getElementById('krJabatan').value = k.jabatan || '';
  document.getElementById('krId').value      = k.id;
  window.scrollTo(0,0);
}

async function saveKaryawan() {
  const id      = document.getElementById('krId').value;
  const nama    = document.getElementById('krNama').value.trim();
  const jabatan = document.getElementById('krJabatan').value.trim();
  if (!nama) { showToast('Nama karyawan wajib diisi', true); return; }
  const d = await api({ action:'saveKaryawan', id, nama, jabatan });
  showToast(d.msg, !d.ok);
  if (d.ok) { clearKaryawanForm(); loadKaryawanAdmin(); }
}

async function deleteKaryawan(id) {
  if (!confirm('Hapus karyawan ini?')) return;
  const d = await api({ action:'deleteKaryawan', id });
  showToast(d.msg, !d.ok);
  if (d.ok) loadKaryawanAdmin();
}

function clearKaryawanForm() {
  ['krId','krNama','krJabatan'].forEach(i => { const el=document.getElementById(i); if(el) el.value=''; });
}

// WALI KELAS
function renderWaliTable() {
  const search = (document.getElementById('searchWali')?.value || '').toLowerCase();
  let rows = DATA.kelas.map(kelas => {
    const kode = (DATA.wali || {})[kelas] || '';
    const g = kode ? DATA.guru[kode] : null;
    return { kelas, kode, nama: g ? g.nama : '' };
  });
  if (search) rows = rows.filter(r => r.kelas.toLowerCase().includes(search) || r.nama.toLowerCase().includes(search));
  const guruOptions = Object.entries(DATA.guru).sort((a,b)=>parseInt(a[0])-parseInt(b[0]))
    .map(([k,g])=>\`<option value="\${k}">[\${k}] \${g.nama} — \${g.mapel}</option>\`).join('');
  document.getElementById('waliTable').innerHTML = rows.map(r => \`<tr>
    <td><strong>\${r.kelas}</strong></td>
    <td>
      <select onchange="saveWaliInline('\${r.kelas}', this.value)" style="padding:6px 8px;border:1px solid #e2e8f0;border-radius:6px;font-size:12px;min-width:220px">
        <option value="">— Belum ditentukan —</option>
        \${guruOptions.replace(\`value="\${r.kode}"\`, \`value="\${r.kode}" selected\`)}
      </select>
    </td>
    <td>\${r.kode ? \`<span class="pill pill-green">\${r.kode}</span>\` : '<span class="pill pill-gray">—</span>'}</td>
  </tr>\`).join('') || '<tr><td colspan="3" style="text-align:center;color:#999;padding:20px">Tidak ditemukan</td></tr>';
}

async function saveWaliInline(kelas, kodeGuru) {
  const d = await api({ action:'saveWali', kelas, kodeGuru });
  showToast(d.msg, !d.ok);
  if (d.ok) { DATA.wali = DATA.wali || {}; DATA.wali[kelas] = kodeGuru; renderWaliTable(); }
}

// PIKET GURU
function renderPiketGrid() {
  const guruOptions = Object.entries(DATA.guru).sort((a,b)=>parseInt(a[0])-parseInt(b[0]))
    .map(([k,g])=>\`<option value="\${k}">[\${k}] \${g.nama}</option>\`).join('');
  const karyawanOptions = (DATA.karyawan||[]).sort((a,b)=>a.nama.localeCompare(b.nama))
    .map(k=>\`<option value="\${k.id}">\${k.nama}\${k.jabatan?' — '+k.jabatan:''}</option>\`).join('');
  const allOptions = \`<optgroup label="── Guru ──">\${guruOptions}</optgroup>
    \${karyawanOptions ? \`<optgroup label="── Karyawan ──">\${karyawanOptions}</optgroup>\` : ''}\`;

  let html = '';
  HARI.forEach(hari => {
    const namaHari = hari.charAt(0) + hari.slice(1).toLowerCase();
    const slots = (DATA.piket || {})[hari] || [];
    // Pastikan selalu 15 slot
    const rows15 = Array.from({length:15}, (_,i) => slots[i] || {kode:'', pin:false, ket:''});

    html += \`<div style="background:#fff;border:1px solid var(--border);border-radius:10px;padding:14px;margin-bottom:12px">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
        <div style="font-weight:700;font-size:14px;color:var(--green)">\${namaHari}</div>
        <button class="btn btn-green btn-sm" onclick="savePiketHari('\${hari}')">💾 Simpan \${namaHari}</button>
      </div>
      <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead><tr style="background:#f7f8fa">
          <th style="padding:6px 8px;text-align:center;width:36px;border:1px solid var(--border)">#</th>
          <th style="padding:6px 8px;text-align:left;border:1px solid var(--border)">Guru</th>
          <th style="padding:6px 8px;text-align:center;width:56px;border:1px solid var(--border)">📌 Pin</th>
          <th style="padding:6px 8px;text-align:left;border:1px solid var(--border)">Keterangan</th>
        </tr></thead>
        <tbody id="piketBody_\${hari}">
          \${rows15.map((s,i) => \`<tr>
            <td style="padding:4px 6px;text-align:center;border:1px solid var(--border);color:#999">\${i+1}</td>
            <td style="padding:4px 6px;border:1px solid var(--border)">
              <select id="piket_kode_\${hari}_\${i}" style="width:100%;padding:5px;border:1px solid var(--border);border-radius:5px;font-size:11px">
                <option value="">— Pilih guru / karyawan —</option>
                \${allOptions.replace(\`value="\${s.kode}"\`, \`value="\${s.kode}" selected\`)}
              </select>
            </td>
            <td style="padding:4px 6px;text-align:center;border:1px solid var(--border)">
              <input type="checkbox" id="piket_pin_\${hari}_\${i}" \${s.pin?'checked':''} style="width:16px;height:16px;cursor:pointer">
            </td>
            <td style="padding:4px 6px;border:1px solid var(--border)">
              <input type="text" id="piket_ket_\${hari}_\${i}" value="\${s.ket||''}" placeholder="cth: Koordinator Piket" style="width:100%;padding:5px;border:1px solid var(--border);border-radius:5px;font-size:11px">
            </td>
          </tr>\`).join('')}
        </tbody>
      </table>
      </div>
    </div>\`;
  });
  document.getElementById('piketGrid').innerHTML = html;
}

async function savePiketHari(hari) {
  const slots = Array.from({length:15}, (_,i) => {
    const kode = document.getElementById(\`piket_kode_\${hari}_\${i}\`)?.value || '';
    const pin  = document.getElementById(\`piket_pin_\${hari}_\${i}\`)?.checked || false;
    const ket  = document.getElementById(\`piket_ket_\${hari}_\${i}\`)?.value.trim() || '';
    return kode ? { kode, pin, ket } : null;
  }).filter(Boolean);

  const d = await api({ action:'savePiket', hari, slots });
  showToast(d.msg, !d.ok);
  if (d.ok) {
    DATA.piket = DATA.piket || {};
    DATA.piket[hari] = slots;
  }
}

// KALENDER AKADEMIK ADMIN
const KAT_CONFIG = {
  'libur-nasional': { label:'🇮🇩 Libur Nasional', pill:'pill-red',    color:'#dc2626' },
  'libur-sekolah':  { label:'🏫 Libur Sekolah',   pill:'pill-amber',  color:'#d97706' },
  'ujian':          { label:'📝 Ujian',            pill:'pill-purple', color:'#7c3aed' },
  'kegiatan':       { label:'🎉 Kegiatan',         pill:'pill-green',  color:'#15875c' },
  'lainnya':        { label:'📌 Lainnya',          pill:'pill-blue',   color:'#1d5fa8' },
};

let _kalenderAll = [];

async function loadKalenderAdmin() {
  // Isi dropdown tahun
  const selTahun = document.getElementById('filterKalenderTahun');
  const curTahun = selTahun.value || String(new Date().getFullYear());
  const years = [];
  const thisYear = new Date().getFullYear();
  for (let y = thisYear - 1; y <= thisYear + 2; y++) years.push(y);
  selTahun.innerHTML = years.map(y => \`<option value="\${y}" \${String(y)===curTahun?'selected':''}>\${y}</option>\`).join('');

  const d = await api({ action:'getAllKalender' });
  if (!d.ok) { showToast('Gagal memuat kalender', true); return; }
  _kalenderAll = d.list || [];
  renderKalenderTable();
}

function renderKalenderTable() {
  const tahun = document.getElementById('filterKalenderTahun')?.value;
  const kat   = document.getElementById('filterKalenderKat')?.value || '';
  let list = _kalenderAll.filter(e => {
    const matchTahun = !tahun || (e.tglMulai && e.tglMulai.startsWith(tahun));
    const matchKat   = !kat   || e.kategori === kat;
    return matchTahun && matchKat;
  });
  list.sort((a,b) => (a.tglMulai||'').localeCompare(b.tglMulai||''));
  const tb = document.getElementById('kalenderTable');
  if (!list.length) {
    tb.innerHTML = \`<tr><td colspan="5" style="text-align:center;color:#999;padding:20px">Belum ada event untuk tahun \${tahun}</td></tr>\`;
    return;
  }
  tb.innerHTML = list.map(e => {
    const k = KAT_CONFIG[e.kategori] || KAT_CONFIG.lainnya;
    const tgl = e.tglMulai === e.tglSelesai || !e.tglSelesai
      ? formatTgl(e.tglMulai)
      : formatTgl(e.tglMulai) + ' – ' + formatTgl(e.tglSelesai);
    return \`<tr>
      <td><strong>\${e.judul}</strong></td>
      <td><span class="pill \${k.pill}" style="font-size:11px">\${k.label}</span></td>
      <td style="font-size:12px;white-space:nowrap">\${tgl}</td>
      <td style="font-size:12px;color:#718096;max-width:200px">\${e.keterangan||'—'}</td>
      <td style="white-space:nowrap">
        <button class="btn btn-blue btn-sm" onclick='editKalender(\${JSON.stringify(e).replace(/'/g,"&#39;")})'>✏️</button>
        <button class="btn btn-red btn-sm" onclick="deleteKalender('\${e.id}')">🗑</button>
      </td>
    </tr>\`;
  }).join('');
}

function formatTgl(str) {
  if (!str) return '—';
  const d = new Date(str + 'T00:00:00');
  return d.toLocaleDateString('id-ID', { day:'numeric', month:'short', year:'numeric' });
}

function editKalender(e) {
  document.getElementById('kId').value      = e.id;
  document.getElementById('kJudul').value   = e.judul;
  document.getElementById('kKet').value     = e.keterangan||'';
  document.getElementById('kKategori').value= e.kategori;
  document.getElementById('kMulai').value   = e.tglMulai;
  document.getElementById('kSelesai').value = e.tglSelesai||'';
  window.scrollTo(0,0);
}

async function saveKalender() {
  const id        = document.getElementById('kId').value;
  const judul     = document.getElementById('kJudul').value.trim();
  const keterangan= document.getElementById('kKet').value.trim();
  const kategori  = document.getElementById('kKategori').value;
  const tglMulai  = document.getElementById('kMulai').value;
  const tglSelesai= document.getElementById('kSelesai').value;
  if (!judul)    { showToast('Judul wajib diisi', true); return; }
  if (!tglMulai) { showToast('Tanggal mulai wajib diisi', true); return; }
  const d = await api({ action:'saveKalender', id, judul, keterangan, kategori,
    tglMulai, tglSelesai: tglSelesai || tglMulai });
  showToast(d.msg, !d.ok);
  if (d.ok) { clearKalenderForm(); loadKalenderAdmin(); }
}

async function deleteKalender(id) {
  if (!confirm('Hapus event ini?')) return;
  const d = await api({ action:'deleteKalender', id });
  showToast(d.msg, !d.ok);
  if (d.ok) loadKalenderAdmin();
}

function clearKalenderForm() {
  ['kId','kJudul','kKet','kMulai','kSelesai'].forEach(i => { const el=document.getElementById(i); if(el) el.value=''; });
  document.getElementById('kKategori').value = 'libur-nasional';
}

// PENGUMUMAN ADMIN
const TIPE_ADMIN = {
  info:     { label:'ℹ️ Info',     pill:'pill-blue'  },
  penting:  { label:'⚠️ Penting',  pill:'pill-amber' },
  libur:    { label:'🏖️ Libur',    pill:'pill-green' },
  kegiatan: { label:'🎉 Kegiatan', pill:'pill-purple' },
};

async function loadPengumumanAdmin() {
  const d = await api({ action:'getAllPengumuman' });
  if (!d.ok) { showToast('Gagal memuat pengumuman', true); return; }
  const tb = document.getElementById('pengumumanTable');
  if (!d.list.length) {
    tb.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#999;padding:20px">Belum ada pengumuman</td></tr>';
    return;
  }
  tb.innerHTML = d.list.map(p => {
    const t    = TIPE_ADMIN[p.tipe] || TIPE_ADMIN.info;
    const masa = p.tglMulai ? (p.tglMulai + (p.tglSelesai ? ' – '+p.tglSelesai : ' – ∞')) : '—';
    return \`<tr>
      <td><strong>\${p.judul}</strong>
        \${p.isi ? '<div style="font-size:11px;color:#718096;margin-top:2px;max-width:280px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+p.isi+'</div>' : ''}
      </td>
      <td><span class="pill \${t.pill}" style="font-size:11px">\${t.label}</span></td>
      <td><span class="pill \${p.aktif==='ya'?'pill-green':'pill-gray'}">\${p.aktif==='ya'?'Aktif':'Nonaktif'}</span></td>
      <td style="font-size:12px;white-space:nowrap">\${masa}</td>
      <td style="font-size:11px;color:#999;white-space:nowrap">\${p.dibuat}</td>
      <td style="white-space:nowrap">
        <button class="btn btn-blue btn-sm" onclick='editPengumuman(\${JSON.stringify(p)})'>✏️</button>
        <button class="btn btn-red btn-sm" onclick="deletePengumuman('\${p.id}')">🗑</button>
      </td>
    </tr>\`;
  }).join('');
}

function editPengumuman(p) {
  document.getElementById('pId').value     = p.id;
  document.getElementById('pJudul').value  = p.judul;
  document.getElementById('pIsi').value    = p.isi;
  document.getElementById('pTipe').value   = p.tipe;
  document.getElementById('pAktif').value  = p.aktif;
  document.getElementById('pMulai').value  = p.tglMulai;
  document.getElementById('pSelesai').value = p.tglSelesai;
  window.scrollTo(0,0);
}

async function savePengumuman() {
  const id       = document.getElementById('pId').value;
  const judul    = document.getElementById('pJudul').value.trim();
  const isi      = document.getElementById('pIsi').value.trim();
  const tipe     = document.getElementById('pTipe').value;
  const aktif    = document.getElementById('pAktif').value;
  const tglMulai   = document.getElementById('pMulai').value;
  const tglSelesai = document.getElementById('pSelesai').value;
  if (!judul) { showToast('Judul wajib diisi', true); return; }
  const d = await api({ action:'savePengumuman', id, judul, isi, tipe, aktif, tglMulai, tglSelesai });
  showToast(d.msg, !d.ok);
  if (d.ok) { clearPengumumanForm(); loadPengumumanAdmin(); }
}

async function deletePengumuman(id) {
  if (!confirm('Hapus pengumuman ini?')) return;
  const d = await api({ action:'deletePengumuman', id });
  showToast(d.msg, !d.ok);
  if (d.ok) loadPengumumanAdmin();
}

function clearPengumumanForm() {
  ['pId','pJudul','pIsi','pMulai','pSelesai'].forEach(i => { const el=document.getElementById(i); if(el) el.value=''; });
  document.getElementById('pTipe').value  = 'info';
  document.getElementById('pAktif').value = 'ya';
}

// UI HELPERS
function showTab(id) {
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.nav button').forEach(b=>b.classList.remove('active'));
  const target = document.getElementById('tab-'+id);
  if (!target) return;
  target.classList.add('active');
  event?.target?.classList.add('active');
  if (id==='jadwal')     renderJadwal();
  if (id==='konflik')    renderKonflik();
  if (id==='beban')      renderBebanAdmin();
  if (id==='wali')       renderWaliTable();
  if (id==='piket')      renderPiketGrid();
  if (id==='kalender')   loadKalenderAdmin();
  if (id==='pengumuman') loadPengumumanAdmin();
  if (id==='karyawan')   loadKaryawanAdmin();
}

function showToast(msg, isErr=false, persist=false) {
  const t = document.getElementById('toast');
  t.textContent = msg; t.style.display='block';
  t.className = 'toast' + (isErr?' err':'');
  if (!persist) setTimeout(()=>{t.style.display='none'},2800);
}
function hideToast() { document.getElementById('toast').style.display='none'; }

function logout() {
  sessionStorage.removeItem('adminToken');
  window.location.href = SCRIPT_URL + '?page=admin';
}

// INIT
loadData();
</script></body></html>`;
}

function publicPage() {
  return `<!DOCTYPE html><html lang="id"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Aplikasi Jadwal KBM | Demo Sekolah</title>
<meta name="description" content="Jadwal pelajaran sekolah — dapat diakses publik">
<script src="https://cdn.tailwindcss.com"></script>
<script>
  tailwind.config = {
    theme: {
      extend: {
        colors: {
          brand: {
            50:'#eaf1fb',100:'#cfe0f5',200:'#a3c4ec',300:'#6fa3e0',
            400:'#3f80d4',500:'#2b63b3',600:'#1f4e8f',700:'#1a4078',
            800:'#163562',900:'#122a4d'
          }
        },
        fontFamily: { sans: ['Inter','ui-sans-serif','system-ui','Segoe UI','Arial','sans-serif'] }
      }
    }
  }
</script>
<style>
  body{font-family:Inter,'Segoe UI',system-ui,Arial,sans-serif}
  .slot{transition:transform .12s ease,filter .12s ease}
  .slot:hover{filter:brightness(.94);transform:translateY(-1px)}
  .scrollbar-thin::-webkit-scrollbar{height:6px;width:6px}
  .scrollbar-thin::-webkit-scrollbar-thumb{background:transparent;border-radius:99px}
  .scrollbar-thin:hover::-webkit-scrollbar-thumb{background:#cbd5e1}
  .scrollbar-thin::-webkit-scrollbar-track{background:transparent}
  @keyframes spin{to{transform:rotate(360deg)}}
  .spinner{animation:spin .8s linear infinite}
  table.jadwal-tbl{border-collapse:separate;border-spacing:0}
  table.jadwal-tbl th:first-child{border-top-left-radius:.5rem}
  table.jadwal-tbl th:last-child{border-top-right-radius:.5rem}
  .hari-divider td{border-top:3px solid #1f4e8f !important}
  .search-cat{background:#f1f5f9;color:#64748b}
  .search-cat:hover{background:#e2e8f0;color:#334155}
  .active-cat{background:#1a4078 !important;color:#fff !important}
  .search-result-item{display:flex;align-items:flex-start;gap:12px;padding:12px 16px;cursor:pointer;transition:background .1s}
  .search-result-item:hover,.search-result-item.focused{background:#f0f4f8}
  .search-result-icon{width:36px;height:36px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;margin-top:1px}
  .search-result-title{font-weight:600;font-size:14px;color:#1a202c;margin-bottom:2px}
  .search-result-sub{font-size:12px;color:#718096;line-height:1.5}
  mark{background:#dbeafe;color:#1e40af;border-radius:2px;padding:0 1px}
  .search-empty{text-align:center;padding:40px 16px;color:#94a3b8}
  table.jadwal-tbl thead th{position:sticky;top:0;z-index:10;box-shadow:0 1px 0 rgba(0,0,0,.08)}
  .pg-info    {border-left:4px solid #2b63b3;background:#eaf1fb}
  .pg-penting {border-left:4px solid #d97706;background:#fffbeb}
  .pg-libur   {border-left:4px solid #15875c;background:#eefcf5}
  .pg-kegiatan{border-left:4px solid #7c3aed;background:#f5f3ff}
</style>
</head>
<body class="bg-slate-100 text-slate-800 text-[15px] min-h-screen">

<!-- HEADER -->
<header class="bg-gradient-to-r from-brand-700 to-brand-600 text-white shadow-lg shadow-brand-900/10 sticky top-0 z-30">
  <div class="max-w-7xl mx-auto px-4 sm:px-6 py-4 flex items-center gap-3">
    <span class="text-3xl leading-none">📅</span>
    <div class="min-w-0 flex-1">
      <h1 class="text-lg sm:text-xl font-bold tracking-tight truncate">Aplikasi Jadwal KBM</h1>
      <p id="headerSub" class="text-xs sm:text-sm text-brand-100/90">Memuat data...</p>
    </div>
    <button onclick="openSearch()" title="Pencarian Global" class="flex items-center gap-2 bg-white/15 hover:bg-white/25 border border-white/30 text-white px-3 py-2 rounded-xl text-sm font-semibold transition-all shrink-0">
      <span class="text-base">🔍</span>
      <span class="hidden sm:inline">Cari</span>
      <kbd class="hidden sm:inline text-[10px] bg-white/20 px-1.5 py-0.5 rounded font-mono">Ctrl+K</kbd>
    </button>
  </div>
</header>

<!-- SEARCH OVERLAY -->
<div id="searchOverlay" onclick="if(event.target===this)closeSearch()" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[200] items-start justify-center pt-[10vh] px-4">
  <div class="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden">

    <!-- Search input -->
    <div class="flex items-center gap-3 px-4 py-3.5 border-b border-slate-200">
      <span class="text-xl shrink-0">🔍</span>
      <input id="searchInput" type="text" placeholder="Cari guru, kelas, atau mata pelajaran..." oninput="runSearch()" onkeydown="handleSearchKey(event)"
        class="flex-1 text-base text-slate-800 placeholder-slate-400 outline-none bg-transparent">
      <button onclick="closeSearch()" class="text-slate-400 hover:text-slate-600 text-xl leading-none shrink-0">✕</button>
    </div>

    <!-- Kategori filter -->
    <div class="flex gap-2 px-4 py-2.5 border-b border-slate-100 overflow-x-auto">
      <button class="search-cat active-cat shrink-0 px-3 py-1 rounded-full text-xs font-bold transition-colors" data-cat="semua" onclick="setCat('semua',this)">Semua</button>
      <button class="search-cat shrink-0 px-3 py-1 rounded-full text-xs font-bold transition-colors" data-cat="guru" onclick="setCat('guru',this)">👤 Guru</button>
      <button class="search-cat shrink-0 px-3 py-1 rounded-full text-xs font-bold transition-colors" data-cat="kelas" onclick="setCat('kelas',this)">🏫 Kelas</button>
      <button class="search-cat shrink-0 px-3 py-1 rounded-full text-xs font-bold transition-colors" data-cat="mapel" onclick="setCat('mapel',this)">📖 Mapel</button>
      <button class="search-cat shrink-0 px-3 py-1 rounded-full text-xs font-bold transition-colors" data-cat="wali" onclick="setCat('wali',this)">🧑‍🏫 Wali Kelas</button>
      <button class="search-cat shrink-0 px-3 py-1 rounded-full text-xs font-bold transition-colors" data-cat="piket" onclick="setCat('piket',this)">🗓️ Piket</button>
    </div>

    <!-- Results -->
    <div id="searchResults" class="overflow-y-auto max-h-[55vh] divide-y divide-slate-100 scrollbar-thin"></div>

    <!-- Footer hint -->
    <div class="px-4 py-2.5 bg-slate-50 border-t border-slate-100 flex items-center gap-4 text-[11px] text-slate-400">
      <span><kbd class="bg-slate-200 px-1.5 py-0.5 rounded font-mono">↑↓</kbd> navigasi</span>
      <span><kbd class="bg-slate-200 px-1.5 py-0.5 rounded font-mono">Enter</kbd> buka</span>
      <span><kbd class="bg-slate-200 px-1.5 py-0.5 rounded font-mono">Esc</kbd> tutup</span>
    </div>
  </div>
</div>

<!-- LOADING -->
<div id="loading" class="flex flex-col items-center justify-center py-24 text-slate-400 gap-3">
  <div class="spinner w-9 h-9 rounded-full border-4 border-slate-200 border-t-brand-600"></div>
  <span class="text-base">Memuat jadwal...</span>
</div>

<!-- APP -->
<div id="app" class="hidden">

  <!-- NAV TABS -->
  <nav class="bg-white border-b border-slate-200 sticky top-[65px] sm:top-[69px] z-20 shadow-sm">
    <div class="max-w-7xl mx-auto flex overflow-x-auto scrollbar-thin">
      <button class="nav-btn shrink-0 px-5 sm:px-6 py-3.5 text-[15px] font-semibold text-brand-700 border-b-[3px] border-brand-600 -mb-px transition-colors" data-tab="hari-ini" onclick="showTab('hari-ini',this)">📅 Hari Ini</button>
      <button class="nav-btn shrink-0 px-5 sm:px-6 py-3.5 text-[15px] font-semibold text-slate-500 border-b-[3px] border-transparent hover:text-slate-700 hover:bg-slate-50 -mb-px transition-colors" data-tab="jadwal-kelas" onclick="showTab('jadwal-kelas',this)">🏫 Jadwal Kelas</button>
      <button class="nav-btn shrink-0 px-5 sm:px-6 py-3.5 text-[15px] font-semibold text-slate-500 border-b-[3px] border-transparent hover:text-slate-700 hover:bg-slate-50 -mb-px transition-colors" data-tab="jadwal-guru" onclick="showTab('jadwal-guru',this)">👤 Jadwal Guru</button>
      <button class="nav-btn shrink-0 px-5 sm:px-6 py-3.5 text-[15px] font-semibold text-slate-500 border-b-[3px] border-transparent hover:text-slate-700 hover:bg-slate-50 -mb-px transition-colors" data-tab="jadwal-mapel" onclick="showTab('jadwal-mapel',this)">📖 Jadwal Mapel</button>
      <button class="nav-btn shrink-0 px-5 sm:px-6 py-3.5 text-[15px] font-semibold text-slate-500 border-b-[3px] border-transparent hover:text-slate-700 hover:bg-slate-50 -mb-px transition-colors" data-tab="wali-kelas" onclick="showTab('wali-kelas',this)">🧑‍🏫 Wali Kelas</button>
      <button class="nav-btn shrink-0 px-5 sm:px-6 py-3.5 text-[15px] font-semibold text-slate-500 border-b-[3px] border-transparent hover:text-slate-700 hover:bg-slate-50 -mb-px transition-colors" data-tab="piket-guru" onclick="showTab('piket-guru',this)">🗓️ Piket Harian</button>
      <button class="nav-btn shrink-0 px-5 sm:px-6 py-3.5 text-[15px] font-semibold text-slate-500 border-b-[3px] border-transparent hover:text-slate-700 hover:bg-slate-50 -mb-px transition-colors" data-tab="pengumuman" onclick="showTab('pengumuman',this)">📢 Pengumuman</button>
    </div>
  </nav>

  <main class="max-w-7xl mx-auto px-3 sm:px-6 py-5">

    <!-- TAB: HARI INI -->
    <section id="tab-hari-ini" class="tab-section block">

      <!-- Banner info hari & jam sekarang -->
      <div class="bg-gradient-to-r from-brand-700 to-brand-600 text-white rounded-xl p-4 mb-4 flex flex-wrap items-center gap-4 shadow-sm">
        <div class="text-3xl">📅</div>
        <div class="flex-1 min-w-[160px]">
          <div id="hariIniLabel" class="font-bold text-lg sm:text-xl"></div>
          <div id="hariIniTanggal" class="text-sm text-brand-100/90 mt-0.5"></div>
        </div>
        <div class="text-right">
          <div id="jamSekarang" class="font-mono text-2xl font-bold tabular-nums"></div>
          <div id="statusJP" class="text-xs text-brand-100/80 mt-0.5"></div>
        </div>
      </div>

      <!-- Filter kelas & tingkat -->
      <div class="bg-white border border-slate-200 rounded-xl p-4 mb-4 flex flex-wrap gap-4 items-center shadow-sm">
        <div class="flex items-center gap-2">
          <label class="text-sm font-semibold text-slate-500">Tingkat</label>
          <select id="filterTingkatHariIni" onchange="renderHariIni()" class="rounded-lg border border-slate-300 text-base px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-brand-500">
            <option value="">Semua Tingkat</option>
            <option value="KELAS X">Kelas X</option>
            <option value="KELAS XI">Kelas XI</option>
            <option value="KELAS XII">Kelas XII</option>
          </select>
        </div>
        <div class="flex items-center gap-2">
          <label class="text-sm font-semibold text-slate-500">Tampilkan JP</label>
          <select id="filterJPHariIni" onchange="renderHariIni()" class="rounded-lg border border-slate-300 text-base px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-brand-500">
            <option value="semua">Semua JP</option>
            <option value="sekarang">JP Sekarang Saja</option>
            <option value="sisa">JP Tersisa Hari Ini</option>
          </select>
        </div>
        <div id="hariIniStats" class="flex gap-4 ml-auto flex-wrap"></div>
      </div>

      <!-- Piket hari ini -->
      <!-- Banner pengumuman aktif -->
      <div id="pengumumanBanner" class="mb-4"></div>

      <!-- Piket hari ini -->
      <div id="piketHariIni" class="mb-4"></div>

      <!-- Jadwal hari ini -->
      <div class="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
        <div class="overflow-auto max-h-[65vh] scrollbar-thin" id="jadwalHariIniWrap"></div>
      </div>

    </section>

    <!-- TAB: JADWAL KELAS -->
    <section id="tab-jadwal-kelas" class="tab-section hidden">
      <div class="bg-white border border-slate-200 rounded-xl p-4 mb-4 flex flex-wrap gap-4 items-center shadow-sm">
        <div class="flex items-center gap-2">
          <label class="text-sm font-semibold text-slate-500">Kelas</label>
          <select id="filterKelas" onchange="renderJadwalKelas()" class="rounded-lg border border-slate-300 text-base px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-brand-500 focus:border-brand-500"></select>
        </div>
        <div class="flex items-center gap-2">
          <label class="text-sm font-semibold text-slate-500">Hari</label>
          <select id="filterHari" onchange="renderJadwalKelas()" class="rounded-lg border border-slate-300 text-base px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-brand-500 focus:border-brand-500">
            <option value="">Semua Hari</option>
            <option>SENIN</option><option>SELASA</option><option>RABU</option><option>KAMIS</option><option>JUMAT</option>
          </select>
        </div>
      </div>

      <div id="legendKelas" class="flex flex-wrap gap-x-4 gap-y-2 mb-3 px-1 text-sm"></div>

      <div class="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
        <div class="overflow-auto max-h-[70vh] scrollbar-thin" id="jadwalKelasWrap"></div>
      </div>
    </section>

    <!-- TAB: JADWAL GURU -->
    <section id="tab-jadwal-guru" class="tab-section hidden">
      <div class="bg-white border border-slate-200 rounded-xl p-4 mb-4 flex flex-wrap gap-4 items-center shadow-sm">
        <div class="flex items-center gap-2 flex-1 min-w-[260px]">
          <label class="text-sm font-semibold text-slate-500 shrink-0">Guru</label>
          <select id="filterGuru" onchange="renderJadwalGuru()" class="w-full rounded-lg border border-slate-300 text-base px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-brand-500 focus:border-brand-500"></select>
        </div>
      </div>

      <div id="guruInfoBar" class="mb-4"></div>

      <div class="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
        <div class="overflow-auto max-h-[70vh] scrollbar-thin" id="jadwalGuruWrap"></div>
      </div>
    </section>

    <!-- TAB: JADWAL MAPEL -->
    <section id="tab-jadwal-mapel" class="tab-section hidden">
      <div class="bg-white border border-slate-200 rounded-xl p-4 mb-4 flex flex-wrap gap-4 items-center shadow-sm">
        <div class="flex items-center gap-2 flex-1 min-w-[220px]">
          <label class="text-sm font-semibold text-slate-500 shrink-0">Mapel</label>
          <select id="filterMapel" onchange="renderJadwalMapel()" class="w-full rounded-lg border border-slate-300 text-base px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-brand-500 focus:border-brand-500"></select>
        </div>
        <div class="flex items-center gap-2">
          <label class="text-sm font-semibold text-slate-500">Hari</label>
          <select id="filterHariMapel" onchange="renderJadwalMapel()" class="rounded-lg border border-slate-300 text-base px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-brand-500 focus:border-brand-500">
            <option value="">Semua Hari</option>
            <option>SENIN</option><option>SELASA</option><option>RABU</option><option>KAMIS</option><option>JUMAT</option>
          </select>
        </div>
      </div>

      <div id="mapelInfoBar" class="mb-4"></div>

      <div class="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden mb-4">
        <div class="px-4 pt-3 pb-1 text-sm font-bold text-slate-500">📊 Grid Ringkas — Kelas yang Diajar per Hari &amp; Jam</div>
        <div class="overflow-auto max-h-[50vh] scrollbar-thin" id="jadwalMapelGrid"></div>
      </div>

      <div class="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
        <div class="px-4 pt-3 pb-1 text-sm font-bold text-slate-500">📋 Tabel Detail</div>
        <div class="overflow-auto max-h-[70vh] scrollbar-thin" id="jadwalMapelWrap"></div>
      </div>
    </section>

    <!-- TAB: WALI KELAS -->
    <section id="tab-wali-kelas" class="tab-section hidden">
      <div class="bg-white border border-slate-200 rounded-xl p-4 mb-4 flex flex-wrap gap-4 items-center shadow-sm">
        <input id="searchWaliPublik" placeholder="🔍 Cari kelas atau nama guru..." oninput="renderWaliPublik()" class="flex-1 min-w-[200px] max-w-xs rounded-lg border border-slate-300 text-base px-3 py-2 focus:outline-none focus:ring-2 focus:ring-brand-500 focus:border-brand-500">
      </div>
      <div class="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
        <div class="overflow-auto max-h-[70vh] scrollbar-thin">
          <table class="jadwal-tbl w-full text-sm">
            <thead><tr>
              <th class="px-3 py-3 text-left font-semibold bg-brand-700 text-white">Kelas</th>
              <th class="px-3 py-3 text-left font-semibold bg-brand-700 text-white">Wali Kelas / PA</th>
              <th class="px-3 py-3 text-center font-semibold bg-brand-700 text-white w-20">Kode</th>
            </tr></thead>
            <tbody id="waliPublikBody"></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- TAB: PIKET GURU -->
    <section id="tab-piket-guru" class="tab-section hidden">
      <div id="piketPublikContent" class="grid gap-4"></div>
    </section>

    <!-- TAB: PENGUMUMAN -->
    <section id="tab-pengumuman" class="tab-section hidden">
      <div id="pengumumanListPublik"></div>
    </section>

  </main>
</div>

<!-- MODAL DETAIL -->
<div id="modalBg" onclick="if(event.target===this)closeModal()" class="hidden fixed inset-0 bg-slate-900/50 backdrop-blur-[2px] z-50 items-center justify-center p-4">
  <div class="bg-white rounded-2xl p-5 sm:p-6 max-w-sm w-full shadow-2xl">
    <h3 id="modalTitle" class="text-lg font-bold text-brand-700 mb-3">Detail</h3>
    <div id="modalBody" class="divide-y divide-slate-100"></div>
    <button onclick="closeModal()" class="mt-5 w-full py-2.5 bg-brand-600 hover:bg-brand-700 text-white text-base font-semibold rounded-lg transition-colors">Tutup</button>
  </div>
</div>

<script>
const SCRIPT_URL = '${ScriptApp.getService().getUrl()}';
const HARI  = ['SENIN','SELASA','RABU','KAMIS','JUMAT'];
const HARI_LENGKAP = {SENIN:'Senin',SELASA:'Selasa',RABU:'Rabu',KAMIS:'Kamis',JUMAT:'Jumat'};
const JAM   = [1,2,3,4,5,6,7,8,9,10];
const COLORS  = ['#e6f5ee','#e8f0fb','#fff3e0','#f3e8ff','#fdecea','#e0f7fa','#f9fbe7','#fce4ec','#ede7f6','#e0f2f1','#fff8e1','#efebe9','#e8eaf6','#f1f8e9','#fbe9e7'];
const BORDERS = ['#2ea874','#4a90d9','#f09b00','#9b59b6','#e74c3c','#00acc1','#9ccc65','#e91e8c','#7e57c2','#26a69a','#ffca28','#8d6e63','#5c6bc0','#7cb342','#ff7043'];
let D = {};

function jamUntukHari(hari) {
  // Jumat tidak punya JP 9 dan 10
  if (hari === 'JUMAT') return JAM.filter(j => j !== 9 && j !== 10);
  return JAM;
}

function labelUntukHari(hari) {
  return hari === 'JUMAT' ? (D.jam_label_jumat || {}) : (D.jam_label_reguler || D.jam_label || {});
}

// ── PENCARIAN GLOBAL ──────────────────────────────────────────

let searchCat = 'semua';
let searchFocusIdx = -1;

function openSearch() {
  const overlay = document.getElementById('searchOverlay');
  overlay.classList.remove('hidden');
  overlay.classList.add('flex');
  setTimeout(() => {
    const inp = document.getElementById('searchInput');
    inp.value = '';
    inp.focus();
    document.getElementById('searchResults').innerHTML =
      \`<div class="search-empty"><div class="icon">🔍</div><div class="font-semibold text-slate-500">Ketik untuk mencari...</div><div class="text-xs mt-1">Cari nama guru, kelas, atau mata pelajaran</div></div>\`;
  }, 50);
}

function closeSearch() {
  const overlay = document.getElementById('searchOverlay');
  overlay.classList.add('hidden');
  overlay.classList.remove('flex');
  searchFocusIdx = -1;
}

function setCat(cat, btn) {
  searchCat = cat;
  document.querySelectorAll('.search-cat').forEach(b => b.classList.remove('active-cat'));
  btn.classList.add('active-cat');
  runSearch();
}

function highlight(text, q) {
  if (!q) return text;
  const re = new RegExp('(' + q.replace(/[.*+?^${}()|[\]\\\\]/g,'\\\\$&') + ')', 'gi');
  return String(text).replace(re, '<mark>$1</mark>');
}

function runSearch() {
  const q   = (document.getElementById('searchInput').value || '').trim().toLowerCase();
  const el  = document.getElementById('searchResults');
  searchFocusIdx = -1;

  if (!q) {
    el.innerHTML = \`<div class="search-empty"><div class="icon">🔍</div><div class="font-semibold text-slate-500">Ketik untuk mencari...</div><div class="text-xs mt-1">Cari nama guru, kelas, atau mata pelajaran</div></div>\`;
    return;
  }
  if (!D.guru) { el.innerHTML = '<div class="search-empty"><div class="icon">⏳</div><div>Data belum dimuat</div></div>'; return; }

  let results = [];

  // ── Guru ──────────────────────────────────────────────────
  if (searchCat === 'semua' || searchCat === 'guru') {
    Object.entries(D.guru).forEach(([kode, g]) => {
      if (g.nama.toLowerCase().includes(q) || kode.includes(q) || g.mapel.toLowerCase().includes(q)) {
        const beban = computeBebanMap()[kode] || 0;
        results.push({
          type:'guru', icon:'👤', iconBg:'#e8f0fb', iconColor:'#1d5fa8',
          title: g.nama,
          sub: \`Kode <b>\${kode}</b> · \${g.mapel} · \${beban} JP/minggu\`,
          action: () => { closeSearch(); showTab('jadwal-guru', document.querySelector('[data-tab=jadwal-guru]')); setTimeout(()=>{ const s=document.getElementById('filterGuru'); s.value=kode; renderJadwalGuru(); },100); },
          q
        });
      }
    });
  }

  // ── Kelas ─────────────────────────────────────────────────
  if (searchCat === 'semua' || searchCat === 'kelas') {
    (D.kelas || []).forEach(kelas => {
      if (kelas.toLowerCase().includes(q)) {
        const kodeWali = (D.wali || {})[kelas] || '';
        const gWali    = kodeWali ? D.guru[kodeWali] : null;
        results.push({
          type:'kelas', icon:'🏫', iconBg:'#e6f5ee', iconColor:'#15875c',
          title: kelas,
          sub: \`Wali Kelas: \${gWali ? gWali.nama : 'Belum ditentukan'}\`,
          action: () => { closeSearch(); showTab('jadwal-kelas', document.querySelector('[data-tab=jadwal-kelas]')); setTimeout(()=>{ const s=document.getElementById('filterKelas'); s.value=kelas; renderJadwalKelas(); },100); },
          q
        });
      }
    });
  }

  // ── Mata Pelajaran ────────────────────────────────────────
  if (searchCat === 'semua' || searchCat === 'mapel') {
    const mapelSet = {};
    Object.values(D.guru).forEach(g => {
      if (!mapelSet[g.mapel]) mapelSet[g.mapel] = 0;
      mapelSet[g.mapel]++;
    });
    Object.keys(mapelSet).sort().forEach(mapel => {
      if (mapel.toLowerCase().includes(q)) {
        const jumlahGuru = mapelSet[mapel];
        results.push({
          type:'mapel', icon:'📖', iconBg:'#f3e8ff', iconColor:'#7c3aed',
          title: mapel,
          sub: \`Diampu oleh \${jumlahGuru} guru\`,
          action: () => { closeSearch(); showTab('jadwal-mapel', document.querySelector('[data-tab=jadwal-mapel]')); setTimeout(()=>{ const s=document.getElementById('filterMapel'); s.value=mapel; renderJadwalMapel(); },100); },
          q
        });
      }
    });
  }

  // ── Wali Kelas ────────────────────────────────────────────
  if (searchCat === 'semua' || searchCat === 'wali') {
    Object.entries(D.wali || {}).forEach(([kelas, kode]) => {
      const g = kode ? D.guru[kode] : null;
      if (!g) return;
      if (g.nama.toLowerCase().includes(q) || kelas.toLowerCase().includes(q)) {
        results.push({
          type:'wali', icon:'🧑‍🏫', iconBg:'#fff3e0', iconColor:'#b06a00',
          title: \`Wali \${kelas}\`,
          sub: \`\${g.nama} · \${g.mapel}\`,
          action: () => { closeSearch(); showTab('wali-kelas', document.querySelector('[data-tab=wali-kelas]')); },
          q
        });
      }
    });
  }

  // ── Piket ─────────────────────────────────────────────────
  if (searchCat === 'semua' || searchCat === 'piket') {
    Object.entries(D.piket || {}).forEach(([hari, kodeList]) => {
      kodeList.forEach(kode => {
        const g = kode ? D.guru[kode] : null;
        if (!g) return;
        if (g.nama.toLowerCase().includes(q) || kode.includes(q)) {
          results.push({
            type:'piket', icon:'🗓️', iconBg:'#fdecea', iconColor:'#c0392b',
            title: g.nama,
            sub: \`Piket hari \${HARI_LENGKAP[hari] || hari}\`,
            action: () => { closeSearch(); showTab('piket-guru', document.querySelector('[data-tab=piket-guru]')); },
            q
          });
        }
      });
    });
  }

  if (!results.length) {
    el.innerHTML = \`<div class="search-empty"><div class="icon">😕</div><div class="font-semibold text-slate-500">Tidak ditemukan</div><div class="text-xs mt-1">Coba kata kunci lain</div></div>\`;
    return;
  }

  // Batas 40 hasil
  const shown = results.slice(0, 40);
  const typeLabel = { guru:'Guru', kelas:'Kelas', mapel:'Mata Pelajaran', wali:'Wali Kelas', piket:'Piket' };
  el.innerHTML = shown.map((r, idx) => {
    const hl = (t) => highlight(t, r.q);
    return \`<div class="search-result-item" data-idx="\${idx}" onclick="searchResultClick(\${idx})">
      <div class="search-result-icon" style="background:\${r.iconBg}">\${r.icon}</div>
      <div class="min-w-0 flex-1">
        <div class="search-result-title">\${hl(r.title)}</div>
        <div class="search-result-sub">\${r.sub}</div>
      </div>
      <span class="shrink-0 mt-1 text-[10px] px-2 py-0.5 rounded-full font-semibold" style="background:\${r.iconBg};color:\${r.iconColor}">\${typeLabel[r.type]||r.type}</span>
    </div>\`;
  }).join('');

  if (results.length > 40) {
    el.innerHTML += \`<div class="text-center py-3 text-xs text-slate-400">… dan \${results.length-40} hasil lainnya. Persempit pencarian.</div>\`;
  }

  // simpan actions di closure
  el._searchActions = shown.map(r => r.action);
}

function searchResultClick(idx) {
  const el = document.getElementById('searchResults');
  if (el._searchActions && el._searchActions[idx]) el._searchActions[idx]();
}

function handleSearchKey(e) {
  const el    = document.getElementById('searchResults');
  const items = el.querySelectorAll('.search-result-item');
  if (!items.length) {
    if (e.key === 'Escape') closeSearch();
    return;
  }
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    searchFocusIdx = Math.min(searchFocusIdx + 1, items.length - 1);
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    searchFocusIdx = Math.max(searchFocusIdx - 1, 0);
  } else if (e.key === 'Enter') {
    if (searchFocusIdx >= 0) { searchResultClick(searchFocusIdx); return; }
    if (items.length === 1) { searchResultClick(0); return; }
  } else if (e.key === 'Escape') {
    closeSearch(); return;
  } else { return; }
  items.forEach((it, i) => {
    it.classList.toggle('focused', i === searchFocusIdx);
    if (i === searchFocusIdx) it.scrollIntoView({ block:'nearest' });
  });
}

// Keyboard shortcut Ctrl+K / Cmd+K
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'k') { e.preventDefault(); openSearch(); }
  if (e.key === 'Escape' && !document.getElementById('searchOverlay').classList.contains('hidden')) closeSearch();
});

// ── PENGUMUMAN PUBLIK ─────────────────────────────────────────

const TIPE_PG = {
  info:     { icon:'ℹ️', cls:'pg-info',     tc:'#1e40af' },
  penting:  { icon:'⚠️', cls:'pg-penting',  tc:'#92400e' },
  libur:    { icon:'🏖️', cls:'pg-libur',    tc:'#065f46' },
  kegiatan: { icon:'🎉', cls:'pg-kegiatan', tc:'#5b21b6' },
};

function buildPgCard(p, compact) {
  const t = TIPE_PG[p.tipe] || TIPE_PG.info;
  const masa = p.tglMulai ? (p.tglMulai + (p.tglSelesai ? ' – '+p.tglSelesai : '')) : '';
  if (compact) return \`<div class="\${t.cls} rounded-xl px-4 py-3 flex items-start gap-3">
    <span class="text-xl shrink-0 mt-0.5">\${t.icon}</span>
    <div class="flex-1 min-w-0">
      <div class="font-bold text-sm" style="color:\${t.tc}">\${p.judul}</div>
      \${p.isi ? \`<div class="text-xs text-slate-600 mt-0.5">\${p.isi}</div>\` : ''}
    </div>
    \${masa ? \`<div class="text-[11px] text-slate-400 whitespace-nowrap shrink-0 mt-0.5">\${masa}</div>\` : ''}
  </div>\`;
  return \`<div class="\${t.cls} rounded-xl p-5 shadow-sm mb-3">
    <div class="flex items-start gap-3 mb-2">
      <span class="text-2xl shrink-0">\${t.icon}</span>
      <div class="flex-1 min-w-0">
        <div class="font-bold text-base" style="color:\${t.tc}">\${p.judul}</div>
        \${masa ? \`<div class="text-xs text-slate-400 mt-0.5">\${masa}</div>\` : ''}
      </div>
    </div>
    \${p.isi ? \`<div class="text-sm text-slate-700 leading-relaxed whitespace-pre-wrap">\${p.isi}</div>\` : ''}
  </div>\`;
}

function renderPengumumanBanner() {
  const el   = document.getElementById('pengumumanBanner');
  if (!el) return;
  const list = D.pengumuman || [];
  if (!list.length) { el.innerHTML=''; return; }
  const shown = list.slice(0,3);
  const more  = list.length > 3 ? \`<button onclick="showTab('pengumuman',document.querySelector('[data-tab=pengumuman]'))" class="text-xs font-semibold text-brand-600 hover:underline mt-1 text-right block">Lihat semua \${list.length} pengumuman →</button>\` : '';
  el.innerHTML = \`<div class="flex flex-col gap-2">\${shown.map(p=>buildPgCard(p,true)).join('')}\${more}</div>\`;
}

function renderPengumumanPublik() {
  const el   = document.getElementById('pengumumanListPublik');
  const list = D.pengumuman || [];
  if (!list.length) {
    el.innerHTML = \`<div class="text-center py-16 text-slate-400"><div class="text-5xl mb-3">📭</div><div class="font-semibold">Tidak ada pengumuman aktif saat ini</div></div>\`;
    return;
  }
  el.innerHTML = list.map(p => buildPgCard(p, false)).join('');
}

async function fetchData() {
  try {
    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'getPublicData' })
    });
    D = await res.json();
    if (!D.ok) throw new Error(D.msg || 'Server mengembalikan error');
  } catch(err) {
    document.getElementById('loading').innerHTML =
      '<div class="flex flex-col items-center gap-2 text-red-600"><span class="text-2xl">⚠️</span><span class="font-semibold">Gagal memuat data</span><span class="text-xs text-slate-400">'+err.message+'</span></div>';
    return;
  }
  document.getElementById('loading').style.display='none';
  document.getElementById('app').classList.remove('hidden');
  document.getElementById('headerSub').textContent = 'Demo Sekolah — Tahun Ajaran Sekarang';
  initSelects();
  renderHariIni();
  renderPengumumanBanner();
  renderJadwalKelas();
  renderJadwalGuru();
  // Tick jam setiap detik
  tickJam();
  setInterval(tickJam, 1000);
}

function initSelects() {
  const groups={};
  D.kelas.forEach(k=>{ const t=k.split(' - ')[0]; if(!groups[t])groups[t]=[]; groups[t].push(k); });
  document.getElementById('filterKelas').innerHTML = Object.entries(groups)
    .map(([t,ks])=>\`<optgroup label="\${t}">\${ks.map(k=>\`<option value="\${k}">\${k}</option>\`).join('')}</optgroup>\`).join('');
  document.getElementById('filterGuru').innerHTML = Object.entries(D.guru)
    .sort((a,b)=>parseInt(a[0])-parseInt(b[0]))
    .map(([k,g])=>\`<option value="\${k}">[\${k}] \${g.nama} — \${g.mapel}</option>\`).join('');
  const mapelList = [...new Set(Object.values(D.guru).map(g=>g.mapel))].sort((a,b)=>a.localeCompare(b));
  document.getElementById('filterMapel').innerHTML = mapelList
    .map(m=>\`<option value="\${m}">\${m}</option>\`).join('');
}

function getMapelIdx(nama) {
  const list = [...new Set(Object.values(D.guru).map(g=>g.mapel))];
  const i = list.indexOf(nama); return i>=0?i:0;
}

function renderJadwalKelas() {
  const kelasId = document.getElementById('filterKelas').value;
  const hariFilter = document.getElementById('filterHari').value;
  const hariList = hariFilter ? [hariFilter] : HARI;
  const isSemuaHari = !hariFilter;

  const usedMapel = new Set();
  hariList.forEach(h=>jamUntukHari(h).forEach(j=>{const kode=(D.jadwal[h]||{})[j]||{}; const k=kode[kelasId]; if(k){const g=D.guru[String(k)]; if(g)usedMapel.add(g.mapel);}}));
  const mapelArr = [...usedMapel];

  document.getElementById('legendKelas').innerHTML = mapelArr.map(m=>{
    const i=getMapelIdx(m);
    return \`<div class="flex items-center gap-1.5 text-slate-500"><span class="w-3.5 h-3.5 rounded-sm shrink-0" style="background:\${COLORS[i%15]};border:1.5px solid \${BORDERS[i%15]}"></span><span>\${m}</span></div>\`;
  }).join('');

  let html = \`<table class="jadwal-tbl w-full text-sm min-w-[620px]">
    <thead><tr>
      <th class="px-3 py-3 text-center font-semibold w-24 bg-brand-700 text-white">Hari</th>
      <th class="px-2 py-3 text-center font-semibold w-10 bg-brand-700 text-white">JP</th>
      <th class="px-2 py-3 text-center font-semibold w-24 bg-brand-700 text-white">Waktu</th>
      <th class="px-3 py-3 text-left font-semibold bg-brand-700 text-white">Mata Pelajaran</th>
      <th class="px-3 py-3 text-left font-semibold bg-brand-700 text-white">Guru</th>
    </tr></thead><tbody>\`;
  hariList.forEach((hari, hIdx)=>{
    const jamList = jamUntukHari(hari);
    jamList.forEach((jam,ji)=>{
      const kd = (D.jadwal[hari]||{})[jam]||{};
      const kode = kd[kelasId];
      const g = kode ? D.guru[String(kode)] : null;
      const isConf = D.conflicts.some(c=>c.hari===hari&&c.jam===jam&&c.kode===parseInt(kode));
      const i = g ? getMapelIdx(g.mapel) : 0;
      // Garis pembatas tebal di pergantian hari (saat "Semua Hari" dipilih)
      const isDividerRow = isSemuaHari && ji===0 && hIdx>0;
      html += \`<tr class="border-t border-slate-100 hover:bg-slate-50/70 \${isDividerRow?'hari-divider':''}">
        <td class="px-3 py-2 text-center font-bold text-slate-600 bg-slate-100 \${ji===0?'border-t-2 border-t-brand-500':''}">\${ji===0?HARI_LENGKAP[hari]:''}</td>
        <td class="px-2 py-2 text-center font-bold text-slate-600 bg-slate-100">\${jam}</td>
        <td class="px-2 py-2 text-center text-slate-400 whitespace-nowrap text-xs">\${(labelUntukHari(hari)[jam]||'')}</td>\`;
      if (g) {
        html += \`<td class="px-2 py-2">
          <div class="slot rounded-md px-3 py-2 cursor-pointer \${isConf?'ring-2 ring-red-500':''}" style="background:\${COLORS[i%15]};border-left:3px solid \${BORDERS[i%15]}" onclick="showDetail('\${hari}',\${jam},'\${kelasId}',\${kode})">
            <div class="font-bold text-sm leading-tight truncate" style="color:\${BORDERS[i%15]}cc">\${g.mapel}</div>
          </div></td>
          <td class="px-3 py-2 text-sm text-slate-600">\${g.nama}\${isConf?' <span class="inline-block ml-1 px-2 py-0.5 rounded-full bg-red-100 text-red-700 text-xs font-bold">⚠ Bentrok</span>':''}</td>\`;
      } else {
        html += \`<td class="px-2 py-2"><div class="h-[42px] rounded-md border border-dashed border-slate-200 bg-slate-50/60"></div></td><td class="px-3 py-2 text-slate-300 text-sm">—</td>\`;
      }
      html += '</tr>';
      if(jam===4||jam===6) html+=\`<tr><td colspan="5" class="text-center text-xs font-bold text-amber-700 bg-amber-50 py-2">— ISTIRAHAT —</td></tr>\`;
    });
  });
  html += '</tbody></table>';
  document.getElementById('jadwalKelasWrap').innerHTML = html;
}

function renderJadwalGuru() {
  const kode = document.getElementById('filterGuru').value;
  const g = D.guru[kode];
  if (!g) return;
  const beban = computeBebanMap()[kode]||0;
  const kelasSet = new Set();
  HARI.forEach(h=>jamUntukHari(h).forEach(j=>{const kd=(D.jadwal[h]||{})[j]||{}; Object.entries(kd).forEach(([kls,k])=>{if(String(k)===kode)kelasSet.add(kls);});}));
  const i = getMapelIdx(g.mapel);
  document.getElementById('guruInfoBar').innerHTML = \`
    <div class="bg-white border border-slate-200 rounded-xl p-4 flex gap-4 items-center flex-wrap shadow-sm">
      <div class="w-14 h-14 rounded-full flex items-center justify-center text-xl font-extrabold shrink-0" style="background:\${COLORS[i%15]};color:\${BORDERS[i%15]}">\${kode}</div>
      <div class="flex-1 min-w-[160px]">
        <div class="font-bold text-base text-slate-800">\${g.nama}</div>
        <div class="text-sm text-slate-500 mt-0.5">\${g.mapel}</div>
      </div>
      <div class="text-center px-4">
        <div class="text-2xl font-bold text-brand-600">\${beban}</div>
        <div class="text-xs text-slate-400 font-medium">JP/Minggu</div>
      </div>
      <div class="text-center px-4">
        <div class="text-2xl font-bold text-brand-600">\${kelasSet.size}</div>
        <div class="text-xs text-slate-400 font-medium">Kelas</div>
      </div>
    </div>\`;

  let html = \`<table class="jadwal-tbl w-full text-sm min-w-[620px]">
    <thead><tr>
      <th class="px-2 py-3 text-center font-semibold w-12 bg-brand-700 text-white">Jam</th>
      \${HARI.map(h=>\`<th class="px-2 py-3 text-center font-semibold bg-brand-700 text-white">\${HARI_LENGKAP[h]}</th>\`).join('')}
    </tr></thead><tbody>\`;
  JAM.forEach(jam=>{
    html += \`<tr class="border-t border-slate-100">
      <td class="px-2 py-2 text-center font-bold text-slate-600 bg-slate-100">\${jam}</td>\`;
    HARI.forEach(hari=>{
      const jamTersedia = jamUntukHari(hari).includes(jam);
      if (!jamTersedia) {
        html += \`<td class="px-2 py-2"><div class="h-[40px] rounded-md bg-slate-100/70"></div></td>\`;
        return;
      }
      let foundKelas=null;
      const kd=(D.jadwal[hari]||{})[jam]||{};
      Object.entries(kd).forEach(([kls,k])=>{if(String(k)===kode)foundKelas=kls;});
      const isConf=D.conflicts.some(c=>c.hari===hari&&c.jam===jam&&c.kode===parseInt(kode));
      if(foundKelas){
        html+=\`<td class="px-2 py-2">
          <div class="slot rounded-md px-2 py-2 cursor-pointer \${isConf?'ring-2 ring-red-500':''}" style="background:\${COLORS[i%15]};border-left:3px solid \${BORDERS[i%15]}" onclick="showDetail('\${hari}',\${jam},'\${foundKelas}',\${kode})">
            <div class="font-bold text-xs leading-tight truncate" style="color:\${BORDERS[i%15]}cc">\${g.mapel}</div>
            <div class="text-[11px] text-slate-600 truncate">\${foundKelas.replace('KELAS ','')}</div>
          </div></td>\`;
      } else {html+=\`<td class="px-2 py-2"><div class="h-[40px] rounded-md bg-slate-50"></div></td>\`;}
    });
    html+='</tr>';
    if(jam===4||jam===6)html+=\`<tr><td colspan="6" class="text-center text-xs font-bold text-amber-700 bg-amber-50 py-2">— ISTIRAHAT —</td></tr>\`;
  });
  html+='</tbody></table>';
  document.getElementById('jadwalGuruWrap').innerHTML=html;
}

// ── HARI INI ──────────────────────────────────────────────────

const NAMA_HARI_JS  = ['MINGGU','SENIN','SELASA','RABU','KAMIS','JUMAT','SABTU'];
const NAMA_BULAN    = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];

function getHariIni() {
  const now = new Date();
  return NAMA_HARI_JS[now.getDay()]; // 'SENIN' dst.
}

function getJamSekarang() {
  const now = new Date();
  return now.getHours() * 60 + now.getMinutes(); // total menit sejak tengah malam
}

function parseJam(str) {
  // "07:15" → 435 menit
  if (!str) return 0;
  const part = str.split('–')[0].trim();
  const [h, m] = part.split(':').map(Number);
  return h * 60 + m;
}

function parseJamAkhir(str) {
  if (!str) return 0;
  const part = str.split('–')[1]?.trim() || str.split('-')[1]?.trim() || '';
  const [h, m] = part.split(':').map(Number);
  return h * 60 + m;
}

function getJPSekarang(hari) {
  const sekarang = getJamSekarang();
  const labels   = labelUntukHari(hari);
  for (const jam of jamUntukHari(hari)) {
    const mulai  = parseJam(labels[jam]);
    const selesai = parseJamAkhir(labels[jam]);
    if (sekarang >= mulai && sekarang < selesai) return jam;
  }
  return null; // di luar jam pelajaran
}

function tickJam() {
  const now    = new Date();
  const h      = String(now.getHours()).padStart(2,'0');
  const m      = String(now.getMinutes()).padStart(2,'0');
  const s      = String(now.getSeconds()).padStart(2,'0');
  const elJam  = document.getElementById('jamSekarang');
  if (elJam) elJam.textContent = \`\${h}:\${m}:\${s}\`;

  const hari    = getHariIni();
  const jpNow   = getJPSekarang(hari);
  const elStatus = document.getElementById('statusJP');
  if (elStatus) {
    if (!HARI.includes(hari)) {
      elStatus.textContent = 'Bukan hari sekolah';
    } else if (jpNow) {
      const label = labelUntukHari(hari)[jpNow] || '';
      elStatus.textContent = \`Sedang berlangsung: JP \${jpNow} (\${label})\`;
    } else {
      const sekarang = getJamSekarang();
      const labels   = labelUntukHari(hari);
      const jamList  = jamUntukHari(hari);
      const jamMulai = parseJam(labels[jamList[0]]);
      const jamAkhir = parseJamAkhir(labels[jamList[jamList.length-1]]);
      if (sekarang < jamMulai)      elStatus.textContent = 'Belum mulai — KBM belum dimulai';
      else if (sekarang >= jamAkhir) elStatus.textContent = 'KBM telah selesai hari ini';
      else                           elStatus.textContent = 'Sedang istirahat';
    }
  }
}

function renderHariIni() {
  const hari = getHariIni();
  const now  = new Date();
  const tgl  = \`\${now.getDate()} \${NAMA_BULAN[now.getMonth()]} \${now.getFullYear()}\`;

  // Update header banner
  const elHari = document.getElementById('hariIniLabel');
  if (elHari) elHari.textContent = HARI.includes(hari)
    ? \`Hari \${HARI_LENGKAP[hari]}\`
    : \`\${hari.charAt(0)+hari.slice(1).toLowerCase()} — Bukan Hari Sekolah\`;
  const elTgl = document.getElementById('hariIniTanggal');
  if (elTgl) elTgl.textContent = tgl;

  // Piket hari ini
  const piketList = (D.piket || {})[hari] || [];
  const piketEl   = document.getElementById('piketHariIni');
  if (piketEl) {
    if (HARI.includes(hari) && piketList.length) {
      const chips = piketList.map(kode => {
        const g = kode ? D.guru[kode] : null;
        return g ? \`<div class="flex items-center gap-2 bg-white border border-slate-200 rounded-lg px-3 py-1.5 text-sm"><span class="font-extrabold text-brand-700">\${kode}</span><span class="text-slate-700">\${g.nama.split(',')[0]}</span></div>\` : '';
      }).filter(Boolean).join('');
      piketEl.innerHTML = \`<div class="bg-white border border-slate-200 rounded-xl p-4 shadow-sm">
        <div class="text-sm font-bold text-slate-500 mb-2">📋 Guru Piket Hari Ini</div>
        <div class="flex flex-wrap gap-2">\${chips || '<span class="text-slate-400 text-sm">Belum ada data piket</span>'}</div>
      </div>\`;
    } else {
      piketEl.innerHTML = '';
    }
  }

  if (!HARI.includes(hari)) {
    document.getElementById('jadwalHariIniWrap').innerHTML =
      \`<div class="text-center py-16 text-slate-400"><div class="text-5xl mb-3">🏖️</div><div class="font-semibold text-base">Hari ini adalah hari libur</div><div class="text-sm mt-1">Tidak ada jadwal KBM</div></div>\`;
    document.getElementById('hariIniStats').innerHTML = '';
    return;
  }

  const filterTingkat = document.getElementById('filterTingkatHariIni')?.value || '';
  const filterJP      = document.getElementById('filterJPHariIni')?.value || 'semua';
  const sekarang      = getJamSekarang();
  const jpNow         = getJPSekarang(hari);
  const labels        = labelUntukHari(hari);

  // Tentukan jam yang ditampilkan
  let jamTampil = jamUntukHari(hari);
  if (filterJP === 'sekarang') {
    jamTampil = jpNow ? [jpNow] : [];
  } else if (filterJP === 'sisa') {
    jamTampil = jamTampil.filter(j => parseJamAkhir(labels[j]) > sekarang);
  }

  // Kumpulkan semua entri jadwal hari ini
  let entries = [];
  jamTampil.forEach(jam => {
    const kd = (D.jadwal[hari] || {})[jam] || {};
    Object.entries(kd).forEach(([kelas, kode]) => {
      if (!kode) return;
      if (filterTingkat && !kelas.startsWith(filterTingkat)) return;
      const g = D.guru[String(kode)];
      entries.push({ jam, kelas, kode: String(kode), guru: g ? g.nama : \`Kode \${kode}\`, mapel: g ? g.mapel : '—' });
    });
  });

  // Stats
  const totalKelas = new Set(entries.map(e=>e.kelas)).size;
  const totalGuru  = new Set(entries.map(e=>e.kode)).size;
  document.getElementById('hariIniStats').innerHTML = \`
    <div class="text-center"><div class="font-bold text-brand-700 text-lg">\${entries.length}</div><div class="text-xs text-slate-400">JP berjalan</div></div>
    <div class="text-center"><div class="font-bold text-brand-700 text-lg">\${totalKelas}</div><div class="text-xs text-slate-400">Kelas aktif</div></div>
    <div class="text-center"><div class="font-bold text-brand-700 text-lg">\${totalGuru}</div><div class="text-xs text-slate-400">Guru mengajar</div></div>\`;

  if (!entries.length) {
    document.getElementById('jadwalHariIniWrap').innerHTML =
      \`<div class="text-center py-12 text-slate-400"><div class="text-4xl mb-3">📭</div><div class="font-semibold">Tidak ada jadwal yang sesuai filter</div></div>\`;
    return;
  }

  // Urutkan: jam → kelas
  entries.sort((a,b) => a.jam-b.jam || a.kelas.localeCompare(b.kelas));

  // Render tabel
  let lastJam = null;
  let html = \`<table class="jadwal-tbl w-full text-sm min-w-[520px]">
    <thead><tr>
      <th class="px-2 py-3 text-center font-semibold w-10 bg-brand-700 text-white">JP</th>
      <th class="px-3 py-3 text-center font-semibold w-28 bg-brand-700 text-white">Waktu</th>
      <th class="px-3 py-3 text-left font-semibold bg-brand-700 text-white">Kelas</th>
      <th class="px-3 py-3 text-left font-semibold bg-brand-700 text-white">Mata Pelajaran</th>
      <th class="px-3 py-3 text-left font-semibold bg-brand-700 text-white">Guru</th>
    </tr></thead><tbody>\`;

  entries.forEach(e => {
    const isNewJam   = e.jam !== lastJam;
    const isJPNow    = e.jam === jpNow;
    const i          = getMapelIdx(e.mapel);
    const isConf     = D.conflicts.some(c=>c.hari===hari&&c.jam===e.jam&&c.kode===parseInt(e.kode));
    const rowBg      = isJPNow ? 'bg-brand-50/60' : '';
    const jpNowBadge = isJPNow && isNewJam
      ? \` <span class="ml-1 inline-block px-2 py-0.5 rounded-full bg-brand-600 text-white text-[10px] font-bold align-middle">▶ Sekarang</span>\`
      : '';

    html += \`<tr class="border-t \${isNewJam ? 'border-t-2 border-t-brand-400' : 'border-slate-100'} hover:bg-slate-50/70 \${rowBg}">
      <td class="px-2 py-2.5 text-center font-bold text-slate-600 bg-slate-100">\${isNewJam ? e.jam+jpNowBadge : ''}</td>
      <td class="px-3 py-2.5 text-center text-slate-400 whitespace-nowrap text-xs">\${isNewJam ? labels[e.jam]||'' : ''}</td>
      <td class="px-3 py-2.5">
        <div class="slot rounded-md px-2.5 py-1.5 cursor-pointer inline-block \${isConf?'ring-2 ring-red-500':''}" style="background:\${COLORS[i%15]};border-left:3px solid \${BORDERS[i%15]}" onclick="showDetail('\${hari}',\${e.jam},'\${e.kelas}',\${e.kode})">
          <div class="font-bold text-xs" style="color:\${BORDERS[i%15]}cc">\${e.kelas.replace('KELAS ','')}</div>
        </div>
      </td>
      <td class="px-3 py-2.5 font-semibold text-slate-700">\${e.mapel}\${isConf?' <span class="ml-1 px-1.5 py-0.5 rounded bg-red-100 text-red-700 text-[10px] font-bold">⚠</span>':''}</td>
      <td class="px-3 py-2.5 text-slate-600">\${e.guru}</td>
    </tr>\`;
    lastJam = e.jam;
  });

  html += '</tbody></table>';
  document.getElementById('jadwalHariIniWrap').innerHTML = html;
}

function renderJadwalMapel() {
  const sel = document.getElementById('filterMapel');
  if (!sel.value && sel.options.length) sel.value = sel.options[0].value;
  const mapel = sel.value;
  const hariFilter = document.getElementById('filterHariMapel').value;
  if (!mapel) return;
  const hariList = hariFilter ? [hariFilter] : HARI;
  const i = getMapelIdx(mapel);

  // Kumpulkan semua slot untuk mapel ini, lintas kelas & hari
  let entries = [];
  hariList.forEach(hari => {
    jamUntukHari(hari).forEach(jam => {
      const kd = (D.jadwal[hari]||{})[jam] || {};
      Object.entries(kd).forEach(([kelas, kode]) => {
        if (!kode) return;
        const g = D.guru[String(kode)];
        if (g && g.mapel === mapel) {
          entries.push({ hari, jam, kelas, kode: String(kode), guru: g.nama });
        }
      });
    });
  });

  // Info ringkas
  const kelasSet = new Set(entries.map(e=>e.kelas));
  const guruSet  = new Set(entries.map(e=>e.kode));
  document.getElementById('mapelInfoBar').innerHTML = \`
    <div class="bg-white border border-slate-200 rounded-xl p-4 flex gap-4 items-center flex-wrap shadow-sm">
      <div class="w-14 h-14 rounded-full flex items-center justify-center text-2xl shrink-0" style="background:\${COLORS[i%15]};color:\${BORDERS[i%15]}">📖</div>
      <div class="flex-1 min-w-[160px]">
        <div class="font-bold text-base text-slate-800">\${mapel}</div>
        <div class="text-sm text-slate-500 mt-0.5">\${hariFilter ? HARI_LENGKAP[hariFilter] : 'Semua Hari'}</div>
      </div>
      <div class="text-center px-4">
        <div class="text-2xl font-bold text-brand-600">\${entries.length}</div>
        <div class="text-xs text-slate-400 font-medium">Total JP</div>
      </div>
      <div class="text-center px-4">
        <div class="text-2xl font-bold text-brand-600">\${kelasSet.size}</div>
        <div class="text-xs text-slate-400 font-medium">Kelas</div>
      </div>
      <div class="text-center px-4">
        <div class="text-2xl font-bold text-brand-600">\${guruSet.size}</div>
        <div class="text-xs text-slate-400 font-medium">Guru Pengampu</div>
      </div>
    </div>\`;

  if (!entries.length) {
    document.getElementById('jadwalMapelGrid').innerHTML = '';
    document.getElementById('jadwalMapelWrap').innerHTML =
      \`<div class="text-center text-slate-400 py-10 text-sm">Tidak ada jadwal untuk mata pelajaran ini\${hariFilter ? ' pada hari ' + HARI_LENGKAP[hariFilter] : ''}.</div>\`;
    return;
  }

  // ── GRID RINGKAS (Hari × Jam, multi-kelas per slot) ──────────
  let gridHtml = \`<table class="jadwal-tbl w-full text-sm min-w-[620px]">
    <thead><tr>
      <th class="px-2 py-3 text-center font-semibold w-12 bg-brand-700 text-white">Jam</th>
      \${hariList.map(h=>\`<th class="px-2 py-3 text-center font-semibold bg-brand-700 text-white">\${HARI_LENGKAP[h]}</th>\`).join('')}
    </tr></thead><tbody>\`;
  const semuaJam = [...new Set(hariList.flatMap(h=>jamUntukHari(h)))].sort((a,b)=>a-b);
  semuaJam.forEach(jam => {
    gridHtml += \`<tr class="border-t border-slate-100">
      <td class="px-2 py-2 text-center font-bold text-slate-600 bg-slate-100">\${jam}</td>\`;
    hariList.forEach(hari => {
      if (!jamUntukHari(hari).includes(jam)) { gridHtml += \`<td class="px-2 py-2"><div class="h-[40px] rounded-md bg-slate-100/70"></div></td>\`; return; }
      const slotEntries = entries.filter(e=>e.hari===hari&&e.jam===jam);
      if (!slotEntries.length) { gridHtml += \`<td class="px-2 py-2"><div class="h-[40px] rounded-md bg-slate-50"></div></td>\`; return; }
      const chips = slotEntries.map(e=>{
        const isConf = D.conflicts.some(c=>c.hari===e.hari&&c.jam===e.jam&&c.kode===parseInt(e.kode));
        return \`<span class="inline-block px-1.5 py-0.5 rounded text-[10px] font-bold cursor-pointer \${isConf?'ring-2 ring-red-500':''}" style="background:\${COLORS[i%15]};color:\${BORDERS[i%15]}" onclick="showDetail('\${e.hari}',\${e.jam},'\${e.kelas}',\${e.kode})">\${e.kelas.replace('KELAS ','')}</span>\`;
      }).join(' ');
      gridHtml += \`<td class="px-2 py-2"><div class="flex flex-wrap gap-1 items-center min-h-[40px] py-1">\${chips}</div></td>\`;
    });
    gridHtml += '</tr>';
  });
  gridHtml += '</tbody></table>';
  document.getElementById('jadwalMapelGrid').innerHTML = gridHtml;

  // ── TABEL DETAIL ──────────────────────────────────────────────
  // Urutkan: hari → jam
  const hariIdx = h => HARI.indexOf(h);
  entries.sort((a,b) => hariIdx(a.hari)-hariIdx(b.hari) || a.jam-b.jam || a.kelas.localeCompare(b.kelas));

  let html = \`<table class="jadwal-tbl w-full text-sm min-w-[620px]">
    <thead><tr>
      <th class="px-3 py-3 text-center font-semibold w-24 bg-brand-700 text-white">Hari</th>
      <th class="px-2 py-3 text-center font-semibold w-10 bg-brand-700 text-white">JP</th>
      <th class="px-2 py-3 text-center font-semibold w-24 bg-brand-700 text-white">Waktu</th>
      <th class="px-3 py-3 text-left font-semibold bg-brand-700 text-white">Kelas</th>
      <th class="px-3 py-3 text-left font-semibold bg-brand-700 text-white">Guru</th>
    </tr></thead><tbody>\`;

  let lastHari = null;
  entries.forEach(e => {
    const isNewHari = e.hari !== lastHari;
    const isConf = D.conflicts.some(c=>c.hari===e.hari&&c.jam===e.jam&&c.kode===parseInt(e.kode));
    html += \`<tr class="border-t border-slate-100 hover:bg-slate-50/70 \${isNewHari && lastHari!==null ? 'hari-divider' : ''}">
      <td class="px-3 py-2 text-center font-bold text-slate-600 bg-slate-100">\${isNewHari ? HARI_LENGKAP[e.hari] : ''}</td>
      <td class="px-2 py-2 text-center font-bold text-slate-600 bg-slate-100">\${e.jam}</td>
      <td class="px-2 py-2 text-center text-slate-400 whitespace-nowrap text-xs">\${labelUntukHari(e.hari)[e.jam]||''}</td>
      <td class="px-3 py-2">
        <div class="slot rounded-md px-3 py-2 cursor-pointer inline-block \${isConf?'ring-2 ring-red-500':''}" style="background:\${COLORS[i%15]};border-left:3px solid \${BORDERS[i%15]}" onclick="showDetail('\${e.hari}',\${e.jam},'\${e.kelas}',\${e.kode})">
          <div class="font-bold text-sm" style="color:\${BORDERS[i%15]}cc">\${e.kelas.replace('KELAS ','')}</div>
        </div>
      </td>
      <td class="px-3 py-2 text-sm text-slate-600">\${e.guru}\${isConf?' <span class="inline-block ml-1 px-2 py-0.5 rounded-full bg-red-100 text-red-700 text-xs font-bold">⚠ Bentrok</span>':''}</td>
    </tr>\`;
    lastHari = e.hari;
  });

  html += '</tbody></table>';
  document.getElementById('jadwalMapelWrap').innerHTML = html;
}

function computeBebanMap() {
  const b={};
  Object.values(D.jadwal).forEach(jams=>Object.values(jams).forEach(kd=>Object.values(kd).forEach(kode=>{if(kode){const k=String(kode);b[k]=(b[k]||0)+1;}})));
  return b;
}

function showDetail(hari,jam,kelas,kode) {
  const g=D.guru[String(kode)];
  const isConf=D.conflicts.some(c=>c.hari===hari&&c.jam===jam&&c.kode===parseInt(kode));
  document.getElementById('modalTitle').textContent=\`\${kelas}\`;
  const row = (key,val) => \`<div class="flex justify-between items-center py-2.5 text-sm"><span class="text-slate-400 font-medium text-sm">\${key}</span><span class="text-slate-700">\${val}</span></div>\`;
  document.getElementById('modalBody').innerHTML=
    row('Hari / Jam', \`\${HARI_LENGKAP[hari]}, Jam \${jam}\`) +
    row('Waktu', labelUntukHari(hari)[jam]||'') +
    row('Mata Pelajaran', \`<strong>\${g?g.mapel:'?'}</strong>\`) +
    row('Guru', g?g.nama:'Kode '+kode) +
    row('Kode', \`<span class="font-extrabold text-brand-700">\${kode}</span>\`) +
    (isConf ? row('Status', \`<span class="px-2.5 py-0.5 rounded-full bg-red-100 text-red-700 text-xs font-bold">⚠ BENTROK JADWAL</span>\`) : '');
  const modal = document.getElementById('modalBg');
  modal.classList.remove('hidden');
  modal.classList.add('flex');
}
function closeModal(){
  const modal = document.getElementById('modalBg');
  modal.classList.add('hidden');
  modal.classList.remove('flex');
}

function renderWaliPublik() {
  const search = (document.getElementById('searchWaliPublik')?.value || '').toLowerCase();
  let rows = (D.kelas||[]).map(kelas => {
    const kode = (D.wali || {})[kelas] || '';
    const g = kode ? D.guru[kode] : null;
    return { kelas, kode, nama: g ? g.nama : '', mapel: g ? g.mapel : '' };
  });
  if (search) rows = rows.filter(r => r.kelas.toLowerCase().includes(search) || r.nama.toLowerCase().includes(search));
  document.getElementById('waliPublikBody').innerHTML = rows.map(r => \`<tr class="border-t border-slate-100 hover:bg-slate-50/70">
    <td class="px-3 py-2.5 font-semibold text-slate-700">\${r.kelas}</td>
    <td class="px-3 py-2.5">\${r.nama ? \`<div><span class="font-semibold text-slate-800">\${r.nama}</span><span class="text-xs text-slate-400 ml-1">— \${r.mapel}</span></div>\` : '<span class="text-slate-300 text-sm">Belum ditentukan</span>'}</td>
    <td class="px-3 py-2.5 text-center">\${r.kode ? \`<span class="font-extrabold text-brand-700">\${r.kode}</span>\` : '—'}</td>
  </tr>\`).join('') || '<tr><td colspan="3" class="text-center text-slate-400 py-6">Tidak ditemukan</td></tr>';
}

function renderPiketPublik() {
  const cont = document.getElementById('piketPublikContent');
  cont.innerHTML = HARI.map(hari => {
    const slotList = (D.piket || {})[hari] || [];
    const namaHari = HARI_LENGKAP[hari];
    const items = Array.from({length:7}, (_,i) => {
      const kode = slotList[i];
      const g = kode ? D.guru[kode] : null;
      return g
        ? \`<div class="flex items-center gap-2 bg-slate-50 rounded-lg px-3 py-2"><span class="font-extrabold text-brand-700 text-sm w-6">\${kode}</span><span class="text-sm text-slate-700 truncate">\${g.nama}</span></div>\`
        : \`<div class="flex items-center gap-2 bg-slate-50/50 rounded-lg px-3 py-2 border border-dashed border-slate-200"><span class="text-sm text-slate-300">Slot \${i+1} — belum ditentukan</span></div>\`;
    }).join('');
    return \`<div class="bg-white border border-slate-200 rounded-xl p-4 shadow-sm">
      <div class="font-bold text-base text-brand-700 mb-3">📌 \${namaHari}</div>
      <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">\${items}</div>
    </div>\`;
  }).join('');
}

function showTab(id, btn){
  document.querySelectorAll('.tab-section').forEach(s=>{s.classList.add('hidden');s.classList.remove('block');});
  document.querySelectorAll('.nav-btn').forEach(b=>{
    b.classList.remove('text-brand-700','border-brand-600');
    b.classList.add('text-slate-500','border-transparent');
  });
  const target = document.getElementById('tab-'+id);
  target.classList.remove('hidden');
  target.classList.add('block');
  if (btn) {
    btn.classList.remove('text-slate-500','border-transparent');
    btn.classList.add('text-brand-700','border-brand-600');
  }
  if(id==='hari-ini')renderHariIni();
  if(id==='jadwal-guru')renderJadwalGuru();
  if(id==='jadwal-mapel'){ renderJadwalMapel(); }
  if(id==='wali-kelas')renderWaliPublik();
  if(id==='piket-guru')renderPiketPublik();
  if(id==='pengumuman')renderPengumumanPublik();
}

fetchData();
</script></body></html>`;
}
