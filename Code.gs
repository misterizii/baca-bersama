/**
 * ============================================================
 *  BACA BERSAMA — Backend (Google Apps Script)
 *  SD Negeri Geger, Kedungadem - Bojonegoro
 * ============================================================
 *  Fungsi:
 *   - Database  : Google Sheets yang terikat dengan script ini
 *   - File      : Google Drive (folder "LMS_Membaca_Files")
 *   - API       : Web App JSON (dipanggil dari frontend di Vercel via fetch)
 *   - Peran     : admin (kepala sekolah), guru, ortu
 *   - Keamanan  : semua validasi peran & kelas dilakukan DI SINI (server)
 *
 *  CARA PAKAI SINGKAT (detail lengkap ada di README.md):
 *   1. Buat Google Sheet kosong > Extensions > Apps Script
 *   2. Tempel seluruh isi file ini ke Code.gs
 *   3. Jalankan fungsi setupDatabase() satu kali (setujui izin akses)
 *   4. Deploy > New deployment > Web app
 *        Execute as : Me
 *        Who has access : Anyone
 *   5. Salin URL Web App (berakhiran /exec) ke API_URL di index.html
 * ============================================================
 */

// ---------------------------------------------------------------
// KONFIGURASI
// ---------------------------------------------------------------
var DRIVE_FOLDER_NAME = 'LMS_Membaca_Files';
var SESSION_DAYS = 14;                 // masa berlaku sesi login
var MAX_FILE_BYTES = 4 * 1024 * 1024;  // batas lampiran 4 MB
var MAX_LOGIN_FAIL = 5;                // maksimal salah PIN berturut-turut
var LOCKOUT_SECONDS = 600;             // terkunci 10 menit setelah salah PIN
var KELAS_ADMIN = 'SEMUA';             // nilai kolom Kelas untuk akun admin

var SHEETS = {
  Users: ['UserID', 'Username', 'PIN', 'NamaLengkap', 'NamaAnak', 'Role', 'Kelas'],
  MateriBacaan: ['MateriID', 'Kelas', 'Judul', 'Instruksi', 'TeksBacaan', 'FileURL', 'TanggalUpload', 'GuruID'],
  LaporanBaca: ['LaporanID', 'MateriID', 'Kelas', 'OrtuID', 'NamaAnak', 'TingkatKelancaran', 'CatatanOrtu', 'TanggalLapor'],
  TugasSoal: ['TugasID', 'Kelas', 'JudulTugas', 'DataPG_JSON', 'DataEsai_JSON', 'TanggalBuat', 'GuruID'],
  JawabanSiswa: ['JawabanID', 'TugasID', 'Kelas', 'OrtuID', 'NamaAnak', 'JawabanPG_JSON', 'SkorPG', 'JawabanEsai_JSON', 'TanggalKumpul'],
  Diskusi: ['DiskusiID', 'Kelas', 'ParentBubbleID', 'PengirimID', 'NamaPengirim', 'RolePengirim', 'IsiPesan', 'MentionedOrtuIDs', 'Waktu']
};

var ID_PREFIX = {
  Users: 'U', MateriBacaan: 'M', LaporanBaca: 'L',
  TugasSoal: 'T', JawabanSiswa: 'J', Diskusi: 'D'
};

var TINGKAT = ['Masih Mengeja', 'Mulai Lancar', 'Sangat Lancar'];

// ---------------------------------------------------------------
// SETUP DATABASE  (jalankan sekali dari editor Apps Script)
// ---------------------------------------------------------------
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SHEETS).forEach(function (name) {
    var headers = SHEETS[name];
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);

    // Semua sel diformat teks agar isi pesan seperti "=1+1" tidak dianggap rumus
    sheet.getRange(1, 1, sheet.getMaxRows(), Math.max(sheet.getMaxColumns(), headers.length)).setNumberFormat('@');

    var current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    var empty = current.every(function (c) { return String(c) === ''; });
    if (empty) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold').setBackground('#FDE8E8');
    }
  });

  // Hapus sheet bawaan kosong
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1') || ss.getSheetByName('Lembar1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);

  getSecret_();      // buat kunci rahasia sesi
  getFolder_();      // buat folder Drive (sekaligus meminta izin Drive)

  // Akun contoh (hanya jika tabel Users masih kosong)
  var users = ss.getSheetByName('Users');
  if (users.getLastRow() < 2) {
    var samples = [
      ['U001', 'kepsek', '123456', 'Kepala Sekolah', '', 'admin', KELAS_ADMIN],
      ['U002', 'guru1',  '111111', 'Bu Guru Contoh', '', 'guru', '1A'],
      ['U003', 'ortu1',  '222222', 'Bapak Budi', 'Budi', 'ortu', '1A'],
      ['U004', 'ortu2',  '333333', 'Ibu Sari', 'Sari', 'ortu', '2B']
    ];
    samples.forEach(function (s) {
      appendObj_('Users', {
        UserID: s[0], Username: s[1], PIN: hashPin_(s[0], s[2]),
        NamaLengkap: s[3], NamaAnak: s[4], Role: s[5], Kelas: s[6]
      });
    });
  }

  // Contoh materi agar langsung ada yang bisa dilihat saat uji coba
  var materi = ss.getSheetByName('MateriBacaan');
  if (materi.getLastRow() < 2) {
    appendObj_('MateriBacaan', {
      MateriID: 'M001', Kelas: '1A', Judul: 'Ini Budi',
      Instruksi: 'Dampingi anak membaca 2 kali. Minta anak menunjuk tiap kata dengan jari.',
      TeksBacaan: 'Ini Budi.\nBudi anak baik.\nBudi suka membaca buku.\nIbu Budi senang.',
      FileURL: '', TanggalUpload: nowIso_(), GuruID: 'U002'
    });
  }

  Logger.log('Selesai. Akun contoh: kepsek/123456 (admin), guru1/111111 (guru 1A), ortu1/222222 (ortu 1A), ortu2/333333 (ortu 2B). Segera ganti PIN setelah login pertama.');
}

// ---------------------------------------------------------------
// ENTRY POINT WEB APP
// ---------------------------------------------------------------
function doGet() {
  return json_({ ok: true, app: 'Baca Bersama API', time: nowIso_() });
}

function doPost(e) {
  var lock = null;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var action = String(req.action || '');
    var handler = ROUTES[action];
    if (!handler) throw new Error('Aksi tidak dikenal.');

    var user = null;
    if (action !== 'login') user = auth_(req.token);

    // Aksi yang mengubah data diserialkan agar tidak bentrok
    if (WRITE_ACTIONS[action]) {
      lock = LockService.getScriptLock();
      lock.waitLock(25000);
    }
    var data = handler(req, user);
    return json_({ ok: true, data: data });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err), code: err && err.code ? err.code : '' });
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (x) {} }
  }
}

var WRITE_ACTIONS = {
  login: false, changePin: true, saveUser: true, deleteUser: true,
  saveMateri: true, deleteMateri: true, submitLaporan: true,
  saveTugas: true, deleteTugas: true, submitJawaban: true,
  postDiskusi: true, deleteDiskusi: true
};

var ROUTES = {
  login: actLogin_,
  me: actMe_,
  changePin: actChangePin_,
  kelasList: actKelasList_,
  listUsers: actListUsers_,
  saveUser: actSaveUser_,
  deleteUser: actDeleteUser_,
  getMateri: actGetMateri_,
  saveMateri: actSaveMateri_,
  deleteMateri: actDeleteMateri_,
  getRekap: actGetRekap_,
  submitLaporan: actSubmitLaporan_,
  getTugas: actGetTugas_,
  saveTugas: actSaveTugas_,
  deleteTugas: actDeleteTugas_,
  getHasil: actGetHasil_,
  submitJawaban: actSubmitJawaban_,
  getDiskusi: actGetDiskusi_,
  postDiskusi: actPostDiskusi_,
  deleteDiskusi: actDeleteDiskusi_,
  notif: actNotif_
};

// ---------------------------------------------------------------
// UTILITAS UMUM
// ---------------------------------------------------------------
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function nowIso_() {
  return Utilities.formatDate(new Date(), 'Asia/Jakarta', "yyyy-MM-dd'T'HH:mm:ss");
}

function fail_(msg, code) {
  var e = new Error(msg);
  if (code) e.code = code;
  throw e;
}

function str_(v, max) {
  var s = (v === null || v === undefined) ? '' : String(v);
  s = s.replace(/\r\n/g, '\n').replace(/^\s+|\s+$/g, '');
  if (max && s.length > max) s = s.substring(0, max);
  return s;
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name) {
  var s = ss_().getSheetByName(name);
  if (!s) fail_('Database belum disiapkan. Jalankan setupDatabase() di Apps Script.');
  return s;
}

function readAll_(name) {
  var sheet = sheet_(name);
  var headers = SHEETS[name];
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var values = sheet.getRange(2, 1, last - 1, headers.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (String(row[0]) === '') continue;
    var o = { _row: i + 2 };
    for (var c = 0; c < headers.length; c++) o[headers[c]] = String(row[c]);
    out.push(o);
  }
  return out;
}

function appendObj_(name, obj) {
  var headers = SHEETS[name];
  var sheet = sheet_(name);
  var row = headers.map(function (h) { return obj[h] === undefined ? '' : String(obj[h]); });
  var next = sheet.getLastRow() + 1;
  sheet.getRange(next, 1, 1, headers.length).setNumberFormat('@').setValues([row]);
}

function updateObj_(name, rowIndex, obj) {
  var headers = SHEETS[name];
  var sheet = sheet_(name);
  var row = headers.map(function (h) { return obj[h] === undefined ? '' : String(obj[h]); });
  sheet.getRange(rowIndex, 1, 1, headers.length).setNumberFormat('@').setValues([row]);
}

function deleteRows_(name, rowIndexes) {
  var sheet = sheet_(name);
  rowIndexes.slice().sort(function (a, b) { return b - a; }).forEach(function (r) {
    sheet.deleteRow(r);
  });
}

function nextId_(name) {
  var prefix = ID_PREFIX[name];
  var rows = readAll_(name);
  var max = 0;
  rows.forEach(function (r) {
    var id = r[SHEETS[name][0]];
    var n = parseInt(id.substring(prefix.length), 10);
    if (!isNaN(n) && n > max) max = n;
  });
  var num = String(max + 1);
  while (num.length < 3) num = '0' + num;
  return prefix + num;
}

function parseJson_(s, fallback) {
  try { return s ? JSON.parse(s) : fallback; } catch (e) { return fallback; }
}

// ---------------------------------------------------------------
// KEAMANAN: PIN, TOKEN, SESI
// ---------------------------------------------------------------
function getSecret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('APP_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('APP_SECRET', s);
  }
  return s;
}

function toHex_(bytes) {
  return bytes.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function hashPin_(userId, pin) {
  var raw = getSecret_() + '|' + userId + '|' + pin;
  return toHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8));
}

function isHash_(s) { return /^[a-f0-9]{64}$/.test(s); }

function checkPin_(u, pin) {
  var stored = String(u.PIN);
  if (isHash_(stored)) return stored === hashPin_(u.UserID, pin);
  // PIN yang diketik manual di sheet (belum di-hash): cocokkan lalu upgrade otomatis
  if (stored !== '' && stored === String(pin)) {
    u.PIN = hashPin_(u.UserID, pin);
    updateObj_('Users', u._row, u);
    return true;
  }
  return false;
}

function sign_(payload) {
  var sig = Utilities.computeHmacSha256Signature(payload, getSecret_());
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '');
}

function makeToken_(userId) {
  var exp = Date.now() + SESSION_DAYS * 24 * 3600 * 1000;
  var payload = userId + '.' + exp;
  return payload + '.' + sign_(payload);
}

function auth_(token) {
  token = String(token || '');
  var parts = token.split('.');
  if (parts.length !== 3) fail_('Sesi berakhir. Silakan masuk lagi.', 'SESSION');
  var payload = parts[0] + '.' + parts[1];
  if (sign_(payload) !== parts[2]) fail_('Sesi tidak valid. Silakan masuk lagi.', 'SESSION');
  if (Number(parts[1]) < Date.now()) fail_('Sesi berakhir. Silakan masuk lagi.', 'SESSION');
  var user = findUserById_(parts[0]);
  if (!user) fail_('Akun tidak ditemukan. Silakan masuk lagi.', 'SESSION');
  return user;
}

function findUserById_(id) {
  var rows = readAll_('Users');
  for (var i = 0; i < rows.length; i++) if (rows[i].UserID === id) return rows[i];
  return null;
}

function publicUser_(u) {
  return {
    id: u.UserID, username: u.Username, nama: u.NamaLengkap,
    namaAnak: u.NamaAnak, role: u.Role, kelas: u.Kelas
  };
}

function requireRole_(user, roles) {
  if (roles.indexOf(user.Role) < 0) fail_('Anda tidak memiliki akses untuk tindakan ini.');
}

/**
 * ISOLASI KELAS — satu-satunya pintu penentu kelas yang boleh diakses.
 *  - guru/ortu : selalu kelasnya sendiri; permintaan kelas lain ditolak
 *  - admin     : boleh kelas apa pun; '*' berarti semua kelas (hanya bila allowAll)
 */
function scope_(user, requested, allowAll) {
  requested = str_(requested, 40);
  if (user.Role === 'admin') {
    if (requested === '*' || requested === '') {
      if (allowAll) return '*';
      fail_('Pilih kelas terlebih dahulu.');
    }
    return requested;
  }
  if (requested !== '' && requested !== '*' && requested !== user.Kelas) fail_('Akses ke kelas lain tidak diizinkan.');
  return user.Kelas;
}

function inScope_(rowKelas, sc) { return sc === '*' || rowKelas === sc; }

function assertKelasAccess_(user, kelas) {
  if (user.Role === 'admin') return;
  if (user.Kelas !== kelas) fail_('Akses ke kelas lain tidak diizinkan.');
}

function ortuInKelas_(kelas) {
  return readAll_('Users').filter(function (u) { return u.Role === 'ortu' && u.Kelas === kelas; });
}

// ---------------------------------------------------------------
// AKSI: LOGIN, PROFIL
// ---------------------------------------------------------------
function actLogin_(req) {
  var username = str_(req.username, 60).toLowerCase();
  var pin = str_(req.pin, 20);
  if (!username || !pin) fail_('Isi username dan PIN.');

  var cache = CacheService.getScriptCache();
  var key = 'fail_' + username.replace(/[^a-z0-9_.-]/g, '');
  var fails = parseInt(cache.get(key) || '0', 10);
  if (fails >= MAX_LOGIN_FAIL) fail_('Terlalu banyak percobaan salah. Coba lagi sekitar 10 menit.');

  var rows = readAll_('Users');
  var user = null;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].Username.toLowerCase() === username) { user = rows[i]; break; }
  }
  if (!user || !checkPin_(user, pin)) {
    cache.put(key, String(fails + 1), LOCKOUT_SECONDS);
    fail_('Username atau PIN salah.');
  }
  cache.remove(key);
  return { token: makeToken_(user.UserID), user: publicUser_(user) };
}

function actMe_(req, user) {
  return { user: publicUser_(user) };
}

function actChangePin_(req, user) {
  var oldPin = str_(req.oldPin, 20);
  var newPin = str_(req.newPin, 20);
  if (!/^\d{4,8}$/.test(newPin)) fail_('PIN baru harus 4 sampai 8 angka.');
  if (!checkPin_(user, oldPin)) fail_('PIN lama salah.');
  user.PIN = hashPin_(user.UserID, newPin);
  updateObj_('Users', user._row, user);
  return { done: true };
}

function actKelasList_(req, user) {
  requireRole_(user, ['admin']);
  var seen = {};
  readAll_('Users').forEach(function (u) {
    if (u.Role !== 'admin' && u.Kelas) seen[u.Kelas] = true;
  });
  return { kelas: Object.keys(seen).sort() };
}

// ---------------------------------------------------------------
// AKSI: MANAJEMEN PENGGUNA (admin / kepala sekolah)
// ---------------------------------------------------------------
function actListUsers_(req, user) {
  requireRole_(user, ['admin']);
  var order = { admin: 0, guru: 1, ortu: 2 };
  var list = readAll_('Users').map(function (u) {
    return { id: u.UserID, username: u.Username, nama: u.NamaLengkap, namaAnak: u.NamaAnak, role: u.Role, kelas: u.Kelas };
  });
  list.sort(function (a, b) {
    if (order[a.role] !== order[b.role]) return order[a.role] - order[b.role];
    if (a.kelas !== b.kelas) return a.kelas < b.kelas ? -1 : 1;
    return a.nama.toLowerCase() < b.nama.toLowerCase() ? -1 : 1;
  });
  return { users: list };
}

function actSaveUser_(req, user) {
  requireRole_(user, ['admin']);
  var id = str_(req.id, 20);
  var username = str_(req.username, 40);
  var pin = str_(req.pin, 20);
  var nama = str_(req.nama, 80);
  var namaAnak = str_(req.namaAnak, 80);
  var role = str_(req.role, 10);
  var kelas = str_(req.kelas, 20).toUpperCase();

  if (['admin', 'guru', 'ortu'].indexOf(role) < 0) fail_('Peran tidak valid.');
  if (!/^[A-Za-z0-9_.-]{3,40}$/.test(username)) fail_('Username 3-40 karakter: huruf, angka, titik, strip, atau garis bawah.');
  if (!nama) fail_('Nama lengkap wajib diisi.');
  if (role === 'admin') { kelas = KELAS_ADMIN; namaAnak = ''; }
  if (role === 'guru') { namaAnak = ''; if (!kelas) fail_('Kelas guru wajib diisi.'); }
  if (role === 'ortu') {
    if (!kelas) fail_('Kelas wajib diisi.');
    if (!namaAnak) fail_('Nama anak wajib diisi untuk orang tua.');
  }
  if (kelas === '*' || (role !== 'admin' && kelas === KELAS_ADMIN)) fail_('Nama kelas tidak valid.');

  var users = readAll_('Users');
  users.forEach(function (u) {
    if (u.Username.toLowerCase() === username.toLowerCase() && u.UserID !== id) fail_('Username sudah dipakai.');
  });

  if (id) {
    var target = null;
    users.forEach(function (u) { if (u.UserID === id) target = u; });
    if (!target) fail_('Pengguna tidak ditemukan.');
    if (target.Role === 'admin' && role !== 'admin') {
      var admins = users.filter(function (u) { return u.Role === 'admin'; });
      if (admins.length <= 1) fail_('Harus ada minimal satu admin.');
    }
    if (target.UserID === user.UserID && role !== 'admin') fail_('Anda tidak dapat mengubah peran akun Anda sendiri.');
    if (pin) {
      if (!/^\d{4,8}$/.test(pin)) fail_('PIN harus 4 sampai 8 angka.');
      target.PIN = hashPin_(target.UserID, pin);
    }
    target.Username = username; target.NamaLengkap = nama; target.NamaAnak = namaAnak;
    target.Role = role; target.Kelas = kelas;
    updateObj_('Users', target._row, target);
    return { id: target.UserID };
  }

  if (!/^\d{4,8}$/.test(pin)) fail_('PIN wajib diisi: 4 sampai 8 angka.');
  var newId = nextId_('Users');
  appendObj_('Users', {
    UserID: newId, Username: username, PIN: hashPin_(newId, pin),
    NamaLengkap: nama, NamaAnak: namaAnak, Role: role, Kelas: kelas
  });
  return { id: newId };
}

function actDeleteUser_(req, user) {
  requireRole_(user, ['admin']);
  var id = str_(req.id, 20);
  if (id === user.UserID) fail_('Anda tidak dapat menghapus akun Anda sendiri.');
  var users = readAll_('Users');
  var target = null;
  users.forEach(function (u) { if (u.UserID === id) target = u; });
  if (!target) fail_('Pengguna tidak ditemukan.');
  if (target.Role === 'admin' && users.filter(function (u) { return u.Role === 'admin'; }).length <= 1) {
    fail_('Harus ada minimal satu admin.');
  }
  deleteRows_('Users', [target._row]);
  return { done: true };
}

// ---------------------------------------------------------------
// AKSI: MATERI BACAAN
// ---------------------------------------------------------------
function actGetMateri_(req, user) {
  var sc = scope_(user, req.kelas, true);
  var rows = readAll_('MateriBacaan').filter(function (m) { return inScope_(m.Kelas, sc); });
  rows.sort(function (a, b) { return a.TanggalUpload < b.TanggalUpload ? 1 : -1; });

  var laporan = readAll_('LaporanBaca').filter(function (l) { return inScope_(l.Kelas, sc); });
  var out;

  if (user.Role === 'ortu') {
    out = rows.map(function (m) {
      var mine = null;
      laporan.forEach(function (l) { if (l.MateriID === m.MateriID && l.OrtuID === user.UserID) mine = l; });
      return {
        id: m.MateriID, kelas: m.Kelas, judul: m.Judul, instruksi: m.Instruksi,
        teks: m.TeksBacaan, fileUrl: m.FileURL, tanggal: m.TanggalUpload,
        laporanSaya: mine ? { tingkat: mine.TingkatKelancaran, catatan: mine.CatatanOrtu, tanggal: mine.TanggalLapor } : null
      };
    });
    return { materi: out };
  }

  var users = readAll_('Users');
  out = rows.map(function (m) {
    var jumlahOrtu = users.filter(function (u) { return u.Role === 'ortu' && u.Kelas === m.Kelas; }).length;
    var jumlahLapor = laporan.filter(function (l) { return l.MateriID === m.MateriID; }).length;
    return {
      id: m.MateriID, kelas: m.Kelas, judul: m.Judul, instruksi: m.Instruksi,
      teks: m.TeksBacaan, fileUrl: m.FileURL, tanggal: m.TanggalUpload,
      jumlahOrtu: jumlahOrtu, jumlahLapor: jumlahLapor
    };
  });
  return { materi: out };
}

function getFolder_() {
  var it = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function saveFile_(file) {
  var allowed = { 'application/pdf': 'pdf', 'image/jpeg': 'jpg', 'image/png': 'png' };
  var mime = str_(file.mime, 60);
  if (!allowed[mime]) fail_('Lampiran harus PDF, JPG, atau PNG.');
  var b64 = String(file.data || '');
  if (!b64) fail_('Lampiran kosong.');
  var bytes = Utilities.base64Decode(b64);
  if (bytes.length > MAX_FILE_BYTES) fail_('Ukuran lampiran maksimal 4 MB.');
  var base = str_(file.name, 80).replace(/[^A-Za-z0-9._ -]/g, '_').replace(/\.[A-Za-z0-9]+$/, '');
  if (!base) base = 'lampiran';
  var blob = Utilities.newBlob(bytes, mime, base + '_' + Date.now() + '.' + allowed[mime]);
  var f = getFolder_().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f.getUrl();
}

function trashFileByUrl_(url) {
  try {
    var m = /\/d\/([A-Za-z0-9_-]+)/.exec(url) || /[?&]id=([A-Za-z0-9_-]+)/.exec(url);
    if (m) DriveApp.getFileById(m[1]).setTrashed(true);
  } catch (e) { /* abaikan jika file sudah hilang */ }
}

function actSaveMateri_(req, user) {
  requireRole_(user, ['guru', 'admin']);
  var kelas = scope_(user, req.kelas, false);
  var judul = str_(req.judul, 150);
  var instruksi = str_(req.instruksi, 1000);
  var teks = str_(req.teks, 8000);
  if (!judul) fail_('Judul wajib diisi.');
  if (!teks && !(req.file && req.file.data)) fail_('Isi teks bacaan atau unggah lampiran.');

  var url = '';
  if (req.file && req.file.data) url = saveFile_(req.file);

  var id = nextId_('MateriBacaan');
  appendObj_('MateriBacaan', {
    MateriID: id, Kelas: kelas, Judul: judul, Instruksi: instruksi, TeksBacaan: teks,
    FileURL: url, TanggalUpload: nowIso_(), GuruID: user.UserID
  });
  return { id: id };
}

function actDeleteMateri_(req, user) {
  requireRole_(user, ['guru', 'admin']);
  var id = str_(req.id, 20);
  var m = null;
  readAll_('MateriBacaan').forEach(function (r) { if (r.MateriID === id) m = r; });
  if (!m) fail_('Materi tidak ditemukan.');
  assertKelasAccess_(user, m.Kelas);

  var laporanRows = readAll_('LaporanBaca').filter(function (l) { return l.MateriID === id; }).map(function (l) { return l._row; });
  if (laporanRows.length) deleteRows_('LaporanBaca', laporanRows);
  deleteRows_('MateriBacaan', [m._row]);
  if (m.FileURL) trashFileByUrl_(m.FileURL);
  return { done: true };
}

function actGetRekap_(req, user) {
  requireRole_(user, ['guru', 'admin']);
  var id = str_(req.materiId, 20);
  var m = null;
  readAll_('MateriBacaan').forEach(function (r) { if (r.MateriID === id) m = r; });
  if (!m) fail_('Materi tidak ditemukan.');
  assertKelasAccess_(user, m.Kelas);

  var laporan = readAll_('LaporanBaca').filter(function (l) { return l.MateriID === id; });
  var list = ortuInKelas_(m.Kelas).map(function (o) {
    var l = null;
    laporan.forEach(function (x) { if (x.OrtuID === o.UserID) l = x; });
    return {
      ortuId: o.UserID, namaOrtu: o.NamaLengkap, namaAnak: o.NamaAnak,
      tingkat: l ? l.TingkatKelancaran : '', catatan: l ? l.CatatanOrtu : '', tanggal: l ? l.TanggalLapor : ''
    };
  });
  list.sort(function (a, b) { return a.namaAnak.toLowerCase() < b.namaAnak.toLowerCase() ? -1 : 1; });
  return { materiId: id, judul: m.Judul, kelas: m.Kelas, daftar: list };
}

function actSubmitLaporan_(req, user) {
  requireRole_(user, ['ortu']);
  var id = str_(req.materiId, 20);
  var tingkat = str_(req.tingkat, 30);
  var catatan = str_(req.catatan, 1500);
  if (TINGKAT.indexOf(tingkat) < 0) fail_('Pilih tingkat kelancaran membaca.');

  var m = null;
  readAll_('MateriBacaan').forEach(function (r) { if (r.MateriID === id) m = r; });
  if (!m || m.Kelas !== user.Kelas) fail_('Materi tidak ditemukan.');

  var dup = readAll_('LaporanBaca').filter(function (l) { return l.MateriID === id && l.OrtuID === user.UserID; });
  if (dup.length) fail_('Laporan untuk materi ini sudah dikirim.');

  appendObj_('LaporanBaca', {
    LaporanID: nextId_('LaporanBaca'), MateriID: id, Kelas: user.Kelas, OrtuID: user.UserID,
    NamaAnak: user.NamaAnak, TingkatKelancaran: tingkat, CatatanOrtu: catatan, TanggalLapor: nowIso_()
  });
  return { done: true };
}

// ---------------------------------------------------------------
// AKSI: TUGAS SOAL
// ---------------------------------------------------------------
function actGetTugas_(req, user) {
  var sc = scope_(user, req.kelas, true);
  var rows = readAll_('TugasSoal').filter(function (t) { return inScope_(t.Kelas, sc); });
  rows.sort(function (a, b) { return a.TanggalBuat < b.TanggalBuat ? 1 : -1; });
  var jawaban = readAll_('JawabanSiswa').filter(function (j) { return inScope_(j.Kelas, sc); });
  var users = readAll_('Users');

  var out = rows.map(function (t) {
    var pg = parseJson_(t.DataPG_JSON, []);
    var esai = parseJson_(t.DataEsai_JSON, []);

    if (user.Role === 'ortu') {
      // Kunci jawaban TIDAK dikirim ke orang tua
      var mine = null;
      jawaban.forEach(function (j) { if (j.TugasID === t.TugasID && j.OrtuID === user.UserID) mine = j; });
      var pgInfo = mine ? parseJson_(mine.JawabanPG_JSON, {}) : null;
      return {
        id: t.TugasID, kelas: t.Kelas, judul: t.JudulTugas, tanggal: t.TanggalBuat,
        pg: pg.map(function (p) { return { q: p.q, a: p.a, b: p.b, c: p.c, d: p.d }; }),
        esai: esai.map(function (e) { return { q: e.q }; }),
        hasilSaya: mine ? {
          skor: mine.SkorPG, benar: pgInfo ? pgInfo.benar : 0, total: pgInfo ? pgInfo.total : 0, tanggal: mine.TanggalKumpul
        } : null
      };
    }

    var jumlahOrtu = users.filter(function (u) { return u.Role === 'ortu' && u.Kelas === t.Kelas; }).length;
    var jumlahKumpul = jawaban.filter(function (j) { return j.TugasID === t.TugasID; }).length;
    return {
      id: t.TugasID, kelas: t.Kelas, judul: t.JudulTugas, tanggal: t.TanggalBuat,
      jumlahPG: pg.length, jumlahEsai: esai.length, jumlahOrtu: jumlahOrtu, jumlahKumpul: jumlahKumpul
    };
  });
  return { tugas: out };
}

function actSaveTugas_(req, user) {
  requireRole_(user, ['guru', 'admin']);
  var kelas = scope_(user, req.kelas, false);
  var judul = str_(req.judul, 150);
  if (!judul) fail_('Judul tugas wajib diisi.');

  var pgIn = Array.isArray(req.pg) ? req.pg : [];
  var esaiIn = Array.isArray(req.esai) ? req.esai : [];
  if (pgIn.length > 10) fail_('Maksimal 10 soal pilihan ganda.');
  if (esaiIn.length > 5) fail_('Maksimal 5 soal esai.');

  var pg = [];
  pgIn.forEach(function (p, idx) {
    var q = str_(p && p.q, 500);
    var a = str_(p && p.a, 200), b = str_(p && p.b, 200), c = str_(p && p.c, 200), d = str_(p && p.d, 200);
    var k = str_(p && p.k, 1).toUpperCase();
    var any = q || a || b || c || d || k;
    if (!any) return; // slot kosong dilewati
    if (!q || !a || !b || !c || !d) fail_('Soal PG nomor ' + (idx + 1) + ' belum lengkap (pertanyaan dan opsi A-D).');
    if (['A', 'B', 'C', 'D'].indexOf(k) < 0) fail_('Soal PG nomor ' + (idx + 1) + ' belum punya kunci jawaban.');
    pg.push({ q: q, a: a, b: b, c: c, d: d, k: k });
  });

  var esai = [];
  esaiIn.forEach(function (e) {
    var q = str_(e && e.q, 500);
    if (q) esai.push({ q: q });
  });

  if (!pg.length && !esai.length) fail_('Isi minimal satu soal.');

  var id = nextId_('TugasSoal');
  appendObj_('TugasSoal', {
    TugasID: id, Kelas: kelas, JudulTugas: judul,
    DataPG_JSON: JSON.stringify(pg), DataEsai_JSON: JSON.stringify(esai),
    TanggalBuat: nowIso_(), GuruID: user.UserID
  });
  return { id: id, jumlahPG: pg.length, jumlahEsai: esai.length };
}

function actDeleteTugas_(req, user) {
  requireRole_(user, ['guru', 'admin']);
  var id = str_(req.id, 20);
  var t = null;
  readAll_('TugasSoal').forEach(function (r) { if (r.TugasID === id) t = r; });
  if (!t) fail_('Tugas tidak ditemukan.');
  assertKelasAccess_(user, t.Kelas);

  var rows = readAll_('JawabanSiswa').filter(function (j) { return j.TugasID === id; }).map(function (j) { return j._row; });
  if (rows.length) deleteRows_('JawabanSiswa', rows);
  deleteRows_('TugasSoal', [t._row]);
  return { done: true };
}

function actGetHasil_(req, user) {
  requireRole_(user, ['guru', 'admin']);
  var id = str_(req.tugasId, 20);
  var t = null;
  readAll_('TugasSoal').forEach(function (r) { if (r.TugasID === id) t = r; });
  if (!t) fail_('Tugas tidak ditemukan.');
  assertKelasAccess_(user, t.Kelas);

  var pg = parseJson_(t.DataPG_JSON, []);
  var esai = parseJson_(t.DataEsai_JSON, []);
  var jawaban = readAll_('JawabanSiswa').filter(function (j) { return j.TugasID === id; });

  var daftar = ortuInKelas_(t.Kelas).map(function (o) {
    var j = null;
    jawaban.forEach(function (x) { if (x.OrtuID === o.UserID) j = x; });
    if (!j) return { ortuId: o.UserID, namaAnak: o.NamaAnak, namaOrtu: o.NamaLengkap, terkirim: false };
    var pgInfo = parseJson_(j.JawabanPG_JSON, {});
    var esaiJawab = parseJson_(j.JawabanEsai_JSON, []);
    return {
      ortuId: o.UserID, namaAnak: o.NamaAnak, namaOrtu: o.NamaLengkap, terkirim: true,
      skor: j.SkorPG, benar: pgInfo.benar || 0, total: pgInfo.total || 0, tanggal: j.TanggalKumpul,
      esai: esai.map(function (e, i) { return { q: e.q, a: esaiJawab[i] || '' }; })
    };
  });
  daftar.sort(function (a, b) { return a.namaAnak.toLowerCase() < b.namaAnak.toLowerCase() ? -1 : 1; });
  return { tugasId: id, judul: t.JudulTugas, kelas: t.Kelas, jumlahPG: pg.length, jumlahEsai: esai.length, daftar: daftar };
}

function actSubmitJawaban_(req, user) {
  requireRole_(user, ['ortu']);
  var id = str_(req.tugasId, 20);
  var t = null;
  readAll_('TugasSoal').forEach(function (r) { if (r.TugasID === id) t = r; });
  if (!t || t.Kelas !== user.Kelas) fail_('Tugas tidak ditemukan.');

  var dup = readAll_('JawabanSiswa').filter(function (j) { return j.TugasID === id && j.OrtuID === user.UserID; });
  if (dup.length) fail_('Tugas ini sudah dikumpulkan.');

  var pg = parseJson_(t.DataPG_JSON, []);
  var esai = parseJson_(t.DataEsai_JSON, []);
  var pgIn = Array.isArray(req.pg) ? req.pg : [];
  var esaiIn = Array.isArray(req.esai) ? req.esai : [];

  var jawabPG = [];
  var benar = 0;
  for (var i = 0; i < pg.length; i++) {
    var a = str_(pgIn[i], 1).toUpperCase();
    if (['A', 'B', 'C', 'D'].indexOf(a) < 0) fail_('Soal PG nomor ' + (i + 1) + ' belum dijawab.');
    jawabPG.push(a);
    if (a === pg[i].k) benar++;
  }
  var jawabEsai = [];
  for (var j = 0; j < esai.length; j++) {
    var e = str_(esaiIn[j], 2000);
    if (!e) fail_('Soal esai nomor ' + (j + 1) + ' belum dijawab.');
    jawabEsai.push(e);
  }

  var skor = pg.length ? Math.round(benar / pg.length * 100) : '';
  appendObj_('JawabanSiswa', {
    JawabanID: nextId_('JawabanSiswa'), TugasID: id, Kelas: user.Kelas, OrtuID: user.UserID,
    NamaAnak: user.NamaAnak,
    JawabanPG_JSON: JSON.stringify({ a: jawabPG, benar: benar, total: pg.length }),
    SkorPG: skor, JawabanEsai_JSON: JSON.stringify(jawabEsai), TanggalKumpul: nowIso_()
  });
  return { skor: skor, benar: benar, total: pg.length, adaEsai: esai.length > 0 };
}

// ---------------------------------------------------------------
// AKSI: DISKUSI KELAS
// ---------------------------------------------------------------
function splitIds_(s) {
  return String(s || '').split(',').map(function (x) { return x.replace(/\s/g, ''); }).filter(function (x) { return x; });
}

function actGetDiskusi_(req, user) {
  var kelas = scope_(user, req.kelas, false);
  var rows = readAll_('Diskusi').filter(function (d) { return d.Kelas === kelas; });
  rows.sort(function (a, b) { return a.Waktu < b.Waktu ? -1 : 1; });

  var ortu = ortuInKelas_(kelas);
  var ortuMap = {};
  ortu.forEach(function (o) { ortuMap[o.UserID] = o.NamaLengkap + ' (ortu ' + o.NamaAnak + ')'; });

  var topics = [];
  var byId = {};
  rows.forEach(function (d) {
    if (!d.ParentBubbleID) {
      var t = {
        id: d.DiskusiID, pengirimId: d.PengirimID, nama: d.NamaPengirim, role: d.RolePengirim,
        isi: d.IsiPesan, mentions: splitIds_(d.MentionedOrtuIDs), waktu: d.Waktu, balasan: []
      };
      byId[d.DiskusiID] = t;
      topics.push(t);
    }
  });
  rows.forEach(function (d) {
    if (d.ParentBubbleID && byId[d.ParentBubbleID]) {
      byId[d.ParentBubbleID].balasan.push({
        id: d.DiskusiID, pengirimId: d.PengirimID, nama: d.NamaPengirim, role: d.RolePengirim,
        isi: d.IsiPesan, waktu: d.Waktu
      });
    }
  });
  topics.sort(function (a, b) { return a.waktu < b.waktu ? 1 : -1; });

  topics.forEach(function (t) {
    // Aturan kunci: hanya ortu yang di-mention yang boleh membalas; guru/admin selalu boleh
    t.bisaBalas = user.Role !== 'ortu' || t.mentions.indexOf(user.UserID) >= 0;
  });

  var result = { kelas: kelas, topik: topics, ortuMap: ortuMap };
  if (user.Role !== 'ortu') {
    result.ortuList = ortu.map(function (o) { return { id: o.UserID, nama: o.NamaLengkap, namaAnak: o.NamaAnak }; });
  }
  return result;
}

function actPostDiskusi_(req, user) {
  var kelas = scope_(user, req.kelas, false);
  var isi = str_(req.isi, 1500);
  if (!isi) fail_('Pesan tidak boleh kosong.');
  var parentId = str_(req.parentId, 20);

  var all = readAll_('Diskusi');
  var parent = null;
  if (parentId) {
    all.forEach(function (d) { if (d.DiskusiID === parentId) parent = d; });
    if (!parent || parent.ParentBubbleID || parent.Kelas !== kelas) fail_('Pesan yang dibalas tidak ditemukan.');
  }

  var mentions = [];
  if (user.Role === 'ortu') {
    if (!parent) fail_('Orang tua hanya dapat membalas pesan yang menyebut Anda.');
    if (splitIds_(parent.MentionedOrtuIDs).indexOf(user.UserID) < 0) {
      fail_('Anda hanya dapat membalas pesan yang menyebut Anda.');
    }
  } else if (!parent) {
    // Topik baru dari guru/admin: validasi daftar mention harus orang tua di kelas ini
    var valid = {};
    ortuInKelas_(kelas).forEach(function (o) { valid[o.UserID] = true; });
    (Array.isArray(req.mentions) ? req.mentions : []).forEach(function (m) {
      m = str_(m, 20);
      if (valid[m] && mentions.indexOf(m) < 0) mentions.push(m);
    });
  }

  var sender = user.NamaLengkap;
  if (user.Role === 'ortu') sender += ' (ortu ' + user.NamaAnak + ')';
  var id = nextId_('Diskusi');
  appendObj_('Diskusi', {
    DiskusiID: id, Kelas: kelas, ParentBubbleID: parentId, PengirimID: user.UserID,
    NamaPengirim: sender, RolePengirim: user.Role, IsiPesan: isi,
    MentionedOrtuIDs: mentions.join(','), Waktu: nowIso_()
  });
  return { id: id };
}

function actDeleteDiskusi_(req, user) {
  requireRole_(user, ['guru', 'admin']);
  var id = str_(req.id, 20);
  var all = readAll_('Diskusi');
  var target = null;
  all.forEach(function (d) { if (d.DiskusiID === id) target = d; });
  if (!target) fail_('Pesan tidak ditemukan.');
  assertKelasAccess_(user, target.Kelas);

  var rows = [target._row];
  if (!target.ParentBubbleID) {
    all.forEach(function (d) { if (d.ParentBubbleID === id) rows.push(d._row); });
  }
  deleteRows_('Diskusi', rows);
  return { done: true };
}

/** Jumlah topik yang menyebut orang tua dan belum dibalas olehnya (untuk banner). */
function actNotif_(req, user) {
  if (user.Role !== 'ortu') return { count: 0 };
  var rows = readAll_('Diskusi').filter(function (d) { return d.Kelas === user.Kelas; });
  var replied = {};
  rows.forEach(function (d) {
    if (d.ParentBubbleID && d.PengirimID === user.UserID) replied[d.ParentBubbleID] = true;
  });
  var count = 0;
  rows.forEach(function (d) {
    if (!d.ParentBubbleID && splitIds_(d.MentionedOrtuIDs).indexOf(user.UserID) >= 0 && !replied[d.DiskusiID]) count++;
  });
  return { count: count };
}
