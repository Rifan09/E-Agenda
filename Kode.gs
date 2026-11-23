/**
 * E-AGENDA SURAT RUTAN BARABAI - BACKEND CONTROLLER
 * Fitur: Penomoran Urut Terpisah Per Klasifikasi + Asal Surat & Sifat Surat
 * E-AGENDA SURAT RUTAN BARABAI — backend + approval edit
 */

function doGet(e) {
  var page = e && e.parameter && e.parameter.page;
  var isAdmin = page === 'admin';

  var template = HtmlService.createTemplateFromFile(
    isAdmin ? 'Admin' : 'Index'
  );

  // URL dasar deployment, tanpa parameter ?page=admin
  template.baseUrl = ScriptApp.getService().getUrl();

  return template.evaluate()
    .setTitle(
      isAdmin
        ? 'Admin E-Agenda Surat'
        : 'E-Agenda Surat Rutan Barabai'
    )
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function initDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ==========================================
  // 1. Setup Sheet: Database_Agenda
  // ==========================================
  var a = ss.getSheetByName('Database_Agenda');
  if (!a) {
    a = ss.insertSheet('Database_Agenda');
    a.appendRow([
      'No Urut Klas',
      'Tanggal',
      'Nomor Surat',
      'Klasifikasi',
      'Asal Surat',
      'Sifat_Surat',
      'Tujuan',
      'Perihal',
      'Nama Petugas',
      'Format Naskah'
    ]);

    a.getRange('A1:J1')
      .setFontWeight('bold')
      .setBackground('#1e3a8a')
      .setFontColor('#ffffff');
  } else {
    // Tambahkan kolom Format Naskah jika belum tersedia.
    var lastColumn = a.getLastColumn();
    var header = a.getRange(1, 1, 1, lastColumn).getValues()[0];

    if (header.indexOf('Format Naskah') === -1) {
      var newColumn = lastColumn + 1;

      a.getRange(1, newColumn).setValue('Format Naskah');
      a.getRange(1, newColumn)
        .setFontWeight('bold')
        .setBackground('#1e3a8a')
        .setFontColor('#ffffff');
    }
  }

  // ==========================================
  // 2. Setup Sheet: Klasifikasi
  // ==========================================
  var k = ss.getSheetByName('Klasifikasi');
  if (!k) {
    k = ss.insertSheet('Klasifikasi');
    k.appendRow(['Kode', 'Keterangan']);
    k.appendRow(['OT.02.01', 'ORGANISASI DAN TATA LAKSANA']);
    k.appendRow(['UM.01.01', 'URUSAN RUMAH TANGGA DAN PERLENGKAPAN']);

    k.getRange('A1:B1')
      .setFontWeight('bold')
      .setBackground('#1e3a8a')
      .setFontColor('#ffffff');
  }

  // ==========================================
  // 3. Setup Sheet: Asal_Surat
  // ==========================================
  var asal = ss.getSheetByName('Asal_Surat');
  if (!asal) {
    asal = ss.insertSheet('Asal_Surat');
    asal.appendRow(['Kode/Singkatan', 'Nama Instansi / Pengirim']);
    asal.appendRow(['LO-LA', 'PENGELOLAAN RUTAN']);
    asal.appendRow(['PELTAH', 'PELAYANAN TAHANAN RUTAN']);
    asal.appendRow(['KPR', 'KESATUAN PENGAMANAN RUTAN']);
    asal.appendRow(['DW', 'DHARMA WANITA']);
    asal.appendRow(['DLL', 'DAN LAINYA']);

    asal.getRange('A1:B1')
      .setFontWeight('bold')
      .setBackground('#1e3a8a')
      .setFontColor('#ffffff');
  }

  // ==========================================
  // 4. Setup Sheet: Sifat_Surat
  // ==========================================
  var sifat = ss.getSheetByName('Sifat_Surat');
  if (!sifat) {
    sifat = ss.insertSheet('Sifat_Surat');
    sifat.appendRow(['Sifat', 'Keterangan']);

    [
      ['Biasa', 'B'],
      ['Terbatas', 'T'],
      ['Rahasia', 'R'],
      ['Sangat Rahasia', 'SR']
    ].forEach(function(r) {
      sifat.appendRow(r);
    });

    sifat.getRange('A1:B1')
      .setFontWeight('bold')
      .setBackground('#1e3a8a')
      .setFontColor('#ffffff');
  }

  ensureWorkflowSheets_();
}


function ensureWorkflowSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var r = ss.getSheetByName('Edit_Requests');
  if (!r) {
    r = ss.insertSheet('Edit_Requests');
    r.appendRow(['Request ID', 'Nomor Surat', 'Status', 'Waktu Permintaan', 'Pemohon', 'Waktu Keputusan', 'Admin']);
    r.getRange('A1:G1').setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  }
  
  var l = ss.getSheetByName('Audit_Edit');
  if (!l) {
    l = ss.insertSheet('Audit_Edit');
    l.appendRow(['Waktu', 'Request ID', 'Nomor Surat', 'Admin', 'Tujuan Lama', 'Tujuan Baru', 'Perihal Lama', 'Perihal Baru', 'Aksi']);
    l.getRange('A1:I1').setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  }
}


function getInitialData() {
  initDatabase();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Database_Agenda');

  var rows = sh.getDataRange().getValues();
  var maxIds = {};
  var agendas = [];


  // ==========================================
  // 1. HITUNG NOMOR URUT TERBESAR PER KELOMPOK
  // ==========================================

  var maxFormatLama = {};

  // Ambil nomor terbesar dari setiap format dan klasifikasi.
  for (var i = 1; i < rows.length; i++) {
    var klas = String(rows[i][3] || '').trim();
    var num = Number(rows[i][0]) || 0;
    var fmt = String(rows[i][9] || '').trim();

    // Data tanpa klasifikasi atau format tidak dihitung.
    if (!klas || !fmt) continue;

    var key = fmt + '|' + klas;

    if (!maxFormatLama[key] || num > maxFormatLama[key]) {
      maxFormatLama[key] = num;
    }
  }

  // Masukkan nomor terbesar ke kelompok masing-masing.
  for (var key in maxFormatLama) {
    var bagian = key.split('|');
    var fmt = bagian[0];
    var klas = bagian[1];

    var kelompok = (fmt === '3' || fmt === '4') ? '34' : fmt;
    var targetKey = kelompok + '|' + klas;

    if (!maxIds[targetKey] || maxFormatLama[key] > maxIds[targetKey]) {
      maxIds[targetKey] = maxFormatLama[key];
    }
  }

  // Format 1 dan 2 juga menjadi acuan nomor Format 3 dan 4.
  for (var key in maxFormatLama) {
    var bagian = key.split('|');
    var fmt = bagian[0];
    var klas = bagian[1];

    if (fmt === '1' || fmt === '2') {
      var targetKey = '34|' + klas;

      if (!maxIds[targetKey] || maxFormatLama[key] > maxIds[targetKey]) {
        maxIds[targetKey] = maxFormatLama[key];
      }
    }
  }



  // ==========================================
  // 2. AMBIL SELURUH DATA AGENDA
  // ==========================================

  for (var j = rows.length - 1; j >= 1; j--) {
    agendas.push({
      no: rows[j][0],

      tanggal: rows[j][1] instanceof Date
        ? Utilities.formatDate(
            rows[j][1],
            Session.getScriptTimeZone(),
            'yyyy-MM-dd HH:mm:ss'
          )
        : String(rows[j][1] || ''),

      nomorSurat: String(rows[j][2] || ''),
      klasifikasi: String(rows[j][3] || ''),
      asalSurat: String(rows[j][4] || '-'),
      sifatSurat: String(rows[j][5] || '-'),
      tujuan: String(rows[j][6] || ''),
      perihal: String(rows[j][7] || ''),
      petugas: String(rows[j][8] || '')
    });
  }

  // ==========================================
  // 3. AMBIL DATA MASTER
  // ==========================================

  function list_(name) {
    var s = ss.getSheetByName(name);

    if (!s) {
      throw new Error('Sheet ' + name + ' tidak ditemukan.');
    }

    var v = s.getDataRange().getValues();
    var out = [];

    for (var n = 1; n < v.length; n++) {
      out.push(v[n]);
    }

    return out;
  }

  // ==========================================
  // 4. KIRIM DATA KE FRONTEND
  // ==========================================

  return {
    agendas: agendas,

    klasifikasis: list_('Klasifikasi').map(function(r) {
      return {
        kode: r[0],
        keterangan: r[1]
      };
    }),

    asalSurats: list_('Asal_Surat').map(function(r) {
      return {
        kode: r[0],
        instansi: r[1]
      };
    }),

    sifatSurats: list_('Sifat_Surat').map(function(r) {
      return {
        sifat: r[0],
        keterangan: r[1]
      };
    }),

    maxIds: maxIds,
    totalArsip: rows.length - 1
  };
}



function addAgendaRecord(data) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);
    initDatabase();

    // ==========================================
    // 1. VALIDASI DATA
    // ==========================================
    if (!data || !String(data.klasifikasi || '').trim()) {
      throw new Error('Pilih klasifikasi surat.');
    }

    if (
      !String(data.petugas || '').trim() ||
      !String(data.tujuan || '').trim() ||
      !String(data.perihal || '').trim()
    ) {
      throw new Error('Petugas, tujuan, dan perihal wajib diisi.');
    }

    var sh = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('Database_Agenda');

    if (!sh) {
      throw new Error('Sheet Database_Agenda tidak ditemukan.');
    }

    // ==========================================
    // 2. KONFIGURASI FORMAT DAN KLASIFIKASI
    // ==========================================
    var rows = sh.getDataRange().getValues();

    var format = String(data.formatNomor || '1');
    var klasifikasi = String(data.klasifikasi || '').trim();

    var kelompokFormat = {
      '1': ['1'],
      '2': ['2'],
      '3': ['1', '2', '3', '4'],
      '4': ['1', '2', '3', '4'],
      '5': ['5'],
      '6': ['6']
    };

    if (!kelompokFormat[format]) {
      throw new Error('Format nomor tidak valid.');
    }

    // ==========================================
    // 3. HITUNG NOMOR URUT BERDASARKAN KLASIFIKASI
    // ==========================================
    var formatYangDihitung = kelompokFormat[format];
    var next = 1;

    for (var i = 1; i < rows.length; i++) {
      // Kolom D = Klasifikasi
      var klasLama = String(rows[i][3] || '').trim();

      // Kolom A = No Urut
      var nomorLama = Number(rows[i][0]) || 0;

      // Kolom J = Format Naskah
      var formatLama = String(rows[i][9] || '').trim();

      // Hanya menghitung klasifikasi yang sama
      if (klasLama !== klasifikasi) {
        continue;
      }

      // Hanya menghitung format yang termasuk kelompok
      if (formatYangDihitung.indexOf(formatLama) === -1) {
        continue;
      }

      // Mengambil nomor urut terbesar
      next = Math.max(next, nomorLama + 1);
    }

    // ==========================================
    // 4. GENERATE NOMOR SURAT
    // ==========================================
    var now = new Date();
    var year = now.getFullYear();
    var number = '';

    switch (format) {
      // Format 1: Surat Perintah
      case '1':
        number = 'WP.19.PAS.9/' + klasifikasi + '/' + next;
        break;

      // Format 2: Naskah Keputusan/SK
      case '2':
        number = 'WP.19.PAS.9/' + next + '/' + klasifikasi + '/' + year;
        break;

      // Format 3: Korespondensi Internal
      case '3':
        number = 'WP.19.PAS.9/' + klasifikasi + '/' + next + '/' + year;
        break;

      // Format 4: Korespondensi Eksternal
      case '4':
        var sec = String(
          data.klasifikasiKeamanan || ''
        ).toUpperCase();

        if (['SR', 'R', 'T', 'B'].indexOf(sec) < 0) {
          throw new Error(
            'Pilih klasifikasi keamanan SR, R, T, atau B.'
          );
        }

        number = 'WP.19.PAS.9/' + sec + '/' + klasifikasi + '/' + next + '/' + year;
        break;

      // Format 5: Dinas Khusus
      case '5':
        number = 'WP.19.PAS.9/' + klasifikasi + '/' + next + '/' + year;
        break;

      // Format 6: SOP
      case '6':
        number = 'WP.19.PAS.9/' + next + '/' + klasifikasi + '/' + year;
        break;

      default:
        throw new Error('Format nomor tidak valid.');
    }

    // ==========================================
    // 5. SIMPAN DATA KE SPREADSHEET
    // ==========================================
    sh.appendRow([
      next,
      Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd HH:mm:ss'
      ),
      number,
      klasifikasi,
      String(data.asalSurat || '-').toUpperCase(),
      String(data.sifatSurat || '-').toUpperCase(),
      String(data.tujuan).trim().toUpperCase(),
      String(data.perihal).trim().toUpperCase(),
      String(data.petugas).trim().toUpperCase(),
      format
    ]);

    return {
      success: true,
      nomorSurat: number
    };

  } catch (e) {
    return {
      success: false,
      message: e.message || String(e)
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}


// Pengguna mengajukan permintaan. Tidak ada perubahan arsip pada tahap ini.
function createEditRequest(nomorSurat) {
  initDatabase();
  nomorSurat = String(nomorSurat || '').trim();
  
  if (!nomorSurat) throw new Error('Nomor surat tidak valid.');
  if (!findAgendaRow_(nomorSurat)) throw new Error('Arsip tidak ditemukan.');
  
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Edit_Requests'),
      id = Utilities.getUuid();
      
  sh.appendRow([id, nomorSurat, 'PENDING', new Date(), safeRequester_(), '', '']);
  return { success: true, requestId: id };
}

function getEditRequestStatus(requestId) {
  initDatabase();
  var r = findRequest_(requestId);
  
  if (!r) throw new Error('Permintaan tidak ditemukan.');
  
  var statusMsg = '';
  if (r.values[2] === 'APPROVED') {
    statusMsg = 'Admin menyetujui. Kamu hanya dapat mengubah Tujuan dan Perihal.';
  } else if (r.values[2] === 'REJECTED') {
    statusMsg = 'Permintaan ditolak admin.';
  } else {
    statusMsg = 'Menunggu persetujuan admin.';
  }
  
  return { success: true, status: String(r.values[2]), message: statusMsg };
}

// Admin PIN harus disimpan di Project Settings > Script properties sebagai E_AGENDA_ADMIN_PIN.
function adminAuthenticate(pin) {
  var isValid = validAdminPin_(pin);
  return { 
    success: isValid, 
    message: isValid ? 'Autentikasi berhasil.' : 'PIN admin salah atau belum dikonfigurasi.' 
  };
}

function adminListEditRequests(pin) {
  requireAdmin_(pin);
  initDatabase();
  
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Edit_Requests'),
      v = sh.getDataRange().getValues(),
      out = [];
      
  for (var i = 1; i < v.length; i++) {
    out.push({
      requestId: String(v[i][0]),
      nomorSurat: String(v[i][1]),
      status: String(v[i][2]),
      createdAt: v[i][3] instanceof Date ? Utilities.formatDate(v[i][3], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(v[i][3] || ''),
      requester: String(v[i][4] || '')
    });
  }
  
  out.reverse();
  return { success: true, requests: out };
}

function adminDecideEditRequest(pin, requestId, decision) {
  requireAdmin_(pin);
  
  if (['APPROVED', 'REJECTED'].indexOf(decision) < 0) {
    throw new Error('Keputusan tidak valid.');
  }
  
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    var req = findRequest_(requestId);
    if (!req) throw new Error('Permintaan tidak ditemukan.');
    if (req.values[2] !== 'PENDING') throw new Error('Permintaan ini sudah diproses.');
    if (!findAgendaRow_(String(req.values[1]))) throw new Error('Arsip sudah tidak ditemukan.');
    
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Edit_Requests');
    sh.getRange(req.row, 3).setValue(decision);
    sh.getRange(req.row, 6, 1, 2).setValues([[new Date(), safeRequester_()]]);
    
    return { success: true, message: 'Permintaan ' + (decision === 'APPROVED' ? 'disetujui.' : 'ditolak.') };
  } finally {
    lock.releaseLock();
  }
}

// Server hanya menerima dua field yang diizinkan dan mencari kolom berdasarkan nomor surat.
function saveApprovedEdit(payload) {
  if (!payload) throw new Error('Data edit tidak tersedia.');
  
  var id = String(payload.requestId || ''),
      number = String(payload.nomorSurat || ''),
      tujuan = String(payload.tujuan || '').trim(),
      perihal = String(payload.perihal || '').trim();
      
  if (!id || !number) throw new Error('ID permintaan atau nomor surat tidak valid.');
  if (!tujuan || !perihal) throw new Error('Tujuan dan perihal wajib diisi.');
  if (tujuan.length > 500 || perihal.length > 2000) throw new Error('Tujuan maksimal 500 karakter dan perihal maksimal 2.000 karakter.');
  
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    var req = findRequest_(id);
    if (!req || String(req.values[1]) !== number) throw new Error('Permintaan tidak cocok dengan arsip.');
    if (req.values[2] !== 'APPROVED') throw new Error('Admin belum menyetujui permintaan edit.');
    
    var found = findAgendaRow_(number);
    if (!found) throw new Error('Arsip tidak ditemukan.');
    
    var sh = found.sheet,
        oldTujuan = String(found.values[6] || ''),
        oldPerihal = String(found.values[7] || '');
        
    sh.getRange(found.row, 7, 1, 2).setValues([[tujuan.toUpperCase(), perihal.toUpperCase()]]);
    
    var audit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit_Edit');
    audit.appendRow([
      new Date(), 
      id, 
      number, 
      'Admin terautentikasi', 
      oldTujuan, 
      tujuan.toUpperCase(), 
      oldPerihal, 
      perihal.toUpperCase(), 
      'EDIT_TUJUAN_PERIHAL'
    ]);
    
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Edit_Requests').getRange(req.row, 3).setValue('COMPLETED');
    
    return { success: true, message: 'Tujuan dan perihal diperbarui.' };
  } finally {
    lock.releaseLock();
  }
}

function validAdminPin_(pin) {
  var expected = PropertiesService
    .getScriptProperties()
    .getProperty('E_AGENDA_ADMIN_PIN');

  if (!expected || !pin) {
    return false;
  }

  return String(pin).trim() === String(expected).trim();
}

function requireAdmin_(pin) {
  if (!validAdminPin_(pin)) throw new Error('Akses admin ditolak. PIN salah atau belum dikonfigurasi.');
}

function findRequest_(id) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Edit_Requests');
  if (!sh) return null;
  
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(id)) return { sheet: sh, row: i + 1, values: v[i] };
  }
  return null;
}

function findAgendaRow_(number) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Database_Agenda');
  if (!sh) return null;
  
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][2]) === String(number)) return { sheet: sh, row: i + 1, values: v[i] };
  }
  return null;
}

function safeRequester_() {
  try {
    return Session.getActiveUser().getEmail() || 'Pengguna E-Agenda';
  } catch (e) {
    return 'Pengguna E-Agenda';
  }
}




/**
 * ============================================================
 * MIGRASI OTOMATIS ARSIP E-AGENDA LAMA
 * ============================================================
 *
 * Format 1 = Surat Perintah
 * Format 2 = Naskah Keputusan
 *
 * Kolom:
 * A = No
 * C = Nomor Surat
 * D = Klasifikasi
 * J = Format Naskah
 *
 * Tidak mengubah nomor surat, tanggal, atau data arsip lainnya.
 */

// ------------------------------------------------------------
// 1. DETEKSI FORMAT BERDASARKAN POLA NOMOR SURAT LAMA
// ------------------------------------------------------------

function suggestLegacyFormat_(nomor, klas) {
  nomor = String(nomor || '').trim();
  klas = String(klas || '').trim();

  if (!nomor || !klas) return '';

  var escapedKlas = klas.replace(
    /[.*+?^${}()|[\]\\]/g,
    '\\$&'
  );

  // Format 1: Surat Perintah
  // WP.19.PAS.9-KLASIFIKASI-NOMOR
  var pola1 = new RegExp(
    '^WP\\.19\\.PAS\\.9-' +
    escapedKlas +
    '-\\d+$'
  );

  if (pola1.test(nomor)) return '1';

  // Format 2: Naskah Keputusan
  // WP.19.PAS.9-NOMOR-KLASIFIKASI-TAHUN
  var pola2 = new RegExp(
    '^WP\\.19\\.PAS\\.9-\\d+-' +
    escapedKlas +
    '-\\d{4}$'
  );

  if (pola2.test(nomor)) return '2';

  return '';
}


// ------------------------------------------------------------
// 2. PREVIEW SEBELUM MIGRASI
// ------------------------------------------------------------

function previewLegacyAgendaFormats() {
  var sh = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Database_Agenda');

  if (!sh) {
    throw new Error('Sheet Database_Agenda tidak ditemukan.');
  }

  var data = sh.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log('Belum ada data agenda.');
    return [];
  }

  var headers = data[0].map(function(v) {
    return String(v || '').trim().toLowerCase();
  });

  var col = {
    no: headers.indexOf('no'),
    nomor: headers.indexOf('nomor surat'),
    klas: headers.indexOf('klasifikasi'),
    format: headers.indexOf('format naskah')
  };

  if (Object.values(col).some(function(v) {
    return v < 0;
  })) {
    throw new Error(
      'Header harus memuat No, Nomor Surat, Klasifikasi, dan Format Naskah.'
    );
  }

  var result = [];
  var summary = {
    totalData: 0,
    format1: 0,
    format2: 0,
    sudahTerisi: 0,
    perluPemeriksaan: 0
  };

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var nomor = String(row[col.nomor] || '').trim();
    var klas = String(row[col.klas] || '').trim();
    var existing = String(row[col.format] || '').trim();

    if (!nomor && !klas) continue;

    summary.totalData++;

    var suggestion = suggestLegacyFormat_(nomor, klas);

    if (existing) {
      summary.sudahTerisi++;
    } else if (suggestion === '1') {
      summary.format1++;
    } else if (suggestion === '2') {
      summary.format2++;
    } else {
      summary.perluPemeriksaan++;
    }

    result.push({
      baris: i + 1,
      no: row[col.no],
      nomorSurat: nomor,
      klasifikasi: klas,
      formatSaatIni: existing || '(kosong)',
      saranFormat: existing || suggestion || 'PERLU CEK MANUAL'
    });
  }

  Logger.log('RINGKASAN PREVIEW:');
  Logger.log(JSON.stringify(summary, null, 2));

  Logger.log('CONTOH MAKSIMAL 30 DATA:');
  Logger.log(JSON.stringify(result.slice(0, 30), null, 2));

  return result;
}


// ------------------------------------------------------------
// 3. MIGRASI OTOMATIS TANPA PEMETAAN MANUAL
// ------------------------------------------------------------

function migrateLegacyAgendaFormats() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName('Database_Agenda');

    if (!sh) {
      throw new Error('Sheet Database_Agenda tidak ditemukan.');
    }

    var data = sh.getDataRange().getValues();

    if (data.length < 2) {
      throw new Error('Belum ada data agenda untuk dimigrasikan.');
    }

    var headers = data[0].map(function(v) {
      return String(v || '').trim().toLowerCase();
    });

    var colNomor = headers.indexOf('nomor surat');
    var colKlas = headers.indexOf('klasifikasi');
    var colFormat = headers.indexOf('format naskah');

    if (colNomor < 0 || colKlas < 0 || colFormat < 0) {
      throw new Error(
        'Kolom Nomor Surat, Klasifikasi, atau Format Naskah tidak ditemukan.'
      );
    }

    var updates = [];
    var skipped = [];
    var unmatched = [];

    // Periksa seluruh arsip sebelum menulis.
    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      var nomor = String(row[colNomor] || '').trim();
      var klas = String(row[colKlas] || '').trim();
      var existing = String(row[colFormat] || '').trim();

      if (!nomor && !klas) continue;

      // Jangan menimpa format yang sudah terisi.
      if (existing) {
        skipped.push(i + 1);
        continue;
      }

      var format = suggestLegacyFormat_(nomor, klas);

      if (format === '1' || format === '2') {
        updates.push({
          row: i + 1,
          format: format
        });
      } else {
        unmatched.push({
          baris: i + 1,
          nomorSurat: nomor,
          klasifikasi: klas
        });
      }
    }

    // Isi hanya kolom Format Naskah.
    updates.forEach(function(item) {
      sh.getRange(item.row, colFormat + 1)
        .setValue(item.format);
    });

    SpreadsheetApp.flush();

    var hasil = {
      success: true,
      jumlahFormat1: updates.filter(function(x) {
        return x.format === '1';
      }).length,
      jumlahFormat2: updates.filter(function(x) {
        return x.format === '2';
      }).length,
      jumlahMigrasi: updates.length,
      jumlahDilewati: skipped.length,
      jumlahPerluPemeriksaan: unmatched.length,
      barisTidakDikenali: unmatched,
      pesan: 'Migrasi selesai. Nomor surat dan data arsip lainnya tidak diubah.'
    };

    Logger.log(JSON.stringify(hasil, null, 2));

    return hasil;

  } finally {
    lock.releaseLock();
  }
}
