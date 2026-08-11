function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistem Jurnal KBM & Presensi Siswa')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * MENGAMBIL SELURUH DATA MASTER DENGAN AMAN
 */
function getInitialData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Helper untuk mengambil data tab sheet dengan aman
    const getDataFromSheet = (sheetName) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];
      const values = sheet.getDataRange().getValues();
      if (values.length <= 1) return []; // Jika cuma ada header / kosong
      values.shift(); // Buang baris header
      return values;
    };

    // Ambil data per sheet (Sesuaikan nama tab di spreadsheet jika beda)
    const rawGuru = getDataFromSheet('Data_Guru');
    const rawKelas = getDataFromSheet('Data_Kelas');
    const rawMapel = getDataFromSheet('Data_Mapel');
    const rawMateri = getDataFromSheet('Data_Materi');
    const rawSiswa = getDataFromSheet('Data_Siswa');

    // Mapping ke format Objek JavaScript
    const gurus = rawGuru.map(r => ({ nip: String(r[0]), nama: String(r[1]) }));
    const kelases = rawKelas.map(r => ({ id: String(r[0]), nama: String(r[1]) }));
    const mapels = rawMapel.map(r => ({ id: String(r[0]), nama: String(r[1]) }));
    const materis = rawMateri.map(r => ({ id: String(r[0]), mapelId: String(r[1]), kelasId: String(r[2]), nama: String(r[3]) }));
    const siswas = rawSiswa.map(r => ({ nis: String(r[0]), nama: String(r[1]), kelasId: String(r[2]) }));

    return {
      status: 'success',
      data: { gurus, kelases, mapels, materis, siswas }
    };

  } catch (err) {
    return {
      status: 'error',
      message: err.toString()
    };
  }
}

/**
 * SIMPAN JURNAL DAN PRESENSI
 */
function simpanJurnalAndPresensi(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Simpan ke Sheet Jurnal_KBM
    let sheetJurnal = ss.getSheetByName('Jurnal_KBM');
    if (!sheetJurnal) {
      sheetJurnal = ss.insertSheet('Jurnal_KBM');
      sheetJurnal.appendRow(['ID', 'Tanggal', 'NIP', 'KelasID', 'MapelID', 'JamKe', 'TotalJP', 'MateriID', 'Catatan']);
    }
    const idJurnal = 'JRN-' + new Date().getTime();
    sheetJurnal.appendRow([
      idJurnal, payload.tanggal, payload.nip, payload.kelasId, 
      payload.mapelId, payload.jamKe, payload.totalJP, payload.materiId, payload.catatan
    ]);

    // 2. Simpan ke Sheet Presensi_Siswa
    let sheetPresensi = ss.getSheetByName('Presensi_Siswa');
    if (!sheetPresensi) {
      sheetPresensi = ss.insertSheet('Presensi_Siswa');
      sheetPresensi.appendRow(['ID_Jurnal', 'Tanggal', 'KelasID', 'NIS', 'Nama', 'Status']);
    }
    
    if (payload.presensi && payload.presensi.length > 0) {
      payload.presensi.forEach(p => {
        sheetPresensi.appendRow([idJurnal, payload.tanggal, payload.kelasId, p.nis, p.nama, p.status]);
      });
    }

    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * REKAP PRESENSI SISWA
 */
function getRekapSiswa(kelasId, tipe, tglAwal, tglAkhir) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa = ss.getSheetByName('Data_Siswa');
    const sheetPresensi = ss.getSheetByName('Presensi_Siswa');

    if (!sheetSiswa) return { status: 'success', detailSiswa: [] };

    const dataSiswa = sheetSiswa.getDataRange().getValues();
    dataSiswa.shift();
    const siswasInKelas = dataSiswa.filter(r => String(r[2]) === String(kelasId));

    let presensiRows = [];
    if (sheetPresensi) {
      presensiRows = sheetPresensi.getDataRange().getValues();
      presensiRows.shift();
    }

    const detailSiswa = siswasInKelas.map(s => {
      const nis = String(s[0]);
      const nama = String(s[1]);
      
      let h = 0, i = 0, sCount = 0, a = 0;

      presensiRows.forEach(p => {
        const pTgl = String(p[1]);
        const pNis = String(p[3]);
        const pStatus = String(p[5]);

        if (pNis === nis && pTgl >= tglAwal && pTgl <= tglAkhir) {
          if (pStatus === 'H') h++;
          else if (pStatus === 'I') i++;
          else if (pStatus === 'S') sCount++;
          else if (pStatus === 'A') a++;
        }
      });

      return { nis, nama, h, i, s: sCount, a };
    });

    return { status: 'success', detailSiswa };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * REKAP JURNAL GURU
 */
function getRekapGuru(nip, tipe, tglAwal, tglAkhir) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetJurnal = ss.getSheetByName('Jurnal_KBM');

    if (!sheetJurnal) return { status: 'success', jurnal: [], statistik: { totalSesi: 0, totalJP: 0 } };

    const rows = sheetJurnal.getDataRange().getValues();
    rows.shift();

    let totalSesi = 0;
    let totalJP = 0;
    const listJurnal = [];

    rows.forEach(r => {
      const tgl = String(r[1]);
      const rNip = String(r[2]);

      if (rNip === String(nip) && tgl >= tglAwal && tgl <= tglAkhir) {
        totalSesi++;
        const jp = parseInt(r[6]) || 0;
        totalJP += jp;

        listJurnal.push({
          tanggal: tgl,
          kelas: r[3],
          mapel: r[4],
          jamKe: r[5],
          materi: r[7],
          catatan: r[8]
        });
      }
    });

    return {
      status: 'success',
      jurnal: listJurnal,
      statistik: { totalSesi, totalJP }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
