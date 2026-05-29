// ==================== KONFIGURASI ====================
const CONFIG = {
  SPREADSHEET_ID: '1p61twRCan2BTkxpfmst2kG0es1EGY0uiB5M94i2SjyE',
  SHEET_DATA: 'Data',
  SHEET_DOSEN: 'Dosen',
  ADMIN_PASSWORD: 'admin123',
  RATE_KOREKSI_UTS: 3000,
  RATE_KOREKSI_UAS: 4000,
  RATE_NASKAH: 16000
};

// ==================== HELPERS ====================
function getSS() {
  try {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch (e) {
    try { return SpreadsheetApp.getActiveSpreadsheet(); }
    catch (e2) { throw new Error('Tidak dapat mengakses Spreadsheet!'); }
  }
}

function findSheet(names) {
  const ss = getSS();
  const sheets = ss.getSheets();
  const lowerNames = names.map(n => n.toLowerCase());
  for (let s of sheets) {
    if (lowerNames.includes(s.getName().trim().toLowerCase())) return s;
  }
  return null;
}

function createDosenSheet() {
  const ss = getSS();
  const sheet = ss.insertSheet(CONFIG.SHEET_DOSEN);
  sheet.getRange('A1:D1').setValues([['No', 'Nama', 'NIDN', 'PPH']])
    .setFontWeight('bold').setBackground('#4361ee').setFontColor('#fff');
  return sheet;
}

function createDataSheet() {
  const ss = getSS();
  const sheet = ss.insertSheet(CONFIG.SHEET_DATA);
  sheet.getRange('A1:N1').setValues([['No','Tanggal','Jenis Ujian','Dosen','Matakuliah','Prodi','Sem/Kelas',
    'Jlh Mahasiswa','Jlh Naskah','H. Koreksi','H. Naskah','Jumlah','PPH','Jumlah Diterima']])
    .setFontWeight('bold').setBackground('#4361ee').setFontColor('#fff');
  return sheet;
}

function reNumberSheet(sheet) {
  if (!sheet) return;
  const lr = sheet.getLastRow();
  if (lr <= 1) return;
  for (let i = 2; i <= lr; i++) sheet.getRange(i, 1).setValue(i - 1);
}

function parseNumber(val) {
  if (typeof val === 'number') return val;
  return parseFloat(String(val).replace('%','').replace(',','.').trim()) || 0;
}

function formatDate(d) {
  try {
    if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    return String(d || '');
  } catch (e) { return String(d || ''); }
}

// ==================== WEB APP ====================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistem Honorer Dosen')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== SETUP ====================
function testConnection() {
  try {
    const ss = getSS();
    const sheets = ss.getSheets();
    let info = `=== KONEKSI BERHASIL ===\nNama: ${ss.getName()}\n`;
    sheets.forEach((s, i) => {
      info += `${i+1}. "${s.getName()}" - ${s.getLastRow()} baris\n`;
    });
    Logger.log(info);
    return info;
  } catch (error) {
    return 'Error: ' + error.toString();
  }
}

function setupSheets() {
  try {
    let ds = findSheet(['dosen', 'data dosen']);
    if (!ds) {
      ds = createDosenSheet();
      ds.getRange('A2:D4').setValues([
        [1, 'Dr. Ahmad Fauzi, M.Pd', '1234567890', 5],
        [2, 'Prof. Siti Aminah, M.Si', '0987654321', 15],
        [3, 'Budi Santoso, S.Pd', '1122334455', 5]
      ]);
    }
    if (!findSheet(['data', 'data input'])) createDataSheet();
    return 'Setup selesai!';
  } catch (error) {
    return 'Error: ' + error.toString();
  }
}

// MIGRASI: Tambahkan kolom Jenis Ujian jika belum ada
function migrateDataSheet() {
  try {
    const sheet = findSheet(['data', 'data input']);
    if (!sheet) return { success: false, message: 'Sheet Data tidak ditemukan' };
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const hasJenisUjian = headers.some(h => String(h).toLowerCase().trim() === 'jenis ujian');
    
    if (hasJenisUjian) return { success: true, message: 'Sudah ada kolom Jenis Ujian' };
    
    // Insert kolom Jenis Ujian setelah kolom Tanggal (column 3)
    sheet.insertColumnAfter(2);
    sheet.getRange(1, 3).setValue('Jenis Ujian').setFontWeight('bold').setBackground('#4361ee').setFontColor('#fff');
    
    // Set default value untuk data lama
    const lr = sheet.getLastRow();
    if (lr > 1) {
      const range = sheet.getRange(2, 3, lr - 1, 1);
      range.setValue('UJIAN TENGAH SEMESTER');
    }
    
    return { success: true, message: 'Migrasi berhasil! Kolom "Jenis Ujian" ditambahkan.' };
  } catch (error) {
    return { success: false, message: 'Error migrasi: ' + error.message };
  }
}

// ==================== DOSEN CRUD ====================
function getDosenList() {
  try {
    let sheet = findSheet(['dosen', 'data dosen']);
    if (!sheet) { createDosenSheet(); return []; }
    
    const lr = sheet.getLastRow(), lc = sheet.getLastColumn();
    if (lr <= 1) return [];
    
    const allData = sheet.getRange(1, 1, lr, lc).getValues();
    const headers = allData[0].map(h => String(h).toLowerCase().trim());
    
    const colMap = {
      no: headers.findIndex(h => ['no','nomor'].includes(h)),
      nama: headers.findIndex(h => ['nama','name','nama dosen','dosen'].includes(h)),
      nidn: headers.findIndex(h => ['nidn','nip'].includes(h)),
      pph: headers.findIndex(h => ['pph','pajak','pph 21','pph21'].includes(h))
    };
    
    if (colMap.nama === -1) colMap.nama = 1;
    if (colMap.nidn === -1) colMap.nidn = 2;
    if (colMap.pph === -1) colMap.pph = 3;
    if (colMap.no === -1) colMap.no = 0;
    
    const result = [];
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const nama = String(row[colMap.nama] || '').trim();
      if (nama && !['undefined','null','nama','nama dosen'].includes(nama.toLowerCase())) {
        result.push({
          no: row[colMap.no] || i,
          nama: nama,
          nidn: String(row[colMap.nidn] || '').trim(),
          pph: parseNumber(row[colMap.pph])
        });
      }
    }
    return result;
  } catch (error) {
    throw new Error('Gagal memuat dosen: ' + error.message);
  }
}

function addDosen(nama, nidn, pph) {
  try {
    let sheet = findSheet(['dosen']) || createDosenSheet();
    const lr = sheet.getLastRow();
    sheet.appendRow([lr < 1 ? 1 : lr, nama, nidn, parseFloat(pph) || 0]);
    reNumberSheet(sheet);
    return { success: true, message: `Dosen "${nama}" berhasil ditambahkan!` };
  } catch (error) {
    return { success: false, message: 'Gagal: ' + error.message };
  }
}

function deleteDosen(nama) {
  try {
    const sheet = findSheet(['dosen']);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan!' };
    const lr = sheet.getLastRow();
    if (lr <= 1) return { success: false, message: 'Data kosong!' };
    
    const data = sheet.getRange(2, 1, lr - 1, sheet.getLastColumn()).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase().trim() === nama.toLowerCase().trim()) {
        sheet.deleteRow(i + 2);
        reNumberSheet(sheet);
        return { success: true, message: `Dosen "${nama}" dihapus!` };
      }
    }
    return { success: false, message: 'Dosen tidak ditemukan!' };
  } catch (error) {
    return { success: false, message: 'Gagal: ' + error.message };
  }
}

function updateDosen(originalName, nama, nidn, pph) {
  try {
    const sheet = findSheet(['dosen']);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan!' };
    const lr = sheet.getLastRow();
    if (lr <= 1) return { success: false, message: 'Data kosong!' };
    
    const data = sheet.getRange(2, 1, lr - 1, sheet.getLastColumn()).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase().trim() === originalName.toLowerCase().trim()) {
        sheet.getRange(i + 2, 2).setValue(nama);
        sheet.getRange(i + 2, 3).setValue(nidn);
        sheet.getRange(i + 2, 4).setValue(parseFloat(pph) || 0);
        return { success: true, message: 'Data dosen diperbarui!' };
      }
    }
    return { success: false, message: 'Dosen tidak ditemukan!' };
  } catch (error) {
    return { success: false, message: 'Gagal: ' + error.message };
  }
}

// ==================== DATA CRUD ====================
function submitData(formData) {
  try {
    let sheet = findSheet(['data']) || createDataSheet();
    
    // Auto-migrate jika kolom Jenis Ujian belum ada
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const hasJenisUjian = headers.some(h => String(h).toLowerCase().trim() === 'jenis ujian');
    if (!hasJenisUjian) migrateDataSheet();
    
    const lr = sheet.getLastRow();
    const jm = parseInt(formData.jlhMahasiswa) || 0;
    const jn = parseInt(formData.jlhNaskah) || 0;
    const jenisUjian = formData.jenisUjian || 'UJIAN TENGAH SEMESTER';
    const rateKoreksi = jenisUjian === 'UJIAN AKHIR SEMESTER' ? CONFIG.RATE_KOREKSI_UAS : CONFIG.RATE_KOREKSI_UTS;
    const hk = jm * rateKoreksi;
    const hn = jn * CONFIG.RATE_NASKAH;
    const jml = hk + hn;
    const pph = parseFloat(formData.pph) || 0;
    const diterima = jml - (jml * pph / 100);
    
    sheet.appendRow([
      lr < 1 ? 1 : lr, new Date(), jenisUjian, formData.dosen, formData.matakuliah,
      formData.prodi, formData.semKelas, jm, jn, hk, hn, jml, pph + '%', diterima
    ]);
    
    const nlr = sheet.getLastRow();
    [10, 11, 12, 14].forEach(c => sheet.getRange(nlr, c).setNumberFormat('#,##0'));
    sheet.getRange(nlr, 2).setNumberFormat('dd/MM/yyyy HH:mm');
    reNumberSheet(sheet);
    
    return { success: true, message: 'Data berhasil disimpan!' };
  } catch (error) {
    return { success: false, message: 'Gagal: ' + error.message };
  }
}

function getSubmittedData() {
  try {
    const sheet = findSheet(['data']);
    if (!sheet) return [];
    const lr = sheet.getLastRow();
    if (lr <= 1) return [];
    
    const lc = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h).toLowerCase().trim());
    const hasJenisUjian = headers.includes('jenis ujian');
    
    const data = sheet.getRange(2, 1, lr - 1, lc).getValues();
    
    return data.map((row, i) => {
      if (hasJenisUjian) {
        // Format baru dengan kolom Jenis Ujian (14 kolom)
        return {
          no: row[0] || (i + 1),
          tanggal: formatDate(row[1]),
          jenisUjian: String(row[2] || 'UJIAN TENGAH SEMESTER'),
          dosen: String(row[3] || ''),
          matakuliah: String(row[4] || ''),
          prodi: String(row[5] || ''),
          semKelas: String(row[6] || ''),
          jlhMahasiswa: row[7] || 0,
          jlhNaskah: row[8] || 0,
          hKoreksi: row[9] || 0,
          hNaskah: row[10] || 0,
          jumlah: row[11] || 0,
          pph: String(row[12] || '0%'),
          jumlahDiterima: row[13] || 0
        };
      } else {
        // Format lama (13 kolom)
        return {
          no: row[0] || (i + 1),
          tanggal: formatDate(row[1]),
          jenisUjian: 'UJIAN TENGAH SEMESTER',
          dosen: String(row[2] || ''),
          matakuliah: String(row[3] || ''),
          prodi: String(row[4] || ''),
          semKelas: String(row[5] || ''),
          jlhMahasiswa: row[6] || 0,
          jlhNaskah: row[7] || 0,
          hKoreksi: row[8] || 0,
          hNaskah: row[9] || 0,
          jumlah: row[10] || 0,
          pph: String(row[11] || '0%'),
          jumlahDiterima: row[12] || 0
        };
      }
    });
  } catch (error) {
    return [];
  }
}

function deleteSubmittedData(rowIndex) {
  try {
    const sheet = findSheet(['data']);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan!' };
    sheet.deleteRow(rowIndex + 2);
    reNumberSheet(sheet);
    return { success: true, message: 'Data dihapus!' };
  } catch (error) {
    return { success: false, message: 'Gagal: ' + error.message };
  }
}

// BARU: Bulk delete
function deleteBulkData(indices) {
  try {
    const sheet = findSheet(['data']);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan!' };
    if (!indices || !indices.length) return { success: false, message: 'Tidak ada data dipilih!' };
    
    // Sort descending agar deleteRow tidak shift index lain
    const sortedIndices = indices.slice().sort((a, b) => b - a);
    let deletedCount = 0;
    
    sortedIndices.forEach(idx => {
      try {
        sheet.deleteRow(idx + 2);
        deletedCount++;
      } catch (e) {}
    });
    
    reNumberSheet(sheet);
    return { 
      success: true, 
      message: `${deletedCount} data berhasil dihapus!`,
      count: deletedCount
    };
  } catch (error) {
    return { success: false, message: 'Gagal: ' + error.message };
  }
}

function updateSubmittedData(rowIndex, formData) {
  try {
    const sheet = findSheet(['data']);
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan!' };
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const hasJenisUjian = headers.some(h => String(h).toLowerCase().trim() === 'jenis ujian');
    
    const ar = rowIndex + 2;
    const jm = parseInt(formData.jlhMahasiswa) || 0;
    const jn = parseInt(formData.jlhNaskah) || 0;
    const jenisUjian = formData.jenisUjian || 'UJIAN TENGAH SEMESTER';
    const rateKoreksi = jenisUjian === 'UJIAN AKHIR SEMESTER' ? CONFIG.RATE_KOREKSI_UAS : CONFIG.RATE_KOREKSI_UTS;
    const hk = jm * rateKoreksi;
    const hn = jn * CONFIG.RATE_NASKAH;
    const jml = hk + hn;
    const pph = parseFloat(formData.pph) || 0;
    const diterima = jml - (jml * pph / 100);
    
    if (hasJenisUjian) {
      // Format baru
      const updates = [
        [3, jenisUjian], [4, formData.dosen], [5, formData.matakuliah], [6, formData.prodi],
        [7, formData.semKelas], [8, jm], [9, jn], [10, hk], [11, hn],
        [12, jml], [13, pph + '%'], [14, diterima]
      ];
      updates.forEach(([c, v]) => sheet.getRange(ar, c).setValue(v));
      [10, 11, 12, 14].forEach(c => sheet.getRange(ar, c).setNumberFormat('#,##0'));
    } else {
      // Format lama
      const updates = [
        [3, formData.dosen], [4, formData.matakuliah], [5, formData.prodi],
        [6, formData.semKelas], [7, jm], [8, jn], [9, hk], [10, hn],
        [11, jml], [12, pph + '%'], [13, diterima]
      ];
      updates.forEach(([c, v]) => sheet.getRange(ar, c).setValue(v));
      [9, 10, 11, 13].forEach(c => sheet.getRange(ar, c).setNumberFormat('#,##0'));
    }
    
    return { success: true, message: 'Data diperbarui!' };
  } catch (error) {
    return { success: false, message: 'Gagal: ' + error.message };
  }
}

function getFilteredDataByDosen(dosenName, jenisUjian) {
  const all = getSubmittedData();
  let filtered = all;
  
  // Filter berdasarkan jenis ujian jika diberikan
  if (jenisUjian && jenisUjian !== 'all') {
    filtered = filtered.filter(d => d.jenisUjian === jenisUjian);
  }
  
  // Filter berdasarkan nama dosen jika diberikan
  if (dosenName && dosenName !== 'all') {
    filtered = filtered.filter(d => d.dosen.toLowerCase() === dosenName.toLowerCase());
  }
  
  return filtered;
}

function getUniqueDosenFromData() {
  try {
    const all = getSubmittedData();
    const unique = {};
    all.forEach(d => {
      if (d.dosen) {
        if (!unique[d.dosen]) unique[d.dosen] = { nama: d.dosen, totalDiterima: 0, count: 0 };
        unique[d.dosen].totalDiterima += parseFloat(d.jumlahDiterima) || 0;
        unique[d.dosen].count++;
      }
    });
    return Object.values(unique).sort((a, b) => a.nama.localeCompare(b.nama));
  } catch (e) { return []; }
}

function verifyAdmin(password) {
  return password === CONFIG.ADMIN_PASSWORD;
}

function getStatistics() {
  try {
    const data = getSubmittedData();
    const dosen = getDosenList();
    const totalDiterima = data.reduce((s, d) => s + (parseFloat(d.jumlahDiterima) || 0), 0);
    const totalMhs = data.reduce((s, d) => s + (parseInt(d.jlhMahasiswa) || 0), 0);
    const totalNaskah = data.reduce((s, d) => s + (parseInt(d.jlhNaskah) || 0), 0);
    const totalUTS = data.filter(d => d.jenisUjian === 'UJIAN TENGAH SEMESTER').length;
    const totalUAS = data.filter(d => d.jenisUjian === 'UJIAN AKHIR SEMESTER').length;
    return {
      totalData: data.length,
      totalDosen: dosen.length,
      totalDiterima: totalDiterima,
      totalMhs: totalMhs,
      totalNaskah: totalNaskah,
      totalUTS: totalUTS,
      totalUAS: totalUAS
    };
  } catch (e) {
    return { totalData: 0, totalDosen: 0, totalDiterima: 0, totalMhs: 0, totalNaskah: 0, totalUTS: 0, totalUAS: 0 };
  }
}