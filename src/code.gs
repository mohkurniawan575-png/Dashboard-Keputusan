/**
 * code.gs — Backend untuk Dashboard SK
 * Google Apps Script Web App
 *
 * Sumber data: Spreadsheet aktif (tanpa hardcode ID).
 * Deteksi sheet, header, dan kolom secara dinamis.
 */

// ═══════════════════════════════════════════════════════════════════
//  ENTRY POINT
// ═══════════════════════════════════════════════════════════════════

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Dashboard Surat Keputusan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ═══════════════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════════════

var BULAN = [
  'Januari','Februari','Maret','April','Mei','Juni',
  'Juli','Agustus','September','Oktober','November','Desember'
];

/**
 * parseDate_ — Ubah nilai sel jadi Date object.
 * Support: Date, dd/MM/yyyy, yyyy-MM-dd, dd-MM-yyyy.
 */
function parseDate_(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;

  var s = val.toString().trim();
  var p, d;
  if (s.includes('/')) {
    p = s.split('/');
    d = new Date(+p[2], +p[1] - 1, +p[0]);
  } else if (s.includes('-')) {
    p = s.split('-');
    d = p[0].length === 4 ? new Date(s) : new Date(+p[2], +p[1] - 1, +p[0]);
  } else {
    d = new Date(s);
  }
  return (d && !isNaN(d.getTime())) ? d : null;
}

/**
 * colIndex_ — Cari index kolom pertama yang cocok dengan salah satu keyword.
 * Case-insensitive, partial match. Return -1 jika tidak ada.
 */
function colIndex_(headers, keywords) {
  for (var i = 0; i < headers.length; i++) {
    var h = (headers[i] || '').toString().toUpperCase();
    for (var k = 0; k < keywords.length; k++) {
      if (h.includes(keywords[k].toUpperCase())) return i;
    }
  }
  return -1;
}

/**
 * colIndexAll_ — Cari SEMUA index kolom yang mengandung keyword.
 */
function colIndexAll_(headers, keyword) {
  var kw = keyword.toUpperCase();
  var out = [];
  for (var i = 0; i < headers.length; i++) {
    if ((headers[i] || '').toString().toUpperCase().includes(kw)) out.push(i);
  }
  return out;
}

/**
 * detectHeaderRow_ — Deteksi baris header di sheet.
 * Scan 5 baris pertama, pilih yang paling banyak sel non-kosong.
 * Return { row: 1-based, headers: string[] }
 */
function detectHeaderRow_(sheet) {
  var maxScan = Math.min(sheet.getLastRow(), 5);
  var lastCol = sheet.getLastColumn();
  var bestRow = 1, bestCount = 0, bestHeaders = [];

  for (var r = 1; r <= maxScan; r++) {
    var row = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    var count = 0;
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== '' && row[c] !== null && row[c] !== undefined) count++;
    }
    if (count > bestCount) {
      bestCount = count;
      bestRow = r;
      bestHeaders = row;
    }
  }
  return { row: bestRow, headers: bestHeaders };
}

/**
 * normStatus_ — Normalisasi status string.
 */
function normStatus_(s) {
  return (s || '').toString().toLowerCase().trim();
}

// ═══════════════════════════════════════════════════════════════════
//  SHEET RESOLVER
// ═══════════════════════════════════════════════════════════════════

/**
 * resolveSheet_ — Tentukan sheet sumber data.
 * Prioritas: "Pelacakan SK" → sheet aktif.
 */
function resolveSheet_(ss) {
  var sheet = ss.getSheetByName('Pelacakan SK');
  if (sheet) return sheet;

  // Coba beberapa nama alternatif
  var altNames = ['Data SK', 'Tracking SK', 'Pelacakan', 'Data'];
  for (var i = 0; i < altNames.length; i++) {
    sheet = ss.getSheetByName(altNames[i]);
    if (sheet) return sheet;
  }

  // Fallback ke sheet aktif
  return ss.getActiveSheet();
}

// ═══════════════════════════════════════════════════════════════════
//  ANALYTICS ENGINE
// ═══════════════════════════════════════════════════════════════════

/**
 * computeAnalytics_ — Hitung semua metrik dashboard dari data mentah.
 *
 * @param {string[]} headers — Header kolom.
 * @param {Array[]}  rows    — Data baris (2D array).
 * @returns {Object} Analytics object lengkap.
 */
function computeAnalytics_(headers, rows) {
  // ── Deteksi kolom ──────────────────────────────────────────────
  var noIdx    = colIndex_(headers, ['NO.', 'NO ', 'NOMOR']);
  var judulIdx = colIndex_(headers, ['JUDUL', 'NAMA KEPUTUSAN', 'PERIHAL']);
  var statusIdx = colIndex_(headers, ['STATUS PENGERJAAN', 'STATUS']);
  var jenisIdx  = colIndex_(headers, ['JENIS SK', 'JENIS KEPUTUSAN', 'JENIS']);

  // Tanggal masuk & selesai — cari semua kolom TANGGAL lalu bedakan
  var dateCols = colIndexAll_(headers, 'TANGGAL');
  var tglMasukIdx = -1, tglSelesaiIdx = -1;

  for (var i = 0; i < dateCols.length; i++) {
    var h = (headers[dateCols[i]] || '').toString().toUpperCase();
    if (h.includes('MASUK') || h.includes('DITERIMA') || h.includes('TERIMA') || h.includes('TERIMA')) {
      tglMasukIdx = dateCols[i];
    } else if (h.includes('SELESAI') || h.includes('DITETAPKAN') || h.includes('TETAP') || h.includes('KELUAR')) {
      tglSelesaiIdx = dateCols[i];
    }
  }
  // Fallback: juga cek keyword bahasa Inggris
  if (tglMasukIdx === -1) tglMasukIdx = colIndex_(headers, ['TANGGAL MASUK', 'DATE IN', 'TANGGAL_INPUT', 'TANGGAL TERIMA']);
  if (tglSelesaiIdx === -1) tglSelesaiIdx = colIndex_(headers, ['TANGGAL SELESAI', 'DATE DONE', 'TANGGAL_TETAP', 'TANGGAL DITETAPKAN']);
  // Last resort: kolom tanggal pertama & terakhir
  if (tglMasukIdx === -1 && dateCols.length >= 1) tglMasukIdx = dateCols[0];
  if (tglSelesaiIdx === -1 && dateCols.length >= 2) tglSelesaiIdx = dateCols[dateCols.length - 1];

  // Kolom "last update" — cari spesifik, fallback ke tglMasuk
  var updateIdx = colIndex_(headers, ['LAST UPDATE', 'UPDATE', 'TERAKHIR DIPERBARUI', 'TANGGAL UPDATE', 'TANGGAL PERUBAHAN']);

  // Perangkat daerah
  var deptIdx = colIndex_(headers, [
    'PERANGKAT DAERAH', 'PERANGKAT', 'INSTANSI', 'SKPD', 'OPD',
    'ORGANISASI', 'DINAS', 'BADAN', 'BIRO', 'KANTOR',
    'UNIT KERJA', 'LEMBAGA', 'DEPARTEMEN'
  ]);

  // ── Inisialisasi accumulator ───────────────────────────────────
  var counts = { total: 0, done: 0, proc: 0, pend: 0, rej: 0 };
  var typeMap = {};       // jenis SK → count
  var monthMap = {};      // 'YYYY-MM' → count
  var procDays = [];      // array selisih hari
  var deptMap = {};       // dept → { gubernur, sekda, total }
  var latestDate = null;  // tanggal terbaru untuk "last update"

  var now = new Date();
  var curYear = now.getFullYear();
  var curMonth = now.getMonth(); // 0-based

  // ── Iterasi data ───────────────────────────────────────────────
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];

    // Skip baris kosong / header berulang
    var noVal  = noIdx >= 0 ? row[noIdx] : (r + 1);
    var judVal = judulIdx >= 0 ? row[judulIdx] : true;
    if (!noVal || !judVal) continue;
    var noStr = (noVal || '').toString().toUpperCase();
    if (noStr.includes('URAIAN') || noStr.includes('JUMLAH') || noStr.includes('TOTAL')) continue;

    counts.total++;

    // ── Status ───────────────────────────────────────────────────
    var status = statusIdx >= 0 ? normStatus_(row[statusIdx]) : '';
    if (status === 'done' || status === 'selesai' || status === 'completed') {
      counts.done++;
    } else if (status.includes('progres') || status.includes('process') || status.includes('proses')) {
      counts.proc++;
    } else if (status === 'pending' || status === 'tertunda') {
      counts.pend++;
    } else if (status === 'rejected' || status === 'ditolak') {
      counts.rej++;
    }

    // ── Jenis keputusan ──────────────────────────────────────────
    var jenis = jenisIdx >= 0 ? (row[jenisIdx] || '').toString().trim() : '';
    if (jenis) {
      var jl = jenis.toLowerCase();
      var key;
      if (jl.includes('gubernur')) key = 'SK Gubernur';
      else if (jl.includes('sekda') || jl.includes('sekretaris daerah')) key = 'SK Sekda';
      else key = jenis.length > 35 ? jenis.substring(0, 35) + '…' : jenis;
      typeMap[key] = (typeMap[key] || 0) + 1;
    }

    // ── Tanggal masuk → monthly trend ────────────────────────────
    var dMasuk = tglMasukIdx >= 0 ? parseDate_(row[tglMasukIdx]) : null;
    if (dMasuk) {
      var mk = dMasuk.getFullYear() + '-' + String(dMasuk.getMonth() + 1).padStart(2, '0');
      monthMap[mk] = (monthMap[mk] || 0) + 1;
    }

    // ── Processing time (completed only) ─────────────────────────
    var dSelesai = tglSelesaiIdx >= 0 ? parseDate_(row[tglSelesaiIdx]) : null;
    if (dMasuk && dSelesai && (status === 'done' || status === 'selesai' || status === 'completed')) {
      var days = Math.round((dSelesai.getTime() - dMasuk.getTime()) / 86400000);
      if (days >= 0) procDays.push(days);
    }

    // ── Department ───────────────────────────────────────────────
    var dept = deptIdx >= 0 ? (row[deptIdx] || '').toString().trim() : '';
    if (dept) {
      if (!deptMap[dept]) deptMap[dept] = { gubernur: 0, sekda: 0, total: 0 };
      deptMap[dept].total++;
      if (jenis && jenis.toLowerCase().includes('gubernur')) deptMap[dept].gubernur++;
      else if (jenis && (jenis.toLowerCase().includes('sekda') || jenis.toLowerCase().includes('sekretaris daerah'))) deptMap[dept].sekda++;
    }

    // ── Last update detection ────────────────────────────────────
    var updateVal = updateIdx >= 0 ? parseDate_(row[updateIdx]) : null;
    if (!updateVal && tglMasukIdx >= 0) updateVal = dMasuk; // fallback
    if (updateVal && (!latestDate || updateVal > latestDate)) latestDate = updateVal;
  }

  // ── Post: Monthly average ──────────────────────────────────────
  var curKey = curYear + '-' + String(curMonth + 1).padStart(2, '0');
  var curMonthCount = monthMap[curKey] || 0;
  var activeMo = 0, totalYear = 0;
  for (var k in monthMap) {
    if (k.indexOf(String(curYear)) === 0) { activeMo++; totalYear += monthMap[k]; }
  }
  var monthAvg = activeMo > 0 ? Math.round(totalYear / activeMo) : 0;

  // ── Post: Average processing time ──────────────────────────────
  var avgProc = 0;
  if (procDays.length > 0) {
    var sum = 0;
    for (var p = 0; p < procDays.length; p++) sum += procDays[p];
    avgProc = Math.round(sum / procDays.length);
  }

  // ── Post: Monthly trend (sorted) ──────────────────────────────
  var sortedKeys = Object.keys(monthMap).sort();
  var tLabels = [], tValues = [];
  for (var m = 0; m < sortedKeys.length; m++) {
    var pp = sortedKeys[m].split('-');
    tLabels.push(BULAN[parseInt(pp[1], 10) - 1]);
    tValues.push(monthMap[sortedKeys[m]]);
  }

  // ── Post: Department ranking (sorted desc) ────────────────────
  var dEntries = [];
  for (var dn in deptMap) {
    dEntries.push({
      name: dn.length > 45 ? dn.substring(0, 45) + '…' : dn,
      gubernur: deptMap[dn].gubernur,
      sekda: deptMap[dn].sekda,
      total: deptMap[dn].total
    });
  }
  dEntries.sort(function(a, b) { return b.total - a.total; });
  var dLabels = [], dGov = [], dSek = [];
  for (var d = 0; d < dEntries.length; d++) {
    dLabels.push(dEntries[d].name);
    dGov.push(dEntries[d].gubernur);
    dSek.push(dEntries[d].sekda);
  }

  // ── Return analytics ───────────────────────────────────────────
  return {
    statusCounts: { total: counts.total, done: counts.done, proc: counts.proc, pend: counts.pend, rej: counts.rej },
    monthlyIncoming: { currentMonthCount: curMonthCount, average: monthAvg, activeMonths: activeMo, currentMonthName: BULAN[curMonth] },
    processingTime: { average: avgProc, sampleSize: procDays.length },
    decisionTypes: { labels: Object.keys(typeMap), values: Object.keys(typeMap).map(function(x) { return typeMap[x]; }) },
    statusMonitoring: { pending: counts.pend, onProgress: counts.proc, completed: counts.done, rejected: counts.rej },
    monthlyTrend: { labels: tLabels, values: tValues },
    departmentRanking: { labels: dLabels, gubernur: dGov, sekda: dSek },
    lastUpdate: latestDate
  };
}

// ═══════════════════════════════════════════════════════════════════
//  RAWDATA KPI READER
// ═══════════════════════════════════════════════════════════════════

/**
 * readKPIFromRawData_ — Baca nilai KPI dari sheet "Raw Data".
 *
 * Mapping cell:
 *   B24 → Completed SK
 *   B25 → On Progress
 *   B26 → Pending
 *   B27 → Rejected
 *   B31 → Rata-rata SK Masuk/Bulan
 *   B32 → Rata-rata Lama Pengerjaan SK (hari)
 *
 * @param {Spreadsheet} ss — Active spreadsheet.
 * @returns {Object|null} KPI values atau null jika sheet tidak ada.
 */
function readKPIFromRawData_(ss) {
  var sheet = ss.getSheetByName('Raw Data');
  if (!sheet) return null;

  var done = parseInt(sheet.getRange('B24').getValue(), 10) || 0;
  var proc = parseInt(sheet.getRange('B25').getValue(), 10) || 0;
  var pend = parseInt(sheet.getRange('B26').getValue(), 10) || 0;
  var rej  = parseInt(sheet.getRange('B27').getValue(), 10) || 0;

  var avgMonthly   = parseFloat(sheet.getRange('B31').getValue()) || 0;
  var avgProcTime  = parseFloat(sheet.getRange('B32').getValue()) || 0;

  return {
    done: done, proc: proc, pend: pend, rej: rej,
    total: done + proc + pend + rej,
    avgMonthly: avgMonthly,
    avgProcTime: avgProcTime
  };
}

// ═══════════════════════════════════════════════════════════════════
//  PUBLIC API
// ═══════════════════════════════════════════════════════════════════

/**
 * emptyResponse_ — Respons kosong yang valid (tidak pernah null).
 */
function emptyResponse_(msg) {
  return { analytics: null, lastUpdate: '', error: msg || null };
}

/**
 * getDashboardData — API utama yang dipanggil frontend via google.script.run.
 * Return: { analytics, lastUpdate, error }
 */
function getDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return emptyResponse_('Spreadsheet tidak ditemukan.');

    var tz = Session.getScriptTimeZone();
    var sheet = resolveSheet_(ss);
    if (!sheet || sheet.getLastRow() < 2) return emptyResponse_('Sheet data kosong atau tidak ditemukan.');

    // ── Deteksi header & baca data ──────────────────────────────
    var info = detectHeaderRow_(sheet);
    var headers = info.headers;
    var dataStartRow = info.row + 1;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < dataStartRow || lastCol < 1) {
      return emptyResponse_('Tidak ada data di bawah baris header.');
    }

    var rows = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol).getValues();

    // ── Hitung analytics ────────────────────────────────────────
    var analytics = computeAnalytics_(headers, rows);

    // ── Override KPI dari sheet Raw Data (cell B24-B27, B31-B32) ─
    var rawKPI = readKPIFromRawData_(ss);
    if (rawKPI) {
      analytics.statusCounts = {
        total: rawKPI.total,
        done:  rawKPI.done,
        proc:  rawKPI.proc,
        pend:  rawKPI.pend,
        rej:   rawKPI.rej
      };
      analytics.statusMonitoring = {
        completed:  rawKPI.done,
        onProgress: rawKPI.proc,
        pending:    rawKPI.pend,
        rejected:   rawKPI.rej
      };
      // Override rata-rata dari B31 & B32
      if (analytics.monthlyIncoming) {
        analytics.monthlyIncoming.average = rawKPI.avgMonthly;
      }
      if (analytics.processingTime) {
        analytics.processingTime.average = rawKPI.avgProcTime;
      }
    }

    // ── Format last update ──────────────────────────────────────
    var lastUpdateStr = '';
    if (analytics.lastUpdate) {
      lastUpdateStr = Utilities.formatDate(analytics.lastUpdate, tz, 'dd MMMM yyyy');
    }
    // Hapus field Date object sebelum dikirim ke client (tidak serializable)
    delete analytics.lastUpdate;

    return { analytics: analytics, lastUpdate: lastUpdateStr, spreadsheetUrl: ss.getUrl(), error: null };

  } catch (e) {
    console.error('getDashboardData error:', e);
    return emptyResponse_('Error server: ' + (e.message || e));
  }
}
