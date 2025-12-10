/*
 * File: commands.js
 * Fungsi: Logika background final untuk Dashboard Oscar
 * Update: Fix mapping DetailDowntimeTable (Matrix)
 */

Office.onReady(() => {});

async function populateDashboard(event) {
  try {
    await Excel.run(async (context) => {
      
      // --- 1. DEFINISI SHEET & TABEL ---
      const sheetDash = context.workbook.worksheets.getItemOrNullObject("Dash Oscar");
      const sheetShiftly = context.workbook.worksheets.getItemOrNullObject("Input Shiftly");
      const sheetDowntime = context.workbook.worksheets.getItemOrNullObject("Input Downtime");

      // Tabel Utama
      const tblMain = sheetShiftly.tables.getItemOrNullObject("TableLaporanAkhir");
      
      // Tabel Matrix (Mesin 1-13) -> DetailDowntimeTable
      const tblMatrix = sheetDowntime.tables.getItemOrNullObject("DetailDowntimeTable");
      
      // Tabel List (Rincian Kejadian) -> DowntimeTable
      const tblDetailList = sheetDowntime.tables.getItemOrNullObject("DowntimeTable");

      // Load properti
      sheetDash.load("isNullObject");
      tblMain.load("isNullObject");
      tblMatrix.load("isNullObject");
      tblDetailList.load("isNullObject");

      await context.sync();

      // --- SET NOTIFIKASI LOADING ---
      const statusCell = sheetDash.getRange("AG2");
      statusCell.values = [["⏳ Memuat..."]];
      statusCell.format.font.color = "blue";
      await context.sync();

      // --- 2. AMBIL ID DARI DASHBOARD (AG1) ---
      const searchRange = sheetDash.getRange("AG1");
      searchRange.load("values");
      await context.sync();

      // Pastikan ID bersih dari spasi
      const searchID = String(searchRange.values[0][0]).trim();

      if (!searchID) {
        statusCell.values = [["⚠️ ID Kosong"]];
        return;
      }

      // --- 3. LOAD DATA (BATCHING) ---
      // Load Main Table
      const rangeMainHead = tblMain.getHeaderRowRange().load("values");
      const rangeMainBody = tblMain.getDataBodyRange().load("values");

      // Load Matrix Table
      let rangeMatrixHead = null, rangeMatrixBody = null;
      if (!tblMatrix.isNullObject) {
        rangeMatrixHead = tblMatrix.getHeaderRowRange().load("values");
        rangeMatrixBody = tblMatrix.getDataBodyRange().load("values");
      }

      // Load Detail List Table
      let rangeDetailHead = null, rangeDetailBody = null;
      if (!tblDetailList.isNullObject) {
        rangeDetailHead = tblDetailList.getHeaderRowRange().load("values");
        rangeDetailBody = tblDetailList.getDataBodyRange().load("values");
      }

      await context.sync();

      // --- 4. HELPER FUNCTIONS ---
      // Membuat Map: "NamaKolom" -> Index (0, 1, 2...)
      function createColMap(headers) {
        let map = {};
        for (let i = 0; i < headers.length; i++) {
          // Simpan dengan Huruf Besar & Trim agar aman
          map[String(headers[i]).trim().toUpperCase()] = i;
        }
        return map;
      }
      
      // Ambil Nilai dari Baris berdasarkan Nama Kolom
      function getVal(row, map, colName) {
        const idx = map[colName.toUpperCase()]; // Cari pakai Huruf Besar
        if (idx === undefined) return ""; // Jika kolom tidak ada, return kosong
        return (row[idx] !== null) ? row[idx] : "";
      }

      // --- 5. PROSES TABEL UTAMA (TableLaporanAkhir) ---
      const headersMain = rangeMainHead.values[0];
      const bodyMain = rangeMainBody.values;
      const mapMain = createColMap(headersMain);
      const idxSourceMain = mapMain["SOURCE"];

      // Cari Baris ID
      let rowMain = null;
      for (let i = 0; i < bodyMain.length; i++) {
        if (String(bodyMain[i][idxSourceMain]).trim() === searchID) {
          rowMain = bodyMain[i];
          break;
        }
      }

      if (!rowMain) {
        statusCell.values = [["❌ ID TIDAK DITEMUKAN"]];
        statusCell.format.font.color = "red";
        await context.sync();
        return;
      }

      // Tulis Header ke Dashboard
      sheetDash.getRange("K1").values = [[getVal(rowMain, mapMain, "DATE")]];
      sheetDash.getRange("N1").values = [[getVal(rowMain, mapMain, "SHIFT(1)")]];
      sheetDash.getRange("E1").values = [[getVal(rowMain, mapMain, "HARI")]];
      sheetDash.getRange("S1").values = [[getVal(rowMain, mapMain, "LEADER")]];
      sheetDash.getRange("R6").values = [[getVal(rowMain, mapMain, "TEAM")]];
      sheetDash.getRange("AB1").values = [[getVal(rowMain, mapMain, "SPV")]];
      sheetDash.getRange("K2").values = [[getVal(rowMain, mapMain, "LINE")]];
      sheetDash.getRange("N2").values = [[getVal(rowMain, mapMain, "SKU NAME")]];
      sheetDash.getRange("S2").values = [[getVal(rowMain, mapMain, "TARGET OEE")]];
      sheetDash.getRange("Q23").values = [[getVal(rowMain, mapMain, "NO SO")]];
      sheetDash.getRange("AD91").values = [[getVal(rowMain, mapMain, "START")]];
      sheetDash.getRange("AD92").values = [[getVal(rowMain, mapMain, "FINISH")]];
      sheetDash.getRange("AA75").values = [[getVal(rowMain, mapMain, "ISI 1 DUS")]];
      sheetDash.getRange("F6").values = [[getVal(rowMain, mapMain, "PLAN")]];
      sheetDash.getRange("M23").values = [[getVal(rowMain, mapMain, "TOTAL QUALITY")]];
      sheetDash.getRange("O23").values = [[getVal(rowMain, mapMain, "TOTAL SAFETY")]];
      sheetDash.getRange("AA74").values = [[getVal(rowMain, mapMain, "SPEED / JAM")]];

      // Tulis Data Per Jam (Loop 1-10)
      const targetRowsMain = [10, 11, 12, 13, 15, 16, 17, 19, 20, 21];
      let hourRanges = [];

      for (let i = 1; i <= 10; i++) {
        let r = targetRowsMain[i-1];
        let hVal = getVal(rowMain, mapMain, `HOUR(${i})`);
        
        sheetDash.getRange("B" + r).values = [[hVal]];
        sheetDash.getRange("H" + r).values = [[getVal(rowMain, mapMain, `ACTUAL(${i})`)]];
        sheetDash.getRange("M" + r).values = [[getVal(rowMain, mapMain, `QUALITY(${i})`)]];
        sheetDash.getRange("O" + r).values = [[getVal(rowMain, mapMain, `SAFETY(${i})`)]];
        sheetDash.getRange("U" + r).values = [[getVal(rowMain, mapMain, `WASTE(${i})`)]];
        sheetDash.getRange("D" + r).values = [[getVal(rowMain, mapMain, `STANDART(${i})`)]];

        hourRanges.push(parseTimeRange(hVal));
      }

      // Tulis Waste Tambahan
      sheetDash.getRange("X10").values = [[getVal(rowMain, mapMain, "WASTE(11)")]];
      sheetDash.getRange("X13").values = [[getVal(rowMain, mapMain, "WASTE(12)")]];
      sheetDash.getRange("X15").values = [[getVal(rowMain, mapMain, "WASTE(13)")]];
      sheetDash.getRange("X17").values = [[getVal(rowMain, mapMain, "WASTE(14)")]];
      sheetDash.getRange("X19").values = [[getVal(rowMain, mapMain, "WASTE(15)")]];


      // =========================================================
      // BAGIAN II: MATRIX DOWNTIME (DetailDowntimeTable)
      // =========================================================
      if (!tblMatrix.isNullObject && rangeMatrixBody) {
        
        const headersMatrix = rangeMatrixHead.values[0];
        const bodyMatrix = rangeMatrixBody.values;
        const mapMatrix = createColMap(headersMatrix);
        const idxSourceMatrix = mapMatrix["SOURCE"];

        // Cari Baris ID di Tabel Matrix
        let rowMatrix = null;
        for (let i = 0; i < bodyMatrix.length; i++) {
          if (String(bodyMatrix[i][idxSourceMatrix]).trim() === searchID) {
            rowMatrix = bodyMatrix[i];
            break;
          }
        }

        if (rowMatrix) {
          // --- MAPPING BARIS TUJUAN (Hardcoded sesuai Request) ---
          const grp1Rows = [7, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21];
          const grp2Rows = [30, 31, 33, 35, 37, 39, 40, 41, 43, 45, 46, 47, 48];
          const grp3Rows = [55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67];

          // Loop 13 Mesin (Machine1 s/d Machine13)
          for (let m = 1; m <= 13; m++) {
            let idx = m - 1; // Array index mulai 0
            
            // GROUP 1: Unscheduled Loss -> Ke Kolom Z, AA, AB, AC
            // Mengambil kolom: MACHINE1, UTND1, CIMOH1, NPT1 (Otomatis Uppercase di fungsi getVal)
            let r1 = grp1Rows[idx];
            sheetDash.getRange("Z" + r1).values  = [[getVal(rowMatrix, mapMatrix, "MACHINE" + m)]];
            sheetDash.getRange("AA" + r1).values = [[getVal(rowMatrix, mapMatrix, "UTND" + m)]];
            sheetDash.getRange("AB" + r1).values = [[getVal(rowMatrix, mapMatrix, "CIMOH" + m)]];
            sheetDash.getRange("AC" + r1).values = [[getVal(rowMatrix, mapMatrix, "NPT" + m)]];

            // GROUP 2: Planned Downtime -> Ke Kolom AA, AB, AC, AD
            // Mengambil kolom: PM1, PS1, PCO1, BM1
            let r2 = grp2Rows[idx];
            sheetDash.getRange("AA" + r2).values = [[getVal(rowMatrix, mapMatrix, "PM" + m)]];
            sheetDash.getRange("AB" + r2).values = [[getVal(rowMatrix, mapMatrix, "PS" + m)]];
            sheetDash.getRange("AC" + r2).values = [[getVal(rowMatrix, mapMatrix, "PCO" + m)]];
            sheetDash.getRange("AD" + r2).values = [[getVal(rowMatrix, mapMatrix, "BM" + m)]];

            // GROUP 3: Unplanned Downtime -> Ke Kolom AA, AB, AC, AD, AE
            // Mengambil kolom: OLPS1, EQFB1, LOG1, PRL1, QUAL1
            let r3 = grp3Rows[idx];
            sheetDash.getRange("AA" + r3).values = [[getVal(rowMatrix, mapMatrix, "OLPS" + m)]];
            sheetDash.getRange("AB" + r3).values = [[getVal(rowMatrix, mapMatrix, "EQFB" + m)]];
            sheetDash.getRange("AC" + r3).values = [[getVal(rowMatrix, mapMatrix, "LOG" + m)]];
            sheetDash.getRange("AD" + r3).values = [[getVal(rowMatrix, mapMatrix, "PRL" + m)]];
            sheetDash.getRange("AE" + r3).values = [[getVal(rowMatrix, mapMatrix, "QUAL" + m)]];
          }
        }
      }

      // =========================================================
      // BAGIAN III: DETAIL DOWNTIME (DowntimeTable List)
      // =========================================================
      if (!tblDetailList.isNullObject && rangeDetailBody) {
        const headersDetail = rangeDetailHead.values[0];
        const bodyDetail = rangeDetailBody.values;
        const mapDetail = createColMap(headersDetail);
        const idxSourceDetail = mapDetail["SOURCE"];

        // Filter baris yang ID-nya cocok
        const matchingRows = bodyDetail.filter(r => String(r[idxSourceDetail]).trim() === searchID);

        // Siapkan 10 "Keranjang" (Buckets) untuk Jam 1-10
        let buckets = [];
        for(let i=0; i<10; i++) buckets.push({ F: [], P: [], U: [], W: [] });

        matchingRows.forEach(row => {
          let startVal = getVal(row, mapDetail, "START");
          let timeDec = 0;

          // Normalisasi Time (Excel Serial -> Decimal)
          if (typeof startVal === 'number') {
             timeDec = (startVal - Math.floor(startVal)) * 24;
          }

          // Cek Jam Kejadian masuk ke Keranjang mana (1-10)
          let foundBucketIdx = -1;
          for (let i = 0; i < 10; i++) {
            let range = hourRanges[i]; // Dari tabel utama
            if (range && timeDec >= range.start && timeDec < range.end) {
              foundBucketIdx = i;
              break;
            }
          }

          // Jika ketemu jamnya, masukkan datanya ke keranjang
          if (foundBucketIdx > -1) {
            let b = buckets[foundBucketIdx];
            let mach = getVal(row, mapDetail, "MACHINE");
            let desc = getVal(row, mapDetail, "DESCRIPTION");
            let dur = getVal(row, mapDetail, "DURASI");
            let act = getVal(row, mapDetail, "ACTION");
            let pic = getVal(row, mapDetail, "PIC");
            let stat = getVal(row, mapDetail, "STATUS");

            // Format String: "Mesin: Masalah (Durasi)"
            b.F.push(`${mach}: ${desc} (${dur})`);
            
            if (act && act !== "NONE") b.P.push(act);
            if (pic && pic !== "NONE") b.U.push(pic);
            if (stat && stat !== "NONE") b.W.push(stat);
          }
        });

        // Tulis ke Dashboard (Baris 59, 62, 65...)
        const dtTargetRows = [59, 62, 65, 67, 69, 71, 73, 75, 77, 79];
        
        for(let i=0; i<10; i++) {
          let r = dtTargetRows[i];
          let b = buckets[i];

          // Gabungkan data jika ada lebih dari 1 kejadian di jam yang sama
          let valF = b.F.length > 0 ? b.F.join(", ") : "NONE";
          let valP = b.P.length > 0 ? b.P.join(", ") : "NONE";
          let valU = b.U.length > 0 ? [...new Set(b.U)].join(" & ") : "NONE"; // Hapus duplikat nama
          let valW = b.W.length > 0 ? [...new Set(b.W)].join(" & ") : "NONE"; 

          sheetDash.getRange("F" + r).values = [[valF]];
          sheetDash.getRange("P" + r).values = [[valP]];
          sheetDash.getRange("U" + r).values = [[valU]];
          sheetDash.getRange("W" + r).values = [[valW]];
        }
      }

      // --- SELESAI ---
      statusCell.values = [["✅ DATA BERHASIL"]];
      statusCell.format.font.color = "green";
      statusCell.format.font.bold = true;
      sheetDash.getRange("K2").select();
      
      await context.sync();

    });
  } catch (error) {
    console.error("Error populateDashboard: " + error);
  } finally {
    if (event) event.completed();
  }
}

// Helper: Parse Range Jam ("07.00 - 08.00") jadi angka desimal
function parseTimeRange(rangeStr) {
  if (!rangeStr || typeof rangeStr !== 'string' || rangeStr.indexOf("-") === -1) {
    return null;
  }
  try {
    let parts = rangeStr.split("-");
    let startStr = parts[0].trim().replace(".", ":");
    let endStr = parts[1].trim().replace(".", ":");

    let start = timeStrToDecimal(startStr);
    let end = timeStrToDecimal(endStr);

    if (end < start) end += 24; 

    return { start: start, end: end };
  } catch (e) {
    return null;
  }
}

function timeStrToDecimal(tStr) {
  let p = tStr.split(":");
  let h = parseInt(p[0]);
  let m = p.length > 1 ? parseInt(p[1]) : 0;
  return h + (m / 60);
}

Office.actions.associate("populateDashboard", populateDashboard);