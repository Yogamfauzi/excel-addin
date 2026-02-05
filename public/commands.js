/*
 * File: commands.js
 * Fungsi: Logika pencarian dinamis untuk multi-sheet Dashboard (A-G)
 * Deskripsi: Mendeteksi Active Worksheet untuk pengisian data otomatis
 */

Office.onReady(() => {});

async function populateDashboard(event) {
  try {
    await Excel.run(async (context) => {
      
      // --- 1. DETEKSI SHEET AKTIF (Dinamis untuk Dash Oscar A-G) ---
      const sheetDash = context.workbook.worksheets.getActiveWorksheet();
      
      // Referensi Database Statis
      const sheetShiftly = context.workbook.worksheets.getItemOrNullObject("Input Shiftly");
      const sheetDowntime = context.workbook.worksheets.getItemOrNullObject("Input Downtime");

      const tblMain = sheetShiftly.tables.getItemOrNullObject("TableLaporanAkhir");
      const tblMatrix = sheetDowntime.tables.getItemOrNullObject("DetailDowntimeTable");
      const tblDetailList = sheetDowntime.tables.getItemOrNullObject("DowntimeTable");
      const tblReject = sheetDowntime.tables.getItemOrNullObject("IsiRejectTable");

      // Load properti sheet aktif dan tabel
      sheetDash.load("name");
      tblMain.load("isNullObject");
      tblMatrix.load("isNullObject");
      tblDetailList.load("isNullObject");
      tblReject.load("isNullObject");

      await context.sync();

      // --- SET NOTIFIKASI LOADING DI SHEET AKTIF ---
      const statusCell = sheetDash.getRange("AG2");
      statusCell.values = [["⏳ Memuat ke " + sheetDash.name + "..."]];
      statusCell.format.font.color = "blue";
      statusCell.format.font.bold = false;
      await context.sync();

      // --- 2. AMBIL ID DARI DASHBOARD AKTIF (AG1) ---
      const searchRange = sheetDash.getRange("AG1");
      searchRange.load("values");
      await context.sync();

      const searchID = String(searchRange.values[0][0]).trim();

      if (!searchID || searchID === "") {
        statusCell.values = [["⚠️ Masukkan ID di AG1"]];
        statusCell.format.font.color = "orange";
        return;
      }

      // --- 3. LOAD DATA DARI DATABASE ---
      const rangeMainHead = tblMain.getHeaderRowRange().load("values");
      const rangeMainBody = tblMain.getDataBodyRange().load("values");

      let rangeMatrixHead = null, rangeMatrixBody = null;
      if (!tblMatrix.isNullObject) {
        rangeMatrixHead = tblMatrix.getHeaderRowRange().load("values");
        rangeMatrixBody = tblMatrix.getDataBodyRange().load("values");
      }

      let rangeDetailHead = null, rangeDetailBody = null;
      if (!tblDetailList.isNullObject) {
        rangeDetailHead = tblDetailList.getHeaderRowRange().load("values");
        rangeDetailBody = tblDetailList.getDataBodyRange().load("values");
      }

      let rangeRejectHead = null, rangeRejectBody = null;
      if (!tblReject.isNullObject) {
        rangeRejectHead = tblReject.getHeaderRowRange().load("values");
        rangeRejectBody = tblReject.getDataBodyRange().load("values");
      }

      await context.sync();

      // --- 4. HELPER FUNCTIONS ---
      function createColMap(headers) {
        let map = {};
        for (let i = 0; i < headers.length; i++) {
          map[String(headers[i]).trim().toUpperCase()] = i;
        }
        return map;
      }
      
      function getVal(row, map, colName) {
        const idx = map[colName.toUpperCase()];
        if (idx === undefined) return "";
        return (row[idx] !== null && row[idx] !== undefined) ? row[idx] : "";
      }

      // --- 5. PROSES TABEL UTAMA ---
      const headersMain = rangeMainHead.values[0];
      const bodyMain = rangeMainBody.values;
      const mapMain = createColMap(headersMain);
      const idxSourceMain = mapMain["SOURCE"];

      let rowMain = null;
      for (let i = 0; i < bodyMain.length; i++) {
        if (String(bodyMain[i][idxSourceMain]).trim() === searchID) {
          rowMain = bodyMain[i];
          break;
        }
      }

      if (!rowMain) {
        statusCell.values = [["❌ ID " + searchID + " Tidak Ada"]];
        statusCell.format.font.color = "red";
        await context.sync();
        return;
      }

      // Tulis ke Dashboard Aktif
      sheetDash.getRange("K1").values = [[getVal(rowMain, mapMain, "DATE")]];
      sheetDash.getRange("N1").values = [[getVal(rowMain, mapMain, "SHIFT")]];
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
      sheetDash.getRange("U6").values = [[getVal(rowMain, mapMain, "TOTAL JAM")]]; 
      sheetDash.getRange("AA74").values = [[getVal(rowMain, mapMain, "SPEED / JAM")]];

      // Data Per Jam (1-10)
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

      // Waste 11-15
      for (let i = 11; i <= 15; i++) {
        const wasteRows = {11: "X10", 12: "X13", 13: "X15", 14: "X17", 15: "X19"};
        sheetDash.getRange(wasteRows[i]).values = [[getVal(rowMain, mapMain, `WASTE(${i})`)]];
      }

      // --- 6. PROSES MATRIX DOWNTIME ---
      if (!tblMatrix.isNullObject && rangeMatrixBody) {
        const headersMatrix = rangeMatrixHead.values[0];
        const bodyMatrix = rangeMatrixBody.values;
        const mapMatrix = createColMap(headersMatrix);
        const idxSourceMatrix = mapMatrix["SOURCE"];

        let rowMatrix = bodyMatrix.find(r => String(r[idxSourceMatrix]).trim() === searchID);

        if (rowMatrix) {
          const grpRows = {
            1: [7, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21],
            2: [30, 31, 33, 35, 37, 39, 40, 41, 43, 45, 46, 47, 48],
            3: [55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67]
          };

          for (let m = 1; m <= 13; m++) {
            let idx = m - 1;
            sheetDash.getRange("Z" + grpRows[1][idx]).values = [[getVal(rowMatrix, mapMatrix, "MACHINE" + m)]];
            sheetDash.getRange("AA" + grpRows[1][idx]).values = [[getVal(rowMatrix, mapMatrix, "UTND" + m)]];
            sheetDash.getRange("AB" + grpRows[1][idx]).values = [[getVal(rowMatrix, mapMatrix, "CIMOH" + m)]];
            sheetDash.getRange("AC" + grpRows[1][idx]).values = [[getVal(rowMatrix, mapMatrix, "NPT" + m)]];
            sheetDash.getRange("AA" + grpRows[2][idx]).values = [[getVal(rowMatrix, mapMatrix, "PM" + m)]];
            sheetDash.getRange("AB" + grpRows[2][idx]).values = [[getVal(rowMatrix, mapMatrix, "PS" + m)]];
            sheetDash.getRange("AC" + grpRows[2][idx]).values = [[getVal(rowMatrix, mapMatrix, "PCO" + m)]];
            sheetDash.getRange("AD" + grpRows[2][idx]).values = [[getVal(rowMatrix, mapMatrix, "BM" + m)]];
            sheetDash.getRange("AA" + grpRows[3][idx]).values = [[getVal(rowMatrix, mapMatrix, "OLPS" + m)]];
            sheetDash.getRange("AB" + grpRows[3][idx]).values = [[getVal(rowMatrix, mapMatrix, "EQFB" + m)]];
            sheetDash.getRange("AC" + grpRows[3][idx]).values = [[getVal(rowMatrix, mapMatrix, "LOG" + m)]];
            sheetDash.getRange("AD" + grpRows[3][idx]).values = [[getVal(rowMatrix, mapMatrix, "PRL" + m)]];
            sheetDash.getRange("AE" + grpRows[3][idx]).values = [[getVal(rowMatrix, mapMatrix, "QUAL" + m)]];
          }
        }
      }

      // --- 7. PROSES DETAIL DOWNTIME LIST ---
      if (!tblDetailList.isNullObject && rangeDetailBody) {
        const headersDetail = rangeDetailHead.values[0];
        const bodyDetail = rangeDetailBody.values;
        const mapDetail = createColMap(headersDetail);
        const idxSourceDetail = mapDetail["SOURCE"];

        const matchingRows = bodyDetail.filter(r => String(r[idxSourceDetail]).trim() === searchID);
        let buckets = Array.from({length: 10}, () => ({F: [], P: [], U: [], W: []}));

        matchingRows.forEach(row => {
          let startVal = getVal(row, mapDetail, "START");
          let timeDec = (typeof startVal === 'number') ? (startVal - Math.floor(startVal)) * 24 : 0;

          let bIdx = hourRanges.findIndex(r => r && timeDec >= r.start && timeDec < r.end);
          if (bIdx > -1) {
            let b = buckets[bIdx];
            b.F.push(`${getVal(row, mapDetail, "MACHINE")}: ${getVal(row, mapDetail, "DESCRIPTION")} (${getVal(row, mapDetail, "DURASI")})`);
            if (getVal(row, mapDetail, "ACTION") !== "NONE") b.P.push(getVal(row, mapDetail, "ACTION"));
            if (getVal(row, mapDetail, "PIC") !== "NONE") b.U.push(getVal(row, mapDetail, "PIC"));
            if (getVal(row, mapDetail, "STATUS") !== "NONE") b.W.push(getVal(row, mapDetail, "STATUS"));
          }
        });

        const dtRows = [59, 62, 65, 67, 69, 71, 73, 75, 77, 79];
        buckets.forEach((b, i) => {
          sheetDash.getRange("F" + dtRows[i]).values = [[b.F.join(", ") || "NONE"]];
          sheetDash.getRange("P" + dtRows[i]).values = [[b.P.join(", ") || "NONE"]];
          sheetDash.getRange("U" + dtRows[i]).values = [[([...new Set(b.U)]).join(" & ") || "NONE"]];
          sheetDash.getRange("W" + dtRows[i]).values = [[([...new Set(b.W)]).join(" & ") || "NONE"]];
        });
      }

      // --- 8. PROSES REJECT ---
      if (!tblReject.isNullObject && rangeRejectBody) {
        const headersReject = rangeRejectHead.values[0];
        const bodyReject = rangeRejectBody.values;
        const mapReject = createColMap(headersReject);
        const idxSourceReject = mapReject["SOURCE"];

        let rowReject = bodyReject.find(r => String(r[idxSourceReject]).trim() === searchID);
        if (rowReject) {
          const cols = ["E", "H", "K", "L", "N", "Q", "R", "S", "T", "W", "AB", "AD"];
          cols.forEach((c, i) => {
            sheetDash.getRange(c + "113").values = [[getVal(rowReject, mapReject, "REJECT" + (i+1))]];
            sheetDash.getRange(c + "114").values = [[getVal(rowReject, mapReject, "ISI" + (i+1))]];
          });
        }
      }

      // --- SELESAI ---
      statusCell.values = [["✅ BERHASIL DI LOAD"]];
      statusCell.format.font.color = "green";
      statusCell.format.font.bold = true;
      await context.sync();

    });
  } catch (error) {
    console.error(error);
  } finally {
    if (event) event.completed();
  }
}

function parseTimeRange(rangeStr) {
  if (!rangeStr || typeof rangeStr !== 'string' || !rangeStr.includes("-")) return null;
  try {
    let parts = rangeStr.split("-");
    let start = timeStrToDecimal(parts[0].trim().replace(".", ":"));
    let end = timeStrToDecimal(parts[1].trim().replace(".", ":"));
    if (end < start) end += 24; 
    return { start, end };
  } catch (e) { return null; }
}

function timeStrToDecimal(tStr) {
  let p = tStr.split(":");
  return parseInt(p[0]) + (p.length > 1 ? parseInt(p[1])/60 : 0);
}

Office.actions.associate("populateDashboard", populateDashboard);