/*
 * File: commands.js
 * Fungsi: Logika background lengkap dengan Notifikasi Status di Cell AG2
 */

Office.onReady(() => {});

async function populateDashboard(event) {
  try {
    await Excel.run(async (context) => {
      
      // --- 1. DEFINISI SHEET & TABEL ---
      const sheetDash = context.workbook.worksheets.getItemOrNullObject("Dash Oscar");
      const sheetShiftly = context.workbook.worksheets.getItemOrNullObject("Input Shiftly");
      const sheetDowntime = context.workbook.worksheets.getItemOrNullObject("Input Downtime");

      const tblMain = sheetShiftly.tables.getItemOrNullObject("TableLaporanAkhir");
      const tblMatrix = sheetDowntime.tables.getItemOrNullObject("DowntimeTable2");
      const tblDetail = sheetDowntime.tables.getItemOrNullObject("DowntimeTable");

      sheetDash.load("isNullObject");
      sheetShiftly.load("isNullObject");
      sheetDowntime.load("isNullObject");
      tblMain.load("isNullObject");
      tblMatrix.load("isNullObject");
      tblDetail.load("isNullObject");

      await context.sync();

      if (sheetDash.isNullObject || tblMain.isNullObject) {
        console.log("Error: Sheet utama tidak ditemukan.");
        return;
      }

      // --- SET NOTIFIKASI AWAL (LOADING) ---
      const statusCell = sheetDash.getRange("AG2");
      statusCell.values = [["⏳ Memuat..."]];
      statusCell.format.font.color = "blue";
      await context.sync();

      // --- 2. AMBIL ID DARI AG1 ---
      const searchRange = sheetDash.getRange("AG1");
      searchRange.load("values");
      await context.sync();

      const searchID = String(searchRange.values[0][0]).trim();
      if (!searchID) {
        statusCell.values = [["⚠️ ID Kosong"]];
        statusCell.format.font.color = "orange";
        return;
      }

      // --- 3. LOAD DATA ---
      const rangeMainHead = tblMain.getHeaderRowRange().load("values");
      const rangeMainBody = tblMain.getDataBodyRange().load("values");

      let rangeMatrixHead = null, rangeMatrixBody = null;
      if (!tblMatrix.isNullObject) {
        rangeMatrixHead = tblMatrix.getHeaderRowRange().load("values");
        rangeMatrixBody = tblMatrix.getDataBodyRange().load("values");
      }

      let rangeDetailHead = null, rangeDetailBody = null;
      if (!tblDetail.isNullObject) {
        rangeDetailHead = tblDetail.getHeaderRowRange().load("values");
        rangeDetailBody = tblDetail.getDataBodyRange().load("values");
      }

      await context.sync();

      // --- 4. MAPPING & PENCARIAN ---
      function createColMap(headers) {
        let map = {};
        for (let i = 0; i < headers.length; i++) {
          map[String(headers[i]).trim().toUpperCase()] = i;
        }
        return map;
      }

      function getVal(row, map, colName) {
        const idx = map[colName.toUpperCase()];
        return (idx !== undefined && row[idx] !== null) ? row[idx] : "";
      }

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
        // NOTIFIKASI GAGAL
        statusCell.values = [["❌ ID TIDAK DITEMUKAN"]];
        statusCell.format.font.color = "red";
        statusCell.format.font.bold = true;
        await context.sync();
        return;
      }

      // =========================================================
      // BAGIAN I: DATA UTAMA
      // =========================================================
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

      const targetRows = [10, 11, 12, 13, 15, 16, 17, 19, 20, 21];
      let hourRanges = [];

      for (let i = 1; i <= 10; i++) {
        let r = targetRows[i-1];
        let hVal = getVal(rowMain, mapMain, `HOUR(${i})`);
        
        sheetDash.getRange("B" + r).values = [[hVal]];
        sheetDash.getRange("H" + r).values = [[getVal(rowMain, mapMain, `ACTUAL(${i})`)]];
        sheetDash.getRange("M" + r).values = [[getVal(rowMain, mapMain, `QUALITY(${i})`)]];
        sheetDash.getRange("O" + r).values = [[getVal(rowMain, mapMain, `SAFETY(${i})`)]];
        sheetDash.getRange("U" + r).values = [[getVal(rowMain, mapMain, `WASTE(${i})`)]];
        sheetDash.getRange("D" + r).values = [[getVal(rowMain, mapMain, `STANDART(${i})`)]];

        hourRanges.push(parseTimeRange(hVal));
      }

      sheetDash.getRange("X10").values = [[getVal(rowMain, mapMain, "WASTE(11)")]];
      sheetDash.getRange("X13").values = [[getVal(rowMain, mapMain, "WASTE(12)")]];
      sheetDash.getRange("X15").values = [[getVal(rowMain, mapMain, "WASTE(13)")]];
      sheetDash.getRange("X17").values = [[getVal(rowMain, mapMain, "WASTE(14)")]];
      sheetDash.getRange("X19").values = [[getVal(rowMain, mapMain, "WASTE(15)")]];

      // =========================================================
      // BAGIAN II: MATRIX DOWNTIME
      // =========================================================
      if (!tblMatrix.isNullObject && rangeMatrixBody) {
        const headersMatrix = rangeMatrixHead.values[0];
        const bodyMatrix = rangeMatrixBody.values;
        const mapMatrix = createColMap(headersMatrix);
        const idxSourceMatrix = mapMatrix["SOURCE"];

        let rowMatrix = null;
        for (let i = 0; i < bodyMatrix.length; i++) {
          if (String(bodyMatrix[i][idxSourceMatrix]).trim() === searchID) {
            rowMatrix = bodyMatrix[i];
            break;
          }
        }

        if (rowMatrix) {
          const rowsGrp1 = [7, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21];
          const rowsGrp2 = [30, 31, 33, 35, 37, 39, 40, 41, 43, 45, 46, 47, 48];
          const rowsGrp3 = [55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67];

          for (let m = 1; m <= 13; m++) {
            let idx = m - 1;
            sheetDash.getRange("Z" + rowsGrp1[idx]).values = [[getVal(rowMatrix, mapMatrix, "MACHINE" + m)]];
            sheetDash.getRange("AA" + rowsGrp1[idx]).values = [[getVal(rowMatrix, mapMatrix, "UTND" + m)]];
            sheetDash.getRange("AB" + rowsGrp1[idx]).values = [[getVal(rowMatrix, mapMatrix, "CIMOH" + m)]];
            sheetDash.getRange("AC" + rowsGrp1[idx]).values = [[getVal(rowMatrix, mapMatrix, "NPT" + m)]];

            sheetDash.getRange("AA" + rowsGrp2[idx]).values = [[getVal(rowMatrix, mapMatrix, "PM" + m)]];
            sheetDash.getRange("AB" + rowsGrp2[idx]).values = [[getVal(rowMatrix, mapMatrix, "PS" + m)]];
            sheetDash.getRange("AC" + rowsGrp2[idx]).values = [[getVal(rowMatrix, mapMatrix, "PCO" + m)]];
            sheetDash.getRange("AD" + rowsGrp2[idx]).values = [[getVal(rowMatrix, mapMatrix, "BM" + m)]];

            sheetDash.getRange("AA" + rowsGrp3[idx]).values = [[getVal(rowMatrix, mapMatrix, "OLPS" + m)]];
            sheetDash.getRange("AB" + rowsGrp3[idx]).values = [[getVal(rowMatrix, mapMatrix, "EQFB" + m)]];
            sheetDash.getRange("AC" + rowsGrp3[idx]).values = [[getVal(rowMatrix, mapMatrix, "LOG" + m)]];
            sheetDash.getRange("AD" + rowsGrp3[idx]).values = [[getVal(rowMatrix, mapMatrix, "PRL" + m)]];
            sheetDash.getRange("AE" + rowsGrp3[idx]).values = [[getVal(rowMatrix, mapMatrix, "QUAL" + m)]];
          }
        }
      }

      // =========================================================
      // BAGIAN III: DETAIL DOWNTIME
      // =========================================================
      if (!tblDetail.isNullObject && rangeDetailBody) {
        const headersDetail = rangeDetailHead.values[0];
        const bodyDetail = rangeDetailBody.values;
        const mapDetail = createColMap(headersDetail);
        const idxSourceDetail = mapDetail["SOURCE"];

        const matchingRows = bodyDetail.filter(r => String(r[idxSourceDetail]).trim() === searchID);

        let buckets = [];
        for(let i=0; i<10; i++) buckets.push({ F: [], P: [], U: [], W: [] });

        matchingRows.forEach(row => {
          let startVal = getVal(row, mapDetail, "START");
          let timeDec = 0;

          if (typeof startVal === 'number') {
             timeDec = (startVal - Math.floor(startVal)) * 24;
          }

          let foundBucketIdx = -1;
          for (let i = 0; i < 10; i++) {
            let range = hourRanges[i];
            if (range && timeDec >= range.start && timeDec < range.end) {
              foundBucketIdx = i;
              break;
            }
          }

          if (foundBucketIdx > -1) {
            let b = buckets[foundBucketIdx];
            let mach = getVal(row, mapDetail, "MACHINE");
            let desc = getVal(row, mapDetail, "DESCRIPTION");
            let dur = getVal(row, mapDetail, "DURASI");
            let act = getVal(row, mapDetail, "ACTION");
            let pic = getVal(row, mapDetail, "PIC");
            let stat = getVal(row, mapDetail, "STATUS");

            b.F.push(`${mach}: ${desc} (${dur})`);
            if (act && act !== "NONE") b.P.push(act);
            if (pic && pic !== "NONE") b.U.push(pic);
            if (stat && stat !== "NONE") b.W.push(stat);
          }
        });

        const dtTargetRows = [59, 62, 65, 67, 69, 71, 73, 75, 77, 79];
        
        for(let i=0; i<10; i++) {
          let r = dtTargetRows[i];
          let b = buckets[i];

          let valF = b.F.length > 0 ? b.F.join(", ") : "NONE";
          let valP = b.P.length > 0 ? b.P.join(", ") : "NONE";
          let valU = b.U.length > 0 ? [...new Set(b.U)].join(" & ") : "NONE";
          let valW = b.W.length > 0 ? [...new Set(b.W)].join(" & ") : "NONE";

          sheetDash.getRange("F" + r).values = [[valF]];
          sheetDash.getRange("P" + r).values = [[valP]];
          sheetDash.getRange("U" + r).values = [[valU]];
          sheetDash.getRange("W" + r).values = [[valW]];
        }
      }

      // --- NOTIFIKASI SUKSES ---
      statusCell.values = [["✅ DATA BERHASIL DIMUAT"]];
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