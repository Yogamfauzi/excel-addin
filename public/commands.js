/*
 * File: commands.js
 * Fungsi: Logika background untuk memindahkan data dari TableLaporanAkhir ke Dash Oscar
 * Referensi Logic: VBA Sub PanggilDataKeOscar
 */

Office.onReady(() => {
  // Office siap
});

async function populateDashboard(event) {
  try {
    await Excel.run(async (context) => {
      
      // 1. Setup Sheet & Table
      const sheetDash = context.workbook.worksheets.getItemOrNullObject("Dash Oscar");
      const sheetSource = context.workbook.worksheets.getItemOrNullObject("Input Shiftly");
      const tableData = sheetSource.tables.getItemOrNullObject("TableLaporanAkhir");

      // Load properti untuk validasi
      sheetDash.load("isNullObject");
      sheetSource.load("isNullObject");
      tableData.load("isNullObject");

      await context.sync();

      if (sheetDash.isNullObject || sheetSource.isNullObject || tableData.isNullObject) {
        console.log("Error: Sheet 'Dash Oscar', 'Input Shiftly' atau 'TableLaporanAkhir' tidak ditemukan.");
        return;
      }

      // 2. Ambil ID dari Dash Oscar (AG1)
      const searchRange = sheetDash.getRange("AG1");
      searchRange.load("values");
      await context.sync();

      const searchID = searchRange.values[0][0];

      if (!searchID || searchID.toString().trim() === "") {
        console.log("Info: Cell AG1 kosong.");
        return;
      }

      // 3. Ambil Data Sumber (Header & Body)
      const headerRange = tableData.getHeaderRowRange().load("values");
      const bodyRange = tableData.getDataBodyRange().load("values");
      await context.sync();

      const headers = headerRange.values[0];
      const body = bodyRange.values;

      // 4. Mapping Index Kolom (Case Insensitive)
      // Helper untuk mencari index kolom berdasarkan nama
      let colMap = {};
      for (let i = 0; i < headers.length; i++) {
        colMap[String(headers[i]).trim().toUpperCase()] = i;
      }

      // --- DEFINISI NAMA KOLOM (SESUAI VBA CONST) ---
      const idxSource = colMap["SOURCE"];
      
      // Validasi Source
      if (idxSource === undefined) {
        console.log("Error: Kolom 'Source' tidak ditemukan.");
        return;
      }

      // 5. Cari Baris Data (Looping)
      let foundRow = null;
      for (let i = 0; i < body.length; i++) {
        // Konversi ke string & trim untuk pencarian yang akurat
        if (String(body[i][idxSource]).trim() === String(searchID).trim()) {
          foundRow = body[i];
          break;
        }
      }

      // Helper function untuk mengambil nilai dari foundRow secara aman
      function getVal(colName) {
        const idx = colMap[colName.toUpperCase()];
        return (idx !== undefined && foundRow[idx] !== null) ? foundRow[idx] : "";
      }

      // 6. Tulis ke Dashboard jika data ditemukan
      if (foundRow) {
        console.log(`Data ditemukan untuk ID: ${searchID}`);

        // --- BAGIAN I: HEADER DATA ---
        // Mapping sesuai VBA & Request Anda
        
        // K1: Date
        let valDate = getVal("DATE");
        // Excel JS mengembalikan tanggal sebagai serial number (int), biarkan Excel formatting yang handle
        sheetDash.getRange("K1").values = [[valDate]]; 

        sheetDash.getRange("N1").values = [[getVal("SHIFT(1)")]];
        sheetDash.getRange("E1").values = [[getVal("HARI")]];
        sheetDash.getRange("S1").values = [[getVal("LEADER")]];
        sheetDash.getRange("R6").values = [[getVal("TEAM")]];
        sheetDash.getRange("AB1").values = [[getVal("SPV")]];
        sheetDash.getRange("K2").values = [[getVal("LINE")]];
        sheetDash.getRange("N2").values = [[getVal("SKU NAME")]]; // Request: SKU Name
        sheetDash.getRange("S2").values = [[getVal("TARGET OEE")]];
        
        sheetDash.getRange("Q23").values = [[getVal("NO SO")]];
        sheetDash.getRange("AD91").values = [[getVal("START")]];
        sheetDash.getRange("AD92").values = [[getVal("FINISH")]];
        sheetDash.getRange("AA75").values = [[getVal("ISI 1 DUS")]];
        sheetDash.getRange("F6").values = [[getVal("PLAN")]];
        sheetDash.getRange("M23").values = [[getVal("TOTAL QUALITY")]];
        sheetDash.getRange("O23").values = [[getVal("TOTAL SAFETY")]];
        
        // Request Tambahan
        sheetDash.getRange("AA74").values = [[getVal("SPEED / JAM")]]; 

        // --- BAGIAN II: DATA PER JAM (LOOP 1-10) ---
        // Sesuai VBA: targetRows = Array(10, 11, 12, 13, 15, 16, 17, 19, 20, 21)
        const targetRows = [10, 11, 12, 13, 15, 16, 17, 19, 20, 21];

        // Kita siapkan array range untuk mempercepat penulisan (batching per kolom)
        // Namun demi kesederhanaan logika VBA, kita tulis per cell (aman untuk volume data kecil ini)
        
        for (let i = 1; i <= 10; i++) {
            let rowNum = targetRows[i-1]; // Array js mulai dari 0
            
            // Kolom B: Hour(i)
            sheetDash.getRange("B" + rowNum).values = [[getVal(`HOUR(${i})`)]];
            // Kolom H: Actual(i)
            sheetDash.getRange("H" + rowNum).values = [[getVal(`ACTUAL(${i})`)]];
            // Kolom M: Quality(i)
            sheetDash.getRange("M" + rowNum).values = [[getVal(`QUALITY(${i})`)]];
            // Kolom O: Safety(i)
            sheetDash.getRange("O" + rowNum).values = [[getVal(`SAFETY(${i})`)]];
            // Kolom U: Waste(i)
            sheetDash.getRange("U" + rowNum).values = [[getVal(`WASTE(${i})`)]];
            // Kolom D: Standart(i)
            sheetDash.getRange("D" + rowNum).values = [[getVal(`STANDART(${i})`)]];
        }

        // --- BAGIAN III: WASTE TAMBAHAN (11-15) ---
        // Sesuai VBA
        sheetDash.getRange("X10").values = [[getVal("WASTE(11)")]];
        sheetDash.getRange("X13").values = [[getVal("WASTE(12)")]];
        sheetDash.getRange("X15").values = [[getVal("WASTE(13)")]];
        sheetDash.getRange("X17").values = [[getVal("WASTE(14)")]];
        sheetDash.getRange("X19").values = [[getVal("WASTE(15)")]];

        // Selesai, arahkan kursor ke K2
        sheetDash.getRange("K2").select();
        
      } else {
        console.log("ID tidak ditemukan.");
        // Kosongkan field utama saja sebagai indikator
        const blank = [[""]];
        sheetDash.getRange("K2").values = blank; // Line
        sheetDash.getRange("N2").values = blank; // SKU
        sheetDash.getRange("F6").values = blank; // Plan
        // Anda bisa menambahkan logika pengosongan cell lain jika diperlukan
      }

      await context.sync();
    });
  } catch (error) {
    console.error("Error populateDashboard: " + error);
  } finally {
    if (event) event.completed();
  }
}

Office.actions.associate("populateDashboard", populateDashboard);