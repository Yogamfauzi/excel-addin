/*
 * File: clear.js
 * Fungsi: Membersihkan data pada halaman Dashboard (Dash Oscar)
 * Cakupan: Header, Hourly Data, Matrix Downtime, Detail List, Reject
 */

Office.onReady(() => {});

async function clearDashboard(event) {
  try {
    await Excel.run(async (context) => {
      
      // --- 1. DEFINISI SHEET ---
      const sheetDash = context.workbook.worksheets.getItemOrNullObject("Dash Oscar");
      sheetDash.load("isNullObject");
      await context.sync();

      if (sheetDash.isNullObject) {
        console.error("Sheet 'Dash Oscar' tidak ditemukan.");
        return;
      }

      // --- SET STATUS AGAR USER TAHU SEDANG PROSES ---
      const statusCell = sheetDash.getRange("AG2");
      statusCell.values = [["🧹 Cleaning..."]];
      statusCell.format.font.color = "orange";
      await context.sync();

      // =========================================================
      // BAGIAN I: HEADER UTAMA (Single Cells)
      // =========================================================
      // Daftar sel header sesuai populateDashboard
      const headerCells = [
        "K1", "N1", "E1", "S1", "R6", "AB1", // Baris 1 & Leader/Team
        "K2", "N2", "S2",                    // Baris 2
        "Q23", "M23", "O23",                 // Info SO & Totals
        "AD91", "AD92",                      // Start/Finish Time
        "AA75", "AA74",                      // Isi 1 Dus & Speed
        "F6"                                 // Plan
      ];

      // Hapus konten header sekaligus
      // Menggunakan getRanges (pisahkan koma) untuk efisiensi
      sheetDash.getRange(headerCells.join(",")).clear(Excel.ClearApplyTo.contents);

      // =========================================================
      // BAGIAN II: DATA PER JAM (Loop 1-10) & WASTE EXTRA
      // =========================================================
      const targetRowsMain = [10, 11, 12, 13, 15, 16, 17, 19, 20, 21];
      
      // Kita kumpulkan alamat sel untuk Hourly Data
      let hourlyAddresses = [];
      targetRowsMain.forEach(r => {
        // Kolom B, H, M, O, U, D
        hourlyAddresses.push(`B${r}`);
        hourlyAddresses.push(`H${r}`);
        hourlyAddresses.push(`M${r}`);
        hourlyAddresses.push(`O${r}`);
        hourlyAddresses.push(`U${r}`);
        hourlyAddresses.push(`D${r}`);
      });

      // Tambahan Waste Khusus (X10, X13, dll)
      const wasteExtra = ["X10", "X13", "X15", "X17", "X19"];
      
      // Gabungkan dan hapus
      const allHourly = hourlyAddresses.concat(wasteExtra);
      // Batasi panjang string range jika terlalu banyak, tapi untuk ini masih aman
      sheetDash.getRange(allHourly.join(",")).clear(Excel.ClearApplyTo.contents);

      // =========================================================
      // BAGIAN III: MATRIX DOWNTIME (DetailDowntimeTable)
      // =========================================================
      // Baris-baris sesuai grup di commands.js
      const grp1Rows = [7, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21];
      const grp2Rows = [30, 31, 33, 35, 37, 39, 40, 41, 43, 45, 46, 47, 48];
      const grp3Rows = [55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67];

      let matrixRanges = [];

      // Group 1: Cols Z, AA, AB, AC
      grp1Rows.forEach(r => matrixRanges.push(`Z${r}:AC${r}`));

      // Group 2: Cols AA, AB, AC, AD
      grp2Rows.forEach(r => matrixRanges.push(`AA${r}:AD${r}`));

      // Group 3: Cols AA, AB, AC, AD, AE
      grp3Rows.forEach(r => matrixRanges.push(`AA${r}:AE${r}`));

      // Eksekusi hapus Matrix (gunakan getRanges untuk multi-area)
      // Note: getRange dengan koma kadang punya limit panjang string,
      // kita loop delete per area agar lebih aman jika range sangat banyak
      if(matrixRanges.length > 0) {
         // Cara cepat: sheetDash.getRange(matrixRanges.join(",")).clear(...)
         // Tapi untuk keamanan, kita loop batching atau per item:
         matrixRanges.forEach(address => {
             sheetDash.getRange(address).clear(Excel.ClearApplyTo.contents);
         });
      }

      // =========================================================
      // BAGIAN IV: DETAIL DOWNTIME LIST (Deskripsi, Action, PIC, Status)
      // =========================================================
      const dtTargetRows = [59, 62, 65, 67, 69, 71, 73, 75, 77, 79];
      let detailListAddresses = [];

      dtTargetRows.forEach(r => {
        detailListAddresses.push(`F${r}`); // Desc (Gabungan)
        detailListAddresses.push(`P${r}`); // Action
        detailListAddresses.push(`U${r}`); // PIC
        detailListAddresses.push(`W${r}`); // Status
      });

      sheetDash.getRange(detailListAddresses.join(",")).clear(Excel.ClearApplyTo.contents);

      // =========================================================
      // BAGIAN V: DATA REJECT (IsiRejectTable)
      // =========================================================
      // Kolom target Reject
      const rejectCols = ["E", "H", "K", "L", "N", "Q", "R", "S", "T", "W", "AB", "AD"];
      
      let rejectAddresses = [];
      rejectCols.forEach(col => {
        // Baris 113 (Nama) dan 114 (Jumlah)
        rejectAddresses.push(`${col}113`);
        rejectAddresses.push(`${col}114`);
      });

      sheetDash.getRange(rejectAddresses.join(",")).clear(Excel.ClearApplyTo.contents);

      // =========================================================
      // OPSIONAL: HAPUS INPUT SEARCH ID (AG1) JUGA?
      // Hapus tanda komentar di bawah jika ingin kolom search (AG1) ikut bersih
      // =========================================================
      
      // sheetDash.getRange("AG1").clear(Excel.ClearApplyTo.contents); 


      // --- SELESAI ---
      statusCell.values = [["✨ READY"]]; // Status kembali normal
      statusCell.format.font.color = "black";
      statusCell.format.font.bold = false;
      
      // Kembalikan kursor ke AG1 agar siap scan/ketik lagi
      sheetDash.getRange("AG1").select();

      await context.sync();

    });
  } catch (error) {
    console.error("Error clearDashboard: " + error);
  } finally {
    if (event) event.completed();
  }
}

// Pastikan mendaftarkan fungsi ini di manifest.xml atau taskpane HTML
Office.actions.associate("clearDashboard", clearDashboard);