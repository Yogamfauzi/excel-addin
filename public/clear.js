/*
 * File: clear.js (VERSI INPUT NULL/EMPTY)
 * Fungsi: Menimpa data lama dengan data kosong ("")
 * Kelebihan: Biasanya lebih cepat daripada .clear()
 */

Office.onReady(() => {});

async function clearDashboard(event) {
  try {
    await Excel.run(async (context) => {
      
      const sheetDash = context.workbook.worksheets.getItemOrNullObject("Dash Oscar");
      sheetDash.load("isNullObject");
      await context.sync();

      if (sheetDash.isNullObject) return;

      // Status: Cleaning
      const statusCell = sheetDash.getRange("AG2");
      statusCell.values = [["⚡ Cleaning..."]]; // Ubah icon jadi petir biar beda
      statusCell.format.font.color = "purple";
      await context.sync();

      // --- PERSIAPAN DATA KOSONG ---
      const valEmpty = [[""]];                  // Untuk 1 sel
      const valEmpty4 = [["", "", "", ""]];     // Untuk 4 kolom (Z-AC)
      const valEmpty5 = [["", "", "", "", ""]]; // Untuk 5 kolom (AA-AE)

      // =========================================================
      // 1. TIMPA HEADER (Single Cells)
      // =========================================================
      const headerCells = [
        "K1", "N1", "E1", "S1", "R6", "AB1", 
        "K2", "N2", "S2", 
        "Q23", "M23", "O23", 
        "AD91", "AD92", 
        "AA75", "AA74", 
        "F6"
      ];

      // Loop cepat untuk menimpa dengan ""
      headerCells.forEach(addr => {
        sheetDash.getRange(addr).values = valEmpty;
      });

      // =========================================================
      // 2. TIMPA DATA PER JAM & WASTE
      // =========================================================
      const targetRowsMain = [10, 11, 12, 13, 15, 16, 17, 19, 20, 21];
      const wasteExtra = ["X10", "X13", "X15", "X17", "X19"];

      targetRowsMain.forEach(r => {
        // Kita timpa satu-satu, ini sangat ringan untuk Excel
        sheetDash.getRange(`B${r}`).values = valEmpty;
        sheetDash.getRange(`H${r}`).values = valEmpty;
        sheetDash.getRange(`M${r}`).values = valEmpty;
        sheetDash.getRange(`O${r}`).values = valEmpty;
        sheetDash.getRange(`U${r}`).values = valEmpty;
        sheetDash.getRange(`D${r}`).values = valEmpty;
      });

      wasteExtra.forEach(addr => {
        sheetDash.getRange(addr).values = valEmpty;
      });

      // =========================================================
      // 3. TIMPA MATRIX DOWNTIME (Block Range)
      // =========================================================
      // Kita gunakan array kosong 4 kolom & 5 kolom agar pas ukurannya
      
      const grp1Rows = [7, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21];
      const grp2Rows = [30, 31, 33, 35, 37, 39, 40, 41, 43, 45, 46, 47, 48];
      const grp3Rows = [55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67];

      // Group 1 (Z sampai AC = 4 kolom)
      grp1Rows.forEach(r => {
          sheetDash.getRange(`Z${r}:AC${r}`).values = valEmpty4;
      });

      // Group 2 (AA sampai AD = 4 kolom)
      grp2Rows.forEach(r => {
          sheetDash.getRange(`AA${r}:AD${r}`).values = valEmpty4;
      });

      // Group 3 (AA sampai AE = 5 kolom)
      grp3Rows.forEach(r => {
          sheetDash.getRange(`AA${r}:AE${r}`).values = valEmpty5;
      });

      // =========================================================
      // 4. TIMPA DETAIL LIST
      // =========================================================
      const dtTargetRows = [59, 62, 65, 67, 69, 71, 73, 75, 77, 79];
      dtTargetRows.forEach(r => {
        sheetDash.getRange(`F${r}`).values = valEmpty;
        sheetDash.getRange(`P${r}`).values = valEmpty;
        sheetDash.getRange(`U${r}`).values = valEmpty;
        sheetDash.getRange(`W${r}`).values = valEmpty;
      });

      // =========================================================
      // 5. TIMPA DATA REJECT
      // =========================================================
      const rejectCols = ["E", "H", "K", "L", "N", "Q", "R", "S", "T", "W", "AB", "AD"];
      rejectCols.forEach(col => {
        sheetDash.getRange(`${col}113`).values = valEmpty;
        sheetDash.getRange(`${col}114`).values = valEmpty;
      });

      // --- SELESAI ---
      statusCell.values = [["✨ READY"]];
      statusCell.format.font.color = "black";
      
      // Select kembali ke AG1
      sheetDash.getRange("AG1").select();

      await context.sync();

    });
  } catch (error) {
    console.error("Error: " + error);
    // Jika error, paksa tulis error di cell
     Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem("Dash Oscar");
        s.getRange("AG2").values = [["❌ ERROR"]];
        s.getRange("AG2").format.font.color = "red";
    });
  } finally {
    if (event) event.completed();
  }
}

Office.actions.associate("clearDashboard", clearDashboard);