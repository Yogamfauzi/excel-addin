/*
 * File: commands.js
 * Fungsi: Menangani logika tombol Ribbon yang berjalan di background (tanpa UI).
 */

Office.onReady(() => {
  // Office siap digunakan
});

/**
 * Fungsi: populateDashboard
 * Deskripsi: Membaca ID dari sheet 'Dash Oscar' (AG1), mencari data di 'Input Shiftly',
 * lalu mengisi sel K2, S1, dan N1 di 'Dash Oscar'.
 */
async function populateDashboard(event) {
  try {
    await Excel.run(async (context) => {
      
      // 1. Definisikan Sheet & Tabel
      const sheetDash = context.workbook.worksheets.getItemOrNullObject("Dash Oscar");
      const sheetSource = context.workbook.worksheets.getItemOrNullObject("Input Shiftly");
      const tableData = sheetSource.tables.getItemOrNullObject("TableLaporanAkhir");

      // Load properti isNullObject untuk pengecekan
      sheetDash.load("isNullObject");
      sheetSource.load("isNullObject");
      tableData.load("isNullObject");

      await context.sync();

      // 2. Validasi Ketersediaan Sheet & Tabel
      if (sheetDash.isNullObject) {
        console.log("Error: Sheet 'Dash Oscar' tidak ditemukan.");
        return;
      }
      if (sheetSource.isNullObject || tableData.isNullObject) {
        console.log("Error: Sheet 'Input Shiftly' atau 'TableLaporanAkhir' tidak ditemukan.");
        return;
      }

      // 3. Ambil ID Pencarian dari Cell AG1 di Dashboard
      const searchRange = sheetDash.getRange("AG1");
      searchRange.load("values");
      await context.sync();

      const searchID = searchRange.values[0][0]; // Nilai di AG1

      if (!searchID || searchID.toString().trim() === "") {
        console.log("Info: Cell AG1 kosong. Harap isi ID terlebih dahulu.");
        return;
      }

      // 4. Ambil Data Header & Body dari Tabel Sumber
      const headerRange = tableData.getHeaderRowRange().load("values");
      const bodyRange = tableData.getDataBodyRange().load("values");
      await context.sync();

      const headers = headerRange.values[0]; // Array Header (Baris 1)
      const body = bodyRange.values;         // Array Data (Baris-baris data)

      // 5. Peta Kolom (Mapping Index Kolom berdasarkan Nama Header)
      let colMap = {};
      for (let i = 0; i < headers.length; i++) {
        // Simpan nama header dalam huruf besar agar tidak sensitif case
        colMap[String(headers[i]).trim().toUpperCase()] = i;
      }

      // Cari Index kolom yang dibutuhkan
      const idxSource = colMap["SOURCE"];
      const idxLine = colMap["LINE"];
      const idxLeader = colMap["LEADER"];
      
      // Coba cari kolom "SHIFT" atau "SHIFT(1)" atau "SHIFT (1)"
      let idxShift = colMap["SHIFT"];
      if (idxShift === undefined) idxShift = colMap["SHIFT(1)"];
      if (idxShift === undefined) idxShift = colMap["SHIFT (1)"];

      // Validasi jika kolom Source (ID) tidak ketemu
      if (idxSource === undefined) {
        console.log("Error: Kolom 'Source' tidak ditemukan di tabel sumber.");
        return;
      }

      // 6. Loop Pencarian Data (Matching ID)
      let foundRow = null;
      for (let i = 0; i < body.length; i++) {
        // Bandingkan ID (Convert ke string & trim biar aman)
        if (String(body[i][idxSource]).trim() === String(searchID).trim()) {
          foundRow = body[i];
          break; // Stop loop jika ketemu
        }
      }

      // 7. Jika Data Ditemukan, Tulis ke Dashboard
      if (foundRow) {
        // Ambil value (jika index kolom valid)
        const valLine = (idxLine !== undefined) ? foundRow[idxLine] : "";
        const valLeader = (idxLeader !== undefined) ? foundRow[idxLeader] : "";
        const valShift = (idxShift !== undefined) ? foundRow[idxShift] : "";

        // Tulis ke sel tujuan
        sheetDash.getRange("K2").values = [[valLine]];   // Line -> K2
        sheetDash.getRange("S1").values = [[valLeader]]; // Leader -> S1
        sheetDash.getRange("N1").values = [[valShift]];  // Shift -> N1

        // (Opsional) Select cell K2 agar user tahu data berubah
        sheetDash.getRange("K2").select();

        console.log(`Sukses: Data ID ${searchID} berhasil dimuat.`);
      } else {
        console.log(`Info: ID ${searchID} tidak ditemukan di database.`);
        // (Opsional) Kosongkan field jika tidak ketemu
        sheetDash.getRange("K2").values = [[""]];
        sheetDash.getRange("S1").values = [[""]];
        sheetDash.getRange("N1").values = [[""]];
      }

      await context.sync();
    });
  } catch (error) {
    console.error("Error di populateDashboard: " + error);
  } finally {
    // Wajib dipanggil untuk memberitahu Excel bahwa proses selesai
    if (event) {
      event.completed();
    }
  }
}

// Mendaftarkan fungsi ke Office.actions agar bisa dipanggil dari Manifest XML
// Nama string pertama ("populateDashboard") harus SAMA PERSIS dengan <FunctionName> di XML
Office.actions.associate("populateDashboard", populateDashboard);