function doGet(e) {
  // Ganti dengan ID Spreadsheet Anda yang bisa didapatkan dari URL Spreadsheet
  var sheetId = '1lrk1JFZwxoxiwfjOyoMsGQ65tvp5swtIHjkMU-Vd6s4'; 
  
  // Mengambil sheet pertama (tab paling kiri) secara otomatis
  var sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  
  var nikToSearch = e.parameter.nik;
  
  // CORS Helper
  var output = ContentService.createTextOutput();
  
  if (!nikToSearch) {
    output.setContent(JSON.stringify({
      status: 'error',
      message: 'Parameter NIK tidak diberikan'
    })).setMimeType(ContentService.MimeType.JSON);
    return output;
  }

  // Header row diasumsikan berada pada index 0
  // Urutan kolom: RESOT, ANGGOTA HARI, TANGGAL, NAMA, NIK, ALAMAT, NO HP, STATUS
  
  var result = null;
  // Looping data mulai dari baris 2 (index 1) untuk melewati header
  for (var i = 1; i < data.length; i++) {
    // Berdasarkan gambar:
    // Kolom A (NO) = 0
    // Kolom B (RESOT) = 1
    // Kolom C (ANGGOTA HARI) = 2
    // Kolom D (TANGGAL) = 3
    // Kolom E (NAMA) = 4
    // Kolom F (NIK) = 5
    // Kolom G (ALAMAT) = 6
    // Kolom H (NO HP) = 7
    // Kolom I (STATUS) = 8
    
    // Bersihkan NIK dari karakter aneh (spasi, tanda kutip, dll) agar lebih akurat
    var sheetNik = String(data[i][5]).replace(/[^0-9]/g, '');
    var searchNik = String(nikToSearch).replace(/[^0-9]/g, '');

    // NIK berada di kolom ke-6 (index 5)
    if (sheetNik === searchNik && searchNik !== '') {
      result = {
        resot: data[i][1],
        anggota_hari: data[i][2],
        tanggal: data[i][3],
        nama: data[i][4],
        nik: data[i][5],
        alamat: data[i][6],
        no_hp: data[i][7],
        status: data[i][8]
      };
      break;
    }
  }

  if (result) {
    output.setContent(JSON.stringify({
      status: 'success',
      data: result
    })).setMimeType(ContentService.MimeType.JSON);
  } else {
    output.setContent(JSON.stringify({
      status: 'not_found',
      message: 'Data NIK tidak ditemukan di database.'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  return output;
}
