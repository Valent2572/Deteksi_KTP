function doGet(e) {
  // Ganti dengan ID Spreadsheet Anda yang bisa didapatkan dari URL Spreadsheet
  var sheetId = '1FYM2Th8Tgu-5bIq4Ec9IA8iLNDMfHphy6Z9iveSubPY'; 
  
  // Mengambil sheet pertama (tab paling kiri) secara otomatis
  var sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
  var data = sheet.getDataRange().getValues();
  
  var action = e.parameter.action || 'search';
  var output = ContentService.createTextOutput();
  
  if (action === 'search') {
    var nikToSearch = e.parameter.nik;
    
    if (!nikToSearch) {
      output.setContent(JSON.stringify({
        status: 'error',
        message: 'Parameter NIK tidak diberikan'
      })).setMimeType(ContentService.MimeType.JSON);
      return output;
    }

    var result = null;
    // Looping data mulai dari baris 2 (index 1) untuk melewati header
    for (var i = 1; i < data.length; i++) {
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
  else if (action === 'add') {
    try {
      var lastRow = sheet.getLastRow();
      // Angka NO diset sama dengan jumlah baris saat ini (karena baris 1 header)
      var newNo = lastRow; 
      
      var newRow = [
        newNo,
        e.parameter.resot || '',
        e.parameter.anggota_hari || '',
        e.parameter.tanggal || '',
        e.parameter.nama || '',
        e.parameter.nik || '',
        e.parameter.alamat || '',
        e.parameter.no_hp || '',
        e.parameter.status || ''
      ];
      
      sheet.appendRow(newRow);
      
      output.setContent(JSON.stringify({
        status: 'success',
        message: 'Data peminjam baru berhasil ditambahkan.'
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      output.setContent(JSON.stringify({
        status: 'error',
        message: 'Gagal menambahkan data: ' + err.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return output;
  }
  else if (action === 'edit') {
    try {
      var nikToEdit = e.parameter.nik;
      var edited = false;
      
      for (var i = 1; i < data.length; i++) {
        var sheetNik = String(data[i][5]).replace(/[^0-9]/g, '');
        var searchNik = String(nikToEdit).replace(/[^0-9]/g, '');
        
        if (sheetNik === searchNik && searchNik !== '') {
          var rowToEdit = i + 1; // +1 karena index array mulai dari 0, row google sheet mulai dari 1
          
          if(e.parameter.resot !== undefined) sheet.getRange(rowToEdit, 2).setValue(e.parameter.resot);
          if(e.parameter.anggota_hari !== undefined) sheet.getRange(rowToEdit, 3).setValue(e.parameter.anggota_hari);
          if(e.parameter.tanggal !== undefined) sheet.getRange(rowToEdit, 4).setValue(e.parameter.tanggal);
          if(e.parameter.nama !== undefined) sheet.getRange(rowToEdit, 5).setValue(e.parameter.nama);
          // Kolom 6 adalah NIK, kita tidak ubah karena ini adalah kuncinya.
          if(e.parameter.alamat !== undefined) sheet.getRange(rowToEdit, 7).setValue(e.parameter.alamat);
          if(e.parameter.no_hp !== undefined) sheet.getRange(rowToEdit, 8).setValue(e.parameter.no_hp);
          if(e.parameter.status !== undefined) sheet.getRange(rowToEdit, 9).setValue(e.parameter.status);
          
          edited = true;
          break;
        }
      }
      
      if (edited) {
        output.setContent(JSON.stringify({
          status: 'success',
          message: 'Data peminjam berhasil diperbarui.'
        })).setMimeType(ContentService.MimeType.JSON);
      } else {
        output.setContent(JSON.stringify({
          status: 'error',
          message: 'Data NIK tidak ditemukan untuk diedit.'
        })).setMimeType(ContentService.MimeType.JSON);
      }
    } catch (err) {
      output.setContent(JSON.stringify({
        status: 'error',
        message: 'Gagal mengedit data: ' + err.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return output;
  }
}
