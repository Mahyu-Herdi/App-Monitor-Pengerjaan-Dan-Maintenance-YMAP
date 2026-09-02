const SHEET_ID = '1J0Qb5VNOlLs3_s98REMyL5tyRBqf6dZxrPAP5xcLV8M'; 
const FOLDER_ID = '1a71rnxUYlqUzAtQGEVsdvjrAGNko79Qk';

// FUNGSI INIT JIKA SHEET BELUM SIAP
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  if (!ss.getSheetByName('Monitoring')) {
    ss.insertSheet('Monitoring').appendRow([
      'noUrut', 'tglMulai', 'namaDapur', 'pic', 'anggaranAwal', 'dibayar', 'sisa', 'status',
      'fotoAwal', 'fotoSedang', 'fotoSelesai', 'fotoBon', 'fotoKwitansi', 'invoiceUrl',
      'tglTerbayar', 'sumberDana', 'fotoStrukBayar', 'bankNama', 'bankRek', 'bankAn', 'judulProyek'
    ]);
  }
  if (!ss.getSheetByName('Data_Pesanan')) {
    ss.insertSheet('Data_Pesanan').appendRow([
      'noUrut', 'tglMulai', 'tglSelesai', 'namaDapur', 'pic', 'namaBarang', 'qty', 'satuan', 'hargaSatuan', 'total'
    ]);
  }
  if (!ss.getSheetByName('Pengaturan')) {
    ss.insertSheet('Pengaturan').appendRow([
      'lokasi', 'jabatan', 'bankNama', 'bankRek', 'bankAn', 'kopUrl'
    ]);
  }
}

// ── GET DATA UNTUK SINKRONISASI ALL DEVICE ──
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    initSheets();
    
    const sheetMonitor = ss.getSheetByName('Monitoring');
    const sheetPesanan = ss.getSheetByName('Data_Pesanan');
    const sheetSetting = ss.getSheetByName('Pengaturan');

    // ── Monitoring ──
    const dataMonitor = sheetMonitor.getDataRange().getValues();
    let jsonMonitor = [];
    
    for (let i = 1; i < dataMonitor.length; i++) {
      if(dataMonitor[i][0] === "") continue;
      
      let tglMulai = dataMonitor[i][1];
      if(tglMulai instanceof Date) {
        tglMulai = Utilities.formatDate(tglMulai, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        tglMulai = String(tglMulai || "");
      }
      
      jsonMonitor.push({
        noUrut: dataMonitor[i][0],
        tglMulai: tglMulai,
        namaDapur: dataMonitor[i][2],
        pic: dataMonitor[i][3],
        anggaranAwal: dataMonitor[i][4],
        dibayar: dataMonitor[i][5],
        sisa: dataMonitor[i][6],
        status: dataMonitor[i][7],
        fotoAwal: dataMonitor[i][8] ? String(dataMonitor[i][8]).split(" \n").filter(Boolean) : [],
        fotoSedang: dataMonitor[i][9] ? String(dataMonitor[i][9]).split(" \n").filter(Boolean) : [],
        fotoSelesai: dataMonitor[i][10] ? String(dataMonitor[i][10]).split(" \n").filter(Boolean) : [],
        fotoBon: dataMonitor[i][11] ? String(dataMonitor[i][11]).split(" \n").filter(Boolean) : [],
        fotoKwitansi: dataMonitor[i][12] ? String(dataMonitor[i][12]).split(" \n").filter(Boolean) : [],
        invoiceUrl: dataMonitor[i][13] || "",
        tglTerbayar: dataMonitor[i][14] ? String(dataMonitor[i][14]).split(" \n").filter(Boolean) : [],
        sumberDana: dataMonitor[i][15] ? String(dataMonitor[i][15]).split(" \n").filter(Boolean) : [],
        fotoStrukBayar: dataMonitor[i][16] ? String(dataMonitor[i][16]).split(" \n").filter(Boolean) : [],
        bankNama: dataMonitor[i][17] || "",
        bankRek: dataMonitor[i][18] || "",
        bankAn: dataMonitor[i][19] || "",
        judulProyek: dataMonitor[i][20] || ""
      });
    }

    // ── Pesanan ──
    const dataPesanan = sheetPesanan.getDataRange().getValues();
    let jsonPesanan = [];
    for (let i = 1; i < dataPesanan.length; i++) {
      if(dataPesanan[i][0] === "") continue;
      let tglMulai = dataPesanan[i][1];
      let tglSelesai = dataPesanan[i][2];
      if(tglMulai instanceof Date) tglMulai = Utilities.formatDate(tglMulai, Session.getScriptTimeZone(), "yyyy-MM-dd");
      if(tglSelesai instanceof Date) tglSelesai = Utilities.formatDate(tglSelesai, Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      jsonPesanan.push({
        noUrut: dataPesanan[i][0],
        tglMulai: String(tglMulai || ""),
        tglSelesai: String(tglSelesai || ""),
        namaDapur: dataPesanan[i][3],
        pic: dataPesanan[i][4],
        namaBarang: dataPesanan[i][5],
        qty: dataPesanan[i][6],
        satuan: dataPesanan[i][7],
        hargaSatuan: dataPesanan[i][8],
        total: dataPesanan[i][9]
      });
    }

    // ── Pengaturan ──
    const dataSetting = sheetSetting.getDataRange().getValues();
    let jsonSetting = {};
    let daftarRekening = [];
    
    if(dataSetting.length > 1) {
      jsonSetting = {
        lok: String(dataSetting[1][0] || ""),
        jab: String(dataSetting[1][1] || ""),
        bnk: String(dataSetting[1][2] || ""),
        rek: String(dataSetting[1][3] || ""),
        an: String(dataSetting[1][4] || ""),
        kop: String(dataSetting[1][5] || "")
      };
      
      for (let i = 1; i < dataSetting.length; i++) {
        if(dataSetting[i][4] === "") continue;
        daftarRekening.push({
          lok: String(dataSetting[i][0] || ""),
          jab: String(dataSetting[i][1] || ""),
          bnk: String(dataSetting[i][2] || ""),
          rek: String(dataSetting[i][3] || ""),
          an: String(dataSetting[i][4] || ""),
          kop: String(dataSetting[i][5] || "")
        });
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({
        pesanan: jsonPesanan,
        monitoring: jsonMonitor,
        pengaturan: jsonSetting,
        daftarRekening: daftarRekening
      })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({status: "Gagal", pesan: err.message})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── POST DATA ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetMonitor = ss.getSheetByName('Monitoring');
    const sheetPesanan = ss.getSheetByName('Data_Pesanan');
    const sheetSetting = ss.getSheetByName('Pengaturan');
    
    // 1. INPUT USULAN BARU
    if(data.action === "simpanUsulan") {
      sheetMonitor.appendRow([ data.noUrut, '', data.namaDapur, '', 0, 0, 0, 'Pending Usulan', '', '', '', '', '', '', '', '', '', '', '', '', data.judulProyek || '' ]);
      data.items.forEach(item => {
        sheetPesanan.appendRow([ data.noUrut, '', '', data.namaDapur, '', item.namaBarang, item.qty, item.satuan, 0, 0 ]);
      });
      return ContentService.createTextOutput(JSON.stringify({status: "Sukses"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 2. PROSES PENETAPAN SP
    if(data.action === "prosesSP") {
      const dbMon = sheetMonitor.getDataRange().getValues();
      for (let i = 1; i < dbMon.length; i++) {
        if(dbMon[i][0] == data.noUrut) {
          sheetMonitor.getRange(i + 1, 2).setValue(data.tglMulai);
          sheetMonitor.getRange(i + 1, 4).setValue(data.pic);
          sheetMonitor.getRange(i + 1, 5).setValue(data.anggaranAwal);
          sheetMonitor.getRange(i + 1, 7).setValue(data.anggaranAwal);
          sheetMonitor.getRange(i + 1, 8).setValue('Dalam Pengerjaan');
          break;
        }
      }
      hapusItemPesanan(sheetPesanan, data.noUrut);
      data.items.forEach(item => {
        sheetPesanan.appendRow([ data.noUrut, data.tglMulai, data.tglSelesai, data.namaDapur, data.pic, item.namaBarang, item.qty, item.satuan, item.hargaSatuan, item.total ]);
      });
      return ContentService.createTextOutput(JSON.stringify({status: "Sukses"})).setMimeType(ContentService.MimeType.JSON);
    }

    // 3. UPLOAD VALIDASI DOKUMEN (OTOMATIS BIKIN SUB FOLDER)
    if(data.action === "uploadValidasi") {
      const dbMon = sheetMonitor.getDataRange().getValues();
      let dDapur = "", dPic = "", dJudul = "";
      let rowIndex = -1;
      
      // Cari data Dapur, PIC, dan Judul dari NoUrut yang dikirim HTML
      for (let i = 1; i < dbMon.length; i++) {
        if(dbMon[i][0] == data.noUrut) {
          rowIndex = i;
          dDapur = dbMon[i][2];
          dPic = dbMon[i][3];
          dJudul = dbMon[i][20];
          break;
        }
      }
      
      // Ambil folder spesifik proyek ini
      let targetFolder = null;
      if (rowIndex !== -1) {
         targetFolder = getTargetFolder(data.noUrut, dDapur, dPic, dJudul);
      } else {
         targetFolder = DriveApp.getFolderById(FOLDER_ID); // fallback
      }

      let lAwal=[], lSedang=[], lSelesai=[], lBon=[], lKwitansi=[];
      // Kirim targetFolder ke uploadKeDrive
      if(data.fAwal) data.fAwal.forEach((b64,i) => lAwal.push(uploadKeDrive(b64, "Awal_"+data.noUrut+"_"+i, targetFolder)));
      if(data.fSedang) data.fSedang.forEach((b64,i) => lSedang.push(uploadKeDrive(b64, "Sedang_"+data.noUrut+"_"+i, targetFolder)));
      if(data.fSelesai) data.fSelesai.forEach((b64,i) => lSelesai.push(uploadKeDrive(b64, "Selesai_"+data.noUrut+"_"+i, targetFolder)));
      if(data.fBon) data.fBon.forEach((b64,i) => lBon.push(uploadKeDrive(b64, "Bon_"+data.noUrut+"_"+i, targetFolder)));
      if(data.fKwitansi) data.fKwitansi.forEach((b64,i) => lKwitansi.push(uploadKeDrive(b64, "Kwitansi_"+data.noUrut+"_"+i, targetFolder)));

      if (rowIndex !== -1) {
        if(lAwal.length) sheetMonitor.getRange(rowIndex+1, 9).setValue(gabungString(dbMon[rowIndex][8], lAwal));
        if(lSedang.length) sheetMonitor.getRange(rowIndex+1, 10).setValue(gabungString(dbMon[rowIndex][9], lSedang));
        if(lSelesai.length) sheetMonitor.getRange(rowIndex+1, 11).setValue(gabungString(dbMon[rowIndex][10], lSelesai));
        if(lBon.length) sheetMonitor.getRange(rowIndex+1, 12).setValue(gabungString(dbMon[rowIndex][11], lBon));
        if(lKwitansi.length) sheetMonitor.getRange(rowIndex+1, 13).setValue(gabungString(dbMon[rowIndex][12], lKwitansi));
      }
      return ContentService.createTextOutput(JSON.stringify({status: "Sukses"})).setMimeType(ContentService.MimeType.JSON);
    }

    // 4. CATAT PEMBAYARAN (OTOMATIS STRUK MASUK SUB FOLDER)
    if(data.action === "submitPembayaran") {
      const dbMon = sheetMonitor.getDataRange().getValues();
      let rowIndex = -1;
      let dDapur = "", dPic = "", dJudul = "";
      
      // Sama, cari info proyeknya dulu
      for (let i = 1; i < dbMon.length; i++) {
        if(dbMon[i][0] == data.noUrut) {
          rowIndex = i;
          dDapur = dbMon[i][2];
          dPic = dbMon[i][3];
          dJudul = dbMon[i][20];
          break;
        }
      }

      let urlStruk = "";
      if(data.fotoStrukBase64) {
        let targetFolder = null;
        if(rowIndex !== -1) {
           targetFolder = getTargetFolder(data.noUrut, dDapur, dPic, dJudul);
        }
        urlStruk = uploadKeDrive(data.fotoStrukBase64, "Struk_"+data.noUrut+"_"+Date.now(), targetFolder);
      }

      if (rowIndex !== -1) {
        sheetMonitor.getRange(rowIndex+1, 6).setValue(data.dibayar);
        sheetMonitor.getRange(rowIndex+1, 7).setValue(data.sisa);    
        sheetMonitor.getRange(rowIndex+1, 8).setValue(data.status);  
        sheetMonitor.getRange(rowIndex+1, 15).setValue(gabungString(dbMon[rowIndex][14], [data.tglBaru]));
        sheetMonitor.getRange(rowIndex+1, 16).setValue(gabungString(dbMon[rowIndex][15], [data.danaBaru]));
        if(urlStruk) {
          sheetMonitor.getRange(rowIndex+1, 17).setValue(gabungString(dbMon[rowIndex][16], [urlStruk]));
        }
        sheetMonitor.getRange(rowIndex+1, 18).setValue(data.bankNama || '');
        sheetMonitor.getRange(rowIndex+1, 19).setValue(data.bankRek || '');
        sheetMonitor.getRange(rowIndex+1, 20).setValue(data.bankAn || '');
      }
      return ContentService.createTextOutput(JSON.stringify({status: "Sukses"})).setMimeType(ContentService.MimeType.JSON);
    }

    // 5. UPDATE/EDIT PROYEK
    if(data.action === "updateProyek") {
      const dbMon = sheetMonitor.getDataRange().getValues();
      for (let i = 1; i < dbMon.length; i++) {
        if(dbMon[i][0] == data.noUrut) {
          sheetMonitor.getRange(i+1, 2).setValue(data.tglMulai);
          sheetMonitor.getRange(i+1, 3).setValue(data.namaDapur);
          sheetMonitor.getRange(i+1, 4).setValue(data.pic);
          sheetMonitor.getRange(i+1, 5).setValue(data.anggaranAwal);
          sheetMonitor.getRange(i+1, 6).setValue(data.dibayar);
          sheetMonitor.getRange(i+1, 7).setValue(data.sisa);
          sheetMonitor.getRange(i+1, 8).setValue(data.status);
          sheetMonitor.getRange(i+1, 21).setValue(data.judulProyek || ''); 
          break;
        }
      }
      hapusItemPesanan(sheetPesanan, data.noUrut);
      data.items.forEach(item => {
        sheetPesanan.appendRow([ data.noUrut, data.tglMulai, data.tglSelesai || '', data.namaDapur, data.pic, item.namaBarang, item.qty, item.satuan, item.hargaSatuan, item.total ]);
      });
      return ContentService.createTextOutput(JSON.stringify({status: "Sukses"})).setMimeType(ContentService.MimeType.JSON);
    }

    // 6. SIMPAN PENGATURAN
    if(data.action === "simpanPengaturan") {
      const existing = sheetSetting.getDataRange().getValues();
      if(existing.length > 1) {
        sheetSetting.getRange(2, 1).setValue(data.lok);
        sheetSetting.getRange(2, 2).setValue(data.jab);
        sheetSetting.getRange(2, 3).setValue(data.bnk);
        sheetSetting.getRange(2, 4).setValue(data.rek);
        sheetSetting.getRange(2, 5).setValue(data.an);
        sheetSetting.getRange(2, 6).setValue(data.kop);
      } else {
        sheetSetting.appendRow([data.lok, data.jab, data.bnk, data.rek, data.an, data.kop]);
      }
      return ContentService.createTextOutput(JSON.stringify({status: "Sukses"})).setMimeType(ContentService.MimeType.JSON);
    }

    // 7. TAMBAH REKENING BARU
    if(data.action === "tambahRekening") {
      sheetSetting.appendRow([data.lok, data.jab, data.bnk, data.rek, data.an, data.kop]);
      return ContentService.createTextOutput(JSON.stringify({status: "Sukses"})).setMimeType(ContentService.MimeType.JSON);
    }

    // 8. HAPUS FOTO DARI DATABASE DAN GOOGLE DRIVE
    if(data.action === "hapusFoto") {
      const dbMon = sheetMonitor.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < dbMon.length; i++) {
        if(dbMon[i][0] == data.noUrut) {
          rowIndex = i;
          break;
        }
      }
      if (rowIndex !== -1) {
        // Kolom fotoAwal (9), fotoSedang (10), fotoSelesai (11), fotoBon (12), fotoKwitansi (13), fotoStrukBayar (17)
        const colIndices = [9, 10, 11, 12, 13, 17];
        for (let j = 0; j < colIndices.length; j++) {
          let colIdx = colIndices[j];
          let cellVal = String(dbMon[rowIndex][colIdx - 1] || "");
          let urls = cellVal.split(" \n").filter(Boolean);
          const foundIdx = urls.indexOf(data.fotoUrl);
          if (foundIdx !== -1) {
            urls.splice(foundIdx, 1);
            sheetMonitor.getRange(rowIndex + 1, colIdx).setValue(urls.join(" \n"));
            
            // Hapus file dari Drive
            try {
              const fileId = ambilFileId(data.fotoUrl);
              if (fileId) {
                DriveApp.getFileById(fileId).setTrashed(true);
              }
            } catch(e) {
              // Abaikan jika gagal menghapus file fisik di Drive
            }
            break;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status: "Sukses"})).setMimeType(ContentService.MimeType.JSON);
    }


  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({status: "Gagal", pesan: error.message})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}


// ==========================================
// ── HELPER FUNCTIONS (MODIFIED UTK FOLDER)
// ==========================================

function gabungString(dataLama, arrayBaru) {
  let lama = String(dataLama || "").trim();
  let baru = arrayBaru.join(" \n");
  if (lama === "") return baru;
  return lama + " \n" + baru;
}

function hapusItemPesanan(sheet, noUrut) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] == noUrut) sheet.deleteRow(i + 1);
  }
}

// LOGIKA UPLOAD DIPERBARUI AGAR BISA MENERIMA TARGET FOLDER
function uploadKeDrive(base64Data, namaFile, targetFolder) {
  if(!base64Data || !base64Data.includes(',')) return base64Data;
  
  // Kalau target folder dikirim, pakai itu. Kalau tidak, pakai folder utama.
  const folder = targetFolder ? targetFolder : DriveApp.getFolderById(FOLDER_ID); 
  
  const splitBase = base64Data.split(',');
  const byteCharacters = Utilities.base64Decode(splitBase[1]);
  const blob = Utilities.newBlob(byteCharacters, 'image/jpeg', namaFile + ".jpg");
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ==========================================
// ── FUNGSI PENCIPTA STRUKTUR FOLDER
// ==========================================

function getOrCreateSubFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

function getTargetFolder(noUrut, namaDapur, pic, judulProyek) {
  const root = DriveApp.getFolderById(FOLDER_ID);
  
  // 1. Folder Dapur
  const safeDapur = namaDapur || "Dapur_Tanpa_Nama";
  const folderDapur = getOrCreateSubFolder(root, safeDapur);
  
  // 2. Folder PIC (di dalam folder Dapur)
  const safePic = pic || "PIC_Tanpa_Nama";
  const folderPic = getOrCreateSubFolder(folderDapur, safePic);
  
  // 3. Folder Proyek (di dalam folder PIC)
  const safeJudul = judulProyek || "Proyek_Umum";
  const namaFolderProyek = noUrut + " - " + safeJudul;
  const folderProyek = getOrCreateSubFolder(folderPic, namaFolderProyek);
  
  return folderProyek;
}

function ambilFileId(url) {
  if (!url) return null;
  // Format drive.google.com/open?id=...
  let match = url.match(/id=([^&]+)/);
  if (match) return match[1];
  
  // Format drive.google.com/file/d/...
  match = url.match(/\/file\/d\/([^/]+)/);
  if (match) return match[1];
  
  return null;
}
