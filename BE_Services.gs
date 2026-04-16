// ==============================================================================
// FILE: BE_Services.gs
// TIPE: SERVER-SIDE SCRIPT
// DESKRIPSI: API Endpoints (Routing, Login, Berangkat, dan Absen Pulang)
// UPDATE: Dynamic Status Absensi & Lookup Master_Status_Absensi
// ==============================================================================

function doGet(e) {
  // 1. Logika Verifikasi Publik (Tanpa Login)
  if (e.parameter.verify_st && e.parameter.nrpp) {
    // Kita panggil fungsi render dan tambahkan header bypass
    return renderPublicVerification(e.parameter.verify_st, e.parameter.nrpp)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 2. Logika Dashboard Internal (Tetap Aman)
  return HtmlService.createTemplateFromFile('UI_Base')
    .evaluate()
    .setTitle('C.E.P.U - Enterprise Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- MODUL LOOKUP DATA MASTER ---

function api_getInitFormData(user) {
  try {
    const dbMaster = SpreadsheetApp.openById(DB_MASTER_ID);
    const sheetStatus = dbMaster.getSheetByName("Master_Status_Absensi");
    const dataStatus = sheetStatus.getDataRange().getValues();
    
    let options = [];
    let userLokasi = user.lokasi ? user.lokasi.toString().toUpperCase().trim() : "";
    let userJabatan = user.jabatan ? user.jabatan.toString().toUpperCase().trim() : "";

    for (let i = 1; i < dataStatus.length; i++) {
      let filterArea = dataStatus[i][2] ? dataStatus[i][2].toString().toUpperCase() : "";
      if (filterArea.includes(userLokasi) || filterArea.includes(userJabatan)) {
        options.push({ code: dataStatus[i][0], desc: dataStatus[i][1] });
      }
    }

    // [UPDATE FIX]: Smart Routing DB Sales vs UPD untuk History & Active Trip
    const isSales = (userLokasi === "SALES" || userJabatan === "SALES");
    const dbApp = isSales ? SpreadsheetApp.openById(DB_SALES_ID) : SpreadsheetApp.openById(DB_UPD_ID);
    let targetSheetName = isSales ? "Log_Sales" : "Log_" + user.lokasi.trim();
    let sheetLog = dbApp.getSheetByName(targetSheetName);
    
    let history = [];
    let activeTrip = null; // [NEW]: Penampung Trip Aktif

    if (sheetLog) {
      const dataLog = sheetLog.getDataRange().getValues();
      let stSet = new Set();
      
      for (let i = dataLog.length - 1; i >= 1; i--) {
        if (dataLog[i][1].toString() === user.nrpp.toString()) { 
          
          // [NEW PULL ACTIVE STATE]: Cek apakah ada trip gantung (Sedang Jalan)
          let statusPerjalanan = isSales ? dataLog[i][15] : dataLog[i][17];
          if (!activeTrip && statusPerjalanan === "SEDANG JALAN") {
              let rawTime = dataLog[i][isSales ? 10 : 11];
              let waktuFormat = (rawTime instanceof Date) ? Utilities.formatDate(rawTime, "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss") : rawTime.toString();
              
              activeTrip = {
                  idTransaksi: dataLog[i][0].toString(),
                  waktuKeluar: waktuFormat,
                  lokasi: dataLog[i][isSales ? 9 : 10].toString(),
                  customer: dataLog[i][isSales ? 8 : 9].toString(),
                  statusAbsensi: "SEDANG JALAN" 
              };
          }

          let noSTRaw = isSales ? "" : dataLog[i][8];
          let noST = noSTRaw ? String(noSTRaw).replace(/^'/, '').trim() : ""; 
          
          if (noST !== "" && !stSet.has(noST)) {
            stSet.add(noST);
            history.push({
              noST: noST,
              customer: dataLog[i][isSales ? 8 : 9] ? dataLog[i][isSales ? 8 : 9].toString() : "", 
              lokasi: dataLog[i][isSales ? 9 : 10] ? dataLog[i][isSales ? 9 : 10].toString() : ""   
            });
          }
        }
      }
    }

    return { status: "success", data: { statusOptions: options, historyST: history, activeTrip: activeTrip } };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

// ==========================================================
// MODUL LOGIN & SECURITY: STRICT DEVICE BINDING (ANTI-TITIP ABSEN)
// ==========================================================
function api_verifyLogin(nrpp, password, deviceId) {
  try {
    const dbMaster = SpreadsheetApp.openById(DB_MASTER_ID);
    const sheetLogin = dbMaster.getSheetByName("Master_Login");
    const dataLogin = sheetLogin.getDataRange().getValues();
    const headers = dataLogin[0].map(h => h.toString().toUpperCase().trim());
    
    const idx = {
      nrpp: headers.indexOf("NRPP"),
      pass: headers.indexOf("PASSWORD"),
      devId: headers.indexOf("DEVICE_ID"),
      name: headers.indexOf("NAMA"),
      lastLogin: headers.indexOf("LAST_LOGIN")
    };
    
    if (idx.nrpp === -1 || idx.pass === -1) throw new Error("Struktur kolom Master_Login rusak!");

    let userFound = false;
    let userData = null;

    for (let i = 1; i < dataLogin.length; i++) {
      let row = dataLogin[i];
      
      // Pengecekan NRPP dan Password
      if (row[idx.nrpp].toString() === nrpp.toString() && row[idx.pass].toString() === password.toString()) {
        
        let registeredDeviceId = row[idx.devId] ? row[idx.devId].toString().trim() : "";
        let rowIndex = i + 1;

        // ==========================================================
        // [SECURITY LOCK]: PENCEGAHAN TITIP ABSEN MUTLAK
        // ==========================================================
        if (registeredDeviceId !== "" && registeredDeviceId !== deviceId.toString().trim()) {
            // JIKA ID DATABASE ADA ISINYA, DAN TIDAK SAMA DENGAN HP YANG SEDANG DIPAKAI: BLOKIR!
            return { 
                status: "error", 
                message: "⛔ SECURITY LOCK: Akun Anda sudah terikat di perangkat lain. Dilarang titip absen! Hubungi Admin jika Anda mengganti HP." 
            };
        } else if (registeredDeviceId === "") {
            // JIKA ID KOSONG (HP Baru / Habis Direset Admin): DAFTARKAN HP INI!
            sheetLogin.getRange(rowIndex, idx.devId + 1).setValue(deviceId);
        }
        // ==========================================================

        userFound = true;
        sheetLogin.getRange(rowIndex, idx.lastLogin + 1).setValue(new Date());
        
        const sheetKaryawan = dbMaster.getSheetByName("Master_Karyawan");
        const dataKar = sheetKaryawan.getDataRange().getValues();
        const headKar = dataKar[0].map(h => h.toString().toUpperCase().trim());
        const iK = {
          nrpp: headKar.indexOf("NRPP"),
          jabatan: headKar.indexOf("JABATAN"),
          gol: headKar.indexOf("GOLONGAN"),
          status: headKar.indexOf("STATUS_KARYAWAN"),
          dept: headKar.indexOf("DEPARTEMEN"),
          loc: headKar.indexOf("LOKASI")
        };
        let details = dataKar.find(r => r[iK.nrpp].toString() === nrpp.toString());
        
        userData = {
          nrpp: nrpp,
          nama: row[idx.name],
          jabatan: details ? details[iK.jabatan] : "User",
          golongan: details ? details[iK.gol] : "-",
          statusKaryawan: details ? details[iK.status] : "-",
          departemen: details ? details[iK.dept] : "-",
          lokasi: details ? details[iK.loc] : "OFFICE"
        };
        
        // Pengecekan Hak Akses UPD
        const sheetUPD = dbMaster.getSheetByName("Master_UPD");
        if (sheetUPD) {
            const dataUPD = sheetUPD.getDataRange().getValues();
            let headUPD = dataUPD[0].map(h => h.toString().toUpperCase().trim());
            let idxJabUPD = headUPD.indexOf("JABATAN");
            if(idxJabUPD !== -1) {
                userData.isUpdEligible = dataUPD.some(r => r[idxJabUPD].toString().trim().toUpperCase() === userData.jabatan.toString().trim().toUpperCase());
            } else {
                userData.isUpdEligible = false;
            }
        }
        break;
      }
    }

    if (!userFound) return { status: "error", message: "NRPP atau Password salah!" };
    return { status: "success", data: userData };

  } catch (error) {
    return { status: "error", message: "System Error: " + error.toString() };
  }
}

// --- CARI DAN GANTI FUNGSI INI DI BE_Services.gs ---

function api_submitPerjalananDinas(payload, user) {
  try {
    if (!user.lokasi) throw new Error("Data Lokasi Karyawan kosong di Master Database!");
    const userLokasiUpper = user.lokasi.toString().trim().toUpperCase();
    const isSales = (userLokasiUpper === "SALES" || user.jabatan.toString().trim().toUpperCase() === "SALES");
    
    const dbApp = isSales ? SpreadsheetApp.openById(DB_SALES_ID) : SpreadsheetApp.openById(DB_UPD_ID);
    let targetSheetName = isSales ? "Log_Sales" : "Log_" + user.lokasi.trim(); 
    
    let sheet = dbApp.getSheetByName(targetSheetName);
    if (!sheet) throw new Error("Sheet database tujuan ('" + targetSheetName + "') tidak ditemukan di database!");
    
    const timestamp = new Date();
    const timeToSave = Utilities.formatDate(timestamp, "Asia/Jakarta", "yyyy/MM/dd HH:mm:ss");
    
    const d_wib = Utilities.formatDate(timestamp, "Asia/Jakarta", "dd");
    const m_wib = Utilities.formatDate(timestamp, "Asia/Jakarta", "MM");
    const y_wib = Utilities.formatDate(timestamp, "Asia/Jakarta", "yyyy");

    // ==========================================================
    // PROTOKOL ODOC (One-Day-One-Checkin) - DYNAMIC COLUMN VALIDATOR
    // ==========================================================
    if (user.jabatan !== "Super Admin" && user.jabatan !== "Administrator" && user.jabatan !== "HRD") {
        const dataLog = sheet.getDataRange().getValues();
        
        let nrppIndex = 1;
        let waktuKeluarIndex = isSales ? 10 : 11; // Default Index
        
        if (dataLog.length > 0) {
            let header = dataLog[0];
            for(let c = 0; c < header.length; c++) {
                let colName = header[c].toString().toUpperCase().trim();
                if(colName === "NRPP") nrppIndex = c;
                // [UPDATE]: Mencari Index Kolom Waktu Keluar secara dinamis
                if(colName === "WAKTU_KELUAR" || colName === "WAKTU KELUAR") waktuKeluarIndex = c;
            }
        }

        for (let i = 1; i < dataLog.length; i++) {
            let dbNRPP = String(dataLog[i][nrppIndex]).trim().toUpperCase();
            let reqNRPP = String(user.nrpp).trim().toUpperCase();
            
            let isUserMatch = (dbNRPP === reqNRPP);
            if (!isUserMatch && dbNRPP !== "" && reqNRPP !== "" && !isNaN(dbNRPP) && !isNaN(reqNRPP)) {
                isUserMatch = (Number(dbNRPP) === Number(reqNRPP));
            }
            
            if (isUserMatch) {
                let isAlreadyClockedIn = false;
                // [KUNCI ABSOLUT]: Ekstrak Waktu Langsung dari Kolom Visual (Agar bisa dites Admin)
                let rawTime = dataLog[i][waktuKeluarIndex];
                
                if (rawTime) {
                    let logDate = (rawTime instanceof Date) ? rawTime : new Date(rawTime.toString() + " +0700");
                    
                    if (!isNaN(logDate.getTime()) && logDate.getFullYear() > 2000) {
                        let r_y = Utilities.formatDate(logDate, "Asia/Jakarta", "yyyy");
                        let r_m = Utilities.formatDate(logDate, "Asia/Jakarta", "MM");
                        let r_d = Utilities.formatDate(logDate, "Asia/Jakarta", "dd");
                        
                        // Jika ada data dengan tanggal yang sama seperti hari ini, BLOKIR.
                        if (r_y === y_wib && r_m === m_wib && r_d === d_wib) {
                            isAlreadyClockedIn = true;
                        }
                    } else if (typeof rawTime === "string" || typeof rawTime === "number") {
                        // Fallback Plan: String Scanner
                        let timeStr = String(rawTime);
                        let d_nz = parseInt(d_wib, 10).toString();
                        let m_nz = parseInt(m_wib, 10).toString();
                        const checkFormats = [
                            `${y_wib}/${m_wib}/${d_wib}`, `${d_wib}/${m_wib}/${y_wib}`,
                            `${y_wib}-${m_wib}-${d_wib}`, `${d_wib}-${m_wib}-${y_wib}`,
                            `${d_nz}/${m_nz}/${y_wib}`, `${m_nz}/${d_nz}/${y_wib}`, `${y_wib}/${m_nz}/${d_nz}`
                        ];
                        if (checkFormats.some(fmt => timeStr.includes(fmt))) {
                            isAlreadyClockedIn = true;
                        }
                    }
                }

                if (isAlreadyClockedIn) {
                    return { status: "error", message: "⛔ FRAUD ALERT: Anda sudah melakukan absensi keberangkatan hari ini. (Limit 1x/Hari)" };
                }
            }
        }
    }
    // ==========================================================

    const idTransaksi = "TRX-" + timestamp.getTime();
    const statusAbsensi = payload.statusAbsensi || "H";

    // ==========================================================
    // [NEW] PROTOKOL VALIDASI RADIUS LATLONG & HAVERSINE ENGINE
    // ==========================================================
    if (userLokasiUpper === "RFMC" && statusAbsensi !== "BKF" && statusAbsensi !== "BKS") {
        
        const dbMaster = SpreadsheetApp.openById(DB_MASTER_ID);
        const sheetLatlong = dbMaster.getSheetByName("Master_Latlong");
        
        if (sheetLatlong) {
            const dataLatlong = sheetLatlong.getDataRange().getValues();
            let targetLat = null, targetLon = null, maxRadius = 0, isLatlongActive = false;

            // Memindai konfigurasi untuk lokasi RFMC
            for (let i = 1; i < dataLatlong.length; i++) {
                if (dataLatlong[i][0].toString().trim().toUpperCase() === "RFMC") {
                    targetLat = parseFloat(dataLatlong[i][1]);
                    targetLon = parseFloat(dataLatlong[i][2]);
                    maxRadius = parseFloat(dataLatlong[i][3] || 50); // Default 50 meter jika kosong
                    isLatlongActive = dataLatlong[i][4].toString().trim().toUpperCase() === "AKTIF";
                    break;
                }
            }

            // Eksekusi Kalkulasi Jarak jika mode Latlong diaktifkan oleh Admin
            if (isLatlongActive && targetLat !== null && targetLon !== null) {
                const userKordinat = payload.kordinat.toString().split(",");
                
                if (userKordinat.length === 2) {
                    const userLat = parseFloat(userKordinat[0].trim());
                    const userLon = parseFloat(userKordinat[1].trim());

                    // Haversine Formula (Menghitung jarak lengkung bumi dalam Meter)
                    const R = 6371e3; // Radius bumi (Meter)
                    const rad = Math.PI / 180;
                    const dLat = (targetLat - userLat) * rad;
                    const dLon = (targetLon - userLon) * rad;
                    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
                              Math.cos(userLat * rad) * Math.cos(targetLat * rad) *
                              Math.sin(dLon / 2) * Math.sin(dLon / 2);
                    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
                    const distance = R * c;

                    // Blokir jika jarak melebihi batas radius yang diizinkan
                    if (distance > maxRadius) {
                        return { status: "error", message: `⛔ FRAUD ALERT: Posisi Anda (${distance.toFixed(0)} Meter) berada di luar radius POS Security RFMC (Maksimal: ${maxRadius} Meter)!` };
                    }
                } else {
                    return { status: "error", message: "Gagal memindai kordinat GPS atau format tidak valid!" };
                }
            }
        }
    }
    // ==========================================================

    let rowData = [];
    if (isSales) {
        rowData = [
          idTransaksi, user.nrpp, user.nama, user.jabatan, user.golongan, user.statusKaryawan, user.departemen, user.lokasi,
          payload.customer, payload.lokasi, timeToSave, payload.kordinat, "", "", "", "SEDANG JALAN"
        ];
    } else {
        rowData = [
          idTransaksi, user.nrpp, user.nama, user.jabatan, user.golongan, user.statusKaryawan, user.departemen, user.lokasi, 
          ("'" + payload.noST), payload.customer, payload.lokasi, timeToSave, payload.kordinat, "", "", "", "", "SEDANG JALAN", "BELUM KLAIM", ""
        ];
    }
    
    sheet.appendRow(rowData);

    // [AUTO-GRID DB_REKAP]
    try {
      const dbRekap = SpreadsheetApp.openById(DB_REKAP_ID);
      let rekapSheetName = "";
      if (isSales) rekapSheetName = "Rekap_Absensi_Sales";
      else if (userLokasiUpper === "FMC") rekapSheetName = "Rekap_Absesnsi_FMC";
      else if (userLokasiUpper === "SATELITE") rekapSheetName = "Rekap_Absesnsi_Satelite";
      else rekapSheetName = "Rekap_Absensi_" + user.lokasi.trim();

      const rekapSheet = dbRekap.getSheetByName(rekapSheetName);
      if (rekapSheet) {
        const todayStr = Utilities.formatDate(timestamp, "Asia/Jakarta", "dd/MM/yyyy");
        const timeStr = Utilities.formatDate(timestamp, "Asia/Jakarta", "HH:mm");
        const dataRekap = rekapSheet.getDataRange().getValues();
        let targetRow = -1; let targetCol = -1;
        
        for (let r = 1; r < dataRekap.length; r++) {
          if (dataRekap[r][0].toString() === user.nrpp.toString()) { targetRow = r + 1; break; }
        }
        
        if (targetRow === -1) {
          targetRow = rekapSheet.getLastRow() + 1;
          if (targetRow < 3) targetRow = 3; 
          rekapSheet.getRange(targetRow, 1).setValue(user.nrpp);
          rekapSheet.getRange(targetRow, 2).setValue(user.nama);
          if (rekapSheet.getRange("A1").getValue() === "") {
             rekapSheet.getRange("A1:A2").merge().setValue("NRPP").setBackground("#000000").setFontColor("#FFFFFF").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
             rekapSheet.getRange("B1:B2").merge().setValue("Nama").setBackground("#000000").setFontColor("#FFFFFF").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
          }
        }

        let headerRow = dataRekap.length > 0 ? dataRekap[0] : [];
        for (let c = 2; c < headerRow.length; c += 4) {
          let cellDateStr = (headerRow[c] instanceof Date) ? Utilities.formatDate(headerRow[c], "Asia/Jakarta", "dd/MM/yyyy") : headerRow[c].toString();
          if (cellDateStr.includes(todayStr)) { targetCol = c + 1; break; }
        }

        if (targetCol === -1) {
          targetCol = rekapSheet.getLastColumn() + 1;
          if (targetCol < 3) targetCol = 3;
          rekapSheet.getRange(1, targetCol).setValue(todayStr);
          rekapSheet.getRange(1, targetCol, 1, 4).mergeAcross().setBackground("#000000").setFontColor("#FFFFFF").setHorizontalAlignment("center").setFontWeight("bold");
          rekapSheet.getRange(2, targetCol).setValue("IN");
          rekapSheet.getRange(2, targetCol + 1).setValue("OUT");
          rekapSheet.getRange(2, targetCol + 2).setValue("STATUS");
          rekapSheet.getRange(2, targetCol + 3).setValue("DURASI");
          rekapSheet.getRange(2, targetCol, 1, 4).setBackground("#000000").setFontColor("#FFFFFF").setHorizontalAlignment("center").setFontWeight("bold");
        }

        rekapSheet.getRange(targetRow, targetCol).setValue(timeStr); 
        rekapSheet.getRange(targetRow, targetCol + 2).setValue(statusAbsensi);
      }
    } catch(e) { console.error("Error DB_REKAP IN: " + e.message); }

    return {
      status: "success",
      message: "Keberangkatan Berhasil. Status: " + statusAbsensi,
      data: { idTransaksi: idTransaksi, waktuKeluar: timestamp.toLocaleString('id-ID'), lokasi: payload.lokasi, customer: payload.customer, statusAbsensi: statusAbsensi }
    };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

function api_submitPulangDinas(payload, user) {
  try {
    if (!user.lokasi) throw new Error("Data Lokasi Karyawan kosong!");
    const userLokasiUpper = user.lokasi.toString().trim().toUpperCase();
    const isSales = (userLokasiUpper === "SALES" || user.jabatan.toString().trim().toUpperCase() === "SALES");
    
    const dbApp = isSales ? SpreadsheetApp.openById(DB_SALES_ID) : SpreadsheetApp.openById(DB_UPD_ID);
    let targetSheetName = isSales ? "Log_Sales" : "Log_" + user.lokasi.trim(); 
    
    const sheet = dbApp.getSheetByName(targetSheetName);
    if (!sheet) throw new Error("Sheet tujuan tidak ditemukan.");

    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    let waktuKeluar = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === payload.idTransaksi) {
        targetRow = i + 1;
        let rawKeluar = data[i][isSales ? 10 : 11]; 
        
        // [UPDATE BUG FIX] Memaksa JavaScript membaca String sebagai Zona Waktu Jakarta (+0700)
        waktuKeluar = (rawKeluar instanceof Date) ? rawKeluar : new Date(rawKeluar.toString() + " +0700");
        
        // [KILLER BUG FIX]: Koreksi jika Admin edit jam manual di G-Sheets (Tahun menjadi 1899)
        if (waktuKeluar.getFullYear() < 2000) {
            const today = new Date();
            waktuKeluar.setFullYear(today.getFullYear(), today.getMonth(), today.getDate());
        }
        
        break;
      }
    }

    if (targetRow === -1) throw new Error("ID Transaksi tidak ditemukan.");

    const waktuMasuk = new Date();
    const timeMasukToSave = Utilities.formatDate(waktuMasuk, "Asia/Jakarta", "yyyy/MM/dd HH:mm:ss");
    
    const diffMs = waktuMasuk - waktuKeluar;
    const durasiJam = diffMs / (1000 * 60 * 60);

    let nominalUPD = 0;
    if (!isSales) {
        const dbMaster = SpreadsheetApp.openById(DB_MASTER_ID);
        const masterUPD = dbMaster.getSheetByName("Master_UPD").getDataRange().getValues();
        
        let baseUPD = 0, uangMakan = 0, mknSiangLibur = 0, lainKerja = 0, lainLibur = 0;
        let keyJabatan = user.jabatan ? user.jabatan.toString().trim().toUpperCase() : "";
        let keyGolongan = user.golongan ? user.golongan.toString().trim().toUpperCase() : "";
        let keyStatusKaryawan = user.statusKaryawan ? user.statusKaryawan.toString().trim().toUpperCase() : ""; 

        for (let i = 1; i < masterUPD.length; i++) {
          let dbJabatan = masterUPD[i][0] ? masterUPD[i][0].toString().trim().toUpperCase() : "";
          let dbGolongan = masterUPD[i][1] ? masterUPD[i][1].toString().trim().toUpperCase() : "";
          let dbStatusJabatan = masterUPD[i][2] ? masterUPD[i][2].toString().trim().toUpperCase() : ""; 

          // [STRICT MATCHER LOGIC]: Anti Overlap "NON" vs "PROJECT"
          let isStatusMatch = false;
          if (dbStatusJabatan === "") {
              isStatusMatch = true; 
          } else if (dbStatusJabatan === "PROJECT") {
              if (keyStatusKaryawan.includes("PROJECT") && !keyStatusKaryawan.includes("NON")) isStatusMatch = true;
          } else if (dbStatusJabatan === "NON PROJECT") {
              if (keyStatusKaryawan.includes("NON") || !keyStatusKaryawan.includes("PROJECT")) isStatusMatch = true;
          } else if (keyStatusKaryawan.includes(dbStatusJabatan)) {
              isStatusMatch = true;
          }

          if (dbJabatan === keyJabatan && dbGolongan === keyGolongan && isStatusMatch) {
            baseUPD = (durasiJam >= 8) ? parseFloat(masterUPD[i][3] || 0) : parseFloat(masterUPD[i][4] || 0); 
            uangMakan = parseFloat(masterUPD[i][5] || 0);
            mknSiangLibur = parseFloat(masterUPD[i][6] || 0);
            lainKerja = parseFloat(masterUPD[i][7] || 0);
            lainLibur = parseFloat(masterUPD[i][8] || 0);
            break;
          }
        }

        const isWeekend = (waktuMasuk.getDay() === 0 || waktuMasuk.getDay() === 6);
        if (isWeekend) nominalUPD = baseUPD + mknSiangLibur + uangMakan + lainLibur;
        else nominalUPD = baseUPD + uangMakan + lainKerja;

        sheet.getRange(targetRow, 14).setValue(timeMasukToSave);
        sheet.getRange(targetRow, 15).setValue(payload.kordinatMasuk);
        sheet.getRange(targetRow, 16).setValue(durasiJam.toFixed(2));
        sheet.getRange(targetRow, 17).setValue(nominalUPD);
        sheet.getRange(targetRow, 18).setValue("SELESAI");
    } else {
        sheet.getRange(targetRow, 13).setValue(timeMasukToSave);
        sheet.getRange(targetRow, 14).setValue(payload.kordinatMasuk);
        sheet.getRange(targetRow, 15).setValue(durasiJam.toFixed(2));
        sheet.getRange(targetRow, 16).setValue("SELESAI");
    }

    // [AUTO-GRID DB_REKAP]
    try {
      const dbRekap = SpreadsheetApp.openById(DB_REKAP_ID);
      let rekapSheetName = "";
      if (isSales) rekapSheetName = "Rekap_Absensi_Sales";
      else if (userLokasiUpper === "FMC") rekapSheetName = "Rekap_Absesnsi_FMC";
      else if (userLokasiUpper === "SATELITE") rekapSheetName = "Rekap_Absesnsi_Satelite";
      else rekapSheetName = "Rekap_Absensi_" + user.lokasi.trim();

      const rekapSheet = dbRekap.getSheetByName(rekapSheetName);
      if (rekapSheet) {
        const todayStr = Utilities.formatDate(waktuMasuk, "Asia/Jakarta", "dd/MM/yyyy");
        const timeStr = Utilities.formatDate(waktuMasuk, "Asia/Jakarta", "HH:mm");
        const dataRekap = rekapSheet.getDataRange().getValues();
        let tRow = -1; let tCol = -1;
        
        for (let r = 1; r < dataRekap.length; r++) {
          if (dataRekap[r][0].toString() === user.nrpp.toString()) { tRow = r + 1; break; }
        }

        let headerRow = dataRekap.length > 0 ? dataRekap[0] : [];
        for (let c = 2; c < headerRow.length; c += 4) {
          let cellDateStr = (headerRow[c] instanceof Date) ? Utilities.formatDate(headerRow[c], "Asia/Jakarta", "dd/MM/yyyy") : headerRow[c].toString();
          if (cellDateStr.includes(todayStr)) { tCol = c + 1; break; }
        }

        if (tRow !== -1 && tCol !== -1) {
          rekapSheet.getRange(tRow, tCol + 1).setValue(timeStr);
          rekapSheet.getRange(tRow, tCol + 3).setValue(durasiJam.toFixed(2)); 
        }
      }
    } catch(e) { console.error("Error DB_REKAP OUT: " + e.message); }

    return { 
        status: "success", 
        message: "Selesai!", 
        data: { durasi: durasiJam.toFixed(2), nominal: nominalUPD } 
    };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

// ==========================================================
// MODUL SUPER ADMIN: UNIVERSAL CRUD MASTER DATA
// ==========================================================


function api_adminGetMaster(sheetName) {
  try {
    const db = SpreadsheetApp.openById(DB_MASTER_ID);
    const sheet = db.getSheetByName(sheetName);
    if (!sheet) throw new Error("Tabel Master '" + sheetName + "' tidak ditemukan!");

    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return { status: "success", data: { headers: [], rows: [] } };

    // [KILLER BUG FIX]: Memaksa seluruh elemen sel menjadi teks untuk mencegah JSON Serialization Crash dari google.script.run
    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1).map(row => {
        return row.map(cell => {
            if (cell instanceof Date) return Utilities.formatDate(cell, "Asia/Jakarta", "yyyy/MM/dd HH:mm:ss");
            if (cell === null || cell === undefined) return "";
            return String(cell); 
        });
    });

    return {
      status: "success",
      data: { headers: headers, rows: rows } 
    };
  } catch(e) { 
    return { status: "error", message: e.toString() }; 
  }
}

function api_adminMutateMaster(action, sheetName, payload, rowIndex) {
  try {
    // [DEBUG LOGGING]
    console.log("=== API MUTATE EXECUTED ===");
    console.log("Action:", action);
    console.log("Sheet:", sheetName);
    console.log("Payload:", payload);
    console.log("RowIndex:", rowIndex);

    const db = SpreadsheetApp.openById(DB_MASTER_ID);
    const sheet = db.getSheetByName(sheetName);
    if (!sheet) throw new Error("Tabel Master '" + sheetName + "' tidak ditemukan!");

    if (action === "CREATE") {
      sheet.appendRow(payload);
      
      // Protokol Auto-Login
      if (sheetName === "Master_Karyawan") {
          const loginSheet = db.getSheetByName("Master_Login");
          if (loginSheet) {
              loginSheet.appendRow([payload[0], payload[1], payload[0], "", ""]);
              console.log("Auto-Login Created for NRPP:", payload[0]);
          } else {
              console.warn("Sheet Master_Login tidak ditemukan untuk Auto-Login!");
          }
      }

      return { status: "success", message: "Entitas baru berhasil direkam ke " + sheetName };
    } 
    else if (action === "UPDATE") {
      const targetRow = parseInt(rowIndex) + 2; 
      if(isNaN(targetRow)) throw new Error("Target baris tidak valid!");
      
      sheet.getRange(targetRow, 1, 1, payload.length).setValues([payload]);
      return { status: "success", message: "Entitas pada baris " + targetRow + " berhasil diperbarui!" };
    } 
    else if (action === "DELETE") {
      const targetRow = parseInt(rowIndex) + 2;
      if(isNaN(targetRow)) throw new Error("Target baris tidak valid!");
      
      sheet.deleteRow(targetRow);
      return { status: "success", message: "Entitas berhasil dimusnahkan." };
    } 
    else {
      throw new Error("Protokol Mutasi (CRUD) tidak dikenali: " + action);
    }
  } catch(e) { 
    console.error("Backend Error:", e.toString());
    return { status: "error", message: e.toString() }; 
  }
}

// ==========================================================
// MODUL SUPER ADMIN: SMART HRIS & LIVE TRACKING
// ==========================================================

function api_adminGetKaryawanDetails() {
  try {
    const db = SpreadsheetApp.openById(DB_MASTER_ID);
    const karSheet = db.getSheetByName("Master_Karyawan");
    const logSheet = db.getSheetByName("Master_Login");
    
    const karData = karSheet.getDataRange().getValues();
    const logData = logSheet.getDataRange().getValues();

    // Pemetaan data Login untuk digabungkan dengan Master_Karyawan (Relational Mapping)
    let logMap = {};
    for(let i = 1; i < logData.length; i++) {
        logMap[logData[i][0].toString()] = {
            deviceId: logData[i][3] ? logData[i][3].toString() : "",
            lastLogin: logData[i][4] instanceof Date ? Utilities.formatDate(logData[i][4], "Asia/Jakarta", "dd/MM/yyyy HH:mm") : logData[i][4].toString()
        };
    }

    let results = [];
    for(let i = 1; i < karData.length; i++) {
        let nrpp = karData[i][0].toString();
        if(!nrpp) continue;
        results.push({
            nrpp: nrpp, nama: karData[i][1], jabatan: karData[i][2],
            golongan: karData[i][3], departemen: karData[i][5], lokasi: karData[i][6],
            deviceId: logMap[nrpp] ? logMap[nrpp].deviceId : "",
            lastLogin: logMap[nrpp] ? logMap[nrpp].lastLogin : "Belum Pernah Login"
        });
    }
    return { status: "success", data: results };
  } catch(e) { return { status: "error", message: e.toString() }; }
}

function api_adminResetDevice(nrpp) {
  try {
    const db = SpreadsheetApp.openById(DB_MASTER_ID);
    const sheet = db.getSheetByName("Master_Login");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === nrpp.toString()) {
            sheet.getRange(i + 1, 4).setValue(""); // Kolom ke-4 adalah Device_ID
            return { status: "success", message: "Security Lock untuk NRPP " + nrpp + " berhasil dihancurkan!" };
        }
    }
    throw new Error("Data Kredensial tidak ditemukan!");
  } catch(e) { return { status: "error", message: e.toString() }; }
}

// ==========================================================
// MODUL SUPER ADMIN: GLOBAL LIVE TRACKING (RADAR AKTIF)
// ==========================================================
function api_adminGetLiveLogs() {
  try {
    const timestamp = new Date();
    const todayWIB = Utilities.formatDate(timestamp, "Asia/Jakarta", "yyyy/MM/dd");
    let liveData = [];

    // Fungsi Scraper Internal untuk mengekstrak multi-sheet
    function extractLogs(dbId, isSales) {
      const db = SpreadsheetApp.openById(dbId);
      const sheets = db.getSheets();
      
      sheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if(!sheetName.startsWith("Log_")) return; 
        
        const data = sheet.getDataRange().getValues();
        if(data.length > 1) {
          let timeIndex = isSales ? 10 : 11;
          for(let i = data.length - 1; i >= 1; i--) { 
            
            // [KILLER FIX]: Cegat dari awal! Hanya proses yang statusnya "SEDANG JALAN"
            let currentStatus = isSales ? (data[i][15] || "SEDANG JALAN") : (data[i][17] || "SEDANG JALAN");
            if (currentStatus.toString().toUpperCase() !== "SEDANG JALAN") continue; // Abaikan yang sudah SELESAI

            let wKeluar = data[i][timeIndex];
            if (!wKeluar) continue;
            
            let wKeluarDate = (wKeluar instanceof Date) ? wKeluar : new Date(wKeluar.toString() + " +0700");
            if (isNaN(wKeluarDate.getTime())) continue;
            
            let wKeluarStr = Utilities.formatDate(wKeluarDate, "Asia/Jakarta", "yyyy/MM/dd");
            
            if(wKeluarStr === todayWIB) {
              liveData.push({
                nrpp: data[i][1], nama: data[i][2],
                divisi: isSales ? "SALES" : "OPS (" + data[i][7] + ")",
                customer: isSales ? data[i][8] : data[i][9],
                waktuKeluar: Utilities.formatDate(wKeluarDate, "Asia/Jakarta", "HH:mm"),
                status: currentStatus.toString().toUpperCase()
              });
            }
          }
        }
      });
    }
    
    extractLogs(DB_UPD_ID, false);
    extractLogs(DB_SALES_ID, true);

    return { status: "success", data: liveData };
  } catch(e) { return { status: "error", message: e.toString() }; }
}

// ==========================================================
// MODUL SUPER ADMIN: REKAPITULASI & EKSPOR DATA
// ==========================================================

function api_adminGetRekapSheets() {
  try {
    const db = SpreadsheetApp.openById(DB_REKAP_ID);
    const sheets = db.getSheets();
    // Tarik semua nama Sheet yang ada di dalam DB_REKAP
    const sheetNames = sheets.map(s => s.getName());
    return { status: "success", data: sheetNames };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

function api_adminGetRekapData(sheetName) {
  try {
    const db = SpreadsheetApp.openById(DB_REKAP_ID);
    const sheet = db.getSheetByName(sheetName);
    if(!sheet) throw new Error("Sheet Rekap '" + sheetName + "' tidak ditemukan!");
    
    // [KILLER FEATURE]: Menggunakan getDisplayValues() alih-alih getValues(). 
    // Ini memaksa Google Sheets mengirim data persis seperti yang Anda lihat di layar (Teks),
    // mencegah kerusakan format jam (-13.99) atau Date Object Object saat diekspor.
    const data = sheet.getDataRange().getDisplayValues(); 
    
    return { status: "success", data: data, sheetName: sheetName };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

// ==========================================================
// MODUL USER: RIWAYAT PERJALANAN DINAS & UPD (GROUPING ST)
// ==========================================================

function api_getLogPribadi(user, filterBulan, filterTahun) {
  try {
    const dbUpd = SpreadsheetApp.openById(DB_UPD_ID);
    let targetSheetName = "Log_" + user.lokasi.trim();
    let sheet = dbUpd.getSheetByName(targetSheetName);
    if(!sheet) return { status: "success", data: {} };
    
    const data = sheet.getDataRange().getValues();
    if(data.length <= 1) return { status: "success", data: {} };

    const dbMaster = SpreadsheetApp.openById(DB_MASTER_ID);
    const masterUPD = dbMaster.getSheetByName("Master_UPD").getDataRange().getValues();
    
    let keyJabatan = user.jabatan ? user.jabatan.toString().trim().toUpperCase() : "";
    let keyGolongan = user.golongan ? user.golongan.toString().trim().toUpperCase() : "";
    let keyStatusKaryawan = user.statusKaryawan ? user.statusKaryawan.toString().trim().toUpperCase() : "";
    
    let upd_ge8 = 0, upd_lt8 = 0, uangMakan = 0, mknSiangLibur = 0, lainKerja = 0, lainLibur = 0;

    for (let i = 1; i < masterUPD.length; i++) {
      let dbJab = masterUPD[i][0] ? masterUPD[i][0].toString().trim().toUpperCase() : "";
      let dbGol = masterUPD[i][1] ? masterUPD[i][1].toString().trim().toUpperCase() : "";
      let dbStatusJabatan = masterUPD[i][2] ? masterUPD[i][2].toString().trim().toUpperCase() : ""; 

      let isStatusMatch = false;
      if (dbStatusJabatan === "") {
          isStatusMatch = true; 
      } else if (dbStatusJabatan === "PROJECT") {
          if (keyStatusKaryawan.includes("PROJECT") && !keyStatusKaryawan.includes("NON")) isStatusMatch = true;
      } else if (dbStatusJabatan === "NON PROJECT") {
          if (keyStatusKaryawan.includes("NON") || !keyStatusKaryawan.includes("PROJECT")) isStatusMatch = true;
      } else if (keyStatusKaryawan.includes(dbStatusJabatan)) {
          isStatusMatch = true;
      }

      if (dbJab === keyJabatan && dbGol === keyGolongan && isStatusMatch) {
        upd_ge8 = parseFloat(masterUPD[i][3] || 0); 
        upd_lt8 = parseFloat(masterUPD[i][4] || 0);
        uangMakan = parseFloat(masterUPD[i][5] || 0);
        mknSiangLibur = parseFloat(masterUPD[i][6] || 0);
        lainKerja = parseFloat(masterUPD[i][7] || 0);
        lainLibur = parseFloat(masterUPD[i][8] || 0);
        break;
      }
    }

    const headers = data[0].map(h => h.toString().toUpperCase().trim());
    const iNoST = headers.indexOf("NO_ST");
    const iCust = headers.indexOf("CUSTOMER");
    const iLokasi = headers.lastIndexOf("LOKASI");
    const iWaktuKeluar = headers.indexOf("WAKTU_KELUAR") !== -1 ? headers.indexOf("WAKTU_KELUAR") : headers.indexOf("WAKTU KELUAR");
    const iWaktuMasuk = headers.indexOf("WAKTU_MASUK") !== -1 ? headers.indexOf("WAKTU_MASUK") : headers.indexOf("WAKTU MASUK");
    const iDurasi = headers.indexOf("DURASI_JAM") !== -1 ? headers.indexOf("DURASI_JAM") : headers.indexOf("DURASI JAM");
    const iNominal = headers.indexOf("NOMINAL_UPD") !== -1 ? headers.indexOf("NOMINAL_UPD") : headers.indexOf("NOMINAL UPD");
    const iStatus = headers.indexOf("STATUS_PERJALANAN") !== -1 ? headers.indexOf("STATUS_PERJALANAN") : headers.indexOf("STATUS PERJALANAN");
    const iKlaim = headers.indexOf("STATUS_KLAIM") !== -1 ? headers.indexOf("STATUS_KLAIM") : headers.indexOf("STATUS KLAIM");

    let groupedData = {};

    // [FIX BUG 2: DEFAULT PARAMETER AMAN] - Memastikan parameter adalah angka valid
    const now = new Date();
    let pBulan = parseInt(filterBulan, 10);
    let pTahun = parseInt(filterTahun, 10);

    let fixBulan = (isNaN(pBulan) || pBulan === 0) ? (now.getMonth() + 1) : pBulan;
    let fixTahun = (isNaN(pTahun) || pTahun === 0) ? now.getFullYear() : pTahun;

    const tBulanStr = String(fixBulan).padStart(2, '0');
    const tTahunStr = String(fixTahun);
    
    for(let i = data.length - 1; i >= 1; i--) {
       if(data[i][1].toString() === user.nrpp.toString()) {
           
           let wKeluar = (iWaktuKeluar !== -1) ? data[i][iWaktuKeluar] : "";
           let wKeluarStr = (wKeluar instanceof Date) ? Utilities.formatDate(wKeluar, "Asia/Jakarta", "dd/MM/yyyy HH:mm") : wKeluar.toString().trim();
           
           let wMasuk = (iWaktuMasuk !== -1 && data[i][iWaktuMasuk]) ? data[i][iWaktuMasuk] : "";
           let wMasukStr = (wMasuk instanceof Date) ? Utilities.formatDate(wMasuk, "Asia/Jakarta", "dd/MM/yyyy HH:mm") : (wMasuk ? wMasuk.toString().trim() : "");
           
           // [FIX BUG 1: TUTUP LUBANG DATA BOCOR]
           let dateToCheck = wKeluarStr !== "" ? wKeluarStr : wMasukStr;
           let isMatch = false;

           if (dateToCheck !== "") {
               isMatch = dateToCheck.includes(tTahunStr + "/" + tBulanStr) || 
                         dateToCheck.includes(tBulanStr + "/" + tTahunStr) || 
                         dateToCheck.includes(tTahunStr + "-" + tBulanStr) || 
                         dateToCheck.includes(tBulanStr + "-" + tTahunStr);
           } else {
               // Jika tanggal kosong (sedang jalan tapi belum ke-record), TAMPILKAN HANYA di filter bulan berjalan ini.
               if (fixBulan === (now.getMonth() + 1) && fixTahun === now.getFullYear()) {
                   isMatch = true;
               }
           }

           if (!isMatch) continue; // KUNCI MATI: Buang data yang tidak cocok!
           // ------------------------------------------------------------------

           let st = (iNoST !== -1 && data[i][iNoST]) ? data[i][iNoST].toString().trim() : "TANPA ST";
           if(!st) st = "TANPA ST";
           if(!groupedData[st]) groupedData[st] = [];
           
           let durVal = (iDurasi !== -1 && data[i][iDurasi]) ? parseFloat(data[i][iDurasi]) : 0;
           let dbNominal = (iNominal !== -1 && data[i][iNominal]) ? parseFloat(data[i][iNominal]) : 0;
           
           let logDateObj = (wKeluar instanceof Date) ? wKeluar : new Date(); 
           if (wMasuk instanceof Date) logDateObj = wMasuk; 
           else if (typeof wMasuk === "string" && wMasuk.length > 5) logDateObj = new Date(wMasuk.toString().replace(/-/g, "/") + " +0700");
           
           let isWeekend = (logDateObj.getDay() === 0 || logDateObj.getDay() === 6);

           let calc_upd = (durVal >= 8) ? upd_ge8 : upd_lt8;
           let calc_makanTotal = uangMakan;
           let calc_makanSiang = isWeekend ? mknSiangLibur : 0;
           let calc_lain = isWeekend ? lainLibur : lainKerja;
           let calc_total = calc_upd + calc_makanTotal + calc_makanSiang + calc_lain;

           if(durVal === 0) {
               calc_upd = 0; calc_makanTotal = 0; calc_makanSiang = 0; calc_lain = 0; calc_total = 0;
           }

           groupedData[st].push({
               idTransaksi: (data[i][0]) ? data[i][0].toString() : "", 
               customer: (iCust !== -1 && data[i][iCust]) ? data[i][iCust].toString() : "-",
               lokasi: (iLokasi !== -1 && data[i][iLokasi]) ? data[i][iLokasi].toString() : "-",
               waktuKeluar: wKeluarStr,
               waktuMasuk: wMasukStr,
               durasi: durVal.toFixed(2),
               nominal: dbNominal,
               status: (iStatus !== -1 && data[i][iStatus]) ? data[i][iStatus].toString() : "SEDANG JALAN",
               klaim: (iKlaim !== -1 && data[i][iKlaim]) ? data[i][iKlaim].toString() : "BELUM KLAIM",
               persetujuan: data[i][20] ? data[i][20].toString().toUpperCase() : "PENDING", 
               breakdown: {
                   upd: calc_upd, makanTotal: calc_makanTotal, makanSiang: calc_makanSiang, lain: calc_lain, total: calc_total
               }
           });
       }
    }
    return { status: "success", data: groupedData };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

// ==========================================================
// MODUL USER: PRINT LOCK & ID KLAIM GENERATOR (AUTO-TARGETING)
// ==========================================================
function api_updatePrintStatus(stNumber, user) {
  try {
    const dbUpd = SpreadsheetApp.openById(DB_UPD_ID);
    let targetSheetName = "Log_" + user.lokasi.trim();
    let sheet = dbUpd.getSheetByName(targetSheetName);
    if(!sheet) throw new Error("Sheet log tidak ditemukan.");

    const data = sheet.getDataRange().getValues();
    const idKlaim = "PRN-" + new Date().getTime() + "-" + user.nrpp;

    // [KILLER FIX]: Dynamic Header Targeting (Rudal Pencari Kolom Otomatis)
    const headers = data[0].map(h => String(h).toUpperCase().trim());
    
    let idxST = headers.indexOf("NO_ST");
    let idxNRPP = headers.indexOf("NRPP");
    
    // Cari persis di mana Kolom Status Klaim berada
    let idxStatusKlaim = headers.indexOf("STATUS_KLAIM");
    if (idxStatusKlaim === -1) idxStatusKlaim = headers.indexOf("STATUS KLAIM");
    if (idxStatusKlaim === -1) idxStatusKlaim = 18; // Fallback jika nama beda (Kolom S)

    // Cari persis di mana Kolom ID Klaim berada
    let idxIdKlaim = headers.indexOf("ID_KLAIM");
    if (idxIdKlaim === -1) idxIdKlaim = headers.indexOf("ID KLAIM");
    if (idxIdKlaim === -1) idxIdKlaim = 19; // Fallback jika nama beda (Kolom T)

    if (idxST === -1 || idxNRPP === -1) throw new Error("Kolom NO_ST atau NRPP tidak ditemukan di database!");

    // [KILLER FIX 2]: Normalisasi String Mutlak untuk menghancurkan Ilusi Tanda Kutip (')
    const targetST = String(stNumber).replace(/^'/, '').trim();
    const targetNRPP = String(user.nrpp).trim();
    let updatedCount = 0;

    for(let i = 1; i < data.length; i++) {
      let dbST = data[i][idxST] ? String(data[i][idxST]).replace(/^'/, '').trim() : "TANPA ST";
      let dbNRPP = data[i][idxNRPP] ? String(data[i][idxNRPP]).trim() : "";

      // Jika ST dan NRPP terbukti identik, KUNCI MUTLAK!
      if(dbST === targetST && dbNRPP === targetNRPP) {
        sheet.getRange(i + 1, idxStatusKlaim + 1).setValue("SUDAH PRINT"); // +1 karena getRange dihitung dari 1
        sheet.getRange(i + 1, idxIdKlaim + 1).setValue(idKlaim);
        updatedCount++;
      }
    }
    
    console.log(`Berhasil mengunci ${updatedCount} baris untuk ST ${targetST}`);
    return { status: "success", message: `Terkunci ${updatedCount} baris` };
    
  } catch(e) {
    console.error("Print Lock Error: ", e.message);
    return { status: "error", message: e.toString() };
  }
}

// ==========================================================
// MODUL OTOMASI: ROBOT SAPU BERSIH (AUTO-SWEEPER)
// Trigger: Cron Job Harian pukul 23:50 atau 23:59
// ==========================================================

function trigger_AutoSweeper() {
  try {
    const dbUpd = SpreadsheetApp.openById(DB_UPD_ID);
    const dbSales = SpreadsheetApp.openById(DB_SALES_ID);
    const dbRekap = SpreadsheetApp.openById(DB_REKAP_ID);
    
    // 1. Eksekusi Sapu Bersih untuk Pasukan Operasional (DB_UPD)
    dbUpd.getSheets().forEach(sheet => {
      if (!sheet.getName().startsWith("Log_")) return;
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return;

      for (let i = 1; i < data.length; i++) {
        if (data[i][17] === "SEDANG JALAN") { // Kolom R (Index 17)
          let wKeluar = data[i][11]; // Kolom L (Index 11)
          let dateBase = (wKeluar instanceof Date) ? wKeluar : new Date(wKeluar);
          if (isNaN(dateBase.getTime())) dateBase = new Date(); // Fallback

          // Kunci Mutlak Checkout Paksa
          let forcedTimeStr = Utilities.formatDate(dateBase, "Asia/Jakarta", "yyyy/MM/dd") + " 23:59:00";
          
          sheet.getRange(i + 1, 14).setValue(forcedTimeStr); // Waktu Masuk
          sheet.getRange(i + 1, 15).setValue("SYSTEM_AUTO_SWEEP"); // Kordinat
          sheet.getRange(i + 1, 16).setValue(0); // Durasi = 0 (Hangus)
          sheet.getRange(i + 1, 17).setValue(0); // Nominal = 0 (Ditahan)
          sheet.getRange(i + 1, 18).setValue("SELESAI"); // Status ODOC
          sheet.getRange(i + 1, 19).setValue("PENDING"); // Red Flag ke HRD

          updateRekapSweeper(dbRekap, data[i][1], dateBase, sheet.getName().replace("Log_", ""));
        }
      }
    });

    // 2. Eksekusi Sapu Bersih untuk Pasukan Sales (DB_SALES)
    dbSales.getSheets().forEach(sheet => {
      if (!sheet.getName().startsWith("Log_")) return;
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return;

      for (let i = 1; i < data.length; i++) {
        if (data[i][15] === "SEDANG JALAN") { // Kolom P (Index 15)
          let wKeluar = data[i][10]; // Kolom K (Index 10)
          let dateBase = (wKeluar instanceof Date) ? wKeluar : new Date(wKeluar);
          if (isNaN(dateBase.getTime())) dateBase = new Date(); 

          let forcedTimeStr = Utilities.formatDate(dateBase, "Asia/Jakarta", "yyyy/MM/dd") + " 23:59:00";
          
          sheet.getRange(i + 1, 13).setValue(forcedTimeStr);
          sheet.getRange(i + 1, 14).setValue("SYSTEM_AUTO_SWEEP");
          sheet.getRange(i + 1, 15).setValue(0); 
          sheet.getRange(i + 1, 16).setValue("SELESAI");

          updateRekapSweeper(dbRekap, data[i][1], dateBase, "Sales");
        }
      }
    });
  } catch(e) { console.error("Sistem Sweeper Gagal: " + e.message); }
}

// ==========================================================
// MODUL SUPER ADMIN: LIST APPROVAL & BULK APPROVE
// ==========================================================
function api_adminGetApprovalList() {
  try {
    let approvalData = [];
    const dbUpd = SpreadsheetApp.openById(DB_UPD_ID);

    dbUpd.getSheets().forEach(sheet => {
      if (!sheet.getName().startsWith("Log_")) return;
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return;

      for(let i = data.length - 1; i >= 1; i--) {
        let statusJalan = data[i][17] ? data[i][17].toString().toUpperCase() : "";
        let statusKlaim = data[i][18] ? data[i][18].toString().toUpperCase() : "BELUM KLAIM"; // Kolom S
        let statusApprove = data[i][20] ? data[i][20].toString().toUpperCase() : "PENDING";   // Kolom U
        
        // [SMART FILTER FIX]:
        // Tampilkan jika Status Jalan SELESAI DAN (Belum Di-Approve ATAU Masih Terkunci Print)
        const butuhApproval = statusApprove.includes("PENDING");
        const masihTerkunci = statusKlaim.includes("SUDAH");

        if(statusJalan === "SELESAI" && (butuhApproval || masihTerkunci)) {
          let wMasuk = data[i][13] ? data[i][13].toString() : "";
          let isAnomali = wMasuk.includes("23:59");
          
          let fmtKeluar = (data[i][11] instanceof Date) ? Utilities.formatDate(data[i][11], "Asia/Jakarta", "yyyy-MM-dd'T'HH:mm") : "";
          let fmtMasuk = (data[i][13] instanceof Date) ? Utilities.formatDate(data[i][13], "Asia/Jakarta", "yyyy-MM-dd'T'HH:mm") : "";

          approvalData.push({
            idTransaksi: data[i][0].toString(), nrpp: data[i][1], nama: data[i][2],
            jabatan: data[i][3], golongan: data[i][4], divisi: "OPS (" + data[i][7] + ")", customer: data[i][9],
            waktuKeluar: (data[i][11] instanceof Date) ? Utilities.formatDate(data[i][11], "Asia/Jakarta", "dd/MM/yyyy HH:mm") : data[i][11].toString(),
            waktuMasuk: (data[i][13] instanceof Date) ? Utilities.formatDate(data[i][13], "Asia/Jakarta", "dd/MM/yyyy HH:mm") : (wMasuk || "-"),
            rawKeluar: fmtKeluar, rawMasuk: fmtMasuk, sheetName: sheet.getName(),
            durasi: data[i][15], nominal: data[i][16], isAnomali: isAnomali, 
            statusKlaim: statusKlaim, persetujuan: statusApprove
          });
        }
      }
    });
    return { status: "success", data: approvalData };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function api_adminEditLog(payload) {
  try {
    const db = SpreadsheetApp.openById(DB_UPD_ID);
    const sheet = db.getSheetByName(payload.sheetName);
    if(!sheet) throw new Error("Sheet asal tidak ditemukan.");

    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    for(let i=1; i<data.length; i++) {
      if(data[i][0].toString() === payload.idTransaksi) { targetRow = i + 1; break; }
    }
    if(targetRow === -1) throw new Error("ID Transaksi tidak valid.");

    // Kalkulasi Waktu Baru
    const dKeluar = new Date(payload.wKeluar.replace("T", " ") + ":00 +0700");
    const dMasuk = new Date(payload.wMasuk.replace("T", " ") + ":00 +0700");
    const diffMs = dMasuk - dKeluar;
    if(diffMs < 0) throw new Error("Waktu masuk tidak boleh lebih awal dari keluar.");
    const durasiJam = diffMs / (1000 * 60 * 60);

  // Kalkulasi Ulang UPD
    let nominalUPD = 0;
    const dbMaster = SpreadsheetApp.openById(DB_MASTER_ID);
    const masterUPD = dbMaster.getSheetByName("Master_UPD").getDataRange().getValues();
    let baseUPD = 0, uangMakan = 0, mknSiangLibur = 0, lainKerja = 0, lainLibur = 0;
    
    let keyStatusKaryawan = data[targetRow - 1][5] ? data[targetRow - 1][5].toString().trim().toUpperCase() : "";

    for(let i=1; i<masterUPD.length; i++){
      let dbJabatan = masterUPD[i][0] ? masterUPD[i][0].toString().trim().toUpperCase() : "";
      let dbGolongan = masterUPD[i][1] ? masterUPD[i][1].toString().trim().toUpperCase() : "";
      let dbStatusJabatan = masterUPD[i][2] ? masterUPD[i][2].toString().trim().toUpperCase() : ""; 

      // [STRICT MATCHER LOGIC]: Anti Overlap "NON" vs "PROJECT"
      let isStatusMatch = false;
      if (dbStatusJabatan === "") {
          isStatusMatch = true; 
      } else if (dbStatusJabatan === "PROJECT") {
          if (keyStatusKaryawan.includes("PROJECT") && !keyStatusKaryawan.includes("NON")) isStatusMatch = true;
      } else if (dbStatusJabatan === "NON PROJECT") {
          if (keyStatusKaryawan.includes("NON") || !keyStatusKaryawan.includes("PROJECT")) isStatusMatch = true;
      } else if (keyStatusKaryawan.includes(dbStatusJabatan)) {
          isStatusMatch = true;
      }

      if(dbJabatan === payload.jabatan.toUpperCase() && dbGolongan === payload.golongan.toUpperCase() && isStatusMatch){
        baseUPD = (durasiJam >= 8) ? parseFloat(masterUPD[i][3] || 0) : parseFloat(masterUPD[i][4] || 0); 
        uangMakan = parseFloat(masterUPD[i][5] || 0);
        mknSiangLibur = parseFloat(masterUPD[i][6] || 0);
        lainKerja = parseFloat(masterUPD[i][7] || 0);
        lainLibur = parseFloat(masterUPD[i][8] || 0);
        break;
      }
    }
    const isWeekend = (dMasuk.getDay() === 0 || dMasuk.getDay() === 6);
    nominalUPD = isWeekend ? (baseUPD + mknSiangLibur + uangMakan + lainLibur) : (baseUPD + uangMakan + lainKerja);

    const timeKeluarSave = Utilities.formatDate(dKeluar, "Asia/Jakarta", "yyyy/MM/dd HH:mm:ss");
    const timeMasukSave = Utilities.formatDate(dMasuk, "Asia/Jakarta", "yyyy/MM/dd HH:mm:ss");

    // Timpa Database
    sheet.getRange(targetRow, 12).setValue(timeKeluarSave); // L (Keluar)
    sheet.getRange(targetRow, 14).setValue(timeMasukSave);  // N (Masuk)
    sheet.getRange(targetRow, 16).setValue(durasiJam.toFixed(2)); // P (Durasi)
    sheet.getRange(targetRow, 17).setValue(nominalUPD); // Q (Nominal)
    sheet.getRange(targetRow, 19).setValue("BELUM KLAIM"); // Reset Klaim agar gembok terbuka

    return {status: "success", message: `Data dikoreksi! Durasi baru: ${durasiJam.toFixed(2)} Jam.`};
  } catch(e) { return {status: "error", message: e.toString()}; }
}

function api_adminBulkApprove(trxIds) {
  try {
    const dbUpd = SpreadsheetApp.openById(DB_UPD_ID);
    let count = 0;

    dbUpd.getSheets().forEach(sheet => {
      if (!sheet.getName().startsWith("Log_")) return;
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        let currentId = data[i][0].toString();
        if (trxIds.includes(currentId)) {
          // Eksekusi TEPAT di Kolom U (Kolom ke-21)
          sheet.getRange(i + 1, 21).setValue("APPROVED"); 
          count++;
        }
      }
    });

    return { status: "success", message: `${count} Dokumen perjalanan berhasil disetujui!` };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function updateRekapSweeper(dbRekap, nrpp, dateBase, lokasiStr) {
  try {
    let rekapSheetName = lokasiStr === "Sales" ? "Rekap_Absensi_Sales" : "Rekap_Absensi_" + lokasiStr;
    let sheet = dbRekap.getSheetByName(rekapSheetName);
    if(!sheet) return;

    let data = sheet.getDataRange().getValues();
    let targetRow = -1; let targetCol = -1;
    let todayStr = Utilities.formatDate(dateBase, "Asia/Jakarta", "dd/MM/yyyy");

    for (let r = 1; r < data.length; r++) {
      if (data[r][0].toString() === nrpp.toString()) { targetRow = r + 1; break; }
    }
    if(data.length > 0) {
      for (let c = 2; c < data[0].length; c += 4) {
        let cellDateStr = (data[0][c] instanceof Date) ? Utilities.formatDate(data[0][c], "Asia/Jakarta", "dd/MM/yyyy") : data[0][c].toString();
        if (cellDateStr.includes(todayStr)) { targetCol = c + 1; break; }
      }
    }
    if(targetRow !== -1 && targetCol !== -1) {
      sheet.getRange(targetRow, targetCol + 1).setValue("23:59"); // Waktu Masuk Anomali
      sheet.getRange(targetRow, targetCol + 3).setValue("0"); // Durasi Anomali
    }
  } catch(e) {}
}

// ==========================================================
// MODUL SUPER ADMIN: UNLOCK PRINT STATUS (RESET CLAIM)
// ==========================================================
function api_adminUnlockPrint(stNumber, nrpp) {
  try {
    const dbUpd = SpreadsheetApp.openById(DB_UPD_ID);
    let count = 0;
    
    // Scan seluruh sheet log operasional
    dbUpd.getSheets().forEach(sheet => {
      if (!sheet.getName().startsWith("Log_")) return;
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).toUpperCase().trim());
      
      const idxST = headers.indexOf("NO_ST");
      const idxNRPP = headers.indexOf("NRPP");
      const idxStatusKlaim = headers.indexOf("STATUS_KLAIM") !== -1 ? headers.indexOf("STATUS_KLAIM") : 18;
      const idxIdKlaim = headers.indexOf("ID_KLAIM") !== -1 ? headers.indexOf("ID_KLAIM") : 19;

      const targetST = String(stNumber).replace(/^'/, '').trim();
      const targetNRPP = String(nrpp).trim();

      for(let i = 1; i < data.length; i++) {
        let dbST = data[i][idxST] ? String(data[i][idxST]).replace(/^'/, '').trim() : "";
        let dbNRPP = data[i][idxNRPP] ? String(data[i][idxNRPP]).trim() : "";

        if(dbST === targetST && dbNRPP === targetNRPP) {
          sheet.getRange(i + 1, idxStatusKlaim + 1).setValue("BELUM PRINT"); // Reset Status
          sheet.getRange(i + 1, idxIdKlaim + 1).setValue("");               // Hapus ID Klaim
          count++;
        }
      }
    });

    return { status: "success", message: `Gembok ST ${stNumber} berhasil dibuka (${count} baris).` };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

function api_adminBulkUnlockPrint(trxIds) {
  try {
    const dbUpd = SpreadsheetApp.openById(DB_UPD_ID);
    let count = 0;

    dbUpd.getSheets().forEach(sheet => {
      if (!sheet.getName().startsWith("Log_")) return;
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).toUpperCase().trim());
      
      const idxStatusKlaim = headers.indexOf("STATUS_KLAIM") !== -1 ? headers.indexOf("STATUS_KLAIM") : 18;
      const idxIdKlaim = headers.indexOf("ID_KLAIM") !== -1 ? headers.indexOf("ID_KLAIM") : 19;

      for (let i = 1; i < data.length; i++) {
        let currentId = data[i][0].toString();
        if (trxIds.includes(currentId)) {
          sheet.getRange(i + 1, idxStatusKlaim + 1).setValue("BELUM PRINT");
          sheet.getRange(i + 1, idxIdKlaim + 1).setValue("");
          count++;
        }
      }
    });

    return { status: "success", message: `${count} Gembok berhasil dibuka secara masal!` };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

// ==========================================================
// MODUL HRIS: ENGINE SCANNER LOG INDIVIDU (30 HARI)
// ==========================================================
function api_adminGetIndividualLogs(nrpp, month, year) {
  try {
    const dbUpd = SpreadsheetApp.openById(DB_UPD_ID);
    let logs = [];
    
    dbUpd.getSheets().forEach(sheet => {
      const sheetName = sheet.getName();
      if (!sheetName.startsWith("Log_")) return;
      
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return;
      
      const headers = data[0].map(h => String(h).toUpperCase().trim());
      const idxNRPP = headers.indexOf("NRPP");
      const idxTgl = headers.indexOf("WAKTU_KELUAR");
      const idxCust = headers.indexOf("CUSTOMER");
      const idxLok = headers.indexOf("LOKASI");
      const idxDur = headers.indexOf("DURASI_JAM");
      const idxNom = headers.indexOf("NOMINAL_UPD");
      const idxStat = headers.indexOf("STATUS_PERJALANAN");
      const idxApp = headers.indexOf("STATUS_APPROVE");
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idxNRPP]) === String(nrpp)) {
          let rawTgl = data[i][idxTgl];
          let dateObj = (rawTgl instanceof Date) ? rawTgl : new Date(rawTgl);
          
          // Filter Berdasarkan Bulan dan Tahun Pilihan HRD
          if (dateObj.getMonth() === parseInt(month) && dateObj.getFullYear() === parseInt(year)) {
            logs.push({
              tanggal: Utilities.formatDate(dateObj, "Asia/Jakarta", "dd/MM/yyyy HH:mm"),
              customer: data[i][idxCust] || "-",
              lokasi: data[i][idxLok] || "-",
              durasi: parseFloat(data[i][idxDur] || 0),
              nominal: parseFloat(data[i][idxNom] || 0),
              status: data[i][idxStat] || "SELESAI",
              persetujuan: data[i][idxApp] || "PENDING",
              rawDate: dateObj.getTime()
            });
          }
        }
      }
    });
    
    logs.sort((a, b) => a.rawDate - b.rawDate); // Urutkan dari tanggal awal bulan
    return { status: "success", data: logs, filter: { month, year } };
  } catch (e) { return { status: "error", message: e.toString() }; }
}

function renderPublicVerification(stNumber, nrpp) {
  let foundData = null;
  const targetST = String(stNumber).replace(/^'/, '').trim();
  const targetNRPP = String(nrpp).trim();

  try {
    // Mencari data di DB_UPD dan DB_SALES
    const dbs = [DB_UPD_ID, DB_SALES_ID];
    for (const dbId of dbs) {
      if (foundData) break;
      const db = SpreadsheetApp.openById(dbId);
      const sheets = db.getSheets();
      
      for (const sheet of sheets) {
        const values = sheet.getDataRange().getValues();
        if (values.length < 2) continue;
        const headers = values[0].map(h => String(h).toUpperCase().trim());
        const iST = headers.indexOf("NO_ST") !== -1 ? headers.indexOf("NO_ST") : 8;
        const iNRPP = headers.indexOf("NRPP") !== -1 ? headers.indexOf("NRPP") : 1;
        
        for (let i = 1; i < values.length; i++) {
          if (String(values[i][iST]).includes(targetST) && String(values[i][iNRPP]) === targetNRPP) {
            foundData = {
              nama: values[i][2],
              customer: values[i][headers.indexOf("CUSTOMER") || 9],
              status: values[i][headers.indexOf("STATUS_PERJALANAN") || 17] || "SELESAI",
              waktu: values[i][headers.indexOf("WAKTU_KELUAR") || 11]
            };
            break;
          }
        }
      }
    }

    const color = foundData ? "emerald" : "rose";
    const statusText = foundData ? "DOKUMEN VALID" : "DATA TIDAK DITEMUKAN";

    let html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <script src="https://cdn.tailwindcss.com"></script>
      </head>
      <body class="bg-slate-100 flex items-center justify-center min-h-screen p-4">
        <div class="max-w-xs w-full bg-white rounded-3xl shadow-xl border border-slate-200 overflow-hidden">
          <div class="p-6 text-center">
            <div class="w-16 h-16 bg-${color}-100 text-${color}-600 rounded-full flex items-center justify-center mx-auto mb-4">
              <svg class="w-10 h-10" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="3" d="${foundData ? 'M5 13l4 4L19 7' : 'M6 18L18 6M6 6l12 12'}"></path></svg>
            </div>
            <h1 class="text-xl font-black text-slate-800">${statusText}</h1>
            ${foundData ? `
              <div class="mt-4 text-left text-sm space-y-2 bg-slate-50 p-4 rounded-2xl">
                <p class="text-[10px] font-bold text-slate-400 uppercase">Nama</p><p class="font-bold text-slate-700">${foundData.nama}</p>
                <p class="text-[10px] font-bold text-slate-400 uppercase">No ST</p><p class="font-bold text-indigo-600">${targetST}</p>
                <p class="text-[10px] font-bold text-slate-400 uppercase">Customer</p><p class="font-bold text-slate-700">${foundData.customer}</p>
              </div>
            ` : `<p class="text-slate-500 text-xs mt-2">Nomor ST atau NRPP tidak cocok dengan database kami.</p>`}
          </div>
          <div class="bg-slate-800 p-3 text-center text-[9px] text-white font-bold tracking-widest uppercase">C.E.P.U VERIFICATION SYSTEM</div>
        </div>
      </body>
      </html>`;
    return HtmlService.createHtmlOutput(html);
  } catch (e) {
    return HtmlService.createHtmlOutput("<p>Error: " + e.message + "</p>");
  }
}

// ==========================================================
// MODUL USER: UPDATE NOMOR ST EPICOR (FIXED VERSION)
// ==========================================================
function api_userUpdateST(trxIds, newST, user) {
  try {
    const dbUpd = SpreadsheetApp.openById(DB_UPD_ID);
    let targetSheetName = "Log_" + user.lokasi.trim();
    let sheet = dbUpd.getSheetByName(targetSheetName);
    
    if(!sheet) throw new Error("Sheet " + targetSheetName + " tidak ditemukan.");
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toUpperCase().trim());
    let idxST = headers.indexOf("NO_ST");
    
    // Safety check jika kolom NO_ST tidak ditemukan
    if (idxST === -1) idxST = 8; 
    
    let count = 0;
    for(let i = 1; i < data.length; i++) {
        // Bandingkan ID Transaksi (Kolom A)
        let currentId = String(data[i][0]).trim();
        if(trxIds.includes(currentId)) {
            let targetCell = sheet.getRange(i + 1, idxST + 1);
            
            // SOLUSI: Set format sel ke Plain Text dulu, baru isi nilainya (Tanpa tanda petik)
            targetCell.setNumberFormat("@"); 
            targetCell.setValue(newST.toString().toUpperCase().trim());
            
            count++;
        }
    }
    
    SpreadsheetApp.flush(); // Paksa sinkronisasi database
    return { status: "success", message: "Berhasil update " + count + " trip ke ST: " + newST };
    
  } catch(e) {
    console.error("Error api_userUpdateST: " + e.message);
    return { status: "error", message: "Gagal: " + e.message };
  }
}

// GANTI FUNGSI INI DI BE_Services.gs
function api_getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Master Config: Sumber kebenaran versi aplikasi.
 * Tambahkan/Update fungsi ini di BE_Services.gs
 */
function api_getSystemConfig() {
  return {
    latestVersion: "v1.2.16", // <-- UBAH KE v1.2.12 SEKARANG
    scriptUrl: ScriptApp.getService().getUrl()
  };
}
