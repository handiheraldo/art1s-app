// =========================================================================
// KONFIGURASI TABEL JADWAL
// Menyimpan peta antara nama Tab di Sheets dengan format data di Web App
// =========================================================================
var SCHEDULE_CONFIGS = [
  { sheetName: "Jadwal Rabu", key: "petugas", headers: ["Tanggal", "Pianis", "Pemimpin Acara", "Renungan", "Doa Syafaat", "Diakon"] },
  { sheetName: "Jadwal SS", key: "sekolahSabat", headers: ["Tanggal", "Pianist", "Pembawa Acara", "Ayat Inti & Doa Buka", "Berita Misi", "Pelayanan Perorangan"] },
  { sheetName: "Jadwal Khotbah", key: "khotbah", headers: ["Tanggal", "Khotbah", "Pendamping 1", "Pendamping 2", "Cerita Anak-anak", "Song Leader", "Lagu Pujian"] },
  { sheetName: "Jadwal Diakon", key: "diakon", headers: ["Tanggal", "Diakon 1", "Diakon 2", "Diakones 1", "Diakones 2"] },
  { sheetName: "Jadwal Musik", key: "musik", headers: ["Tanggal", "Pianis", "Keyboardis", "Gitaris", "Bassist", "Saxophonist", "Violinist"] },
  { sheetName: "Jadwal Perjamuan", key: "perjamuan", headers: [
    "Tanggal",
    "P. Roti & Anggur 1", "P. Roti & Anggur 2", "P. Roti & Anggur 3", "P. Roti & Anggur 4", "P. Roti & Anggur 5",
    "P. Basuh Kaki 1", "P. Basuh Kaki 2", "P. Basuh Kaki 3",
    "Pelayan Basuh Kaki 1", "Pelayan Basuh Kaki 2", "Pelayan Basuh Kaki 3",
    "Pelayan Perjamuan (L1)", "Pelayan Perjamuan (L2)", "Pelayan Perjamuan (P1)", "Pelayan Perjamuan (P2)",
    "Cuci Baskom 1", "Cuci Baskom 2", "Cuci Baskom 3", "Cuci Baskom 4",
    "Cuci Alat Perjamuan"
  ]}
];

// =========================================================================
// FUNGSI INISIALISASI: Membuat format tabel otomatis jika belum ada
// =========================================================================
function checkAndInitSheets() {
  // Mengarahkan database langsung ke ID Google Sheet spesifik milikmu
  var ss = SpreadsheetApp.openById("1PKgBirWyy3HFtfmWJc1P2jkN5ZaxwMWraApductXMn0");
  
  // 1. Sheet Pengaturan
  var sPengaturan = ss.getSheetByName("Pengaturan");
  if (!sPengaturan) {
    sPengaturan = ss.insertSheet("Pengaturan");
    sPengaturan.appendRow(["Konfigurasi", "Nilai"]);
    sPengaturan.appendRow(["PASSWORD", "admin"]);
    sPengaturan.appendRow(["YOUTUBE_URL", "https://www.youtube-nocookie.com/embed?listType=playlist&list=UUz6rQ_5zP0Y0c8V7aKx2jLQ"]);
    sPengaturan.appendRow(["AUTO_DETECT_YOUTUBE", "YA"]);
    sPengaturan.appendRow(["KEUANGAN_SALDO_PEMBUKUAN", "98982221"]);
    sPengaturan.appendRow(["KEUANGAN_SALDO_BCA", "98984480"]);
    sPengaturan.appendRow(["KEUANGAN_SIGNATURES", "Ditandatangani: Ketua Jemaat — Septha Domona • Gembala Jemaat — Pdt. David Indra Utomo • Bendahara — Ari Wattimena"]);
    sPengaturan.getRange("A1:B1").setFontWeight("bold");
    sPengaturan.setColumnWidth(1, 150);
    sPengaturan.setColumnWidth(2, 400);
  } else {
    var data = sPengaturan.getDataRange().getValues();
    var keys = data.map(function(r) { return r[0]; });
    if (keys.indexOf("AUTO_DETECT_YOUTUBE") === -1) {
      sPengaturan.appendRow(["AUTO_DETECT_YOUTUBE", "YA"]);
    }
    if (keys.indexOf("KEUANGAN_SALDO_PEMBUKUAN") === -1) {
      sPengaturan.appendRow(["KEUANGAN_SALDO_PEMBUKUAN", "98982221"]);
    }
    if (keys.indexOf("KEUANGAN_SALDO_BCA") === -1) {
      sPengaturan.appendRow(["KEUANGAN_SALDO_BCA", "98984480"]);
    }
    if (keys.indexOf("KEUANGAN_SIGNATURES") === -1) {
      sPengaturan.appendRow(["KEUANGAN_SIGNATURES", "Ditandatangani: Ketua Jemaat — Septha Domona • Gembala Jemaat — Pdt. David Indra Utomo • Bendahara — Ari Wattimena"]);
    }
  }
  
  // 2. Sheet Pejabat
  var sPejabat = ss.getSheetByName("Pejabat");
  if (!sPejabat) {
    sPejabat = ss.insertSheet("Pejabat");
    sPejabat.appendRow(["ID", "Jabatan", "Nama", "WhatsApp", "Link Foto"]);
    sPejabat.getRange("A1:E1").setFontWeight("bold");
    sPejabat.setFrozenRows(1);
    
    var initialPejabat = [
      ["gembala", "Gembala Jemaat", "Pdt. [Nama Gembala]", "62800000000", "https://ui-avatars.com/api/?name=Gembala+Jemaat&background=eff6ff&color=1e3a8a&size=128"],
      ["ketua1", "Ketua Jemaat 1", "Bpk. [Nama Ketua 1]", "62800000000", "https://ui-avatars.com/api/?name=Ketua+1&background=eff6ff&color=1e3a8a&size=128"],
      ["ketua2", "Ketua Jemaat 2", "Bpk. [Nama Ketua 2]", "62800000000", "https://ui-avatars.com/api/?name=Ketua+2&background=eff6ff&color=1e3a8a&size=128"],
      ["sekretaris1", "Sekretaris 1", "Ibu [Nama Sekretaris 1]", "62800000000", "https://ui-avatars.com/api/?name=Sekretaris+1&background=f0fdf4&color=14532d&size=128"],
      ["sekretaris2", "Sekretaris 2", "Bpk. [Nama Sekretaris 2]", "62800000000", "https://ui-avatars.com/api/?name=Sekretaris+2&background=f0fdf4&color=14532d&size=128"],
      ["bendahara1", "Bendahara 1", "Ibu Ari Pahlawani", "62800000000", "https://ui-avatars.com/api/?name=Ari+Pahlawani&background=fffbeb&color=78350f&size=128"],
      ["bendahara2", "Bendahara 2", "Ibu Kasfia Naibaho", "62800000000", "https://ui-avatars.com/api/?name=Kasfia+Naibaho&background=fffbeb&color=78350f&size=128"],
      ["multimedia", "Multimedia", "Sdr. [Nama Multimedia]", "62800000000", "https://ui-avatars.com/api/?name=Multimedia&background=e0e7ff&color=3730a3&size=128"],
      ["sound_system", "Sound System", "Sdr. [Nama Sound System]", "62800000000", "https://ui-avatars.com/api/?name=Sound+System&background=e0e7ff&color=3730a3&size=128"],
      ["diakon", "Ketua Diakon & Diakones", "Bpk. [Nama Ketua Diakon]", "62800000000", "https://ui-avatars.com/api/?name=Diakon+Diakones&background=f3f4f6&color=1f2937&size=128"],
      ["pemuda", "Ketua Pemuda", "Sdr. [Nama Ketua Pemuda]", "62800000000", "https://ui-avatars.com/api/?name=Ketua+Pemuda&background=faf5ff&color=581c87&size=128"]
    ];
    sPejabat.getRange(2, 1, initialPejabat.length, 5).setValues(initialPejabat);
  }
  
  // 3. Loop untuk membuat semua Tab Jadwal jika belum ada
  for (var i = 0; i < SCHEDULE_CONFIGS.length; i++) {
    var conf = SCHEDULE_CONFIGS[i];
    var sheet = ss.getSheetByName(conf.sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(conf.sheetName);
      sheet.appendRow(conf.headers);
      sheet.getRange(1, 1, 1, conf.headers.length).setFontWeight("bold").setBackground("#eef2f6");
      sheet.setFrozenRows(1);
    }
  }

  // 4. Sheet Keuangan_KasOps
  var sKasOps = ss.getSheetByName("Keuangan_KasOps");
  if (!sKasOps) {
    sKasOps = ss.insertSheet("Keuangan_KasOps");
    sKasOps.appendRow(["Bulan", "Tahun", "Saldo Awal", "Debet", "Kredit", "Saldo Akhir"]);
    sKasOps.getRange("A1:F1").setFontWeight("bold").setBackground("#eef2f6");
    sKasOps.setFrozenRows(1);
    
    var initialKasOps = [
      ["Des", "2025", 10611132, 10855000, 11805150, 9660982],
      ["Jan", "2026", 9660982, 10615496, 7859322, 12417156],
      ["Feb", "2026", 12417156, 11464300, 8944386, 14937070],
      ["Mar", "2026", 14937070, 8255225, 11823456, 11368839],
      ["Apr", "2026", 11368839, 7986200, 8159786, 11195253],
      ["Mei", "2026", 11195253, 11778000, 9114626, 13858627]
    ];
    sKasOps.getRange(2, 1, initialKasOps.length, 6).setValues(initialKasOps);
  }

  // 5. Sheet Keuangan_KasTotal
  var sKasTotal = ss.getSheetByName("Keuangan_KasTotal");
  if (!sKasTotal) {
    sKasTotal = ss.insertSheet("Keuangan_KasTotal");
    sKasTotal.appendRow(["Nama Kas", "Saldo"]);
    sKasTotal.getRange("A1:B1").setFontWeight("bold").setBackground("#eef2f6");
    sKasTotal.setFrozenRows(1);
    
    var initialKasTotal = [
      ["Kas Operasional Gereja", 13858627],
      ["Kas Pembangunan Gereja", 3760725],
      ["Kas PA (Bukan Kas Gereja)", 11227954],
      ["Kas Kursus Gratis", 5725998],
      ["Kas SSA", 1503967],
      ["Kas Dana Sosial", 1227200],
      ["Kas BWA", 4073000],
      ["Kas Penginjilan", 2824750],
      ["Kas Administrasi Gereja", 28500000],
      ["Kas ART1S Senior", 812500],
      ["Kas Pembelian Angklung", 250000],
      ["Kas Choir ART1S Pemuda", 618000],
      ["Kas Pembelian A/C", 1000000]
    ];
    sKasTotal.getRange(2, 1, initialKasTotal.length, 2).setValues(initialKasTotal);
  }

  // 6. Sheet Keuangan_KJKT
  var sKJKT = ss.getSheetByName("Keuangan_KJKT");
  if (!sKJKT) {
    sKJKT = ss.insertSheet("Keuangan_KJKT");
    sKJKT.appendRow(["Tanggal", "Keterangan", "Perpuluhan", "Terpadu", "Jumlah"]);
    sKJKT.getRange("A1:E1").setFontWeight("bold").setBackground("#eef2f6");
    sKJKT.setFrozenRows(1);
    
    var initialKJKT = [
      ["02/05", "Persembahan SS & Khotbah", "", 831500, 831500],
      ["09/05", "Persembahan SS & Khotbah", "", 759500, 759500],
      ["16/05", "Persembahan SS & Khotbah", "", 706000, 706000],
      ["23/05", "Persembahan SS & Khotbah", "", 866000, 866000],
      ["30/05", "Persembahan SS & Khotbah", "", 486000, 486000],
      ["", "Perpuluhan & Persembahan Mei '26 (Amplop Persan & Tf.)", 19122500, 828000, 19950500]
    ];
    sKJKT.getRange(2, 1, initialKJKT.length, 5).setValues(initialKJKT);
  }
  
  return ss;
}

// =========================================================================
// MENDETEKSI LIVE STREAM ATAU VIDEO TERBARU SECARA OTOMATIS
// =========================================================================
function getLatestYoutubeVideo() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("latest_youtube_video_data");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  var fallbackData = {
    videoId: "kdBIXxl3pfM",
    title: "Live Streaming ART1STV",
    isLive: false,
    url: "https://www.youtube.com/embed/kdBIXxl3pfM"
  };

  try {
    var channelUrl = "https://www.youtube.com/@art1stv/live";
    var response = UrlFetchApp.fetch(channelUrl, {
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
      }
    });

    if (response.getResponseCode() !== 200) {
      return fallbackData;
    }

    var html = response.getContentText();
    var startText = "ytInitialData = ";
    var startIndex = html.indexOf(startText);
    if (startIndex !== -1) {
      var start = startIndex + startText.length;
      var end = html.indexOf(";</script>", start);
      if (end === -1) {
        end = html.indexOf(";\n", start);
      }
      if (end !== -1) {
        var jsonStr = html.substring(start, end);
        var data = JSON.parse(jsonStr);
        
        var videos = [];
        function findLockups(obj) {
          if (!obj || typeof obj !== 'object') return;
          if (obj.lockupViewModel) {
            videos.push(obj.lockupViewModel);
          }
          for (var key in obj) {
            if (obj.hasOwnProperty(key)) {
              findLockups(obj[key]);
            }
          }
        }
        findLockups(data);
        
        if (videos.length > 0) {
          var firstVideo = videos[0];
          var videoId = firstVideo.contentId;
          var title = firstVideo.metadata && firstVideo.metadata.lockupMetadataViewModel && firstVideo.metadata.lockupMetadataViewModel.title && firstVideo.metadata.lockupMetadataViewModel.title.content || "";
          
          var isLive = false;
          var overlays = firstVideo.contentImage && firstVideo.contentImage.thumbnailViewModel && firstVideo.contentImage.thumbnailViewModel.overlays;
          if (overlays && Array.isArray(overlays)) {
            for (var i = 0; i < overlays.length; i++) {
              var badge = overlays[i].thumbnailBottomOverlayViewModel && overlays[i].thumbnailBottomOverlayViewModel.badges && overlays[i].thumbnailBottomOverlayViewModel.badges[0] && overlays[i].thumbnailBottomOverlayViewModel.badges[0].thumbnailBadgeViewModel;
              if (badge) {
                if (badge.badgeStyle === 'THUMBNAIL_OVERLAY_BADGE_STYLE_LIVE') {
                  isLive = true;
                }
              }
            }
          }
          
          var result = {
            videoId: videoId,
            title: title,
            isLive: isLive,
            url: "https://www.youtube.com/embed/" + videoId
          };
          
          // Cache the result for 5 minutes (300 seconds)
          cache.put("latest_youtube_video_data", JSON.stringify(result), 300);
          return result;
        }
      }
    }
  } catch (err) {
    Logger.log("Error in getLatestYoutubeVideo: " + err.toString());
  }

  return fallbackData;
}

// =========================================================================
// MENGAMBIL DATA: Membaca tabel dan mengubahnya jadi objek JSON ke Web
// =========================================================================
function doGet(e) {
  var ss = checkAndInitSheets();
  
  // --- Baca Pengaturan ---
  var sPengaturan = ss.getSheetByName("Pengaturan");
  var pengData = sPengaturan.getDataRange().getValues();
  var youtubeUrl = "https://www.youtube-nocookie.com/embed?listType=playlist&list=UUz6rQ_5zP0Y0c8V7aKx2jLQ";
  var heroImageUrl = "";
  var kategoriPejabat = ["Kepemimpinan", "Keuangan", "Departemen & Pelayanan", "Lainnya"];
  var autoDetectYoutube = true;
  
  var saldoPembukuan = 98982221;
  var saldoBca = 98984480;
  var signaturesKeuangan = "Ditandatangani: Ketua Jemaat — Septha Domona • Gembala Jemaat — Pdt. David Indra Utomo • Bendahara — Ari Wattimena";
  
  for (var i = 1; i < pengData.length; i++) {
    if (pengData[i][0] === "YOUTUBE_URL") youtubeUrl = pengData[i][1].toString();
    if (pengData[i][0] === "HERO_IMAGE_URL") heroImageUrl = pengData[i][1].toString();
    if (pengData[i][0] === "AUTO_DETECT_YOUTUBE") autoDetectYoutube = pengData[i][1].toString() === "YA";
    if (pengData[i][0] === "KEUANGAN_SALDO_PEMBUKUAN") saldoPembukuan = Number(pengData[i][1]) || 0;
    if (pengData[i][0] === "KEUANGAN_SALDO_BCA") saldoBca = Number(pengData[i][1]) || 0;
    if (pengData[i][0] === "KEUANGAN_SIGNATURES") signaturesKeuangan = pengData[i][1].toString();
    if (pengData[i][0] === "KATEGORI_PEJABAT") {
      try {
        kategoriPejabat = JSON.parse(pengData[i][1].toString());
      } catch (e) {}
    }
  }

  var isLiveYoutube = false;
  var youtubeTitle = "";
  
  if (autoDetectYoutube) {
    var ytData = getLatestYoutubeVideo();
    youtubeUrl = ytData.url;
    isLiveYoutube = ytData.isLive;
    youtubeTitle = ytData.title;
  }
  
  // --- Baca Data Pejabat ---
  var sPejabat = ss.getSheetByName("Pejabat");
  var pData = sPejabat.getDataRange().getValues();
  var dataPejabat = [];
  for (var i = 1; i < pData.length; i++) {
    if (pData[i][0]) {
      dataPejabat.push({
        id: pData[i][0].toString(),
        jabatan: pData[i][1].toString(),
        nama: pData[i][2].toString(),
        wa: pData[i][3].toString().replace(/'/g, ''),
        img: pData[i][4].toString(),
        kategori: pData[i][5] ? pData[i][5].toString() : 'Umum'
      });
    }
  }
  
  // --- Baca Data Jadwal dari Berbagai Tab ---
  var jadwalDB = {};
  
  for (var i = 0; i < SCHEDULE_CONFIGS.length; i++) {
    var conf = SCHEDULE_CONFIGS[i];
    var sheet = ss.getSheetByName(conf.sheetName);
    if (!sheet) continue;
    
    var data = sheet.getDataRange().getValues();
    
    for (var r = 1; r < data.length; r++) {
      var tglObj = data[r][0];
      if (!tglObj || tglObj === "") continue;
      
      var dateStr = typeof tglObj === 'object' ? Utilities.formatDate(tglObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(tglObj);
      
      // Setup objek default untuk hari ini jika belum ada
      if (!jadwalDB[dateStr]) {
        var isRabu = new Date(dateStr + "T00:00:00").getDay() === 3;
        if (isRabu) {
          jadwalDB[dateStr] = { title: "Ibadah Permintaan Doa (Rabu)", time: "19:00 WIB - selesai", petugas: [] };
        } else {
          jadwalDB[dateStr] = { title: "Ibadah Sabat (Sabtu)", time: "09:00 - 12:00 WIB", sekolahSabatTime: "09:00 - 10:30 WIB", khotbahTime: "10:30 - 12:00 WIB", sekolahSabat: [], khotbah: [], diakon: [], musik: [], perjamuan: [] };
        }
      }
      
      // Ambil nilai setiap kolom dan hubungkan kembali dengan nama tugasnya
      var taskArray = [];
      for (var c = 1; c < conf.headers.length; c++) {
        taskArray.push({
          tugas: conf.headers[c],
          nama: data[r][c] ? data[r][c].toString() : ""
        });
      }
      
      jadwalDB[dateStr][conf.key] = taskArray;
    }
  }
  
  // --- Baca Susunan Lagu Khusus ---
  var sheetSusunan = ss.getSheetByName("Susunan_Lagu");
  if (sheetSusunan) {
    var dataSusunan = sheetSusunan.getDataRange().getValues();
    for (var r = 1; r < dataSusunan.length; r++) {
      var tglObj = dataSusunan[r][0];
      if (!tglObj || tglObj === "") continue;
      
      var dateStr = typeof tglObj === 'object' ? Utilities.formatDate(tglObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(tglObj);
      
      if (!jadwalDB[dateStr]) {
        var isRabu = new Date(dateStr + "T00:00:00").getDay() === 3;
        if (isRabu) {
          jadwalDB[dateStr] = { title: "Ibadah Permintaan Doa (Rabu)", time: "19:00 WIB - selesai", petugas: [] };
        } else {
          jadwalDB[dateStr] = { title: "Ibadah Sabat (Sabtu)", time: "09:00 - 12:00 WIB", sekolahSabatTime: "09:00 - 10:30 WIB", khotbahTime: "10:30 - 12:00 WIB", sekolahSabat: [], khotbah: [], diakon: [], musik: [], perjamuan: [] };
        }
      }
      
      jadwalDB[dateStr].susunan = {
        ssLaguBuka: dataSusunan[r][1] ? String(dataSusunan[r][1]) : "",
        ssLaguTutup: dataSusunan[r][2] ? String(dataSusunan[r][2]) : "",
        kAyatBersahutan: dataSusunan[r][3] ? String(dataSusunan[r][3]) : "",
        kLaguBuka: dataSusunan[r][4] ? String(dataSusunan[r][4]) : "",
        kLaguPujian1_show: dataSusunan[r][5] === "YA",
        kLaguPujian1_judul: dataSusunan[r][6] ? String(dataSusunan[r][6]) : "",
        kLaguPujian2_show: dataSusunan[r][7] === "YA",
        kLaguPujian2_judul: dataSusunan[r][8] ? String(dataSusunan[r][8]) : "",
        kLaguPujian3_show: dataSusunan[r][9] === "YA",
        kLaguPujian3_judul: dataSusunan[r][10] ? String(dataSusunan[r][10]) : "",
        kAyatInti: dataSusunan[r][11] ? String(dataSusunan[r][11]) : "",
        kLaguTutup: dataSusunan[r][12] ? String(dataSusunan[r][12]) : ""
      };
    }
  }

  // --- Baca Keuangan KasOps ---
  var sKasOps = ss.getSheetByName("Keuangan_KasOps");
  var kasOpsData = [];
  if (sKasOps) {
    var data = sKasOps.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== "") {
        kasOpsData.push({
          bulan: data[i][0].toString(),
          tahun: data[i][1].toString(),
          saldoAwal: Number(data[i][2]) || 0,
          debet: Number(data[i][3]) || 0,
          kredit: Number(data[i][4]) || 0,
          saldo: Number(data[i][5]) || 0
        });
      }
    }
  }
  
  // --- Baca Keuangan KasTotal ---
  var sKasTotal = ss.getSheetByName("Keuangan_KasTotal");
  var kasTotalData = [];
  if (sKasTotal) {
    var data = sKasTotal.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== "") {
        kasTotalData.push({
          nama: data[i][0].toString(),
          saldo: Number(data[i][1]) || 0
        });
      }
    }
  }
  
  // --- Baca Keuangan KJKT ---
  var sKJKT = ss.getSheetByName("Keuangan_KJKT");
  var kjktData = [];
  if (sKJKT) {
    var data = sKJKT.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] !== "") {
        kjktData.push({
          tgl: data[i][0].toString(),
          ket: data[i][1].toString(),
          perpuluhan: data[i][2] === "" ? null : Number(data[i][2]),
          terpadu: data[i][3] === "" ? null : Number(data[i][3]),
          jumlah: Number(data[i][4]) || 0
        });
      }
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    dataPejabat: dataPejabat,
    jadwalDB: jadwalDB,
    youtubeUrl: youtubeUrl,
    isLiveYoutube: isLiveYoutube,
    youtubeTitle: youtubeTitle,
    autoDetectYoutube: autoDetectYoutube,
    heroImageUrl: heroImageUrl,
    kategoriPejabat: kategoriPejabat,
    keuangan: {
      kasOps: kasOpsData,
      kasTotal: kasTotalData,
      kjktRows: kjktData,
      saldoPembukuan: saldoPembukuan,
      saldoBca: saldoBca,
      signatures: signaturesKeuangan
    }
  })).setMimeType(ContentService.MimeType.JSON);
}

// =========================================================================
// MENYIMPAN DATA: Menerima JSON dari Web dan menuliskannya di Tabel Sheets
// =========================================================================
function doPost(e) {
  var ss = checkAndInitSheets();
  var payload = JSON.parse(e.postData.contents);
  var action = payload.action;
  
  var sPengaturan = ss.getSheetByName("Pengaturan");
  var currentPassword = sPengaturan.getRange("B2").getValue().toString();
  
  // --- Aksi: Verifikasi Login ---
  if (action === "verifyPassword") {
    if (payload.password === currentPassword) {
      return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "Password salah"})).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // --- Aksi: Ganti Password ---
  if (action === "changePassword") {
    if (payload.oldPassword === currentPassword) {
      sPengaturan.getRange("B2").setValue(payload.newPassword);
      return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "Password lama salah"})).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // --- Aksi: Simpan URL/Pengaturan YouTube ---
  if (action === "saveYoutubeUrl" || action === "saveYoutubeSettings") {
    if (payload.password !== currentPassword) { return ContentService.createTextOutput(JSON.stringify({success: false, message: "Akses Ditolak"})).setMimeType(ContentService.MimeType.JSON); }
    
    var pengData = sPengaturan.getDataRange().getValues();
    
    // Simpan YOUTUBE_URL jika diberikan
    if (payload.url !== undefined) {
      var foundUrl = false;
      for (var i = 1; i < pengData.length; i++) {
        if (pengData[i][0] === "YOUTUBE_URL") {
          sPengaturan.getRange(i + 1, 2).setValue(payload.url);
          foundUrl = true;
          break;
        }
      }
      if (!foundUrl) { sPengaturan.appendRow(["YOUTUBE_URL", payload.url]); }
    }
    
    // Simpan AUTO_DETECT_YOUTUBE jika diberikan
    if (payload.autoDetect !== undefined) {
      var val = payload.autoDetect ? "YA" : "TIDAK";
      var foundAuto = false;
      for (var i = 1; i < pengData.length; i++) {
        if (pengData[i][0] === "AUTO_DETECT_YOUTUBE") {
          sPengaturan.getRange(i + 1, 2).setValue(val);
          foundAuto = true;
          break;
        }
      }
      if (!foundAuto) { sPengaturan.appendRow(["AUTO_DETECT_YOUTUBE", val]); }
    }
    
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  }

  // --- Aksi: Simpan Hero Image ---
  if (action === "saveHeroImage") {
    if (payload.password !== currentPassword) { return ContentService.createTextOutput(JSON.stringify({success: false, message: "Akses Ditolak"})).setMimeType(ContentService.MimeType.JSON); }
    
    var pengData = sPengaturan.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < pengData.length; i++) {
      if (pengData[i][0] === "HERO_IMAGE_URL") {
        sPengaturan.getRange(i + 1, 2).setValue(payload.url);
        found = true;
        break;
      }
    }
    if (!found) { sPengaturan.appendRow(["HERO_IMAGE_URL", payload.url]); }
    
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // --- Aksi: Simpan Jadwal (Memisahkan data ke tab yang tepat) ---
  if (action === "saveJadwal") {
    if (payload.password !== currentPassword) { return ContentService.createTextOutput(JSON.stringify({success: false, message: "Akses Ditolak"})).setMimeType(ContentService.MimeType.JSON); }
    
    // PERBAIKAN: Menggunakan 'payload' (karena variabelnya bernama payload di baris 150)
    // dan menyertakan 'ss' ke dalam fungsi
    if (payload.data && payload.data.susunan) {
      simpanSusunanAcaraKeTab(ss, payload.tanggal, payload.data.susunan);
    }

    var targetDateObj = new Date(payload.tanggal + "T00:00:00");
    var isRabu = targetDateObj.getDay() === 3;
    
    // Loop melalui semua konfigurasi tab
    for (var i = 0; i < SCHEDULE_CONFIGS.length; i++) {
      var conf = SCHEDULE_CONFIGS[i];
      
      // Skip tab yang tidak sesuai harinya (Rabu hanya update 'petugas', Sabat update yang lain)
      if (isRabu && conf.key !== "petugas") continue;
      if (!isRabu && conf.key === "petugas") continue;
      
      var sheet = ss.getSheetByName(conf.sheetName);
      if (!sheet) continue;
      
      // Ambil array tugas dari payload frontend, jika tidak ada (kosong) jadikan array kosong
      var tasksFromPayload = payload.data[conf.key] || [];
      
      // Siapkan baris data baru sesuai urutan header kolom
      var rowDataToSave = ["'" + payload.tanggal];
      
      // Mulai dari indeks 1 karena indeks 0 adalah Tanggal
      for (var c = 1; c < conf.headers.length; c++) {
        var taskHeader = conf.headers[c];
        var personName = "";
        
        // Cari nama petugas berdasarkan nama tugasnya di array payload
        for (var p = 0; p < tasksFromPayload.length; p++) {
          if (tasksFromPayload[p].tugas === taskHeader) {
            personName = tasksFromPayload[p].nama;
            break;
          }
        }
        rowDataToSave.push(personName);
      }
      
      // Cari apakah tanggal ini sudah ada di dalam Sheet (untuk Update)
      var sheetData = sheet.getDataRange().getValues();
      var foundRow = -1;
      for (var r = 1; r < sheetData.length; r++) {
        var dStr = typeof sheetData[r][0] === 'object' ? Utilities.formatDate(sheetData[r][0], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(sheetData[r][0]);
        if (dStr === payload.tanggal) {
          foundRow = r + 1; // Ditambah 1 karena array mulai dari 0, baris sheet mulai dari 1
          break;
        }
      }
      
      // Jika ketemu tanggalnya, timpa barisnya. Jika belum ada, append baris baru.
      if (foundRow > -1) {
        sheet.getRange(foundRow, 1, 1, rowDataToSave.length).setValues([rowDataToSave]);
      } else {
        sheet.appendRow(rowDataToSave);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // --- Aksi: Simpan Pejabat ---
  if (action === "savePejabat") {
    if (payload.password !== currentPassword) { return ContentService.createTextOutput(JSON.stringify({success: false, message: "Akses Ditolak"})).setMimeType(ContentService.MimeType.JSON); }
    
    var sPejabat = ss.getSheetByName("Pejabat");
    
    // Bersihkan isi sheet Pejabat kecuali Header (Ubah sampai kolom ke-6)
    if (sPejabat.getLastRow() > 1) {
      sPejabat.getRange(2, 1, sPejabat.getLastRow() - 1, 6).clearContent();
    }
    
    var newRows = [];
    for (var i = 0; i < payload.data.length; i++) {
      var p = payload.data[i];
      newRows.push([p.id, p.jabatan, p.nama, "'" + p.wa, p.img, p.kategori || 'Umum']);
    }
    
    if (newRows.length > 0) {
      sPejabat.getRange(2, 1, newRows.length, 6).setValues(newRows);
    }

    // Simpan Kategori Pejabat jika disertakan
    if (payload.kategoriPejabat) {
      var pengData = sPengaturan.getDataRange().getValues();
      var foundKat = false;
      for (var i = 1; i < pengData.length; i++) {
        if (pengData[i][0] === "KATEGORI_PEJABAT") {
          sPengaturan.getRange(i + 1, 2).setValue(JSON.stringify(payload.kategoriPejabat));
          foundKat = true;
          break;
        }
      }
      if (!foundKat) { sPengaturan.appendRow(["KATEGORI_PEJABAT", JSON.stringify(payload.kategoriPejabat)]); }
    }
    
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  }
  
  // --- Aksi: Simpan Keuangan ---
  if (action === "saveKeuangan") {
    if (payload.password !== currentPassword) { return ContentService.createTextOutput(JSON.stringify({success: false, message: "Akses Ditolak"})).setMimeType(ContentService.MimeType.JSON); }
    
    // 1. Simpan ke Keuangan_KasOps
    var sKasOps = ss.getSheetByName("Keuangan_KasOps");
    if (sKasOps) {
      if (sKasOps.getLastRow() > 1) {
        sKasOps.getRange(2, 1, sKasOps.getLastRow() - 1, 6).clearContent();
      }
      var newOpsRows = [];
      for (var i = 0; i < payload.data.kasOps.length; i++) {
        var d = payload.data.kasOps[i];
        newOpsRows.push([d.bulan, d.tahun, d.saldoAwal, d.debet, d.kredit, d.saldo]);
      }
      if (newOpsRows.length > 0) {
        sKasOps.getRange(2, 1, newOpsRows.length, 6).setValues(newOpsRows);
      }
    }
    
    // 2. Simpan ke Keuangan_KasTotal
    var sKasTotal = ss.getSheetByName("Keuangan_KasTotal");
    if (sKasTotal) {
      if (sKasTotal.getLastRow() > 1) {
        sKasTotal.getRange(2, 1, sKasTotal.getLastRow() - 1, 2).clearContent();
      }
      var newTotalRows = [];
      for (var i = 0; i < payload.data.kasTotal.length; i++) {
        var d = payload.data.kasTotal[i];
        newTotalRows.push([d.nama, d.saldo]);
      }
      if (newTotalRows.length > 0) {
        sKasTotal.getRange(2, 1, newTotalRows.length, 2).setValues(newTotalRows);
      }
    }
    
    // 3. Simpan ke Keuangan_KJKT
    var sKJKT = ss.getSheetByName("Keuangan_KJKT");
    if (sKJKT) {
      if (sKJKT.getLastRow() > 1) {
        sKJKT.getRange(2, 1, sKJKT.getLastRow() - 1, 5).clearContent();
      }
      var newKjktRows = [];
      for (var i = 0; i < payload.data.kjktRows.length; i++) {
        var d = payload.data.kjktRows[i];
        var perp = d.perpuluhan === null ? "" : d.perpuluhan;
        var terp = d.terpadu === null ? "" : d.terpadu;
        newKjktRows.push([d.tgl || "", d.ket, perp, terp, d.jumlah]);
      }
      if (newKjktRows.length > 0) {
        sKJKT.getRange(2, 1, newKjktRows.length, 5).setValues(newKjktRows);
      }
    }
    
    // 4. Simpan Pengaturan Keuangan (Saldo Pembukuan, Saldo BCA, Signatures)
    var pengData = sPengaturan.getDataRange().getValues();
    var keysToUpdate = {
      "KEUANGAN_SALDO_PEMBUKUAN": payload.data.saldoPembukuan,
      "KEUANGAN_SALDO_BCA": payload.data.saldoBca,
      "KEUANGAN_SIGNATURES": payload.data.signatures
    };
    
    for (var k in keysToUpdate) {
      var found = false;
      for (var i = 1; i < pengData.length; i++) {
        if (pengData[i][0] === k) {
          sPengaturan.getRange(i + 1, 2).setValue(keysToUpdate[k]);
          found = true;
          break;
        }
      }
      if (!found) {
        sPengaturan.appendRow([k, keysToUpdate[k]]);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({success: false, message: "Aksi tidak dikenali"})).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Fungsi untuk menyimpan atau memperbarui data Susunan Acara ke tab terpisah
 * PERBAIKAN: Menambahkan parameter `ss` agar fungsi ini memakai koneksi spreadsheet yang sama
 */
function simpanSusunanAcaraKeTab(ss, tanggal, susunan) {
  var sheetName = "Susunan_Lagu";
  var sheet = ss.getSheetByName(sheetName);
  
  // Jika tab "Susunan_Lagu" belum ada, buat otomatis beserta Header kolomnya
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow([
      "Tanggal", 
      "SS Lagu Buka", 
      "SS Lagu Tutup", 
      "Khotbah Ayat Bersahutan", 
      "Khotbah Lagu Buka", 
      "Pujian 1 Tampil", 
      "Pujian 1 Judul", 
      "Pujian 2 Tampil", 
      "Pujian 2 Judul", 
      "Pujian 3 Tampil", 
      "Pujian 3 Judul", 
      "Ayat Inti", 
      "Lagu Tutup"
    ]);
    // Bekukan baris pertama agar rapi saat di-scroll
    sheet.setFrozenRows(1);
  }
  
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  // Cek apakah tanggal ini sudah ada di database
  for (var i = 1; i < data.length; i++) {
    var rowDate = typeof data[i][0] === 'object' ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(data[i][0]);
    if (rowDate === tanggal) {
      rowIndex = i + 1; // +1 karena index array dari 0, sedangkan baris sheet dari 1
      break;
    }
  }
  
  // Susun data per kolom yang akan dimasukkan ke spreadsheet
  var rowData = [
    "'" + tanggal, // Gunakan tanda kutip agar diformat sebagai text/string murni di Sheets
    susunan.ssLaguBuka || "",
    susunan.ssLaguTutup || "",
    susunan.kAyatBersahutan || "",
    susunan.kLaguBuka || "",
    susunan.kLaguPujian1_show ? "YA" : "TIDAK",
    susunan.kLaguPujian1_judul || "",
    susunan.kLaguPujian2_show ? "YA" : "TIDAK",
    susunan.kLaguPujian2_judul || "",
    susunan.kLaguPujian3_show ? "YA" : "TIDAK",
    susunan.kLaguPujian3_judul || "",
    susunan.kAyatInti || "",
    susunan.kLaguTutup || ""
  ];
  
  // Jika tanggal sudah ada, timpa (update) baris tersebut.
  // Jika belum ada, tambahkan baris baru di bawah.
  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}
