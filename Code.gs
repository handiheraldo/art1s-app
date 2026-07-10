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
    sPengaturan.appendRow(["YOUTUBE_API_KEY", ""]);
    sPengaturan.appendRow(["YOUTUBE_CHANNEL_ID", "UCNwVpE7CqpcKVcaUnZhUWTQ"]);
    sPengaturan.getRange("A1:B1").setFontWeight("bold");
    sPengaturan.setColumnWidth(1, 150);
    sPengaturan.setColumnWidth(2, 400);
  } else {
    var data = sPengaturan.getDataRange().getValues();
    var keys = data.map(function(r) { return r[0]; });
    if (keys.indexOf("AUTO_DETECT_YOUTUBE") === -1) {
      sPengaturan.appendRow(["AUTO_DETECT_YOUTUBE", "YA"]);
    }
    if (keys.indexOf("YOUTUBE_API_KEY") === -1) {
      sPengaturan.appendRow(["YOUTUBE_API_KEY", ""]);
    }
    if (keys.indexOf("YOUTUBE_CHANNEL_ID") === -1) {
      sPengaturan.appendRow(["YOUTUBE_CHANNEL_ID", "UCNwVpE7CqpcKVcaUnZhUWTQ"]);
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

  // 4. Sheet Buku Tamu jika belum ada
  var sBukuTamu = ss.getSheetByName("Buku_Tamu");
  if (!sBukuTamu) {
    sBukuTamu = ss.insertSheet("Buku_Tamu");
    sBukuTamu.appendRow(["ID", "Tanggal", "Nama", "WhatsApp", "Asal Jemaat", "Kunjungan", "Sumber Info", "Pesan", "Status Follow-up"]);
    sBukuTamu.getRange("A1:I1").setFontWeight("bold").setBackground("#eef2f6");
    sBukuTamu.setFrozenRows(1);
  }

  // 5. Sheet Perlawatan jika belum ada
  var sPerlawatan = ss.getSheetByName("Perlawatan");
  if (!sPerlawatan) {
    sPerlawatan = ss.insertSheet("Perlawatan");
    sPerlawatan.appendRow(["ID", "Tanggal Pengajuan", "Nama Pemohon", "WhatsApp", "Alamat / Lokasi", "Tanggal & Waktu Rencana", "Tujuan / Alasan Perlawatan", "Keterangan", "Status"]);
    sPerlawatan.getRange("A1:I1").setFontWeight("bold").setBackground("#eef2f6");
    sPerlawatan.setFrozenRows(1);
  }

  return ss;
}


// =========================================================================
// MENDETEKSI LIVE STREAM ATAU VIDEO TERBARU SECARA OTOMATIS
// =========================================================================
function getLatestYoutubeVideo(apiKey, channelId) {
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

  if (!apiKey) {
    Logger.log("getLatestYoutubeVideo: API Key is empty");
    return fallbackData;
  }
  apiKey = apiKey.toString().trim();
  
  if (!channelId) {
    channelId = "UCNwVpE7CqpcKVcaUnZhUWTQ";
  }
  channelId = channelId.toString().trim();

  try {
    // 1. Cek apakah ada live stream aktif
    var liveUrl = "https://www.googleapis.com/youtube/v3/search"
      + "?part=snippet"
      + "&channelId=" + channelId
      + "&eventType=live"
      + "&type=video"
      + "&key=" + apiKey;

    var response = UrlFetchApp.fetch(liveUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      if (data.items && data.items.length > 0) {
        var firstVideo = data.items[0];
        var videoId = firstVideo.id.videoId;
        var title = firstVideo.snippet.title;
        
        var result = {
          videoId: videoId,
          title: title,
          isLive: true,
          url: "https://www.youtube.com/embed/" + videoId
        };
        
        // Cache selama 5 menit
        cache.put("latest_youtube_video_data", JSON.stringify(result), 300);
        return result;
      }
    } else {
      Logger.log("YouTube API Live Check failed: " + response.getContentText());
    }

    // 2. Jika tidak sedang live, ambil video rekaman terbaru
    var latestUrl = "https://www.googleapis.com/youtube/v3/search"
      + "?part=snippet"
      + "&channelId=" + channelId
      + "&order=date"
      + "&type=video"
      + "&maxResults=1"
      + "&key=" + apiKey;

    var responseLatest = UrlFetchApp.fetch(latestUrl, { muteHttpExceptions: true });
    if (responseLatest.getResponseCode() === 200) {
      var dataLatest = JSON.parse(responseLatest.getContentText());
      if (dataLatest.items && dataLatest.items.length > 0) {
        var firstVideo = dataLatest.items[0];
        var videoId = firstVideo.id.videoId;
        var title = firstVideo.snippet.title;
        
        var result = {
          videoId: videoId,
          title: title,
          isLive: false,
          url: "https://www.youtube.com/embed/" + videoId
        };
        
        // Cache selama 5 menit
        cache.put("latest_youtube_video_data", JSON.stringify(result), 300);
        return result;
      }
    } else {
      Logger.log("YouTube API Latest Check failed: " + responseLatest.getContentText());
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
  
  // --- DEBUG ENDPOINT FOR YOUTUBE ---
  if (e && e.parameter && e.parameter.debug === "youtube") {
    var sPengaturan = ss.getSheetByName("Pengaturan");
    var pengData = sPengaturan.getDataRange().getValues();
    var apiKey = "";
    var channelId = "UCNwVpE7CqpcKVcaUnZhUWTQ";
    for (var i = 1; i < pengData.length; i++) {
      if (pengData[i][0] === "YOUTUBE_API_KEY") apiKey = pengData[i][1].toString().trim();
      if (pengData[i][0] === "YOUTUBE_CHANNEL_ID") channelId = pengData[i][1].toString().trim();
    }
    
    var debugInfo = {
      apiKeyProvided: !!apiKey,
      apiKeyLength: apiKey.length,
      channelId: channelId,
      liveUrlSample: "https://www.googleapis.com/youtube/v3/search?part=snippet&channelId=" + channelId + "&eventType=live&type=video&key=...",
      liveResponse: "",
      latestResponse: ""
    };
    
    if (apiKey) {
      try {
        var liveUrl = "https://www.googleapis.com/youtube/v3/search?part=snippet&channelId=" + channelId + "&eventType=live&type=video&key=" + apiKey;
        var res = UrlFetchApp.fetch(liveUrl, { muteHttpExceptions: true });
        debugInfo.liveResponse = {
          code: res.getResponseCode(),
          body: JSON.parse(res.getContentText())
        };
      } catch (err) {
        debugInfo.liveResponse = "Error: " + err.toString();
      }
      
      try {
        var latestUrl = "https://www.googleapis.com/youtube/v3/search?part=snippet&channelId=" + channelId + "&order=date&type=video&maxResults=1&key=" + apiKey;
        var resLatest = UrlFetchApp.fetch(latestUrl, { muteHttpExceptions: true });
        debugInfo.latestResponse = {
          code: resLatest.getResponseCode(),
          body: JSON.parse(resLatest.getContentText())
        };
      } catch (err) {
        debugInfo.latestResponse = "Error: " + err.toString();
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify(debugInfo)).setMimeType(ContentService.MimeType.JSON);
  }
  
  // --- Baca Pengaturan ---
  var sPengaturan = ss.getSheetByName("Pengaturan");
  var pengData = sPengaturan.getDataRange().getValues();
  var youtubeUrl = "https://www.youtube-nocookie.com/embed?listType=playlist&list=UUz6rQ_5zP0Y0c8V7aKx2jLQ";
  var heroImageUrl = "";
  var gdriveUrl = "https://drive.google.com";
  var kategoriPejabat = ["Kepemimpinan", "Keuangan", "Departemen & Pelayanan", "Lainnya"];
  var autoDetectYoutube = true;
  var youtubeApiKey = "";
  var youtubeChannelId = "UCNwVpE7CqpcKVcaUnZhUWTQ";
  
  for (var i = 1; i < pengData.length; i++) {
    if (pengData[i][0] === "YOUTUBE_URL") youtubeUrl = pengData[i][1].toString().trim();
    if (pengData[i][0] === "HERO_IMAGE_URL") heroImageUrl = pengData[i][1].toString().trim();
    if (pengData[i][0] === "GDRIVE_URL") gdriveUrl = pengData[i][1].toString().trim();
    if (pengData[i][0] === "AUTO_DETECT_YOUTUBE") autoDetectYoutube = pengData[i][1].toString().trim() === "YA";
    if (pengData[i][0] === "YOUTUBE_API_KEY") youtubeApiKey = pengData[i][1].toString().trim();
    if (pengData[i][0] === "YOUTUBE_CHANNEL_ID") youtubeChannelId = pengData[i][1].toString().trim();
    if (pengData[i][0] === "KATEGORI_PEJABAT") {
      try {
        kategoriPejabat = JSON.parse(pengData[i][1].toString());
      } catch (e) {}
    }
  }

  var isLiveYoutube = false;
  var youtubeTitle = "";
  
  if (autoDetectYoutube) {
    var ytData = getLatestYoutubeVideo(youtubeApiKey, youtubeChannelId);
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

  return ContentService.createTextOutput(JSON.stringify({
    dataPejabat: dataPejabat,
    jadwalDB: jadwalDB,
    youtubeUrl: youtubeUrl,
    isLiveYoutube: isLiveYoutube,
    youtubeTitle: youtubeTitle,
    autoDetectYoutube: autoDetectYoutube,
    youtubeChannelId: youtubeChannelId,
    heroImageUrl: heroImageUrl,
    gdriveUrl: gdriveUrl,
    kategoriPejabat: kategoriPejabat
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
      var youtubeApiKey = "";
      var youtubeChannelId = "UCNwVpE7CqpcKVcaUnZhUWTQ";
      var pengData = sPengaturan.getDataRange().getValues();
      for (var i = 1; i < pengData.length; i++) {
        if (pengData[i][0] === "YOUTUBE_API_KEY") youtubeApiKey = pengData[i][1].toString();
        if (pengData[i][0] === "YOUTUBE_CHANNEL_ID") youtubeChannelId = pengData[i][1].toString();
      }
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        youtubeApiKey: youtubeApiKey,
        youtubeChannelId: youtubeChannelId
      })).setMimeType(ContentService.MimeType.JSON);
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

    // Simpan YOUTUBE_API_KEY jika diberikan
    if (payload.apiKey !== undefined) {
      var foundKey = false;
      for (var i = 1; i < pengData.length; i++) {
        if (pengData[i][0] === "YOUTUBE_API_KEY") {
          sPengaturan.getRange(i + 1, 2).setValue(payload.apiKey);
          foundKey = true;
          break;
        }
      }
      if (!foundKey) { sPengaturan.appendRow(["YOUTUBE_API_KEY", payload.apiKey]); }
    }

    // Simpan YOUTUBE_CHANNEL_ID jika diberikan
    if (payload.channelId !== undefined) {
      var foundChan = false;
      for (var i = 1; i < pengData.length; i++) {
        if (pengData[i][0] === "YOUTUBE_CHANNEL_ID") {
          sPengaturan.getRange(i + 1, 2).setValue(payload.channelId);
          foundChan = true;
          break;
        }
      }
      if (!foundChan) { sPengaturan.appendRow(["YOUTUBE_CHANNEL_ID", payload.channelId]); }
    }
    
    // Hapus cache agar deteksi YouTube langsung mengambil data terbaru dari API saat di-refresh
    var cache = CacheService.getScriptCache();
    cache.remove("latest_youtube_video_data");
    
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

  // --- Aksi: Simpan Google Drive URL ---
  if (action === "saveGdriveUrl") {
    if (payload.password !== currentPassword) { return ContentService.createTextOutput(JSON.stringify({success: false, message: "Akses Ditolak"})).setMimeType(ContentService.MimeType.JSON); }
    
    var pengData = sPengaturan.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < pengData.length; i++) {
      if (pengData[i][0] === "GDRIVE_URL") {
        sPengaturan.getRange(i + 1, 2).setValue(payload.url);
        found = true;
        break;
      }
    }
    if (!found) { sPengaturan.appendRow(["GDRIVE_URL", payload.url]); }
    
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
  
  // --- Aksi: Kirim Buku Tamu (Public) ---
  if (action === "submitBukuTamu") {
    // Verifikasi matematika sederhana untuk mencegah spam
    var num1 = payload.num1;
    var num2 = payload.num2;
    var captchaAnswer = payload.captchaAnswer;
    if (num1 === undefined || num2 === undefined || captchaAnswer === undefined || 
        Number(captchaAnswer) !== (Number(num1) + Number(num2))) {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "Jawaban pertanyaan keamanan salah."})).setMimeType(ContentService.MimeType.JSON);
    }

    var sBukuTamu = ss.getSheetByName("Buku_Tamu");
    if (!sBukuTamu) {
      sBukuTamu = ss.insertSheet("Buku_Tamu");
      sBukuTamu.appendRow(["ID", "Tanggal", "Nama", "WhatsApp", "Asal Jemaat", "Kunjungan", "Sumber Info", "Pesan", "Status Follow-up"]);
      sBukuTamu.getRange("A1:I1").setFontWeight("bold").setBackground("#eef2f6");
      sBukuTamu.setFrozenRows(1);
    }
    
    var id = "BT_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
    var tanggal = payload.tanggal || new Date().toISOString().split('T')[0];
    var nama = payload.nama || "";
    var wa = payload.wa || "";
    var asalJemaat = payload.asalJemaat || "";
    var kunjungan = payload.kunjungan || "Pertama kali";
    var sumberInfo = payload.sumberInfo || "";
    var pesan = payload.pesan || "";
    var statusFollowUp = "Belum di-follow up";
    
    sBukuTamu.appendRow([id, tanggal, nama, "'" + wa, asalJemaat, kunjungan, sumberInfo, pesan, statusFollowUp]);
    return ContentService.createTextOutput(JSON.stringify({success: true, id: id})).setMimeType(ContentService.MimeType.JSON);
  }

  // --- Aksi: Ambil Data Buku Tamu Publik 7 Hari Terakhir (Public) ---
  if (action === "getPublicBukuTamu") {
    var sBukuTamu = ss.getSheetByName("Buku_Tamu");
    var data = [];
    if (sBukuTamu && sBukuTamu.getLastRow() > 1) {
      var range = sBukuTamu.getDataRange();
      var values = range.getValues();
      var sevenDaysAgo = new Date();
      sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
      sevenDaysAgo.setHours(0, 0, 0, 0);
      
      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        if (row[0]) {
          var dateVal = row[1];
          var dateObj;
          if (dateVal instanceof Date) {
            dateObj = dateVal;
          } else if (dateVal) {
            dateObj = new Date(String(dateVal).split('T')[0] + "T00:00:00");
          }
          
          if (dateObj && dateObj >= sevenDaysAgo) {
            var dateString = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
            data.push({
              tanggal: dateString,
              nama: row[2] ? row[2].toString() : "",
              asalJemaat: row[4] ? row[4].toString() : "",
              kunjungan: row[5] ? row[5].toString() : ""
            });
          }
        }
      }
    }
    data.reverse(); // Newest first
    return ContentService.createTextOutput(JSON.stringify({success: true, data: data})).setMimeType(ContentService.MimeType.JSON);
  }

  // --- Aksi: Ambil Data Buku Tamu (Admin) ---
  if (action === "getBukuTamu") {
    if (payload.password !== currentPassword) { 
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "Akses Ditolak"})).setMimeType(ContentService.MimeType.JSON); 
    }
    var sBukuTamu = ss.getSheetByName("Buku_Tamu");
    var data = [];
    if (sBukuTamu && sBukuTamu.getLastRow() > 1) {
      var range = sBukuTamu.getDataRange();
      var values = range.getValues();
      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        if (row[0]) {
          var dateVal = row[1];
          var dateString = "";
          if (dateVal instanceof Date) {
            dateString = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else if (dateVal) {
            dateString = String(dateVal).split('T')[0];
          }
          
          data.push({
            id: row[0].toString(),
            tanggal: dateString,
            nama: row[2] ? row[2].toString() : "",
            wa: row[3] ? row[3].toString().replace(/'/g, '') : "",
            asalJemaat: row[4] ? row[4].toString() : "",
            kunjungan: row[5] ? row[5].toString() : "",
            sumberInfo: row[6] ? row[6].toString() : "",
            pesan: row[7] ? row[7].toString() : "",
            statusFollowUp: row[8] ? row[8].toString() : "Belum di-follow up"
          });
        }
      }
    }
    data.reverse(); // Newest first
    return ContentService.createTextOutput(JSON.stringify({success: true, data: data})).setMimeType(ContentService.MimeType.JSON);
  }

  // --- Aksi: Update Status Follow-up Buku Tamu (Admin) ---
  if (action === "updateBukuTamuStatus") {
    if (payload.password !== currentPassword) { 
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "Akses Ditolak"})).setMimeType(ContentService.MimeType.JSON); 
    }
    var sBukuTamu = ss.getSheetByName("Buku_Tamu");
    if (!sBukuTamu) {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "Sheet Buku_Tamu tidak ditemukan"})).setMimeType(ContentService.MimeType.JSON);
    }
    var targetId = payload.id;
    var newStatus = payload.status;
    var values = sBukuTamu.getDataRange().getValues();
    var foundRow = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] && values[i][0].toString() === targetId) {
        foundRow = i + 1; // 1-based index
        break;
      }
    }
    if (foundRow > -1) {
      sBukuTamu.getRange(foundRow, 9).setValue(newStatus); // Column 9 is "Status Follow-up"
      return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "Data tamu tidak ditemukan"})).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // --- Aksi: Kirim Permintaan Perlawatan (Public) ---
  if (action === "submitPerlawatan") {
    var num1 = payload.num1;
    var num2 = payload.num2;
    var captchaAnswer = payload.captchaAnswer;
    if (num1 === undefined || num2 === undefined || captchaAnswer === undefined || 
        Number(captchaAnswer) !== (Number(num1) + Number(num2))) {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "Jawaban pertanyaan keamanan salah."})).setMimeType(ContentService.MimeType.JSON);
    }

    var sPerlawatan = ss.getSheetByName("Perlawatan");
    if (!sPerlawatan) {
      sPerlawatan = ss.insertSheet("Perlawatan");
      sPerlawatan.appendRow(["ID", "Tanggal Pengajuan", "Nama Pemohon", "WhatsApp", "Alamat / Lokasi", "Tanggal & Waktu Rencana", "Tujuan / Alasan Perlawatan", "Keterangan", "Status"]);
      sPerlawatan.getRange("A1:I1").setFontWeight("bold").setBackground("#eef2f6");
      sPerlawatan.setFrozenRows(1);
    }
    
    var id = "PL_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
    var tglPengajuan = payload.tanggal || new Date().toISOString().split('T')[0];
    var nama = payload.nama || "";
    var wa = payload.wa || "";
    var lokasi = payload.lokasi || "";
    var rencanaTgl = payload.rencanaTgl || "";
    var tujuan = payload.tujuan || "";
    var keterangan = payload.keterangan || "";
    var status = "Belum dijadwalkan";
    
    sPerlawatan.appendRow([id, tglPengajuan, nama, "'" + wa, lokasi, rencanaTgl, tujuan, keterangan, status]);
    return ContentService.createTextOutput(JSON.stringify({success: true, id: id})).setMimeType(ContentService.MimeType.JSON);
  }

  // --- Aksi: Ambil Data Perlawatan (Admin) ---
  if (action === "getPerlawatan") {
    if (payload.password !== currentPassword) { 
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "Akses Ditolak"})).setMimeType(ContentService.MimeType.JSON); 
    }
    var sPerlawatan = ss.getSheetByName("Perlawatan");
    var data = [];
    if (sPerlawatan && sPerlawatan.getLastRow() > 1) {
      var range = sPerlawatan.getDataRange();
      var values = range.getValues();
      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        if (row[0]) {
          var dateVal = row[1];
          var dateString = "";
          if (dateVal instanceof Date) {
            dateString = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else if (dateVal) {
            dateString = String(dateVal).split('T')[0];
          }
          
          data.push({
            id: row[0].toString(),
            tanggalPengajuan: dateString,
            nama: row[2] ? row[2].toString() : "",
            wa: row[3] ? row[3].toString().replace(/'/g, '') : "",
            lokasi: row[4] ? row[4].toString() : "",
            rencanaTgl: row[5] ? row[5].toString() : "",
            tujuan: row[6] ? row[6].toString() : "",
            keterangan: row[7] ? row[7].toString() : "",
            status: row[8] ? row[8].toString() : "Belum dijadwalkan"
          });
        }
      }
    }
    data.reverse(); // Newest first
    return ContentService.createTextOutput(JSON.stringify({success: true, data: data})).setMimeType(ContentService.MimeType.JSON);
  }

  // --- Aksi: Update Status Perlawatan (Admin atau Public Transition) ---
  if (action === "updatePerlawatanStatus") {
    var targetId = payload.id;
    var newStatus = payload.status;
    var isPublicTransition = (newStatus === "Sudah dijadwalkan" && !payload.password);
    
    if (!isPublicTransition && payload.password !== currentPassword) { 
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "Akses Ditolak"})).setMimeType(ContentService.MimeType.JSON); 
    }
    
    var sPerlawatan = ss.getSheetByName("Perlawatan");
    if (!sPerlawatan) {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "Sheet Perlawatan tidak ditemukan"})).setMimeType(ContentService.MimeType.JSON);
    }
    
    var values = sPerlawatan.getDataRange().getValues();
    var foundRow = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] && values[i][0].toString() === targetId) {
        foundRow = i + 1; // 1-based index
        break;
      }
    }
    
    if (foundRow > -1) {
      if (isPublicTransition) {
        var currentStatus = values[foundRow - 1][8] ? values[foundRow - 1][8].toString() : "";
        if (currentStatus !== "Belum dijadwalkan") {
          return ContentService.createTextOutput(JSON.stringify({success: false, message: "Status tidak dapat diubah secara publik"})).setMimeType(ContentService.MimeType.JSON);
        }
      }
      sPerlawatan.getRange(foundRow, 9).setValue(newStatus); // Column 9 is "Status"
      return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "Data perlawatan tidak ditemukan"})).setMimeType(ContentService.MimeType.JSON);
    }
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

// =========================================================================
// OTORISASI: Jalankan fungsi ini sekali saja di editor Apps Script 
// untuk memicu izin akses UrlFetchApp (koneksi keluar ke API YouTube)
// =========================================================================
function triggerAuthorization() {
  try {
    UrlFetchApp.fetch("https://www.googleapis.com/youtube/v3/search");
  } catch (e) {
    Logger.log("Otorisasi eksternal berhasil dipicu: " + e.toString());
  }
}
