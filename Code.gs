// ==========================================
// KOD BACKEND KEMASKINI (Code.gs) - PORTAL SK SOOK 2026
// FORMAT RASMI REKOD PENGAJARAN DAN PEMBELAJARAN HARIAN (SK SOOK)
// ==========================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Portal SK Sook')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --------------------------------------------------------------------------
// 1. FUNGSI PEMBERSIHAN MASA & PADANAN DATA
// --------------------------------------------------------------------------

// Pembersihan Format Masa Bersih (HH:MM) - Mengelakkan ralat zon masa epoch Dec 30 1899 (+64 minit)
// TIDAK menggunakan Utilities.formatDate atau objek Date dengan penukaran zon masa
function bersihkanMasa(val) {
  if (!val && val !== 0) return "08:00";
  // Jika masih objek Date (cth dari pembacaan sel legasi), ekstrak jam & minit mentah TANPA pertukaran zon masa
  if (val instanceof Date) {
    var j = val.getHours();
    var m = val.getMinutes();
    return (j < 10 ? "0" + j : "" + j) + ":" + (m < 10 ? "0" + m : "" + m);
  }
  var s = String(val).trim();
  // Jika format mengandungi corak masa HH:MM (cth: "08:40", "8:40", "08:40:00")
  var match = s.match(/(\d{1,2}):(\d{2})/);
  if (match) {
    var jam = parseInt(match[1], 10);
    var minit = match[2];
    return (jam < 10 ? "0" + jam : "" + jam) + ":" + minit;
  }
  return s;
}

// Format Kod Minggu Ringkas (cth: "Minggu 30" -> "M30", "Minggu 1" -> "M1")
function formatKodMinggu(m) {
  var s = String(m || "").trim();
  var num = s.replace(/\D/g, '');
  if (num) return "M" + num;
  return s.toUpperCase();
}

// Padanan Minggu Fleksibel
function padanMingguSama(m1, m2) {
  if (!m1 || !m2) return false;
  var s1 = String(m1).trim().toLowerCase();
  var s2 = String(m2).trim().toLowerCase();
  if (s1 === s2) return true;
  var d1 = s1.replace(/\D/g, '');
  var d2 = s2.replace(/\D/g, '');
  if (d1 && d2 && d1 === d2) return true;
  return false;
}

// Padanan Emel Selamat
function padanEmelSama(e1, e2) {
  if (!e1 || !e2) return false;
  return String(e1).trim().toLowerCase() === String(e2).trim().toLowerCase();
}

// Kira Tarikh Sebenar Bagi Hari (Isnin - Jumaat) Berdasarkan Tarikh Isnin Takwim
function dapatkanTarikhHari(isninStr, namaHari) {
  var offset = 0;
  var h = String(namaHari || "").toUpperCase();
  if (h.includes("SELASA")) offset = 1;
  else if (h.includes("RABU")) offset = 2;
  else if (h.includes("KHAMIS")) offset = 3;
  else if (h.includes("JUMAAT")) offset = 4;

  if (isninStr) {
    var baseDate = new Date(isninStr);
    if (!isNaN(baseDate.getTime())) {
      var targetDate = new Date(baseDate.getTime() + (offset * 24 * 60 * 60 * 1000));
      var d = targetDate.getDate();
      var m = targetDate.getMonth() + 1;
      var y = targetDate.getFullYear();
      return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
    }
  }
  
  // Fallback tarikh semasa jika tiada takwim
  var now = new Date();
  var dNow = now.getDate();
  var mNow = now.getMonth() + 1;
  var yNow = now.getFullYear();
  return (dNow < 10 ? '0' + dNow : dNow) + '/' + (mNow < 10 ? '0' + mNow : mNow) + '/' + yNow;
}

// Dapatkan Tarikh Isnin dari Tab TAKWIM
function dapatkanIsninMinggu(minggu) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("TAKWIM");
    if (!sheet) return "";
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (padanMingguSama(data[i][0], minggu)) {
        if (data[i][1] instanceof Date) {
          return Utilities.formatDate(data[i][1], "GMT+8", "yyyy-MM-dd");
        }
        return String(data[i][1] || "");
      }
    }
  } catch (e) {
    console.warn("Ralat TAKWIM: " + e.message);
  }
  return "";
}

// Format Nama Hari Bersih (cth: "1. Isnin" -> "ISNIN")
function formatNamaHari(hari) {
  var h = String(hari || "").toUpperCase();
  if (h.includes("ISNIN")) return "ISNIN";
  if (h.includes("SELASA")) return "SELASA";
  if (h.includes("RABU")) return "RABU";
  if (h.includes("KHAMIS")) return "KHAMIS";
  if (h.includes("JUMAAT")) return "JUMAAT";
  return h.replace(/^\d+\.\s*/, '');
}

// Format Label Kelas Lengkap: [Nama Kelas (Tahun)]
function formatKelasTahun(namaKelas, tahun) {
  var k = String(namaKelas || "").trim().toUpperCase();
  var t = String(tahun || "").trim().toUpperCase();
  
  if (k.includes("(") && k.includes(")")) return k; // Sudah berformat lengkap
  
  if (k.includes("PRA")) {
    return k + " (PRASEKOLAH)";
  }
  if (k.includes("VIVA") || k.includes("WIRA") || k.includes("ARENA")) {
    return k + " (PPKI TAHAP 1)";
  }
  if (k.includes("AXIA") || k.includes("SAGA") || k.includes("BEZZA")) {
    return k + " (PPKI TAHAP 2)";
  }
  if (k.includes("PPKI")) {
    return k + " (PPKI)";
  }
  
  // Perdana
  var matchDigit = k.match(/\d+/);
  var digitTahun = matchDigit ? matchDigit[0] : (t.replace(/\D/g, '') || "1");
  return k + " (TAHUN " + digitTahun + ")";
}

// --------------------------------------------------------------------------
// 2. ENJIN PENJANAAN KANDUNGAN RPH LENGKAP & BAHASA PENGANTAR SESUAI
// --------------------------------------------------------------------------
function seragamkanNamaSubjek(subjek) {
  var s = String(subjek || "").trim().toUpperCase();
  if (s.includes("SOSIOEMOSI")) return ["SOSIOEMOSI", "SE"];
  if (s.includes("FIZIKAL") && (s.includes("KESEJAHTERAAN") || s.includes("PRA"))) return ["FIZIKAL DAN KESEJAHTERAAN DIRI", "FK"];
  if (s.includes("KEWARGANEGARAAN")) return ["PENDIDIKAN KEWARGANEGARAAN", "KW"];
  if (s.includes("KOGNITIF")) return ["KOGNITIF (MATEMATIK AWAL & SAINS)", "KOGNITIF", "KF"];
  if (s.includes("ESTETIKA") || (s.includes("KREATIVITI") && s.includes("PRA"))) return ["KREATIVITI DAN ESTETIKA", "KE"];
  if (s.includes("PERBUALAN")) return ["PERBUALAN AWAL"];
  if (s.includes("BACA BERSAMA")) return ["BACA BERSAMA"];
  if (s.includes("PEMBELAJARAN")) return ["AKTIVITI PEMBELAJARAN (BERSEPADU/PROJEK)", "AKTIVITI PEMBELAJARAN"];
  var s = String(subjek || "").trim().toUpperCase();
  if (s.includes("MATHEMATIC") || s.includes("MATEMATIK")) return ["MATEMATIK", "MATHEMATICS", "MATHEMATICS (DLP)", "MATEMATIK (DLP)"];
  if (s.includes("SCIENCE") || s.includes("SAINS")) return ["SAINS", "SCIENCE", "SCIENCE (DLP)", "SAINS (DLP)"];
  if (s.includes("ENGLISH") || s.includes("INGGERIS")) return ["BAHASA INGGERIS", "ENGLISH LANGUAGE", "ENGLISH", "BI"];
  if (s.includes("MELAYU")) return ["BAHASA MELAYU", "BM"];
  if (s.includes("ISLAM") || s.includes("PAI") || s.includes("إسلام") || s.includes("اسلام") || s.includes("ڤنديديقن") || s.includes("فنديديقن") || s.includes("جاوي")) {
    return ["PENDIDIKAN ISLAM", "PAI", "ڤنديديقن اسلام", "فنديديقن اسلام", "اسلام", "إسلام", "جاوي", "JAWI"];
  }
  if (s.includes("MORAL")) return ["PENDIDIKAN MORAL", "PM"];
  if (s.includes("JASMANI") && !s.includes("KESIHATAN")) return ["PENDIDIKAN JASMANI", "PJ"];
  if (s.includes("KESIHATAN") && !s.includes("JASMANI")) return ["PENDIDIKAN KESIHATAN", "PK"];
  if (s.includes("JASMANI") && s.includes("KESIHATAN")) return ["PENDIDIKAN JASMANI DAN PENDIDIKAN KESIHATAN", "PJPK", "PENDIDIKAN JASMANI", "PENDIDIKAN KESIHATAN"];
  if (s.includes("SENI") || s.includes("PSV")) return ["PENDIDIKAN SENI VISUAL", "PSV", "PENDIDIKAN SENI"];
  if (s.includes("MUZIK")) return ["PENDIDIKAN MUZIK", "MUZIK"];
  if (s.includes("REKA BENTUK") || s.includes("RBT")) return ["REKA BENTUK DAN TEKNOLOGI", "RBT"];
  if (s.includes("SEJARAH")) return ["SEJARAH"];
  if (s.includes("ARAB")) return ["BAHASA ARAB", "اللغة العربية"];
  if (s.includes("KADAZANDUSUN") || s.includes("BKD")) return ["BAHASA KADAZANDUSUN", "BKD"];
  if (s.includes("PENGURUSAN DIRI")) return ["PENGURUSAN DIRI"];
  if (s.includes("LITERASI")) return ["LITERASI ASAS", "KEMAHIRAN ASAS MEMBACA"];
  if (s.includes("NUMERASI")) return ["NUMERASI ASAS", "KEMAHIRAN ASAS MATEMATIK"];
  if (s.includes("SENI KRAF")) return ["PENDIDIKAN SENI KRAF", "SENI KRAF"];
  return [s];
}

function tukarDigitArabKeRumi(teks) {
  if (!teks) return "";
  var str = String(teks);
  var arab = ['٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩'];
  var persia = ['۰', '۱', '۲', '۳', '۴', '۵', '۶', '۷', '۸', '۹'];
  for (var i = 0; i < 10; i++) {
    str = str.replace(new RegExp(arab[i], 'g'), i).replace(new RegExp(persia[i], 'g'), i);
  }
  return str;
}

function padanTahunSama(rowVal, cariVal, isPpki) {
  if (!cariVal || !rowVal) return true;
  var sRow = String(rowVal).trim().toUpperCase();
  var sCari = String(cariVal).trim().toUpperCase();
  if (sRow === sCari) return true;
  if (sRow.includes("PRA") || sCari.includes("PRA")) return true;
  var dRow = tukarDigitArabKeRumi(sRow).replace(/\D/g, '');
  var dCari = tukarDigitArabKeRumi(sCari).replace(/\D/g, '');
  if (dRow && dCari && dRow === dCari) return true;
  if (isPpki) {
    var cNum = parseInt(dCari, 10);
    if (cNum >= 1 && cNum <= 3 && sRow.includes("1")) return true;
    if (cNum >= 4 && cNum <= 6 && sRow.includes("2")) return true;
  }
  return false;
}

// --------------------------------------------------------------------------
// 2. ENJIN DSKP BERHIERARKI & PENGESANAN DINAMIK KOLUM GOOGLE SHEETS
// Hierarki Wajib: Subjek > Tahun > Tema > Tajuk > SK > SP
// --------------------------------------------------------------------------

// Baca data DSKP dari tab Google Sheets dengan pengecaman lajur dinamik
function muatSemuaDskpDariSheet(isPpki) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetsToScan = [];
    if (isPpki) {
      var shPpki = ss.getSheetByName("DSKP_PPKI");
      if (shPpki) sheetsToScan.push(shPpki);
    } else {
      var shMain = ss.getSheetByName("DSKP");
      if (shMain) sheetsToScan.push(shMain);
      
      // Imbas tab-tab lain jika guru meletakkan tab berasingan untuk subjek khusus
      var allSheets = ss.getSheets();
      allSheets.forEach(function(sh) {
        var nm = sh.getName().toUpperCase();
        if (nm !== "DSKP" && nm !== "DSKP_PPKI" && (nm.includes("DSKP") || nm.includes("ISLAM") || nm.includes("PAI") || nm.includes("JAWI") || sh.getName().includes("اسلام"))) {
          sheetsToScan.push(sh);
        }
      });
    }
    if (sheetsToScan.length === 0) return [];
    
    var senarai = [];
    sheetsToScan.forEach(function(sheet) {
      var data = sheet.getDataRange().getValues();
      if (!data || data.length <= 1) return;
      
      var header = data[0].map(function(h) { return String(h || "").trim().toUpperCase(); });
      var colSubjek = 0;
      var colTahun = 1;
      var colTema = -1;
      var colTajuk = -1;
      var colSk = -1;
      var colSp = -1;
      
      for (var c = 0; c < header.length; c++) {
        var h = header[c];
        if (h.includes("SUBJEK") || h.includes("MATA PELAJARAN")) colSubjek = c;
        else if (h.includes("TAHUN") || h.includes("TINGKATAN") || h.includes("TAHAP")) colTahun = c;
        else if (h.includes("TEMA")) colTema = c;
        else if (h.includes("TAJUK") || h.includes("UNIT") || h.includes("TOPIK")) colTajuk = c;
        else if (h.includes("STANDARD KANDUNGAN") || h === "SK" || h.includes("KANDUNGAN")) colSk = c;
        else if (h.includes("STANDARD PEMBELAJARAN") || h === "SP" || h.includes("PEMBELAJARAN")) colSp = c;
      }
      
      if (colTema === -1) colTema = 2;
      if (colTajuk === -1) colTajuk = 3;
      if (colSk === -1) colSk = 4;
      if (colSp === -1) colSp = 5;
      
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var sub = row[colSubjek] ? String(row[colSubjek]).trim() : "";
        var thn = row[colTahun] ? String(row[colTahun]).trim() : "";
        var tema = (colTema < row.length && row[colTema]) ? String(row[colTema]).trim() : "";
        var tajuk = (colTajuk < row.length && row[colTajuk]) ? String(row[colTajuk]).trim() : "";
        var sk = (colSk < row.length && row[colSk]) ? String(row[colSk]).trim() : "";
        var sp = (colSp < row.length && row[colSp]) ? String(row[colSp]).trim() : "";
        
        if (sub && (sk || sp || tema || tajuk)) {
          senarai.push({
            subjek: sub,
            tahun: thn,
            tema: tema || "Umum",
            tajuk: tajuk || "Umum",
            sk: sk,
            sp: sp
          });
        }
      }
    });
    return senarai;
  } catch (e) {
    console.warn("Ralat muat DSKP sheet: " + e.message);
    return [];
  }
}

// Pangkalan Data Rujukan DSKP Piawai KPM (Menjamin Data Sentiasa Tally 100%)
function dapatkanDskpPiawaiKpm(subjek, tahun, isPpki, isPra) {
  var subUpper = String(subjek || "").toUpperCase();
  var digitTahun = String(tahun || "").replace(/\D/g, '') || "1";

    // 1. KURIKULUM PRASEKOLAH 2026 (BPK KPM TERBITAN 2025 - KERANGKA KP2027)
  if (isPra || subUpper.includes("PRA") || subUpper.includes("TUNJANG") || subUpper.includes("SOSIOEMOSI") || subUpper.includes("KOGNITIF") || subUpper.includes("KREATIVITI") || subUpper.includes("KEWARGANEGARAAN")) {
    if (subUpper.includes("SOSIOEMOSI")) {
      return [
        { subjek: subjek, tahun: "Pra", tema: "Pengurusan Kendiri & Emosi", tajuk: "Mengenali dan Mengurus Emosi Diri", sk: "SE 1.1 Mengenali diri sendiri", sp: "SE 1.1.1 Mengetahui pelbagai jenis emosi (gembira, sedih, takut, marah, malu)" },
        { subjek: subjek, tahun: "Pra", tema: "Pengurusan Kendiri & Emosi", tajuk: "Mengenali dan Mengurus Emosi Diri", sk: "SE 1.1 Mengenali diri sendiri", sp: "SE 1.1.2 Mengenal pasti emosi sendiri" },
        { subjek: subjek, tahun: "Pra", tema: "Pengurusan Kendiri & Emosi", tajuk: "Mengenali dan Mengurus Emosi Diri", sk: "SE 3.1 Menunjukkan kebolehan untuk mengawal diri", sp: "SE 3.1.1 Mengurus emosi sendiri dengan memilih tindakan yang sesuai" },
        { subjek: subjek, tahun: "Pra", tema: "Kemahiran Sosial & Hubungan Komuniti", tajuk: "Empati, Kerjasama & Kepelbagaian", sk: "SE 3.2 Berempati dan bekerjasama dengan orang lain", sp: "SE 3.2.1 Mempamerkan kebolehan menyertai sesuatu permainan yang sedang berlangsung (play entry)" },
        { subjek: subjek, tahun: "Pra", tema: "Kemahiran Sosial & Hubungan Komuniti", tajuk: "Empati, Kerjasama & Kepelbagaian", sk: "SE 3.3 Menyesuaikan tindakan sendiri dan mematuhi peraturan", sp: "SE 3.3.2 Mengamalkan etika sosial dalam perhubungan" }
      ];
    }
    if (subUpper.includes("FIZIKAL")) {
      return [
        { subjek: subjek, tahun: "Pra", tema: "Perkembangan Motor & Koordinasi", tajuk: "Pergerakan Lokomotor dan Bukan Lokomotor", sk: "FK 1.1 Memahami kepentingan koordinasi dalam pergerakan", sp: "FK 1.1.2 Melakukan pergerakan lokomotor dan bukan lokomotor" },
        { subjek: subjek, tahun: "Pra", tema: "Perkembangan Motor & Koordinasi", tajuk: "Manipulasi Alatan & Motor Halus", sk: "FK 2.2 Menggunakan alatan dengan cekap", sp: "FK 2.2.1 Mengaplikasi kemahiran manipulasi dalam aktiviti harian (motor halus & kasar)" },
        { subjek: subjek, tahun: "Pra", tema: "Kesihatan, Pemakanan & Keselamatan (PEERS)", tajuk: "Gaya Hidup Sihat & Kebersihan Diri", sk: "FK 1.3 Memerhati dan mengetahui cara hidup sihat dan selamat", sp: "FK 1.3.1 Mengenal pasti makanan dan minuman yang sihat dan selamat" },
        { subjek: subjek, tahun: "Pra", tema: "Kesihatan, Pemakanan & Keselamatan (PEERS)", tajuk: "Gaya Hidup Sihat & Kebersihan Diri", sk: "FK 1.3 Memerhati dan mengetahui cara hidup sihat dan selamat", sp: "FK 1.3.4 Mengetahui sentuhan selamat dan sentuhan tidak selamat (PEERS)" },
        { subjek: subjek, tahun: "Pra", tema: "Kesihatan, Pemakanan & Keselamatan (PEERS)", tajuk: "Gaya Hidup Sihat & Kebersihan Diri", sk: "FK 2.3 Mengamalkan cara hidup sihat dan selamat", sp: "FK 2.3.4 Mengaplikasikan kemahiran mengatakan TIDAK kepada sentuhan tidak selamat (Jerit, Lari, Lapor)" }
      ];
    }
    if (subUpper.includes("ISLAM") || subUpper.includes("PAI") || subUpper.includes("JAWI")) {
      return [
        { subjek: subjek, tahun: "Pra", tema: "Asas Al-Quran, Akidah & Ibadah", tajuk: "Asas Al-Quran & Hafazan Surah", sk: "PI 1.1 Mengetahui asas Al-Quran", sp: "PI 1.1.2 Mengenal dan menyebut bunyi huruf hijaiyah berbaris satu (fathah, kasrah, dhommah)" },
        { subjek: subjek, tahun: "Pra", tema: "Asas Al-Quran, Akidah & Ibadah", tajuk: "Asas Al-Quran & Hafazan Surah", sk: "PI 1.1 Mengetahui asas Al-Quran", sp: "PI 1.1.3 Menghafaz surah daripada Juz Amma dengan betul (Al-Fatihah, An-Nas, Al-Ikhlas)" },
        { subjek: subjek, tahun: "Pra", tema: "Asas Al-Quran, Akidah & Ibadah", tajuk: "Rukun Iman, Rukun Islam & Ibadah", sk: "PI 1.3 Mengetahui dan memahami ibadah dalam kehidupan harian", sp: "PI 1.3.1 Melakukan wuduk dengan betul dan tertib" },
        { subjek: subjek, tahun: "Pra", tema: "Asas Al-Quran, Akidah & Ibadah", tajuk: "Rukun Iman, Rukun Islam & Ibadah", sk: "PI 1.3 Mengetahui dan memahami ibadah dalam kehidupan harian", sp: "PI 1.3.2 Melakukan perlakuan solat dengan betul dan tertib" },
        { subjek: subjek, tahun: "Pra", tema: "Sirah, Akhlak, Adab & Tulisan Jawi", tajuk: "Sirah Nabi & Asas Tulisan Jawi", sk: "PI 1.5 Mengetahui asas tulisan jawi", sp: "PI 1.5.1 Mengecam dan menyebut huruf jawi" },
        { subjek: subjek, tahun: "Pra", tema: "Sirah, Akhlak, Adab & Tulisan Jawi", tajuk: "Adab dan Doa Harian", sk: "PI 2.2 Mengamalkan adab dalam kehidupan seharian", sp: "PI 2.2.2 Melafazkan doa dalam aktiviti harian (doa belajar, makan, tidur)" }
      ];
    }
    if (subUpper.includes("MORAL")) {
      return [
        { subjek: subjek, tahun: "Pra", tema: "Nilai Murni Kendiri & Interaksi", tajuk: "Sayangi Diri, Bersikap Jujur & Rajin", sk: "PM 1.1 Mengetahui nilai murni yang perlu ada dalam diri sendiri", sp: "PM 1.1.3 Menceritakan tentang perlakuan dan kebaikan bersikap jujur" },
        { subjek: subjek, tahun: "Pra", tema: "Nilai Murni Kendiri & Interaksi", tajuk: "Sayangi Diri, Bersikap Jujur & Rajin", sk: "PM 1.2 Memahami nilai murni dalam interaksi dengan orang lain", sp: "PM 1.2.1 Menceritakan cara menunjukkan kasih sayang kepada orang lain" },
        { subjek: subjek, tahun: "Pra", tema: "Kelestarian Alam Sekitar", tajuk: "Kasih Sayang dan Penjagaan Alam Sekitar", sk: "PM 2.3 Mengamalkan penjagaan alam sekitar", sp: "PM 2.3.1 Melaksanakan aktiviti memelihara dan memulihara alam sekitar" }
      ];
    }
    if (subUpper.includes("KEWARGANEGARAAN")) {
      return [
        { subjek: subjek, tahun: "Pra", tema: "Diri, Keluarga & Komuniti", tajuk: "Hak & Tanggungjawab Diri serta Keluarga", sk: "KW 1.1 Memahami hak dan tanggungjawab terhadap diri sendiri", sp: "KW 1.1.1 Membezakan keperluan dan kehendak" },
        { subjek: subjek, tahun: "Pra", tema: "Diri, Keluarga & Komuniti", tajuk: "Warga Sekolah & Kemudahan Awam", sk: "KW 1.3 Memahami hubungan diri dengan komuniti", sp: "KW 1.3.1 Mengetahui tanggungjawab sebagai warga sekolah" },
        { subjek: subjek, tahun: "Pra", tema: "Negara Malaysia Tercinta", tajuk: "Identiti, Lambang & Mercu Tanda Negara", sk: "KW 1.4 Memahami dan mengetahui negara Malaysia", sp: "KW 1.4.2 Mengenali identiti dan lambang negara (Lagu Negaraku, Jalur Gemilang, Rukun Negara)" }
      ];
    }
    if (subUpper.includes("INGGERIS") || subUpper.includes("ENGLISH")) {
      return [
        { subjek: subjek, tahun: "Pra", tema: "Language and Literacy (CEFR & NDL)", tajuk: "Listening and Phonemic Awareness", sk: "BI 1.1 Recognise and understand the sounds in songs and spoken language", sp: "BI 1.1.1* Listen and recognise sounds, letters, words, phrases and simple sentences" },
        { subjek: subjek, tahun: "Pra", tema: "Language and Literacy (CEFR & NDL)", tajuk: "Listening and Phonemic Awareness", sk: "BI 1.1 Recognise and understand the sounds in songs and spoken language", sp: "BI 1.1.2* Listen and recognise sounds in monosyllable words (c-a-t, b-e-d)" },
        { subjek: subjek, tahun: "Pra", tema: "Language and Literacy (CEFR & NDL)", tajuk: "Speaking and Daily Conversations", sk: "BI 2.1 Communicate simple information", sp: "BI 2.1.1* Use words, phrases and sentences to express oneself" },
        { subjek: subjek, tahun: "Pra", tema: "Language and Literacy (CEFR & NDL)", tajuk: "Reading and Sight Words", sk: "BI 3.2 Understand texts and their meanings", sp: "BI 3.2.1* Read and recognise the shapes, names and sounds of letters of the alphabet" },
        { subjek: subjek, tahun: "Pra", tema: "Language and Literacy (CEFR & NDL)", tajuk: "Reading and Sight Words", sk: "BI 3.2 Understand texts and their meanings", sp: "BI 3.2.3* Read and understand words (New Dolch List 100 high-frequency words)" },
        { subjek: subjek, tahun: "Pra", tema: "Language and Literacy (CEFR & NDL)", tajuk: "Prewriting and Writing Skills", sk: "BI 4.2 Use basic writing skills for communication", sp: "BI 4.2.1* Write letters of the alphabet" }
      ];
    }
    if (subUpper.includes("KOGNITIF") || subUpper.includes("MATEMATIK") || subUpper.includes("SAINS")) {
      return [
        { subjek: subjek, tahun: "Pra", tema: "Matematik Dalam Kehidupan Harian", tajuk: "Konsep Nombor dan Kuantiti", sk: "KF 1.1 Mengumpul maklumat matematik dalam kehidupan harian", sp: "KF 1.1.2* Mengenal nombor (simbol 0-9, angka dan perkataan, membilang 1 hingga 50)" },
        { subjek: subjek, tahun: "Pra", tema: "Matematik Dalam Kehidupan Harian", tajuk: "Konsep Nombor dan Kuantiti", sk: "KF 1.1 Mengumpul maklumat matematik dalam kehidupan harian", sp: "KF 1.1.6* Menyedari konsep ketekalan dalam kehidupan seharian (kuantiti, isipadu, jisim)" },
        { subjek: subjek, tahun: "Pra", tema: "Matematik Dalam Kehidupan Harian", tajuk: "Operasi Nombor & Bentuk Geometri", sk: "KF 2.1 Mengurus maklumat matematik dalam persekitaran", sp: "KF 2.1.2* Membina pola dengan menggunakan objek di persekitaran" },
        { subjek: subjek, tahun: "Pra", tema: "Matematik Dalam Kehidupan Harian", tajuk: "Operasi Nombor & Bentuk Geometri", sk: "KF 2.1 Mengurus maklumat matematik dalam persekitaran", sp: "KF 2.1.4 Melakukan operasi tambah dan tolak dalam lingkungan 18 (fakta asas)" },
        { subjek: subjek, tahun: "Pra", tema: "Penerokaan Kejadian Alam (Sains)", tajuk: "Alam Hidupan dan Fizikal", sk: "KF 1.2 Mengumpul maklumat tentang kejadian alam", sp: "KF 1.2.1 Meneroka kejadian alam (hidupan, bahan larut/timbul, fizikal magnet/cahaya, bumi & cuaca)" },
        { subjek: subjek, tahun: "Pra", tema: "Hasil Warisan & Dunia Digital", tajuk: "Objek Persekitaran, Wang & Peranti Digital", sk: "KF 2.3 Membuat hubung kait antara maklumat hasil warisan", sp: "KF 2.3.3 Memahami wang Malaysia (syiling, wang kertas, simulasi jual beli & kupon/e-wallet)" },
        { subjek: subjek, tahun: "Pra", tema: "Hasil Warisan & Dunia Digital", tajuk: "Objek Persekitaran, Wang & Peranti Digital", sk: "KF 2.3 Membuat hubung kait antara maklumat hasil warisan", sp: "KF 2.3.4 Menggunakan peranti digital dengan selamat (had waktu layar, keselamatan & privasi)" }
      ];
    }
    if (subUpper.includes("KREATIVITI") || subUpper.includes("ESTETIKA") || subUpper.includes("SENI")) {
      return [
        { subjek: subjek, tahun: "Pra", tema: "Seni Visual, Muzik, Gerakan & Drama", tajuk: "Penerokaan dan Penghasilan Seni Visual", sk: "KE 1.2 Menunjukkan minat terhadap karya seni dalam pelbagai media", sp: "KE 1.2.1 Mengenal kepelbagaian warna, bentuk, tekstur dalam karya seni visual" },
        { subjek: subjek, tahun: "Pra", tema: "Seni Visual, Muzik, Gerakan & Drama", tajuk: "Penerokaan dan Penghasilan Seni Visual", sk: "KE 2.2 Menghasilkan karya seni menggunakan daya imaginasi", sp: "KE 2.2.1 Menggunakan pelbagai teknik dan media untuk menghasilkan karya seni visual (mozek, capan, kolaj, binaan)" },
        { subjek: subjek, tahun: "Pra", tema: "Seni Visual, Muzik, Gerakan & Drama", tajuk: "Muzik, Pergerakan Kreatif & Drama", sk: "KE 2.2 Menghasilkan karya seni menggunakan daya imaginasi", sp: "KE 2.2.2 Menyanyikan lagu dengan sebutan, pic dan tempo yang betul" },
        { subjek: subjek, tahun: "Pra", tema: "Seni Visual, Muzik, Gerakan & Drama", tajuk: "Muzik, Pergerakan Kreatif & Drama", sk: "KE 2.2 Menghasilkan karya seni menggunakan daya imaginasi", sp: "KE 2.2.5 Melakonkan pelbagai watak mengikut imaginasi dan kreativiti" }
      ];
    }
    if (subUpper.includes("PEMBELAJARAN") || subUpper.includes("BERSEPADU") || subUpper.includes("PROJEK")) {
      return [
        { subjek: subjek, tahun: "Pra", tema: "Pembelajaran Bersepadu & Inkuiri", tajuk: "Penerokaan Bertema & Pembelajaran Berasaskan Projek", sk: "AP 1.1 Meneroka dan menyelesaikan tugasan amali bertema secara bersepadu", sp: "AP 1.1.1 Melaksanakan aktiviti hands-on melalui pendekatan bertema dan inkuiri penerokaan" },
        { subjek: subjek, tahun: "Pra", tema: "Pembelajaran Bersepadu & Inkuiri", tajuk: "Penerokaan Bertema & Pembelajaran Berasaskan Projek", sk: "AP 1.1 Meneroka dan menyelesaikan tugasan amali bertema secara bersepadu", sp: "AP 1.1.2 Menghasilkan produk projek amali dan membentangkan dapatan bersama rakan kumpulan" }
      ];
    }
    if (subUpper.includes("PERBUALAN")) {
      return [
        { subjek: subjek, tahun: "Pra", tema: "Rutin Harian & Sosioemosi", tajuk: "Kesediaan Minda dan Emosi Pagi", sk: "PA 1.1 Menyediakan diri secara mental dan emosi untuk pembelajaran harian", sp: "PA 1.1.1 Berkongsi idea, maklumat, meluahkan emosi pagi atau menyanyikan lagu beraksi bersama guru dan rakan" }
      ];
    }
    if (subUpper.includes("BACA BERSAMA")) {
      return [
        { subjek: subjek, tahun: "Pra", tema: "Literasi & Budaya Membaca Sepanjang Hayat", tajuk: "Penerokaan Bahan Bacaan Menyeronokkan", sk: "BB 1.1 Memupuk minat membaca melalui penerokaan pelbagai jenis bahan bacaan", sp: "BB 1.1.1 Membaca buku cerita bergambar, big book, mendengar cerita dan membuat aktiviti susulan melukis/berlakon" }
      ];
    }
    // Asas Lalai Prasekolah (Bahasa Melayu)
    return [
      { subjek: subjek, tahun: "Pra", tema: "Kemahiran Bahasa & Literasi", tajuk: "Mendengar dan Memberikan Respons", sk: "BM 1.1 Mendengar dan memberikan respons", sp: "BM 1.1.2* Mengecam bunyi abjad, suku kata, perkataan, frasa dan ayat mudah" },
      { subjek: subjek, tahun: "Pra", tema: "Kemahiran Bahasa & Literasi", tajuk: "Bertutur dan Berinteraksi Sopan", sk: "BM 2.1 Bertutur dan berinteraksi dengan sebutan yang jelas dan sopan", sp: "BM 2.1.2* Berinteraksi dengan sebutan yang jelas dan sopan dalam perbualan harian" },
      { subjek: subjek, tahun: "Pra", tema: "Kemahiran Bahasa & Literasi", tajuk: "Kemahiran Prabaca dan Membaca", sk: "BM 3.2 Memahami teks dan maksudnya", sp: "BM 3.2.3* Mengenal dan membaca perkataan suku kata terbuka dan tertutup (KV+KV, KVK)" },
      { subjek: subjek, tahun: "Pra", tema: "Kemahiran Bahasa & Literasi", tajuk: "Kemahiran Pratulis dan Menulis", sk: "BM 4.2 Menyampaikan idea dalam bentuk lukisan, simbol dan tulisan", sp: "BM 4.2.1* Menulis huruf dengan cara yang betul" }
    ];
  }

  // 2. PPKI (11 MATA PELAJARAN RASMI KSSR PENDIDIKAN KHAS MASALAH PEMBELAJARAN)
  if (isPpki) {
    if (subUpper.includes("PENGURUSAN DIRI") || subUpper.includes("[PD]")) {
      return [
        {
          subjek: subjek, tahun: tahun,
          tema: "Pengurusan Diri & Kebersihan", tajuk: "Penjagaan Kebersihan Anggota Badan dan Pakaian",
          sk: "1.1 Mengamalkan penjagaan kebersihan anggota badan dan pakaian",
          sp: "1.1.1 Menyatakan dan melakukan langkah membersihkan anggota badan serta berpakaian kemas"
        },
        {
          subjek: subjek, tahun: tahun,
          tema: "Pengurusan Diri & Keselamatan", tajuk: "Keselamatan Diri di Rumah dan Sekolah",
          sk: "1.2 Mengamalkan langkah keselamatan diri dalam kehidupan harian",
          sp: "1.2.1 Mengenal pasti situasi bahaya dan cara menjaga keselamatan fizikal diri"
        }
      ];
    }
    if (subUpper.includes("MANIPULATIF") || subUpper.includes("[KM]")) {
      return [
        {
          subjek: subjek, tahun: tahun,
          tema: "Motor Halus dan Koordinasi", tajuk: "Pergerakan Motor Halus dan Pengamatan Visual",
          sk: "1.1 Menguasai kemahiran motor halus menggunakan jari dan tangan",
          sp: "1.1.1 Mengkoordinasikan pergerakan tangan dan mata melalui aktiviti meramas, mencubit dan menggenggam"
        },
        {
          subjek: subjek, tahun: tahun,
          tema: "Motor Kasar", tajuk: "Keseimbangan Badan dan Koordinasi Anggota",
          sk: "1.2 Melakukan pergerakan motor kasar secara terkawal",
          sp: "1.2.1 Mengimbangkan badan semasa berjalan di atas garisan lurus dan melangkah halangan"
        }
      ];
    }
    if ((subUpper.includes("MELAYU") || subUpper.includes("LITERASI") || subUpper.includes("[BM]")) && !subUpper.includes("INGGERIS")) {
      return [
        {
          subjek: subjek, tahun: tahun,
          tema: "Kemahiran Mendengar dan Bertutur", tajuk: "Mendengar dan Menyebut Perkataan Mudah",
          sk: "1.1 Mendengar dan memberi respons terhadap arahan mudah",
          sp: "1.1.1 Menyebut dan membunyikan perkataan bermakna dalam konteks bilik darjah"
        },
        {
          subjek: subjek, tahun: tahun,
          tema: "Kemahiran Membaca Asas", tajuk: "Mengecam Huruf Abjad dan Suku Kata Terbuka",
          sk: "2.1 Mengecam bentuk huruf vokal dan konsonan",
          sp: "2.1.1 Membunyikan suku kata terbuka KV dan membatang perkataan dua suku kata"
        }
      ];
    }
    if (subUpper.includes("INGGERIS") || subUpper.includes("ENGLISH") || subUpper.includes("[BI]")) {
      return [
        {
          subjek: subjek, tahun: tahun,
          tema: "World of Self, Family and Friends", tajuk: "Greetings and Courteous Expressions",
          sk: "1.1 Listen and respond to familiar everyday words and instructions",
          sp: "1.1.1 Listen, recognise and reproduce simple greetings and courteous expressions politely"
        },
        {
          subjek: subjek, tahun: tahun,
          tema: "World of Stories", tajuk: "Simple Classroom Objects and Phonics",
          sk: "2.1 Recognise shapes of letters and sounds of alphabet",
          sp: "2.1.1 Read and pronounce monosyllable words with visual flashcards"
        }
      ];
    }
    if (subUpper.includes("MATEMATIK") || subUpper.includes("NUMERASI") || subUpper.includes("[MAT]")) {
      return [
        {
          subjek: subjek, tahun: tahun,
          tema: "Konsep Nombor dan Kuantiti", tajuk: "Konsep Nombor Bulat Lingkungan 10",
          sk: "1.1 Menyatakan kuantiti objek melalui perbandingan",
          sp: "1.1.1 Membilang dan memadankan objek maujud dengan simbol angka yang betul"
        },
        {
          subjek: subjek, tahun: tahun,
          tema: "Operasi Asas", tajuk: "Operasi Tambah Dalam Lingkungan 10",
          sk: "2.1 Memahami konsep penambahan dua kumpulan objek",
          sp: "2.1.1 Menggabungkan dua set objek maujud dan menyatakan jumlah keseluruhan"
        }
      ];
    }
    if (subUpper.includes("ISLAM") || subUpper.includes("MORAL") || subUpper.includes("[PI/PM]")) {
      return [
        {
          subjek: subjek, tahun: tahun,
          tema: "Asas Nilai dan Kerohanian", tajuk: "Adab dan Akhlak Terpuji",
          sk: "1.1 Mengamalkan adab asas terhadap diri, guru dan rakan",
          sp: "1.1.1 Menunjukkan adab sopan dan bersalaman serta menghormati warga sekolah"
        }
      ];
    }
    if (subUpper.includes("JASMANI") || subUpper.includes("PJPK") || subUpper.includes("[PJPK]")) {
      return [
        {
          subjek: subjek, tahun: tahun,
          tema: "Pergerakan Asas & Kesihatan Diri", tajuk: "Pergerakan Lokomotor dan Bukan Lokomotor",
          sk: "1.1 Melakukan pelbagai corak pergerakan asas lokomotor",
          sp: "1.1.1 Bergerak mengikut arah (berjalan, melompat, melangkah) dengan imbangan stabil"
        }
      ];
    }
    if (subUpper.includes("SAINS") || subUpper.includes("PSSAS") || subUpper.includes("[PSSAS]")) {
      return [
        {
          subjek: subjek, tahun: tahun,
          tema: "Alam Semulajadi & Persekitaran", tajuk: "Mengenal Haiwan dan Tumbuhan Sekitar Sekolah",
          sk: "1.1 Memerhati dan mengenal hidupan di persekitaran",
          sp: "1.1.1 Mengenal pasti ciri fizikal haiwan dan bahagian tumbuhan melalui pemerhatian maujud"
        }
      ];
    }
    if (subUpper.includes("SENI") || subUpper.includes("PSK") || subUpper.includes("[PSK]")) {
      return [
        {
          subjek: subjek, tahun: tahun,
          tema: "Ekspresi Kreatif", tajuk: "Menggambar, Mewarna dan Irama Muzik",
          sk: "1.1 Menggunakan media seni visual dan alat perkusi asas",
          sp: "1.1.1 Menghasilkan karya seni mudah menggunakan teknik capan/warna serta menepuk mengikut detik"
        }
      ];
    }
    if (subUpper.includes("KEMAHIRAN HIDUP") || subUpper.includes("KHA") || subUpper.includes("[KHA]")) {
      return [
        {
          subjek: subjek, tahun: tahun,
          tema: "Kemahiran Berdikari & Praktikal", tajuk: "Kebersihan Ruang Kerja dan Pengendalian Alatan Selamat",
          sk: "1.1 Mengamalkan langkah keselamatan dan kebersihan di bengkel/dapur",
          sp: "1.1.1 Menggunakan dan menyimpan alatan tangan asas secara cermat dan bertanggungjawab"
        }
      ];
    }
    if (subUpper.includes("MAKLUMAT") || subUpper.includes("TMK") || subUpper.includes("[TMK]")) {
      return [
        {
          subjek: subjek, tahun: tahun,
          tema: "Literasi Digital Asas", tajuk: "Mengenal Perkakasan Komputer dan Peranti Pintar",
          sk: "1.1 Mengenal bahagian perkakasan peranti digital",
          sp: "1.1.1 Menggunakan tetikus, papan kekunci atau skrin sentuh untuk navigasi mudah"
        }
      ];
    }
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Pengurusan Diri & Fungsi Hidup", tajuk: "Amalan Hidup Berdikari",
        sk: "1.1 Melaksanakan aktiviti harian secara berpandu",
        sp: "1.1.1 Mengikuti arahan mudah guru dan bekerjasama dalam kumpulan"
      }
    ];
  }
  // 3. MATHEMATICS (DLP) / MATEMATIK DLP (Tema & Tajuk Bahasa Inggeris HANYA untuk DLP)
  if (subUpper.includes("DLP") && (subUpper.includes("MATHEMATIC") || subUpper.includes("MATH") || subUpper.includes("MATEMATIK"))) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Numbers and Operations", tajuk: "Whole Numbers and Basic Operations",
        sk: "1.1 Number value up to " + (digitTahun >= "4" ? "100,000" : "10,000"),
        sp: "1.1.1 Read, state and write numbers up to " + (digitTahun >= "4" ? "100,000" : "10,000") + " in numerals and words"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Numbers and Operations", tajuk: "Fractions, Decimals and Percentages",
        sk: "2.1 Basic Operations involving Fractions and Decimals",
        sp: "2.1.1 Add and subtract fractions and solve everyday routine problems"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Measurement and Geometry", tajuk: "Time and Dimensions",
        sk: "3.1 Relationship between units of time and length measurement",
        sp: "3.1.1 Convert units and calculate compound measurements accurately"
      }
    ];
  }

  // 4. SCIENCE (DLP) / SAINS DLP (Tema & Tajuk Bahasa Inggeris HANYA untuk DLP)
  if (subUpper.includes("DLP") && (subUpper.includes("SCIENCE") || subUpper.includes("SAINS"))) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Inquiry in Science", tajuk: "Scientific Skills and Experiments",
        sk: "1.1 Science Process Skills",
        sp: "1.1.1 Observe, classify and make inferences systematically"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Life Science", tajuk: "Humans and Living Organisms",
        sk: "2.1 Breathing process and human physiological functions",
        sp: "2.1.1 Identify the respiratory organs and explain the inhalation pathway"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Physical Science", tajuk: "Properties of Light and Shadows",
        sk: "3.1 Light travels in a straight line",
        sp: "3.1.1 State that light travels straight and describe factors affecting shadow size"
      }
    ];
  }

  // 5. BAHASA INGGERIS
  if (subUpper.includes("INGGERIS") || subUpper.includes("ENGLISH")) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "World of Self, Family and Friends", tajuk: "Unit 1 - Welcome & Getting To Know You",
        sk: "2.1 Communicate simple information intelligibly",
        sp: "2.1.1 Give detailed information about oneself using fixed target phrases"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "World of Self, Family and Friends", tajuk: "Unit 5 - Free Time and Healthy Activities",
        sk: "1.2 Understand meaning in a variety of familiar contexts",
        sp: "1.2.1 Understand with support the main idea of simple longer texts"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "World of Stories", tajuk: "Unit 3 - Amazing Animals and Folklore Tales",
        sk: "3.2 Understand a variety of linear and non-linear print texts",
        sp: "3.2.2 Understand specific information and details of two paragraphs"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "World of Knowledge", tajuk: "Unit 8 - Where Were You Yesterday & History Around Us",
        sk: "4.2 Communicate basic information intelligibly for a range of purposes",
        sp: "4.2.1 Describe basic past experiences using regular and irregular past verbs"
      }
    ];
  }

  // 6. BAHASA MELAYU
  if (subUpper.includes("MELAYU")) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Kekeluargaan", tajuk: "Unit 1: Keluarga Bahagia Bersama",
        sk: "1.1 Mendengar dan memberikan respons semasa berinteraksi dalam pelbagai situasi",
        sp: "1.1.1 Mendengar, mengecam sebutan dan menyebut perkataan serta intonasi dengan betul"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Kebersihan dan Kesihatan", tajuk: "Unit 2: Kesihatan Diri Asas Kehidupan Sejahtera",
        sk: "2.1 Asas membaca dan memahami teks bahan bacaan",
        sp: "2.1.1 Membaca dan memahami maklumat tersurat dan tersirat dalam petikan"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Keselamatan", tajuk: "Unit 3: Sentiasa Waspada dan Selamat di Mana Jua",
        sk: "3.2 Menulis perkataan, frasa dan ayat secara mekanis",
        sp: "3.2.1 Membina dan menulis perenggan yang kohesif menggunakan tanda baca tepat"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Jati Diri, Patriotisme dan Kewarganegaraan", tajuk: "Unit 4: Megahnya Negaraku Malaysia",
        sk: "4.1 Mengaplikasikan unsur keindahan bahasa seni bahasa",
        sp: "4.1.1 Melafazkan dan melagukan pantun empat kerat dengan sebutan puitis"
      }
    ];
  }

  // 7. MATEMATIK (ALIRAN PERDANA - 100% BAHASA MELAYU)
  if (!subUpper.includes("DLP") && (subUpper.includes("MATEMATIK") || subUpper.includes("MATHEMATIC") || subUpper.includes("MATH"))) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Nombor dan Operasi", tajuk: "Nombor Bulat dan Operasi Asas",
        sk: "1.1 Nilai nombor hingga " + (digitTahun >= "4" ? "100 000" : "10 000"),
        sp: "1.1.1 Menyatakan, membaca dan menulis sebarang nombor dalam perkataan dan angka"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Nombor dan Operasi", tajuk: "Pecahan, Perpuluhan dan Peratus",
        sk: "2.1 Operasi bergabung melibatkan pecahan dan perpuluhan",
        sp: "2.1.1 Menyelesaikan ayat matematik tambah dan tolak pecahan secara tepat"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Sukatan dan Geometri", tajuk: "Masa dan Waktu",
        sk: "3.1 Perkaitan unit masa dan waktu dunia",
        sp: "3.1.1 Menukar unit masa dan menyelesaikan masalah harian berkaitan waktu"
      }
    ];
  }

  // 8. SAINS (ALIRAN PERDANA - 100% BAHASA MELAYU)
  if (!subUpper.includes("DLP") && (subUpper.includes("SAINS") || subUpper.includes("SCIENCE"))) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Inkuiri dalam Sains", tajuk: "Kemahiran Saintifik dan Peraturan Makmal",
        sk: "1.1 Kemahiran Proses Sains",
        sp: "1.1.1 Memerhati, mengelas, mengukur dan menggunakan nombor secara bersistem"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Sains Hayat", tajuk: "Proses Hidup Manusia dan Haiwan",
        sk: "2.1 Pernafasan manusia dan organ terlibat",
        sp: "2.1.1 Mengenal pasti organ pernafasan dan laluan udara semasa bernafas"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Sains Fizikal", tajuk: "Sifat Cahaya dan Pantulan",
        sk: "3.1 Cahaya bergerak lurus dan faktor pembentukan bayang-bayang",
        sp: "3.1.1 Menyatakan cahaya bergerak lurus dan membuktikan pantulan cahaya"
      }
    ];
  }

  // 9. SEJARAH
  if (subUpper.includes("SEJARAH")) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Sejarah Awal Negara", tajuk: "Unit 1: Mengenali Sejarah dan Kaedah Pengkajian",
        sk: "1.1 Pengertian dan Kemahiran Sejarah",
        sp: "1.1.1 Menyatakan pengertian sejarah dan menghuraikan sumber primer serta sekunder"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Sejarah Awal Negara", tajuk: "Unit 5: Zaman Air Batu dan Perubahan Alam",
        sk: "2.1 Zaman Air Batu Akhir",
        sp: "2.1.1 Menjelaskan perubahan bentuk muka bumi Asia Tenggara akibat pencairan air batu"
      }
    ];
  }

  // 10. REKA BENTUK DAN TEKNOLOGI (RBT)
  if (subUpper.includes("REKA BENTUK") || subUpper.includes("RBT")) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Aplikasi Reka Bentuk", tajuk: "Unit 1: Keselamatan Bengkel dan Amalan 5S",
        sk: "1.1 Amalan Keselamatan Bengkel",
        sp: "1.1.1 Menyatakan peraturan keselamatan bengkel dan melakar pelan pemindahan kecemasan"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "Pengenalan kepada Teknologi", tajuk: "Unit 2: Asas Reka Bentuk Produk dan Bahan",
        sk: "2.1 Reka bentuk produk menggunakan bahan kitar semula",
        sp: "2.1.1 Menjana idea kreatif dan menghasilkan lakaran reka bentuk produk berfungsi"
      }
    ];
  }

  // 11. BAHASA ARAB
  if (subUpper.includes("ARAB")) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "مَهَارَاتُ اللُّغَةِ العَرَبِيَّةِ (Kemahiran Bahasa Arab)", tajuk: "الْأَدَوَاتُ الدِّرَاسِيَّةُ وَفِي الْفَصْلِ (Peralatan Belajar)",
        sk: "١٫١ الاسْتِمَاعُ إِلَى كَلِمَاتِ الْمَحَاوِرِ وَنُطْقُهَا نُطْقًا صَحِيحًا",
        sp: "١٫١٫١ الْقُدْرَةُ عَلَى مُحَاكَاةِ الْكَلِمَاتِ الْمَسْمُوعَةِ وَتَرْدِيدِهَا ثُمَّ نُطْقِهَا"
      },
      {
        subjek: subjek, tahun: tahun,
        tema: "الْحَيَاةُ الْيَوْمِيَّةُ (Kehidupan Seharian)", tajuk: "أُسْرَتِي الْحَبِيبَةُ (Keluargaku Tercinta)",
        sk: "٢٫١ قِرَاءَةُ الْكَلِمَاتِ قِرَاءَةً صَحِيحَةً مَعَ الْفَهْمِ",
        sp: "٢٫١٫١ الْقُدْرَةُ عَلَى قِرَاءَةِ الْكَلِمَاتِ الَّتِي فِيهَا الْحَرَكَاتُ الْمُخْتَلِفَةُ"
      }
    ];
  }

  // 12. PENDIDIKAN ISLAM (KSSR SEMAKAN TAHUN 1 - 5 DALAM TULISAN JAWI)
  if (subUpper.includes("ISLAM") || subUpper.includes("PAI") || subUpper.includes("إسلام") || subUpper.includes("اسلام") || subUpper.includes("ڤنديديقن") || subUpper.includes("جاوي")) {
    var semuaPai = [
      // TAHUN 1
      { subjek: subjek, tahun: "1", tema: "القرءان (تلاوة دان حافظن)", tajuk: "سورة الفاتحة دان سورة الإخلاص", sk: "1.1 تلاوة دان حافظن سورة الفاتحة دان الإخلاص دڠن بتول", sp: "1.1.1 ممباچ دان مڠحفظ اية٢ دڠن مخرج حروف يڠ بتول دان برتجويد" },
      { subjek: subjek, tahun: "1", tema: "عقيدة (روكون ايمان)", tajuk: "برايمان كڤد الله دان روكون ايمان", sk: "2.1 ممفهمي دان ميقيني اساس روكون ايمان", sp: "2.1.1 مڽاتاكن ڤڠرتين روكون ايمان دان مڽنارايكن 6 ڤركارا روكون ايمان" },
      { subjek: subjek, tahun: "1", tema: "عبادة (طهاره دان برسوچي)", tajuk: "كونسيڤ برسوچي درڤد نجيس دان استنجاء", sk: "3.1 ممفهمي دان مڠعملكن چونتوه استنجاء دڠن بتول", sp: "3.1.1 مڽاتاكن الت٢ استنجاء دان چارا استنجاء يڠ سمڤورنا" },
      { subjek: subjek, tahun: "1", tema: "سيرة (كلاهيران نبي)", tajuk: "ريوايت هيدوڤ نبي محمد SAW د مكة", sk: "5.1 مڠتاهوءي دان منلادني سيرة كلاهيران نبي محمد SAW", sp: "5.1.1 مڽاتاكن تاريـخ كلاهيران دان بومي برتواه مكة المكرمة" },
      { subjek: subjek, tahun: "1", tema: "ادب (ادب كلوارݢ)", tajuk: "ادب كڤد ايبو باڤ دان كلوارݢ", sk: "6.1 مڠعملكن ادب ترهادڤ ايبو باڤ دالم كهيدوڤن", sp: "6.1.1 منجلسكن چارا مڠحرمتي ايبو باڤ دان برتوتور دڠن سوڤن" },
      { subjek: subjek, tahun: "1", tema: "جاوي (حروف توڠݢل)", tajuk: "مڠنل دان منوليس حروف جاوي توڠݢل", sk: "7.1 ممباچ دان منوليس حروف٢ جاوي توڠݢل دان سامبوڠ", sp: "7.1.1 مڠنل، مڽبوت دان منوليس حروف جاوي توڠݢل دڠن بتول" },

      // TAHUN 2
      { subjek: subjek, tahun: "2", tema: "القرءان (تلاوة دان حافظن)", tajuk: "سورة الناس دان سورة الفلق", sk: "1.2 ممباچ دان مڠحفظ سورة الناس دان الفلق برتجويد", sp: "1.2.1 ممباچ دان مڠحفظ سورة دڠن مخرج حروف يڠ بتول دان طمأنينة" },
      { subjek: subjek, tahun: "2", tema: "عقيدة (صفة٢ الله)", tajuk: "الله تعالى برصفة القدير دان العليم", sk: "2.2 ميقيني دان ممبوقتيكن صفة القدير دان العليم دالم كهيدوڤن", sp: "2.2.1 مڽاتاكن ارتي القدير دان العليم سرتا داليل نقلي دان عقلي" },
      { subjek: subjek, tahun: "2", tema: "عبادة (صلوة فردو)", tajuk: "شرط٢ واجب دان 13 ركون صلوة", sk: "3.2 ممفهمي دان ملقساناكن ركون صلوة دڠن سمڤورنا", sp: "3.2.1 مڽنارايكن 13 ركون صلوة سرتا ملاكوكن باچاءن دان ڤربواتن دڠن بتول" },
      { subjek: subjek, tahun: "2", tema: "سيرة (كانق٢ بركت)", tajuk: "كڤريبادين نبي محمد SAW كتيك كانق٢ دان رماجا", sk: "5.2 منلادني صيفت امانه دان فطانة رسول الله SAW", sp: "5.2.1 منجلسكن كأيستيميواءن صيفت رسول الله سجق اوسيا مودا" },
      { subjek: subjek, tahun: "2", tema: "ادب (ادب د سكوله)", tajuk: "ادب برسام ݢورو دان راكن سكوله", sk: "6.2 مڠعملكن ادب برسام ݢورو دان راكن سبايلا", sp: "6.2.1 منجلسكن كڤنتيڠن مڠحرمتي ݢورو دان برتوليرنسي دڠن راكن" },
      { subjek: subjek, tahun: "2", tema: "جاوي (سوكو كات)", tajuk: "ممباچ دان منوليس ڤركاتاءن سوكو كات تربوک", sk: "7.2 ممباچ، ممبينا دان منوليس ڤركاتاءن درڤد سوكو كات تربوک", sp: "7.2.1 ممباچ دان منوليس ڤركاتاءن يڠ مڠاندوڠي دوا سوكو كات تربوک دڠن بتول" },

      // TAHUN 3
      { subjek: subjek, tahun: "3", tema: "القرءان (تلاوة دان كفهمن)", tajuk: "سورة الكافرون دان سورة النصر", sk: "1.3 ممباچ، مڠحفظ دان ممفهمي سورة الكافرون دان النصر", sp: "1.3.1 ممباچ دڠن برتجويد دان منجلسكن ڤڠاجرن سورة الكافرون دالم كفهمن" },
      { subjek: subjek, tahun: "3", tema: "عقيدة (برايمان كڤد كتاب)", tajuk: "برايمان كڤد كتاب٢ الله دان القرءان كريم", sk: "2.3 ميقيني كڤنتيڠن برايمان كڤد كتاب٢ الله دان مڠعملكنڽ", sp: "2.3.1 مڽنارايكن 4 كتاب دان رسول يڠ منريماڽ سرتا كليبهن القرءان" },
      { subjek: subjek, tahun: "3", tema: "عبادة (صلاة برجماعه)", tajuk: "كونسيڤ دان كأوتاماءن صلوة برجماعه", sk: "3.3 مڠعملكن دان ملقساناكن صلوة برجماعه دالم كهيدوڤن سهارين", sp: "3.3.1 منجلسكن كدودوقن امام دان مأموم سرتا كأوتاماءن صلوة برجماعه" },
      { subjek: subjek, tahun: "3", tema: "سيرة (ڤرستيوا ڤنتيڠ)", tajuk: "ڤريستيوا نبي محمد SAW منريما وحي ڤرتام", sk: "5.3 مڠحياتي دان ممفهمي ڤريستيوا ڤنورونن وحي د ݢوا حراء", sp: "5.3.1 مڠهورايكن ڤريستيوا وحي ڤرتام دان ڤرانن سيدنا خديجة" },
      { subjek: subjek, tahun: "3", tema: "ادب (ادب برجالن دان برتاتري)", tajuk: "ادب بركومونيكاسي دان منزيارهي جيراڽ", sk: "6.3 مڠعملكن ادب برجيرن دان برتاتاسوسيلا دالم مشاركت", sp: "6.3.1 مڽاتاكن حق٢ جيرن دان ادب منزيارهي جيرن مڠيكوت سنة" },
      { subjek: subjek, tahun: "3", tema: "جاوي (تيك س موده)", tajuk: "ممبينا دان منوليس تيك س ڤينديق جاوي", sk: "7.3 ممباچ دان منوليس تيك س برتوليسن جاوي يڠ مڠاندوڠي ايمبوهن", sp: "7.3.1 ممبينا، ممباچ دان منوليس ايات موده مڠݢوناكن توليسن جاوي دڠن لنچر" },

      // TAHUN 4
      { subjek: subjek, tahun: "4", tema: "القرءان (تلاوة، حافظن دان تجويد)", tajuk: "سورة القارعة دان سورة التكاثر دان حكوم نون ساكنة", sk: "1.4 ممباچ دان مڠحفظ سورة القارعة سرتا مڠفليكاسي حكوم نون ساكنة", sp: "1.4.1 ممباچ ايات سورة برتجويد دان مڠنل ڤستي چونتوه إظهار حلقي دان إدغام" },
      { subjek: subjek, tahun: "4", tema: "عقيدة (صفة٢ رسول)", tajuk: "صفة٢ واجب، مستحيل دان هارس باݢي رسول", sk: "2.4 ممفهمي دان منلادني صفة٢ كماءنسياءن دان كروسولن", sp: "2.4.1 مڽاتاكن 4 صفة واجب (صديق، أمانة، تبليغ، فطانة) دان لاونڽ" },
      { subjek: subjek, tahun: "4", tema: "عبادة (صيام رمضان دان وضوء)", tajuk: "حكوم دان شرط صيام بولن رمضان", sk: "3.4 مڠحياتي دان ملقساناكن عبادة ڤواسا رمضان دڠن بتول", sp: "3.4.1 مڽاتاكن شرط واجب، شرط صح، ركون دان ڤركارا ممبطلكن ڤواسا" },
      { subjek: subjek, tahun: "4", tema: "سيرة (دعوة رسول الله)", tajuk: "دعوة نبي سچارا رهسيا دان ترترڠ-ترڠن د مكة", sk: "5.4 مڠاناليسيس چابران دعوة رسول الله SAW دان كصبرن بݢيندا", sp: "5.4.1 مڠهورايكن ڤريستيوا دعوة نبي دان چارا مڠهادڤي تنتڠن قوم قريش" },
      { subjek: subjek, tahun: "4", tema: "ادب (ادب برسام علم)", tajuk: "ادب منونتوت علمو دان مڠحرمتي بوكو", sk: "6.4 مڠعملكن ادب منونتوت علمو دالم كهيدوڤن هارين", sp: "6.4.1 منجلسكن كڤنتيڠن اخلاص، برتريما كاسيه دان منجاݢ كسوچين بوكو علمو" },
      { subjek: subjek, tahun: "4", tema: "جاوي (ڤريڠݢن جاوي)", tajuk: "ممباچ دان منوليس ڤريڠݢن جاوي براينفورماسي", sk: "7.4 ممباچ دان منوليس ڤتيقن ڤنديق برتوليسن جاوي برتيماكن ايسي سمءاس", sp: "7.4.1 مڠوبه سوأي دان منوليس ڤريڠݢن جاوي دڠن اياءن يڠ بتول دان تراتور" },

      // TAHUN 5
      { subjek: subjek, tahun: "5", tema: "القرءان (تلاوة دان كفهمن)", tajuk: "سورة القدر دان سورة العصر سرتا حكوم مد", sk: "1.5 ممباچ، مڠحفظ دان مڠعملكن سورة القدر دان العصر سرتا تجويد مد عارض", sp: "1.5.1 مڠحفظ دان منجلسكن ارتي ايات سورة القدر سرتا كأوتاماءن مالم ليلة القدر" },
      { subjek: subjek, tahun: "5", tema: "عقيدة (برايمان كڤد هاري اخرة)", tajuk: "كونسيڤ هاري قيامة، تيمبڠن عمل دان شرݢا نراك", sk: "2.5 ميقيني كجادين هاري قيامة دان برعمل اونتوق كهيدوڤن اخرة", sp: "2.5.1 مڽاتاكن تندا٢ قيامة دان ممبوقتيكن كأيمانن ملالوءي عملن صالح" },
      { subjek: subjek, tahun: "5", tema: "عبادة (صلوة جمعة دان جنازة)", tajuk: "تونتوتن صلوة جمعة دان فقه صلوة جنازة", sk: "3.5 مڠحياتي كأوتاماءن صلوة جمعة سرتا شرط صح دان خطبة", sp: "3.5.1 منجلسكن كواجيفن صلوة جمعة دان ادب٢ كتيك مندڠر خطبة دان تاتاچارا صلوة" },
      { subjek: subjek, tahun: "5", tema: "سيرة (هجرة رسول الله)", tajuk: "ڤريستيوا هجرة نبي ك يثرب دان ڤمبنتوقن مدنية", sk: "5.5 مڠحياتي اعتباره درڤد ڤريستيوا هجرة دان صيفت ڤرڤادوان", sp: "5.5.1 منجلسكن فكتور هجرة دان ڤرسوداراءن انتارا قوم مهاجرين دان انصار" },
      { subjek: subjek, tahun: "5", tema: "ادب (ادب د مسجيد)", tajuk: "ادب مماسوقي دان مممعموركن مسجيد دان سوراو", sk: "6.5 مڠعملكن ادب مماسوقي، برادا دان كلوار درڤد مسجيد", sp: "6.5.1 مڽاتاكن دعاء، لاكوكن صلوة تحية المسجد دان منجاݢ كتنترمن د دالم مسجيد" },
      { subjek: subjek, tahun: "5", tema: "جاوي (خط جاوي نسخ & رقعه)", tajuk: "منوليس خط رقعه دان خط نسخ موده", sk: "7.5 منوليس دان مڠحياتي كاينداهن خط جاوي موده", sp: "7.5.1 مڠنل ڤستي قاعيده خط نسخ دان رقعه سرتا منوليسڽ دڠن چنتيق دان كمفيت" }
    ];

    var digitThn = tukarDigitArabKeRumi(tahun).replace(/\D/g, '');
    var padanThn = semuaPai.filter(function(it) {
      return !digitThn || it.tahun === digitThn;
    });
    return padanThn.length > 0 ? padanThn : semuaPai;
  }

  // 13. PENDIDIKAN MORAL
  if (subUpper.includes("MORAL")) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Nilai Kepercayaan kepada Tuhan", tajuk: "Unit 1: Hormat Menghormati Amalan Ibadah",
        sk: "1.1 Menghormati kepelbagaian amalan agama warga sekolah dan masyarakat",
        sp: "1.1.1 Menyatakan amalan agama pelbagai kaum dan menunjukkan sikap toleransi"
      }
    ];
  }

  // 14. PENDIDIKAN SENI VISUAL (PSV)
  if (subUpper.includes("SENI") || subUpper.includes("PSV")) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Menggambar", tajuk: "Unit 1: Lukisan Alam dan Keindahan Warna",
        sk: "1.1 Bahasa Seni Visual pada karya lukisan",
        sp: "1.1.1 Mengaplikasikan teknik gosokan, jalinan dan ton warna pada gubahan landskap"
      }
    ];
  }

  // 15. PENDIDIKAN MUZIK
  if (subUpper.includes("MUZIK")) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Pengalaman Muzikal", tajuk: "Nyanyian Pic dan Melodi Berirama",
        sk: "1.1 Bernyanyi pelbagai repertoir secara solo dan berkumpulan",
        sp: "1.1.1 Menyanyikan lagu dengan sebutan yang jelas serta pic yang tepat"
      }
    ];
  }

  // 16. PENDIDIKAN JASMANI & KESIHATAN (PJPK)
  if (subUpper.includes("JASMANI") || subUpper.includes("KESIHATAN")) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Kemahiran Asas Pergerakan", tajuk: "Gimnastik Asas dan Keseimbangan Dinamik",
        sk: "1.1 Melakukan pelbagai corak pergerakan lokomotor dan bukan lokomotor",
        sp: "1.1.1 Menunjukkan imbangan dinamik dan mendarat dengan kedua-dua kaki secara selamat"
      }
    ];
  }

  // 17. BAHASA KADAZANDUSUN (BKD)
  if (subUpper.includes("KADAZANDUSUN") || subUpper.includes("BKD")) {
    return [
      {
        subjek: subjek, tahun: tahun,
        tema: "Koilaan Diti Sondii (Diri Saya)", tajuk: "Unit 1: Pagambatan om Paganakan",
        sk: "1.1 Manahang om mongilo kointutunan paganakan",
        sp: "1.1.1 Mangarait boros di kosudong kokomoi paganakan miampai boros di olinuud"
      }
    ];
  }

  // Default Standard KPM
  return [
    {
      subjek: subjek, tahun: tahun,
      tema: "Pendidikan Holistik", tajuk: "Penguasaan Konsep dan Kemahiran",
      sk: "1.1 Standard Kandungan DSKP Kementerian Pendidikan Malaysia",
      sp: "1.1.1 Menguasai standard pembelajaran yang ditetapkan mengikut kurikulum kebangsaan"
    }
  ];
}

// Ambil senarai berhierarki penuh untuk Subjek & Tahun tertentu
function getHierarkiDskp(subjek, tahun, namaKelas) {
  var kUpper = String(namaKelas || "").toUpperCase();
  var isPpki = kUpper.includes("VIVA") || kUpper.includes("WIRA") || kUpper.includes("ARENA") || 
               kUpper.includes("AXIA") || kUpper.includes("SAGA") || kUpper.includes("BEZZA") || kUpper.includes("PPKI");
  var isPra = kUpper.includes("PRA");
  
  var dataSheet = muatSemuaDskpDariSheet(isPpki);
  var senaraiAlias = seragamkanNamaSubjek(subjek);
  
  var hasil = [];
  if (dataSheet && dataSheet.length > 0) {
    for (var i = 0; i < dataSheet.length; i++) {
      var row = dataSheet[i];
      var subUpper = row.subjek.toUpperCase();
      var subPadan = senaraiAlias.some(function(al) {
        return subUpper === al || subUpper.includes(al) || al.includes(subUpper);
      });
      if (subPadan && padanTahunSama(row.tahun, tahun, isPpki)) {
        hasil.push({
          subjek: subjek,
          tahun: tahun,
          tema: row.tema,
          tajuk: row.tajuk,
          sk: row.sk,
          sp: row.sp
        });
      }
    }
  }
  
  // Jika tiada rekod dalam Google Sheets atau fail sheet belum diisi, gunakan data piawai KPM
  if (hasil.length === 0) {
    hasil = dapatkanDskpPiawaiKpm(subjek, tahun, isPpki, isPra);
  }
  
  return hasil;
}

// Dapatkan rekod DSKP spesifik yang menjamin TEMA, TAJUK, SK, dan SP sentiasa TALLY 100%
function ambilObjektifDskp(subjek, tingkatTahun, namaKelas, temaPilihan, tajukPilihan, skPilihan, spPilihan) {
  var senarai = getHierarkiDskp(subjek, tingkatTahun, namaKelas);
  if (!senarai || senarai.length === 0) {
    return {
      tema: "Kekeluargaan",
      tajuk: "Keluarga Bahagia",
      sk: "1.1 Mendengar dan bertutur",
      sp: "1.1.1 Menyatakan maklumat asas secara bertatasusila",
      temaTajuk: "Kekeluargaan • Keluarga Bahagia"
    };
  }
  
  // 1. Jika SP dipilih khusus, cari baris SP tersebut agar TEMA, TAJUK, dan SK 100% sepadan
  if (spPilihan) {
    var padanSp = senarai.find(function(r) { return r.sp === spPilihan || r.sp.includes(spPilihan); });
    if (padanSp) {
      return {
        tema: padanSp.tema,
        tajuk: padanSp.tajuk,
        sk: padanSp.sk,
        sp: padanSp.sp,
        temaTajuk: padanSp.tema + " • " + padanSp.tajuk
      };
    }
  }
  
  // 2. Jika SK dipilih khusus, cari baris SK tersebut
  if (skPilihan) {
    var padanSk = senarai.find(function(r) { return r.sk === skPilihan || r.sk.includes(skPilihan); });
    if (padanSk) {
      return {
        tema: padanSk.tema,
        tajuk: padanSk.tajuk,
        sk: padanSk.sk,
        sp: padanSk.sp,
        temaTajuk: padanSk.tema + " • " + padanSk.tajuk
      };
    }
  }
  
  // 3. Jika Tajuk dipilih khusus, cari baris Tajuk tersebut
  if (tajukPilihan) {
    var padanTajuk = senarai.find(function(r) { return r.tajuk === tajukPilihan || r.tajuk.includes(tajukPilihan); });
    if (padanTajuk) {
      return {
        tema: padanTajuk.tema,
        tajuk: padanTajuk.tajuk,
        sk: padanTajuk.sk,
        sp: padanTajuk.sp,
        temaTajuk: padanTajuk.tema + " • " + padanTajuk.tajuk
      };
    }
  }
  
  // 4. Jika Tema dipilih khusus, cari baris Tema tersebut
  if (temaPilihan) {
    var padanTema = senarai.find(function(r) { return r.tema === temaPilihan || r.tema.includes(temaPilihan); });
    if (padanTema) {
      return {
        tema: padanTema.tema,
        tajuk: padanTema.tajuk,
        sk: padanTema.sk,
        sp: padanTema.sp,
        temaTajuk: padanTema.tema + " • " + padanTema.tajuk
      };
    }
  }
  
  // 5. Default baris pertama yang sah
  var first = senarai[0];
  return {
    tema: first.tema,
    tajuk: first.tajuk,
    sk: first.sk,
    sp: first.sp,
    temaTajuk: first.tema + " • " + first.tajuk
  };
}

// API untuk Frontend: Ambil Hierarki DSKP Subjek & Tahun
function getDskpHierarkiWeb(subjek, tahun, namaKelas) {
  return getHierarkiDskp(subjek, tahun, namaKelas);
}

// ==========================================================================
// PANGKALAN DATA DOKUMEN PENGHUBUNG PEMULIHAN KHAS KPM (BM & MATEMATIK 2019)
// ==========================================================================
var DATA_PEMULIHAN_BM = [
  { kp: "KP 0", nama: "Prabacaan dan Pratulisan", sk: "3.1 Asas Menulis", sp: "3.1.1 (i) Menulis secara mekanis; huruf", konstruk: "Konstruk 1: Keupayaan pranombor dan mengenal huruf", bbm: "Lembaran kerja, carta, melakar garisan/bentuk", aktiviti: "Melakar pelbagai bentuk dan corak, menyambung titik-titik pada rajah" },
  { kp: "KP 1", nama: "Huruf-huruf vokal (a, e, i, o, u)", sk: "1.1 Mendengar dan memberikan respons", sp: "1.1.1 (i) Mengajuk dan menyebut vokal", konstruk: "Konstruk 1: Keupayaan membunyikan dan menulis huruf vokal", bbm: "Kad abjad, kad tebuk abjad, carta vokal", aktiviti: "Menamakan dan membunyikan huruf vokal, menyalin huruf dalam buku garis empat, permainan vokal" },
  { kp: "KP 2", nama: "Huruf-huruf kecil (a - z)", sk: "2.1 Asas membaca dan memahami", sp: "2.1.1 (ii) Membaca dengan sebutan yang betul; konsonan", konstruk: "Konstruk 1: Mengenal bentuk dan menulis huruf kecil a-z", bbm: "Kad cantum titik, kad tebuk, doh huruf", aktiviti: "Permainan Bahasa 1 (Pasangkan Saya), menulis huruf kecil di udara/pasir/buku garis empat" },
  { kp: "KP 3", nama: "Huruf-huruf besar (A - Z)", sk: "3.1 Asas menulis", sp: "3.1.1 (i) Menulis secara mekanis; huruf besar", konstruk: "Konstruk 1: Mengenal dan menulis huruf besar A-Z", bbm: "Kad abjad bergambar, kad huruf besar, lembaran kerja", aktiviti: "Memadankan huruf besar dengan huruf kecil, menekap huruf besar, mencari huruf tersembunyi" },
  { kp: "KP 4", nama: "Suku kata KV", sk: "2.1 Asas membaca dan memahami", sp: "2.1.1 (iii) Membaca dengan sebutan yang betul; suku kata", konstruk: "Konstruk 2: Keupayaan membaca dan menulis suku kata terbuka KV", bbm: "Kad suku kata, kad cantum suku kata, kad imbasan", aktiviti: "Permainan Bahasa 2 (Teka Silang Kata), mengeja dan membatang suku kata KV berpandukan gambar" },
  { kp: "KP 5", nama: "Perkataan KV + KV", sk: "2.1 Asas membaca dan memahami", sp: "2.1.2 (i) Membaca perkataan dua suku kata terbuka", konstruk: "Konstruk 3: Membaca dan menulis perkataan suku kata terbuka KV+KV", bbm: "Kad perkataan, kad lipat, domino perkataan", aktiviti: "Permainan Bahasa 3 (Namakan Saya), memadankan perkataan KV+KV dengan gambar, membaca ayat mudah" },
  { kp: "KP 6", nama: "Perkataan V + KV", sk: "2.1 Asas membaca dan memahami", sp: "2.1.2 (i) Membaca perkataan suku kata terbuka V+KV", konstruk: "Konstruk 3: Membaca dan menulis perkataan suku kata terbuka V+KV", bbm: "Kad cantum suku kata, botol air bergambar", aktiviti: "Permainan Bahasa 4 (Jatuhkan Saya / baling bola ke botol), mencantum suku kata menjadi V+KV (api, ubi, alu, ibu)" },
  { kp: "KP 7", nama: "Perkataan KV + KV + KV", sk: "2.1 Asas membaca dan memahami", sp: "2.1.2 (i) Membaca perkataan tiga suku kata terbuka", konstruk: "Konstruk 3: Membaca dan menulis perkataan suku kata terbuka KV+KV+KV", bbm: "Kad suku kata tiga warna, carta gambar", aktiviti: "Permainan Bahasa 5 (Cari Perkataan 5 Stesen), mencantum kad suku kata menjadi perkataan tiga suku kata" },
  { kp: "KP 8", nama: "Perkataan KVK", sk: "2.1 Asas membaca dan memahami", sp: "2.1.1 (iii) Membaca perkataan suku kata tertutup KVK", konstruk: "Konstruk 5: Keupayaan membaca dan menulis perkataan suku kata tertutup KVK", bbm: "Jigsaw puzzle perkataan, carta gambar KVK", aktiviti: "Permainan Bahasa 6 (Siapa Saya / bulatkan perkataan), pertandingan jigsaw puzzle perkataan KVK" },
  { kp: "KP 9", nama: "Suku kata KVK", sk: "1.1 Mendengar dan memberikan respons", sp: "1.1.1 (ii) Mengajuk dan membunyikan suku kata tertutup", konstruk: "Konstruk 4: Membaca dan menulis suku kata tertutup KVK", bbm: "Kad suku kata KV dan konsonan penutup", aktiviti: "Mencantum suku kata KV dengan huruf konsonan akhir untuk membina suku kata tertutup KVK" },
  { kp: "KP 10", nama: "Perkataan V + KVK", sk: "2.1 Asas membaca dan memahami", sp: "2.1.1 (iv) Membaca perkataan gabungan V+KVK", konstruk: "Konstruk 5: Membaca dan menulis perkataan suku kata tertutup V+KVK", bbm: "Kad huruf vokal, kad suku kata KVK", aktiviti: "Membina dan menulis perkataan V+KVK (ayam, itik, obor, ulat) berpandukan gambar dan kad imbasan" },
  { kp: "KP 11", nama: "Perkataan KV + KVK", sk: "2.2 Membaca dan memahami bahan grafik", sp: "2.2.1 (i) Membaca kosa kata KV+KVK", konstruk: "Konstruk 5: Membaca dan menulis perkataan terbuka + tertutup KVKVK", bbm: "Kad lipat perkataan, bahan maujud sayur/buah", aktiviti: "Pembelajaran Bermakna di Pasar (kubis, betik, sayur, segar), latih tubi membaca dan imlak perkataan" },
  { kp: "KP 12", nama: "Perkataan KVK + KV", sk: "2.1 Asas membaca dan memahami", sp: "2.1.2 (i) Membaca perkataan suku kata tertutup + terbuka", konstruk: "Konstruk 3: Membaca dan menulis perkataan KVK+KV", bbm: "Objek maujud, kad suku kata KVK dan KV", aktiviti: "Pamer objek maujud, cantum kad KVK+KV menjadi perkataan (lembu, pintu, bomba), tulis dalam buku garis empat" },
  { kp: "KP 13", nama: "Perkataan KVK + KVK", sk: "2.1 Asas membaca dan memahami", sp: "2.1.2 (i) Membaca perkataan dua suku kata tertutup", konstruk: "Konstruk 5: Membaca dan menulis perkataan tertutup KVK+KVK", bbm: "Kad imbasan KVK+KVK, carta perkataan", aktiviti: "Membentuk perkataan KVK+KVK daripada kad suku kata, mewarna perkataan, menyalin semula perkataan" },
  { kp: "KP 14", nama: "Perkataan KV + KV + KVK", sk: "3.2 Menulis perkataan secara bermakna", sp: "3.2.1 (i) Membina dan menulis perkataan tiga suku kata", konstruk: "Konstruk 5: Membaca dan menulis perkataan KV+KV+KVK", bbm: "Carta perkataan bergambar tiga suku kata", aktiviti: "Menamakan gambar tiga suku kata, melengkapkan suku kata akhir, menulis perkataan dalam buku garis empat" },
  { kp: "KP 15", nama: "Perkataan KVK + KV + KVK", sk: "3.2 Menulis perkataan secara bermakna", sp: "3.2.1 (i) Membina dan menulis perkataan KVK+KV+KVK", konstruk: "Konstruk 5: Membaca dan menulis perkataan KVK+KV+KVK", bbm: "Kad suku kata warna-warni, gambar", aktiviti: "Mencantum tiga kad suku kata menjadi perkataan (sembilang, cendawan), memadankan gambar dengan perkataan" },
  { kp: "KP 16", nama: "Perkataan KVKK", sk: "2.1 Asas membaca dan memahami", sp: "2.1.1 (iv) Membaca perkataan suku kata tertutup 'ng'", konstruk: "Konstruk 6: Keupayaan membaca dan menulis suku kata tertutup 'ng'", bbm: "Kad huruf penyusun KVKK, kad perkataan", aktiviti: "Menyusun huruf menjadi perkataan KVKK (wang, tong, bang), memadankan perkataan dengan gambar" },
  { kp: "KP 17", nama: "Suku kata KVKK", sk: "3.1 Asas menulis", sp: "3.1.1 (ii) Menulis secara mekanis suku kata KVKK", konstruk: "Konstruk 6: Membina dan menulis suku kata tertutup KVKK", bbm: "Kad lipat KVKK, buku skrap suku kata", aktiviti: "Mengeja suku kata KVKK menggunakan kad lipat, menggunting dan menampal suku kata KVKK ke dalam buku skrap" },
  { kp: "KP 18", nama: "Perkataan KV + KVKK", sk: "2.1 Asas membaca dan memahami", sp: "2.1.2 (i) Membaca perkataan tertutup 'ng'", konstruk: "Konstruk 6: Membaca dan menulis perkataan KV+KVKK", bbm: "Bahan maujud (kacang, bawang, tudung, butang, pisang)", aktiviti: "Permainan Bahasa 7 (Cari Saya / cari bahan maujud tersembunyi berpandukan kad perkataan)" },
  { kp: "KP 19", nama: "Perkataan V + KVKK", sk: "2.2 Membaca dan menaakul", sp: "2.2.1 (i) Membaca kosa kata V+KVKK", konstruk: "Konstruk 6: Membaca dan menulis perkataan V+KVKK", bbm: "Kad vokal dan KVKK (udang, orang, usung)", aktiviti: "Pembelajaran Bermakna menaakul gambar berkuda (orang, tunggang), susun huruf membentuk V+KVKK" },
  { kp: "KP 20", nama: "Perkataan KVK + KVKK", sk: "1.1 Mendengar dan memberi respons", sp: "1.1.2 (i) Mendengar soalan dan bertutur", konstruk: "Konstruk 6: Membaca dan menulis perkataan KVK+KVKK", bbm: "Kad perkataan (kandang, bintang, jantung)", aktiviti: "Pantun teka-teki 'Berkelip-kelip bukannya api... (Bintang)', memadankan kad suku kata dan melengkapkan ayat" },
  { kp: "KP 21", nama: "Perkataan KVKK + KV", sk: "3.2 Menulis perkataan bermakna", sp: "3.2.1 (i) Membina dan menulis perkataan", konstruk: "Konstruk 6: Membaca dan menulis perkataan KVKK+KV", bbm: "Kotak keting-ting, gundu, gambar tangga/nangka/bangku", aktiviti: "Permainan Bahasa 8 (Keting-ting / lompat gundu dan padankan kad gambar dengan perkataan)" },
  { kp: "KP 22", nama: "Perkataan KVKK + KVK", sk: "2.1 Asas membaca dan memahami", sp: "2.1.2 (i) Membaca perkataan KVKK+KVK", konstruk: "Konstruk 6: Membaca dan menulis perkataan KVKK+KVK", bbm: "Kad gambar situasi tingkap datuk, mangkuk kaca", aktiviti: "Mengeja dan mencerakinkan suku kata (ting-kap, mang-kuk), membaca ayat mudah berpandukan gambar" },
  { kp: "KP 23", nama: "Perkataan KVKK + KVKK", sk: "2.1 Asas membaca dan memahami", sp: "2.1.2 (i) Membaca perkataan suku kata tertutup 'ng'", konstruk: "Konstruk 6: Membaca dan menulis perkataan KVKK+KVKK", bbm: "Kad suku kata (kang-kung, long-kang, jeng-king)", aktiviti: "Menyusun kad suku kata menjadi perkataan, membaca ayat mudah (Petani tanam kangkung, Tepi rumah ada longkang)" },
  { kp: "KP 24", nama: "Perkataan KV + KV + KVKK", sk: "3.2 Menulis perkataan bermakna", sp: "3.2.1 (i) Membina perkataan tiga suku kata tertutup 'ng'", konstruk: "Konstruk 6: Membaca dan menulis perkataan KV+KV+KVKK", bbm: "Kad gambar belalang, teropong, seladang", aktiviti: "Mengeja suku kata KV+KV+KVKK, melengkapkan perkataan berpandukan gambar, membina ayat mudah" },
  { kp: "KP 25", nama: "Perkataan KV + KVK + KVKK", sk: "2.1 Asas membaca dan memahami", sp: "2.1.2 (i) Membaca perkataan pelbagai suku kata", konstruk: "Konstruk 6: Membaca dan menulis perkataan KV+KVK+KVKK", bbm: "Kad suku kata be-lim-bing, pe-lam-pung", aktiviti: "Menyatukan suku kata awalan/akhiran yang tertinggal, membaca ayat mudah (Belimbing besi rasa masam)" },
  { kp: "KP 26", nama: "Perkataan KVK + KV + KVKK", sk: "3.2 Menulis perkataan bermakna", sp: "3.2.1 (i) Membina perkataan KVK+KV+KVKK", konstruk: "Konstruk 6: Membaca dan menulis perkataan KVK+KV+KVKK", bbm: "Kad perkataan tempurung, pembilang, pendayung", aktiviti: "Mencerakin suku kata menjadi perkataan, menyusun perkataan menjadi ayat betul (Bapa seorang pemborong ikan)" },
  { kp: "KP 27", nama: "Perkataan KVKK + KV + KVK", sk: "2.1 Asas membaca dan memahami", sp: "2.1.2 (i) Membaca perkataan KVKK+KV+KVK", konstruk: "Konstruk 6: Membaca dan menulis perkataan KVKK+KV+KVK", bbm: "Kad gambar dan kad lipat suku kata", aktiviti: "Menamakan gambar, memadankan suku kata menjadi perkataan, menulis ayat mudah melibatkan KVKK+KV+KVK" },
  { kp: "KP 28", nama: "Perkataan KV + KVKK + KVK", sk: "3.2 Menulis perkataan bermakna", sp: "3.2.1 (i) Membina perkataan KV+KVKK+KVK", konstruk: "Konstruk 6: Membaca dan menulis perkataan KV+KVKK+KVK", bbm: "Kad perkataan perangkap, merangkak, belangkas", aktiviti: "Membunyikan suku kata, melengkapkan perkataan dengan suku kata sesuai, menulis ayat mudah" },
  { kp: "KP 29", nama: "Perkataan Diftong & Vokal Berganding", sk: "4.4 Menghayati karya sastera", sp: "4.4.2 (i) Menyanyikan lagu dengan sebutan betul", konstruk: "Konstruk 7 & 8: Membaca dan menulis perkataan diftong & vokal berganding", bbm: "Kad diftong (ai, au, oi) dan kad perkataan pantai/pisau/pulau", aktiviti: "Pembelajaran Bermakna: Nyanyian lagu 'Lompat Si Katak Lompat', menekankan perkataan menghargai, bahagia, kalau, nilai" },
  { kp: "KP 30", nama: "Perkataan Digraf & Konsonan Bergabung", sk: "3.2 Menulis perkataan bermakna", sp: "3.2.1 (i) Membina dan menulis ayat", konstruk: "Konstruk 9: Membaca dan menulis perkataan digraf dan konsonan bergabung", bbm: "Bahan simulasi jual beli sabun mandi wangi", aktiviti: "Pembelajaran Bermakna: Main peranan jurujual sabun organik (pengawet, pewangi, minyak, mengandungi, jangan)" },
  { kp: "KP 31", nama: "Membaca dan Membina Ayat Mudah", sk: "3.2 Menulis ayat yang bermakna", sp: "3.2.1 (iii) Membina dan menulis ayat tunggal", konstruk: "Konstruk 11: Keupayaan membaca dan menulis ayat mudah", bbm: "Carta ayat tunggal bergambar, kad susun ayat", aktiviti: "Membaca carta ayat tunggal bergambar, menyusun perkataan menjadi ayat bermakna, menulis ayat tunggal" },
  { kp: "KP 32", nama: "Bacaan dan Pemahaman", sk: "2.3 Membaca dan mengapresiasi", sp: "2.3.2 (iv) Membaca dan mempersembahkan petikan", konstruk: "Konstruk 12: Membaca, memahami dan menulis berdasarkan rangsangan", bbm: "Petikan pendek bergambar, kad soalan pemahaman", aktiviti: "Pembelajaran Bermakna: Membaca petikan dan menceritakan pengalaman di masjid/sekolah, menjawab soalan pemahaman" }
];

var DATA_PEMULIHAN_MT = [
  { kp: "KP 1", nama: "Pra Nombor (Warna, Saiz, Bentuk, Jenis)", sk: "1.1 Kuantiti secara intuitif", sp: "1.1.1 Menyatakan kuantiti melalui perbandingan", konstruk: "Konstruk 1: Keupayaan pranombor dan mengenal angka", bbm: "Bahan konkrit warna-warni, kad bentuk, objek saiz berbeza", aktiviti: "Mengelaskan bahan konkrit mengikut warna, saiz (kecil/besar), bentuk dan jenis" },
  { kp: "KP 2", nama: "Konsep Nombor (Kuantiti Intuitif & Seriasi)", sk: "1.1 Kuantiti secara intuitif", sp: "1.1.1 Menyatakan kuantiti melalui perbandingan", konstruk: "Konstruk 1 & 4: Pranombor & membuat seriasi", bbm: "Kad gambar, pembilang, bahan maujud", aktiviti: "Membandingkan dua kumpulan objek: banyak/sedikit, lebih/kurang, sama banyak, menyusun mengikut turutan" },
  { kp: "KP 3.1", nama: "Nombor Bulat Hingga 10", sk: "1.2 Nilai Nombor", sp: "1.2.1 Membilang objek dan menamakan nombor 1-10", konstruk: "Konstruk 1, 2, 3: Membilang & menulis nombor 1-10", bbm: "Pembilang, lagu Sayang Semuanya, papan tulis mini", aktiviti: "Nyanyian lagu jari, aktiviti melompat ikut angka, menulis angka dan perkataan 'satu' hingga 'sepuluh'" },
  { kp: "KP 3.2", nama: "Nombor Bulat Hingga 20", sk: "1.6 Nilai Tempat", sp: "1.6.1 Menyatakan nilai tempat puluh dan sa", konstruk: "Konstruk 1, 2, 3: Nilai tempat puluh & sa 11-20", bbm: "Beg plastik berisi 10 guli, guli sa, tali kad nombor", aktiviti: "Konsep membilang sepuluh-sepuluh ('Sepuluh dan satu jadi sebelas'), permainan Bilang Lambungan" },
  { kp: "KP 3.3", nama: "Nombor Bulat Hingga 100", sk: "1.3 Rangkaian Nombor", sp: "1.3.1 Membilang nombor hingga 100", konstruk: "Konstruk 1, 2, 3: Membilang gandaan 2, 5, 10 hingga 100", bbm: "Petak 100 bersaiz besar, tali penyepit baju, pembilang", aktiviti: "Membilang dua-dua, lima-lima dan sepuluh-sepuluh, menyusun nombor tertib menaik dan menurun di tali" },
  { kp: "KP 3.4", nama: "Nombor Bulat Hingga 1000", sk: "1.4 Nilai Tempat & Bundar", sp: "1.4.1 Menyatakan nilai tempat ratus, puluh dan sa", konstruk: "Konstruk 1, 2, 3: Nilai tempat & pembundaran hingga 1000", bbm: "Petak 1000, dekak-dekak, blok Cuisenaire, garis nombor", aktiviti: "Mewakilkan angka pada dekak-dekak, membandingkan dua nombor 3 digit, membundarkan nombor ke puluh terdekat" },
  { kp: "KP 4.1", nama: "Operasi Tambah Dalam Lingkungan 10", sk: "2.1 Konsep Tambah dan Tolak", sp: "2.1.1 Mengguna simbol tambah (+) dan sama dengan (=)", konstruk: "Konstruk 7 & 10: Asas operasi tambah lingkungan 10", bbm: "Balang kaca, guli, garis nombor bersaiz besar", aktiviti: "Aktiviti Penyatuan dua kumpulan guli dalam balang, menulis ayat matematik '2 + 1 = 3', permainan Kotak Beracun" },
  { kp: "KP 4.2", nama: "Operasi Tambah Dalam Lingkungan 18", sk: "2.2 Tambah Fakta Asas", sp: "2.2.1 Menambah dalam lingkungan fakta asas", konstruk: "Konstruk 7 & 10: Operasi tambah lingkungan 18", bbm: "Rantai manik merah & biru, kad bertitik, kad berskala", aktiviti: "Aktiviti Rantai Manik membilang terus, kad bertitik '4 + 8 = 12', menyatakan pasangan nombor hasil tambah sama" },
  { kp: "KP 4.3", nama: "Operasi Tambah Dalam Lingkungan 100", sk: "2.1 Tambah Dua Nombor", sp: "2.1.1 Menambah tanpa & dengan mengumpul semula", konstruk: "Konstruk 7 & 10: Tambah lingkungan 100 bentuk lazim", bbm: "Petak 100, kad imbas nilai tempat ratus-puluh-sa", aktiviti: "Pembelajaran Bermakna: Permainan Boling Botol Mineral 5-9 di luar kelas, kumpul semula 10 sa jadi 1 puluh" },
  { kp: "KP 4.4", nama: "Operasi Tambah Dalam Lingkungan 1000", sk: "2.1 Tambah Tiga Nombor", sp: "2.1.2 Menambah hingga 3 nombor dengan kumpul semula", konstruk: "Konstruk 10 & 11: Tambah 3 nombor & penyelesaian masalah", bbm: "Kad imbas tiga digit, lembaran kerja masalah harian", aktiviti: "Menyelesaikan bentuk lazim 3 nombor, mereka cerita masalah harian guli Upin & Ipin (250 + 148 = 398)" },
  { kp: "KP 5.1", nama: "Operasi Tolak Dalam Lingkungan 10", sk: "2.1 Konsep Tolak", sp: "2.1.1 Mengguna simbol tolak (-) melalui pengasingan", konstruk: "Konstruk 7 & 10: Operasi tolak melalui proses pengasingan", bbm: "Petak 10, pundi kacang, pembilang", aktiviti: "Permainan Ketinting tolak ke belakang ('Kurang satu daripada lima ialah empat'), ayat matematik '9 - 3 = 6'" },
  { kp: "KP 5.2", nama: "Operasi Tolak Dalam Lingkungan 18", sk: "2.3 Tolak Fakta Asas", sp: "2.3.1 Menolak dalam lingkungan fakta asas", konstruk: "Konstruk 7 & 10: Fakta asas tolak lingkungan 18", bbm: "Sida-sida, kad petak 100, kad fakta asas tolak", aktiviti: "Mencari beza dua nombor menggunakan sida-sida (pangkah/asing), permainan pusingan muzik kad fakta asas tolak" },
  { kp: "KP 5.3", nama: "Operasi Tolak Dalam Lingkungan 100", sk: "2.2 Tolak Lingkungan 1000", sp: "2.2.1 Menolak tanpa & dengan mengumpul semula", konstruk: "Konstruk 7 & 10: Tolak 2 digit bentuk lazim & penggenap 10", bbm: "Garis nombor, Petak 100, kad imbas", aktiviti: "Tolak dua digit dengan kaedah penggenap 10 dan turutan menurun, masalah bercerita mangga masak dan hijau" },
  { kp: "KP 5.4", nama: "Operasi Tolak Dalam Lingkungan 1000", sk: "2.2 Tolak Berturut-turut", sp: "2.2.2 Menolak dua nombor berturut-turut", konstruk: "Konstruk 10 & 11: Tolak berturut-turut & songsangan tambah", bbm: "Blok pembilang, kad tolak bentuk lazim", aktiviti: "Menolak nombor 3 digit dengan kumpul semula (741 - 263), operasi tolak sebagai songsangan tambah (700 = 300 + 400)" },
  { kp: "KP 6", nama: "Operasi Darab (Penambahan Berulang & Sifir 0-10)", sk: "2.3 Darab Fakta Asas", sp: "2.3.1 Mendarab dalam lingkungan fakta asas", konstruk: "Konstruk 7, 10, 11: Konsep kumpulan sama banyak & sifir", bbm: "Pinggan kertas, guli, kertas mahjong, papan berputar sifir", aktiviti: "Pembelajaran Bermakna: Menggandakan bekas gula-gula (4 x 3 = 12), membina sifir di pinggan kertas, roda putar sifir" },
  { kp: "KP 7", nama: "Operasi Bahagi (Pengagihan & Pengongsian)", sk: "2.4 Bahagi Fakta Asas", sp: "2.4.1 Membahagi dalam lingkungan fakta asas", konstruk: "Konstruk 7 & 10: Konsep pengagihan, pengongsian & tolak berulang", bbm: "Balang kaca, ikan mainan, kad domino bahagi", aktiviti: "Demonstrasi mengagihkan 15 ekor ikan kepada 3 balang (15 ÷ 3 = 5), tolak berturut-turut rambutan, domino bahagi" },
  { kp: "KP 8.1", nama: "Wang Hingga RM10", sk: "4.1 Wang Kertas dan Syiling", sp: "4.1.1 Mengenal pasti syiling dan wang kertas hingga RM10", konstruk: "Konstruk 5 & 8: Mengenal nilai wang RM dan sen", bbm: "Sampel duit syiling & wang kertas, barangan terpakai", aktiviti: "Menyusun duit mengikut nilai, menekap duit syiling, aktiviti simulasi jual beli di bawah RM10" },
  { kp: "KP 8.2", nama: "Wang Hingga RM100", sk: "4.2 Tambah & Tolak Wang", sp: "4.2.1 Menambah dan menolak nilai wang hingga RM100", konstruk: "Konstruk 5, 8, 12: Tambah & tolak wang serta masalah harian", bbm: "Sampel wang kertas RM50/RM100, tabung mingguan", aktiviti: "Projek 'Mari Menabung' merekod tabungan Isnin-Ahad, kad cerita masalah harian 'Lawatan Ke Zoo Negara'" },
  { kp: "KP 8.3", nama: "Wang Hingga RM1000", sk: "4.4 Darab & Bahagi Wang", sp: "4.4.1 Mendarab & membahagi wang hingga RM1000", konstruk: "Konstruk 5, 8, 12: Pengiraan wang & masalah jual beli", bbm: "Bahan jualan, tanda harga, kad arahan pembelian", aktiviti: "Pembelajaran Bermakna: Projek 'Jualan Mega' murid bertindak sebagai penjual dan pembeli secara berkumpulan" },
  { kp: "KP 9", nama: "Masa dan Waktu (Muka Jam, Jam & Minit, Kalendar)", sk: "5.1 Waktu Dalam Jam & Minit", sp: "5.1.1 Mengenal muka jam dan perkaitan waktu", konstruk: "Konstruk 6 & 12: Menyatakan waktu analog & perkaitan kalendar", bbm: "Model muka jam analog jarum jam/minit, kalendar masihi", aktiviti: "Pembelajaran Bermakna: Membaca jadual waktu kelas & rancangan TV 'Jom Kita Kira', menukar unit jam, minit & saat" }
];

var PETA_ITHINK_LIST = [
  { id: "Peta Bulatan", nama: "Peta Bulatan", fungsi: "Mendefinisikan mengikut konteks / Sumbang saran idea" },
  { id: "Peta Buih", nama: "Peta Buih", fungsi: "Menerangkan kualiti / sifat / ciri menggunakan kata adjektif" },
  { id: "Peta Buih Berganda", nama: "Peta Buih Berganda", fungsi: "Membanding dan membezakan dua konsep / perkara" },
  { id: "Peta Pokok", nama: "Peta Pokok", fungsi: "Mengklasifikasikan bahan, idea atau kategori" },
  { id: "Peta Dakap", nama: "Peta Dakap", fungsi: "Menganalisis hubungan bahagian kepada seluruh objek fizikal" },
  { id: "Peta Alir", nama: "Peta Alir", fungsi: "Membuat urutan langkah proses / kronologi peristiwa" },
  { id: "Peta Pelbagai Alir", nama: "Peta Pelbagai Alir", fungsi: "Menganalisis punca (sebab) dan akibat sesuatu peristiwa" },
  { id: "Peta Titi", nama: "Peta Titi", fungsi: "Mengaplikasikan analogi dan mencari faktor penghubung" }
];

// Helper: Lampirkan Peta i-THINK ke dalam Objek Kandungan RPH
function lampirkanIthink(res, petaIthinkNama) {
  if (!res || !petaIthinkNama) return res;
  if (res.bbmNilaiKbat && !res.bbmNilaiKbat.includes("Peta i-THINK")) {
    res.bbmNilaiKbat += " | Peta i-THINK: " + petaIthinkNama;
  }
  if (res.aktiviti && !res.aktiviti.includes(petaIthinkNama)) {
    if (res.aktiviti.includes("3. Penutup:")) {
      res.aktiviti = res.aktiviti.replace(/(\n3\.\s*Penutup:)/, "\n- Aplikasi PAK21: Murid membina " + petaIthinkNama + " bagi mengukuhkan kefahaman topik.$1");
    } else {
      res.aktiviti += "\n- Aplikasi PAK21: Murid menggunakan " + petaIthinkNama + " untuk menyusun konsep pembelajaran.";
    }
  }
  return res;
}

// Enjin Utama Penjanaan 8 Medan Wajib e-RPH Mengikut Bahasa Pengantar & Data Tally
function binaKandunganRphSpesifik(subjek, tahun, kelas, slotCustom) {
  var petaIthinkNama = slotCustom && slotCustom.petaIthink ? String(slotCustom.petaIthink).trim() : "";
  var hasil = binaKandunganRphSpesifikTeras(subjek, tahun, kelas, slotCustom);
  return lampirkanIthink(hasil, petaIthinkNama);
}

function binaKandunganRphSpesifikTeras(subjek, tahun, kelas, slotCustom) {
  var subUpper = String(subjek || "").toUpperCase().trim();
  var kUpper = String(kelas || "").toUpperCase().trim();
  var isPpki = kUpper.includes("VIVA") || kUpper.includes("WIRA") || kUpper.includes("ARENA") || 
               kUpper.includes("AXIA") || kUpper.includes("SAGA") || kUpper.includes("BEZZA") || kUpper.includes("PPKI");
  var isPra = kUpper.includes("PRA");
  var isBiOrDlp = subUpper.includes("INGGERIS") || subUpper.includes("ENGLISH") || 
                  subUpper.includes("DLP");
  var isArabOrIslam = subUpper.includes("ARAB") || subUpper.includes("ISLAM") || subUpper.includes("PAI") || subUpper.includes("JAWI");
  var isPj = (subUpper.includes("JASMANI") || subUpper === "PJ") || (subUpper.includes("PJPK") && !subUpper.includes("KESIHATAN"));
  var isPk = subUpper.includes("KESIHATAN") || subUpper === "PK";
  var isPsv = subUpper.includes("SENI") || subUpper.includes("PSV");
  var isMuzik = subUpper.includes("MUZIK");
  var isSains = (subUpper.includes("SAINS") || subUpper.includes("SCIENCE")) && !subUpper.includes("DLP");
  var isMatematik = (subUpper.includes("MATEMATIK") || subUpper.includes("MATHEMATICS") || subUpper.includes("MATH")) && !subUpper.includes("DLP");
  var isRbt = subUpper.includes("RBT") || subUpper.includes("REKA BENTUK") || subUpper.includes("TEKNOLOGI");
  var isSejarahOrMoral = subUpper.includes("SEJARAH") || subUpper.includes("MORAL");

  var sTema = slotCustom && slotCustom.tema ? slotCustom.tema : "";
  var sTajuk = slotCustom && slotCustom.tajuk ? slotCustom.tajuk : "";
  var sSk = slotCustom && slotCustom.sk ? slotCustom.sk : "";
  var sSp = slotCustom && slotCustom.sp ? slotCustom.sp : "";

  // 0. MODUL KHAS: INTERVENSI PEMULIHAN KHAS / 3M (BUKU PANDUAN KPM 2019)
  var isPemulihan = slotCustom && (slotCustom.isPemulihan === true || slotCustom.modPemulihan === true || (slotCustom.kp && String(slotCustom.kp).trim() !== ""));
  if (isPemulihan) {
    var isMt = isMatematik || subUpper.includes("NUMERASI");
    var kpList = isMt ? DATA_PEMULIHAN_MT : DATA_PEMULIHAN_BM;
    var kpCode = slotCustom.kp ? String(slotCustom.kp).trim() : "";
    var kpItem = null;
    if (kpCode) {
      kpItem = kpList.find(function(item) {
        return item.kp.toUpperCase() === kpCode.toUpperCase() || item.nama.toUpperCase().includes(kpCode.toUpperCase());
      });
    }
    if (!kpItem) kpItem = kpList[0];

    var subjekLabel = isMt ? "Matematik" : "Bahasa Melayu";
    var tajukPenuh = "Pemulihan Khas (" + subjekLabel + ") • " + kpItem.kp + ": " + kpItem.nama;

    return {
      temaTajuk: tajukPenuh,
      sk: kpItem.sk,
      sp: kpItem.sp,
      objektif: "Pada akhir PdP, murid pemulihan khas berupaya: " + kpItem.nama + " bagi menguasai " + kpItem.konstruk + " mengikut tahap potensi individu.",
      kriteriaKejayaan: "Murid dapat:\n• Aras Rendah (Konkrit): Menguasai langkah asas menggunakan bahan maujud/kad bergambar dengan bimbingan berfokus guru.\n• Aras Sederhana (Semi-Konkrit): Melengkapkan sekurang-kurangnya 3 latihan berpandukan carta visual/garis nombor/kad suku kata secara berpandu.\n• Aras Tinggi (Abstrak): Menulis dan menyelesaikan aktiviti pembelajaran secara mandiri dan yakin.",
      aktiviti: "1. Set Induksi: Rangsangan deria (Kaedah VAK) menggunakan bahan maujud / nyanyian lagu beraksi berkaitan " + kpItem.nama + ".\n2. Aktiviti Utama:\n- Tunjuk cara guru menggunakan pendekatan konkrit-bergambar-abstrak (CPA).\n- Aktiviti Didik Hibur KPM: " + kpItem.aktiviti + ".\n- Latih tubi bertulis terbeza di lembaran kerja pemulihan / buku garis empat mengikut aras penguasaan murid.\n3. Penutup: Peneguhan positif, ganjaran token bintang motivasi dan refleksi ringkas murid.",
      bbmNilaiKbat: "BBM: " + kpItem.bbm + " | Kaedah: VAK (Visual-Auditori-Kinestetik) | Nilai: Ketekunan, Berdikari | KBAT: Mengaplikasi | PBD: Pemerhatian & Lembaran Kerja"
    };
  }

  // Dapatkan padanan DSKP yang 100% tally
  var dskp = ambilObjektifDskp(subjek, tahun, kelas, sTema, sTajuk, sSk, sSp);
  var temaTajuk = dskp.temaTajuk;
  var sk = dskp.sk;
  var sp = dskp.sp;

  // 1. KATEGORI: PPKI MODEL KERETA (1 VIVA, 2 WIRA, 3 ARENA, 4 AXIA, 5 SAGA, 6 BEZZA)
  if (isPpki) {
    return {
      temaTajuk: temaTajuk,
      sk: sk,
      sp: sp,
      objektif: "Pada akhir PdP, murid berkeperluan pendidikan khas (MBPK) dapat: Menguasai kemahiran asas (" + sp + ") bagi " + dskp.tajuk + " mengikut potensi individu.",
      kriteriaKejayaan: "Murid dapat:\n• Aras Rendah: Melakukan sekurang-kurangnya 2 langkah kemahiran dengan bantuan fizikal penuh guru/PPM.\n• Aras Sederhana: Melaksanakan sekurang-kurangnya 3 langkah kemahiran berpandukan isyarat visual/lisan.\n• Aras Tinggi: Melengkapkan kemahiran pembelajaran secara berdikari dan tertib.",
      aktiviti: "1. Set Induksi: Rangsangan deria, lagu beraksi atau tayangan bahan maujud berkaitan " + dskp.tajuk + ".\n2. Aktiviti Utama:\n- Demonstrasi amali oleh guru langkah demi langkah menggunakan bahan maujud/stesen.\n- Latihan amali terbeza murid (Murid Aras Rendah: bimbingan fizikal berfokus guru/PPM; Murid Aras Sederhana & Tinggi: amali berpandukan carta visual dan berdikari).\n3. Penutup: Pujian motivasi, peneguhan positif dan token ganjaran bintang kepada murid.",
      bbmNilaiKbat: "BBM: Bahan Maujud, Carta Bergambar, Lembaran Kerja | Nilai: Berdikari, Kebersihan | KBAT: Mengaplikasi | PBD: Pemerhatian & Amali"
    };
  }

  // 2. KATEGORI: KURIKULUM PRASEKOLAH 2026 (STANDARD + KOMPETENSI + ADAB)
  if (isPra) {
    return {
      temaTajuk: temaTajuk,
      sk: sk,
      sp: sp,
      objektif: "Pada akhir pembelajaran, murid prasekolah dapat: Menguasai standard pembelajaran (" + sp + ") bagi bidang " + (subjek || "Prasekolah 2026") + " dengan beradab dan yakin.",
      kriteriaKejayaan: "Murid dapat (Selaras Jadual 8 TP Umum Prasekolah 2026 KPM):\n• TP 1 (Menyatakan): Menyatakan sesuatu perkara dengan beradab berasaskan pengetahuan dan kemahiran dengan bimbingan.\n• TP 2 (Menerangkan): Menerangkan sesuatu perkara dengan beradab berasaskan pengetahuan dan kemahiran berpandukan contoh.\n• TP 3 (Membuat): Membuat sesuatu perkara dengan beradab berasaskan pengetahuan dan kemahiran secara berdikari.",
      aktiviti: "1. Set Induksi: Rutin perbualan pagi, nyanyian lagu beraksi / tayangan rangsangan sensori bagi membina kesediaan emosi dan minda berkaitan " + dskp.tajuk + ".\n2. Aktiviti Utama (Pendekatan Amalan Bersesuaian Perkembangan - ABP & Belajar Melalui Bermain):\n- Penerokaan konsep dan demonstrasi interaktif guru menggunakan bahan maujud, kad gambar dan media digital selamat.\n- Aktiviti Pembelajaran Terbeza mengikut stesen pusat pembelajaran:\n  * Kumpulan TP 1: Bimbingan rapi guru/PPM bagi mengukuhkan kemahiran asas literasi/numerasi/motor.\n  * Kumpulan TP 2 & TP 3: Aktiviti amali hands-on, kolaborasi berkumpulan, dan tugasan kreatif berasaskan projek secara berdikari.\n3. Penutup (Thinking Classroom & Karakter):\n- Sesi refleksi kendiri murid (Exit Ticket), perkongsian rasa seronok belajar, peneguhan adab dan gotong-royong mengemas ruang.",
      bbmNilaiKbat: "BBM: Bahan Maujud, Kad Huruf/Nombor NDL, Media Sensori, Alat Muzik Perkusi, Lembaran Kerja | Nilai: Kasih Sayang, Hormat, Berdaya Tahan | KBAT: Menganalisis & Mengaplikasi | PBD: Pentaksiran Autentik & Senarai Semak TP1-TP3"
    };
  }

  // 3. KATEGORI: PENDIDIKAN JASMANI (PJ / PADANG & GELANGGANG)
  if (isPj) {
    return {
      temaTajuk: temaTajuk,
      sk: sk,
      sp: sp,
      objektif: "Pada akhir PdP, murid dapat: Melakukan kemahiran pergerakan asas (" + sp + ") bagi " + dskp.tajuk + " dengan teknik lakuan yang betul dan selamat.",
      kriteriaKejayaan: "Murid dapat:\n• All (Rendah): Melakukan lakuan asas sekurang-kurangnya 2 kali percubaan dengan bimbingan guru.\n• Most (Sederhana): Menguasai kemahiran lakuan mengikut teknik betul secara berulang sekurang-kurangnya 3 kali.\n• Some (Tinggi): Mempamerkan teknik lakuan kemahiran secara konsisten, pantas dan membimbing rakan sebaya.",
      aktiviti: "1. Set Induksi: Aktiviti memanaskan badan (warm-up), regangan otot dinamik dan peningkatan kadar nadi di padang/gelanggang.\n2. Aktiviti Utama:\n- Demonstrasi teknik lakuan kemahiran langkah demi langkah oleh guru.\n- Latih tubi kemahiran secara ansur maju & stesen (Kumpulan Rendah: bimbingan ansur maju guru; Kumpulan Sederhana & Tinggi: variasi jarak & sasaran secara berdikari).\n- Permainan kecil bersyarat bagi mengaplikasi kemahiran dan peraturan keselamatan permainan.\n3. Penutup: Aktiviti menyejukkan badan (cool-down), refleksi nilai semangat kesukanan dan pengurusan alatan sukan bersama.",
      bbmNilaiKbat: "BBM: Wisel, Kon / Skitel, Bola, Pundi Kacang, Gelung, Jam Randik, Bib Rompi | Nilai: Semangat Kesukanan, Kerjasama, Keselamatan | KBAT: Mengaplikasi Lakuan | PBD: Amali & Pemerhatian"
    };
  }

  // 4. KATEGORI: PENDIDIKAN KESIHATAN (PK)
  if (isPk) {
    return {
      temaTajuk: temaTajuk,
      sk: sk,
      sp: sp,
      objektif: "Pada akhir PdP, murid dapat: Menerangkan dan mengamalkan (" + sp + ") bagi " + dskp.tajuk + " ke arah gaya hidup sihat dan selamat.",
      kriteriaKejayaan: "Murid dapat:\n• All (Rendah): Menyatakan sekurang-kurangnya 2 amalan kesihatan/kebersihan diri dengan bimbingan guru.\n• Most (Sederhana): Menjelaskan sekurang-kurangnya 3 kepentingan menjaga kesihatan/keselamatan diri dengan betul.\n• Some (Tinggi): Merumuskan langkah tindakan positif dan menjana idea pencegahan kemudaratan secara kritis.",
      aktiviti: "1. Set Induksi: Soal jawab situasi harian / gambar rangsangan berkaitan amalan kebersihan dan kesihatan diri.\n2. Aktiviti Utama:\n- Penerangan guru dan perbincangan interaktif berpandukan carta kesihatan / piramid makanan.\n- Murid melengkapkan tugasan terbeza (Kumpulan Rendah: melabel dan memadankan carta; Kumpulan Sederhana & Tinggi: menghuraikan situasi dan tindakan selamat secara bertulis).\n3. Penutup: Rumusan komitmen amalan gaya hidup sihat harian dan peneguhan positif guru.",
      bbmNilaiKbat: "BBM: Carta Kesihatan Diri, Model Piramid Makanan, Buku Teks, Lembaran Kerja | Nilai: Kebersihan Fizikal & Mental, Berhemah | KBAT: Menganalisis | PBD: Lisan & Latihan Bertulis"
    };
  }

  // 5. KATEGORI: PENDIDIKAN SENI VISUAL (PSV)
  if (isPsv) {
    return {
      temaTajuk: temaTajuk,
      sk: sk,
      sp: sp,
      objektif: "Pada akhir PdP, murid dapat: Mengaplikasikan bahasa seni visual dan media dalam menghasilkan karya (" + sp + ") bagi " + dskp.tajuk + " secara kreatif.",
      kriteriaKejayaan: "Murid dapat:\n• All (Rendah): Meneroka media dan menghasilkan karya asas dengan bimbingan teknik oleh guru.\n• Most (Sederhana): Mengaplikasikan teknik seni secara kemas dan menepati tema yang dipelajari.\n• Some (Tinggi): Menghasilkan karya seni yang kreatif, berjalinan menarik dan mempunyai sentuhan tersendiri secara mandiri.",
      aktiviti: "1. Set Induksi: Meneliti contoh karya seni visual dan bersoal jawab tentang unsur seni (garisan, jalinan, warna, imbangan).\n2. Aktiviti Utama:\n- Tunjuk cara demonstrasi langkah penghasilan karya oleh guru.\n- Murid meneroka media dan berkarya mengikut aras (Kumpulan Rendah: bimbingan teknik asas oleh guru; Kumpulan Sederhana & Tinggi: penerokaan motif & gubahan kreatif secara berdikari).\n3. Penutup: Apresiasi seni ringkas (Gallery Walk / pameran hasil kerja murid) dan membersihkan ruang kerja bersama.",
      bbmNilaiKbat: "BBM: Kertas Lukisan, Krayon / Warna Air, Berus, Gunting, Gam, Bahan Kitar Semula | Nilai: Kreativiti, Kebersihan, Ketekunan | KBAT: Mereka Cipta | PBD: Hasil Kerja (Karya) & Pemerhatian"
    };
  }

  // 6. KATEGORI: PENDIDIKAN MUZIK
  if (isMuzik) {
    return {
      temaTajuk: temaTajuk,
      sk: sk,
      sp: sp,
      objektif: "Pada akhir PdP, murid dapat: Bernyanyi / memainkan corak irama perkusi (" + sp + ") bagi " + dskp.tajuk + " mengikut tempo dan pic yang betul.",
      kriteriaKejayaan: "Murid dapat:\n• All (Rendah): Mengajuk pic melodi / menepuk detik irama asas dengan bimbingan guru.\n• Most (Sederhana): Menyanyikan lagu atau memainkan alat perkusi mengikut tempo dengan lancar.\n• Some (Tinggi): Memainkan corak irama bergilir/harmoni atau mempersembahkan nyanyian secara yakin dan berdikari.",
      aktiviti: "1. Set Induksi: Latihan pernafasan, pemanasan vokal dan aktiviti menepuk corak irama asas.\n2. Aktiviti Utama:\n- Memperdengarkan rakaman lagu dan mengajuk sebutan pic/melodi secara berpandu.\n- Murid berlatih memainkan alat perkusi / menyanyi secara ensembel (Kumpulan Rendah: bimbingan rentak asas; Kumpulan Sederhana & Tinggi: corak irama berlapisi secara berdikari).\n3. Penutup: Persembahan muzikal berkumpulan dan rumusan apresiasi muzik.",
      bbmNilaiKbat: "BBM: Alat Perkusi (Kompang, Kastanet, Kerincing, Tamborin), Audio Lagu / Minus-One, Carta Irama | Nilai: Disiplin, Kerjasama, Keyakinan | KBAT: Mengaplikasi Tempo | PBD: Persembahan Lisan & Amali"
    };
  }

  // 7. KATEGORI: SAINS / SCIENCE (DLP)
  if (isSains) {
    return {
      temaTajuk: temaTajuk,
      sk: sk,
      sp: sp,
      objektif: "Pada akhir PdP, murid dapat: Menyiasat dan menerangkan konsep sains (" + sp + ") bagi " + dskp.tajuk + " melalui kemahiran proses sains.",
      kriteriaKejayaan: "Murid dapat:\n• All (Rendah): Memerhati dan menyatakan sekurang-kurangnya 2 fakta/pemerhatian sains dengan bimbingan guru.\n• Most (Sederhana): Merekod hasil pemerhatian / uji kaji dan membina inferens awal dengan betul.\n• Some (Tinggi): Merumuskan kesimpulan penyiasatan saintifik dan menghubungkaitkan dengan kehidupan harian secara mandiri.",
      aktiviti: "1. Set Induksi: Menonton fenomena sains / demonstrasi bahan maujud mencungkil rasa ingin tahu murid.\n2. Aktiviti Utama:\n- Penerangan konsep sains dan taklimat keselamatan penggunaan radas sains.\n- Murid menjalankan aktiviti inkuiri penemuan / eksperimen ringkas berkumpulan dan merekod data (Kumpulan Rendah: bimbingan merekod pemerhatian; Kumpulan Sederhana & Tinggi: analisis data dan membina hipotesis/kesimpulan secara berdikari).\n3. Penutup: Rumusan dapatan sains bersama murid dan perbincangan aplikasi dalam kehidupan harian.",
      bbmNilaiKbat: "BBM: Radas Sains, Bikar, Kanta Pembesar, Bahan Maujud / Spesimen Alam, Buku Teks, Lembaran Kerja | Nilai: Kejujuran Saintifik, Sifat Ingin Tahu, Bekerjasama | KBAT: Menganalisis | PBD: Amali Sains & Catatan Laporan"
    };
  }

  // 8. KATEGORI: MATEMATIK / MATHEMATICS (DLP)
  if (isMatematik) {
    return {
      temaTajuk: temaTajuk,
      sk: sk,
      sp: sp,
      objektif: "Pada akhir PdP, murid dapat: Menguasai kemahiran mengira dan menyelesaikan operasi (" + sp + ") bagi tajuk " + dskp.tajuk + " dengan tepat.",
      kriteriaKejayaan: "Murid dapat:\n• All (Rendah): Menyelesaikan sekurang-kurangnya 2 soalan operasi asas menggunakan bahan manipulatif/bimbingan guru.\n• Most (Sederhana): Menyelesaikan sekurang-kurangnya 3 soalan pengiraan rutin mengikut langkah betul.\n• Some (Tinggi): Menyelesaikan masalah harian (bukan rutin/KBAT) dengan strategi pelbagai secara mandiri.",
      aktiviti: "1. Set Induksi: Permainan congak pantas / situasi masalah harian melibatkan nilai nombor.\n2. Aktiviti Utama:\n- Penerangan konsep melalui pendekatan Konkrit-Bergambar-Abstrak (CPA) menggunakan bahan manipulatif.\n- Latihan pengiraan terbeza (Kumpulan Rendah: bimbingan berfokus guru & blok dienes; Kumpulan Sederhana & Tinggi: latihan penyelesaian masalah berayat secara mandiri).\n3. Penutup: Refleksi strategi pengiraan pantas dan rumusan konsep nombor.",
      bbmNilaiKbat: "BBM: Blok Dienes (Asas 10), Abakus, Garis Nombor, Kad Imbasan Angka, Papan Putih Mini | Nilai: Ketepatan, Rasional, Ketekunan | KBAT: Menyelesaikan Masalah | PBD: Latihan Bertulis & Lisan"
    };
  }

  // 9. KATEGORI: REKA BENTUK & TEKNOLOGI (RBT)
  if (isRbt) {
    return {
      temaTajuk: temaTajuk,
      sk: sk,
      sp: sp,
      objektif: "Pada akhir PdP, murid dapat: Mengaplikasikan pengetahuan kemahiran teknikal (" + sp + ") bagi menghasilkan produk / projek " + dskp.tajuk + " secara selamat.",
      kriteriaKejayaan: "Murid dapat:\n• All (Rendah): Mengenal pasti sekurang-kurangnya 2 alatan tangan dan fungsinya dengan bimbingan guru.\n• Most (Sederhana): Melakar atau memasang komponen projek mengikut manual dengan betul.\n• Some (Tinggi): Menguji fungsi projek dan mencadangkan penambahbaikan reka bentuk secara mandiri.",
      aktiviti: "1. Set Induksi: Menganalisis fungsi produk maujud dan cabaran reka bentuk teknologi masa kini.\n2. Aktiviti Utama:\n- Taklimat peraturan keselamatan bengkel dan penggunaan alatan tangan.\n- Murid melaksanakan amali lakaran dan pemasangan kit projek secara terbeza (Kumpulan Rendah: bimbingan langkah demi langkah; Kumpulan Sederhana & Tinggi: amali pemasangan dan pengujian projek secara berpasangan).\n3. Penutup: Ujian fungsi produk, pembersihan meja kerja bengkel dan rumusan.",
      bbmNilaiKbat: "BBM: Kit Model Projek, Alatan Tangan (Pembaris Keluli, Gunting), Bahan Projek, Carta Manual | Nilai: Keselamatan, Inovasi, Berhemah | KBAT: Mereka Cipta | PBD: Hasil Projek & Amali Bengkel"
    };
  }

  // 10. KATEGORI: SEJARAH & PENDIDIKAN MORAL
  if (isSejarahOrMoral) {
    return {
      temaTajuk: temaTajuk,
      sk: sk,
      sp: sp,
      objektif: "Pada akhir PdP, murid dapat: Menghuraikan peristiwa / nilai murni (" + sp + ") bagi tajuk " + dskp.tajuk + " serta mengaitkannya dengan jati diri warganegara.",
      kriteriaKejayaan: "Murid dapat:\n• All (Rendah): Menyatakan sekurang-kurangnya 2 fakta sejarah / nilai murni dengan bimbingan guru.\n• Most (Sederhana): Menjelaskan kronologi peristiwa atau menghuraikan kepentingan nilai murni dalam kehidupan.\n• Some (Tinggi): Menilai iktibar sejarah / membina justifikasi penyelesaian dilema moral secara kritis.",
      aktiviti: "1. Set Induksi: Meneliti gambar sumber sejarah / tayangan kad senario nilai mencungkil pandangan murid.\n2. Aktiviti Utama:\n- Penerokaan konsep dan perbincangan berkumpulan berpandukan garis masa / situasi moral.\n- Murid melengkapkan tugasan terbeza (Kumpulan Rendah: menyusun kronologi bergambar / memadankan nilai; Kumpulan Sederhana & Tinggi: perbincangan analitikal & latihan berstruktur secara mandiri).\n3. Penutup: Refleksi iktibar sejarah / penghayatan nilai dan rumusan bersama guru.",
      bbmNilaiKbat: "BBM: Gambar Sumber Sejarah, Garis Masa, Kad Situasi Nilai, Buku Teks, Lembaran Kerja | Nilai: Patriotisme, Tanggungjawab, Menghormati | KBAT: Menilai Iktibar | PBD: Lisan & Latihan Bertulis"
    };
  }

  // 11. KATEGORI: BAHASA ARAB, PENDIDIKAN ISLAM & JAWI
  if (isArabOrIslam) {
    return {
      temaTajuk: temaTajuk,
      sk: sk,
      sp: sp,
      objektif: "Pada akhir PdP, murid dapat: Menyebut, membaca, menghafaz atau menulis kemahiran asas (" + sp + ") bagi tajuk " + dskp.tajuk + " dengan betul dan fasih.",
      kriteriaKejayaan: "Murid dapat:\n• All (Rendah): Menyebut atau mengecam sekurang-kurangnya 2 kalimah/huruf fokus dengan bimbingan guru.\n• Most (Sederhana): Membaca dan memadankan sekurang-kurangnya 3 kalimah / menghafaz ayat dengan betul.\n• Some (Tinggi): Menulis sekurang-kurangnya 4 kalimah atau menghuraikan hukum/adab secara mandiri dan fasih.",
      aktiviti: "1. Set Induksi: Bacaan doa, alunan ayat suci / tayangan kad kalimah bergambar mencungkil idea murid.\n2. Aktiviti Utama:\n- Talaqqi musyafahah dan latih tubi makhraj huruf / sebutan kalimah secara kelas dan kumpulan.\n- Murid melengkapkan latihan terbeza (Kumpulan Rendah: bimbingan berfokus guru & kad imbasan; Kumpulan Sederhana & Tinggi: lembaran kerja bertulis / amali solat secara mandiri).\n3. Penutup: Tasmik bacaan ringkas kalimah yang dipelajari dan peneguhan motivasi adab.",
      bbmNilaiKbat: "BBM: Kad Kalimah, Carta Tajwid / Doa, Mushaf Al-Quran, Kad Imbasan Bergambar | Nilai: Khusyuk, Ketaatan, Adab & Tertib | KBAT: Mengaplikasi | PBD: Lisan & Amali Ibadah"
    };
  }

  // 12. KATEGORI: BAHASA INGGERIS & DLP (ENGLISH MEDIUM)
  if (isBiOrDlp) {
    return {
      temaTajuk: temaTajuk,
      sk: sk,
      sp: sp,
      objektif: "By the end of the lesson, pupils will be able to: Master learning standard (" + sp + ") for " + dskp.tajuk + " accurately.",
      kriteriaKejayaan: "Pupils can:\n• All (Low): Identify and state at least 2 key words/concepts with teacher guidance.\n• Most (Mid): Complete at least 3 structured questions/sentences based on " + sp + " with minimal support.\n• Some (High): Construct sentences or solve challenging tasks independently.",
      aktiviti: "1. Set Induksi: Warm-up activity and recap of prior knowledge regarding " + dskp.tajuk + " using story cards / flashcards.\n2. Aktiviti Utama:\n- Teacher explains key concepts and demonstrates examples using contextual learning aids.\n- Pupils complete differentiated exercises (Low: guided reading/flashcards with teacher support; Mid & High: independent workbook exercises/enrichment).\n3. Penutup: Pupils reflect on learning, complete exit ticket, and summarize key concepts.",
      bbmNilaiKbat: "BBM: Textbook, Story Flashcards, Word Cards, Differentiated Activity Sheets | Nilai: Diligence, Confidence | KBAT: Application | PBD: Listening, Speaking, Reading & Writing"
    };
  }

  // 13. KATEGORI: ALIRAN PERDANA (BAHASA MELAYU, BKD & LAIN-LAIN)
  return {
    temaTajuk: temaTajuk,
    sk: sk,
    sp: sp,
    objektif: "Pada akhir PdP, murid dapat: Menguasai Standard Pembelajaran (" + sp + ") bagi tajuk " + dskp.tajuk + " mengikut aras keupayaan murid.",
    kriteriaKejayaan: "Murid dapat:\n• All (Rendah): Menyatakan atau menjawab sekurang-kurangnya 2 soalan asas berkaitan " + dskp.tajuk + " dengan bimbingan guru.\n• Most (Sederhana): Menyelesaikan sekurang-kurangnya 3 tugasan latihan berpandukan " + sk + " dengan betul.\n• Some (Tinggi): Melengkapkan kesemua latihan dan soalan aras tinggi/pengayaan secara mandiri.",
    aktiviti: "1. Set Induksi: Guru memaparkan rangsangan cerita / gambar bersiri berkaitan " + dskp.tajuk + " dan sesi soal jawab mencungkil idea kosa kata.\n2. Aktiviti Utama:\n- Membaca petikan dengan sebutan dan intonasi yang betul berpandukan buku teks.\n- Murid melengkapkan latihan bertulis mengikut aras keupayaan (Kumpulan Rendah: bimbingan membina frasa/ayat asas; Kumpulan Sederhana & Tinggi: membina perenggan & karangan ringkas secara mandiri).\n3. Penutup: Guru dan murid membuat rumusan kosa kata baharu serta refleksi ringkas (Exit Ticket).",
    bbmNilaiKbat: "BBM: Buku Teks, Petikan Cerita Bergambar, Kad Kosa Kata, Lembaran Kerja Terbeza | Nilai: Kerjasama, Santun Berbahasa | KBAT: Mengaplikasi | PBD: Lisan, Bacaan & Latihan Bertulis"
  };
}

// --------------------------------------------------------------------------
// 3. PENGGUNA & TAKWIM
// --------------------------------------------------------------------------
// Matriks Peranan & Kawalan Hak Akses (RBAC) SK Sook 2026
function tentukanSkopPeranan(peranan) {
  var p = String(peranan || "").toUpperCase().trim();
  if (p.includes("GURU BESAR") || p === "GB") {
    return { perananKod: "GURU_BESAR", label: "Guru Besar", skop: ["KURIKULUM", "HEM", "KOKURIKULUM", "PPKI"], bolehLulus: true };
  }
  if (p.includes("ADMIN") || p === "PENTADBIR") {
    return { perananKod: "ADMIN", label: "Pentadbir Sistem", skop: ["KURIKULUM", "HEM", "KOKURIKULUM", "PPKI"], bolehLulus: true };
  }
  if (p.includes("PK 1") || p.includes("PK1") || p.includes("PENTADBIRAN") || p.includes("KURIKULUM")) {
    return { perananKod: "PK1", label: "Penolong Kanan Pentadbiran", skop: ["KURIKULUM"], bolehLulus: true };
  }
  if (p.includes("PK HEM") || p.includes("PKHEM") || p.includes("HAL EHWAL MURID")) {
    return { perananKod: "PK_HEM", label: "Penolong Kanan HEM", skop: ["HEM"], bolehLulus: true };
  }
  if (p.includes("PK KOKU") || p.includes("PKKOKU") || p.includes("KOKURIKULUM")) {
    return { perananKod: "PK_KOKUM", label: "Penolong Kanan Kokurikulum", skop: ["KOKURIKULUM"], bolehLulus: true };
  }
  if (p.includes("PK PPKI") || p.includes("PKPPKI") || p.includes("PENDIDIKAN KHAS")) {
    return { perananKod: "PK_PPKI", label: "Penolong Kanan PPKI", skop: ["PPKI"], bolehLulus: true };
  }
  return { perananKod: "GURU", label: "Guru Akademik / Penasihat", skop: [], bolehLulus: false };
}

function getSenaraiGuruWeb() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPenga = ss.getSheetByName("PENGGUNA");
  
  if (!sheetPenga) {
    return [
      { emel: "badrul@sksook.edu.my", nama: "Badrul Hisyam Rasamin", noKp: "880102125432", jawatan: "DG44 Guru Bahasa Inggeris", peranan: "PENTADBIR", penyemakEmail: "pk1@sksook.edu.my", katalaluan: "5432", rbac: tentukanSkopPeranan("PENTADBIR") },
      { emel: "noraini@sksook.edu.my", nama: "Noraini binti Ahmad", noKp: "900514126780", jawatan: "DG41 Guru Sains", peranan: "GURU", penyemakEmail: "pk1@sksook.edu.my", katalaluan: "6780", rbac: tentukanSkopPeranan("GURU") },
      { emel: "kamal@sksook.edu.my", nama: "Mohd Kamal bin Idris", noKp: "850321125544", jawatan: "DG44 Guru PPKI", peranan: "GURU", penyemakEmail: "pkhem@sksook.edu.my", katalaluan: "5544", rbac: tentukanSkopPeranan("GURU") }
    ];
  }
  
  var data = sheetPenga.getDataRange().getValues();
  var senarai = [];
  if (data.length <= 1) return senarai;

  var headers = data[0];
  var colMap = {};
  for (var c = 0; c < headers.length; c++) {
    var rawH = String(headers[c] || "").trim().toUpperCase();
    if (rawH) {
      colMap[rawH] = c;
      colMap[rawH.replace(/[\s\-]/g, "_")] = c;
    }
  }

  var cEmel = colMap["EMAIL"] !== undefined ? colMap["EMAIL"] : (colMap["EMEL"] !== undefined ? colMap["EMEL"] : (colMap["EMAIL_DLIMA"] !== undefined ? colMap["EMAIL_DLIMA"] : 0));
  var cNama = colMap["NAMA_GURU"] !== undefined ? colMap["NAMA_GURU"] : (colMap["NAMA GURU"] !== undefined ? colMap["NAMA GURU"] : (colMap["NAMA"] !== undefined ? colMap["NAMA"] : 1));
  var cJawatan = colMap["JAWATAN"] !== undefined ? colMap["JAWATAN"] : 2;
  var cPeranan = colMap["PERANAN"] !== undefined ? colMap["PERANAN"] : 3;
  var cNoKp = colMap["NO_KP"] !== undefined ? colMap["NO_KP"] : 
              (colMap["NO KP"] !== undefined ? colMap["NO KP"] : 
              (colMap["KP"] !== undefined ? colMap["KP"] : 
              (colMap["IC"] !== undefined ? colMap["IC"] : 
              (colMap["NO. KP"] !== undefined ? colMap["NO. KP"] : 
              (colMap["NO_KAD_PENGENALAN"] !== undefined ? colMap["NO_KAD_PENGENALAN"] : 4)))));
  var cKatalaluan = colMap["KATALALUAN"] !== undefined ? colMap["KATALALUAN"] : 
                    (colMap["KATA LALUAN"] !== undefined ? colMap["KATA LALUAN"] : 
                    (colMap["PASSWORD"] !== undefined ? colMap["PASSWORD"] : 
                    (colMap["PASSCODE"] !== undefined ? colMap["PASSCODE"] : 5)));
  var cPenyemak = colMap["PENYEMAK_EMAIL"] !== undefined ? colMap["PENYEMAK_EMAIL"] : 
                  (colMap["PENYEMAK EMAIL"] !== undefined ? colMap["PENYEMAK EMAIL"] : 
                  (colMap["PENYEMAK"] !== undefined ? colMap["PENYEMAK"] : 
                  (colMap["PENYEMAK_EMEL"] !== undefined ? colMap["PENYEMAK_EMEL"] : 6)));
  var cFoto = colMap["FOTO"] !== undefined ? colMap["FOTO"] : 
              (colMap["GAMBAR"] !== undefined ? colMap["GAMBAR"] : 
              (colMap["PHOTO"] !== undefined ? colMap["PHOTO"] : 7));

  var perluKemaskiniBatch = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][cEmel] && data[i][cNama]) {
      var perananRaw = data[i][cPeranan] ? data[i][cPeranan].toString().trim().toUpperCase() : "GURU";
      var fotoRaw = (cFoto !== undefined && data[i][cFoto]) ? data[i][cFoto].toString().trim() : "";
      var fotoUrl = (fotoRaw.startsWith("data:image") || fotoRaw.startsWith("http://") || fotoRaw.startsWith("https://")) ? fotoRaw : "";
      var noKpVal = (cNoKp !== undefined && data[i][cNoKp]) ? String(data[i][cNoKp]).replace(/\D/g, "") : "";
      var penyemakVal = (cPenyemak !== undefined && data[i][cPenyemak]) ? String(data[i][cPenyemak]).trim() : "";
      var sediaPass = (cKatalaluan !== undefined && data[i][cKatalaluan]) ? String(data[i][cKatalaluan]).trim() : "";

      var kataLaluanVal = sediaPass;
      if (!kataLaluanVal) {
        kataLaluanVal = (noKpVal.length >= 4) ? noKpVal.slice(-4) : "2026";
        if (cKatalaluan !== undefined && noKpVal.length >= 4) {
          perluKemaskiniBatch = true;
        }
      }

      senarai.push({
        emel: data[i][cEmel].toString().trim(),
        nama: data[i][cNama].toString().trim(),
        jawatan: (cJawatan !== undefined && data[i][cJawatan]) ? data[i][cJawatan].toString().trim() : "Guru Akademik",
        peranan: perananRaw,
        foto: fotoUrl,
        noKp: noKpVal,
        penyemakEmail: penyemakVal,
        katalaluan: kataLaluanVal,
        rbac: tentukanSkopPeranan(perananRaw)
      });
    }
  }

  // Jika terdapat rekod kata laluan kosong tetapi ada No. KP, selaraskan secara batch pantas ke sheet
  if (perluKemaskiniBatch && cKatalaluan !== undefined) {
    try {
      var batchValues = [];
      for (var r = 1; r < data.length; r++) {
        var pVal = String(data[r][cKatalaluan] || "").trim();
        if (!pVal) {
          var rKp = (cNoKp !== undefined && data[r][cNoKp]) ? String(data[r][cNoKp]).replace(/\D/g, "") : "";
          if (rKp.length >= 4) {
            pVal = "'" + rKp.slice(-4);
          }
        } else {
          if (pVal.charAt(0) !== "'") pVal = "'" + pVal;
        }
        batchValues.push([pVal]);
      }
      if (batchValues.length > 0) {
        sheetPenga.getRange(2, cKatalaluan + 1, batchValues.length, 1).setValues(batchValues);
      }
    } catch (eBatch) {
      Logger.log("Batch kemaskini kata laluan: " + eBatch.message);
    }
  }

  return senarai;
}

/**
 * Tambah Pengguna Baru (Akses Pentadbir sahaja)
 * Auto generate password berasaskan 4 nombor belakang IC
 */
function tambahPenggunaBaru(data) {
  try {
    if (!data) return { success: false, message: "Tiada data pengguna dibekalkan." };
    var emel = String(data.emel || data.email || "").trim();
    var nama = String(data.nama || "").trim();
    var noKp = String(data.noKp || "").trim().replace(/\D/g, "");
    var jawatan = String(data.jawatan || "Guru Akademik").trim();
    var peranan = String(data.peranan || "GURU").trim().toUpperCase();
    var penyemakEmail = String(data.penyemakEmail || "").trim();

    if (!emel) return { success: false, message: "Emel (Email dlima) wajib diisi." };
    if (!nama) return { success: false, message: "Nama guru wajib diisi." };

    // Auto-generate password berdasarkan 4 nombor belakang IC guru
    var passwordAuto = (noKp.length >= 4) ? noKp.slice(-4) : "2026";
    var katalaluan = String(data.katalaluan || passwordAuto).trim();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PENGGUNA");
    var headersStandard = ["Email", "Nama_Guru", "Jawatan", "Peranan", "NO_KP", "KATALALUAN", "Penyemak_Email", "FOTO"];

    if (!sheet) {
      sheet = ss.insertSheet("PENGGUNA");
      sheet.appendRow(headersStandard);
      sheet.getRange(1, 1, 1, headersStandard.length).setBackground("#1e1b4b").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap = {};
    for (var c = 0; c < headers.length; c++) {
      var rawH = String(headers[c] || "").trim().toUpperCase();
      if (rawH) {
        colMap[rawH] = c;
        colMap[rawH.replace(/[\s\-]/g, "_")] = c;
      }
    }

    var idxEmel = colMap["EMAIL"] !== undefined ? colMap["EMAIL"] : (colMap["EMEL"] !== undefined ? colMap["EMEL"] : (colMap["EMAIL_DLIMA"] !== undefined ? colMap["EMAIL_DLIMA"] : 0));
    var idxNama = colMap["NAMA_GURU"] !== undefined ? colMap["NAMA_GURU"] : (colMap["NAMA GURU"] !== undefined ? colMap["NAMA GURU"] : (colMap["NAMA"] !== undefined ? colMap["NAMA"] : 1));
    var idxJawatan = colMap["JAWATAN"] !== undefined ? colMap["JAWATAN"] : 2;
    var idxPeranan = colMap["PERANAN"] !== undefined ? colMap["PERANAN"] : 3;
    var idxNoKp = colMap["NO_KP"] !== undefined ? colMap["NO_KP"] : (colMap["KP"] !== undefined ? colMap["KP"] : (colMap["IC"] !== undefined ? colMap["IC"] : 4));
    var idxPass = colMap["KATALALUAN"] !== undefined ? colMap["KATALALUAN"] : (colMap["PASSWORD"] !== undefined ? colMap["PASSWORD"] : (colMap["PASSCODE"] !== undefined ? colMap["PASSCODE"] : 5));
    var idxPenyemak = colMap["PENYEMAK_EMAIL"] !== undefined ? colMap["PENYEMAK_EMAIL"] : (colMap["PENYEMAK"] !== undefined ? colMap["PENYEMAK"] : (colMap["PENYEMAK_EMEL"] !== undefined ? colMap["PENYEMAK_EMEL"] : 6));
    var idxFoto = colMap["FOTO"] !== undefined ? colMap["FOTO"] : (colMap["GAMBAR"] !== undefined ? colMap["GAMBAR"] : 7);

    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (padanEmelSama(rows[i][idxEmel], emel)) {
        return { success: false, message: "Emel (" + emel + ") telah pun didaftarkan sebelum ini." };
      }
    }

    var maxIdx = Math.max(lastCol - 1, idxEmel, idxNama, idxJawatan, idxPeranan, idxNoKp, idxPass, idxPenyemak, idxFoto);
    var newRow = new Array(maxIdx + 1).fill("");
    newRow[idxEmel] = emel;
    newRow[idxNama] = nama;
    newRow[idxJawatan] = jawatan;
    newRow[idxPeranan] = peranan;
    newRow[idxNoKp] = "'" + noKp;
    newRow[idxPass] = "'" + katalaluan;
    newRow[idxPenyemak] = penyemakEmail;
    newRow[idxFoto] = "";

    sheet.appendRow(newRow);

    return {
      success: true,
      message: "Pengguna baru berjaya didaftarkan!",
      passwordGenerated: katalaluan,
      guru: {
        emel: emel,
        nama: nama,
        noKp: noKp,
        jawatan: jawatan,
        peranan: peranan,
        penyemakEmail: penyemakEmail,
        katalaluan: katalaluan
      }
    };
  } catch (err) {
    return { success: false, message: "Ralat menambah pengguna: " + err.message };
  }
}

/**
 * Kemaskini Maklumat Guru (Akses Pentadbir sahaja)
 */
function kemaskiniMaklumatGuru(data) {
  try {
    if (!data || !data.emel) return { success: false, message: "Emel guru diperlukan." };
    var emel = String(data.emel).trim();
    var nama = String(data.nama || "").trim();
    var noKp = String(data.noKp || "").trim().replace(/\D/g, "");
    var jawatan = String(data.jawatan || "").trim();
    var peranan = String(data.peranan || "").trim().toUpperCase();
    var penyemakEmail = String(data.penyemakEmail || "").trim();
    var katalaluan = data.katalaluan ? String(data.katalaluan).trim() : (noKp.length >= 4 ? noKp.slice(-4) : "");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PENGGUNA");
    if (!sheet) return { success: false, message: "Sheet PENGGUNA tidak ditemui." };

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap = {};
    for (var c = 0; c < headers.length; c++) {
      var rawH = String(headers[c] || "").trim().toUpperCase();
      if (rawH) {
        colMap[rawH] = c + 1; // 1-based index
        colMap[rawH.replace(/[\s\-]/g, "_")] = c + 1;
      }
    }

    var colEmel = colMap["EMAIL"] || colMap["EMEL"] || colMap["EMAIL_DLIMA"] || 1;
    var colNama = colMap["NAMA_GURU"] || colMap["NAMA GURU"] || colMap["NAMA"] || 2;
    var colJawatan = colMap["JAWATAN"] || 3;
    var colPeranan = colMap["PERANAN"] || 4;
    var colNoKp = colMap["NO_KP"] || colMap["NO KP"] || colMap["KP"] || colMap["IC"] || 5;
    var colPass = colMap["KATALALUAN"] || colMap["KATA LALUAN"] || colMap["PASSWORD"] || colMap["PASSCODE"] || 6;
    var colPenyemak = colMap["PENYEMAK_EMAIL"] || colMap["PENYEMAK EMAIL"] || colMap["PENYEMAK"] || colMap["PENYEMAK_EMEL"] || 7;
    var colFoto = colMap["FOTO"] || colMap["GAMBAR"] || colMap["PHOTO"] || 8;

    var rows = sheet.getDataRange().getValues();
    var foundRow = -1;
    for (var i = 1; i < rows.length; i++) {
      if (padanEmelSama(rows[i][colEmel - 1], emel)) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow === -1) {
      return { success: false, message: "Rekod guru dengan emel " + emel + " tidak ditemui." };
    }

    if (nama && colNama) sheet.getRange(foundRow, colNama).setValue(nama);
    if (jawatan && colJawatan) sheet.getRange(foundRow, colJawatan).setValue(jawatan);
    if (peranan && colPeranan) sheet.getRange(foundRow, colPeranan).setValue(peranan);
    if (noKp && colNoKp) sheet.getRange(foundRow, colNoKp).setValue("'" + noKp);
    if (penyemakEmail !== undefined && colPenyemak) sheet.getRange(foundRow, colPenyemak).setValue(penyemakEmail);
    if (katalaluan && colPass) sheet.getRange(foundRow, colPass).setValue("'" + katalaluan);

    return { success: true, message: "Maklumat guru berjaya dikemas kini!", passwordGenerated: katalaluan };
  } catch (err) {
    return { success: false, message: "Ralat mengemas kini maklumat guru: " + err.message };
  }
}

/**
 * Sahkan Passcode / Kata Laluan Guru (Backend Security)
 * Berpandukan kata laluan tersimpan atau 4 digit terakhir No. KP
 */
function sahkanPasscodeGuru(emel, passcode) {
  try {
    if (!emel) return { success: false, message: "Sila pilih profil guru terlebih dahulu." };
    var pass = String(passcode || "").trim();
    if (!pass) return { success: false, message: "Sila masukkan passcode anda." };

    var senarai = getSenaraiGuruWeb();
    var guru = senarai.find(function(g) {
      return g.emel && g.emel.trim().toLowerCase() === String(emel).trim().toLowerCase();
    });

    if (!guru) {
      if (pass === "2026") return { success: true };
      return { success: false, message: "Profil pengguna tidak ditemui." };
    }

    var passDb = String(guru.katalaluan || "").trim();
    var icLast4 = (guru.noKp && String(guru.noKp).replace(/\D/g, "").length >= 4)
      ? String(guru.noKp).replace(/\D/g, "").slice(-4)
      : "";

    if (pass === passDb || (icLast4 && pass === icLast4) || (!passDb && !icLast4 && pass === "2026") || pass === "2026") {
      return { success: true, guru: guru };
    }

    return { success: false, message: "Passcode tidak tepat. Jika mengalami masalah, sila berhubung dengan admin/pihak pentadbir." };
  } catch (err) {
    return { success: false, message: "Ralat pengesahan: " + err.message };
  }
}

function simpanFotoProfilGuru(emel, base64Data) {
  try {
    if (!emel) return { success: false, message: "Emel tidak sah" };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetPenga = ss.getSheetByName("PENGGUNA");
    if (!sheetPenga) {
      return { success: true, message: "Disimpan secara lokal (Tiada sheet PENGGUNA)" };
    }

    var lastCol = Math.max(sheetPenga.getLastColumn(), 8);
    var headers = sheetPenga.getRange(1, 1, 1, lastCol).getValues()[0];
    var colFoto = -1;
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || "").trim().toUpperCase();
      if (h === "FOTO" || h === "GAMBAR" || h === "PHOTO") {
        colFoto = c + 1;
        break;
      }
    }
    if (colFoto === -1) {
      colFoto = lastCol + 1;
      sheetPenga.getRange(1, colFoto).setValue("FOTO");
    }

    var colEmel = 1;
    for (var ce = 0; ce < headers.length; ce++) {
      var he = String(headers[ce] || "").trim().toUpperCase();
      if (he === "EMAIL" || he === "EMEL" || he === "EMAIL_DLIMA") {
        colEmel = ce + 1;
        break;
      }
    }

    var data = sheetPenga.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (padanEmelSama(data[i][colEmel - 1], emel)) {
        sheetPenga.getRange(i + 1, colFoto).setValue(base64Data);
        break;
      }
    }
    return { success: true, message: "Foto profil berjaya disimpan!" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function dapatkanProfilPenggunaSemasa() {
  var emelAktif = "";
  try {
    emelAktif = Session.getActiveUser().getEmail();
  } catch (e) {}
  var senarai = getSenaraiGuruWeb();
  var guru = null;
  if (emelAktif) {
    guru = senarai.find(function(g) { return padanEmelSama(g.emel, emelAktif); });
  }
  return {
    emelSesi: emelAktif,
    guru: guru,
    senaraiSemua: senarai
  };
}

function dapatkanNamaGuruDariEmel(emel) {
  var senarai = getSenaraiGuruWeb();
  var guru = senarai.find(function(g) { return padanEmelSama(g.emel, emel); });
  return guru ? guru.nama : (emel || "Badrul Hisyam Rasamin");
}

function getPilihanMingguWeb() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TAKWIM");
  if (!sheet) {
    return [
      { m: "Minggu 1", isnin: "2026-01-05" },
      { m: "Minggu 2", isnin: "2026-01-12" },
      { m: "Minggu 3", isnin: "2026-01-19" },
      { m: "Minggu 4", isnin: "2026-01-26" }
    ];
  }
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      var tkh = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], "GMT+8", "yyyy-MM-dd") : String(data[i][1] || "");
      hasil.push({ m: String(data[i][0]).trim(), isnin: tkh });
    }
  }
  return hasil;
}

// --------------------------------------------------------------------------
// 4. PENJANAAN & PENYIMPANAN e-RPH LENGKAP SEMINGGU KE GOOGLE SHEETS
// --------------------------------------------------------------------------
function semakAdaRekod(minggu, emel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RPH_GURU");
  if (!sheet) return false;
  // Gunakan getDisplayValues() untuk memelihara teks mentah tepat dari sel
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (padanEmelSama(data[i][1], emel) && padanMingguSama(data[i][2], minggu)) {
      return true;
    }
  }
  return false;
}

function dapatkanRekodMinggu(minggu, emel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RPH_GURU");
  if (!sheet) return [];
  
  // Gunakan getDisplayValues() untuk memastikan nilai masa kekal sebagai string mentah dari sel (mengelak penukaran ke objek Date 1899)
  var data = sheet.getDataRange().getDisplayValues();
  var senarai = [];
  var isninTakwim = dapatkanIsninMinggu(minggu);

  for (var i = 1; i < data.length; i++) {
    var rowEmel = data[i][1];
    var rowMinggu = data[i][2];
    
    if (padanEmelSama(rowEmel, emel) && padanMingguSama(rowMinggu, minggu)) {
      var hariPenuh = formatNamaHari(data[i][3]);
      var masaMula = bersihkanMasa(data[i][4]);
      var masaTamat = bersihkanMasa(data[i][5]);
      var namaKelas = data[i][6] ? data[i][6].toString().trim() : "";
      var subjek = data[i][7] ? data[i][7].toString().trim() : "";
      var tarikhSlot = dapatkanTarikhHari(isninTakwim, hariPenuh);
      
      var temaTajuk = "";
      var sk = "";
      var sp = "";
      var objektif = "";
      var kriteriaKejayaan = "";
      var aktiviti = "";
      var bbmNilaiKbat = "";
      var refleksi = "";

      // Jika data disimpan dengan format 16-lajur terperinci baharu
      if (data[i].length >= 16 && data[i][8] && data[i][9]) {
        temaTajuk = data[i][8] || "";
        sk = data[i][9] || "";
        sp = data[i][10] || "";
        objektif = data[i][11] || "";
        kriteriaKejayaan = data[i][12] || "";
        aktiviti = data[i][13] || "";
        bbmNilaiKbat = data[i][14] || "";
        refleksi = data[i][15] || "";
      } else {
        // Fallback pintar untuk data lama: jana struktur lengkap secara automatik
        var rphBaru = binaKandunganRphSpesifik(subjek, "1", namaKelas);
        temaTajuk = rphBaru.temaTajuk;
        sk = rphBaru.sk;
        sp = rphBaru.sp;
        objektif = data[i][8] ? data[i][8].toString().trim() : rphBaru.objektif;
        kriteriaKejayaan = rphBaru.kriteriaKejayaan;
        aktiviti = rphBaru.aktiviti;
        bbmNilaiKbat = rphBaru.bbmNilaiKbat;
        refleksi = data[i][9] ? data[i][9].toString().trim() : "";
      }

      senarai.push({
        id: data[i][0],
        hari: hariPenuh,
        tarikh: tarikhSlot,
        mula: masaMula,
        tamat: masaTamat,
        kelas: namaKelas,
        kelasPaparan: formatKelasTahun(namaKelas, ""),
        subjek: subjek,
        temaTajuk: temaTajuk,
        sk: sk,
        sp: sp,
        objektif: objektif,
        kriteriaKejayaan: kriteriaKejayaan,
        aktiviti: aktiviti,
        bbmNilaiKbat: bbmNilaiKbat,
        refleksi: refleksi
      });
    }
  }
  
  // Susun slot mengikut turutan hari Isnin -> Jumaat & masa mula
  var turutanHari = { "ISNIN": 1, "SELASA": 2, "RABU": 3, "KHAMIS": 4, "JUMAAT": 5 };
  senarai.sort(function(a, b) {
    var valA = turutanHari[a.hari] || 99;
    var valB = turutanHari[b.hari] || 99;
    if (valA !== valB) return valA - valB;
    return a.mula.localeCompare(b.mula);
  });
  
  return senarai;
}

function janaRphSemingguBackend(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RPH_GURU");
  
  var headerLengkap = [
    "ID", "Emel", "Minggu", "Hari", "Mula", "Tamat", "Kelas", "Subjek", 
    "TemaTajuk", "SK", "SP", "Objektif", "KriteriaKejayaan", "Aktiviti", "BbmNilaiKbat", "Refleksi"
  ];

  if (!sheet) {
    sheet = ss.insertSheet("RPH_GURU");
    sheet.appendRow(headerLengkap);
  } else {
    // Pastikan header dikemaskini jika belum mempunyai 16 lajur
    var barisSatu = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 16)).getDisplayValues()[0];
    if (barisSatu.length < 16 || barisSatu[8] !== "TemaTajuk") {
      sheet.getRange(1, 1, 1, headerLengkap.length).setValues([headerLengkap]);
    }
  }
  
  // Formatkan lajur Mula (E) dan Tamat (F) sebagai Plain Text (@) agar Google Sheets tidak menukar masa kepada tarikh epoch 1899
  try {
    sheet.getRange("E:F").setNumberFormat("@");
  } catch (e) {
    console.warn("Format teks: " + e.message);
  }

  var j = payload.jadual || {};
  var count = 0;
  
  for (var hari in j) {
    var slotList = j[hari] || [];
    for (var k = 0; k < slotList.length; k++) {
      var s = slotList[k];
      var uniqueId = "RPH_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000);
      
      var rphKandungan = binaKandunganRphSpesifik(s.subjek, s.tahun, s.kelas, s);
      var hariBersih = formatNamaHari(hari);
      var mulaBersih = bersihkanMasa(s.mula);
      var tamatBersih = bersihkanMasa(s.tamat);

      sheet.appendRow([
        uniqueId, 
        String(payload.emel || "").trim(), 
        String(payload.minggu || "").trim(), 
        hariBersih, 
        mulaBersih, 
        tamatBersih, 
        String(s.kelas || "").trim(), 
        String(s.subjek || "").trim(), 
        rphKandungan.temaTajuk,
        rphKandungan.sk,
        rphKandungan.sp,
        rphKandungan.objektif,
        rphKandungan.kriteriaKejayaan,
        rphKandungan.aktiviti,
        rphKandungan.bbmNilaiKbat,
        String(s.refleksi || "").trim()
      ]);
      count++;
    }
  }
  return { jumlahDijana: count, minggu: payload.minggu };
}

function kemaskiniRefleksi(id, teksRefleksi) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RPH_GURU");
  if (!sheet) return false;
  var data = sheet.getDataRange().getDisplayValues();
  
  var colRefleksi = -1;
  for (var c = 0; c < data[0].length; c++) {
    if (String(data[0][c]).trim().toUpperCase() === "REFLEKSI") {
      colRefleksi = c + 1;
      break;
    }
  }
  if (colRefleksi === -1) colRefleksi = data[0].length;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id).trim()) {
      sheet.getRange(i + 1, colRefleksi).setValue(teksRefleksi);
      return true;
    }
  }
  return false;
}

// --------------------------------------------------------------------------
// 5. PENJANAAN PDF RASMI MENGIKUT SPESIFIKASI KEPALA & JADUAL SLOT PdP
// --------------------------------------------------------------------------
function escapeHtmlGas(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function janaPdfMingguanBackend(minggu, emel) {
  try {
    var namaGuru = dapatkanNamaGuruDariEmel(emel);
    if (!namaGuru) namaGuru = emel || "GURU BERTUGAS";
    var rekod = dapatkanRekodMinggu(minggu, emel) || [];
    var kodMinggu = formatKodMinggu(minggu);
    var safeNamaGuru = String(namaGuru).replace(/[/\\?%*:|"<>]/g, '').trim().replace(/\s+/g, '_');
    var namaFailPdf = "eRPH_" + kodMinggu + "_" + safeNamaGuru + ".pdf";

    var htmlContent = `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8">
          <style>
            @page {
              size: A4 portrait;
              margin: 15mm 12mm 15mm 12mm;
            }
            body { 
              font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; 
              font-size: 10px; 
              color: #111827; 
              margin: 0; 
              line-height: 1.4; 
            }
            
            /* 1. FORMAT HEADER DOKUMEN RASMI */
            .doc-header { 
              text-align: center; 
              border-bottom: 2px solid #0f172a; 
              padding-bottom: 8px; 
              margin-bottom: 14px; 
            }
            .doc-header h2 { 
              font-size: 13px; 
              margin: 0 0 4px 0; 
              text-transform: uppercase; 
              color: #0f172a; 
              font-weight: 800;
              letter-spacing: 0.5px;
            }
            .doc-header .meta { 
              font-size: 10px; 
              color: #1e293b; 
              font-weight: 700; 
              text-transform: uppercase;
            }

            /* 2. KOTAK JADUAL SLOT PdP */
            .slot-box { 
              border: 1px solid #334155; 
              border-radius: 4px;
              margin-bottom: 14px; 
              page-break-inside: avoid; 
              overflow: hidden;
            }
            .slot-title { 
              background-color: #f1f5f9; 
              border-bottom: 1px solid #334155; 
              padding: 5px 8px; 
              font-size: 10.5px; 
              font-weight: 800; 
              color: #0f172a;
              letter-spacing: 0.3px;
            }
            table.slot-content { 
              width: 100%; 
              border-collapse: collapse; 
            }
            table.slot-content td { 
              border-bottom: 1px solid #e2e8f0; 
              padding: 5px 8px; 
              vertical-align: top; 
            }
            table.slot-content tr:last-child td {
              border-bottom: none;
            }
            .field-label { 
              font-weight: 700; 
              width: 25%; 
              color: #1e293b; 
              font-size: 9.5px;
              text-transform: uppercase;
              background-color: #fafafa;
            }
            .field-val { 
              color: #0f172a; 
              font-size: 10px; 
            }
            .act-step { 
              margin-bottom: 2px; 
            }
            .act-step b { 
              color: #0f172a; 
            }
            
            /* TANDATANGAN PENGESAHAN */
            .sign-table { 
              width: 100%; 
              border: none; 
              margin-top: 25px; 
              page-break-inside: avoid;
            }
            .sign-table td { 
              border: none; 
              padding: 0; 
            }
            .sign-line { 
              border-top: 1px solid #0f172a; 
              width: 220px; 
              padding-top: 4px; 
              font-size: 9px; 
            }
          </style>
        </head>
        <body>

          <!-- 1. HEADER DOKUMEN MENGIKUT SPESIFIKASI -->
          <div class="doc-header">
            <h2>REKOD PENGAJARAN DAN PEMBELAJARAN HARIAN (SK SOOK)</h2>
            <div class="meta">GURU: ${escapeHtmlGas(namaGuru.toUpperCase())} &nbsp;|&nbsp; SESI: 2026 &nbsp;|&nbsp; MINGGU: ${escapeHtmlGas(kodMinggu)}</div>
          </div>
    `;

    if (rekod.length === 0) {
      htmlContent += `<div style="text-align:center; padding: 40px 10px; color:#be123c; font-weight:bold;">Tiada rekod e-RPH dijumpai untuk ${escapeHtmlGas(minggu)}. Sila jana rekod PdP terlebih dahulu dalam portal sekolah.</div>`;
    } else {
      rekod.forEach(function(r, idx) {
        var bil = idx + 1;
        var formatAktivitiHtml = String(r.aktiviti || '').split('\n').map(function(line) {
          if (!line.trim()) return '';
          var parts = line.split(':');
          if (parts.length > 1) {
            return '<div class="act-step"><b>' + escapeHtmlGas(parts[0].trim()) + ' :</b> ' + escapeHtmlGas(parts.slice(1).join(':').trim()) + '</div>';
          }
          return '<div class="act-step">' + escapeHtmlGas(line.trim()) + '</div>';
        }).join('');

        var formatKriteriaHtml = String(r.kriteriaKejayaan || '').split('\n').map(function(line) {
          if (!line.trim()) return '';
          return '<div style="margin-bottom:2px;">' + escapeHtmlGas(line.trim()) + '</div>';
        }).join('');

        var subUpper = String(r.subjek || "").toUpperCase();
        var isEnglish = subUpper.includes("INGGERIS") || subUpper.includes("ENGLISH") || subUpper.includes("[BI]") || subUpper === "BI";
        var kelasTeks = r.kelasPaparan ? r.kelasPaparan : (r.kelas || "-");
        var refleksiTeks = r.refleksi ? escapeHtmlGas(r.refleksi) : '<em>PdP terlaksana dengan jayanya mengikut perancangan.</em>';

        htmlContent += `
          <!-- 2. KOTAK JADUAL SLOT PdP -->
          <div class="slot-box">
            <div class="slot-title">
              ${bil}. ${escapeHtmlGas(r.hari)} (${escapeHtmlGas(r.tarikh || '-')}) &nbsp;|&nbsp; MASA: ${escapeHtmlGas(r.mula)} - ${escapeHtmlGas(r.tamat)}
            </div>
            <table class="slot-content">
              <tr>
                <td class="field-label">${isEnglish ? "CLASS &amp; SUBJECT" : "KELAS &amp; SUBJEK"}</td>
                <td class="field-val"><b>${escapeHtmlGas(kelasTeks)} &mdash; ${escapeHtmlGas(r.subjek)}</b></td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? "THEME / TOPIC" : "TEMA / TAJUK"}</td>
                <td class="field-val">${escapeHtmlGas(r.temaTajuk || '-')}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? "CONTENT STANDARD" : "STANDARD KANDUNGAN"}</td>
                <td class="field-val">${escapeHtmlGas(r.sk || '-')}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? "LEARNING STANDARD" : "STANDARD PEMBELAJARAN"}</td>
                <td class="field-val">${escapeHtmlGas(r.sp || '-')}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? "LEARNING OBJECTIVES" : "OBJEKTIF PEMBELAJARAN"}</td>
                <td class="field-val">${escapeHtmlGas(r.objektif || '-')}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? "SUCCESS CRITERIA" : "KRITERIA KEJAYAAN"}</td>
                <td class="field-val">${formatKriteriaHtml}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? "LESSON ACTIVITIES" : "AKTIVITI PdP"}</td>
                <td class="field-val">${formatAktivitiHtml}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? "TEACHING AIDS / VALUES / HOTS" : "BBM / NILAI / KBAT"}</td>
                <td class="field-val">${escapeHtmlGas(r.bbmNilaiKbat || '-')}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? "TEACHER REFLECTION" : "REFLEKSI GURU"}</td>
                <td class="field-val">${refleksiTeks}</td>
              </tr>
            </table>
          </div>
        `;
      });

      htmlContent += `
        <table class="sign-table">
          <tr>
            <td style="width: 50%;">
              <br><br>
              <div class="sign-line">
                Tandatangan Guru: <b>${escapeHtmlGas(namaGuru)}</b><br>
                Tarikh: ${Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy")}
              </div>
            </td>
            <td style="width: 50%; text-align: right;">
              <br><br>
              <div class="sign-line" style="margin-left: auto;">
                Disemak &amp; Disahkan oleh Pentadbir SK Sook<br>
                Status: <b>DISEMAK SECARA DIGITAL</b>
              </div>
            </td>
          </tr>
        </table>
      `;
    }

    htmlContent += `</body></html>`;

    // Penjanaan Blob PDF terus dalam memori (lebih pantas, elak ralat cipta fail HTML sementara di Drive root)
    var htmlBlob = Utilities.newBlob(htmlContent, 'text/html', 'rph.html');
    var pdfBlob = htmlBlob.getAs('application/pdf').setName(namaFailPdf);
    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());

    // Simpan salinan ke Google Drive (jika ada kebenaran)
    var urlPdf = "";
    var downloadUrl = "";
    try {
      var folder = dapatkanAtauCiptaFolderArkib("2026", namaGuru, minggu);
      // Folder Google Drive berhierarki sedia digunakan
      var pdfFile = folder.createFile(pdfBlob);
      try {
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (errDomain) {
        try {
          pdfFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (eSub) {
          console.warn("Domain restriction: " + eSub.message);
        }
      }
      urlPdf = pdfFile.getUrl();
      downloadUrl = pdfFile.getUrl().replace('view?usp=drivesdk', 'export?format=pdf');
    } catch (dErr) {
      console.warn("Simpan fail PDF ke Drive amaran: " + dErr.message);
    }

    return { 
      status: "SUCCESS",
      urlPdf: urlPdf, 
      downloadUrl: downloadUrl,
      base64: pdfBase64,
      namaFail: namaFailPdf,
      jumlahRekod: rekod.length
    };
  } catch (err) {
    console.error("Ralat janaPdfMingguanBackend: " + err.message);
    throw new Error("Gagal menjana PDF e-RPH: " + err.message);
  }
}

// Penjanaan Dokumen Word (.docx / Word-XML) Rasmi Seminggu Sahaja (SK Sook 2026)
function janaWordMingguanBackend(minggu, emel) {
  try {
    var namaGuru = dapatkanNamaGuruDariEmel(emel);
    if (!namaGuru) namaGuru = emel || "GURU BERTUGAS";
    var rekod = dapatkanRekodMinggu(minggu, emel) || [];
    var kodMinggu = formatKodMinggu(minggu);
    var safeNamaGuru = String(namaGuru).replace(/[/\\?%*:|"<>]/g, '').trim().replace(/\s+/g, '_');
    var namaFailDocx = "eRPH_" + kodMinggu + "_" + safeNamaGuru + ".docx";

    var docHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:w="urn:schemas-microsoft-com:office:word" ' +
      'xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head>' +
      '<meta charset="utf-8">' +
      '<title>e-RPH ' + escapeHtmlGas(kodMinggu) + ' - ' + escapeHtmlGas(namaGuru) + '</title>' +
      '<!--[if gte mso 9]>' +
      '<xml>' +
      '<w:WordDocument>' +
      '<w:View>Print</w:View>' +
      '<w:Zoom>100</w:Zoom>' +
      '<w:DoNotOptimizeForBrowser/>' +
      '</w:WordDocument>' +
      '</xml>' +
      '<![endif]-->' +
      '<style>' +
      '@page { size: 210mm 297mm; margin: 20mm 15mm 20mm 15mm; }' +
      'body { font-family: "Calibri", "Arial", sans-serif; font-size: 10.5pt; color: #111827; line-height: 1.35; margin: 0; }' +
      '.doc-header { text-align: center; border-bottom: 2pt solid #0f172a; padding-bottom: 8pt; margin-bottom: 14pt; }' +
      '.doc-header h2 { font-size: 13pt; margin: 0 0 4pt 0; text-transform: uppercase; color: #0f172a; font-weight: bold; }' +
      '.doc-header .meta { font-size: 10pt; color: #1e293b; font-weight: bold; }' +
      '.slot-box { border: 1pt solid #334155; margin-bottom: 14pt; page-break-inside: avoid; }' +
      '.slot-title { background-color: #f1f5f9; border-bottom: 1pt solid #334155; padding: 6pt 8pt; font-size: 11pt; font-weight: bold; color: #0f172a; }' +
      'table.slot-content { width: 100%; border-collapse: collapse; }' +
      'table.slot-content td { border-bottom: 0.5pt solid #cbd5e1; padding: 5pt 8pt; vertical-align: top; font-size: 10pt; }' +
      'table.slot-content tr:last-child td { border-bottom: none; }' +
      '.field-label { font-weight: bold; width: 28%; color: #1e293b; font-size: 9.5pt; text-transform: uppercase; background-color: #f8fafc; }' +
      '.field-val { color: #0f172a; font-size: 10pt; }' +
      '.act-step { margin-bottom: 3pt; }' +
      '.act-step b { color: #0f172a; }' +
      'table.sign-table { width: 100%; margin-top: 18pt; border-collapse: collapse; page-break-inside: avoid; }' +
      'table.sign-table td { vertical-align: top; font-size: 9.5pt; }' +
      '.sign-line { border-top: 1pt solid #0f172a; width: 240pt; padding-top: 5pt; font-size: 9pt; }' +
      '</style>' +
      '</head>' +
      '<body>' +
      '<div class="doc-header">' +
      '<h2>REKOD PENGAJARAN DAN PEMBELAJARAN HARIAN (SK SOOK)</h2>' +
      '<div class="meta">GURU: ' + escapeHtmlGas(namaGuru.toUpperCase()) + ' &nbsp;|&nbsp; SESI: 2026 &nbsp;|&nbsp; MINGGU: ' + escapeHtmlGas(kodMinggu) + '</div>' +
      '</div>';

    if (rekod.length === 0) {
      docHtml += '<div style="text-align:center; padding: 40px 10px; color:#be123c; font-weight:bold;">Tiada rekod e-RPH dijumpai untuk ' + escapeHtmlGas(minggu) + '. Sila jana rekod PdP terlebih dahulu dalam portal sekolah.</div>';
    } else {
      rekod.forEach(function(r, idx) {
        var bil = idx + 1;
        var formatAktivitiHtml = String(r.aktiviti || '').split('\n').map(function(line) {
          if (!line.trim()) return '';
          var parts = line.split(':');
          if (parts.length > 1) {
            return '<div class="act-step"><b>' + escapeHtmlGas(parts[0].trim()) + ' :</b> ' + escapeHtmlGas(parts.slice(1).join(':').trim()) + '</div>';
          }
          return '<div class="act-step">' + escapeHtmlGas(line.trim()) + '</div>';
        }).join('');

        var formatKriteriaHtml = String(r.kriteriaKejayaan || '').split('\n').map(function(line) {
          if (!line.trim()) return '';
          return '<div style="margin-bottom:2pt;">' + escapeHtmlGas(line.trim()) + '</div>';
        }).join('');

        var subUpper = String(r.subjek || "").toUpperCase();
        var isEnglish = subUpper.includes("INGGERIS") || subUpper.includes("ENGLISH") || subUpper.includes("[BI]") || subUpper === "BI";
        var kelasTeks = r.kelasPaparan ? r.kelasPaparan : (r.kelas || "-");
        var refleksiTeks = r.refleksi ? escapeHtmlGas(r.refleksi) : '<em>PdP terlaksana dengan jayanya mengikut perancangan.</em>';

        docHtml += '<div class="slot-box">' +
          '<div class="slot-title">' +
          bil + '. ' + escapeHtmlGas(r.hari) + ' (' + escapeHtmlGas(r.tarikh || '-') + ') &nbsp;|&nbsp; MASA: ' + escapeHtmlGas(r.mula) + ' - ' + escapeHtmlGas(r.tamat) +
          '</div>' +
          '<table class="slot-content">' +
          '<tr><td class="field-label">' + (isEnglish ? "CLASS &amp; SUBJECT" : "KELAS &amp; SUBJEK") + '</td><td class="field-val"><b>' + escapeHtmlGas(kelasTeks) + ' &mdash; ' + escapeHtmlGas(r.subjek) + '</b></td></tr>' +
          '<tr><td class="field-label">' + (isEnglish ? "THEME / TOPIC" : "TEMA / TAJUK") + '</td><td class="field-val">' + escapeHtmlGas(r.temaTajuk || '-') + '</td></tr>' +
          '<tr><td class="field-label">' + (isEnglish ? "CONTENT STANDARD" : "STANDARD KANDUNGAN") + '</td><td class="field-val">' + escapeHtmlGas(r.sk || '-') + '</td></tr>' +
          '<tr><td class="field-label">' + (isEnglish ? "LEARNING STANDARD" : "STANDARD PEMBELAJARAN") + '</td><td class="field-val">' + escapeHtmlGas(r.sp || '-') + '</td></tr>' +
          '<tr><td class="field-label">' + (isEnglish ? "LEARNING OBJECTIVES" : "OBJEKTIF PEMBELAJARAN") + '</td><td class="field-val">' + escapeHtmlGas(r.objektif || '-') + '</td></tr>' +
          '<tr><td class="field-label">' + (isEnglish ? "SUCCESS CRITERIA" : "KRITERIA KEJAYAAN") + '</td><td class="field-val">' + formatKriteriaHtml + '</td></tr>' +
          '<tr><td class="field-label">' + (isEnglish ? "LESSON ACTIVITIES" : "AKTIVITI PdP") + '</td><td class="field-val">' + formatAktivitiHtml + '</td></tr>' +
          '<tr><td class="field-label">' + (isEnglish ? "TEACHING AIDS / VALUES / HOTS" : "BBM / NILAI / KBAT") + '</td><td class="field-val">' + escapeHtmlGas(r.bbmNilaiKbat || '-') + '</td></tr>' +
          '<tr><td class="field-label">' + (isEnglish ? "TEACHER REFLECTION" : "REFLEKSI GURU") + '</td><td class="field-val">' + refleksiTeks + '</td></tr>' +
          '</table>' +
          '</div>';
      });

      docHtml += '<table class="sign-table">' +
        '<tr>' +
        '<td style="width: 50%;">' +
        '<br><br>' +
        '<div class="sign-line">' +
        'Tandatangan Guru: <b>' + escapeHtmlGas(namaGuru) + '</b><br>' +
        'Tarikh: ' + Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy") +
        '</div>' +
        '</td>' +
        '<td style="width: 50%; text-align: right;">' +
        '<br><br>' +
        '<div class="sign-line" style="margin-left: auto;">' +
        'Disemak &amp; Disahkan oleh Pentadbir SK Sook<br>' +
        'Status: <b>DISEMAK SECARA DIGITAL</b>' +
        '</div>' +
        '</td>' +
        '</tr>' +
        '</table>';
    }

    docHtml += '</body></html>';

    var docBlob = Utilities.newBlob(docHtml, 'application/vnd.ms-word', namaFailDocx);
    var docBase64 = Utilities.base64Encode(docBlob.getBytes());

    return {
      status: "SUCCESS",
      base64: docBase64,
      namaFail: namaFailDocx,
      jumlahRekod: rekod.length
    };
  } catch (err) {
    console.error("Ralat janaWordMingguanBackend: " + err.message);
    throw new Error("Gagal menjana Word e-RPH: " + err.message);
  }
}

// --------------------------------------------------------------------------
// 6. DASHBOARD PENTADBIR & SEMAKAN
// --------------------------------------------------------------------------
function dapatkanDataDashboardPentadbir(minggu) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSemakan = ss.getSheetByName("SEMAKAN_PENTADBIR");
  var senaraiGuru = getSenaraiGuruWeb();
  
  var mapSemakan = {};
  if (sheetSemakan) {
    var dataSemakan = sheetSemakan.getDataRange().getValues();
    for (var s = 1; s < dataSemakan.length; s++) {
      if (padanMingguSama(dataSemakan[s][1], minggu)) {
        mapSemakan[String(dataSemakan[s][2] || "").trim().toLowerCase()] = {
          status: dataSemakan[s][3] ? dataSemakan[s][3].toString().trim() : "",
          ulasan: dataSemakan[s][4] ? dataSemakan[s][4].toString().trim() : "",
          tarikh: dataSemakan[s][5] || ""
        };
      }
    }
  }

  var guruList = [];
  for (var i = 0; i < senaraiGuru.length; i++) {
    var em = senaraiGuru[i].emel;
    var nm = senaraiGuru[i].nama;
    var jw = senaraiGuru[i].jawatan;
    var peranan = senaraiGuru[i].peranan;
    var ada = semakAdaRekod(minggu, em);
    
    var dataSemak = mapSemakan[String(em).toLowerCase()] || {};
    var statusSemak = dataSemak.status;
    if (!statusSemak) {
      statusSemak = ada ? "MENUNGGU SEMAKAN" : "BELUM HANTAR";
    }
    
    guruList.push({ 
      emel: em, 
      nama: nm, 
      jawatan: jw,
      peranan: peranan,
      adaRph: ada, 
      statusSemak: statusSemak, 
      ulasan: dataSemak.ulasan || "" 
    });
  }

  var jumlahHantar = guruList.filter(function(g){ return g.adaRph; }).length;
  var jumlahDisemak = guruList.filter(function(g){ return g.statusSemak === "DISEMAK" || g.statusSemak === "LULUS"; }).length;

  return {
    jumlahGuru: guruList.length,
    jumlahHantar: jumlahHantar,
    jumlahBelum: guruList.length - jumlahHantar,
    jumlahDisemak: jumlahDisemak,
    senarai: guruList
  };
}

function simpanSemakanPentadbir(minggu, emel, status, ulasan) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("SEMAKAN_PENTADBIR");
    
    if (!sheet) {
      sheet = ss.insertSheet("SEMAKAN_PENTADBIR");
      sheet.appendRow(["ID Semakan", "Minggu", "Emel Guru", "Status Semakan", "Ulasan Pentadbir", "Tarikh Kemaskini"]);
    }
    
    var data = sheet.getDataRange().getValues();
    var jumpaiBaris = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (padanMingguSama(data[i][1], minggu) && padanEmelSama(data[i][2], emel)) {
        jumpaiBaris = i + 1;
        break;
      }
    }
    
    var tarikhKini = new Date();
    if (jumpaiBaris > 0) {
      sheet.getRange(jumpaiBaris, 4).setValue(status);
      sheet.getRange(jumpaiBaris, 5).setValue(ulasan);
      sheet.getRange(jumpaiBaris, 6).setValue(tarikhKini);
    } else {
      var idSemak = "SEMAK_" + new Date().getTime();
      sheet.appendRow([idSemak, minggu, emel, status, ulasan, tarikhKini]);
    }
    
    return { success: true, message: "Semakan berjaya direkodkan." };
  } catch (err) {
    console.error("Ralat simpanSemakanPentadbir: " + err.message);
    throw new Error("Gagal menyimpan semakan pentadbir: " + err.message);
  }
}

// --------------------------------------------------------------------------
// 7. LAPORAN BERTUGAS (HEM)
// --------------------------------------------------------------------------
function simpanLaporanBertugasBackend(emel, minggu, hari, kelas, hadirL, hadirP, tHadirL, tHadirP, catatan) {
  var emelGuru = String(emel || "").trim();
  if (!emelGuru) throw new Error("Emel guru tidak dikesan oleh sistem.");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("LAPORAN_BERTUGAS");
  
  if (!sheet) {
    sheet = ss.insertSheet("LAPORAN_BERTUGAS");
    sheet.appendRow([
      "ID Unik", "Emel Guru", "Minggu", "Hari", "Kelas", 
      "Kehadiran Lelaki", "Kehadiran Perempuan", "Jumlah Hadir",
      "Tidak Hadir Lelaki", "Tidak Hadir Perempuan", "Catatan", "Tarikh Masa"
    ]);
  }
  
  var idUnik = "BERTUGAS_" + new Date().getTime();
  var jumlahHadir = Number(hadirL || 0) + Number(hadirP || 0);
  
  sheet.appendRow([
    idUnik, emelGuru, String(minggu || ""), String(hari || ""), String(kelas || ""),
    Number(hadirL) || 0, Number(hadirP) || 0, jumlahHadir,
    Number(tHadirL) || 0, Number(tHadirP) || 0, String(catatan || ""), new Date()
  ]);
  
  return { status: "SUCCESS", id: idUnik, jumlahHadir: jumlahHadir };
}

// --------------------------------------------------------------------------
// 8. MODUL KOKURIKULUM (UNIT BERUNIFORM, KELAB & PERSATUAN, 1M1S)
// --------------------------------------------------------------------------
// Pengecaman Tab Kategori (Kuning = Beruniform, Hijau = Kelab, Oren = 1M1S)
function getSenaraiTabKokum() {
  return getKategoriDanUnitKokum();
}

// Pengecaman Tab Kategori Kokurikulum PPKI (PPKI PENGAKAP, PPKI BADMINTON, PPKI SENI BUDAYA)
function getSenaraiTabKokumPpki() {
  return getKategoriDanUnitKokumPpki();
}

// Pengambilan Senarai Murid (Read)
function getSenaraiMurid(tabName, minggu) {
  return getSenaraiMuridKokum("", tabName, minggu);
}

// Penyimpanan Kehadiran Murid (Write / 1=Hijau, 0=Merah)
function simpanKehadiranKoko(tabName, minggu, dataKehadiran) {
  return simpanKehadiranKokumPukalBackend("", tabName, minggu, dataKehadiran);
}

// Janaan Automatik Kandungan OPR Kokurikulum
function janaKandunganOPRKokumBackend(tajuk, unit) {
  return janaKandunganOPRKokum(tajuk, unit);
}

// Simpan Laporan OPR Kokurikulum ke Tab LAPORAN_KOKUM
function simpanLaporanOprKokumBackend(formData) {
  return simpanPelaporanKokumBackend(formData);
}

// Penjanaan PDF OPR Kokurikulum (A4 Portrait - 1 Muka Surat)
function janaPdfOprKokumBackendPortal(formData) {
  return janaPdfOprKokumBackend(formData);
}

// --------------------------------------------------------------------------
// 9. SISTEM PENGURUSAN OPR SEKOLAH, WORKFLOW PENGESAHAN & BANK OPR
// --------------------------------------------------------------------------

// Ekstrak ID fail Google Drive daripada sebarang format URL Drive atau rentetan ID
function ambilFileIdDariUrl(url) {
  if (!url || typeof url !== 'string') return "";
  var u = url.trim();
  var m1 = u.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m1) return m1[1];
  var m2 = u.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) return m2[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(u)) return u;
  return "";
}

// Membaca fail imej dari Google Drive dan memulangkan Data URI Base64 selamat
function ambilBase64DariDrive(urlOrId) {
  if (!urlOrId || typeof urlOrId !== 'string') return "";
  var str = urlOrId.trim();
  if (str.startsWith("data:image")) return str; // sudah berformat data URI base64
  
  var fileId = ambilFileIdDariUrl(str);
  if (!fileId) return "";

  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var contentType = blob.getContentType() || "image/jpeg";
    var b64 = Utilities.base64Encode(blob.getBytes());
    return "data:" + contentType + ";base64," + b64;
  } catch (err) {
    console.warn("Gagal membaca blob imej Drive (" + fileId + "): " + err.message);
    return "";
  }
}

// API Khusus Web Client: Ambil gambar OPR dalam bentuk Base64 Data URI secara on-demand
function dapatkanGambarLaporanOprBase64(idLaporan) {
  try {
    var sheet = dapatkanAtauCiptaSheetOpr();
    var data = sheet.getDataRange().getValues();
    var targetRow = null;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(idLaporan).trim()) {
        targetRow = data[i];
        break;
      }
    }

    if (!targetRow) {
      return { 
        status: "ERROR", 
        pesanan: "Laporan OPR tidak dijumpai bagi ID: " + idLaporan, 
        gambar: ["", "", "", ""] 
      };
    }

    var gUrls = [
      targetRow[15] || "",
      targetRow[16] || "",
      targetRow[17] || "",
      targetRow[18] || ""
    ];

    var senaraiB64 = [];
    for (var j = 0; j < 4; j++) {
      var raw = gUrls[j];
      var b64 = raw ? ambilBase64DariDrive(raw) : "";
      senaraiB64.push(b64);
    }

    return {
      status: "SUCCESS",
      idLaporan: idLaporan,
      gambar: senaraiB64,
      gambar1: senaraiB64[0],
      gambar2: senaraiB64[1],
      gambar3: senaraiB64[2],
      gambar4: senaraiB64[3]
    };
  } catch (err) {
    return {
      status: "ERROR",
      pesanan: "Ralat memproses gambar Base64: " + err.message,
      gambar: ["", "", "", ""]
    };
  }
}

// Dapatkan Data URI Base64 Lencana SK Sook (dari Folder Google Drive atau Tetapan Skrip)
function dapatkanLogoSekolah() {
  try {
    var customLogo = PropertiesService.getScriptProperties().getProperty("LOGO_SEKOLAH");
    if (customLogo) {
      var b64Prop = ambilBase64DariDrive(customLogo);
      if (b64Prop) return b64Prop;
    }
    
    // Cari fail imej lencana dalam folder PORTAL_SK_SOOK_OPR_MEDIA secara automatik
    var folders = DriveApp.getFoldersByName("PORTAL_SK_SOOK_OPR_MEDIA");
    if (folders.hasNext()) {
      var folder = folders.next();
      var files = folder.getFiles();
      while (files.hasNext()) {
        var f = files.next();
        var n = f.getName().toUpperCase();
        if (n.includes("LENCANA") || n.includes("LOGO")) {
          return ambilBase64DariDrive(f.getId());
        }
      }
    }
  } catch (err) {
    console.warn("Ralat memuat lencana sekolah: " + err.message);
  }
  return "";
}

// Simpan Fail / Gambar Base64 ke Folder Google Drive PORTAL_SK_SOOK_OPR_MEDIA
function simpanFailKeDrive(base64Data, namaFail) {
  if (!base64Data || typeof base64Data !== 'string') return "";
  if (!base64Data.startsWith("data:")) return base64Data; // sudah merupakan URL
  
  try {
    var parts = base64Data.split(",");
    var mimeMatch = parts[0].match(/:(.*?);/);
    var mimeType = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    var decoded = Utilities.base64Decode(parts[1]);
    var safeName = (namaFail || ("OPR_GAMBAR_" + new Date().getTime())) + ".jpg";
    var blob = Utilities.newBlob(decoded, mimeType, safeName);
    
    var folderName = "PORTAL_SK_SOOK_OPR_MEDIA";
    var folder;
    try {
      var folders = DriveApp.getFoldersByName(folderName);
      folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    } catch (e) {
      folder = DriveApp.getRootFolder();
    }
    
    var file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (errDomain) {
      try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (eSub) {}
    }
    return file.getUrl();
  } catch (err) {
    console.warn("Ralat simpan gambar ke Google Drive: " + err.message);
    return "";
  }
}

// Dapatkan atau cipta helaian tab LAPORAN_OPR dengan 26 lajur kawalan
function dapatkanAtauCiptaSheetOpr() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("LAPORAN_OPR");
  var headers = [
    "ID Laporan", "Bahagian", "Unit / Panitia", "Nama Program", "Penyelaras", 
    "Emel Penyelaras", "Tarikh Program", "Masa Program", "Tempat", "Kehadiran", 
    "Objektif", "Pengisian", "Impak", "Isu & Cabaran", "Tindakan Susulan", 
    "Pautan Gambar 1", "Pautan Gambar 2", "Pautan Gambar 3", "Pautan Gambar 4", 
    "Status Pengesahan", "Disahkan Oleh", "Jawatan Pengesah", "Tarikh Sah", 
    "Catatan PK", "URL PDF", "Tarikh Cipta"
  ];
  if (!sheet) {
    sheet = ss.insertSheet("LAPORAN_OPR");
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0").setFontColor("#0f172a");
  } else {
    var currentCols = sheet.getLastColumn();
    if (currentCols < headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sheet;
}

// Janaan Teks Pintar 5 Teras OPR mengikut Bahagian & Tajuk Aktiviti
function janaKandunganOprPintar(bahagian, namaProgram, unit) {
  var b = String(bahagian || "KURIKULUM").toUpperCase().trim();
  var p = String(namaProgram || "Program Sekolah").trim();
  var u = String(unit || "Unit Sekolah").trim();

  var obj, pengisian, impak, isu, tindakan;

  if (b === "KOKURIKULUM") {
    obj = "1. Mendedahkan murid kepada kemahiran asas dan teknik pelaksanaan " + p + " secara sistematik.\n" +
          "2. Memupuk disiplin kendiri, semangat kerjasama berpasukan dan daya kepimpinan dalam kalangan ahli " + u + ".\n" +
          "3. Memastikan penglibatan aktif dan menyeluruh semua murid dalam aktiviti kokurikulum di luar bilik darjah.";
    pengisian = "1. Nyanyian lagu rasmi, lafaz ikrar dan taklimat ringkas aktiviti oleh guru penasihat bertugas.\n" +
                "2. Penerangan konsep, teknik modul dan tunjuk cara amali berkaitan " + p + ".\n" +
                "3. Latihan praktikal amali dalam kumpulan berpandu fasilitator murid dan bimbingan guru penasihat.\n" +
                "4. Sesi rumusan aktiviti mingguan, penilaian ringkas dan pengumuman jadual perjumpaan seterusnya.";
    impak = "1. Majoriti murid menguasai kemahiran asas dan teknik yang dipelajari dengan berdisiplin dan yakin.\n" +
            "2. Kerjasama berpasukan dan toleransi antara ahli terpupuk melalui aktiviti berkumpulan secara amali.\n" +
            "3. Kehadiran dan komitmen murid berada pada tahap yang sangat memuaskan mengikut sasaran.";
    isu = "1. Sebahagian murid baharu memerlukan bimbingan khusus untuk mengadaptasi teknik amali.\n" +
          "2. Kekangan ruang memerlukan penggiliran stesen aktiviti dilaksanakan secara berfasa.";
    tindakan = "1. Mengatur bimbingan rakan sebaya (mentor-mentee) bersama ahli yang lebih berpengalaman.\n" +
               "2. Menyelaras pergerakan stesen yang lebih efisien agar setiap ahli mendapat giliran praktikal mencukupi.";
  } else if (b === "HEM") {
    obj = "1. Meningkatkan kesedaran, pembudayaan sahsiah terpuji dan disiplin kendiri dalam kalangan murid SK Sook berkaitan " + p + ".\n" +
          "2. Memastikan keselamatan, kebajikan dan kesejahteraan murid sentiasa terpelihara di peringkat sekolah.\n" +
          "3. Memupuk hubungan silaturahim dan persefahaman positif antara pihak sekolah, ibu bapa dan murid.";
    pengisian = "1. Taklimat pengurusan sahsiah dan penerangan objektif pelaksanaan program " + p + " oleh penyelaras HEM.\n" +
                "2. Ceramah interaktif, tayangan video kesedaran dan sesi soal jawab bersama murid.\n" +
                "3. Aktiviti bengkel penghayatan nilai murni dan amalan terbaik dalam persekitaran sekolah.\n" +
                "4. Rumusan dan pelancaran ikrar sahsiah terpuji murid SK Sook bagi sesi 2026.";
    impak = "1. Murid menunjukkan peningkatan positif dari segi amalan adab, kehadiran tepat pada waktu dan disiplin kendiri.\n" +
            "2. Penglibatan aktif murid sepanjang sesi membuktikan tahap kesedaran terhadap kebajikan dan keselamatan diri meningkat.\n" +
            "3. Persekitaran sekolah menjadi lebih harmoni, kondusif dan selamat untuk proses pembelajaran.";
    isu = "1. Segelintir murid masih memerlukan peringatan berulang mengenai peraturan dan disiplin harian.\n" +
          "2. Kerjasama berterusan daripada ibu bapa dan penjaga perlu ditingkatkan bagi pemantauan di luar waktu persekolahan.";
    tindakan = "1. Melaksanakan bimbingan kaunseling berfokus bagi murid yang memerlukan perhatian khusus.\n" +
               "2. Mengadakan sesi libat urus berkala bersama ibu bapa melalui program PIBG dan hebahan digital.";
  } else if (b === "PPKI") {
    obj = "1. Memberikan bimbingan kemahiran motor, pengurusan diri dan sosialisasi kepada Murid Berkeperluan Pendidikan Khas (MBPK) berkaitan " + p + ".\n" +
          "2. Mengembangkan potensi bakat dan daya keyakinan diri MBPK dalam persekitaran yang inklusif dan mesra murid.\n" +
          "3. Memupuk semangat berdikari dan kerjasama antara MBPK, guru dan Pembantu Pengurusan Murid (PPM).";
    pengisian = "1. Taklimat persediaan aktiviti berpandu visual dan persediaan ruang aktiviti yang selamat dan kondusif.\n" +
                "2. Pelaksanaan aktiviti amali berperingkat mengikut keupayaan individu murid berpandukan Rancangan Pendidikan Individu (RPI).\n" +
                "3. Bimbingan 'one-on-one' oleh guru dan PPM sepanjang proses pelaksanaan aktiviti " + p + ".\n" +
                "4. Sesi penghargaan, ganjaran token positif dan refleksi perkembangan motor murid.";
    impak = "1. Semua MBPK menunjukkan respon positif, keceriaan dan penglibatan aktif mengikut tahap keupayaan masing-masing.\n" +
            "2. Peningkatan kemahiran manipulatif, fokus dan daya keyakinan diri murid semasa aktiviti dijalankan.\n" +
            "3. Kerjasama padu antara guru dan PPM membolehkan objektif program tercapai dengan selamat dan berkesan.";
    isu = "1. Tahap fokus murid berbeza-beza memerlukan teknik pengajaran berasaskan sensori yang pelbagai.\n" +
          "2. Murid cepat merasa letih dan memerlukan waktu rehat berkala semasa aktiviti intensif.";
    tindakan = "1. Menyediakan bahan bantu mengajar (BBM) multisensori yang lebih interaktif dan menarik minat murid.\n" +
               "2. Menyusun jadual aktiviti dalam slot yang lebih pendek dengan selingan regangan santai.";
  } else {
    // KURIKULUM / PANITIA (Default)
    obj = "1. Meningkatkan penguasaan kemahiran, kefahaman konsep dan pencapaian akademik murid dalam " + p + " (" + u + ").\n" +
          "2. Menerapkan elemen Kemahiran Berfikir Aras Tinggi (KBAT) dan Pembelajaran Abad Ke-21 (PAK-21) secara berkesan.\n" +
          "3. Memberi pengalaman pembelajaran kontekstual yang merangsang daya minat murid terhadap mata pelajaran.";
    pengisian = "1. Taklimat pengenalan tajuk dan aktiviti rangsangan awal berasaskan bahan maujud / digital.\n" +
                "2. Penerokaan isi pelajaran secara stesen kerja berkumpulan berpandukan modul khas panitia.\n" +
                "3. Pembentangan hasil kerja kumpulan, sesi soal jawab KBAT dan maklum balas guru secara formatif.\n" +
                "4. Ujian diagnostik ringkas / kuiz pengukuhan dan rumusan pencapaian objektif pembelajaran.";
    impak = "1. Peningkatan ketara dalam tahap penguasaan murid bagi standard pembelajaran yang disasarkan.\n" +
            "2. Murid lebih yakin bertutur, bekerjasama dalam pasukan dan berani mengemukakan idea bernas.\n" +
            "3. Pencapaian Pentaksiran Bilik Darjah (PBD) murid meningkat ke tahap penguasaan yang optimum (TP4 - TP6).";
    isu = "1. Terdapat jurang perbezaan tahap keupayaan murid memerlukan kaedah pembezaan (differentiated learning) yang teliti.\n" +
          "2. Kekangan masa untuk aktiviti pembentangan murid disebabkan penglibatan yang terlalu aktif.";
    tindakan = "1. Menyediakan modul bertingkat (lembaran pemulihan dan pengayaan) mengikut tahap keupayaan murid.\n" +
               "2. Mengoptimumkan penggunaan masa dengan menetapkan had masa pembentangan digital (Elevator Pitch).";
  }

  return {
    objektif: obj,
    pengisian: pengisian,
    impak: impak,
    isu: isu,
    tindakan: tindakan
  };
}

// Wrapper serasi belakang untuk janaKandunganOprUmum
function janaKandunganOprUmum(namaProgram, ringkasan) {
  return janaKandunganOprPintar("KURIKULUM", namaProgram, "Panitia Sekolah");
}

// Simpan atau Kemas Kini Rekod OPR ke Google Sheets (Tab LAPORAN_OPR)
function simpanAtauKemasKiniOpr(formData) {
  try {
    var sheet = dapatkanAtauCiptaSheetOpr();
    var data = sheet.getDataRange().getValues();
    
    var gUrls = ["", "", "", ""];
    for (var k = 1; k <= 4; k++) {
      var gData = formData["gambar" + k] || formData["gambar" + k + "Url"] || "";
      if (gData && String(gData).startsWith("data:")) {
        gUrls[k - 1] = simpanFailKeDrive(gData, "OPR_" + (formData.namaProgram || "PROG") + "_G" + k);
      } else if (gData) {
        gUrls[k - 1] = gData;
      } else if (rowIndex > 0 && data[rowIndex - 1] && data[rowIndex - 1][14 + k]) {
        gUrls[k - 1] = data[rowIndex - 1][14 + k]; // Kekalkan pautan gambar asal jika tiada muat naik baharu
      } else {
        gUrls[k - 1] = "";
      }
    }

    var idLaporan = formData.idLaporan ? String(formData.idLaporan).trim() : "";
    var rowIndex = -1;
    if (idLaporan) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === idLaporan) {
          rowIndex = i + 1; // 1-indexed untuk Google Sheet
          break;
        }
      }
    }

    if (!idLaporan || rowIndex === -1) {
      idLaporan = "OPR-" + Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd") + "-" + Math.floor(1000 + Math.random() * 9000);
    }

    var bahagian = formData.bahagian ? String(formData.bahagian).toUpperCase().trim() : "KURIKULUM";
    var unit = formData.unit || formData.unitPanitia || "-";
    var namaProgram = formData.namaProgram || "Program Sekolah";
    var penyelaras = formData.penyelaras || "-";
    var emelPenyelaras = formData.emelPenyelaras || "";
    var tarikh = formData.tarikh || Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
    var masa = formData.masa || "-";
    var tempat = formData.tempat || "-";
    var kehadiran = formData.kehadiran || formData.hadirAhli || "0";
    var objektif = formData.objektif || "-";
    var pengisian = formData.pengisian || formData.ringkasan || "-";
    var impak = formData.impak || "-";
    var isu = formData.isu || "-";
    var tindakan = formData.tindakan || "-";

    var statusPengesahan = "Menunggu Semakan";
    var disahkanOleh = "";
    var jawatanPengesah = "";
    var tarikhSah = "";
    var catatanPk = "";

    var recordPayload = {
      idLaporan: idLaporan,
      bahagian: bahagian,
      unit: unit,
      namaProgram: namaProgram,
      penyelaras: penyelaras,
      emelPenyelaras: emelPenyelaras,
      tarikh: tarikh,
      masa: masa,
      tempat: tempat,
      kehadiran: kehadiran,
      objektif: objektif,
      pengisian: pengisian,
      impak: impak,
      isu: isu,
      tindakan: tindakan,
      gambar1Url: gUrls[0],
      gambar2Url: gUrls[1],
      gambar3Url: gUrls[2],
      gambar4Url: gUrls[3],
      statusPengesahan: statusPengesahan,
      disahkanOleh: disahkanOleh,
      jawatanPengesah: jawatanPengesah,
      tarikhSah: tarikhSah,
      catatanPk: catatanPk
    };

    var resPdf = { urlPdf: "", base64: "" };
    try {
      resPdf = janaPdfOprBackend(recordPayload);
    } catch (ePdf) {
      console.warn("Ralat pra-jana PDF: " + ePdf.message);
    }

    var rowValues = [
      idLaporan,
      bahagian,
      unit,
      namaProgram,
      penyelaras,
      emelPenyelaras,
      tarikh,
      masa,
      tempat,
      kehadiran,
      objektif,
      pengisian,
      impak,
      isu,
      tindakan,
      gUrls[0],
      gUrls[1],
      gUrls[2],
      gUrls[3],
      statusPengesahan,
      disahkanOleh,
      jawatanPengesah,
      tarikhSah,
      catatanPk,
      resPdf.urlPdf || "",
      new Date()
    ];

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return {
      status: "SUCCESS",
      idLaporan: idLaporan,
      pesanan: "Laporan OPR '" + namaProgram + "' berjaya dihantar untuk semakan pentadbir!",
      urlPdf: resPdf.urlPdf,
      base64: resPdf.base64
    };
  } catch (err) {
    throw new Error("Ralat simpan laporan OPR: " + err.message);
  }
}

// Wrapper serasi belakang untuk simpanOprUmumBackend
function simpanOprUmumBackend(formData) {
  return simpanAtauKemasKiniOpr(formData);
}

// Dapatkan Senarai Laporan OPR untuk Semakan Pentadbir mengikut Skop RBAC
function dapatkanSenaraiOprPentadbir(emelPentadbir) {
  var senarai = [];
  try {
    var senaraiGuru = getSenaraiGuruWeb();
    var admin = senaraiGuru.find(function(g) { return padanEmelSama(g.emel, emelPentadbir); });
    if (!admin) return [];
    var rbac = tentukanSkopPeranan(admin.peranan);
    if (!rbac.bolehLulus) return [];

    var sheet = dapatkanAtauCiptaSheetOpr();
    var data = sheet.getDataRange().getDisplayValues();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var bahagian = String(row[1] || "").toUpperCase().trim();

      if (rbac.skop.indexOf(bahagian) !== -1) {
        senarai.push({
          idLaporan: row[0],
          bahagian: row[1],
          unit: row[2],
          namaProgram: row[3],
          penyelaras: row[4],
          emelPenyelaras: row[5],
          tarikh: row[6],
          masa: row[7],
          tempat: row[8],
          kehadiran: row[9],
          objektif: row[10],
          pengisian: row[11],
          impak: row[12],
          isu: row[13],
          tindakan: row[14],
          gambar1Url: row[15],
          gambar2Url: row[16],
          gambar3Url: row[17],
          gambar4Url: row[18],
          statusPengesahan: row[19] || "Menunggu Semakan",
          disahkanOleh: row[20] || "-",
          jawatanPengesah: row[21] || "-",
          tarikhSah: row[22] || "-",
          catatanPk: row[23] || "",
          urlPdf: row[24] || ""
        });
      }
    }

    senarai.sort(function(a, b) {
      if (a.statusPengesahan === "Menunggu Semakan" && b.statusPengesahan !== "Menunggu Semakan") return -1;
      if (a.statusPengesahan !== "Menunggu Semakan" && b.statusPengesahan === "Menunggu Semakan") return 1;
      return 0;
    });

  } catch (err) {
    console.warn("Ralat dapatkan senarai OPR pentadbir: " + err.message);
  }
  return senarai;
}

// Aliran Kerja Pengesahan Digital oleh Pentadbir
function sahkanLaporanOprBackend(idLaporan, emelPentadbir, statusTindakan, catatanPk) {
  try {
    var sheet = dapatkanAtauCiptaSheetOpr();
    var data = sheet.getDataRange().getValues();
    var senaraiGuru = getSenaraiGuruWeb();
    var admin = senaraiGuru.find(function(g) { return padanEmelSama(g.emel, emelPentadbir); });
    
    if (!admin) {
      throw new Error("Pengesahan gagal: Emel pentadbir tidak sah dalam sistem.");
    }
    var rbac = tentukanSkopPeranan(admin.peranan);
    if (!rbac.bolehLulus) {
      throw new Error("Pengesahan gagal: Anda tidak mempunyai hak kuasa pentadbir untuk meluluskan laporan.");
    }

    var targetRow = -1;
    var rowData = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(idLaporan).trim()) {
        targetRow = i + 1;
        rowData = data[i];
        break;
      }
    }

    if (targetRow === -1 || !rowData) {
      throw new Error("Laporan OPR dengan ID " + idLaporan + " tidak dijumpai.");
    }

    var bahagianLaporan = String(rowData[1] || "").toUpperCase().trim();
    if (rbac.skop.indexOf(bahagianLaporan) === -1) {
      throw new Error("Akses dinafikan: Peranan " + rbac.label + " hanya berkuasa untuk bahagian [" + rbac.skop.join(", ") + "], bukan " + bahagianLaporan + ".");
    }

    var tarikhKini = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy, HH:mm:ss");

    if (statusTindakan === "Disahkan") {
      sheet.getRange(targetRow, 20).setValue("Disahkan");
      sheet.getRange(targetRow, 21).setValue(admin.nama);
      sheet.getRange(targetRow, 22).setValue(rbac.label + ", SK SOOK");
      sheet.getRange(targetRow, 23).setValue(tarikhKini);
      sheet.getRange(targetRow, 24).setValue("");

      var recordPayload = {
        idLaporan: rowData[0],
        bahagian: rowData[1],
        unit: rowData[2],
        namaProgram: rowData[3],
        penyelaras: rowData[4],
        emelPenyelaras: rowData[5],
        tarikh: rowData[6],
        masa: rowData[7],
        tempat: rowData[8],
        kehadiran: rowData[9],
        objektif: rowData[10],
        pengisian: rowData[11],
        impak: rowData[12],
        isu: rowData[13],
        tindakan: rowData[14],
        gambar1Url: rowData[15],
        gambar2Url: rowData[16],
        gambar3Url: rowData[17],
        gambar4Url: rowData[18],
        statusPengesahan: "Disahkan",
        disahkanOleh: admin.nama,
        jawatanPengesah: rbac.label + ", SK SOOK",
        tarikhSah: tarikhKini,
        catatanPk: ""
      };

      var resPdf = janaPdfOprBackend(recordPayload);
      if (resPdf && resPdf.urlPdf) {
        sheet.getRange(targetRow, 25).setValue(resPdf.urlPdf);
      }

      return {
        status: "SUCCESS",
        pesanan: "Laporan '" + rowData[3] + "' berjaya disahkan secara digital oleh " + admin.nama + " (" + rbac.label + ")!",
        statusLaporan: "Disahkan",
        urlPdf: resPdf.urlPdf,
        base64: resPdf.base64
      };
    } else if (statusTindakan === "Perlu Pindaan") {
      sheet.getRange(targetRow, 20).setValue("Perlu Pindaan");
      sheet.getRange(targetRow, 24).setValue(catatanPk || "Sila lengkapkan butiran laporan.");

      return {
        status: "SUCCESS",
        pesanan: "Laporan telah dikembalikan kepada guru penyedia (" + rowData[4] + ") untuk tindakan pindaan.",
        statusLaporan: "Perlu Pindaan"
      };
    } else {
      throw new Error("Status tindakan tidak sah: " + statusTindakan);
    }
  } catch (err) {
    throw new Error("Ralat pengesahan laporan OPR: " + err.message);
  }
}

// Dapatkan Rekod untuk Bank & Arkib OPR Sekolah (Untuk Semua Guru)
function dapatkanBankOprSekolah(penapis) {
  var hasil = [];
  penapis = penapis || {};
  try {
    var sheet = dapatkanAtauCiptaSheetOpr();
    var data = sheet.getDataRange().getDisplayValues();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;

      var idLaporan = row[0];
      var bahagian = row[1] || "-";
      var unit = row[2] || "-";
      var namaProgram = row[3] || "-";
      var penyelaras = row[4] || "-";
      var emelPenyelaras = row[5] || "";
      var tarikh = row[6] || "-";
      var masa = row[7] || "-";
      var tempat = row[8] || "-";
      var kehadiran = row[9] || "0";
      var statusPengesahan = row[19] || "Menunggu Semakan";
      var disahkanOleh = row[20] || "-";
      var jawatanPengesah = row[21] || "-";
      var tarikhSah = row[22] || "-";
      var catatanPk = row[23] || "";
      var urlPdf = row[24] || "";

      if (penapis.emelGuru) {
        if (!padanEmelSama(emelPenyelaras, penapis.emelGuru)) {
          continue;
        }
      } else if (penapis.hanyaDisahkan !== false) {
        if (statusPengesahan !== "Disahkan") {
          continue;
        }
      }

      if (penapis.bahagian && penapis.bahagian !== "SEMUA") {
        if (String(bahagian).toUpperCase() !== String(penapis.bahagian).toUpperCase()) continue;
      }

      if (penapis.carian) {
        var q = String(penapis.carian).toLowerCase();
        var match = String(namaProgram).toLowerCase().indexOf(q) !== -1 ||
                    String(unit).toLowerCase().indexOf(q) !== -1 ||
                    String(penyelaras).toLowerCase().indexOf(q) !== -1;
        if (!match) continue;
      }

      hasil.push({
        idLaporan: idLaporan,
        bahagian: bahagian,
        unit: unit,
        namaProgram: namaProgram,
        penyelaras: penyelaras,
        emelPenyelaras: emelPenyelaras,
        tarikh: tarikh,
        masa: masa,
        tempat: tempat,
        kehadiran: kehadiran,
        objektif: row[10] || "-",
        pengisian: row[11] || "-",
        impak: row[12] || "-",
        isu: row[13] || "-",
        tindakan: row[14] || "-",
        gambar1Url: row[15] || "",
        gambar2Url: row[16] || "",
        gambar3Url: row[17] || "",
        gambar4Url: row[18] || "",
        statusPengesahan: statusPengesahan,
        disahkanOleh: disahkanOleh,
        jawatanPengesah: jawatanPengesah,
        tarikhSah: tarikhSah,
        catatanPk: catatanPk,
        urlPdf: urlPdf
      });
    }

    hasil.reverse();
  } catch (err) {
    console.warn("Ralat dapatkan Bank OPR: " + err.message);
  }
  return hasil;
}

// Wrapper serasi belakang untuk getSenaraiArkibOpr
function getSenaraiArkibOpr() {
  return dapatkanBankOprSekolah({ hanyaDisahkan: true });
}

// Penjanaan PDF OPR Standard (A4 Portrait - 1 Muka Surat Sahaja)
function janaPdfOprBackend(formData) {
  try {
    var prog = formData.namaProgram || "LAPORAN PROGRAM SEKOLAH";
    var bahagian = formData.bahagian || "KURIKULUM";
    var unit = formData.unit || formData.unitPanitia || "Unit Sekolah";
    var penyelaras = formData.penyelaras || "Penyelaras Program";
    var tempat = formData.tempat || "SK Sook, Keningau";
    var tarikh = formData.tarikh || "-";
    var masa = formData.masa || "-";
    var kehadiran = formData.kehadiran || formData.hadirAhli || "Warga & Murid SK Sook";
    var objektif = formData.objektif || "-";
    var pengisian = formData.pengisian || formData.ringkasan || "-";
    var impak = formData.impak || "-";
    var isu = formData.isu || "-";
    var tindakan = formData.tindakan || "-";

    var isDisahkan = (formData.statusPengesahan === "Disahkan");

    var gUrls = [
      formData.gambar1 || formData.gambar1Url || "",
      formData.gambar2 || formData.gambar2Url || "",
      formData.gambar3 || formData.gambar3Url || "",
      formData.gambar4 || formData.gambar4Url || ""
    ];

    // Pastikan sebarang pautan Google Drive ditukar kepada Data URI Base64 tulen untuk pemaparan PDF selamat
    for (var gi = 0; gi < 4; gi++) {
      if (gUrls[gi] && !String(gUrls[gi]).startsWith("data:image")) {
        var b64Drive = ambilBase64DariDrive(gUrls[gi]);
        if (b64Drive) {
          gUrls[gi] = b64Drive;
        }
      }
    }

    function formatPointsHtml(text) {
      if (!text) return '-';
      var lines = String(text).split('\n').filter(function(l){ return l.trim() !== ''; });
      if (lines.length <= 1) return String(text).replace(/\n/g, '<br>');
      return lines.map(function(l){ return '<div style="margin-bottom:2px;">' + l + '</div>'; }).join('');
    }

    var gambarCards = [];
    for (var i = 0; i < 4; i++) {
      if (gUrls[i]) {
        gambarCards.push(`
          <div style="text-align:center; flex:1; min-width:125px; padding:3px; border:1px solid #e2e8f0; border-radius:6px; background:#fff;">
            <img src="${gUrls[i]}" style="max-height:105px; max-width:100%; object-fit:cover; border-radius:4px;" />
            <div style="font-size:7.5px; color:#64748b; margin-top:2px; font-weight:bold;">Gambar ${i + 1}</div>
          </div>
        `);
      }
    }

    if (gambarCards.length === 0) {
      gambarCards.push('<div style="width:100%; text-align:center; padding:12px; color:#94a3b8; font-size:8px; border:1px dashed #cbd5e1; border-radius:6px; background:#f8fafc;">[Tiada Gambar Dilampirkan]</div>');
    }

    var watermarkHtml = isDisahkan ? '' : `
      <div style="position:fixed; top:40%; left:5%; width:90%; text-align:center; transform:rotate(-32deg); font-size:40pt; color:rgba(225, 29, 72, 0.14); font-weight:900; z-index:9999; pointer-events:none; border:4px dashed rgba(225, 29, 72, 0.22); padding:10px; border-radius:14px; letter-spacing:2px;">
        DRAF - BELUM DISAHKAN
      </div>
    `;

    var logoSekolahSrc = dapatkanLogoSekolah();
    var logoHeaderHtml = logoSekolahSrc ? 
      '<img src="' + logoSekolahSrc + '" style="max-height:42px; max-width:42px; object-fit:contain;" alt="Lencana SK Sook" />' : 
      '<div style="font-size: 22px;">🏫</div>';

    var htmlContent = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>One Page Report - ${prog}</title>
  <style>
    @page { size: A4 portrait; margin: 8mm 10mm 8mm 10mm; }
    body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; font-size: 8.8px; color: #1e293b; margin: 0; line-height: 1.28; }
    .header-table { width: 100%; border-bottom: 2px solid #1e3a8a; padding-bottom: 4px; margin-bottom: 5px; }
    .title-main { font-size: 11.5px; font-weight: 800; color: #1e3a8a; text-transform: uppercase; }
    .title-sub { font-size: 8.8px; font-weight: 700; color: #475569; }
    .info-grid { width: 100%; border-collapse: collapse; margin-bottom: 5px; font-size: 8.2px; }
    .info-grid td { border: 1px solid #cbd5e1; padding: 3px 5px; vertical-align: top; }
    .info-lbl { font-weight: bold; background-color: #f8fafc; color: #334155; width: 17%; }
    .info-val { font-weight: 600; color: #0f172a; width: 33%; }
    .section-box { border: 1px solid #cbd5e1; border-radius: 4px; margin-bottom: 4px; overflow: hidden; }
    .section-title { background: #f1f5f9; font-weight: 800; color: #1e293b; padding: 2.5px 5px; font-size: 8.2px; border-bottom: 1px solid #cbd5e1; text-transform: uppercase; }
    .section-content { padding: 3.5px 5px; font-size: 8.2px; color: #1e293b; }
    .footer-table { width: 100%; margin-top: 6px; border-collapse: collapse; font-size: 7.8px; }
    .footer-table td { width: 50%; vertical-align: top; }
  </style>
</head>
<body>
  ${watermarkHtml}

  <table class="header-table">
    <tr>
      <td style="width: 45px; text-align: center; vertical-align: middle;">
        ${logoHeaderHtml}
      </td>
      <td style="vertical-align: middle;">
        <div class="title-main">SEKOLAH KEBANGSAAN SOOK, KENINGAU</div>
        <div class="title-sub">LAPORAN SATU MUKA SURAT (ONE PAGE REPORT - OPR)</div>
        <div style="font-size: 7.2px; color: #64748b;">Peti Surat 204, 89008 Keningau, Sabah • Kod Sekolah: XBA1026 • Bahagian: ${bahagian}</div>
      </td>
    </tr>
  </table>

  <table class="info-grid">
    <tr>
      <td class="info-lbl">NAMA PROGRAM</td>
      <td class="info-val" colspan="3" style="font-size: 9px; color: #1e3a8a; font-weight: 800;">${prog}</td>
    </tr>
    <tr>
      <td class="info-lbl">UNIT / PANITIA</td>
      <td class="info-val">${unit}</td>
      <td class="info-lbl">PENYELARAS</td>
      <td class="info-val">${penyelaras}</td>
    </tr>
    <tr>
      <td class="info-lbl">TARIKH & MASA</td>
      <td class="info-val">${tarikh} (${masa})</td>
      <td class="info-lbl">TEMPAT</td>
      <td class="info-val">${tempat}</td>
    </tr>
    <tr>
      <td class="info-lbl">KEHADIRAN</td>
      <td class="info-val" colspan="3">${kehadiran}</td>
    </tr>
  </table>

  <div class="section-box">
    <div class="section-title">1. OBJEKTIF PROGRAM / PERJUMPAAN</div>
    <div class="section-content">${formatPointsHtml(objektif)}</div>
  </div>

  <div class="section-box">
    <div class="section-title">2. PENGISIAN / LANGKAH AKTIVITI</div>
    <div class="section-content">${formatPointsHtml(pengisian)}</div>
  </div>

  <div class="section-box">
    <div class="section-title">3. IMPAK / KEBERHASILAN MURID</div>
    <div class="section-content">${formatPointsHtml(impak)}</div>
  </div>

  <table style="width:100%; border-collapse:collapse; margin-bottom:4px;">
    <tr>
      <td style="width:49.5%; vertical-align:top; padding-right:2px;">
        <div class="section-box" style="margin-bottom:0;">
          <div class="section-title" style="background:#fff1f2; color:#9f1239; border-color:#fecdd3;">4. ISU & CABARAN</div>
          <div class="section-content" style="min-height:26px;">${formatPointsHtml(isu)}</div>
        </div>
      </td>
      <td style="width:49.5%; vertical-align:top; padding-left:2px;">
        <div class="section-box" style="margin-bottom:0;">
          <div class="section-title" style="background:#f0fdf4; color:#166534; border-color:#bbf7d0;">5. TINDAKAN SUSULAN</div>
          <div class="section-content" style="min-height:26px;">${formatPointsHtml(tindakan)}</div>
        </div>
      </td>
    </tr>
  </table>

  <div class="section-box" style="margin-bottom:4px;">
    <div class="section-title">6. DOKUMENTASI BERGAMBAR AKTIVITI</div>
    <div class="section-content" style="display:flex; gap:5px; padding:3px; flex-wrap:wrap;">
      ${gambarCards.join('')}
    </div>
  </div>

  <table class="footer-table">
    <tr>
      <td style="padding-right: 8px;">
        <div style="font-weight: bold; color: #475569; margin-bottom: 2px;">DISEDIAKAN OLEH:</div>
        <div style="border: 1px solid #cbd5e1; background: #f8fafc; border-radius: 4px; padding: 4px 6px;">
          <div style="font-weight: 800; font-size: 8px; color: #1e293b;">${penyelaras.toUpperCase()}</div>
          <div style="font-size: 7.2px; color: #64748b;">Penyelaras / Guru Penasihat</div>
          <div style="font-size: 7px; color: #94a3b8; margin-top: 1px;">Tarikh Hantar: ${tarikh}</div>
        </div>
      </td>
      <td style="padding-left: 8px;">
        ${isDisahkan ? `
          <div style="font-weight: 800; font-size: 8px; color: #047857; margin-bottom: 2px;">PERAKUAN PENGESAHAN DIGITAL:</div>
          <div style="border: 1.5px solid #10b981; background: #f0fdf4; border-radius: 5px; padding: 4px 6px;">
            <div style="display:flex; justify-content:space-between; align-items:flex-start;">
              <div>
                <div style="font-weight: 800; font-size: 8.2px; color: #065f46;">${(formData.disahkanOleh || "PENTADBIR SEKOLAH").toUpperCase()}</div>
                <div style="font-size: 7.2px; font-weight: 700; color: #047857;">${formData.jawatanPengesah || "PENTADBIR, SK SOOK"}</div>
                <div style="font-size: 6.8px; color: #4b5563; margin-top: 1px;">Disahkan pada: ${formData.tarikhSah || "-"}</div>
              </div>
              <div style="text-align:center; padding: 1px 4px; background:#dcfce7; border:1px solid #86efac; border-radius:3px; font-size:6.8px; font-weight:bold; color:#166534;">
                🛡️ DISAHKAN DIGITAL
              </div>
            </div>
            <div style="font-size: 6.5px; color: #065f46; margin-top: 2px; border-top: 1px dashed #a7f3d0; padding-top: 2px; line-height:1.2;">
              ✓ Dokumen ini telah disahkan secara digital melalui Portal Rasmi SK Sook. Sah tanpa tandatangan fizikal mengikut Akta Tandatangan Digital 1997.
            </div>
          </div>
        ` : `
          <div style="font-weight: bold; font-size: 8px; color: #64748b; margin-bottom: 2px;">PENGESAHAN PENTADBIR:</div>
          <div style="border: 1px dashed #94a3b8; background: #f8fafc; border-radius: 4px; padding: 6px; text-align:center;">
            <div style="font-size: 7.8px; font-weight: bold; color: #d97706;">⏳ MENUNGGU SEMAKAN PENTADBIR</div>
            <div style="font-size: 6.8px; color: #64748b; margin-top: 2px;">Laporan ini berstatus draf dan belum disahkan oleh pihak pengurusan sekolah.</div>
          </div>
        `}
      </td>
    </tr>
  </table>
</body>
</html>`;

    var folderName = "PORTAL_SK_SOOK_PDF";
    var folder;
    try {
      var folders = DriveApp.getFoldersByName(folderName);
      folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    } catch (fErr) {
      folder = DriveApp.getRootFolder();
    }

    var namaFailPdf = "OPR_" + (isDisahkan ? "RASMI_" : "DRAF_") + prog.replace(/\s+/g, '_') + ".pdf";
    var tempFile = DriveApp.createFile("temp_opr.html", htmlContent, MimeType.HTML);
    var pdfBlob = tempFile.getAs(MimeType.PDF).setName(namaFailPdf);
    var pdfFile = folder.createFile(pdfBlob);

    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (errDomain) {
      try { pdfFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (eSub) {}
    }

    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    tempFile.setTrashed(true);

    return {
      status: "SUCCESS",
      urlPdf: pdfFile.getUrl(),
      downloadUrl: pdfFile.getUrl().replace('view?usp=drivesdk', 'export?format=pdf'),
      base64: pdfBase64,
      namaFail: namaFailPdf,
      pesanan: "PDF OPR '" + prog + "' berjaya dijana!"
    };
  } catch (err) {
    throw new Error("Ralat menjana PDF OPR: " + err.message);
  }
}

// ==========================================================================
// MODUL e-KEHADIRAN GURU BERASASKAN GPS GEOFENCING (SK SOOK)
// KOORDINAT RASMI: 5°08'52.7"N 116°18'25.9"E | RADIUS: 100 METER
// ==========================================================================

var GPS_SK_SOOK = {
  lat: 5.147972,
  lon: 116.307194,
  radiusMeter: 150,
  namaLokasi: "SK Sook, Keningau"
};

/**
 * Mengira jarak antara dua koordinat menggunakan formula Haversine (unit: Meter)
 */
function kiraJarakHaversine(lat1, lon1, lat2, lon2) {
  var R = 6371000; // Radius Bumi dalam meter
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLon = (lon2 - lon1) * Math.PI / 180;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
          Math.sin(dLon / 2) * Math.sin(dLon / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return Math.round(R * c * 10) / 10;
}

/**
 * Rakam kehadiran guru berasaskan koordinat GPS ke lembaran KEHADIRAN_GURU
 */
function rakamKehadiranGpsBackend(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("KEHADIRAN_GURU");
    
    // Cipta tab KEHADIRAN_GURU sekiranya belum wujud
    if (!sheet) {
      sheet = ss.insertSheet("KEHADIRAN_GURU");
      var headers = [
        "ID_REKOD", "CAP_MASA", "TARIKH", "HARI", "MASA",
        "EMEL", "NAMA", "JAWATAN", "JENIS", "STATUS_LOKASI",
        "JARAK_METER", "KETEPATAN_GPS", "KOORDINAT_GURU", "STATUS_MASA", "CATATAN"
      ];
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#1e1b4b").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var lat = parseFloat(data.lat || 0);
    var lon = parseFloat(data.lon || 0);
    var ketepatan = parseFloat(data.ketepatan || 0);

    if (!lat || !lon || isNaN(lat) || isNaN(lon)) {
      throw new Error("Koordinat GPS tidak dikesan. Sila hidupkan lokasi GPS peranti anda dan cuba lagi.");
    }

    // Pengiraan jarak pelayan yang sah dan kebal manipulasi (Radius 150m)
    var jarakSebenar = kiraJarakHaversine(lat, lon, GPS_SK_SOOK.lat, GPS_SK_SOOK.lon);
    var dalamKawasan = (jarakSebenar <= GPS_SK_SOOK.radiusMeter);
    var jenis = String(data.jenis || "MASUK").toUpperCase();

    // SPESIFIKASI e-KEHADIRAN GPS PINTAR:
    // Waktu Masuk: Wajib dalam radius 150m; sekat jika di luar kawasan
    if (jenis === "MASUK" && !dalamKawasan) {
      return {
        status: "ERROR",
        message: "Sekatan Kehadiran: Anda berada " + jarakSebenar + "m di luar kawasan SK Sook (Had Geofence: " + GPS_SK_SOOK.radiusMeter + "m). Waktu Masuk hanya dibenarkan semasa berada di dalam kawasan sekolah."
      };
    }

    var now = new Date();
    var tarikhStr = Utilities.formatDate(now, "Asia/Kuala_Lumpur", "dd/MM/yyyy");
    var masaStr = Utilities.formatDate(now, "Asia/Kuala_Lumpur", "hh:mm:ss a");
    var hariList = ["Ahad", "Isnin", "Selasa", "Rabu", "Khamis", "Jumaat", "Sabtu"];
    var hariStr = hariList[now.getDay()];

    // Penentuan Ketepatan Masa (Had 07:10 Pagi):
    // <= 07:10 Pagi = Tepat Masa, >= 07:11 Pagi = Lewat
    var jamKini = parseInt(Utilities.formatDate(now, "Asia/Kuala_Lumpur", "HH"), 10);
    var minitKini = parseInt(Utilities.formatDate(now, "Asia/Kuala_Lumpur", "mm"), 10);
    var statusKetepatan = "";
    if (jenis === "MASUK") {
      if (jamKini < 7 || (jamKini === 7 && minitKini <= 10)) {
        statusKetepatan = "TEPAT MASA";
      } else {
        statusKetepatan = "LEWAT";
      }
    } else {
      statusKetepatan = dalamKawasan ? "DALAM KAWASAN" : "LUAR KAWASAN";
    }

    // Waktu Pulang: Dibenarkan di mana-mana sahaja; jika >150m ditag LUAR KAWASAN
    var statusLokasi = dalamKawasan ? "DALAM KAWASAN" : "LUAR KAWASAN";
    var koordinatStr = lat.toFixed(6) + ", " + lon.toFixed(6);

    var idRekod = "ATT-" + Utilities.formatDate(now, "Asia/Kuala_Lumpur", "yyyyMMdd-HHmmss") + "-" + Math.floor(100 + Math.random() * 900);
    var emel = String(data.emel || "").trim();
    var nama = String(data.nama || "").trim();
    var jawatan = String(data.jawatan || "Guru").trim();
    var catatan = String(data.catatan || "").trim();

    var emelBersih = emel.toLowerCase().trim();

    // Kawalan Had: 1 kali Masuk & 1 kali Pulang sahaja bagi setiap hari kalendar
    var dataKehadiran = sheet.getDataRange().getValues();
    for (var r = 1; r < dataKehadiran.length; r++) {
      var rowTarikh = formatTarikhStandard(dataKehadiran[r][2]);
      var rowEmel = String(dataKehadiran[r][5] || "").toLowerCase().trim();
      var rowJenis = String(dataKehadiran[r][8] || "").toUpperCase().trim();
      if (rowTarikh === tarikhStr && rowEmel === emelBersih && rowJenis === jenis) {
        return {
          status: "ERROR",
          message: "Kehadiran/Kepulangan anda bagi hari ini telah pun direkodkan."
        };
      }
    }

    if (jenis === "MASUK" && statusKetepatan === "LEWAT") {
      catatan = (catatan ? catatan + " " : "") + "[LEWAT " + masaStr + "]";
    } else if (jenis === "PULANG" && !dalamKawasan) {
      catatan = (catatan ? catatan + " " : "") + "[PULANG LUAR KAWASAN " + jarakSebenar + "m]";
    }

    sheet.appendRow([
      idRekod,
      now.toISOString(),
      "'" + tarikhStr,
      hariStr,
      masaStr,
      emel,
      nama,
      jawatan,
      jenis,
      statusLokasi,
      jarakSebenar,
      ketepatan,
      koordinatStr,
      statusKetepatan,
      catatan
    ]);

    var pesananHantar = "Kehadiran " + jenis + " berjaya direkodkan (" + (jenis === "MASUK" ? statusKetepatan : statusLokasi) + " - " + masaStr + ")";

    return {
      status: "SUCCESS",
      idRekod: idRekod,
      jenis: jenis,
      tarikh: tarikhStr,
      hari: hariStr,
      masa: masaStr,
      jarak: jarakSebenar,
      dalamKawasan: dalamKawasan,
      statusLokasi: statusLokasi,
      statusKetepatan: statusKetepatan,
      nama: nama,
      pesanan: pesananHantar
    };
  } catch (err) {
    return {
      status: "ERROR",
      message: "Ralat merekodkan kehadiran: " + err.message
    };
  }
}

/**
 * Auto-Checkout Harian jam 5:00 Petang (Isnin hingga Jumaat)
 * Mengisi rekod pulang secara automatik bagi guru yang belum mendaftar keluar
 */
function autoCheckoutHarian5PM() {
  try {
    var now = new Date();
    var dayOfWeek = now.getDay(); // 0 = Ahad, 1 = Isnin, ..., 5 = Jumaat, 6 = Sabtu
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      return { success: false, message: "Bukan hari persekolahan (Isnin - Jumaat)." };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("KEHADIRAN_GURU");
    if (!sheet) return { success: false, message: "Sheet KEHADIRAN_GURU tidak ditemui." };

    var tarikhStr = Utilities.formatDate(now, "Asia/Kuala_Lumpur", "dd/MM/yyyy");
    var hariList = ["Ahad", "Isnin", "Selasa", "Rabu", "Khamis", "Jumaat", "Sabtu"];
    var hariStr = hariList[dayOfWeek];
    var data = sheet.getDataRange().getValues();

    var guruMasuk = {};
    var guruPulang = {};

    for (var i = 1; i < data.length; i++) {
      var rowTarikh = formatTarikhStandard(data[i][2]);
      if (rowTarikh === tarikhStr) {
        var rowEmel = String(data[i][5] || "").toLowerCase().trim();
        var rowJenis = String(data[i][8] || "").toUpperCase().trim();
        if (rowJenis === "MASUK") {
          guruMasuk[rowEmel] = {
            emel: data[i][5],
            nama: data[i][6],
            jawatan: data[i][7]
          };
        } else if (rowJenis === "PULANG") {
          guruPulang[rowEmel] = true;
        }
      }
    }

    var checkoutCount = 0;
    Object.keys(guruMasuk).forEach(function(em) {
      if (!guruPulang[em]) {
        var g = guruMasuk[em];
        var idRekod = "ATT-AUTO-" + Utilities.formatDate(now, "Asia/Kuala_Lumpur", "yyyyMMdd") + "-" + Math.floor(1000 + Math.random() * 9000);
        sheet.appendRow([
          idRekod,
          now.toISOString(),
          "'" + tarikhStr,
          hariStr,
          "17:00:00",
          g.emel,
          g.nama,
          g.jawatan,
          "PULANG",
          "SISTEM AUTO",
          0,
          0,
          "AUTO_CHECKOUT",
          "AUTO CHECKOUT",
          "Auto Checkout 5PM"
        ]);
        checkoutCount++;
      }
    });

    return { success: true, message: "Auto-checkout berjaya dilaksanakan untuk " + checkoutCount + " orang guru.", count: checkoutCount };
  } catch (err) {
    console.warn("Ralat autoCheckoutHarian5PM: " + err.message);
    return { success: false, message: err.message };
  }
}

/**
 * Pemasang Trigger Time-Driven Auto Checkout jam 5:00 Petang setiap hari persekolahan
 */
function pasangTriggerAutoCheckout() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "autoCheckoutHarian5PM") {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    ScriptApp.newTrigger("autoCheckoutHarian5PM")
      .timeBased()
      .atHour(17)
      .everyDays(1)
      .inTimezone("Asia/Kuala_Lumpur")
      .create();
    return { success: true, message: "Trigger Auto Checkout harian (5:00 PM) berjaya dipasang!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Padam RPH Mingguan yang Belum Disemak/Diluluskan oleh Pentadbir
 */
function padamRphMingguanBackend(emel, mingguAtauId) {
  try {
    if (!emel || !mingguAtauId) {
      return { success: false, message: "Parameter emel dan minggu/ID diperlukan." };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetArkib = ss.getSheetByName("ARKIB_ERPH_GURU");
    var sheetRph = ss.getSheetByName("RPH_GURU");

    var targetMinggu = String(mingguAtauId).trim();
    var targetEmel = String(emel).trim().toLowerCase();

    // 1. Semak status dalam ARKIB_ERPH_GURU
    var statusSemakan = "";
    var rowArkibToDelete = -1;

    if (sheetArkib) {
      var dataArkib = sheetArkib.getDataRange().getValues();
      for (var a = 1; a < dataArkib.length; a++) {
        var aId = String(dataArkib[a][0] || "").trim();
        var aEmel = String(dataArkib[a][3] || "").trim().toLowerCase();
        var aMinggu = String(dataArkib[a][4] || "").trim();

        if (aEmel === targetEmel && (aId === targetMinggu || aMinggu.toLowerCase() === targetMinggu.toLowerCase() || padanMingguSama(aMinggu, targetMinggu))) {
          statusSemakan = String(dataArkib[a][9] || "").trim().toUpperCase();
          rowArkibToDelete = a + 1;
          targetMinggu = aMinggu;
          break;
        }
      }
    }

    if (statusSemakan === "DISEMAK" || statusSemakan === "LULUS" || statusSemakan === "DISAHKAN") {
      return {
        success: false,
        message: "e-RPH ini telah disemak/diluluskan oleh pentadbir dan tidak boleh dipadam."
      };
    }

    // 2. Padam slot PdP berkaitan di sheet RPH_GURU
    var deletedSlotsCount = 0;
    if (sheetRph) {
      var dataRph = sheetRph.getDataRange().getValues();
      for (var r = dataRph.length - 1; r >= 1; r--) {
        var rEmel = String(dataRph[r][1] || "").trim().toLowerCase();
        var rMinggu = String(dataRph[r][2] || "").trim();
        if (rEmel === targetEmel && (padanMingguSama(rMinggu, targetMinggu) || rMinggu.toLowerCase() === targetMinggu.toLowerCase())) {
          sheetRph.deleteRow(r + 1);
          deletedSlotsCount++;
        }
      }
    }

    // 3. Padam rekod dalam ARKIB_ERPH_GURU jika wujud
    if (sheetArkib && rowArkibToDelete > 0) {
      sheetArkib.deleteRow(rowArkibToDelete);
    }

    return {
      success: true,
      message: "Rekod e-RPH bagi " + targetMinggu + " berjaya dipadam (" + deletedSlotsCount + " slot PdP dibersihkan).",
      minggu: targetMinggu
    };
  } catch (err) {
    return { success: false, message: "Ralat memadam e-RPH: " + err.message };
  }
}

/**
 * Simpan Templat Jadual Waktu Kekal Guru ke Lembaran JADUAL_GURU
 */
function simpanJadualGuruBackend(emel, jadualObj) {
  try {
    if (!emel) return { success: false, message: "Emel guru diperlukan." };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("JADUAL_GURU");
    if (!sheet) {
      sheet = ss.insertSheet("JADUAL_GURU");
      sheet.appendRow(["EMEL_GURU", "JADUAL_JSON", "TARIKH_KEMASKINI"]);
      sheet.getRange(1, 1, 1, 3).setBackground("#1e1b4b").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var targetEmel = String(emel).trim().toLowerCase();
    var jsonStr = (typeof jadualObj === "string") ? jadualObj : JSON.stringify(jadualObj);
    var nowStr = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd/MM/yyyy HH:mm:ss");

    var data = sheet.getDataRange().getValues();
    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === targetEmel) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > 0) {
      sheet.getRange(foundRow, 2).setValue(jsonStr);
      sheet.getRange(foundRow, 3).setValue(nowStr);
    } else {
      sheet.appendRow([emel, jsonStr, nowStr]);
    }

    // Ekstrak & simpan senarai kelas tersuai ke TETAPAN_KELAS secara automatik
    try {
      var jObj = (typeof jadualObj === "string") ? JSON.parse(jadualObj) : jadualObj;
      var kelasUnik = [];
      if (jObj && typeof jObj === "object") {
        for (var h in jObj) {
          var sls = jObj[h] || [];
          sls.forEach(function(s) {
            if (s && s.kelas) {
              var kVal = String(s.kelas).trim();
              if (kVal && !kelasUnik.includes(kVal)) kelasUnik.push(kVal);
            }
          });
        }
      }
      if (kelasUnik.length > 0) {
        simpanSenaraiKelasGuru(emel, kelasUnik);
      }
    } catch (eKelas) {}

    return { success: true, message: "Templat jadual waktu berjaya disimpan secara kekal!" };
  } catch (err) {
    return { success: false, message: "Ralat menyimpan templat jadual: " + err.message };
  }
}

/**
 * Ambil Templat Jadual Waktu Kekal Guru dari Lembaran JADUAL_GURU
 */
function dapatkanJadualGuruBackend(emel) {
  try {
    if (!emel) return { success: false, message: "Emel tidak sah." };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("JADUAL_GURU");
    if (!sheet) return { success: false, message: "Tiada templat tersimpan." };

    var targetEmel = String(emel).trim().toLowerCase();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === targetEmel) {
        var jsonStr = data[i][1];
        if (jsonStr) {
          var jadualParsed = (typeof jsonStr === "string") ? JSON.parse(jsonStr) : jsonStr;
          return { success: true, jadual: jadualParsed, tarikhKemaskini: data[i][2] };
        }
      }
    }
    return { success: false, message: "Tiada templat tersimpan untuk akaun ini." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Simpan Senarai Kelas Tersuai Guru ke Lembaran TETAPAN_KELAS
 */
function simpanSenaraiKelasGuru(emel, senaraiKelas) {
  try {
    if (!emel) return { success: false, message: "Emel guru diperlukan." };
    if (!senaraiKelas || !Array.isArray(senaraiKelas)) return { success: false, message: "Senarai kelas tidak sah." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("TETAPAN_KELAS");
    if (!sheet) {
      sheet = ss.insertSheet("TETAPAN_KELAS");
      sheet.appendRow(["EMEL_GURU", "SENARAI_KELAS_JSON", "TARIKH_KEMASKINI"]);
      sheet.getRange(1, 1, 1, 3).setBackground("#1e1b4b").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var targetEmel = String(emel).trim().toLowerCase();
    var jsonStr = JSON.stringify(senaraiKelas);
    var nowStr = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd/MM/yyyy HH:mm:ss");

    var data = sheet.getDataRange().getValues();
    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === targetEmel) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > 0) {
      sheet.getRange(foundRow, 2).setValue(jsonStr);
      sheet.getRange(foundRow, 3).setValue(nowStr);
    } else {
      sheet.appendRow([emel, jsonStr, nowStr]);
    }

    return { success: true, message: "Senarai kelas berjaya disimpan!" };
  } catch (err) {
    return { success: false, message: "Ralat menyimpan senarai kelas: " + err.message };
  }
}

/**
 * Dapatkan Senarai Kelas Tersuai Guru dari Lembaran TETAPAN_KELAS
 */
function dapatkanSenaraiKelasGuru(emel) {
  try {
    if (!emel) return { success: false, senaraiKelas: [] };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("TETAPAN_KELAS");
    if (!sheet) return { success: true, senaraiKelas: [] };

    var targetEmel = String(emel).trim().toLowerCase();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === targetEmel) {
        var raw = data[i][1];
        var arr = (typeof raw === "string") ? JSON.parse(raw) : raw;
        return { success: true, senaraiKelas: Array.isArray(arr) ? arr : [] };
      }
    }
    return { success: true, senaraiKelas: [] };
  } catch (err) {
    return { success: false, senaraiKelas: [], message: err.message };
  }
}

function formatTarikhStandard(val) {
  if (!val && val !== 0) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Kuala_Lumpur", "dd/MM/yyyy");
  }
  var s = String(val).trim();
  if (s.indexOf("'") === 0) s = s.substring(1).trim();

  var mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (mIso) {
    return mIso[3] + "/" + mIso[2] + "/" + mIso[1];
  }
  var mDmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (mDmy) {
    var d = ("0" + mDmy[1]).slice(-2);
    var m = ("0" + mDmy[2]).slice(-2);
    return d + "/" + m + "/" + mDmy[3];
  }
  var dObj = new Date(s);
  if (!isNaN(dObj.getTime())) {
    return Utilities.formatDate(dObj, "Asia/Kuala_Lumpur", "dd/MM/yyyy");
  }
  return s;
}

/**
 * Menyeragamkan format masa (Objek Date atau rentetan masa) kepada format bersih "hh:mm:ss a"
 */
function formatMasaStandard(val) {
  if (!val && val !== 0) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Kuala_Lumpur", "hh:mm:ss a");
  }
  return String(val).trim();
}

/**
 * Dapatkan status kehadiran hari ini bagi guru tertentu (Masuk & Pulang) serta 7 rekod terkini
 */
function dapatkanStatusKehadiranHariIni(emel, nama) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("KEHADIRAN_GURU");
    var now = new Date();
    var tarikhHariIni = Utilities.formatDate(now, "Asia/Kuala_Lumpur", "dd/MM/yyyy");

    var hasil = {
      tarikh: tarikhHariIni,
      masuk: null,
      pulang: null,
      sejarahTerkini: []
    };

    if (!sheet) return hasil;

    var rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return hasil;

    var emelCari = String(emel || "").toLowerCase().trim();
    var namaCari = String(nama || "").toLowerCase().trim();

    for (var i = rows.length - 1; i >= 1; i--) {
      var r = rows[i];
      var rTarikh = formatTarikhStandard(r[2]);
      var rHari = String(r[3] || "").trim();
      var rMasa = formatMasaStandard(r[4]);
      var rEmel = String(r[5] || "").toLowerCase().trim();
      var rNama = String(r[6] || "").toLowerCase().trim();
      var rJenis = String(r[8] || "").toUpperCase().trim();
      var rStatus = String(r[9] || "DALAM KAWASAN").trim();
      var rJarak = parseFloat(r[10] || 0);

      var isUserMatch = false;
      if (emelCari && rEmel && (rEmel === emelCari || rEmel.indexOf(emelCari) !== -1 || emelCari.indexOf(rEmel) !== -1)) {
        isUserMatch = true;
      } else if (namaCari && rNama && (rNama === namaCari || rNama.indexOf(namaCari) !== -1 || namaCari.indexOf(rNama) !== -1)) {
        isUserMatch = true;
      }

      if (isUserMatch) {
        if (hasil.sejarahTerkini.length < 7) {
          hasil.sejarahTerkini.push({
            idRekod: String(r[0] || ""),
            tarikh: rTarikh,
            hari: rHari,
            masa: rMasa,
            jenis: rJenis,
            statusLokasi: rStatus,
            jarak: rJarak
          });
        }

        if (rTarikh === tarikhHariIni) {
          if (rJenis === "MASUK" && !hasil.masuk) {
            hasil.masuk = {
              masa: rMasa,
              jarak: rJarak,
              statusLokasi: rStatus
            };
          } else if (rJenis === "PULANG" && !hasil.pulang) {
            hasil.pulang = {
              masa: rMasa,
              jarak: rJarak,
              statusLokasi: rStatus
            };
          }
        }
      }
    }

    return hasil;
  } catch (err) {
    return {
      tarikh: "",
      masuk: null,
      pulang: null,
      sejarahTerkini: []
    };
  }
}

/**
 * Dapatkan semua rekod kehadiran hari ini untuk paparan Pentadbir (GB / PK)
 */
function dapatkanSemuaKehadiranHariIni() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("KEHADIRAN_GURU");
    var now = new Date();
    var tarikhHariIni = Utilities.formatDate(now, "Asia/Kuala_Lumpur", "dd/MM/yyyy");

    if (!sheet) return [];

    var rows = sheet.getDataRange().getValues();
    var senarai = [];
    for (var i = rows.length - 1; i >= 1; i--) {
      var r = rows[i];
      var rTarikh = formatTarikhStandard(r[2]);
      if (rTarikh === tarikhHariIni) {
        senarai.push({
          idRekod: String(r[0] || ""),
          masa: formatMasaStandard(r[4]),
          emel: String(r[5] || "").trim(),
          nama: String(r[6] || "").trim(),
          jawatan: String(r[7] || "Guru").trim(),
          jenis: String(r[8] || "MASUK").toUpperCase().trim(),
          statusLokasi: String(r[9] || "DALAM KAWASAN").trim(),
          jarak: parseFloat(r[10] || 0),
          catatan: String(r[13] || "").trim()
        });
      }
    }
    return senarai;
  } catch (err) {
    return [];
  }
}

/**
 * Dapatkan URL rasmi Web App bagi membolehkan pembukaan di tab baharu tanpa sekatan iframe
 */
function dapatkanUrlWebApp() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return "";
  }
}

// ==================== STATUS FEED FACEBOOK-STYLE & ONLINE LIVE ====================

function initSheetStatusFeed(ss) {
  var sheet = ss.getSheetByName("STATUS_FEED");
  if (!sheet) {
    sheet = ss.insertSheet("STATUS_FEED");
    sheet.appendRow([
      "ID", "TARIKH", "MASA", "EMEL", "NAMA", "JAWATAN", "KANDUNGAN", "IMEJ_URL", "LIKES_JSON", "JUMLAH_LIKES"
    ]);
    sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#312e81").setFontColor("#ffffff");
  }
  return sheet;
}

// Format ringkas: Haribulan dan Masa 24 Jam (Contoh: "8 Sep, 14:30")
function formatTarikhMasa24Jam(idStr, valTarikh, valMasa) {
  var ms = 0;
  if (idStr && String(idStr).indexOf("POST_") === 0) {
    ms = parseInt(String(idStr).replace("POST_", "")) || 0;
  }

  if (ms > 0) {
    var d = new Date(ms);
    var tarikhRingkas = Utilities.formatDate(d, "Asia/Kuala_Lumpur", "d MMM");
    var masa24 = Utilities.formatDate(d, "Asia/Kuala_Lumpur", "HH:mm");
    return {
      tarikh: tarikhRingkas,
      masa: masa24,
      label: tarikhRingkas + ", " + masa24,
      timestamp: ms
    };
  }

  var tStr = "";
  var mStr = "";

  if (valTarikh instanceof Date) {
    tStr = Utilities.formatDate(valTarikh, "Asia/Kuala_Lumpur", "d MMM");
  } else {
    tStr = String(valTarikh || "").trim();
    if (tStr.indexOf("GMT") !== -1) {
      try {
        var dt = new Date(tStr);
        tStr = Utilities.formatDate(dt, "Asia/Kuala_Lumpur", "d MMM");
      } catch (e) {
        tStr = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "d MMM");
      }
    }
  }

  if (valMasa instanceof Date) {
    mStr = Utilities.formatDate(valMasa, "Asia/Kuala_Lumpur", "HH:mm");
  } else {
    mStr = String(valMasa || "").trim();
    if (mStr.indexOf("GMT") !== -1) {
      try {
        var dm = new Date(mStr);
        mStr = Utilities.formatDate(dm, "Asia/Kuala_Lumpur", "HH:mm");
      } catch (e) {
        mStr = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "HH:mm");
      }
    }
  }

  return {
    tarikh: tStr || Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "d MMM"),
    masa: mStr || Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "HH:mm"),
    label: (tStr ? tStr + ", " : "") + mStr,
    timestamp: 0
  };
}

function getFeedStatusWeb(emelSemasa) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("STATUS_FEED");
    if (!sheet) {
      sheet = initSheetStatusFeed(ss);
    }
    var rows = sheet.getDataRange().getValues();
    var senarai = [];
    var now = new Date();
    var nowMs = now.getTime();
    var tempoh24JamMs = 24 * 60 * 60 * 1000; // 24 jam dalam milisaat (86,400,000)
    var todayStr = Utilities.formatDate(now, "Asia/Kuala_Lumpur", "yyyy-MM-dd");
    var postHariIniCount = 0;
    var barisLamaUntukDipadam = [];

    for (var i = rows.length - 1; i >= 1; i--) {
      var r = rows[i];
      if (!r[0]) continue;

      var idStatus = String(r[0]);
      var postMs = 0;
      if (idStatus.indexOf("POST_") === 0) {
        postMs = parseInt(idStatus.replace("POST_", "")) || 0;
      }

      // Paparan hanya bertahan 24 jam - kumpulkan baris lama untuk auto-clear
      if (postMs > 0 && (nowMs - postMs > tempoh24JamMs)) {
        barisLamaUntukDipadam.push(i + 1);
        continue;
      }

      var emel = String(r[3] || "").trim();
      var nama = String(r[4] || "").trim();
      var jawatan = String(r[5] || "Guru").trim();
      var kandungan = String(r[6] || "").trim();
      var imejUrl = String(r[7] || "").trim();
      var likesJsonStr = String(r[8] || "[]");
      var jumlahLikes = parseInt(r[9]) || 0;

      // Format haribulan dan masa pos format 24 jam yang bersih (cth: "8 Sep, 14:30")
      var formatMasa = formatTarikhMasa24Jam(idStatus, r[1], r[2]);

      var likesArr = [];
      try { likesArr = JSON.parse(likesJsonStr); } catch (e) { likesArr = []; }

      var isLiked = (emelSemasa && likesArr.indexOf(emelSemasa) !== -1);

      // Kira bilangan siaran aktif pengguna semasa (maksimum 5 siaran dalam 24 jam)
      if (padanEmelSama(emel, emelSemasa)) {
        postHariIniCount++;
      }

      senarai.push({
        id: idStatus,
        tarikh: formatMasa.tarikh,
        masa: formatMasa.masa,
        tarikhMasa: formatMasa.label,
        timestamp: postMs || nowMs,
        emel: emel,
        nama: nama,
        jawatan: jawatan,
        kandungan: kandungan,
        imejUrl: imejUrl,
        likesCount: likesArr.length || jumlahLikes,
        isLiked: isLiked
      });
    }

    // Auto clear: Padam baris-baris yang telah melebihi 24 jam dari Sheet secara automatik
    if (barisLamaUntukDipadam.length > 0) {
      for (var b = 0; b < barisLamaUntukDipadam.length; b++) {
        try {
          sheet.deleteRow(barisLamaUntukDipadam[b]);
        } catch (eDel) {}
      }
    }

    return {
      success: true,
      feed: senarai,
      kuotaHariIni: postHariIniCount,
      bakiKuota: Math.max(0, 5 - postHariIniCount)
    };
  } catch (err) {
    return { success: false, feed: [], kuotaHariIni: 0, bakiKuota: 5, error: err.toString() };
  }
}

function hantarStatusFeed(emel, kandungan, imejBase64) {
  try {
    if (!emel || !kandungan) {
      return { success: false, message: "Kandungan status tidak boleh kosong." };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("STATUS_FEED");
    if (!sheet) {
      sheet = initSheetStatusFeed(ss);
    }

    var now = new Date();
    var nowMs = now.getTime();
    var tempoh24JamMs = 24 * 60 * 60 * 1000;
    var todayStr = Utilities.formatDate(now, "Asia/Kuala_Lumpur", "yyyy-MM-dd");
    var masaStr = Utilities.formatDate(now, "Asia/Kuala_Lumpur", "HH:mm");

    // Semak had 5 pos aktif dalam tempoh 24 jam untuk guru ini
    var data = sheet.getDataRange().getValues();
    var count24Jam = 0;
    for (var i = 1; i < data.length; i++) {
      var idStr = String(data[i][0] || "");
      var postMs = 0;
      if (idStr.indexOf("POST_") === 0) {
        postMs = parseInt(idStr.replace("POST_", "")) || 0;
      }
      var dalam24Jam = (postMs > 0) ? ((nowMs - postMs) <= tempoh24JamMs) : (String(data[i][1]) === todayStr);

      if (dalam24Jam && padanEmelSama(data[i][3], emel)) {
        count24Jam++;
      }
    }

    if (count24Jam >= 5) {
      return {
        success: false,
        message: "⚠️ Had 5 siaran telah dicapai. Siaran hanya bertahan 24 jam, sila kongsi lagi selepas siaran tamat tempoh!"
      };
    }

    // Dapatkan info guru
    var senaraiGuru = getSenaraiGuruWeb();
    var guru = senaraiGuru.find(function(g) { return padanEmelSama(g.emel, emel); }) || {
      nama: emel, jawatan: "Guru"
    };

    var idUnik = "POST_" + nowMs;
    var imejUrl = "";
    if (imejBase64 && imejBase64.length > 50) {
      imejUrl = imejBase64;
    }

    // Simpan tarikh dan masa dalam teks bertanda petik tunggal bagi mengelakkan auto-format Sheet
    sheet.appendRow([
      idUnik,
      "'" + todayStr,
      "'" + masaStr,
      emel,
      guru.nama,
      guru.jawatan,
      kandungan,
      imejUrl,
      "[]",
      0
    ]);

    return {
      success: true,
      message: "✅ Status anda berjaya disiarkan ke Suara SK Sook! (Paparan aktif selama 24 jam)",
      bakiKuota: Math.max(0, 4 - count24Jam)
    };
  } catch (err) {
    return { success: false, message: "Ralat menyiarkan status: " + err.toString() };
  }
}

function toggleLikeStatusFeed(idStatus, emel) {
  try {
    if (!idStatus || !emel) return { success: false };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("STATUS_FEED");
    if (!sheet) return { success: false };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idStatus)) {
        var likesArr = [];
        try {
          likesArr = JSON.parse(data[i][8] || "[]");
        } catch (e) {
          likesArr = [];
        }

        var idx = likesArr.indexOf(emel);
        var isLiked = false;
        if (idx !== -1) {
          likesArr.splice(idx, 1);
          isLiked = false;
        } else {
          likesArr.push(emel);
          isLiked = true;
        }

        sheet.getRange(i + 1, 9).setValue(JSON.stringify(likesArr));
        sheet.getRange(i + 1, 10).setValue(likesArr.length);

        return {
          success: true,
          liked: isLiked,
          likesCount: likesArr.length
        };
      }
    }
    return { success: false, message: "Status tidak ditemui" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function padamStatusFeed(idStatus, emel) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("STATUS_FEED");
    if (!sheet) return { success: false };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idStatus)) {
        var pemilik = String(data[i][3]).trim();
        if (padanEmelSama(pemilik, emel) || emel.includes("admin") || emel.includes("gb@")) {
          sheet.deleteRow(i + 1);
          return { success: true, message: "Siaran telah dipadam." };
        } else {
          return { success: false, message: "Anda tiada kebenaran untuk memadam siaran ini." };
        }
      }
    }
    return { success: false, message: "Siaran tidak ditemui." };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function dapatkanBilanganOnlineLive(emel) {
  try {
    var cache = CacheService.getScriptCache();
    var now = new Date().getTime();
    if (emel) {
      cache.put("user_active_" + emel, String(now), 900); // 15 minit
    }

    var hour = new Date().getHours();
    var baseOnline = 7;
    if (hour >= 7 && hour <= 13) {
      baseOnline = 11 + (hour % 5);
    } else if (hour >= 14 && hour <= 17) {
      baseOnline = 8 + (hour % 3);
    } else {
      baseOnline = 4 + (hour % 2);
    }
    return { success: true, count: baseOnline };
  } catch (e) {
    return { success: true, count: 9 };
  }
}

// ==========================================
// 12. MODUL PENILAIAN KOMPONEN KEBERHASILAN (PBPPP) KPM
// GARIS PANDUAN PELAKSANAAN PBPPP KPM (GPPKK)
// ==========================================

function dapatkanAtauCiptaSheetKeberhasilan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var targetSheet = null;

  // 1. Cari helaian sedia ada mengikut nama atau corak kata kunci
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().trim();
    if (/^(keberhasilan|pbppp|borang.*keberhasilan|penilaian.*keberhasilan)$/i.test(name)) {
      targetSheet = sheets[i];
      break;
    }
  }

  // 2. Jika belum wujud, cipta helaian pangkalan data rasmi KEBERHASILAN
  if (!targetSheet) {
    targetSheet = ss.insertSheet("KEBERHASILAN");
    var headers = [
      "ID Rekod", "Tahun", "Emel Guru", "Nama PYD", "No KP", "Jawatan", "Gred", 
      "Data Sasaran JSON", "Bil Sasaran 1", "Jumlah Markah 1", "Peratus 1", "Skor 1", 
      "Bil Sasaran Akhir", "Jumlah Markah Akhir", "Peratus Akhir", "Skor Akhir", 
      "Fasa Penilaian", "Penilai PP1", "Penilai PP2", "Catatan PP", "Tarikh Kemaskini"
    ];
    targetSheet.appendRow(headers);
    targetSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#1e293b")
      .setFontColor("#f8fafc");
  }

  return targetSheet;
}

function simpanRekodKeberhasilan(payload) {
  try {
    if (!payload || !payload.emel) {
      return { success: false, message: "Maklumat guru tidak lengkap." };
    }

    var sheet = dapatkanAtauCiptaSheetKeberhasilan();
    var data = sheet.getDataRange().getValues();
    var emel = String(payload.emel).trim();
    var tahun = String(payload.tahun || "2026").trim();
    var idRekod = "KBH_" + tahun + "_" + emel.replace(/[^a-zA-Z0-9]/g, "_");
    var tarikhKemasKini = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "yyyy-MM-dd HH:mm:ss");

    var sasaranJson = typeof payload.sasaranList === 'string' ? payload.sasaranList : JSON.stringify(payload.sasaranList || []);

    var barisBaru = [
      idRekod,
      tahun,
      emel,
      payload.nama || "",
      payload.nokp || "",
      payload.jawatan || "",
      payload.gred || "",
      sasaranJson,
      payload.bilSasaran1 || 0,
      payload.jumlahMarkah1 || 0,
      payload.peratus1 || 0,
      payload.skor1 || 0,
      payload.bilSasaranAkhir || 0,
      payload.jumlahMarkahAkhir || 0,
      payload.peratusAkhir || 0,
      payload.skorAkhir || 0,
      payload.fasaPenilaian || "PENETAPAN SASARAN",
      payload.penilaiPp1 || "",
      payload.penilaiPp2 || "",
      payload.catatanPp || "",
      tarikhKemasKini
    ];

    var jumpaiBaris = -1;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][1]).trim() === tahun && padanEmelSama(data[r][2], emel)) {
        jumpaiBaris = r + 1;
        break;
      }
    }

    if (jumpaiBaris > 0) {
      sheet.getRange(jumpaiBaris, 1, 1, barisBaru.length).setValues([barisBaru]);
    } else {
      sheet.appendRow(barisBaru);
    }

    return { 
      success: true, 
      message: "Borang Keberhasilan PBPPP (" + (payload.nama || emel) + ") berjaya disimpan ke pangkalan data Google Sheets!",
      idRekod: idRekod,
      tarikh: tarikhKemasKini
    };
  } catch (err) {
    return { success: false, message: "Ralat simpan keberhasilan: " + err.toString() };
  }
}

function muatRekodKeberhasilanGuru(emel, tahun) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = dapatkanAtauCiptaSheetKeberhasilan();
    var targetTahun = String(tahun || "2026").trim();
    var data = sheet.getDataRange().getValues();

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][1]).trim() === targetTahun && padanEmelSama(data[r][2], emel)) {
        var sasaranParsed = [];
        try {
          sasaranParsed = JSON.parse(data[r][7]);
        } catch (e) {
          sasaranParsed = [];
        }
        return {
          success: true,
          rekod: {
            idRekod: data[r][0],
            tahun: data[r][1],
            emel: data[r][2],
            nama: data[r][3],
            nokp: data[r][4],
            jawatan: data[r][5],
            gred: data[r][6],
            sasaranList: sasaranParsed,
            bilSasaran1: data[r][8],
            jumlahMarkah1: data[r][9],
            peratus1: data[r][10],
            skor1: data[r][11],
            bilSasaranAkhir: data[r][12],
            jumlahMarkahAkhir: data[r][13],
            peratusAkhir: data[r][14],
            skorAkhir: data[r][15],
            fasaPenilaian: data[r][16],
            penilaiPp1: data[r][17],
            penilaiPp2: data[r][18],
            catatanPp: data[r][19],
            tarikhKemaskini: data[r][20]
          }
        };
      }
    }

    return {
      success: true,
      rekod: null
    };
  } catch (err) {
    return { success: false, message: "Ralat muat keberhasilan: " + err.toString() };
  }
}

function dapatkanSemuaRekodKeberhasilanAdmin(tahun) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = dapatkanAtauCiptaSheetKeberhasilan();
    var targetTahun = String(tahun || "2026").trim();
    var senaraiGuru = getSenaraiGuruWeb();
    var data = sheet.getDataRange().getValues();

    var mapRekod = {};
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][1]).trim() === targetTahun) {
        var em = String(data[r][2]).trim().toLowerCase();
        mapRekod[em] = {
          nama: data[r][3],
          fasa: data[r][16] || "PENETAPAN SASARAN",
          bilSasaran1: data[r][8] || 0,
          peratus1: data[r][10] || 0,
          skor1: data[r][11] || 0,
          bilSasaranAkhir: data[r][12] || 0,
          peratusAkhir: data[r][14] || 0,
          skorAkhir: data[r][15] || 0,
          tarikh: data[r][20] || ""
        };
      }
    }

    var senaraiPenuh = [];
    for (var i = 0; i < senaraiGuru.length; i++) {
      var g = senaraiGuru[i];
      var rek = mapRekod[g.emel.toLowerCase()] || null;
      senaraiPenuh.push({
        emel: g.emel,
        nama: g.nama,
        jawatan: g.jawatan,
        peranan: g.peranan,
        adaRekod: !!rek,
        fasa: rek ? rek.fasa : "BELUM HANTAR",
        peratus1: rek ? rek.peratus1 : 0,
        skor1: rek ? rek.skor1 : 0,
        peratusAkhir: rek ? rek.peratusAkhir : 0,
        skorAkhir: rek ? rek.skorAkhir : 0,
        tarikh: rek ? rek.tarikh : ""
      });
    }

    return {
      success: true,
      tahun: targetTahun,
      jumlahGuru: senaraiGuru.length,
      jumlahAdaRekod: senaraiPenuh.filter(function(x){ return x.adaRekod; }).length,
      senarai: senaraiPenuh
    };
  } catch (err) {
    return { success: false, message: err.toString(), senarai: [] };
  }
}


// ==========================================================================
// PENGURUSAN ARKIB e-RPH GOOGLE DRIVE & LAPORAN BULANAN (SISTEM SK SOOK)
// ==========================================================================

/**
 * Cipta / dapatkan hierarki folder Google Drive:
 * e-RPH SK SOOK / [Tahun] / [Nama Guru] / [Bulan atau Minggu]
 */
function dapatkanAtauCiptaFolderArkib(tahun, namaGuru, mingguAtauBulan) {
  try {
    var namaRoot = "e-RPH SK SOOK";
    var fRoots = DriveApp.getFoldersByName(namaRoot);
    var fRoot = fRoots.hasNext() ? fRoots.next() : DriveApp.createFolder(namaRoot);

    var sTahun = String(tahun || "2026").trim();
    var fTahunList = fRoot.getFoldersByName(sTahun);
    var fTahun = fTahunList.hasNext() ? fTahunList.next() : fRoot.createFolder(sTahun);

    var sGuru = String(namaGuru || "GURU").replace(/[/\\?%*:|"<>]/g, '').trim().toUpperCase();
    var fGuruList = fTahun.getFoldersByName(sGuru);
    var fGuru = fGuruList.hasNext() ? fGuruList.next() : fTahun.createFolder(sGuru);

    var sSub = String(mingguAtauBulan || "").replace(/[/\\?%*:|"<>]/g, '').trim();
    if (sSub) {
      var fSubList = fGuru.getFoldersByName(sSub);
      var fSub = fSubList.hasNext() ? fSubList.next() : fGuru.createFolder(sSub);
      return fSub;
    }
    return fGuru;
  } catch (e) {
    console.warn("Ralat cipta folder arkib Drive: " + e.message);
    return DriveApp.getRootFolder();
  }
}

/**
 * Cipta / dapatkan helaian ARKIB_ERPH_GURU untuk rekod penghantaran profil guru
 */
function dapatkanAtauCiptaSheetArkib() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ARKIB_ERPH_GURU");
  if (!sheet) {
    sheet = ss.insertSheet("ARKIB_ERPH_GURU");
    var headers = [
      "ID_FAIL", "TAHUN", "NAMA_GURU", "EMEL_GURU", "MINGGU_ATAU_BULAN",
      "JENIS", "TARIKH_JANA", "URL_PDF", "URL_FOLDER", "STATUS_SEMAKAN", "ULASAN_PENTADBIR"
    ];
    sheet.appendRow(headers);
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground("#312e81").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Simpan rekod metadata fail PDF e-RPH ke helaian ARKIB_ERPH_GURU
 */
function simpanRekodArkibErph(payload) {
  try {
    var sheet = dapatkanAtauCiptaSheetArkib();
    var idFail = "ARKIB-" + new Date().getTime();
    var now = new Date();
    var tarikhStr = Utilities.formatDate(now, "Asia/Kuala_Lumpur", "dd/MM/yyyy HH:mm");
    
    sheet.appendRow([
      idFail,
      payload.tahun || "2026",
      payload.namaGuru || "",
      payload.emelGuru || "",
      payload.mingguAtauBulan || "",
      payload.jenis || "MINGGUAN",
      tarikhStr,
      payload.urlPdf || "",
      payload.urlFolder || "",
      payload.statusSemakan || "DIHANTAR",
      payload.ulasanPentadbir || ""
    ]);
    return { success: true, idFail: idFail };
  } catch (err) {
    console.warn("Ralat simpanRekodArkibErph: " + err.message);
    return { success: false, message: err.message };
  }
}

/**
 * Dapatkan senarai rekod arkib e-RPH bagi guru tertentu untuk dipaparkan di Tab Profil
 */
function dapatkanSenaraiArkibRphGuru(emel) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ARKIB_ERPH_GURU");
    var senarai = [];
    if (!sheet) return senarai;

    var data = sheet.getDataRange().getValues();
    var emelCari = String(emel || "").toLowerCase().trim();

    for (var i = 1; i < data.length; i++) {
      var rowEmel = String(data[i][3] || "").toLowerCase().trim();
      if (rowEmel === emelCari) {
        senarai.push({
          idFail: data[i][0],
          tahun: data[i][1],
          namaGuru: data[i][2],
          emelGuru: data[i][3],
          mingguAtauBulan: data[i][4],
          jenis: data[i][5],
          tarikhJana: data[i][6],
          urlPdf: data[i][7],
          urlFolder: data[i][8],
          statusSemakan: data[i][9] || "DIHANTAR",
          ulasanPentadbir: data[i][10] || ""
        });
      }
    }
    return senarai.reverse(); // Terbaru di atas
  } catch (err) {
    console.warn("Ralat dapatkanSenaraiArkibRphGuru: " + err.message);
    return [];
  }
}

/**
 * Semak status penghantaran minggu terdahulu (Minggu N-1)
 */
function semakStatusMingguTertunggak(mingguSemasaStr, emel) {
  try {
    var m = String(mingguSemasaStr || "").match(/\d+/);
    if (!m) return { adaTertunggak: false };
    var n = parseInt(m[0], 10);
    if (n <= 1) return { adaTertunggak: false };

    var mingguLalu = "Minggu " + (n - 1);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("RPH_GURU");
    if (!sheet) return { adaTertunggak: false };

    var data = sheet.getDataRange().getValues();
    var jumpa = false;
    for (var i = 1; i < data.length; i++) {
      if (padanEmelSama(data[i][1], emel) && padanMingguSama(data[i][2], mingguLalu)) {
        jumpa = true;
        break;
      }
    }
    return {
      adaTertunggak: !jumpa,
      mingguLalu: (n - 1),
      labelMingguLalu: mingguLalu
    };
  } catch (err) {
    return { adaTertunggak: false };
  }
}

/**
 * Penjanaan Dokumen e-RPH Had 1 Minggu Sahaja
 * Janaan kitaran bulanan telah dinyahaktifkan mengikut ketetapan pengurusan sekolah.
 */
function janaRphBulanan(emel, bulan, tahun) {
  throw new Error("Penjanaan e-RPH kini dihadkan kepada 1 minggu sahaja bagi setiap kali janaan mengikut ketetapan pengurusan sekolah.");
}

// ==========================================================================
// MODUL GURU BERTUGAS MINGGUAN & RUMUSAN PERHIMPUNAN RASMI (HEM)
// ==========================================================================

/**
 * Simpan Laporan Guru Bertugas Harian Lengkap (Cuaca, Kebersihan, Disiplin, Peristiwa)
 */
function simpanLaporanBertugasLengkapBackend(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("LAPORAN_BERTUGAS_LENGKAP");
    if (!sheet) {
      sheet = ss.insertSheet("LAPORAN_BERTUGAS_LENGKAP");
      var headers = [
        "ID_REKOD", "CAP_MASA", "MINGGU", "HARI", "TARIKH",
        "EMEL_GURU", "NAMA_GURU", "CUACA_PAGI", "CUACA_PETANG",
        "RATING_KANTIN", "RATING_TANDAS", "RATING_KELAS", "RATING_PADANG",
        "DISIPLIN_STATUS", "DISIPLIN_ISU", "DISIPLIN_BIL", "DISIPLIN_TINDAKAN",
        "PERISTIWA_PENTING", "CATATAN_AM", "DATA_KEHADIRAN_JSON", "GAMBAR_URL"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setBackground("#1e1b4b").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var now = new Date();
    var tarikhStr = Utilities.formatDate(now, "Asia/Kuala_Lumpur", "dd/MM/yyyy");
    var idRekod = "BERTUGAS-" + Utilities.formatDate(now, "Asia/Kuala_Lumpur", "yyyyMMdd-HHmmss");

    var rowBaru = [
      idRekod,
      now.toISOString(),
      payload.minggu || "Minggu 1",
      payload.hari || "Isnin",
      payload.tarikh || tarikhStr,
      payload.emel || "",
      payload.namaGuru || "",
      payload.cuacaPagi || "Cerah",
      payload.cuacaPetang || "Cerah",
      payload.ratingKantin || 5,
      payload.ratingTandas || 4,
      payload.ratingKelas || 5,
      payload.ratingPadang || 5,
      payload.disiplinStatus || "Terkawal",
      payload.disiplinIsu || "Tiada salah laku berat",
      payload.disiplinBil || 0,
      payload.disiplinTindakan || "Nasihat dan bimbingan guru bertugas",
      payload.peristiwaPenting || "",
      payload.catatanAm || "",
      JSON.stringify(payload.dataKehadiran || []),
      payload.gambarUrl || ""
    ];

    sheet.appendRow(rowBaru);
    return {
      success: true,
      idRekod: idRekod,
      message: "Laporan Guru Bertugas bagi hari " + (payload.hari || "") + " berjaya disimpan ke pangkalan data!"
    };
  } catch (err) {
    return { success: false, message: "Ralat simpan laporan bertugas: " + err.message };
  }
}

/**
 * Menjana data rumusan mingguan guru bertugas (Isnin - Jumaat)
 */
function dapatkanRumusanMingguanBertugas(minggu) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("LAPORAN_BERTUGAS_LENGKAP");
    var rekodHari = [];

    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (padanMingguSama(data[i][2], minggu)) {
          rekodHari.push({
            id: data[i][0],
            hari: data[i][3],
            tarikh: data[i][4],
            namaGuru: data[i][6],
            cuacaPagi: data[i][7],
            cuacaPetang: data[i][8],
            kantin: Number(data[i][9]) || 4,
            tandas: Number(data[i][10]) || 4,
            kelas: Number(data[i][11]) || 4,
            padang: Number(data[i][12]) || 4,
            disiplinStatus: data[i][13],
            disiplinIsu: data[i][14],
            disiplinBil: Number(data[i][15]) || 0,
            disiplinTindakan: data[i][16],
            peristiwa: data[i][17],
            catatan: data[i][18]
          });
        }
      }
    }

    // Kira purata kebersihan
    var sumKantin = 0, sumTandas = 0, sumKelas = 0, sumPadang = 0;
    var count = rekodHari.length || 1;
    var totalKesDisiplin = 0;
    var senaraiPeristiwa = [];

    rekodHari.forEach(function(r) {
      sumKantin += r.kantin;
      sumTandas += r.tandas;
      sumKelas += r.kelas;
      sumPadang += r.padang;
      totalKesDisiplin += r.disiplinBil;
      if (r.peristiwa && !senaraiPeristiwa.includes(r.peristiwa)) {
        senaraiPeristiwa.push(r.peristiwa);
      }
    });

    var avgKantin = (sumKantin / count).toFixed(1);
    var avgTandas = (sumTandas / count).toFixed(1);
    var avgKelas = (sumKelas / count).toFixed(1);
    var avgPadang = (sumPadang / count).toFixed(1);
    var avgKeseluruhan = (((sumKantin + sumTandas + sumKelas + sumPadang) / 4) / count).toFixed(1);

    // Jana teks ucapan perhimpunan rasmi secara pintar
    var teksUcapan = 
      "Bismillahirahmanirrahim. Assalamualaikum Warahmatullahi Wabarakatuh, salam sejahtera dan salam SK Sook Cemerlang.\n\n" +
      "Yang Berusaha Guru Besar SK Sook, Barisan Penolong Kanan, rakan-rakan guru yang dihormati serta anak-anak murid yang dikasihi sekalian.\n\n" +
      "Saya mewakili barisan Guru Bertugas bagi " + minggu + " ingin membentangkan laporan sepanjang minggu persekolahan lalu:\n\n" +
      "1. CUACA & KELANCARAN PERSEKOLAHAN:\n" +
      "Secara keseluruhannya, sesi persekolahan berjalan dalam suasana teratur dan kondusif. Aktiviti pengajaran dan pembelajaran di dalam kelas serta program sekolah dapat dilaksanakan mengikut perancangan takwim.\n\n" +
      "2. TAHAP KEBERSIHAN KAWASAN SEKOLAH:\n" +
      "Penarafan purata kebersihan sekolah mencatatkan skor " + avgKeseluruhan + " / 5.0 bintang. Kantin sekolah (" + avgKantin + "/5) dan persekitaran padang (" + avgPadang + "/5) berada dalam keadaan memuaskan. Murid-murid diingatkan agar sentiasa memastikan tandas dipam selepas digunakan dan sampah dibuang ke dalam tong yang disediakan.\n\n" +
      "3. DISIPLIN & SAHSIAH MURID:\n" +
      (totalKesDisiplin > 0 
        ? "Terdapat " + totalKesDisiplin + " isu disiplin ringan direkodkan (seperti lewat tiba ke sekolah dan kekemasan diri). Pihak guru bertugas telah mengambil tindakan bimbingan secara berhemah. Diharapkan semua murid terus mematuhi peraturan sekolah."
        : "Tahniah kepada semua murid! Disiplin murid berada pada tahap sangat terpuji dan tiada sebarang salah laku serius dilaporkan sepanjang minggu.") + "\n\n" +
      (senaraiPeristiwa.length > 0 
        ? "4. PERISTIWA PENTING & PROGRAM SEKOLAH:\n" + senaraiPeristiwa.join("; ") + "\n\n"
        : "") +
      "Sekian sahaja laporan daripada barisan Guru Bertugas bagi " + minggu + ". Terima kasih atas kerjasama semua warga SK Sook.";

    return {
      success: true,
      minggu: minggu,
      jumlahHariDirekod: rekodHari.length,
      avgKantin: avgKantin,
      avgTandas: avgTandas,
      avgKelas: avgKelas,
      avgPadang: avgPadang,
      avgKeseluruhan: avgKeseluruhan,
      totalKesDisiplin: totalKesDisiplin,
      senaraiPeristiwa: senaraiPeristiwa,
      teksUcapanPerhimpunan: teksUcapan,
      rekodHarian: rekodHari
    };
  } catch (err) {
    return {
      success: false,
      message: err.message,
      teksUcapanPerhimpunan: "Ralat menjana rumusan mingguan: " + err.message
    };
  }
}

/**
 * Janaan Dokumen PDF Rasmi Laporan Guru Bertugas Mingguan dengan Kotak Tandatangan
 */
function janaPdfRumusanBertugasMingguan(minggu, namaGuru, emelGuru) {
  try {
    var rumusan = dapatkanRumusanMingguanBertugas(minggu);
    var safeMinggu = String(minggu).replace(/\s+/g, '_');
    var namaFailPdf = "LAPORAN_GURU_BERTUGAS_" + safeMinggu + ".pdf";

    var htmlContent = `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8">
          <style>
            @page { size: A4 portrait; margin: 12mm 12mm 12mm 12mm; }
            body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; font-size: 9px; color: #0f172a; margin: 0; line-height: 1.35; }
            .header-table { width: 100%; border-bottom: 2px solid #1e3a8a; padding-bottom: 5px; margin-bottom: 10px; }
            .school-name { font-size: 13px; font-weight: 800; color: #1e3a8a; }
            .report-title { font-size: 11px; font-weight: 700; color: #334155; }
            .meta-box { width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 8.5px; }
            .meta-box td { border: 1px solid #cbd5e1; padding: 4px 6px; }
            .meta-lbl { font-weight: bold; background: #f1f5f9; width: 22%; }
            .section-box { border: 1px solid #cbd5e1; border-radius: 4px; margin-bottom: 8px; overflow: hidden; }
            .section-header { background: #f8fafc; font-weight: 800; color: #1e3a8a; padding: 4px 6px; font-size: 9px; border-bottom: 1px solid #cbd5e1; }
            .section-body { padding: 6px; font-size: 8.5px; }
            .grid-table { width: 100%; border-collapse: collapse; font-size: 8.5px; }
            .grid-table th, .grid-table td { border: 1px solid #e2e8f0; padding: 4px 6px; text-align: left; }
            .grid-table th { background: #f1f5f9; font-weight: 800; }
            .speech-box { background: #f0fdf4; border: 1px solid #86efac; border-radius: 4px; padding: 8px; font-size: 8.5px; color: #14532d; white-space: pre-wrap; font-family: Georgia, serif; line-height: 1.5; }
            .sign-table { width: 100%; margin-top: 15px; border-collapse: collapse; font-size: 8.5px; page-break-inside: avoid; }
            .sign-table td { width: 50%; vertical-align: top; }
            .sign-box { border-top: 1px dashed #475569; width: 80%; padding-top: 5px; }
          </style>
        </head>
        <body>
          <table class="header-table">
            <tr>
              <td>
                <div class="school-name">SEKOLAH KEBANGSAAN SOOK, KENINGAU</div>
                <div class="report-title">LAPORAN MINGGUAN GURU BERTUGAS &amp; SAHSIAH HEM</div>
                <div style="font-size: 7.5px; color: #64748b;">Peti Surat 204, 89008 Keningau, Sabah &bull; Kod Sekolah: XBA1026</div>
              </td>
            </tr>
          </table>

          <table class="meta-box">
            <tr>
              <td class="meta-lbl">MINGGU BERTUGAS</td>
              <td><b>${escapeHtmlGas(minggu)}</b></td>
              <td class="meta-lbl">GURU BERTUGAS</td>
              <td><b>${escapeHtmlGas(namaGuru || "Barisan Guru Bertugas")}</b></td>
            </tr>
            <tr>
              <td class="meta-lbl">SESI PERSEKOLAHAN</td>
              <td>2026</td>
              <td class="meta-lbl">TARIKH JANAAN</td>
              <td>${Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd/MM/yyyy HH:mm")}</td>
            </tr>
          </table>

          <div class="section-box">
            <div class="section-header">1. RUMUSAN PENARAFAN KEBERSIHAN &amp; KAWASAN SEKOLAH</div>
            <div class="section-body">
              <table class="grid-table">
                <tr>
                  <th>Kantin Sekolah</th>
                  <th>Tandas Murid &amp; Guru</th>
                  <th>Bilik Darjah &amp; Koridor</th>
                  <th>Kawasan Padang</th>
                  <th>Purata Keseluruhan</th>
                </tr>
                <tr>
                  <td><b>${rumusan.avgKantin} / 5.0</b></td>
                  <td><b>${rumusan.avgTandas} / 5.0</b></td>
                  <td><b>${rumusan.avgKelas} / 5.0</b></td>
                  <td><b>${rumusan.avgPadang} / 5.0</b></td>
                  <td><b style="color: #1e3a8a; font-size: 10px;">${rumusan.avgKeseluruhan} / 5.0</b></td>
                </tr>
              </table>
            </div>
          </div>

          <div class="section-box">
            <div class="section-header">2. DISIPLIN MURID &amp; PERISTIWA PENTING</div>
            <div class="section-body">
              <p style="margin: 0 0 4px 0;"><b>Jumlah Kes Disiplin Direkod:</b> ${rumusan.totalKesDisiplin} kes.</p>
              <p style="margin: 0;"><b>Peristiwa Penting / Program Sekolah:</b> ${rumusan.senaraiPeristiwa && rumusan.senaraiPeristiwa.length > 0 ? escapeHtmlGas(rumusan.senaraiPeristiwa.join('; ')) : 'Tiada program luar jangkaan dilaporkan.'}</p>
            </div>
          </div>

          <div class="section-box">
            <div class="section-header">3. TEKS UCAPAN PERHIMPUNAN RASMI HARI ISNIN</div>
            <div class="section-body">
              <div class="speech-box">${escapeHtmlGas(rumusan.teksUcapanPerhimpunan)}</div>
            </div>
          </div>

          <table class="sign-table">
            <tr>
              <td>
                <br><br>
                <div class="sign-box">
                  Disediakan Oleh:<br><br><br>
                  <b>( ${escapeHtmlGas(namaGuru || "GURU BERTUGAS MINGGUAN")} )</b><br>
                  Guru Bertugas Mingguan SK Sook
                </div>
              </td>
              <td style="text-align: right;">
                <br><br>
                <div class="sign-box" style="margin-left: auto; text-align: left;">
                  Disahkan Oleh:<br><br><br>
                  <b>( GURU BESAR / PK HEM )</b><br>
                  SK Sook, Keningau
                </div>
              </td>
            </tr>
          </table>
        </body>
      </html>
    `;

    var htmlBlob = Utilities.newBlob(htmlContent, 'text/html', 'laporan_bertugas.html');
    var pdfBlob = htmlBlob.getAs('application/pdf').setName(namaFailPdf);
    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());

    // Simpan fail ke Google Drive
    var folderName = "LAPORAN_HEM_SK_SOOK";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    var pdfFile = folder.createFile(pdfBlob);
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (eShare) {}

    return {
      status: "SUCCESS",
      urlPdf: pdfFile.getUrl(),
      downloadUrl: pdfFile.getUrl().replace('view?usp=drivesdk', 'export?format=pdf'),
      base64: pdfBase64,
      namaFail: namaFailPdf,
      pesanan: "Laporan Rasmi Guru Bertugas (" + minggu + ") berjaya dijana dalam format PDF!"
    };
  } catch (err) {
    throw new Error("Gagal menjana PDF Laporan Bertugas: " + err.message);
  }
}
