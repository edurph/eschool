function mulaSediakanDaftar() {
  return sediakanDaftarSaas_();
}

function mulaCiptaFolder() {
  return provisionSekolahEditor_();
}

// FUNGSI KHAS UNTUK DAFTAR SEKOLAH BAHARU:
function daftarSekolahLotong() {
  PropertiesService.getScriptProperties().setProperty('TARGET_SCHOOL_ID', 'XBA1111');
  return provisionSekolahEditor_();
}

function mulaPasangAutomasi() {
  return pasangAutomasiSaas_();
}

function mulaPasangCheckout() {
  return pasangCheckoutSaas_();
}
