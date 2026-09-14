// ===== Code.gs =====
// Portal Sekolah SaaS — daftar pusat tidak boleh dikongsi kepada pengguna portal.
// Semua helper berakhir dengan _ supaya tidak boleh dipanggil melalui google.script.run.
var CTX_ = null;
var ROUTE_ = null;
var STRUCTURE_VERSION_ = '1';
var SCHOOL_HEADERS_ = ['Sekolah_ID','Kod_Sekolah','Nama_Sekolah','Folder_Drive_ID','Spreadsheet_ID','Ada_PPKI','Status_Langganan','Tarikh_Tamat','Status_Provisioning','Versi_Struktur','Ralat_Terakhir','Tarikh_Provisioning','Alamat','Logo_File_ID','Panitia_Tambahan_JSON','Unit_Koko_JSON','GPS_Lat','GPS_Lon','GPS_Radius'];
var PANITIA_ = [
  ['BM','01_Panitia_Bahasa_Melayu'],['BI','02_Panitia_Bahasa_Inggeris'],
  ['MT','03_Panitia_Matematik'],['SN','04_Panitia_Sains'],
  ['PI_PM','05_Panitia_Pendidikan_Islam_Moral'],['PJPK','06_Panitia_PJPK'],
  ['SENI_MUZIK','07_Panitia_Seni_Muzik'],['RBT_SEJ','08_Panitia_RBT_Sejarah']
];
var PANITIA_DOC_ = {DASAR:'01_Dokumen_Dasar',MESYUARAT:'02_Mesyuarat_Panitia',OPR:'03_Program_Kecemerlangan_OPR',PENTAKSIRAN:'04_Pentaksiran_PBD_UASA',BBM:'05_Bank_Soalan_BBM'};
var KOKO_DIR_ = {SUKAN:'OPR_Sukan_Permainan',KELAB:'OPR_Kelab_Persatuan',BERUNIFORM:'OPR_Badan_Beruniform'};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Portal Sekolah').addMetaTag('viewport','width=device-width, initial-scale=1');
}
function include_(file) { return HtmlService.createHtmlOutputFromFile(file).getContent(); }
function prop_(key) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('Konfigurasi belum lengkap: ' + key);
  return v;
}
function master_() { return SpreadsheetApp.openById(prop_('MASTER_SPREADSHEET_ID')); }
function tab_(name, headers) {
  var ss=master_(), sh=ss.getSheetByName(name);
  if (!sh) { sh=ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  var actual=sh.getRange(1,1,1,headers.length).getDisplayValues()[0];
  if (JSON.stringify(actual)!==JSON.stringify(headers)) throw new Error('Header tidak sepadan: '+name);
  return sh;
}
function rows_(sh) {
  var values=sh.getDataRange().getValues(), heads=values.shift();
  return values.map(function(r,i){var o={_row:i+2};heads.forEach(function(h,j){o[h]=r[j];});return o;});
}
function sekolahRows_() { return rows_(tab_('Langganan_Sekolah',SCHOOL_HEADERS_)); }
function norm_(s) {return String(s || '').trim();}
function email_(s) {return norm_(s).toLowerCase();}
function yes_(v) {return v===true || /^(true|ya|1)$/i.test(norm_(v));}
function hash_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(value),Utilities.Charset.UTF_8)
    .map(function(b){return ('0'+((b+256)%256).toString(16)).slice(-2);}).join('');
}
function unique_() { return Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,''); }
function lock_(fn) {var l=LockService.getScriptLock();l.waitLock(20000);try{return fn();}finally{l.releaseLock();}}
function sekolah_(id, allowInactive) {
  var matches=sekolahRows_().filter(function(r){return norm_(r.Sekolah_ID)===norm_(id);});
  if(matches.length!==1) throw new Error('Sekolah tidak berdaftar atau ID bertindih.');
  var s=matches[0];
  if(!allowInactive) {
    if(norm_(s.Status_Langganan).toUpperCase()!=='AKTIF') throw new Error('Langganan tidak aktif.');
    if(s.Tarikh_Tamat) {
      var end=s.Tarikh_Tamat instanceof Date ? Utilities.formatDate(s.Tarikh_Tamat,'Asia/Kuala_Lumpur','yyyy-MM-dd') : norm_(s.Tarikh_Tamat);
      if(!/^\d{4}-\d{2}-\d{2}$/.test(end)) throw new Error('Tarikh_Tamat mesti YYYY-MM-DD.');
      if(end<Utilities.formatDate(new Date(),'Asia/Kuala_Lumpur','yyyy-MM-dd')) throw new Error('Langganan telah tamat.');
    }
  }
  return s;
}
function konteks_() {if(!CTX_)throw new Error('Sesi sekolah diperlukan.');return CTX_;}
function sekolahSpreadsheet_() {return SpreadsheetApp.openById(konteks_().school.Spreadsheet_ID);}
function operator_() {
  var effective=email_(Session.getEffectiveUser().getEmail()), active=email_(Session.getActiveUser().getEmail());
  if(!active || active!==effective || active!==email_(prop_('OPERATOR_EMAIL'))) throw new Error('Jalankan fungsi ini dalam editor menggunakan akaun operator.');
}
function statusSekolah_(s,status,error) {
  var sh=tab_('Langganan_Sekolah',SCHOOL_HEADERS_);
  sh.getRange(s._row,9,1,4).setValues([[status,status==='READY'?STRUCTURE_VERSION_:s.Versi_Struktur||'',String(error||'').slice(0,800),new Date()]]);
}
function jsonConfig_(text,fallback) {if(!text)return fallback;try{return JSON.parse(String(text));}catch(e){throw new Error('Konfigurasi JSON sekolah tidak sah.');}}
function panitia_(s) {
  var extra=jsonConfig_(s.Panitia_Tambahan_JSON,[]), list=PANITIA_.map(function(p){return p.slice();}), seen={};
  if(!Array.isArray(extra))throw new Error('Panitia_Tambahan_JSON mesti array.');
  list.concat(extra).forEach(function(p){
    if(!Array.isArray(p)||p.length!==2||!/^\w{2,30}$/.test(p[0])||!p[1]||/[\/\\]/.test(p[1])||seen[p[0]])throw new Error('Kod/nama panitia tidak sah atau bertindih.');
    seen[p[0]]=true;
  });
  list=list.concat(extra);
  if(new Set(list.map(function(p){return p[1];})).size!==list.length)throw new Error('Nama folder panitia bertindih.');
  return list;
}
function structure_(s) {
  var out=[['KURIKULUM','01_KURIKULUM'],['RPH','01_KURIKULUM/e-RPH'],['RPH_DOCX','01_KURIKULUM/e-RPH/RPH_Word_Docx'],['RPH_PDF','01_KURIKULUM/e-RPH/Arkib_PDF'],['OPR_KURIKULUM','01_KURIKULUM/OPR_Kurikulum'],['PANITIA','01_KURIKULUM/e-Panitia']];
  panitia_(s).forEach(function(p){var base='01_KURIKULUM/e-Panitia/'+p[1];out.push(['PANITIA_'+p[0],base]);Object.keys(PANITIA_DOC_).forEach(function(k){out.push(['PANITIA_'+p[0]+'_'+k,base+'/'+PANITIA_DOC_[k]]);});});
  out.push(['HEM','02_HAL_EHWAL_MURID_HEM'],['OPR_HEM','02_HAL_EHWAL_MURID_HEM/OPR_HEM'],['KOKURIKULUM','03_KOKURIKULUM']);
  Object.keys(KOKO_DIR_).forEach(function(k){out.push(['OPR_KOKO_'+k,'03_KOKURIKULUM/'+KOKO_DIR_[k]]);});
  if(yes_(s.Ada_PPKI))out.push(['PPKI','04_PPKI'],['RPI','04_PPKI/e-RPI_RPH_Inklusif'],['OPR_PPKI','04_PPKI/OPR_PPKI']);
  return out;
}
function driveGet_(id) {return Drive.Files.get(id,{supportsAllDrives:true,fields:'id,name,mimeType,parents,trashed,driveId,capabilities(canAddChildren),appProperties,webViewLink'});}
function driveList_(q) {
  var result=[],token;
  do {var p=Drive.Files.list({q:q,pageSize:1000,pageToken:token,spaces:'drive',includeItemsFromAllDrives:true,supportsAllDrives:true,fields:'nextPageToken,files(id,name,mimeType,parents,trashed,appProperties,webViewLink)'});result=result.concat(p.files||[]);token=p.nextPageToken;}while(token);
  return result;
}
function q_(s) {return String(s).replace(/\\/g,'\\\\').replace(/'/g,"\\'");}
function inRoot_(id,root) {
  var seen={},pending=[id],count=0;
  while(pending.length) {
    var current=pending.pop();if(current===root)return true;
    if(seen[current])continue;seen[current]=true;
    if(++count>100)throw new Error('Hierarki Drive terlalu dalam.');
    var f=driveGet_(current);if(f.trashed)continue;
    pending=pending.concat(f.parents||[]);
  }
  return false;
}
function validateRoot_(s) {
  if(!s.Folder_Drive_ID||!s.Spreadsheet_ID)throw new Error('Folder_Drive_ID dan Spreadsheet_ID diperlukan.');
  if(s.Spreadsheet_ID===prop_('MASTER_SPREADSHEET_ID'))throw new Error('Spreadsheet operasi tidak boleh menggunakan daftar pusat.');
  var root=driveGet_(s.Folder_Drive_ID);
  if(root.trashed||root.mimeType!=='application/vnd.google-apps.folder'||!root.capabilities||!root.capabilities.canAddChildren)throw new Error('Folder akar tiada kebenaran menambah fail.');
  sekolahRows_().filter(function(r){return r.Sekolah_ID!==s.Sekolah_ID;}).forEach(function(other){
    if(other.Spreadsheet_ID===s.Spreadsheet_ID)throw new Error('Spreadsheet sudah digunakan sekolah lain.');
    if(other.Folder_Drive_ID) {
      if(inRoot_(root.id,other.Folder_Drive_ID)||inRoot_(other.Folder_Drive_ID,root.id))throw new Error('Folder akar bertindih dengan sekolah lain.');
    }
  });
  SpreadsheetApp.openById(s.Spreadsheet_ID).getName();
  return root;
}
// Caller mesti memegang ScriptLock. Carian hanya di bawah parent yang disahkan.
function dapatkanAtauCiptaFolder_(parentFolder,folderName) {
  var parentId=typeof parentFolder==='string'?parentFolder:parentFolder.id;
  var matches=driveList_("'"+q_(parentId)+"' in parents and trashed = false and mimeType = 'application/vnd.google-apps.folder' and name = '"+q_(folderName)+"'");
  if(matches.length>1)throw new Error('Folder pendua memerlukan semakan: '+folderName);
  if(matches.length===1)return matches[0];
  return Drive.Files.create({name:folderName,mimeType:'application/vnd.google-apps.folder',parents:[parentId]},null,{supportsAllDrives:true,fields:'id,name,parents'});
}
function folderIndex_() {return tab_('Direktori_Sekolah',['Sekolah_ID','Root_ID','Kod_Laluan','Path','Folder_ID','Versi']);}
function pathFolder_(s,path,key) {
  var sh=folderIndex_(), indexed=rows_(sh).filter(function(r){return r.Sekolah_ID===s.Sekolah_ID&&r.Root_ID===s.Folder_Drive_ID&&r.Kod_Laluan===key;});
  if(indexed.length>1)throw new Error('Indeks direktori bertindih.');
  if(indexed.length) {
    var existing=driveGet_(indexed[0].Folder_ID);
    if(existing.trashed||existing.mimeType!=='application/vnd.google-apps.folder'||!inRoot_(existing.id,s.Folder_Drive_ID))throw new Error('Folder berdaftar dipadam/dipindahkan. Semakan diperlukan: '+key);
    var expectedParts=path.split('/'),current=existing;
    for(var p=expectedParts.length-1;p>=0;p--) {
      if(current.name!==expectedParts[p]||(current.parents||[]).length!==1)throw new Error('Nama/hierarki folder berdaftar berubah: '+key);
      var siblings=driveList_("'"+q_(current.parents[0])+"' in parents and trashed = false and mimeType = 'application/vnd.google-apps.folder' and name = '"+q_(current.name)+"'");
      if(siblings.length!==1)throw new Error('Folder pendua memerlukan semakan: '+key);
      current=p===0?{id:current.parents[0]}:driveGet_(current.parents[0]);
    }
    if(current.id!==s.Folder_Drive_ID)throw new Error('Folder tidak lagi berada di laluan berdaftar.');
    return existing.id;
  }
  var parent=s.Folder_Drive_ID,parts=path.split('/');
  // Traverse untuk mengesan pindahan, rename dan folder pendua; jangan percaya ID cache semata-mata.
  parts.forEach(function(name){parent=dapatkanAtauCiptaFolder_(parent,name).id;});
  if(indexed.length && indexed[0].Folder_ID!==parent)throw new Error('Folder berdaftar berubah. Semak indeks sebelum membaiki: '+key);
  if(!indexed.length)sh.appendRow([s.Sekolah_ID,s.Folder_Drive_ID,key,path,parent,STRUCTURE_VERSION_]);
  return parent;
}
function provision_(s,deadline) {
  validateRoot_(s);var plan=structure_(s);
  var progressKey='provision:'+hash_(s.Sekolah_ID+'|'+s.Folder_Drive_ID+'|'+JSON.stringify(plan));
  var props=PropertiesService.getScriptProperties(),start=Number(props.getProperty(progressKey)||0);
  statusSekolah_(s,'PROVISIONING','');
  for(var i=start;i<plan.length;i++) {
    if(deadline&&Date.now()>deadline){statusSekolah_(s,'PENDING','Sambung pada worker seterusnya.');return false;}
    pathFolder_(s,plan[i][1],plan[i][0]);
    props.setProperty(progressKey,String(i+1));
  }
  props.deleteProperty(progressKey);
  statusSekolah_(s,'READY','');return true;
}
// Fungsi editor, bukan endpoint web. Tetapkan TARGET_SCHOOL_ID sebelum menjalankannya.
function sediakanDaftarSaas_() {
  operator_();lock_(function(){tab_('Langganan_Sekolah',SCHOOL_HEADERS_);folderIndex_();artifactIndex_();authTable_();});
}
function provisionSekolahEditor_() {
  operator_();return lock_(function(){var s=sekolah_(prop_('TARGET_SCHOOL_ID'));try{return provision_(s,Date.now()+240000);}catch(e){statusSekolah_(s,'ERROR',e.message);throw e;}});
}
function workerProvisioning_() {
  return lock_(function(){var deadline=Date.now()+45000;sekolahRows_().forEach(function(r){
    if(Date.now()>deadline||!['','PENDING','PROVISIONING'].includes(norm_(r.Status_Provisioning)))return;
    try {provision_(sekolah_(r.Sekolah_ID),deadline);}catch(e){statusSekolah_(r,'ERROR',e.message);}
  });});
}
function daftarBerubah_(e) {
  if(!e||!e.range||e.source.getId()!==prop_('MASTER_SPREADSHEET_ID')||e.range.getSheet().getName()!=='Langganan_Sekolah'||e.range.getRow()<2)return;
  if(e.range.getColumn()>8 && e.range.getColumn()<13)return;
  lock_(function(){for(var r=e.range.getRow();r<=e.range.getLastRow();r++)e.range.getSheet().getRange(r,9).setValue('PENDING');});
}
function pasangAutomasiSaas_() {
  operator_();var own=['workerProvisioning_','daftarBerubah_'];
  ScriptApp.getProjectTriggers().forEach(function(t){if(own.includes(t.getHandlerFunction()))ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('workerProvisioning_').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('daftarBerubah_').forSpreadsheet(prop_('MASTER_SPREADSHEET_ID')).onEdit().create();
}

function routeKey_(s,m) {
  m=m||{};var type=norm_(m.jenisDokumen).toUpperCase(),section=norm_(m.bahagian).toUpperCase();
  if(type==='RPI'||type==='RPH_INKLUSIF'||section==='PPKI') {if(!yes_(s.Ada_PPKI))throw new Error('Sekolah tidak mengaktifkan PPKI.');if(!['RPI','RPH','RPH_INKLUSIF','OPR'].includes(type))throw new Error('Jenis dokumen PPKI tidak sah.');return type==='OPR'?'OPR_PPKI':'RPI';}
  if(type==='RPH'){if(!['PDF','DOCX'].includes(m.format))throw new Error('Format RPH tidak sah.');return m.format==='DOCX'?'RPH_DOCX':'RPH_PDF';}
  if(m.panitiaKod) {
    if(!panitia_(s).some(function(p){return p[0]===m.panitiaKod;})||!PANITIA_DOC_[type])throw new Error('Panitia/jenis dokumen tidak sah.');
    if(section!=='KURIKULUM')throw new Error('Dokumen panitia mesti di bahagian Kurikulum.');
    return 'PANITIA_'+m.panitiaKod+'_'+type;
  }
  if(type==='OPR'&&section==='KURIKULUM')return 'OPR_KURIKULUM';
  if(type==='OPR'&&section==='HEM')return 'OPR_HEM';
  if(type==='OPR'&&section==='KOKURIKULUM') {
    var category=m.kategoriKoko;
    if(!category)category=jsonConfig_(s.Unit_Koko_JSON,{})[norm_(m.unit)];
    if(!KOKO_DIR_[category])throw new Error('Pilih kategori Kokurikulum: SUKAN, KELAB atau BERUNIFORM.');
    return 'OPR_KOKO_'+category;
  }
  if(type==='MEDIA_PORTAL')return 'MEDIA_PORTAL';
  throw new Error('Jenis dokumen/bahagian belum ditentukan.');
}
function routeFolder_(m) {
  var s=konteks_().school;validateRoot_(s);var key=routeKey_(s,m);
  var plan=structure_(s), match=plan.filter(function(p){return p[0]===key;})[0];
  var path=key==='MEDIA_PORTAL'?'05_SISTEM/Media_Portal':match&&match[1];
  if(!path)throw new Error('Laluan tidak tersedia.');
  if(m.idLaporan)path+='/'+safeName_(m.idLaporan);
  if(m.tahun)path+='/'+safeName_(m.tahun);
  if(m.guruId)path+='/'+safeName_(m.guruId);
  return pathFolder_(s,path,key+':'+path);
}
function safeName_(name) {var s=norm_(name).replace(/[\\/\x00-\x1f]/g,'_').slice(0,140);if(!s||s==='.'||s==='..')throw new Error('Nama fail/folder tidak sah.');return s;}
function artifactIndex_() {return tab_('Fail_Sekolah',['Sekolah_ID','Root_ID','Kunci','File_ID','Folder_ID','Jenis','Nama','Tarikh']);}
function saveBlob_(blob,meta) {
  var s=konteks_().school,m=meta||ROUTE_;
  if(!m)throw new Error('Metadata destinasi fail diperlukan.');
  var parent=routeFolder_(m),bytes=blob.getBytes();
  var contentHash=Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,bytes));
  var key=hash_(s.Sekolah_ID+'|'+parent+'|'+blob.getName()+'|'+contentHash);
  var sh=artifactIndex_(),old=rows_(sh).filter(function(r){return r.Sekolah_ID===s.Sekolah_ID&&r.Root_ID===s.Folder_Drive_ID&&r.Kunci===key;});
  if(old.length>1)throw new Error('Indeks fail bertindih.');
  var f;
  if(old.length){f=driveGet_(old[0].File_ID);if(f.trashed||!(f.parents||[]).includes(parent))throw new Error('Fail arkib dipadam/dipindahkan. Semakan diperlukan.');}
  else {
    var found=driveList_("'"+q_(parent)+"' in parents and trashed = false and appProperties has { key='artifactKey' and value='"+key+"' }");
    if(found.length>1)throw new Error('Fail pendua dikesan.');
    f=found[0]||Drive.Files.create({name:safeName_(blob.getName()),parents:[parent],appProperties:{artifactKey:key,sekolahId:String(s.Sekolah_ID)}},blob,{supportsAllDrives:true,fields:'id,name,parents,webViewLink'});
    sh.appendRow([s.Sekolah_ID,s.Folder_Drive_ID,key,f.id,parent,m.jenisDokumen,blob.getName(),new Date()]);
  }
  // Adapter untuk penjana legasi; semua ciptaan sebenar melalui Advanced Drive.
  return {getId:function(){return f.id;},getUrl:function(){return 'https://drive.google.com/file/d/'+f.id+'/view';}};
}
function scopedFile_(id) {
  var s=konteks_().school;
  if(!inRoot_(id,s.Folder_Drive_ID))throw new Error('Fail di luar folder sekolah.');
  var f=driveGet_(id);if(f.trashed)throw new Error('Fail berada dalam sampah.');
  return DriveApp.getFileById(id);
}
function imageBlob_(data,name) {
  var match=String(data||'').match(/^data:(image\/(?:jpeg|png|gif));base64,([A-Za-z0-9+/=\r\n]+)$/);
  if(!match||match[2].length>11000000)throw new Error('Imej tidak sah atau melebihi 8 MB.');
  return Utilities.newBlob(Utilities.base64Decode(match[2]),match[1],safeName_(name).replace(/\.(jpg|jpeg|png|gif)$/i,'')+'.'+({ 'image/jpeg':'jpg','image/png':'png','image/gif':'gif'}[match[1]]));
}
function simpanMedia_(data,name) {
  if(!data)return '';
  if(!String(data).startsWith('data:')) {var id=ambilFileIdDariUrl_(data);scopedFile_(id);return data;}
  return saveBlob_(imageBlob_(data,name),ROUTE_||{jenisDokumen:'MEDIA_PORTAL'}).getUrl();
}
function metaOpr_(form) {
  var s=konteks_().school;
  var saved=metaTable_(),found=rows_(saved).filter(function(r){return r.Sekolah_ID===s.Sekolah_ID&&r.ID_Laporan===form.idLaporan;});
  var m=found.length?JSON.parse(found[0].Metadata):{jenisDokumen:'OPR',bahagian:form.bahagian||'KOKURIKULUM',panitiaKod:form.panitiaKod||'',kategoriKoko:form.kategoriKoko||'',unit:form.unit||form.unitPanitia||'',owner:konteks_().user.emel};
  if(found.length)['bahagian','panitiaKod','kategoriKoko'].forEach(function(key){
    if(Object.prototype.hasOwnProperty.call(form,key)&&norm_(form[key])!==norm_(m[key]))throw new Error('Kategori laporan sedia ada tidak boleh diubah. Cipta laporan baharu untuk destinasi lain.');
  });
  m.idLaporan=form.idLaporan||('OPR-'+Utilities.getUuid());form.idLaporan=m.idLaporan;
  routeKey_(s,m);
  if(!found.length)saved.appendRow([s.Sekolah_ID,m.idLaporan,JSON.stringify(m)]);
  return m;
}
function metaTable_(){return tab_('Metadata_Dokumen',['Sekolah_ID','ID_Laporan','Metadata']);}


// ===== Auth.gs =====
// OTP emel; tiada kata laluan atau senarai pengguna dihantar sebelum log masuk.
function authTable_() {return tab_('Auth_State',['Key','Value','Expires']);}
function authGet_(key) {
  var sh=authTable_(), r=rows_(sh).filter(function(x){return x.Key===key;})[0];
  if(!r||Number(r.Expires)<Date.now())return null;
  return JSON.parse(r.Value);
}
function authPut_(key,value,expires) {
  var sh=authTable_(),all=rows_(sh),r=all.filter(function(x){return x.Key===key;})[0];
  if(!r)r=all.filter(function(x){return Number(x.Expires)<Date.now();})[0];
  var values=[key,JSON.stringify(value),expires];
  if(r)sh.getRange(r._row,1,1,3).setValues([values]);else sh.appendRow(values);
}
function users_(s) {
  var sh=SpreadsheetApp.openById(s.Spreadsheet_ID).getSheetByName('PENGGUNA');
  if(!sh)throw new Error('Tab PENGGUNA belum disediakan.');
  var data=sh.getDataRange().getDisplayValues(),heads=data.shift().map(function(h){return norm_(h).toUpperCase().replace(/[ -]/g,'_');});
  function col(names,required){var i=-1;names.some(function(n){i=heads.indexOf(n);return i>=0;});if(i<0&&required)throw new Error('Header PENGGUNA diperlukan: '+names[0]);return i;}
  var e=col(['EMEL','EMAIL','EMAIL_DLIMA'],true),n=col(['NAMA_GURU','NAMA'],true),p=col(['PERANAN'],true),j=col(['JAWATAN']),review=col(['PENYEMAK_EMAIL','PENYEMAK_EMEL','PENYEMAK']),foto=col(['FOTO','GAMBAR','PHOTO']),status=col(['STATUS']);
  var seen={};
  return data.filter(function(r){return r[e];}).map(function(r){
    var emel=email_(r[e]);if(seen[emel])throw new Error('Emel pengguna bertindih.');seen[emel]=true;
    var role=norm_(r[p]).toUpperCase(),rbac=roleMap_(role);
    return {emel:emel,nama:r[n],jawatan:r[j]||'',peranan:role,penyemakEmail:r[review]||'',foto:r[foto]||'',aktif:status<0||!r[status]||norm_(r[status]).toUpperCase()==='AKTIF',rbac:rbac};
  });
}
function roleMap_(role) {
  var aliases={'PK PENTADBIRAN':'PK1','PENOLONG KANAN PENTADBIRAN':'PK1','PENOLONG KANAN HEM':'PK HEM','PENOLONG KANAN KOKURIKULUM':'PK KOKURIKULUM','PENOLONG KANAN PPKI':'PK PPKI'};
  role=aliases[role]||role;
  var map={ADMIN:['ADMIN',['KURIKULUM','HEM','KOKURIKULUM','PPKI']],PENTADBIR:['ADMIN',['KURIKULUM','HEM','KOKURIKULUM','PPKI']],GB:['GURU_BESAR',['KURIKULUM','HEM','KOKURIKULUM','PPKI']],'GURU BESAR':['GURU_BESAR',['KURIKULUM','HEM','KOKURIKULUM','PPKI']],PK1:['PK1',['KURIKULUM']],'PK 1':['PK1',['KURIKULUM']],'PK HEM':['PK_HEM',['HEM']],PK_HEM:['PK_HEM',['HEM']],'PK KOKUM':['PK_KOKUM',['KOKURIKULUM']],'PK KOKU':['PK_KOKUM',['KOKURIKULUM']],'PK KOKURIKULUM':['PK_KOKUM',['KOKURIKULUM']],PK_KOKUM:['PK_KOKUM',['KOKURIKULUM']],'PK PPKI':['PK_PPKI',['PPKI']],PK_PPKI:['PK_PPKI',['PPKI']]};
  var m=map[role]||['GURU',[]];return {perananKod:m[0],label:role||'GURU',skop:m[1],bolehLulus:m[1].length>0};
}
function portalMintaKod(sekolahId,emel) {
  return lock_(function(){
    var s=sekolah_(sekolahId),e=email_(emel);
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)||e.length>200)throw new Error('Emel tidak sah.');
    var now=Date.now(),key='otp:'+hash_(s.Sekolah_ID+'|'+e),last=authGet_(key);
    if(last&&now-last.sent<60000)throw new Error('Tunggu satu minit sebelum meminta kod baharu.');
    var bucket='rate:'+hash_(s.Sekolah_ID+'|'+Math.floor(now/3600000)),rate=authGet_(bucket)||{count:0};
    if(rate.count>=20)throw new Error('Had permintaan kod sekolah dicapai. Cuba kemudian.');
    rate.count++;authPut_(bucket,rate,now+3600000);
    var user=users_(s).filter(function(u){return u.emel===e&&u.aktif;})[0];
    var challenge=unique_(),code=String(parseInt(hash_(unique_()).slice(0,12),16)%1000000).padStart(6,'0');
    authPut_(key,{challenge:challenge,hash:hash_(challenge+'|'+code),attempts:0,sent:now,valid:!!user},now+600000);
    if(user) {
      if(MailApp.getRemainingDailyQuota()<1)throw new Error('Kuota emel habis. Hubungi operator.');
      MailApp.sendEmail({to:e,subject:'Kod log masuk '+s.Nama_Sekolah,body:'Kod sekali guna anda: '+code+'\nSah selama 10 minit. Jangan kongsikan kod ini. Abaikan jika anda tidak memintanya.'});
    }
    return {challenge:challenge,message:'Jika emel berdaftar dan aktif, kod telah dihantar. Semak peti masuk dan spam.'};
  });
}
function portalSahkanKod(sekolahId,emel,challenge,kod) {
  return lock_(function(){
    var s=sekolah_(sekolahId),e=email_(emel),key='otp:'+hash_(s.Sekolah_ID+'|'+e),v=authGet_(key);
    if(!v||v.challenge!==challenge||v.attempts>=5)throw new Error('Kod tidak sah/tamat. Minta kod baharu.');
    v.attempts++;authPut_(key,v,v.sent+600000);
    if(!v.valid||!/^[0-9]{6}$/.test(String(kod))||hash_(challenge+'|'+kod)!==v.hash)throw new Error('Kod tidak sah.');
    var user=users_(s).filter(function(u){return u.emel===e&&u.aktif;})[0];if(!user)throw new Error('Pengguna tidak aktif.');
    authPut_(key,{used:true},Date.now()-1);
    var token=unique_(),expires=Date.now()+21600000;
    authPut_('session:'+hash_(token),{school:s.Sekolah_ID,email:e,root:s.Folder_Drive_ID,spreadsheet:s.Spreadsheet_ID},expires);
    return {token:token,expires:expires,guru:user,sekolah:{id:s.Sekolah_ID,nama:s.Nama_Sekolah}};
  });
}
function session_(token) {
  if(typeof token!=='string'||!/^[a-f0-9]{64}$/i.test(token))throw new Error('Sila log masuk.');
  var v=authGet_('session:'+hash_(token));if(!v)throw new Error('Sesi tamat. Sila log masuk semula.');
  var s=sekolah_(v.school);
  if(v.root!==s.Folder_Drive_ID||v.spreadsheet!==s.Spreadsheet_ID)throw new Error('Konfigurasi sekolah berubah. Log masuk semula.');
  var u=users_(s).filter(function(x){return x.emel===v.email&&x.aktif;})[0];if(!u)throw new Error('Akses pengguna telah dibatalkan.');
  return {school:s,user:u};
}
function portalLogKeluar(token) {return lock_(function(){if(typeof token==='string')authPut_('session:'+hash_(token),{},Date.now()-1);return true;});}
function ownEmail_(email,adminAllowed) {
  var u=konteks_().user;
  if(email_(email)!==u.emel&&!(adminAllowed&&u.rbac.bolehLulus))throw new Error('Akses rekod guru lain tidak dibenarkan.');
}
function admin_(section,superOnly) {
  var r=konteks_().user.rbac;
  if(superOnly&&!['ADMIN','GURU_BESAR'].includes(r.perananKod))throw new Error('Akses pentadbir sistem diperlukan.');
  if(!r.bolehLulus||(section&&!r.skop.includes(section)))throw new Error('Tiada hak pentadbir untuk bahagian ini.');
}
function ownedRow_(sheet,id,emailCol,allowAdmin) {
  var sh=sekolahSpreadsheet_().getSheetByName(sheet);if(!sh)throw new Error('Rekod tidak ditemui.');
  var rows=sh.getDataRange().getDisplayValues().slice(1),r=rows.filter(function(x){return String(x[0])===String(id);})[0];
  if(!r)throw new Error('Rekod tidak ditemui.');ownEmail_(r[emailCol],allowAdmin);return r;
}
function policy_(method,args) {
  var u=konteks_().user,p=args[0]||{};
  var adminMethods=['dapatkanDataDashboardPentadbir','simpanSemakanPentadbir','dapatkanSemuaKehadiranHariIni','dapatkanSenaraiOprPentadbir','sahkanLaporanOprBackend','dapatkanSemuaRekodKeberhasilanAdmin'];
  if(adminMethods.includes(method))admin_();
  if(['tambahPenggunaBaru','kemaskiniMaklumatGuru'].includes(method))admin_('',true);
  if(['dapatkanDataDashboardPentadbir','simpanSemakanPentadbir'].includes(method))admin_('KURIKULUM');
  var ownAt={simpanFotoProfilGuru:0,semakAdaRekod:1,dapatkanRekodMinggu:1,janaPdfMingguanBackend:1,simpanLaporanBertugasBackend:0,padamRphMingguanBackend:0,simpanJadualGuruBackend:0,dapatkanJadualGuruBackend:0,dapatkanStatusKehadiranHariIni:0,getFeedStatusWeb:0,hantarStatusFeed:0,toggleLikeStatusFeed:1,padamStatusFeed:1,dapatkanBilanganOnlineLive:0,muatRekodKeberhasilanGuru:0,dapatkanSenaraiArkibRphGuru:0,semakStatusMingguTertunggak:1,janaRphBulanan:0,janaPdfRumusanBertugasMingguan:2};
  if(Object.prototype.hasOwnProperty.call(ownAt,method))ownEmail_(args[ownAt[method]],['janaPdfMingguanBackend','dapatkanRekodMinggu','muatRekodKeberhasilanGuru','dapatkanSenaraiArkibRphGuru'].includes(method));
  if(['janaRphSemingguBackend','simpanRekodKeberhasilan','simpanLaporanBertugasLengkapBackend','rakamKehadiranGpsBackend'].includes(method)){ownEmail_(p.emel);p.nama=u.nama;p.namaGuru=u.nama;p.jawatan=u.jawatan;}
  if(method==='kemaskiniRefleksi')ownedRow_('RPH_GURU',args[0],1,false);
  if(method==='padamStatusFeed')ownedRow_('STATUS_FEED',args[0],3,false);
  if(method==='dapatkanSenaraiOprPentadbir')args[0]=u.emel;
  if(method==='sahkanLaporanOprBackend'){args[1]=u.emel;if(!['Disahkan','Perlu Pindaan'].includes(args[2]))throw new Error('Status pengesahan tidak sah.');}
  if(method==='dapatkanBankOprSekolah'){
    if(p.emelGuru)ownEmail_(p.emelGuru,false);
    else if(!u.rbac.bolehLulus)p.hanyaDisahkan=true;
    args[0]=p;
  }
  if(method==='dapatkanGambarLaporanOprBase64') {
    var sh=sekolahSpreadsheet_().getSheetByName('LAPORAN_OPR');var r=sh&&sh.getDataRange().getDisplayValues().slice(1).filter(function(x){return x[0]===args[0];})[0];
    if(!r)throw new Error('Laporan tidak ditemui.');
    if(r[19]!=='Disahkan'&&email_(r[5])!==u.emel)admin_(r[1]);
  }
  if(['simpanAtauKemasKiniOpr','simpanLaporanOprKokumBackend','janaPdfOprKokumBackendPortal'].includes(method)) {
    if(p.idLaporan) {
      var oprSh=sekolahSpreadsheet_().getSheetByName('LAPORAN_OPR');
      var oprRow=oprSh&&oprSh.getDataRange().getDisplayValues().slice(1).filter(function(r){return r[0]===p.idLaporan;})[0];
      if(oprRow)ownEmail_(oprRow[5],false);
      else {
        if(!/^OPR-[a-f0-9-]{36}$/i.test(p.idLaporan))throw new Error('ID laporan baharu tidak sah.');
        var pending=rows_(metaTable_()).filter(function(r){return r.Sekolah_ID===konteks_().school.Sekolah_ID&&r.ID_Laporan===p.idLaporan;})[0];
        if(pending)ownEmail_(JSON.parse(pending.Metadata).owner,false);
      }
    }
    p.emelPenyelaras=u.emel;p.emelGuru=u.emel;p.penyelaras=u.nama;
    ROUTE_=metaOpr_(p);
  }
  if(method==='simpanFotoProfilGuru')imageBlob_(args[1],'profil');
  if(method==='getSenaraiTabKokumPpki'&&!yes_(konteks_().school.Ada_PPKI))throw new Error('PPKI tidak aktif.');
  if(['janaPdfMingguanBackend','janaRphBulanan'].includes(method)) {
    var owner=method==='janaRphBulanan'?args[0]:args[1];
    ROUTE_={jenisDokumen:'RPH',format:'PDF',tahun:method==='janaRphBulanan'?String(args[2]):Utilities.formatDate(new Date(),'Asia/Kuala_Lumpur','yyyy'),guruId:hash_(owner).slice(0,12)};
  }
  // Cegah pembacaan helaian bukan Kokurikulum melalui parameter tabName.
  if(['getSenaraiMurid','simpanKehadiranKoko'].includes(method)) {
    var sh=sekolahSpreadsheet_().getSheetByName(String(args[0]));
    var units=jsonConfig_(konteks_().school.Unit_Koko_JSON,{});
    if(!sh||!Object.prototype.hasOwnProperty.call(units,String(args[0]))||!KOKO_DIR_[units[String(args[0])]])throw new Error('Tab belum didaftarkan dalam Unit_Koko_JSON sekolah.');
    if(/PPKI/i.test(sh.getName())&&!yes_(konteks_().school.Ada_PPKI))throw new Error('PPKI tidak aktif.');
  }
}
function portalRpc(token,method,args) {
  if(!Object.prototype.hasOwnProperty.call(RPC_,method))throw new Error('Operasi tidak dibenarkan.');
  if(!Array.isArray(args)||args.length>15)throw new Error('Parameter tidak sah.');
  return lock_(function(){
    CTX_=session_(token);ROUTE_=null;
    try {policy_(method,args);var result=RPC_[method].apply(null,args);
      if(['dapatkanBankOprSekolah','dapatkanSenaraiOprPentadbir'].includes(method)&&Array.isArray(result)) {
        var metas=rows_(metaTable_()).filter(function(r){return r.Sekolah_ID===CTX_.school.Sekolah_ID;});
        result.forEach(function(item){var found=metas.filter(function(r){return r.ID_Laporan===item.idLaporan;})[0];if(found){var m=JSON.parse(found.Metadata);item.panitiaKod=m.panitiaKod||'';item.kategoriKoko=m.kategoriKoko||'';}});
      }
      return result;
    }
    finally {CTX_=null;ROUTE_=null;}
  });
}
function portalBootstrap(token) {
  return lock_(function(){CTX_=session_(token);try{
    var s=CTX_.school, dirs=rows_(folderIndex_()).filter(function(r){return r.Sekolah_ID===s.Sekolah_ID&&r.Root_ID===s.Folder_Drive_ID;});
    return {guru:CTX_.user,sekolah:{id:s.Sekolah_ID,nama:s.Nama_Sekolah,kod:s.Kod_Sekolah,alamat:s.Alamat,adaPpki:yes_(s.Ada_PPKI),status:s.Status_Provisioning,gps:{lat:s.GPS_Lat,lon:s.GPS_Lon,radiusMeter:s.GPS_Radius,nama:s.Nama_Sekolah}},panitia:panitia_(s),direktori:dirs.map(function(r){return {kod:r.Kod_Laluan,path:r.Path,url:'https://drive.google.com/drive/folders/'+r.Folder_ID};})};
  }finally{CTX_=null;}});
}
function portalMuatNaik(token,meta,dataUrl,namaFail) {
  return lock_(function(){CTX_=session_(token);try{
    if(!meta||!['RPH','RPI','RPH_INKLUSIF','OPR','DASAR','MESYUARAT','PENTAKSIRAN','BBM'].includes(meta.jenisDokumen))throw new Error('Jenis dokumen tidak sah.');
    var data=String(dataUrl||'').match(/^data:([^;,]+);base64,([A-Za-z0-9+/=\r\n]+)$/);
    if(!data||data[2].length>11000000)throw new Error('Fail tidak sah atau melebihi 8 MB.');
    var allowed=['application/pdf','application/vnd.openxmlformats-officedocument.wordprocessingml.document'];
    if(!allowed.includes(data[1]))throw new Error('Hanya PDF atau DOCX dibenarkan.');
    var bytes=Utilities.base64Decode(data[2]),blob=Utilities.newBlob(bytes,data[1],safeName_(namaFail));
    var docx=data[1]===allowed[1];
    if(!(docx?/\.docx$/i:/\.pdf$/i).test(namaFail))throw new Error('Sambungan nama fail tidak sepadan dengan format.');
    if(docx) {
      var zip=Utilities.unzip(Utilities.newBlob(bytes,'application/zip','file.zip'));
      if(!zip.some(function(f){return f.getName()==='word/document.xml';}))throw new Error('Fail bukan DOCX sebenar.');
    }else if(String.fromCharCode.apply(null,bytes.slice(0,5))!=='%PDF-')throw new Error('Fail bukan PDF sebenar.');
    meta.format=docx?'DOCX':'PDF';
    if(meta.jenisDokumen==='RPH'){meta.guruId=hash_(CTX_.user.emel).slice(0,12);meta.tahun=Utilities.formatDate(new Date(),'Asia/Kuala_Lumpur','yyyy');}
    var file=saveBlob_(blob,meta);return {status:'SUCCESS',url:file.getUrl()};
  }finally{CTX_=null;}});
}


// ===== Integration.gs =====
function brandHtml_(html) {
  var s=konteks_().school;
  return String(html).replaceAll('__NAMA_SEKOLAH__',escapeHtmlGas_(s.Nama_Sekolah)).replaceAll('__ALAMAT_SEKOLAH__',escapeHtmlGas_(s.Alamat||'')).replaceAll('__KOD_SEKOLAH__',escapeHtmlGas_(s.Kod_Sekolah||''));
}
function gpsSekolah_() {
  var s=konteks_().school;
  if(s.GPS_Lat===''||s.GPS_Lon===''||!Number.isFinite(Number(s.GPS_Lat))||!Number.isFinite(Number(s.GPS_Lon))||!(Number(s.GPS_Radius)>0))throw new Error('Koordinat GPS sekolah belum dikonfigurasi.');
  return {lat:Number(s.GPS_Lat),lon:Number(s.GPS_Lon),radiusMeter:Number(s.GPS_Radius),namaLokasi:s.Nama_Sekolah};
}
function configuredKoko_(ppki) {
  var s=konteks_().school,result={BERUNIFORM:[],KELAB:[],'1M1S':[]};
  if(ppki&&!yes_(s.Ada_PPKI))return result;
  var config=jsonConfig_(s.Unit_Koko_JSON,{});
  Object.keys(config).forEach(function(name){
    if(!KOKO_DIR_[config[name]])throw new Error('Kategori Unit_Koko_JSON tidak sah: '+name);
    if(/PPKI/i.test(name)!==ppki)return;
    if(!sekolahSpreadsheet_().getSheetByName(name))throw new Error('Tab unit belum disediakan: '+name);
    var key=config[name]==='SUKAN'?'1M1S':config[name];
    result[key].push({id:name,nama:name,warna:key==='BERUNIFORM'?'Kuning':key==='KELAB'?'Hijau':'Oren'});
  });
  return result;
}
function manageUser_(data,update) {
  admin_('',true);
  var e=email_(data.emel||data.email),name=norm_(data.nama||data.namaGuru);
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)||!name)throw new Error('Nama dan emel diperlukan.');
  var sh=sekolahSpreadsheet_().getSheetByName('PENGGUNA');
  if(!sh)throw new Error('Sediakan PENGGUNA mengikut panduan pemasangan.');
  var values=sh.getDataRange().getDisplayValues(),headers=values[0].map(function(h){return norm_(h).toUpperCase().replace(/[ -]/g,'_');});
  var eCol=headers.findIndex(function(h){return ['EMAIL','EMEL','EMAIL_DLIMA'].includes(h);});
  if(eCol<0)throw new Error('Header EMEL tiada.');
  var row=values.findIndex(function(r,i){return i>0&&email_(r[eCol])===e;});
  if(update&&row<1)throw new Error('Pengguna tidak ditemui.');
  if(!update&&row>0)throw new Error('Emel sudah berdaftar.');
  var record=row>0?values[row]:headers.map(function(){return '';});
  var fields={EMEL:e,EMAIL:e,EMAIL_DLIMA:e,NAMA:name,NAMA_GURU:name,JAWATAN:norm_(data.jawatan),PERANAN:norm_(data.peranan||'GURU').toUpperCase(),PENYEMAK_EMAIL:email_(data.penyemakEmail),PENYEMAK_EMEL:email_(data.penyemakEmail)};
  headers.forEach(function(h,i){if(Object.prototype.hasOwnProperty.call(fields,h))record[i]=fields[h];if(['KATALALUAN','PASSWORD','PASSCODE','KATA_LALUAN'].includes(h))record[i]='';});
  if(row>0)sh.getRange(row+1,1,1,record.length).setValues([record]);else sh.appendRow(record);
  return {success:true,message:'Pengguna disimpan. Log masuk menggunakan kod emel sekali guna.'};
}
function pasangCheckoutSaas_() {
  operator_();ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='checkoutSemuaSekolah_')ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('checkoutSemuaSekolah_').timeBased().atHour(17).everyDays(1).inTimezone('Asia/Kuala_Lumpur').create();
}
function checkoutSemuaSekolah_() {
  sekolahRows_().forEach(function(r){try{lock_(function(){CTX_={school:sekolah_(r.Sekolah_ID),user:{emel:'',rbac:{bolehLulus:false}}};try{autoCheckoutHarian5PM_();}finally{CTX_=null;}});}catch(e){console.warn('Checkout '+r.Sekolah_ID+': '+e.message);}});
}


// ===== LegacyPortal.gs =====
// ==========================================
// KOD BACKEND KEMASKINI (Code.gs) - PORTAL __NAMA_SEKOLAH__ 2026
// FORMAT RASMI REKOD PENGAJARAN DAN PEMBELAJARAN HARIAN (__NAMA_SEKOLAH__)
// ==========================================



// --------------------------------------------------------------------------
// 1. FUNGSI PEMBERSIHAN MASA & PADANAN DATA
// --------------------------------------------------------------------------

// Pembersihan Format Masa Bersih (HH:MM) - Mengelakkan ralat zon masa epoch Dec 30 1899 (+64 minit)
// TIDAK menggunakan Utilities.formatDate atau objek Date dengan penukaran zon masa
function bersihkanMasa_(val) {
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
function formatKodMinggu_(m) {
  var s = String(m || "").trim();
  var num = s.replace(/\D/g, '');
  if (num) return "M" + num;
  return s.toUpperCase();
}

// Padanan Minggu Fleksibel
function padanMingguSama_(m1, m2) {
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
function padanEmelSama_(e1, e2) {
  if (!e1 || !e2) return false;
  return String(e1).trim().toLowerCase() === String(e2).trim().toLowerCase();
}

// Kira Tarikh Sebenar Bagi Hari (Isnin - Jumaat) Berdasarkan Tarikh Isnin Takwim
function dapatkanTarikhHari_(isninStr, namaHari) {
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
function dapatkanIsninMinggu_(minggu) {
  try {
    var ss = sekolahSpreadsheet_();
    var sheet = ss.getSheetByName("TAKWIM");
    if (!sheet) return "";
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (padanMingguSama_(data[i][0], minggu)) {
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
function formatNamaHari_(hari) {
  var h = String(hari || "").toUpperCase();
  if (h.includes("ISNIN")) return "ISNIN";
  if (h.includes("SELASA")) return "SELASA";
  if (h.includes("RABU")) return "RABU";
  if (h.includes("KHAMIS")) return "KHAMIS";
  if (h.includes("JUMAAT")) return "JUMAAT";
  return h.replace(/^\d+\.\s*/, '');
}

// Format Label Kelas Lengkap: [Nama Kelas (Tahun)]
function formatKelasTahun_(namaKelas, tahun) {
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
function seragamkanNamaSubjek_(subjek) {
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

function tukarDigitArabKeRumi_(teks) {
  if (!teks) return "";
  var str = String(teks);
  var arab = ['٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩'];
  var persia = ['۰', '۱', '۲', '۳', '۴', '۵', '۶', '۷', '۸', '۹'];
  for (var i = 0; i < 10; i++) {
    str = str.replace(new RegExp(arab[i], 'g'), i).replace(new RegExp(persia[i], 'g'), i);
  }
  return str;
}

function padanTahunSama_(rowVal, cariVal, isPpki) {
  if (!cariVal || !rowVal) return true;
  var sRow = String(rowVal).trim().toUpperCase();
  var sCari = String(cariVal).trim().toUpperCase();
  if (sRow === sCari) return true;
  if (sRow.includes("PRA") || sCari.includes("PRA")) return true;
  var dRow = tukarDigitArabKeRumi_(sRow).replace(/\D/g, '');
  var dCari = tukarDigitArabKeRumi_(sCari).replace(/\D/g, '');
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
function muatSemuaDskpDariSheet_(isPpki) {
  try {
    var ss = sekolahSpreadsheet_();
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
function dapatkanDskpPiawaiKpm_(subjek, tahun, isPpki, isPra) {
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

    var digitThn = tukarDigitArabKeRumi_(tahun).replace(/\D/g, '');
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
function getHierarkiDskp_(subjek, tahun, namaKelas) {
  var kUpper = String(namaKelas || "").toUpperCase();
  var isPpki = kUpper.includes("VIVA") || kUpper.includes("WIRA") || kUpper.includes("ARENA") || 
               kUpper.includes("AXIA") || kUpper.includes("SAGA") || kUpper.includes("BEZZA") || kUpper.includes("PPKI");
  var isPra = kUpper.includes("PRA");
  
  var dataSheet = muatSemuaDskpDariSheet_(isPpki);
  var senaraiAlias = seragamkanNamaSubjek_(subjek);
  
  var hasil = [];
  if (dataSheet && dataSheet.length > 0) {
    for (var i = 0; i < dataSheet.length; i++) {
      var row = dataSheet[i];
      var subUpper = row.subjek.toUpperCase();
      var subPadan = senaraiAlias.some(function(al) {
        return subUpper === al || subUpper.includes(al) || al.includes(subUpper);
      });
      if (subPadan && padanTahunSama_(row.tahun, tahun, isPpki)) {
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
    hasil = dapatkanDskpPiawaiKpm_(subjek, tahun, isPpki, isPra);
  }
  
  return hasil;
}

// Dapatkan rekod DSKP spesifik yang menjamin TEMA, TAJUK, SK, dan SP sentiasa TALLY 100%
function ambilObjektifDskp_(subjek, tingkatTahun, namaKelas, temaPilihan, tajukPilihan, skPilihan, spPilihan) {
  var senarai = getHierarkiDskp_(subjek, tingkatTahun, namaKelas);
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
function getDskpHierarkiWeb_(subjek, tahun, namaKelas) {
  return getHierarkiDskp_(subjek, tahun, namaKelas);
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
function lampirkanIthink_(res, petaIthinkNama) {
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
function binaKandunganRphSpesifik_(subjek, tahun, kelas, slotCustom) {
  var petaIthinkNama = slotCustom && slotCustom.petaIthink ? String(slotCustom.petaIthink).trim() : "";
  var hasil = binaKandunganRphSpesifikTeras_(subjek, tahun, kelas, slotCustom);
  return lampirkanIthink_(hasil, petaIthinkNama);
}

function binaKandunganRphSpesifikTeras_(subjek, tahun, kelas, slotCustom) {
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
  var dskp = ambilObjektifDskp_(subjek, tahun, kelas, sTema, sTajuk, sSk, sSp);
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
// Matriks Peranan & Kawalan Hak Akses (RBAC) __NAMA_SEKOLAH__ 2026
function tentukanSkopPeranan_(peranan) { return roleMap_(String(peranan||'').toUpperCase().trim()); }

function getSenaraiGuruWeb_() { return users_(konteks_().school).filter(function(u){return u.aktif;}); }

/**
 * Tambah Pengguna Baru (Akses Pentadbir sahaja)
 * Auto generate password berasaskan 4 nombor belakang IC
 */
function tambahPenggunaBaru_(data) { return manageUser_(data,false); }

/**
 * Kemaskini Maklumat Guru (Akses Pentadbir sahaja)
 */
function kemaskiniMaklumatGuru_(data) { return manageUser_(data,true); }

/**
 * Sahkan Passcode / Kata Laluan Guru (Backend Security)
 * Berpandukan kata laluan tersimpan atau 4 digit terakhir No. KP
 */
function sahkanPasscodeGuru_() { throw new Error('Gunakan kod emel sekali guna.'); }

function simpanFotoProfilGuru_(emel, base64Data) {
  try {
    if (!emel) return { success: false, message: "Emel tidak sah" };
    var ss = sekolahSpreadsheet_();
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
      if (padanEmelSama_(data[i][colEmel - 1], emel)) {
        sheetPenga.getRange(i + 1, colFoto).setValue(base64Data);
        break;
      }
    }
    return { success: true, message: "Foto profil berjaya disimpan!" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function dapatkanProfilPenggunaSemasa_() { return {emelSesi:konteks_().user.emel,guru:konteks_().user}; }

function dapatkanNamaGuruDariEmel_(emel) {
  var senarai = getSenaraiGuruWeb_();
  var guru = senarai.find(function(g) { return padanEmelSama_(g.emel, emel); });
  return guru ? guru.nama : (emel || "Badrul Hisyam Rasamin");
}

function getPilihanMingguWeb_() {
  var ss = sekolahSpreadsheet_();
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
function semakAdaRekod_(minggu, emel) {
  var ss = sekolahSpreadsheet_();
  var sheet = ss.getSheetByName("RPH_GURU");
  if (!sheet) return false;
  // Gunakan getDisplayValues() untuk memelihara teks mentah tepat dari sel
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (padanEmelSama_(data[i][1], emel) && padanMingguSama_(data[i][2], minggu)) {
      return true;
    }
  }
  return false;
}

function dapatkanRekodMinggu_(minggu, emel) {
  var ss = sekolahSpreadsheet_();
  var sheet = ss.getSheetByName("RPH_GURU");
  if (!sheet) return [];
  
  // Gunakan getDisplayValues() untuk memastikan nilai masa kekal sebagai string mentah dari sel (mengelak penukaran ke objek Date 1899)
  var data = sheet.getDataRange().getDisplayValues();
  var senarai = [];
  var isninTakwim = dapatkanIsninMinggu_(minggu);

  for (var i = 1; i < data.length; i++) {
    var rowEmel = data[i][1];
    var rowMinggu = data[i][2];
    
    if (padanEmelSama_(rowEmel, emel) && padanMingguSama_(rowMinggu, minggu)) {
      var hariPenuh = formatNamaHari_(data[i][3]);
      var masaMula = bersihkanMasa_(data[i][4]);
      var masaTamat = bersihkanMasa_(data[i][5]);
      var namaKelas = data[i][6] ? data[i][6].toString().trim() : "";
      var subjek = data[i][7] ? data[i][7].toString().trim() : "";
      var tarikhSlot = dapatkanTarikhHari_(isninTakwim, hariPenuh);
      
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
        var rphBaru = binaKandunganRphSpesifik_(subjek, "1", namaKelas);
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
        kelasPaparan: formatKelasTahun_(namaKelas, ""),
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

function janaRphSemingguBackend_(payload) {
  var ss = sekolahSpreadsheet_();
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
      
      var rphKandungan = binaKandunganRphSpesifik_(s.subjek, s.tahun, s.kelas, s);
      var hariBersih = formatNamaHari_(hari);
      var mulaBersih = bersihkanMasa_(s.mula);
      var tamatBersih = bersihkanMasa_(s.tamat);

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

function kemaskiniRefleksi_(id, teksRefleksi) {
  var ss = sekolahSpreadsheet_();
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
function escapeHtmlGas_(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function janaPdfMingguanBackend_(minggu, emel) {
  try {
    var namaGuru = dapatkanNamaGuruDariEmel_(emel);
    if (!namaGuru) namaGuru = emel || "GURU BERTUGAS";
    var rekod = dapatkanRekodMinggu_(minggu, emel) || [];
    var kodMinggu = formatKodMinggu_(minggu);
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
            <h2>REKOD PENGAJARAN DAN PEMBELAJARAN HARIAN (__NAMA_SEKOLAH__)</h2>
            <div class="meta">GURU: ${escapeHtmlGas_(namaGuru.toUpperCase())} &nbsp;|&nbsp; SESI: 2026 &nbsp;|&nbsp; MINGGU: ${escapeHtmlGas_(kodMinggu)}</div>
          </div>
    `;

    if (rekod.length === 0) {
      htmlContent += `<div style="text-align:center; padding: 40px 10px; color:#be123c; font-weight:bold;">Tiada rekod e-RPH dijumpai untuk ${escapeHtmlGas_(minggu)}. Sila jana rekod PdP terlebih dahulu dalam portal sekolah.</div>`;
    } else {
      rekod.forEach(function(r, idx) {
        var bil = idx + 1;
        var formatAktivitiHtml = String(r.aktiviti || '').split('\n').map(function(line) {
          if (!line.trim()) return '';
          var parts = line.split(':');
          if (parts.length > 1) {
            return '<div class="act-step"><b>' + escapeHtmlGas_(parts[0].trim()) + ' :</b> ' + escapeHtmlGas_(parts.slice(1).join(':').trim()) + '</div>';
          }
          return '<div class="act-step">' + escapeHtmlGas_(line.trim()) + '</div>';
        }).join('');

        var formatKriteriaHtml = String(r.kriteriaKejayaan || '').split('\n').map(function(line) {
          if (!line.trim()) return '';
          return '<div style="margin-bottom:2px;">' + escapeHtmlGas_(line.trim()) + '</div>';
        }).join('');

        var subUpper = String(r.subjek || "").toUpperCase();
        var isEnglish = subUpper.includes("INGGERIS") || subUpper.includes("ENGLISH") || subUpper.includes("[BI]") || subUpper === "BI";
        var kelasTeks = r.kelasPaparan ? r.kelasPaparan : (r.kelas || "-");
        var refleksiTeks = r.refleksi ? escapeHtmlGas_(r.refleksi) : '<em>PdP terlaksana dengan jayanya mengikut perancangan.</em>';

        htmlContent += `
          <!-- 2. KOTAK JADUAL SLOT PdP -->
          <div class="slot-box">
            <div class="slot-title">
              ${bil}. ${escapeHtmlGas_(r.hari)} (${escapeHtmlGas_(r.tarikh || '-')}) &nbsp;|&nbsp; MASA: ${escapeHtmlGas_(r.mula)} - ${escapeHtmlGas_(r.tamat)}
            </div>
            <table class="slot-content">
              <tr>
                <td class="field-label">${isEnglish ? "CLASS &amp; SUBJECT" : "KELAS &amp; SUBJEK"}</td>
                <td class="field-val"><b>${escapeHtmlGas_(kelasTeks)} &mdash; ${escapeHtmlGas_(r.subjek)}</b></td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? "THEME / TOPIC" : "TEMA / TAJUK"}</td>
                <td class="field-val">${escapeHtmlGas_(r.temaTajuk || '-')}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? "CONTENT STANDARD" : "STANDARD KANDUNGAN"}</td>
                <td class="field-val">${escapeHtmlGas_(r.sk || '-')}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? "LEARNING STANDARD" : "STANDARD PEMBELAJARAN"}</td>
                <td class="field-val">${escapeHtmlGas_(r.sp || '-')}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? "LEARNING OBJECTIVES" : "OBJEKTIF PEMBELAJARAN"}</td>
                <td class="field-val">${escapeHtmlGas_(r.objektif || '-')}</td>
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
                <td class="field-val">${escapeHtmlGas_(r.bbmNilaiKbat || '-')}</td>
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
                Tandatangan Guru: <b>${escapeHtmlGas_(namaGuru)}</b><br>
                Tarikh: ${Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy")}
              </div>
            </td>
            <td style="width: 50%; text-align: right;">
              <br><br>
              <div class="sign-line" style="margin-left: auto;">
                Disemak &amp; Disahkan oleh Pentadbir __NAMA_SEKOLAH__<br>
                Status: <b>RUJUK REKOD SEMAKAN DALAM PORTAL</b>
              </div>
            </td>
          </tr>
        </table>
      `;
    }

    htmlContent += `</body></html>`;

    // Penjanaan Blob PDF terus dalam memori (lebih pantas, elak ralat cipta fail HTML sementara di Drive root)
    var htmlBlob = Utilities.newBlob(brandHtml_(htmlContent), 'text/html', 'rph.html');
    var pdfBlob = htmlBlob.getAs('application/pdf').setName(namaFailPdf);
    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());

    // Simpan salinan ke Google Drive (jika ada kebenaran)
    var urlPdf = "";
    var downloadUrl = "";
    try {
      var folder = dapatkanAtauCiptaFolderArkib_("2026", namaGuru, minggu);
      // Folder Google Drive berhierarki sedia digunakan
      var pdfFile = folder.createFile(pdfBlob);
      try {
        
      } catch (errDomain) {
        try {
          
        } catch (eSub) {
          console.warn("Domain restriction: " + eSub.message);
        }
      }
      urlPdf = pdfFile.getUrl();
      downloadUrl = pdfFile.getUrl().replace('view?usp=drivesdk', 'export?format=pdf');
    } catch (dErr) {
      throw new Error("PDF dijana tetapi arkib Drive gagal: " + dErr.message);
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
    console.error("Ralat janaPdfMingguanBackend_: " + err.message);
    throw new Error("Gagal menjana PDF e-RPH: " + err.message);
  }
}

// --------------------------------------------------------------------------
// 6. DASHBOARD PENTADBIR & SEMAKAN
// --------------------------------------------------------------------------
function dapatkanDataDashboardPentadbir_(minggu) {
  var ss = sekolahSpreadsheet_();
  var sheetSemakan = ss.getSheetByName("SEMAKAN_PENTADBIR");
  var senaraiGuru = getSenaraiGuruWeb_();
  
  var mapSemakan = {};
  if (sheetSemakan) {
    var dataSemakan = sheetSemakan.getDataRange().getValues();
    for (var s = 1; s < dataSemakan.length; s++) {
      if (padanMingguSama_(dataSemakan[s][1], minggu)) {
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
    var ada = semakAdaRekod_(minggu, em);
    
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

function simpanSemakanPentadbir_(minggu, emel, status, ulasan) {
  try {
    var ss = sekolahSpreadsheet_();
    var sheet = ss.getSheetByName("SEMAKAN_PENTADBIR");
    
    if (!sheet) {
      sheet = ss.insertSheet("SEMAKAN_PENTADBIR");
      sheet.appendRow(["ID Semakan", "Minggu", "Emel Guru", "Status Semakan", "Ulasan Pentadbir", "Tarikh Kemaskini"]);
    }
    
    var data = sheet.getDataRange().getValues();
    var jumpaiBaris = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (padanMingguSama_(data[i][1], minggu) && padanEmelSama_(data[i][2], emel)) {
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
    console.error("Ralat simpanSemakanPentadbir_: " + err.message);
    throw new Error("Gagal menyimpan semakan pentadbir: " + err.message);
  }
}

// --------------------------------------------------------------------------
// 7. LAPORAN BERTUGAS (HEM)
// --------------------------------------------------------------------------
function simpanLaporanBertugasBackend_(emel, minggu, hari, kelas, hadirL, hadirP, tHadirL, tHadirP, catatan) {
  var emelGuru = String(emel || "").trim();
  if (!emelGuru) throw new Error("Emel guru tidak dikesan oleh sistem.");

  var ss = sekolahSpreadsheet_();
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


// Pengecaman Tab Kategori Kokurikulum PPKI (PPKI PENGAKAP, PPKI BADMINTON, PPKI SENI BUDAYA)


// Pengambilan Senarai Murid (Read)


// Penyimpanan Kehadiran Murid (Write / 1=Hijau, 0=Merah)


// Janaan Automatik Kandungan OPR Kokurikulum
function janaKandunganOPRKokumBackend_(tajuk, unit) {
  return janaKandunganOPRKokum_(tajuk, unit);
}

// Simpan Laporan OPR Kokurikulum ke Tab LAPORAN_KOKUM
function simpanLaporanOprKokumBackend_(formData) {
  return simpanPelaporanKokumBackend_(formData);
}

// Penjanaan PDF OPR Kokurikulum (A4 Portrait - 1 Muka Surat)
function janaPdfOprKokumBackendPortal_(formData) {
  return janaPdfOprKokumBackend_(formData);
}

// --------------------------------------------------------------------------
// 9. SISTEM PENGURUSAN OPR SEKOLAH, WORKFLOW PENGESAHAN & BANK OPR
// --------------------------------------------------------------------------

// Ekstrak ID fail Google Drive daripada sebarang format URL Drive atau rentetan ID
function ambilFileIdDariUrl_(url) {
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
function ambilBase64DariDrive_(urlOrId) {
  if (!urlOrId || typeof urlOrId !== 'string') return "";
  var str = urlOrId.trim();
  if (str.startsWith("data:image")) return str; // sudah berformat data URI base64
  
  var fileId = ambilFileIdDariUrl_(str);
  if (!fileId) return "";

  try {
    var file = scopedFile_(fileId);
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
function dapatkanGambarLaporanOprBase64_(idLaporan) {
  try {
    var sheet = dapatkanAtauCiptaSheetOpr_();
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
      var b64 = raw ? ambilBase64DariDrive_(raw) : "";
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

// Dapatkan Data URI Base64 Lencana __NAMA_SEKOLAH__ (dari Folder Google Drive atau Tetapan Skrip)
function dapatkanLogoSekolah_() { var id=konteks_().school.Logo_File_ID;return id?ambilBase64DariDrive_(id):''; }

// Simpan Fail / Gambar Base64 ke Folder Google Drive PORTAL_SK_SOOK_OPR_MEDIA
function simpanFailKeDrive_(base64Data,namaFail) { return simpanMedia_(base64Data,namaFail); }

// Dapatkan atau cipta helaian tab LAPORAN_OPR dengan 26 lajur kawalan
function dapatkanAtauCiptaSheetOpr_() {
  var ss = sekolahSpreadsheet_();
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
function janaKandunganOprPintar_(bahagian, namaProgram, unit) {
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
    obj = "1. Meningkatkan kesedaran, pembudayaan sahsiah terpuji dan disiplin kendiri dalam kalangan murid __NAMA_SEKOLAH__ berkaitan " + p + ".\n" +
          "2. Memastikan keselamatan, kebajikan dan kesejahteraan murid sentiasa terpelihara di peringkat sekolah.\n" +
          "3. Memupuk hubungan silaturahim dan persefahaman positif antara pihak sekolah, ibu bapa dan murid.";
    pengisian = "1. Taklimat pengurusan sahsiah dan penerangan objektif pelaksanaan program " + p + " oleh penyelaras HEM.\n" +
                "2. Ceramah interaktif, tayangan video kesedaran dan sesi soal jawab bersama murid.\n" +
                "3. Aktiviti bengkel penghayatan nilai murni dan amalan terbaik dalam persekitaran sekolah.\n" +
                "4. Rumusan dan pelancaran ikrar sahsiah terpuji murid __NAMA_SEKOLAH__ bagi sesi 2026.";
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

// Wrapper serasi belakang untuk janaKandunganOprUmum_
function janaKandunganOprUmum_(namaProgram, ringkasan) {
  return janaKandunganOprPintar_("KURIKULUM", namaProgram, "Panitia Sekolah");
}

// Simpan atau Kemas Kini Rekod OPR ke Google Sheets (Tab LAPORAN_OPR)
function simpanAtauKemasKiniOpr_(formData) {
  try {
    var sheet = dapatkanAtauCiptaSheetOpr_();
    var data = sheet.getDataRange().getValues();
    
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

    if (!idLaporan) {
      idLaporan = "OPR-" + Utilities.getUuid();
      formData.idLaporan=idLaporan;
    }

    ROUTE_=metaOpr_(formData);
    idLaporan=formData.idLaporan;
    var gUrls = ["", "", "", ""];
    for (var k = 1; k <= 4; k++) {
      var gData = formData["gambar" + k] || formData["gambar" + k + "Url"] || "";
      if (gData && String(gData).startsWith("data:")) {
        gUrls[k - 1] = simpanFailKeDrive_(gData, "OPR_" + (formData.namaProgram || "PROG") + "_G" + k);
      } else if (gData) {
        scopedFile_(ambilFileIdDariUrl_(gData));
        gUrls[k - 1] = gData;
      } else if (rowIndex > 0 && data[rowIndex - 1] && data[rowIndex - 1][14 + k]) {
        gUrls[k - 1] = data[rowIndex - 1][14 + k]; // Kekalkan pautan gambar asal jika tiada muat naik baharu
      } else {
        gUrls[k - 1] = "";
      }
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
      resPdf = janaPdfOprBackend_(recordPayload);
    } catch (ePdf) {
      throw new Error("Arkib PDF OPR gagal: " + ePdf.message);
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

// Wrapper serasi belakang untuk simpanOprUmumBackend_
function simpanOprUmumBackend_(formData) {
  return simpanAtauKemasKiniOpr_(formData);
}

// Dapatkan Senarai Laporan OPR untuk Semakan Pentadbir mengikut Skop RBAC
function dapatkanSenaraiOprPentadbir_(emelPentadbir) {
  var senarai = [];
  try {
    var senaraiGuru = getSenaraiGuruWeb_();
    var admin = senaraiGuru.find(function(g) { return padanEmelSama_(g.emel, emelPentadbir); });
    if (!admin) return [];
    var rbac = tentukanSkopPeranan_(admin.peranan);
    if (!rbac.bolehLulus) return [];

    var sheet = dapatkanAtauCiptaSheetOpr_();
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
function sahkanLaporanOprBackend_(idLaporan, emelPentadbir, statusTindakan, catatanPk) {
  try {
    var sheet = dapatkanAtauCiptaSheetOpr_();
    var data = sheet.getDataRange().getValues();
    var senaraiGuru = getSenaraiGuruWeb_();
    var admin = senaraiGuru.find(function(g) { return padanEmelSama_(g.emel, emelPentadbir); });
    
    if (!admin) {
      throw new Error("Pengesahan gagal: Emel pentadbir tidak sah dalam sistem.");
    }
    var rbac = tentukanSkopPeranan_(admin.peranan);
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
        jawatanPengesah: rbac.label + ", __NAMA_SEKOLAH__",
        tarikhSah: tarikhKini,
        catatanPk: ""
      };

      var resPdf = janaPdfOprBackend_(recordPayload);
      sheet.getRange(targetRow, 20).setValue("Disahkan");
      sheet.getRange(targetRow, 21).setValue(admin.nama);
      sheet.getRange(targetRow, 22).setValue(rbac.label + ", __NAMA_SEKOLAH__");
      sheet.getRange(targetRow, 23).setValue(tarikhKini);
      sheet.getRange(targetRow, 24).setValue("");

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
function dapatkanBankOprSekolah_(penapis) {
  var hasil = [];
  penapis = penapis || {};
  try {
    var sheet = dapatkanAtauCiptaSheetOpr_();
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
        if (!padanEmelSama_(emelPenyelaras, penapis.emelGuru)) {
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

// Wrapper serasi belakang untuk getSenaraiArkibOpr_
function getSenaraiArkibOpr_() {
  return dapatkanBankOprSekolah_({ hanyaDisahkan: true });
}

// Penjanaan PDF OPR Standard (A4 Portrait - 1 Muka Surat Sahaja)
function janaPdfOprBackend_(formData) {
  try {
    var prog = formData.namaProgram || "LAPORAN PROGRAM SEKOLAH";
    var bahagian = formData.bahagian || "KURIKULUM";
    var unit = formData.unit || formData.unitPanitia || "Unit Sekolah";
    var penyelaras = formData.penyelaras || "Penyelaras Program";
    var tempat = formData.tempat || "__NAMA_SEKOLAH__";
    var tarikh = formData.tarikh || "-";
    var masa = formData.masa || "-";
    var kehadiran = formData.kehadiran || formData.hadirAhli || "Warga & Murid __NAMA_SEKOLAH__";
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
        var b64Drive = ambilBase64DariDrive_(gUrls[gi]);
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

    var logoSekolahSrc = dapatkanLogoSekolah_();
    var logoHeaderHtml = logoSekolahSrc ? 
      '<img src="' + logoSekolahSrc + '" style="max-height:42px; max-width:42px; object-fit:contain;" alt="Lencana __NAMA_SEKOLAH__" />' : 
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
        <div class="title-main">__NAMA_SEKOLAH__</div>
        <div class="title-sub">LAPORAN SATU MUKA SURAT (ONE PAGE REPORT - OPR)</div>
        <div style="font-size: 7.2px; color: #64748b;">__ALAMAT_SEKOLAH__ • Kod Sekolah: __KOD_SEKOLAH__ • Bahagian: ${bahagian}</div>
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
                <div style="font-size: 7.2px; font-weight: 700; color: #047857;">${formData.jawatanPengesah || "PENTADBIR, __NAMA_SEKOLAH__"}</div>
                <div style="font-size: 6.8px; color: #4b5563; margin-top: 1px;">Disahkan pada: ${formData.tarikhSah || "-"}</div>
              </div>
              <div style="text-align:center; padding: 1px 4px; background:#dcfce7; border:1px solid #86efac; border-radius:3px; font-size:6.8px; font-weight:bold; color:#166534;">
                🛡️ DISAHKAN DIGITAL
              </div>
            </div>
            <div style="font-size: 6.5px; color: #065f46; margin-top: 2px; border-top: 1px dashed #a7f3d0; padding-top: 2px; line-height:1.2;">
              ✓ Dokumen ini telah disahkan secara digital melalui Portal Rasmi __NAMA_SEKOLAH__. Rekod pengesahan dalaman sekolah.
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

    ROUTE_=metaOpr_(formData);

    var namaFailPdf = "OPR_" + (isDisahkan ? "RASMI_" : "DRAF_") + prog.replace(/\s+/g, '_') + ".pdf";
    var pdfBlob = Utilities.newBlob(brandHtml_(htmlContent), 'text/html', 'document.html').getAs(MimeType.PDF).setName(namaFailPdf);
    var pdfFile = saveBlob_(pdfBlob,ROUTE_);

    try {
      
    } catch (errDomain) {
      try {  } catch (eSub) {}
    }

    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());


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
// MODUL e-KEHADIRAN GURU BERASASKAN GPS GEOFENCING (__NAMA_SEKOLAH__)
// KOORDINAT RASMI: 5°08'52.7"N 116°18'25.9"E | RADIUS: 100 METER
// ==========================================================================



/**
 * Mengira jarak antara dua koordinat menggunakan formula Haversine (unit: Meter)
 */
function kiraJarakHaversine_(lat1, lon1, lat2, lon2) {
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
function rakamKehadiranGpsBackend_(data) {
  try {
    var ss = sekolahSpreadsheet_();
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
    var jarakSebenar = kiraJarakHaversine_(lat, lon, gpsSekolah_().lat, gpsSekolah_().lon);
    var dalamKawasan = (jarakSebenar <= gpsSekolah_().radiusMeter);
    var jenis = String(data.jenis || "MASUK").toUpperCase();

    // SPESIFIKASI e-KEHADIRAN GPS PINTAR:
    // Waktu Masuk: Wajib dalam radius 150m; sekat jika di luar kawasan
    if (jenis === "MASUK" && !dalamKawasan) {
      return {
        status: "ERROR",
        message: "Sekatan Kehadiran: Anda berada " + jarakSebenar + "m di luar kawasan __NAMA_SEKOLAH__ (Had Geofence: " + gpsSekolah_().radiusMeter + "m). Waktu Masuk hanya dibenarkan semasa berada di dalam kawasan sekolah."
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
      var rowTarikh = formatTarikhStandard_(dataKehadiran[r][2]);
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
function autoCheckoutHarian5PM_() {
  try {
    var now = new Date();
    var dayOfWeek = now.getDay(); // 0 = Ahad, 1 = Isnin, ..., 5 = Jumaat, 6 = Sabtu
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      return { success: false, message: "Bukan hari persekolahan (Isnin - Jumaat)." };
    }

    var ss = sekolahSpreadsheet_();
    var sheet = ss.getSheetByName("KEHADIRAN_GURU");
    if (!sheet) return { success: false, message: "Sheet KEHADIRAN_GURU tidak ditemui." };

    var tarikhStr = Utilities.formatDate(now, "Asia/Kuala_Lumpur", "dd/MM/yyyy");
    var hariList = ["Ahad", "Isnin", "Selasa", "Rabu", "Khamis", "Jumaat", "Sabtu"];
    var hariStr = hariList[dayOfWeek];
    var data = sheet.getDataRange().getValues();

    var guruMasuk = {};
    var guruPulang = {};

    for (var i = 1; i < data.length; i++) {
      var rowTarikh = formatTarikhStandard_(data[i][2]);
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
    console.warn("Ralat autoCheckoutHarian5PM_: " + err.message);
    return { success: false, message: err.message };
  }
}

/**
 * Pemasang Trigger Time-Driven Auto Checkout jam 5:00 Petang setiap hari persekolahan
 */
function pasangTriggerAutoCheckout_() { throw new Error('Gunakan pasangCheckoutSaas_ dalam editor.'); }

/**
 * Padam RPH Mingguan yang Belum Disemak/Diluluskan oleh Pentadbir
 */
function padamRphMingguanBackend_(emel, mingguAtauId) {
  try {
    if (!emel || !mingguAtauId) {
      return { success: false, message: "Parameter emel dan minggu/ID diperlukan." };
    }
    var ss = sekolahSpreadsheet_();
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

        if (aEmel === targetEmel && (aId === targetMinggu || aMinggu.toLowerCase() === targetMinggu.toLowerCase() || padanMingguSama_(aMinggu, targetMinggu))) {
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
        if (rEmel === targetEmel && (padanMingguSama_(rMinggu, targetMinggu) || rMinggu.toLowerCase() === targetMinggu.toLowerCase())) {
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
function simpanJadualGuruBackend_(emel, jadualObj) {
  try {
    if (!emel) return { success: false, message: "Emel guru diperlukan." };
    var ss = sekolahSpreadsheet_();
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

    return { success: true, message: "Templat jadual waktu berjaya disimpan secara kekal!" };
  } catch (err) {
    return { success: false, message: "Ralat menyimpan templat jadual: " + err.message };
  }
}

/**
 * Ambil Templat Jadual Waktu Kekal Guru dari Lembaran JADUAL_GURU
 */
function dapatkanJadualGuruBackend_(emel) {
  try {
    if (!emel) return { success: false, message: "Emel tidak sah." };
    var ss = sekolahSpreadsheet_();
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

function formatTarikhStandard_(val) {
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
function formatMasaStandard_(val) {
  if (!val && val !== 0) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Kuala_Lumpur", "hh:mm:ss a");
  }
  return String(val).trim();
}

/**
 * Dapatkan status kehadiran hari ini bagi guru tertentu (Masuk & Pulang) serta 7 rekod terkini
 */
function dapatkanStatusKehadiranHariIni_(emel, nama) {
  try {
    var ss = sekolahSpreadsheet_();
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
      var rTarikh = formatTarikhStandard_(r[2]);
      var rHari = String(r[3] || "").trim();
      var rMasa = formatMasaStandard_(r[4]);
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
function dapatkanSemuaKehadiranHariIni_() {
  try {
    var ss = sekolahSpreadsheet_();
    var sheet = ss.getSheetByName("KEHADIRAN_GURU");
    var now = new Date();
    var tarikhHariIni = Utilities.formatDate(now, "Asia/Kuala_Lumpur", "dd/MM/yyyy");

    if (!sheet) return [];

    var rows = sheet.getDataRange().getValues();
    var senarai = [];
    for (var i = rows.length - 1; i >= 1; i--) {
      var r = rows[i];
      var rTarikh = formatTarikhStandard_(r[2]);
      if (rTarikh === tarikhHariIni) {
        senarai.push({
          idRekod: String(r[0] || ""),
          masa: formatMasaStandard_(r[4]),
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
function dapatkanUrlWebApp_() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return "";
  }
}

// ==================== STATUS FEED FACEBOOK-STYLE & ONLINE LIVE ====================

function initSheetStatusFeed_(ss) {
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
function formatTarikhMasa24Jam_(idStr, valTarikh, valMasa) {
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

function getFeedStatusWeb_(emelSemasa) {
  try {
    var ss = sekolahSpreadsheet_();
    var sheet = ss.getSheetByName("STATUS_FEED");
    if (!sheet) {
      sheet = initSheetStatusFeed_(ss);
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
      var formatMasa = formatTarikhMasa24Jam_(idStatus, r[1], r[2]);

      var likesArr = [];
      try { likesArr = JSON.parse(likesJsonStr); } catch (e) { likesArr = []; }

      var isLiked = (emelSemasa && likesArr.indexOf(emelSemasa) !== -1);

      // Kira bilangan siaran aktif pengguna semasa (maksimum 5 siaran dalam 24 jam)
      if (padanEmelSama_(emel, emelSemasa)) {
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

function hantarStatusFeed_(emel, kandungan, imejBase64) {
  try {
    if (!emel || !kandungan) {
      return { success: false, message: "Kandungan status tidak boleh kosong." };
    }
    var ss = sekolahSpreadsheet_();
    var sheet = ss.getSheetByName("STATUS_FEED");
    if (!sheet) {
      sheet = initSheetStatusFeed_(ss);
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

      if (dalam24Jam && padanEmelSama_(data[i][3], emel)) {
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
    var senaraiGuru = getSenaraiGuruWeb_();
    var guru = senaraiGuru.find(function(g) { return padanEmelSama_(g.emel, emel); }) || {
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
      message: "✅ Status anda berjaya disiarkan ke Suara __NAMA_SEKOLAH__! (Paparan aktif selama 24 jam)",
      bakiKuota: Math.max(0, 4 - count24Jam)
    };
  } catch (err) {
    return { success: false, message: "Ralat menyiarkan status: " + err.toString() };
  }
}

function toggleLikeStatusFeed_(idStatus, emel) {
  try {
    if (!idStatus || !emel) return { success: false };
    var ss = sekolahSpreadsheet_();
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

function padamStatusFeed_(idStatus, emel) {
  try {
    var ss = sekolahSpreadsheet_();
    var sheet = ss.getSheetByName("STATUS_FEED");
    if (!sheet) return { success: false };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idStatus)) {
        var pemilik = String(data[i][3]).trim();
        if (padanEmelSama_(pemilik, emel) || emel.includes("admin") || emel.includes("gb@")) {
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

function dapatkanBilanganOnlineLive_(emel) {
  var s=konteks_().school,cache=CacheService.getScriptCache(),prefix='online:'+s.Sekolah_ID+':';
  cache.put(prefix+hash_(emel),'1',900);
  var count=users_(s).filter(function(u){return u.aktif&&cache.get(prefix+hash_(u.emel));}).length;
  return {success:true,count:count};
}

// ==========================================
// 12. MODUL PENILAIAN KOMPONEN KEBERHASILAN (PBPPP) KPM
// GARIS PANDUAN PELAKSANAAN PBPPP KPM (GPPKK)
// ==========================================

function dapatkanAtauCiptaSheetKeberhasilan_() {
  var ss = sekolahSpreadsheet_();
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

function simpanRekodKeberhasilan_(payload) {
  try {
    if (!payload || !payload.emel) {
      return { success: false, message: "Maklumat guru tidak lengkap." };
    }

    var sheet = dapatkanAtauCiptaSheetKeberhasilan_();
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
      if (String(data[r][1]).trim() === tahun && padanEmelSama_(data[r][2], emel)) {
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

function muatRekodKeberhasilanGuru_(emel, tahun) {
  try {
    var ss = sekolahSpreadsheet_();
    var sheet = dapatkanAtauCiptaSheetKeberhasilan_();
    var targetTahun = String(tahun || "2026").trim();
    var data = sheet.getDataRange().getValues();

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][1]).trim() === targetTahun && padanEmelSama_(data[r][2], emel)) {
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

function dapatkanSemuaRekodKeberhasilanAdmin_(tahun) {
  try {
    var ss = sekolahSpreadsheet_();
    var sheet = dapatkanAtauCiptaSheetKeberhasilan_();
    var targetTahun = String(tahun || "2026").trim();
    var senaraiGuru = getSenaraiGuruWeb_();
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
// PENGURUSAN ARKIB e-RPH GOOGLE DRIVE & LAPORAN BULANAN (SISTEM __NAMA_SEKOLAH__)
// ==========================================================================

/**
 * Cipta / dapatkan hierarki folder Google Drive:
 * e-RPH __NAMA_SEKOLAH__ / [Tahun] / [Nama Guru] / [Bulan atau Minggu]
 */
function dapatkanAtauCiptaFolderArkib_(tahun,namaGuru,mingguAtauBulan) {
  var meta=ROUTE_||{jenisDokumen:'RPH',format:'PDF',tahun:String(tahun),guruId:hash_(konteks_().user.emel).slice(0,12)};
  return {createFile:function(blob){return saveBlob_(blob,meta);},getUrl:function(){return 'https://drive.google.com/drive/folders/'+routeFolder_(meta);}};
}

/**
 * Cipta / dapatkan helaian ARKIB_ERPH_GURU untuk rekod penghantaran profil guru
 */
function dapatkanAtauCiptaSheetArkib_() {
  var ss = sekolahSpreadsheet_();
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
function simpanRekodArkibErph_(payload) {
  try {
    var sheet = dapatkanAtauCiptaSheetArkib_();
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
    console.warn("Ralat simpanRekodArkibErph_: " + err.message);
    return { success: false, message: err.message };
  }
}

/**
 * Dapatkan senarai rekod arkib e-RPH bagi guru tertentu untuk dipaparkan di Tab Profil
 */
function dapatkanSenaraiArkibRphGuru_(emel) {
  try {
    var ss = sekolahSpreadsheet_();
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
    console.warn("Ralat dapatkanSenaraiArkibRphGuru_: " + err.message);
    return [];
  }
}

/**
 * Semak status penghantaran minggu terdahulu (Minggu N-1)
 */
function semakStatusMingguTertunggak_(mingguSemasaStr, emel) {
  try {
    var m = String(mingguSemasaStr || "").match(/\d+/);
    if (!m) return { adaTertunggak: false };
    var n = parseInt(m[0], 10);
    if (n <= 1) return { adaTertunggak: false };

    var mingguLalu = "Minggu " + (n - 1);
    var ss = sekolahSpreadsheet_();
    var sheet = ss.getSheetByName("RPH_GURU");
    if (!sheet) return { adaTertunggak: false };

    var data = sheet.getDataRange().getValues();
    var jumpa = false;
    for (var i = 1; i < data.length; i++) {
      if (padanEmelSama_(data[i][1], emel) && padanMingguSama_(data[i][2], mingguLalu)) {
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
 * Janaan Kompilasi e-RPH Sebulan Penuh (Bulanan) ke dalam 1 Dokumen PDF
 */
function janaRphBulanan_(emel, bulan, tahun) {
  try {
    var namaGuru = dapatkanNamaGuruDariEmel_(emel);
    if (!namaGuru) namaGuru = emel || "GURU";
    var sTahun = String(tahun || "2026").trim();
    var sBulan = String(bulan || "Januari").trim();

    // Petakan bulan kepada minggu anggaran takwim
    var mapBulanKeMinggu = {
      "Januari": ["Minggu 1", "Minggu 2", "Minggu 3", "Minggu 4"],
      "Februari": ["Minggu 5", "Minggu 6", "Minggu 7", "Minggu 8"],
      "Mac": ["Minggu 9", "Minggu 10", "Minggu 11", "Minggu 12"],
      "April": ["Minggu 13", "Minggu 14", "Minggu 15", "Minggu 16"],
      "Mei": ["Minggu 17", "Minggu 18", "Minggu 19", "Minggu 20"],
      "Jun": ["Minggu 21", "Minggu 22", "Minggu 23", "Minggu 24"],
      "Julai": ["Minggu 25", "Minggu 26", "Minggu 27", "Minggu 28"],
      "Ogos": ["Minggu 29", "Minggu 30", "Minggu 31", "Minggu 32"],
      "September": ["Minggu 33", "Minggu 34", "Minggu 35", "Minggu 36"],
      "Oktober": ["Minggu 37", "Minggu 38", "Minggu 39", "Minggu 40"],
      "November": ["Minggu 41", "Minggu 42", "Minggu 43"],
      "Disember": ["Minggu 44", "Minggu 45"]
    };

    var targetMingguList = mapBulanKeMinggu[sBulan] || ["Minggu 1", "Minggu 2", "Minggu 3", "Minggu 4"];
    var semuaRekodBulan = [];

    for (var mIdx = 0; mIdx < targetMingguList.length; mIdx++) {
      var w = targetMingguList[mIdx];
      var wRekod = dapatkanRekodMinggu_(w, emel) || [];
      if (wRekod.length > 0) {
        semuaRekodBulan.push({ minggu: w, rekod: wRekod });
      }
    }

    if (semuaRekodBulan.length === 0) {
      throw new Error("Tiada rekod e-RPH ditemui bagi bulan " + sBulan + " (" + targetMingguList.join(", ") + "). Sila pastikan anda telah menjana RPH bagi minggu berkaitan.");
    }

    var safeNamaGuru = String(namaGuru).replace(/[/\\?%*:|"<>]/g, '').trim().replace(/\s+/g, '_');
    var namaFailPdf = "eRPH_BULANAN_" + sBulan.toUpperCase() + "_" + sTahun + "_" + safeNamaGuru + ".pdf";

    // Bina HTML Kompilasi Sebulan
    var htmlContent = `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8">
          <style>
            @page { size: A4 portrait; margin: 12mm 10mm 12mm 10mm; }
            body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; font-size: 9.5px; color: #0f172a; margin: 0; line-height: 1.35; }
            .doc-header { text-align: center; border-bottom: 2px solid #1e3a8a; padding-bottom: 6px; margin-bottom: 12px; }
            .school-title { font-size: 13px; font-weight: 800; color: #1e3a8a; text-transform: uppercase; }
            .sub-title { font-size: 10.5px; font-weight: 700; color: #334155; }
            .meta-table { width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 8.5px; }
            .meta-table td { border: 1px solid #cbd5e1; padding: 3px 6px; }
            .meta-lbl { font-weight: bold; background-color: #f1f5f9; width: 22%; }
            .section-badge { background: #1e3a8a; color: white; padding: 4px 8px; font-weight: 800; font-size: 10px; border-radius: 4px; margin: 10px 0 6px 0; }
            .slot-box { border: 1px solid #cbd5e1; border-radius: 4px; margin-bottom: 8px; page-break-inside: avoid; }
            .slot-title { background: #f8fafc; font-weight: 800; color: #1e3a8a; padding: 3px 6px; font-size: 9px; border-bottom: 1px solid #cbd5e1; }
            .slot-content { width: 100%; border-collapse: collapse; font-size: 8.5px; }
            .slot-content td { border: 1px solid #e2e8f0; padding: 3px 5px; vertical-align: top; }
            .field-label { font-weight: bold; width: 22%; background: #f8fafc; color: #334155; }
            .sign-table { width: 100%; margin-top: 15px; border-collapse: collapse; font-size: 8.5px; page-break-inside: avoid; }
            .sign-line { border-top: 1px dashed #475569; width: 75%; padding-top: 4px; }
          </style>
        </head>
        <body>
          <div class="doc-header">
            <div class="school-title">__NAMA_SEKOLAH__</div>
            <div class="sub-title">REKOD PENGAJARAN &amp; PEMBELAJARAN HARIAN (e-RPH) &mdash; JANAAN SEBULAN PENUH</div>
            <div style="font-size: 7.5px; color: #64748b;">__ALAMAT_SEKOLAH__ &bull; Kod Sekolah: __KOD_SEKOLAH__</div>
          </div>
          <table class="meta-table">
            <tr>
              <td class="meta-lbl">NAMA GURU</td>
              <td><b>${escapeHtmlGas_(namaGuru)}</b></td>
              <td class="meta-lbl">BULAN &amp; TAHUN</td>
              <td><b>${escapeHtmlGas_(sBulan.toUpperCase())} ${sTahun}</b></td>
            </tr>
            <tr>
              <td class="meta-lbl">EMEL GURU</td>
              <td>${escapeHtmlGas_(emel)}</td>
              <td class="meta-lbl">JUMLAH MINGGU</td>
              <td><b>${semuaRekodBulan.length} Minggu Lengkap</b></td>
            </tr>
          </table>
    `;

    semuaRekodBulan.forEach(function(item) {
      htmlContent += `<div class="section-badge">&#128197; ${escapeHtmlGas_(item.minggu.toUpperCase())}</div>`;
      item.rekod.forEach(function(r, idx) {
        var subUpper = String(r.subjek || "").toUpperCase();
        var isEnglish = subUpper.includes("INGGERIS") || subUpper.includes("ENGLISH") || subUpper.includes("[BI]") || subUpper === "BI";
        htmlContent += `
          <div class="slot-box">
            <div class="slot-title">
              ${idx+1}. ${escapeHtmlGas_(r.hari)} (${escapeHtmlGas_(r.tarikh || '-')}) | MASA: ${escapeHtmlGas_(r.mula)} - ${escapeHtmlGas_(r.tamat)} | KELAS: ${escapeHtmlGas_(r.kelas)} &mdash; ${escapeHtmlGas_(r.subjek)}
            </div>
            <table class="slot-content">
              <tr>
                <td class="field-label">${isEnglish ? 'THEME / TOPIC' : 'TEMA / TAJUK'}</td>
                <td>${escapeHtmlGas_(r.temaTajuk || '-')}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? 'LEARNING STANDARD' : 'STANDARD PEMBELAJARAN'}</td>
                <td>${escapeHtmlGas_(r.sp || '-')}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? 'OBJECTIVES' : 'OBJEKTIF PEMBELAJARAN'}</td>
                <td>${escapeHtmlGas_(r.objektif || '-')}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? 'ACTIVITIES' : 'AKTIVITI PdP'}</td>
                <td>${escapeHtmlGas_(r.aktiviti || '-').replace(/\n/g, '<br>')}</td>
              </tr>
              <tr>
                <td class="field-label">${isEnglish ? 'REFLECTION' : 'REFLEKSI'}</td>
                <td>${escapeHtmlGas_(r.refleksi || 'Refleksi direkodkan secara digital.')}</td>
              </tr>
            </table>
          </div>
        `;
      });
    });

    htmlContent += `
        <table class="sign-table">
          <tr>
            <td style="width: 50%;">
              <br><br>
              <div class="sign-line">
                Tandatangan Guru: <b>${escapeHtmlGas_(namaGuru)}</b><br>
                Tarikh: ${Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd/MM/yyyy")}
              </div>
            </td>
            <td style="width: 50%; text-align: right;">
              <br><br>
              <div class="sign-line" style="margin-left: auto;">
                Pengesahan Pentadbir __NAMA_SEKOLAH__<br>
                Status: <b>DIARKIBKAN — RUJUK REKOD SEMAKAN</b>
              </div>
            </td>
          </tr>
        </table>
      </body>
      </html>
    `;

    var htmlBlob = Utilities.newBlob(brandHtml_(htmlContent), 'text/html', 'rph_bulanan.html');
    var pdfBlob = htmlBlob.getAs('application/pdf').setName(namaFailPdf);
    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());

    // Simpan ke hierarki Google Drive rasmi
    var folderArkib = dapatkanAtauCiptaFolderArkib_(sTahun, namaGuru, sBulan);
    var pdfFile = folderArkib.createFile(pdfBlob);
    try {
      
    } catch (eShare) {
      console.warn("Share link info: " + eShare.message);
    }

    var urlPdf = pdfFile.getUrl();
    var downloadUrl = pdfFile.getUrl().replace('view?usp=drivesdk', 'export?format=pdf');
    var urlFolder = folderArkib.getUrl();

    // Catat ke lembaran ARKIB_ERPH_GURU
    simpanRekodArkibErph_({
      tahun: sTahun,
      namaGuru: namaGuru,
      emelGuru: emel,
      mingguAtauBulan: sBulan + " (Sebulan Penuh)",
      jenis: "BULANAN",
      urlPdf: urlPdf,
      urlFolder: urlFolder,
      statusSemakan: "DIARKIBKAN",
      ulasanPentadbir: "Kompilasi sebulan penuh telah berjaya diarkibkan."
    });

    return {
      status: "SUCCESS",
      urlPdf: urlPdf,
      downloadUrl: downloadUrl,
      urlFolder: urlFolder,
      base64: pdfBase64,
      namaFail: namaFailPdf,
      bulan: sBulan,
      tahun: sTahun,
      pesanan: "Kompilasi e-RPH sebulan penuh bagi bulan " + sBulan + " berjaya dijana dan disimpan ke Google Drive!"
    };
  } catch (err) {
    throw new Error("Gagal menjana e-RPH Bulanan: " + err.message);
  }
}

// ==========================================================================
// MODUL GURU BERTUGAS MINGGUAN & RUMUSAN PERHIMPUNAN RASMI (HEM)
// ==========================================================================

/**
 * Simpan Laporan Guru Bertugas Harian Lengkap (Cuaca, Kebersihan, Disiplin, Peristiwa)
 */
function simpanLaporanBertugasLengkapBackend_(payload) {
  try {
    var ss = sekolahSpreadsheet_();
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
function dapatkanRumusanMingguanBertugas_(minggu) {
  try {
    var ss = sekolahSpreadsheet_();
    var sheet = ss.getSheetByName("LAPORAN_BERTUGAS_LENGKAP");
    var rekodHari = [];

    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (padanMingguSama_(data[i][2], minggu)) {
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
      "Bismillahirahmanirrahim. Assalamualaikum Warahmatullahi Wabarakatuh, salam sejahtera dan salam __NAMA_SEKOLAH__ Cemerlang.\n\n" +
      "Yang Berusaha Guru Besar __NAMA_SEKOLAH__, Barisan Penolong Kanan, rakan-rakan guru yang dihormati serta anak-anak murid yang dikasihi sekalian.\n\n" +
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
      "Sekian sahaja laporan daripada barisan Guru Bertugas bagi " + minggu + ". Terima kasih atas kerjasama semua warga __NAMA_SEKOLAH__.";

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
function janaPdfRumusanBertugasMingguan_(minggu, namaGuru, emelGuru) {
  try {
    var rumusan = dapatkanRumusanMingguanBertugas_(minggu);
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
                <div class="school-name">__NAMA_SEKOLAH__</div>
                <div class="report-title">LAPORAN MINGGUAN GURU BERTUGAS &amp; SAHSIAH HEM</div>
                <div style="font-size: 7.5px; color: #64748b;">__ALAMAT_SEKOLAH__ &bull; Kod Sekolah: __KOD_SEKOLAH__</div>
              </td>
            </tr>
          </table>

          <table class="meta-box">
            <tr>
              <td class="meta-lbl">MINGGU BERTUGAS</td>
              <td><b>${escapeHtmlGas_(minggu)}</b></td>
              <td class="meta-lbl">GURU BERTUGAS</td>
              <td><b>${escapeHtmlGas_(namaGuru || "Barisan Guru Bertugas")}</b></td>
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
              <p style="margin: 0;"><b>Peristiwa Penting / Program Sekolah:</b> ${rumusan.senaraiPeristiwa && rumusan.senaraiPeristiwa.length > 0 ? escapeHtmlGas_(rumusan.senaraiPeristiwa.join('; ')) : 'Tiada program luar jangkaan dilaporkan.'}</p>
            </div>
          </div>

          <div class="section-box">
            <div class="section-header">3. TEKS UCAPAN PERHIMPUNAN RASMI HARI ISNIN</div>
            <div class="section-body">
              <div class="speech-box">${escapeHtmlGas_(rumusan.teksUcapanPerhimpunan)}</div>
            </div>
          </div>

          <table class="sign-table">
            <tr>
              <td>
                <br><br>
                <div class="sign-box">
                  Disediakan Oleh:<br><br><br>
                  <b>( ${escapeHtmlGas_(namaGuru || "GURU BERTUGAS MINGGUAN")} )</b><br>
                  Guru Bertugas Mingguan __NAMA_SEKOLAH__
                </div>
              </td>
              <td style="text-align: right;">
                <br><br>
                <div class="sign-box" style="margin-left: auto; text-align: left;">
                  Disahkan Oleh:<br><br><br>
                  <b>( GURU BESAR / PK HEM )</b><br>
                  __NAMA_SEKOLAH__
                </div>
              </td>
            </tr>
          </table>
        </body>
      </html>
    `;

    var htmlBlob = Utilities.newBlob(brandHtml_(htmlContent), 'text/html', 'laporan_bertugas.html');
    var pdfBlob = htmlBlob.getAs('application/pdf').setName(namaFailPdf);
    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());

    // Simpan fail ke Google Drive
    var pdfFile = saveBlob_(pdfBlob,{jenisDokumen:"OPR",bahagian:"HEM"});
    try {
      
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


// ===== Kokurikulum.gs =====
// ==========================================
// KOD BACKEND KOKURIKULUM & KEHADIRAN (Kokurikulum.gs)
// PORTAL DIGITAL __NAMA_SEKOLAH__ 2026
// PENGURUSAN UNIT BERUNIFORM, KELAB & PERSATUAN, 1M1S
// ==========================================

// --------------------------------------------------------------------------
// 1. SISTEM PENGECAMAN KATEGORI & TAB SHEET GOOGLE SHEETS
// --------------------------------------------------------------------------

// Pengecaman Kod Warna Tab Google Sheets (RGB -> HSL)
// 1. Unit Beruniform: Tab Kuning (Yellow)
// 2. Kelab & Persatuan: Tab Hijau (Green)
// 3. 1M1S: Tab Oren (Orange)
function klasifikasiWarnaTab_(colorHex) {
  if (!colorHex) return null;
  var hex = String(colorHex).toLowerCase().replace('#', '').trim();
  if (hex.length === 3) {
    hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
  }
  if (hex.length !== 6) return null;

  var r = parseInt(hex.substring(0, 2), 16);
  var g = parseInt(hex.substring(2, 4), 16);
  var b = parseInt(hex.substring(4, 6), 16);

  var rNorm = r / 255, gNorm = g / 255, bNorm = b / 255;
  var max = Math.max(rNorm, gNorm, bNorm), min = Math.min(rNorm, gNorm, bNorm);
  var h = 0, s = 0, l = (max + min) / 2;

  if (max !== min) {
    var d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case rNorm: h = (gNorm - bNorm) / d + (gNorm < bNorm ? 6 : 0); break;
      case gNorm: h = (bNorm - rNorm) / d + 2; break;
      case bNorm: h = (rNorm - gNorm) / d + 4; break;
    }
    h = h * 60; // Sudut Hue 0 - 360 darjah
  }

  // Pengelasan mengikut Hue:
  // Oren (Orange): Hue 14Â° hingga 44Â° -> 1M1S
  if (h >= 14 && h < 44 && s > 0.20) {
    return "1M1S";
  }

  // Kuning (Yellow): Hue 44Â° hingga 72Â° -> Unit Beruniform
  if (h >= 44 && h < 72 && s > 0.20) {
    return "BERUNIFORM";
  }

  // Hijau (Green): Hue 72Â° hingga 175Â° -> Kelab & Persatuan
  if (h >= 72 && h <= 175 && s > 0.18) {
    return "KELAB";
  }

  return null;
}

// Pengecaman Pintar Kategori Tab Mengikut Warna Tab & Kata Kunci Nama Tab
function kenalPastiKategoriTab_(sheet) {
  var nama = sheet.getName().trim();
  var namaUpper = nama.toUpperCase();

  // Tab-tab sistem yang dikecualikan
  var tabSistem = [
    "LAPORAN_OPR", "KEHADIRAN_GURU", "STATUS_FEED", "KEBERHASILAN_GURU", "ARKIB_ERPH_GURU", "LAPORAN_BERTUGAS_LENGKAP", "PENGGUNA", "RPH_GURU", "TAKWIM", "DSKP", "DSKP_PPKI", 
    "LAPORAN_BERTUGAS", "LAPORAN_KOKUM", "SEMAKAN_RPH", "DATABASE", "KEHADIRAN", "INDEX"
  ];
  if (tabSistem.includes(namaUpper)) return null;

  // 1. Keutamaan Pertama: Warna Tab Google Sheets
  try {
    var tabColor = sheet.getTabColor();
    if (tabColor) {
      var katWarna = klasifikasiWarnaTab_(tabColor);
      if (katWarna) return katWarna;
    }
  } catch (e) {
    console.warn("Ralat semak getTabColor: " + e.message);
  }

  // 2. Keutamaan Kedua: Pengecaman Berdasarkan Kata Kunci Nama Tab
  if (namaUpper.includes("PENGAKAP") || namaUpper.includes("TKRS") || namaUpper.includes("BSMM") || 
      namaUpper.includes("PUTERI") || namaUpper.includes("KADET") || namaUpper.includes("BERUNIFORM") || 
      namaUpper.includes("PPIM") || namaUpper.includes("BULAN SABIT") || namaUpper.includes("PANDU") || 
      namaUpper.includes("UNIFORM")) {
    return "BERUNIFORM";
  }

  if (namaUpper.includes("1M1S") || namaUpper.includes("BOLA") || namaUpper.includes("TAKRAW") || 
      namaUpper.includes("BADMINTON") || namaUpper.includes("OLAHRAGA") || namaUpper.includes("FUTSAL") || 
      namaUpper.includes("JARING") || namaUpper.includes("PING PONG") || namaUpper.includes("CATUR") || 
      namaUpper.includes("HOKI") || namaUpper.includes("SUKAN") || namaUpper.includes("PERMAINAN")) {
    return "1M1S";
  }

  if (namaUpper.includes("KELAB") || namaUpper.includes("PERSATUAN") || namaUpper.includes("STEM") || 
      namaUpper.includes("DOKTOR") || namaUpper.includes("BAHASA") || namaUpper.includes("KESENIAN") || 
      namaUpper.includes("KEBUDAYAAN") || namaUpper.includes("BUDAYA") || namaUpper.includes("SENI") || namaUpper.includes("AGAMA") || namaUpper.includes("ISLAM") || 
      namaUpper.includes("SPBT") || namaUpper.includes("RUKUN NEGARA") || namaUpper.includes("ROBOTIK")) {
    return "KELAB";
  }

  return null;
}

// Senarai Lalai Rasmi Unit Kokurikulum __NAMA_SEKOLAH__ (Fallback Pintar)
var SENARAI_DEFAULT_KOKUM = {
  "BERUNIFORM": [
    { id: "Pengakap Kanak-Kanak", nama: "Pengakap Kanak-Kanak", warna: "Kuning" },
    { id: "Tunas Kadet Remaja Sekolah", nama: "Tunas Kadet Remaja Sekolah (TKRS)", warna: "Kuning" },
    { id: "Bulan Sabit Merah Malaysia", nama: "Bulan Sabit Merah Malaysia (BSMM)", warna: "Kuning" },
    { id: "Pandu Puteri Tunas", nama: "Pandu Puteri Tunas", warna: "Kuning" },
    { id: "Pergerakan Puteri Islam", nama: "Pergerakan Puteri Islam Malaysia (PPIM)", warna: "Kuning" }
  ],
  "KELAB": [
    { id: "Persatuan Bahasa Melayu", nama: "Persatuan Bahasa Melayu", warna: "Hijau" },
    { id: "Persatuan Bahasa Inggeris", nama: "Persatuan Bahasa Inggeris", warna: "Hijau" },
    { id: "Kelab STEM & Sains", nama: "Kelab STEM & Sains", warna: "Hijau" },
    { id: "Persatuan Pendidikan Islam", nama: "Persatuan Pendidikan Islam", warna: "Hijau" },
    { id: "Kelab Doktor Muda", nama: "Kelab Doktor Muda", warna: "Hijau" },
    { id: "Kelab Kesenian & Kebudayaan", nama: "Kelab Kesenian & Kebudayaan", warna: "Hijau" }
  ],
  "1M1S": [
    { id: "Kelab Bola Sepak", nama: "Kelab Bola Sepak / Futsal", warna: "Oren" },
    { id: "Kelab Bola Jaring", nama: "Kelab Bola Jaring", warna: "Oren" },
    { id: "Kelab Sepak Takraw", nama: "Kelab Sepak Takraw", warna: "Oren" },
    { id: "Kelab Badminton", nama: "Kelab Badminton", warna: "Oren" },
    { id: "Kelab Olahraga", nama: "Kelab Olahraga & Merentas Desa", warna: "Oren" }
  ]
};

// Fungsi Utama: Mengimbas & Menyenaraikan Semua Tab Mengikut 3 Kategori Utama
function getKategoriDanUnitKokum_() { return configuredKoko_(false); }

// Alias untuk keserasian panggilan sedia ada
function getSenaraiTabKokum_() {
  return getKategoriDanUnitKokum_();
}

function getSenaraiUnitKokumWeb_() {
  return getKategoriDanUnitKokum_();
}

// ==========================================================================
// PENGURUSAN UNIT KOKURIKULUM PPKI (PENDIDIKAN KHAS MASALAH PEMBELAJARAN)
// ==========================================================================
var SENARAI_DEFAULT_KOKUM_PPKI = {
  "BERUNIFORM": [
    { id: "PPKI PENGAKAP", nama: "PPKI PENGAKAP", warna: "Kuning" }
  ],
  "1M1S": [
    { id: "PPKI BADMINTON", nama: "PPKI BADMINTON", warna: "Oren" }
  ],
  "KELAB": [
    { id: "PPKI SENI BUDAYA", nama: "PPKI SENI BUDAYA", warna: "Hijau" }
  ]
};

// Imbas Google Sheets Khusus Bagi Unit Kokurikulum PPKI
function getKategoriDanUnitKokumPpki_() { return configuredKoko_(true); }

function getSenaraiTabKokumPpki_() {
  return getKategoriDanUnitKokumPpki_();
}

// --------------------------------------------------------------------------
// 2. PENGAMBILAN DATA MURID & STATUS KEHADIRAN SEMASA (READ)
// --------------------------------------------------------------------------

// Helper: Cari Lajur Minggu / Perjumpaan Pintar ("MINGGU 1", "M1", "PERJUMPAAN 1", "P1", dsb.)
function cariLajurMinggu_(headerBaris, mingguKe) {
  if (!headerBaris || !mingguKe) return -1;
  var numStr = String(mingguKe).replace(/\D/g, '') || String(mingguKe).trim();
  
  var corakCari = [
    "MINGGU " + numStr,
    "M" + numStr,
    "M " + numStr,
    "PERJUMPAAN " + numStr,
    "P" + numStr,
    "P " + numStr,
    "BIL " + numStr,
    "BIL. " + numStr,
    numStr
  ];

  // 1. Padanan Tepat
  for (var c = 0; c < headerBaris.length; c++) {
    var val = String(headerBaris[c] || "").trim().toUpperCase();
    for (var k = 0; k < corakCari.length; k++) {
      if (val === corakCari[k].toUpperCase()) {
        return c + 1; // 1-indexed
      }
    }
  }

  // 2. Padanan Mengandungi Nombor Minggu
  for (var c = 0; c < headerBaris.length; c++) {
    var val = String(headerBaris[c] || "").trim().toUpperCase();
    if (val.includes("MINGGU") || val.includes("PERJUMPAAN") || val.includes("M") || val.includes("P")) {
      var digitDalamSel = val.replace(/\D/g, '');
      if (digitDalamSel === numStr) {
        return c + 1;
      }
    }
  }

  return -1;
}

// Dapatkan atau Cipta Lajur Minggu Secara Automatik jika belum wujud
function dapatkanAtauCiptaLajurMinggu_(sheet, headerRowIdx, mingguKe) {
  var data = sheet.getDataRange().getDisplayValues();
  var header = data[headerRowIdx] || [];
  var targetCol = cariLajurMinggu_(header, mingguKe);
  if (targetCol !== -1) return targetCol;

  // Jika tiada pada baris header semasa, semak baris 0 atau baris 1
  if (headerRowIdx > 0) {
    targetCol = cariLajurMinggu_(data[0], mingguKe);
    if (targetCol !== -1) return targetCol;
  }

  // Jika belum wujud, cipta lajur baharu di penghujung helaian
  var newCol = Math.max(sheet.getLastColumn() + 1, 4);
  var labelHeader = "M" + mingguKe;
  sheet.getRange(headerRowIdx + 1, newCol).setValue(labelHeader);
  sheet.getRange(headerRowIdx + 1, newCol).setFontWeight("bold").setHorizontalAlignment("center");
  return newCol;
}

// Ambil Senarai Murid dan Semak Status Kehadiran Minggu Terpilih
// Menggunakan .getDisplayValues() untuk memelihara teks mentah tanpa masalah Date zon masa
// Ambil Senarai Murid dan Semak Status Kehadiran Minggu Terpilih
// Menggunakan .getDisplayValues() untuk memelihara teks mentah tanpa masalah Date zon masa
function getSenaraiMurid_(tabName, mingguKe) {
  try {
    var ss = sekolahSpreadsheet_();
    var sheet = ss.getSheetByName(tabName);
    
    // 1. Jika helaian wujud dalam spreadsheet, baca data sebenar daripada helaian guru
    if (sheet) {
      var data = sheet.getDataRange().getDisplayValues();
      if (data && data.length >= 2) {
        // Pengesanan Lajur Nama & Kelas secara dinamik
        var colNama = 1;      // Default: Lajur B (indeks 1)
        var colKelas = 2;     // Default: Lajur C (indeks 2)
        var headerRowIdx = 1; // Default: Baris 2 (indeks 1)

        // Imbas 4 baris teratas untuk mengenal pasti struktur jadual
        for (var r = 0; r < Math.min(data.length, 4); r++) {
          for (var c = 0; c < data[r].length; c++) {
            var val = String(data[r][c] || "").trim().toUpperCase();
            if (val === "NAMA" || val === "NAMA AHLI" || val === "NAMA MURID" || val.includes("NAMA")) {
              colNama = c;
              headerRowIdx = r;
            }
            if (val === "KELAS" || val === "TAHUN" || val.includes("KELAS")) {
              colKelas = c;
            }
          }
        }

        var rowMula = headerRowIdx + 1;

        // Cari lajur minggu / perjumpaan yang dipilih jika dibekalkan
        var colMinggu = -1;
        if (mingguKe) {
          colMinggu = cariLajurMinggu_(data[headerRowIdx], mingguKe);
          if (colMinggu === -1 && headerRowIdx > 0) {
            colMinggu = cariLajurMinggu_(data[0], mingguKe);
          }
        }

        var senarai = [];
        var noBil = 1;

        for (var i = rowMula; i < data.length; i++) {
          var namaMurid = data[i][colNama] ? data[i][colNama].trim() : "";
          var kelasMurid = data[i][colKelas] ? data[i][colKelas].trim() : "PPKI";

          // Abaikan baris kosong, baris tajuk, atau baris jumlah
          if (namaMurid !== "" && 
              !namaMurid.toUpperCase().startsWith("NAMA") && 
              !namaMurid.toUpperCase().startsWith("JUMLAH") && 
              !namaMurid.toUpperCase().startsWith("CATATAN")) {

            var statusHadir = 1; // Default ditandakan hadir
            if (colMinggu !== -1 && (colMinggu - 1) < data[i].length) {
              var valTeks = String(data[i][colMinggu - 1] || "").trim();
              if (valTeks === "0") {
                statusHadir = 0;
              } else if (valTeks === "1") {
                statusHadir = 1;
              }
            }

            senarai.push({
              bil: noBil++,
              nama: namaMurid,
              kelas: kelasMurid,
              hadir: statusHadir
            });
          }
        }

        if (senarai.length > 0) {
          return senarai;
        }
      }
    }

    return [];
  } catch (err) {
    console.error("Ralat ambil senarai murid: " + err.message);
    return [];
  }
}

// Alias untuk keserasian panggilan sedia ada
function getSenaraiMuridKokum_(kategori, tabName, mingguKe) {
  return getSenaraiMurid_(tabName, mingguKe);
}

// --------------------------------------------------------------------------
// 3. PENYIMPANAN KEHADIRAN (WRITE) & PEMFORMATAN BERSYARAT (1=Hijau, 0=Merah)
// --------------------------------------------------------------------------

// Pasang Peraturan Pemformatan Bersyarat (Conditional Formatting) Automatik pada Google Sheets
// Nilai 1 -> Latar Belakang HIJAU lembut (#dcfce7), Teks Hijau Gelap (#15803d)
// Nilai 0 -> Latar Belakang MERAH lembut (#fee2e2), Teks Merah Gelap (#b91c1c)
function pasangConditionalFormattingKoko_(sheet, colIndex, startRow, numRows) {
  try {
    var range = sheet.getRange(startRow, colIndex, numRows, 1);
    var rules = sheet.getConditionalFormatRules();
    
    var ruleHadirWujud = false;
    var ruleTidakHadirWujud = false;

    for (var i = 0; i < rules.length; i++) {
      var r = rules[i];
      var cond = r.getBooleanCondition();
      if (cond && cond.getCriteriaType() === SpreadsheetApp.BooleanCriteria.NUMBER_EQUAL_TO) {
        var args = cond.getCriteriaValues();
        if (args && args[0] == 1) ruleHadirWujud = true;
        if (args && args[0] == 0) ruleTidakHadirWujud = true;
      }
    }

    if (!ruleHadirWujud) {
      var ruleHadir = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberEqualTo(1)
        .setBackground("#dcfce7") // Hijau lembut
        .setFontColor("#15803d")   // Teks hijau gelap
        .setBold(true)
        .setRanges([range])
        .build();
      rules.push(ruleHadir);
    }

    if (!ruleTidakHadirWujud) {
      var ruleTidakHadir = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberEqualTo(0)
        .setBackground("#fee2e2") // Merah lembut
        .setFontColor("#b91c1c")   // Teks merah gelap
        .setBold(true)
        .setRanges([range])
        .build();
      rules.push(ruleTidakHadir);
    }

    sheet.setConditionalFormatRules(rules);
  } catch (e) {
    console.warn("Ralat pasang conditional formatting: " + e.message);
  }
}

// Simpan Kehadiran Koko Pukal: Mengemas kini lajur minggu dan menetapkan warna 1=Hijau, 0=Merah
function simpanKehadiranKoko_(tabName, mingguKe, dataKehadiran) {
  try {
    var ss = sekolahSpreadsheet_();
    var sheet = ss.getSheetByName(tabName);

    if (!sheet) {
      throw new Error("Tab unit '" + tabName + "' tidak dijumpai dalam Google Sheets.");
    }

    var data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) {
      throw new Error("Format tab '" + tabName + "' tidak lengkap.");
    }

    // Kenal pasti baris header & lajur nama
    var colNama = 1;
    var headerRowIdx = 1;

    for (var r = 0; r < Math.min(data.length, 3); r++) {
      for (var c = 0; c < data[r].length; c++) {
        var val = String(data[r][c] || "").trim().toUpperCase();
        if (val === "NAMA" || val === "NAMA AHLI" || val === "NAMA MURID" || val.includes("NAMA")) {
          colNama = c;
          headerRowIdx = r;
          break;
        }
      }
    }

    // Dapatkan atau cipta lajur minggu secara automatik
    var targetCol = dapatkanAtauCiptaLajurMinggu_(sheet, headerRowIdx, mingguKe);

    // Petakan setiap nama murid kepada nombor baris dalam helaian (1-indexed)
    var mapNamaKeBaris = {};
    for (var r = headerRowIdx + 1; r < data.length; r++) {
      if (data[r][colNama]) {
        var nNorm = String(data[r][colNama]).trim().toUpperCase();
        mapNamaKeBaris[nNorm] = r + 1;
      }
    }

    var jumlahDikemaskini = 0;
    var rowMin = 999999;
    var rowMax = 0;

    // Catatkan nilai 1 atau 0 serta warna sel secara langsung
    dataKehadiran.forEach(function(item) {
      var nCari = String(item.nama || "").trim().toUpperCase();
      var baris = mapNamaKeBaris[nCari];

      if (baris) {
        var nilaiAngka = (item.hadir === true || item.hadir === 1 || item.hadir === "1") ? 1 : 0;
        var cell = sheet.getRange(baris, targetCol);
        
        cell.setValue(nilaiAngka);
        cell.setHorizontalAlignment("center");
        cell.setFontWeight("bold");

        // Warna Latar Lembut: Hijau (#dcfce7) untuk 1, Merah (#fee2e2) untuk 0
        if (nilaiAngka === 1) {
          cell.setBackground("#dcfce7");
          cell.setFontColor("#15803d");
        } else {
          cell.setBackground("#fee2e2");
          cell.setFontColor("#b91c1c");
        }

        if (baris < rowMin) rowMin = baris;
        if (baris > rowMax) rowMax = baris;
        jumlahDikemaskini++;
      }
    });

    // Pasangkan juga Conditional Formatting Rule pada lajur tersebut
    if (rowMin <= rowMax) {
      var totalBaris = (rowMax - rowMin) + 1;
      pasangConditionalFormattingKoko_(sheet, targetCol, rowMin, totalBaris);
    }

    return {
      status: "SUCCESS",
      jumlahDikemaskini: jumlahDikemaskini,
      pesanan: `Kehadiran ${jumlahDikemaskini} orang murid bagi ${tabName} (Minggu ${mingguKe}) berjaya disimpan!`
    };
  } catch (err) {
    throw new Error("Ralat simpan kehadiran: " + err.message);
  }
}

// Alias untuk keserasian panggilan sedia ada
function simpanKehadiranKokumPukalBackend_(kategori, tabNama, mingguKe, senaraiKehadiran) {
  return simpanKehadiranKoko_(tabNama, mingguKe, senaraiKehadiran);
}

// Simpan kehadiran individu (Fallback)
function simpanKehadiranKokumBackend_(kategori, tabNama, namaMurid, mingguKe, statusHadir) {
  var nilai = (statusHadir === true || statusHadir === 1 || statusHadir === "1") ? 1 : 0;
  return simpanKehadiranKoko_(tabNama, mingguKe, [{ nama: namaMurid, hadir: nilai }]);
}

// --------------------------------------------------------------------------
// 4. MODUL LAPORAN OPR KOKURIKULUM (ONE PAGE REPORT)
// --------------------------------------------------------------------------

// Janaan Teks Automatik bagi 5 Teras OPR Kokurikulum
function janaKandunganOPRKokum_(tajuk, unit) {
  var tUpper = String(tajuk || "").toUpperCase();
  var uUpper = String(unit || "").toUpperCase();
  
  var obj1 = "Mendedahkan murid kepada pengetahuan dan kemahiran asas mengenai " + (tajuk || "aktiviti unit") + ".";
  var obj2 = "Meningkatkan tahap disiplin, kepimpinan dan semangat berpasukan dalam kalangan ahli.";
  var obj3 = "Memupuk nilai jati diri, keberanian serta penglibatan aktif semua murid.";
  
  var langkah1 = "Guru penasihat memberi penerangan konsep, taklimat keselamatan dan demonstrasi aktiviti.";
  var langkah2 = "Murid dibahagikan kepada kumpulan kecil bagi sesi amali dan latihan berpandu.";
  var langkah3 = "Setiap kumpulan melaksanakan tugasan secara bergilir-gilir di bawah bimbingan guru penasihat.";
  var langkah4 = "Sesi rumusan aktiviti, refleksi murid dan maklum balas guru bertugas.";
  
  var impak1 = "Majoriti murid berjaya menguasai kemahiran yang dipelajari dengan yakin dan selamat.";
  var impak2 = "Kehadiran dan kerjasama murid amat memuaskan serta menunjukkan minat mendalam.";
  var impak3 = "Disiplin dan adab murid sepanjang perjumpaan terkawal dan mematuhi peraturan.";
  
  var isu1 = "Segelintir murid memerlukan bimbingan secara individu bagi menguasai teknik secara teliti.";
  var isu2 = "Kekangan masa dan giliran peralatan memerlukan susunan stesen yang lebih terperinci.";
  
  var tind1 = "Mengadakan bimbingan rakan sebaya (peer coaching) bersama AJK murid senior.";
  var tind2 = "Menambah bilangan stesen amali agar masa menunggu dapat dikurangkan pada perjumpaan seterusnya.";

  // Penyesuaian khusus mengikut kata kunci topik
  if (tUpper.includes("IKATAN") || tUpper.includes("SIMPULAN") || tUpper.includes("BUKU SILA")) {
    obj1 = "Mendedahkan murid kepada teknik ikatan asas dan simpulan buku sila dengan kaedah yang betul.";
    langkah2 = "Latihan amali mengikat tali secara berpasangan dengan bimbingan guru dan demonstrasi visual.";
    impak1 = "Murid mampu menghasilkan ikatan buku sila dan bunga geti dengan kemas dan pantas.";
  } else if (tUpper.includes("KAWAD") || tUpper.includes("PERBARISAN")) {
    obj1 = "Meningkatkan disiplin, keseragaman pergerakan kawad statik dan kawad dinamik murid.";
    langkah2 = "Latihan kawad kaki mengikut arahan hukuman: sedia, senang diri, luruskan barisan dan pusing.";
    impak1 = "Keseragaman langkah dan keyakinan murid memberi arahan hukuman kawad semakin mantap.";
  } else if (tUpper.includes("PERTOLONGAN") || tUpper.includes("BALUTAN") || tUpper.includes("ANDUH")) {
    obj1 = "Memberi pendedahan praktikal berkenaan prinsip pertolongan cemas dan rawatan awal kecederaan.";
    langkah2 = "Sesi amali menggunakan kain anduh untuk balutan kepala, lengan dan tapak tangan.";
    impak1 = "Murid memahami prosedur keselamatan asas dan cara menggunakan peti pertolongan cemas.";
  } else if (tUpper.includes("BOLA") || tUpper.includes("TAKRAW") || tUpper.includes("BADMINTON") || tUpper.includes("OLAHRAGA")) {
    obj1 = "Menguasai teknik asas kemahiran permainan, kawalan bola/raket dan ketahanan fizikal.";
    langkah2 = "Sesi pemanasan badan, latihan tubi hantaran/pukulan/kawalan dan simulasi permainan kecil.";
    impak1 = "Murid dapat mengaplikasikan undang-undang asas permainan dan memupuk semangat kesukanan.";
  } else if (tUpper.includes("STEM") || tUpper.includes("SAINS") || tUpper.includes("ROBOTIK")) {
    obj1 = "Merangsang daya pemikiran kritis, kreativiti dan penyelesaian masalah melalui projek STEM ringkas.";
    langkah2 = "Aktiviti penerokaan sains secara amali dalam kumpulan menggunakan bahan kitar semula / modul.";
    impak1 = "Murid berupaya membina prototaip dan menerangkan prinsip saintifik di sebalik projek.";
  }

  return {
    objektif: "1. " + obj1 + "\n2. " + obj2 + "\n3. " + obj3,
    pengisian: "1. " + langkah1 + "\n2. " + langkah2 + "\n3. " + langkah3 + "\n4. " + langkah4,
    impak: "1. " + impak1 + "\n2. " + impak2 + "\n3. " + impak3,
    isu: "1. " + isu1 + "\n2. " + isu2,
    tindakan: "1. " + tind1 + "\n2. " + tind2
  };
}

// Simpan Laporan OPR Kokurikulum ke Tab LAPORAN_KOKUM
function simpanPelaporanKokumBackend_(formData) {
  return simpanAtauKemasKiniOpr_({idLaporan:formData.idLaporan,bahagian:ROUTE_&&ROUTE_.bahagian||'KOKURIKULUM',unit:formData.unit,namaProgram:formData.tajuk,penyelaras:konteks_().user.nama,emelPenyelaras:konteks_().user.emel,tarikh:formData.tarikh,masa:formData.minggu?'Minggu '+formData.minggu:'-',tempat:formData.tempat,kehadiran:formData.hadirAhli,objektif:formData.objektif||formData.setInduksi,pengisian:formData.pengisian,impak:formData.impak,isu:formData.isu,tindakan:formData.tindakan||formData.ulasan,gambar1:formData.gambar1||formData.gambar1Url,gambar2:formData.gambar2||formData.gambar2Url,kategoriKoko:formData.kategoriKoko});
}

// Penjanaan PDF OPR Kokurikulum (A4 Portrait - 1 Muka Surat)
function janaPdfOprKokumBackend_(formData) {
  try {
    var unit = formData.unit || "UNIT KOKURIKULUM";
    var tajuk = formData.tajuk || "AKTIVITI PERJUMPAAN KOKURIKULUM";
    var tarikh = formData.tarikh || "-";
    var minggu = formData.minggu ? ("Minggu " + formData.minggu) : "-";
    var tempat = formData.tempat || "Kawasan Sekolah";
    var hadir = formData.hadirAhli || "0";
    var guru = formData.guruHadir || "Guru Penasihat";
    var objektif = formData.objektif || "-";
    var pengisian = formData.pengisian || "-";
    var impak = formData.impak || "-";
    var isu = formData.isu || "-";
    var tindakan = formData.tindakan || "-";
    var g1 = formData.gambar1 || formData.gambar1Url || "";
    var g2 = formData.gambar2 || formData.gambar2Url || "";

    // Tukar pautan gambar Google Drive kepada Base64 Data URI untuk pemaparan PDF selamat
    if (g1 && !String(g1).startsWith("data:image") && typeof ambilBase64DariDrive_ === 'function') {
      var b64_1 = ambilBase64DariDrive_(g1);
      if (b64_1) g1 = b64_1;
    }
    if (g2 && !String(g2).startsWith("data:image") && typeof ambilBase64DariDrive_ === 'function') {
      var b64_2 = ambilBase64DariDrive_(g2);
      if (b64_2) g2 = b64_2;
    }

    function formatPointsHtml(text) {
      if (!text) return '-';
      var lines = String(text).split('\n').filter(function(l){ return l.trim() !== ''; });
      if (lines.length <= 1) return String(text).replace(/\n/g, '<br>');
      return lines.map(function(l){ return '<div style="margin-bottom:2px;">' + l + '</div>'; }).join('');
    }

    var img1Html = g1 ? '<div style="text-align:center; flex:1; padding:3px; border:1px solid #e2e8f0; border-radius:6px; background:#fff;"><img src="' + g1 + '" style="max-height:120px; max-width:100%; object-fit:cover; border-radius:4px;" /><div style="font-size:8px; color:#64748b; margin-top:2px; font-weight:bold;">Gambar 1: Semasa Aktiviti</div></div>' : '<div style="flex:1; border:1px dashed #cbd5e1; border-radius:6px; height:100px; display:flex; align-items:center; justify-content:center; color:#94a3b8; font-size:8.5px; background:#f8fafc;">[Tiada Gambar 1]</div>';
    
    var img2Html = g2 ? '<div style="text-align:center; flex:1; padding:3px; border:1px solid #e2e8f0; border-radius:6px; background:#fff;"><img src="' + g2 + '" style="max-height:120px; max-width:100%; object-fit:cover; border-radius:4px;" /><div style="font-size:8px; color:#64748b; margin-top:2px; font-weight:bold;">Gambar 2: Penglibatan Ahli</div></div>' : '<div style="flex:1; border:1px dashed #cbd5e1; border-radius:6px; height:100px; display:flex; align-items:center; justify-content:center; color:#94a3b8; font-size:8.5px; background:#f8fafc;">[Tiada Gambar 2]</div>';

    var htmlContent = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>OPR Kokurikulum - ${unit}</title>
  <style>
    @page { size: A4 portrait; margin: 8mm 10mm 8mm 10mm; }
    body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; font-size: 9px; color: #1e293b; margin: 0; line-height: 1.3; }
    .header-table { width: 100%; border-bottom: 2px solid #1e3a8a; padding-bottom: 4px; margin-bottom: 6px; }
    .title-main { font-size: 12px; font-weight: 800; color: #1e3a8a; text-transform: uppercase; }
    .title-sub { font-size: 9px; font-weight: 700; color: #475569; }
    .info-grid { width: 100%; border-collapse: collapse; margin-bottom: 6px; font-size: 8.5px; }
    .info-grid td { border: 1px solid #cbd5e1; padding: 3.5px 6px; vertical-align: top; }
    .info-lbl { font-weight: bold; background-color: #f8fafc; color: #334155; width: 18%; }
    .info-val { font-weight: 600; color: #0f172a; width: 32%; }
    .section-box { border: 1px solid #cbd5e1; border-radius: 5px; margin-bottom: 5px; overflow: hidden; }
    .section-title { background: #f1f5f9; font-weight: 800; color: #1e293b; padding: 3px 6px; font-size: 8.5px; border-bottom: 1px solid #cbd5e1; text-transform: uppercase; }
    .section-content { padding: 4px 6px; font-size: 8.5px; color: #1e293b; }
    .footer-table { width: 100%; margin-top: 8px; border-collapse: collapse; font-size: 8px; }
    .footer-table td { width: 50%; vertical-align: top; }
  </style>
</head>
<body>
  <table class="header-table">
    <tr>
      <td style="width: 50px; text-align: center; vertical-align: middle;">
        ${typeof dapatkanLogoSekolah_ === 'function' && dapatkanLogoSekolah_() ? 
          '<img src="' + dapatkanLogoSekolah_() + '" style="max-height:45px; max-width:45px; object-fit:contain;" alt="Lencana __NAMA_SEKOLAH__" />' : 
          '<div style="font-size: 24px;">🏫</div>'}
      </td>
      <td style="vertical-align: middle;">
        <div class="title-main">__NAMA_SEKOLAH__</div>
        <div class="title-sub">LAPORAN SATU MUKA SURAT (ONE PAGE REPORT - OPR) KOKURIKULUM</div>
        <div style="font-size: 7.5px; color: #64748b;">__ALAMAT_SEKOLAH__ • Kod Sekolah: __KOD_SEKOLAH__</div>
      </td>
    </tr>
  </table>

  <table class="info-grid">
    <tr>
      <td class="info-lbl">UNIT KOKURIKULUM</td>
      <td class="info-val">${unit}</td>
      <td class="info-lbl">PERJUMPAAN / MINGGU</td>
      <td class="info-val">${minggu}</td>
    </tr>
    <tr>
      <td class="info-lbl">TARIKH & MASA</td>
      <td class="info-val">${tarikh}</td>
      <td class="info-lbl">TEMPAT PELAKSANAAN</td>
      <td class="info-val">${tempat}</td>
    </tr>
    <tr>
      <td class="info-lbl">GURU PENASIHAT</td>
      <td class="info-val">${guru}</td>
      <td class="info-lbl">KEHADIRAN AHLI</td>
      <td class="info-val"><strong>${hadir} Orang Murid</strong></td>
    </tr>
    <tr>
      <td class="info-lbl">TAJUK / TOPIK</td>
      <td class="info-val" colspan="3" style="font-size: 9px; color: #1e3a8a; font-weight: 800;">${tajuk}</td>
    </tr>
  </table>

  <div class="section-box">
    <div class="section-title">1. OBJEKTIF PERJUMPAAN</div>
    <div class="section-content">${formatPointsHtml(objektif)}</div>
  </div>

  <div class="section-box">
    <div class="section-title">2. PENGISIAN & RINGKASAN AKTIVITI</div>
    <div class="section-content">${formatPointsHtml(pengisian)}</div>
  </div>

  <div class="section-box">
    <div class="section-title">3. IMPAK & KEBERHASILAN MURID</div>
    <div class="section-content">${formatPointsHtml(impak)}</div>
  </div>

  <table style="width:100%; border-collapse:collapse; margin-bottom:5px;">
    <tr>
      <td style="width:49.5%; vertical-align:top; padding-right:3px;">
        <div class="section-box" style="margin-bottom:0;">
          <div class="section-title" style="background:#fff1f2; color:#9f1239; border-color:#fecdd3;">4. ISU & CABARAN</div>
          <div class="section-content" style="min-height:30px;">${formatPointsHtml(isu)}</div>
        </div>
      </td>
      <td style="width:49.5%; vertical-align:top; padding-left:3px;">
        <div class="section-box" style="margin-bottom:0;">
          <div class="section-title" style="background:#f0fdf4; color:#166534; border-color:#bbf7d0;">5. TINDAKAN SUSULAN</div>
          <div class="section-content" style="min-height:30px;">${formatPointsHtml(tindakan)}</div>
        </div>
      </td>
    </tr>
  </table>

  <div class="section-box" style="margin-bottom:5px;">
    <div class="section-title">6. DOKUMENTASI BERGAMBAR AKTIVITI</div>
    <div class="section-content" style="display:flex; gap:6px; padding:4px;">
      ${img1Html}
      ${img2Html}
    </div>
  </div>

  <table class="footer-table">
    <tr>
      <td style="padding-right: 12px;">
        <div>Disediakan Oleh:</div>
        <div style="height: 24px;"></div>
        <div style="font-weight: bold; border-top: 1px dotted #94a3b8; padding-top: 2px;">${guru.toUpperCase()}</div>
        <div style="color: #64748b;">Guru Penasihat ${unit}</div>
        <div style="color: #64748b;">Tarikh: ${tarikh}</div>
      </td>
      <td style="padding-left: 12px;">
        <div>Disahkan Oleh:</div>
        <div style="height: 24px;"></div>
        <div style="font-weight: bold; border-top: 1px dotted #94a3b8; padding-top: 2px;">PENOLONG KANAN KOKURIKULUM</div>
        <div style="color: #64748b;">__NAMA_SEKOLAH__</div>
        <div style="color: #64748b;">Tarikh:</div>
      </td>
    </tr>
  </table>
</body>
</html>`;

    ROUTE_=metaOpr_(formData);

    var namaFailPdf = "OPR_KOKU_" + unit.replace(/\s+/g, '_') + "_" + (formData.minggu ? ("M" + formData.minggu) : "LAPORAN") + ".pdf";
    var pdfBlob = Utilities.newBlob(brandHtml_(htmlContent), 'text/html', 'document.html').getAs(MimeType.PDF).setName(namaFailPdf);
    var pdfFile = saveBlob_(pdfBlob,ROUTE_);

    try {
      
    } catch (errDomain) {
      try {  } catch (eSub) {}
    }

    var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());


    return {
      status: "SUCCESS",
      urlPdf: pdfFile.getUrl(),
      downloadUrl: pdfFile.getUrl().replace('view?usp=drivesdk', 'export?format=pdf'),
      base64: pdfBase64,
      namaFail: namaFailPdf,
      pesanan: "PDF OPR Kokurikulum berjaya dijana!"
    };
  } catch (err) {
    throw new Error("Ralat menjana PDF OPR Kokum: " + err.message);
  }
}


// ===== RpcMap.gs =====
// Generated allowlist. No arbitrary function lookup.
var RPC_ = {
  "muatSemuaDskpDariSheet": muatSemuaDskpDariSheet_,
  "getSenaraiGuruWeb": getSenaraiGuruWeb_,
  "tambahPenggunaBaru": tambahPenggunaBaru_,
  "kemaskiniMaklumatGuru": kemaskiniMaklumatGuru_,
  "simpanFotoProfilGuru": simpanFotoProfilGuru_,
  "getPilihanMingguWeb": getPilihanMingguWeb_,
  "semakAdaRekod": semakAdaRekod_,
  "dapatkanRekodMinggu": dapatkanRekodMinggu_,
  "janaRphSemingguBackend": janaRphSemingguBackend_,
  "kemaskiniRefleksi": kemaskiniRefleksi_,
  "janaPdfMingguanBackend": janaPdfMingguanBackend_,
  "dapatkanDataDashboardPentadbir": dapatkanDataDashboardPentadbir_,
  "simpanSemakanPentadbir": simpanSemakanPentadbir_,
  "simpanLaporanBertugasBackend": simpanLaporanBertugasBackend_,
  "getSenaraiTabKokum": getSenaraiTabKokum_,
  "getSenaraiTabKokumPpki": getSenaraiTabKokumPpki_,
  "getSenaraiMurid": getSenaraiMurid_,
  "simpanKehadiranKoko": simpanKehadiranKoko_,
  "janaKandunganOPRKokumBackend": janaKandunganOPRKokumBackend_,
  "simpanLaporanOprKokumBackend": simpanLaporanOprKokumBackend_,
  "janaPdfOprKokumBackendPortal": janaPdfOprKokumBackendPortal_,
  "dapatkanGambarLaporanOprBase64": dapatkanGambarLaporanOprBase64_,
  "dapatkanLogoSekolah": dapatkanLogoSekolah_,
  "janaKandunganOprPintar": janaKandunganOprPintar_,
  "simpanAtauKemasKiniOpr": simpanAtauKemasKiniOpr_,
  "dapatkanSenaraiOprPentadbir": dapatkanSenaraiOprPentadbir_,
  "sahkanLaporanOprBackend": sahkanLaporanOprBackend_,
  "dapatkanBankOprSekolah": dapatkanBankOprSekolah_,
  "rakamKehadiranGpsBackend": rakamKehadiranGpsBackend_,
  "padamRphMingguanBackend": padamRphMingguanBackend_,
  "simpanJadualGuruBackend": simpanJadualGuruBackend_,
  "dapatkanJadualGuruBackend": dapatkanJadualGuruBackend_,
  "dapatkanStatusKehadiranHariIni": dapatkanStatusKehadiranHariIni_,
  "dapatkanSemuaKehadiranHariIni": dapatkanSemuaKehadiranHariIni_,
  "dapatkanUrlWebApp": dapatkanUrlWebApp_,
  "getFeedStatusWeb": getFeedStatusWeb_,
  "hantarStatusFeed": hantarStatusFeed_,
  "toggleLikeStatusFeed": toggleLikeStatusFeed_,
  "padamStatusFeed": padamStatusFeed_,
  "dapatkanBilanganOnlineLive": dapatkanBilanganOnlineLive_,
  "simpanRekodKeberhasilan": simpanRekodKeberhasilan_,
  "muatRekodKeberhasilanGuru": muatRekodKeberhasilanGuru_,
  "dapatkanSemuaRekodKeberhasilanAdmin": dapatkanSemuaRekodKeberhasilanAdmin_,
  "dapatkanSenaraiArkibRphGuru": dapatkanSenaraiArkibRphGuru_,
  "semakStatusMingguTertunggak": semakStatusMingguTertunggak_,
  "janaRphBulanan": janaRphBulanan_,
  "simpanLaporanBertugasLengkapBackend": simpanLaporanBertugasLengkapBackend_,
  "dapatkanRumusanMingguanBertugas": dapatkanRumusanMingguanBertugas_,
  "janaPdfRumusanBertugasMingguan": janaPdfRumusanBertugasMingguan_
};
