/*****************************************************************
 * MERCADO CIÉNEGA · Backend en Google Apps Script + Sheets
 *
 * Hojas (se crean solas con action=setup):
 *   Usuarios : username | hash | nombre | telefono | rol | must_change
 *   Sesiones : token | username | creado | expira
 *   Categorias : nombre
 *   Productos : id | titulo | categoria | precio | moneda | descripcion |
 *               imagen_ids | vendedor | telefono | whatsapp | municipio |
 *               fecha_pub | fecha_exp | destacado | destacado_hasta | estado
 *   Pedidos : id | fecha | comprador | comprador_tel | items_json |
 *             total_cup | total_usd | nota | estado
 *   Pagos : id | fecha | producto_id | titulo | usuario | metodo |
 *           referencia | monto_cup | estado
 *   Config : clave | valor
 *
 * Las imágenes se guardan en la carpeta de Drive
 * "Imágenes de Mercado Ciénega" (se crea sola al subir la primera).
 *
 * Despliegue: Implementar > Nueva implementación > App web
 *   Ejecutar como: Yo · Acceso: Cualquiera
 *****************************************************************/

var SH = {
  USUARIOS: 'Usuarios', SESIONES: 'Sesiones', CATEGORIAS: 'Categorias',
  PRODUCTOS: 'Productos', PEDIDOS: 'Pedidos', PAGOS: 'Pagos', CONFIG: 'Config'
};
var SESSION_HOURS = 72;
var DRIVE_FOLDER = 'Imágenes de Mercado Ciénega';

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function ensureSheets_() {
  var s = ss_();
  var defs = {};
  defs[SH.USUARIOS]   = ['username','hash','nombre','telefono','rol','must_change'];
  defs[SH.SESIONES]   = ['token','username','creado','expira'];
  defs[SH.CATEGORIAS] = ['nombre'];
  defs[SH.PRODUCTOS]  = ['id','titulo','categoria','precio','moneda','descripcion','imagen_ids',
                         'vendedor','telefono','whatsapp','municipio','fecha_pub','fecha_exp',
                         'destacado','destacado_hasta','estado'];
  defs[SH.PEDIDOS]    = ['id','fecha','comprador','comprador_tel','items_json','total_cup','total_usd','nota','estado'];
  defs[SH.PAGOS]      = ['id','fecha','producto_id','titulo','usuario','metodo','referencia','monto_cup','estado','tipo'];
  defs[SH.CONFIG]     = ['clave','valor'];
  Object.keys(defs).forEach(function(n){
    var sh = s.getSheetByName(n);
    if (!sh) { sh = s.insertSheet(n); sh.appendRow(defs[n]); sh.setFrozenRows(1); }
  });
  // migración: la hoja Pagos anterior tenía 9 columnas, agregar 'tipo'
  try{
    var pg = s.getSheetByName(SH.PAGOS);
    if (pg && pg.getLastColumn()===9) pg.getRange(1,10,1,1).setValue('tipo');
  }catch(e){}
}
function sh_(n){ ensureSheets_(); var s = ss_().getSheetByName(n); if(!s) throw new Error('Falta hoja '+n); return s; }
function rows_(n){
  var v = sh_(n).getDataRange().getValues(); v.shift();
  return v.filter(function(r){ return String(r[0]) !== ''; });
}
function findRow_(n, col, val){
  var v = sh_(n).getDataRange().getValues();
  for (var i=1;i<v.length;i++) if (String(v[i][col])===String(val)) return i+1;
  return -1;
}
function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hoy_(){ var d=new Date(); return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2); }
function sumarMeses_(fechaStr, m){
  var p = fechaStr.split('-'); var d = new Date(Number(p[0]), Number(p[1])-1+m, Number(p[2]));
  return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);
}
function sumarDias_(fechaStr, d_){
  var p = fechaStr.split('-'); var d = new Date(Number(p[0]), Number(p[1])-1, Number(p[2])+d_);
  return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);
}
function newId_(pref){ return pref + Date.now().toString(36).toUpperCase() + Math.floor(Math.random()*1296).toString(36).toUpperCase(); }

/* ---------- hash contraseñas: sha256$salt$hex ---------- */
function hashPw_(pw){
  var salt = Utilities.getUuid().replace(/-/g,'').slice(0,16);
  var hex = bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt+'::'+pw, Utilities.Charset.UTF_8));
  return 'sha256$'+salt+'$'+hex;
}
function verifyPw_(pw, stored){
  try{
    var p = String(stored||'').split('$'); if(p.length!==3||p[0]!=='sha256') return false;
    var hex = bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, p[1]+'::'+pw, Utilities.Charset.UTF_8));
    return hex === p[2];
  }catch(e){ return false; }
}
function bytesToHex_(b){ return b.map(function(x){ var h=(x<0?x+256:x).toString(16); return h.length===1?'0'+h:h; }).join(''); }

/* ---------- config ---------- */
function cfgDefaults_(){
  return {
    tasa_usd: 770, tasa_fecha: '2026-10-02', tasa_fuente: 'El Toque',
    pago_usdt_direccion: 'TBAMgPefcgznuNkanfWmxEkRx6LAmKchUT',
    pago_transfermovil_numero: '',
    pago_efectivo_nota: 'Coordina en persona con el administrador.',
    moderador_token: '',
    destacado_precio_cup: 500, destacado_dias: 30, anuncio_meses: 1,
    publicar_precio_cup: 100
  };
}
function cfgGet_(){
  var d = cfgDefaults_();
  rows_(SH.CONFIG).forEach(function(v){
    var k = String(v[0]); if(!(k in d)) return;
    var val = v[1];
    if (k==='tasa_fecha'||k==='tasa_fuente'||k==='pago_usdt_direccion'||k==='pago_transfermovil_numero'||k==='pago_efectivo_nota') d[k]=String(val);
    else { var n=Number(val); if(!isNaN(n)) d[k]=n; }
  });
  return d;
}
function cfgSet_(k, v){
  var r = findRow_(SH.CONFIG, 0, k), s = sh_(SH.CONFIG);
  if (r<0) s.appendRow([k, v]); else s.getRange(r,2,1,1).setValue(v);
}

/* ---------- usuarios y sesiones ---------- */
function userGet_(username){
  var r = findRow_(SH.USUARIOS, 0, username); if(r<0) return null;
  var v = sh_(SH.USUARIOS).getRange(r,1,1,6).getValues()[0];
  return { username:String(v[0]), hash:String(v[1]), nombre:String(v[2]),
           telefono:String(v[3]), rol:String(v[4]), must_change:String(v[5])==='1' };
}
function sessionCleanup_(){
  try{
    var s=sh_(SH.SESIONES), v=s.getDataRange().getValues(), now=Date.now();
    for(var i=v.length-1;i>=1;i--){ if(!v[i][0] || Number(v[i][3])<now) s.deleteRow(i+1); }
  }catch(e){}
}
function sessionCreate_(username){
  sessionCleanup_();
  var tok = Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
  var now = Date.now();
  sh_(SH.SESIONES).appendRow([tok, username, now, now + SESSION_HOURS*3600*1000]);
  return tok;
}
function sessionGet_(tok){
  if(!tok) return null;
  var r = findRow_(SH.SESIONES, 0, tok); if(r<0) return null;
  var v = sh_(SH.SESIONES).getRange(r,1,1,4).getValues()[0];
  if (Number(v[3]) < Date.now()) { try{ sh_(SH.SESIONES).deleteRow(r); }catch(e){} return null; }
  return String(v[1]);
}
function sessionDelete_(tok){ var r=findRow_(SH.SESIONES,0,tok); if(r>0){ try{ sh_(SH.SESIONES).deleteRow(r); }catch(e){} } }
function authUser_(tok){
  var username = sessionGet_(tok);
  if(!username) return { err:'Sesión inválida o vencida. Entra de nuevo.' };
  var u = userGet_(username);
  if(!u) return { err:'Usuario no existe.' };
  return { user:u };
}
function needAdmin_(tok){
  var a = authUser_(tok);
  if(a.err) return a;
  if(a.user.rol!=='admin') return { err:'Solo el administrador.' };
  return a;
}

/* ---------- productos ---------- */
function prodRow_(v){
  return {
    id:String(v[0]), titulo:String(v[1]), categoria:String(v[2]),
    precio:Number(v[3])||0, moneda:String(v[4])==='USD'?'USD':'CUP',
    descripcion:String(v[5]), imagen_ids:String(v[6]),
    vendedor:String(v[7]), telefono:String(v[8]), whatsapp:String(v[9]),
    municipio:String(v[10]), fecha_pub:String(v[11]), fecha_exp:String(v[12]),
    destacado:String(v[13])==='1', destacado_hasta:String(v[14]),
    estado:String(v[15]||'activo')
  };
}
function prodActivo_(p){
  return p.estado==='activo' && p.fecha_exp >= hoy_();
}
function prodEsDestacado_(p){
  return p.destacado && p.destacado_hasta >= hoy_();
}
function thumb_(id, w){ return 'https://drive.google.com/thumbnail?id='+id+'&sz=w'+(w||800); }
function prodPub_(p){
  var ids = p.imagen_ids ? p.imagen_ids.split(',').filter(function(x){return x;}) : [];
  return {
    id:p.id, titulo:p.titulo, categoria:p.categoria, precio:p.precio, moneda:p.moneda,
    descripcion:p.descripcion, vendedor:p.vendedor, telefono:p.telefono,
    whatsapp:p.whatsapp, municipio:p.municipio, fecha_pub:p.fecha_pub,
    destacado:prodEsDestacado_(p),
    imagenes: ids.map(function(id){ return thumb_(id, 900); }),
    mini: ids.length ? thumb_(ids[0], 400) : ''
  };
}

/* ---------- Drive ---------- */
function driveFolder_(){
  var it = DriveApp.getFoldersByName(DRIVE_FOLDER);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(DRIVE_FOLDER);
}
function guardarImagenes_(imagenes){
  // imagenes: [{nombre, tipo, datos(base64 sin prefijo)}] — ya comprimidas en el navegador
  var folder = driveFolder_(), ids = [];
  (imagenes||[]).slice(0,8).forEach(function(im, i){
    try{
      var bytes = Utilities.base64Decode(String(im.datos||'').replace(/^data:[^,]+,/, ''));
      var ext = (im.tipo||'').indexOf('webp')>=0 ? 'webp' : 'jpg';
      var blob = Utilities.newBlob(bytes, im.tipo||'image/jpeg', 'mc_'+Date.now()+'_'+i+'.'+ext);
      var f = folder.createFile(blob);
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      ids.push(f.getId());
    }catch(e){}
  });
  return ids;
}

/* ---------- tasa El Toque (intento automático, sin bloquear) ---------- */
function tasaAuto_(){
  var cfg = cfgGet_();
  try{
    var r = UrlFetchApp.fetch('https://api.eltoque.com/v1/trmi', {muteHttpExceptions:true, timeout:15000});
    if (r.getResponseCode()===200){
      var j = JSON.parse(r.getContentText());
      var usd = null;
      // varios formatos posibles
      if (j && j.tasas && j.tasas.USD) usd = Number(j.tasas.USD.informal || j.tasas.USD.valor || j.tasas.USD);
      else if (j && j.USD) usd = Number(j.USD.informal || j.USD);
      else if (j && j.usd) usd = Number(j.usd);
      if (usd && usd>50 && usd<5000){
        cfgSet_('tasa_usd', usd); cfgSet_('tasa_fecha', hoy_()); cfgSet_('tasa_fuente', 'El Toque (auto)');
        return { ok:true, tasa:usd, auto:true };
      }
    }
  }catch(e){}
  return { ok:true, tasa:Number(cfg.tasa_usd)||770, fecha:String(cfg.tasa_fecha),
           fuente:String(cfg.tasa_fuente), auto:false };
}

/* ---------- setup inicial (una sola vez) ---------- */
function setup_(p){
  ensureSheets_();
  if (rows_(SH.USUARIOS).length > 0) return { ok:false, error:'Ya está inicializado.' };
  var cats = ['Alimentos','Ropa y calzado','Electrónica','Hogar','Belleza y cuidado','Servicios','Vehículos y repuestos','Mascotas','Otros'];
  var cs = sh_(SH.CATEGORIAS);
  cats.forEach(function(c){ cs.appendRow([c]); });
  var d = cfgDefaults_();
  var cf = sh_(SH.CONFIG);
  Object.keys(d).forEach(function(k){ cf.appendRow([k, d[k]]); });
  var au = String((p&&p.admin_user)||'admin');
  var ap = String((p&&p.admin_pass)||('MC'+Math.floor(100000+Math.random()*900000)));
  sh_(SH.USUARIOS).appendRow([au, hashPw_(ap), 'Administrador', '', 'admin', '1']);
  return { ok:true, admin_user:au, msg:'Mercado Ciénega inicializado. Cambia la clave al entrar.' };
}

/* ---------- enrutador ---------- */
function doGet(e){
  var p = (e&&e.parameter)||{};
  if (p.action==='setup') return out_(setup_({admin_user:p.admin_user, admin_pass:p.admin_pass}));
  if (p.action==='ping') return out_({ok:true, app:'Mercado Ciénega', time:new Date().toISOString()});
  return out_({ok:false, error:'Usa POST.'});
}

function doPost(e){
  try{
    var b = JSON.parse(e.postData.contents);
    var a = b.action, tok = b.token;
    switch(a){

      case 'ping': return out_({ok:true, app:'Mercado Ciénega'});

      /* ----- auth ----- */
      case 'register': {
        var un = String(b.username||'').trim().toLowerCase();
        var pw = String(b.password||'');
        if(!/^[a-z0-9_]{3,20}$/.test(un)) return out_({ok:false, error:'Usuario: 3-20 caracteres (letras, números, _).'});
        if(pw.length<4) return out_({ok:false, error:'La clave debe tener al menos 4 caracteres.'});
        if(userGet_(un)) return out_({ok:false, error:'Ese usuario ya existe.'});
        sh_(SH.USUARIOS).appendRow([un, hashPw_(pw), String(b.nombre||'').slice(0,60),
          String(b.telefono||'').slice(0,20), 'user', '']);
        var t0 = sessionCreate_(un);
        return out_({ok:true, token:t0, user:{username:un, nombre:String(b.nombre||''), rol:'user'}});
      }
      case 'login': {
        var u = userGet_(String(b.username||'').trim().toLowerCase());
        if(!u || !verifyPw_(String(b.password||''), u.hash))
          return out_({ok:false, error:'Usuario o clave incorrectos.'});
        var t = sessionCreate_(u.username);
        return out_({ok:true, token:t,
          user:{username:u.username, nombre:u.nombre, telefono:u.telefono, rol:u.rol, must_change:u.must_change}});
      }
      case 'logout': sessionDelete_(tok); return out_({ok:true});
      case 'me': {
        var m = authUser_(tok); if(m.err) return out_({ok:false, error:m.err});
        var uu = m.user;
        return out_({ok:true, user:{username:uu.username, nombre:uu.nombre, telefono:uu.telefono, rol:uu.rol, must_change:uu.must_change}});
      }
      case 'pass_change': {
        var pc = authUser_(tok); if(pc.err) return out_({ok:false, error:pc.err});
        if(!verifyPw_(String(b.actual||''), pc.user.hash)) return out_({ok:false, error:'La clave actual no es correcta.'});
        if(String(b.nueva||'').length<4) return out_({ok:false, error:'La nueva clave debe tener al menos 4 caracteres.'});
        var r = findRow_(SH.USUARIOS, 0, pc.user.username);
        sh_(SH.USUARIOS).getRange(r,2,1,1).setValue(hashPw_(String(b.nueva)));
        sh_(SH.USUARIOS).getRange(r,6,1,1).setValue('');
        return out_({ok:true});
      }

      /* ----- tasa ----- */
      case 'tasa_get': {
        var c = cfgGet_();
        return out_({ok:true, tasa:Number(c.tasa_usd)||770, fecha:String(c.tasa_fecha), fuente:String(c.tasa_fuente)});
      }
      case 'tasa_set': {
        var ta = needAdmin_(tok); if(ta.err) return out_({ok:false, error:ta.err});
        var tv = Number(b.tasa);
        if(!(tv>50 && tv<5000)) return out_({ok:false, error:'Tasa inválida.'});
        cfgSet_('tasa_usd', tv); cfgSet_('tasa_fecha', hoy_()); cfgSet_('tasa_fuente', 'Manual (admin)');
        return out_({ok:true});
      }
      case 'tasa_auto': return out_(tasaAuto_());

      /* ----- categorías ----- */
      case 'categorias_list':
        return out_({ok:true, categorias: rows_(SH.CATEGORIAS).map(function(v){return String(v[0]);})});
      case 'categoria_add': {
        var ca = needAdmin_(tok); if(ca.err) return out_({ok:false, error:ca.err});
        var cn = String(b.nombre||'').trim().slice(0,40);
        if(!cn) return out_({ok:false, error:'Nombre vacío.'});
        sh_(SH.CATEGORIAS).appendRow([cn]);
        return out_({ok:true});
      }
      case 'categoria_del': {
        var cd = needAdmin_(tok); if(cd.err) return out_({ok:false, error:cd.err});
        var dr = findRow_(SH.CATEGORIAS, 0, b.nombre);
        if(dr>0) sh_(SH.CATEGORIAS).deleteRow(dr);
        return out_({ok:true});
      }

      /* ----- productos ----- */
      case 'productos_list': {
        var list = rows_(SH.PRODUCTOS).map(prodRow_).filter(prodActivo_);
        if(b.categoria) list = list.filter(function(p){return p.categoria===b.categoria;});
        if(b.q){
          var q = String(b.q).toLowerCase();
          list = list.filter(function(p){ return (p.titulo+' '+p.descripcion).toLowerCase().indexOf(q)>=0; });
        }
        list.sort(function(x,y){
          var dx = prodEsDestacado_(x)?0:1, dy = prodEsDestacado_(y)?0:1;
          if(dx!==dy) return dx-dy;
          return y.fecha_pub < x.fecha_pub ? -1 : 1;
        });
        return out_({ok:true, productos:list.map(prodPub_)});
      }
      case 'producto_get': {
        var gr = findRow_(SH.PRODUCTOS, 0, b.id);
        if(gr<0) return out_({ok:false, error:'No existe.'});
        var gp = prodRow_(sh_(SH.PRODUCTOS).getRange(gr,1,1,16).getValues()[0]);
        if(!prodActivo_(gp)) return out_({ok:false, error:'Anuncio no disponible.'});
        return out_({ok:true, producto:prodPub_(gp)});
      }
      case 'producto_create': {
        var cu = authUser_(tok); if(cu.err) return out_({ok:false, error:cu.err});
        var tit = String(b.titulo||'').trim().slice(0,80);
        var pr = Number(b.precio);
        if(!tit) return out_({ok:false, error:'Falta el título.'});
        if(!(pr>0)) return out_({ok:false, error:'Precio inválido.'});
        var met0 = String(b.metodo||'');
        if(['transfermovil','usdt','efectivo'].indexOf(met0)<0)
          return out_({ok:false, error:'Elige un método de pago.'});
        var ids = guardarImagenes_(b.imagenes);
        var cfg0 = cfgGet_();
        var pid0 = newId_('P');
        // el anuncio nace pendiente: solo se muestra cuando el admin confirma el pago
        var row = [pid0, tit, String(b.categoria||'Otros').slice(0,40), pr,
          String(b.moneda||'CUP')==='USD'?'USD':'CUP', String(b.descripcion||'').slice(0,2000),
          ids.join(','), cu.user.username, String(b.telefono||cu.user.telefono||'').slice(0,20),
          String(b.whatsapp||'')==='1'?'1':'', String(b.municipio||'').slice(0,40),
          '', '', '', '', 'pendiente'];
        sh_(SH.PRODUCTOS).appendRow(row);
        sh_(SH.PAGOS).appendRow([newId_('PG'), hoy_(), pid0, tit, cu.user.username,
          met0, String(b.referencia||'').slice(0,60), Number(cfg0.publicar_precio_cup)||500,
          'pendiente', 'publicar']);
        return out_({ok:true, id:pid0, imagenes:ids.length,
          msg:'Anuncio recibido. Se publicará cuando el administrador confirme tu pago.'});
      }
      case 'producto_update': {
        var uu2 = authUser_(tok); if(uu2.err) return out_({ok:false, error:uu2.err});
        var ur = findRow_(SH.PRODUCTOS, 0, b.id);
        if(ur<0) return out_({ok:false, error:'No existe.'});
        var up = prodRow_(sh_(SH.PRODUCTOS).getRange(ur,1,1,16).getValues()[0]);
        if(up.vendedor!==uu2.user.username && uu2.user.rol!=='admin')
          return out_({ok:false, error:'No es tu anuncio.'});
        var s2 = sh_(SH.PRODUCTOS);
        if(b.titulo) s2.getRange(ur,2,1,1).setValue(String(b.titulo).slice(0,80));
        if(b.categoria) s2.getRange(ur,3,1,1).setValue(String(b.categoria).slice(0,40));
        if(b.precio && Number(b.precio)>0) s2.getRange(ur,4,1,1).setValue(Number(b.precio));
        if(b.moneda) s2.getRange(ur,5,1,1).setValue(String(b.moneda)==='USD'?'USD':'CUP');
        if(b.descripcion!==undefined) s2.getRange(ur,6,1,1).setValue(String(b.descripcion).slice(0,2000));
        if(b.telefono!==undefined) s2.getRange(ur,9,1,1).setValue(String(b.telefono).slice(0,20));
        if(b.municipio!==undefined) s2.getRange(ur,11,1,1).setValue(String(b.municipio).slice(0,40));
        if(b.imagenes && b.imagenes.length){
          var nids = guardarImagenes_(b.imagenes);
          if(nids.length) s2.getRange(ur,7,1,1).setValue(nids.join(','));
        }
        return out_({ok:true});
      }
      case 'producto_delete': {
        var du = authUser_(tok); if(du.err) return out_({ok:false, error:du.err});
        var dlr = findRow_(SH.PRODUCTOS, 0, b.id);
        if(dlr<0) return out_({ok:false, error:'No existe.'});
        var dp = prodRow_(sh_(SH.PRODUCTOS).getRange(dlr,1,1,16).getValues()[0]);
        if(dp.vendedor!==du.user.username && du.user.rol!=='admin')
          return out_({ok:false, error:'No es tu anuncio.'});
        sh_(SH.PRODUCTOS).getRange(dlr,16,1,1).setValue('inactivo');
        return out_({ok:true});
      }
      case 'mis_productos': {
        var mu = authUser_(tok); if(mu.err) return out_({ok:false, error:mu.err});
        var mine = rows_(SH.PRODUCTOS).map(prodRow_)
          .filter(function(p){return p.vendedor===mu.user.username && p.estado!=='inactivo';})
          .map(function(p){ var q=prodPub_(p); q.fecha_exp=p.fecha_exp; q.activo=prodActivo_(p); q.estado=p.estado; return q; });
        return out_({ok:true, productos:mine});
      }
      case 'producto_renovar': {
        var ru = authUser_(tok); if(ru.err) return out_({ok:false, error:ru.err});
        var rr = findRow_(SH.PRODUCTOS, 0, b.id);
        if(rr<0) return out_({ok:false, error:'No existe.'});
        var rp = prodRow_(sh_(SH.PRODUCTOS).getRange(rr,1,1,16).getValues()[0]);
        if(rp.vendedor!==ru.user.username && ru.user.rol!=='admin')
          return out_({ok:false, error:'No es tu anuncio.'});
        var met3 = String(b.metodo||'');
        if(['transfermovil','usdt','efectivo'].indexOf(met3)<0)
          return out_({ok:false, error:'Elige un método de pago.'});
        var cfg3 = cfgGet_();
        sh_(SH.PAGOS).appendRow([newId_('PG'), hoy_(), rp.id, rp.titulo, ru.user.username,
          met3, String(b.referencia||'').slice(0,60), Number(cfg3.publicar_precio_cup)||100,
          'pendiente', 'renovar']);
        return out_({ok:true, msg:'Solicitud de renovación enviada. Se extenderá 1 mes cuando se apruebe el pago.'});
      }

      /* ----- destacar (pagado) ----- */
      case 'destacar_info': {
        var cfi = cfgGet_();
        return out_({ok:true,
          precio_cup:Number(cfi.destacado_precio_cup)||500,
          publicar_precio_cup:Number(cfi.publicar_precio_cup)||500,
          dias:Number(cfi.destacado_dias)||30,
          metodos:{
            transfermovil:String(cfi.pago_transfermovil_numero||''),
            usdt:String(cfi.pago_usdt_direccion||''),
            efectivo:String(cfi.pago_efectivo_nota||'')
          }});
      }
      case 'destacar_solicitar': {
        var su = authUser_(tok); if(su.err) return out_({ok:false, error:su.err});
        var sr = findRow_(SH.PRODUCTOS, 0, b.producto_id);
        if(sr<0) return out_({ok:false, error:'No existe.'});
        var sp = prodRow_(sh_(SH.PRODUCTOS).getRange(sr,1,1,16).getValues()[0]);
        if(sp.vendedor!==su.user.username) return out_({ok:false, error:'No es tu anuncio.'});
        var met = String(b.metodo||'');
        if(['transfermovil','usdt','efectivo'].indexOf(met)<0)
          return out_({ok:false, error:'Método de pago inválido.'});
        var cfd = cfgGet_();
        sh_(SH.PAGOS).appendRow([newId_('PG'), hoy_(), sp.id, sp.titulo, su.user.username,
          met, String(b.referencia||'').slice(0,60), Number(cfd.destacado_precio_cup)||500, 'pendiente', 'destacar']);
        return out_({ok:true, msg:'Solicitud enviada. Tu anuncio se destacará cuando se apruebe el pago.'});
      }
      case 'pagos_list': {
        var pl = needAdmin_(tok); if(pl.err) return out_({ok:false, error:pl.err});
        var pays = rows_(SH.PAGOS).map(function(v){
          return {id:String(v[0]), fecha:String(v[1]), producto_id:String(v[2]), titulo:String(v[3]),
                  usuario:String(v[4]), metodo:String(v[5]), referencia:String(v[6]),
                  monto_cup:Number(v[7])||0, estado:String(v[8]), tipo:String(v[9]||'destacar')};
        });
        pays.sort(function(a,b){ return b.fecha<a.fecha?-1:1; });
        return out_({ok:true, pagos:pays});
      }
      case 'pago_aprobar': {
        var pa = needAdmin_(tok); if(pa.err) return out_({ok:false, error:pa.err});
        var pr2 = findRow_(SH.PAGOS, 0, b.id);
        if(pr2<0) return out_({ok:false, error:'No existe.'});
        var pv = sh_(SH.PAGOS).getRange(pr2,1,1,10).getValues()[0];
        if(String(pv[8])!=='pendiente') return out_({ok:false, error:'Ya fue procesado.'});
        sh_(SH.PAGOS).getRange(pr2,9,1,1).setValue('aprobado');
        var cfa = cfgGet_();
        var fr = findRow_(SH.PRODUCTOS, 0, String(pv[2]));
        var tipo = String(pv[9]||'destacar');
        if(fr>0){
          if(tipo==='publicar'){
            // al confirmar el pago el anuncio sale publicado por 1 mes
            sh_(SH.PRODUCTOS).getRange(fr,12,1,1).setValue(hoy_());
            sh_(SH.PRODUCTOS).getRange(fr,13,1,1).setValue(sumarMeses_(hoy_(), Number(cfa.anuncio_meses)||1));
            sh_(SH.PRODUCTOS).getRange(fr,16,1,1).setValue('activo');
          }else if(tipo==='renovar'){
            // extender 1 mes desde el vencimiento actual (o desde hoy si ya venció)
            var expV = sh_(SH.PRODUCTOS).getRange(fr,13,1,1).getValue();
            var expS = '';
            try{ expS = Utilities.formatDate(new Date(expV), 'GMT-4', 'yyyy-MM-dd'); }catch(e){}
            var base = (expS && expS>=hoy_()) ? expS : hoy_();
            sh_(SH.PRODUCTOS).getRange(fr,13,1,1).setValue(sumarMeses_(base, Number(cfa.anuncio_meses)||1));
            sh_(SH.PRODUCTOS).getRange(fr,16,1,1).setValue('activo');
          }else{
            sh_(SH.PRODUCTOS).getRange(fr,14,1,1).setValue('1');
            sh_(SH.PRODUCTOS).getRange(fr,15,1,1).setValue(sumarDias_(hoy_(), Number(cfa.destacado_dias)||30));
          }
        }
        return out_({ok:true});
      }
      case 'pago_rechazar': {
        var prj = needAdmin_(tok); if(prj.err) return out_({ok:false, error:prj.err});
        var rj = findRow_(SH.PAGOS, 0, b.id);
        if(rj<0) return out_({ok:false, error:'No existe.'});
        sh_(SH.PAGOS).getRange(rj,9,1,1).setValue('rechazado');
        return out_({ok:true});
      }

      /* ----- moderador automático (revisión diaria de categorías) ----- */
      case 'moderador_recategorizar': {
        var cfgm = cfgGet_();
        var realTok = String(cfgm.moderador_token||'');
        var dadoTok = String(b.mod_token||'');
        if(!dadoTok && e && e.parameter) dadoTok = String(e.parameter.mod_token||'');
        if(!realTok || dadoTok!==realTok) return out_({ok:false, error:'No autorizado.'});
        var fr = findRow_(SH.PRODUCTOS, 0, b.id);
        if(fr<0) return out_({ok:false, error:'No existe.'});
        var catN = String(b.categoria||'');
        if(CATS.indexOf(catN)<0) return out_({ok:false, error:'Categoría inválida.'});
        var oldC = String(sh_(SH.PRODUCTOS).getRange(fr,3,1,1).getValue()||'');
        if(oldC===catN) return out_({ok:true, sin_cambio:true, categoria:oldC});
        sh_(SH.PRODUCTOS).getRange(fr,3,1,1).setValue(catN);
        return out_({ok:true, anterior:oldC, nueva:catN});
      }

      /* ----- pedidos (carrito) ----- */
      case 'pedido_create': {
        var ou = authUser_(tok); if(ou.err) return out_({ok:false, error:ou.err});
        var items = b.items||[];
        if(!items.length) return out_({ok:false, error:'Carrito vacío.'});
        var cfgp = cfgGet_(), tasa = Number(cfgp.tasa_usd)||770;
        var det = [], tCup = 0;
        for (var i=0;i<items.length;i++){
          var ir = findRow_(SH.PRODUCTOS, 0, items[i].id);
          if(ir<0) continue;
          var ip = prodRow_(sh_(SH.PRODUCTOS).getRange(ir,1,1,16).getValues()[0]);
          if(!prodActivo_(ip)) continue;
          var cant = Math.max(1, Math.min(99, Number(items[i].cantidad)||1));
          var enCup = ip.moneda==='USD' ? ip.precio*tasa : ip.precio;
          tCup += enCup*cant;
          det.push({id:ip.id, titulo:ip.titulo, precio:ip.precio, moneda:ip.moneda,
                    cantidad:cant, vendedor:ip.vendedor, telefono:ip.telefono});
        }
        if(!det.length) return out_({ok:false, error:'Ningún producto disponible.'});
        var pid = newId_('PD');
        sh_(SH.PEDIDOS).appendRow([pid, hoy_(), ou.user.username, ou.user.telefono,
          JSON.stringify(det), Math.round(tCup), Math.round(tCup/tasa*100)/100,
          String(b.nota||'').slice(0,300), 'nuevo']);
        return out_({ok:true, pedido:pid, items:det});
      }
      case 'pedidos_mios': {
        var pm = authUser_(tok); if(pm.err) return out_({ok:false, error:pm.err});
        var mine2 = rows_(SH.PEDIDOS)
          .filter(function(v){return String(v[2])===pm.user.username;})
          .map(function(v){ return {id:String(v[0]), fecha:String(v[1]), items:JSON.parse(String(v[4]||'[]')),
            total_cup:Number(v[5])||0, total_usd:Number(v[6])||0, estado:String(v[8])}; });
        return out_({ok:true, pedidos:mine2.reverse()});
      }

      /* ----- admin: config ----- */
      case 'config_get': {
        var cg = needAdmin_(tok); if(cg.err) return out_({ok:false, error:cg.err});
        return out_({ok:true, config:cfgGet_()});
      }
      case 'config_set': {
        var csu = needAdmin_(tok); if(csu.err) return out_({ok:false, error:csu.err});
        var d0 = cfgDefaults_(), k0 = String(b.clave||'');
        if(!(k0 in d0)) return out_({ok:false, error:'Clave desconocida.'});
        cfgSet_(k0, b.valor);
        return out_({ok:true});
      }
      case 'productos_admin': {
        var adm = needAdmin_(tok); if(adm.err) return out_({ok:false, error:adm.err});
        var all = rows_(SH.PRODUCTOS).map(function(v){ var p=prodRow_(v); var q=prodPub_(p);
          q.fecha_exp=p.fecha_exp; q.activo=prodActivo_(p); q.estado=p.estado; return q; });
        return out_({ok:true, productos:all.reverse()});
      }

      default: return out_({ok:false, error:'Acción desconocida.'});
    }
  }catch(err){
    return out_({ok:false, error:'Error: '+err.message});
  }
}
