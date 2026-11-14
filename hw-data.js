/* Datos de Hardware (compartido por las vistas «Hardware» y «Mapa de activos»).
   Lee el Excel «HARWARE ACTUALIZADO …» de la carpeta HW (OneDrive de acastro) vía Microsoft Graph, hoja BBDD, y calcula resúmenes y comparaciones entre cortes. */
window.HW = (function(){
"use strict";
var CFG = {
  clientId: "2dd44541-d3ac-459c-afa0-aa873590f7e0",
  tenant:   "9d85a34e-57e5-4f82-9341-7dd9300077b7",
  scopes:   ["Files.Read.All"],
  driveId:  "b!ftpx1PHMU02Xyd_BvRi5gZ4jKtzFMjZHt1vZTICVWd0sPgcGYMRfTZopqpx_5vTB",
  folderId: "015HP6ULJFS5ISEVNR5JC2CPNZJLQQB4NO",   // PLAN / PROYECTOS ACTUALES / HW  (el archivo se renombra en cada corte: se toma el más reciente)
  sheet:    "BBDD",
  /* Excel histórico de cortes (lo escribe bat-merch/tools/guardar_corte_hw.py). Está en el OneDrive de Claudia: hay que compartirlo con quien use la vista. */
  cortes:   { driveId: "b!NpQ7a93VQkOL-yt8XpCvELSu8zBzw0hChMPffAg5mBs_yxFrMcD7QKFa-_Ao0MIL", itemId: "01FD7R7ECU6CYHD6WMPJDZVQEFCEPJZ5IO" },
  maxQty:   10                                       // cantidades mayores se consideran error de digitación (se cuenta 1 y se avisa)
};
var G = "https://graph.microsoft.com/v1.0/drives/"+CFG.driveId;

function norm(s){ return String(s==null?"":s).normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/\s+/g," ").trim().toUpperCase(); }
function esc(s){ return String(s==null?"":s).replace(/[&<>"]/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c];}); }
function num(v){ var n=parseFloat(String(v==null?"":v).replace(",",".")); return isNaN(n)?0:n; }
var vacio = function(v){ var n=norm(v); return n===""||n==="N/A"||n==="NA"||n==="-"||n==="SIN ACTIVO BAT"||n==="SIN ACTIVO"||n==="NO"; };

/* ═══ Coordenadas ═══ */
var CITY = { "LIMA":[-12.046,-77.043],"AREQUIPA":[-16.409,-71.537],"TRUJILLO":[-8.112,-79.029],"PIURA":[-5.194,-80.633],"CHICLAYO":[-6.771,-79.841],"CUSCO":[-13.532,-71.967],
  "ICA":[-14.068,-75.729],"HUANCAYO":[-12.065,-75.205],"CHIMBOTE":[-9.085,-78.578],"NUEVO CHIMBOTE":[-9.12,-78.52],"CHINCHA":[-13.41,-76.132],"IQUITOS":[-3.749,-73.253],"LORETO":[-3.749,-73.253],
  "HUACHO":[-11.107,-77.605],"TARAPOTO":[-6.488,-76.36],"TACNA":[-18.014,-70.253],"HUANUCO":[-9.93,-76.242],"PUCALLPA":[-8.379,-74.554],"UCAYALI":[-8.379,-74.554],"TUMBES":[-3.567,-80.451],
  "BARRANCA":[-10.752,-77.761],"CHANCAY":[-11.571,-77.27],"HUARAL":[-11.495,-77.208],"HUAURA":[-11.07,-77.6],"CASMA":[-9.474,-78.305],"AYACUCHO":[-13.158,-74.223],"ILO":[-17.639,-71.337],
  "CAJAMARCA":[-7.157,-78.5],"HUARMEY":[-10.068,-78.15],"HUARAZ":[-9.526,-77.528],"LA MERCED":[-11.055,-75.326],"CHURIN":[-10.8,-76.9],"JAEN":[-5.708,-78.807],"OXAPAMPA":[-10.573,-75.397] };
/* Acepta número, "-12,1044" o "-12.125.776.230" (puntos como separador de miles) y exige que caiga dentro de Perú. */
function coord(v, isLat){
  var n;
  if(typeof v==="number") n=v;
  else { var s=String(v==null?"":v).trim().replace(/\s/g,""); if(!s) return null;
    var dots=(s.match(/\./g)||[]).length; if(dots>1){ var i=s.indexOf("."); s=s.slice(0,i+1)+s.slice(i+1).replace(/\./g,""); }
    n=parseFloat(s.replace(",",".")); }
  if(isNaN(n)) return null;
  return isLat ? (n>=-18.6&&n<=0.2?n:null) : (n>=-81.5&&n<=-68.4?n:null);
}

/* ═══ Lectura del libro ═══ */
var COLS = { id:"ID STORE", zona:"LIMA/KC/NKC", canal:"CANAL", ciudad:"CIUDAD", comCig:"COMUNICACION DE CIGARRERA", cadena:"CADENA", pos:"POS", status:"STATUS", tipoCig:"TIPO DE CIG",
  tvQ:"CANTIDAD PANTALLA", cigA:"DETALLE DE CIGARRERA - A", cigAQ:"CANTIDAD CIG - A", cigB:"DETALLE DE CIGARRERA - B", cigBQ:"CANTIDAD CIG - B", cigC:"DETALLE DE CIGARRERA - C", cigCQ:"CANTIDAD CIG - C",
  comDisp:"COMUNICACION DISPENSER", dispA:"DISPENSER A", dispB:"DISPENSER B", dispQ:"CANTIDAD DISP.", gFmc:"G. FMC", gVuse:"G. VUSE", gVelo:"G. VELO", gVuseQ:"CANT. GLORI VUSE",
  vol:"VOLUMETRICO", volQ:"CANTIDAD VOLUMETRICOS", espejo:"ESPEJO", comodato:"COMODATO", viajero:"POS VIAJERO", distrito:"DISTRITO", lat:"LATITUD", lon:"LONGITUD" };

function assetsOf(r, warn){
  var out=[], q=function(v){ var n=num(v); if(n>CFG.maxQty){ warn.n++; return 1; } return n; };
  [["cigA","cigAQ"],["cigB","cigBQ"],["cigC","cigCQ"]].forEach(function(p){ if(!vacio(r[p[0]])){ var n=q(r[p[1]])||1; out.push({cat:"Cigarrera",act:norm(r[p[0]]),n:n}); } });
  var hasA=!vacio(r.dispA), hasB=!vacio(r.dispB), dq=q(r.dispQ);
  if(hasA||hasB){ var nB=hasB?1:0, nA=hasA?Math.max((dq||(nB+1))-nB,1):0;
    if(hasA) out.push({cat:"Dispenser",act:norm(r.dispA),n:nA}); if(hasB) out.push({cat:"Dispenser",act:norm(r.dispB),n:nB}); }
  if(!vacio(r.gFmc))  out.push({cat:"Glorificador",act:"FMC · "+norm(r.gFmc),n:1});
  if(!vacio(r.gVuse)) out.push({cat:"Glorificador",act:"VUSE · "+norm(r.gVuse),n:q(r.gVuseQ)||1});
  if(!vacio(r.gVelo)) out.push({cat:"Glorificador",act:"VELO · "+norm(r.gVelo),n:1});
  if(!vacio(r.vol)){ out.push({cat:"Volumétrico",act:"VOLUMÉTRICO",n:q(r.volQ)||1}); }
  var tv=q(r.tvQ); if(tv>0) out.push({cat:"TV",act:"TV",n:tv});
  if(norm(r.espejo)==="SI") out.push({cat:"Espejo",act:"ESPEJO",n:1});
  return out;
}
function findHeader(rows){
  for(var i=0;i<Math.min(rows.length,8);i++){ var h=rows[i].map(norm); if(h.indexOf("ID STORE")>=0 && h.indexOf("CIUDAD")>=0 && h.indexOf("CADENA")>=0) return {hi:i,h:h}; }
  return null;
}
/* Devuelve {rows, meta:{filas, avisos}} desde un ArrayBuffer. Solo lee la hoja BBDD (o la hoja con esos encabezados). */
function parseBook(buf){
  var cheap=XLSX.read(buf,{type:"array",bookSheets:true}), name=cheap.SheetNames.indexOf(CFG.sheet)>=0?CFG.sheet:null, ws, wb;
  if(name){ wb=XLSX.read(buf,{type:"array",sheets:[name]}); ws=wb.Sheets[name]; }
  else { wb=XLSX.read(buf,{type:"array"}); for(var i=0;i<wb.SheetNames.length&&!ws;i++){ var w=wb.Sheets[wb.SheetNames[i]], rr=XLSX.utils.sheet_to_json(w,{header:1,range:0,defval:""}).slice(0,8); if(findHeader(rr)) ws=w; } }
  if(!ws) throw new Error("No encontré la hoja BBDD (ID STORE / CIUDAD / CADENA) en el archivo.");
  var rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:""}), H=findHeader(rows);
  if(!H) throw new Error("No encontré los encabezados de la hoja BBDD.");
  var ix={}; Object.keys(COLS).forEach(function(k){ ix[k]=H.h.indexOf(COLS[k]); });
  var out=[], warn={n:0}, seen={}, dup=0;
  for(var r=H.hi+1;r<rows.length;r++){
    var row=rows[r], g=function(k){ return ix[k]<0?"":row[ix[k]]; }, id=String(g("id")).trim(); if(!id) continue;
    if(seen[id]){ dup++; continue; } seen[id]=1;
    var o={}; Object.keys(COLS).forEach(function(k){ o[k]=g(k); });
    o.id=id; o.ciudad=norm(o.ciudad)||"SIN CIUDAD"; o.cadena=norm(o.cadena)||"SIN CADENA"; o.zona=norm(o.zona)||"SIN ZONA"; o.canal=norm(o.canal)||"SIN CANAL"; o.pos=String(o.pos).trim();
    o.comCig=norm(o.comCig)||"N/A"; o.comDisp=norm(o.comDisp)||"N/A"; o.distrito=norm(o.distrito);
    var la=coord(o.lat,true), lo=coord(o.lon,false); o.geo = la!=null&&lo!=null ? [la,lo] : null; o.geoExacta=!!o.geo; if(!o.geo && CITY[o.ciudad]) o.geo=CITY[o.ciudad].slice();
    o.assets=assetsOf(o,warn);
    o.u={cig:0,disp:0,glor:0,vol:0,tv:0,esp:0,total:0};
    o.assets.forEach(function(a){ var k={Cigarrera:"cig",Dispenser:"disp",Glorificador:"glor","Volumétrico":"vol",TV:"tv",Espejo:"esp"}[a.cat]; o.u[k]+=a.n; o.u.total+=a.n; });
    out.push(o);
  }
  return {rows:out, meta:{filas:out.length, duplicados:dup, cantidadesRaras:warn.n}};
}

/* ═══ Resúmenes ═══ */
var DIMS = { ciudad:"Ciudad", cadena:"Cadena", zona:"Zona (Lima/KC/NKC)", canal:"Canal", comCig:"Comunicación cigarrera", comDisp:"Comunicación dispenser" };
function byDim(rows, key){
  var m={}; rows.forEach(function(r){ var k=r[key]; var o=m[k]||(m[k]={k:k,pdv:0,cig:0,disp:0,glor:0,vol:0,tv:0,esp:0,total:0}); o.pdv++; ["cig","disp","glor","vol","tv","esp","total"].forEach(function(f){ o[f]+=r.u[f]; }); });
  return Object.keys(m).map(function(k){return m[k];}).sort(function(a,b){ return b.pdv-a.pdv || (a.k>b.k?1:-1); });
}
function byActivo(rows){
  var m={}; rows.forEach(function(r){ r.assets.forEach(function(a){ var k=a.cat+"|"+a.act, o=m[k]||(m[k]={cat:a.cat,act:a.act,pdv:0,units:0}); o.pdv++; o.units+=a.n; }); });
  return Object.keys(m).map(function(k){return m[k];}).sort(function(a,b){ return a.cat===b.cat ? b.units-a.units : (a.cat>b.cat?1:-1); });
}
function totals(rows){ var t={pdv:rows.length,cig:0,disp:0,glor:0,vol:0,tv:0,esp:0,total:0,ciudades:{},geoExacta:0,comodato:0,viajero:0};
  rows.forEach(function(r){ ["cig","disp","glor","vol","tv","esp","total"].forEach(function(f){ t[f]+=r.u[f]; }); t.ciudades[r.ciudad]=1; if(r.geoExacta) t.geoExacta++; if(norm(r.comodato)==="SI") t.comodato++; if(norm(r.viajero)==="SI") t.viajero++; });
  t.ciudades=Object.keys(t.ciudades).length; return t; }

/* ═══ Cortes: un «snapshot» es el resumen de un corte (totales, indicadores, activos y una fila por PDV con su huella de hardware) ═══ */
function hash(str){ var h=5381; for(var i=0;i<str.length;i++) h=((h<<5)+h+str.charCodeAt(i))|0; return (h>>>0).toString(36); }   // misma huella que tools/guardar_corte_hw.py
function sig(r){ return r.assets.map(function(a){return a.cat+":"+a.act+"x"+a.n;}).sort().join("|"); }
function snapshot(rows, meta){
  var dict={ciudad:[],cadena:[],zona:[]}, idx={ciudad:{},cadena:{},zona:{}}, pos={};
  var di=function(k,v){ var m=idx[k]; if(m[v]==null){ m[v]=dict[k].length; dict[k].push(v); } return m[v]; };
  var pdv=rows.map(function(r){ pos[r.id]=r.pos; return [r.id,di("ciudad",r.ciudad),di("cadena",r.cadena),di("zona",r.zona),hash(sig(r)),r.u.total]; });
  return { corte:(meta&&meta.corte)||"", archivo:(meta&&meta.archivo)||"", tot:totals(rows), dims:{ciudad:byDim(rows,"ciudad"),cadena:byDim(rows,"cadena"),zona:byDim(rows,"zona")}, activos:byActivo(rows), dict:dict, pdv:pdv, pos:pos };
}
/* Compara dos snapshots: A (anterior) vs B (actual) */
function compareSnaps(A, B){
  var ma={}, mb={}; A.pdv.forEach(function(p){ ma[p[0]]=p; }); B.pdv.forEach(function(p){ mb[p[0]]=p; });
  var info=function(S,p){ return {id:p[0],pos:(B.pos&&B.pos[p[0]])||(A.pos&&A.pos[p[0]])||"",ciudad:S.dict.ciudad[p[1]],cadena:S.dict.cadena[p[2]],zona:S.dict.zona[p[3]],total:p[5]}; };
  var nuevos=B.pdv.filter(function(p){return !ma[p[0]];}).map(function(p){return info(B,p);});
  var salieron=A.pdv.filter(function(p){return !mb[p[0]];}).map(function(p){return info(A,p);});
  var cambios=B.pdv.filter(function(p){ return ma[p[0]] && ma[p[0]][4]!==p[4]; }).map(function(p){ var o=info(B,p); o.antes=ma[p[0]][5]; o.ahora=p[5]; return o; });
  function delta(key){ var da={}, db={}; (A.dims[key]||[]).forEach(function(o){da[o.k]=o;}); (B.dims[key]||[]).forEach(function(o){db[o.k]=o;});
    var ks={}; Object.keys(da).concat(Object.keys(db)).forEach(function(k){ks[k]=1;});
    return Object.keys(ks).map(function(k){ var a=da[k]||{pdv:0,total:0}, b=db[k]||{pdv:0,total:0}; return {k:k,pdvA:a.pdv,pdvB:b.pdv,pdvD:b.pdv-a.pdv,totA:a.total,totB:b.total,totD:b.total-a.total}; })
      .sort(function(x,y){ return Math.abs(y.totD)+Math.abs(y.pdvD)-Math.abs(x.totD)-Math.abs(x.pdvD); }); }
  var aa={}, ab={}; A.activos.forEach(function(o){aa[o.cat+"|"+o.act]=o;}); B.activos.forEach(function(o){ab[o.cat+"|"+o.act]=o;});
  var ks={}; Object.keys(aa).concat(Object.keys(ab)).forEach(function(k){ks[k]=1;});
  var activos=Object.keys(ks).map(function(k){ var a=aa[k]||{units:0,pdv:0}, b=ab[k]||{units:0,pdv:0}, p=k.split("|"); return {cat:p[0],act:p[1],uA:a.units,uB:b.units,uD:b.units-a.units,pA:a.pdv,pB:b.pdv}; })
    .filter(function(x){return x.uD!==0||x.pA!==x.pB;}).sort(function(x,y){ return Math.abs(y.uD)-Math.abs(x.uD); });
  return {totA:A.tot,totB:B.tot,nuevos:nuevos,salieron:salieron,cambios:cambios,ciudad:delta("ciudad"),cadena:delta("cadena"),zona:delta("zona"),activos:activos};
}

/* ═══ Microsoft Graph + MSAL ═══ */
var msalApp=null;
function initAuth(onAccount){
  if(typeof msal==="undefined") return Promise.resolve(false);
  msalApp=new msal.PublicClientApplication({auth:{clientId:CFG.clientId,authority:"https://login.microsoftonline.com/"+CFG.tenant,redirectUri:location.origin+location.pathname},cache:{cacheLocation:"localStorage"}});
  return msalApp.initialize().then(function(){return msalApp.handleRedirectPromise();}).then(function(resp){
    if(resp&&resp.account) msalApp.setActiveAccount(resp.account);
    var acc=msalApp.getActiveAccount()||msalApp.getAllAccounts()[0];
    if(acc){ msalApp.setActiveAccount(acc); onAccount&&onAccount(acc.username); return true; } return false;
  }).catch(function(e){ console.warn("MSAL init:",e&&e.message); return false; });
}
function token(interactive){
  var req={scopes:CFG.scopes}, acc=msalApp.getAllAccounts()[0];
  if(acc){ req.account=acc; return msalApp.acquireTokenSilent(req).then(function(r){return r.accessToken;}).catch(function(){
    if(!interactive) return Promise.reject(new Error("silent-failed")); return msalApp.acquireTokenPopup(req).then(function(r){return r.accessToken;}); }); }
  if(!interactive) return Promise.reject(new Error("no-account"));
  return msalApp.loginPopup(req).then(function(r){ return r.accessToken; });
}
function gfetch(url,t){ return fetch(url,{headers:{Authorization:"Bearer "+t}}).then(function(r){ if(!r.ok) throw new Error("Graph HTTP "+r.status); return r; }); }
/* El archivo vigente: el .xlsx más reciente de la carpeta HW cuyo nombre contenga HARWARE/HARDWARE */
function findFile(t){
  return gfetch(G+"/items/"+CFG.folderId+"/children?$select=id,name,size,file,lastModifiedDateTime&$top=200",t).then(function(r){return r.json();}).then(function(j){
    var f=(j.value||[]).filter(function(i){ return i.file && /\.xls[xm]$/i.test(i.name) && i.name.indexOf("~$")!==0 && /HAR?D?WARE/i.test(i.name); })
      .sort(function(a,b){ return a.lastModifiedDateTime<b.lastModifiedDateTime?1:-1; })[0];
    if(!f) throw new Error("No encontré el Excel de Hardware en la carpeta HW.");
    return f; });
}
function download(t,item,versionId){
  var u=G+"/items/"+item.id+(versionId?"/versions/"+encodeURIComponent(versionId):"")+"/content";
  return gfetch(u,t).then(function(r){return r.arrayBuffer();});
}
function versions(t,item){
  return gfetch(G+"/items/"+item.id+"/versions",t).then(function(r){return r.json();}).then(function(j){
    return (j.value||[]).map(function(v){ return {id:v.id,fecha:v.lastModifiedDateTime,size:v.size,por:(v.lastModifiedBy&&v.lastModifiedBy.user&&v.lastModifiedBy.user.displayName)||""}; })
      .sort(function(a,b){ return a.fecha<b.fecha?1:-1; }); });
}
/* Excel histórico de cortes: hojas Cortes / Resumen / Activos / PDV (ver tools/guardar_corte_hw.py). Devuelve {lista, get(id)→snapshot}. */
function cortesLoad(t){
  var c=CFG.cortes; if(!c||!c.itemId) return Promise.reject(new Error("No hay Excel de cortes configurado."));
  return gfetch("https://graph.microsoft.com/v1.0/drives/"+c.driveId+"/items/"+c.itemId+"/content",t).then(function(r){return r.arrayBuffer();}).then(function(buf){
    var wb=XLSX.read(buf,{type:"array"});
    var rows=function(n){ return wb.Sheets[n] ? XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true,defval:""}).slice(1).filter(function(r){return r[0]!=="";}) : []; };
    var cortes=rows("Cortes").map(function(r){ return {id:String(r[0]),archivo:String(r[1]),pdv:+r[2]||0,activos:+r[9]||0,tot:{pdv:+r[2]||0,cig:+r[3]||0,disp:+r[4]||0,glor:+r[5]||0,vol:+r[6]||0,tv:+r[7]||0,esp:+r[8]||0,total:+r[9]||0,ciudades:0},guardado:String(r[11]||"")}; })
      .sort(function(a,b){ return a.id<b.id?1:-1; });
    var resumen=rows("Resumen"), activos=rows("Activos"), pdv=rows("PDV");
    function get(id){
      var cut=cortes.filter(function(x){return x.id===id;})[0]; if(!cut) return Promise.reject(new Error("Corte no encontrado en el Excel de cortes."));
      var dims={ciudad:[],cadena:[],zona:[]};
      resumen.forEach(function(r){ if(String(r[0])===id && dims[r[1]]) dims[r[1]].push({k:String(r[2]),pdv:+r[3],cig:+r[4],disp:+r[5],glor:+r[6],vol:+r[7],tv:+r[8],esp:+r[9],total:+r[10]}); });
      cut.tot.ciudades=dims.ciudad.length;
      var act=[]; activos.forEach(function(r){ if(String(r[0])===id) act.push({cat:String(r[1]),act:String(r[2]),pdv:+r[3],units:+r[4]}); });
      var dict={ciudad:[],cadena:[],zona:[]}, idx={ciudad:{},cadena:{},zona:{}}, pos={}, list=[];
      var di=function(k,v){ var m=idx[k]; if(m[v]==null){ m[v]=dict[k].length; dict[k].push(v); } return m[v]; };
      pdv.forEach(function(r){ if(String(r[0])!==id) return; var pid=String(r[1]); pos[pid]=String(r[2]); list.push([pid,di("ciudad",String(r[5])),di("cadena",String(r[6])),di("zona",String(r[3])),String(r[8]),+r[7]]); });
      return Promise.resolve({corte:id,archivo:cut.archivo,tot:cut.tot,dims:dims,activos:act,dict:dict,pdv:list,pos:pos});
    }
    return {lista:cortes,get:get};
  });
}
function errorMsg(err){
  var m=(err&&err.message)||String(err);
  if(m.indexOf("403")>=0) return "Tu cuenta no tiene acceso a la carpeta HW (403). Pide a acastro que te comparta PLAN/PROYECTOS ACTUALES/HW, o usa la carga manual.";
  if(m.indexOf("404")>=0) return "No se encontró la carpeta o el archivo (404): lo movieron o borraron. Usa la carga manual.";
  if(m.indexOf("interaction_in_progress")>=0) return "Hay un inicio de sesión en curso. Recarga (F5) y completa el login.";
  if(m.indexOf("popup")>=0) return "El navegador bloqueó la ventana de inicio de sesión: permite ventanas emergentes para este sitio.";
  return m;
}
return { CFG:CFG, norm:norm, esc:esc, parseBook:parseBook, byDim:byDim, byActivo:byActivo, totals:totals, snapshot:snapshot, compareSnaps:compareSnaps, DIMS:DIMS, CITY:CITY,
  auth:{init:initAuth,token:token}, graph:{findFile:findFile,download:download,versions:versions,cortes:cortesLoad}, errorMsg:errorMsg };
})();
