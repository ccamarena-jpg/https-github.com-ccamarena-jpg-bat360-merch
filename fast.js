/* Rapidez de las vistas: caché local (IndexedDB) de lo ya leído de cada Excel.
   Medido con los archivos reales: leer el Excel tarda poco (Consolidado 13 MB ≈ 1,7 s); lo lento es DESCARGARLO de SharePoint en cada apertura.
   Por eso cada vista guarda el resultado ya procesado junto con la «versión» del archivo (id + fecha de modificación + tamaño) y solo
   vuelve a descargar los archivos que cambiaron. Además muestra al instante lo último guardado mientras revisa si hay algo nuevo.
   La caché vive en este navegador (no sale del equipo). Si falla (modo privado, sin espacio), todo sigue funcionando sin ella. */
window.FAST = (function(){
"use strict";
var DB = "bat360merch", ST = "kv";
function open(){
  return new Promise(function(ok,no){
    try{ var r=indexedDB.open(DB,1); r.onupgradeneeded=function(){ r.result.createObjectStore(ST); }; r.onsuccess=function(){ ok(r.result); }; r.onerror=function(){ no(r.error); }; }
    catch(e){ no(e); }
  });
}
function get(key){
  return open().then(function(db){ return new Promise(function(ok){ var q=db.transaction(ST).objectStore(ST).get(key); q.onsuccess=function(){ ok(q.result||null); }; q.onerror=function(){ ok(null); }; }); })
    .catch(function(){ return null; });
}
function set(key,val){
  return open().then(function(db){ return new Promise(function(ok){ var t=db.transaction(ST,"readwrite"); t.objectStore(ST).put(val,key); t.oncomplete=function(){ ok(true); }; t.onerror=t.onabort=function(){ ok(false); }; }); })
    .catch(function(){ return false; });
}
function del(key){
  return open().then(function(db){ return new Promise(function(ok){ var t=db.transaction(ST,"readwrite"); t.objectStore(ST).delete(key); t.oncomplete=function(){ ok(true); }; t.onerror=function(){ ok(false); }; }); }).catch(function(){ return false; });
}
/* Versión de un archivo de Graph: cambia cuando alguien lo guarda. «extra» permite invalidar por otros motivos (p. ej. el día). */
function ver(item, extra){ return [item.id, item.lastModifiedDateTime||"", item.size||"", extra||""].join("|"); }
function day(){ var d=new Date(); return d.getFullYear()+"-"+(d.getMonth()+1)+"-"+d.getDate(); }
function ago(ts){
  var m=Math.round((Date.now()-ts)/60000);
  if(m<1) return "hace un momento"; if(m<60) return "hace "+m+" min";
  var h=Math.round(m/60); if(h<24) return "hace "+h+" h"; var d=Math.round(h/24); return "hace "+d+" día"+(d===1?"":"s");
}
/* Opciones de lectura de SheetJS sin lo que estas vistas no usan (HTML por celda, fórmulas, estilos): más rápido y con menos memoria. */
var LEAN = {cellHTML:false,cellFormula:false,cellStyles:false,cellNF:false,cellText:false,bookDeps:false};
function lean(o){ var r={}; Object.keys(LEAN).forEach(function(k){ r[k]=LEAN[k]; }); Object.keys(o||{}).forEach(function(k){ r[k]=o[k]; }); return r; }
return { get:get, set:set, del:del, ver:ver, day:day, ago:ago, lean:lean };
})();
