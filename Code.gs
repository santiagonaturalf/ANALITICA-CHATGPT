/**
 * @OnlyCurrentDoc
 * PRIORIDAD: toda hoja *_Corregido domina sobre lo importado por IMPORTRANGE.
 *
 * Flujo general:
 *  1) prepararInputs()  -> Data_SKU, Data_Ventas, Data_Adquisiciones
 *  2) generarSKU_A()
 *  3) generarAnalisisMargenes()
 *  4) _abrirDashboardMargenesInterno()
 *  5) _abrirDashboardPedidosHoyInterno()
 */

// ============== CONFIG ==============
const SOURCE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1hPyDsDHo6Sll6mYY_4YGcPJ4I9FPpG1kQINcidMM-s4/edit#gid=0";
const NEW_ACQ_URL      = "https://docs.google.com/spreadsheets/d/13DdVj-xf5PnjyHcGp2fV0GD65pHbgErtvORbWpFAaPw/edit?gid=0#gid=0";

// Hojas de corrección (persisten ediciones del dashboard)
const SHEET_CORR_PRECIO    = "Origen_Adquisicion_Corregido";   // [A]PrecioCompra, [B]Producto Base
const SHEET_CORR_FORMATO   = "Formato_Adquisicion_Corregido";  // [A]Producto Base, [B]Formato, [C]Cantidad, [D]Unidad
const SHEET_MAP_NP_BASE    = "SKU_Map_Corregido";              // [A]Nombre Producto, [B]Producto Base
const SHEET_VENTA_CORR     = "SKU_Venta_Corregido";            // [A]Nombre, [B]Cantidad Venta, [C]Unidad Venta
const SHEET_LASTPRICE_CORR = "Ventas_PrecioUltimo_Corregido";  // [A]Nombre, [B]Precio Ultimo
const SHEET_COSTO_NP_CORR  = "CostoAdquisicion_Corregido";     // [A]Nombre, [B]Costo de Adq
const SHEET_MARGEN_HOY     = "MargenHoy_Corregido";            // [A]Nombre, [B]Margen $ (hoy)
const SHEET_REVIEW         = "Margenes_Revision_Corregido";    // [A]Nombre Producto (revisión)

// Columnas relevantes en Data_Ventas
const VENTAS_COL_ORDEN    = 0;
const VENTAS_COL_CLIENTE  = 1;
const VENTAS_COL_NOMBRE   = 9;
const VENTAS_COL_CANTIDAD = 10;
const VENTAS_COL_PRECIO   = 11;
const VENTAS_COL_TOTAL    = 12;
const VENTAS_COL_FECHA    = 8;

// Lógicas de negocio
const TARGET_MARGIN_PCT = 0.20;  // sugiere "Revisar Precio" si el margen% < 20%
const DASH_LOW_PCT_RED  = 0.15;  // pinta fila roja si margen% < 15%

// ============== MENÚ ==============
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔥 Analítica Pro (simple)')
    .addItem('1) Preparar INPUTS (IMPORTRANGE)', 'prepararInputs')
    .addItem('2) Generar SKU_A', 'generarSKU_A')
    .addItem('3) Generar Análisis de Márgenes', 'generarAnalisisMargenes')
    .addSeparator()
    .addItem('4) Abrir Dashboard de Márgenes (interno)', '_abrirDashboardMargenesInterno')
    .addItem('5) Abrir Dashboard Pedidos de HOY (interno)', '_abrirDashboardPedidosHoyInterno')
    .addSeparator()
    .addItem('6) Realizar una cotización', 'abrirDashboardCotizacion')
    .addToUi();
}

// ============== 1) INPUTS (IMPORTRANGE) ==============
function prepararInputs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pairs = [
    ["Data_SKU",           "SKU!A:J"],
    ["Data_Ventas",        "Orders!A:N"],
    ["Data_Adquisiciones", "compras!A:AQ"]
  ];
  pairs.forEach(([dest, origin]) => {
    const sh = getOrCreateSheet(ss, dest, null, true);
    const url = (dest === "Data_Adquisiciones") ? NEW_ACQ_URL : SOURCE_SHEET_URL;
    sh.getRange("A1").setFormula(`=IMPORTRANGE("${url}"; "${origin}")`);
  });
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast('INPUTS listos con IMPORTRANGE', 'OK', 4);
}

// ============== 2) SKU_A ==============
// costo = precioCompra(base) / cantAdq(base) * cantVenta(nombre), con overrides.
function generarSKU_A() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shSKU = ss.getSheetByName("Data_SKU");
  const shADQ = ss.getSheetByName("Data_Adquisiciones");
  if (!shSKU || shSKU.getLastRow() < 2) throw new Error("Data_SKU vacío.");
  if (!shADQ || shADQ.getLastRow() < 2) throw new Error("Data_Adquisiciones vacío.");

  // precio reportado por base
  const precioBase = {};
  shADQ.getRange(2,1,shADQ.getLastRow()-1,2).getValues().forEach(r=>{
    const p=toNumber(r[0]); const b=String(r[1]||"").trim();
    if(b && !isNaN(p)) precioBase[b]=p;
  });

  // overrides
  const precioCorr  = readPrecioCorrecciones();   // base -> precio
  const formatoCorr = readFormatoCorrecciones();  // base -> {formato,cantidad,unidad}
  const npBaseCorr  = readNombreProductoBaseMap(); // nombre -> base
  const ventaCorr   = readVentaOverrides();       // nombre -> {cantVenta, unidadVenta}
  const costoNPCorr = readCostoAdqOverrides();    // nombre -> costo adq
  const precioFinal = Object.assign({}, precioBase, precioCorr);

  const outHeaders = [
    "Nombre Producto","Producto Base","Formato Adquisicion","Cantidad Adquisicion","Unidad Adquisicion",
    "Categoria","Cantidad Venta","Unidad Venta","Proveedor","Número de Teléfono",
    "Costo de Adquisicion","Fecha de Actualizacion"
  ];
  const out=[outHeaders];
  const now=new Date();

  shSKU.getRange(2,1,shSKU.getLastRow()-1,10).getValues().forEach(row=>{
    const nombre=row[0];
    let   base  = npBaseCorr[nombre] || row[1];
    let   formato=row[2];
    let   cantAdq=toNumber(row[3]);
    let   uniAdq =row[4];
    const cat    =row[5];
    let   cantV  =toNumber(row[6]);
    let   uniV   =row[7];
    const prov   =row[8];
    const fono   =row[9];

    // venta override (nombre)
    if (ventaCorr[nombre]) {
      if (!isNaN(toNumber(ventaCorr[nombre].cantVenta))) cantV = toNumber(ventaCorr[nombre].cantVenta);
      if (ventaCorr[nombre].unidadVenta) uniV = ventaCorr[nombre].unidadVenta;
    }
    // formato override (base)
    if (formatoCorr[base]) {
      formato = formatoCorr[base].formato || formato;
      cantAdq = !isNaN(toNumber(formatoCorr[base].cantidad)) ? toNumber(formatoCorr[base].cantidad) : cantAdq;
      uniAdq  = formatoCorr[base].unidad || uniAdq;
    }

    // precio compra por base (con override)
    const precio = precioFinal[base];
    let costo = 0;
    if (!isNaN(precio) && cantAdq > 0 && !isNaN(cantV)) costo = (precio / cantAdq) * cantV;

    // override directo por nombre
    if (!isNaN(toNumber(costoNPCorr[nombre]))) costo = toNumber(costoNPCorr[nombre]);

    out.push([nombre, base, formato, cantAdq, uniAdq, cat, cantV, uniV, prov, fono, costo, now]);
  });

  const dest = getOrCreateSheet(ss, "SKU_A", "CYAN", true);
  dest.getRange(1,1,out.length,out[0].length).setValues(out);
  dest.getRange("K:K").setNumberFormat("$#,##0");
  dest.getRange("L:L").setNumberFormat("yyyy-mm-dd HH:mm:ss");
  dest.autoResizeColumns(1, outHeaders.length);
  SpreadsheetApp.getActive().toast('SKU_A generado', 'OK', 4);
}

// ============== 3) Analisis_Margenes ==============
// agrega CostoCompra; recalcula J (Costo de Adquisicion) al final:
// override por nombre -> J,
// si no -> J=CostoCompra(C)/CantidadAdq(E)*CantidadVenta(H)
function generarAnalisisMargenes() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const ventas = ss.getSheetByName("Data_Ventas");
  const skua   = ss.getSheetByName("SKU_A");
  const adq    = ss.getSheetByName("Data_Adquisiciones");
  if (!ventas || ventas.getLastRow()<2) throw new Error("Data_Ventas vacío.");
  if (!skua   || skua.getLastRow()<2)   throw new Error("SKU_A vacío.");

  // precio por base (reportado + override)
  const precioBase={};
  if (adq && adq.getLastRow()>1){
    adq.getRange(2,1,adq.getLastRow()-1,2).getValues().forEach(r=>{
      const p=toNumber(r[0]); const b=String(r[1]||"").trim();
      if(b && !isNaN(p)) precioBase[b]=p;
    });
  }
  const precioCorr  = readPrecioCorrecciones();
  const precioFinal = Object.assign({}, precioBase, precioCorr);

  // desde SKU_A
  const costoMap={}; const skuInfo={};
  skua.getRange(2,1,skua.getLastRow()-1,12).getValues().forEach(r=>{
    const n=String(r[0]||"").trim(); if(!n) return;
    costoMap[n]=toNumber(r[10])||0;
    skuInfo[n]={productoBase:r[1], formatoAdq:r[2], cantAdq:r[3], unidadAdq:r[4], categoria:r[5], cantVenta:r[6], unidadVenta:r[7], fechaUpd:r[11]};
  });

  // overrides por nombre
  const ventaCorr    = readVentaOverrides();
  const lastPriceCorr= readPrecioUltimoOverrides();
  const costoNPCorr  = readCostoAdqOverrides();
  const margenHoyCorr= readMargenHoyOverrides();

  // agregado ventas (promedio/moda/ultimo)
  const agg={};
  const vv=ventas.getDataRange().getValues();
  for (let i=1;i<vv.length;i++){
    const nombre=String(vv[i][VENTAS_COL_NOMBRE]||"").trim();
    const p=toNumber(vv[i][VENTAS_COL_PRECIO]);
    if(!nombre||isNaN(p)) continue;
    if(!agg[nombre]) agg[nombre]={sum:0,count:0,freq:{},last:null};
    agg[nombre].sum+=p; agg[nombre].count++;
    const key=p.toFixed(2);
    agg[nombre].freq[key]=(agg[nombre].freq[key]||0)+1;
    agg[nombre].last=p;
  }

  const headers=[
    "Nombre Producto","Producto Base","CostoCompra",
    "Formato Adquisicion","Cantidad Adquisicion","Unidad Adquisicion","Categoria",
    "Cantidad Venta","Unidad Venta","Costo de Adquisicion","Fecha de Actualizacion",
    "Precio Venta Promedio (hoy)","Precio Venta Moda (hoy)","Precio Venta Último (hoy)",
    "Margen $ (promedio)","Margen $ (hoy)","Margen % (promedio)","Sugerencia"
  ];
  const out=[headers];

  Object.keys(agg).forEach(nombre=>{
    const a=agg[nombre];
    const avg=a.sum/a.count;
    let mode=0, maxf=0;
    Object.keys(a.freq).forEach(k=>{
      if(a.freq[k]>maxf){ maxf=a.freq[k]; mode=parseFloat(k); }
    });

    const s=skuInfo[nombre]||{};
    const base=String(s.productoBase||"").trim();
    const costoCompra = toNumber(precioFinal[base]);

    let cantV=s.cantVenta, uniV=s.unidadVenta;
    if (ventaCorr[nombre]) {
      if (!isNaN(toNumber(ventaCorr[nombre].cantVenta))) cantV = toNumber(ventaCorr[nombre].cantVenta);
      if (ventaCorr[nombre].unidadVenta) uniV = ventaCorr[nombre].unidadVenta;
    }

    let costo=costoMap[nombre]||0;
    if (!isNaN(toNumber(costoNPCorr[nombre]))) costo = toNumber(costoNPCorr[nombre]);

    let last=a.last;
    if (!isNaN(toNumber(lastPriceCorr[nombre]))) last = toNumber(lastPriceCorr[nombre]);

    const margenProm$ = avg - costo;
    let   margenHoy$  = (isNaN(last)?0:last - costo);
    if (!isNaN(toNumber(margenHoyCorr[nombre]))) margenHoy$ = toNumber(margenHoyCorr[nombre]);

    const margenPct   = (avg>0) ? (margenProm$/avg) : 0;
    const sugerencia  = (margenPct < TARGET_MARGIN_PCT) ? "Revisar Precio" : "OK";

    out.push([
      nombre, base, (isNaN(costoCompra) ? "" : costoCompra),
      s.formatoAdq||"", s.cantAdq||"", s.unidadAdq||"", s.categoria||"",
      cantV||"", uniV||"", costo, s.fechaUpd||"",
      avg, mode, last,
      margenProm$, margenHoy$, margenPct, sugerencia
    ]);
  });

  const dest=getOrCreateSheet(ss,"Analisis_Margenes","GREEN",true);
  dest.getRange(1,1,out.length,out[0].length).setValues(out);

  // formatos
  dest.getRange("C:C").setNumberFormat("$#,##0");    // CostoCompra
  dest.getRange("J:J").setNumberFormat("$#,##0");    // Costo de Adq
  dest.getRange("K:K").setNumberFormat("yyyy-mm-dd HH:mm:ss");
  dest.getRange("L:N").setNumberFormat("$#,##0");    // Precios
  dest.getRange("O:P").setNumberFormat("$#,##0");    // Márgenes $
  dest.getRange("Q:Q").setNumberFormat("0%");        // Márgen %

  // formato condicional por sugerencia
  dest.clearConditionalFormatRules();
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Revisar Precio").setBackground("#f4cccc").setBold(true)
    .setRanges([dest.getRange(2,18,Math.max(out.length-1,1),1)])
    .build();
  dest.setConditionalFormatRules([rule]);

  // Recalculo final de J (Costo de Adq)
  const lastRow=dest.getLastRow();
  if (lastRow>1){
    const n=lastRow-1;
    const colNombre=dest.getRange(2,1,n,1).getValues();
    const colC=dest.getRange(2,3,n,1).getValues(); // CostoCompra
    const colE=dest.getRange(2,5,n,1).getValues(); // CantAdq
    const colH=dest.getRange(2,8,n,1).getValues(); // CantVenta
    const overrideMap=readCostoAdqOverrides();     // por nombre

    const newJ=[];
    for (let i=0;i<n;i++){
      const nombre=String(colNombre[i][0]||"").trim();
      if (!isNaN(toNumber(overrideMap[nombre]))) {
        newJ.push([toNumber(overrideMap[nombre])]);
        continue;
      }
      const C=toNumber(colC[i][0]), E=toNumber(colE[i][0]), H=toNumber(colH[i][0]);
      let J="";
      if(!isNaN(C)&&C>0&&!isNaN(E)&&E>0&&!isNaN(H)&&H>=0) J=(C/E)*H;
      newJ.push([J]);
    }
    dest.getRange(2,10,n,1).setValues(newJ);
    dest.getRange("J:J").setNumberFormat("$#,##0");
  }

  dest.autoResizeColumns(1, headers.length);
  SpreadsheetApp.getActive().toast('Analisis_Margenes actualizado', 'OK', 4);
}

// ============== 4) DASHBOARD MÁRGENES (interno) ==============
function _abrirDashboardMargenesInterno() {
  const html = doGet();
  html.setWidth(1800).setHeight(980); // pantalla grande para comodidad
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard de Márgenes (interno)');
}

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Analisis_Margenes");
  if (!sh || sh.getLastRow()<2) {
    return HtmlService.createHtmlOutput("<h2>Genera primero la hoja Analisis_Margenes.</h2>");
  }

  const bases = collectProductoBaseSugerencias();
  const revisadoMap = readReviewMap(); // nombre -> true/false

  const vals = sh.getDataRange().getValues();
  const headers = vals[0].map(h=>String(h||"").trim());
  const idx = n => headers.indexOf(n);

  const iNombre   = idx("Nombre Producto");
  const iBase     = idx("Producto Base");
  const iCostoC   = idx("CostoCompra");
  const iForm     = idx("Formato Adquisicion");
  const iCantAdq  = idx("Cantidad Adquisicion");
  const iUniAdq   = idx("Unidad Adquisicion");
  const iCantV    = idx("Cantidad Venta");
  const iUniV     = idx("Unidad Venta");
  const iLast     = idx("Precio Venta Último (hoy)");
  const iMHoy     = idx("Margen $ (hoy)");
  const iCostoAdq = idx("Costo de Adquisicion");

  const rows=[];
  const summary = { venta: 0, costo: 0, margen: 0 };

  for (let r=1;r<vals.length;r++){
    const v=vals[r];
    const precioUlt = toNumber(v[iLast]);
    const costoAdq  = toNumber(v[iCostoAdq]);
    let   margenHoy = toNumber(v[iMHoy]);

    // Si el margen de la hoja es invalido, recalcularlo para la fila.
    if (isNaN(margenHoy)) {
      margenHoy = (!isNaN(precioUlt) && !isNaN(costoAdq)) ? (precioUlt - costoAdq) : 0;
    }
    const mPctHoy = (!isNaN(precioUlt) && precioUlt > 0) ? (margenHoy / precioUlt) : 0;

    // Sumar para el resumen
    if (!isNaN(precioUlt)) summary.venta += precioUlt;
    if (!isNaN(costoAdq))  summary.costo += costoAdq;

    const nombre = v[iNombre];
    rows.push({
      nombreProducto:   nombre,
      productoBase:     v[iBase],
      costoCompra:      v[iCostoC] || "",
      formatoAdq:       v[iForm],
      cantidadAdq:      v[iCantAdq],
      unidadAdq:        v[iUniAdq],
      cantidadVenta:    v[iCantV],
      unidadVenta:      v[iUniV],
      precioUltimo:     v[iLast],
      costoAdquisicion: v[iCostoAdq],
      margenHoy:        margenHoy,
      margenPctHoy:     mPctHoy,
      revisado:         !!revisadoMap[nombre]
    });
  }

  // Calcular el margen total a partir de los totales.
  summary.margen = summary.venta - summary.costo;

  const tpl = HtmlService.createTemplateFromFile('dashboard_margenes.html');
  tpl.data    = rows;
  tpl.bases   = bases;
  tpl.lowPct  = DASH_LOW_PCT_RED; // 0.15
  tpl.summary = summary;
  return tpl.evaluate()
            .setTitle("Dashboard de Márgenes (interno)")
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============== 5) DASHBOARD PEDIDOS HOY (interno) ==============
function _abrirDashboardPedidosHoyInterno() {
  const html = doGetPedidos();
  html.setWidth(1800).setHeight(980);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Pedidos de HOY');
}

function doGetPedidos() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const ventas = ss.getSheetByName("Data_Ventas");
  const skua   = ss.getSheetByName("SKU_A");
  if (!ventas || ventas.getLastRow()<2) return HtmlService.createHtmlOutput("<h2>Falta hoja Data_Ventas.</h2>");
  if (!skua   || skua.getLastRow()<2)   return HtmlService.createHtmlOutput("<h2>Falta hoja SKU_A.</h2>");

  // Mapa de revisión por Nombre Producto
  const revisadoMap = readReviewMap();

  // Mapas desde SKU_A
  const maps = getSkuMapsFromSheet(skua);   // {costoMap, skuInfo}
  const costoMap = maps.costoMap;
  const skuInfo  = maps.skuInfo;

  const tz = Session.getScriptTimeZone() || "America/Santiago";
  const today = new Date(Utilities.formatDate(new Date(), tz, "yyyy-MM-dd")); // 00:00 local
  const start = new Date(today.getTime());
  const end   = new Date(today.getTime());
  end.setDate(end.getDate()+1); // < end

  const vv = ventas.getDataRange().getValues();
  /** orders: { [numPedido]: { numero, cliente, lineas:[], totVenta, totCosto, totMargen, pct } } */
  const orders = {};

  for (let i=1;i<vv.length;i++){
    const row = vv[i];
    const f   = row[VENTAS_COL_FECHA];
    const d   = parseAsDate(f);
    if (!d || d < start || d >= end) continue; // solo pedidos de hoy

    const pedido  = String(row[VENTAS_COL_ORDEN]||"").trim();
    const cliente = String(row[VENTAS_COL_CLIENTE]||"").trim();
    const nombre  = String(row[VENTAS_COL_NOMBRE]||"").trim();
    const cant    = toNumber(row[VENTAS_COL_CANTIDAD]);
    const pUnit   = toNumber(row[VENTAS_COL_PRECIO]);
    const totalLn = toNumber(row[VENTAS_COL_TOTAL]);

    if (!pedido || !nombre || isNaN(cant)) continue;

    const cUnit   = toNumber(costoMap[nombre]); // costo por unidad de venta
    const cLinea  = (!isNaN(cUnit) && !isNaN(cant)) ? Math.round(cUnit * cant) : 0;
    const ventaLn = !isNaN(totalLn) ? totalLn : (isNaN(pUnit) ? 0 : Math.round(pUnit * cant));
    const margenL = ventaLn - cLinea;
    const pctUni  = (!isNaN(pUnit) && pUnit>0) ? ( (pUnit - (isNaN(cUnit)?0:cUnit)) / pUnit ) : 0;

    if (!orders[pedido]) orders[pedido] = {numero:pedido, cliente:cliente, lineas:[], totVenta:0, totCosto:0, totMargen:0};
    orders[pedido].lineas.push({
      nombreProducto: nombre,
      cantidad: cant,
      precioUnit: pUnit,
      totalLinea: ventaLn,
      costoUnit: cUnit,
      costoLinea: cLinea,
      margenLinea: margenL,
      pctUnidad: pctUni,
      // información adicional para el subformulario
      base:        (skuInfo[nombre]||{}).productoBase || "",
      costoCompra: (skuInfo[nombre]||{}).costoCompra  || "",
      formatoAdq:  (skuInfo[nombre]||{}).formatoAdq   || "",
      cantAdq:     (skuInfo[nombre]||{}).cantAdq      || "",
      uniAdq:      (skuInfo[nombre]||{}).unidadAdq    || "",
      cantVenta:   (skuInfo[nombre]||{}).cantVenta    || "",
      uniVenta:    (skuInfo[nombre]||{}).unidadVenta  || "",
      fechaUpd:    (skuInfo[nombre]||{}).fechaUpd     || "",
      revisado:    !!revisadoMap[nombre],
      categoria:   (skuInfo[nombre]||{}).categoria || ""
    });
    orders[pedido].totVenta  += ventaLn;
    orders[pedido].totCosto  += cLinea;
    orders[pedido].totMargen += margenL;
  }

  const listado = Object.values(orders).map(o=>{
    const pct = (o.totVenta>0) ? (o.totMargen/o.totVenta) : 0;
    return Object.assign(o,{pct});
  }).sort((a,b)=>a.numero.localeCompare(b.numero,'es'));

  // resumen general
  const resumen = listado.reduce((acc,o)=>{
    acc.pedidos += 1;
    acc.venta   += o.totVenta;
    acc.costo   += o.totCosto;
    acc.margen  += o.totMargen;
    return acc;
  }, {pedidos:0, venta:0, costo:0, margen:0});
  resumen.pct = (resumen.venta>0)? (resumen.margen/resumen.venta) : 0;
  resumen.ticket = (resumen.pedidos>0)? Math.round(resumen.venta/resumen.pedidos) : 0;

  const bases = collectProductoBaseSugerencias();

  // Lista de categorías únicas para el filtro
  const catSet = new Set();
  listado.forEach(o => {
    o.lineas.forEach(ln => { if (ln.categoria) catSet.add(ln.categoria); });
  });
  const categorias = Array.from(catSet).sort((a,b)=>a.localeCompare(b,'es'));

  const tpl = HtmlService.createTemplateFromFile('dashboard_pedidos.html');
  tpl.pedidos    = listado;
  tpl.resumen    = resumen;
  tpl.lowPct     = DASH_LOW_PCT_RED;
  tpl.bases      = bases;
  tpl.categorias = categorias;
  return tpl.evaluate()
            .setTitle("Dashboard Pedidos de HOY")
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============== GUARDADO (desde dashboard) ==============
/**
 * Recibe payload:
 * { nombreProducto, productoBase, costoCompra, formato, cantidadAdq, unidadAdq,
 *   cantidadVenta, unidadVenta, precioUltimo, costoAdquisicion, margenHoy }
 * Modifica las tablas de correcciones y recalcula hojas.
 */
function saveDashboardEdits(payload) {
  try {
    const p = payload || {};
    const nombre = String(p.nombreProducto || "").trim();
    if (!nombre) return "Error: Falta Nombre Producto.";

    // Mapa Nombre -> Base
    if (p.productoBase) upsertNombreProductoBase(nombre, String(p.productoBase).trim());

    // CostoCompra por Base
    if (!isNaN(toNumber(p.costoCompra)) && toNumber(p.costoCompra) > 0) {
      upsertPrecioCorreccion(String(p.productoBase || "").trim(), toNumber(p.costoCompra));
    }

    // Formato por Base
    if (p.productoBase && (p.formato || p.cantidadAdq || p.unidadAdq)) {
      const cant = (!isNaN(toNumber(p.cantidadAdq)) ? toNumber(p.cantidadAdq) : "");
      upsertFormatoCorreccion(String(p.productoBase).trim(), p.formato || "", cant, p.unidadAdq || "");
    }

    // Venta por Nombre
    if (!isNaN(toNumber(p.cantidadVenta)) || p.unidadVenta) {
      upsertVentaOverride(nombre, (!isNaN(toNumber(p.cantidadVenta)) ? toNumber(p.cantidadVenta) : ""), p.unidadVenta || "");
    }

    // Precio Último / Costo / Margen por Nombre
    if (!isNaN(toNumber(p.precioUltimo)))      upsertPrecioUltimoOverride(nombre, toNumber(p.precioUltimo));
    if (!isNaN(toNumber(p.costoAdquisicion)))  upsertCostoAdqOverride(nombre, toNumber(p.costoAdquisicion));
    if (!isNaN(toNumber(p.margenHoy)))         upsertMargenHoyOverride(nombre, toNumber(p.margenHoy));

    // Recalcular salidas
    generarSKU_A();
    generarAnalisisMargenes();
    return "¡Cambios guardados y hojas recalculadas!";
  } catch (e) {
    return "Error al guardar: " + e.message;
  }
}

// ============== REVISIÓN (toggle verde, sin recalcular) ==============
function readReviewMap(){
  const sh = getOrCreateSheet(SpreadsheetApp.getActive(), SHEET_REVIEW, "LIGHTGREEN", false);
  const map = {};
  if (sh.getLastRow()<2){
    sh.getRange(1,1,1,2).setValues([["Nombre Producto","Revisado"]]);
    return map;
  }
  sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(r=>{
    const n=String(r[0]||"").trim(); const f=String(r[1]||"").trim();
    if (!n) return;
    const val = (f==="1" || f==="TRUE" || f==="SI" || f==="OK" || f==="✓");
    map[n]=val;
  });
  return map;
}
function upsertReview(nombre,flag){
  const sh = getOrCreateSheet(SpreadsheetApp.getActive(), SHEET_REVIEW, "LIGHTGREEN", false);
  if (sh.getLastRow()<1) sh.getRange(1,1,1,2).setValues([["Nombre Producto","Revisado"]]);
  const last=sh.getLastRow();
  if (last>1){
    const vals=sh.getRange(2,1,last-1,1).getValues();
    for (let i=0;i<vals.length;i++){
      if (String(vals[i][0]||"").trim()===nombre){
        sh.getRange(i+2,2).setValue(flag?1:0);
        return;
      }
    }
  }
  sh.appendRow([nombre, flag?1:0]);
}
function toggleRevision(nombre, flag){
  upsertReview(String(nombre||"").trim(), !!flag);
  return "OK";
}

// ============== CRUD de hojas *_Corregido ==============
function readPrecioCorrecciones(){
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_CORR_PRECIO,"ORANGE",false);
  const map={};
  if (sh.getLastRow()<2){
    sh.getRange(1,1,1,2).setValues([["Precio","Producto Base"]]);
    return map;
  }
  sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(r=>{
    const p=toNumber(r[0]); const b=String(r[1]||"").trim();
    if(b && !isNaN(p)) map[b]=p;
  });
  return map;
}
function upsertPrecioCorreccion(base,precio){
  if(!base) return;
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_CORR_PRECIO,"ORANGE",false);
  if (sh.getLastRow()<1) sh.getRange(1,1,1,2).setValues([["Precio","Producto Base"]]);
  const last=sh.getLastRow();
  if (last>1){
    const vals=sh.getRange(2,1,last-1,2).getValues();
    for (let i=0;i<vals.length;i++){
      if (String(vals[i][1]||"").trim()===base){
        sh.getRange(i+2,1).setValue(precio);
        return;
      }
    }
  }
  sh.appendRow([precio,base]);
}

function readFormatoCorrecciones(){
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_CORR_FORMATO,"ORANGE",false);
  const map={};
  if (sh.getLastRow()<2){
    sh.getRange(1,1,1,4).setValues([["Producto Base","Formato Adquisicion","Cantidad Adquisicion","Unidad Adquisicion"]]);
    return map;
  }
  sh.getRange(2,1,sh.getLastRow()-1,4).getValues().forEach(r=>{
    const b=String(r[0]||"").trim();
    if(!b) return;
    map[b]={formato:r[1], cantidad:toNumber(r[2]), unidad:r[3]};
  });
  return map;
}
function upsertFormatoCorreccion(base,formato,cantidad,unidad){
  if(!base) return;
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_CORR_FORMATO,"ORANGE",false);
  if (sh.getLastRow()<1) sh.getRange(1,1,1,4).setValues([["Producto Base","Formato Adquisicion","Cantidad Adquisicion","Unidad Adquisicion"]]);
  const last=sh.getLastRow();
  if (last>1){
    const vals=sh.getRange(2,1,last-1,4).getValues();
    for (let i=0;i<vals.length;i++){
      if (String(vals[i][0]||"").trim()===base){
        sh.getRange(i+2,2,1,3).setValues([[formato,cantidad,unidad]]);
        return;
      }
    }
  }
  sh.appendRow([base,formato,cantidad,unidad]);
}

function readNombreProductoBaseMap(){
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_MAP_NP_BASE,"ORANGE",false);
  const map={};
  if (sh.getLastRow()<2){
    sh.getRange(1,1,1,2).setValues([["Nombre Producto","Producto Base"]]);
    return map;
  }
  sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(r=>{
    const n=String(r[0]||"").trim();
    const b=String(r[1]||"").trim();
    if(n && b) map[n]=b;
  });
  return map;
}
function upsertNombreProductoBase(nombre,base){
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_MAP_NP_BASE,"ORANGE",false);
  if (sh.getLastRow()<1) sh.getRange(1,1,1,2).setValues([["Nombre Producto","Producto Base"]]);
  const last=sh.getLastRow();
  if (last>1){
    const vals=sh.getRange(2,1,last-1,2).getValues();
    for (let i=0;i<vals.length;i++){
      if (String(vals[i][0]||"").trim()===nombre){
        sh.getRange(i+2,2).setValue(base);
        return;
      }
    }
  }
  sh.appendRow([nombre,base]);
}

function readVentaOverrides(){
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_VENTA_CORR,"ORANGE",false);
  const map={};
  if (sh.getLastRow()<2){
    sh.getRange(1,1,1,3).setValues([["Nombre Producto","Cantidad Venta","Unidad Venta"]]);
    return map;
  }
  sh.getRange(2,1,sh.getLastRow()-1,3).getValues().forEach(r=>{
    const n=String(r[0]||"").trim();
    if(!n) return;
    map[n]={cantVenta:r[1], unidadVenta:r[2]};
  });
  return map;
}
function upsertVentaOverride(nombre,cant,unidad){
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_VENTA_CORR,"ORANGE",false);
  if (sh.getLastRow()<1) sh.getRange(1,1,1,3).setValues([["Nombre Producto","Cantidad Venta","Unidad Venta"]]);
  const last=sh.getLastRow();
  if (last>1){
    const vals=sh.getRange(2,1,last-1,3).getValues();
    for (let i=0;i<vals.length;i++){
      if (String(vals[i][0]||"").trim()===nombre){
        sh.getRange(i+2,2,1,2).setValues([[cant,unidad]]);
        return;
      }
    }
  }
  sh.appendRow([nombre,cant,unidad]);
}

function readPrecioUltimoOverrides(){
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_LASTPRICE_CORR,"ORANGE",false);
  const map={};
  if (sh.getLastRow()<2){
    sh.getRange(1,1,1,2).setValues([["Nombre Producto","Precio Último (hoy)"]]);
    return map;
  }
  sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(r=>{
    const n=String(r[0]||"").trim();
    const p=toNumber(r[1]);
    if(n && !isNaN(p)) map[n]=p;
  });
  return map;
}
function upsertPrecioUltimoOverride(nombre,precio){
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_LASTPRICE_CORR,"ORANGE",false);
  if (sh.getLastRow()<1) sh.getRange(1,1,1,2).setValues([["Nombre Producto","Precio Último (hoy)"]]);
  const last=sh.getLastRow();
  if (last>1){
    const vals=sh.getRange(2,1,last-1,2).getValues();
    for (let i=0;i<vals.length;i++){
      if (String(vals[i][0]||"").trim()===nombre){
        sh.getRange(i+2,2).setValue(precio);
        return;
      }
    }
  }
  sh.appendRow([nombre,precio]);
}

function readCostoAdqOverrides(){
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_COSTO_NP_CORR,"ORANGE",false);
  const map={};
  if (sh.getLastRow()<2){
    sh.getRange(1,1,1,2).setValues([["Nombre Producto","Costo de Adquisicion"]]);
    return map;
  }
  sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(r=>{
    const n=String(r[0]||"").trim();
    const c=toNumber(r[1]);
    if(n && !isNaN(c)) map[n]=c;
  });
  return map;
}
function upsertCostoAdqOverride(nombre,costo){
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_COSTO_NP_CORR,"ORANGE",false);
  if (sh.getLastRow()<1) sh.getRange(1,1,1,2).setValues([["Nombre Producto","Costo de Adquisicion"]]);
  const last=sh.getLastRow();
  if (last>1){
    const vals=sh.getRange(2,1,last-1,2).getValues();
    for (let i=0;i<vals.length;i++){
      if (String(vals[i][0]||"").trim()===nombre){
        sh.getRange(i+2,2).setValue(costo);
        return;
      }
    }
  }
  sh.appendRow([nombre,costo]);
}

function readMargenHoyOverrides(){
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_MARGEN_HOY,"ORANGE",false);
  const map={};
  if (sh.getLastRow()<2){
    sh.getRange(1,1,1,2).setValues([["Nombre Producto","Margen $ (hoy)"]]);
    return map;
  }
  sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(r=>{
    const n=String(r[0]||"").trim();
    const m=toNumber(r[1]);
    if(n && !isNaN(m)) map[n]=m;
  });
  return map;
}
function upsertMargenHoyOverride(nombre,margen){
  const sh=getOrCreateSheet(SpreadsheetApp.getActive(),SHEET_MARGEN_HOY,"ORANGE",false);
  if (sh.getLastRow()<1) sh.getRange(1,1,1,2).setValues([["Nombre Producto","Margen $ (hoy)"]]);
  const last=sh.getLastRow();
  if (last>1){
    const vals=sh.getRange(2,1,last-1,2).getValues();
    for (let i=0;i<vals.length;i++){
      if (String(vals[i][0]||"").trim()===nombre){
        sh.getRange(i+2,2).setValue(margen);
        return;
      }
    }
  }
  sh.appendRow([nombre,margen]);
}

function collectProductoBaseSugerencias(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const set=new Set();
  const shSKU=ss.getSheetByName("Data_SKU");
  if (shSKU && shSKU.getLastRow()>1){
    shSKU.getRange(2,2,shSKU.getLastRow()-1,1).getValues().forEach(r=>{
      const s=String(r[0]||"").trim();
      if (s) set.add(s);
    });
  }
  Object.keys(readPrecioCorrecciones()).forEach(b=>set.add(b));
  Object.keys(readFormatoCorrecciones()).forEach(b=>set.add(b));
  Object.values(readNombreProductoBaseMap()).forEach(b=>set.add(b));
  return Array.from(set).sort((a,b)=>a.localeCompare(b,'es'));
}

// ============== 6) COTIZACIONES (Dashboard) ==============
function abrirDashboardCotizacion() {
  const productos = getProductosParaCotizar();
  if (!productos) return; // Error message is handled inside the function

  const tpl = HtmlService.createTemplateFromFile('dashboard_cotizacion.html');
  tpl.productos = productos;

  const html = tpl.evaluate().setWidth(1000).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generador de Cotizaciones');
}

function getProductosParaCotizar() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SKU_A");

  if (!sheet) {
    ui.alert('No se encontró la hoja "SKU_A". Por favor, generela primero usando la opción del menú.');
    return null;
  }

  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11); // Read until column K
  const values = dataRange.getValues();

  const productos = values.map(function(row) {
    const nombre = row[0]; // Column A
    const costo = toNumber(row[10]); // Column K
    if (nombre && !isNaN(costo)) {
      return { nombre: nombre, costo: costo };
    }
    return null;
  }).filter(function(p) { return p; });

  if (productos.length === 0) {
    ui.alert('No se encontraron productos con costo válido en la hoja "SKU_A".');
    return null;
  }

  return productos;
}

function crearHojaDeCotizacion(datosCotizacion) {
  if (!datosCotizacion || !Array.isArray(datosCotizacion) || datosCotizacion.length === 0) {
    return "Error: No se recibieron datos para generar la cotización.";
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const date = new Date();
  const formattedDate = (date.getMonth() + 1) + '-' + date.getDate() + '-' + date.getFullYear();
  const newSheetName = 'Cotizacion - ' + formattedDate;

  const newSheet = ss.insertSheet(newSheetName);

  const headers = ["Producto", "Costo Base", "% IVA", "% Margen", "Precio de Venta"];
  const outputData = [headers];

  datosCotizacion.forEach(function(item) {
    outputData.push([
      item.nombre,
      item.costo,
      item.ivaPct,
      item.margenPct,
      item.precioVenta
    ]);
  });

  newSheet.getRange(1, 1, outputData.length, headers.length).setValues(outputData);

  // Format columns
  newSheet.getRange(2, 2, outputData.length - 1, 1).setNumberFormat("$#,##0"); // Costo Base
  newSheet.getRange(2, 3, outputData.length - 1, 2).setNumberFormat("0.00%"); // IVA % and Margen %
  newSheet.getRange(2, 5, outputData.length - 1, 1).setNumberFormat("$#,##0"); // Precio de Venta

  newSheet.autoResizeColumns(1, headers.length);

  return 'Cotización generada exitosamente en la hoja "' + newSheetName + '".';
}

// ============== HELPERS ==============
/**
 * Convierte un valor a número (acepta formatos "12.345,67", "$1.234", etc.)
 */
function toNumber(v){
  if (typeof v === 'number') return v;
  if (v == null) return NaN;
  let s = String(v).trim();
  if (!s) return NaN;

  s = s.replace(/\s+/g,'').replace(/\$/g,'');

  const hasDot   = s.indexOf('.') >= 0;
  const hasComma = s.indexOf(',') >= 0;

  if (hasDot && hasComma) {
    // Formato con miles y decimales (CL): 12.345,67 -> 12345.67
    s = s.replace(/\./g,'').replace(',', '.');
  } else if (hasComma && !hasDot) {
    // Solo coma (decimales)
    s = s.replace(',', '.');
  }
  const n = parseFloat(s);
  return isNaN(n) ? NaN : n;
}

/**
 * Crea u obtiene una hoja; si clear=true, la limpia.
 */
function getOrCreateSheet(ss, name, color /*=null*/, clear /*=false*/) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (clear) {
    sh.clear();
    if (sh.getMaxRows() > 2000) sh.deleteRows(2001, sh.getMaxRows() - 2000);
    if (sh.getMaxColumns() > 26) sh.deleteColumns(27, sh.getMaxColumns() - 26);
  }
  if (color) sh.setTabColor(color);
  return sh;
}

/**
 * Lee la SKU_A y construye mapas reutilizables (costoMap y skuInfo).
 */
function getSkuMapsFromSheet(skuaSheet){
  const costoMap={}; const skuInfo={};
  const vals = skuaSheet.getRange(2,1,skuaSheet.getLastRow()-1,12).getValues();
  vals.forEach(r=>{
    const n=String(r[0]||"").trim();
    if(!n) return;
    costoMap[n]=toNumber(r[10])||0;
    skuInfo[n]={
      productoBase:r[1], formatoAdq:r[2], cantAdq:r[3], unidadAdq:r[4], categoria:r[5],
      cantVenta:r[6], unidadVenta:r[7], costoCompra:r[2]?null:null,
      fechaUpd:r[11]
    };
  });
  // CostoCompra no viene en SKU_A; si se necesita, se completa desde Analisis_Margenes.
  const am = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analisis_Margenes");
  if (am && am.getLastRow()>1){
    const headers = am.getRange(1,1,1,am.getLastColumn()).getValues()[0].map(h=>String(h||"").trim());
    const iNombre = headers.indexOf("Nombre Producto");
    const iBase   = headers.indexOf("Producto Base");
    const iCostoC = headers.indexOf("CostoCompra");
    const rows = am.getRange(2,1,am.getLastRow()-1,am.getLastColumn()).getValues();
    rows.forEach(v=>{
      const n = String(v[iNombre]||"").trim();
      if (!n) return;
      if (!skuInfo[n]) skuInfo[n]={};
      skuInfo[n].productoBase = skuInfo[n].productoBase || v[iBase];
      skuInfo[n].costoCompra  = skuInfo[n].costoCompra  || toNumber(v[iCostoC]);
    });
  }
  return {costoMap, skuInfo};
}

/**
 * Intenta parsear cualquier valor fecha de Data_Ventas a Date local.
 */
function parseAsDate(v){
  if (v instanceof Date) return v;
  const s = String(v||"").trim();
  if (!s) return null;
  // formatos típicos: "2025-08-28 9:25", "2025/08/28 09:25:00", etc.
  const s2 = s.replace('T',' ').replace(/\//g,'-');
  const d  = new Date(s2);
  return isNaN(d.getTime()) ? null : d;
}