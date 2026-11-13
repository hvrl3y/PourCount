/* Tallypour exports: Excel workbook, PDF report, CSV files.
   index.html calls tallypourExports(T) with its state and helpers and gets back {exportSheet, exRun}. */
window.tallypourExports = function(T){
"use strict";
const {calc,totalOz,sortedProducts,OZ,toast,$,ic,I,esc,fmtDate,openSheet} = T;

/* ---------- loading ---------- */
const EX_LIBS = {};
function exLoad(src){
  if(EX_LIBS[src]) return EX_LIBS[src];
  EX_LIBS[src] = new Promise((res,rej)=>{ const s=document.createElement("script"); s.src=src; s.onload=res; s.onerror=()=>{ delete EX_LIBS[src]; rej(new Error("load "+src)); }; document.head.appendChild(s); });
  return EX_LIBS[src];
}
const exLoadExcel = () => exLoad("vendor/exceljs.min.js");
const exLoadPdf = () => exLoad("vendor/jspdf.umd.min.js").then(()=>exLoad("vendor/jspdf.plugin.autotable.min.js"));

/* ---------- data ---------- */
function exShelfName(id){ const s=(T.S.shelves||[]).find(x=>x.id===id); return s?s.name:""; }
function exBarName(){ return (T.S.settings.bar||"").trim(); }
function exStamp(ts){ const d=new Date(ts); const p=n=>String(n).padStart(2,"0"); return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}`; }
function exFileBase(ci){
  const c=T.S.counts[ci], bar=exBarName().replace(/[^A-Za-z0-9]+/g,"-").replace(/^-|-$/g,"");
  return ["Tallypour", bar||null, exStamp(c?c.date:Date.now())].filter(Boolean).join("_");
}
function exLongDate(ts){ return new Date(ts).toLocaleString("en-US",{weekday:"short",month:"short",day:"numeric",year:"numeric",hour:"numeric",minute:"2-digit"}); }

/* Everything the reports need, for count index ci (compared with the count before it). */
function exData(ci){
  if(ci==null) ci = T.S.counts.length-1;
  const c = T.S.counts[ci]||null, pv = ci>0 ? T.S.counts[ci-1] : null;
  const target = T.S.settings.target;
  const rows = sortedProducts().map(p=>{
    const k = calc(p), bOz = k.bOz;
    const e = c && c.items[p.id];
    const onOz = totalOz(p,e);
    const r = {p,k,bOz,cat:p.cat||"Other",shelf:exShelfName(p.shelf),vendor:(p.vendor||"").trim(),par:p.par>0?p.par:null,
      counted:!!e, sealed:e?(e.sealed||0):null, open:e?(e.open||0):null, onOz,
      onBtl:onOz!=null&&bOz>0?onOz/bOz:null, value:onOz!=null?onOz*k.costOz:null,
      price:p.price>0?p.price:null};
    r.status = k.pc==null ? "No price" : k.pc<=target ? "On target" : k.pc<=target+5 ? "Close" : "Over target";
    r.toOrder = r.par!=null ? Math.max(0, Math.ceil(r.par - (r.onBtl||0) - 1e-9)) : null;
    r.orderCost = r.toOrder!=null ? r.toOrder*(p.cost||0) : null;
    if(pv){
      const a = totalOz(p,pv.items[p.id]), purch=(c.purchases||{})[p.id]||0, sold=(c.sales||{})[p.id];
      if(a!=null || purch){
        r.u = {start:a||0, recv:purch, end:onOz, sold:sold!=null?sold:null};
        r.u.used = onOz==null ? null : r.u.start + purch*bOz - onOz;
        r.u.exp = sold!=null ? sold*k.pour : null;
        r.u.v = r.u.used!=null && r.u.exp!=null ? r.u.used - r.u.exp : null;
        r.u.usedCost = r.u.used!=null ? r.u.used*k.costOz : null;
        r.u.vCost = r.u.v!=null ? r.u.v*k.costOz : null;
        r.u.vPct = r.u.v!=null && r.u.exp>0 ? r.u.v/r.u.exp : null;
        r.u.rev = sold!=null && r.price ? sold*r.price : null;
      }
    }
    return r;
  });
  const sum = (arr,f)=>arr.reduce((s,x)=>{ const v=f(x); return s+(v==null||isNaN(v)?0:v); },0);
  const usage = rows.filter(r=>r.u);
  const priced = rows.filter(r=>r.k.pc!=null);
  const t = {
    products: rows.length, counted: rows.filter(r=>r.counted).length,
    value: sum(rows,r=>r.value),
    avgPc: priced.length ? sum(priced,r=>r.k.pc)/priced.length : null,
    over: rows.filter(r=>r.status==="Over target").length,
    usedCost: sum(usage,r=>r.u.usedCost), vCost: sum(usage,r=>r.u.vCost), rev: sum(usage,r=>r.u.rev),
    soldCost: sum(usage,r=>r.u.exp!=null?r.u.exp*r.k.costOz:null),
    orderCost: sum(rows,r=>r.orderCost), orderItems: sum(rows,r=>r.toOrder), hasPar: rows.some(r=>r.par!=null)
  };
  t.actualPc = t.rev>0 ? t.usedCost/t.rev : null;   // what was poured vs what was rung up
  t.idealPc = t.rev>0 ? t.soldCost/t.rev : null;    // what it should have cost
  const cats = {};
  rows.forEach(r=>{ const g = cats[r.cat] || (cats[r.cat]={cat:r.cat,n:0,value:0,usedCost:0,vCost:0,rev:0});
    g.n++; g.value+=r.value||0; if(r.u){ g.usedCost+=r.u.usedCost||0; g.vCost+=r.u.vCost||0; g.rev+=r.u.rev||0; } });
  return {ci,c,pv,rows,usage,t,cats:Object.values(cats).sort((a,b)=>b.value-a.value),target};
}

/* ---------- delivering a file ---------- */
async function exDeliver(blob, name){
  try{
    const file = new File([blob], name, {type: blob.type});
    if(navigator.canShare && navigator.canShare({files:[file]})){
      try{ await navigator.share({files:[file], title:name}); return "shared"; }
      catch(e){ if(e && e.name==="AbortError") return "cancelled"; }
    }
  }catch(e){}
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; a.rel="noopener";
  document.body.appendChild(a); a.click(); setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 8000);
  return "downloaded";
}

/* ---------- Excel workbook ---------- */
const XL = {
  ink:"FF1C1C1E", band:"FF2C2C2E", amber:"FFFF9F0A", amberSoft:"FFFFF4E0", stripe:"FFF6F6F8", line:"FFD9D9DE", sub:"FF6E6E73",
  okF:"FFE3F6E8", okT:"FF1E7A34", warnF:"FFFFF1D6", warnT:"FF9A5B00", badF:"FFFDE2E1", badT:"FFB3261E",
  money:'"$"#,##0.00;[Red]-"$"#,##0.00', money0:'"$"#,##0;[Red]-"$"#,##0', oz:'#,##0.00', oz1:'#,##0.0', int:'#,##0', pct:'0.0%;[Red]-0.0%', btl:'#,##0.00'
};
const BOZ = OZ; // ml per oz, used inside formulas as 29.5735

function xlTitle(ws, title, sub, lastCol){
  ws.mergeCells(1,1,1,lastCol); ws.mergeCells(2,1,2,lastCol);
  const a = ws.getCell(1,1); a.value = title; a.font = {name:"Calibri",size:16,bold:true,color:{argb:"FFFFFFFF"}};
  a.fill = {type:"pattern",pattern:"solid",fgColor:{argb:XL.ink}}; a.alignment = {vertical:"middle",indent:1};
  const b = ws.getCell(2,1); b.value = sub; b.font = {size:10,color:{argb:XL.sub}}; b.alignment={indent:1};
  ws.getRow(1).height = 30; ws.getRow(2).height = 18;
}
function xlHeader(ws, rowNo, cols){
  const row = ws.getRow(rowNo);
  cols.forEach((c,i)=>{ const cell = row.getCell(i+1); cell.value = c.h;
    cell.font = {bold:true,color:{argb:"FFFFFFFF"},size:10.5};
    cell.fill = {type:"pattern",pattern:"solid",fgColor:{argb:XL.band}};
    cell.alignment = {vertical:"middle",horizontal:c.num?"right":"left",wrapText:true,indent:c.num?0:1};
    cell.border = {bottom:{style:"medium",color:{argb:XL.amber}}};
    ws.getColumn(i+1).width = c.w||12;
    if(c.fmt) ws.getColumn(i+1).numFmt = c.fmt;
  });
  row.height = 32;
}
function xlBody(ws, from, to, cols){
  for(let r=from;r<=to;r++){
    const row = ws.getRow(r);
    cols.forEach((c,i)=>{ const cell=row.getCell(i+1);
      if(c.fmt) cell.numFmt = c.fmt;
      cell.alignment = {vertical:"middle",horizontal:c.num?"right":"left",indent:c.num?0:1};
      cell.border = {bottom:{style:"thin",color:{argb:XL.line}}};
      if((r-from)%2===1) cell.fill = {type:"pattern",pattern:"solid",fgColor:{argb:XL.stripe}};
      if(c.calc) cell.font = {color:{argb:"FF000000"}};
      if(c.input) cell.font = {color:{argb:"FF0A4FB3"}};
    });
    row.height = 19;
  }
}
function xlTotal(ws, rowNo, nCols, label, cells){
  const row = ws.getRow(rowNo);
  for(let i=1;i<=nCols;i++){ const cell=row.getCell(i);
    cell.fill = {type:"pattern",pattern:"solid",fgColor:{argb:XL.amberSoft}};
    cell.border = {top:{style:"medium",color:{argb:XL.amber}}};
    cell.font = {bold:true}; }
  row.getCell(1).value = label; row.getCell(1).alignment = {indent:1};
  Object.entries(cells).forEach(([col,v])=>{ const cell=row.getCell(+col); cell.value=v.v; if(v.fmt) cell.numFmt=v.fmt; cell.alignment={horizontal:"right"}; });
  row.height = 22;
}
const F = (formula, result) => ({formula, result: (result==null||(typeof result==="number"&&isNaN(result))) ? "" : result});
const blankIf = v => v==null||(typeof v==="number"&&isNaN(v)) ? null : v;
function colL(n){ let s=""; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; }
function xlStatusFmt(ws, ref, field){
  // Colors a status text column
  ws.addConditionalFormatting({ref, rules:[
    {type:"containsText",operator:"containsText",text:"Over",style:{font:{color:{argb:XL.badT},bold:true},fill:{type:"pattern",pattern:"solid",bgColor:{argb:XL.badF}}}},
    {type:"containsText",operator:"containsText",text:"Close",style:{font:{color:{argb:XL.warnT}},fill:{type:"pattern",pattern:"solid",bgColor:{argb:XL.warnF}}}},
    {type:"containsText",operator:"containsText",text:"On target",style:{font:{color:{argb:XL.okT}},fill:{type:"pattern",pattern:"solid",bgColor:{argb:XL.okF}}}}
  ]});
}

async function buildWorkbook(ci){
  await exLoadExcel();
  const D = exData(ci), {t,c,pv} = D, bar = exBarName();
  const wb = new ExcelJS.Workbook();
  wb.creator = "Tallypour"; wb.created = new Date(); wb.title = "Tallypour inventory report";
  const sm = wb.addWorksheet("Summary",{views:[{showGridLines:false}],properties:{tabColor:{argb:XL.ink}}}); // first tab, filled in last
  const subLine = [bar||null, c?("Count: "+exLongDate(c.date)):"No count yet", pv?("Compared with: "+exLongDate(pv.date)):null].filter(Boolean).join("   ·   ");
  const view = r => [{state:"frozen",ySplit:r,xSplit:0,showGridLines:false}];
  const setup = ws => { ws.pageSetup = {orientation:"landscape",fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.4,right:.4,top:.5,bottom:.5,header:.3,footer:.3}};
    ws.headerFooter = {oddFooter:"&L&8Tallypour&C&8&A&R&8Page &P of &N"}; };

  /* ----- Inventory ----- */
  const inv = wb.addWorksheet("Inventory",{views:view(4),properties:{tabColor:{argb:"FF30D158"}}}); setup(inv);
  const invCols = [{h:"Shelf",w:13},{h:"Category",w:13},{h:"Product",w:28},{h:"Size (ml)",w:9,num:1,fmt:XL.int},{h:"Bottle cost",w:11,num:1,fmt:XL.money},
    {h:"Cost / oz",w:10,num:1,fmt:XL.money},{h:"Sealed bottles",w:10,num:1,fmt:XL.int},{h:"Open bottle (oz)",w:11,num:1,fmt:XL.oz},
    {h:"On hand (oz)",w:11,num:1,fmt:XL.oz},{h:"On hand (bottles)",w:11,num:1,fmt:XL.btl},{h:"Value",w:12,num:1,fmt:XL.money},{h:"Counted",w:10}];
  xlTitle(inv,"Inventory on hand",subLine,invCols.length); xlHeader(inv,4,invCols);
  let r0 = 5;
  D.rows.forEach((x,i)=>{ const r=r0+i;
    inv.getRow(r).values = [x.shelf, x.cat, x.p.name, x.p.ml||0, x.p.cost||0,
      F(`IF(D${r}>0,E${r}/(D${r}/${BOZ}),0)`, x.k.costOz),
      blankIf(x.sealed), blankIf(x.open),
      F(`IF(L${r}="Yes",G${r}*D${r}/${BOZ}+H${r},"")`, x.counted?x.onOz:null),
      F(`IF(I${r}="","",I${r}/(D${r}/${BOZ}))`, x.counted?x.onBtl:null),
      F(`IF(I${r}="","",I${r}*F${r})`, x.counted?x.value:null),
      x.counted?"Yes":"Not counted"];
  });
  const invEnd = r0+D.rows.length-1;
  xlBody(inv,r0,invEnd,invCols);
  xlTotal(inv,invEnd+1,invCols.length,"Total",{7:{v:F(`SUM(G${r0}:G${invEnd})`,D.rows.reduce((s,x)=>s+(x.sealed||0),0)),fmt:XL.int},
    11:{v:F(`SUM(K${r0}:K${invEnd})`,t.value),fmt:XL.money}});
  if(D.rows.length) inv.autoFilter = {from:{row:4,column:1},to:{row:invEnd,column:invCols.length}};
  inv.addConditionalFormatting({ref:`L${r0}:L${invEnd}`,rules:[{type:"containsText",operator:"containsText",text:"Not",style:{font:{color:{argb:XL.warnT},italic:true}}}]});
  const INV_TOTAL = `Inventory!$K$${invEnd+1}`;

  /* ----- Pricing ----- */
  const pr = wb.addWorksheet("Pricing",{views:view(4),properties:{tabColor:{argb:XL.amber}}}); setup(pr);
  const prCols = [{h:"Category",w:13},{h:"Product",w:28},{h:"Size (ml)",w:9,num:1,fmt:XL.int},{h:"Bottle cost",w:11,num:1,fmt:XL.money},
    {h:"Pour (oz)",w:9,num:1,fmt:XL.oz},{h:"Cost / oz",w:10,num:1,fmt:XL.money},{h:"Cost / pour",w:11,num:1,fmt:XL.money},{h:"Menu price",w:11,num:1,fmt:XL.money},
    {h:"Pour cost %",w:11,num:1,fmt:XL.pct},{h:"Suggested price",w:12,num:1,fmt:XL.money},{h:"Profit / pour",w:11,num:1,fmt:XL.money},
    {h:"Pours / bottle",w:10,num:1,fmt:XL.oz1},{h:"Status",w:12}];
  xlTitle(pr,"Pour cost & pricing",subLine,prCols.length);
  pr.getCell("A3").value = "Target pour cost"; pr.getCell("A3").font={bold:true}; pr.getCell("A3").alignment={indent:1};
  pr.getCell("C3").value = D.target/100; pr.getCell("C3").numFmt = "0.0%"; pr.getCell("C3").font={bold:true,color:{argb:"FF0A4FB3"}};
  pr.getCell("C3").fill={type:"pattern",pattern:"solid",fgColor:{argb:XL.amberSoft}};
  pr.getCell("D3").value = "← change this to re-price every drink"; pr.getCell("D3").font={italic:true,size:9,color:{argb:XL.sub}};
  xlHeader(pr,4,prCols);
  D.rows.forEach((x,i)=>{ const r=r0+i;
    pr.getRow(r).values = [x.cat, x.p.name, x.p.ml||0, x.p.cost||0, x.k.pour,
      F(`IF(C${r}>0,D${r}/(C${r}/${BOZ}),0)`, x.k.costOz), F(`F${r}*E${r}`, x.k.costPour), blankIf(x.price),
      F(`IF(H${r}>0,G${r}/H${r},"")`, x.k.pc!=null?x.k.pc/100:null),
      F(`IF($C$3>0,G${r}/$C$3,"")`, x.k.suggested),
      F(`IF(H${r}>0,H${r}-G${r},"")`, x.k.profit),
      F(`IF(E${r}>0,(C${r}/${BOZ})/E${r},"")`, x.k.pours),
      F(`IF(I${r}="","No price",IF(I${r}<=$C$3,"On target",IF(I${r}<=$C$3+0.05,"Close","Over target")))`, x.status)];
  });
  const prEnd = r0+D.rows.length-1;
  xlBody(pr,r0,prEnd,prCols);
  if(D.rows.length){ pr.autoFilter = {from:{row:4,column:1},to:{row:prEnd,column:prCols.length}}; xlStatusFmt(pr,`M${r0}:M${prEnd}`);
    pr.addConditionalFormatting({ref:`I${r0}:I${prEnd}`,rules:[{type:"expression",formulae:[`AND(ISNUMBER(I${r0}),I${r0}>$C$3+0.05)`],style:{font:{color:{argb:XL.badT},bold:true}}}]}); }

  /* ----- Usage & Variance ----- */
  const us = wb.addWorksheet("Usage & Variance",{views:view(4),properties:{tabColor:{argb:"FFFF453A"}}}); setup(us);
  const usCols = [{h:"Category",w:13},{h:"Product",w:26},{h:"Bottle (oz)",w:9,num:1,fmt:XL.oz},{h:"Cost / oz",w:9,num:1,fmt:XL.money},
    {h:"Start (oz)",w:10,num:1,fmt:XL.oz},{h:"Received (bottles)",w:10,num:1,fmt:XL.int},{h:"End (oz)",w:10,num:1,fmt:XL.oz},
    {h:"Used (oz)",w:10,num:1,fmt:XL.oz},{h:"Used cost",w:11,num:1,fmt:XL.money},{h:"Pours sold",w:9,num:1,fmt:XL.int},{h:"Pour (oz)",w:8,num:1,fmt:XL.oz},
    {h:"Menu price",w:10,num:1,fmt:XL.money},{h:"Should have used (oz)",w:11,num:1,fmt:XL.oz},{h:"Over / under (oz)",w:11,num:1,fmt:XL.oz},
    {h:"Over / under ($)",w:11,num:1,fmt:XL.money},{h:"Over / under %",w:10,num:1,fmt:XL.pct},{h:"Sales",w:11,num:1,fmt:XL.money}];
  xlTitle(us,"Usage & variance",pv?subLine:"Needs two counts. Do another count to see usage.",usCols.length); xlHeader(us,4,usCols);
  D.usage.forEach((x,i)=>{ const r=r0+i, u=x.u;
    us.getRow(r).values = [x.cat, x.p.name, x.bOz, x.k.costOz, u.start, u.recv, blankIf(u.end),
      F(`IF(G${r}="","",E${r}+F${r}*C${r}-G${r})`, u.used), F(`IF(H${r}="","",H${r}*D${r})`, u.usedCost),
      blankIf(u.sold), x.k.pour, blankIf(x.price),
      F(`IF(J${r}="","",J${r}*K${r})`, u.exp), F(`IF(OR(H${r}="",M${r}=""),"",H${r}-M${r})`, u.v),
      F(`IF(N${r}="","",N${r}*D${r})`, u.vCost), F(`IF(OR(N${r}="",M${r}="",M${r}=0),"",N${r}/M${r})`, u.vPct),
      F(`IF(OR(J${r}="",L${r}=""),"",J${r}*L${r})`, u.rev)];
  });
  const usEnd = r0+Math.max(D.usage.length,1)-1;
  if(D.usage.length){ xlBody(us,r0,usEnd,usCols);
    xlTotal(us,usEnd+1,usCols.length,"Total",{9:{v:F(`SUM(I${r0}:I${usEnd})`,t.usedCost),fmt:XL.money},15:{v:F(`SUM(O${r0}:O${usEnd})`,t.vCost),fmt:XL.money},17:{v:F(`SUM(Q${r0}:Q${usEnd})`,t.rev),fmt:XL.money}});
    us.autoFilter = {from:{row:4,column:1},to:{row:usEnd,column:usCols.length}};
    us.addConditionalFormatting({ref:`N${r0}:P${usEnd}`,rules:[
      {type:"expression",formulae:[`AND(ISNUMBER($P${r0}),$P${r0}>0.05)`],style:{font:{color:{argb:XL.badT},bold:true},fill:{type:"pattern",pattern:"solid",bgColor:{argb:XL.badF}}}},
      {type:"expression",formulae:[`AND(ISNUMBER($P${r0}),$P${r0}<-0.05)`],style:{font:{color:{argb:XL.warnT}}}}]});
    const nr = usEnd+3;
    us.getCell(`A${nr}`).value = "Over / under = what left the bottles minus what was rung up. Positive means more was poured than sold (over-pouring, comps, spills or loss). Negative usually means a missed sale or a count error.";
    us.mergeCells(nr,1,nr,usCols.length); us.getCell(`A${nr}`).font={italic:true,size:9,color:{argb:XL.sub}}; us.getCell(`A${nr}`).alignment={wrapText:true,indent:1}; us.getRow(nr).height=28;
  } else { us.getCell("A5").value = "No usage yet. Usage compares two counts, so it appears after your second count."; us.mergeCells(5,1,5,usCols.length); us.getCell("A5").font={italic:true,color:{argb:XL.sub}}; }
  const US_TOTAL = D.usage.length ? usEnd+1 : null;

  /* ----- Order Sheet ----- */
  const od = wb.addWorksheet("Order Sheet",{views:view(4),properties:{tabColor:{argb:"FF0A84FF"}}}); setup(od);
  od.pageSetup.orientation = "portrait";
  const odCols = [{h:"Vendor",w:18},{h:"Category",w:13},{h:"Product",w:28},{h:"Size (ml)",w:9,num:1,fmt:XL.int},{h:"Par (bottles)",w:10,num:1,fmt:XL.oz1},
    {h:"On hand (bottles)",w:11,num:1,fmt:XL.btl},{h:"Order (bottles)",w:10,num:1,fmt:XL.int},{h:"Unit cost",w:10,num:1,fmt:XL.money},{h:"Est. cost",w:11,num:1,fmt:XL.money},{h:"✓",w:5}];
  xlTitle(od,"Order sheet",subLine,odCols.length);
  od.getCell("A3").value = t.hasPar ? "Order = par minus what's on hand, rounded up. Change a par here and the order updates." : "Set a par level on each bottle in Tallypour (Edit Bottle → Ordering) and this sheet fills in the order for you. You can also type pars in column E.";
  od.mergeCells(3,1,3,odCols.length); od.getCell("A3").font={italic:true,size:9,color:{argb:XL.sub}}; od.getCell("A3").alignment={wrapText:true,indent:1}; od.getRow(3).height=26;
  xlHeader(od,4,odCols);
  const ordRows = D.rows.slice().sort((a,b)=>(a.vendor||"~").localeCompare(b.vendor||"~")||a.cat.localeCompare(b.cat)||a.p.name.localeCompare(b.p.name));
  ordRows.forEach((x,i)=>{ const r=r0+i;
    od.getRow(r).values = [x.vendor||"—", x.cat, x.p.name, x.p.ml||0, blankIf(x.par), x.onBtl!=null?+x.onBtl.toFixed(2):0,
      F(`IF(E${r}="","",MAX(0,ROUNDUP(E${r}-F${r},0)))`, x.toOrder), x.p.cost||0, F(`IF(G${r}="","",G${r}*H${r})`, x.orderCost), ""];
  });
  const odEnd = r0+ordRows.length-1;
  xlBody(od,r0,odEnd,odCols);
  for(let r=r0;r<=odEnd;r++) od.getCell(`E${r}`).font = {color:{argb:"FF0A4FB3"}};
  xlTotal(od,odEnd+1,odCols.length,"Total",{7:{v:F(`SUM(G${r0}:G${odEnd})`,t.orderItems),fmt:XL.int},9:{v:F(`SUM(I${r0}:I${odEnd})`,t.orderCost),fmt:XL.money}});
  if(ordRows.length){ od.autoFilter = {from:{row:4,column:1},to:{row:odEnd,column:odCols.length}};
    od.addConditionalFormatting({ref:`G${r0}:G${odEnd}`,rules:[{type:"cellIs",operator:"greaterThan",formulae:[0],style:{font:{bold:true,color:{argb:"FF0A4FB3"}},fill:{type:"pattern",pattern:"solid",bgColor:{argb:"FFE1EEFF"}}}}]}); }
  const OD_TOTAL = `'Order Sheet'!$I$${odEnd+1}`;

  /* ----- Count History ----- */
  const hs = wb.addWorksheet("Count History",{views:view(4),properties:{tabColor:{argb:"FF8E8E93"}}}); setup(hs);
  const last = T.S.counts.slice(0, D.ci+1).slice(-12);
  const hCols = [{h:"Category",w:13},{h:"Product",w:28}].concat(last.map(cc=>({h:new Date(cc.date).toLocaleDateString("en-US",{month:"short",day:"numeric",year:"2-digit"}),w:11,num:1,fmt:XL.money})));
  xlTitle(hs,"Inventory value by count",(bar?bar+"   ·   ":"")+`Last ${last.length} count${last.length===1?"":"s"}, value on hand per bottle`,Math.max(hCols.length,4)); xlHeader(hs,4,hCols);
  D.rows.forEach((x,i)=>{ const r=r0+i;
    hs.getRow(r).values = [x.cat, x.p.name].concat(last.map(cc=>{ const o=totalOz(x.p,cc.items[x.p.id]); return o==null?null:+(o*x.k.costOz).toFixed(2); }));
  });
  const hEnd = r0+D.rows.length-1;
  xlBody(hs,r0,hEnd,hCols);
  const hTot = {}; last.forEach((cc,j)=>{ const col=colL(j+3); let s=0; D.rows.forEach(x=>{ const o=totalOz(x.p,cc.items[x.p.id]); if(o!=null) s+=o*x.k.costOz; }); hTot[j+3]={v:F(`SUM(${col}${r0}:${col}${hEnd})`,s),fmt:XL.money}; });
  xlTotal(hs,hEnd+1,hCols.length,"Total value",hTot);

  /* ----- Products ----- */
  const pl = wb.addWorksheet("Product List",{views:view(4),properties:{tabColor:{argb:"FF5E5CE6"}}}); setup(pl);
  const plCols = [{h:"Product",w:28},{h:"Category",w:13},{h:"Size (ml)",w:9,num:1,fmt:XL.int},{h:"Bottle cost",w:11,num:1,fmt:XL.money},{h:"Pour (oz)",w:9,num:1,fmt:XL.oz},
    {h:"Menu price",w:11,num:1,fmt:XL.money},{h:"Shelf",w:13},{h:"Vendor",w:18},{h:"Par (bottles)",w:10,num:1,fmt:XL.oz1},{h:"Barcode",w:16},{h:"Empty bottle (g)",w:11,num:1,fmt:XL.int}];
  xlTitle(pl,"Product list",(bar?bar+"   ·   ":"")+`${D.rows.length} bottles`,plCols.length); xlHeader(pl,4,plCols);
  D.rows.forEach((x,i)=>{ pl.getRow(r0+i).values = [x.p.name,x.cat,x.p.ml||0,x.p.cost||0,x.k.pour,blankIf(x.price),x.shelf,x.vendor,blankIf(x.par),x.p.barcode?String(x.p.barcode):"",x.p.tare>0?x.p.tare:null]; });
  xlBody(pl,r0,r0+D.rows.length-1,plCols);
  for(let r=r0;r<r0+D.rows.length;r++) pl.getCell(`J${r}`).numFmt="@";
  if(D.rows.length) pl.autoFilter = {from:{row:4,column:1},to:{row:r0+D.rows.length-1,column:plCols.length}};

  /* ----- Summary (first tab) ----- */
  sm.pageSetup = {orientation:"portrait",fitToPage:true,fitToWidth:1,fitToHeight:0};
  sm.headerFooter = {oddFooter:"&L&8Tallypour&R&8Page &P of &N"};
  [26,16,16,16,16,16].forEach((w,i)=>sm.getColumn(i+1).width=w);
  xlTitle(sm,(bar?bar+" — ":"")+"Inventory report",subLine,6);
  sm.getCell("A3").value = "Prepared "+exLongDate(Date.now())+" with Tallypour"; sm.getCell("A3").font={size:9,italic:true,color:{argb:XL.sub}}; sm.getCell("A3").alignment={indent:1};
  const kp = [
    ["Inventory value", F(INV_TOTAL, t.value), XL.money, "Everything on the shelf, at cost"],
    ["Bottles counted", `${t.counted} of ${t.products}`, null, t.counted<t.products?"Some bottles weren't counted":"Every bottle counted"],
    ["Average pour cost", t.avgPc!=null?t.avgPc/100:"", "0.0%", `Target ${D.target}%`],
    ["Drinks over target", t.over, XL.int, "See the Pricing tab for suggested prices"]];
  if(pv) kp.push(
    ["Liquor used (at cost)", US_TOTAL?F(`'Usage & Variance'!$I$${US_TOTAL}`,t.usedCost):0, XL.money, "Since "+exLongDate(pv.date)],
    ["Sales from logged pours", US_TOTAL?F(`'Usage & Variance'!$Q$${US_TOTAL}`,t.rev):0, XL.money, t.rev>0?"Pours sold × menu price":"Log pours sold to see this"],
    ["Actual pour cost", t.actualPc!=null?t.actualPc:"", "0.0%", t.idealPc!=null?`Ideal ${(t.idealPc*100).toFixed(1)}% if every pour matched a sale`:"Needs pours sold"],
    ["Over / under ($)", US_TOTAL?F(`'Usage & Variance'!$O$${US_TOTAL}`,t.vCost):0, XL.money, t.vCost>0?"Poured but not rung up":"Within what was sold"]);
  kp.push(["Estimated order", F(OD_TOTAL, t.orderCost), XL.money, t.hasPar?`${t.orderItems} bottles to reach par`:"Set par levels to build an order"]);
  sm.getCell("A5").value="KEY NUMBERS"; sm.getCell("A5").font={bold:true,size:9,color:{argb:XL.sub}}; sm.getCell("A5").alignment={indent:1};
  kp.forEach((k,i)=>{ const r=6+i;
    sm.getCell(`A${r}`).value=k[0]; sm.getCell(`A${r}`).font={bold:true}; sm.getCell(`A${r}`).alignment={indent:1,vertical:"middle"};
    sm.getCell(`B${r}`).value=k[1]; if(k[2]) sm.getCell(`B${r}`).numFmt=k[2]; sm.getCell(`B${r}`).font={bold:true,size:13}; sm.getCell(`B${r}`).alignment={horizontal:"right",vertical:"middle"};
    sm.mergeCells(r,3,r,6); sm.getCell(`C${r}`).value=k[3]; sm.getCell(`C${r}`).font={size:9.5,color:{argb:XL.sub}}; sm.getCell(`C${r}`).alignment={indent:1,vertical:"middle"};
    for(let col=1;col<=6;col++){ const cell=sm.getRow(r).getCell(col); cell.border={bottom:{style:"thin",color:{argb:XL.line}}}; if(i%2) cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:XL.stripe}}; }
    sm.getRow(r).height=24; });
  if(pv && t.vCost>0) sm.getCell(`B${6+kp.findIndex(k=>k[0]==="Over / under ($)")}`).font={bold:true,size:13,color:{argb:XL.badT}};
  // by category
  let cr = 6+kp.length+2;
  sm.getCell(`A${cr-1}`).value="BY CATEGORY"; sm.getCell(`A${cr-1}`).font={bold:true,size:9,color:{argb:XL.sub}}; sm.getCell(`A${cr-1}`).alignment={indent:1};
  const cCols=[{h:"Category"},{h:"Bottles",num:1},{h:"Value",num:1},{h:"Used cost",num:1},{h:"Over / under ($)",num:1},{h:"Sales",num:1}];
  const hr = sm.getRow(cr); cCols.forEach((c,i)=>{ const cell=hr.getCell(i+1); cell.value=c.h; cell.font={bold:true,color:{argb:"FFFFFFFF"},size:10}; cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:XL.band}}; cell.alignment={horizontal:c.num?"right":"left",indent:c.num?0:1,vertical:"middle"}; cell.border={bottom:{style:"medium",color:{argb:XL.amber}}}; }); hr.height=22;
  const cs = cr+1;
  D.cats.forEach((g,i)=>{ const r=cs+i, rng=(sh,col,end)=>`'${sh}'!$${col}$5:$${col}$${end}`;
    sm.getRow(r).values=[g.cat,
      F(`COUNTIF(${rng("Inventory","B",invEnd)},A${r})`,g.n),
      F(`SUMIF(${rng("Inventory","B",invEnd)},A${r},${rng("Inventory","K",invEnd)})`,g.value),
      D.usage.length?F(`SUMIF(${rng("Usage & Variance","A",usEnd)},A${r},${rng("Usage & Variance","I",usEnd)})`,g.usedCost):null,
      D.usage.length?F(`SUMIF(${rng("Usage & Variance","A",usEnd)},A${r},${rng("Usage & Variance","O",usEnd)})`,g.vCost):null,
      D.usage.length?F(`SUMIF(${rng("Usage & Variance","A",usEnd)},A${r},${rng("Usage & Variance","Q",usEnd)})`,g.rev):null];
    [2].forEach(c=>sm.getRow(r).getCell(c).numFmt=XL.int); [3,4,5,6].forEach(c=>sm.getRow(r).getCell(c).numFmt=XL.money);
    for(let col=1;col<=6;col++){ const cell=sm.getRow(r).getCell(col); cell.border={bottom:{style:"thin",color:{argb:XL.line}}}; cell.alignment={horizontal:col>1?"right":"left",indent:col>1?0:1}; if(i%2) cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:XL.stripe}}; }
  });
  const ce = cs+D.cats.length-1;
  if(D.cats.length){ const tr=ce+1, row=sm.getRow(tr);
    row.values=["Total",F(`SUM(B${cs}:B${ce})`,t.products),F(`SUM(C${cs}:C${ce})`,t.value),D.usage.length?F(`SUM(D${cs}:D${ce})`,t.usedCost):null,D.usage.length?F(`SUM(E${cs}:E${ce})`,t.vCost):null,D.usage.length?F(`SUM(F${cs}:F${ce})`,t.rev):null];
    for(let col=1;col<=6;col++){ const cell=row.getCell(col); cell.font={bold:true}; cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:XL.amberSoft}}; cell.border={top:{style:"medium",color:{argb:XL.amber}}}; cell.alignment={horizontal:col>1?"right":"left",indent:col>1?0:1}; if(col===2) cell.numFmt=XL.int; else if(col>2) cell.numFmt=XL.money; }
    cr = tr+3; } else cr = cs+2;
  // biggest variances
  const worst = D.usage.filter(x=>x.u.vCost!=null && x.u.vCost>0.005).sort((a,b)=>b.u.vCost-a.u.vCost).slice(0,8);
  if(worst.length){
    sm.getCell(`A${cr-1}`).value="BIGGEST OVER-POURS"; sm.getCell(`A${cr-1}`).font={bold:true,size:9,color:{argb:XL.sub}}; sm.getCell(`A${cr-1}`).alignment={indent:1};
    const wh=sm.getRow(cr); ["Product","Used (oz)","Sold (oz)","Over (oz)","Over %","Cost"].forEach((h,i)=>{ const cell=wh.getCell(i+1); cell.value=h; cell.font={bold:true,color:{argb:"FFFFFFFF"},size:10}; cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:XL.band}}; cell.alignment={horizontal:i?"right":"left",indent:i?0:1}; cell.border={bottom:{style:"medium",color:{argb:XL.amber}}}; }); wh.height=22;
    worst.forEach((x,i)=>{ const r=cr+1+i, row=sm.getRow(r);
      row.values=[x.p.name,x.u.used,x.u.exp,x.u.v,x.u.vPct,x.u.vCost];
      [2,3,4].forEach(c=>row.getCell(c).numFmt=XL.oz); row.getCell(5).numFmt=XL.pct; row.getCell(6).numFmt=XL.money; row.getCell(6).font={color:{argb:XL.badT},bold:true};
      for(let col=1;col<=6;col++){ const cell=row.getCell(col); cell.border={bottom:{style:"thin",color:{argb:XL.line}}}; if(col===1) cell.alignment={indent:1}; if(i%2) cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:XL.stripe}}; } });
  }
  wb.views = [{activeTab:0,firstSheet:0,visibility:"visible"}];
  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}

/* ---------- PDF report ---------- */
async function buildPdf(ci){
  await exLoadPdf();
  const D = exData(ci), {t,c,pv} = D, bar = exBarName();
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({unit:"pt",format:"letter"});
  const W = doc.internal.pageSize.getWidth(), M = 40;
  const ink=[28,28,30], amber=[255,159,10], sub=[110,110,115], red=[179,38,30], green=[30,122,52];
  const m = n => n==null||isNaN(n) ? "—" : (n<0?"-$":"$")+Math.abs(n).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
  const pct = n => n==null||isNaN(n) ? "—" : (n*100).toFixed(1)+"%";
  const o2 = n => n==null||isNaN(n) ? "—" : n.toFixed(1);
  // header band
  doc.setFillColor(...ink); doc.rect(0,0,W,78,"F");
  doc.setFillColor(...amber); doc.rect(0,78,W,3,"F");
  doc.setTextColor(255); doc.setFont("helvetica","bold"); doc.setFontSize(20); doc.text((bar||"Bar")+" — Inventory Report", M, 38);
  doc.setFont("helvetica","normal"); doc.setFontSize(9.5); doc.setTextColor(200);
  doc.text([c?("Count: "+exLongDate(c.date)):"No count yet", pv?("Compared with: "+exLongDate(pv.date)):null].filter(Boolean).join("     ·     "), M, 58);
  doc.setFontSize(9); doc.setTextColor(...amber); doc.text("TALLYPOUR", W-M, 38, {align:"right"});
  // KPI tiles
  const tiles = [["Inventory value",m(t.value),`${t.counted} of ${t.products} bottles counted`],
    pv?["Liquor used",m(t.usedCost),"at cost, since last count"]:["Average pour cost",t.avgPc!=null?t.avgPc.toFixed(1)+"%":"—",`target ${D.target}%`],
    pv?["Actual pour cost",pct(t.actualPc),t.idealPc!=null?`ideal ${pct(t.idealPc)}`:"log pours sold to see this"]:["Drinks over target",String(t.over),"see pricing below"],
    pv?["Over / under",m(t.vCost),t.vCost>0?"poured but not rung up":"within what was sold"]:["Estimated order",m(t.orderCost),t.hasPar?`${t.orderItems} bottles to par`:"set par levels"]];
  const tw = (W-2*M-3*10)/4; let y = 100;
  tiles.forEach((k,i)=>{ const x=M+i*(tw+10);
    doc.setFillColor(246,246,248); doc.roundedRect(x,y,tw,66,6,6,"F");
    doc.setFont("helvetica","normal"); doc.setFontSize(8); doc.setTextColor(...sub); doc.text(k[0].toUpperCase(), x+10, y+16);
    doc.setFont("helvetica","bold"); doc.setFontSize(17); doc.setTextColor(...(k[0]==="Over / under"&&t.vCost>0?red:ink)); doc.text(k[1], x+10, y+40);
    doc.setFont("helvetica","normal"); doc.setFontSize(7.5); doc.setTextColor(...sub); doc.text(doc.splitTextToSize(k[2],tw-20)[0], x+10, y+56); });
  y += 86;
  const head = {fillColor:[44,44,46],textColor:255,fontStyle:"bold",fontSize:8.5};
  const base = {theme:"plain",margin:{left:M,right:M,bottom:44,top:44},styles:{fontSize:8.5,cellPadding:{top:4,bottom:4,left:5,right:5},textColor:ink,lineColor:[217,217,222],lineWidth:{bottom:.5}},headStyles:head,alternateRowStyles:{fillColor:[246,246,248]}};
  const section = (title, note) => { if(y > doc.internal.pageSize.getHeight()-120){ doc.addPage(); y = 50; }
    doc.setFont("helvetica","bold"); doc.setFontSize(12); doc.setTextColor(...ink); doc.text(title, M, y);
    if(note){ doc.setFont("helvetica","normal"); doc.setFontSize(8); doc.setTextColor(...sub); doc.text(note, M, y+12); y+=6; }
    y += 10; };
  const after = () => { y = doc.lastAutoTable.finalY + 26; };
  const R = {halign:"right"};
  // categories
  section("By category");
  doc.autoTable(Object.assign({}, base, {startY:y,
    head:[["Category","Bottles","Value"].concat(pv?["Used cost","Over / under","Sales"]:[])],
    body:D.cats.map(g=>[g.cat,g.n,m(g.value)].concat(pv?[m(g.usedCost),m(g.vCost),m(g.rev)]:[])),
    foot:[["Total",t.products,m(t.value)].concat(pv?[m(t.usedCost),m(t.vCost),m(t.rev)]:[])],
    footStyles:{fillColor:[255,244,224],textColor:ink,fontStyle:"bold"},
    columnStyles:{1:R,2:R,3:R,4:R,5:R}, didParseCell:h=>{ if(h.column.index>0) h.cell.styles.halign="right"; }})); after();
  // over-pours
  if(pv){
    const worst = D.usage.filter(x=>x.u.vCost!=null && x.u.vCost>0.005).sort((a,b)=>b.u.vCost-a.u.vCost).slice(0,10);
    section("Biggest over-pours", worst.length?"Poured but not rung up, at cost. Check pour habits, comps and spills.":"Nothing poured beyond what was sold. Nice.");
    if(worst.length){ doc.autoTable(Object.assign({}, base, {startY:y,
      head:[["Product","Used (oz)","Sold (oz)","Over (oz)","Over %","Cost"]],
      body:worst.map(x=>[x.p.name,o2(x.u.used),o2(x.u.exp),o2(x.u.v),pct(x.u.vPct),m(x.u.vCost)]),
      didParseCell:h=>{ if(h.column.index>0) h.cell.styles.halign="right"; if(h.section==="body"&&h.column.index===5){ h.cell.styles.textColor=red; h.cell.styles.fontStyle="bold"; } }})); after(); }
    else y += 8;
  }
  // pricing flags
  const over = D.rows.filter(x=>x.status==="Over target"||x.status==="Close").sort((a,b)=>b.k.pc-a.k.pc);
  section("Drinks to re-price", over.length?`Pour cost above the ${D.target}% target. Suggested prices are rounded up to the nearest quarter.`:`Every priced drink is at or under the ${D.target}% target.`);
  if(over.length){ doc.autoTable(Object.assign({}, base, {startY:y,
    head:[["Product","Cost / pour","Menu price","Pour cost","Suggested price"]],
    body:over.map(x=>[x.p.name,m(x.k.costPour),m(x.price),x.k.pc.toFixed(1)+"%",m(Math.ceil(x.k.suggested*4)/4)]),
    didParseCell:h=>{ if(h.column.index>0) h.cell.styles.halign="right"; if(h.section==="body"&&h.column.index===3){ const v=parseFloat(h.cell.raw); h.cell.styles.textColor = v>D.target+5?red:[154,91,0]; h.cell.styles.fontStyle="bold"; } }})); after(); }
  else y += 8;
  // order
  if(t.hasPar){
    const ord = D.rows.filter(x=>x.toOrder>0).sort((a,b)=>(a.vendor||"~").localeCompare(b.vendor||"~")||a.p.name.localeCompare(b.p.name));
    section("Order to par", ord.length?`${t.orderItems} bottles · about ${m(t.orderCost)}`:"Everything is at or above par.");
    if(ord.length){ doc.autoTable(Object.assign({}, base, {startY:y,
      head:[["Vendor","Product","Par","On hand","Order","Est. cost"]],
      body:ord.map(x=>[x.vendor||"—",x.p.name,x.par,x.onBtl!=null?x.onBtl.toFixed(1):"0",x.toOrder,m(x.orderCost)]),
      didParseCell:h=>{ if(h.column.index>1) h.cell.styles.halign="right"; }})); after(); }
  }
  // full inventory
  doc.addPage(); y = 50;
  section("Full inventory");
  doc.autoTable(Object.assign({}, base, {startY:y,
    head:[["Product","Category","Size","Sealed","Open (oz)","On hand (oz)","Cost / oz","Value"]],
    body:D.rows.map(x=>[x.p.name,x.cat,(x.p.ml||0)+" ml",x.counted?x.sealed:"—",x.counted?o2(x.open):"—",x.counted?o2(x.onOz):"not counted",m(x.k.costOz),x.counted?m(x.value):"—"]),
    foot:[["Total","","",String(D.rows.reduce((s,x)=>s+(x.sealed||0),0)),"","","",m(t.value)]],
    footStyles:{fillColor:[255,244,224],textColor:ink,fontStyle:"bold"},
    didParseCell:h=>{ if(h.column.index>1) h.cell.styles.halign="right"; if(h.section==="body"&&h.cell.raw==="not counted"){ h.cell.styles.textColor=[154,91,0]; h.cell.styles.fontStyle="italic"; } }}));
  // footer on every page
  const n = doc.internal.getNumberOfPages(), H = doc.internal.pageSize.getHeight();
  const stamp = "Generated with Tallypour · "+exLongDate(Date.now());
  for(let i=1;i<=n;i++){ doc.setPage(i); doc.setFont("helvetica","normal"); doc.setFontSize(7.5); doc.setTextColor(...sub); doc.text(stamp, M, H-20); doc.text(`Page ${i} of ${n}`, W-M, H-20, {align:"right"}); }
  return doc.output("blob");
}

/* ---------- CSV ---------- */
function exCsv(rows){
  const q = v => { if(v==null||(typeof v==="number"&&isNaN(v))) return ""; v = typeof v==="number" ? String(+v.toFixed(4)) : String(v); return /[",\n\r]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v; };
  return new Blob(["﻿"+rows.map(r=>r.map(q).join(",")).join("\r\n")],{type:"text/csv;charset=utf-8"});
}
const r2 = n => n==null||isNaN(n) ? null : Math.round(n*100)/100;
const EX_CSV = {
  inventory:{name:"Inventory", build:D=>[["Shelf","Category","Product","Size ml","Bottle cost","Cost per oz","Sealed bottles","Open oz","On hand oz","On hand bottles","Value","Counted"]]
    .concat(D.rows.map(x=>[x.shelf,x.cat,x.p.name,x.p.ml,r2(x.p.cost||0),+x.k.costOz.toFixed(4),x.sealed,r2(x.open),r2(x.onOz),r2(x.onBtl),r2(x.value),x.counted?"Yes":"No"]))},
  usage:{name:"Usage", build:D=>[["Category","Product","Start oz","Received bottles","End oz","Used oz","Used cost","Pours sold","Should have used oz","Over under oz","Over under $","Over under %","Sales"]]
    .concat(D.usage.map(x=>{ const u=x.u; return [x.cat,x.p.name,r2(u.start),u.recv,r2(u.end),r2(u.used),r2(u.usedCost),u.sold,r2(u.exp),r2(u.v),r2(u.vCost),u.vPct==null?null:+(u.vPct*100).toFixed(1),r2(u.rev)]; }))},
  pricing:{name:"Pricing", build:D=>[["Category","Product","Size ml","Bottle cost","Pour oz","Cost per pour","Menu price","Pour cost %","Suggested price","Profit per pour","Status"]]
    .concat(D.rows.map(x=>[x.cat,x.p.name,x.p.ml,r2(x.p.cost||0),x.k.pour,r2(x.k.costPour),x.price,x.k.pc==null?null:+x.k.pc.toFixed(1),r2(x.k.suggested),r2(x.k.profit),x.status]))},
  order:{name:"Order", build:D=>[["Vendor","Category","Product","Size ml","Par bottles","On hand bottles","Order bottles","Unit cost","Est cost"]]
    .concat(D.rows.slice().sort((a,b)=>(a.vendor||"~").localeCompare(b.vendor||"~")||a.p.name.localeCompare(b.p.name))
      .map(x=>[x.vendor,x.cat,x.p.name,x.p.ml,x.par,r2(x.onBtl||0),x.toOrder,r2(x.p.cost||0),r2(x.orderCost)]))},
  products:{name:"Products", build:D=>[["Product","Category","Size ml","Bottle cost","Pour oz","Menu price","Shelf","Vendor","Par bottles","Barcode"]]
    .concat(D.rows.map(x=>[x.p.name,x.cat,x.p.ml,r2(x.p.cost||0),x.k.pour,x.price,x.shelf,x.vendor,x.par,x.p.barcode||""]))}
};

/* ---------- UI ---------- */
let exCi = null, exBusy = false;
function exportSheet(){
  if(exCi==null || exCi>=T.S.counts.length) exCi = T.S.counts.length-1;
  exLoadExcel().catch(()=>{}); exLoadPdf().catch(()=>{}); // warm up so sharing keeps the tap
  const opts = T.S.counts.map((cc,i)=>({i,l:fmtDate(cc.date)})).reverse();
  const row = (a,v,icon,col,title,sub) => `<div class="row tap hi" data-a="${a}" data-v="${v||""}">${ic(icon,col)}<div class="main"><div class="name">${title}</div><div class="meta" style="white-space:normal">${sub}</div></div>${I.chev}</div>`;
  const noProducts = !T.S.products.length;
  openSheet(`<div class="sh-head"><button class="l" data-a="closeSheet">Done</button><div class="t">Export</div><button class="r" style="visibility:hidden">Done</button></div>
  <div class="sh-body">
    ${noProducts?`<div class="empty">Add bottles first, then come back to export.</div>`:`
    <div class="sec-h">Report For</div><div class="group">
      <div class="row hi">${ic("clock")}<label>Count</label>${opts.length?`<select id="exCount">${opts.map(o=>`<option value="${o.i}" ${o.i===exCi?"selected":""}>${o.l}${o.i===T.S.counts.length-1?" (latest)":""}</option>`).join("")}</select>`:`<div class="val">No counts yet</div>`}</div>
      <div class="row hi">${ic("tag")}<label>Bar name</label><input data-t="bar" value="${esc(T.S.settings.bar||"")}" placeholder="Shown on reports" style="color:var(--label)"></div>
    </div>
    <div class="sec-f">${exCi>0?"Usage compares this count with the one before it.":"Usage and over/under appear once you have two counts."}</div>
    <div class="sec-h">Full Reports</div><div class="group">
      ${row("exXlsx","","grid","#30d158","Excel Workbook","Summary, inventory, usage &amp; variance, pricing, order sheet, count history and product list. Live formulas.")}
      ${row("exPdf","","export","#ff453a","PDF Report","A clean summary for the owner: key numbers, over-pours, drinks to re-price, order and full inventory.")}
    </div>
    <div class="sec-h">Single Files (CSV)</div><div class="group">
      ${row("exCsv","inventory","box","#ac8e68","Inventory","For accounting or month-end inventory value")}
      ${row("exCsv","usage","chart","#ff453a","Usage &amp; Variance","Used, sold and over/under per bottle")}
      ${row("exCsv","pricing","percent","#bf5af2","Pricing","Cost per pour, pour cost % and suggested prices")}
      ${row("exCsv","order","truck","#0a84ff","Order Sheet","What to order to get back to par, by vendor")}
      ${row("exCsv","products","bottle","#64d2ff","Product List","All bottles, for importing into other tools")}
    </div>
    <div class="sec-f">Files open in Excel, Google Sheets and Numbers. On a phone you'll get the share menu, so you can email, text, AirDrop or save to Files.</div>`}
  </div>`);
}
async function exRun(kind, arg){
  if(exBusy) return; exBusy = true;
  const sel = $("#exCount"); if(sel) exCi = +sel.value;
  const ci = T.S.counts.length ? exCi : null, base = exFileBase(ci==null?0:ci);
  try{
    let blob, name;
    if(kind==="xlsx"){ toast("Building workbook…"); blob = await buildWorkbook(ci); name = base+".xlsx"; }
    else if(kind==="pdf"){ toast("Building report…"); blob = await buildPdf(ci); name = base+".pdf"; }
    else { const def = EX_CSV[arg]; const D = exData(ci); if(arg==="usage" && !D.usage.length){ toast("Usage needs two counts"); return; } blob = exCsv(def.build(D)); name = base+"_"+def.name+".csv"; }
    const how = await exDeliver(blob, name);
    if(how==="downloaded") toast(kind==="xlsx"?"Workbook saved":kind==="pdf"?"Report saved":"CSV saved");
  }catch(e){ console.error(e); toast("Couldn't build that file. Check your connection and try again."); }
  finally{ exBusy = false; }
}

return {exportSheet, exRun, buildWorkbook, buildPdf, exData};
};
