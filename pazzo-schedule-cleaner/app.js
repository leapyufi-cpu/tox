"use strict";
const $ = id => document.getElementById(id);
const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
let inputFile = null, sourceZip = null, sheets = [], shared = [], analysis = null;
const child = (el, name) => [...(el?.children || [])].find(e => e.localName === name) || null;
const children = (el, name) => [...(el?.children || [])].filter(e => e.localName === name);
const textOf = el => el?.textContent || "";
const xml = str => {
  const doc = new DOMParser().parseFromString(str, "application/xml");
  if (doc.querySelector("parsererror")) throw new Error("無法讀取 Excel XML，請確認檔案未損毀。");
  return doc;
};
const serialize = doc => new XMLSerializer().serializeToString(doc);
const colNum = letters => [...letters].reduce((n,ch) => n * 26 + ch.charCodeAt(0) - 64,0);
const colLetters = n => {let x="";while(n){x=String.fromCharCode(65+(n-1)%26)+x;n=Math.floor((n-1)/26);}return x;};
const parseAddr = addr => {
  const m = /^(\$?)([A-Z]+)(\$?)(\d+)$/.exec(addr || "");
  return m ? {col:colNum(m[2]),row:+m[4]} : null;
};
const newText = (doc, value) => {
  const is=doc.createElementNS(NS,"is"), t=doc.createElementNS(NS,"t");
  if (/^\s|\s$/.test(value)) t.setAttributeNS("http://www.w3.org/XML/1998/namespace","xml:space","preserve");
  t.textContent=value;is.append(t);return is;
};
const dateString = serial => {
  const d=new Date(Date.UTC(1899,11,30)+Math.floor(+serial)*86400000);
  return Number.isFinite(d.getTime()) ? (d.getUTCMonth()+1)+"月"+d.getUTCDate()+"日" : String(serial);
};
const resolvePath = (base,target) => {
  if (target.startsWith("/")) return target.slice(1);
  const p=base.split("/");p.pop();
  for (const part of target.split("/")){if(part==="..")p.pop();else if(part&&part!==".")p.push(part);}
  return p.join("/");
};
async function getXML(zip,path){const file=zip.file(path);if(!file)throw new Error("缺少 Excel 內部檔案："+path);return xml(await file.async("string"));}
async function mapSheets(zip){
  const book=await getXML(zip,"xl/workbook.xml");
  const links=await getXML(zip,"xl/_rels/workbook.xml.rels");
  const rels=new Map(children(links.documentElement,"Relationship").map(e=>[e.getAttribute("Id"),e.getAttribute("Target")]));
  return children(child(book.documentElement,"sheets"),"sheet").map(s=>({
    name:s.getAttribute("name"),rid:s.getAttributeNS(REL,"id"),
    path:resolvePath("xl/workbook.xml",rels.get(s.getAttributeNS(REL,"id"))||"")
  })).filter(s=>zip.file(s.path));
}
async function getShared(zip){
  const f=zip.file("xl/sharedStrings.xml");
  if(!f)return [];
  const d=xml(await f.async("string"));
  return children(d.documentElement,"si").map(si=>textOf(si));
}
function cellValue(cell){
  if(!cell)return "";
  const t=cell.getAttribute("t");
  if(t==="inlineStr")return textOf(child(cell,"is"));
  const v=textOf(child(cell,"v"));
  return t==="s" ? (shared[+v]??"") : v;
}
function getRows(sheetDoc){
  const data=child(sheetDoc.documentElement,"sheetData");
  const result=[];
  for(const row of children(data,"row")){
    const index=+row.getAttribute("r");
    const cells=new Map();
    for(const cell of children(row,"c")){
      const p=parseAddr(cell.getAttribute("r"));if(p)cells.set(p.col,cell);
    }
    result.push({index,row,cells,label:cellValue(cells.get(1)).trim()});
  }
  return result;
}
function groupLayouts(rows){
  const starts=rows.filter(r=>r.label==="商品貨號");
  if(!starts.length)throw new Error("找不到「商品貨號」區段，不符合目前支援的 PAZZO 範本。");
  const result=[];
  for(const s of starts){
    const start=s.index, next=starts.find(r=>r.index>start)?.index ?? Number.MAX_SAFE_INTEGER;
    const findLabel=label=>rows.find(r=>r.index>=start&&r.index<next&&r.label===label)?.index;
    const r={code:start,image:findLabel("商品圖片"),activity:findLabel("活動組合 / 活動價 / 毛利"),price:findLabel("成本 / 定價 / 毛利"),color:findLabel("顏色 / 件數"),size:findLabel("尺寸"),name:findLabel("商品名稱"),description:findLabel("商品描述")};
    if(Object.values(r).some(x=>x===undefined)||!(r.code<r.image&&r.image<r.activity&&r.activity<r.price&&r.price<r.color&&r.color<r.size&&r.size<r.name&&r.name<r.description)){
      throw new Error("第 "+start+" 列的商品區段不完整或順序不同，已停止處理以避免內容錯置。");
    }
    const byIndex=new Map(rows.map(x=>[x.index,x]));
    const productCols=[...byIndex.get(start).cells.entries()].filter(([col,c])=>col>=3&&(col-3)%3===0&&/^P[A-Z0-9]{6,}$/.test(cellValue(c).trim())).map(([col])=>col);
    if(!productCols.length && result.length) continue;
    // Check visible color and OFFLINE quantity across ALL products. Do not count ONLINE quantities or formatting-only cells.
    // Keep a partially populated row (orange-marked case); drop a row only when all product groups are empty (blue-marked case).
    const colorRows=Array.from({length:r.size-r.color},(_,j)=>r.color+j).filter(i=>{
      const row=byIndex.get(i);
      return row && [...row.cells.entries()].some(([c,node])=>
        c>=2 && c<=31 && ((c-2)%3===0 || (c-2)%3===2) && !!cellValue(node).trim()
      );
    });
    if(!colorRows.length)colorRows.push(r.color); // Preserve the label when a group has no colors yet.
    const highlight=rows.find(x=>x.index> (result.at(-1)?.description||0) && x.index<start &&
      !/關鍵字\s*[\/／]?\s*備註/.test(x.label) &&
      [...x.cells.values()].some(c=>/A\+\s*主推|GIFT|贈品|滿額/i.test(cellValue(c))));
    result.push({...r,colorRows,productCols,highlight:highlight?.index});
  }
  return result;
}
async function drawingStats(zip,sheet){
  const relPath=sheet.path.replace("worksheets/","worksheets/_rels/")+".rels";
  const file=zip.file(relPath);if(!file)return {count:0,drawPath:null,drawRelsPath:null,rels:null};
  const rels=xml(await file.async("string"));
  const drawings=children(rels.documentElement,"Relationship").filter(e=>(e.getAttribute("Type")||"").endsWith("/drawing"));
  if(!drawings.length)return {count:0,drawPath:null,drawRelsPath:null,rels:null};
  const drawPath=resolvePath(sheet.path,drawings[0].getAttribute("Target"));
  const drawDoc=await getXML(zip,drawPath);
  const count=[...drawDoc.documentElement.children].filter(a=>a.localName.endsWith("Anchor")&&a.querySelector("pic")).length;
  return {count,drawPath,drawDoc,drawRelsPath:drawPath.replace("drawings/","drawings/_rels/")+".rels",sheetDrawRelId:drawings[0].getAttribute("Id")};
}
function showStatus(message,error=false){
  const box=$("output");box.classList.remove("hide");box.classList.toggle("error",error);box.replaceChildren();
  const strong=document.createElement("strong");strong.textContent=error?"無法整理":"處理結果";box.append(strong);
  const p=document.createElement("div");p.textContent=message;box.append(p);
}
async function inspectFile(file){
  inputFile=file;sourceZip=null;analysis=null;
  $("run").disabled=true;$("fileBox").classList.remove("hide");$("fileName").textContent=file.name;
  $("fileSize").textContent=(file.size/1024/1024).toFixed(1)+" MB";
  $("output").classList.add("hide");$("sheetField").classList.add("hide");$("stats").classList.add("hide");
  if(!/\.xlsx$/i.test(file.name)){showStatus("請選擇 .xlsx 格式。",true);return;}
  if(!window.JSZip){showStatus("壓縮檔處理程式尚未載入，請刷新頁面並確認網路連線。",true);return;}
  try{
    sourceZip=await JSZip.loadAsync(file);
    shared=await getShared(sourceZip);
    sheets=await mapSheets(sourceZip);
    if(!sheets.length)throw new Error("找不到可處理的工作表。");
    $("sheetSelect").replaceChildren();
    for(let i=0;i<sheets.length;i++){
      const opt=document.createElement("option");opt.value=String(i);opt.textContent=sheets[i].name;
      $("sheetSelect").append(opt);
    }
    const preferred=sheets.findIndex(x=>!/總表|彙整|summary/i.test(x.name));
    $("sheetSelect").value=String(Math.max(preferred,0));
    $("sheetField").classList.toggle("hide",sheets.length<2);
    await refreshStats();
  }catch(e){showStatus(e.message||"無法讀取檔案。",true);}
}
async function refreshStats(){
  try{
    const sheet=sheets[+$("sheetSelect").value];
    const doc=await getXML(sourceZip,sheet.path),rows=getRows(doc),groups=groupLayouts(rows);
    const draw=await drawingStats(sourceZip,sheet);
    const productCount=groups.reduce((sum,g)=>sum+g.productCols.length,0);
    analysis={sheet,rows,groups,draw,doc,productCount};
    $("groupCount").textContent=groups.length;
    $("productCount").textContent=productCount;
    $("imageCount").textContent=draw.count;
    $("stats").classList.remove("hide");
    $("run").disabled=false;
  }catch(e){analysis=null;$("run").disabled=true;showStatus(e.message,true);}
}
function getKeepRows(analysis){
  const {rows,groups}=analysis;
  const source=new Map(rows.map(r=>[r.index,r]));
  const keep=[1,2,3,4];
  const specials=new Set();
  for(const g of groups){
    if(g.highlight && ( $("keepHighlights").checked || $("keepGifts").checked)){
      const row=source.get(g.highlight);
      const matching=[...row.cells.values()].some(c=>{
        const v=cellValue(c);
        return ($("keepHighlights").checked && /A\+\s*主推/.test(v)) || ($("keepGifts").checked && /GIFT|贈品|滿額/i.test(v));
      });
      if(matching && !/關鍵字\s*[\/／]?\s*備註/.test(row.label)){
        keep.push(g.highlight);specials.add(g.highlight);
      }
    }
    keep.push(g.code,g.image,g.activity,g.price);
    keep.push(...g.colorRows);
    keep.push(g.size,g.name,g.description);
  }
  // Force-remove the whole keyword/notes row, even when it contains GIFT or other marketing terms.
  const forbidden=new Set(rows.filter(r=>/關鍵字\s*[\/／]?\s*備註/.test(r.label)).map(r=>r.index));
  return {keep:[...new Set(keep)].filter(i=>!forbidden.has(i)).sort((a,b)=>a-b),specials};
}
function filteredCell(rowType, col, rawValue, rowNumber){
  if(rowType==="header"){
    if(rowNumber===1 && col===2) return rawValue;
    if(rowNumber===1 && col===1 && rawValue && !/^\d/.test(rawValue))return rawValue;
    return null;
  }
  if(col===1)return rawValue;
  if(rowType==="highlight"){
    return (/A\+\s*主推/.test(rawValue)&&$("keepHighlights").checked)||(/GIFT|贈品|滿額/i.test(rawValue)&&$("keepGifts").checked)?rawValue:null;
  }
  if(rowType==="code")return col>=3 && (col-3)%3===0 && /^P[A-Z0-9]{6,}$/.test(rawValue.trim()) ? rawValue : null;
  if(rowType==="price")return col>=3 && (col-3)%3===0 ? rawValue : null;
  if(rowType==="activity")return col>=2 && (col-2)%3!==2 && /^\s*-?[\d,]+(?:\.\d+)?\s*$/.test(rawValue) ? rawValue : null;
  if(rowType==="color")return col>=2 && (col-2)%3!==1 ? rawValue : null;
  if(rowType==="image")return rawValue==="*" ? rawValue : null;
  return rawValue;
}
function rewriteCell(cell, value, asString=false){
  for(const c of [...cell.children]) if(["f","v","is"].includes(c.localName))cell.removeChild(c);
  cell.removeAttribute("cm");cell.removeAttribute("vm");
  if(value===null || value==="") {cell.removeAttribute("t");return;}
  if(!asString && /^-?\d+(?:\.\d+)?$/.test(String(value))){
    cell.removeAttribute("t");const v=cell.ownerDocument.createElementNS(NS,"v");v.textContent=String(value);cell.append(v);
  }else{
    cell.setAttribute("t","inlineStr");cell.append(newText(cell.ownerDocument,String(value)));
  }
}
function transformSheet(analysis){
  const doc=analysis.doc, root=doc.documentElement, {keep,specials}=getKeepRows(analysis);
  const rank=new Map(keep.map((old,i)=>[old,i+1]));
  const types=new Map();
  for(const g of analysis.groups){
    if(specials.has(g.highlight))types.set(g.highlight,"highlight");
    for(const k of ["code","image","activity","price","size","name","description"])types.set(g[k],k);
    for(const i of g.colorRows)types.set(i,"color");
  }
  const sheetData=child(root,"sheetData");
  const original=[...sheetData.children].filter(e=>e.localName==="row");
  for(const row of original) {
    const old=+row.getAttribute("r"),nr=rank.get(old);
    if(!nr){sheetData.removeChild(row);continue;}
    row.setAttribute("r",String(nr));
    const type=types.get(old)||"header";
    for(const c of [...row.children].filter(e=>e.localName==="c")){
      const pos=parseAddr(c.getAttribute("r"));
      if(!pos){row.removeChild(c);continue;}
      if(pos.col>31){row.removeChild(c);continue;} // AF, AG totals are internal
      const value=cellValue(c);
      const safe=filteredCell(type,pos.col,value,old);
      if(safe===null||safe===""){
        // Preserve empty worksheet cells (styles/fills), then add consistent borders in the styling pass.
        // Do not leave sensitive values or formulas inside them.
        if(["code","activity","price","color","size","name","description"].includes(type)){
          rewriteCell(c,null);
          c.setAttribute("r",colLetters(pos.col)+nr);
        }else row.removeChild(c);
        continue;
      }
      let result=safe,asString=false;
      if(type==="header"&&old===1&&pos.col===2&&/^\d{5}$/.test(safe)){result=dateString(safe);asString=true;}
      if(type==="header"&&old===1&&pos.col===1){asString=true;}
      if(type==="code"||type==="size"||type==="name"||type==="description"||type==="color"||type==="highlight"||pos.col===1)asString=true;
      rewriteCell(c,result,asString);
      c.setAttribute("r",colLetters(pos.col)+nr);
    }
    if(type==="header"&&old===1&&row.children.length===0){
      // keep row format only
    }
    // Remove row metadata that may reveal source row numbering
    row.removeAttribute("spans");
  }
  // Clear non-essential nodes that could include sensitive internal content.
  for(const el of [...root.children]){
    if(["conditionalFormatting","dataValidations","hyperlinks","autoFilter","sortState","extLst","legacyDrawing","oleObjects","controls","tableParts","smartTags","customSheetViews","protectedRanges","sheetProtection","pageSetup","printOptions","rowBreaks","colBreaks","ignoredErrors"].includes(el.localName))root.removeChild(el);
  }
  let mergeSection=child(root,"mergeCells");
  if(mergeSection){
    for(const m of [...mergeSection.children]){
      const ref=m.getAttribute("ref")||"", [a,b]=ref.split(":");
      const p=parseAddr(a),q=parseAddr(b||a);
      if(!p||!q){mergeSection.removeChild(m);continue;}
      const start=rank.get(p.row);
      // Color labels (e.g. A21:A26) are vertically merged. Shorten the merge to the
      // retained rows rather than dropping its border when blank color rows are deleted.
      const mappedRows=[...rank.entries()].filter(([old])=>old>=p.row&&old<=q.row).map(([,now])=>now);
      const end=mappedRows.at(-1);
      if(!start||!end||p.col>31||q.col>31){
        mergeSection.removeChild(m);continue;
      }
      if(start===end && p.col===q.col){mergeSection.removeChild(m);continue;}
      m.setAttribute("ref",colLetters(p.col)+start+":"+colLetters(q.col)+end);
    }
    mergeSection.setAttribute("count",String(mergeSection.children.length));
    if(!mergeSection.children.length)root.removeChild(mergeSection);
  }
  const dim=child(root,"dimension");
  if(dim)dim.setAttribute("ref","A1:AE"+keep.length);
  const cols=child(root,"cols");
  if(cols)for(const col of [...cols.children]){
    if(+col.getAttribute("min")>31)cols.removeChild(col);
    else if(+col.getAttribute("max")>31)col.setAttribute("max","31");
  }
  const rowTypes=new Map([...types.entries()].filter(([old])=>rank.has(old)).map(([old,type])=>[rank.get(old),type]));
  return {doc,rank,rowTypes,rowCount:keep.length};
}
function transformDrawing(draw,rank){
  if(!draw.drawDoc)return {doc:null,images:0};
  const doc=draw.drawDoc,root=doc.documentElement;
  let kept=0;
  for(const a of [...root.children]){
    if(!a.localName.endsWith("Anchor")){root.removeChild(a);continue;}
    const from=child(a,"from"),to=child(a,"to");
    const fr=child(from,"row"),tr=child(to,"row");
    const old=fr?+fr.textContent+1:0;
    if(!rank.has(old)||!a.querySelector("pic")){root.removeChild(a);continue;}
    fr.textContent=String(rank.get(old)-1);
    if(tr){const oldTo=+tr.textContent+1;if(!rank.has(oldTo)){root.removeChild(a);continue;}tr.textContent=String(rank.get(oldTo)-1);}
    kept++;
  }
  return {doc,images:kept};
}

function applyCleanStyles(sheetDoc, stylesDoc, groups, rank){
  // Preserve existing cell layout, merges, fills, and interior borders.
  // Only complete missing OUTER edges on each product card (B:D, E:G, etc).
  const root=stylesDoc.documentElement;
  const fonts=child(root,"fonts");
  if(fonts)for(const font of children(fonts,"font")){
    let name=child(font,"name");
    if(!name){name=stylesDoc.createElementNS(NS,"name");font.append(name);}
    name.setAttribute("val","微軟正黑體");
  }
  const borders=child(root,"borders"),xfs=child(root,"cellXfs");
  if(!borders||!xfs)return;
  const originalXfs=children(xfs,"xf"),originalBorders=children(borders,"border");
  const styleCache=new Map();
  const edgeName=["left","right","top","bottom"];
  function outlineStyle(styleId,edges){
    const requested=edgeName.filter(side=>edges.has(side));
    if(!requested.length)return styleId;
    const key=styleId+"|"+requested.join(",");
    if(styleCache.has(key))return styleCache.get(key);
    const xf=originalXfs[styleId]||originalXfs[0];
    if(!xf)return styleId;
    const origBorder=originalBorders[+(xf.getAttribute("borderId")||0)]||originalBorders[0];
    if(!origBorder)return styleId;
    const border=origBorder.cloneNode(true);
    let changed=false;
    for(const side of requested){
      let el=child(border,side);
      if(!el){el=stylesDoc.createElementNS(NS,side);border.append(el);}
      // Existing lines are never restyled.
      if(el.getAttribute("style"))continue;
      el.setAttribute("style","thin");
      const color=stylesDoc.createElementNS(NS,"color");
      color.setAttribute("rgb","FF8E9398");
      el.append(color);
      changed=true;
    }
    if(!changed){styleCache.set(key,styleId);return styleId;}
    const borderId=borders.children.length;
    borders.append(border);
    const copy=xf.cloneNode(true);
    copy.setAttribute("borderId",String(borderId));
    copy.setAttribute("applyBorder","1");
    const newStyle=xfs.children.length;
    xfs.append(copy);
    styleCache.set(key,newStyle);
    return newStyle;
  }
  const edgesByRow=new Map();
  const addEdge=(row,col,edge)=>{
    if(!row||col<2||col>31)return;
    if(!edgesByRow.has(row))edgesByRow.set(row,new Map());
    const byCol=edgesByRow.get(row);
    if(!byCol.has(col))byCol.set(col,new Set());
    byCol.get(col).add(edge);
  };
  for(const g of groups){
    const top=rank.get(g.code),bottom=rank.get(g.description);
    if(!top||!bottom)continue;
    for(const center of g.productCols){
      const left=center-1,right=center+1;
      for(let row=top;row<=bottom;row++){
        addEdge(row,left,"left");
        addEdge(row,right,"right");
      }
      for(let col=left;col<=right;col++){
        addEdge(top,col,"top");
        addEdge(bottom,col,"bottom");
      }
    }
  }
  // Do not create, remove or merge any cells. Only add a missing border edge
  // where a cell already exists on the card perimeter.
  for(const row of children(child(sheetDoc.documentElement,"sheetData"),"row")){
    const byCol=edgesByRow.get(+row.getAttribute("r"));
    if(!byCol)continue;
    for(const cell of children(row,"c")){
      const col=parseAddr(cell.getAttribute("r"))?.col;
      const edges=byCol.get(col);
      if(!edges)continue;
      const before=+(cell.getAttribute("s")||0);
      const after=outlineStyle(before,edges);
      if(after!==before)cell.setAttribute("s",String(after));
    }
  }
  borders.setAttribute("count",String(borders.children.length));
  xfs.setAttribute("count",String(xfs.children.length));
}

function makeRelDoc(records){
  const d=document.implementation.createDocument(PKG,"Relationships",null);
  for(const [id,type,target] of records){
    const e=d.createElementNS(PKG,"Relationship");e.setAttribute("Id",id);e.setAttribute("Type",type);e.setAttribute("Target",target);d.documentElement.append(e);
  }
  return d;
}
async function buildOutput(){
  const {sheet,draw}=analysis;
  const {doc,rank,rowTypes,rowCount}=transformSheet(analysis);
  const {doc:drawDoc,images}=transformDrawing(draw,rank);
  if(draw.count && images<analysis.productCount)throw new Error("圖片座標數量異常（"+images+"），為避免遺失圖片，已停止下載。");
  const zip=new JSZip();
  const R="http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  const baseCT=await getXML(sourceZip,"[Content_Types].xml");
  const outCT=baseCT.implementation.createDocument(CT,"Types",null);
  const def=(ext,type)=>{const el=outCT.createElementNS(CT,"Default");el.setAttribute("Extension",ext);el.setAttribute("ContentType",type);outCT.documentElement.append(el);};
  const override=(path,type)=>{const el=outCT.createElementNS(CT,"Override");el.setAttribute("PartName","/"+path);el.setAttribute("ContentType",type);outCT.documentElement.append(el);};
  const origTypes=new Map(children(baseCT.documentElement,"Override").map(e=>[e.getAttribute("PartName").replace(/^\//,""),e.getAttribute("ContentType")]));
  for(const e of children(baseCT.documentElement,"Default")){
    const ext=e.getAttribute("Extension");
    if(["png","jpg","jpeg","gif","webp","bmp","emf","wmf","svg"].includes(ext))def(ext,e.getAttribute("ContentType"));
  }
  def("rels","application/vnd.openxmlformats-package.relationships+xml");def("xml","application/xml");
  override("xl/workbook.xml","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
  override(sheet.path,"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
  if(sourceZip.file("xl/styles.xml"))override("xl/styles.xml",origTypes.get("xl/styles.xml")||"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
  if(sourceZip.file("xl/theme/theme1.xml"))override("xl/theme/theme1.xml",origTypes.get("xl/theme/theme1.xml")||"application/vnd.openxmlformats-officedocument.theme+xml");
  if(drawDoc)override(draw.drawPath,origTypes.get(draw.drawPath)||"application/vnd.openxmlformats-officedocument.drawing+xml");
  zip.file("[Content_Types].xml",serialize(outCT));
  zip.file("_rels/.rels",serialize(makeRelDoc([["rId1",R+"officeDocument","xl/workbook.xml"]])));
  const workbook=await getXML(sourceZip,"xl/workbook.xml");
  for(const el of [...workbook.documentElement.children]){
    if(["definedNames","externalReferences","pivotCaches","calcPr","extLst","smartTagPr","smartTagTypes","webPublishing","fileSharing"].includes(el.localName))workbook.documentElement.removeChild(el);
  }
  const sheetList=child(workbook.documentElement,"sheets");
  for(const s of [...sheetList.children])sheetList.removeChild(s);
  const only=workbook.createElementNS(NS,"sheet");
  only.setAttribute("name",sheet.name);only.setAttribute("sheetId","1");only.setAttributeNS(REL,"r:id","rId1");sheetList.append(only);
  zip.file("xl/workbook.xml",serialize(workbook));
  const workbookRels=[["rId1",R+"worksheet",sheet.path.replace(/^xl\//,"")]];
  if(sourceZip.file("xl/styles.xml")){
    const stylesDoc=await getXML(sourceZip,"xl/styles.xml");
    applyCleanStyles(doc,stylesDoc,analysis.groups,rank);
    zip.file("xl/styles.xml",serialize(stylesDoc));
    workbookRels.push(["rId2",R+"styles","styles.xml"]);
  }
  if(sourceZip.file("xl/theme/theme1.xml")){zip.file("xl/theme/theme1.xml",await sourceZip.file("xl/theme/theme1.xml").async("uint8array"));workbookRels.push(["rId3",R+"theme","theme/theme1.xml"]);}
  zip.file("xl/_rels/workbook.xml.rels",serialize(makeRelDoc(workbookRels)));
  zip.file(sheet.path,serialize(doc));
  if(drawDoc){
    zip.file(draw.drawPath,serialize(drawDoc));
    const relFile=sourceZip.file(draw.drawRelsPath);
    const origDrawRels=relFile ? xml(await relFile.async("string")) : null;
    const relsUsed=new Set([...drawDoc.documentElement.querySelectorAll("*")].map(x=>x.getAttributeNS(REL,"embed")).filter(Boolean));
    const relRecords=[];
    if(origDrawRels)for(const item of children(origDrawRels.documentElement,"Relationship")){
      const id=item.getAttribute("Id");if(!relsUsed.has(id))continue;
      const path=resolvePath(draw.drawPath,item.getAttribute("Target"));
      if(sourceZip.file(path)){
        zip.file(path,await sourceZip.file(path).async("uint8array"));
        relRecords.push([id,item.getAttribute("Type"),item.getAttribute("Target")]);
      }
    }
    if(relsUsed.size>relRecords.length)throw new Error("有商品圖片無法正確連結，已停止匯出。");
    if(relRecords.length)zip.file(draw.drawRelsPath,serialize(makeRelDoc(relRecords)));
    const sheetRelPath=sheet.path.replace("worksheets/","worksheets/_rels/")+".rels";
    zip.file(sheetRelPath,serialize(makeRelDoc([[draw.sheetDrawRelId,R+"drawing",draw.drawPath.substring(draw.drawPath.indexOf("drawings/")-3).replace(/^xl\//,"../")]])));
    // Relative to xl/worksheets, the drawing is always one directory above.
    const rel=xml(await zip.file(sheetRelPath).async("string"));const r=child(rel.documentElement,"Relationship");
    r.setAttribute("Target","../"+draw.drawPath.replace(/^xl\//,""));zip.file(sheetRelPath,serialize(rel));
  }else{
    for(const el of [...doc.documentElement.children])if(el.localName==="drawing")doc.documentElement.removeChild(el);
    zip.file(sheet.path,serialize(doc));
  }
  const blob=await zip.generateAsync({type:"blob",compression:"DEFLATE",compressionOptions:{level:5}});
  return {blob,rowCount,images,bytes:blob.size};
}
async function run(){
  if(!analysis||!inputFile)return;
  const btn=$("run");btn.disabled=true;btn.textContent="正在整理 Excel…";
  try{
    // Re-read the workbook to ensure the saved source is not mutated across retries.
    const doc=await getXML(sourceZip,analysis.sheet.path);
    analysis.doc=doc;
    const draw=await drawingStats(sourceZip,analysis.sheet);analysis.draw=draw;
    const result=await buildOutput();
    const filename=inputFile.name.replace(/\.xlsx$/i,"")+"_跨部門分享版.xlsx";
    const url=URL.createObjectURL(result.blob),a=document.createElement("a");
    a.href=url;a.download=filename;document.body.append(a);a.click();a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),60000);
    showStatus("已下載："+filename+"。商品批次："+analysis.groups.length+"，商品："+analysis.productCount+" 款，圖片："+result.images+" 個，整理後："+result.rowCount+" 列。請檢查圖片及售價後再分享。");
  }catch(e){console.error(e);showStatus(e?.message||"處理過程失敗。",true);}
  finally{btn.disabled=false;btn.textContent="整理並下載分享版 Excel";}
}
document.addEventListener("DOMContentLoaded",()=>{
  const file=$("fileInput"),drop=$("drop");
  $("pick").addEventListener("click",()=>file.click());
  drop.addEventListener("click",e=>{if(e.target.id!=="pick")file.click();});
  drop.addEventListener("keydown",e=>{if(e.key==="Enter"||e.key===" "){e.preventDefault();file.click();}});
  file.addEventListener("change",()=>{if(file.files?.[0])inspectFile(file.files[0]);});
  drop.addEventListener("dragover",e=>{e.preventDefault();drop.classList.add("drag");});
  drop.addEventListener("dragleave",()=>drop.classList.remove("drag"));
  drop.addEventListener("drop",e=>{e.preventDefault();drop.classList.remove("drag");if(e.dataTransfer.files[0])inspectFile(e.dataTransfer.files[0]);});
  $("sheetSelect").addEventListener("change",refreshStats);
  $("run").addEventListener("click",run);
});
