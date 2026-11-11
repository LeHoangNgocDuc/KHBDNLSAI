import {addRichRuns, richFragments, latexToOmml, WORD_NS} from './word-math.js';
const JSZip = globalThis.JSZip;
if (!JSZip) throw new Error('Không tải được thư viện JSZip.');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const XML = 'http://www.w3.org/XML/1998/namespace';
const COLORS = { NLS: 'DC2626', NLAI: '2563EB', HSHN: '7C3AED' };
const TYPES = new Set(Object.keys(COLORS));
const normalize = v => String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/gi, 'd').toLowerCase().replace(/[^\p{L}\p{N}]+/gu, ' ').trim();
const digitsOnly = v => String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
const elems = (root, tag) => Array.from(root.getElementsByTagNameNS(W, tag));
const first = (root, tag) => elems(root, tag)[0] || null;
const tag = n => n?.localName || n?.nodeName?.split(':').pop() || '';
const parentOfTag = (node, wanted) => { for (let n=node?.parentNode; n; n=n.parentNode) if (tag(n)===wanted && n.namespaceURI===W) return n; return null; };
const kids = (node, wanted) => Array.from(node?.childNodes || []).filter(n => n.nodeType === 1 && n.namespaceURI === W && (!wanted || tag(n)===wanted));
const pText = p => elems(p, 't').map(n=>n.textContent).join('');
const xmlParser = xml => {
  const doc = new DOMParser().parseFromString(xml,'application/xml');
  const bad = doc.getElementsByTagName('parsererror');
  if (bad.length) throw new Error('Tài liệu Word có XML không hợp lệ: ' + bad[0].textContent.slice(0,180));
  return doc;
};
function newEl(doc, kind, attributes={}) {
  const n=doc.createElementNS(W, `w:${kind}`);
  for(const [key,value] of Object.entries(attributes)) n.setAttributeNS(W, `w:${key}`, value);
  return n;
}
function getDocumentXml(zip) { return zip.file('word/document.xml')?.async('string'); }
function sectionName(text) {
  const s=normalize(text);
  if (/^i muc tieu\b/.test(s)) return 'I';
  if (/^ii thiet bi\b/.test(s)) return 'II';
  if (/^iii tien trinh\b/.test(s)) return 'III';
  return null;
}
function activityName(text) {
  const s=text.trim();
  // Tin học: 1. Hoạt động 1, 2.1 Hoạt động 1, 2.1.2. Hoạt động 2.
  // Toán: A. HOẠT ĐỘNG, Hoạt động 1, Hoạt động 2.
  return /^(?:(?:[A-D]|\d+(?:\.\d+)*)\s*\.?\s*)?hoạt\s+động(?:\s*\d+)?\b/i.test(s) ? s.slice(0,140) : null;
}
function isUnitTitle(text){
  const s=text.trim();
  if(s.length>190 || !/^(?:TIẾT[^\n]{0,48}?BÀI|BÀI)\s*\d+/i.test(s))return false;
  const letters=s.match(/[a-zA-ZÀ-Ỹà-ỹĐđ]/g)||[];
  const upper=letters.filter(c=>c===c.toUpperCase()&&c!==c.toLowerCase()).length;
  return letters.length>0 && upper/letters.length>.63;
}
function getCellMeta(p, tables) {
  const tc=parentOfTag(p,'tc'); if(!tc)return {cell:'',column:-1,table:-1,role:''};
  const tr=parentOfTag(tc,'tr'), tbl=parentOfTag(tc,'tbl');
  const columns=kids(tr,'tc');const column=columns.indexOf(tc);
  const index=tables.indexOf(tbl);
  const headerRow=kids(tbl,'tr')[0];
  const headings=kids(headerRow,'tc');
  const heading=headings[column] ? elems(headings[column],'p').map(pText).join(' ').slice(0,150) : '';
  const label=normalize(heading);
  const role = /(gv.*hs|giao vien.*hoc sinh|to chuc thuc hien|hoat dong cua gv|hoat dong cua giao vien)/i.test(label) ? 'teacher' : /(san pham|ket qua|noi dung ghi bang)/i.test(label) ? 'product' : '';
  return {cell:heading,column,table:index,role};
}
function snapshotFrom(doc) {
  const body=first(doc,'body');if(!body)throw new Error('Không tìm thấy nội dung chính của giáo án.');
  const tables=elems(body,'tbl');
  let section='',activity='',unitId='u0',unitIndex=0;
  const units=[];
  const paragraphs=elems(body,'p');
  const blocks=paragraphs.map((p,i)=>{
    const text=pText(p).replace(/\s+/g,' ').trim();
    const insideTable=!!parentOfTag(p,'tc');
    if(!insideTable){
      if(isUnitTitle(text)){
        unitId=`u${++unitIndex}`;units.push({id:unitId,title:text,index:unitIndex});
        section='';activity='';
      }
      const sec=sectionName(text);if(sec){section=sec;if(sec!=='III')activity='';}
      if(section==='III'){const a=activityName(text);if(a)activity=a;}
    }
    const found=/^Bước\s*([1-4])\s*[:.\-]/i.exec(text);
    return {id:'p'+i,element:p,text,section,activity,unitId,step:found?`Bước ${found[1]}`:'',...getCellMeta(p,tables)};
  });
  if(!units.length){units.push({id:'u0',title:'Giáo án hiện tại',index:0});}
  const anchors=[];
  for(const [ui,unit] of units.entries()){
    const objectives=blocks.filter(b=>b.unitId===unit.id&&b.section==='I'&&b.table===-1);
    const pham=objectives.findIndex(b=>/^3\s*\.\s*(?:về\s*)?phẩm chất/i.test(b.text));
    const comp=objectives.findIndex(b=>/^2\s*\.\s*(?:về\s*)?năng lực/i.test(b.text));
    const suffix=ui===0?'':`-${unit.id}`;
    if(comp>=0){
      let end=pham>comp?pham-1:objectives.length-1;
      if(end>=comp){const b=objectives[end];anchors.push({id:`goals-competency${suffix}`,kind:'goal',blockId:b.id,unitId:unit.id,section:'I',activity:'Mục tiêu → Năng lực',step:'',column:'',preview:b.text.slice(0,120)});}
    }
    if(pham>=0){
      const tail=objectives.slice(pham+1);
      const hsk=tail.findIndex(b=>/mục tiêu dành cho học sinh khuyết tật|điều chỉnh mục tiêu giáo dục hòa nhập/i.test(normalize(b.text)));
      const regular= hsk>=0?tail.slice(0,hsk):tail;
      const b=regular.at(-1)||objectives[pham];
      anchors.push({id:`goals-inclusion${suffix}`,kind:'goal',blockId:b.id,unitId:unit.id,section:'I',activity:'Mục tiêu → Sau Phẩm chất',step:'',column:'',preview:b.text.slice(0,120)});
    }
    const home=blocks.find(b=>b.unitId===unit.id&&b.table<0&&/^(?:[*•-]\s*)?(?:HƯỚNG DẪN VỀ NHÀ|HƯỚNG DẪN TỰ HỌC Ở NHÀ)/i.test(b.text));
    if(home)anchors.push({id:`homework-${home.id}`,kind:'homework',blockId:home.id,unitId:unit.id,section:'III',activity:'Hướng dẫn về nhà',step:'',column:'',preview:home.text});
  }
  for(const b of blocks){
    if(b.section!=='III'||!b.step||!b.activity)continue;
    if(b.table>=0&&b.role!=='teacher')continue;
    if(b.step==='Bước 1'||b.step==='Bước 2')
      anchors.push({id:'step-'+b.id,kind:'activity',blockId:b.id,unitId:b.unitId,section:'III',activity:b.activity,step:b.step,column:b.cell,preview:b.text.slice(0,150)});
  }
  // Keep step 1 summaries inside the *same unit, activity, table and teacher column*.
  for(const b of blocks.filter(b=>b.section==='III'&&b.step==='Bước 1'&&(b.table<0||b.role==='teacher'))){
    const group=blocks.filter(x=>x.section==='III'&&x.unitId===b.unitId&&x.activity===b.activity&&x.table===b.table&&x.column===b.column);
    const pos=group.indexOf(b),next=group.findIndex((x,j)=>j>pos&&!!x.step);
    if(next<=pos+1)continue;
    const tail=group[next-1];
    if(tail&&tail!==b)anchors.push({id:'step1-tail-'+tail.id,kind:'activity-end',blockId:tail.id,unitId:b.unitId,section:'III',activity:b.activity,step:'Cuối Bước 1',column:b.cell,preview:tail.text.slice(0,150)});
  }
  return {body,tables,blocks,anchors,units};
}
function makeParagraph(doc,anchorP,text,color,heading=false){
  const p=newEl(doc,'p');
  const ppr=kids(anchorP,'pPr')[0];
  if(ppr&&!heading){const clone=ppr.cloneNode(true);for(const num of kids(clone,'numPr'))clone.removeChild(num);p.appendChild(clone);}
  if(heading){const pPr=newEl(doc,'pPr');pPr.appendChild(newEl(doc,'spacing',{before:'120',after:'80'}));p.appendChild(pPr);}
  const origSz=elems(anchorP,'sz')[0]?.getAttributeNS(W,'val');
  const pointSize=(origSz && /^\d{2,3}$/.test(origSz))?origSz:'26';
  addRichRuns(doc,p,text,color,pointSize);
  const goalHeader=/^(?:\*\s*MỤC TIÊU DÀNH CHO|\*\s*Điều chỉnh mục tiêu|[123]\.\s*(?:Kiến thức|Năng lực|Phẩm chất)|Năng lực chung:|Năng lực riêng)/i.test(text);
  if(goalHeader)for(const run of kids(p,'r')){
    const props=kids(run,'rPr')[0];if(props)props.appendChild(newEl(doc,'b'));
  }
  if(heading){const r=kids(p,'r')[0];const rPr=r?kids(r,'rPr')[0]:null;if(rPr)rPr.appendChild(newEl(doc,'b'));}
  return p;
}
function proposalLines(p){
  const c=p.code ? ` [${p.code}]`:'';
  if(p.type==='HSHN' && p.slot==='goal')return p.content.trim().split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
  const label=p.type==='NLS'?'NLS':p.type==='NLAI'?'Năng lực AI':'GDHN';
  const prefix=p.slot==='step1-summary'?'[Tích hợp NLS]':p.slot==='homework'?'Nhiệm vụ NLS':p.type==='HSHN'&&p.slot==='activity'?'[GDHN] Biện pháp hỗ trợ HSKT':`[Tích hợp ${label}]${c}`;
  const raw=p.content.trim().split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
  if(!raw.length)return [];
  if(p.type==='HSHN'&&p.slot==='activity')return [`- ${prefix}: ${raw[0]}`,...raw.slice(1)];
  if(p.slot==='homework')return [`- ${prefix}: ${raw[0]}`,...raw.slice(1)];
  if(p.slot==='step1-summary')return [prefix,...raw];
  return [`- ${prefix}: ${raw[0]}`,...raw.slice(1)];
}
function formatProposal(p){
  const c=p.code ? ` [${p.code}]` : '';
  const label=p.type==='NLS'?'NLS':p.type==='NLAI'?'Năng lực AI':'HSHN';
  return `- [Tích hợp ${label}]${c}: ${p.content.trim()}`;
}
function isDuplicate(proposal, anchor, originalSnapshot) {
  const block=originalSnapshot.blocks.find(b=>b.id===anchor.blockId);
  if(!block)return false;
  const relevant=originalSnapshot.blocks.filter(b=>b.unitId===block.unitId && (anchor.kind==='goal' ? b.section==='I' : b.section==='III'&&b.activity===block.activity));
  if(proposal.code){
    const wanted=digitsOnly(proposal.code);
    return relevant.some(b=>digitsOnly(b.text).includes(wanted) && (b.text.includes('NLS')||b.text.includes('AI')||b.text.toUpperCase().includes(proposal.code.toUpperCase())));
  }
  if(proposal.type==='HSHN'){
    if(anchor.kind==='goal'){
      if(relevant.some(b=>/mục tiêu dành cho học sinh khuyết tật|giáo dục hòa nhập|GDHN\.KT/i.test(b.text)))return true;
    }else{
      // A HSHN support already placed in Bước 1 must not prevent a new
      // accommodation at the beginning of Bước 2 of the same activity.
      const same=originalSnapshot.blocks.filter(b=>b.unitId===block.unitId && b.activity===block.activity && b.table===block.table && b.column===block.column);
      const start=same.findIndex(b=>b.id===block.id);
      const end=same.findIndex((b,i)=>i>start && !!b.step);
      const region=start>=0?same.slice(start,end>=0?end:undefined):[];
      if(region.some(b=>/\[GDHN\]|\[T[ií]ch h[oợ]p GDHN\]|\[T[ií]ch h[oợ]p HSHN\]|\[HSHN\]/i.test(b.text)))return true;
    }
  }
  const needle=normalize(proposal.content).slice(0,70);
  return needle.length>25 && relevant.some(b=>normalize(b.text).includes(needle));
}
const count = (node,name) => node.getElementsByTagNameNS(name==='oMath'||name==='oMathPara'?M:W,name).length;
export async function openDocx(input) {
  const zip=await JSZip.loadAsync(input);
  const oldXml=await getDocumentXml(zip);
  if(!oldXml)throw new Error('Tệp không chứa word/document.xml. Chỉ hỗ trợ định dạng .docx thật.');
  const doc=xmlParser(oldXml),snapshot=snapshotFrom(doc);
  const info={name:input?.name||'Giao_an.docx',size:input?.size||0,
    blocks:snapshot.blocks.map(({element,...rest})=>rest),
    anchors:snapshot.anchors,units:snapshot.units,
    counts:{tables:snapshot.tables.length,paragraphs:snapshot.blocks.length,equations:count(doc,'oMath')+count(doc,'oMathPara'),pictures:count(doc,'drawing'),embeddedObjects:count(doc,'object')}};
  return {zip,oldXml,doc,snapshot,info};
}
export function reportContext(opened,{unitId}={}){
  let {blocks,anchors}=opened.snapshot;
  if(unitId){blocks=blocks.filter(b=>b.unitId===unitId);anchors=anchors.filter(a=>a.unitId===unitId);}
  const headlines=blocks.filter(b=>b.table===-1 && (b.section==='I'||b.section==='III'))
    .filter(b=>/Bài\s*\d+|Hoạt động|Bước|Mục tiêu|Năng lực|Phẩm chất/i.test(b.text))
    .slice(0,65).map(b=>b.text.slice(0,180));
  const activities=[];
  for(const a of anchors.filter(a=>a.kind==='activity')){
    const items=blocks.filter(b=>b.section==='III' && b.activity===a.activity && (b.role==='teacher'||b.table===-1)).slice(0,18);
    const candidate={id:a.id,activity:a.activity,step:a.step,column:a.column,preview:items.map(x=>x.text).filter(Boolean).join(' ').slice(0,980)};
    if(!activities.some(x=>x.id===candidate.id))activities.push(candidate);
  }
  return {headlines,anchors:anchors.map(({id,kind,activity,step,column,preview})=>({id,kind,activity,step,column,preview})),activities:activities.slice(0,40),
    existing:blocks.filter(b=>/\[Tích hợp|năng lực số|năng lực (AI|trí tuệ nhân tạo)|giáo dục hòa nhập/i.test(b.text)).slice(0,40).map(b=>({section:b.section,activity:b.activity,text:b.text.slice(0,320)}))};
}

function convertOneRunLatex(doc,r){
  const texts=kids(r,'t');if(texts.length!==1)return 0;
  const raw=texts[0].textContent;
  if(!/\$[^$\n]+\$/.test(raw))return 0;
  const pieces=richFragments(raw);
  if(!pieces.some(x=>x.kind==='math'))return 0;
  const parent=r.parentNode,after=r.nextSibling;
  let count=0;
  for(const piece of pieces){
    if(piece.kind==='math'){parent.insertBefore(latexToOmml(doc,piece.value),after);count++;}
    else if(piece.value){const clone=r.cloneNode(true);const node=kids(clone,'t')[0];node.textContent=piece.value;parent.insertBefore(clone,after);}
  }
  parent.removeChild(r);return count;
}
function convertExistingDocxLatex(doc){
  let converted=0,failed=0;
  for(const p of elems(first(doc,'body'),'p')){
    for(const r of kids(p,'r')){
      try{converted+=convertOneRunLatex(doc,r);}catch(e){failed++;}
    }
  }
  return {converted,failed};
}

export async function exportDocx(opened,proposals,{allowUnplaced=true,convertExistingLatex=false,unitId=null}={}) {
  const doc=xmlParser(opened.oldXml),before=snapshotFrom(doc);
  const serial=new XMLSerializer();
  const existingOriginal=before.blocks.map(b=>({node:b.element,xml:serial.serializeToString(b.element),text:b.text}));
  const beforeTables=before.tables.length;
  const beforeMath=count(doc,'oMath')+count(doc,'oMathPara');
  const beforeDraw=count(doc,'drawing');
  const beforeObject=count(doc,'object');
  const anchors=new Map(before.anchors.map(x=>[x.id,x]));
  const blocks=new Map(before.blocks.map(x=>[x.id,x]));
  const allRecords=[];const pending=[];
  const lastForAnchor=new Map();
  const alreadyInserted=new Set();
  for(const raw of proposals){
    if(raw.accepted===false)continue;
    const p={type:String(raw.type||''),code:String(raw.code||'').trim(),content:String(raw.content||'').replace(/<br\s*\/?\s*>/gi,'\n').trim().slice(0,4000),anchorId:String(raw.anchorId||'').trim(),reason:String(raw.reason||'').trim(),slot:String(raw.slot||'')};
    if(!TYPES.has(p.type)||!p.content) continue;
    const rawAnchor=anchors.get(p.anchorId);
    const anchor=(unitId && rawAnchor?.unitId!==unitId)?null:rawAnchor;
    if(rawAnchor&&!anchor)p.reason='Vị trí đề xuất thuộc bài khác với bài đang chọn trong file Word.';
    const b=anchor&&blocks.get(anchor.blockId);
    const sig=[p.type,digitsOnly(p.code),normalize(p.content),p.slot].join('|');
    if(alreadyInserted.has(sig)) {allRecords.push({...p,status:'ALREADY_EXISTS',reason:'Đề xuất trùng trong lượt xử lý.'});continue;}
    if(!b||!anchor){
      const duplicateUnplaced=before.blocks.some(item=>normalize(item.text).includes(normalize(p.content).slice(0,70)) && normalize(p.content).length>25);
      if(duplicateUnplaced){allRecords.push({...p,status:'ALREADY_EXISTS',reason:'Nội dung này đã xuất hiện trong tài liệu hoặc phụ lục trước đó.'});continue;}
      pending.push({...p,status:'APPENDED_UNPLACED',reason:p.reason||'Không tìm thấy vị trí chèn hợp lệ trong cấu trúc Word.'});continue;
    }
    if(isDuplicate(p,anchor,before)) {allRecords.push({...p,status:'ALREADY_EXISTS',reason:'Mã hoặc nội dung tương đương đã có tại mục/hoạt động tương ứng.'});continue;}
    // Re-validate user-adjusted positions at export time, not just in the AI response.
    const slot=p.slot||'activity';
    const matchesSlot=slot==='goal'?anchor.kind==='goal'&&anchor.id.startsWith(p.type==='HSHN'?'goals-inclusion':'goals-competency'):
      slot==='activity'?anchor.kind==='activity'&&anchor.step===(p.type==='HSHN'?'Bước 2':'Bước 1'):
      slot==='step1-summary'?p.type==='NLS'&&anchor.kind==='activity-end':
      slot==='homework'?p.type==='NLS'&&anchor.kind==='homework':false;
    if(!matchesSlot){pending.push({...p,status:'APPENDED_UNPLACED',reason:'Vị trí được chọn không phù hợp loại năng lực hoặc bước tích hợp.'});continue;}
    if(anchor.kind==='activity' || anchor.kind==='activity-end'){
      if(b.table>=0 && b.role!=='teacher'){pending.push({...p,status:'APPENDED_UNPLACED',reason:'Ô bảng không phải cột hoạt động GV và HS.'});continue;}
    }
    if(p.slot==='homework'&&anchor.kind!=='homework'){pending.push({...p,status:'APPENDED_UNPLACED',reason:'Không xác định được mục Hướng dẫn về nhà.'});continue;}
    const current=lastForAnchor.get(anchor.id)||b.element;
    if(!current.parentNode){pending.push({...p,status:'APPENDED_UNPLACED',reason:'Đoạn được chọn không còn hợp lệ.'});continue;}
    try{
      for(const line of proposalLines({...p,anchorId:anchor.id})){
        const para=makeParagraph(doc,b.element,line,COLORS[p.type]);
        current.parentNode.insertBefore(para,(lastForAnchor.get(anchor.id)||current).nextSibling);
        lastForAnchor.set(anchor.id,para);
      }
    }catch(e){throw Error('Không chuyển được công thức LaTeX thành Equation tại '+anchor.activity+': '+e.message);}
    alreadyInserted.add(sig);
    allRecords.push({...p,status:'INSERTED',location:anchor.activity+(anchor.step?' → '+anchor.step:'')});
  }
  if(pending.length && allowUnplaced){
    const body=before.body;
    const sect=kids(body,'sectPr')[0];
    const last=kids(body).filter(n=>tag(n)==='p').at(-1)||first(doc,'p');
    const append = p => body.insertBefore(p,sect||null);
    append(makeParagraph(doc,last,'PHỤ LỤC: NỘI DUNG CHƯA XÁC ĐỊNH ĐƯỢC VỊ TRÍ TÍCH HỢP','222222',true));
    append(makeParagraph(doc,last,'Các nội dung sau chưa được tích hợp vào hoạt động. Giáo viên cần kiểm tra và tự quyết định vị trí.','444444'));
    pending.forEach((p,i)=>{
      append(makeParagraph(doc,last,`${i+1}. ${p.type}${p.code?' ['+p.code+']':''} – Vị trí dự kiến: ${p.anchorId||'Chưa xác định'}. Lý do: ${p.reason}`,'444444'));
      for(const line of proposalLines(p))append(makeParagraph(doc,last,line,COLORS[p.type]));
      allRecords.push({...p,status:'APPENDED_UNPLACED'});
    });
  }
  // Existing LaTeX conversion is opt-in because it necessarily modifies original XML.
  let latexReport={converted:0,failed:0};
  if(convertExistingLatex)latexReport=convertExistingDocxLatex(doc);
  if(!convertExistingLatex)for(const x of existingOriginal) if(serial.serializeToString(x.node)!==x.xml)throw new Error('Kiểm thử bảo toàn thất bại: một đoạn văn gốc đã bị thay đổi.');
  if(count(doc,'tbl')!==beforeTables||count(doc,'oMath')+count(doc,'oMathPara')<beforeMath||count(doc,'drawing')!==beforeDraw||count(doc,'object')!==beforeObject)
    throw new Error('Kiểm thử bảo toàn thất bại: bảng, công thức hoặc đối tượng Word đã thay đổi.');
  const finished=serial.serializeToString(doc);
  const reparse=xmlParser(finished);
  if(count(reparse,'tbl')!==beforeTables||count(reparse,'oMath')+count(reparse,'oMathPara')!==count(doc,'oMath')+count(doc,'oMathPara'))
    throw new Error('Kiểm thử XML cuối cùng không đạt.');
  const zip=await JSZip.loadAsync(await opened.zip.generateAsync({type:'arraybuffer'}));
  zip.file('word/document.xml',finished);
  const blob=await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});
  return {blob,records:allRecords,summary:{inserted:allRecords.filter(x=>x.status==='INSERTED').length,existing:allRecords.filter(x=>x.status==='ALREADY_EXISTS').length,unplaced:pending.length,originalTables:beforeTables,originalEquations:beforeMath,latexConverted:latexReport.converted,latexSkipped:latexReport.failed}};
}
export function filenameFor(name){return name.replace(/\.docx$/i,'')+'_tich_hop_NLS_AI_HSHN.docx';}
export {normalize};
