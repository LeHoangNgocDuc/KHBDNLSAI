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
  if (/^I\s*[.\-:]\s*M[ỤU]C\s*TI[ÊE]U/i.test(text)) return 'I';
  if (/^II\s*[.\-:]\s*THI[ẾE]T\s*B[ỊI]/i.test(text)) return 'II';
  if (/^III\s*[.\-:]\s*TI[ẾE]N\s*TR[ÌI]NH/i.test(text)) return 'III';
  return null;
}
function activityName(text) {
  const s = text.trim();
  if (/^(?:[A-D]\s*[.]\s*HOẠT ĐỘNG|HOẠT ĐỘNG\s*\d+\s*[:.])/i.test(s)) return s.slice(0,120);
  // Paragraph headings may be uppercase and contain no spaces after a letter.
  if (/^[A-D]\s*\.\s*HOAT DONG/i.test(normalize(s).toUpperCase())) return s.slice(0,120);
  return null;
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
  let section='', activity='';
  const paragraphs=elems(body,'p');
  const blocks = paragraphs.map((p,i)=>{
    const text=pText(p).replace(/\s+/g,' ').trim();
    const insideTable=!!parentOfTag(p,'tc');
    if(!insideTable){
      const s=sectionName(text);if(s) {section=s; if(s!=='III')activity='';}
      if(section==='III') {const a=activityName(text); if(a)activity=a;}
    }
    const found=/^Bước\s*([1-4])\s*[:.\-]/i.exec(text);
    return {id:'p'+i,element:p,text,section,activity,step:found ? `Bước ${found[1]}`:'', ...getCellMeta(p,tables)};
  });
  const anchors=[];
  const objectives=blocks.filter(b=>b.section==='I' && b.table===-1);
  const firstPham=objectives.findIndex(b=>/^3\s*[.]\s*phẩm chất/i.test(b.text));
  const idx2=objectives.findIndex(b=>/^2\s*[.]\s*năng lực/i.test(b.text));
  if(idx2>=0){
    const end=firstPham>idx2 ? firstPham-1 : objectives.length-1;
    const b=objectives[end];
    if(b)anchors.push({id:'goals-competency',kind:'goal',blockId:b.id,section:'I',activity:'Mục tiêu → Năng lực',step:'',column:'',preview:b.text.slice(0,120)});
  }
  if(firstPham>=0){
    const b=objectives[objectives.length-1];
    if(b)anchors.push({id:'goals-inclusion',kind:'goal',blockId:b.id,section:'I',activity:'Mục tiêu → Phẩm chất',step:'',column:'',preview:b.text.slice(0,120)});
  }
  // Add an explicit anchor at the end of Bước 1 only when another step follows.
  // Homework is a standalone section; never infer it from a table cell.
  const home=blocks.find(b=>b.table<0 && /^(?:[*•-]\s*)?HƯỚNG DẪN VỀ NHÀ/i.test(b.text));
  if(home)anchors.push({id:'homework',kind:'homework',blockId:home.id,section:'III',activity:'Hướng dẫn về nhà',step:'',column:'',preview:home.text});
  let lastActivity='';
  blocks.forEach(b=>{
    if(b.section!=='III'||!b.step) return;
    if(b.table>=0 && b.role!=='teacher') return; // Never guess the instructional column of a table.
    if(!b.activity) return;
    if(b.step==='Bước 1'||b.step==='Bước 2'){
      if(anchors.some(a=>a.kind==='activity' && a.blockId===b.id))return;
      // Activity headers and step markers are anchors; material is added directly after the step.
      anchors.push({id:'step-'+b.id,kind:'activity',blockId:b.id,section:'III',activity:b.activity,step:b.step,column:b.cell,preview:b.text.slice(0,150)});
    }
    lastActivity=b.activity;
  });
  for(const b of blocks.filter(b=>b.section==='III' && b.step==='Bước 1' && (b.table<0||b.role==='teacher'))){
    const group=blocks.filter(x=>x.section==='III'&&x.activity===b.activity&&x.table===b.table&&x.column===b.column);
    const pos=group.indexOf(b),nextStep=group.findIndex((x,j)=>j>pos && !!x.step);
    if(nextStep<=pos+1)continue;
    const tail=group[nextStep-1];
    if(tail&&tail!==b){anchors.push({id:'step1-tail-'+tail.id,kind:'activity-end',blockId:tail.id,section:'III',activity:b.activity,step:'Cuối Bước 1',column:b.cell,preview:tail.text.slice(0,150)});}
  }
  return {body,tables,blocks,anchors};
}
function makeParagraph(doc,anchorP,text,color,heading=false){
  const p=newEl(doc,'p');
  const ppr=kids(anchorP,'pPr')[0];
  if(ppr&&!heading){const clone=ppr.cloneNode(true);for(const num of kids(clone,'numPr'))clone.removeChild(num);p.appendChild(clone);}
  if(heading){const pPr=newEl(doc,'pPr');pPr.appendChild(newEl(doc,'spacing',{before:'120',after:'80'}));p.appendChild(pPr);}
  addRichRuns(doc,p,text,color);
  if(heading){const r=kids(p,'r')[0];const rPr=r?kids(r,'rPr')[0]:null;if(rPr)rPr.appendChild(newEl(doc,'b'));}
  return p;
}
function proposalLines(p){
  const c=p.code ? ` [${p.code}]`:'';
  if(p.type==='HSHN' && p.anchorId==='goals-inclusion')return p.content.trim().split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
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
  const relevant=originalSnapshot.blocks.filter(b=>anchor.kind==='goal' ? b.section==='I' : b.section==='III'&&b.activity===block.activity);
  if(proposal.code){
    const wanted=digitsOnly(proposal.code);
    return relevant.some(b=>digitsOnly(b.text).includes(wanted) && (b.text.includes('NLS')||b.text.includes('AI')||b.text.toUpperCase().includes(proposal.code.toUpperCase())));
  }
  if(proposal.type==='HSHN'){
    const marker=anchor.kind==='goal' ? /gi[aá]o d[uụ]c h[oò]a nh[aậ]p|GDHN\.KT|\[HSHN\]/i : /\[T[ií]ch h[oợ]p GDHN\]|\[T[ií]ch h[oợ]p HSHN\]|\[HSHN\]/i;
    if(relevant.some(b=>marker.test(b.text)))return true;
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
    anchors:snapshot.anchors,
    counts:{tables:snapshot.tables.length,paragraphs:snapshot.blocks.length,equations:count(doc,'oMath')+count(doc,'oMathPara'),pictures:count(doc,'drawing'),embeddedObjects:count(doc,'object')}};
  return {zip,oldXml,doc,snapshot,info};
}
export function reportContext(opened){
  const {blocks,anchors}=opened.snapshot;
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

export async function exportDocx(opened,proposals,{allowUnplaced=true,convertExistingLatex=false}={}) {
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
    const anchor=anchors.get(p.anchorId), b=anchor&&blocks.get(anchor.blockId);
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
    const matchesSlot=slot==='goal'?anchor.kind==='goal'&&anchor.id===(p.type==='HSHN'?'goals-inclusion':'goals-competency'):
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
