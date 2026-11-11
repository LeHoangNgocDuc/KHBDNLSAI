import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {createRequire} from 'node:module';
const require=createRequire(import.meta.url);
try{globalThis.JSZip=require('jszip')}catch{globalThis.JSZip=require('/usr/local/slides_js/node_modules/jszip');}
let dom;try{dom=require('@xmldom/xmldom')}catch{dom=require('/usr/local/slides_js/node_modules/@xmldom/xmldom');}
globalThis.DOMParser=dom.DOMParser;globalThis.XMLSerializer=dom.XMLSerializer;
const {openDocx,exportDocx}=await import('../public/src/docx-engine.js');
const {latexToOmml,richFragments}=await import('../public/src/word-math.js');
const {inclusiveObjectives,inclusiveActivity}=await import('../public/src/inclusion-policy.js');
const {offlineProposals}=await import('../public/src/catalog.js');
const sample=process.env.TEST_FIXTURE_ONLY==='1'?new URL('./fixture.docx',import.meta.url):(fs.existsSync('/mnt/data/DS9_T9_T23_24_B6.docx')?'/mnt/data/DS9_T9_T23_24_B6.docx':new URL('./fixture.docx',import.meta.url));
const file=fs.readFileSync(sample);
const lesson={subject:'Toán',name:'Bất phương trình bậc nhất một ẩn',nls:[{code:'1.1.TC2a',text:'Tìm kiếm dữ liệu trực tuyến'}],ai:[{code:'9.D1.1',text:'Kiểm tra kết quả AI'}]};
const ns='http://schemas.openxmlformats.org/officeDocument/2006/math';
test('LaTeX becomes Word-native OMML fractions, radicals and scripts',()=>{
 const document=new DOMParser().parseFromString('<root/>','text/xml');
 const math=latexToOmml(document,String.raw`x^2+\frac{a_1}{\sqrt{b}}\ge 3`);
 assert.equal(math.getElementsByTagNameNS(ns,'f').length,1);
 assert.equal(math.getElementsByTagNameNS(ns,'rad').length,1);
 assert.equal(math.getElementsByTagNameNS(ns,'sSup').length,1);
 assert.equal(math.getElementsByTagNameNS(ns,'sSub').length,1);
 assert.ok(!new XMLSerializer().serializeToString(math).includes('\\frac'));
});
test('HTML br is normalized to newline and unbalanced dollar rejected',()=>{
 assert.equal(richFragments('Đọc<br>Tiếp')[0].value,'Đọc\nTiếp');
 assert.throws(()=>richFragments('Một biểu thức $x+1'),/thiếu dấu/i);
});
test('No-NLS HSKT only receives non-digital support',()=>{
 const x=inclusiveActivity(lesson,{types:['mobility','vision'],digital:false});
 assert.match(x,/máy tính cầm tay/);
 assert.doesNotMatch(x,/GeoGebra|Excel|thao tác phần mềm.*HSKT/i);
 const y=inclusiveActivity(lesson,{types:['mobility','vision'],digital:true});
 assert.match(y,/chuột, bàn phím/);
});
test('HSKT goals contain mandatory subheadings using actual Word objectives',async()=>{
 const opened=await openDocx(file);
 const goal=inclusiveObjectives(opened,lesson,{types:['vision','mobility'],digital:false});
 for(const str of ['1. Kiến thức','2. Năng lực','Năng lực chung','Năng lực riêng','3. Phẩm chất','[GDHN.KT]','[GDHN.KN]'])assert.ok(goal.includes(str),str);
 assert.ok(goal.includes('Thị giác'));
});
test('offline proposal places HSKT goal in I, support at step2 and adds NLS homework',async()=>{
 const opened=await openDocx(file);
 const output=offlineProposals(lesson,opened,['NLS','HSHN'],{types:['intellectual'],digital:false});
 assert.ok(output.some(x=>x.type==='HSHN'&&x.slot==='goal'&&x.anchorId==='goals-inclusion'));
 assert.ok(output.some(x=>x.type==='HSHN'&&x.slot==='activity'));
 if(opened.snapshot.anchors.some(x=>x.kind==='activity'&&x.step==='Bước 2'))assert.ok(output.some(x=>x.type==='HSHN'&&x.slot==='activity'&&x.anchorId));
 const homework=output.find(x=>x.slot==='homework');assert.ok(homework);
 // The document may not have a structurally reliable homework anchor, in which case it must be appended.
 if(!opened.snapshot.anchors.some(x=>x.kind==='homework'))assert.equal(homework.anchorId,'');
});
test('native Equation from new LaTeX survives DOCX export with original Word parts untouched',async()=>{
 const original=await openDocx(file);
 const a=original.snapshot.anchors.find(x=>x.kind==='activity'&&x.step==='Bước 1');assert.ok(a);
 const props=[{type:'NLAI',code:'9.D1.1',content:String.raw`HS xét công thức $x^2+\frac{1}{2}\ge 3$ rồi kiểm chứng kết quả do AI đề xuất.`,anchorId:a.id,slot:'activity',accepted:true}];
 const result=await exportDocx(original,props);
 assert.equal(result.summary.inserted,1);
 const zip=await globalThis.JSZip.loadAsync(Buffer.from(await result.blob.arrayBuffer()));
 const originalZip=await globalThis.JSZip.loadAsync(file);
 const text=await zip.file('word/document.xml').async('string');
 assert.ok(text.includes('m:oMath'));
 assert.ok(text.includes('m:f'));
 assert.ok(!text.includes('$x^2'));
 for(const key of Object.keys(originalZip.files).filter(x=>x!=='word/document.xml'&&!x.endsWith('/'))){
  assert.ok(Buffer.from(await originalZip.file(key).async('nodebuffer')).equals(Buffer.from(await zip.file(key).async('nodebuffer'))),key);
 }
 fs.writeFileSync(new URL('./sample-equation-after.docx',import.meta.url),Buffer.from(await result.blob.arrayBuffer()));
});
test('opt-in converts existing LaTeX runs into actual Word Equation',async()=>{
 const source=fs.readFileSync(new URL('./latex-fixture.docx',import.meta.url));
 const original=await openDocx(source);
 const result=await exportDocx(original,[],{convertExistingLatex:true});
 assert.equal(result.summary.latexConverted,2);
 const zip=await globalThis.JSZip.loadAsync(Buffer.from(await result.blob.arrayBuffer()));
 const xml=await zip.file('word/document.xml').async('string');
 assert.ok(xml.includes('m:oMath'));
 assert.ok(!xml.includes('$x^2'));
 assert.ok(!xml.includes('$\\sqrt'));
 const other=await globalThis.JSZip.loadAsync(source);
 for(const key of Object.keys(other.files).filter(x=>x!=='word/document.xml'&&!x.endsWith('/'))){
   assert.ok(Buffer.from(await other.file(key).async('nodebuffer')).equals(Buffer.from(await zip.file(key).async('nodebuffer'))),key);
 }
});
test('unsupported LaTeX must not silently become a broken text formula',async()=>{
 const opened=await openDocx(file);
 const anchor=opened.snapshot.anchors.find(x=>x.kind==='activity'&&x.step==='Bước 1');
 await assert.rejects(()=>exportDocx(opened,[{type:'NLAI',code:'TEST',content:String.raw`Biểu thức $\foobar{x}$ cần xử lý.`,anchorId:anchor.id,slot:'activity',accepted:true}]),/Không chuyển được công thức/);
});
