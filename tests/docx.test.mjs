import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {createRequire} from 'node:module';
const require=createRequire(import.meta.url);
// The production web app runs in browsers with native DOMParser/JSZip vendor script.
// For local container tests, libraries are available in the preinstalled toolchain.
try{globalThis.JSZip=require('jszip')}catch{globalThis.JSZip=require('/usr/local/slides_js/node_modules/jszip');}
let dom;try{dom=require('@xmldom/xmldom')}catch{dom=require('/usr/local/slides_js/node_modules/@xmldom/xmldom');}
globalThis.DOMParser=dom.DOMParser;globalThis.XMLSerializer=dom.XMLSerializer;
const {openDocx,exportDocx,reportContext}=await import('../public/src/docx-engine.js');
const sample='/mnt/data/DS9_T9_T23_24_B6.docx';
const source=fs.existsSync(sample)?sample:new URL('./fixture.docx',import.meta.url);
const original=fs.readFileSync(source);
const readDoc=async(blob)=>globalThis.JSZip.loadAsync(Buffer.from(await blob.arrayBuffer()));

test('sample Word OOXML analysis finds table activity and objective anchors',async()=>{
  const file=await openDocx(original);
  assert.equal(file.info.counts.tables,3);
  assert.ok(file.info.counts.equations>=0);
  assert.ok(file.snapshot.anchors.some(x=>x.id==='goals-competency'));
  assert.ok(file.snapshot.anchors.some(x=>x.kind==='activity' && x.column.includes('GV')));
  assert.ok(reportContext(file).activities.length>0);
});

test('editing existing DOCX adds paragraphs without modifying all other ZIP parts',async()=>{
  const file=await openDocx(original);
  const anchor=file.snapshot.anchors.find(x=>x.kind==='activity' && x.activity.includes('Hoạt động 2') && x.step==='Bước 1')||file.snapshot.anchors.find(x=>x.kind==='activity');
  assert.ok(anchor);
  const proposals=[
    {type:'NLAI',code:'9.TEST.1',content:'GV hướng dẫn HS đối chiếu kết quả của công cụ AI với cách làm của nhóm để đánh giá tính đúng đắn.',anchorId:anchor.id,accepted:true},
    {type:'HSHN',code:'',content:'GV chia nhiệm vụ thành các bước ngắn phù hợp và sử dụng phiếu gợi ý.',anchorId:'',reason:'Bước cần chèn không tìm thấy.',accepted:true}
  ];
  const result=await exportDocx(file,proposals);
  assert.equal(result.summary.inserted,1);assert.equal(result.summary.unplaced,1);
  const origZip=await globalThis.JSZip.loadAsync(original);
  const edited=await readDoc(result.blob);
  for(const name of Object.keys(origZip.files).filter(k=>!k.endsWith('/')&&k!=='word/document.xml')){
    const [oldBytes,newBytes]=await Promise.all([origZip.file(name).async('nodebuffer'),edited.file(name).async('nodebuffer')]);
    assert.ok(oldBytes.equals(newBytes),name+' binary changed');
  }
  const xml=await edited.file('word/document.xml').async('string');
  assert.ok(xml.includes('9.TEST.1'));
  assert.ok(xml.includes('PHỤ LỤC'));
  fs.writeFileSync(new URL('./tmp-edited-sample.docx',import.meta.url),Buffer.from(await result.blob.arrayBuffer()));
});

test('second run prevents reinserting same NLS code inside the same activity',async()=>{
  const file=await openDocx(original);
  const anchor=file.snapshot.anchors.find(x=>x.kind==='activity' && x.activity.includes('Hoạt động 2') && x.step==='Bước 1');
  assert.ok(anchor);
  const result=await exportDocx(file,[{type:'NLS',code:'5.2.TC2c',content:'GV thêm mô phỏng tập nghiệm theo đúng chỉ số.',anchorId:anchor.id,accepted:true}]);
  assert.equal(result.summary.existing,1);
  assert.equal(result.summary.inserted,0);
});

test('table anchors follow semantic teacher column even when on the right',async()=>{
  const sample=fs.readFileSync(new URL('./right-column-fixture.docx',import.meta.url));
  const opened=await openDocx(sample);
  const target=opened.snapshot.anchors.find(x=>x.kind==='activity'&&x.step==='Bước 1');
  assert.ok(target);
  assert.match(target.column,/HĐ CỦA GV VÀ HS/);
  const answer=await exportDocx(opened,[{type:'NLS',code:'3.1.TC1a',content:'GV hướng dẫn HS sắp xếp thư mục và kiểm tra lại cấu trúc tệp.',anchorId:target.id,accepted:true}]);
  assert.equal(answer.summary.inserted,1);
});

test('existing HSHN activity is not duplicated',async()=>{
  const opened=await openDocx(original);
  const target=opened.snapshot.anchors.find(x=>x.kind==='activity'&&x.activity.includes('Hoạt động 2')&&x.step==='Bước 2');
  assert.ok(target);
  const answer=await exportDocx(opened,[{type:'HSHN',content:'GV hỗ trợ học sinh qua phiếu chia bước.',anchorId:target.id,accepted:true}]);
  assert.equal(answer.summary.existing,1);
});
