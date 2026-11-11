/** HSHN uses the actual objectives of the selected lesson, never a generic AI-written copy. */
export const DISABILITY_TYPES=Object.freeze([
  {id:'vision',label:'Khuyết tật nhìn (Thị giác)'},
  {id:'hearing',label:'Khuyết tật nghe (Thính giác)'},
  {id:'mobility',label:'Khuyết tật vận động'},
  {id:'intellectual',label:'Khuyết tật trí tuệ / Phát triển'},
  {id:'communication',label:'Khuyết tật ngôn ngữ / Giao tiếp'},
  {id:'other',label:'Khuyết tật khác'}
]);
export const shortTypes=ids=>DISABILITY_TYPES.filter(x=>(ids||[]).includes(x.id)).map(x=>x.label);
const fold=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/gi,'d').toLowerCase().replace(/\s+/g,' ').trim();
const bullet=s=>String(s||'').replace(/^[\s\-•+–]+/,'').trim();
const isHeading=s=>{
 const t=fold(s);
 return /^(?:[1-5](?:\.\d+)*\s*\.?\s*)?(?:(?:ve\s*)?(?:kien thuc|nang luc|pham chat)|nang luc chung|nang luc rieng|nang luc tin hoc)(?:\s*[:.]?\s*)$/.test(t)
 || /^(?:[1-5](?:\.\d+)*\s*\.?\s*)?(?:nang luc so|nang luc tri tue nhan tao)/.test(t)
 || /^nang luc [abcde]\s*\(nl[abcde]\)/.test(t);
};
const slice=xs=>xs.filter(Boolean).map(bullet);
const options={
 vision:'đọc hoặc trình bày thông tin bằng lời, học liệu tương phản/chữ lớn hoặc dạng tiếp cận phù hợp với nhu cầu thực tế',
 hearing:'hướng dẫn bằng văn bản ngắn và sơ đồ rõ ràng, kiểm tra việc hiểu nhiệm vụ thay vì chỉ giao việc bằng lời',
 mobility:'bố trí chỗ ngồi và phương tiện thao tác thuận tiện, phân công bạn hỗ trợ khi học sinh mong muốn',
 intellectual:'chia nhiệm vụ thành bước ngắn, mẫu làm sẵn và giảm số thao tác đồng thời',
 communication:'cho phép trả lời bằng cách chọn, chỉ, viết hoặc công cụ giao tiếp phù hợp',
 other:'dựa vào kế hoạch giáo dục cá nhân, không suy đoán mức độ hỗ trợ'
};
export const hskSupportDescription=types=>(types?.length?types:['other']).map(x=>options[x]).filter(Boolean).join('; ');
export function goalSource(opened,{unitId}={}){
  let blocks=(opened?.snapshot?.blocks||[]).filter(b=>b.table===-1&&b.section==='I'&&(!unitId||b.unitId===unitId));
  const cut=blocks.findIndex(b=>/muc tieu danh cho hoc sinh khuyet tat|dieu chinh muc tieu.*hoa nhap/.test(fold(b.text)));
  if(cut>=0)blocks=blocks.slice(0,cut);
  const ix=(re,start=0)=>blocks.findIndex((b,i)=>i>=start&&re.test(fold(b.text)));
  const iK=ix(/^1\s*\.\s*(?:ve\s*)?kien thuc\b/);
  const iC=ix(/^2\s*\.\s*(?:ve\s*)?nang luc\b/);
  const iP=ix(/^3\s*\.\s*(?:ve\s*)?pham chat\b/);
  if(iK<0||iC<=iK||iP<=iC) return {valid:false,reason:'Không xác định đủ ba phần Kiến thức – Năng lực – Phẩm chất của đúng bài học.',knowledge:[],general:[],special:[],qualities:[]};
  const isReal=s=>s.length>12&&!/^(hoc xong|thuc hien bai hoc nay|hoc sinh se|cac nang luc)/.test(fold(s));
  const knowledge=slice(blocks.slice(iK+1,iC).map(b=>b.text)).filter(s=>isReal(s)&&!isHeading(s)).slice(0,2);
  const competencies=blocks.slice(iC+1,iP).map(b=>b.text);
  const iG=competencies.findIndex(s=>/^(2\s*\.\s*1\s*\.?\s*)?nang luc chung/i.test(fold(s)));
  const iS=competencies.findIndex(s=>/^(2\s*\.\s*2\s*\.?\s*)?(?:nang luc tin hoc|nang luc rieng)/i.test(fold(s)));
  const iStop=competencies.findIndex((s,i)=>i>iS&&(/^(?:(?:2\s*\.\s*[34]\b|[345]\s*\.\s*)?nang luc so|(?:2\s*\.\s*[34]\b|[345]\s*\.\s*)?nang luc tri tue)/i.test(fold(s))));
  const general=slice(competencies.slice(iG>=0?iG+1:0,iS>iG?iS:competencies.length)).filter(s=>isReal(s)&&!isHeading(s))
    .filter(s=>/tu chu|tu hoc/i.test(fold(s))).slice(0,1);
  const rawSpecial=slice(competencies.slice(iS>=0?iS+1:0,iStop>iS?iStop:competencies.length));
  const normalSpecial=rawSpecial.filter(s=>isReal(s)&&!isHeading(s)&&!/^(nang luc so|nang luc tri tue|\d+(?:\.\d+)*\s*\.\s*tc|\d+\s*\.\s*[abcd]\d)/i.test(fold(s)));
  let special=normalSpecial.filter(s=>/tu duy|su dung cong cu|su dung va quan ly|phat trien nang luc su dung|thao tac|nhan biet/i.test(fold(s)));
  if(!special.length)special=normalSpecial;
  special=special.slice(0,2);
  const qualities=slice(blocks.slice(iP+1).map(b=>b.text)).filter(s=>isReal(s)&&!isHeading(s)).slice(0,2);
  return {valid:knowledge.length>0&&qualities.length>0,reason:knowledge.length&&qualities.length?'':'Không trích đủ yêu cầu nhận biết/phẩm chất từ mục tiêu gốc.',knowledge,general,special,qualities};
}
function simplifyCompetence(s,subject){
 const f=fold(s);
 if(subject==='Toán'&&f.includes('giai quyet duoc cac van de thuc tien')&&f.includes('the tich'))
   return 'Nhận biết cách sử dụng công thức tính thể tích và diện tích xung quanh của hình hộp chữ nhật, hình lập phương trong bài tập cơ bản.';
 if(subject==='Tin'&&f.includes('thao tac thanh thao voi tep va thu muc'))
   return 'Thực hiện được một thao tác tệp hoặc thư mục đơn giản theo các bước hướng dẫn.';
 return s;
}
export function inclusiveObjectives(opened,lesson,{types=[],digital=false,unitId}={}){
  const src=goalSource(opened,{unitId});
  if(!src.valid)throw new Error(src.reason);
  const who=shortTypes(types).join('; ')||'chưa chọn loại khuyết tật – giáo viên xác định biện pháp cụ thể';
  const subject=lesson?.subject||'Toán';
  let g=src.general[0]||'Tự chủ và tự học thông qua nhiệm vụ cơ bản có hướng dẫn.';
  if(subject==='Tin'&&g.length>145)g='Năng lực tự học: Đọc SGK và trả lời câu hỏi nhận biết của bài học với gợi ý của giáo viên.';
  const specials=(digital?src.special:src.special.slice(0,1)).length?(digital?src.special:src.special.slice(0,1)):[subject==='Tin'?'Thực hiện thao tác cơ bản của bài học theo hướng dẫn.':'Nhận biết và sử dụng kiến thức toán học cơ bản trong nhiệm vụ được giao.'];
  const fact=src.knowledge[0].replace(/[.;:]$/,'').replace(/^(HS|Học sinh)\s+/i,'');
  const supports=subject==='Tin'
    ?`${fact} với ví dụ, tranh hoặc thiết bị trực quan; thực hiện một thao tác cốt lõi theo từng bước với sự hỗ trợ khi cần.`
    :`${fact} với nhiệm vụ vừa sức; dùng máy tính cầm tay để kiểm tra phép tính hoặc thử giá trị đơn giản khi phù hợp.`;
  return [
    '* MỤC TIÊU DÀNH CHO HỌC SINH KHUYẾT TẬT',
    '1. Kiến thức:',...src.knowledge.map(x=>'- '+x),
    '2. Năng lực',
    'Năng lực chung:', '- '+g,
    subject==='Tin'?'Năng lực riêng (Tin học):':'Năng lực riêng:',...specials.map(x=>'- '+simplifyCompetence(x,subject)),
    '3. Phẩm chất',...src.qualities.map(x=>'- '+x),
    `* Điều chỉnh mục tiêu Giáo dục hòa nhập (Cho HSKT ${who}):`,
    `+ [GDHN.KT]: ${supports}${digital?' HSKT có thể tham gia nhiệm vụ số bổ sung nếu có phương tiện tiếp cận và hỗ trợ phù hợp.':''}`,
    '+ [GDHN.KN]: Rèn luyện sự tự tin, chủ động tương tác với bạn đồng hành, tham gia hoạt động và trình bày kết quả bằng hình thức phù hợp.'
  ].join('\n');
}
export function inclusiveActivity(lesson,{types=[],digital=false}={},context=''){
  const accommodations=hskSupportDescription(types),s=fold(context+' '+(lesson?.name||'')),subject=lesson?.subject||'Toán';
  if(subject==='Tin'){
    const task=/thiet bi vao|thiet bi ra/.test(s)?'nhận biết và phân loại chuột, bàn phím, màn hình hoặc loa trong bài SGK'
      :/he dieu hanh|phan mem/.test(s)?'nhận biết biểu tượng hoặc chọn ví dụ phần mềm/hệ điều hành trong SGK'
      :/du lieu|tep|thu muc/.test(s)?'nhận diện tệp, thư mục hoặc chọn thao tác đúng trong bài SGK'
      :'thực hiện một nhiệm vụ cơ bản từ bài học Tin học theo phiếu minh họa';
    return digital
      ?`GV/bạn đồng hành điều chỉnh bàn phím, chuột và giao diện tiếp cận; ${accommodations}. HSKT tham gia thao tác số phù hợp để ${task}, có thể dùng hình thức trả lời thay thế.`
      :`GV/bạn đồng hành hỗ trợ: ${accommodations}. HSKT ${task}, được trả lời bằng chỉ hình, ghép thẻ hoặc thao tác tối thiểu theo yêu cầu cốt lõi của môn; không giao thêm chỉ số NLS riêng.`;
  }
  const task=/hinh hop|hinh lap phuong|the tich|dien tich/.test(s)
    ?'nhận biết hình hộp hoặc hình lập phương bằng mô hình và tính một giá trị đơn giản từ số đo trong SGK'
    :'đọc đề, xác định dữ kiện và thay số nguyên đơn giản vào biểu thức hoặc phương trình của bài SGK';
  return digital
   ?`GV/bạn đồng hành hỗ trợ HSKT thao tác chuột, bàn phím hoặc tiếp cận thông tin màn hình theo nhu cầu; ${accommodations}. Đồng thời HS ${task}, sử dụng máy tính cầm tay để kiểm tra khi phù hợp.`
   :`GV/bạn đồng hành hỗ trợ HSKT: ${accommodations}. HS ${task}; có thể sử dụng máy tính cầm tay để kiểm tra, không yêu cầu học sinh thực hiện thêm nhiệm vụ năng lực số riêng.`;
}
export function checkInclusiveOutput(content,{digital=false}={},subject='Toán'){
  const s=String(content||'');
  if(!digital&&subject==='Toán'&&/(GeoGebra|Desmos|Excel|Google Sheets|Internet|trang web|website|chatbot|thao tác phần mềm)/i.test(s))return false;
  return true;
}

export function safeInclusiveObjectives(opened,lesson,settings={}){
 try{return {content:inclusiveObjectives(opened,lesson,settings),reason:''};}
 catch(e){return {content:'* MỤC TIÊU DÀNH CHO HỌC SINH KHUYẾT TẬT\n[CẦN GIÁO VIÊN BỔ SUNG: không trích được mục tiêu gốc phù hợp. Không tự tạo mục tiêu thay cho giáo viên.]',reason:e.message||'Không xác định được mục tiêu gốc.'};}
}
