/** Classroom accommodations: not a medical diagnosis, never infer specific pupil details. */
export const DISABILITY_TYPES=Object.freeze([
  {id:'vision',label:'Khuyết tật nhìn (Thị giác)'},
  {id:'hearing',label:'Khuyết tật nghe (Thính giác)'},
  {id:'mobility',label:'Khuyết tật vận động'},
  {id:'intellectual',label:'Khuyết tật trí tuệ / Phát triển'},
  {id:'communication',label:'Khuyết tật ngôn ngữ / Giao tiếp'},
  {id:'other',label:'Khuyết tật khác'}
]);
export const shortTypes=ids=>DISABILITY_TYPES.filter(x=>ids.includes(x.id)).map(x=>x.label);
const options={
  vision:'học liệu chữ lớn hoặc đọc bằng giọng nói, diễn giải thông tin trên đồ thị bằng lời; không mặc định phóng to là phù hợp với mọi HSKT nhìn',
  hearing:'hướng dẫn bằng chữ rõ ràng, hình minh họa và kiểm tra việc hiểu nhiệm vụ; không chỉ hướng dẫn bằng lời',
  mobility:'bố trí thiết bị/đồ dùng phù hợp và bạn hỗ trợ thao tác khi người học cần, không mặc định HSKT không thể tự dùng máy',
  intellectual:'chia bài tập thành từng bước ngắn với ví dụ mẫu, số liệu đơn giản và thời gian bổ sung',
  communication:'cho phép diễn đạt bằng cách chỉ, chọn, viết hoặc công cụ giao tiếp phù hợp',
  other:'điều chỉnh dựa vào nhu cầu thực tế và kế hoạch giáo dục cá nhân, tránh tự suy đoán tình trạng'
};
export function hskSupportDescription(types=[]){return (types.length?types:['other']).map(x=>options[x]).filter(Boolean).join('; ');}
function goalsFromDocx(opened){
  const blocks=opened?.snapshot?.blocks||[];
  let goals=blocks.filter(b=>b.table===-1&&b.section==='I');
  const hsk=goals.findIndex(b=>/MỤC TIÊU DÀNH CHO|Điều chỉnh mục tiêu.*hòa nhập/i.test(b.text));
  if(hsk>=0)goals=goals.slice(0,hsk);
  const knowledge=goals.filter(b=>/(nhận biết|nêu được|biết được|thực hiện được)/i.test(b.text))
    .filter(b=>!/Năng lực chung|Năng lực riêng|Năng lực số|Năng lực trí tuệ|Phẩm chất/i.test(b.text));
  const firstKnowledge=knowledge.slice(0,2).map(b=>b.text.replace(/^[•\-]\s*/,'').trim().slice(0,210));
  const qualities=goals.findIndex(b=>/^3\s*[.]\s*Phẩm chất/i.test(b.text));
  const qualitiesLines=qualities<0?[]:goals.slice(qualities+1,qualities+5).map(b=>b.text).filter(x=>/(tích cực|trách nhiệm)/i.test(x)).slice(0,2).map(x=>x.replace(/^[•\-]\s*/,''));
  return {firstKnowledge,qualitiesLines};
}
export function inclusiveObjectives(opened,lesson,{types=[],digital=false}={}){
  const g=goalsFromDocx(opened);
  const who=shortTypes(types).join('; ')||'theo nhu cầu hỗ trợ thực tế';
  const source=g.firstKnowledge.length?g.firstKnowledge:['Nhận biết kiến thức trọng tâm của bài học ở mức độ cơ bản.'];
  const knowledge=source.map(s=>'- '+s).join('\n');
  const pc=g.qualitiesLines.length?g.qualitiesLines:['Tích cực tham gia nhiệm vụ phù hợp.','Có tinh thần trách nhiệm với công việc được giao.'];
  const tool=lesson?.subject==='Toán'?'Sử dụng máy tính cầm tay để thử lại phép tính hoặc nghiệm đơn giản khi cần hỗ trợ.':'Thực hiện thao tác tin học cơ bản theo từng bước và khả năng tiếp cận của học sinh.';
  const aim=lesson?.subject==='Toán'?'Nhận biết khái niệm, công thức cơ bản; thử lại nghiệm hoặc phép tính đơn giản bằng máy tính cầm tay với sự hỗ trợ.':'Nhận biết khái niệm, thao tác hoặc quy trình cơ bản của bài; thực hành nhiệm vụ nhỏ với hỗ trợ phù hợp.';
  return [
    '* MỤC TIÊU DÀNH CHO HỌC SINH KHUYẾT TẬT',
    '1. Kiến thức:',knowledge,
    '2. Năng lực',
    '- Năng lực chung: Tự chủ và tự học trong nhiệm vụ vừa sức với sự hỗ trợ cần thiết.',
    '- Năng lực riêng: Tư duy và lập luận cơ bản; '+tool,
    '3. Phẩm chất',...pc.map(x=>'- '+x),
    '* Điều chỉnh mục tiêu Giáo dục hòa nhập (Cho HSKT '+who+'):',
    '+ [GDHN.KT]: '+aim+(digital?' Có thể tham gia nội dung NLS với công cụ hỗ trợ tiếp cận, không yêu cầu tự thao tác nếu chưa phù hợp.':''),
    '+ [GDHN.KN]: Rèn luyện sự tự tin, tương tác với bạn đồng hành và tham gia hoạt động cùng lớp bằng hình thức phù hợp.'
  ].join('\n');
}
export function inclusiveActivity(lesson,{types=[],digital=false}={}){
  const accommodations=hskSupportDescription(types);
  const task=lesson?.subject==='Toán'
   ?'HS đọc hoặc nghe đề theo khả năng, chọn số nguyên đơn giản để thay vào bài SGK và kiểm tra một bước bằng máy tính cầm tay khi cần.'
   :'HS thực hiện thao tác cơ bản của ví dụ/bài tập SGK theo phiếu bước ngắn, có thể trả lời bằng hình thức phù hợp.';
  if(digital)return 'GV/bạn đồng hành hỗ trợ HSKT truy cập và thao tác thiết bị số (chuột, bàn phím, phần mềm, hoặc mô tả bằng lời nội dung màn hình khi phù hợp); '+accommodations+'. Kết hợp nhiệm vụ đơn giản: '+task;
  return 'GV/bạn đồng hành hỗ trợ HSKT: '+accommodations+'. '+task+' Mục tiêu nhiệm vụ là hiểu bài và tương tác cùng lớp, không đánh giá HSKT qua hoạt động NLS.';
}
export function checkInclusiveOutput(content,{digital=false}={}){
  const s=String(content||'');
  if(!digital&&/(GeoGebra|Desmos|Excel|Google Sheets|Internet|trang web|website|chatbot|thao tác phần mềm|phần mềm đồ thị)/i.test(s))return false;
  return true;
}
