import {normalize} from './docx-engine.js';
const significant = s => normalize(s).split(' ').filter(x=>x.length>=3 && !['bai','chuong','tiet','tap','hoc','toan','tin','lop'].includes(x));
export const lessonKey = x => `${x.subject}|${x.grade}|${x.track}|${x.name}|${x.periods.join(',')}`;
export function getTitleMetadata(opened) {
  const initial=opened.snapshot.blocks.filter(x=>x.table<0).slice(0,75).map(x=>x.text).join('\n');
  const titleLine=initial.split('\n').find(x=>/\bBài\s*\d+[a-z]?\s*[.:]/i.test(x))||'';
  const title=titleLine.replace(/^.*?(?=Bài\s*\d+[a-z]?\s*[.:])/i,'').slice(0,190);
  const num=title.match(/Bài\s*(\d+[a-z]?)/i)?.[1]?.toLowerCase()||'';
  const header=opened.snapshot.blocks.filter(x=>x.table<0).slice(0,15).map(x=>x.text).join(' ');
  const headNorm=normalize(header);
  let subject=/(?:^| )toan(?: |$)/.test(headNorm)?'Toán':/(?:^| )tin(?: hoc)?(?: |$)/.test(headNorm)?'Tin':'';
  let grade=+(header.match(/\b(?:toán|tin(?:\s+học)?)\s*([6-9])\b/i)?.[1]||0);
  if(!grade){const m=opened.info.name.match(/(?:toan|tin|DS)\s*([6-9])/i);if(m)grade=+m[1];}
  if(!subject){if(/Toán|DS\d/i.test(opened.info.name))subject='Toán';if(/Tin/i.test(opened.info.name))subject='Tin';}
  return {title,num,grade,subject,initial};
}
export function ranking(catalog,meta,subject,grade){
  const selected=catalog.lessons.filter(x=>x.subject===subject&&x.grade===+grade);
  const titleWords=significant(meta.title);
  return selected.map(x=>{
    const name=normalize(x.name);
    const current=x.name.match(/Bài\s*(\d+[a-z]?)/i)?.[1]?.toLowerCase()||'';
    const nameWords=significant(x.name);
    let score=titleWords.length?nameWords.filter(w=>titleWords.includes(w)).length/Math.max(1,new Set([...titleWords,...nameWords]).size):0;
    if(meta.num&&current){if(meta.num===current)score+=1.8;else score-=1;}
    if(meta.title&&name.includes(normalize(meta.title).slice(0,35)))score+=0.8;
    return {lesson:x,score};
  }).sort((a,b)=>b.score-a.score);
}
export function sourceList(lesson,types){
  const list=[];
  if(types.includes('NLS'))list.push(...lesson.nls.map(x=>({...x,type:'NLS'})));
  if(types.includes('NLAI'))list.push(...lesson.ai.map(x=>({...x,type:'NLAI'})));
  return list;
}
export function offlineProposals(lesson,opened,types){
  const anchors=opened.snapshot.anchors;
  const firstGoal=anchors.find(x=>x.id==='goals-competency');
  const hGoal=anchors.find(x=>x.id==='goals-inclusion');
  const act=anchors.find(x=>x.kind==='activity' && x.step==='Bước 1');
  const act2=anchors.find(x=>x.kind==='activity' && x.step==='Bước 2')||act;
  const res=[];
  for(const type of ['NLS','NLAI']) {
    if(!types.includes(type))continue;
    const codes=(type==='NLS'?lesson.nls:lesson.ai).slice(0,2);
    for(const item of codes){
      const raw=item.text.replace(/^Học sinh\s*/i,'HS ').replace(/^HS\s*/i,'HS ');
      res.push({type,code:item.code,anchorId:firstGoal?.id||'',content:raw||`Phát triển ${type} ${item.code} trong bài học.`,reason:'Không tìm thấy mục tiêu năng lực phù hợp.',accepted:true});
      const target=type==='NLS'?act:act2;
      res.push({type,code:item.code,anchorId:target?.id||'',content:`GV giao nhiệm vụ cho HS: ${raw.charAt(0).toLowerCase()+raw.slice(1)}`,reason:'Không tìm được bước thực hiện phù hợp với hoạt động.',accepted:true});
    }
  }
  if(types.includes('HSHN')){
    res.push({type:'HSHN',code:'',anchorId:hGoal?.id||'',content:`HS được hỗ trợ đạt yêu cầu trọng tâm của bài ${lesson.name} thông qua nhiệm vụ chia nhỏ và hướng dẫn từng bước; mức điều chỉnh do GV quyết định.`,reason:'Không có tiểu mục mục tiêu giáo dục hòa nhập.',accepted:true});
    res.push({type:'HSHN',code:'',anchorId:act2?.id||'',content:'GV cung cấp phiếu có gợi ý từng bước, cho HS thêm thời gian, bố trí bạn học hỗ trợ và đánh giá theo khả năng thực hiện thực tế.',reason:'Không xác định được Bước 2 hoặc hoạt động phù hợp.',accepted:true});
  }
  return res;
}
