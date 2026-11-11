import {safeInclusiveObjectives,inclusiveActivity} from './inclusion-policy.js';
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
    const current=x.name.match(/Bài\s*(\d+[a-z]?)/i)?.[1]?.toLowerCase().replace(/^0+(?=\d)/,'')||'';
    const nameWords=significant(x.name);
    let score=titleWords.length?nameWords.filter(w=>titleWords.includes(w)).length/Math.max(1,new Set([...titleWords,...nameWords]).size):0;
    if(meta.num&&current){if(String(meta.num).replace(/^0+(?=\d)/,'')===current)score+=1.8;else score-=1;}
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
export function offlineProposals(lesson,opened,types, inclusive={}){
  const unitId=inclusive.unitId||opened.snapshot.units?.[0]?.id;
  const anchors=opened.snapshot.anchors.filter(a=>!unitId||a.unitId===unitId);
  const firstGoal=anchors.find(x=>x.id.startsWith('goals-competency'));
  const hGoal=anchors.find(x=>x.id.startsWith('goals-inclusion'));
  const act=anchors.find(x=>x.kind==='activity' && x.step==='Bước 1');
  const act2=anchors.find(x=>x.kind==='activity' && x.step==='Bước 2');
  const tail=anchors.find(x=>x.kind==='activity-end')||null;
  const home=anchors.find(x=>x.kind==='homework')||null;
  const res=[];
  for(const type of ['NLS','NLAI']) {
    if(!types.includes(type))continue;
    const codes=(type==='NLS'?lesson.nls:lesson.ai).slice(0,2);
    for(const item of codes){
      const raw=item.text.replace(/^Học sinh\s*/i,'HS ').replace(/^HS\s*/i,'HS ');
      res.push({type,code:item.code,anchorId:firstGoal?.id||'',content:raw||`Phát triển ${type} ${item.code} trong bài học.`,reason:'Không tìm thấy mục tiêu năng lực phù hợp.',accepted:true,slot:'goal'});
      const target=act; // NLS and AI tasks belong to Bước 1, not Bước 2.
      res.push({type,code:item.code,anchorId:target?.id||'',content:`GV giao nhiệm vụ cho HS: ${raw.charAt(0).toLowerCase()+raw.slice(1)}`,reason:'Không tìm được bước thực hiện phù hợp với hoạt động.',accepted:true,slot:'activity'});
      if(type==='NLS' && item===codes[0]){
        res.push({type,code:'',anchorId:tail?.id||'',content:'HS trình bày kết quả nhiệm vụ NLS ngay trong bài tập SGK và nêu cách đã kiểm chứng.',reason:'Không xác định được cuối Bước 1.',accepted:true,slot:'step1-summary'});
        res.push({type,code:'',anchorId:home?.id||'',content:'Ôn lại bài tập SGK và thực hành một nhiệm vụ số ngắn phù hợp điều kiện thiết bị của lớp.',reason:'Không xác định mục Hướng dẫn về nhà.',accepted:true,slot:'homework'});
      }
    }
  }
  if(types.includes('HSHN')){
    const hGoals=safeInclusiveObjectives(opened,lesson,inclusive);
    res.push({type:'HSHN',code:'',anchorId:hGoals.reason?'':hGoal?.id||'',content:hGoals.content,reason:hGoals.reason||(!hGoal?'Không xác định đúng mục Phẩm chất.':''),accepted:true,slot:'goal'});
    const seen=new Set();
    for(const step2 of anchors.filter(a=>a.kind==='activity'&&a.step==='Bước 2')){
      if(seen.has(step2.activity))continue;seen.add(step2.activity);
      res.push({type:'HSHN',code:'',anchorId:step2.id,content:inclusiveActivity(lesson,inclusive,step2.activity),reason:'',accepted:true,slot:'activity'});
    }
    if(!seen.size)res.push({type:'HSHN',code:'',anchorId:'',content:inclusiveActivity(lesson,inclusive),reason:'Không xác định được Bước 2 trong cột Tổ chức thực hiện của bài.',accepted:true,slot:'activity'});
  }
  return res;
}
