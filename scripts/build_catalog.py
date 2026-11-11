"""Build immutable lesson catalogue from the source plans supplied by the teacher.
Usage: python scripts/build_catalog.py (run from project root in this environment).
Source files are deliberately NOT copied into the public web app.
"""
import json,re,unicodedata,os
from pathlib import Path
from docx import Document
import fitz
SRC=Path('/mnt/data')
OUT=Path('public/data/catalog.json')
NLS_RE=re.compile(r'(?<![\w.])\b(?:[1-6])\s*\.\s*[1-6]\s*\.?\s*TC\s*[12]\s*[a-d]\b',re.I)
AI_RE=re.compile(r'(?<!\w)\b[6-9]\s*\.\s*[A-D]\s*\d+(?:\s*\.\s*(?:MR)?\d+)\b',re.I)

def norm(s):
 s=unicodedata.normalize('NFD',s.lower())
 s=''.join(c for c in s if unicodedata.category(c)!='Mn')
 return re.sub(r'\W+',' ',s).strip()

def parse_codes(txt,typ):
 txt=re.sub(r'\s+',' ',txt.strip())
 if not txt:return []
 rx=NLS_RE if typ=='nls' else AI_RE
 matches=list(rx.finditer(txt))
 res=[]
 for i,m in enumerate(matches):
  code=re.sub(r'\s+','',m.group()).upper()
  descr=txt[m.end():matches[i+1].start() if i+1<len(matches) else len(txt)].lstrip(' .:-–*')
  if len(descr)>420:descr=descr[:420].rsplit(' ',1)[0]+'…'
  if not any(x['code']==code and x['text']==descr for x in res):res.append({'code':code,'text':descr})
 return res

def periods(txt,n):
 # PPCT column may contain e.g. '3.4' typo for '3,4'; allow spaced and comma
 nums=[int(v) for v in re.findall(r'\d+',str(txt))]
 if len(nums)<=n+1:return nums[:n] if n and len(nums)>n else nums
 return nums[:min(8,len(nums))]

def split_lesson(title):
 s=title.replace('\n',' ')
 m=re.search(r'\bBài\s*\d+[a-z]?\b[^\n]*',s,re.I)
 if m:return re.sub(r'\s+',' ',m.group(0)).strip()
 return re.sub(r'\s+',' ',s).strip()

def build_tin(grade):
 matches=list(SRC.glob(f'KHDH*TIN*{grade}*cá nhân.docx'))
 if not matches:matches=[x for x in SRC.glob('*.docx') if 'cá nhân' in x.name.lower() and f'{grade}(' in x.name and 'tin' in x.name.lower()]
 assert matches,(grade,'no personal plan')
 d=Document(matches[0]);t=max(d.tables,key=lambda x:len(x.rows)); data=[]
 for row in t.rows[1:]:
  cells=[c.text.strip() for c in row.cells]
  if len(cells)<8:continue
  name=split_lesson(cells[1]);ppct=cells[3];num=int(re.findall(r'\d+',cells[2])[0]) if re.search(r'\d+',cells[2]) else 0
  if not name or not re.search(r'\d+',ppct):continue
  if norm(name).startswith(('hoc ki','kiem tra','on tap')):continue
  # Deduplicate repeating merged cells by actual text rather than column indexes
  nls_txt=cells[7]
  ai_txt=cells[-1]
  data.append(dict(subject='Tin',grade=grade,track='Tin học',name=name,periods=periods(ppct,num),nls=parse_codes(nls_txt,'nls'),ai=parse_codes(ai_txt,'ai'),source=matches[0].name,rawNote='KHDH cá nhân'))
 return data

def build_toan(grade):
 matches=list(SRC.glob(f'KHDH-TOÁN {grade}*.docx')); assert matches,grade
 d=Document(matches[0]);data=[]
 for t in d.tables:
  if len(t.columns)!=5 or len(t.rows)<10:continue
  track='Đại số' if not data else 'Hình học'
  for row in t.rows[1:]:
   cells=[c.text.strip() for c in row.cells]
   if not cells[1] or not re.search(r'\d+',cells[3]):continue
   if not re.search(r'\d',cells[0]) and 'bài' not in cells[1].lower() and 'luyện' not in cells[1].lower():continue
   num=int(re.findall(r'\d+',cells[2])[0]) if re.search(r'\d+',cells[2]) else 0
   text=cells[4].replace('\n',' ')
   # AI and NLS may be mixed; code match boundaries used below
   entry=dict(subject='Toán',grade=grade,track=track,name=split_lesson(cells[1]),periods=periods(cells[3],num),nls=parse_codes(text,'nls'),ai=parse_codes(text,'ai'),source=matches[0].name,rawNote='KHDH tổ')
   if entry['name']:data.append(entry)
 return data

def build_toan6():
 pdf=SRC/'KHDH Toan 6_26_27_Dieu chinh_8_9_2026.pdf'
 d=fitz.open(pdf);result=[];curr=None;track='Số học'
 for pno,page in enumerate(d):
  for table in page.find_tables().tables:
   rows=table.extract()
   if not rows or len(rows[0])<7:continue
   heading=rows[0]
   def header_index(term):
    return next((j for j,v in enumerate(heading) if v and term.lower() in str(v).lower()),-1)
   nameIdx=header_index('Bài học');perIdx=header_index('Tiết')
   if nameIdx<0 or perIdx<0:continue
   for cells in rows[1:]:
    values=[(x or '').strip() for x in cells]
    if any('mot so hinh hoc phang' in norm(v) for v in values): track='Hình học'
    ident=next((v for v in values[:3] if re.fullmatch(r'\d{1,3}',v)),None)
    # Find lesson title within a narrow band near the header's Bài học column.
    positions=range(max(0,nameIdx-1),min(len(values),nameIdx+3))
    named=[(j,values[j]) for j in positions if re.search(r'Bài\s*\d+|Luyện tập|Ôn tập|Kiểm tra|Trả bài|Hoạt động thực hành|Bài tập cuối', values[j],re.I)]
    if ident and named:
     j,name=named[0]
     # Period is generally the rightmost numeric cell between name and the first long description.
     candidates=[]
     for k,v in enumerate(values):
      if k<=j or k>min(len(values)-1,perIdx+3):continue
      if re.fullmatch(r'\d{1,3}(?:\s*[,.;]\s*\d{1,3})*',v):candidates.append((k,v))
     closest=sorted(candidates,key=lambda it:(abs(it[0]-perIdx),-it[0]))
     ptext=closest[0][1] if closest else ''
     per=[int(x) for x in re.findall(r'\d+',ptext)]
     curr=dict(subject='Toán',grade=6,track=track,name=split_lesson(name),periods=per,nls=[],ai=[],source=pdf.name,rawNote='KHDH tổ - PDF')
     result.append(curr)
    # A table's last text-bearing cell normally contains curricular outcomes.
    if curr:
     trailing=[(j,v) for j,v in enumerate(values) if j>perIdx and len(v)>35]
     if trailing:
      content=max(trailing,key=lambda item:item[0])[1]
      for typ in ['nls','ai']:
       for c in parse_codes(content,typ):
        if not any(x['code']==c['code'] for x in curr[typ]):curr[typ].append(c)
 return result

if __name__=='__main__':
 all_lessons=[]
 for g in (6,7,8,9):all_lessons+=build_tin(g)
 all_lessons+=build_toan6()
 for g in (7,8,9):all_lessons+=build_toan(g)
 # Teacher-approved example: these indicators are present in the supplied lesson-plan sample,
 # but NOT in the Toán 9 KHDH entry. Keep the provenance explicit.
 for lesson in all_lessons:
  if lesson['subject']=='Toán' and lesson['grade']==9 and lesson['name'].lower().startswith('bài 6. bất phương trình bậc nhất một ẩn') and lesson['track']=='Đại số':
   lesson['referenceSource']='Giáo án mẫu DS9_T9_T23_24_B6.docx (tham khảo; không phải mã trích từ KHDH Toán 9)'
   lesson['nls']=[
    {'code':'5.2.TC2c','text':'Biết điều chỉnh môi trường số để biểu diễn trực quan tập nghiệm bất phương trình trên trục số ảo.','provenance':'sample'},
    {'code':'3.4.TC2a','text':'Liệt kê các hướng dẫn logic cho máy tính thực hiện các bước giải bất phương trình bậc nhất một ẩn.','provenance':'sample'},
    {'code':'4.3.TC2a','text':'Giải thích cách tránh mỏi mắt, điều chỉnh độ sáng và khoảng cách nhìn khi làm bài trên thiết bị số.','provenance':'sample'}
   ]
   lesson['ai']=[{'code':'9.D1.1','text':'Nhận biết vai trò quyết định của con người khi kiểm chứng kết quả giải toán do AI đưa ra; phát hiện lỗi quên đổi chiều bất đẳng thức khi chia cho số âm.','provenance':'sample'}]
   break
 out=dict(version='2026-2027',school='THCS Nam Nha Trang',lessons=all_lessons)
 OUT.parent.mkdir(parents=True,exist_ok=True)
 OUT.write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
 for su in ('Tin','Toán'):
  for g in (6,7,8,9):
   subset=[x for x in all_lessons if x['subject']==su and x['grade']==g]
   print(su,g,len(subset),'NLS',sum(len(x['nls']) for x in subset),'AI',sum(len(x['ai']) for x in subset),'with periods',sum(bool(x['periods']) for x in subset))
 print('Total',len(all_lessons),'file',OUT.stat().st_size)
