/** Minimal, strict LaTeX -> Word OMML converter for classroom algebra.
 * Unsupported LaTeX fails loudly rather than exporting literal dollar signs.
 * Already existing OMML equations in the source DOCX are NEVER rewritten here.
 */
export const MATH_NS='http://schemas.openxmlformats.org/officeDocument/2006/math';
export const WORD_NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NS='http://www.w3.org/XML/1998/namespace';
const symbols={le:'≤',leq:'≤',ge:'≥',geq:'≥',neq:'≠',ne:'≠',times:'×',cdot:'·',pm:'±',mp:'∓',in:'∈',notin:'∉',subset:'⊂',subseteq:'⊆',cup:'∪',cap:'∩',div:'÷',to:'→',Rightarrow:'⇒',implies:'⇒',approx:'≈',equiv:'≡',infty:'∞',ldots:'…',cdots:'⋯',forall:'∀',exists:'∃',land:'∧',lor:'∨',alpha:'α',beta:'β',gamma:'γ',delta:'δ',Delta:'Δ',pi:'π',theta:'θ',lambda:'λ',mu:'μ',omega:'ω',Omega:'Ω',sum:'∑',int:'∫',angle:'∠',perp:'⊥',parallel:'∥',circ:'°',triangle:'△',degree:'°',prime:'′',qquad:'  ',quad:' '};
const el=(doc,local,attributes={})=>{const n=doc.createElementNS(MATH_NS,'m:'+local);for(const [k,v] of Object.entries(attributes))n.setAttributeNS(MATH_NS,'m:'+k,v);return n;};
function run(doc,text){const r=el(doc,'r'),t=el(doc,'t');t.setAttributeNS(XML_NS,'xml:space','preserve');t.appendChild(doc.createTextNode(text));r.appendChild(t);return r;}
const append=(n,arr)=>{for(const a of arr)n.appendChild(a);return n;};
function box(doc,name,arr){return append(el(doc,name),arr);}
function mathFrac(doc,a,b){const n=el(doc,'f');n.appendChild(box(doc,'num',a));n.appendChild(box(doc,'den',b));return n;}
function mathRad(doc,e,degree){const n=el(doc,'rad'),pr=el(doc,'radPr');pr.appendChild(el(doc,'degHide',{val:degree?.length?'0':'1'}));n.appendChild(pr);n.appendChild(box(doc,'deg',degree||[]));n.appendChild(box(doc,'e',e));return n;}
function mathScripts(doc,base,sub,sup){if(sub&&sup){const s=el(doc,'sSubSup');s.appendChild(box(doc,'e',[base]));s.appendChild(box(doc,'sub',sub));s.appendChild(box(doc,'sup',sup));return s;}if(sub){const s=el(doc,'sSub');s.appendChild(box(doc,'e',[base]));s.appendChild(box(doc,'sub',sub));return s;}const s=el(doc,'sSup');s.appendChild(box(doc,'e',[base]));s.appendChild(box(doc,'sup',sup));return s;}
class Reader {
  constructor(source,doc){this.s=source.trim();this.i=0;this.doc=doc;this.depth=0;}
  ws(){while(this.i<this.s.length&&/\s/.test(this.s[this.i]))this.i++;}
  group(){this.ws();if(this.s[this.i]==='{'){this.i++;const got=this.seq('}');if(this.s[this.i]!=='}')throw Error('Thiếu dấu } trong LaTeX');this.i++;return got;}return this.atom();}
  atom(){this.ws();if(this.i>=this.s.length)return[];const c=this.s[this.i];if(c==='{' )return this.group();if(c==='}'||c===']')return[];
    if(c==='\\'){
      this.i++;const start=this.i;while(/[A-Za-z]/.test(this.s[this.i]||''))this.i++;let command=this.s.slice(start,this.i);
      if(!command){command=this.s[this.i++]||'';if([',',';',':','!',' '].includes(command))return [run(this.doc,' ')];if(command==='{')return[run(this.doc,'{')];if(command==='}')return[run(this.doc,'}')];if(command==='\\')return[run(this.doc,' ')];throw Error('Lệnh LaTeX không hỗ trợ: \\'+command);}
      if(['left','right','big','Big','bigg','Bigg'].includes(command))return[];
      if(['frac','dfrac','tfrac'].includes(command))return[mathFrac(this.doc,this.group(),this.group())];
      if(command==='sqrt'){
        this.ws();let degree=[];if(this.s[this.i]==='['){this.i++;degree=this.seq(']');if(this.s[this.i]!==']')throw Error('Thiếu ] sau căn');this.i++;}
        return[mathRad(this.doc,this.group(),degree)];
      }
      if(command==='overline'||command==='bar'){
        const n=el(this.doc,'bar');const pr=el(this.doc,'barPr');pr.appendChild(el(this.doc,'pos',{val:'top'}));n.appendChild(pr);n.appendChild(box(this.doc,'e',this.group()));return[n];
      }
      if(command==='vec'){
        const n=el(this.doc,'acc');const pr=el(this.doc,'accPr');pr.appendChild(el(this.doc,'chr',{val:'⃗'}));n.appendChild(pr);n.appendChild(box(this.doc,'e',this.group()));return[n];
      }
      if(['mathrm','mathbf','mathit','text','operatorname','mathbb'].includes(command)){
        if(command==='mathbb'){const group=this.group();const text=group.map(n=>n.textContent).join('');if(text==='R')return[run(this.doc,'ℝ')];if(text==='N')return[run(this.doc,'ℕ')];if(text==='Z')return[run(this.doc,'ℤ')];if(text==='Q')return[run(this.doc,'ℚ')];return group;}
        return this.group();
      }
      if(symbols[command])return[run(this.doc,symbols[command])];
      throw Error('Lệnh LaTeX không hỗ trợ: \\'+command);
    }
    if(c==='^'||c==='_')throw Error('Chỉ số dưới/trên không có biểu thức cơ sở');
    if(c==='&')throw Error('Cú pháp bảng/matrix không được hỗ trợ tại đây');
    this.i++;return[run(this.doc,c)];
  }
  seq(end){if(++this.depth>45)throw Error('LaTeX lồng nhau quá sâu');const result=[];
    while(this.i<this.s.length){this.ws();if(this.i>=this.s.length||this.s[this.i]===end)break;
      const start=this.i,atoms=this.atom();if(this.i===start)throw Error('Cú pháp LaTeX không hợp lệ');
      for(let base of atoms){let sub=null,sup=null;let next=this.s[this.i];
        while(next==='^'||next==='_'){this.i++;const script=this.group();if(next==='^')sup=script;else sub=script;this.ws();next=this.s[this.i];}
        result.push(sub||sup?mathScripts(this.doc,base,sub,sup):base);
      }
    }
    this.depth--;return result;
  }
  parse(){const out=this.seq('');if(this.i!==this.s.length)throw Error('Còn ký tự LaTeX chưa xử lý: '+this.s.slice(this.i,30));if(!out.length)throw Error('Công thức LaTeX rỗng');return out;}
}
export function latexToOmml(doc,latex,color){
  if(!latex||latex.length>1200)throw Error('Công thức LaTeX rỗng hoặc dài quá 1200 ký tự');
  const root=el(doc,'oMath');
  const cases=/^\s*\\begin\{cases\}([\s\S]*?)\\end\{cases\}\s*$/.exec(latex);
  if(cases){
    const lines=cases[1].split(/\\\\/).map(x=>x.replace(/&/g,'').trim()).filter(Boolean);
    if(!lines.length||lines.length>8)throw Error('Hệ phương trình cần 1-8 dòng');
    const delim=el(doc,'d'),dp=el(doc,'dPr');dp.appendChild(el(doc,'begChr',{val:'{'}));dp.appendChild(el(doc,'endChr',{val:''}));delim.appendChild(dp);
    const expr=el(doc,'e'),array=el(doc,'eqArr');
    for(const line of lines)array.appendChild(box(doc,'e',new Reader(line,doc).parse()));
    expr.appendChild(array);delim.appendChild(expr);root.appendChild(delim);if(color)applyMathColor(doc,root,color);return root;
  }
  append(root,new Reader(latex,doc).parse());if(color)applyMathColor(doc,root,color);return root;
}
function applyMathColor(doc,root,color){
  const all=Array.from(root.getElementsByTagNameNS(MATH_NS,'r'));
  for(const mathRun of all){
    const rp=doc.createElementNS(WORD_NS,'w:rPr');
    const col=doc.createElementNS(WORD_NS,'w:color');col.setAttributeNS(WORD_NS,'w:val',color);
    const fonts=doc.createElementNS(WORD_NS,'w:rFonts');for(const key of ['ascii','hAnsi','eastAsia','cs'])fonts.setAttributeNS(WORD_NS,'w:'+key,'Times New Roman');
    rp.appendChild(fonts);rp.appendChild(col);
    mathRun.insertBefore(rp,mathRun.firstChild);
  }
}
export function richFragments(content){const value=String(content||'').replace(/<br\s*\/?\s*>/gi,'\n').replace(/&lt;br\s*\/?&gt;/gi,'\n');
  const pieces=[];let start=0;const re=/(?<!\\)\$([^$\n]+)(?<!\\)\$/g;let m;
  while((m=re.exec(value))){if(m.index>start)pieces.push({kind:'text',value:value.slice(start,m.index)});pieces.push({kind:'math',value:m[1]});start=re.lastIndex;}
  if(start<value.length)pieces.push({kind:'text',value:value.slice(start)});
  if(pieces.filter(x=>x.kind==='text').some(x=>/(?<!\\)\$/.test(x.value)))throw Error('Công thức thiếu dấu $ đóng hoặc chứa xuống dòng.');
  return pieces;
}
export function addRichRuns(doc,p,text,color='2563EB',fontSize='26'){
  const pieces=richFragments(text);let count=0;
  for(const piece of pieces){if(piece.kind==='math'){p.appendChild(latexToOmml(doc,piece.value,color));count++;continue;}
    if(!piece.value)continue;const r=doc.createElementNS(WORD_NS,'w:r');const rpr=doc.createElementNS(WORD_NS,'w:rPr');
    const font=doc.createElementNS(WORD_NS,'w:rFonts');for(const key of ['ascii','hAnsi','eastAsia','cs'])font.setAttributeNS(WORD_NS,'w:'+key,'Times New Roman');
    const col=doc.createElementNS(WORD_NS,'w:color');col.setAttributeNS(WORD_NS,'w:val',color);rpr.appendChild(font);rpr.appendChild(col);const sz=doc.createElementNS(WORD_NS,'w:sz');sz.setAttributeNS(WORD_NS,'w:val',fontSize);rpr.appendChild(sz);r.appendChild(rpr);
    const t=doc.createElementNS(WORD_NS,'w:t');t.setAttributeNS(XML_NS,'xml:space','preserve');t.appendChild(doc.createTextNode(piece.value));r.appendChild(t);p.appendChild(r);
  }return count;
}
