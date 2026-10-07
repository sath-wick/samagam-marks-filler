/* Self-contained XLSX reader. Supports multiple worksheets. */
(function(){
  const td=new TextDecoder('utf-8');
  const u16=(a,o)=>a[o]|a[o+1]<<8;
  const u32=(a,o)=>(a[o]|a[o+1]<<8|a[o+2]<<16|a[o+3]<<24)>>>0;
  function eocd(a){for(let i=a.length-22;i>=Math.max(0,a.length-65558);i--)if(u32(a,i)===0x06054b50)return i;throw Error('Invalid XLSX: ZIP end record not found.');}
  async function inflate(b){const ds=new DecompressionStream('deflate-raw');return new Uint8Array(await new Response(new Blob([b]).stream().pipeThrough(ds)).arrayBuffer())}
  async function zip(ab){const a=new Uint8Array(ab),e=eocd(a),size=u32(a,e+12),off=u32(a,e+16);let p=off;const out=new Map();while(p<off+size){if(u32(a,p)!==0x02014b50)throw Error('Invalid XLSX: ZIP directory error.');const m=u16(a,p+10),cs=u32(a,p+20),nl=u16(a,p+28),xl=u16(a,p+30),cl=u16(a,p+32),lo=u32(a,p+42),name=td.decode(a.slice(p+46,p+46+nl));if(!name.endsWith('/')){if(u32(a,lo)!==0x04034b50)throw Error('Invalid XLSX: local header error.');const lnl=u16(a,lo+26),lxl=u16(a,lo+28);let d=a.slice(lo+30+lnl+lxl,lo+30+lnl+lxl+cs);if(m===8)d=await inflate(d);else if(m!==0)throw Error('Unsupported ZIP compression method: '+m);out.set(name,d)}p+=46+nl+xl+cl}return out}
  const xml=b=>new DOMParser().parseFromString(td.decode(b),'application/xml');
  const all=(el,n)=>Array.from(el.getElementsByTagName(n).length?el.getElementsByTagName(n):el.getElementsByTagNameNS('*',n));
  const first=(el,n)=>el.getElementsByTagName(n)[0]||el.getElementsByTagNameNS('*',n)[0];
  const ci=r=>{let s=r.replace(/\d/g,''),n=0;for(const c of s)n=n*26+c.charCodeAt(0)-64;return n-1};
  const nh=s=>String(s??'').trim().replace(/\s+/g,' ').toLowerCase();
  function cv(c,shared){const t=c.getAttribute('t'),v=first(c,'v'),raw=v?v.textContent:'';if(t==='s')return shared[Number(raw)]??'';if(t==='inlineStr'){const is=first(c,'is');return is?all(is,'t').map(x=>x.textContent).join(''):''}if(t==='b')return raw==='1'?'TRUE':'FALSE';return raw}
  function resolvePath(target){target=target.replace(/^\//,'');return target.startsWith('xl/')?target:'xl/'+target}
  async function parseSheet(z,path,name,shared){if(!z.has(path))throw Error('Could not find worksheet: '+name);const doc=xml(z.get(path)),matrix=[];for(const row of all(doc,'row')){const r=[];for(const c of all(row,'c')){const ref=c.getAttribute('r')||'';r[ci(ref)]=cv(c,shared)}matrix.push(r)}if(!matrix.length)return null;const headers=matrix[0].map(x=>String(x??'').trim()),nameCol=headers.findIndex(h=>nh(h)==='student name');if(nameCol<0)return null;const examCols=headers.map((h,i)=>({name:h,index:i})).filter(x=>x.index!==nameCol&&x.name);const records=[];for(let i=1;i<matrix.length;i++){const r=matrix[i]||[],student=String(r[nameCol]??'').trim();if(!student)continue;const marks={};for(const c of examCols)marks[c.name]=r[c.index]??'';records.push({name:student,marks,row:i+1})}if(!examCols.length)throw Error('Worksheet "'+name+'" has no mark columns.');return {name,headers,examCols:examCols.map(x=>x.name),records}}
  async function parse(ab){const z=await zip(ab),shared=[];if(z.has('xl/sharedStrings.xml')){const d=xml(z.get('xl/sharedStrings.xml'));for(const si of all(d,'si'))shared.push(all(si,'t').map(x=>x.textContent).join(''))}
    const wb=xml(z.get('xl/workbook.xml')),sheets=all(wb,'sheet');if(!sheets.length)throw Error('No worksheets found.');const rels=z.has('xl/_rels/workbook.xml.rels')?xml(z.get('xl/_rels/workbook.xml.rels')):null;const result=[];
    for(const sh of sheets){const id=sh.getAttribute('r:id')||sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');const rel=rels&&all(rels,'Relationship').find(x=>x.getAttribute('Id')===id);const path=rel?resolvePath(rel.getAttribute('Target')):'xl/worksheets/sheet'+(result.length+1)+'.xml';const parsed=await parseSheet(z,path,sh.getAttribute('name')||('Sheet '+(result.length+1)),shared);if(parsed)result.push(parsed)}
    if(!result.length)throw Error('No worksheets contain a "Student Name" column.');return {sheets:result};
  }
  window.MarksXLSX={parse};
})();
