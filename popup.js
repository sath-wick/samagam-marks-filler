const $=id=>document.getElementById(id);let workbook=null,pageRows=null,verification=null;
function normalizeName(s){return String(s??'').toUpperCase().normalize('NFKD').replace(/[\u0300-\u036f]/g,'').replace(/&/g,' AND ').replace(/[^A-Z0-9]+/g,'').trim()}
function esc(s){return String(s).replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}
function validMark(v){if(v===''||v==null)return{ok:true,blank:true};const s=String(v).trim();if(/^(AB|ABSENT|NA|N\/A)$/i.test(s))return{ok:true,value:s};const n=Number(s);return Number.isFinite(n)&&n>=0&&n<=40?{ok:true,value:s}:{ok:false,reason:'Invalid mark: '+s}}
async function activeTab(){const t=await chrome.tabs.query({active:true,currentWindow:true});if(!t[0]?.id)throw Error('Could not access the active tab.');return t[0].id}
async function scanPage(){const id=await activeTab(),[{result}]=await chrome.scripting.executeScript({target:{tabId:id},func:()=>{const table=document.querySelector('#stageMarksGrid');if(!table)return{error:'Could not find #stageMarksGrid on this page.'};return{rows:[...table.querySelectorAll('tbody tr')].map((tr,i)=>{const c=tr.querySelectorAll('td'),input=tr.querySelector('input.stage-mark-cell');return{index:i+1,name:(c[3]?.textContent||'').trim().replace(/\s+/g,' '),admission:(c[2]?.textContent||'').trim(),hasInput:!!input}})}}});if(result?.error)throw Error(result.error);return result.rows}
function currentRecords(){return workbook.sheets.find(s=>s.name===$('sheetSelect').value)}
function populateSheets(){ $('sheetSelect').innerHTML=workbook.sheets.map(s=>`<option value="${esc(s.name)}">${esc(s.name)}</option>`).join('');$('sheetSelect').disabled=false;updateColumns() }
function updateColumns(){const s=currentRecords();$('examSelect').innerHTML=s.examCols.map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join('');$('examSelect').disabled=false;$('sheetInfo').textContent=`${workbook.sheets.length} sheet${workbook.sheets.length===1?'':'s'} · ${s.records.length} students in selected sheet · ${s.examCols.length} mark columns`;$('scanBtn').disabled=false;$('fillBtn').disabled=true;pageRows=null;verification=null;$('verification').classList.add('hidden');$('scanStatus').textContent='Not scanned';$('scanStatus').className='status neutral';$('summary').textContent='Select an exam, then scan the website.'}
function buildVerification(){const records=currentRecords(),map=new Map(records.records.map(r=>[normalizeName(r.name),r])),matched=[],missing=[],duplicates=[],invalid=[],seen=new Set();for(const p of pageRows){const k=normalizeName(p.name);if(seen.has(k))duplicates.push(p.name);seen.add(k);const e=map.get(k);if(!e){missing.push({name:p.name});continue}const v=validMark(e.marks[$('examSelect').value]);if(!v.ok)invalid.push({name:e.name,reason:v.reason});matched.push({page:p,excel:e,mark:v.blank?'':v.value})}for(const e of records.records)if(!seen.has(normalizeName(e.name)))missing.push({name:e.name});return{matched,missing,duplicates,invalid,total:records.records.length}}
function renderVerification(){verification=buildVerification();const ok=verification.missing.length===0&&verification.duplicates.length===0&&verification.invalid.length===0&&verification.matched.length===pageRows.length;const st=$('scanStatus');st.textContent=ok?'Verified':'Needs review';st.className='status '+(ok?'good':'warn');$('verification').classList.remove('hidden');$('verification').innerHTML=`<div class="verify-grid"><div class="verify-item"><b>${verification.matched.length}/${pageRows.length}</b><span>website names matched</span></div><div class="verify-item"><b>${verification.total}</b><span>sheet students</span></div></div>${verification.missing.length?`<div class="issues"><b>Unmatched:</b> ${verification.missing.map(x=>esc(x.name)).join(', ')}</div>`:''}${verification.duplicates.length?`<div class="issues"><b>Duplicate website names:</b> ${verification.duplicates.map(esc).join(', ')}</div>`:''}${verification.invalid.length?`<div class="issues"><b>Invalid marks:</b> ${verification.invalid.map(x=>esc(x.name+' ('+x.reason+')')).join(', ')}</div>`:''}${ok?'<div class="result success">All names match exactly after normalization. Fill is unlocked.</div>':'<div class="result warning">Filling is blocked until every student has an unambiguous match and valid marks.</div>'}`;$('fillBtn').disabled=!ok;$('summary').textContent=`${currentRecords().name} · ${$('examSelect').value} · ${verification.matched.length}/${pageRows.length} verified`}
async function loadBuffer(buf,label){workbook=await MarksXLSX.parse(buf);$('fileLabel').textContent=label;populateSheets()}
$('localTab').addEventListener('click',()=>{ $('localSource').classList.remove('hidden'); $('googleSource').classList.add('hidden'); $('localTab').classList.add('active'); $('googleTab').classList.remove('active'); });
$('googleTab').addEventListener('click',()=>{ $('googleSource').classList.remove('hidden'); $('localSource').classList.add('hidden'); $('googleTab').classList.add('active'); $('localTab').classList.remove('active'); });

$('fileInput').addEventListener('change',async e=>{const f=e.target.files[0];if(!f)return;try{await loadBuffer(await f.arrayBuffer(),f.name);$('result').textContent=''}catch(err){workbook=null;$('sheetInfo').textContent='Could not read workbook.';$('result').innerHTML=`<span class="error">${esc(err.message)}</span>`}});
$('sheetSelect').addEventListener('change',()=>{updateColumns();saveState()});$('examSelect').addEventListener('change',()=>{if(pageRows)renderVerification();saveState()});
$('googleBtn').addEventListener('click',async()=>{const url=$('googleUrl').value.trim();if(!url){$('googleResult').innerHTML='<span class="error">Paste a Google Sheets URL first.</span>';return}try{const m=url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);if(!m)throw Error('That does not look like a Google Sheets URL.');const exportUrl=`https://docs.google.com/spreadsheets/d/${m[1]}/export?format=xlsx`;const response=await new Promise((resolve,reject)=>{chrome.runtime.sendMessage({type:'fetchGoogleSheet',id:m[1]},resolve)});if(chrome.runtime.lastError)throw Error(chrome.runtime.lastError.message);if(!response?.ok)throw Error(response?.error||'Could not fetch the Google Sheet.');const binary=atob(response.base64||'');const bytes=new Uint8Array(binary.length);for(let i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);await loadBuffer(bytes.buffer,'Google Sheets workbook');$('googleResult').innerHTML='<span class="success">Google Sheet loaded.</span>'}catch(err){$('googleResult').innerHTML=`<span class="error">${esc(err.message)}</span>`}});
$('scanBtn').addEventListener('click',async()=>{const b=$('scanBtn');b.disabled=true;b.textContent='Scanning…';try{pageRows=await scanPage();renderVerification();await saveState()}catch(err){$('result').innerHTML=`<span class="error">${esc(err.message)}</span>`;$('scanStatus').textContent='Scan failed';$('scanStatus').className='status bad';await saveState()}finally{b.disabled=false;b.textContent='Scan & verify names'}});
$('fillBtn').addEventListener('click',async()=>{if(!verification)return;const marks={};for(const x of verification.matched)marks[normalizeName(x.excel.name)]=x.mark;const id=await activeTab(),[{result}]=await chrome.scripting.executeScript({target:{tabId:id},args:[marks],func:marks=>{const table=document.querySelector('#stageMarksGrid');if(!table)return{error:'Marks table disappeared.'};const normalize=s=>String(s??'').toUpperCase().normalize('NFKD').replace(/[\u0300-\u036f]/g,'').replace(/&/g,' AND ').replace(/[^A-Z0-9]+/g,'').trim();let filled=0,blank=0;for(const tr of table.querySelectorAll('tbody tr')){const n=(tr.querySelector('td:nth-child(4)')?.textContent||'').trim(),k=normalize(n),input=tr.querySelector('input.stage-mark-cell');if(!(k in marks)||!input)continue;const v=marks[k];if(v===''){blank++;continue}const setter=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value')?.set;setter?setter.call(input,String(v)):input.value=String(v);input.dispatchEvent(new Event('input',{bubbles:true}));input.dispatchEvent(new Event('change',{bubbles:true}));input.dispatchEvent(new Event('blur',{bubbles:true}));filled++}return{filled,blank}}});if(result?.error){$('result').innerHTML=`<span class="error">${esc(result.error)}</span>`;return}$('result').innerHTML=`<span class="success">Filled <b>${result.filled}</b> marks for ${esc($('examSelect').value)}.</span>${result.blank?` ${result.blank} blank cells were left untouched.`:''}<br><span class="hint">Review the page, then use the website's own Save/Submit button.</span>`;await saveState()});

// Keep popup state in memory while the browser session is open. Chrome closes
// action popups whenever focus leaves them, so reopening should not wipe the work.
const STATE_KEY='marksFillerState';
let restoringState=false;

async function saveState(){
  if(restoringState) return;
  try{
    const state={
      source:$('googleTab').classList.contains('active')?'google':'local',
      googleUrl:$('googleUrl').value,
      fileLabel:$('fileLabel').textContent,
      workbook,
      sheet:workbook?$('sheetSelect').value:'',
      exam:workbook?$('examSelect').value:'',
      pageRows,
      verification,
      tabId:pageRows?await activeTab().catch(()=>null):null
    };
    await chrome.storage.session.set({[STATE_KEY]:state});
  }catch(_){/* State persistence is convenience only. */}
}

async function restoreState(){
  restoringState=true;
  try{
    const data=await chrome.storage.session.get(STATE_KEY);
    const s=data?.[STATE_KEY];
    if(!s) return;

    if(s.source==='google'){
      $('googleTab').click();
    }else{
      $('localTab').click();
    }

    if(s.googleUrl) $('googleUrl').value=s.googleUrl;
    if(s.fileLabel && s.fileLabel!=='Choose Excel file') $('fileLabel').textContent=s.fileLabel;

    if(s.workbook?.sheets?.length){
      workbook=s.workbook;
      populateSheets();
      if(s.sheet && workbook.sheets.some(x=>x.name===s.sheet)){
        $('sheetSelect').value=s.sheet;
        updateColumns();
      }
      if(s.exam && [...$('examSelect').options].some(x=>x.value===s.exam)){
        $('examSelect').value=s.exam;
      }
    }

    let sameTab=false;
    if(s.pageRows && s.tabId!=null){
      try{sameTab=(await activeTab())===s.tabId}catch(_){sameTab=false}
    }
    if(sameTab){
      pageRows=s.pageRows;
      verification=s.verification;
      if(verification){
        renderVerification();
      }
    }
  }finally{
    restoringState=false;
  }
}

// Persist only UI/workflow state, never credentials or Google authentication.
['googleUrl','sheetSelect','examSelect'].forEach(id=>$(id).addEventListener('input',saveState));
$('googleUrl').addEventListener('change',saveState);
$('fileInput').addEventListener('change',saveState);
$('localTab').addEventListener('click',saveState);
$('googleTab').addEventListener('click',saveState);

// Wrap the existing workflow handlers with persistence at the important state changes.
const originalLoadBuffer=loadBuffer;
loadBuffer=async function(buf,label){await originalLoadBuffer(buf,label);await saveState()};

restoreState();
