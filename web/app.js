'use strict';
const $=id=>document.getElementById(id);
const fields=[['font','Font','Times New Roman'],['size','Body size (pt)',13],['line','Line spacing',1.5],['h1','Heading 1 (pt)',16],['h2','Heading 2 (pt)',14],['h3','Heading 3–9 (pt)',13],['left','Left margin (cm)',3],['right','Right margin (cm)',2],['top','Top margin (cm)',2],['bottom','Bottom margin (cm)',2]];
fields.push(['before','Space before (pt)',0],['after','Space after (pt)',6],['first_indent','First line indent (cm)',0],['table_size','Table font (pt)',12],['caption_size','Caption font (pt)',12],['page_start','First page number',1]);
const switches=[['numbering','Native multilevel numbering',true],['toc','Insert table of contents',true],['tables','Format table text',true],['figures','Center figures and captions',true],['chapter_break','Page break before chapters',false]];
switches.push(['caption_numbers','Number existing captions',true],['figure_list','List of figures',true],['table_list','List of tables',true],['page_numbers','Page numbers',true],['update_fields','Request field refresh',true],['heading_uppercase','Uppercase chapter headings',false],['heading_bold','Bold headings',true]);
switches.push(['format_body','Format body text',true],['format_headings','Format headings',true],['format_layout','Format page layout',true],['protect_cover','Protect cover section',false],['cleanup_text','Clean repeated spaces',false],['format_citations','Align citation paragraphs safely',false],['table_fit','Fit oversized tables',false],['image_resize','Resize oversized inline images',false],['smart_pages','Smart page numbering (3 sections)',false]);
const enums=[['paper','Paper size',['A4','A5','Letter']],['alignment','Body alignment',['justify','left','center']],['numbering_style','Heading numbering',['decimal','chapter','roman']],['language','Caption language',['en','vi']],['page_position','Page number position',['center','right','left']],['page_format','Page number style',['decimal','lowerRoman','upperRoman']],['reference_style','Reference style',['none','APA','Harvard','IEEE']]];
const typographyGroups=[['body','Body text'],['heading','Headings'],['caption','Captions'],['table','Table text']];
const typographyDefaults={};
for(const [group,label] of typographyGroups){
 fields.push([group+'_color',label+' color','#000000']);
 typographyDefaults[group+'_color']='#000000';
 for(const [attr,title] of [['bold','Bold'],['italic','Italic'],['underline','Underline']]){
  const key=group+'_'+attr+'_mode',initial=group==='heading'?(attr==='bold'?'on':'off'):'preserve';
  enums.push([key,title,['preserve','on','off']]);typographyDefaults[key]=initial;
 }
}
let capabilities={history:false,pdf:false,ai:false,multipart:false,inspection:false};
let uploadId=null, lastReport=null, batchEntries=[], compareMode=false;
let originalFull=null, formattedFull=null, previewUrls=[], previewEpoch=0;
let wizardStep=0, processed=null;
let selectedFile=null, previewBlocks=[], undoStack=[], redoStack=[], proposed=true, resultCache=null;
let documentData='', name='', rows=[], all=false, busy=false, generation=0;
function node(tag,text,cls){const e=document.createElement(tag);if(text!==undefined)e.textContent=text;if(cls)e.className=cls;return e;}
const grid=node('div',undefined,'fieldgrid');$('settings').append(grid);const extra=node('div',undefined,'fieldgrid');$('extraSettings').append(extra);
for(const [key,label,value] of fields){if(key.endsWith('_color'))continue;const l=node('label',label,key==='font'?'wide':'');const i=node('input');i.id=key;i.value=value;if(typeof value==='number'){i.type='number';i.step='0.1';}l.append(i);(['font','size','line','h1','h2','h3','left','right','top','bottom'].includes(key)?grid:extra).append(l);}
for(const [key,label,value] of switches){const l=node('label');const i=node('input');i.type='checkbox';i.id=key;i.checked=value;l.append(i,document.createTextNode(label));$('switches').append(l);}
for(const [key,label,values] of enums){if(key.endsWith('_mode'))continue;const l=node('label',label);const i=node('select');i.id=key;for(const v of values){const o=node('option',v);o.value=v;i.append(o);}l.append(i);extra.append(l);}
for(const [group,label] of typographyGroups){
 const card=node('div',undefined,'type-card');card.append(node('h3',label));
 const example=node('div','Aa — Sample text','type-example');card.append(example);
 const controls=node('div',undefined,'type-controls');
 for(const [attr,title] of [['bold','Bold'],['italic','Italic'],['underline','Underline']]){
  const key=group+'_'+attr+'_mode',l=node('label',title),select=node('select');select.id=key;
  for(const [v,t] of [['preserve','Keep'],['on','On'],['off','Off']]){const o=node('option',t);o.value=v;select.append(o);}
  select.value=typographyDefaults[key];l.append(select);controls.append(l);
 }
 const colorLabel=node('label','Text color'),input=node('input');input.id=group+'_color';input.value='#000000';input.placeholder='#000000 or preserve';input.maxLength=8;
 const picker=node('input');picker.type='color';picker.value='#000000';picker.setAttribute('aria-label',label+' color picker');
 const keep=node('button','Keep original','secondary');keep.type='button';keep.onclick=()=>{input.value='preserve';input.dispatchEvent(new Event('change',{bubbles:true}));};
 picker.oninput=()=>{input.value=picker.value;input.dispatchEvent(new Event('change',{bubbles:true}));};
 const colorRow=node('div',undefined,'type-color');colorRow.append(picker,input,keep);colorLabel.append(colorRow);card.append(controls,colorLabel);
 const update=()=>{const color=input.value.trim();if(/^#[0-9a-f]{6}$/i.test(color)){picker.value=color;example.style.color=color;}else example.style.color='';example.style.fontWeight=$(group+'_bold_mode').value==='on'?'700':'400';example.style.fontStyle=$(group+'_italic_mode').value==='on'?'italic':'normal';example.style.textDecoration=$(group+'_underline_mode').value==='on'?'underline':'none';};
 card.addEventListener('change',update);card.addEventListener('input',update);$('typography').append(card);update();
}
function config(){return Object.fromEntries([...fields.map(([k,,v])=>[k,typeof v==='number'?Number($(k).value):$(k).value]),...switches.map(([k])=>[k,$(k).checked]),...enums.map(([k])=>[k,$(k).value])]);}
function status(msg,error=false){$('status').textContent=msg;$('status').hidden=false;$('status').className='status'+(error?' error':'');}
function lock(v){busy=v;$('refreshPreview').disabled=v||!selectedFile;document.body.classList.toggle('busy',v);$('batchRun').disabled=v;$('batchFiles').disabled=v;$('doiVerify').disabled=v||!selectedFile||!capabilities.doi;$('file').disabled=v;$('analyze').disabled=v;$('export').disabled=v||!processed||processed.key!==previewKey();$('processDocument').disabled=v||!selectedFile;document.querySelectorAll('[data-step],[data-go]').forEach(b=>b.disabled=v);$('showAll').disabled=v||!selectedFile;for(const id of ['audit','preview','downloadPdf'])$(id).disabled=v||!documentData;for(const id of ['parseRules','aiRules','reqFile','sampleFile','sampleRules'])$(id).disabled=v; $('aiRules').disabled=v||!capabilities.ai; for(const id of ['preview','downloadPdf'])$(id).disabled=v||!documentData||!capabilities.pdf;$('references').disabled=v||!selectedFile||!capabilities.inspection;$('sampleRules').disabled=v||!capabilities.inspection;}
function render(){const root=$('rows');root.replaceChildren();const query=$('headingSearch').value.toLowerCase().trim();const filtered=rows.filter(r=>(all||r.level>0)&&(!query||r.text.toLowerCase().includes(query)));if(!filtered.length)root.append(node('p','No heading candidates. Show all paragraphs to assign headings manually.','empty'));for(const r of filtered){const item=node('div',undefined,'item');const left=node('div');left.append(node('p',r.text),node('small','#'+(r.index+1)+' · '+(r.protected?'Protected':r.reason||'Body text')+(r.confidence?' · '+r.confidence:'')));const s=node('select');s.setAttribute('aria-label','Heading level for paragraph '+(r.index+1));for(let n=0;n<=9;n++){const o=node('option',n?'Heading '+n:'Body text');o.value=n;s.append(o);}s.value=r.level;s.disabled=r.protected;s.addEventListener('change',()=>{undoStack.push({index:r.index,from:r.level,to:Number(s.value)});redoStack=[];r.level=Number(s.value);resultCache=null;updateStats();updateUndo();renderPreview();});item.append(left,s);root.append(item);}}
function updateStats(){const box=$('stats');box.replaceChildren();for(const [label,num] of [['Paragraphs',rows.length],['Headings',rows.filter(r=>r.level).length],['Protected',rows.filter(r=>r.protected).length]])box.append(node('div',num+' '+label,'stat'));}
function encode(file){return new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result.split(',')[1]);reader.onerror=reject;reader.readAsDataURL(file);});}
async function request(path,payload){let headers={'X-ReportStudio':'1'},body;if(payload.document instanceof Blob){if(capabilities.multipart){body=new FormData();body.append('document',payload.document);const {document:omit,...rest}=payload;body.append('payload',JSON.stringify(rest));}else{payload={...payload,document:await encode(payload.document)};}}if(!body){headers['Content-Type']='application/json';body=JSON.stringify(payload);}const res=await fetch(path,{method:'POST',headers,body});let data;try{data=await res.json();}catch{throw Error('Server unavailable. Please retry.');}if(!res.ok){const e=Error(data.error||'Request failed');e.code=data.code;throw e;}return data;}
$('file').addEventListener('change',()=>{generation++;processed=null;resetFullPreview();showStep(0);lastReport=null;$('changeReport').disabled=true;if(uploadId)fetch('/api/uploads/'+encodeURIComponent(uploadId),{method:'DELETE',headers:{'X-ReportStudio':'1'}}).catch(()=>{});uploadId=null;selectedFile=null;documentData='';rows=[];previewBlocks=[];undoStack=[];redoStack=[];resultCache=null;all=false;updateUndo();renderPreview();$('filename').textContent=$('file').files[0]?.name||'No document selected';$('rows').replaceChildren(node('p','Analyze the selected document to continue.','empty'));$('stats').replaceChildren();$('result').textContent='';lock(false);});
$('analyze').addEventListener('click',async()=>{if(busy)return;const file=$('file').files[0];if(!file)return status('Select a DOCX file first.',true);if(file.size>15*1024*1024)return status('Maximum file size is 15 MB.',true);lock(true);resetFullPreview();selectedFile=null;documentData='';rows=[];previewBlocks=[];resultCache=null;$('result').textContent='';status('Analyzing document…');try{const result=await request('/api/analyze',{document:file,outline:$('outline').value,cache_upload:!!capabilities.upload_sessions});uploadId=result.upload_id||null;selectedFile=file;documentData='ready';name=file.name;rows=result.rows;previewBlocks=result.preview?.blocks||rows; $('previewNote').textContent=result.preview?.note||'Approximate text only.';undoStack=[];redoStack=[];resultCache=null;updateUndo();renderPreview();all=false;$('showAll').textContent='Show all paragraphs';updateStats();render();status('Analysis complete. Choose formatting settings.');showStep(1);}catch(e){status(e.message,true);}finally{lock(false);}});
$('showAll').addEventListener('click',()=>{all=!all;$('showAll').textContent=all?'Show headings only':'Show all paragraphs';render();});
$('export').addEventListener('click',async()=>{
 if(busy||!processed||processed.key!==previewKey()){status('Apply formatting before downloading.',true);return;}
 lock(true);
 try{
  let result=processed.result;
  if(capabilities.history&&$('saveHistory').checked&&!result.saved){result=await formatted(true);if(result.report.input_sha256!==processed.result.report.input_sha256||result.report.output_sha256!==processed.result.report.output_sha256){status('History copy was regenerated; downloading the verified preview copy instead.');result=processed.result;}else processed={key:previewKey(),result};}
  if(!result.report.integrity?.passed||result.report.change_report?.unchanged)throw Error('No formatting changes detected. Adjust settings and process again.');
  downloadBlob(new Blob([Uint8Array.from(atob(result.document),c=>c.charCodeAt(0))],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}),name.replace(/\.docx$/i,'')+'_formatted.docx');
  status(result.storage_warning||'Formatted document downloaded.',!!result.storage_warning);if(result.saved)refreshHistory();
 }catch(e){status(e.message,true);}finally{lock(false);}
});



let suggested={}, pdfUrl='', sampleBaseline=null, samplePrevious=null, sampleName='';
function sampleState(){
 const box=$('sampleState');if(!box)return;
 box.hidden=!sampleBaseline;
 if(!sampleBaseline)return;
 const current=config(),modified=Object.keys(sampleBaseline).filter(k=>current[k]!==sampleBaseline[k]);
 $('sampleSource').textContent=sampleName;
 $('sampleCounts').textContent=Object.keys(sampleBaseline).length+' imported settings · '+modified.length+' manually changed';
 $('restoreSample').disabled=busy;$('undoSample').disabled=busy||!samplePrevious;
}
function validSettings(obj){
 const clean={};if(!obj||typeof obj!=='object')return clean;
 for(const [k,,v] of fields){if(typeof obj[k]!==typeof v)continue;if(typeof v==='number'&&!Number.isFinite(obj[k]))continue;if(k.endsWith('_color')&&!/^(#[0-9a-f]{6}|preserve)$/i.test(obj[k]))continue;clean[k]=obj[k];}
 for(const [k] of switches)if(typeof obj[k]==='boolean')clean[k]=obj[k];
 for(const [k,,values] of enums)if(values.includes(obj[k]))clean[k]=obj[k];
 return clean;
}
function invalidateSettings(){resultCache=null;sampleState();if($('textPreviewView')&&!$('textPreviewView').hidden)renderPreview();}
$('restoreSample').onclick=()=>{if(!busy&&sampleBaseline){setConfig(sampleBaseline);status('Sample formatting restored.');}};
$('undoSample').onclick=()=>{if(!busy&&samplePrevious){const previous=samplePrevious;sampleBaseline=null;samplePrevious=null;sampleName='';setConfig(previous);sampleState();status('Settings before sample restored.');}};
$('formatPanel').addEventListener('input',e=>{if(fields.some(([k])=>k===e.target.id)||enums.some(([k])=>k===e.target.id)||switches.some(([k])=>k===e.target.id)){resultCache=null;sampleState();}});
$('formatPanel').addEventListener('change',sampleState);

function renderAudit(a){const root=$('auditResult');root.replaceChildren(node('div',a.score+'/100','audit-score'),node('p',a.failures+' failed checks out of '+a.checks+' · '+a.protected_skipped+' protected paragraphs skipped','muted'),node('p',a.note,'muted'));for(const c of a.categories)root.append(node('div',c.name+': '+c.failed+' / '+c.checks+' issues','issue'));for(const i of a.issues)root.append(node('div',i.location+' — '+i.detail,'issue'));}
async function parseRequirements(ai){if(busy)return;if(ai&&!$('aiConsent').checked)return status('Confirm AI consent first.',true);lock(true);try{const file=$('reqFile').files[0];const text=$('reqText').value.trim();const payload={ai,consent:$('aiConsent').checked,text};if(file){if(file.size>15*1024*1024)throw Error('Maximum file size is 15 MB.');payload.document=file;payload.filename=file.name;}const result=await request('/api/requirements',payload);suggested=result.rules;$('reqText').value=result.text;const root=$('ruleReview');root.replaceChildren();for(const [k,v] of Object.entries(suggested)){if(!$(k))continue;const l=node('label');const i=node('input');i.type='checkbox';i.checked=true;i.dataset.rule=k;l.append(i,document.createTextNode(k+': '+v),node('small',result.evidence[k]||'Review this suggestion'));root.append(l);}for(const w of result.warnings)root.append(node('p',w,'error'));$('applyRules').disabled=!Object.keys(suggested).length;status('Review extracted requirements before applying.');}catch(e){status(e.message,true);}finally{lock(false);}}
$('parseRules').onclick=()=>parseRequirements(false);$('aiRules').onclick=()=>parseRequirements(true);
$('applyRules').onclick=()=>{if(busy)return;const chosen={};for(const i of $('ruleReview').querySelectorAll('input:checked'))chosen[i.dataset.rule]=suggested[i.dataset.rule];setConfig(validSettings(chosen));sampleState();status('Selected rules applied to manual settings.');};
$('audit').onclick=async()=>{if(busy||!documentData)return;lock(true);try{renderAudit(await documentRequest('/api/audit',{config:config(),outline:$('outline').value}));status('Original document checked.');}catch(e){status(e.message,true);}finally{lock(false);}};
async function makePdf(download){if(busy||!documentData)return;lock(true);status('Rendering formatted document…');try{const r=await formatted(false);renderAudit(r.report.audit);const p=await request('/api/pdf',{document:new Blob([Uint8Array.from(atob(r.document),c=>c.charCodeAt(0))])});const bytes=Uint8Array.from(atob(p.pdf),c=>c.charCodeAt(0));if(pdfUrl)URL.revokeObjectURL(pdfUrl);pdfUrl=URL.createObjectURL(new Blob([bytes],{type:'application/pdf'}));if(download){const a=node('a');a.href=pdfUrl;a.download=name.replace(/\.docx$/i,'')+'_formatted.pdf';a.click();}else{$('pdfFrame').src=pdfUrl;$('pdfFrame').hidden=false;}status('PDF rendered with LibreOffice. Layout may differ from Word.');}catch(e){status(e.message,true);}finally{lock(false);}}
$('preview').onclick=()=>makePdf(false);$('downloadPdf').onclick=()=>makePdf(true);
$('file').addEventListener('change',()=>{$('auditResult').replaceChildren();$('referenceResult').replaceChildren();$('pdfFrame').hidden=true;$('pdfFrame').removeAttribute('src');if(pdfUrl){URL.revokeObjectURL(pdfUrl);pdfUrl='';}});

async function refreshHistory(){if(!capabilities.history)return;const root=$('historyRows');try{const res=await fetch('/api/history');if(!res.ok)throw Error('Unable to load history');const data=await res.json();root.replaceChildren();if(!data.files.length)root.append(node('p','No saved files.','muted'));for(const f of data.files){const item=node('div',undefined,'item');const info=node('div');info.append(node('p',f.filename),node('small',Math.ceil(f.size_bytes/1024)+' KB · Expires '+new Date(f.expires_at).toLocaleString()));const countdown=node('small');countdown.dataset.expires=f.expires_at;info.append(countdown);const actions=node('div');const a=node('a','Download');a.href='/api/history/'+encodeURIComponent(f.id);const b=node('button','Delete','secondary');b.onclick=async()=>{if(!confirm('Delete this saved copy?'))return;try{const r=await fetch('/api/history/'+encodeURIComponent(f.id),{method:'DELETE',headers:{'X-ReportStudio':'1'}});if(!r.ok)throw Error('Unable to delete');await refreshHistory();}catch(e){status(e.message,true);}};actions.append(a,b);item.append(info,actions);root.append(item);}}catch(e){root.replaceChildren(node('p',e.message,'error'));}}
$('refreshHistory').onclick=refreshHistory;
(async()=>{try{const r=await fetch('/api/capabilities');if(r.ok)capabilities=await r.json();}catch{}$('doiVerify').hidden=!capabilities.doi;$('historyPanel').hidden=wizardStep!==5||!capabilities.history;$('historyNav').hidden=!capabilities.history;$('pdfActions').hidden=!capabilities.pdf;$('aiRules').hidden=!capabilities.ai;$('aiConsentLabel').hidden=!capabilities.ai;$('retentionLabel').hidden=!capabilities.history;$('aiConsent').disabled=!capabilities.ai;lock(false);refreshHistory();})();

function setConfig(c){c=validSettings(c);for(const [k] of [...fields,...enums])if(k in c)$(k).value=c[k];for(const [k] of switches)if(k in c)$(k).checked=c[k];for(const [group] of typographyGroups)$(group+'_color').dispatchEvent(new Event('change',{bubbles:true}));invalidateSettings();}
const builtin={academic:{font:'Times New Roman',size:13,h1:16,h2:14,h3:13,line:1.5,left:3,right:2,top:2,bottom:2,paper:'A4',alignment:'justify'},business:{font:'Arial',size:11,h1:16,h2:13,h3:12,line:1.15,left:2.54,right:2.54,top:2.54,bottom:2.54,paper:'A4',alignment:'left'},proposal:{font:'Times New Roman',size:12,h1:16,h2:14,h3:12,line:1.5,left:2.54,right:2.54,top:2.54,bottom:2.54,paper:'A4',alignment:'justify'}};
$('presetSelect').onchange=()=>{if(builtin[$('presetSelect').value])setConfig(builtin[$('presetSelect').value]);};
function presets(){try{return JSON.parse(localStorage.getItem('reportstudio-presets')||'{}');}catch{return {};}}
function fillPresets(){const sel=$('savedPresets');sel.replaceChildren(node('option','Choose a preset'));sel.firstChild.value='';for(const key of Object.keys(presets())){const o=node('option',key);o.value=key;sel.append(o);}}
$('savePreset').onclick=()=>{const key=$('presetName').value.trim().slice(0,60);if(!key)return status('Enter a preset name.',true);try{const p=presets();Object.defineProperty(p,key,{value:config(),enumerable:true,configurable:true,writable:true});localStorage.setItem('reportstudio-presets',JSON.stringify(p));fillPresets();$('savedPresets').value=key;status('Complete formatting preset saved, including imported and manual settings.');}catch(e){status(e.message,true);}};
$('loadPreset').onclick=()=>{const p=presets(),key=$('savedPresets').value;if(key&&Object.hasOwn(p,key)){sampleBaseline=null;samplePrevious=null;sampleName='';setConfig(p[key]);status('Complete formatting preset loaded.');}};
$('deletePreset').onclick=()=>{const p=presets(),key=$('savedPresets').value;if(key&&Object.hasOwn(p,key)){delete p[key];try{localStorage.setItem('reportstudio-presets',JSON.stringify(p));fillPresets();}catch(e){status(e.message,true);}}};fillPresets();
function setMode(quick){document.body.classList.remove('quick');$('quickMode').setAttribute('aria-pressed',quick);$('advancedMode').setAttribute('aria-pressed',!quick);$('quickMode').classList.toggle('secondary',!quick);$('advancedMode').classList.toggle('secondary',quick);}
$('quickMode').onclick=()=>setMode(true);$('advancedMode').onclick=()=>setMode(false);setMode(true);

function theme(dark){document.body.classList.toggle('dark',dark);$('theme').textContent=dark?'Light mode':'Dark mode';try{localStorage.setItem('reportstudio-theme',dark?'dark':'light');}catch{}}
$('theme').onclick=()=>theme(!document.body.classList.contains('dark'));try{theme(localStorage.getItem('reportstudio-theme')==='dark');}catch{}
const drop=$('dropzone');for(const evt of ['dragenter','dragover'])drop.addEventListener(evt,e=>{e.preventDefault();if(!busy)drop.classList.add('dragging');});for(const evt of ['dragleave','drop'])drop.addEventListener(evt,e=>{e.preventDefault();drop.classList.remove('dragging');});drop.addEventListener('drop',e=>{if(busy)return;const f=e.dataTransfer.files[0];if(!f||!/\.docx$/i.test(f.name))return status('Select a DOCX file.',true);const dt=new DataTransfer();dt.items.add(f);$('file').files=dt.files;$('file').dispatchEvent(new Event('change'));});
$('headingSearch').oninput=render;
function updateUndo(){$('undoHeading').disabled=!undoStack.length;$('redoHeading').disabled=!redoStack.length;}
function moveHeading(from,to,reverse){const change=from.pop();if(!change)return;const r=rows.find(x=>x.index===change.index);if(!r)return;r.level=reverse?change.from:change.to;to.push(change);resultCache=null;updateUndo();updateStats();render();renderPreview();}
$('undoHeading').onclick=()=>moveHeading(undoStack,redoStack,true);$('redoHeading').onclick=()=>moveHeading(redoStack,undoStack,false);
function tab(preview){$('outlineView').hidden=preview;$('textPreviewView').hidden=!preview;$('outlineTab').setAttribute('aria-pressed',!preview);$('previewTab').setAttribute('aria-pressed',preview);if(preview){renderPreview();loadFullPreview();}}
$('outlineTab').onclick=()=>tab(false);$('previewTab').onclick=()=>tab(true);
$('beforePreview').onclick=()=>{compareMode=false;proposed=false;renderPreview();loadFullPreview();};$('afterPreview').onclick=()=>{compareMode=false;proposed=true;renderPreview();loadFullPreview();};
function resetFullPreview(){previewEpoch++;originalFull=null;formattedFull=null;for(const url of previewUrls)URL.revokeObjectURL(url);previewUrls=[];}
function previewKey(){return JSON.stringify({config:config(),outline:$('outline').value,decisions:rows.map(r=>[r.index,r.level])});}
function imageUrl(model,key){if(!model.urls)model.urls=new Map();if(model.urls.has(key))return model.urls.get(key);const a=model.assets?.[key];if(!a)return '';const bytes=Uint8Array.from(atob(a.base64),c=>c.charCodeAt(0));const url=URL.createObjectURL(new Blob([bytes],{type:a.mime}));model.urls.set(key,url);previewUrls.push(url);return url;}
function releaseModel(model){if(!model?.urls)return;for(const url of model.urls.values()){URL.revokeObjectURL(url);previewUrls=previewUrls.filter(u=>u!==url);}}
function renderPreview(){
 const root=$('textPreview');const scrolls=Array.from(root.querySelectorAll('.preview-scroll')).map(e=>[e.scrollTop,e.scrollLeft]);root.replaceChildren();root.classList.toggle('compare',compareMode);
 $('beforePreview').setAttribute('aria-pressed',!proposed&&!compareMode);$('afterPreview').setAttribute('aria-pressed',proposed&&!compareMode);
 if(!selectedFile){root.append(node('p','Analyze a document first.','empty'));return;}
 const key=previewKey(),afterModel=formattedFull?.key===key?formattedFull.model:null,zoom=Number($('zoom').value)/100;
 const number=(v,a,b,f)=>Number.isFinite(v)?Math.max(a,Math.min(b,v)):f;
 const font=v=>typeof v==='string'&&/^[\p{L}\p{N} ,\-]{1,80}$/u.test(v)?v:null;
 function appendBlocks(target,blocks,model){
  for(const b of blocks){
   if(b.kind==='table'){const table=node('table',undefined,'preview-table'),body=node('tbody');for(const row of b.rows){const tr=node('tr');for(const cell of row){const td=node('td');td.colSpan=number(cell.span,1,100,1);appendBlocks(td,cell.blocks,model);tr.append(td);}body.append(tr);}table.append(body);target.append(table);continue;}
   const p=node('p',undefined,'preview-paragraph');if(b.page_break)p.classList.add('preview-pagebreak');if(b.level)p.classList.add('preview-heading');if(b.alignment)p.style.textAlign=b.alignment;if(b.line)p.style.lineHeight=number(b.line,.5,4,1.5);if(font(b.font))p.style.fontFamily=font(b.font);if(b.size)p.style.fontSize=number(b.size,6,72,13)+'pt';if(b.bold!==undefined)p.style.fontWeight=b.bold?'bold':'normal';if(b.italic!==undefined)p.style.fontStyle=b.italic?'italic':'normal';if(b.underline!==undefined)p.style.textDecoration=b.underline?'underline':'none';if(/^#[0-9a-f]{6}$/i.test(b.color||''))p.style.color=b.color;for(const [k,css] of [['left_indent','marginLeft'],['right_indent','marginRight'],['first_indent','textIndent'],['before','marginTop'],['after','marginBottom']])if(Number.isFinite(b[k]))p.style[css]=number(b[k],-150,300,0)+'pt';if(b.number)p.append(node('span',b.number+' '));
   for(const seg of b.segments||[]){if(seg.kind==='break'){p.append(node('span','Page break','preview-page-marker'));continue;}if(seg.kind==='image'){const url=seg.asset?imageUrl(model,seg.asset):'';if(url){const img=node('img');img.src=url;img.alt='Embedded document image';img.loading='lazy';img.decoding='async';if(seg.width)img.style.width=number(seg.width,1,2000,400)+'px';p.append(img);}else p.append(node('span',seg.missing||'Image unavailable in preview','preview-image-placeholder'));continue;}const span=node('span',seg.text||'');if(seg.bold!==undefined)span.style.fontWeight=seg.bold?'bold':'normal';if(seg.italic!==undefined)span.style.fontStyle=seg.italic?'italic':'normal';if(seg.underline!==undefined)span.style.textDecoration=seg.underline?'underline':'none';if(/^#[0-9a-f]{6}$/i.test(seg.color||''))span.style.color=seg.color;if(/^#[0-9a-f]{6}$/i.test(seg.highlight||''))span.style.backgroundColor=seg.highlight;if(font(seg.font))span.style.fontFamily=font(seg.font);if(seg.size)span.style.fontSize=number(seg.size,6,72,13)+'pt';p.append(span);}
   if(!p.childNodes.length)p.append(document.createTextNode('\u00a0'));target.append(p);
  }
 }
 for(const [i,after] of (compareMode?[false,true]:[proposed]).entries()){
  const model=after?afterModel:originalFull,pane=node('div',undefined,'preview-pane'),head=node('div',undefined,'preview-pane-title');head.append(node('strong',after?'Formatted document':'Original document'));pane.append(head);
  const scroll=node('div',undefined,'preview-scroll');if(!model){scroll.append(node('p',after?'Generate formatted preview to see the complete result.':'Load original preview to see the complete document.','empty'));pane.append(scroll);root.append(pane);continue;}
  const page=model.page||{},paper=node('div',undefined,'paper full-paper');paper.style.width=number(page.page_width,10,60,21)+'cm';paper.style.minHeight=number(page.page_height,10,60,29.7)+'cm';paper.style.padding=number(page.top_margin,0,8,2)+'cm '+number(page.right_margin,0,8,2)+'cm '+number(page.bottom_margin,0,8,2)+'cm '+number(page.left_margin,0,8,3)+'cm';paper.style.zoom=zoom;
  appendBlocks(paper,model.blocks,model);scroll.append(paper);pane.append(scroll);root.append(pane);if(scrolls[i]){scroll.scrollTop=scrolls[i][0];scroll.scrollLeft=scrolls[i][1];}
  if(model.warnings?.length)pane.append(node('p',model.warnings.join(' · '),'muted preview-warning'));
 }
 $('previewNote').textContent='Full body preview: images and tables included. Word pagination, headers, footers and complex floating objects may differ.';
}
async function loadFullPreview(force=false){
 if(busy||!selectedFile)return;const epoch=previewEpoch,key=previewKey();const needOriginal=!originalFull,needAfter=(proposed||compareMode)&&(force||formattedFull?.key!==key);
 if(!needOriginal&&!needAfter){renderPreview();return;}
 lock(true);status('Loading full document preview…');
 try{
  if(needOriginal){const model=await documentRequest('/api/preview',{});if(epoch!==previewEpoch)return;originalFull=model;}
  if(needAfter){const r=await formatted(false);const blob=new Blob([Uint8Array.from(atob(r.document),c=>c.charCodeAt(0))]);const model=await request('/api/preview',{document:blob});if(epoch!==previewEpoch)return;releaseModel(formattedFull?.model);formattedFull={key,model};}
  renderPreview();status('Full document preview ready.');
 }catch(e){status(e.message,true);renderPreview();}finally{lock(false);}
}
$('typography').addEventListener('change',()=>{resultCache=null;renderPreview();});
$('settings').addEventListener('change',()=>{resultCache=null;renderPreview();});$('extraSettings').addEventListener('change',()=>{resultCache=null;renderPreview();});$('switches').addEventListener('change',()=>{resultCache=null;renderPreview();});
async function formatted(save){const payload={filename:name,config:config(),outline:$('outline').value,decisions:Object.fromEntries(rows.map(r=>[r.index,r.level]))};const key=JSON.stringify(payload);if(!save&&resultCache?.key===key)return resultCache.result;const result=await request('/api/format',{...payload,document:selectedFile,save_history:save});resultCache={key,result};return result;}
$('sampleRules').onclick=async()=>{if(busy)return;const file=$('sampleFile').files[0];if(!file)return status('Select a sample DOCX first.',true);if(file.size>15*1024*1024)return status('Maximum file size is 15 MB.',true);lock(true);status('Reading sample formatting…');try{const result=await request('/api/sample',{document:file});suggested=validSettings(result.rules);const root=$('ruleReview');root.replaceChildren();for(const [k,v] of Object.entries(suggested)){const l=node('label');const i=node('input');i.type='checkbox';i.checked=true;i.dataset.rule=k;l.append(i,document.createTextNode(k+': '+v),node('small',result.evidence?.[k]||''));root.append(l);}for(const w of result.warnings||[])root.append(node('p',w,'muted'));$('applyRules').disabled=!Object.keys(suggested).length;if(!Object.keys(suggested).length){status('No supported sample settings detected.',true);return;}samplePrevious=config();sampleBaseline={...suggested};sampleName=file.name;setConfig(sampleBaseline);$('presetSelect').value='';if(!$('presetName').value.trim())$('presetName').value=file.name.replace(/\.(docx|dotx)$/i,'').slice(0,60);status('Sample formatting applied. Edit manual settings below and save one complete preset.');$('sampleState').scrollIntoView({block:'nearest',behavior:'auto'});}catch(e){status(e.message,true);}finally{lock(false);sampleState();}};
$('sampleFile').addEventListener('change',()=>{if($('sampleFile').files[0]&&!busy&&capabilities.inspection)$('sampleRules').click();});
$('references').onclick=async()=>{if(busy||!selectedFile)return;lock(true);status('Checking references…');try{const r=await documentRequest('/api/references',{});const root=$('referenceResult');root.replaceChildren(node('p',r.reference_entries+' reference entries · '+r.duplicate_entries.length+' duplicate candidates · '+r.doi_candidates.length+' DOI candidates'),node('p',r.note,'muted'));for(const t of [...r.duplicate_entries,...r.doi_candidates])root.append(node('div',t,'issue'));status('Reference check complete.');}catch(e){status(e.message,true);}finally{lock(false);}};

async function documentRequest(path,payload){
  if(uploadId){try{return await request(path,{...payload,upload_id:uploadId});}catch(e){if(e.code!=='upload_expired')throw e;uploadId=null;}}
  if(capabilities.upload_sessions&&selectedFile){const a=await request('/api/analyze',{document:selectedFile,outline:$('outline').value,cache_upload:true});uploadId=a.upload_id||null;if(uploadId)return request(path,{...payload,upload_id:uploadId});}
  return request(path,{...payload,document:selectedFile});
}
$('batchRun').onclick=async()=>{
 if(busy)return;for(const row of $('batchResults').querySelectorAll('[data-url]'))URL.revokeObjectURL(row.dataset.url);const files=Array.from($('batchFiles').files);batchEntries=[];$('batchZip').disabled=true;
 if(!files.length||files.length>5)return status('Select 1–5 DOCX files.',true);
 if(files.some(f=>f.size>15*1024*1024||!/\.docx$/i.test(f.name)))return status('Each file must be a DOCX under 15 MB.',true);
 lock(true);$('batchProgress').hidden=false;$('batchProgress').max=files.length;$('batchProgress').value=0;const root=$('batchResults');root.replaceChildren();const cfg=config();
 try{for(const [i,file] of files.entries()){
   const row=node('div',file.name+' — processing…','issue');root.append(row);status('Batch '+(i+1)+' / '+files.length);
   try{const a=await request('/api/analyze',{document:file,cache_upload:!!capabilities.upload_sessions});const decisions=Object.fromEntries(a.rows.filter(r=>!r.protected&&r.confidence==='high').map(r=>[r.index,r.level]));let result;
   try{result=await request('/api/format',{...(a.upload_id?{upload_id:a.upload_id}:{document:file}),filename:file.name,config:cfg,decisions,save_history:false});}finally{if(a.upload_id)fetch('/api/uploads/'+encodeURIComponent(a.upload_id),{method:'DELETE',headers:{'X-ReportStudio':'1'}}).catch(()=>{});}
   const bytes=Uint8Array.from(atob(result.document),c=>c.charCodeAt(0));const url=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}));const link=node('a','Download DOCX');link.href=url;link.download=file.name.replace(/\.docx$/i,'')+'_formatted.docx';row.replaceChildren(node('p',file.name+' — ready; review headings in Word'),link);row.dataset.url=url;batchEntries.push({name:link.download,bytes});
   }catch(e){row.textContent=file.name+' — '+e.message;}finally{$('batchProgress').value=i+1;$('batchZip').disabled=!batchEntries.length;}
 }status('Batch processing finished.');}finally{lock(false);}
};
$('batchFiles').onchange=()=>{batchEntries=[];$('batchZip').disabled=true;for(const row of $('batchResults').querySelectorAll('[data-url]'))URL.revokeObjectURL(row.dataset.url);$('batchResults').replaceChildren();};
$('doiVerify').onclick=async()=>{if(busy||!selectedFile)return;lock(true);const root=$('doiResults');root.replaceChildren();try{const refs=await documentRequest('/api/references',{});const dois=refs.doi_candidates; if(!dois.length){root.append(node('p','No DOI candidates found.'));return;}for(let i=0;i<dois.length;i+=10){status('Checking DOI '+(i+1)+'–'+Math.min(i+10,dois.length)+' / '+dois.length);const batch=await request('/api/doi',{dois:dois.slice(i,i+10)});for(const r of batch.results){const line=node('div',undefined,'issue');line.append(node('p',r.doi+' — '+r.status+(r.source?' · '+r.source:'')));if(r.title)line.append(node('small',r.title+(r.year?' ('+r.year+')':'')));root.append(line);}}status('DOI check complete.');}catch(e){status(e.message,true);}finally{lock(false);}};

$('comparePreview').onclick=()=>{compareMode=!compareMode;renderPreview();loadFullPreview();};
$('zoom').onchange=renderPreview;
$('refreshPreview').onclick=()=>loadFullPreview(true);
$('expandPreview').onclick=()=>{const view=$('textPreviewView');const expanded=view.classList.toggle('preview-expanded');document.body.classList.toggle('preview-open',expanded);$('expandPreview').textContent=expanded?'Close full screen':'Full screen';};
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&$('textPreviewView').classList.contains('preview-expanded'))$('expandPreview').click();});
function downloadBlob(blob,filename){const url=URL.createObjectURL(blob);const a=node('a');a.href=url;a.download=filename;a.click();setTimeout(()=>URL.revokeObjectURL(url),2000);}
$('changeReport').onclick=()=>{if(lastReport)downloadBlob(new Blob([JSON.stringify(lastReport,null,2)],{type:'application/json'}),'formatting-report.json');};
$('exportPreset').onclick=()=>downloadBlob(new Blob([JSON.stringify(config(),null,2)],{type:'application/json'}),'reportstudio-preset.json');
$('importPreset').onchange=async()=>{try{const file=$('importPreset').files[0];if(!file||file.size>64000)throw Error('Invalid preset file.');const obj=JSON.parse(await file.text());if(!obj||typeof obj!=='object'||Array.isArray(obj))throw Error('Invalid preset file.');sampleBaseline=null;samplePrevious=null;sampleName='';setConfig(validSettings(obj));status('Preset imported.');}catch(e){status(e.message,true);}finally{$('importPreset').value='';}};
function crc32(bytes){let crc=-1;for(const b of bytes){crc^=b;for(let i=0;i<8;i++)crc=(crc>>>1)^((crc&1)?0xedb88320:0);}return (crc^-1)>>>0;}
function makeZip(entries){const enc=new TextEncoder(),parts=[],central=[];let offset=0,centralSize=0;const names=new Set();for(const item of entries){let filename=item.name.replace(/[\\/]/g,'_'),n=1;while(names.has(filename))filename=item.name.replace(/\.docx$/i,'')+'_'+n+++'.docx';names.add(filename);const name=enc.encode(filename),bytes=item.bytes,crc=crc32(bytes);const h=new Uint8Array(30+name.length),v=new DataView(h.buffer);v.setUint32(0,0x04034b50,true);v.setUint16(4,20,true);v.setUint16(6,0x800,true);v.setUint16(12,33,true);v.setUint32(14,crc,true);v.setUint32(18,bytes.length,true);v.setUint32(22,bytes.length,true);v.setUint16(26,name.length,true);h.set(name,30);parts.push(h,bytes);const c=new Uint8Array(46+name.length),d=new DataView(c.buffer);d.setUint32(0,0x02014b50,true);d.setUint16(4,20,true);d.setUint16(6,20,true);d.setUint16(8,0x800,true);d.setUint16(14,33,true);d.setUint32(16,crc,true);d.setUint32(20,bytes.length,true);d.setUint32(24,bytes.length,true);d.setUint16(28,name.length,true);d.setUint32(42,offset,true);c.set(name,46);central.push(c);centralSize+=c.length;offset+=h.length+bytes.length;}const end=new Uint8Array(22),e=new DataView(end.buffer);e.setUint32(0,0x06054b50,true);e.setUint16(8,entries.length,true);e.setUint16(10,entries.length,true);e.setUint32(12,centralSize,true);e.setUint32(16,offset,true);return new Blob([...parts,...central,end],{type:'application/zip'});}
$('batchZip').onclick=()=>{if(batchEntries.length)downloadBlob(makeZip(batchEntries),'reportstudio-batch.zip');};
function tickHistory(){for(const e of document.querySelectorAll('[data-expires]')){const ms=Date.parse(e.dataset.expires)-Date.now();e.textContent=ms<=0?'Expired':Math.floor(ms/3600000)+'h '+Math.floor(ms/60000)%60+'m remaining';if(ms<=0){const a=e.closest('.item')?.querySelector('a');if(a)a.removeAttribute('href');}}}
setInterval(tickHistory,1000);

function showStep(step){
 if(step>0&&step<4&&!selectedFile){status('Upload and analyze a document first.',true);return;}
 if(step===3&&(!processed||processed.key!==previewKey())){status('Apply formatting before downloading.',true);return;}
 wizardStep=step;document.body.dataset.step=String(step);const labels=['Upload your Word document','Choose formatting settings','Review your document','Your formatted document','Batch formatting','Recent files'];$('heroTitle').textContent=labels[step];$('heroProgress').style.width=((Math.min(step,3)+1)*25)+'%';
 const groups=[['importPanel'],['formatPanel','requirementsPanel'],['reviewPanel','auditPanel'],['exportPanel'],['batchPanel'],['historyPanel']];
 for(const [i,ids] of groups.entries())for(const id of ids)$(id).hidden=i!==step;
 document.querySelectorAll('[data-step]').forEach(b=>{b.classList.toggle('active',Number(b.dataset.step)===step);b.setAttribute('aria-current',Number(b.dataset.step)===step?'step':'false');});
 $('wizardFile').textContent=selectedFile?.name||$('file').files[0]?.name||'No document selected';
 $('wizardState').textContent=step<4?'Step '+(step+1)+' of 4':'Tools';
 if(step===2)render();window.scrollTo({top:0,behavior:'instant'});
}
document.querySelectorAll('[data-step]').forEach(b=>b.onclick=()=>{if(!busy)showStep(Number(b.dataset.step));});
document.querySelectorAll('[data-go]').forEach(b=>b.onclick=()=>{if(!busy)showStep(Number(b.dataset.go));});
$('newDocument').onclick=()=>{if(busy)return;$('file').value='';$('file').dispatchEvent(new Event('change'));showStep(0);};
$('processDocument').onclick=async()=>{
 if(busy||!selectedFile)return;lock(true);processed=null;status('Formatting document…');
 try{
  const key=previewKey(),r=await formatted(false);
  if(key!==previewKey())throw Error('Settings changed during processing. Apply formatting again.');
  if(!r.report.integrity?.passed)throw Error('Document integrity check failed.');
  if(!r.report.input_sha256||!r.report.output_sha256)throw Error('Server version is outdated. Deploy the latest commit.');
  const sha=async bytes=>Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes)),x=>x.toString(16).padStart(2,'0')).join('');
  const inputHash=await sha(await selectedFile.arrayBuffer());
  const outputHash=await sha(Uint8Array.from(atob(r.document),c=>c.charCodeAt(0)));
  if(inputHash!==r.report.input_sha256||outputHash!==r.report.output_sha256)throw Error('Document verification failed. Analyze the original again.');
  if(inputHash===outputHash)throw Error('The output is identical to the original. No file downloaded.');
  if(r.report.change_report?.unchanged)throw Error('No formatting changes detected. Adjust settings and process again.');
  processed={key,result:r};lastReport=r.report;$('changeReport').disabled=false;
  const report=r.report;const summary=$('outputSummary');summary.replaceChildren();
  summary.append(node('strong','Formatted document ready'),node('p','Headings: '+report.heading+' · Body paragraphs: '+report.body+' · Tables: '+report.tables),node('p','Image references: '+(report.integrity.image_integrity?.before??0)+' → '+(report.integrity.image_integrity?.after??0)),node('p','Integrity: passed'),node('p','Applied: '+report.applied_settings.font+' · '+report.applied_settings.size+' pt · '+report.applied_settings.line+' line spacing'),node('p','Changed parts: '+report.change_report.changed_parts.join(', ')),node('p','Output SHA-256: '+report.output_sha256.slice(0,16)+'…'));
  $('result').textContent=(report.warnings||[]).join('\n');renderAudit(report.audit);showStep(3);status('Formatting complete. Download your new document.');
 }catch(e){status(e.message,true);}finally{lock(false);}
};
document.addEventListener('change',e=>{if(e.target.closest('#formatPanel,#requirementsPanel,#outlineView,#importPanel')){if(processed&&processed.key!==previewKey()){processed=null;$('export').disabled=true;}}});
showStep(0);
