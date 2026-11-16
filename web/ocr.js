'use strict';
let ocrSetup=null, ocrWasRunning=false;
const ocrLanguageNames={eng:'English',ara:'Arabic',fra:'French',deu:'German',spa:'Spanish',ita:'Italian',por:'Portuguese',rus:'Russian',chi_sim:'Chinese (simplified)'};

function showOCRSetup(setup) {
  ocrSetup=setup;$('ocr-tessdata').value=setup.tessdata;
  $('ocr-setup-status').textContent=setup.message;
  $('ocr-language-list').replaceChildren();
  for(const code of setup.languages) {
    const label=el('label',undefined,'check-label'), input=el('input');input.type='checkbox';input.value=code;input.name='language';input.checked=code==='eng';
    label.append(input,el('span',`${ocrLanguageNames[code]||code} (${code})`));$('ocr-language-list').append(label);
  }
  if(setup.languages.length&&!setup.languages.includes('eng'))$('ocr-language-list').querySelector('input').checked=true;
  if(!setup.available)$('ocr-language-list').append(el('p','Install language data, then choose its tessdata folder above. See the README for setup on Windows, macOS, and Linux.','help'));
  $('ocr-start').disabled=!setup.available;
}
async function openOCR(doc=null) {
  const [setup,result]=await Promise.all([api('/api/ocr/setup'),api('/api/documents')]);
  showOCRSetup(setup);
  const documents=result.documents.filter(item=>item.filename.toLowerCase().endsWith('.pdf'));
  if(doc&&!documents.some(item=>item.id===doc.id))documents.unshift(doc);
  $('ocr-document').replaceChildren(new Option('Choose a PDF',''),...documents.map(item=>new Option(item.filename,item.id)));
  if(doc)$('ocr-document').value=doc.id;
  if(!$('ocr-dialog').open)$('ocr-dialog').showModal();
  await updateOCRStatus();
}
async function updateOCRStatus() {
  const state=await api('/api/ocr');
  $('ocr-job-status').textContent=(state.filename?state.filename+' · ':'')+state.message;
  $('ocr-progress').max=Math.max(1,state.total);$('ocr-progress').value=state.page||0;
  $('ocr-cancel').disabled=!state.running;
  $('ocr-start').disabled=state.running || !ocrSetup?.available;
  for(const id of ['ocr-document','ocr-mode','ocr-dpi','ocr-languages'])$(id).disabled=state.running;
  $('ocr-job-summary').textContent=state.result&&!state.result.error&&!state.result.cancelled?`${state.result.ocr_pages} page(s) recognized · ${state.result.skipped} native-text page(s) retained · ${state.result.language} · ${state.result.dpi} DPI`:state.running?`Page ${state.page} of ${state.total || '…'}. Cancellation takes effect after the current page.`:'';
  if(ocrWasRunning&&!state.running){
    await refresh();
    if(readerState&&readerState.doc.id===state.doc_id&&!readerState.busy)await loadReaderPage(readerState.page);
  }
  ocrWasRunning=state.running;
}
action($('ocr-nav'),()=>openOCR());
$('ocr-setup-form').addEventListener('submit',async event=>{
  event.preventDefault();try{showOCRSetup(await api('/api/ocr/setup',{tessdata:$('ocr-tessdata').value}));await updateOCRStatus();}catch(error){toast(error.message);}
});
$('ocr-form').addEventListener('submit',async event=>{
  event.preventDefault();
  const languages=Array.from($('ocr-language-list').querySelectorAll('input:checked')).map(input=>input.value);
  try{await api('/api/ocr',{id:Number($('ocr-document').value),languages,dpi:Number($('ocr-dpi').value),mode:$('ocr-mode').value});ocrWasRunning=true;await updateOCRStatus();}catch(error){toast(error.message);}
});
action($('ocr-cancel'),async()=>{await api('/api/ocr/cancel',{});await updateOCRStatus();});
async function pollOCR(){try{await updateOCRStatus();}catch(error){if($('ocr-dialog').open)$('ocr-job-status').textContent='Keep the Navigator app running to check recognition progress.';}setTimeout(pollOCR,1500);}
pollOCR();
