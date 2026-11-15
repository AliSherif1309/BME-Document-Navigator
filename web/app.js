'use strict';
const $ = id => document.getElementById(id);
const token = document.querySelector('meta[name="navigator-token"]').content;
let offset = 0, total = 0, favorites = false, selected = null, searchVersion = 0;
let wasRunning = false, toastTimer, searchTimer;
const labels = {manufacturer:'Manufacturer', device_model:'Device model', document_type:'Document type', revision_number:'Revision', revision_date:'Revision date', status:'Document status', applicable_models:'Applicable models', associated_test_equipment:'Test equipment', keywords:'Keywords'};

function el(tag, text, className) {
  const item = document.createElement(tag);
  if (text !== undefined) item.textContent = text;
  if (className) item.className = className;
  return item;
}
function toast(message) {
  $('toast').textContent = message; $('toast').hidden = false;
  clearTimeout(toastTimer); toastTimer = setTimeout(() => $('toast').hidden = true, 6000);
}
async function api(path, body) {
  const response = await fetch(path, {method:body ? 'POST' : 'GET', headers:{'X-Navigator-Token':token, ...(body ? {'Content-Type':'application/json'} : {})}, ...(body ? {body:JSON.stringify(body)} : {})});
  const data = await response.json();
  if (!response.ok) throw Error(data.error || 'Request failed.');
  return data;
}
function action(button, fn) { button.addEventListener('click', () => Promise.resolve().then(fn).catch(error => toast(error.message))); }
function download(path) { const a = el('a'); a.href = `${path}${path.includes('?') ? '&' : '?'}token=${encodeURIComponent(token)}`; a.target = '_blank'; a.rel = 'noopener'; a.click(); }
function chooseOptions(id, values, placeholder) {
  const select = $(id), previous = select.value;
  select.replaceChildren(new Option(placeholder, ''), ...values.map(value => new Option(value,value)));
  if (values.includes(previous)) select.value = previous;
}
async function refreshStats() {
  const stats = await api('/api/stats');
  for (const key of ['documents','indexed','favorites','notes']) $(`stat-${key}`).textContent = stats[key].toLocaleString();
  chooseOptions('manufacturer', stats.manufacturers, 'All manufacturers');
  chooseOptions('kind', stats.types, 'All document types');
  $('folder-list').replaceChildren();
  for (const path of stats.paths) {
    const row = el('div',undefined,'path-row'), remove = el('button','Remove');
    row.append(el('span',path),remove);
    action(remove, async () => { await api('/api/paths/remove',{path}); await refreshStats(); toast('Folder removed. Library history retained.'); });
    $('folder-list').append(row);
  }
  if (!stats.paths.length) $('folder-list').append(el('p','No folders yet. Add a folder to start your library.','help'));
}
async function search() {
  const version = ++searchVersion;
  const params = new URLSearchParams({q:$('query').value, manufacturer:$('manufacturer').value, kind:$('kind').value, favorites:favorites?'1':'0', offset});
  const data = await api('/api/documents?' + params);
  if (version !== searchVersion) return;
  total = data.total;
  $('result-count').textContent = `${total.toLocaleString()} document${total===1?'':'s'}`;
  $('documents').replaceChildren();
  for (const doc of data.documents) {
    const row = el('article',undefined,'document-row');
    const summary = el('div',undefined,'doc-summary');
    const title = el('button',doc.filename,'doc-title');
    action(title, () => showDocument(doc.id, doc.page_number));
    const meta = el('div',undefined,'doc-meta');
    meta.append(el('span',doc.manufacturer || 'Manufacturer unassigned'), el('span',doc.device_model || 'Model unassigned'));
    if (doc.document_type) meta.append(el('span',doc.document_type,'badge'));
    if (doc.revision_number) meta.append(el('span','Rev. ' + doc.revision_number));
    summary.append(title,meta);
    if (doc.snippet) summary.append(el('p',`Page ${Number(doc.page_number)+1} · ${doc.snippet}`,'snippet'));
    const star = el('button',doc.favorite?'★':'☆','star'); star.setAttribute('aria-label',doc.favorite?'Remove favorite':'Add favorite');
    action(star, async () => { await api('/api/favorite',{id:doc.id}); await Promise.all([search(),refreshStats()]); });
    row.append(el('span',doc.filename.split('.').pop().slice(0,5),'doc-icon'),summary,star);
    $('documents').append(row);
  }
  if (!data.documents.length) {
    const empty = el('div',undefined,'empty');
    empty.append(el('div','▤','empty-icon'),el('h2',favorites?'Keep your key references close':($('query').value?'No matching documents':'Your library starts here')),
      el('p',favorites?'Use the star beside a document to save it here.':'Add your document folders, scan the library, then search filenames, equipment metadata, and the text inside your manuals.'));
    const button = el('button','Manage library folders','primary'); action(button,()=> $('folders-dialog').showModal()); empty.append(button); $('documents').append(empty);
  }
  $('previous').disabled = offset === 0; $('next').disabled = offset+100 >= total;
  $('page-number').textContent = total ? `${offset+1}–${Math.min(offset+100,total)} of ${total}` : '0 documents';
}
async function refresh() { await refreshStats(); await search(); }
async function showDocument(id, page=0) {
  const doc = await api('/api/document?id=' + id); selected = id;
  $('document-title').textContent = doc.filename;
  const content = $('document-content'); content.replaceChildren();
  content.append(el('p',doc.filepath,'file-path'));
  const actions = el('div',undefined,'detail-actions');
  const open = el('a',doc.filename.toLowerCase().endsWith('.pdf')?'Open PDF':'Download document','link-button');
  open.href = `/api/file?id=${id}&token=${encodeURIComponent(token)}#page=${Number(page)+1}`; open.target='_blank'; open.rel='noopener';
  if (doc.available) actions.append(open); else actions.append(el('span','File offline — reconnect its original folder.','badge'));
  const star = el('button',doc.favorite?'★ Saved':'☆ Save favorite');
  action(star,async()=>{await api('/api/favorite',{id}); await showDocument(id,page); await refresh();});
  actions.append(star); content.append(actions);
  content.append(el('p',doc.indexed_pages?`${doc.indexed_pages} text page(s) indexed. Verify document revisions against your approved source before service.`:'No searchable text indexed. Image-only PDFs require OCR; other file types may be listed by metadata only.','help'));
  const form = el('form'), grid = el('div',undefined,'metadata-grid');
  for (const [field,label] of Object.entries(labels)) {
    const item = el('label',label), input = el('input'); input.name=field; input.value=doc[field]||''; input.maxLength=4000;
    item.append(input); grid.append(item);
  }
  const save = el('button','Save metadata','primary'); form.append(grid,save);
  form.addEventListener('submit',async event=>{event.preventDefault();try{await api('/api/metadata',{id,...Object.fromEntries(new FormData(form))}); await refresh(); toast('Metadata saved.');}catch(error){toast(error.message);}});
  content.append(el('h3','Equipment & document metadata','section-title'),form,el('h3','Engineering notes','section-title'));
  for (const note of doc.notes) {
    const block = el('div',note.note_text,'note'), remove=el('button','Delete note');
    block.append(el('small',`Page ${Number(note.page_number || 0)+1} · ${new Date(note.created_timestamp*1000).toLocaleString()}`));
    action(remove,async()=>{if(!confirm('Delete this note?'))return;await api('/api/notes/delete',{note_id:note.note_id});await showDocument(id,page);await refreshStats();});
    block.append(remove); content.append(block);
  }
  if (!doc.notes.length) content.append(el('p','No notes yet. Record references and observations here.','help'));
  const noteForm=el('form',undefined,'note-form'), textLabel=el('label','Note'), text=el('textarea'), pageLabel=el('label','Page','page-input'), pageInput=el('input');
  text.required=true; text.maxLength=20000; text.rows=3; text.name='text'; textLabel.append(text);
  pageInput.type='number'; pageInput.min='1'; pageInput.required=true; pageInput.value=Number(page)+1; pageLabel.append(pageInput);
  noteForm.append(textLabel,pageLabel,el('button','Add note','primary'));
  noteForm.addEventListener('submit',async event=>{event.preventDefault();try{await api('/api/notes',{id,text:text.value,page:Number(pageInput.value)});await showDocument(id,page);await refreshStats();}catch(error){toast(error.message);}});
  content.append(noteForm);
  if (!$('document-dialog').open) $('document-dialog').showModal();
}
async function poll() {
  try {
    const state = await api('/api/scan');
    $('scan-status').textContent = state.message + (state.result ? ` ${state.result.added} added · ${state.result.updated} updated · ${state.result.errors} errors` : '');
    $('scan').disabled=state.running; $('scan').classList.toggle('busy',state.running);
    $('scan').textContent=state.running?'↻ Scanning…':'↻ Scan library';
    $('scan-report').textContent = state.result ? JSON.stringify(state.result,null,2) : state.message;
    if (wasRunning && !state.running) await refresh();
    wasRunning=state.running;
  } catch(error) { $('scan-status').textContent='Connection lost. Keep the Navigator application running, then reload this page.'; }
  setTimeout(poll,1500);
}
action($('scan'),async()=>{await api('/api/scan',{});wasRunning=true;});
action($('folders-nav'),()=> $('folders-dialog').showModal());
action($('scan-details'),()=> $('scan-dialog').showModal());
action($('export'),()=> download('/api/export'));
action($('backup'),()=> download('/api/backup'));
for (const [id,value] of [['library-nav',false],['favorites-nav',true]]) action($(id),async()=>{
  favorites=value; offset=0; $('page-title').textContent=value?'Favorite documents':'Document library';
  $('library-nav').classList.toggle('active',!value); $('favorites-nav').classList.toggle('active',value); await search();
});
document.querySelectorAll('[data-close]').forEach(button=>action(button,()=>$(button.dataset.close).close()));
$('folder-form').addEventListener('submit',async event=>{event.preventDefault();try{await api('/api/paths/add',{path:$('folder-path').value.trim()});$('folder-path').value='';await refreshStats();await api('/api/scan',{});wasRunning=true;toast('Folder added. Scanning your documents.');}catch(error){toast(error.message);}});
$('query').addEventListener('input',()=>{clearTimeout(searchTimer);searchTimer=setTimeout(()=>{offset=0;search().catch(error=>toast(error.message));},250);});
for(const id of ['manufacturer','kind']) $(id).addEventListener('change',()=>{offset=0;search().catch(error=>toast(error.message));});
action($('clear'),async()=>{$('query').value='';$('manufacturer').value='';$('kind').value='';offset=0;await search();});
action($('previous'),async()=>{offset=Math.max(0,offset-100);await search();});
action($('next'),async()=>{offset+=100;await search();});
document.addEventListener('keydown',event=>{if(event.key==='/'&&!document.querySelector('dialog[open]')&&!['INPUT','TEXTAREA','SELECT'].includes(document.activeElement.tagName)){event.preventDefault();$('query').focus();}});
refresh().catch(error=>toast(error.message)); poll();
