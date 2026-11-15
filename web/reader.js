'use strict';
let readerState = null, readerGeneration = 0, readerImageURL = null;

async function openReader(doc, requestedPage) {
  const generation = ++readerGeneration;
  const info = await api('/api/reader?id=' + doc.id);
  if (generation !== readerGeneration) return;
  readerState = {doc, page:requestedPage ?? info.page, zoom:info.zoom, count:info.page_count, busy:false, query:'', text:false};
  $('reader-title').textContent=doc.filename;
  $('reader-count').textContent='of ' + info.page_count;
  $('reader-page').max=info.page_count;
  const zoomValue=String(info.zoom);
  if(!Array.from($('reader-zoom').options).some(option=>option.value===zoomValue)) $('reader-zoom').add(new Option(`${Math.round(info.zoom*100)}%`,zoomValue));
  $('reader-zoom').value=zoomValue;
  $('reader-find').value=''; $('reader-page-text').hidden=true;
  $('reader-text-toggle').setAttribute('aria-expanded','false');
  $('reader-outline').replaceChildren();
  for(const section of info.outline) {
    const button=el('button',section.title,'outline-link'); button.dataset.level=Math.min(3,section.level);
    button.title='Page ' + section.page; action(button,()=>loadReaderPage(section.page)); $('reader-outline').append(button);
  }
  if(!info.outline.length) $('reader-outline').append(el('p','This PDF has no document outline.','help'));
  await renderReaderBookmarks();
  if(!$('reader-dialog').open) $('reader-dialog').showModal();
  await loadReaderPage(readerState.page);
}
function readerControls(busy) {
  if(!readerState)return;
  readerState.busy=busy;
  $('reader-previous').disabled=busy || readerState.page<=1;
  $('reader-next').disabled=busy || readerState.page>=readerState.count;
  for(const id of ['reader-page','reader-zoom','reader-add-bookmark','reader-text-toggle']) $(id).disabled=busy;
  document.querySelector('.reader-page-panel').setAttribute('aria-busy',String(busy));
}
async function loadReaderPage(page) {
  if(!readerState || readerState.busy)return;
  if(!Number.isInteger(page) || page<1 || page>readerState.count) throw Error(`Choose a page between 1 and ${readerState.count}.`);
  const state=readerState, generation=readerGeneration;
  const zoom=Number($('reader-zoom').value), query=state.query;
  readerControls(true); $('reader-status').textContent=`Loading page ${page}…`;
  try {
    const params=new URLSearchParams({id:state.doc.id,page,zoom,q:query});
    const response=await fetch('/api/reader/page?'+params,{headers:{'X-Navigator-Token':token}});
    if(!response.ok){const result=await response.json();throw Error(result.error || 'Could not load PDF page.');}
    const blob=await response.blob();
    if(generation!==readerGeneration)return;
    const nextURL=URL.createObjectURL(blob);
    if(readerImageURL)URL.revokeObjectURL(readerImageURL);
    readerImageURL=nextURL;
    $('reader-image').src=nextURL; $('reader-image').alt=`${state.doc.filename}, page ${page} of ${state.count}`;
    $('reader-image').hidden=false;
    state.page=page;state.zoom=zoom;$('reader-page').value=page;
    $('reader-original').href=`/api/file?id=${state.doc.id}&token=${encodeURIComponent(token)}#page=${page}`;
    const matches=Number(response.headers.get('X-Match-Count'));
    $('reader-status').textContent=`Page ${page} of ${state.count}`+(query?` · ${matches}${matches===200?'+':''} highlighted match(es)`:'');
    $('reader-page-text').textContent='';
    if(state.text) await loadReaderText();
    if(generation!==readerGeneration)return;
    try {await api('/api/reader/position',{id:state.doc.id,page,zoom});}
    catch(error){toast('Page opened, but the reading position could not be saved: '+error.message);}
  } catch(error) {
    if(generation===readerGeneration){$('reader-status').textContent=error.message;$('reader-page').value=state.page;}
    throw error;
  } finally {
    if(generation===readerGeneration) readerControls(false);
  }
}
async function loadReaderText() {
  if(!readerState)return;
  const generation=readerGeneration, state=readerState, page=readerState.page;
  const result=await api(`/api/reader/text?id=${state.doc.id}&page=${page}`);
  if(generation!==readerGeneration || readerState.page!==page)return;
  $('reader-page-text').textContent=result.text || 'No extractable text on this page. An image-only scan needs OCR.';
}
async function renderReaderBookmarks() {
  if(!readerState)return;
  const state=readerState;
  const {bookmarks}=await api('/api/bookmarks?id='+state.doc.id);
  if(readerState!==state)return;
  $('reader-bookmarks').replaceChildren();
  for(const bookmark of bookmarks) {
    const button=el('button',`${bookmark.name} · p. ${bookmark.page_number+1}`,'outline-link');
    action(button,()=>loadReaderPage(bookmark.page_number+1));$('reader-bookmarks').append(button);
  }
  if(!bookmarks.length)$('reader-bookmarks').append(el('p','Bookmark useful pages for quick access.','help'));
}
async function openBookmarks() {
  const {bookmarks}=await api('/api/bookmarks');$('bookmarks-list').replaceChildren();
  for(const bookmark of bookmarks) {
    const row=el('article',undefined,'path-row'), summary=el('div'), open=el('button',bookmark.name,'saved-name'), remove=el('button','Delete bookmark');
    action(open,async()=>{
      const doc=await api('/api/document?id='+bookmark.doc_id);
      if(doc.filename.toLowerCase().endsWith('.pdf')) await openReader(doc,bookmark.page_number+1);
      else await showDocument(bookmark.doc_id,bookmark.page_number);
    });
    summary.append(open,el('p',`${bookmark.filename} · page ${bookmark.page_number+1}`,'help'));
    action(remove,async()=>{if(!confirm('Delete this page bookmark?'))return;await api('/api/bookmarks/delete',{bookmark_id:bookmark.fav_id});await openBookmarks();await refresh();});
    row.append(summary,remove);$('bookmarks-list').append(row);
  }
  if(!bookmarks.length)$('bookmarks-list').append(el('p','No page bookmarks yet. Open a PDF in the reader and choose Bookmark page.','help'));
  if(!$('bookmarks-dialog').open)$('bookmarks-dialog').showModal();
}

async function renderDocumentLinks(container, doc, page, links) {
  container.append(el('h3','Related documents','section-title'));
  for(const link of links) {
    const row=el('div',undefined,'path-row'), summary=el('div'), open=el('button',link.filename,'saved-name'), remove=el('button','Remove link');
    action(open,()=>showDocument(link.doc_id)); summary.append(open);
    if(link.description)summary.append(el('p',link.description,'help'));
    action(remove,async()=>{if(!confirm('Remove this document link?'))return;await api('/api/links/delete',{id:doc.id,link_id:link.link_id});await showDocument(doc.id,page);});
    row.append(summary,remove);container.append(row);
  }
  if(!links.length)container.append(el('p','Link this manual to a related SOP, datasheet, or service reference.','help'));
  const findForm=el('form',undefined,'link-search-form'), queryLabel=el('label','Find a document to link'), query=el('input');
  query.type='search';query.placeholder='Search by filename, model, or content';queryLabel.append(query);findForm.append(queryLabel,el('button','Find documents'));
  const linkForm=el('form',undefined,'link-add-form'), targetLabel=el('label','Matching document'), target=el('select'), descriptionLabel=el('label','Link description'), description=el('input'), add=el('button','Add document link','primary');
  target.required=true;target.append(new Option('Search for a document first',''));targetLabel.append(target);description.maxLength=2000;descriptionLabel.append(description);add.disabled=true;linkForm.append(targetLabel,descriptionLabel,add);
  findForm.addEventListener('submit',async event=>{
    event.preventDefault();try {
      const {documents}=await api('/api/documents?q='+encodeURIComponent(query.value));
      const candidates=documents.filter(item=>item.id!==doc.id&&!links.some(link=>link.doc_id===item.id));
      target.replaceChildren(new Option(candidates.length?'Choose a document':'No unlinked matches',''),...candidates.map(item=>new Option(item.filename,item.id)));
      add.disabled=!candidates.length;
    }catch(error){toast(error.message);}
  });
  linkForm.addEventListener('submit',async event=>{
    event.preventDefault();try {await api('/api/links',{id:doc.id,target_id:Number(target.value),description:description.value});await showDocument(doc.id,page);toast('Related document linked.');}catch(error){toast(error.message);}
  });
  container.append(findForm,linkForm);
}

action($('bookmarks-nav'),openBookmarks);
action($('reader-previous'),()=>loadReaderPage(readerState.page-1));
action($('reader-next'),()=>loadReaderPage(readerState.page+1));
$('reader-page-form').addEventListener('submit',event=>{event.preventDefault();loadReaderPage(Number($('reader-page').value)).catch(error=>toast(error.message));});
$('reader-zoom').addEventListener('change',()=>loadReaderPage(readerState.page).catch(error=>toast(error.message)));
$('reader-find-form').addEventListener('submit',event=>{event.preventDefault();if(!readerState)return;readerState.query=$('reader-find').value.trim();loadReaderPage(readerState.page).catch(error=>toast(error.message));});
action($('reader-find-clear'),()=>{if(!readerState)return;readerState.query='';$('reader-find').value='';return loadReaderPage(readerState.page);});
action($('reader-text-toggle'),async()=>{
  if(!readerState)return;readerState.text=!readerState.text;$('reader-page-text').hidden=!readerState.text;$('reader-text-toggle').setAttribute('aria-expanded',String(readerState.text));if(readerState.text)await loadReaderText();
});
action($('reader-add-bookmark'),()=>{
  if(!readerState)return;
  $('bookmark-name').value=`${readerState.doc.filename} — page ${readerState.page}`;
  $('bookmark-page-summary').textContent=`${readerState.doc.filename} · page ${readerState.page}`;
  $('bookmark-save-dialog').showModal();
});
$('bookmark-form').addEventListener('submit',async event=>{
  event.preventDefault();if(!readerState)return;
  try{await api('/api/bookmarks',{id:readerState.doc.id,page:readerState.page,name:$('bookmark-name').value});$('bookmark-save-dialog').close();await renderReaderBookmarks();await refreshStats();toast('Page bookmark saved.');}catch(error){toast(error.message);}
});
$('reader-dialog').addEventListener('close',()=>{
  readerGeneration++;readerState=null;
  if(readerImageURL)URL.revokeObjectURL(readerImageURL);readerImageURL=null;$('reader-image').removeAttribute('src');$('reader-image').hidden=true;
});
$('reader-dialog').addEventListener('keydown',event=>{
  if(!readerState||readerState.busy||['INPUT','TEXTAREA','SELECT'].includes(document.activeElement.tagName))return;
  if(event.key==='ArrowRight'&&readerState.page<readerState.count){event.preventDefault();loadReaderPage(readerState.page+1).catch(error=>toast(error.message));}
  if(event.key==='ArrowLeft'&&readerState.page>1){event.preventDefault();loadReaderPage(readerState.page-1).catch(error=>toast(error.message));}
});
