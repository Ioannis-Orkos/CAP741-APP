    (function(){
      /*
       * CAP 741 Logbook App
       *
       * High-level flow:
       * 1. Load workbook data from cap741-data.xlsx into in-memory row objects.
       * 2. Group rows by aircraft type + ATA chapter.
       * 3. Paginate each group into CAP 741-sized pages.
       * 4. Re-render the page HTML after edits.
       * 5. Write the current in-memory state back to the workbook.
       *
       * Why the file feels "single page":
       * - HTML holds the shell and modal containers.
       * - CSS handles the printable CAP 741 layout.
       * - This file owns data loading, rendering, editing, filtering, and saving.
       *
       * Navigation guide:
       * - constants/state: top of file
       * - filters: lines near the "Filter state" section
       * - row model + rendering: the middle of the file
       * - workbook I/O + autosave: lower-middle
       * - event wiring + startup: end of file
       */
      var rows = [];
      var AIRCRAFT_MAP = Object.create(null);
      var CHAPTER_OPTIONS = [];
      var LOG_OWNER_INFO = { name: '', signature: '', stamp: '' };
      var SUPERVISOR_OPTIONS = [];
      var SUPERVISOR_LOOKUP = Object.create(null);
      var AIRCRAFT_GROUP_ROWS = [];
      var SUPERVISOR_RECORDS = [];
      var PAGE_SLOTS = 6;
      var ROW_TASK_CHARS = 60;
      var ROW_LINES_PER_SLOT = 5;
      var DB_NAME = 'cap741-file-handles';
      var DB_STORE = 'handles';
      var LINKED_FILE_KEY = 'cap741-main-file';
      var DEFAULT_WORKBOOK_PATH = './cap741-data.xlsx';
      var BLANK_CHAPTER_FILTER = 'No Chapter';
      var LOG_HEADERS = ['Aircraft Type','A/C Reg','Chapter','Chapter Description','Date','Job No','FAULT','Task Detail','Rewriten for cap741','Approval Name','Approval stamp','Aprroval Licence No.'];
      var DATE_PLACEHOLDER = 'dd/MMM/yyyy';
      var FILTER_KEYS = ['aircraftType','aircraftReg','supervisor','chapter'];
      var NEW_WORKBOOK_SUPERVISOR_NAME = 'Ioannis Orkos';
      var NEW_WORKBOOK_SUPERVISOR_LICENCE = 'UK.XX.XXXXXXX';
      var NEW_WORKBOOK_TASK_TEXT = 'Dummy data test one';
      var NEW_WORKBOOK_OWNER_NAME = 'User User';

      // ---- DOM refs ----
      var errorBox = document.getElementById('errorBox');
      var loadingOverlay = document.getElementById('loadingOverlay');
      var loadingTitleEl = document.getElementById('loadingTitle');
      var loadingTextEl = document.getElementById('loadingText');
      var pagesEl = document.getElementById('pages');
      var searchInput = document.getElementById('searchInput');
      var clearSearchBtn = document.getElementById('clearSearchBtn');
      var loadBtn = document.getElementById('loadBtn');
      var loadOptionsEl = document.getElementById('loadOptions');
      var loadExistingBtn = document.getElementById('loadExistingBtn');
      var createNewWorkbookBtn = document.getElementById('createNewWorkbookBtn');
      var filterBtn = document.getElementById('filterBtn');
      var filterCountEl = document.getElementById('filterCount');
      var filterStripEl = document.getElementById('filterStrip');
      var addBtn = document.getElementById('addBlankPage');
      var printBtn = document.getElementById('printBtn');
      var printOptionsEl = document.getElementById('printOptions');
      var printCurrentBtn = document.getElementById('printCurrentBtn');
      var printAllBtn = document.getElementById('printAllBtn');
      var saveFileBtn = document.getElementById('saveFileBtn');
      var infoBtn = document.getElementById('infoBtn');
      var modal = document.getElementById('blankModal');
      var modalBody = document.getElementById('blankModalBody');
      var infoModal = document.getElementById('infoModal');
      var closeInfoModalBtn = document.getElementById('closeInfoModal');
      var filterModal = document.getElementById('filterModal');
      var closeFilterPanelBtn = document.getElementById('closeFilterPanel');
      var filterForm = document.getElementById('filterForm');
      var filterAircraftTypeInput = document.getElementById('filterAircraftType');
      var filterAircraftRegInput = document.getElementById('filterAircraftReg');
      var filterSupervisorInput = document.getElementById('filterSupervisor');
      var filterChapterInput = document.getElementById('filterChapter');
      var filterAircraftTypeChipsEl = document.getElementById('filterAircraftTypeChips');
      var filterAircraftRegChipsEl = document.getElementById('filterAircraftRegChips');
      var filterSupervisorChipsEl = document.getElementById('filterSupervisorChips');
      var filterChapterChipsEl = document.getElementById('filterChapterChips');
      var clearFiltersBtn = document.getElementById('clearFilters');
      var filterAircraftRegListEl = document.getElementById('filter-aircraft-reg-list');
      var filterChapterListEl = document.getElementById('filter-chapter-list');
      var taskDetailModal = document.getElementById('taskDetailModal');
      var confirmModal = document.getElementById('confirmModal');
      var confirmTitleEl = document.getElementById('confirmTitle');
      var confirmTextEl = document.getElementById('confirmText');
      var confirmCancelBtn = document.getElementById('confirmCancelBtn');
      var confirmOkBtn = document.getElementById('confirmOkBtn');
      var closeTaskDetailBtn = document.getElementById('closeTaskDetail');
      var detailFaultEl = document.getElementById('detailFault');
      var detailTaskEl = document.getElementById('detailTask');
      var detailRewriteEl = document.getElementById('detailRewrite');
      var sharedListsEl = document.getElementById('sharedLists');
      var saveBtn = document.getElementById('saveBlankPage');
      var cancelBtn = document.getElementById('cancelBlankPage');
      var modalShell = modal.querySelector('.modal-shell');
      var modalActions = modal.querySelector('.modal-actions');
      var settingsBtn = document.getElementById('settingsBtn');
      var settingsModal = document.getElementById('settingsModal');
      var closeSettingsModalBtn = document.getElementById('closeSettingsModal');
      var settingsBodyEl = document.getElementById('settingsBody');
      var saveSettingsBtn = document.getElementById('saveSettingsBtn');
      var saveTaskDetailBtn = document.getElementById('saveTaskDetail');
      var detailChapterEl = document.getElementById('detailChapter');

      // ---- Runtime state ----
      var layoutTimer = null;
      var autoSaveTimer = null;
      var liveLayoutEditorState = null;
      var saveInFlight = false;
      var saveQueued = false;
      var hasUnsavedChanges = false;
      var lastSavedLogbookText = '';
      var confirmResolver = null;
      var settingsDirty = false;
      var lastTaskDetailFocus = null;
      var lastTaskDetailRowId = null;
      var taskDetailRewriteDirty = false;
      var taskDetailOriginalState = null;
      var rowsById = Object.create(null);
      var nextRowIdValue = 0;
      var sharedDatalistsCache = '';
      var activeFilters = emptyFilterState();
      var draftFilters = emptyFilterState();
      var searchQuery = '';
      var printMode = '';
      var settingsActiveTab = 'owner';
      var loadButtonMode = 'load';

      // ---- Filter state ----
      function emptyFilterState(){ return { aircraftType:[], aircraftReg:[], supervisor:[], chapter:[] }; }
      function cloneFilterState(state){ return { aircraftType:(state.aircraftType||[]).slice(), aircraftReg:(state.aircraftReg||[]).slice(), supervisor:(state.supervisor||[]).slice(), chapter:(state.chapter||[]).slice() }; }
      function filterValues(state, key){ return (state && state[key] && state[key].length) ? state[key] : []; }
      function totalFilterValueCount(state){ var total=0; for(var i=0;i<FILTER_KEYS.length;i++) total += filterValues(state,FILTER_KEYS[i]).length; return total; }
      function hasActiveFilters(){ return totalFilterValueCount(activeFilters) > 0; }
      function hasActiveSearch(){ return !!s(searchQuery); }
      function activeFilterCount(){ return totalFilterValueCount(activeFilters); }
      function activeFilterChips(){ var chips=[]; for(var i=0;i<activeFilters.aircraftType.length;i++) chips.push({label:'Type',value:activeFilters.aircraftType[i]}); for(var j=0;j<activeFilters.aircraftReg.length;j++) chips.push({label:'A/C',value:activeFilters.aircraftReg[j]}); for(var k=0;k<activeFilters.supervisor.length;k++) chips.push({label:'Supervisor',value:activeFilters.supervisor[k]}); for(var m=0;m<activeFilters.chapter.length;m++) chips.push({label:'Chapter',value:activeFilters.chapter[m]}); return chips; }
      function syncFilterButtonState(){ if(!filterBtn||!filterCountEl) return; var count=activeFilterCount(); filterBtn.classList.toggle('active',count>0); filterCountEl.hidden=count<1; filterCountEl.textContent=String(count); }
      function syncSearchUi(){ if(!searchInput||!clearSearchBtn) return; searchInput.value=searchQuery; clearSearchBtn.hidden=!hasActiveSearch(); }
      function renderFilterStrip(){ if(!filterStripEl) return; var chips=activeFilterChips(); if(!chips.length){ filterStripEl.className='filter-strip'; filterStripEl.innerHTML=''; return; } var html=[]; for(var i=0;i<chips.length;i++) html.push('<span class="filter-chip">'+esc(chips[i].label+': '+chips[i].value)+'</span>'); filterStripEl.className='filter-strip open'; filterStripEl.innerHTML='<div class="filter-strip-text"><strong>Filters:</strong> '+html.join('')+'</div><button type="button" data-clear-filters="1">Clear filters</button>'; }
      function normalizeFilterEntry(key, value){ var raw=s(value); if(!raw) return ''; if(key==='aircraftReg') return raw.toUpperCase(); if(key==='supervisor'){ var sv=normalizeSupervisorValue(raw); return sv.name||raw; } if(key==='chapter'){ var normalized=normalizedText(raw); if(normalized==='no chapter'||normalized==='blank chapter'||normalized==='empty chapter'||normalized==='[no chapter]') return BLANK_CHAPTER_FILTER; } return raw; }
      function uniqueFilterValues(values){ var out=[],seen=Object.create(null); for(var i=0;i<values.length;i++){ var n=normalizedText(values[i]); if(!n||seen[n]) continue; seen[n]=true; out.push(values[i]); } return out; }
      function filterInputForKey(key){ if(key==='aircraftType') return filterAircraftTypeInput; if(key==='aircraftReg') return filterAircraftRegInput; if(key==='supervisor') return filterSupervisorInput; if(key==='chapter') return filterChapterInput; return null; }
      function filterChipHostForKey(key){ if(key==='aircraftType') return filterAircraftTypeChipsEl; if(key==='aircraftReg') return filterAircraftRegChipsEl; if(key==='supervisor') return filterSupervisorChipsEl; if(key==='chapter') return filterChapterChipsEl; return null; }
      function aircraftOptionsHtmlForTypes(types){ if(!types||!types.length) return aircraftOptionsHtml(); var seen=Object.create(null),regs=[]; for(var reg in AIRCRAFT_MAP){ if(!Object.prototype.hasOwnProperty.call(AIRCRAFT_MAP,reg)) continue; for(var i=0;i<types.length;i++){ if(AIRCRAFT_MAP[reg]===types[i]&&!seen[reg]){ seen[reg]=true; regs.push(reg); } } } regs.sort(); if(!regs.length) return aircraftOptionsHtml(); var html=''; for(var j=0;j<regs.length;j++) html+='<option value="'+esc(regs[j])+'"></option>'; return html; }
      function syncFilterRegList(){ if(!filterAircraftRegListEl) return; filterAircraftRegListEl.innerHTML=aircraftOptionsHtmlForTypes(draftFilters.aircraftType); }
      function syncFilterChapterList(){ if(!filterChapterListEl) return; filterChapterListEl.innerHTML='<option value="'+esc(BLANK_CHAPTER_FILTER)+'"></option>'+chapterOptionsHtml(); }
      function renderDraftFilterField(key){ var host=filterChipHostForKey(key),input=filterInputForKey(key); if(!host||!input) return; var values=filterValues(draftFilters,key),html=''; for(var i=0;i<values.length;i++) html+='<span class="multi-filter-chip"><span>'+esc(values[i])+'</span><button type="button" data-remove-filter-key="'+key+'" data-remove-filter-index="'+i+'" aria-label="Remove '+esc(values[i])+'">x</button></span>'; host.innerHTML=html; }
      function renderDraftFilters(){ for(var i=0;i<FILTER_KEYS.length;i++) renderDraftFilterField(FILTER_KEYS[i]); syncFilterRegList(); syncFilterChapterList(); }
      function addDraftFilterValue(key, value){ var rawParts=s(value).split(',').map(function(p){ return s(p); }).filter(Boolean); if(!rawParts.length) return; var nextValues=filterValues(draftFilters,key).slice(); for(var i=0;i<rawParts.length;i++){ var normalized=normalizeFilterEntry(key,rawParts[i]); if(normalized) nextValues.push(normalized); } draftFilters[key]=uniqueFilterValues(nextValues); var input=filterInputForKey(key); if(input) input.value=''; renderDraftFilterField(key); if(key==='aircraftType') syncFilterRegList(); }
      function commitPendingDraftInputs(){ for(var i=0;i<FILTER_KEYS.length;i++){ var key=FILTER_KEYS[i],input=filterInputForKey(key); if(input&&s(input.value)) addDraftFilterValue(key,input.value); } }
      function removeDraftFilterValue(key, index){ var values=filterValues(draftFilters,key).slice(); values.splice(index,1); draftFilters[key]=values; renderDraftFilterField(key); if(key==='aircraftType') syncFilterRegList(); }
      function resetDraftFilters(){ draftFilters=emptyFilterState(); for(var i=0;i<FILTER_KEYS.length;i++){ var input=filterInputForKey(FILTER_KEYS[i]); if(input) input.value=''; } renderDraftFilters(); }
      function openFilterPanel(){ if(!filterModal) return; draftFilters=cloneFilterState(activeFilters); for(var i=0;i<FILTER_KEYS.length;i++){ var input=filterInputForKey(FILTER_KEYS[i]); if(input) input.value=''; } renderDraftFilters(); filterModal.className='modal-backdrop filter-backdrop open'; setTimeout(function(){ if(filterAircraftTypeInput) filterAircraftTypeInput.focus(); },0); }
      function closeFilterPanel(){ if(filterModal) filterModal.className='modal-backdrop filter-backdrop'; }
      function readFilterForm(){ commitPendingDraftInputs(); return cloneFilterState(draftFilters); }
      function clearFilters(){ activeFilters=emptyFilterState(); draftFilters=emptyFilterState(); resetDraftFilters(); renderAll(); }
      function clearSearch(){ searchQuery=''; syncSearchUi(); renderAll(); }
      function rowMatchesFilters(row){ var i; if(activeFilters.aircraftType.length){ var typeMatch=false; for(i=0;i<activeFilters.aircraftType.length;i++){ if(normalizedText(aircraftLabel(row)).indexOf(normalizedText(activeFilters.aircraftType[i]))!==-1){ typeMatch=true; break; } } if(!typeMatch) return false; } if(activeFilters.aircraftReg.length){ var regMatch=false; for(i=0;i<activeFilters.aircraftReg.length;i++){ if(normalizedText(s(row['A/C Reg'])).indexOf(normalizedText(activeFilters.aircraftReg[i]))!==-1){ regMatch=true; break; } } if(!regMatch) return false; } if(activeFilters.supervisor.length){ var supervisorName=normalizedText(s(row['Approval Name'])),supervisorFull=normalizedText([s(row['Approval Name']),s(row['Approval stamp']),s(row['Aprroval Licence No.'])].filter(Boolean).join(' | ')),supervisorMatch=false; for(i=0;i<activeFilters.supervisor.length;i++){ var supervisorNeedle=normalizedText(activeFilters.supervisor[i]); if(supervisorName.indexOf(supervisorNeedle)!==-1||supervisorFull.indexOf(supervisorNeedle)!==-1){ supervisorMatch=true; break; } } if(!supervisorMatch) return false; } if(activeFilters.chapter.length){ var chapterMatch=false; for(i=0;i<activeFilters.chapter.length;i++){ var chapterNeedle=activeFilters.chapter[i]; if(chapterNeedle===BLANK_CHAPTER_FILTER){ if(!s(row['Chapter'])){ chapterMatch=true; break; } continue; } chapterNeedle=normalizedText(chapterNeedle); if(normalizedText(chapterLabelText(row)).indexOf(chapterNeedle)!==-1||normalizedText(s(row['Chapter']))===chapterNeedle){ chapterMatch=true; break; } } if(!chapterMatch) return false; } return true; }
      function rowMatchesSearch(row){ var needle=normalizedText(searchQuery); if(!needle) return true; return normalizedText(s(row['Job No'])).indexOf(needle)!==-1||normalizedText(s(row['Task Detail'])).indexOf(needle)!==-1||normalizedText(s(row['Rewriten for cap741'])).indexOf(needle)!==-1; }
      function renderEmptyState(){ if(!hasActiveFilters()&&!hasActiveSearch()) return ''; var title='No pages match your search'; var copy='Try a different Job No or task detail search.'; var button='<button type="button" data-clear-search="1">Clear search</button>'; if(hasActiveFilters()&&hasActiveSearch()){ title='No pages match these filters and search'; copy='Try a broader search, change the filters, or clear everything to show the full logbook again.'; button='<button type="button" data-clear-all-results="1">Clear search and filters</button>'; } else if(hasActiveFilters()){ title='No pages match these filters'; copy='Try a broader mix of aircraft type, registration, supervisor, or chapter, or clear the filters to show the full logbook again.'; button='<button type="button" data-clear-filters="1">Clear filters</button>'; } return '<div class="empty-state" data-transition-key="empty-state"><div class="empty-state-title">'+title+'</div><div class="empty-state-copy">'+copy+'</div>'+button+'</div>'; }

      // ---- Utilities ----
      function setLoadingState(active, title, text){ if(loadingTitleEl&&title!=null) loadingTitleEl.textContent=title; if(loadingTextEl&&text!=null) loadingTextEl.textContent=text; if(loadingOverlay){ loadingOverlay.className=active?'loading-overlay open':'loading-overlay'; loadingOverlay.setAttribute('aria-hidden',active?'false':'true'); } document.body.classList.toggle('busy',!!active); if(loadBtn){ loadBtn.disabled=!!active; } }
      function nextPaint(){ return new Promise(function(resolve){ requestAnimationFrame(function(){ requestAnimationFrame(resolve); }); }); }
      function renderedLayoutElements(){ return Array.prototype.slice.call(pagesEl.querySelectorAll('.page, .empty-state, tr[data-row-key]')); }
      function renderedLayoutKey(el){ return el?(el.getAttribute('data-row-key')||el.getAttribute('data-page-key')||el.getAttribute('data-transition-key')||''):''; }
      function captureRenderedPositions(){ var map=Object.create(null),els=renderedLayoutElements(); for(var i=0;i<els.length;i++){ var key=renderedLayoutKey(els[i]); if(key) map[key]=els[i].getBoundingClientRect(); } return map; }
      function animateRenderedPositions(previous){ if(!previous||!pagesEl||typeof requestAnimationFrame!=='function') return; requestAnimationFrame(function(){ var els=renderedLayoutElements(); for(var i=0;i<els.length;i++){ var el=els[i],key=renderedLayoutKey(el),before=previous[key],after=el.getBoundingClientRect(); if(before){ var deltaY=before.top-after.top; if(Math.abs(deltaY)>1&&typeof el.animate==='function'){ el.animate([{ transform:'translateY('+deltaY+'px)' },{ transform:'translateY(0)' }],{ duration:240, easing:'cubic-bezier(.2,.8,.2,1)' }); } } else if(typeof el.animate==='function'){ el.animate([{ opacity:.35, transform:'translateY(14px)' },{ opacity:1, transform:'translateY(0)' }],{ duration:220, easing:'ease-out' }); } } }); }
      function renderAllWithMotion(){ var previous=captureRenderedPositions(); renderAll(); animateRenderedPositions(previous); }
      function markSharedDatalistsDirty(){ sharedDatalistsCache=''; }
      function normalizedText(value){ return s(value).toLowerCase(); }
      function chapterLabelText(row){ var chapter=s(row['Chapter']),desc=s(row['Chapter Description']); return desc?chapter+' - '+desc:chapter; }
      function fieldAffectsRowLayout(field){ return field==='Task Detail'||field==='Rewriten for cap741'||field==='Approval Name'||field==='Aprroval Licence No.'; }
      function fieldNeedsLiveLayoutRefresh(field){ return field==='Task Detail'||field==='Approval Name'||field==='Aprroval Licence No.'; }
      function captureContentEditableSelection(root){ var sel=window.getSelection&&window.getSelection(); if(!sel||!sel.rangeCount) return {start:0,end:0}; var range=sel.getRangeAt(0); if(!root.contains(range.startContainer)||!root.contains(range.endContainer)) return {start:0,end:0}; var startRange=document.createRange(); startRange.selectNodeContents(root); startRange.setEnd(range.startContainer,range.startOffset); var endRange=document.createRange(); endRange.selectNodeContents(root); endRange.setEnd(range.endContainer,range.endOffset); return {start:startRange.toString().length,end:endRange.toString().length}; }
      function setContentEditableSelection(root, start, end){ var walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT,null); var node,pos=0,startNode=null,endNode=null,startOffset=0,endOffset=0; while((node=walker.nextNode())){ var next=pos+node.nodeValue.length; if(startNode==null&&start<=next){ startNode=node; startOffset=Math.max(0,start-pos); } if(endNode==null&&end<=next){ endNode=node; endOffset=Math.max(0,end-pos); break; } pos=next; } var range=document.createRange(); if(startNode&&endNode){ range.setStart(startNode,startOffset); range.setEnd(endNode,endOffset); } else { range.selectNodeContents(root); range.collapse(false); } var sel=window.getSelection&&window.getSelection(); if(sel){ sel.removeAllRanges(); sel.addRange(range); } }
      function captureEditorSnapshot(target){ var el=target&&target.closest?(target.closest('.editable-cell')||target.closest('input.field-input[data-edit-field], input.field-input[data-new-row]')):null; if(!el) return null; var field=el.getAttribute('data-edit-field'),rowId=el.getAttribute('data-row-id'); if(!field||rowId==null) return null; var snapshot={field:field,rowId:rowId,isInput:el.tagName==='INPUT',scrollX:window.scrollX,scrollY:window.scrollY}; if(snapshot.isInput){ snapshot.start=typeof el.selectionStart==='number'?el.selectionStart:null; snapshot.end=typeof el.selectionEnd==='number'?el.selectionEnd:snapshot.start; } else { var selection=captureContentEditableSelection(el); snapshot.start=selection.start; snapshot.end=selection.end; } return snapshot; }
      function restoreEditorSnapshot(snapshot){ if(!snapshot) return; var selector='[data-row-id="'+snapshot.rowId+'"][data-edit-field="'+snapshot.field+'"]'; var el=pagesEl.querySelector(selector); if(!el||typeof el.focus!=='function') return; try { el.focus({preventScroll:true}); } catch(e){ el.focus(); } if(snapshot.isInput){ if(typeof el.setSelectionRange==='function'&&snapshot.start!=null){ try { el.setSelectionRange(snapshot.start,snapshot.end==null?snapshot.start:snapshot.end); } catch(e){} } } else { setContentEditableSelection(el,snapshot.start||0,snapshot.end||snapshot.start||0); } if(typeof window.scrollTo==='function') window.scrollTo(snapshot.scrollX||0,snapshot.scrollY||0); }
      function refreshLayoutPreservingEditor(){ var snapshot=liveLayoutEditorState; liveLayoutEditorState=null; renderAll(); restoreEditorSnapshot(snapshot); }
      function scheduleLiveLayoutRefresh(snapshot, delay){ liveLayoutEditorState=snapshot; clearTimeout(layoutTimer); layoutTimer=setTimeout(refreshLayoutPreservingEditor,typeof delay==='number'?delay:120); }
      function refreshLayoutIfIdle(){ if(editorIsActive()){ scheduleLayoutRefresh(350); return; } renderAll(); }
      function scheduleLayoutRefresh(delay){ clearTimeout(layoutTimer); layoutTimer=setTimeout(refreshLayoutIfIdle,typeof delay==='number'?delay:300); }
      function scheduleLocalDraftPersist(){ return; }
      function refreshUnsavedChangesState(){ hasUnsavedChanges=settingsDirty||fullLogbookText()!==(lastSavedLogbookText||''); syncSaveButtonState(false); }
      function syncSaveButtonState(isSaving){ if(!saveFileBtn) return; saveFileBtn.classList.toggle('open',!!hasUnsavedChanges||!!isSaving); saveFileBtn.classList.toggle('saving',!!isSaving); saveFileBtn.disabled=!!isSaving; saveFileBtn.setAttribute('aria-label',isSaving?'Saving changes to file':'Save changes to file'); saveFileBtn.title=isSaving?'Saving cap741-data.xlsx...':'Save changes to cap741-data.xlsx'; }
      function setPrintOptionsOpen(open){ if(!printOptionsEl||!printBtn) return; printOptionsEl.classList.toggle('open',!!open); printOptionsEl.setAttribute('aria-hidden',open?'false':'true'); printBtn.classList.toggle('active',!!open); }
      function syncLoadOptionLabels(){ if(loadExistingBtn) loadExistingBtn.textContent=loadButtonMode==='link'?'Link Existing File':'Load Existing File'; }
      function setLoadOptionsOpen(open){ if(!loadOptionsEl||!loadBtn) return; syncLoadOptionLabels(); loadOptionsEl.classList.toggle('open',!!open); loadOptionsEl.setAttribute('aria-hidden',open?'false':'true'); loadBtn.classList.toggle('active',!!open); }
      function fail(msg){ errorBox.style.display='block'; errorBox.textContent=msg; document.body.classList.add('has-top-error'); }
      function clearFail(){ errorBox.style.display='none'; errorBox.textContent=''; document.body.classList.remove('has-top-error'); }
      function saveFailureMessage(error){ var message='Could not save: '+(error&&error.message?error.message:'Unknown error.'); if(message.toLowerCase().indexOf('close it in excel')===-1&&message.toLowerCase().indexOf('open in excel')===-1) message+=' If cap741-data.xlsx is open in Excel, close it and try again.'; return message; }
      function filePickerSupported(){ return typeof window.showOpenFilePicker==='function'; }
      function fileSavePickerSupported(){ return typeof window.showSaveFilePicker==='function'; }
      function setLoadButtonMode(mode){ if(!loadBtn) return; loadButtonMode=mode||'load'; loadBtn.setAttribute('data-mode',loadButtonMode); if(loadButtonMode==='hidden'){ loadBtn.style.display='none'; setLoadOptionsOpen(false); return; } loadBtn.style.display='block'; loadBtn.textContent=loadButtonMode==='link'?'Link':'Load'; loadBtn.title=loadButtonMode==='link'?'Link Excel workbook for saving or create a new CAP741 file':'Load an existing logbook or create a new CAP741 file'; loadBtn.setAttribute('aria-label',loadButtonMode==='link'?'Link Excel workbook for saving or create a new CAP741 file':'Load an existing logbook or create a new CAP741 file'); syncLoadOptionLabels(); }
      function s(v){ return v==null?'':String(v).trim(); }
      function esc(v){ return s(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }

      // ---- Date ----
      function padDatePart(value){ return String(value).padStart(2,'0'); }
      function monthNumberFromName(name){ var months={jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12'}; return months[s(name).slice(0,3).toLowerCase()]||''; }
      function isValidIsoDateParts(year, month, day){ var y=Number(year),m=Number(month),d=Number(day),date=new Date(Date.UTC(y,m-1,d)); return date.getUTCFullYear()===y&&date.getUTCMonth()===(m-1)&&date.getUTCDate()===d; }
      function normalizeDateYear(year){ var y=s(year); if(/^\d{2}$/.test(y)) return String(Number(y)>=70?1900+Number(y):2000+Number(y)); return y; }
      function isoFromDateParts(year, month, day){ var y=normalizeDateYear(year),m=padDatePart(month),d=padDatePart(day); return isValidIsoDateParts(y,m,d)?(y+'-'+m+'-'+d):''; }
      function excelSerialToIso(value){ var serial=Number(value); if(!isFinite(serial)||serial<=0) return ''; var whole=Math.floor(serial); var utc=(whole-25569)*86400000; var date=new Date(utc); if(!isFinite(date.getTime())) return ''; return isoFromDateParts(date.getUTCFullYear(),date.getUTCMonth()+1,date.getUTCDate()); }
      function parseDate(v){ var iso=toIsoInputDate(v); return iso?new Date(iso+'T00:00:00').getTime():8640000000000000; }
      function todayIsoDate(){ var now=new Date(); return now.getFullYear()+'-'+padDatePart(now.getMonth()+1)+'-'+padDatePart(now.getDate()); }
      function toDisplayDate(v){ var m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(s(v)); if(!m) return s(v); var months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']; return m[3]+'/'+months[(+m[2])-1]+'/'+m[1]; }
      function toIsoInputDate(v){ var src=s(v),m,iso=''; if(!src) return ''; if(/^(\d{4})-(\d{2})-(\d{2})$/.test(src)) return src; if(/^\d+(?:\.\d+)?$/.test(src)){ iso=excelSerialToIso(src); if(iso) return iso; } m=/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/.exec(src); if(m) return isoFromDateParts(m[1],m[2],m[3]); m=/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2}|\d{4})$/.exec(src); if(m){ var first=Number(m[1]),second=Number(m[2]),day=first,month=second; if(first<=12&&second>12){ day=second; month=first; } return isoFromDateParts(m[3],month,day); } m=/^(\d{1,2})[\s\/.-]+([A-Za-z]{3,9})[\s\/.-]+(\d{2}|\d{4})$/.exec(src); if(m){ var namedMonth=monthNumberFromName(m[2]); if(namedMonth) return isoFromDateParts(m[3],namedMonth,m[1]); } m=/^([A-Za-z]{3,9})[\s\/.-]+(\d{1,2})(?:,)?[\s\/.-]+(\d{2}|\d{4})$/.exec(src); if(m){ var leadingMonth=monthNumberFromName(m[1]); if(leadingMonth) return isoFromDateParts(m[3],leadingMonth,m[2]); } return ''; }
      function formatDateDisplay(v){ var iso=toIsoInputDate(v); return iso?toDisplayDate(iso):s(v); }
      function parseChapterValue(raw){ var value=s(raw),parts=value.split(' - '); return {chapter:s(parts.shift()),chapterDesc:s(parts.join(' - '))}; }
      function workbookDateValue(row){ return row&&row.__dateDirty?row['Date']:(s(row&&row.__rawDate)||s(row&&row['Date'])); }
      function normalizeLoadedRow(row){ var rawDate=s(row&&row['Date']); row['Date']=formatDateDisplay(rawDate); row.__rawDate=rawDate; row.__dateDirty=false; return row; }

      // ---- Row model ----
      function emptyLogRow(type, chapter, chapterDesc){ return {__rowId:nextRowId(),'Aircraft Type':s(type),'A/C Reg':'','Chapter':s(chapter),'Chapter Description':s(chapterDesc),'Date':'','Job No':'','FAULT':'','Task Detail':'','Rewriten for cap741':'','Approval Name':'','Approval stamp':'','Aprroval Licence No.':''}; }
      function rowHasEntryContent(row){ return !!(s(row['Date'])||s(row['A/C Reg'])||s(row['Job No'])||s(row['FAULT'])||s(row['Task Detail'])||s(row['Rewriten for cap741'])||s(row['Approval Name'])||s(row['Approval stamp'])||s(row['Aprroval Licence No.'])); }
      function nonEmptyRows(list){ var out=[]; for(var i=0;i<(list||[]).length;i++){ if(rowHasEntryContent(list[i]||{})) out.push(list[i]); } return out; }
      function normalizeRows(list){ list=nonEmptyRows(list); rowsById=Object.create(null); var max=-1; for(var i=0;i<list.length;i++){ var id=Number(list[i].__rowId); if(!isFinite(id)||id<0) id=i; list[i].__rowId=id; rowsById[String(id)]=list[i]; if(id>max) max=id; } nextRowIdValue=max+1; return list; }
      function appendRows(list){ for(var i=0;i<list.length;i++){ var row=list[i]; var id=Number(row.__rowId); if(!isFinite(id)||id<0) id=nextRowId(); if(id>=nextRowIdValue) nextRowIdValue=id+1; row.__rowId=id; rowsById[String(id)]=row; rows.push(row); } }
      function nextRowId(){ return nextRowIdValue++; }
      function rowById(id){ return rowsById[String(id)]||null; }
      function removeRowById(id){ var key=String(id); delete rowsById[key]; for(var i=rows.length-1;i>=0;i--){ if(String(rows[i].__rowId)===key){ rows.splice(i,1); break; } } }
      function rowsByGroupKey(key){ var out=[]; for(var i=0;i<rows.length;i++){ var row=rows[i]; if((aircraftLabel(row)+'||'+s(row['Chapter']))===key) out.push(row); } return out; }
      function fillAircraftTypeFromReg(row){ if(!row) return row; var reg=s(row['A/C Reg']).toUpperCase(); if(reg) row['A/C Reg']=reg; if(!s(row['Aircraft Type'])&&reg&&AIRCRAFT_MAP[reg]) row['Aircraft Type']=AIRCRAFT_MAP[reg]; return row; }
      function randomBaAircraftRecord(){
        var matches=[];
        for(var i=0;i<(AIRCRAFT_GROUP_ROWS||[]).length;i++){
          var item=AIRCRAFT_GROUP_ROWS[i]||{};
          if(normalizedText(item.group)!=='ba') continue;
          if(!s(item.reg)||!s(item.type)) continue;
          matches.push({ group:s(item.group), reg:s(item.reg).toUpperCase(), type:s(item.type) });
        }
        if(!matches.length) return null;
        return matches[Math.floor(Math.random()*matches.length)];
      }
      function syncAllRowAircraftTypes(){ for(var i=0;i<rows.length;i++) fillAircraftTypeFromReg(rows[i]); }
      function tsvLineFromRow(row){ return LOG_HEADERS.map(function(key){ var value=key==='Date'?workbookDateValue(row):row[key]; return s(value).replace(/\r?\n/g,' ').replace(/\t/g,' '); }).join('\t'); }
      function fullLogbookText(){ var header=LOG_HEADERS.join('\t'),body=nonEmptyRows(rows).map(tsvLineFromRow).join('\r\n'); return header+'\r\n'+body+(body?'\r\n':''); }

      // ---- Supervisor helpers ----
      function supervisorRecordFor(value){ var key=s(value).toLowerCase(); return key?(SUPERVISOR_LOOKUP[key]||null):null; }
      function normalizeSupervisorValue(value){ var raw=s(value),record=supervisorRecordFor(value); if(record) return {name:record.name,stamp:record.stamp,licence:record.licence}; var parts=raw.split('|').map(function(x){ return s(x); }).filter(Boolean); if(parts.length) return {name:parts[0]||'',stamp:parts[1]||'',licence:parts[2]||''}; return {name:raw,stamp:'',licence:''}; }
      function extractSupervisorParts(value){ var resolved=normalizeSupervisorValue(value); return {name:resolved.name,stamp:resolved.stamp,licence:resolved.licence}; }
      function fillSupervisorFields(nameInput, licenceInput, row){ if(!nameInput) return null; var resolved=normalizeSupervisorValue(nameInput.value); if(resolved.name) nameInput.value=resolved.name; if(licenceInput&&resolved.licence) licenceInput.value=resolved.licence; if(row){ row['Approval Name']=resolved.name; row['Approval stamp']=resolved.stamp; if(resolved.licence) row['Aprroval Licence No.']=resolved.licence; } return resolved; }
      function setRowSupervisorFields(row, nameValue, licenceValue){ var resolved=normalizeSupervisorValue(nameValue); row['Approval Name']=resolved.name; row['Approval stamp']=resolved.stamp; row['Aprroval Licence No.']=licenceValue||resolved.licence||''; }

      // ---- Aircraft / Chapter options HTML ----
      function safeIdPart(value){ return s(value).toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-+|-+$/g,'')||'group'; }
      function aircraftLabel(row){ var reg=s(row['A/C Reg']); return AIRCRAFT_MAP[reg]||s(row['Aircraft Type']); }
      function aircraftOptionsHtml(){ var regs=Object.keys(AIRCRAFT_MAP).sort(),html=''; for(var i=0;i<regs.length;i++) html+='<option value="'+esc(regs[i])+'"></option>'; return html; }
      function aircraftOptionsHtmlForType(type){ var regs=[]; for(var reg in AIRCRAFT_MAP){ if(Object.prototype.hasOwnProperty.call(AIRCRAFT_MAP,reg)&&AIRCRAFT_MAP[reg]===type) regs.push(reg); } regs.sort(); if(!regs.length) return aircraftOptionsHtml(); var html=''; for(var i=0;i<regs.length;i++) html+='<option value="'+esc(regs[i])+'"></option>'; return html; }
      function aircraftTypeOptionsHtml(){ var seen={},vals=[]; for(var k in AIRCRAFT_MAP){ if(Object.prototype.hasOwnProperty.call(AIRCRAFT_MAP,k)&&!seen[AIRCRAFT_MAP[k]]){ seen[AIRCRAFT_MAP[k]]=true; vals.push(AIRCRAFT_MAP[k]); } } vals.sort(); var html=''; for(var i=0;i<vals.length;i++) html+='<option value="'+esc(vals[i])+'"></option>'; return html; }
      function chapterOptionsHtml(){ var html=''; for(var i=0;i<CHAPTER_OPTIONS.length;i++) html+='<option value="'+esc(CHAPTER_OPTIONS[i])+'"></option>'; return html; }
      function supervisorOptionsHtml(){ var html=''; for(var i=0;i<SUPERVISOR_OPTIONS.length;i++) html+='<option value="'+esc(SUPERVISOR_OPTIONS[i])+'"></option>'; return html; }
      function sharedDatalistsHtml(){ if(!sharedDatalistsCache) sharedDatalistsCache='<datalist id="aircraft-reg-list">'+aircraftOptionsHtml()+'</datalist><datalist id="aircraft-type-list">'+aircraftTypeOptionsHtml()+'</datalist><datalist id="chapter-list">'+chapterOptionsHtml()+'</datalist><datalist id="supervisor-list">'+supervisorOptionsHtml()+'</datalist>'; return sharedDatalistsCache; }
      function aircraftRegListIdForGroup(group){ return 'aircraft-reg-list-'+safeIdPart(group.type)+'-'+safeIdPart(group.chapter); }
      function groupAircraftRegDatalistHtml(group){ return '<datalist id="'+aircraftRegListIdForGroup(group)+'">'+aircraftOptionsHtmlForType(group.type)+'</datalist>'; }
      function usedAircraftTypes(){ var seen={},vals=[]; for(var i=0;i<rows.length;i++){ var type=aircraftLabel(rows[i]); if(type&&!seen[type]){ seen[type]=true; vals.push(type); } } vals.sort(); return vals; }
      function usedAircraftTypeOptionsHtml(){ var vals=usedAircraftTypes(); if(!vals.length) return aircraftTypeOptionsHtml(); var html=''; for(var i=0;i<vals.length;i++) html+='<option value="'+esc(vals[i])+'"></option>'; return html; }
      function modalAircraftTypeListId(){ return 'modal-aircraft-type-list'; }
      function modalAircraftRegListId(){ return 'modal-aircraft-reg-list'; }
      function modalAircraftTypeDatalistHtml(){ return '<datalist id="'+modalAircraftTypeListId()+'">'+aircraftTypeOptionsHtml()+'</datalist>'; }
      function modalAircraftRegDatalistHtml(type){ return '<datalist id="'+modalAircraftRegListId()+'">'+aircraftOptionsHtmlForType(type)+'</datalist>'; }

      // ---- Render helpers ----
      function linesFor(text, width){ var t=s(text); if(!t) return 1; var parts=t.split(/\r?\n/); var n=0; for(var i=0;i<parts.length;i++) n+=Math.max(1,Math.ceil(parts[i].length/Math.max(1,width))); return n; }
      function mainPageTaskText(row){ return s(row['Rewriten for cap741']||row['Task Detail']); }
      function unitsFor(row){ var task=linesFor(mainPageTaskText(row),ROW_TASK_CHARS),base=Math.max(task,2); return Math.max(1,Math.min(PAGE_SLOTS,Math.ceil(base/ROW_LINES_PER_SLOT))); }
      function dotsInputSize(value){ return Math.max(8,Math.min(56,s(value).length+1)); }
      function renderDotsInput(value, extraAttrs){ return '<span class="dots-value"><input class="field-input dots-input" type="text" size="'+dotsInputSize(value)+'" value="'+esc(value||'')+'"'+(extraAttrs||'')+'></span>'; }
      function syncDotsInputSize(input){ if(!input||!input.classList||!input.classList.contains('dots-input')) return; input.size=dotsInputSize(valueOf(input)); }
      function editableTextInput(field, rowId, value, placeholder, extraClass, listId){ return '<input class="field-input '+(extraClass||'')+'" type="text"'+(listId?' list="'+listId+'"':'')+' data-row-id="'+rowId+'" data-edit-field="'+field+'" value="'+esc(value||'')+'"'+(placeholder?' placeholder="'+esc(placeholder)+'"':'')+'>';}
      function editableCell(field, rowId, value, cls){ return '<div class="editable-cell '+(cls||'')+'" contenteditable="true" data-row-id="'+rowId+'" data-edit-field="'+field+'">'+(esc(value)||'&nbsp;')+'</div>'; }
      function clearSupervisorButtonHtml(){ return '<button class="sup-clear" type="button" data-clear-supervisor="1" aria-label="Clear supervisor details">Clear</button>'; }
      function editableSupervisorCell(row){ return '<div class="sup"><span class="star">*</span>'+editableTextInput('Approval Name',row.__rowId,row['Approval Name'],'Supervisor','name','supervisor-list')+editableTextInput('Aprroval Licence No.',row.__rowId,row['Aprroval Licence No.'],'Licence number','licence')+clearSupervisorButtonHtml()+'</div>'; }
      function taskCellHtml(row){ return '<div class="task-wrap"><div class="task">'+editableCell('Rewriten for cap741',row.__rowId,mainPageTaskText(row),'task-input')+'</div><button class="task-expand" type="button" data-open-task="1" data-row-id="'+row.__rowId+'" aria-label="Show full task detail">&#x2197;</button></div>'; }
      function blankTaskCellHtml(type, chapter, chapterDesc, regListId){ return '<div class="task-wrap"><div class="task">'+blankEditableCell('Rewriten for cap741',type,chapter,chapterDesc,regListId)+'</div><button class="task-expand" type="button" data-open-task-new="1" aria-label="Show full task detail">&#x2197;</button></div>'; }
      function dateControlHtml(extraAttrs, placeholder, displayValue, isoValue){ return '<div class="date-entry"><input class="field-input date-text" type="text" data-date-text="1" placeholder="'+(placeholder||DATE_PLACEHOLDER)+'" value="'+esc(displayValue||'')+('"'+extraAttrs)+'><input class="date-native" type="date" data-date-picker="1" value="'+esc(isoValue||'')+'"></div>'; }
      // Rows are grouped exactly how the printed logbook is grouped: one section
      // per aircraft type + ATA chapter combination.
      function groupRows(list){ var map={}; for(var i=0;i<list.length;i++){ var row=list[i],key=aircraftLabel(row)+'||'+s(row['Chapter']); if(!map[key]) map[key]={type:aircraftLabel(row),chapter:s(row['Chapter']),chapterDesc:s(row['Chapter Description']),rows:[]}; map[key].rows.push(row); } var out=[]; for(var k in map){ if(Object.prototype.hasOwnProperty.call(map,k)) out.push(map[k]); } out.sort(function(a,b){ if(a.type===b.type) return a.chapter.localeCompare(b.chapter,undefined,{numeric:true}); return a.type.localeCompare(b.type); }); return out; }
      // Each task consumes one or more vertical "slots" on a page, so pagination is
      // based on rendered space rather than raw row count.
      function paginate(list){ var sorted=list.slice(); sorted.sort(function(a,b){ var da=parseDate(a['Date']),db=parseDate(b['Date']); if(da!==db) return da-db; return (Number(a.__rowId)||0)-(Number(b.__rowId)||0); }); var out=[],page=[],used=0; for(var i=0;i<sorted.length;i++){ var row=sorted[i],u=unitsFor(row); if(page.length&&used+u>PAGE_SLOTS){ out.push(page); page=[]; used=0; } page.push({row:row,units:u}); used+=u; } if(page.length) out.push(page); return out; }
      function blankEditableCell(field, type, chapter, chapterDesc, regListId){ var common=' data-new-row="1" data-edit-field="'+field+'" data-new-type="'+esc(type)+'" data-new-chapter="'+esc(chapter)+'" data-new-chapter-desc="'+esc(chapterDesc||'')+'"'; if(field==='Date') return dateControlHtml(common,DATE_PLACEHOLDER); if(field==='A/C Reg') return '<input class="field-input" type="text" list="'+esc(regListId||'aircraft-reg-list')+'" placeholder="G-XXXX"'+common+'>'; if(field==='Job No') return '<input class="field-input" type="text" placeholder="Job No"'+common+'>'; if(field==='Task Detail'||field==='Rewriten for cap741') return '<div class="editable-cell task-input" contenteditable="true"'+common+'></div>'; return '<div class="editable-cell" contenteditable="true"'+common+'></div>'; }
      function blankSupervisorCell(type, chapter, chapterDesc){ return '<div class="sup"><span class="star">*</span><input class="field-input name" type="text" list="supervisor-list" placeholder="Supervisor" data-new-row="1" data-edit-field="Approval Name" data-new-type="'+esc(type)+'" data-new-chapter="'+esc(chapter)+'" data-new-chapter-desc="'+esc(chapterDesc||'')+'"><input class="field-input licence" type="text" placeholder="Licence number" data-new-row="1" data-edit-field="Aprroval Licence No." data-new-type="'+esc(type)+'" data-new-chapter="'+esc(chapter)+'" data-new-chapter-desc="'+esc(chapterDesc||'')+'">'+clearSupervisorButtonHtml()+'</div>'; }
      function makeRows(items, group){ var html='',consumed=0,regListId=aircraftRegListIdForGroup(group); for(var i=0;i<items.length;i++){ var item=items[i]; consumed+=item.units; html+='<tr class="slot'+(item.units>1?' merged-slot':'')+'" data-row-key="row-'+item.row.__rowId+'" style="height:calc(var(--slot-h) * '+item.units+')"><td class="c-date">'+dateControlHtml(' data-row-id="'+item.row.__rowId+'" data-edit-field="Date"',DATE_PLACEHOLDER,formatDateDisplay(item.row['Date']),toIsoInputDate(item.row['Date']))+'</td><td class="c-reg">'+editableTextInput('A/C Reg',item.row.__rowId,item.row['A/C Reg'],'G-XXXX','',regListId)+'</td><td class="c-job">'+editableTextInput('Job No',item.row.__rowId,item.row['Job No'],'Job No')+'</td><td class="c-task">'+taskCellHtml(item.row)+'</td><td class="c-sup">'+editableSupervisorCell(item.row)+'</td></tr>'; } for(var j=consumed;j<PAGE_SLOTS;j++) html+='<tr class="slot"><td class="c-date">'+blankEditableCell('Date',group.type,group.chapter,group.chapterDesc,regListId)+'</td><td class="c-reg">'+blankEditableCell('A/C Reg',group.type,group.chapter,group.chapterDesc,regListId)+'</td><td class="c-job">'+blankEditableCell('Job No',group.type,group.chapter,group.chapterDesc,regListId)+'</td><td class="c-task">'+blankTaskCellHtml(group.type,group.chapter,group.chapterDesc,regListId)+'</td><td class="c-sup">'+blankSupervisorCell(group.type,group.chapter,group.chapterDesc)+'</td></tr>'; return html; }
      function renderDeclaration(){ return '<tfoot><tr class="declaration-row"><td colspan="5"><div class="declaration"><div class="declaration-star">*</div><div class="declaration-text">The above work has been carried out correctly by the logbook owner under my supervision and in accordance with the<br>appropriate technical documentation.</div></div></td></tr></tfoot>'; }
      function renderPage(type, chapter, rowsHtml, owner, sign){ return '<section class="page"><div class="headrow"><div>CAP 741</div><div>Aircraft Maintenance Engineer\'s Logbook</div></div><div class="topline"></div><div class="title">Section 3.1&nbsp;&nbsp; Maintenance Experience</div><div class="dots-row"><div class="field-stack"><div class="dots-field"><span class="dots-label">Aircraft Type:</span><span class="dots-line">'+type+'</span></div><div class="subnote">(Aircraft/Engine combination)</div></div><div class="field-stack top-pad"><div class="dots-field"><span class="dots-label">ATA Chapter:</span><span class="dots-line">'+chapter+'</span></div></div></div><div class="frame"><table class="sheet"><thead><tr><th class="c-date">Date</th><th class="c-reg">A/C Reg</th><th class="c-job">Job No</th><th class="c-task">Task Detail</th><th class="c-sup">Supervisor&rsquo;s Name Signature,<br>and Licence Number</th></tr></thead><tbody>'+rowsHtml+'</tbody>'+renderDeclaration()+'</table></div><div class="owner-row"><div class="dots-field"><span class="dots-label">Logbook Owner\'s Name:</span><span class="dots-line"><span>'+owner+'</span></span></div><div class="dots-field"><span class="dots-label">Signature:</span><span class="dots-line"><span>'+sign+'</span></span></div></div><div style="margin-top:18px" class="bottomline"></div><div class="footer-id">Section 3.1</div></section>'; }
      function renderEditablePage(group, rowsHtml, owner, sign, pageKey){ return '<section class="page" data-group-key="'+esc(group.type+'||'+group.chapter)+'" data-page-key="'+esc(pageKey||'')+'">'+groupAircraftRegDatalistHtml(group)+'<div class="headrow"><div>CAP 741</div><div>Aircraft Maintenance Engineer\'s Logbook</div></div><div class="topline"></div><div class="title">Section 3.1&nbsp;&nbsp; Maintenance Experience</div><div class="dots-row"><div class="field-stack"><div class="dots-field"><span class="dots-label">Aircraft Type:</span><span class="dots-line editable-dots-line">'+renderDotsInput(group.type,' data-group-field="Aircraft Type" list="aircraft-type-list"')+'</span></div><div class="subnote">(Aircraft/Engine combination)</div></div><div class="field-stack top-pad"><div class="dots-field"><span class="dots-label">ATA Chapter:</span><span class="dots-line editable-dots-line">'+renderDotsInput(group.chapter+(group.chapterDesc?' - '+group.chapterDesc:''),' data-group-field="Chapter" list="chapter-list"')+'</span></div></div></div><div class="frame"><table class="sheet"><thead><tr><th class="c-date">Date</th><th class="c-reg">A/C Reg</th><th class="c-job">Job No</th><th class="c-task">Task Detail</th><th class="c-sup">Supervisor&rsquo;s Name, Signature<br>and Licence Number</th></tr></thead><tbody>'+rowsHtml+'</tbody>'+renderDeclaration()+'</table></div><div class="owner-row"><div class="dots-field"><span class="dots-label">Logbook Owner\'s Name:</span><span class="dots-line"><span>'+owner+'</span></span></div><div class="dots-field"><span class="dots-label">Signature:</span><span class="dots-line"><span>'+sign+'</span></span></div></div><div style="margin-top:18px" class="bottomline"></div><div class="footer-id">Section 3.1</div></section>'; }
      function renderDataPage(group, items, pageKey){ return renderEditablePage(group,makeRows(items,group),esc(LOG_OWNER_INFO.name),esc(LOG_OWNER_INFO.signature),pageKey); }
      // Main UI render pass. This rebuilds the visible pages from the current in-memory
      // state instead of diffing small DOM fragments, which keeps layout logic simpler.
      function renderAll(){ try { if(sharedListsEl) sharedListsEl.innerHTML=sharedDatalistsHtml(); syncFilterButtonState(); syncSearchUi(); renderFilterStrip(); var activeRows=nonEmptyRows(rows),visibleRows=activeRows.filter(function(row){ return rowMatchesSearch(row)&&(!hasActiveFilters()||rowMatchesFilters(row)); }); if(!visibleRows.length){ pagesEl.innerHTML=renderEmptyState(); return; } var grps=groupRows(visibleRows),html=[]; for(var i=0;i<grps.length;i++){ var pages=paginate(grps[i].rows); for(var j=0;j<pages.length;j++){ var pageKey=(grps[i].type+'||'+grps[i].chapter+'||'+pages[j].map(function(item){ return item.row.__rowId; }).join('-')); html.push(renderDataPage(grps[i],pages[j],pageKey)); } } pagesEl.innerHTML=html.join(''); wireDateControls(pagesEl); } catch(e){ fail('Could not render pages: '+e.message); } }
      async function renderAllWithLoading(title, text){ setLoadingState(true,title||'Loading logbook',text||'Rendering logbook pages...'); await nextPaint(); renderAll(); }

      // ---- Modal editor ----
      function renderModalEditor(){
        var rowsHtml='';
        for(var i=0;i<PAGE_SLOTS;i++) rowsHtml+='<tr class="slot"><td class="c-date">'+dateControlHtml(' data-field="date"',DATE_PLACEHOLDER)+'</td><td class="c-reg"><input class="field-input" type="text" data-field="reg" list="'+modalAircraftRegListId()+'" placeholder="G-XXXX"></td><td class="c-job"><div class="editable" contenteditable="true" data-field="job"></div></td><td class="c-task"><div class="editable" contenteditable="true" data-field="task"></div></td><td class="c-sup"><div class="sup"><span class="star">*</span><input class="field-input name" type="text" data-field="sup_name" list="supervisor-list" placeholder="Supervisor"><input class="field-input licence" type="text" data-field="sup_licence" placeholder="Licence number"></div></td></tr>';
        return modalAircraftTypeDatalistHtml()+modalAircraftRegDatalistHtml('')+
          '<div class="modal-preview-wrap"><div class="modal-preview">'+
          renderPage(
            renderDotsInput('',' data-head="type" list="'+modalAircraftTypeListId()+'"'),
            renderDotsInput('',' data-head="chapter" list="chapter-list"'),
            rowsHtml,
            esc(LOG_OWNER_INFO.name),
            esc(LOG_OWNER_INFO.signature)
          )+'</div></div>';
      }
      function fitModalPreview(){ var preview=modalBody.querySelector('.modal-preview'),page=modalBody.querySelector('.page'); if(!preview||!page||!modalBody) return; page.style.transform='none'; page.style.transformOrigin='top left'; var availableWidth=Math.max(0,modalBody.clientWidth-16),availableHeight=Math.max(0,modalBody.clientHeight-16),pageWidth=page.offsetWidth||1123,pageHeight=page.offsetHeight||794,scale=Math.min(availableWidth/pageWidth,availableHeight/pageHeight,1); if(!isFinite(scale)||scale<=0) scale=0.5; preview.style.width=Math.round(pageWidth*scale)+'px'; preview.style.height=Math.round(pageHeight*scale)+'px'; page.style.transform='scale('+scale+')'; page.style.margin='0'; if(modalActions&&modalShell){ var previewRect=preview.getBoundingClientRect(),shellRect=modalShell.getBoundingClientRect(),contentRightInset=Math.round(3*scale),contentTopInset=Math.round(0*scale); modalActions.style.width='auto'; modalActions.style.top=Math.round(previewRect.top-shellRect.top+contentTopInset)+'px'; modalActions.style.left=Math.round(previewRect.right-shellRect.left-modalActions.offsetWidth-contentRightInset)+'px'; } }
      function textOf(node){ return node&&typeof node.innerText==='string'?node.innerText.trim():''; }
      function valueOf(node){ if(!node) return ''; if(typeof node.value==='string') return node.value.trim(); return textOf(node); }
      function focusInputAtEnd(input){ if(!input||typeof input.focus!=='function') return; input.focus(); if(typeof input.setSelectionRange==='function'){ var end=input.value.length; try { input.setSelectionRange(end,end); } catch(e){} } }

      // ---- Date control wiring ----
      function syncDateControl(entry, isoValue){ if(!entry) return; var textInput=entry.querySelector('[data-date-text]'),picker=entry.querySelector('[data-date-picker]'); if(textInput) textInput.value=toDisplayDate(isoValue||''); if(picker&&isoValue!=null) picker.value=isoValue; }
      function wireDateControls(scope){ var entries=scope.querySelectorAll('.date-entry'); for(var i=0;i<entries.length;i++){ (function(entry){ var textInput=entry.querySelector('[data-date-text]'),picker=entry.querySelector('[data-date-picker]'); if(!textInput||!picker) return; var openPicker=function(){ textInput.focus(); if(typeof picker.showPicker==='function'){ try { picker.showPicker(); } catch(e){} } else { picker.click(); } }; entry.addEventListener('click',function(ev){ if(ev.target!==picker) openPicker(); }); picker.addEventListener('change',function(){ syncDateControl(entry,picker.value); }); })(entries[i]); } }
      function openDatePicker(entry){ if(!entry) return; var textInput=entry.querySelector('[data-date-text]'),picker=entry.querySelector('[data-date-picker]'); if(!textInput||!picker) return; textInput.focus(); if(typeof picker.showPicker==='function'){ try { picker.showPicker(); } catch(e){ picker.click(); } } else { picker.click(); } }
      function autoSizeDetailTextarea(textarea){ if(!textarea) return; textarea.style.height='auto'; textarea.style.height=Math.max(textarea.scrollHeight,0)+'px'; }
      function autoSizeDetailTextareas(){ autoSizeDetailTextarea(detailFaultEl); autoSizeDetailTextarea(detailTaskEl); autoSizeDetailTextarea(detailRewriteEl); }

      // ---- Modal field wiring ----
      function applySupervisorSuggestion(nameInput, licenceInput){ return fillSupervisorFields(nameInput,licenceInput,null); }
      function syncModalAircraftRegList(type){ var list=modalBody.querySelector('#'+modalAircraftRegListId()); if(list) list.innerHTML=aircraftOptionsHtmlForType(type||''); }
      function wireModalFields(){
        wireDateControls(modalBody);
        var typeHead=modalBody.querySelector('[data-head="type"]');
        syncDotsInputSize(typeHead);
        syncDotsInputSize(modalBody.querySelector('[data-head="chapter"]'));
        syncModalAircraftRegList(typeHead?valueOf(typeHead):'');
        if(typeHead){ typeHead.addEventListener('input',function(){ syncDotsInputSize(this); syncModalAircraftRegList(valueOf(this)); }); typeHead.addEventListener('change',function(){ syncModalAircraftRegList(valueOf(this)); }); }
        var chapterHead=modalBody.querySelector('[data-head="chapter"]');
        if(chapterHead) chapterHead.addEventListener('input',function(){ syncDotsInputSize(this); });
        var headerInputs=modalBody.querySelectorAll('[data-head]');
        for(var h=0;h<headerInputs.length;h++){
          (function(input){
            var line=input.closest('.dots-line');
            if(!line) return;
            line.addEventListener('click',function(ev){
              if(ev.target!==input) focusInputAtEnd(input);
            });
          })(headerInputs[h]);
        }
        var regFields=modalBody.querySelectorAll('[data-field="reg"]');
        for(var i=0;i<regFields.length;i++){ regFields[i].addEventListener('input',function(){ var th=modalBody.querySelector('[data-head="type"]'); if(th){ th.value=AIRCRAFT_MAP[s(this.value).toUpperCase()]||''; syncDotsInputSize(th); syncModalAircraftRegList(valueOf(th)); } }); }
        var clickToEditFields=['job','task'];
        for(var j=0;j<clickToEditFields.length;j++){ var nodes=modalBody.querySelectorAll('[data-field="'+clickToEditFields[j]+'"]'); for(var m=0;m<nodes.length;m++){ nodes[m].closest('td').addEventListener('click',function(){ var editable=this.querySelector('.editable'); if(!editable) return; editable.focus(); var sel=window.getSelection&&window.getSelection(); if(sel&&document.createRange){ var range=document.createRange(); range.selectNodeContents(editable); range.collapse(false); sel.removeAllRanges(); sel.addRange(range); } }); } }
        var supFields=modalBody.querySelectorAll('[data-field="sup_name"]');
        for(var n=0;n<supFields.length;n++){ supFields[n].closest('td').addEventListener('click',function(){ var input=this.querySelector('[data-field="sup_name"]'); if(!input) return; input.focus(); }); supFields[n].addEventListener('change',function(){ var licenceInput=this.closest('td').querySelector('[data-field="sup_licence"]'); applySupervisorSuggestion(this,licenceInput); }); }
      }
      function collectModalPage(){ var heads={},headNodes=modalBody.querySelectorAll('[data-head]'); for(var i=0;i<headNodes.length;i++) heads[headNodes[i].getAttribute('data-head')]=valueOf(headNodes[i]); var resultRows=[],trs=modalBody.querySelectorAll('tbody tr'); for(var j=0;j<trs.length;j++){ var tr=trs[j],rawSupName=valueOf(tr.querySelector('[data-field="sup_name"]')),parsedSup=extractSupervisorParts(rawSupName); var item={date:formatDateDisplay(valueOf(tr.querySelector('[data-field="date"]'))),reg:s(valueOf(tr.querySelector('[data-field="reg"]'))).toUpperCase(),job:textOf(tr.querySelector('[data-field="job"]')),task:textOf(tr.querySelector('[data-field="task"]')),supName:parsedSup.name||rawSupName,supLicence:valueOf(tr.querySelector('[data-field="sup_licence"]'))||parsedSup.licence,supStamp:parsedSup.stamp||''}; if(item.date||item.reg||item.job||item.task||item.supName||item.supLicence) resultRows.push(item); } return {heads:heads,rows:resultRows}; }
      function manualPageToLogRows(page){ var heads=page.heads||{},chapterInfo=parseChapterValue(heads.chapter),out=[]; for(var i=0;i<(page.rows||[]).length;i++){ var item=page.rows[i]; out.push({'Aircraft Type':s(heads.type)||AIRCRAFT_MAP[s(item.reg).toUpperCase()]||'','A/C Reg':s(item.reg).toUpperCase(),'Chapter':chapterInfo.chapter,'Chapter Description':chapterInfo.chapterDesc,'Date':formatDateDisplay(item.date),'Job No':s(item.job),'FAULT':'','Task Detail':s(item.task),'Rewriten for cap741':s(item.task),'Approval Name':s(item.supName),'Approval stamp':s(item.supStamp)||(supervisorRecordFor(item.supName)||{}).stamp||'','Aprroval Licence No.':s(item.supLicence)||(supervisorRecordFor(item.supName)||{}).licence||''}); } return out; }

      // ---- Print ----
      function pageElements(){ return Array.prototype.slice.call(pagesEl.querySelectorAll('.page')); }
      function currentVisiblePage(){ var pages=pageElements(); if(!pages.length) return null; var viewportMid=window.innerHeight/2,best=pages[0],bestDistance=Infinity; for(var i=0;i<pages.length;i++){ var rect=pages[i].getBoundingClientRect(),center=rect.top+(rect.height/2),distance=Math.abs(center-viewportMid); if(rect.top<=viewportMid&&rect.bottom>=viewportMid) return pages[i]; if(distance<bestDistance){ bestDistance=distance; best=pages[i]; } } return best; }
      function clearPrintSelection(){ var pages=pageElements(); document.body.classList.remove('print-current'); for(var i=0;i<pages.length;i++) pages[i].classList.remove('print-exclude'); printMode=''; }
      function printCurrentPage(){ var current=currentVisiblePage(),pages=pageElements(); clearPrintSelection(); if(!current||!pages.length){ window.print(); return; } document.body.classList.add('print-current'); for(var i=0;i<pages.length;i++){ if(pages[i]!==current) pages[i].classList.add('print-exclude'); } printMode='current'; window.print(); }
      function printAllPages(){ clearPrintSelection(); printMode='all'; window.print(); }

      // ---- Editor active state ----
      function editorIsActive(){ var active=document.activeElement; if(!active) return false; return !!(pagesEl.contains(active)&&(active.matches('input, textarea, [contenteditable="true"]')||active.classList.contains('editable-cell'))); }
      function captureActiveEditorState(){ var cell=activeEditorCell(); if(cell) updateRowFromEditor(cell); }
      function activeEditorCell(){ var active=document.activeElement; if(!active||!active.closest) return null; return active.closest('.editable-cell')||active.closest('[data-row-id]')||active.closest('[data-new-row]')||null; }

      // ---- Row editing ----
      function createRowFromBlankCell(cell){ var tr=cell.closest('tr'),first=tr.querySelector('[data-new-row="1"]'); if(!first) return null; var row=emptyLogRow(first.getAttribute('data-new-type')||'',first.getAttribute('data-new-chapter')||'',first.getAttribute('data-new-chapter-desc')||''); rows.push(row); rowsById[String(row.__rowId)]=row; var nodes=tr.querySelectorAll('[data-new-row="1"]'); for(var i=0;i<nodes.length;i++){ nodes[i].setAttribute('data-row-id',row.__rowId); nodes[i].removeAttribute('data-new-row'); nodes[i].removeAttribute('data-new-type'); nodes[i].removeAttribute('data-new-chapter'); nodes[i].removeAttribute('data-new-chapter-desc'); } return row; }
      // Sync a single edited control back into the canonical row object.
      function updateRowFromEditor(cell){ if(!cell) return null; var row=rowById(cell.getAttribute('data-row-id')); if(!row&&cell.hasAttribute('data-new-row')) row=createRowFromBlankCell(cell); if(!row) return null; var field=cell.getAttribute('data-edit-field'),value=(cell.tagName==='INPUT')?valueOf(cell):textOf(cell); if(field==='Date'){ var entry=cell.closest('.date-entry'),picker=entry&&entry.querySelector('[data-date-picker]'),rawDate=(picker&&picker.value)||value,iso=toIsoInputDate(rawDate),originalDisplay=formatDateDisplay(s(row.__rawDate)); value=iso?toDisplayDate(iso):s(rawDate); syncDateControl(entry,iso); row.__dateDirty=!!value ? value!==originalDisplay : !!s(row.__rawDate); } row[field]=value; if(field==='Task Detail') row['Rewriten for cap741']=value; if(field==='A/C Reg'){ row['A/C Reg']=value.toUpperCase(); if(cell.value!==row['A/C Reg']) cell.value=row['A/C Reg']; var mapped=AIRCRAFT_MAP[row['A/C Reg']]; if(mapped) row['Aircraft Type']=mapped; } if(field==='Approval Name'){ var licenceInput=cell.closest('td')&&cell.closest('td').querySelector('[data-edit-field="Aprroval Licence No."]'); var resolvedSup=fillSupervisorFields(cell,licenceInput,row); if(resolvedSup){ row['Approval Name']=resolvedSup.name; row['Approval stamp']=resolvedSup.stamp; row['Aprroval Licence No.']=licenceInput?s(licenceInput.value):(resolvedSup.licence||''); } } if(field==='Aprroval Licence No.'){ var nameInput=cell.closest('td')&&cell.closest('td').querySelector('[data-edit-field="Approval Name"]'); setRowSupervisorFields(row,nameInput?nameInput.value:row['Approval Name'],value); } if(!value&&cell.classList&&cell.classList.contains('editable-cell')) cell.innerHTML='&nbsp;'; refreshUnsavedChangesState(); return row; }
      async function clearSupervisorFields(button){
        var tr=button&&button.closest?button.closest('tr'):null;
        if(!tr) return;
        if(!await showConfirmDialog('Clear Row','This will remove all content in this row from the logbook page.','Clear row')) return;
        var rowId=button.getAttribute('data-row-id')||'';
        if(!rowId){
          var rowBoundField=tr.querySelector('[data-row-id]');
          rowId=rowBoundField?rowBoundField.getAttribute('data-row-id'):'';
        }
        if(rowId&&rowById(rowId)){
          removeRowById(rowId);
          refreshUnsavedChangesState();
          renderAllWithMotion();
          scheduleAutoSave();
          return;
        }
        var textInputs=tr.querySelectorAll('input.field-input[data-new-row]');
        for(var i=0;i<textInputs.length;i++) textInputs[i].value='';
        var editableCells=tr.querySelectorAll('.editable-cell[data-new-row]');
        for(var j=0;j<editableCells.length;j++) editableCells[j].innerHTML='&nbsp;';
        var datePickers=tr.querySelectorAll('.date-native');
        for(var k=0;k<datePickers.length;k++) datePickers[k].value='';
      }
      function openTaskDetailFromButton(button){
        var rowId=button&&button.getAttribute&&button.getAttribute('data-row-id');
        if(rowId){ openTaskDetail(rowId); return; }
        var tr=button&&button.closest?button.closest('tr'):null;
        var taskCell=tr&&tr.querySelector('.editable-cell.task-input[data-edit-field="Rewriten for cap741"]');
        var row=taskCell&&updateRowFromEditor(taskCell);
        if(!row||!tr) return;
        var editors=tr.querySelectorAll('[data-row-id="'+row.__rowId+'"][data-edit-field]');
        for(var i=0;i<editors.length;i++) updateRowFromEditor(editors[i]);
        openTaskDetail(row.__rowId);
      }
      function syncBlankRowMetadata(page, type, chapter, chapterDesc){ var blanks=page.querySelectorAll('[data-new-row="1"]'); for(var i=0;i<blanks.length;i++){ blanks[i].setAttribute('data-new-type',type); blanks[i].setAttribute('data-new-chapter',chapter); blanks[i].setAttribute('data-new-chapter-desc',chapterDesc); } page.setAttribute('data-group-key',type+'||'+chapter); }
      function saveModalPageRows(rowsToAppend){ if(!rowsToAppend.length) return; appendRows(rowsToAppend); refreshUnsavedChangesState(); }

      // ---- Task detail modal ----
      function openInfoModal(){ if(infoModal) infoModal.className='modal-backdrop open'; }
      function closeInfoModal(){ if(infoModal) infoModal.className='modal-backdrop'; }
      function taskDetailStateFromRow(row){ return { chapter:s(row['Chapter']), chapterDesc:s(row['Chapter Description']), fault:s(row['FAULT']), task:s(row['Task Detail']), rewrite:s(row['Rewriten for cap741']) }; }
      function restoreTaskDetailState(row, state){ if(!row||!state) return; row['Chapter']=state.chapter; row['Chapter Description']=state.chapterDesc; row['FAULT']=state.fault; row['Task Detail']=state.task; row['Rewriten for cap741']=state.rewrite; }
      function previewTaskDetailForm(){ var row=rowById(lastTaskDetailRowId); if(!row) return; applyTaskDetailForm(row,readTaskDetailForm()); renderAll(); }
      function openTaskDetail(rowId){ lastTaskDetailFocus=document.activeElement&&pagesEl.contains(document.activeElement)?document.activeElement:null; captureActiveEditorState(); var row=rowById(rowId); if(!row) return; lastTaskDetailRowId=rowId; taskDetailOriginalState=taskDetailStateFromRow(row); taskDetailRewriteDirty=false; detailChapterEl.value=chapterLabelText(row); detailFaultEl.value=s(row['FAULT']); detailTaskEl.value=s(row['Task Detail']); detailRewriteEl.value=s(row['Rewriten for cap741']||row['Task Detail']); taskDetailModal.className='modal-backdrop open'; if(typeof requestAnimationFrame==='function') requestAnimationFrame(autoSizeDetailTextareas); else autoSizeDetailTextareas(); }
      function showConfirmDialog(title, text, okLabel){ return new Promise(function(resolve){ confirmResolver=resolve; if(confirmTitleEl) confirmTitleEl.textContent=title||'Confirm'; if(confirmTextEl) confirmTextEl.textContent=text||'Are you sure?'; if(confirmOkBtn) confirmOkBtn.textContent=okLabel||'Confirm'; if(confirmModal) confirmModal.className='modal-backdrop open'; }); }
      function closeConfirmDialog(result){ if(confirmModal) confirmModal.className='modal-backdrop'; if(confirmResolver){ var resolve=confirmResolver; confirmResolver=null; resolve(!!result); } }
      function readTaskDetailForm(){
        return {
          chapter:s(detailChapterEl.value),
          fault:s(detailFaultEl.value),
          task:s(detailTaskEl.value),
          rewrite:s(detailRewriteEl.value)
        };
      }
      function applyTaskDetailForm(row, form){
        var parsedChapter=parseChapterValue(form.chapter);
        row['FAULT']=form.fault;
        row['Task Detail']=form.task;
        row['Rewriten for cap741']=form.rewrite||form.task;
        if(parsedChapter.chapter){
          row['Chapter']=parsedChapter.chapter;
          row['Chapter Description']=parsedChapter.chapterDesc;
        }
      }
      function saveTaskDetail(){
        var row=rowById(lastTaskDetailRowId);
        if(!row) return;
        applyTaskDetailForm(row,readTaskDetailForm());
        refreshUnsavedChangesState();
        renderAll();
        closeTaskDetail(true);
        scheduleAutoSave();
      }
      function closeTaskDetail(keepPreviewChanges){ var row=lastTaskDetailRowId&&rowById(lastTaskDetailRowId); if(!keepPreviewChanges&&row&&taskDetailOriginalState){ restoreTaskDetailState(row,taskDetailOriginalState); renderAll(); } taskDetailModal.className='modal-backdrop'; lastTaskDetailRowId=null; taskDetailOriginalState=null; taskDetailRewriteDirty=false; if(lastTaskDetailFocus&&document.contains(lastTaskDetailFocus)&&typeof lastTaskDetailFocus.focus==='function'){ try { lastTaskDetailFocus.focus({preventScroll:true}); } catch(e){ try { lastTaskDetailFocus.focus(); } catch(err){} } } lastTaskDetailFocus=null; }

      // ---- IndexedDB ----
      function withHandleDb(mode){ return new Promise(function(resolve,reject){ if(!window.indexedDB){ reject(new Error('IndexedDB unavailable')); return; } var request=indexedDB.open(DB_NAME,1); request.onupgradeneeded=function(){ var db=request.result; if(!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE); }; request.onerror=function(){ reject(request.error||new Error('Could not open file-handle store')); }; request.onsuccess=function(){ var db=request.result,tx=db.transaction(DB_STORE,mode),store=tx.objectStore(DB_STORE); resolve({db:db,tx:tx,store:store}); }; }); }
      async function loadStoredHandle(key){ try { var ctx=await withHandleDb('readonly'); return await new Promise(function(resolve,reject){ var req=ctx.store.get(key); req.onsuccess=function(){ ctx.db.close(); resolve(req.result||null); }; req.onerror=function(){ ctx.db.close(); reject(req.error||new Error('Could not read stored file handle')); }; }); } catch(e){ return null; } }
      async function storeHandle(key, handle){ try { var ctx=await withHandleDb('readwrite'); return await new Promise(function(resolve,reject){ var req=ctx.store.put(handle,key); req.onsuccess=function(){ ctx.db.close(); resolve(true); }; req.onerror=function(){ ctx.db.close(); reject(req.error||new Error('Could not store file handle')); }; }); } catch(e){ return false; } }
      async function ensurePermission(handle){ if(!handle||typeof handle.queryPermission!=='function') return false; var opts={mode:'readwrite'}; if(await handle.queryPermission(opts)==='granted') return true; return await handle.requestPermission(opts)==='granted'; }
      async function pickWorkbookHandle(){
        if(!filePickerSupported()) throw new Error('File picker not supported. Open this page via a local web server (e.g. VS Code Live Server) and use Chrome or Edge.');
        var picked=await window.showOpenFilePicker({multiple:false,types:[{description:'CAP741 Excel workbook',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}],excludeAcceptAllOption:true});
        var handle=picked&&picked[0];
        if(!handle) return null;
        if(!handleIsWorkbook(handle)) throw new Error('Please choose a .xlsx Excel file.');
        if(!await ensurePermission(handle)) return null;
        await storeHandle(LINKED_FILE_KEY,handle);
        return handle;
      }
      async function pickNewWorkbookHandle(){
        if(!fileSavePickerSupported()) throw new Error('Save file picker not supported. Use Chrome or Edge over a local web server to create a new Excel file.');
        var handle=await window.showSaveFilePicker({suggestedName:'cap741-data.xlsx',types:[{description:'CAP741 Excel workbook',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}],excludeAcceptAllOption:true});
        if(!handle) return null;
        if(!handleIsWorkbook(handle)) throw new Error('Please save the new workbook as a .xlsx Excel file.');
        if(!await ensurePermission(handle)) return null;
        await storeHandle(LINKED_FILE_KEY,handle);
        return handle;
      }

      // ---- Reference data state builders (from xlsx) ----
      function parseSupervisorRecordsText(text){ var lines=String(text||'').split(/\r?\n/),out=[]; for(var i=0;i<lines.length;i++){ var line=s(lines[i]); if(!line||/^id\s+/i.test(line)) continue; var cols=line.split('\t').map(function(x){ return s(x); }); while(cols.length&&!s(cols[cols.length-1])) cols.pop(); if(cols.length<4) continue; out.push({id:cols[0]||'',name:cols[1]||'',stamp:cols[2]||'',licence:cols[3]||'',scope:cols[4]||'',date:cols[5]||''}); } return out; }
      function rebuildSupervisorState(records){ var out=[],options=[],lookup=Object.create(null); records=records||[]; for(var i=0;i<records.length;i++){ var record=records[i]||{}; var clean={id:s(record.id||record.ID||''),name:s(record.name||record['Signatory Name']||record['Name']||''),stamp:s(record.stamp||record['Stamp']||''),licence:s(record.licence||record['License Number']||record['Licence Number']||''),scope:s(record.scope||record['Scope / Limitations']||''),date:s(record.date||record['Date']||'')}; if(!clean.name) continue; out.push(clean); var label=clean.name+' | '+clean.stamp+' | '+clean.licence; options.push(label); lookup[clean.name.toLowerCase()]=clean; lookup[label.toLowerCase()]=clean; } SUPERVISOR_RECORDS=out; SUPERVISOR_OPTIONS=options.sort(function(a,b){ return a.localeCompare(b); }); SUPERVISOR_LOOKUP=lookup; }
      function applyAircraftGroupRows(records){ var map=Object.create(null),out=[]; for(var i=0;i<(records||[]).length;i++){ var record=records[i]||{}; var clean={group:s(record.group||record.Group||''),reg:s(record.reg||record['A/C Reg']||'').toUpperCase(),type:s(record.type||record['Aircraft Type']||'')}; if(!clean.reg||!clean.type) continue; out.push(clean); map[clean.reg]=clean.type; } AIRCRAFT_GROUP_ROWS=out; AIRCRAFT_MAP=map; }
      function applyChapterRows(records){ CHAPTER_OPTIONS=[]; for(var i=0;i<(records||[]).length;i++){ var chapter=s(records[i].chapter||records[i].Chapter||''),desc=s(records[i].description||records[i].Description||''); if(chapter) CHAPTER_OPTIONS.push(desc?chapter+' - '+desc:chapter); } }
      function aircraftWorkbookRows(){ return AIRCRAFT_GROUP_ROWS.map(function(item){ return {Group:s(item.group),'A/C Reg':s(item.reg),'Aircraft Type':s(item.type)}; }); }
      function chapterWorkbookRows(){ return CHAPTER_OPTIONS.map(function(label){ var p=parseChapterValue(label); return {Chapter:p.chapter,Description:p.chapterDesc}; }); }
      function supervisorWorkbookRows(){ return SUPERVISOR_RECORDS.map(function(item){ return {ID:s(item.id),'Signatory Name':s(item.name),Stamp:s(item.stamp),'License Number':s(item.licence),'Scope / Limitations':s(item.scope),Date:s(item.date)}; }); }

      // ---- XLSX core ----
      function handleIsWorkbook(handle){ return !!(handle&&/\.xlsx$/i.test(s(handle.name))); }
      function workbookSheetObjects(workbook, sheetName){ var sheet=workbook&&workbook.Sheets?workbook.Sheets[sheetName]:null; if(!sheet||!window.XLSX) return []; return XLSX.utils.sheet_to_json(sheet,{defval:'',raw:false}); }
      // Workbook sheets are the source of truth on disk; this function translates them
      // into the smaller in-memory structures the UI works with.
      function loadWorkbookFromArrayBuffer(buffer){ var workbook=XLSX.read(buffer,{type:'array'}); var logRows=workbookSheetObjects(workbook,'Logbook'); if(logRows.length){ var parsed=[]; for(var i=0;i<logRows.length;i++){ var row={}; for(var j=0;j<LOG_HEADERS.length;j++) row[LOG_HEADERS[j]]=s(logRows[i][LOG_HEADERS[j]]); parsed.push(normalizeLoadedRow(row)); } rows=normalizeRows(parsed); } applyAircraftGroupRows(workbookSheetObjects(workbook,'Aircraft').map(function(r){ return {group:r.Group,reg:r['A/C Reg'],type:r['Aircraft Type']}; })); syncAllRowAircraftTypes(); applyChapterRows(workbookSheetObjects(workbook,'Chapters').map(function(r){ return {chapter:r.Chapter,description:r.Description}; })); rebuildSupervisorState(workbookSheetObjects(workbook,'Supervisors').map(function(r){ return {id:r.ID,name:r['Signatory Name'],stamp:r.Stamp,licence:r['License Number'],scope:r['Scope / Limitations'],date:r.Date}; })); var infoRows=workbookSheetObjects(workbook,'Info'); for(var k=0;k<infoRows.length;k++){ var key=normalizedText(infoRows[k].Key||infoRows[k].key),value=s(infoRows[k].Value||infoRows[k].value); if(key==='name') LOG_OWNER_INFO.name=value; if(key==='signature') LOG_OWNER_INFO.signature=value; if(key==='stamp') LOG_OWNER_INFO.stamp=value; } markSharedDatalistsDirty(); lastSavedLogbookText=fullLogbookText(); settingsDirty=false; }
      async function loadDefaultWorkbookData(){ var res=await fetch(DEFAULT_WORKBOOK_PATH,{cache:'no-store'}); if(!res.ok) throw new Error('Excel workbook returned '+res.status); loadWorkbookFromArrayBuffer(await res.arrayBuffer()); }
      function buildWorkbookFromState(){ syncAllRowAircraftTypes(); var wb=XLSX.utils.book_new(); var logRows=rows.map(function(row){ var out={}; for(var i=0;i<LOG_HEADERS.length;i++) out[LOG_HEADERS[i]]=s(LOG_HEADERS[i]==='Date'?workbookDateValue(row):row[LOG_HEADERS[i]]); return out; }); XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(logRows,{header:LOG_HEADERS}),'Logbook'); XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(aircraftWorkbookRows(),{header:['Group','A/C Reg','Aircraft Type']}),'Aircraft'); XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(chapterWorkbookRows(),{header:['Chapter','Description']}),'Chapters'); XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(supervisorWorkbookRows(),{header:['ID','Signatory Name','Stamp','License Number','Scope / Limitations','Date']}),'Supervisors'); XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet([{Key:'Name',Value:s(LOG_OWNER_INFO.name)},{Key:'Signature',Value:s(LOG_OWNER_INFO.signature)},{Key:'Stamp',Value:s(LOG_OWNER_INFO.stamp)}],{header:['Key','Value']}),'Info'); return wb; }
      async function getXlsxHandle(){ try { var stored=await loadStoredHandle(LINKED_FILE_KEY); if(stored&&handleIsWorkbook(stored)){ var perm=await stored.queryPermission({mode:'readwrite'}); if(perm==='granted') return stored; perm=await stored.requestPermission({mode:'readwrite'}); if(perm==='granted') return stored; } } catch(e){} return null; }
      function isStaleHandleError(e){ return !!(e && (e.name==='InvalidStateError' || (e.message&&e.message.indexOf('state cached')!==-1))); }
      async function writeWorkbookToHandle(handle){
        if(!handle) throw new Error('No Excel file linked. Click Save or Link and choose cap741-data.xlsx once to keep using it.');
        var data=XLSX.write(buildWorkbookFromState(),{bookType:'xlsx',type:'array'});
        var writable=null;
        for(var attempt=0;attempt<2;attempt++){
          try {
            writable=await handle.createWritable({keepExistingData:false});
            break;
          } catch(e){
            if(attempt===0 && isStaleHandleError(e)){
              try { await handle.requestPermission({mode:'readwrite'}); } catch(permErr){}
              continue;
            }
            if(isStaleHandleError(e)){
              throw new Error('cap741-data.xlsx is open in Excel. Close it in Excel, then click Save again.');
            }
            throw e;
          }
        }
        if(!writable) throw new Error('cap741-data.xlsx is open in Excel. Close it in Excel, then click Save again.');
        try {
          await writable.write(data);
          await writable.close();
        } catch(writeErr){
          try { await writable.abort(); } catch(abortErr){}
          throw writeErr;
        }
        lastSavedLogbookText=fullLogbookText();
        settingsDirty=false;
      }
      function initializeNewWorkbookState(){
        var starterAircraft=randomBaAircraftRecord();
        var starterRow=normalizeLoadedRow({
          'Aircraft Type':starterAircraft?starterAircraft.type:'',
          'A/C Reg':starterAircraft?starterAircraft.reg:'',
          'Chapter':'',
          'Chapter Description':'',
          'Date':todayIsoDate(),
          'Job No':'TEST-001',
          'FAULT':'',
          'Task Detail':NEW_WORKBOOK_TASK_TEXT,
          'Rewriten for cap741':NEW_WORKBOOK_TASK_TEXT,
          'Approval Name':NEW_WORKBOOK_SUPERVISOR_NAME,
          'Approval stamp':NEW_WORKBOOK_SUPERVISOR_NAME,
          'Aprroval Licence No.':NEW_WORKBOOK_SUPERVISOR_LICENCE
        });
        rows=normalizeRows([starterRow]);
        AIRCRAFT_GROUP_ROWS=starterAircraft?[starterAircraft]:[];
        AIRCRAFT_MAP=starterAircraft?(function(){ var map=Object.create(null); map[starterAircraft.reg]=starterAircraft.type; return map; })():Object.create(null);
        CHAPTER_OPTIONS=[];
        LOG_OWNER_INFO={ name:NEW_WORKBOOK_OWNER_NAME, signature:NEW_WORKBOOK_OWNER_NAME, stamp:NEW_WORKBOOK_OWNER_NAME };
        activeFilters=emptyFilterState();
        draftFilters=emptyFilterState();
        searchQuery='';
        rebuildSupervisorState([{ id:'1', name:NEW_WORKBOOK_SUPERVISOR_NAME, stamp:NEW_WORKBOOK_SUPERVISOR_NAME, licence:NEW_WORKBOOK_SUPERVISOR_LICENCE, scope:'', date:todaySupervisorDate() }]);
        markSharedDatalistsDirty();
        settingsDirty=false;
        lastSavedLogbookText='';
      }
      async function writeXlsx(allowPicker){
        var handle=await getXlsxHandle();
        if(!handle&&allowPicker){
          setLoadingState(true,'Linking file','Choose the Excel workbook to save to...');
          handle=await pickWorkbookHandle();
          if(handle) setLoadingState(true,'Saving','Writing changes to cap741-data.xlsx...');
        }
        await writeWorkbookToHandle(handle);
      }
      async function createNewWorkbookFile(){
        var handle=await pickNewWorkbookHandle();
        if(!handle) return false;
        initializeNewWorkbookState();
        await writeWorkbookToHandle(handle);
        setLoadButtonMode('hidden');
        await renderAllWithLoading('Creating logbook','Rendering starter CAP741 pages...');
        refreshUnsavedChangesState();
        return true;
      }

      // ---- Auto-save after chapter/data changes ----
      function scheduleAutoSave(){ clearTimeout(autoSaveTimer); autoSaveTimer=setTimeout(async function(){ if(!hasUnsavedChanges) return; try { await writeXlsx(false); refreshUnsavedChangesState(); } catch(e){ /* silently fail - save button still available */ } },1500); }

      // ---- Flush / save ----
      async function flushLinkedRewrite(force){ if(!hasUnsavedChanges&&!force) return; captureActiveEditorState(); if(saveInFlight){ saveQueued=true; return; } saveInFlight=true; syncSaveButtonState(true); try { clearFail(); await writeXlsx(!!force); refreshUnsavedChangesState(); } catch(e){ if(e&&e.name==='AbortError') fail('Save cancelled. Click Save again and choose cap741-data.xlsx to link it.'); else fail(saveFailureMessage(e)); } finally { saveInFlight=false; syncSaveButtonState(false); if(saveQueued){ saveQueued=false; flushLinkedRewrite(true); } } }

      // ---- Settings modal with tabs ----
      function settingsTableRow(cells, kind, rowAttrs){ var html='<tr'+(rowAttrs?' '+rowAttrs:'')+'>'; for(var i=0;i<cells.length;i++) html+='<td>'+cells[i]+'</td>'; html+='<td><button type="button" class="settings-remove-btn" data-settings-remove="'+esc(kind)+'">&#x2715;</button></td></tr>'; return html; }
      function nextSupervisorNumericId(records){ var max=0; for(var i=0;i<(records||[]).length;i++){ var id=parseInt(s(records[i]&&records[i].id),10); if(isFinite(id)&&id>max) max=id; } return max+1; }
      function todaySupervisorDate(){ var now=new Date(); var iso=now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0')+'-'+String(now.getDate()).padStart(2,'0'); return formatDateDisplay(iso); }
      function renderSettingsRows(kind){ var html=''; if(kind==='aircraft'){ var list=aircraftWorkbookRows(); for(var i=0;i<list.length;i++) html+=settingsTableRow(['<input type="text" data-col="Group" value="'+esc(list[i].Group)+'">','<input type="text" data-col="A/C Reg" value="'+esc(list[i]['A/C Reg'])+'">','<input type="text" data-col="Aircraft Type" value="'+esc(list[i]['Aircraft Type'])+'">'],kind); }
      if(kind==='chapters'){ var ch=chapterWorkbookRows(); for(var j=0;j<ch.length;j++) html+=settingsTableRow(['<input type="text" data-col="Chapter" value="'+esc(ch[j].Chapter)+'">','<input type="text" data-col="Description" value="'+esc(ch[j].Description)+'">'],kind); }
      if(kind==='supervisors'){ var su=supervisorWorkbookRows(); for(var k=0;k<su.length;k++) html+=settingsTableRow(['<input type="text" data-col="Signatory Name" value="'+esc(su[k]['Signatory Name'])+'">','<input type="text" data-col="Stamp" value="'+esc(su[k].Stamp)+'">','<input type="text" data-col="License Number" value="'+esc(su[k]['License Number'])+'">','<input type="text" data-col="Scope / Limitations" value="'+esc(su[k]['Scope / Limitations'])+'">'],kind,'data-supervisor-id="'+esc(su[k].ID)+'" data-supervisor-date="'+esc(su[k].Date)+'"'); }
      return html; }
      function collectSettingsTable(kind){ var tbody=settingsBodyEl&&settingsBodyEl.querySelector('[data-settings-table="'+kind+'"] tbody'),out=[]; if(!tbody) return out; var trs=tbody.querySelectorAll('tr'); for(var i=0;i<trs.length;i++){ var obj={},hasValue=false,inputs=trs[i].querySelectorAll('[data-col]'); for(var j=0;j<inputs.length;j++){ var key=inputs[j].getAttribute('data-col'),value=s(inputs[j].value); if(value) hasValue=true; obj[key]=value; } if(hasValue) out.push(obj); } return out; }
      function collectSupervisorSettingsTable(){
        var tbody=settingsBodyEl&&settingsBodyEl.querySelector('[data-settings-table="supervisors"] tbody'),out=[];
        if(!tbody) return out;
        var nextId=nextSupervisorNumericId(SUPERVISOR_RECORDS),trs=tbody.querySelectorAll('tr');
        for(var i=0;i<trs.length;i++){
          var tr=trs[i],obj={},hasValue=false,inputs=tr.querySelectorAll('[data-col]');
          for(var j=0;j<inputs.length;j++){
            var key=inputs[j].getAttribute('data-col'),value=s(inputs[j].value);
            if(value) hasValue=true;
            obj[key]=value;
          }
          if(!hasValue) continue;
          obj.ID=s(tr.getAttribute('data-supervisor-id'))||String(nextId++);
          obj.Date=s(tr.getAttribute('data-supervisor-date'))||todaySupervisorDate();
          out.push(obj);
        }
        return out;
      }
      function renderSettingsBody(tab){
        settingsActiveTab=tab||'owner';
        var TABS=[{id:'owner',label:'Owner'},{id:'aircraft',label:'Aircraft'},{id:'supervisors',label:'Supervisors'},{id:'chapters',label:'Chapters'}];
        var tabsHtml='<div class="settings-tabs-nav">';
        for(var t=0;t<TABS.length;t++) tabsHtml+='<button type="button" class="settings-tab-btn'+(TABS[t].id===settingsActiveTab?' active':'')+'" data-settings-tab="'+TABS[t].id+'">'+TABS[t].label+'</button>';
        tabsHtml+='</div>';
        var panelHtml='';
        if(settingsActiveTab==='owner'){
          panelHtml='<div class="settings-tab-panel"><p class="settings-panel-copy">Used on the CAP 741 page footer.</p><div class="settings-grid"><div class="settings-field"><label>Name</label><input class="settings-input" id="settingsOwnerName" type="text" value="'+esc(LOG_OWNER_INFO.name)+'"></div><div class="settings-field"><label>Stamp</label><input class="settings-input" id="settingsOwnerStamp" type="text" value="'+esc(LOG_OWNER_INFO.stamp)+'"></div></div></div>';
        } else if(settingsActiveTab==='aircraft'){
          panelHtml='<div class="settings-tab-panel"><p class="settings-panel-copy">Manage aircraft registration, type, and group. Registration is used to auto-fill Aircraft Type when entering A/C Reg.</p><div class="settings-panel-toolbar"><button class="settings-add-row" type="button" data-settings-add="aircraft">+ Add aircraft</button></div><div class="settings-table-wrap"><table class="settings-table" data-settings-table="aircraft"><thead><tr><th>Group</th><th>A/C Reg</th><th>Aircraft Type</th><th></th></tr></thead><tbody>'+renderSettingsRows('aircraft')+'</tbody></table></div></div>';
        } else if(settingsActiveTab==='supervisors'){
          panelHtml='<div class="settings-tab-panel"><p class="settings-panel-copy">Manage supervisor names, stamps, and licence numbers. The app keeps ID numbers and save dates automatically in the background.</p><div class="settings-panel-toolbar"><button class="settings-add-row" type="button" data-settings-add="supervisors">+ Add supervisor</button></div><div class="settings-table-wrap"><table class="settings-table" data-settings-table="supervisors"><thead><tr><th>Name</th><th>Stamp</th><th>Licence</th><th>Scope</th><th></th></tr></thead><tbody>'+renderSettingsRows('supervisors')+'</tbody></table></div></div>';
        } else if(settingsActiveTab==='chapters'){
          panelHtml='<div class="settings-tab-panel"><p class="settings-panel-copy">Manage ATA chapter numbers and descriptions. These appear in the Chapter dropdown on pages and filters.</p><div class="settings-panel-toolbar"><button class="settings-add-row" type="button" data-settings-add="chapters">+ Add chapter</button></div><div class="settings-table-wrap"><table class="settings-table" data-settings-table="chapters"><thead><tr><th>Chapter</th><th>Description</th><th></th></tr></thead><tbody>'+renderSettingsRows('chapters')+'</tbody></table></div></div>';
        }
        settingsBodyEl.innerHTML=tabsHtml+panelHtml;
        // Wire tab buttons
        var tabBtns=settingsBodyEl.querySelectorAll('.settings-tab-btn');
        for(var i=0;i<tabBtns.length;i++){ (function(btn){ btn.addEventListener('click',function(){ var nextTab=btn.getAttribute('data-settings-tab'); if(nextTab===settingsActiveTab) return; renderSettingsBody(nextTab); }); })(tabBtns[i]); }
      }
      function openSettingsModal(){ if(!settingsModal||!settingsBodyEl) return; renderSettingsBody(settingsActiveTab); settingsModal.className='modal-backdrop open'; }
      function closeSettingsModal(){ if(settingsModal) settingsModal.className='modal-backdrop'; }
      function addSettingsRow(kind){ var tbody=settingsBodyEl&&settingsBodyEl.querySelector('[data-settings-table="'+kind+'"] tbody'); if(!tbody) return; var rowHtml=kind==='aircraft'?settingsTableRow(['<input type="text" data-col="Group" placeholder="Group">','<input type="text" data-col="A/C Reg" placeholder="G-XXXX">','<input type="text" data-col="Aircraft Type" placeholder="Boeing 777-300ER - GE90">'],kind):(kind==='chapters'?settingsTableRow(['<input type="text" data-col="Chapter" placeholder="e.g. 71">','<input type="text" data-col="Description" placeholder="e.g. Power Plant">'],kind):settingsTableRow(['<input type="text" data-col="Signatory Name" placeholder="Name">','<input type="text" data-col="Stamp" placeholder="Stamp">','<input type="text" data-col="License Number" placeholder="Licence No.">','<input type="text" data-col="Scope / Limitations" placeholder="Scope">'],kind)); tbody.insertAdjacentHTML('beforeend',rowHtml); tbody.lastElementChild.querySelector('input[data-col]') && tbody.lastElementChild.querySelector('input[data-col]').focus(); }
      function saveSettingsFromModal(){
        // Save owner info (only available on owner tab; store from DOM if on that tab, else use cached)
        var ownerNameEl=settingsBodyEl.querySelector('#settingsOwnerName');
        var ownerStampEl=settingsBodyEl.querySelector('#settingsOwnerStamp');
        if(ownerNameEl) LOG_OWNER_INFO.name=s(ownerNameEl.value);
        if(ownerStampEl) LOG_OWNER_INFO.stamp=s(ownerStampEl.value);
        if(settingsActiveTab==='aircraft'){
          applyAircraftGroupRows(collectSettingsTable('aircraft').map(function(r){ return {group:r.Group,reg:r['A/C Reg'],type:r['Aircraft Type']}; }));
        } else if(settingsActiveTab==='chapters'){
          applyChapterRows(collectSettingsTable('chapters').map(function(r){ return {chapter:r.Chapter,description:r.Description}; }));
        } else if(settingsActiveTab==='supervisors'){
          rebuildSupervisorState(collectSupervisorSettingsTable().map(function(r){ return {id:r.ID,name:r['Signatory Name'],stamp:r.Stamp,licence:r['License Number'],scope:r['Scope / Limitations'],date:r.Date}; }));
        }
        markSharedDatalistsDirty();
        settingsDirty=true;
        refreshUnsavedChangesState();
        renderAll();
        closeSettingsModal();
        scheduleAutoSave();
      }

      // ---- Event handlers ----
      filterBtn.onclick=function(){ openFilterPanel(); };
      closeFilterPanelBtn.onclick=function(){ closeFilterPanel(); };
      clearFiltersBtn.onclick=function(){ resetDraftFilters(); };
      filterForm.onsubmit=function(ev){ ev.preventDefault(); activeFilters=readFilterForm(); closeFilterPanel(); renderAll(); };
      filterModal.onclick=function(ev){ if(ev.target===filterModal) closeFilterPanel(); };
      filterStripEl.onclick=function(ev){ if(ev.target&&ev.target.closest&&ev.target.closest('[data-clear-filters="1"]')) clearFilters(); };
      if(searchInput) searchInput.addEventListener('input',function(){ searchQuery=s(this.value); syncSearchUi(); renderAll(); });
      if(searchInput) searchInput.addEventListener('keydown',function(ev){ if(ev.key==='Escape'&&hasActiveSearch()){ ev.preventDefault(); clearSearch(); } });
      if(clearSearchBtn) clearSearchBtn.onclick=function(){ clearSearch(); if(searchInput) searchInput.focus(); };
      filterForm.addEventListener('keydown',function(ev){ var input=ev.target&&ev.target.closest&&ev.target.closest('[data-filter-key]'); if(!input) return; if(ev.key==='Enter'||ev.key===','){ ev.preventDefault(); addDraftFilterValue(input.getAttribute('data-filter-key'),input.value); } if(ev.key==='Backspace'&&!s(input.value)){ var key=input.getAttribute('data-filter-key'),values=filterValues(draftFilters,key); if(values.length) removeDraftFilterValue(key,values.length-1); } });
      filterForm.addEventListener('change',function(ev){ var input=ev.target&&ev.target.closest&&ev.target.closest('[data-filter-key]'); if(!input) return; addDraftFilterValue(input.getAttribute('data-filter-key'),input.value); });
      filterForm.addEventListener('click',function(ev){ var removeBtn=ev.target&&ev.target.closest&&ev.target.closest('[data-remove-filter-key]'); if(removeBtn){ removeDraftFilterValue(removeBtn.getAttribute('data-remove-filter-key'),Number(removeBtn.getAttribute('data-remove-filter-index'))); return; } var box=ev.target&&ev.target.closest&&ev.target.closest('[data-filter-box]'); if(box){ var filterInput=box.querySelector('[data-filter-key]'); if(filterInput) filterInput.focus(); } });

      addBtn.onclick=function(){ modalBody.innerHTML=renderModalEditor(); wireModalFields(); modal.className='modal-backdrop open'; setTimeout(fitModalPreview,0); };

      // Save modal page - auto-save xlsx after adding
      saveBtn.onclick=async function(){ try { clearFail(); var pageData=collectModalPage(),rowsToAppend=manualPageToLogRows(pageData); if(!rowsToAppend.length){ modal.className='modal-backdrop'; return; } saveModalPageRows(rowsToAppend); renderAll(); modal.className='modal-backdrop'; scheduleAutoSave(); } catch(e){ fail('Page added but could not update: '+e.message); } };

      cancelBtn.onclick=function(){ modal.className='modal-backdrop'; };
      modal.onclick=function(ev){ var insidePage=ev.target&&ev.target.closest&&(ev.target.closest('.page')||ev.target.closest('.modal-actions')); if(!insidePage) modal.className='modal-backdrop'; };
      closeTaskDetailBtn.onclick=function(){ closeTaskDetail(); };
      taskDetailModal.onclick=function(ev){ if(!ev.target.closest('.detail-modal-card')) closeTaskDetail(); };
      closeTaskDetailBtn.onmousedown=function(ev){ ev.preventDefault(); };
      taskDetailModal.onmousedown=function(ev){ if(ev.target===taskDetailModal||ev.target.closest('.detail-modal-close')) ev.preventDefault(); };
      if(saveTaskDetailBtn) saveTaskDetailBtn.onclick=saveTaskDetail;
      if(detailChapterEl) detailChapterEl.addEventListener('input',previewTaskDetailForm);
      if(detailFaultEl) detailFaultEl.addEventListener('input',function(){ autoSizeDetailTextarea(detailFaultEl); previewTaskDetailForm(); });
      if(detailTaskEl) detailTaskEl.addEventListener('input',function(){ if(!taskDetailRewriteDirty) detailRewriteEl.value=detailTaskEl.value; autoSizeDetailTextarea(detailTaskEl); autoSizeDetailTextarea(detailRewriteEl); previewTaskDetailForm(); });
      if(detailRewriteEl) detailRewriteEl.addEventListener('input',function(){ taskDetailRewriteDirty=true; autoSizeDetailTextarea(detailRewriteEl); previewTaskDetailForm(); });
      if(confirmCancelBtn) confirmCancelBtn.onclick=function(){ closeConfirmDialog(false); };
      if(confirmOkBtn) confirmOkBtn.onclick=function(){ closeConfirmDialog(true); };
      if(confirmModal) confirmModal.onclick=function(ev){ if(ev.target===confirmModal) closeConfirmDialog(false); };

      infoBtn.onclick=function(){ openInfoModal(); };
      closeInfoModalBtn.onclick=function(){ closeInfoModal(); };
      infoModal.onclick=function(ev){ if(ev.target===infoModal) closeInfoModal(); };

      printBtn.onclick=function(ev){ if(ev) ev.stopPropagation(); setLoadOptionsOpen(false); setPrintOptionsOpen(!(printOptionsEl&&printOptionsEl.classList.contains('open'))); };
      printCurrentBtn.onclick=function(){ setPrintOptionsOpen(false); printCurrentPage(); };
      printAllBtn.onclick=function(){ setPrintOptionsOpen(false); printAllPages(); };
      document.addEventListener('click',function(ev){
        var insidePrint=!!(ev.target.closest&&(ev.target.closest('#printBtn')||ev.target.closest('#printOptions')));
        var insideLoad=!!(ev.target.closest&&(ev.target.closest('#loadBtn')||ev.target.closest('#loadOptions')));
        if(!insidePrint) setPrintOptionsOpen(false);
        if(!insideLoad) setLoadOptionsOpen(false);
      });

      // Save button - write xlsx
      saveFileBtn.onclick=async function(){ captureActiveEditorState(); setLoadingState(true,'Saving','Writing changes to cap741-data.xlsx...'); try { await flushLinkedRewrite(true); } finally { setLoadingState(false); } };

      // Load/Link button menu
      loadBtn.onclick=function(ev){
        if(ev) ev.stopPropagation();
        setPrintOptionsOpen(false);
        setLoadOptionsOpen(!(loadOptionsEl&&loadOptionsEl.classList.contains('open')));
      };
      if(loadExistingBtn) loadExistingBtn.onclick=async function(ev){
        if(ev) ev.stopPropagation();
        if(!filePickerSupported()){ fail('File picker not supported. Open this page via a local web server (e.g. VS Code Live Server) and use Chrome or Edge.'); return; }
        try {
          setLoadOptionsOpen(false);
          clearFail();
          setLoadingState(true,loadButtonMode==='link'?'Linking file':'Loading','Waiting for Excel file selection...');
          var handle=await pickWorkbookHandle();
          if(!handle) return;
          if(loadButtonMode==='link'){
            setLoadButtonMode('hidden');
            return;
          }
          setLoadingState(true,'Loading','Reading workbook...');
          var file=await handle.getFile();
          loadWorkbookFromArrayBuffer(await file.arrayBuffer());
          setLoadButtonMode('hidden');
          await renderAllWithLoading('Loading logbook','Rendering pages...');
          refreshUnsavedChangesState();
        } catch(e){
          if(e.name!=='AbortError') fail('Could not open Excel file: '+e.message);
        } finally {
          setLoadingState(false);
        }
      };
      if(createNewWorkbookBtn) createNewWorkbookBtn.onclick=async function(ev){
        if(ev) ev.stopPropagation();
        try {
          setLoadOptionsOpen(false);
          clearFail();
          setLoadingState(true,'Creating logbook','Choose where to save the new cap741-data.xlsx file...');
          await createNewWorkbookFile();
        } catch(e){
          if(e.name!=='AbortError') fail('Could not create Excel file: '+e.message);
        } finally {
          setLoadingState(false);
        }
      };

      // Settings
      if(settingsBtn) settingsBtn.onclick=openSettingsModal;
      if(closeSettingsModalBtn) closeSettingsModalBtn.onclick=closeSettingsModal;
      if(saveSettingsBtn) saveSettingsBtn.onclick=saveSettingsFromModal;
      if(settingsModal) settingsModal.onclick=function(ev){ if(ev.target===settingsModal) closeSettingsModal(); };
      if(settingsBodyEl) settingsBodyEl.addEventListener('click',function(ev){
        var add=ev.target.closest&&ev.target.closest('[data-settings-add]');
        if(add){ addSettingsRow(add.getAttribute('data-settings-add')); return; }
        var rem=ev.target.closest&&ev.target.closest('[data-settings-remove]');
        if(rem){ var tr=rem.closest('tr'); if(tr) tr.remove(); }
      });

      // Page editing events
      pagesEl.addEventListener('focusin',function(ev){ var cell=ev.target.closest&&ev.target.closest('.editable-cell'); if(cell&&cell.innerHTML==='&nbsp;') cell.innerHTML=''; });
      pagesEl.addEventListener('click',function(ev){
        if(ev.target&&ev.target.closest&&ev.target.closest('[data-clear-search="1"]')){ clearSearch(); return; }
        if(ev.target&&ev.target.closest&&ev.target.closest('[data-clear-all-results="1"]')){ clearFilters(); clearSearch(); return; }
        if(ev.target&&ev.target.closest&&ev.target.closest('[data-clear-filters="1"]')){ clearFilters(); return; }
        var clearSupervisorBtn=ev.target.closest&&ev.target.closest('[data-clear-supervisor]');
        if(clearSupervisorBtn){ clearSupervisorFields(clearSupervisorBtn); return; }
        var taskExpandBtn=ev.target.closest&&ev.target.closest('[data-open-task], [data-open-task-new]');
        if(taskExpandBtn){ openTaskDetailFromButton(taskExpandBtn); return; }
        var editableHeaderLine=ev.target.closest&&ev.target.closest('.editable-dots-line');
        if(editableHeaderLine){
          var headerInput=editableHeaderLine.querySelector('[data-group-field]');
          if(headerInput&&ev.target!==headerInput) focusInputAtEnd(headerInput);
          return;
        }
        var dateEntry=ev.target.closest&&ev.target.closest('.date-entry');
        if(dateEntry){ if(!pagesEl.contains(dateEntry)) return; if(ev.target!==dateEntry.querySelector('[data-date-picker]')) openDatePicker(dateEntry); return; }
        var td=ev.target.closest&&ev.target.closest('td');
        if(!td) return;
        if(td.classList.contains('c-sup')){ if(!ev.target.closest('input')){ var supNameInput=td.querySelector('[data-edit-field="Approval Name"], [data-new-row][data-edit-field="Approval Name"]'); if(supNameInput) supNameInput.focus(); } return; }
        var textInput=td.querySelector('[data-date-text]');
        if(textInput&&!ev.target.closest('.editable-cell')&&ev.target!==textInput){ textInput.focus(); return; }
        var fieldInput=td.querySelector('input.field-input[data-edit-field], input.field-input[data-new-row]');
        if(fieldInput&&ev.target!==fieldInput) fieldInput.focus();
        var taskCell=td.querySelector('.editable-cell.task-input');
        if(taskCell&&ev.target!==taskCell&&!ev.target.closest('input')){ taskCell.focus(); var sel=window.getSelection&&window.getSelection(); if(sel&&document.createRange){ var range=document.createRange(); range.selectNodeContents(taskCell); range.collapse(false); sel.removeAllRanges(); sel.addRange(range); } }
      });
      pagesEl.addEventListener('mousedown',function(ev){ var actionBtn=ev.target.closest&&ev.target.closest('[data-open-task], [data-open-task-new], [data-clear-supervisor]'); if(actionBtn) ev.preventDefault(); });
      pagesEl.addEventListener('change',function(ev){
        var datePicker=ev.target.closest&&ev.target.closest('[data-date-picker]');
        if(datePicker){ var dateEntry=datePicker.closest('.date-entry'); syncDateControl(dateEntry,datePicker.value); var dateInput=dateEntry&&dateEntry.querySelector('[data-date-text]'); if(dateInput&&updateRowFromEditor(dateInput)) syncSaveButtonState(false); }
        var supervisorInput=ev.target.closest&&ev.target.closest('[data-edit-field="Approval Name"], [data-new-row][data-edit-field="Approval Name"]');
        if(supervisorInput){ var supTd=supervisorInput.closest('td'),licenceInput=supTd&&supTd.querySelector('[data-edit-field="Aprroval Licence No."], [data-new-row][data-edit-field="Aprroval Licence No."]'); applySupervisorSuggestion(supervisorInput,licenceInput); updateRowFromEditor(supervisorInput); if(licenceInput) updateRowFromEditor(licenceInput); syncSaveButtonState(false); }
        var groupInput=ev.target.closest&&ev.target.closest('[data-group-field]');
        if(!groupInput) return;
        var page=groupInput.closest('.page');
        if(!page) return;
        var groupKey=page.getAttribute('data-group-key'),grpRows=rowsByGroupKey(groupKey);
        if(!grpRows.length) return;
        var nextType=grpRows[0]['Aircraft Type'],nextChapter=grpRows[0]['Chapter'],nextChapterDesc=grpRows[0]['Chapter Description'];
        if(groupInput.getAttribute('data-group-field')==='Aircraft Type'){
          nextType=valueOf(groupInput);
          for(var i=0;i<grpRows.length;i++) grpRows[i]['Aircraft Type']=nextType;
        } else if(groupInput.getAttribute('data-group-field')==='Chapter'){
          var parsedChapter=parseChapterValue(valueOf(groupInput));
          nextChapter=parsedChapter.chapter; nextChapterDesc=parsedChapter.chapterDesc;
          for(var j=0;j<grpRows.length;j++){ grpRows[j]['Chapter']=nextChapter; grpRows[j]['Chapter Description']=nextChapterDesc; }
        }
        syncBlankRowMetadata(page,nextType,nextChapter,nextChapterDesc);
        renderAll();
        refreshUnsavedChangesState();
        scheduleAutoSave();
      });
      pagesEl.addEventListener('input',function(ev){
        if(ev.target&&ev.target.classList&&ev.target.classList.contains('dots-input')) syncDotsInputSize(ev.target);
        var cell=ev.target.closest&&(ev.target.closest('.editable-cell')||ev.target.closest('[data-row-id]')||ev.target.closest('[data-new-row]'));
        if(!cell) return;
        if(ev.target.matches&&ev.target.matches('[data-group-field]')) return;
        updateRowFromEditor(cell);
        if(fieldNeedsLiveLayoutRefresh(cell.getAttribute('data-edit-field'))){
          if(hasActiveFilters()) return;
          scheduleLiveLayoutRefresh(captureEditorSnapshot(ev.target),120);
        }
      });
      pagesEl.addEventListener('blur',function(ev){
        var cell=ev.target.closest&&(ev.target.closest('.editable-cell')||ev.target.closest('[data-row-id]')||ev.target.closest('[data-new-row]'));
        if(!cell) return;
        var wasNewRow=cell.hasAttribute('data-new-row'),field=cell.getAttribute('data-edit-field');
        if(!updateRowFromEditor(cell)) return;
        if(wasNewRow||fieldAffectsRowLayout(field)) scheduleLayoutRefresh(250);
      },true);
      window.addEventListener('afterprint',function(){ clearPrintSelection(); setPrintOptionsOpen(false); });
      window.addEventListener('resize',function(){ if(modal.className.indexOf('open')!==-1) fitModalPreview(); });
      window.addEventListener('beforeunload',function(){ captureActiveEditorState(); });

      // ---- Startup ----
      rows=normalizeRows(rows);

      (async function cap741Startup(){
        setLoadingState(true,'Loading logbook','Reading cap741-data.xlsx...');
        await nextPaint();
        var loaded=false;
        var linked=false;

        // 1. Try direct fetch (works when served via HTTP / local dev server)
        try {
          await loadDefaultWorkbookData();
          loaded=true;
        } catch(fetchErr){
          // 2. Try stored file handle (user already picked file before)
          try {
            var storedHandle=await loadStoredHandle(LINKED_FILE_KEY);
            if(storedHandle&&handleIsWorkbook(storedHandle)&&await ensurePermission(storedHandle)){
              var file=await storedHandle.getFile();
              loadWorkbookFromArrayBuffer(await file.arrayBuffer());
              loaded=true;
              linked=true;
            }
          } catch(handleErr){}
        }

        if(!loaded){
          // Show load button for user to pick the xlsx file once
          setLoadButtonMode('load');
          setLoadingState(false);
          renderAll();
          return;
        }

        if(!linked){
          try {
            var rememberedHandle=await loadStoredHandle(LINKED_FILE_KEY);
            linked=!!(rememberedHandle&&handleIsWorkbook(rememberedHandle));
          } catch(linkStateErr){}
        }
        if(linked) setLoadButtonMode('hidden');
        else if(filePickerSupported()) setLoadButtonMode('link');
        renderAll();
        refreshUnsavedChangesState();
        setLoadingState(false);
      })();
    })();


