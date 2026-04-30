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
      var FLAG_RECORDS = [];
      var LOG_OWNER_INFO = { name: '', signature: '', stamp: '' };
      var PAGE_GROUPING_TYPE = 'type';
      var PAGE_GROUPING_GROUP = 'group';
      var FLAG_SECTION_PRIMARY = 'Primary';
      var FLAG_SECTION_MORE = 'More';
      var DEFAULT_APP_VIEW_SETTINGS = { showMindMap: false, pageGrouping: PAGE_GROUPING_TYPE, referenceOnlySave: true };
      var APP_VIEW_SETTINGS = { showMindMap: DEFAULT_APP_VIEW_SETTINGS.showMindMap, pageGrouping: DEFAULT_APP_VIEW_SETTINGS.pageGrouping, referenceOnlySave: DEFAULT_APP_VIEW_SETTINGS.referenceOnlySave };
      var SUPERVISOR_OPTIONS = [];
      var SUPERVISOR_LOOKUP = Object.create(null);
      var AIRCRAFT_GROUP_ROWS = [];
      var SUPERVISOR_RECORDS = [];
      var PAGE_SLOTS = 6;
      var ROW_TASK_LINES_PER_SLOT = 4;
      var TASK_TEXT_MEASURE_WIDTH_FALLBACK_PX = 318;
      var TASK_TEXT_LINE_HEIGHT_FALLBACK_PX = 9.44;
      var taskTextMeasureEl = null;
      var taskTextMeasureCache = null;
      var DB_NAME = 'cap741-file-handles';
      var DB_STORE = 'handles';
      var LINKED_FILE_KEY = 'cap741-main-file';
      var AUTO_LOAD_DEFAULT_KEY = 'cap741-auto-load-default';
      var STORAGE_SOURCE_KEY = 'cap741-storage-source';
      var GOOGLE_CLIENT_ID_KEY = 'cap741-google-client-id';
      var GOOGLE_CLIENT_ID = '647645362385-rj453g1g2g79hh9vorp1guvp7e9c8b9b.apps.googleusercontent.com';
      var DEFAULT_WORKBOOK_PATH = './cap741-data.xlsx';
      var BLANK_CHAPTER_FILTER = 'No Chapter';
      var SUPERVISOR_ID_FIELD = 'Supervisor ID';
      var LOG_HEADERS = ['Aircraft Type','A/C Reg','Chapter','Chapter Description','Date','Job No','FAULT','Task Detail','Rewriten for cap741','Flags',SUPERVISOR_ID_FIELD,'Approval Name','Approval stamp','Aprroval Licence No.','Signed'];
      var DATE_PLACEHOLDER = 'dd/MMM/yyyy';
      var FILTER_KEYS = ['aircraftType','aircraftReg','supervisor','chapter'];
      var STORAGE_SOURCE_NONE = 'none';
      var STORAGE_SOURCE_DEFAULT = 'default-excel';
      var STORAGE_SOURCE_EXCEL = 'excel';
      var STORAGE_SOURCE_GOOGLE = 'google-sheet';
      var GOOGLE_SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';
      var GOOGLE_SHEET_TITLES = ['Logbook','Aircraft','Chapters','Supervisors','Flags','Info'];
      var NEW_WORKBOOK_SUPERVISOR_NAME = 'Ioannis Orkos';
      var NEW_WORKBOOK_SUPERVISOR_LICENCE = 'UK.XX.XXXXXXX';
      var NEW_WORKBOOK_TASK_TEXT = 'Dummy data test one';
      var NEW_WORKBOOK_OWNER_NAME = 'User User';
      var OTHER_LAYOUT_DEFAULTS = { top: 36, headerHeight: 11, left: 14.5, dateStart: 14.5, regStart: 31.5, jobStart: 53.5, taskStart: 69, superStart: 142, end: 193, rowHeight: 13, textTop: 3, textLeft: 1 };
      var SUPERVISOR_LIST_ROWS_PER_PAGE = 10;

      // ---- DOM refs ----
      var errorBox = document.getElementById('errorBox');
      var errorTextEl = document.getElementById('errorText');
      var errorOkBtn = document.getElementById('errorOkBtn');
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
      var loadGoogleSheetBtn = document.getElementById('loadGoogleSheetBtn');
      var createGoogleSheetBtn = document.getElementById('createGoogleSheetBtn');
      var workbookOpenInput = document.getElementById('workbookOpenInput');
      var filterBtn = document.getElementById('filterBtn');
      var mindMapBtn = document.getElementById('mindMapBtn');
      var filterCountEl = document.getElementById('filterCount');
      var filterStripEl = document.getElementById('filterStrip');
      var addBtn = document.getElementById('addBlankPage');
      var printBtn = document.getElementById('printBtn');
      var printOptionsEl = document.getElementById('printOptions');
      var printCurrentBtn = document.getElementById('printCurrentBtn');
      var printCurrentOverlayBtn = document.getElementById('printCurrentOverlayBtn');
      var printOtherLayoutBtn = document.getElementById('printOtherLayoutBtn');
      var printAllBtn = document.getElementById('printAllBtn');
      var saveFileBtn = document.getElementById('saveFileBtn');
      var infoBtn = document.getElementById('infoBtn');
      var modal = document.getElementById('blankModal');
      var modalBody = document.getElementById('blankModalBody');
      var infoModal = document.getElementById('infoModal');
      var closeInfoModalBtn = document.getElementById('closeInfoModal');
      var mindMapModal = document.getElementById('mindMapModal');
      var closeMindMapModalBtn = document.getElementById('closeMindMapModal');
      var mindMapCanvasEl = document.getElementById('mindMapCanvas');
      var mindMapDetailEl = document.getElementById('mindMapDetail');
      var mindMapShellEl = mindMapCanvasEl ? mindMapCanvasEl.parentElement : null;
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
      var filterAircraftModeListEl = document.getElementById('filter-aircraft-mode-list');
      var filterAircraftRegListEl = document.getElementById('filter-aircraft-reg-list');
      var filterChapterListEl = document.getElementById('filter-chapter-list');
      var filterAircraftTypeFieldEl = filterAircraftTypeInput ? filterAircraftTypeInput.closest('.filter-field') : null;
      var filterAircraftTypeLabelEl = filterAircraftTypeFieldEl ? filterAircraftTypeFieldEl.querySelector('label[for="filterAircraftType"]') : null;
      var filterAircraftTypeHintEl = filterAircraftTypeFieldEl ? filterAircraftTypeFieldEl.querySelector('.filter-hint') : null;
      var taskDetailModal = document.getElementById('taskDetailModal');
      var confirmModal = document.getElementById('confirmModal');
      var confirmTitleEl = document.getElementById('confirmTitle');
      var confirmTextEl = document.getElementById('confirmText');
      var confirmCancelBtn = document.getElementById('confirmCancelBtn');
      var confirmOkBtn = document.getElementById('confirmOkBtn');
      var otherLayoutModal = document.getElementById('otherLayoutModal');
      var closeOtherLayoutModalBtn = document.getElementById('closeOtherLayoutModal');
      var resetOtherLayoutDefaultsBtn = document.getElementById('resetOtherLayoutDefaults');
      var printOtherLayoutConfirmBtn = document.getElementById('printOtherLayoutConfirm');
      var otherLayoutPreviewEl = document.getElementById('otherLayoutPreview');
      var otherOverlayTopMeasureValueEl = document.getElementById('otherOverlayTopMeasureValue');
      var otherOverlayRowMeasureValueEl = document.getElementById('otherOverlayRowMeasureValue');
      var otherOverlayTaskMeasureValueEl = document.getElementById('otherOverlayTaskMeasureValue');
      var otherOverlayTaskWidthMeasureValueEl = document.getElementById('otherOverlayTaskWidthMeasureValue');
      var otherOverlayDateMeasureValueEl = document.getElementById('otherOverlayDateMeasureValue');
      var otherOverlayRegMeasureValueEl = document.getElementById('otherOverlayRegMeasureValue');
      var otherOverlayJobMeasureValueEl = document.getElementById('otherOverlayJobMeasureValue');
      var otherOverlaySuperMeasureValueEl = document.getElementById('otherOverlaySuperMeasureValue');
      var otherOverlaySampleFrameEl = document.getElementById('otherOverlaySampleFrame');
      var otherOverlaySampleRowEl = document.getElementById('otherOverlaySampleRow');
      var otherOverlaySampleHeaderEl = otherOverlaySampleFrameEl ? otherOverlaySampleFrameEl.querySelector('.other-overlay-row-sample-header') : null;
      var otherOverlayDateHeaderEl = document.getElementById('otherOverlayDateHeader');
      var otherOverlayRegHeaderEl = document.getElementById('otherOverlayRegHeader');
      var otherOverlayTaskHeaderEl = document.getElementById('otherOverlayTaskHeader');
      var otherOverlaySuperHeaderEl = document.getElementById('otherOverlaySuperHeader');
      var googleSheetModal = document.getElementById('googleSheetModal');
      var closeGoogleSheetModalBtn = document.getElementById('closeGoogleSheetModal');
      var googleSheetModalTitleEl = document.getElementById('googleSheetModalTitle');
      var googleSheetModalCopyEl = document.getElementById('googleSheetModalCopy');
      var googleSheetInputWrapEl = document.getElementById('googleSheetInputWrap');
      var googleSheetInputLabelEl = document.getElementById('googleSheetInputLabel');
      var googleSheetUrlInputEl = document.getElementById('googleSheetUrlInput');
      var googleSheetResultRowEl = document.getElementById('googleSheetResultRow');
      var googleSheetResultLinkEl = document.getElementById('googleSheetResultLink');
      var googleSheetModalNoteEl = document.getElementById('googleSheetModalNote');
      var googleSheetCancelBtn = document.getElementById('googleSheetCancelBtn');
      var googleSheetOkBtn = document.getElementById('googleSheetOkBtn');
      var closeTaskDetailBtn = document.getElementById('closeTaskDetail');
      var detailFaultEl = document.getElementById('detailFault');
      var detailTaskEl = document.getElementById('detailTask');
      var detailRewriteEl = document.getElementById('detailRewrite');
      var sharedListsEl = document.getElementById('sharedLists');
      var detailFlagsPrimaryEl = document.getElementById('detailFlagsPrimary');
      var detailFlagsMoreWrapEl = document.getElementById('detailFlagsMoreWrap');
      var detailFlagsMoreEl = document.getElementById('detailFlagsMore');
      var detailFlagsEmptyEl = document.getElementById('detailFlagsEmpty');
      var saveBtn = document.getElementById('saveBlankPage');
      var cancelBtn = document.getElementById('cancelBlankPage');
      var modalShell = modal.querySelector('.modal-shell');
      var modalActions = modal.querySelector('.modal-actions');
      var settingsBtn = document.getElementById('settingsBtn');
      var settingsModal = document.getElementById('settingsModal');
      var closeSettingsModalBtn = document.getElementById('closeSettingsModal');
      var settingsBodyEl = document.getElementById('settingsBody');
      var saveSettingsBtn = document.getElementById('saveSettingsBtn');
      var printSupervisorsBtn = document.getElementById('printSupervisorsBtn');
      var supervisorPrintHost = document.getElementById('supervisorPrintHost');
      var saveTaskDetailBtn = document.getElementById('saveTaskDetail');
      var detailChapterEl = document.getElementById('detailChapter');

      // ---- Runtime state ----
      var layoutTimer = null;
      var autoSaveTimer = null;
      var liveLayoutEditorState = null;
      var saveInFlight = false;
      var saveQueued = false;
      var hasUnsavedChanges = false;
      var dataDirty = false;
      var lastSavedLogbookText = '';
      var confirmResolver = null;
      var googleSheetModalResolver = null;
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
      var normalizedSearchQuery = '';
      var searchRenderTimer = 0;
      var printMode = '';
      var settingsActiveTab = 'owner';
      var loadButtonMode = 'load';
      var workbookOpenInFlight = false;
      var linkedWorkbookName = '';
      var activeStorageSource = { type: STORAGE_SOURCE_NONE };
      var googleAccessToken = '';
      var googleTokenClient = null;
      var otherLayoutMeasurements = cloneMeasurementState(OTHER_LAYOUT_DEFAULTS);
      var savedLogbookRowSignatures = Object.create(null);
      var dirtyLogbookRowIds = Object.create(null);
      var savedLogbookRowOrder = '';
      var comparableOrderDirty = false;
      var aircraftOptionsByTypeCache = Object.create(null);
      var mindMapState = { selectedKind:'root', selectedId:'overview', summary:null };
      var mindMapRowHighlightTimer = 0;

      // ---- Filter state ----
      function emptyFilterState(){ return { aircraftType:[], aircraftReg:[], supervisor:[], chapter:[] }; }
      function cloneFilterState(state){ return { aircraftType:(state.aircraftType||[]).slice(), aircraftReg:(state.aircraftReg||[]).slice(), supervisor:(state.supervisor||[]).slice(), chapter:(state.chapter||[]).slice() }; }
      function filterValues(state, key){ return (state && state[key] && state[key].length) ? state[key] : []; }
      function totalFilterValueCount(state){ var total=0; for(var i=0;i<FILTER_KEYS.length;i++) total += filterValues(state,FILTER_KEYS[i]).length; return total; }
      function hasActiveFilters(){ return totalFilterValueCount(activeFilters) > 0; }
      function hasActiveSearch(){ return !!normalizedSearchQuery; }
      function activeFilterCount(){ return totalFilterValueCount(activeFilters); }
      function currentPageGroupingLabel(){ return currentPageGrouping()===PAGE_GROUPING_GROUP ? 'Aircraft Group' : 'Aircraft Type'; }
      function aircraftFilterChipLabel(){ return currentPageGrouping()===PAGE_GROUPING_GROUP ? 'Group' : 'Type'; }
      function aircraftFilterPlaceholderText(){ return currentPageGrouping()===PAGE_GROUPING_GROUP ? 'Add aircraft group' : 'Add aircraft type'; }
      function aircraftFilterValueForRow(row){ return currentPageGrouping()===PAGE_GROUPING_GROUP ? rowAircraftGroupLabel(row) : aircraftLabel(row); }
      function availableAircraftGroupValues(){
        var seen=Object.create(null),vals=[],i,item,group;
        for(i=0;i<(AIRCRAFT_GROUP_ROWS||[]).length;i++){
          item=AIRCRAFT_GROUP_ROWS[i]||{};
          group=normalizeMindMapGroupLabel(s(item.group));
          if(!group||seen[group]) continue;
          seen[group]=true;
          vals.push(group);
        }
        for(i=0;i<rows.length;i++){
          group=rowAircraftGroupLabel(rows[i]);
          if(!group||seen[group]) continue;
          seen[group]=true;
          vals.push(group);
        }
        vals.sort(mindMapGroupSort);
        return vals;
      }
      function filterAircraftModeOptionsHtml(){
        var vals=currentPageGrouping()===PAGE_GROUPING_GROUP ? availableAircraftGroupValues() : usedAircraftTypes(),html='',i;
        if(!vals.length) vals=currentPageGrouping()===PAGE_GROUPING_GROUP ? availableAircraftGroupValues() : [];
        if(!vals.length && currentPageGrouping()!==PAGE_GROUPING_GROUP) return aircraftTypeOptionsHtml();
        for(i=0;i<vals.length;i++) html+='<option value="'+esc(vals[i])+'"></option>';
        return html;
      }
      function syncFilterAircraftModeUi(){
        if(filterAircraftTypeLabelEl) filterAircraftTypeLabelEl.textContent=currentPageGroupingLabel();
        if(filterAircraftTypeInput) filterAircraftTypeInput.placeholder=aircraftFilterPlaceholderText();
        if(filterAircraftTypeHintEl) filterAircraftTypeHintEl.textContent='Press Enter or comma to add more than one.';
        if(filterAircraftModeListEl) filterAircraftModeListEl.innerHTML=filterAircraftModeOptionsHtml();
      }
      function activeFilterChips(){ var chips=[]; for(var i=0;i<activeFilters.aircraftType.length;i++) chips.push({label:aircraftFilterChipLabel(),value:activeFilters.aircraftType[i]}); for(var j=0;j<activeFilters.aircraftReg.length;j++) chips.push({label:'A/C',value:activeFilters.aircraftReg[j]}); for(var k=0;k<activeFilters.supervisor.length;k++) chips.push({label:'Supervisor',value:activeFilters.supervisor[k]}); for(var m=0;m<activeFilters.chapter.length;m++) chips.push({label:'Chapter',value:activeFilters.chapter[m]}); return chips; }
      function syncFilterButtonState(){ if(!filterBtn||!filterCountEl) return; var count=activeFilterCount(); filterBtn.classList.toggle('active',count>0); filterCountEl.hidden=count<1; filterCountEl.textContent=String(count); }
      function syncSearchUi(){ if(!searchInput||!clearSearchBtn) return; if(searchInput.value!==searchQuery) searchInput.value=searchQuery; clearSearchBtn.hidden=!hasActiveSearch(); }
      function renderFilterStrip(){ if(!filterStripEl) return; var chips=activeFilterChips(); if(!chips.length){ filterStripEl.className='filter-strip'; filterStripEl.innerHTML=''; return; } var html=[]; for(var i=0;i<chips.length;i++) html.push('<span class="filter-chip">'+esc(chips[i].label+': '+chips[i].value)+'</span>'); filterStripEl.className='filter-strip open'; filterStripEl.innerHTML='<div class="filter-strip-text"><strong>Filters:</strong> '+html.join('')+'</div><button type="button" data-clear-filters="1">Clear filters</button>'; }
      function normalizeFilterEntry(key, value){ var raw=s(value); if(!raw) return ''; if(key==='aircraftReg') return raw.toUpperCase(); if(key==='supervisor'){ var sv=normalizeSupervisorValue(raw); return sv.name||raw; } if(key==='chapter'){ var normalized=normalizedText(raw); if(normalized==='no chapter'||normalized==='blank chapter'||normalized==='empty chapter'||normalized==='[no chapter]') return BLANK_CHAPTER_FILTER; } return raw; }
      function uniqueFilterValues(values){ var out=[],seen=Object.create(null); for(var i=0;i<values.length;i++){ var n=normalizedText(values[i]); if(!n||seen[n]) continue; seen[n]=true; out.push(values[i]); } return out; }
      function filterInputForKey(key){ if(key==='aircraftType') return filterAircraftTypeInput; if(key==='aircraftReg') return filterAircraftRegInput; if(key==='supervisor') return filterSupervisorInput; if(key==='chapter') return filterChapterInput; return null; }
      function filterChipHostForKey(key){ if(key==='aircraftType') return filterAircraftTypeChipsEl; if(key==='aircraftReg') return filterAircraftRegChipsEl; if(key==='supervisor') return filterSupervisorChipsEl; if(key==='chapter') return filterChapterChipsEl; return null; }
      function aircraftOptionsHtmlForTypes(types){ if(!types||!types.length) return aircraftOptionsHtml(); var seen=Object.create(null),regs=[]; for(var reg in AIRCRAFT_MAP){ if(!Object.prototype.hasOwnProperty.call(AIRCRAFT_MAP,reg)) continue; for(var i=0;i<types.length;i++){ if(AIRCRAFT_MAP[reg]===types[i]&&!seen[reg]){ seen[reg]=true; regs.push(reg); } } } regs.sort(); if(!regs.length) return aircraftOptionsHtml(); var html=''; for(var j=0;j<regs.length;j++) html+='<option value="'+esc(regs[j])+'"></option>'; return html; }
      function aircraftOptionsHtmlForGroups(groups){
        if(!groups||!groups.length) return aircraftOptionsHtml();
        var seen=Object.create(null),wanted=Object.create(null),regs=[],i,item,row,group,reg;
        for(i=0;i<groups.length;i++) wanted[normalizeMindMapGroupLabel(groups[i])]=true;
        for(i=0;i<(AIRCRAFT_GROUP_ROWS||[]).length;i++){
          item=AIRCRAFT_GROUP_ROWS[i]||{};
          group=normalizeMindMapGroupLabel(s(item.group));
          reg=s(item.reg).toUpperCase();
          if(!reg||!wanted[group]||seen[reg]) continue;
          seen[reg]=true;
          regs.push(reg);
        }
        for(i=0;i<rows.length;i++){
          row=rows[i]||{};
          reg=s(row['A/C Reg']).toUpperCase();
          group=rowAircraftGroupLabel(row);
          if(!reg||!wanted[group]||seen[reg]) continue;
          seen[reg]=true;
          regs.push(reg);
        }
        regs.sort();
        if(!regs.length) return aircraftOptionsHtml();
        var html='';
        for(i=0;i<regs.length;i++) html+='<option value="'+esc(regs[i])+'"></option>';
        return html;
      }
      function aircraftOptionsHtmlForGroupLabel(groupLabel){
        groupLabel=normalizeMindMapGroupLabel(groupLabel);
        if(!groupLabel) return aircraftOptionsHtml();
        return aircraftOptionsHtmlForGroups([groupLabel]);
      }
      function syncFilterRegList(){ if(!filterAircraftRegListEl) return; filterAircraftRegListEl.innerHTML=currentPageGrouping()===PAGE_GROUPING_GROUP?aircraftOptionsHtmlForGroups(draftFilters.aircraftType):aircraftOptionsHtmlForTypes(draftFilters.aircraftType); }
      function syncFilterChapterList(){ if(!filterChapterListEl) return; filterChapterListEl.innerHTML='<option value="'+esc(BLANK_CHAPTER_FILTER)+'"></option>'+chapterOptionsHtml(); }
      function renderDraftFilterField(key){ var host=filterChipHostForKey(key),input=filterInputForKey(key); if(!host||!input) return; var values=filterValues(draftFilters,key),html=''; for(var i=0;i<values.length;i++) html+='<span class="multi-filter-chip"><span>'+esc(values[i])+'</span><button type="button" data-remove-filter-key="'+key+'" data-remove-filter-index="'+i+'" aria-label="Remove '+esc(values[i])+'">x</button></span>'; host.innerHTML=html; }
      function renderDraftFilters(){ syncFilterAircraftModeUi(); for(var i=0;i<FILTER_KEYS.length;i++) renderDraftFilterField(FILTER_KEYS[i]); syncFilterRegList(); syncFilterChapterList(); }
      function addDraftFilterValue(key, value){ var rawParts=s(value).split(',').map(function(p){ return s(p); }).filter(Boolean); if(!rawParts.length) return; var nextValues=filterValues(draftFilters,key).slice(); for(var i=0;i<rawParts.length;i++){ var normalized=normalizeFilterEntry(key,rawParts[i]); if(normalized) nextValues.push(normalized); } draftFilters[key]=uniqueFilterValues(nextValues); var input=filterInputForKey(key); if(input) input.value=''; renderDraftFilterField(key); if(key==='aircraftType') syncFilterRegList(); }
      function commitPendingDraftInputs(){ for(var i=0;i<FILTER_KEYS.length;i++){ var key=FILTER_KEYS[i],input=filterInputForKey(key); if(input&&s(input.value)) addDraftFilterValue(key,input.value); } }
      function removeDraftFilterValue(key, index){ var values=filterValues(draftFilters,key).slice(); values.splice(index,1); draftFilters[key]=values; renderDraftFilterField(key); if(key==='aircraftType') syncFilterRegList(); }
      function resetDraftFilters(){ draftFilters=emptyFilterState(); for(var i=0;i<FILTER_KEYS.length;i++){ var input=filterInputForKey(FILTER_KEYS[i]); if(input) input.value=''; } renderDraftFilters(); }
      function openFilterPanel(){ if(!filterModal) return; draftFilters=cloneFilterState(activeFilters); for(var i=0;i<FILTER_KEYS.length;i++){ var input=filterInputForKey(FILTER_KEYS[i]); if(input) input.value=''; } renderDraftFilters(); filterModal.className='modal-backdrop filter-backdrop open'; setTimeout(function(){ if(filterAircraftTypeInput) filterAircraftTypeInput.focus(); },0); }
      function closeFilterPanel(){ if(filterModal) filterModal.className='modal-backdrop filter-backdrop'; }
      function mindMapClip(text, limit){ text=s(text); limit=Number(limit)||0; if(!text||!limit||text.length<=limit) return text; return text.slice(0,Math.max(0,limit-3))+'...'; }
      function mindMapTextSort(a,b){ return s(a).localeCompare(s(b),undefined,{numeric:true}); }
      function isMindMapUngroupedChapter(value){
        var label=s(value);
        return !label||label===BLANK_CHAPTER_FILTER||label==='No Chapter'||label==='Ungrouped';
      }
      function mindMapChapterLabel(chapterCode, chapterDesc){
        var code=s(chapterCode),desc=s(chapterDesc);
        return code?(code+(desc?' - '+desc:'')):'Ungrouped';
      }
      function mindMapChapterTextSort(a,b){ var left=s(a),right=s(b); if(isMindMapUngroupedChapter(left)) return isMindMapUngroupedChapter(right)?0:1; if(isMindMapUngroupedChapter(right)) return -1; return left.localeCompare(right,undefined,{numeric:true}); }
      function mindMapChapterSort(a,b){ var left=s(a&&a.chapter),right=s(b&&b.chapter); if(isMindMapUngroupedChapter(left)) return isMindMapUngroupedChapter(right)?0:1; if(isMindMapUngroupedChapter(right)) return -1; return left.localeCompare(right,undefined,{numeric:true}); }
      function mindMapChapterLookup(){
        var lookup=Object.create(null),source=chapterDataStore(),i;
        function remember(chapter, description){
          chapter=s(chapter);
          description=s(description);
          if(!chapter) return;
          if(!lookup[chapter]) lookup[chapter]={ chapter:chapter, description:description };
          else if(!lookup[chapter].description&&description) lookup[chapter].description=description;
        }
        for(i=0;i<source.length;i++) remember(source[i]&&source[i].chapter,source[i]&&source[i].description);
        for(i=0;i<CHAPTER_OPTIONS.length;i++){ var parsed=parseChapterValue(CHAPTER_OPTIONS[i]); remember(parsed.chapter,parsed.chapterDesc); }
        return lookup;
      }
      function normalizeMindMapGroupLabel(value){
        var label=s(value),normalized=normalizedText(label);
        if(!label||normalized==='unassigned group'||normalized==='ungrouped'||normalized==='unassigned') return 'Ungrouped';
        return label;
      }
      function normalizePageGroupLabel(value){
        var label=s(value),normalized=normalizedText(label);
        if(!label||normalized==='unassigned group'||normalized==='unassigned') return '';
        if(normalized==='ungrouped') return 'Ungrouped';
        return label;
      }
      function resolvedRowPageGroupLabel(row){
        var reg=s(row&&row['A/C Reg']).toUpperCase(),type=s(aircraftLabel(row)),record=reg?aircraftReferenceRecordForReg(reg):null,label='';
        if(reg) label=s(record&&record.group)||s(aircraftGroupForType(type));
        else if(type) label=s(aircraftGroupForType(type));
        return normalizePageGroupLabel(label);
      }
      function rememberedRowPageGroupLabel(row){ return s(row&&row.__pageGroupLabel); }
      function syncRowPageGroupLabel(row, sourceEl){
        if(!row) return '';
        var page=sourceEl&&sourceEl.closest&&sourceEl.closest('.page'),label=resolvedRowPageGroupLabel(row);
        if(!label) label=s(page&&page.getAttribute('data-page-group-label'))||rememberedRowPageGroupLabel(row);
        row.__pageGroupLabel=s(label);
        return row.__pageGroupLabel;
      }
      function rowAircraftGroupLabel(row){
        var reg=s(row&&row['A/C Reg']).toUpperCase(),type=s(aircraftLabel(row)),record=reg?aircraftReferenceRecordForReg(reg):null,label='';
        if(reg) label=s(record&&record.group)||s(aircraftGroupForType(type));
        else if(type) label=s(aircraftGroupForType(type));
        return label?normalizeMindMapGroupLabel(label):'';
      }
      function rowAircraftTypeLabel(row){ return s(aircraftLabel(row))||'Unassigned Aircraft Type'; }
      function mindMapGroupSort(a,b){
        var left=normalizeMindMapGroupLabel(a&&a.label||a),right=normalizeMindMapGroupLabel(b&&b.label||b);
        if(left==='Ungrouped'&&right!=='Ungrouped') return 1;
        if(right==='Ungrouped'&&left!=='Ungrouped') return -1;
        return left.localeCompare(right,undefined,{numeric:true});
      }
      function mindMapExpandedGroups(){
        if(!mindMapState.expandedGroups) mindMapState.expandedGroups=Object.create(null);
        return mindMapState.expandedGroups;
      }
      function mindMapNodeOffsets(){
        if(!mindMapState.nodeOffsets) mindMapState.nodeOffsets=Object.create(null);
        return mindMapState.nodeOffsets;
      }
      function mindMapNodeOffset(key){
        var offsets=mindMapNodeOffsets();
        if(!offsets[key]) offsets[key]={ x:0, y:0 };
        return offsets[key];
      }
      function mindMapDomKey(kind, id){ return String(kind||'node')+'-'+encodeURIComponent(String(id==null?'':id)); }
      function mindMapJobSort(a,b){
        var left=Number(a&&a.sortDate),right=Number(b&&b.sortDate),blank=8640000000000000;
        if(!isFinite(left)||left===blank) left=-1;
        if(!isFinite(right)||right===blank) right=-1;
        if(left!==right) return right-left;
        return mindMapTextSort(a&&a.label,b&&b.label);
      }
      function buildMindMapGroupTypeItems(group){
        var jobs=(group&&group.jobs)||[],lookup=Object.create(null),list=[],i,job,typeLabel,item,chapterLabel;
        for(i=0;i<jobs.length;i++){
          job=jobs[i]||{};
          typeLabel=s(job.type)||'Unassigned Aircraft Type';
          if(!lookup[typeLabel]) lookup[typeLabel]={ id:typeLabel, label:typeLabel, entryCount:0, jobs:[], _chapters:Object.create(null), _regs:Object.create(null) };
          item=lookup[typeLabel];
          item.entryCount++;
          item.jobs.push(job);
          chapterLabel=s(job.chapterLabel);
          if(chapterLabel) item._chapters[chapterLabel]=chapterLabel;
          if(s(job.reg)) item._regs[s(job.reg)]=s(job.reg);
        }
        for(typeLabel in lookup) if(Object.prototype.hasOwnProperty.call(lookup,typeLabel)){
          item=lookup[typeLabel];
          item.jobs.sort(mindMapJobSort);
          item.chapters=Object.keys(item._chapters).sort(mindMapChapterTextSort);
          item.registrations=Object.keys(item._regs).sort(mindMapTextSort);
          item.chapterItems=buildMindMapTypeChapterItems(item);
          item.regItems=item.chapterItems;
          delete item._chapters;
          delete item._regs;
          list.push(item);
        }
        list.sort(function(a,b){ return mindMapTextSort(a&&a.label,b&&b.label); });
        return list;
      }
      function mindMapMetricHtml(label, value){ return '<div class="mindmap-metric"><span class="mindmap-metric-value">'+esc(value)+'</span><span class="mindmap-metric-label">'+esc(label)+'</span></div>'; }
      function mindMapPillListHtml(items, emptyText){
        if(!items||!items.length) return '<div class="mindmap-empty-mini">'+esc(emptyText||'Nothing recorded yet.')+'</div>';
        var html=[];
        for(var i=0;i<items.length;i++) html.push('<span class="mindmap-pill">'+esc(items[i])+'</span>');
        return '<div class="mindmap-pill-list">'+html.join('')+'</div>';
      }
      function mindMapDetailCardHtml(kind, id, title, meta, copy){
        return '<button class="mindmap-detail-card" type="button" data-mindmap-node-kind="'+esc(kind)+'" data-mindmap-node-id="'+esc(id)+'"><strong>'+esc(title)+'</strong>'+(meta?'<span class="mindmap-detail-card-meta">'+esc(meta)+'</span>':'')+(copy?'<span class="mindmap-detail-card-copy">'+esc(copy)+'</span>':'')+'</button>';
      }
      function mindMapJobLinkHtml(job){
        var meta=[s(job.date),s(job.reg),s(job.chapterLabel)].filter(Boolean).join(' | ');
        return '<button class="mindmap-job-link" type="button" data-mindmap-job-row-id="'+esc(job.rowId)+'"><strong>'+esc(job.label)+'</strong>'+(meta?'<span class="mindmap-job-link-meta">'+esc(meta)+'</span>':'')+'</button>';
      }
      function mindMapChapterAccordionHtml(chapter){
        var jobs=[],meta=[],notes=[],i;
        for(i=0;i<chapter.jobs.length;i++) jobs.push(mindMapJobLinkHtml(chapter.jobs[i]));
        meta.push(String(Number(chapter.entryCount)||0)+' entries');
        meta.push(String((chapter.jobs||[]).length)+' jobs');
        if(chapter.registrations&&chapter.registrations.length) meta.push(String(chapter.registrations.length)+' regs');
        if(s(chapter.description)) notes.push(s(chapter.description));
        if(chapter.registrations&&chapter.registrations.length) notes.push('Regs: '+mindMapClip(chapter.registrations.join(', '),120));
        return '<details class="mindmap-accordion"><summary class="mindmap-accordion-summary"><span class="mindmap-accordion-head"><strong class="mindmap-accordion-title">'+esc(chapter.label)+'</strong><span class="mindmap-accordion-meta">'+esc(meta.join(' | '))+'</span></span><span class="mindmap-accordion-toggle" aria-hidden="true"></span></summary>'+(notes.length?'<div class="mindmap-accordion-copy">'+esc(notes.join(' '))+'</div>':'')+'<div class="mindmap-accordion-body"><div class="mindmap-job-links">'+(jobs.join('')||'<div class="mindmap-empty-mini">No jobs recorded in this chapter.</div>')+'</div></div></details>';
      }
      function mindMapChapterAccordionListHtml(items, emptyText){
        if(!items||!items.length) return '<div class="mindmap-empty-mini">'+esc(emptyText||'No chapters recorded yet.')+'</div>';
        var html=[],i;
        for(i=0;i<items.length;i++) html.push(mindMapChapterAccordionHtml(items[i]));
        return '<div class="mindmap-accordion-list">'+html.join('')+'</div>';
      }
      function mindMapBranchHtml(title, count, itemsHtml, branchClass){
        var total=Number(count)||0;
        return '<section class="mindmap-branch '+esc(branchClass||'')+'"><div class="mindmap-node mindmap-node-hub" aria-hidden="true"><span class="mindmap-node-title">'+esc(title)+'</span><span class="mindmap-node-meta">'+esc(String(total)+' nodes in this branch')+'</span></div><div class="mindmap-node-list">'+(itemsHtml||'<div class="mindmap-empty-mini">Nothing recorded yet.</div>')+'</div></section>';
      }
      function renderMindMapNode(kind, id, title, meta, extraClass){
        var active=mindMapState.selectedKind===kind&&mindMapState.selectedId===id;
        return '<button class="mindmap-node '+(extraClass||'')+(active?' is-active':'')+'" type="button" data-mindmap-node-kind="'+esc(kind)+'" data-mindmap-node-id="'+esc(id)+'"><span class="mindmap-node-title">'+esc(title)+'</span>'+(meta?'<span class="mindmap-node-meta">'+esc(meta)+'</span>':'')+'</button>';
      }
      function mindMapNodePositionStyle(x, y, delay){
        var style='left:'+Math.round(Number(x)||0)+'px;top:'+Math.round(Number(y)||0)+'px;';
        if(delay!=null) style+='--mindmap-delay:'+String(delay)+'s;';
        return style;
      }
      function mindMapConnectorPath(fromX, fromY, toX, toY){
        var dx=(Number(toX)||0)-(Number(fromX)||0),curve=Math.max(86,Math.round(Math.abs(dx)*0.34)),dir=dx>=0?1:-1,c1x=(Number(fromX)||0)+(curve*dir),c2x=(Number(toX)||0)-(curve*dir);
        return 'M'+Math.round(Number(fromX)||0)+' '+Math.round(Number(fromY)||0)+' C '+Math.round(c1x)+' '+Math.round(Number(fromY)||0)+' '+Math.round(c2x)+' '+Math.round(Number(toY)||0)+' '+Math.round(Number(toX)||0)+' '+Math.round(Number(toY)||0);
      }
      function mindMapCenteredPositions(count, centerY, gap){
        var out=[],i,start;
        count=Number(count)||0;
        if(count<1) return out;
        start=(Number(centerY)||0)-(((count-1)*(Number(gap)||0))/2);
        for(i=0;i<count;i++) out.push(Math.round(start+(i*(Number(gap)||0))));
        return out;
      }
      function mindMapFanPositions(count, centerX, centerY, radiusX, radiusY, startDeg, endDeg){
        var out=[],i,angle,start,end;
        count=Number(count)||0;
        if(count<1) return out;
        start=Number(startDeg);
        end=Number(endDeg);
        if(!isFinite(start)) start=-72;
        if(!isFinite(end)) end=72;
        if(count===1){
          angle=((start+end)/2)*(Math.PI/180);
          out.push({ x:Math.round((Number(centerX)||0)+(Math.cos(angle)*(Number(radiusX)||0))), y:Math.round((Number(centerY)||0)+(Math.sin(angle)*(Number(radiusY)||0))) });
          return out;
        }
        for(i=0;i<count;i++){
          angle=(start+(((end-start)*i)/(count-1)))*(Math.PI/180);
          out.push({ x:Math.round((Number(centerX)||0)+(Math.cos(angle)*(Number(radiusX)||0))), y:Math.round((Number(centerY)||0)+(Math.sin(angle)*(Number(radiusY)||0))) });
        }
        return out;
      }
      function mindMapCirclePositions(count, centerX, centerY, radiusX, radiusY, startDeg){
        var out=[],i,angle,start,step;
        count=Number(count)||0;
        if(count<1) return out;
        start=Number(startDeg);
        if(!isFinite(start)) start=-90;
        step=360/count;
        for(i=0;i<count;i++){
          angle=(start+(step*i))*(Math.PI/180);
          out.push({ x:Math.round((Number(centerX)||0)+(Math.cos(angle)*(Number(radiusX)||0))), y:Math.round((Number(centerY)||0)+(Math.sin(angle)*(Number(radiusY)||0))), angle:angle });
        }
        return out;
      }
      function findMindMapGroup(summary, groupId){
        var groups=(summary&&summary.groups)||[],i,target=s(groupId);
        if(!target) return null;
        for(i=0;i<groups.length;i++) if(s(groups[i]&&groups[i].id)===target) return groups[i];
        return null;
      }
      function findMindMapGroupType(group, typeId){
        var items=(group&&group.typeItems)||[],i,target=s(typeId);
        if(!target) return null;
        for(i=0;i<items.length;i++) if(s(items[i]&&items[i].id)===target) return items[i];
        return null;
      }
      function findMindMapTypeChapter(typeItem, chapterId){
        var items=(typeItem&&(typeItem.chapterItems||typeItem.regItems))||[],i,target=s(chapterId);
        if(!target) return null;
        for(i=0;i<items.length;i++) if(s(items[i]&&items[i].id)===target) return items[i];
        return null;
      }
      function findMindMapRegChapter(regItem, chapterId){
        var items=(regItem&&regItem.chapterItems)||[],i,target=s(chapterId);
        if(!target) return null;
        for(i=0;i<items.length;i++) if(s(items[i]&&items[i].id)===target) return items[i];
        return null;
      }
      function buildMindMapRegChapterItems(regItem){
        var jobs=(regItem&&regItem.jobs)||[],lookup=Object.create(null),list=[],i,job,chapterId,chapterLabel,item;
        for(i=0;i<jobs.length;i++){
          job=jobs[i]||{};
          chapterId=s(job.chapter)||BLANK_CHAPTER_FILTER;
          chapterLabel=s(job.chapterLabel)||'Ungrouped';
          if(!lookup[chapterId]) lookup[chapterId]={ id:chapterId, label:chapterLabel, description:s(job.chapterDesc), entryCount:0, jobs:[] };
          item=lookup[chapterId];
          item.entryCount++;
          item.jobs.push(job);
          if(!item.description&&s(job.chapterDesc)) item.description=s(job.chapterDesc);
        }
        for(chapterId in lookup) if(Object.prototype.hasOwnProperty.call(lookup,chapterId)){
          item=lookup[chapterId];
          item.jobs.sort(mindMapJobSort);
          list.push(item);
        }
        list.sort(function(a,b){ return mindMapChapterTextSort(a&&a.label,b&&b.label); });
        return list;
      }
      function buildMindMapTypeChapterItems(typeItem){
        var jobs=(typeItem&&typeItem.jobs)||[],lookup=Object.create(null),list=[],i,job,chapterId,chapterLabel,item,regLabel;
        for(i=0;i<jobs.length;i++){
          job=jobs[i]||{};
          chapterId=s(job.chapter)||BLANK_CHAPTER_FILTER;
          chapterLabel=s(job.chapterLabel)||'Ungrouped';
          if(!lookup[chapterId]) lookup[chapterId]={ id:chapterId, label:chapterLabel, description:s(job.chapterDesc), entryCount:0, jobs:[], _regs:Object.create(null) };
          item=lookup[chapterId];
          item.entryCount++;
          item.jobs.push(job);
          regLabel=s(job.reg);
          if(regLabel) item._regs[regLabel]=regLabel;
          if(!item.description&&s(job.chapterDesc)) item.description=s(job.chapterDesc);
        }
        for(chapterId in lookup) if(Object.prototype.hasOwnProperty.call(lookup,chapterId)){
          item=lookup[chapterId];
          item.jobs.sort(mindMapJobSort);
          item.registrations=Object.keys(item._regs).sort(mindMapTextSort);
          delete item._regs;
          list.push(item);
        }
        list.sort(function(a,b){ return mindMapChapterTextSort(a&&a.label,b&&b.label); });
        return list;
      }
      function buildMindMapLayout(summary){
        var sceneWidth=4080,hubX=2040,groups=summary.groups||[],expandedGroups=mindMapExpandedGroups(),groupOrbitX=Math.max(620,Math.min(980,560+(groups.length*22))),groupOrbitY=Math.max(430,Math.min(760,360+(groups.length*18))),i,domKey,baseX,baseY,offset,node,groupPositions,typePositions,groupItem,typeItems,ringPos,groupAngleDeg,typeRadiusX,typeRadiusY,minX,maxX,minY,maxY,shiftX,shiftY,layout={sceneWidth:sceneWidth,sceneHeight:2400,rootX:hubX,rootY:1200,hubs:[],nodes:[],links:[]};
        layout.hubs.push({ key:'groups-hub', domKey:'groups-hub', label:'Aircraft Groups', meta:String(Number(summary.groupCount)||0)+' groups', x:hubX, y:layout.rootY, className:'mindmap-scene-hub-groups' });
        groupPositions=mindMapCirclePositions(groups.length,hubX,layout.rootY,groupOrbitX,groupOrbitY,-90);
        minX=hubX-180;
        maxX=hubX+180;
        minY=layout.rootY-120;
        maxY=layout.rootY+120;
        for(i=0;i<groups.length;i++){
          groupItem=groups[i];
          ringPos=groupPositions[i]||{ x:hubX, y:layout.rootY, angle:-Math.PI/2 };
          domKey=mindMapDomKey('group',groupItem.id);
          baseX=Math.round(ringPos.x);
          baseY=Math.round(ringPos.y);
          offset=mindMapNodeOffset(domKey);
          node={ kind:'group', id:groupItem.id, domKey:domKey, label:groupItem.label, meta:groupItem.entryCount+' entries | '+groupItem.typeItems.length+' types', x:baseX+offset.x, y:baseY+offset.y, baseX:baseX, baseY:baseY, orbitAngle:ringPos.angle, linkFromX:hubX, linkFromY:layout.rootY, className:'mindmap-scene-node-group'+(expandedGroups[groupItem.id]?' is-branch-open':''), delay:(i%8)*0.05, group:groupItem };
          layout.nodes.push(node);
          layout.links.push({ domKey:domKey, fromX:hubX, fromY:layout.rootY, toX:node.x, toY:node.y, className:'mindmap-link-branch' });
          minX=Math.min(minX,node.x-210);
          maxX=Math.max(maxX,node.x+210);
          minY=Math.min(minY,node.y-90);
          maxY=Math.max(maxY,node.y+90);
          if(expandedGroups[groupItem.id]){
            typeItems=groupItem.typeItems||[];
            groupAngleDeg=(node.orbitAngle*180/Math.PI);
            typeRadiusX=Math.max(240,Math.min(420,180+(typeItems.length*18)));
            typeRadiusY=Math.max(170,Math.min(340,130+(typeItems.length*15)));
            typePositions=mindMapFanPositions(typeItems.length,node.x+(Math.cos(node.orbitAngle)*140),node.y+(Math.sin(node.orbitAngle)*140),typeRadiusX,typeRadiusY,groupAngleDeg-86,groupAngleDeg+86);
            for(var typeIndex=0;typeIndex<typeItems.length;typeIndex++){
              node={ kind:'group-type', groupId:groupItem.id, id:typeItems[typeIndex].id, domKey:mindMapDomKey('group-type',groupItem.id+'::'+typeItems[typeIndex].id), label:typeItems[typeIndex].label, meta:typeItems[typeIndex].entryCount+' entries | '+typeItems[typeIndex].chapterItems.length+' chapters', x:Math.round((typePositions[typeIndex]&&typePositions[typeIndex].x)||(node.x+260)), y:Math.round((typePositions[typeIndex]&&typePositions[typeIndex].y)||node.y), className:'mindmap-scene-node-type', delay:(typeIndex%8)*0.05, typeItem:typeItems[typeIndex], linkFromX:baseX+offset.x, linkFromY:baseY+offset.y };
              layout.nodes.push(node);
              layout.links.push({ domKey:node.domKey, fromX:node.linkFromX, fromY:node.linkFromY, toX:node.x, toY:node.y, className:'mindmap-link-type' });
              minX=Math.min(minX,node.x-190);
              maxX=Math.max(maxX,node.x+190);
              minY=Math.min(minY,node.y-72);
              maxY=Math.max(maxY,node.y+72);
            }
          }
        }
        layout.sceneWidth=Math.max(2200,Math.round(maxX-minX+420));
        layout.sceneHeight=Math.max(1600,Math.round(maxY-minY+380));
        shiftX=Math.round(-minX+210);
        shiftY=Math.round(-minY+160);
        layout.rootX=Math.round(layout.rootX+shiftX);
        layout.rootY=Math.round(layout.rootY+shiftY);
        for(i=0;i<layout.hubs.length;i++){ layout.hubs[i].x+=shiftX; layout.hubs[i].y+=shiftY; }
        for(i=0;i<layout.nodes.length;i++){ layout.nodes[i].x+=shiftX; layout.nodes[i].y+=shiftY; if(layout.nodes[i].baseX!=null) layout.nodes[i].baseX+=shiftX; if(layout.nodes[i].baseY!=null) layout.nodes[i].baseY+=shiftY; if(layout.nodes[i].linkFromX!=null) layout.nodes[i].linkFromX+=shiftX; if(layout.nodes[i].linkFromY!=null) layout.nodes[i].linkFromY+=shiftY; }
        for(i=0;i<layout.links.length;i++){ layout.links[i].fromX+=shiftX; layout.links[i].toX+=shiftX; layout.links[i].fromY+=shiftY; layout.links[i].toY+=shiftY; }
        return layout;
      }
      function renderMindMapSceneHub(hub){
        return '<div class="mindmap-node mindmap-scene-hub '+esc(hub.className||'')+'" data-mindmap-node-key="'+esc(hub.domKey||hub.key||'')+'" style="'+mindMapNodePositionStyle(hub.x,hub.y)+'" aria-hidden="true"><span class="mindmap-node-title">'+esc(hub.label)+'</span>'+(hub.meta?'<span class="mindmap-node-meta">'+esc(hub.meta)+'</span>':'')+'</div>';
      }
      function renderMindMapGroupSceneNode(node){
        var active=s(mindMapState.focusGroupId)===s(node.id),expanded=!!mindMapExpandedGroups()[node.id];
        return '<article class="mindmap-node mindmap-scene-node '+esc(node.className||'')+(active?' is-active':'')+'" tabindex="0" role="button" aria-expanded="'+(expanded?'true':'false')+'" data-mindmap-node-kind="group" data-mindmap-node-id="'+esc(node.id)+'" data-mindmap-node-key="'+esc(node.domKey)+'" data-mindmap-draggable="1" style="'+mindMapNodePositionStyle(node.x,node.y,node.delay)+'"><div class="mindmap-node-head"><div class="mindmap-node-head-copy"><span class="mindmap-node-title">'+esc(node.label)+'</span>'+(node.meta?'<span class="mindmap-node-meta">'+esc(node.meta)+'</span>':'')+'</div><span class="mindmap-node-toggle">'+(expanded?'Collapse':'Expand')+'</span></div></article>';
      }
      function renderMindMapSceneNode(node){
        var active=mindMapState.selectedKind===node.kind&&mindMapState.selectedId===node.id;
        if(node.kind==='group') return renderMindMapGroupSceneNode(node);
        if(node.kind==='group-type'){
          active=s(mindMapState.focusGroupId)===s(node.groupId)&&s(mindMapState.focusTypeId)===s(node.id);
          return '<article class="mindmap-node mindmap-scene-node '+esc(node.className||'')+(active?' is-active':'')+'" tabindex="0" role="button" data-mindmap-node-kind="group-type" data-mindmap-group-id="'+esc(node.groupId)+'" data-mindmap-node-id="'+esc(node.id)+'" data-mindmap-node-key="'+esc(node.domKey)+'" style="'+mindMapNodePositionStyle(node.x,node.y,node.delay)+'"><div class="mindmap-node-head"><div class="mindmap-node-head-copy"><span class="mindmap-node-title">'+esc(node.label)+'</span>'+(node.meta?'<span class="mindmap-node-meta">'+esc(node.meta)+'</span>':'')+'</div><span class="mindmap-node-toggle">'+(active?'Active':'Open Notes')+'</span></div></article>';
        }
        if(node.kind==='group-reg'){
          active=s(mindMapState.focusGroupId)===s(node.groupId)&&s(mindMapState.focusTypeId)===s(node.typeId)&&s(mindMapState.focusRegId)===s(node.id);
          return '<article class="mindmap-node mindmap-scene-node '+esc(node.className||'')+(active?' is-active':'')+'" tabindex="0" role="button" data-mindmap-node-kind="group-reg" data-mindmap-group-id="'+esc(node.groupId)+'" data-mindmap-type-id="'+esc(node.typeId)+'" data-mindmap-node-id="'+esc(node.id)+'" data-mindmap-node-key="'+esc(node.domKey)+'" style="'+mindMapNodePositionStyle(node.x,node.y,node.delay)+'"><div class="mindmap-node-head"><div class="mindmap-node-head-copy"><span class="mindmap-node-title">'+esc(node.label)+'</span>'+(node.meta?'<span class="mindmap-node-meta">'+esc(node.meta)+'</span>':'')+'</div><span class="mindmap-node-toggle">'+(active?'Close':'Jobs')+'</span></div></article>';
        }
        if(node.kind==='group-chapter'){
          active=s(mindMapState.focusGroupId)===s(node.groupId)&&s(mindMapState.focusTypeId)===s(node.typeId)&&s(mindMapState.focusRegId)===s(node.regId)&&s(mindMapState.focusChapterId)===s(node.id);
          return '<article class="mindmap-node mindmap-scene-node '+esc(node.className||'')+(active?' is-active':'')+'" tabindex="0" role="button" data-mindmap-node-kind="group-chapter" data-mindmap-group-id="'+esc(node.groupId)+'" data-mindmap-type-id="'+esc(node.typeId)+'" data-mindmap-reg-id="'+esc(node.regId)+'" data-mindmap-node-id="'+esc(node.id)+'" data-mindmap-node-key="'+esc(node.domKey)+'" style="'+mindMapNodePositionStyle(node.x,node.y,node.delay)+'"><div class="mindmap-node-head"><div class="mindmap-node-head-copy"><span class="mindmap-node-title">'+esc(node.label)+'</span>'+(node.meta?'<span class="mindmap-node-meta">'+esc(node.meta)+'</span>':'')+'</div><span class="mindmap-node-toggle">'+(active?'Close':'Jobs')+'</span></div></article>';
        }
        if(node.rowId) return '<button class="mindmap-node mindmap-scene-node '+esc(node.className||'')+'" type="button" data-mindmap-job-row-id="'+esc(node.rowId)+'" data-mindmap-node-key="'+esc(node.domKey||'')+'" style="'+mindMapNodePositionStyle(node.x,node.y,node.delay)+'"><span class="mindmap-node-title">'+esc(node.label)+'</span>'+(node.meta?'<span class="mindmap-node-meta">'+esc(node.meta)+'</span>':'')+'</button>';
        return '<article class="mindmap-node mindmap-scene-node '+esc(node.className||'')+(active?' is-active':'')+'" tabindex="0" role="button" data-mindmap-node-kind="'+esc(node.kind)+'" data-mindmap-node-id="'+esc(node.id)+'" data-mindmap-node-key="'+esc(node.domKey||'')+'" style="'+mindMapNodePositionStyle(node.x,node.y,node.delay)+'"><span class="mindmap-node-title">'+esc(node.label)+'</span>'+(node.meta?'<span class="mindmap-node-meta">'+esc(node.meta)+'</span>':'')+'</article>';
      }
      function mindMapViewState(){
        if(!mindMapState.view) mindMapState.view={ x:0, y:0, scale:0.76, dragging:false, pointerId:null, startClientX:0, startClientY:0, startX:0, startY:0, needsCenter:true };
        return mindMapState.view;
      }
      function currentMindMapViewport(){ return mindMapCanvasEl&&mindMapCanvasEl.querySelector('[data-mindmap-viewport="1"]'); }
      function currentMindMapScene(){ return mindMapCanvasEl&&mindMapCanvasEl.querySelector('[data-mindmap-scene="1"]'); }
      function currentMindMapNodeEl(domKey){ return mindMapCanvasEl&&mindMapCanvasEl.querySelector('[data-mindmap-node-key="'+domKey+'"]'); }
      function currentMindMapLinkEl(domKey){ return mindMapCanvasEl&&mindMapCanvasEl.querySelector('[data-mindmap-link-key="'+domKey+'"]'); }
      function currentMindMapLayoutNode(domKey){
        var layout=mindMapState.layout||null,i;
        if(!layout||!layout.nodes) return null;
        for(i=0;i<layout.nodes.length;i++) if(layout.nodes[i].domKey===domKey) return layout.nodes[i];
        return null;
      }
      function clampMindMapScale(scale){ return Math.max(0.24,Math.min(1.7,Number(scale)||0.76)); }
      function fitMindMapScale(maxScale){
        var viewport=currentMindMapViewport(),scene=currentMindMapScene(),view=mindMapViewState(),availableWidth=0,availableHeight=0,nextScale=0;
        if(!viewport||!scene) return view.scale;
        availableWidth=Math.max(160,viewport.clientWidth-96);
        availableHeight=Math.max(160,viewport.clientHeight-86);
        nextScale=Math.min(availableWidth/Math.max(scene.offsetWidth,1),availableHeight/Math.max(scene.offsetHeight,1),Number(maxScale)||0.76);
        view.scale=clampMindMapScale(nextScale);
        updateMindMapZoomLabel();
        return view.scale;
      }
      function updateMindMapZoomLabel(){ var label=mindMapCanvasEl&&mindMapCanvasEl.querySelector('[data-mindmap-zoom-label="1"]'); if(label) label.textContent=Math.round(mindMapViewState().scale*100)+'%'; }
      function syncMindMapView(){
        var scene=currentMindMapScene(),viewport=currentMindMapViewport(),view=mindMapViewState();
        if(scene) scene.style.transform='translate('+view.x.toFixed(2)+'px,'+view.y.toFixed(2)+'px) scale('+view.scale.toFixed(3)+')';
        if(viewport) viewport.classList.toggle('is-dragging',!!view.dragging);
        updateMindMapZoomLabel();
      }
      function centerMindMapView(force){
        var viewport=currentMindMapViewport(),scene=currentMindMapScene(),view=mindMapViewState();
        if(!viewport||!scene) return;
        if(!force&&!view.needsCenter) return;
        var rootX=Number(scene.getAttribute('data-root-x'))||0,rootY=Number(scene.getAttribute('data-root-y'))||0;
        view.x=(viewport.clientWidth/2)-(rootX*view.scale);
        view.y=(viewport.clientHeight/2)-(rootY*view.scale);
        view.needsCenter=false;
        syncMindMapView();
      }
      function resetMindMapView(){
        var view=mindMapViewState();
        view.scale=0.76;
        view.dragging=false;
        view.pointerId=null;
        view.needsCenter=true;
        fitMindMapScale(0.76);
        centerMindMapView(true);
      }
      function zoomMindMapAt(factor, clientX, clientY){
        var viewport=currentMindMapViewport(),view=mindMapViewState();
        if(!viewport) return;
        var rect=viewport.getBoundingClientRect(),localX=clientX-rect.left,localY=clientY-rect.top,newScale=clampMindMapScale(view.scale*(Number(factor)||1));
        if(newScale===view.scale) return;
        var worldX=(localX-view.x)/view.scale,worldY=(localY-view.y)/view.scale;
        view.scale=newScale;
        view.x=localX-(worldX*newScale);
        view.y=localY-(worldY*newScale);
        view.needsCenter=false;
        syncMindMapView();
      }
      function zoomMindMapBy(factor){
        var viewport=currentMindMapViewport();
        if(!viewport) return;
        var rect=viewport.getBoundingClientRect();
        zoomMindMapAt(factor,rect.left+(rect.width/2),rect.top+(rect.height/2));
      }
      function mindMapTouchPointers(){
        if(!mindMapState.touchPointers) mindMapState.touchPointers=Object.create(null);
        return mindMapState.touchPointers;
      }
      function setMindMapTouchPointer(pointerId, clientX, clientY){
        mindMapTouchPointers()[String(pointerId)]={ id:pointerId, x:Number(clientX)||0, y:Number(clientY)||0 };
      }
      function clearMindMapTouchPointer(pointerId){
        var pointers=mindMapTouchPointers();
        delete pointers[String(pointerId)];
      }
      function clearMindMapTouchPointers(){ mindMapState.touchPointers=Object.create(null); }
      function mindMapTouchPointList(pointerIds){
        var pointers=mindMapTouchPointers(),list=[],i,key;
        if(pointerIds&&pointerIds.length){
          for(i=0;i<pointerIds.length;i++){
            key=String(pointerIds[i]);
            if(pointers[key]) list.push(pointers[key]);
          }
          return list;
        }
        for(key in pointers) list.push(pointers[key]);
        return list;
      }
      function mindMapPinchMetrics(points){
        if(!points||points.length<2) return null;
        var left=points[0],right=points[1],dx=(right.x-left.x),dy=(right.y-left.y);
        return { centerX:(left.x+right.x)/2, centerY:(left.y+right.y)/2, distance:Math.sqrt((dx*dx)+(dy*dy)) };
      }
      function beginMindMapPinch(){
        var viewport=currentMindMapViewport(),view=mindMapViewState(),points=mindMapTouchPointList(),metrics=mindMapPinchMetrics(points),rect,i,worldX,worldY;
        if(!viewport||!metrics||points.length<2||metrics.distance<12) return false;
        if(mindMapState.nodeDrag) endMindMapNodeDrag(mindMapState.nodeDrag.pointerId,{skipRender:true});
        if(view.dragging) endMindMapDrag(view.pointerId);
        rect=viewport.getBoundingClientRect();
        worldX=((metrics.centerX-rect.left)-view.x)/Math.max(view.scale,0.0001);
        worldY=((metrics.centerY-rect.top)-view.y)/Math.max(view.scale,0.0001);
        mindMapState.pinch={ pointerIds:[points[0].id,points[1].id], startDistance:metrics.distance, startScale:view.scale, worldX:worldX, worldY:worldY };
        view.needsCenter=false;
        if(viewport.setPointerCapture){
          for(i=0;i<mindMapState.pinch.pointerIds.length;i++){
            try { viewport.setPointerCapture(mindMapState.pinch.pointerIds[i]); } catch(err){}
          }
        }
        return true;
      }
      function updateMindMapPinch(){
        var viewport=currentMindMapViewport(),view=mindMapViewState(),pinch=mindMapState.pinch,points=mindMapTouchPointList(pinch&&pinch.pointerIds),metrics=mindMapPinchMetrics(points),rect,localX,localY,newScale;
        if(!viewport||!pinch||!metrics||points.length<2||metrics.distance<12) return false;
        rect=viewport.getBoundingClientRect();
        localX=metrics.centerX-rect.left;
        localY=metrics.centerY-rect.top;
        newScale=clampMindMapScale(pinch.startScale*(metrics.distance/Math.max(pinch.startDistance,1)));
        view.scale=newScale;
        view.x=localX-(pinch.worldX*newScale);
        view.y=localY-(pinch.worldY*newScale);
        view.needsCenter=false;
        syncMindMapView();
        return true;
      }
      function endMindMapPinch(pointerId){
        var viewport=currentMindMapViewport(),pinch=mindMapState.pinch,ids,i;
        if(!pinch) return false;
        ids=pinch.pointerIds||[];
        if(pointerId!=null&&ids.length&&ids.indexOf(pointerId)===-1&&mindMapTouchPointList(ids).length>=2) return false;
        if(viewport&&viewport.releasePointerCapture){
          for(i=0;i<ids.length;i++){
            try {
              if(!viewport.hasPointerCapture||viewport.hasPointerCapture(ids[i])) viewport.releasePointerCapture(ids[i]);
            } catch(err){}
          }
        }
        mindMapState.pinch=null;
        return true;
      }
      function syncMindMapDraggedNode(domKey){
        var node=currentMindMapLayoutNode(domKey),offset,nodeEl,linkEl;
        if(!node) return;
        offset=mindMapNodeOffset(domKey);
        node.x=node.baseX+offset.x;
        node.y=node.baseY+offset.y;
        nodeEl=currentMindMapNodeEl(domKey);
        if(nodeEl) nodeEl.style.cssText=mindMapNodePositionStyle(node.x,node.y,node.delay);
        linkEl=currentMindMapLinkEl(domKey);
        if(linkEl) linkEl.setAttribute('d',mindMapConnectorPath(node.linkFromX,node.linkFromY,node.x,node.y));
      }
      function beginMindMapNodeDrag(domKey, pointerId, clientX, clientY){
        var nodeEl=currentMindMapNodeEl(domKey),offset=mindMapNodeOffset(domKey);
        mindMapState.nodeDrag={ domKey:domKey, pointerId:pointerId, startClientX:Number(clientX)||0, startClientY:Number(clientY)||0, startX:offset.x, startY:offset.y, moved:false };
        if(nodeEl&&nodeEl.setPointerCapture){
          try { nodeEl.setPointerCapture(pointerId); } catch(err){}
        }
      }
      function updateMindMapNodeDrag(clientX, clientY){
        var drag=mindMapState.nodeDrag,offset,deltaX,deltaY;
        if(!drag) return;
        deltaX=(Number(clientX)||0)-drag.startClientX;
        deltaY=(Number(clientY)||0)-drag.startClientY;
        if(!drag.moved&&Math.abs(deltaX)<4&&Math.abs(deltaY)<4) return;
        drag.moved=true;
        offset=mindMapNodeOffset(drag.domKey);
        offset.x=drag.startX+deltaX;
        offset.y=drag.startY+deltaY;
        syncMindMapDraggedNode(drag.domKey);
      }
      function endMindMapNodeDrag(pointerId, options){
        var drag=mindMapState.nodeDrag,nodeEl;
        options=options||{};
        if(!drag||pointerId!=null&&drag.pointerId!==pointerId) return false;
        nodeEl=currentMindMapNodeEl(drag.domKey);
        if(nodeEl&&nodeEl.releasePointerCapture){
          try {
            if(!nodeEl.hasPointerCapture||nodeEl.hasPointerCapture(drag.pointerId)) nodeEl.releasePointerCapture(drag.pointerId);
          } catch(err){}
        }
        mindMapState.nodeDrag=null;
        if(drag.moved) mindMapState.suppressClickUntil=Date.now()+220;
        if(drag.moved&&!options.skipRender) renderMindMapModal();
        return !!drag.moved;
      }
      function canStartMindMapDrag(target){
        return !(target&&target.closest&&target.closest('button,[role="button"],[data-mindmap-node-kind]'));
      }
      function beginMindMapDrag(pointerId, clientX, clientY){
        var viewport=currentMindMapViewport(),view=mindMapViewState();
        if(!viewport) return false;
        view.dragging=true;
        view.pointerId=pointerId;
        view.startClientX=Number(clientX)||0;
        view.startClientY=Number(clientY)||0;
        view.startX=view.x;
        view.startY=view.y;
        view.needsCenter=false;
        if(viewport.setPointerCapture){
          try { viewport.setPointerCapture(pointerId); } catch(err){}
        }
        syncMindMapView();
        return true;
      }
      function updateMindMapDrag(clientX, clientY){
        var view=mindMapViewState();
        if(!view.dragging) return;
        view.x=view.startX+((Number(clientX)||0)-view.startClientX);
        view.y=view.startY+((Number(clientY)||0)-view.startClientY);
        syncMindMapView();
      }
      function endMindMapDrag(pointerId){
        var viewport=currentMindMapViewport(),view=mindMapViewState(),activePointer=view.pointerId;
        if(pointerId!=null&&activePointer!=null&&pointerId!==activePointer) return;
        view.dragging=false;
        view.pointerId=null;
        if(viewport&&viewport.releasePointerCapture&&activePointer!=null){
          try {
            if(!viewport.hasPointerCapture||viewport.hasPointerCapture(activePointer)) viewport.releasePointerCapture(activePointer);
          } catch(err){}
        }
        syncMindMapView();
      }
      function handleMindMapWheelEvent(ev){
        var viewport=currentMindMapViewport(),delta,factor;
        if(!viewport||!viewport.contains(ev.target)||!mindMapModal||mindMapModal.className.indexOf('open')===-1) return;
        ev.preventDefault();
        delta=Number(ev.deltaY)||0;
        if(ev.deltaMode===1) delta*=18;
        else if(ev.deltaMode===2) delta*=Math.max(viewport.clientHeight,1);
        factor=Math.exp((-1*delta)/640);
        zoomMindMapAt(factor,ev.clientX,ev.clientY);
      }
      function renderMindMapToolbar(){
        return '<div class="mindmap-toolbar"><button class="mindmap-tool-btn" type="button" data-mindmap-zoom="out" aria-label="Zoom out">-</button><span class="mindmap-zoom-label" data-mindmap-zoom-label="1">'+Math.round(mindMapViewState().scale*100)+'%</span><button class="mindmap-tool-btn" type="button" data-mindmap-zoom="in" aria-label="Zoom in">+</button><button class="mindmap-tool-btn mindmap-tool-btn-reset" type="button" data-mindmap-zoom="reset">Center</button><button class="mindmap-tool-btn mindmap-tool-btn-details" type="button" data-mindmap-detail-toggle="1">'+(mindMapState.detailClosed?'Show Notes':'Hide Notes')+'</button></div>';
      }
      function buildMindMapSummary(){
        var list=nonEmptyRows(rows).slice(),chapterLookup=mindMapChapterLookup(),groupsMap=Object.create(null),typesMap=Object.create(null),chaptersMap=Object.create(null),jobs=[],i;
        function ensureGroupBucket(label){
          var normalized=normalizeMindMapGroupLabel(label);
          if(!groupsMap[normalized]) groupsMap[normalized]={ id:normalized, label:normalized, entryCount:0, jobs:[], _types:Object.create(null), _chapters:Object.create(null), _chapterItems:Object.create(null), _regs:Object.create(null) };
          return groupsMap[normalized];
        }
        for(i=0;i<list.length;i++){
          var row=list[i]||{},type=rowAircraftTypeLabel(row),group=rowAircraftGroupLabel(row),groupBucket=ensureGroupBucket(group),reg=s(row['A/C Reg']).toUpperCase(),chapterCode=s(row['Chapter']),chapterInfo=chapterLookup[chapterCode]||null,chapterDesc=s(row['Chapter Description'])||s(chapterInfo&&chapterInfo.description),chapterId=chapterCode||BLANK_CHAPTER_FILTER,chapterLabel=mindMapChapterLabel(chapterCode,chapterDesc),task=mainPageTaskText(row),jobNo=s(row['Job No']),job={ id:String(row.__rowId), rowId:String(row.__rowId), label:jobNo||mindMapClip(task,40)||('Entry '+String(Number(row.__rowId)+1)), jobNo:jobNo, task:task, date:formatDateDisplay(row['Date']), sortDate:parseDate(row['Date']), reg:reg, type:type, group:group, chapter:chapterCode, chapterDesc:chapterDesc, chapterLabel:chapterLabel };
          jobs.push(job);
          groupBucket.entryCount++;
          groupBucket.jobs.push(job);
          groupBucket._types[type]=type;
          groupBucket._chapters[chapterLabel]=chapterLabel;
          groupBucket._chapterItems[chapterId]={ id:chapterId, label:chapterLabel };
          if(reg) groupBucket._regs[reg]=reg;
          if(!typesMap[type]) typesMap[type]={ id:type, label:type, entryCount:0, jobs:[], _groups:Object.create(null), _chapters:Object.create(null), _regs:Object.create(null) };
          typesMap[type].entryCount++;
          typesMap[type].jobs.push(job);
          typesMap[type]._groups[group]=group;
          typesMap[type]._chapters[chapterLabel]=chapterLabel;
          if(reg) typesMap[type]._regs[reg]=reg;
          if(!chaptersMap[chapterId]) chaptersMap[chapterId]={ id:chapterId, chapter:chapterCode||BLANK_CHAPTER_FILTER, description:chapterDesc, label:chapterLabel, entryCount:0, jobs:[], _groups:Object.create(null), _types:Object.create(null), _regs:Object.create(null) };
          chaptersMap[chapterId].entryCount++;
          chaptersMap[chapterId].jobs.push(job);
          chaptersMap[chapterId]._groups[group]=group;
          chaptersMap[chapterId]._types[type]=type;
          if(reg) chaptersMap[chapterId]._regs[reg]=reg;
          if(!chaptersMap[chapterId].description&&chapterDesc) chaptersMap[chapterId].description=chapterDesc;
        }
        ensureGroupBucket('Ungrouped');
        var groups=[],groupKey;
        for(groupKey in groupsMap) if(Object.prototype.hasOwnProperty.call(groupsMap,groupKey)){
          var groupItem=groupsMap[groupKey];
          groupItem.label=normalizeMindMapGroupLabel(groupItem.label);
          groupItem.id=groupItem.label;
          groupItem.aircraftTypes=Object.keys(groupItem._types).sort(mindMapTextSort);
          groupItem.chapters=Object.keys(groupItem._chapters).sort(mindMapChapterTextSort);
          groupItem.chapterItems=Object.keys(groupItem._chapterItems).map(function(itemKey){ return groupItem._chapterItems[itemKey]; }).sort(function(a,b){ return mindMapChapterTextSort(a&&a.label,b&&b.label); });
          groupItem.registrations=Object.keys(groupItem._regs).sort(mindMapTextSort);
          groupItem.jobs.sort(mindMapJobSort);
          groupItem.typeItems=buildMindMapGroupTypeItems(groupItem);
          delete groupItem._types;
          delete groupItem._chapters;
          delete groupItem._chapterItems;
          delete groupItem._regs;
          groups.push(groupItem);
        }
        groups.sort(mindMapGroupSort);
        var types=[],typeKey;
        for(typeKey in typesMap) if(Object.prototype.hasOwnProperty.call(typesMap,typeKey)){
          var typeItem=typesMap[typeKey];
          typeItem.aircraftGroups=Object.keys(typeItem._groups).sort(mindMapTextSort);
          typeItem.chapters=Object.keys(typeItem._chapters).sort(mindMapChapterTextSort);
          typeItem.registrations=Object.keys(typeItem._regs).sort(mindMapTextSort);
          typeItem.jobs.sort(mindMapJobSort);
          delete typeItem._groups;
          delete typeItem._chapters;
          delete typeItem._regs;
          types.push(typeItem);
        }
        types.sort(function(a,b){ return mindMapTextSort(a.label,b.label); });
        var chapters=[],chapterKey;
        for(chapterKey in chaptersMap) if(Object.prototype.hasOwnProperty.call(chaptersMap,chapterKey)){
          var chapterItem=chaptersMap[chapterKey];
          chapterItem.aircraftGroups=Object.keys(chapterItem._groups).sort(mindMapTextSort);
          chapterItem.aircraftTypes=Object.keys(chapterItem._types).sort(mindMapTextSort);
          chapterItem.registrations=Object.keys(chapterItem._regs).sort(mindMapTextSort);
          chapterItem.jobs.sort(mindMapJobSort);
          delete chapterItem._groups;
          delete chapterItem._types;
          delete chapterItem._regs;
          chapters.push(chapterItem);
        }
        chapters.sort(mindMapChapterSort);
        jobs.sort(mindMapJobSort);
        return { entryCount:list.length, groupCount:groups.length, typeCount:types.length, chapterCount:chapters.length, jobCount:jobs.length, groups:groups, types:types, chapters:chapters, jobs:jobs };
      }
      function resolveMindMapSelection(summary){
        var kind=s(mindMapState.selectedKind)||'root',id=s(mindMapState.selectedId)||'overview',item=null,i,list=null,group=null,typeItem=null,regItem=null;
        if(kind==='group') list=summary.groups;
        else if(kind==='type') list=summary.types;
        else if(kind==='chapter') list=summary.chapters;
        else if(kind==='group-type'){
          group=findMindMapGroup(summary,mindMapState.focusGroupId);
          list=group&&group.typeItems;
        } else if(kind==='group-reg'){
          group=findMindMapGroup(summary,mindMapState.focusGroupId);
          typeItem=findMindMapGroupType(group,mindMapState.focusTypeId);
          list=typeItem&&(typeItem.chapterItems||typeItem.regItems);
        } else if(kind==='group-chapter'){
          group=findMindMapGroup(summary,mindMapState.focusGroupId);
          typeItem=findMindMapGroupType(group,mindMapState.focusTypeId);
          regItem=findMindMapTypeChapter(typeItem,mindMapState.focusRegId);
          list=regItem&&regItem.chapterItems;
        }
        if(list){ for(i=0;i<list.length;i++){ if(s(list[i].id)===id){ item=list[i]; break; } } }
        if(!item){ kind='root'; id='overview'; item=summary; }
        mindMapState.selectedKind=kind;
        mindMapState.selectedId=id;
        return { kind:kind, id:id, item:item, group:group, type:typeItem, reg:regItem };
      }
      function renderMindMapCanvas(summary){
        var layout=buildMindMapLayout(summary),links=[],hubs=[],nodes=[],i;
        mindMapState.layout=layout;
        for(i=0;i<layout.links.length;i++) links.push('<path class="mindmap-link '+esc(layout.links[i].className||'')+'" data-mindmap-link-key="'+esc(layout.links[i].domKey||'')+'" d="'+esc(mindMapConnectorPath(layout.links[i].fromX,layout.links[i].fromY,layout.links[i].toX,layout.links[i].toY))+'"></path>');
        for(i=0;i<layout.hubs.length;i++) hubs.push(renderMindMapSceneHub(layout.hubs[i]));
        for(i=0;i<layout.nodes.length;i++) nodes.push(renderMindMapSceneNode(layout.nodes[i]));
        return '<div class="mindmap-stage">'+
          renderMindMapToolbar()+
          '<div class="mindmap-viewport" data-mindmap-viewport="1">'+
            '<div class="mindmap-scene" data-mindmap-scene="1" data-root-x="'+esc(layout.rootX)+'" data-root-y="'+esc(layout.rootY)+'" style="width:'+layout.sceneWidth+'px;height:'+layout.sceneHeight+'px;">'+
              '<svg class="mindmap-links" viewBox="0 0 '+layout.sceneWidth+' '+layout.sceneHeight+'" width="'+layout.sceneWidth+'" height="'+layout.sceneHeight+'" aria-hidden="true">'+links.join('')+'</svg>'+
              hubs.join('')+
              nodes.join('')+
            '</div>'+
          '</div>'+
          '<div class="mindmap-scene-hint">Drag the canvas to move around, drag a group node to reposition it, and keep several groups open at once. Pick an A/C type to load its chapters and jobs in the side notes panel.</div>'+
        '</div>';
      }
      function renderMindMapOverviewDetail(summary){
        if(!summary.entryCount) return '<div class="mindmap-empty">No logbook entries are loaded yet. Import or add CAP 741 rows and the mind map will build itself from the live data.</div>';
        var chapterCards=[],i;
        for(i=0;i<summary.chapters.length;i++){
          var chapter=summary.chapters[i],includes=chapter.aircraftTypes.length?('Includes: '+mindMapClip(chapter.aircraftTypes.join(', '),110)):'No aircraft types recorded yet.',copy=(chapter.description?chapter.description+'. ':'')+includes;
          chapterCards.push(mindMapDetailCardHtml('chapter',chapter.id,chapter.label,chapter.entryCount+' entries | '+chapter.aircraftGroups.length+' groups',copy));
        }
        return '<p class="mindmap-detail-head-label">Overview</p><h3 class="mindmap-detail-title">Logbook Snapshot</h3><p class="mindmap-detail-copy">Live totals.</p><div class="mindmap-metrics">'+mindMapMetricHtml('Entries',summary.entryCount)+mindMapMetricHtml('Chapters',summary.chapterCount)+mindMapMetricHtml('Groups',summary.groupCount)+mindMapMetricHtml('Types',summary.typeCount)+'</div><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">What Each Chapter Includes</h4><div class="mindmap-detail-cards">'+chapterCards.join('')+'</div></section>';
      }
      function renderMindMapGroupDetail(group){
        return '<p class="mindmap-detail-head-label">Aircraft Group</p><h3 class="mindmap-detail-title">'+esc(group.label)+'</h3><p class="mindmap-detail-copy">Group summary.</p><div class="mindmap-metrics">'+mindMapMetricHtml('Entries',group.entryCount)+mindMapMetricHtml('A/C Types',group.aircraftTypes.length)+mindMapMetricHtml('Regs',group.registrations.length)+mindMapMetricHtml('Chapters',group.chapters.length)+'</div><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Aircraft Types</h4>'+mindMapPillListHtml(group.aircraftTypes,'No aircraft types recorded in this group.')+'</section><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Registrations</h4>'+mindMapPillListHtml(group.registrations,'No registrations recorded in this group.')+'</section><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Chapters</h4>'+mindMapPillListHtml(group.chapters,'No chapters recorded in this group.')+'</section>';
      }
      function renderMindMapTypeDetail(type){
        return '<p class="mindmap-detail-head-label">Aircraft Type</p><h3 class="mindmap-detail-title">'+esc(type.label)+'</h3><p class="mindmap-detail-copy">Type summary.</p><div class="mindmap-metrics">'+mindMapMetricHtml('Entries',type.entryCount)+mindMapMetricHtml('Groups',type.aircraftGroups.length)+mindMapMetricHtml('Regs',type.registrations.length)+mindMapMetricHtml('Chapters',type.chapters.length)+'</div><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Aircraft Groups</h4>'+mindMapPillListHtml(type.aircraftGroups,'No aircraft groups recorded for this type.')+'</section><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Registrations</h4>'+mindMapPillListHtml(type.registrations,'No registrations recorded for this type.')+'</section><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Chapters</h4>'+mindMapPillListHtml(type.chapters,'No chapters recorded for this type.')+'</section>';
      }
      function renderMindMapGroupTypeDetail(type, group){
        return '<p class="mindmap-detail-head-label">Group Type</p><h3 class="mindmap-detail-title">'+esc(type.label)+'</h3><p class="mindmap-detail-copy">Pick a chapter below.</p><div class="mindmap-metrics">'+mindMapMetricHtml('Entries',type.entryCount)+mindMapMetricHtml('Chapters',(type.chapterItems||type.regItems||[]).length)+mindMapMetricHtml('Regs',type.registrations.length)+mindMapMetricHtml('Jobs',type.jobs.length)+'</div><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Chapter List</h4>'+mindMapChapterAccordionListHtml(type.chapterItems||type.regItems,'No chapters recorded for this type in the selected group.')+'</section><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Registrations</h4>'+mindMapPillListHtml(type.registrations,'No registrations recorded for this type in the selected group.')+'</section>';
      }
      function renderMindMapGroupRegDetail(regItem, type, group){
        var jobLinks=[],i;
        for(i=0;i<regItem.jobs.length;i++) jobLinks.push(mindMapJobLinkHtml(regItem.jobs[i]));
        return '<p class="mindmap-detail-head-label">Chapter</p><h3 class="mindmap-detail-title">'+esc(regItem.label)+'</h3><p class="mindmap-detail-copy">Chapter summary.</p><div class="mindmap-metrics">'+mindMapMetricHtml('Entries',regItem.entryCount)+mindMapMetricHtml('Jobs',regItem.jobs.length)+mindMapMetricHtml('Regs',(regItem.registrations||[]).length)+mindMapMetricHtml('Type',type&&type.label||'')+'</div><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Registrations</h4>'+mindMapPillListHtml(regItem.registrations,'No registrations recorded in this chapter.')+'</section><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Jobs In This Chapter</h4><div class="mindmap-job-links">'+(jobLinks.join('')||'<div class="mindmap-empty-mini">No jobs recorded in this chapter.</div>')+'</div></section>';
      }
      function renderMindMapGroupChapterDetail(chapter, regItem, type, group){
        var jobLinks=[],i;
        for(i=0;i<chapter.jobs.length;i++) jobLinks.push(mindMapJobLinkHtml(chapter.jobs[i]));
        return '<p class="mindmap-detail-head-label">Chapter</p><h3 class="mindmap-detail-title">'+esc(chapter.label)+'</h3><p class="mindmap-detail-copy">This chapter is under '+esc(regItem&&regItem.label||'the selected registration')+' for '+esc(type&&type.label||'the selected type')+' in '+esc(group&&group.label||'the selected group')+'.</p><div class="mindmap-metrics">'+mindMapMetricHtml('Entries',chapter.entryCount)+mindMapMetricHtml('Jobs',chapter.jobs.length)+mindMapMetricHtml('Registration',regItem&&regItem.label||'')+mindMapMetricHtml('Group',group&&group.label||'')+'</div><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Jobs In This Chapter</h4><div class="mindmap-job-links">'+(jobLinks.join('')||'<div class="mindmap-empty-mini">No jobs recorded in this chapter.</div>')+'</div></section>';
      }
      function renderMindMapChapterDetail(chapter){
        var jobLinks=[],i;
        for(i=0;i<chapter.jobs.length;i++) jobLinks.push(mindMapJobLinkHtml(chapter.jobs[i]));
        return '<p class="mindmap-detail-head-label">Chapter</p><h3 class="mindmap-detail-title">'+esc(chapter.label)+'</h3><p class="mindmap-detail-copy">'+esc(chapter.description||'Chapter usage.')+'</p><div class="mindmap-metrics">'+mindMapMetricHtml('Entries',chapter.entryCount)+mindMapMetricHtml('Groups',chapter.aircraftGroups.length)+mindMapMetricHtml('Types',chapter.aircraftTypes.length)+mindMapMetricHtml('Jobs',chapter.jobs.length)+'</div><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Aircraft Groups</h4>'+mindMapPillListHtml(chapter.aircraftGroups,'No aircraft groups recorded in this chapter.')+'</section><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Aircraft Types</h4>'+mindMapPillListHtml(chapter.aircraftTypes,'No aircraft types recorded in this chapter.')+'</section><section class="mindmap-detail-section"><h4 class="mindmap-detail-section-title">Jobs In This Chapter</h4><div class="mindmap-job-links">'+(jobLinks.join('')||'<div class="mindmap-empty-mini">No jobs recorded in this chapter.</div>')+'</div></section>';
      }
      function renderMindMapDetail(summary){
        var selection=resolveMindMapSelection(summary);
        if(selection.kind==='group') return renderMindMapGroupDetail(selection.item);
        if(selection.kind==='group-type') return renderMindMapGroupTypeDetail(selection.item,selection.group);
        if(selection.kind==='group-reg') return renderMindMapGroupRegDetail(selection.item,selection.type,selection.group);
        if(selection.kind==='group-chapter') return renderMindMapGroupChapterDetail(selection.item,selection.reg,selection.type,selection.group);
        if(selection.kind==='type') return renderMindMapTypeDetail(selection.item);
        if(selection.kind==='chapter') return renderMindMapChapterDetail(selection.item);
        return renderMindMapOverviewDetail(summary);
      }
      function renderMindMapDetailPanel(summary){
        return '<div class="mindmap-detail-bar"><p class="mindmap-detail-bar-label">Side Notes</p><button class="mindmap-detail-close" type="button" data-mindmap-detail-toggle="1">Hide</button></div>'+renderMindMapDetail(summary);
      }
      function renderMindMapModal(){
        var summary=mindMapState.summary||buildMindMapSummary();
        mindMapState.summary=summary;
        if(mindMapCanvasEl){ mindMapCanvasEl.innerHTML=renderMindMapCanvas(summary); syncMindMapView(); }
        if(mindMapShellEl) mindMapShellEl.classList.toggle('is-detail-closed',!!mindMapState.detailClosed);
        if(mindMapDetailEl){
          mindMapDetailEl.className='mindmap-detail'+(mindMapState.detailClosed?' is-closed':'');
          mindMapDetailEl.hidden=!!mindMapState.detailClosed;
          mindMapDetailEl.innerHTML=mindMapState.detailClosed?'':renderMindMapDetailPanel(summary);
        }
      }
      function setMindMapDetailClosed(closed){
        mindMapState.detailClosed=!!closed;
        renderMindMapModal();
      }
      function openMindMapFeature(){
        var view=mindMapViewState();
        view.scale=0.76;
        view.x=0;
        view.y=0;
        view.dragging=false;
        view.pointerId=null;
        view.needsCenter=true;
        mindMapState.expandedGroups=Object.create(null);
        mindMapState.nodeOffsets=Object.create(null);
        mindMapState.nodeDrag=null;
        mindMapState.pinch=null;
        clearMindMapTouchPointers();
        mindMapState.detailClosed=false;
        mindMapState.suppressClickUntil=0;
        mindMapState.focusGroupId='';
        mindMapState.focusTypeId='';
        mindMapState.focusRegId='';
        mindMapState.focusChapterId='';
        mindMapState.summary=buildMindMapSummary();
        mindMapState.selectedKind='root';
        mindMapState.selectedId='overview';
        if(mindMapModal) mindMapModal.className='modal-backdrop open';
        renderMindMapModal();
        setTimeout(function(){
          fitMindMapScale(0.76);
          centerMindMapView(true);
          var firstNode=mindMapCanvasEl&&mindMapCanvasEl.querySelector('[data-mindmap-node-kind="group"]');
          if(firstNode&&typeof firstNode.focus==='function') firstNode.focus();
        },0);
      }
      function closeMindMapModal(){ var view=mindMapViewState(); endMindMapPinch(); view.dragging=false; view.pointerId=null; mindMapState.nodeDrag=null; mindMapState.pinch=null; clearMindMapTouchPointers(); if(mindMapModal) mindMapModal.className='modal-backdrop'; syncMindMapView(); }
      function selectMindMapNode(kind, id, groupId, typeId, regId){
        var nextKind=s(kind)||'root',nextId=s(id)||'overview',nextGroup=s(groupId),sameGroup,nextType,nextReg,sameType,sameReg,sameChapter,expandedGroups=mindMapExpandedGroups();
        if(nextKind==='group'){
          sameGroup=!!expandedGroups[nextId];
          if(sameGroup){
            delete expandedGroups[nextId];
            if(s(mindMapState.focusGroupId)===nextId){
              mindMapState.focusGroupId='';
              mindMapState.focusTypeId='';
              mindMapState.focusRegId='';
              mindMapState.focusChapterId='';
              mindMapState.selectedKind='root';
              mindMapState.selectedId='overview';
            }
          } else {
            expandedGroups[nextId]=1;
            mindMapState.focusGroupId=nextId;
            mindMapState.focusTypeId='';
            mindMapState.focusRegId='';
            mindMapState.focusChapterId='';
            mindMapState.selectedKind='group';
            mindMapState.selectedId=nextId;
          }
          renderMindMapModal();
          return;
        }
        if(nextKind==='group-type'){
          nextType=nextId;
          if(!nextGroup) nextGroup=s(mindMapState.focusGroupId);
          expandedGroups[nextGroup]=1;
          sameType=s(mindMapState.focusGroupId)===nextGroup&&s(mindMapState.focusTypeId)===nextType&&s(mindMapState.focusRegId)==='';
          mindMapState.focusGroupId=nextGroup;
          mindMapState.focusTypeId=sameType?'':nextType;
          mindMapState.focusRegId='';
          mindMapState.focusChapterId='';
          mindMapState.selectedKind=sameType?'group':'group-type';
          mindMapState.selectedId=sameType?nextGroup:nextType;
          renderMindMapModal();
          return;
        }
        if(nextKind==='group-reg'){
          nextReg=nextId;
          if(!nextGroup) nextGroup=s(mindMapState.focusGroupId);
          if(!typeId) typeId=s(mindMapState.focusTypeId);
          sameReg=s(mindMapState.focusGroupId)===nextGroup&&s(mindMapState.focusTypeId)===s(typeId)&&s(mindMapState.focusRegId)===nextReg&&s(mindMapState.focusChapterId)==='';
          mindMapState.focusGroupId=nextGroup;
          mindMapState.focusTypeId=s(typeId);
          mindMapState.focusRegId=sameReg?'':nextReg;
          mindMapState.focusChapterId='';
          mindMapState.selectedKind=sameReg?'group-type':'group-reg';
          mindMapState.selectedId=sameReg?s(typeId):nextReg;
          renderMindMapModal();
          return;
        }
        if(nextKind==='group-chapter'){
          if(!nextGroup) nextGroup=s(mindMapState.focusGroupId);
          if(!typeId) typeId=s(mindMapState.focusTypeId);
          if(!regId) regId=s(mindMapState.focusRegId);
          sameChapter=s(mindMapState.focusGroupId)===nextGroup&&s(mindMapState.focusTypeId)===s(typeId)&&s(mindMapState.focusRegId)===s(regId)&&s(mindMapState.focusChapterId)===nextId;
          mindMapState.focusGroupId=nextGroup;
          mindMapState.focusTypeId=s(typeId);
          mindMapState.focusRegId=s(regId);
          mindMapState.focusChapterId=sameChapter?'':nextId;
          mindMapState.selectedKind=sameChapter?'group-reg':'group-chapter';
          mindMapState.selectedId=sameChapter?s(regId):nextId;
          renderMindMapModal();
          return;
        }
        if(nextKind!=='group') mindMapState.focusGroupId='';
        mindMapState.focusTypeId='';
        mindMapState.focusRegId='';
        mindMapState.focusChapterId='';
        mindMapState.selectedKind=nextKind;
        mindMapState.selectedId=nextId;
        renderMindMapModal();
      }
      function clearMindMapRowHighlight(){
        if(!pagesEl) return;
        var highlighted=pagesEl.querySelectorAll('.mindmap-row-target');
        for(var i=0;i<highlighted.length;i++) highlighted[i].classList.remove('mindmap-row-target');
      }
      async function jumpToMindMapRow(rowId){
        rowId=s(rowId);
        if(!rowId||!rowById(rowId)){ fail('Could not find that logbook job.'); return; }
        closeMindMapModal();
        var rowSelector='[data-row-key="row-'+rowId+'"]',rowEl=pagesEl&&pagesEl.querySelector(rowSelector);
        if(!rowEl&&(hasActiveFilters()||hasActiveSearch())){
          activeFilters=emptyFilterState();
          draftFilters=emptyFilterState();
          applySearchQuery('');
          renderAll();
          await nextPaint();
          note('Filters and search were cleared so the selected job could be shown.');
          rowEl=pagesEl&&pagesEl.querySelector(rowSelector);
        }
        if(!rowEl){
          renderAll();
          await nextPaint();
          rowEl=pagesEl&&pagesEl.querySelector(rowSelector);
        }
        if(!rowEl){ fail('Could not jump to the selected job in the logbook.'); return; }
        clearMindMapRowHighlight();
        rowEl.classList.add('mindmap-row-target');
        if(typeof rowEl.scrollIntoView==='function') rowEl.scrollIntoView({behavior:'smooth',block:'center'});
        clearTimeout(mindMapRowHighlightTimer);
        mindMapRowHighlightTimer=setTimeout(function(){ rowEl.classList.remove('mindmap-row-target'); },2200);
      }
      function handleMindMapInteraction(target){
        if((Number(mindMapState.suppressClickUntil)||0)>Date.now()) return true;
        var zoomBtn=target&&target.closest&&target.closest('[data-mindmap-zoom]');
        if(zoomBtn){
          var action=zoomBtn.getAttribute('data-mindmap-zoom');
          if(action==='in') zoomMindMapBy(1.12);
          else if(action==='out') zoomMindMapBy(1/1.12);
          else resetMindMapView();
          return true;
        }
        var detailToggleBtn=target&&target.closest&&target.closest('[data-mindmap-detail-toggle]');
        if(detailToggleBtn){ setMindMapDetailClosed(!mindMapState.detailClosed); return true; }
        var jobBtn=target&&target.closest&&target.closest('[data-mindmap-job-row-id]');
        if(jobBtn){ jumpToMindMapRow(jobBtn.getAttribute('data-mindmap-job-row-id')); return true; }
        var nodeBtn=target&&target.closest&&target.closest('[data-mindmap-node-kind]');
        if(nodeBtn){ selectMindMapNode(nodeBtn.getAttribute('data-mindmap-node-kind'),nodeBtn.getAttribute('data-mindmap-node-id'),nodeBtn.getAttribute('data-mindmap-group-id'),nodeBtn.getAttribute('data-mindmap-type-id'),nodeBtn.getAttribute('data-mindmap-reg-id')); return true; }
        return false;
      }
      function readFilterForm(){ commitPendingDraftInputs(); return cloneFilterState(draftFilters); }
      function clearFilters(){ activeFilters=emptyFilterState(); draftFilters=emptyFilterState(); resetDraftFilters(); renderAll(); }
      function applySearchQuery(value){
        searchQuery=String(value==null?'':value);
        normalizedSearchQuery=normalizeSearchText(searchQuery);
        syncSearchUi();
      }
      function scheduleSearchRender(delay){
        clearTimeout(searchRenderTimer);
        searchRenderTimer=setTimeout(function(){ renderAll(); },typeof delay==='number'?delay:120);
      }
      function clearSearch(){ applySearchQuery(''); clearTimeout(searchRenderTimer); renderAll(); }
      function cloneMeasurementState(source){
        source=source||OTHER_LAYOUT_DEFAULTS;
        return {
          top:Number(source.top)||0,
          headerHeight:Number(source.headerHeight)||0,
          left:Number(source.left)||0,
          dateStart:Number(source.dateStart)||0,
          regStart:Number(source.regStart)||0,
          jobStart:Number(source.jobStart)||0,
          taskStart:Number(source.taskStart)||0,
          superStart:Number(source.superStart)||0,
          end:Number(source.end)||0,
          rowHeight:Number(source.rowHeight)||0,
          textTop:Number(source.textTop)||0,
          textLeft:Number(source.textLeft)||0
        };
      }
      function writeMeasurementCssVars(prefix, measurements, target){
        target=target||document.documentElement;
        measurements=cloneMeasurementState(measurements);
        measurements.dateStart=measurements.left;
        function setVar(name, value, unit){ target.style.setProperty(prefix+String(name).replace(/^--/,'-'),String(value)+(unit||'')); }
        setVar('--top', measurements.top, 'mm');
        setVar('--header-height', measurements.headerHeight, 'mm');
        setVar('--left', measurements.left, 'mm');
        setVar('--frame-width', Math.max(1,measurements.end-measurements.left), 'mm');
        setVar('--row-height', measurements.rowHeight, 'mm');
        setVar('--text-top', measurements.textTop, 'mm');
        setVar('--text-left', measurements.textLeft, 'mm');
        setVar('--date-width', Math.max(1,measurements.regStart-measurements.dateStart), 'mm');
        setVar('--reg-width', Math.max(1,measurements.jobStart-measurements.regStart), 'mm');
        setVar('--job-width', Math.max(1,measurements.taskStart-measurements.jobStart), 'mm');
        setVar('--task-width', Math.max(1,measurements.superStart-measurements.taskStart), 'mm');
        setVar('--sup-width', Math.max(1,measurements.end-measurements.superStart), 'mm');
      }
      function pagePxToMm(pageRect, distancePx){
        return Number(((Math.max(0,distancePx)*210)/Math.max(pageRect&&pageRect.width||0,1)).toFixed(2));
      }
      function currentOverlayPrintMeasurements(page){
        if(!page||!page.getBoundingClientRect) return null;
        var pageRect=page.getBoundingClientRect(),frame=page.querySelector('.frame'),table=page.querySelector('table.sheet');
        if(!pageRect.width||!frame||!table||!table.tHead||!table.tHead.rows||!table.tHead.rows.length) return null;
        var headerRow=table.tHead.rows[0],firstRow=table.tBodies&&table.tBodies[0]?table.tBodies[0].querySelector('tr.slot'):null,cells=headerRow.cells||[];
        if(cells.length<5) return null;
        var frameRect=frame.getBoundingClientRect(),headerRect=headerRow.getBoundingClientRect(),firstRowRect=firstRow?firstRow.getBoundingClientRect():null;
        return {
          top:pagePxToMm(pageRect,frameRect.top-pageRect.top),
          headerHeight:pagePxToMm(pageRect,headerRect.height),
          left:pagePxToMm(pageRect,frameRect.left-pageRect.left),
          dateStart:pagePxToMm(pageRect,cells[0].getBoundingClientRect().left-pageRect.left),
          regStart:pagePxToMm(pageRect,cells[1].getBoundingClientRect().left-pageRect.left),
          jobStart:pagePxToMm(pageRect,cells[2].getBoundingClientRect().left-pageRect.left),
          taskStart:pagePxToMm(pageRect,cells[3].getBoundingClientRect().left-pageRect.left),
          superStart:pagePxToMm(pageRect,cells[4].getBoundingClientRect().left-pageRect.left),
          end:pagePxToMm(pageRect,frameRect.right-pageRect.left),
          rowHeight:firstRowRect?pagePxToMm(pageRect,firstRowRect.height):OTHER_LAYOUT_DEFAULTS.rowHeight,
          textTop:0,
          textLeft:0
        };
      }
      function writeOtherLayoutPrintCssVars(page, measurements){
        var target=document.documentElement;
        measurements=cloneMeasurementState(measurements);
        writeMeasurementCssVars('--other-layout',measurements,target);
      }
      function updateOtherLayoutPreview(measurements){
        if(!otherLayoutPreviewEl) return;
        measurements=cloneMeasurementState(measurements);
        measurements.dateStart=measurements.left;
        var pageWidth=210,pageHeight=148,frameWidth=Math.max(1,measurements.end-measurements.left),rowHeight=Math.max(.1,measurements.rowHeight),headerHeight=Math.max(.1,measurements.headerHeight),frameTop=Math.max(0,measurements.top),dateWidth=Math.max(0,(measurements.regStart||0)-(measurements.left||0)),regWidth=Math.max(0,(measurements.jobStart||0)-(measurements.regStart||0)),jobWidth=Math.max(0,(measurements.taskStart||0)-(measurements.jobStart||0)),taskWidth=Math.max(0,(measurements.superStart||0)-(measurements.taskStart||0)),superWidth=Math.max(0,(measurements.end||0)-(measurements.superStart||0));
        function pct(num, den){ return (Math.max(0,num)/Math.max(.1,den)*100).toFixed(3)+'%'; }
        if(otherOverlayTopMeasureValueEl){
          if(document.activeElement!==otherOverlayTopMeasureValueEl){
            otherOverlayTopMeasureValueEl.value=String(Number((measurements.top||0).toFixed(1)));
          }
        }
        if(otherOverlayRowMeasureValueEl){
          if(document.activeElement!==otherOverlayRowMeasureValueEl){
            otherOverlayRowMeasureValueEl.value=String(Number((measurements.rowHeight||0).toFixed(1)));
          }
        }
        if(otherOverlayTaskMeasureValueEl){
          if(document.activeElement!==otherOverlayTaskMeasureValueEl){
            otherOverlayTaskMeasureValueEl.value=String(Number((measurements.taskStart||0).toFixed(1)));
          }
        }
        if(otherOverlayTaskWidthMeasureValueEl){
          if(document.activeElement!==otherOverlayTaskWidthMeasureValueEl){
            otherOverlayTaskWidthMeasureValueEl.value=String(Number(taskWidth.toFixed(1)));
          }
        }
        if(otherOverlayDateMeasureValueEl){
          if(document.activeElement!==otherOverlayDateMeasureValueEl){
            otherOverlayDateMeasureValueEl.value=String(Number(dateWidth.toFixed(1)));
          }
        }
        if(otherOverlayRegMeasureValueEl){
          if(document.activeElement!==otherOverlayRegMeasureValueEl){
            otherOverlayRegMeasureValueEl.value=String(Number(regWidth.toFixed(1)));
          }
        }
        if(otherOverlayJobMeasureValueEl){
          if(document.activeElement!==otherOverlayJobMeasureValueEl){
            otherOverlayJobMeasureValueEl.value=String(Number(jobWidth.toFixed(1)));
          }
        }
        if(otherOverlaySuperMeasureValueEl){
          if(document.activeElement!==otherOverlaySuperMeasureValueEl){
            otherOverlaySuperMeasureValueEl.value=String(Number(superWidth.toFixed(1)));
          }
        }
        otherLayoutPreviewEl.style.setProperty('--preview-frame-top', pct(frameTop,pageHeight));
        otherLayoutPreviewEl.style.setProperty('--preview-frame-left', pct(measurements.left,pageWidth));
        otherLayoutPreviewEl.style.setProperty('--preview-frame-width', pct(frameWidth,pageWidth));
        otherLayoutPreviewEl.style.setProperty('--preview-header-height', pct(headerHeight,pageHeight));
        otherLayoutPreviewEl.style.setProperty('--preview-row-height', pct(rowHeight,pageHeight));
        otherLayoutPreviewEl.style.setProperty('--preview-date-width', pct(dateWidth,frameWidth));
        otherLayoutPreviewEl.style.setProperty('--preview-reg-width', pct(regWidth,frameWidth));
        otherLayoutPreviewEl.style.setProperty('--preview-job-width', pct(jobWidth,frameWidth));
        otherLayoutPreviewEl.style.setProperty('--preview-task-width', pct(taskWidth,frameWidth));
        otherLayoutPreviewEl.style.setProperty('--preview-sup-width', pct(superWidth,frameWidth));
        requestAnimationFrame(syncOtherOverlaySampleGuide);
      }
      function syncOtherOverlaySampleGuide(){
        if(!otherOverlaySampleFrameEl||!otherOverlaySampleHeaderEl||!otherOverlaySampleRowEl||!otherOverlayTaskHeaderEl) return;
        var frameRect=otherOverlaySampleFrameEl.getBoundingClientRect();
        var headerRect=otherOverlaySampleHeaderEl.getBoundingClientRect();
        var rowRect=otherOverlaySampleRowEl.getBoundingClientRect();
        var taskCell=otherOverlaySampleRowEl.children&&otherOverlaySampleRowEl.children[3]?otherOverlaySampleRowEl.children[3]:null;
        var taskRect=taskCell?taskCell.getBoundingClientRect():otherOverlayTaskHeaderEl.getBoundingClientRect();
        if(!frameRect.width||!headerRect.width||!rowRect.width||!taskRect.width) return;
        var measureHeight=Math.max(0,rowRect.top-frameRect.top+1);
        var measureLeft=Math.max(0,taskRect.left-frameRect.left+(taskRect.width/2));
        var taskStartWidth=Math.max(0,taskRect.left-frameRect.left+1);
        var taskWidth=Math.max(0,taskRect.width-1);
        otherOverlaySampleFrameEl.style.setProperty('--other-overlay-measure-height',measureHeight.toFixed(2)+'px');
        otherOverlaySampleFrameEl.style.setProperty('--other-overlay-measure-left',measureLeft.toFixed(2)+'px');
        otherOverlaySampleFrameEl.style.setProperty('--other-overlay-row-measure-top',Math.max(0,rowRect.top-frameRect.top+1).toFixed(2)+'px');
        otherOverlaySampleFrameEl.style.setProperty('--other-overlay-row-measure-height',Math.max(0,rowRect.height-1).toFixed(2)+'px');
        otherOverlaySampleFrameEl.style.setProperty('--other-overlay-row-measure-left',measureLeft.toFixed(2)+'px');
        otherOverlaySampleFrameEl.style.setProperty('--other-overlay-task-measure-width',taskStartWidth.toFixed(2)+'px');
        otherOverlaySampleFrameEl.style.setProperty('--other-overlay-task-measure-top',Math.max(0,rowRect.top-frameRect.top+rowRect.height+6).toFixed(2)+'px');
        otherOverlaySampleFrameEl.style.setProperty('--other-overlay-task-width-measure-left',Math.max(0,taskRect.left-frameRect.left).toFixed(2)+'px');
        otherOverlaySampleFrameEl.style.setProperty('--other-overlay-task-width-measure-width',taskWidth.toFixed(2)+'px');
        otherOverlaySampleFrameEl.style.setProperty('--other-overlay-task-width-measure-top',Math.max(0,rowRect.top-frameRect.top+rowRect.height+6).toFixed(2)+'px');
      }
      function validatedOtherLayoutMeasurements(measurements){
        measurements=cloneMeasurementState(measurements);
        measurements.dateStart=measurements.left;
        if(!(measurements.top>=0)){
          throw new Error('Top value must be zero or greater.');
        }
        if(!(measurements.left<=measurements.dateStart&&measurements.dateStart<measurements.regStart&&measurements.regStart<measurements.jobStart&&measurements.jobStart<measurements.taskStart&&measurements.taskStart<measurements.superStart&&measurements.superStart<measurements.end)){
          throw new Error('Layout values must increase from left edge to end edge.');
        }
        return measurements;
      }
      function openOtherLayoutModal(){
        updateOtherLayoutPreview(otherLayoutMeasurements);
        if(otherLayoutModal) otherLayoutModal.className='modal-backdrop open';
        if(otherOverlayTopMeasureValueEl) setTimeout(function(){ otherOverlayTopMeasureValueEl.focus(); if(otherOverlayTopMeasureValueEl.select) otherOverlayTopMeasureValueEl.select(); },0);
      }
      function bindOtherLayoutMeasurementInput(input, applyValue){
        if(!input) return;
        input.addEventListener('input',function(){
          var value=Number(input.value);
          if(!isFinite(value)||value<=0) return;
          applyValue(value);
          updateOtherLayoutPreview(otherLayoutMeasurements);
        });
        input.addEventListener('blur',function(){ updateOtherLayoutPreview(otherLayoutMeasurements); });
      }
      function closeOtherLayoutModal(){ if(otherLayoutModal) otherLayoutModal.className='modal-backdrop'; }
      function currentSingleAircraftRegFilter(){ return activeFilters.aircraftReg.length===1?s(activeFilters.aircraftReg[0]).toUpperCase():''; }
      function rowAircraftRegFilterValue(row){ return s(row&&row['A/C Reg']).toUpperCase()||s(row&&row.__pageFilterAircraftReg).toUpperCase(); }
      function rowMatchesFilters(row){ var i; if(activeFilters.aircraftType.length){ var typeMatch=false,rowAircraftFilterText=normalizedText(aircraftFilterValueForRow(row)); for(i=0;i<activeFilters.aircraftType.length;i++){ if(rowAircraftFilterText.indexOf(normalizedText(activeFilters.aircraftType[i]))!==-1){ typeMatch=true; break; } } if(!typeMatch) return false; } if(activeFilters.aircraftReg.length){ var regMatch=false,rowRegFilterText=normalizedText(rowAircraftRegFilterValue(row)); for(i=0;i<activeFilters.aircraftReg.length;i++){ if(rowRegFilterText.indexOf(normalizedText(activeFilters.aircraftReg[i]))!==-1){ regMatch=true; break; } } if(!regMatch) return false; } if(activeFilters.supervisor.length){ var supervisorName=normalizedText(supervisorNameView(row)),supervisorFull=normalizedText([supervisorNameView(row),supervisorStampView(row),supervisorLicenceView(row)].filter(Boolean).join(' | ')),supervisorMatch=false; for(i=0;i<activeFilters.supervisor.length;i++){ var supervisorNeedle=normalizedText(activeFilters.supervisor[i]); if(supervisorName.indexOf(supervisorNeedle)!==-1||supervisorFull.indexOf(supervisorNeedle)!==-1){ supervisorMatch=true; break; } } if(!supervisorMatch) return false; } if(activeFilters.chapter.length){ var chapterMatch=false; for(i=0;i<activeFilters.chapter.length;i++){ var chapterNeedle=activeFilters.chapter[i]; if(chapterNeedle===BLANK_CHAPTER_FILTER){ if(!s(row['Chapter'])){ chapterMatch=true; break; } continue; } chapterNeedle=normalizedText(chapterNeedle); if(normalizedText(chapterLabelText(row)).indexOf(chapterNeedle)!==-1||normalizedText(s(row['Chapter']))===chapterNeedle){ chapterMatch=true; break; } } if(!chapterMatch) return false; } return true; }
      function normalizeSearchText(value){ return String(value==null?'':value).toLowerCase().replace(/\s+/g,' ').trim(); }
      function rowSearchHaystack(row){
        row=row||{};
        var key=[row['Job No']||'',row['Task Detail']||'',row['Rewriten for cap741']||'',row['Flags']||''].join('\u0001');
        if(row.__searchCacheKey!==key){
          row.__searchCacheKey=key;
          row.__searchCacheValue=normalizeSearchText((row['Job No']||'')+' '+(row['Task Detail']||'')+' '+(row['Rewriten for cap741']||'')+' '+(row['Flags']||''));
        }
        return row.__searchCacheValue||'';
      }
      function rowMatchesSearch(row){ if(!normalizedSearchQuery) return true; return rowSearchHaystack(row).indexOf(normalizedSearchQuery)!==-1; }
      function shouldPreserveFilteredRowSlots(){ return !!activeFilters.supervisor.length; }
      function buildRenderedGroups(activeRows, visibleRows){
        var rendered=[],i,j,k,group,pages,pageItems,filteredItems,visibleById;
        if(!shouldPreserveFilteredRowSlots()){
          var visibleGroups=groupRows(visibleRows);
          for(i=0;i<visibleGroups.length;i++) rendered.push({ group:visibleGroups[i], pages:paginate(visibleGroups[i].rows) });
          return rendered;
        }
        visibleById=Object.create(null);
        for(i=0;i<visibleRows.length;i++) visibleById[String(visibleRows[i].__rowId)]=true;
        var allGroups=groupRows(activeRows);
        for(i=0;i<allGroups.length;i++){
          group=allGroups[i];
          pages=paginate(group.rows);
          filteredItems=[];
          for(j=0;j<pages.length;j++){
            pageItems=[];
            for(k=0;k<pages[j].length;k++){
              if(visibleById[String(pages[j][k].row.__rowId)]) pageItems.push(pages[j][k]);
            }
            if(pageItems.length) filteredItems.push(pageItems);
          }
          if(filteredItems.length) rendered.push({ group:group, pages:filteredItems });
        }
        return rendered;
      }
      function renderEmptyState(){ if(!hasWorkbookDataLoaded()&&!hasActiveFilters()&&!hasActiveSearch()) return '<div class="empty-state empty-state-blank" data-transition-key="empty-state"><div class="empty-state-title">No workbook loaded</div></div>'; var title='No pages match your search'; var copy='Try a different Job No or task detail search.'; var button='<button type="button" data-clear-search="1">Clear search</button>'; if(hasActiveFilters()&&hasActiveSearch()){ title='No pages match these filters and search'; copy='Try a broader search, change the filters, or clear everything to show the full logbook again.'; button='<button type="button" data-clear-all-results="1">Clear search and filters</button>'; } else if(hasActiveFilters()){ title='No pages match these filters'; copy='Try a broader mix of aircraft type, registration, supervisor, or chapter, or clear the filters to show the full logbook again.'; button='<button type="button" data-clear-filters="1">Clear filters</button>'; } return '<div class="empty-state" data-transition-key="empty-state"><div class="empty-state-title">'+title+'</div><div class="empty-state-copy">'+copy+'</div>'+button+'</div>'; }

      // ---- Utilities ----
      function setLoadingState(active, title, text){ if(loadingTitleEl&&title!=null) loadingTitleEl.textContent=title; if(loadingTextEl&&text!=null) loadingTextEl.textContent=text; if(loadingOverlay){ loadingOverlay.className=active?'loading-overlay open':'loading-overlay'; loadingOverlay.setAttribute('aria-hidden',active?'false':'true'); } document.body.classList.toggle('busy',!!active); if(loadBtn){ loadBtn.disabled=!!active; } }
      function nextPaint(){ return new Promise(function(resolve){ requestAnimationFrame(function(){ requestAnimationFrame(resolve); }); }); }
      function renderedLayoutElements(){ return Array.prototype.slice.call(pagesEl.querySelectorAll('.page, .empty-state, tr[data-row-key]')); }
      function renderedLayoutKey(el){ return el?(el.getAttribute('data-row-key')||el.getAttribute('data-page-key')||el.getAttribute('data-transition-key')||''):''; }
      function captureRenderedPositions(){ var map=Object.create(null),els=renderedLayoutElements(); for(var i=0;i<els.length;i++){ var key=renderedLayoutKey(els[i]); if(key) map[key]=els[i].getBoundingClientRect(); } return map; }
      function animateRenderedPositions(previous){ if(!previous||!pagesEl||typeof requestAnimationFrame!=='function') return; requestAnimationFrame(function(){ var els=renderedLayoutElements(); for(var i=0;i<els.length;i++){ var el=els[i],key=renderedLayoutKey(el),before=previous[key],after=el.getBoundingClientRect(); if(before){ var deltaY=before.top-after.top; if(Math.abs(deltaY)>1&&typeof el.animate==='function'){ el.animate([{ transform:'translateY('+deltaY+'px)' },{ transform:'translateY(0)' }],{ duration:240, easing:'cubic-bezier(.2,.8,.2,1)' }); } } else if(typeof el.animate==='function'){ el.animate([{ opacity:.35, transform:'translateY(14px)' },{ opacity:1, transform:'translateY(0)' }],{ duration:220, easing:'ease-out' }); } } }); }
      function renderAllWithMotion(){ var previous=captureRenderedPositions(); renderAll(); animateRenderedPositions(previous); }
      function markSharedDatalistsDirty(){
        sharedDatalistsCache='';
        aircraftOptionsByTypeCache=Object.create(null);
        if(sharedListsEl) sharedListsEl.__renderedHtml='';
      }
      function syncSharedDatalists(html){
        if(!sharedListsEl) return;
        var nextHtml=String(html==null?sharedDatalistsHtml():html);
        if(sharedListsEl.__renderedHtml===nextHtml) return;
        sharedListsEl.innerHTML=nextHtml;
        sharedListsEl.__renderedHtml=nextHtml;
      }
      function normalizedText(value){ return s(value).toLowerCase(); }
      function chapterLabelText(row){ var completed=completeChapterParts(row&&row['Chapter'],row&&row['Chapter Description']); return completed.chapterDesc?(completed.chapter+' - '+completed.chapterDesc):completed.chapter; }
      function rowChapterDescriptionView(row){ return s(row&&row['Chapter Description'])||chapterDescriptionForCode(row&&row['Chapter']); }
      function supervisorReferenceForRow(row){
        var supervisorId=s(row&&row[SUPERVISOR_ID_FIELD]),record=supervisorId?supervisorRecordForId(supervisorId):null;
        if(!record&&s(row&&row['Approval Name'])) record=supervisorRecordFor(row['Approval Name']);
        return record||null;
      }
      function supervisorNameView(row){ var record=supervisorReferenceForRow(row); return s(row&&row['Approval Name'])||s(record&&record.name); }
      function supervisorStampView(row){ var record=supervisorReferenceForRow(row); return s(row&&row['Approval stamp'])||s(record&&record.stamp); }
      function supervisorLicenceView(row){ var record=supervisorReferenceForRow(row); return s(row&&row['Aprroval Licence No.'])||s(record&&record.licence); }
      function fieldAffectsRowLayout(field){ return field==='Task Detail'||field==='Rewriten for cap741'; }
      function fieldNeedsLiveLayoutRefresh(field){ return field==='Task Detail'||field==='Rewriten for cap741'; }
      function liveLayoutUnitsForField(row, field){
        if(!row||!(field==='Task Detail'||field==='Rewriten for cap741')) return 0;
        return unitsFor(row);
      }
      function captureContentEditableSelection(root){ var sel=window.getSelection&&window.getSelection(); if(!sel||!sel.rangeCount) return {start:0,end:0}; var range=sel.getRangeAt(0); if(!root.contains(range.startContainer)||!root.contains(range.endContainer)) return {start:0,end:0}; var startRange=document.createRange(); startRange.selectNodeContents(root); startRange.setEnd(range.startContainer,range.startOffset); var endRange=document.createRange(); endRange.selectNodeContents(root); endRange.setEnd(range.endContainer,range.endOffset); return {start:startRange.toString().length,end:endRange.toString().length}; }
      function setContentEditableSelection(root, start, end){ var walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT,null); var node,pos=0,startNode=null,endNode=null,startOffset=0,endOffset=0; while((node=walker.nextNode())){ var next=pos+node.nodeValue.length; if(startNode==null&&start<=next){ startNode=node; startOffset=Math.max(0,start-pos); } if(endNode==null&&end<=next){ endNode=node; endOffset=Math.max(0,end-pos); break; } pos=next; } var range=document.createRange(); if(startNode&&endNode){ range.setStart(startNode,startOffset); range.setEnd(endNode,endOffset); } else { range.selectNodeContents(root); range.collapse(false); } var sel=window.getSelection&&window.getSelection(); if(sel){ sel.removeAllRanges(); sel.addRange(range); } }
      function captureEditorSnapshot(target){ var el=target&&target.closest?(target.closest('.editable-cell')||target.closest('input.field-input[data-edit-field], input.field-input[data-new-row]')):null; if(!el) return null; var field=el.getAttribute('data-edit-field'),rowId=el.getAttribute('data-row-id'); if(!field||rowId==null) return null; var snapshot={field:field,rowId:rowId,isInput:el.tagName==='INPUT',scrollX:window.scrollX,scrollY:window.scrollY}; if(snapshot.isInput){ snapshot.start=typeof el.selectionStart==='number'?el.selectionStart:null; snapshot.end=typeof el.selectionEnd==='number'?el.selectionEnd:snapshot.start; } else { var selection=captureContentEditableSelection(el); snapshot.start=selection.start; snapshot.end=selection.end; } return snapshot; }
      function restoreEditorSnapshot(snapshot){ if(!snapshot) return; var selector='[data-row-id="'+snapshot.rowId+'"][data-edit-field="'+snapshot.field+'"]'; var el=pagesEl.querySelector(selector); if(!el||typeof el.focus!=='function') return; try { el.focus({preventScroll:true}); } catch(e){ el.focus(); } if(snapshot.isInput){ if(typeof el.setSelectionRange==='function'&&snapshot.start!=null){ try { el.setSelectionRange(snapshot.start,snapshot.end==null?snapshot.start:snapshot.end); } catch(e){} } } else { setContentEditableSelection(el,snapshot.start||0,snapshot.end||snapshot.start||0); } if(typeof window.scrollTo==='function') window.scrollTo(snapshot.scrollX||0,snapshot.scrollY||0); }
      function refreshLayoutPreservingEditor(){ var snapshot=liveLayoutEditorState; liveLayoutEditorState=null; renderAll(); restoreEditorSnapshot(snapshot); }
      function scheduleLiveLayoutRefresh(snapshot, delay){ liveLayoutEditorState=snapshot; clearTimeout(layoutTimer); layoutTimer=setTimeout(refreshLayoutPreservingEditor,typeof delay==='number'?delay:60); }
      function refreshLayoutIfIdle(){ if(editorIsActive()){ scheduleLayoutRefresh(350); return; } renderAll(); }
      function scheduleLayoutRefresh(delay){ clearTimeout(layoutTimer); layoutTimer=setTimeout(refreshLayoutIfIdle,typeof delay==='number'?delay:300); }
      function scheduleLocalDraftPersist(){ return; }
      function refreshUnsavedChangesState(){ hasUnsavedChanges=settingsDirty||dataDirty; syncSaveButtonState(false); }
      function workbookMimeType(){ return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'; }
      function workbookAcceptValue(){ return '.xlsx,'+workbookMimeType(); }
      function workbookFileName(name){ var fileName=s(name)||'cap741-data.xlsx'; return /\.xlsx$/i.test(fileName)?fileName:(fileName+'.xlsx'); }
      function currentWorkbookFileName(){ return workbookFileName(s(activeStorageSource&&activeStorageSource.name)||linkedWorkbookName||'cap741-data.xlsx'); }
      function persistentExcelLinkingSupported(){ return filePickerSupported()&&fileSavePickerSupported(); }
      function usingExcelDownloadFallback(){ return sourceType(activeStorageSource)!==STORAGE_SOURCE_GOOGLE&&!fileSavePickerSupported(); }
      function setSessionExcelSource(name, downloadOnly){
        var fileName=workbookFileName(name);
        setAutoLoadDefaultWorkbook(false);
        writeStoredJson(STORAGE_SOURCE_KEY,null);
        removeStoredHandle(LINKED_FILE_KEY);
        setLinkedWorkbookName(null);
        setActiveStorageSource({type:STORAGE_SOURCE_EXCEL,name:fileName,downloadOnly:!!downloadOnly},false);
        return fileName;
      }
      function downloadWorkbookFile(name){
        var fileName=workbookFileName(name),data=XLSX.write(buildWorkbookFromState(),{bookType:'xlsx',type:'array'}),blob=new Blob([data],{type:workbookMimeType()}),url=(window.URL||window.webkitURL).createObjectURL(blob),link=document.createElement('a');
        link.href=url;
        link.download=fileName;
        link.rel='noopener';
        link.style.position='fixed';
        link.style.left='-9999px';
        document.body.appendChild(link);
        try {
          link.click();
        } finally {
          setTimeout(function(){
            if(link.parentNode) link.parentNode.removeChild(link);
            (window.URL||window.webkitURL).revokeObjectURL(url);
          },800);
        }
        resetSavedLogbookState();
        settingsDirty=false;
        return fileName;
      }
      function showExcelDownloadNote(fileName, prefix){ note((prefix||'Excel export ready')+': '+workbookFileName(fileName)+'. If Safari opens a preview, use Share and choose Save to Files.'); }
      function syncSaveButtonState(isSaving){ if(!saveFileBtn) return; var canShow=hasWorkbookDataLoaded(),targetLabel=sourceType(activeStorageSource)===STORAGE_SOURCE_GOOGLE?(s(activeStorageSource.title)||'Google Sheet'):currentWorkbookFileName(); saveFileBtn.classList.toggle('open',canShow&&(!!hasUnsavedChanges||!!isSaving)); saveFileBtn.classList.toggle('saving',!!isSaving); saveFileBtn.disabled=!canShow||!!isSaving; saveFileBtn.setAttribute('aria-label',isSaving?'Saving changes to file':'Save changes to file'); saveFileBtn.title=isSaving?('Saving '+targetLabel+'...'):('Save changes to '+targetLabel); }
      function setPrintOptionsOpen(open){ if(!printOptionsEl||!printBtn) return; printOptionsEl.classList.toggle('open',!!open); printOptionsEl.setAttribute('aria-hidden',open?'false':'true'); printBtn.classList.toggle('active',!!open); }
      function syncLoadOptionLabels(){
        var loadExistingTextEl=loadExistingBtn&&loadExistingBtn.querySelector&&loadExistingBtn.querySelector('span');
        if(loadExistingTextEl) loadExistingTextEl.textContent=loadButtonMode==='link'&&persistentExcelLinkingSupported()?'Link Existing File':'Load Existing File';
        if(loadGoogleSheetBtn) loadGoogleSheetBtn.textContent=loadButtonMode==='link'?'Link Google Sheet':'Load Google Sheet';
      }
      function setLoadOptionsOpen(open){ if(!loadOptionsEl||!loadBtn) return; syncLoadOptionLabels(); loadOptionsEl.classList.toggle('open',!!open); loadOptionsEl.setAttribute('aria-hidden',open?'false':'true'); loadBtn.classList.toggle('active',!!open); }
      function hasWorkbookDataLoaded(){ return !!(sourceType(activeStorageSource)!==STORAGE_SOURCE_NONE||rows.length||AIRCRAFT_GROUP_ROWS.length||SUPERVISOR_RECORDS.length||s(LOG_OWNER_INFO.name)||s(LOG_OWNER_INFO.signature)||s(LOG_OWNER_INFO.stamp)||(CHAPTER_OPTIONS.length&&!chapterOptionsMatchDefaults())); }
      function setLinkedWorkbookName(handle){ linkedWorkbookName=handle&&handleIsWorkbook(handle)?s(handle.name):''; }
      function readStoredJson(key){ try { return window.localStorage ? JSON.parse(window.localStorage.getItem(key)||'null') : null; } catch(e){ return null; } }
      function writeStoredJson(key, value){ try { if(!window.localStorage) return; if(value==null) window.localStorage.removeItem(key); else window.localStorage.setItem(key,JSON.stringify(value)); } catch(e){} }
      function sourceType(source){ return source&&source.type ? source.type : STORAGE_SOURCE_NONE; }
      function setActiveStorageSource(source, persist){ activeStorageSource=source||{ type: STORAGE_SOURCE_NONE }; if(persist!==false){ if(sourceType(activeStorageSource)===STORAGE_SOURCE_NONE) writeStoredJson(STORAGE_SOURCE_KEY,null); else writeStoredJson(STORAGE_SOURCE_KEY,activeStorageSource); } }
      function loadStoredSource(){ var stored=readStoredJson(STORAGE_SOURCE_KEY); return stored&&stored.type ? stored : { type: STORAGE_SOURCE_NONE }; }
      function googleSheetUrl(source){ var spreadsheetId=s(source&&source.spreadsheetId); return spreadsheetId?('https://docs.google.com/spreadsheets/d/'+encodeURIComponent(spreadsheetId)+'/edit') : ''; }
      function googleClientId(){ try { return s(GOOGLE_CLIENT_ID||(window.CAP741_GOOGLE_CONFIG&&window.CAP741_GOOGLE_CONFIG.clientId)||window.CAP741_GOOGLE_CLIENT_ID||window.localStorage&&window.localStorage.getItem(GOOGLE_CLIENT_ID_KEY)||''); } catch(e){ return ''; } }
      function ensureGoogleClientId(interactive){ var clientId=googleClientId(); if(clientId) return clientId; if(!interactive) return ''; clientId=s(window.prompt('Enter your Google OAuth Client ID for Google Sheets access.',clientId||'')); if(clientId&&window.localStorage) window.localStorage.setItem(GOOGLE_CLIENT_ID_KEY,clientId); return clientId; }
      function shouldAutoLoadDefaultWorkbook(){ try { return window.localStorage ? window.localStorage.getItem(AUTO_LOAD_DEFAULT_KEY)!=='0' : true; } catch(e){ return true; } }
      function setAutoLoadDefaultWorkbook(enabled){ try { if(window.localStorage) window.localStorage.setItem(AUTO_LOAD_DEFAULT_KEY,enabled?'1':'0'); } catch(e){} }
      function showTopMessage(msg, kind){
        if(!errorBox) return;
        errorBox.className='error'+(kind&&kind!=='error'?' '+kind:'');
        errorBox.style.display='block';
        if(errorTextEl) errorTextEl.textContent=msg;
        else errorBox.textContent=msg;
        document.body.classList.add('has-top-error');
      }
      function fail(msg){ showTopMessage(msg,'error'); }
      function note(msg){ showTopMessage(msg,'info'); }
      function success(msg){ showTopMessage(msg,'success'); }
      function clearFail(){ if(!errorBox) return; errorBox.style.display='none'; if(errorTextEl) errorTextEl.textContent=''; else errorBox.textContent=''; errorBox.className='error'; document.body.classList.remove('has-top-error'); }
      function saveFailureMessage(error){ var message='Could not save: '+(error&&error.message?error.message:'Unknown error.'); if(sourceType(activeStorageSource)===STORAGE_SOURCE_GOOGLE) return message; if(fileSavePickerSupported()&&message.toLowerCase().indexOf('close it in excel')===-1&&message.toLowerCase().indexOf('open in excel')===-1) message+=' If cap741-data.xlsx is open in Excel, close it and try again.'; return message; }
      async function copyTextToClipboard(text){
        if(!s(text)) return false;
        if(navigator.clipboard&&typeof navigator.clipboard.writeText==='function'){
          await navigator.clipboard.writeText(text);
          return true;
        }
        var input=document.createElement('textarea');
        input.value=text;
        input.setAttribute('readonly','readonly');
        input.style.position='fixed';
        input.style.left='-9999px';
        document.body.appendChild(input);
        input.select();
        try {
          return !!document.execCommand('copy');
        } finally {
          document.body.removeChild(input);
        }
      }
      function flashStorageCopyButton(btn){
        if(!btn) return;
        var originalLabel=btn.getAttribute('aria-label')||'Copy link';
        btn.classList.add('copied');
        btn.title='Copied';
        btn.setAttribute('aria-label','Link copied');
        clearTimeout(btn.__copiedTimer);
        btn.__copiedTimer=setTimeout(function(){
          btn.classList.remove('copied');
          btn.title='Copy link';
          btn.setAttribute('aria-label',originalLabel);
        },1200);
      }
      function showGoogleSheetModal(options){
        options=options||{};
        return new Promise(function(resolve){
          googleSheetModalResolver=resolve;
          if(googleSheetModalTitleEl) googleSheetModalTitleEl.textContent=options.title||'Google Sheet';
          if(googleSheetModalCopyEl) googleSheetModalCopyEl.textContent=options.copy||'';
          if(googleSheetModalNoteEl) googleSheetModalNoteEl.textContent=options.note||'';
          if(googleSheetInputLabelEl) googleSheetInputLabelEl.textContent=options.label||'Sheet URL or ID';
          if(googleSheetInputWrapEl){
            googleSheetInputWrapEl.hidden=!options.input;
            googleSheetInputWrapEl.style.display=options.input?'grid':'none';
          }
          if(googleSheetUrlInputEl){
            googleSheetUrlInputEl.type=options.inputType||'text';
            googleSheetUrlInputEl.autocomplete=options.autocomplete||'off';
            googleSheetUrlInputEl.value=options.input?(options.value||''):'';
            googleSheetUrlInputEl.placeholder=options.placeholder||'Paste Google Sheet URL or ID';
          }
          var linkUrl=s(options.linkUrl);
          if(googleSheetResultRowEl) googleSheetResultRowEl.hidden=!linkUrl;
          if(googleSheetResultLinkEl){
            googleSheetResultLinkEl.href=linkUrl||'#';
            googleSheetResultLinkEl.textContent=linkUrl;
          }
          if(googleSheetCancelBtn) googleSheetCancelBtn.style.display=options.hideCancel?'none':'inline-flex';
          if(googleSheetOkBtn) googleSheetOkBtn.textContent=options.okLabel||'Continue';
          if(googleSheetModal) googleSheetModal.className='modal-backdrop open';
          setTimeout(function(){
            if(options.input&&googleSheetUrlInputEl){
              googleSheetUrlInputEl.focus();
              try { googleSheetUrlInputEl.select(); } catch(e){}
            }
            else if(googleSheetOkBtn) googleSheetOkBtn.focus();
          },0);
        });
      }
      function closeGoogleSheetModal(result){
        if(googleSheetModal) googleSheetModal.className='modal-backdrop';
        if(googleSheetModalResolver){
          var resolve=googleSheetModalResolver;
          googleSheetModalResolver=null;
          resolve(result==null?null:result);
        }
      }
      async function requestGoogleSheetIdFromModal(title, copyText, okLabel){
        var value=await showGoogleSheetModal({
          title:title||'Connect Google Sheet',
          copy:copyText||'Paste the Google Sheet URL or just the sheet ID.',
          note:'You can paste either the full Google Sheet URL or the sheet ID.',
          input:true,
          okLabel:okLabel||'Connect'
        });
        return googleSheetIdFromInput(value||'');
      }
      async function showCreatedGoogleSheetNotice(source){
        var url=googleSheetUrl(source);
        if(!url) return;
        await showGoogleSheetModal({
          title:'Google Sheet Created',
          copy:'Your new Google Sheet is ready.',
          note:'Take a note of this URL. You can use it later to reconnect this sheet from the app.',
          linkUrl:url,
          okLabel:'Close',
          hideCancel:true
        });
      }
      function protectedDataStore(){ return window.CAP741_PROTECTED_DATA||null; }
      function chapterDataStore(){
        var store=window.CAP741_CHAPTER_DATA||window.CAP741_PREFILLED_CHAPTERS||[];
        return Array.isArray(store)?store:[];
      }
      function flagDataStore(){
        var store=window.CAP741_FLAG_DATA||window.CAP741_PREFILLED_FLAGS||[];
        return Array.isArray(store)?store:[];
      }
      function protectedPayloadValue(payload){ return typeof payload==='string' ? s(payload) : s(payload&&payload.cipher); }
      function protectedDataAvailable(){
        var store=protectedDataStore();
        if(!store) return false;
        if(s(store.scheme)==='pbkdf2-aes-cbc-v1') return !!(protectedPayloadValue(store.aircraftPayload)&&protectedPayloadValue(store.supervisorPayload)&&s(store.passwordSalt)&&s(store.passwordVerifier));
        return !!(store&&s(store.password)&&protectedPayloadValue(store.aircraftPayload)&&protectedPayloadValue(store.supervisorPayload));
      }
      function decodeProtectedPayload(payload, key){
        var text=s(payload),pass=s(key);
        if(!text||!pass) throw new Error('Protected data could not be decoded.');
        var raw=window.atob(text),passCodes=[],i,decodedBytes,decoder;
        for(i=0;i<pass.length;i++) passCodes.push(pass.charCodeAt(i));
        decodedBytes=new Uint8Array(raw.length);
        for(i=0;i<raw.length;i++) decodedBytes[i]=raw.charCodeAt(i)^passCodes[i%passCodes.length];
        if(window.TextDecoder){
          decoder=new TextDecoder('utf-8');
          return decoder.decode(decodedBytes);
        }
        var out='';
        for(i=0;i<decodedBytes.length;i++) out+=String.fromCharCode(decodedBytes[i]);
        try { return decodeURIComponent(escape(out)); } catch(e){ return out; }
      }
      function bytesFromBase64(text){
        var raw=window.atob(s(text)),bytes=new Uint8Array(raw.length),i;
        for(i=0;i<raw.length;i++) bytes[i]=raw.charCodeAt(i);
        return bytes;
      }
      function bytesFromUtf8(text){ return window.TextEncoder ? new TextEncoder().encode(String(text==null?'':text)) : new Uint8Array(unescape(encodeURIComponent(String(text==null?'':text))).split('').map(function(ch){ return ch.charCodeAt(0); })); }
      function concatBytes(a, b){ var left=a instanceof Uint8Array?a:new Uint8Array(a||[]),right=b instanceof Uint8Array?b:new Uint8Array(b||[]),out=new Uint8Array(left.length+right.length); out.set(left,0); out.set(right,left.length); return out; }
      function bytesEqual(a, b){ var left=a instanceof Uint8Array?a:new Uint8Array(a||[]),right=b instanceof Uint8Array?b:new Uint8Array(b||[]),i,diff=0; if(left.length!==right.length) return false; for(i=0;i<left.length;i++) diff|=(left[i]^right[i]); return diff===0; }
      async function sha256Bytes(bytes){
        if(!window.crypto||!window.crypto.subtle) throw new Error('Secure protected data import requires a modern browser.');
        return new Uint8Array(await window.crypto.subtle.digest('SHA-256',bytes));
      }
      async function deriveProtectedMasterBytes(password, saltBase64, iterations){
        if(!window.crypto||!window.crypto.subtle) throw new Error('Secure protected data import requires a modern browser.');
        var baseKey=await window.crypto.subtle.importKey('raw',bytesFromUtf8(password),{name:'PBKDF2'},false,['deriveBits']);
        return new Uint8Array(await window.crypto.subtle.deriveBits({name:'PBKDF2',salt:bytesFromBase64(saltBase64),iterations:Number(iterations)||200000,hash:'SHA-256'},baseKey,256));
      }
      async function protectedVerifierBytes(masterBytes){ return await sha256Bytes(concatBytes(masterBytes,bytesFromUtf8('verify'))); }
      async function protectedEncryptionKey(masterBytes){
        var keyBytes=await sha256Bytes(concatBytes(masterBytes,bytesFromUtf8('encrypt')));
        return await window.crypto.subtle.importKey('raw',keyBytes,{name:'AES-CBC'},false,['decrypt']);
      }
      async function decodeProtectedPayloadV2(payload, password, store){
        var masterBytes=await deriveProtectedMasterBytes(password,store.passwordSalt,store.kdfIterations);
        var expected=bytesFromBase64(store.passwordVerifier);
        var actual=await protectedVerifierBytes(masterBytes);
        if(!bytesEqual(actual,expected)) throw new Error('Incorrect password.');
        var key=await protectedEncryptionKey(masterBytes);
        var plainBuffer=await window.crypto.subtle.decrypt({name:'AES-CBC',iv:bytesFromBase64(payload&&payload.iv)},key,bytesFromBase64(payload&&payload.cipher));
        if(window.TextDecoder) return new TextDecoder('utf-8').decode(new Uint8Array(plainBuffer));
        var out='',decodedBytes=new Uint8Array(plainBuffer),i;
        for(i=0;i<decodedBytes.length;i++) out+=String.fromCharCode(decodedBytes[i]);
        try { return decodeURIComponent(escape(out)); } catch(e){ return out; }
      }
      async function requestProtectedImportPassword(kindLabel){
        return await showGoogleSheetModal({
          title:'Unlock Protected '+kindLabel,
          copy:'Enter the password to import the protected '+kindLabel.toLowerCase()+' data into this logbook.',
          note:'This imports the prefilled reference data from the protected local data.js file.',
          input:true,
          inputType:'password',
          label:'Password',
          placeholder:'Enter password',
          okLabel:'Import',
          autocomplete:'current-password'
        });
      }
      function filePickerSupported(){ return typeof window.showOpenFilePicker==='function'; }
      function fileSavePickerSupported(){ return typeof window.showSaveFilePicker==='function'; }
      function syncLoadButtonAvailability(isLinked){ if(isLinked){ setLoadButtonMode('hidden'); return; } setLoadButtonMode(hasWorkbookDataLoaded()?'link':'load'); }
      function setLoadButtonMode(mode){ if(!loadBtn) return; loadButtonMode=mode||'load'; loadBtn.setAttribute('data-mode',loadButtonMode); if(loadButtonMode==='hidden'){ loadBtn.style.display='none'; setLoadOptionsOpen(false); return; } loadBtn.style.display='block'; loadBtn.textContent=loadButtonMode==='link'&&persistentExcelLinkingSupported()?'Link':'Load'; loadBtn.title=loadButtonMode==='link'?(persistentExcelLinkingSupported()?'Link an Excel file or Google Sheet for saving':'Load an Excel file or link a Google Sheet for saving'):'Load an existing CAP741 source or create a new one'; loadBtn.setAttribute('aria-label',loadButtonMode==='link'?(persistentExcelLinkingSupported()?'Link an Excel file or Google Sheet for saving':'Load an Excel file or link a Google Sheet for saving'):'Load an existing CAP741 source or create a new one'); syncLoadOptionLabels(); }
      function s(v){ return v==null?'':String(v).trim(); }
      function esc(v){ return s(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }
      function normalizePageGrouping(value){ return s(value).toLowerCase()===PAGE_GROUPING_GROUP?PAGE_GROUPING_GROUP:PAGE_GROUPING_TYPE; }
      function cloneAppViewSettings(settings){
        settings=settings||{};
        return {
          showMindMap: settings.showMindMap!==false,
          pageGrouping: normalizePageGrouping(settings.pageGrouping||DEFAULT_APP_VIEW_SETTINGS.pageGrouping),
          referenceOnlySave: boolSettingValue(settings.referenceOnlySave,DEFAULT_APP_VIEW_SETTINGS.referenceOnlySave)
        };
      }
      function boolSettingValue(value, fallback){
        var raw=normalizedText(value);
        if(!raw) return !!fallback;
        if(raw==='0'||raw==='false'||raw==='no'||raw==='off') return false;
        if(raw==='1'||raw==='true'||raw==='yes'||raw==='on') return true;
        return !!fallback;
      }
      function currentPageGrouping(){ return normalizePageGrouping(APP_VIEW_SETTINGS.pageGrouping); }
      function referenceOnlySaveEnabled(){ return APP_VIEW_SETTINGS.referenceOnlySave!==false; }
      function pageGroupingDisplayLabel(value){ return normalizePageGrouping(value)===PAGE_GROUPING_GROUP ? 'Aircraft Group' : 'Aircraft Type'; }
      function syncMindMapButtonVisibility(){
        var hidden=APP_VIEW_SETTINGS.showMindMap===false;
        if(mindMapBtn) mindMapBtn.hidden=hidden;
        document.body.classList.toggle('mindmap-fab-hidden',hidden);
      }
      function infoWorkbookRows(){
        return [
          {Key:'Name',Value:s(LOG_OWNER_INFO.name)},
          {Key:'Signature',Value:s(LOG_OWNER_INFO.signature)},
          {Key:'Stamp',Value:s(LOG_OWNER_INFO.stamp)},
          {Key:'Show Mind Map',Value:APP_VIEW_SETTINGS.showMindMap?'true':'false'},
          {Key:'Page Grouping',Value:currentPageGrouping()},
          {Key:'Reference Save',Value:referenceOnlySaveEnabled()?'true':'false'}
        ];
      }
      function cloneFlagRecords(records){
        return (records||[]).map(function(record){
          return { section:normalizeFlagSection(record&&record.section), flag:s(record&&record.flag), color:s(record&&record.color) };
        }).filter(function(record){ return !!record.flag; });
      }
      function defaultFlagRecords(){
        return cloneFlagRecords(flagDataStore().map(function(record){
          return {
            section:record&&(
              record.section||
              record.Section||
              record.group||
              record.Group||
              record.Flags
            ),
            flag:record&&(record.flag||record.Flag||record.label||record.Label||record.name||record.Name||record.More),
            color:record&&(record.color||record.Color||record.colour||record.Colour||record.Gold)
          };
        }));
      }
      function normalizeFlagToken(value){
        return normalizedText(s(value).replace(/[–—]/g,'-').replace(/\s+/g,' '));
      }
      function normalizeFlagSection(value){
        var token=normalizeFlagToken(value);
        if(token==='more'||token==='expanded'||token==='secondary'||token==='default'||token==='options') return FLAG_SECTION_MORE;
        return FLAG_SECTION_PRIMARY;
      }
      function flagShortLabel(value){
        var label=s(value),match=label.match(/^(.+?)\s+[–-]\s+/);
        if(match&&match[1]) return s(match[1]);
        if(label.length<=18) return label;
        return label.slice(0,15).replace(/\s+$/,'')+'...';
      }
      function flagBadgeLetter(value){
        var label=flagShortLabel(value)||s(value),match=label.match(/[A-Za-z0-9]/);
        if(match&&match[0]) return match[0].toUpperCase();
        return label?label.charAt(0).toUpperCase():'';
      }
      function flagSortIndex(label){
        var token=normalizeFlagToken(label),defaults=defaultFlagRecords();
        for(var i=0;i<defaults.length;i++) if(normalizeFlagToken(defaults[i].flag)===token) return i;
        return defaults.length+1000;
      }
      function flagColorForLabel(label){
        var token=normalizeFlagToken(label),i,defaults=defaultFlagRecords();
        for(i=0;i<FLAG_RECORDS.length;i++) if(normalizeFlagToken(FLAG_RECORDS[i].flag)===token) return s(FLAG_RECORDS[i].color);
        for(i=0;i<defaults.length;i++) if(normalizeFlagToken(defaults[i].flag)===token) return s(defaults[i].color);
        return '';
      }
      function flagSectionForLabel(label){
        var token=normalizeFlagToken(label),i,defaults=defaultFlagRecords();
        for(i=0;i<FLAG_RECORDS.length;i++) if(normalizeFlagToken(FLAG_RECORDS[i].flag)===token) return normalizeFlagSection(FLAG_RECORDS[i].section);
        for(i=0;i<defaults.length;i++) if(normalizeFlagToken(defaults[i].flag)===token) return normalizeFlagSection(defaults[i].section);
        return FLAG_SECTION_PRIMARY;
      }
      function flagRecordForToken(value){
        var token=normalizeFlagToken(value),i,record,defaults=defaultFlagRecords();
        if(!token) return null;
        for(i=0;i<FLAG_RECORDS.length;i++){
          record=FLAG_RECORDS[i];
          if(normalizeFlagToken(record.flag)===token||normalizeFlagToken(flagShortLabel(record.flag))===token) return record;
        }
        for(i=0;i<defaults.length;i++){
          record=defaults[i];
          if(normalizeFlagToken(record.flag)===token||normalizeFlagToken(flagShortLabel(record.flag))===token) return record;
        }
        return null;
      }
      function normalizeFlagSelection(values){
        var parts=Array.isArray(values)?values.slice():String(values==null?'':values).split(/\s*(?:;|\r?\n|\|)\s*/),seen=Object.create(null),matched=[],extra=[],i,raw,record,label,token;
        for(i=0;i<parts.length;i++){
          raw=s(parts[i]);
          if(!raw) continue;
          record=flagRecordForToken(raw);
          label=s(record&&record.flag)||raw;
          token=normalizeFlagToken(label);
          if(!token||seen[token]) continue;
          seen[token]=true;
          if(record) matched.push(label);
          else extra.push(label);
        }
        matched.sort(function(a,b){
          var left=flagSortIndex(a),right=flagSortIndex(b);
          if(left!==right) return left-right;
          return s(a).localeCompare(s(b),undefined,{numeric:true});
        });
        extra.sort(function(a,b){ return s(a).localeCompare(s(b),undefined,{numeric:true}); });
        return matched.concat(extra);
      }
      function serializeFlagSelection(values){ return normalizeFlagSelection(values).join('; '); }
      function rowFlagLabels(row){ return normalizeFlagSelection(row&&row['Flags']); }
      function setRowFlags(row, values){ if(row) row['Flags']=serializeFlagSelection(values); }
      function flagSummaryRecords(row){
        return rowFlagLabels(row).map(function(label){
          var record=flagRecordForToken(label);
          return { flag:s(record&&record.flag)||label, color:s(record&&record.color)||flagColorForLabel(label) };
        });
      }
      function flagSectionSelectHtml(value){
        value=normalizeFlagSection(value);
        return '<select data-col="Section"><option value="'+FLAG_SECTION_PRIMARY+'"'+(value===FLAG_SECTION_PRIMARY?' selected':'')+'>'+FLAG_SECTION_PRIMARY+'</option><option value="'+FLAG_SECTION_MORE+'"'+(value===FLAG_SECTION_MORE?' selected':'')+'>'+FLAG_SECTION_MORE+'</option></select>';
      }
      function renderDetailFlagOption(record, selectedTokens){
        var label=s(record&&record.flag),checked=!!selectedTokens[normalizeFlagToken(label)],color=s(record&&record.color)||'#8091a0';
        return '<label class="detail-flag-option"><input class="detail-flag-check" type="checkbox" data-flag-option="1" value="'+esc(label)+'"'+(checked?' checked':'')+'><span class="detail-flag-chip"><span class="detail-flag-dot" style="background-color:'+esc(color)+'"></span><span class="detail-flag-chip-text">'+esc(label)+'</span></span></label>';
      }
      function renderTaskDetailFlagOptions(selectedValues){
        var selectedTokens=Object.create(null),i,record,primaryHtml=[],moreHtml=[],section;
        for(i=0;i<(selectedValues||[]).length;i++) selectedTokens[normalizeFlagToken(selectedValues[i])]=true;
        for(i=0;i<FLAG_RECORDS.length;i++){
          record=FLAG_RECORDS[i];
          section=normalizeFlagSection(record.section);
          if(section===FLAG_SECTION_MORE) moreHtml.push(renderDetailFlagOption(record,selectedTokens));
          else primaryHtml.push(renderDetailFlagOption(record,selectedTokens));
        }
        if(detailFlagsPrimaryEl) detailFlagsPrimaryEl.innerHTML=primaryHtml.join('');
        if(detailFlagsMoreEl) detailFlagsMoreEl.innerHTML=moreHtml.join('');
        if(detailFlagsMoreWrapEl){
          detailFlagsMoreWrapEl.hidden=!moreHtml.length;
          detailFlagsMoreWrapEl.open=false;
        }
        if(detailFlagsEmptyEl) detailFlagsEmptyEl.hidden=!!(primaryHtml.length||moreHtml.length);
      }
      function readTaskDetailSelectedFlags(){
        var checks=taskDetailModal?taskDetailModal.querySelectorAll('[data-flag-option]:checked'):[],values=[];
        for(var i=0;i<checks.length;i++) values.push(s(checks[i].value));
        return normalizeFlagSelection(values);
      }
      function flagWorkbookRows(){
        return FLAG_RECORDS.map(function(record){
          return { Section:normalizeFlagSection(record.section), Flag:s(record.flag), Color:s(record.color) };
        });
      }
      function applyFlagRows(records){
        records=records||[];
        var parsed=[],seen=Object.create(null),legacy=false,i,record,label,color,section,token,defaults=defaultFlagRecords();
        for(i=0;i<records.length;i++){
          record=records[i]||{};
          if(Object.prototype.hasOwnProperty.call(record,'Flags')||Object.prototype.hasOwnProperty.call(record,'More')||Object.prototype.hasOwnProperty.call(record,'Gold')) legacy=true;
          label=s(record.flag||record.Flag||record.label||record.Label||record.name||record.Name||record.More);
          color=s(record.color||record.Color||record.colour||record.Colour||record.Gold);
          section=normalizeFlagSection(record.section||record.Section||record.group||record.Group||record.Flags);
          if(normalizeFlagToken(label)==='options') continue;
          if(!label) continue;
          token=normalizeFlagToken(label);
          if(seen[token]) continue;
          seen[token]=true;
          parsed.push({ section:section, flag:label, color:color||flagColorForLabel(label) });
        }
        if(legacy){
          for(i=0;i<defaults.length;i++){
            label=defaults[i].flag;
            token=normalizeFlagToken(label);
            if(seen[token]) continue;
            seen[token]=true;
            parsed.push({ section:normalizeFlagSection(defaults[i].section), flag:label, color:s(defaults[i].color) });
          }
        }
        parsed=cloneFlagRecords(parsed);
        parsed.sort(function(a,b){
          var sectionDiff=(normalizeFlagSection(a.section)===FLAG_SECTION_MORE?1:0)-(normalizeFlagSection(b.section)===FLAG_SECTION_MORE?1:0);
          if(sectionDiff) return sectionDiff;
          var left=flagSortIndex(a.flag),right=flagSortIndex(b.flag);
          if(left!==right) return left-right;
          return s(a.flag).localeCompare(s(b.flag),undefined,{numeric:true});
        });
        FLAG_RECORDS=parsed.length?parsed:defaultFlagRecords();
      }

      // ---- Date ----
      function padDatePart(value){ return String(value).padStart(2,'0'); }
      function monthNumberFromName(name){ var months={jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12'}; return months[s(name).slice(0,3).toLowerCase()]||''; }
      function isValidIsoDateParts(year, month, day){ var y=Number(year),m=Number(month),d=Number(day),date=new Date(Date.UTC(y,m-1,d)); return date.getUTCFullYear()===y&&date.getUTCMonth()===(m-1)&&date.getUTCDate()===d; }
      function normalizeDateYear(year){ var y=s(year); if(/^\d{2}$/.test(y)) return String(Number(y)>=70?1900+Number(y):2000+Number(y)); return y; }
      function isoFromDateParts(year, month, day){ var y=normalizeDateYear(year),m=padDatePart(month),d=padDatePart(day); return isValidIsoDateParts(y,m,d)?(y+'-'+m+'-'+d):''; }
      function excelSerialToIso(value){ var serial=Number(value); if(!isFinite(serial)||serial<=0) return ''; var whole=Math.floor(serial); var utc=(whole-25569)*86400000; var date=new Date(utc); if(!isFinite(date.getTime())) return ''; return isoFromDateParts(date.getUTCFullYear(),date.getUTCMonth()+1,date.getUTCDate()); }
      function parseDate(v){ var iso=toIsoInputDate(v); return iso?new Date(iso+'T00:00:00').getTime():8640000000000000; }
      function sortableRowDateValue(row){
        var iso=toIsoInputDate(row&&row['Date']);
        return iso ? new Date(iso+'T00:00:00').getTime() : -8640000000000000;
      }
      function compareRowsNewestFirst(a,b){
        var leftIso=toIsoInputDate(a&&a['Date']),rightIso=toIsoInputDate(b&&b['Date']),da=leftIso?new Date(leftIso+'T00:00:00').getTime():null,db=rightIso?new Date(rightIso+'T00:00:00').getTime():null;
        if(da==null&&db==null){
          var leftEmpty=Number(a&&a.__rowId),rightEmpty=Number(b&&b.__rowId);
          if(isFinite(leftEmpty)&&isFinite(rightEmpty)&&leftEmpty!==rightEmpty) return rightEmpty-leftEmpty;
          return s(b&&b['Job No']).localeCompare(s(a&&a['Job No']),undefined,{numeric:true});
        }
        if(da==null) return -1;
        if(db==null) return 1;
        if(da!==db) return db-da;
        var left=Number(a&&a.__rowId),right=Number(b&&b.__rowId);
        if(isFinite(left)&&isFinite(right)&&left!==right) return right-left;
        return s(b&&b['Job No']).localeCompare(s(a&&a['Job No']),undefined,{numeric:true});
      }
      function compareRowsOldestFirst(a,b){
        var da=parseDate(a&&a['Date']),db=parseDate(b&&b['Date']);
        if(da!==db) return da-db;
        return (Number(a&&a.__rowId)||0)-(Number(b&&b.__rowId)||0);
      }
      function todayIsoDate(){ var now=new Date(); return now.getFullYear()+'-'+padDatePart(now.getMonth()+1)+'-'+padDatePart(now.getDate()); }
      function toDisplayDate(v){ var m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(s(v)); if(!m) return s(v); var months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']; return m[3]+'/'+months[(+m[2])-1]+'/'+m[1]; }
      function toIsoInputDate(v){ var src=s(v),m,iso=''; if(!src) return ''; src=src.replace(/[T\s]+\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?$/,''); if(/^(\d{4})-(\d{2})-(\d{2})$/.test(src)) return src; if(/^\d+(?:\.\d+)?$/.test(src)){ iso=excelSerialToIso(src); if(iso) return iso; } m=/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/.exec(src); if(m) return isoFromDateParts(m[1],m[2],m[3]); m=/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2}|\d{4})$/.exec(src); if(m){ var first=Number(m[1]),second=Number(m[2]),day=first,month=second; if(first<=12&&second>12){ day=second; month=first; } return isoFromDateParts(m[3],month,day); } m=/^(\d{1,2})[\s\/.-]+([A-Za-z]{3,9})[\s\/.-]+(\d{2}|\d{4})$/.exec(src); if(m){ var namedMonth=monthNumberFromName(m[2]); if(namedMonth) return isoFromDateParts(m[3],namedMonth,m[1]); } m=/^([A-Za-z]{3,9})[\s\/.-]+(\d{1,2})(?:,)?[\s\/.-]+(\d{2}|\d{4})$/.exec(src); if(m){ var leadingMonth=monthNumberFromName(m[1]); if(leadingMonth) return isoFromDateParts(m[3],leadingMonth,m[2]); } return ''; }
      function formatDateDisplay(v){ var iso=toIsoInputDate(v); return iso?toDisplayDate(iso):s(v); }
      function parseChapterValue(raw){ var value=s(raw),parts=value.split(' - '); return {chapter:s(parts.shift()),chapterDesc:s(parts.join(' - '))}; }
      function chapterDescriptionForCode(chapter){
        var code=s(chapter);
        if(!code) return '';
        for(var i=0;i<CHAPTER_OPTIONS.length;i++){
          var parsed=parseChapterValue(CHAPTER_OPTIONS[i]);
          if(parsed.chapter===code&&parsed.chapterDesc) return parsed.chapterDesc;
        }
        return '';
      }
      function completeChapterParts(chapter, chapterDesc){
        var code=s(chapter),desc=s(chapterDesc);
        if(code&&!desc) desc=chapterDescriptionForCode(code);
        return { chapter:code, chapterDesc:desc };
      }
      function applyRowChapterReference(row){
        if(!row) return { chapter:'', chapterDesc:'' };
        var completed=completeChapterParts(row['Chapter'],row['Chapter Description']);
        row['Chapter']=completed.chapter;
        if(!referenceOnlySaveEnabled()||s(row['Chapter Description'])) row['Chapter Description']=completed.chapterDesc;
        return completed;
      }
      function syncRowsToReferenceFillMode(){
        for(var i=0;i<rows.length;i++){
          var row=rows[i];
          if(!row) continue;
          if(referenceOnlySaveEnabled()){
            var chapter=s(row['Chapter']),referenceDesc=chapterDescriptionForCode(chapter),supervisorRecord=null;
            if(!row.__manualChapterDescription&&s(row['Chapter Description'])===referenceDesc) row['Chapter Description']='';
            supervisorRecord=supervisorRecordForId(row[SUPERVISOR_ID_FIELD])||(!s(row[SUPERVISOR_ID_FIELD])&&s(row['Approval Name'])?supervisorRecordFor(row['Approval Name']):null);
            if(supervisorRecord&&!row.__manualApprovalLicenceNo&&s(row['Aprroval Licence No.'])===s(supervisorRecord.licence)) row['Aprroval Licence No.']='';
          } else {
            applyRowReferenceData(row);
          }
        }
      }
      function workbookDateValue(row){ return row&&row.__dateDirty?row['Date']:(s(row&&row.__rawDate)||s(row&&row['Date'])); }
      function workbookSavedDateValue(row){
        var iso=toIsoInputDate(workbookDateValue(row));
        return iso?(iso.replace(/-/g,'/')+' 00:00'):'';
      }
      function manualReferenceFieldFlag(row, field){
        if(!row) return false;
        if(field==='Chapter Description') return !!row.__manualChapterDescription;
        if(field==='Approval Name') return !!row.__manualApprovalName;
        if(field==='Approval stamp') return !!row.__manualApprovalStamp;
        if(field==='Aprroval Licence No.') return !!row.__manualApprovalLicenceNo;
        return false;
      }
      function workbookSavedChapterDescriptionValue(row){
        var chapter=s(row&&row['Chapter']),desc=s(row&&row['Chapter Description']),referenceDesc=chapterDescriptionForCode(chapter);
        if(!referenceOnlySaveEnabled()) return desc||referenceDesc;
        if(manualReferenceFieldFlag(row,'Chapter Description')) return desc;
        if(chapter&&desc&&desc===referenceDesc) return '';
        return desc;
      }
      function workbookSavedSupervisorFieldValue(row, field){
        var value=s(row&&row[field]),supervisorId=s(row&&row[SUPERVISOR_ID_FIELD]),record=supervisorId?supervisorRecordForId(supervisorId):null,referenceValue='';
        if(!record) return value;
        if(field==='Approval Name') referenceValue=s(record.name);
        else if(field==='Approval stamp') referenceValue=s(record.stamp);
        else if(field==='Aprroval Licence No.') referenceValue=s(record.licence);
        if(!referenceOnlySaveEnabled()) return value||referenceValue;
        if(manualReferenceFieldFlag(row,field)) return value;
        return value&&value!==referenceValue ? value : '';
      }
      function workbookSavedFieldValue(row, field){
        if(field==='Date') return workbookSavedDateValue(row);
        if(field==='Chapter Description') return workbookSavedChapterDescriptionValue(row);
        if(field==='Approval Name'||field==='Approval stamp'||field==='Aprroval Licence No.') return workbookSavedSupervisorFieldValue(row,field);
        return s(row&&row[field]);
      }
      function normalizeLoadedRow(row){ var rawDate=s(row&&row['Date']); row['Date']=formatDateDisplay(rawDate); row['Flags']=serializeFlagSelection(row&&row['Flags']); row[SUPERVISOR_ID_FIELD]=s(row&&row[SUPERVISOR_ID_FIELD]); row.__manualChapterDescription=!!s(row&&row['Chapter Description']); row.__manualApprovalName=!!s(row&&row['Approval Name']); row.__manualApprovalStamp=!!s(row&&row['Approval stamp']); row.__manualApprovalLicenceNo=!!s(row&&row['Aprroval Licence No.']); row.__rawDate=rawDate; row.__dateDirty=false; return row; }
      function clearWorkbookState(){ rows=normalizeRows([]); AIRCRAFT_GROUP_ROWS=[]; AIRCRAFT_MAP=Object.create(null); CHAPTER_OPTIONS=defaultChapterOptions(); FLAG_RECORDS=defaultFlagRecords(); LOG_OWNER_INFO={ name:'', signature:'', stamp:'' }; APP_VIEW_SETTINGS=cloneAppViewSettings(DEFAULT_APP_VIEW_SETTINGS); rebuildSupervisorState([]); activeFilters=emptyFilterState(); draftFilters=emptyFilterState(); applySearchQuery(''); markSharedDatalistsDirty(); settingsDirty=false; resetSavedLogbookState(); syncMindMapButtonVisibility(); }

      // ---- Row model ----
      function emptyLogRow(type, chapter, chapterDesc){ return {__rowId:nextRowId(),'Aircraft Type':s(type),'A/C Reg':'','Chapter':s(chapter),'Chapter Description':s(chapterDesc),'Date':'','Job No':'','FAULT':'','Task Detail':'','Rewriten for cap741':'','Flags':'',[SUPERVISOR_ID_FIELD]:'','Approval Name':'','Approval stamp':'','Aprroval Licence No.':'','Signed':'',__manualChapterDescription:false,__manualApprovalName:false,__manualApprovalStamp:false,__manualApprovalLicenceNo:false,__trackedComparablePresent:false}; }
      function rowHasEntryContent(row){ return !!(s(row['Date'])||s(row['A/C Reg'])||s(row['Job No'])||s(row['FAULT'])||s(row['Task Detail'])||s(row['Rewriten for cap741'])||s(row['Flags'])||s(row[SUPERVISOR_ID_FIELD])||s(row['Approval Name'])||s(row['Approval stamp'])||s(row['Aprroval Licence No.'])); }
      function rowHasWorkbookContent(row){ return !!(rowHasEntryContent(row)||s(row['Aircraft Type'])||s(row['Chapter'])||s(row['Chapter Description'])); }
      function nonEmptyRows(list){ var out=[]; for(var i=0;i<(list||[]).length;i++){ if(rowHasEntryContent(list[i]||{})) out.push(list[i]); } return out; }
      function workbookContentRows(list){ var out=[]; for(var i=0;i<(list||[]).length;i++){ if(rowHasWorkbookContent(list[i]||{})) out.push(list[i]); } return out; }
      function normalizeRows(list){ list=workbookContentRows(list); rowsById=Object.create(null); var max=-1; for(var i=0;i<list.length;i++){ var id=Number(list[i].__rowId); if(!isFinite(id)||id<0) id=i; list[i].__rowId=id; list[i]['Signed']=isRowSigned(list[i])?'true':''; if(!isFinite(Number(list[i].__signedSlot))) list[i].__signedSlot=-1; list[i].__trackedComparablePresent=rowHasWorkbookContent(list[i]); rowsById[String(id)]=list[i]; if(id>max) max=id; } nextRowIdValue=max+1; return list; }
      function appendRows(list){ for(var i=0;i<list.length;i++){ var row=list[i]; var id=Number(row.__rowId); if(!isFinite(id)||id<0) id=nextRowId(); if(id>=nextRowIdValue) nextRowIdValue=id+1; row.__rowId=id; row.__trackedComparablePresent=rowHasWorkbookContent(row); rowsById[String(id)]=row; rows.push(row); } rebuildDataDirtyTracking(); }
      function nextRowId(){ return nextRowIdValue++; }
      function rowById(id){ return rowsById[String(id)]||null; }
      function removeRowById(id){ var key=String(id); delete rowsById[key]; for(var i=rows.length-1;i>=0;i--){ if(String(rows[i].__rowId)===key){ rows.splice(i,1); break; } } updateRemovedRowDirtyState(key); }
      function rowsByGroupKey(key){ var out=[],mode=currentPageGrouping(); for(var i=0;i<rows.length;i++){ var row=rows[i]; if(rowPageGroupingKey(row,mode)===key) out.push(row); } return out; }
      function aircraftReferenceRecordForReg(reg){
        reg=s(reg).toUpperCase();
        if(!reg) return null;
        for(var i=0;i<(AIRCRAFT_GROUP_ROWS||[]).length;i++){
          var item=AIRCRAFT_GROUP_ROWS[i]||{};
          if(s(item.reg).toUpperCase()===reg) return item;
        }
        return null;
      }
      function aircraftGroupForType(type){
        type=s(type);
        if(!type) return '';
        for(var i=0;i<(AIRCRAFT_GROUP_ROWS||[]).length;i++){
          var item=AIRCRAFT_GROUP_ROWS[i]||{};
          if(s(item.type)===type && s(item.group)) return s(item.group);
        }
        return '';
      }
      function upsertAircraftReferenceRecord(reg, type, preferredGroup){
        reg=s(reg).toUpperCase();
        type=s(type);
        preferredGroup=s(preferredGroup);
        if(!reg || !type) return null;
        var existing=aircraftReferenceRecordForReg(reg);
        if(existing){
          existing.type=type;
          if(preferredGroup) existing.group=preferredGroup;
          else if(!s(existing.group)) existing.group=aircraftGroupForType(type);
          AIRCRAFT_MAP[reg]=type;
          markSharedDatalistsDirty();
          return existing;
        }
        var record={ group:preferredGroup||aircraftGroupForType(type), reg:reg, type:type };
        AIRCRAFT_GROUP_ROWS.push(record);
        AIRCRAFT_MAP[reg]=type;
        markSharedDatalistsDirty();
        return record;
      }
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
      function fullLogbookText(){ var header=LOG_HEADERS.join('\t'),body=workbookContentRows(rows).map(tsvLineFromRow).join('\r\n'); return header+'\r\n'+body+(body?'\r\n':''); }
      function comparableRowSignature(row){ return rowHasWorkbookContent(row||{}) ? tsvLineFromRow(row) : ''; }
      function savedComparableRowSignature(row){ var key=String(row&&row.__rowId); return Object.prototype.hasOwnProperty.call(savedLogbookRowSignatures,key)?savedLogbookRowSignatures[key]:''; }
      function comparableRowOrder(){ var ids=[]; for(var i=0;i<rows.length;i++) if(rowHasWorkbookContent(rows[i]||{})) ids.push(String(rows[i].__rowId)); return ids.join('|'); }
      function hasDirtyLogbookRows(){ for(var key in dirtyLogbookRowIds){ if(Object.prototype.hasOwnProperty.call(dirtyLogbookRowIds,key)) return true; } return false; }
      function syncDataDirtyFlag(){ dataDirty=comparableOrderDirty||hasDirtyLogbookRows(); }
      function refreshComparableOrderDirty(){ comparableOrderDirty=comparableRowOrder()!==savedLogbookRowOrder; syncDataDirtyFlag(); }
      function updateRowDirtyState(row, deferSync){
        if(!row||row.__rowId==null) return false;
        var key=String(row.__rowId),signature=comparableRowSignature(row),wasPresent=!!row.__trackedComparablePresent,isPresent=!!signature,savedSignature=Object.prototype.hasOwnProperty.call(savedLogbookRowSignatures,key)?savedLogbookRowSignatures[key]:'';
        row.__trackedComparablePresent=isPresent;
        if((signature||'')===(savedSignature||'')) delete dirtyLogbookRowIds[key];
        else dirtyLogbookRowIds[key]=true;
        if(deferSync===true) return wasPresent!==isPresent;
        if(wasPresent!==isPresent) refreshComparableOrderDirty();
        else syncDataDirtyFlag();
        return wasPresent!==isPresent;
      }
      function updateRowsDirtyState(list){
        var structureChanged=false;
        for(var i=0;i<(list||[]).length;i++) if(updateRowDirtyState(list[i],true)) structureChanged=true;
        if(structureChanged) refreshComparableOrderDirty();
        else syncDataDirtyFlag();
      }
      function rebuildDataDirtyTracking(){
        var currentSeen=Object.create(null);
        dirtyLogbookRowIds=Object.create(null);
        for(var i=0;i<rows.length;i++){
          var row=rows[i],key=String(row.__rowId),signature=comparableRowSignature(row),savedSignature=Object.prototype.hasOwnProperty.call(savedLogbookRowSignatures,key)?savedLogbookRowSignatures[key]:'';
          row.__trackedComparablePresent=!!signature;
          currentSeen[key]=true;
          if((signature||'')!==(savedSignature||'')) dirtyLogbookRowIds[key]=true;
        }
        for(var savedKey in savedLogbookRowSignatures){
          if(Object.prototype.hasOwnProperty.call(savedLogbookRowSignatures,savedKey)&&savedLogbookRowSignatures[savedKey]&&!currentSeen[savedKey]) dirtyLogbookRowIds[savedKey]=true;
        }
        comparableOrderDirty=comparableRowOrder()!==savedLogbookRowOrder;
        syncDataDirtyFlag();
      }
      function resetSavedLogbookState(){
        var ids=[];
        savedLogbookRowSignatures=Object.create(null);
        dirtyLogbookRowIds=Object.create(null);
        for(var i=0;i<rows.length;i++){
          var row=rows[i],key=String(row.__rowId),signature=comparableRowSignature(row);
          row.__trackedComparablePresent=!!signature;
          if(!signature) continue;
          savedLogbookRowSignatures[key]=signature;
          ids.push(key);
        }
        savedLogbookRowOrder=ids.join('|');
        comparableOrderDirty=false;
        dataDirty=false;
        lastSavedLogbookText=fullLogbookText();
      }
      function updateRemovedRowDirtyState(rowId){
        var key=String(rowId),savedSignature=Object.prototype.hasOwnProperty.call(savedLogbookRowSignatures,key)?savedLogbookRowSignatures[key]:'';
        if(savedSignature) dirtyLogbookRowIds[key]=true;
        else delete dirtyLogbookRowIds[key];
        refreshComparableOrderDirty();
      }

      // ---- Supervisor helpers ----
      function supervisorLookupKey(kind, value){ return kind+'::'+s(value).toLowerCase(); }
      function supervisorRecordFor(value){
        var key=s(value);
        if(!key) return null;
        return SUPERVISOR_LOOKUP[supervisorLookupKey('name',key)]||SUPERVISOR_LOOKUP[supervisorLookupKey('label',key)]||SUPERVISOR_LOOKUP[supervisorLookupKey('id',key)]||null;
      }
      function supervisorRecordForId(value){
        var key=s(value);
        return key?(SUPERVISOR_LOOKUP[supervisorLookupKey('id',key)]||null):null;
      }
      function normalizeSupervisorValue(value){
        var raw=s(value),record=supervisorRecordFor(value);
        if(record) return {id:s(record.id),name:record.name,stamp:record.stamp,licence:record.licence};
        var parts=raw.split('|').map(function(x){ return s(x); }).filter(Boolean);
        if(parts.length) return {id:'',name:parts[0]||'',stamp:parts[1]||'',licence:parts[2]||''};
        return {id:'',name:raw,stamp:'',licence:''};
      }
      function extractSupervisorParts(value){ var resolved=normalizeSupervisorValue(value); return {id:resolved.id,name:resolved.name,stamp:resolved.stamp,licence:resolved.licence}; }
      function applyRowSupervisorReference(row){
        if(!row) return null;
        var record=supervisorRecordForId(row[SUPERVISOR_ID_FIELD]);
        if(!record&&s(row['Approval Name'])) record=supervisorRecordFor(row['Approval Name']);
        if(record){
          if(!s(row[SUPERVISOR_ID_FIELD])) row[SUPERVISOR_ID_FIELD]=s(record.id);
          if(!s(row['Approval Name'])) row['Approval Name']=s(record.name);
          if(!s(row['Approval stamp'])) row['Approval stamp']=s(record.stamp);
          if(!s(row['Aprroval Licence No.'])&&!referenceOnlySaveEnabled()) row['Aprroval Licence No.']=s(record.licence);
          return record;
        }
        return null;
      }
      function applyRowReferenceData(row){ applyRowChapterReference(row); applyRowSupervisorReference(row); return row; }
      function fillSupervisorFields(nameInput, licenceInput, row){
        if(!nameInput) return null;
        var resolved=normalizeSupervisorValue(nameInput.value);
        if(resolved.name) nameInput.value=resolved.name;
        syncFieldInputViewState(nameInput);
        if(licenceInput){
          if(resolved.licence&&!referenceOnlySaveEnabled()) licenceInput.value=resolved.licence;
          var licenceView=licenceInput.nextElementSibling&&licenceInput.nextElementSibling.classList&&licenceInput.nextElementSibling.classList.contains('field-input-view')?licenceInput.nextElementSibling:null;
          if(licenceView) licenceView.textContent=resolved.licence||'';
          syncFieldInputViewState(licenceInput);
        }
        if(row){
          row[SUPERVISOR_ID_FIELD]=resolved.id||'';
          row['Approval Name']=resolved.name;
          row['Approval stamp']=resolved.stamp;
          row.__manualApprovalName=!!resolved.name;
          row.__manualApprovalStamp=false;
          if(resolved.licence&&!referenceOnlySaveEnabled()) row['Aprroval Licence No.']=resolved.licence;
        }
        return resolved;
      }
      function setRowSupervisorFields(row, nameValue, licenceValue){
        var resolved=normalizeSupervisorValue(nameValue);
        row[SUPERVISOR_ID_FIELD]=resolved.id||'';
        row['Approval Name']=resolved.name;
        row['Approval stamp']=resolved.stamp;
        row['Aprroval Licence No.']=licenceValue||(!referenceOnlySaveEnabled()?resolved.licence:'')||'';
        row.__manualApprovalName=!!resolved.name;
        row.__manualApprovalStamp=false;
        row.__manualApprovalLicenceNo=!!s(licenceValue);
      }
      function isRowSigned(row){ var value=normalizedText(row&&row['Signed']); return value==='yes'||value==='signed'||value==='true'||value==='1'; }
      function signedSlotFor(row){ var slot=Number(row&&row.__signedSlot); return isFinite(slot)&&slot>=0?slot:-1; }
      function setRowSignedState(row, signed, slotStart){ if(!row) return; row['Signed']=signed?'true':''; row.__signedSlot=signed&&isFinite(slotStart)&&slotStart>=0?slotStart:-1; }
      function initializeSignedSlots(list){
        function nextSlotFit(start, units){ var offset=start%PAGE_SLOTS; if(offset&&offset+units>PAGE_SLOTS) return start+(PAGE_SLOTS-offset); return start; }
        var groups=groupRows(list||[]);
        for(var i=0;i<groups.length;i++){
          var groupRowsSorted=(groups[i].rows||[]).slice().sort(compareRowsOldestFirst),cursor=0;
          for(var j=0;j<groupRowsSorted.length;j++){
            var row=groupRowsSorted[j],slot=signedSlotFor(row),units=unitsFor(row);
            if(slot<0&&isRowSigned(row)){
              cursor=nextSlotFit(cursor,units);
              row.__signedSlot=cursor;
            }
            cursor=nextSlotFit(cursor,units)+units;
          }
        }
      }

      // ---- Aircraft / Chapter options HTML ----
      function safeIdPart(value){ return s(value).toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-+|-+$/g,'')||'group'; }
      function aircraftLabel(row){ var type=s(row&&row['Aircraft Type']),reg=s(row&&row['A/C Reg']).toUpperCase(); return type||(AIRCRAFT_MAP[reg]||''); }
      function aircraftOptionsHtml(){ var regs=Object.keys(AIRCRAFT_MAP).sort(),html=''; for(var i=0;i<regs.length;i++) html+='<option value="'+esc(regs[i])+'"></option>'; return html; }
      function aircraftOptionsHtmlForType(type){ type=s(type); if(!type) return aircraftOptionsHtml(); if(Object.prototype.hasOwnProperty.call(aircraftOptionsByTypeCache,type)) return aircraftOptionsByTypeCache[type]; var regs=[]; for(var reg in AIRCRAFT_MAP){ if(Object.prototype.hasOwnProperty.call(AIRCRAFT_MAP,reg)&&AIRCRAFT_MAP[reg]===type) regs.push(reg); } regs.sort(); if(!regs.length) return aircraftOptionsByTypeCache[type]=aircraftOptionsHtml(); var html=''; for(var i=0;i<regs.length;i++) html+='<option value="'+esc(regs[i])+'"></option>'; aircraftOptionsByTypeCache[type]=html; return html; }
      function aircraftTypeOptionsHtml(){ var seen={},vals=[]; for(var k in AIRCRAFT_MAP){ if(Object.prototype.hasOwnProperty.call(AIRCRAFT_MAP,k)&&!seen[AIRCRAFT_MAP[k]]){ seen[AIRCRAFT_MAP[k]]=true; vals.push(AIRCRAFT_MAP[k]); } } vals.sort(); var html=''; for(var i=0;i<vals.length;i++) html+='<option value="'+esc(vals[i])+'"></option>'; return html; }
      function chapterOptionsHtml(){ var html=''; for(var i=0;i<CHAPTER_OPTIONS.length;i++) html+='<option value="'+esc(CHAPTER_OPTIONS[i])+'"></option>'; return html; }
      function supervisorOptionsHtml(){ var html=''; for(var i=0;i<SUPERVISOR_OPTIONS.length;i++) html+='<option value="'+esc(SUPERVISOR_OPTIONS[i])+'"></option>'; return html; }
      function sharedDatalistsHtml(){ if(!sharedDatalistsCache) sharedDatalistsCache='<datalist id="aircraft-reg-list">'+aircraftOptionsHtml()+'</datalist><datalist id="aircraft-type-list">'+aircraftTypeOptionsHtml()+'</datalist><datalist id="chapter-list">'+chapterOptionsHtml()+'</datalist><datalist id="supervisor-list">'+supervisorOptionsHtml()+'</datalist>'; return sharedDatalistsCache; }
      function aircraftRegListIdForGroup(group){ return 'aircraft-reg-list-'+safeIdPart(s(group&&group.key)||((s(group&&group.type)+'-'+s(group&&group.chapter))||'group')); }
      function groupAircraftRegDatalistHtml(group){
        var options='';
        if(group&&group.mode===PAGE_GROUPING_GROUP) options=aircraftOptionsHtmlForGroupLabel(group.groupLabel);
        else options=aircraftOptionsHtmlForTypes(((group&&group.typeList)||[]).filter(Boolean));
        return '<datalist id="'+aircraftRegListIdForGroup(group)+'">'+options+'</datalist>';
      }
      function renderedGroupDatalistsHtml(renderedGroups){ var html=[],seen=Object.create(null); for(var i=0;i<(renderedGroups||[]).length;i++){ var group=renderedGroups[i]&&renderedGroups[i].group,key=s(group&&group.key); if(!key||seen[key]) continue; seen[key]=true; html.push(groupAircraftRegDatalistHtml(group)); } return html.join(''); }
      function usedAircraftTypes(){ var seen={},vals=[]; for(var i=0;i<rows.length;i++){ var type=aircraftLabel(rows[i]); if(type&&!seen[type]){ seen[type]=true; vals.push(type); } } vals.sort(); return vals; }
      function usedAircraftTypeOptionsHtml(){ var vals=usedAircraftTypes(); if(!vals.length) return aircraftTypeOptionsHtml(); var html=''; for(var i=0;i<vals.length;i++) html+='<option value="'+esc(vals[i])+'"></option>'; return html; }
      function modalAircraftTypeListId(){ return 'modal-aircraft-type-list'; }
      function modalAircraftRegListId(){ return 'modal-aircraft-reg-list'; }
      function modalAircraftTypeDatalistHtml(){ return '<datalist id="'+modalAircraftTypeListId()+'">'+aircraftTypeOptionsHtml()+'</datalist>'; }
      function modalAircraftRegDatalistHtml(type){ return '<datalist id="'+modalAircraftRegListId()+'">'+aircraftOptionsHtmlForType(type)+'</datalist>'; }

      // ---- Render helpers ----
      function mainPageTaskText(row){ return s(row['Rewriten for cap741']||row['Task Detail']); }
      function taskTextMeasureSample(){ return pagesEl&&pagesEl.querySelector?pagesEl.querySelector('.sheet td.c-task .task-wrap .task > .editable-cell, .sheet td.c-task .task-wrap .task > .locked-cell'):null; }
      function taskTextMeasureMetrics(){
        if(taskTextMeasureCache) return taskTextMeasureCache;
        var sample=taskTextMeasureSample(),width=TASK_TEXT_MEASURE_WIDTH_FALLBACK_PX,lineHeight=TASK_TEXT_LINE_HEIGHT_FALLBACK_PX,fontFamily='Arial, Helvetica, sans-serif',fontSize='8px',fontWeight='400',fontStyle='normal',fontVariant='normal',letterSpacing='normal';
        if(sample&&window.getComputedStyle){
          var computed=window.getComputedStyle(sample),sampleRect=sample.getBoundingClientRect();
          width=(sampleRect&&sampleRect.width)||sample.clientWidth||width;
          lineHeight=parseFloat(computed.lineHeight)||lineHeight;
          fontFamily=computed.fontFamily||fontFamily;
          fontSize=computed.fontSize||fontSize;
          fontWeight=computed.fontWeight||fontWeight;
          fontStyle=computed.fontStyle||fontStyle;
          fontVariant=computed.fontVariant||fontVariant;
          letterSpacing=computed.letterSpacing||letterSpacing;
        }
        width=Math.max(1,width);
        lineHeight=Math.max(1,lineHeight);
        taskTextMeasureCache={
          width:width,
          lineHeight:lineHeight,
          fontFamily:fontFamily,
          fontSize:fontSize,
          fontWeight:fontWeight,
          fontStyle:fontStyle,
          fontVariant:fontVariant,
          letterSpacing:letterSpacing,
          signature:[Math.round(width*100)/100,lineHeight,fontFamily,fontSize,fontWeight,fontStyle,fontVariant,letterSpacing,ROW_TASK_LINES_PER_SLOT].join('|')
        };
        return taskTextMeasureCache;
      }
      function ensureTaskTextMeasureEl(){
        if(taskTextMeasureEl||!document||!document.body) return taskTextMeasureEl;
        taskTextMeasureEl=document.createElement('div');
        taskTextMeasureEl.setAttribute('aria-hidden','true');
        taskTextMeasureEl.style.position='absolute';
        taskTextMeasureEl.style.left='-9999px';
        taskTextMeasureEl.style.top='0';
        taskTextMeasureEl.style.visibility='hidden';
        taskTextMeasureEl.style.pointerEvents='none';
        taskTextMeasureEl.style.padding='0';
        taskTextMeasureEl.style.border='0';
        taskTextMeasureEl.style.margin='0';
        taskTextMeasureEl.style.whiteSpace='pre-wrap';
        taskTextMeasureEl.style.wordBreak='break-word';
        taskTextMeasureEl.style.overflowWrap='break-word';
        taskTextMeasureEl.style.boxSizing='border-box';
        document.body.appendChild(taskTextMeasureEl);
        return taskTextMeasureEl;
      }
      function taskLinesFor(text){
        var t=s(text);
        if(!t) return 1;
        var measureEl=ensureTaskTextMeasureEl(),metrics=taskTextMeasureMetrics();
        if(!measureEl) return Math.max(1,Math.ceil(t.length/64));
        measureEl.style.width=Math.max(1,Math.round(metrics.width))+'px';
        measureEl.style.lineHeight=metrics.lineHeight+'px';
        measureEl.style.fontFamily=metrics.fontFamily;
        measureEl.style.fontSize=metrics.fontSize;
        measureEl.style.fontWeight=metrics.fontWeight;
        measureEl.style.fontStyle=metrics.fontStyle;
        measureEl.style.fontVariant=metrics.fontVariant;
        measureEl.style.letterSpacing=metrics.letterSpacing;
        measureEl.textContent=t;
        var measuredHeight=measureEl.scrollHeight||measureEl.getBoundingClientRect().height||metrics.lineHeight;
        measureEl.textContent='';
        return Math.max(1,Math.round(measuredHeight/metrics.lineHeight));
      }
      function taskLineCountForRow(row){
        var key=mainPageTaskText(row||{}),metrics=taskTextMeasureMetrics(),cacheKey=metrics.signature+'||'+key;
        if(row&&row.__taskLineCountCacheKey===cacheKey&&typeof row.__taskLineCount==='number') return row.__taskLineCount;
        var lineCount=taskLinesFor(key);
        if(row){
          row.__taskLineCountCacheKey=cacheKey;
          row.__taskLineCount=lineCount;
        }
        return lineCount;
      }
      function unitsFor(row){
        var key=mainPageTaskText(row||{}),metrics=taskTextMeasureMetrics(),cacheKey=metrics.signature+'||'+key;
        if(row&&row.__unitsCacheKey===cacheKey&&typeof row.__unitsCacheValue==='number') return row.__unitsCacheValue;
        var lineCount=taskLineCountForRow(row),value=Math.max(1,Math.min(PAGE_SLOTS,Math.ceil(lineCount/Math.max(1,ROW_TASK_LINES_PER_SLOT))));
        if(row){
          row.__unitsCacheKey=cacheKey;
          row.__unitsCacheValue=value;
        }
        return value;
      }
      function taskWrapClassName(row){
        var classes=['task-wrap'],lineCount=taskLineCountForRow(row),units=unitsFor(row);
        if(isRowSigned(row)) classes.push('task-wrap-locked');
        if(units===1&&lineCount===ROW_TASK_LINES_PER_SLOT) classes.push('task-wrap-four-lines');
        return classes.join(' ');
      }
      function dotsInputSize(value){ return Math.max(8,Math.min(56,s(value).length+1)); }
      function dotsInputFontSize(value){
        var length=s(value).length;
        if(length<=16) return 11;
        if(length>=40) return 9;
        return Math.round((11-((length-16)*(2/24)))*100)/100;
      }
      function renderDotsInput(value, extraAttrs){
        var text=s(value),size=dotsInputSize(text),fontSize=dotsInputFontSize(text);
        return '<span class="dots-value"><input class="field-input dots-input" type="text" size="'+size+'" value="'+esc(text||'')+'" style="--dots-input-font-size:'+fontSize+'px;"'+(extraAttrs||'')+'></span>';
      }
      function renderDotsStatic(value){ return '<span>'+(esc(value)||'&nbsp;')+'</span>'; }
      function renderPageHeaderDots(value, field, listId, editable){
        return '<span class="dots-line'+(editable?' editable-dots-line':'')+'">'+(editable?renderDotsInput(value,' data-group-field="'+field+'"'+(listId?' list="'+listId+'"':'')):renderDotsStatic(value))+'</span>';
      }
      function syncDotsInputSize(input){
        if(!input||!input.classList||!input.classList.contains('dots-input')) return;
        var value=valueOf(input),fontSize=dotsInputFontSize(value);
        input.size=dotsInputSize(value);
        input.style.setProperty('--dots-input-font-size',fontSize+'px');
      }
      function syncFieldInputViewState(input){
        var wrap=input&&input.parentElement,view=input&&input.nextElementSibling;
        if(!wrap||!wrap.classList||!wrap.classList.contains('field-input-view-wrap')) return;
        wrap.classList.toggle('has-input-value',!!s(valueOf(input)));
        wrap.classList.toggle('has-view-value',!!s(view&&view.textContent));
      }
      function editableTextInput(field, rowId, value, placeholder, extraClass, listId){ return '<input class="field-input '+(extraClass||'')+'" type="text"'+(listId?' list="'+listId+'"':'')+' data-row-id="'+rowId+'" data-edit-field="'+field+'" value="'+esc(value||'')+'"'+(placeholder?' placeholder="'+esc(placeholder)+'"':'')+'>';}
      function editableTextInputWithView(field, rowId, value, placeholder, extraClass, listId, viewValue){
        var text=s(value),viewText=s(viewValue),classes=['field-input-view-wrap'];
        if(extraClass) classes.push(s(extraClass));
        if(text) classes.push('has-input-value');
        if(viewText) classes.push('has-view-value');
        return '<span class="'+esc(classes.join(' '))+'"><input class="field-input ref-view-input" type="text"'+(listId?' list="'+listId+'"':'')+' data-row-id="'+rowId+'" data-edit-field="'+field+'" value="'+esc(text)+'"'+(placeholder?' placeholder="'+esc(placeholder)+'"':'')+'><span class="field-input-view" aria-hidden="true">'+esc(viewText||'')+'</span></span>';
      }
      function editableCell(field, rowId, value, cls){ return '<div class="editable-cell '+(cls||'')+'" contenteditable="true" data-row-id="'+rowId+'" data-edit-field="'+field+'">'+(esc(value)||'&nbsp;')+'</div>'; }
      function staticTextCell(value, cls){ return '<div class="locked-cell'+(cls?' '+cls:'')+'">'+(esc(value)||'&nbsp;')+'</div>'; }
      function placeholderCellHtml(extraClass){ return '<div class="'+(extraClass||'placeholder-cell')+'">&nbsp;</div>'; }
      function rowLockIconSvg(signed){
        if(signed) return '<svg viewBox="0 0 24 24" class="row-lock-icon" aria-hidden="true" focusable="false"><path d="M7 10V8a5 5 0 1 1 10 0v2h1a2 2 0 0 1 2 2v7a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-7a2 2 0 0 1 2-2h1Zm2 0h6V8a3 3 0 1 0-6 0v2Z" fill="currentColor"/></svg>';
        return '<svg viewBox="0 0 24 24" class="row-lock-icon" aria-hidden="true" focusable="false"><path d="M17 10h1a2 2 0 0 1 2 2v7a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-7a2 2 0 0 1 2-2h8V8a2 2 0 1 0-4 0v1H8V8a4 4 0 1 1 8 0v2Zm1 2H6v7h12v-7Z" fill="currentColor"/></svg>';
      }
      function signToggleButtonHtml(row){
        var signed=isRowSigned(row);
        return '<button class="row-lock'+(signed?' is-signed':'')+'" type="button" data-toggle-signed="1" data-row-id="'+row.__rowId+'" aria-label="'+(signed?'Unlock signed row':'Lock and sign row')+'" title="'+(signed?'Unlock row':'Lock row')+'">'+rowLockIconSvg(signed)+'</button>';
      }
      function clearRowButtonHtml(row){
        return '<button class="sup-clear" type="button" data-clear-supervisor="1" data-row-id="'+row.__rowId+'" aria-label="Clear row" title="Clear row">Clear</button>';
      }
      function renderRowFlagsSummary(row){
        var flags=flagSummaryRecords(row),html=[],visibleCount=Math.min(flags.length,3),extraCount=flags.length-visibleCount,i,item;
        if(!flags.length) return '';
        for(i=0;i<visibleCount;i++){
          item=flags[i]||{};
          html.push('<span class="sup-flag-pill" title="'+esc(item.flag)+'"><span class="sup-flag-dot" style="background-color:'+esc(s(item.color)||'#7f93a1')+'"></span><span class="sup-flag-pill-text">'+esc(flagBadgeLetter(item.flag))+'</span></span>');
        }
        if(extraCount>0) html.push('<span class="sup-flag-pill sup-flag-pill-more" title="'+esc(flags.slice(visibleCount).map(function(flag){ return flag.flag; }).join(', '))+'">+'+extraCount+'</span>');
        return '<div class="sup-flags" title="'+esc(flags.map(function(flag){ return flag.flag; }).join(', '))+'">'+html.join('')+'</div>';
      }
      function editableSupervisorCell(row){
        var clearAction=clearRowButtonHtml(row),flagsHtml=renderRowFlagsSummary(row),classes='sup'+(flagsHtml?' has-flags':'');
        if(isRowSigned(row)) return '<div class="'+classes+' sup-locked"><span class="star">*</span>'+staticTextCell(supervisorNameView(row),'name locked-text')+staticTextCell(supervisorLicenceView(row),'licence locked-text')+flagsHtml+clearAction+'</div>';
        return '<div class="'+classes+'"><span class="star">*</span>'+editableTextInputWithView('Approval Name',row.__rowId,row['Approval Name'],'Supervisor','name','supervisor-list',supervisorNameView(row))+editableTextInputWithView('Aprroval Licence No.',row.__rowId,row['Aprroval Licence No.'],'Licence number','licence','',supervisorLicenceView(row))+flagsHtml+clearAction+'</div>';
      }
      function dateCellHtml(row){
        var dateContent=isRowSigned(row)
          ? staticTextCell(formatDateDisplay(row['Date']),'locked-cell date-locked')
          : dateControlHtml(' data-row-id="'+row.__rowId+'" data-edit-field="Date"',DATE_PLACEHOLDER,formatDateDisplay(row['Date']),toIsoInputDate(row['Date']));
        return '<div class="date-cell-wrap">'+signToggleButtonHtml(row)+dateContent+'</div>';
      }
      function taskCellHtml(row){ var wrapClass=taskWrapClassName(row); if(isRowSigned(row)) return '<div class="'+wrapClass+'"><div class="task">'+staticTextCell(mainPageTaskText(row),'task-input locked-text')+'</div></div>'; return '<div class="'+wrapClass+'"><div class="task">'+editableCell('Rewriten for cap741',row.__rowId,mainPageTaskText(row),'task-input')+'</div><button class="task-expand" type="button" data-open-task="1" data-row-id="'+row.__rowId+'" aria-label="Show full task detail">&#x2197;</button></div>'; }
      function blankTaskCellHtml(type, chapter, chapterDesc, groupLabel, regListId, filterReg){ return '<div class="task-wrap"><div class="task">'+blankEditableCell('Rewriten for cap741',type,chapter,chapterDesc,groupLabel,regListId,filterReg)+'</div><button class="task-expand" type="button" data-open-task-new="1" aria-label="Show full task detail">&#x2197;</button></div>'; }
      function dateControlHtml(extraAttrs, placeholder, displayValue, isoValue){ return '<div class="date-entry"><input class="field-input date-text" type="text" data-date-text="1" placeholder="'+(placeholder||DATE_PLACEHOLDER)+'" value="'+esc(displayValue||'')+('"'+extraAttrs)+'><input class="date-native" type="date" data-date-picker="1" value="'+esc(isoValue||'')+'"></div>'; }
      function splitAircraftTypeParts(label){
        label=s(label);
        var idx=label.lastIndexOf(' - ');
        return idx>0 ? { airframe:s(label.slice(0,idx)), engine:s(label.slice(idx+3)) } : { airframe:label, engine:'' };
      }
      function uniqueTextList(list){
        var seen=Object.create(null),out=[],i,value;
        for(i=0;i<(list||[]).length;i++){
          value=s(list[i]);
          if(!value||seen[value]) continue;
          seen[value]=true;
          out.push(value);
        }
        return out;
      }
      function combinedAircraftAirframeLabel(list){
        var airframes=uniqueTextList(list),parts=[],sharedLead='',i,splitAt;
        if(!airframes.length) return '';
        if(airframes.length===1) return airframes[0];
        for(i=0;i<airframes.length;i++){
          splitAt=airframes[i].lastIndexOf(' ');
          if(splitAt<=0) return airframes.join(', ');
          parts.push({ lead:s(airframes[i].slice(0,splitAt)), tail:s(airframes[i].slice(splitAt+1)) });
        }
        sharedLead=parts[0].lead;
        for(i=0;i<parts.length;i++) if(parts[i].lead!==sharedLead||!parts[i].tail) return airframes.join(', ');
        return sharedLead+' '+parts.map(function(item){ return item.tail; }).join(', ');
      }
      function defaultNewAircraftTypeForRows(list, mode){
        var seen=Object.create(null),i,type='',count=0,lastType='';
        mode=mode||currentPageGrouping();
        if(mode===PAGE_GROUPING_TYPE) return s(aircraftLabel(list&&list[0]));
        for(i=0;i<(list||[]).length;i++){
          type=s(aircraftLabel(list[i]));
          if(!type||seen[type]) continue;
          seen[type]=true;
          lastType=type;
          count++;
          if(count>1) return '';
        }
        return lastType;
      }
      function combinedAircraftTypeHeader(list, fallback){
        var types=uniqueTextList(list),parsed=[],engines=[],i,airframeText='',engineText='';
        if(!types.length) return s(fallback)||'';
        if(types.length===1) return types[0];
        for(i=0;i<types.length;i++) parsed.push(splitAircraftTypeParts(types[i]));
        for(i=0;i<parsed.length;i++) if(parsed[i].engine) engines.push(parsed[i].engine);
        airframeText=combinedAircraftAirframeLabel(parsed.map(function(item){ return item.airframe; }));
        engineText=uniqueTextList(engines).join(', ');
        return airframeText+(engineText?', '+engineText:'');
      }
      function rowPageGroupingLabel(row, mode){
        mode=mode||currentPageGrouping();
        if(mode===PAGE_GROUPING_GROUP) return resolvedRowPageGroupLabel(row)||rememberedRowPageGroupLabel(row);
        return s(aircraftLabel(row))||'';
      }
      function rowPageGroupingKey(row, mode){
        mode=mode||currentPageGrouping();
        return mode+'||'+rowPageGroupingLabel(row,mode)+'||'+s(row&&row['Chapter']);
      }
      // Rows are grouped exactly how the printed logbook is grouped: one section
      // per aircraft type + ATA chapter, or aircraft group + ATA chapter.
      function groupRows(list){
        var mode=currentPageGrouping(),map=Object.create(null),i,row,key,bucket,label,typeLabel,out=[],sortLeft='',sortRight='';
        for(i=0;i<(list||[]).length;i++){
          row=list[i];
          label=s(rowPageGroupingLabel(row,mode));
          key=rowPageGroupingKey(row,mode);
          if(!map[key]) map[key]={ key:key, mode:mode, groupLabel:label, chapter:s(row['Chapter']), chapterDesc:rowChapterDescriptionView(row), rows:[], typeList:[], _typeSeen:Object.create(null) };
          bucket=map[key];
          if(!bucket.chapterDesc) bucket.chapterDesc=rowChapterDescriptionView(row);
          bucket.rows.push(row);
          typeLabel=s(aircraftLabel(row));
          if(typeLabel&&!bucket._typeSeen[typeLabel]){
            bucket._typeSeen[typeLabel]=true;
            bucket.typeList.push(typeLabel);
          }
        }
        for(key in map){
          if(!Object.prototype.hasOwnProperty.call(map,key)) continue;
          bucket=map[key];
          bucket.typeList.sort();
          bucket.type=mode===PAGE_GROUPING_GROUP?combinedAircraftTypeHeader(bucket.typeList,''):(bucket.typeList[0]||bucket.groupLabel||'');
          bucket.typeEditable=mode===PAGE_GROUPING_TYPE;
          bucket.defaultNewType=defaultNewAircraftTypeForRows(bucket.rows,mode);
          delete bucket._typeSeen;
          out.push(bucket);
        }
        out.sort(function(a,b){
          sortLeft=mode===PAGE_GROUPING_GROUP?s(a.groupLabel):s(a.type);
          sortRight=mode===PAGE_GROUPING_GROUP?s(b.groupLabel):s(b.type);
          if(sortLeft===sortRight) return a.chapter.localeCompare(b.chapter,undefined,{numeric:true});
          return sortLeft.localeCompare(sortRight,undefined,{numeric:true});
        });
        return out;
      }
      // Each task consumes one or more vertical "slots" on a page, so pagination is
      // based on rendered space rather than raw row count.
      function paginate(list){
        function nextSlotFit(start, units){ var offset=start%PAGE_SLOTS; if(offset&&offset+units>PAGE_SLOTS) return start+(PAGE_SLOTS-offset); return start; }
        var signed=[],unsigned=[],placed=[],cursor=0,unsignedIndex=0,pages=[];
        for(var i=0;i<list.length;i++){
          var entry={row:list[i],units:unitsFor(list[i]),signed:isRowSigned(list[i]),signedSlot:signedSlotFor(list[i])};
          if(entry.signed) signed.push(entry);
          else unsigned.push(entry);
        }
        signed.sort(function(a,b){ if(a.signedSlot!==b.signedSlot) return a.signedSlot-b.signedSlot; return compareRowsOldestFirst(a.row,b.row); });
        unsigned.sort(function(a,b){ return compareRowsOldestFirst(a.row,b.row); });
        function placeUnsignedUntil(target){
          while(unsignedIndex<unsigned.length){
            var next=unsigned[unsignedIndex],start=nextSlotFit(cursor,next.units);
            if(start+next.units>target) break;
            cursor=start;
            placed.push({row:next.row,units:next.units,slotStart:cursor});
            cursor+=next.units;
            unsignedIndex++;
          }
          if(cursor<target) cursor=target;
        }
        for(var j=0;j<signed.length;j++){
          var signedEntry=signed[j],target=signedEntry.signedSlot>=0?signedEntry.signedSlot:cursor,startSlot;
          placeUnsignedUntil(target);
          startSlot=nextSlotFit(Math.max(cursor,target),signedEntry.units);
          placed.push({row:signedEntry.row,units:signedEntry.units,slotStart:startSlot,signed:true});
          cursor=startSlot+signedEntry.units;
        }
        while(unsignedIndex<unsigned.length){
          var freeEntry=unsigned[unsignedIndex++];
          cursor=nextSlotFit(cursor,freeEntry.units);
          placed.push({row:freeEntry.row,units:freeEntry.units,slotStart:cursor});
          cursor+=freeEntry.units;
        }
        for(var k=0;k<placed.length;k++){
          var pageIndex=Math.floor(placed[k].slotStart/PAGE_SLOTS);
          placed[k].pageSlotStart=placed[k].slotStart%PAGE_SLOTS;
          if(!pages[pageIndex]) pages[pageIndex]=[];
          pages[pageIndex].push(placed[k]);
        }
        return pages.filter(Boolean);
      }
      function blankEditableCell(field, type, chapter, chapterDesc, groupLabel, regListId, filterReg){ var common=' data-new-row="1" data-edit-field="'+field+'" data-new-type="'+esc(type)+'" data-new-chapter="'+esc(chapter)+'" data-new-chapter-desc="'+esc(chapterDesc||'')+'" data-new-page-group="'+esc(groupLabel||'')+'" data-new-filter-reg="'+esc(filterReg||'')+'"'; if(field==='Date') return dateControlHtml(common,DATE_PLACEHOLDER); if(field==='A/C Reg') return '<input class="field-input" type="text" list="'+esc(regListId||'aircraft-reg-list')+'" placeholder="G-XXXX"'+common+'>'; if(field==='Job No') return '<input class="field-input" type="text" placeholder="Job No"'+common+'>'; if(field==='Task Detail'||field==='Rewriten for cap741') return '<div class="editable-cell task-input" contenteditable="true"'+common+'></div>'; return '<div class="editable-cell" contenteditable="true"'+common+'></div>'; }
      function blankSupervisorInputWithView(field, type, chapter, chapterDesc, groupLabel, filterReg, placeholder, extraClass, listId){
        var attrs=' data-new-row="1" data-edit-field="'+field+'" data-new-type="'+esc(type)+'" data-new-chapter="'+esc(chapter)+'" data-new-chapter-desc="'+esc(chapterDesc||'')+'" data-new-page-group="'+esc(groupLabel||'')+'" data-new-filter-reg="'+esc(filterReg||'')+'"';
        return '<span class="field-input-view-wrap '+esc(extraClass||'')+'"><input class="field-input ref-view-input" type="text"'+(listId?' list="'+listId+'"':'')+' placeholder="'+esc(placeholder||'')+'"'+attrs+'><span class="field-input-view" aria-hidden="true"></span></span>';
      }
      function blankSupervisorCell(type, chapter, chapterDesc, groupLabel, filterReg){ return '<div class="sup"><span class="star">*</span>'+blankSupervisorInputWithView('Approval Name',type,chapter,chapterDesc,groupLabel,filterReg,'Supervisor','name','supervisor-list')+blankSupervisorInputWithView('Aprroval Licence No.',type,chapter,chapterDesc,groupLabel,filterReg,'Licence number','licence','')+'</div>'; }
      function makeBlankSlot(group, regListId, preserveOnly){ if(preserveOnly) return '<tr class="slot slot-placeholder"><td class="c-date">'+placeholderCellHtml('placeholder-cell')+'</td><td class="c-reg">'+placeholderCellHtml('placeholder-cell')+'</td><td class="c-job">'+placeholderCellHtml('placeholder-cell')+'</td><td class="c-task">'+placeholderCellHtml('placeholder-cell')+'</td><td class="c-sup">'+placeholderCellHtml('placeholder-cell')+'</td></tr>'; var newType=s(group&&group.defaultNewType),groupLabel=group&&group.mode===PAGE_GROUPING_GROUP?s(group.groupLabel):'',filterReg=currentSingleAircraftRegFilter(); return '<tr class="slot"><td class="c-date">'+blankEditableCell('Date',newType,group.chapter,group.chapterDesc,groupLabel,regListId,filterReg)+'</td><td class="c-reg">'+blankEditableCell('A/C Reg',newType,group.chapter,group.chapterDesc,groupLabel,regListId,filterReg)+'</td><td class="c-job">'+blankEditableCell('Job No',newType,group.chapter,group.chapterDesc,groupLabel,regListId,filterReg)+'</td><td class="c-task">'+blankTaskCellHtml(newType,group.chapter,group.chapterDesc,groupLabel,regListId,filterReg)+'</td><td class="c-sup">'+blankSupervisorCell(newType,group.chapter,group.chapterDesc,groupLabel,filterReg)+'</td></tr>'; }
      function makeRows(items, group, preserveOnly){
        var html='',consumed=0,regListId=aircraftRegListIdForGroup(group);
        for(var i=0;i<items.length;i++){
          var item=items[i],row=item.row,signed=isRowSigned(row);
          while(consumed<item.pageSlotStart){ html+=makeBlankSlot(group,regListId,preserveOnly); consumed++; }
          html+='<tr class="slot'+(item.units>1?' merged-slot':'')+(signed?' signed-row':'')+'" data-row-key="row-'+row.__rowId+'" data-slot-start="'+item.slotStart+'" style="height:calc(var(--slot-h) * '+item.units+')"><td class="c-date">'+dateCellHtml(row)+'</td><td class="c-reg">'+(signed?staticTextCell(row['A/C Reg'],'locked-cell'):editableTextInput('A/C Reg',row.__rowId,row['A/C Reg'],'G-XXXX','',regListId))+'</td><td class="c-job">'+(signed?staticTextCell(row['Job No'],'locked-cell'):editableTextInput('Job No',row.__rowId,row['Job No'],'Job No'))+'</td><td class="c-task">'+taskCellHtml(row)+'</td><td class="c-sup">'+editableSupervisorCell(row)+'</td></tr>';
          consumed=item.pageSlotStart+item.units;
        }
        for(var j=consumed;j<PAGE_SLOTS;j++) html+=makeBlankSlot(group,regListId,preserveOnly);
        return html;
      }
      function renderDeclaration(){ return '<tfoot><tr class="declaration-row"><td colspan="5"><div class="declaration"><div class="declaration-star">*</div><div class="declaration-text">The above work has been carried out correctly by the logbook owner under my supervision and in accordance with the<br>appropriate technical documentation.</div></div></td></tr></tfoot>'; }
      function renderPageFooter(owner, sign, footerId){
        var ownerText=esc(owner||'')||'&nbsp;';
        var signText=esc(sign||'')||'&nbsp;';
        return '<div class="owner-row"><div class="dots-field"><span class="dots-label">Logbook Owner\'s Name:</span><span class="dots-line"><span>'+ownerText+'</span></span></div><div class="dots-field"><span class="dots-label">Signature:</span><span class="dots-line"><span>'+signText+'</span></span></div></div><div style="margin-top:18px" class="bottomline"></div><div class="footer-id">'+esc(footerId||'')+'</div>';
      }
      function renderPage(type, chapter, rowsHtml, owner, sign){ return '<section class="page"><div class="headrow"><div>CAP 741</div><div>Aircraft Maintenance Engineer\'s Logbook</div></div><div class="topline"></div><div class="title">Section 3.1&nbsp;&nbsp; Maintenance Experience</div><div class="dots-row"><div class="field-stack"><div class="dots-field"><span class="dots-label">Aircraft Type:</span><span class="dots-line">'+type+'</span></div><div class="subnote">(Aircraft/Engine combination)</div></div><div class="field-stack top-pad"><div class="dots-field"><span class="dots-label">ATA Chapter:</span><span class="dots-line">'+chapter+'</span></div></div></div><div class="frame"><table class="sheet"><thead><tr><th class="c-date">Date</th><th class="c-reg">A/C Reg</th><th class="c-job">Job No</th><th class="c-task">Task Detail</th><th class="c-sup">Supervisor&rsquo;s Name Signature,<br>and Licence Number</th></tr></thead><tbody>'+rowsHtml+'</tbody>'+renderDeclaration()+'</table></div>'+renderPageFooter(owner,sign,'Section 3.1')+'</section>'; }
      function renderEditablePage(group, rowsHtml, owner, sign, pageKey){
        var aircraftLabelText=(group&&group.typeList&&group.typeList.length>1)?'Aircraft Types:':'Aircraft Type:';
        return '<section class="page" data-group-key="'+esc(s(group&&group.key))+'" data-page-group-label="'+esc(group&&group.mode===PAGE_GROUPING_GROUP?s(group.groupLabel):'')+'" data-page-key="'+esc(pageKey||'')+'"><div class="headrow"><div>CAP 741</div><div>Aircraft Maintenance Engineer\'s Logbook</div></div><div class="topline"></div><div class="title">Section 3.1&nbsp;&nbsp; Maintenance Experience</div><div class="dots-row"><div class="field-stack"><div class="dots-field"><span class="dots-label">'+aircraftLabelText+'</span>'+renderPageHeaderDots(group.type,'Aircraft Type','aircraft-type-list',!!group.typeEditable)+'</div><div class="subnote">(Aircraft/Engine combination)</div></div><div class="field-stack top-pad"><div class="dots-field"><span class="dots-label">ATA Chapter:</span>'+renderPageHeaderDots(group.chapter+(group.chapterDesc?' - '+group.chapterDesc:''),'Chapter','chapter-list',true)+'</div></div></div><div class="frame"><table class="sheet"><thead><tr><th class="c-date">Date</th><th class="c-reg">A/C Reg</th><th class="c-job">Job No</th><th class="c-task">Task Detail</th><th class="c-sup">Supervisor&rsquo;s Name, Signature<br>and Licence Number</th></tr></thead><tbody>'+rowsHtml+'</tbody>'+renderDeclaration()+'</table></div>'+renderPageFooter(owner,sign,'Section 3.1')+'</section>';
      }
      function renderDataPage(group, items, pageKey, preserveOnly){ return renderEditablePage(group,makeRows(items,group,preserveOnly),esc(LOG_OWNER_INFO.name),esc(LOG_OWNER_INFO.signature),pageKey); }
      function supervisorListSourceRows(){
        if(settingsModal&&settingsModal.className.indexOf('open')!==-1&&settingsActiveTab==='supervisors') return collectSupervisorPrintRows();
        return [];
      }
      function blankSupervisorListRowHtml(){
        return '<tr class="slot supervisor-list-slot"><td class="c-name"><div class="supervisor-list-name">&nbsp;</div></td><td class="c-stamp"><span class="blank-line">&nbsp;</span></td><td class="c-licence"><div class="supervisor-list-licence">&nbsp;</div></td><td class="c-signature"><span class="blank-line">&nbsp;</span></td></tr>';
      }
      function supervisorListRowHtml(item){
        var name=s(item&&item['Signatory Name']);
        var licence=s(item&&(item['License Number']||item['Licence Number']));
        return '<tr class="slot supervisor-list-slot"><td class="c-name"><div class="supervisor-list-name">'+(esc(name)||'&nbsp;')+'</div></td><td class="c-stamp"><span class="blank-line">&nbsp;</span></td><td class="c-licence"><div class="supervisor-list-licence">'+(esc(licence)||'&nbsp;')+'</div></td><td class="c-signature"><span class="blank-line">&nbsp;</span></td></tr>';
      }
      function renderSupervisorListPage(items, pageNumber, totalPages){
        var rowsHtml='',list=items||[];
        for(var i=0;i<list.length;i++) rowsHtml+=supervisorListRowHtml(list[i]);
        for(var j=list.length;j<SUPERVISOR_LIST_ROWS_PER_PAGE;j++) rowsHtml+=blankSupervisorListRowHtml();
        return '<section class="page supervisor-list-page" data-page-key="supervisor-list-'+pageNumber+'">'+
          '<div class="headrow"><div>CAP 741</div><div>Aircraft Maintenance Engineer\'s Logbook</div></div>'+
          '<div class="topline"></div>'+
          '<div class="title">Section Supervisor\'s list</div>'+
          '<div class="frame"><table class="sheet supervisor-list-sheet"><thead><tr><th class="c-name">Supervisor&rsquo;s Name</th><th class="c-stamp">Stamp</th><th class="c-licence">Licence Number</th><th class="c-signature">Signature</th></tr></thead><tbody>'+rowsHtml+'</tbody></table></div>'+
          renderPageFooter(NEW_WORKBOOK_SUPERVISOR_NAME,'','Section Supervisor\'s list')+
          '</section>';
      }
      function renderSupervisorListPages(){
        var list=supervisorListSourceRows(),html=[],pageItems,pageNumber=0,totalPages=Math.max(1,Math.ceil((list.length||1)/SUPERVISOR_LIST_ROWS_PER_PAGE));
        if(!list.length) return renderSupervisorListPage([],1,1);
        for(var i=0;i<list.length;i+=SUPERVISOR_LIST_ROWS_PER_PAGE){
          pageNumber++;
          pageItems=list.slice(i,i+SUPERVISOR_LIST_ROWS_PER_PAGE);
          html.push(renderSupervisorListPage(pageItems,pageNumber,totalPages));
        }
        return html.join('');
      }
      // Main UI render pass. This rebuilds the visible pages from the current in-memory
      // state instead of diffing small DOM fragments, which keeps layout logic simpler.
      function renderAll(){ try { taskTextMeasureCache=null; syncMindMapButtonVisibility(); syncFilterAircraftModeUi(); syncFilterButtonState(); syncSearchUi(); renderFilterStrip(); var activeRows=workbookContentRows(rows),visibleRows=activeRows.filter(function(row){ return rowMatchesSearch(row)&&(!hasActiveFilters()||rowMatchesFilters(row)); }),preserveFilteredSlots=shouldPreserveFilteredRowSlots(); if(!visibleRows.length){ syncSharedDatalists(sharedDatalistsHtml()); pagesEl.innerHTML=renderEmptyState(); return; } var renderedGroups=buildRenderedGroups(activeRows,visibleRows),html=[]; syncSharedDatalists(sharedDatalistsHtml()+renderedGroupDatalistsHtml(renderedGroups)); for(var i=0;i<renderedGroups.length;i++){ var group=renderedGroups[i].group,pages=renderedGroups[i].pages; for(var j=0;j<pages.length;j++){ var pageKey=(s(group.key)+'||'+pages[j].map(function(item){ return item.row.__rowId; }).join('-')); html.push(renderDataPage(group,pages[j],pageKey,preserveFilteredSlots)); } } pagesEl.innerHTML=html.join(''); } catch(e){ fail('Could not render pages: '+e.message); } }
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
      function manualPageToLogRows(page){
        var heads=page.heads||{},chapterInfo=completeChapterParts(parseChapterValue(heads.chapter).chapter,parseChapterValue(heads.chapter).chapterDesc),out=[];
        for(var i=0;i<(page.rows||[]).length;i++){
          var item=page.rows[i],supervisorRecord=supervisorRecordFor(item.supName)||{},row={'Aircraft Type':s(heads.type)||AIRCRAFT_MAP[s(item.reg).toUpperCase()]||'','A/C Reg':s(item.reg).toUpperCase(),'Chapter':chapterInfo.chapter,'Chapter Description':referenceOnlySaveEnabled()?'':chapterInfo.chapterDesc,'Date':formatDateDisplay(item.date),'Job No':s(item.job),'FAULT':'','Task Detail':s(item.task),'Rewriten for cap741':s(item.task),'Flags':'',[SUPERVISOR_ID_FIELD]:'','Approval Name':s(item.supName),'Approval stamp':s(item.supStamp)||s(supervisorRecord.stamp)||'','Aprroval Licence No.':s(item.supLicence)||(!referenceOnlySaveEnabled()?s(supervisorRecord.licence):'')};
          applyRowReferenceData(row);
          out.push(row);
        }
        return out;
      }

      // ---- Print ----
      function pageElements(){ return Array.prototype.slice.call(pagesEl.querySelectorAll('.page')); }
      function currentVisiblePage(){ var pages=pageElements(); if(!pages.length) return null; var viewportMid=window.innerHeight/2,best=pages[0],bestDistance=Infinity; for(var i=0;i<pages.length;i++){ var rect=pages[i].getBoundingClientRect(),center=rect.top+(rect.height/2),distance=Math.abs(center-viewportMid); if(rect.top<=viewportMid&&rect.bottom>=viewportMid) return pages[i]; if(distance<bestDistance){ bestDistance=distance; best=pages[i]; } } return best; }
      function clearPrintSelection(){ var pages=pageElements(); document.body.classList.remove('print-current','print-current-overlay','print-current-other-layout','print-supervisor-list'); for(var i=0;i<pages.length;i++) pages[i].classList.remove('print-exclude'); if(supervisorPrintHost) supervisorPrintHost.innerHTML=''; printMode=''; }
      function printCurrentPage(){ var current=currentVisiblePage(),pages=pageElements(); clearPrintSelection(); if(!current||!pages.length){ window.print(); return; } document.body.classList.add('print-current'); for(var i=0;i<pages.length;i++){ if(pages[i]!==current) pages[i].classList.add('print-exclude'); } printMode='current'; window.print(); }
      function printCurrentPageOverlay(){ var current=currentVisiblePage(),pages=pageElements(); clearPrintSelection(); if(!current||!pages.length){ window.print(); return; } document.body.classList.add('print-current-overlay'); for(var i=0;i<pages.length;i++){ if(pages[i]!==current) pages[i].classList.add('print-exclude'); } printMode='current-overlay'; window.print(); }
      function printCurrentOtherLayout(){ var current=currentVisiblePage(),pages=pageElements(); clearPrintSelection(); if(!current||!pages.length){ window.print(); return; } writeOtherLayoutPrintCssVars(current,otherLayoutMeasurements); document.body.classList.add('print-current-other-layout'); for(var i=0;i<pages.length;i++){ if(pages[i]!==current) pages[i].classList.add('print-exclude'); } printMode='current-other-layout'; window.print(); }
      function printSupervisorList(){ clearPrintSelection(); if(!supervisorPrintHost){ window.print(); return; } supervisorPrintHost.innerHTML=renderSupervisorListPages(); document.body.classList.add('print-supervisor-list'); printMode='supervisor-list'; window.print(); }
      function printAllPages(){ clearPrintSelection(); printMode='all'; window.print(); }
      function confirmOtherLayoutPrint(){ otherLayoutMeasurements=validatedOtherLayoutMeasurements(otherLayoutMeasurements); closeOtherLayoutModal(); printCurrentOtherLayout(); }

      // ---- Editor active state ----
      function editorIsActive(){ var active=document.activeElement; if(!active) return false; return !!(pagesEl.contains(active)&&(active.matches('input, textarea, [contenteditable="true"]')||active.classList.contains('editable-cell'))); }
      function captureActiveEditorState(){ var cell=activeEditorCell(); if(cell) updateRowFromEditor(cell); }
      function activeEditorCell(){ var active=document.activeElement; if(!active||!active.closest) return null; return active.closest('.editable-cell')||active.closest('[data-row-id]')||active.closest('[data-new-row]')||null; }

      // ---- Row editing ----
      function createRowFromBlankCell(cell){ var tr=cell.closest('tr'),first=tr.querySelector('[data-new-row="1"]'); if(!first) return null; var row=emptyLogRow(first.getAttribute('data-new-type')||'',first.getAttribute('data-new-chapter')||'',first.getAttribute('data-new-chapter-desc')||''); row.__pageGroupLabel=s(first.getAttribute('data-new-page-group')); row.__pageFilterAircraftReg=s(first.getAttribute('data-new-filter-reg')).toUpperCase(); syncRowPageGroupLabel(row,first); rows.push(row); rowsById[String(row.__rowId)]=row; var nodes=tr.querySelectorAll('[data-new-row="1"]'); for(var i=0;i<nodes.length;i++){ nodes[i].setAttribute('data-row-id',row.__rowId); nodes[i].removeAttribute('data-new-row'); nodes[i].removeAttribute('data-new-type'); nodes[i].removeAttribute('data-new-chapter'); nodes[i].removeAttribute('data-new-chapter-desc'); nodes[i].removeAttribute('data-new-page-group'); nodes[i].removeAttribute('data-new-filter-reg'); } return row; }
      // Sync a single edited control back into the canonical row object.
      function updateRowFromEditor(cell){ if(!cell) return null; var row=rowById(cell.getAttribute('data-row-id')); if(!row&&cell.hasAttribute('data-new-row')) row=createRowFromBlankCell(cell); if(!row) return null; var field=cell.getAttribute('data-edit-field'),value=(cell.tagName==='INPUT')?valueOf(cell):textOf(cell); if(field==='Date'){ var entry=cell.closest('.date-entry'),picker=entry&&entry.querySelector('[data-date-picker]'),rawDate=(picker&&picker.value)||value,iso=toIsoInputDate(rawDate),originalDisplay=formatDateDisplay(s(row.__rawDate)); value=iso?toDisplayDate(iso):s(rawDate); syncDateControl(entry,iso); row.__dateDirty=!!value ? value!==originalDisplay : !!s(row.__rawDate); } row[field]=value; if(field==='Task Detail') row['Rewriten for cap741']=value; if(field==='A/C Reg'){ row['A/C Reg']=value.toUpperCase(); if(cell.value!==row['A/C Reg']) cell.value=row['A/C Reg']; var mapped=AIRCRAFT_MAP[row['A/C Reg']]; if(mapped) row['Aircraft Type']=mapped; } if(field==='Approval Name'){ var licenceInput=cell.closest('td')&&cell.closest('td').querySelector('[data-edit-field="Aprroval Licence No."]'); var resolvedSup=fillSupervisorFields(cell,licenceInput,row); if(resolvedSup){ row['Approval Name']=resolvedSup.name; row['Approval stamp']=resolvedSup.stamp; row['Aprroval Licence No.']=licenceInput?s(licenceInput.value):(resolvedSup.licence||''); row.__manualApprovalLicenceNo=false; } } if(field==='Aprroval Licence No.'){ var nameInput=cell.closest('td')&&cell.closest('td').querySelector('[data-edit-field="Approval Name"]'); setRowSupervisorFields(row,nameInput?nameInput.value:row['Approval Name'],value); } syncRowPageGroupLabel(row,cell); if(!value&&cell.classList&&cell.classList.contains('editable-cell')) cell.innerHTML='&nbsp;'; updateRowDirtyState(row); refreshUnsavedChangesState(); return row; }
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
      function rowSlotStartFromButton(button){ var tr=button&&button.closest?button.closest('tr'):null,slot=Number(tr&&tr.getAttribute('data-slot-start')); return isFinite(slot)&&slot>=0?slot:-1; }
      function toggleSignedRow(button){
        var rowId=button&&button.getAttribute?button.getAttribute('data-row-id'):'',row=rowId?rowById(rowId):null;
        if(!row) return;
        if(isRowSigned(row)){
          setRowSignedState(row,false,-1);
        } else {
          setRowSignedState(row,true,rowSlotStartFromButton(button));
        }
        updateRowDirtyState(row);
        refreshUnsavedChangesState();
        renderAllWithMotion();
        scheduleAutoSave();
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
      function syncBlankRowMetadata(page, type, chapter, chapterDesc, groupKey, groupLabel, filterReg){ var blanks=page.querySelectorAll('[data-new-row="1"]'); for(var i=0;i<blanks.length;i++){ blanks[i].setAttribute('data-new-type',type); blanks[i].setAttribute('data-new-chapter',chapter); blanks[i].setAttribute('data-new-chapter-desc',chapterDesc); blanks[i].setAttribute('data-new-page-group',s(groupLabel)); blanks[i].setAttribute('data-new-filter-reg',s(filterReg||'')); } page.setAttribute('data-group-key',s(groupKey)); page.setAttribute('data-page-group-label',s(groupLabel)); }
      function saveModalPageRows(rowsToAppend){ if(!rowsToAppend.length) return; appendRows(rowsToAppend); refreshUnsavedChangesState(); }

      // ---- Task detail modal ----
      function openInfoModal(){ if(infoModal) infoModal.className='modal-backdrop open'; }
      function closeInfoModal(){ if(infoModal) infoModal.className='modal-backdrop'; }
      function taskDetailStateFromRow(row){ return { chapter:s(row['Chapter']), chapterDesc:s(row['Chapter Description']), fault:s(row['FAULT']), task:s(row['Task Detail']), rewrite:s(row['Rewriten for cap741']), flags:s(row['Flags']) }; }
      function restoreTaskDetailState(row, state){ if(!row||!state) return; row['Chapter']=state.chapter; row['Chapter Description']=state.chapterDesc; row['FAULT']=state.fault; row['Task Detail']=state.task; row['Rewriten for cap741']=state.rewrite; row['Flags']=serializeFlagSelection(state.flags); }
      function previewTaskDetailForm(){ var row=rowById(lastTaskDetailRowId); if(!row) return; applyTaskDetailForm(row,readTaskDetailForm()); updateRowDirtyState(row); refreshUnsavedChangesState(); }
      function openTaskDetail(rowId){ lastTaskDetailFocus=document.activeElement&&pagesEl.contains(document.activeElement)?document.activeElement:null; captureActiveEditorState(); var row=rowById(rowId); if(!row) return; lastTaskDetailRowId=rowId; taskDetailOriginalState=taskDetailStateFromRow(row); taskDetailRewriteDirty=false; detailChapterEl.value=chapterLabelText(row); detailFaultEl.value=s(row['FAULT']); detailTaskEl.value=s(row['Task Detail']); detailRewriteEl.value=s(row['Rewriten for cap741']||row['Task Detail']); renderTaskDetailFlagOptions(rowFlagLabels(row)); taskDetailModal.className='modal-backdrop open'; if(typeof requestAnimationFrame==='function') requestAnimationFrame(autoSizeDetailTextareas); else autoSizeDetailTextareas(); }
      function showConfirmDialog(title, text, okLabel){ return new Promise(function(resolve){ confirmResolver=resolve; if(confirmTitleEl) confirmTitleEl.textContent=title||'Confirm'; if(confirmTextEl) confirmTextEl.textContent=text||'Are you sure?'; if(confirmOkBtn) confirmOkBtn.textContent=okLabel||'Confirm'; if(confirmModal) confirmModal.className='modal-backdrop open'; }); }
      function closeConfirmDialog(result){ if(confirmModal) confirmModal.className='modal-backdrop'; if(confirmResolver){ var resolve=confirmResolver; confirmResolver=null; resolve(!!result); } }
      function readTaskDetailForm(){
        return {
          chapter:s(detailChapterEl.value),
          fault:s(detailFaultEl.value),
          task:s(detailTaskEl.value),
          rewrite:s(detailRewriteEl.value),
          flags:readTaskDetailSelectedFlags()
        };
      }
      function applyTaskDetailForm(row, form){
        var parsedChapter=parseChapterValue(form.chapter);
        row['FAULT']=form.fault;
        row['Task Detail']=form.task;
        row['Rewriten for cap741']=form.rewrite||form.task;
        setRowFlags(row,form.flags);
        if(parsedChapter.chapter){
          row['Chapter']=parsedChapter.chapter;
          row['Chapter Description']=parsedChapter.chapterDesc;
          applyRowChapterReference(row);
        }
      }
      function saveTaskDetail(){
        var row=rowById(lastTaskDetailRowId);
        if(!row) return;
        applyTaskDetailForm(row,readTaskDetailForm());
        updateRowDirtyState(row);
        refreshUnsavedChangesState();
        if(!rowHasEntryContent(row)&&comparableRowSignature(row)!==(savedComparableRowSignature(row)||'')){
          removeRowById(row.__rowId);
        }
        renderAll();
        closeTaskDetail(true);
        scheduleAutoSave();
      }
      function closeTaskDetail(keepPreviewChanges){ var row=lastTaskDetailRowId&&rowById(lastTaskDetailRowId); if(!keepPreviewChanges&&row&&taskDetailOriginalState){ restoreTaskDetailState(row,taskDetailOriginalState); updateRowDirtyState(row); refreshUnsavedChangesState(); renderAll(); } taskDetailModal.className='modal-backdrop'; lastTaskDetailRowId=null; taskDetailOriginalState=null; taskDetailRewriteDirty=false; if(lastTaskDetailFocus&&document.contains(lastTaskDetailFocus)&&typeof lastTaskDetailFocus.focus==='function'){ try { lastTaskDetailFocus.focus({preventScroll:true}); } catch(e){ try { lastTaskDetailFocus.focus(); } catch(err){} } } lastTaskDetailFocus=null; }

      // ---- IndexedDB ----
      function withHandleDb(mode){ return new Promise(function(resolve,reject){ if(!window.indexedDB){ reject(new Error('IndexedDB unavailable')); return; } var request=indexedDB.open(DB_NAME,1); request.onupgradeneeded=function(){ var db=request.result; if(!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE); }; request.onerror=function(){ reject(request.error||new Error('Could not open file-handle store')); }; request.onsuccess=function(){ var db=request.result,tx=db.transaction(DB_STORE,mode),store=tx.objectStore(DB_STORE); resolve({db:db,tx:tx,store:store}); }; }); }
      async function loadStoredHandle(key){ try { var ctx=await withHandleDb('readonly'); return await new Promise(function(resolve,reject){ var req=ctx.store.get(key); req.onsuccess=function(){ ctx.db.close(); resolve(req.result||null); }; req.onerror=function(){ ctx.db.close(); reject(req.error||new Error('Could not read stored file handle')); }; }); } catch(e){ return null; } }
      async function storeHandle(key, handle){ try { var ctx=await withHandleDb('readwrite'); return await new Promise(function(resolve,reject){ var req=ctx.store.put(handle,key); req.onsuccess=function(){ ctx.db.close(); resolve(true); }; req.onerror=function(){ ctx.db.close(); reject(req.error||new Error('Could not store file handle')); }; }); } catch(e){ return false; } }
      async function removeStoredHandle(key){ try { var ctx=await withHandleDb('readwrite'); return await new Promise(function(resolve,reject){ var req=ctx.store.delete(key); req.onsuccess=function(){ ctx.db.close(); resolve(true); }; req.onerror=function(){ ctx.db.close(); reject(req.error||new Error('Could not remove stored file handle')); }; }); } catch(e){ return false; } }
      async function ensurePermission(handle, mode){
        mode=mode==='readwrite'?'readwrite':'read';
        if(!handle) return false;
        if(typeof handle.queryPermission!=='function'){
          return mode==='readwrite' ? typeof handle.createWritable==='function' : typeof handle.getFile==='function';
        }
        var opts={mode:mode};
        if(await handle.queryPermission(opts)==='granted') return true;
        if(typeof handle.requestPermission!=='function'){
          return mode==='readwrite' ? typeof handle.createWritable==='function' : typeof handle.getFile==='function';
        }
        return await handle.requestPermission(opts)==='granted';
      }
      async function pickWorkbookHandle(){
        if(!filePickerSupported()) throw new Error('File picker not supported. Open this page via a local web server (e.g. VS Code Live Server) and use Chrome or Edge.');
        var picked=await window.showOpenFilePicker({multiple:false,types:[{description:'CAP741 Excel workbook',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}],excludeAcceptAllOption:true});
        var handle=picked&&picked[0];
        if(!handle) return null;
        if(!handleIsWorkbook(handle)) throw new Error('Please choose a .xlsx Excel file.');
        if(!await ensurePermission(handle,'readwrite')) return null;
        await storeHandle(LINKED_FILE_KEY,handle);
        setLinkedWorkbookName(handle);
        return handle;
      }
      async function pickWorkbookHandleTransient(){
        if(!filePickerSupported()) throw new Error('File picker not supported. Open this page via a local web server (e.g. VS Code Live Server) and use Chrome or Edge.');
        var picked=await window.showOpenFilePicker({multiple:false,types:[{description:'Excel workbook',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}],excludeAcceptAllOption:true});
        var handle=picked&&picked[0];
        if(!handle) return null;
        if(!handleIsWorkbook(handle)) throw new Error('Please choose a .xlsx Excel file.');
        if(!await ensurePermission(handle,'read')) return null;
        return handle;
      }
      function pickWorkbookFileInput(){
        return new Promise(function(resolve,reject){
          var input=workbookOpenInput,settled=false;
          if(!input){ reject(new Error('Workbook file input is not available.')); return; }
          function cleanup(){
            input.removeEventListener('change',onChange);
            input.removeEventListener('cancel',onCancel);
          }
          function finish(file){
            if(settled) return;
            settled=true;
            cleanup();
            resolve(file||null);
          }
          function onChange(){ finish(input.files&&input.files[0]?input.files[0]:null); }
          function onCancel(){ finish(null); }
          input.accept=workbookAcceptValue();
          input.value='';
          input.addEventListener('change',onChange);
          input.addEventListener('cancel',onCancel);
          try {
            input.click();
          } catch(e){
            cleanup();
            reject(e);
          }
        });
      }
      async function pickWorkbookSource(persistHandle){
        if(filePickerSupported()&&(!persistHandle||persistentExcelLinkingSupported())){
          var handle=persistHandle?await pickWorkbookHandle():await pickWorkbookHandleTransient();
          if(!handle) return null;
          return { handle:handle, file:await handle.getFile(), name:s(handle.name) };
        }
        var file=await pickWorkbookFileInput();
        if(!file) return null;
        if(!/\.xlsx$/i.test(s(file.name))) throw new Error('Please choose a .xlsx Excel file.');
        return { handle:null, file:file, name:s(file.name) };
      }
      function readFileArrayBuffer(file){
        if(!file) return Promise.reject(new Error('No file selected.'));
        if(typeof file.arrayBuffer==='function') return file.arrayBuffer();
        return new Promise(function(resolve,reject){
          var reader=new FileReader();
          reader.onload=function(){ resolve(reader.result); };
          reader.onerror=function(){ reject(reader.error||new Error('Could not read the selected file.')); };
          reader.readAsArrayBuffer(file);
        });
      }
      async function pickNewWorkbookHandle(){
        if(!fileSavePickerSupported()) throw new Error('Save file picker not supported. Use Chrome or Edge over a local web server to create a new Excel file.');
        var handle=await window.showSaveFilePicker({suggestedName:'cap741-data.xlsx',types:[{description:'CAP741 Excel workbook',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}],excludeAcceptAllOption:true});
        if(!handle) return null;
        if(!handleIsWorkbook(handle)) throw new Error('Please save the new workbook as a .xlsx Excel file.');
        if(!await ensurePermission(handle)) return null;
        try {
          var existingFile=await handle.getFile();
          if(existingFile&&existingFile.size>0) throw new Error('That Excel file already exists. Choose a new filename so the app does not overwrite it.');
        } catch(e){
          if(e&&e.name&&e.name!=='NotFoundError') throw e;
        }
        await storeHandle(LINKED_FILE_KEY,handle);
        setLinkedWorkbookName(handle);
        return handle;
      }
      function googleSheetIdFromInput(value){
        var raw=s(value);
        if(!raw) return '';
        var match=/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/.exec(raw);
        return match ? match[1] : raw;
      }
      function waitForGoogleIdentity(timeoutMs){
        timeoutMs=timeoutMs||10000;
        return new Promise(function(resolve,reject){
          var started=Date.now();
          (function check(){
            if(window.google&&window.google.accounts&&window.google.accounts.oauth2){ resolve(true); return; }
            if(Date.now()-started>=timeoutMs){ reject(new Error('Google Sign-In library did not load. Refresh the page and try again.')); return; }
            setTimeout(check,50);
          })();
        });
      }
      async function requestGoogleAccessToken(interactive){
        if(googleAccessToken) return googleAccessToken;
        var clientId=ensureGoogleClientId(!!interactive);
        if(!clientId) throw new Error('A Google OAuth Client ID is required before using Google Sheets.');
        await waitForGoogleIdentity();
        return await new Promise(function(resolve,reject){
          googleTokenClient=window.google.accounts.oauth2.initTokenClient({
            client_id: clientId,
            scope: GOOGLE_SHEETS_SCOPE,
            callback: function(response){
              if(response&&response.error){
                googleAccessToken='';
                reject(new Error(response.error_description||response.error));
                return;
              }
              googleAccessToken=s(response&&response.access_token);
              if(!googleAccessToken){
                reject(new Error('Google did not return an access token.'));
                return;
              }
              resolve(googleAccessToken);
            }
          });
          googleTokenClient.requestAccessToken({prompt:interactive?'consent':''});
        });
      }
      function googleApiErrorMessage(status, bodyText){
        var detail='';
        try {
          var parsed=JSON.parse(bodyText||'{}');
          detail=s(parsed&&parsed.error&&parsed.error.message);
        } catch(e){}
        return detail||('Google Sheets request failed ('+status+').');
      }
      async function googleApiRequest(method, url, body, interactive){
        var token=await requestGoogleAccessToken(!!interactive);
        var response=await fetch(url,{method:method,headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},body:body==null?undefined:JSON.stringify(body)});
        if(response.status===401&&interactive!==true){
          googleAccessToken='';
          token=await requestGoogleAccessToken(true);
          response=await fetch(url,{method:method,headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},body:body==null?undefined:JSON.stringify(body)});
        }
        if(!response.ok) throw new Error(googleApiErrorMessage(response.status,await response.text()));
        return response;
      }
      async function googleApiJson(method, url, body, interactive){
        var response=await googleApiRequest(method,url,body,interactive);
        if(response.status===204) return null;
        return await response.json();
      }
      async function googleSpreadsheetMetadata(spreadsheetId, interactive){
        return await googleApiJson('GET','https://sheets.googleapis.com/v4/spreadsheets/'+encodeURIComponent(spreadsheetId)+'?fields=spreadsheetId,spreadsheetUrl,properties.title,sheets.properties.title',null,interactive);
      }
      async function googleEnsureSheetTabs(spreadsheetId, interactive){
        var metadata=await googleSpreadsheetMetadata(spreadsheetId,interactive),titles=Object.create(null),requests=[];
        for(var i=0;i<((metadata&&metadata.sheets)||[]).length;i++) titles[s((((metadata.sheets[i]||{}).properties||{}).title))]=true;
        for(var j=0;j<GOOGLE_SHEET_TITLES.length;j++) if(!titles[GOOGLE_SHEET_TITLES[j]]) requests.push({addSheet:{properties:{title:GOOGLE_SHEET_TITLES[j]}}});
        if(requests.length) await googleApiJson('POST','https://sheets.googleapis.com/v4/spreadsheets/'+encodeURIComponent(spreadsheetId)+':batchUpdate',{requests:requests},interactive);
        return metadata;
      }
      async function fetchGoogleSheetObjects(spreadsheetId, interactive){
        var metadata=await googleEnsureSheetTabs(spreadsheetId,interactive);
        var params=new URLSearchParams();
        for(var i=0;i<GOOGLE_SHEET_TITLES.length;i++) params.append('ranges',GOOGLE_SHEET_TITLES[i]+'!A:Z');
        var payload=await googleApiJson('GET','https://sheets.googleapis.com/v4/spreadsheets/'+encodeURIComponent(spreadsheetId)+'/values:batchGet?'+params.toString(),null,interactive);
        var sheets=Object.create(null),ranges=(payload&&payload.valueRanges)||[];
        for(var j=0;j<ranges.length;j++){
          var rangeTitle=s(((ranges[j]||{}).range||'').split('!')[0]);
          if(rangeTitle) sheets[rangeTitle]=objectsFromMatrix(ranges[j].values||[]);
        }
        return { metadata: metadata, sheets: sheets };
      }
      async function loadGoogleSheetState(source, interactive){
        var spreadsheetId=s(source&&source.spreadsheetId);
        if(!spreadsheetId) throw new Error('Google Sheet ID is missing.');
        var loaded=await fetchGoogleSheetObjects(spreadsheetId,interactive);
        applySheetObjects(loaded.sheets);
        setAutoLoadDefaultWorkbook(false);
        setActiveStorageSource({type:STORAGE_SOURCE_GOOGLE,spreadsheetId:spreadsheetId,title:s(loaded.metadata&&loaded.metadata.properties&&loaded.metadata.properties.title)||s(source&&source.title)},true);
        setLinkedWorkbookName(null);
      }
      async function saveGoogleSheetState(source, interactive){
        var spreadsheetId=s(source&&source.spreadsheetId);
        if(!spreadsheetId) throw new Error('No Google Sheet linked.');
        await googleEnsureSheetTabs(spreadsheetId,interactive);
        var defs=stateSheetDefinitions(),clearRanges=[],data=[];
        for(var i=0;i<defs.length;i++){
          clearRanges.push(defs[i].title+'!A:Z');
          data.push({range:defs[i].title+'!A1',majorDimension:'ROWS',values:matrixFromObjects(defs[i].headers,defs[i].rows)});
        }
        await googleApiJson('POST','https://sheets.googleapis.com/v4/spreadsheets/'+encodeURIComponent(spreadsheetId)+'/values:batchClear',{ranges:clearRanges},interactive);
        await googleApiJson('POST','https://sheets.googleapis.com/v4/spreadsheets/'+encodeURIComponent(spreadsheetId)+'/values:batchUpdate',{valueInputOption:'RAW',data:data},interactive);
        resetSavedLogbookState();
        settingsDirty=false;
      }
      async function connectExistingGoogleSheet(){
        var spreadsheetId=await requestGoogleSheetIdFromModal('Connect Google Sheet','Paste the Google Sheet URL or just the sheet ID to connect it.','Connect');
        if(!spreadsheetId) return false;
        setLoadingState(true,'Loading Google Sheet','Waiting for Google authorization...');
        await loadGoogleSheetState({spreadsheetId:spreadsheetId},true);
        setLoadButtonMode('hidden');
        await renderAllWithLoading('Loading Google Sheet','Rendering pages from Google Sheets...');
        refreshUnsavedChangesState();
        return true;
      }
      function markImportedDataAsPendingSave(){
        settingsDirty=true;
        refreshUnsavedChangesState();
      }
      async function confirmReplaceStorageLoad(label){
        if(!hasWorkbookDataLoaded()) return true;
        return await showConfirmDialog('Import Into Current Source','Importing from '+label+' will replace the current CAP741 data shown in the app, but it will keep the current linked source for saving.','Import data');
      }
      async function importDataFromExcelForLinkedSource(){
        if(sourceType(activeStorageSource)===STORAGE_SOURCE_NONE) throw new Error('Link a main storage source first, then import data into it.');
        if(!await confirmReplaceStorageLoad('an Excel file')) return false;
        if(filePickerSupported()) setLoadingState(true,'Importing data','Waiting for Excel file selection...');
        var source=await pickWorkbookSource(false);
        if(!source) return false;
        setLoadingState(true,'Importing data','Reading workbook...');
        loadWorkbookFromArrayBuffer(await source.file.arrayBuffer());
        markImportedDataAsPendingSave();
        await renderAllWithLoading('Importing data','Rendering imported CAP741 pages...');
        return true;
      }
      async function importDataFromGoogleSheetForLinkedSource(){
        if(sourceType(activeStorageSource)===STORAGE_SOURCE_NONE) throw new Error('Link a main storage source first, then import data into it.');
        if(!await confirmReplaceStorageLoad('a Google Sheet')) return false;
        var spreadsheetId=await requestGoogleSheetIdFromModal('Import Google Sheet','Paste the Google Sheet URL or just the sheet ID to import its CAP741 data into the current linked source.','Import data');
        if(!spreadsheetId) return false;
        setLoadingState(true,'Importing data','Waiting for Google authorization...');
        var loaded=await fetchGoogleSheetObjects(spreadsheetId,true);
        applySheetObjects(loaded.sheets);
        markImportedDataAsPendingSave();
        await renderAllWithLoading('Importing data','Rendering imported CAP741 pages...');
        return true;
      }
      function cleanProtectedAircraftRecord(record){
        record=record||{};
        return { group:s(record.group), reg:s(record.reg).toUpperCase(), type:s(record.type) };
      }
      function cleanProtectedSupervisorRecord(record){
        record=record||{};
        return { id:s(record.id), name:s(record.name), stamp:s(record.stamp), licence:s(record.licence), scope:s(record.scope), date:s(record.date) };
      }
      async function protectedRecordsFromStore(kind, password){
        var store=protectedDataStore(),payload='',decoded='';
        if(!store) throw new Error('Protected import data is not available.');
        payload=kind==='aircraft'?store.aircraftPayload:store.supervisorPayload;
        if(!protectedPayloadValue(payload)) throw new Error('Protected '+kind+' data is not available.');
        if(s(store.scheme)==='pbkdf2-aes-cbc-v1'){
          decoded=await decodeProtectedPayloadV2(payload,password,store);
        } else {
          if(s(password)!==s(store.password)) throw new Error('Incorrect password.');
          decoded=decodeProtectedPayload(payload, password);
        }
        return JSON.parse(decoded||'[]');
      }
      function aircraftImportSummary(result){
        result=result||{};
        var parts=['Prefilled A/C data imported'],details=[];
        if(Number(result.addedCount)||0) details.push((Number(result.addedCount)||0)+' added');
        if(Number(result.updatedCount)||0) details.push((Number(result.updatedCount)||0)+' updated');
        if(Number(result.filledCount)||0) details.push((Number(result.filledCount)||0)+' row types filled');
        return parts.join('')+(details.length?(': '+details.join(', ')):'.')+(details.length?'.':'');
      }
      function supervisorImportSummary(result){
        result=result||{};
        var details=[];
        if(Number(result.addedCount)||0) details.push((Number(result.addedCount)||0)+' added');
        if(Number(result.updatedCount)||0) details.push((Number(result.updatedCount)||0)+' updated');
        return 'Prefilled supervisors imported'+(details.length?(': '+details.join(', ')):'.')+(details.length?'.':'');
      }
      async function finalizeReferenceImport(importTitle, renderText, saveText, successText){
        markImportedDataAsPendingSave();
        await renderAllWithLoading(importTitle,renderText);
        if(usingExcelDownloadFallback()){
          success(successText+' Tap Save to export the updated Excel file.');
          return true;
        }
        setLoadingState(true,importTitle,saveText);
        await saveActiveStorage(true);
        refreshUnsavedChangesState();
        success(successText);
        return true;
      }
      function importProtectedAircraftRecords(records){
        var addedCount=0,updatedCount=0,filledCount=0,i,beforeType,beforeGroup,existing,row,clean,updated;
        for(i=0;i<(records||[]).length;i++){
          clean=cleanProtectedAircraftRecord(records[i]);
          if(!clean.reg||!clean.type) continue;
          existing=aircraftReferenceRecordForReg(clean.reg);
          beforeType=existing?s(existing.type):'';
          beforeGroup=existing?s(existing.group):'';
          upsertAircraftReferenceRecord(clean.reg,clean.type,clean.group);
          if(existing){
            updated=aircraftReferenceRecordForReg(clean.reg);
            if(beforeType!==s(updated&&updated.type)||beforeGroup!==s(updated&&updated.group)) updatedCount++;
          } else addedCount++;
        }
        for(i=0;i<rows.length;i++){
          row=rows[i];
          beforeType=s(row['Aircraft Type']);
          if(beforeType) continue;
          fillAircraftTypeFromReg(row);
          if(s(row['Aircraft Type'])&&!beforeType) filledCount++;
        }
        markSharedDatalistsDirty();
        return { addedCount:addedCount, updatedCount:updatedCount, filledCount:filledCount };
      }
      function mergeProtectedSupervisorRecords(records){
        var merged=SUPERVISOR_RECORDS.map(function(item){ return { id:s(item.id), name:s(item.name), stamp:s(item.stamp), licence:s(item.licence), scope:s(item.scope), date:s(item.date) }; });
        var byId=Object.create(null),byName=Object.create(null),addedCount=0,updatedCount=0,i,clean,target,before;
        function reindex(record){
          if(s(record.id)) byId[s(record.id)]=record;
          if(s(record.name)) byName[normalizedText(record.name)]=record;
        }
        for(i=0;i<merged.length;i++) reindex(merged[i]);
        for(i=0;i<(records||[]).length;i++){
          clean=cleanProtectedSupervisorRecord(records[i]);
          if(!clean.name) continue;
          target=(clean.id&&byId[clean.id])||byName[normalizedText(clean.name)]||null;
          if(target){
            before=JSON.stringify(target);
            target.id=clean.id||target.id;
            target.name=clean.name||target.name;
            target.stamp=clean.stamp||target.stamp;
            target.licence=clean.licence||target.licence;
            target.scope=clean.scope||target.scope;
            target.date=clean.date||target.date;
            if(before!==JSON.stringify(target)) updatedCount++;
            reindex(target);
            continue;
          }
          merged.push(clean);
          reindex(clean);
          addedCount++;
        }
        rebuildSupervisorState(merged);
        markSharedDatalistsDirty();
        return { addedCount:addedCount, updatedCount:updatedCount };
      }
      function compactTextKey(value){ return s(value).replace(/\s+/g,' ').toLowerCase(); }
      function ultraMainHeaderKey(value){ return s(value).toLowerCase().replace(/[^a-z0-9]+/g,''); }
      function ultraMainNormalizedRow(row){
        var out=Object.create(null);
        row=row||{};
        for(var key in row){
          if(!Object.prototype.hasOwnProperty.call(row,key)) continue;
          out[ultraMainHeaderKey(key)]=row[key];
        }
        return out;
      }
      function ultraMainDateIso(value){
        var raw=s(value),iso='';
        if(!raw) return '';
        iso=toIsoInputDate(raw);
        if(iso) return iso;
        iso=toIsoInputDate(raw.split(/[T\s]/)[0]);
        if(iso) return iso;
        var match=/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})/.exec(raw);
        return match?isoFromDateParts(match[1],match[2],match[3]):'';
      }
      function ultraMainDuplicateKey(dateValue, jobNo){
        var iso=ultraMainDateIso(dateValue),job=compactTextKey(jobNo);
        return (iso&&job)?(iso+'||'+job):'';
      }
      function ultraMainWorkbookRows(workbook){
        var out=[],sheetNames=(workbook&&workbook.SheetNames)||[];
        for(var i=0;i<sheetNames.length;i++){
          var records=workbookSheetObjects(workbook,sheetNames[i]);
          if(!records.length) continue;
          var probe=ultraMainNormalizedRow(records[0]);
          if(!('tailnumber' in probe) || !('recordname' in probe) || !('datetime' in probe)) continue;
          for(var j=0;j<records.length;j++){
            var source=ultraMainNormalizedRow(records[j]);
            var row=emptyLogRow('','','');
            row['A/C Reg']=s(source.tailnumber).toUpperCase();
            row['Job No']=s(source.recordname);
            row['FAULT']=s(source.entry);
            row['Task Detail']=s(source.comments);
            row['Rewriten for cap741']=row['Task Detail'];
            row['Date']=ultraMainDateIso(source.datetime)||s(source.datetime);
            row.__signedSlot=-1;
            fillAircraftTypeFromReg(row);
            row=normalizeLoadedRow(row);
            if(!rowHasEntryContent(row)) continue;
            out.push(row);
          }
        }
        return out;
      }
      function mergeUltraMainRows(importedRows){
        var existingKeys=Object.create(null),addedRows=[],skippedCount=0;
        for(var i=0;i<rows.length;i++){
          var existingKey=ultraMainDuplicateKey(workbookDateValue(rows[i]),rows[i]['Job No']);
          if(existingKey) existingKeys[existingKey]=true;
        }
        for(var j=0;j<(importedRows||[]).length;j++){
          var importedRow=importedRows[j];
          var importKey=ultraMainDuplicateKey(workbookDateValue(importedRow),importedRow['Job No']);
          if(importKey&&existingKeys[importKey]){
            skippedCount++;
            continue;
          }
          if(importKey) existingKeys[importKey]=true;
          addedRows.push(importedRow);
        }
        if(addedRows.length) appendRows(addedRows);
        syncAllRowAircraftTypes();
        initializeSignedSlots(rows);
        return { addedCount:addedRows.length, skippedCount:skippedCount };
      }
      function ultraMainImportSummary(merged){
        merged=merged||{};
        var addedCount=Math.max(0,Number(merged.addedCount)||0),skippedCount=Math.max(0,Number(merged.skippedCount)||0),parts=[];
        parts.push('UltraMain import complete: '+addedCount+' imported');
        if(skippedCount) parts.push(skippedCount+' skipped because matching Job No + Date already exists');
        return parts.join(', ')+'.';
      }
      async function importUltraMainForLinkedSource(){
        if(sourceType(activeStorageSource)===STORAGE_SOURCE_NONE) throw new Error('Link a main storage source first, then import data into it.');
        if(filePickerSupported()) setLoadingState(true,'Importing UltraMain','Waiting for Excel file selection...');
        var source=await pickWorkbookSource(false);
        if(!source) return false;
        setLoadingState(true,'Importing UltraMain','Reading UltraMain workbook...');
        var workbook=XLSX.read(await source.file.arrayBuffer(),{type:'array'});
        var importedRows=ultraMainWorkbookRows(workbook);
        if(!importedRows.length) throw new Error('No UltraMain maintenance rows were found in the selected workbook.');
        var merged=mergeUltraMainRows(importedRows);
        if(!merged.addedCount){
          if(merged.skippedCount){
            note(ultraMainImportSummary(merged));
            return false;
          }
          throw new Error('No UltraMain entries were imported.');
        }
        markImportedDataAsPendingSave();
        await renderAllWithLoading('Importing UltraMain','Rendering imported CAP741 pages...');
        success(ultraMainImportSummary(merged));
        return true;
      }
      async function importProtectedAircraftForLinkedSource(){
        if(sourceType(activeStorageSource)===STORAGE_SOURCE_NONE) throw new Error('Link a main storage source first, then import data into it.');
        if(!protectedDataAvailable()) throw new Error('Protected aircraft data is not available.');
        var password=await requestProtectedImportPassword('Aircraft Data');
        if(password==null) return false;
        var result=importProtectedAircraftRecords(await protectedRecordsFromStore('aircraft',password));
        if(!(result.addedCount||result.updatedCount||result.filledCount)) throw new Error('No aircraft data changes were imported.');
        return await finalizeReferenceImport(
          'Importing aircraft data',
          'Updating aircraft references and filling empty CAP741 aircraft types...',
          sourceType(activeStorageSource)===STORAGE_SOURCE_GOOGLE?'Saving imported aircraft data to Google Sheets...':'Saving imported aircraft data to cap741-data.xlsx...',
          aircraftImportSummary(result)+' Saved to the linked source.'
        );
      }
      async function importProtectedSupervisorsForLinkedSource(){
        if(sourceType(activeStorageSource)===STORAGE_SOURCE_NONE) throw new Error('Link a main storage source first, then import data into it.');
        if(!protectedDataAvailable()) throw new Error('Protected supervisor data is not available.');
        var password=await requestProtectedImportPassword('Supervisor Data');
        if(password==null) return false;
        var result=mergeProtectedSupervisorRecords(await protectedRecordsFromStore('supervisors',password));
        if(!(result.addedCount||result.updatedCount)) throw new Error('No supervisor data changes were imported.');
        return await finalizeReferenceImport(
          'Importing supervisors',
          'Updating the protected supervisor list...',
          sourceType(activeStorageSource)===STORAGE_SOURCE_GOOGLE?'Saving imported supervisors to Google Sheets...':'Saving imported supervisors to cap741-data.xlsx...',
          supervisorImportSummary(result)+' Saved to the linked source.'
        );
      }
      async function createNewGoogleSheet(){
        if(!hasWorkbookDataLoaded()) initializeNewWorkbookState();
        var created=await googleApiJson('POST','https://sheets.googleapis.com/v4/spreadsheets',{properties:{title:'cap741-data'},sheets:GOOGLE_SHEET_TITLES.map(function(title){ return {properties:{title:title}}; })},true);
        var source={type:STORAGE_SOURCE_GOOGLE,spreadsheetId:s(created&&created.spreadsheetId),title:s(created&&created.properties&&created.properties.title)};
        await saveGoogleSheetState(source,true);
        setAutoLoadDefaultWorkbook(false);
        setActiveStorageSource(source,true);
        setLoadButtonMode('hidden');
        await renderAllWithLoading('Creating Google Sheet','Rendering starter CAP741 pages...');
        refreshUnsavedChangesState();
        setLoadingState(false);
        await showCreatedGoogleSheetNotice(source);
        return source;
      }

      // ---- Reference data state builders (from xlsx) ----
      function parseSupervisorRecordsText(text){ var lines=String(text||'').split(/\r?\n/),out=[]; for(var i=0;i<lines.length;i++){ var line=s(lines[i]); if(!line||/^id\s+/i.test(line)) continue; var cols=line.split('\t').map(function(x){ return s(x); }); while(cols.length&&!s(cols[cols.length-1])) cols.pop(); if(cols.length<4) continue; out.push({id:cols[0]||'',name:cols[1]||'',stamp:cols[2]||'',licence:cols[3]||'',scope:cols[4]||'',date:cols[5]||''}); } return out; }
      function rebuildSupervisorState(records){ var out=[],options=[],lookup=Object.create(null); records=records||[]; for(var i=0;i<records.length;i++){ var record=records[i]||{}; var clean={id:s(record.id||record.ID||''),name:s(record.name||record['Signatory Name']||record['Name']||''),stamp:s(record.stamp||record['Stamp']||''),licence:s(record.licence||record['License Number']||record['Licence Number']||''),scope:s(record.scope||record['Scope / Limitations']||''),date:s(record.date||record['Date']||'')}; if(!clean.name) continue; out.push(clean); var label=clean.name+' | '+clean.stamp+' | '+clean.licence; options.push(label); lookup[supervisorLookupKey('name',clean.name)]=clean; lookup[supervisorLookupKey('label',label)]=clean; if(clean.id) lookup[supervisorLookupKey('id',clean.id)]=clean; } SUPERVISOR_RECORDS=out; SUPERVISOR_OPTIONS=options.sort(function(a,b){ return a.localeCompare(b); }); SUPERVISOR_LOOKUP=lookup; }
      function applyAircraftGroupRows(records){ var map=Object.create(null),out=[]; for(var i=0;i<(records||[]).length;i++){ var record=records[i]||{}; var clean={group:s(record.group||record.Group||''),reg:s(record.reg||record['A/C Reg']||'').toUpperCase(),type:s(record.type||record['Aircraft Type']||'')}; if(!clean.reg||!clean.type) continue; out.push(clean); map[clean.reg]=clean.type; } AIRCRAFT_GROUP_ROWS=out; AIRCRAFT_MAP=map; }
      function defaultChapterRows(){
        return chapterDataStore().map(function(record){
          return {
            chapter:s(record&&(
              record.chapter||
              record.Chapter
            )),
            description:s(record&&(
              record.description||
              record.Description
            ))
          };
        }).filter(function(record){ return !!record.chapter; });
      }
      function defaultChapterOptions(){
        return defaultChapterRows().map(function(record){
          return record.description ? (record.chapter+' - '+record.description) : record.chapter;
        });
      }
      function chapterOptionsMatchDefaults(){
        var defaults=defaultChapterOptions();
        if(CHAPTER_OPTIONS.length!==defaults.length) return false;
        for(var i=0;i<defaults.length;i++){
          if(s(CHAPTER_OPTIONS[i])!==s(defaults[i])) return false;
        }
        return true;
      }
      function applyChapterRows(records){
        var source=(records&&records.length)?records:defaultChapterRows();
        CHAPTER_OPTIONS=[];
        for(var i=0;i<(source||[]).length;i++){
          var chapter=s(source[i].chapter||source[i].Chapter||''),desc=s(source[i].description||source[i].Description||'');
          if(chapter) CHAPTER_OPTIONS.push(desc?chapter+' - '+desc:chapter);
        }
      }
      function aircraftWorkbookRows(){ return AIRCRAFT_GROUP_ROWS.map(function(item){ return {Group:s(item.group),'A/C Reg':s(item.reg),'Aircraft Type':s(item.type)}; }); }
      function chapterWorkbookRows(){ return CHAPTER_OPTIONS.map(function(label){ var p=parseChapterValue(label); return {Chapter:p.chapter,Description:p.chapterDesc}; }); }
      function supervisorWorkbookRows(){ return SUPERVISOR_RECORDS.map(function(item){ return {ID:s(item.id),'Signatory Name':s(item.name),Stamp:s(item.stamp),'License Number':s(item.licence),'Scope / Limitations':s(item.scope),Date:s(item.date)}; }); }
      function stateSheetDefinitions(){
        syncAllRowAircraftTypes();
        var logRows=rows.slice().sort(compareRowsNewestFirst).map(function(row){ var out={}; for(var i=0;i<LOG_HEADERS.length;i++) out[LOG_HEADERS[i]]=workbookSavedFieldValue(row,LOG_HEADERS[i]); return out; });
        return [
          { title:'Logbook', headers:LOG_HEADERS, rows:logRows },
          { title:'Aircraft', headers:['Group','A/C Reg','Aircraft Type'], rows:aircraftWorkbookRows() },
          { title:'Chapters', headers:['Chapter','Description'], rows:chapterWorkbookRows() },
          { title:'Supervisors', headers:['ID','Signatory Name','Stamp','License Number','Scope / Limitations','Date'], rows:supervisorWorkbookRows() },
          { title:'Flags', headers:['Section','Flag','Color'], rows:flagWorkbookRows() },
          { title:'Info', headers:['Key','Value'], rows:infoWorkbookRows() }
        ];
      }
      function applySheetObjects(sheetObjects){
        sheetObjects=sheetObjects||{};
        var logRows=sheetObjects.Logbook||[];
        if(logRows.length){
          var parsed=[];
          for(var i=0;i<logRows.length;i++){
            var row={};
            for(var j=0;j<LOG_HEADERS.length;j++) row[LOG_HEADERS[j]]=s(logRows[i][LOG_HEADERS[j]]);
            row[SUPERVISOR_ID_FIELD]=s(logRows[i][SUPERVISOR_ID_FIELD]||logRows[i]['Approval ID']||logRows[i]['ID']||row[SUPERVISOR_ID_FIELD]);
            row.__signedSlot=(function(value){ var slot=parseInt(s(value),10); return isFinite(slot)&&slot>=0?slot:-1; })(logRows[i]['Signed Slot']);
            parsed.push(normalizeLoadedRow(row));
          }
          rows=normalizeRows(parsed);
        } else {
          rows=normalizeRows([]);
        }
        applyAircraftGroupRows((sheetObjects.Aircraft||[]).map(function(r){ return {group:r.Group,reg:r['A/C Reg'],type:r['Aircraft Type']}; }));
        syncAllRowAircraftTypes();
        initializeSignedSlots(rows);
        applyChapterRows((sheetObjects.Chapters||[]).map(function(r){ return {chapter:r.Chapter,description:r.Description}; }));
        rebuildSupervisorState((sheetObjects.Supervisors||[]).map(function(r){ return {id:r.ID,name:r['Signatory Name'],stamp:r.Stamp,licence:r['License Number'],scope:r['Scope / Limitations'],date:r.Date}; }));
        applyFlagRows(sheetObjects.Flags||[]);
        LOG_OWNER_INFO={ name:'', signature:'', stamp:'' };
        APP_VIEW_SETTINGS=cloneAppViewSettings(DEFAULT_APP_VIEW_SETTINGS);
        var infoRows=sheetObjects.Info||[];
        for(var k=0;k<infoRows.length;k++){
          var key=normalizedText(infoRows[k].Key||infoRows[k].key),value=s(infoRows[k].Value||infoRows[k].value);
          if(key==='name') LOG_OWNER_INFO.name=value;
          if(key==='signature') LOG_OWNER_INFO.signature=value;
          if(key==='stamp') LOG_OWNER_INFO.stamp=value;
          if(key==='show mind map'||key==='show mindmap'||key==='mind map icon') APP_VIEW_SETTINGS.showMindMap=boolSettingValue(value,DEFAULT_APP_VIEW_SETTINGS.showMindMap);
          if(key==='page grouping'||key==='organize pages by'||key==='page grouping mode') APP_VIEW_SETTINGS.pageGrouping=normalizePageGrouping(value);
          if(key==='reference save'||key==='save reference values'||key==='reference-only save'||key==='reference only save') APP_VIEW_SETTINGS.referenceOnlySave=boolSettingValue(value,DEFAULT_APP_VIEW_SETTINGS.referenceOnlySave);
        }
        syncRowsToReferenceFillMode();
        markSharedDatalistsDirty();
        resetSavedLogbookState();
        settingsDirty=false;
        syncMindMapButtonVisibility();
      }
      function matrixFromObjects(headers, rows){
        var values=[headers.slice()];
        for(var i=0;i<(rows||[]).length;i++) values.push(headers.map(function(header){ return s(rows[i][header]); }));
        return values;
      }
      function objectsFromMatrix(values){
        values=values||[];
        if(!values.length) return [];
        var headers=(values[0]||[]).map(function(v){ return s(v); }),rowsOut=[];
        for(var i=1;i<values.length;i++){
          var rowValues=values[i]||[],obj={},hasValue=false;
          for(var j=0;j<headers.length;j++){
            var key=headers[j];
            if(!key) continue;
            obj[key]=s(rowValues[j]);
            if(obj[key]) hasValue=true;
          }
          if(hasValue) rowsOut.push(obj);
        }
        return rowsOut;
      }

      // ---- XLSX core ----
      function handleIsWorkbook(handle){ return !!(handle&&/\.xlsx$/i.test(s(handle.name))); }
      function workbookSheetObjects(workbook, sheetName){ var sheet=workbook&&workbook.Sheets?workbook.Sheets[sheetName]:null; if(!sheet||!window.XLSX) return []; return XLSX.utils.sheet_to_json(sheet,{defval:'',raw:false}); }
      // Workbook sheets are the source of truth on disk; this function translates them
      // into the smaller in-memory structures the UI works with.
      function loadWorkbookFromArrayBuffer(buffer){ var workbook=XLSX.read(buffer,{type:'array'}); applySheetObjects({ Logbook:workbookSheetObjects(workbook,'Logbook'), Aircraft:workbookSheetObjects(workbook,'Aircraft'), Chapters:workbookSheetObjects(workbook,'Chapters'), Supervisors:workbookSheetObjects(workbook,'Supervisors'), Flags:workbookSheetObjects(workbook,'Flags'), Info:workbookSheetObjects(workbook,'Info') }); }
      async function loadDefaultWorkbookData(){ var res=await fetch(DEFAULT_WORKBOOK_PATH,{cache:'no-store'}); if(!res.ok) throw new Error('Excel workbook returned '+res.status); loadWorkbookFromArrayBuffer(await res.arrayBuffer()); }
      function buildWorkbookFromState(){ syncAllRowAircraftTypes(); var wb=XLSX.utils.book_new(),defs=stateSheetDefinitions(); for(var i=0;i<defs.length;i++) XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(defs[i].rows,{header:defs[i].headers}),defs[i].title); return wb; }
      async function getXlsxHandle(){ try { var stored=await loadStoredHandle(LINKED_FILE_KEY); if(stored&&handleIsWorkbook(stored)){ setLinkedWorkbookName(stored); var perm=await stored.queryPermission({mode:'readwrite'}); if(perm==='granted') return stored; perm=await stored.requestPermission({mode:'readwrite'}); if(perm==='granted') return stored; } } catch(e){} return null; }
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
        resetSavedLogbookState();
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
          [SUPERVISOR_ID_FIELD]:'1',
          'Approval Name':NEW_WORKBOOK_SUPERVISOR_NAME,
          'Approval stamp':NEW_WORKBOOK_SUPERVISOR_NAME,
          'Aprroval Licence No.':NEW_WORKBOOK_SUPERVISOR_LICENCE
        });
        rows=normalizeRows([starterRow]);
        AIRCRAFT_GROUP_ROWS=starterAircraft?[starterAircraft]:[];
        AIRCRAFT_MAP=starterAircraft?(function(){ var map=Object.create(null); map[starterAircraft.reg]=starterAircraft.type; return map; })():Object.create(null);
        CHAPTER_OPTIONS=defaultChapterOptions();
        FLAG_RECORDS=defaultFlagRecords();
        LOG_OWNER_INFO={ name:NEW_WORKBOOK_OWNER_NAME, signature:'', stamp:NEW_WORKBOOK_OWNER_NAME };
        APP_VIEW_SETTINGS=cloneAppViewSettings(DEFAULT_APP_VIEW_SETTINGS);
        activeFilters=emptyFilterState();
        draftFilters=emptyFilterState();
        applySearchQuery('');
        rebuildSupervisorState([{ id:'1', name:NEW_WORKBOOK_SUPERVISOR_NAME, stamp:NEW_WORKBOOK_SUPERVISOR_NAME, licence:NEW_WORKBOOK_SUPERVISOR_LICENCE, scope:'', date:todaySupervisorDate() }]);
        markSharedDatalistsDirty();
        settingsDirty=false;
        resetSavedLogbookState();
        syncMindMapButtonVisibility();
      }
      async function writeXlsx(allowPicker){
        if(!fileSavePickerSupported()){
          var fallbackName=setSessionExcelSource(currentWorkbookFileName(),true);
          downloadWorkbookFile(fallbackName);
          return;
        }
        var handle=await getXlsxHandle();
        if(!handle&&allowPicker){
          setLoadingState(true,'Linking file','Choose the Excel workbook to save to...');
          handle=await pickWorkbookHandle();
          if(handle) setLoadingState(true,'Saving','Writing changes to cap741-data.xlsx...');
        }
        if(handle){ setAutoLoadDefaultWorkbook(false); setActiveStorageSource({type:STORAGE_SOURCE_EXCEL,name:s(handle.name)},true); }
        await writeWorkbookToHandle(handle);
      }
      async function createNewWorkbookFile(){
        if(!fileSavePickerSupported()){
          initializeNewWorkbookState();
          var fallbackName=setSessionExcelSource('cap741-data.xlsx',true);
          downloadWorkbookFile(fallbackName);
          syncLoadButtonAvailability(false);
          await renderAllWithLoading('Creating logbook','Rendering starter CAP741 pages...');
          refreshUnsavedChangesState();
          showExcelDownloadNote(fallbackName,'New Excel file ready');
          return true;
        }
        var handle=await pickNewWorkbookHandle();
        if(!handle) return false;
        initializeNewWorkbookState();
        setAutoLoadDefaultWorkbook(false);
        setActiveStorageSource({type:STORAGE_SOURCE_EXCEL,name:s(handle.name)},true);
        await writeWorkbookToHandle(handle);
        setLoadButtonMode('hidden');
        await renderAllWithLoading('Creating logbook','Rendering starter CAP741 pages...');
        refreshUnsavedChangesState();
        return true;
      }
      async function migrateCurrentDataToExcel(){
        if(!hasWorkbookDataLoaded()) throw new Error('Load a CAP741 source first before migrating data.');
        if(!fileSavePickerSupported()){
          var fallbackName=setSessionExcelSource(currentWorkbookFileName(),true);
          downloadWorkbookFile(fallbackName);
          syncLoadButtonAvailability(false);
          refreshUnsavedChangesState();
          showExcelDownloadNote(fallbackName,'Migrated Excel file ready');
          return true;
        }
        var handle=await pickNewWorkbookHandle();
        if(!handle) return false;
        setLoadingState(true,'Migrating storage','Copying the current data into the new Excel file...');
        setAutoLoadDefaultWorkbook(false);
        setActiveStorageSource({type:STORAGE_SOURCE_EXCEL,name:s(handle.name)},true);
        await writeWorkbookToHandle(handle);
        setLoadButtonMode('hidden');
        refreshUnsavedChangesState();
        return true;
      }
      async function migrateCurrentDataToGoogleSheet(){
        if(!hasWorkbookDataLoaded()) throw new Error('Load a CAP741 source first before migrating data.');
        var source=await createNewGoogleSheet();
        if(!source) return false;
        await removeStoredHandle(LINKED_FILE_KEY);
        setLinkedWorkbookName(null);
        refreshUnsavedChangesState();
        return true;
      }
      async function saveActiveStorage(allowPicker){
        if(sourceType(activeStorageSource)===STORAGE_SOURCE_GOOGLE){
          await saveGoogleSheetState(activeStorageSource,!!allowPicker);
          return;
        }
        await writeXlsx(allowPicker);
      }

      // ---- Auto-save after chapter/data changes ----
      function scheduleAutoSave(){ clearTimeout(autoSaveTimer); refreshUnsavedChangesState(); }

      // ---- Flush / save ----
      async function flushLinkedRewrite(force){ if(!hasUnsavedChanges&&!force) return; captureActiveEditorState(); if(saveInFlight){ saveQueued=true; return; } saveInFlight=true; syncSaveButtonState(true); try { clearFail(); await saveActiveStorage(!!force); refreshUnsavedChangesState(); } catch(e){ if(e&&e.name==='AbortError') fail('Save cancelled. Choose the storage source again and try saving once more.'); else fail(saveFailureMessage(e)); } finally { saveInFlight=false; syncSaveButtonState(false); if(saveQueued){ saveQueued=false; flushLinkedRewrite(true); } } }

      // ---- Settings modal with tabs ----
      function settingsTableRow(cells, kind, rowAttrs){ var html='<tr'+(rowAttrs?' '+rowAttrs:'')+'>'; for(var i=0;i<cells.length;i++) html+='<td>'+cells[i]+'</td>'; html+='<td><button type="button" class="settings-remove-btn" data-settings-remove="'+esc(kind)+'">&#x2715;</button></td></tr>'; return html; }
      function nextSupervisorNumericId(records){ var max=0; for(var i=0;i<(records||[]).length;i++){ var id=parseInt(s(records[i]&&records[i].id),10); if(isFinite(id)&&id>max) max=id; } return max+1; }
      function todaySupervisorDate(){ var now=new Date(); var iso=now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0')+'-'+String(now.getDate()).padStart(2,'0'); return formatDateDisplay(iso); }
      function supervisorPrintCheckboxHtml(checked){ return '<label class="settings-print-check"><input type="checkbox" data-supervisor-print="1"'+(checked?' checked':'')+'><span>&#10003;</span></label>'; }
      function settingsFlagColorCell(value){ return '<div class="settings-flag-color-cell"><span class="settings-flag-color-dot" style="background-color:'+esc(s(value)||'#7f93a1')+'"></span><input type="text" data-col="Color" value="'+esc(value)+'" placeholder="Gold"></div>'; }
      function renderSettingsRows(kind){ var html=''; if(kind==='aircraft'){ var list=aircraftWorkbookRows(); for(var i=0;i<list.length;i++) html+=settingsTableRow(['<input type="text" data-col="Group" value="'+esc(list[i].Group)+'">','<input type="text" data-col="A/C Reg" value="'+esc(list[i]['A/C Reg'])+'">','<input type="text" data-col="Aircraft Type" value="'+esc(list[i]['Aircraft Type'])+'">'],kind); }
      if(kind==='chapters'){ var ch=chapterWorkbookRows(); for(var j=0;j<ch.length;j++) html+=settingsTableRow(['<input type="text" data-col="Chapter" value="'+esc(ch[j].Chapter)+'">','<input type="text" data-col="Description" value="'+esc(ch[j].Description)+'">'],kind); }
      if(kind==='flags'){ var fl=flagWorkbookRows(); for(var f=0;f<fl.length;f++) html+=settingsTableRow([flagSectionSelectHtml(fl[f].Section),'<input type="text" data-col="Flag" value="'+esc(fl[f].Flag)+'">',settingsFlagColorCell(fl[f].Color)],kind); }
      if(kind==='supervisors'){ var su=supervisorWorkbookRows(); for(var k=0;k<su.length;k++) html+=settingsTableRow([supervisorPrintCheckboxHtml(false),'<input type="text" data-col="Signatory Name" value="'+esc(su[k]['Signatory Name'])+'">','<input type="text" data-col="Stamp" value="'+esc(su[k].Stamp)+'">','<input type="text" data-col="License Number" value="'+esc(su[k]['License Number'])+'">','<input type="text" data-col="Scope / Limitations" value="'+esc(su[k]['Scope / Limitations'])+'">'],kind,'data-supervisor-id="'+esc(su[k].ID)+'" data-supervisor-date="'+esc(su[k].Date)+'"'); }
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
      function collectSupervisorPrintRows(){
        var tbody=settingsBodyEl&&settingsBodyEl.querySelector('[data-settings-table="supervisors"] tbody'),out=[];
        if(!tbody) return out;
        var trs=tbody.querySelectorAll('tr');
        for(var i=0;i<trs.length;i++){
          var tr=trs[i],printToggle=tr.querySelector('[data-supervisor-print="1"]');
          if(!printToggle||!printToggle.checked) continue;
          var nameInput=tr.querySelector('[data-col="Signatory Name"]');
          var licenceInput=tr.querySelector('[data-col="License Number"]');
          var name=s(nameInput&&nameInput.value),licence=s(licenceInput&&licenceInput.value);
          if(!name&&!licence) continue;
          out.push({'Signatory Name':name,'License Number':licence});
        }
        return out;
      }
      function collectAppViewSettingsFromModal(){
        var mindMapToggle=settingsBodyEl&&settingsBodyEl.querySelector('#settingsShowMindMap');
        var referenceSaveToggle=settingsBodyEl&&settingsBodyEl.querySelector('#settingsReferenceOnlySave');
        var groupingChoice=settingsBodyEl&&settingsBodyEl.querySelector('input[name="settingsPageGrouping"]:checked');
        return {
          showMindMap: !mindMapToggle || !!mindMapToggle.checked,
          pageGrouping: normalizePageGrouping(groupingChoice&&groupingChoice.value),
          referenceOnlySave: !referenceSaveToggle || !referenceSaveToggle.checked
        };
      }
      function syncSettingsPageGroupingUi(){
        if(!settingsBodyEl||settingsActiveTab!=='view') return;
        var radios=settingsBodyEl.querySelectorAll('input[name="settingsPageGrouping"]'),selected=currentPageGrouping(),i,label,currentEl;
        for(i=0;i<radios.length;i++){
          if(radios[i].checked) selected=normalizePageGrouping(radios[i].value);
          label=radios[i].closest&&radios[i].closest('.settings-toggle-option');
          if(label) label.classList.toggle('is-active',!!radios[i].checked);
        }
        currentEl=settingsBodyEl.querySelector('[data-settings-page-grouping-current]');
        if(currentEl) currentEl.textContent='Selected: '+pageGroupingDisplayLabel(selected);
      }
      function renderSettingsBody(tab){
        settingsActiveTab=tab||'owner';
        var TABS=[{id:'owner',label:'Owner'},{id:'storage',label:'Storage'},{id:'aircraft',label:'Aircraft'},{id:'supervisors',label:'Supervisors'},{id:'chapters',label:'Chapters'},{id:'flags',label:'Flags'},{id:'view',label:'Features'}];
        var tabsHtml='<div class="settings-tabs-nav">';
        for(var t=0;t<TABS.length;t++) tabsHtml+='<button type="button" class="settings-tab-btn'+(TABS[t].id===settingsActiveTab?' active':'')+'" data-settings-tab="'+TABS[t].id+'">'+TABS[t].label+'</button>';
        tabsHtml+='</div>';
        var panelHtml='';
        if(settingsActiveTab==='owner'){
          panelHtml='<div class="settings-tab-panel"><p class="settings-panel-copy">Used on the CAP 741 page footer.</p><div class="settings-grid"><div class="settings-field"><label>Name</label><input class="settings-input" id="settingsOwnerName" type="text" value="'+esc(LOG_OWNER_INFO.name)+'"></div><div class="settings-field"><label>Stamp</label><input class="settings-input" id="settingsOwnerStamp" type="text" value="'+esc(LOG_OWNER_INFO.stamp)+'"></div></div></div>';
        } else if(settingsActiveTab==='view'){
          panelHtml='<div class="settings-tab-panel"><p class="settings-panel-copy">Control extra page tools, save behavior, and how CAP 741 pages are organized.</p><div class="settings-view-grid"><div class="settings-check-card"><div class="settings-view-card-head"><h3 class="settings-view-card-title">Mind icon</h3><p class="settings-view-card-copy">Show or hide the floating shortcut on the main screen.</p></div><label class="settings-check-main"><span class="settings-switch"><input id="settingsShowMindMap" type="checkbox" aria-label="Mind icon"'+(APP_VIEW_SETTINGS.showMindMap?' checked':'')+'><span class="settings-switch-ui" aria-hidden="true"></span></span></label></div><div class="settings-segment-card"><div class="settings-view-card-head"><h3 class="settings-view-card-title">Page Grouping</h3><p class="settings-view-card-copy">Choose how the printed CAP 741 pages are built and titled.</p></div><div class="settings-toggle-group" role="radiogroup" aria-label="Organize pages by"><label class="settings-toggle-option'+(currentPageGrouping()===PAGE_GROUPING_TYPE?' is-active':'')+'"><input type="radio" name="settingsPageGrouping" value="'+PAGE_GROUPING_TYPE+'"'+(currentPageGrouping()===PAGE_GROUPING_TYPE?' checked':'')+'>Aircraft Type</label><label class="settings-toggle-option'+(currentPageGrouping()===PAGE_GROUPING_GROUP?' is-active':'')+'"><input type="radio" name="settingsPageGrouping" value="'+PAGE_GROUPING_GROUP+'"'+(currentPageGrouping()===PAGE_GROUPING_GROUP?' checked':'')+'>Aircraft Group</label></div><p class="settings-toggle-current" data-settings-page-grouping-current="1">Selected: '+pageGroupingDisplayLabel(currentPageGrouping())+'</p><p class="settings-toggle-copy">Group mode keeps related aircraft together and shows the aircraft variants from that group in the page header.</p></div><div class="settings-check-card"><div class="settings-view-card-head"><h3 class="settings-view-card-title">Fill empty refs</h3><p class="settings-view-card-copy">When on, empty Chapter Description and Licence No. are completed from references and saved. When off, they stay view-only.</p></div><label class="settings-check-main"><span class="settings-switch"><input id="settingsReferenceOnlySave" type="checkbox" aria-label="Fill empty references"'+(!referenceOnlySaveEnabled()?' checked':'')+'><span class="settings-switch-ui" aria-hidden="true"></span></span></label></div></div>';
        } else if(settingsActiveTab==='storage'){
          panelHtml='<div class="settings-tab-panel"><p class="settings-panel-copy">See which storage source is connected, import CAP741 data from another source into it, import UltraMain maintenance actions into the current logbook, load prefilled aircraft and supervisor reference data, or migrate the current data between local Excel and Google Sheets.</p><div class="settings-linked-card"><div class="settings-linked-title">Current Linked Storage</div><p class="settings-linked-copy" id="settingsStorageSummary">Checking remembered storage...</p><div class="settings-storage-link" id="settingsStorageLinkRow" hidden><a class="settings-storage-anchor" id="settingsStorageLink" href="#" target="_blank" rel="noreferrer noopener"></a><button class="settings-copy-btn" id="settingsCopyStorageLinkBtn" data-settings-copy-link="1" type="button" aria-label="Copy link" title="Copy link"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M9 9h9v11H9z"></path><path d="M6 4h9v3H8v9H6z"></path></svg></button></div><div class="settings-linked-note" id="settingsStorageNote">Migration copies the current in-browser data to a new source and then switches the app to that source.</div><div class="settings-storage-section"><div class="settings-storage-actions"><button class="settings-secondary-btn" id="loadStorageExcelBtn" data-settings-storage-action="import-excel" type="button">Import Excel File</button><button class="settings-secondary-btn" id="loadStorageGoogleBtn" data-settings-storage-action="import-google" type="button">Import Google Sheet</button><button class="settings-secondary-btn" id="loadStorageUltraMainBtn" data-settings-storage-action="import-ultramain" type="button">Import UltraMain Report</button><button class="settings-secondary-btn" id="loadProtectedAircraftBtn" data-settings-storage-action="import-protected-aircraft" type="button">Load Prefilled A/C Data</button><button class="settings-secondary-btn" id="loadProtectedSupervisorsBtn" data-settings-storage-action="import-protected-supervisors" type="button">Load Prefilled Supervisors</button></div></div><div class="settings-storage-section"><div class="settings-linked-title">Migrate Current Data</div><div class="settings-storage-actions"><button class="settings-secondary-btn" id="migrateToExcelBtn" data-settings-storage-action="migrate-excel" type="button">Migrate To Excel</button><button class="settings-secondary-btn" id="migrateToGoogleBtn" data-settings-storage-action="migrate-google" type="button">Migrate To Google Sheet</button><button class="settings-secondary-btn" id="unlinkWorkbookBtn" data-settings-unlink="1" type="button">Unlink Source</button></div></div></div></div>';
        } else if(settingsActiveTab==='aircraft'){
          panelHtml='<div class="settings-tab-panel"><p class="settings-panel-copy">Manage aircraft registration, type, and group. Registration is used to auto-fill Aircraft Type when entering A/C Reg.</p><div class="settings-panel-toolbar"><button class="settings-add-row" type="button" data-settings-add="aircraft">+ Add aircraft</button></div><div class="settings-table-wrap"><table class="settings-table" data-settings-table="aircraft"><thead><tr><th>Group</th><th>A/C Reg</th><th>Aircraft Type</th><th></th></tr></thead><tbody>'+renderSettingsRows('aircraft')+'</tbody></table></div></div>';
        } else if(settingsActiveTab==='supervisors'){
          panelHtml='<div class="settings-tab-panel"><p class="settings-panel-copy">Manage supervisor names, stamps, and licence numbers. Tick the rows you want to include in the printed supervisor list.</p><div class="settings-panel-toolbar"><button class="settings-add-row" type="button" data-settings-add="supervisors">+ Add supervisor</button></div><div class="settings-table-wrap"><table class="settings-table" data-settings-table="supervisors"><thead><tr><th class="settings-supervisor-print-col"></th><th>Name</th><th>Stamp</th><th>Licence</th><th>Scope</th><th></th></tr></thead><tbody>'+renderSettingsRows('supervisors')+'</tbody></table></div></div>';
        } else if(settingsActiveTab==='chapters'){
          panelHtml='<div class="settings-tab-panel"><p class="settings-panel-copy">Manage ATA chapter numbers and descriptions. These appear in the Chapter dropdown on pages and filters.</p><div class="settings-panel-toolbar"><button class="settings-add-row" type="button" data-settings-add="chapters">+ Add chapter</button></div><div class="settings-table-wrap"><table class="settings-table" data-settings-table="chapters"><thead><tr><th>Chapter</th><th>Description</th><th></th></tr></thead><tbody>'+renderSettingsRows('chapters')+'</tbody></table></div></div>';
        } else if(settingsActiveTab==='flags'){
          panelHtml='<div class="settings-tab-panel"><p class="settings-panel-copy">Manage the flag list, whether a flag appears in the main list or under More Flags, and the color used when it is shown on the page.</p><div class="settings-panel-toolbar"><button class="settings-add-row" type="button" data-settings-add="flags">+ Add flag</button></div><div class="settings-table-wrap"><table class="settings-table" data-settings-table="flags"><thead><tr><th>Section</th><th>Flag</th><th>Color</th><th></th></tr></thead><tbody>'+renderSettingsRows('flags')+'</tbody></table></div></div>';
        }
        settingsBodyEl.innerHTML=tabsHtml+panelHtml;
        if(printSupervisorsBtn) printSupervisorsBtn.hidden=settingsActiveTab!=='supervisors';
        // Wire tab buttons
        var tabBtns=settingsBodyEl.querySelectorAll('.settings-tab-btn');
        for(var i=0;i<tabBtns.length;i++){ (function(btn){ btn.addEventListener('click',function(){ var nextTab=btn.getAttribute('data-settings-tab'); if(nextTab===settingsActiveTab) return; renderSettingsBody(nextTab); }); })(tabBtns[i]); }
        if(settingsActiveTab==='storage') updateSettingsStorageUi();
        if(settingsActiveTab==='view') syncSettingsPageGroupingUi();
      }
      async function updateSettingsStorageUi(){
        if(!settingsBodyEl||settingsActiveTab!=='storage') return;
        var textEl=settingsBodyEl.querySelector('#settingsStorageSummary');
        var noteEl=settingsBodyEl.querySelector('#settingsStorageNote');
        var linkRowEl=settingsBodyEl.querySelector('#settingsStorageLinkRow');
        var linkEl=settingsBodyEl.querySelector('#settingsStorageLink');
        var copyBtn=settingsBodyEl.querySelector('#settingsCopyStorageLinkBtn');
        var unlinkBtn=settingsBodyEl.querySelector('#unlinkWorkbookBtn');
        var loadExcelBtn=settingsBodyEl.querySelector('#loadStorageExcelBtn');
        var loadGoogleBtn=settingsBodyEl.querySelector('#loadStorageGoogleBtn');
        var loadUltraMainBtn=settingsBodyEl.querySelector('#loadStorageUltraMainBtn');
        var loadProtectedAircraftBtn=settingsBodyEl.querySelector('#loadProtectedAircraftBtn');
        var loadProtectedSupervisorsBtn=settingsBodyEl.querySelector('#loadProtectedSupervisorsBtn');
        var migrateExcelBtn=settingsBodyEl.querySelector('#migrateToExcelBtn');
        var migrateGoogleBtn=settingsBodyEl.querySelector('#migrateToGoogleBtn');
        if(!textEl||!unlinkBtn||!loadExcelBtn||!loadGoogleBtn||!loadUltraMainBtn||!loadProtectedAircraftBtn||!loadProtectedSupervisorsBtn||!migrateExcelBtn||!migrateGoogleBtn) return;
        var source=sourceType(activeStorageSource)!==STORAGE_SOURCE_NONE?activeStorageSource:loadStoredSource(),handle=await loadStoredHandle(LINKED_FILE_KEY);
        if(!settingsBodyEl.contains(textEl)) return;
        var downloadMode=!!(source&&source.downloadOnly&&sourceType(source)===STORAGE_SOURCE_EXCEL&&s(source.name));
        var linked=sourceType(source)===STORAGE_SOURCE_GOOGLE?!!s(source.spreadsheetId):(downloadMode||!!(handle&&handleIsWorkbook(handle)));
        if(sourceType(source)===STORAGE_SOURCE_GOOGLE){
          setLinkedWorkbookName(null);
          textEl.textContent=linked?('Google Sheet: '+(s(source.title)||s(source.spreadsheetId))+'.'):'No storage source is linked right now.';
          var url=googleSheetUrl(source);
          if(linkRowEl) linkRowEl.hidden=!url;
          if(linkEl){
            linkEl.href=url||'#';
            linkEl.textContent=url;
          }
          if(copyBtn) copyBtn.disabled=!url;
        } else if(downloadMode){
          setLinkedWorkbookName(null);
          textEl.textContent='Excel download mode: '+workbookFileName(source.name)+'.';
          if(linkRowEl) linkRowEl.hidden=true;
          if(linkEl){
            linkEl.removeAttribute('href');
            linkEl.textContent='';
          }
          if(copyBtn) copyBtn.disabled=true;
        } else {
          if(linked) setLinkedWorkbookName(handle);
          else setLinkedWorkbookName(null);
          textEl.textContent=linked?('Excel file: '+(linkedWorkbookName||'cap741-data.xlsx')+'.'):'No storage source is linked right now.';
          if(linkRowEl) linkRowEl.hidden=true;
          if(linkEl){
            linkEl.removeAttribute('href');
            linkEl.textContent='';
          }
          if(copyBtn) copyBtn.disabled=true;
        }
        unlinkBtn.disabled=!linked;
        loadExcelBtn.disabled=!linked;
        loadGoogleBtn.disabled=!linked;
        loadUltraMainBtn.disabled=!linked;
        loadProtectedAircraftBtn.disabled=!linked||!protectedDataAvailable();
        loadProtectedSupervisorsBtn.disabled=!linked||!protectedDataAvailable();
        migrateExcelBtn.disabled=!hasWorkbookDataLoaded()||sourceType(activeStorageSource)===STORAGE_SOURCE_EXCEL;
        migrateGoogleBtn.disabled=!hasWorkbookDataLoaded()||sourceType(activeStorageSource)===STORAGE_SOURCE_GOOGLE;
        if(noteEl){
          if(!linked) noteEl.textContent='Link a main storage source first. Import keeps that linked source before you save. UltraMain import adds mapped rows and skips matching Job No + Date entries. Prefilled aircraft and supervisor imports save straight into the linked source after import.';
          else if(downloadMode) noteEl.textContent='This browser saves Excel changes by downloading a fresh .xlsx file each time you tap Save. If Safari opens a preview first, use Share and choose Save to Files.';
          else if(!protectedDataAvailable()) noteEl.textContent='Prefilled aircraft and supervisor data is not available. The regular import and migration actions still work normally.';
          else if(!hasWorkbookDataLoaded()) noteEl.textContent='Import loads CAP741 data from another source into the app but keeps the current linked source for saving.';
          else if(sourceType(activeStorageSource)===STORAGE_SOURCE_GOOGLE) noteEl.textContent='Import replaces the in-app data from another source but keeps this Google Sheet as the save target. Prefilled aircraft and supervisor imports save straight into this linked source. Use "Migrate To Excel" if you want to switch to a new Excel file instead.';
          else if(sourceType(activeStorageSource)===STORAGE_SOURCE_EXCEL||sourceType(activeStorageSource)===STORAGE_SOURCE_DEFAULT) noteEl.textContent='Import replaces the in-app data from another source but keeps this Excel file as the save target. Prefilled aircraft and supervisor imports save straight into this linked source. Use "Migrate To Google Sheet" if you want to switch to a new Google Sheet instead.';
          else noteEl.textContent='Import replaces the in-app data but keeps the current linked source for saving. Migration copies the current data to a new source and switches the app to it.';
        }
      }
      async function unlinkRememberedWorkbook(){
        var removed=await removeStoredHandle(LINKED_FILE_KEY);
        var source=sourceType(activeStorageSource)!==STORAGE_SOURCE_NONE?activeStorageSource:loadStoredSource();
        if(!removed&&sourceType(source)!==STORAGE_SOURCE_GOOGLE&&!source.downloadOnly) throw new Error('The remembered workbook could not be cleared.');
        setAutoLoadDefaultWorkbook(false);
        setLinkedWorkbookName(null);
        googleAccessToken='';
        setActiveStorageSource({type:STORAGE_SOURCE_NONE},true);
        clearWorkbookState();
        syncLoadButtonAvailability(false);
        refreshUnsavedChangesState();
        renderAll();
        if(settingsModal&&settingsModal.className.indexOf('open')!==-1) renderSettingsBody(settingsActiveTab);
      }
      function openSettingsModal(){ if(!settingsModal||!settingsBodyEl) return; renderSettingsBody(settingsActiveTab); settingsModal.className='modal-backdrop open'; }
      function closeSettingsModal(){ if(settingsModal) settingsModal.className='modal-backdrop'; }
      function addSettingsRow(kind){ var tbody=settingsBodyEl&&settingsBodyEl.querySelector('[data-settings-table="'+kind+'"] tbody'); if(!tbody) return; var rowHtml=kind==='aircraft'?settingsTableRow(['<input type="text" data-col="Group" placeholder="Group">','<input type="text" data-col="A/C Reg" placeholder="G-XXXX">','<input type="text" data-col="Aircraft Type" placeholder="Boeing 777-300ER - GE90">'],kind):(kind==='chapters'?settingsTableRow(['<input type="text" data-col="Chapter" placeholder="e.g. 71">','<input type="text" data-col="Description" placeholder="e.g. Power Plant">'],kind):(kind==='flags'?settingsTableRow([flagSectionSelectHtml(FLAG_SECTION_PRIMARY),'<input type="text" data-col="Flag" placeholder="INSP - Inspections">',settingsFlagColorCell('Gold')],kind):settingsTableRow([supervisorPrintCheckboxHtml(false),'<input type="text" data-col="Signatory Name" placeholder="Name">','<input type="text" data-col="Stamp" placeholder="Stamp">','<input type="text" data-col="License Number" placeholder="Licence No.">','<input type="text" data-col="Scope / Limitations" placeholder="Scope">'],kind))); tbody.insertAdjacentHTML('beforeend',rowHtml); tbody.lastElementChild.querySelector('input[data-col], select[data-col]') && tbody.lastElementChild.querySelector('input[data-col], select[data-col]').focus(); }
      async function saveSettingsFromModal(){
        // Save owner info (only available on owner tab; store from DOM if on that tab, else use cached)
        var ownerNameEl=settingsBodyEl.querySelector('#settingsOwnerName');
        var ownerStampEl=settingsBodyEl.querySelector('#settingsOwnerStamp');
        if(ownerNameEl) LOG_OWNER_INFO.name=s(ownerNameEl.value);
        if(ownerStampEl) LOG_OWNER_INFO.stamp=s(ownerStampEl.value);
        if(settingsActiveTab==='view'){
          APP_VIEW_SETTINGS=cloneAppViewSettings(collectAppViewSettingsFromModal());
          syncMindMapButtonVisibility();
          syncRowsToReferenceFillMode();
        } else if(settingsActiveTab==='aircraft'){
          applyAircraftGroupRows(collectSettingsTable('aircraft').map(function(r){ return {group:r.Group,reg:r['A/C Reg'],type:r['Aircraft Type']}; }));
        } else if(settingsActiveTab==='chapters'){
          applyChapterRows(collectSettingsTable('chapters').map(function(r){ return {chapter:r.Chapter,description:r.Description}; }));
        } else if(settingsActiveTab==='flags'){
          applyFlagRows(collectSettingsTable('flags').map(function(r){ return {section:r.Section,flag:r.Flag,color:r.Color}; }));
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
      if(mindMapBtn) mindMapBtn.onclick=function(ev){ if(ev) ev.stopPropagation(); setLoadOptionsOpen(false); setPrintOptionsOpen(false); openMindMapFeature(); };
      if(closeMindMapModalBtn) closeMindMapModalBtn.onclick=function(){ closeMindMapModal(); };
      if(mindMapModal) mindMapModal.onclick=function(ev){ if(handleMindMapInteraction(ev.target)) return; if(ev.target===mindMapModal) closeMindMapModal(); };
      if(mindMapCanvasEl) mindMapCanvasEl.addEventListener('wheel',handleMindMapWheelEvent,{passive:false});
      if(mindMapCanvasEl) mindMapCanvasEl.addEventListener('pointerdown',function(ev){
        var viewport=currentMindMapViewport(),draggableNode=ev.target&&ev.target.closest&&ev.target.closest('[data-mindmap-draggable="1"]');
        if(!viewport||!viewport.contains(ev.target)) return;
        if(ev.pointerType==='touch'){
          setMindMapTouchPointer(ev.pointerId,ev.clientX,ev.clientY);
          if(mindMapTouchPointList().length>=2){
            ev.preventDefault();
            beginMindMapPinch();
            return;
          }
        }
        if(typeof ev.button==='number'&&ev.button!==0&&ev.pointerType!=='touch') return;
        if(draggableNode&&!(ev.target&&ev.target.closest&&ev.target.closest('[data-mindmap-inline-action="1"]'))){
          ev.preventDefault();
          beginMindMapNodeDrag(draggableNode.getAttribute('data-mindmap-node-key'),ev.pointerId,ev.clientX,ev.clientY);
          return;
        }
        if(!canStartMindMapDrag(ev.target)) return;
        ev.preventDefault();
        beginMindMapDrag(ev.pointerId,ev.clientX,ev.clientY);
      });
      if(mindMapCanvasEl) mindMapCanvasEl.addEventListener('pointermove',function(ev){
        var view=mindMapViewState(),nodeDrag=mindMapState.nodeDrag;
        if(ev.pointerType==='touch') setMindMapTouchPointer(ev.pointerId,ev.clientX,ev.clientY);
        if(mindMapState.pinch){
          ev.preventDefault();
          updateMindMapPinch();
          return;
        }
        if(nodeDrag&&nodeDrag.pointerId===ev.pointerId){
          ev.preventDefault();
          updateMindMapNodeDrag(ev.clientX,ev.clientY);
          return;
        }
        if(!view.dragging||view.pointerId!==ev.pointerId) return;
        ev.preventDefault();
        updateMindMapDrag(ev.clientX,ev.clientY);
      });
      if(mindMapCanvasEl) ['pointerup','pointercancel','lostpointercapture'].forEach(function(eventName){
        mindMapCanvasEl.addEventListener(eventName,function(ev){
          var view=mindMapViewState(),nodeDrag=mindMapState.nodeDrag;
          if(ev.pointerType==='touch') clearMindMapTouchPointer(ev.pointerId);
          if(mindMapState.pinch){
            endMindMapPinch(ev.pointerId);
            return;
          }
          if(nodeDrag&&nodeDrag.pointerId===ev.pointerId){
            endMindMapNodeDrag(ev.pointerId);
            return;
          }
          if(!view.dragging||view.pointerId!==ev.pointerId) return;
          endMindMapDrag(ev.pointerId);
        });
      });
      closeFilterPanelBtn.onclick=function(){ closeFilterPanel(); };
      clearFiltersBtn.onclick=function(){ resetDraftFilters(); };
      filterForm.onsubmit=function(ev){ ev.preventDefault(); activeFilters=readFilterForm(); closeFilterPanel(); renderAll(); };
      filterModal.onclick=function(ev){ if(ev.target===filterModal) closeFilterPanel(); };
      filterStripEl.onclick=function(ev){ if(ev.target&&ev.target.closest&&ev.target.closest('[data-clear-filters="1"]')) clearFilters(); };
      if(searchInput) searchInput.addEventListener('input',function(){ applySearchQuery(this.value); scheduleSearchRender(normalizedSearchQuery?110:0); });
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
      if(taskDetailModal) taskDetailModal.addEventListener('change',function(ev){ if(ev.target&&ev.target.matches&&ev.target.matches('[data-flag-option]')) previewTaskDetailForm(); });
      if(confirmCancelBtn) confirmCancelBtn.onclick=function(){ closeConfirmDialog(false); };
      if(confirmOkBtn) confirmOkBtn.onclick=function(){ closeConfirmDialog(true); };
      if(confirmModal) confirmModal.onclick=function(ev){ if(ev.target===confirmModal) closeConfirmDialog(false); };
      if(closeOtherLayoutModalBtn) closeOtherLayoutModalBtn.onclick=closeOtherLayoutModal;
      if(resetOtherLayoutDefaultsBtn) resetOtherLayoutDefaultsBtn.onclick=function(){ otherLayoutMeasurements=cloneMeasurementState(OTHER_LAYOUT_DEFAULTS); updateOtherLayoutPreview(otherLayoutMeasurements); if(otherOverlayTopMeasureValueEl) otherOverlayTopMeasureValueEl.focus(); };
      if(printOtherLayoutConfirmBtn) printOtherLayoutConfirmBtn.onclick=confirmOtherLayoutPrint;
      if(otherLayoutModal) otherLayoutModal.onclick=function(ev){ if(ev.target===otherLayoutModal) closeOtherLayoutModal(); };
      bindOtherLayoutMeasurementInput(otherOverlayTopMeasureValueEl,function(value){ otherLayoutMeasurements.top=value; });
      bindOtherLayoutMeasurementInput(otherOverlayRowMeasureValueEl,function(value){ otherLayoutMeasurements.rowHeight=value; });
      bindOtherLayoutMeasurementInput(otherOverlayTaskMeasureValueEl,function(value){ otherLayoutMeasurements.taskStart=value; });
      bindOtherLayoutMeasurementInput(otherOverlayTaskWidthMeasureValueEl,function(value){ otherLayoutMeasurements.superStart=otherLayoutMeasurements.taskStart+value; });
      bindOtherLayoutMeasurementInput(otherOverlayDateMeasureValueEl,function(value){
        otherLayoutMeasurements.dateStart=otherLayoutMeasurements.left;
        otherLayoutMeasurements.regStart=otherLayoutMeasurements.left+value;
      });
      bindOtherLayoutMeasurementInput(otherOverlayRegMeasureValueEl,function(value){ otherLayoutMeasurements.jobStart=otherLayoutMeasurements.regStart+value; });
      bindOtherLayoutMeasurementInput(otherOverlayJobMeasureValueEl,function(value){ otherLayoutMeasurements.taskStart=otherLayoutMeasurements.jobStart+value; });
      bindOtherLayoutMeasurementInput(otherOverlaySuperMeasureValueEl,function(value){ otherLayoutMeasurements.end=otherLayoutMeasurements.superStart+value; });
      window.addEventListener('resize',function(){ if(otherLayoutModal&&otherLayoutModal.classList.contains('open')) requestAnimationFrame(syncOtherOverlaySampleGuide); });
      window.addEventListener('resize',function(){
        if(!mindMapModal||!mindMapModal.classList.contains('open')) return;
        requestAnimationFrame(function(){
          centerMindMapView(false);
          syncMindMapView();
        });
      });
      if(closeGoogleSheetModalBtn) closeGoogleSheetModalBtn.onclick=function(){ closeGoogleSheetModal(null); };
      if(googleSheetCancelBtn) googleSheetCancelBtn.onclick=function(){ closeGoogleSheetModal(null); };
      if(googleSheetOkBtn) googleSheetOkBtn.onclick=function(){ closeGoogleSheetModal(googleSheetInputWrapEl&&!googleSheetInputWrapEl.hidden&&googleSheetUrlInputEl?googleSheetUrlInputEl.value:true); };
      if(googleSheetUrlInputEl) googleSheetUrlInputEl.addEventListener('keydown',function(ev){ if(ev.key==='Enter'){ ev.preventDefault(); closeGoogleSheetModal(googleSheetUrlInputEl.value); } });
      if(googleSheetModal) googleSheetModal.onclick=function(ev){ if(ev.target===googleSheetModal) closeGoogleSheetModal(null); };

      infoBtn.onclick=function(){ openInfoModal(); };
      closeInfoModalBtn.onclick=function(){ closeInfoModal(); };
      infoModal.onclick=function(ev){ if(ev.target===infoModal) closeInfoModal(); };

      printBtn.onclick=function(ev){ if(ev) ev.stopPropagation(); setLoadOptionsOpen(false); setPrintOptionsOpen(!(printOptionsEl&&printOptionsEl.classList.contains('open'))); };
      printCurrentBtn.onclick=function(){ setPrintOptionsOpen(false); printCurrentPage(); };
      if(printCurrentOverlayBtn) printCurrentOverlayBtn.onclick=function(){ setPrintOptionsOpen(false); printCurrentPageOverlay(); };
      if(printOtherLayoutBtn) printOtherLayoutBtn.onclick=function(){ setPrintOptionsOpen(false); openOtherLayoutModal(); };
      printAllBtn.onclick=function(){ setPrintOptionsOpen(false); printAllPages(); };
      document.addEventListener('click',function(ev){
        var insidePrint=!!(ev.target.closest&&(ev.target.closest('#printBtn')||ev.target.closest('#printOptions')));
        var insideLoad=!!(ev.target.closest&&(ev.target.closest('#loadBtn')||ev.target.closest('#loadOptions')));
        if(!insidePrint) setPrintOptionsOpen(false);
        if(!insideLoad) setLoadOptionsOpen(false);
      });
      document.addEventListener('keydown',function(ev){
        if(ev.key==='Escape'&&mindMapModal&&mindMapModal.className.indexOf('open')!==-1){
          ev.preventDefault();
          closeMindMapModal();
        }
      });

      // Save button - write xlsx
      saveFileBtn.onclick=async function(){ captureActiveEditorState(); setLoadingState(true,'Saving',sourceType(activeStorageSource)===STORAGE_SOURCE_GOOGLE?'Writing changes to Google Sheets...':('Writing changes to '+currentWorkbookFileName()+'...')); try { await flushLinkedRewrite(true); if(usingExcelDownloadFallback()) showExcelDownloadNote(currentWorkbookFileName(),'Excel file saved'); } finally { setLoadingState(false); } };

      // Load/Link button menu
      async function openExistingWorkbookSelection(sourceOverride){
        if(workbookOpenInFlight) return;
        workbookOpenInFlight=true;
        try {
          var persistHandle=persistentExcelLinkingSupported();
          setLoadOptionsOpen(false);
          clearFail();
          if(persistHandle) setLoadingState(true,loadButtonMode==='link'?'Linking file':'Loading','Waiting for Excel file selection...');
          var source=sourceOverride||await pickWorkbookSource(persistHandle);
          if(!source) return;
          if(!/\.xlsx$/i.test(s(source.name))) throw new Error('Please choose a .xlsx Excel file.');
          setLoadingState(true,'Loading','Reading workbook...');
          loadWorkbookFromArrayBuffer(await readFileArrayBuffer(source.file));
          if(persistHandle&&source.handle){
            setAutoLoadDefaultWorkbook(false);
            setActiveStorageSource({type:STORAGE_SOURCE_EXCEL,name:s(source.name)},true);
            syncLoadButtonAvailability(true);
          } else {
            setSessionExcelSource(source.name,false);
            syncLoadButtonAvailability(false);
          }
          await renderAllWithLoading('Loading logbook','Rendering pages...');
          refreshUnsavedChangesState();
        } catch(e){
          if(e.name!=='AbortError') fail('Could not open Excel file: '+e.message);
        } finally {
          if(workbookOpenInput) workbookOpenInput.value='';
          workbookOpenInFlight=false;
          setLoadingState(false);
        }
      }
      loadBtn.onclick=function(ev){
        if(ev) ev.stopPropagation();
        setPrintOptionsOpen(false);
        setLoadOptionsOpen(!(loadOptionsEl&&loadOptionsEl.classList.contains('open')));
      };
      if(loadExistingBtn) loadExistingBtn.addEventListener('click',function(ev){
        if(!persistentExcelLinkingSupported()) return;
        ev.preventDefault();
        ev.stopPropagation();
        openExistingWorkbookSelection();
      });
      if(workbookOpenInput) workbookOpenInput.addEventListener('click',function(){ if(!persistentExcelLinkingSupported()){ clearFail(); workbookOpenInput.value=''; } });
      if(workbookOpenInput) workbookOpenInput.addEventListener('change',async function(){
        if(persistentExcelLinkingSupported()) return;
        var file=workbookOpenInput.files&&workbookOpenInput.files[0];
        workbookOpenInput.value='';
        if(!file) return;
        try {
          await openExistingWorkbookSelection({handle:null,file:file,name:s(file.name)});
        } catch(e){
          if(e.name!=='AbortError') fail('Could not open Excel file: '+e.message);
        }
      });
      if(createNewWorkbookBtn) createNewWorkbookBtn.onclick=async function(ev){
        if(ev) ev.stopPropagation();
        try {
          setLoadOptionsOpen(false);
          clearFail();
          setLoadingState(true,'Creating logbook',fileSavePickerSupported()?'Choose where to save the new cap741-data.xlsx file...':'Preparing a new CAP741 Excel download...');
          await createNewWorkbookFile();
        } catch(e){
          if(e.name!=='AbortError') fail('Could not create Excel file: '+e.message);
        } finally {
          setLoadingState(false);
        }
      };
      if(loadGoogleSheetBtn) loadGoogleSheetBtn.onclick=async function(ev){
        if(ev) ev.stopPropagation();
        try {
          setLoadOptionsOpen(false);
          clearFail();
          await connectExistingGoogleSheet();
        } catch(e){
          if(e.name!=='AbortError') fail('Could not open Google Sheet: '+e.message);
        } finally {
          setLoadingState(false);
        }
      };
      if(createGoogleSheetBtn) createGoogleSheetBtn.onclick=async function(ev){
        if(ev) ev.stopPropagation();
        try {
          setLoadOptionsOpen(false);
          clearFail();
          setLoadingState(true,'Creating Google Sheet','Waiting for Google authorization...');
          await createNewGoogleSheet();
        } catch(e){
          if(e.name!=='AbortError') fail('Could not create Google Sheet: '+e.message);
        } finally {
          setLoadingState(false);
        }
      };

      // Settings
      if(settingsBtn) settingsBtn.onclick=openSettingsModal;
      if(closeSettingsModalBtn) closeSettingsModalBtn.onclick=closeSettingsModal;
      if(saveSettingsBtn) saveSettingsBtn.onclick=saveSettingsFromModal;
      if(printSupervisorsBtn) printSupervisorsBtn.onclick=printSupervisorList;
      if(errorOkBtn) errorOkBtn.onclick=function(ev){ if(ev) ev.stopPropagation(); clearFail(); };
      if(errorBox) errorBox.addEventListener('click',function(ev){
        if(!(ev.target&&ev.target.closest&&ev.target.closest('.error-ok'))) clearFail();
      });
      if(settingsModal) settingsModal.onclick=function(ev){ if(ev.target===settingsModal) closeSettingsModal(); };
      if(settingsBodyEl) settingsBodyEl.addEventListener('change',function(ev){
        var groupingInput=ev.target&&ev.target.closest&&ev.target.closest('input[name="settingsPageGrouping"]');
        if(groupingInput){
          syncSettingsPageGroupingUi();
          return;
        }
      });
      if(settingsBodyEl) settingsBodyEl.addEventListener('input',function(ev){
        var colorInput=ev.target&&ev.target.closest&&ev.target.closest('[data-col="Color"]');
        if(!colorInput) return;
        var dot=colorInput.parentElement&&colorInput.parentElement.querySelector('.settings-flag-color-dot');
        if(dot) dot.style.backgroundColor=s(colorInput.value)||'#7f93a1';
      });
      if(settingsBodyEl) settingsBodyEl.addEventListener('click',function(ev){
        var unlink=ev.target.closest&&ev.target.closest('[data-settings-unlink]');
        if(unlink){
          clearFail();
          unlinkRememberedWorkbook().catch(function(e){ fail('Could not unlink storage source: '+e.message); });
          return;
        }
        var copyLinkBtn=ev.target.closest&&ev.target.closest('[data-settings-copy-link]');
        if(copyLinkBtn){
          var linkEl=settingsBodyEl.querySelector('#settingsStorageLink');
          copyTextToClipboard(linkEl&&linkEl.textContent||'').then(function(copied){
            if(!copied) throw new Error('The Google Sheet link could not be copied.');
            flashStorageCopyButton(copyLinkBtn);
          }).catch(function(e){
            fail('Could not copy link: '+e.message);
          });
          return;
        }
        var storageAction=ev.target.closest&&ev.target.closest('[data-settings-storage-action]');
        if(storageAction){
          var action=storageAction.getAttribute('data-settings-storage-action');
          clearFail();
          captureActiveEditorState();
          var actionPromise;
          if(action==='import-google'){
            actionPromise=importDataFromGoogleSheetForLinkedSource();
          } else if(action==='import-excel'){
            actionPromise=importDataFromExcelForLinkedSource();
          } else if(action==='import-ultramain'){
            actionPromise=importUltraMainForLinkedSource();
          } else if(action==='import-protected-aircraft'){
            actionPromise=importProtectedAircraftForLinkedSource();
          } else if(action==='import-protected-supervisors'){
            actionPromise=importProtectedSupervisorsForLinkedSource();
          } else if(action==='migrate-google'){
            setLoadingState(true,'Migrating storage','Creating a Google Sheet from the current CAP741 data...');
            actionPromise=migrateCurrentDataToGoogleSheet();
          } else {
            setLoadingState(true,'Migrating storage',fileSavePickerSupported()?'Choosing where to save the migrated Excel file...':'Preparing the migrated Excel download...');
            actionPromise=migrateCurrentDataToExcel();
          }
          actionPromise.then(function(changed){
            if(!changed) return;
            renderAll();
            if(action.indexOf('import-')===0){
              closeSettingsModal();
              return;
            }
            if(settingsModal&&settingsModal.className.indexOf('open')!==-1) renderSettingsBody(settingsActiveTab);
          }).catch(function(e){
            if(e&&e.name!=='AbortError') fail((action.indexOf('import-')===0?'Could not import data: ':'Could not update storage: ')+e.message);
          }).finally(function(){
            setLoadingState(false);
          });
          return;
        }
        var add=ev.target.closest&&ev.target.closest('[data-settings-add]');
        if(add){ addSettingsRow(add.getAttribute('data-settings-add')); return; }
        var rem=ev.target.closest&&ev.target.closest('[data-settings-remove]');
        if(rem){ var tr=rem.closest('tr'); if(tr) tr.remove(); }
      });

      // Page editing events
      pagesEl.addEventListener('focusin',function(ev){ var cell=ev.target.closest&&ev.target.closest('.editable-cell'); if(cell&&cell.innerHTML==='&nbsp;') cell.innerHTML=''; });
      pagesEl.addEventListener('click',function(ev){
        if(ev.target&&ev.target.closest&&ev.target.closest('[data-open-load-menu="1"]')){ if(loadBtn&&loadBtn.style.display!=='none'){ setLoadOptionsOpen(true); } return; }
        if(ev.target&&ev.target.closest&&ev.target.closest('[data-clear-search="1"]')){ clearSearch(); return; }
        if(ev.target&&ev.target.closest&&ev.target.closest('[data-clear-all-results="1"]')){ clearFilters(); clearSearch(); return; }
        if(ev.target&&ev.target.closest&&ev.target.closest('[data-clear-filters="1"]')){ clearFilters(); return; }
        var signToggleBtn=ev.target.closest&&ev.target.closest('[data-toggle-signed]');
        if(signToggleBtn){ toggleSignedRow(signToggleBtn); return; }
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
      pagesEl.addEventListener('mousedown',function(ev){ var actionBtn=ev.target.closest&&ev.target.closest('[data-open-task], [data-open-task-new], [data-clear-supervisor], [data-toggle-signed]'); if(actionBtn) ev.preventDefault(); });
      pagesEl.addEventListener('change',function(ev){
        var datePicker=ev.target.closest&&ev.target.closest('[data-date-picker]');
        if(datePicker){ var dateEntry=datePicker.closest('.date-entry'); syncDateControl(dateEntry,datePicker.value); var dateInput=dateEntry&&dateEntry.querySelector('[data-date-text]'); if(dateInput&&updateRowFromEditor(dateInput)) syncSaveButtonState(false); }
        var supervisorInput=ev.target.closest&&ev.target.closest('[data-edit-field="Approval Name"], [data-new-row][data-edit-field="Approval Name"]');
        if(supervisorInput){ var supTd=supervisorInput.closest('td'),licenceInput=supTd&&supTd.querySelector('[data-edit-field="Aprroval Licence No."], [data-new-row][data-edit-field="Aprroval Licence No."]'); applySupervisorSuggestion(supervisorInput,licenceInput); updateRowFromEditor(supervisorInput); syncSaveButtonState(false); }
        var groupInput=ev.target.closest&&ev.target.closest('[data-group-field]');
        if(!groupInput) return;
        var page=groupInput.closest('.page');
        if(!page) return;
        var groupKey=page.getAttribute('data-group-key'),grpRows=rowsByGroupKey(groupKey);
        if(!grpRows.length) return;
        var mode=currentPageGrouping(),nextType=defaultNewAircraftTypeForRows(grpRows,mode),nextChapter=grpRows[0]['Chapter'],nextChapterDesc=grpRows[0]['Chapter Description'],nextGroupLabel=mode===PAGE_GROUPING_GROUP?(s(page.getAttribute('data-page-group-label'))||rowPageGroupingLabel(grpRows[0],mode)):'';
        if(groupInput.getAttribute('data-group-field')==='Aircraft Type'){
          nextType=valueOf(groupInput);
          for(var i=0;i<grpRows.length;i++) grpRows[i]['Aircraft Type']=nextType;
        } else if(groupInput.getAttribute('data-group-field')==='Chapter'){
          var parsedChapter=parseChapterValue(valueOf(groupInput)),completedChapter=completeChapterParts(parsedChapter.chapter,parsedChapter.chapterDesc);
          nextChapter=completedChapter.chapter; nextChapterDesc=referenceOnlySaveEnabled()?s(parsedChapter.chapterDesc):completedChapter.chapterDesc;
          for(var j=0;j<grpRows.length;j++){ grpRows[j]['Chapter']=nextChapter; grpRows[j]['Chapter Description']=nextChapterDesc; grpRows[j].__manualChapterDescription=!!s(parsedChapter.chapterDesc); }
        }
        if(mode===PAGE_GROUPING_GROUP){
          for(var k=0;k<grpRows.length;k++) grpRows[k].__pageGroupLabel=nextGroupLabel;
        }
        updateRowsDirtyState(grpRows);
        syncBlankRowMetadata(page,nextType,nextChapter,nextChapterDesc,rowPageGroupingKey(grpRows[0],mode),nextGroupLabel,currentSingleAircraftRegFilter());
        renderAll();
        refreshUnsavedChangesState();
        scheduleAutoSave();
      });
      pagesEl.addEventListener('input',function(ev){
        if(ev.target&&ev.target.classList&&ev.target.classList.contains('dots-input')) syncDotsInputSize(ev.target);
        if(ev.target&&ev.target.classList&&ev.target.classList.contains('ref-view-input')) syncFieldInputViewState(ev.target);
        var cell=ev.target.closest&&(ev.target.closest('.editable-cell')||ev.target.closest('[data-row-id]')||ev.target.closest('[data-new-row]'));
        if(!cell) return;
        if(ev.target.matches&&ev.target.matches('[data-group-field]')) return;
        var field=cell.getAttribute('data-edit-field');
        var existingRow=rowById(cell.getAttribute('data-row-id'));
        var unitsBefore=liveLayoutUnitsForField(existingRow,field);
        var row=updateRowFromEditor(cell);
        if(fieldNeedsLiveLayoutRefresh(field)&&liveLayoutUnitsForField(row,field)!==unitsBefore){
          scheduleLiveLayoutRefresh(captureEditorSnapshot(ev.target),140);
        }
      });
      pagesEl.addEventListener('blur',function(ev){
        var cell=ev.target.closest&&(ev.target.closest('.editable-cell')||ev.target.closest('[data-row-id]')||ev.target.closest('[data-new-row]'));
        if(!cell) return;
        var wasNewRow=cell.hasAttribute('data-new-row'),field=cell.getAttribute('data-edit-field');
        var row=updateRowFromEditor(cell);
        if(!row) return;
        if(!rowHasEntryContent(row)&&comparableRowSignature(row)!==(savedComparableRowSignature(row)||'')){
          removeRowById(row.__rowId);
          refreshUnsavedChangesState();
          renderAllWithMotion();
          scheduleAutoSave();
          return;
        }
        if(wasNewRow||fieldAffectsRowLayout(field)) scheduleLayoutRefresh(250);
      },true);
      window.addEventListener('afterprint',function(){ clearPrintSelection(); setPrintOptionsOpen(false); });
      window.addEventListener('resize',function(){ if(modal.className.indexOf('open')!==-1) fitModalPreview(); });
      window.addEventListener('beforeunload',function(){ captureActiveEditorState(); });

      // ---- Startup ----
      rows=normalizeRows(rows);
      FLAG_RECORDS=defaultFlagRecords();

      (async function cap741Startup(){
        setLoadingState(true,'Loading logbook','Reading cap741-data.xlsx...');
        await nextPaint();
        var loaded=false;
        var linked=false;
        var storedSource=loadStoredSource();

        // 1. Prefer the remembered Google Sheet when one exists.
        if(sourceType(storedSource)===STORAGE_SOURCE_GOOGLE){
          try {
            await loadGoogleSheetState(storedSource,false);
            loaded=true;
            linked=true;
          } catch(googleErr){
            googleAccessToken='';
            setActiveStorageSource({type:STORAGE_SOURCE_NONE},false);
          }
        }

        // 2. Prefer the remembered linked workbook when one exists.
        if(!loaded){
          try {
            var storedHandle=await loadStoredHandle(LINKED_FILE_KEY);
            if(storedHandle&&handleIsWorkbook(storedHandle)&&await ensurePermission(storedHandle)){
              setLinkedWorkbookName(storedHandle);
              setActiveStorageSource({type:STORAGE_SOURCE_EXCEL,name:s(storedHandle.name)},false);
              var file=await storedHandle.getFile();
              loadWorkbookFromArrayBuffer(await file.arrayBuffer());
              loaded=true;
              linked=true;
            }
          } catch(handleErr){}
        }

        if(!loaded){
          // If nothing is linked, stay blank and let the user choose a source.
          setLinkedWorkbookName(null);
          setActiveStorageSource({type:STORAGE_SOURCE_NONE},false);
          clearWorkbookState();
          syncLoadButtonAvailability(false);
          setLoadingState(false);
          renderAll();
          return;
        }

        syncLoadButtonAvailability(linked);
        renderAll();
        refreshUnsavedChangesState();
        setLoadingState(false);
      })();
    })();
