/**
 * VennFlow AI - Client Application Logic
 * Dual-Mode (Offline Local ML vs AselixAI API), Excel 5-Row Preview & Column Picker,
 * Copy-Paste Multi-Row Comparison, Interactive SVG Venn Diagram & QA Assistant.
 */

(function () {
  'use strict';

  // Config State (LocalStorage backed)
  const STORAGE_KEY = 'vennflow_aselix_config_v2';
  const DEFAULT_ALLOWED_ORIGINS = [
    'http://localhost:3004',
    'https://localhost:3004',
    'http://127.0.0.1:3004',
    'https://kokpit.sirketiniz.com',
    'https://kokpit.aselsan.com.tr',
    'http://kokpit.aselsan.com.tr'
  ];
  const DEFAULT_CONFIG = {
    network: 'milli',
    baseUrl: 'https://api.aselixai.aselsan.com.tr/v1',
    apiKey: '',
    model: 'aselix-llm-v1',
    embeddingModel: 'text-embedding-3-small',
    engine: 'local', // 'local', 'embedding', or 'aselix'
    threshold: 80,   // %50 - %100
    batchSize: 15,   // 10, 15, 20 satırlık güvenli batch boyutu
    inputMode: 'excel', // 'excel' or 'paste'
    allowedOrigins: DEFAULT_ALLOWED_ORIGINS.join(', ')
  };

  let appConfig = loadConfig();
  let file1 = null;
  let file2 = null;
  let previewData1 = null;
  let previewData2 = null;
  let currentAnalysis = null;
  let currentFilter = 'all';

  // Kokpit SSO / Iframe Push Identity & External Inputs State
  let kokpitIdentity = {
    user: null,
    token: null,
    loading: true,
    source: null,
    origin: null,
    timestamp: null,
    externalInputs: {}
  };

  // Network URLs mapping
  const NETWORK_URLS = {
    milli: 'https://api.aselixai.aselsan.com.tr/v1',
    aynet: 'https://agate.aselsan.com.tr/send/milli/api.aselixai/v1'
  };

  // Theme Elements
  const btnThemeToggle = document.getElementById('btnThemeToggle');
  const themeToggleIcon = document.getElementById('themeToggleIcon');
  const themeToggleText = document.getElementById('themeToggleText');

  // DOM Elements
  const engineStatusPill = document.getElementById('engineStatusPill');
  const engineStatusLabel = document.getElementById('engineStatusLabel');
  const btnEngineLocal = document.getElementById('btnEngineLocal');
  const btnEngineEmbedding = document.getElementById('btnEngineEmbedding');
  const btnEngineAselix = document.getElementById('btnEngineAselix');
  const thresholdSlider = document.getElementById('thresholdSlider');
  const thresholdValueBadge = document.getElementById('thresholdValueBadge');
  const currentModelDisplay = document.getElementById('currentModelDisplay');
  const chatModelNameDisplay = document.getElementById('chatModelNameDisplay');

  // Input Tabs
  const tabBtnExcel = document.getElementById('tabBtnExcel');
  const tabBtnPaste = document.getElementById('tabBtnPaste');
  const tabContentExcel = document.getElementById('tabContentExcel');
  const tabContentPaste = document.getElementById('tabContentPaste');

  // Excel Dropzones & Column Pickers
  const dropzone1 = document.getElementById('dropzone1');
  const dropzone2 = document.getElementById('dropzone2');
  const fileInput1 = document.getElementById('fileInput1');
  const fileInput2 = document.getElementById('fileInput2');
  const dropzoneContent1 = document.getElementById('dropzoneContent1');
  const dropzoneContent2 = document.getElementById('dropzoneContent2');
  const fileView1 = document.getElementById('fileView1');
  const fileView2 = document.getElementById('fileView2');
  const fileName1 = document.getElementById('fileName1');
  const fileName2 = document.getElementById('fileName2');
  const fileSub1 = document.getElementById('fileSub1');
  const fileSub2 = document.getElementById('fileSub2');
  const btnRemoveFile1 = document.getElementById('btnRemoveFile1');
  const btnRemoveFile2 = document.getElementById('btnRemoveFile2');

  const colPickBox1 = document.getElementById('colPickBox1');
  const colPickBox2 = document.getElementById('colPickBox2');
  const colSelect1 = document.getElementById('colSelect1');
  const colSelect2 = document.getElementById('colSelect2');
  const btnTogglePreview1 = document.getElementById('btnTogglePreview1');
  const btnTogglePreview2 = document.getElementById('btnTogglePreview2');
  const previewCard1 = document.getElementById('previewCard1');
  const previewCard2 = document.getElementById('previewCard2');
  const previewRowCountBadge1 = document.getElementById('previewRowCountBadge1');
  const previewRowCountBadge2 = document.getElementById('previewRowCountBadge2');
  const previewTableScroll1 = document.getElementById('previewTableScroll1');
  const previewTableScroll2 = document.getElementById('previewTableScroll2');

  // Copy-Paste Area
  const textPaste1 = document.getElementById('textPaste1');
  const textPaste2 = document.getElementById('textPaste2');
  const lineCount1 = document.getElementById('lineCount1');
  const lineCount2 = document.getElementById('lineCount2');

  // Action Buttons
  const btnRunComparison = document.getElementById('btnRunComparison');
  const btnRunText = document.getElementById('btnRunText');
  const btnUseSampleFiles = document.getElementById('btnUseSampleFiles');
  const progressBanner = document.getElementById('progressBanner');
  const progressStatusText = document.getElementById('progressStatusText');
  const progressSubText = document.getElementById('progressSubText');
  const progressBarWrapper = document.getElementById('progressBarWrapper');
  const progressBarFill = document.getElementById('progressBarFill');
  const progressPctBadge = document.getElementById('progressPctBadge');
  const actionHintText = document.getElementById('actionHintText');
  const batchLogPanel = document.getElementById('batchLogPanel');
  const btnCloseProgress = document.getElementById('btnCloseProgress');

  // Results Section
  const resultsSection = document.getElementById('resultsSection');
  const resultsSubTitle = document.getElementById('resultsSubTitle');
  const summaryEngineBadge = document.getElementById('summaryEngineBadge');
  const summaryText = document.getElementById('summaryText');
  const btnCopySummary = document.getElementById('btnCopySummary');
  const vennContainer = document.getElementById('vennContainer');

  // Metrics
  const statTotal = document.getElementById('statTotal');
  const statIntersection = document.getElementById('statIntersection');
  const statJaccard = document.getElementById('statJaccard');
  const statOnlyA = document.getElementById('statOnlyA');
  const statOnlyASub = document.getElementById('statOnlyASub');
  const statOnlyB = document.getElementById('statOnlyB');
  const statOnlyBSub = document.getElementById('statOnlyBSub');
  const legendTextA = document.getElementById('legendTextA');
  const legendTextB = document.getElementById('legendTextB');

  // Table & Filters
  const countAll = document.getElementById('countAll');
  const countIntersection = document.getElementById('countIntersection');
  const countA = document.getElementById('countA');
  const countB = document.getElementById('countB');
  const countDeclined = document.getElementById('countDeclined');
  const tabFilterAll = document.getElementById('tabFilterAll');
  const tabFilterIntersection = document.getElementById('tabFilterIntersection');
  const tabFilterA = document.getElementById('tabFilterA');
  const tabFilterB = document.getElementById('tabFilterB');
  const tabFilterDeclined = document.getElementById('tabFilterDeclined');
  const tableSearchInput = document.getElementById('tableSearchInput');
  const colFilterDoc1 = document.getElementById('colFilterDoc1');
  const colFilterDoc2 = document.getElementById('colFilterDoc2');
  const colFilterStatus = document.getElementById('colFilterStatus');
  const btnClearDetailColFilters = document.getElementById('btnClearDetailColFilters');
  const detailFilteredCountBadge = document.getElementById('detailFilteredCountBadge');
  const detailsTableBody = document.getElementById('detailsTableBody');
  const btnExportCSV = document.getElementById('btnExportCSV');
  const btnExportJSON = document.getElementById('btnExportJSON');
  const btnHeaderKpiTable = document.getElementById('btnHeaderKpiTable');
  const btnSaveToKpiDb = document.getElementById('btnSaveToKpiDb');
  const btnToggleKpiRecords = document.getElementById('btnToggleKpiRecords');
  const btnRefreshKpiRecords = document.getElementById('btnRefreshKpiRecords');
  const btnClearKpiFilters = document.getElementById('btnClearKpiFilters');
  const kpiRecordsCountBadge = document.getElementById('kpiRecordsCountBadge');
  const kpiFilterProje = document.getElementById('kpiFilterProje');
  const kpiFilterSd = document.getElementById('kpiFilterSd');
  const kpiFilterKalem = document.getElementById('kpiFilterKalem');
  const kpiFilterFaaliyet2 = document.getElementById('kpiFilterFaaliyet2');
  const kpiFilterUrun = document.getElementById('kpiFilterUrun');
  const kpiFilterSearch = document.getElementById('kpiFilterSearch');
  const dbKpiPreviewStrip = document.getElementById('dbKpiPreviewStrip');
  const kpiSaveStatusBanner = document.getElementById('kpiSaveStatusBanner');
  const kpiRecordsModal = document.getElementById('kpiRecordsModal');
  const btnCloseKpiModal = document.getElementById('btnCloseKpiModal');
  const kpiModalJsonPreviewBox = document.getElementById('kpiModalJsonPreviewBox');
  const kpiModalJsonTitle = document.getElementById('kpiModalJsonTitle');
  const kpiModalJsonContent = document.getElementById('kpiModalJsonContent');
  const btnCloseKpiJsonPreview = document.getElementById('btnCloseKpiJsonPreview');
  const kpiSavedRecordsBody = document.getElementById('kpiSavedRecordsBody');

  // Main Page Tabs & Home Page Elements
  const pageTabBtnHome = document.getElementById('pageTabBtnHome');
  const pageTabBtnVennflow = document.getElementById('pageTabBtnVennflow');
  const pageTabBtnBulk = document.getElementById('pageTabBtnBulk');
  const btnNavOpenKpiPopup = document.getElementById('btnNavOpenKpiPopup');
  const pageViewHome = document.getElementById('pageViewHome');
  const pageViewVennflow = document.getElementById('pageViewVennflow');
  const pageViewBulk = document.getElementById('pageViewBulk');
  const sharedConfigBarSection = document.getElementById('sharedConfigBarSection');
  const btnHeroGoVennflow = document.getElementById('btnHeroGoVennflow');
  const btnHeroGoBulk = document.getElementById('btnHeroGoBulk');
  const btnHeroOpenKpiPopup = document.getElementById('btnHeroOpenKpiPopup');
  const btnCardGoVennflow = document.getElementById('btnCardGoVennflow');
  const btnCardGoBulk = document.getElementById('btnCardGoBulk');
  const btnCardOpenKpiPopup = document.getElementById('btnCardOpenKpiPopup');
  const homeStatKpiCount = document.getElementById('homeStatKpiCount');
  const homeStatBakiyeCount = document.getElementById('homeStatBakiyeCount');
  const homeStatOnlyA = document.getElementById('homeStatOnlyA');
  const homeStatCommon = document.getElementById('homeStatCommon');
  const homeStatOnlyB = document.getElementById('homeStatOnlyB');

  // Bulk Data Entry (Toplu Veri Girişi) Elements
  const bulkTemplateProjeSelect = document.getElementById('bulkTemplateProjeSelect');
  const btnDownloadBulkTemplate = document.getElementById('btnDownloadBulkTemplate');
  const bulkFallbackBakiyeSelect = document.getElementById('bulkFallbackBakiyeSelect');
  const bulkDropzone = document.getElementById('bulkDropzone');
  const bulkFileInput = document.getElementById('bulkFileInput');
  const bulkDropzoneContent = document.getElementById('bulkDropzoneContent');
  const bulkFileView = document.getElementById('bulkFileView');
  const bulkFileName = document.getElementById('bulkFileName');
  const bulkFileSub = document.getElementById('bulkFileSub');
  const btnRemoveBulkFile = document.getElementById('btnRemoveBulkFile');
  const btnRunBulkAnalysis = document.getElementById('btnRunBulkAnalysis');
  const bulkResultsSection = document.getElementById('bulkResultsSection');
  const bulkResultsSubtitle = document.getElementById('bulkResultsSubtitle');
  const btnSaveAllBulkGroups = document.getElementById('btnSaveAllBulkGroups');
  const bulkSaveStatusBanner = document.getElementById('bulkSaveStatusBanner');
  const bulkGroupsContainer = document.getElementById('bulkGroupsContainer');

  let bulkFile = null;
  let bulkAnalyzedGroups = [];

  // Chat Section
  const chatMessages = document.getElementById('chatMessages');
  const chatInput = document.getElementById('chatInput');
  const btnSendChat = document.getElementById('btnSendChat');
  const btnClearChat = document.getElementById('btnClearChat');

  // Settings Modal & Admin-Only Elements
  const btnOpenSettings = document.getElementById('btnOpenSettings');
  const btnSwaggerDocs = document.getElementById('btnSwaggerDocs');
  const configPillsGroup = document.getElementById('configPillsGroup');
  const settingsModal = document.getElementById('settingsModal');
  const btnCloseSettings = document.getElementById('btnCloseSettings');
  const btnSaveSettings = document.getElementById('btnSaveSettings');
  const btnResetSettings = document.getElementById('btnResetSettings');
  const btnTestConn = document.getElementById('btnTestConn');
  const btnFetchModels = document.getElementById('btnFetchModels');
  const cfgBaseUrl = document.getElementById('cfgBaseUrl');
  const cfgApiKey = document.getElementById('cfgApiKey');
  const btnToggleApiKey = document.getElementById('btnToggleApiKey');
  const cfgModelSelect = document.getElementById('cfgModelSelect');
  const cfgBatchSize = document.getElementById('cfgBatchSize');
  const cfgEmbeddingModel = document.getElementById('cfgEmbeddingModel');
  const fetchModelsStatus = document.getElementById('fetchModelsStatus');
  const testConnResult = document.getElementById('testConnResult');
  const radioMilli = document.getElementById('radioMilli');
  const radioAynet = document.getElementById('radioAynet');
  const radioCustom = document.getElementById('radioCustom');

  // Kokpit SSO / Identity & External Input Elements
  const btnOpenIdentity = document.getElementById('btnOpenIdentity');
  const identityModal = document.getElementById('identityModal');
  const btnCloseIdentity = document.getElementById('btnCloseIdentity');
  const authStatusEl = document.getElementById('auth-status');
  const identityAvatar = document.getElementById('identityAvatar');
  const identityUserMeta = document.getElementById('identityUserMeta');
  const modalIdentityAvatar = document.getElementById('modalIdentityAvatar');
  const modalIdentityName = document.getElementById('modalIdentityName');
  const modalIdentitySub = document.getElementById('modalIdentitySub');
  const modalIdentityBadge = document.getElementById('modalIdentityBadge');
  const identityDetailsGrid = document.getElementById('identityDetailsGrid');
  const cfgAllowedOrigins = document.getElementById('cfgAllowedOrigins');
  const simPostMessagePayload = document.getElementById('simPostMessagePayload');
  const btnFillStandardUserPayload = document.getElementById('btnFillStandardUserPayload');
  const btnFillDefaultIdentityPayload = document.getElementById('btnFillDefaultIdentityPayload');
  const btnSimulatePostMessage = document.getElementById('btnSimulatePostMessage');
  const btnLogoutIdentity = document.getElementById('btnLogoutIdentity');
  const btnSaveIdentityConfig = document.getElementById('btnSaveIdentityConfig');
  const identityActionStatus = document.getElementById('identityActionStatus');

  // UGESUDKOKPİT Bakiye Sipariş (teslimat_md.bakiye_siparis) & Faz Seçici Elements
  const bakiyeProjeSelect1 = document.getElementById('bakiyeProjeSelect1');
  const bakiyeProjeSelect2 = document.getElementById('bakiyeProjeSelect2');
  const bakiyeSearchInput1 = document.getElementById('bakiyeSearchInput1');
  const bakiyeSearchInput2 = document.getElementById('bakiyeSearchInput2');
  const bakiyeItemSelect1 = document.getElementById('bakiyeItemSelect1');
  const bakiyeItemSelect2 = document.getElementById('bakiyeItemSelect2');
  const infoProje1 = document.getElementById('infoProje1');
  const infoProje2 = document.getElementById('infoProje2');
  const infoSdKalem1 = document.getElementById('infoSdKalem1');
  const infoSdKalem2 = document.getElementById('infoSdKalem2');
  const infoUrun1 = document.getElementById('infoUrun1');
  const infoUrun2 = document.getElementById('infoUrun2');
  const phaseGroup1 = document.getElementById('phaseGroup1');
  const phaseGroup2 = document.getElementById('phaseGroup2');
  const btnCopyBakiye1To2 = document.getElementById('btnCopyBakiye1To2');
  const dropBadge1 = document.getElementById('dropBadge1');
  const dropBadge2 = document.getElementById('dropBadge2');
  const bakiyeDbSource1 = document.getElementById('bakiyeDbSource1');
  const comparisonMetaBanner = document.getElementById('comparisonMetaBanner');

  let bakiyeState = {
    allItems: [],
    projeler: [],
    selectedItem1: null,
    selectedItem2: null,
    phase1: 'Dry-Run',
    phase2: 'FKT'
  };

  // --------------------------------------------------------------------------
  // Initialization & Theme Handling
  // --------------------------------------------------------------------------
  function init() {
    initTheme();
    applyConfigToUI();
    applyAdminVisibility();
    setupEventListeners();
    initKokpitIdentityBridge();
    loadBakiyeSiparisItems();
    refreshHomePageStats();
    switchMainPageTab('home');
  }

  function switchMainPageTab(tabName) {
    const activeTab = ['home', 'vennflow', 'bulk'].includes(tabName) ? tabName : 'home';
    document.body.classList.toggle('on-home-page', activeTab === 'home');

    [pageTabBtnHome, pageTabBtnVennflow, pageTabBtnBulk].forEach(btn => {
      if (btn) btn.classList.toggle('active', btn.dataset.pageTab === activeTab);
    });

    if (pageViewHome) pageViewHome.style.display = activeTab === 'home' ? 'block' : 'none';
    if (pageViewVennflow) pageViewVennflow.style.display = activeTab === 'vennflow' ? 'block' : 'none';
    if (pageViewBulk) pageViewBulk.style.display = activeTab === 'bulk' ? 'block' : 'none';
    if (sharedConfigBarSection) {
      sharedConfigBarSection.style.display = (activeTab === 'vennflow' || activeTab === 'bulk') ? 'block' : 'none';
    }

    if (activeTab === 'home') {
      refreshHomePageStats();
    } else {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  }

  async function refreshHomePageStats() {
    if (homeStatBakiyeCount) {
      homeStatBakiyeCount.textContent = String((bakiyeState && bakiyeState.allItems ? bakiyeState.allItems.length : 0));
    }
    try {
      const resp = await fetch('/api/bakiye-siparis/kpi-records?limit=200');
      const data = await resp.json();
      const rows = data.kayitlar || [];
      kpiLoadedRecordsCache = rows;
      if (homeStatKpiCount) homeStatKpiCount.textContent = String(rows.length);

      let sumA = 0, sumCommon = 0, sumB = 0;
      rows.forEach(r => {
        sumA += Number(r.ilk_faaliyet_benzersiz_sayisi || 0);
        sumCommon += Number(r.ortak_uygunsuzluk_sayisi || 0);
        sumB += Number(r.ikinci_faaliyet_benzersiz_sayisi || 0);
      });
      if (homeStatOnlyA) homeStatOnlyA.textContent = String(sumA);
      if (homeStatCommon) homeStatCommon.textContent = String(sumCommon);
      if (homeStatOnlyB) homeStatOnlyB.textContent = String(sumB);
    } catch (_) {}
  }

  function isCurrentUserAdmin() {
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('admin') === '1' || (urlParams.get('role') || '').toLowerCase() === 'admin') {
      return true;
    }
    const u = kokpitIdentity && kokpitIdentity.user;
    if (!u || typeof u !== 'object') return false;
    if (u.is_admin === true || u.isAdmin === true) return true;

    const adminRoles = ['admin', 'yönetici', 'yonetici', 'superuser', 'system_admin', 'sistem yöneticisi'];
    const r = String(u.role || u.rol || '').trim().toLowerCase();
    if (r && adminRoles.includes(r)) return true;

    if (Array.isArray(u.roles)) {
      for (const item of u.roles) {
        if (adminRoles.includes(String(item || '').trim().toLowerCase())) {
          return true;
        }
      }
    }
    return false;
  }

  function applyAdminVisibility() {
    const isAdmin = isCurrentUserAdmin();
    document.body.classList.toggle('is-admin', isAdmin);

    if (btnOpenSettings) {
      btnOpenSettings.style.display = isAdmin ? 'inline-flex' : 'none';
    }
    if (btnSwaggerDocs) {
      btnSwaggerDocs.style.display = isAdmin ? 'inline-flex' : 'none';
    }
    if (configPillsGroup) {
      configPillsGroup.style.display = isAdmin ? 'flex' : 'none';
    }
    if (!isAdmin && settingsModal && settingsModal.style.display !== 'none') {
      settingsModal.style.display = 'none';
    }
    if (currentAnalysis) {
      renderSummarySection(currentAnalysis);
      renderTableRows();
    }
    if (kpiRecordsModal && kpiRecordsModal.style.display !== 'none') {
      renderFilteredKpiRecordsTable();
    }
  }

  function initTheme() {
    const savedTheme = localStorage.getItem('vennflow_theme') || 'light';
    applyTheme(savedTheme);
  }

  function applyTheme(theme) {
    const portalThemeIcon = document.getElementById('portalThemeIcon');
    if (theme === 'dark') {
      document.body.classList.remove('light-theme');
      document.body.classList.add('dark-theme');
      if (themeToggleIcon) themeToggleIcon.textContent = '☀️';
      if (themeToggleText) themeToggleText.textContent = 'Açık Tema';
      if (portalThemeIcon) portalThemeIcon.textContent = '☀️';
    } else {
      document.body.classList.remove('dark-theme');
      document.body.classList.add('light-theme');
      if (themeToggleIcon) themeToggleIcon.textContent = '🌙';
      if (themeToggleText) themeToggleText.textContent = 'Koyu Tema';
      if (portalThemeIcon) portalThemeIcon.textContent = '🌙';
    }
    localStorage.setItem('vennflow_theme', theme);

    if (currentAnalysis && currentAnalysis.venn_verileri) {
      drawVennDiagram(currentAnalysis.venn_verileri, currentAnalysis.dosya_1_adi, currentAnalysis.dosya_2_adi);
    }
  }

  function toggleTheme() {
    const isDark = document.body.classList.contains('dark-theme');
    applyTheme(isDark ? 'light' : 'dark');
  }

  function setupEventListeners() {
    // Theme toggle
    if (btnThemeToggle) {
      btnThemeToggle.addEventListener('click', toggleTheme);
    }

    // Engine mode toggle
    btnEngineLocal.addEventListener('click', () => setEngineMode('local'));
    if (btnEngineEmbedding) {
      btnEngineEmbedding.addEventListener('click', () => setEngineMode('embedding'));
    }
    btnEngineAselix.addEventListener('click', () => setEngineMode('aselix'));

    // Threshold slider
    thresholdSlider.addEventListener('input', (e) => {
      const val = parseInt(e.target.value, 10);
      thresholdValueBadge.textContent = `%${val}`;
      appConfig.threshold = val;
      saveConfig();
    });

    // Input mode tabs
    tabBtnExcel.addEventListener('click', () => setInputMode('excel'));
    tabBtnPaste.addEventListener('click', () => setInputMode('paste'));

    // Drag and Drop
    setupDragAndDrop(dropzone1, fileInput1, (f) => handleExcelUpload(1, f));
    setupDragAndDrop(dropzone2, fileInput2, (f) => handleExcelUpload(2, f));

    btnRemoveFile1.addEventListener('click', (e) => { e.stopPropagation(); removeExcelFile(1); });
    btnRemoveFile2.addEventListener('click', (e) => { e.stopPropagation(); removeExcelFile(2); });

    // Column change in select
    colSelect1.addEventListener('change', () => updateActiveColumnHighlight(1));
    colSelect2.addEventListener('change', () => updateActiveColumnHighlight(2));

    // Preview table toggles
    btnTogglePreview1.addEventListener('click', () => togglePreviewCard(1));
    btnTogglePreview2.addEventListener('click', () => togglePreviewCard(2));

    // Live row counter for paste
    textPaste1.addEventListener('input', updatePasteCounters);
    textPaste2.addEventListener('input', updatePasteCounters);

    // Run Comparison
    btnRunComparison.addEventListener('click', runUnifiedComparison);
    btnUseSampleFiles.addEventListener('click', loadSampleDemoData);
    if (btnCloseProgress) btnCloseProgress.addEventListener('click', () => { progressBanner.style.display = 'none'; });

    // Filter Buttons
    tabFilterAll.addEventListener('click', () => setTableFilter('all'));
    tabFilterIntersection.addEventListener('click', () => setTableFilter('intersection'));
    tabFilterA.addEventListener('click', () => setTableFilter('only_a'));
    tabFilterB.addEventListener('click', () => setTableFilter('only_b'));
    if (tabFilterDeclined) {
      tabFilterDeclined.addEventListener('click', () => setTableFilter('declined'));
    }

    // Decline / Restore / Admin Delete Similarity click delegation on table body
    if (detailsTableBody) {
      detailsTableBody.addEventListener('click', (e) => {
        const declineBtn = e.target.closest('.btn-decline-match');
        if (declineBtn) {
          const pairIdx = parseInt(declineBtn.dataset.pairIdx, 10);
          if (!isNaN(pairIdx)) declineMatchedPair(pairIdx);
          return;
        }
        const restoreBtn = e.target.closest('.btn-restore-match');
        if (restoreBtn) {
          const decIdx = parseInt(restoreBtn.dataset.declinedIdx, 10);
          if (!isNaN(decIdx)) restoreDeclinedPair(decIdx);
          return;
        }
        const delDetailBtn = e.target.closest('.btn-delete-detail-row');
        if (delDetailBtn) {
          deleteDetailRow(
            delDetailBtn.dataset.region,
            delDetailBtn.dataset.idx !== undefined ? parseInt(delDetailBtn.dataset.idx, 10) : -1,
            delDetailBtn.dataset.itemText || ''
          );
        }
      });
    }

    // Detail table column filters
    if (colFilterDoc1) colFilterDoc1.addEventListener('input', () => renderTableRows());
    if (colFilterDoc2) colFilterDoc2.addEventListener('input', () => renderTableRows());
    if (colFilterStatus) colFilterStatus.addEventListener('change', () => renderTableRows());
    if (btnClearDetailColFilters) {
      btnClearDetailColFilters.addEventListener('click', () => {
        if (colFilterDoc1) colFilterDoc1.value = '';
        if (colFilterDoc2) colFilterDoc2.value = '';
        if (colFilterStatus) colFilterStatus.value = 'all';
        if (tableSearchInput) tableSearchInput.value = '';
        renderTableRows();
      });
    }

    // Main Page Tabs & Standalone Fighter-Select Portal Navigation
    if (pageTabBtnHome) pageTabBtnHome.addEventListener('click', () => switchMainPageTab('home'));
    if (pageTabBtnVennflow) pageTabBtnVennflow.addEventListener('click', () => switchMainPageTab('vennflow'));
    if (pageTabBtnBulk) pageTabBtnBulk.addEventListener('click', () => switchMainPageTab('bulk'));
    if (btnHeroGoVennflow) btnHeroGoVennflow.addEventListener('click', (e) => { e.stopPropagation(); switchMainPageTab('vennflow'); });
    if (btnHeroGoBulk) btnHeroGoBulk.addEventListener('click', (e) => { e.stopPropagation(); switchMainPageTab('bulk'); });
    if (btnCardGoVennflow) btnCardGoVennflow.addEventListener('click', (e) => { e.stopPropagation(); switchMainPageTab('vennflow'); });
    if (btnCardGoBulk) btnCardGoBulk.addEventListener('click', (e) => { e.stopPropagation(); switchMainPageTab('bulk'); });

    // Header brand click -> return to standalone Home Portal
    const headerBrandEl = document.querySelector('.header-brand');
    if (headerBrandEl) {
      headerBrandEl.style.cursor = 'pointer';
      headerBrandEl.title = 'Ana Sayfaya (Modül Seçim Ekranına) Dön';
      headerBrandEl.addEventListener('click', () => switchMainPageTab('home'));
    }

    // Standalone Portal Floating HUD Controls & Fighter Blades
    const btnPortalIdentity = document.getElementById('btnPortalIdentity');
    const btnPortalThemeToggle = document.getElementById('btnPortalThemeToggle');
    const btnPortalSettings = document.getElementById('btnPortalSettings');
    const btnPortalQuickDownloadTpl = document.getElementById('btnPortalQuickDownloadTpl');
    const cardModuleVennflow = document.getElementById('cardModuleVennflow');
    const cardModuleBulk = document.getElementById('cardModuleBulk');
    const cardModuleKpiPopup = document.getElementById('cardModuleKpiPopup');

    if (btnPortalIdentity) btnPortalIdentity.addEventListener('click', openIdentityModal);
    if (btnPortalThemeToggle) btnPortalThemeToggle.addEventListener('click', toggleTheme);
    if (btnPortalSettings) btnPortalSettings.addEventListener('click', openSettingsModal);
    if (btnPortalQuickDownloadTpl) {
      btnPortalQuickDownloadTpl.addEventListener('click', (e) => {
        e.stopPropagation();
        downloadBulkTemplateExcel();
      });
    }
    if (cardModuleVennflow) {
      cardModuleVennflow.addEventListener('click', () => switchMainPageTab('vennflow'));
    }
    if (cardModuleBulk) {
      cardModuleBulk.addEventListener('click', () => switchMainPageTab('bulk'));
    }
    if (cardModuleKpiPopup) {
      cardModuleKpiPopup.addEventListener('click', (e) => {
        e.stopPropagation();
        openKpiRecordsModal();
      });
    }

    // KPI Database Save, Pop-Up Modal, Filtering & Admin Delete
    [btnHeaderKpiTable, btnNavOpenKpiPopup, btnHeroOpenKpiPopup, btnCardOpenKpiPopup, btnToggleKpiRecords].forEach(btn => {
      if (btn) btn.addEventListener('click', (e) => { e.stopPropagation(); openKpiRecordsModal(); });
    });
    if (btnCloseKpiModal) btnCloseKpiModal.addEventListener('click', closeKpiRecordsModal);
    if (kpiRecordsModal) {
      kpiRecordsModal.addEventListener('click', (e) => {
        if (e.target === kpiRecordsModal) closeKpiRecordsModal();
      });
    }
    if (btnCloseKpiJsonPreview) {
      btnCloseKpiJsonPreview.addEventListener('click', () => {
        if (kpiModalJsonPreviewBox) kpiModalJsonPreviewBox.style.display = 'none';
      });
    }
    if (btnSaveToKpiDb) {
      btnSaveToKpiDb.addEventListener('click', saveCurrentDataToKpiDatabase);
    }
    if (btnRefreshKpiRecords) {
      btnRefreshKpiRecords.addEventListener('click', loadAndRenderKpiRecords);
    }
    if (btnClearKpiFilters) {
      btnClearKpiFilters.addEventListener('click', clearKpiTableFilters);
    }
    if (kpiFilterProje) kpiFilterProje.addEventListener('change', renderFilteredKpiRecordsTable);
    if (kpiFilterFaaliyet2) kpiFilterFaaliyet2.addEventListener('change', renderFilteredKpiRecordsTable);
    if (kpiFilterSd) kpiFilterSd.addEventListener('input', renderFilteredKpiRecordsTable);
    if (kpiFilterKalem) kpiFilterKalem.addEventListener('input', renderFilteredKpiRecordsTable);
    if (kpiFilterUrun) kpiFilterUrun.addEventListener('input', renderFilteredKpiRecordsTable);
    if (kpiFilterSearch) kpiFilterSearch.addEventListener('input', renderFilteredKpiRecordsTable);
    if (kpiSavedRecordsBody) {
      kpiSavedRecordsBody.addEventListener('click', (e) => {
        const delBtn = e.target.closest('.btn-delete-kpi-record');
        if (delBtn) {
          const recIdx = parseInt(delBtn.dataset.recIdx, 10);
          if (!isNaN(recIdx)) deleteKpiRecordByIndex(recIdx);
          return;
        }
        const jsonBtn = e.target.closest('.btn-view-json');
        if (jsonBtn && jsonBtn.dataset.encodedJson) {
          showKpiJsonPreviewInModal(
            jsonBtn.dataset.recTitle || 'Kayıt JSON Uygunsuzluk Detayları',
            decodeURIComponent(jsonBtn.dataset.encodedJson)
          );
        }
      });
    }

    // Bulk Data Entry (Toplu Veri Girişi) Listeners
    if (btnDownloadBulkTemplate) {
      btnDownloadBulkTemplate.addEventListener('click', downloadBulkTemplateExcel);
    }
    if (bulkDropzone && bulkFileInput) {
      setupDragAndDrop(bulkDropzone, bulkFileInput, handleBulkFileSelected);
    }
    if (btnRemoveBulkFile) {
      btnRemoveBulkFile.addEventListener('click', (e) => {
        e.stopPropagation();
        removeBulkFile();
      });
    }
    if (btnRunBulkAnalysis) {
      btnRunBulkAnalysis.addEventListener('click', runBulkExcelAnalysis);
    }
    if (btnSaveAllBulkGroups) {
      btnSaveAllBulkGroups.addEventListener('click', saveAllBulkGroupsToDb);
    }
    if (bulkGroupsContainer) {
      bulkGroupsContainer.addEventListener('click', handleBulkGroupsContainerClick);
    }

    // Metric Cards click filters
    document.getElementById('cardIntersection')?.addEventListener('click', () => setTableFilter('intersection'));
    document.getElementById('cardOnlyA')?.addEventListener('click', () => setTableFilter('only_a'));
    document.getElementById('cardOnlyB')?.addEventListener('click', () => setTableFilter('only_b'));

    // Search and Export
    tableSearchInput.addEventListener('input', () => renderTableRows());
    btnExportCSV.addEventListener('click', exportTableToCSV);
    btnExportJSON.addEventListener('click', exportTableToJSON);
    btnCopySummary.addEventListener('click', copySummaryToClipboard);

    // Chat
    btnSendChat.addEventListener('click', handleSendChatMessage);
    chatInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        handleSendChatMessage();
      }
    });
    btnClearChat.addEventListener('click', () => {
      chatMessages.innerHTML = `
        <div class="chat-msg bot-msg">
          <div class="msg-avatar">🤖</div>
          <div class="msg-bubble"><p>Sohbet temizlendi. Kıyaslama sonuçlarınızla ilgili yeni bir soru sorabilirsiniz.</p></div>
        </div>
      `;
    });

    // Suggestion chips
    document.querySelectorAll('.chip-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        chatInput.value = btn.dataset.query;
        handleSendChatMessage();
      });
    });

    // Settings Modal
    btnOpenSettings.addEventListener('click', openSettingsModal);
    btnCloseSettings.addEventListener('click', closeSettingsModal);
    settingsModal.addEventListener('click', (e) => {
      if (e.target === settingsModal) closeSettingsModal();
    });
    btnSaveSettings.addEventListener('click', saveSettingsFromModal);
    btnResetSettings.addEventListener('click', resetSettingsToDefault);
    btnToggleApiKey.addEventListener('click', toggleApiKeyVisibility);
    btnFetchModels.addEventListener('click', fetchModelsFromApi);
    btnTestConn.addEventListener('click', testApiConnection);

    // Kokpit Identity Modal
    if (btnOpenIdentity) btnOpenIdentity.addEventListener('click', openIdentityModal);
    if (btnCloseIdentity) btnCloseIdentity.addEventListener('click', closeIdentityModal);
    if (identityModal) {
      identityModal.addEventListener('click', (e) => {
        if (e.target === identityModal) closeIdentityModal();
      });
    }
    if (btnFillStandardUserPayload) btnFillStandardUserPayload.addEventListener('click', fillStandardUserIdentityPayload);
    if (btnFillDefaultIdentityPayload) btnFillDefaultIdentityPayload.addEventListener('click', fillDefaultIdentityPayload);
    if (btnSimulatePostMessage) btnSimulatePostMessage.addEventListener('click', simulateKokpitPostMessage);
    if (btnLogoutIdentity) btnLogoutIdentity.addEventListener('click', clearKokpitIdentitySession);
    if (btnSaveIdentityConfig) btnSaveIdentityConfig.addEventListener('click', saveIdentityModalConfig);

    // Network radio change
    document.querySelectorAll('input[name="networkChoice"]').forEach(radio => {
      radio.addEventListener('change', (e) => {
        const choice = e.target.value;
        if (choice === 'milli') {
          cfgBaseUrl.value = NETWORK_URLS.milli;
        } else if (choice === 'aynet') {
          cfgBaseUrl.value = NETWORK_URLS.aynet;
        }
      });
    });

    // UGESUDKOKPİT Bakiye Sipariş (teslimat_md.bakiye_siparis) & Faz Seçici
    if (bakiyeProjeSelect1) {
      bakiyeProjeSelect1.addEventListener('change', () => populateBakiyeItemSelect(1, true));
    }
    if (bakiyeProjeSelect2) {
      bakiyeProjeSelect2.addEventListener('change', () => populateBakiyeItemSelect(2, true));
    }
    if (bakiyeSearchInput1) {
      bakiyeSearchInput1.addEventListener('input', () => populateBakiyeItemSelect(1, true));
    }
    if (bakiyeSearchInput2) {
      bakiyeSearchInput2.addEventListener('input', () => populateBakiyeItemSelect(2, true));
    }
    if (bakiyeItemSelect1) {
      bakiyeItemSelect1.addEventListener('change', () => onBakiyeItemChanged(1));
    }
    if (bakiyeItemSelect2) {
      bakiyeItemSelect2.addEventListener('change', () => onBakiyeItemChanged(2));
    }
    if (btnCopyBakiye1To2) {
      btnCopyBakiye1To2.addEventListener('click', copyBakiyeSelection1To2);
    }
    if (phaseGroup1) {
      phaseGroup1.querySelectorAll('.phase-btn').forEach(btn => {
        btn.addEventListener('click', () => setBakiyePhase(1, btn.dataset.phase));
      });
    }
    if (phaseGroup2) {
      phaseGroup2.querySelectorAll('.phase-btn').forEach(btn => {
        btn.addEventListener('click', () => setBakiyePhase(2, btn.dataset.phase));
      });
    }
  }

  // --------------------------------------------------------------------------
  // Config & State Management
  // --------------------------------------------------------------------------
  function loadConfig() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored) {
        return { ...DEFAULT_CONFIG, ...JSON.parse(stored) };
      }
    } catch (e) {
      console.warn('LocalStorage error:', e);
    }
    return { ...DEFAULT_CONFIG };
  }

  function saveConfig() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(appConfig));
    } catch (e) {
      console.warn('LocalStorage save error:', e);
    }
  }

  function applyConfigToUI() {
    // Engine
    setEngineMode(appConfig.engine, false);

    // Threshold
    thresholdSlider.value = appConfig.threshold || 80;
    thresholdValueBadge.textContent = `%${appConfig.threshold || 80}`;

    // Model name
    currentModelDisplay.textContent = appConfig.engine === 'embedding'
      ? (appConfig.embeddingModel || 'text-embedding-3-small')
      : (appConfig.model || 'aselix-llm-v1');
    chatModelNameDisplay.textContent = appConfig.model || 'aselix-llm-v1';

    // Settings modal inputs
    cfgBaseUrl.value = appConfig.baseUrl || NETWORK_URLS.milli;
    cfgApiKey.value = appConfig.apiKey || '';
    if (cfgBatchSize) cfgBatchSize.value = String(appConfig.batchSize || 15);
    if (cfgEmbeddingModel) cfgEmbeddingModel.value = appConfig.embeddingModel || 'text-embedding-3-small';
    if (cfgAllowedOrigins) cfgAllowedOrigins.value = appConfig.allowedOrigins || DEFAULT_ALLOWED_ORIGINS.join(', ');

    if (appConfig.network === 'aynet') {
      radioAynet.checked = true;
    } else if (appConfig.network === 'custom') {
      radioCustom.checked = true;
    } else {
      radioMilli.checked = true;
    }

    // Input mode
    setInputMode(appConfig.inputMode || 'excel');
  }

  function setEngineMode(mode, save = true) {
    appConfig.engine = mode;
    btnEngineLocal.classList.remove('active');
    if (btnEngineEmbedding) btnEngineEmbedding.classList.remove('active');
    btnEngineAselix.classList.remove('active');

    if (mode === 'local') {
      btnEngineLocal.classList.add('active');
      engineStatusLabel.textContent = '⚡ Yerel ML (Çevrimdışı)';
      engineStatusPill.className = 'flow-status-pill';
      btnRunText.textContent = 'Kıyasla & Venn Şemasını Çiz (Yerel)';
      currentModelDisplay.textContent = appConfig.model || 'aselix-llm-v1';
    } else if (mode === 'embedding') {
      if (btnEngineEmbedding) btnEngineEmbedding.classList.add('active');
      engineStatusLabel.textContent = '🧬 Vektör Embedding Aktif';
      engineStatusPill.className = 'flow-status-pill online';
      btnRunText.textContent = 'Embedding ile Kıyasla & Venn Çiz';
      currentModelDisplay.textContent = appConfig.embeddingModel || 'text-embedding-3-small';
    } else {
      btnEngineAselix.classList.add('active');
      engineStatusLabel.textContent = '🤖 AselixAI API Aktif';
      engineStatusPill.className = 'flow-status-pill online';
      btnRunText.textContent = 'Kıyasla, Analiz Et & Venn Çiz';
      currentModelDisplay.textContent = appConfig.model || 'aselix-llm-v1';
    }
    if (save) saveConfig();
  }

  function setInputMode(mode) {
    appConfig.inputMode = mode;
    if (mode === 'excel') {
      tabBtnExcel.classList.add('active');
      tabBtnPaste.classList.remove('active');
      tabContentExcel.style.display = 'block';
      tabContentPaste.style.display = 'none';
      actionHintText.textContent = 'Excel dosyalarını seçin ve kıyaslanacak sütunları belirleyin.';
    } else {
      tabBtnPaste.classList.add('active');
      tabBtnExcel.classList.remove('active');
      tabContentExcel.style.display = 'none';
      tabContentPaste.style.display = 'block';
      actionHintText.textContent = 'Her iki alana da satır satır veri yapıştırarak kıyaslamayı başlatın.';
    }
    saveConfig();
  }

  // --------------------------------------------------------------------------
  // Drag and Drop & Excel Preview Upload
  // --------------------------------------------------------------------------
  function setupDragAndDrop(dropzone, fileInput, onFileSelected) {
    ['dragenter', 'dragover'].forEach(eventName => {
      dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.add('drag-over');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.remove('drag-over');
      });
    });

    dropzone.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      if (dt.files && dt.files.length > 0) {
        onFileSelected(dt.files[0]);
      }
    });

    dropzone.addEventListener('click', (e) => {
      if (!e.target.closest('.btn-remove')) {
        fileInput.click();
      }
    });

    fileInput.addEventListener('change', () => {
      if (fileInput.files && fileInput.files.length > 0) {
        onFileSelected(fileInput.files[0]);
      }
    });
  }

  async function handleExcelUpload(index, file) {
    if (!file.name.match(/\.(xlsx|xls|xlsm)$/i)) {
      alert('Lütfen geçerli bir Excel (.xlsx, .xls) dosyası seçin.');
      return;
    }

    if (index === 1) {
      file1 = file;
      fileName1.textContent = file.name;
      fileSub1.textContent = `${formatBytes(file.size)} • Yükleniyor...`;
      dropzoneContent1.style.display = 'none';
      fileView1.style.display = 'flex';
    } else {
      file2 = file;
      fileName2.textContent = file.name;
      fileSub2.textContent = `${formatBytes(file.size)} • Yükleniyor...`;
      dropzoneContent2.style.display = 'none';
      fileView2.style.display = 'flex';
    }

    // Call /api/compare/excel-preview
    try {
      const formData = new FormData();
      formData.append('file', file);

      const resp = await fetch('/api/compare/excel-preview', {
        method: 'POST',
        body: formData
      });
      const data = await resp.json();

      if (data.durum === 'basarili') {
        if (data.tespit_edilen_bakiye) {
          applyAutoDetectedBakiyeFromExcel(index, data.tespit_edilen_bakiye, data.tespit_edilen_faz);
        }
        if (index === 1) {
          previewData1 = data;
          fileSub1.textContent = `${formatBytes(file.size)} • ${data.toplam_satir} satır okundu`;
          populateColumnSelector(1, data.sutunlar, data.onerilen_sutun);
          renderPreviewMiniTable(1, data.sutunlar, data.onerilen_sutun, data.onizleme_satirlari, data.toplam_satir);
          colPickBox1.style.display = 'flex';
        } else {
          previewData2 = data;
          fileSub2.textContent = `${formatBytes(file.size)} • ${data.toplam_satir} satır okundu`;
          populateColumnSelector(2, data.sutunlar, data.onerilen_sutun);
          renderPreviewMiniTable(2, data.sutunlar, data.onerilen_sutun, data.onizleme_satirlari, data.toplam_satir);
          colPickBox2.style.display = 'flex';
        }
      } else {
        alert(`Dosya önizleme uyarısı: ${data.mesaj}`);
      }
    } catch (err) {
      console.error('Preview error:', err);
      if (index === 1) fileSub1.textContent = `${formatBytes(file.size)} • Hazır`;
      else fileSub2.textContent = `${formatBytes(file.size)} • Hazır`;
    }
  }

  function removeExcelFile(index) {
    if (index === 1) {
      file1 = null;
      previewData1 = null;
      fileInput1.value = '';
      dropzoneContent1.style.display = 'flex';
      fileView1.style.display = 'none';
      colPickBox1.style.display = 'none';
      previewCard1.style.display = 'none';
    } else {
      file2 = null;
      previewData2 = null;
      fileInput2.value = '';
      dropzoneContent2.style.display = 'flex';
      fileView2.style.display = 'none';
      colPickBox2.style.display = 'none';
      previewCard2.style.display = 'none';
    }
  }

  function populateColumnSelector(index, columns, suggestedCol) {
    const sel = index === 1 ? colSelect1 : colSelect2;
    sel.innerHTML = '';

    columns.forEach(col => {
      const opt = document.createElement('option');
      opt.value = col;
      opt.textContent = col + (col === suggestedCol ? ' ★ (Algılanan Uygunsuzluk Sütunu)' : '');
      if (col === suggestedCol) {
        opt.selected = true;
      }
      sel.appendChild(opt);
    });
  }

  function renderPreviewMiniTable(index, columns, activeCol, rows, totalRows) {
    const scrollEl = index === 1 ? previewTableScroll1 : previewTableScroll2;
    const badgeEl = index === 1 ? previewRowCountBadge1 : previewRowCountBadge2;

    badgeEl.textContent = `Toplam ${totalRows} Satır (İlk 5 Gösteriliyor)`;

    if (!rows || rows.length === 0) {
      scrollEl.innerHTML = '<p style="padding:10px; color:var(--text-muted);">Önizleme verisi yok.</p>';
      return;
    }

    let thHtml = columns.map(c => `
      <th class="${c === activeCol ? 'active-col' : ''}">${escapeHtml(c)}</th>
    `).join('');

    let trHtml = rows.map(r => `
      <tr>
        ${columns.map(c => `<td class="${c === activeCol ? 'active-col' : ''}">${escapeHtml(String(r[c] !== null && r[c] !== undefined ? r[c] : ''))}</td>`).join('')}
      </tr>
    `).join('');

    scrollEl.innerHTML = `
      <table class="preview-mini-table">
        <thead><tr>${thHtml}</tr></thead>
        <tbody>${trHtml}</tbody>
      </table>
    `;
  }

  function updateActiveColumnHighlight(index) {
    const sel = index === 1 ? colSelect1 : colSelect2;
    const previewData = index === 1 ? previewData1 : previewData2;
    if (previewData && previewData.sutunlar) {
      renderPreviewMiniTable(index, previewData.sutunlar, sel.value, previewData.onizleme_satirlari, previewData.toplam_satir);
    }
  }

  function togglePreviewCard(index) {
    const card = index === 1 ? previewCard1 : previewCard2;
    const btn = index === 1 ? btnTogglePreview1 : btnTogglePreview2;
    if (card.style.display === 'none') {
      card.style.display = 'block';
      btn.textContent = '❌ Önizlemeyi Gizle';
    } else {
      card.style.display = 'none';
      btn.textContent = '👁️ İlk 5 Satırı Göster';
    }
  }

  // --------------------------------------------------------------------------
  // Copy-Paste Textarea Counter
  // --------------------------------------------------------------------------
  function updatePasteCounters() {
    const lines1 = textPaste1.value.split('\n').filter(l => l.trim().length > 0);
    const lines2 = textPaste2.value.split('\n').filter(l => l.trim().length > 0);
    lineCount1.textContent = `${lines1.length} satır`;
    lineCount2.textContent = `${lines2.length} satır`;
  }

  // --------------------------------------------------------------------------
  // Demo Sample Data Loader
  // --------------------------------------------------------------------------
  function loadSampleDemoData() {
    if (appConfig.inputMode === 'paste') {
      textPaste1.value = [
        'Pano içi etiketleme eksik',
        'Kablo kesiti standart dışı (2.5mm yerine 1.5mm)',
        'Topraklama barası cıvatası gevşek',
        'Acil durdurma butonu mekanik takılma yapıyor',
        'PLC haberleşme modülü zaman aşımı hatası',
        'Hidrolik rekor sızdırmazlık contası deforme',
        'Röle soket kilitleme tırnağı kırık'
      ].join('\n');

      textPaste2.value = [
        'etiketleme eksik pano içi', // Fuzzy match with typo/reordered words
        'Topraklama barası cıvatası gevşek', // Exact match
        'Acil stop butonu mekanik takılıyor', // Fuzzy match
        'Optik sensör kalibrasyon sapması', // New error in 2
        'Solenoid valf bobini aşırı ısınıyor', // New error in 2
        'Ethernet soketi RJ45 tırnağı oturmamış' // New error in 2
      ].join('\n');

      updatePasteCounters();
      runUnifiedComparison();
    } else {
      // Load sample excel through demo endpoint
      triggerSampleExcelCompare();
    }
  }

  async function triggerSampleExcelCompare() {
    progressBanner.style.display = 'flex';
    progressStatusText.textContent = 'Sunucudaki hazır örnek uygunsuzluk Excel verileri yükleniyor...';
    progressSubText.textContent = 'Dry-run ve SKT fazları karşılaştırma verisi hazırlanıyor...';

    try {
      const resp = await fetch('/api/flow/sample-compare', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          base_url: appConfig.baseUrl,
          flow_id: 'c16ecc83-ff64-4fc7-bb52-0f594beaf70a',
          api_key: appConfig.apiKey
        })
      });

      const data = await resp.json();
      progressBanner.style.display = 'none';

      if (data.durum === 'basarili') {
        fileName1.textContent = data.dosya_1_adi || 'Ornek_Faz1_DryRun.xlsx';
        fileSub1.textContent = 'Hazır Test Fazı';
        dropzoneContent1.style.display = 'none';
        fileView1.style.display = 'flex';

        fileName2.textContent = data.dosya_2_adi || 'Ornek_Faz2_SKT.xlsx';
        fileSub2.textContent = 'Hazır Test Fazı';
        dropzoneContent2.style.display = 'none';
        fileView2.style.display = 'flex';

        renderAnalysisResults(data);
      } else {
        alert('Örnek veri yükleme uyarısı: ' + data.mesaj);
      }
    } catch (e) {
      progressBanner.style.display = 'none';
      alert('Bağlantı hatası: ' + e.message);
    }
  }

  // --------------------------------------------------------------------------
  // AselixAI Batch Pipeline (her istek < 1 dk, canlı log ile)
  // --------------------------------------------------------------------------
  const REQUEST_HARD_LIMIT_MS = 58000; // Tarayıcı tarafı güvenlik: 58 sn'de isteği iptal et

  function setProgress(pct) {
    const p = Math.max(0, Math.min(100, Math.round(pct)));
    if (progressBarFill) progressBarFill.style.width = `${p}%`;
    if (progressPctBadge) progressPctBadge.textContent = `%${p}`;
  }

  function batchLog(msg, level = 'info') {
    if (!batchLogPanel) return;
    batchLogPanel.style.display = 'block';
    const line = document.createElement('div');
    line.className = `batch-log-line log-${level}`;
    const ts = new Date().toLocaleTimeString('tr-TR');
    line.innerHTML = `<span class="log-ts">${ts}</span> ${escapeHtml(msg)}`;
    batchLogPanel.appendChild(line);
    batchLogPanel.scrollTop = batchLogPanel.scrollHeight;
  }

  function resetBatchLog() {
    if (!batchLogPanel) return;
    batchLogPanel.innerHTML = '';
    batchLogPanel.style.display = 'none';
    if (btnCloseProgress) btnCloseProgress.style.display = 'none';
    progressBanner.classList.remove('progress-error');
  }

  function failPipeline(msg) {
    batchLog(msg, 'error');
    progressStatusText.textContent = '❌ İşlem durdu: ' + msg;
    progressBanner.classList.add('progress-error');
    if (btnCloseProgress) btnCloseProgress.style.display = 'inline-flex';
    btnRunComparison.disabled = false;
  }

  // Zaman sınırlı fetch + canlı geçen süre sayacı
  async function fetchWithLimit(url, options, liveLabel) {
    const controller = new AbortController();
    const t0 = Date.now();
    const timer = setTimeout(() => controller.abort(), REQUEST_HARD_LIMIT_MS);
    const ticker = setInterval(() => {
      const sec = Math.round((Date.now() - t0) / 1000);
      progressSubText.textContent = `${liveLabel} • Geçen süre: ${sec} sn / limit 50 sn`;
    }, 1000);
    try {
      const resp = await fetch(url, { ...options, signal: controller.signal });
      const text = await resp.text();
      let data;
      try { data = JSON.parse(text); } catch (_) {
        throw new Error(`Sunucu JSON dışı yanıt döndü (HTTP ${resp.status}): ${text.slice(0, 200)}`);
      }
      if (!resp.ok) {
        const detail = data && (data.mesaj || data.detail);
        throw new Error(`HTTP ${resp.status}: ${typeof detail === 'string' ? detail : JSON.stringify(detail).slice(0, 300)}`);
      }
      return { data, elapsed: (Date.now() - t0) / 1000 };
    } catch (err) {
      if (err.name === 'AbortError') {
        throw new Error(`İstek ${REQUEST_HARD_LIMIT_MS / 1000} sn içinde tamamlanmadı ve iptal edildi.`);
      }
      throw err;
    } finally {
      clearTimeout(timer);
      clearInterval(ticker);
    }
  }

  async function runAselixBatchPipeline(inputMode, thresholdDec, batchSize) {
    resetBatchLog();
    setProgress(2);
    const model = appConfig.model || 'aselix-llm-v1';

    // 1) Verileri hazırla
    let itemsA = [], itemsB = [], doc1 = '1. Metin Seti', doc2 = '2. Metin Seti', col1 = 'Metin', col2 = 'Metin';
    try {
      if (inputMode === 'excel') {
        progressStatusText.textContent = 'Excel sütunları okunuyor...';
        batchLog(`Excel okunuyor: ${file1.name} [${colSelect1.value || 'auto'}] ve ${file2.name} [${colSelect2.value || 'auto'}]`);
        const fd = new FormData();
        fd.append('file_1', file1);
        fd.append('file_2', file2);
        fd.append('col_1', colSelect1.value || 'auto');
        fd.append('col_2', colSelect2.value || 'auto');
        const { data } = await fetchWithLimit('/api/compare/extract-items', { method: 'POST', body: fd }, 'Excel okunuyor');
        if (data.durum !== 'basarili') throw new Error(data.mesaj || 'Excel okunamadı');
        itemsA = data.items_1; itemsB = data.items_2;
        doc1 = data.dosya_1_adi; doc2 = data.dosya_2_adi; col1 = data.sutun_1; col2 = data.sutun_2;
      } else {
        itemsA = textPaste1.value.split('\n').map(l => l.trim()).filter(Boolean);
        itemsB = textPaste2.value.split('\n').map(l => l.trim()).filter(Boolean);
      }
    } catch (err) {
      failPipeline('Veri hazırlama hatası: ' + err.message);
      return;
    }

    // Tekrarlayan satırları tekilleştir (gereksiz API yükünü önler)
    itemsA = [...new Set(itemsA)];
    itemsB = [...new Set(itemsB)];
    if (!itemsA.length || !itemsB.length) {
      failPipeline('Kıyaslanacak geçerli satır bulunamadı.');
      return;
    }

    const chunks = [];
    for (let i = 0; i < itemsA.length; i += batchSize) chunks.push(itemsA.slice(i, i + batchSize));
    const total = chunks.length;
    batchLog(`Hazır: A=${itemsA.length} tekil satır, B=${itemsB.length} tekil satır → ${total} batch × ${batchSize} satır (model: ${model})`);

    // 2) Batch döngüsü
    const allPairs = [];
    const usedB = new Set();
    let aiOk = 0, fallbackCount = 0, failCount = 0;

    for (let i = 0; i < total; i++) {
      const chunk = chunks[i];
      const label = `Batch ${i + 1}/${total}`;
      setProgress(5 + (i / total) * 80);
      progressStatusText.textContent = `${label} AselixAI'ye gönderildi (${chunk.length} satır)...`;

      try {
        const { data, elapsed } = await fetchWithLimit('/api/compare/batch-chunk', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            base_url: appConfig.baseUrl,
            api_key: appConfig.apiKey,
            model,
            items_a_batch: chunk,
            candidates_b: itemsB.filter(b => !usedB.has(b)),
            threshold: thresholdDec,
            timeout_sec: 50,
            batch_index: i + 1,
            total_batches: total
          })
        }, label);

        const pairs = Array.isArray(data.matched_pairs) ? data.matched_pairs : [];
        allPairs.push(...pairs);
        pairs.forEach(p => usedB.add(p.item_b));
        const d = data.tanilama || {};
        const extra = d.shortlist_sayisi ? `, ${d.shortlist_sayisi} B adayı` : '';

        if (data.kaynak === 'aselix_ai') {
          aiOk++;
          batchLog(`✓ ${label}: ${elapsed.toFixed(1)} sn, ${pairs.length} eşleşme (AI: ${d.llm_eslesme ?? '-'}, fuzzy: ${d.fuzzy_ek_eslesme ?? 0}${extra})`, 'ok');
        } else {
          fallbackCount++;
          batchLog(`⚠ ${label}: ${elapsed.toFixed(1)} sn, yerel ML ile tamamlandı (${pairs.length} eşleşme). Sebep: ${data.api_hatasi || data.kaynak}`, 'warn');
          if (d.llm_yanit_onizleme) batchLog(`   LLM yanıt önizleme: ${d.llm_yanit_onizleme}`, 'warn');
        }
      } catch (err) {
        failCount++;
        batchLog(`✗ ${label} başarısız: ${err.message}`, 'error');
        // İlk 2 batch'in ikisi de tamamen başarısızsa (ağ/sunucu sorunu) devam etmek anlamsız
        if (i === 1 && failCount === 2) {
          failPipeline('İlk iki batch başarısız oldu. Sunucu/ağ bağlantısını ve API ayarlarını kontrol edin.');
          return;
        }
      }
    }

    batchLog(`Batch'ler tamamlandı: ${aiOk} AI, ${fallbackCount} yerel yedek, ${failCount} hatalı. Toplam ${allPairs.length} eşleşme.`, failCount ? 'warn' : 'ok');

    // 3) Birleştir + yönetici özeti (ayrı istek, < 1 dk)
    setProgress(88);
    progressStatusText.textContent = 'Sonuçlar birleştiriliyor ve yönetici özeti üretiliyor...';
    let finalData;
    try {
      const res = await fetchWithLimit('/api/compare/batch-finalize', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          base_url: appConfig.baseUrl,
          api_key: appConfig.apiKey,
          model,
          threshold: thresholdDec,
          all_items_a: itemsA,
          all_items_b: itemsB,
          matched_pairs: allPairs,
          doc1_name: doc1,
          doc2_name: doc2,
          col1_name: col1,
          col2_name: col2,
          need_executive_summary: true,
          meta_1: getBakiyeMetaPayload(1),
          meta_2: getBakiyeMetaPayload(2)
        })
      }, 'Yönetici özeti');
      finalData = res.data;
      batchLog(`✓ Birleştirme + özet: ${res.elapsed.toFixed(1)} sn (özet durumu: ${finalData.ai_durum})`, finalData.ai_durum === 'basarili' ? 'ok' : 'warn');
    } catch (err) {
      failPipeline('Birleştirme/özet hatası: ' + err.message);
      return;
    }

    if (finalData.durum !== 'basarili') {
      failPipeline(finalData.mesaj || 'Birleştirme başarısız.');
      return;
    }

    setProgress(100);
    progressStatusText.textContent = '✓ Kıyaslama tamamlandı';
    btnRunComparison.disabled = false;
    // Uyarı/hata olduysa log görünür kalsın, kullanıcı kapatsın
    if (fallbackCount || failCount) {
      if (btnCloseProgress) btnCloseProgress.style.display = 'inline-flex';
    } else {
      setTimeout(() => { progressBanner.style.display = 'none'; }, 1200);
    }
    renderAnalysisResults(finalData);
  }
  // --------------------------------------------------------------------------
  // Run Unified Comparison (/api/compare/run)
  // --------------------------------------------------------------------------
  async function runUnifiedComparison() {
    const inputMode = appConfig.inputMode || 'excel';

    if (inputMode === 'excel') {
      if (!file1 || !file2) {
        alert('Lütfen kıyaslamak için her iki Excel dosyasını da seçin veya sürükleyin.');
        return;
      }
    } else {
      const lines1 = textPaste1.value.split('\n').filter(l => l.trim().length > 0);
      const lines2 = textPaste2.value.split('\n').filter(l => l.trim().length > 0);
      if (lines1.length === 0 || lines2.length === 0) {
        alert('Lütfen her iki metin alanına da kıyaslanacak en az birer satır yapıştırın.');
        return;
      }
    }

    progressBanner.style.display = 'flex';
    resetBatchLog();
    if (progressBarWrapper) progressBarWrapper.style.display = 'block';
    if (progressBarFill) progressBarFill.style.width = '5%';
    if (progressPctBadge) progressPctBadge.textContent = '%5';
    btnRunComparison.disabled = true;

    const thresholdDec = (appConfig.threshold || 80) / 100.0;
    const isAselix = appConfig.engine === 'aselix';
    const isEmbedding = appConfig.engine === 'embedding';
    const batchSize = parseInt(appConfig.batchSize, 10) || 15;
    const effectiveApiKey = (appConfig.apiKey || '').trim() || (kokpitIdentity.token || '').trim();

    // AselixAI modu: Excel VE metin için istemci tarafında gerçek batch döngüsü.
    // Her istek ayrı bir HTTP çağrısıdır -> hiçbir istek 1 dakikayı aşmaz.
    if (isAselix) {
      if (!effectiveApiKey) {
        progressBanner.style.display = 'none';
        btnRunComparison.disabled = false;
        if (isCurrentUserAdmin()) {
          alert('AselixAI modu için Ayarlar\'dan API anahtarı girmelisiniz veya Kokpit SSO Token ile oturum açmalısınız (veya Yerel ML / Embedding moduna geçin).');
        } else {
          alert('AselixAI bulut analizi için sistem yöneticisi tarafından API bağlantısı tanımlanmalıdır veya Kokpit SSO ile oturum açılmalıdır. Lütfen Yerel ML veya Vektör Embedding modunu kullanın.');
        }
        return;
      }
      await runAselixBatchPipeline(inputMode, thresholdDec, batchSize);
      return;
    }

    // Embedding veya Yerel ML Modu
    progressStatusText.textContent = isEmbedding
      ? 'Veriler Vektör Embedding uzayına dönüştürülüyor ve kosinüs benzerliği hesaplanıyor...'
      : 'Veriler yerel mesafe ve fuzzy matching algoritması ile kıyaslanıyor...';
    progressSubText.textContent = `Benzerlik toleransı: %${appConfig.threshold} • Motor: ${isEmbedding ? 'Vektör Embedding (' + (appConfig.embeddingModel || 'text-embedding-3-small') + ')' : 'Yerel ML'}`;

    let progressSimInterval = null;

    try {
      const formData = new FormData();
      formData.append('input_mode', inputMode);
      formData.append('engine_mode', appConfig.engine);
      formData.append('threshold', thresholdDec);
      formData.append('base_url', appConfig.baseUrl);
      formData.append('api_key', effectiveApiKey);
      formData.append('model', isEmbedding ? (appConfig.embeddingModel || 'text-embedding-3-small') : (appConfig.model || 'aselix-llm-v1'));

      if (inputMode === 'excel') {
        formData.append('file_1', file1);
        formData.append('file_2', file2);
        formData.append('col_1', colSelect1.value || 'auto');
        formData.append('col_2', colSelect2.value || 'auto');
      } else {
        formData.append('text_1', textPaste1.value);
        formData.append('text_2', textPaste2.value);
        formData.append('name_1', '1. Metin Seti');
        formData.append('name_2', '2. Metin Seti');
      }

      formData.append('meta_1', JSON.stringify(getBakiyeMetaPayload(1)));
      formData.append('meta_2', JSON.stringify(getBakiyeMetaPayload(2)));

      const fetchHeaders = {};
      if (kokpitIdentity.token) {
        fetchHeaders['Authorization'] = `Bearer ${kokpitIdentity.token}`;
      }

      const resp = await fetch('/api/compare/run', {
        method: 'POST',
        headers: fetchHeaders,
        body: formData
      });

      if (progressSimInterval) clearInterval(progressSimInterval);
      if (progressBarFill) progressBarFill.style.width = '100%';
      if (progressPctBadge) progressPctBadge.textContent = '%100';

      const data = await resp.json();
      btnRunComparison.disabled = false;
      setTimeout(() => { progressBanner.style.display = 'none'; }, 400);

      if (data.durum === 'basarili') {
        renderAnalysisResults(data);
      } else {
        alert('Kıyaslama Hatası: ' + (data.mesaj || 'İşlem tamamlanamadı.'));
      }
    } catch (err) {
      if (progressSimInterval) clearInterval(progressSimInterval);
      btnRunComparison.disabled = false;
      progressBanner.style.display = 'none';
      alert('Sunucu iletişim hatası: ' + err.message);
    }
  }

  // --------------------------------------------------------------------------
  // Render Results, Venn Diagram & Table
  // --------------------------------------------------------------------------
  function sanitizeAdminTextForUser(rawText) {
    if (!rawText) return '';
    const lines = String(rawText).split('\n');
    const filtered = lines.filter(line => {
      const l = line.toLowerCase();
      if (l.includes('llm suite flow bağlantısı')) return false;
      if (l.includes('hedef akış:') || l.includes('/api/v1/run/')) return false;
      if (l.includes('ayarlar') && (l.includes('api key') || l.includes('x-api-key') || l.includes('anahtar'))) return false;
      if (l.includes('x-api-key')) return false;
      return true;
    });
    return filtered.join('\n').replace(/\n{3,}/g, '\n\n').trim();
  }

  function renderSummarySection(data) {
    if (!data) return;
    const v = data.venn_verileri || {};
    const isAselix = data.engine_mode === 'aselix';
    const isEmbedding = data.engine_mode === 'embedding';
    const isAdmin = isCurrentUserAdmin();

    // Summary Engine Badge
    if (isAselix) {
      summaryEngineBadge.innerHTML = `<span class="engine-icon">🤖</span><span>${isAdmin ? `ASELSAN AselixAI (${data.model_kullanilan || appConfig.model}) Yönetici Özeti` : 'Yapay Zeka Yönetici Özeti'}</span>`;
    } else if (isEmbedding) {
      summaryEngineBadge.innerHTML = `<span class="engine-icon">🧬</span><span>Vektör Embedding (%${Math.round((data.esik_orani || 0.8) * 100)} Eşik)</span>`;
    } else {
      summaryEngineBadge.innerHTML = `<span class="engine-icon">⚡</span><span>Yerel ML & String Distance (%${Math.round((data.esik_orani || 0.8) * 100)} Eşik)</span>`;
    }

    const rawFlowSummary = (data.flow_analiz_ozeti || '').trim();
    const cleanFlowSummary = isAdmin ? rawFlowSummary : sanitizeAdminTextForUser(rawFlowSummary);

    // Eğer standart kullanıcı için geriye yalnızca jenerik başlık kaldıysa ozet_yorum'u da ekle
    if (cleanFlowSummary) {
      if (!isAdmin && data.ozet_yorum && !cleanFlowSummary.includes(data.ozet_yorum)) {
        summaryText.innerHTML = `
          ${markdownToHtml(cleanFlowSummary)}
          <div style="margin-top: 8px; line-height: 1.7;">
            <p><strong>${escapeHtml(data.ozet_yorum)}</strong></p>
          </div>
        `;
      } else {
        summaryText.innerHTML = markdownToHtml(cleanFlowSummary);
      }
    } else {
      const adminHintHtml = isAdmin
        ? `
          <p style="margin-top: 8px; color: var(--text-muted); font-size: 13px;">
            ℹ️ <strong>[Yönetici Bilgisi]</strong> AselixAI ile derinlemesine kök neden analizi ve önleyici faaliyet önerileri almak için yukarıdan 
            <strong>AselixAI API</strong> modunu aktif edebilir ve <strong>⚙️ Ayarlar</strong>'dan API anahtarınızı tanımlayabilirsiniz.
          </p>
        `
        : '';

      summaryText.innerHTML = `
        <div style="line-height: 1.7;">
          <p><strong>${escapeHtml(data.ozet_yorum || '')}</strong></p>
          ${adminHintHtml}
        </div>
      `;
    }
  }

  function renderAnalysisResults(data) {
    currentAnalysis = data;
    resultsSection.style.display = 'block';
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });

    const v = data.venn_verileri;

    // Render Summary Badge & Text (Admin vs Standard User aware)
    renderSummarySection(data);

    // Render Comparison Meta Banner (Proje, SD/Kalem, Ürün Kısa Metni, Faz: Dry-Run/FKT/SKT)
    const m1 = (data.belge_1_meta && Object.keys(data.belge_1_meta).length > 0) ? data.belge_1_meta : getBakiyeMetaPayload(1);
    const m2 = (data.belge_2_meta && Object.keys(data.belge_2_meta).length > 0) ? data.belge_2_meta : getBakiyeMetaPayload(2);
    renderComparisonMetaBanner(m1, m2, data.dosya_1_adi, data.dosya_2_adi);

    const vennLabel1 = formatShortMetaLabel(m1, data.dosya_1_adi || '1. Belge');
    const vennLabel2 = formatShortMetaLabel(m2, data.dosya_2_adi || '2. Belge');

    // Metrics
    statTotal.textContent = v.total_unique;
    statIntersection.textContent = v.intersection_count;
    statJaccard.textContent = `Jaccard Benzerlik Oranı: %${Math.round((v.jaccard_similarity || 0) * 100)}`;
    statOnlyA.textContent = v.only_a_count;
    statOnlyB.textContent = v.only_b_count;

    statOnlyASub.textContent = `${vennLabel1} içinde tekil`;
    statOnlyBSub.textContent = `${vennLabel2} içinde yeni`;

    legendTextA.textContent = truncateStr(vennLabel1, 26);
    legendTextB.textContent = truncateStr(vennLabel2, 26);

    // Update Filter Tab Badges
    countAll.textContent = v.total_unique;
    countIntersection.textContent = v.intersection_count;
    countA.textContent = v.only_a_count;
    countB.textContent = v.only_b_count;
    if (!Array.isArray(v.declined_pairs)) v.declined_pairs = [];
    if (countDeclined) countDeclined.textContent = v.declined_pairs.length;
    if (tabFilterDeclined) {
      tabFilterDeclined.style.display = v.declined_pairs.length > 0 ? 'inline-flex' : 'none';
    }

    // Draw Venn SVG
    drawVennDiagram(v, vennLabel1, vennLabel2);

    // Render Table & DB KPI Preview Strip
    currentFilter = 'all';
    updateFilterTabButtons();
    renderTableRows();
    renderDbKpiPreviewStrip();
    if (kpiSaveStatusBanner) kpiSaveStatusBanner.style.display = 'none';
  }

  // --------------------------------------------------------------------------
  // Draw Interactive SVG Venn Diagram
  // --------------------------------------------------------------------------
  function drawVennDiagram(vennData, nameA, nameB) {
    const labelA = truncateStr(nameA || '1. Belge', 16);
    const labelB = truncateStr(nameB || '2. Belge', 16);
    const isDark = document.body.classList.contains('dark-theme');

    const gradA_start = '#3b82f6';
    const gradA_end = '#1d4ed8';
    const opA_start = isDark ? '0.55' : '0.26';
    const opA_end = isDark ? '0.35' : '0.14';
    const strokeA = isDark ? '#60a5fa' : '#2563eb';

    const gradB_start = '#ec4899';
    const gradB_end = '#8b5cf6';
    const opB_start = isDark ? '0.55' : '0.26';
    const opB_end = isDark ? '0.35' : '0.14';
    const strokeB = isDark ? '#f472b6' : '#db2777';

    const gradC_start = isDark ? '#60a5fa' : '#6366f1';
    const gradC_end = isDark ? '#c084fc' : '#a855f7';
    const opC = isDark ? '0.75' : '0.38';
    const strokeC = isDark ? '#a78bfa' : '#7c3aed';

    const countColorA = isDark ? '#93c5fd' : '#1d4ed8';
    const labelColorA = isDark ? '#cbd5e1' : '#334155';
    const subColorA = isDark ? '#93c5fd' : '#2563eb';

    const countColorCommon = isDark ? '#ffffff' : '#4338ca';
    const labelColorCommon = isDark ? '#ffffff' : '#1e1b4b';
    const subColorCommon = isDark ? '#e9d5ff' : '#6d28d9';
    const filterCommon = isDark ? 'filter: drop-shadow(0 0 10px rgba(255,255,255,0.8));' : 'filter: drop-shadow(0 1px 2px rgba(67,56,202,0.15));';

    const countColorB = isDark ? '#f9a8d4' : '#be185d';
    const labelColorB = isDark ? '#cbd5e1' : '#334155';
    const subColorB = isDark ? '#f9a8d4' : '#db2777';

    const svgContent = `
      <svg viewBox="0 0 640 360" width="100%" height="100%" style="overflow: visible;">
        <defs>
          <linearGradient id="gradCircleA" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="${gradA_start}" stop-opacity="${opA_start}"/>
            <stop offset="100%" stop-color="${gradA_end}" stop-opacity="${opA_end}"/>
          </linearGradient>

          <linearGradient id="gradCircleB" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="${gradB_start}" stop-opacity="${opB_start}"/>
            <stop offset="100%" stop-color="${gradB_end}" stop-opacity="${opB_end}"/>
          </linearGradient>

          <linearGradient id="gradCommon" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="${gradC_start}" stop-opacity="${opC}"/>
            <stop offset="100%" stop-color="${gradC_end}" stop-opacity="${opC}"/>
          </linearGradient>

          <clipPath id="clipCircleA">
            <circle cx="230" cy="180" r="140" />
          </clipPath>
        </defs>

        <!-- Circle A Full Disc -->
        <circle id="svgCircleA" cx="230" cy="180" r="140"
                fill="url(#gradCircleA)" stroke="${strokeA}" stroke-width="2.5"
                class="venn-circle" data-region="only_a" style="cursor: pointer;" />

        <!-- Circle B Full Disc -->
        <circle id="svgCircleB" cx="410" cy="180" r="140"
                fill="url(#gradCircleB)" stroke="${strokeB}" stroke-width="2.5"
                class="venn-circle" data-region="only_b" style="cursor: pointer;" />

        <!-- Intersection Lens (B clipped inside A) -->
        <circle id="svgIntersection" cx="410" cy="180" r="140"
                clip-path="url(#clipCircleA)" fill="url(#gradCommon)"
                stroke="${strokeC}" stroke-width="2.5"
                class="venn-intersection" data-region="intersection" style="cursor: pointer;" />

        <!-- Set A Only Count & Label -->
        <text x="160" y="172" class="venn-text-count" style="font-family: var(--font-heading); font-size: 28px; font-weight: 800; fill: ${countColorA}; text-anchor: middle;">${vennData.only_a_count}</text>
        <text x="160" y="202" class="venn-text-label" style="font-size: 13px; font-weight: 600; fill: ${labelColorA}; text-anchor: middle;">Yalnızca 1. Belge</text>
        <text x="160" y="220" class="venn-text-label" style="font-size: 11px; font-weight: 600; fill: ${subColorA}; text-anchor: middle;">(${escapeHtml(labelA)})</text>

        <!-- Intersection Count & Label -->
        <text x="320" y="165" class="venn-text-count" style="font-family: var(--font-heading); fill: ${countColorCommon}; font-size: 34px; font-weight: 800; text-anchor: middle; ${filterCommon}">${vennData.intersection_count}</text>
        <text x="320" y="198" class="venn-text-label" style="font-size: 14px; font-weight: 700; fill: ${labelColorCommon}; text-anchor: middle;">Ortak Kesişim</text>
        <text x="320" y="216" class="venn-text-label" style="font-size: 11.5px; font-weight: 600; fill: ${subColorCommon}; text-anchor: middle;">(Tekrarlayanlar)</text>

        <!-- Set B Only Count & Label -->
        <text x="480" y="172" class="venn-text-count" style="font-family: var(--font-heading); font-size: 28px; font-weight: 800; fill: ${countColorB}; text-anchor: middle;">${vennData.only_b_count}</text>
        <text x="480" y="202" class="venn-text-label" style="font-size: 13px; font-weight: 600; fill: ${labelColorB}; text-anchor: middle;">Yalnızca 2. Belge</text>
        <text x="480" y="220" class="venn-text-label" style="font-size: 11px; font-weight: 600; fill: ${subColorB}; text-anchor: middle;">(${escapeHtml(labelB)})</text>
      </svg>
    `;

    vennContainer.innerHTML = svgContent;

    // Attach click events on SVG regions
    vennContainer.querySelectorAll('[data-region]').forEach(el => {
      el.addEventListener('click', () => {
        setTableFilter(el.dataset.region);
      });
    });
  }

  // --------------------------------------------------------------------------
  // Table Rows & Filtering
  // --------------------------------------------------------------------------
  function setTableFilter(filter) {
    currentFilter = filter;
    updateFilterTabButtons();
    renderTableRows();
  }

  function updateFilterTabButtons() {
    [tabFilterAll, tabFilterIntersection, tabFilterA, tabFilterB, tabFilterDeclined].forEach(btn => {
      if (btn) btn.classList.remove('active');
    });
    if (currentFilter === 'all') tabFilterAll.classList.add('active');
    else if (currentFilter === 'intersection') tabFilterIntersection.classList.add('active');
    else if (currentFilter === 'only_a') tabFilterA.classList.add('active');
    else if (currentFilter === 'only_b') tabFilterB.classList.add('active');
    else if (currentFilter === 'declined' && tabFilterDeclined) tabFilterDeclined.classList.add('active');
  }

  function declineMatchedPair(pairIdx) {
    if (!currentAnalysis || !currentAnalysis.venn_verileri) return;
    const v = currentAnalysis.venn_verileri;
    if (!Array.isArray(v.matched_pairs) || pairIdx < 0 || pairIdx >= v.matched_pairs.length) return;

    const removedPair = v.matched_pairs.splice(pairIdx, 1)[0];
    if (!Array.isArray(v.declined_pairs)) v.declined_pairs = [];
    v.declined_pairs.push(removedPair);

    // Reddedilen eşleşmedeki A kaydını 1. Belge Benzersiz (Dry-Run), B kaydını 2. Belge Benzersiz listesine taşı
    if (!Array.isArray(v.only_a_items)) v.only_a_items = [];
    if (!Array.isArray(v.only_b_items)) v.only_b_items = [];

    if (removedPair.item_a && !v.only_a_items.includes(removedPair.item_a)) {
      v.only_a_items.unshift(removedPair.item_a);
    }
    if (removedPair.item_b && !v.only_b_items.includes(removedPair.item_b)) {
      v.only_b_items.unshift(removedPair.item_b);
    }

    refreshVennMetricsAndDiagram();
  }

  function restoreDeclinedPair(declinedIdx) {
    if (!currentAnalysis || !currentAnalysis.venn_verileri) return;
    const v = currentAnalysis.venn_verileri;
    if (!Array.isArray(v.declined_pairs) || declinedIdx < 0 || declinedIdx >= v.declined_pairs.length) return;

    const restoredPair = v.declined_pairs.splice(declinedIdx, 1)[0];
    if (!Array.isArray(v.matched_pairs)) v.matched_pairs = [];
    v.matched_pairs.push(restoredPair);

    if (Array.isArray(v.only_a_items)) {
      v.only_a_items = v.only_a_items.filter(x => x !== restoredPair.item_a);
    }
    if (Array.isArray(v.only_b_items)) {
      v.only_b_items = v.only_b_items.filter(x => x !== restoredPair.item_b);
    }

    refreshVennMetricsAndDiagram();
  }

  function refreshVennMetricsAndDiagram() {
    if (!currentAnalysis || !currentAnalysis.venn_verileri) return;
    const v = currentAnalysis.venn_verileri;

    v.intersection_count = (v.matched_pairs || []).length;
    v.only_a_count = (v.only_a_items || []).length;
    v.only_b_count = (v.only_b_items || []).length;
    v.total_unique = v.intersection_count + v.only_a_count + v.only_b_count;
    v.jaccard_similarity = v.total_unique > 0 ? (v.intersection_count / v.total_unique) : 0;
    v.intersection_items = (v.matched_pairs || []).map(p =>
      p.item_a !== p.item_b ? `${p.item_a} ↔ ${p.item_b} (%${p.similarity || 100})` : p.item_a
    );

    statTotal.textContent = v.total_unique;
    statIntersection.textContent = v.intersection_count;
    statJaccard.textContent = `Jaccard Benzerlik Oranı: %${Math.round(v.jaccard_similarity * 100)}`;
    statOnlyA.textContent = v.only_a_count;
    statOnlyB.textContent = v.only_b_count;

    countAll.textContent = v.total_unique;
    countIntersection.textContent = v.intersection_count;
    countA.textContent = v.only_a_count;
    countB.textContent = v.only_b_count;

    const declinedLen = (v.declined_pairs || []).length;
    if (countDeclined) countDeclined.textContent = declinedLen;
    if (tabFilterDeclined) {
      tabFilterDeclined.style.display = declinedLen > 0 ? 'inline-flex' : 'none';
      if (declinedLen === 0 && currentFilter === 'declined') {
        currentFilter = 'all';
        updateFilterTabButtons();
      }
    }

    const m1 = (currentAnalysis.belge_1_meta && Object.keys(currentAnalysis.belge_1_meta).length > 0)
      ? currentAnalysis.belge_1_meta
      : getBakiyeMetaPayload(1);
    const m2 = (currentAnalysis.belge_2_meta && Object.keys(currentAnalysis.belge_2_meta).length > 0)
      ? currentAnalysis.belge_2_meta
      : getBakiyeMetaPayload(2);

    const vennLabel1 = formatShortMetaLabel(m1, currentAnalysis.dosya_1_adi || '1. Belge');
    const vennLabel2 = formatShortMetaLabel(m2, currentAnalysis.dosya_2_adi || '2. Belge');
    drawVennDiagram(v, vennLabel1, vennLabel2);
    renderTableRows();
    renderDbKpiPreviewStrip();
  }

  function deleteDetailRow(region, idx, itemText) {
    if (!isCurrentUserAdmin()) {
      alert('Kayıt silme yetkisi yalnızca Admin kullanıcılara aittir.');
      return;
    }
    if (!currentAnalysis || !currentAnalysis.venn_verileri) return;
    const v = currentAnalysis.venn_verileri;

    if (region === 'intersection' && Array.isArray(v.matched_pairs) && idx >= 0 && idx < v.matched_pairs.length) {
      v.matched_pairs.splice(idx, 1);
    } else if (region === 'declined' && Array.isArray(v.declined_pairs) && idx >= 0 && idx < v.declined_pairs.length) {
      v.declined_pairs.splice(idx, 1);
    } else if (region === 'only_a' && Array.isArray(v.only_a_items)) {
      if (idx >= 0 && idx < v.only_a_items.length) v.only_a_items.splice(idx, 1);
      else if (itemText) v.only_a_items = v.only_a_items.filter(x => x !== itemText);
    } else if (region === 'only_b' && Array.isArray(v.only_b_items)) {
      if (idx >= 0 && idx < v.only_b_items.length) v.only_b_items.splice(idx, 1);
      else if (itemText) v.only_b_items = v.only_b_items.filter(x => x !== itemText);
    }

    refreshVennMetricsAndDiagram();
  }

  function renderTableRows() {
    if (!currentAnalysis) return;

    const v = currentAnalysis.venn_verileri;
    const query = tableSearchInput.value.toLowerCase().trim();
    const col1Query = colFilterDoc1 ? colFilterDoc1.value.toLowerCase().trim() : '';
    const col2Query = colFilterDoc2 ? colFilterDoc2.value.toLowerCase().trim() : '';
    const statusFilter = colFilterStatus ? colFilterStatus.value : 'all';
    const isAdmin = isCurrentUserAdmin();

    let rows = [];

    if (currentFilter === 'declined') {
      (v.declined_pairs || []).forEach((p, dIdx) => {
        rows.push({
          itemA: p.item_a,
          itemB: p.item_b,
          similarity: p.similarity,
          region: 'declined',
          matchType: 'Kullanıcı Reddetti (Benzer Değil)',
          declinedIdx: dIdx,
          rawIdx: dIdx
        });
      });
    } else {
      // 1. Matched Pairs (Intersection)
      (v.matched_pairs || []).forEach((p, pIdx) => {
        rows.push({
          itemA: p.item_a,
          itemB: p.item_b,
          similarity: p.similarity,
          region: 'intersection',
          matchType: p.match_type,
          pairIdx: pIdx,
          rawIdx: pIdx
        });
      });

      // 2. Only A
      (v.only_a_items || []).forEach((item, aIdx) => {
        rows.push({
          itemA: item,
          itemB: '-',
          similarity: 0,
          region: 'only_a',
          matchType: 'Yalnızca 1. Belge (Dry-Run)',
          rawIdx: aIdx
        });
      });

      // 3. Only B
      (v.only_b_items || []).forEach((item, bIdx) => {
        rows.push({
          itemA: '-',
          itemB: item,
          similarity: 0,
          region: 'only_b',
          matchType: 'Yalnızca 2. Belge',
          rawIdx: bIdx
        });
      });

      // Filter by Region Tab
      if (currentFilter !== 'all') {
        rows = rows.filter(r => r.region === currentFilter);
      }
    }

    // Filter by General Search Query
    if (query) {
      rows = rows.filter(r => (
        (r.itemA && r.itemA.toLowerCase().includes(query)) ||
        (r.itemB && r.itemB.toLowerCase().includes(query)) ||
        (r.matchType && r.matchType.toLowerCase().includes(query))
      ));
    }

    // Filter by Column 1 (1. Belge)
    if (col1Query) {
      rows = rows.filter(r => r.itemA && r.itemA.toLowerCase().includes(col1Query));
    }

    // Filter by Column 2 (2. Belge)
    if (col2Query) {
      rows = rows.filter(r => r.itemB && r.itemB.toLowerCase().includes(col2Query));
    }

    // Filter by Status Dropdown
    if (statusFilter && statusFilter !== 'all') {
      if (statusFilter === 'exact') {
        rows = rows.filter(r => r.region === 'intersection' && r.similarity >= 99.5);
      } else if (statusFilter === 'fuzzy') {
        rows = rows.filter(r => r.region === 'intersection' && r.similarity < 99.5);
      } else if (statusFilter === 'only_a') {
        rows = rows.filter(r => r.region === 'only_a');
      } else if (statusFilter === 'only_b') {
        rows = rows.filter(r => r.region === 'only_b');
      }
    }

    if (detailFilteredCountBadge) {
      detailFilteredCountBadge.textContent = `${rows.length} kayıt`;
    }

    if (rows.length === 0) {
      detailsTableBody.innerHTML = `
        <tr>
          <td colspan="5" style="text-align: center; padding: 24px; color: var(--text-muted);">
            Filtre ve arama kriterlerine uygun kayıt bulunamadı.
          </td>
        </tr>
      `;
      return;
    }

    detailsTableBody.innerHTML = rows.map((r, idx) => {
      let badgeHtml = '';
      let primaryBtnHtml = '';

      if (r.region === 'intersection') {
        if (r.similarity >= 99.5) {
          badgeHtml = `<span class="badge-similarity-exact">✓ %100 Tam Eşleşme</span>`;
        } else {
          badgeHtml = `<span class="badge-similarity-fuzzy">≈ %${r.similarity} Benzer</span>`;
        }
        primaryBtnHtml = `
          <button type="button" class="btn-decline-match" data-pair-idx="${r.pairIdx}" title="Bu iki kaydın benzer olmadığını işaretle ve ayrı benzersiz uygunsuzluklara ayır">
            ✕ Benzer Değil
          </button>
        `;
      } else if (r.region === 'declined') {
        badgeHtml = `<span class="badge-similarity-declined">✕ Reddedildi (%${r.similarity})</span>`;
        primaryBtnHtml = `
          <button type="button" class="btn-restore-match" data-declined-idx="${r.declinedIdx}" title="Reddedilen eşleşmeyi tekrar ortak kesişime geri al">
            ↩ Eşleşmeyi Geri Al
          </button>
        `;
      } else if (r.region === 'only_a') {
        badgeHtml = `<span class="badge-similarity-only a">Yalnızca 1. Belge</span>`;
      } else {
        badgeHtml = `<span class="badge-similarity-only b">Yalnızca 2. Belge</span>`;
      }

      const adminDeleteBtnHtml = isAdmin
        ? `<button type="button" class="btn-delete-detail-row admin-only-item" data-region="${r.region}" data-idx="${r.rawIdx}" title="Admin Yetkisi: Bu satırı kıyaslama listesinden sil">🗑️ Sil</button>`
        : '';

      const actionHtml = (primaryBtnHtml || adminDeleteBtnHtml)
        ? `<div style="display: inline-flex; gap: 6px; align-items: center; justify-content: flex-end;">${primaryBtnHtml}${adminDeleteBtnHtml}</div>`
        : `<span style="color: var(--text-dim); font-size: 11.5px;">—</span>`;

      return `
        <tr>
          <td style="color: var(--text-dim); font-family: var(--font-mono);">${idx + 1}</td>
          <td style="font-weight: ${r.itemA !== '-' ? '500' : 'normal'}; color: ${r.itemA === '-' ? 'var(--text-dim)' : 'var(--text-main)'};">
            ${escapeHtml(r.itemA)}
          </td>
          <td style="font-weight: ${r.itemB !== '-' ? '500' : 'normal'}; color: ${r.itemB === '-' ? 'var(--text-dim)' : 'var(--text-main)'};">
            ${escapeHtml(r.itemB)}
          </td>
          <td>${badgeHtml}</td>
          <td style="text-align: right;">${actionHtml}</td>
        </tr>
      `;
    }).join('');
  }

  // --------------------------------------------------------------------------
  // Export CSV & JSON
  // --------------------------------------------------------------------------
  function exportTableToCSV() {
    if (!currentAnalysis) return;
    const v = currentAnalysis.venn_verileri;

    let csvContent = 'Sıra,1. Belge Kaydı,2. Belge Karşılığı,Benzerlik Puanı,Durum / Bölge\n';
    let rowIdx = 1;

    (v.matched_pairs || []).forEach(p => {
      csvContent += `"${rowIdx++}","${csvEscape(p.item_a)}","${csvEscape(p.item_b)}","%${p.similarity}","Ortak Kesişim"\n`;
    });
    (v.only_a_items || []).forEach(item => {
      csvContent += `"${rowIdx++}","${csvEscape(item)}","-","0%","Yalnızca 1. Belge"\n`;
    });
    (v.only_b_items || []).forEach(item => {
      csvContent += `"${rowIdx++}","-","${csvEscape(item)}","0%","Yalnızca 2. Belge"\n`;
    });

    downloadBlob(csvContent, 'venn_kiyaslama_raporu.csv', 'text/csv;charset=utf-8;');
  }

  function exportTableToJSON() {
    if (!currentAnalysis) return;
    const jsonStr = JSON.stringify(currentAnalysis, null, 2);
    downloadBlob(jsonStr, 'venn_kiyaslama_raporu.json', 'application/json;charset=utf-8;');
  }

  function downloadBlob(content, filename, type) {
    const blob = new Blob(['\uFEFF' + content], { type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function copySummaryToClipboard() {
    if (!summaryText.textContent) return;
    navigator.clipboard.writeText(summaryText.textContent.trim()).then(() => {
      btnCopySummary.innerHTML = `<span>✓ Kopyalandı!</span>`;
      setTimeout(() => {
        btnCopySummary.innerHTML = `
          <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2">
            <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
            <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
          </svg>
          <span>Kopyala</span>
        `;
      }, 2000);
    });
  }

  // --------------------------------------------------------------------------
  // Chat Assistant with AselixAI
  // --------------------------------------------------------------------------
  async function handleSendChatMessage() {
    const msg = chatInput.value.trim();
    if (!msg) return;

    // Append user message
    appendChatMessage('user', msg);
    chatInput.value = '';
    btnSendChat.disabled = true;

    // Typing placeholder
    const typingId = appendTypingIndicator();

    // Prepare context
    let contextObj = null;
    if (currentAnalysis) {
      contextObj = {
        dosya_1: currentAnalysis.dosya_1_adi,
        dosya_2: currentAnalysis.dosya_2_adi,
        intersection_count: currentAnalysis.venn_verileri?.intersection_count || 0,
        only_a_count: currentAnalysis.venn_verileri?.only_a_count || 0,
        only_b_count: currentAnalysis.venn_verileri?.only_b_count || 0,
        jaccard_similarity: currentAnalysis.venn_verileri?.jaccard_similarity || 0,
        esik_orani: currentAnalysis.esik_orani || 0.8
      };
    }

    try {
      const resp = await fetch('/api/aselix/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          base_url: appConfig.baseUrl,
          api_key: appConfig.apiKey,
          model: appConfig.model || 'aselix-llm-v1',
          messages: [{ role: 'user', content: msg }],
          temperature: 0.2,
          context: contextObj
        })
      });

      const data = await resp.json();
      removeTypingIndicator(typingId);
      btnSendChat.disabled = false;

      if (data.durum === 'basarili') {
        appendChatMessage('bot', data.cevap);
      } else {
        const fallbackErrMsg = isCurrentUserAdmin()
          ? 'Cevap üretilemedi. Lütfen Ayarlar menüsünden API anahtarınızı ve ağ seçiminizi kontrol edin.'
          : 'Şu anda yapay zeka sohbet servisine ulaşılamıyor. Lütfen sistem yöneticinizle iletişime geçin.';
        const cleanMsg = isCurrentUserAdmin()
          ? (data.mesaj || fallbackErrMsg)
          : (sanitizeAdminTextForUser(data.mesaj) || fallbackErrMsg);
        appendChatMessage('bot', `⚠️ ${cleanMsg}`);
      }
    } catch (err) {
      removeTypingIndicator(typingId);
      btnSendChat.disabled = false;
      appendChatMessage('bot', `⚠️ Bağlantı hatası: ${err.message}`);
    }
  }

  function appendChatMessage(sender, text) {
    const msgDiv = document.createElement('div');
    msgDiv.className = `chat-msg ${sender === 'user' ? 'user-msg' : 'bot-msg'}`;

    const avatar = sender === 'user' ? '👤' : '🤖';
    msgDiv.innerHTML = `
      <div class="msg-avatar">${avatar}</div>
      <div class="msg-bubble">${sender === 'user' ? `<p>${escapeHtml(text)}</p>` : markdownToHtml(text)}</div>
    `;

    chatMessages.appendChild(msgDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }

  function appendTypingIndicator() {
    const id = 'typing_' + Date.now();
    const div = document.createElement('div');
    div.id = id;
    div.className = 'chat-msg bot-msg';
    div.innerHTML = `
      <div class="msg-avatar">🤖</div>
      <div class="msg-bubble"><span class="typing-dots"><span>.</span><span>.</span><span>.</span> AselixAI düşünüyor</span></div>
    `;
    chatMessages.appendChild(div);
    chatMessages.scrollTop = chatMessages.scrollHeight;
    return id;
  }

  function removeTypingIndicator(id) {
    const el = document.getElementById(id);
    if (el) el.remove();
  }

  // --------------------------------------------------------------------------
  // Settings Modal Functions (Yalnızca Admin Yetkisi)
  // --------------------------------------------------------------------------
  function openSettingsModal() {
    if (!isCurrentUserAdmin()) {
      alert('Bu bölüm yalnızca sistem yöneticisi (Admin) yetkisine sahip kullanıcılara açıktır.');
      return;
    }
    cfgBaseUrl.value = appConfig.baseUrl || NETWORK_URLS.milli;
    cfgApiKey.value = appConfig.apiKey || '';
    if (cfgBatchSize) cfgBatchSize.value = String(appConfig.batchSize || 15);

    if (appConfig.network === 'aynet') radioAynet.checked = true;
    else if (appConfig.network === 'custom') radioCustom.checked = true;
    else radioMilli.checked = true;

    testConnResult.style.display = 'none';
    fetchModelsStatus.style.display = 'none';
    settingsModal.style.display = 'flex';
  }

  function closeSettingsModal() {
    settingsModal.style.display = 'none';
  }

  function saveSettingsFromModal() {
    let network = 'milli';
    if (radioAynet.checked) network = 'aynet';
    else if (radioCustom.checked) network = 'custom';

    appConfig.network = network;
    appConfig.baseUrl = cfgBaseUrl.value.trim() || NETWORK_URLS.milli;
    appConfig.apiKey = cfgApiKey.value.trim();
    appConfig.model = cfgModelSelect.value || 'aselix-llm-v1';
    if (cfgBatchSize) appConfig.batchSize = parseInt(cfgBatchSize.value, 10) || 15;
    if (cfgEmbeddingModel) appConfig.embeddingModel = cfgEmbeddingModel.value.trim() || 'text-embedding-3-small';

    saveConfig();
    applyConfigToUI();
    closeSettingsModal();
  }

  // --------------------------------------------------------------------------
  // Kokpit SSO / Iframe Push & External Input Bridge (KIMLIK_AKTARIMI_ENTEGRASYON_REHBERI)
  // --------------------------------------------------------------------------
  function getAllowedOriginsList() {
    const raw = (appConfig.allowedOrigins || DEFAULT_ALLOWED_ORIGINS.join(',')).split(',');
    const list = raw.map(s => s.trim()).filter(Boolean);
    DEFAULT_ALLOWED_ORIGINS.forEach(o => {
      if (!list.includes(o)) list.push(o);
    });
    if (window.location && window.location.origin && !list.includes(window.location.origin)) {
      list.push(window.location.origin);
    }
    return list;
  }

  function isOriginAllowed(origin) {
    if (!origin) return true;
    const allowed = getAllowedOriginsList();
    if (allowed.includes('*') || allowed.includes(origin)) return true;
    if (origin.includes(':3004') || origin.endsWith('.aselsan.com.tr') || origin === window.location.origin) {
      return true;
    }
    return false;
  }

  function initKokpitIdentityBridge() {
    // 1. Iframe içinde gömülü (embedded) olup olmadığını algıla
    const urlParams = new URLSearchParams(window.location.search);
    const isIframe = (window.self !== window.top) || urlParams.get('embed') === '1' || urlParams.get('iframe') === '1';
    if (isIframe) {
      document.body.classList.add('iframe-embedded');
    }

    // 2. window.postMessage dinleyicisini kur (Rehber Bölüm 5.A & 5.B)
    let resolvedFromMessage = false;
    window.addEventListener('message', async function (event) {
      if (!isOriginAllowed(event.origin)) {
        console.warn('[Kokpit Bridge] İzin verilmeyen origin reddedildi:', event.origin);
        return;
      }

      const d = event.data;
      if (!d || typeof d !== 'object') return;

      // A) Rehber Bölüm 4: Standart kokpit-identity paketi
      if (d.type === 'kokpit-identity') {
        if (d.source && d.source !== 'kokpit-iframe-bridge') return;
        if (d.status && d.status !== 'success') return;

        resolvedFromMessage = true;
        await handleReceivedKokpitIdentity(d, event.origin);
      }
      // B) Dışarıdan doğrudan gönderilen veri / embedding / input paketleri
      else if (['kokpit-input', 'kokpit-data', 'external-input', 'vennflow-input'].includes(d.type)) {
        resolvedFromMessage = true;
        await handleReceivedKokpitIdentity(d, event.origin);
      }
    });

    // 3. URL Query Parametreleri ile dışarıdan gelen input / kimlik varsa al
    parseUrlExternalInputs(urlParams);

    // 4. Ana pencereye (Kokpit) iframe'in hazır olduğunu bildir
    if (window.parent && window.parent !== window.self) {
      try {
        window.parent.postMessage({
          type: 'kokpit-iframe-ready',
          source: 'vennflow-app',
          timestamp: new Date().toISOString()
        }, '*');
      } catch (_) {}
    }

    // 5. Rehber Bölüm 5.B: 4 saniye içinde mesaj gelmezse sessionStorage veya Backend oturumuna bak
    setTimeout(async () => {
      if (!resolvedFromMessage && !kokpitIdentity.user) {
        const cachedUserStr = sessionStorage.getItem('kokpit_user') || sessionStorage.getItem('current_user');
        const cachedToken = sessionStorage.getItem('kokpit_token') || sessionStorage.getItem('auth_token');
        if (cachedUserStr) {
          try {
            const cachedUser = JSON.parse(cachedUserStr);
            kokpitIdentity = {
              ...kokpitIdentity,
              user: cachedUser,
              token: cachedToken || null,
              loading: false,
              source: 'sessionStorage'
            };
            updateIdentityUI();
            pushIdentityToBackend({ user: cachedUser, token: cachedToken, source: 'kokpit-iframe-bridge' });
            return;
          } catch (_) {}
        }

        // Backend'de aktif kimlik veya dış input var mı kontrol et
        try {
          const resp = await fetch('/api/identity');
          const data = await resp.json();
          if (data && data.user) {
            kokpitIdentity = {
              user: data.user,
              token: kokpitIdentity.token || null,
              loading: false,
              source: data.source || 'backend-session',
              origin: data.origin || null,
              timestamp: data.timestamp || null,
              externalInputs: data.external_inputs || {}
            };
            updateIdentityUI();
            if (data.external_inputs && Object.keys(data.external_inputs).length > 0) {
              applyExternalInputsToApp(data.external_inputs);
            }
            return;
          }
        } catch (_) {}

        kokpitIdentity.loading = false;
        updateIdentityUI();
      }
    }, 4000);
  }

  function parseUrlExternalInputs(urlParams) {
    const ext = {};
    const token = urlParams.get('token') || urlParams.get('auth_token');
    const sicil = urlParams.get('sicil');
    const isim = urlParams.get('isim') || urlParams.get('name');
    const username = urlParams.get('username');
    const mudurluk = urlParams.get('mudurluk');
    const role = urlParams.get('role');

    ['api_key', 'base_url', 'model', 'embedding_model', 'engine', 'threshold', 'text_1', 'text_2', 'name_1', 'name_2', 'theme', 'auto_run'].forEach(k => {
      if (urlParams.has(k)) ext[k] = urlParams.get(k);
    });

    if (token || sicil || isim || Object.keys(ext).length > 0) {
      const payload = {
        type: 'kokpit-identity',
        status: 'success',
        source: 'kokpit-iframe-bridge',
        token: token || undefined,
        user: (sicil || isim || username) ? {
          sicil: sicil || '',
          isim: isim || username || '',
          username: username || '',
          mudurluk: mudurluk || '',
          role: role || 'User'
        } : undefined,
        inputs: ext
      };
      handleReceivedKokpitIdentity(payload, window.location.origin);
    }
  }

  async function handleReceivedKokpitIdentity(payload, origin) {
    const token = payload.token || payload.auth_token || kokpitIdentity.token || null;
    const user = payload.user || kokpitIdentity.user || null;

    // Rehber Bölüm 5.A & 5.B: Oturum süresince sessionStorage'da sakla
    if (token) {
      sessionStorage.setItem('auth_token', token);
      sessionStorage.setItem('kokpit_token', token);
    }
    if (user) {
      sessionStorage.setItem('current_user', JSON.stringify(user));
      sessionStorage.setItem('kokpit_user', JSON.stringify(user));
      console.log('✅ Kimlik başarıyla alındı:', user.isim || user.username, `(Sicil: ${user.sicil || '-'})`);
    }

    // Dışarıdan gelen tüm ekstra inputları topla
    const mergedInputs = {
      ...(kokpitIdentity.externalInputs || {}),
      ...(payload.inputs || {}),
      ...(payload.config || {}),
      ...(payload.data || {})
    };
    [
      'api_key', 'apiKey', 'base_url', 'baseUrl', 'model', 'embedding_model', 'embeddingModel',
      'engine', 'engine_mode', 'threshold', 'batchSize', 'batch_size',
      'text_1', 'text_2', 'text1', 'text2', 'items_1', 'items_2', 'items_a', 'items_b',
      'name_1', 'name_2', 'embeddings_1', 'embeddings_2', 'embeddings_a', 'embeddings_b',
      'theme', 'auto_run', 'autoRun'
    ].forEach(k => {
      if (payload[k] !== undefined && payload[k] !== null) {
        mergedInputs[k] = payload[k];
      }
    });

    kokpitIdentity = {
      user,
      token,
      loading: false,
      source: payload.source || 'kokpit-iframe-bridge',
      origin: origin || window.location.origin,
      timestamp: payload.timestamp || new Date().toISOString(),
      externalInputs: mergedInputs
    };

    // UI'ı güncelle
    updateIdentityUI();

    // Dışarıdan gelen input/embedding/ayar parametrelerini uygulamaya uygula
    if (Object.keys(mergedInputs).length > 0) {
      applyExternalInputsToApp(mergedInputs);
    }

    // Rehber Bölüm 5.C (Tier-2): Kendi Backend'imize (/api/identity) POST et
    await pushIdentityToBackend({
      type: payload.type || 'kokpit-identity',
      status: 'success',
      source: payload.source || 'kokpit-iframe-bridge',
      timestamp: kokpitIdentity.timestamp,
      origin: kokpitIdentity.origin,
      user,
      token,
      inputs: mergedInputs
    });

    // Ana pencereye (Kokpit) alındı teyidi gönder
    if (window.parent && window.parent !== window.self) {
      try {
        window.parent.postMessage({
          type: 'kokpit-identity-ack',
          status: 'success',
          sicil: user ? user.sicil : null
        }, '*');
      } catch (_) {}
    }
  }

  async function pushIdentityToBackend(payload) {
    try {
      const resp = await fetch('/api/identity', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await resp.json();
      if (data && data.user && !kokpitIdentity.user) {
        kokpitIdentity.user = data.user;
        updateIdentityUI();
      }
      return data;
    } catch (err) {
      console.warn('[Kokpit Bridge] Backend /api/identity aktarım uyarısı:', err);
      return null;
    }
  }

  function applyExternalInputsToApp(ext) {
    if (!ext || typeof ext !== 'object') return;

    if (ext.theme === 'dark' || ext.theme === 'light') {
      applyTheme(ext.theme);
    }
    if (ext.base_url || ext.baseUrl) {
      appConfig.baseUrl = ext.base_url || ext.baseUrl;
    }
    if (ext.api_key || ext.apiKey) {
      appConfig.apiKey = ext.api_key || ext.apiKey;
    }
    if (ext.model) {
      appConfig.model = ext.model;
    }
    if (ext.embedding_model || ext.embeddingModel) {
      appConfig.embeddingModel = ext.embedding_model || ext.embeddingModel;
    }
    if (ext.threshold !== undefined) {
      let thrVal = parseFloat(ext.threshold);
      if (thrVal <= 1.0) thrVal = Math.round(thrVal * 100);
      appConfig.threshold = Math.max(50, Math.min(100, thrVal));
    }
    const eng = ext.engine_mode || ext.engine;
    if (eng && ['local', 'embedding', 'aselix'].includes(eng)) {
      appConfig.engine = eng;
    }

    // Dışarıdan metin veya liste verisi geldiyse Copy-Paste sekmesine doldur
    const list1 = Array.isArray(ext.items_1 || ext.items_a) ? (ext.items_1 || ext.items_a).join('\n') : (ext.text_1 || ext.text1);
    const list2 = Array.isArray(ext.items_2 || ext.items_b) ? (ext.items_2 || ext.items_b).join('\n') : (ext.text_2 || ext.text2);

    if (list1 || list2) {
      switchMainPageTab('vennflow');
      setInputMode('paste');
      if (list1 && textPaste1) textPaste1.value = list1;
      if (list2 && textPaste2) textPaste2.value = list2;
      updatePasteCounters();
    }

    applyConfigToUI();

    if ((ext.auto_run === true || ext.auto_run === '1' || ext.autoRun === true) && textPaste1.value.trim() && textPaste2.value.trim()) {
      setTimeout(() => runUnifiedComparison(), 250);
    }
  }

  function getInitials(name) {
    if (!name) return '👤';
    const parts = String(name).trim().split(/\s+/);
    if (parts.length >= 2) {
      return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    }
    return String(name).slice(0, 2).toUpperCase();
  }

  function updateIdentityUI() {
    const u = kokpitIdentity.user;
    const portalIdentityText = document.getElementById('portalIdentityText');
    if (u && (u.isim || u.username || u.sicil)) {
      const displayName = u.isim || u.username || `Sicil: ${u.sicil}`;
      const subInfo = u.mudurluk || u.birim || u.unvan || u.role || 'Kokpit SSO';

      if (btnOpenIdentity) {
        btnOpenIdentity.classList.remove('waiting');
        btnOpenIdentity.classList.add('authenticated');
      }
      if (authStatusEl) {
        authStatusEl.innerText = `Hoş geldiniz, ${displayName} (${u.mudurluk || u.role || 'SSO'})`;
      }
      if (portalIdentityText) {
        portalIdentityText.textContent = `${displayName}${u.role ? ' (' + u.role + ')' : ''}`;
      }
      if (identityUserMeta) {
        identityUserMeta.textContent = `${u.sicil ? 'Sicil: ' + u.sicil + ' • ' : ''}${subInfo}`;
      }
      if (identityAvatar) {
        identityAvatar.textContent = getInitials(displayName);
      }
    } else {
      if (btnOpenIdentity) {
        btnOpenIdentity.classList.remove('authenticated');
        btnOpenIdentity.classList.add('waiting');
      }
      if (authStatusEl) {
        authStatusEl.innerText = kokpitIdentity.loading
          ? 'Kimlik bilgisi bekleniyor...'
          : 'Bağımsız Mod (Kokpit SSO)';
      }
      if (portalIdentityText) {
        portalIdentityText.textContent = 'Kokpit SSO';
      }
      if (identityUserMeta) {
        identityUserMeta.textContent = kokpitIdentity.loading
          ? 'Kokpit Iframe Köprüsü dinleniyor'
          : 'Simüle etmek veya görmek için tıklayın';
      }
      if (identityAvatar) {
        identityAvatar.textContent = '🔐';
      }
    }

    applyAdminVisibility();
    renderIdentityModalContent();
  }

  function renderIdentityModalContent() {
    if (!identityDetailsGrid) return;
    const u = kokpitIdentity.user;
    if (u && (u.isim || u.username || u.sicil)) {
      const displayName = u.isim || u.username || u.sicil;
      if (modalIdentityAvatar) modalIdentityAvatar.textContent = getInitials(displayName);
      if (modalIdentityName) modalIdentityName.textContent = `${displayName} (${u.role || 'User'})`;
      if (modalIdentitySub) modalIdentitySub.textContent = `${u.unvan || ''} • ${u.mudurluk || u.direktorluk || ''}`;
      if (modalIdentityBadge) {
        modalIdentityBadge.textContent = '✓ SSO Aktif';
        modalIdentityBadge.style.background = 'var(--color-emerald-subtle)';
        modalIdentityBadge.style.color = 'var(--color-emerald)';
      }

      const fields = [
        ['Ad Soyad (isim)', u.isim || '-'],
        ['Sicil No', u.sicil || '-'],
        ['Kullanıcı Adı', u.username || '-'],
        ['E-Posta (mail)', u.mail || '-'],
        ['Rol (role)', u.role || '-'],
        ['Unvan', u.unvan || '-'],
        ['Müdürlük', u.mudurluk ? `${u.mudurluk} ${u.mudurluk_id ? '(#' + u.mudurluk_id + ')' : ''}` : '-'],
        ['Direktörlük / Birim', [u.direktorluk, u.birim].filter(Boolean).join(' / ') || '-'],
        ['Hiyerarşik Seviye', u.hierarchicLevel !== undefined && u.hierarchicLevel !== null ? String(u.hierarchicLevel) : '-'],
        ['JWT Bearer Token', kokpitIdentity.token ? `✓ Mevcut (${kokpitIdentity.token.slice(0, 18)}…)` : 'Yok'],
        ['Kaynak (source)', kokpitIdentity.source || 'kokpit-iframe-bridge'],
        ['Dış Input Alanları', Object.keys(kokpitIdentity.externalInputs || {}).join(', ') || 'Yok']
      ];

      identityDetailsGrid.innerHTML = fields.map(([lbl, val]) => `
        <div class="identity-kv">
          <span class="identity-kv-label">${escapeHtml(lbl)}</span>
          <span class="identity-kv-val">${escapeHtml(String(val))}</span>
        </div>
      `).join('');
    } else {
      if (modalIdentityAvatar) modalIdentityAvatar.textContent = '🔐';
      if (modalIdentityName) modalIdentityName.textContent = 'Henüz Kokpit Kimlik Paketi Alınmadı';
      if (modalIdentitySub) modalIdentitySub.textContent = 'Ana portaldan postMessage({ type: "kokpit-identity" }) bekleniyor veya aşağıdan test edebilirsiniz.';
      if (modalIdentityBadge) {
        modalIdentityBadge.textContent = kokpitIdentity.loading ? 'Dinleniyor...' : 'Bağımsız Mod';
        modalIdentityBadge.style.background = 'var(--color-primary-subtle)';
        modalIdentityBadge.style.color = 'var(--color-primary)';
      }
      identityDetailsGrid.innerHTML = `
        <div class="identity-kv" style="grid-column: span 2;">
          <span class="identity-kv-label">Durum</span>
          <span class="identity-kv-val">Ana sistem (Kokpit) içerisine iframe olarak gömüldüğünde kimlik ve dış inputlar otomatik aktarılır.</span>
        </div>
      `;
    }
  }

  function getDefaultIdentitySamplePayload(roleType = 'Admin') {
    const isAdminRole = String(roleType || '').toLowerCase() === 'admin';
    return {
      type: 'kokpit-identity',
      status: 'success',
      source: 'kokpit-iframe-bridge',
      timestamp: new Date().toISOString(),
      token: 'eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NSIsInByZWZlcnJlZF91c2VybmFtZSI6ImFobWV0LnlpbG1heiIsIm5hbWUiOiJBaG1ldCBZxLFsbWF6In0.signature',
      user: {
        id: isAdminRole ? 145 : 210,
        sicil: isAdminRole ? '12345' : '67890',
        username: isAdminRole ? 'ahmet.yilmaz' : 'mehmet.kaya',
        isim: isAdminRole ? 'Ahmet Yılmaz' : 'Mehmet Kaya',
        mail: isAdminRole ? 'ahmet.yilmaz@sirket.com' : 'mehmet.kaya@sirket.com',
        picture: isAdminRole ? '/uploads/photos/12345.jpg' : '/uploads/photos/67890.jpg',
        role: isAdminRole ? 'Admin' : 'User',
        mudurluk: 'UGES Teslimat Yönetim Müdürlüğü',
        mudurluk_id: 191148,
        direktorluk: 'Program Direktörlüğü',
        birim: 'Planlama Birimi',
        unvan: isAdminRole ? 'Sistem Yöneticisi / Uzman' : 'Kalite Mühendisi',
        hierarchicLevel: isAdminRole ? 2 : 3
      },
      inputs: {
        engine_mode: 'embedding',
        embedding_model: 'text-embedding-3-small',
        threshold: 80,
        items_1: [
          'Pano içi etiketleme eksik',
          'Topraklama barası cıvatası gevşek',
          'Acil durdurma butonu mekanik takılma yapıyor'
        ],
        items_2: [
          'etiketleme eksik pano içi',
          'Topraklama barası cıvatası gevşek',
          'Optik sensör kalibrasyon sapması'
        ],
        auto_run: false
      }
    };
  }

  function openIdentityModal() {
    if (!identityModal) return;
    if (cfgAllowedOrigins) {
      cfgAllowedOrigins.value = appConfig.allowedOrigins || DEFAULT_ALLOWED_ORIGINS.join(', ');
    }
    if (simPostMessagePayload && !simPostMessagePayload.value.trim()) {
      simPostMessagePayload.value = JSON.stringify(getDefaultIdentitySamplePayload('User'), null, 2);
    }
    if (identityActionStatus) identityActionStatus.style.display = 'none';
    renderIdentityModalContent();
    identityModal.style.display = 'flex';
  }

  function closeIdentityModal() {
    if (identityModal) identityModal.style.display = 'none';
  }

  function fillStandardUserIdentityPayload() {
    if (simPostMessagePayload) {
      simPostMessagePayload.value = JSON.stringify(getDefaultIdentitySamplePayload('User'), null, 2);
    }
  }

  function fillDefaultIdentityPayload() {
    if (simPostMessagePayload) {
      simPostMessagePayload.value = JSON.stringify(getDefaultIdentitySamplePayload('Admin'), null, 2);
    }
  }

  function simulateKokpitPostMessage() {
    if (!simPostMessagePayload) return;
    try {
      const parsed = JSON.parse(simPostMessagePayload.value);
      window.postMessage(parsed, window.location.origin);
      if (identityActionStatus) {
        identityActionStatus.style.display = 'block';
        identityActionStatus.style.background = 'rgba(16, 185, 129, 0.15)';
        identityActionStatus.style.border = '1px solid rgba(16, 185, 129, 0.4)';
        identityActionStatus.style.color = 'var(--color-emerald)';
        identityActionStatus.innerHTML = '✓ <strong>kokpit-identity</strong> mesajı <code>window.postMessage</code> üzerinden gönderildi, oturum ve dış inputlar aktarıldı!';
      }
    } catch (e) {
      if (identityActionStatus) {
        identityActionStatus.style.display = 'block';
        identityActionStatus.style.background = 'rgba(239, 68, 68, 0.15)';
        identityActionStatus.style.border = '1px solid rgba(239, 68, 68, 0.4)';
        identityActionStatus.style.color = 'var(--color-danger)';
        identityActionStatus.textContent = 'JSON Format Hatası: ' + e.message;
      }
    }
  }

  async function clearKokpitIdentitySession() {
    sessionStorage.removeItem('auth_token');
    sessionStorage.removeItem('current_user');
    sessionStorage.removeItem('kokpit_token');
    sessionStorage.removeItem('kokpit_user');
    kokpitIdentity = {
      user: null,
      token: null,
      loading: false,
      source: null,
      origin: null,
      timestamp: null,
      externalInputs: {}
    };
    try {
      await fetch('/api/identity', { method: 'DELETE' });
    } catch (_) {}
    updateIdentityUI();
    if (identityActionStatus) {
      identityActionStatus.style.display = 'block';
      identityActionStatus.style.background = 'var(--color-primary-subtle)';
      identityActionStatus.style.border = '1px solid rgba(37, 99, 235, 0.3)';
      identityActionStatus.style.color = 'var(--color-primary)';
      identityActionStatus.textContent = 'ℹ️ Oturum ve dış input verileri temizlendi.';
    }
  }

  function saveIdentityModalConfig() {
    if (cfgAllowedOrigins) {
      appConfig.allowedOrigins = cfgAllowedOrigins.value.trim() || DEFAULT_ALLOWED_ORIGINS.join(', ');
      saveConfig();
    }
    closeIdentityModal();
  }

  // Global Bridge nesnesi (Dış pencere veya konsol erişimi için)
  window.VennFlowKokpitBridge = {
    getIdentity: () => kokpitIdentity,
    pushIdentity: (payload) => handleReceivedKokpitIdentity(payload, window.location.origin),
    applyExternalInputs: applyExternalInputsToApp,
    clearIdentity: clearKokpitIdentitySession
  };

  function resetSettingsToDefault() {
    appConfig = { ...DEFAULT_CONFIG };
    saveConfig();
    applyConfigToUI();
    openSettingsModal();
  }

  function toggleApiKeyVisibility() {
    if (cfgApiKey.type === 'password') {
      cfgApiKey.type = 'text';
      btnToggleApiKey.textContent = 'Gizle';
    } else {
      cfgApiKey.type = 'password';
      btnToggleApiKey.textContent = 'Göster';
    }
  }

  async function fetchModelsFromApi() {
    btnFetchModels.disabled = true;
    btnFetchModels.textContent = '⏳ Çekiliyor...';
    fetchModelsStatus.style.display = 'none';

    try {
      const resp = await fetch('/api/aselix/models', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          base_url: cfgBaseUrl.value.trim(),
          api_key: cfgApiKey.value.trim()
        })
      });

      const data = await resp.json();
      btnFetchModels.disabled = false;
      btnFetchModels.textContent = '🔄 Modelleri Listele / Getir';

      if (data.modeller && data.modeller.length > 0) {
        cfgModelSelect.innerHTML = '';
        data.modeller.forEach(m => {
          const opt = document.createElement('option');
          opt.value = m;
          opt.textContent = m;
          if (m === appConfig.model) opt.selected = true;
          cfgModelSelect.appendChild(opt);
        });

        fetchModelsStatus.style.display = 'block';
        fetchModelsStatus.style.background = 'rgba(16, 185, 129, 0.15)';
        fetchModelsStatus.style.color = '#34d399';
        fetchModelsStatus.style.border = '1px solid rgba(16, 185, 129, 0.3)';
        fetchModelsStatus.textContent = `✓ ${data.modeller.length} model listelendi (${data.kaynak === 'canli_api' ? 'Canlı API' : 'Önbellek'}).`;
      } else {
        fetchModelsStatus.style.display = 'block';
        fetchModelsStatus.style.background = 'rgba(239, 68, 68, 0.15)';
        fetchModelsStatus.style.color = '#f87171';
        fetchModelsStatus.style.border = '1px solid rgba(239, 68, 68, 0.3)';
        fetchModelsStatus.textContent = `Uyarı: ${data.mesaj}`;
      }
    } catch (e) {
      btnFetchModels.disabled = false;
      btnFetchModels.textContent = '🔄 Modelleri Listele / Getir';
      fetchModelsStatus.style.display = 'block';
      fetchModelsStatus.style.background = 'rgba(239, 68, 68, 0.15)';
      fetchModelsStatus.style.color = '#f87171';
      fetchModelsStatus.style.border = '1px solid rgba(239, 68, 68, 0.3)';
      fetchModelsStatus.textContent = `Hata: ${e.message}`;
    }
  }

  async function testApiConnection() {
    btnTestConn.disabled = true;
    btnTestConn.textContent = '⏳ Test Ediliyor...';
    testConnResult.style.display = 'none';

    try {
      const resp = await fetch('/api/aselix/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          base_url: cfgBaseUrl.value.trim(),
          api_key: cfgApiKey.value.trim(),
          model: cfgModelSelect.value || 'aselix-llm-v1',
          messages: [{ role: 'user', content: 'Merhaba, bağlantı testi. 1 cümle ile teyit et.' }]
        })
      });

      const data = await resp.json();
      btnTestConn.disabled = false;
      btnTestConn.textContent = '⚡ API Testi Yap';

      testConnResult.style.display = 'block';
      if (data.durum === 'basarili') {
        testConnResult.style.background = 'rgba(16, 185, 129, 0.15)';
        testConnResult.style.border = '1px solid rgba(16, 185, 129, 0.4)';
        testConnResult.style.color = '#34d399';
        testConnResult.innerHTML = `<strong>✓ Bağlantı Başarılı!</strong><br><small>${escapeHtml(data.cevap)}</small>`;
      } else {
        testConnResult.style.background = 'rgba(239, 68, 68, 0.15)';
        testConnResult.style.border = '1px solid rgba(239, 68, 68, 0.4)';
        testConnResult.style.color = '#f87171';
        testConnResult.innerHTML = `<strong>⚠️ Bağlantı Yanıtı:</strong><br><small>${escapeHtml(data.mesaj || 'Ulaşılamadı')}</small>`;
      }
    } catch (e) {
      btnTestConn.disabled = false;
      btnTestConn.textContent = '⚡ API Testi Yap';
      testConnResult.style.display = 'block';
      testConnResult.style.background = 'rgba(239, 68, 68, 0.15)';
      testConnResult.style.border = '1px solid rgba(239, 68, 68, 0.4)';
      testConnResult.style.color = '#f87171';
      testConnResult.innerHTML = `<strong>⚠️ Test Hatası:</strong><br><small>${escapeHtml(e.message)}</small>`;
    }
  }

  // --------------------------------------------------------------------------
  // UGESUDKOKPİT Bakiye Sipariş (ud_kokpit -> teslimat_md.bakiye_siparis) & Faz Seçici
  // --------------------------------------------------------------------------
  async function loadBakiyeSiparisItems() {
    try {
      const resp = await fetch('/api/bakiye-siparis/items');
      const data = await resp.json();
      if (data && data.durum === 'basarili') {
        bakiyeState.allItems = Array.isArray(data.kayitlar) ? data.kayitlar : [];
        bakiyeState.projeler = Array.isArray(data.projeler) ? data.projeler : [];

        if (bakiyeDbSource1) {
          const isPg = String(data.kaynak || '').toLowerCase().includes('postgres');
          bakiyeDbSource1.textContent = isPg
            ? `🗄️ ud_kokpit • bakiye_siparis (${bakiyeState.allItems.length} kalem)`
            : `📂 bakiye_siparis (${bakiyeState.allItems.length} kalem)`;
        }

        if (homeStatBakiyeCount) {
          homeStatBakiyeCount.textContent = String(bakiyeState.allItems.length);
        }

        populateBakiyeProjectSelects();
        populateBulkBakiyeSelectors();
        populateBakiyeItemSelect(1, true);
        populateBakiyeItemSelect(2, true);
      }
    } catch (err) {
      console.warn('Bakiye sipariş verileri alınamadı:', err);
    }
  }

  function populateBulkBakiyeSelectors() {
    if (bulkTemplateProjeSelect) {
      bulkTemplateProjeSelect.innerHTML = `<option value="ALL">Tüm Projeler (${bakiyeState.projeler.length} proje / ${bakiyeState.allItems.length} kalem)</option>` +
        bakiyeState.projeler.map(p => {
          const cnt = bakiyeState.allItems.filter(x => x.proje_adi === p).length;
          return `<option value="${escapeHtml(p)}">${escapeHtml(p)} [${cnt} kalem]</option>`;
        }).join('');
    }
    if (bulkFallbackBakiyeSelect) {
      bulkFallbackBakiyeSelect.innerHTML = `<option value="">-- Excel İçindeki SD / Kalem Sütununu Kullan (Otomatik) --</option>` +
        bakiyeState.allItems.map(it => {
          const val = `${it.sd}|${it.kalem}`;
          const lbl = `${it.sd} / ${it.kalem} • ${it.proje_adi} • ${truncateStr(it.urun_kisa_metni || '', 38)}`;
          return `<option value="${escapeHtml(val)}">${escapeHtml(lbl)}</option>`;
        }).join('');
    }
  }

  function populateBakiyeProjectSelects() {
    [bakiyeProjeSelect1, bakiyeProjeSelect2].forEach(sel => {
      if (!sel) return;
      const prev = sel.value;
      sel.innerHTML = `<option value="">Tüm Projeler (${bakiyeState.projeler.length} proje / ${bakiyeState.allItems.length} kalem)</option>`;
      bakiyeState.projeler.forEach(p => {
        const count = bakiyeState.allItems.filter(x => x.proje_adi === p).length;
        const sample = bakiyeState.allItems.find(x => x.proje_adi === p && x.uges_proje_tanim);
        const ugesSuffix = sample && sample.uges_proje_tanim ? ` (${sample.uges_proje_tanim})` : '';
        const opt = document.createElement('option');
        opt.value = p;
        opt.textContent = `${p}${ugesSuffix} [${count} kalem]`;
        sel.appendChild(opt);
      });
      if (prev && bakiyeState.projeler.includes(prev)) {
        sel.value = prev;
      }
    });
  }

  function filterBakiyeItemsForDoc(docIndex) {
    const projeSel = docIndex === 1 ? bakiyeProjeSelect1 : bakiyeProjeSelect2;
    const searchIn = docIndex === 1 ? bakiyeSearchInput1 : bakiyeSearchInput2;
    const selectedProje = (projeSel && projeSel.value) ? projeSel.value.trim() : '';
    const query = (searchIn && searchIn.value) ? searchIn.value.trim().toLowerCase() : '';

    return bakiyeState.allItems.filter(item => {
      if (selectedProje && item.proje_adi !== selectedProje) return false;
      if (query) {
        const hay = `${item.proje_adi || ''} ${item.uges_proje_tanim || ''} ${item.sd || ''} ${item.kalem || ''} ${item.urun_kisa_metni || ''} ${item.malzeme || ''}`.toLowerCase();
        if (!hay.includes(query)) return false;
      }
      return true;
    });
  }

  function populateBakiyeItemSelect(docIndex, autoSelectFirst = false) {
    const itemSel = docIndex === 1 ? bakiyeItemSelect1 : bakiyeItemSelect2;
    if (!itemSel) return;

    const prevId = itemSel.value;
    const filtered = filterBakiyeItemsForDoc(docIndex);
    itemSel.innerHTML = '';

    if (filtered.length === 0) {
      const emptyOpt = document.createElement('option');
      emptyOpt.value = '';
      emptyOpt.textContent = 'Eşleşen sipariş kalemi bulunamadı';
      itemSel.appendChild(emptyOpt);
      if (docIndex === 1) bakiyeState.selectedItem1 = null;
      else bakiyeState.selectedItem2 = null;
      renderBakiyeSelectedInfo(docIndex);
      return;
    }

    filtered.forEach(item => {
      const opt = document.createElement('option');
      opt.value = item.id;
      opt.textContent = `${item.proje_adi || '-'} • SD: ${item.sd || '-'} / Kalem: ${item.kalem || '-'} • ${item.urun_kisa_metni || '-'}`;
      itemSel.appendChild(opt);
    });

    if (prevId && filtered.some(x => String(x.id) === String(prevId))) {
      itemSel.value = prevId;
    } else if (autoSelectFirst && filtered.length > 0) {
      itemSel.value = filtered[0].id;
    }

    onBakiyeItemChanged(docIndex);
  }

  function onBakiyeItemChanged(docIndex) {
    const itemSel = docIndex === 1 ? bakiyeItemSelect1 : bakiyeItemSelect2;
    if (!itemSel) return;
    const idVal = itemSel.value;
    const found = bakiyeState.allItems.find(x => String(x.id) === String(idVal)) || null;
    if (docIndex === 1) {
      bakiyeState.selectedItem1 = found;
    } else {
      bakiyeState.selectedItem2 = found;
    }
    renderBakiyeSelectedInfo(docIndex);
    if (currentAnalysis) {
      renderDbKpiPreviewStrip();
    }
  }

  function renderBakiyeSelectedInfo(docIndex) {
    const item = docIndex === 1 ? bakiyeState.selectedItem1 : bakiyeState.selectedItem2;
    const phase = docIndex === 1 ? bakiyeState.phase1 : bakiyeState.phase2;
    const elProje = docIndex === 1 ? infoProje1 : infoProje2;
    const elSdKalem = docIndex === 1 ? infoSdKalem1 : infoSdKalem2;
    const elUrun = docIndex === 1 ? infoUrun1 : infoUrun2;
    const badgeEl = docIndex === 1 ? dropBadge1 : dropBadge2;

    if (!item) {
      if (elProje) elProje.textContent = 'Seçilmedi';
      if (elSdKalem) elSdKalem.textContent = '- / -';
      if (elUrun) elUrun.textContent = '-';
      if (badgeEl) {
        badgeEl.textContent = `${docIndex}. Belge (${phase})`;
      }
      return;
    }

    const projeDisplay = item.uges_proje_tanim
      ? `${item.proje_adi} (${item.uges_proje_tanim})`
      : (item.proje_adi || '-');
    const sdKalemDisplay = `SD: ${item.sd || '-'} / Kalem: ${item.kalem || '-'}`;
    const urunDisplay = item.malzeme
      ? `${item.urun_kisa_metni || '-'} [${item.malzeme}]`
      : (item.urun_kisa_metni || '-');

    if (elProje) elProje.textContent = projeDisplay;
    if (elSdKalem) elSdKalem.textContent = sdKalemDisplay;
    if (elUrun) elUrun.textContent = urunDisplay;

    if (badgeEl) {
      badgeEl.textContent = `${docIndex}. Belge • ${item.proje_adi || ''} • SD:${item.sd}/${item.kalem} • ${phase}`;
    }
  }

  function setBakiyePhase(docIndex, phase) {
    if (!phase) return;
    if (docIndex === 1) {
      // Kural: İlk faaliyet hep Dry-Run olmalı
      bakiyeState.phase1 = 'Dry-Run';
    } else {
      // Kural: 2. faaliyet türü yalnızca FKT ya da SKT olmalı
      bakiyeState.phase2 = (phase === 'SKT') ? 'SKT' : 'FKT';
    }
    const groupEl = docIndex === 1 ? phaseGroup1 : phaseGroup2;
    const effectivePhase = docIndex === 1 ? 'Dry-Run' : bakiyeState.phase2;
    if (groupEl) {
      groupEl.querySelectorAll('.phase-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.phase === effectivePhase);
      });
    }
    renderBakiyeSelectedInfo(docIndex);
    if (currentAnalysis) {
      renderDbKpiPreviewStrip();
    }
  }

  function copyBakiyeSelection1To2() {
    if (bakiyeProjeSelect1 && bakiyeProjeSelect2) {
      bakiyeProjeSelect2.value = bakiyeProjeSelect1.value;
    }
    if (bakiyeSearchInput1 && bakiyeSearchInput2) {
      bakiyeSearchInput2.value = bakiyeSearchInput1.value;
    }
    populateBakiyeItemSelect(2, false);
    if (bakiyeItemSelect1 && bakiyeItemSelect2) {
      bakiyeItemSelect2.value = bakiyeItemSelect1.value;
      onBakiyeItemChanged(2);
    }
  }

  function applyAutoDetectedBakiyeFromExcel(docIndex, detectedBakiye, detectedFaz) {
    if (!detectedBakiye || (!detectedBakiye.sd && !detectedBakiye.kalem)) return;
    const sdStr = String(detectedBakiye.sd || '').trim();
    const kalemStr = String(detectedBakiye.kalem || '').trim();

    let matchedItem = bakiyeState.allItems.find(
      x => String(x.sd || '').trim() === sdStr && String(x.kalem || '').trim() === kalemStr
    );
    if (!matchedItem && sdStr && !kalemStr) {
      matchedItem = bakiyeState.allItems.find(x => String(x.sd || '').trim() === sdStr);
    }

    if (!matchedItem) {
      matchedItem = {
        id: detectedBakiye.id || `auto_${sdStr}_${kalemStr}`,
        proje_adi: detectedBakiye.proje_adi || 'Bilinmeyen Proje',
        uges_proje_tanim: detectedBakiye.uges_proje_tanim || '',
        sd: sdStr,
        kalem: kalemStr,
        sd_kalem: `${sdStr} / ${kalemStr}`,
        urun_kisa_metni: detectedBakiye.urun_kisa_metni || '-',
        malzeme: detectedBakiye.malzeme || ''
      };
      bakiyeState.allItems.push(matchedItem);
      if (matchedItem.proje_adi && !bakiyeState.projeler.includes(matchedItem.proje_adi)) {
        bakiyeState.projeler.push(matchedItem.proje_adi);
        populateBakiyeProjectFilters();
      }
    }

    const projeSel = docIndex === 1 ? bakiyeProjeSelect1 : bakiyeProjeSelect2;
    const searchIn = docIndex === 1 ? bakiyeSearchInput1 : bakiyeSearchInput2;
    const itemSel = docIndex === 1 ? bakiyeItemSelect1 : bakiyeItemSelect2;

    if (projeSel && matchedItem.proje_adi && bakiyeState.projeler.includes(matchedItem.proje_adi)) {
      projeSel.value = matchedItem.proje_adi;
    } else if (projeSel) {
      projeSel.value = '';
    }
    if (searchIn) searchIn.value = '';

    populateBakiyeItemSelect(docIndex, false);
    if (itemSel) {
      itemSel.value = matchedItem.id;
    }
    if (docIndex === 1) {
      bakiyeState.selectedItem1 = matchedItem;
    } else {
      bakiyeState.selectedItem2 = matchedItem;
      if (detectedFaz === 'FKT' || detectedFaz === 'SKT') {
        setBakiyePhase(2, detectedFaz);
      }
    }
    renderBakiyeSelectedInfo(docIndex);
  }

  function getBakiyeMetaPayload(docIndex) {
    const item = docIndex === 1 ? bakiyeState.selectedItem1 : bakiyeState.selectedItem2;
    const phase = docIndex === 1 ? bakiyeState.phase1 : bakiyeState.phase2;
    if (!item) {
      return { faz: phase };
    }
    return {
      id: item.id,
      proje_adi: item.proje_adi || '',
      uges_proje_tanim: item.uges_proje_tanim || '',
      sd: item.sd || '',
      kalem: item.kalem || '',
      sd_kalem: item.sd_kalem || `${item.sd || '-'} / ${item.kalem || '-'}`,
      urun_kisa_metni: item.urun_kisa_metni || '',
      malzeme: item.malzeme || '',
      faz: phase
    };
  }

  function formatShortMetaLabel(meta, fallbackName) {
    if (!meta || typeof meta !== 'object') return fallbackName;
    const faz = meta.faz || '';
    const proje = meta.proje_adi || '';
    const sd = meta.sd || '';
    const kalem = meta.kalem || '';
    if (proje && sd) {
      return `${proje} (SD:${sd}/${kalem}) [${faz}]`;
    }
    if (faz) {
      return `${fallbackName} [${faz}]`;
    }
    return fallbackName;
  }

  function renderComparisonMetaBanner(m1, m2, doc1Name, doc2Name) {
    if (!comparisonMetaBanner) return;
    const hasM1 = m1 && (m1.proje_adi || m1.sd || m1.faz);
    const hasM2 = m2 && (m2.proje_adi || m2.sd || m2.faz);
    if (!hasM1 && !hasM2) {
      comparisonMetaBanner.style.display = 'none';
      return;
    }

    const buildBox = (m, docName, idx, cls, phaseCls) => {
      const proje = m?.proje_adi ? (m.uges_proje_tanim ? `${m.proje_adi} (${m.uges_proje_tanim})` : m.proje_adi) : '-';
      const sdKalem = (m?.sd || m?.kalem) ? `${m.sd || '-'} / ${m.kalem || '-'}` : '-';
      const urun = m?.urun_kisa_metni ? (m.malzeme ? `${m.urun_kisa_metni} [${m.malzeme}]` : m.urun_kisa_metni) : '-';
      const faz = m?.faz || '-';
      return `
        <div class="meta-doc-box ${cls}">
          <div class="meta-doc-top">
            <span class="meta-doc-title">${idx}. Excel: ${escapeHtml(docName || (idx + '. Belge'))}</span>
            <span class="meta-phase-pill ${phaseCls}">Faz: ${escapeHtml(faz)}</span>
          </div>
          <div class="meta-doc-details">
            <span><strong>Proje İsmi:</strong> ${escapeHtml(proje)}</span>
            <span><strong>SD / Kalem:</strong> <code>${escapeHtml(sdKalem)}</code></span>
            <span><strong>Ürün Kısa Metni:</strong> ${escapeHtml(urun)}</span>
          </div>
        </div>
      `;
    };

    comparisonMetaBanner.innerHTML = `
      ${buildBox(m1, doc1Name, 1, 'doc-a', 'phase-a')}
      <div class="meta-vs-chip">${escapeHtml(m1?.faz || 'Dry-Run')} ↔ ${escapeHtml(m2?.faz || '2. Faz')}</div>
      ${buildBox(m2, doc2Name, 2, 'doc-b', 'phase-b')}
    `;
    comparisonMetaBanner.style.display = 'grid';
  }

  function buildCurrentKpiExportPayload() {
    const v = (currentAnalysis && currentAnalysis.venn_verileri) ? currentAnalysis.venn_verileri : {};
    const m1 = getBakiyeMetaPayload(1);
    const m2 = getBakiyeMetaPayload(2);

    // Anahtar bilgiler için öncelikle 1. Belge seçimi, yoksa 2. Belge seçimi kullanılır
    const projeAdi = m1.proje_adi || m2.proje_adi || 'Belirtilmedi';
    const sd = m1.sd || m2.sd || '-';
    const kalem = m1.kalem || m2.kalem || '-';
    const urunKisaMetni = m1.urun_kisa_metni || m2.urun_kisa_metni || '-';
    const ugesProjeTanim = m1.uges_proje_tanim || m2.uges_proje_tanim || '';
    const malzeme = m1.malzeme || m2.malzeme || '';

    // 1. Faaliyetin Türü her zaman Dry-Run, 2. Faaliyetin Türü seçilen 2. Excel fazı (FKT / SKT)
    const faaliyet1Turu = 'Dry-Run';
    const faaliyet2Turu = m2.faz || bakiyeState.phase2 || 'FKT';

    return {
      proje_adi: projeAdi,
      sd: sd,
      kalem: kalem,
      faaliyet_1_turu: faaliyet1Turu,
      faaliyet_2_turu: faaliyet2Turu,
      urun_kisa_metni: urunKisaMetni,
      uges_proje_tanim: ugesProjeTanim,
      malzeme: malzeme,
      ilk_faaliyet_benzersiz_uygunsuzluklar: v.only_a_items || [],
      ortak_uygunsuzluklar: (v.matched_pairs || []).map(p => ({
        ilk_faaliyet_uygunsuzluk: p.item_a,
        ikinci_faaliyet_uygunsuzluk: p.item_b,
        benzerlik_yuzdesi: p.similarity,
        eslesme_tipi: p.match_type || ''
      })),
      ikinci_faaliyet_benzersiz_uygunsuzluklar: v.only_b_items || [],
      reddedilen_eslesmeler: (v.declined_pairs || []).map(p => ({
        ilk_faaliyet_uygunsuzluk: p.item_a,
        ikinci_faaliyet_uygunsuzluk: p.item_b,
        eski_benzerlik_yuzdesi: p.similarity
      }))
    };
  }

  function renderDbKpiPreviewStrip() {
    if (!dbKpiPreviewStrip || !currentAnalysis) return;
    const p = buildCurrentKpiExportPayload();
    const onlyACount = p.ilk_faaliyet_benzersiz_uygunsuzluklar.length;
    const commonCount = p.ortak_uygunsuzluklar.length;
    const onlyBCount = p.ikinci_faaliyet_benzersiz_uygunsuzluklar.length;

    dbKpiPreviewStrip.innerHTML = `
      <div class="kpi-strip-cell">
        <span class="kpi-strip-label">Proje Adı</span>
        <span class="kpi-strip-val">${escapeHtml(p.proje_adi)}</span>
      </div>
      <div class="kpi-strip-cell">
        <span class="kpi-strip-label">SD / Kalem</span>
        <span class="kpi-strip-val mono">${escapeHtml(p.sd)} / ${escapeHtml(p.kalem)}</span>
      </div>
      <div class="kpi-strip-cell">
        <span class="kpi-strip-label">1. Faaliyetin Türü</span>
        <span class="kpi-strip-val mono">🔒 ${escapeHtml(p.faaliyet_1_turu)}</span>
      </div>
      <div class="kpi-strip-cell">
        <span class="kpi-strip-label">2. Faaliyetin Türü</span>
        <span class="kpi-strip-val mono">${escapeHtml(p.faaliyet_2_turu)}</span>
      </div>
      <div class="kpi-strip-cell">
        <span class="kpi-strip-label">Ürün Kısa Metni</span>
        <span class="kpi-strip-val">${escapeHtml(p.urun_kisa_metni)}</span>
      </div>
      <div class="kpi-strip-cell">
        <span class="kpi-strip-label">İlk Faaliyet Benzersiz (Dry-Run)</span>
        <span class="kpi-strip-val mono" style="color: var(--color-primary);">${onlyACount} adet (JSON)</span>
      </div>
      <div class="kpi-strip-cell">
        <span class="kpi-strip-label">Ortak Uygunsuzluklar</span>
        <span class="kpi-strip-val mono" style="color: var(--color-emerald);">${commonCount} adet (JSON)</span>
      </div>
      <div class="kpi-strip-cell">
        <span class="kpi-strip-label">2. Faaliyet Benzersiz (${escapeHtml(p.faaliyet_2_turu)})</span>
        <span class="kpi-strip-val mono" style="color: var(--color-purple);">${onlyBCount} adet (JSON)</span>
      </div>
    `;
  }

  async function saveCurrentDataToKpiDatabase() {
    if (!currentAnalysis || !currentAnalysis.venn_verileri) {
      alert('Önce bir kıyaslama çalıştırın.');
      return;
    }

    const payload = buildCurrentKpiExportPayload();
    if (btnSaveToKpiDb) {
      btnSaveToKpiDb.disabled = true;
      btnSaveToKpiDb.textContent = '⏳ Veritabanına Kaydediliyor...';
    }

    try {
      const resp = await fetch('/api/bakiye-siparis/save-kpi', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await resp.json();

      if (btnSaveToKpiDb) {
        btnSaveToKpiDb.disabled = false;
        btnSaveToKpiDb.textContent = '💾 Mevcut Verileri Veritabanına Aktar';
      }

      if (kpiSaveStatusBanner) {
        kpiSaveStatusBanner.style.display = 'block';
        if (data.durum === 'basarili' || data.durum === 'uyari') {
          kpiSaveStatusBanner.className = 'db-kpi-status-msg ok';
          kpiSaveStatusBanner.innerHTML = `
            ✓ <strong>${escapeHtml(data.mesaj)}</strong>
            (İlk Faaliyet [Dry-Run] Benzersiz: <strong>${payload.ilk_faaliyet_benzersiz_uygunsuzluklar.length}</strong>,
            Ortak: <strong>${payload.ortak_uygunsuzluklar.length}</strong>,
            2. Faaliyet [${escapeHtml(payload.faaliyet_2_turu)}] Benzersiz: <strong>${payload.ikinci_faaliyet_benzersiz_uygunsuzluklar.length}</strong> •
            JSON listeleri kaydedildi.)
          `;
          refreshHomePageStats();
        } else {
          kpiSaveStatusBanner.className = 'db-kpi-status-msg err';
          kpiSaveStatusBanner.textContent = '❌ Kayıt hatası: ' + (data.mesaj || 'Bilinmeyen hata');
        }
      }
    } catch (err) {
      if (btnSaveToKpiDb) {
        btnSaveToKpiDb.disabled = false;
        btnSaveToKpiDb.textContent = '💾 Mevcut Verileri Veritabanına Aktar';
      }
      if (kpiSaveStatusBanner) {
        kpiSaveStatusBanner.style.display = 'block';
        kpiSaveStatusBanner.className = 'db-kpi-status-msg err';
        kpiSaveStatusBanner.textContent = '❌ Sunucu iletişim hatası: ' + err.message;
      }
    }
  }

  let kpiLoadedRecordsCache = [];

  function openKpiRecordsModal() {
    if (!kpiRecordsModal) return;
    if (kpiModalJsonPreviewBox) kpiModalJsonPreviewBox.style.display = 'none';
    kpiRecordsModal.style.display = 'flex';
    loadAndRenderKpiRecords();
  }

  function closeKpiRecordsModal() {
    if (!kpiRecordsModal) return;
    kpiRecordsModal.style.display = 'none';
  }

  function showKpiJsonPreviewInModal(titleText, jsonString) {
    if (!kpiModalJsonPreviewBox || !kpiModalJsonContent) return;
    if (kpiModalJsonTitle) kpiModalJsonTitle.textContent = `📄 ${titleText}`;
    kpiModalJsonContent.textContent = jsonString;
    kpiModalJsonPreviewBox.style.display = 'block';
  }

  function clearKpiTableFilters() {
    if (kpiFilterProje) kpiFilterProje.value = 'ALL';
    if (kpiFilterSd) kpiFilterSd.value = '';
    if (kpiFilterKalem) kpiFilterKalem.value = '';
    if (kpiFilterFaaliyet2) kpiFilterFaaliyet2.value = 'ALL';
    if (kpiFilterUrun) kpiFilterUrun.value = '';
    if (kpiFilterSearch) kpiFilterSearch.value = '';
    renderFilteredKpiRecordsTable();
  }

  async function loadAndRenderKpiRecords() {
    if (!kpiSavedRecordsBody) return;
    kpiSavedRecordsBody.innerHTML = `<tr><td colspan="11" style="text-align:center; padding:18px;">Veritabanından okunuyor...</td></tr>`;
    try {
      const resp = await fetch('/api/bakiye-siparis/kpi-records?limit=200');
      const data = await resp.json();
      kpiLoadedRecordsCache = data.kayitlar || [];

      if (kpiFilterProje) {
        const prevVal = kpiFilterProje.value || 'ALL';
        const distinctProj = Array.isArray(data.projeler) && data.projeler.length > 0
          ? data.projeler
          : [...new Set(kpiLoadedRecordsCache.map(r => r.proje_adi).filter(Boolean))].sort();
        kpiFilterProje.innerHTML = `<option value="ALL">Tüm Projeler (${distinctProj.length})</option>` +
          distinctProj.map(p => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`).join('');
        if ([...kpiFilterProje.options].some(o => o.value === prevVal)) {
          kpiFilterProje.value = prevVal;
        }
      }

      renderFilteredKpiRecordsTable();
    } catch (err) {
      kpiSavedRecordsBody.innerHTML = `<tr><td colspan="11" style="text-align:center; padding:18px; color:#dc2626;">Hata: ${escapeHtml(err.message)}</td></tr>`;
    }
  }

  function renderFilteredKpiRecordsTable() {
    if (!kpiSavedRecordsBody) return;
    const isAdmin = isCurrentUserAdmin();

    const fProje = kpiFilterProje ? kpiFilterProje.value : 'ALL';
    const fSd = kpiFilterSd ? kpiFilterSd.value.toLowerCase().trim() : '';
    const fKalem = kpiFilterKalem ? kpiFilterKalem.value.toLowerCase().trim() : '';
    const fF2 = kpiFilterFaaliyet2 ? kpiFilterFaaliyet2.value : 'ALL';
    const fUrun = kpiFilterUrun ? kpiFilterUrun.value.toLowerCase().trim() : '';
    const fSearch = kpiFilterSearch ? kpiFilterSearch.value.toLowerCase().trim() : '';

    const filtered = kpiLoadedRecordsCache
      .map((r, originalIdx) => ({ r, originalIdx }))
      .filter(({ r }) => {
        if (fProje && fProje !== 'ALL' && String(r.proje_adi || '') !== fProje) return false;
        if (fSd && !String(r.sd || '').toLowerCase().includes(fSd)) return false;
        if (fKalem && !String(r.kalem || '').toLowerCase().includes(fKalem)) return false;
        if (fF2 && fF2 !== 'ALL' && String(r.faaliyet_2_turu || '').toUpperCase() !== fF2.toUpperCase()) return false;
        if (fUrun && !String(r.urun_kisa_metni || '').toLowerCase().includes(fUrun)) return false;
        if (fSearch) {
          const haystack = JSON.stringify({
            proje_adi: r.proje_adi,
            sd: r.sd,
            kalem: r.kalem,
            urun_kisa_metni: r.urun_kisa_metni,
            kaydeden_kullanici: r.kaydeden_kullanici,
            ilk_faaliyet_benzersiz_uygunsuzluklar: r.ilk_faaliyet_benzersiz_uygunsuzluklar,
            ortak_uygunsuzluklar: r.ortak_uygunsuzluklar,
            ikinci_faaliyet_benzersiz_uygunsuzluklar: r.ikinci_faaliyet_benzersiz_uygunsuzluklar
          }).toLowerCase();
          if (!haystack.includes(fSearch)) return false;
        }
        return true;
      });

    if (kpiRecordsCountBadge) {
      kpiRecordsCountBadge.textContent = `${filtered.length} / ${kpiLoadedRecordsCache.length} Kayıt`;
    }

    if (filtered.length === 0) {
      kpiSavedRecordsBody.innerHTML = `<tr><td colspan="11" style="text-align:center; padding:20px; color:var(--text-muted);">Filtre kriterlerine uygun kayıt bulunamadı.</td></tr>`;
      return;
    }

    kpiSavedRecordsBody.innerHTML = filtered.map(({ r, originalIdx }) => {
      const jsonPreviewObj = {
        ilk_faaliyet_benzersiz_uygunsuzluklar: r.ilk_faaliyet_benzersiz_uygunsuzluklar,
        ortak_uygunsuzluklar: r.ortak_uygunsuzluklar,
        ikinci_faaliyet_benzersiz_uygunsuzluklar: r.ikinci_faaliyet_benzersiz_uygunsuzluklar,
        reddedilen_eslesmeler: r.reddedilen_eslesmeler
      };
      const encodedJson = encodeURIComponent(JSON.stringify(jsonPreviewObj, null, 2));
      const recTitle = `${r.proje_adi} • SD: ${r.sd}/${r.kalem} (${r.faaliyet_1_turu} vs ${r.faaliyet_2_turu})`;
      const dateStr = r.guncellenme_zamani ? String(r.guncellenme_zamani).slice(0, 16) : '';
      const adminDelCell = `
        <td class="admin-only-item" style="text-align: right; ${isAdmin ? '' : 'display: none;'}">
          <button type="button" class="btn-delete-kpi-record" data-rec-idx="${originalIdx}" title="Admin Yetkisi: Bu kaydı veritabanından kalıcı olarak sil">
            🗑️ Sil
          </button>
        </td>
      `;

      return `
        <tr>
          <td><strong>${escapeHtml(r.proje_adi)}</strong></td>
          <td style="font-family: var(--font-mono);">${escapeHtml(r.sd)} / ${escapeHtml(r.kalem)}</td>
          <td><span class="badge-similarity-only a">${escapeHtml(r.faaliyet_1_turu)}</span></td>
          <td><span class="badge-similarity-only b">${escapeHtml(r.faaliyet_2_turu)}</span></td>
          <td>${escapeHtml(r.urun_kisa_metni)}</td>
          <td style="font-family: var(--font-mono); font-weight:700;">${r.ilk_faaliyet_benzersiz_sayisi}</td>
          <td style="font-family: var(--font-mono); font-weight:700; color: var(--color-emerald);">${r.ortak_uygunsuzluk_sayisi}</td>
          <td style="font-family: var(--font-mono); font-weight:700;">${r.ikinci_faaliyet_benzersiz_sayisi}</td>
          <td style="font-size: 11.5px; color: var(--text-muted);">
            <div>${escapeHtml(r.kaydeden_kullanici || '-')}</div>
            <div style="font-family: var(--font-mono); font-size: 10.5px;">${escapeHtml(dateStr)}</div>
          </td>
          <td>
            <button type="button" class="btn-view-json" data-rec-title="${escapeHtml(recTitle)}" data-encoded-json="${encodedJson}">
              { } JSON Gör
            </button>
          </td>
          ${adminDelCell}
        </tr>
      `;
    }).join('');
  }

  async function deleteKpiRecordByIndex(recIdx) {
    if (!isCurrentUserAdmin()) {
      alert('Kayıt silme işlemi yalnızca Admin (Yönetici) yetkisine sahip kullanıcılar tarafından yapılabilir.');
      return;
    }
    const r = kpiLoadedRecordsCache[recIdx];
    if (!r) return;

    const confirmMsg = `Aşağıdaki KPI kaydını veritabanından (teslimat_md.dryrun_fkt_skt_kpi) silmek istediğinize emin misiniz?\n\n` +
      `• Proje Adı: ${r.proje_adi}\n` +
      `• SD / Kalem: ${r.sd} / ${r.kalem}\n` +
      `• Faaliyetler: ${r.faaliyet_1_turu} vs ${r.faaliyet_2_turu}\n` +
      `• Ürün Kısa Metni: ${r.urun_kisa_metni}`;

    if (!window.confirm(confirmMsg)) return;

    try {
      const resp = await fetch('/api/bakiye-siparis/delete-kpi', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          proje_adi: r.proje_adi,
          sd: r.sd,
          kalem: r.kalem,
          faaliyet_1_turu: r.faaliyet_1_turu,
          faaliyet_2_turu: r.faaliyet_2_turu,
          urun_kisa_metni: r.urun_kisa_metni
        })
      });
      const data = await resp.json();

      if (kpiSaveStatusBanner) {
        kpiSaveStatusBanner.style.display = 'block';
        if (resp.ok && data.durum === 'basarili') {
          kpiSaveStatusBanner.className = 'db-kpi-status-msg ok';
          kpiSaveStatusBanner.innerHTML = `🗑️ <strong>${escapeHtml(data.mesaj)}</strong>`;
        } else {
          kpiSaveStatusBanner.className = 'db-kpi-status-msg err';
          kpiSaveStatusBanner.textContent = '❌ Silme hatası: ' + (data.mesaj || 'İşlem başarısız');
        }
      }

      await loadAndRenderKpiRecords();
      refreshHomePageStats();
    } catch (err) {
      alert('Sunucu iletişim hatası: ' + err.message);
    }
  }

  // --------------------------------------------------------------------------
  // Toplu Veri Girişi (Şablon Excel İndir / Yükle / SD-Kalem Arka Plan Eşleştirme)
  // --------------------------------------------------------------------------
  function downloadBulkTemplateExcel() {
    const selectedProje = bulkTemplateProjeSelect ? bulkTemplateProjeSelect.value : 'ALL';
    const q = (selectedProje && selectedProje !== 'ALL') ? `?proje=${encodeURIComponent(selectedProje)}` : '';
    window.location.href = `/api/bulk/template-excel${q}`;
  }

  function handleBulkFileSelected(file) {
    if (!file || !file.name.match(/\.(xlsx|xls|xlsm)$/i)) {
      alert('Lütfen geçerli bir Excel (.xlsx, .xls) şablon dosyası seçin.');
      return;
    }
    bulkFile = file;
    if (bulkFileName) bulkFileName.textContent = file.name;
    if (bulkFileSub) bulkFileSub.textContent = `${formatBytes(file.size)} • Analize Hazır`;
    if (bulkDropzoneContent) bulkDropzoneContent.style.display = 'none';
    if (bulkFileView) bulkFileView.style.display = 'flex';
  }

  function removeBulkFile() {
    bulkFile = null;
    if (bulkFileInput) bulkFileInput.value = '';
    if (bulkDropzoneContent) bulkDropzoneContent.style.display = 'flex';
    if (bulkFileView) bulkFileView.style.display = 'none';
  }

  async function runBulkExcelAnalysis() {
    if (btnRunBulkAnalysis) {
      btnRunBulkAnalysis.disabled = true;
      btnRunBulkAnalysis.textContent = '⏳ SD/Kalem Eşleştiriliyor ve Venn Hesaplanıyor...';
    }

    try {
      let fileToUpload = bulkFile;
      // Kullanıcı henüz dosya seçmediyse örnek şablonu arka planda indirip onunla çalıştır
      if (!fileToUpload) {
        const selectedProje = bulkTemplateProjeSelect ? bulkTemplateProjeSelect.value : 'ALL';
        const q = (selectedProje && selectedProje !== 'ALL') ? `?proje=${encodeURIComponent(selectedProje)}` : '';
        const tplResp = await fetch(`/api/bulk/template-excel${q}`);
        const blob = await tplResp.blob();
        fileToUpload = new File([blob], 'VennFlow_Ornek_Sablon.xlsx', {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        handleBulkFileSelected(fileToUpload);
      }

      const fd = new FormData();
      fd.append('file', fileToUpload);
      fd.append('engine_mode', appConfig.engine === 'embedding' ? 'embedding' : 'local');
      fd.append('threshold', String((appConfig.threshold || 80) / 100.0));
      fd.append('base_url', appConfig.baseUrl || '');
      fd.append('api_key', (appConfig.apiKey || '').trim() || (kokpitIdentity.token || '').trim());
      fd.append('embedding_model', appConfig.embeddingModel || 'text-embedding-3-small');

      if (bulkFallbackBakiyeSelect && bulkFallbackBakiyeSelect.value && bulkFallbackBakiyeSelect.value.includes('|')) {
        const [defSd, defKalem] = bulkFallbackBakiyeSelect.value.split('|');
        fd.append('default_sd', defSd || '');
        fd.append('default_kalem', defKalem || '');
      }

      const resp = await fetch('/api/bulk/analyze-excel', {
        method: 'POST',
        body: fd
      });
      const data = await resp.json();

      if (btnRunBulkAnalysis) {
        btnRunBulkAnalysis.disabled = false;
        btnRunBulkAnalysis.textContent = '🚀 Şablonu Analiz Et & SD/Kalem Eşleştir';
      }

      if (data.durum !== 'basarili') {
        alert('Toplu Analiz Hatası: ' + (data.mesaj || 'İşlem başarısız.'));
        return;
      }

      bulkAnalyzedGroups = Array.isArray(data.gruplar) ? data.gruplar : [];
      if (bulkResultsSection) {
        bulkResultsSection.style.display = 'block';
        bulkResultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
      if (bulkResultsSubtitle) {
        bulkResultsSubtitle.innerHTML = `
          <strong>${escapeHtml(data.dosya_adi)}</strong> dosyasından <strong>${data.toplam_okunan_satir}</strong> uygunsuzluk satırı okundu ve
          <code>teslimat_md.bakiye_siparis</code> ile arka planda eşleştirilerek <strong>${bulkAnalyzedGroups.length}</strong> adet
          <strong>SD / Kalem Faaliyet Kıyaslama Grubu</strong> oluşturuldu.
        `;
      }
      if (bulkSaveStatusBanner) bulkSaveStatusBanner.style.display = 'none';
      renderBulkGroupsResults();
    } catch (err) {
      if (btnRunBulkAnalysis) {
        btnRunBulkAnalysis.disabled = false;
        btnRunBulkAnalysis.textContent = '🚀 Şablonu Analiz Et & SD/Kalem Eşleştir';
      }
      alert('Sunucu iletişim hatası: ' + err.message);
    }
  }

  function recalculateBulkGroupMetrics(g) {
    const v = g.venn_verileri;
    if (!v) return;
    v.intersection_count = (v.matched_pairs || []).length;
    v.only_a_count = (v.only_a_items || []).length;
    v.only_b_count = (v.only_b_items || []).length;
    v.total_unique = v.intersection_count + v.only_a_count + v.only_b_count;
    v.jaccard_similarity = v.total_unique > 0 ? (v.intersection_count / v.total_unique) : 0;
  }

  function renderBulkGroupsResults() {
    if (!bulkGroupsContainer) return;
    if (!bulkAnalyzedGroups.length) {
      bulkGroupsContainer.innerHTML = `<div class="panel-card" style="padding:20px; text-align:center;">Grup bulunamadı.</div>`;
      return;
    }

    bulkGroupsContainer.innerHTML = bulkAnalyzedGroups.map((g, gIdx) => {
      const v = g.venn_verileri || {};
      const matchedPairs = v.matched_pairs || [];
      const declinedPairs = v.declined_pairs || [];
      const onlyA = v.only_a_items || [];
      const onlyB = v.only_b_items || [];

      const matchRowsHtml = matchedPairs.map((p, pIdx) => `
        <tr>
          <td>${escapeHtml(p.item_a)}</td>
          <td>${escapeHtml(p.item_b)}</td>
          <td><span class="${p.similarity >= 99.5 ? 'badge-similarity-exact' : 'badge-similarity-fuzzy'}">%${p.similarity} Benzer</span></td>
          <td style="text-align:right;">
            <button type="button" class="btn-decline-match" data-bulk-action="decline" data-g-idx="${gIdx}" data-p-idx="${pIdx}">
              ✕ Benzer Değil
            </button>
          </td>
        </tr>
      `).join('');

      const declinedRowsHtml = declinedPairs.map((p, dIdx) => `
        <tr style="opacity: 0.78;">
          <td>${escapeHtml(p.item_a)}</td>
          <td>${escapeHtml(p.item_b)}</td>
          <td><span class="badge-similarity-declined">Reddedildi (%${p.similarity})</span></td>
          <td style="text-align:right;">
            <button type="button" class="btn-restore-match" data-bulk-action="restore" data-g-idx="${gIdx}" data-d-idx="${dIdx}">
              ↩ Geri Al
            </button>
          </td>
        </tr>
      `).join('');

      const onlyARowsHtml = onlyA.map(item => `
        <tr>
          <td>${escapeHtml(item)}</td>
          <td style="color:var(--text-dim);">-</td>
          <td><span class="badge-similarity-only a">Yalnızca 1. (${escapeHtml(g.faaliyet_1_turu)})</span></td>
          <td style="text-align:right; color:var(--text-dim);">—</td>
        </tr>
      `).join('');

      const onlyBRowsHtml = onlyB.map(item => `
        <tr>
          <td style="color:var(--text-dim);">-</td>
          <td>${escapeHtml(item)}</td>
          <td><span class="badge-similarity-only b">Yalnızca 2. (${escapeHtml(g.faaliyet_2_turu)})</span></td>
          <td style="text-align:right; color:var(--text-dim);">—</td>
        </tr>
      `).join('');

      return `
        <div class="panel-card bulk-group-card">
          <div class="bulk-group-header">
            <div class="bulk-group-meta">
              <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                <span class="badge-accent">${escapeHtml(g.proje_adi)}</span>
                <span class="bakiye-db-tag">SD: ${escapeHtml(g.sd)} / Kalem: ${escapeHtml(g.kalem)}</span>
                <span class="badge-similarity-only a">🔒 1. Faaliyet: ${escapeHtml(g.faaliyet_1_turu)}</span>
                <span class="badge-similarity-only b">2. Faaliyet: ${escapeHtml(g.faaliyet_2_turu)}</span>
                ${g.bakiye_eslesti ? '<span class="badge-similarity-exact">✓ Bakiye Sipariş Eşleşti</span>' : '<span class="badge-similarity-fuzzy">ℹ️ Manuel SD/Kalem</span>'}
              </div>
              <h4 style="margin:6px 0 2px; font-size:15px;">Ürün Kısa Metni: ${escapeHtml(g.urun_kisa_metni)}</h4>
            </div>

            <div class="bulk-group-kpis">
              <div class="bulk-kpi-pill">
                <span>1. Benzersiz (Dry-Run):</span>
                <strong>${v.only_a_count || 0}</strong>
              </div>
              <div class="bulk-kpi-pill common">
                <span>Ortak Uygunsuzluk:</span>
                <strong>${v.intersection_count || 0}</strong>
              </div>
              <div class="bulk-kpi-pill">
                <span>2. Benzersiz (${escapeHtml(g.faaliyet_2_turu)}):</span>
                <strong>${v.only_b_count || 0}</strong>
              </div>
              <button type="button" class="btn btn-secondary btn-sm" data-bulk-action="open-vennflow" data-g-idx="${gIdx}" title="Bu SD/Kalem grubunu VennFlow sekmesinde Venn şeması ile aç">
                🔵 VennFlow'da Aç
              </button>
              <button type="button" class="btn btn-primary btn-sm" data-bulk-action="save-single" data-g-idx="${gIdx}">
                💾 Bu Grubu Kaydet
              </button>
            </div>
          </div>

          <div class="table-responsive" style="max-height: 260px; margin-top: 10px;">
            <table class="data-table">
              <thead>
                <tr>
                  <th>1. Faaliyet (Dry-Run) Uygunsuzluğu</th>
                  <th>2. Faaliyet (${escapeHtml(g.faaliyet_2_turu)}) Uygunsuzluğu</th>
                  <th style="width: 165px;">Durum / Benzerlik</th>
                  <th style="width: 155px; text-align:right;">Karar</th>
                </tr>
              </thead>
              <tbody>
                ${matchRowsHtml}
                ${declinedRowsHtml}
                ${onlyARowsHtml}
                ${onlyBRowsHtml}
              </tbody>
            </table>
          </div>
        </div>
      `;
    }).join('');
  }

  function handleBulkGroupsContainerClick(e) {
    const btn = e.target.closest('[data-bulk-action]');
    if (!btn) return;
    const action = btn.dataset.bulkAction;
    const gIdx = parseInt(btn.dataset.gIdx, 10);
    if (isNaN(gIdx) || !bulkAnalyzedGroups[gIdx]) return;
    const g = bulkAnalyzedGroups[gIdx];
    const v = g.venn_verileri;

    if (action === 'decline') {
      const pIdx = parseInt(btn.dataset.pIdx, 10);
      if (isNaN(pIdx) || !v.matched_pairs || !v.matched_pairs[pIdx]) return;
      const rem = v.matched_pairs.splice(pIdx, 1)[0];
      if (!Array.isArray(v.declined_pairs)) v.declined_pairs = [];
      v.declined_pairs.push(rem);
      if (rem.item_a && !v.only_a_items.includes(rem.item_a)) v.only_a_items.unshift(rem.item_a);
      if (rem.item_b && !v.only_b_items.includes(rem.item_b)) v.only_b_items.unshift(rem.item_b);
      recalculateBulkGroupMetrics(g);
      renderBulkGroupsResults();
    } else if (action === 'restore') {
      const dIdx = parseInt(btn.dataset.dIdx, 10);
      if (isNaN(dIdx) || !v.declined_pairs || !v.declined_pairs[dIdx]) return;
      const rest = v.declined_pairs.splice(dIdx, 1)[0];
      v.matched_pairs.push(rest);
      v.only_a_items = (v.only_a_items || []).filter(x => x !== rest.item_a);
      v.only_b_items = (v.only_b_items || []).filter(x => x !== rest.item_b);
      recalculateBulkGroupMetrics(g);
      renderBulkGroupsResults();
    } else if (action === 'open-vennflow') {
      const meta1 = {
        proje_adi: g.proje_adi,
        sd: g.sd,
        kalem: g.kalem,
        urun_kisa_metni: g.urun_kisa_metni,
        malzeme: g.malzeme || '',
        uges_proje_tanim: g.uges_proje_tanim || '',
        faz: 'Dry-Run'
      };
      const meta2 = { ...meta1, faz: g.faaliyet_2_turu };
      const syntheticData = {
        durum: 'basarili',
        engine_mode: g.engine_mode || 'local',
        dosya_1_adi: `${g.sd}/${g.kalem} (Dry-Run)`,
        dosya_2_adi: `${g.sd}/${g.kalem} (${g.faaliyet_2_turu})`,
        belge_1_meta: meta1,
        belge_2_meta: meta2,
        esik_orani: g.esik_orani || 0.8,
        ozet_yorum: `${g.proje_adi} • SD: ${g.sd} / Kalem: ${g.kalem} (${g.urun_kisa_metni}) için Dry-Run ve ${g.faaliyet_2_turu} faaliyetleri kıyaslandı.`,
        venn_verileri: g.venn_verileri
      };
      switchMainPageTab('vennflow');
      renderAnalysisResults(syntheticData);
    } else if (action === 'save-single') {
      saveSingleBulkGroupToDb(g);
    }
  }

  function buildKpiPayloadFromBulkGroup(g) {
    const v = g.venn_verileri || {};
    return {
      proje_adi: g.proje_adi,
      sd: g.sd,
      kalem: g.kalem,
      faaliyet_1_turu: 'Dry-Run',
      faaliyet_2_turu: g.faaliyet_2_turu,
      urun_kisa_metni: g.urun_kisa_metni,
      uges_proje_tanim: g.uges_proje_tanim || '',
      malzeme: g.malzeme || '',
      ilk_faaliyet_benzersiz_uygunsuzluklar: v.only_a_items || [],
      ortak_uygunsuzluklar: (v.matched_pairs || []).map(p => ({
        ilk_faaliyet_uygunsuzluk: p.item_a,
        ikinci_faaliyet_uygunsuzluk: p.item_b,
        benzerlik_yuzdesi: p.similarity,
        eslesme_tipi: p.match_type || ''
      })),
      ikinci_faaliyet_benzersiz_uygunsuzluklar: v.only_b_items || [],
      reddedilen_eslesmeler: (v.declined_pairs || []).map(p => ({
        ilk_faaliyet_uygunsuzluk: p.item_a,
        ikinci_faaliyet_uygunsuzluk: p.item_b,
        eski_benzerlik_yuzdesi: p.similarity
      }))
    };
  }

  async function saveSingleBulkGroupToDb(g) {
    try {
      const payload = buildKpiPayloadFromBulkGroup(g);
      const resp = await fetch('/api/bakiye-siparis/save-kpi', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await resp.json();
      if (bulkSaveStatusBanner) {
        bulkSaveStatusBanner.style.display = 'block';
        bulkSaveStatusBanner.className = (data.durum === 'basarili' || data.durum === 'uyari') ? 'db-kpi-status-msg ok' : 'db-kpi-status-msg err';
        bulkSaveStatusBanner.innerHTML = `✓ <strong>${escapeHtml(data.mesaj || 'Kaydedildi')}</strong>`;
      }
      refreshHomePageStats();
    } catch (err) {
      alert('Kayıt hatası: ' + err.message);
    }
  }

  async function saveAllBulkGroupsToDb() {
    if (!bulkAnalyzedGroups.length) {
      alert('Kaydedilecek analiz grubu bulunamadı.');
      return;
    }
    if (btnSaveAllBulkGroups) {
      btnSaveAllBulkGroups.disabled = true;
      btnSaveAllBulkGroups.textContent = '⏳ Tüm Gruplar Kaydediliyor...';
    }
    try {
      const payloads = bulkAnalyzedGroups.map(buildKpiPayloadFromBulkGroup);
      const resp = await fetch('/api/bulk/save-all-kpi', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ gruplar: payloads })
      });
      const data = await resp.json();
      if (btnSaveAllBulkGroups) {
        btnSaveAllBulkGroups.disabled = false;
        btnSaveAllBulkGroups.innerHTML = '💾 Tüm Grupları Veritabanına Aktar (<code>dryrun_fkt_skt_kpi</code>)';
      }
      if (bulkSaveStatusBanner) {
        bulkSaveStatusBanner.style.display = 'block';
        bulkSaveStatusBanner.className = data.durum === 'basarili' ? 'db-kpi-status-msg ok' : 'db-kpi-status-msg err';
        bulkSaveStatusBanner.innerHTML = `✓ <strong>${escapeHtml(data.mesaj || '')}</strong>`;
      }
      refreshHomePageStats();
    } catch (err) {
      if (btnSaveAllBulkGroups) {
        btnSaveAllBulkGroups.disabled = false;
        btnSaveAllBulkGroups.innerHTML = '💾 Tüm Grupları Veritabanına Aktar (<code>dryrun_fkt_skt_kpi</code>)';
      }
      alert('Toplu kayıt hatası: ' + err.message);
    }
  }

  // --------------------------------------------------------------------------
  // Utility Functions
  // --------------------------------------------------------------------------
  function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  function csvEscape(text) {
    if (!text) return '';
    return String(text).replace(/"/g, '""');
  }

  function truncateStr(str, n) {
    return (str && str.length > n) ? str.substr(0, n - 1) + '…' : (str || '');
  }

  function formatBytes(bytes) {
    if (!bytes || bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  function markdownToHtml(md) {
    if (!md) return '';
    let html = escapeHtml(md);

    // Headers
    html = html.replace(/^### (.*$)/gm, '<h3 style="margin: 12px 0 6px; font-size: 15px; font-weight: 700; color: var(--text-main);">$1</h3>');
    html = html.replace(/^## (.*$)/gm, '<h2 style="margin: 14px 0 8px; font-size: 16px; font-weight: 700; color: var(--text-main);">$1</h2>');
    html = html.replace(/^# (.*$)/gm, '<h1 style="margin: 16px 0 10px; font-size: 18px; font-weight: 700; color: var(--text-main);">$1</h1>');

    // Inline code
    html = html.replace(/`([^`]+)`/g, '<code style="font-family: var(--font-mono); font-size: 12px; background: var(--bg-chip); padding: 2px 6px; border-radius: 4px; border: 1px solid var(--border-subtle); color: var(--color-primary);">$1</code>');

    // Bold
    html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    // Italic
    html = html.replace(/\*(.*?)\*/g, '<em>$1</em>');
    // Bullet list items
    html = html.replace(/^[•\-\*]\s+(.*)$/gm, '<li>$1</li>');
    html = html.replace(/(<li>.*<\/li>)/s, '<ul style="padding-left: 20px; margin-bottom: 10px;">$1</ul>');
    // Numbered list items
    html = html.replace(/^(\d+)\.\s+(.*)$/gm, '<li><strong>$1.</strong> $2</li>');
    // Paragraphs
    html = html.split('\n\n').map(p => {
      if (p.startsWith('<ul>') || p.startsWith('<li>') || p.startsWith('<h1') || p.startsWith('<h2') || p.startsWith('<h3')) return p;
      return `<p style="margin-bottom: 8px;">${p.replace(/\n/g, '<br>')}</p>`;
    }).join('');

    return html;
  }

  // Start the application
  document.addEventListener('DOMContentLoaded', init);

})();
