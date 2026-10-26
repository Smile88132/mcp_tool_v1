/**
 * VennFlow AI - Client Application Logic
 * Dual-Mode (Offline Local ML vs AselixAI API), Excel 5-Row Preview & Column Picker,
 * Copy-Paste Multi-Row Comparison, Interactive SVG Venn Diagram & QA Assistant.
 */

(function () {
  'use strict';

  // Config State (LocalStorage backed)
  const STORAGE_KEY = 'vennflow_aselix_config_v2';
  const DEFAULT_CONFIG = {
    network: 'milli',
    baseUrl: 'https://api.aselixai.aselsan.com.tr/v1',
    apiKey: '',
    model: 'aselix-llm-v1',
    engine: 'local', // 'local' or 'aselix'
    threshold: 80,   // %50 - %100
    batchSize: 15,   // 10, 15, 20 satırlık güvenli batch boyutu
    inputMode: 'excel' // 'excel' or 'paste'
  };

  let appConfig = loadConfig();
  let file1 = null;
  let file2 = null;
  let previewData1 = null;
  let previewData2 = null;
  let currentAnalysis = null;
  let currentFilter = 'all';

  // Network URLs mapping
  const NETWORK_URLS = {
    milli: 'https://api.aselixai.aselsan.com.tr/v1',
    aynet: 'https://agate.aselsan.com.tr/send/milli/api.aselixai/v1'
  };

  // DOM Elements
  const engineStatusPill = document.getElementById('engineStatusPill');
  const engineStatusLabel = document.getElementById('engineStatusLabel');
  const btnEngineLocal = document.getElementById('btnEngineLocal');
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
  const tabFilterAll = document.getElementById('tabFilterAll');
  const tabFilterIntersection = document.getElementById('tabFilterIntersection');
  const tabFilterA = document.getElementById('tabFilterA');
  const tabFilterB = document.getElementById('tabFilterB');
  const tableSearchInput = document.getElementById('tableSearchInput');
  const detailsTableBody = document.getElementById('detailsTableBody');
  const btnExportCSV = document.getElementById('btnExportCSV');
  const btnExportJSON = document.getElementById('btnExportJSON');

  // Chat Section
  const chatMessages = document.getElementById('chatMessages');
  const chatInput = document.getElementById('chatInput');
  const btnSendChat = document.getElementById('btnSendChat');
  const btnClearChat = document.getElementById('btnClearChat');

  // Settings Modal
  const btnOpenSettings = document.getElementById('btnOpenSettings');
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
  const fetchModelsStatus = document.getElementById('fetchModelsStatus');
  const testConnResult = document.getElementById('testConnResult');
  const radioMilli = document.getElementById('radioMilli');
  const radioAynet = document.getElementById('radioAynet');
  const radioCustom = document.getElementById('radioCustom');

  // --------------------------------------------------------------------------
  // Initialization
  // --------------------------------------------------------------------------
  function init() {
    applyConfigToUI();
    setupEventListeners();
  }

  function setupEventListeners() {
    // Engine mode toggle
    btnEngineLocal.addEventListener('click', () => setEngineMode('local'));
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

    // Filter Buttons
    tabFilterAll.addEventListener('click', () => setTableFilter('all'));
    tabFilterIntersection.addEventListener('click', () => setTableFilter('intersection'));
    tabFilterA.addEventListener('click', () => setTableFilter('only_a'));
    tabFilterB.addEventListener('click', () => setTableFilter('only_b'));

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
    currentModelDisplay.textContent = appConfig.model || 'aselix-llm-v1';
    chatModelNameDisplay.textContent = appConfig.model || 'aselix-llm-v1';

    // Settings modal inputs
    cfgBaseUrl.value = appConfig.baseUrl || NETWORK_URLS.milli;
    cfgApiKey.value = appConfig.apiKey || '';
    if (cfgBatchSize) cfgBatchSize.value = String(appConfig.batchSize || 15);

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
    if (mode === 'local') {
      btnEngineLocal.classList.add('active');
      btnEngineAselix.classList.remove('active');
      engineStatusLabel.textContent = '⚡ Yerel ML (Çevrimdışı)';
      engineStatusPill.className = 'flow-status-pill';
      btnRunText.textContent = 'Kıyasla & Venn Şemasını Çiz (Yerel)';
    } else {
      btnEngineAselix.classList.add('active');
      btnEngineLocal.classList.remove('active');
      engineStatusLabel.textContent = '🤖 AselixAI API Aktif';
      engineStatusPill.className = 'flow-status-pill online';
      btnRunText.textContent = 'Kıyasla, Analiz Et & Venn Çiz';
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
    if (progressBarWrapper) progressBarWrapper.style.display = 'block';
    if (progressBarFill) progressBarFill.style.width = '5%';
    if (progressPctBadge) progressPctBadge.textContent = '%5';
    btnRunComparison.disabled = true;

    const thresholdDec = (appConfig.threshold || 80) / 100.0;
    const isAselix = appConfig.engine === 'aselix';
    const batchSize = parseInt(appConfig.batchSize, 10) || 15;

    // Metin Modunda ve Aselix seçildiğinde: İstemci tarafında gerçek zamanlı chunking döngüsü
    if (inputMode === 'paste' && isAselix && appConfig.apiKey && appConfig.apiKey.trim()) {
      try {
        const lines1 = textPaste1.value.split('\n').filter(l => l.trim().length > 0);
        const lines2 = textPaste2.value.split('\n').filter(l => l.trim().length > 0);

        const chunks = [];
        for (let i = 0; i < lines1.length; i += batchSize) {
          chunks.push(lines1.slice(i, i + batchSize));
        }
        const totalBatches = chunks.length;
        let allMatchedPairs = [];

        for (let i = 0; i < totalBatches; i++) {
          const chunk = chunks[i];
          const startPct = Math.round((i / (totalBatches + 1)) * 100);
          if (progressBarFill) progressBarFill.style.width = `${Math.max(5, startPct)}%`;
          if (progressPctBadge) progressPctBadge.textContent = `%${Math.max(5, startPct)}`;

          progressStatusText.textContent = `Batch ${i + 1}/${totalBatches} AselixAI ile işleniyor (${chunk.length} satır)...`;
          progressSubText.textContent = `1 dakikalık zaman aşımı koruması devrede (Timeout: 50 sn) • Model: ${appConfig.model || 'aselix-llm-v1'}`;

          const chunkResp = await fetch('/api/compare/batch-chunk', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              base_url: appConfig.baseUrl,
              api_key: appConfig.apiKey,
              model: appConfig.model || 'aselix-llm-v1',
              items_a_batch: chunk,
              candidates_b: lines2,
              threshold: thresholdDec,
              timeout_sec: 50,
              batch_index: i + 1,
              total_batches: totalBatches
            })
          });

          const chunkData = await chunkResp.json();
          if (chunkData.matched_pairs && Array.isArray(chunkData.matched_pairs)) {
            allMatchedPairs.push(...chunkData.matched_pairs);
          }
        }

        // Finalize: Kümeleri birleştir & Yönetici Özeti üret
        progressStatusText.textContent = 'Tüm batch sonuçları birleştiriliyor ve AselixAI Yönetici Özeti üretiliyor...';
        progressSubText.textContent = `Toplam ${allMatchedPairs.length} eşleşme bulundu. Kök neden ve risk analizi derleniyor...`;
        if (progressBarFill) progressBarFill.style.width = '90%';
        if (progressPctBadge) progressPctBadge.textContent = '%90';

        const finalResp = await fetch('/api/compare/batch-finalize', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            base_url: appConfig.baseUrl,
            api_key: appConfig.apiKey,
            model: appConfig.model || 'aselix-llm-v1',
            threshold: thresholdDec,
            all_items_a: lines1,
            all_items_b: lines2,
            matched_pairs: allMatchedPairs,
            doc1_name: '1. Metin Seti',
            doc2_name: '2. Metin Seti',
            col1_name: 'Metin',
            col2_name: 'Metin',
            need_executive_summary: true
          })
        });

        const finalData = await finalResp.json();
        if (progressBarFill) progressBarFill.style.width = '100%';
        if (progressPctBadge) progressPctBadge.textContent = '%100';

        btnRunComparison.disabled = false;
        setTimeout(() => { progressBanner.style.display = 'none'; }, 400);

        if (finalData.durum === 'basarili') {
          renderAnalysisResults(finalData);
        } else {
          alert('Kıyaslama Hatası: ' + (finalData.mesaj || 'İşlem tamamlanamadı.'));
        }
        return;
      } catch (err) {
        btnRunComparison.disabled = false;
        progressBanner.style.display = 'none';
        alert('Batch iletişim hatası: ' + err.message);
        return;
      }
    }

    // Excel veya Yerel ML Modu
    progressStatusText.textContent = isAselix 
      ? 'Veriler 1 dakikalık güvenli batch\'ler halinde AselixAI API\'ye aktarılıyor...'
      : 'Veriler yerel mesafe ve fuzzy matching algoritması ile kıyaslanıyor...';
    progressSubText.textContent = `Benzerlik toleransı: %${appConfig.threshold} • Motor: ${isAselix ? 'AselixAI (Batch boyutu: ' + batchSize + ')' : 'Yerel ML'}`;

    let progressSimInterval = null;
    if (isAselix) {
      let currentSimPct = 10;
      progressSimInterval = setInterval(() => {
        if (currentSimPct < 85) {
          currentSimPct += 5;
          if (progressBarFill) progressBarFill.style.width = `${currentSimPct}%`;
          if (progressPctBadge) progressPctBadge.textContent = `%${currentSimPct}`;
        }
      }, 1500);
    }

    try {
      const formData = new FormData();
      formData.append('input_mode', inputMode);
      formData.append('engine_mode', appConfig.engine);
      formData.append('threshold', thresholdDec);
      formData.append('base_url', appConfig.baseUrl);
      formData.append('api_key', appConfig.apiKey || '');
      formData.append('model', appConfig.model || 'aselix-llm-v1');

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

      const resp = await fetch('/api/compare/run', {
        method: 'POST',
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
  function renderAnalysisResults(data) {
    currentAnalysis = data;
    resultsSection.style.display = 'block';
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });

    const v = data.venn_verileri;
    const isAselix = data.engine_mode === 'aselix';

    // Summary Engine Badge
    summaryEngineBadge.innerHTML = isAselix
      ? `<span class="engine-icon">🤖</span><span>ASELSAN AselixAI (${data.model_kullanilan || appConfig.model}) Yönetici Özeti</span>`
      : `<span class="engine-icon">⚡</span><span>Yerel ML & String Distance (%${Math.round((data.esik_orani || 0.8) * 100)} Eşik)</span>`;

    // Summary Text
    if (data.flow_analiz_ozeti && data.flow_analiz_ozeti.trim()) {
      summaryText.innerHTML = markdownToHtml(data.flow_analiz_ozeti);
    } else {
      summaryText.innerHTML = `
        <div style="line-height: 1.7;">
          <p><strong>${escapeHtml(data.ozet_yorum || '')}</strong></p>
          <p style="margin-top: 8px; color: var(--text-muted); font-size: 13px;">
            ℹ️ AselixAI ile derinlemesine kök neden analizi ve önleyici faaliyet önerileri almak için yukarıdan 
            <strong>AselixAI API</strong> modunu aktif edebilir ve Ayarlar'dan API anahtarınızı tanımlayabilirsiniz.
          </p>
        </div>
      `;
    }

    // Metrics
    statTotal.textContent = v.total_unique;
    statIntersection.textContent = v.intersection_count;
    statJaccard.textContent = `Jaccard Benzerlik Oranı: %${Math.round((v.jaccard_similarity || 0) * 100)}`;
    statOnlyA.textContent = v.only_a_count;
    statOnlyB.textContent = v.only_b_count;

    statOnlyASub.textContent = `${data.dosya_1_adi || '1. Belge'} içinde tekil`;
    statOnlyBSub.textContent = `${data.dosya_2_adi || '2. Belge'} içinde yeni`;

    legendTextA.textContent = truncateStr(data.dosya_1_adi || '1. Belge', 20);
    legendTextB.textContent = truncateStr(data.dosya_2_adi || '2. Belge', 20);

    // Update Filter Tab Badges
    countAll.textContent = v.total_unique;
    countIntersection.textContent = v.intersection_count;
    countA.textContent = v.only_a_count;
    countB.textContent = v.only_b_count;

    // Draw Venn SVG
    drawVennDiagram(v, data.dosya_1_adi, data.dosya_2_adi);

    // Render Table
    currentFilter = 'all';
    updateFilterTabButtons();
    renderTableRows();
  }

  // --------------------------------------------------------------------------
  // Draw Interactive SVG Venn Diagram
  // --------------------------------------------------------------------------
  function drawVennDiagram(vennData, nameA, nameB) {
    const labelA = truncateStr(nameA || '1. Belge', 16);
    const labelB = truncateStr(nameB || '2. Belge', 16);

    const svgContent = `
      <svg viewBox="0 0 640 360" width="100%" height="100%" style="overflow: visible;">
        <defs>
          <linearGradient id="gradCircleA" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="#3b82f6" stop-opacity="0.55"/>
            <stop offset="100%" stop-color="#1d4ed8" stop-opacity="0.35"/>
          </linearGradient>

          <linearGradient id="gradCircleB" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="#ec4899" stop-opacity="0.55"/>
            <stop offset="100%" stop-color="#8b5cf6" stop-opacity="0.35"/>
          </linearGradient>

          <linearGradient id="gradCommon" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="#60a5fa" stop-opacity="0.75"/>
            <stop offset="100%" stop-color="#c084fc" stop-opacity="0.75"/>
          </linearGradient>

          <clipPath id="clipCircleA">
            <circle cx="230" cy="180" r="140" />
          </clipPath>
        </defs>

        <!-- Circle A Full Disc -->
        <circle id="svgCircleA" cx="230" cy="180" r="140"
                fill="url(#gradCircleA)" stroke="#60a5fa" stroke-width="2.5"
                class="venn-circle" data-region="only_a" style="cursor: pointer;" />

        <!-- Circle B Full Disc -->
        <circle id="svgCircleB" cx="410" cy="180" r="140"
                fill="url(#gradCircleB)" stroke="#f472b6" stroke-width="2.5"
                class="venn-circle" data-region="only_b" style="cursor: pointer;" />

        <!-- Intersection Lens (B clipped inside A) -->
        <circle id="svgIntersection" cx="410" cy="180" r="140"
                clip-path="url(#clipCircleA)" fill="url(#gradCommon)"
                stroke="#a78bfa" stroke-width="2"
                class="venn-intersection" data-region="intersection" style="cursor: pointer;" />

        <!-- Set A Only Count & Label -->
        <text x="160" y="172" class="venn-text-count" style="font-family: var(--font-heading); font-size: 28px; font-weight: 700; fill: #93c5fd; text-anchor: middle;">${vennData.only_a_count}</text>
        <text x="160" y="202" class="venn-text-label" style="font-size: 13px; font-weight: 600; fill: #cbd5e1; text-anchor: middle;">Yalnızca 1. Belge</text>
        <text x="160" y="220" class="venn-text-label" style="font-size: 11px; fill: #93c5fd; text-anchor: middle;">(${escapeHtml(labelA)})</text>

        <!-- Intersection Count & Label -->
        <text x="320" y="165" class="venn-text-count" style="font-family: var(--font-heading); fill: #ffffff; font-size: 34px; font-weight: 800; text-anchor: middle; filter: drop-shadow(0 0 10px rgba(255,255,255,0.8));">${vennData.intersection_count}</text>
        <text x="320" y="198" class="venn-text-label" style="font-size: 14px; font-weight: 700; fill: #ffffff; text-anchor: middle;">Ortak Kesişim</text>
        <text x="320" y="216" class="venn-text-label" style="font-size: 11px; fill: #e9d5ff; text-anchor: middle;">(Tekrarlayanlar)</text>

        <!-- Set B Only Count & Label -->
        <text x="480" y="172" class="venn-text-count" style="font-family: var(--font-heading); font-size: 28px; font-weight: 700; fill: #f9a8d4; text-anchor: middle;">${vennData.only_b_count}</text>
        <text x="480" y="202" class="venn-text-label" style="font-size: 13px; font-weight: 600; fill: #cbd5e1; text-anchor: middle;">Yalnızca 2. Belge</text>
        <text x="480" y="220" class="venn-text-label" style="font-size: 11px; fill: #f9a8d4; text-anchor: middle;">(${escapeHtml(labelB)})</text>
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
    [tabFilterAll, tabFilterIntersection, tabFilterA, tabFilterB].forEach(btn => btn.classList.remove('active'));
    if (currentFilter === 'all') tabFilterAll.classList.add('active');
    else if (currentFilter === 'intersection') tabFilterIntersection.classList.add('active');
    else if (currentFilter === 'only_a') tabFilterA.classList.add('active');
    else if (currentFilter === 'only_b') tabFilterB.classList.add('active');
  }

  function renderTableRows() {
    if (!currentAnalysis) return;

    const v = currentAnalysis.venn_verileri;
    const query = tableSearchInput.value.toLowerCase().trim();

    let rows = [];

    // 1. Matched Pairs (Intersection)
    (v.matched_pairs || []).forEach(p => {
      rows.push({
        itemA: p.item_a,
        itemB: p.item_b,
        similarity: p.similarity,
        region: 'intersection',
        matchType: p.match_type
      });
    });

    // 2. Only A
    (v.only_a_items || []).forEach(item => {
      rows.push({
        itemA: item,
        itemB: '-',
        similarity: 0,
        region: 'only_a',
        matchType: 'Yalnızca 1. Belge'
      });
    });

    // 3. Only B
    (v.only_b_items || []).forEach(item => {
      rows.push({
        itemA: '-',
        itemB: item,
        similarity: 0,
        region: 'only_b',
        matchType: 'Yalnızca 2. Belge'
      });
    });

    // Filter by Region
    if (currentFilter !== 'all') {
      rows = rows.filter(r => r.region === currentFilter);
    }

    // Filter by Search Query
    if (query) {
      rows = rows.filter(r => (
        (r.itemA && r.itemA.toLowerCase().includes(query)) ||
        (r.itemB && r.itemB.toLowerCase().includes(query)) ||
        (r.matchType && r.matchType.toLowerCase().includes(query))
      ));
    }

    if (rows.length === 0) {
      detailsTableBody.innerHTML = `
        <tr>
          <td colspan="4" style="text-align: center; padding: 24px; color: var(--text-muted);">
            Arama kriterine uygun kayıt bulunamadı.
          </td>
        </tr>
      `;
      return;
    }

    detailsTableBody.innerHTML = rows.map((r, idx) => {
      let badgeHtml = '';
      if (r.region === 'intersection') {
        if (r.similarity >= 99.5) {
          badgeHtml = `<span class="badge-similarity-exact">✓ %100 Tam Eşleşme</span>`;
        } else {
          badgeHtml = `<span class="badge-similarity-fuzzy">≈ %${r.similarity} Benzer</span>`;
        }
      } else if (r.region === 'only_a') {
        badgeHtml = `<span class="badge-similarity-only a">Yalnızca 1. Belge</span>`;
      } else {
        badgeHtml = `<span class="badge-similarity-only b">Yalnızca 2. Belge</span>`;
      }

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
        appendChatMessage('bot', `⚠️ AselixAI Yanıtı: ${data.mesaj || 'Cevap üretilemedi. Lütfen Ayarlar menüsünden API anahtarınızı ve ağ seçiminizi kontrol edin.'}`);
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
  // Settings Modal Functions
  // --------------------------------------------------------------------------
  function openSettingsModal() {
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

    saveConfig();
    applyConfigToUI();
    closeSettingsModal();
  }

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

    // Bold
    html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    // Italic
    html = html.replace(/\*(.*?)\*/g, '<em>$1</em>');
    // Bullet list items
    html = html.replace(/^[•\-\*]\s+(.*)$/gm, '<li>$1</li>');
    html = html.replace(/(<li>.*<\/li>)/s, '<ul>$1</ul>');
    // Numbered list items
    html = html.replace(/^(\d+)\.\s+(.*)$/gm, '<li><strong>$1.</strong> $2</li>');
    // Paragraphs
    html = html.split('\n\n').map(p => {
      if (p.startsWith('<ul>') || p.startsWith('<li>')) return p;
      return `<p>${p.replace(/\n/g, '<br>')}</p>`;
    }).join('');

    return html;
  }

  // Start the application
  document.addEventListener('DOMContentLoaded', init);

})();
