/**
 * VennFlow AI - Client Application Logic
 * Drag & Drop, Interactive SVG Venn Diagram, ASELSAN LLM Suite Flow API Client
 */

(function () {
  'use strict';

  // Config State (LocalStorage backed)
  const DEFAULT_CONFIG = {
    baseUrl: 'https://llmsuiteflow.aiplatform.aselsan.com.tr',
    flowId: 'c16ecc83-ff64-4fc7-bb52-0f594beaf70a',
    apiKey: '',
    tweaksInput: 'ChatInput-jugL4'
  };

  let appConfig = loadConfig();
  let file1 = null;
  let file2 = null;
  let currentAnalysis = null;
  let currentFilter = 'all';
  let activeSessionId = null;

  // DOM Elements
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

  const btnRunComparison = document.getElementById('btnRunComparison');
  const btnUseSampleFiles = document.getElementById('btnUseSampleFiles');
  const progressBanner = document.getElementById('progressBanner');
  const progressStatusText = document.getElementById('progressStatusText');
  const progressSubText = document.getElementById('progressSubText');
  const resultsSection = document.getElementById('resultsSection');

  // Metrics
  const valOnlyA = document.getElementById('valOnlyA');
  const valCommon = document.getElementById('valCommon');
  const valOnlyB = document.getElementById('valOnlyB');
  const valSimilarity = document.getElementById('valSimilarity');
  const valTotalUnique = document.getElementById('valTotalUnique');
  const labelDocA = document.getElementById('labelDocA');
  const labelDocB = document.getElementById('labelDocB');
  const legendTitleA = document.getElementById('legendTitleA');
  const legendTitleB = document.getElementById('legendTitleB');

  // Venn SVG
  const vennSvg = document.getElementById('vennSvg');
  const countAll = document.getElementById('countAll');
  const countOnlyA = document.getElementById('countOnlyA');
  const countCommon = document.getElementById('countCommon');
  const countOnlyB = document.getElementById('countOnlyB');

  // AI Summary & Chat
  const aiSummaryBody = document.getElementById('aiSummaryBody');
  const flowModelMeta = document.getElementById('flowModelMeta');
  const btnCopyAiSummary = document.getElementById('btnCopyAiSummary');
  const chatMessages = document.getElementById('chatMessages');
  const chatInput = document.getElementById('chatInput');
  const btnSendChat = document.getElementById('btnSendChat');
  const btnClearChat = document.getElementById('btnClearChat');

  // Table
  const tableBody = document.getElementById('tableBody');
  const tableSearchInput = document.getElementById('tableSearchInput');
  const btnExportCsv = document.getElementById('btnExportCsv');

  // Settings Modal
  const btnOpenSettings = document.getElementById('btnOpenSettings');
  const settingsModal = document.getElementById('settingsModal');
  const btnCloseSettings = document.getElementById('btnCloseSettings');
  const btnSaveSettings = document.getElementById('btnSaveSettings');
  const btnResetSettings = document.getElementById('btnResetSettings');
  const cfgBaseUrl = document.getElementById('cfgBaseUrl');
  const cfgFlowId = document.getElementById('cfgFlowId');
  const cfgApiKey = document.getElementById('cfgApiKey');
  const cfgTweaksInput = document.getElementById('cfgTweaksInput');
  const btnToggleApiKey = document.getElementById('btnToggleApiKey');

  // --------------------------------------------------------------------------
  // Initialization
  // --------------------------------------------------------------------------
  function init() {
    setupDragAndDrop(dropzone1, fileInput1, (file) => handleFileSelect(1, file));
    setupDragAndDrop(dropzone2, fileInput2, (file) => handleFileSelect(2, file));

    btnRemoveFile1.addEventListener('click', (e) => { e.stopPropagation(); removeFile(1); });
    btnRemoveFile2.addEventListener('click', (e) => { e.stopPropagation(); removeFile(2); });

    btnRunComparison.addEventListener('click', runComparisonFlow);
    btnUseSampleFiles.addEventListener('click', loadSampleComparison);

    // Filter Buttons
    document.querySelectorAll('.legend-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.legend-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentFilter = btn.dataset.filter;
        renderDetailsTable();
      });
    });

    tableSearchInput.addEventListener('input', renderDetailsTable);
    btnExportCsv.addEventListener('click', exportToCSV);

    // Chat
    btnSendChat.addEventListener('click', sendChatMessage);
    chatInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        sendChatMessage();
      }
    });

    btnClearChat.addEventListener('click', () => {
      chatMessages.innerHTML = `
        <div class="chat-msg bot-msg">
          <div class="msg-avatar">🤖</div>
          <div class="msg-bubble">
            <p>Sohbet geçmişi temizlendi. İki Excel dosyasının kıyaslama sonuçları hakkında istediğiniz soruyu sorabilirsiniz.</p>
          </div>
        </div>
      `;
    });

    document.querySelectorAll('.chip-btn').forEach(chip => {
      chip.addEventListener('click', () => {
        chatInput.value = chip.dataset.query;
        chatInput.focus();
        sendChatMessage();
      });
    });

    btnCopyAiSummary.addEventListener('click', () => {
      if (currentAnalysis && currentAnalysis.flow_analiz_ozeti) {
        navigator.clipboard.writeText(currentAnalysis.flow_analiz_ozeti);
        const originalText = btnCopyAiSummary.querySelector('span').textContent;
        btnCopyAiSummary.querySelector('span').textContent = 'Kopyalandı!';
        setTimeout(() => {
          btnCopyAiSummary.querySelector('span').textContent = originalText;
        }, 2000);
      }
    });

    // Settings
    setupSettingsModal();
    updateFlowStatusPill();
  }

  // --------------------------------------------------------------------------
  // Drag & Drop Handlers
  // --------------------------------------------------------------------------
  function setupDragAndDrop(dropzone, input, onFile) {
    dropzone.addEventListener('click', () => input.click());

    input.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        onFile(e.target.files[0]);
      }
    });

    ['dragenter', 'dragover'].forEach(eventName => {
      dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.remove('dragover');
      });
    });

    dropzone.addEventListener('drop', (e) => {
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        onFile(e.dataTransfer.files[0]);
      }
    });
  }

  function handleFileSelect(index, file) {
    const validExts = ['.xlsx', '.xls', '.xlsm'];
    const hasValidExt = validExts.some(ext => file.name.toLowerCase().endsWith(ext));
    if (!hasValidExt) {
      alert('Lütfen geçerli bir Excel (.xlsx, .xls) dosyası seçin.');
      return;
    }

    if (index === 1) {
      file1 = file;
      fileName1.textContent = file.name;
      fileSub1.textContent = `${formatBytes(file.size)} • Seçildi`;
      dropzoneContent1.style.display = 'none';
      fileView1.style.display = 'flex';
      labelDocA.textContent = `Yalnızca ${file.name.replace(/\.[^/.]+$/, "")}`;
      legendTitleA.textContent = `Sadece ${truncateStr(file.name, 14)}`;
    } else {
      file2 = file;
      fileName2.textContent = file.name;
      fileSub2.textContent = `${formatBytes(file.size)} • Seçildi`;
      dropzoneContent2.style.display = 'none';
      fileView2.style.display = 'flex';
      labelDocB.textContent = `Yalnızca ${file.name.replace(/\.[^/.]+$/, "")}`;
      legendTitleB.textContent = `Sadece ${truncateStr(file.name, 14)}`;
    }
  }

  function removeFile(index) {
    if (index === 1) {
      file1 = null;
      fileInput1.value = '';
      dropzoneContent1.style.display = 'flex';
      fileView1.style.display = 'none';
      labelDocA.textContent = 'Yalnızca 1. Belgede';
      legendTitleA.textContent = 'Sadece 1. Belge';
    } else {
      file2 = null;
      fileInput2.value = '';
      dropzoneContent2.style.display = 'flex';
      fileView2.style.display = 'none';
      labelDocB.textContent = 'Yalnızca 2. Belgede';
      legendTitleB.textContent = 'Sadece 2. Belge';
    }
  }

  function formatBytes(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  function truncateStr(str, n) {
    return (str.length > n) ? str.substr(0, n - 1) + '…' : str;
  }

  // --------------------------------------------------------------------------
  // Sample Data Loader
  // --------------------------------------------------------------------------
  async function loadSampleComparison() {
    progressBanner.style.display = 'flex';
    progressStatusText.textContent = 'Örnek uygunsuzluk verileri sunucudan yükleniyor...';
    progressSubText.textContent = 'Dry-run ve SKT fazları karşılaştırma verisi hazırlanıyor...';

    try {
      const resp = await fetch('/api/flow/sample-compare', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          base_url: appConfig.baseUrl,
          flow_id: appConfig.flowId,
          api_key: appConfig.apiKey,
          tweaks_input: appConfig.tweaksInput
        })
      });

      const data = await resp.json();
      progressBanner.style.display = 'none';

      if (data.durum === 'basarili') {
        fileName1.textContent = data.dosya_1_adi || 'Örnek_Faz_1_DryRun.xlsx';
        fileSub1.textContent = 'Hazır Test Verisi';
        dropzoneContent1.style.display = 'none';
        fileView1.style.display = 'flex';

        fileName2.textContent = data.dosya_2_adi || 'Örnek_Faz_2_SKT.xlsx';
        fileSub2.textContent = 'Hazır Test Verisi';
        dropzoneContent2.style.display = 'none';
        fileView2.style.display = 'flex';

        renderAnalysisResults(data);
      } else {
        alert('Hata: ' + (data.mesaj || 'Örnek veri yüklenemedi.'));
      }
    } catch (err) {
      progressBanner.style.display = 'none';
      alert('Bağlantı hatası: ' + err.message);
    }
  }

  // --------------------------------------------------------------------------
  // Run Comparison Flow
  // --------------------------------------------------------------------------
  async function runComparisonFlow() {
    if (!file1 || !file2) {
      alert('Lütfen kıyaslamak için her iki Excel dosyasını da sürükleyip bırakın veya seçin.');
      return;
    }

    progressBanner.style.display = 'flex';
    progressStatusText.textContent = 'Excel belgeleri okunuyor ve küme analizi yapılıyor...';
    progressSubText.textContent = `ASELSAN LLM Suite Flow (${appConfig.flowId.slice(0, 8)}...) çağrılıyor...`;

    btnRunComparison.disabled = true;

    try {
      const formData = new FormData();
      formData.append('file_1', file1);
      formData.append('file_2', file2);
      formData.append('base_url', appConfig.baseUrl);
      formData.append('flow_id', appConfig.flowId);
      formData.append('api_key', appConfig.apiKey || '');
      formData.append('tweaks_input', appConfig.tweaksInput || 'ChatInput-jugL4');

      const colMatch = document.getElementById('colMatchSelect').value;
      if (colMatch && colMatch !== 'auto') {
        formData.append('column_name', colMatch);
      }

      const resp = await fetch('/api/flow/compare', {
        method: 'POST',
        body: formData
      });

      const data = await resp.json();
      btnRunComparison.disabled = false;
      progressBanner.style.display = 'none';

      if (data.durum === 'basarili') {
        renderAnalysisResults(data);
      } else {
        alert('Analiz Hatası: ' + (data.mesaj || 'Bilinmeyen bir hata oluştu.'));
      }
    } catch (err) {
      btnRunComparison.disabled = false;
      progressBanner.style.display = 'none';
      alert('Sunucu ile iletişim kurulurken hata oluştu: ' + err.message);
    }
  }

  // --------------------------------------------------------------------------
  // Render Results & Interactive Venn
  // --------------------------------------------------------------------------
  function renderAnalysisResults(data) {
    currentAnalysis = data;
    activeSessionId = data.session_id || generateUUID();
    resultsSection.style.display = 'block';

    // Scroll smoothly to results
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });

    const v = data.venn_verileri;

    // Update Metrics
    valOnlyA.textContent = v.only_a_count;
    valCommon.textContent = v.intersection_count;
    valOnlyB.textContent = v.only_b_count;
    valSimilarity.textContent = `%${Math.round((v.jaccard_similarity || 0) * 100)}`;
    valTotalUnique.textContent = `Toplam ${v.total_unique} benzersiz uygunsuzluk analiz edildi.`;

    // Update Counts on Filter Buttons
    countAll.textContent = v.total_unique;
    countOnlyA.textContent = v.only_a_count;
    countCommon.textContent = v.intersection_count;
    countOnlyB.textContent = v.only_b_count;

    // AI Summary
    if (data.flow_analiz_ozeti) {
      aiSummaryBody.innerHTML = markdownToHtml(data.flow_analiz_ozeti);
      flowModelMeta.textContent = `Flow ID: ${appConfig.flowId.slice(0, 8)}... • Yapay Zeka Çıktısı`;
    } else {
      aiSummaryBody.innerHTML = `
        <div class="note-box">
          <p><strong>Matematiksel Venn Şeması Başarıyla Çizildi!</strong></p>
          <p>ASELSAN LLM Suite Flow'dan detaylı yapay zeka analiz özeti almak için sağ üstteki <strong>Ayarlar</strong> menüsünden API anahtarınızı (x-api-key) tanımlayabilirsiniz.</p>
        </div>
      `;
    }

    // Render Dynamic SVG Venn Diagram
    drawVennDiagram(v, data.dosya_1_adi, data.dosya_2_adi);

    // Render Table
    renderDetailsTable();
  }

  // --------------------------------------------------------------------------
  // Draw Interactive SVG Venn Diagram
  // --------------------------------------------------------------------------
  function drawVennDiagram(vennData, nameA, nameB) {
    const labelA = truncateStr(nameA || '1. Belge', 16);
    const labelB = truncateStr(nameB || '2. Belge', 16);

    // SVG geometry: width=640, height=360
    // Circle A center=(230, 180), r=140
    // Circle B center=(410, 180), r=140
    // Distance between centers = 180 (creates perfect realistic Venn overlap)
    const svgContent = `
      <defs>
        <!-- Gradients -->
        <linearGradient id="gradCircleA" x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stop-color="#3b82f6" stop-opacity="0.55"/>
          <stop offset="100%" stop-color="#1d4ed8" stop-opacity="0.35"/>
        </linearGradient>

        <linearGradient id="gradCircleB" x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stop-color="#ec4899" stop-opacity="0.55"/>
          <stop offset="100%" stop-color="#8b5cf6" stop-opacity="0.35"/>
        </linearGradient>

        <linearGradient id="gradCommon" x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stop-color="#60a5fa" stop-opacity="0.7"/>
          <stop offset="100%" stop-color="#c084fc" stop-opacity="0.7"/>
        </linearGradient>

        <!-- Filters for glow -->
        <filter id="glowA" x="-20%" y="-20%" width="140%" height="140%">
          <feGaussianBlur stdDeviation="6" result="blur"/>
          <feComposite in="SourceGraphic" in2="blur" operator="over"/>
        </filter>
        <filter id="glowCommon" x="-20%" y="-20%" width="140%" height="140%">
          <feGaussianBlur stdDeviation="8" result="blur"/>
          <feComposite in="SourceGraphic" in2="blur" operator="over"/>
        </filter>

        <!-- Clip Paths for exact click detection -->
        <clipPath id="clipCircleA">
          <circle cx="230" cy="180" r="140" />
        </clipPath>
        <clipPath id="clipCircleB">
          <circle cx="410" cy="180" r="140" />
        </clipPath>
      </defs>

      <!-- Circle A Full Disc -->
      <circle id="svgCircleA" cx="230" cy="180" r="140"
              fill="url(#gradCircleA)" stroke="#60a5fa" stroke-width="2.5"
              class="venn-circle" data-region="only_a" />

      <!-- Circle B Full Disc -->
      <circle id="svgCircleB" cx="410" cy="180" r="140"
              fill="url(#gradCircleB)" stroke="#f472b6" stroke-width="2.5"
              class="venn-circle" data-region="only_b" />

      <!-- Exact Intersection Lens (B clipped inside A) -->
      <circle id="svgIntersection" cx="410" cy="180" r="140"
              clip-path="url(#clipCircleA)" fill="url(#gradCommon)"
              stroke="#a78bfa" stroke-width="2"
              class="venn-intersection" data-region="common" />

      <!-- Set A Only Count & Label -->
      <text x="160" y="172" class="venn-text-count">${vennData.only_a_count}</text>
      <text x="160" y="202" class="venn-text-label">Yalnızca 1. Belge</text>
      <text x="160" y="220" class="venn-text-label" style="fill: #93c5fd; font-size: 11px;">(${labelA})</text>

      <!-- Intersection Count & Label -->
      <text x="320" y="165" class="venn-text-count" style="fill: #ffffff; font-size: 32px; filter: drop-shadow(0 0 8px rgba(255,255,255,0.7));">${vennData.intersection_count}</text>
      <text x="320" y="198" class="venn-text-label" style="fill: #ffffff; font-weight: 700;">Ortak Kesişim</text>
      <text x="320" y="216" class="venn-text-label" style="fill: #e9d5ff; font-size: 11px;">Her İki Belgede Ortak</text>

      <!-- Set B Only Count & Label -->
      <text x="480" y="172" class="venn-text-count">${vennData.only_b_count}</text>
      <text x="480" y="202" class="venn-text-label">Yalnızca 2. Belge</text>
      <text x="480" y="220" class="venn-text-label" style="fill: #f9a8d4; font-size: 11px;">(${labelB})</text>
    `;

    vennSvg.innerHTML = svgContent;

    // Attach click events on SVG regions
    document.querySelectorAll('#vennSvg [data-region]').forEach(el => {
      el.addEventListener('click', () => {
        const reg = el.dataset.region;
        const targetBtn = document.querySelector(`.legend-btn[data-filter="${reg}"]`);
        if (targetBtn) targetBtn.click();
      });
    });
  }

  // --------------------------------------------------------------------------
  // Render Details Table with Search & Filter
  // --------------------------------------------------------------------------
  function renderDetailsTable() {
    if (!currentAnalysis) return;

    const v = currentAnalysis.venn_verileri;
    const searchTerm = tableSearchInput.value.toLowerCase().trim();

    let allItems = [];

    // Tag each item
    (v.only_a_items || []).forEach(item => {
      allItems.push({
        text: item,
        region: 'only_a',
        badgeClass: 'badge-set-a',
        badgeLabel: 'Yalnızca 1. Belge',
        source: currentAnalysis.dosya_1_adi || '1. Belge'
      });
    });

    (v.intersection_items || []).forEach(item => {
      allItems.push({
        text: item,
        region: 'common',
        badgeClass: 'badge-common',
        badgeLabel: 'Ortak Kesişim',
        source: 'Her İki Belge'
      });
    });

    (v.only_b_items || []).forEach(item => {
      allItems.push({
        text: item,
        region: 'only_b',
        badgeClass: 'badge-set-b',
        badgeLabel: 'Yalnızca 2. Belge',
        source: currentAnalysis.dosya_2_adi || '2. Belge'
      });
    });

    // Filter by Region
    if (currentFilter !== 'all') {
      allItems = allItems.filter(i => i.region === currentFilter);
    }

    // Filter by Search
    if (searchTerm) {
      allItems = allItems.filter(i => i.text.toLowerCase().includes(searchTerm));
    }

    if (allItems.length === 0) {
      tableBody.innerHTML = `
        <tr>
          <td colspan="4" class="text-center text-muted" style="padding: 24px;">
            Arama kriterine veya seçili filtreye uygun kayıt bulunamadı.
          </td>
        </tr>
      `;
      return;
    }

    tableBody.innerHTML = allItems.map((item, idx) => `
      <tr>
        <td style="color: var(--text-dim); font-family: var(--font-mono); font-size: 11px;">${idx + 1}</td>
        <td><span class="table-badge ${item.badgeClass}">${item.badgeLabel}</span></td>
        <td style="font-weight: 500;">${escapeHtml(item.text)}</td>
        <td style="color: var(--text-muted); font-size: 12px;">${escapeHtml(truncateStr(item.source, 20))}</td>
      </tr>
    `).join('');
  }

  function exportToCSV() {
    if (!currentAnalysis) return;
    const v = currentAnalysis.venn_verileri;

    let rows = [["Sira", "Durum Kumesi", "Uygunsuzluk Tanimi", "Kaynak"]];
    let idx = 1;

    (v.only_a_items || []).forEach(item => rows.push([idx++, "Yalnizca 1. Belge", `"${item.replace(/"/g, '""')}"`, currentAnalysis.dosya_1_adi]));
    (v.intersection_items || []).forEach(item => rows.push([idx++, "Ortak Kesisim", `"${item.replace(/"/g, '""')}"`, "Her Iki Belge"]));
    (v.only_b_items || []).forEach(item => rows.push([idx++, "Yalnizca 2. Belge", `"${item.replace(/"/g, '""')}"`, currentAnalysis.dosya_2_adi]));

    const csvContent = "\uFEFF" + rows.map(e => e.join(";")).join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `Venn_Kiyaslama_Raporu_${Date.now()}.csv`;
    link.click();
  }

  // --------------------------------------------------------------------------
  // Chat Integration with Flow API
  // --------------------------------------------------------------------------
  async function sendChatMessage() {
    const text = chatInput.value.trim();
    if (!text) return;

    chatInput.value = '';

    // Append User Message
    appendChatMessage('user', text);

    // Append Temporary Bot Loading Message
    const loadingId = 'loading-' + Date.now();
    const loadingEl = document.createElement('div');
    loadingEl.className = 'chat-msg bot-msg';
    loadingEl.id = loadingId;
    loadingEl.innerHTML = `
      <div class="msg-avatar">🤖</div>
      <div class="msg-bubble">
        <div style="display: flex; align-items: center; gap: 8px;">
          <div class="spinner" style="width: 16px; height: 16px; border-width: 2px;"></div>
          <span>ASELSAN LLM Suite yanıt hazırlıyor...</span>
        </div>
      </div>
    `;
    chatMessages.appendChild(loadingEl);
    chatMessages.scrollTop = chatMessages.scrollHeight;

    try {
      const resp = await fetch('/api/flow/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: text,
          session_id: activeSessionId,
          base_url: appConfig.baseUrl,
          flow_id: appConfig.flowId,
          api_key: appConfig.apiKey,
          tweaks_input: appConfig.tweaksInput,
          context: currentAnalysis ? {
            dosya_1: currentAnalysis.dosya_1_adi,
            dosya_2: currentAnalysis.dosya_2_adi,
            only_a_count: currentAnalysis.venn_verileri.only_a_count,
            intersection_count: currentAnalysis.venn_verileri.intersection_count,
            only_b_count: currentAnalysis.venn_verileri.only_b_count,
            jaccard: currentAnalysis.venn_verileri.jaccard_similarity
          } : null
        })
      });

      const data = await resp.json();
      const loadNode = document.getElementById(loadingId);
      if (loadNode) loadNode.remove();

      if (data.durum === 'basarili' && data.cevap) {
        appendChatMessage('bot', data.cevap);
      } else {
        appendChatMessage('bot', '⚠️ ' + (data.mesaj || 'Modelden yanıt alınamadı. Lütfen API Key ayarlarınızı kontrol edin.'));
      }
    } catch (err) {
      const loadNode = document.getElementById(loadingId);
      if (loadNode) loadNode.remove();
      appendChatMessage('bot', '⚠️ Bağlantı hatası: ' + err.message);
    }
  }

  function appendChatMessage(sender, text) {
    const msgEl = document.createElement('div');
    msgEl.className = `chat-msg ${sender === 'user' ? 'user-msg' : 'bot-msg'}`;
    const avatar = sender === 'user' ? '👤' : '🤖';
    msgEl.innerHTML = `
      <div class="msg-avatar">${avatar}</div>
      <div class="msg-bubble">${markdownToHtml(text)}</div>
    `;
    chatMessages.appendChild(msgEl);
    chatMessages.scrollTop = chatMessages.scrollHeight;
  }

  // --------------------------------------------------------------------------
  // Settings Modal Handlers
  // --------------------------------------------------------------------------
  function setupSettingsModal() {
    btnOpenSettings.addEventListener('click', () => {
      cfgBaseUrl.value = appConfig.baseUrl;
      cfgFlowId.value = appConfig.flowId;
      cfgApiKey.value = appConfig.apiKey;
      cfgTweaksInput.value = appConfig.tweaksInput;
      settingsModal.style.display = 'flex';
    });

    btnCloseSettings.addEventListener('click', () => {
      settingsModal.style.display = 'none';
    });

    settingsModal.addEventListener('click', (e) => {
      if (e.target === settingsModal) {
        settingsModal.style.display = 'none';
      }
    });

    btnToggleApiKey.addEventListener('click', () => {
      if (cfgApiKey.type === 'password') {
        cfgApiKey.type = 'text';
        btnToggleApiKey.textContent = 'Gizle';
      } else {
        cfgApiKey.type = 'password';
        btnToggleApiKey.textContent = 'Göster';
      }
    });

    btnResetSettings.addEventListener('click', () => {
      cfgBaseUrl.value = DEFAULT_CONFIG.baseUrl;
      cfgFlowId.value = DEFAULT_CONFIG.flowId;
      cfgApiKey.value = '';
      cfgTweaksInput.value = DEFAULT_CONFIG.tweaksInput;
    });

    btnSaveSettings.addEventListener('click', () => {
      appConfig = {
        baseUrl: cfgBaseUrl.value.trim() || DEFAULT_CONFIG.baseUrl,
        flowId: cfgFlowId.value.trim() || DEFAULT_CONFIG.flowId,
        apiKey: cfgApiKey.value.trim(),
        tweaksInput: cfgTweaksInput.value.trim() || DEFAULT_CONFIG.tweaksInput
      };
      saveConfig(appConfig);
      settingsModal.style.display = 'none';
      updateFlowStatusPill();
      alert('Ayarlar kaydedildi!');
    });
  }

  function loadConfig() {
    try {
      const stored = localStorage.getItem('aselsan_flow_config');
      if (stored) {
        return { ...DEFAULT_CONFIG, ...JSON.parse(stored) };
      }
    } catch (e) { }
    return { ...DEFAULT_CONFIG };
  }

  function saveConfig(cfg) {
    try {
      localStorage.setItem('aselsan_flow_config', JSON.stringify(cfg));
    } catch (e) { }
  }

  function updateFlowStatusPill() {
    const pill = document.getElementById('flowStatusPill');
    const label = pill.querySelector('.status-label');
    if (appConfig.apiKey) {
      label.textContent = 'LLM Flow (API Hazır)';
      pill.style.borderColor = 'rgba(16, 185, 129, 0.4)';
    } else {
      label.textContent = 'LLM Suite Flow';
      pill.style.borderColor = 'var(--border-subtle)';
    }
  }

  // --------------------------------------------------------------------------
  // Utility: Markdown to Simple Safe HTML
  // --------------------------------------------------------------------------
  function markdownToHtml(md) {
    if (!md) return '';
    let html = escapeHtml(md);

    // Headers
    html = html.replace(/^### (.*$)/gim, '<h3>$1</h3>');
    html = html.replace(/^## (.*$)/gim, '<h2>$1</h2>');
    html = html.replace(/^# (.*$)/gim, '<h1>$1</h1>');

    // Bold & Italic
    html = html.replace(/\*\*(.*?)\*\*/gim, '<strong>$1</strong>');
    html = html.replace(/\*(.*?)\*/gim, '<em>$1</em>');

    // Code blocks & inline code
    html = html.replace(/`([^`]+)`/gim, '<code style="background: rgba(255,255,255,0.08); padding: 2px 6px; border-radius: 4px; font-family: var(--font-mono); font-size: 12px;">$1</code>');

    // Lists
    html = html.replace(/^\s*[-•*]\s+(.*$)/gim, '<li>$1</li>');
    html = html.replace(/(<li>.*<\/li>)/s, '<ul>$1</ul>');

    // Numbered lists
    html = html.replace(/^\s*\d+\.\s+(.*$)/gim, '<li>$1</li>');

    // Line breaks to paragraphs
    html = html.split('\n\n').map(p => {
      p = p.trim();
      if (!p) return '';
      if (p.startsWith('<h1>') || p.startsWith('<h2>') || p.startsWith('<h3>') || p.startsWith('<ul>') || p.startsWith('<ol>')) {
        return p;
      }
      return `<p>${p.replace(/\n/g, '<br>')}</p>`;
    }).join('');

    return html;
  }

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  function generateUUID() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      const r = Math.random() * 16 | 0;
      const v = c === 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }

  // Run on DOM ready
  document.addEventListener('DOMContentLoaded', init);
})();
