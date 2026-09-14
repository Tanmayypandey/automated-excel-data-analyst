/**
 * Automated Excel Data Analyst - Frontend JavaScript
 * Handles drag & drop, file uploads, progressive stepper feedback,
 * dynamic KPI rendering, and Excel report download generation.
 */

document.addEventListener('DOMContentLoaded', () => {
  // DOM Elements
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const browseBtn = document.getElementById('browseBtn');
  const fileDetailsBar = document.getElementById('fileDetailsBar');
  const fileNameLabel = document.getElementById('fileNameLabel');
  const fileSizeLabel = document.getElementById('fileSizeLabel');
  const fileTypeBadge = document.getElementById('fileTypeBadge');
  const clearFileBtn = document.getElementById('clearFileBtn');
  const analyzeBtn = document.getElementById('analyzeBtn');

  const alertBox = document.getElementById('alertBox');
  const alertTitle = document.getElementById('alertTitle');
  const alertMessage = document.getElementById('alertMessage');
  const alertClose = document.getElementById('alertClose');

  const loadingCard = document.getElementById('loadingCard');
  const loadingStepText = document.getElementById('loadingStepText');
  const resultsSection = document.getElementById('resultsSection');

  // Results elements
  const resultFileTitle = document.getElementById('resultFileTitle');
  const metaFileType = document.getElementById('metaFileType');
  const metaSheetName = document.getElementById('metaSheetName');
  const metaChartsCount = document.getElementById('metaChartsCount');
  const downloadBtn = document.getElementById('downloadBtn');
  const kpiCardsContainer = document.getElementById('kpiCardsContainer');
  const chartsContainer = document.getElementById('chartsContainer');
  const auditList = document.getElementById('auditList');

  // Stats elements
  const statOrigRows = document.getElementById('statOrigRows');
  const statCleanRows = document.getElementById('statCleanRows');
  const statCols = document.getElementById('statCols');
  const statDupes = document.getElementById('statDupes');
  const statNullsBefore = document.getElementById('statNullsBefore');
  const statNullsAfter = document.getElementById('statNullsAfter');
  const statPreservedIds = document.getElementById('statPreservedIds');

  let currentFile = null;
  let stepperInterval = null;

  // Stepper messages
  const steps = [
    { text: 'Validating file format and size...', step: 1 },
    { text: 'Cleaning data & normalizing types...', step: 2 },
    { text: 'Inferring semantic roles & detecting KPIs...', step: 3 },
    { text: 'Selecting optimal chart configurations...', step: 4 },
    { text: 'Building native Excel dashboard & report...', step: 5 }
  ];

  // Helper: Format bytes
  function formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  // Show Notification
  function showAlert(title, message, isError = true) {
    alertTitle.textContent = title;
    alertMessage.textContent = message;
    if (isError) {
      alertBox.style.backgroundColor = '#fef2f2';
      alertBox.style.borderColor = '#fecaca';
      alertBox.style.color = '#991b1b';
    } else {
      alertBox.style.backgroundColor = '#ecfdf5';
      alertBox.style.borderColor = '#a7f3d0';
      alertBox.style.color = '#065f46';
    }
    alertBox.classList.remove('hidden');
    alertBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  function hideAlert() {
    alertBox.classList.add('hidden');
  }

  alertClose.addEventListener('click', hideAlert);

  // File Selection Handler
  function handleFileSelected(file) {
    hideAlert();
    if (!file) return;

    const ext = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();
    if (ext !== '.xlsx' && ext !== '.csv') {
      showAlert('Unsupported File Type', 'Please select a valid Microsoft Excel (.xlsx) or CSV (.csv) file.');
      clearSelectedFile();
      return;
    }

    if (file.size === 0) {
      showAlert('Empty File', 'The selected file is 0 bytes.');
      clearSelectedFile();
      return;
    }

    if (file.size > 100 * 1024 * 1024) {
      showAlert('File Too Large', 'File exceeds the maximum allowable limit of 100 MB.');
      clearSelectedFile();
      return;
    }

    currentFile = file;
    fileNameLabel.textContent = file.name;
    fileSizeLabel.textContent = formatBytes(file.size);
    fileTypeBadge.textContent = ext.replace('.', '').toUpperCase();

    fileDetailsBar.classList.remove('hidden');
    analyzeBtn.removeAttribute('disabled');
  }

  function clearSelectedFile() {
    currentFile = null;
    fileInput.value = '';
    fileDetailsBar.classList.add('hidden');
    analyzeBtn.setAttribute('disabled', 'true');
  }

  // File Input Listeners
  browseBtn.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('click', (e) => {
    if (e.target !== browseBtn) fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFileSelected(e.target.files[0]);
    }
  });

  clearFileBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    clearSelectedFile();
  });

  // Drag & Drop
  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add('dragover');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('dragover');
    }, false);
  });

  dropZone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt.files && dt.files.length > 0) {
      handleFileSelected(dt.files[0]);
    }
  });

  // Stepper Controller
  function startProgressStepper() {
    loadingCard.classList.remove('hidden');
    resultsSection.classList.add('hidden');
    analyzeBtn.setAttribute('disabled', 'true');

    // Reset step styles
    for (let i = 1; i <= 5; i++) {
      const stepEl = document.getElementById(`step${i}`);
      stepEl.classList.remove('active', 'completed');
    }

    let currentIndex = 0;
    const updateStepUI = (index) => {
      const item = steps[index];
      loadingStepText.textContent = item.text;
      for (let i = 1; i <= 5; i++) {
        const stepEl = document.getElementById(`step${i}`);
        if (i < item.step) {
          stepEl.classList.remove('active');
          stepEl.classList.add('completed');
        } else if (i === item.step) {
          stepEl.classList.add('active');
          stepEl.classList.remove('completed');
        } else {
          stepEl.classList.remove('active', 'completed');
        }
      }
    };

    updateStepUI(0);
    stepperInterval = setInterval(() => {
      if (currentIndex < steps.length - 1) {
        currentIndex++;
        updateStepUI(currentIndex);
      }
    }, 1500);
  }

  function stopProgressStepper() {
    clearInterval(stepperInterval);
    loadingCard.classList.add('hidden');
    analyzeBtn.removeAttribute('disabled');
  }

  // Analyze Action
  analyzeBtn.addEventListener('click', async () => {
    if (!currentFile) return;

    hideAlert();
    startProgressStepper();

    const formData = new FormData();
    formData.append('file', currentFile);

    try {
      const response = await fetch('/analyze', {
        method: 'POST',
        body: formData
      });

      const data = await response.json();

      stopProgressStepper();

      if (!response.ok || data.status !== 'success') {
        const errorMsg = data.detail || data.error || data.message || 'Data analysis failed.';
        showAlert('Analysis Error', errorMsg, true);
        return;
      }

      // Populate & Display Results
      renderResults(data);

    } catch (err) {
      stopProgressStepper();
      showAlert('Network/Server Error', 'Failed to connect to the analysis backend. Please verify the server is running.', true);
    }
  });

  // Render Pipeline Results
  function renderResults(res) {
    resultFileTitle.textContent = `${res.original_filename} Analysis`;
    metaFileType.textContent = `Type: ${res.input_type || 'Spreadsheet'}`;
    metaSheetName.textContent = `Sheet: ${res.selected_sheet || 'Raw Data'}`;
    metaChartsCount.textContent = `Charts: ${res.charts_generated || 0} Native Charts`;

    downloadBtn.href = res.download_url;
    downloadBtn.setAttribute('download', res.output_filename);

    // KPI Cards
    kpiCardsContainer.innerHTML = '';
    const kpis = res.kpis || [];
    kpis.forEach(kpi => {
      const card = document.createElement('div');
      card.className = 'kpi-card';
      card.innerHTML = `
        <div class="kpi-top-row">
          <span class="kpi-label">${escapeHtml(kpi.label)}</span>
          <span class="kpi-badge">${escapeHtml(kpi.aggregation || 'metric')}</span>
        </div>
        <div class="kpi-value">${escapeHtml(kpi.formatted_value || String(kpi.value))}</div>
      `;
      kpiCardsContainer.appendChild(card);
    });

    // Summary Stats
    const cleanRep = res.cleaning_report || {};
    const valRep = res.validation_report || {};
    statOrigRows.textContent = (cleanRep.original_rows || valRep.total_rows || 0).toLocaleString();
    statCleanRows.textContent = (cleanRep.cleaned_rows || 0).toLocaleString();
    statCols.textContent = (cleanRep.cleaned_columns || valRep.total_columns || 0).toLocaleString();
    statDupes.textContent = (cleanRep.duplicates_removed || 0).toLocaleString();

    statNullsBefore.textContent = (cleanRep.missing_values_before || 0).toLocaleString();
    statNullsAfter.textContent = `${(cleanRep.missing_values_after || 0).toLocaleString()} (Cleaned/Imputed)`;

    const preservedIds = cleanRep.preserved_id_columns || [];
    statPreservedIds.textContent = preservedIds.length > 0 ? preservedIds.join(', ') : 'None';

    // Charts
    chartsContainer.innerHTML = '';
    const charts = res.selected_charts || [];
    charts.forEach(chart => {
      const cEl = document.createElement('div');
      cEl.className = 'chart-card';
      cEl.innerHTML = `
        <div class="chart-card-header">
          <span class="chart-type-tag">${escapeHtml(chart.chart_type)}</span>
        </div>
        <h4 class="chart-card-title">${escapeHtml(chart.title)}</h4>
        <p class="chart-card-desc">${escapeHtml(chart.description || '')}</p>
      `;
      chartsContainer.appendChild(cEl);
    });

    // Audit Log
    auditList.innerHTML = '';
    const actions = cleanRep.major_actions || [];
    if (actions.length > 0) {
      actions.forEach(act => {
        const li = document.createElement('li');
        li.className = 'audit-item';
        li.textContent = act;
        auditList.appendChild(li);
      });
    } else {
      const li = document.createElement('li');
      li.className = 'audit-item';
      li.textContent = 'Standardized column headers and validated cell formatting.';
      auditList.appendChild(li);
    }

    resultsSection.classList.remove('hidden');
    resultsSection.scrollIntoView({ behavior: 'smooth' });
  }

  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }
});
