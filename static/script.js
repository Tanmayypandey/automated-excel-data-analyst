/**
 * Automated Excel Data Analyst - Frontend JavaScript
 *
 * Handles:
 * - Drag & drop
 * - CSV/XLSX upload
 * - File validation
 * - Processing UI
 * - KPI rendering
 * - Chart information
 * - Cleaning report
 * - Excel report download
 */

document.addEventListener('DOMContentLoaded', () => {

  // =========================================================
  // API CONFIGURATION
  // =========================================================

  /**
   * IMPORTANT:
   *
   * During local development:
   * window.location.origin = http://127.0.0.1:8000
   *
   * This prevents requests from accidentally going to
   * the old Render deployment.
   */
  const API_BASE_URL = window.location.origin;

  console.log('API Base URL:', API_BASE_URL);


  // =========================================================
  // DOM ELEMENTS
  // =========================================================

  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const browseBtn = document.getElementById('browseBtn');

  const fileDetailsBar =
    document.getElementById('fileDetailsBar');

  const fileNameLabel =
    document.getElementById('fileNameLabel');

  const fileSizeLabel =
    document.getElementById('fileSizeLabel');

  const fileTypeBadge =
    document.getElementById('fileTypeBadge');

  const clearFileBtn =
    document.getElementById('clearFileBtn');

  const analyzeBtn =
    document.getElementById('analyzeBtn');


  // Alerts
  const alertBox =
    document.getElementById('alertBox');

  const alertTitle =
    document.getElementById('alertTitle');

  const alertMessage =
    document.getElementById('alertMessage');

  const alertClose =
    document.getElementById('alertClose');


  // Loading
  const loadingCard =
    document.getElementById('loadingCard');

  const loadingStepText =
    document.getElementById('loadingStepText');


  // Results
  const resultsSection =
    document.getElementById('resultsSection');

  const resultFileTitle =
    document.getElementById('resultFileTitle');

  const metaFileType =
    document.getElementById('metaFileType');

  const metaSheetName =
    document.getElementById('metaSheetName');

  const metaChartsCount =
    document.getElementById('metaChartsCount');

  const downloadBtn =
    document.getElementById('downloadBtn');

  const kpiCardsContainer =
    document.getElementById('kpiCardsContainer');

  const chartsContainer =
    document.getElementById('chartsContainer');

  const auditList =
    document.getElementById('auditList');


  // Statistics
  const statOrigRows =
    document.getElementById('statOrigRows');

  const statCleanRows =
    document.getElementById('statCleanRows');

  const statCols =
    document.getElementById('statCols');

  const statDupes =
    document.getElementById('statDupes');

  const statNullsBefore =
    document.getElementById('statNullsBefore');

  const statNullsAfter =
    document.getElementById('statNullsAfter');

  const statPreservedIds =
    document.getElementById('statPreservedIds');


  // =========================================================
  // STATE
  // =========================================================

  let currentFile = null;
  let stepperInterval = null;


  // =========================================================
  // PROCESSING STEPS
  // =========================================================

  const steps = [

    {
      text: 'Validating file format and size...',
      step: 1
    },

    {
      text: 'Cleaning data & normalizing types...',
      step: 2
    },

    {
      text: 'Inferring semantic roles & detecting KPIs...',
      step: 3
    },

    {
      text: 'Selecting optimal chart configurations...',
      step: 4
    },

    {
      text: 'Building native Excel dashboard & report...',
      step: 5
    }

  ];


  // =========================================================
  // FORMAT BYTES
  // =========================================================

  function formatBytes(bytes) {

    if (bytes === 0) {
      return '0 Bytes';
    }

    const k = 1024;

    const sizes = [
      'Bytes',
      'KB',
      'MB',
      'GB'
    ];

    const i = Math.floor(
      Math.log(bytes) / Math.log(k)
    );

    return (
      parseFloat(
        (
          bytes /
          Math.pow(k, i)
        ).toFixed(2)
      )
      +
      ' '
      +
      sizes[i]
    );
  }


  // =========================================================
  // ALERTS
  // =========================================================

  function showAlert(
    title,
    message,
    isError = true
  ) {

    alertTitle.textContent = title;
    alertMessage.textContent = message;

    if (isError) {

      alertBox.style.backgroundColor =
        '#fef2f2';

      alertBox.style.borderColor =
        '#fecaca';

      alertBox.style.color =
        '#991b1b';

    } else {

      alertBox.style.backgroundColor =
        '#ecfdf5';

      alertBox.style.borderColor =
        '#a7f3d0';

      alertBox.style.color =
        '#065f46';
    }

    alertBox.classList.remove('hidden');

    alertBox.scrollIntoView({
      behavior: 'smooth',
      block: 'nearest'
    });
  }


  function hideAlert() {

    alertBox.classList.add('hidden');
  }


  if (alertClose) {

    alertClose.addEventListener(
      'click',
      hideAlert
    );
  }


  // =========================================================
  // FILE SELECTION
  // =========================================================

  function handleFileSelected(file) {

    hideAlert();

    if (!file) {
      return;
    }

    const lastDot =
      file.name.lastIndexOf('.');

    const ext =
      lastDot >= 0
        ? file.name.slice(lastDot).toLowerCase()
        : '';


    // -------------------------------------------------------
    // TYPE VALIDATION
    // -------------------------------------------------------

    if (
      ext !== '.xlsx' &&
      ext !== '.csv'
    ) {

      showAlert(
        'Unsupported File Type',
        'Please select a valid Microsoft Excel (.xlsx) or CSV (.csv) file.'
      );

      clearSelectedFile();

      return;
    }


    // -------------------------------------------------------
    // EMPTY FILE
    // -------------------------------------------------------

    if (file.size === 0) {

      showAlert(
        'Empty File',
        'The selected file is empty.'
      );

      clearSelectedFile();

      return;
    }


    // -------------------------------------------------------
    // PRACTICAL LOCAL LIMITS
    // -------------------------------------------------------

    const CSV_MAX_SIZE =
      30 * 1024 * 1024;

    const XLSX_MAX_SIZE =
      15 * 1024 * 1024;


    if (
      ext === '.csv' &&
      file.size > CSV_MAX_SIZE
    ) {

      showAlert(
        'CSV File Too Large',
        'For the current version, CSV files should be 30 MB or smaller.'
      );

      clearSelectedFile();

      return;
    }


    if (
      ext === '.xlsx' &&
      file.size > XLSX_MAX_SIZE
    ) {

      showAlert(
        'Excel File Too Large',
        'For the current version, Excel files should be 15 MB or smaller.'
      );

      clearSelectedFile();

      return;
    }


    // -------------------------------------------------------
    // ACCEPT FILE
    // -------------------------------------------------------

    currentFile = file;

    fileNameLabel.textContent =
      file.name;

    fileSizeLabel.textContent =
      formatBytes(file.size);

    fileTypeBadge.textContent =
      ext
        .replace('.', '')
        .toUpperCase();

    fileDetailsBar.classList.remove(
      'hidden'
    );

    analyzeBtn.removeAttribute(
      'disabled'
    );
  }


  // =========================================================
  // CLEAR FILE
  // =========================================================

  function clearSelectedFile() {

    currentFile = null;

    fileInput.value = '';

    fileDetailsBar.classList.add(
      'hidden'
    );

    analyzeBtn.setAttribute(
      'disabled',
      'true'
    );
  }


  // =========================================================
  // FILE INPUT EVENTS
  // =========================================================

  browseBtn.addEventListener(
    'click',
    (event) => {

      event.stopPropagation();

      fileInput.click();
    }
  );


  dropZone.addEventListener(
    'click',
    (event) => {

      if (event.target !== browseBtn) {

        fileInput.click();
      }
    }
  );


  fileInput.addEventListener(
    'change',
    (event) => {

      if (
        event.target.files &&
        event.target.files.length > 0
      ) {

        handleFileSelected(
          event.target.files[0]
        );
      }
    }
  );


  clearFileBtn.addEventListener(
    'click',
    (event) => {

      event.stopPropagation();

      clearSelectedFile();
    }
  );


  // =========================================================
  // DRAG AND DROP
  // =========================================================

  [
    'dragenter',
    'dragover'
  ].forEach(eventName => {

    dropZone.addEventListener(
      eventName,
      (event) => {

        event.preventDefault();
        event.stopPropagation();

        dropZone.classList.add(
          'dragover'
        );
      },
      false
    );
  });


  [
    'dragleave',
    'drop'
  ].forEach(eventName => {

    dropZone.addEventListener(
      eventName,
      (event) => {

        event.preventDefault();
        event.stopPropagation();

        dropZone.classList.remove(
          'dragover'
        );
      },
      false
    );
  });


  dropZone.addEventListener(
    'drop',
    (event) => {

      const dt =
        event.dataTransfer;

      if (
        dt.files &&
        dt.files.length > 0
      ) {

        handleFileSelected(
          dt.files[0]
        );
      }
    }
  );


  // =========================================================
  // PROGRESS STEPPER
  // =========================================================

  function startProgressStepper() {

    loadingCard.classList.remove(
      'hidden'
    );

    resultsSection.classList.add(
      'hidden'
    );

    analyzeBtn.setAttribute(
      'disabled',
      'true'
    );


    // Reset all steps
    for (
      let i = 1;
      i <= 5;
      i++
    ) {

      const stepElement =
        document.getElementById(
          `step${i}`
        );

      if (stepElement) {

        stepElement.classList.remove(
          'active',
          'completed'
        );
      }
    }


    let currentIndex = 0;


    function updateStepUI(index) {

      const item =
        steps[index];

      loadingStepText.textContent =
        item.text;


      for (
        let i = 1;
        i <= 5;
        i++
      ) {

        const stepElement =
          document.getElementById(
            `step${i}`
          );

        if (!stepElement) {
          continue;
        }


        if (i < item.step) {

          stepElement.classList.remove(
            'active'
          );

          stepElement.classList.add(
            'completed'
          );

        } else if (
          i === item.step
        ) {

          stepElement.classList.add(
            'active'
          );

          stepElement.classList.remove(
            'completed'
          );

        } else {

          stepElement.classList.remove(
            'active',
            'completed'
          );
        }
      }
    }


    updateStepUI(0);


    stepperInterval =
      setInterval(() => {

        if (
          currentIndex <
          steps.length - 1
        ) {

          currentIndex++;

          updateStepUI(
            currentIndex
          );
        }

      }, 1500);
  }


  function stopProgressStepper() {

    if (stepperInterval) {

      clearInterval(
        stepperInterval
      );

      stepperInterval = null;
    }

    loadingCard.classList.add(
      'hidden'
    );

    analyzeBtn.removeAttribute(
      'disabled'
    );
  }


  // =========================================================
  // ANALYZE FILE
  // =========================================================

  analyzeBtn.addEventListener(
    'click',
    async () => {

      if (!currentFile) {
        return;
      }

      hideAlert();

      startProgressStepper();


      const formData =
        new FormData();

      formData.append(
        'file',
        currentFile
      );


      try {

        console.log(
          'Uploading to:',
          `${API_BASE_URL}/analyze`
        );


        const response =
          await fetch(
            `${API_BASE_URL}/analyze`,
            {
              method: 'POST',
              body: formData
            }
          );


        let data;

        try {

          data =
            await response.json();

        } catch (jsonError) {

          throw new Error(
            'The backend returned an invalid response.'
          );
        }


        stopProgressStepper();


        if (
          !response.ok ||
          data.status !== 'success'
        ) {

          const errorMessage =
            data.detail ||
            data.error ||
            data.message ||
            'Data analysis failed.';


          showAlert(
            'Analysis Error',
            errorMessage,
            true
          );

          return;
        }


        console.log(
          'Analysis completed:',
          data
        );


        renderResults(data);


      } catch (error) {

        console.error(
          'Analysis request failed:',
          error
        );


        stopProgressStepper();


        showAlert(
          'Network/Server Error',
          error.message ||
          'Failed to connect to the analysis backend. Please try again.',
          true
        );
      }
    }
  );


  // =========================================================
  // RENDER RESULTS
  // =========================================================

  function renderResults(res) {

    resultFileTitle.textContent =
      `${res.original_filename || 'File'} Analysis`;


    metaFileType.textContent =
      `Type: ${res.input_type || 'Spreadsheet'}`;


    metaSheetName.textContent =
      `Sheet: ${res.selected_sheet || 'Raw Data'}`;


    metaChartsCount.textContent =
      `Charts: ${res.charts_generated || 0} Native Charts`;


    // =====================================================
    // DOWNLOAD URL
    // =====================================================

    if (res.download_url) {

      if (
        res.download_url.startsWith(
          'http://'
        ) ||
        res.download_url.startsWith(
          'https://'
        )
      ) {

        downloadBtn.href =
          res.download_url;

      } else {

        const relativePath =
          res.download_url.startsWith('/')
            ? res.download_url
            : `/${res.download_url}`;


        downloadBtn.href =
          `${API_BASE_URL}${relativePath}`;
      }


      console.log(
        'Download URL:',
        downloadBtn.href
      );
    }


    if (res.output_filename) {

      downloadBtn.setAttribute(
        'download',
        res.output_filename
      );
    }


    // =====================================================
    // KPI CARDS
    // =====================================================

    kpiCardsContainer.innerHTML =
      '';


    const kpis =
      res.kpis || [];


    kpis.forEach(kpi => {

      const card =
        document.createElement(
          'div'
        );


      card.className =
        'kpi-card';


      card.innerHTML = `
        <div class="kpi-top-row">

          <span class="kpi-label">
            ${escapeHtml(
              kpi.label || ''
            )}
          </span>

          <span class="kpi-badge">
            ${escapeHtml(
              kpi.aggregation ||
              'metric'
            )}
          </span>

        </div>

        <div class="kpi-value">

          ${escapeHtml(
            kpi.formatted_value ??
            kpi.value ??
            ''
          )}

        </div>
      `;


      kpiCardsContainer.appendChild(
        card
      );
    });


    // =====================================================
    // CLEANING STATISTICS
    // =====================================================

    const cleanRep =
      res.cleaning_report || {};


    const valRep =
      res.validation_report || {};


    const originalRows =
      cleanRep.original_rows ??
      valRep.total_rows ??
      0;


    const cleanedRows =
      cleanRep.cleaned_rows ??
      0;


    const columns =
      cleanRep.cleaned_columns ??
      valRep.total_columns ??
      0;


    const duplicates =
      cleanRep.duplicates_removed ??
      0;


    const missingBefore =
      cleanRep.missing_values_before ??
      0;


    const missingAfter =
      cleanRep.missing_values_after ??
      0;


    statOrigRows.textContent =
      Number(
        originalRows
      ).toLocaleString();


    statCleanRows.textContent =
      Number(
        cleanedRows
      ).toLocaleString();


    statCols.textContent =
      Number(
        columns
      ).toLocaleString();


    statDupes.textContent =
      Number(
        duplicates
      ).toLocaleString();


    statNullsBefore.textContent =
      Number(
        missingBefore
      ).toLocaleString();


    statNullsAfter.textContent =
      `${Number(
        missingAfter
      ).toLocaleString()} (Remaining)`;


    const preservedIds =
      cleanRep.preserved_id_columns ||
      [];


    statPreservedIds.textContent =
      preservedIds.length > 0
        ? preservedIds.join(', ')
        : 'None';


    // =====================================================
    // CHART INFORMATION
    // =====================================================

    chartsContainer.innerHTML =
      '';


    const charts =
      res.selected_charts || [];


    charts.forEach(chart => {

      const chartElement =
        document.createElement(
          'div'
        );


      chartElement.className =
        'chart-card';


      chartElement.innerHTML = `
        <div class="chart-card-header">

          <span class="chart-type-tag">

            ${escapeHtml(
              chart.chart_type ||
              'chart'
            )}

          </span>

        </div>

        <h4 class="chart-card-title">

          ${escapeHtml(
            chart.title ||
            'Chart'
          )}

        </h4>

        <p class="chart-card-desc">

          ${escapeHtml(
            chart.description ||
            ''
          )}

        </p>
      `;


      chartsContainer.appendChild(
        chartElement
      );
    });


    // =====================================================
    // AUDIT LOG
    // =====================================================

    auditList.innerHTML =
      '';


    const actions =
      cleanRep.major_actions ||
      [];


    if (actions.length > 0) {

      actions.forEach(action => {

        const li =
          document.createElement(
            'li'
          );


        li.className =
          'audit-item';


        li.textContent =
          action;


        auditList.appendChild(
          li
        );
      });

    } else {

      const li =
        document.createElement(
          'li'
        );


      li.className =
        'audit-item';


      li.textContent =
        'Dataset validated and standardized successfully.';


      auditList.appendChild(
        li
      );
    }


    // =====================================================
    // DISPLAY RESULTS
    // =====================================================

    resultsSection.classList.remove(
      'hidden'
    );


    resultsSection.scrollIntoView({
      behavior: 'smooth',
      block: 'start'
    });
  }


  // =========================================================
  // HTML ESCAPING
  // =========================================================

  function escapeHtml(value) {

    if (
      value === null ||
      value === undefined
    ) {

      return '';
    }


    return String(value)

      .replace(
        /&/g,
        '&amp;'
      )

      .replace(
        /</g,
        '&lt;'
      )

      .replace(
        />/g,
        '&gt;'
      )

      .replace(
        /"/g,
        '&quot;'
      )

      .replace(
        /'/g,
        '&#039;'
      );
  }

});