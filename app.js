/**
 * Statistical Data Analysis Dashboard - Client-Side Engine
 * 100% Offline & Client-Side Execution for Data Privacy
 */

// Global State
const appState = {
  rawRows: [],
  columns: [],
  numericColumns: [],
  timeColumn: null,
  fileName: "Random_Dummy_Data.xlsx",
  currentTab: "eda",
  tableSearch: "",
  tablePage: 1,
  tablePageSize: 15,
  charts: {},
  corrMethod: "pearson"
};

// Initialize Application
document.addEventListener("DOMContentLoaded", async () => {
  // Render Lucide icons
  if (window.lucide) {
    lucide.createIcons();
  }

  // Setup UI event listeners
  setupEventListeners();

  // Attempt to load default Random_Dummy_Data.xlsx
  await loadDefaultDataset();
});

// Setup Event Listeners
function setupEventListeners() {
  // Tab Navigation
  document.querySelectorAll(".nav-tab").forEach(tab => {
    tab.addEventListener("click", () => {
      document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
      tab.classList.add("active");
      const targetTab = tab.getAttribute("data-tab");
      switchTab(targetTab);
    });
  });

  // File Input
  const fileInput = document.getElementById("fileInput");
  fileInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files[0]) {
      handleFile(e.target.files[0]);
    }
  });

  // Sample Reload Button
  document.getElementById("btnReloadSample").addEventListener("click", () => {
    loadDefaultDataset();
  });

  // Print/Report Button
  document.getElementById("btnExportReport").addEventListener("click", () => {
    window.print();
  });

  // Drag & Drop
  window.addEventListener("dragover", (e) => {
    e.preventDefault();
    document.getElementById("dropZoneOverlay").classList.remove("hidden");
  });
  window.addEventListener("dragleave", (e) => {
    if (e.relatedTarget === null) {
      document.getElementById("dropZoneOverlay").classList.add("hidden");
    }
  });
  window.addEventListener("drop", (e) => {
    e.preventDefault();
    document.getElementById("dropZoneOverlay").classList.add("hidden");
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFile(e.dataTransfer.files[0]);
    }
  });

  // Variable Selectors Change Listeners
  document.getElementById("distVarSelect")?.addEventListener("change", renderDistributionTab);
  document.getElementById("distBinSelect")?.addEventListener("change", renderDistributionTab);
  
  document.getElementById("btnCorrPearson")?.addEventListener("click", () => {
    appState.corrMethod = "pearson";
    updateCorrButtonStyles();
    renderCorrelationTab();
  });
  document.getElementById("btnCorrSpearman")?.addEventListener("click", () => {
    appState.corrMethod = "spearman";
    updateCorrButtonStyles();
    renderCorrelationTab();
  });

  document.getElementById("regVarX")?.addEventListener("change", updateRegressionPlot);
  document.getElementById("regVarY")?.addEventListener("change", updateRegressionPlot);

  document.getElementById("tsVarSelect")?.addEventListener("change", renderTimeSeriesTab);
  document.getElementById("tsWindowInput")?.addEventListener("change", renderTimeSeriesTab);

  document.getElementById("spcVarSelect")?.addEventListener("change", renderSPCTab);

  document.getElementById("anovaVarSelect")?.addEventListener("change", renderAnovaTab);
  document.getElementById("anovaSplitSelect")?.addEventListener("change", renderAnovaTab);

  document.getElementById("clusterKSelect")?.addEventListener("change", renderClusterTab);

  // Table Search & Pagination
  document.getElementById("tableSearchInput")?.addEventListener("input", (e) => {
    appState.tableSearch = e.target.value.toLowerCase();
    appState.tablePage = 1;
    renderDataTable();
  });
  document.getElementById("btnPrevPage")?.addEventListener("click", () => {
    if (appState.tablePage > 1) {
      appState.tablePage--;
      renderDataTable();
    }
  });
  document.getElementById("btnNextPage")?.addEventListener("click", () => {
    const totalPages = Math.ceil(getFilteredRows().length / appState.tablePageSize);
    if (appState.tablePage < totalPages) {
      appState.tablePage++;
      renderDataTable();
    }
  });

  // Export CSV Buttons
  document.getElementById("btnExportDataCsv")?.addEventListener("click", exportRawDataCsv);
  document.getElementById("btnExportStatsCsv")?.addEventListener("click", exportStatsCsv);
}

function updateCorrButtonStyles() {
  const pBtn = document.getElementById("btnCorrPearson");
  const sBtn = document.getElementById("btnCorrSpearman");
  if (appState.corrMethod === "pearson") {
    pBtn.className = "px-3 py-1.5 rounded-lg text-xs font-semibold bg-indigo-600 text-white";
    sBtn.className = "px-3 py-1.5 rounded-lg text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-300";
  } else {
    sBtn.className = "px-3 py-1.5 rounded-lg text-xs font-semibold bg-indigo-600 text-white";
    pBtn.className = "px-3 py-1.5 rounded-lg text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-300";
  }
}

// Switch Active Tab
function switchTab(tabId) {
  appState.currentTab = tabId;
  document.querySelectorAll(".tab-pane").forEach(pane => pane.classList.add("hidden"));
  const targetPane = document.getElementById(`tab-${tabId}`);
  if (targetPane) {
    targetPane.classList.remove("hidden");
  }

  // Trigger re-render of chart if necessary for correct sizing
  setTimeout(() => {
    switch (tabId) {
      case "eda": renderEdaTab(); break;
      case "distribution": renderDistributionTab(); break;
      case "correlation": renderCorrelationTab(); break;
      case "timeseries": renderTimeSeriesTab(); break;
      case "spc": renderSPCTab(); break;
      case "anova": renderAnovaTab(); break;
      case "cluster": renderClusterTab(); break;
      case "data-view": renderDataTable(); break;
    }
  }, 50);
}

// Load Default Dataset (Random_Dummy_Data.xlsx or Embedded Sample Data)
async function loadDefaultDataset() {
  // If embedded data is preloaded, use it directly (works flawlessly when opening index.html directly via file:///)
  if (window.EMBEDDED_SAMPLE_DATA && window.EMBEDDED_SAMPLE_DATA.length > 0) {
    loadJsonData(window.EMBEDDED_SAMPLE_DATA, "Random_Dummy_Data.xlsx");
    return;
  }

  try {
    const response = await fetch("Random_Dummy_Data.xlsx");
    if (!response.ok) throw new Error("Default file fetch failed");
    const arrayBuffer = await response.arrayBuffer();
    parseExcelArrayBuffer(arrayBuffer, "Random_Dummy_Data.xlsx");
  } catch (err) {
    console.warn("Auto-load from server/local path failed, waiting for user file drop", err);
  }
}

// Load JSON data directly into appState
function loadJsonData(json, fileName) {
  if (!json || json.length === 0) {
    alert("데이터를 읽을 수 없거나 파일이 비어 있습니다.");
    return;
  }

  appState.fileName = fileName;
  appState.rawRows = json;
  appState.columns = Object.keys(json[0]);

  // Detect numeric columns & time columns
  appState.numericColumns = appState.columns.filter(col => {
    const nonNulls = json.map(r => r[col]).filter(v => v !== null && v !== "");
    if (nonNulls.length === 0) return false;
    const numCount = nonNulls.filter(v => !isNaN(Number(v))).length;
    return numCount / nonNulls.length >= 0.8;
  });

  appState.timeColumn = appState.columns.find(c => 
    c.toLowerCase().includes("time") || 
    c.includes("시각") || 
    c.includes("일시") || 
    c.includes("날짜")
  ) || appState.columns[0];

  // Update UI Badges
  document.getElementById("loadedFileName").textContent = appState.fileName;
  document.getElementById("loadedRowCount").textContent = `${appState.rawRows.length}행`;
  document.getElementById("fileInfoBadge").classList.remove("hidden");

  // Populate Select Elements
  populateSelects();

  // Render Global Metrics
  renderGlobalMetrics();

  // Re-render current active tab
  switchTab(appState.currentTab);
}

// File Handler
function handleFile(file) {
  appState.fileName = file.name;
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    parseExcelArrayBuffer(data, file.name);
  };
  reader.readAsArrayBuffer(file);
}

// Parse Excel or CSV using SheetJS
function parseExcelArrayBuffer(data, fileName) {
  try {
    const workbook = XLSX.read(data, { type: "array", cellDates: true });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const json = XLSX.utils.sheet_to_json(worksheet, { defval: null });

    if (!json || json.length === 0) {
      alert("데이터를 읽을 수 없거나 파일이 비어 있습니다.");
      return;
    }

    appState.fileName = fileName;
    appState.rawRows = json;
    appState.columns = Object.keys(json[0]);

    // Detect numeric columns & time columns
    appState.numericColumns = appState.columns.filter(col => {
      // Check if at least 70% of non-null values are numeric and column is not purely an index or date string
      const nonNulls = json.map(r => r[col]).filter(v => v !== null && v !== "");
      if (nonNulls.length === 0) return false;
      const numCount = nonNulls.filter(v => !isNaN(Number(v))).length;
      return numCount / nonNulls.length >= 0.8;
    });

    appState.timeColumn = appState.columns.find(c => 
      c.toLowerCase().includes("time") || 
      c.includes("시각") || 
      c.includes("일시") || 
      c.includes("날짜")
    ) || appState.columns[0];

    // Update UI Badges
    document.getElementById("loadedFileName").textContent = appState.fileName;
    document.getElementById("loadedRowCount").textContent = `${appState.rawRows.length}행`;
    document.getElementById("fileInfoBadge").classList.remove("hidden");

    // Populate Select Elements
    populateSelects();

    // Render Global Metrics
    renderGlobalMetrics();

    // Re-render current active tab
    switchTab(appState.currentTab);
  } catch (err) {
    console.error("Parse Error:", err);
    alert("파일 파싱 중 오류가 발생했습니다: " + err.message);
  }
}

// Populate Select Dropdowns
function populateSelects() {
  const selects = [
    { id: "distVarSelect", cols: appState.numericColumns },
    { id: "regVarX", cols: appState.numericColumns, defaultIdx: 0 },
    { id: "regVarY", cols: appState.numericColumns, defaultIdx: 1 },
    { id: "tsVarSelect", cols: appState.numericColumns },
    { id: "spcVarSelect", cols: appState.numericColumns },
    { id: "anovaVarSelect", cols: appState.numericColumns }
  ];

  selects.forEach(({ id, cols, defaultIdx }) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.innerHTML = "";
    cols.forEach((col, idx) => {
      const opt = document.createElement("option");
      opt.value = col;
      opt.textContent = col;
      if (defaultIdx !== undefined && idx === defaultIdx) {
        opt.selected = true;
      }
      el.appendChild(opt);
    });
  });

  // Ensure regVarY is different if possible
  const regX = document.getElementById("regVarX");
  const regY = document.getElementById("regVarY");
  if (regX && regY && regY.options.length > 1 && regX.value === regY.value) {
    regY.selectedIndex = 1;
  }
}

// Global Metric Calculation
function renderGlobalMetrics() {
  const rows = appState.rawRows;
  const nRows = rows.length;
  document.getElementById("statTotalRows").textContent = nRows.toLocaleString();
  document.getElementById("statNumCols").textContent = `${appState.numericColumns.length}개`;

  // Missing values
  let missing = 0;
  let totalCells = nRows * appState.columns.length;
  rows.forEach(r => {
    appState.columns.forEach(c => {
      if (r[c] === null || r[c] === undefined || r[c] === "") missing++;
    });
  });
  const missingPct = totalCells > 0 ? ((missing / totalCells) * 100).toFixed(1) : 0;
  const missingEl = document.getElementById("statMissingCells");
  missingEl.textContent = `${missing}개 (${missingPct}%)`;
  if (missing > 0) {
    missingEl.className = "text-xl font-bold font-mono text-amber-400 mt-1";
  } else {
    missingEl.className = "text-xl font-bold font-mono text-emerald-400 mt-1";
  }

  // Quick summary for Temp, Humidity, Vibration if available
  const tempCol = appState.numericColumns.find(c => c.includes("온도") || c.toLowerCase().includes("temp"));
  const humCol = appState.numericColumns.find(c => c.includes("습도") || c.toLowerCase().includes("hum"));
  const vibCol = appState.numericColumns.find(c => c.includes("진동") || c.toLowerCase().includes("vib"));

  if (tempCol) {
    const vals = getColValues(tempCol);
    document.getElementById("statTempMean").textContent = ss.mean(vals).toFixed(1) + "°C";
  } else {
    document.getElementById("statTempMean").textContent = "-";
  }

  if (humCol) {
    const vals = getColValues(humCol);
    document.getElementById("statHumMean").textContent = ss.mean(vals).toFixed(1) + "%";
  } else {
    document.getElementById("statHumMean").textContent = "-";
  }

  if (vibCol) {
    const vals = getColValues(vibCol);
    document.getElementById("statVibMean").textContent = ss.mean(vals).toFixed(2) + "Hz";
  } else {
    document.getElementById("statVibMean").textContent = "-";
  }

  if (window.lucide) lucide.createIcons();
}

// Helper: Extract valid numeric values from a column
function getColValues(colName) {
  return appState.rawRows
    .map(r => parseFloat(r[colName]))
    .filter(v => !isNaN(v) && v !== null && v !== undefined);
}

/* ==========================================================================
   MODULE 1: EDA & 요약 통계량
   ========================================================================== */
function renderEdaTab() {
  const tbody = document.getElementById("edaStatsTableBody");
  tbody.innerHTML = "";

  const statsList = [];
  const cvLabels = [];
  const cvValues = [];

  appState.numericColumns.forEach(col => {
    const vals = getColValues(col);
    if (vals.length === 0) return;

    const n = vals.length;
    const mean = ss.mean(vals);
    const sd = vals.length > 1 ? ss.standardDeviation(vals) : 0;
    const median = ss.median(vals);
    const min = ss.min(vals);
    const max = ss.max(vals);
    const q1 = ss.quantile(vals, 0.25);
    const q3 = ss.quantile(vals, 0.75);
    const iqr = q3 - q1;
    const skewness = ss.sampleSkewness ? ss.sampleSkewness(vals) : calculateSkewness(vals, mean, sd);
    const kurtosis = calculateKurtosis(vals, mean, sd);
    const cv = mean !== 0 ? (sd / Math.abs(mean)) * 100 : 0;

    statsList.push({ col, n, mean, sd, median, min, q1, q3, max, iqr, skewness, kurtosis, cv });

    cvLabels.push(col);
    cvValues.push(cv);

    const tr = document.createElement("tr");
    tr.className = "hover:bg-slate-800/40 transition";
    tr.innerHTML = `
      <td class="py-2.5 px-3 font-semibold text-indigo-300">${col}</td>
      <td class="py-2.5 px-2 text-right">${n}</td>
      <td class="py-2.5 px-2 text-right text-emerald-300 font-bold">${formatNum(mean)}</td>
      <td class="py-2.5 px-2 text-right text-slate-300">${formatNum(sd)}</td>
      <td class="py-2.5 px-2 text-right">${formatNum(median)}</td>
      <td class="py-2.5 px-2 text-right text-slate-400">${formatNum(min)}</td>
      <td class="py-2.5 px-2 text-right">${formatNum(q1)}</td>
      <td class="py-2.5 px-2 text-right">${formatNum(q3)}</td>
      <td class="py-2.5 px-2 text-right text-slate-400">${formatNum(max)}</td>
      <td class="py-2.5 px-2 text-right text-amber-300">${formatNum(iqr)}</td>
      <td class="py-2.5 px-2 text-right ${Math.abs(skewness) > 1 ? 'text-rose-400' : 'text-slate-300'}">${formatNum(skewness)}</td>
      <td class="py-2.5 px-2 text-right">${formatNum(kurtosis)}</td>
    `;
    tbody.appendChild(tr);
  });

  // Chart 1: Normalized Z-Score Distribution (Box-like bar distribution)
  renderEdaNormalizedChart();

  // Chart 2: Coefficient of Variation (CV %)
  renderEdaCVChart(cvLabels, cvValues);
}

function renderEdaNormalizedChart() {
  const ctx = document.getElementById("chartEdaNormalized");
  if (!ctx) return;
  if (appState.charts.edaNorm) appState.charts.edaNorm.destroy();

  const labels = appState.numericColumns;
  const meanData = [];
  const stdData = [];

  labels.forEach(col => {
    const vals = getColValues(col);
    const mean = ss.mean(vals);
    const sd = ss.standardDeviation(vals);
    meanData.push(mean);
    stdData.push(sd);
  });

  appState.charts.edaNorm = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: '평균값 (Mean)',
          data: meanData,
          backgroundColor: 'rgba(99, 102, 241, 0.7)',
          borderColor: 'rgb(99, 102, 241)',
          borderWidth: 1,
          borderRadius: 4
        },
        {
          label: '표준편차 (Std Dev)',
          data: stdData,
          backgroundColor: 'rgba(236, 72, 153, 0.7)',
          borderColor: 'rgb(236, 72, 153)',
          borderWidth: 1,
          borderRadius: 4
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#94a3b8', font: { size: 11 } } },
        tooltip: {
          backgroundColor: '#0f172a',
          borderColor: '#334155',
          borderWidth: 1
        }
      },
      scales: {
        x: { ticks: { color: '#94a3b8', font: { size: 10 } }, grid: { color: '#1e293b' } },
        y: { ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } }
      }
    }
  });
}

function renderEdaCVChart(labels, values) {
  const ctx = document.getElementById("chartEdaCV");
  if (!ctx) return;
  if (appState.charts.edaCV) appState.charts.edaCV.destroy();

  appState.charts.edaCV = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: '변동계수 (CV %)',
        data: values,
        backgroundColor: values.map(v => v > 30 ? 'rgba(244, 63, 94, 0.75)' : 'rgba(6, 182, 212, 0.75)'),
        borderColor: values.map(v => v > 30 ? 'rgb(244, 63, 94)' : 'rgb(6, 182, 212)'),
        borderWidth: 1,
        borderRadius: 4
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => `CV: ${ctx.raw.toFixed(2)}% (${ctx.raw > 30 ? '높은 변동성' : '안정적'})`
          }
        }
      },
      scales: {
        x: { ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } },
        y: { ticks: { color: '#94a3b8', font: { size: 10 } }, grid: { color: '#1e293b' } }
      }
    }
  });
}

/* ==========================================================================
   MODULE 2: 분포 & 정규성 검정 (Normality & Distribution)
   ========================================================================== */
function renderDistributionTab() {
  const varSelect = document.getElementById("distVarSelect");
  const binSelect = document.getElementById("distBinSelect");
  if (!varSelect || !varSelect.value) return;

  const col = varSelect.value;
  const numBins = parseInt(binSelect.value) || 15;
  const vals = getColValues(col);
  if (vals.length === 0) return;

  const n = vals.length;
  const mean = ss.mean(vals);
  const sd = ss.standardDeviation(vals);
  const min = ss.min(vals);
  const max = ss.max(vals);

  const skewness = ss.sampleSkewness ? ss.sampleSkewness(vals) : calculateSkewness(vals, mean, sd);
  const kurtosis = calculateKurtosis(vals, mean, sd); // Excess kurtosis

  // Jarque-Bera Test Statistic: JB = n/6 * (S^2 + (K^2 / 4))
  const jbStat = (n / 6) * (Math.pow(skewness, 2) + Math.pow(kurtosis, 2) / 4);
  // Chi-Square with 2 degrees of freedom: p-value = exp(-JB / 2)
  const pValue = Math.exp(-jbStat / 2);

  // Update Badges
  document.getElementById("distSkewVal").textContent = skewness.toFixed(3);
  document.getElementById("distSkewDesc").textContent = 
    Math.abs(skewness) < 0.5 ? "대칭적 분포 (정규성 적합)" : skewness > 0 ? "우측 꼬리가 긴 분포 (양의 왜도)" : "좌측 꼬리가 긴 분포 (음의 왜도)";

  document.getElementById("distKurtVal").textContent = kurtosis.toFixed(3);
  document.getElementById("distKurtDesc").textContent = 
    Math.abs(kurtosis) < 0.5 ? "정규분포 수준의 첨도 (Mesokurtic)" : kurtosis > 0 ? "중심 집중 및 두꺼운 꼬리 (Leptokurtic)" : "평평한 분포 (Platykurtic)";

  document.getElementById("distJBStat").textContent = jbStat.toFixed(2);
  document.getElementById("distJBPval").textContent = `p-value: ${pValue < 0.001 ? '< 0.001' : pValue.toFixed(4)}`;

  const verdictEl = document.getElementById("distNormalityVerdict");
  const adviceEl = document.getElementById("distNormalityAdvice");
  if (pValue >= 0.05) {
    verdictEl.textContent = "정규성 만족 (H0 채택)";
    verdictEl.className = "text-base font-bold font-mono mt-1 text-emerald-400";
    adviceEl.textContent = "모수 통계 분석(t-test, ANOVA, 선형회귀 등)을 신뢰성 있게 적용할 수 있습니다.";
  } else {
    verdictEl.textContent = "정규성 기각 (비정규 분포)";
    verdictEl.className = "text-base font-bold font-mono mt-1 text-rose-400";
    adviceEl.textContent = "표본이 비정규성을 띱니다. 비모수 분석 또는 데이터 변환을 고려할 수 있습니다.";
  }

  // Create Bins for Histogram
  const binWidth = (max - min) / numBins;
  const binCounts = new Array(numBins).fill(0);
  const binLabels = [];
  const normalCurveData = [];

  for (let i = 0; i < numBins; i++) {
    const bStart = min + i * binWidth;
    const bEnd = bStart + binWidth;
    binLabels.push(`${bStart.toFixed(1)}~${bEnd.toFixed(1)}`);
  }

  vals.forEach(v => {
    let bIdx = Math.floor((v - min) / binWidth);
    if (bIdx >= numBins) bIdx = numBins - 1;
    if (bIdx < 0) bIdx = 0;
    binCounts[bIdx]++;
  });

  // Calculate Normal Density Curve scaled to counts
  for (let i = 0; i < numBins; i++) {
    const mid = min + (i + 0.5) * binWidth;
    const z = (mid - mean) / sd;
    const pdf = (1 / (sd * Math.sqrt(2 * Math.PI))) * Math.exp(-0.5 * z * z);
    const expectedCount = pdf * n * binWidth;
    normalCurveData.push(expectedCount);
  }

  // Render Histogram & Normal Curve
  const histCtx = document.getElementById("chartHistogram");
  if (appState.charts.distHist) appState.charts.distHist.destroy();

  appState.charts.distHist = new Chart(histCtx, {
    data: {
      labels: binLabels,
      datasets: [
        {
          type: 'bar',
          label: '실측 빈도수 (Frequency)',
          data: binCounts,
          backgroundColor: 'rgba(16, 185, 129, 0.65)',
          borderColor: 'rgb(16, 185, 129)',
          borderWidth: 1,
          borderRadius: 4
        },
        {
          type: 'line',
          label: '이론적 정규분포 곡선',
          data: normalCurveData,
          borderColor: 'rgb(244, 63, 94)',
          borderWidth: 2,
          pointRadius: 2,
          tension: 0.4
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#94a3b8' } }
      },
      scales: {
        x: { ticks: { color: '#94a3b8', font: { size: 9 }, maxRotation: 45 }, grid: { color: '#1e293b' } },
        y: { ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } }
      }
    }
  });

  // Render Empirical CDF vs Normal CDF
  const sortedVals = [...vals].sort((a, b) => a - b);
  const cdfPoints = [];
  const normalCdfPoints = [];

  for (let i = 0; i < sortedVals.length; i++) {
    const x = sortedVals[i];
    const empiricalP = (i + 1) / sortedVals.length;
    const z = (x - mean) / sd;
    const theoreticalP = ss.cumulativeStdNormalProbability(z);

    cdfPoints.push({ x: x, y: empiricalP });
    normalCdfPoints.push({ x: x, y: theoreticalP });
  }

  const cdfCtx = document.getElementById("chartCDF");
  if (appState.charts.distCDF) appState.charts.distCDF.destroy();

  appState.charts.distCDF = new Chart(cdfCtx, {
    type: 'scatter',
    data: {
      datasets: [
        {
          label: '경험적 누적분포 (Empirical CDF)',
          data: cdfPoints,
          borderColor: 'rgb(6, 182, 212)',
          backgroundColor: 'rgba(6, 182, 212, 0.5)',
          showLine: true,
          pointRadius: 2,
          tension: 0.1
        },
        {
          label: '이론적 정규 누적분포 (Normal CDF)',
          data: normalCdfPoints,
          borderColor: 'rgb(168, 85, 247)',
          borderDash: [5, 5],
          showLine: true,
          pointRadius: 0,
          tension: 0.3
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#94a3b8' } }
      },
      scales: {
        x: { title: { display: true, text: col, color: '#94a3b8' }, ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } },
        y: { title: { display: true, text: '누적확률 P(X <= x)', color: '#94a3b8' }, min: 0, max: 1, ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } }
      }
    }
  });
}

/* ==========================================================================
   MODULE 3: 상관관계 매트릭스 & 회귀분석 (Correlation & Regression)
   ========================================================================== */
function renderCorrelationTab() {
  const cols = appState.numericColumns;
  if (cols.length === 0) return;

  // Build Correlation Matrix Table / Heatmap Container
  const container = document.getElementById("corrMatrixContainer");
  container.innerHTML = "";

  const table = document.createElement("table");
  table.className = "w-full text-xs font-mono border-collapse";

  // Header
  const thead = document.createElement("thead");
  let trHead = `<tr class="border-b border-slate-800"><th class="p-2 text-left text-slate-400"></th>`;
  cols.forEach(c => {
    trHead += `<th class="p-2 text-center text-slate-400 text-[10px] max-w-[70px] truncate" title="${c}">${truncate(c, 7)}</th>`;
  });
  trHead += `</tr>`;
  thead.innerHTML = trHead;
  table.appendChild(thead);

  // Body
  const tbody = document.createElement("tbody");
  const matrixData = [];

  cols.forEach((col1, i) => {
    const vals1 = getColValues(col1);
    const tr = document.createElement("tr");
    tr.className = "border-b border-slate-800/40";
    
    let rowHtml = `<td class="p-2 font-semibold text-slate-300 text-[11px] truncate max-w-[80px]" title="${col1}">${truncate(col1, 8)}</td>`;

    cols.forEach((col2, j) => {
      const vals2 = getColValues(col2);
      let r = 0;
      if (i === j) {
        r = 1.0;
      } else {
        r = appState.corrMethod === "pearson" ? ss.sampleCorrelation(vals1, vals2) : calculateSpearman(vals1, vals2);
      }
      if (isNaN(r)) r = 0;

      const bgColor = getCorrColor(r);
      const textColor = Math.abs(r) > 0.4 ? 'text-white' : 'text-slate-300';
      
      rowHtml += `
        <td class="p-1 text-center">
          <div class="heatmap-cell p-1.5 rounded transition cursor-pointer" 
               style="background-color: ${bgColor};" 
               data-x="${col1}" data-y="${col2}" 
               title="${col1} vs ${col2}\n${appState.corrMethod.toUpperCase()} r = ${r.toFixed(3)}">
            <span class="${textColor} font-bold">${r.toFixed(2)}</span>
          </div>
        </td>
      `;
    });

    tr.innerHTML = rowHtml;
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  container.appendChild(table);

  // Heatmap click listener to update scatter regression pair
  container.querySelectorAll(".heatmap-cell").forEach(cell => {
    cell.addEventListener("click", () => {
      const colX = cell.getAttribute("data-x");
      const colY = cell.getAttribute("data-y");
      document.getElementById("regVarX").value = colX;
      document.getElementById("regVarY").value = colY;
      updateRegressionPlot();
    });
  });

  updateRegressionPlot();
}

function updateRegressionPlot() {
  const colX = document.getElementById("regVarX").value;
  const colY = document.getElementById("regVarY").value;
  if (!colX || !colY) return;

  const valsX = getColValues(colX);
  const valsY = getColValues(colY);
  const paired = [];
  for (let i = 0; i < Math.min(valsX.length, valsY.length); i++) {
    paired.push([valsX[i], valsY[i]]);
  }

  // Linear Regression
  const reg = ss.linearRegression(paired);
  const regLine = ss.linearRegressionLine(reg);
  const r = ss.sampleCorrelation(valsX, valsY);
  const r2 = ss.rSquared(paired, regLine);

  // Update formula and stats UI
  const sign = reg.b >= 0 ? "+" : "-";
  document.getElementById("regEquation").textContent = `Y = ${reg.m.toFixed(4)}X ${sign} ${Math.abs(reg.b).toFixed(4)}`;
  document.getElementById("regR2").textContent = r2.toFixed(4);
  document.getElementById("regR").textContent = (isNaN(r) ? 0 : r).toFixed(4);

  // Prepare Chart Data
  const scatterPoints = paired.map(p => ({ x: p[0], y: p[1] }));
  const minX = ss.min(valsX);
  const maxX = ss.max(valsX);
  const linePoints = [
    { x: minX, y: regLine(minX) },
    { x: maxX, y: regLine(maxX) }
  ];

  const ctx = document.getElementById("chartRegression");
  if (appState.charts.regression) appState.charts.regression.destroy();

  appState.charts.regression = new Chart(ctx, {
    data: {
      datasets: [
        {
          type: 'scatter',
          label: `${colX} vs ${colY}`,
          data: scatterPoints,
          backgroundColor: 'rgba(99, 102, 241, 0.7)',
          borderColor: 'rgb(99, 102, 241)',
          pointRadius: 3.5,
          pointHoverRadius: 6
        },
        {
          type: 'line',
          label: `회귀선 (R²=${r2.toFixed(3)})`,
          data: linePoints,
          borderColor: 'rgb(244, 63, 94)',
          borderWidth: 2,
          pointRadius: 0
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#94a3b8' } }
      },
      scales: {
        x: { title: { display: true, text: colX, color: '#94a3b8' }, ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } },
        y: { title: { display: true, text: colY, color: '#94a3b8' }, ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } }
      }
    }
  });
}

function getCorrColor(r) {
  if (r >= 0) {
    // Blue/Indigo scale (0 to 1)
    const alpha = Math.min(0.9, Math.max(0.1, Math.abs(r) * 0.9));
    return `rgba(99, 102, 241, ${alpha})`;
  } else {
    // Rose/Red scale (-1 to 0)
    const alpha = Math.min(0.9, Math.max(0.1, Math.abs(r) * 0.9));
    return `rgba(244, 63, 94, ${alpha})`;
  }
}

function calculateSpearman(arr1, arr2) {
  const rank1 = getRanks(arr1);
  const rank2 = getRanks(arr2);
  return ss.sampleCorrelation(rank1, rank2);
}

function getRanks(arr) {
  const sorted = arr.map((v, i) => ({ v, i })).sort((a, b) => a.v - b.v);
  const ranks = new Array(arr.length);
  for (let i = 0; i < sorted.length; i++) {
    ranks[sorted[i].i] = i + 1;
  }
  return ranks;
}

/* ==========================================================================
   MODULE 4: 시계열 & 추세 분석 (Time Series & Trend)
   ========================================================================== */
function renderTimeSeriesTab() {
  const col = document.getElementById("tsVarSelect").value;
  const windowSize = parseInt(document.getElementById("tsWindowInput").value) || 5;
  if (!col) return;

  const vals = getColValues(col);
  const labels = appState.rawRows.map((r, i) => {
    return r[appState.timeColumn] ? String(r[appState.timeColumn]) : `#${i + 1}`;
  });

  // Calculate Simple Moving Average (SMA)
  const sma = [];
  for (let i = 0; i < vals.length; i++) {
    if (i < windowSize - 1) {
      sma.push(null);
    } else {
      const windowSlice = vals.slice(i - windowSize + 1, i + 1);
      sma.push(ss.mean(windowSlice));
    }
  }

  // Calculate Exponential Moving Average (EMA)
  const ema = [];
  const k = 2 / (windowSize + 1);
  let prevEma = vals[0];
  ema.push(prevEma);
  for (let i = 1; i < vals.length; i++) {
    const currentEma = vals[i] * k + prevEma * (1 - k);
    ema.push(currentEma);
    prevEma = currentEma;
  }

  const ctx = document.getElementById("chartTimeSeries");
  if (appState.charts.timeSeries) appState.charts.timeSeries.destroy();

  appState.charts.timeSeries = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [
        {
          label: `${col} (Raw Data)`,
          data: vals,
          borderColor: 'rgba(148, 163, 184, 0.45)',
          backgroundColor: 'transparent',
          borderWidth: 1.5,
          pointRadius: 2,
          pointHoverRadius: 5
        },
        {
          label: `단순이동평균 SMA (${windowSize})`,
          data: sma,
          borderColor: 'rgb(245, 158, 11)',
          backgroundColor: 'transparent',
          borderWidth: 2.5,
          pointRadius: 0,
          tension: 0.3
        },
        {
          label: `지수이동평균 EMA (${windowSize})`,
          data: ema,
          borderColor: 'rgb(6, 182, 212)',
          backgroundColor: 'transparent',
          borderWidth: 2,
          borderDash: [4, 4],
          pointRadius: 0,
          tension: 0.3
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { labels: { color: '#94a3b8' } }
      },
      scales: {
        x: { ticks: { color: '#94a3b8', maxTicksLimit: 12, font: { size: 10 } }, grid: { color: '#1e293b' } },
        y: { title: { display: true, text: col, color: '#94a3b8' }, ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } }
      }
    }
  });
}

/* ==========================================================================
   MODULE 5: 공정관리 & 이상치 (SPC & Anomaly Detection)
   ========================================================================== */
function renderSPCTab() {
  const col = document.getElementById("spcVarSelect").value;
  if (!col) return;

  const vals = getColValues(col);
  const n = vals.length;
  if (n === 0) return;

  const mean = ss.mean(vals);
  const sd = ss.standardDeviation(vals);
  const ucl = mean + 3 * sd;
  const lcl = mean - 3 * sd;

  document.getElementById("spcCL").textContent = mean.toFixed(2);
  document.getElementById("spcUCL").textContent = ucl.toFixed(2);
  document.getElementById("spcLCL").textContent = lcl.toFixed(2);

  // Detect Outliers (Z-score >= 3 or <= -3)
  const outliers = [];
  vals.forEach((v, i) => {
    const z = (v - mean) / sd;
    if (Math.abs(z) >= 3 || v > ucl || v < lcl) {
      const timeVal = appState.rawRows[i][appState.timeColumn] || `#${i + 1}`;
      outliers.push({
        index: i + 1,
        time: timeVal,
        val: v,
        z: z,
        reason: v > ucl ? '상부한계(UCL) 초과' : '하부한계(LCL) 미달'
      });
    }
  });

  document.getElementById("spcOutlierCount").textContent = `${outliers.length}건`;

  // Populate Outlier Table
  const tbody = document.getElementById("spcOutlierTableBody");
  tbody.innerHTML = "";
  if (outliers.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" class="p-4 text-center text-emerald-400 font-sans">🎉 3&sigma; 관리 한계를 벗어난 공정 이상치가 검출되지 않았습니다.</td></tr>`;
  } else {
    outliers.forEach(o => {
      const tr = document.createElement("tr");
      tr.className = "hover:bg-slate-800/40";
      tr.innerHTML = `
        <td class="p-2 font-bold text-slate-300">#${o.index}</td>
        <td class="p-2 text-slate-400">${o.time}</td>
        <td class="p-2 text-right text-rose-400 font-bold">${o.val.toFixed(2)}</td>
        <td class="p-2 text-right font-bold ${o.z > 0 ? 'text-rose-400' : 'text-cyan-400'}">${o.z.toFixed(2)}</td>
        <td class="p-2"><span class="px-2 py-0.5 rounded bg-rose-500/20 text-rose-300 text-[11px]">${o.reason}</span></td>
      `;
      tbody.appendChild(tr);
    });
  }

  // Render SPC Control Chart
  const labels = appState.rawRows.map((r, i) => `#${i + 1}`);
  const pointColors = vals.map(v => (v > ucl || v < lcl) ? 'rgb(244, 63, 94)' : 'rgb(99, 102, 241)');
  const pointRadii = vals.map(v => (v > ucl || v < lcl) ? 6 : 3);

  const ctx = document.getElementById("chartSPC");
  if (appState.charts.spc) appState.charts.spc.destroy();

  appState.charts.spc = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [
        {
          label: '측정값',
          data: vals,
          borderColor: 'rgb(99, 102, 241)',
          pointBackgroundColor: pointColors,
          pointBorderColor: pointColors,
          pointRadius: pointRadii,
          borderWidth: 1.5,
          tension: 0.1
        },
        {
          label: 'UCL (+3σ)',
          data: new Array(n).fill(ucl),
          borderColor: 'rgba(244, 63, 94, 0.8)',
          borderDash: [5, 5],
          pointRadius: 0,
          borderWidth: 2
        },
        {
          label: 'CL (중심선)',
          data: new Array(n).fill(mean),
          borderColor: 'rgba(6, 182, 212, 0.8)',
          pointRadius: 0,
          borderWidth: 1.5
        },
        {
          label: 'LCL (-3σ)',
          data: new Array(n).fill(lcl),
          borderColor: 'rgba(244, 63, 94, 0.8)',
          borderDash: [5, 5],
          pointRadius: 0,
          borderWidth: 2
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#94a3b8' } }
      },
      scales: {
        x: { ticks: { color: '#94a3b8', maxTicksLimit: 15 }, grid: { color: '#1e293b' } },
        y: { title: { display: true, text: col, color: '#94a3b8' }, ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } }
      }
    }
  });
}

/* ==========================================================================
   MODULE 6: 가설검정 (t-Test & One-way ANOVA)
   ========================================================================== */
function renderAnovaTab() {
  const col = document.getElementById("anovaVarSelect").value;
  const splitType = document.getElementById("anovaSplitSelect").value;
  if (!col) return;

  const vals = getColValues(col);
  const n = vals.length;
  if (n < 4) return;

  let groups = [];
  let groupLabels = [];

  if (splitType === "halves") {
    // 2 Groups
    const half = Math.floor(n / 2);
    groups = [vals.slice(0, half), vals.slice(half)];
    groupLabels = ["전반부 구간", "후반부 구간"];
  } else if (splitType === "tertiles") {
    // 3 Groups
    const size = Math.floor(n / 3);
    groups = [vals.slice(0, size), vals.slice(size, size * 2), vals.slice(size * 2)];
    groupLabels = ["초기 구간 (T1)", "중기 구간 (T2)", "후기 구간 (T3)"];
  } else {
    // 4 Quartiles
    const size = Math.floor(n / 4);
    groups = [
      vals.slice(0, size),
      vals.slice(size, size * 2),
      vals.slice(size * 2, size * 3),
      vals.slice(size * 3)
    ];
    groupLabels = ["1분위 (Q1)", "2분위 (Q2)", "3분위 (Q3)", "4분위 (Q4)"];
  }

  // Calculate ANOVA / t-test
  const k = groups.length;
  const groupMeans = groups.map(g => ss.mean(g));
  const groupSds = groups.map(g => ss.standardDeviation(g));
  const grandMean = ss.mean(vals);

  // Sum of Squares Between (SSB) and Within (SSW)
  let ssb = 0;
  let ssw = 0;
  groups.forEach((g, idx) => {
    const ni = g.length;
    ssb += ni * Math.pow(groupMeans[idx] - grandMean, 2);
    g.forEach(v => {
      ssw += Math.pow(v - groupMeans[idx], 2);
    });
  });

  const dfBetween = k - 1;
  const dfWithin = n - k;
  const msb = ssb / dfBetween;
  const msw = ssw / dfWithin;
  const fStat = msb / msw;

  // Approximate p-value for F-distribution
  const pVal = approximateF_PValue(fStat, dfBetween, dfWithin);

  if (k === 2) {
    // For 2 groups, t = sqrt(F)
    const tStat = Math.sqrt(fStat) * (groupMeans[0] < groupMeans[1] ? -1 : 1);
    document.getElementById("anovaStatType").textContent = "독립표본 t-통계량 (t)";
    document.getElementById("anovaStatVal").textContent = tStat.toFixed(3);
  } else {
    document.getElementById("anovaStatType").textContent = "F-통계량 (ANOVA)";
    document.getElementById("anovaStatVal").textContent = fStat.toFixed(3);
  }

  document.getElementById("anovaPVal").textContent = pVal < 0.001 ? "< 0.001" : pVal.toFixed(4);

  const verdictEl = document.getElementById("anovaVerdict");
  const interpEl = document.getElementById("anovaInterpretation");

  if (pVal < 0.05) {
    verdictEl.textContent = "집단 간 통계적으로 유의미한 차이 있음 (p < 0.05)";
    verdictEl.className = "text-sm font-semibold font-mono text-rose-400 mt-1";
    interpEl.textContent = `구간 분할에 따른 [${col}]의 평균 차이가 우연에 의한 것일 확률이 5% 미만으로, 시간대/조건별로 공정 특성에 실질적인 변화가 발생했습니다.`;
  } else {
    verdictEl.textContent = "집단 간 유의미한 차이 없음 (귀무가설 유지)";
    verdictEl.className = "text-sm font-semibold font-mono text-emerald-400 mt-1";
    interpEl.textContent = `구간별 [${col}] 평균의 차이가 통계적 유의수준(α=0.05)을 만족하지 못하므로, 전 시간대에 걸쳐 안정적으로 유지되고 있습니다.`;
  }

  // Render Group Means & Error Margin (95% CI)
  const ci95 = groups.map((g, idx) => 1.96 * (groupSds[idx] / Math.sqrt(g.length)));

  const ctx = document.getElementById("chartAnovaGroups");
  if (appState.charts.anova) appState.charts.anova.destroy();

  appState.charts.anova = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: groupLabels,
      datasets: [
        {
          label: '집단별 평균 (Mean)',
          data: groupMeans,
          backgroundColor: 'rgba(168, 85, 247, 0.65)',
          borderColor: 'rgb(168, 85, 247)',
          borderWidth: 1.5,
          borderRadius: 6
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#94a3b8' } },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const idx = ctx.dataIndex;
              return `평균: ${groupMeans[idx].toFixed(2)} (95% CI: ±${ci95[idx].toFixed(2)})`;
            }
          }
        }
      },
      scales: {
        x: { ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } },
        y: { title: { display: true, text: col, color: '#94a3b8' }, ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } }
      }
    }
  });
}

function approximateF_PValue(F, df1, df2) {
  if (isNaN(F) || F <= 0) return 1.0;
  // Beta incomplete approx
  const x = df2 / (df2 + df1 * F);
  const a = df2 / 2;
  const b = df1 / 2;
  // Simple approximation for UI feedback
  const z = (Math.pow((1 - 2 / (9 * df2)) * Math.pow(F, 1/3) - (1 - 2 / (9 * df1)), 1) / Math.sqrt(2 / (9 * df2) * Math.pow(F, 2/3) + 2 / (9 * df1)));
  const p = 1 - ss.cumulativeStdNormalProbability(z);
  return Math.max(0.0001, Math.min(1.0, p));
}

/* ==========================================================================
   MODULE 7: 군집(K-Means) & 주성분 분석(PCA)
   ========================================================================== */
function renderClusterTab() {
  const k = parseInt(document.getElementById("clusterKSelect").value) || 3;
  const cols = appState.numericColumns;
  if (cols.length < 2) return;

  const rows = appState.rawRows;
  const n = rows.length;

  // 1. Standardize Data Matrix
  const matrix = [];
  const colStats = cols.map(c => {
    const vals = getColValues(c);
    return { mean: ss.mean(vals), sd: ss.standardDeviation(vals) || 1 };
  });

  rows.forEach(r => {
    const rowVec = cols.map((c, idx) => {
      const v = parseFloat(r[c]) || 0;
      return (v - colStats[idx].mean) / colStats[idx].sd;
    });
    matrix.push(rowVec);
  });

  // 2. K-Means Clustering
  const { clusters, centroids } = runKMeans(matrix, k);

  // 3. Simple PCA (First 2 Principal Components via Covariance Eigenvector approx)
  const pcaResults = computePCA2D(matrix);

  document.getElementById("pcaVar1").textContent = `${(pcaResults.varRatio1 * 100).toFixed(1)}%`;
  document.getElementById("pcaVar2").textContent = `${(pcaResults.varRatio2 * 100).toFixed(1)}%`;
  document.getElementById("pcaVarTotal").textContent = `${((pcaResults.varRatio1 + pcaResults.varRatio2) * 100).toFixed(1)}%`;

  // Build Scatter Datasets by Cluster
  const clusterColors = [
    { bg: 'rgba(236, 72, 153, 0.7)', border: 'rgb(236, 72, 153)' },
    { bg: 'rgba(99, 102, 241, 0.7)', border: 'rgb(99, 102, 241)' },
    { bg: 'rgba(16, 185, 129, 0.7)', border: 'rgb(16, 185, 129)' },
    { bg: 'rgba(245, 158, 11, 0.7)', border: 'rgb(245, 158, 11)' },
    { bg: 'rgba(6, 182, 212, 0.7)', border: 'rgb(6, 182, 212)' }
  ];

  const datasets = [];
  for (let cIdx = 0; cIdx < k; cIdx++) {
    const clusterPoints = [];
    pcaResults.projected.forEach((pt, rowIdx) => {
      if (clusters[rowIdx] === cIdx) {
        clusterPoints.push({ x: pt[0], y: pt[1] });
      }
    });

    datasets.push({
      label: `군집 ${cIdx + 1} (${clusterPoints.length}개)`,
      data: clusterPoints,
      backgroundColor: clusterColors[cIdx % clusterColors.length].bg,
      borderColor: clusterColors[cIdx % clusterColors.length].border,
      pointRadius: 4,
      pointHoverRadius: 7
    });
  }

  // Render 2D PCA Scatter Chart
  const ctx = document.getElementById("chartClusterPCA");
  if (appState.charts.pca) appState.charts.pca.destroy();

  appState.charts.pca = new Chart(ctx, {
    type: 'scatter',
    data: { datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#94a3b8' } }
      },
      scales: {
        x: { title: { display: true, text: `주성분 1 (PC1: ${(pcaResults.varRatio1 * 100).toFixed(1)}%)`, color: '#94a3b8' }, ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } },
        y: { title: { display: true, text: `주성분 2 (PC2: ${(pcaResults.varRatio2 * 100).toFixed(1)}%)`, color: '#94a3b8' }, ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } }
      }
    }
  });

  // Cluster Profile Breakdown Card
  const profileContainer = document.getElementById("clusterProfileContainer");
  profileContainer.innerHTML = "";

  for (let cIdx = 0; cIdx < k; cIdx++) {
    const memberIndices = clusters.map((c, i) => c === cIdx ? i : -1).filter(i => i !== -1);
    const count = memberIndices.length;

    let profileHtml = `
      <div class="p-3 bg-slate-950/60 border border-slate-800 rounded-lg space-y-1.5">
        <div class="flex items-center justify-between">
          <span class="font-bold text-pink-400">군집 #${cIdx + 1}</span>
          <span class="bg-slate-800 px-2 py-0.5 rounded text-[10px] text-slate-300 font-mono">${count}개 (${((count/n)*100).toFixed(1)}%)</span>
        </div>
        <div class="text-[11px] text-slate-400 space-y-0.5">
    `;

    cols.slice(0, 4).forEach((c, idx) => {
      const rawAvg = ss.mean(memberIndices.map(i => parseFloat(rows[i][c]) || 0));
      profileHtml += `<div class="flex justify-between font-mono"><span>${truncate(c, 10)}:</span> <span class="text-slate-200 font-semibold">${rawAvg.toFixed(2)}</span></div>`;
    });

    profileHtml += `</div></div>`;
    profileContainer.innerHTML += profileHtml;
  }
}

// K-Means Implementation
function runKMeans(data, k, maxIter = 50) {
  const n = data.length;
  const dim = data[0].length;
  let centroids = [];

  // Random initialization (Forgy method)
  const chosenIndices = new Set();
  while (centroids.length < k && centroids.length < n) {
    const idx = Math.floor(Math.random() * n);
    if (!chosenIndices.has(idx)) {
      chosenIndices.add(idx);
      centroids.push([...data[idx]]);
    }
  }

  let clusters = new Array(n).fill(0);

  for (let iter = 0; iter < maxIter; iter++) {
    let changed = false;

    // Assignment Step
    for (let i = 0; i < n; i++) {
      let minDist = Infinity;
      let bestC = 0;
      for (let c = 0; c < k; c++) {
        const d = euclideanDist(data[i], centroids[c]);
        if (d < minDist) {
          minDist = d;
          bestC = c;
        }
      }
      if (clusters[i] !== bestC) {
        clusters[i] = bestC;
        changed = true;
      }
    }

    if (!changed && iter > 0) break;

    // Update Step
    const counts = new Array(k).fill(0);
    const newCentroids = Array.from({ length: k }, () => new Array(dim).fill(0));

    for (let i = 0; i < n; i++) {
      const c = clusters[i];
      counts[c]++;
      for (let d = 0; d < dim; d++) {
        newCentroids[c][d] += data[i][d];
      }
    }

    for (let c = 0; c < k; c++) {
      if (counts[c] > 0) {
        for (let d = 0; d < dim; d++) {
          centroids[c][d] = newCentroids[c][d] / counts[c];
        }
      }
    }
  }

  return { clusters, centroids };
}

function euclideanDist(v1, v2) {
  let sum = 0;
  for (let i = 0; i < v1.length; i++) {
    sum += Math.pow(v1[i] - v2[i], 2);
  }
  return Math.sqrt(sum);
}

// 2D PCA Computation via Power Iteration
function computePCA2D(matrix) {
  const n = matrix.length;
  const p = matrix[0].length;

  // Covariance Matrix (matrix is already standardized mean=0, std=1)
  const cov = Array.from({ length: p }, () => new Array(p).fill(0));
  for (let i = 0; i < p; i++) {
    for (let j = 0; j < p; j++) {
      let sum = 0;
      for (let r = 0; r < n; r++) {
        sum += matrix[r][i] * matrix[r][j];
      }
      cov[i][j] = sum / (n - 1);
    }
  }

  // 1st Principal Component
  const v1 = powerIteration(cov, 50);
  const lambda1 = rayleighQuotient(cov, v1);

  // Deflate Covariance Matrix for 2nd PC
  const covDeflated = cov.map((row, i) =>
    row.map((val, j) => val - lambda1 * v1[i] * v1[j])
  );
  const v2 = powerIteration(covDeflated, 50);
  const lambda2 = rayleighQuotient(covDeflated, v2);

  const totalVar = cov.reduce((sum, row, i) => sum + row[i], 0) || 1;

  // Project points
  const projected = matrix.map(row => {
    const pc1 = row.reduce((sum, val, idx) => sum + val * v1[idx], 0);
    const pc2 = row.reduce((sum, val, idx) => sum + val * v2[idx], 0);
    return [pc1, pc2];
  });

  return {
    projected,
    varRatio1: Math.min(1.0, Math.max(0.1, lambda1 / totalVar)),
    varRatio2: Math.min(1.0, Math.max(0.05, lambda2 / totalVar))
  };
}

function powerIteration(mat, iter = 40) {
  const dim = mat.length;
  let b = Array.from({ length: dim }, () => Math.random());
  for (let it = 0; it < iter; it++) {
    const bNext = new Array(dim).fill(0);
    for (let i = 0; i < dim; i++) {
      for (let j = 0; j < dim; j++) {
        bNext[i] += mat[i][j] * b[j];
      }
    }
    const norm = Math.sqrt(bNext.reduce((s, v) => s + v * v, 0)) || 1;
    b = bNext.map(v => v / norm);
  }
  return b;
}

function rayleighQuotient(mat, v) {
  let num = 0;
  for (let i = 0; i < mat.length; i++) {
    for (let j = 0; j < mat.length; j++) {
      num += v[i] * mat[i][j] * v[j];
    }
  }
  return Math.max(0, num);
}

/* ==========================================================================
   MODULE 8: 원본 데이터 테이블 뷰어 (Raw Dataset Explorer)
   ========================================================================== */
function getFilteredRows() {
  if (!appState.tableSearch) return appState.rawRows;
  const q = appState.tableSearch;
  return appState.rawRows.filter(r => {
    return Object.values(r).some(v => String(v).toLowerCase().includes(q));
  });
}

function renderDataTable() {
  const thead = document.getElementById("rawTableHead");
  const tbody = document.getElementById("rawTableBody");
  if (!thead || !tbody) return;

  // Table Headers
  thead.innerHTML = `<tr>${appState.columns.map(c => `<th class="p-3 text-slate-300 font-semibold">${c}</th>`).join('')}</tr>`;

  // Pagination
  const filtered = getFilteredRows();
  const totalRows = filtered.length;
  const totalPages = Math.ceil(totalRows / appState.tablePageSize) || 1;
  if (appState.tablePage > totalPages) appState.tablePage = totalPages;

  const startIdx = (appState.tablePage - 1) * appState.tablePageSize;
  const pageRows = filtered.slice(startIdx, startIdx + appState.tablePageSize);

  tbody.innerHTML = "";
  pageRows.forEach(r => {
    const tr = document.createElement("tr");
    tr.className = "hover:bg-slate-800/40 transition";
    tr.innerHTML = appState.columns.map(c => `<td class="p-2.5 text-slate-300">${r[c] !== null && r[c] !== undefined ? r[c] : '-'}</td>`).join('');
    tbody.appendChild(tr);
  });

  document.getElementById("tableRecordCount").textContent = `표시 중: ${filtered.length > 0 ? startIdx + 1 : 0} ~ ${Math.min(startIdx + appState.tablePageSize, totalRows)} / 총 ${totalRows} 레코드`;
  document.getElementById("tablePageNum").textContent = `${appState.tablePage} / ${totalPages}`;
  document.getElementById("btnPrevPage").disabled = appState.tablePage <= 1;
  document.getElementById("btnNextPage").disabled = appState.tablePage >= totalPages;
}

/* ==========================================================================
   EXPORTERS & UTILITIES
   ========================================================================== */
function exportRawDataCsv() {
  if (appState.rawRows.length === 0) return;
  const worksheet = XLSX.utils.json_to_sheet(appState.rawRows);
  const csv = XLSX.utils.sheet_to_csv(worksheet);
  downloadBlob(csv, "dataset_export.csv", "text/csv;charset=utf-8;");
}

function exportStatsCsv() {
  const statsRows = [];
  appState.numericColumns.forEach(col => {
    const vals = getColValues(col);
    if (vals.length === 0) return;
    const mean = ss.mean(vals);
    const sd = ss.standardDeviation(vals);
    statsRows.push({
      "변수명": col,
      "표본수(N)": vals.length,
      "평균(Mean)": mean.toFixed(4),
      "표준편차(SD)": sd.toFixed(4),
      "중앙값(Median)": ss.median(vals).toFixed(4),
      "최솟값(Min)": ss.min(vals).toFixed(4),
      "Q1": ss.quantile(vals, 0.25).toFixed(4),
      "Q3": ss.quantile(vals, 0.75).toFixed(4),
      "최댓값(Max)": ss.max(vals).toFixed(4),
      "왜도(Skewness)": calculateSkewness(vals, mean, sd).toFixed(4),
      "첨도(Kurtosis)": calculateKurtosis(vals, mean, sd).toFixed(4)
    });
  });

  const ws = XLSX.utils.json_to_sheet(statsRows);
  const csv = XLSX.utils.sheet_to_csv(ws);
  downloadBlob(csv, "statistical_summary_report.csv", "text/csv;charset=utf-8;");
}

function downloadBlob(content, filename, mimeType) {
  const blob = new Blob(["\uFEFF" + content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function calculateSkewness(vals, mean, sd) {
  const n = vals.length;
  if (n < 3 || sd === 0) return 0;
  const m3 = vals.reduce((acc, v) => acc + Math.pow((v - mean) / sd, 3), 0) / n;
  return (Math.sqrt(n * (n - 1)) / (n - 2)) * m3;
}

function calculateKurtosis(vals, mean, sd) {
  const n = vals.length;
  if (n < 4 || sd === 0) return 0;
  const m4 = vals.reduce((acc, v) => acc + Math.pow((v - mean) / sd, 4), 0) / n;
  return m4 - 3; // Excess kurtosis
}

function formatNum(num) {
  if (num === null || num === undefined || isNaN(num)) return "-";
  return Number(num).toFixed(2);
}

function truncate(str, maxLen) {
  if (!str) return "";
  return str.length > maxLen ? str.substring(0, maxLen) + ".." : str;
}
