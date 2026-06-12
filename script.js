// ===== НАЛАШТУВАННЯ ЕЛЕМЕНТІВ =====
let managersAgg = [];

const fileInput = document.getElementById("file-input");
const managerSelect = document.getElementById("filter-manager");
const categorySelect = document.getElementById("filter-category");
const searchInput = document.getElementById("search-input");
const tableBody = document.getElementById("managers-body");

const kpiPhones = document.getElementById("kpi-phones");
const kpiGlassQty = document.getElementById("kpi-glass-qty");
const kpiGlassPen = document.getElementById("kpi-glass-pen");
const kpiCasesQty = document.getElementById("kpi-cases-qty");
const kpiCasesPen = document.getElementById("kpi-cases-pen");
const kpiCameraQty = document.getElementById("kpi-camera-qty");
const kpiCameraPen = document.getElementById("kpi-camera-pen");

let topGlassChart;

// ===== 1. ЗАВАНТАЖЕННЯ XLS/XLSX З КОМПА =====
fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  console.log("Обрано файл:", file.name);

  const reader = new FileReader();

  reader.onload = (evt) => {
    console.log("Файл прочитано в пам'ять");
    const data = evt.target.result;

    let workbook;
    try {
      workbook = XLSX.read(data, { type: "array" });
    } catch (err) {
      console.error("Помилка XLSX.read:", err);
      alert("Не вдалося прочитати файл Excel. Перевір формат.");
      return;
    }

    const firstSheetName = workbook.SheetNames[0];
    console.log("Перший лист:", firstSheetName);

    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    console.log("Рядків у листі:", rows.length);
    console.log("Перший рядок:", rows[0]);

    managersAgg = aggregateByManager(rows);
    console.log("Після агрегації, менеджерів:", managersAgg.length);

    fillManagerFilter(managersAgg);
    updateAll();
  };

  reader.onerror = (err) => {
    console.error("FileReader error:", err);
    alert("Помилка читання файлу");
  };

  reader.readAsArrayBuffer(file);
});

// ===== 2. АГРЕГАЦІЯ ПО МЕНЕДЖЕРАХ =====
function aggregateByManager(rows) {
  const map = new Map();

  rows.forEach((row, idx) => {
    const manager = (row["Продавець сайту"] || "").trim();
    if (!manager || manager === "Разом") return;

    const phones = toNum(row["Телефони"]);
    const glass = toNum(row["Скло"]);
    const glassProfit = moneyToNum(row["Скло валовий прибуток"]);
    const cases = toNum(row["К-сть чохли"]);
    const camera = toNum(row["К-сть скло на камеру"]);

    if (!map.has(manager)) {
      map.set(manager, {
        manager,
        phones: 0,
        glass: 0,
        glassProfit: 0,
        cases: 0,
        camera: 0
      });
    }

    const agg = map.get(manager);
    agg.phones += phones;
    agg.glass += glass;
    agg.glassProfit += glassProfit;
    agg.cases += cases;
    agg.camera += camera;
  });

  return Array.from(map.values()).map((r) => {
    const p = r.phones || 1;
    return {
      ...r,
      glassPen: (r.glass / p) * 100,
      casesPen: (r.cases / p) * 100,
      cameraPen: (r.camera / p) * 100
    };
  });
}

function toNum(v) {
  if (v === null || v === undefined) return 0;
  const n = Number(String(v).replace(",", ".").replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? 0 : n;
}

function moneyToNum(v) {
  if (!v) return 0;
  const s = String(v).replace("$", "").replace(",", ".").replace(/[^0-9.\-]/g, "");
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

// ===== 3. ФІЛЬТР МЕНЕДЖЕРІВ =====
function fillManagerFilter(data) {
  const managers = [...new Set(data.map((r) => r.manager))].sort();
  managerSelect.innerHTML = '<option value="all">Всі менеджери</option>';
  managers.forEach((m) => {
    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = m;
    managerSelect.appendChild(opt);
  });
}

managerSelect.addEventListener("change", updateAll);
categorySelect.addEventListener("change", updateAll);
if (searchInput) {
  searchInput.addEventListener("input", updateAll);
}

// ===== 4. ОНОВЛЕННЯ ТАБЛИЦІ + KPI + ЧАРТУ =====
function updateAll() {
  if (!managersAgg.length) {
    tableBody.innerHTML = "<tr><td colspan='9'>Немає даних. Завантаж файл Excel.</td></tr>";
    resetKpi();
    return;
  }

  const selectedManager = managerSelect.value;
  const search = (searchInput?.value || "").toLowerCase();

  let rows = managersAgg.filter((r) => {
    const byManager = selectedManager === "all" || r.manager === selectedManager;
    const bySearch = !search || r.manager.toLowerCase().includes(search);
    return byManager && bySearch;
  });

  renderTable(rows);
  updateKpi(rows);
  updateChart(rows);
}

function renderTable(rows) {
  tableBody.innerHTML = "";
  rows.forEach((r) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.manager}</td>
      <td>${r.phones}</td>
      <td>${r.glass}</td>
      <td>$${r.glassProfit.toFixed(0)}</td>
      <td>${r.glassPen.toFixed(1)}%</td>
      <td>${r.cases}</td>
      <td>${r.casesPen.toFixed(1)}%</td>
      <td>${r.camera}</td>
      <td>${r.cameraPen.toFixed(1)}%</td>
    `;
    tableBody.appendChild(tr);
  });
}

function resetKpi() {
  kpiPhones.textContent = "0";
  kpiGlassQty.textContent = "0";
  kpiGlassPen.textContent = "0% до телефонів";
  kpiCasesQty.textContent = "0";
  kpiCasesPen.textContent = "0% до телефонів";
  kpiCameraQty.textContent = "0";
  kpiCameraPen.textContent = "0% до телефонів";
}

function updateKpi(rows) {
  const totals = rows.reduce(
    (acc, r) => {
      acc.phones += r.phones;
      acc.glass += r.glass;
      acc.cases += r.cases;
      acc.camera += r.camera;
      return acc;
    },
    { phones: 0, glass: 0, cases: 0, camera: 0 }
  );

  const p = totals.phones || 1;

  kpiPhones.textContent = totals.phones;
  kpiGlassQty.textContent = totals.glass;
  kpiGlassPen.textContent = `${((totals.glass / p) * 100).toFixed(1)}% до телефонів`;
  kpiCasesQty.textContent = totals.cases;
  kpiCasesPen.textContent = `${((totals.cases / p) * 100).toFixed(1)}% до телефонів`;
  kpiCameraQty.textContent = totals.camera;
  kpiCameraPen.textContent = `${((totals.camera / p) * 100).toFixed(1)}% до телефонів`;
}

// простий стовпчиковий графік
function updateChart(rows) {
  const sorted = [...rows].sort((a, b) => b.glassProfit - a.glassProfit).slice(0, 10);

  const labels = sorted.map((r) => r.manager);
  const values = sorted.map((r) => r.glassProfit);

  const ctx = document.getElementById("topGlassChart");
  if (!ctx) return;

  if (topGlassChart) topGlassChart.destroy();

  topGlassChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Валовий прибуток зі скла, $",
          data: values,
          backgroundColor: "#8b5cf6"
        }
      ]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false }
      },
      scales: {
        y: { beginAtZero: true }
      }
    }
  });
}

// перший виклик
updateAll();
