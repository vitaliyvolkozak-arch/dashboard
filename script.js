// Усі агреговані дані по менеджерах
let managersAgg = [];

const fileInput = document.getElementById("file-input");
const managerSelect = document.getElementById("filter-manager");
const categorySelect = document.getElementById("filter-category");
const searchInput = document.getElementById("search-input");
const tableBody = document.getElementById("managers-body");

// KPI елементи
const kpiPhones = document.getElementById("kpi-phones");
const kpiGlassQty = document.getElementById("kpi-glass-qty");
const kpiGlassPen = document.getElementById("kpi-glass-pen");
const kpiCasesQty = document.getElementById("kpi-cases-qty");
const kpiCasesPen = document.getElementById("kpi-cases-pen");
const kpiCameraQty = document.getElementById("kpi-camera-qty");
const kpiCameraPen = document.getElementById("kpi-camera-pen");

// Діаграма
let topGlassChart;

// ===== 1. ЗАВАНТАЖЕННЯ XLS/XLSX =====
fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();

  reader.onload = (evt) => {
    const data = evt.target.result;
    // читаємо файл як масив байтів
    const workbook = XLSX.read(data, { type: "array" });

    // беремо ПЕРШИЙ лист
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];

    // конвертуємо в масив об'єктів (рядок = об'єкт, ключі = назви колонок)
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    managersAgg = aggregateByManager(rows);
    fillManagerFilter(managersAgg);
    updateAll();
  };

  reader.onerror = () => {
    alert("Помилка читання файлу Excel");
  };

  reader.readAsArrayBuffer(file);
});
