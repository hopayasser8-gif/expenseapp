const form = document.getElementById("submit-form");
const statusEl = document.getElementById("status");
const submitBtn = document.getElementById("submit-btn");
const expenseSelect = document.getElementById("expense");
const subexpenseSelect = document.getElementById("subexpense");
const balanceValueEl = document.getElementById("balance-value");
const monthExpenseValueEl = document.getElementById("month-expense-value");
const lastRowsTable = document.getElementById("last-rows-table");
const lastRowsHead = lastRowsTable.querySelector("thead");
const lastRowsBody = lastRowsTable.querySelector("tbody");
const DASHBOARD_REFRESH_MS = 15000;

const EXPENSES = {
  "الأولاد": [
    "Allowance / Pocket money (مصروف)",
    "Transportation (مواصلات)",
    "Other (أخرى)",
    "Haircut (حلاقة)"
  ],
  التعليم: [
    "Schools (مدارس)",
    "Study supplies (مستلزمات دراسة)",
    "Universities (جامعات)",
    "Certificate authentication (تصديق شهادات)",
    "Courses (كورسات)",
    "Exam fees (رسوم امتحانات)",
    "Nanny (ماما)"
  ],
  العلاج: [
    "Medical insurance (تأمين طبي)",
    "Doctor visit / checkup (كشف)",
    "Medication (أدوية)",
    "X-rays & lab tests (أشعة وتحاليل)",
    "Dental (أسنان)"
  ],
  المنزل: [
    "Electricity (كهرباء)",
    "Water (مياه)",
    "Gas (غاز)",
    "Phone & Internet (تلفون وانترنت)",
    "Mobile bills (فواتير موبايل)",
    "Maintenance (صيانة)",
    "IPTV",
    "Household tools (أدوات منزل)",
    "Tips / gratuities (اكراميات)"
  ],
  سيارة: [
    "Fuel (بنزين)",
    "Maintenance (صيانة)",
    "Licensing (تراخيص)",
    "Car wash & oil change (غسيل وزيت)",
    "Tires (كاوتش)"
  ],
  "أكل ومشروبات": [
    "Meat (لحوم)",
    "Water (ماء)",
    "Groceries (بقالة)",
    "Fruits & vegetables (فواكه وخضروات)",
    "Bread (خبز)"
  ],
  تسوق: [
    "Clothes (ملابس)",
    "Shoes (احذية)",
    "Books (كتب)",
    "Electronics / devices (أجهزة)",
    "Yasser (ياسر)",
    "Maha (مها)"
  ],
  ترفيه: [
    "Eating out / restaurants (أكل خارجي)",
    "Cinema & theater (سينما ومسرح)",
    "Summer vacation (مصيف)",
    "Cosmetics / beauty (مكياجات)",
    "Club membership (اشتراك النادي)",
    "Bonus / reward (مكافأة)"
  ],
  زكاة: ["Zakat on money (زكاة المال)", "Zakat al-Fitr (زكاة الفطر)"]
};

function setStatus(text, isError = false) {
  statusEl.textContent = text;
  statusEl.style.color = isError ? "#b42318" : "#166534";
}

function setSelectOptions(selectEl, values) {
  selectEl.innerHTML = "";
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    selectEl.appendChild(option);
  });
}

function initExpenseOptions() {
  const expenseNames = Object.keys(EXPENSES);
  setSelectOptions(expenseSelect, expenseNames);
  setSelectOptions(subexpenseSelect, EXPENSES[expenseNames[0]]);

  expenseSelect.addEventListener("change", () => {
    const selectedExpense = expenseSelect.value;
    setSelectOptions(subexpenseSelect, EXPENSES[selectedExpense] || []);
  });
}

function formatNumber(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) {
    return String(value ?? "");
  }
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 }).format(n);
}

function renderLastRows(rows) {
  lastRowsHead.innerHTML = "";
  lastRowsBody.innerHTML = "";

  if (!Array.isArray(rows) || rows.length === 0) {
    lastRowsBody.innerHTML = '<tr><td>No rows found.</td></tr>';
    return;
  }

  const columns = Object.keys(rows[0]);

  const headRow = document.createElement("tr");
  columns.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    headRow.appendChild(th);
  });
  lastRowsHead.appendChild(headRow);

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    columns.forEach((col) => {
      const td = document.createElement("td");
      td.textContent = String(row[col] ?? "");
      tr.appendChild(td);
    });
    lastRowsBody.appendChild(tr);
  });
}

async function loadDashboard() {
  try {
    const response = await fetch("/api/dashboard");
    const result = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(result.error || "Failed to fetch dashboard.");
    }

    balanceValueEl.textContent = formatNumber(result.balance);
    monthExpenseValueEl.textContent = formatNumber(result.monthExpense);
    renderLastRows(result.rows || []);
  } catch (error) {
    setStatus(`Dashboard error: ${error.message}`, true);
  }
}

initExpenseOptions();
loadDashboard();
setInterval(() => {
  loadDashboard();
}, DASHBOARD_REFRESH_MS);

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  submitBtn.disabled = true;
  setStatus("Submitting...");

  const fd = new FormData(form);
  const payload = {
    date: String(fd.get("date") || "").trim(),
    expense: String(fd.get("expense") || "").trim(),
    subexpense: String(fd.get("subexpense") || "").trim(),
    amount: String(fd.get("amount") || "").trim(),
    note: String(fd.get("note") || "").trim()
  };

  try {
    const response = await fetch("/api/submit", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const result = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(result.error || "Request failed");
    }

    form.reset();
    initExpenseOptions();
    setStatus("Success: row added to Excel.");
    await loadDashboard();
  } catch (error) {
    setStatus(`Error: ${error.message}`, true);
  } finally {
    submitBtn.disabled = false;
  }
});
