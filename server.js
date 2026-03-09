import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import path from "node:path";
import { fileURLToPath } from "node:url";

dotenv.config();

const app = express();
const port = Number(process.env.PORT || 4000);

const MS_TENANT_ID = String(process.env.MS_TENANT_ID || "").trim();
const MS_CLIENT_ID = String(process.env.MS_CLIENT_ID || "").trim();
const MS_CLIENT_SECRET = String(process.env.MS_CLIENT_SECRET || "").trim();
const MS_DRIVE_ID = String(process.env.MS_DRIVE_ID || "").trim();
const MS_FILE_ITEM_ID = String(process.env.MS_FILE_ITEM_ID || "").trim();
const MS_TABLE_NAME = String(process.env.MS_TABLE_NAME || "").trim();
const SUMMARY_SHEET_NAME = String(process.env.SUMMARY_SHEET_NAME || "Exp26").trim();
const SUMMARY_BALANCE_CELL = String(process.env.SUMMARY_BALANCE_CELL || "M2").trim();
const SUMMARY_MONTH_EXPENSE_CELL = String(process.env.SUMMARY_MONTH_EXPENSE_CELL || "N2").trim();
const AUTO_UPDATE_MONTH_EXPENSE = String(process.env.AUTO_UPDATE_MONTH_EXPENSE || "true").toLowerCase() !== "false";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

function configIsComplete() {
  return Boolean(
    MS_TENANT_ID &&
      MS_CLIENT_ID &&
      MS_CLIENT_SECRET &&
      MS_DRIVE_ID &&
      MS_FILE_ITEM_ID &&
      MS_TABLE_NAME
  );
}

async function getGraphAccessToken() {
  if (!configIsComplete()) {
    throw new Error("Missing Microsoft Graph env values in .env");
  }

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(MS_TENANT_ID)}/oauth2/v2.0/token`;

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "client_credentials",
      client_id: MS_CLIENT_ID,
      client_secret: MS_CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default"
    })
  });

  if (!response.ok) {
    throw new Error(`Token request failed: ${response.status} ${await response.text()}`);
  }

  const data = await response.json();
  if (!data?.access_token) {
    throw new Error("No access token returned by Microsoft Graph auth.");
  }

  return data.access_token;
}

async function addRowToExcel(values) {
  const accessToken = await getGraphAccessToken();
  const base = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(MS_DRIVE_ID)}/items/${encodeURIComponent(MS_FILE_ITEM_ID)}/workbook/tables/${encodeURIComponent(MS_TABLE_NAME)}`;
  const columnsEndpoint = `${base}/columns?$top=200`;
  const endpoint = `${base}/rows/add`;

  const columnsResponse = await fetch(columnsEndpoint, {
    method: "GET",
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  if (!columnsResponse.ok) {
    throw new Error(`Failed to read table columns: ${columnsResponse.status} ${await columnsResponse.text()}`);
  }

  const columnsPayload = await columnsResponse.json();
  const columnCount = Array.isArray(columnsPayload?.value) ? columnsPayload.value.length : 0;
  if (columnCount <= 0) {
    throw new Error("Excel table has no columns.");
  }

  const normalizedValues = [...values];
  if (normalizedValues.length < columnCount) {
    normalizedValues.push(...new Array(columnCount - normalizedValues.length).fill(""));
  } else if (normalizedValues.length > columnCount) {
    normalizedValues.length = columnCount;
  }

  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ values: [normalizedValues] })
  });

  if (!response.ok) {
    throw new Error(`Excel insert failed: ${response.status} ${await response.text()}`);
  }

  return response.json();
}

async function graphJsonRequest(accessToken, method, endpoint, body) {
  const response = await fetch(endpoint, {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: body ? JSON.stringify(body) : undefined
  });

  if (!response.ok) {
    throw new Error(`Graph request failed: ${response.status} ${await response.text()}`);
  }

  return response.json();
}

async function getDashboardData() {
  const accessToken = await getGraphAccessToken();
  const workbookBase = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(MS_DRIVE_ID)}/items/${encodeURIComponent(MS_FILE_ITEM_ID)}/workbook`;
  const tableBase = `${workbookBase}/tables/${encodeURIComponent(MS_TABLE_NAME)}`;

  const [balanceRange, monthExpenseRange, columnsPayload, rowsPayload] = await Promise.all([
    graphJsonRequest(
      accessToken,
      "GET",
      `${workbookBase}/worksheets/${encodeURIComponent(SUMMARY_SHEET_NAME)}/range(address='${encodeURIComponent(SUMMARY_BALANCE_CELL)}')`
    ),
    graphJsonRequest(
      accessToken,
      "GET",
      `${workbookBase}/worksheets/${encodeURIComponent(SUMMARY_SHEET_NAME)}/range(address='${encodeURIComponent(SUMMARY_MONTH_EXPENSE_CELL)}')`
    ),
    graphJsonRequest(accessToken, "GET", `${tableBase}/columns?$top=200`),
    graphJsonRequest(accessToken, "GET", `${tableBase}/rows?$top=5000`)
  ]);

  const balance = balanceRange?.values?.[0]?.[0] ?? "";
  const monthExpense = monthExpenseRange?.values?.[0]?.[0] ?? "";

  const columnNames = Array.isArray(columnsPayload?.value)
    ? columnsPayload.value.map((column) => String(column.name || ""))
    : [];
  const rowValues = Array.isArray(rowsPayload?.value) ? rowsPayload.value : [];
  const last5 = rowValues.slice(-5).reverse();

  const rows = last5.map((row) => {
    const values = Array.isArray(row?.values?.[0]) ? row.values[0] : [];
    const item = {};
    for (let i = 0; i < columnNames.length; i += 1) {
      item[columnNames[i] || `col_${i + 1}`] = values[i] ?? "";
    }
    return item;
  });

  return {
    balance,
    monthExpense,
    rows
  };
}

function toNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  const parsed = Number(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : 0;
}

async function incrementMonthExpense(amountToAdd) {
  const accessToken = await getGraphAccessToken();
  const workbookBase = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(MS_DRIVE_ID)}/items/${encodeURIComponent(MS_FILE_ITEM_ID)}/workbook`;
  const rangeUrl = `${workbookBase}/worksheets/${encodeURIComponent(SUMMARY_SHEET_NAME)}/range(address='${encodeURIComponent(SUMMARY_MONTH_EXPENSE_CELL)}')`;

  const currentRange = await graphJsonRequest(accessToken, "GET", rangeUrl);
  const currentValue = currentRange?.values?.[0]?.[0];
  const nextValue = toNumber(currentValue) + toNumber(amountToAdd);

  await graphJsonRequest(accessToken, "PATCH", rangeUrl, { values: [[nextValue]] });
  return nextValue;
}

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, excelConfigured: configIsComplete() });
});

app.post("/api/submit", async (req, res) => {
  try {
    const date = String(req.body?.date || "").trim();
    const expense = String(req.body?.expense || "").trim();
    const subexpense = String(req.body?.subexpense || "").trim();
    const amountRaw = String(req.body?.amount || "").trim();
    const note = String(req.body?.note || "").trim();
    const amount = Number(amountRaw);

    if (!date || !expense || !subexpense || !amountRaw) {
      return res.status(400).json({ error: "Date, expense, subexpense, and amount are required." });
    }

    if (!Number.isFinite(amount) || amount < 0) {
      return res.status(400).json({ error: "Amount must be a valid non-negative number." });
    }

    const rowValues = [date, expense, subexpense, amount, note];
    const graphResult = await addRowToExcel(rowValues);
    let updatedMonthExpense = null;

    if (AUTO_UPDATE_MONTH_EXPENSE) {
      updatedMonthExpense = await incrementMonthExpense(amount);
    }

    return res.status(201).json({
      ok: true,
      message: "Row added.",
      graphResult,
      updatedMonthExpense
    });
  } catch (error) {
    console.error(error.message);
    return res.status(500).json({
      ok: false,
      error: `Failed to add row. ${error.message}`
    });
  }
});

app.get("/api/dashboard", async (_req, res) => {
  try {
    const data = await getDashboardData();
    return res.json({ ok: true, ...data });
  } catch (error) {
    console.error(error.message);
    return res.status(500).json({
      ok: false,
      error: `Failed to fetch dashboard data. ${error.message}`
    });
  }
});

app.listen(port, () => {
  console.log(`App running at http://localhost:${port}`);
});
