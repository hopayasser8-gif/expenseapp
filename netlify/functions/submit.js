import {
  AUTO_UPDATE_MONTH_EXPENSE,
  addRowToExcel,
  incrementMonthExpense,
  jsonResponse,
  parseSubmitInput
} from "./_graph.js";

export async function handler(event) {
  if (event.httpMethod !== "POST") {
    return jsonResponse(405, { ok: false, error: "Method not allowed." });
  }

  try {
    const body = event.body ? JSON.parse(event.body) : {};
    const { date, expense, subexpense, amount, note } = parseSubmitInput(body);

    const rowValues = [date, expense, subexpense, amount, note];
    const graphResult = await addRowToExcel(rowValues);
    let updatedMonthExpense = null;

    if (AUTO_UPDATE_MONTH_EXPENSE) {
      updatedMonthExpense = await incrementMonthExpense(amount);
    }

    return jsonResponse(201, {
      ok: true,
      message: "Row added.",
      graphResult,
      updatedMonthExpense
    });
  } catch (error) {
    return jsonResponse(500, {
      ok: false,
      error: `Failed to add row. ${error.message}`
    });
  }
}
