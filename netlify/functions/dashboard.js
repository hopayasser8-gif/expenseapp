import { getDashboardData, jsonResponse } from "./_graph.js";

export async function handler() {
  try {
    const data = await getDashboardData();
    return jsonResponse(200, { ok: true, ...data });
  } catch (error) {
    return jsonResponse(500, {
      ok: false,
      error: `Failed to fetch dashboard data. ${error.message}`
    });
  }
}
