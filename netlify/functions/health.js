import { configIsComplete, jsonResponse } from "./_graph.js";

export async function handler() {
  return jsonResponse(200, {
    ok: true,
    excelConfigured: configIsComplete()
  });
}
