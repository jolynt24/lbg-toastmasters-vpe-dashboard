// Vercel serverless function — /api/data
// GET  → returns the stored dashboard JSON
// POST → overwrites it (requires x-write-token header to match WRITE_TOKEN env var)

const { put, get } = require("@vercel/blob");

const BLOB_PATH = "vpe-dashboard.json";

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-write-token");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  // ── GET: fetch the blob directly by pathname and return its JSON ─────────
  if (req.method === "GET") {
    try {
      // useCache: false — always read the latest write, never a stale CDN copy.
      const result = await get(BLOB_PATH, { access: "private", useCache: false });
      if (!result) return res.status(200).json(null);

      const text = await new Response(result.stream).text();
      return res.status(200).json(JSON.parse(text));
    } catch (err) {
      console.error("Blob read error:", err.message);
      return res.status(503).json({ error: "Storage unavailable" });
    }
  }

  // ── POST: validate token then overwrite the blob ──────────────────────────
  if (req.method === "POST") {
    const serverToken = process.env.WRITE_TOKEN;
    if (serverToken && req.headers["x-write-token"] !== serverToken) {
      return res.status(401).json({ error: "Unauthorized" });
    }

    try {
      await put(BLOB_PATH, JSON.stringify(req.body), {
        access: "private",
        allowOverwrite: true,
        addRandomSuffix: false,
        contentType: "application/json",
      });
      return res.status(200).json({ ok: true });
    } catch (err) {
      console.error("Blob write error:", err.message);
      return res.status(503).json({ error: "Storage unavailable" });
    }
  }

  return res.status(405).json({ error: "Method not allowed" });
};
