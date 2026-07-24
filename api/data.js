// Vercel serverless function — /api/data
// GET  → returns the stored dashboard JSON
// POST → overwrites it (requires x-write-token header to match WRITE_TOKEN env var)

const { put, list } = require("@vercel/blob");

const BLOB_PATH = "vpe-dashboard.json";

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-write-token");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  // ── GET: find the blob and return its JSON ───────────────────────────────
  if (req.method === "GET") {
    try {
      const { blobs } = await list({ prefix: BLOB_PATH });
      if (blobs.length === 0) return res.status(200).json(null);

      // downloadUrl is a pre-signed URL valid for this request
      const blob = blobs[0];
      const fetchUrl = blob.downloadUrl || blob.url;
      const dataRes = await fetch(fetchUrl, {
        headers: blob.downloadUrl
          ? {} // pre-signed, no auth needed
          : { Authorization: `Bearer ${process.env.BLOB_READ_WRITE_TOKEN}` },
      });

      if (!dataRes.ok) return res.status(200).json(null);
      const data = await dataRes.json();
      return res.status(200).json(data);
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
