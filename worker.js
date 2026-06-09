// ============================================================
// ART1S Multimedia Worker — Cloudflare Worker + R2
// Handles: upload, download, delete for multimedia file sharing
// ============================================================

const MAX_FILE_SIZE = 100 * 1024 * 1024; // 100 MB
const RETENTION_DAYS = 30;

// GAS API URL — same as frontend
const GAS_API_URL = "https://script.google.com/macros/s/AKfycbwI5T1i2ylq6h1vXM7P94sawNyBuc7m7JSurUZAxz4PVl08Y2Rk9V0Rc0nKp6tA2i4ATg/exec";

// Allowed MIME types
const ALLOWED_TYPES = new Set([
  // Images
  "image/jpeg", "image/png", "image/gif", "image/webp", "image/svg+xml", "image/bmp", "image/tiff",
  // Videos
  "video/mp4", "video/quicktime", "video/x-msvideo", "video/x-matroska", "video/webm", "video/mpeg",
  // Audio
  "audio/mpeg", "audio/wav", "audio/x-wav", "audio/mp4", "audio/aac", "audio/ogg", "audio/flac",
  // Documents
  "application/pdf",
  "application/msword",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.ms-excel",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-powerpoint",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  "text/plain",
  // Archives (common for multimedia bundles)
  "application/zip",
  "application/x-zip-compressed",
]);

// Allowed file extensions as fallback (some browsers send wrong MIME)
const ALLOWED_EXTENSIONS = new Set([
  ".jpg", ".jpeg", ".png", ".gif", ".webp", ".svg", ".bmp", ".tiff",
  ".mp4", ".mov", ".avi", ".mkv", ".webm", ".mpg", ".mpeg",
  ".mp3", ".wav", ".m4a", ".aac", ".ogg", ".flac",
  ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt",
  ".zip",
]);

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
  });
}

function isAllowedFile(filename, mimeType) {
  const ext = "." + filename.split(".").pop().toLowerCase();
  return ALLOWED_TYPES.has(mimeType) || ALLOWED_EXTENSIONS.has(ext);
}

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

// ============================================================
// UPLOAD HANDLER
// ============================================================
async function handleUpload(request, env) {
  let formData;
  try {
    formData = await request.formData();
  } catch (e) {
    return jsonResponse({ success: false, message: "Format data tidak valid." }, 400);
  }

  const file = formData.get("file");
  const keperluan = formData.get("keperluan") || "";

  if (!file || typeof file === "string") {
    return jsonResponse({ success: false, message: "Tidak ada file yang dipilih." }, 400);
  }

  // Validate file size
  if (file.size > MAX_FILE_SIZE) {
    return jsonResponse({
      success: false,
      message: `Ukuran file terlalu besar. Maksimum ${formatBytes(MAX_FILE_SIZE)}.`
    }, 413);
  }

  // Validate file type
  if (!isAllowedFile(file.name, file.type)) {
    return jsonResponse({
      success: false,
      message: "Tipe file tidak diizinkan. Harap upload gambar, video, dokumen, atau audio."
    }, 415);
  }

  // Generate unique key: timestamp + sanitized filename
  const timestamp = Date.now();
  const safeName = file.name.replace(/[^a-zA-Z0-9._-]/g, "_");
  const r2Key = `${timestamp}_${safeName}`;

  // Upload to R2
  try {
    const fileBuffer = await file.arrayBuffer();
    const expiresAt = new Date(timestamp + RETENTION_DAYS * 24 * 60 * 60 * 1000).toISOString();

    await env.MULTIMEDIA_BUCKET.put(r2Key, fileBuffer, {
      httpMetadata: {
        contentType: file.type || "application/octet-stream",
        contentDisposition: `attachment; filename="${file.name}"`,
      },
      customMetadata: {
        originalName: file.name,
        keperluan: keperluan,
        uploadedAt: new Date(timestamp).toISOString(),
        expiresAt: expiresAt,
      },
    });
  } catch (e) {
    return jsonResponse({ success: false, message: "Gagal menyimpan file. Coba lagi." }, 500);
  }

  // Build the download URL (served by this Worker)
  const workerUrl = new URL(request.url);
  const downloadUrl = `${workerUrl.origin}/api/download/${r2Key}`;

  // Save metadata to Google Apps Script
  try {
    const metaPayload = {
      action: "saveMultimediaFile",
      data: {
        r2Key: r2Key,
        originalName: file.name,
        keperluan: keperluan,
        ukuran: file.size,
        ukuranFormatted: formatBytes(file.size),
        mimeType: file.type,
        downloadUrl: downloadUrl,
        uploadedAt: new Date(timestamp).toISOString(),
        expiresAt: new Date(timestamp + RETENTION_DAYS * 24 * 60 * 60 * 1000).toISOString(),
      }
    };

    await fetch(GAS_API_URL, {
      method: "POST",
      body: JSON.stringify(metaPayload),
      headers: { "Content-Type": "application/json" },
    });
  } catch (e) {
    // Metadata save failure is non-critical — file is still uploaded
    console.error("Failed to save metadata to GAS:", e);
  }

  return jsonResponse({
    success: true,
    message: "File berhasil dikirim ke tim multimedia!",
    r2Key: r2Key,
    downloadUrl: downloadUrl,
  });
}

// ============================================================
// DOWNLOAD HANDLER
// ============================================================
async function handleDownload(request, env, r2Key) {
  if (!r2Key) {
    return jsonResponse({ success: false, message: "File tidak ditemukan." }, 404);
  }

  const object = await env.MULTIMEDIA_BUCKET.get(r2Key);
  if (!object) {
    return jsonResponse({ success: false, message: "File tidak ditemukan atau sudah kedaluwarsa." }, 404);
  }

  const headers = new Headers();
  object.writeHttpMetadata(headers);
  headers.set("Access-Control-Allow-Origin", "*");
  // Force download
  if (!headers.has("Content-Disposition")) {
    const meta = object.customMetadata || {};
    const filename = meta.originalName || r2Key;
    headers.set("Content-Disposition", `attachment; filename="${filename}"`);
  }

  return new Response(object.body, { headers });
}

// ============================================================
// DELETE HANDLER (Admin only — requires password)
// ============================================================
async function handleDelete(request, env) {
  let body;
  try {
    body = await request.json();
  } catch (e) {
    return jsonResponse({ success: false, message: "Format request tidak valid." }, 400);
  }

  const { r2Key, password } = body;
  if (!r2Key || !password) {
    return jsonResponse({ success: false, message: "Data tidak lengkap." }, 400);
  }

  // Verify password against GAS before deleting
  try {
    const verifyRes = await fetch(GAS_API_URL, {
      method: "POST",
      body: JSON.stringify({ action: "verifyPassword", password }),
      headers: { "Content-Type": "application/json" },
    });
    const verifyData = await verifyRes.json();
    if (!verifyData.success) {
      return jsonResponse({ success: false, message: "Akses ditolak. Password salah." }, 403);
    }
  } catch (e) {
    return jsonResponse({ success: false, message: "Gagal memverifikasi akses." }, 500);
  }

  // Delete from R2
  try {
    await env.MULTIMEDIA_BUCKET.delete(r2Key);
  } catch (e) {
    return jsonResponse({ success: false, message: "Gagal menghapus file dari penyimpanan." }, 500);
  }

  // Delete metadata from GAS
  try {
    await fetch(GAS_API_URL, {
      method: "POST",
      body: JSON.stringify({ action: "deleteMultimediaFile", r2Key, password }),
      headers: { "Content-Type": "application/json" },
    });
  } catch (e) {
    console.error("Failed to delete metadata from GAS:", e);
  }

  return jsonResponse({ success: true, message: "File berhasil dihapus." });
}

// ============================================================
// CLEANUP EXPIRED FILES (called by cron or on-demand)
// ============================================================
async function handleCleanupExpired(env) {
  const cutoffDate = new Date(Date.now() - RETENTION_DAYS * 24 * 60 * 60 * 1000);
  let deleted = 0;

  try {
    const list = await env.MULTIMEDIA_BUCKET.list();
    for (const obj of list.objects) {
      const meta = obj.customMetadata || {};
      const uploadedAt = meta.uploadedAt ? new Date(meta.uploadedAt) : null;
      if (uploadedAt && uploadedAt < cutoffDate) {
        await env.MULTIMEDIA_BUCKET.delete(obj.key);
        deleted++;
      }
    }
  } catch (e) {
    console.error("Cleanup error:", e);
  }

  return jsonResponse({ success: true, deleted });
}

// ============================================================
// MAIN ENTRY POINT
// ============================================================
export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const { pathname } = url;

    // Handle CORS preflight
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: CORS_HEADERS });
    }

    // API routes
    if (pathname === "/api/upload" && request.method === "POST") {
      return handleUpload(request, env);
    }

    if (pathname.startsWith("/api/download/") && request.method === "GET") {
      const r2Key = decodeURIComponent(pathname.replace("/api/download/", ""));
      return handleDownload(request, env, r2Key);
    }

    if (pathname === "/api/files/delete" && request.method === "POST") {
      return handleDelete(request, env);
    }

    if (pathname === "/api/cleanup" && request.method === "POST") {
      return handleCleanupExpired(env);
    }

    // Pass through to static assets (the index.html app)
    return env.ASSETS.fetch(request);
  },

  // Scheduled cleanup: runs daily at midnight UTC
  async scheduled(event, env, ctx) {
    ctx.waitUntil(handleCleanupExpired(env));
  },
};
