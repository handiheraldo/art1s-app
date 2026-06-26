// ============================================================
// ART1S Church App — Cloudflare Pages Function
// Handles /api route: caching proxy to Google Apps Script
// ============================================================

const GAS_API_URL = "https://script.google.com/macros/s/AKfycbyT8u3ESo5wuQr9w50bUCFd_al5tau5aTTIVtmh7_jDifkoO8Jp_7y86JuFhsNdux3bOg/exec";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, HEAD, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

function addCorsHeaders(response) {
  const newResponse = new Response(response.body, response);
  Object.entries(corsHeaders).forEach(([key, value]) => {
    newResponse.headers.set(key, value);
  });
  return newResponse;
}

// --- CORS Preflight ---
export async function onRequestOptions() {
  return new Response(null, { status: 204, headers: corsHeaders });
}

// --- GET: Cached proxy to GAS ---
export async function onRequestGet(context) {
  const { request } = context;
  const url = new URL(request.url);
  const cache = caches.default;
  const targetUrl = new URL(GAS_API_URL);
  targetUrl.search = url.search;

  // Check Cloudflare Cache
  let cachedResponse = await cache.match(request);
  if (cachedResponse) {
    const headers = new Headers(cachedResponse.headers);
    headers.set("X-Cache", "HIT");
    return addCorsHeaders(new Response(cachedResponse.body, {
      status: cachedResponse.status,
      statusText: cachedResponse.statusText,
      headers
    }));
  }

  // Cache miss: Fetch from Google Apps Script
  try {
    const response = await fetch(targetUrl.toString(), {
      method: "GET",
      headers: request.headers,
    });

    if (response.ok) {
      const cacheResponse = new Response(response.body, {
        status: response.status,
        statusText: response.statusText,
        headers: response.headers,
      });

      // Cache for 1 hour (3600 seconds)
      cacheResponse.headers.set("Cache-Control", "public, max-age=3600");

      // Store in Cloudflare edge cache
      context.waitUntil(cache.put(request, cacheResponse.clone()));

      const clientHeaders = new Headers(cacheResponse.headers);
      clientHeaders.set("X-Cache", "MISS");
      return addCorsHeaders(new Response(cacheResponse.body, {
        status: cacheResponse.status,
        statusText: cacheResponse.statusText,
        headers: clientHeaders,
      }));
    }

    return addCorsHeaders(response);
  } catch (error) {
    return new Response(JSON.stringify({ error: "Failed to fetch from backend: " + error.message }), {
      status: 500,
      headers: { "Content-Type": "application/json", ...corsHeaders },
    });
  }
}

// --- POST: Forward to GAS + invalidate cache ---
export async function onRequestPost(context) {
  const { request } = context;
  const url = new URL(request.url);
  const cache = caches.default;
  const targetUrl = new URL(GAS_API_URL);
  targetUrl.search = url.search;

  try {
    const response = await fetch(targetUrl.toString(), {
      method: "POST",
      headers: request.headers,
      body: await request.clone().text(),
    });

    if (response.ok) {
      // Invalidate cached GET /api?action=getData after successful POST
      const getRequestUrl = new URL(url.origin + "/api");
      getRequestUrl.searchParams.set("action", "getData");
      const cacheKey = new Request(getRequestUrl.toString(), { method: "GET" });
      context.waitUntil(cache.delete(cacheKey));
    }

    return addCorsHeaders(response);
  } catch (error) {
    return new Response(JSON.stringify({ error: "Failed to post to backend: " + error.message }), {
      status: 500,
      headers: { "Content-Type": "application/json", ...corsHeaders },
    });
  }
}
