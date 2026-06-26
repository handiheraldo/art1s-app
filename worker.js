// ============================================================
// ART1S Church App Worker
// Handles static assets routing and API caching
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

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    // Intercept requests to /api
    if (url.pathname === "/api") {
      // Handle CORS preflight options request
      if (request.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: corsHeaders,
        });
      }

      const cache = caches.default;
      const targetUrl = new URL(GAS_API_URL);
      targetUrl.search = url.search;

      // We will cache GET requests
      if (request.method === "GET") {
        // Match request in Cloudflare Cache
        let cachedResponse = await cache.match(request);
        if (cachedResponse) {
          // Add a header to indicate cache hit (useful for debugging)
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
            // Clone the response to modify headers and save to cache
            // Cloudflare Cache requires Cache-Control headers to cache the response.
            const cacheResponse = new Response(response.body, {
              status: response.status,
              statusText: response.statusText,
              headers: response.headers,
            });

            // Set Cache-Control header to cache for 1 hour (3600 seconds)
            cacheResponse.headers.set("Cache-Control", "public, max-age=3600");

            // Store the cloned response in the cache
            ctx.waitUntil(cache.put(request, cacheResponse.clone()));

            // Return response to client with HIT/MISS indicator
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
            headers: { 
              "Content-Type": "application/json",
              ...corsHeaders
            },
          });
        }
      }

      // Handle POST requests (updating/saving data)
      if (request.method === "POST") {
        try {
          // Forward the POST request to GAS
          const response = await fetch(targetUrl.toString(), {
            method: "POST",
            headers: request.headers,
            body: await request.clone().text(),
          });

          if (response.ok) {
            // Invalidate the cache for GET requests to /api (regardless of query params)
            const getRequestUrl = new URL(url.origin + "/api");
            getRequestUrl.searchParams.set("action", "getData");
            const cacheKey = new Request(getRequestUrl.toString(), { method: "GET" });
            ctx.waitUntil(cache.delete(cacheKey));
          }

          return addCorsHeaders(response);
        } catch (error) {
          return new Response(JSON.stringify({ error: "Failed to post to backend: " + error.message }), {
            status: 500,
            headers: { 
              "Content-Type": "application/json",
              ...corsHeaders
            },
          });
        }
      }
    }

    // Pass through to static assets (the index.html app)
    return env.ASSETS.fetch(request);
  },
};
