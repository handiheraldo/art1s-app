// ============================================================
// ART1S Church App Worker
// Handles static assets routing
// ============================================================

export default {
  async fetch(request, env, ctx) {
    // Pass through to static assets (the index.html app)
    return env.ASSETS.fetch(request);
  },
};
