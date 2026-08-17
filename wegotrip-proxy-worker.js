// wegotrip-proxy-worker.js
//
// Why this exists:
// WeGoTrip's public partner API (app.wegotrip.com/api/v2) doesn't send CORS
// headers, so a browser calling it directly from your app gets silently
// blocked — that's why the Activities tab showed "0 cities loaded". This
// worker calls WeGoTrip server-side (where CORS doesn't apply) and re-sends
// the response with CORS opened up for your app.
//
// SETUP (5 minutes, free):
// 1. Go to https://dash.cloudflare.com -> Workers & Pages -> Create -> Create Worker.
// 2. Delete the default starter code, paste in everything below, click "Deploy".
// 3. Copy the worker's URL — it looks like:
//    https://wegotrip-proxy.<your-subdomain>.workers.dev
// 4. In App.jsx, set WEGOTRIP_PROXY_BASE to that exact URL (see the comment
//    next to WEGOTRIP_API in App.jsx).
//
// That's it — city search and activity listings will start returning real
// data once the URL is in place.

const WEGOTRIP_API = "https://app.wegotrip.com/api/v2";

// Only these documented WeGoTrip endpoints are relayed, so the worker can't
// be used as an open proxy to arbitrary hosts.
const ALLOWED_PREFIXES = [
  "/cities/",
  "/products/",
  "/search/",
  "/languages/",
  "/currencies/",
  "/countries/",
  "/attractions/",
];

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Accept-Language",
  };
}

export default {
  async fetch(request) {
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders() });
    }

    const url = new URL(request.url);
    // Expecting requests shaped like /api/v2/<wegotrip path>?<query>
    const path = url.pathname.replace(/^\/api\/v2/, "");

    if (!ALLOWED_PREFIXES.some((p) => path.startsWith(p))) {
      return new Response(JSON.stringify({ error: "Path not allowed" }), {
        status: 403,
        headers: { "Content-Type": "application/json", ...corsHeaders() },
      });
    }

    const target = `${WEGOTRIP_API}${path}${url.search}`;

    try {
      const upstream = await fetch(target, {
        headers: {
          "Accept-Language": request.headers.get("Accept-Language") || "en",
        },
      });
      const body = await upstream.text();
      return new Response(body, {
        status: upstream.status,
        headers: {
          "Content-Type": upstream.headers.get("Content-Type") || "application/json",
          ...corsHeaders(),
        },
      });
    } catch (e) {
      return new Response(JSON.stringify({ error: "Upstream request failed" }), {
        status: 502,
        headers: { "Content-Type": "application/json", ...corsHeaders() },
      });
    }
  },
};
