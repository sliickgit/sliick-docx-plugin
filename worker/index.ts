/**
 * Cloudflare Worker for the Sliick Docs Office add-in.
 *
 * Responsibilities:
 *  - /api/token : server-side OAuth token exchange proxy. Salesforce's
 *    /services/oauth2/* endpoints do not return CORS headers, so a browser
 *    cannot call them directly (fetch throws "Load failed"). The task pane
 *    POSTs the form body here (same origin); we forward it to the org's token
 *    endpoint and pass the JSON response straight back.
 *  - everything else : served from the static asset bundle (./dist) via the
 *    ASSETS binding, preserving the extensionless-URL behavior the manifest and
 *    auth dialog depend on.
 */

interface Env {
  ASSETS: { fetch: (request: Request) => Promise<Response> };
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);
    if (url.pathname === "/api/token") {
      return handleTokenExchange(request, url);
    }
    return env.ASSETS.fetch(request);
  },
};

/**
 * Proxy the OAuth token request to the Salesforce org named in `?org=`.
 * The org host is validated to be an https *.salesforce.com domain so this
 * can't be abused as an open relay.
 */
async function handleTokenExchange(request: Request, url: URL): Promise<Response> {
  if (request.method !== "POST") {
    return jsonError("method_not_allowed", "Use POST.", 405);
  }

  const org = url.searchParams.get("org") ?? "";
  let target: URL;
  try {
    target = new URL(org);
  } catch {
    return jsonError("invalid_org", "Missing or malformed org URL.", 400);
  }
  if (target.protocol !== "https:" || !/\.salesforce\.com$/i.test(target.hostname)) {
    return jsonError("invalid_org_host", "org must be an https *.salesforce.com URL.", 400);
  }

  const body = await request.text();
  const sfResp = await fetch(`${target.origin}/services/oauth2/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  // Pass Salesforce's status + JSON body through unchanged. Same-origin call,
  // so no CORS headers are needed on this response.
  const text = await sfResp.text();
  return new Response(text, {
    status: sfResp.status,
    headers: {
      "Content-Type": sfResp.headers.get("Content-Type") ?? "application/json",
      "Cache-Control": "no-store",
    },
  });
}

function jsonError(error: string, description: string, status: number): Response {
  return new Response(JSON.stringify({ error, error_description: description }), {
    status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
  });
}
