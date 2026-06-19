/**
 * Cloudflare Worker for the Sliick Docs Office add-in — static asset host.
 *
 * Serves the Vite build (./dist) via the ASSETS binding, preserving the
 * extensionless-URL behavior the manifest and auth dialog depend on.
 *
 * The OAuth token exchange runs directly in the browser against the org's
 * /services/oauth2/token (the org enables "CORS for OAuth Endpoints" with
 * office.sliick.com allowlisted), so there is no server-side token proxy here.
 */

interface Env {
  ASSETS: { fetch: (request: Request) => Promise<Response> };
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    return env.ASSETS.fetch(request);
  },
};
