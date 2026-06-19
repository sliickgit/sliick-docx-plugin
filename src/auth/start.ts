/**
 * Auth-dialog bootstrap page, opened inside the Office auth dialog.
 *
 * Office's displayDialogAsync only allows opening a URL on our own domain (the
 * source location / AppDomains). The Salesforce authorize endpoint lives on a
 * per-org *.my.salesforce.com domain we can't enumerate in the manifest, so we
 * open THIS page (same origin, always allowed) and immediately redirect to the
 * authorize URL passed in `?u=`. AppDomains governs only the *initial* dialog
 * URL; navigation after the dialog opens is unrestricted.
 */

function isSalesforceAuthorizeUrl(raw: string): boolean {
  try {
    const u = new URL(raw);
    return (
      u.protocol === "https:" &&
      /\.salesforce\.com$/i.test(u.hostname) &&
      u.pathname === "/services/oauth2/authorize"
    );
  } catch {
    return false;
  }
}

const target = new URLSearchParams(window.location.search).get("u") ?? "";
if (isSalesforceAuthorizeUrl(target)) {
  window.location.replace(target);
} else {
  const msg = document.getElementById("msg");
  if (msg) msg.textContent = "Invalid sign-in request. Close this window and try again.";
}
