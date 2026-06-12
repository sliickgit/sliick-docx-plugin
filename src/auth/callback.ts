/**
 * OAuth redirect landing page, opened inside the Office auth dialog.
 * Relays ?code/&state (or the error) back to the task pane via messageParent,
 * then the task pane closes the dialog and exchanges the code.
 */

Office.onReady(() => {
  const params = new URLSearchParams(window.location.search);
  const payload = JSON.stringify({
    code: params.get("code") ?? undefined,
    state: params.get("state") ?? undefined,
    error: params.get("error") ?? undefined,
    errorDescription: params.get("error_description") ?? undefined,
  });
  try {
    Office.context.ui.messageParent(payload);
  } catch {
    const msg = document.getElementById("msg");
    if (msg) {
      msg.textContent =
        "Sign-in finished, but this window couldn't talk to the task pane. Close it and try again.";
    }
  }
});
