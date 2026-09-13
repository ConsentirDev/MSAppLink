const $ = id => document.getElementById(id);
async function load() {
  const packaged = await fetch(chrome.runtime.getURL("runtime-config.json")).then(r => r.json());
  const local = await chrome.storage.local.get(null);
  $("origin").value = (local.allowedOrigins ?? packaged.origins ?? [""])[0] || "";
  $("path").value = (local.allowedPathPrefixes ?? packaged.pathPrefixes ?? ["/"])[0] || "/";
  $("mode").value = local.openMode ?? packaged.openMode ?? "edit";
  $("resolve").checked = local.resolveDocShareLinks ?? packaged.resolveDocShareLinks ?? true;
  $("debug").checked = local.debug ?? packaged.debug ?? false;
  $("enabled").checked = local.enabled ?? true;
}
$("save").addEventListener("click", async () => {
  const status = $("status");
  try {
    const origin = new URL($("origin").value.trim());
    if (origin.protocol !== "https:") throw new Error("Origin must use HTTPS");
    await chrome.storage.local.set({
      allowedOrigins: [origin.origin],
      allowedPathPrefixes: [$("path").value.trim() || "/"],
      openMode: $("mode").value,
      resolveDocShareLinks: $("resolve").checked,
      debug: $("debug").checked,
      enabled: $("enabled").checked
    });
    status.textContent = "Saved.";
    setTimeout(() => status.textContent = "", 1800);
  } catch (err) {
    status.textContent = err.message || String(err);
  }
});
load();
