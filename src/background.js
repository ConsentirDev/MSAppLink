import { buildOfficeUri, extractDocumentUrl } from "./lib/url-rules.js";

let cachedConfig = null;
let configPromise = null;
const launched = new Map();

function log(cfg, ...args) {
  if (cfg?.debug) console.log("[SP Office Launcher]", ...args);
}

async function loadConfig() {
  if (cachedConfig) return cachedConfig;
  if (configPromise) return configPromise;
  configPromise = (async () => {
    const packaged = await fetch(chrome.runtime.getURL("runtime-config.json")).then(r => r.json());
    let managed = {};
    let local = {};
    try { managed = await chrome.storage.managed.get(null); } catch { managed = {}; }
    try { local = await chrome.storage.local.get(null); } catch { local = {}; }
    const cfg = {
      enabled: managed.enabled ?? local.enabled ?? true,
      allowedOrigins: managed.allowedOrigins ?? local.allowedOrigins ?? packaged.origins,
      allowedPathPrefixes: managed.allowedPathPrefixes ?? local.allowedPathPrefixes ?? packaged.pathPrefixes ?? ["/"],
      openMode: managed.openMode ?? local.openMode ?? packaged.openMode ?? "edit",
      resolveDocShareLinks: managed.resolveDocShareLinks ?? local.resolveDocShareLinks ?? packaged.resolveDocShareLinks ?? true,
      debug: managed.debug ?? local.debug ?? packaged.debug ?? false
    };
    cfg.allowedOrigins = cfg.allowedOrigins.map(x => new URL(x).origin.toLowerCase());
    cachedConfig = cfg;
    return cfg;
  })();
  return configPromise;
}

chrome.storage?.onChanged?.addListener((changes, area) => {
  if (area === "managed" || area === "local") {
    cachedConfig = null;
    configPromise = null;
  }
});

function recentlyLaunched(tabId, sourceUrl) {
  const key = `${tabId}|${sourceUrl}`;
  const now = Date.now();
  const last = launched.get(key) || 0;
  launched.set(key, now);
  for (const [k, t] of launched) if (now - t > 5000) launched.delete(k);
  return now - last < 1500;
}

async function launch(tabId, sourceUrl, documentUrl, cfg) {
  if (tabId < 0 || recentlyLaunched(tabId, sourceUrl)) return;
  const officeUri = buildOfficeUri(documentUrl, cfg.openMode);
  if (!officeUri) return;
  log(cfg, "Launching", officeUri, "from", sourceUrl);
  try {
    await chrome.tabs.update(tabId, { url: officeUri });
  } catch (err) {
    log(cfg, "Direct protocol navigation failed; using fallback launcher", err?.message || err);
    const fallback = new URL(chrome.runtime.getURL("launcher.html"));
    fallback.searchParams.set("target", officeUri);
    fallback.searchParams.set("source", sourceUrl);
    fallback.searchParams.set("document", documentUrl);
    await chrome.tabs.update(tabId, { url: fallback.href });
  }
}

// Enterprise mode: policy-installed MV3 extensions may retain webRequestBlocking.
if (chrome.webRequest?.onBeforeRequest) {
  try {
    chrome.webRequest.onBeforeRequest.addListener(
      (details) => {
        if (details.type !== "main_frame" || details.tabId < 0) return;
        // Blocking listeners must return synchronously. Config is primed at service worker startup.
        if (!cachedConfig) return;
        const docUrl = extractDocumentUrl(details.url, cachedConfig);
        if (!docUrl) return;
        queueMicrotask(() => launch(details.tabId, details.url, docUrl, cachedConfig));
        return { cancel: true };
      },
      { urls: ["https://*/*"], types: ["main_frame"] },
      ["blocking"]
    );
  } catch (err) {
    console.warn("[SP Office Launcher] webRequest blocking unavailable; using webNavigation fallback.", err?.message || err);
  }
}

// Development/fallback mode: observe the navigation and immediately replace it.
if (chrome.webNavigation?.onBeforeNavigate) {
  chrome.webNavigation.onBeforeNavigate.addListener(async (details) => {
    if (details.frameId !== 0 || details.tabId < 0) return;
    const cfg = await loadConfig();
    const docUrl = extractDocumentUrl(details.url, cfg);
    if (!docUrl) return;
    await launch(details.tabId, details.url, docUrl, cfg);
  });
}

// Prime configuration so the synchronous enterprise listener can operate.
loadConfig().then(cfg => log(cfg, "Loaded config", cfg)).catch(err => console.error("[SP Office Launcher] Config load failed", err));
