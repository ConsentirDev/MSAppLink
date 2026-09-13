(async () => {
  function ext(name) {
    const m = String(name || "").match(/\.([A-Za-z0-9]+)$/);
    return m ? m[1].toLowerCase() : "";
  }
  const protocols = {
    doc: "ms-word", docx: "ms-word", docm: "ms-word", dot: "ms-word", dotx: "ms-word", dotm: "ms-word", rtf: "ms-word",
    xls: "ms-excel", xlsx: "ms-excel", xlsm: "ms-excel", xlsb: "ms-excel", xlt: "ms-excel", xltx: "ms-excel", xltm: "ms-excel", csv: "ms-excel",
    ppt: "ms-powerpoint", pptx: "ms-powerpoint", pptm: "ms-powerpoint", pps: "ms-powerpoint", ppsx: "ms-powerpoint", ppsm: "ms-powerpoint", pot: "ms-powerpoint", potx: "ms-powerpoint", potm: "ms-powerpoint",
    vsd: "ms-visio", vsdx: "ms-visio", vsdm: "ms-visio"
  };
  try {
    const cfg = await fetch(chrome.runtime.getURL("runtime-config.json")).then(r => r.json());
    let managed = {};
    let local = {};
    try { managed = await chrome.storage.managed.get(null); } catch { managed = {}; }
    try { local = await chrome.storage.local.get(null); } catch { local = {}; }
    if ((managed.enabled ?? local.enabled ?? true) === false || (managed.resolveDocShareLinks ?? local.resolveDocShareLinks ?? cfg.resolveDocShareLinks ?? true) === false) return;

    const u = new URL(location.href);
    const allowedOrigins = (managed.allowedOrigins ?? local.allowedOrigins ?? cfg.origins).map(x => new URL(x).origin.toLowerCase());
    if (!allowedOrigins.includes(u.origin.toLowerCase())) return;
    if (!/\/_layouts\/15\/doc\.aspx$/i.test(u.pathname)) return;
    const guid = (u.searchParams.get("sourcedoc") || "").replace(/[{}]/g, "").trim();
    const filename = u.searchParams.get("file") || "";
    const protocol = protocols[ext(filename)];
    if (!guid || !protocol) return;

    const marker = "/_layouts/15/doc.aspx";
    const idx = u.pathname.toLowerCase().lastIndexOf(marker);
    const webPath = idx >= 0 ? u.pathname.slice(0, idx) : "";
    const siteBase = `${u.origin}${webPath}`;
    const endpoint = `${siteBase}/_api/web/GetFileById(guid'${encodeURIComponent(guid)}')?$select=ServerRelativeUrl`;
    const response = await fetch(endpoint, {
      method: "GET",
      credentials: "same-origin",
      headers: { "Accept": "application/json;odata=verbose" }
    });
    if (!response.ok) return;
    const data = await response.json();
    const rel = data?.d?.ServerRelativeUrl || data?.ServerRelativeUrl;
    if (!rel) return;
    const documentUrl = new URL(rel, u.origin).href;
    const mode = (managed.openMode ?? local.openMode ?? cfg.openMode ?? "edit") === "view" ? "ofv" : "ofe";
    window.stop();
    location.replace(`${protocol}:${mode}|u|${documentUrl}`);
  } catch {
    // Fail open: SharePoint continues with its native behavior.
  }
})();
