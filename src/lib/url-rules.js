const OFFICE_TYPES = new Map([
  ["doc", "ms-word"], ["docx", "ms-word"], ["docm", "ms-word"], ["dot", "ms-word"], ["dotx", "ms-word"], ["dotm", "ms-word"], ["rtf", "ms-word"],
  ["xls", "ms-excel"], ["xlsx", "ms-excel"], ["xlsm", "ms-excel"], ["xlsb", "ms-excel"], ["xlt", "ms-excel"], ["xltx", "ms-excel"], ["xltm", "ms-excel"], ["csv", "ms-excel"],
  ["ppt", "ms-powerpoint"], ["pptx", "ms-powerpoint"], ["pptm", "ms-powerpoint"], ["pps", "ms-powerpoint"], ["ppsx", "ms-powerpoint"], ["ppsm", "ms-powerpoint"], ["pot", "ms-powerpoint"], ["potx", "ms-powerpoint"], ["potm", "ms-powerpoint"],
  ["vsd", "ms-visio"], ["vsdx", "ms-visio"], ["vsdm", "ms-visio"]
]);

export const OFFICE_PROTOCOLS = new Set([...new Set(OFFICE_TYPES.values())]);

export function normaliseOrigin(value) {
  const u = new URL(value);
  if (u.protocol !== "https:") throw new Error(`Only HTTPS origins are supported: ${value}`);
  return u.origin.toLowerCase();
}

export function extensionOf(urlOrPath) {
  let pathname = urlOrPath;
  try { pathname = new URL(urlOrPath).pathname; } catch { /* path only */ }
  const leaf = pathname.split("/").pop() || "";
  const idx = leaf.lastIndexOf(".");
  return idx >= 0 ? leaf.slice(idx + 1).toLowerCase() : "";
}

export function officeProtocolFor(urlOrPath) {
  return OFFICE_TYPES.get(extensionOf(urlOrPath)) || null;
}

export function buildOfficeUri(documentUrl, mode = "edit") {
  const protocol = officeProtocolFor(documentUrl);
  if (!protocol) return null;
  const command = mode === "view" ? "ofv" : "ofe";
  return `${protocol}:${command}|u|${documentUrl}`;
}

function decodeSafely(value) {
  let current = value;
  for (let i = 0; i < 2; i += 1) {
    try {
      const decoded = decodeURIComponent(current);
      if (decoded === current) break;
      current = decoded;
    } catch {
      break;
    }
  }
  return current;
}

function resolveCandidate(baseUrl, rawValue) {
  if (!rawValue) return null;
  const value = decodeSafely(rawValue).trim();
  try {
    const absolute = new URL(value, baseUrl.origin);
    if (absolute.origin !== baseUrl.origin) return null;
    return absolute.href;
  } catch {
    return null;
  }
}

function allowedByConfig(url, cfg) {
  if (!cfg.enabled) return false;
  const origin = url.origin.toLowerCase();
  if (!cfg.allowedOrigins.map(x => x.toLowerCase()).includes(origin)) return false;
  const prefixes = cfg.allowedPathPrefixes?.length ? cfg.allowedPathPrefixes : ["/"];
  return prefixes.some(prefix => url.pathname.startsWith(prefix));
}

/**
 * Return a concrete Office document URL if the navigation is safe to convert.
 * Deliberately does not treat Doc.aspx?sourcedoc={GUID}&file=name.docx as a
 * concrete URL: a filename alone is insufficient to locate the file safely.
 */
export function extractDocumentUrl(rawUrl, cfg) {
  let url;
  try { url = new URL(rawUrl); } catch { return null; }
  if (url.protocol !== "https:") return null;
  if (!allowedByConfig(url, cfg)) return null;

  if (officeProtocolFor(url.href)) return url.href;

  const candidateKeys = [
    "SourceUrl", "sourceurl", "sourceUrl", "FileRef", "fileref",
    "RootFolder", "rootfolder", "id", "url"
  ];
  for (const key of candidateKeys) {
    const raw = url.searchParams.get(key);
    const candidate = resolveCandidate(url, raw);
    if (candidate && officeProtocolFor(candidate)) return candidate;
  }

  // 'file=' is useful only when it contains a path/URL, not just a filename.
  const fileValue = url.searchParams.get("file");
  if (fileValue && /[\\/]/.test(decodeSafely(fileValue))) {
    const candidate = resolveCandidate(url, fileValue);
    if (candidate && officeProtocolFor(candidate)) return candidate;
  }

  return null;
}

export function isDocAspxShareLink(rawUrl, cfg) {
  let url;
  try { url = new URL(rawUrl); } catch { return false; }
  if (!allowedByConfig(url, cfg)) return false;
  if (!/\/_layouts\/15\/doc\.aspx$/i.test(url.pathname)) return false;
  return Boolean(url.searchParams.get("sourcedoc") && url.searchParams.get("file") && officeProtocolFor(url.searchParams.get("file")));
}

export function siteBaseForDocAspx(rawUrl) {
  const url = new URL(rawUrl);
  const marker = "/_layouts/15/doc.aspx";
  const lower = url.pathname.toLowerCase();
  const idx = lower.lastIndexOf(marker);
  if (idx < 0) return null;
  const webPath = url.pathname.slice(0, idx) || "";
  return `${url.origin}${webPath}`;
}

export function sourcedocGuid(rawUrl) {
  const url = new URL(rawUrl);
  const raw = url.searchParams.get("sourcedoc") || "";
  return raw.replace(/[{}]/g, "").trim();
}
