import { readFile, writeFile, rm, mkdir, cp } from "node:fs/promises";
import path from "node:path";

const configPath = process.argv[2] || "config/targets.json";
const mode = process.argv[3] || "enterprise";
if (!["enterprise", "dev"].includes(mode)) throw new Error("Mode must be 'enterprise' or 'dev'");
const cfg = JSON.parse(await readFile(configPath, "utf8"));
if (!Array.isArray(cfg.origins) || cfg.origins.length === 0) throw new Error("config.origins must contain at least one HTTPS origin");
const origins = cfg.origins.map(value => {
  const u = new URL(value);
  if (u.protocol !== "https:") throw new Error(`Origin must use HTTPS: ${value}`);
  return u.origin;
});

const hostPermissions = mode === "dev" ? ["https://*/*"] : origins.map(x => `${x}/*`);
const matches = mode === "dev" ? ["https://*/*"] : origins.map(x => `${x}/*`);

// IMPORTANT: webRequestBlocking is valid in Manifest V3 only when the extension
// is installed by enterprise policy. Unpacked/developer-mode MV3 extensions
// must not request it or Chrome will reject the manifest.
const permissions = mode === "enterprise"
  ? ["storage", "tabs", "webNavigation", "webRequest", "webRequestBlocking"]
  : ["storage", "tabs", "webNavigation"];

const contentScripts = [{
  matches,
  js: ["doc-resolver.js"],
  run_at: "document_start"
}];

let manifest = await readFile("src/manifest.template.json", "utf8");
manifest = manifest
  .replace("__EXTENSION_NAME__", (cfg.name || "SharePoint Office Launcher").replaceAll('"', ""))
  .replace("__PERMISSIONS__", JSON.stringify(permissions, null, 2))
  .replace("__HOST_PERMISSIONS__", JSON.stringify(hostPermissions, null, 2))
  .replace("__CONTENT_SCRIPTS__", JSON.stringify(contentScripts, null, 2));

// Keep developer and enterprise artefacts separate so an enterprise manifest
// cannot accidentally be loaded unpacked during POC testing.
const outputDir = mode === "dev" ? "dist-dev" : "dist-enterprise";
await rm(outputDir, { recursive: true, force: true });
await mkdir(path.join(outputDir, "lib"), { recursive: true });
for (const file of ["background.js", "doc-resolver.js", "launcher.html", "launcher.js", "options.html", "options.js", "managed_schema.json"]) {
  await cp(path.join("src", file), path.join(outputDir, file));
}
await cp("src/lib/url-rules.js", path.join(outputDir, "lib/url-rules.js"));
await writeFile(path.join(outputDir, "manifest.json"), manifest);
await writeFile(path.join(outputDir, "runtime-config.json"), JSON.stringify({
  origins,
  pathPrefixes: cfg.pathPrefixes || ["/"],
  openMode: cfg.openMode || "edit",
  resolveDocShareLinks: cfg.resolveDocShareLinks ?? true,
  debug: cfg.debug ?? false,
  buildMode: mode
}, null, 2));
console.log(`Built ${mode} extension in ${outputDir}/ for: ${origins.join(", ")}`);
if (mode === "enterprise") {
  console.log("NOTE: dist-enterprise is for policy/force installation only; Chrome will reject webRequestBlocking if loaded unpacked.");
}
