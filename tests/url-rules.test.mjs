import test from "node:test";
import assert from "node:assert/strict";
import { buildOfficeUri, extractDocumentUrl, isDocAspxShareLink, officeProtocolFor } from "../src/lib/url-rules.js";

const cfg = {
  enabled: true,
  allowedOrigins: ["https://workspace.example.gov.au"],
  allowedPathPrefixes: ["/sites/workspaces/"],
  openMode: "edit"
};

test("maps common Office extensions", () => {
  assert.equal(officeProtocolFor("https://x/a.docx"), "ms-word");
  assert.equal(officeProtocolFor("https://x/a.xlsx"), "ms-excel");
  assert.equal(officeProtocolFor("https://x/a.pptx"), "ms-powerpoint");
  assert.equal(officeProtocolFor("https://x/a.pdf"), null);
});

test("builds Office edit URI", () => {
  const u = "https://workspace.example.gov.au/sites/workspaces/Docs/Test.docx";
  assert.equal(buildOfficeUri(u, "edit"), `ms-word:ofe|u|${u}`);
});

test("accepts a direct approved document URL", () => {
  const u = "https://workspace.example.gov.au/sites/workspaces/Docs/Test.docx";
  assert.equal(extractDocumentUrl(u, cfg), u);
});

test("rejects a different origin", () => {
  const u = "https://evil.example/sites/workspaces/Docs/Test.docx";
  assert.equal(extractDocumentUrl(u, cfg), null);
});

test("rejects paths outside configured scope", () => {
  const u = "https://workspace.example.gov.au/personal/user/Docs/Test.docx";
  assert.equal(extractDocumentUrl(u, cfg), null);
});

test("extracts SourceUrl from SharePoint download links", () => {
  const u = "https://workspace.example.gov.au/sites/workspaces/_layouts/15/download.aspx?SourceUrl=%2Fsites%2Fworkspaces%2FDocs%2FTest.docx";
  assert.equal(extractDocumentUrl(u, cfg), "https://workspace.example.gov.au/sites/workspaces/Docs/Test.docx");
});

test("does not guess a path from Doc.aspx file=name only", () => {
  const u = "https://workspace.example.gov.au/sites/workspaces/_layouts/15/Doc.aspx?sourcedoc=%7B11111111-1111-1111-1111-111111111111%7D&file=Test.docx&action=default";
  assert.equal(extractDocumentUrl(u, cfg), null);
  assert.equal(isDocAspxShareLink(u, cfg), true);
});

test("accepts real GRDC TEST direct SharePoint document link and preserves query string", () => {
  const grdcCfg = {
    enabled: true,
    allowedOrigins: [
      "https://workspaces.grdc.com.au",
      "https://workspaces.test.grdc.com.au",
      "https://workspaces.dev.grdc.com.au"
    ],
    allowedPathPrefixes: ["/"],
    openMode: "edit"
  };
  const u = "https://workspaces.test.grdc.com.au/sites/procurement-9177490/Shared%20Documents/Gone%20To%20Market/PROC-9177490%20Application%20Planning%20Document%20-%20Not%20for%20Submission.docx?e=4%3A047420abc19b46f3a6d1e3755d541d88";
  assert.equal(officeProtocolFor(u), "ms-word");
  assert.equal(extractDocumentUrl(u, grdcCfg), u);
  assert.equal(buildOfficeUri(u, "edit"), `ms-word:ofe|u|${u}`);
});
