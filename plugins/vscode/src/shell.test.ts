import { describe, expect, test } from "bun:test";

import { createNonce, editorContentSecurityPolicy, editorShellHtml } from "./shell";

describe("createNonce", () => {
  test("is fresh each time", () => {
    const nonces = new Set(Array.from({ length: 20 }, createNonce));

    expect(nonces.size).toBe(20);
    for (const nonce of nonces) expect(nonce).toMatch(/^[A-Za-z0-9+/]{24}$/u);
  });
});

describe("editorShellHtml", () => {
  const source = "https://file+.vscode-resource.vscode-cdn.net";
  const options = {
    nonce: "n0nce",
    cspSource: source,
    scriptUri: `${source}/ext/dist/editor/editor.js`,
    styleUri: `${source}/ext/dist/editor/editor.css`,
    fileName: 'Q3 "draft".docx',
  };
  const html = editorShellHtml(options);

  test("admits the bundle, its stylesheet and fonts, and no worker or fetch", () => {
    expect(editorContentSecurityPolicy("n0nce", source).split("; ")).toEqual([
      "default-src 'none'",
      `script-src 'nonce-n0nce' ${source}`,
      `style-src ${source} 'unsafe-inline'`,
      `font-src ${source} data: blob:`,
      `img-src ${source} data: blob:`,
      "worker-src 'none'",
      "connect-src 'none'",
    ]);
    expect(html).toContain(`content="default-src 'none'; script-src 'nonce-n0nce' ${source};`);
  });

  test("links the stylesheet and nonces the one script", () => {
    expect(html).toContain(`<link rel="stylesheet" href="${source}/ext/dist/editor/editor.css">`);
    expect(html.match(/<script/gu)).toHaveLength(1);
    expect(html).toContain(`<script nonce="n0nce" src="${source}/ext/dist/editor/editor.js">`);
    expect(html).toContain("Q3 &quot;draft&quot;.docx");
    expect(html).not.toContain("folio-test-type");
  });

  test("carries the test script only in test mode, nonced", () => {
    const testHtml = editorShellHtml({ ...options, testMode: true });

    expect(testHtml.match(/<script nonce="n0nce"/gu)).toHaveLength(2);
    expect(testHtml).toContain("folio-test-type");
  });
});
