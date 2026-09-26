import { describe, expect, test } from "bun:test";

import { contentSecurityPolicy, createNonce, shellHtml } from "./shell";

describe("contentSecurityPolicy", () => {
  test("admits only the nonced script and inlined fonts and images", () => {
    const policy = contentSecurityPolicy("abc123");

    expect(policy.split("; ")).toEqual([
      "default-src 'none'",
      "script-src 'nonce-abc123'",
      "style-src 'unsafe-inline'",
      "font-src data:",
      "img-src data:",
    ]);
  });
});

describe("createNonce", () => {
  test("is fresh each time", () => {
    const nonces = new Set(Array.from({ length: 20 }, createNonce));

    expect(nonces.size).toBe(20);
    for (const nonce of nonces) expect(nonce).toMatch(/^[A-Za-z0-9+/]{24}$/u);
  });
});

describe("shellHtml", () => {
  const html = shellHtml({
    nonce: "n0nce",
    scriptUri: "https://file+.vscode-resource.vscode-cdn.net/ext/dist/webview/main.js",
    fileName: 'Q3 "draft" <final>.docx',
  });

  test("carries the policy and nonces its one script", () => {
    expect(html).toContain(
      `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'nonce-n0nce';`,
    );
    expect(html.match(/<script/gu)).toHaveLength(1);
    expect(html).toContain(
      '<script nonce="n0nce" src="https://file+.vscode-resource.vscode-cdn.net/ext/dist/webview/main.js">',
    );
  });

  test("escapes the file name", () => {
    expect(html).toContain("Q3 &quot;draft&quot; &lt;final&gt;.docx");
    expect(html).not.toContain("<final>");
  });

  test("loads nothing from the network", () => {
    expect(html).not.toMatch(/(?:src|href)="http:/u);
    expect(html).not.toContain("<link");
  });
});
