import { expect, test } from "bun:test";
import { transform } from "lightningcss";
import { resolve } from "node:path";
import { readdir } from "node:fs/promises";
import { FONT_MAPPING } from "../../packages/core/src/utils/fontLoader";

import { scopeEditorCss } from "./scoped-editor-css";

const authoredFamilies = new Set(Object.keys(FONT_MAPPING).map((family) => family.toLowerCase()));

const assertNoAuthoredFontAliases = (source: string) => {
  for (const face of source.matchAll(/@font-face\s*\{(?<body>[^}]+)\}/gu)) {
    const family = face.groups?.body?.match(/font-family:\s*["']?(?<family>[^;"']+)/u)?.groups
      ?.family;
    expect(family).toBeDefined();
    expect(authoredFamilies.has(family?.trim().toLowerCase() ?? "")).toBe(false);
  }
};

test("every adapter stylesheet can be scoped without serializing declaration ASTs", async () => {
  let sheets = 0;
  for (const adapter of ["react", "vue"]) {
    const directory = resolve(import.meta.dir, `../../packages/${adapter}/src/styles`);
    for (const filename of await readdir(directory)) {
      if (!filename.endsWith(".css")) continue;
      const source = await Bun.file(resolve(directory, filename)).text();
      assertNoAuthoredFontAliases(source);
      const scoped = scopeEditorCss(source);
      expect(scopeEditorCss(scoped)).toBe(scoped);
      sheets += 1;
    }
  }
  expect(sheets).toBeGreaterThan(0);
});

test("the alias guard rejects every authored font family and allows scoped registrations", () => {
  for (const family of Object.keys(FONT_MAPPING)) {
    expect(() =>
      assertNoAuthoredFontAliases(
        `@font-face { font-family: "${family}"; src: local("Carlito"); }`,
      ),
    ).toThrow();
  }
  assertNoAuthoredFontAliases(
    '@font-face { font-family: "folio-document-Calibri"; src: url(face); }',
  );
});

test("all upstream ProseMirror selectors are scoped without changing declarations", async () => {
  const source = await Bun.file(
    Bun.resolveSync(
      "prosemirror-view/style/prosemirror.css",
      resolve(import.meta.dir, "../../packages/react"),
    ),
  ).text();
  const collect = (css: string) => {
    const selectors: string[] = [];
    const declarations: string[] = [];
    transform({
      filename: "fixture.css",
      code: Buffer.from(css),
      visitor: {
        Rule: {
          style(rule) {
            for (const selector of rule.value.selectors) selectors.push(JSON.stringify(selector));
            declarations.push(JSON.stringify(rule.value.declarations));
          },
        },
      },
    });
    return { selectors, declarations };
  };
  // Normalize both sides through the same optimizer (it merges repeated rules).
  const canonicalSource = new TextDecoder().decode(
    transform({ filename: "fixture.css", code: Buffer.from(source) }).code,
  );
  const original = collect(canonicalSource);
  const scoped = scopeEditorCss(source);
  const result = collect(scoped);
  expect(original.selectors.length).toBeGreaterThan(0);
  expect(result.selectors).toHaveLength(original.selectors.length);
  for (const selector of result.selectors) expect(selector).toContain('"name":"folio-root"');
  expect(result.declarations).toEqual(original.declarations);
  expect(scopeEditorCss(scoped)).toBe(scoped);
});

test("scopes every generic branch and preserves already scoped and portaled rules", () => {
  const css = `
    @media (min-width: 1px) {
      .ProseMirror, [contenteditable="true"]:focus { outline: none; }
      :is(.ProseMirror-focused, .ProseMirror-gapcursor) { color: red; }
      :is(.folio-root, .ProseMirror) { opacity: .5; }
    }
    :where(.folio-root) .ProseMirror { color: blue; }
    .folio-default-dialog-popup { position: fixed; }
  `;
  const scoped = scopeEditorCss(css);
  expect(scoped).toContain(':where(.folio-root) [contenteditable="true"]:focus');
  expect(scoped).toContain(":where(.folio-root) :is(.ProseMirror-focused, .ProseMirror-gapcursor)");
  expect(scoped).toContain(":where(.folio-root) :is(.folio-root, .ProseMirror)");
  expect(scoped).toContain(".folio-default-dialog-popup {");
  expect(scopeEditorCss(scoped)).toBe(scoped);
});

test("leaves sheets without generic selectors byte-for-byte intact", () => {
  const css =
    "/* bundled font attribution */\n@font-face { font-family: Folio; src: url(font.woff2); }";
  expect(scopeEditorCss(css)).toBe(css);
});
