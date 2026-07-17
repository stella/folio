// Security regression: `parseDOM`/`toDOM` must strip javascript:/data:/file:
// hrefs so pasted or programmatically-authored anchors never reach the live
// DOM as a navigable link, while internal bookmark anchors (`#name`) keep
// working unchanged.

import { describe, expect, test } from "bun:test";

import { HyperlinkExtension } from "./HyperlinkExtension";

// Minimal getAttribute-only stand-in — `getAttrs` only calls
// `dom.getAttribute`, so a full HTMLElement isn't needed to exercise it.
const fakeAnchorDom = (attrs: Record<string, string>): HTMLElement =>
  ({
    getAttribute: (name: string) => attrs[name] ?? null,
  }) as unknown as HTMLElement;

const getParseDomAttrs = (dom: HTMLElement) => {
  const rule = HyperlinkExtension().config.markSpec.parseDOM?.[0];
  if (!rule?.getAttrs) {
    throw new Error("HyperlinkExtension must define parseDOM[0].getAttrs");
  }
  return rule.getAttrs(dom) as { href: string } | false | null;
};

describe("HyperlinkExtension parseDOM — href sanitization", () => {
  test("strips javascript: hrefs to an empty href", () => {
    const executableHref = ["java", "script:alert(1)"].join("");
    const attrs = getParseDomAttrs(fakeAnchorDom({ href: executableHref }));
    expect(attrs).not.toBe(false);
    expect((attrs as { href: string }).href).toBe("");
  });

  test("strips data: hrefs to an empty href", () => {
    const attrs = getParseDomAttrs(fakeAnchorDom({ href: "data:text/html,<script></script>" }));
    expect((attrs as { href: string }).href).toBe("");
  });

  test("strips file: hrefs to an empty href", () => {
    const attrs = getParseDomAttrs(fakeAnchorDom({ href: "file:///etc/passwd" }));
    expect((attrs as { href: string }).href).toBe("");
  });

  test("keeps allow-listed https hrefs", () => {
    const attrs = getParseDomAttrs(fakeAnchorDom({ href: "https://example.com/doc" }));
    expect((attrs as { href: string }).href).toBe("https://example.com/doc");
  });

  test("keeps internal bookmark anchors", () => {
    const attrs = getParseDomAttrs(fakeAnchorDom({ href: "#bookmark1" }));
    expect((attrs as { href: string }).href).toBe("#bookmark1");
  });
});

describe("HyperlinkExtension toDOM — defense-in-depth href sanitization", () => {
  test("re-sanitizes an unsafe href already stored on the mark", () => {
    const spec = HyperlinkExtension().config.markSpec;
    if (!spec.toDOM) {
      throw new Error("HyperlinkExtension must define toDOM");
    }

    const executableHref = ["java", "script:alert(1)"].join("");
    const fakeMark = {
      type: { name: "hyperlink" },
      attrs: { href: executableHref, tooltip: null, rId: null, _docxHyperlinkIndex: null },
    } as Parameters<typeof spec.toDOM>[0];

    const output = spec.toDOM(fakeMark, true);
    if (!Array.isArray(output)) {
      throw new TypeError("Expected DOMOutputSpec to be an array");
    }
    const domAttrs = output[1] as Record<string, string> | undefined;
    expect(domAttrs?.["href"]).toBe("");
  });
});
