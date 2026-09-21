/**
 * No anchor the hyperlink mark renders may target a frame outside the
 * allow-list, whatever a document, a paste or a collaborator supplies.
 *
 * A `w:tgtFrame` is kept verbatim in the model so a save writes back what the
 * source said. That makes the clamp this side's job: a named frame addresses
 * another browsing context, and `_blank` without `rel="noopener noreferrer"`
 * hands the opened context a handle on the opener.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../../../test/property-testing";

import { HyperlinkExtension } from "./HyperlinkExtension";

const ALLOWED_TARGETS = ["_blank", "_self", "_parent", "_top"];

/** `getAttrs` reads attributes and nothing else, so this stands in for a pasted anchor. */
const pastedAnchor = (attrs: Record<string, string>): HTMLElement =>
  ({ getAttribute: (name: string) => attrs[name] ?? null }) as unknown as HTMLElement;

const markSpec = HyperlinkExtension().config.markSpec;

const renderedAnchorAttrs = (attrs: Record<string, unknown>): Record<string, string> => {
  const toDOM = markSpec.toDOM;
  if (!toDOM) {
    throw new Error("HyperlinkExtension must define toDOM");
  }
  // The reader checks the mark type before reading attrs, so the stand-in
  // names it; nothing else about a real Mark is consulted.
  const rendered = toDOM({ type: { name: "hyperlink" }, attrs } as never, true);
  const [tag, domAttrs] = rendered as [string, Record<string, string>];
  expect(tag).toBe("a");
  return domAttrs;
};

const frameArbitrary = fc.oneof(
  fc.constantFrom(...ALLOWED_TARGETS),
  fc.constantFrom("_top ", "topFrame", "__proto__", "javascript:alert(1)", "", "MainWindow"),
  fc.string({ maxLength: 24 }),
);

describe("hyperlink anchors (property)", () => {
  test(
    "keep allow-listed targets and clamp every other target to a safe context",
    () => {
      fc.assert(
        fc.property(frameArbitrary, fc.webUrl(), (frame, href) => {
          // Authored: a frame the parser kept verbatim off `w:tgtFrame`.
          const authored = renderedAnchorAttrs({ href, tooltip: null, target: frame });
          // Pasted: an anchor carrying its own `target` through `parseDOM`.
          const pastedAttrs = markSpec.parseDOM?.[0]?.getAttrs?.(
            pastedAnchor({ href, target: frame }),
          );
          const pasted = renderedAnchorAttrs({
            tooltip: null,
            ...(pastedAttrs === false || pastedAttrs === null || pastedAttrs === undefined
              ? { href }
              : pastedAttrs),
          });
          // Collaboration-style: an attr map merged onto an existing mark.
          const collaborated = renderedAnchorAttrs({
            href,
            tooltip: null,
            target: frame,
            rel: "",
          });
          const expectedTarget = ALLOWED_TARGETS.includes(frame) ? frame : "_blank";

          for (const anchor of [authored, pasted, collaborated]) {
            expect(anchor["target"]).toBe(expectedTarget);
            expect(anchor["rel"]).toBe("noopener noreferrer");
          }
        }),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(30_000),
  );
});
