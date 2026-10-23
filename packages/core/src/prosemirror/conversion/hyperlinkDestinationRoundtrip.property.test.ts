import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document, Hyperlink } from "../../types/document";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const hyperlinks = (document: Document): Hyperlink[] => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Expected a paragraph");
  return paragraph.content.filter((item) => item.type === "hyperlink");
};

describe("hyperlink destinations survive the editor", () => {
  test(
    "keeps each external document's bookmark alongside its URL",
    () => {
      fc.assert(
        fc.property(
          fc.webUrl(),
          fc.nat({ max: 1_000_000 }),
          fc.integer({ min: 1, max: 5 }),
          (href, firstBookmark, count) => {
            const links = Array.from({ length: count }, (_, index) => ({
              type: "hyperlink",
              href,
              anchor: `_Toc${firstBookmark + index}`,
              children: [
                { type: "run", content: [{ type: "text", text: "Section" }] },
              ],
            })) satisfies Hyperlink[];
            const source = {
              package: {
                document: {
                  content: [{ type: "paragraph", content: links }],
                },
              },
            } satisfies Document;
            const saved = fromProseDoc(toProseDoc(source), source, {
              stylesheetSource: { type: "package" },
            });
            expect(hyperlinks(saved)).toEqual(links);
            expect(
              hyperlinks(
                fromProseDoc(toProseDoc(saved), saved, {
                  stylesheetSource: { type: "package" },
                }),
              ),
            ).toEqual(links);
          },
        ),
        propertyConfig({ numRuns: 50 }),
      );
    },
    propertyTestTimeout(30_000),
  );

  test("distinct bookmarks do not merge when their external URL matches", () => {
    const href = "https://example.test/document.docx";
    const first = schema.mark("hyperlink", { href, anchor: "First" });
    const second = schema.mark("hyperlink", { href, anchor: "Second" });
    const document = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.text("First", [first]),
        schema.text("Second", [second]),
      ]),
    ]);
    const saved = fromProseDoc(document, undefined, {
      stylesheetSource: { type: "package" },
    });

    expect(hyperlinks(saved).map(({ anchor }) => anchor)).toEqual([
      "First",
      "Second",
    ]);
  });
});
