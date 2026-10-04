import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import type { Document } from "../../model/document";
import { compileEditorIntent } from "../editorIntent";
import { applyDocumentOps } from "../apply";
import { OP_STORIES } from "../types";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";

setDefaultTimeout(propertyTestTimeout(30_000));
test("TOC identities are package-wide, deterministic and exact under inverse", () => {
  assertProperty(
    fc.property(
      fc.constantFrom(0, 1, 3),
      fc.constantFrom(0, 1, 2, 3, 4, 5, 6, 7, 8),
      fc.boolean(),
      (offset, level, existing) => {
        const document = {
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "12345678",
                  content: [
                    ...(existing
                      ? [{ type: "bookmarkStart", id: 7, name: "_TocExisting" } as const]
                      : []),
                    { type: "run", content: [{ type: "text", text: "Head" }] },
                    ...(existing ? [{ type: "bookmarkEnd", id: 7 } as const] : []),
                  ],
                },
                {
                  type: "paragraph",
                  paraId: "23456789",
                  content: [{ type: "run", content: [{ type: "text", text: "a😀" }] }],
                },
              ],
            },
            footnotes: [
              {
                type: "footnote",
                id: 1,
                content: [
                  {
                    type: "paragraph",
                    paraId: "00000001",
                    content: [
                      { type: "bookmarkStart", id: 1, name: "_Toc1" },
                      { type: "run", content: [{ type: "text", text: "Note" }] },
                      { type: "bookmarkEnd", id: 1 },
                    ],
                  },
                ],
              },
            ],
          },
        } satisfies Document;
        const intent = {
          type: "generateTOC",
          at: { story: OP_STORIES.MAIN, blockId: "23456789", offset },
          title: "Contents",
          headings: [{ blockId: "12345678", text: "Head", level }],
          tabPosition: 9360,
        } as const;
        const compiled = compileEditorIntent(document, {
          intent,
          mode: { type: "editing" },
        }).unwrap();
        expect(
          compileEditorIntent(document, { intent, mode: { type: "editing" } }).unwrap(),
        ).toEqual(compiled);
        const applied = applyDocumentOps(document, compiled.ops).unwrap();
        const ids = applied.document.package.document.content.map((item) =>
          item.type === "paragraph" ? item.paraId : panic("Unexpected TOC block"),
        );
        expect(new Set(ids).size).toBe(ids.length);
        expect(ids).not.toContain("00000001");
        const heading = applied.document.package.document.content.find(
          (item) => item.type === "paragraph" && item.paraId === "12345678",
        );
        if (heading?.type !== "paragraph") return panic("Missing TOC heading");
        const bookmarks = heading.content.filter((item) => item.type === "bookmarkStart");
        expect(bookmarks).toHaveLength(1);
        if (existing)
          expect(bookmarks.at(0)).toEqual({ type: "bookmarkStart", id: 7, name: "_TocExisting" });
        else {
          expect(bookmarks.at(0)?.id).not.toBe(1);
          expect(bookmarks.at(0)?.name).not.toBe("_Toc1");
        }
        const undone = applyDocumentOps(applied.document, applied.inverse).unwrap();
        expect(undone.document).toStrictEqual(document);
        expect(applyDocumentOps(undone.document, undone.inverse).unwrap().document).toStrictEqual(
          applied.document,
        );
        const suggested = compileEditorIntent(document, {
          intent,
          mode: {
            type: "suggesting",
            revision: { id: 20, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
            newIds: {},
          },
        });
        expect(suggested.isErr()).toBe(true);
      },
    ),
    { numRuns: 24 },
  );
});
