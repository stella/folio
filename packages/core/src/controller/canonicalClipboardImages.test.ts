import { expect, test } from "bun:test";
import { Fragment, Slice, type Node as PMNode } from "prosemirror-model";
import JSZip from "jszip";
import { panic } from "better-result";

import { createCanonicalEditorHarness } from "../../../../test/canonicalEditorHarness";
import { shapeArrayBuffer } from "../__tests__/documentShapes";
import { modelMarkdown, parseShapeDocument } from "../__tests__/editorHarness";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseRelationships, resolveRelativePath } from "../docx/relsParser";
import { toMarkdownResult } from "../markdown";
import { visitParagraphRuns } from "../docx/paragraphTraversal";
import type { Document, DrawingContent } from "../types/document";

const drawings = (document: Document) => {
  const result: DrawingContent[] = [];
  for (const paragraph of document.package.document.content) {
    if (paragraph.type !== "paragraph") continue;
    visitParagraphRuns(paragraph, (run) => {
      for (const item of run.content) if (item.type === "drawing") result.push(item);
    });
  }
  return result;
};

const CASES = (["editing", "suggesting"] as const).flatMap((mode) =>
  (["package", "dataUrl", "http"] as const).map((resource) => ({ mode, resource })),
);

test.each(CASES)(
  "canonical copied $resource image retains owning relationship through save in $mode",
  async ({ mode, resource }) => {
    const source = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer("image")));
    const driver = createCanonicalEditorHarness(source, mode);
    try {
      const before = driver.snapshot();
      let slice = new Slice(driver.state.doc.content, 0, 0);
      if (resource !== "package") {
        const imageNodes: PMNode[] = [];
        driver.state.doc.descendants((node) => {
          if (node.type.name === "image") imageNodes.push(node);
        });
        const imageNode = imageNodes.at(0);
        if (imageNode === undefined) throw new TypeError("Image fixture has no projected picture.");
        const imageType = driver.state.schema.nodes["image"];
        if (imageType === undefined) throw new TypeError("Image fixture has no image schema.");
        slice = new Slice(
          Fragment.from(
            imageType.create({
              ...imageNode.attrs,
              rId: null,
              src:
                resource === "http"
                  ? "https://example.com/copied.png"
                  : drawings(source).at(0)?.image.src,
            }),
          ),
          0,
          0,
        );
      }
      const end = driver.state.doc.content.size - 1;
      driver.history.setSelection(end, end);
      driver.paste(slice);
      expect(driver.refusals).toEqual([]);
      const edited = driver.snapshot();
      expect(drawings(edited)).toHaveLength(2);
      if (resource !== "http") expect(modelMarkdown(edited)).not.toContain("data:image/");
      const snapshot =
        driver.history.captureCanonicalSave() ??
        panic("Canonical image save snapshot unavailable.");
      const { buffer: bytes } = await serializeCanonicalSave({
        snapshot,
        options: { mode: "full" },
      });
      const zip = await JSZip.loadAsync(bytes);
      const xml = await zip.file("word/document.xml")?.async("string");
      const rels = await zip.file("word/_rels/document.xml.rels")?.async("string");
      if (rels === undefined) throw new TypeError("Saved image fixture has no relationship part.");
      const savedRelationships = parseRelationships(rels);
      await Promise.all(
        drawings(edited).map(async ({ image }) => {
          const id = image.rId;
          if (id === undefined) throw new TypeError("Copied image has no allocated relationship.");
          const relationship = edited.package.relationships?.get(id);
          expect(savedRelationships.get(id)).toEqual(relationship);
          if (relationship === undefined || relationship.targetMode === "External") return;
          const path = resolveRelativePath("word/document.xml", relationship.target);
          const media = edited.package.media?.get(path);
          if (media === undefined) throw new TypeError("Copied image lacks canonical media bytes.");
          expect(await zip.file(path)?.async("uint8array")).toEqual(new Uint8Array(media.data));
        }),
      );
      if (resource !== "http") {
        expect(toMarkdownResult(edited).images.size).toBe(2);
      }
      const reopened = await parseShapeDocument(new Uint8Array(bytes));
      expect({ count: drawings(reopened).length, xml, rels }).toMatchObject({ count: 2 });
      if (resource === "http") {
        const copied = drawings(reopened).at(-1)?.image;
        expect(copied?.src).toBeUndefined();
        // The parser's sanitizer keeps network pictures unpainted; their relationship remains a fact.
        expect(
          copied?.rId === undefined ? undefined : reopened.package.relationships?.get(copied.rId),
        ).toMatchObject({ targetMode: "External", target: "https://example.com/copied.png" });
      } else {
        expect(drawings(reopened).map(({ image }) => image.src)).toEqual(
          drawings(edited).map(({ image }) => image.src),
        );
      }
      expect(driver.history.undo()).toBe(true);
      expect(driver.snapshot()).toEqual(before);
      expect(driver.history.redo()).toBe(true);
      expect(driver.snapshot()).toEqual(edited);
    } finally {
      driver.dispose();
    }
  },
);
