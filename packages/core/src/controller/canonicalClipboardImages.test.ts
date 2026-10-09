import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { expect, setDefaultTimeout, test } from "bun:test";
import { DOMParser, DOMSerializer, Fragment, Slice, type Node as PMNode } from "prosemirror-model";
import JSZip from "jszip";
import { panic } from "better-result";

import { createCanonicalEditorHarness } from "../../../../test/canonicalEditorHarness";
import { shapeArrayBuffer } from "../__tests__/documentShapes";
import { modelMarkdown, parseShapeDocument } from "../__tests__/editorHarness";
import { createDocx } from "../docx/rezip";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { parseRelationships, resolveRelativePath } from "../docx/relsParser";
import { toMarkdownResult } from "../markdown";
import { visitParagraphRuns } from "../docx/paragraphTraversal";
import type { Document, DrawingContent, NonVisualDrawingNames } from "../types/document";

setDefaultTimeout(propertyTestTimeout(30_000));

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

type CopiedImageRoundtripOptions = (typeof CASES)[number] & {
  names?: NonVisualDrawingNames;
  clipboard: "pm" | "html";
};

const copiedImageRoundtrip = async ({
  mode,
  resource,
  names,
  clipboard,
}: CopiedImageRoundtripOptions) => {
  let source = await parseShapeDocument(new Uint8Array(await shapeArrayBuffer("image")));
  if (names !== undefined) {
    const image = drawings(source).at(0)?.image ?? panic("The generated picture has no source.");
    image.pictureNames = { ...names };
    image.docPrName = "Drawing name";
    image.alt = "Drawing description";
    image.title = "Drawing title";
    source = await parseShapeDocument(new Uint8Array(await createDocx(source)));
    expect(drawings(source).at(0)?.image.pictureNames).toEqual(names);
  }
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
    if (clipboard === "html") {
      const element = document.createElement("div");
      element.append(
        DOMSerializer.fromSchema(driver.state.schema).serializeFragment(slice.content),
      );
      const parsed = DOMParser.fromSchema(driver.state.schema).parseSlice(element);
      slice = new Slice(parsed.content, 0, 0);
    }
    const end = driver.state.doc.content.size - 1;
    driver.history.setSelection(end, end);
    driver.paste(slice);
    expect(driver.refusals).toEqual([]);
    const edited = driver.snapshot();
    expect(drawings(edited)).toHaveLength(2);
    if (names !== undefined) {
      for (const { image } of drawings(edited)) {
        expect(image.pictureNames).toEqual(names);
        expect({ name: image.docPrName, alt: image.alt, title: image.title }).toEqual({
          name: "Drawing name",
          alt: "Drawing description",
          title: "Drawing title",
        });
      }
    }
    if (resource !== "http") expect(modelMarkdown(edited)).not.toContain("data:image/");
    const snapshot =
      driver.history.captureCanonicalSave() ?? panic("Canonical image save snapshot unavailable.");
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
    if (names !== undefined)
      expect(drawings(reopened).map(({ image }) => image.pictureNames)).toEqual([names, names]);
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
};

test.each(CASES)(
  "canonical copied $resource image retains owning relationship through save in $mode",
  (options) => copiedImageRoundtrip({ ...options, clipboard: "pm" }),
);

test("generated clipboard image metadata preserves picture and drawing names independently", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.record({
        name: fc.constantFrom("dot.png", "図😀.png", 'scan<&".png'),
        alt: fc.option(fc.constantFrom("", "Description<&>", "図の説明"), { nil: undefined }),
        title: fc.option(fc.constantFrom("", "Title<&>", "図の題名"), { nil: undefined }),
      }),
      async ({ name, alt, title }) => {
        const names = {
          name,
          ...(alt === undefined ? {} : { alt }),
          ...(title === undefined ? {} : { title }),
        };
        for (const options of CASES)
          await copiedImageRoundtrip({ ...options, names, clipboard: "html" });
      },
    ),
    { numRuns: 6 },
  );
});
