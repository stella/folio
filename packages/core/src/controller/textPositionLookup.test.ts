import { describe, expect, test } from "bun:test";

import { OP_STORIES, type TextPosition } from "@stll/docx-core/ops";

import { createCanonicalSession } from "../canonical/canonicalSession";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { findParagraphByParaId } from "../prosemirror/utils/findParagraphByParaId";
import { readModel } from "../types/documentModel";
import { legacyEditorPositionAt } from "./textPositionLookup";

/** A document from Markdown whose paragraphs carry ids, as a parsed DOCX's do. */
const withParagraphIds = (markdown: string) => {
  const document = fromMarkdown(markdown);
  for (const [index, block] of document.package.document.content.entries())
    if (block.type === "paragraph")
      block.paraId = (index + 1).toString(16).toUpperCase().padStart(8, "0");
  return document;
};

const openSession = (markdown: string) => {
  const created = createCanonicalSession(withParagraphIds(markdown));
  if (created.isErr()) throw created.error;
  return created.value;
};

const bodyParagraphIds = (session: ReturnType<typeof openSession>) =>
  readModel(session.document).package.document.content.flatMap((block) =>
    block.type === "paragraph" && block.paraId !== undefined
      ? [{ id: block.paraId }]
      : [],
  );

describe("legacyEditorPositionAt", () => {
  test("names the same text the canonical projection names at every offset", () => {
    const session = openSession(
      "Alpha **bold** and *slanted* text\n\nSecond paragraph here",
    );
    const legacy = toProseDoc(readModel(session.document));
    const projection = session.projection;
    const ids = bodyParagraphIds(session);
    expect(ids.length).toBe(2);
    for (const { id } of ids) {
      const paragraphText =
        findParagraphByParaId(legacy, id)?.node.textContent ?? "";
      expect(paragraphText.length).toBeGreaterThan(0);
      for (let offset = 0; offset <= paragraphText.length; offset += 1) {
        const position: TextPosition = {
          story: OP_STORIES.MAIN,
          blockId: id,
          offset,
        };
        const legacyAt = legacyEditorPositionAt(legacy, position);
        const canonicalAt = projection.positionAt(position);
        expect(canonicalAt.isOk()).toBe(true);
        if (legacyAt === null || canonicalAt.isErr())
          throw new Error(`No position for ${id}@${offset}`);
        const before = (doc: typeof legacy, at: number) =>
          doc.resolve(at).parent.textBetween(0, doc.resolve(at).parentOffset);
        expect(before(legacy, legacyAt)).toBe(paragraphText.slice(0, offset));
        expect(before(projection.doc, canonicalAt.value)).toBe(
          paragraphText.slice(0, offset),
        );
      }
    }
  });

  test("refuses an unknown paragraph, an offset past the end and another story", () => {
    const session = openSession("Only paragraph");
    const legacy = toProseDoc(readModel(session.document));
    const [first] = bodyParagraphIds(session);
    if (first === undefined) throw new Error("No paragraph id");
    const at = (position: TextPosition) =>
      legacyEditorPositionAt(legacy, position);
    expect(
      at({ story: OP_STORIES.MAIN, blockId: "NOSUCHID", offset: 0 }),
    ).toBeNull();
    expect(
      at({ story: OP_STORIES.MAIN, blockId: first.id, offset: 99 }),
    ).toBeNull();
    expect(
      at({ story: OP_STORIES.MAIN, blockId: first.id, offset: -1 }),
    ).toBeNull();
    expect(
      at({
        story: { kind: "footnote", id: 1 },
        blockId: first.id,
        offset: 0,
      }),
    ).toBeNull();
  });
});
