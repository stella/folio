import { describe, expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphFormatting,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { cloneModel } from "../modelClone";
import { registerParagraphPropertySnapshot } from "../paragraphPropertySnapshot";
import { paragraphPropertyChange } from "../review";
import { DOCUMENT_OP_TYPES, INHERIT_RUN_PROPS, OP_STORIES } from "../types";

const info = { id: 1, author: "Reviewer", date: "2026-05-06T07:08:09Z" };
const paragraphOf = (formatting: ParagraphFormatting): Paragraph => ({
  type: "paragraph",
  formatting,
  content: [],
});

describe("authored paragraph property snapshots", () => {
  test.each([
    { patch: { keepNext: true }, snapshot: { keepNext: true } },
    {
      patch: { keepNext: true, indentLeft: 720 },
      snapshot: { keepNext: true, indentLeft: 720 },
    },
  ])(
    "a direct property edit retains its authored baseline for a later suggestion: %j",
    ({ patch, snapshot }) => {
      const paragraph = paragraphOf({ indentLeft: 720 });
      paragraph.paraId = "00000001";
      registerParagraphPropertySnapshot({ paragraph, snapshot: undefined });
      const document: Document = {
        package: { document: { content: [paragraph] } },
      };
      const applied = applyDocumentOps(document, [
        {
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          story: OP_STORIES.MAIN,
          blockId: "00000001",
          patch,
        },
        {
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          story: OP_STORIES.MAIN,
          blockId: "00000001",
          patch: { indentLeft: 1440 },
          revision: info,
        },
      ]).unwrap();
      const edited = applied.document.package.document.content.at(0);
      expect(edited?.type).toBe("paragraph");
      if (edited?.type !== "paragraph") return;
      expect(edited.propertyChanges?.at(0)?.previousFormatting).toEqual(
        snapshot,
      );
    },
  );

  test("a content edit before a property suggestion retains its authored baseline", () => {
    const paragraph = paragraphOf({ indentLeft: 720, keepNext: true });
    paragraph.paraId = "00000001";
    registerParagraphPropertySnapshot({
      paragraph,
      snapshot: { keepNext: true },
    });
    const document: Document = {
      package: { document: { content: [paragraph] } },
    };
    const applied = applyDocumentOps(document, [
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
        text: "Content",
        runProps: INHERIT_RUN_PROPS,
      },
      {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        patch: { indentLeft: 1440 },
        revision: info,
      },
    ]).unwrap();
    const edited = applied.document.package.document.content.at(0);
    expect(edited?.type).toBe("paragraph");
    if (edited?.type !== "paragraph") return;
    expect(edited.propertyChanges?.at(0)?.previousFormatting).toEqual({
      keepNext: true,
    });
  });

  test("keeps paragraph ownership when effective formatting is shared", () => {
    const formatting = { indentLeft: 720, keepNext: true };
    const inherited = paragraphOf(formatting);
    const direct = paragraphOf(formatting);
    registerParagraphPropertySnapshot({
      paragraph: inherited,
      snapshot: { keepNext: true },
    });
    registerParagraphPropertySnapshot({
      paragraph: direct,
      snapshot: formatting,
    });
    expect(Object.getOwnPropertySymbols(inherited)).toEqual([]);
    expect(inherited).toStrictEqual(paragraphOf(formatting));
    expect(structuredClone(inherited)).toEqual(inherited);
    const model = { paragraphs: [inherited, direct] };
    const cloned = cloneModel(model);
    expect(cloned.paragraphs.at(0)?.formatting).not.toBe(formatting);
    for (const value of [model, cloned]) {
      expect(
        value.paragraphs.map((paragraph) =>
          paragraphPropertyChange(info, paragraph),
        ),
      ).toEqual([
        {
          type: "paragraphPropertyChange",
          info,
          previousFormatting: { keepNext: true },
        },
        {
          type: "paragraphPropertyChange",
          info,
          previousFormatting: formatting,
        },
      ]);
    }
    expect(
      paragraphPropertyChange(info, structuredClone(inherited))
        .previousFormatting,
    ).toEqual(formatting);
  });

  test("retains cloned bindings and reads changed formatting directly", () => {
    const paragraph = paragraphOf({ indentLeft: 720 });
    registerParagraphPropertySnapshot({ paragraph, snapshot: undefined });
    for (const value of [paragraph, cloneModel(paragraph)]) {
      expect(paragraphPropertyChange(info, value)).not.toHaveProperty(
        "previousFormatting",
      );
    }
    const changed = cloneModel(paragraph);
    changed.formatting = { indentLeft: 0, hangingIndent: false };
    expect(paragraphPropertyChange(info, changed).previousFormatting).toEqual(
      changed.formatting,
    );
  });

  test("records unbound direct formatting", () => {
    const formatting = {
      indentLeft: 0,
      indentFirstLine: 120,
      hangingIndent: false,
    };
    expect(
      paragraphPropertyChange(info, paragraphOf(formatting)).previousFormatting,
    ).toEqual(formatting);
  });
});
