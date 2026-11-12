/**
 * The host's copy of the document the editor shows is the committed document
 * with the live body, numbering and notes in place of its own, copied once.
 * Over every document shape, as committed document and as live stories, it
 * equals the copy put together from two copies (the committed document's, and
 * the live stories' grafted onto it), saves to the same bytes, writes neither
 * input, and shares no object with either.
 */

import { panic } from "better-result";
import {
  afterAll,
  beforeAll,
  describe,
  expect,
  setSystemTime,
  test,
} from "bun:test";

import { deepFreeze } from "@stll/folio-core/__tests__/deepFreeze";
import { DOCUMENT_SHAPES } from "@stll/folio-core/__tests__/documentShapes";
import { sharedObjects } from "@stll/folio-core/__tests__/sharedObjects";
import { prepareCanonicalDocxInput } from "@stll/folio-core/docx/canonicalSessionInput";
import { cloneDocumentWithParagraphPropertySources } from "@stll/folio-core/docx/document-clone";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { repackDocx } from "@stll/folio-core/docx/rezip";
import type { LayoutInstrumentation } from "@stll/folio-core/layout-engine/layoutInstrumentation";
import type { Document } from "@stll/folio-core/types/document";
import type { ReadonlyDocument } from "@stll/folio-core/types/documentModel";

import { copyHostDocument } from "./hostDocument";

const parseShape = async (build: () => Promise<Uint8Array>) =>
  parseDocx(
    (await prepareCanonicalDocxInput((await build()).slice().buffer)).unwrap(),
    { preloadFonts: false, detectVariables: false },
  );

type Shape = { id: string; committed: Document; live: Document };
let shapes: Shape[] = [];
const previousInstrumentation: LayoutInstrumentation | undefined =
  globalThis.__folioLayoutInstrumentation;
let copies = 0;

beforeAll(async () => {
  // Two parses per shape: the live stories are objects of their own.
  shapes = await Promise.all(
    DOCUMENT_SHAPES.map(async (shape) => ({
      id: shape.id,
      committed: deepFreeze(await parseShape(shape.build)),
      live: deepFreeze(await parseShape(shape.build)),
    })),
  );
  globalThis.__folioLayoutInstrumentation = {
    ...previousInstrumentation,
    onDocumentBuilt: (event) => {
      previousInstrumentation?.onDocumentBuilt?.(event);
      if (event.kind === "copy") copies += 1;
    },
  };
});
afterAll(() => {
  globalThis.__folioLayoutInstrumentation = previousInstrumentation;
});

/** The copy as two copies put it together: the committed one, then the stories'. */
const twoCopies = (
  committed: ReadonlyDocument,
  live: ReadonlyDocument,
): Document => {
  const doc = cloneDocumentWithParagraphPropertySources(committed);
  const stories = cloneDocumentWithParagraphPropertySources({
    package: {
      document: { content: live.package.document.content },
      ...(live.package.numbering && { numbering: live.package.numbering }),
      ...(live.package.footnotes && { footnotes: live.package.footnotes }),
      ...(live.package.endnotes && { endnotes: live.package.endnotes }),
    },
  });
  doc.package.document.content = stories.package.document.content;
  if (stories.package.numbering) {
    doc.package.numbering = stories.package.numbering;
  } else {
    delete doc.package.numbering;
  }
  if (stories.package.footnotes !== undefined) {
    doc.package.footnotes = stories.package.footnotes;
  }
  if (stories.package.endnotes !== undefined) {
    doc.package.endnotes = stories.package.endnotes;
  }
  return doc;
};

const shapeIds = DOCUMENT_SHAPES.map((shape) => shape.id);
const shapeById = (id: string) =>
  shapes.find((shape) => shape.id === id) ?? panic(`No shape ${id}`);

describe("the host's copy of the editor's document", () => {
  test.each(shapeIds)(
    "with %s committed, for the live stories of every shape, equals the two-copy result, takes one copy and shares nothing",
    (committedId) => {
      const { committed } = shapeById(committedId);
      for (const { live } of shapes) {
        copies = 0;
        const copied = copyHostDocument(committed, live);
        expect(copies).toBe(1);
        expect(copied).toStrictEqual(twoCopies(committed, live));
        expect(sharedObjects(committed, copied)).toEqual([]);
        expect(sharedObjects(live, copied)).toEqual([]);
        expect(() => {
          copied.package.document.content.length = 0;
          copied.package.footnotes?.splice(0);
          copied.package.headers?.clear();
        }).not.toThrow();
      }
    },
  );

  test.each(shapeIds)(
    "of %s saves to the bytes of the two-copy result",
    async (id) => {
      const { committed, live } = shapeById(id);
      const save = async (doc: Document) => {
        try {
          return { bytes: new Uint8Array(await repackDocx(doc)) };
        } catch (error) {
          return { refused: String(error) };
        }
      };
      // A save stamps the package's modified time: both read one clock.
      setSystemTime(new Date("2026-01-01T00:00:00Z"));
      try {
        const saved = await save(copyHostDocument(committed, live));
        expect("bytes" in saved).toBe(true);
        expect(saved).toEqual(await save(twoCopies(committed, live)));
      } finally {
        setSystemTime();
      }
    },
  );
});
