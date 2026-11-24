/**
 * A watermark set or cleared through the canonical session is one journaled
 * operation: it lands in the headers that need it, creates the ones a section
 * lacks, and undoes back to the document it started from.
 */
import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { assertExactModel } from "../../../../test/exactModel";
import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import { schema } from "../prosemirror/schema";
import type { Document, HeaderFooter } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { getDocumentWatermark } from "../watermark/index";
import { watermarkOperations } from "./canonicalWatermark";

const WATERMARK = { kind: "text", text: "CONFIDENTIAL" } as const;

/** A default header and a title-page section that lacks a first header. */
const titled = (): Document => {
  const document = createEmptyDocument();
  const header: HeaderFooter = {
    type: "header",
    hdrFtrType: "default",
    content: [{ type: "paragraph", paraId: "0000A001", content: [] }],
  };
  document.package.headers = new Map([["rIdHdr", header]]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    titlePg: true,
    headerReferences: [{ type: "default", rId: "rIdHdr" }],
  };
  return document;
};

const apply = (source: Document, watermark: typeof WATERMARK | undefined) => {
  const session = createCanonicalSession(source).unwrap();
  const before = cloneDocumentWithParagraphPropertySources(session.document);
  let state = EditorState.create({ schema, doc: session.projection.doc });
  const prepared = session.prepareOperations(
    state,
    watermarkOperations(
      cloneDocumentWithParagraphPropertySources(session.document),
      watermark,
    ),
  );
  if (prepared.isErr()) throw prepared.error;
  state = publishCanonicalProjection({
    session,
    state,
    commit: prepared.value,
  }).unwrap().state;
  return { session, state, before };
};

test("a watermark set through the session covers every header and undoes", () => {
  const { session, state, before } = apply(titled(), WATERMARK);
  const after = session.document;
  expect(getDocumentWatermark(after)).toEqual(WATERMARK);
  // The title page gets a first header of its own carrying the watermark.
  const references =
    after.package.document.finalSectionProperties?.headerReferences ?? [];
  expect(references.map(({ type }) => type).toSorted()).toEqual([
    "default",
    "first",
  ]);
  for (const { rId } of references)
    expect(after.package.headers?.get(rId)?.watermark).toEqual(WATERMARK);
  publishCanonicalProjection({
    session,
    state,
    commit: session.prepareUndo(state).unwrap(),
  }).unwrap();
  assertExactModel(session.document, before);
});

test("clearing the watermark through the session removes it from every header", () => {
  const { session } = apply(titled(), WATERMARK);
  const cleared = apply(
    cloneDocumentWithParagraphPropertySources(session.document),
    undefined,
  );
  expect(getDocumentWatermark(cleared.session.document)).toBeUndefined();
});
