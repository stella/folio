import { panic } from "better-result";
import { expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import {
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  storyParagraphs,
} from "@stll/docx-core/ops";

import { deepFreeze } from "../__tests__/deepFreeze";
import {
  CANONICAL_DOCX_PARSE_OPTIONS,
  openCanonicalDocx,
  openOwnedCanonicalDocx,
  takeOwnedDocument,
} from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type { LayoutInstrumentation } from "../layout-engine/layoutInstrumentation";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { readModel } from "../types/documentModel";
import {
  createCanonicalSession,
  createCanonicalSessionTaking,
  publishCanonicalProjection,
} from "./canonicalSession";

const SOURCE =
  "First clause.\n\nSecond clause.\n\n| A | B |\n|---|---|\n| 1 | 2 |";

const opened = async (owned: boolean) => {
  const bytes = await createDocx(fromMarkdown(SOURCE));
  const result = await (owned ? openOwnedCanonicalDocx : openCanonicalDocx)(
    bytes,
    CANONICAL_DOCX_PARSE_OPTIONS,
  );
  return result.unwrap().document;
};

/** The whole-document copies made while `run` runs. */
const copiesDuring = <T>(run: () => T): { value: T; copies: number } => {
  const previous: LayoutInstrumentation | undefined =
    globalThis.__folioLayoutInstrumentation;
  let copies = 0;
  globalThis.__folioLayoutInstrumentation = {
    ...previous,
    onDocumentBuilt: (event) => {
      previous?.onDocumentBuilt?.(event);
      if (event.kind === "copy") copies += 1;
    },
  };
  try {
    return { value: run(), copies };
  } finally {
    globalThis.__folioLayoutInstrumentation = previous;
  }
};

test("a session takes a document handed to it without a copy, and the public one copies", async () => {
  const handed = await opened(false);
  const taking = copiesDuring(() => createCanonicalSessionTaking(handed));
  expect(taking.value.isOk()).toBe(true);
  expect(taking.copies).toBe(0);

  const copying = copiesDuring(() => createCanonicalSession(handed));
  expect(copying.value.isOk()).toBe(true);
  expect(copying.copies).toBe(1);
});

test("a session never writes the document it takes, through its seed and its edits", async () => {
  const handed = deepFreeze(await opened(false));
  const session = createCanonicalSessionTaking(handed).unwrap();
  const first =
    storyParagraphs(readModel(session.document).package.document).at(0)
      ?.paragraph.paraId ?? panic("The source needs a first paragraph.");
  const position = {
    story: OP_STORIES.MAIN,
    blockId: first,
    offset: 0,
  } as const;
  const state = EditorState.create({ doc: session.projection.doc });
  const prepared = session
    .prepareOps(
      state,
      [
        {
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          story: OP_STORIES.MAIN,
          blockId: first,
          patch: { alignment: "center" },
        },
        {
          type: DOCUMENT_OP_TYPES.INSERT_TEXT,
          at: position,
          text: "Edited ",
          runProps: INHERIT_RUN_PROPS,
        },
      ],
      { type: "text", anchor: position, head: position },
    )
    .unwrap();
  publishCanonicalProjection({ session, state, commit: prepared }).unwrap();

  const edited = storyParagraphs(
    readModel(session.document).package.document,
  ).at(0)?.paragraph;
  expect(edited?.formatting?.alignment).toBe("center");
  // The handed document is still the one read.
  expect(
    storyParagraphs(handed.package.document).at(0)?.paragraph.formatting
      ?.alignment,
  ).toBeUndefined();
});

test("a host that transfers the package's bytes away leaves the session's own", async () => {
  const handed = await opened(false);
  const session = createCanonicalSessionTaking(handed).unwrap();
  const bytes = handed.originalBuffer ?? panic("The package keeps its bytes.");
  const length = bytes.byteLength;
  structuredClone(bytes, { transfer: [bytes] });
  expect(bytes.byteLength).toBe(0);
  const kept = readModel(session.document).originalBuffer;
  expect(kept?.byteLength).toBe(length);
  expect(() => structuredClone(readModel(session.document))).not.toThrow();
});

test("an editor's own open hands its document to the first session only", async () => {
  const owned = await opened(true);
  expect(takeOwnedDocument(owned)).toBe(true);
  expect(takeOwnedDocument(owned)).toBe(false);
  expect(takeOwnedDocument(await opened(false))).toBe(false);
});
