import {
  afterAll,
  beforeAll,
  describe,
  expect,
  setSystemTime,
  test,
} from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { EditorState } from "prosemirror-state";

import {
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  storyParagraphs,
} from "@stll/docx-core/ops";

import { serializeCanonicalSave } from "../docx/canonicalSave";
import {
  CANONICAL_DOCX_PARSE_OPTIONS,
  openCanonicalDocx,
  openOwnedCanonicalDocx,
} from "../docx/canonicalSessionInput";
import { getDocumentSourceBaseline } from "../docx/headerFooterVerbatim";
import type { Document } from "../types/document";
import { readModel } from "../types/documentModel";
import {
  type CanonicalSession,
  createCanonicalSession,
  createCanonicalSessionTaking,
  publishCanonicalProjection,
} from "./canonicalSession";

// Compare saved bytes with the writer's modification timestamp held fixed.
beforeAll(() => setSystemTime(new Date("2026-01-02T03:04:05Z")));
afterAll(() => setSystemTime());

const CORPUS = path.join(
  import.meta.dir,
  "../docx/__tests__/__fixtures__/corpus",
);
const FIXTURES = readdirSync(CORPUS)
  .filter((name) => name.endsWith(".docx"))
  .sort();

const read = (name: string) =>
  new Uint8Array(readFileSync(path.join(CORPUS, name)));

/** Insert text at the start of the body's first paragraph. */
const edit = (session: CanonicalSession): boolean => {
  const first = storyParagraphs(
    readModel(session.document).package.document,
  ).at(0)?.paragraph.paraId;
  if (first === undefined) return false;
  const position = {
    story: OP_STORIES.MAIN,
    blockId: first,
    offset: 0,
  } as const;
  const state = EditorState.create({ doc: session.projection.doc });
  const prepared = session.prepareOps(
    state,
    [
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: position,
        text: "Edited ",
        runProps: INHERIT_RUN_PROPS,
      },
    ],
    { type: "text", anchor: position, head: position },
  );
  if (prepared.isErr()) return false;
  return publishCanonicalProjection({
    session,
    state,
    commit: prepared.value,
  }).isOk();
};

const saved = async (session: CanonicalSession) => {
  const result = await serializeCanonicalSave({
    snapshot: session.captureSaveSnapshot(),
  });
  return { path: result.path, bytes: new Uint8Array(result.buffer) };
};

describe("an editor's own open compares its saves against the parsed body itself", () => {
  test("the owned open keeps the parsed body as the baseline, the public open a copy", async () => {
    const [name] = FIXTURES;
    if (name === undefined) throw new Error("the corpus has no fixture");
    const owned = (
      await openOwnedCanonicalDocx(read(name), CANONICAL_DOCX_PARSE_OPTIONS)
    ).unwrap().document;
    const copied = (
      await openCanonicalDocx(read(name), CANONICAL_DOCX_PARSE_OPTIONS)
    ).unwrap().document;
    const baselineOf = (document: Document) => {
      const baseline = getDocumentSourceBaseline(document);
      if (baseline.type !== "captured") throw new Error("no source baseline");
      return baseline.body;
    };
    expect(baselineOf(owned)).toBe(owned.package.document);
    expect(baselineOf(copied)).not.toBe(copied.package.document);
  });

  test.each(FIXTURES)(
    "%s saves the same bytes, unedited and edited, either way it was opened",
    async (name) => {
      const owned = await openOwnedCanonicalDocx(
        read(name),
        CANONICAL_DOCX_PARSE_OPTIONS,
      );
      const copied = await openCanonicalDocx(
        read(name),
        CANONICAL_DOCX_PARSE_OPTIONS,
      );
      expect(owned.isOk()).toBe(copied.isOk());
      if (owned.isErr() || copied.isErr()) return;
      const taking = createCanonicalSessionTaking(owned.value.document);
      const copying = createCanonicalSession(copied.value.document);
      expect(taking.isOk()).toBe(copying.isOk());
      if (taking.isErr() || copying.isErr()) return;

      expect(await saved(taking.value)).toEqual(await saved(copying.value));
      expect(edit(taking.value)).toBe(edit(copying.value));
      expect(await saved(taking.value)).toEqual(await saved(copying.value));
    },
  );
});
