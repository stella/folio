/**
 * A story editor's snapshot of the legacy session's model is built over the
 * model: what no edit touched is the model's own, so the snapshot is handed
 * out read-only and a writer copies it first. Each test deep-freezes a model,
 * edits its stories, takes the snapshot, and checks that taking it wrote
 * nothing, that it reads as the snapshot of the same model unfrozen, and that
 * the copy a writer takes of it shares no object with the model.
 */

import { panic } from "better-result";
import {
  afterAll,
  beforeAll,
  describe,
  expect,
  setDefaultTimeout,
  test,
} from "bun:test";
import fc from "fast-check";
import type { EditorView } from "prosemirror-view";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { deepFreeze } from "../__tests__/deepFreeze";
import { DOCUMENT_SHAPES } from "../__tests__/documentShapes";
import { sharedObjects } from "../__tests__/sharedObjects";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import { parseDocx } from "../docx/parser";
import type { Document } from "../types/document";
import { readModel, type ReadonlyDocument } from "../types/documentModel";
import { createHeaderFooterEditorManager } from "./headerFooterEditorManager";
import { createNoteEditorManager } from "./noteEditorManager";

setDefaultTimeout(propertyTestTimeout(120_000));

let releaseDom: ReturnType<typeof acquireHarnessDom>;
beforeAll(() => {
  releaseDom = acquireHarnessDom();
});
afterAll(() => releaseDom());

const loadShapes = () =>
  Promise.all(
    DOCUMENT_SHAPES.map(async (shape) => ({
      id: shape.id,
      document: await parseDocx(
        (
          await prepareCanonicalDocxInput((await shape.build()).slice().buffer)
        ).unwrap(),
        { preloadFonts: false, detectVariables: false },
      ),
    })),
  );

/** Types `text` at the start of the story's first text block. */
const typeInto = (view: EditorView, text: string): void => {
  let position = -1;
  view.state.doc.descendants((node, pos) => {
    if (position !== -1) return false;
    if (node.isTextblock) {
      position = pos + 1;
      return false;
    }
    return true;
  });
  if (position === -1) panic("The story has no text block.");
  view.dispatch(view.state.tr.insertText(text, position));
};

type StoryKind = "notes" | "parts";

/**
 * The snapshot of `model` after typing `text` into each story `edits` picks,
 * and how many stories were typed into.
 */
const snapshotAfterEdits = (
  kind: StoryKind,
  model: ReadonlyDocument,
  edits: readonly boolean[],
  text: string,
): { snapshot: ReadonlyDocument; edited: number } => {
  const deps = {
    getHost: () => document.body,
    getDocument: () => readModel(model),
    readCanonicalDocument: () => null,
    getStyles: () => readModel(model).package.styles,
    getTheme: () => null,
    getExperimentalSession: () => "legacy" as const,
  };
  let edited = 0;
  const picked = (index: number) => edits.at(index % edits.length) ?? false;
  if (kind === "notes") {
    const manager = createNoteEditorManager(deps);
    for (const [index, story] of manager.listStories().entries()) {
      const view = manager.activate(story);
      if (!view || !picked(index)) continue;
      typeInto(view, text);
      edited += 1;
    }
    const snapshot = manager.snapshotDocument(model);
    manager.destroy();
    return { snapshot, edited };
  }
  const manager = createHeaderFooterEditorManager(deps);
  manager.sync();
  for (const [index, { rId }] of manager.listSlots().entries()) {
    const view = manager.getView(rId);
    if (!view || !picked(index)) continue;
    typeInto(view, text);
    edited += 1;
  }
  const snapshot = manager.snapshotDocument(model);
  manager.destroy();
  return { snapshot, edited };
};

const expectSnapshotContract = (
  kind: StoryKind,
  source: Document,
  edits: readonly boolean[],
  text: string,
): number => {
  const model = deepFreeze(cloneDocumentWithParagraphPropertySources(source));
  const frozen = snapshotAfterEdits(kind, model, edits, text);
  const plain = snapshotAfterEdits(
    kind,
    cloneDocumentWithParagraphPropertySources(source),
    edits,
    text,
  );
  expect(frozen.snapshot).toEqual(plain.snapshot);
  const owned = cloneDocumentWithParagraphPropertySources(frozen.snapshot);
  expect(sharedObjects(model, owned)).toEqual([]);
  expect(() => {
    owned.package.document.content.length = 0;
    owned.package.footnotes?.splice(0);
    owned.package.headers?.clear();
  }).not.toThrow();
  return frozen.edited;
};

describe("a story snapshot in the legacy session", () => {
  test.each([
    ["notes", "notes"],
    ["parts", "header-footer"],
  ] as const)(
    "of the %s reads a frozen model, and a writer's copy of it shares nothing with the model",
    async (kind, shapeId) => {
      const shapes = await loadShapes();
      const shape =
        shapes.find(({ id }) => id === shapeId) ??
        panic(`No document shape ${shapeId}`);
      expect(
        expectSnapshotContract(kind, shape.document, [true], "edited "),
      ).toBeGreaterThan(0);
    },
  );

  test("over every document shape and any stories edited", async () => {
    const shapes = await loadShapes();
    await assertProperty(
      fc.property(
        fc.constantFrom(...shapes),
        fc.constantFrom<StoryKind>("notes", "parts"),
        fc.array(fc.boolean(), { minLength: 1, maxLength: 4 }),
        fc.constantFrom("x", "edited ", "Note "),
        ({ document }, kind, edits, text) => {
          expectSnapshotContract(kind, document, edits, text);
        },
      ),
      { id: "story snapshots never write the model", numRuns: 40 },
    );
  });
});
