/**
 * Before the canonical session exists, the view only paints the loaded model.
 * The API returns that model even if the view holds different text. Writers
 * copy the model before changing it.
 */

import { panic } from "better-result";
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

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
import { singletonManager } from "../prosemirror/schema";
import type { Document } from "../types/document";
import { readModel, type ReadonlyDocument } from "../types/documentModel";
import { createEmptyDocument } from "../utils/createDocument";
import { createHiddenEditorApi } from "./hiddenEditorApi";
import { createHiddenEditorState } from "./hiddenEditorManager";

setDefaultTimeout(propertyTestTimeout(120_000));

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

/** `state` with `text` typed at the start of its `index`th text block. */
const typed = (state: EditorState, index: number, text: string) => {
  const starts: number[] = [];
  state.doc.descendants((node, pos) => {
    if (node.isTextblock) {
      starts.push(pos + 1);
      return false;
    }
    return true;
  });
  if (starts.length === 0) return state;
  const position = starts.at(index % starts.length) ?? panic("No text block");
  return state.apply(state.tr.insertText(text, position));
};

/** The hidden editor API over `model`, its view holding `state`. */
const apiOver = (model: ReadonlyDocument, state: EditorState) => {
  const view = { state };
  return createHiddenEditorApi({
    // SAFETY: This fixture supplies only the view state to distinguish it from the loaded model.
    // eslint-disable-next-line typescript/no-unsafe-type-assertion
    getView: () => view as EditorView,
    getDocumentContext: () => model,
    isDestroying: () => false,
    ensureView: () => {},
    isViewRequested: () => true,
  });
};

/** Read the loaded model with a view that may hold different text. */
const readLoadedDocument = (
  model: ReadonlyDocument,
  edit: { index: number; text: string } | null,
): ReadonlyDocument => {
  const state = createHiddenEditorState({
    document: model,
    styles: readModel(model).package.styles ?? null,
    manager: singletonManager,
  });
  const api = apiOver(
    model,
    edit ? typed(state, edit.index, edit.text) : state,
  );
  return api.getDocument() ?? panic("Expected the editor's document");
};

const expectDocumentContract = (
  source: Document,
  edit: { index: number; text: string } | null,
): void => {
  const model = deepFreeze(cloneDocumentWithParagraphPropertySources(source));
  const loaded = readLoadedDocument(model, edit);
  expect(loaded).toBe(model);
  const owned = cloneDocumentWithParagraphPropertySources(loaded);
  expect(sharedObjects(model, owned)).toEqual([]);
  expect(() => {
    owned.package.document.content.length = 0;
    if (owned.package.document.finalSectionProperties)
      owned.package.document.finalSectionProperties.marginLeft = 4321;
    owned.package.styles?.styles.splice(0);
    owned.package.footnotes?.splice(0);
    owned.package.headers?.clear();
  }).not.toThrow();
};

describe("the hidden editor's document before a canonical session", () => {
  test("returns the loaded model across document shapes and view edits", async () => {
    const shapes = await loadShapes();
    await assertProperty(
      fc.property(
        fc.constantFrom(...shapes),
        fc.option(
          fc.record({
            index: fc.nat({ max: 8 }),
            text: fc.constantFrom("x", "edited ", "Note "),
          }),
        ),
        ({ document }, edit) => {
          expectDocumentContract(document, edit);
        },
      ),
      {
        id: "the hidden editor returns the loaded model before activation",
        numRuns: 40,
      },
    );
  });
});

describe("the hidden editor's document in a canonical session", () => {
  test("is the session's own document, read without a copy", () => {
    const session = deepFreeze(
      cloneDocumentWithParagraphPropertySources(
        createEmptyDocument({ initialText: "Session" }),
      ),
    );
    const state = createHiddenEditorState({
      document: session,
      styles: readModel(session).package.styles ?? null,
      manager: singletonManager,
    });
    const api = createHiddenEditorApi({
      // SAFETY: This fixture supplies only the view state to distinguish it from the loaded model.
      // eslint-disable-next-line typescript/no-unsafe-type-assertion
      getView: () => ({ state }) as EditorView,
      getDocumentContext: () => session,
      getCanonicalDocument: () => session,
      isDestroying: () => false,
      ensureView: () => {},
      isViewRequested: () => true,
    });
    expect(api.getDocument()).toBe(session);
  });
});
