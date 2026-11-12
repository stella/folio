/**
 * The hidden editor's `getDocument` builds the default session's document
 * from the editor state over the model: the parts no edit reaches (section
 * properties, styles, theme, fonts, the original package) are the model's
 * own, so the document is handed out read-only and a writer copies it first.
 * Each test deep-freezes a model, edits the body, takes the document, and
 * checks that taking it wrote nothing, that it reads as the document of the
 * same model unfrozen, and that the copy a writer takes of it shares no
 * object with the model.
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
    // SAFETY: getDocument reads only the view's state.
    // eslint-disable-next-line typescript/no-unsafe-type-assertion
    getView: () => view as EditorView,
    getDocumentContext: () => model,
    isDestroying: () => false,
    ensureView: () => {},
    isViewRequested: () => true,
  });
};

/** The document `getDocument` builds over `model` after the edit. */
const documentAfterEdit = (
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
  const frozen = documentAfterEdit(model, edit);
  const plain = documentAfterEdit(
    cloneDocumentWithParagraphPropertySources(source),
    edit,
  );
  expect(frozen).toEqual(plain);
  const owned = cloneDocumentWithParagraphPropertySources(frozen);
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

describe("the hidden editor's document in the default session", () => {
  test("is built over the model: it holds the model's own section properties", () => {
    const model = deepFreeze(
      cloneDocumentWithParagraphPropertySources(
        createEmptyDocument({ initialText: "Body" }),
      ),
    );
    const handed = documentAfterEdit(model, { index: 0, text: "edited " });
    const sections =
      model.package.document.finalSectionProperties ??
      panic("Expected section properties");
    // What no edit touched is the model's own: read it, never write it.
    expect(handed.package.document.finalSectionProperties).toBe(sections);
    expect(JSON.stringify(handed.package.document.content)).toContain(
      "edited ",
    );
  });

  test("over every document shape, edited or not, reads a frozen model and a writer's copy of it shares nothing with the model", async () => {
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
        id: "the hidden editor's document never writes the model",
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
      // SAFETY: getDocument reads only the view's state.
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
