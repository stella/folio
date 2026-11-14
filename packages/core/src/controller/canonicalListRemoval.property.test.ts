import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  type CanonicalSession,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { serializeCanonicalSave } from "../docx/canonicalSave";
import { mintListInstance } from "../docx/listNumberingInstances";
import { NO_NUMBERING_IDS_IN_USE } from "../docx/numberingIds";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { CanonicalCommandIntent } from "../prosemirror/canonicalCommands";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph, Style } from "../types/document";
import { FOLIO_DOCX_SERIALIZATION_MODE } from "../types/docxSerialization";
import {
  isCanonicalListParagraph,
  prepareCanonicalCommands,
} from "./canonicalStructure";

/**
 * Paragraph styles by whether their chain numbers a paragraph: none named, a
 * heading without numbering, a numbered style, and a style based on it.
 */
const STYLES = {
  none: { styleId: undefined, numbers: false },
  heading: { styleId: "Heading1", numbers: false },
  numbered: { styleId: "Numbered", numbers: true },
  derived: { styleId: "Derived", numbers: true },
} as const;

const fixture = async (styleId: string | undefined): Promise<Document> => {
  const minted = mintListInstance(
    undefined,
    { kind: "numbered" },
    NO_NUMBERING_IDS_IN_USE,
  );
  const styles: Style[] = [
    { type: "paragraph", styleId: "Heading1", name: "heading 1" },
    {
      type: "paragraph",
      styleId: "Numbered",
      name: "Numbered",
      pPr: { numPr: { kind: "reference", numId: minted.numId, ilvl: 0 } },
    },
    {
      type: "paragraph",
      styleId: "Derived",
      name: "Derived",
      basedOn: "Numbered",
    },
  ];
  const paragraph = (paraId: string, text: string): Paragraph => ({
    type: "paragraph",
    paraId,
    ...(styleId === undefined ? {} : { formatting: { styleId } }),
    content: [{ type: "run", content: [{ type: "text", text }] }],
  });
  return parseDocx(
    await createDocx({
      package: {
        numbering: minted.definitions,
        styles: { styles },
        document: {
          content: [
            paragraph("10000001", "Item"),
            paragraph("10000002", "Next"),
          ],
        },
      },
    }),
    { preloadFonts: false },
  );
};

const first = (document: Document): Paragraph => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") return panic("Expected a paragraph.");
  return paragraph;
};

const run = (
  session: CanonicalSession,
  state: EditorState,
  command: CanonicalCommandIntent,
): EditorState =>
  publishCanonicalProjection({
    session,
    state,
    commit: prepareCanonicalCommands(session, state, [command]).unwrap(),
  }).unwrap().state;

test(
  "taking a paragraph out of its list restores the numbering it stated, unless its style numbers it",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom(...Object.values(STYLES)),
        fc.array(fc.constantFrom("bullet", "decimal"), {
          minLength: 1,
          maxLength: 4,
        }),
        fc.constantFrom("toggle", "removeList", "outdent"),
        async (style, toggles, clear) => {
          const opened = await fixture(style.styleId);
          const original = structuredClone(first(opened).formatting?.numPr);
          // A paragraph numbered by its style starts in the list; any other
          // paragraph states no numbering of its own.
          expect(first(opened).listRendering !== undefined).toBe(style.numbers);
          if (!style.numbers) expect(original).toBeUndefined();

          const session = createCanonicalSession(opened).unwrap();
          // Enter and Backspace take a paragraph its style numbers out of the
          // list as they take one numbered by its own properties.
          expect(isCanonicalListParagraph(first(session.document))).toBe(
            style.numbers,
          );
          let state = EditorState.create({
            schema,
            doc: session.projection.doc,
          });
          state = state.apply(
            state.tr.setSelection(TextSelection.create(state.doc, 1)),
          );
          for (const kind of toggles)
            state = run(session, state, { type: "toggleList", kind });
          const listed = first(session.document).listRendering;
          if (listed !== undefined) {
            const clears = {
              toggle: {
                type: "toggleList",
                kind: listed.isBullet ? "bullet" : "decimal",
              },
              removeList: { type: "removeList" },
              outdent: { type: "changeListLevel", direction: "decrease" },
            } as const satisfies Record<typeof clear, CanonicalCommandIntent>;
            state = run(session, state, clears[clear]);
          }

          const cleared = first(session.document);
          expect(cleared.listRendering).toBeUndefined();
          expect(isCanonicalListParagraph(cleared)).toBe(false);
          // Only a style that numbers the paragraph needs the cancellation.
          expect(cleared.formatting?.numPr).toEqual(
            style.numbers ? { kind: "none" } : original,
          );
          expect(cleared.formatting?.styleId).toBe(style.styleId);

          for (const mode of Object.values(FOLIO_DOCX_SERIALIZATION_MODE)) {
            const saved = await serializeCanonicalSave({
              snapshot: session.captureSaveSnapshot(),
              options: { mode },
            });
            const reopened = first(
              await parseDocx(saved.buffer, { preloadFonts: false }),
            );
            expect(reopened.listRendering).toBeUndefined();
            expect(reopened.formatting?.numPr).toEqual(
              style.numbers ? { kind: "none" } : original,
            );
            expect(reopened.formatting?.styleId).toBe(style.styleId);
          }
        },
      ),
      {
        id: "taking a paragraph out of its list restores the numbering it stated, unless its style numbers it",
        examples: [
          [STYLES.heading, ["bullet", "decimal"], "toggle"],
          [STYLES.numbered, ["bullet"], "removeList"],
          [STYLES.derived, ["decimal"], "outdent"],
        ],
      },
    );
  },
  propertyTestTimeout(30_000),
);

test("the removal law catches the previous rules", async () => {
  // The previous removal: the cancellation whatever the style says.
  const opened = await fixture(STYLES.heading.styleId);
  const session = createCanonicalSession(opened).unwrap();
  let state = EditorState.create({ schema, doc: session.projection.doc });
  state = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, 1)),
  );
  state = run(session, state, { type: "toggleList", kind: "bullet" });
  const at = session.projection.inputAddressAt(1).unwrap();
  state = publishCanonicalProjection({
    session,
    state,
    commit: session
      .prepareIntent(state, {
        type: "formatParagraph",
        at,
        patch: { numPr: { kind: "none" } },
      })
      .unwrap(),
  }).unwrap().state;
  expect(first(session.document).listRendering).toBeUndefined();
  expect(first(session.document).formatting?.numPr).not.toEqual(
    first(opened).formatting?.numPr,
  );
  // The previous membership: only a paragraph's own numbering.
  const styled = first(
    createCanonicalSession(await fixture(STYLES.numbered.styleId)).unwrap()
      .document,
  );
  expect(styled.formatting?.numPr?.kind === "reference").toBe(false);
  expect(isCanonicalListParagraph(styled)).toBe(true);
});
