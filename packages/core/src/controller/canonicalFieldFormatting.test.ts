import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
} from "@stll/docx-core/ops";

import {
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { schema, singletonManager } from "../prosemirror/schema";
import type { Document, ParagraphContent, Run } from "../types/document";
import { prepareCanonicalCommands } from "./canonicalStructure";

const run = (text: string, formatting?: Run["formatting"]): Run => ({
  type: "run",
  ...(formatting === undefined ? {} : { formatting }),
  content: [{ type: "text", text }],
});

// "ab" | PAGE "7" | "cd" | NUMPAGES "1" "2" in two formattings | "ef".
const seed = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          content: [
            run("ab"),
            {
              type: "simpleField",
              instruction: "PAGE",
              fieldType: "PAGE",
              content: [run("7", { italic: true })],
            },
            run("cd"),
            {
              type: "complexField",
              instruction: " NUMPAGES ",
              fieldType: "NUMPAGES",
              fieldCode: [
                {
                  type: "run",
                  content: [{ type: "instrText", text: " NUMPAGES " }],
                },
              ],
              fieldResult: [run("1", { italic: true }), run("2")],
            },
            run("ef"),
          ],
        },
      ],
    },
  },
});

const contentOf = (document: Document): readonly ParagraphContent[] => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") return panic("Missing paragraph.");
  return paragraph.content;
};

/** What each field's result runs state: their properties and review records. */
const fieldResults = (document: Document) =>
  contentOf(document).flatMap((item) => {
    if (item.type !== "simpleField" && item.type !== "complexField") return [];
    const result =
      item.type === "simpleField" ? item.content : item.fieldResult;
    return [
      result.map((entry) =>
        entry.type === "run"
          ? {
              formatting: entry.formatting,
              previous: entry.propertyChanges?.map(
                ({ previousFormatting }) => previousFormatting,
              ),
            }
          : entry.type,
      ),
    ];
  });

const reopened = async (document: Document): Promise<Document> =>
  parseDocx(await createDocx(document), { preloadFonts: false });

const toggleBold = singletonManager.requireCommand("toggleBold")();

/** Bold over the paragraph's units `[from, to)` on the canonical session. */
const formatOnSession = (
  mode: "editing" | "suggesting",
  from: number,
  to: number,
) => {
  const session = createCanonicalSession(seed()).unwrap();
  if (mode === "suggesting")
    session.setMode({ type: "suggesting", author: "Reviewer" });
  const position = (offset: number) =>
    session.projection
      .positionAt({ story: "main", blockId: "12345678", offset })
      .unwrap();
  const state = EditorState.create({
    schema,
    doc: session.projection.doc,
    selection: TextSelection.create(
      session.projection.doc,
      position(from),
      position(to),
    ),
  });
  const before = session.document;
  const intents =
    getCanonicalCommandIntents(toggleBold, state) ??
    panic("Bold has no canonical descriptor.");
  const after = publishCanonicalProjection({
    session,
    state,
    commit: prepareCanonicalCommands(session, state, intents).unwrap(),
  }).unwrap().state;
  return { session, state, before, after };
};

describe("formatting a selection that holds a field on the session", () => {
  test("bold reaches the field's result as the legacy editor writes it, and saves and reopens so", async () => {
    const { session, state } = formatOnSession("editing", 0, 8);
    let legacy = state;
    expect(
      toggleBold(state, (transaction) => {
        legacy = state.apply(transaction);
      }),
    ).toBe(true);
    const legacyDocument: Document = {
      package: { document: { content: proseDocToBlocks(legacy.doc) } },
    };
    const expected = [
      [{ formatting: { italic: true, bold: true }, previous: undefined }],
      [
        { formatting: { italic: true, bold: true }, previous: undefined },
        { formatting: { bold: true }, previous: undefined },
      ],
    ];
    expect(fieldResults(session.document)).toEqual(expected);
    expect(fieldResults(await reopened(session.document))).toEqual(
      fieldResults(await reopened(legacyDocument)),
    );
    expect(fieldResults(await reopened(session.document))).toEqual(expected);
  });

  test("a selection ending at a field's end reaches it; one ending at its start does not", () => {
    expect(
      fieldResults(formatOnSession("editing", 1, 3).session.document),
    ).toEqual([
      [{ formatting: { italic: true, bold: true }, previous: undefined }],
      [
        { formatting: { italic: true }, previous: undefined },
        { formatting: undefined, previous: undefined },
      ],
    ]);
    expect(
      fieldResults(formatOnSession("editing", 0, 2).session.document),
    ).toEqual(fieldResults(seed()));
  });

  test("a suggested bold records the change on each result run; it saves, reopens, accepts, rejects and undoes exactly", async () => {
    const { session, after, before } = formatOnSession("suggesting", 0, 8);
    const expected = [
      [
        {
          formatting: { italic: true, bold: true },
          previous: [{ italic: true }],
        },
      ],
      [
        {
          formatting: { italic: true, bold: true },
          previous: [{ italic: true }],
        },
        { formatting: { bold: true }, previous: [undefined] },
      ],
    ];
    expect(fieldResults(session.document)).toEqual(expected);
    expect(fieldResults(await reopened(session.document))).toEqual(expected);
    const suggested = session.document;
    const undone = publishCanonicalProjection({
      session,
      state: after,
      commit: session.prepareUndo(after).unwrap(),
    }).unwrap().state;
    expect(session.document).toStrictEqual(before);
    publishCanonicalProjection({
      session,
      state: undone,
      commit: session.prepareRedo(undone).unwrap(),
    }).unwrap();
    expect(session.document).toStrictEqual(suggested);
    for (const decision of ["accept", "reject"] as const) {
      const revisionIds = [
        ...new Set(
          contentOf(suggested).flatMap((item) =>
            item.type === "simpleField" || item.type === "complexField"
              ? (item.type === "simpleField"
                  ? item.content
                  : item.fieldResult
                ).flatMap((entry) =>
                  entry.type === "run"
                    ? (entry.propertyChanges?.map(({ info }) => info.id) ?? [])
                    : [],
                )
              : [],
          ),
        ),
      ];
      expect(revisionIds.length).toBeGreaterThan(0);
      const resolved = applyDocumentOp(suggested, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds,
        decision,
      }).unwrap().document;
      expect(fieldResults(resolved)).toEqual(
        decision === "reject"
          ? fieldResults(seed())
          : fieldResults(formatOnSession("editing", 0, 8).session.document),
      );
    }
  });
});
