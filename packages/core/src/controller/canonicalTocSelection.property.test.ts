import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import {
  applyDocumentOps,
  compileEditorIntent,
  OP_STORIES,
  zeroWidthLeavesAt,
} from "@stll/docx-core/ops";
import type { Document } from "@stll/docx-core/model";
import { headingOutlineLevel } from "@stll/docx-core/model";
import { EditorState, TextSelection } from "prosemirror-state";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { mapTocSelection } from "./canonicalTocSelection";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";
import { prepareCanonicalCommands } from "./canonicalStructure";
import { schema, singletonManager } from "../prosemirror/schema";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";

setDefaultTimeout(propertyTestTimeout(120_000));

// The TOC history generator previously varied text offsets but omitted marker ordinals.
test("TOC splits preserve every endpoint between source bookmark markers", () => {
  assertProperty(
    fc.property(fc.integer({ min: 1, max: 12 }), (markerCount) => {
      const markers = Array.from(
        { length: markerCount },
        (_, id) =>
          ({
            type: "bookmarkStart",
            id,
            name: `Bookmark${id}`,
          }) as const,
      );
      const document = {
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                paraId: "12345678",
                formatting: { outlineLevel: headingOutlineLevel(0) },
                content: [{ type: "run", content: [{ type: "text", text: "Heading" }] }],
              },
              {
                type: "paragraph",
                paraId: "23456789",
                content: [
                  { type: "run", content: [{ type: "text", text: "ab" }] },
                  ...markers,
                  { type: "run", content: [{ type: "text", text: "cd" }] },
                  ...markers.toReversed().map(({ id }) => ({ type: "bookmarkEnd", id }) as const),
                ],
              },
            ],
          },
        },
      } satisfies Document;
      for (let splitOrdinal = 0; splitOrdinal <= markerCount; splitOrdinal += 1) {
        const at = {
          story: OP_STORIES.MAIN,
          blockId: "23456789",
          offset: 2,
          zeroWidthBefore: splitOrdinal,
        } as const;
        const compiled = compileEditorIntent(document, {
          intent: {
            type: "generateTOC",
            at,
            title: "Contents",
            headings: [{ blockId: "12345678", text: "Heading", level: 0 }],
            tabPosition: 9360,
          },
          mode: { type: "editing" },
        }).unwrap();
        const applied = applyDocumentOps(document, compiled.ops).unwrap().document;
        const mapping = { at, after: compiled.selection, ops: compiled.ops };
        for (let ordinal = 0; ordinal <= markerCount; ordinal += 1) {
          const mapped = mapTocSelection({ ...at, zeroWidthBefore: ordinal }, mapping);
          const paragraph = applied.package.document.content.find(
            (block) => block.type === "paragraph" && block.paraId === mapped.blockId,
          );
          if (paragraph?.type !== "paragraph") return panic("Mapped TOC paragraph missing");
          const remaining = zeroWidthLeavesAt(paragraph.content, mapped.offset).slice(
            mapped.zeroWidthBefore ?? 0,
          );
          // The source markers after an endpoint must stay after it in the retained half.
          const endOrdinal = ordinal < splitOrdinal ? splitOrdinal : markerCount;
          expect(remaining).toEqual(markers.slice(ordinal, endOrdinal));
          expect(mapped.offset).toBe(ordinal < splitOrdinal ? 2 : 0);
          expect(mapped.zeroWidthBefore ?? 0).toBe(
            ordinal < splitOrdinal ? ordinal : ordinal - splitOrdinal,
          );
        }
        for (const reverse of [false, true]) {
          const session = createCanonicalSession(document).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          const from = session.projection.positionAt(at).unwrap();
          const to = session.projection
            .positionAt({ ...at, zeroWidthBefore: markerCount })
            .unwrap();
          state = state.apply(
            state.tr.setSelection(
              TextSelection.create(state.doc, reverse ? to : from, reverse ? from : to),
            ),
          );
          const command = singletonManager.requireCommand("generateTOC")({ title: "Contents" });
          command(state);
          const intents =
            getCanonicalCommandIntents(command, state) ?? panic("Missing TOC descriptor");
          const commit = prepareCanonicalCommands(session, state, intents).unwrap();
          state = publishCanonicalProjection({ session, state, commit }).unwrap().state;
          const selection = session.projection.selectionAt(state).unwrap();
          const start = reverse ? selection.head : selection.anchor;
          const end = reverse ? selection.anchor : selection.head;
          expect(start.blockId).toBe("23456789");
          expect(end.blockId).toBe(start.blockId);
          expect(start.offset).toBe(0);
          expect(end.offset).toBe(0);
          expect(start.zeroWidthBefore ?? 0).toBe(0);
          expect(end.zeroWidthBefore ?? 0).toBe(markerCount - splitOrdinal);
          // Bookmark selections have no text; count native marker nodes to catch boundary drift.
          expect(state.selection.to - state.selection.from).toBe(markerCount - splitOrdinal);
        }
      }
    }),
  );
});
