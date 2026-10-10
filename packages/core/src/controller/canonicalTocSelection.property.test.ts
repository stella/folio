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
import {
  readBookmarkBoundaryAttrs,
  bookmarkMarkerFromAttrs,
} from "../prosemirror/bookmarkBoundaryAttrs";

setDefaultTimeout(propertyTestTimeout(120_000));

// The TOC history generator previously varied text offsets but omitted marker ordinals.
test("TOC splits preserve every endpoint between source bookmark markers", () => {
  assertProperty(
    fc.property(
      fc.integer({ min: 0, max: 12 }),
      fc.constantFrom(0, 2, 4),
      fc.constantFrom("body", "heading"),
      (markerCount, offset, targetKind) => {
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
                  ...(targetKind === "heading"
                    ? { formatting: { outlineLevel: headingOutlineLevel(1) } }
                    : {}),
                  content: [
                    ...(offset === 0
                      ? []
                      : [
                          {
                            type: "run",
                            content: [{ type: "text", text: "abcd".slice(0, offset) }],
                          } as const,
                        ]),
                    ...markers,
                    ...(offset === 4
                      ? []
                      : [
                          {
                            type: "run",
                            content: [{ type: "text", text: "abcd".slice(offset) }],
                          } as const,
                        ]),
                  ],
                },
                {
                  type: "paragraph",
                  paraId: "3456789A",
                  content: [
                    ...markers.toReversed().map(({ id }) => ({ type: "bookmarkEnd", id }) as const),
                    { type: "run", content: [{ type: "text", text: "End" }] },
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
            offset,
            zeroWidthBefore: splitOrdinal,
          } as const;
          const compiled = compileEditorIntent(document, {
            intent: {
              type: "generateTOC",
              at,
              title: "Contents",
              headings: [
                { blockId: "12345678", text: "Heading", level: 0 },
                ...(targetKind === "heading"
                  ? [{ blockId: "23456789", text: "abcd", level: 1 } as const]
                  : []),
              ],
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
            const remaining = zeroWidthLeavesAt(paragraph.content, mapped.offset)
              .slice(mapped.zeroWidthBefore ?? 0)
              .filter((marker) => marker.type === "bookmarkStart" && marker.id < markerCount);
            // The source markers after an endpoint must stay after it in the retained half.
            const split = offset === 2;
            const endOrdinal = split && ordinal < splitOrdinal ? splitOrdinal : markerCount;
            expect(remaining).toEqual(markers.slice(ordinal, endOrdinal));
            expect(mapped.offset).toBe(split && ordinal >= splitOrdinal ? 0 : offset);
            expect(mapped.zeroWidthBefore ?? 0).toBe(
              split && ordinal >= splitOrdinal
                ? ordinal - splitOrdinal
                : ordinal + (targetKind === "heading" && offset === 0 ? 1 : 0),
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
            expect(start.offset).toBe(offset === 2 ? 0 : offset);
            expect(end.offset).toBe(start.offset);
            const prepended = targetKind === "heading" && offset === 0 ? 1 : 0;
            expect(start.zeroWidthBefore ?? 0).toBe(offset === 2 ? 0 : splitOrdinal + prepended);
            expect(end.zeroWidthBefore ?? 0).toBe(
              offset === 2 ? markerCount - splitOrdinal : markerCount + prepended,
            );
            // Bookmark selections have no text; count native marker nodes to catch boundary drift.
            expect(state.selection.to - state.selection.from).toBe(markerCount - splitOrdinal);
            const selectedMarkers = state.selection
              .content()
              .content.content.flatMap((node) => (node.isTextblock ? node.content.content : [node]))
              .map((node) => {
                const attrs = readBookmarkBoundaryAttrs(node);
                if (!attrs.ok) return panic("Selection contains a non-bookmark node");
                return bookmarkMarkerFromAttrs(attrs.value);
              });
            expect(selectedMarkers).toEqual(markers.slice(splitOrdinal));
          }
        }
      },
    ),
    { id: "TOC splits preserve every endpoint between source bookmark markers" },
  );
});
