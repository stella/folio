/**
 * The tracker's structure counts, kept per block, equal counting the whole
 * document afresh: the same paragraph and section counts, the same section
 * records and endpoint fingerprint, and fingerprints that are equal for two
 * documents exactly when the full records are.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import { DOCUMENT_SHAPES } from "../../../__tests__/documentShapes";
import {
  createHarnessState,
  HeadlessEditorView,
  parseShapeDocument,
} from "../../../__tests__/editorHarness";
import type { SectionProperties } from "../../../types/document";
import { canonicalJson } from "../../../utils/canonicalJson";
import { sectionPropertiesOf } from "../../sectionCarrier";
import { countDocumentStructure } from "./ParagraphChangeTrackerExtension";

setDefaultTimeout(propertyTestTimeout(120_000));

/** The whole-document count the cached one replaces, kept as the reference. */
const countAfresh = (doc: PMNode) => {
  let paragraphs = 0;
  const sectionRecords = new Set<SectionProperties>();
  const blockRecords: unknown[] = [];
  const endpointRecords: unknown[] = [];
  const visit = (parent: PMNode, parentPath: string): void => {
    parent.forEach((node, _offset, index) => {
      const path =
        parentPath.length === 0 ? `${index}` : `${parentPath}.${index}`;
      blockRecords.push(
        node.type.name === "paragraph"
          ? [path, node.type.name, node.attrs["paraId"] ?? null]
          : [path, node.type.name, canonicalJson(node.attrs)],
      );
      if (node.type.name === "paragraph") {
        paragraphs++;
        const sectionProperties = sectionPropertiesOf(node);
        if (sectionProperties !== null) {
          sectionRecords.add(sectionProperties);
          endpointRecords.push({
            path,
            paraId: node.attrs["paraId"] ?? null,
            sectionProperties,
          });
        }
        return;
      }
      if (node.childCount > 0) visit(node, path);
    });
  };
  visit(doc, "");
  return {
    blockStructureFingerprint: JSON.stringify(blockRecords),
    sectionRecords,
    paragraphs,
    sectionEndpoints: sectionRecords.size,
    sectionEndpointFingerprint: canonicalJson(endpointRecords),
  };
};

const SHAPES = DOCUMENT_SHAPES.filter(({ id }) =>
  ["tables", "sections", "mixed-lists", "comments", "notes"].includes(id),
);

const editArbitrary = fc.record({
  kind: fc.constantFrom("type", "delete", "enter", "undo", "move", "copy"),
  from: fc.double({ min: 0, max: 1, noNaN: true }),
  span: fc.double({ min: 0, max: 0.3, noNaN: true }),
  to: fc.double({ min: 0, max: 1, noNaN: true }),
});
type Edit = fc.RecordValue<typeof editArbitrary>;

const select = (view: HeadlessEditorView, edit: Edit): void => {
  const { doc } = view.state;
  const size = doc.content.size;
  const from = Math.max(0, Math.min(size, Math.floor(edit.from * size)));
  const to = Math.min(size, from + Math.floor(edit.span * size));
  view.dispatch(
    view.state.tr.setSelection(
      TextSelection.between(doc.resolve(from), doc.resolve(to)),
    ),
  );
};

const applyEdit = (view: HeadlessEditorView, edit: Edit): void => {
  switch (edit.kind) {
    case "type":
      select(view, edit);
      view.typeText("x");
      return;
    case "delete":
      select(view, edit);
      view.pressKey("Backspace");
      return;
    case "enter":
      select(view, { ...edit, span: 0 });
      view.pressKey("Enter");
      return;
    case "undo":
      view.pressKey("Mod-z");
      return;
    case "move":
    case "copy": {
      // The same block node at another top-level gap: cached by identity.
      const { doc } = view.state;
      if (doc.childCount < 2) return;
      const index = Math.min(
        doc.childCount - 1,
        Math.floor(edit.from * doc.childCount),
      );
      const node = doc.child(index);
      let start = 0;
      for (let i = 0; i < index; i += 1) start += doc.child(i).nodeSize;
      const target = Math.min(
        doc.childCount,
        Math.floor(edit.to * (doc.childCount + 1)),
      );
      let gap = 0;
      for (let i = 0; i < target; i += 1) gap += doc.child(i).nodeSize;
      const tr = view.state.tr.insert(gap, node);
      if (edit.kind === "move") {
        const moved = tr.mapping.map(start, 1);
        tr.delete(moved, moved + node.nodeSize);
      }
      view.dispatch(tr);
      return;
    }
    default: {
      const exhaustive: never = edit.kind;
      return exhaustive;
    }
  }
};

describe("document structure counts", () => {
  test("kept per block, they equal counting the whole document", async () => {
    const documents = await Promise.all(
      SHAPES.map(async ({ build }) => parseShapeDocument(await build())),
    );
    await assertProperty(
      fc.asyncProperty(
        fc.integer({ min: 0, max: documents.length - 1 }),
        fc.boolean(),
        fc.array(editArbitrary, { minLength: 1, maxLength: 12 }),
        async (which, suggesting, edits) => {
          const view = new HeadlessEditorView(
            createHarnessState(
              documents[which]!,
              suggesting ? "suggesting" : "editing",
            ),
          );
          const docs = [view.state.doc];
          for (const edit of edits) {
            applyEdit(view, edit);
            docs.push(view.state.doc);
          }
          for (const [index, doc] of docs.entries()) {
            const kept = countDocumentStructure(doc);
            const fresh = countAfresh(doc);
            expect(kept.paragraphs).toBe(fresh.paragraphs);
            expect(kept.sectionEndpoints).toBe(fresh.sectionEndpoints);
            expect([...kept.sectionRecords]).toEqual([...fresh.sectionRecords]);
            expect(kept.sectionEndpointFingerprint).toBe(
              fresh.sectionEndpointFingerprint,
            );
            const previous = docs[index - 1];
            if (previous === undefined) continue;
            expect(
              countDocumentStructure(previous).blockStructureFingerprint ===
                kept.blockStructureFingerprint,
            ).toBe(
              countAfresh(previous).blockStructureFingerprint ===
                fresh.blockStructureFingerprint,
            );
          }
        },
      ),
      { id: "document-structure-counts-equal-whole-count", numRuns: 60 },
    );
  });
});
