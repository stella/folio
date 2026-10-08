import { assertExactModel } from "../../../../test/exactModel";
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import { Fragment, Slice, type Node as ProseNode } from "prosemirror-model";
import { EditorState, Plugin, TextSelection } from "prosemirror-state";
import { paragraphLogicalText } from "@stll/docx-core/ops";
import { paragraphNumberingReference, relationshipIdOf } from "@stll/docx-core/model";
import { listRenderingAttrPatch } from "../prosemirror/listRenderingAttrs";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Document } from "../types/document";
import { schema } from "../prosemirror/schema";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";
import { toFlowBlocks } from "../layout-bridge/convert/toFlowBlocks";
import {
  assignDocumentParagraphPropertySourceContract,
  cloneDocumentWithParagraphPropertySources,
  getParagraphPropertySourceToken,
} from "../docx/paragraphPropertySource";
import { prepareCanonicalPaste } from "./canonicalClipboard";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(120_000));

const seed = (): Document => {
  const document: Document = {
    package: {
      document: {
        content: ["ab😀cd", "EFgh"].map((text, index) => ({
          type: "paragraph",
          paraId: index === 0 ? "12345678" : "87654321",
          formatting: { alignment: "center", keepNext: true },
          preservedAttributes: [{ name: "rsidR", value: "00112233" }],
          content: [
            { type: "run", formatting: { italic: true }, content: [{ type: "text", text }] },
          ],
        })),
      },
    },
  };
  assignDocumentParagraphPropertySourceContract(document, "b".repeat(64));
  return document;
};

const modelTexts = (document: Document) =>
  document.package.document.content.map((paragraph) => {
    if (paragraph.type !== "paragraph") throw new TypeError("Expected clipboard paragraph.");
    return paragraphLogicalText(paragraph);
  });
const projectionTexts = (state: EditorState) => {
  const texts: string[] = [];
  state.doc.forEach((paragraph) =>
    texts.push(paragraph.textBetween(0, paragraph.content.size, "", "\uFFFC")),
  );
  return texts;
};

const authoredTextMarks = (doc: ProseNode) => {
  const characters: { text: string; marks: unknown[] }[] = [];
  doc.descendants((node) => {
    if (!node.isText) return;
    const marks = node.marks
      .filter((mark) => ["bold", "italic", "underline"].includes(mark.type.name))
      .map((mark) => mark.toJSON());
    for (const character of node.text ?? "") characters.push({ text: character, marks });
  });
  return characters;
};

const sourceFacts = (document: Document) =>
  document.package.document.content.map((paragraph) => {
    if (paragraph.type !== "paragraph")
      throw new TypeError("Clipboard source paragraph disappeared.");
    return { paraId: paragraph.paraId, sourceToken: getParagraphPropertySourceToken(paragraph) };
  });

const assertClipboardModel = (actual: Document, expected: Document) => {
  assertExactModel(actual, expected);
  assertExactModel(sourceFacts(actual), sourceFacts(expected));
};

const assertClipboardSources = (document: Document, before: Document) => {
  const priorIds = new Set(sourceFacts(before).map(({ paraId }) => paraId));
  const sources = sourceFacts(document);
  const tokens = sources.flatMap(({ sourceToken }) =>
    sourceToken === undefined ? [] : [sourceToken],
  );
  expect(new Set(tokens).size).toBe(tokens.length);
  for (const { paraId, sourceToken } of sources) {
    if (!priorIds.has(paraId)) expect(sourceToken).toBeUndefined();
  }
};

describe("canonical clipboard", () => {
  test("copied comment and note ids cannot bind colliding destination story parts in either editing mode", () => {
    for (const kind of ["comment", "footnote", "endnote"] as const) {
      for (const suggesting of [false, true])
        for (const reverse of [false, true]) {
          const destination = seed();
          const sourceDocument = seed();
          const ownedContent = [
            {
              type: "paragraph",
              paraId: "13572468",
              content: [{ type: "run", content: [{ type: "text", text: "owned story" }] }],
            },
          ] satisfies Paragraph[];
          const foreignContent = [
            {
              type: "paragraph",
              paraId: "13572468",
              content: [{ type: "run", content: [{ type: "text", text: "foreign story" }] }],
            },
          ] satisfies Paragraph[];
          switch (kind) {
            case "comment":
              destination.package.document.comments = [
                { id: 7, author: "Owned", content: ownedContent },
              ];
              sourceDocument.package.document.comments = [
                { id: 7, author: "Foreign", content: foreignContent },
              ];
              break;
            case "footnote":
              destination.package.footnotes = [{ type: "footnote", id: 7, content: ownedContent }];
              sourceDocument.package.footnotes = [
                { type: "footnote", id: 7, content: foreignContent },
              ];
              break;
            case "endnote":
              destination.package.endnotes = [{ type: "endnote", id: 7, content: ownedContent }];
              sourceDocument.package.endnotes = [
                { type: "endnote", id: 7, content: foreignContent },
              ];
              break;
            default:
              kind satisfies never;
          }
          const session = createCanonicalSession(destination).unwrap();
          session.setMode(
            suggesting ? { type: "suggesting", author: "Clipboard" } : { type: "editing" },
          );
          let state = EditorState.create({ schema, doc: session.projection.doc });
          state = state.apply(
            state.tr.setSelection(
              TextSelection.create(state.doc, reverse ? 2 : 1, reverse ? 1 : 2),
            ),
          );
          const copied =
            kind === "comment"
              ? schema.node("commentReference", { commentId: 7 })
              : schema.text("7", [
                  schema.marks["footnoteRef"].create({
                    occurrenceId: "fixture-note",
                    id: "7",
                    noteType: kind,
                  }),
                ]);
          const slice = new Slice(Fragment.from(schema.node("paragraph", null, copied)), 1, 1);
          const before = session.document;
          const version = session.version;
          const projection = session.projection;
          const selection = state.selection.toJSON();
          const sourceSnapshot = cloneDocumentWithParagraphPropertySources(sourceDocument);
          assertExactModel(sourceDocument, sourceSnapshot);
          const prepared = prepareCanonicalPaste({ session, state, slice, sourceDocument });
          expect(prepared.isErr()).toBe(true);
          if (prepared.isOk()) throw new TypeError("Foreign story reference unexpectedly pasted.");
          expect(prepared.error.reason).toBe("refused");
          expect(prepared.error.message).toContain("source story parts");
          assertClipboardModel(session.document, before);
          assertExactModel(sourceDocument, sourceSnapshot);
          assertExactModel(state.selection.toJSON(), selection);
          expect(session.projection).toBe(projection);
          expect(session.version).toBe(version);
          expect(session.canUndo).toBe(false);
          expect(session.canRedo).toBe(false);
        }
    }
  });

  test("same-session comment and note moves retain their owned story references and exact history", () => {
    for (const kind of ["comment", "footnote", "endnote"] as const) {
      const destination = seed();
      const first = destination.package.document.content.at(0);
      if (first?.type !== "paragraph") throw new TypeError("Owned story move fixture disappeared.");
      first.content = [
        { type: "run", content: [{ type: "text", text: "a" }] },
        ...(kind === "comment"
          ? [{ type: "commentReference", id: 7 } as const]
          : [
              {
                type: "run",
                content: [{ type: kind === "footnote" ? "footnoteRef" : "endnoteRef", id: 7 }],
              } as const,
            ]),
        { type: "run", content: [{ type: "text", text: "b" }] },
      ];
      const content = [
        {
          type: "paragraph",
          paraId: "13572468",
          content: [{ type: "run", content: [{ type: "text", text: "owned story" }] }],
        },
      ] satisfies Document["package"]["document"]["content"];
      switch (kind) {
        case "comment":
          destination.package.document.comments = [{ id: 7, author: "Owner", content }];
          break;
        case "footnote":
          destination.package.footnotes = [{ type: "footnote", id: 7, content }];
          break;
        case "endnote":
          destination.package.endnotes = [{ type: "endnote", id: 7, content }];
          break;
      }
      const session = createCanonicalSession(destination).unwrap();
      let state = EditorState.create({ schema, doc: session.projection.doc });
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2, 3)));
      const before = session.document;
      const selection = state.selection.toJSON();
      const slice = state.selection.content();
      state = publishCanonicalProjection({
        session,
        state,
        commit: prepareCanonicalPaste({ session, state, slice, moveTarget: 4 }).unwrap(),
      }).unwrap().state;
      const moved = session.document;
      const paragraph = moved.package.document.content.at(0);
      if (paragraph?.type !== "paragraph")
        throw new TypeError("Moved owned reference paragraph disappeared.");
      let beforeReference = "";
      let references = 0;
      for (const child of paragraph.content) {
        if (child.type === "commentReference") {
          expect(kind).toBe("comment");
          expect(child.id).toBe(7);
          references += 1;
          expect(beforeReference).toBe("ab");
          continue;
        }
        if (child.type !== "run")
          throw new TypeError("Owned reference move introduced an unexpected wrapper.");
        for (const leaf of child.content) {
          if (leaf.type === "text") beforeReference += leaf.text;
          else if (leaf.type === "footnoteRef" || leaf.type === "endnoteRef") {
            expect(leaf.type).toBe(kind === "footnote" ? "footnoteRef" : "endnoteRef");
            expect(leaf.id).toBe(7);
            references += 1;
          } else throw new TypeError("Owned reference move introduced an unexpected leaf.");
          if (references !== 0) expect(beforeReference).toBe("ab");
        }
      }
      expect(references).toBe(1);
      expect(beforeReference).toBe("ab");
      assertExactModel(moved.package.document.comments, before.package.document.comments);
      assertExactModel(moved.package.footnotes, before.package.footnotes);
      assertExactModel(moved.package.endnotes, before.package.endnotes);
      state = publishCanonicalProjection({
        session,
        state,
        commit: session.prepareUndo(state).unwrap(),
      }).unwrap().state;
      assertClipboardModel(session.document, before);
      assertExactModel(state.selection.toJSON(), selection);
      expect(session.canUndo).toBe(false);
      state = publishCanonicalProjection({
        session,
        state,
        commit: session.prepareRedo(state).unwrap(),
      }).unwrap().state;
      assertClipboardModel(session.document, moved);
    }
  });

  test("generated owned story-reference moves preserve arbitrary-target sequences and count atomic refusals", () => {
    assertProperty(
      fc.property(
        fc.constantFrom("comment", "footnote", "endnote"),
        fc.array(fc.nat(), { minLength: 8, maxLength: 16 }),
        (kind, targets) => {
          const destination = seed();
          const first = destination.package.document.content.at(0);
          if (first?.type !== "paragraph")
            throw new TypeError("Owned story move fixture disappeared.");
          first.content = [
            { type: "run", content: [{ type: "text", text: "a" }] },
            ...(kind === "comment"
              ? [{ type: "commentReference", id: 7 } as const]
              : [
                  {
                    type: "run",
                    content: [{ type: kind === "footnote" ? "footnoteRef" : "endnoteRef", id: 7 }],
                  } as const,
                ]),
            { type: "run", content: [{ type: "text", text: "b" }] },
          ];
          const content = [
            {
              type: "paragraph",
              paraId: "13572468",
              content: [{ type: "run", content: [{ type: "text", text: "owned story" }] }],
            },
          ] satisfies Document["package"]["document"]["content"];
          switch (kind) {
            case "comment":
              destination.package.document.comments = [{ id: 7, author: "Owner", content }];
              break;
            case "footnote":
              destination.package.footnotes = [{ type: "footnote", id: 7, content }];
              break;
            case "endnote":
              destination.package.endnotes = [{ type: "endnote", id: 7, content }];
              break;
          }
          const session = createCanonicalSession(destination).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });

          const original = session.document;
          const journal: { before: Document; after: Document }[] = [];
          const refusals = new Map<string, number>();
          let expectedRefusals = 0;
          for (const targetSeed of targets) {
            let sourceFrom: number | undefined;
            let sourceTo: number | undefined;
            state.doc.descendants((node, position) => {
              if (
                node.type.name === "commentReference" ||
                (node.isText && node.marks.some((mark) => mark.type.name === "footnoteRef"))
              ) {
                if (sourceFrom !== undefined)
                  throw new TypeError("Owned reference duplicated during move sequence.");
                sourceFrom = position;
                sourceTo = position + node.nodeSize;
              }
            });
            if (sourceFrom === undefined || sourceTo === undefined)
              throw new TypeError("Owned reference disappeared during move sequence.");
            const gaps: number[] = [];
            state.doc.forEach((paragraph, offset) => {
              for (let gap = 0; gap <= paragraph.content.size; gap += 1)
                gaps.push(offset + 1 + gap);
            });
            const target = gaps.at(targetSeed % gaps.length);
            if (target === undefined) throw new TypeError("Owned reference target disappeared.");
            state = state.apply(
              state.tr.setSelection(TextSelection.create(state.doc, sourceFrom, sourceTo)),
            );
            const before = session.document;
            const beforeSelection = state.selection.toJSON();
            const projection = session.projection;
            const version = session.version;
            const slice = state.selection.content();
            const prepared = prepareCanonicalPaste({ session, state, slice, moveTarget: target });
            if (prepared.isErr()) {
              refusals.set(prepared.error.reason, (refusals.get(prepared.error.reason) ?? 0) + 1);
              expect(target >= sourceFrom && target <= sourceTo).toBe(true);
              expect(prepared.error.reason).toBe("noChange");
              expectedRefusals += 1;
              assertClipboardModel(session.document, before);
              assertExactModel(state.selection.toJSON(), beforeSelection);
              expect(session.projection).toBe(projection);
              expect(session.version).toBe(version);
              continue;
            }
            expect(target < sourceFrom || target > sourceTo).toBe(true);
            const oracle = state.tr.deleteRange(sourceFrom, sourceTo);
            const adjusted = oracle.mapping.map(target);
            oracle.replaceRange(adjusted, adjusted, slice);
            state = publishCanonicalProjection({ session, state, commit: prepared.value }).unwrap()
              .state;
            expect(projectionTexts(state)).toEqual(
              projectionTexts(EditorState.create({ schema, doc: oracle.doc })),
            );
            const after = session.document;
            assertExactModel(after.package.document.comments, original.package.document.comments);
            assertExactModel(after.package.footnotes, original.package.footnotes);
            assertExactModel(after.package.endnotes, original.package.endnotes);
            state = publishCanonicalProjection({
              session,
              state,
              commit: session.prepareUndo(state).unwrap(),
            }).unwrap().state;
            assertClipboardModel(session.document, before);
            assertExactModel(state.selection.toJSON(), beforeSelection);
            state = publishCanonicalProjection({
              session,
              state,
              commit: session.prepareRedo(state).unwrap(),
            }).unwrap().state;
            assertClipboardModel(session.document, after);
            journal.push({ before, after });
          }
          expect([...refusals.keys()]).toEqual(expectedRefusals === 0 ? [] : ["noChange"]);
          expect(refusals.get("noChange") ?? 0).toBe(expectedRefusals);
          for (const entry of journal.toReversed()) {
            state = publishCanonicalProjection({
              session,
              state,
              commit: session.prepareUndo(state).unwrap(),
            }).unwrap().state;
            assertClipboardModel(session.document, entry.before);
          }
          assertClipboardModel(session.document, original);
          for (const entry of journal) {
            state = publishCanonicalProjection({
              session,
              state,
              commit: session.prepareRedo(state).unwrap(),
            }).unwrap().state;
            assertClipboardModel(session.document, entry.after);
          }
        },
      ),
      { numRuns: 100 },
    );
  });

  test("generated multi-digit note moves preserve owned stories and exact arbitrary-target histories", () => {
    assertProperty(
      fc.property(
        fc.constantFrom("footnote", "endnote"),
        fc.integer({ min: 1, max: 2147483647 }),
        fc.array(fc.nat(), { minLength: 8, maxLength: 16 }),
        (kind, noteId, targets) => {
          const destination = seed();
          const first = destination.package.document.content.at(0);
          if (first?.type !== "paragraph")
            throw new TypeError("Owned story move fixture disappeared.");
          first.content = [
            { type: "run", content: [{ type: "text", text: "a" }] },
            {
              type: "run",
              content: [{ type: kind === "footnote" ? "footnoteRef" : "endnoteRef", id: noteId }],
            },
            { type: "run", content: [{ type: "text", text: "b" }] },
          ];
          const content = [
            {
              type: "paragraph",
              paraId: "13572468",
              content: [{ type: "run", content: [{ type: "text", text: "owned story" }] }],
            },
          ] satisfies Document["package"]["document"]["content"];
          switch (kind) {
            case "footnote":
              destination.package.footnotes = [{ type: "footnote", id: noteId, content }];
              break;
            case "endnote":
              destination.package.endnotes = [{ type: "endnote", id: noteId, content }];
              break;
          }
          const session = createCanonicalSession(destination).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });

          const assertOwnedNote = (document: Document) => {
            const references: { type: string; id: number }[] = [];
            for (const paragraph of document.package.document.content) {
              if (paragraph.type !== "paragraph")
                throw new TypeError("Note move left plain story.");
              for (const inline of paragraph.content) {
                if (inline.type !== "run") continue;
                for (const leaf of inline.content) {
                  if (leaf.type === "footnoteRef" || leaf.type === "endnoteRef")
                    references.push({ type: leaf.type, id: leaf.id });
                }
              }
            }
            assertExactModel(references, [
              { type: kind === "footnote" ? "footnoteRef" : "endnoteRef", id: noteId },
            ]);
          };
          const original = session.document;
          assertOwnedNote(original);
          const journal: { before: Document; after: Document }[] = [];
          const refusals = new Map<string, number>();
          let expectedRefusals = 0;
          for (const targetSeed of targets) {
            let sourceFrom: number | undefined;
            let sourceTo: number | undefined;
            state.doc.descendants((node, position) => {
              if (node.isText && node.marks.some((mark) => mark.type.name === "footnoteRef")) {
                if (sourceFrom !== undefined)
                  throw new TypeError("Owned reference duplicated during move sequence.");
                const mark = node.marks.find((entry) => entry.type.name === "footnoteRef");
                expect(mark?.attrs["id"]).toBe(String(noteId));
                expect(mark?.attrs["noteType"]).toBe(kind);
                expect(node.text).toBe(String(noteId));
                expect(node.nodeSize).toBe(String(noteId).length);
                sourceFrom = position;
                sourceTo = position + node.nodeSize;
              }
            });
            if (sourceFrom === undefined || sourceTo === undefined)
              throw new TypeError("Owned reference disappeared during move sequence.");
            const gaps: number[] = [];
            state.doc.forEach((paragraph, offset) => {
              for (let gap = 0; gap <= paragraph.content.size; gap += 1)
                if (session.projection.addressAt(offset + 1 + gap).isOk())
                  gaps.push(offset + 1 + gap);
            });
            const target = gaps.at(targetSeed % gaps.length);
            if (target === undefined) throw new TypeError("Owned reference target disappeared.");
            state = state.apply(
              state.tr.setSelection(TextSelection.create(state.doc, sourceFrom, sourceTo)),
            );
            const before = session.document;
            const beforeSelection = state.selection.toJSON();
            const projection = session.projection;
            const version = session.version;
            const slice = state.selection.content();
            const prepared = prepareCanonicalPaste({ session, state, slice, moveTarget: target });
            if (prepared.isErr()) {
              refusals.set(prepared.error.reason, (refusals.get(prepared.error.reason) ?? 0) + 1);
              expect(target >= sourceFrom && target <= sourceTo).toBe(true);
              expect(prepared.error.reason).toBe("noChange");
              expectedRefusals += 1;
              assertClipboardModel(session.document, before);
              assertExactModel(state.selection.toJSON(), beforeSelection);
              expect(session.projection).toBe(projection);
              expect(session.version).toBe(version);
              continue;
            }
            expect(target < sourceFrom || target > sourceTo).toBe(true);
            const oracle = state.tr.deleteRange(sourceFrom, sourceTo);
            const adjusted = oracle.mapping.map(target);
            oracle.replaceRange(adjusted, adjusted, slice);
            state = publishCanonicalProjection({ session, state, commit: prepared.value }).unwrap()
              .state;
            expect(projectionTexts(state)).toEqual(
              projectionTexts(EditorState.create({ schema, doc: oracle.doc })),
            );
            const after = session.document;
            assertOwnedNote(after);
            assertExactModel(after.package.document.comments, original.package.document.comments);
            assertExactModel(after.package.footnotes, original.package.footnotes);
            assertExactModel(after.package.endnotes, original.package.endnotes);
            state = publishCanonicalProjection({
              session,
              state,
              commit: session.prepareUndo(state).unwrap(),
            }).unwrap().state;
            assertClipboardModel(session.document, before);
            assertExactModel(state.selection.toJSON(), beforeSelection);
            state = publishCanonicalProjection({
              session,
              state,
              commit: session.prepareRedo(state).unwrap(),
            }).unwrap().state;
            assertClipboardModel(session.document, after);
            journal.push({ before, after });
          }
          expect([...refusals.keys()]).toEqual(expectedRefusals === 0 ? [] : ["noChange"]);
          expect(refusals.get("noChange") ?? 0).toBe(expectedRefusals);
          for (const entry of journal.toReversed()) {
            state = publishCanonicalProjection({
              session,
              state,
              commit: session.prepareUndo(state).unwrap(),
            }).unwrap().state;
            assertClipboardModel(session.document, entry.before);
          }
          assertClipboardModel(session.document, original);
          for (const entry of journal) {
            state = publishCanonicalProjection({
              session,
              state,
              commit: session.prepareRedo(state).unwrap(),
            }).unwrap().state;
            assertClipboardModel(session.document, entry.after);
            assertOwnedNote(session.document);
          }
        },
      ),
      { numRuns: 100 },
    );
  });

  test("generated Strict and Transitional image link caches rebind ids while preserving authored metadata", () => {
    assertProperty(
      fc.property(fc.constantFrom("a", "draw", "alternate"), fc.boolean(), (prefix, strict) => {
        const drawingNamespace = strict
          ? "http://purl.oclc.org/ooxml/drawingml/main"
          : "http://schemas.openxmlformats.org/drawingml/2006/main";
        const relationshipNamespace = strict
          ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
          : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const hyperlinkType =
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        const destination = seed();
        destination.package.relationships = new Map(
          [1, 2].map((id) => [
            `rId${id}`,
            {
              id: `rId${id}`,
              type: hyperlinkType,
              target: `https://owned.example.test/${id}`,
              targetMode: "External",
            },
          ]),
        );
        const sourceDocument = seed();
        sourceDocument.package.relationships = new Map(
          [1, 2].map((id) => [
            `rId${id}`,
            {
              id: `rId${id}`,
              type: hyperlinkType,
              target: `https://foreign.example.test/${id}`,
              targetMode: "External",
            },
          ]),
        );
        const session = createCanonicalSession(destination).unwrap();
        let state = EditorState.create({ schema, doc: session.projection.doc });
        const before = session.document;
        const png =
          "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=";
        const slice = new Slice(
          Fragment.from(
            schema.node(
              "paragraph",
              null,
              schema.node("image", {
                src: `data:image/png;base64,${png}`,
                width: 1,
                height: 1,
                hlinkRId: "rId1",
                hlinkClickSource: {
                  rId: relationshipIdOf("rId1"),
                  xml: `<${prefix}:hlinkClick xmlns:${prefix}="${drawingNamespace}" xmlns:rel="${relationshipNamespace}" rel:id="rId1" tooltip="authored tip" tgtFrame="authored-frame" history="0"/>`,
                },
                hlinkHoverXml: `<${prefix}:hlinkHover xmlns:${prefix}="${drawingNamespace}" xmlns:rel="${relationshipNamespace}" rel:id="rId2" tooltip="hover tip"/>`,
              }),
            ),
          ),
          1,
          1,
        );
        state = publishCanonicalProjection({
          session,
          state,
          commit: prepareCanonicalPaste({ session, state, slice, sourceDocument }).unwrap(),
        }).unwrap().state;
        const drawings = session.document.package.document.content.flatMap((paragraph) =>
          paragraph.type !== "paragraph"
            ? []
            : paragraph.content.flatMap((content) =>
                content.type !== "run"
                  ? []
                  : content.content.filter((leaf) => leaf.type === "drawing"),
              ),
        );
        const drawing = drawings.at(0);
        if (drawing?.rawXmlMode !== undefined || drawing === undefined)
          throw new TypeError("Imported link-cache drawing disappeared.");
        const image = drawing.image;
        const click = [...(session.document.package.relationships?.values() ?? [])].find(
          (relation) => relation.target === "https://foreign.example.test/1",
        );
        const hover = [...(session.document.package.relationships?.values() ?? [])].find(
          (relation) => relation.target === "https://foreign.example.test/2",
        );
        expect(click?.id).not.toBe("rId1");
        expect(hover?.id).not.toBe("rId2");
        expect(image.hlinkClickSource?.rId).toBe(click?.id);
        expect(image.hlinkClickSource?.xml).toContain(`="${click?.id}"`);
        expect(image.hlinkClickSource?.xml).toContain('tooltip="authored tip"');
        expect(image.hlinkClickSource?.xml).toContain('tgtFrame="authored-frame"');
        expect(image.hlinkClickSource?.xml).toContain('history="0"');
        expect(image.hlinkHoverXml).toContain(`="${hover?.id}"`);
        expect(image.hlinkHoverXml).toContain('tooltip="hover tip"');
        const after = session.document;
        state = publishCanonicalProjection({
          session,
          state,
          commit: session.prepareUndo(state).unwrap(),
        }).unwrap().state;
        assertClipboardModel(session.document, before);
        state = publishCanonicalProjection({
          session,
          state,
          commit: session.prepareRedo(state).unwrap(),
        }).unwrap().state;
        assertClipboardModel(session.document, after);
      }),
      { numRuns: 25 },
    );
  });

  test("foreign text and image hyperlink relationships never bind colliding destination ids", async () => {
    const hyperlinkType =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    const png =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=";
    await assertProperty(
      fc.asyncProperty(
        fc.array(
          fc.record({
            kind: fc.constantFrom("text", "image"),
            sourceKind: fc.constantFrom("modeled", "package", "missing"),
            sourceBuffer: fc.boolean(),
            profile: fc.constantFrom("strict", "transitional"),
          }),
          { minLength: 8, maxLength: 16 },
        ),
        async (trace) => {
          let refusalCount = 0;
          for (const { kind, sourceKind, sourceBuffer, profile } of trace) {
            let destination = seed();
            destination.package.relationships = new Map([
              [
                "rId1",
                {
                  id: "rId1",
                  type: hyperlinkType,
                  target: "https://owned.example.test",
                  targetMode: "External",
                },
              ],
            ]);
            if (sourceBuffer) {
              const paragraph = destination.package.document.content.at(0);
              if (paragraph?.type !== "paragraph")
                throw new TypeError("Expected hyperlink destination paragraph.");
              paragraph.content.push({
                type: "hyperlink",
                rId: "rId1",
                href: "https://owned.example.test",
                children: [{ type: "run", content: [{ type: "text", text: "owned" }] }],
              });
              const zip = await JSZip.loadAsync(await createDocx(destination));
              await Promise.all(
                Object.values(zip.files)
                  .filter(
                    (file) =>
                      !file.dir && (file.name.endsWith(".xml") || file.name.endsWith(".rels")),
                  )
                  .map(async (file) => {
                    let xml = await file.async("text");
                    if (profile === "strict") {
                      xml = xml
                        .replaceAll(
                          "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                          "http://purl.oclc.org/ooxml/wordprocessingml/main",
                        )
                        .replaceAll(
                          "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                          "http://purl.oclc.org/ooxml/officeDocument/relationships",
                        );
                    }
                    if (file.name.endsWith(".rels")) {
                      xml = xml
                        .replace("<Relationships xmlns=", "<pkg:Relationships xmlns:pkg=")
                        .replaceAll("<Relationship ", "<pkg:Relationship ")
                        .replace("</Relationships>", "</pkg:Relationships>");
                    }
                    zip.file(file.name, xml);
                  }),
              );
              destination = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
                preloadFonts: false,
                detectVariables: false,
              });
            }
            const collisionId = [...(destination.package.relationships?.values() ?? [])].find(
              (relation) => relation.target === "https://owned.example.test",
            )?.id;
            if (collisionId === undefined)
              throw new TypeError("Expected owned hyperlink collision.");
            const sourceDocument = seed();
            // A modeled source must supply the defaults needed by a parsed destination.
            if (destination.package.styles !== undefined)
              sourceDocument.package.styles = structuredClone(destination.package.styles);
            sourceDocument.package.relationships = new Map([
              [
                collisionId,
                {
                  id: collisionId,
                  type: hyperlinkType,
                  target: "https://foreign.example.test",
                  targetMode: "External",
                },
              ],
            ]);
            const session = createCanonicalSession(destination).unwrap();
            let state = EditorState.create({ schema, doc: session.projection.doc });
            const before = session.document;
            const node =
              kind === "text"
                ? schema.text("foreign link", [
                    schema.marks["hyperlink"].create({
                      href: sourceKind === "modeled" ? "https://foreign.example.test" : "",
                      rId: collisionId,
                    }),
                  ])
                : schema.node("image", {
                    src: `data:image/png;base64,${png}`,
                    width: 1,
                    height: 1,
                    hlinkRId: collisionId,
                    hlinkClickSource: {
                      rId: collisionId,
                      xml: `<d:hlinkClick xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:link="${profile === "strict" ? "http://purl.oclc.org/ooxml/officeDocument/relationships" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}" link:id="${collisionId}" tooltip="Clipboard tooltip" history="0"/>`,
                    },
                    ...(sourceKind === "modeled"
                      ? { hlinkHref: "https://foreign.example.test" }
                      : {}),
                  });
            const slice = new Slice(Fragment.from(schema.node("paragraph", null, node)), 1, 1);
            const prepared = prepareCanonicalPaste({
              session,
              state,
              slice,
              ...(sourceKind === "package" ? { sourceDocument } : {}),
            });
            if (sourceKind === "missing") {
              if (prepared.isOk())
                throw new TypeError("Missing foreign hyperlink source must refuse.");
              expect(prepared.error.reason).toBe("refused");
              refusalCount += 1;
              assertClipboardModel(session.document, before);
              expect(session.canUndo).toBe(false);
            } else {
              state = publishCanonicalProjection({
                session,
                state,
                commit: prepared.unwrap(),
              }).unwrap().state;
              const after = session.document;
              const foreignRelations = [...(after.package.relationships?.values() ?? [])].filter(
                (relation) =>
                  relation.type === hyperlinkType &&
                  relation.target === "https://foreign.example.test",
              );
              expect(foreignRelations).toHaveLength(1);
              expect(foreignRelations.at(0)?.id).not.toBe(collisionId);
              assertExactModel(
                after.package.relationships?.get(collisionId),
                before.package.relationships?.get(collisionId),
              );
              const exactBeforeSave = cloneDocumentWithParagraphPropertySources(after);
              assertExactModel(after, exactBeforeSave);
              const sourceBytes =
                after.originalBuffer === undefined
                  ? undefined
                  : [...new Uint8Array(after.originalBuffer)];
              const saved = await createDocx(after);
              assertExactModel(after, exactBeforeSave);
              if (sourceBytes !== undefined)
                expect([...new Uint8Array(after.originalBuffer ?? new ArrayBuffer(0))]).toEqual(
                  sourceBytes,
                );
              const reopened = await parseDocx(saved, {
                preloadFonts: false,
                detectVariables: false,
              });
              expect(
                [...(reopened.package.relationships?.values() ?? [])].some(
                  (relation) =>
                    relation.type === hyperlinkType &&
                    relation.target === "https://foreign.example.test",
                ),
              ).toBe(true);
              const paragraphs = reopened.package.document.content.filter(
                (paragraph) => paragraph.type === "paragraph",
              );
              if (kind === "text") {
                expect(
                  paragraphs.some((paragraph) =>
                    paragraph.content.some(
                      (content) =>
                        content.type === "hyperlink" &&
                        content.href === "https://foreign.example.test",
                    ),
                  ),
                ).toBe(true);
              } else {
                const drawings = paragraphs.flatMap((paragraph) =>
                  paragraph.content.flatMap((content) =>
                    content.type === "run"
                      ? content.content.filter((leaf) => leaf.type === "drawing")
                      : [],
                  ),
                );
                expect(drawings).toHaveLength(1);
                const drawing = drawings.at(0);
                if (drawing === undefined)
                  throw new TypeError("Reopened clipboard drawing disappeared.");
                expect(drawing.rawXmlMode).toBeUndefined();
                // Drawing links use the URL parser's normalized href on reopen.
                expect(drawing.image.hlinkHref).toBe(new URL("https://foreign.example.test").href);
                expect(drawing.image.hlinkClickSource?.xml).toContain(
                  'tooltip="Clipboard tooltip"',
                );
                expect(drawing.image.hlinkClickSource?.xml).toContain('history="0"');
              }
              state = publishCanonicalProjection({
                session,
                state,
                commit: session.prepareUndo(state).unwrap(),
              }).unwrap().state;
              assertClipboardModel(session.document, before);
              state = publishCanonicalProjection({
                session,
                state,
                commit: session.prepareRedo(state).unwrap(),
              }).unwrap().state;
              assertClipboardModel(session.document, after);
            }
          }
          expect(refusalCount).toBe(
            trace.filter(({ sourceKind }) => sourceKind === "missing").length,
          );
        },
      ),
      { numRuns: 8 },
    );
  });

  test("divergent paste after undo never reallocates retired paragraph identities", () => {
    const session = createCanonicalSession(seed()).unwrap();
    let state = EditorState.create({ schema, doc: session.projection.doc });
    const before = session.document;
    const originalIds = new Set(sourceFacts(before).map(({ paraId }) => paraId));
    const slice = (text: string) =>
      new Slice(Fragment.from(schema.node("paragraph", null, schema.text(text))), 0, 0);
    state = publishCanonicalProjection({
      session,
      state,
      commit: prepareCanonicalPaste({ session, state, slice: slice("first") }).unwrap(),
    }).unwrap().state;
    const retiredIds = sourceFacts(session.document)
      .map(({ paraId }) => paraId)
      .filter((id) => !originalIds.has(id));
    expect(retiredIds.length).toBeGreaterThan(0);
    state = publishCanonicalProjection({
      session,
      state,
      commit: session.prepareUndo(state).unwrap(),
    }).unwrap().state;
    assertClipboardModel(session.document, before);
    state = publishCanonicalProjection({
      session,
      state,
      commit: prepareCanonicalPaste({ session, state, slice: slice("divergent") }).unwrap(),
    }).unwrap().state;
    const divergentIds = sourceFacts(session.document)
      .map(({ paraId }) => paraId)
      .filter((id) => !originalIds.has(id));
    expect(divergentIds.length).toBeGreaterThan(0);
    expect(divergentIds.some((id) => retiredIds.includes(id))).toBe(false);
    expect(session.canRedo).toBe(false);
  });

  test("unknown clipboard style ids flatten explicit formatting instead of inheriting colliding local styles", () => {
    const destination = seed();
    destination.package.styles = {
      styles: [{ styleId: "Foreign", type: "paragraph", rPr: { italic: true } }],
    };
    const session = createCanonicalSession(destination).unwrap();
    let state = EditorState.create({ schema, doc: session.projection.doc });
    const before = session.document;
    const slice = new Slice(
      Fragment.from(
        schema.node(
          "paragraph",
          { styleId: "Foreign", alignment: "right" },
          schema.text("unowned style", [schema.marks["bold"].create()]),
        ),
      ),
      0,
      0,
    );
    state = publishCanonicalProjection({
      session,
      state,
      commit: prepareCanonicalPaste({ session, state, slice }).unwrap(),
    }).unwrap().state;
    const pasted = session.document.package.document.content.find(
      (paragraph) =>
        paragraph.type === "paragraph" && paragraphLogicalText(paragraph).includes("unowned style"),
    );
    if (pasted?.type !== "paragraph")
      throw new TypeError("Flattened clipboard paragraph disappeared.");
    expect(pasted.formatting?.styleId).toBeUndefined();
    expect(pasted.formatting?.alignment).toBe("right");
    const textRun = pasted.content.find((content) => content.type === "run");
    if (textRun?.type !== "run") throw new TypeError("Flattened clipboard run disappeared.");
    expect(textRun.formatting?.bold).toBe(true);
    expect(textRun.formatting?.italic ?? false).toBe(false);
    assertExactModel(session.document.package.styles, before.package.styles);
  });

  test("foreign numbering collisions preserve each observed format, start, level and marker formatting", () => {
    for (const numFmt of ["decimal", "lowerLetter", "upperRoman"] as const) {
      for (const level of [0, 1, 2]) {
        const destination = seed();
        destination.package.numbering = {
          nums: [{ numId: 1, abstractNumId: 1 }],
          abstractNums: [
            {
              abstractNumId: 1,
              levels: [{ ilvl: 0, start: 1, numFmt: "decimal", lvlText: "%1." }],
            },
          ],
        };
        const session = createCanonicalSession(destination).unwrap();
        let state = EditorState.create({ schema, doc: session.projection.doc });
        const before = session.document;
        const ownedNumbering = structuredClone(before.package.numbering);
        const markerTemplate = `Clause %${level + 1})`;
        const starts = Array.from({ length: level + 1 }, (_, index) => 3 + index);
        const slice = new Slice(
          Fragment.from(
            schema.node(
              "paragraph",
              {
                numPr: paragraphNumberingReference({ numId: 1, ilvl: level }),
                ...listRenderingAttrPatch({
                  numId: 1,
                  abstractNumId: 1,
                  level,
                  marker: markerTemplate,
                  markerTemplate,
                  numFmt,
                  isBullet: false,
                  levelStarts: starts,
                  levelNumFmts: Array.from({ length: level + 1 }, () => numFmt),
                  markerFormatting: { italic: true, fontSize: 24 },
                }),
              },
              schema.text("foreign list"),
            ),
          ),
          0,
          0,
        );
        state = publishCanonicalProjection({
          session,
          state,
          commit: prepareCanonicalPaste({ session, state, slice }).unwrap(),
        }).unwrap().state;
        const pasted = session.document.package.document.content.find(
          (paragraph) =>
            paragraph.type === "paragraph" &&
            paragraphLogicalText(paragraph).includes("foreign list"),
        );
        if (pasted?.type !== "paragraph" || pasted.formatting?.numPr?.kind !== "reference")
          throw new TypeError("Imported list reference disappeared.");
        const numId = pasted.formatting.numPr.numId;
        expect(numId).not.toBe(1);
        expect(pasted.formatting.numPr.ilvl).toBe(level);
        const numbering = session.document.package.numbering;
        const importedNum = numbering?.nums.find((num) => num.numId === numId);
        const imported = numbering?.abstractNums.find(
          (abstract) => abstract.abstractNumId === importedNum?.abstractNumId,
        );
        const importedLevel = imported?.levels.find((entry) => entry.ilvl === level);
        expect(imported?.abstractNumId).not.toBe(1);
        expect(importedLevel?.numFmt).toBe(numFmt);
        expect(importedLevel?.start).toBe(starts.at(level));
        expect(importedLevel?.lvlText).toBe(markerTemplate);
        expect(importedLevel?.rPr?.italic).toBe(true);
        expect(importedLevel?.rPr?.fontSize).toBe(24);
        assertExactModel(
          numbering?.nums.find((num) => num.numId === 1),
          ownedNumbering?.nums.at(0),
        );
        assertExactModel(
          numbering?.abstractNums.find((abstract) => abstract.abstractNumId === 1),
          ownedNumbering?.abstractNums.at(0),
        );
        const after = session.document;
        state = publishCanonicalProjection({
          session,
          state,
          commit: session.prepareUndo(state).unwrap(),
        }).unwrap().state;
        assertClipboardModel(session.document, before);
        state = publishCanonicalProjection({
          session,
          state,
          commit: session.prepareRedo(state).unwrap(),
        }).unwrap().state;
        assertClipboardModel(session.document, after);
      }
    }
  });

  test("source package style collisions import definitions and aliases without changing owned styles", () => {
    const destination = seed();
    destination.package.numbering = {
      nums: [{ numId: 1, abstractNumId: 1 }],
      abstractNums: [
        { abstractNumId: 1, levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }] },
      ],
    };
    destination.package.styles = {
      styles: [{ styleId: "Shared", type: "paragraph", name: "Owned", rPr: { italic: true } }],
    };
    const sourceDocument = seed();
    sourceDocument.package.numbering = {
      nums: [{ numId: 1, abstractNumId: 1 }],
      abstractNums: [
        { abstractNumId: 1, levels: [{ ilvl: 0, start: 8, numFmt: "upperRoman", lvlText: "%1)" }] },
      ],
    };
    sourceDocument.package.styles = {
      styles: [
        {
          styleId: "Shared",
          type: "paragraph",
          name: "Foreign",
          pPr: { numPr: paragraphNumberingReference({ numId: 1, ilvl: 0 }) },
          basedOn: "Base",
          next: "Shared",
          link: "Character",
          rPr: { bold: true },
        },
        { styleId: "Base", type: "paragraph", pPr: { keepNext: true } },
        {
          styleId: "Character",
          type: "character",
          link: "Shared",
          rPr: { underline: { style: "single" } },
        },
      ],
    };
    const session = createCanonicalSession(destination, destination.package.styles).unwrap();
    let state = EditorState.create({ schema, doc: session.projection.doc });
    const before = session.document;
    const slice = new Slice(
      Fragment.from(
        schema.node(
          "paragraph",
          { styleId: "Shared" },
          schema.text("foreign styled", [schema.marks["bold"].create()]),
        ),
      ),
      0,
      0,
    );
    state = publishCanonicalProjection({
      session,
      state,
      commit: prepareCanonicalPaste({ session, state, slice, sourceDocument }).unwrap(),
    }).unwrap().state;
    const pasted = session.document.package.document.content.find(
      (paragraph) =>
        paragraph.type === "paragraph" &&
        paragraphLogicalText(paragraph).includes("foreign styled"),
    );
    if (pasted?.type !== "paragraph") throw new TypeError("Imported style paragraph disappeared.");
    const importedId = pasted.formatting?.styleId;
    expect(importedId).not.toBe("Shared");
    const styles = session.document.package.styles?.styles;
    const imported = styles?.find((style) => style.styleId === importedId);
    let projected: ProseNode | undefined;
    let projectedPosition = -1;
    state.doc.forEach((paragraph, position) => {
      if (paragraph.textContent.includes("foreign styled")) {
        projected = paragraph;
        projectedPosition = position;
      }
    });
    if (!projected) throw new TypeError("Imported style projection disappeared.");
    expect(projected.attrs["listMarker"]).toBe("%1)");
    expect(projected.attrs["listNumFmt"]).toBe("upperRoman");
    expect(projected.attrs["listLevelStarts"]?.at(0)).toBe(8);
    const rendered = toFlowBlocks(state.doc).find(
      (block) => block.kind === "paragraph" && block.pmStart === projectedPosition,
    );
    if (rendered?.kind !== "paragraph")
      throw new TypeError("Imported style layout paragraph disappeared.");
    expect(rendered.attrs?.listMarker).toBe("VIII)");
    expect(imported?.name).toBe("Foreign");
    expect(imported?.rPr?.bold).toBe(true);
    const importedNumbering = imported?.pPr?.numPr;
    if (importedNumbering?.kind !== "reference")
      throw new TypeError("Style numbering dependency disappeared.");
    expect(importedNumbering.numId).not.toBe(1);
    const num = session.document.package.numbering?.nums.find(
      (entry) => entry.numId === importedNumbering.numId,
    );
    const abstract = session.document.package.numbering?.abstractNums.find(
      (entry) => entry.abstractNumId === num?.abstractNumId,
    );
    expect(abstract?.levels.at(0)?.numFmt).toBe("upperRoman");
    expect(abstract?.levels.at(0)?.start).toBe(8);
    expect(imported?.next).toBe(importedId);
    const parent = styles?.find((style) => style.styleId === imported?.basedOn);
    const linked = styles?.find((style) => style.styleId === imported?.link);
    expect(parent?.pPr?.keepNext).toBe(true);
    expect(linked?.rPr?.underline?.style).toBe("single");
    expect(linked?.link).toBe(importedId);
    assertExactModel(
      styles?.find((style) => style.styleId === "Shared"),
      before.package.styles?.styles.at(0),
    );
    const after = session.document;
    state = publishCanonicalProjection({
      session,
      state,
      commit: session.prepareUndo(state).unwrap(),
    }).unwrap().state;
    assertClipboardModel(session.document, before);
    state = publishCanonicalProjection({
      session,
      state,
      commit: session.prepareRedo(state).unwrap(),
    }).unwrap().state;
    assertClipboardModel(session.document, after);
  });

  test("generated clipboard sequences preserve every intermediate journal state at arbitrary targets", () => {
    const refusals = new Map<string, number>();
    const appliedKinds = new Set<string>();
    const text = fc.string({ unit: fc.constantFrom("x", "é", "😀"), minLength: 1, maxLength: 5 });
    assertProperty(
      fc.property(
        fc.array(
          fc.record({
            kind: fc.constantFrom("paste", "move"),
            anchor: fc.nat(),
            head: fc.nat(),
            target: fc.nat(),
            reverse: fc.boolean(),
            openStart: fc.constantFrom(0 as const, 1 as const),
            openEnd: fc.constantFrom(0 as const, 1 as const),
            texts: fc.array(text, { minLength: 1, maxLength: 3 }),
            mark: fc.constantFrom("bold", "italic", "underline"),
          }),
          { minLength: 8, maxLength: 16 },
        ),
        (steps) => {
          const session = createCanonicalSession(seed()).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          const baseline = session.document;
          const issuedIds = new Set(sourceFacts(baseline).map(({ paraId }) => paraId));
          const journal: {
            before: Document;
            after: Document;
            beforeSelection: unknown;
            afterSelection: unknown;
          }[] = [];
          for (const [stepIndex, step] of steps.entries()) {
            const gaps: number[] = [];
            state.doc.forEach((paragraph, offset) => {
              gaps.push(offset + 1);
              let width = 0;
              for (const character of paragraph.textContent) {
                width += character.length;
                gaps.push(offset + 1 + width);
              }
            });
            const anchor = gaps.at(step.anchor % gaps.length);
            const head = gaps.at(step.head % gaps.length);
            const target = gaps.at(step.target % gaps.length);
            if (anchor === undefined || head === undefined || target === undefined)
              throw new TypeError("Generated clipboard target disappeared.");
            const from = Math.min(anchor, head);
            const to = Math.max(anchor, head);
            state = state.apply(
              state.tr.setSelection(
                TextSelection.create(state.doc, step.reverse ? to : from, step.reverse ? from : to),
              ),
            );
            const before = session.document;
            const beforeSelection = state.selection.toJSON();
            const slice =
              step.kind === "move"
                ? state.doc.slice(from, to)
                : new Slice(
                    Fragment.fromArray(
                      step.texts.map((token, index) =>
                        schema.node(
                          "paragraph",
                          { paraId: index === 0 ? "12345678" : "87654321" },
                          schema.text(`paste${token}`, [schema.marks[step.mark].create()]),
                        ),
                      ),
                    ),
                    step.openStart,
                    step.openEnd,
                  );
            const prepared = prepareCanonicalPaste({
              session,
              state,
              slice,
              ...(step.kind === "move" ? { moveTarget: target } : {}),
            });
            if (prepared.isErr()) {
              refusals.set(prepared.error.reason, (refusals.get(prepared.error.reason) ?? 0) + 1);
              // The only generated refusal/no-op is a drop inside its source or an empty move.
              if (step.kind !== "move" || !(from === to || (target >= from && target <= to))) {
                Object.assign(prepared.error, {
                  message: `Clipboard step ${stepIndex} (${step.kind}, from=${from}, to=${to}, target=${target}) refused: ${prepared.error.message}`,
                });
                throw prepared.error;
              }
              expect(step.kind).toBe("move");
              expect(from === to || (target >= from && target <= to)).toBe(true);
              expect(["noChange", "refused"]).toContain(prepared.error.reason);
              assertClipboardModel(session.document, before);
              assertExactModel(state.selection.toJSON(), beforeSelection);
              expect(session.canUndo).toBe(journal.length > 0);
            } else {
              const oracleTransaction = state.tr;
              if (step.kind === "move") {
                oracleTransaction.delete(from, to);
                oracleTransaction.replaceRange(
                  target > to ? target - (to - from) : target,
                  target > to ? target - (to - from) : target,
                  slice,
                );
              } else oracleTransaction.replaceSelection(slice);
              state = publishCanonicalProjection({
                session,
                state,
                commit: prepared.value,
              }).unwrap().state;
              const oracle = projectionTexts(
                EditorState.create({ schema, doc: oracleTransaction.doc }),
              );
              expect(modelTexts(session.document)).toEqual(oracle);
              expect(projectionTexts(state)).toEqual(oracle);
              if (step.kind === "paste")
                assertExactModel(state.selection.toJSON(), oracleTransaction.selection.toJSON());
              expect(authoredTextMarks(state.doc)).toEqual(
                authoredTextMarks(oracleTransaction.doc),
              );
              const ids = session.document.package.document.content.map((paragraph) =>
                paragraph.type === "paragraph" ? paragraph.paraId : undefined,
              );
              expect(ids.every((id) => typeof id === "string")).toBe(true);
              expect(new Set(ids).size).toBe(ids.length);
              const previousIds = new Set(sourceFacts(before).map(({ paraId }) => paraId));
              for (const id of ids) {
                if (!previousIds.has(id)) expect(issuedIds.has(id)).toBe(false);
                issuedIds.add(id);
              }
              expect(state.doc.eq(session.projection.doc)).toBe(true);
              assertClipboardSources(session.document, before);
              const after = session.document;
              const afterSelection = state.selection.toJSON();
              state = publishCanonicalProjection({
                session,
                state,
                commit: session.prepareUndo(state).unwrap(),
              }).unwrap().state;
              assertClipboardModel(session.document, before);
              assertExactModel(state.selection.toJSON(), beforeSelection);
              expect(session.canUndo).toBe(journal.length > 0);
              state = publishCanonicalProjection({
                session,
                state,
                commit: session.prepareRedo(state).unwrap(),
              }).unwrap().state;
              assertClipboardModel(session.document, after);
              assertExactModel(state.selection.toJSON(), afterSelection);
              journal.push({ before, after, beforeSelection, afterSelection });
              appliedKinds.add(step.kind);
            }
          }
          for (const entry of journal.toReversed()) {
            state = publishCanonicalProjection({
              session,
              state,
              commit: session.prepareUndo(state).unwrap(),
            }).unwrap().state;
            assertClipboardModel(session.document, entry.before);
            assertExactModel(state.selection.toJSON(), entry.beforeSelection);
          }
          assertClipboardModel(session.document, baseline);
          expect(session.canUndo).toBe(false);
          for (const entry of journal) {
            state = publishCanonicalProjection({
              session,
              state,
              commit: session.prepareRedo(state).unwrap(),
            }).unwrap().state;
            assertClipboardModel(session.document, entry.after);
            assertExactModel(state.selection.toJSON(), entry.afterSelection);
          }
        },
      ),
      { numRuns: 50 },
    );
    expect([...appliedKinds].sort()).toEqual(["move", "paste"]);
    expect(
      [...refusals.keys()].every((reason) => reason === "refused" || reason === "noChange"),
    ).toBe(true);
    expect([...refusals.values()].reduce((sum, count) => sum + count, 0)).toBeGreaterThan(0);
  });

  test("generated marked slices obey all open edges, replacement directions and exact history", () => {
    assertProperty(
      fc.property(
        fc.array(fc.string({ unit: fc.constantFrom("x", "é", "😀"), minLength: 1, maxLength: 5 }), {
          minLength: 2,
          maxLength: 3,
        }),
        (generatedTokens) => {
          const tokens = generatedTokens.map((token) => `paste${token}`);
          const exercised = new Set<string>();
          for (const openStart of [0, 1])
            for (const openEnd of [0, 1]) {
              for (const cross of [false, true])
                for (const reverse of [false, true]) {
                  exercised.add(`${openStart}:${openEnd}:${cross}:${reverse}`);
                  const session = createCanonicalSession(seed()).unwrap();
                  const second = session.projection.paragraph("87654321");
                  if (!second) throw new TypeError("Missing second destination paragraph.");
                  const from = 2;
                  const to = cross ? second.start + 2 : 5;
                  let state = EditorState.create({ schema, doc: session.projection.doc });
                  state = state.apply(
                    state.tr.setSelection(
                      TextSelection.create(state.doc, reverse ? to : from, reverse ? from : to),
                    ),
                  );
                  const marks = [schema.marks["bold"].create()];
                  const importedIds = tokens.map((_, index) => `4AFE000${index + 1}`);
                  const slice = new Slice(
                    Fragment.fromArray(
                      tokens.map((text, index) =>
                        schema.node(
                          "paragraph",
                          { paraId: importedIds.at(index) },
                          schema.text(text, marks),
                        ),
                      ),
                    ),
                    openStart,
                    openEnd,
                  );
                  // PM is an independent slice-fitting oracle; model operations must preserve its authored text.
                  const oracleState = state.apply(state.tr.replaceSelection(slice));
                  const oracle = projectionTexts(oracleState);
                  const before = session.document;
                  const beforeSelection = state.selection.toJSON();
                  const commit = prepareCanonicalPaste({ session, state, slice }).unwrap();
                  expect(session.document).toBe(before);
                  state = publishCanonicalProjection({ session, state, commit }).unwrap().state;
                  expect(modelTexts(session.document)).toEqual(oracle);
                  assertExactModel(state.selection.toJSON(), oracleState.selection.toJSON());
                  expect(projectionTexts(state)).toEqual(oracle);
                  expect(state.doc.eq(session.projection.doc)).toBe(true);
                  assertClipboardSources(session.document, before);
                  const paragraphs = session.document.package.document.content;
                  const ids = paragraphs.map((paragraph) =>
                    paragraph.type === "paragraph" ? paragraph.paraId : undefined,
                  );
                  expect(new Set(ids).size).toBe(ids.length);
                  expect(ids.every((id) => id !== undefined && !importedIds.includes(id))).toBe(
                    true,
                  );
                  for (const paragraph of paragraphs) {
                    if (paragraph.type !== "paragraph")
                      throw new TypeError("Lost clipboard paragraph.");
                    for (const run of paragraph.content) {
                      if (run.type !== "run") throw new TypeError("Lost clipboard run.");
                      if (
                        run.content.some(
                          (leaf) => leaf.type === "text" && leaf.text.includes("paste"),
                        )
                      ) {
                        expect(run.formatting?.bold).toBe(true);
                        expect(run.formatting?.italic ?? false).toBe(false);
                      }
                    }
                  }
                  const after = session.document;
                  const afterSelection = state.selection.toJSON();
                  state = publishCanonicalProjection({
                    session,
                    state,
                    commit: session.prepareUndo(state).unwrap(),
                  }).unwrap().state;
                  assertClipboardModel(session.document, before);
                  expect(state.selection.toJSON()).toEqual(beforeSelection);
                  expect(session.canUndo).toBe(false);
                  state = publishCanonicalProjection({
                    session,
                    state,
                    commit: session.prepareRedo(state).unwrap(),
                  }).unwrap().state;
                  assertClipboardModel(session.document, after);
                  expect(state.selection.toJSON()).toEqual(afterSelection);
                }
            }
          expect(exercised.size).toBe(16);
        },
      ),
      { numRuns: 25 },
    );
  });

  test("malformed resources and plugin staging leave authority, projection, selection and journal unchanged", () => {
    for (const failure of ["resource", "filter", "append", "throw"] as const) {
      const session = createCanonicalSession(seed()).unwrap();
      let state = EditorState.create({ schema, doc: session.projection.doc });
      let filterTransaction = failure === "filter" ? () => false : undefined;
      if (failure === "throw")
        filterTransaction = () => {
          throw new TypeError("Clipboard staging refused");
        };
      if (failure !== "resource")
        state = state.reconfigure({
          plugins: [
            new Plugin({
              filterTransaction,
              appendTransaction:
                failure === "append"
                  ? (_transactions, _previous, next) => next.tr.insertText("foreign", 1)
                  : undefined,
            }),
          ],
        });
      const before = {
        document: session.document,
        projection: session.projection,
        version: session.version,
        undo: session.canUndo,
        redo: session.canRedo,
        selection: state.selection.toJSON(),
      };
      const content =
        failure === "resource"
          ? schema.node("image", { src: "data:image/png;base64,A", width: 1, height: 1 })
          : schema.text("pasted");
      const slice = new Slice(Fragment.from(schema.node("paragraph", null, content)), 1, 1);
      const prepared = prepareCanonicalPaste({ session, state, slice });
      if (failure === "resource") expect(prepared.isErr()).toBe(true);
      else
        expect(
          publishCanonicalProjection({ session, state, commit: prepared.unwrap() }).isErr(),
        ).toBe(true);
      expect(session.document).toBe(before.document);
      expect(session.projection).toBe(before.projection);
      expect(session.version).toBe(before.version);
      expect(session.canUndo).toBe(before.undo);
      expect(session.canRedo).toBe(before.redo);
      expect(state.selection.toJSON()).toEqual(before.selection);
    }
  });

  test.each([false, true])(
    "inline image import journals package resources and survives save and reopen (hyperlink %s)",
    async (linked) => {
      const png =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=";
      const session = createCanonicalSession(seed()).unwrap();
      let state = EditorState.create({ schema, doc: session.projection.doc });
      const before = session.document;
      const slice = new Slice(
        Fragment.from(
          schema.node("paragraph", { paraId: "4AFE0001" }, [
            schema.text("image"),
            schema.node(
              "image",
              { src: `data:image/png;base64,${png}`, width: 1, height: 1 },
              null,
              linked
                ? [schema.marks["hyperlink"].create({ href: "https://example.test/image" })]
                : [],
            ),
          ]),
        ),
        1,
        1,
      );
      state = publishCanonicalProjection({
        session,
        state,
        commit: prepareCanonicalPaste({ session, state, slice }).unwrap(),
      }).unwrap().state;
      const after = session.document;
      expect(modelTexts(after).at(0)).toContain("image\uFFFC");
      const media = [...(after.package.media?.values() ?? [])];
      expect(media).toHaveLength(1);
      expect([...new Uint8Array(media.at(0)?.data ?? new ArrayBuffer(0))]).toEqual([
        ...Uint8Array.from(atob(png), (character) => character.charCodeAt(0)),
      ]);
      const reopened = await parseDocx(await createDocx(after), {
        preloadFonts: false,
        detectVariables: false,
      });
      expect(modelTexts(reopened)).toEqual(modelTexts(after));
      expect([
        ...new Uint8Array(
          [...(reopened.package.media?.values() ?? [])].at(0)?.data ?? new ArrayBuffer(0),
        ),
      ]).toEqual([...new Uint8Array(media.at(0)?.data ?? new ArrayBuffer(0))]);
      state = publishCanonicalProjection({
        session,
        state,
        commit: session.prepareUndo(state).unwrap(),
      }).unwrap().state;
      assertClipboardModel(session.document, before);
      state = publishCanonicalProjection({
        session,
        state,
        commit: session.prepareRedo(state).unwrap(),
      }).unwrap().state;
      assertClipboardModel(session.document, after);
      expect(state.doc.eq(session.projection.doc)).toBe(true);
    },
  );

  test("moves to either side of their source are one exact journal entry", () => {
    for (const target of [1, 7]) {
      const session = createCanonicalSession(seed()).unwrap();
      let state = EditorState.create({ schema, doc: session.projection.doc });
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2, 5)));
      const before = session.document;
      const beforeSelection = state.selection.toJSON();
      const source = state.doc.textBetween(2, 5);
      const original = modelTexts(before).at(0) ?? "";
      const offset = target - 1;
      const without = original.slice(0, 1) + original.slice(4);
      const adjusted = offset > 4 ? offset - 3 : offset;
      const expected = without.slice(0, adjusted) + source + without.slice(adjusted);
      const slice = new Slice(
        Fragment.from(schema.node("paragraph", null, schema.text(source))),
        1,
        1,
      );
      state = publishCanonicalProjection({
        session,
        state,
        commit: prepareCanonicalPaste({ session, state, slice, moveTarget: target }).unwrap(),
      }).unwrap().state;
      expect(modelTexts(session.document).at(0)).toBe(expected);
      const after = session.document;
      const afterSelection = state.selection.toJSON();
      state = publishCanonicalProjection({
        session,
        state,
        commit: session.prepareUndo(state).unwrap(),
      }).unwrap().state;
      assertClipboardModel(session.document, before);
      expect(state.selection.toJSON()).toEqual(beforeSelection);
      expect(session.canUndo).toBe(false);
      state = publishCanonicalProjection({
        session,
        state,
        commit: session.prepareRedo(state).unwrap(),
      }).unwrap().state;
      assertClipboardModel(session.document, after);
      expect(state.selection.toJSON()).toEqual(afterSelection);
    }
  });
});
