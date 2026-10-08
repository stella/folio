import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import { headingOutlineLevel } from "@stll/docx-core/model";
import { panic } from "better-result";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Document, Paragraph } from "../types/document";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { schema, singletonManager } from "../prosemirror/schema";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";
import { prepareCanonicalCommands } from "./canonicalStructure";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";

setDefaultTimeout(propertyTestTimeout(120_000));
const fixture = (level: number, localized: boolean, width: number): Document => ({
  package: {
    styles: {
      styles: [
        {
          styleId: localized ? "Nadpis" : "Heading",
          type: "paragraph",
          name: `heading ${level + 1}`,
        },
        { styleId: "Entry", type: "paragraph", name: `toc ${level + 1}` },
        {
          styleId: "Title",
          type: "paragraph",
          name: "TOC Heading",
          pPr: { outlineLevel: { kind: "bodyText" } },
        },
      ],
    },
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          formatting: { styleId: localized ? "Nadpis" : "Heading" },
          content: [
            { type: "bookmarkStart", id: 7, name: "_TocExisting" },
            { type: "bookmarkStart", id: 8, name: "unrelated" },
            {
              type: "run",
              formatting: { bold: true },
              content: [{ type: "text", text: "Heading😀" }],
            },
            { type: "bookmarkEnd", id: 8 },
            { type: "bookmarkEnd", id: 7 },
          ],
        },
        {
          type: "paragraph",
          paraId: "23456789",
          content: [{ type: "run", content: [{ type: "text", text: "Body😀text" }] }],
          sectionProperties: { pageWidth: width, marginLeft: 1000, marginRight: 2000 },
        },
        {
          type: "paragraph",
          paraId: "3456789A",
          formatting: { outlineLevel: headingOutlineLevel(level) },
          content: [{ type: "run", content: [{ type: "text", text: "Second" }] }],
        },
      ],
    },
  },
});
const paragraphs = (document: Document): Paragraph[] =>
  document.package.document.content.map((item) =>
    item.type === "paragraph" ? item : panic("Unexpected TOC block"),
  );
const semantic = (document: Document) =>
  paragraphs(document).map((paragraph) => ({
    id: paragraph.paraId,
    style: paragraph.formatting?.styleId,
    tabs: paragraph.formatting?.tabs,
    content: paragraph.content.map((item) => {
      switch (item.type) {
        case "run":
          return { type: item.type, content: item.content };
        case "hyperlink":
          return {
            type: item.type,
            anchor: item.anchor,
            children: item.children.map((run) => (run.type === "run" ? run.content : run)),
          };
        case "complexField":
          return {
            type: item.type,
            instruction: item.instruction,
            dirty: item.dirty,
            text: item.fieldResult
              .flatMap((run) => run.content)
              .filter((child) => child.type === "text")
              .map((child) => child.text)
              .join(""),
          };
        default:
          return item;
      }
    }),
  }));

test("generated TOC histories preserve anchors, canonical styles, fields, section widths and exact inverse", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.integer({ min: 0, max: 8 }),
      fc.boolean(),
      fc.integer({ min: 6000, max: 20000 }),
      fc.constantFrom(0, 4, 10),
      fc.boolean(),
      async (level, localized, width, offset, reverse) => {
        const session = createCanonicalSession(fixture(level, localized, width)).unwrap();
        let state = EditorState.create({ schema, doc: session.projection.doc });
        const start = session.projection
          .positionAt({ story: "main", blockId: "23456789", offset })
          .unwrap();
        const end = session.projection
          .positionAt({ story: "main", blockId: "23456789", offset: 10 })
          .unwrap();
        state = state.apply(
          state.tr.setSelection(
            TextSelection.create(state.doc, reverse ? end : start, reverse ? start : end),
          ),
        );
        const history = [];
        for (let iteration = 0; iteration < 2; iteration += 1) {
          const before = session.document;
          const beforeSelection = state.selection.toJSON();
          const {
            anchor: _beforeAnchor,
            head: _beforeHead,
            ...selectionMetadata
          } = session.projection.selectionAt(state).unwrap();
          const version = session.version;
          const command = singletonManager.requireCommand("generateTOC")({ title: "Contents" });
          command(state);
          const intents =
            getCanonicalCommandIntents(command, state) ?? panic("Missing TOC descriptor");
          expect(getCanonicalCommandIntents(command, state)).toEqual(intents);
          expect(session.document).toBe(before);
          expect(session.version).toBe(version);
          expect(state.selection.toJSON()).toEqual(beforeSelection);
          session.breakUndoGroup();
          const prepared = prepareCanonicalCommands(session, state, intents).unwrap();
          state = publishCanonicalProjection({ session, state, commit: prepared }).unwrap().state;
          expect(state.doc.eq(session.projection.doc)).toBe(true);
          const {
            anchor: _afterAnchor,
            head: _afterHead,
            ...mappedSelectionMetadata
          } = session.projection.selectionAt(state).unwrap();
          expect(mappedSelectionMetadata).toEqual(selectionMetadata);
          expect(state.selection.empty).toBe(offset === 10);
          expect(
            state.selection.$from.parent.textBetween(
              state.selection.$from.parentOffset,
              state.selection.$to.parentOffset,
            ),
          ).toBe("Body😀text".slice(offset));
          const all = paragraphs(session.document);
          const headings = all.filter(
            (paragraph) => paragraph.paraId === "12345678" || paragraph.paraId === "3456789A",
          );
          const names = headings.flatMap((paragraph) =>
            paragraph.content
              .filter((item) => item.type === "bookmarkStart")
              .map((item) => item.name),
          );
          expect(names).toContain("_TocExisting");
          expect(names).toContain("unrelated");
          expect(names.filter((name) => name.startsWith("_Toc"))).toHaveLength(2);
          const entries = all.filter((paragraph) => paragraph.formatting?.styleId === "Entry");
          expect(entries).toHaveLength((iteration + 1) * 2);
          for (const entry of entries) {
            expect(entry.formatting?.tabs).toEqual([
              { position: width - 3000, alignment: "right", leader: "dot" },
            ]);
            const link = entry.content.find((item) => item.type === "hyperlink");
            const field = entry.content.find((item) => item.type === "complexField");
            expect(link?.type === "hyperlink" && names.includes(link.anchor ?? "")).toBe(true);
            expect(field?.type === "complexField" ? field.instruction : undefined).toBe(
              `PAGEREF ${link?.type === "hyperlink" ? link.anchor : ""} \\h`,
            );
            const address =
              session.projection.paragraph(entry.paraId ?? "") ?? panic("Missing field address");
            expect(address.text.endsWith("\uFFFC\uFFFC")).toBe(true);
            const boundaries = [0];
            let boundary = 0;
            for (const character of address.text) {
              boundary += character.length;
              boundaries.push(boundary);
            }
            for (const index of boundaries) {
              const position = { story: "main", blockId: address.blockId, offset: index } as const;
              expect(
                session.projection
                  .addressAt(session.projection.positionAt(position).unwrap())
                  .unwrap().offset,
              ).toBe(index);
            }
          }
          history.push({
            before,
            beforeSelection,
            after: session.document,
            afterSelection: state.selection.toJSON(),
          });
        }
        expect(semantic(await parseDocx(await createDocx(session.document)))).toEqual(
          semantic(session.document),
        );
        for (const entry of history.toReversed()) {
          state = publishCanonicalProjection({
            session,
            state,
            commit: session.prepareUndo(state).unwrap(),
          }).unwrap().state;
          expect(session.document).toStrictEqual(entry.before);
          expect(state.selection.toJSON()).toEqual(entry.beforeSelection);
        }
        for (const entry of history) {
          state = publishCanonicalProjection({
            session,
            state,
            commit: session.prepareRedo(state).unwrap(),
          }).unwrap().state;
          expect(session.document).toStrictEqual(entry.after);
          expect(state.selection.toJSON()).toEqual(entry.afterSelection);
        }
      },
    ),
    { numRuns: 18 },
  );
});

test("TOC suggesting retains the named expected refusal without publishing headings or allocating ids", () => {
  assertProperty(
    fc.property(fc.integer({ min: 0, max: 8 }), fc.boolean(), (level, localized) => {
      const session = createCanonicalSession(fixture(level, localized, 12240)).unwrap();
      session.setMode({ type: "suggesting", author: "Reviewer" });
      const state = EditorState.create({ schema, doc: session.projection.doc });
      const before = session.document;
      const command = singletonManager.requireCommand("generateTOC")({ title: "Contents" });
      const prepared = prepareCanonicalCommands(
        session,
        state,
        getCanonicalCommandIntents(command, state) ?? [],
      );
      expect(prepared.isErr()).toBe(true);
      if (prepared.isOk()) return panic("Expected TOC refusal stopped reproducing");
      expect(prepared.error.gap).toBe(CANONICAL_GAP.trackedHyperlinkResolution);
      expect(session.document).toBe(before);
      expect(session.version).toBe(0);
      expect(session.canUndo).toBe(false);
    }),
    { numRuns: 18 },
  );
});

test("a document without headings has no TOC changes", () => {
  const source = fixture(0, true, 12240);
  source.package.document.content = paragraphs(source).filter(
    (paragraph) => paragraph.paraId === "23456789",
  );
  const session = createCanonicalSession(source).unwrap();
  const state = EditorState.create({ schema, doc: session.projection.doc });
  const command = singletonManager.requireCommand("generateTOC")({ title: "Contents" });
  expect(command(state)).toBe(false);
  const prepared = prepareCanonicalCommands(
    session,
    state,
    getCanonicalCommandIntents(command, state) ?? [],
  );
  expect(prepared.isErr()).toBe(true);
  if (prepared.isErr()) expect(prepared.error.reason).toBe("noChange");
  expect(session.version).toBe(0);
});

test("PAGEREF cache length never changes atom address width", () => {
  assertProperty(
    fc.property(fc.constantFrom("", "1", "12345", "😀"), (cache) => {
      const source = fixture(0, true, 12240);
      source.package.document.content = [
        {
          type: "paragraph",
          paraId: "12345678",
          content: [
            { type: "run", content: [{ type: "text", text: "a" }] },
            {
              type: "complexField",
              fieldType: "PAGEREF",
              instruction: "PAGEREF anchor \\h",
              fieldCode: [],
              fieldResult: [{ type: "run", content: [{ type: "text", text: cache }] }],
            },
            { type: "run", content: [{ type: "text", text: "b" }] },
          ],
        },
      ];
      const session = createCanonicalSession(source).unwrap();
      expect(session.projection.paragraph("12345678")?.text).toBe("a\uFFFCb");
      for (const offset of [0, 1, 2, 3]) {
        const position = session.projection
          .positionAt({ story: "main", blockId: "12345678", offset })
          .unwrap();
        expect(position).toBe(offset + 1);
        expect(session.projection.addressAt(position).unwrap().offset).toBe(offset);
      }
    }),
    { numRuns: 12 },
  );
});

test("TOC after undo never reuses retired paragraph identities", () => {
  assertProperty(
    fc.property(fc.boolean(), fc.integer({ min: 0, max: 8 }), (localized, level) => {
      const session = createCanonicalSession(fixture(level, localized, 12240)).unwrap();
      let state = EditorState.create({ schema, doc: session.projection.doc });
      const at = session.projection
        .positionAt({ story: "main", blockId: "23456789", offset: 4 })
        .unwrap();
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, at)));
      const originalIds = new Set(
        paragraphs(session.document).map((paragraph) => paragraph.paraId),
      );
      const execute = (title: string) => {
        const command = singletonManager.requireCommand("generateTOC")({ title });
        const intents =
          getCanonicalCommandIntents(command, state) ?? panic("Missing TOC descriptor");
        session.breakUndoGroup();
        state = publishCanonicalProjection({
          session,
          state,
          commit: prepareCanonicalCommands(session, state, intents).unwrap(),
        }).unwrap().state;
      };
      execute("Contents");
      const retired = new Set(
        paragraphs(session.document)
          .filter((paragraph) => !originalIds.has(paragraph.paraId))
          .map((paragraph) => paragraph.paraId),
      );
      state = publishCanonicalProjection({
        session,
        state,
        commit: session.prepareUndo(state).unwrap(),
      }).unwrap().state;
      execute("New contents");
      const fresh = paragraphs(session.document)
        .filter((paragraph) => !originalIds.has(paragraph.paraId))
        .map((paragraph) => paragraph.paraId);
      expect(fresh.some((id) => retired.has(id))).toBe(false);
      expect(session.canRedo).toBe(false);
    }),
    { numRuns: 12 },
  );
});
