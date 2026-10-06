import { expect, test, setDefaultTimeout } from "bun:test";
import fc from "fast-check";
import { panic } from "better-result";
import { AllSelection, EditorState, TextSelection, type Command } from "prosemirror-state";
import type { Document, ParagraphContent } from "../types/document";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { schema, singletonManager } from "../prosemirror/schema";
import { getCanonicalCommandIntents } from "../prosemirror/canonicalCommands";
import type { CanonicalCommandIntent } from "../prosemirror/canonicalCommands";
import { insertPageBreak } from "../prosemirror/commands/pageBreak";
import { createCanonicalSession, publishCanonicalProjection } from "./canonicalSession";
import { prepareCanonicalCommands } from "./canonicalStructure";
import { createDocx } from "../docx/rezip";
import { parseDocx } from "../docx/parser";

setDefaultTimeout(propertyTestTimeout(120_000));
const seed = (bookmarks: boolean): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          content: [
            ...(bookmarks ? [{ type: "bookmarkStart", id: 7, name: "anchor" } as const] : []),
            {
              type: "run",
              formatting: { bold: true, color: { rgb: "123456" } },
              content: [{ type: "text", text: "ab" }],
            },
            {
              type: "hyperlink",
              href: "https://old.example/",
              tooltip: "Old",
              target: "_blank",
              children: [
                {
                  type: "run",
                  formatting: { italic: true },
                  content: [{ type: "text", text: "😀cd" }],
                },
              ],
            },
            { type: "hyperlink", href: "https://empty.example/", children: [] },
            ...(bookmarks ? [{ type: "bookmarkEnd", id: 7 } as const] : []),
          ],
        },
      ],
    },
  },
});

/** Per-character link targets and bookmark positions ignore serializer run merging and rIds. */
const authored = (document: Document) => {
  const tokens: unknown[] = [];
  const walk = (
    items: readonly ParagraphContent[],
    link?: { href?: string; anchor?: string; tooltip?: string; target?: string },
  ) => {
    for (const item of items) {
      switch (item.type) {
        case "run":
          for (const child of item.content)
            if (child.type === "text") for (const text of child.text) tokens.push({ text, link });
          break;
        case "hyperlink":
          if (item.children.length === 0)
            tokens.push({ type: "emptyHyperlink", href: item.href, anchor: item.anchor });
          walk(item.children, {
            href: item.href ?? (item.anchor === undefined ? undefined : `#${item.anchor}`),
            anchor: item.anchor,
            tooltip: item.tooltip,
            target: item.target,
          });
          break;
        case "bookmarkStart":
          tokens.push({ type: item.type, id: item.id, name: item.name });
          break;
        case "bookmarkEnd":
          tokens.push({ type: item.type, id: item.id });
          break;
        case "insertion":
          walk(item.content, link);
          break;
        case "deletion":
          break;
        default:
          return panic("Unexpected generated hyperlink content.");
      }
    }
  };
  for (const paragraph of document.package.document.content) {
    if (paragraph.type !== "paragraph") return panic("Hyperlink history lost its paragraph.");
    walk(paragraph.content);
  }
  return tokens;
};
const inputArbitrary = fc.record({
  left: fc.nat(),
  right: fc.nat(),
  reverse: fc.boolean(),
  href: fc.constantFrom("example.com", "#anchor", "https://new.example/path", "javascript:void 0"),
  tooltip: fc.constantFrom(undefined, "Tip", ""),
});
type HyperlinkInput = ReturnType<typeof inputArbitrary.generate>["value"];
const factories = {
  setHyperlink: ({ href, tooltip }: HyperlinkInput) =>
    singletonManager.requireCommand("setHyperlink")(href, tooltip),
  removeHyperlink: () => singletonManager.requireCommand("removeHyperlink")(),
  insertHyperlink: ({ href, tooltip }: HyperlinkInput) =>
    singletonManager.requireCommand("insertHyperlink")("New😀", href, tooltip),
};

const allSelectionCommands = {
  setHyperlink: [() => singletonManager.requireCommand("setHyperlink")("https://new.example/")],
  removeHyperlink: [() => singletonManager.requireCommand("removeHyperlink")()],
  insertHyperlink: [() => singletonManager.requireCommand("insertHyperlink")("New😀", "#anchor")],
  formatRun: [
    () => singletonManager.requireCommand("toggleUnderline")(),
    () => singletonManager.requireCommand("setFontFamily")("New Font"),
  ],
  insertBreak: [() => insertPageBreak],
} as const satisfies Record<
  Extract<CanonicalCommandIntent, { from: number; to: number }>["type"],
  readonly (() => Command)[]
>;

test("every range intent handles AllSelection and restores exact undo redo", () => {
  assertProperty(
    fc.property(fc.boolean(), fc.integer({ min: 1, max: 3 }), (bookmarks, paragraphCount) => {
      const exercised = new Set<string>();
      for (const [name, commandFactories] of Object.entries(allSelectionCommands)) {
        for (const factory of commandFactories) {
          exercised.add(name);
          const document = seed(bookmarks);
          for (let index = 1; index < paragraphCount; index += 1) {
            document.package.document.content.push({
              type: "paragraph",
              paraId: (0x12345678 + index).toString(16),
              content: [{ type: "run", content: [{ type: "text", text: "tail😀" }] }],
            });
          }
          const session = createCanonicalSession(document).unwrap();
          let state = EditorState.create({ schema, doc: session.projection.doc });
          state = state.apply(state.tr.setSelection(new AllSelection(state.doc)));
          const before = session.document;
          const beforeSelection = state.selection.toJSON();
          const intents =
            getCanonicalCommandIntents(factory(), state) ?? panic("Missing range descriptor.");
          // Page breaks require a caret, so whole-document selection has no intent.
          if (name === "insertBreak") {
            expect(intents).toEqual([]);
            continue;
          }
          expect(intents.length).toBeGreaterThan(0);
          const prepared = prepareCanonicalCommands(session, state, intents).unwrap();
          state = publishCanonicalProjection({ session, state, commit: prepared }).unwrap().state;
          const after = session.document;
          const afterSelection = state.selection.toJSON();
          if (name === "insertHyperlink") {
            expect(state.doc.textContent).toBe("New😀");
          } else {
            expect(afterSelection).toEqual(beforeSelection);
          }
          state = publishCanonicalProjection({
            session,
            state,
            commit: session.prepareUndo(state).unwrap(),
          }).unwrap().state;
          expect(session.document).toStrictEqual(before);
          expect(state.selection.toJSON()).toEqual(beforeSelection);
          state = publishCanonicalProjection({
            session,
            state,
            commit: session.prepareRedo(state).unwrap(),
          }).unwrap().state;
          expect(session.document).toStrictEqual(after);
          expect(state.selection.toJSON()).toEqual(afterSelection);
        }
      }
      expect([...exercised].sort()).toEqual(Object.keys(allSelectionCommands).sort());
    }),
    { numRuns: 12 },
  );
});

test("generated hyperlink command histories preserve targets, bookmarks, probes and exact inverse", async () => {
  await assertProperty(
    fc.asyncProperty(
      fc.boolean(),
      fc.array(inputArbitrary, { minLength: 3, maxLength: 9 }),
      async (bookmarks, inputs) => {
        for (const mode of [
          { type: "editing" },
          { type: "suggesting", author: "Reviewer" },
        ] as const) {
          const document = seed(bookmarks);
          const session = createCanonicalSession(document).unwrap();
          session.setMode(mode);
          let state = EditorState.create({ schema, doc: session.projection.doc });
          const baseline = session.document;
          const history: {
            before: Document;
            after: Document;
            beforeSelection: ReturnType<typeof state.selection.toJSON>;
            afterSelection: ReturnType<typeof state.selection.toJSON>;
          }[] = [];
          const commands = Object.entries(factories);
          const exercised = new Set<string>();
          for (const [index, input] of inputs.entries()) {
            const [name, factory] =
              commands.at(index % commands.length) ?? panic("Missing hyperlink command factory.");
            exercised.add(name);
            const address =
              session.projection.paragraph("12345678") ?? panic("Missing hyperlink paragraph.");
            const text = address.text;
            const boundaries = [0];
            let offset = 0;
            for (const char of text) {
              offset += char.length;
              boundaries.push(offset);
            }
            const left =
              boundaries.at(input.left % boundaries.length) ?? panic("Missing range start.");
            const right =
              boundaries.at(input.right % boundaries.length) ?? panic("Missing range end.");
            const position = (logicalOffset: number) =>
              session.projection
                .positionAt({ story: "main", blockId: address.blockId, offset: logicalOffset })
                .unwrap();
            state = state.apply(
              state.tr.setSelection(
                TextSelection.create(
                  state.doc,
                  position(input.reverse ? right : left),
                  position(input.reverse ? left : right),
                ),
              ),
            );
            const before = session.document;
            const version = session.version;
            const beforeSelection = state.selection.toJSON();
            const marks = state.storedMarks;
            const command = factory(input);
            command(state);
            const intents =
              getCanonicalCommandIntents(command, state) ?? panic(`Missing ${name} descriptor.`);
            expect(getCanonicalCommandIntents(command, state)).toEqual(intents);
            expect(session.document).toBe(before);
            expect(session.version).toBe(version);
            expect(state.selection.toJSON()).toEqual(beforeSelection);
            expect(state.storedMarks).toBe(marks);
            session.breakUndoGroup();
            const prepared = prepareCanonicalCommands(session, state, intents);
            if (prepared.isErr()) {
              if (mode.type === "editing" && prepared.error.reason === "refused")
                panic(prepared.error.message);
              expect(["refused", "noChange"]).toContain(prepared.error.reason);
              if (mode.type === "suggesting" && intents.length > 0) {
                expect(prepared.error.gap).toBe(CANONICAL_GAP.trackedHyperlinkResolution);
              }
              expect(session.document).toBe(before);
              expect(session.version).toBe(version);
              expect(state.selection.toJSON()).toEqual(beforeSelection);
              continue;
            }
            state = publishCanonicalProjection({ session, state, commit: prepared.value }).unwrap()
              .state;
            expect(state.doc.eq(session.projection.doc)).toBe(true);
            if (name !== "insertHyperlink")
              expect(state.selection.toJSON()).toEqual(beforeSelection);
            history.push({
              before,
              after: session.document,
              beforeSelection,
              afterSelection: state.selection.toJSON(),
            });
          }
          if (mode.type === "editing") expect(history.length).toBeGreaterThan(0);
          else expect(history).toEqual([]);
          expect([...exercised].sort()).toEqual(Object.keys(factories).sort());
          expect(authored(await parseDocx(await createDocx(session.document)))).toEqual(
            authored(session.document),
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
          expect(session.document).toStrictEqual(baseline);
          for (const entry of history) {
            state = publishCanonicalProjection({
              session,
              state,
              commit: session.prepareRedo(state).unwrap(),
            }).unwrap().state;
            expect(session.document).toStrictEqual(entry.after);
            expect(state.selection.toJSON()).toEqual(entry.afterSelection);
          }
        }
      },
    ),
    { numRuns: 15 },
  );
});

test("tracked input addresses skip rendered deletions after bookmark atoms", () => {
  assertProperty(
    fc.property(fc.boolean(), fc.constantFrom("a", "😀ab"), (bookmarks, prefix) => {
      const document = seed(bookmarks);
      const paragraph =
        document.package.document.content.at(0) ?? panic("Missing address fixture.");
      if (paragraph.type !== "paragraph") panic("Missing paragraph.");
      const start = bookmarks ? [{ type: "bookmarkStart", id: 7, name: "anchor" } as const] : [];
      const end = bookmarks ? [{ type: "bookmarkEnd", id: 7 } as const] : [];
      paragraph.content = [
        ...start,
        { type: "run", content: [{ type: "text", text: prefix }] },
        {
          type: "deletion",
          info: { id: 1, author: "Reviewer" },
          content: [{ type: "run", content: [{ type: "text", text: "deleted" }] }],
        },
        { type: "run", content: [{ type: "text", text: "tail" }] },
        ...end,
      ];
      const session = createCanonicalSession(document).unwrap();
      let exercised = 0;
      session.projection.doc.descendants((node, position) => {
        if (!node.isText || !node.marks.some((mark) => mark.type.name === "deletion")) return;
        for (let offset = 1; offset < node.nodeSize; offset += 1) {
          expect(session.projection.inputAddressAt(position + offset).unwrap()).toEqual(
            session.projection.addressAt(position + node.nodeSize).unwrap(),
          );
          exercised += 1;
        }
      });
      expect(exercised).toBeGreaterThan(0);
    }),
    { numRuns: 10 },
  );
});
