/**
 * What a run's identity does not extend to, and what carries it where.
 *
 * `inclusive: false` settles the two edges of a marked span. It says nothing
 * about the interior: ProseMirror's stored marks at a position inside a span
 * are those of the preceding character, so a character typed in the middle of
 * an authored run inherits that run's `w:rsidR` — a statement that the editing
 * session named there wrote it.
 *
 * folio may not make that statement. An rsid names a session registered in the
 * package's own `settings.xml`, `serializeSettingsXml` is reachable only from
 * the fresh-document builder, and every repack copies `settings.xml` through,
 * so folio can neither register a session nor merge two documents' tables. It
 * therefore writes no rsid it did not read, and text it inserts becomes a run
 * that states none — which `w:rsidRDefault` on the parent `w:p` already
 * answers for, on 94% of authored paragraphs.
 *
 * The same argument closes the clipboard: a pasted span's source session is
 * not in this package's table, so `toDOM` carries the id and nothing else.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { Window } from "happy-dom";
import { DOMParser, DOMSerializer, type Mark } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

import { propertyConfig } from "../../../../test/property-testing";

import type { Paragraph, Run } from "../types/document";
import { proseDocToBlocks } from "./conversion/fromProseDoc";
import { toProseDoc } from "./conversion/toProseDoc";
import { createStarterKit } from "./extensions/StarterKit";
import { ExtensionManager } from "./extensions/ExtensionManager";
import { RUN_IDENTITY_ATTRIBUTE, runIdentityAttrs } from "./runIdentity";
import { schema } from "./schema";

const RSID = "00ABCDEF";
const AUTHORED = "authored";

/** One paragraph, one run, one remainder: the shape every case here starts in. */
const documentWithOneAuthoredRun = () => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [{ type: "text", text: AUTHORED }],
              preservedAttributes: [{ name: "rsidR", value: RSID }],
            } satisfies Run,
          ],
        } satisfies Paragraph,
      ],
    },
  },
});

/**
 * An editor state carrying the plugins the real editor runs, so the strip is
 * exercised through the hook that sees every insertion rather than through a
 * helper written for the test.
 */
const editorStateOf = (doc: ReturnType<typeof toProseDoc>): EditorState => {
  const manager = new ExtensionManager(createStarterKit());
  manager.buildSchema();
  manager.initializeRuntime();
  return EditorState.create({ doc, plugins: manager.getPlugins() });
};

// SAFETY: the fixture states exactly this shape; `toProseDoc` accepts it.
const authoredDocument = () => toProseDoc(documentWithOneAuthoredRun() as never);

/** The saved runs of the one paragraph, as `(text, carries the rsid)` pairs. */
const savedRuns = (state: EditorState): [string, boolean][] => {
  const [block] = proseDocToBlocks(state.doc);
  if (block?.type !== "paragraph") {
    throw new Error("the state did not save as one paragraph");
  }
  return block.content
    .filter((item): item is Run => item.type === "run")
    .map((run) => [
      run.content.map((item) => (item.type === "text" ? item.text : "")).join(""),
      run.preservedAttributes?.some(({ value }) => value === RSID) ?? false,
    ]);
};

/** The position of the authored text's first character. */
const authoredStart = (state: EditorState): number => {
  let found: number | undefined;
  state.doc.descendants((node, pos) => {
    if (found === undefined && node.isText && node.text === AUTHORED) {
      found = pos;
    }
    return found === undefined;
  });
  if (found === undefined) {
    throw new Error("the projection produced no authored text");
  }
  return found;
};

describe("text the editor inserts claims no editing session", () => {
  test("at every offset of an authored run, including both edges", () => {
    fc.assert(
      fc.property(fc.integer({ min: 0, max: AUTHORED.length }), (offset) => {
        const before = editorStateOf(authoredDocument());
        const at = authoredStart(before) + offset;
        const typed = before.apply(
          before.tr.setSelection(TextSelection.create(before.doc, at)).insertText("X", at),
        );

        // Whatever the cut, the inserted character is a run of its own that
        // states no session, and each authored piece is a run that states the
        // session it was authored in. Stated as the whole partition: a weaker
        // assertion passes when the three pieces coalesce back into one run
        // carrying the authored rsid, which is the defect.
        expect(savedRuns(typed)).toEqual(
          (
            [
              [AUTHORED.slice(0, offset), true],
              ["X", false],
              [AUTHORED.slice(offset), true],
            ] as [string, boolean][]
          ).filter(([text]) => text.length > 0),
        );
      }),
      propertyConfig({ numRuns: AUTHORED.length + 1 }),
    );
  });

  test("a typing burst coalesces into one run rather than one run per keystroke", () => {
    let state = editorStateOf(authoredDocument());
    const at = authoredStart(state) + 3;
    for (const [index, character] of [..."new"].entries()) {
      state = state.apply(state.tr.insertText(character, at + index));
    }

    expect(savedRuns(state)).toEqual([
      ["aut", true],
      ["new", false],
      ["hored", true],
    ]);
  });
});

// ---------------------------------------------------------------------------

const window = new Window();
const document = window.document as unknown as globalThis.Document;

/** The mark set of the single leaf in a paragraph serialized and parsed back. */
const throughTheDom = (marks: readonly Mark[]): readonly Mark[] => {
  const host = document.createElement("div");
  host.append(
    DOMSerializer.fromSchema(schema).serializeFragment(
      schema.node("paragraph", undefined, [schema.text("x", [...marks])]).content,
      { document },
    ),
  );
  const parsed = DOMParser.fromSchema(schema).parse(host);
  let found: readonly Mark[] = [];
  parsed.descendants((node) => {
    if (node.isText) {
      found = node.marks;
    }
    return true;
  });
  return found;
};

/** The save leg's own grouping key, spelled as `getMarksKey` spells it. */
const marksKey = (marks: readonly Mark[]): string =>
  marks
    .filter(({ type }) => type.name !== "hyperlink")
    .map((mark) => `${mark.type.name}:${JSON.stringify(mark.attrs)}`)
    .toSorted()
    .join("|");

const payloadArbitrary = fc.record(
  {
    preservedAttributes: fc.array(
      fc.record(
        {
          namespace: fc.constantFrom(
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            undefined,
          ),
          name: fc.constantFrom("rsidR", "rsidRPr", "rsidDel"),
          value: fc.string({
            minLength: 8,
            maxLength: 8,
            unit: fc.constantFrom(..."0123456789ABCDEF"),
          }),
        },
        { requiredKeys: ["name", "value"] },
      ),
      { maxLength: 3 },
    ),
    preserved: fc.record({
      children: fc.array(
        fc.record({ index: fc.nat({ max: 4 }), xml: fc.constantFrom("<w:u/>", "<w:bdr/>") }),
        { maxLength: 2 },
      ),
    }),
  },
  { requiredKeys: [] },
);

describe("the identity's key", () => {
  test("a mark built by the factory keys the same however its payload was spelled", () => {
    fc.assert(
      fc.property(fc.nat({ max: 1000 }), payloadArbitrary, (id, payload) => {
        const direct = schema.mark("runIdentity", runIdentityAttrs(id, payload));
        // The same payload with its keys in the other order: what `JSON.parse`
        // hands back from a stored attrs object, and what the key class is.
        const reordered = schema.mark(
          "runIdentity",
          runIdentityAttrs(id, {
            preserved: payload.preserved,
            preservedAttributes: payload.preservedAttributes?.map(({ value, name, namespace }) =>
              namespace === undefined ? { value, name } : { value, name, namespace },
            ),
          }),
        );

        expect(marksKey([reordered])).toBe(marksKey([direct]));
      }),
      propertyConfig({ numRuns: 50 }),
    );
  });

  test("a pasted span keys differently from its source, because the payload is not in the DOM", () => {
    fc.assert(
      fc.property(
        fc.nat({ max: 1000 }),
        payloadArbitrary.filter(
          ({ preservedAttributes, preserved }) =>
            (preservedAttributes?.length ?? 0) > 0 || (preserved?.children?.length ?? 0) > 0,
        ),
        (id, payload) => {
          const source = schema.mark("runIdentity", runIdentityAttrs(id, payload));
          const pasted = throughTheDom([source]);

          // The id travels, so a copied page-break-bearing run still rejoins.
          expect(pasted.map(({ type }) => type.name)).toEqual(["runIdentity"]);
          expect(pasted[0]?.attrs["id"]).toBe(id);
          // The payload does not, so the paste is its own run with no
          // attributes — the same answer typing gets, by a second route.
          expect(marksKey(pasted)).not.toBe(marksKey([source]));
          expect(marksKey(pasted)).toBe(
            marksKey([schema.mark("runIdentity", runIdentityAttrs(id))]),
          );
        },
      ),
      propertyConfig({ numRuns: 50 }),
    );
  });

  test("the DOM spelling is the id alone", () => {
    const host = document.createElement("div");
    host.append(
      DOMSerializer.fromSchema(schema).serializeFragment(
        schema.node("paragraph", undefined, [
          schema.text("x", [
            schema.mark(
              "runIdentity",
              runIdentityAttrs(9, { preservedAttributes: [{ name: "rsidR", value: RSID }] }),
            ),
          ]),
        ]).content,
        { document },
      ),
    );
    const span = host.querySelector(`span[${RUN_IDENTITY_ATTRIBUTE}]`);

    expect(span?.getAttribute(RUN_IDENTITY_ATTRIBUTE)).toBe("9");
    expect(span?.outerHTML).not.toContain(RSID);
  });
});
