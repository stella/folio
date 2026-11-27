/**
 * A custom XML element's own tags, suggested: a pasted element records the
 * insertion of its tags on itself beside its content's records, a moved one
 * the deletion of its tags at the source, and accepting or rejecting the
 * element's record and its content's, in any order, leaves one document.
 */
import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  BlockCustomXml,
  Document,
  Paragraph,
  Run,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { validateOpsDocument } from "../contract";
import { compileEditorIntent, type EditorIntentMode } from "../editorIntent";
import { identityKeysIn } from "../ids";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});
const paragraph = (text: string, paraId?: string): Paragraph => ({
  type: "paragraph",
  ...(paraId === undefined ? {} : { paraId }),
  content: text === "" ? [] : [run(text)],
});

const PROPERTIES: BlockContent = {
  type: "preservedBlock",
  xml: '<w:customXmlPr><w:attr w:name="role" w:val="operative"/></w:customXmlPr>',
};

const element = (content: BlockContent[]): BlockCustomXml => ({
  type: "blockCustomXml",
  openingXml: '<w:customXml w:element="clause">',
  closingXml: "</w:customXml>",
  content,
});

const source = (): Document => ({
  package: {
    document: {
      content: [
        paragraph("before after", "00000001"),
        paragraph("next", "00000002"),
      ],
    },
  },
});

const SUGGESTING: EditorIntentMode = {
  type: "suggesting",
  revision: { id: 90, author: "R", date: "2026-01-02T00:00:00Z" },
  newIds: { revision: Array.from({ length: 80 }, (_, index) => 91 + index) },
};
const DIRECT: EditorIntentMode = { type: "editing" };

const apply = (document: Document, ops: readonly DocumentOp[]) => {
  const applied = applyDocumentOps(document, ops);
  if (applied.isErr()) throw applied.error;
  return applied.value;
};

const paste = (blocks: BlockContent[], mode: EditorIntentMode) => {
  const document = source();
  const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 7 };
  const compiled = compileEditorIntent(document, {
    intent: {
      type: "pasteBlocks",
      from: at,
      to: at,
      blocks: [
        paragraph("Lead"),
        element([PROPERTIES, ...blocks]),
        paragraph("Trail"),
      ],
      openStart: 1,
      openEnd: 1,
    },
    mode,
  });
  if (compiled.isErr()) throw compiled.error;
  return { document, applied: apply(document, compiled.value.ops) };
};

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
) =>
  apply(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds,
      decision,
    },
  ]).document;

/** The document as a reader sees it: text by paragraph, wrappers as nesting. */
const shapeOf = (blocks: readonly BlockContent[]): unknown[] =>
  blocks.flatMap((block): unknown[] => {
    if (block.type === "blockCustomXml")
      return [
        {
          element: shapeOf(block.content),
          ...(block.insertion === undefined ? {} : { inserted: true }),
          ...(block.deletion === undefined ? {} : { deleted: true }),
        },
      ];
    if (block.type === "preservedBlock") return ["properties"];
    if (block.type !== "paragraph") return [block.type];
    return [
      block.content
        .flatMap((item) => {
          let runs: Run[] = [];
          let prefix = "";
          if (item.type === "run") runs = [item];
          else if (item.type === "insertion" || item.type === "deletion") {
            runs = item.content.flatMap((inner) =>
              inner.type === "run" ? [inner] : [],
            );
            prefix = item.type === "insertion" ? "+" : "-";
          }
          return runs.flatMap((one) =>
            one.content.flatMap((leaf) =>
              leaf.type === "text" ? [prefix + leaf.text] : [],
            ),
          );
        })
        .join(""),
    ];
  });

const revisionIdsOf = (document: Document): number[] =>
  [
    ...new Set(
      identityKeysIn(document.package.document).flatMap((key) =>
        key.startsWith("revision:") ? [Number(key.slice(9))] : [],
      ),
    ),
  ].toSorted((left, right) => left - right);

const elementIn = (document: Document): BlockCustomXml => {
  const found = document.package.document.content.find(
    (block): block is BlockCustomXml => block.type === "blockCustomXml",
  );
  if (found === undefined) throw new TypeError("No custom XML element.");
  return found;
};

type Decision = RevisionDecision;
const DECISIONS: readonly Decision[] = [
  REVISION_DECISIONS.ACCEPT,
  REVISION_DECISIONS.REJECT,
];

/** Every order of `items`. */
const orders = <Item>(items: readonly Item[]): Item[][] =>
  items.length <= 1
    ? [[...items]]
    : items.flatMap((item, index) =>
        orders([...items.slice(0, index), ...items.slice(index + 1)]).map(
          (rest) => [item].concat(rest),
        ),
      );

/**
 * What resolving each group with its decision leaves, in every order, and
 * resolving groups that share a decision at once: one shape, or the first
 * that differs.
 */
const resolvedShapes = (
  document: Document,
  groups: readonly (readonly number[])[],
  decisions: readonly Decision[],
): Set<string> => {
  const outcomes = new Set<string>();
  const steps = groups.map((ids, index): [readonly number[], Decision] => [
    ids,
    decisions[index] ?? REVISION_DECISIONS.ACCEPT,
  ]);
  for (const order of orders(steps)) {
    let current = document;
    for (const [ids, decision] of order)
      current = resolve(current, ids, decision);
    expect(validateOpsDocument(current).isOk()).toBe(true);
    outcomes.add(JSON.stringify(shapeOf(current.package.document.content)));
  }
  for (const decision of DECISIONS) {
    const together = steps.filter(([, each]) => each === decision);
    if (together.length !== steps.length) continue;
    outcomes.add(
      JSON.stringify(
        shapeOf(
          resolve(
            document,
            together.flatMap(([ids]) => ids),
            decision,
          ).package.document.content,
        ),
      ),
    );
  }
  return outcomes;
};

describe("a suggested paste of a custom XML element", () => {
  const inner = [paragraph("Inside one"), paragraph("Inside two")];

  test("records the insertion of its tags on the element, reported with the paste's", () => {
    const { applied } = paste(inner, SUGGESTING);
    expect(validateOpsDocument(applied.document).isOk()).toBe(true);
    const pasted = elementIn(applied.document);
    expect(pasted.insertion?.info).toEqual({
      id: pasted.insertion?.info.id ?? -1,
      author: "R",
      date: "2026-01-02T00:00:00Z",
    });
    expect(applied.revisions).toContain(pasted.insertion?.info.id ?? -1);
    expect(applied.revisions.toSorted((a, b) => a - b)).toEqual(
      revisionIdsOf(applied.document),
    );
    expect(shapeOf(applied.document.package.document.content)).toEqual([
      "before +Lead",
      { element: ["properties", "+Inside one", "+Inside two"], inserted: true },
      "+Trailafter",
      "next",
    ]);
  });

  test("undoes to the document it was pasted into", () => {
    const { document, applied } = paste(inner, SUGGESTING);
    expect(apply(applied.document, applied.inverse).document).toEqual(document);
  });

  test.each([
    {
      element: REVISION_DECISIONS.ACCEPT,
      content: REVISION_DECISIONS.ACCEPT,
      shape: [
        "before Lead",
        { element: ["properties", "Inside one", "Inside two"] },
        "Trailafter",
        "next",
      ],
    },
    {
      element: REVISION_DECISIONS.REJECT,
      content: REVISION_DECISIONS.REJECT,
      shape: ["before after", "next"],
    },
    {
      element: REVISION_DECISIONS.REJECT,
      content: REVISION_DECISIONS.ACCEPT,
      shape: ["before Lead", "Inside one", "Inside two", "Trailafter", "next"],
    },
    // What the tags wrapped joins across them, and with all of it gone the
    // tags go too: no empty custom XML element is left where the
    // suggestion added none.
    {
      element: REVISION_DECISIONS.ACCEPT,
      content: REVISION_DECISIONS.REJECT,
      shape: ["before after", "next"],
    },
  ])(
    "resolves to one document whichever goes first: tags $element, content $content",
    ({ element: tags, content, shape }) => {
      const { applied } = paste(inner, SUGGESTING);
      const own = elementIn(applied.document).insertion?.info.id ?? -1;
      const rest = revisionIdsOf(applied.document).filter((id) => id !== own);
      expect([
        ...resolvedShapes(applied.document, [[own], rest], [tags, content]),
      ]).toEqual([JSON.stringify(shape)]);
    },
  );

  test("accepted, leaves what the direct paste leaves", () => {
    const suggested = paste(inner, SUGGESTING).applied.document;
    const direct = paste(inner, DIRECT).applied.document;
    expect(
      shapeOf(
        resolve(suggested, revisionIdsOf(suggested), REVISION_DECISIONS.ACCEPT)
          .package.document.content,
      ),
    ).toEqual(shapeOf(direct.package.document.content));
  });

  test("drops the records a copied element carries from its own document", () => {
    const copied: BlockCustomXml = {
      ...element([PROPERTIES, paragraph("Copied")]),
      insertion: { info: { id: 7, author: "Elsewhere" } },
      deletion: { info: { id: 8, author: "Elsewhere" } },
    };
    const document = source();
    const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 };
    const compiled = compileEditorIntent(document, {
      intent: {
        type: "pasteBlocks",
        from: at,
        to: at,
        blocks: [copied],
        openStart: 0,
        openEnd: 0,
      },
      mode: DIRECT,
    });
    if (compiled.isErr()) throw compiled.error;
    const pasted = elementIn(apply(document, compiled.value.ops).document);
    expect(pasted.insertion).toBeUndefined();
    expect(pasted.deletion).toBeUndefined();
  });

  test.each([
    {
      kind: "a block content control",
      held: {
        type: "blockSdt",
        properties: { sdtType: "richText", id: 8 },
        content: [paragraph("Nested")],
      } satisfies BlockContent,
    },
    {
      kind: "kept block markup",
      held: {
        type: "preservedBlock",
        xml: "<w:altChunk/>",
      } satisfies BlockContent,
    },
  ])("refuses, as the format's rule, an element holding $kind", ({ held }) => {
    const document = source();
    const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 7 };
    const compiled = compileEditorIntent(document, {
      intent: {
        type: "pasteBlocks",
        from: at,
        to: at,
        blocks: [element([PROPERTIES, held, paragraph("Beside")])],
        openStart: 0,
        openEnd: 0,
      },
      mode: SUGGESTING,
    });
    expect(compiled.isErr()).toBe(true);
    if (compiled.isErr())
      expect(compiled.error.reason).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.NO_REVISION_RECORD,
      );
  });
});

describe("a suggested move of a custom XML element", () => {
  const document = (): Document => ({
    package: {
      document: {
        content: [
          paragraph("zero", "00000001"),
          paragraph("one", "00000002"),
          element([PROPERTIES, paragraph("held", "00000005")]),
          paragraph("two", "00000003"),
          paragraph("three", "00000004"),
        ],
      },
    },
  });
  const at = (blockId: string, offset: number) => ({
    story: OP_STORIES.MAIN,
    blockId,
    offset,
  });
  const drag = (mode: EditorIntentMode) => {
    const dragged = document();
    const compiled = compileEditorIntent(dragged, {
      intent: {
        type: "moveBlocks",
        from: at("00000002", 3),
        to: at("00000003", 0),
        target: at("00000004", 2),
        blocks: [
          paragraph(""),
          element([PROPERTIES, paragraph("held")]),
          paragraph(""),
        ],
        openStart: 1,
        openEnd: 1,
      },
      mode,
    });
    if (compiled.isErr()) throw compiled.error;
    return { dragged, applied: apply(dragged, compiled.value.ops) };
  };
  const ORIGINAL = [
    "zero",
    "one",
    { element: ["properties", "held"] },
    "two",
    "three",
  ];
  const MOVED = [
    "zero",
    "onetwo",
    "th",
    { element: ["properties", "held"] },
    "ree",
  ];

  test("records the deletion of its tags at the source and their insertion at the target", () => {
    const { applied } = drag(SUGGESTING);
    expect(validateOpsDocument(applied.document).isOk()).toBe(true);
    expect(shapeOf(applied.document.package.document.content)).toEqual([
      "zero",
      "one",
      { element: ["properties", "-held"], deleted: true },
      "two",
      "th",
      { element: ["properties", "+held"], inserted: true },
      "ree",
    ]);
    expect(applied.revisions.toSorted((a, b) => a - b)).toEqual(
      revisionIdsOf(applied.document),
    );
    expect(
      shapeOf(drag(DIRECT).applied.document.package.document.content),
    ).toEqual(MOVED);
  });

  test("resolves to one document in every order, the direct move accepted and the original rejected", () => {
    const { applied } = drag(SUGGESTING);
    const elements = applied.document.package.document.content.filter(
      (block): block is BlockCustomXml => block.type === "blockCustomXml",
    );
    const moved = elements.at(0)?.deletion?.info.id ?? -1;
    const target = elements.at(1)?.insertion?.info.id ?? -1;
    const rest = revisionIdsOf(applied.document).filter(
      (id) => id !== moved && id !== target,
    );
    for (const first of DECISIONS)
      for (const second of DECISIONS)
        for (const third of DECISIONS) {
          const outcomes = resolvedShapes(
            applied.document,
            [[moved], [target], rest],
            [first, second, third],
          );
          expect({ first, second, third, outcomes: outcomes.size }).toEqual({
            first,
            second,
            third,
            outcomes: 1,
          });
          if (first === second && second === third)
            expect([...outcomes]).toEqual([
              JSON.stringify(
                first === REVISION_DECISIONS.ACCEPT ? MOVED : ORIGINAL,
              ),
            ]);
        }
  });
});
