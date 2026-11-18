import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  Paragraph,
  Table,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { normalizeForOps } from "../contract";
import {
  compileEditorIntent,
  createEditorIntentIdAllocator,
} from "../editorIntent";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type EditorIntent,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const table = (paraId: string): Table => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [paragraph(paraId, "Cell")] }],
    },
  ],
});
const CONTAINERS = ["body", "cell", "sdt", "customXml"] as const;
type Container = (typeof CONTAINERS)[number];
const wrap = (
  content: BlockContent[],
  container: Container,
): BlockContent[] => {
  switch (container) {
    case "body":
      return content;
    case "cell":
      return [
        {
          type: "table",
          rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
        },
      ];
    case "sdt":
      return [
        {
          type: "blockSdt",
          properties: { sdtType: "richText", id: 77 },
          content,
        },
      ];
    case "customXml":
      return [
        {
          type: "blockCustomXml",
          openingXml: '<w:customXml w:element="clause">',
          closingXml: "</w:customXml>",
          content,
        },
      ];
    default: {
      const unreachable: never = container;
      return unreachable;
    }
  }
};
const SOURCE_IDS = ["10000001", "10000011", "10000022", "10000002"] as const;
const source = (container: Container): Document =>
  normalizeForOps({
    package: {
      document: {
        content: wrap(
          [
            paragraph("10000001", "Left"),
            table("10000011"),
            table("10000022"),
            paragraph("10000002", "Right"),
          ],
          container,
        ),
      },
    },
  });

/**
 * Every text a paragraph shows, pending insertions and deletions included,
 * leaving out the pending insertions whose ids `without` names.
 */
const textOf = (
  nodes: readonly unknown[],
  without: ReadonlySet<number> = new Set(),
): string =>
  nodes
    .map((node) => {
      if (typeof node !== "object" || node === null) return "";
      if (
        "type" in node &&
        node.type === "insertion" &&
        "info" in node &&
        typeof node.info === "object" &&
        node.info !== null &&
        "id" in node.info &&
        typeof node.info.id === "number" &&
        without.has(node.info.id)
      )
        return "";
      if ("text" in node && typeof node.text === "string") return node.text;
      if ("content" in node && Array.isArray(node.content))
        return textOf(node.content, without);
      return "";
    })
    .join("");
const paragraphsOf = (
  document: Document,
  without: ReadonlySet<number> = new Set(),
) =>
  storyParagraphs(document.package.document).map(({ paragraph: item }) => ({
    paraId: item.paraId,
    text: textOf(item.content, without),
  }));

const DATE = "2026-01-02T03:04:05Z";
/** The source table the pending deletion removes, and its revisions. */
const pendingDeletion = (document: Document) =>
  applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.DELETE_TABLE,
    story: OP_STORIES.MAIN,
    blockId: "10000022",
    revision: { id: 10, author: "Reviewer", date: DATE },
    newIds: { revision: Array.from({ length: 20 }, (_, index) => index + 11) },
  }).unwrap();

type Scenario = {
  container: Container;
  side: "beforeTable" | "afterTable";
  words: readonly string[];
  pending: boolean;
};

/**
 * A tracked table-side insertion, possibly beside a pending source table
 * deletion, with every revision id named by the change it belongs to.
 */
const producer = ({ container, side, words, pending }: Scenario) => {
  const clean = source(container);
  const deletion = pending ? pendingDeletion(clean) : undefined;
  const before = deletion?.document ?? clean;
  const intent = {
    type: "insertParagraphs",
    story: OP_STORIES.MAIN,
    at: {
      type: side,
      blockId: side === "beforeTable" ? "10000022" : "10000011",
    },
    paragraphs: words.map((word) => paragraph("00000000", word)),
  } as const satisfies EditorIntent;
  const ids = createEditorIntentIdAllocator()(before, intent);
  const compiled = compileEditorIntent(before, {
    intent,
    mode: {
      type: "suggesting",
      revision: { id: 100, author: "Reviewer", date: DATE },
      newIds: ids.newIds,
    },
  }).unwrap();
  const edit = applyDocumentOps(before, compiled.ops).unwrap();
  const direct = compileEditorIntent(clean, {
    intent,
    mode: { type: "editing", newIds: ids.newIds },
  }).unwrap();
  const insertedIds = paragraphsOf(edit.document)
    .map(({ paraId }) => paraId)
    .filter(
      (paraId): paraId is string =>
        paraId !== undefined &&
        !(SOURCE_IDS as readonly string[]).includes(paraId),
    );
  expect(insertedIds).toHaveLength(words.length);
  return {
    clean,
    edited: edit.document,
    direct,
    insertedIds,
    words,
    changes: [
      { change: "insertion" as const, ids: edit.revisions },
      ...(deletion === undefined
        ? []
        : [{ change: "deletion" as const, ids: deletion.revisions }]),
    ],
  };
};

type Produced = ReturnType<typeof producer>;

/** What review owes: the clean source with each change applied directly as decided. */
const expectedParagraphs = (
  produced: Produced,
  decisions: Readonly<Record<"insertion" | "deletion", RevisionDecision>>,
) => {
  let document = produced.clean;
  if (decisions.insertion === REVISION_DECISIONS.ACCEPT)
    document = applyDocumentOps(document, produced.direct.ops).unwrap()
      .document;
  if (
    produced.changes.some(({ change }) => change === "deletion") &&
    decisions.deletion === REVISION_DECISIONS.ACCEPT
  )
    document = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.DELETE_TABLE,
      story: OP_STORIES.MAIN,
      blockId: "10000022",
    }).unwrap().document;
  return paragraphsOf(document);
};

/**
 * Resolve revision ids one at a time in `order`, retrying a refused id after
 * the others once. Every step keeps each surviving paragraph's identity and
 * each source paragraph's own text: a rejected inserted mark may run the
 * insertion's pending content into the paragraph after it (a table's first
 * cell included) and an accepted source deletion may empty its cell, but the
 * insertion never takes source text and review never mints a paragraph.
 */
const review = (
  produced: Produced,
  steps: readonly { id: number; decision: RevisionDecision }[],
) => {
  const sourceText = new Map(
    paragraphsOf(produced.clean).map(({ paraId, text }) => [paraId, text]),
  );
  const insertedText = new Map(
    paragraphsOf(produced.edited)
      .filter(({ paraId }) => produced.insertedIds.includes(paraId ?? ""))
      .map(({ paraId, text }) => [paraId, text]),
  );
  const insertion = new Set(
    produced.changes.find(({ change }) => change === "insertion")?.ids,
  );
  const deletion = new Set(
    produced.changes.find(({ change }) => change === "deletion")?.ids,
  );
  // Generated words hold no pattern syntax.
  const insertedWords = new RegExp(produced.words.join("|"), "gu");
  let document = produced.edited;
  let queue = [...steps];
  const refused: typeof queue = [];
  for (let round = 0; round < 2 && queue.length > 0; round++) {
    for (const { id, decision } of queue) {
      const resolved = applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [id],
        decision,
      });
      if (resolved.isErr()) {
        expect(resolved.error.reason).toBe(
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        );
        refused.push({ id, decision });
        continue;
      }
      document = resolved.value.document;
      for (const { paraId, text } of paragraphsOf(document, insertion)) {
        if (!sourceText.has(paraId) && !insertedText.has(paraId))
          panic("Review minted a paragraph identity.", { paraId });
        const own = sourceText.get(paraId);
        if (own === undefined) continue;
        // Accepting the source deletion's text record may empty the deleted
        // table's cell while its row is still pending.
        const kept =
          paraId === "10000022" && deletion.size > 0 && !text.endsWith(own)
            ? ""
            : own;
        // Inserted words accepted while their mark is rejected run into the
        // following paragraph ahead of its own text; nothing else may.
        const runIn = text.slice(0, text.length - kept.length);
        expect({
          paraId,
          text,
          own: text.endsWith(kept),
          runIn: runIn.replaceAll(insertedWords, ""),
        }).toEqual({ paraId, text, own: true, runIn: "" });
      }
    }
    queue = refused.splice(0);
  }
  return { document, unresolved: queue };
};

const scenario = fc.record({
  container: fc.constantFrom(...CONTAINERS),
  side: fc.constantFrom("beforeTable" as const, "afterTable" as const),
  words: fc.array(fc.constantFrom("Inserted", "More", "ž😀"), {
    minLength: 1,
    maxLength: 3,
  }),
  pending: fc.boolean(),
});
const decision = fc.constantFrom(
  REVISION_DECISIONS.ACCEPT,
  REVISION_DECISIONS.REJECT,
);

test(
  "selective review of a table-side insertion keeps source identities and completes as decided",
  async () => {
    await assertProperty(
      fc.property(
        scenario,
        decision,
        decision,
        fc.infiniteStream(fc.nat()),
        (setting, insertion, deletion, shuffle) => {
          const produced = producer(setting);
          const decisions = { insertion, deletion };
          const steps = produced.changes
            .flatMap(({ change, ids }) =>
              ids.map((id) => ({ id, decision: decisions[change] })),
            )
            .map((step) => ({ step, key: shuffle.next().value }))
            .toSorted((left, right) => left.key - right.key)
            .map(({ step }) => step);
          const reviewed = review(produced, steps);
          // Decisions that agree within each change always complete.
          expect(reviewed.unresolved).toEqual([]);
          expect(paragraphsOf(reviewed.document)).toEqual(
            expectedParagraphs(produced, decisions),
          );
        },
      ),
      {
        id: "selective review of a table-side insertion keeps source identities and completes as decided",
        numRuns: 200,
      },
    );
  },
  propertyTestTimeout(60_000),
);

test(
  "selective review with split decisions keeps source identities or refuses unchanged",
  async () => {
    await assertProperty(
      fc.property(
        scenario,
        fc.infiniteStream(fc.tuple(fc.nat(), decision)),
        (setting, draws) => {
          const produced = producer(setting);
          const steps = produced.changes
            .flatMap(({ ids }) => ids)
            .map((id) => {
              const [key, chosen] = draws.next().value;
              return { key, step: { id, decision: chosen } };
            })
            .toSorted((left, right) => left.key - right.key)
            .map(({ step }) => step);
          // `review` holds every step to the identity laws; a split decision
          // may leave a refused id pending, but a refusal never reports
          // anything other than an untrackable review.
          const reviewed = review(produced, steps);
          for (const { id } of reviewed.unresolved)
            expect(steps.some((step) => step.id === id)).toBe(true);
        },
      ),
      {
        id: "selective review with split decisions keeps source identities or refuses unchanged",
        numRuns: 200,
      },
    );
  },
  propertyTestTimeout(60_000),
);

for (const side of ["beforeTable", "afterTable"] as const)
  test(`${side}: every ordering of single-id decisions over a pending source table completes as decided`, () => {
    for (const insertion of [
      REVISION_DECISIONS.ACCEPT,
      REVISION_DECISIONS.REJECT,
    ])
      for (const deletion of [
        REVISION_DECISIONS.ACCEPT,
        REVISION_DECISIONS.REJECT,
      ]) {
        const produced = producer({
          container: "body",
          side,
          words: ["Inserted"],
          pending: true,
        });
        const decisions = { insertion, deletion };
        const steps = produced.changes.flatMap(({ change, ids }) =>
          ids.map((id) => ({ id, decision: decisions[change] })),
        );
        for (const ordered of [steps, steps.toReversed()]) {
          const reviewed = review(produced, ordered);
          expect(reviewed.unresolved).toEqual([]);
          expect(paragraphsOf(reviewed.document)).toEqual(
            expectedParagraphs(produced, decisions),
          );
        }
      }
  });
