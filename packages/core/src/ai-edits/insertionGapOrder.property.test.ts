/**
 * Blocks a reviewer puts beside one anchor land in the order they were
 * requested, applied directly, as tracked changes accepted, and as
 * suggestions accepted and then accepted as tracked changes, whatever else
 * the batch does there.
 *
 * The anchor may be the story's last paragraph, or have a paragraph after
 * it, with a table between the two that stays or that a first batch deletes
 * (as a tracked change still pending, unless applied directly); or the
 * paragraph after it ends the story and a first batch deletes it, so the
 * anchor's break is pending removal. Every batch puts one to five blocks of
 * mixed kinds (a paragraph, two paragraphs from one text, a paragraph
 * starting a new list, a table or a signature table, whichever the mode has;
 * no operation inserts a block content control) before or after the anchor,
 * after the table, or before the next paragraph. A paragraph asks for a
 * heading style, an alignment or nothing. The last batch may delete the
 * anchor too. Between batches the reviewer may save and reopen the package,
 * accept every change, or both.
 *
 * A block put after a block lands right after it; one put before a block
 * lands before the blocks right before it that accepting removes whole. The
 * blocks a batch puts at one place keep the batch's order (`requestedOrder`),
 * and each keeps the properties it asked for: one that asked for none reads
 * as the anchor, never as another block of the batch.
 *
 * A case saves and reopens a package a few times, about 50 ms; the budget is
 * 120 s for the stated runs, scaled by the nightly factor.
 */

import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

type Mode = "direct" | "tracked-changes" | "suggested";
type Kind = "paragraph" | "paragraphs" | "newList" | "table" | "signature";
type Side = "beforeAnchor" | "afterAnchor" | "afterTable" | "beforeNext";
type Between = "nothing" | "reopen" | "accept" | "reopenAndAccept";
/** What a paragraph asks for: a heading style, an alignment, or nothing. */
type Look = "plain" | "heading" | "centered";
type Insertion = { kind: Kind; side: Side; look: Look };
type Batch = { insertions: Insertion[]; deletesAnchor: boolean };
type Case = {
  anchorIsLast: boolean;
  /** A table between the anchor and the next paragraph, and what happens to it. */
  table: "none" | "kept" | "removed";
  /** The next paragraph ends the story, and a first batch deletes it. */
  nextRemoved: boolean;
  mode: Exclude<Mode, "direct">;
  batches: Batch[];
  between: Between[];
};

/** A block once accepted: its words, whether a table holds it, and how it looks. */
type Block = {
  text: string;
  inTable: boolean;
  styleId: string | undefined;
  alignment: string | undefined;
};
/** What the requested order names: an insertion, or a block that was there. */
type Placed = { label: string; table: boolean };

const OPENING = "Opening clause.";
/** The style the opening heading defines, which a paragraph may ask for. */
const HEADING = "Heading2";
const ANCHOR = "Anchor clause.";
const NEXT = "Next clause.";
const TABLE = "TBL";
const LABEL = /\b(?:B\dI\d|TBL)\b/u;

/**
 * Who put `block` there: the insertion its words name, or the paragraph it
 * is. A signature table's other cells (rules, captions) belong to the party
 * named before them.
 */
const labelsOf = (blocks: readonly Block[]): string[] => {
  const labels: string[] = [];
  for (const block of blocks) {
    const label =
      LABEL.exec(block.text)?.[0] ??
      (block.inTable ? labels.at(-1) : block.text);
    if (label !== undefined && label !== labels.at(-1)) labels.push(label);
  }
  return labels;
};

const looking = (look: Look) => {
  switch (look) {
    case "plain":
      return {};
    case "heading":
      return { styleId: HEADING };
    case "centered":
      return { alignment: "center" } as const;
  }
};

const insertionOf = (
  { kind, look }: Insertion,
  label: string,
  anchor: { blockId: string; after: boolean },
): FolioDocumentOperation => {
  const id = `insert-${label}`;
  const position = anchor.after ? "after" : "before";
  const beside = {
    type: anchor.after ? "insertAfterBlock" : "insertBeforeBlock",
    blockId: anchor.blockId,
  } as const;
  switch (kind) {
    case "paragraph":
      return { id, ...beside, text: `${label} alone.`, ...looking(look) };
    case "paragraphs":
      return {
        id,
        ...beside,
        text: `${label} first.\n${label} second.`,
        ...looking(look),
        // Both paragraphs ask for the look, so both are checked for it.
        formattingScope: "allParagraphs",
      };
    case "newList":
      return {
        id,
        ...beside,
        text: `${label} listed.`,
        numbering: { kind: "newList", format: "numbered" },
      };
    case "table":
      return {
        id,
        type: "insertTable",
        blockId: anchor.blockId,
        position,
        rows: [[`${label} term`, `${label} value`]],
      };
    case "signature":
      return {
        id,
        type: "insertSignatureTable",
        blockId: anchor.blockId,
        position,
        parties: [{ name: `${label} signer` }],
      };
  }
};

/** What stands between the anchor and the next paragraph, in order. */
type Stretch = (Placed & { isTable?: true })[];

/**
 * Each inserted paragraph reads as it asked: in the heading style when it
 * asked for it, centered when it asked for that, and otherwise as the plain
 * anchor, whatever the other blocks of its batch asked for.
 */
const looksOf = (blocks: readonly Block[], looks: ReadonlyMap<string, Look>) =>
  blocks.flatMap((block) => {
    const named = LABEL.exec(block.text)?.[0];
    const look = named === undefined ? undefined : looks.get(named);
    if (look === undefined || block.inTable) return [];
    return [
      {
        label: named,
        asked: look,
        heading: block.styleId === HEADING,
        centered: block.alignment === "center",
      },
    ];
  });

/**
 * The places a batch puts `insertions` at, read before it runs, then each
 * place's blocks in the batch's order, right to left. After the anchor is
 * right after it, after the table right after the table; before the next
 * paragraph is before the table right before it when accepting removes it.
 */
const requestedOrder = (
  stretch: Stretch,
  insertions: readonly (Placed & { side: Side })[],
  tableGoes: boolean,
): Stretch => {
  const tableAt = stretch.findIndex((entry) => entry.isTable === true);
  const placeOf = (side: Side): number => {
    if (side === "afterAnchor") return 0;
    if (side === "afterTable") return tableAt + 1;
    let at = stretch.length;
    if (tableGoes && at > 0 && stretch[at - 1]?.isTable) at -= 1;
    return at;
  };
  const places = new Map<number, Placed[]>();
  for (const { side, ...placed } of insertions) {
    const place = placeOf(side);
    places.set(place, [...(places.get(place) ?? []), placed]);
  }
  const result = [...stretch];
  for (const place of [...places.keys()].toSorted((a, b) => b - a)) {
    result.splice(place, 0, ...(places.get(place) ?? []));
  }
  return result;
};

const open = async (setup: Case): Promise<FolioDocxReviewer> => {
  const texts = [
    `## ${OPENING}`,
    ANCHOR,
    ...(setup.table === "none"
      ? []
      : [
          `| ${TABLE} one | ${TABLE} two |\n| --- | --- |\n| ${TABLE} three | ${TABLE} four |`,
        ]),
    ...(setup.anchorIsLast ? [] : [NEXT]),
  ];
  const bytes = (
    await ensureParaIds(await createDocx(fromMarkdown(texts.join("\n\n"))))
  ).docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });
};

const reopen = async (
  reviewer: FolioDocxReviewer,
): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
    author: "Reviewer",
  });

/** The blocks a reader lists once every change is accepted and saved. */
const acceptedBlocks = async (
  reviewer: FolioDocxReviewer,
): Promise<Block[]> => {
  reviewer.acceptAll();
  return (await reopen(reviewer))
    .snapshot()
    .blocks.filter((block) => block.text.length > 0)
    .map((block) => ({
      text: block.text,
      inTable: block.table !== undefined,
      styleId: block.styleId ?? undefined,
      alignment: block.directAlignment ?? undefined,
    }));
};

/** Apply `operations` in `mode`, accepting what it suggests as tracked changes. */
const applyBatch = (
  reviewer: FolioDocxReviewer,
  mode: Mode,
  operations: FolioDocumentOperation[],
) => {
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      atomic: false,
      operations,
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  // A suggestion stays out of the package until accepted, which turns it
  // into an ordinary tracked change.
  const suggestionIds = new Set(
    result.applied.flatMap(({ suggestionId }) => suggestionId ?? []),
  );
  for (const suggestionId of suggestionIds) {
    expect(reviewer.acceptSuggestion(suggestionId)).toBe(true);
  }
};

/**
 * Replay `setup` in `mode`: the blocks accepting leaves, and the blocks the
 * requested order says.
 */
const replay = async (setup: Case, mode: Mode) => {
  let reviewer = await open(setup);
  const idOf = (text: string): string => {
    const block = reviewer.snapshot().blocks.find((b) => b.text === text);
    if (!block) throw new Error(`No block reads "${text}"`);
    return block.id;
  };
  const anchorId = idOf(ANCHOR);
  const nextId = setup.anchorIsLast ? undefined : idOf(NEXT);
  const tableId = setup.table === "none" ? undefined : idOf(`${TABLE} one`);
  let stretch: Stretch =
    setup.table === "none"
      ? []
      : [{ label: TABLE, table: true, isTable: true }];
  if (setup.nextRemoved && nextId !== undefined) {
    // The story's last paragraph, deleted: the anchor's break goes instead.
    applyBatch(reviewer, mode === "direct" ? "direct" : "tracked-changes", [
      { id: "delete-next", type: "deleteBlock", blockId: nextId },
    ]);
  }
  if (setup.table === "removed" && tableId !== undefined) {
    // Another reviewer's pending deletion, unless applied directly.
    applyBatch(reviewer, mode === "direct" ? "direct" : "tracked-changes", [
      { id: "delete-table", type: "deleteTable", blockId: tableId },
    ]);
  }
  const anchors: Record<Side, string | undefined> = {
    beforeAnchor: anchorId,
    afterAnchor: anchorId,
    afterTable: tableId,
    beforeNext: nextId,
  };
  const anchorOf = (side: Side) => {
    const blockId = anchors[side];
    if (blockId === undefined) throw new Error(`No block for ${side}`);
    return { blockId, after: side !== "beforeNext" && side !== "beforeAnchor" };
  };
  /** What stands right before the anchor, after the opening, in order. */
  const ahead: Placed[] = [];
  const looks = new Map<string, Look>();
  for (const [index, batch] of setup.batches.entries()) {
    const placed = batch.insertions.map(({ kind, side }, at) => ({
      label: `B${index}I${at}`,
      table: kind === "table" || kind === "signature",
      side,
    }));
    for (const [at, insertion] of batch.insertions.entries()) {
      if (insertion.kind === "paragraph" || insertion.kind === "paragraphs")
        looks.set(`B${index}I${at}`, insertion.look);
    }
    const operations = batch.insertions.map((insertion, at) =>
      insertionOf(insertion, `B${index}I${at}`, anchorOf(insertion.side)),
    );
    if (batch.deletesAnchor) {
      operations.push({ id: "delete", type: "deleteBlock", blockId: anchorId });
    }
    applyBatch(reviewer, mode, operations);
    // Before the anchor is one gap, after what earlier batches put there.
    for (const { side, label: name, table } of placed) {
      if (side === "beforeAnchor") ahead.push({ label: name, table });
    }
    stretch = requestedOrder(
      stretch,
      placed.filter(({ side }) => side !== "beforeAnchor"),
      setup.table === "removed",
    );
    const step = setup.between[index] ?? "nothing";
    if (step === "reopen" || step === "reopenAndAccept") {
      reviewer = await reopen(reviewer);
    }
    if (step === "accept" || step === "reopenAndAccept") {
      reviewer.acceptAll();
    }
  }
  const deleted = setup.batches.some((batch) => batch.deletesAnchor);
  const accepted = await acceptedBlocks(reviewer);
  return {
    accepted,
    looks: looksOf(accepted, looks),
    /** What stands between the anchor and the next paragraph while pending. */
    pending: stretch.map(({ label, table }) => ({ label, table })),
    /** What stands before the anchor, after the opening. */
    ahead,
    requested: [
      { label: OPENING, table: false },
      ...ahead,
      ...(deleted ? [] : [{ label: ANCHOR, table: false }]),
      ...stretch
        .filter((entry) => !(entry.isTable && setup.table === "removed"))
        .map(({ label, table }) => ({ label, table })),
      ...(setup.anchorIsLast || setup.nextRemoved
        ? []
        : [{ label: NEXT, table: false }]),
    ],
  };
};

const kindsIn = (mode: Case["mode"]) =>
  fc.constantFrom<Kind[]>(
    "paragraph",
    "paragraphs",
    "newList",
    // A table has a tracked form and a signature table a suggested one.
    mode === "suggested" ? "signature" : "table",
  );

const caseArbitrary: fc.Arbitrary<Case> = fc
  .constantFrom<Case["mode"][]>("tracked-changes", "suggested")
  .chain((mode) =>
    // A suggestion beside the story's last paragraph fails to accept, or goes
    // with a suggested deletion of that paragraph, which is not about order:
    // suggestions keep a paragraph after the anchor. A table stands before
    // a paragraph.
    (mode === "suggested" ? fc.constant(false) : fc.boolean()).chain(
      (anchorIsLast) =>
        fc
          .record({
            anchorIsLast: fc.constant(anchorIsLast),
            table: anchorIsLast
              ? fc.constant<Case["table"]>("none")
              : fc.constantFrom<Case["table"][]>("none", "kept", "removed"),
            mode: fc.constant(mode),
          })
          .chain((setup) =>
            fc.record({
              anchorIsLast: fc.constant(setup.anchorIsLast),
              table: fc.constant(setup.table),
              mode: fc.constant(setup.mode),
              // Tracked only: a suggestion beside what then ends the story
              // fails to accept, as above.
              nextRemoved:
                setup.anchorIsLast ||
                setup.table !== "none" ||
                setup.mode === "suggested"
                  ? fc.constant(false)
                  : fc.boolean(),
            }),
          ),
    ),
  )
  .chain(({ anchorIsLast, table, mode, nextRemoved }) => {
    // A suggestion is refused beside a block pending deletion.
    const besideTable =
      table === "kept" || (table === "removed" && mode !== "suggested");
    const sides: Side[] = [
      "beforeAnchor",
      "afterAnchor",
      ...(besideTable ? (["afterTable"] as const) : []),
      ...(anchorIsLast || nextRemoved ? [] : (["beforeNext"] as const)),
    ];
    return fc
      .array(
        fc.array(
          fc.record({
            kind: kindsIn(mode),
            side: fc.constantFrom(...sides),
            look: fc.constantFrom<Look[]>("plain", "heading", "centered"),
          }),
          { minLength: 1, maxLength: 5 },
        ),
        { minLength: 1, maxLength: 3 },
      )
      .chain((batches) =>
        fc.record({
          anchorIsLast: fc.constant(anchorIsLast),
          table: fc.constant(table),
          mode: fc.constant(mode),
          nextRemoved: fc.constant(nextRemoved),
          // The anchor's own break pending removal is not deleted again.
          batches: (nextRemoved ? fc.constant(false) : fc.boolean()).map(
            (deletes) =>
              batches.map((insertions, index) => {
                const deletesAnchor = deletes && index === batches.length - 1;
                return {
                  // A suggestion is refused beside a block pending deletion.
                  insertions:
                    deletesAnchor && mode === "suggested"
                      ? insertions.map((insertion) =>
                          insertion.side === "beforeAnchor"
                            ? { ...insertion, side: "afterAnchor" as const }
                            : insertion,
                        )
                      : insertions,
                  deletesAnchor,
                };
              }),
          ),
          // Accepting the table's deletion takes away the table blocks are
          // put after.
          between: fc.array(
            table === "removed"
              ? fc.constantFrom<Between[]>("nothing", "reopen")
              : fc.constantFrom<Between[]>(
                  "nothing",
                  "reopen",
                  "accept",
                  "reopenAndAccept",
                ),
            { minLength: batches.length, maxLength: batches.length },
          ),
        }),
      );
  });

/**
 * Two tables with nothing between them read as one table once saved, and a
 * story ends with a paragraph: neither is about order.
 */
const readsApart = (placed: readonly Placed[]): boolean =>
  placed.at(-1)?.table !== true &&
  placed.every((entry, index) => !entry.table || !placed[index + 1]?.table);

const label = (placed: readonly Placed[]): string[] =>
  placed.map((entry) => entry.label);

const examples: [Case][] = [
  // A paragraph after the story's last paragraph, pending; then a paragraph
  // and a table after it in a batch that deletes it.
  [
    {
      anchorIsLast: true,
      table: "none",
      nextRemoved: false,
      mode: "tracked-changes",
      batches: [
        {
          insertions: [
            { kind: "paragraph", side: "afterAnchor", look: "centered" },
          ],
          deletesAnchor: false,
        },
        {
          insertions: [
            { kind: "paragraph", side: "afterAnchor", look: "centered" },
            { kind: "table", side: "afterAnchor", look: "plain" },
          ],
          deletesAnchor: true,
        },
      ],
      between: ["nothing", "nothing"],
    },
  ],
  // A paragraph after a table pending deletion, and a table after the
  // paragraph before it, in one batch.
  [
    {
      anchorIsLast: false,
      table: "removed",
      nextRemoved: false,
      mode: "tracked-changes",
      batches: [
        {
          insertions: [
            { kind: "paragraph", side: "afterTable", look: "centered" },
            { kind: "table", side: "afterAnchor", look: "plain" },
          ],
          deletesAnchor: false,
        },
      ],
      between: ["nothing"],
    },
  ],
  // The story's last paragraph deleted, so the anchor's break goes; then a
  // heading after the anchor and a plain paragraph before it in one batch.
  // The plain one reads as the anchor, not as the heading.
  [
    {
      anchorIsLast: false,
      table: "none",
      nextRemoved: true,
      mode: "tracked-changes",
      batches: [
        {
          insertions: [
            { kind: "paragraph", side: "afterAnchor", look: "heading" },
            { kind: "paragraph", side: "beforeAnchor", look: "plain" },
          ],
          deletesAnchor: false,
        },
      ],
      between: ["nothing"],
    },
  ],
];

test(
  "blocks put beside one anchor land in requested order in every mode",
  async () => {
    await assertProperty(
      fc.asyncProperty(caseArbitrary, async (setup) => {
        const reviewed = await replay(setup, setup.mode);
        fc.pre(
          readsApart(reviewed.requested) &&
            readsApart([...reviewed.pending, { label: NEXT, table: false }]),
        );
        expect(labelsOf(reviewed.accepted)).toEqual(label(reviewed.requested));
        expect(reviewed.looks).toEqual(
          reviewed.looks.map((look) => ({
            ...look,
            heading: look.asked === "heading",
            centered: look.asked === "centered",
          })),
        );
        // Applied directly, a removed table is gone before anything is put
        // after it.
        const afterRemovedTable =
          setup.table === "removed" &&
          setup.batches.some((batch) =>
            batch.insertions.some(({ side }) => side === "afterTable"),
          );
        if (afterRemovedTable) return;
        const direct = await replay(setup, "direct");
        expect(direct.accepted).toEqual(reviewed.accepted);
      }),
      {
        id: "blocks put beside one anchor land in requested order in every mode",
        numRuns: 80,
        examples,
      },
    );
  },
  propertyTestTimeout(120_000),
);
