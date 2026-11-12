/**
 * The passes a story conversion makes over its top-level blocks, kept from
 * the last pass, equal the passes over every block.
 *
 * Carrier continuation keeps what the last pass gave the blocks it was
 * given, and validation checks only the children the last valid document did
 * not hold. Random runs of story versions rearrange real blocks (tracked
 * splits, carriers, bookmarks across paragraphs, tables, comments) and add
 * invalid ones, and each version is run both ways.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { Fragment, type Mark, type Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  CarrierContinuationCache,
  continueSourceCarrierIdentitiesAcrossInsertedParagraphMarks,
  continueSourceCarrierIdentitiesKeeping,
} from "../prosemirror/runIdentityAcrossRevisions";
import { schema } from "../prosemirror/schema";
import {
  assertValidProseMirrorDocumentKeeping,
  type ValidatedStory,
  validateProseMirrorDocument,
} from "../prosemirror/validation";
import {
  applyEdit,
  type Edit,
  parsedSources,
} from "./__tests__/canonicalEditRuns";
import { createCanonicalSession } from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(240_000));

/** Edits that leave tracked paragraph splits, carriers and lists behind. */
const SEEDING_EDITS: readonly Edit[] = [
  { kind: "split", from: 0.3, span: 0, text: "", suggesting: true },
  { kind: "type", from: 0.3, span: 0, text: "x", suggesting: true },
  { kind: "split", from: 0.6, span: 0, text: "", suggesting: true },
  {
    kind: "replace",
    from: 0.5,
    span: 0.1,
    text: "new words",
    suggesting: true,
  },
  { kind: "split", from: 0.8, span: 0, text: "", suggesting: false },
  { kind: "numbered", from: 0.2, span: 0, text: "", suggesting: true },
];

/** `node` with every source carrier given another identity, as a reopened half has. */
const reidentified = (node: PMNode): PMNode => {
  const mark = (candidate: Mark): Mark => {
    const attrs: Record<string, unknown> = candidate.attrs;
    if (
      candidate.type.name === "runIdentity" &&
      typeof attrs["id"] === "number"
    )
      return candidate.type.create({ ...attrs, id: attrs["id"] + 100_000 });
    const index = attrs["_docxHyperlinkIndex"];
    if (candidate.type.name === "hyperlink" && typeof index === "number")
      return candidate.type.create({
        ...attrs,
        _docxHyperlinkIndex: index + 100_000,
      });
    return candidate;
  };
  const marked = node.mark(node.marks.map(mark));
  if (node.isText || node.childCount === 0) return marked;
  const children: PMNode[] = [];
  node.forEach((child) => children.push(reidentified(child)));
  return marked.copy(Fragment.fromArray(children));
};

const endsWithInsertedMark = (node: PMNode): boolean => {
  const mark: unknown = node.attrs["pPrMark"];
  return (
    node.type.name === "paragraph" &&
    typeof mark === "object" &&
    mark !== null &&
    "kind" in mark &&
    (mark.kind === "ins" || mark.kind === "moveTo")
  );
};

/** Paragraphs ending with an inserted mark, each with the block after it carrying other identities. */
const splitPairs: [PMNode, PMNode][] = [];

/** Top-level blocks of real story documents, with invalid ones beside them. */
const blockPool = async (): Promise<PMNode[]> => {
  const pool: PMNode[] = [];
  for (const { document } of await parsedSources()) {
    const session = createCanonicalSession(structuredClone(document)).unwrap();
    let state = EditorState.create({ schema, doc: session.projection.doc });
    const take = () => {
      let previous: PMNode | undefined;
      session.projection.doc.forEach((child) => {
        pool.push(child);
        if (previous !== undefined && endsWithInsertedMark(previous))
          splitPairs.push([previous, reidentified(child)]);
        previous = child;
      });
    };
    take();
    for (const edit of SEEDING_EDITS) {
      state = applyEdit(session, state, edit);
      take();
    }
  }
  pool.push(
    schema.node("paragraph", { paraId: 12 } as never, [schema.text("bad")]),
    schema.node("paragraph", { bookmarks: "x" } as never, [schema.text("bad")]),
  );
  return pool;
};

const editArbitrary = fc.record({
  // A split pair often, so blocks continue across inserted marks.
  kind: fc.constantFrom(
    "replace",
    "insert",
    "remove",
    "repeat",
    "move",
    "pair",
    "pair",
    "pair",
  ),
  at: fc.nat(),
  to: fc.nat(),
  block: fc.nat(),
  count: fc.integer({ min: 1, max: 3 }),
});

/** A new story version sharing every block the edit leaves. */
const edited = (
  blocks: readonly PMNode[],
  pool: readonly PMNode[],
  edit: fc.RecordValue<typeof editArbitrary>,
): PMNode[] => {
  const next = [...blocks];
  const at = edit.at % Math.max(1, next.length);
  const block = pool[edit.block % pool.length] ?? next[0];
  if (block === undefined) return next;
  switch (edit.kind) {
    case "replace":
      next.splice(at, edit.count, block);
      break;
    case "insert":
      next.splice(at, 0, ...Array.from({ length: edit.count }, () => block));
      break;
    case "remove":
      if (next.length > edit.count) next.splice(at, edit.count);
      break;
    case "pair": {
      const pair = splitPairs[edit.block % Math.max(1, splitPairs.length)];
      if (pair === undefined) break;
      // In a table's cell every other time: containers continue their own blocks.
      next.splice(
        at,
        0,
        ...(edit.to % 2 === 0
          ? pair
          : [
              schema.node("table", null, [
                schema.node("tableRow", null, [
                  schema.node("tableCell", null, pair),
                ]),
              ]),
            ]),
      );
      break;
    }
    case "repeat": {
      const kept = next[at];
      if (kept !== undefined) next.splice(at, 0, kept);
      break;
    }
    case "move": {
      const [moved] = next.splice(at, 1);
      if (moved !== undefined)
        next.splice(edit.to % (next.length + 1), 0, moved);
      break;
    }
    default: {
      const unreachable: never = edit.kind;
      return unreachable;
    }
  }
  return next.length === 0 ? [block] : next;
};

const json = (blocks: readonly PMNode[]) =>
  blocks.map((block) => block.toJSON());

describe("a story conversion's passes kept from the last one", () => {
  test("equal the passes over every block across story versions", async () => {
    const pool = await blockPool();
    let keptContinuation = 0;
    let keptValidation = 0;
    let refused = 0;
    let continuing = 0;
    await assertProperty(
      fc.property(
        fc.array(fc.nat(), { minLength: 1, maxLength: 12 }),
        fc.array(editArbitrary, { minLength: 1, maxLength: 16 }),
        (initial, edits) => {
          const carriers = new CarrierContinuationCache();
          let validated: ValidatedStory | undefined;
          let blocks = initial.map(
            (index) => pool[index % pool.length] ?? panicMissing(),
          );
          for (const edit of [undefined, ...edits]) {
            if (edit !== undefined) blocks = edited(blocks, pool, edit);
            const last = carriers.last?.continued;
            const continued = continueSourceCarrierIdentitiesKeeping(
              blocks,
              carriers,
            );
            if (last?.some((node) => continued.includes(node)))
              keptContinuation += 1;
            const given = blocks;
            if (continued.some((node, index) => node !== given[index]))
              continuing += 1;
            expect(json(continued)).toEqual(
              json(
                continueSourceCarrierIdentitiesAcrossInsertedParagraphMarks(
                  blocks,
                ),
              ),
            );
            const doc = schema.node("doc", null, [...continued]);
            const full = validateProseMirrorDocument(doc);
            let message: string | undefined;
            const before = validated;
            try {
              validated = assertValidProseMirrorDocumentKeeping(
                doc,
                "invalid",
                validated,
              );
            } catch (error) {
              message = error instanceof Error ? error.message : String(error);
            }
            if (before !== undefined) keptValidation += 1;
            if (full.valid) {
              expect(message).toBeUndefined();
              expect(validated?.doc).toBe(doc);
            } else {
              refused += 1;
              expect(message).toBe(
                `invalid:\n${full.issues.map((issue) => `ProseMirror document error at ${issue.path}: ${issue.message}`).join("\n")}`,
              );
            }
          }
        },
      ),
      {
        id: "kept story conversion passes equal full passes",
        numRuns: 300,
      },
    );
    // The runs kept results and met invalid documents, not only full passes.
    expect(keptContinuation).toBeGreaterThan(300);
    expect(keptValidation).toBeGreaterThan(300);
    expect(refused).toBeGreaterThan(30);
    expect(continuing).toBeGreaterThan(40);
  });
});

const panicMissing = (): never => {
  throw new Error("Missing pool block");
};
