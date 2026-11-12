/**
 * A projection that reuses the blocks earlier projections converted equals a
 * projection converting every block.
 *
 * The block cache reuses a paragraph's nodes where every story fact its
 * conversion read still holds: the counters it drew from, the comments open
 * at its start, the bookmark and comment ids it asked about. Random runs of
 * edits through a canonical session, typing, deleting across paragraphs,
 * splitting and joining, editing and suggesting, over documents with lists,
 * comments across paragraphs, bookmarks, links, fields, notes, tables and
 * text boxes, are projected both ways after every edit.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  InlineProjectionCache,
  projectCanonicalInline,
} from "../canonical/canonicalInlineProjection";
import {
  type CanonicalProjection,
  createCanonicalSession,
  newStoryProjectionCaches,
  projectCanonicalStory,
} from "../canonical/canonicalSession";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import {
  CommentMarkerCache,
  rangedCommentIds,
} from "../docx/commentAnchorIndex";
import { parseDocx } from "../docx/parser";
import {
  collectPairedBookmarkIds,
  ProjectionBlockCache,
  toProseDoc,
} from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import type { BlockContent, Document, Paragraph } from "../types/document";
import {
  applyEdit,
  type Edit,
  editArbitrary,
  editModel,
  modelEditArbitrary,
  parsedSources,
  SOURCES,
} from "./__tests__/canonicalEditRuns";

setDefaultTimeout(propertyTestTimeout(300_000));

/** Text-box ids with their conversion's salt left out: each projection draws its own. */
const withoutSalt = (value: unknown): unknown => {
  if (typeof value === "string")
    return value.replace(/^[a-z0-9]{1,12}(?=:\d+(?::\d+)?$)/u, "salt");
  if (Array.isArray(value)) return value.map(withoutSalt);
  if (typeof value === "object" && value !== null)
    return Object.fromEntries(
      Object.entries(value).map(([key, entry]) => [key, withoutSalt(entry)]),
    );
  return value;
};

const projections = (document: Document, cache: ProjectionBlockCache) => {
  const options = {
    ...(document.package.styles === undefined
      ? {}
      : { styles: document.package.styles }),
  };
  return {
    cached: toProseDoc(document, { ...options, blockCache: cache }),
    full: toProseDoc(document, { ...options, textBoxGroupSalt: cache.salt }),
  };
};

/**
 * Each paragraph's address in a projection: where it starts, its text, and
 * the source gaps of its editor positions, with whether it still names the
 * same source records.
 */
const addressesOf = (
  projection: CanonicalProjection,
  records: CanonicalProjection = projection,
) => {
  const addresses: unknown[] = [];
  projection.doc.descendants((node) => {
    const paraId: unknown = node.attrs["paraId"];
    if (node.type.name !== "paragraph" || typeof paraId !== "string")
      return true;
    const address = projection.paragraph(paraId);
    const other = records.paragraph(paraId);
    addresses.push(
      address === undefined
        ? { paraId }
        : {
            paraId,
            start: address.start,
            text: address.text,
            boundaries: address.boundaries,
            seams: address.seams,
            sameSource: address.source === other?.source,
            sameContainer: address.container === other?.container,
          },
    );
    return true;
  });
  return addresses;
};

describe("a projection reusing converted blocks", () => {
  test("equals a projection converting every block over session edits", async () => {
    const sources = await parsedSources();
    await assertProperty(
      fc.property(
        fc.constantFrom(...sources),
        fc.array(editArbitrary, { minLength: 1, maxLength: 16 }),
        ({ document }, edits) => {
          const session = createCanonicalSession(
            structuredClone(document),
          ).unwrap();
          const cache = new ProjectionBlockCache();
          let state = EditorState.create({
            schema,
            doc: session.projection.doc,
          });
          const first = projections(session.document, cache);
          expect(first.cached.eq(first.full)).toBe(true);
          for (const edit of edits) {
            state = applyEdit(session, state, edit);
            const { cached, full } = projections(session.document, cache);
            expect(cached.toJSON()).toEqual(full.toJSON());
            // The session's own projection, built with its cache while it
            // staged the edit, against its document converted afresh.
            expect(withoutSalt(session.projection.doc.toJSON())).toEqual(
              withoutSalt(full.toJSON()),
            );
            // Its addresses, built from the paragraphs' kept gaps, against
            // every paragraph's gaps computed afresh.
            const fresh = projectCanonicalStory({
              document: session.document,
              styles: session.document.package.styles,
            }).unwrap();
            expect(addressesOf(session.projection, fresh)).toEqual(
              addressesOf(fresh),
            );
          }
        },
      ),
      {
        id: "projection block cache equals full conversion over session edits",
        numRuns: 300,
      },
    );
  });
});

/** The full projection, or nothing when the version is one the converter refuses. */
const convertedInFull = (document: Document, salt: string) => {
  try {
    return toProseDoc(document, { textBoxGroupSalt: salt });
  } catch {
    return undefined;
  }
};

describe("a projection reusing converted blocks across story versions", () => {
  test("equals a projection converting every block over story versions", async () => {
    const sources = await parsedSources();
    const sorted = (ids: ReadonlySet<number>) =>
      [...ids].toSorted((left, right) => left - right);
    const law = (counts: { marked: number }) =>
      fc.property(
        fc.constantFrom(...sources),
        fc.array(modelEditArbitrary, { minLength: 1, maxLength: 16 }),
        ({ document }, edits) => {
          const cache = new ProjectionBlockCache();
          const caches = newStoryProjectionCaches();
          let version = document;
          for (const edit of [undefined, ...edits]) {
            if (edit !== undefined) version = editModel(version, edit);
            // The story's bookmark and comment ids, from the blocks kept,
            // against a walk of every block.
            const blocks = version.package.document.content;
            const markers = cache.storyMarkers(blocks);
            expect(sorted(markers.pairedBookmarkIds)).toEqual(
              sorted(collectPairedBookmarkIds(blocks)),
            );
            expect(sorted(markers.rangedCommentIds)).toEqual(
              sorted(rangedCommentIds(blocks)),
            );
            if (
              markers.pairedBookmarkIds.size > 0 &&
              markers.rangedCommentIds.size > 0
            )
              counts.marked += 1;
            const full = convertedInFull(version, cache.salt);
            if (full === undefined) continue;
            const cached = toProseDoc(version, { blockCache: cache });
            expect(cached.toJSON()).toEqual(full.toJSON());
            // The story's addresses, each paragraph's gaps kept from the
            // versions before, against every paragraph's gaps afresh.
            const fresh = projectCanonicalStory({
              document: version,
              styles: version.package.styles,
            });
            const kept = projectCanonicalStory({
              document: version,
              styles: version.package.styles,
              caches,
            });
            expect(kept.isOk()).toBe(fresh.isOk());
            if (fresh.isOk() && kept.isOk())
              expect(addressesOf(kept.value, fresh.value)).toEqual(
                addressesOf(fresh.value),
              );
          }
        },
      );
    await assertProperty(law({ marked: 0 }), {
      id: "projection block cache equals full conversion over story versions",
      numRuns: 400,
    });
    // Versions of a fixed sample held both paired bookmarks and ranged
    // comments. Each case converts every version twice, so a quarter of the
    // runs, at the same ratio.
    const RUNS = 100;
    const counts = { marked: 0 };
    checkCoverageSample(law(counts), RUNS);
    expect(counts.marked).toBeGreaterThan(RUNS / 8);
  });
});

const withMainContent = (
  document: Document,
  content: BlockContent[],
): Document => ({
  ...document,
  package: {
    ...document.package,
    document: { ...document.package.document, content },
  },
});

describe("a projection reusing a block after the comments open before it change", () => {
  const text = (value: string) => ({
    type: "run" as const,
    content: [{ type: "text" as const, text: value }],
  });
  // The block after the first paragraph is the same object in every version:
  // only the comments the first paragraph opens before it change.
  const kept: Paragraph = {
    type: "paragraph",
    paraId: "7A000002",
    content: [text("kept")],
  };
  const versionOpening = (document: Document, ids: readonly number[]) =>
    withMainContent(document, [
      {
        type: "paragraph",
        paraId: "7A000001",
        content: [
          ...ids.map((id) => ({ type: "commentRangeStart" as const, id })),
          text("opening"),
        ],
      },
      kept,
      {
        type: "paragraph",
        paraId: "7A000003",
        content: [
          text("closing"),
          ...ids.map((id) => ({ type: "commentRangeEnd" as const, id })),
        ],
      },
    ]);
  const cases: [string, readonly number[], readonly number[]][] = [
    ["other comments, as many", [1], [2]],
    ["the same comments in another order", [1, 2], [2, 1]],
  ];
  for (const [name, before, after] of cases)
    test(`converts it again: ${name}`, async () => {
      const [source] = await parsedSources();
      if (source === undefined) throw new Error("No source document.");
      const cache = new ProjectionBlockCache();
      toProseDoc(versionOpening(source.document, before), {
        blockCache: cache,
      });
      const version = versionOpening(source.document, after);
      expect(toProseDoc(version, { blockCache: cache }).toJSON()).toEqual(
        toProseDoc(version, { textBoxGroupSalt: cache.salt }).toJSON(),
      );
    });
});

describe("a projection reusing converted blocks through a list suggestion", () => {
  for (const source of SOURCES)
    test(`equals a projection converting every block once resolved, ${source.id}`, async () => {
      const document = await parseDocx(
        (
          await prepareCanonicalDocxInput((await source.bytes()).slice().buffer)
        ).unwrap(),
        { preloadFonts: false, detectVariables: false },
      );
      for (const kind of ["bullet", "numbered"] as const)
        for (const resolution of ["accept", "reject"] as const)
          for (const from of [0.05, 0.3, 0.6, 0.9]) {
            // Sessions opened on the same parsed records, one after another.
            const session = createCanonicalSession(document).unwrap();
            const cache = new ProjectionBlockCache();
            let state = EditorState.create({
              schema,
              doc: session.projection.doc,
            });
            // A list suggested and typed into, then resolved in a session
            // opened on the same document, as a reviewer opens it: the two
            // sessions share the document's records and the facts kept for
            // them, and resolving brings back earlier blocks.
            for (const edit of [
              { kind, from, span: 0, text: "x", suggesting: true },
              { kind: "type", from, span: 0, text: "y", suggesting: true },
            ] as const satisfies readonly Edit[])
              state = applyEdit(session, state, edit);
            const reviewer = createCanonicalSession(session.document).unwrap();
            let reviewState = EditorState.create({
              schema,
              doc: reviewer.projection.doc,
            });
            for (const edit of [
              { kind: resolution, from, span: 0, text: "", suggesting: true },
              { kind: "type", from, span: 0, text: "z", suggesting: false },
              { kind: "undo", from, span: 0, text: "", suggesting: false },
            ] as const satisfies readonly Edit[]) {
              reviewState = applyEdit(reviewer, reviewState, edit);
              const { cached, full } = projections(reviewer.document, cache);
              expect(cached.toJSON()).toEqual(full.toJSON());
              expect(withoutSalt(reviewer.projection.doc.toJSON())).toEqual(
                withoutSalt(full.toJSON()),
              );
              const fresh = projectCanonicalStory({
                document: reviewer.document,
                styles: reviewer.document.package.styles,
              }).unwrap();
              expect(addressesOf(reviewer.projection, fresh)).toEqual(
                addressesOf(fresh),
              );
            }
          }
    });
});

type InlineArgs = Parameters<typeof projectCanonicalInline>[0];

/** An inline cache that also keeps every input a projection gave it. */
class CapturingInlineCache extends InlineProjectionCache {
  readonly calls: InlineArgs[] = [];
  override project(args: InlineArgs) {
    const context = args.commentContext;
    this.calls.push(
      context === undefined
        ? args
        : {
            ...args,
            commentContext: { ...context, openIds: new Set(context.openIds) },
          },
    );
    return super.project(args);
  }
}

const PERTURBATIONS = [
  "none",
  "bookmark",
  "ranged",
  "point",
  "open",
  "reorder",
  "node",
] as const;

/** One membership with one id's answer turned around. */
const toggled = (
  base: Pick<ReadonlySet<number>, "has">,
  id: number,
): Pick<ReadonlySet<number>, "has"> => ({
  has: (value) => (value === id ? !base.has(value) : base.has(value)),
});

/** A paragraph's real inputs with one story fact, or its editor node, changed. */
const perturbed = (
  args: InlineArgs,
  kind: (typeof PERTURBATIONS)[number],
  id: number,
  otherNode: InlineArgs["paragraph"],
): InlineArgs => {
  const context = args.commentContext;
  switch (kind) {
    case "none":
      return args;
    case "bookmark":
      return {
        ...args,
        pairedBookmarkIds: toggled(args.pairedBookmarkIds, id),
      };
    case "node":
      return { ...args, paragraph: otherNode };
    default:
      break;
  }
  if (context === undefined) return args;
  switch (kind) {
    case "ranged":
      return {
        ...args,
        commentContext: {
          ...context,
          rangedIds: toggled(context.rangedIds, id),
        },
      };
    case "point":
      return {
        ...args,
        commentContext: { ...context, pointIds: toggled(context.pointIds, id) },
      };
    case "open": {
      const openIds = new Set(context.openIds);
      if (!openIds.delete(id)) openIds.add(id);
      return { ...args, commentContext: { ...context, openIds } };
    }
    case "reorder":
      return {
        ...args,
        commentContext: {
          ...context,
          openIds: new Set([...context.openIds].reverse()),
        },
      };
    default:
      return kind satisfies never;
  }
};

/** Every number a paragraph's records or its editor node name as an id. */
const idsNamedBy = (args: InlineArgs): number[] => [
  ...new Set(
    [
      ...`${JSON.stringify(args.source)}${JSON.stringify(args.paragraph.toJSON())}`.matchAll(
        /"(?:id|commentId)":(\d+)/gu,
      ),
    ].map(([, id]) => Number(id)),
  ),
];

describe("a paragraph projection reusing its kept gaps", () => {
  test("equals the paragraph's gaps computed afresh under changed story facts", async () => {
    const sources = await parsedSources();
    const captured = sources.map(({ document }) => {
      const inline = new CapturingInlineCache();
      projectCanonicalStory({
        document,
        styles: document.package.styles,
        caches: {
          blocks: new ProjectionBlockCache(),
          inline,
          comments: new CommentMarkerCache(),
        },
      }).unwrap();
      return inline.calls;
    });
    await assertProperty(
      fc.property(
        fc.nat({ max: sources.length - 1 }),
        fc.array(
          fc.record({
            call: fc.nat(),
            // Two together reach a fact one alone leaves unread: a comment
            // no longer ranged in the story is read as a point comment.
            kinds: fc.array(fc.constantFrom(...PERTURBATIONS), {
              minLength: 1,
              maxLength: 2,
            }),
            id: fc.nat(),
            other: fc.nat(),
          }),
          { minLength: 1, maxLength: 24 },
        ),
        (sourceIndex, steps) => {
          const calls = captured[sourceIndex] ?? [];
          if (calls.length === 0) return;
          const cache = new InlineProjectionCache();
          for (const step of steps) {
            const real = calls[step.call % calls.length];
            const other = calls[step.other % calls.length];
            if (real === undefined || other === undefined) continue;
            const ids = idsNamedBy(real);
            const id =
              ids.length === 0 ? step.id % 4 : ids[step.id % ids.length];
            // The paragraph as the story projected it, with one fact changed,
            // then another, and back: none may take another's kept gaps.
            const chain = [real];
            for (const kind of step.kinds)
              chain.push(
                perturbed(chain.at(-1) ?? real, kind, id ?? 0, other.paragraph),
              );
            for (const input of [...chain, ...chain.toReversed()]) {
              const fresh = projectCanonicalInline(input);
              const kept = cache.project(input);
              expect(kept.isOk()).toBe(fresh.isOk());
              if (fresh.isOk() && kept.isOk())
                expect(kept.value).toEqual(fresh.value);
              else if (fresh.isErr() && kept.isErr())
                expect(kept.error.message).toBe(fresh.error.message);
            }
          }
        },
      ),
      {
        id: "inline projection cache equals fresh paragraph projection under changed story facts",
        numRuns: 400,
      },
    );
  });
});
