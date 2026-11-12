/**
 * A projection keeping the top-level blocks of the one before it equals a
 * projection walking every block.
 *
 * A projection made from the one before keeps each block an edit left, its
 * paragraphs' starts relative to the block, and walks only the blocks between
 * the kept runs. Random runs of edits through a canonical session, and story
 * versions that change, remove, insert, repeat and move blocks with markers
 * that change what the blocks after them read, are projected both ways after
 * every edit and compared through every address the projection answers.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { schema } from "../prosemirror/schema";
import type { Document } from "../types/document";
import {
  applyEdit,
  editArbitrary,
  editModel,
  modelEditArbitrary,
  parsedSources,
} from "./__tests__/canonicalEditRuns";
import {
  type CanonicalAddressEdge,
  type CanonicalProjection,
  createCanonicalSession,
  newStoryProjectionCaches,
  projectCanonicalStory,
} from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(300_000));

const EDGES: readonly CanonicalAddressEdge[] = [
  "caret",
  "afterMarkers",
  "beforeMarkers",
];

/** Text-box ids with their conversion's salt left out: each conversion draws its own. */
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

const answer = <T>(result: {
  isOk: () => boolean;
  value?: T;
  error?: { message: string };
}) =>
  result.isOk() ? { ok: result.value } : { refused: result.error?.message };

/**
 * Everything a projection answers about its document: each paragraph's
 * address and the paragraph after it in its container, each editor
 * position's source gap at every edge, the paragraph starting there, and the
 * position each paragraph gap maps back to. Records are compared by whether
 * they are the ones `records` names.
 */
const answersOf = (
  projection: CanonicalProjection,
  records: CanonicalProjection,
) => {
  const paragraphs: unknown[] = [];
  projection.doc.descendants((node) => {
    const paraId: unknown = node.attrs["paraId"];
    if (node.type.name !== "paragraph" || typeof paraId !== "string")
      return true;
    const address = projection.paragraph(paraId);
    const other = records.paragraph(paraId);
    paragraphs.push(
      address === undefined
        ? { paraId }
        : {
            paraId,
            blockId: address.blockId,
            start: address.start,
            text: address.text,
            boundaries: address.boundaries,
            seams: address.seams,
            node: withoutSalt(address.node.toJSON()),
            sameSource: address.source === other?.source,
            sameContainer: address.container === other?.container,
            next: projection.nextParagraphInContainer(paraId)?.blockId,
            positions: Array.from(
              { length: address.text.length + 1 },
              (_, offset) =>
                answer(
                  projection.positionAt({
                    story: projection.story,
                    blockId: paraId,
                    offset,
                  }),
                ),
            ),
          },
    );
    return true;
  });
  const positions = Array.from(
    { length: projection.doc.content.size + 1 },
    (_, position) => ({
      position,
      addresses: EDGES.map((edge) =>
        answer(projection.addressAt(position, edge)),
      ),
      startingHere: projection.paragraphStartingAt(position)?.blockId,
      edges: projection.rangeEdgesAt(
        position,
        Math.min(projection.doc.content.size, position + 7),
      ),
    }),
  );
  return {
    first: projection.firstParagraph?.blockId,
    paragraphs,
    positions,
  };
};

/** Whether `projection` holds any block record `previous` held. */
const keepsBlocks = (
  projection: CanonicalProjection,
  previous: CanonicalProjection,
): boolean =>
  projection.blocks.some((block) => previous.blocks.includes(block));

describe("a projection keeping the blocks of the one before", () => {
  test("equals a projection walking every block over session edits", async () => {
    const sources = await parsedSources();
    const law = (counts: { kept: number }) =>
      fc.property(
        fc.constantFrom(...sources),
        fc.array(editArbitrary, { minLength: 1, maxLength: 16 }),
        ({ document }, edits) => {
          const session = createCanonicalSession(
            structuredClone(document),
          ).unwrap();
          let state = EditorState.create({
            schema,
            doc: session.projection.doc,
          });
          for (const edit of edits) {
            const previous = session.projection;
            state = applyEdit(session, state, edit);
            const projection = session.projection;
            if (projection !== previous && keepsBlocks(projection, previous))
              counts.kept += 1;
            const full = projectCanonicalStory({
              document: session.document,
              styles: session.document.package.styles,
            }).unwrap();
            expect(answersOf(projection, full)).toEqual(answersOf(full, full));
          }
        },
      );
    await assertProperty(law({ kept: 0 }), {
      id: "kept projection equals full projection over session edits",
      numRuns: 200,
    });
    // A fixed sample reached the projection that keeps blocks, not only the
    // full walk. Each case projects every edit twice, so a quarter of the
    // runs, at the same ratio.
    const RUNS = 50;
    const counts = { kept: 0 };
    checkCoverageSample(law(counts), RUNS);
    expect(counts.kept).toBeGreaterThan(RUNS);
  });

  test("equals a projection walking every block over story versions", async () => {
    const sources = await parsedSources();
    const law = (counts: { kept: number }) =>
      fc.property(
        fc.constantFrom(...sources),
        fc.array(modelEditArbitrary, { minLength: 1, maxLength: 16 }),
        ({ document }, edits) => {
          const caches = newStoryProjectionCaches();
          let version: Document = document;
          let previous: CanonicalProjection | undefined;
          for (const edit of [undefined, ...edits]) {
            if (edit !== undefined) version = editModel(version, edit);
            const full = projectCanonicalStory({
              document: version,
              styles: version.package.styles,
            });
            const projected = projectCanonicalStory({
              document: version,
              styles: version.package.styles,
              caches,
              ...(previous === undefined ? {} : { previous }),
            });
            expect(projected.isOk()).toBe(full.isOk());
            if (!projected.isOk() || !full.isOk()) continue;
            if (
              previous !== undefined &&
              keepsBlocks(projected.value, previous)
            )
              counts.kept += 1;
            expect(answersOf(projected.value, full.value)).toEqual(
              answersOf(full.value, full.value),
            );
            previous = projected.value;
          }
        },
      );
    await assertProperty(law({ kept: 0 }), {
      id: "kept projection equals full projection over story versions",
      numRuns: 300,
    });
    // As above: a quarter of the runs, at the same ratio.
    const RUNS = 75;
    const counts = { kept: 0 };
    checkCoverageSample(law(counts), RUNS);
    expect(counts.kept).toBeGreaterThan(RUNS);
  });
});
