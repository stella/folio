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
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { DOCUMENT_SHAPES } from "../__tests__/documentShapes";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { mintListInstance } from "../docx/listNumberingInstances";
import { NO_NUMBERING_IDS_IN_USE } from "../docx/numberingIds";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import {
  ProjectionBlockCache,
  toProseDoc,
} from "../prosemirror/conversion/toProseDoc";
import { schema } from "../prosemirror/schema";
import type { Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import {
  InlineProjectionCache,
  projectCanonicalInline,
} from "./canonicalInlineProjection";
import {
  type CanonicalProjection,
  type CanonicalSession,
  createCanonicalSession,
  newStoryProjectionCaches,
  projectCanonicalStory,
} from "./canonicalSession";
import { prepareCanonicalCommands } from "./canonicalStructure";

setDefaultTimeout(propertyTestTimeout(240_000));

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="${R}" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
const box = (paraId: string, text: string) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="4000000" cy="2000000"/><wp:docPr id="${Number.parseInt(paraId, 16) % 1000}" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const paragraph = (paraId: string, inner: string) =>
  `<w:p w14:paraId="${paraId}">${inner}</w:p>`;
const text = (value: string) =>
  `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;

/** Boxes beside text and boxes standing alone, between plain paragraphs. */
const textBoxPackage = async (): Promise<Uint8Array> => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>` +
      paragraph("10000001", text("Before the boxes")) +
      paragraph("10000002", text("Host ") + box("20000001", "Inside one")) +
      paragraph("10000003", text("Between")) +
      paragraph("10000004", box("20000002", "Standing")) +
      paragraph("10000005", text("After ") + box("20000003", "Inside two")) +
      paragraph("10000006", text("Last paragraph")) +
      "<w:sectPr/></w:body></w:document>",
  );
  return zip.generate({ compression: "store" });
};

const run = (value: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text: value }],
});

/**
 * Story facts that cross paragraphs: a comment and a bookmark each opening in
 * one paragraph and closing several later, a link in every paragraph (each
 * drawing the next hyperlink instance) and notes referenced along the way.
 */
const crossingMarkersPackage = async (): Promise<Uint8Array> => {
  const document = createEmptyDocument();
  const paragraphs: Paragraph[] = Array.from({ length: 8 }, (_, index) => ({
    type: "paragraph",
    paraId: (0x30000000 + index).toString(16).toUpperCase(),
    content: [
      run(`Clause ${index} opens `),
      {
        type: "hyperlink",
        href: `https://example.com/${index}`,
        children: [run(`link ${index}`)],
      },
      run(" and closes."),
    ],
  }));
  const at = (index: number) => paragraphs[index]?.content ?? [];
  at(1).splice(1, 0, { type: "commentRangeStart", id: 1 });
  at(5).push(
    { type: "commentRangeEnd", id: 1 },
    { type: "commentReference", id: 1 },
  );
  at(2).splice(0, 0, { type: "bookmarkStart", id: 7, name: "Span" });
  at(6).push({ type: "bookmarkEnd", id: 7 });
  at(3).push({ type: "run", content: [{ type: "footnoteRef", id: 1 }] });
  at(4).push({ type: "run", content: [{ type: "footnoteRef", id: 2 }] });
  document.package.document.content = paragraphs;
  document.package.document.comments = [
    {
      id: 1,
      anchorKind: "range",
      author: "Reviewer",
      content: [
        { type: "paragraph", paraId: "40000001", content: [run("Note")] },
      ],
    },
  ];
  document.package.footnotes = [1, 2].map((id) => ({
    type: "footnote" as const,
    id,
    content: [
      {
        type: "paragraph" as const,
        paraId: (0x50000000 + id).toString(16).toUpperCase(),
        content: [
          {
            type: "run" as const,
            content: [
              { type: "noteMarker" as const, kind: "footnote" as const },
            ],
          },
          run(`Footnote ${id}`),
        ],
      },
    ],
  }));
  return new Uint8Array(await createDocx(document));
};

/** Direct list numbering at two levels, the indentation left to the list. */
const listLevelsPackage = async (): Promise<Uint8Array> => {
  const minted = mintListInstance(
    undefined,
    { kind: "bullet", start: 0 },
    NO_NUMBERING_IDS_IN_USE,
  );
  return new Uint8Array(
    await createDocx({
      package: {
        numbering: minted.definitions,
        document: {
          content: [0, 1, 0, undefined].map((ilvl, index) => {
            const item: Paragraph = {
              type: "paragraph",
              paraId: (index + 1).toString(16).padStart(8, "0"),
              content: [run("Alpha")],
            };
            if (ilvl !== undefined)
              item.formatting = {
                numPr: { kind: "reference", numId: minted.numId, ilvl },
              };
            return item;
          }),
        },
      },
    }),
  );
};

const SOURCES = [
  { id: "list-levels", bytes: listLevelsPackage },
  { id: "crossing-markers", bytes: crossingMarkersPackage },
  ...DOCUMENT_SHAPES.filter(({ id }) =>
    [
      "mixed-lists",
      "tables",
      "notes",
      "comments",
      "invisible-markers",
      "tracked-changes",
      "fields-links-bookmarks",
    ].includes(id),
  ).map(({ id, build }) => ({ id, bytes: build })),
  { id: "text-boxes", bytes: textBoxPackage },
];

const editArbitrary = fc.record({
  kind: fc.constantFrom(
    "type",
    "replace",
    "delete",
    "split",
    "join",
    "bullet",
    "numbered",
    "accept",
    "reject",
    "undo",
    "redo",
  ),
  from: fc.double({ min: 0, max: 1, noNaN: true }),
  // Wide enough to take markers, links and paragraph breaks with it.
  span: fc.double({ min: 0, max: 0.6, noNaN: true }),
  text: fc.constantFrom("x", "new words", " "),
  suggesting: fc.boolean(),
});
type Edit = fc.RecordValue<typeof editArbitrary>;

/** Every revision id the document's records carry. */
const revisionIdsOf = (document: Document): number[] => [
  ...new Set(
    [
      ...JSON.stringify(document.package.document).matchAll(
        /"info":\{"id":(\d+)/gu,
      ),
    ].map(([, id]) => Number(id)),
  ),
];

/** Apply one edit at the positions it names, when the session accepts it. */
const applyEdit = (
  session: CanonicalSession,
  state: EditorState,
  edit: Edit,
): EditorState => {
  session.setMode(
    edit.suggesting
      ? { type: "suggesting", author: "Reviewer" }
      : { type: "editing" },
  );
  const size = state.doc.content.size;
  const from = Math.max(1, Math.min(size - 1, Math.floor(edit.from * size)));
  const to =
    edit.kind === "replace" || edit.kind === "delete"
      ? Math.min(size - 1, from + Math.floor(edit.span * size))
      : from;
  const selected = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, from, to)),
  );
  const prepared = (() => {
    switch (edit.kind) {
      case "type":
      case "replace":
        return session.prepareReplace(selected, { from, to, text: edit.text });
      case "delete":
        return session.prepareReplace(selected, { from, to, text: "" });
      case "split":
        return session.prepareSplit(selected);
      case "join":
        return session.prepareJoin(selected, "backward");
      case "bullet":
      case "numbered":
        return prepareCanonicalCommands(session, selected, [
          {
            type: "toggleList",
            kind: edit.kind === "bullet" ? "bullet" : "decimal",
          },
        ]);
      // Resolving and undoing bring back earlier blocks and their sources.
      case "accept":
      case "reject":
        return session.prepareResolve(selected, {
          revisionIds: revisionIdsOf(session.document),
          resolution: edit.kind,
        });
      case "undo":
        return session.prepareUndo(selected);
      case "redo":
        return session.prepareRedo(selected);
      default: {
        const unreachable: never = edit.kind;
        return unreachable;
      }
    }
  })();
  if (prepared.isErr()) return state;
  const next = selected.apply(prepared.value.transaction);
  return prepared.value.publish().isOk() ? next : state;
};

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

/** Every source parsed as a canonical session loads it. */
const parsedSources = async () =>
  Promise.all(
    SOURCES.map(async (source) => ({
      id: source.id,
      document: await parseDocx(
        (
          await prepareCanonicalDocxInput((await source.bytes()).slice().buffer)
        ).unwrap(),
        { preloadFonts: false, detectVariables: false },
      ),
    })),
  );

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

const MARKERS = [
  () => ({
    type: "hyperlink" as const,
    href: "https://example.com/new",
    children: [run("new link")],
  }),
  () => ({ type: "commentRangeStart" as const, id: 1 }),
  () => ({ type: "commentRangeEnd" as const, id: 1 }),
  () => ({ type: "bookmarkStart" as const, id: 7, name: "Span" }),
  () => ({ type: "bookmarkEnd" as const, id: 7 }),
  () => ({ type: "bookmarkStart" as const, id: 9, name: "Other" }),
  () => ({
    type: "run" as const,
    content: [{ type: "footnoteRef" as const, id: 1 }],
  }),
  () => ({ type: "commentReference" as const, id: 1 }),
] as const;

const modelEditArbitrary = fc.record({
  kind: fc.constantFrom("touch", "remove", "insert", "repeat", "move"),
  at: fc.nat(),
  to: fc.nat(),
  markers: fc.array(fc.nat({ max: MARKERS.length - 1 }), { maxLength: 3 }),
});
type ModelEdit = fc.RecordValue<typeof modelEditArbitrary>;

/**
 * A new version of the story sharing every block the edit leaves, as the
 * operations make one: a block changed, removed, inserted with markers that
 * move what the blocks after it read, repeated, or moved.
 */
const editModel = (document: Document, edit: ModelEdit): Document => {
  const blocks = [...document.package.document.content];
  if (blocks.length === 0) return document;
  const index = edit.at % blocks.length;
  const markers = edit.markers.map((choice) => MARKERS[choice]?.() ?? run("?"));
  const fresh = (): Paragraph => ({
    type: "paragraph",
    content: [run("Inserted "), ...markers, run(" text")],
  });
  switch (edit.kind) {
    case "touch": {
      const block = blocks[index];
      if (block?.type === "paragraph")
        blocks[index] = { ...block, content: [...block.content, ...markers] };
      break;
    }
    case "remove":
      if (blocks.length > 1) blocks.splice(index, 1);
      break;
    case "insert":
      blocks.splice(index, 0, fresh());
      break;
    case "repeat": {
      const block = blocks[index];
      if (block !== undefined) blocks.splice(index, 0, block);
      break;
    }
    case "move": {
      const [block] = blocks.splice(index, 1);
      if (block !== undefined)
        blocks.splice(edit.to % (blocks.length + 1), 0, block);
      break;
    }
    default: {
      const unreachable: never = edit.kind;
      return unreachable;
    }
  }
  return {
    ...document,
    package: {
      ...document.package,
      document: { ...document.package.document, content: blocks },
    },
  };
};

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
    await assertProperty(
      fc.property(
        fc.constantFrom(...sources),
        fc.array(modelEditArbitrary, { minLength: 1, maxLength: 16 }),
        ({ document }, edits) => {
          const cache = new ProjectionBlockCache();
          const caches = newStoryProjectionCaches();
          let version = document;
          for (const edit of [undefined, ...edits]) {
            if (edit !== undefined) version = editModel(version, edit);
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
      ),
      {
        id: "projection block cache equals full conversion over story versions",
        numRuns: 400,
      },
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
        caches: { blocks: new ProjectionBlockCache(), inline },
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
