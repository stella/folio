/**
 * Sources and random edit runs shared by the properties that hold an
 * incremental pass of a canonical session equal to its full computation.
 *
 * The sources cover what crosses paragraphs or numbers them: lists at
 * several levels, a comment and a bookmark spanning paragraphs, links and
 * notes, tables, invisible markers, tracked changes and text boxes. Edits
 * run through a session, or make new versions of a story sharing every
 * block they leave, as the operations do.
 */
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import { DOCUMENT_SHAPES } from "../../__tests__/documentShapes";
import { prepareCanonicalDocxInput } from "../../docx/canonicalSessionInput";
import { mintListInstance } from "../../docx/listNumberingInstances";
import { NO_NUMBERING_IDS_IN_USE } from "../../docx/numberingIds";
import { carryParagraphPropertySource } from "../../docx/paragraphPropertySource";
import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import type { BlockContent, Document, Paragraph } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { type CanonicalSession } from "../canonicalSession";
import { prepareCanonicalCommands } from "../canonicalStructure";

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

export const run = (value: string) => ({
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

export const SOURCES = [
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

export const editArbitrary = fc.record({
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
export type Edit = fc.RecordValue<typeof editArbitrary>;

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
export const applyEdit = (
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

/** Every source parsed as a canonical session loads it. */
export const parsedSources = () =>
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

export const MARKERS = [
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

/** Every paragraph of a story, in its tables too, in document order. */
const storyParagraphsOf = (blocks: readonly BlockContent[]): Paragraph[] =>
  blocks.flatMap((block) => {
    switch (block.type) {
      case "paragraph":
        return [block];
      case "table":
        return block.rows.flatMap((row) =>
          row.cells.flatMap((cell) => storyParagraphsOf(cell.content)),
        );
      case "blockSdt":
      case "blockCustomXml":
        return storyParagraphsOf(block.content);
      default:
        return [];
    }
  });

/** `blocks` with `from` replaced by `to` wherever it is held, as a new version holds it. */
const replacingParagraph = (
  blocks: readonly BlockContent[],
  from: Paragraph,
  to: Paragraph,
): BlockContent[] =>
  blocks.map((block): BlockContent => {
    switch (block.type) {
      case "paragraph":
        return block === from ? to : block;
      case "table":
        return {
          ...block,
          rows: block.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({
              ...cell,
              content: replacingParagraph(cell.content, from, to),
            })),
          })),
        };
      case "blockSdt":
      case "blockCustomXml":
        return {
          ...block,
          content: replacingParagraph(block.content, from, to),
        };
      default:
        return block;
    }
  });

export const modelEditArbitrary = fc.record({
  kind: fc.constantFrom(
    "touch",
    "touchCell",
    "restore",
    "remove",
    "insert",
    "repeat",
    "move",
  ),
  at: fc.nat(),
  to: fc.nat(),
  markers: fc.array(fc.nat({ max: MARKERS.length - 1 }), { maxLength: 3 }),
});
export type ModelEdit = fc.RecordValue<typeof modelEditArbitrary>;

/**
 * A new version of the story sharing every block the edit leaves, as the
 * operations make one: a block changed, a table's first cell changed,
 * removed, inserted with markers that move what the blocks after it read,
 * repeated, or moved. `restore` gives one paragraph, in a table too, another
 * paragraph's source facts and indentation projection, as history restores
 * them: in place while no reader kept the paragraph, else on a new paragraph
 * in its place.
 */
export const editModel = (document: Document, edit: ModelEdit): Document => {
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
    case "touchCell": {
      const table = blocks.find(
        (block, at) => block.type === "table" && at >= index,
      );
      const [row, ...rows] = table?.type === "table" ? table.rows : [];
      const [cell, ...cells] = row?.cells ?? [];
      const [first, ...rest] = cell?.content ?? [];
      if (
        table?.type !== "table" ||
        row === undefined ||
        cell === undefined ||
        first?.type !== "paragraph"
      )
        break;
      blocks[blocks.indexOf(table)] = {
        ...table,
        rows: [
          {
            ...row,
            cells: [
              {
                ...cell,
                content: [
                  { ...first, content: [...first.content, ...markers] },
                  ...rest,
                ],
              },
              ...cells,
            ],
          },
          ...rows,
        ],
      };
      break;
    }
    case "restore": {
      const paragraphs = storyParagraphsOf(blocks);
      const target = paragraphs[index % Math.max(1, paragraphs.length)];
      const source = paragraphs[edit.to % Math.max(1, paragraphs.length)];
      if (target === undefined || source === undefined) return document;
      const restored = carryParagraphPropertySource({
        target,
        source,
        indentation: "restore",
      });
      if (restored === target) return document;
      const replaced = replacingParagraph(blocks, target, restored);
      blocks.splice(0, blocks.length, ...replaced);
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
