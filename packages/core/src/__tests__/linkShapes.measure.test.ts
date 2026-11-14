// Measurement only: counts of link shapes a suggested link change refuses,
// and what the legacy editor does with each. Never merged.
import { expect, setDefaultTimeout, test } from "bun:test";
import { readdir, readFile } from "node:fs/promises";
import { join } from "node:path";
import { TextSelection } from "prosemirror-state";

import {
  COLLISION_FIXTURES,
  FIXTURES,
  STORY_FIXTURES,
} from "../../../../test/consumer-scenarios/support/documents";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import type { Document, ParagraphContent } from "../types/document";
import { DOCUMENT_SHAPES, shapeArrayBuffer } from "./documentShapes";
import {
  createHarnessState,
  harnessRuntimeManager,
  parseShapeDocument,
  resolveAllChanges,
} from "./editorHarness";

setDefaultTimeout(600_000);

const MARKERS = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "commentReference",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);
const IDENTIFIED = new Set([
  "footnoteRef",
  "endnoteRef",
  "drawing",
  "shape",
  "preservedXml",
]);

type Census = { links: number; markerInside: number; identified: number };

const census = (value: unknown, out: Census): void => {
  if (Array.isArray(value)) {
    for (const item of value) census(item, out);
    return;
  }
  if (typeof value !== "object" || value === null) return;
  const record = value as { type?: unknown; children?: unknown };
  if (record.type === "hyperlink" && Array.isArray(record.children)) {
    out.links += 1;
    const children = record.children as { type: string; content?: unknown }[];
    if (
      children.some(({ type }) => MARKERS.has(type)) &&
      children.some(({ type }) => type === "run")
    )
      out.markerInside += 1;
    if (
      children.some(
        (child) =>
          child.type === "run" &&
          Array.isArray(child.content) &&
          (child.content as { type: string }[]).some(({ type }) =>
            IDENTIFIED.has(type),
          ),
      )
    )
      out.identified += 1;
  }
  for (const field of Object.values(value)) census(field, out);
};

const tokens = (document: Document): string[] => {
  const out: string[] = [];
  const walk = (items: readonly ParagraphContent[], link: string): void => {
    for (const item of items) {
      switch (item.type) {
        case "run":
          for (const part of item.content) {
            if (part.type === "text")
              for (const text of part.text) out.push(`${text}${link}`);
            else if (IDENTIFIED.has(part.type))
              out.push(`[${part.type}]${link}`);
          }
          break;
        case "hyperlink":
          walk(item.children, `@${item.href ?? item.anchor ?? ""}`);
          break;
        case "insertion":
        case "deletion":
        case "moveFrom":
        case "moveTo":
          walk(item.content, link);
          break;
        case "bookmarkStart":
        case "bookmarkEnd":
          out.push(`<${item.type}:${String(item.id)}>`);
          break;
        case "commentRangeStart":
        case "commentRangeEnd":
          out.push(`<${item.type}:${String(item.id)}>`);
          break;
        default:
          break;
      }
    }
  };
  for (const block of document.package.document.content)
    if (block.type === "paragraph") walk(block.content, "");
  return out;
};

/** The first linked range of the legacy document, in PM positions. */
const linkRange = (
  state: ReturnType<typeof createHarnessState>,
): { from: number; to: number } | undefined => {
  let range: { from: number; to: number } | undefined;
  state.doc.descendants((node, position) => {
    if (!node.isInline) return true;
    if (!node.marks.some((mark) => mark.type.name === "hyperlink")) {
      if (range !== undefined) return false;
      return true;
    }
    if (range === undefined)
      range = { from: position, to: position + node.nodeSize };
    else if (range.to === position) range.to = position + node.nodeSize;
    return true;
  });
  return range;
};

type Verdict = string;

const legacyRemoval = (document: Document): Verdict => {
  try {
    const run = (mode: "editing" | "suggesting") => {
      let state = createHarnessState(document, mode);
      const range = linkRange(state);
      if (range === undefined) return undefined;
      state = state.apply(
        state.tr.setSelection(
          TextSelection.create(state.doc, range.from, range.to),
        ),
      );
      const command =
        harnessRuntimeManager(state).requireCommand("removeHyperlink")();
      let applied = false;
      command(state, (tr) => {
        state = state.apply(tr);
        applied = true;
      });
      return applied ? state : null;
    };
    const direct = run("editing");
    const tracked = run("suggesting");
    if (direct === undefined) return "no-link";
    if (tracked === null || direct === null) return "refuses";
    if (tracked === undefined) return "no-link";
    const model = (state: NonNullable<typeof tracked>) =>
      tokens(fromProseDoc(state.doc, document));
    const pending = model(tracked);
    const accepted = model(resolveAllChanges(tracked, "accept"));
    const directAccepted = model(resolveAllChanges(direct, "accept"));
    const rejected = model(resolveAllChanges(tracked, "reject"));
    const original = model(
      resolveAllChanges(createHarnessState(document, "editing"), "reject"),
    );
    const duplicated = (list: readonly string[]) =>
      list.filter((token) => token.startsWith("<")).length !==
      new Set(list.filter((token) => token.startsWith("<"))).size;
    const notes: string[] = [];
    if (JSON.stringify(accepted) !== JSON.stringify(directAccepted))
      notes.push(
        `accept-differs ${JSON.stringify(accepted)} vs ${JSON.stringify(directAccepted)}`,
      );
    if (JSON.stringify(rejected) !== JSON.stringify(original))
      notes.push(
        `reject-differs ${JSON.stringify(rejected)} vs ${JSON.stringify(original)}`,
      );
    if (duplicated(pending)) notes.push("pending-duplicates-markers");
    if (pending.length === tokens(document).length)
      notes.push("untracked-or-unchanged");
    return notes.length === 0 ? "applies-correctly" : notes.join("; ");
  } catch (error) {
    return `throws ${String(error)}`;
  }
};

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const synthetic: Record<string, ParagraphContent[]> = {
  bookmarkBetween: [
    run("before "),
    {
      type: "hyperlink",
      href: "https://a.example/",
      children: [
        run("ab"),
        { type: "bookmarkStart", id: 7, name: "mark" },
        { type: "bookmarkEnd", id: 7 },
        run("cd"),
      ],
    },
    run(" after"),
  ],
  bookmarkAtEnd: [
    run("before "),
    {
      type: "hyperlink",
      href: "https://a.example/",
      children: [
        run("ab"),
        { type: "bookmarkStart", id: 7, name: "mark" },
        { type: "bookmarkEnd", id: 7 },
      ],
    },
    run(" after"),
  ],
  noteInLink: [
    run("before "),
    {
      type: "hyperlink",
      href: "https://a.example/",
      children: [
        run("ab"),
        { type: "run", content: [{ type: "footnoteRef", id: 1 }] },
      ],
    },
    run(" after"),
  ],
};

test("measure link shapes", async () => {
  const report: Record<string, unknown> = {};
  const sets: Record<string, Record<string, () => Promise<Uint8Array>>> = {
    consumer: { ...FIXTURES, ...COLLISION_FIXTURES, ...STORY_FIXTURES },
    shapes: Object.fromEntries(
      DOCUMENT_SHAPES.map((shape) => [
        shape.id,
        async () => new Uint8Array(await shapeArrayBuffer(shape)),
      ]),
    ),
  };
  const corpusDir = join(
    import.meta.dir,
    "../docx/__tests__/__fixtures__/corpus",
  );
  sets["corpus"] = Object.fromEntries(
    (await readdir(corpusDir))
      .filter((name) => name.endsWith(".docx"))
      .map((name) => [
        name,
        async () => new Uint8Array(await readFile(join(corpusDir, name))),
      ]),
  );
  for (const [setName, fixtures] of Object.entries(sets)) {
    const total: Census = { links: 0, markerInside: 0, identified: 0 };
    const hits: string[] = [];
    for (const [name, load] of Object.entries(fixtures)) {
      try {
        const document = await parseShapeDocument(await load());
        const own: Census = { links: 0, markerInside: 0, identified: 0 };
        census(document.package, own);
        total.links += own.links;
        total.markerInside += own.markerInside;
        total.identified += own.identified;
        if (own.markerInside + own.identified > 0)
          hits.push(`${name}:${JSON.stringify(own)}`);
      } catch (error) {
        hits.push(`${name}: unreadable ${String(error)}`);
      }
    }
    report[setName] = { fixtures: Object.keys(fixtures).length, total, hits };
  }
  for (const [name, content] of Object.entries(synthetic)) {
    const document: Document = {
      package: {
        document: {
          content: [{ type: "paragraph", paraId: "12345678", content }],
        },
      },
    };
    report[`legacy-remove:${name}`] = legacyRemoval(document);
  }
  console.log(`S20MEASURE ${JSON.stringify(report)}`);
  expect(true).toBe(true);
});
