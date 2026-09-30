/** Shared π/π′ oracle for operations; no editor or serializer dependency. */
import type {
  BlockContent,
  Document,
  Paragraph,
  Run,
} from "../packages/docx-core/src/model/document";
import {
  asParagraphContent,
  childNodes,
  type InlineNode,
  rebuildNode,
} from "../packages/docx-core/src/ops/leaves";
import { mergeAtSeam } from "../packages/docx-core/src/ops/resolve";
import { IDENTITY_SPACES, identityKeysIn } from "../packages/docx-core/src/ops/ids";
import { PARAGRAPH_MARK_FORMATTING_KEYS } from "../packages/docx-core/src/ops/types";

/** Empty formatting containers (including a captured empty trPr) state no properties. */
const authoredFormatting = <Formatting extends object>(formatting: Formatting | undefined) =>
  formatting &&
  Object.entries(formatting).some(([key, value]) => key !== "sourceXml" && value !== undefined)
    ? formatting
    : undefined;

const canonicalParagraphFormatting = (formatting: Paragraph["formatting"]) =>
  formatting
    ? authoredFormatting({
        ...formatting,
        runProperties: authoredFormatting(formatting.runProperties),
      })
    : undefined;

/** currentFormatting is an optional capture of the owning node's current properties. */
const canonicalRun = (run: Run): Run => ({
  ...run,
  formatting: authoredFormatting(run.formatting),
  ...(run.propertyChanges
    ? {
        propertyChanges: run.propertyChanges.map((change) => ({
          ...change,
          previousFormatting: authoredFormatting(change.previousFormatting),
          currentFormatting: authoredFormatting(change.currentFormatting ?? run.formatting),
        })),
      }
    : {}),
});

const canonicalList = (nodes: readonly InlineNode[]): InlineNode[] => {
  const out: InlineNode[] = [];
  for (const node of nodes) {
    const children = childNodes(node);
    const rebuilt = children === undefined ? node : rebuildNode(node, canonicalList(children));
    const own = rebuilt.type === "run" ? canonicalRun(rebuilt) : rebuilt;
    const last = out.at(-1);
    if (last === undefined) out.push(own);
    else out.splice(-1, 1, ...mergeAtSeam(last, own));
  }
  return out;
};

const canonicalParagraph = (paragraph: Paragraph): Paragraph => {
  const next = { ...paragraph, content: asParagraphContent(canonicalList(paragraph.content)) };
  next.formatting = canonicalParagraphFormatting(paragraph.formatting);
  if (next.propertyChanges) {
    next.propertyChanges = next.propertyChanges.map((change) => ({
      ...change,
      previousFormatting: canonicalParagraphFormatting(change.previousFormatting),
      currentFormatting: canonicalParagraphFormatting(change.currentFormatting ?? next.formatting),
    }));
  }
  delete next.listRendering;
  delete next.renderedPageBreakBefore;
  return next;
};

export const canonicalReviewBlocks = (blocks: readonly BlockContent[]): BlockContent[] =>
  blocks.map((block): BlockContent => {
    switch (block.type) {
      case "paragraph":
        return canonicalParagraph(block);
      case "table":
        return {
          ...block,
          rows: block.rows.map((row) => {
            const next = {
              ...row,
              cells: row.cells.map((cell) => ({
                ...cell,
                content: canonicalReviewBlocks(cell.content),
              })),
            };
            if (next.formatting && !authoredFormatting(next.formatting)) delete next.formatting;
            return next;
          }),
        };
      case "blockSdt":
      case "blockCustomXml":
        return { ...block, content: canonicalReviewBlocks(block.content) };
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        return block;
      default:
        return block satisfies never;
    }
  });

const MARK_KEYS: ReadonlySet<string> = new Set(PARAGRAPH_MARK_FORMATTING_KEYS);
const PARAGRAPH_IDENTITY_KEYS = new Set(["paraId", "textId", "preservedAttributes"]);

type ProjectReviewOptions = {
  document: Document;
  projection?: "π" | "π′";
};

/** The existing ops-law projection; L8 uses canonicalReviewBlocks to preserve ids. */
export const projectReview = ({ document, projection = "π" }: ProjectReviewOptions): unknown => {
  const ids = new Map<string, number>();
  const number = (key: string): number => {
    const known = ids.get(key);
    if (known !== undefined) return known;
    ids.set(key, ids.size + 1);
    return ids.size;
  };
  const walk = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(walk);
    if (typeof value !== "object" || value === null) return value;
    const paragraph = Reflect.get(value, "type") === "paragraph";
    const out: Record<string, unknown> = {};
    for (const [key, field] of Object.entries(value)) {
      // Optional undefined fields serialize exactly as absent fields.
      if (field === undefined) continue;
      if (paragraph && projection === "π′" && PARAGRAPH_IDENTITY_KEYS.has(key)) continue;
      out[key] = walk(field);
    }
    if (paragraph && projection === "π′") {
      const formatting = out["formatting"];
      if (typeof formatting === "object" && formatting !== null) {
        const retained = Object.fromEntries(
          Object.entries(formatting).filter(([key]) => !MARK_KEYS.has(key)),
        );
        if (Object.keys(retained).length === 0) delete out["formatting"];
        else out["formatting"] = retained;
      }
    }
    const info = out["info"];
    if (
      typeof info === "object" &&
      info !== null &&
      typeof Reflect.get(info, "author") === "string"
    ) {
      out["info"] = { ...info, id: number(`r:${String(Reflect.get(info, "id"))}`) };
    }
    if (typeof out["sdtType"] === "string" && typeof out["id"] === "number") {
      out["id"] = number(`c:${out["id"]}`);
    }
    return out;
  };
  return walk(canonicalReviewBlocks(document.package.document.content));
};

/** Test-only census shared by model and editor oracles. */
export const storyRevisionIds = (document: Document): number[] => {
  const prefix = `${IDENTITY_SPACES.REVISION}:`;
  return [
    ...new Set(
      identityKeysIn(document.package.document.content).flatMap((key) =>
        key.startsWith(prefix) ? [Number(key.slice(prefix.length))] : [],
      ),
    ),
  ].sort((left, right) => left - right);
};
