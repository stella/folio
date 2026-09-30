/** Shared π/π′ oracle for operations; no editor or serializer dependency. */
import type { BlockContent, Document, Paragraph } from "../../model/document";
import { asParagraphContent, childNodes, type InlineNode, rebuildNode } from "../leaves";
import { mergeAtSeam } from "../resolve";
import { PARAGRAPH_MARK_FORMATTING_KEYS } from "../types";

const canonicalList = (nodes: readonly InlineNode[]): InlineNode[] => {
  const out: InlineNode[] = [];
  for (const node of nodes) {
    const children = childNodes(node);
    const own = children === undefined ? node : rebuildNode(node, canonicalList(children));
    const last = out.at(-1);
    if (last === undefined) out.push(own);
    else out.splice(-1, 1, ...mergeAtSeam(last, own));
  }
  return out;
};

const canonicalParagraph = (paragraph: Paragraph): Paragraph => {
  const next = { ...paragraph, content: asParagraphContent(canonicalList(paragraph.content)) };
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
          rows: block.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({
              ...cell,
              content: canonicalReviewBlocks(cell.content),
            })),
          })),
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
