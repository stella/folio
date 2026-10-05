import { Result, panic } from "better-result";
import { sanitizeExternalUrl } from "../markdown/href";
import type { Document, Hyperlink, ParagraphContent, Run } from "../model/document";
import { applyDocumentOps } from "./apply";
import type { EditorIntentMode } from "./editorIntent";
import { resolveGap } from "./gaps";
import { asParagraphContent, partitionContent } from "./leaves";
import { paragraphLength } from "./offsets";
import { idKey } from "./ids";
import { selectedParagraphRuns } from "./plan";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { structurallyEqual } from "./equality";
import { DOCUMENT_OP_TYPES, type DocumentOp, type TextPosition } from "./types";

export type HyperlinkEditorIntent =
  | { type: "setHyperlink"; from: TextPosition; to: TextPosition; href: string; tooltip?: string }
  | { type: "removeHyperlink"; from: TextPosition; to: TextPosition; hyperlinkStyleId?: string }
  | {
      type: "insertHyperlink";
      from: TextPosition;
      to: TextPosition;
      text: string;
      href: string;
      tooltip?: string;
    };

const refuse = (
  message: string,
  reason: DocumentOpRefusal["reason"] = DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
) =>
  Result.err(new DocumentOpRefusal({ message, reason, opType: DOCUMENT_OP_TYPES.REPLACE_INLINE }));

/** Cutting these records cannot duplicate review identities or unmodeled wrappers. */
const editable = (items: readonly ParagraphContent[]): boolean =>
  items.every((item) => {
    switch (item.type) {
      case "run":
        return (item.propertyChanges?.length ?? 0) === 0;
      case "bookmarkStart":
      case "bookmarkEnd":
        return true;
      case "hyperlink":
        return editable(item.children);
      default:
        return false;
    }
  });

const HYPERLINK_STYLE_ID = "Hyperlink";

const styledRun = (
  run: Run,
  intent: Exclude<HyperlinkEditorIntent, { type: "insertHyperlink" }>,
): Run => {
  if (run.formatting === undefined) return run;
  const formatting = { ...run.formatting };
  if (intent.type === "setHyperlink") delete formatting.color;
  else if (
    formatting.styleId === HYPERLINK_STYLE_ID ||
    (intent.hyperlinkStyleId !== undefined && formatting.styleId === intent.hyperlinkStyleId)
  )
    delete formatting.styleId;
  if (structurallyEqual(formatting, run.formatting)) return run;
  const next = { ...run };
  if (Object.keys(formatting).length === 0) delete next.formatting;
  else next.formatting = formatting;
  return next;
};

/** Flatten only the edited hyperlinks; unrelated run formatting and zero-width markers stay. */
const linkedChildren = (
  items: readonly ParagraphContent[],
  intent: Exclude<HyperlinkEditorIntent, { type: "insertHyperlink" }>,
): ParagraphContent[] => {
  const children: ParagraphContent[] = [];
  for (const item of items) {
    switch (item.type) {
      case "run":
        children.push(intent.type === "setHyperlink" ? styledRun(item, intent) : item);
        break;
      case "bookmarkStart":
      case "bookmarkEnd":
        children.push(item);
        break;
      case "hyperlink":
        if (item.children.length === 0) {
          children.push(item);
          break;
        }
        for (const child of item.children) {
          if (child.type === "run") children.push(styledRun(child, intent));
          else children.push(child);
        }
        break;
      default:
        return panic("Validated hyperlink edit contains an unsupported inline node.");
    }
  }
  return children;
};

const hyperlink = (
  intent: Exclude<HyperlinkEditorIntent, { type: "removeHyperlink" }>,
  children: Hyperlink["children"],
): Hyperlink => {
  const result: Hyperlink = { type: "hyperlink", children };
  if (intent.href.startsWith("#")) result.anchor = intent.href.slice(1);
  else if (intent.href !== "") result.href = intent.href;
  if (intent.tooltip !== undefined) result.tooltip = intent.tooltip;
  return result;
};

/** Link text segments while structural markers and empty links retain their authored positions. */
const wrapLinkedChildren = (
  intent: Extract<HyperlinkEditorIntent, { type: "setHyperlink" }>,
  children: readonly ParagraphContent[],
): ParagraphContent[] => {
  const result: ParagraphContent[] = [];
  let link: Hyperlink | undefined;
  for (const child of children) {
    if (child.type === "run") {
      if (link === undefined) {
        link = hyperlink(intent, []);
        result.push(link);
      }
      link.children.push(child);
      continue;
    }
    link = undefined;
    result.push(child);
  }
  return result;
};

type CompileHyperlinkIntentOptions = {
  intent: HyperlinkEditorIntent;
  mode: EditorIntentMode;
  compileEmptyReplacement: (
    document: Document,
    range: Pick<HyperlinkEditorIntent, "from" | "to">,
  ) => Result<{ ops: DocumentOp[]; selection: TextPosition }, DocumentOpRefusal>;
};

/** Wrapper edits compile to exact-inverse content ops, never to a projection snapshot. */
export const compileHyperlinkIntent = (
  document: Document,
  { intent: sourceIntent, mode, compileEmptyReplacement }: CompileHyperlinkIntentOptions,
): Result<{ ops: DocumentOp[]; selection: TextPosition }, DocumentOpRefusal> => {
  let intent = sourceIntent;
  if (intent.type !== "removeHyperlink" && intent.href !== "" && !intent.href.startsWith("#")) {
    const href = sanitizeExternalUrl(intent.href);
    if (href === undefined)
      return refuse(
        "Hyperlink targets must use a supported URL scheme.",
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
      );
    intent = { ...intent, href };
  }
  const selected = selectedParagraphRuns(document, intent.from, intent.to);
  if (selected.isErr()) return Result.err(selected.error);
  const locations = selected.value.flat();
  const pieces = [];
  for (const { paragraph } of locations) {
    if (!editable(paragraph.content))
      return refuse(
        "Hyperlink edits cannot cut pending review identities or unsupported inline wrappers.",
      );
    const blockId = paragraph.paraId ?? "";
    const from =
      idKey(blockId) === idKey(intent.from.blockId)
        ? intent.from
        : { story: intent.from.story, blockId, offset: 0 };
    const to =
      idKey(blockId) === idKey(intent.to.blockId)
        ? intent.to
        : { story: intent.to.story, blockId, offset: paragraphLength(paragraph) };
    const start = resolveGap({ paragraph, position: from, fallback: "insertion" });
    if (typeof start !== "object") return refuse("Invalid hyperlink range start.", start);
    const end = resolveGap({ paragraph, position: to, fallback: "insertion" });
    if (typeof end !== "object") return refuse("Invalid hyperlink range end.", end);
    if (
      start.offset > end.offset ||
      (start.offset === end.offset && start.zeroWidthBefore > end.zeroWidthBefore)
    )
      return refuse(
        "Hyperlink range endpoints are reversed.",
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
      );
    const [before = [], middle = [], after = []] = partitionContent(paragraph.content, [
      start,
      end,
    ]);
    pieces.push({
      paragraph,
      from,
      to,
      before: asParagraphContent(before),
      middle: asParagraphContent(middle),
      after: asParagraphContent(after),
    });
  }
  // The serializer rejects resolutionJoins repartitioned across hyperlink wrappers.
  // Preserve exact review provenance by refusing before any operation is published.
  if (mode.type === "suggesting")
    return refuse("Hyperlink suggestions require serializable wrapper review provenance.");
  if (intent.type === "insertHyperlink") {
    if (intent.text === "")
      return refuse("A hyperlink insertion needs text.", DOCUMENT_OP_REFUSAL_REASONS.EMPTY_CONTENT);
    const content = [
      hyperlink(intent, [{ type: "run", content: [{ type: "text", text: intent.text }] }]),
    ];
    const deletion = compileEmptyReplacement(document, { from: intent.from, to: intent.to });
    if (deletion.isErr()) return deletion;
    const applied = applyDocumentOps(document, deletion.value.ops);
    if (applied.isErr()) return Result.err(applied.error);
    const at = deletion.value.selection;
    const target = selectedParagraphRuns(applied.value.document, at, at);
    if (target.isErr()) return Result.err(target.error);
    const paragraph = target.value.at(0)?.at(0)?.paragraph;
    if (!paragraph) return refuse("Hyperlink insertion lost its paragraph.");
    const gap = resolveGap({ paragraph, position: at, fallback: "insertion" });
    if (typeof gap !== "object") return refuse("Invalid hyperlink insertion gap.", gap);
    const [before = [], after = []] = partitionContent(paragraph.content, [gap]);
    const replacement = asParagraphContent(before).concat(content, asParagraphContent(after));
    return Result.ok({
      ops: deletion.value.ops.concat([
        {
          type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
          story: at.story,
          blockId: at.blockId,
          expected: paragraph.content,
          content: replacement,
        },
      ]),
      selection: { ...at, offset: at.offset + intent.text.length, zeroWidthBefore: 0 },
    });
  }
  const ops: DocumentOp[] = [];
  for (const piece of pieces) {
    if (piece.from.offset === piece.to.offset) continue;
    const children = linkedChildren(piece.middle, intent);
    const replacement: ParagraphContent[] =
      intent.type === "setHyperlink" ? wrapLinkedChildren(intent, children) : children;
    if (structurallyEqual(piece.middle, replacement)) continue;
    ops.push({
      type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
      story: intent.from.story,
      blockId: piece.from.blockId,
      expected: piece.paragraph.content,
      content: piece.before.concat(replacement, piece.after),
    });
  }
  return Result.ok({ ops, selection: intent.to });
};
