import { mergeRunContent } from "@stll/docx-core/ops";

import {
  removeCommentMarkersThroughTextBoxes,
  visitCommentMarkers,
} from "../docx/commentAnchorIndex";
import { paragraphFormattingWithAuthoredIndentation } from "../docx/paragraphPropertySource";
import {
  BLOCK_TREE_DESCENT,
  visitBlockTreeRecords,
  visitInlineContentSlots,
  visitParagraphRuns,
  withInlineContentRewritten,
} from "../docx/paragraphTraversal";
import {
  consolidateParagraphContent,
  formattingEquals,
} from "../docx/runConsolidator";
import { serializeDocumentPart } from "../docx/serializer/documentSerializer";
import { WORDPROCESSINGML_NAMESPACE_URIS } from "../docx/xmlParser";
import type {
  BlockContent,
  Paragraph,
  ParagraphContent,
  ParagraphFormatting,
} from "../types/document";

const protectionKinds = ["note", "drawing", "field", "comment"] as const;
type ProtectionKind = (typeof protectionKinds)[number];
export type Inventory = Record<ProtectionKind, Map<string, number>>;

const add = (items: Map<string, number>, key: string): void => {
  items.set(key, (items.get(key) ?? 0) + 1);
};

/**
 * The protected content `after` holds as a pending deletion and `before` did
 * not: an edit the reviewer sees as deleted, which Accept All may remove.
 */
export const newlyDeletedProtectedContent = (
  before: Inventory,
  after: Inventory,
): Inventory => {
  const inventory: Inventory = {
    note: new Map(),
    drawing: new Map(),
    field: new Map(),
    comment: new Map(),
  };
  for (const kind of protectionKinds) {
    for (const [key, count] of after[kind]) {
      const added = count - (before[kind].get(key) ?? 0);
      if (added > 0) inventory[kind].set(key, added);
    }
  }
  return inventory;
};

export const missingProtectedContent = (
  expected: Inventory,
  actual: Inventory,
  retiredCommentIds: readonly number[],
  deleted?: Inventory,
): ProtectionKind | null => {
  const retiredCommentKeys = new Set(
    retiredCommentIds.flatMap((id) => [String(id), `reference:${String(id)}`]),
  );
  return (
    protectionKinds.find((kind) => {
      for (const [key, count] of expected[kind]) {
        if (kind === "comment" && retiredCommentKeys.has(key)) continue;
        const removed = deleted?.[kind].get(key) ?? 0;
        if ((actual[kind].get(key) ?? 0) + removed < count) return true;
      }
      return false;
    }) ?? null
  );
};

// A resolved paragraph break can change the package's editing-session owner.
const PARAGRAPH_SESSION_ATTRIBUTES: ReadonlySet<string> = new Set([
  "rsidR",
  "rsidRPr",
  "rsidDel",
  "rsidP",
  "rsidRDefault",
]);

/**
 * `content` with adjacent runs of the same properties as one run: where a
 * run boundary falls is no content or formatting of the paragraph, so a
 * comment's markers taken out between two runs (around a line break, say)
 * leave the paragraph the original was. A run carrying a property revision
 * keeps its boundaries, which delimit the revision.
 */
const withoutRunBoundaries = (
  content: Paragraph["content"],
): Paragraph["content"] => {
  const joined: Paragraph["content"] = [];
  for (const item of content) {
    const previous = joined.at(-1);
    if (
      previous?.type === "run" &&
      item.type === "run" &&
      previous.propertyChanges === undefined &&
      item.propertyChanges === undefined &&
      formattingEquals(previous.formatting, item.formatting) &&
      JSON.stringify(previous.preservedAttributes) ===
        JSON.stringify(item.preservedAttributes)
    ) {
      joined[joined.length - 1] = {
        ...previous,
        content: mergeRunContent(previous.content, item.content),
      };
      continue;
    }
    joined.push(item);
  }
  return joined;
};

const authoredBlock = (block: BlockContent): string => {
  // A native reject restores authored properties, while the live model still
  // projects list indentation. Read the authored values before cloning drops
  // the provenance that distinguishes inherited indentation from direct edits.
  const formatting: (ParagraphFormatting | undefined)[] = [];
  visitBlockTreeRecords([block], (record) => {
    if (record.type === "paragraph")
      formatting.push(paragraphFormattingWithAuthoredIndentation(record));
    return BLOCK_TREE_DESCENT.descend;
  });
  const authored = structuredClone(block);
  let paragraphIndex = 0;
  visitBlockTreeRecords([authored], (record) => {
    if (record.type === "paragraph") {
      const properties = formatting.at(paragraphIndex);
      paragraphIndex += 1;
      if (properties === undefined) delete record.formatting;
      else record.formatting = properties;
      delete record.paraId;
      delete record.textId;
      delete record.renderedPageBreakBefore;
      delete record.listRendering;
      delete record.reviewCarrier;
      if (record.preservedAttributes !== undefined) {
        record.preservedAttributes = record.preservedAttributes.filter(
          ({ name, namespace }) =>
            !WORDPROCESSINGML_NAMESPACE_URIS.has(namespace ?? "") ||
            !PARAGRAPH_SESSION_ATTRIBUTES.has(name),
        );
      }
      if (record.formatting && Object.keys(record.formatting).length === 0)
        delete record.formatting;
      record.content = withoutRunBoundaries(
        consolidateParagraphContent(record.content),
      );
    }
    return BLOCK_TREE_DESCENT.descend;
  });
  return serializeDocumentPart({
    package: { document: { content: [authored] } },
  }).xml;
};

export const sameOriginalBlocks = (
  before: readonly BlockContent[],
  after: readonly BlockContent[],
  addedCommentIds: ReadonlySet<number>,
): boolean => {
  const restored = removeCommentMarkersThroughTextBoxes(after, addedCommentIds);
  return (
    before.length === restored.length &&
    before.every((block, index) => {
      const other = restored.at(index);
      return (
        other !== undefined &&
        (block === other || authoredBlock(block) === authoredBlock(other))
      );
    })
  );
};

const modelInventory = (blocks: readonly BlockContent[]): Inventory => {
  const inventory: Inventory = {
    note: new Map(),
    drawing: new Map(),
    field: new Map(),
    comment: new Map(),
  };
  visitBlockTreeRecords(blocks, (record) => {
    if (record.type !== "paragraph") return BLOCK_TREE_DESCENT.descend;
    visitInlineContentSlots(record, ({ item }) => {
      if (item.type === "simpleField" || item.type === "complexField")
        add(inventory.field, item.instruction);
    });
    visitParagraphRuns(record, ({ content }) => {
      for (const item of content) {
        switch (item.type) {
          case "footnoteRef":
          case "endnoteRef":
            inventory.note.set(JSON.stringify([item.type, item.id]), 1);
            break;
          case "drawing":
            add(
              inventory.drawing,
              JSON.stringify([
                "drawing",
                item.image.id,
                item.image.rId ?? item.rawXml ?? item.image.src,
              ]),
            );
            break;
          case "shape":
            add(
              inventory.drawing,
              JSON.stringify([
                "shape",
                item.shape.id,
                item.shape.name,
                item.shape.shapeType,
              ]),
            );
            break;
        }
      }
    });
    return BLOCK_TREE_DESCENT.descend;
  });
  visitCommentMarkers(blocks, ({ item }) => {
    inventory.comment.set(String(item.id), 1);
    if (item.type === "commentReference")
      inventory.comment.set(`reference:${String(item.id)}`, 1);
  });
  return inventory;
};

/** The notes, drawings and fields `blocks` hold inside tracked deletions. */
const deletedModelInventory = (blocks: readonly BlockContent[]): Inventory => {
  const deleted: ParagraphContent[] = [];
  const collect = (content: readonly ParagraphContent[]): void => {
    for (const item of content) {
      if (item.type === "deletion") deleted.push(...item.content);
      else
        withInlineContentRewritten(item, (children) => {
          collect(children);
          return children;
        });
    }
  };
  visitBlockTreeRecords(blocks, (record) => {
    if (record.type === "paragraph") collect(record.content);
    return BLOCK_TREE_DESCENT.descend;
  });
  const inventory = modelInventory([{ type: "paragraph", content: deleted }]);
  inventory.comment.clear();
  return inventory;
};

/**
 * The protected content the edited blocks hold as a new tracked deletion,
 * which Accept All removes as the reviewer sees it.
 */
export const newlyDeletedProtectedBlocks = (
  before: readonly BlockContent[],
  after: readonly BlockContent[],
): Inventory =>
  newlyDeletedProtectedContent(
    deletedModelInventory(before),
    deletedModelInventory(after),
  );

/** The kind of protected content missing from the accepted blocks. */
export const missingProtectedBlockKind = (
  before: readonly BlockContent[],
  after: readonly BlockContent[],
  retiredCommentIds: readonly number[] = [],
  deleted?: Inventory,
): ProtectionKind | null =>
  missingProtectedContent(
    modelInventory(before),
    modelInventory(after),
    retiredCommentIds,
    deleted,
  );
