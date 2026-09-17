/**
 * Whether a parsed package can be edited without losing content on save.
 *
 * Every part walked here is re-serialized from the model at some point: the
 * repack always rebuilds `word/document.xml`, headers and footers replay their
 * captured bytes only while their content fingerprint is unchanged, and note
 * parts are preserved except for edited paragraphs spliced back by `w:paraId`.
 * So an opaque drawing anywhere in this walk is a real risk, while a drawing the
 * serializer replays verbatim is not, whichever part holds it.
 */

import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";

import {
  type BlockContent,
  type Document,
  type DocxConformanceClass,
  type Hyperlink,
  type HeaderFooter,
  type ParagraphContent,
  type Run,
} from "../types/document";
import {
  classifyDrawingSafety,
  DRAWING_SAFETY_CLASSES,
  type DrawingSafetyClass,
} from "./imageRawXml";

export type DocxCompatibilityReason = "opaqueDrawing";

export type FolioDocxCompatibilityHost = "browser" | "server" | "unknown";
export type FolioDocxCompatibilityProfile = DocxConformanceClass;

export type DocxCompatibilityContext = {
  host: FolioDocxCompatibilityHost;
  profile: FolioDocxCompatibilityProfile;
};

export type InspectDocxCompatibilityOptions = Partial<DocxCompatibilityContext>;

export type DocxCompatibilityPart =
  | { type: "document" }
  | { type: "header" | "footer"; relationshipId: string }
  | { type: "footnote" | "endnote"; id: number };

export type DocxCompatibilityLocation = {
  part: DocxCompatibilityPart;
  path: string;
  blockId?: string;
};

export type DocxCompatibilityIssue = {
  code: DocxCompatibilityReason;
  location: DocxCompatibilityLocation;
};

/** A drawing and the fate the run serializer holds for it. */
export type DocxDrawingClassification = {
  class: DrawingSafetyClass;
  location: DocxCompatibilityLocation;
};

/**
 * Schema 2 replaced the per-drawing boolean judgement with
 * {@link DocxDrawingClassification} and stopped reporting a replayable drawing
 * as an issue; `issues` now holds opaque drawings only.
 */
export type DocxCompatibility = {
  schemaVersion: 2;
  context: DocxCompatibilityContext;
  canSafelyEdit: boolean;
  /** Every drawing the walk reached, classified. */
  drawings: DocxDrawingClassification[];
  /** The blocking subset: drawings a save can neither replay nor regenerate. */
  issues: DocxCompatibilityIssue[];
  reasons: DocxCompatibilityReason[];
  unsupportedContentCount: number;
};

type InspectionLocationContext = {
  blockId?: string;
  part: DocxCompatibilityPart;
  path: string;
};

type RecordDrawing = (drawing: DocxDrawingClassification) => void;

const resolveCompatibilityContext = (
  doc: Document,
  options: InspectDocxCompatibilityOptions,
): DocxCompatibilityContext => ({
  host: options.host ?? "unknown",
  profile: options.profile ?? doc.package.conformanceClass ?? DOCX_CONFORMANCE_CLASSES.UNKNOWN,
});

export const inspectDocxCompatibility = (
  doc: Document,
  options: InspectDocxCompatibilityOptions = {},
): DocxCompatibility => {
  const context = resolveCompatibilityContext(doc, options);
  const reasons = new Set<DocxCompatibilityReason>();
  const drawings: DocxDrawingClassification[] = [];
  const issues: DocxCompatibilityIssue[] = [];

  const record: RecordDrawing = (drawing) => {
    drawings.push(drawing);
    if (drawing.class !== DRAWING_SAFETY_CLASSES.OPAQUE) {
      return;
    }
    const code = "opaqueDrawing";
    reasons.add(code);
    issues.push({
      code,
      location: drawing.location,
    });
  };

  inspectBlocks(doc.package.document.content, {
    part: { type: "document" },
    path: "package.document.content",
    record,
  });
  for (const [relationshipId, header] of doc.package.headers?.entries() ?? []) {
    inspectHeaderFooter(header, {
      part: { type: "header", relationshipId },
      path: `package.headers.get(${JSON.stringify(relationshipId)}).content`,
      record,
    });
  }
  for (const [relationshipId, footer] of doc.package.footers?.entries() ?? []) {
    inspectHeaderFooter(footer, {
      part: { type: "footer", relationshipId },
      path: `package.footers.get(${JSON.stringify(relationshipId)}).content`,
      record,
    });
  }
  for (const footnote of doc.package.footnotes ?? []) {
    inspectBlocks(footnote.content, {
      part: { type: "footnote", id: footnote.id },
      path: `package.footnotes[id=${footnote.id}].content`,
      record,
    });
  }
  for (const endnote of doc.package.endnotes ?? []) {
    inspectBlocks(endnote.content, {
      part: { type: "endnote", id: endnote.id },
      path: `package.endnotes[id=${endnote.id}].content`,
      record,
    });
  }

  return {
    schemaVersion: 2,
    context,
    canSafelyEdit: issues.length === 0,
    drawings,
    issues,
    reasons: Array.from(reasons),
    unsupportedContentCount: issues.length,
  };
};

function inspectBlocks(
  blocks: BlockContent[],
  context: InspectionLocationContext & { record: RecordDrawing },
): void {
  for (const [blockIndex, block] of blocks.entries()) {
    const blockPath = `${context.path}[${blockIndex}]`;
    if (block.type === "paragraph") {
      inspectParagraphContent(block.content, {
        ...(block.paraId === undefined ? {} : { blockId: block.paraId }),
        part: context.part,
        path: `${blockPath}.content`,
        record: context.record,
      });
      continue;
    }

    if (block.type === "table") {
      for (const [rowIndex, row] of block.rows.entries()) {
        for (const [cellIndex, cell] of row.cells.entries()) {
          inspectBlocks(cell.content, {
            part: context.part,
            path: `${blockPath}.rows[${rowIndex}].cells[${cellIndex}].content`,
            record: context.record,
          });
        }
      }
      continue;
    }

    inspectBlocks(block.content, {
      part: context.part,
      path: `${blockPath}.content`,
      record: context.record,
    });
  }
}

function inspectHeaderFooter(
  headerFooter: HeaderFooter,
  context: InspectionLocationContext & { record: RecordDrawing },
): void {
  inspectBlocks(headerFooter.content, context);
}

function inspectParagraphContent(
  content: ParagraphContent[],
  context: InspectionLocationContext & { record: RecordDrawing },
): void {
  for (const [itemIndex, item] of content.entries()) {
    const itemContext = {
      ...context,
      path: `${context.path}[${itemIndex}]`,
    };
    if (item.type === "run") {
      inspectRun(item, itemContext);
      continue;
    }

    if (item.type === "hyperlink") {
      inspectHyperlink(item, itemContext);
      continue;
    }

    if (item.type === "inlineSdt") {
      inspectParagraphContent(item.content, {
        ...itemContext,
        path: `${itemContext.path}.content`,
      });
      continue;
    }

    if (
      item.type === "insertion" ||
      item.type === "deletion" ||
      item.type === "moveFrom" ||
      item.type === "moveTo"
    ) {
      inspectParagraphContent(item.content, {
        ...itemContext,
        path: `${itemContext.path}.content`,
      });
      continue;
    }

    if (item.type === "simpleField") {
      inspectParagraphContent(item.content, {
        ...itemContext,
        path: `${itemContext.path}.content`,
      });
      continue;
    }

    if (item.type === "complexField") {
      for (const [runIndex, run] of item.fieldCode.entries()) {
        inspectRun(run, {
          ...itemContext,
          path: `${itemContext.path}.fieldCode[${runIndex}]`,
        });
      }
      for (const [runIndex, run] of item.fieldResult.entries()) {
        inspectRun(run, {
          ...itemContext,
          path: `${itemContext.path}.fieldResult[${runIndex}]`,
        });
      }
    }
  }
}

function inspectHyperlink(
  hyperlink: Hyperlink,
  context: InspectionLocationContext & { record: RecordDrawing },
): void {
  for (const [childIndex, child] of hyperlink.children.entries()) {
    if (child.type === "run") {
      inspectRun(child, {
        ...context,
        path: `${context.path}.children[${childIndex}]`,
      });
    }
  }
}

function inspectRun(
  run: Run,
  context: InspectionLocationContext & { record: RecordDrawing },
): void {
  for (const [contentIndex, content] of run.content.entries()) {
    if (content.type !== "drawing") {
      continue;
    }
    context.record({
      class: classifyDrawingSafety(content),
      location: {
        ...(context.blockId === undefined ? {} : { blockId: context.blockId }),
        part: context.part,
        path: `${context.path}.content[${contentIndex}]`,
      },
    });
  }
}
