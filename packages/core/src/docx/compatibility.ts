import type {
  BlockContent,
  Document,
  Hyperlink,
  HeaderFooter,
  ParagraphContent,
  Run,
} from "../types/document";
import {
  getFolioDocxCapability,
  type FolioDocxCompatibilityHost,
  type FolioDocxCompatibilityProfile,
  type FolioDocxFeatureCapability,
} from "./capabilities";

export type DocxCompatibilityReason = "opaqueDrawing";

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
  capability: FolioDocxFeatureCapability;
  coverage: {
    host: "covered" | "unknown" | "unverified";
    profile: "covered" | "unknown" | "unverified";
  };
  location: DocxCompatibilityLocation;
};

export type DocxCompatibility = {
  schemaVersion: 1;
  context: DocxCompatibilityContext;
  canSafelyEdit: boolean;
  issues: DocxCompatibilityIssue[];
  reasons: DocxCompatibilityReason[];
  unsupportedContentCount: number;
};

type InspectionLocationContext = {
  blockId?: string;
  part: DocxCompatibilityPart;
  path: string;
};

type RecordIssue = (location: DocxCompatibilityLocation) => void;

const resolveCompatibilityContext = (
  options: InspectDocxCompatibilityOptions,
): DocxCompatibilityContext => ({
  host: options.host ?? "unknown",
  profile: options.profile ?? "unknown",
});

const getCoverageState = (
  value: string,
  coveredValues: readonly string[],
): "covered" | "unknown" | "unverified" => {
  if (value === "unknown") {
    return "unknown";
  }
  return coveredValues.includes(value) ? "covered" : "unverified";
};

export const inspectDocxCompatibility = (
  doc: Document,
  options: InspectDocxCompatibilityOptions = {},
): DocxCompatibility => {
  const context = resolveCompatibilityContext(options);
  const reasons = new Set<DocxCompatibilityReason>();
  const issues: DocxCompatibilityIssue[] = [];

  const record: RecordIssue = (location) => {
    const code = "opaqueDrawing";
    const capability = getFolioDocxCapability(code);
    reasons.add(code);
    issues.push({
      code,
      capability,
      coverage: {
        host: getCoverageState(context.host, capability.hosts),
        profile: getCoverageState(context.profile, capability.profiles),
      },
      location,
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
    schemaVersion: 1,
    context,
    canSafelyEdit: issues.length === 0,
    issues,
    reasons: Array.from(reasons),
    unsupportedContentCount: issues.length,
  };
};

function inspectBlocks(
  blocks: BlockContent[],
  context: InspectionLocationContext & { record: RecordIssue },
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
  context: InspectionLocationContext & { record: RecordIssue },
): void {
  inspectBlocks(headerFooter.content, context);
}

function inspectParagraphContent(
  content: ParagraphContent[],
  context: InspectionLocationContext & { record: RecordIssue },
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
  context: InspectionLocationContext & { record: RecordIssue },
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

function inspectRun(run: Run, context: InspectionLocationContext & { record: RecordIssue }): void {
  for (const [contentIndex, content] of run.content.entries()) {
    if (content.type === "drawing" && content.rawXml) {
      context.record({
        ...(context.blockId === undefined ? {} : { blockId: context.blockId }),
        part: context.part,
        path: `${context.path}.content[${contentIndex}]`,
      });
    }
  }
}
