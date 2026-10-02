/** Operation × touched document feature × selection coverage shared by the
 * consumer flows and the editor-command conformance suite. */

export const DOCUMENT_FEATURES = [
  "plain-table",
  "merged-cells",
  "nested-table",
  "bullet-list",
  "numbered-list",
  "multi-level-list",
  "field",
  "content-control",
  "footnote",
  "endnote",
  "comment-anchor",
  "tracked-change",
  "section-break",
  "header-footer",
  "heading",
  "none",
] as const;
export type DocumentFeature = (typeof DOCUMENT_FEATURES)[number];

export const SELECTION_TYPES = [
  "caret",
  "paragraph-range",
  "cross-paragraph",
  "whole-document",
  "cell-selection",
  "node-selection",
  "none",
] as const;
export type SelectionType = (typeof SELECTION_TYPES)[number];

export type FeatureCell = { operation: string; feature: DocumentFeature; selection: SelectionType };
export type FeatureCounts = Record<string, number>;
export type FeatureCoverage = { version: 1; cells: FeatureCounts; operations: string[] };

const DELIMITER = " | ";
const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** Validate and snapshot only generation inputs; derived report fields are discarded. */
export const parseFeatureCoverage = (value: unknown): FeatureCoverage => {
  if (
    !isRecord(value) ||
    value["version"] !== 1 ||
    !isRecord(value["cells"]) ||
    !Array.isArray(value["operations"])
  ) {
    throw new TypeError("feature coverage needs version 1, cells, and operations");
  }
  const operations = value["operations"].filter(
    (operation): operation is string => typeof operation === "string" && operation.length > 0,
  );
  if (
    operations.length !== value["operations"].length ||
    new Set(operations).size !== operations.length
  ) {
    throw new TypeError("feature coverage operations must be distinct nonempty strings");
  }
  const cells: FeatureCounts = {};
  for (const [key, count] of Object.entries(value["cells"])) {
    const [operation, feature, selection, extra] = key.split(DELIMITER);
    if (
      !operation ||
      !operations.includes(operation) ||
      !(DOCUMENT_FEATURES as readonly unknown[]).includes(feature) ||
      !(SELECTION_TYPES as readonly unknown[]).includes(selection) ||
      extra !== undefined ||
      typeof count !== "number" ||
      !Number.isSafeInteger(count) ||
      count < 0
    ) {
      throw new TypeError(`invalid feature coverage cell ${key}`);
    }
    cells[key] = count;
  }
  return { version: 1, cells, operations };
};

/** Generation's operation weights must derive from the flow's recorded input. */
export const featureOperationHits = (report: FeatureCoverage): FeatureCounts => {
  const hits: FeatureCounts = {};
  for (const [key, count] of Object.entries(report.cells)) {
    const operation = key.split(DELIMITER).at(0);
    if (!operation) throw new TypeError("feature coverage cell has no operation");
    hits[operation] = (hits[operation] ?? 0) + count;
  }
  return hits;
};

export const featureCellKey = ({ operation, feature, selection }: FeatureCell): string =>
  [operation, feature, selection].join(DELIMITER);

export const addFeatureHit = (report: FeatureCoverage, cell: FeatureCell): void => {
  const key = featureCellKey(cell);
  report.cells[key] = (report.cells[key] ?? 0) + 1;
  if (!report.operations.includes(cell.operation)) report.operations.push(cell.operation);
};

export const emptyFeatureCoverage = (): FeatureCoverage => ({
  version: 1,
  cells: {},
  operations: [],
});

export const mergeFeatureCoverage = (reports: readonly FeatureCoverage[]): FeatureCoverage => {
  const merged = emptyFeatureCoverage();
  for (const report of reports) {
    for (const operation of report.operations) {
      if (!merged.operations.includes(operation)) merged.operations.push(operation);
    }
    for (const [key, count] of Object.entries(report.cells)) {
      merged.cells[key] = (merged.cells[key] ?? 0) + count;
    }
  }
  merged.operations.sort();
  return merged;
};

export const emptyFeatureCells = (report: FeatureCoverage): string[] =>
  report.operations.flatMap((operation) =>
    DOCUMENT_FEATURES.flatMap((feature) =>
      SELECTION_TYPES.flatMap((selection) => {
        const key = featureCellKey({ operation, feature, selection });
        return report.cells[key] ? [] : [key];
      }),
    ),
  );

export const summarizeFeatureCoverage = (report: FeatureCoverage) => ({
  ...report,
  emptyCells: emptyFeatureCells(report),
});

/** Bounded inverse-frequency choice, with no extra random draws when disabled. */
export const weightedChoice = <T>(
  items: readonly T[],
  random: { next: () => number },
  weight: (item: T) => number,
): T => {
  const weights = items.map((item) => Math.max(1, Math.min(9, weight(item))));
  const total = weights.reduce((sum, value) => sum + value, 0);
  let draw = random.next() * total;
  for (let index = 0; index < items.length; index += 1) {
    draw -= weights[index] ?? 0;
    if (draw < 0) return items[index]!;
  }
  return items.at(-1)!;
};

export const gapWeight = (hits: number): number => 1 + Math.floor(8 / (1 + hits));

const SHAPE_FEATURES: Readonly<Record<string, DocumentFeature>> = {
  table: "plain-table",
  "list-bullet": "bullet-list",
  "list-decimal": "numbered-list",
  "style-numbering": "heading",
  "outline-level": "heading",
  "content-control": "content-control",
  footnote: "footnote",
  comment: "comment-anchor",
  "tracked-insertion": "tracked-change",
  "tracked-deletion": "tracked-change",
  "paragraph-property-change": "tracked-change",
  "run-property-change": "tracked-change",
  sections: "section-break",
};

/** Shape metadata is document-wide; ambiguous combinations are not attributed to its focus. */
export const shapeFeatureSignature = (features: readonly string[]): DocumentFeature[] => {
  const found = new Set<DocumentFeature>();
  for (const feature of features) {
    const mapped = SHAPE_FEATURES[feature];
    if (mapped) found.add(mapped);
  }
  return found.size === 1 ? [...found] : ["none"];
};

export const placementSelection = (placement: string, isCellSelection = false): SelectionType => {
  if (isCellSelection) return "cell-selection";
  if (placement === "node") return "node-selection";
  if (placement.startsWith("caret-")) return "caret";
  if (placement === "word" || placement === "paragraph") return "paragraph-range";
  if (placement === "cross-paragraph") return "cross-paragraph";
  if (placement === "document") return "whole-document";
  return "none";
};

export const operationSelection = (operation: Record<string, unknown>): SelectionType => {
  const range = operation["range"];
  if (typeof range === "object" && range !== null) return "paragraph-range";
  if (typeof operation["find"] === "string" || typeof operation["quote"] === "string") {
    return "paragraph-range";
  }
  if (typeof operation["offset"] === "number") return "caret";
  return "none";
};

export const generatedSelection = (type: string): SelectionType => {
  if (
    ["replaceRange", "formatRange", "commentOnRange", "replaceInBlock", "commentOnBlock"].includes(
      type,
    )
  ) {
    return "paragraph-range";
  }
  if (type === "splitBlock") return "caret";
  return "none";
};

type TargetSignature = {
  kind?: string;
  headingLevel?: number;
  listLevel?: number;
  listReference?: unknown;
  displayLabel?: string;
  table?: unknown;
};

/** Classify only features observable at the touched target. */
export const targetFeatureSignature = (
  target: TargetSignature | undefined,
  known: ReadonlySet<string>,
  story: string,
): DocumentFeature[] => {
  const found = new Set<DocumentFeature>();
  const table = target?.table;
  if (typeof table === "object" && table !== null) {
    found.add("plain-table");
    if (
      ("columnSpan" in table && typeof table.columnSpan === "number" && table.columnSpan > 1) ||
      ("rowSpan" in table && typeof table.rowSpan === "number" && table.rowSpan > 1)
    )
      found.add("merged-cells");
    if (
      "outerTableIndex" in table &&
      typeof table.outerTableIndex === "number" &&
      "tableIndex" in table &&
      typeof table.tableIndex === "number" &&
      table.outerTableIndex !== table.tableIndex
    )
      found.add("nested-table");
  }
  if (target?.kind === "listItem" || target?.listReference !== undefined) {
    found.add(target.displayLabel?.match(/^\d/u) ? "numbered-list" : "bullet-list");
    if ((target.listLevel ?? 0) > 0) found.add("multi-level-list");
  }
  if (target?.kind === "heading" || target?.headingLevel !== undefined) found.add("heading");
  if (known.has("field")) found.add("field");
  if (known.has("contentControl")) found.add("content-control");
  if (story === "footnote") found.add("footnote");
  if (story === "endnote") found.add("endnote");
  if (known.has("commentAnchor")) found.add("comment-anchor");
  if (known.has("pendingRevision")) found.add("tracked-change");
  if (known.has("sectionCarrier")) found.add("section-break");
  if (story === "header" || story === "footer") found.add("header-footer");
  return found.size > 0 ? [...found] : ["none"];
};
