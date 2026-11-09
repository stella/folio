// Read and write a painted element's editor positions through one owner:
// `packages/core/src/utils/paintedPositions.ts`.
//
// The painter marks what it draws with the editor document positions it
// shows, and hit testing, caret and selection drawing, scrolling and the hosts'
// pointer handling read them back. How a position is stored on an element is
// the owner's decision alone, so storing it differently (for instance relative
// to the painted block) changes one module instead of every reader. A reader
// that spells the attribute out keeps reading the old storage after the
// change, and nothing fails until a click lands in the wrong place.
//
// Flagged, outside the owner:
//   el.dataset["pmStart"], el.dataset.pmEnd, el.dataset["tablePmStart"]
//   el.getAttribute("data-pm-start")
//   root.querySelectorAll("span[data-pm-start][data-pm-end]")
//   `[data-pm-start="${pos}"]`
//
// Not flagged:
//   readPaintedStart(el), hasPaintedRange(el), writePaintedRange(el, start, end)
//   `span${PAINTED_RANGE_SELECTOR}`, paintedStartSelector(pos)
//   fragment.pmStart (a layout value, not a painted attribute)

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  filename: string;
  report: (descriptor: {
    node: unknown;
    messageId: "paintedPositionOutsideOwner";
  }) => void;
};

const OWNER = "packages/core/src/utils/paintedPositions.ts";

const DATASET_KEYS: ReadonlySet<string> = new Set([
  "pmStart",
  "pmEnd",
  "tablePmStart",
  "sdtPmPos",
]);

const ATTRIBUTE_NAMES = [
  "data-pm-start",
  "data-pm-end",
  "data-table-pm-start",
  "data-sdt-pm-pos",
];

const isNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof value.type === "string";

const isDataset = (node: unknown): boolean =>
  isNode(node) &&
  node.type === "MemberExpression" &&
  node.computed !== true &&
  isNode(node.property) &&
  node.property.type === "Identifier" &&
  node.property["name"] === "dataset";

// `x.dataset.pmStart` or `x.dataset["pmStart"]`.
const datasetKeyOf = (node: AstNode): string | null => {
  if (!isDataset(node["object"])) return null;
  const property = node["property"];
  if (!isNode(property)) return null;
  if (node["computed"] !== true)
    return property.type === "Identifier" &&
      typeof property["name"] === "string"
      ? property["name"]
      : null;
  return property.type === "Literal" && typeof property["value"] === "string"
    ? property["value"]
    : null;
};

const namesAttribute = (text: unknown): boolean =>
  typeof text === "string" &&
  ATTRIBUTE_NAMES.some((name) => text.includes(name));

export default {
  meta: { name: "folio-painted-positions" },
  rules: {
    "position-attribute-owner": {
      meta: {
        type: "problem",
        messages: {
          paintedPositionOutsideOwner:
            "A painted element's editor positions are stored by " +
            "`packages/core/src/utils/paintedPositions.ts` alone. Read them " +
            "with readPaintedStart / readPaintedEnd / readPaintedTableStart / " +
            "hasPaintedRange, write them with writePaintedRange, and select " +
            "elements with PAINTED_RANGE_SELECTOR, PAINTED_START_SELECTOR or " +
            "paintedStartSelector(pos), so a change of storage reaches every " +
            "reader.",
        },
      },
      create(context: RuleContext) {
        if (context.filename.replaceAll("\\", "/").endsWith(OWNER)) return {};
        const report = (node: unknown) =>
          context.report({ node, messageId: "paintedPositionOutsideOwner" });
        return {
          MemberExpression: (node: unknown) => {
            if (!isNode(node)) return;
            const key = datasetKeyOf(node);
            if (key !== null && DATASET_KEYS.has(key)) report(node);
          },
          Literal: (node: unknown) => {
            if (isNode(node) && namesAttribute(node["value"])) report(node);
          },
          TemplateElement: (node: unknown) => {
            if (!isNode(node)) return;
            const value = node["value"];
            if (
              typeof value === "object" &&
              value !== null &&
              "raw" in value &&
              namesAttribute(value.raw)
            )
              report(node);
          },
        };
      },
    },
  },
};
