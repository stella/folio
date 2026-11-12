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
//   el.dataset[editing ? "pmStart" : "decorationStart"]
//   const { pmStart } = el.dataset
//   { pmPos: "sdtPmPos" }   (a key table naming a dataset-only key)
//   el.getAttribute("data-pm-start")
//   root.querySelectorAll("span[data-pm-start][data-pm-end]")
//   `[data-pm-start="${pos}"]`
//
// Not flagged (no scope analysis): `const data = el.dataset; data.pmStart`.
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

// Dataset keys no layout field shares, so any spelling of them is a read or
// write of the painted attribute.
const DATASET_ONLY_KEYS: ReadonlySet<string> = new Set([
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

// The string literals a computed key can evaluate to: `"pmStart"`,
// `editing ? "pmStart" : "decorationStart"`, `key ?? "pmEnd"`.
const literalKeysOf = (node: unknown): string[] => {
  if (!isNode(node)) return [];
  if (node.type === "Literal")
    return typeof node["value"] === "string" ? [node["value"]] : [];
  if (node.type === "ConditionalExpression")
    return [
      ...literalKeysOf(node["consequent"]),
      ...literalKeysOf(node["alternate"]),
    ];
  if (node.type === "LogicalExpression")
    return [...literalKeysOf(node["left"]), ...literalKeysOf(node["right"])];
  return [];
};

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

// The literal is the key of a `dataset[...]` access the MemberExpression
// visitor already reports.
const isDatasetKey = (node: AstNode): boolean => {
  const parent = node["parent"];
  return (
    isNode(parent) &&
    parent.type === "MemberExpression" &&
    parent["property"] === node &&
    isDataset(parent["object"])
  );
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
            if (key !== null && DATASET_KEYS.has(key)) {
              report(node);
              return;
            }
            if (
              node["computed"] === true &&
              isDataset(node["object"]) &&
              literalKeysOf(node["property"]).some((k) => DATASET_KEYS.has(k))
            )
              report(node);
          },
          VariableDeclarator: (node: unknown) => {
            if (!isNode(node) || !isDataset(node["init"])) return;
            const pattern = node["id"];
            if (!isNode(pattern) || pattern.type !== "ObjectPattern") return;
            const properties = pattern["properties"];
            if (!Array.isArray(properties)) return;
            for (const property of properties) {
              if (!isNode(property) || property.type !== "Property") continue;
              const key = property["key"];
              const name = isNode(key)
                ? key.type === "Identifier"
                  ? key["name"]
                  : key["value"]
                : undefined;
              if (typeof name === "string" && DATASET_KEYS.has(name))
                report(property);
            }
          },
          Literal: (node: unknown) => {
            if (!isNode(node)) return;
            const value = node["value"];
            if (
              namesAttribute(value) ||
              (typeof value === "string" &&
                DATASET_ONLY_KEYS.has(value) &&
                !isDatasetKey(node))
            )
              report(node);
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
