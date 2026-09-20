// Keep the reading of an OOXML reserved value in one place.
//
// A reserved value is a value a slot accepts whose meaning is not the value
// itself: `w:numId` 0 names no numbering definition, `w:outlineLvl` 9 means
// body text, `w:u w:val="none"` cancels an inherited underline, a `w:tcW`
// number is meaningless under `w:type="auto"`. Every leak folio has had here
// came from a duplicated reader whose copies drifted apart, never from one
// reader that forgot: four shading parsers where one keeps `auto`, two
// `textFormattingToMarks` where one checks `"none"`, three table-width
// resolvers where one treats `auto` as `dxa`.
//
// So each model field records its decision in
// `specifications/reserved-values`, naming the one function that owns the
// read, and this rule flags a bare comparison against that field's sentinel
// anywhere else. The field/sentinel pairs and the owning modules are read from
// the registry at lint time; nothing here mirrors them.
//
// Flagged:
//   if (paragraph.formatting.outlineLevel === 9) { ... }
//                                          ^^^^^^^^^ outside headingCollector.ts
//   const isMerged = cell.vMerge === "continue";
//
// Safe:
//   if (!isNumberingReference(numPr.numId)) { ... }   // through the owning reader
//   if (op.action === "clear") { ... }                // not a model field
//   if (widthType === "dxa") { ... }                  // not a reserved value

// Explicit `.ts`: plain `bunx oxlint` loads this plugin through Node, whose ESM
// resolver refuses an extensionless specifier. See
// `scripts/oxlint-config-loaders.test.ts`.
import { reservedValueEntries } from "../specifications/reserved-values/registry.ts";

type AstNode = Record<string, unknown> & { type: string };

type ReservedCompareContext = {
  filename: string;
  report: (descriptor: {
    node: unknown;
    messageId: "bareReservedCompare";
    data: { field: string; literal: string; reader: string };
  }) => void;
};

// Keys that never lead to child nodes (or lead back up the tree).
const SKIP_KEYS = new Set(["parent", "loc", "range", "start", "end", "type"]);

const EQUALITY_OPERATORS = new Set(["===", "!==", "==", "!="]);
const RELATIONAL_OPERATORS = new Set(["<", "<=", ">", ">="]);

/**
 * Sentinels that name a structural rule rather than a literal: an absent
 * `w:vMerge w:val`, a `w:ind` carrying `@w:hanging` and `@w:firstLine` at once,
 * a `w:tcW` number under `w:type="auto"`, a style id nothing resolves. There is
 * no literal for a comparison to be bare against.
 */
const STRUCTURAL_SENTINELS = new Set([
  "absent",
  "both-present",
  "meaningless-under-auto",
  "superseded-by-flag",
  "unresolvable-styleid",
]);

/**
 * Field names that are also everyday local-variable names. The rule is
 * syntactic, so a bare `id` or `start` reads the same whether it holds a note
 * id or a string index; requiring a property access (`note.id`, `level.start`)
 * keeps the domain reads and drops the rest. The four `w:tblLook` positions are
 * here for the same reason: a table walk names its bounds `firstRow` and
 * `lastColumn` too, and `firstRow === 0` is then an index test, not a flag
 * read. `separator` is the column-rule toggle here and the index of a `:` in
 * half the string parsers. Every other registry field name (`numId`,
 * `outlineLevel`, `gridSpan`, `vMerge`, `suffix`, `leader`) is distinctive
 * enough to stand alone.
 */
const AMBIGUOUS_BARE_FIELDS = new Set([
  "id",
  "start",
  "firstRow",
  "lastRow",
  "firstColumn",
  "lastColumn",
  "separator",
]);

type FieldRule = { literals: Set<string>; numbers: Set<number>; readers: Set<string> };

/** Field name -> the sentinels recorded for it and the modules that own them. */
const buildFieldRules = (): Map<string, FieldRule> => {
  const rules = new Map<string, FieldRule>();
  for (const { field, disposition } of reservedValueEntries()) {
    if (disposition === "no-reserved-value" || disposition.disposition !== "reader-owned") {
      continue;
    }
    const sentinels = disposition.sentinel
      .split("|")
      .filter((sentinel) => !STRUCTURAL_SENTINELS.has(sentinel));
    if (sentinels.length === 0) {
      continue;
    }
    let rule = rules.get(field);
    if (rule === undefined) {
      rule = { literals: new Set(), numbers: new Set(), readers: new Set() };
      rules.set(field, rule);
    }
    for (const sentinel of sentinels) {
      rule.literals.add(sentinel);
      const asNumber = Number(sentinel);
      if (sentinel !== "" && Number.isInteger(asNumber)) {
        rule.numbers.add(asNumber);
      }
    }
    rule.readers.add(disposition.reader);
  }
  return rules;
};

const FIELD_RULES = buildFieldRules();

/** Repo-relative module paths that own at least one reserved value. */
const READER_MODULES = new Set(
  [...FIELD_RULES.values()].flatMap((rule) =>
    [...rule.readers].map((reader) => reader.split("#").at(0) ?? reader),
  ),
);

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const forEachChild = (node: AstNode, visit: (child: AstNode) => void): void => {
  for (const [key, value] of Object.entries(node)) {
    if (SKIP_KEYS.has(key)) {
      continue;
    }
    if (Array.isArray(value)) {
      for (const item of value) {
        if (isAstNode(item)) {
          visit(item);
        }
      }
      continue;
    }
    if (isAstNode(value)) {
      visit(value);
    }
  }
};

const normalized = (filename: string): string => filename.replaceAll("\\", "/");

/**
 * The module that declares a reserved value is where the comparison belongs.
 * Exempt the whole file rather than the one function: a reader routinely spans
 * a private helper and its exported entry point.
 */
const isOwningModule = (filename: string): boolean => {
  const path = normalized(filename);
  return [...READER_MODULES].some((module) => path.endsWith(module));
};

/** The registry itself spells every sentinel out; it compares none of them. */
const isRegistryModule = (filename: string): boolean =>
  normalized(filename).includes("/specifications/reserved-values/");

/**
 * Receivers that never hold a model value, whatever the field is called.
 *
 * `type` is a field name the model shares with half the editor's own tagged
 * unions, and `"none"` is a value both vocabularies use: `fill.type === "none"`
 * is a shape with no fill, `target.type === "none"` is the undo stack with no
 * focused view. The receiver is the only syntactic signal that tells them
 * apart, and `dataset` is the DOM's own string bag.
 */
const NON_DOMAIN_RECEIVERS = new Set(["dataset", "source", "target"]);

/** The name the receiver of a member access is reached through. */
const receiverName = (node: unknown): string | null => {
  if (!isAstNode(node)) {
    return null;
  }
  if (node.type === "Identifier" && typeof node["name"] === "string") {
    return node["name"];
  }
  if (node.type === "TSNonNullExpression" || node.type === "ChainExpression") {
    return receiverName(node["expression"]);
  }
  if (node.type !== "MemberExpression" || node["computed"] === true) {
    return null;
  }
  const property = node["property"];
  return isAstNode(property) &&
    property.type === "Identifier" &&
    typeof property["name"] === "string"
    ? property["name"]
    : null;
};

/**
 * The name a comparison reads a value through: `cell.vMerge` -> `vMerge`.
 *
 * `throughProperty` is false for a bare identifier, which is what
 * {@link AMBIGUOUS_BARE_FIELDS} turns on.
 */
const comparedFieldName = (node: unknown, throughProperty = false): string | null => {
  if (!isAstNode(node)) {
    return null;
  }
  if (node.type === "Identifier" && typeof node["name"] === "string") {
    return throughProperty || !AMBIGUOUS_BARE_FIELDS.has(node["name"]) ? node["name"] : null;
  }
  if (node.type === "TSNonNullExpression" || node.type === "ChainExpression") {
    return comparedFieldName(node["expression"], throughProperty);
  }
  if (node.type !== "MemberExpression") {
    return null;
  }
  const receiver = receiverName(node["object"]);
  if (receiver !== null && NON_DOMAIN_RECEIVERS.has(receiver)) {
    return null;
  }
  if (node["computed"] === true) {
    // `attrs["vMerge"]` still names the field.
    const property = node["property"];
    return isAstNode(property) &&
      property.type === "Literal" &&
      typeof property["value"] === "string"
      ? property["value"]
      : null;
  }
  const property = node["property"];
  return isAstNode(property) &&
    property.type === "Identifier" &&
    typeof property["name"] === "string"
    ? property["name"]
    : null;
};

type LiteralOperand = { text: string; value: number | null };

/** The literal a comparison tests against, including a negated number. */
const literalOperand = (node: unknown): LiteralOperand | null => {
  if (!isAstNode(node)) {
    return null;
  }
  if (node.type === "UnaryExpression" && node["operator"] === "-") {
    const inner = literalOperand(node["argument"]);
    return inner === null || inner.value === null
      ? null
      : { text: `-${inner.text}`, value: -inner.value };
  }
  if (node.type !== "Literal") {
    return null;
  }
  const value = node["value"];
  if (typeof value === "string") {
    return { text: value, value: null };
  }
  if (typeof value === "number") {
    return { text: String(value), value };
  }
  return null;
};

/**
 * A relational guard against a numeric sentinel is the same decision spelled as
 * a bound: `outlineLevel <= 8`, `ilvl > 8` and `outlineLevel < 9` all ask
 * whether the value is the `9` sentinel.
 */
const guardsNumericSentinel = (rule: FieldRule, literal: LiteralOperand): boolean =>
  literal.value !== null &&
  [...rule.numbers].some(
    (sentinel) => sentinel === literal.value || sentinel - 1 === literal.value,
  );

const checkComparison = (node: AstNode, context: ReservedCompareContext): void => {
  const operator = node["operator"];
  if (typeof operator !== "string") {
    return;
  }
  const isEquality = EQUALITY_OPERATORS.has(operator);
  if (!isEquality && !RELATIONAL_OPERATORS.has(operator)) {
    return;
  }
  for (const [fieldSide, literalSide] of [
    [node["left"], node["right"]],
    [node["right"], node["left"]],
  ]) {
    const field = comparedFieldName(fieldSide);
    const literal = literalOperand(literalSide);
    if (field === null || literal === null) {
      continue;
    }
    const rule = FIELD_RULES.get(field);
    if (rule === undefined) {
      continue;
    }
    const matches = isEquality
      ? rule.literals.has(literal.text)
      : guardsNumericSentinel(rule, literal);
    if (!matches) {
      continue;
    }
    context.report({
      node,
      messageId: "bareReservedCompare",
      data: {
        field,
        literal: literal.text,
        reader: [...rule.readers].sort().join(", "),
      },
    });
    return;
  }
};

export default {
  meta: { name: "folio-reserved-values" },
  rules: {
    "no-bare-reserved-compare": {
      meta: {
        type: "problem",
        messages: {
          bareReservedCompare:
            "`{{field}}` carries an OOXML reserved value and `{{literal}}` is it. Read it through " +
            "{{reader}} instead: a second copy of this test is how the sentinel gets handled in " +
            "one place and missed in the next. The decision is recorded in " +
            "`specifications/reserved-values`.",
        },
      },
      create(context: ReservedCompareContext) {
        return {
          Program: (node: unknown) => {
            if (!isAstNode(node)) {
              return;
            }
            if (isOwningModule(context.filename) || isRegistryModule(context.filename)) {
              return;
            }
            const visit = (child: AstNode): void => {
              if (child.type === "BinaryExpression") {
                checkComparison(child, context);
              }
              forEachChild(child, visit);
            };
            visit(node);
          },
        };
      },
    },
  },
};
