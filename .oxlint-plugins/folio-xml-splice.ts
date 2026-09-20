// Forbid hand-rolled XML splices in folio source: a part is rewritten through
// one owner, `spliceXml` in `packages/core/src/docx/selectiveXmlPatch.ts`.
//
// `xml.slice(0, start) + replacement + xml.slice(end)` cuts a region out of a
// serialized part and drops something else in. A DOCX part's markup is not
// free-form text: `w:commentRangeStart`/`w:commentRangeEnd`,
// `w:bookmarkStart`/`w:bookmarkEnd` and the tracked-move range markers come in
// halves that must both survive. A selective patch rewrites only the regions
// an edit touched and keeps the rest byte-for-byte, so a splice that removes
// the half the edit moved writes the other half alone: invalid OOXML, and a
// comment anchored to nothing. The owner applies the same offsets and refuses
// a result that would, so the refusal is a property of the operation rather
// than a check each call site remembers.
//
// The rule fires on a `+` concatenation whose slices of one string leave a gap:
// the gap is the region deleted. Slices that meet — `slice(0, at)` beside
// `slice(at)` — cover the string end to end and only insert, so they cannot
// lose a marker and are not the rule's business. Neither is a lone slice, nor a
// chain assembled from parts already extracted.
//
// Flagged:
//   result.slice(0, start) + newXml + result.slice(end)
//   `${xml.slice(0, from)}${element}${xml.slice(to)}`
//
// Not flagged:
//   xml.slice(0, at) + run + xml.slice(at)           (meets: inserts only)
//   xml.slice(0, rootClose) + definitions            (one slice: appends)
//   head + minted + middle + added + tail            (no slice in the chain)
//   first.slice(0, at) + second.slice(at)            (two strings, one read each)

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  options?: unknown[];
  report: (descriptor: { node: unknown; messageId: "handRolledXmlSplice" }) => void;
};

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

/**
 * A structural key for the small expression grammar offsets are written in, so
 * `slash + 1` in one slice can be recognized as the same offset as `slash + 1`
 * in the next. Null for anything this cannot compare.
 */
const offsetKey = (node: unknown): string | null => {
  if (!isAstNode(node)) {
    return null;
  }
  switch (node.type) {
    case "Identifier":
      return typeof node["name"] === "string" ? `id:${node["name"]}` : null;
    case "Literal":
      return `lit:${String(node["value"])}`;
    case "MemberExpression": {
      const object = offsetKey(node["object"]);
      const property = offsetKey(node["property"]);
      return object === null || property === null
        ? null
        : `${object}${node["computed"] === true ? `[${property}]` : `.${property}`}`;
    }
    case "UnaryExpression": {
      const argument = offsetKey(node["argument"]);
      return argument === null ? null : `${String(node["operator"])}(${argument})`;
    }
    case "BinaryExpression": {
      const left = offsetKey(node["left"]);
      const right = offsetKey(node["right"]);
      return left === null || right === null
        ? null
        : `(${left}${String(node["operator"])}${right})`;
    }
    default:
      return null;
  }
};

/** The string sliced and the region taken, for an `x.slice(...)` call. */
type SliceRead = { source: string; from: string | null; to: string | null };

/** `x.slice(from, to)` read off a call expression, else null. */
const sliceRead = (node: unknown): SliceRead | null => {
  if (!isAstNode(node) || node.type !== "CallExpression") {
    return null;
  }
  const callee = node["callee"];
  if (!isAstNode(callee) || callee.type !== "MemberExpression" || callee["computed"] === true) {
    return null;
  }
  const property = callee["property"];
  const object = callee["object"];
  if (!isAstNode(property) || property.type !== "Identifier" || property["name"] !== "slice") {
    return null;
  }
  if (!isAstNode(object) || object.type !== "Identifier" || typeof object["name"] !== "string") {
    return null;
  }
  const args = Array.isArray(node["arguments"]) ? node["arguments"] : [];
  return {
    source: object["name"],
    from: args.length === 0 ? "lit:0" : offsetKey(args[0]),
    // A missing second argument runs to the end of the string.
    to: args.length < 2 ? null : offsetKey(args[1]),
  };
};

/**
 * Whether the reads tile the string end to end: the first opens at 0, each
 * hands its end to the next, and the last runs to the end. Nothing between
 * them is dropped, so no range marker can be lost.
 */
const coversWholeString = (reads: readonly SliceRead[]): boolean => {
  if (reads.at(0)?.from !== "lit:0" || reads.at(-1)?.to !== null) {
    return false;
  }
  return reads.every((read, index) => {
    const next = reads[index + 1];
    return next === undefined || (read.to !== null && read.to === next.from);
  });
};

/** Every operand of a `+` chain or template literal, flattened left to right. */
const concatenationOperands = (node: AstNode): unknown[] => {
  if (node.type === "TemplateLiteral") {
    return Array.isArray(node["expressions"]) ? node["expressions"] : [];
  }
  const operands: unknown[] = [];
  const visit = (part: unknown): void => {
    if (isAstNode(part) && part.type === "BinaryExpression" && part["operator"] === "+") {
      visit(part["left"]);
      visit(part["right"]);
      return;
    }
    operands.push(part);
  };
  visit(node["left"]);
  visit(node["right"]);
  return operands;
};

/** Whether an enclosing `+` chain already covers this node, so it reports once. */
const isNestedConcatenation = (node: AstNode): boolean => {
  const parent = node["parent"];
  return isAstNode(parent) && parent.type === "BinaryExpression" && parent["operator"] === "+";
};

/** The name a function is declared or bound under, walking out from `node`. */
const enclosingFunctionNames = (node: AstNode): string[] => {
  const names: string[] = [];
  let ancestor = node["parent"];
  while (isAstNode(ancestor)) {
    if (ancestor.type === "FunctionDeclaration" || ancestor.type === "TSDeclareFunction") {
      const id = ancestor["id"];
      if (isAstNode(id) && typeof id["name"] === "string") {
        names.push(id["name"]);
      }
    }
    if (ancestor.type === "VariableDeclarator") {
      const id = ancestor["id"];
      if (isAstNode(id) && typeof id["name"] === "string") {
        names.push(id["name"]);
      }
    }
    if (
      (ancestor.type === "MethodDefinition" || ancestor.type === "PropertyDefinition") &&
      isAstNode(ancestor["key"]) &&
      typeof (ancestor["key"] as AstNode)["name"] === "string"
    ) {
      names.push((ancestor["key"] as AstNode)["name"] as string);
    }
    ancestor = ancestor["parent"];
  }
  return names;
};

const stringArrayOption = (options: Record<string, unknown>, key: string): string[] => {
  const value = options[key];
  return Array.isArray(value)
    ? value.filter((entry): entry is string => typeof entry === "string")
    : [];
};

export default {
  meta: { name: "folio-xml-splice" },
  rules: {
    "no-hand-rolled-splice": {
      meta: {
        type: "problem",
        messages: {
          handRolledXmlSplice:
            "Cutting a region out of a serialized part by slicing it twice can " +
            "drop one half of a comment range, a bookmark or a tracked move, " +
            "leaving the other half alone in otherwise byte-preserved markup. " +
            "Apply the replacements through `spliceXml` in " +
            "`packages/core/src/docx/selectiveXmlPatch.ts`, which refuses a " +
            "result that would, or add the function to `allowedFunctions` with " +
            "the reason it cannot.",
        },
        schema: [
          {
            type: "object",
            properties: {
              allowedFunctions: { type: "array", items: { type: "string" } },
            },
            additionalProperties: false,
          },
        ],
      },
      create(context: RuleContext) {
        const options = isRecord(context.options?.[0]) ? context.options[0] : {};
        const allowedFunctions = new Set(stringArrayOption(options, "allowedFunctions"));

        const check = (node: AstNode): void => {
          const reads = concatenationOperands(node)
            .map((operand) => sliceRead(operand))
            .filter((read): read is SliceRead => read !== null);
          // One read cannot cut a region out, and reads of different strings
          // are reads, not one string rewritten.
          if (reads.length < 2 || new Set(reads.map(({ source }) => source)).size !== 1) {
            return;
          }
          if (coversWholeString(reads)) {
            return;
          }
          if (enclosingFunctionNames(node).some((name) => allowedFunctions.has(name))) {
            return;
          }
          context.report({ node, messageId: "handRolledXmlSplice" });
        };

        return {
          BinaryExpression: (node: unknown) => {
            if (isAstNode(node) && node["operator"] === "+" && !isNestedConcatenation(node)) {
              check(node);
            }
          },
          TemplateLiteral: (node: unknown) => {
            if (isAstNode(node)) {
              check(node);
            }
          },
        };
      },
    },
  },
};
