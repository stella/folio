// Forbid the empty string as a relationship reference in folio source.
//
// A relationship reference (`r:embed`, `r:id`, `r:link`) is a key into the
// owning part's relationships, and absence has one spelling, `undefined`. `""`
// is a second: it is not a key any relationship answers to, so it reads as
// present everywhere that tests for a value and as absent only where someone
// remembered to test for both. Three producers wrote it, and the drawing that
// carried it was saved as `<a:blip r:embed=""/>`.
//
// `Image.rId` is a `RelationshipId`, so the compiler answers for that field.
// The other references (`Hyperlink.rId`, `Image.hlinkRId`, `ImageDocPrLink`)
// are still plain strings, and a `""` in a ProseMirror attr or an untyped
// literal reaches them without the type system having an opinion. This rule is
// what holds the whole class while those fields are strings.
//
// Flagged:
//   { rId: "" }
//   { hlinkRId: attrs.hlinkRId || "" }
//   image.rId = "";
//
// Not flagged:
//   { rId: undefined }              (the one spelling of absence)
//   { rId: resolved ?? fallbackId } (a reference either way)

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  report: (descriptor: { node: unknown; messageId: "emptyRelationshipId" }) => void;
};

/** Every model and attr field whose value is a relationship reference. */
const RELATIONSHIP_ID_FIELDS = new Set(["rId", "hlinkRId"]);

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

const isEmptyStringLiteral = (value: unknown): boolean =>
  isAstNode(value) && value.type === "Literal" && value["value"] === "";

/**
 * Whether the expression can evaluate to the empty string as written: the
 * literal itself, or a fallback chain whose last resort is one. `a || ""` and
 * `a ?? ""` are how the empty spelling reached the model from an attr.
 */
const yieldsEmptyString = (value: unknown): boolean => {
  if (isEmptyStringLiteral(value)) {
    return true;
  }
  if (!isAstNode(value) || value.type !== "LogicalExpression") {
    return false;
  }
  const operator = value["operator"];
  if (operator !== "||" && operator !== "??") {
    return false;
  }
  return yieldsEmptyString(value["right"]);
};

/** The name a property or member access is written with, ignoring computed keys. */
const staticKeyName = (node: AstNode, keyField: "key" | "property"): string | undefined => {
  if (node["computed"] === true) {
    return undefined;
  }
  const key = node[keyField];
  if (!isAstNode(key)) {
    return undefined;
  }
  if (key.type === "Identifier" && typeof key["name"] === "string") {
    return key["name"];
  }
  return key.type === "Literal" && typeof key["value"] === "string" ? key["value"] : undefined;
};

export default {
  meta: { name: "folio-relationship-ids" },
  rules: {
    "no-empty-relationship-id": {
      meta: {
        type: "problem",
        messages: {
          emptyRelationshipId:
            "The empty string is not a relationship reference: it is a key no " +
            "relationship answers to, and a drawing that carried it was saved " +
            'as `r:embed=""`. A reference the source did not write is ' +
            "`undefined`, spelled by omitting the field or by narrowing the " +
            "value with `relationshipIdOf` from `@stll/docx-core/model`.",
        },
      },
      create(context: RuleContext) {
        return {
          Property: (node: unknown) => {
            if (!isAstNode(node)) {
              return;
            }
            const name = staticKeyName(node, "key");
            if (
              name !== undefined &&
              RELATIONSHIP_ID_FIELDS.has(name) &&
              yieldsEmptyString(node["value"])
            ) {
              context.report({ node, messageId: "emptyRelationshipId" });
            }
          },
          AssignmentExpression: (node: unknown) => {
            if (!isAstNode(node) || !yieldsEmptyString(node["right"])) {
              return;
            }
            const target = node["left"];
            if (!isAstNode(target) || target.type !== "MemberExpression") {
              return;
            }
            const name = staticKeyName(target, "property");
            if (name !== undefined && RELATIONSHIP_ID_FIELDS.has(name)) {
              context.report({ node, messageId: "emptyRelationshipId" });
            }
          },
        };
      },
    },
  },
};
