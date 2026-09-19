// A parser that walks a container's children with its own `switch` or
// `if`/`else if` chain over child local names decides two things and writes
// down one. Which children are modelled is in the cases; what happens to the
// rest is the `default`, and a `default` that does nothing is how folio loses
// markup silently — the container contract counts 347 `never-parsed` pairs
// from exactly this shape.
//
// `docx/containerChildren.ts` is the one way to do it: a handler map the
// compiler makes total over the children the schema declares, and an ordered
// verbatim sink for everything else. This rule bans a new hand-rolled chain in
// the DOCX parsers and holds the unmigrated ones at their current count, so the
// list can only shrink.
//
// The signal is a dispatch over a *child name*: a `switch` whose discriminant
// is `getLocalName(...)` (or a variable a `getLocalName(...)` call initialised)
// with three or more cases. Three is the floor because a two-case switch is a
// choice between two known things, not a walk over a content model.

import baseline from "./folio-container-children.baseline.json" with { type: "json" };

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  filename: string;
  report: (descriptor: { node: unknown; messageId: "handRolledChildSwitch" }) => void;
};

/** Files whose chains predate the dispatcher, by count. The list only shrinks. */
const BASELINE: Readonly<Record<string, number>> = baseline;

const OWNER = "docx/containerChildren.ts";

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

/** `getLocalName(x)`, the repository's one way to read a child's name. */
const isLocalNameCall = (node: unknown): boolean => {
  if (!isAstNode(node) || node.type !== "CallExpression") {
    return false;
  }
  const callee = node.callee;
  return isAstNode(callee) && callee.type === "Identifier" && callee.name === "getLocalName";
};

/**
 * Names a parser binds a child's local name to. Deliberately narrow: a bare
 * `name` is any string in the repository, and matching it turned a
 * property-name guard into a container walk.
 */
const CHILD_NAME_IDENTIFIERS = new Set(["localName", "childName", "contentName"]);

/**
 * Whether a `switch` discriminant is a child's local name: the call itself, or
 * an identifier the file conventionally binds it to.
 */
const isChildNameDiscriminant = (discriminant: unknown): boolean => {
  if (isLocalNameCall(discriminant)) {
    return true;
  }
  return (
    isAstNode(discriminant) &&
    discriminant.type === "Identifier" &&
    typeof discriminant.name === "string" &&
    CHILD_NAME_IDENTIFIERS.has(discriminant.name)
  );
};

const MINIMUM_CASES = 3;

const normalize = (filename: string): string => {
  const path = filename.replaceAll("\\", "/");
  const index = path.indexOf("packages/");
  return index === -1 ? path : path.slice(index);
};

export default {
  meta: { name: "folio-container-children" },
  rules: {
    "no-hand-rolled-child-dispatch": {
      meta: {
        type: "problem",
        messages: {
          handRolledChildSwitch:
            "Walk a container's children with `dispatchChildren` from `docx/containerChildren`. " +
            "A hand-rolled switch over child local names states which children are modelled and " +
            "leaves what happens to the rest in a `default`, which is how unmodelled markup is " +
            "dropped without a decision. If this container is not migrated yet, raise its count " +
            "in `.oxlint-plugins/folio-container-children.baseline.json` — the baseline only shrinks.",
        },
      },
      create(context: RuleContext) {
        const file = normalize(context.filename);
        if (file.endsWith(OWNER)) {
          return {};
        }
        const allowed = BASELINE[file] ?? 0;
        let seen = 0;

        return {
          SwitchStatement: (node: unknown) => {
            if (!isAstNode(node) || !isChildNameDiscriminant(node.discriminant)) {
              return;
            }
            // `default` does not count: it is the branch this rule is about,
            // not evidence of a walk. A switch with two real cases is a choice
            // between two known things.
            const cases = node.cases;
            if (!Array.isArray(cases)) {
              return;
            }
            const named = cases.filter((entry) => isAstNode(entry) && entry.test !== null).length;
            if (named < MINIMUM_CASES) {
              return;
            }
            seen += 1;
            if (seen > allowed) {
              context.report({ node, messageId: "handRolledChildSwitch" });
            }
          },
        };
      },
    },
  },
};
