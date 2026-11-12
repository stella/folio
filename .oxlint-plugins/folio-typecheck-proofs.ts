// A compile-time claim in a test file is checked by nothing. Every package's
// `typecheck` runs `bun check` over `tsconfig.build.json`, which excludes
// `**/*.test.ts`, so a `@ts-expect-error` there is read by no compiler: the
// claim it makes can stop holding and no check reports it, and the directive
// itself can go stale without the usual "unused '@ts-expect-error'" error.
// Proofs belong in a `*.typecheck.ts` file, which `typecheck` does compile.

type CommentNode = {
  range: [number, number];
  type: string;
  value: string;
};

type ProgramNode = {
  comments?: CommentNode[];
  type: string;
};

type RuleContext = {
  filename: string;
  options?: unknown[];
  report: (diagnostic: {
    node: { range: [number, number] };
    messageId: "typeSuppressionInTest";
    data: { directive: string };
  }) => void;
};

// Mirrors TypeScript's own directive scan: only a comment that opens with the
// directive suppresses anything, so prose naming one is left alone.
const TS_SUPPRESSION =
  /^[/*]*\s*@ts-(?<directive>expect-error|ignore|nocheck)\b/u;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const stringArrayOption = (
  options: Record<string, unknown>,
  key: string,
): string[] => {
  const value = options[key];
  return Array.isArray(value)
    ? value.filter((entry) => typeof entry === "string")
    : [];
};

const normalizedPath = (filename: string): string =>
  filename.replaceAll("\\", "/");

type AstNode = Record<string, unknown> & { type: string };

type GateContext = {
  filename: string;
  options?: unknown[];
  report: (diagnostic: {
    node: unknown;
    messageId: "uncheckedFieldGate";
  }) => void;
};

const isAstNode = (value: unknown): value is AstNode =>
  isRecord(value) && typeof value["type"] === "string";

const SKIP_KEYS = new Set(["parent", "loc", "range", "start", "end", "type"]);

const forEachChild = (node: AstNode, visit: (child: AstNode) => void): void => {
  for (const [key, value] of Object.entries(node)) {
    if (SKIP_KEYS.has(key)) {
      continue;
    }
    for (const item of Array.isArray(value) ? value : [value]) {
      if (isAstNode(item)) {
        visit(item);
      }
    }
  }
};

const typeArguments = (node: AstNode): unknown[] => {
  const instantiation = node["typeArguments"];
  const params = isAstNode(instantiation) ? instantiation["params"] : null;
  return Array.isArray(params) ? params : [];
};

const referenceName = (node: unknown): string | null => {
  if (!isAstNode(node) || node.type !== "TSTypeReference") {
    return null;
  }
  const name = node["typeName"];
  return isAstNode(name) && typeof name["name"] === "string"
    ? name["name"]
    : null;
};

/** `keyof Model` or `RecordField<Model>` over a named type, not a value's
 * `typeof`: the field list of a declared type. */
const isFieldListOfType = (node: unknown): boolean => {
  if (!isAstNode(node)) {
    return false;
  }
  if (node.type === "TSTypeOperator" && node["operator"] === "keyof") {
    const operand = node["typeAnnotation"];
    return isAstNode(operand) && operand.type === "TSTypeReference";
  }
  return referenceName(node) === "RecordField";
};

/** A type that holds a value total over a declared type's fields:
 * `Required<Model>`, `Record<keyof Model, …>`, `{ [Field in keyof Model]-?: … }`. */
const isFieldTotalityGate = (node: unknown): boolean => {
  if (!isAstNode(node)) {
    return false;
  }
  const name = referenceName(node);
  if (name === "Required") {
    return true;
  }
  if (name === "Record") {
    return isFieldListOfType(typeArguments(node)[0]);
  }
  if (node.type === "TSMappedType") {
    const constraint =
      node["constraint"] ??
      (isAstNode(node["typeParameter"])
        ? node["typeParameter"]["constraint"]
        : null);
    return isFieldListOfType(constraint);
  }
  return false;
};

export default {
  meta: { name: "folio-typecheck-proofs" },
  rules: {
    "no-type-suppression-in-test": {
      meta: {
        type: "problem",
        messages: {
          typeSuppressionInTest:
            "`@ts-{{directive}}` in a test file is checked by nothing: a package's `typecheck` " +
            "runs over `tsconfig.build.json`, which excludes `**/*.test.ts`. State the claim in a " +
            "`*.typecheck.ts` proof beside the module it is about, or give the value a type that " +
            "says what it really holds.",
        },
        schema: [
          {
            type: "object",
            properties: {
              // Paths (or path suffixes) exempt from the rule. Empty: a test
              // file that needs a suppression is a claim in the wrong file, so
              // an exception is a reviewed change to this list, not a local
              // disable comment.
              allowedFiles: {
                type: "array",
                items: { type: "string" },
              },
            },
            additionalProperties: false,
          },
        ],
      },
      create(context: RuleContext) {
        const options = isRecord(context.options?.[0])
          ? context.options[0]
          : {};
        const allowedFiles = stringArrayOption(options, "allowedFiles");
        const filename = normalizedPath(context.filename);
        if (
          allowedFiles.some((allowed) =>
            filename.endsWith(normalizedPath(allowed)),
          )
        ) {
          return {};
        }

        return {
          Program: (node: unknown) => {
            if (!isRecord(node)) {
              return;
            }
            const { comments } = node as ProgramNode;
            if (comments === undefined) {
              return;
            }
            for (const comment of comments) {
              const directive = TS_SUPPRESSION.exec(comment.value)?.groups?.[
                "directive"
              ];
              if (directive !== undefined) {
                context.report({
                  node: comment,
                  messageId: "typeSuppressionInTest",
                  data: { directive },
                });
              }
            }
          },
        };
      },
    },
    "no-unchecked-field-gate": {
      meta: {
        type: "problem",
        messages: {
          uncheckedFieldGate:
            "A field-totality gate (`satisfies Required<…>`, `Record<keyof …>`, " +
            "`Record<RecordField<…>, …>` or a mapped type over a type's fields) in a test file " +
            "is checked by nothing: a package's `typecheck` excludes `**/*.test.ts`, so the " +
            "gate can stop holding unnoticed. Move the table into a `__fixtures__` module " +
            "beside the test, which the package's `typecheck` compiles, and import it.",
        },
      },
      create(context: GateContext) {
        return {
          Program: (node: unknown) => {
            if (!isAstNode(node)) {
              return;
            }
            const visit = (child: AstNode): void => {
              if (
                child.type === "TSSatisfiesExpression" &&
                isFieldTotalityGate(child["typeAnnotation"])
              ) {
                context.report({
                  node: child,
                  messageId: "uncheckedFieldGate",
                });
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
