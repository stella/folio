// A compile-time claim in a test file is checked by nothing. Every package's
// `typecheck` runs `tsc` over `tsconfig.build.json`, which excludes
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
const TS_SUPPRESSION = /^[/*]*\s*@ts-(?<directive>expect-error|ignore|nocheck)\b/u;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const stringArrayOption = (options: Record<string, unknown>, key: string): string[] => {
  const value = options[key];
  return Array.isArray(value) ? value.filter((entry) => typeof entry === "string") : [];
};

const normalizedPath = (filename: string): string => filename.replaceAll("\\", "/");

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
        const options = isRecord(context.options?.[0]) ? context.options[0] : {};
        const allowedFiles = stringArrayOption(options, "allowedFiles");
        const filename = normalizedPath(context.filename);
        if (allowedFiles.some((allowed) => filename.endsWith(normalizedPath(allowed)))) {
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
              const directive = TS_SUPPRESSION.exec(comment.value)?.groups?.["directive"];
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
  },
};
