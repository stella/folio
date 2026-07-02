// Forbid `new URL("<relative>.ts", import.meta.url)` (and the `.tsx` / `.cts` /
// `.mts` / `.jsx` variants) in folio source.
//
// A worker or asset referenced via `new URL("<spec>", import.meta.url)` — most
// often wrapped in `new Worker(new URL(...))` — is resolved LITERALLY by the
// consuming bundler against the emitted module. The package build (tsdown /
// rolldown) renames every `*.ts` source to `*.js` in dist, so a specifier that
// still points at a source extension resolves to a file that never ships. A
// downstream Vite/rolldown build then treats the worker as a build entry and
// aborts with UNRESOLVED_ENTRY (the same class of failure font-metrics.worker.ts
// caused). Reference the emitted name (`.js`) instead: in-repo source builds
// still resolve it via TypeScript extension resolution (`.js` -> `.ts`), and the
// published dist ships the real `.js`.
//
// The matching architecture test
// (`packages/core/src/__tests__/asset-url-extensions.test.ts`) walks the same
// call sites with `Bun.Glob` and asserts the same rule, so a loosened lint
// config cannot silently re-introduce a dangling source-extension URL.
//
// Flagged:
//   new Worker(new URL("font-metrics.worker.ts", import.meta.url))
//                       ^^^ .ts is not emitted into dist
//   new URL("./thing.tsx", import.meta.url)
//
// Safe:
//   new Worker(new URL("font-metrics.worker.js", import.meta.url))
//   new URL("https://example.com/x.ts")            (absolute URL, not a build asset)
//   new URL(someVariable, import.meta.url)          (dynamic; out of scope)

type AstNode = Record<string, unknown> & { type: string };

type RuleContext = {
  report: (descriptor: { node: unknown; messageId: "sourceExtensionUrl" }) => void;
};

const isAstNode = (value: unknown): value is AstNode =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  typeof (value as { type: unknown }).type === "string";

// Source extensions the package build renames to `.js` in dist. A URL asset
// pointing at any of these resolves to a file that never ships.
const SOURCE_EXTENSIONS = [".ts", ".tsx", ".cts", ".mts", ".jsx"];

// A string that names a build asset with a source extension. Strips a trailing
// `?query` / `#hash` (e.g. `./worker.ts?worker`) before matching, and skips
// absolute URLs (`https:`, `data:`, `file:`) which are not build assets.
const hasSourceExtension = (specifier: string): boolean => {
  if (/^[a-z][a-z0-9+.-]*:/iu.test(specifier)) {
    return false;
  }
  const base = specifier.split(/[?#]/u)[0] ?? specifier;
  return SOURCE_EXTENSIONS.some((ext) => base.endsWith(ext));
};

// True when a node is `import.meta.url`. Scoping the rule to import.meta-relative
// URLs keeps it to build-asset references and off unrelated `new URL(str, base)`
// calls.
const isImportMetaUrl = (node: unknown): boolean => {
  if (!isAstNode(node) || node.type !== "MemberExpression") {
    return false;
  }
  const property = node.property;
  const object = node.object;
  const propIsUrl =
    isAstNode(property) && property.type === "Identifier" && property["name"] === "url";
  const objIsImportMeta = isAstNode(object) && object.type === "MetaProperty";
  return propIsUrl && objIsImportMeta;
};

const literalStringValue = (node: unknown): string | null => {
  if (!isAstNode(node) || node.type !== "Literal") {
    return null;
  }
  return typeof node.value === "string" ? node.value : null;
};

const checkUrlConstruction = (context: RuleContext, node: AstNode): void => {
  const callee = node.callee;
  if (!isAstNode(callee) || callee.type !== "Identifier" || callee["name"] !== "URL") {
    return;
  }
  const args = Array.isArray(node.arguments) ? node.arguments : [];
  const [specifierArg, baseArg] = args;
  if (!isImportMetaUrl(baseArg)) {
    return;
  }
  const specifier = literalStringValue(specifierArg);
  if (specifier === null || !hasSourceExtension(specifier)) {
    return;
  }
  context.report({ node, messageId: "sourceExtensionUrl" });
};

export default {
  meta: { name: "folio-asset-urls" },
  rules: {
    "no-source-extension-url": {
      meta: {
        type: "problem",
        messages: {
          sourceExtensionUrl:
            "A `new URL(..., import.meta.url)` build asset (e.g. a worker entry) " +
            "must reference the name the package build emits into dist, not a " +
            "source extension. The build renames `.ts`/`.tsx` to `.js`, so a " +
            "source-extension specifier resolves to a file that never ships and " +
            "aborts a downstream bundler with UNRESOLVED_ENTRY. Point it at the " +
            "emitted `.js` (in-repo source builds still resolve `.js` -> `.ts` " +
            "via TypeScript extension resolution).",
        },
      },
      create(context: RuleContext) {
        return {
          NewExpression: (node: unknown) => {
            if (isAstNode(node)) {
              checkUrlConstruction(context, node);
            }
          },
        };
      },
    },
  },
};
