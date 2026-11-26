// Measure the static JavaScript closure of the public mount API for each adapter.
// The adapters use their built entry points. Workspace dependencies resolve to
// source so esbuild can attribute docx-core operation bytes to their modules.

import { build, type Metafile } from "esbuild";
import path from "node:path";
import { gzipSync } from "node:zlib";

import { FOLIO_LOCALES } from "@stll/folio-core/i18n/messages/locales";

const catalogDir = "packages/core/src/i18n/messages";

const root = path.resolve(import.meta.dir, "../../..");
const ADAPTERS = ["react", "vue"] as const;
type Adapter = (typeof ADAPTERS)[number];
type Metrics = { minified: number; gzip: number; docxOps: number };

const BUDGETS = {
  react: { minified: 3_980_000, gzip: 1_257_000, docxOps: 663_000 },
  vue: { minified: 3_641_000, gzip: 1_130_000, docxOps: 664_000 },
} as const satisfies Record<Adapter, Metrics>;

const HOST_PEERS = [
  "prosemirror-*",
  "react",
  "react-dom",
  "vue",
  "yjs",
  "y-prosemirror",
  "y-protocols",
];

const eagerOutputs = ({ outputs }: Metafile): Set<string> => {
  const entry = Object.keys(outputs).find(
    (name) => outputs[name]?.entryPoint === "<stdin>",
  );
  if (entry === undefined) throw new Error("Missing consumer entry");
  const eager = new Set<string>();
  const pending = [entry];
  for (const name of pending) {
    if (eager.has(name)) continue;
    const output = outputs[name];
    if (output === undefined) throw new Error(`Missing eager output ${name}`);
    eager.add(name);
    for (const edge of output.imports) {
      if (edge.kind !== "import-statement" || edge.external) continue;
      pending.push(
        Object.hasOwn(outputs, edge.path)
          ? edge.path
          : path.posix.join(path.posix.dirname(name), edge.path),
      );
    }
  }
  return eager;
};

const checkCatalogs = (metafile: Metafile, eager: Set<string>) => {
  const eagerInputs = new Set<string>();
  const lazyInputs = new Set<string>();
  for (const [name, output] of Object.entries(metafile.outputs)) {
    const inputs = eager.has(name) ? eagerInputs : lazyInputs;
    for (const [source, input] of Object.entries(output.inputs)) {
      if (input.bytesInOutput > 0 && source.startsWith(`${catalogDir}/`))
        inputs.add(source);
    }
  }
  if (eagerInputs.has(`${catalogDir}/catalogs.gen.ts`))
    throw new Error("The aggregate message catalog is eager");
  for (const locale of FOLIO_LOCALES) {
    const input = `${catalogDir}/${locale}.gen.ts`;
    const isEnglish = locale === "en";
    if (
      isEnglish
        ? !eagerInputs.has(input)
        : eagerInputs.has(input) || !lazyInputs.has(input)
    )
      throw new Error(`Unexpected bundle placement for the ${locale} catalog`);
  }
};

const measure = async (adapter: Adapter): Promise<Metrics> => {
  const packageName = `@stll/folio-${adapter}`;
  const buildResult = await build({
    absWorkingDir: root,
    stdin: {
      contents: `
        import { renderAsync } from "${packageName}";
        const input = await fetch("/document.docx").then((response) => response.arrayBuffer());
        await renderAsync(input, document.body, { locale: "en" });
      `,
      resolveDir: root,
    },
    alias: {
      [packageName]: path.join(root, "packages", adapter, "dist/index.js"),
    },
    bundle: true,
    format: "esm",
    platform: "browser",
    target: "es2022",
    minify: true,
    splitting: true,
    metafile: true,
    write: false,
    outdir: path.join(root, ".cache/editor-eager-budget", adapter),
    external: HOST_PEERS,
    loader: { ".css": "empty" },
  });
  const eager = eagerOutputs(buildResult.metafile);
  checkCatalogs(buildResult.metafile, eager);
  const bytes = new Map(
    buildResult.outputFiles.map((file) => [
      path.relative(root, file.path).split(path.sep).join("/"),
      file.contents,
    ]),
  );
  const metrics = { minified: 0, gzip: 0, docxOps: 0 };
  for (const name of eager) {
    const contents = bytes.get(name);
    const output = buildResult.metafile.outputs[name];
    if (contents === undefined || output === undefined)
      throw new Error(`Missing contents for eager output ${name}`);
    metrics.minified += contents.byteLength;
    metrics.gzip += gzipSync(contents, { level: 9 }).byteLength;
    for (const [source, input] of Object.entries(output.inputs)) {
      if (
        source === "packages/docx-core/src/validate/schemaAttributes.gen.ts" &&
        input.bytesInOutput > 0
      ) {
        throw new Error(
          `${packageName} loaded the package attribute schema eagerly`,
        );
      }
      if (source.startsWith("packages/docx-core/src/ops/"))
        metrics.docxOps += input.bytesInOutput;
    }
  }
  if (metrics.docxOps === 0)
    throw new Error(`${packageName} did not include docx-core ops`);
  return metrics;
};

const measureOnly = process.argv.includes("--measure");
for (const adapter of ADAPTERS) {
  const metrics = await measure(adapter);
  const budget = BUDGETS[adapter];
  console.log(`@stll/folio-${adapter}`);
  console.table({
    "Eager JavaScript, minified": {
      bytes: metrics.minified,
      budget: budget.minified,
    },
    "Eager JavaScript, gzip": { bytes: metrics.gzip, budget: budget.gzip },
    "Eager docx-core ops, minified": {
      bytes: metrics.docxOps,
      budget: budget.docxOps,
    },
  });
  if (
    !measureOnly &&
    (metrics.minified > budget.minified ||
      metrics.gzip > budget.gzip ||
      metrics.docxOps > budget.docxOps)
  )
    process.exitCode = 1;
}
