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
  // docxOps +1.1 kB: table row and table removal keep the ends of ranges
  // reaching past them. +0.5 kB: rejecting changes at once carries a cut
  // through the pieces it empties. +2.2 kB: moving blocks, tables included,
  // is one intent. The docxOps ceilings then keep about 5 kB of room, so a
  // correctness fix in the ops can land without editing this line; they come
  // down again once the queue is quiet.
  // minified +1.3 kB, gzip +0.5 kB: a session opening a story aligns each
  // paragraph's text when it is first addressed. docxOps +3.8 kB, past the
  // room left: a pasted or dragged block container (content control, text
  // box) is placed whole, with its notes and comments. docxOps +6.5 kB: room
  // again for ops correctness fixes; the ops measure about 674 kB. docxOps
  // +3.5 kB: a suggested deletion keeps the comment ranges it crosses, a
  // suggested join or paste keeps the records at its seam, and rejecting a
  // link insertion that cut an inline control gives one result in every order.
  // Room for the tracked table records now landing in a series (rows, cells,
  // nested tables and their moves; the ops measure about 694 kB), so
  // each one needs no edit here: docxOps to 715 kB, minified +18.5 kB, gzip
  // +6 kB. All six come down to the measured size plus a small margin once
  // the series has landed. minified +5 kB, gzip +1 kB: a pasted note or
  // comment keeps the parts its kept markup names in its own part's
  // relationships. The series outgrew its room (the ops measure about 715 kB
  // with five of its changes still to land, each about 6 kB): docxOps to
  // 745 kB, minified +30 kB, gzip +10 kB.
  react: { minified: 3_805_000, gzip: 1_198_000, docxOps: 745_000 },
  vue: { minified: 3_500_000, gzip: 1_094_000, docxOps: 745_000 },
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
      // Only the clipboard paste sanitizer uses DOMPurify; opening a
      // document must not load it.
      if (source.includes("node_modules/dompurify/") && input.bytesInOutput > 0)
        throw new Error(`${packageName} loaded DOMPurify eagerly`);
      if (source.startsWith("packages/docx-core/src/ops/"))
        metrics.docxOps += input.bytesInOutput;
    }
  }
  // Hyphenation patterns are data files fetched when a document needs them;
  // no chunk, eager or lazy, may carry a pattern set.
  for (const output of Object.values(buildResult.metafile.outputs)) {
    for (const [source, input] of Object.entries(output.inputs)) {
      if (
        source.includes("node_modules/hyphen/patterns/") &&
        input.bytesInOutput > 0
      )
        throw new Error(
          `${packageName} bundled hyphenation patterns (${source})`,
        );
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
