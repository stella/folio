/**
 * Turns the engine measurements (one JSON file per host, document and engine,
 * written by tests/visual/engine-measurements.spec.ts) into one JSON summary
 * and markdown tables: per-key typing latency, open time and save time, the
 * default engine against the canonical session. Report-only: no number fails.
 *
 * Usage: bun scripts/engine-measurements-report.ts <measurements dir>
 * Writes <dir>/report.json and <dir>/report.md.
 */
import { readdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import process from "node:process";

type SaveSample = { ms: number; path: string; bytes: number | null };
type Measurement = {
  host: string;
  engine: "legacy" | "canonical";
  document: { id: string; label: string; bytes: number };
  mode: string;
  runner: { cpu: string; cores: number };
  open: {
    firstPageMs: number | null;
    laidOutMs: number | null;
    pages: number;
  }[];
  typing: {
    keys: number;
    frameMs: number[];
    layoutsPerKey: number;
    measuredBlocksPerKey: number;
  };
  save: {
    text: SaveSample[];
    structural: SaveSample[];
    textAfterStructural: SaveSample[];
    full: SaveSample[];
  };
  errors: string[];
};

const ENGINES = ["legacy", "canonical"] as const;

const percentile = (values: readonly number[], rank: number): number | null => {
  const sorted = values.toSorted((left, right) => left - right);
  if (sorted.length === 0) return null;
  const index = Math.min(
    sorted.length - 1,
    Math.max(0, Math.ceil(rank * sorted.length) - 1),
  );
  return sorted[index] ?? null;
};
const median = (values: readonly number[]) => percentile(values, 0.5);
const numbers = (values: readonly (number | null)[]): number[] =>
  values.filter((value): value is number => value !== null);

const summarize = (measurement: Measurement) => {
  const paths = (samples: readonly SaveSample[]) => [
    ...new Set(samples.map((sample) => sample.path)),
  ];
  return {
    typing: {
      keys: measurement.typing.keys,
      medianMs: median(measurement.typing.frameMs),
      p95Ms: percentile(measurement.typing.frameMs, 0.95),
      layoutsPerKey: measurement.typing.layoutsPerKey,
      measuredBlocksPerKey: measurement.typing.measuredBlocksPerKey,
    },
    open: {
      firstPageMs: median(
        numbers(measurement.open.map((sample) => sample.firstPageMs)),
      ),
      laidOutMs: median(
        numbers(measurement.open.map((sample) => sample.laidOutMs)),
      ),
      pages: measurement.open.at(-1)?.pages ?? null,
    },
    save: {
      textMs: median(measurement.save.text.map((sample) => sample.ms)),
      textPath: paths(measurement.save.text),
      structuralMs: median(
        measurement.save.structural.map((sample) => sample.ms),
      ),
      structuralPath: paths(measurement.save.structural),
      textAfterStructuralMs: median(
        measurement.save.textAfterStructural.map((sample) => sample.ms),
      ),
      textAfterStructuralPath: paths(measurement.save.textAfterStructural),
      fullMs: median(measurement.save.full.map((sample) => sample.ms)),
    },
    errors: measurement.errors,
  };
};
type Summary = ReturnType<typeof summarize>;

const directory = process.argv[2];
if (directory === undefined)
  throw new TypeError("Usage: engine-measurements-report.ts <dir>");

const measurements: Measurement[] = readdirSync(directory)
  .filter((file) => file.endsWith(".json") && file !== "report.json")
  .map((file) => JSON.parse(readFileSync(join(directory, file), "utf8")));

const hosts = [...new Set(measurements.map((entry) => entry.host))].toSorted();
const documents = [
  ...new Map(
    measurements.map((entry) => [entry.document.id, entry.document]),
  ).values(),
];

const find = (host: string, document: string, engine: string) =>
  measurements.find(
    (entry) =>
      entry.host === host &&
      entry.document.id === document &&
      entry.engine === engine,
  );

const rows = hosts.flatMap((host) =>
  documents.map((document) => {
    const engines = Object.fromEntries(
      ENGINES.map((engine) => {
        const measurement = find(host, document.id, engine);
        return [engine, measurement ? summarize(measurement) : null];
      }),
    ) as Record<(typeof ENGINES)[number], Summary | null>;
    return { host, document, engines };
  }),
);

const ms = (value: number | null | undefined) =>
  value === null || value === undefined ? "n/a" : `${Math.round(value)} ms`;
const ratio = (
  canonical: number | null | undefined,
  legacy: number | null | undefined,
) =>
  canonical === null ||
  canonical === undefined ||
  legacy === null ||
  legacy === undefined ||
  legacy === 0
    ? "n/a"
    : `${(canonical / legacy).toFixed(2)}x`;

const table = (header: readonly string[], body: readonly string[][]) =>
  [
    `| ${header.join(" | ")} |`,
    `|${header.map((_, index) => (index === 0 ? "---" : "---:")).join("|")}|`,
    ...body.map((row) => `| ${row.join(" | ")} |`),
  ].join("\n");

const first = measurements[0];
const lines = [
  "# Engine measurements",
  "",
  `Runner: ${first?.runner.cpu ?? "unknown"}, ${String(first?.runner.cores ?? 0)} cores. Mode: ${first?.mode ?? "unknown"}. Ratio = canonical / legacy (above 1.00x: canonical is slower).`,
  "",
];
for (const host of hosts) {
  const hostRows = rows.filter((row) => row.host === host);
  lines.push(
    `## ${host} host`,
    "",
    "### Typing: per-key latency (keydown to the next frame)",
    "",
    table(
      [
        "Document",
        "Legacy median",
        "Legacy p95",
        "Canonical median",
        "Canonical p95",
        "Ratio median",
        "Ratio p95",
      ],
      hostRows.map(({ document, engines: { legacy, canonical } }) => [
        document.label,
        ms(legacy?.typing.medianMs),
        ms(legacy?.typing.p95Ms),
        ms(canonical?.typing.medianMs),
        ms(canonical?.typing.p95Ms),
        ratio(canonical?.typing.medianMs, legacy?.typing.medianMs),
        ratio(canonical?.typing.p95Ms, legacy?.typing.p95Ms),
      ]),
    ),
    "",
    "### Open: bytes in to first page painted and to fully laid out (median of 3 opens)",
    "",
    table(
      [
        "Document",
        "Pages",
        "Legacy first page",
        "Canonical first page",
        "Ratio",
        "Legacy laid out",
        "Canonical laid out",
        "Ratio",
      ],
      hostRows.map(({ document, engines: { legacy, canonical } }) => [
        document.label,
        `${String(legacy?.open.pages ?? "n/a")} / ${String(canonical?.open.pages ?? "n/a")}`,
        ms(legacy?.open.firstPageMs),
        ms(canonical?.open.firstPageMs),
        ratio(canonical?.open.firstPageMs, legacy?.open.firstPageMs),
        ms(legacy?.open.laidOutMs),
        ms(canonical?.open.laidOutMs),
        ratio(canonical?.open.laidOutMs, legacy?.open.laidOutMs),
      ]),
    ),
    "",
    "### Save in one session: rounds of a text edit, Enter, a text edit, each saved (median of 3 rounds; the paths the saves took, in round order, in brackets)",
    "",
    table(
      [
        "Document",
        "Engine",
        "Text edit",
        "Enter",
        "Text edit after Enter",
        "Forced full repack",
      ],
      hostRows.flatMap(({ document, engines }) =>
        ENGINES.map((engine) => {
          const save = engines[engine]?.save;
          const cell = (
            value: number | null | undefined,
            path: readonly string[] | undefined,
          ) => `${ms(value)} (${path?.join(", ") ?? "n/a"})`;
          return [
            document.label,
            engine,
            cell(save?.textMs, save?.textPath),
            cell(save?.structuralMs, save?.structuralPath),
            cell(save?.textAfterStructuralMs, save?.textAfterStructuralPath),
            ms(save?.fullMs),
          ];
        }),
      ),
    ),
    "",
    "### Save ratios (canonical / legacy, medians)",
    "",
    table(
      ["Document", "Text edit", "Enter", "Text edit after Enter"],
      hostRows.map(({ document, engines: { legacy, canonical } }) => [
        document.label,
        ratio(canonical?.save.textMs, legacy?.save.textMs),
        ratio(canonical?.save.structuralMs, legacy?.save.structuralMs),
        ratio(
          canonical?.save.textAfterStructuralMs,
          legacy?.save.textAfterStructuralMs,
        ),
      ]),
    ),
    "",
  );
}
const errors = rows.flatMap((row) =>
  ENGINES.flatMap((engine) =>
    (row.engines[engine]?.errors ?? []).map(
      (error) => `- ${row.host}, ${row.document.label}, ${engine}: ${error}`,
    ),
  ),
);
const absent = rows.flatMap((row) =>
  ENGINES.filter((engine) => row.engines[engine] === null).map(
    (engine) =>
      `- ${row.host}, ${row.document.label}, ${engine}: no measurement`,
  ),
);
if (errors.length > 0 || absent.length > 0)
  lines.push("## Steps that did not run", "", ...errors, ...absent, "");

writeFileSync(
  join(directory, "report.json"),
  `${JSON.stringify(rows, null, 2)}\n`,
);
writeFileSync(join(directory, "report.md"), `${lines.join("\n")}\n`);
