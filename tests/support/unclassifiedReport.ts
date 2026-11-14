// Scratch instrumentation: print the unclassified-transaction refusals the
// run collected, counted by site, message, input and transaction shape.
import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

const DIR = path.resolve("test-results/unclassified");

export default function unclassifiedReport(): void {
  const rows: {
    file: string;
    test: string;
    record: {
      site: string;
      message: string;
      lastInput?: string;
      context?: unknown;
      stack?: string[];
    };
  }[] = [];
  if (existsSync(DIR))
    for (const name of readdirSync(DIR))
      if (name.endsWith(".jsonl"))
        for (const line of readFileSync(path.join(DIR, name), "utf8").split(
          "\n",
        ))
          if (line.trim() !== "") rows.push(JSON.parse(line));
  const groups = new Map<string, { count: number; tests: Set<string> }>();
  for (const { file, test, record } of rows) {
    const key = JSON.stringify({
      site: record.site.startsWith("story:") ? "story" : record.site,
      message: record.message,
      input: record.lastInput ?? null,
      context: record.context ?? null,
    });
    const group = groups.get(key) ?? { count: 0, tests: new Set<string>() };
    group.count += 1;
    group.tests.add(`${file} > ${test}`);
    groups.set(key, group);
  }
  const ordered = [...groups].toSorted(([, a], [, b]) => b.count - a.count);
  const lines = [
    `[folio-unclassified-report] ${rows.length} refusals in ${groups.size} shapes`,
    ...ordered.map(
      ([key, group]) =>
        `[folio-unclassified-report] ${group.count}x ${key} tests=${JSON.stringify([...group.tests].slice(0, 4))}`,
    ),
  ];
  // oxlint-disable-next-line no-console -- scratch instrumentation
  console.log(lines.join("\n"));
  writeFileSync(
    path.resolve("test-results/unclassified-report.txt"),
    lines.join("\n") + "\n",
  );
}
