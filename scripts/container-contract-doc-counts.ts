/**
 * The counts in `docs/container-contract.md`, derived from the contract.
 *
 * The prose used to state how many pairs carry a mechanism, and the census
 * moved those numbers every time a parser did. A hand-updated number beside a
 * generated file is a mirror, and a mirror drifts: the doc claimed 118, 128 and
 * 108 pairs against a contract that held 289.
 *
 * So the numbers are derived. Each one sits between a marker naming the query
 * that produces it:
 *
 * ```md
 * <!--count:reason=editorProjection-->245<!--/count--> pairs carry this mechanism
 * ```
 *
 * `write` recomputes every one from `specifications/container-contract/contract.json`.
 * `check` fails when a number has drifted, when a filter names a key or a value
 * the contract does not define, and when a query selects nothing at all — a
 * query that matches no pair is a stale key rather than a true zero, and
 * letting it read as zero is how the drift started.
 *
 * Usage:
 *   bun scripts/container-contract-doc-counts.ts check
 *   bun scripts/container-contract-doc-counts.ts write
 */

import path from "node:path";

import { TaggedError } from "better-result";

import type {
  ContainerContract,
  ContractEntry,
} from "../specifications/container-contract/dispositions";
import { DISPOSITIONS, DROP_REASONS } from "../specifications/container-contract/dispositions";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const CONTRACT_PATH = path.join(REPOSITORY_ROOT, "specifications/container-contract/contract.json");
const DOC_PATH = path.join(REPOSITORY_ROOT, "docs/container-contract.md");

class DocCountError extends TaggedError("DocCountError")<{ message: string }> {}

/** The pair key with its namespaces removed: `bdo|CT_BdoContentRun/ins`. */
const QUALIFIER = /\{[^}]*\}/gu;

type Pair = {
  /** `element|Type`, namespaces removed. */
  container: string;
  kind: "child" | "attribute";
  /** The child element's or the attribute's local name. */
  subject: string;
  entry: ContractEntry;
};

const pairOf = (key: string, entry: ContractEntry): Pair => {
  const local = key.replaceAll(QUALIFIER, "");
  const attribute = local.lastIndexOf("@");
  if (attribute !== -1) {
    return {
      container: local.slice(0, attribute),
      kind: "attribute",
      subject: local.slice(attribute + 1),
      entry,
    };
  }
  const child = local.lastIndexOf("/");
  if (child === -1) {
    throw new DocCountError({
      message: `the contract key names neither a child nor an attribute: ${key}`,
    });
  }
  return {
    container: local.slice(0, child),
    kind: "child",
    subject: local.slice(child + 1),
    entry,
  };
};

/**
 * What a filter may ask about a pair.
 *
 * Total by construction: a key absent here is refused rather than ignored, so
 * a marker cannot quietly select every pair because its filter was misspelled.
 */
const FIELDS = {
  disposition: (pair: Pair) => [pair.entry.disposition],
  reason: (pair: Pair) =>
    pair.entry.disposition === DISPOSITIONS.dropped ? [pair.entry.reason] : [],
  kind: (pair: Pair) => [pair.kind],
  container: (pair: Pair) => [pair.container],
  subject: (pair: Pair) => [pair.subject],
} as const satisfies Record<string, (pair: Pair) => readonly string[]>;

type Field = keyof typeof FIELDS;

/** The values each field may be compared against, so a typo fails rather than matching nothing. */
const vocabularyOf = (pairs: readonly Pair[], field: Field): ReadonlySet<string> => {
  const values = new Set<string>();
  for (const pair of pairs) {
    for (const value of FIELDS[field](pair)) {
      values.add(value);
    }
  }
  if (field === "reason") {
    for (const reason of Object.keys(DROP_REASONS)) {
      values.add(reason);
    }
  }
  return values;
};

type Filter = { field: Field; negated: boolean; values: readonly string[] };

const isField = (name: string): name is Field => Object.hasOwn(FIELDS, name);

export const parseQuery = (query: string): Filter[] =>
  query
    .split("&")
    .map((term) => term.trim())
    .filter((term) => term !== "")
    .map((term) => {
      const negated = term.includes("!=");
      const comparison = negated ? "!=" : "=";
      const parts = term.split(comparison).map((part) => part.trim());
      if (parts.length !== 2 || parts.some((part) => part === "")) {
        throw new DocCountError({ message: `a count filter is not \`field=value\`: ${term}` });
      }
      const [fieldName, values] = parts;
      if (fieldName === undefined || values === undefined) {
        throw new DocCountError({ message: `a count filter is not \`field=value\`: ${term}` });
      }
      const field = fieldName.trim();
      if (!isField(field)) {
        throw new DocCountError({
          message: `a count filter names no contract field: ${field} (have ${Object.keys(FIELDS).join(", ")})`,
        });
      }
      return { field, negated, values: values.split(",").map((value) => value.trim()) };
    });

const matches = (pair: Pair, filters: readonly Filter[]): boolean =>
  filters.every(({ field, negated, values }) => {
    const held = FIELDS[field](pair);
    const any = values.some((value) => held.includes(value));
    return negated ? !any : any;
  });

const countFor = (pairs: readonly Pair[], query: string): number => {
  const filters = parseQuery(query);
  if (filters.length === 0) {
    throw new DocCountError({ message: "a count marker states no filter" });
  }
  for (const { field, values } of filters) {
    const vocabulary = vocabularyOf(pairs, field);
    for (const value of values) {
      if (!vocabulary.has(value)) {
        throw new DocCountError({
          message: `no pair has ${field} ${value}; the contract knows ${[...vocabulary].toSorted().join(", ")}`,
        });
      }
    }
  }
  const count = pairs.filter((pair) => matches(pair, filters)).length;
  if (count === 0) {
    throw new DocCountError({
      message: `count:${query} selects no pair; a marker that matches nothing is a stale query, not a zero`,
    });
  }
  return count;
};

const MARKER = /<!--count:(?<query>[^>]*?)-->(?<text>.*?)<!--\/count-->/gsu;

type Rewrite = { markdown: string; problems: string[] };

const rewrite = (markdown: string, pairs: readonly Pair[]): Rewrite => {
  const problems: string[] = [];
  const next = markdown.replaceAll(MARKER, (_match, query: string, text: string) => {
    const count = String(countFor(pairs, query.trim()));
    if (text !== count) {
      problems.push(`count:${query.trim()} reads ${text}, the contract says ${count}`);
    }
    return `<!--count:${query.trim()}-->${count}<!--/count-->`;
  });
  return { markdown: next, problems };
};

const main = async (): Promise<number> => {
  const mode = Bun.argv.at(2);
  if (mode !== "check" && mode !== "write") {
    throw new DocCountError({ message: "usage: container-contract-doc-counts.ts check|write" });
  }
  const contract = (await Bun.file(CONTRACT_PATH).json()) as ContainerContract;
  const pairs = Object.entries(contract.entries).map(([key, entry]) => pairOf(key, entry));
  const markdown = await Bun.file(DOC_PATH).text();
  const { markdown: next, problems } = rewrite(markdown, pairs);

  if (mode === "write") {
    if (next !== markdown) {
      await Bun.write(DOC_PATH, next);
    }
    console.log(
      problems.length === 0
        ? "docs/container-contract.md: every derived count already agreed"
        : `docs/container-contract.md: rewrote ${problems.length} count(s)`,
    );
    return 0;
  }
  for (const problem of problems) {
    console.error(`  ${problem}`);
  }
  if (problems.length > 0) {
    console.error("  re-run with `bun run container-contract:doc-counts`");
    return 1;
  }
  console.log("docs/container-contract.md: every derived count agrees with the contract");
  return 0;
};

if (import.meta.main) {
  process.exitCode = await main();
}
