#!/usr/bin/env bun
/**
 * Report the dated policy exceptions that expire soon.
 *
 * Two checks fail on a date rather than on a change: the shared toolchain
 * policy (`stll-toolchain.json` entries with an `expires` date, read by
 * `bun run check:toolchain`) and the temporary quarantine excludes
 * (`# quarantine-expires:` in each bunfig.toml, read by
 * `bun run check:quarantine-excludes`). Once an entry expires, every run on
 * every branch fails until it is renewed or removed. This report runs on a
 * schedule and names each entry expiring within the window (or already
 * expired), as an annotation and in the job summary, so it is handled first.
 *
 * The toolchain file belongs to a shared policy, so this reads every object in
 * it that carries an `expires` date instead of a fixed list of its sections.
 *
 *   bun scripts/expiring-exceptions.ts [--days 7]
 *
 * Exits 1 when any entry is inside the window, so the scheduled run is red
 * days before the checks themselves would be.
 */

import { appendFileSync, existsSync, readFileSync } from "node:fs";
import path from "node:path";

import {
  quarantineInstallRoots,
  readTemporaryExcludes,
} from "./check-quarantine-excludes";

const ROOT = path.resolve(import.meta.dir, "..");
const DAY_MS = 24 * 60 * 60 * 1000;
const TOOLCHAIN_FILE = "stll-toolchain.json";
/** Fields that name a toolchain entry, in the order they are tried. */
const NAME_FIELDS = ["pin", "path", "dependency", "name", "package", "rule"];

export type DatedException = {
  /** The file that declares it, repo-relative. */
  file: string;
  /** What it exempts, as the file names it. */
  name: string;
  /** When the checks start failing on it, as written in the file. */
  expires: string;
  /** That instant in milliseconds. A bare date expires at its start (UTC). */
  expiresMs: number;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** Every object in a toolchain policy document with an `expires` date. */
export const toolchainExceptions = (
  document: unknown,
  file = TOOLCHAIN_FILE,
): DatedException[] => {
  const found: DatedException[] = [];
  const visit = (value: unknown, where: string): void => {
    if (Array.isArray(value)) {
      for (const [index, item] of value.entries())
        visit(item, `${where}[${String(index)}]`);
      return;
    }
    if (!isRecord(value)) return;
    const expires = value["expires"];
    if (typeof expires === "string") {
      const expiresMs = Date.parse(expires);
      if (Number.isNaN(expiresMs))
        throw new Error(`${file} ${where}: unreadable expiry ${expires}`);
      const label = NAME_FIELDS.map((field) => value[field]).find(
        (candidate): candidate is string => typeof candidate === "string",
      );
      found.push({
        file,
        name: label === undefined ? where : `${where} ${label}`,
        expires,
        expiresMs,
      });
    }
    for (const [key, child] of Object.entries(value))
      if (key !== "expires")
        visit(child, where === "" ? key : `${where}.${key}`);
  };
  visit(document, "");
  return found;
};

/** The temporary quarantine excludes of one bunfig.toml. */
export const quarantineExceptions = (
  bunfig: string,
  file: string,
): DatedException[] => {
  const { entries, errors } = readTemporaryExcludes(bunfig);
  if (errors.length > 0) throw new Error(`${file}: ${errors.join("; ")}`);
  return entries.map(({ name, expiresAt }) => ({
    file,
    name,
    expires: expiresAt,
    expiresMs: Date.parse(expiresAt),
  }));
};

/** The exceptions that expire before `now` plus `days`, soonest first. */
export const expiringWithin = (
  exceptions: readonly DatedException[],
  now: Date,
  days: number,
): DatedException[] =>
  exceptions
    .filter(({ expiresMs }) => expiresMs < now.getTime() + days * DAY_MS)
    .toSorted((a, b) => a.expiresMs - b.expiresMs);

/** One line per exception: the annotation and the summary row. */
export const describeException = (
  exception: DatedException,
  now: Date,
): { annotation: string; row: string } => {
  const days = Math.ceil((exception.expiresMs - now.getTime()) / DAY_MS);
  const when =
    days <= 0
      ? `expired ${exception.expires}`
      : `expires ${exception.expires} (in ${String(days)} d)`;
  return {
    annotation: `::warning file=${exception.file},title=Dated exception ${days <= 0 ? "expired" : "expires soon"}::${exception.name} ${when}; renew or remove it before the checks fail every run.`,
    row: `| ${exception.file} | ${exception.name} | ${when} |`,
  };
};

const readAll = (): DatedException[] => {
  const exceptions: DatedException[] = [];
  const toolchainPath = path.join(ROOT, TOOLCHAIN_FILE);
  if (existsSync(toolchainPath))
    exceptions.push(
      ...toolchainExceptions(JSON.parse(readFileSync(toolchainPath, "utf8"))),
    );
  for (const directory of quarantineInstallRoots) {
    const file = path.posix.join(directory, "bunfig.toml");
    const bunfigPath = path.join(ROOT, file);
    if (existsSync(bunfigPath))
      exceptions.push(
        ...quarantineExceptions(readFileSync(bunfigPath, "utf8"), file),
      );
  }
  return exceptions;
};

const main = (): void => {
  const index = process.argv.indexOf("--days");
  const days = index === -1 ? 7 : Number(process.argv[index + 1]);
  if (!Number.isInteger(days) || days < 1)
    throw new Error("--days must be a whole number of days");
  const now = new Date();
  const all = readAll();
  const soon = expiringWithin(all, now, days);
  const summary = process.env["GITHUB_STEP_SUMMARY"];
  const lines = [
    `## Dated exceptions expiring within ${String(days)} days`,
    "",
  ];
  if (soon.length === 0) {
    lines.push(
      `None of the ${String(all.length)} dated exceptions expires within ${String(days)} days.`,
    );
  } else {
    lines.push("| File | Exception | Expiry |", "| --- | --- | --- |");
    for (const exception of soon) {
      const { annotation, row } = describeException(exception, now);
      console.log(annotation);
      lines.push(row);
    }
  }
  const text = `${lines.join("\n")}\n`;
  if (summary !== undefined && summary !== "") appendFileSync(summary, text);
  console.log(text);
  if (soon.length > 0) process.exitCode = 1;
};

if (import.meta.main) main();
