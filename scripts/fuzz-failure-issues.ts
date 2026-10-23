#!/usr/bin/env bun

import { panic, TaggedError } from "better-result";
import { $ } from "bun";
import { createHash } from "node:crypto";
import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import type {
  FailureMarker,
  FailureRecord,
} from "../test/consumer-scenarios/support/failure-fingerprints";
import {
  CONSUMER_SCENARIO_REPLAY_INPUTS,
  consumerScenarioReplay,
} from "./consumer-scenario-args";
import { extractFailureMarkers } from "./failure-fingerprints";
import { normalizeFailureMessage } from "./fuzz-failure-class";
import {
  type ClosedIssueDisposition,
  fileClasses,
  type Issue,
  type IssueStore,
} from "./fuzz-issue-classes";

export const LABEL = {
  name: "fuzz-failure",
  color: "b60205",
  description: "A failure a fuzz run found, one issue per invariant class",
};

export const TITLE_PREFIX = "Fuzz failure";
export const TITLE_FORMAT = `${TITLE_PREFIX} [<signature>]: <test>: <failure>`;

export const AUTO_CLOSE_CLEAN_RUNS = 7;
export const MAX_LISTED_SEEDS = 100;

export const RUN_TRIGGER = {
  scheduled: "scheduled",
  manual: "manual",
  mergeGroup: "merge-group",
} as const;
export type RunTrigger = (typeof RUN_TRIGGER)[keyof typeof RUN_TRIGGER];

export const RUN_RESULT = { success: "success", failure: "failure" } as const;
type RunResult = (typeof RUN_RESULT)[keyof typeof RUN_RESULT];

const LANE = /^[a-z0-9][a-z0-9-]*$/u;
const SIGNATURE = /^[0-9a-f]{16}$/u;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** A failure record read from disk, or null when it is not one. */
export const parseFailureRecord = (value: unknown): FailureRecord | null => {
  if (!isRecord(value) || value["version"] !== 1 || !isRecord(value["marker"]))
    return null;
  let marker: FailureMarker | undefined;
  try {
    [marker] = extractFailureMarkers(
      `FOLIO_FAILURE ${JSON.stringify(value["marker"])}`,
    );
  } catch {
    return null;
  }
  const replays = value["replays"];
  if (
    marker === undefined ||
    !Array.isArray(replays) ||
    replays.length === 0 ||
    !replays.every(
      (replay) => typeof replay === "string" && !replay.includes("\n"),
    ) ||
    typeof value["error"] !== "string"
  ) {
    return null;
  }
  const shrink = value["shrink"];
  return {
    version: 1,
    marker,
    replays: replays as string[],
    error: value["error"],
    ...(value["flow"] === undefined ? {} : { flow: value["flow"] }),
    ...(isRecord(shrink) &&
    typeof shrink["steps"] === "number" &&
    typeof shrink["from"] === "number" &&
    typeof shrink["attempts"] === "number"
      ? {
          shrink: {
            steps: shrink["steps"],
            from: shrink["from"],
            attempts: shrink["attempts"],
          },
        }
      : {}),
  };
};

/** One fingerprint's failures from a run: the most useful record and every seed. */
export type Finding = {
  record: FailureRecord;
  seeds: number[];
  records: FailureRecord[];
};

const stepsOf = (record: FailureRecord): number =>
  record.shrink?.steps ?? Number.POSITIVE_INFINITY;

/**
 * Group records and bare markers by fingerprint. A record beats a bare
 * marker, and a shorter minimized flow beats a longer one.
 */
export const collectFindings = (
  records: readonly FailureRecord[],
  markers: readonly FailureMarker[],
): Finding[] => {
  const findings = new Map<string, Finding>();
  const add = (record: FailureRecord, bare: boolean): void => {
    const { fingerprint, seed } = record.marker;
    const found = findings.get(fingerprint);
    if (found === undefined) {
      findings.set(fingerprint, { record, seeds: [seed], records: [record] });
      return;
    }
    if (!found.seeds.includes(seed)) found.seeds.push(seed);
    const existingIndex = found.records.findIndex(
      ({ marker }) =>
        marker.seed === seed && marker.path === record.marker.path,
    );
    if (existingIndex < 0) found.records.push(record);
    else {
      const existing = found.records.at(existingIndex);
      if (
        existing !== undefined &&
        !bare &&
        stepsOf(record) < stepsOf(existing)
      ) {
        found.records[existingIndex] = record;
      }
    }
    if (!bare && stepsOf(record) < stepsOf(found.record)) found.record = record;
  };
  for (const record of records) add(record, false);
  for (const marker of markers) {
    add(
      { version: 1, marker, replays: [marker.repro], error: marker.assertion },
      true,
    );
  }
  return [...findings.values()];
};

export type Context = {
  runUrl: string | null;
  sha: string | null;
  source: string;
  date: string;
  lane: string;
  trigger: RunTrigger;
};

const runTriggerOf = (value: string): RunTrigger | undefined =>
  Object.values(RUN_TRIGGER).find((trigger) => trigger === value);
const runResultOf = (value: string): RunResult | undefined =>
  Object.values(RUN_RESULT).find((result) => result === value);

class UsageError extends TaggedError("UsageError")<{ message: string }> {}

const parseArgs = (argv: readonly string[]) => {
  const options = {
    logs: [] as string[],
    records: [] as string[],
    runUrl: null as string | null,
    sha: null as string | null,
    source: "a fuzz run",
    dryRun: false,
    result: null as RunResult | null,
    lane: null as string | null,
    trigger: null as RunTrigger | null,
    artifactRoot: null as string | null,
    logPattern: "**/*.log",
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const value = (): string => argv[++index] ?? "";
    if (arg === "--log") options.logs.push(value());
    else if (arg === "--records") options.records.push(value());
    else if (arg === "--run-url") options.runUrl = value() || null;
    else if (arg === "--sha") options.sha = value() || null;
    else if (arg === "--source") options.source = value() || options.source;
    else if (arg === "--artifact-root") options.artifactRoot = value();
    else if (arg === "--log-pattern") options.logPattern = value();
    else if (arg === "--dry-run") options.dryRun = true;
    else if (arg === "--lane") {
      const lane = value();
      if (!LANE.test(lane))
        throw new UsageError({ message: `Invalid --lane ${lane}` });
      options.lane = lane;
    } else if (arg === "--result") {
      const result = value();
      options.result = runResultOf(result) ?? null;
      if (options.result === null)
        throw new UsageError({ message: `Invalid --result ${result}` });
    } else if (arg === "--trigger") {
      const trigger = value();
      options.trigger = runTriggerOf(trigger) ?? null;
      if (options.trigger === null) {
        throw new UsageError({ message: `Invalid --trigger ${trigger}` });
      }
    } else throw new UsageError({ message: `Unknown argument ${String(arg)}` });
  }
  const { lane, result, trigger } = options;
  if (lane === null || result === null || trigger === null) {
    throw new UsageError({
      message: "--lane, --result and --trigger are required",
    });
  }
  return { ...options, lane, result, trigger };
};

const readRecords = (dir: string): FailureRecord[] => {
  if (!existsSync(dir)) return [];
  return readdirSync(dir, { recursive: true })
    .map(String)
    .filter((file) => file.endsWith(".json"))
    .flatMap((file) => {
      try {
        const record = parseFailureRecord(
          JSON.parse(readFileSync(path.join(dir, file), "utf8")),
        );
        return record === null ? [] : [record];
      } catch {
        return [];
      }
    });
};

const writeBody = (body: string): string => {
  const file = path.join(tmpdir(), `fuzz-failure-${String(process.pid)}.md`);
  writeFileSync(file, body);
  return file;
};

class IssueResponseError extends TaggedError("IssueResponseError")<{
  message: string;
}> {}

/** GitHub's close reasons (REST `state_reason`, GraphQL `stateReason`) a closed issue can carry. */
const CLOSE_REASONS = {
  completed: { type: "fixed" },
  not_planned: { type: "superseded" },
  duplicate: { type: "duplicate" },
} as const satisfies Record<string, ClosedIssueDisposition>;
export type CloseReason = keyof typeof CLOSE_REASONS;

const isCloseReason = (reason: string): reason is CloseReason =>
  Object.hasOwn(CLOSE_REASONS, reason);

/** A closed issue's disposition; a close reason this reporter does not know fails. */
const closedIssueDisposition = (
  stateReason: unknown,
): ClosedIssueDisposition => {
  // Issues closed before GitHub recorded close reasons carry none: they were closed as done.
  if (stateReason === null || stateReason === undefined || stateReason === "") {
    return CLOSE_REASONS.completed;
  }
  const reason =
    typeof stateReason === "string" ? stateReason.toLowerCase() : null;
  if (reason === null || !isCloseReason(reason)) {
    throw new IssueResponseError({
      message: `Unknown close reason ${JSON.stringify(stateReason)} on a closed issue`,
    });
  }
  return CLOSE_REASONS[reason];
};

/** Reject malformed API responses before they can affect matching or writes. */
export const parseIssueResponse = (value: unknown): Issue => {
  if (!isRecord(value))
    throw new IssueResponseError({ message: "Invalid GitHub issue response" });
  const number = value["number"];
  const title = value["title"];
  const body = value["body"];
  const state =
    typeof value["state"] === "string" ? value["state"].toLowerCase() : null;
  const closedAt = value["closed_at"] ?? value["closedAt"] ?? null;
  const stateReason = value["state_reason"] ?? value["stateReason"] ?? null;
  if (
    typeof number !== "number" ||
    !Number.isSafeInteger(number) ||
    number <= 0 ||
    typeof title !== "string" ||
    (body !== null && typeof body !== "string") ||
    (state !== "open" && state !== "closed") ||
    (closedAt !== null && typeof closedAt !== "string")
  )
    throw new IssueResponseError({ message: "Invalid GitHub issue fields" });
  if (state === "open") return { number, title, body, state, closedAt: null };
  return {
    number,
    title,
    body,
    state,
    closedAt,
    disposition: closedIssueDisposition(stateReason),
  };
};

/** The bodies of an issue's comments, from paginated GitHub responses. */
export const parseCommentPages = (pages: unknown): string[] => {
  if (!Array.isArray(pages) || !pages.every(Array.isArray)) {
    throw new IssueResponseError({ message: "Invalid GitHub comment pages" });
  }
  return pages.flat().map((comment: unknown) => {
    if (!isRecord(comment) || typeof comment["body"] !== "string") {
      throw new IssueResponseError({
        message: "Invalid GitHub comment fields",
      });
    }
    return comment["body"];
  });
};

/** Validate paginated GitHub responses before selecting an issue to update. */
export const parseIssuePages = (pages: unknown): Issue[] => {
  if (!Array.isArray(pages) || !pages.every(Array.isArray)) {
    throw new IssueResponseError({ message: "Invalid GitHub issue pages" });
  }
  return pages.flat().map(parseIssueResponse);
};

/** Where a finding was filed: `#<number>`, or null when it was not. */
export type Filed = { fingerprint: string; issue: string | null };

export type IssueContent = { title: string; body: string };

export type SignatureIssueStore = Pick<IssueStore, "list" | "create"> & {
  edit: (number: number, content: IssueContent) => Promise<void>;
  reopen: (number: number) => Promise<void>;
  close: (number: number, reason: CloseReason) => Promise<void>;
  closingCommit?: (number: number) => Promise<string | null>;
};

const PAGE_SIZE = 100;

const ghPages = async (endpoint: string): Promise<unknown[]> => {
  const pages: unknown[] = [];
  for (let page = 1; ; page += 1) {
    const separator = endpoint.includes("?") ? "&" : "?";
    const response: unknown = JSON.parse(
      await $`gh api ${`${endpoint}${separator}per_page=${PAGE_SIZE}&page=${page}`}`.text(),
    );
    pages.push(response);
    if (!Array.isArray(response) || response.length < PAGE_SIZE) return pages;
  }
};

/** The commit attached to the latest close event, or null for a manual close. */
export const parseClosingCommitPages = (pages: unknown): string | null => {
  if (!Array.isArray(pages) || !pages.every(Array.isArray)) {
    throw new IssueResponseError({
      message: "Invalid GitHub issue event pages",
    });
  }
  const closing = pages
    .flat()
    .findLast(
      (event: unknown) => isRecord(event) && event["event"] === "closed",
    );
  if (!isRecord(closing)) return null;
  const commit = closing["commit_id"];
  if (commit === null || commit === undefined) return null;
  if (typeof commit !== "string" || !/^[0-9a-f]{40}$/iu.test(commit)) {
    throw new IssueResponseError({ message: "Invalid GitHub closing commit" });
  }
  return commit;
};

export const githubIssueStore = () =>
  ({
    list: async () =>
      parseIssuePages(
        await ghPages(
          `repos/{owner}/{repo}/issues?state=all&labels=${LABEL.name}`,
        ),
      ),
    create: async (title, body) => {
      await $`gh label create ${LABEL.name} --color ${LABEL.color} --description ${LABEL.description} --force`.quiet();
      const url = (
        await $`gh issue create --title ${title} --label ${LABEL.name} --body-file ${writeBody(body)}`.text()
      ).trim();
      return parseIssueResponse(
        JSON.parse(
          await $`gh issue view ${url} --json number,title,state,body,closedAt,stateReason`.text(),
        ),
      );
    },
    edit: async (number, { title, body }) => {
      await $`gh issue edit ${number} --title ${title} --body-file ${writeBody(body)}`.quiet();
    },
    reopen: async (number) => {
      await $`gh issue reopen ${number}`.quiet();
    },
    close: async (number, reason) => {
      await $`gh api -X PATCH ${`repos/{owner}/{repo}/issues/${number}`} -f state=closed -f ${`state_reason=${reason}`}`.quiet();
    },
    closingCommit: async (number) =>
      parseClosingCommitPages(
        await ghPages(`repos/{owner}/{repo}/issues/${number}/events`),
      ),
    comments: async (number: number) =>
      parseCommentPages(
        await ghPages(`repos/{owner}/{repo}/issues/${number}/comments`),
      ),
  }) satisfies SignatureIssueStore & {
    comments: (number: number) => Promise<string[]>;
  };

const githubClassIssueStore = (): IssueStore => {
  const issues = githubIssueStore();
  return {
    list: issues.list,
    create: issues.create,
    edit: async (number, body) => {
      await $`gh issue edit ${number} --body-file ${writeBody(body)}`.quiet();
    },
    reopen: async (number, comment) => {
      await issues.reopen(number);
      await $`gh issue comment ${number} --body ${comment}`.quiet();
    },
    close: (number) => issues.close(number, "completed"),
  };
};

class MissingFailuresError extends TaggedError("MissingFailuresError")<{
  message: string;
}> {}

export const requireFindings = (findings: readonly Finding[]): void => {
  if (findings.length === 0) {
    throw new MissingFailuresError({
      message:
        "Failed fuzz/flow run contained 0 FOLIO_FAILURE fingerprints; check the failure artifact and marker producer",
    });
  }
};

// Bound free text so it cannot displace the complete replay from GitHub's body limit.
const FIELD_LIMIT = 4_000;
const field = (value: string, limit = FIELD_LIMIT): string => {
  if (value.length <= limit) return value;
  let kept = Math.max(0, limit - 24);
  while (true) {
    const marker = `… (${value.length - kept} more characters)`;
    const next = Math.max(0, limit - marker.length);
    if (next === kept) return `${value.slice(0, kept)}${marker}`;
    kept = next;
  }
};
const line = (value: string, limit = FIELD_LIMIT): string =>
  field(value, limit).replaceAll(/\s+/gu, " ").trim();
const code = (value: string, limit = FIELD_LIMIT): string => {
  const bounded = field(value, limit);
  const fence = "`".repeat(
    Math.max(
      3,
      ...[...bounded.matchAll(/`+/gu)].map(([run]) => run.length + 1),
    ),
  );
  return `${fence}\n${bounded}\n${fence}`;
};

/** JSON strings keep label text on one line; escaped pipes keep table cells structural. */
const inlineText = (value: string, limit = FIELD_LIMIT): string =>
  JSON.stringify(field(value, limit)).replaceAll("|", "\\u007c");

const parseInlineText = (
  value: string,
  limit = FIELD_LIMIT,
): string | undefined => {
  let parsed: unknown;
  try {
    parsed = JSON.parse(value);
  } catch {
    return undefined;
  }
  return typeof parsed === "string" && inlineText(parsed, limit) === value
    ? parsed
    : undefined;
};

const parseLegacyField = (
  value: string,
  limit = FIELD_LIMIT,
): string | undefined => {
  // Remove this read-only path once no open fuzz-failure issue has a legacy body.
  return line(value, limit) === value ? value : undefined;
};

export type Signature = { id: string; test: string; failure: string };

const legacySignatureFailure = (assertion: string): string => {
  if (/^Missing numbering format for displayed label/u.test(assertion))
    return "Missing numbering format for displayed label";
  if (/listLevel is .*expected/u.test(assertion)) return "listLevel mismatch";
  if (/\[directTracked\]/u.test(assertion))
    return "accepted tracked batches differ from direct batches";
  if (/\[rejectAll\]/u.test(assertion))
    return "rejected tracked batches differ from fixture";
  if (/\[batchSequential\]/u.test(assertion))
    return "batch differs from sequential operations";
  if (/\[readerStability\]/u.test(assertion))
    return "content reads differ after save";
  if (/\[saveIdempotent\]/u.test(assertion))
    return "saving reopened package changes XML";
  if (/^block texts differ/u.test(assertion)) return "block texts differ";
  return normalizeFailureMessage(assertion);
};

export const signatureOf = ({
  test,
  assertion,
}: Pick<FailureMarker, "test" | "assertion">) => {
  const failure = legacySignatureFailure(assertion);
  const id = createHash("sha256")
    .update(`${test}\0${failure}`)
    .digest("hex")
    .slice(0, 16);
  return { id, test, failure } satisfies Signature;
};

type LaneRow = { lane: string; lastRun: string; cleanRuns: number };
type Repro = { seed: number; command: string; flowPath?: string };

export type SignatureState = {
  signature: Signature;
  occurrences: number;
  seeds: number[];
  lastSeen: { lane: string; run: string };
  lanes: LaneRow[];
  assertion: string;
  repro: Repro;
};

const ISSUE_BODY_LIMIT = 65_535;
const ISSUE_BODY_FORMAT_MARKER = "<!-- fuzz-issue-format: 2 -->";
const LANE_TABLE_HEADER = [
  "| Lane | Last run | Clean scheduled runs |",
  "| --- | --- | --- |",
];

type ShellWord = { raw: string; value: string };

/** Decode the shell syntax emitted by scenario replay commands. */
export const shellWords = (command: string): ShellWord[] | null => {
  const words: ShellWord[] = [];
  let index = 0;
  while (index < command.length) {
    while (command[index] === " " || command[index] === "\t") index += 1;
    if (index === command.length) break;
    const start = index;
    let value = "";
    while (
      index < command.length &&
      command[index] !== " " &&
      command[index] !== "\t"
    ) {
      const character = command[index] ?? "";
      if (character === "'") {
        const end = command.indexOf("'", index + 1);
        if (end === -1) return null;
        value += command.slice(index + 1, end);
        index = end + 1;
        continue;
      }
      if (character === '"') {
        index += 1;
        let closed = false;
        while (index < command.length) {
          const quoted = command[index] ?? "";
          if (quoted === '"') {
            index += 1;
            closed = true;
            break;
          }
          if (quoted === "`" || quoted === "$" || quoted === "\n") return null;
          if (quoted !== "\\") {
            value += quoted;
            index += 1;
            continue;
          }
          const escaped = command[index + 1];
          if (escaped === undefined) return null;
          if (
            escaped === "$" ||
            escaped === "`" ||
            escaped === '"' ||
            escaped === "\\"
          ) {
            value += escaped;
          } else if (escaped !== "\n") {
            value += `\\${escaped}`;
          }
          index += 2;
        }
        if (!closed) return null;
        continue;
      }
      if (character === "\\") {
        const escaped = command[index + 1];
        if (escaped === undefined) return null;
        if (escaped !== "\n") value += escaped;
        index += 2;
        continue;
      }
      if (character === "\n" || /[;|&<>()`$]/u.test(character)) return null;
      value += character;
      index += 1;
    }
    words.push({ raw: command.slice(start, index), value });
  }
  return words;
};

/** Convert old environment-prefixed scenario replays before they reach public issue text. */
export const cliRepro = (command: string): string => {
  const words = shellWords(command);
  if (words === null) return command;
  const inputs: Record<string, string> = {};
  let index = 0;
  while (index < words.length) {
    const word = words[index];
    const assignment = /^(FOLIO_SCENARIO_[A-Z_]+)=(.*)$/su.exec(
      word?.value ?? "",
    );
    if (assignment === null) break;
    const [, name = "", value = ""] = assignment;
    const flag = CONSUMER_SCENARIO_REPLAY_INPUTS.find(
      ({ env }) => env === name,
    )?.flag;
    if (flag === undefined) break;
    const option = flag
      .slice(2)
      .replaceAll(/-([a-z])/gu, (_, letter: string) => letter.toUpperCase());
    inputs[option] = value;
    index += 1;
  }
  if (Object.keys(inputs).length === 0) return command;
  if (
    words[index]?.value !== "bun" ||
    words[index + 1]?.value !== "scripts/consumer-scenarios.ts"
  ) {
    return command;
  }
  const suffix = words.slice(index + 2);
  const flags: string[] = [];
  for (let suffixIndex = 0; suffixIndex < suffix.length; suffixIndex += 1) {
    const word = suffix[suffixIndex];
    if (word?.value === "--") {
      flags.push(...suffix.slice(suffixIndex).map(({ raw }) => raw));
      break;
    }
    const replayInput = CONSUMER_SCENARIO_REPLAY_INPUTS.find(
      ({ flag }) => flag === word?.value,
    );
    if (replayInput === undefined) {
      flags.push(word?.raw ?? "");
      continue;
    }
    const value = suffix[suffixIndex + 1];
    if (value === undefined) return command;
    const option = replayInput.flag
      .slice(2)
      .replaceAll(/-([a-z])/gu, (_, letter: string) => letter.toUpperCase());
    inputs[option] = value.value;
    suffixIndex += 1;
  }
  return consumerScenarioReplay({ inputs, flags: [flags.join(" ")] });
};

const flowArtifactReplay = (
  command: string,
  seed: number,
  artifact: string,
): string => {
  const replay = cliRepro(command);
  const substituted = replay.replace(
    /(--flow\s+)(?:'(?:'\\''|[^'])*'|"(?:\\.|[^"])*"|\S+)/u,
    `$1'${artifact}'`,
  );
  return substituted === replay
    ? consumerScenarioReplay({ inputs: { seed: String(seed), flow: artifact } })
    : substituted;
};

/** A complete fenced repro; this boundary deliberately has no length parameter. */
const wholeRepro = ({ seed, command }: Repro): string =>
  `Latest repro (seed ${seed}):\n${code(cliRepro(command), Number.POSITIVE_INFINITY)}`;

export const renderSignatureIssue = (state: SignatureState): IssueContent => {
  const { signature } = state;
  const title = `${TITLE_PREFIX} [${signature.id}]: ${line(signature.test, 200)}: ${line(signature.failure, 200)}`;
  const prefix = [
    ISSUE_BODY_FORMAT_MARKER,
    `Signature: ${signature.id}`,
    `Test: ${inlineText(signature.test)}`,
    `Failure: ${inlineText(signature.failure)}`,
    `Occurrences: ${state.occurrences}`,
    `Seeds: ${state.seeds.join(", ")}`,
    `Last seen: ${state.lastSeen.lane} ${inlineText(state.lastSeen.run, 500)}`,
    "",
    ...LANE_TABLE_HEADER,
    ...state.lanes.map(
      ({ lane, lastRun, cleanRuns }) =>
        `| ${lane} | ${inlineText(lastRun, 500)} | ${cleanRuns} |`,
    ),
    "",
    `Closes after ${AUTO_CLOSE_CLEAN_RUNS} consecutive clean scheduled runs in every lane.`,
    "",
  ];
  let repro = wholeRepro(state.repro);
  const availableForAssertion =
    ISSUE_BODY_LIMIT - prefix.join("\n").length - repro.length - 24;
  let assertion = code(
    field(
      state.assertion,
      Math.max(0, Math.min(FIELD_LIMIT, availableForAssertion)),
    ),
  );
  let body = [...prefix, `Latest assertion:\n${assertion}`, "", repro].join(
    "\n",
  );
  if (body.length > ISSUE_BODY_LIMIT) {
    const artifact =
      state.repro.flowPath ??
      `findings/${signature.id}-${state.repro.seed}.flow.json`;
    repro = [
      `Flow artifact: ${artifact}`,
      wholeRepro({
        seed: state.repro.seed,
        command: flowArtifactReplay(
          state.repro.command,
          state.repro.seed,
          artifact,
        ),
      }),
    ].join("\n");
    const fallbackAvailable =
      ISSUE_BODY_LIMIT - prefix.join("\n").length - repro.length - 24;
    assertion = code(
      field(
        state.assertion,
        Math.max(0, Math.min(FIELD_LIMIT, fallbackAvailable)),
      ),
    );
    body = [...prefix, `Latest assertion:\n${assertion}`, "", repro].join("\n");
  }
  return {
    title: title.length > 240 ? `${title.slice(0, 239)}…` : title,
    body,
  };
};

class SignatureStateError extends TaggedError("SignatureStateError")<{
  message: string;
}> {}

const fieldOf = (body: string, name: string): string | undefined =>
  new RegExp(`^${name}: (.*)$`, "mu").exec(body)?.at(1);

const seedsOf = (text: string): number[] =>
  text === ""
    ? []
    : text
        .split(", ")
        .map((seed) => (/^-?\d+$/u.test(seed) ? Number(seed) : Number.NaN));

export const parseSignatureState = ({
  number,
  body,
}: Pick<Issue, "number" | "body">) => {
  const id = body === null ? undefined : fieldOf(body, "Signature");
  if (body === null || id === undefined) return null;
  const malformed = (what: string) =>
    new SignatureStateError({
      message: `Issue #${number} has a malformed ${what}`,
    });
  if (!SIGNATURE.test(id)) throw malformed("signature");
  // The writer emits the marker as the first line; a marker elsewhere is body text.
  const format =
    body.split("\n", 1)[0] === ISSUE_BODY_FORMAT_MARKER
      ? "canonical"
      : "legacy";
  const encodedTest = fieldOf(body, "Test") ?? "";
  const encodedFailure = fieldOf(body, "Failure") ?? "";
  const parseBodyField =
    format === "canonical" ? parseInlineText : parseLegacyField;
  const test = parseBodyField(encodedTest);
  const failure = parseBodyField(encodedFailure);
  const occurrences = Number(fieldOf(body, "Occurrences"));
  const seeds = seedsOf(fieldOf(body, "Seeds") ?? "");
  const lastSeen = /^Last seen: (\S+) (.*)$/mu.exec(body);
  const evidence = /^Latest assertion:\n(`{3,})\n([\s\S]*?)\n\1$/mu.exec(body);
  const assertion = evidence?.at(2);
  const repro =
    /^Latest repro \(seed (-?\d+)\):\n(`{3,})\n([\s\S]*?)\n\2$/mu.exec(
      body.slice((evidence?.index ?? 0) + (evidence?.[0].length ?? 0)),
    );
  if (test === undefined || failure === undefined) throw malformed("signature");
  if (!Number.isSafeInteger(occurrences) || occurrences < 0)
    throw malformed("occurrence count");
  if (!seeds.every(Number.isSafeInteger)) throw malformed("seed list");
  const [, seenLane, encodedSeenRun] = lastSeen ?? [];
  const seenRun =
    format === "canonical"
      ? parseInlineText(encodedSeenRun ?? "", 500)
      : parseLegacyField(encodedSeenRun ?? "", 500);
  if (seenLane === undefined || !LANE.test(seenLane) || seenRun === undefined) {
    throw malformed("last seen run");
  }
  const [, reproSeed, , command] = repro ?? [];
  if (
    assertion === undefined ||
    reproSeed === undefined ||
    command === undefined
  ) {
    throw malformed("evidence");
  }
  const tableStart = body.indexOf(LANE_TABLE_HEADER.join("\n"));
  if (tableStart < 0) throw malformed("lane table");
  const rows = body
    .slice(tableStart)
    .split("\n")
    .slice(LANE_TABLE_HEADER.length);
  const end = rows.findIndex((row) => !row.startsWith("|"));
  const lanes = rows.slice(0, end < 0 ? rows.length : end).map((row) => {
    const canonicalRow = /^\| (\S+) \| ((?:\\.|[^|])+) \| (\d+) \|$/u.exec(row);
    const legacyRow = /^\| (\S+) \| (.+) \| (\d+) \|$/u.exec(row);
    const [, lane = "", encodedLastRun = "", cleanRuns = ""] =
      (format === "canonical" ? canonicalRow : legacyRow) ?? [];
    const lastRun =
      format === "canonical"
        ? parseInlineText(encodedLastRun, 500)
        : parseLegacyField(encodedLastRun, 500);
    if (!LANE.test(lane) || lastRun === undefined)
      throw malformed("lane table");
    return { lane, lastRun, cleanRuns: Number(cleanRuns) };
  });
  return {
    signature: { id, test, failure },
    occurrences,
    seeds,
    lastSeen: { lane: seenLane, run: seenRun },
    lanes,
    assertion,
    repro: { seed: Number(reproSeed), command },
  } satisfies SignatureState;
};

export const mergeSeeds = (
  seeds: readonly number[],
  added: readonly number[],
): number[] => {
  const merged = [...seeds];
  for (const seed of added) if (!merged.includes(seed)) merged.push(seed);
  return merged.slice(-MAX_LISTED_SEEDS);
};

type SignatureFinding = {
  signature: Signature;
  seeds: number[];
  fingerprints: string[];
  record: FailureRecord;
};

export const groupBySignature = (
  findings: readonly Finding[],
): SignatureFinding[] => {
  const groups = new Map<string, SignatureFinding>();
  for (const { record, records } of findings) {
    const signature = signatureOf(record.marker);
    const group = groups.get(signature.id);
    const seeds = records.map(({ marker }) => marker.seed);
    if (group === undefined) {
      groups.set(signature.id, {
        signature,
        seeds: [...new Set(seeds)],
        fingerprints: [record.marker.fingerprint],
        record,
      });
      continue;
    }
    for (const seed of seeds)
      if (!group.seeds.includes(seed)) group.seeds.push(seed);
    if (!group.fingerprints.includes(record.marker.fingerprint)) {
      group.fingerprints.push(record.marker.fingerprint);
    }
    if (stepsOf(record) < stepsOf(group.record)) group.record = record;
  }
  return [...groups.values()];
};

const runKeyOf = (context: Context): string =>
  context.runUrl ?? `local:${context.date}`;

type RecordReproductionOptions = {
  state: SignatureState | null;
  finding: SignatureFinding;
  context: Context;
};

export const recordReproduction = ({
  state,
  finding,
  context,
}: RecordReproductionOptions): SignatureState => {
  const run = runKeyOf(context);
  if (state?.lastSeen.lane === context.lane && state.lastSeen.run === run)
    return state;
  const lanes = (state?.lanes ?? []).map(({ lane, lastRun }) => ({
    lane,
    lastRun,
    cleanRuns: 0,
  }));
  if (context.trigger === RUN_TRIGGER.scheduled) {
    const row = { lane: context.lane, lastRun: run, cleanRuns: 0 };
    const index = lanes.findIndex(({ lane }) => lane === context.lane);
    if (index < 0) lanes.push(row);
    else lanes[index] = row;
  }
  const { record } = finding;
  return {
    signature: state?.signature ?? finding.signature,
    occurrences: (state?.occurrences ?? 0) + finding.seeds.length,
    seeds: mergeSeeds(state?.seeds ?? [], finding.seeds),
    lastSeen: { lane: context.lane, run },
    lanes,
    assertion: record.marker.assertion,
    repro: {
      seed: record.marker.seed,
      command: cliRepro(record.replays.at(0) ?? record.marker.repro),
      ...(record.flow === undefined
        ? {}
        : {
            flowPath: `findings/${record.marker.fingerprint}-${record.marker.seed}.flow.json`,
          }),
    },
  };
};

export const recordCleanRun = (
  state: SignatureState,
  context: Context,
): SignatureState => {
  const run = runKeyOf(context);
  const row = state.lanes.find(({ lane }) => lane === context.lane);
  if (
    context.trigger !== RUN_TRIGGER.scheduled ||
    row === undefined ||
    row.lastRun === run
  ) {
    return state;
  }
  return {
    signature: state.signature,
    occurrences: state.occurrences,
    seeds: state.seeds,
    lastSeen: state.lastSeen,
    lanes: state.lanes.map((item) =>
      item.lane === context.lane
        ? { lane: item.lane, lastRun: run, cleanRuns: item.cleanRuns + 1 }
        : item,
    ),
    assertion: state.assertion,
    repro: state.repro,
  };
};

export const resolved = ({ lanes }: SignatureState): boolean =>
  lanes.length > 0 &&
  lanes.every(({ cleanRuns }) => cleanRuns >= AUTO_CLOSE_CLEAN_RUNS);

type Tracked = { issue: Issue; state: SignatureState };

const trackedIssue = (
  tracked: readonly Tracked[],
  id: string,
): Tracked | undefined => {
  const matching = tracked.filter(({ state }) => state.signature.id === id);
  const open = matching
    .filter(({ issue }) => issue.state === "open")
    .toSorted((a, b) => a.issue.number - b.issue.number)
    .at(0);
  if (open !== undefined) return open;
  return matching
    .toSorted(
      (a, b) =>
        (Date.parse(b.issue.closedAt ?? "") || 0) -
          (Date.parse(a.issue.closedAt ?? "") || 0) ||
        b.issue.number - a.issue.number,
    )
    .at(0);
};

type RecurOptions = {
  store: SignatureIssueStore;
  issue: Issue;
  state: SignatureState;
};
const recur = async ({ store, issue, state }: RecurOptions): Promise<void> => {
  if (issue.state === "closed") {
    const { disposition } = issue;
    switch (disposition.type) {
      case "fixed":
        await store.reopen(issue.number);
        break;
      case "superseded":
      case "duplicate":
        break;
      default: {
        const unhandled: never = disposition;
        panic(
          `Unhandled closed issue disposition ${JSON.stringify(unhandled)}`,
        );
      }
    }
  }
  await store.edit(issue.number, renderSignatureIssue(state));
};

type FileFindingsOptions = {
  findings: readonly Finding[];
  context: Context;
  store?: SignatureIssueStore;
  isAncestor?: (
    ancestor: string,
    descendant: string,
  ) => boolean | Promise<boolean>;
};
type FileFindingsResult = { filed: Filed[]; closed: number[] };

export const gitIsAncestor = (
  ancestor: string,
  descendant: string,
): boolean => {
  const result = Bun.spawnSync(
    ["git", "merge-base", "--is-ancestor", ancestor, descendant],
    {
      stdout: "ignore",
      stderr: "pipe",
    },
  );
  if (result.exitCode === 0) return true;
  if (result.exitCode === 1) return false;
  throw new Error(
    result.stderr.toString().trim() || "git merge-base --is-ancestor failed",
  );
};

export const fileFindings = async ({
  findings,
  context,
  store = githubIssueStore(),
  isAncestor = gitIsAncestor,
}: FileFindingsOptions): Promise<FileFindingsResult> => {
  const tracked = (await store.list()).flatMap((issue) => {
    const state = parseSignatureState(issue);
    return state === null ? [] : [{ issue, state }];
  });
  const filed: Filed[] = [];
  const seen = new Set<string>();
  // A caller may hand us repeated records; identity always comes from the fingerprint.
  const unique = collectFindings(
    findings.flatMap(({ records }) => records),
    findings.map(({ record }) => record.marker),
  );
  for (const finding of groupBySignature(unique)) {
    seen.add(finding.signature.id);
    const existing = trackedIssue(tracked, finding.signature.id);
    if (
      existing?.issue.state === "closed" &&
      existing.issue.disposition.type === "fixed"
    ) {
      const closingCommit = await store.closingCommit?.(existing.issue.number);
      // A manual close has no commit to compare, so preserve the established reopen behavior.
      if (
        closingCommit !== null &&
        closingCommit !== undefined &&
        context.sha !== null &&
        !(await isAncestor(closingCommit, context.sha))
      ) {
        for (const fingerprint of finding.fingerprints)
          filed.push({ fingerprint, issue: null });
        continue;
      }
    }
    const state = recordReproduction({
      state: existing?.state ?? null,
      finding,
      context,
    });
    let number: number;
    if (existing === undefined) {
      const { title, body } = renderSignatureIssue(state);
      const issue = await store.create(title, body);
      tracked.push({ issue, state });
      number = issue.number;
    } else {
      if (state !== existing.state)
        await recur({ store, issue: existing.issue, state });
      number = existing.issue.number;
    }
    for (const fingerprint of finding.fingerprints)
      filed.push({ fingerprint, issue: `#${number}` });
  }
  const closed: number[] = [];
  for (const { issue, state } of tracked) {
    if (issue.state !== "open" || seen.has(state.signature.id)) continue;
    const next = recordCleanRun(state, context);
    if (next === state) continue;
    await store.edit(issue.number, renderSignatureIssue(next));
    if (!resolved(next)) continue;
    await store.close(issue.number, "completed");
    closed.push(issue.number);
  }
  return { filed, closed };
};

/** Discover selected logs without merging matrix artifacts with identical filenames. */
export const artifactInputs = (root: string, pattern = "**/*.log") => {
  const logs = [
    ...new Bun.Glob(pattern).scanSync({ cwd: root, onlyFiles: true }),
  ].map((file) => path.join(root, file));
  return {
    logs,
    records: [...new Set(logs.map((file) => path.dirname(file)))],
  };
};

/** The findings in `logs` (their FOLIO_FAILURE lines) and the record directories. */
export const readFindings = (
  logs: readonly string[],
  records: readonly string[],
): Finding[] =>
  collectFindings(
    records.flatMap(readRecords),
    logs
      .filter((log) => existsSync(log))
      .flatMap((log) => extractFailureMarkers(readFileSync(log, "utf8"))),
  );

const main = async (): Promise<void> => {
  const options = parseArgs(process.argv.slice(2));
  const artifacts =
    options.artifactRoot === null
      ? { logs: [], records: [] }
      : artifactInputs(options.artifactRoot, options.logPattern);
  const findings = readFindings(
    [...options.logs, ...artifacts.logs],
    [...options.records, ...artifacts.records],
  );
  if (options.result === RUN_RESULT.failure) requireFindings(findings);
  const context: Context = {
    runUrl: options.runUrl,
    sha: options.sha,
    source: options.source,
    date: new Date().toISOString(),
    lane: options.lane,
    trigger: options.trigger,
  };
  const filed = await fileClasses({
    findings,
    context,
    store: options.dryRun
      ? {
          list: () => Promise.resolve([]),
          create: (title, body) => {
            console.log(`=== ${title}\n${body}\n`);
            return Promise.resolve({
              number: 1,
              title,
              body,
              state: "open",
              closedAt: null,
            });
          },
          edit: () => Promise.resolve(),
          reopen: () => Promise.resolve(),
          close: () => Promise.resolve(),
        }
      : githubClassIssueStore(),
    known: new Map(),
  });
  for (const { fingerprint, issue } of filed) {
    console.log(`${fingerprint}: ${issue ?? "not filed"}`);
  }
};

if (import.meta.main) {
  await main();
}
