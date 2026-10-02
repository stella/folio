/** Stable reporting classes; fingerprints still identify individual replay cases. */
export const normalizeFailureMessage = (message: string): string =>
  message
    .replace(/^step\s+-?\d+:\s*/u, "")
    .replace(/(?:"(?:\\.|[^"\\])*"|(?<!\w)'[^'\n]*'(?!\w)|`[^`]*`)/gu, "<text>")
    .replace(/\b[0-9a-f]{8}-[0-9a-f-]{27,}\b/giu, "<id>")
    .replace(/\b[0-9a-f]{8,}\b/giu, "<id>")
    .replace(/(?:[A-Za-z]:)?(?:\/?[\w.-]+\/)+[\w.-]+/gu, "<path>")
    .replace(/-?\d+(?:\.\d+)?/gu, "<n>")
    .replace(/\s+/gu, " ")
    .replace(/<!--|-->/gu, "")
    .replace(/[:\s]+$/u, "")
    .trim();

// `shape` and `mode` describe each family; only the message identifies it.

const semanticClasses = [
  {
    pattern: /listLevel is .*expected/u,
    area: "List numbering",
    message: "listLevel mismatch",
    shape: "paragraph-numbering",
    mode: "all",
  },
  {
    pattern: /\[directTracked\]/u,
    area: "Revision acceptance",
    message: "accepted tracked batches differ from direct batches",
    shape: "document",
    mode: "direct-tracked",
  },
  {
    pattern: /the reopened package shows something else than the reviewer/u,
    area: "Save and reopen",
    message: "reopened package differs from saved reviewer",
    shape: "document",
    mode: "all",
  },
  {
    pattern: /\[saveIdempotent\]/u,
    area: "Save idempotence",
    message: "saving reopened package changes XML",
    shape: "document",
    mode: "all",
  },
  {
    pattern: /^block texts differ/u,
    area: "Tracked text",
    message: "block texts differ",
    shape: "document",
    mode: "revisions",
  },
  {
    pattern: /\[rejectAll\]/u,
    area: "Revision rejection",
    message: "rejected tracked batches differ from fixture",
    shape: "document",
    mode: "rejection",
  },
  {
    pattern: /\[batchSequential\]/u,
    area: "Batch operations",
    message: "batch differs from sequential operations",
    shape: "document",
    mode: "all",
  },
  {
    pattern: /Cannot serialize changed (?:footnote|endnote|<note>|note) paragraphs/u,
    area: "Note serialization",
    message: "Cannot serialize changed note paragraphs",
    shape: "notes",
    mode: "revisions",
  },
  {
    pattern: /\[readerStability\]/u,
    area: "Save and reopen",
    message: "content reads differ after save",
    shape: "document",
    mode: "all",
  },
] as const;

const MANUAL_AREAS = new Set(["Table input", "Tracked revisions", "Document operations"]);
const TEST_NAME_LIMIT = 80;

/** The test's own name, without the file it lives in. */
const shortTestName = (test: string): string => {
  const name = (test.split("::").at(-1) ?? test).replace(/\s+/gu, " ").trim();
  return name.length > TEST_NAME_LIMIT ? `${name.slice(0, TEST_NAME_LIMIT - 1)}…` : name;
};

type Semantic = { area: string; message: string };
const semanticOf = (assertion: string): Semantic | undefined =>
  semanticClasses.find(({ pattern }) => pattern.test(assertion));

const consumerClass = (
  semantic: Semantic | undefined,
  assertion: string,
  fixture: string,
  mode: string,
) => {
  const message = semantic?.message ?? normalizeFailureMessage(assertion);
  const area = semantic?.area ?? "Consumer flow";
  // One failure is one class wherever it shows: the fixture and the mode are
  // rows of the report, not part of its identity.
  const key = JSON.stringify(["consumer-flow", message]);
  return { key, area, message, flow: "", fixture, mode, title: `${area}: ${message}` };
};

export type FailureClass = ReturnType<typeof consumerClass>;

export const failureClass = (test: string, assertion: string): FailureClass => {
  const consumer = /^consumer flow (.+) \/ (.+)$/u.exec(test);
  if (consumer !== null) {
    return consumerClass(semanticOf(assertion), assertion, consumer[1] ?? "", consumer[2] ?? "all");
  }
  const message = normalizeFailureMessage(assertion);
  const manual = MANUAL_AREAS.has(test);
  const area = manual ? test : "Fuzz checks";
  // Other checks are told apart by the test that failed (file and name), and
  // the title names the test so two tests never share one.
  const key = JSON.stringify(["fuzz-check", test, message]);
  const title = manual ? `${test}: ${message}` : `${area}: ${shortTestName(test)}: ${message}`;
  return { key, area, message, flow: test, fixture: "", mode: "all", title };
};

/**
 * The current key of a class marker. Markers written while fixtures and modes
 * were part of the key carried four parts: kind, fixture or test, mode, message.
 */
export const upgradeClassKey = (parts: readonly string[]): string => {
  if (parts.length !== 4) return JSON.stringify(parts);
  return parts[0] === "consumer-flow"
    ? JSON.stringify(["consumer-flow", parts[3]])
    : JSON.stringify(["fuzz-check", parts[1], parts[3]]);
};

/** Legacy automated and human titles remain discoverable until they close. */
export const classOfTitle = (title: string): FailureClass | null => {
  const automated = /^Fuzz failure \[[0-9a-f]{16}\]: (consumer flow .+? \/ [^:]+): (.*)$/u.exec(
    title,
  );
  if (automated !== null) return failureClass(automated[1] ?? "", automated[2] ?? "");
  const colon = title.indexOf(": ");
  if (colon < 0) return null;
  const area = title.slice(0, colon);
  const rest = title.slice(colon + 2);
  if (area === "Consumer flow") return consumerClass(semanticOf(rest), rest, "", "all");
  const named = semanticClasses.find(
    (semantic) => semantic.area === area && semantic.message === rest,
  );
  if (named !== undefined) return consumerClass(named, rest, "", "all");
  return failureClass(area, rest);
};
