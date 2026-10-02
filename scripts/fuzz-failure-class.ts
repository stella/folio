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

export const failureClass = (test: string, assertion: string) => {
  const consumer = /^consumer flow (.+) \/ (.+)$/u.exec(test);
  const flow = consumer?.[1] ?? test;
  const mode = consumer?.[2] ?? "all";
  const semantic =
    consumer === null ? undefined : semanticClasses.find(({ pattern }) => pattern.test(assertion));
  const message = semantic?.message ?? normalizeFailureMessage(assertion);
  const manualAreas = ["Table input", "Tracked revisions", "Document operations"];
  let area = "Fuzz checks";
  if (consumer !== null) area = "Consumer flow";
  if (manualAreas.includes(test)) area = test;
  if (semantic !== undefined) area = semantic.area;
  const kind = consumer === null ? normalizeFailureMessage(test) : "consumer-flow";
  // Known cross-mode oracles name their own family; unknown failures keep the
  // fixture and mode so unrelated assertions are never collapsed accidentally.
  const key = JSON.stringify([kind, semantic?.shape ?? flow, semantic?.mode ?? mode, message]);
  return { key, area, message, flow, mode };
};

/** Legacy automated and human titles remain discoverable until they close. */
export const classOfTitle = (title: string) => {
  const automated = /^Fuzz failure \[[0-9a-f]{16}\]: (consumer flow .+? \/ [^:]+): (.*)$/u.exec(
    title,
  );
  if (automated !== null) return failureClass(automated[1] ?? "", automated[2] ?? "");
  const colon = title.indexOf(": ");
  if (colon < 0) return null;
  return failureClass(title.slice(0, colon), title.slice(colon + 2));
};
