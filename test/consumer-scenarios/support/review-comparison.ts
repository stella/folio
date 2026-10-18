import assert from "node:assert/strict";
import { canonicalFieldModel } from "./field-model.ts";
import { openReviewer } from "./documents.ts";
import { projectContentPair } from "./identity.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type ReviewSource = Pick<Reviewer, "getContent" | "getChanges" | "toResolvedDocument">;
type Change = ReturnType<Reviewer["getChanges"]>[number];
type Coverage = Pick<Change, "type" | "author" | "date" | "text" | "blockId">;

/** Adjacent fragments in each provenance stream keep lossless character order. */
export const revisionCoverage = (
  changes: readonly Change[],
  aliases: ReadonlyMap<string, string> = new Map(),
): Coverage[] => {
  const streams = new Map<string, Coverage[]>();
  for (const { type, author, date, text, blockId } of changes) {
    const block = blockId === null ? null : (aliases.get(blockId) ?? blockId);
    const key = JSON.stringify([type, author, date, block]);
    const stream = streams.get(key) ?? [];
    const previous = stream.at(-1);
    if (text.length > 0 && previous && previous.text.length > 0) {
      previous.text += text;
    } else {
      stream.push({ type, author, date, text, blockId: block });
    }
    streams.set(key, stream);
  }
  return [...streams.keys()].sort().flatMap((key) => streams.get(key) ?? []);
};

export const assertRevisionCoverage = (left: readonly Change[], right: readonly Change[]) =>
  assert.deepEqual(revisionCoverage(right), revisionCoverage(left));

// Source replay hints are not authored document properties. Every such field
// needs an explicit disposition; unknown capture fields or symbols fail closed.
const SOURCE_FIELDS = new Map(
  Object.entries({
    requiredFonts: { disposition: "exclude", reason: "Derived font-loading inventory" },
    templateVariables: {
      disposition: "exclude",
      reason: "Derived template-token inventory; authored text stays strict",
    },
    warnings: { disposition: "exclude", reason: "Source parse diagnostics" },
    parseWarnings: { disposition: "exclude", reason: "Source parse diagnostics" },
    modified: {
      disposition: "exclude",
      reason: "Package save bookkeeping; revision dates remain strict",
    },
    originalBuffer: { disposition: "exclude", reason: "Original ZIP replay source" },
    verbatimXml: { disposition: "exclude", reason: "Replay bytes for a modeled carrier" },
    verbatimFingerprint: { disposition: "exclude", reason: "Replay cache validity" },
    namespaceDeclarations: { disposition: "exclude", reason: "Lexical namespace spelling" },
    namespaceBindings: { disposition: "exclude", reason: "Lexical namespace spelling" },
    sourceXml: {
      disposition: "exclude",
      reason: "Source property XML replay; modeled properties stay strict",
    },
    gridSourceXml: {
      disposition: "exclude",
      reason: "Source grid XML replay; modeled grid columns stay strict",
    },
    rawWatermarkXml: { disposition: "keep", reason: "Watermark content not fully modeled" },
    rawImageXml: { disposition: "keep", reason: "Drawing content not fully modeled" },
    rawImageFingerprint: { disposition: "exclude", reason: "Drawing replay cache validity" },
  } as const),
);
const SOURCE_SYMBOLS = {
  paragraphPropertyCapture: "Source paragraph property replay handle",
  paragraphPropertySourceBinding: "Source package replay binding",
  paragraphPropertySourceBindingBrand: "Source binding type brand",
  paragraphPropertySourceEmissionFingerprint: "Source emission cache validity",
  paragraphPropertySourceIndentationBaseline: "Source indentation capture",
  paragraphIndentationProjection: "Source indentation projection cache",
  "package-source-baseline": "Source header/footer replay handle",
} as const;

const EMPTY_PART_RELATIONSHIPS = new Map([
  [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "Empty comment part has no modeled comments",
  ],
  [
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    "Empty comment thread part has no modeled comments",
  ],
]);
const hasModeledComments = (value: object): boolean => {
  if (!("document" in value) || typeof value.document !== "object" || value.document === null)
    return true;
  return (
    "comments" in value.document &&
    Array.isArray(value.document.comments) &&
    value.document.comments.length > 0
  );
};

/** One strict modeled-document projection, including nested part collections. */
export const canonicalReviewModel = (
  model: unknown,
  aliases: ReadonlyMap<string, string> = new Map(),
): unknown => {
  const project = (current: unknown, modelPath: string): unknown => {
    const value = current;
    if (Array.isArray(value))
      return value.map((entry, index) => project(entry, `${modelPath}/${index}`));
    if (value instanceof Map)
      return new Map(
        [...value].map(([key, entry]) => [key, project(entry, `${modelPath}/${String(key)}`)]),
      );
    if (value instanceof Set) return new Set([...value].map((entry) => project(entry, modelPath)));
    if (value === null || typeof value !== "object") return value;
    if (Object.getPrototypeOf(value) !== Object.prototype) return value;
    for (const symbol of Object.getOwnPropertySymbols(value))
      assert.ok(
        symbol.description && Object.hasOwn(SOURCE_SYMBOLS, symbol.description),
        `Unclassified model capture symbol: ${String(symbol)}`,
      );
    const entries: [string, unknown][] = [];
    for (const [key, entry] of Object.entries(value)) {
      const field = SOURCE_FIELDS.get(key);
      if (field) {
        if (field.disposition === "exclude") continue;
      } else {
        assert.ok(
          !/capture|verbatim|namespace(?:Declarations|Bindings)|source.*(?:Xml|Token|Fingerprint)|fingerprint/i.test(
            key,
          ),
          `Unclassified model capture field: ${key}`,
        );
      }
      if (entry === undefined) continue;
      if (key === "comments" && Array.isArray(entry) && entry.length === 0) continue;
      if (
        (key === "formatting" || key === "previousFormatting" || key === "currentFormatting") &&
        typeof entry === "object" &&
        entry !== null &&
        Object.getPrototypeOf(entry) === Object.prototype
      ) {
        const keys = Reflect.ownKeys(entry);
        if (
          keys.every((fieldKey) =>
            typeof fieldKey === "string"
              ? SOURCE_FIELDS.get(fieldKey)?.disposition === "exclude"
              : fieldKey.description !== undefined &&
                Object.hasOwn(SOURCE_SYMBOLS, fieldKey.description),
          )
        )
          continue;
      }
      if (key === "relationships" && entry instanceof Map && !hasModeledComments(value)) {
        const meaningful = new Map(
          [...entry].filter(
            ([, relationship]) =>
              typeof relationship !== "object" ||
              relationship === null ||
              !("type" in relationship) ||
              typeof relationship.type !== "string" ||
              !EMPTY_PART_RELATIONSHIPS.has(relationship.type),
          ),
        );
        entries.push([key, project(meaningful, `${modelPath}/${key}`)]);
        continue;
      }
      entries.push([
        key,
        key === "paraId" && typeof entry === "string"
          ? (aliases.get(entry) ?? entry)
          : project(entry, `${modelPath}/${key}`),
      ]);
    }
    const projected = canonicalFieldModel(Object.fromEntries(entries));
    if (
      "type" in value &&
      value.type === "paragraph" &&
      (!("paraId" in value) ||
        value.paraId === undefined ||
        (typeof value.paraId === "string" && aliases.has(value.paraId)))
    ) {
      // Only identities explicitly declared positional by the shared row pairing
      // may be regenerated. Authored paragraph ids remain unchanged and strict.
      projected["paraId"] = `\u0000model:${modelPath}`;
    }
    return projected;
  };
  return project(model, "");
};

type ReviewRoundTripOptions = { live: ReviewSource; reopened: ReviewSource; context: string };

/** The save invariant and reader stability share coverage and both resolution outcomes. */
export const assertReviewRoundTrip = ({
  live,
  reopened,
  context,
}: ReviewRoundTripOptions): Promise<void> =>
  Promise.resolve().then(() => {
    const identities = projectContentPair({
      leftRows: live.getContent(),
      rightRows: reopened.getContent(),
    });
    const leftChanges = live.getChanges();
    const rightChanges = reopened.getChanges();
    assert.deepEqual(
      revisionCoverage(rightChanges, identities.rightAliases),
      revisionCoverage(leftChanges, identities.leftAliases),
      `${context}: revision coverage differs after save`,
    );
    if (leftChanges.length === 0 && rightChanges.length === 0) return undefined;
    for (const decision of ["accept", "reject"] as const) {
      const left = live.toResolvedDocument(decision);
      const right = reopened.toResolvedDocument(decision);
      assert.deepEqual(
        canonicalReviewModel(right, identities.rightAliases),
        canonicalReviewModel(left, identities.leftAliases),
        `${context}: ${decision}-all document differs after save`,
      );
    }
    return undefined;
  });
