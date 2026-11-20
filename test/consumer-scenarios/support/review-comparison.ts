import assert from "node:assert/strict";

import type {
  Document,
  Paragraph,
  ParagraphContent,
} from "@stll/docx-core/model";
import { formattingEquals, mergeRunContent } from "@stll/docx-core/ops";
import { visitDocxParagraphs } from "@stll/folio-core/docx/paragraphTraversal";
import { RELATIONSHIP_TYPES } from "@stll/folio-core/docx/relsParser";
import { consolidateParagraphContent } from "@stll/folio-core/docx/runConsolidator";

import { openReviewer } from "./documents.ts";
import { canonicalFieldModel } from "./field-model.ts";
import { projectContentPair } from "./identity.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type ReviewSource = Pick<
  Reviewer,
  "getContent" | "getChanges" | "toResolvedDocument"
>;
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

export const assertRevisionCoverage = (
  left: readonly Change[],
  right: readonly Change[],
) => assert.deepEqual(revisionCoverage(right), revisionCoverage(left));

// Source replay hints are not authored document properties. Every such field
// needs an explicit disposition; unknown capture fields or symbols fail closed.
const SOURCE_FIELDS = new Map(
  Object.entries({
    requiredFonts: {
      disposition: "exclude",
      reason: "Derived font-loading inventory",
    },
    templateVariables: {
      disposition: "exclude",
      reason: "Derived template-token inventory; authored text stays strict",
    },
    warnings: { disposition: "exclude", reason: "Source parse diagnostics" },
    parseWarnings: {
      disposition: "exclude",
      reason: "Source parse diagnostics",
    },
    modified: {
      disposition: "exclude",
      reason: "Package save bookkeeping; revision dates remain strict",
    },
    originalBuffer: {
      disposition: "exclude",
      reason: "Original ZIP replay source",
    },
    verbatimXml: {
      disposition: "exclude",
      reason: "Replay bytes for a modeled carrier",
    },
    verbatimFingerprint: {
      disposition: "exclude",
      reason: "Replay cache validity",
    },
    verbatimStructureFingerprint: {
      disposition: "exclude",
      reason: "Replay cache validity of a captured shape's structure",
    },
    namespaceDeclarations: {
      disposition: "exclude",
      reason: "Lexical namespace spelling",
    },
    namespaceBindings: {
      disposition: "exclude",
      reason: "Lexical namespace spelling",
    },
    sourceXml: {
      disposition: "exclude",
      reason: "Source property XML replay; modeled properties stay strict",
    },
    gridSourceXml: {
      disposition: "exclude",
      reason: "Source grid XML replay; modeled grid columns stay strict",
    },
    rawWatermarkXml: {
      disposition: "keep",
      reason: "Watermark content not fully modeled",
    },
    rawImageXml: {
      disposition: "keep",
      reason: "Drawing content not fully modeled",
    },
    rawImageFingerprint: {
      disposition: "exclude",
      reason: "Drawing replay cache validity",
    },
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
  if (
    !("document" in value) ||
    typeof value.document !== "object" ||
    value.document === null
  )
    return true;
  return (
    "comments" in value.document &&
    Array.isArray(value.document.comments) &&
    value.document.comments.length > 0
  );
};

/**
 * Parts a package finds by relationship type alone (ECMA-376 implicit
 * relationships). No content names them by id, so the id a save allocates
 * them is no part of what the package reads as.
 */
const IMPLICIT_PART_RELATIONSHIPS = new Set<string>([
  RELATIONSHIP_TYPES.styles,
  RELATIONSHIP_TYPES.numbering,
  RELATIONSHIP_TYPES.settings,
  RELATIONSHIP_TYPES.webSettings,
  RELATIONSHIP_TYPES.fontTable,
  RELATIONSHIP_TYPES.theme,
  RELATIONSHIP_TYPES.footnotes,
  RELATIONSHIP_TYPES.endnotes,
  RELATIONSHIP_TYPES.comments,
  RELATIONSHIP_TYPES.commentsExtended,
]);

const relationshipType = (relationship: unknown): string | undefined =>
  typeof relationship === "object" &&
  relationship !== null &&
  "type" in relationship &&
  typeof relationship.type === "string"
    ? relationship.type
    : undefined;

type MeaningfulRelationshipsOptions = {
  relationships: ReadonlyMap<unknown, unknown>;
  modeledComments: boolean;
};

/** Relationships keyed by what content reads them through: id, or type and target. */
const meaningfulRelationships = ({
  relationships,
  modeledComments,
}: MeaningfulRelationshipsOptions): Map<unknown, unknown> => {
  const meaningful = new Map<unknown, unknown>();
  for (const [id, relationship] of relationships) {
    const type = relationshipType(relationship);
    if (
      !modeledComments &&
      type !== undefined &&
      EMPTY_PART_RELATIONSHIPS.has(type)
    )
      continue;
    if (
      type === undefined ||
      !IMPLICIT_PART_RELATIONSHIPS.has(type) ||
      typeof relationship !== "object" ||
      relationship === null
    ) {
      meaningful.set(id, relationship);
      continue;
    }
    const rest = Object.fromEntries(
      Object.entries(relationship).filter(([field]) => field !== "id"),
    );
    const key = `implicit ${type} ${String(rest["target"])}`;
    let duplicate = 1;
    while (meaningful.has(`${key} #${duplicate}`)) duplicate++;
    meaningful.set(`${key} #${duplicate}`, rest);
  }
  return meaningful;
};

type CanonicalReviewModelOptions = {
  model: unknown;
  aliases?: ReadonlyMap<string, string>;
  paragraphContentOverrides?: WeakMap<object, ParagraphContent[]>;
};

/** One strict modeled-document projection, including nested part collections. */
const projectCanonicalReviewModel = ({
  model,
  aliases = new Map(),
  paragraphContentOverrides,
}: CanonicalReviewModelOptions): unknown => {
  const project = (current: unknown, modelPath: string): unknown => {
    const value = current;
    if (Array.isArray(value))
      return value.map((entry, index) =>
        project(entry, `${modelPath}/${index}`),
      );
    if (value instanceof Map)
      return new Map(
        [...value].map(([key, entry]) => [
          key,
          project(entry, `${modelPath}/${String(key)}`),
        ]),
      );
    if (value instanceof Set)
      return new Set([...value].map((entry) => project(entry, modelPath)));
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
      if (key === "comments" && Array.isArray(entry) && entry.length === 0)
        continue;
      if (
        // An `alternateContent` holding only replay hints (the captured XML
        // and its fingerprints) models nothing a reader can see.
        (key === "formatting" ||
          key === "alternateContent" ||
          key === "previousFormatting" ||
          key === "currentFormatting") &&
        typeof entry === "object" &&
        entry !== null &&
        Object.getPrototypeOf(entry) === Object.prototype
      ) {
        const keys = Reflect.ownKeys(entry);
        if (
          keys.length > 0 &&
          keys.every((fieldKey) =>
            typeof fieldKey === "string"
              ? SOURCE_FIELDS.get(fieldKey)?.disposition === "exclude"
              : fieldKey.description !== undefined &&
                Object.hasOwn(SOURCE_SYMBOLS, fieldKey.description),
          )
        )
          continue;
      }
      if (key === "relationships" && entry instanceof Map) {
        const meaningful = meaningfulRelationships({
          relationships: entry,
          modeledComments: hasModeledComments(value),
        });
        entries.push([key, project(meaningful, `${modelPath}/${key}`)]);
        continue;
      }
      const originalProjection = project(entry, `${modelPath}/${key}`);
      const paragraphContentOverride = paragraphContentOverrides?.get(value);
      const projectedEntry =
        key === "content" && paragraphContentOverride !== undefined
          ? project(paragraphContentOverride, `${modelPath}/${key}`)
          : originalProjection;
      entries.push([
        key,
        key === "paraId" && typeof projectedEntry === "string"
          ? (aliases.get(projectedEntry) ?? projectedEntry)
          : projectedEntry,
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

export const canonicalReviewModel = (
  model: unknown,
  aliases: ReadonlyMap<string, string> = new Map(),
): unknown => projectCanonicalReviewModel({ model, aliases });

type Run = Extract<ParagraphContent, { type: "run" }>;

/** Whether `a` and `b` are runs whose properties a reader cannot tell apart. */
const sameRunProperties = (a: Run, b: Run): boolean =>
  a.content.length > 0 &&
  b.content.length > 0 &&
  (a.propertyChanges?.length ?? 0) === 0 &&
  (b.propertyChanges?.length ?? 0) === 0 &&
  formattingEquals(a.formatting, b.formatting) &&
  JSON.stringify(a.preservedAttributes ?? null) ===
    JSON.stringify(b.preservedAttributes ?? null);

/**
 * `content` with adjacent runs of the same properties read as one run.
 * Where one `w:r` ends and the next begins carries nothing a reader sees
 * when both have the same properties: their text, breaks and tabs are the
 * same content in one run or two. `consolidateParagraphContent` keeps a run
 * holding a break or a tab apart, as an editing boundary; a resolved
 * document is compared by what it holds.
 */
const joinRunsOfSameProperties = (
  content: readonly ParagraphContent[],
): ParagraphContent[] => {
  const joined: ParagraphContent[] = [];
  for (const item of content) {
    const entry: ParagraphContent =
      item.type === "hyperlink"
        ? {
            ...item,
            children: joinRunsOfSameProperties(
              item.children,
            ) as typeof item.children,
          }
        : item;
    const previous = joined.at(-1);
    if (
      previous?.type === "run" &&
      entry.type === "run" &&
      sameRunProperties(previous, entry)
    ) {
      joined[joined.length - 1] = {
        ...previous,
        content: mergeRunContent(previous.content, entry.content),
      };
      continue;
    }
    joined.push(entry);
  }
  return joined;
};

/** Package fields that hold no numbering reference a reader follows. */
const NUMBERING_REFERENCE_FREE_KEYS = new Set(["numbering", "media", "fonts"]);

/** Every `numId` a numbering reference outside the numbering part names. */
const referencedNumIds = (model: Document): Set<number> => {
  const referenced = new Set<number>();
  const visited = new WeakSet<object>();
  const visit = (value: unknown): void => {
    if (typeof value !== "object" || value === null) return;
    if (value instanceof ArrayBuffer || ArrayBuffer.isView(value)) return;
    if (visited.has(value)) return;
    visited.add(value);
    if (value instanceof Map) {
      for (const entry of value.values()) visit(entry);
      return;
    }
    if (
      "kind" in value &&
      value.kind === "reference" &&
      "numId" in value &&
      typeof value.numId === "number"
    )
      referenced.add(value.numId);
    for (const child of Object.values(value)) visit(child);
  };
  for (const [key, value] of Object.entries(model.package))
    if (!NUMBERING_REFERENCE_FREE_KEYS.has(key)) visit(value);
  return referenced;
};

/**
 * `model` with only the numbering definitions a reference reaches: the
 * instances a paragraph, style or recorded previous formatting names, their
 * abstract definitions, and the abstract a level-less one links to by style.
 * A definition nothing reaches changes nothing a reader sees; a save may
 * drop one a session added and no longer uses, and with it a numbering part
 * left with no definition a reference reaches.
 */
const withReachableNumbering = (model: Document): Document => {
  const { numbering } = model.package;
  if (numbering === undefined) return model;
  const referenced = referencedNumIds(model);
  const nums = numbering.nums.filter(({ numId }) => referenced.has(numId));
  const abstractIds = new Set(nums.map(({ abstractNumId }) => abstractNumId));
  const styleLinks = new Set(
    numbering.abstractNums.flatMap(({ abstractNumId, numStyleLink }) =>
      abstractIds.has(abstractNumId) && numStyleLink !== undefined
        ? [numStyleLink]
        : [],
    ),
  );
  const abstractNums = numbering.abstractNums.filter(
    ({ abstractNumId, styleLink }) =>
      abstractIds.has(abstractNumId) ||
      (styleLink !== undefined && styleLinks.has(styleLink)),
  );
  if (nums.length > 0 || abstractNums.length > 0)
    return {
      ...model,
      package: {
        ...model.package,
        numbering: { ...numbering, nums, abstractNums },
      },
    };
  const { relationships } = model.package;
  return {
    ...model,
    package: {
      ...model.package,
      numbering: undefined,
      relationships:
        relationships &&
        new Map(
          [...relationships].filter(
            ([, { type }]) => type !== RELATIONSHIP_TYPES.numbering,
          ),
        ),
    },
  };
};

export const canonicalResolvedReviewModel = (
  model: Document,
  aliases: ReadonlyMap<string, string> = new Map(),
): unknown => {
  const paragraphContentOverrides = new WeakMap<object, ParagraphContent[]>();
  const seenParagraphs = new WeakSet<Paragraph>();
  const visitParagraph = (paragraph: Paragraph): void => {
    if (seenParagraphs.has(paragraph)) return;
    seenParagraphs.add(paragraph);
    paragraphContentOverrides.set(
      paragraph,
      joinRunsOfSameProperties(consolidateParagraphContent(paragraph.content)),
    );
  };
  visitDocxParagraphs(
    {
      documentBody: model.package.document,
      headers: model.package.headers,
      footers: model.package.footers,
      footnotes: model.package.footnotes,
      endnotes: model.package.endnotes,
    },
    visitParagraph,
  );
  for (const section of model.package.document.sections ?? []) {
    visitDocxParagraphs(
      {
        documentBody: { content: [] },
        headers: section.headers,
        footers: section.footers,
      },
      visitParagraph,
    );
  }
  return projectCanonicalReviewModel({
    model: withReachableNumbering(model),
    aliases,
    paragraphContentOverrides,
  });
};

type ReviewRoundTripOptions = {
  live: ReviewSource;
  reopened: ReviewSource;
  context: string;
};

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
        canonicalResolvedReviewModel(right, identities.rightAliases),
        canonicalResolvedReviewModel(left, identities.leftAliases),
        `${context}: ${decision}-all document differs after save`,
      );
    }
    return undefined;
  });
