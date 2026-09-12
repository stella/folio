/**
 * What the round-trip self-check found, and how to say it safely.
 *
 * The check compares two block projections: what accepting the generated
 * revisions leaves against the target, and what rejecting them leaves against
 * the base. A projection carries each block's container, style, list level,
 * direct alignment, direct spacing and text, so WHICH field diverged names
 * which part of the pipeline lost the difference — and that is worth reporting
 * as a typed cause rather than as one opaque "did not reproduce".
 *
 * Every `detail` string here is structural: counts, offsets, container kinds.
 * Never a phrase of either document, because a caller may log it, put it in a
 * report, or quote it in a review.
 */

import { PARAGRAPH_MARK_CHANGE_KINDS, type ParagraphMarkChangeKind } from "@stll/docx-core/model";
import { panic } from "better-result";

import type { FolioDocumentStoryHandle } from "../ai-edits/headless";
import type {
  FolioContentBlock,
  FolioContentPropertySet,
  FolioContentRun,
  FolioContentTableLocation,
} from "./content-types";
import {
  FOLIO_CONTENT_BLOCK_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_CONTAINER_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_IDENTITY_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_PARAGRAPH_FORMATTING_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_RUN_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_STRUCTURAL_BOUNDARY_FIELD_DESCRIPTORS,
  FOLIO_CONTENT_TABLE_FIELD_DESCRIPTORS,
} from "./content-types";

/** The two directions of the round trip, each an invariant of its own. */
export const COMPARE_VERIFICATION_INVARIANTS = Object.freeze([
  "accept-reproduces-target",
  "reject-reproduces-base",
] as const);

export type CompareVerificationInvariant = (typeof COMPARE_VERIFICATION_INVARIANTS)[number];

/**
 * What diverged. Ordered from the structural to the textual: the first cause
 * that fits is the one reported, because a block in the wrong container is a
 * different finding from a block whose words are wrong even when both are true.
 */
export const COMPARE_VERIFICATION_CAUSES = Object.freeze([
  /**
   * Every block is present, in order, at coordinates the block model cannot
   * reach, so no operation could have put a block at the expected ones. Not a
   * lost difference.
   *
   * Blank paragraphs no longer cause this: the snapshot carries them. What is
   * left is the row a package hides, whose whole subtree the snapshot skips on
   * purpose. A table hiding a row on one side only shifts every later row's
   * index, and nothing the comparison can do reaches those positions.
   */
  "invisible-structure",
  "block-count",
  "container",
  /** Explicit inline carriers differ in count, order, position, or semantics. */
  "inline-structure",
  /**
   * Every block is where it should be and a table's own properties are not:
   * `w:tblPr`, the `w:tblGrid` widths, `w:trPr`, `w:tcPr`. No block carries
   * them, so a projection of blocks alone cannot see them go.
   */
  "table-geometry",
  "style",
  "list-level",
  "alignment",
  "spacing",
  "inline-formatting",
  /** A known semantic difference has no supported transport instruction. */
  "unsupported",
  "whitespace",
  "text",
] as const);

export type CompareVerificationCause = (typeof COMPARE_VERIFICATION_CAUSES)[number];

/** The package or exact story whose round-trip invariant did not hold. */
export type CompareVerificationScope =
  | { readonly type: "package" }
  | { readonly type: "story"; readonly story: FolioDocumentStoryHandle };

/** One invariant that did not hold at its typed package or story scope. */
export type CompareVerificationFailure = {
  invariant: CompareVerificationInvariant;
  cause: CompareVerificationCause;
  scope: CompareVerificationScope;
  /** Structural facts only: counts, offsets, container kinds. Safe to quote. */
  detail: string;
};

/**
 * Whether the redline was proven to round-trip.
 *
 * `unverified` is only ever returned when the caller asked for it with
 * `mode: "bestEffort"`; the default refuses instead, because a redline that
 * reads plausibly and is wrong is worse than no redline.
 */
export type CompareVerification =
  | { status: "verified" }
  | { status: "unverified"; failures: readonly CompareVerificationFailure[] };

const collapseWhitespace = (text: string): string => text.replace(/\s+/gu, " ").trim();

const sameStringProjection = (left: readonly string[], right: readonly string[]): boolean =>
  left.length === right.length && left.every((entry, index) => entry === right[index]);

type ClassifyOptions<T> = {
  invariant: CompareVerificationInvariant;
  story: FolioDocumentStoryHandle;
  /** What the redline actually leaves. */
  actual: readonly T[];
  /** What the invariant says it should leave. */
  expected: readonly T[];
};

/**
 * The failure two table-geometry projections describe, or `null` when they
 * agree. One line per table, so the detail names which table diverged and
 * whether the count itself did — never a property value, which could carry a
 * style name either document chose.
 */
export const classifyGeometryMismatch = ({
  invariant,
  story,
  actual,
  expected,
}: ClassifyOptions<string>): CompareVerificationFailure | null => {
  if (sameStringProjection(actual, expected)) {
    return null;
  }
  const detail =
    actual.length === expected.length
      ? `table ${String(actual.findIndex((entry, index) => entry !== expected[index]))} of ${String(actual.length)} carries different properties`
      : `${String(actual.length)} tables against ${String(expected.length)}`;
  return { invariant, cause: "table-geometry", scope: { type: "story", story }, detail };
};

const sameCanonicalValue = (left: unknown, right: unknown): boolean =>
  JSON.stringify(left) === JSON.stringify(right);

const sameCanonicalIdentity = (
  left: FolioContentBlock["identity"],
  right: FolioContentBlock["identity"],
): boolean =>
  Object.values(FOLIO_CONTENT_IDENTITY_FIELD_DESCRIPTORS).every((descriptor) =>
    sameCanonicalValue(Reflect.get(left, descriptor.field), Reflect.get(right, descriptor.field)),
  );

const sameCanonicalTableLocation = (
  left: FolioContentTableLocation | undefined,
  right: FolioContentTableLocation | undefined,
): boolean => {
  if (!left || !right) return left === right;
  for (const descriptor of Object.values(FOLIO_CONTENT_TABLE_FIELD_DESCRIPTORS)) {
    switch (descriptor.verification) {
      case "transport-identity":
        break;
      case "exact":
        if (
          !sameCanonicalValue(
            Reflect.get(left, descriptor.field),
            Reflect.get(right, descriptor.field),
          )
        ) {
          return false;
        }
        break;
      default: {
        const exhaustive: never = descriptor;
        return exhaustive;
      }
    }
  }
  return true;
};

const sameCanonicalContainerPath = (
  left: FolioContentBlock["containerPath"],
  right: FolioContentBlock["containerPath"],
): boolean =>
  left.length === right.length &&
  left.every((entry, index) => {
    const other = right[index];
    if (!other) return false;
    return Object.values(FOLIO_CONTENT_CONTAINER_FIELD_DESCRIPTORS).every((descriptor) => {
      if (descriptor.field === "identity") {
        return sameCanonicalIdentity(entry.identity, other.identity);
      }
      return sameCanonicalValue(
        Reflect.get(entry, descriptor.field),
        Reflect.get(other, descriptor.field),
      );
    });
  });

const sameCanonicalRuns = (
  left: readonly FolioContentRun[],
  right: readonly FolioContentRun[],
): boolean =>
  left.length === right.length &&
  left.every((run, index) => {
    const other = right[index];
    if (!other) return false;
    return Object.values(FOLIO_CONTENT_RUN_FIELD_DESCRIPTORS).every((descriptor) =>
      sameCanonicalValue(Reflect.get(run, descriptor.field), Reflect.get(other, descriptor.field)),
    );
  });

const sameCanonicalParagraphFormatting = (
  left: FolioContentBlock["paragraphFormatting"],
  right: FolioContentBlock["paragraphFormatting"],
): boolean =>
  Object.values(FOLIO_CONTENT_PARAGRAPH_FORMATTING_FIELD_DESCRIPTORS).every((descriptor) =>
    sameCanonicalValue(Reflect.get(left, descriptor.field), Reflect.get(right, descriptor.field)),
  );

const sameCanonicalStructuralBoundaries = (
  left: FolioContentBlock["structuralBoundaries"],
  right: FolioContentBlock["structuralBoundaries"],
): boolean =>
  left.length === right.length &&
  left.every((boundary, index) => {
    const other = right[index];
    if (!other) return false;
    return Object.values(FOLIO_CONTENT_STRUCTURAL_BOUNDARY_FIELD_DESCRIPTORS).every((descriptor) =>
      sameCanonicalValue(
        Reflect.get(boundary, descriptor.field),
        Reflect.get(other, descriptor.field),
      ),
    );
  });

const sameCanonicalProperties = (
  left: FolioContentPropertySet,
  right: FolioContentPropertySet,
): boolean => sameCanonicalValue(left, right);

const paragraphPropertyChanged = (
  left: FolioContentPropertySet,
  right: FolioContentPropertySet,
  keys: ReadonlySet<string>,
): boolean => {
  const selected = (properties: FolioContentPropertySet): FolioContentPropertySet =>
    properties.filter(({ key }) => keys.has(key));
  return !sameCanonicalProperties(selected(left), selected(right));
};

const SPACING_PROPERTY_KEYS = new Set([
  "spaceBefore",
  "spaceAfter",
  "lineSpacing",
  "lineSpacingRule",
  "beforeAutospacing",
  "afterAutospacing",
]);

type CanonicalClassifyOptions = Omit<ClassifyOptions<FolioContentBlock>, "actual" | "expected"> & {
  actual: readonly FolioContentBlock[];
  expected: readonly FolioContentBlock[];
};

/**
 * Compare the complete owned neutral projection produced from a live DOCX
 * model. Stable transport ids are deliberately excluded: accepting a change
 * preserves the base package's ids while the semantic target may carry other
 * ones. Every modeled content, presentation, and ownership field is checked.
 */
export const classifyContentProjectionMismatch = ({
  invariant,
  story,
  actual,
  expected,
}: CanonicalClassifyOptions): CompareVerificationFailure | null => {
  const failure = (
    cause: CompareVerificationCause,
    detail: string,
  ): CompareVerificationFailure => ({
    invariant,
    cause,
    scope: { type: "story", story },
    detail,
  });
  if (actual.length !== expected.length) {
    return failure(
      "block-count",
      `${String(actual.length)} blocks against ${String(expected.length)}`,
    );
  }
  for (const [index, right] of expected.entries()) {
    const left = actual[index];
    if (!left) {
      return failure("block-count", `a block is missing at index ${String(index)}`);
    }
    const at = `at block ${String(index)}/${String(expected.length)}`;
    const mismatches = new Set<keyof FolioContentBlock>();
    for (const descriptor of Object.values(FOLIO_CONTENT_BLOCK_FIELD_DESCRIPTORS)) {
      switch (descriptor.verification) {
        case "transport-identity":
          break;
        case "exact":
          if (
            !sameCanonicalValue(
              Reflect.get(left, descriptor.field),
              Reflect.get(right, descriptor.field),
            )
          ) {
            mismatches.add(descriptor.field);
          }
          break;
        case "container":
          if (!sameCanonicalContainerPath(left.containerPath, right.containerPath)) {
            mismatches.add(descriptor.field);
          }
          break;
        case "nested": {
          const same = (() => {
            switch (descriptor.field) {
              case "paragraphFormatting":
                return sameCanonicalParagraphFormatting(
                  left.paragraphFormatting,
                  right.paragraphFormatting,
                );
              case "runs":
                return sameCanonicalRuns(left.runs, right.runs);
              case "structuralBoundaries":
                return sameCanonicalStructuralBoundaries(
                  left.structuralBoundaries,
                  right.structuralBoundaries,
                );
              case "table":
                return sameCanonicalTableLocation(left.table, right.table);
              default: {
                const exhaustive: never = descriptor;
                return panic("A nested content-verification field has no verifier", {
                  descriptor: exhaustive,
                });
              }
            }
          })();
          if (!same) mismatches.add(descriptor.field);
          break;
        }
        default: {
          const exhaustive: never = descriptor;
          return exhaustive;
        }
      }
    }
    if (mismatches.size === 0) continue;
    if (mismatches.has("table") || mismatches.has("containerPath")) {
      return failure("container", `a block's canonical ownership differs ${at}`);
    }
    if (mismatches.has("structuralBoundaries")) {
      return failure("inline-structure", `a block's inline structure differs ${at}`);
    }
    if (mismatches.has("kind") || mismatches.has("blockProperties")) {
      return failure("unsupported", `a block's modeled semantics differ ${at}`);
    }
    const leftParagraph = left.paragraphFormatting;
    const rightParagraph = right.paragraphFormatting;
    if (mismatches.has("paragraphFormatting")) {
      if (
        paragraphPropertyChanged(
          leftParagraph.authored,
          rightParagraph.authored,
          new Set(["styleId"]),
        )
      ) {
        return failure("style", `the paragraph style differs ${at}`);
      }
      if (
        paragraphPropertyChanged(
          leftParagraph.authored,
          rightParagraph.authored,
          new Set(["numPr"]),
        )
      ) {
        return failure("list-level", `the paragraph numbering differs ${at}`);
      }
      if (
        paragraphPropertyChanged(
          leftParagraph.authored,
          rightParagraph.authored,
          new Set(["alignment"]),
        )
      ) {
        return failure("alignment", `the paragraph alignment differs ${at}`);
      }
      if (
        paragraphPropertyChanged(
          leftParagraph.authored,
          rightParagraph.authored,
          SPACING_PROPERTY_KEYS,
        )
      ) {
        return failure("spacing", `the paragraph spacing differs ${at}`);
      }
      return failure("unsupported", `the paragraph presentation differs ${at}`);
    }
    if (mismatches.has("runs")) {
      return failure("inline-formatting", `the inline presentation differs ${at}`);
    }
    if (mismatches.has("text")) {
      if (collapseWhitespace(left.text) === collapseWhitespace(right.text)) {
        return failure("whitespace", `a block's text differs only in whitespace ${at}`);
      }
      return failure(
        "text",
        `a block's text length ${String(left.text.length)} differs from ${String(right.text.length)} ${at}`,
      );
    }
    return failure("unsupported", `an unclassified modeled block field differs ${at}`);
  }
  return null;
};

/**
 * A container whose last paragraph carries a revision on its mark.
 *
 * Structural facts only — a path through the package model and an index — so
 * the finding is safe to log, report or quote.
 */
export type FinalParagraphMarkRevision = {
  /** Where the container sits in the package model, e.g. `package.document.content`. */
  container: string;
  /** The paragraph's index among its container's children. */
  paragraphIndex: number;
  /** The mark kind found there, `moveFrom` / `moveTo` for a relocation's ends. */
  kind: ParagraphMarkChangeKind;
};

const isAParagraphMarkChangeKind = (value: unknown): value is ParagraphMarkChangeKind =>
  PARAGRAPH_MARK_CHANGE_KINDS.some((kind) => kind === value);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const isParagraph = (value: unknown): value is Record<string, unknown> =>
  isRecord(value) && value["type"] === "paragraph";

/**
 * A row the package marks deleted, whose cells' marks go with it.
 *
 * A deleted table row is written as `w:trPr/w:del` PLUS a deletion on every
 * mark the row's cells end with: accepting takes the whole row away, so those
 * marks never have to join anything and are the shape the format asks for.
 */
const isADeletedTableRow = (value: Record<string, unknown>): boolean => {
  const change = value["structuralChange"];
  return isRecord(change) && change["type"] === "tableRowDeletion";
};

/**
 * Every container in a package whose final paragraph mark carries a revision.
 *
 * A deleted paragraph mark says "join this paragraph with the one after it",
 * and an inserted one says that break was ADDED, so rejecting it closes the
 * paragraph back over the next one. The last paragraph of a body, a table
 * cell, a header or footer, a note or a text box has no paragraph after it, so
 * neither direction states an edit that can be carried out: a consumer refuses
 * the package, or opens it and leaves a revision standing that neither
 * accepting nor rejecting everything can clear. The exception is a cell of a
 * row the package is DELETING: there the mark leaves with its row.
 *
 * The walk is over the package model rather than over a list of the containers
 * known today: a container is any sequence that ends in a paragraph, so a part
 * the model grows later is covered the day it arrives instead of the day
 * someone remembers this function.
 *
 * `since` scopes it to the revisions a comparison MINTED: a base may arrive
 * carrying one of these on a paragraph in a part no story mounts, which folio
 * preserves the way it preserves everything else it parses. What this proves
 * is that the comparison writes none of its own.
 */
export const revisedFinalParagraphMarks = (
  packageModel: unknown,
  { since = 0 }: { since?: number } = {},
): FinalParagraphMarkRevision[] => {
  const found: FinalParagraphMarkRevision[] = [];
  // The path is kept as a stack of raw keys and formatted only where something
  // is found. Building the string at every node instead made the walk cost
  // more than the serialization it guards, on a package where it never has
  // anything to report.
  const trail: (string | number)[] = ["package"];
  const pathOf = (): string => {
    let path = "";
    for (const segment of trail) {
      path += typeof segment === "number" ? `[${String(segment)}]` : `.${segment}`;
    }
    return path.slice(1);
  };
  const visit = (value: unknown, insideADeletedRow: boolean): void => {
    if (Array.isArray(value)) {
      const last: unknown = value.at(-1);
      const mark = isParagraph(last) ? last["pPrMark"] : undefined;
      const kind = isRecord(mark) ? mark["kind"] : undefined;
      const info = isRecord(mark) ? mark["info"] : undefined;
      const revisionId = isRecord(info) ? info["id"] : undefined;
      const isOurs = typeof revisionId === "number" ? revisionId >= since : true;
      if (isAParagraphMarkChangeKind(kind) && isOurs && !insideADeletedRow) {
        found.push({ container: pathOf(), paragraphIndex: value.length - 1, kind });
      }
      for (const [index, item] of value.entries()) {
        trail.push(index);
        visit(item, insideADeletedRow);
        trail.pop();
      }
      return;
    }
    if (value instanceof Map) {
      for (const [key, item] of value) {
        trail.push(String(key));
        visit(item, insideADeletedRow);
        trail.pop();
      }
      return;
    }
    if (!isRecord(value) || value instanceof Date || ArrayBuffer.isView(value)) {
      return;
    }
    const inADeletedRow = insideADeletedRow || isADeletedTableRow(value);
    for (const [key, item] of Object.entries(value)) {
      trail.push(key);
      visit(item, inADeletedRow);
      trail.pop();
    }
  };
  visit(packageModel, false);
  return found;
};
