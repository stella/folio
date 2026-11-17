/**
 * Numbering on demand. `literal` writes the label each numbered paragraph
 * shows (or a house-style label counted continuously) as text at the start
 * of the paragraph and removes its list membership, keeping the level's
 * indentation, tab stops and label run properties as direct formatting, so
 * the paragraph lays out as before and its number no longer depends on how a
 * consumer counts lists across sections. `apply` does the reverse: it puts
 * paragraphs into a numbering instance, optionally stripping a literal label
 * they start with.
 *
 * Labels come from the document model's reading (`displayLabel`,
 * `listReference`); the paragraphs are rewritten in the package XML, found by
 * their `w14:paraId`.
 */

import { Result } from "better-result";

import { SEQUENCE_CHILDREN } from "@stll/docx-core/schema";
import type { FolioDocxReviewer } from "@stll/folio-core/server";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import {
  attributeNamespace,
  childElements,
  createWElement,
  ensureWChild,
  insertChildInOrder,
  isW,
  localNameOf,
  mainDocumentPart,
  type OoxmlPackage,
  RELATIONSHIP_TYPES,
  setWAttribute,
  walkElements,
  wAttribute,
  type XmlElement,
} from "./ooxml-package";

export const NUMBERING_ACTIONS: readonly ["literal", "apply"] = [
  "literal",
  "apply",
] as const;

export const LABEL_SOURCES: readonly ["displayed", "continuous"] = [
  "displayed",
  "continuous",
] as const;

type LabelSource = (typeof LABEL_SOURCES)[number];

/** A paragraph a numbering action changed (or, as a dry run, would change). */
export type NumberedParagraph = {
  blockId: string;
  numId: number;
  level: number;
  /** `literal`: the label written. `apply`: the label stripped, if any. */
  label?: string;
};

export type NumberingReport = {
  action: (typeof NUMBERING_ACTIONS)[number];
  paragraphs: NumberedParagraph[];
};

const PARAGRAPH_PROPERTIES = SEQUENCE_CHILDREN["paragraph-properties"];
const RUN_PROPERTIES = SEQUENCE_CHILDREN["run-properties"];
const NUMBERING_PROPERTIES = ["ilvl", "numId", "numberingChange", "ins"];
const PARAGRAPH_CHILDREN = ["pPr"];

/** Run properties that record review history, never copied into a label. */
const REVISION_PROPERTIES: ReadonlySet<string> = new Set([
  "ins",
  "del",
  "moveFrom",
  "moveTo",
  "rPrChange",
]);

/** Level formats that show no number to write. */
const UNNUMBERED_FORMATS: ReadonlySet<string> = new Set(["bullet", "none"]);

const invalidInput = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message, hint });

const stringList = (
  value: unknown,
  name: string,
): Result<readonly string[] | null, FolioCliError> => {
  if (value === undefined) return Result.ok(null);
  return Array.isArray(value) &&
    value.every((entry) => typeof entry === "string")
    ? Result.ok(value)
    : Result.err(invalidInput(`${name} must be a list of strings.`));
};

const integerList = (
  value: unknown,
  name: string,
): Result<readonly number[] | null, FolioCliError> => {
  if (value === undefined) return Result.ok(null);
  return Array.isArray(value) && value.every((entry) => Number.isInteger(entry))
    ? Result.ok(value)
    : Result.err(invalidInput(`${name} must be a list of integers.`));
};

/** A deep copy that keeps each element's namespace, so lookups by URI still match. */
const cloneElement = (element: XmlElement): XmlElement => {
  const copy: XmlElement = {
    ...element,
    ...(element.attributes !== undefined && {
      attributes: { ...element.attributes },
    }),
    ...(element.elements !== undefined && {
      elements: element.elements.map(cloneElement),
    }),
  };
  for (const property of ["namespaceUri", "namespaceScope"] as const) {
    Object.defineProperty(copy, property, {
      value: element[property],
      enumerable: false,
    });
  }
  return copy;
};

const childW = (
  element: XmlElement | undefined,
  localName: string,
): XmlElement | undefined =>
  element === undefined
    ? undefined
    : childElements(element).find((child) => isW(child, localName));

/** The parts the actions read: the main document, numbering, and styles. */
type NumberingParts = {
  document: string;
  documentRoot: XmlElement;
  numbering: XmlElement | null;
  styles: XmlElement | null;
};

const partsOf = (pkg: OoxmlPackage): Result<NumberingParts, FolioCliError> => {
  const main = mainDocumentPart(pkg);
  if (main.isErr()) return Result.err(main.error);
  if (main.value === null) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidDocument,
        message: "The package has no main document.",
      }),
    );
  }
  const documentRoot = pkg.root(main.value);
  if (documentRoot.isErr()) return Result.err(documentRoot.error);
  if (documentRoot.value === null) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidDocument,
        message: "The main document part is missing.",
      }),
    );
  }
  const relationships = pkg.relationships(main.value);
  if (relationships.isErr()) return Result.err(relationships.error);
  const related = (type: string): Result<XmlElement | null, FolioCliError> => {
    const part = relationships.value.find(
      (relationship) => relationship.type === type,
    )?.part;
    return part === undefined || part === null
      ? Result.ok(null)
      : pkg.root(part);
  };
  const numbering = related(RELATIONSHIP_TYPES.numbering);
  if (numbering.isErr()) return Result.err(numbering.error);
  const styles = related(RELATIONSHIP_TYPES.styles);
  if (styles.isErr()) return Result.err(styles.error);
  return Result.ok({
    document: main.value,
    documentRoot: documentRoot.value,
    numbering: numbering.value,
    styles: styles.value,
  });
};

/** Every paragraph of the main document by its `w14:paraId`. */
const paragraphsById = (root: XmlElement): Map<string, XmlElement> => {
  const found = new Map<string, XmlElement>();
  walkElements(root, (element) => {
    if (!isW(element, "p")) return;
    const id = Object.entries(element.attributes ?? {}).find(
      ([attribute]) =>
        localNameOf(attribute) === "paraId" &&
        attributeNamespace(element, attribute) !== undefined,
    )?.[1];
    if (id !== undefined) found.set(String(id).toUpperCase(), element);
  });
  return found;
};

/** A numbering level's definition: the instance's override first, then its abstract definition. */
const levelDefinition = (
  numbering: XmlElement | null,
  numId: number,
  level: number,
): XmlElement | undefined => {
  if (numbering === null) return undefined;
  const instance = childElements(numbering).find(
    (element) =>
      isW(element, "num") && wAttribute(element, "numId") === String(numId),
  );
  if (instance === undefined) return undefined;
  const levelOf = (
    container: XmlElement | undefined,
  ): XmlElement | undefined =>
    container === undefined
      ? undefined
      : childElements(container).find(
          (element) =>
            isW(element, "lvl") &&
            wAttribute(element, "ilvl") === String(level),
        );
  const override = childElements(instance).find(
    (element) =>
      isW(element, "lvlOverride") &&
      wAttribute(element, "ilvl") === String(level),
  );
  const overridden = levelOf(override);
  if (overridden !== undefined) return overridden;
  const abstractId = wAttribute(childW(instance, "abstractNumId") ?? {}, "val");
  const abstract = childElements(numbering).find(
    (element) =>
      isW(element, "abstractNum") &&
      wAttribute(element, "abstractNumId") === abstractId,
  );
  return levelOf(abstract);
};

/** Whether a paragraph style, or one it is based on, states numbering. */
const styleNumbers = (
  styles: XmlElement | null,
  styleId: string | undefined,
): boolean => {
  if (styles === null) return false;
  const all = childElements(styles).filter((element) => isW(element, "style"));
  let current =
    styleId === undefined
      ? all.find(
          (style) =>
            wAttribute(style, "type") === "paragraph" &&
            ["1", "true", "on"].includes(wAttribute(style, "default") ?? ""),
        )
      : all.find((style) => wAttribute(style, "styleId") === styleId);
  const visited = new Set<XmlElement>();
  while (current !== undefined && !visited.has(current)) {
    visited.add(current);
    if (childW(childW(current, "pPr"), "numPr") !== undefined) return true;
    const basedOn = wAttribute(childW(current, "basedOn") ?? {}, "val");
    current =
      basedOn === undefined
        ? undefined
        : all.find((style) => wAttribute(style, "styleId") === basedOn);
  }
  return false;
};

const INDENT_START = ["start", "left"] as const;
const INDENT_FIRST = ["hanging", "firstLine"] as const;

/** Give the paragraph the level's indentation wherever it states none of its own. */
const keepLevelIndentation = (
  properties: XmlElement,
  level: XmlElement,
): void => {
  const levelIndent = childW(childW(level, "pPr"), "ind");
  if (levelIndent === undefined) return;
  const indent = ensureWChild(properties, "ind", PARAGRAPH_PROPERTIES);
  const has = (names: readonly string[]): boolean =>
    names.some((name) => wAttribute(indent, name) !== undefined);
  for (const group of [INDENT_START, INDENT_FIRST]) {
    if (has(group)) continue;
    for (const name of group) {
      const value = wAttribute(levelIndent, name);
      if (value !== undefined) setWAttribute(indent, name, value);
    }
  }
};

/** Give the paragraph the level's tab stops when it states none. */
const keepLevelTabs = (properties: XmlElement, level: XmlElement): void => {
  const levelTabs = childW(childW(level, "pPr"), "tabs");
  if (levelTabs === undefined || childW(properties, "tabs") !== undefined)
    return;
  insertChildInOrder(properties, cloneElement(levelTabs), PARAGRAPH_PROPERTIES);
};

/** The label run's properties: the paragraph mark's, overlaid with the level's. */
const labelRunProperties = (
  paragraph: XmlElement,
  properties: XmlElement,
  level: XmlElement | undefined,
): XmlElement | null => {
  const merged = createWElement(paragraph, "rPr");
  const sources = [childW(properties, "rPr"), childW(level, "rPr")];
  for (const source of sources) {
    for (const child of childElements(source ?? {})) {
      const name = localNameOf(child.name ?? "");
      if (REVISION_PROPERTIES.has(name)) continue;
      merged.elements = (merged.elements ?? []).filter(
        (existing) => localNameOf(existing.name ?? "") !== name,
      );
      insertChildInOrder(merged, cloneElement(child), RUN_PROPERTIES);
    }
  }
  return (merged.elements ?? []).length === 0 ? null : merged;
};

/** The run holding a label, with the separator the level's `w:suff` asks for. */
const labelRun = (
  paragraph: XmlElement,
  label: string,
  runProperties: XmlElement | null,
  suffix: string,
): XmlElement => {
  const run = createWElement(paragraph, "r");
  const text = createWElement(run, "t", {});
  text.attributes = { "xml:space": "preserve" };
  text.elements = [
    { type: "text", text: suffix === "space" ? `${label} ` : label },
  ];
  run.elements = [
    ...(runProperties === null ? [] : [runProperties]),
    text,
    ...(suffix === "tab" ? [createWElement(run, "tab")] : []),
  ];
  return run;
};

type LiteralTarget = {
  blockId: string;
  numId: number;
  level: number;
  label: string;
  inherits: boolean;
};

/** Which paragraphs `literal` rewrites, with the label each gets. */
const literalTargets = (
  reviewer: FolioDocxReviewer,
  args: Readonly<Record<string, unknown>>,
): Result<LiteralTarget[], FolioCliError> => {
  const source = args["labels"] ?? "displayed";
  if (!LABEL_SOURCES.some((known) => known === source)) {
    return Result.err(
      invalidInput(`labels must be ${LABEL_SOURCES.join(" or ")}.`),
    );
  }
  const labels: LabelSource =
    source === "continuous" ? "continuous" : "displayed";
  const format =
    args["format"] ?? (labels === "continuous" ? "{n}." : "{label}");
  if (typeof format !== "string" || !/\{(?:n|path|label)\}/u.test(format)) {
    return Result.err(
      invalidInput(
        "format must contain {n}, {path}, or {label}.",
        'For example "[{n}.]".',
      ),
    );
  }
  const numIds = integerList(args["numIds"], "numIds");
  if (numIds.isErr()) return Result.err(numIds.error);
  const blockIds = stringList(args["blockIds"], "blockIds");
  if (blockIds.isErr()) return Result.err(blockIds.error);
  const formats = new Map(
    reviewer
      .readNumberingDefinitions()
      .map(({ numId, level, format: levelFormat }) => [
        `${String(numId)}:${String(level)}`,
        levelFormat,
      ]),
  );
  const counters: number[] = [];
  const targets: LiteralTarget[] = [];
  for (const block of reviewer.snapshot().blocks) {
    if (block.kind === "diagnostic" || block.table !== undefined) continue;
    const reference = block.listReference;
    if (reference === undefined || block.displayLabel === undefined) continue;
    const levelFormat = formats.get(
      `${String(reference.numId)}:${String(reference.level)}`,
    );
    if (levelFormat === undefined || UNNUMBERED_FORMATS.has(levelFormat))
      continue;
    if (numIds.value !== null && !numIds.value.includes(reference.numId))
      continue;
    if (blockIds.value !== null && !blockIds.value.includes(block.id)) continue;
    counters.length = reference.level + 1;
    counters[reference.level] = (counters[reference.level] ?? 0) + 1;
    const path = Array.from(
      { length: reference.level + 1 },
      (_, index) => counters[index] ?? 1,
    );
    targets.push({
      blockId: block.id,
      numId: reference.numId,
      level: reference.level,
      label: format
        .replaceAll("{label}", block.displayLabel)
        .replaceAll("{n}", String(counters[reference.level]))
        .replaceAll("{path}", path.join(".")),
      inherits: block.statedNumbering.kind !== "reference",
    });
  }
  if (blockIds.value !== null) {
    const found = new Set(targets.map(({ blockId }) => blockId));
    const missing = blockIds.value.filter((id) => !found.has(id));
    if (missing.length > 0) {
      return Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.notFound,
          message: `No numbered paragraph with id ${missing.join(", ")}.`,
          hint: "Take ids from folio read; a bullet or an unnumbered paragraph has no number to write.",
        }),
      );
    }
  }
  return Result.ok(targets);
};

const literalNumbering = (
  pkg: OoxmlPackage,
  reviewer: FolioDocxReviewer,
  args: Readonly<Record<string, unknown>>,
): Result<NumberingReport, FolioCliError> => {
  const targets = literalTargets(reviewer, args);
  if (targets.isErr()) return Result.err(targets.error);
  const parts = partsOf(pkg);
  if (parts.isErr()) return Result.err(parts.error);
  const { document, documentRoot, numbering, styles } = parts.value;
  const paragraphs = paragraphsById(documentRoot);
  const report: NumberingReport = { action: "literal", paragraphs: [] };
  for (const target of targets.value) {
    const paragraph = paragraphs.get(target.blockId.toUpperCase());
    if (paragraph === undefined) continue;
    const properties = ensureWChild(paragraph, "pPr", PARAGRAPH_CHILDREN);
    const level = levelDefinition(numbering, target.numId, target.level);
    const runProperties = labelRunProperties(paragraph, properties, level);
    properties.elements = (properties.elements ?? []).filter(
      (child) => !isW(child, "numPr"),
    );
    const styleId = wAttribute(childW(properties, "pStyle") ?? {}, "val");
    if (target.inherits || styleNumbers(styles, styleId)) {
      // Numbering the style supplies is switched off by numId 0 (17.9.18).
      const numPr = createWElement(properties, "numPr");
      const numId = createWElement(numPr, "numId", { val: "0" });
      numPr.elements = [numId];
      insertChildInOrder(properties, numPr, PARAGRAPH_PROPERTIES);
    }
    if (level !== undefined) {
      keepLevelIndentation(properties, level);
      keepLevelTabs(properties, level);
    }
    const suffix = wAttribute(childW(level, "suff") ?? {}, "val") ?? "tab";
    const run = labelRun(paragraph, target.label, runProperties, suffix);
    const at = (paragraph.elements ?? []).findIndex((child) =>
      isW(child, "pPr"),
    );
    paragraph.elements?.splice(at + 1, 0, run);
    report.paragraphs.push({
      blockId: target.blockId,
      numId: target.numId,
      level: target.level,
      label: target.label,
    });
  }
  if (report.paragraphs.length > 0) pkg.touch(document);
  return Result.ok(report);
};

/** Remove a leading label the expression matches from a paragraph's runs; what it removed. */
const stripLeadingLabel = (
  paragraph: XmlElement,
  expression: RegExp,
): string | undefined => {
  const pieces: { run: XmlElement; element: XmlElement; text: string }[] = [];
  for (const run of childElements(paragraph)) {
    if (isW(run, "pPr")) continue;
    if (!isW(run, "r")) break;
    for (const element of childElements(run)) {
      if (isW(element, "t")) {
        const text = (element.elements ?? [])
          .map((node) => String(node.text ?? ""))
          .join("");
        pieces.push({ run, element, text });
      } else if (isW(element, "tab")) {
        pieces.push({ run, element, text: "\t" });
      }
    }
  }
  const match = expression.exec(pieces.map(({ text }) => text).join(""));
  if (match?.index !== 0 || match[0] === "") return undefined;
  let remaining = match[0].length;
  for (const { run, element, text } of pieces) {
    if (remaining === 0) break;
    const take = Math.min(remaining, text.length);
    remaining -= take;
    if (take === text.length) {
      run.elements = (run.elements ?? []).filter((child) => child !== element);
      continue;
    }
    element.elements = [{ type: "text", text: text.slice(take) }];
    element.attributes = { ...element.attributes, "xml:space": "preserve" };
  }
  return match[0];
};

const applyNumbering = (
  pkg: OoxmlPackage,
  reviewer: FolioDocxReviewer,
  args: Readonly<Record<string, unknown>>,
): Result<NumberingReport, FolioCliError> => {
  const { numId, level = 0 } = args;
  if (!Number.isInteger(numId) || typeof numId !== "number" || numId < 1) {
    return Result.err(
      invalidInput(
        "apply needs numId, a numbering instance the document defines.",
      ),
    );
  }
  if (
    !Number.isInteger(level) ||
    typeof level !== "number" ||
    level < 0 ||
    level > 8
  ) {
    return Result.err(invalidInput("level must be an integer from 0 to 8."));
  }
  const blockIds = stringList(args["blockIds"], "blockIds");
  if (blockIds.isErr()) return Result.err(blockIds.error);
  if (blockIds.value === null || blockIds.value.length === 0) {
    return Result.err(
      invalidInput("apply needs blockIds, the paragraphs to number."),
    );
  }
  const strip = args["strip"];
  if (strip !== undefined && typeof strip !== "string")
    return Result.err(invalidInput("strip must be an expression."));
  const expression =
    strip === undefined
      ? null
      : Result.try(() => new RegExp(`^(?:${strip})`, "u"));
  if (expression?.isErr())
    return Result.err(
      invalidInput(`strip ${JSON.stringify(strip)} is not a valid expression.`),
    );
  const parts = partsOf(pkg);
  if (parts.isErr()) return Result.err(parts.error);
  const { document, documentRoot, numbering } = parts.value;
  if (levelDefinition(numbering, numId, level) === undefined) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.notFound,
        message: `The document defines no level ${String(level)} of numbering instance ${String(numId)}.`,
      }),
    );
  }
  const known = new Set(reviewer.snapshot().blocks.map(({ id }) => id));
  const paragraphs = paragraphsById(documentRoot);
  const missing = blockIds.value.filter(
    (id) => !known.has(id) || !paragraphs.has(id.toUpperCase()),
  );
  if (missing.length > 0) {
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.notFound,
        message: `No paragraph with id ${missing.join(", ")}.`,
      }),
    );
  }
  const report: NumberingReport = { action: "apply", paragraphs: [] };
  for (const blockId of blockIds.value) {
    const paragraph = paragraphs.get(blockId.toUpperCase());
    if (paragraph === undefined) continue;
    const properties = ensureWChild(paragraph, "pPr", PARAGRAPH_CHILDREN);
    properties.elements = (properties.elements ?? []).filter(
      (child) => !isW(child, "numPr"),
    );
    const numPr = createWElement(properties, "numPr");
    numPr.elements = [];
    insertChildInOrder(
      numPr,
      createWElement(numPr, "ilvl", { val: String(level) }),
      NUMBERING_PROPERTIES,
    );
    insertChildInOrder(
      numPr,
      createWElement(numPr, "numId", { val: String(numId) }),
      NUMBERING_PROPERTIES,
    );
    insertChildInOrder(properties, numPr, PARAGRAPH_PROPERTIES);
    const stripped =
      expression?.isOk() === true
        ? stripLeadingLabel(paragraph, expression.value)
        : undefined;
    report.paragraphs.push({
      blockId,
      numId,
      level,
      ...(stripped !== undefined && { label: stripped }),
    });
  }
  pkg.touch(document);
  return Result.ok(report);
};

/** Run the action `args.action` names (default `literal`). */
export const rewriteNumbering = (
  pkg: OoxmlPackage,
  args: Readonly<Record<string, unknown>>,
  reviewer: FolioDocxReviewer,
): Result<NumberingReport, FolioCliError> => {
  const action = args["action"] ?? "literal";
  if (action !== "literal" && action !== "apply") {
    return Result.err(
      invalidInput(`action must be ${NUMBERING_ACTIONS.join(" or ")}.`),
    );
  }
  return action === "literal"
    ? literalNumbering(pkg, reviewer, args)
    : applyNumbering(pkg, reviewer, args);
};
