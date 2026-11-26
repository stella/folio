/**
 * Raw OOXML builders a template plan is applied with: a hidden bookmark around
 * a paragraph's content, a complex REF field with a cached result, the
 * update-on-open setting, and freezing a field to its cached result.
 */
import { sliceText } from "../../../utils/sliceText";
import { updateSettingsOnOffChild } from "../../settingsHeaderFooterUpdate";
import type { XmlElement } from "../../xmlParser";
import { sha256Hex } from "./package";
import {
  elementChildren,
  escapeXmlAttribute,
  escapeXmlText,
  innerRangeOf,
  isW,
  type PartPrefixes,
  rangeOf,
  type RawPart,
  readRawPart,
  spliceXml,
  walkElements,
  wAttr,
  wChild,
  wName,
  type XmlSplice,
} from "./xml";

const RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REF_BOOKMARK_PREFIX = "_Ref";
const REF_BOOKMARK_DIGITS = 9;
const REF_BOOKMARK_SPACE = 10 ** REF_BOOKMARK_DIGITS;

/**
 * A hidden bookmark name in the form a cross-reference target takes: `_Ref`
 * and nine digits, derived from `seed` and stepped past the names in `taken`.
 * The name is added to `taken`.
 */
export const mintRefBookmarkName = (
  seed: string,
  taken: Set<string>,
): string => {
  let value =
    Number.parseInt(sha256Hex(seed).slice(0, 12), 16) % REF_BOOKMARK_SPACE;
  for (;;) {
    const name = `${REF_BOOKMARK_PREFIX}${String(value).padStart(REF_BOOKMARK_DIGITS, "0")}`;
    if (!taken.has(name)) {
      taken.add(name);
      return name;
    }
    value = (value + 1) % REF_BOOKMARK_SPACE;
  }
};

/** The `w:id` values `w:bookmarkStart` elements under `root` already use. */
export const bookmarkIdsOf = (root: XmlElement): number[] => {
  const ids: number[] = [];
  walkElements(root, (element) => {
    if (!isW(element, "bookmarkStart") && !isW(element, "bookmarkEnd")) {
      return;
    }
    const id = Number.parseInt(wAttr(element, "id") ?? "", 10);
    if (Number.isInteger(id)) {
      ids.push(id);
    }
  });
  return ids;
};

export type BookmarkWrap = { id: number; name: string };

type BookmarkMarkupOptions = BookmarkWrap & { prefix: string };

export const bookmarkStartXml = ({
  prefix,
  id,
  name,
}: BookmarkMarkupOptions): string =>
  `<${prefix}:bookmarkStart ${prefix}:id="${id}" ${prefix}:name="${escapeXmlAttribute(name)}"/>`;

export const bookmarkEndXml = ({
  prefix,
  id,
}: Pick<BookmarkMarkupOptions, "prefix" | "id">): string =>
  `<${prefix}:bookmarkEnd ${prefix}:id="${id}"/>`;

type WrapParagraphOptions = BookmarkWrap & {
  /** One serialized `w:p`. */
  paragraphXml: string;
  prefixes: PartPrefixes;
};

/**
 * `paragraphXml` with its content, everything after `w:pPr`, inside a
 * bookmark. Null when the markup is not one paragraph.
 */
export const wrapParagraphContentInBookmark = ({
  paragraphXml,
  prefixes,
  id,
  name,
}: WrapParagraphOptions): string | null => {
  // The fragment is read on its own, so the prefixes its part's root binds
  // are declared around it.
  const declared = `<wrap xmlns:${prefixes.w}="${prefixes.wUri}" xmlns:${prefixes.r}="${RELATIONSHIPS_NAMESPACE}">${paragraphXml}</wrap>`;
  const offset = declared.indexOf(">") + 1;
  const part = readRawPart("paragraph", declared);
  const paragraph =
    part === null ? undefined : elementChildren(part.root).at(0);
  if (
    part === null ||
    paragraph === undefined ||
    !isW(paragraph, "p") ||
    elementChildren(part.root).length !== 1
  ) {
    return null;
  }
  const start = bookmarkStartXml({ prefix: prefixes.w, id, name });
  const end = bookmarkEndXml({ prefix: prefixes.w, id });
  const inner = innerRangeOf(declared, paragraph);
  if (inner === null) {
    // `<w:p/>`: an empty paragraph still anchors the bookmark.
    const name_ = paragraph.name ?? `${prefixes.w}:p`;
    return `${paragraphXml.slice(0, paragraphXml.lastIndexOf("/>"))}>${start}${end}</${name_}>`;
  }
  const properties = wChild(paragraph, "pPr");
  const propertiesEnd =
    properties === null ? null : (rangeOf(properties)?.end ?? null);
  const startAt = (propertiesEnd ?? inner.start) - offset;
  const endAt = inner.end - offset;
  // Two insertions at one offset would come out in either order.
  return spliceXml(
    paragraphXml,
    startAt === endAt
      ? [{ start: startAt, end: startAt, newXml: `${start}${end}` }]
      : [
          { start: startAt, end: startAt, newXml: start },
          { start: endAt, end: endAt, newXml: end },
        ],
  );
};

type ComplexFieldOptions = {
  prefix: string;
  /** The field code between the field characters, e.g. ` REF _Ref000000001 \n \h `. */
  instruction: string;
  /** Serialized runs holding the cached result. */
  resultRunsXml: string;
  /** A serialized `w:rPr` the field's own runs carry, or the empty string. */
  runPropertiesXml: string;
};

/** A complex field: begin, instruction, separate, the cached result, end. */
export const complexFieldXml = ({
  prefix,
  instruction,
  resultRunsXml,
  runPropertiesXml,
}: ComplexFieldOptions): string => {
  const char = (type: "begin" | "separate" | "end"): string =>
    `<${prefix}:r>${runPropertiesXml}<${prefix}:fldChar ${prefix}:fldCharType="${type}"/></${prefix}:r>`;
  return `${char("begin")}<${prefix}:r>${runPropertiesXml}<${prefix}:instrText xml:space="preserve">${escapeXmlText(instruction)}</${prefix}:instrText></${prefix}:r>${char("separate")}${resultRunsXml}${char("end")}`;
};

/** The REF switches that take no argument. */
const REF_SWITCH = /^\\[fhnprtw]$/u;
/** The general format switch and the letter cases it may name. */
const FORMAT_SWITCH = "\\*";
const CASE_FORMAT = /^(?:lower|upper|firstcap)$/iu;

/** The letter case a REF instruction asks for its result in, or null. */
export const refCaseFormat = (
  instruction: string,
): "lower" | "upper" | "firstcap" | null => {
  const name = /\\\*\s*(?<name>lower|upper|firstcap)\b/iu
    .exec(instruction)
    ?.groups?.["name"]?.toLowerCase();
  return name === "lower" || name === "upper" || name === "firstcap"
    ? name
    : null;
};

/** `text` in the letter case a `\*` switch names, as a field result has it. */
export const applyRefCaseFormat = (
  text: string,
  format: ReturnType<typeof refCaseFormat>,
): string => {
  switch (format) {
    case "lower":
      // Every letter, a roman numeral's included.
      return text.toLowerCase();
    case "upper":
      return text.toUpperCase();
    case "firstcap": {
      const initial = sliceText({ text, units: "graphemes", end: 1 });
      return initial.toUpperCase() + text.slice(initial.length);
    }
    case null:
      return text;
    default:
      return format satisfies never;
  }
};

/**
 * The field code of a REF to `bookmark` with `switches` (`\n \t \h`,
 * `\n \* lower \h`). Null when the switches hold anything but
 * single-letter switches and one letter-case format.
 */
export const refFieldInstruction = (
  bookmark: string,
  switches: string,
): string | null => {
  const parts = switches.trim().split(/\s+/u).filter(Boolean);
  let formats = 0;
  for (const [index, part] of parts.entries()) {
    const afterFormat = parts[index - 1] === FORMAT_SWITCH;
    if (part === FORMAT_SWITCH) {
      formats += 1;
      if (formats > 1 || !CASE_FORMAT.test(parts[index + 1] ?? "")) {
        return null;
      }
    } else if (!afterFormat && !REF_SWITCH.test(part)) {
      return null;
    }
  }
  return ` REF ${bookmark} ${parts.join(" ")}${parts.length === 0 ? "" : " "}`;
};

/** The bookmark a REF, PAGEREF or NOTEREF instruction names. */
export const fieldBookmarkOf = (instruction: string): string | null =>
  /^\s*(?:REF|PAGEREF|NOTEREF)\s+"?(?<name>[^\s"\\]+)/iu.exec(instruction)
    ?.groups?.["name"] ?? null;

/**
 * `settingsXml` with `w:updateFields` set (`true`) or absent (`false`), at its
 * place in the schema's child order. Null when the part is not a settings part.
 */
export const setUpdateFieldsOnOpen = (
  settingsXml: string,
  enabled: boolean,
): string | null =>
  updateSettingsOnOffChild(
    settingsXml,
    "updateFields",
    enabled ? true : undefined,
  );

export type FrozenFieldCount = { kind: string; count: number };

export type FreezeFieldsResult = {
  /** Splices that remove the field machinery and keep the cached result. */
  splices: XmlSplice[];
  frozen: FrozenFieldCount[];
};

const fieldKindOfInstruction = (instruction: string): string =>
  /^\s*\\?(?<name>[A-Za-z][A-Za-z0-9]*)/u
    .exec(instruction)
    ?.groups?.["name"]?.toUpperCase() ?? "";

type OpenFrozenField = {
  /** Markup removed when this field is frozen: its begin and all of its code. */
  pending: XmlElement[];
  instruction: string;
  phase: "code" | "result";
  /** False for a field inside another field's code, which is left as it is. */
  candidate: boolean;
  frozen: boolean;
};

const textOf = (element: XmlElement): string =>
  (element.elements ?? [])
    .map((child) => (child.type === "text" ? String(child.text ?? "") : ""))
    .join("");

/** Drop splices that lie inside an earlier one, so none overlap. */
export const withoutContained = (
  splices: readonly XmlSplice[],
): XmlSplice[] => {
  const kept: XmlSplice[] = [];
  let coveredUntil = -1;
  for (const splice of splices.toSorted(
    (left, right) => left.start - right.start || right.end - left.end,
  )) {
    if (splice.start < coveredUntil) {
      continue;
    }
    kept.push(splice);
    coveredUntil = Math.max(coveredUntil, splice.end);
  }
  return kept;
};

/**
 * Replace every field of the given kinds under `part` by its cached result:
 * a simple field keeps its runs, a complex field loses its field characters
 * and its code. A field nested in a kept result is judged on its own; one
 * nested in another field's code is part of that code.
 */
export const freezeFields = (
  part: RawPart,
  kinds: ReadonlySet<string>,
): FreezeFieldsResult => {
  const splices: XmlSplice[] = [];
  const counts = new Map<string, number>();
  const count = (instruction: string): void => {
    const kind = fieldKindOfInstruction(instruction);
    counts.set(kind, (counts.get(kind) ?? 0) + 1);
  };
  const remove = (element: XmlElement): void => {
    const range = rangeOf(element);
    if (range !== null) {
      splices.push({ ...range, newXml: "" });
    }
  };
  const open: OpenFrozenField[] = [];
  const wanted = (instruction: string): boolean =>
    kinds.has(fieldKindOfInstruction(instruction));
  /** The outermost open field whose code the walk is inside. */
  const codeOwner = (
    fields: readonly OpenFrozenField[] = open,
  ): OpenFrozenField | undefined =>
    fields.find(({ phase }) => phase === "code");
  const freeze = (field: OpenFrozenField, closing: XmlElement): void => {
    field.frozen = true;
    count(field.instruction);
    for (const pending of field.pending) {
      remove(pending);
    }
    remove(closing);
  };

  const visitRunChild = (child: XmlElement): void => {
    if (wName(child) !== "fldChar") {
      const top = open.at(-1);
      if (wName(child) === "instrText" && top?.phase === "code") {
        top.instruction += textOf(child);
      }
      const owner = codeOwner();
      if (owner === undefined) {
        visit(child);
      } else {
        owner.pending.push(child);
      }
      return;
    }
    const type = wAttr(child, "fldCharType");
    if (type === "begin") {
      const owner = codeOwner();
      const field: OpenFrozenField = {
        pending: [],
        instruction: "",
        phase: "code",
        candidate: owner === undefined,
        frozen: false,
      };
      (owner ?? field).pending.push(child);
      open.push(field);
      return;
    }
    const top = open.at(-1);
    if (top === undefined) {
      return;
    }
    if (type === "separate") {
      if (top.phase === "result") {
        // A second separator of one field: malformed, and already decided.
        return;
      }
      const owner = codeOwner(open.slice(0, -1));
      top.phase = "result";
      if (owner !== undefined) {
        owner.pending.push(child);
      } else if (top.candidate && wanted(top.instruction)) {
        freeze(top, child);
      }
      return;
    }
    if (type === "end") {
      open.pop();
      const owner = codeOwner();
      if (owner !== undefined) {
        owner.pending.push(child);
      } else if (top.frozen) {
        remove(child);
      } else if (
        top.candidate &&
        top.phase === "code" &&
        wanted(top.instruction)
      ) {
        // A field without a result freezes to nothing.
        freeze(top, child);
      }
    }
  };

  const visit = (element: XmlElement): void => {
    const name = wName(element);
    if (name === "r") {
      for (const child of elementChildren(element)) {
        if (wName(child) !== "rPr") {
          visitRunChild(child);
        }
      }
      return;
    }
    if (name === "fldSimple") {
      const owner = codeOwner();
      if (owner !== undefined) {
        owner.pending.push(element);
        return;
      }
      const instruction = wAttr(element, "instr") ?? "";
      const range = rangeOf(element);
      // A prefix the field declares itself is in scope for its runs only
      // while the field stands; such a field stays live rather than leave
      // its result with a prefix nothing binds.
      const declaresPrefix = Object.keys(element.attributes ?? {}).some(
        (attribute) => attribute === "xmlns" || attribute.startsWith("xmlns:"),
      );
      if (wanted(instruction) && range !== null && !declaresPrefix) {
        count(instruction);
        const inner = innerRangeOf(part.xml, element);
        if (inner === null) {
          splices.push({ ...range, newXml: "" });
          return;
        }
        splices.push(
          { start: range.start, end: inner.start, newXml: "" },
          { start: inner.end, end: range.end, newXml: "" },
        );
      }
    }
    for (const child of elementChildren(element)) {
      visit(child);
    }
  };
  visit(part.root);

  return {
    splices: withoutContained(splices),
    frozen: [...counts].map(([kind, total]) => ({ kind, count: total })),
  };
};
