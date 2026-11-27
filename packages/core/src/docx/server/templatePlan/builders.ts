/**
 * Raw OOXML builders a template plan is applied with: a hidden bookmark around
 * a paragraph's content, a complex REF field with a cached result, the
 * update-on-open setting, and freezing a field to its cached result.
 */
import { SETTINGS_CHILDREN } from "@stll/docx-core/schema";

import { sliceText } from "../../../utils/sliceText";
import { updateSettingsOnOffChild } from "../../settingsHeaderFooterUpdate";
import type { XmlElement } from "../../xmlParser";
import { sha256Hex } from "./package";
import type { TemplatePlanSectionPage } from "./types";
import {
  elementChildren,
  escapeXmlAttribute,
  escapeXmlText,
  innerRangeOf,
  isW,
  type PartPrefixes,
  partPrefixes,
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
/**
 * What else the general format switch may name: that an update keeps the
 * look of the previous result (ECMA-376 Part 1, 17.16.4.3), so a field
 * set in a face of its own stays so once its result is recalculated.
 */
const LOOK_FORMAT = /^(?:mergeformat|charformat)$/iu;

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
 * single-letter switches, one letter-case format and one format that
 * keeps the result's look.
 */
export const refFieldInstruction = (
  bookmark: string,
  switches: string,
): string | null => {
  const parts = switches.trim().split(/\s+/u).filter(Boolean);
  let cases = 0;
  let looks = 0;
  for (const [index, part] of parts.entries()) {
    const afterFormat = parts[index - 1] === FORMAT_SWITCH;
    if (part === FORMAT_SWITCH) {
      const format = parts[index + 1] ?? "";
      cases += CASE_FORMAT.test(format) ? 1 : 0;
      looks += LOOK_FORMAT.test(format) ? 1 : 0;
      if (
        cases > 1 ||
        looks > 1 ||
        !(CASE_FORMAT.test(format) || LOOK_FORMAT.test(format))
      ) {
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

/**
 * `settingsXml` without `w:trackRevisions`: a document that carries tracked
 * changes of its source records no new ones until its reader switches
 * tracking on. Null when the part is not a settings part.
 */
export const setTrackingOff = (settingsXml: string): string | null =>
  updateSettingsOnOffChild(settingsXml, "trackRevisions", undefined);

const SHIFT_RETURN = "doNotExpandShiftReturn";

/** The children of `w:compat` the schema orders in front of the setting. */
const COMPAT_BEFORE_SHIFT_RETURN: ReadonlySet<string> = new Set([
  "useSingleBorderforContiguousCells",
  "wpJustification",
  "noTabHangInd",
  "noLeading",
  "spaceForUL",
  "noColumnBalance",
  "balanceSingleByteDoubleByteWidth",
  "noExtraLineSpacing",
  "doNotLeaveBackslashAlone",
  "ulTrailSpace",
]);

const SWITCHED_OFF: ReadonlySet<string> = new Set(["0", "false", "off"]);

const SETTINGS_CHILD_ORDINALS: ReadonlyMap<string, number> = new Map(
  SETTINGS_CHILDREN.map((name, index) => [name, index] as const),
);

const settingsRoot = (settingsXml: string): RawPart | null => {
  const part = readRawPart("word/settings.xml", settingsXml);
  return part !== null && isW(part.root, "settings") ? part : null;
};

/** Whether a settings part keeps the spacing of a justified line that ends
 *  in a manual line break (`w:compat/w:doNotExpandShiftReturn`). */
export const keepsShiftReturnSpacing = (settingsXml: string): boolean => {
  const setting = wChild(
    wChild(settingsRoot(settingsXml)?.root, "compat"),
    SHIFT_RETURN,
  );
  return setting !== null && !SWITCHED_OFF.has(wAttr(setting, "val") ?? "1");
};

/** An element of the part's WordprocessingML namespace, bound by the
 *  root's prefix or, under a default namespace, by one of its own. */
const settingsElement = (
  root: XmlElement,
  localName: string,
  inner: string | null,
): string => {
  const prefixes = partPrefixes(root);
  const name =
    prefixes === null ? `folioW:${localName}` : `${prefixes.w}:${localName}`;
  const binding =
    prefixes === null
      ? ` xmlns:folioW="${escapeXmlAttribute(root.namespaceUri ?? "")}"`
      : "";
  return inner === null
    ? `<${name}${binding}/>`
    : `<${name}${binding}>${inner}</${name}>`;
};

/**
 * `settingsXml` with `w:doNotExpandShiftReturn` switched on in `w:compat`,
 * each at its place in the schema's child order; every other byte of the
 * part stays. Null when the part is not a settings part or gives the setting
 * no certain place.
 */
export const setKeepShiftReturnSpacing = (
  settingsXml: string,
): string | null => {
  const part = settingsRoot(settingsXml);
  if (part === null) {
    return null;
  }
  if (keepsShiftReturnSpacing(settingsXml)) {
    return settingsXml;
  }
  const { root } = part;
  const setting = settingsElement(root, SHIFT_RETURN, null);
  const compat = wChild(root, "compat");
  if (compat === null) {
    const compatOrder = SETTINGS_CHILD_ORDINALS.get("compat") ?? 0;
    const wrapper = settingsElement(root, "compat", setting);
    const inner = innerRangeOf(settingsXml, root);
    if (inner === null) {
      // A settings part without a single child is rewritten around it.
      const range = rangeOf(root);
      return range === null || root.name === undefined
        ? null
        : spliceXml(settingsXml, [
            {
              start: range.end - 2,
              end: range.end,
              newXml: `>${wrapper}</${root.name}>`,
            },
          ]);
    }
    const following = elementChildren(root).find((child) => {
      const name = wName(child);
      // A child of another namespace (math properties, later extensions)
      // follows every setting of this one.
      return (
        name === null || (SETTINGS_CHILD_ORDINALS.get(name) ?? -1) > compatOrder
      );
    });
    const at = following === undefined ? inner.end : rangeOf(following)?.start;
    return at === undefined
      ? null
      : spliceXml(settingsXml, [{ start: at, end: at, newXml: wrapper }]);
  }
  const range = rangeOf(compat);
  const inner = innerRangeOf(settingsXml, compat);
  if (range === null) {
    return null;
  }
  if (inner === null) {
    return spliceXml(settingsXml, [
      {
        start: range.start,
        end: range.end,
        newXml: settingsElement(root, "compat", setting),
      },
    ]);
  }
  // One that is there and switched off is replaced where it stands.
  const stated = wChild(compat, SHIFT_RETURN);
  const statedRange = stated === null ? null : rangeOf(stated);
  if (statedRange !== null) {
    return spliceXml(settingsXml, [{ ...statedRange, newXml: setting }]);
  }
  const following = elementChildren(compat).find((child) => {
    const name = wName(child);
    return name === null || !COMPAT_BEFORE_SHIFT_RETURN.has(name);
  });
  const at = following === undefined ? inner.end : rangeOf(following)?.start;
  return at === undefined
    ? null
    : spliceXml(settingsXml, [{ start: at, end: at, newXml: setting }]);
};

/** The children of `w:sectPr` the schema orders in front of `w:type`. */
const SECTION_BEFORE_TYPE: ReadonlySet<string> = new Set([
  "headerReference",
  "footerReference",
  "footnotePr",
  "endnotePr",
]);

const PAGE_NUMBER_KEPT = ["fmt", "chapStyle", "chapSep"] as const;
const MARGINS_KEPT = ["header", "footer", "gutter"] as const;

export type SectionRewrite = {
  /**
   * Whether the section begins the document. Every later one loses the
   * first-page header and a page-number start: they belong to the
   * document's first page, not to each part of it.
   */
  first: boolean;
  /** The page to put in place of the template's, or undefined. */
  page: TemplatePlanSectionPage | undefined;
  /** How the section begins; null keeps what the template states. */
  start: string | null;
  /**
   * Header and footer references the section takes over from the sections
   * in front of it and does not state itself, as markup. Left out, the
   * section is written with its own alone.
   */
  inherited?: readonly string[];
};

/**
 * The header and footer references a body's last section takes over from
 * the sections in front of it: a section that states no header of a kind
 * shows the one of the section before. An output built from the last
 * section alone would lose them, so they are stated on it. Each is the
 * reference of the nearest earlier section that states that kind, as it is
 * written in `xml`; kinds the last section states itself are left out.
 */
export const inheritedReferencesXml = (
  xml: string,
  body: XmlElement,
  last: XmlElement,
): string[] => {
  const keyOf = (reference: XmlElement): string | null => {
    const name = wName(reference);
    return name === "headerReference" || name === "footerReference"
      ? `${name}:${wAttr(reference, "type") ?? "default"}`
      : null;
  };
  const referencesOf = (sectPr: XmlElement | null): [string, XmlElement][] =>
    elementChildren(sectPr).flatMap((child) => {
      const key = keyOf(child);
      return key === null ? [] : [[key, child] as [string, XmlElement]];
    });
  const inEffect = new Map<string, XmlElement>();
  for (const block of elementChildren(body)) {
    const sectPr =
      wName(block) === "p" ? wChild(wChild(block, "pPr"), "sectPr") : null;
    for (const [key, reference] of referencesOf(sectPr)) {
      inEffect.set(key, reference);
    }
  }
  const own = new Set(referencesOf(last).map(([key]) => key));
  return [...inEffect].flatMap(([key, reference]) => {
    const range = own.has(key) ? null : rangeOf(reference);
    return range === null ? [] : [xml.slice(range.start, range.end)];
  });
};

/**
 * A `w:sectPr` built from the template's: the same children in the same
 * order, with the page, the start and the first-section properties changed
 * as `rewrite` says. `xml` is the part `sectPr` was read from. Null when a
 * child cannot be read back.
 */
export const sectionPropertiesXml = (
  xml: string,
  sectPr: XmlElement,
  w: string,
  { first, page, start, inherited = [] }: SectionRewrite,
): string | null => {
  const attributes = (element: XmlElement, names: readonly string[]): string =>
    names
      .map((name) => [name, wAttr(element, name)] as const)
      .filter(([, value]) => value !== null)
      .map(
        ([name, value]) => ` ${w}:${name}="${escapeXmlAttribute(value ?? "")}"`,
      )
      .join("");
  const type = start === null ? "" : `<${w}:type ${w}:val="${start}"/>`;
  // References stand first in a section's properties.
  const out: string[] = [...inherited];
  let typeWritten = type === "";
  for (const child of elementChildren(sectPr)) {
    const name = wName(child);
    if (!typeWritten && (name === null || !SECTION_BEFORE_TYPE.has(name))) {
      out.push(type);
      typeWritten = true;
    }
    if (
      name === "sectPrChange" ||
      (name === "type" && start !== null) ||
      (name === "titlePg" && !first)
    ) {
      continue;
    }
    if (name === "pgSz" && page !== undefined) {
      out.push(
        `<${w}:pgSz ${w}:w="${page.width}" ${w}:h="${page.height}"${page.orientation === "landscape" ? ` ${w}:orient="landscape"` : ""}/>`,
      );
      continue;
    }
    if (name === "pgMar" && page?.margins !== undefined) {
      const { top, right, bottom, left } = page.margins;
      out.push(
        `<${w}:pgMar ${w}:top="${top}" ${w}:right="${right}" ${w}:bottom="${bottom}" ${w}:left="${left}"${attributes(child, MARGINS_KEPT)}/>`,
      );
      continue;
    }
    if (name === "pgNumType" && !first) {
      const kept = attributes(child, PAGE_NUMBER_KEPT);
      if (kept !== "") {
        out.push(`<${w}:pgNumType${kept}/>`);
      }
      continue;
    }
    const range = rangeOf(child);
    if (range === null) {
      return null;
    }
    out.push(xml.slice(range.start, range.end));
  }
  if (!typeWritten) {
    out.push(type);
  }
  return `<${w}:sectPr>${out.join("")}</${w}:sectPr>`;
};

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
