/**
 * The shape of a template plan as it arrives: parsed JSON nobody has typed
 * yet. Everything the resolver later reads off a plan is checked here, so a
 * malformed plan is a `plan-invalid` refusal and never an exception.
 */
import {
  HIGHLIGHT_COLOR_VALUES,
  NUMBER_FORMAT_VALUES,
} from "../../../types/documentEnumValues";

export const OMIT_REASONS: ReadonlySet<string> = new Set([
  "empty",
  "sourceToc",
  "pageArtifact",
  "duplicateLabel",
]);
const MERGES: ReadonlySet<string> = new Set(["lineBreak", "space", "none"]);
const SUFFIXES: ReadonlySet<string> = new Set(["tab", "space", "nothing"]);
const NUMBER_FORMATS: ReadonlySet<string> = new Set(NUMBER_FORMAT_VALUES);
/** The highlight a plan may add; clearing one is not a mark. */
const CLEARED_HIGHLIGHT = "none";
export const PLAN_HIGHLIGHT_COLORS: ReadonlySet<string> = new Set(
  HIGHLIGHT_COLOR_VALUES.filter((color) => color !== CLEARED_HIGHLIGHT),
);
const FIELD_KIND = /^[A-Z][A-Z0-9]*$/u;
const MAX_LIST_LEVEL = 8;

type Json = Record<string, unknown>;

const isObject = (value: unknown): value is Json =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isString = (value: unknown): value is string => typeof value === "string";

const isOptionalBoolean = (value: unknown): boolean =>
  value === undefined || typeof value === "boolean";

const isOptionalString = (value: unknown): boolean =>
  value === undefined || isString(value);

const isCount = (value: unknown): value is number =>
  typeof value === "number" && Number.isSafeInteger(value) && value >= 0;

/** A list id a plan may name: the first real one is 1. */
const isListId = (value: unknown): boolean => isCount(value) && value >= 1;

const isLevel = (value: unknown): boolean =>
  isCount(value) && value <= MAX_LIST_LEVEL;

/**
 * Whether `text` is plain text a `w:t` holds as written: no control
 * characters, no tab or line end, and no half of a surrogate pair.
 */
export const isPlainText = (text: string): boolean => {
  for (let index = 0; index < text.length; index += 1) {
    const unit = text.charCodeAt(index);
    if (unit < 0x20 || unit === 0x7f || unit === 0xfffe || unit === 0xffff) {
      return false;
    }
    const high = unit >= 0xd800 && unit <= 0xdbff;
    const low = unit >= 0xdc00 && unit <= 0xdfff;
    if (high) {
      const next = text.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) {
        return false;
      }
      index += 1;
    } else if (low) {
      return false;
    }
  }
  return true;
};

const numberingIssue = (value: unknown, where: string): string | null => {
  if (value === undefined || value === null) {
    return null;
  }
  if (!isObject(value)) {
    return `${where} names a malformed list.`;
  }
  const { numId, ilvl, restartFrom, baseNumId, group } = value;
  if (numId !== "new" && !isListId(numId)) {
    return `${where} names a list that is neither an id nor "new".`;
  }
  if (!isLevel(ilvl)) {
    return `${where} names a list level outside 0 to ${MAX_LIST_LEVEL}.`;
  }
  if (restartFrom !== undefined && !isCount(restartFrom)) {
    return `${where} restarts a list from a value that is not a count.`;
  }
  if (baseNumId !== undefined && !isListId(baseNumId)) {
    return `${where} bases a list on a malformed id.`;
  }
  if (group !== undefined && !(isString(group) && group !== "")) {
    return `${where} names a malformed list group.`;
  }
  return null;
};

const SECTION_STARTS = new Set(["nextPage", "evenPage", "oddPage"]);
const ORIENTATIONS = new Set(["portrait", "landscape"]);
const isTwips = (value: unknown): boolean =>
  typeof value === "number" && Number.isSafeInteger(value) && value >= 0;

/** Why `section` is not a section of a plan, or null when it is one. */
const sectionIssue = (section: unknown, where: string): string | null => {
  if (section === undefined) {
    return null;
  }
  if (!isObject(section)) {
    return `${where} describes a malformed section.`;
  }
  const { page, start } = section;
  if (start !== undefined && !(isString(start) && SECTION_STARTS.has(start))) {
    return `${where} names an unknown section start.`;
  }
  if (page === undefined) {
    return null;
  }
  if (
    !isObject(page) ||
    !(isString(page["orientation"]) && ORIENTATIONS.has(page["orientation"])) ||
    !isTwips(page["width"]) ||
    !isTwips(page["height"]) ||
    page["width"] === 0 ||
    page["height"] === 0
  ) {
    return `${where} describes a malformed page.`;
  }
  const margins = page["margins"];
  return margins === undefined ||
    (isObject(margins) &&
      ["top", "right", "bottom", "left"].every((side) =>
        isTwips(margins[side]),
      ))
    ? null
    : `${where} describes malformed page margins.`;
};

const INDENT_KEYS: ReadonlySet<string> = new Set([
  "left",
  "hanging",
  "firstLine",
]);
/** The widest indent a plan may state: the widest page WordprocessingML allows. */
const MAX_INDENT = 31_680;

/**
 * Whether `value` is an indent a plan may state: at least one of
 * `left`, `hanging` and `firstLine`, each a whole number of twentieths of a
 * point, and not both a hanging and a first-line indent.
 */
const isIndent = (value: unknown): boolean => {
  if (!isObject(value)) {
    return false;
  }
  const keys = Object.keys(value);
  return (
    keys.length > 0 &&
    keys.every(
      (key) =>
        INDENT_KEYS.has(key) && isCount(value[key]) && value[key] <= MAX_INDENT,
    ) &&
    !(value["hanging"] !== undefined && value["firstLine"] !== undefined)
  );
};

/** The most tab stops a plan may give one paragraph. */
const MAX_TAB_STOPS = 8;

/**
 * Whether `value` is a list of tab stops a plan may state: one to eight
 * whole numbers of twentieths of a point, above zero and ascending.
 */
const isTabStops = (value: unknown): boolean =>
  Array.isArray(value) &&
  value.length > 0 &&
  value.length <= MAX_TAB_STOPS &&
  value.every(
    (stop, index) =>
      isCount(stop) &&
      stop > 0 &&
      stop <= MAX_INDENT &&
      (index === 0 || stop > value[index - 1]),
  );

const sourceItemIssue = (item: Json, where: string): string | null => {
  if (!isString(item["blockId"]) || !isString(item["hash"])) {
    return `${where} names no source block and hash.`;
  }
  for (const key of ["styleId", "cellStyleId", "stripLabel"] as const) {
    if (!isOptionalString(item[key])) {
      return `${where} carries a malformed ${key}.`;
    }
  }
  for (const key of [
    "labelAsText",
    "pageBreakBefore",
    "keepNext",
    "stripLabelBreak",
  ] as const) {
    if (!isOptionalBoolean(item[key])) {
      return `${where} carries a malformed ${key}.`;
    }
  }
  if (item["indent"] !== undefined && !isIndent(item["indent"])) {
    return `${where} carries a malformed indent.`;
  }
  if (item["tabStops"] !== undefined && !isTabStops(item["tabStops"])) {
    return `${where} carries malformed tab stops.`;
  }
  const merge = item["mergeNext"];
  if (merge !== undefined && !(isString(merge) && MERGES.has(merge))) {
    return `${where} names an unknown merge.`;
  }
  const section = sectionIssue(item["endsSection"], where);
  if (section !== null) {
    return section;
  }
  const numbering = numberingIssue(item["numbering"], where);
  if (numbering !== null) {
    return numbering;
  }
  const nested = item["nested"];
  if (nested === undefined) {
    return null;
  }
  if (!Array.isArray(nested)) {
    return `${where} carries malformed nested placements.`;
  }
  const seen = new Set<string>();
  for (const entry of nested) {
    if (!isObject(entry) || !isString(entry["blockId"])) {
      return `${where} carries a nested placement without a block.`;
    }
    if (seen.has(entry["blockId"])) {
      return `${where} places nested block ${entry["blockId"]} twice.`;
    }
    seen.add(entry["blockId"]);
    if (
      !isOptionalString(entry["styleId"]) ||
      !isOptionalString(entry["stripLabel"]) ||
      !isOptionalBoolean(entry["labelAsText"]) ||
      !isOptionalBoolean(entry["keepNext"]) ||
      (entry["indent"] !== undefined && !isIndent(entry["indent"])) ||
      (entry["tabStops"] !== undefined && !isTabStops(entry["tabStops"]))
    ) {
      return `${where} carries a malformed nested placement.`;
    }
    const nestedNumbering = numberingIssue(entry["numbering"], where);
    if (nestedNumbering !== null) {
      return nestedNumbering;
    }
  }
  return null;
};

const bodyIssue = (body: unknown): string | null => {
  if (!Array.isArray(body)) {
    return "The plan has no body.";
  }
  for (const [index, item] of body.entries()) {
    const where = `Body item ${index}`;
    if (!isObject(item)) {
      return `${where} is malformed.`;
    }
    switch (item["kind"]) {
      case "template":
        if (
          !Array.isArray(item["templateBlockIds"]) ||
          !item["templateBlockIds"].every(isString)
        ) {
          return `${where} names no template blocks.`;
        }
        break;
      case "source": {
        const issue = sourceItemIssue(item, where);
        if (issue !== null) {
          return issue;
        }
        break;
      }
      case "slot":
        if (
          !isString(item["templateBlockId"]) ||
          !Array.isArray(item["sourceBlockIds"]) ||
          item["sourceBlockIds"].length === 0 ||
          !item["sourceBlockIds"].every(isString)
        ) {
          return `${where} is a slot without a template paragraph or source blocks.`;
        }
        break;
      default:
        return `${where} has an unknown kind.`;
    }
  }
  return null;
};

const omitIssue = (omit: unknown): string | null => {
  if (!Array.isArray(omit)) {
    return "The plan has no omit list.";
  }
  for (const entry of omit) {
    if (
      !isObject(entry) ||
      !isString(entry["blockId"]) ||
      !isString(entry["reason"]) ||
      !OMIT_REASONS.has(entry["reason"])
    ) {
      return "The plan omits a block without an id or for an unknown reason.";
    }
  }
  return null;
};

const patchesIssue = (patches: unknown): string | null => {
  if (patches === undefined) {
    return null;
  }
  if (!Array.isArray(patches)) {
    return "The plan's numbering patches are malformed.";
  }
  const seen = new Set<string>();
  for (const patch of patches) {
    if (
      !isObject(patch) ||
      !isCount(patch["abstractNumId"]) ||
      !isLevel(patch["ilvl"])
    ) {
      return "A numbering patch names no list definition and level.";
    }
    const key = `${String(patch["abstractNumId"])}:${String(patch["ilvl"])}`;
    if (seen.has(key)) {
      return `Two numbering patches change level ${String(patch["ilvl"])} of list definition ${String(patch["abstractNumId"])}.`;
    }
    seen.add(key);
    const { numFmt, lvlText, start, suff, indent } = patch;
    if (
      (numFmt !== undefined &&
        !(isString(numFmt) && NUMBER_FORMATS.has(numFmt))) ||
      (lvlText !== undefined && !(isString(lvlText) && isPlainText(lvlText))) ||
      (start !== undefined && !isCount(start)) ||
      (suff !== undefined && !(isString(suff) && SUFFIXES.has(suff))) ||
      (indent !== undefined && !isIndent(indent))
    ) {
      return "A numbering patch carries a malformed value.";
    }
  }
  return null;
};

const spansIssue = (plan: Json): string | null => {
  const { references, highlights } = plan;
  if (references !== undefined) {
    if (!Array.isArray(references)) {
      return "The plan's references are malformed.";
    }
    for (const reference of references) {
      if (
        !isObject(reference) ||
        !isString(reference["blockId"]) ||
        !isString(reference["targetBlockId"]) ||
        !isCount(reference["start"]) ||
        !isCount(reference["end"]) ||
        !isString(reference["instruction"])
      ) {
        return "A reference is malformed.";
      }
      const expected = reference["expectedText"];
      if (!isString(expected) || expected === "" || !isPlainText(expected)) {
        return `A reference in ${reference["blockId"]} has no plain cached text.`;
      }
    }
  }
  if (highlights !== undefined) {
    if (!Array.isArray(highlights)) {
      return "The plan's highlights are malformed.";
    }
    for (const highlight of highlights) {
      if (
        !isObject(highlight) ||
        !isString(highlight["blockId"]) ||
        !isCount(highlight["start"]) ||
        !isCount(highlight["end"]) ||
        !isString(highlight["color"]) ||
        !isOptionalString(highlight["note"])
      ) {
        return "A highlight is malformed.";
      }
    }
  }
  return null;
};

const fieldsIssue = (fields: unknown): string | null => {
  if (fields === undefined) {
    return null;
  }
  if (!isObject(fields)) {
    return "The plan's field choices are malformed.";
  }
  const { freezeKinds, updateOnOpen } = fields;
  if (
    freezeKinds !== undefined &&
    !(
      Array.isArray(freezeKinds) &&
      freezeKinds.every((kind) => isString(kind) && FIELD_KIND.test(kind))
    )
  ) {
    return "The plan freezes field kinds that are not field names.";
  }
  if (updateOnOpen !== undefined && typeof updateOnOpen !== "boolean") {
    return "The plan's update-on-open choice is not a boolean.";
  }
  return null;
};

const compatIssue = (compat: unknown): string | null => {
  if (compat === undefined) {
    return null;
  }
  return isObject(compat) && isOptionalBoolean(compat["doNotExpandShiftReturn"])
    ? null
    : "The plan's compatibility settings are malformed.";
};

/** Why `plan` is not a version 1 template plan, or null when it is one. */
export const planShapeIssue = (plan: unknown): string | null => {
  if (
    !isObject(plan) ||
    plan["version"] !== 1 ||
    !isString(plan["sourceSha256"]) ||
    !isString(plan["templateSha256"])
  ) {
    return "The plan is not a version 1 template plan.";
  }
  return (
    bodyIssue(plan["body"]) ??
    omitIssue(plan["omit"]) ??
    patchesIssue(plan["numberingPatches"]) ??
    spansIssue(plan) ??
    fieldsIssue(plan["fields"]) ??
    compatIssue(plan["compat"]) ??
    (plan["trackedChanges"] === undefined ||
    plan["trackedChanges"] === "accept" ||
    plan["trackedChanges"] === "reject"
      ? null
      : "The plan names an unknown view of tracked changes.") ??
    sectionIssue(plan["lastSection"], "The plan's last section") ??
    (plan["textEffects"] === undefined ||
    plan["textEffects"] === "normalize" ||
    plan["textEffects"] === "keep"
      ? null
      : "The plan's choice for text effects is unknown.")
  );
};
