/**
 * The survival law: what folio reads, folio writes back.
 *
 * Four laws run over one synthesised pair. They are reported separately
 * because they fail for different reasons and are fixed in different places.
 *
 * - **L1 parse** — `parseDocx` does not throw on the fixture.
 * - **L2 serialize** — parse, force every serializer to run by removing the
 *   verbatim captures replay would otherwise hand back, save, and find the
 *   subject under the chain the fixture wrote it at, as many times as the
 *   fixture wrote it, with an equal value. For a part a repack copies through
 *   the forcing reaches part level instead, via `PART_REBUILDERS`.
 * - **L3 editor** — the same through `toProseDoc`/`fromProseDoc`, which is the
 *   path every edited document takes, projected with reuse declined
 *   (`./projection.ts`) so a pair that comes back through a reused base block
 *   is not read as one the projection carried. A declaration part is not in
 *   the projection, so on the part leg L3 reports `null` rather than either.
 * - **L4 schema** — the part L2 wrote carries no schema violation the fixture
 *   did not already carry.
 *
 * L2 is the law with teeth. folio replays captured bytes whenever a
 * fingerprint says the model still agrees with them, so a round trip over an
 * untouched document exercises the capture machinery rather than the
 * serializers. Stripping the captures is the same forcing the corpus
 * `reserialize` invariant applies, and it is reused from there rather than
 * restated.
 */

import JSZip from "jszip";

import { parseDocx } from "@stll/folio-core/docx/parser";
import { REVISION_ELEMENT_NAMES } from "@stll/folio-core/docx/revisionIdNormalization";
import { createEmptyDocx, repackDocx } from "@stll/folio-core/docx/rezip";
import { transitionalSlotEncoding } from "@stll/folio-core/docx/transitionalSpelling";
import { universalMeasureAs } from "@stll/folio-core/docx/universalMeasure";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import type { Document } from "@stll/folio-core/types/document";
import { Result } from "better-result";

import { loadSchemaGraph, validateOoxmlPart } from "../corpus-schema-validator";
import { withoutSerializerCaptures } from "../corpus-invariants/reserialize";
import { type CanonicalForms, canonicalFormsOf } from "./canonicalSpellings";
import {
  type BuiltFixture,
  buildFixture,
  IMAGE_RELATIONSHIP_ID,
  modelledCompanionFor,
  spell,
  type Subject,
} from "./fixture";
import { partRebuildFor } from "./partRebuilders";
import { projectWithoutReuse } from "./projection";
import {
  attributeSlotKey,
  childSlotKey,
  containerKey,
  type ContainerSpace,
  qualify,
  type RebuiltPart,
  REBUILT_PARTS,
  WML_NAMESPACE,
} from "./schemaSpace";

export const SURVIVAL_LAWS = {
  parse: "L1-parse",
  serialize: "L2-serialize",
  editor: "L3-editor",
  schema: "L4-schema",
} as const;

export type SurvivalLaw = (typeof SURVIVAL_LAWS)[keyof typeof SURVIVAL_LAWS];

/**
 * Why a pair did not survive.
 *
 * Each value names a different place to fix it, which is the point of
 * separating them: `never-parsed` is a parser change, `replay-rejected` is a
 * replay gate, `lost-in-editor-projection` is the ProseMirror schema.
 */
export const LOSS_MECHANISMS = {
  neverParsed: "never-parsed",
  parsedNotSerialized: "parsed-but-not-serialized",
  replayOnly: "serialized-only-via-verbatim-replay",
  replayRejected: "replay-rejected",
  editorProjection: "lost-in-the-editor-projection",
  respelled: "present-with-a-different-value",
  /**
   * The slot came back, and some of the instances the fixture wrote did not.
   *
   * A reader that keeps the first `wp:lineTo` of a wrap polygon and drops the
   * second writes a shape the source did not describe, and every pair on the
   * slot still survives: the element is there. It is its own mechanism because
   * it is its own fix — a reader or a serializer that handles one instance of a
   * repeated particle and not the rest — and because it can only be observed
   * where the container survived, so it never competes with
   * {@link LOSS_MECHANISMS.containerLost}.
   */
  repeatTruncated: "repeat-truncated",
  /**
   * The pair is lost because its container is, so the defect is the container's.
   *
   * Reported separately because it is not an independent finding: fixing the
   * container fixes every pair under it, and counting them all against the
   * child would put the weight of the census in the wrong place.
   */
  containerLost: "the-container-itself-is-lost",
} as const;

export type LossMechanism = (typeof LOSS_MECHANISMS)[keyof typeof LOSS_MECHANISMS];

export type PairKind = "child" | "attribute";

export type PairOutcome = {
  /** Stable identity: the same string the contract keys its decision by. */
  key: string;
  kind: PairKind;
  container: string;
  /** The child element or attribute under test, qualified. */
  subject: string;
  /**
   * Which laws held. A law that could not run (because an earlier one failed)
   * is `null`, not `false`: a parse that threw says nothing about the editor.
   */
  laws: Record<SurvivalLaw, boolean | null>;
  mechanism: LossMechanism | null;
  /**
   * What carries a surviving pair: the typed model, or markup kept verbatim.
   *
   * The contract records a different disposition for each, and they are fixed
   * in different places: a modelled slot is a parser and a serializer, a
   * captured one is a byte range nothing understands. `null` on a pair that did
   * not survive, `"unknown"` when the probe could not decide.
   */
  carrier: "model" | "capture" | "both" | "unknown" | null;
  /** Present only when the pair could not be tested at all. */
  unrepresentable: string | null;
  detail: string | null;
};

type Presence = "absent" | "equal" | "different" | "truncated";

const ON = new Set(["1", "true", "on"]);
const OFF = new Set(["0", "false", "off"]);

const NUMBERS_PER_PERCENT: Readonly<Record<string, number>> = {
  fiftiethPercent: 50,
  thousandthPercent: 1000,
  wholePercent: 1,
};

/**
 * The Transitional spelling of a Strict-produced value, asked of folio's own table.
 *
 * folio rebuilds every package as Transitional, so `155.85pt` comes back as a
 * twip count and `50%` as whatever integer the slot's unit counts. That is a
 * contract, not a loss, and the law asks `transitionalSlotEncoding` — the same
 * generated table `captureVerbatimXml` consults — rather than restating it, so
 * the two cannot drift.
 */
const transitionalSpelling = (
  written: string,
  element: SubjectSlot | undefined,
  attributeLocalName: string | undefined,
): string | undefined => {
  if (element === undefined) {
    return undefined;
  }
  const encoding = transitionalSlotEncoding(element.namespace, element.name, attributeLocalName);
  if (encoding === undefined) {
    return undefined;
  }
  if (encoding.measure !== undefined) {
    const measure = universalMeasureAs(written, encoding.measure);
    if (measure !== undefined) {
      return String(measure);
    }
  }
  const percentage = /^(-?[0-9]+(?:\.[0-9]+)?)%$/u.exec(written);
  const unit = encoding.percent;
  if (percentage === null || unit === undefined) {
    return undefined;
  }
  // SAFETY: the capture group is present whenever the pattern matched.
  return String(Math.round(Number(percentage[1] as string) * (NUMBERS_PER_PERCENT[unit] ?? 1)));
};

type SubjectSlot = { namespace: string; name: string };

type ValueComparison = {
  written: string;
  read: string;
  /** The element the attribute sits on, for the slots whose spelling depends on it. */
  element: SubjectSlot | undefined;
  attributeLocalName: string | undefined;
};

/**
 * Whether two spellings of a value mean the same thing.
 *
 * Three equalities are folio's design rather than its defects, and calling any
 * of them a loss would bury the losses that are real:
 *
 * - `ST_OnOff` has six spellings of two values, and folio canonicalises them.
 * - A measure and a percentage have a Strict and a Transitional spelling, and
 *   folio rebuilds every package as Transitional.
 * - A revision element's `w:id` is a physical wrapper id, re-minted on every
 *   save by `revisionIdNormalization.ts` so that ids stay unique across a
 *   package. The set of elements that carries one is imported from there.
 */
const sameValue = ({ written, read, element, attributeLocalName }: ValueComparison): boolean => {
  if (written === read) {
    return true;
  }
  if ((ON.has(written) && ON.has(read)) || (OFF.has(written) && OFF.has(read))) {
    return true;
  }
  if (
    attributeLocalName === "id" &&
    element !== undefined &&
    element.namespace === WML_NAMESPACE &&
    REVISION_ELEMENT_NAMES.has(element.name)
  ) {
    return Number.isSafeInteger(Number(read));
  }
  if (transitionalSpelling(written, element, attributeLocalName) === read) {
    return true;
  }
  const writtenNumber = Number(written);
  const readNumber = Number(read);
  return (
    Number.isFinite(writtenNumber) && Number.isFinite(readNumber) && writtenNumber === readNumber
  );
};

/**
 * One start tag: the element's spelling and the attribute text after it.
 *
 * Attribute values are double-quoted and may hold a `>`, so the quoted runs are
 * consumed as units. A processing instruction, a comment and a CDATA section
 * all start with a character no element name may, and are skipped by name.
 */
const TAG_PATTERN = /<(\/?)([^\s/>]+)((?:"[^"]*"|[^>"])*)>/gu;

const QUOTED_VALUE = /"[^"]*"/gu;

/**
 * Whether a start tag closes itself.
 *
 * The slash has to be the one before the `>` rather than any slash in the text,
 * because an attribute value may end in one (`Target="media/"`), and a tag read
 * as self-closing that is not would unbalance every ancestor chain after it.
 */
const closesItself = (attributes: string): boolean =>
  attributes.replaceAll(QUOTED_VALUE, "").trimEnd().endsWith("/");

/**
 * Whether one occurrence sits where the fixture wrote the subject.
 *
 * The pair under test is (container, child): a `w:pgSz` inside the live section
 * is not the `w:pgSz` of the section snapshot a `w:sectPrChange` holds, and a
 * probe that asks only whether the name appears somewhere in the part answers
 * for the wrong one. So the fixture's innermost container has to be the
 * occurrence's own parent, and the chain above it has to appear in order from
 * the part root.
 *
 * The ancestors are a subsequence rather than an exact chain because a save may
 * legitimately wrap what it writes; the parent and the order are what say the
 * element came back where it was written.
 */
const underPath = (stack: readonly string[], path: readonly string[]): boolean => {
  const parent = path.at(-2);
  if (parent === undefined) {
    return stack.length === 1;
  }
  if (stack.at(-2) !== parent) {
    return false;
  }
  let matched = 0;
  for (let level = 0; level < stack.length - 1; level += 1) {
    if (stack[level] === path[matched]) {
      matched += 1;
    }
    if (matched === path.length - 1) {
      return true;
    }
  }
  return false;
};

/**
 * The attribute text of every occurrence of `path`'s last element that sits under `path`.
 *
 * `alsoNamed` is the canonical spelling of that last element, when
 * `CANONICAL_SPELLINGS` gives it one: folio writes `<w:left>` for the
 * `<w:start>` a document authored, and a probe that knows only the authored
 * name reports a rename as a loss.
 */
const occurrencesUnder = (
  xml: string,
  path: readonly string[],
  alsoNamed?: string | undefined,
): string[] => {
  const subject = path.at(-1);
  const found: string[] = [];
  const stack: string[] = [];
  for (const [, closing, name, attributes = ""] of xml.matchAll(TAG_PATTERN)) {
    if (name === undefined || name.startsWith("!") || name.startsWith("?")) {
      continue;
    }
    if (closing === "/") {
      stack.pop();
      continue;
    }
    const empty = closesItself(attributes);
    stack.push(name);
    if ((name === subject || name === alsoNamed) && underPath(stack, path)) {
      found.push(empty ? attributes.slice(0, attributes.lastIndexOf("/")) : attributes);
    }
    if (empty) {
      stack.pop();
    }
  }
  return found;
};

const attributeIn = (attributes: string, spelling: string): string | undefined => {
  const pattern = new RegExp(`\\s${spelling}="([^"]*)"`, "u");
  return pattern.exec(attributes)?.[1];
};

/**
 * What the law looks for, where in the part, and how many times.
 *
 * The chain is the one `fixture.ts` built the package with — the container's
 * own path from the rebuilt part's root — read from the same container space,
 * so the probe cannot search somewhere the fixture did not write.
 */
type ProbeTarget = {
  /** Element spellings from the part root down to the subject itself. */
  path: readonly string[];
  element: SubjectSlot;
  attributeSpelling: string | undefined;
  attributeLocalName: string | undefined;
  /** The value the fixture wrote, for an attribute subject. */
  value: string | undefined;
  /** What folio's canonical output may spell this subject as instead. */
  canonical: CanonicalForms;
};

/** Instances the fixture placed under the chain are what a save owes back. */
type SubjectProbe = ProbeTarget & { expected: number };

/** What the probe found where it searched. */
export type Probe = {
  presence: Presence;
  /** Instances found; for an attribute subject, the ones carrying an equal value. */
  found: number;
  expected: number;
  /** The chain the probe searched, from the part root. */
  location: string;
};

/**
 * Instances of the subject under the chain: the ones that carry the attribute
 * at all, and the ones that carry it with an equal value. A child subject
 * carries itself, so for it the two counts are the occurrences.
 */
const countUnder = (
  xml: string,
  { path, element, attributeSpelling, attributeLocalName, value, canonical }: ProbeTarget,
): { carrying: number; equal: number } => {
  const occurrences = occurrencesUnder(xml, path, canonical.element);
  if (attributeSpelling === undefined || value === undefined) {
    return { carrying: occurrences.length, equal: occurrences.length };
  }
  let carrying = 0;
  let equal = 0;
  for (const attributes of occurrences) {
    const carried =
      attributeIn(attributes, attributeSpelling) ??
      (canonical.attribute === undefined
        ? undefined
        : attributeIn(attributes, canonical.attribute));
    if (carried === undefined) {
      // An entry may say folio writes this value by leaving the attribute out —
      // a bare `<w:keepNext/>` is the on state. The equivalence is the value's,
      // so it reaches only this branch: an occurrence that carries the
      // attribute is compared on what it carries, and an element that did not
      // come back at all is still absent.
      if (canonical.absentAttribute) {
        carrying += 1;
        equal += 1;
      }
      continue;
    }
    carrying += 1;
    if (sameValue({ written: value, read: carried, element, attributeLocalName })) {
      equal += 1;
    }
  }
  return { carrying, equal };
};

/**
 * Whether the saved part still carries the subject, where it was written and as
 * often as it was written.
 *
 * A shortfall is reported rather than rounded up to a survival: a reader that
 * keeps one of a repeated particle's instances and drops the rest writes markup
 * the source did not describe, and the element being present says nothing about
 * it.
 */
const presenceIn = (xml: string, probe: SubjectProbe): Probe => {
  const { carrying, equal } = countUnder(xml, probe);
  const location = probe.path.join("/");
  const report = (presence: Presence, found: number): Probe => ({
    presence,
    found,
    expected: probe.expected,
    location,
  });
  if (carrying === 0) {
    return report("absent", 0);
  }
  if (equal === 0) {
    return report("different", 0);
  }
  return report(equal >= probe.expected ? "equal" : "truncated", equal);
};

/** How a count shortfall reads in a report, and nothing when there is none. */
const shortfall = ({ presence, found, expected, location }: Probe): string | null =>
  presence === "truncated" ? `found ${found} of ${expected} under ${location}` : null;

/**
 * The chain the fixture built the subject's package with, spelled.
 *
 * Read from the container space `fixture.ts` reads, so where the law looks and
 * where the fixture wrote are one derivation rather than two that can drift.
 */
const pathTo = (space: ContainerSpace, subject: Subject): string[] | undefined => {
  const container = space.containers.get(containerKey(subject.slot.container));
  if (container === undefined) {
    return undefined;
  }
  const elements = container.path.map(({ element }) => element);
  const chain: string[] = [];
  for (const element of subject.kind === "child" ? [...elements, subject.slot.child] : elements) {
    const spelled = spell(element);
    if (spelled === undefined) {
      return undefined;
    }
    chain.push(spelled);
  }
  return chain;
};

/**
 * How many instances of the subject the schema admits under one container.
 *
 * The generator can write more than that: a seed and the subject can land on
 * the same particle — `w:numPr` is seeded with the `w:ilvl`/`w:numId` pair that
 * makes it a list, and the `w:numId` pair under test is a second one — and the
 * part validator checks membership and order rather than maxima. Capping the
 * expectation at the schema's own maximum is what keeps the law from charging
 * folio for keeping the one instance the schema allows.
 */
const admittedInstances = (space: ContainerSpace, subject: Subject): number => {
  const declared =
    subject.kind === "child" ? subject.slot.maxOccurs : declaringMaxOccurs(space, subject);
  const admitted = Number.parseInt(declared, 10);
  return Number.isInteger(admitted) ? admitted : Number.POSITIVE_INFINITY;
};

/** The `maxOccurs` of the element an attribute sits on, read where its parent declares it. */
const declaringMaxOccurs = (space: ContainerSpace, subject: Subject): string => {
  const element = qualify(subject.slot.container.element);
  const parentId = space.containers.get(containerKey(subject.slot.container))?.path.at(-2);
  const parent = parentId === undefined ? undefined : space.containers.get(containerKey(parentId));
  return parent?.children.find((slot) => qualify(slot.child) === element)?.maxOccurs ?? "unbounded";
};

type ProbeOptions = { space: ContainerSpace; subject: Subject; fixture: BuiltFixture };

const probeFor = ({ space, subject, fixture }: ProbeOptions): SubjectProbe | undefined => {
  const path = pathTo(space, subject);
  if (path === undefined) {
    return undefined;
  }
  const target: ProbeTarget = {
    path,
    element: fixture.subjectElement,
    attributeSpelling: fixture.attributeSpelling,
    attributeLocalName: fixture.attributeLocalName,
    value: subject.kind === "attribute" ? subject.value : undefined,
    canonical: canonicalFormsOf({
      element: fixture.subjectElement,
      type:
        subject.kind === "child" ? subject.slot.childTypeQName : subject.slot.container.typeQName,
      attribute: subject.kind === "attribute" ? subject.slot.attribute : undefined,
      value: subject.kind === "attribute" ? subject.value : undefined,
    }),
  };
  // The expectation is measured on the fixture with the probe that measures the
  // save, so the two counts cannot disagree about what the generator wrote —
  // a repeated particle, a seed and a partner marker all land in it by
  // themselves — and it is then bounded by what the schema admits.
  const written = countUnder(fixture.documentXml, target).equal;
  return { ...target, expected: Math.min(written, admittedInstances(space, subject)) };
};

/** Where the law looked for a pair and what it found there, for the census's `explain`. */
export const probeOf = (options: ProbeOptions & { xml: string }): Probe | undefined => {
  const probe = probeFor(options);
  return probe === undefined ? undefined : presenceIn(options.xml, probe);
};

/**
 * Every slot in the model that holds markup rather than a parsed shape.
 *
 * A superset of the corpus `reserialize` invariant's list, which strips only
 * the captures that have a model behind them. Here the point is the opposite:
 * clearing all of them says which pairs the typed model could rebuild on its
 * own, and a pair that stops surviving when they are gone is carried by bytes.
 */
const CAPTURE_SLOT_NAMES = new Set([
  "gridSourceXml",
  "numberingChangeXml",
  "numberingInsertionXml",
  "ommlXml",
  "propertiesXml",
  "rawEndPropertiesXml",
  "rawPropertiesXml",
  "rawWatermarkXml",
  "rawXml",
  "sourceXml",
  "verbatimXml",
]);

/**
 * Union members that are a capture rather than a field holding one.
 *
 * The verbatim sink puts a container's unmodelled children in the model's own
 * shape — `RunContent`'s `preservedXml`, `ParagraphContent`'s
 * `preservedInline`, `BlockContent`'s `preservedBlock` —
 * so a pair kept by one is carried by bytes even though nothing on it is
 * spelled `rawSomethingXml`. Without this the carrier question answers
 * "model" for every one of them and the contract records `modelled` for
 * markup the editor cannot touch.
 */
const CAPTURE_MEMBER_TYPES = new Set(["preservedBlock", "preservedInline", "preservedXml"]);

/**
 * The sink itself, and the attribute remainder beside it.
 *
 * `preserved` is the ordered child sink; `preservedAttributes` is the list of
 * attributes an element carried that its record has no field for. Both hold
 * source spellings rather than parsed shapes, so a pair that stops surviving
 * once they are cleared is carried by bytes and the contract records
 * `captured-verbatim` rather than `modelled`.
 *
 * `w:ind` and `w:spacing` were flattened into `ParagraphFormatting`, so their
 * remainders are fields of it rather than of a record of their own, and they
 * are named here for the same reason under the names they actually have.
 */
const CAPTURE_SINK_KEYS = new Set([
  "preserved",
  "preservedAttributes",
  "indentPreservedAttributes",
  "spacingPreservedAttributes",
]);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isCaptureMember = (value: unknown): value is Record<string, unknown> =>
  isRecord(value) && typeof value["type"] === "string" && CAPTURE_MEMBER_TYPES.has(value["type"]);

const clearCapturesInPlace = (value: unknown, seen: WeakSet<object>): void => {
  if (Array.isArray(value)) {
    // Backwards, because a capture member is removed rather than emptied.
    for (let index = value.length - 1; index >= 0; index -= 1) {
      if (isCaptureMember(value[index])) {
        value.splice(index, 1);
        continue;
      }
      clearCapturesInPlace(value[index], seen);
    }
    return;
  }
  if (value instanceof Map) {
    for (const item of value.values()) {
      clearCapturesInPlace(item, seen);
    }
    return;
  }
  if (!isRecord(value) || seen.has(value)) {
    return;
  }
  seen.add(value);
  for (const key of Object.keys(value)) {
    if (CAPTURE_SLOT_NAMES.has(key) || CAPTURE_SINK_KEYS.has(key)) {
      value[key] = undefined;
      continue;
    }
    clearCapturesInPlace(value[key], seen);
  }
};

/** A document with nothing verbatim left: whatever survives this, the model holds. */
const withoutAnyVerbatimMarkup = (document: Document): Document => {
  const cloned = structuredClone(withoutSerializerCaptures(document));
  clearCapturesInPlace(cloned.package, new WeakSet());
  return cloned;
};

/** Every string the model holds in a verbatim-capture slot, concatenated. */
const captureText = (value: unknown, seen: WeakSet<object>, into: string[]): void => {
  if (Array.isArray(value)) {
    for (const item of value) {
      captureText(item, seen, into);
    }
    return;
  }
  if (value instanceof Map) {
    for (const item of value.values()) {
      captureText(item, seen, into);
    }
    return;
  }
  if (typeof value !== "object" || value === null || seen.has(value)) {
    return;
  }
  seen.add(value);
  if (isCaptureMember(value) && typeof value["xml"] === "string") {
    into.push(value["xml"]);
    return;
  }
  for (const [key, item] of Object.entries(value)) {
    if (CAPTURE_SLOT_NAMES.has(key) && typeof item === "string") {
      into.push(item);
      continue;
    }
    captureText(item, seen, into);
  }
};

/**
 * Where the subject's name shows up in the parsed model.
 *
 * This is what separates "the parser never saw it" from "the parser captured it
 * and the serializer refused the capture". It matches on the local name, so it
 * is a probe rather than a proof: a name that is also an ordinary model key
 * reads as modelled. The census says which pairs it decided, and the contract's
 * agreement check is what actually holds the line.
 */
type ModelTrace = { inModel: boolean; inCapture: boolean };

const traceSubject = (parsed: Document, localName: string): ModelTrace => {
  const captures: string[] = [];
  captureText(parsed.package, new WeakSet(), captures);
  const capturedXml = captures.join("");
  const stripped = Result.try({
    try: () => JSON.stringify(withoutSerializerCaptures(parsed).package),
    catch: (cause: unknown) => cause,
  });
  return {
    inModel: stripped.isOk() && stripped.value.includes(localName),
    inCapture: capturedXml.includes(`:${localName}`) || capturedXml.includes(`<${localName}`),
  };
};

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * A numbering definition the fixtures can point at.
 *
 * `w:numPr` that names a `w:numId` no part defines is not a list, and folio
 * drops it. Every pair inside `w:numPr` — `w:numberingChange` among them —
 * would then read as lost because the fixture had no numbering, not because
 * folio lost anything.
 */
const NUMBERING_PART = `${XML_DECLARATION}<w:numbering xmlns:w="${WML_NAMESPACE}">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

const note = (id: number, body: string, type?: string): string =>
  `<w:footnote${type === undefined ? "" : ` w:type="${type}"`} w:id="${id}">${body}</w:footnote>`;

const SEPARATORS =
  `${note(-1, "<w:p><w:r><w:separator/></w:r></w:p>", "separator")}` +
  `${note(0, "<w:p><w:r><w:continuationSeparator/></w:r></w:p>", "continuationSeparator")}`;

/**
 * The notes and comments a fixture's references can point at.
 *
 * A `w:footnoteReference` naming a note no part defines is a dangling
 * reference, and folio refuses the document — correctly. Without these parts
 * the law would report the reference and both its attributes as "the parser
 * throws", which says something about the fixture rather than about folio.
 * `w:id="1"` is the note and comment every reference fixture names, because
 * `representativeValue` gives `ST_DecimalNumber` a `1`.
 */
const SIDE_PARTS: ReadonlyArray<{
  path: string;
  xml: string;
  contentType: string;
  relationship: string;
  relationshipId: string;
}> = [
  {
    path: "word/numbering.xml",
    xml: NUMBERING_PART,
    contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    relationship: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
    relationshipId: "rIdContainerSurvivalNumbering",
  },
  {
    path: "word/footnotes.xml",
    xml:
      `${XML_DECLARATION}<w:footnotes xmlns:w="${WML_NAMESPACE}">${SEPARATORS}` +
      `${note(1, "<w:p><w:r><w:t>note</w:t></w:r></w:p>")}</w:footnotes>`,
    contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
    relationship: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
    relationshipId: "rIdContainerSurvivalFootnotes",
  },
  {
    path: "word/endnotes.xml",
    xml:
      `${XML_DECLARATION}<w:endnotes xmlns:w="${WML_NAMESPACE}">` +
      `${SEPARATORS.replaceAll("footnote", "endnote")}` +
      `${note(1, "<w:p><w:r><w:t>note</w:t></w:r></w:p>").replaceAll("footnote", "endnote")}` +
      "</w:endnotes>",
    contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
    relationship: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
    relationshipId: "rIdContainerSurvivalEndnotes",
  },
  {
    path: "word/comments.xml",
    xml:
      `${XML_DECLARATION}<w:comments xmlns:w="${WML_NAMESPACE}">` +
      '<w:comment w:id="1" w:author="folio" w:date="2024-01-01T00:00:00Z">' +
      "<w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment></w:comments>",
    contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    relationship: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    relationshipId: "rIdContainerSurvivalComments",
  },
];

/**
 * The picture every synthesised drawing points at: a 1×1 PNG, the smallest
 * thing that makes `a:blip r:embed` resolve to real media.
 *
 * Without it a drawing names no picture relationship, folio classifies it
 * preserve-only, and the save replays its captured bytes — so every pair inside
 * a `w:drawing` read as surviving on the strength of a byte copy and the
 * rebuild path was never measured at all.
 */
const MEDIA_PART_PATH = "word/media/folio.png";

const ONE_PIXEL_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  "base64",
);

const IMAGE_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/**
 * Put the media part, its content type and the relationship in the package.
 *
 * The relationship has to live in the rels of the part that holds the drawing,
 * so a fixture rooted at a header declares it there rather than in the
 * document's.
 */
const withMediaPart = async (zip: JSZip, partPath: string): Promise<void> => {
  zip.file(MEDIA_PART_PATH, ONE_PIXEL_PNG);
  const types = await zip.file("[Content_Types].xml")?.async("text");
  if (types === undefined) {
    throw new Error("the empty package lost its packaging parts");
  }
  if (!types.includes('Extension="png"')) {
    zip.file(
      "[Content_Types].xml",
      types.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>'),
    );
  }

  const slash = partPath.lastIndexOf("/");
  const relsPath = `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`;
  const existing = await zip.file(relsPath)?.async("text");
  const relationship =
    `<Relationship Id="${IMAGE_RELATIONSHIP_ID}" Type="${IMAGE_RELATIONSHIP_TYPE}" ` +
    `Target="media/${MEDIA_PART_PATH.slice("word/media/".length)}"/>`;
  if (existing === undefined) {
    zip.file(
      relsPath,
      `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationship}</Relationships>`,
    );
    return;
  }
  if (existing.includes(IMAGE_RELATIONSHIP_ID)) {
    return;
  }
  zip.file(relsPath, existing.replace("</Relationships>", `${relationship}</Relationships>`));
};

let basePackage: Promise<ArrayBuffer> | undefined;

const withSideParts = async (zip: JSZip): Promise<void> => {
  const types = await zip.file("[Content_Types].xml")?.async("text");
  const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
  if (types === undefined || rels === undefined) {
    throw new Error("the empty package lost its packaging parts");
  }
  let overrides = "";
  let relationships = "";
  for (const part of SIDE_PARTS) {
    if (zip.file(part.path) !== null) {
      continue;
    }
    zip.file(part.path, part.xml);
    overrides += `<Override PartName="/${part.path}" ContentType="${part.contentType}"/>`;
    relationships += `<Relationship Id="${part.relationshipId}" Type="${part.relationship}" Target="${part.path.slice("word/".length)}"/>`;
  }
  if (overrides === "") {
    return;
  }
  zip.file("[Content_Types].xml", types.replace("</Types>", `${overrides}</Types>`));
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace("</Relationships>", `${relationships}</Relationships>`),
  );
};

/**
 * Declare a part the base package does not already carry.
 *
 * A fixture rooted at `w:settings` or `w:hdr` is a part of its own: without the
 * content-type override and the relationship, a reader does not find it and the
 * census would report every pair in it as never parsed, which is a fact about
 * the packaging rather than about folio.
 */
const declarePart = async (
  zip: JSZip,
  part: RebuiltPart,
  relationshipId: string,
): Promise<void> => {
  const types = await zip.file("[Content_Types].xml")?.async("text");
  const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
  if (types === undefined || rels === undefined) {
    throw new Error("the empty package lost its packaging parts");
  }
  if (!types.includes(`PartName="/${part.path}"`)) {
    zip.file(
      "[Content_Types].xml",
      types.replace(
        "</Types>",
        `<Override PartName="/${part.path}" ContentType="${part.contentType}"/></Types>`,
      ),
    );
  }
  if (!rels.includes(`Type="${part.relationship}"`)) {
    zip.file(
      "word/_rels/document.xml.rels",
      rels.replace(
        "</Relationships>",
        `<Relationship Id="${relationshipId}" Type="${part.relationship}" Target="${part.path.slice("word/".length)}"/></Relationships>`,
      ),
    );
  }
};

/** The id the subject part is related under, when the base package has none. */
const SUBJECT_RELATIONSHIP_ID = "rIdContainerSurvivalSubject";

/**
 * A `w:hdr` or `w:ftr` part is only read through a section's reference, so the
 * body has to point at it or the part is dead markup a parser never opens.
 */
const SECTION_REFERENCES: Partial<Record<string, string>> = {
  "word/header1.xml": `<w:headerReference r:id="${SUBJECT_RELATIONSHIP_ID}" w:type="default"/>`,
  "word/footer1.xml": `<w:footerReference r:id="${SUBJECT_RELATIONSHIP_ID}" w:type="default"/>`,
};

const withSectionReference = async (zip: JSZip, partPath: string): Promise<void> => {
  const reference = SECTION_REFERENCES[partPath];
  const documentXml = await zip.file("word/document.xml")?.async("text");
  if (reference === undefined || documentXml === undefined) {
    return;
  }
  zip.file("word/document.xml", documentXml.replace("<w:sectPr>", `<w:sectPr>${reference}`));
};

/**
 * The part a `w:headerReference` or `w:footerReference` fixture names.
 *
 * folio removes a reference whose part is missing, which is right — a dangling
 * `r:id` is what makes Word offer to repair the file — so without the part the
 * census reports the reference as never parsed, and that is a fact about the
 * fixture. The id is read out of the fixture rather than restated, because the
 * fixture writes whatever `representativeValue` gives `ST_RelationshipId`.
 *
 * It is added only for the fixture that names it: two side parts cannot share
 * one relationship id, and every reference fixture writes the same one.
 */
const REFERENCED_PARTS: Readonly<Record<string, { part: RebuiltPart; xml: string }>> = {
  "w:headerReference": {
    part: REBUILT_PARTS.hdr,
    xml:
      `${XML_DECLARATION}<w:hdr xmlns:w="${WML_NAMESPACE}">` +
      "<w:p><w:r><w:t>header</w:t></w:r></w:p></w:hdr>",
  },
  "w:footerReference": {
    part: REBUILT_PARTS.ftr,
    xml:
      `${XML_DECLARATION}<w:ftr xmlns:w="${WML_NAMESPACE}">` +
      "<w:p><w:r><w:t>footer</w:t></w:r></w:p></w:ftr>",
  },
};

const withReferencedPart = async (zip: JSZip, fixture: BuiltFixture): Promise<void> => {
  const referenced = REFERENCED_PARTS[fixture.subjectSpelling];
  const relationshipId = new RegExp(`<${fixture.subjectSpelling}[^>]*\\sr:id="([^"]*)"`, "u").exec(
    fixture.documentXml,
  )?.[1];
  if (referenced === undefined || relationshipId === undefined || zip.file(referenced.part.path)) {
    return;
  }
  zip.file(referenced.part.path, referenced.xml);
  await declarePart(zip, referenced.part, relationshipId);
};

const packageFor = async (fixture: BuiltFixture): Promise<ArrayBuffer> => {
  basePackage ??= createEmptyDocx();
  const zip = await JSZip.loadAsync(await basePackage);
  zip.file(fixture.part.path, fixture.documentXml);
  await declarePart(zip, fixture.part, SUBJECT_RELATIONSHIP_ID);
  await withSectionReference(zip, fixture.part.path);
  await withReferencedPart(zip, fixture);
  await withSideParts(zip);
  await withMediaPart(zip, fixture.part.path);
  return zip.generateAsync({ type: "arraybuffer" });
};

const partOf = async (buffer: ArrayBuffer, partPath: string): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  return (await zip.file(partPath)?.async("text")) ?? "";
};

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

/**
 * The fixture's own part, as folio writes it from a model with no captures left.
 *
 * Two legs, and which one runs is decided by the bytes rather than by a list of
 * parts. A repack rebuilds the document part and its neighbours, so for those
 * the save itself is the forcing. It copies every declaration part across, so
 * for those the save hands the fixture back unchanged and the part's own
 * serializer has to be called directly — otherwise every pair in the part would
 * read as surviving on the strength of a file copy.
 *
 * There is a third answer: the part is copied and `PART_REBUILDERS` states no
 * rebuilder for it, so there is nothing to measure and the caller says so
 * rather than passing the pair.
 */
type ForcedPart =
  | { kind: "repacked"; xml: string }
  | { kind: "part-serializer"; xml: string }
  | { kind: "copied-with-no-rebuilder"; reason: string };

const forcePart = async (fixture: BuiltFixture, document: Document): Promise<ForcedPart> => {
  const repacked = await partOf(await save(document), fixture.part.path);
  if (repacked !== fixture.documentXml) {
    return { kind: "repacked", xml: repacked };
  }
  const rebuild = partRebuildFor(fixture.partRoot);
  if (rebuild.kind === "absent") {
    return { kind: "copied-with-no-rebuilder", reason: rebuild.reason };
  }
  // A model with no record for the part rebuilds to nothing, which is a loss
  // the probe reports rather than a case the law skips.
  return { kind: "part-serializer", xml: rebuild.rebuild(document) ?? "" };
};

/** The part a forcing produced, and nothing when it could not force one. */
const forcedXmlOf = (forced: ForcedPart): string | undefined => {
  switch (forced.kind) {
    case "repacked":
    case "part-serializer": {
      return forced.xml;
    }
    case "copied-with-no-rebuilder": {
      return undefined;
    }
    default: {
      const unreachable: never = forced;
      return unreachable;
    }
  }
};

/** The fixture's own part as a forced save writes it, for the census's `explain`. */
export const forcedSavePart = async (fixture: BuiltFixture): Promise<string> => {
  const parsed = await parseDocx(await packageFor(fixture), { preloadFonts: false });
  return forcedXmlOf(await forcePart(fixture, withoutSerializerCaptures(parsed))) ?? "";
};

/**
 * The same, through the editor.
 *
 * `lost-in-the-editor-projection` is the one mechanism the forced save cannot
 * show: the markup is in L2's output and gone from L3's, so `explain` has to
 * print both or the reader is left comparing the input against a part that
 * still has what it is looking for.
 */
export const editorSavePart = async (fixture: BuiltFixture): Promise<string> => {
  const parsed = await parseDocx(await packageFor(fixture), { preloadFonts: false });
  const projected = projectWithoutReuse(toProseDoc(parsed), parsed);
  return forcedXmlOf(await forcePart(fixture, withoutSerializerCaptures(projected))) ?? "";
};

/**
 * What the editor leg found, and `null` when it did not run.
 *
 * A thrown projection is `absent` — the markup did not come back — and a leg
 * the law skipped is neither present nor absent.
 */
const editorPresenceOf = (probe: Result<Probe, unknown> | undefined): Presence | null => {
  if (probe === undefined) {
    return null;
  }
  return probe.isOk() ? probe.value.presence : "absent";
};

/**
 * The body of a part, so a printed comparison is about the pair and not the
 * boilerplate. A part with no `w:body` is already all subject.
 */
export const bodyOf = (xml: string): string => /<w:body>.*<\/w:body>/su.exec(xml)?.[0] ?? xml;

export const subjectKey = (subject: Subject): string =>
  subject.kind === "child" ? childSlotKey(subject.slot) : attributeSlotKey(subject.slot);

const subjectName = (subject: Subject): string =>
  qualify(subject.kind === "child" ? subject.slot.child : subject.slot.attribute);

const outcomeShell = (subject: Subject): PairOutcome => ({
  key: subjectKey(subject),
  kind: subject.kind,
  container: containerKey(subject.slot.container),
  subject: subjectName(subject),
  laws: {
    [SURVIVAL_LAWS.parse]: null,
    [SURVIVAL_LAWS.serialize]: null,
    [SURVIVAL_LAWS.editor]: null,
    [SURVIVAL_LAWS.schema]: null,
  },
  mechanism: null,
  carrier: null,
  unrepresentable: null,
  detail: null,
});

/**
 * Run all four laws over one pair.
 *
 * Nothing here throws: a pair that makes the parser or a serializer throw is a
 * finding, not a crashed census.
 */
export const runSurvivalLaws = async (
  space: ContainerSpace,
  subject: Subject,
): Promise<PairOutcome> => {
  const alone = await runLawsOnce(space, subject);
  if (subject.kind !== "attribute" || alone.unrepresentable !== null || alone.mechanism !== null) {
    return alone;
  }
  const companion = modelledCompanionFor(space, subject.slot);
  if (companion === undefined) {
    return alone;
  }
  // The same pair, stated beside an attribute the element's own record models.
  // A reader that decides a property element whole keeps every attribute of an
  // element it takes nothing from and none of an element it models, so a
  // one-attribute-at-a-time census reports the second case as surviving. The
  // pair survives when it survives both.
  const beside = await runLawsOnce(space, {
    kind: "attribute",
    slot: subject.slot,
    value: subject.value,
    companion,
  });
  if (beside.unrepresentable !== null || beside.mechanism === null) {
    return alone;
  }
  beside.detail = `lost only beside the modelled ${companion.spelled}`;
  return beside;
};

const runLawsOnce = async (space: ContainerSpace, subject: Subject): Promise<PairOutcome> => {
  const outcome = outcomeShell(subject);

  const built = buildFixture(space, subject);
  if (built.status === "unrepresentable") {
    outcome.unrepresentable = built.reason;
    return outcome;
  }
  const { fixture } = built;

  const graph = await loadSchemaGraph();
  const fixtureViolations = validateOoxmlPart({ graph, xml: fixture.documentXml });
  if (fixtureViolations.length > 0) {
    const first = fixtureViolations[0];
    outcome.unrepresentable = `the generated fixture is not schema-valid: ${first?.kind} at ${first?.path} (${first?.name})`;
    return outcome;
  }

  const probe = probeFor({ space, subject, fixture });
  if (probe === undefined) {
    outcome.unrepresentable = "no prefix is bound for an ancestor of the subject";
    return outcome;
  }
  if (probe.expected === 0) {
    // The law would then be vacuous: nothing to find, so everything survives.
    // Counting it as unmeasured says so, where passing it would hide a
    // generator that wrote the subject somewhere other than its own chain.
    outcome.unrepresentable = "the generated fixture writes no subject under its own chain";
    outcome.detail = `no ${fixture.subjectSpelling} under ${probe.path.join("/")}`;
    return outcome;
  }
  const buffer = await packageFor(fixture);

  const parsed = await Result.tryPromise({
    try: () => parseDocx(buffer, { preloadFonts: false }),
    catch: (cause: unknown) => cause,
  });
  if (parsed.isErr()) {
    outcome.laws[SURVIVAL_LAWS.parse] = false;
    outcome.detail = String(parsed.error).slice(0, 200);
    return outcome;
  }
  outcome.laws[SURVIVAL_LAWS.parse] = true;

  const replayed = await Result.tryPromise({
    try: async () => partOf(await save(parsed.value), fixture.part.path),
    catch: (cause: unknown) => cause,
  });
  // Stripping the captures makes the *element* serializers run; it does not
  // make a part serializer run. A repack copies `word/styles.xml`,
  // `word/settings.xml` and the other declaration parts through byte for byte,
  // so for those the forcing has to reach part level: `forcePart` calls the
  // part's own serializer when the save handed the fixture back unchanged.
  // The test is the bytes themselves rather than a list of parts, so a part
  // folio starts rebuilding on the save path needs no change here.
  const forced = await Result.tryPromise({
    try: () => forcePart(fixture, withoutSerializerCaptures(parsed.value)),
    catch: (cause: unknown) => cause,
  });

  if (forced.isErr()) {
    outcome.laws[SURVIVAL_LAWS.serialize] = false;
    outcome.detail = String(forced.error).slice(0, 200);
    return outcome;
  }

  // A copied part folio cannot rebuild has nothing to measure. Recording it as
  // a survival would put a `modelled` disposition on a slot no model holds, so
  // the pair stays unmeasured with the absence named.
  if (forced.value.kind === "copied-with-no-rebuilder") {
    outcome.laws[SURVIVAL_LAWS.parse] = null;
    outcome.unrepresentable = `a repack replays ${fixture.part.path} verbatim, and ${forced.value.reason}`;
    return outcome;
  }
  const forcedXml = forced.value.xml;
  const onPartLeg = forced.value.kind === "part-serializer";

  const forcedProbe = presenceIn(forcedXml, probe);
  outcome.laws[SURVIVAL_LAWS.serialize] = forcedProbe.presence === "equal";
  outcome.laws[SURVIVAL_LAWS.schema] = validateOoxmlPart({ graph, xml: forcedXml }).length === 0;

  // L3 asks what the ProseMirror projection carries, and a declaration part is
  // not in it: the editor holds a document, and `word/fontTable.xml` reaches a
  // save the same way whether or not anything was edited. Reporting `false`
  // there would charge the pair for a leg that never ran, and reporting `true`
  // would claim a projection nobody wrote, so the law reports neither.
  const editorProbe = onPartLeg
    ? undefined
    : await Result.tryPromise({
        try: async () => {
          const projected = projectWithoutReuse(toProseDoc(parsed.value), parsed.value);
          return presenceIn(
            forcedXmlOf(await forcePart(fixture, withoutSerializerCaptures(projected))) ?? "",
            probe,
          );
        },
        catch: (cause: unknown) => cause,
      });
  outcome.laws[SURVIVAL_LAWS.editor] =
    editorProbe === undefined ? null : editorProbe.isOk() && editorProbe.value.presence === "equal";
  if (editorProbe?.isErr()) {
    outcome.detail = String(editorProbe.error).slice(0, 200);
  }
  if (outcome.detail === null) {
    outcome.detail =
      shortfall(forcedProbe) ?? (editorProbe?.isOk() ? shortfall(editorProbe.value) : null);
  }

  // A child's container is the element that declares it; an attribute's is the
  // element it sits on, which is the chain's last step.
  const containerPath = subject.kind === "child" ? probe.path.slice(0, -1) : probe.path;
  // A canonical element spelling belongs to the chain's last step, so it is the
  // container's own spelling only when the subject is an attribute sitting on
  // it. For a child subject the last step is the child, and the container above
  // it is spelled one way.
  const containerSpelling = subject.kind === "child" ? undefined : probe.canonical.element;
  const localName =
    subject.kind === "child" ? subject.slot.child.name : subject.slot.attribute.name;
  const trace = () => traceSubject(parsed.value, localName);
  outcome.mechanism = classify({
    forcedPresence: forcedProbe.presence,
    containerPresent: occurrencesUnder(forcedXml, containerPath, containerSpelling).length > 0,
    replayedPresence: replayed.isOk() ? presenceIn(replayed.value, probe).presence : "absent",
    editorPresence: editorPresenceOf(editorProbe),
    trace,
  });
  if (outcome.mechanism === null) {
    // A pair that survives with every verbatim slot cleared is held by the
    // typed model; one that stops surviving is held by bytes. Asking the
    // question by execution beats guessing from the shape of the model,
    // because a modelled slot rarely keeps the schema's name for itself. The
    // probe goes through the same forcing the pair was measured with, or a
    // part the repack copies would answer "model" for markup no model holds.
    const modelOnly = await Result.tryPromise({
      try: async () =>
        presenceIn(
          forcedXmlOf(await forcePart(fixture, withoutAnyVerbatimMarkup(parsed.value))) ?? "",
          probe,
        ),
      catch: (cause: unknown) => cause,
    });
    if (modelOnly.isErr()) {
      outcome.carrier = "unknown";
    } else {
      outcome.carrier = modelOnly.value.presence === "equal" ? "model" : "capture";
    }
  }
  return outcome;
};

type Classification = {
  forcedPresence: Presence;
  containerPresent: boolean;
  replayedPresence: Presence;
  /** `null` when the editor leg did not run, which a declaration part's never does. */
  editorPresence: Presence | null;
  trace: () => ModelTrace;
};

const classify = ({
  forcedPresence,
  containerPresent,
  replayedPresence,
  editorPresence,
  trace,
}: Classification): LossMechanism | null => {
  if (forcedPresence === "different") {
    return LOSS_MECHANISMS.respelled;
  }
  // A shortfall is only visible where the container came back, so this is never
  // the container's defect wearing a narrower name.
  if (forcedPresence === "truncated") {
    return LOSS_MECHANISMS.repeatTruncated;
  }
  if (forcedPresence === "equal") {
    // A leg that did not run loses nothing: `lost-in-the-editor-projection`
    // names a projection, and a part outside the projection cannot have one.
    return editorPresence === null || editorPresence === "equal"
      ? null
      : LOSS_MECHANISMS.editorProjection;
  }
  if (replayedPresence !== "absent") {
    return LOSS_MECHANISMS.replayOnly;
  }
  if (!containerPresent) {
    return LOSS_MECHANISMS.containerLost;
  }
  const { inModel, inCapture } = trace();
  if (inCapture && !inModel) {
    return LOSS_MECHANISMS.replayRejected;
  }
  return inModel ? LOSS_MECHANISMS.parsedNotSerialized : LOSS_MECHANISMS.neverParsed;
};
