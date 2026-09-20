/**
 * The survival law: what folio reads, folio writes back.
 *
 * Four laws run over one synthesised pair. They are reported separately
 * because they fail for different reasons and are fixed in different places.
 *
 * - **L1 parse** — `parseDocx` does not throw on the fixture.
 * - **L2 serialize** — parse, force every serializer to run by removing the
 *   verbatim captures replay would otherwise hand back, save, and find the
 *   subject in the saved part with an equal value.
 * - **L3 editor** — the same through `toProseDoc`/`fromProseDoc`, which is the
 *   path every edited document takes.
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
import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import type { Document } from "@stll/folio-core/types/document";
import { Result } from "better-result";

import { loadSchemaGraph, validateOoxmlPart } from "../corpus-schema-validator";
import { withoutSerializerCaptures } from "../corpus-invariants/reserialize";
import { type BuiltFixture, buildFixture, spell, type Subject } from "./fixture";
import {
  attributeSlotKey,
  childSlotKey,
  containerKey,
  type ContainerSpace,
  qualify,
  type RebuiltPart,
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

type Presence = "absent" | "equal" | "different";

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

const elementOccurrences = (xml: string, spelling: string): string[] => {
  const pattern = new RegExp(`<${spelling}(\\s[^>]*)?/?>`, "gu");
  return [...xml.matchAll(pattern)].map((match) => match[1] ?? "");
};

const attributeIn = (attributes: string, spelling: string): string | undefined => {
  const pattern = new RegExp(`\\s${spelling}="([^"]*)"`, "u");
  return pattern.exec(attributes)?.[1];
};

/** Whether the saved part still carries the subject, and with which value. */
const presenceIn = (xml: string, fixture: BuiltFixture, expected: string | undefined): Presence => {
  const occurrences = elementOccurrences(xml, fixture.subjectSpelling);
  if (occurrences.length === 0) {
    return "absent";
  }
  if (fixture.attributeSpelling === undefined || expected === undefined) {
    return "equal";
  }
  const read = occurrences
    .map((attributes) => attributeIn(attributes, fixture.attributeSpelling ?? ""))
    .filter((value): value is string => value !== undefined);
  if (read.length === 0) {
    return "absent";
  }
  return read.some((value) =>
    sameValue({
      written: expected,
      read: value,
      element: fixture.subjectElement,
      attributeLocalName: fixture.attributeLocalName,
    }),
  )
    ? "equal"
    : "different";
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
  "gridChangeXml",
  "gridSourceXml",
  "numberingChangeXml",
  "ommlXml",
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

/** The sink itself, on a container whose model holds one kind of child. */
const CAPTURE_SINK_KEYS = new Set(["preserved"]);

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

const packageFor = async (fixture: BuiltFixture): Promise<ArrayBuffer> => {
  basePackage ??= createEmptyDocx();
  const zip = await JSZip.loadAsync(await basePackage);
  zip.file(fixture.part.path, fixture.documentXml);
  await declarePart(zip, fixture.part, SUBJECT_RELATIONSHIP_ID);
  await withSectionReference(zip, fixture.part.path);
  await withSideParts(zip);
  return zip.generateAsync({ type: "arraybuffer" });
};

const partOf = async (buffer: ArrayBuffer, partPath: string): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  return (await zip.file(partPath)?.async("text")) ?? "";
};

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

/** The fixture's own part as a forced save writes it, for the census's `explain`. */
export const forcedSavePart = async (fixture: BuiltFixture): Promise<string> => {
  const parsed = await parseDocx(await packageFor(fixture), { preloadFonts: false });
  return partOf(await save(withoutSerializerCaptures(parsed)), fixture.part.path);
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

  const expected = subject.kind === "attribute" ? subject.value : undefined;
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
  const forced = await Result.tryPromise({
    try: async () => partOf(await save(withoutSerializerCaptures(parsed.value)), fixture.part.path),
    catch: (cause: unknown) => cause,
  });

  if (forced.isErr()) {
    outcome.laws[SURVIVAL_LAWS.serialize] = false;
    outcome.detail = String(forced.error).slice(0, 200);
    return outcome;
  }

  // Stripping the captures makes the *element* serializers run; it does not
  // make a part serializer run. A repack copies `word/styles.xml`,
  // `word/numbering.xml` and the other declaration parts through byte for
  // byte, so the forced leg hands back the fixture unchanged and every pair in
  // them would read as surviving on the strength of a file copy. That is not a
  // survival and recording it as one would put a `modelled` disposition on a
  // slot no model holds. The test is the bytes themselves rather than a list of
  // parts, so a part folio starts rebuilding starts being measured with no
  // change here.
  if (forced.value === fixture.documentXml) {
    outcome.laws[SURVIVAL_LAWS.parse] = null;
    outcome.unrepresentable = `a repack replays ${fixture.part.path} verbatim, so removing the captures does not make its serializer run`;
    return outcome;
  }

  const forcedPresence = presenceIn(forced.value, fixture, expected);
  outcome.laws[SURVIVAL_LAWS.serialize] = forcedPresence === "equal";
  outcome.laws[SURVIVAL_LAWS.schema] = validateOoxmlPart({ graph, xml: forced.value }).length === 0;

  const editorPresence = await Result.tryPromise({
    try: async () => {
      const projected = fromProseDoc(toProseDoc(parsed.value), parsed.value);
      return presenceIn(
        await partOf(await save(withoutSerializerCaptures(projected)), fixture.part.path),
        fixture,
        expected,
      );
    },
    catch: (cause: unknown) => cause,
  });
  outcome.laws[SURVIVAL_LAWS.editor] = editorPresence.isOk() && editorPresence.value === "equal";
  if (editorPresence.isErr()) {
    outcome.detail = String(editorPresence.error).slice(0, 200);
  }

  const containerSpelling = spelledContainer(subject);
  const localName =
    subject.kind === "child" ? subject.slot.child.name : subject.slot.attribute.name;
  const trace = () => traceSubject(parsed.value, localName);
  outcome.mechanism = classify({
    forcedPresence,
    containerPresent:
      containerSpelling === undefined ||
      elementOccurrences(forced.value, containerSpelling).length > 0,
    replayedPresence: replayed.isOk() ? presenceIn(replayed.value, fixture, expected) : "absent",
    editorPresence: editorPresence.isOk() ? editorPresence.value : "absent",
    trace,
  });
  if (outcome.mechanism === null) {
    // A pair that survives with every verbatim slot cleared is held by the
    // typed model; one that stops surviving is held by bytes. Asking the
    // question by execution beats guessing from the shape of the model,
    // because a modelled slot rarely keeps the schema's name for itself.
    const modelOnly = await Result.tryPromise({
      try: async () =>
        presenceIn(
          await partOf(await save(withoutAnyVerbatimMarkup(parsed.value)), fixture.part.path),
          fixture,
          expected,
        ),
      catch: (cause: unknown) => cause,
    });
    if (modelOnly.isErr()) {
      outcome.carrier = "unknown";
    } else {
      outcome.carrier = modelOnly.value === "equal" ? "model" : "capture";
    }
  }
  return outcome;
};

/**
 * The container as the saved part spells it, or nothing when it is the root.
 *
 * A child's container is the element that declares it; an attribute's is the
 * element it sits on, which is the subject's own element.
 */
const spelledContainer = (subject: Subject): string | undefined => {
  const { namespace, name } = subject.slot.container.element;
  if (namespace === WML_NAMESPACE && (name === "document" || name === "body")) {
    return undefined;
  }
  return spell(subject.slot.container.element);
};

type Classification = {
  forcedPresence: Presence;
  containerPresent: boolean;
  replayedPresence: Presence;
  editorPresence: Presence;
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
  if (forcedPresence === "equal") {
    return editorPresence === "equal" ? null : LOSS_MECHANISMS.editorProjection;
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
