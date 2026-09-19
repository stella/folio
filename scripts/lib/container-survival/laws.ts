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
import { createEmptyDocx, repackDocx } from "@stll/folio-core/docx/rezip";
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
  /** Present only when the pair could not be tested at all. */
  unrepresentable: string | null;
  detail: string | null;
};

type Presence = "absent" | "equal" | "different";

const ON = new Set(["1", "true", "on"]);
const OFF = new Set(["0", "false", "off"]);

/**
 * Whether two spellings of a value mean the same thing.
 *
 * `ST_OnOff` has six spellings of two values and folio canonicalises them by
 * design; a measure has a Strict and a Transitional spelling and folio rewrites
 * every package as Transitional. Neither is a loss, and calling them one would
 * bury the losses that are.
 */
const sameValue = (written: string, read: string): boolean => {
  if (written === read) {
    return true;
  }
  if ((ON.has(written) && ON.has(read)) || (OFF.has(written) && OFF.has(read))) {
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
  const written = occurrences
    .map((attributes) => attributeIn(attributes, fixture.attributeSpelling ?? ""))
    .filter((value): value is string => value !== undefined);
  if (written.length === 0) {
    return "absent";
  }
  return written.some((value) => sameValue(expected, value)) ? "equal" : "different";
};

const CAPTURE_SLOT_NAMES = new Set([
  "sourceXml",
  "gridSourceXml",
  "verbatimXml",
  "rawPropertiesXml",
  "rawEndPropertiesXml",
  "rawXml",
  "rawWatermarkXml",
]);

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

const NUMBERING_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const NUMBERING_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

let basePackage: Promise<ArrayBuffer> | undefined;

const withNumbering = async (zip: JSZip): Promise<void> => {
  if (zip.file("word/numbering.xml") !== null) {
    return;
  }
  zip.file("word/numbering.xml", NUMBERING_PART);
  const types = await zip.file("[Content_Types].xml")?.async("text");
  const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
  if (types === undefined || rels === undefined) {
    throw new Error("the empty package lost its packaging parts");
  }
  zip.file(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      `<Override PartName="/word/numbering.xml" ContentType="${NUMBERING_CONTENT_TYPE}"/></Types>`,
    ),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      `<Relationship Id="rIdContainerSurvivalNumbering" Type="${NUMBERING_RELATIONSHIP}" Target="numbering.xml"/></Relationships>`,
    ),
  );
};

const packageFor = async (documentXml: string): Promise<ArrayBuffer> => {
  basePackage ??= createEmptyDocx();
  const zip = await JSZip.loadAsync(await basePackage);
  zip.file("word/document.xml", documentXml);
  await withNumbering(zip);
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentPartOf = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  return (await zip.file("word/document.xml")?.async("text")) ?? "";
};

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

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
  const buffer = await packageFor(fixture.documentXml);

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
    try: async () => documentPartOf(await save(parsed.value)),
    catch: (cause: unknown) => cause,
  });
  const forced = await Result.tryPromise({
    try: async () => documentPartOf(await save(withoutSerializerCaptures(parsed.value))),
    catch: (cause: unknown) => cause,
  });

  if (forced.isErr()) {
    outcome.laws[SURVIVAL_LAWS.serialize] = false;
    outcome.detail = String(forced.error).slice(0, 200);
    return outcome;
  }

  const forcedPresence = presenceIn(forced.value, fixture, expected);
  outcome.laws[SURVIVAL_LAWS.serialize] = forcedPresence === "equal";
  outcome.laws[SURVIVAL_LAWS.schema] = validateOoxmlPart({ graph, xml: forced.value }).length === 0;

  const editorPresence = await Result.tryPromise({
    try: async () => {
      const projected = fromProseDoc(toProseDoc(parsed.value), parsed.value);
      return presenceIn(
        await documentPartOf(await save(withoutSerializerCaptures(projected))),
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
  outcome.mechanism = classify({
    forcedPresence,
    containerPresent:
      containerSpelling === undefined ||
      elementOccurrences(forced.value, containerSpelling).length > 0,
    replayedPresence: replayed.isOk() ? presenceIn(replayed.value, fixture, expected) : "absent",
    editorPresence: editorPresence.isOk() ? editorPresence.value : "absent",
    trace: () =>
      traceSubject(
        parsed.value,
        subject.kind === "child" ? subject.slot.child.name : subject.slot.attribute.name,
      ),
  });
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
