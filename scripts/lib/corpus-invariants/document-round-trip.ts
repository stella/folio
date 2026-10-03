/**
 * Tracked editor saves preserve untouched source blocks by immutable identity.
 * Hand-mutated or cloned models have no source replay guarantee.
 * Compare uncompressed part bytes; ZIP timestamps and compression are transport.
 */
import { EditorState } from "prosemirror-state";
import { fromProseDoc, toProseDoc } from "@stll/folio-core/prosemirror/conversion";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { getSourceReplayToken, repackDocx } from "@stll/folio-core/docx/rezip";
import type { Document } from "@stll/folio-core/types/document";
import { Result, TaggedError } from "better-result";
import JSZip from "jszip";

import {
  getChildElements,
  getLocalName,
  getNamespaceUri,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "../../../packages/core/src/docx/xmlParser";
import { failureFromAssertion, failureFromError } from "../corpus-signature";
import {
  type CorpusInvariantInput,
  type CorpusInvariantOutcome,
  EXTENDED_CORPUS_INVARIANTS,
  timeStage,
} from "./contract";
import { describePackageDifferences, differenceFailures } from "./model-equality";
import { generalizePartPath } from "./save-idempotence";

const INVARIANT = EXTENDED_CORPUS_INVARIANTS.documentRoundTrip;
export const DOCUMENT_ROUND_TRIP_INSERTION = "‸";

class DocumentRoundTripOracleError extends TaggedError("DocumentRoundTripOracleError")<{
  message: string;
}> {}

/** Include unknown parts and binary contents, even when the parser has no carrier. */
const readParts = async (buffer: ArrayBuffer): Promise<Map<string, Uint8Array>> => {
  const zip = await JSZip.loadAsync(buffer);
  return new Map(
    await Promise.all(
      Object.entries(zip.files)
        .filter(([, file]) => !file.dir)
        .map(async ([path, file]) => [path, await file.async("uint8array")] as const),
    ),
  );
};

/** No cap: fixing one loss must not reveal a previously hidden signature. */
export const documentPartDifferences = (
  before: ReadonlyMap<string, Uint8Array>,
  after: ReadonlyMap<string, Uint8Array>,
  touchedPart?: string,
): string[] => {
  const messages = new Set<string>();
  for (const path of new Set([...before.keys(), ...after.keys()])) {
    if (path === touchedPart) continue;
    const original = before.get(path);
    const saved = after.get(path);
    const name = generalizePartPath(path);
    if (original === undefined) messages.add(`a part was added: ${name}`);
    else if (saved === undefined) messages.add(`a part was removed: ${name}`);
    else if (
      original.length !== saved.length ||
      !original.every((byte, index) => byte === saved[index])
    )
      messages.add(`a part changed bytes: ${name}`);
  }
  return [...messages].sort();
};

const XML_TOKEN =
  /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(?:"[^"]*"|'[^']*'|[^'">])*>/gu;

/**
 * Locate the declared first direct body paragraph using namespace-aware XML,
 * then find its lexical span without reserializing any surrounding markup.
 * A failed correspondence is an oracle failure, never an empty projection.
 */
export const withoutFirstBodyParagraph = (xml: string): string => {
  const root = parseXmlDocument(xml);
  if (
    !root ||
    getLocalName(root.name) !== "document" ||
    !WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(root) ?? "")
  )
    throw new DocumentRoundTripOracleError({ message: "The main document root is absent." });
  const rootChildren = getChildElements(root);
  const bodyIndex = rootChildren.findIndex(
    (element) =>
      getLocalName(element.name) === "body" &&
      WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? ""),
  );
  const body = bodyIndex < 0 ? undefined : rootChildren.at(bodyIndex);
  const children = getChildElements(body);
  const targetIndex = children.findIndex(
    (element) =>
      getLocalName(element.name) === "p" &&
      WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? ""),
  );
  if (!body || targetIndex < 0)
    throw new DocumentRoundTripOracleError({ message: "The declared body paragraph is absent." });
  const stack: string[] = [];
  let bodyDepth = -1;
  let rootChildIndex = -1;
  let childIndex = -1;
  let start = -1;
  for (const token of xml.matchAll(XML_TOKEN)) {
    const tag = token[0];
    if (tag.startsWith("<!") || tag.startsWith("<?")) continue;
    const closing = tag.startsWith("</");
    const name = tag
      .slice(closing ? 2 : 1)
      .split(/[\s/>]/u)
      .at(0);
    if (!name) break;
    if (closing) {
      if (stack.pop() !== name) break;
      if (start >= 0 && stack.length === bodyDepth)
        return xml.slice(0, start) + xml.slice(token.index + tag.length);
      continue;
    }
    if (stack.length === 1) {
      rootChildIndex += 1;
      if (rootChildren.at(rootChildIndex)?.name !== name) break;
      if (rootChildIndex === bodyIndex) bodyDepth = 2;
    }
    if (stack.length === bodyDepth) {
      childIndex += 1;
      if (children.at(childIndex)?.name !== name) break;
      if (childIndex === targetIndex) {
        start = token.index;
        if (tag.endsWith("/>")) return xml.slice(0, start) + xml.slice(start + tag.length);
      }
    }
    if (!tag.endsWith("/>")) stack.push(name);
  }
  throw new DocumentRoundTripOracleError({
    message: "The declared paragraph has no corresponding XML span.",
  });
};

/** One deterministic text insertion, requiring no package-wide operation seeding. */
const canonicalEdit = (document: Document): Document | undefined => {
  const projected = toProseDoc(document);
  let insertion: number | undefined;
  projected.forEach((node, offset) => {
    if (insertion === undefined && node.type.name === "paragraph")
      insertion = offset + node.nodeSize - 1;
  });
  if (insertion === undefined) return undefined;
  const edited = EditorState.create({ doc: projected }).tr.insertText(
    DOCUMENT_ROUND_TRIP_INSERTION,
    insertion,
  ).doc;
  return fromProseDoc(edited, document);
};

const save = (document: Document): Promise<ArrayBuffer> => {
  const sourceReplay = getSourceReplayToken(document);
  return repackDocx(document, {
    updateModifiedDate: false,
    ...(sourceReplay === undefined ? {} : { sourceReplay }),
  });
};

export const runDocumentRoundTripInvariant = async ({
  parsed,
  buffer,
  documentPart,
}: CorpusInvariantInput): Promise<CorpusInvariantOutcome> => {
  const timings = {};
  const failures: CorpusInvariantOutcome["failures"] = [];
  const measured = await Result.tryPromise({
    try: async () => {
      const tracked = await timeStage(timings, "tracked-parse", () =>
        parseDocx(buffer, { preloadFonts: false, sourceReplay: "tracked" }),
      );
      const editorDocument = fromProseDoc(toProseDoc(tracked), tracked);
      const original = await timeStage(timings, "original-parts", () => readParts(buffer));
      const saved = await timeStage(timings, "no-edit-save", () => save(editorDocument));
      const [reparsed, savedParts] = await timeStage(timings, "no-edit-read", () =>
        Promise.all([parseDocx(saved, { preloadFonts: false }), readParts(saved)]),
      );
      failures.push(
        ...differenceFailures(
          INVARIANT,
          describePackageDifferences(parsed, reparsed),
          (message) => `no-edit save changed ${message}`,
        ),
        ...documentPartDifferences(original, savedParts).map((message) =>
          failureFromAssertion(INVARIANT, `no-edit save: ${message}`),
        ),
      );
      const edited = await timeStage(timings, "canonical-edit", () => canonicalEdit(tracked));
      if (!edited) return;
      const editedSave = await timeStage(timings, "edit-save", () => save(edited));
      const [editedParsed, editedParts] = await timeStage(timings, "edit-read", () =>
        Promise.all([parseDocx(editedSave, { preloadFonts: false }), readParts(editedSave)]),
      );
      failures.push(
        ...differenceFailures(
          INVARIANT,
          describePackageDifferences(edited, editedParsed),
          (message) => `canonical edit save changed ${message}`,
        ),
        ...documentPartDifferences(original, editedParts, documentPart).map((message) =>
          failureFromAssertion(INVARIANT, `canonical edit: ${message}`),
        ),
      );
      const sourceXml = original.get(documentPart);
      const editedXml = editedParts.get(documentPart);
      if (!sourceXml || !editedXml) {
        failures.push(failureFromAssertion(INVARIANT, "canonical edit lost the main part"));
        return;
      }
      const decoder = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true });
      if (
        withoutFirstBodyParagraph(decoder.decode(sourceXml)) !==
        withoutFirstBodyParagraph(decoder.decode(editedXml))
      )
        failures.push(
          failureFromAssertion(INVARIANT, "canonical edit changed XML outside its declared block"),
        );
    },
    catch: (cause: unknown) => cause,
  });
  if (measured.isErr()) failures.push(failureFromError(INVARIANT, measured.error));
  return { failures, timings };
};
