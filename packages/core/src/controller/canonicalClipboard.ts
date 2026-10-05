import { Result } from "better-result";
import { Fragment, Slice } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { packageResourcesOpOf, type DocumentOp } from "@stll/docx-core/ops";
import {
  relationshipIdOf,
  paragraphNumberingReference,
  NUMBER_FORMATS,
} from "@stll/docx-core/model";
import { createNumberingIdAllocator } from "../docx/numberingIds";
import { visitInlineContentSlots, visitParagraphRuns } from "../docx/paragraphTraversal";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { completeNumberingForDoc } from "../prosemirror/listInstanceReferences";
import { marksToTextFormatting } from "../prosemirror/runFormattingFromMarks";
import type {
  Paragraph,
  Run,
  Document,
  NumberingDefinitions,
  CounterFormat,
  NumberFormat,
} from "../types/document";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import {
  CanonicalSessionError,
  type CanonicalCommit,
  type CanonicalSession,
} from "./canonicalSession";
import {
  flattenClipboardStyleReferences,
  importClipboardStyles,
} from "./canonicalClipboardResources";
import { bytesToDataUrl } from "../utils/base64";
import { captureVerbatimXml } from "../docx/verbatimCapture";
import {
  findAttributeByNamespaceUri,
  getLocalName,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  type XmlElement,
} from "../docx/xmlParser";

const refuse = (message: string) =>
  Result.err(
    new CanonicalSessionError({ gap: CANONICAL_GAP.dispatch, message, reason: "refused" }),
  );
const IMAGE_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const HYPERLINK_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const MAX_CLIPBOARD_IMAGE_BYTES = 16 * 1024 * 1024;

const isNumberFormat = (format: CounterFormat): format is NumberFormat =>
  NUMBER_FORMATS.some((value) => value === format);

const numberFormatFields = (format: CounterFormat) => {
  if (isNumberFormat(format)) return { numFmt: format };
  switch (format) {
    case "decimalZero3":
      return { numFmt: "custom", numFmtFormat: "000" } as const;
    case "decimalZero4":
      return { numFmt: "custom", numFmtFormat: "0000" } as const;
    case "decimalZero5":
      return { numFmt: "custom", numFmtFormat: "00000" } as const;
    default: {
      const exhaustive: never = format;
      return exhaustive;
    }
  }
};

type ImportNumberingOptions = {
  destination: NumberingDefinitions | undefined;
  source: NumberingDefinitions | undefined;
  paragraphs: Paragraph[];
  mode: "copy" | "move";
  sourceKind: "package" | "normalizedSlice";
  extraIds?: ReadonlySet<number>;
  styleIds?: ReadonlyMap<string, string>;
};

/** Clipboard instance ids belong to their source package, so imported lists get fresh ids. */
const importNumbering = ({
  destination,
  source,
  paragraphs,
  mode,
  sourceKind,
  extraIds,
  styleIds,
}: ImportNumberingOptions) => {
  if (mode === "move")
    return Result.ok({ numbering: destination, numIds: new Map<number, number>() });
  const references = paragraphs.flatMap((paragraph) => {
    const formats = [
      paragraph.formatting,
      ...(paragraph.propertyChanges ?? []).map((change) => change.previousFormatting),
    ];
    return formats.flatMap((formatting) =>
      formatting?.numPr?.kind === "reference" ? [formatting.numPr.numId] : [],
    );
  });
  references.push(...(extraIds ?? []));
  if (references.length === 0)
    return Result.ok({ numbering: destination, numIds: new Map<number, number>() });
  if (source === undefined) return refuse("The clipboard list has no numbering definition.");
  const sourceDefinitions = source;
  if (references.some((id) => !sourceDefinitions.nums.some(({ numId }) => numId === id)))
    return refuse("The clipboard list has no numbering definition.");
  const used = new Set(references);
  const sourceNums = source.nums.filter(({ numId }) => used.has(numId));
  const usedAbstracts = new Set(sourceNums.map(({ abstractNumId }) => abstractNumId));
  source = {
    ...source,
    nums: structuredClone(sourceNums),
    abstractNums: structuredClone(
      source.abstractNums.filter(({ abstractNumId }) => usedAbstracts.has(abstractNumId)),
    ),
  };
  if (sourceKind === "normalizedSlice") {
    for (const paragraph of paragraphs) {
      const rendering = paragraph.listRendering;
      if (rendering === undefined) continue;
      const instance = source.nums.find(({ numId }) => numId === rendering.numId);
      const definition = source.abstractNums.find(
        ({ abstractNumId }) => abstractNumId === instance?.abstractNumId,
      );
      if (definition === undefined) return refuse("The clipboard list has no abstract definition.");
      const level = definition.levels.find(({ ilvl }) => ilvl === rendering.level);
      if (level === undefined) return refuse("The clipboard list level could not be imported.");
      Object.assign(level, {
        ...numberFormatFields(rendering.numFmt ?? (rendering.isBullet ? "bullet" : "decimal")),
        lvlText: rendering.markerTemplate ?? rendering.marker,
        start: rendering.levelStarts?.at(rendering.level) ?? level.start ?? 1,
        ...(rendering.isLegal === undefined ? {} : { isLgl: rendering.isLegal }),
        ...(rendering.markerAlignment === undefined ? {} : { lvlJc: rendering.markerAlignment }),
        ...(rendering.markerSuffix === undefined ? {} : { suffix: rendering.markerSuffix }),
        ...(rendering.levelTabs === undefined
          ? {}
          : { pPr: { ...level.pPr, tabs: rendering.levelTabs } }),
        ...(rendering.markerFormatting === undefined ? {} : { rPr: rendering.markerFormatting }),
      });
      if (rendering.markerAllCaps !== undefined || rendering.markerHidden !== undefined) {
        level.rPr = {
          ...level.rPr,
          ...(rendering.markerAllCaps === undefined ? {} : { allCaps: rendering.markerAllCaps }),
          ...(rendering.markerHidden === undefined ? {} : { hidden: rendering.markerHidden }),
        };
      }
    }
  }
  const numAllocator = createNumberingIdAllocator(
    "num",
    (destination?.nums ?? []).map(({ numId }) => numId),
  );
  const abstractAllocator = createNumberingIdAllocator(
    "abstract",
    (destination?.abstractNums ?? []).map(({ abstractNumId }) => abstractNumId),
  );
  const nums = new Map(source.nums.map(({ numId }) => [numId, numAllocator.next()]));
  const abstracts = new Map(
    source.abstractNums.map(({ abstractNumId }) => [abstractNumId, abstractAllocator.next()]),
  );
  for (const paragraph of paragraphs) {
    const formats = [
      paragraph.formatting,
      ...(paragraph.propertyChanges ?? []).map((change) => change.previousFormatting),
    ];
    for (const formatting of formats) {
      if (formatting?.numPr?.kind !== "reference") continue;
      const numId = nums.get(formatting.numPr.numId);
      if (numId === undefined) return refuse("The clipboard list instance could not be imported.");
      formatting.numPr = paragraphNumberingReference({
        numId,
        ...(formatting.numPr.ilvl === undefined ? {} : { ilvl: formatting.numPr.ilvl }),
      });
    }
    if (paragraph.listRendering !== undefined) {
      const numId = nums.get(paragraph.listRendering.numId);
      if (numId !== undefined) paragraph.listRendering.numId = numId;
      const abstractId = paragraph.listRendering.abstractNumId;
      if (abstractId !== undefined) {
        const abstractNumId = abstracts.get(abstractId);
        if (abstractNumId !== undefined) paragraph.listRendering.abstractNumId = abstractNumId;
      }
    }
  }
  const importedNums = [];
  for (const instance of source.nums) {
    const numId = nums.get(instance.numId);
    const abstractNumId = abstracts.get(instance.abstractNumId);
    if (numId === undefined || abstractNumId === undefined)
      return refuse("The clipboard numbering definition could not be imported.");
    const imported = { ...instance, numId, abstractNumId };
    for (const override of imported.levelOverrides ?? []) {
      if (override.lvl?.pStyle === undefined || styleIds === undefined) continue;
      const mapped = styleIds.get(override.lvl.pStyle);
      if (mapped === undefined)
        return refuse("A clipboard numbering style dependency could not be imported.");
      override.lvl.pStyle = mapped;
    }
    importedNums.push(imported);
  }
  const importedAbstracts = [];
  for (const definition of source.abstractNums) {
    const abstractNumId = abstracts.get(definition.abstractNumId);
    if (abstractNumId === undefined)
      return refuse("The clipboard numbering definition could not be imported.");
    const imported = { ...definition, abstractNumId };
    if (styleIds !== undefined) {
      if (imported.numStyleLink !== undefined) {
        const mapped = styleIds.get(imported.numStyleLink);
        if (mapped === undefined)
          return refuse("A clipboard numbering style dependency could not be imported.");
        imported.numStyleLink = mapped;
      }
      if (imported.styleLink !== undefined) {
        const mapped = styleIds.get(imported.styleLink);
        if (mapped === undefined)
          return refuse("A clipboard numbering style dependency could not be imported.");
        imported.styleLink = mapped;
      }
      for (const level of imported.levels) {
        if (level.pStyle === undefined) continue;
        const mapped = styleIds.get(level.pStyle);
        if (mapped === undefined)
          return refuse("A clipboard numbering style dependency could not be imported.");
        level.pStyle = mapped;
      }
    }
    importedAbstracts.push(imported);
  }
  return Result.ok({
    numbering: {
      ...destination,
      nums: [...(destination?.nums ?? []), ...importedNums],
      abstractNums: [...(destination?.abstractNums ?? []), ...importedAbstracts],
    },
    numIds: nums,
  });
};

/** Clipboard slices are input payloads; their paragraph identities and private captures are not owners. */
type PrepareCanonicalPasteOptions = {
  session: CanonicalSession;
  state: EditorState;
  slice: Slice;
  plain?: boolean;
  moveTarget?: number;
  pasteTarget?: number;
  sourceDocument?: Document;
  moveSource?: { from: number; to: number };
};

export const prepareCanonicalPaste = ({
  session,
  state,
  slice,
  plain = false,
  moveTarget,
  pasteTarget,
  sourceDocument,
  moveSource,
}: PrepareCanonicalPasteOptions): Result<CanonicalCommit, CanonicalSessionError> => {
  // Declared: the inferred union of Ok/Err branches is emitted in a nondeterministic order.
  if (session.isComposing) return refuse("Composition must finish before using the clipboard.");
  if (slice.content.size === 0)
    return Result.err(
      new CanonicalSessionError({
        gap: CANONICAL_GAP.dispatch,
        message: "The clipboard contains no content.",
        reason: "noChange",
      }),
    );
  if (slice.openStart > 1 || slice.openEnd > 1)
    return refuse("Clipboard tables and nested block containers require canonical table editing.");
  const paragraphType = state.schema.nodes["paragraph"];
  if (paragraphType === undefined) return refuse("Clipboard input requires a paragraph schema.");
  let inline = true;
  let paragraphsOnly = true;
  slice.content.forEach((node) => {
    inline &&= node.isInline;
    paragraphsOnly &&= node.type === paragraphType;
  });
  if (!inline && !paragraphsOnly)
    return refuse("Clipboard tables and embedded blocks require canonical table editing.");
  // A same-parent document slice omits its paragraph; it still represents open inline edges.
  const importedSlice = inline
    ? new Slice(Fragment.from(paragraphType.create(null, slice.content)), 1, 1)
    : slice;
  const converted = Result.try({
    // Imported fragments mint new paragraph identities; matching source ids
    // must not attach a durable paragraph owner from either document.
    try: () => proseDocToBlocks(state.schema.topNodeType.create(null, importedSlice.content), []),
    catch: (cause) =>
      new CanonicalSessionError({
        gap: CANONICAL_GAP.dispatch,
        reason: "refused",
        message: `Clipboard normalization failed: ${cause instanceof Error ? cause.message : String(cause)}`,
      }),
  });
  if (converted.isErr()) return converted;
  const paragraphs: Paragraph[] = [];
  for (const block of converted.value) {
    if (block.type !== "paragraph") return refuse("The clipboard contains an unsupported block.");
    paragraphs.push(structuredClone(block));
  }
  if (moveTarget === undefined && (plain || sourceDocument === undefined))
    flattenClipboardStyleReferences(paragraphs);
  const from = session.projection.addressAt(
    pasteTarget ?? moveSource?.from ?? state.selection.from,
  );
  if (from.isErr()) return from;
  const to = session.projection.addressAt(pasteTarget ?? moveSource?.to ?? state.selection.to);
  if (to.isErr()) return to;
  let styles = session.document.package.styles;
  let styleIds: ReadonlyMap<string, string> | undefined;
  let extraIds: ReadonlySet<number> | undefined;
  if (sourceDocument !== undefined && moveTarget === undefined && !plain) {
    const importedStyles = importClipboardStyles({
      destination: session.document,
      source: sourceDocument,
      paragraphs,
    });
    if (importedStyles.isErr()) return refuse(importedStyles.error.message);
    styles = importedStyles.value.styles;
    paragraphs.splice(0, paragraphs.length, ...importedStyles.value.paragraphs);
    styleIds = importedStyles.value.styleIds;
    extraIds = importedStyles.value.numberingIds;
  }
  const sourceNumbering =
    sourceDocument?.package.numbering ??
    completeNumberingForDoc(
      undefined,
      state.schema.topNodeType.create(null, importedSlice.content),
    );
  const importedNumbering = importNumbering({
    destination: session.document.package.numbering,
    source: sourceNumbering,
    paragraphs,
    mode: moveTarget === undefined ? "copy" : "move",
    sourceKind: sourceDocument === undefined ? "normalizedSlice" : "package",
    ...(extraIds === undefined ? {} : { extraIds }),
    ...(styleIds === undefined ? {} : { styleIds }),
  });
  if (importedNumbering.isErr()) return importedNumbering;
  const numbering = importedNumbering.value.numbering;
  if (styles !== undefined && styleIds !== undefined) {
    const importedStyleIds = new Set(styleIds.values());
    for (const style of styles.styles) {
      if (!importedStyleIds.has(style.styleId) || style.pPr?.numPr?.kind !== "reference") continue;
      const numId = importedNumbering.value.numIds.get(style.pPr.numPr.numId);
      if (numId === undefined)
        return refuse("A clipboard style numbering dependency could not be imported.");
      style.pPr.numPr = paragraphNumberingReference({
        numId,
        ...(style.pPr.numPr.ilvl === undefined ? {} : { ilvl: style.pPr.numPr.ilvl }),
      });
    }
  }
  const relationships = new Map(session.document.package.relationships);
  const media = new Map(session.document.package.media);
  let importedResource = false;
  let imageIndex = 1;
  let relationshipIndex = 1;
  const importExternalLink = (href: string | undefined, previousId: string | undefined) => {
    const source =
      previousId === undefined ? undefined : sourceDocument?.package.relationships?.get(previousId);
    let target = href;
    if (target === undefined || target.length === 0)
      target = source?.targetMode === "External" ? source.target : undefined;
    if (target === undefined || target.length === 0)
      return refuse("The clipboard hyperlink has no resolved target.");
    while (relationships.has(`rId${relationshipIndex}`)) relationshipIndex += 1;
    const id = `rId${relationshipIndex++}`;
    relationships.set(id, { id, type: HYPERLINK_RELATIONSHIP, target, targetMode: "External" });
    importedResource = true;
    return Result.ok({ id, href: target });
  };
  const importCapturedLinks = (xml: string, remapped: Map<string, string>) => {
    const parsed = Result.try({
      try: () => parseXml(xml, OOXML_NAMESPACE_SCOPE),
      catch: () =>
        new CanonicalSessionError({
          gap: CANONICAL_GAP.dispatch,
          reason: "refused",
          message: "The clipboard hyperlink metadata is malformed.",
        }),
    });
    if (parsed.isErr()) return parsed;
    const visit = (element: XmlElement): Result<void, CanonicalSessionError> => {
      for (const name of Object.keys(element.attributes ?? {})) {
        const attribute = findAttributeByNamespaceUri(
          element,
          OFFICE_RELATIONSHIP_NAMESPACE_URIS,
          getLocalName(name),
        );
        if (attribute === null || attribute.name !== name) continue;
        if (getLocalName(name) !== "id")
          return refuse("Clipboard hyperlink metadata references an unsupported package part.");
        let id = remapped.get(attribute.value);
        if (id === undefined) {
          const imported = importExternalLink(undefined, attribute.value);
          if (imported.isErr()) return imported;
          id = imported.value.id;
          remapped.set(attribute.value, id);
        }
        if (element.attributes !== undefined) element.attributes[name] = id;
      }
      for (const child of element.elements ?? []) {
        const result = visit(child);
        if (result.isErr()) return result;
      }
      return Result.ok(undefined);
    };
    const imported = visit(parsed.value);
    if (imported.isErr()) return imported;
    return Result.ok((parsed.value.elements ?? []).map(captureVerbatimXml).join(""));
  };
  // New inline images have no opaque XML dependencies. Existing package images
  // retain their relationship only when its resolved source belongs to this package.
  for (const paragraph of paragraphs) {
    if (moveTarget === undefined) {
      let requiresStory = false;
      visitInlineContentSlots(paragraph, ({ item }) => {
        if (
          item.type === "commentRangeStart" ||
          item.type === "commentRangeEnd" ||
          item.type === "commentReference"
        )
          requiresStory = true;
      });
      visitParagraphRuns(paragraph, (run) => {
        if (run.content.some((item) => item.type === "footnoteRef" || item.type === "endnoteRef"))
          requiresStory = true;
      });
      if (requiresStory)
        return refuse(
          "Clipboard comments and note references require importing their source story parts.",
        );
      const links: Extract<Paragraph["content"][number], { type: "hyperlink" }>[] = [];
      visitInlineContentSlots(paragraph, ({ item }) => {
        if (item.type === "hyperlink") links.push(item);
      });
      for (const link of links) {
        if (link.anchor !== undefined) {
          delete link.rId;
          continue;
        }
        const imported = importExternalLink(link.href, link.rId);
        if (imported.isErr()) return imported;
        link.href = imported.value.href;
        link.rId = imported.value.id;
      }
    }
    const runs: Run[] = [];
    visitParagraphRuns(paragraph, (run) => runs.push(run));
    for (const item of runs) {
      if (plain)
        item.formatting = marksToTextFormatting(
          state.storedMarks ??
            state.doc.resolve(pasteTarget ?? moveSource?.from ?? state.selection.from).marks(),
        );
      for (const child of item.content) {
        if (child.type !== "drawing") continue;
        if (child.rawXmlMode !== undefined)
          return refuse("Opaque drawings and embedded objects require their source package parts.");
        delete child.rawXml;
        delete child.rawImageFingerprint;
        const image = child.image;
        delete image.id;
        if (
          moveTarget === undefined &&
          (image.hlinkRId !== undefined || image.hlinkHref !== undefined)
        ) {
          const previousId = image.hlinkRId;
          const imported = importExternalLink(image.hlinkHref, previousId);
          if (imported.isErr()) return imported;
          image.hlinkHref = imported.value.href;
          image.hlinkRId = imported.value.id;
          const remapped = new Map<string, string>();
          if (previousId !== undefined) remapped.set(previousId, imported.value.id);
          if (image.hlinkClickSource !== undefined) {
            const rebound = importCapturedLinks(image.hlinkClickSource.xml, remapped);
            if (rebound.isErr()) return rebound;
            image.hlinkClickSource = { xml: rebound.value, rId: imported.value.id };
          }
          if (image.hlinkHoverXml !== undefined) {
            const rebound = importCapturedLinks(image.hlinkHoverXml, remapped);
            if (rebound.isErr()) return rebound;
            image.hlinkHoverXml = rebound.value;
          }
        } else if (moveTarget === undefined && image.hlinkHoverXml !== undefined) {
          const rebound = importCapturedLinks(image.hlinkHoverXml, new Map());
          if (rebound.isErr()) return rebound;
          image.hlinkHoverXml = rebound.value;
        }
        const relationship = image.rId === undefined ? undefined : relationships.get(image.rId);
        const ownedPath = relationship?.target.startsWith("/")
          ? relationship.target.slice(1)
          : `word/${relationship?.target ?? ""}`;
        const ownedMedia = relationship === undefined ? undefined : media.get(ownedPath);
        if (
          ownedMedia !== undefined &&
          (moveTarget !== undefined ||
            (image.src !== undefined && ownedMedia.dataUrl === image.src))
        )
          continue;
        const sourceRelationship =
          image.rId === undefined
            ? undefined
            : sourceDocument?.package.relationships?.get(image.rId);
        const sourcePath = sourceRelationship?.target.startsWith("/")
          ? sourceRelationship.target.slice(1)
          : `word/${sourceRelationship?.target ?? ""}`;
        const sourceMedia = sourceDocument?.package.media?.get(sourcePath);
        if (sourceMedia !== undefined && sourceMedia.data.byteLength > MAX_CLIPBOARD_IMAGE_BYTES)
          return refuse("The clipboard image exceeds the import size limit.");
        const src =
          image.src ??
          (sourceMedia === undefined
            ? undefined
            : bytesToDataUrl(new Uint8Array(sourceMedia.data), sourceMedia.mimeType)) ??
          (sourceRelationship?.targetMode === "External" ? sourceRelationship.target : undefined);
        if (src === undefined) return refuse("The clipboard image has no source bytes or URL.");
        image.src = src;
        while (relationships.has(`rId${relationshipIndex}`)) relationshipIndex += 1;
        const id = `rId${relationshipIndex++}`;
        const data = /^data:(image\/(?:png|jpeg|gif|webp));base64,([A-Za-z0-9+/=\s]+)$/u.exec(src);
        if (data !== null) {
          const mimeType = data.at(1);
          const payload = data.at(2);
          if (mimeType === undefined || payload === undefined)
            return refuse("The clipboard image data is malformed.");
          if (payload.length > Math.ceil((MAX_CLIPBOARD_IMAGE_BYTES * 4) / 3))
            return refuse("The clipboard image exceeds the import size limit.");
          const decoded = Result.try({
            try: () => Uint8Array.from(atob(payload), (char) => char.charCodeAt(0)),
            catch: () =>
              new CanonicalSessionError({
                gap: CANONICAL_GAP.dispatch,
                message: "The clipboard image data is malformed.",
                reason: "refused",
              }),
          });
          if (decoded.isErr()) return decoded;
          const extension = mimeType === "image/jpeg" ? "jpg" : mimeType.slice(6);
          while (media.has(`word/media/clipboard${imageIndex}.${extension}`)) imageIndex += 1;
          const path = `word/media/clipboard${imageIndex++}.${extension}`;
          media.set(path, { path, mimeType, data: decoded.value.buffer, dataUrl: src });
          relationships.set(id, { id, type: IMAGE_RELATIONSHIP, target: path.slice(5) });
        } else if (/^https?:\/\//u.test(src)) {
          relationships.set(id, {
            id,
            type: IMAGE_RELATIONSHIP,
            target: src,
            targetMode: "External",
          });
        } else return refuse("Clipboard images require image bytes or an HTTP image URL.");
        const relationshipId = relationshipIdOf(id);
        if (relationshipId === undefined)
          return refuse("The clipboard image relationship id is invalid.");
        image.rId = relationshipId;
        // Drawing identities from another package must not claim local ownership.
        delete image.id;
        importedResource = true;
      }
    }
  }
  const resourceOps: DocumentOp[] = [];
  if (
    importedResource ||
    numbering !== session.document.package.numbering ||
    styles !== session.document.package.styles
  ) {
    const imported = {
      ...session.document,
      package: {
        ...session.document.package,
        ...(numbering === undefined ? {} : { numbering }),
        ...(styles === undefined ? {} : { styles }),
        ...(importedResource ? { relationships, media } : {}),
      },
    };
    resourceOps.push(packageResourcesOpOf({ before: session.document, after: imported }));
  }
  const fragment = {
    from: from.value,
    to: to.value,
    paragraphs,
    openStart: importedSlice.openStart === 0 ? 0 : 1,
    openEnd: importedSlice.openEnd === 0 ? 0 : 1,
  } as const;
  if (moveTarget === undefined)
    return session.prepareIntents(state, {
      intents: [{ type: "replaceFragment", ...fragment }],
      resourceOps,
      semantic: "paste",
    });
  const target = session.projection.addressAt(moveTarget);
  if (target.isErr()) return target;
  return session.prepareIntents(state, {
    intents: [{ type: "moveFragment", target: target.value, ...fragment }],
    resourceOps,
    semantic: "paste",
  });
};
