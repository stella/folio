/**
 * Apply a source document onto a template from a plan: the output is the
 * template's own package with a new main story. Parts the plan does not
 * touch keep their bytes; the ones it does are spliced, not re-serialized.
 */
import { Result } from "better-result";

import { ZipArchive } from "@stll/docx-core/zip";

import { registerContentTypeParts } from "../../contentTypeRegistry";
import { RELATIONSHIP_TYPES } from "../../relsParser";
import { getLocalName, NAMESPACES, type XmlElement } from "../../xmlParser";
import { appendToXmlRoot } from "../../xmlRootInsertion";
import {
  type AnalyzedDocument,
  type AnalyzedNotes,
  analyzeDocument,
} from "./blocks";
import {
  freezeFields,
  setUpdateFieldsOnOpen,
  withoutContained,
} from "./builders";
import {
  createRelationshipsEditor,
  type EmitContext,
  type EmitShared,
  type NoteKind,
  noteXml,
  paragraphContentXml,
  type ParagraphEmit,
  paragraphXml,
  type RelationshipsEditor,
  type SourceMedia,
  tableXml,
  templateCharacterStyles,
} from "./emit";
import { readHeaderFooterParts } from "./inspect";
import { styleNumbering } from "./labels";
import { guarded, readPartText, relationshipsPathFor } from "./package";
import {
  type ExpectedSegment,
  type PieceOrigin,
  type ResolvedParagraph,
  type ResolvedPlan,
  resolvePlan,
} from "./plan";
import { readStoryResources } from "./resources";
import type {
  ApplyTemplatePlanArgs,
  ApplyTemplatePlanResult,
  TemplatePlan,
  TemplatePlanNumberingPatch,
} from "./types";
import { TemplatePlanError } from "./types";
import { reportTemplatePlanOutput } from "./verify";
import {
  elementChildren,
  escapeXmlAttribute,
  innerRangeOf,
  integerOrNull,
  isW,
  maxOf,
  type PartPrefixes,
  partPrefixes,
  rangeOf,
  type RawPart,
  readRawPart,
  sliceWithSplices,
  spliceXml,
  walkElements,
  wAttr,
  wChild,
  wChildren,
  wName,
  type XmlSplice,
} from "./xml";

const CONTENT_TYPES_PATH = "[Content_Types].xml";
const WORDPROCESSINGML_CONTENT_TYPE_PREFIX =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.";
const RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const MEDIA_CONTENT_TYPES: Readonly<Record<string, string>> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
  svg: "image/svg+xml",
  webp: "image/webp",
};

const NOTE_PARTS = {
  footnote: {
    origin: "footnotes",
    fileName: "footnotes.xml",
    root: "footnotes",
    contentType: `${WORDPROCESSINGML_CONTENT_TYPE_PREFIX}footnotes+xml`,
    relationshipType: RELATIONSHIP_TYPES.footnotes,
  },
  endnote: {
    origin: "endnotes",
    fileName: "endnotes.xml",
    root: "endnotes",
    contentType: `${WORDPROCESSINGML_CONTENT_TYPE_PREFIX}endnotes+xml`,
    relationshipType: RELATIONSHIP_TYPES.endnotes,
  },
} as const satisfies Record<
  NoteKind,
  {
    origin: PieceOrigin;
    fileName: string;
    root: string;
    contentType: string;
    relationshipType: string;
  }
>;

const unreadable = (part: string, message: string): TemplatePlanError =>
  new TemplatePlanError({ code: "xml-unreadable", message, part });

const directoryOf = (path: string): string => {
  const slash = path.lastIndexOf("/");
  return slash === -1 ? "" : path.slice(0, slash + 1);
};

const RANGE_MARKERS = [
  ["bookmarkStart", "bookmarkEnd"],
  ["commentRangeStart", "commentRangeEnd"],
] as const;

type OrphanMarkerOptions = {
  kept: readonly XmlElement[];
  /** Regions of the template the output does not carry. */
  replaced: readonly { start: number; end: number }[];
};

/**
 * Splices removing range markers whose other half the output does not carry:
 * a bookmark or comment range that started in a block the plan dropped.
 */
const orphanMarkerSplices = ({
  kept,
  replaced,
}: OrphanMarkerOptions): XmlSplice[] => {
  const splices: XmlSplice[] = [];
  for (const [startName, endName] of RANGE_MARKERS) {
    const starts = new Map<string, XmlElement[]>();
    const ends = new Map<string, XmlElement[]>();
    for (const block of kept) {
      walkElements(block, (element) => {
        let side: Map<string, XmlElement[]> | null = null;
        if (isW(element, startName)) {
          side = starts;
        } else if (isW(element, endName)) {
          side = ends;
        }
        const range = rangeOf(element);
        const id = wAttr(element, "id");
        if (
          side === null ||
          range === null ||
          id === null ||
          replaced.some(
            ({ start, end }) => range.start >= start && range.end <= end,
          )
        ) {
          return;
        }
        side.set(id, [...(side.get(id) ?? []), element]);
      });
    }
    for (const [own, other] of [
      [starts, ends],
      [ends, starts],
    ] as const) {
      for (const [id, elements] of own) {
        if (other.has(id)) {
          continue;
        }
        for (const element of elements) {
          const range = rangeOf(element);
          if (range !== null) {
            splices.push({ ...range, newXml: "" });
          }
        }
      }
    }
  }
  return splices;
};

const LEVEL_CHILD_ORDER = [
  "start",
  "numFmt",
  "lvlRestart",
  "pStyle",
  "isLgl",
  "suff",
  "lvlText",
  "lvlPicBulletId",
  "legacy",
  "lvlJc",
  "pPr",
  "rPr",
] as const;

/** Children of `w:pPr` that follow `w:ind` in the schema's order. */
const AFTER_INDENT = [
  "contextualSpacing",
  "mirrorIndents",
  "suppressOverlap",
  "jc",
  "textDirection",
  "textAlignment",
  "textboxTightWrap",
  "outlineLvl",
  "divId",
  "cnfStyle",
  "rPr",
  "sectPr",
  "pPrChange",
] as const;

type LevelPropertiesOptions = {
  part: RawPart;
  level: XmlElement;
  w: string;
  indent: NonNullable<TemplatePlanNumberingPatch["indent"]>;
};

/**
 * The `w:pPr` of a level with the planned indent. What the level states and
 * the patch leaves out is kept; a hanging and a first-line indent exclude
 * each other, and an indent counted in characters would override the planned
 * one, so those are not carried over.
 */
const levelPropertiesXml = ({
  part,
  level,
  w,
  indent,
}: LevelPropertiesOptions): string | null => {
  const properties = wChild(level, "pPr");
  const existing = wChild(properties, "ind");
  const stated = (name: string): number | null =>
    integerOrNull(wAttr(existing, name));
  const hanging =
    indent.hanging ??
    (indent.firstLine === undefined ? stated("hanging") : null);
  const firstLine =
    indent.firstLine ??
    (indent.hanging === undefined ? stated("firstLine") : null);
  const values: [string, number | null][] = [
    ["left", indent.left ?? stated("left") ?? stated("start")],
    ["right", stated("right") ?? stated("end")],
    ["hanging", hanging],
    ["firstLine", firstLine],
  ];
  const attributes = values
    .flatMap(([name, value]) =>
      value === null ? [] : [` ${w}:${name}="${String(value)}"`],
    )
    .join("");
  const indentXml = `<${w}:ind${attributes}/>`;
  const range = properties === null ? null : rangeOf(properties);
  if (properties === null || range === null) {
    return `<${w}:pPr>${indentXml}</${w}:pPr>`;
  }
  const existingRange = existing === null ? null : rangeOf(existing);
  if (existingRange !== null) {
    return sliceWithSplices(part.xml, range, [
      { ...existingRange, newXml: indentXml },
    ]);
  }
  const inner = innerRangeOf(part.xml, properties);
  if (inner === null) {
    const name = properties.name ?? `${w}:pPr`;
    return `<${name}>${indentXml}</${name}>`;
  }
  const later = AFTER_INDENT.map((name) => wChild(properties, name)).find(
    (element) => element !== null,
  );
  const laterRange =
    later === undefined || later === null ? null : rangeOf(later);
  const at = laterRange?.start ?? inner.end;
  return `${part.xml.slice(range.start, at)}${indentXml}${part.xml.slice(at, range.end)}`;
};

type NumberingEdit = {
  patches: readonly TemplatePlanNumberingPatch[];
  mints: ResolvedPlan["mints"];
};

/** The template numbering part with level patches and new list instances. */
const editNumbering = (
  part: RawPart,
  { patches, mints }: NumberingEdit,
): Result<string, TemplatePlanError> => {
  const prefixes = partPrefixes(part.root);
  if (prefixes === null) {
    return Result.err(
      unreadable(
        part.path,
        "The numbering part binds no WordprocessingML prefix.",
      ),
    );
  }
  const { w } = prefixes;
  const splices: XmlSplice[] = [];
  const insertions = new Map<number, string[]>();
  const insert = (offset: number, xml: string): void => {
    insertions.set(offset, [...(insertions.get(offset) ?? []), xml]);
  };
  for (const patch of patches) {
    const abstract = wChildren(part.root, "abstractNum").find(
      (element) =>
        integerOrNull(wAttr(element, "abstractNumId")) === patch.abstractNumId,
    );
    const level = wChildren(abstract, "lvl").find(
      (element) => integerOrNull(wAttr(element, "ilvl")) === patch.ilvl,
    );
    const levelRange = level === undefined ? null : rangeOf(level);
    if (level === undefined || levelRange === null) {
      return Result.err(
        new TemplatePlanError({
          code: "numbering-unknown",
          message: `The numbering part has no level ${patch.ilvl} in list definition ${patch.abstractNumId}.`,
        }),
      );
    }
    // A level that wraps a child in a compatibility alternative has two
    // spellings of it; patching one would leave them disagreeing.
    if (elementChildren(level).some((child) => wName(child) === null)) {
      return Result.err(
        new TemplatePlanError({
          code: "unsupported-content",
          message: `Level ${patch.ilvl} of list definition ${patch.abstractNumId} holds markup a patch cannot change.`,
        }),
      );
    }
    const valued = (name: string, value: string): string =>
      `<${w}:${name} ${w}:val="${escapeXmlAttribute(value)}"/>`;
    const properties =
      patch.indent === undefined
        ? undefined
        : levelPropertiesXml({ part, level, w, indent: patch.indent });
    if (properties === null) {
      return Result.err(
        unreadable(part.path, "The numbering part could not be patched."),
      );
    }
    // Each planned child as the markup that states it.
    const planned: Partial<Record<(typeof LEVEL_CHILD_ORDER)[number], string>> =
      {
        ...(patch.start === undefined
          ? {}
          : { start: valued("start", String(patch.start)) }),
        ...(patch.numFmt === undefined
          ? {}
          : { numFmt: valued("numFmt", patch.numFmt) }),
        ...(patch.suff === undefined
          ? {}
          : { suff: valued("suff", patch.suff) }),
        ...(patch.lvlText === undefined
          ? {}
          : { lvlText: valued("lvlText", patch.lvlText) }),
        ...(properties === undefined ? {} : { pPr: properties }),
      };
    const inner = innerRangeOf(part.xml, level);
    for (const [index, name] of LEVEL_CHILD_ORDER.entries()) {
      const xml = planned[name];
      if (xml === undefined) {
        continue;
      }
      const existing = wChild(level, name);
      const existingRange = existing === null ? null : rangeOf(existing);
      if (existingRange !== null) {
        splices.push({ ...existingRange, newXml: xml });
        continue;
      }
      if (inner === null) {
        // `<w:lvl .../>`: open it; the insertions below land inside.
        insert(levelRange.end - 2, "");
        continue;
      }
      const later = LEVEL_CHILD_ORDER.slice(index + 1)
        .map((following) => wChild(level, following))
        .find((element) => element !== null);
      const laterRange =
        later === undefined || later === null ? null : rangeOf(later);
      insert(laterRange?.start ?? inner.end, xml);
    }
    if (inner === null) {
      const children = LEVEL_CHILD_ORDER.flatMap((name) => {
        const xml = planned[name];
        return xml === undefined ? [] : [xml];
      }).join("");
      insertions.delete(levelRange.end - 2);
      splices.push({
        start: levelRange.end - 2,
        end: levelRange.end,
        newXml: `>${children}</${level.name ?? `${w}:lvl`}>`,
      });
    }
  }
  if (mints.length > 0) {
    const nums = wChildren(part.root, "num");
    const markup = mints
      .map(({ numId, baseNumId, abstractNumId, starts }) => {
        // The new instance keeps what its base list overrides, level by level.
        const base = nums.find(
          (element) => integerOrNull(wAttr(element, "numId")) === baseNumId,
        );
        const inherited = new Map<number, XmlElement>();
        for (const override of wChildren(base, "lvlOverride")) {
          const level = integerOrNull(wAttr(override, "ilvl"));
          if (level !== null) {
            inherited.set(level, override);
          }
        }
        const raw = (element: XmlElement | null | undefined): string => {
          const range =
            element === null || element === undefined ? null : rangeOf(element);
          return range === null ? "" : part.xml.slice(range.start, range.end);
        };
        const overrides = [...new Set([...inherited.keys(), ...starts.keys()])]
          .toSorted((left, right) => left - right)
          .map((level) => {
            const start = starts.get(level);
            const kept = inherited.get(level);
            if (start === undefined) {
              return raw(kept);
            }
            const startOverride = (prefix: string): string =>
              `<${prefix}:startOverride ${prefix}:val="${start}"/>`;
            const keptRange = kept === undefined ? null : rangeOf(kept);
            const keptInner =
              kept === undefined ? null : innerRangeOf(part.xml, kept);
            const keptPrefix = kept?.name?.includes(":")
              ? kept.name.slice(0, kept.name.indexOf(":"))
              : null;
            if (
              kept === undefined ||
              keptRange === null ||
              keptInner === null ||
              keptPrefix === null
            ) {
              return `<${w}:lvlOverride ${w}:ilvl="${level}">${startOverride(w)}</${w}:lvlOverride>`;
            }
            // The restart goes into the base override as written, so a
            // prefix that override binds stays bound for the level it holds.
            const replaced = wChild(kept, "startOverride");
            const replacedRange = replaced === null ? null : rangeOf(replaced);
            const rest =
              replacedRange === null
                ? part.xml.slice(keptInner.start, keptInner.end)
                : part.xml.slice(keptInner.start, replacedRange.start) +
                  part.xml.slice(replacedRange.end, keptInner.end);
            return (
              part.xml.slice(keptRange.start, keptInner.start) +
              startOverride(keptPrefix) +
              rest +
              part.xml.slice(keptInner.end, keptRange.end)
            );
          })
          .join("");
        return `<${w}:num ${w}:numId="${numId}"><${w}:abstractNumId ${w}:val="${abstractNumId}"/>${overrides}</${w}:num>`;
      })
      .join("");
    const lastNum = nums.at(-1);
    const cleanup = wChild(part.root, "numIdMacAtCleanup");
    const inner = innerRangeOf(part.xml, part.root);
    const at =
      (lastNum === undefined ? null : rangeOf(lastNum)?.end) ??
      (cleanup === null ? null : rangeOf(cleanup)?.start) ??
      inner?.end;
    if (at === undefined) {
      return Result.err(
        unreadable(
          part.path,
          "The numbering part has no place for a new list.",
        ),
      );
    }
    insert(at, markup);
  }
  for (const [offset, xmls] of insertions) {
    splices.push({ start: offset, end: offset, newXml: xmls.join("") });
  }
  const edited = spliceXml(part.xml, splices);
  return edited === null
    ? Result.err(
        unreadable(part.path, "The numbering part could not be patched."),
      )
    : Result.ok(edited);
};

type NotePartState = {
  kind: NoteKind;
  /** The part the output keeps its notes in; created when the template has none. */
  path: string;
  existing: AnalyzedNotes;
  prefixes: PartPrefixes;
  relationships: RelationshipsEditor;
  nextId: number;
  /** Whether a part the template does not use for notes already has `path`. */
  pathTaken: boolean;
  /** New id by source id, in the order references were written. */
  allocated: Map<string, number>;
};

type WriteArgs = {
  source: AnalyzedDocument;
  template: AnalyzedDocument;
  plan: TemplatePlan;
  resolved: ResolvedPlan;
};

/** Write the output package for a resolved plan. */
const writeOutput = async ({
  source,
  template,
  plan,
  resolved,
}: WriteArgs): Promise<Result<Uint8Array, TemplatePlanError>> => {
  const { pkg } = template;
  const documentPrefixes = partPrefixes(template.document.root);
  const bodyInner = innerRangeOf(template.document.xml, template.body);
  if (documentPrefixes === null || bodyInner === null) {
    return Result.err(
      unreadable(pkg.documentPath, "The template body cannot be rewritten."),
    );
  }
  const documentDirectory = directoryOf(pkg.documentPath);

  // Everything read from either package happens before the first write.
  const parts = new Map<string, string | Uint8Array>();
  const readTemplateText = async (
    path: string | null,
  ): Promise<Result<string | null, TemplatePlanError>> =>
    await readPartText(pkg, path);

  const sourceResources = await readStoryResources(source);
  if (Result.isError(sourceResources)) {
    return sourceResources;
  }
  const sourceMedia = new Map<string, SourceMedia>();
  for (const [path, bytes] of sourceResources.value.media) {
    const dot = path.lastIndexOf(".");
    sourceMedia.set(path, {
      bytes,
      extension: dot === -1 ? "" : path.slice(dot + 1).toLowerCase(),
    });
  }

  const documentRelationshipsXml = await readTemplateText(
    pkg.documentRelationshipsPath,
  );
  if (Result.isError(documentRelationshipsXml)) {
    return documentRelationshipsXml;
  }
  const documentRelationships = createRelationshipsEditor(
    pkg.documentRelationshipsPath,
    documentRelationshipsXml.value,
  );
  if (documentRelationships === null) {
    return Result.err(
      unreadable(
        pkg.documentRelationshipsPath,
        "The template's relationships part is not a relationships part.",
      ),
    );
  }

  // A notes part the template lacks is written under the usual name, the one
  // the engine reads notes back from. Another part may already sit there.
  const takenPaths = new Set([...pkg.paths].map((path) => path.toLowerCase()));
  const notePartState = async (
    kind: NoteKind,
  ): Promise<Result<NotePartState, TemplatePlanError>> => {
    const existing =
      kind === "footnote" ? template.footnotes : template.endnotes;
    const path =
      existing.path ?? `${documentDirectory}${NOTE_PARTS[kind].fileName}`;
    const prefixes =
      existing.part === null
        ? documentPrefixes
        : partPrefixes(existing.part.root);
    if (prefixes === null) {
      return Result.err(
        unreadable(path, "The notes part binds no WordprocessingML prefix."),
      );
    }
    const relationshipsPath = relationshipsPathFor(path);
    const relationshipsXml = await readTemplateText(
      pkg.paths.has(relationshipsPath) ? relationshipsPath : null,
    );
    if (Result.isError(relationshipsXml)) {
      return relationshipsXml;
    }
    const relationships = createRelationshipsEditor(
      relationshipsPath,
      relationshipsXml.value,
    );
    if (relationships === null) {
      return Result.err(
        unreadable(
          relationshipsPath,
          "A notes relationships part is not a relationships part.",
        ),
      );
    }
    return Result.ok({
      kind,
      path,
      existing,
      prefixes,
      relationships,
      nextId: maxOf(existing.usedIds, 0) + 1,
      pathTaken: existing.path === null && takenPaths.has(path.toLowerCase()),
      allocated: new Map(),
    });
  };
  const footnoteState = await notePartState("footnote");
  if (Result.isError(footnoteState)) {
    return footnoteState;
  }
  const endnoteState = await notePartState("endnote");
  if (Result.isError(endnoteState)) {
    return endnoteState;
  }
  const noteParts: Record<NoteKind, NotePartState> = {
    footnote: footnoteState.value,
    endnote: endnoteState.value,
  };

  const media = new Map<string, { path: string; media: SourceMedia }>();
  const mediaPaths = new Set<string>();
  let unsupportedMedia: string | null = null;
  // Drawing ids are unique over the whole document, its notes and its
  // headers and footers; new ones start above every one in use.
  let drawingId = 0;
  const noteDrawingIds = (root: XmlElement): void => {
    walkElements(root, (element) => {
      if (getLocalName(element.name) !== "docPr") {
        return;
      }
      const id = integerOrNull(String(element.attributes?.["id"] ?? ""));
      if (id !== null && id > drawingId) {
        drawingId = id;
      }
    });
  };
  noteDrawingIds(template.document.root);
  for (const notes of [template.footnotes, template.endnotes]) {
    if (notes.part !== null) {
      noteDrawingIds(notes.part.root);
    }
  }
  const headersAndFooters = await readHeaderFooterParts(template);
  if (Result.isError(headersAndFooters)) {
    return headersAndFooters;
  }
  const headerFooterParts = headersAndFooters.value;
  for (const part of headerFooterParts) {
    noteDrawingIds(part.root);
  }
  const shared: EmitShared = {
    styles: templateCharacterStyles(template),
    sourceRelationships: sourceResources.value.relationships,
    sourceMedia,
    resolveSourceTarget: sourceResources.value.resolveTarget,
    noteId: (kind, sourceId) => {
      const state = noteParts[kind];
      const known = state.allocated.get(sourceId);
      if (known !== undefined) {
        return known;
      }
      const id = state.nextId;
      state.nextId += 1;
      state.allocated.set(sourceId, id);
      return id;
    },
    addMedia: (sourcePath, sourceItem) => {
      let entry = media.get(sourcePath);
      if (entry === undefined) {
        if (MEDIA_CONTENT_TYPES[sourceItem.extension] === undefined) {
          unsupportedMedia ??= sourcePath;
        }
        // A name neither the template nor an earlier picture of this run holds.
        let index = media.size + 1;
        let path = `${documentDirectory}media/placed${index}.${sourceItem.extension}`;
        while (pkg.paths.has(path) || mediaPaths.has(path)) {
          index += 1;
          path = `${documentDirectory}media/placed${index}.${sourceItem.extension}`;
        }
        mediaPaths.add(path);
        entry = { path, media: sourceItem };
        media.set(sourcePath, entry);
      }
      return `/${entry.path}`;
    },
    nextDrawingId: () => {
      drawingId += 1;
      return drawingId;
    },
  };
  const documentContext: EmitContext = {
    ...shared,
    prefixes: documentPrefixes,
    relationships: documentRelationships,
  };

  const instructions = new Map(
    resolved.references.map(({ reference, instruction }) => [
      reference,
      instruction,
    ]),
  );
  const instructionOf = (segment: ExpectedSegment): string | null =>
    segment.kind === "reference"
      ? (instructions.get(segment.reference) ?? null)
      : null;
  const emitOf = (paragraph: ResolvedParagraph): ParagraphEmit => ({
    paragraph,
    bookmark: resolved.bookmarks.get(paragraph),
    instructionOf,
    styleBindsList:
      styleNumbering(
        paragraph.styleId ?? template.defaultStyleId,
        template.styles,
        template.numbering,
      )?.kind === "reference",
  });

  // Splices over the template's document: frozen fields, filled slots, and
  // range markers left without their other half.
  const templateSplices: XmlSplice[] = [
    ...freezeFields(template.document, resolved.freezeKinds).splices,
  ];
  const replaced: { start: number; end: number }[] = [];
  const slots = [
    ...resolved.nestedSlots,
    ...resolved.items.flatMap((item) => (item.kind === "slot" ? [item] : [])),
  ];
  for (const { templateBlock, paragraph } of slots) {
    const content = paragraphContentXml(emitOf(paragraph), documentContext);
    if (Result.isError(content)) {
      return content;
    }
    const range = rangeOf(templateBlock.element);
    if (range === null) {
      return Result.err(
        unreadable(pkg.documentPath, "A slot paragraph has no source range."),
      );
    }
    const inner = innerRangeOf(template.document.xml, templateBlock.element);
    if (inner === null) {
      const name = templateBlock.element.name ?? `${documentPrefixes.w}:p`;
      templateSplices.push({
        start: range.end - 2,
        end: range.end,
        newXml: `>${content.value}</${name}>`,
      });
      continue;
    }
    const properties = wChild(templateBlock.element, "pPr");
    const start =
      (properties === null ? null : rangeOf(properties)?.end) ?? inner.start;
    templateSplices.push({ start, end: inner.end, newXml: content.value });
    replaced.push({ start, end: inner.end });
  }
  const bodySection = wChild(template.body, "sectPr");
  const keptElements = [
    ...resolved.keptTemplateBlocks.map(({ element }) => element),
    ...resolved.items.flatMap((item) =>
      item.kind === "slot" ? [item.templateBlock.element] : [],
    ),
    ...(bodySection === null ? [] : [bodySection]),
  ];
  templateSplices.push(
    ...orphanMarkerSplices({ kept: keptElements, replaced }),
  );
  const documentSplices = withoutContained(templateSplices);
  const keptXml = (element: XmlElement): Result<string, TemplatePlanError> => {
    const range = rangeOf(element);
    const xml =
      range === null
        ? null
        : sliceWithSplices(template.document.xml, range, documentSplices);
    return xml === null
      ? Result.err(
          unreadable(
            pkg.documentPath,
            "A template block could not be carried over.",
          ),
        )
      : Result.ok(xml);
  };

  const body: string[] = [];
  for (const item of resolved.items) {
    let xml: Result<string, TemplatePlanError>;
    switch (item.kind) {
      case "template":
        xml = keptXml(item.block.element);
        break;
      case "slot":
        xml = keptXml(item.templateBlock.element);
        break;
      case "paragraph":
        xml = paragraphXml(emitOf(item.paragraph), documentContext);
        break;
      case "table":
        xml = tableXml(
          {
            table: item.block.element,
            styleId: item.styleId,
            paragraphOf: (element) => {
              const paragraph = item.cells.get(element);
              return paragraph === undefined
                ? Result.ok("")
                : paragraphXml(emitOf(paragraph), documentContext);
            },
          },
          documentContext,
        );
        break;
      default:
        return item satisfies never;
    }
    if (Result.isError(xml)) {
      return xml;
    }
    body.push(xml.value);
  }
  if (bodySection !== null) {
    const section = keptXml(bodySection);
    if (Result.isError(section)) {
      return section;
    }
    body.push(section.value);
  }
  const documentXml = spliceXml(template.document.xml, [
    { start: bodyInner.start, end: bodyInner.end, newXml: body.join("") },
  ]);
  if (documentXml === null) {
    return Result.err(
      unreadable(pkg.documentPath, "The new body could not be written."),
    );
  }
  parts.set(pkg.documentPath, documentXml);

  // Notes the placed text refers to, appended to the template's note parts.
  const newContentTypes: { partName: string; contentType: string }[] = [];
  for (const kind of ["footnote", "endnote"] as const) {
    const state = noteParts[kind];
    const definition = NOTE_PARTS[kind];
    const sourceNotes =
      kind === "footnote" ? source.footnotes : source.endnotes;
    const context: EmitContext = {
      ...shared,
      prefixes: state.prefixes,
      relationships: state.relationships,
    };
    const { w } = state.prefixes;
    const markup: string[] = [];
    for (const [sourceId, id] of state.allocated) {
      const note = sourceNotes.byId.get(sourceId);
      const xml =
        note === undefined
          ? Result.ok(`<${w}:${kind} ${w}:id="${id}"><${w}:p/></${w}:${kind}>`)
          : noteXml(
              {
                kind,
                id,
                note,
                source,
                fields: sourceNotes.state.fields,
              },
              context,
            );
      if (Result.isError(xml)) {
        return xml;
      }
      markup.push(xml.value);
    }
    let partXml: string | null = null;
    if (state.existing.part !== null) {
      const frozen = spliceXml(
        state.existing.part.xml,
        freezeFields(state.existing.part, resolved.freezeKinds).splices,
      );
      if (frozen === null) {
        return Result.err(
          unreadable(state.path, "The notes part could not be rewritten."),
        );
      }
      const appended =
        markup.length === 0
          ? frozen
          : appendToXmlRoot(frozen, {
              rootName: state.existing.part.root.name ?? "",
              markup: markup.join(""),
            });
      partXml = appended === state.existing.part.xml ? null : appended;
    } else if (markup.length > 0) {
      if (state.pathTaken) {
        // Writing the notes there would replace that part's bytes.
        return Result.err(
          new TemplatePlanError({
            code: "unsupported-content",
            message: `The template keeps another part at ${state.path}, where its ${kind}s would go.`,
            part: state.path,
          }),
        );
      }
      const separators = [-1, 0]
        .map(
          (id, index) =>
            `<${w}:${kind} ${w}:type="${index === 0 ? "separator" : "continuationSeparator"}" ${w}:id="${id}"><${w}:p><${w}:r><${w}:${index === 0 ? "separator" : "continuationSeparator"}/></${w}:r></${w}:p></${w}:${kind}>`,
        )
        .join("");
      // The ids of the new part start after its two separators.
      partXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${w}:${definition.root} xmlns:${w}="${state.prefixes.wUri}" xmlns:${state.prefixes.r}="${RELATIONSHIPS_NAMESPACE}">${separators}${markup.join("")}</${w}:${definition.root}>`;
      documentRelationships.add({
        type: definition.relationshipType,
        target: definition.fileName,
        external: false,
      });
      newContentTypes.push({
        partName: `/${state.path}`,
        contentType: definition.contentType,
      });
    }
    if (partXml !== null) {
      parts.set(state.path, partXml);
    }
    const relationshipsXml = state.relationships.xml();
    if (relationshipsXml !== null) {
      parts.set(state.relationships.path, relationshipsXml);
    }
  }
  if (unsupportedMedia !== null) {
    return Result.err(
      new TemplatePlanError({
        code: "unsupported-content",
        message:
          "A picture of the source is in a format the output cannot declare.",
        part: unsupportedMedia,
      }),
    );
  }
  for (const { path, media: item } of media.values()) {
    parts.set(path, item.bytes);
    newContentTypes.push({
      partName: `/${path}`,
      contentType: MEDIA_CONTENT_TYPES[item.extension] ?? "",
    });
  }
  const documentRelationshipsOut = documentRelationships.xml();
  if (documentRelationshipsOut !== null) {
    parts.set(pkg.documentRelationshipsPath, documentRelationshipsOut);
  }
  if (newContentTypes.length > 0) {
    const contentTypes = await readTemplateText(CONTENT_TYPES_PATH);
    if (Result.isError(contentTypes)) {
      return contentTypes;
    }
    if (contentTypes.value === null) {
      return Result.err(
        new TemplatePlanError({
          code: "part-missing",
          message: "The template has no content-type registry.",
          part: CONTENT_TYPES_PATH,
        }),
      );
    }
    const typesRoot = readRawPart(CONTENT_TYPES_PATH, contentTypes.value)?.root;
    if (
      typesRoot === undefined ||
      getLocalName(typesRoot.name) !== "Types" ||
      typesRoot.namespaceUri !== NAMESPACES.ct
    ) {
      return Result.err(
        unreadable(
          CONTENT_TYPES_PATH,
          "The template's content-type registry is not one.",
        ),
      );
    }
    parts.set(
      CONTENT_TYPES_PATH,
      registerContentTypeParts(contentTypes.value, newContentTypes),
    );
  }

  const patches = plan.numberingPatches ?? [];
  if (patches.length > 0 || resolved.mints.length > 0) {
    const numberingPart =
      template.numberingPath === null || template.numberingXml === null
        ? null
        : readRawPart(template.numberingPath, template.numberingXml);
    if (numberingPart === null) {
      return Result.err(
        new TemplatePlanError({
          code: "numbering-unknown",
          message: "The template has no numbering part to patch.",
        }),
      );
    }
    const numbering = editNumbering(numberingPart, {
      patches,
      mints: resolved.mints,
    });
    if (Result.isError(numbering)) {
      return numbering;
    }
    parts.set(numberingPart.path, numbering.value);
  }

  if (resolved.freezeKinds.size > 0) {
    for (const part of headerFooterParts) {
      const { splices } = freezeFields(part, resolved.freezeKinds);
      if (splices.length === 0) {
        continue;
      }
      const frozen = spliceXml(part.xml, splices);
      if (frozen === null) {
        return Result.err(
          unreadable(part.path, "A header or footer could not be rewritten."),
        );
      }
      parts.set(part.path, frozen);
    }
  }

  const updateOnOpen = plan.fields?.updateOnOpen;
  if (updateOnOpen !== undefined) {
    const settingsPath = pkg.partPath("settings");
    const settingsXml = await readTemplateText(settingsPath);
    if (Result.isError(settingsXml)) {
      return settingsXml;
    }
    if (settingsPath === null || settingsXml.value === null) {
      if (updateOnOpen) {
        return Result.err(
          new TemplatePlanError({
            code: "part-missing",
            message:
              "The template has no settings part to set update-on-open in.",
          }),
        );
      }
    } else {
      const settings = setUpdateFieldsOnOpen(settingsXml.value, updateOnOpen);
      if (settings === null) {
        return Result.err(
          unreadable(settingsPath, "The settings part could not be updated."),
        );
      }
      if (settings !== settingsXml.value) {
        parts.set(settingsPath, settings);
      }
    }
  }

  return await Result.tryPromise({
    try: async () => {
      const zip = await ZipArchive.read(pkg.bytes);
      for (const [path, content] of parts) {
        const current = zip.entry(path);
        zip.set(
          path,
          content,
          current === undefined ? {} : { modified: current.modified },
        );
      }
      return await zip.generate();
    },
    catch: (cause) =>
      new TemplatePlanError({
        code: "archive-unreadable",
        message: "The output package could not be written.",
        cause,
      }),
  });
};

const applyUnguarded = async ({
  source,
  template,
  plan,
  options,
}: ApplyTemplatePlanArgs): Promise<
  Result<ApplyTemplatePlanResult, TemplatePlanError>
> => {
  const analyzedSource = await analyzeDocument(source);
  if (Result.isError(analyzedSource)) {
    return analyzedSource;
  }
  const analyzedTemplate = await analyzeDocument(template);
  if (Result.isError(analyzedTemplate)) {
    return analyzedTemplate;
  }
  const resolved = resolvePlan(
    analyzedSource.value,
    analyzedTemplate.value,
    plan,
  );
  if (Result.isError(resolved)) {
    return resolved;
  }
  const written = await writeOutput({
    source: analyzedSource.value,
    template: analyzedTemplate.value,
    plan,
    resolved: resolved.value,
  });
  if (Result.isError(written)) {
    return written;
  }
  const report = await reportTemplatePlanOutput({
    source: analyzedSource.value,
    template: analyzedTemplate.value,
    output: written.value,
    plan,
    resolved: resolved.value,
    options,
  });
  if (Result.isError(report)) {
    return report;
  }
  if (!report.value.ok) {
    return Result.err(
      new TemplatePlanError({
        code: "verification-failed",
        message: `The output failed verification: ${report.value.issues
          .slice(0, 3)
          .map(({ message }) => message)
          .join(" ")}`,
        report: report.value,
      }),
    );
  }
  return Result.ok({ docx: written.value, report: report.value });
};

/**
 * Apply `source` onto `template` as `plan` says and verify the result.
 *
 * The plan names blocks by id and pins them by hash; the text itself always
 * comes from the source. The output is refused, with the report attached,
 * unless every placed block reads back identical.
 */
export const applyTemplatePlan = async (
  input: ApplyTemplatePlanArgs,
): Promise<Result<ApplyTemplatePlanResult, TemplatePlanError>> =>
  await guarded(async () => await applyUnguarded(input));
