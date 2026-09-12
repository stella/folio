import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import { folioAIBlockIdStability } from "../../ai-edits/block-identity";
import {
  createFolioAIEditSnapshot,
  hashFolioAIBlockText,
  numberingReferenceKeysOf,
  storyTablesOf,
} from "../../ai-edits/snapshot";
import type {
  FolioAIBlock,
  FolioAIBlockTableLocation,
  FolioAIEditSnapshot,
} from "../../ai-edits/types";
import type { FolioDocumentStoryHandle } from "../../ai-edits/headless";
import type {
  FolioContentBlock,
  FolioContentContainerPathEntry,
  FolioContentPropertySet,
  FolioContentRun,
  FolioContentStructuralBoundary,
} from "../../compare/content-types";
import {
  projectAuthoredParagraphFormatting,
  resolveEffectiveParagraphPresentation,
  type ParagraphPresentationUnsupportedProperty,
  type TableParagraphPresentationProjection,
} from "../../style-engine";
import { createStyleEngine } from "../../style-engine/styleEngine";
import {
  createParagraphRunFormattingResolver,
  resolveEffectiveRunPresentation,
} from "../../style-engine/runPresentation";
import { createTableCellPresentationResolver } from "../../style-engine/tableParagraphPresentation";
import { tableOfContentsStyleLevel } from "../../utils/tableOfContentsStyle";
import { paragraphProjectionParaId } from "../../docx/paragraphPropertySource";
import { projectTableCellRowSpans } from "../../utils/tableRowSpanProjection";
import {
  footnoteToProseDoc,
  headerFooterToProseDoc,
  toProseDoc,
} from "../../prosemirror/conversion/toProseDoc";
import type {
  BlockContent,
  ComplexField,
  Document,
  Hyperlink,
  Image,
  InlineSdt,
  MathEquation,
  Paragraph,
  Run,
  RunContent,
  Shape,
  SimpleField,
  SdtProperties,
  Table,
  TextFormatting,
  TrackedRunContent,
} from "../../types/document";
import type { DocxAuthoredRun } from "./docx-program";
import {
  docxCanonicalPropertyValue,
  docxParagraphFormattingProperties,
  docxTextFormattingProperties,
} from "./docx-model-properties";
import {
  ownContentSnapshot,
  requireOwnedContentSnapshotBlocks,
  type OwnedContentSnapshot,
} from "../../compare/owned-content-snapshot";

const RESOLVED_DOCX_STORY_SNAPSHOT_BRAND: unique symbol = Symbol("resolved-docx-story-snapshot");
const RESOLVED_DOCX_SOURCE_OPERAND_BRAND: unique symbol = Symbol("resolved-docx-source-operand");

/**
 * A private point-in-time projection owned by Folio's DOCX adapter.
 *
 * Its closure-private payload binds canonical neutral content, serializer
 * authorship, operation anchors, and table templates to one exact immutable
 * ProseMirror document. Structural lookalikes cannot manufacture one.
 */
export type ResolvedDocxStorySnapshot = {
  readonly [RESOLVED_DOCX_STORY_SNAPSHOT_BRAND]: true;
};

/** One source block bound by identity to the live story capsule that issued it. */
export type ResolvedDocxSourceOperand = {
  readonly [RESOLVED_DOCX_SOURCE_OPERAND_BRAND]: true;
};

type ResolvedDocxSourceOperandPayload = {
  readonly snapshot: ResolvedDocxStorySnapshot;
  readonly block: FolioContentBlock;
};

type ResolvedDocxStoryPayload = {
  readonly story: FolioDocumentStoryHandle;
  readonly sourceDocument: PMNode;
  readonly operationSnapshot: FolioAIEditSnapshot;
  readonly contentSnapshot: OwnedContentSnapshot;
  readonly authoredRunsByBlockId: ReadonlyMap<string, ResolvedDocxAuthoredRunProjection>;
  readonly unsupportedFieldsByBlockId: ReadonlyMap<string, readonly string[]>;
  readonly sourceOperandsByBlockId: ReadonlyMap<string, ResolvedDocxSourceOperand>;
  readonly tableNodes: ReadonlyMap<number, PMNode>;
  readonly numberingReferenceKeys: readonly string[];
};

type ResolvedDocxAuthoredRunProjection =
  | { readonly status: "exact"; readonly runs: readonly DocxAuthoredRun[] }
  | { readonly status: "unsupported"; readonly reason: "live-text-mismatch" };

const payloadBySnapshot = new WeakMap<ResolvedDocxStorySnapshot, ResolvedDocxStoryPayload>();
const payloadBySourceOperand = new WeakMap<
  ResolvedDocxSourceOperand,
  ResolvedDocxSourceOperandPayload
>();

const payloadOf = (snapshot: ResolvedDocxStorySnapshot): ResolvedDocxStoryPayload =>
  payloadBySnapshot.get(snapshot) ??
  panic("A resolved DOCX story snapshot was not created by Folio");

const freezeRecursively = (value: unknown): void => {
  if (typeof value !== "object" || value === null || Object.isFrozen(value)) return;
  for (const child of Object.values(value)) freezeRecursively(child);
  Object.freeze(value);
};

const ownFormatting = (formatting: TextFormatting | undefined): Readonly<TextFormatting> => {
  const owned = structuredClone(formatting ?? {});
  freezeRecursively(owned);
  return owned;
};

const storyContent = (
  document: Document,
  story: FolioDocumentStoryHandle,
): readonly BlockContent[] | null => {
  switch (story.type) {
    case "main":
      return document.package.document.content;
    case "header":
      return document.package.headers?.get(story.relationshipId)?.content ?? null;
    case "footer":
      return document.package.footers?.get(story.relationshipId)?.content ?? null;
    case "footnote":
      return document.package.footnotes?.find(({ id }) => id === story.noteId)?.content ?? null;
    case "endnote":
      return document.package.endnotes?.find(({ id }) => id === story.noteId)?.content ?? null;
    default: {
      const exhaustive: never = story;
      return exhaustive;
    }
  }
};

const projectStoryDocument = (
  document: Document,
  story: FolioDocumentStoryHandle,
  content: readonly BlockContent[],
): PMNode => {
  const conversionOptions = {
    ...(document.package.styles !== undefined && { styles: document.package.styles }),
    ...(document.package.theme !== undefined && { theme: document.package.theme }),
  };
  switch (story.type) {
    case "main":
      return toProseDoc(document, conversionOptions);
    case "header":
    case "footer":
      return headerFooterToProseDoc([...content], conversionOptions);
    case "footnote":
    case "endnote":
      return footnoteToProseDoc([...content], conversionOptions);
    default: {
      const exhaustive: never = story;
      return exhaustive;
    }
  }
};

const ownOperationSnapshot = (sourceDocument: PMNode): FolioAIEditSnapshot => {
  const snapshot = createFolioAIEditSnapshot(sourceDocument);
  freezeRecursively(snapshot);
  const numberingReferenceKeys = numberingReferenceKeysOf(snapshot);
  Object.freeze(numberingReferenceKeys);
  const storyTables = storyTablesOf(snapshot);
  for (const table of storyTables) Object.freeze(table);
  Object.freeze(storyTables);
  return snapshot;
};

type LiveAuthoredTextRun = {
  text: string;
  formatting: Readonly<TextFormatting>;
  effectiveFormatting: Readonly<TextFormatting>;
};

type EffectiveRunFormattingResolver = (
  formatting: TextFormatting | undefined,
  fieldType?: string,
) => Readonly<TextFormatting>;

type InlineProjection = {
  readonly runs: LiveAuthoredTextRun[];
  readonly structuralBoundaries: FolioContentStructuralBoundary[];
  readonly structure: unknown[];
  readonly textBoxes: Shape[];
  readonly unsupportedFields: Set<string>;
  textLength: number;
};

type InlineFieldDisposition = "nested" | "resource" | "semantic" | "transport" | "unsupported";
type InlineFieldDescriptor<Field extends string> = {
  readonly field: Field;
  readonly disposition: InlineFieldDisposition;
};
type TotalInlineFieldDescriptors<Value> = {
  readonly [Field in keyof Value]-?: InlineFieldDescriptor<Extract<Field, string>>;
};

const HYPERLINK_FIELD_DESCRIPTORS = Object.freeze({
  type: Object.freeze({ field: "type", disposition: "semantic" }),
  rId: Object.freeze({ field: "rId", disposition: "transport" }),
  href: Object.freeze({ field: "href", disposition: "semantic" }),
  anchor: Object.freeze({ field: "anchor", disposition: "semantic" }),
  tooltip: Object.freeze({ field: "tooltip", disposition: "semantic" }),
  target: Object.freeze({ field: "target", disposition: "semantic" }),
  history: Object.freeze({ field: "history", disposition: "semantic" }),
  docLocation: Object.freeze({ field: "docLocation", disposition: "semantic" }),
  children: Object.freeze({ field: "children", disposition: "nested" }),
} as const satisfies TotalInlineFieldDescriptors<Hyperlink>);

const SIMPLE_FIELD_DESCRIPTORS = Object.freeze({
  type: Object.freeze({ field: "type", disposition: "semantic" }),
  instruction: Object.freeze({ field: "instruction", disposition: "semantic" }),
  fieldType: Object.freeze({ field: "fieldType", disposition: "semantic" }),
  content: Object.freeze({ field: "content", disposition: "nested" }),
  fldLock: Object.freeze({ field: "fldLock", disposition: "semantic" }),
  dirty: Object.freeze({ field: "dirty", disposition: "semantic" }),
} as const satisfies TotalInlineFieldDescriptors<SimpleField>);

const COMPLEX_FIELD_DESCRIPTORS = Object.freeze({
  type: Object.freeze({ field: "type", disposition: "semantic" }),
  instruction: Object.freeze({ field: "instruction", disposition: "semantic" }),
  fieldType: Object.freeze({ field: "fieldType", disposition: "semantic" }),
  fieldCode: Object.freeze({ field: "fieldCode", disposition: "unsupported" }),
  fieldResult: Object.freeze({ field: "fieldResult", disposition: "nested" }),
  formatting: Object.freeze({ field: "formatting", disposition: "semantic" }),
  fldLock: Object.freeze({ field: "fldLock", disposition: "semantic" }),
  dirty: Object.freeze({ field: "dirty", disposition: "semantic" }),
} as const satisfies TotalInlineFieldDescriptors<ComplexField>);

const IMAGE_FIELD_DESCRIPTORS = Object.freeze({
  type: Object.freeze({ field: "type", disposition: "semantic" }),
  id: Object.freeze({ field: "id", disposition: "transport" }),
  rId: Object.freeze({ field: "rId", disposition: "transport" }),
  src: Object.freeze({ field: "src", disposition: "resource" }),
  mimeType: Object.freeze({ field: "mimeType", disposition: "semantic" }),
  filename: Object.freeze({ field: "filename", disposition: "semantic" }),
  docPrName: Object.freeze({ field: "docPrName", disposition: "semantic" }),
  alt: Object.freeze({ field: "alt", disposition: "semantic" }),
  title: Object.freeze({ field: "title", disposition: "semantic" }),
  size: Object.freeze({ field: "size", disposition: "semantic" }),
  originalSize: Object.freeze({ field: "originalSize", disposition: "semantic" }),
  wrap: Object.freeze({ field: "wrap", disposition: "semantic" }),
  position: Object.freeze({ field: "position", disposition: "semantic" }),
  transform: Object.freeze({ field: "transform", disposition: "semantic" }),
  padding: Object.freeze({ field: "padding", disposition: "semantic" }),
  crop: Object.freeze({ field: "crop", disposition: "semantic" }),
  opacity: Object.freeze({ field: "opacity", disposition: "semantic" }),
  layoutInCell: Object.freeze({ field: "layoutInCell", disposition: "semantic" }),
  allowOverlap: Object.freeze({ field: "allowOverlap", disposition: "semantic" }),
  decorative: Object.freeze({ field: "decorative", disposition: "semantic" }),
  hlinkHref: Object.freeze({ field: "hlinkHref", disposition: "semantic" }),
  hlinkRId: Object.freeze({ field: "hlinkRId", disposition: "transport" }),
  outline: Object.freeze({ field: "outline", disposition: "semantic" }),
  effects: Object.freeze({ field: "effects", disposition: "semantic" }),
} as const satisfies TotalInlineFieldDescriptors<Image>);

const SHAPE_FIELD_DESCRIPTORS = Object.freeze({
  type: Object.freeze({ field: "type", disposition: "semantic" }),
  shapeType: Object.freeze({ field: "shapeType", disposition: "semantic" }),
  geometryAdjustments: Object.freeze({ field: "geometryAdjustments", disposition: "semantic" }),
  id: Object.freeze({ field: "id", disposition: "transport" }),
  name: Object.freeze({ field: "name", disposition: "semantic" }),
  size: Object.freeze({ field: "size", disposition: "semantic" }),
  position: Object.freeze({ field: "position", disposition: "semantic" }),
  wrap: Object.freeze({ field: "wrap", disposition: "semantic" }),
  fill: Object.freeze({ field: "fill", disposition: "semantic" }),
  outline: Object.freeze({ field: "outline", disposition: "semantic" }),
  transform: Object.freeze({ field: "transform", disposition: "semantic" }),
  textBody: Object.freeze({ field: "textBody", disposition: "nested" }),
  customGeometry: Object.freeze({ field: "customGeometry", disposition: "semantic" }),
} as const satisfies TotalInlineFieldDescriptors<Shape>);

const SDT_PROPERTY_FIELD_DESCRIPTORS = Object.freeze({
  sdtType: Object.freeze({ field: "sdtType", disposition: "semantic" }),
  id: Object.freeze({ field: "id", disposition: "transport" }),
  alias: Object.freeze({ field: "alias", disposition: "semantic" }),
  tag: Object.freeze({ field: "tag", disposition: "semantic" }),
  lock: Object.freeze({ field: "lock", disposition: "semantic" }),
  placeholder: Object.freeze({ field: "placeholder", disposition: "semantic" }),
  showingPlaceholder: Object.freeze({ field: "showingPlaceholder", disposition: "semantic" }),
  dateFormat: Object.freeze({ field: "dateFormat", disposition: "semantic" }),
  dateValueISO: Object.freeze({ field: "dateValueISO", disposition: "semantic" }),
  listItems: Object.freeze({ field: "listItems", disposition: "semantic" }),
  dropdownLastValue: Object.freeze({ field: "dropdownLastValue", disposition: "semantic" }),
  checked: Object.freeze({ field: "checked", disposition: "semantic" }),
  rawPropertiesXml: Object.freeze({ field: "rawPropertiesXml", disposition: "unsupported" }),
  rawEndPropertiesXml: Object.freeze({ field: "rawEndPropertiesXml", disposition: "unsupported" }),
  rawSdtChildrenBeforeContent: Object.freeze({
    field: "rawSdtChildrenBeforeContent",
    disposition: "unsupported",
  }),
  rawSdtChildrenAfterContent: Object.freeze({
    field: "rawSdtChildrenAfterContent",
    disposition: "unsupported",
  }),
} as const satisfies TotalInlineFieldDescriptors<SdtProperties>);

const MATH_FIELD_DESCRIPTORS = Object.freeze({
  type: Object.freeze({ field: "type", disposition: "semantic" }),
  display: Object.freeze({ field: "display", disposition: "semantic" }),
  ommlXml: Object.freeze({ field: "ommlXml", disposition: "unsupported" }),
  plainText: Object.freeze({ field: "plainText", disposition: "semantic" }),
} as const satisfies TotalInlineFieldDescriptors<MathEquation>);

const descriptorSemanticProjection = (
  value: object,
  descriptors: Readonly<Record<string, InlineFieldDescriptor<string>>>,
): unknown => {
  const projected: Record<string, unknown> = {};
  for (const { field, disposition } of Object.values(descriptors)) {
    if (disposition === "transport" || disposition === "nested") continue;
    const fieldValue = Reflect.get(value, field);
    if (fieldValue === undefined) continue;
    projected[field] =
      (disposition === "resource" || disposition === "unsupported") &&
      typeof fieldValue === "string"
        ? {
            codeUnits: fieldValue.length,
            digest: hashFolioAIBlockText(fieldValue),
          }
        : fieldValue;
  }
  return projected;
};

const samePropertySet = (left: FolioContentPropertySet, right: FolioContentPropertySet): boolean =>
  JSON.stringify(left) === JSON.stringify(right);

const appendRunText = (
  projection: InlineProjection,
  text: string,
  formatting: TextFormatting | undefined,
  effectiveFormatting: Readonly<TextFormatting>,
): void => {
  if (text.length === 0) return;
  const owned = ownFormatting(formatting);
  const previous = projection.runs.at(-1);
  if (
    previous &&
    samePropertySet(
      docxTextFormattingProperties(previous.formatting),
      docxTextFormattingProperties(owned),
    ) &&
    samePropertySet(
      docxTextFormattingProperties(previous.effectiveFormatting),
      docxTextFormattingProperties(effectiveFormatting),
    )
  ) {
    previous.text += text;
  } else {
    projection.runs.push({ text, formatting: owned, effectiveFormatting });
  }
  projection.textLength += text.length;
};

const formattingSignature = (formatting: TextFormatting | undefined): unknown =>
  docxTextFormattingProperties(formatting);

const runHasPageBreak = (run: Run): boolean =>
  run.content.some((content) => content.type === "break" && content.breakType === "page");

const fieldHasStructuredProjection = (field: SimpleField | ComplexField): boolean =>
  field.type === "simpleField"
    ? field.content.some((content) =>
        content.type === "hyperlink" ? true : runHasPageBreak(content),
      )
    : field.fieldResult.some(runHasPageBreak);

const visitRunContent = (
  content: RunContent,
  formatting: TextFormatting | undefined,
  effectiveFormatting: Readonly<TextFormatting>,
  projection: InlineProjection,
): void => {
  switch (content.type) {
    case "text":
      appendRunText(projection, content.text, formatting, effectiveFormatting);
      return;
    case "softHyphen":
      appendRunText(projection, "\u00ad", formatting, effectiveFormatting);
      return;
    case "noBreakHyphen":
      appendRunText(projection, "\u2011", formatting, effectiveFormatting);
      return;
    case "footnoteRef":
    case "endnoteRef":
      appendRunText(projection, String(content.id), formatting, effectiveFormatting);
      return;
    case "break":
      if (content.breakType === "page") {
        projection.structuralBoundaries.push({
          type: "pageBreak",
          offset: projection.textLength,
          ...(content.clear !== undefined && { clear: content.clear }),
        });
        return;
      }
      projection.structure.push({
        type: "break",
        breakType: content.breakType ?? null,
        clear: content.clear ?? null,
        formatting: formattingSignature(formatting),
      });
      return;
    case "tab":
      projection.structure.push({
        type: "tab",
        positional: content.positional ?? null,
        offset: projection.textLength,
        formatting: formattingSignature(formatting),
      });
      return;
    case "symbol":
      projection.structure.push({
        type: "symbol",
        font: content.font,
        char: content.char,
        offset: projection.textLength,
        formatting: formattingSignature(formatting),
      });
      return;
    case "drawing":
      projection.unsupportedFields.add("inline.drawing");
      projection.structure.push({
        type: "drawing",
        presentation: descriptorSemanticProjection(content.image, IMAGE_FIELD_DESCRIPTORS),
        offset: projection.textLength,
      });
      return;
    case "shape":
      projection.unsupportedFields.add("inline.shape");
      if (content.shape.textBody !== undefined) {
        projection.textBoxes.push(content.shape);
        projection.structure.push({
          type: "textBoxAnchor",
          presentation: descriptorSemanticProjection(content.shape, SHAPE_FIELD_DESCRIPTORS),
          offset: projection.textLength,
        });
        return;
      }
      projection.structure.push({
        type: "shape",
        presentation: descriptorSemanticProjection(content.shape, SHAPE_FIELD_DESCRIPTORS),
        offset: projection.textLength,
      });
      return;
    case "fieldChar":
      projection.structure.push({
        type: "fieldChar",
        charType: content.charType,
        locked: content.fldLock ?? null,
        dirty: content.dirty ?? null,
        offset: projection.textLength,
      });
      return;
    case "instrText":
      projection.structure.push({
        type: "instructionText",
        text: content.text,
        offset: projection.textLength,
      });
      return;
    case "renderedPageBreak":
      // Cached pagination is not authored document content.
      return;
    default: {
      const exhaustive: never = content;
      return exhaustive;
    }
  }
};

const visitRun = (
  run: Run,
  projection: InlineProjection,
  resolveEffective: EffectiveRunFormattingResolver,
  fieldType?: string,
): void => {
  if (run.formatting?.hidden === true) {
    projection.unsupportedFields.add("inline.hiddenRun");
    projection.structure.push({
      type: "hiddenRun",
      formatting: formattingSignature(run.formatting),
      offset: projection.textLength,
    });
    return;
  }
  const effective = resolveEffective(run.formatting, fieldType);
  for (const content of run.content) {
    visitRunContent(content, run.formatting, effective, projection);
  }
};

const visitHyperlink = (
  hyperlink: Hyperlink,
  projection: InlineProjection,
  resolveEffective: EffectiveRunFormattingResolver,
  fieldType?: string,
): void => {
  projection.structure.push({
    type: "hyperlink",
    presentation: descriptorSemanticProjection(hyperlink, HYPERLINK_FIELD_DESCRIPTORS),
    offset: projection.textLength,
  });
  if (hyperlink.rId !== undefined && hyperlink.href === undefined) {
    projection.unsupportedFields.add("inline.hyperlinkResource");
  }
  for (const child of hyperlink.children) {
    if (child.type === "run") {
      visitRun(child, projection, resolveEffective, fieldType);
    } else {
      projection.unsupportedFields.add("inline.bookmark");
      projection.structure.push({
        type: child.type,
        ...(child.type === "bookmarkStart" && { name: child.name }),
        offset: projection.textLength,
      });
    }
  }
};

const visitField = (
  field: SimpleField | ComplexField,
  projection: InlineProjection,
  resolveEffective: EffectiveRunFormattingResolver,
): void => {
  const structured = fieldHasStructuredProjection(field);
  projection.unsupportedFields.add("inline.field");
  projection.structure.push({
    type: field.type,
    presentation: descriptorSemanticProjection(
      field,
      field.type === "simpleField" ? SIMPLE_FIELD_DESCRIPTORS : COMPLEX_FIELD_DESCRIPTORS,
    ),
    structured,
    offset: projection.textLength,
  });
  if (!structured) return;
  if (field.type === "complexField") {
    for (const run of field.fieldResult) {
      visitRun(run, projection, resolveEffective, field.fieldType);
    }
    return;
  }
  for (const content of field.content) {
    if (content.type === "run") {
      visitRun(content, projection, resolveEffective, field.fieldType);
    } else {
      visitHyperlink(content, projection, resolveEffective, field.fieldType);
    }
  }
};

const sdtSignature = ({ properties }: InlineSdt): unknown =>
  descriptorSemanticProjection(properties, SDT_PROPERTY_FIELD_DESCRIPTORS);

const visitTrackedContent = (
  content: TrackedRunContent,
  projection: InlineProjection,
  resolveEffective: EffectiveRunFormattingResolver,
  fieldType?: string,
): void => {
  switch (content.type) {
    case "run":
      visitRun(content, projection, resolveEffective, fieldType);
      return;
    case "hyperlink":
      visitHyperlink(content, projection, resolveEffective, fieldType);
      return;
    case "simpleField":
    case "complexField":
      visitField(content, projection, resolveEffective);
      return;
    case "bookmarkStart":
    case "bookmarkEnd":
      projection.structure.push({
        type: content.type,
        id: content.id,
        ...(content.type === "bookmarkStart" && { name: content.name }),
        offset: projection.textLength,
      });
      return;
    case "insertion":
    case "moveTo":
      for (const child of content.content) {
        visitTrackedContent(child, projection, resolveEffective, fieldType);
      }
      return;
    case "deletion":
    case "moveFrom":
      projection.unsupportedFields.add(`inline.${content.type}`);
      projection.structure.push({
        type: content.type,
        offset: projection.textLength,
      });
      return;
    default: {
      const exhaustive: never = content;
      return exhaustive;
    }
  }
};

const visitInlineSdt = (
  sdt: InlineSdt,
  projection: InlineProjection,
  resolveEffective: EffectiveRunFormattingResolver,
  fieldType?: string,
): void => {
  projection.unsupportedFields.add("inline.contentControl");
  projection.structure.push({
    type: "inlineSdt",
    properties: sdtSignature(sdt),
    offset: projection.textLength,
  });
  for (const content of sdt.content) {
    switch (content.type) {
      case "run":
        visitRun(content, projection, resolveEffective, fieldType);
        break;
      case "hyperlink":
        visitHyperlink(content, projection, resolveEffective, fieldType);
        break;
      case "simpleField":
      case "complexField":
        visitField(content, projection, resolveEffective);
        break;
      case "inlineSdt":
        visitInlineSdt(content, projection, resolveEffective, fieldType);
        break;
      case "insertion":
      case "moveTo":
        for (const child of content.content) {
          visitTrackedContent(child, projection, resolveEffective, fieldType);
        }
        break;
      case "deletion":
      case "moveFrom":
        projection.unsupportedFields.add(`inline.${content.type}`);
        projection.structure.push({
          type: content.type,
          offset: projection.textLength,
        });
        break;
      case "mathEquation":
        projection.unsupportedFields.add("inline.math");
        projection.structure.push({
          type: "math",
          presentation: descriptorSemanticProjection(content, MATH_FIELD_DESCRIPTORS),
          offset: projection.textLength,
        });
        break;
      default: {
        const exhaustive: never = content;
        return exhaustive;
      }
    }
  }
};

const projectParagraphInline = (
  paragraph: Paragraph,
  resolveEffective: EffectiveRunFormattingResolver,
): InlineProjection => {
  const projection: InlineProjection = {
    runs: [],
    structuralBoundaries: [],
    structure: [],
    textBoxes: [],
    unsupportedFields: new Set(),
    textLength: 0,
  };
  for (const content of paragraph.content) {
    switch (content.type) {
      case "run":
        visitRun(content, projection, resolveEffective);
        break;
      case "hyperlink":
        visitHyperlink(content, projection, resolveEffective);
        break;
      case "simpleField":
      case "complexField":
        visitField(content, projection, resolveEffective);
        break;
      case "inlineSdt":
        visitInlineSdt(content, projection, resolveEffective);
        break;
      case "insertion":
      case "moveTo":
        for (const child of content.content) {
          visitTrackedContent(child, projection, resolveEffective);
        }
        break;
      case "deletion":
      case "moveFrom":
        projection.unsupportedFields.add(`inline.${content.type}`);
        projection.structure.push({
          type: content.type,
          offset: projection.textLength,
        });
        break;
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "commentReference":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        projection.unsupportedFields.add(`inline.${content.type}`);
        projection.structure.push({
          type: content.type,
          ...(content.type === "bookmarkStart" ||
          content.type === "moveFromRangeStart" ||
          content.type === "moveToRangeStart"
            ? { name: content.name }
            : {}),
          offset: projection.textLength,
        });
        break;
      case "mathEquation":
        projection.unsupportedFields.add("inline.math");
        projection.structure.push({
          type: "math",
          presentation: descriptorSemanticProjection(content, MATH_FIELD_DESCRIPTORS),
          offset: projection.textLength,
        });
        break;
      default: {
        const exhaustive: never = content;
        return exhaustive;
      }
    }
  }
  return projection;
};

const sameStructuralBoundaries = (
  left: readonly FolioContentStructuralBoundary[],
  right: readonly FolioContentStructuralBoundary[],
): boolean => JSON.stringify(left) === JSON.stringify(right);

const contentRuns = (
  text: string,
  authored: readonly LiveAuthoredTextRun[],
): {
  readonly content: readonly FolioContentRun[];
  readonly authored: ResolvedDocxAuthoredRunProjection;
} => {
  const authoredIsExact = authored.map(({ text: runText }) => runText).join("") === text;
  const authoredRuns = authoredIsExact
    ? authored
    : [
        {
          text,
          formatting: ownFormatting(undefined),
          effectiveFormatting: ownFormatting(undefined),
        },
      ];
  const boundaries = new Set<number>([0, text.length]);
  let offset = 0;
  for (const run of authoredRuns) {
    offset += run.text.length;
    boundaries.add(offset);
  }
  const ordered = [...boundaries].toSorted((left, right) => left - right);
  const content: FolioContentRun[] = [];
  const authoredRanges: DocxAuthoredRun[] = [];
  let authoredIndex = 0;
  let authoredEnd = authoredRuns.at(0)?.text.length ?? 0;
  for (let index = 0; index + 1 < ordered.length; index++) {
    const start = ordered[index] ?? 0;
    const end = ordered[index + 1] ?? start;
    while (start >= authoredEnd && authoredIndex + 1 < authoredRuns.length) {
      authoredIndex++;
      authoredEnd += authoredRuns[authoredIndex]?.text.length ?? 0;
    }
    const authoredFormatting = authoredRuns[authoredIndex]?.formatting ?? ownFormatting(undefined);
    const effectiveFormatting = authoredRuns[authoredIndex]?.effectiveFormatting;
    content.push({
      text: text.slice(start, end),
      authoredFormatting: docxTextFormattingProperties(authoredFormatting),
      effectiveFormatting: docxTextFormattingProperties(effectiveFormatting),
    });
    authoredRanges.push(
      Object.freeze({ startOffset: start, endOffset: end, formatting: authoredFormatting }),
    );
  }
  if (text.length === 0) {
    authoredRanges.push(
      Object.freeze({ startOffset: 0, endOffset: 0, formatting: ownFormatting(undefined) }),
    );
  }
  return {
    content,
    authored: authoredIsExact
      ? { status: "exact", runs: Object.freeze(authoredRanges) }
      : { status: "unsupported", reason: "live-text-mismatch" },
  };
};

const paragraphHeadingLevel = (
  paragraph: Paragraph,
  effective: ReturnType<typeof resolveEffectiveParagraphPresentation>["effective"],
): number | undefined => {
  if (
    effective.outlineLevel !== undefined &&
    Number.isInteger(effective.outlineLevel) &&
    effective.outlineLevel >= 0 &&
    effective.outlineLevel <= 8
  ) {
    return effective.outlineLevel + 1;
  }
  const match = /^heading(?<level>[1-9])$/iu.exec(paragraph.formatting?.styleId ?? "");
  const level = match?.groups?.["level"];
  return level === undefined ? undefined : Number.parseInt(level, 10);
};

const unsupportedPresentationProperty = (
  unsupported: readonly ParagraphPresentationUnsupportedProperty[],
): FolioContentPropertySet[number] | null =>
  unsupported.length === 0
    ? null
    : {
        key: "docx.unsupportedParagraphPresentation",
        value: docxCanonicalPropertyValue(
          unsupported.map(({ source, field, value }) => ({ source, field, value })),
        ),
      };

type LiveTableOwnership = {
  readonly outerTableIndex: number;
  readonly tableIndex: number;
  readonly rowIndex: number;
  readonly cellIndex: number;
  readonly gridColumnIndex: number;
  readonly columnSpan: number;
  readonly rowSpan: number;
  readonly paragraphIndex: number;
};

type WalkContext = {
  readonly containerPath: readonly FolioContentContainerPathEntry[];
  readonly containerTopology: readonly number[];
  readonly table?: LiveTableOwnership;
  readonly tableParagraphPresentation?: TableParagraphPresentationProjection;
  readonly tableRunFormatting?: TextFormatting;
};

type ProjectionBuilder = {
  readonly blockById: ReadonlyMap<string, FolioAIBlock>;
  readonly projected: FolioContentBlock[];
  readonly authoredRunsByBlockId: Map<string, ResolvedDocxAuthoredRunProjection>;
  readonly unsupportedFieldsByBlockId: Map<string, readonly string[]>;
  readonly consumedBlockIds: Set<string>;
  readonly styleEngine: ReturnType<typeof createStyleEngine>;
  nextTableIndex: number;
};

const liveTableMatches = (
  expected: FolioAIBlockTableLocation | undefined,
  actual: LiveTableOwnership | undefined,
): boolean =>
  expected === undefined || actual === undefined
    ? expected === actual
    : expected.outerTableIndex === actual.outerTableIndex &&
      expected.tableIndex === actual.tableIndex &&
      expected.rowIndex === actual.rowIndex &&
      expected.cellIndex === actual.cellIndex &&
      expected.gridColumnIndex === actual.gridColumnIndex &&
      expected.columnSpan === actual.columnSpan &&
      expected.rowSpan === actual.rowSpan &&
      expected.paragraphIndex === actual.paragraphIndex;

const projectLiveParagraph = (
  paragraph: Paragraph,
  context: WalkContext,
  builder: ProjectionBuilder,
): readonly Shape[] => {
  const id = paragraphProjectionParaId(paragraph);
  const styleId = paragraph.formatting?.styleId;
  const styleName = styleId ? builder.styleEngine.getStyle(styleId)?.name : undefined;
  const runFormatting = createParagraphRunFormattingResolver({
    paragraph,
    styleResolver: builder.styleEngine,
    ...(context.tableRunFormatting !== undefined && {
      extraRunFormatting: context.tableRunFormatting,
    }),
    isTocParagraph:
      tableOfContentsStyleLevel({ styleId, ...(styleName ? { styleName } : {}) }) !== undefined,
  });
  const inline = projectParagraphInline(paragraph, (formatting, fieldType) =>
    ownFormatting(
      resolveEffectiveRunPresentation(
        formatting,
        runFormatting.resolve(formatting, fieldType),
        builder.styleEngine,
      ).effective,
    ),
  );
  if (id === undefined || !builder.blockById.has(id)) return inline.textBoxes;
  if (builder.consumedBlockIds.has(id)) {
    return panic("The live DOCX story repeats a projected paragraph identity", { blockId: id });
  }
  const operationBlock = builder.blockById.get(id);
  if (!operationBlock) return panic("A checked live paragraph lost its operation block");
  if (!liveTableMatches(operationBlock.table, context.table)) {
    return panic(
      "The live DOCX table ownership disagrees with its exact PM projection: " +
        `${JSON.stringify(context.table)} against ${JSON.stringify(operationBlock.table)}`,
      { blockId: id },
    );
  }
  if (
    !sameStructuralBoundaries(
      inline.structuralBoundaries,
      operationBlock.structuralBoundaries ?? [],
    )
  ) {
    return panic("The live DOCX inline structure disagrees with its exact PM projection", {
      blockId: id,
    });
  }
  const authored = projectAuthoredParagraphFormatting(paragraph.formatting);
  const presentation = resolveEffectiveParagraphPresentation({
    authored: paragraph.formatting,
    styleResolver: builder.styleEngine,
    ...(context.tableParagraphPresentation !== undefined && {
      tableParagraphPresentation: context.tableParagraphPresentation,
    }),
  });
  const headingLevel = paragraphHeadingLevel(paragraph, presentation.effective);
  const blockProperties: FolioContentPropertySet[number][] = [];
  if (headingLevel !== undefined) {
    blockProperties.push({ key: "headingLevel", value: headingLevel });
  }
  if (inline.structure.length > 0) {
    blockProperties.push({
      key: "docx.inlineStructure",
      value: JSON.stringify(inline.structure),
    });
  }
  if (inline.runs.map(({ text }) => text).join("") !== operationBlock.text) {
    blockProperties.push({
      key: "docx.authoredRunProjection",
      value: JSON.stringify(
        inline.runs.map(({ text, formatting, effectiveFormatting }) => ({
          text,
          formatting: formattingSignature(formatting),
          effectiveFormatting: formattingSignature(effectiveFormatting),
        })),
      ),
    });
  }
  const unsupported = unsupportedPresentationProperty(presentation.unsupported);
  if (unsupported) blockProperties.push(unsupported);
  const runs = contentRuns(operationBlock.text, inline.runs);
  const unsupportedFields = new Set(inline.unsupportedFields);
  for (const { source, field } of presentation.unsupported) {
    unsupportedFields.add(`paragraph.${source}.${field}`);
  }
  builder.projected.push({
    identity: {
      type: folioAIBlockIdStability(operationBlock) === "stable" ? "persistent-hint" : "positional",
      id,
    },
    // Every DOCX text block is structurally a paragraph. Heading and list
    // presentation live in their canonical property projections, so a
    // property edit cannot also manufacture a contradictory kind change.
    kind: "paragraph",
    text: operationBlock.text,
    blockProperties: blockProperties.toSorted((left, right) => {
      if (left.key < right.key) return -1;
      if (left.key > right.key) return 1;
      return 0;
    }),
    paragraphFormatting: {
      authored: docxParagraphFormattingProperties(authored),
      effective: docxParagraphFormattingProperties(presentation.effective),
    },
    runs: runs.content,
    structuralBoundaries: inline.structuralBoundaries,
    ...(context.table !== undefined && {
      table: {
        outerTableIdentity: {
          type: "positional",
          id: `outer-table-${String(context.table.outerTableIndex)}`,
        },
        tableIdentity: {
          type: "positional",
          id: `table-${String(context.table.tableIndex)}`,
        },
        rowIdentity: {
          type: "positional",
          id: `table-${String(context.table.tableIndex)}-row-${String(context.table.rowIndex)}`,
        },
        cellIdentity: {
          type: "positional",
          id: `table-${String(context.table.tableIndex)}-row-${String(context.table.rowIndex)}-cell-${String(context.table.cellIndex)}`,
        },
        ...context.table,
      },
    }),
    containerPath: context.containerPath,
  });
  builder.authoredRunsByBlockId.set(id, runs.authored);
  builder.unsupportedFieldsByBlockId.set(id, Object.freeze([...unsupportedFields].toSorted()));
  builder.consumedBlockIds.add(id);
  return inline.textBoxes;
};

const nextContainer = (
  kind: string,
  topology: readonly number[],
): FolioContentContainerPathEntry => {
  return {
    kind,
    identity: { type: "positional", id: topology.join(".") },
  };
};

const visitBlocks = (
  blocks: readonly BlockContent[],
  context: WalkContext,
  builder: ProjectionBuilder,
): void => {
  let nextContainerOrdinal = 0;
  for (const [blockIndex, block] of blocks.entries()) {
    if (block.type === "blockSdt") {
      const topology = [...context.containerTopology, nextContainerOrdinal++];
      const entry = nextContainer("blockSdt", topology);
      visitBlocks(
        block.content,
        {
          ...context,
          containerPath: [...context.containerPath, entry],
          containerTopology: topology,
        },
        builder,
      );
      continue;
    }
    if (block.type === "table") {
      visitTable(block, context, builder);
      continue;
    }
    const paragraphContext = context.table
      ? {
          ...context,
          table: { ...context.table, paragraphIndex: blockIndex },
        }
      : context;
    const textBoxes = projectLiveParagraph(block, paragraphContext, builder);
    for (const shape of textBoxes) {
      const textBody = shape.textBody;
      if (!textBody) continue;
      const topology = [...context.containerTopology, nextContainerOrdinal++];
      const entry = nextContainer("textBox", topology);
      visitBlocks(
        textBody.content,
        {
          containerPath: [...context.containerPath, entry],
          containerTopology: topology,
          ...(context.table !== undefined && { table: context.table }),
        },
        builder,
      );
    }
  }
};

const parentTableCellContainer = (table: LiveTableOwnership): FolioContentContainerPathEntry => ({
  kind: "tableCell",
  identity: {
    type: "positional",
    id: `table-${String(table.tableIndex)}-row-${String(table.rowIndex)}-cell-${String(table.cellIndex)}`,
  },
});

const visitTable = (table: Table, context: WalkContext, builder: ProjectionBuilder): void => {
  const tableIndex = builder.nextTableIndex++;
  const outerTableIndex = context.table?.outerTableIndex ?? tableIndex;
  const containerPath = context.table
    ? [...context.containerPath, parentTableCellContainer(context.table)]
    : context.containerPath;
  const resolveTableCellPresentation = createTableCellPresentationResolver({
    table,
    styleResolver: builder.styleEngine,
  });
  const rowSpans = projectTableCellRowSpans(table);
  for (const [rowIndex, row] of table.rows.entries()) {
    if (row.formatting?.hidden === true) continue;
    let gridColumnIndex = row.formatting?.gridBefore ?? 0;
    let projectedCellIndex = 0;
    for (const [sourceCellIndex, cell] of row.cells.entries()) {
      const columnSpan = cell.formatting?.gridSpan ?? 1;
      const rowSpan = rowSpans.get(`${String(rowIndex)}-${String(gridColumnIndex)}`);
      if (rowSpan?.skip === true) {
        gridColumnIndex += columnSpan;
        continue;
      }
      const tablePresentation = resolveTableCellPresentation({
        row,
        rowIndex,
        cell,
        cellIndex: sourceCellIndex,
      });
      visitBlocks(
        cell.content,
        {
          containerPath,
          containerTopology: context.containerTopology,
          table: {
            outerTableIndex,
            tableIndex,
            rowIndex,
            cellIndex: projectedCellIndex,
            gridColumnIndex,
            columnSpan,
            rowSpan: rowSpan?.rowSpan ?? 1,
            paragraphIndex: 0,
          },
          ...(tablePresentation?.paragraph !== undefined && {
            tableParagraphPresentation: tablePresentation.paragraph,
          }),
          ...(tablePresentation?.runFormatting !== undefined && {
            tableRunFormatting: tablePresentation.runFormatting,
          }),
        },
        builder,
      );
      projectedCellIndex++;
      gridColumnIndex += columnSpan;
    }
  }
};

type CreateResolvedDocxStorySnapshotOptions = {
  readonly document: Document;
  readonly story: FolioDocumentStoryHandle;
  readonly sourceDocument: PMNode;
};

/** @internal Capture one live reviewed story exactly once for all compare stages. */
export const createResolvedDocxStorySnapshot = ({
  document,
  story,
  sourceDocument,
}: CreateResolvedDocxStorySnapshotOptions): ResolvedDocxStorySnapshot | null => {
  const content = storyContent(document, story);
  if (content === null) return null;
  const projectedDocument = projectStoryDocument(document, story, content);
  if (!projectedDocument.eq(sourceDocument)) {
    return panic("The live DOCX story and its package projection disagree on source identity", {
      story,
    });
  }
  const operationSnapshot = ownOperationSnapshot(projectedDocument);
  const blockById = new Map(operationSnapshot.blocks.map((block) => [block.id, block]));
  const builder: ProjectionBuilder = {
    blockById,
    projected: [],
    authoredRunsByBlockId: new Map(),
    unsupportedFieldsByBlockId: new Map(),
    consumedBlockIds: new Set(),
    styleEngine: createStyleEngine(document.package.styles),
    nextTableIndex: 0,
  };
  visitBlocks(content, { containerPath: [], containerTopology: [] }, builder);
  if (
    builder.projected.length !== operationSnapshot.blocks.length ||
    operationSnapshot.blocks.some(
      (block, index) => builder.projected[index]?.identity.id !== block.id,
    )
  ) {
    return panic("The live DOCX story and its exact PM projection disagree on block ownership", {
      story,
      liveBlocks: builder.projected.length,
      operationBlocks: operationSnapshot.blocks.length,
    });
  }
  const contentSnapshot = ownContentSnapshot(builder.projected);
  const tableNodes = new Map(
    storyTablesOf(operationSnapshot).map(({ index, node }) => [index, node] as const),
  );
  const snapshot = Object.freeze({
    [RESOLVED_DOCX_STORY_SNAPSHOT_BRAND]: true as const,
  });
  const sourceOperandsByBlockId = new Map<string, ResolvedDocxSourceOperand>();
  for (const block of requireOwnedContentSnapshotBlocks(contentSnapshot)) {
    const source = Object.freeze({
      [RESOLVED_DOCX_SOURCE_OPERAND_BRAND]: true as const,
    });
    payloadBySourceOperand.set(source, { snapshot, block });
    sourceOperandsByBlockId.set(block.identity.id, source);
  }
  payloadBySnapshot.set(snapshot, {
    story: Object.freeze({ ...story }),
    sourceDocument,
    operationSnapshot,
    contentSnapshot,
    authoredRunsByBlockId: builder.authoredRunsByBlockId,
    unsupportedFieldsByBlockId: builder.unsupportedFieldsByBlockId,
    sourceOperandsByBlockId,
    tableNodes,
    numberingReferenceKeys: Object.freeze([...numberingReferenceKeysOf(operationSnapshot)]),
  });
  return snapshot;
};

/** @internal Canonical content capsule consumed by the private comparison fast path. */
export const resolvedDocxContentSnapshot = (
  snapshot: ResolvedDocxStorySnapshot,
): OwnedContentSnapshot => payloadOf(snapshot).contentSnapshot;

/** @internal Exact story handle captured with this canonical projection. */
export const resolvedDocxStoryHandle = (
  snapshot: ResolvedDocxStorySnapshot,
): FolioDocumentStoryHandle => payloadOf(snapshot).story;

/** @internal Canonical owned blocks, without repeating public descriptor capture. */
export const resolvedDocxContentBlocks = (
  snapshot: ResolvedDocxStorySnapshot,
): readonly FolioContentBlock[] =>
  requireOwnedContentSnapshotBlocks(payloadOf(snapshot).contentSnapshot);

/** @internal Nominal source operand for the exact canonical block instance. */
export const resolvedDocxSourceOperand = (
  snapshot: ResolvedDocxStorySnapshot,
  block: FolioContentBlock,
): ResolvedDocxSourceOperand => {
  const payload = payloadOf(snapshot);
  const source = payload.sourceOperandsByBlockId.get(block.identity.id);
  const sourcePayload = source && payloadBySourceOperand.get(source);
  if (!source || sourcePayload?.block !== block) {
    return panic("A DOCX source operand must name its capsule's canonical block", {
      blockId: block.identity.id,
    });
  }
  return source;
};

/** @internal Nominal source operand resolved by this capsule's exact block identity. */
export const resolvedDocxSourceOperandForBlockId = (
  snapshot: ResolvedDocxStorySnapshot,
  blockId: string,
): ResolvedDocxSourceOperand =>
  payloadOf(snapshot).sourceOperandsByBlockId.get(blockId) ??
  panic("A DOCX source operand identity is absent from its story capsule", { blockId });

/** @internal Resolve a nominal source operand only for its issuing story capsule. */
export const resolvedDocxSourceOperandBlock = (
  source: ResolvedDocxSourceOperand,
  snapshot: ResolvedDocxStorySnapshot,
): FolioContentBlock => {
  const payload =
    payloadBySourceOperand.get(source) ?? panic("A DOCX source operand was not created by Folio");
  if (payload.snapshot !== snapshot) {
    return panic("A DOCX source operand belongs to another story snapshot");
  }
  return payload.block;
};

/** @internal Exact story capsule identity carried by a nominal source operand. */
export const resolvedDocxSourceOperandSnapshot = (
  source: ResolvedDocxSourceOperand,
): ResolvedDocxStorySnapshot =>
  (payloadBySourceOperand.get(source) ?? panic("A DOCX source operand was not created by Folio"))
    .snapshot;

/** @internal Operation anchors bound to the same live projection. */
export const resolvedDocxOperationSnapshot = (
  snapshot: ResolvedDocxStorySnapshot,
): FolioAIEditSnapshot => payloadOf(snapshot).operationSnapshot;

/** @internal Exact immutable PM source identity for stale-state refusal. */
export const resolvedDocxSourceDocument = (snapshot: ResolvedDocxStorySnapshot): PMNode =>
  payloadOf(snapshot).sourceDocument;

/** @internal Target table templates captured with the canonical projection. */
export const resolvedDocxTableNodes = (
  snapshot: ResolvedDocxStorySnapshot,
): ReadonlyMap<number, PMNode> => new Map(payloadOf(snapshot).tableNodes);

/** @internal Resolve one exact table without exposing the capsule's ownership map. */
export const resolvedDocxTableNode = (
  snapshot: ResolvedDocxStorySnapshot,
  tableIndex: number,
): PMNode =>
  payloadOf(snapshot).tableNodes.get(tableIndex) ??
  panic("A DOCX story capsule has no table at its canonical index", { tableIndex });

/** @internal Optional table lookup for typed transport preflight. */
export const findResolvedDocxTableNode = (
  snapshot: ResolvedDocxStorySnapshot,
  tableIndex: number,
): PMNode | null => payloadOf(snapshot).tableNodes.get(tableIndex) ?? null;

/** @internal Numbering references captured during the exact source walk. */
export const resolvedDocxNumberingReferenceKeys = (
  snapshot: ResolvedDocxStorySnapshot,
): readonly string[] => payloadOf(snapshot).numberingReferenceKeys;

/** @internal Authored run projection for one complete canonical block. */
export const resolvedDocxAuthoredRunsForBlock = (
  snapshot: ResolvedDocxStorySnapshot,
  block: FolioContentBlock,
): readonly DocxAuthoredRun[] => {
  resolvedDocxSourceOperand(snapshot, block);
  const projection = payloadOf(snapshot).authoredRunsByBlockId.get(block.identity.id);
  if (!projection) {
    return panic("A canonical DOCX block has no live authored-run projection", {
      blockId: block.identity.id,
    });
  }
  if (projection.status === "unsupported") {
    return panic("A lossy DOCX authored-run projection reached transport lowering", {
      blockId: block.identity.id,
      reason: projection.reason,
    });
  }
  return projection.runs;
};

/** @internal Whether transport has exact authored runs for this canonical block. */
export const resolvedDocxHasExactAuthoredRuns = (
  snapshot: ResolvedDocxStorySnapshot,
  block: FolioContentBlock,
): boolean => {
  resolvedDocxSourceOperand(snapshot, block);
  return payloadOf(snapshot).authoredRunsByBlockId.get(block.identity.id)?.status === "exact";
};

/** @internal Explicit projection gaps that prevent a complete DOCX proof. */
export const resolvedDocxUnsupportedProjectionFields = (
  snapshot: ResolvedDocxStorySnapshot,
  block: FolioContentBlock,
): readonly string[] => {
  resolvedDocxSourceOperand(snapshot, block);
  return (
    payloadOf(snapshot).unsupportedFieldsByBlockId.get(block.identity.id) ??
    panic("A canonical DOCX block has no projection-support disposition", {
      blockId: block.identity.id,
    })
  );
};

/** @internal Authored run projection rebased onto one canonical UTF-16 range. */
export const resolvedDocxAuthoredRunsForRange = (
  snapshot: ResolvedDocxStorySnapshot,
  block: FolioContentBlock,
  startOffset: number,
  endOffset: number,
): readonly DocxAuthoredRun[] => {
  if (
    !Number.isSafeInteger(startOffset) ||
    !Number.isSafeInteger(endOffset) ||
    startOffset < 0 ||
    endOffset < startOffset ||
    endOffset > block.text.length
  ) {
    return panic("A DOCX authored-run projection received an invalid range", {
      blockId: block.identity.id,
      startOffset,
      endOffset,
    });
  }
  if (startOffset === endOffset) {
    return Object.freeze([
      Object.freeze({
        startOffset: 0,
        endOffset: 0,
        formatting: resolvedDocxAuthoredFormattingAt(snapshot, block, startOffset),
      }),
    ]);
  }
  const projected: DocxAuthoredRun[] = [];
  for (const run of resolvedDocxAuthoredRunsForBlock(snapshot, block)) {
    const overlapStart = Math.max(startOffset, run.startOffset);
    const overlapEnd = Math.min(endOffset, run.endOffset);
    if (overlapStart >= overlapEnd) continue;
    projected.push(
      Object.freeze({
        startOffset: overlapStart - startOffset,
        endOffset: overlapEnd - startOffset,
        formatting: run.formatting,
      }),
    );
  }
  if (projected.at(-1)?.endOffset !== endOffset - startOffset) {
    return panic("A DOCX authored-run range does not reconstruct canonical text", {
      blockId: block.identity.id,
      startOffset,
      endOffset,
    });
  }
  return Object.freeze(projected);
};

/** @internal Live authored run formatting at one canonical UTF-16 boundary. */
export const resolvedDocxAuthoredFormattingAt = (
  snapshot: ResolvedDocxStorySnapshot,
  block: FolioContentBlock,
  offset: number,
): Readonly<TextFormatting> => {
  const run = resolvedDocxAuthoredRunsForBlock(snapshot, block).find(
    ({ startOffset, endOffset }) =>
      (offset >= startOffset && offset < endOffset) ||
      (offset === block.text.length && endOffset === offset),
  );
  return run?.formatting ?? ownFormatting(undefined);
};
