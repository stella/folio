/**
 * The PDF backend's entry point.
 *
 * Two calls with the same display list and the same timestamp produce the
 * same bytes. That is a hard requirement rather than a nicety: it is what
 * makes an export diffable, cacheable and comparable against a golden file,
 * and every design decision below that looks fussy (one number formatter,
 * resource names from sorted keys, a `/ID` hashed from the body, a fixed
 * deflate level) exists to keep it true. Nothing here reads a clock or a
 * random source.
 */

import { Result, TaggedError } from "better-result";
import { createHash } from "node:crypto";
import type {
  DisplayLink,
  DisplayList,
  DisplayOutlineEntry,
  DisplayPage,
} from "../display-list/types";
import { collectResources, renderContentStream, type ContentResources } from "./contentStream";
import { prepareFonts, type PdfFontSource, type PdfSubstitution } from "./fonts";
import { decodeImage, type PdfImageSamples } from "./images";
import {
  createPdfDocument,
  pdfArray,
  pdfAsciiString,
  pdfDict,
  pdfFlateStream,
  pdfName,
  pdfNull,
  pdfNumber,
  pdfNumberArray,
  pdfStream,
  pdfTextString,
  type PdfDocument,
  type PdfRef,
  type PdfValue,
} from "./objects";
import { collectUsage, paintPage } from "./paint";
import { displayPointToPdf, pxToPt } from "./pageSpace";

export type { PdfFontSource, PdfSubstitution } from "./fonts";

export type WritePdfOptions = {
  readonly fonts: PdfFontSource;
  /**
   * Creation and modification instant stamped into the PDF, ISO-8601.
   * Required rather than defaulted so a caller cannot get a
   * nondeterministic document by omission.
   */
  readonly timestamp: string;
  readonly producer?: string;
};

export type WritePdfResult = {
  readonly bytes: Uint8Array;
  /** Faces the font source could not supply, painted with a base-14 stand-in. */
  readonly substitutions: readonly PdfSubstitution[];
};

export class WritePdfError extends TaggedError("WritePdfError")<{
  message: string;
  cause?: unknown;
}> {}

const DEFAULT_PRODUCER = "folio";

/** Every `/ProcSet` folio's own output can need. */
const PROC_SET = ["PDF", "Text", "ImageB", "ImageC", "ImageI"] as const;

const pad = (value: number, width: number): string => String(value).padStart(width, "0");

const YEAR_DIGITS = 4;
const FIELD_DIGITS = 2;

/**
 * `D:YYYYMMDDHHmmSS+00'00'`. The instant is written in UTC whatever offset
 * the caller's ISO string carried: one instant has one representation here,
 * so two callers who name the same moment get the same bytes.
 */
const formatPdfDate = (timestamp: string): Result<string, WritePdfError> => {
  const milliseconds = Date.parse(timestamp);
  if (Number.isNaN(milliseconds)) {
    return Result.err(
      new WritePdfError({ message: `timestamp is not an ISO-8601 instant: ${timestamp}` }),
    );
  }
  const date = new Date(milliseconds);
  return Result.ok(
    `D:${pad(date.getUTCFullYear(), YEAR_DIGITS)}${pad(date.getUTCMonth() + 1, FIELD_DIGITS)}${pad(date.getUTCDate(), FIELD_DIGITS)}${pad(date.getUTCHours(), FIELD_DIGITS)}${pad(date.getUTCMinutes(), FIELD_DIGITS)}${pad(date.getUTCSeconds(), FIELD_DIGITS)}+00'00'`,
  );
};

/**
 * A `/URI` is a byte string, so anything above ASCII is percent-encoded as
 * UTF-8 rather than written raw.
 */
const asciiUri = (href: string): string => {
  const LAST_ASCII = 0x7e;
  let out = "";
  for (const character of href) {
    out +=
      (character.codePointAt(0) ?? 0) <= LAST_ASCII ? character : encodeURIComponent(character);
  }
  return out;
};

const buildImageObject = (document: PdfDocument, samples: PdfImageSamples): PdfRef => {
  const alphaRef =
    samples.alpha === undefined
      ? undefined
      : document.add(
          pdfFlateStream(
            [
              ["Type", pdfName("XObject")],
              ["Subtype", pdfName("Image")],
              ["Width", pdfNumber(samples.widthPx)],
              ["Height", pdfNumber(samples.heightPx)],
              ["ColorSpace", pdfName("DeviceGray")],
              ["BitsPerComponent", pdfNumber(samples.alpha.bitsPerComponent)],
            ],
            samples.alpha.data,
          ),
        );
  const entries = [
    ["Type", pdfName("XObject")],
    ["Subtype", pdfName("Image")],
    ["Width", pdfNumber(samples.widthPx)],
    ["Height", pdfNumber(samples.heightPx)],
    ["ColorSpace", samples.colorSpace],
    ["BitsPerComponent", pdfNumber(samples.bitsPerComponent)],
    ["Decode", samples.decode === undefined ? undefined : pdfNumberArray(samples.decode)],
    ["SMask", alphaRef],
  ] as const satisfies readonly (readonly [string, PdfValue | undefined])[];
  return samples.encoding === "jpeg"
    ? document.add(pdfStream(pdfDict([...entries, ["Filter", pdfName("DCTDecode")]]), samples.data))
    : document.add(pdfFlateStream(entries, samples.data));
};

type ResourceDictOptions = {
  readonly resources: ContentResources;
  readonly fontRefs: ReadonlyMap<number, PdfRef>;
  readonly imageRefs: ReadonlyMap<number, PdfRef>;
};

const buildResources = ({ resources, fontRefs, imageRefs }: ResourceDictOptions): PdfValue => {
  const fontEntries = [...resources.fontNames].map(
    ([fontIndex, name]) => [name, fontRefs.get(fontIndex)] as const,
  );
  const imageEntries = [...resources.imageNames].map(
    ([imageIndex, name]) => [name, imageRefs.get(imageIndex)] as const,
  );
  const stateEntries = [...resources.extGStateNames].map(
    ([alpha, name]) =>
      [
        name,
        pdfDict([
          ["Type", pdfName("ExtGState")],
          ["ca", pdfNumber(Number(alpha))],
          ["CA", pdfNumber(Number(alpha))],
        ]),
      ] as const,
  );
  return pdfDict([
    ["ProcSet", pdfArray(PROC_SET.map(pdfName))],
    ["Font", fontEntries.length === 0 ? undefined : pdfDict(fontEntries)],
    ["XObject", imageEntries.length === 0 ? undefined : pdfDict(imageEntries)],
    ["ExtGState", stateEntries.length === 0 ? undefined : pdfDict(stateEntries)],
  ]);
};

type DestinationOptions = {
  readonly pageRefs: readonly PdfRef[];
  readonly pages: readonly DisplayPage[];
  readonly pageIndex: number;
  readonly yPx: number;
};

/**
 * `/XYZ` with a null x and zoom: the reader scrolls the point to the top of
 * the window and leaves the horizontal position and magnification alone.
 */
const buildDestination = ({
  pageRefs,
  pages,
  pageIndex,
  yPx,
}: DestinationOptions): Result<PdfValue, WritePdfError> => {
  const pageRef = pageRefs[pageIndex];
  const page = pages[pageIndex];
  if (pageRef === undefined || page === undefined) {
    return Result.err(
      new WritePdfError({
        message: `destination names page ${String(pageIndex)}, which does not exist`,
      }),
    );
  }
  const point = displayPointToPdf(page.heightPx, 0, yPx);
  return Result.ok(pdfArray([pageRef, pdfName("XYZ"), pdfNull, pdfNumber(point.y), pdfNull]));
};

type AnnotationOptions = {
  readonly document: PdfDocument;
  readonly link: DisplayLink;
  readonly page: DisplayPage;
  readonly pages: readonly DisplayPage[];
  readonly pageRefs: readonly PdfRef[];
};

const buildAnnotation = ({
  document,
  link,
  page,
  pages,
  pageRefs,
}: AnnotationOptions): Result<PdfRef, WritePdfError> => {
  // Annotation rectangles live in default user space, untouched by the
  // content stream's CTM, so they run the page matrix themselves.
  const topLeft = displayPointToPdf(page.heightPx, link.rect.xPx, link.rect.yPx);
  const bottomRight = displayPointToPdf(
    page.heightPx,
    link.rect.xPx + link.rect.widthPx,
    link.rect.yPx + link.rect.heightPx,
  );
  let action: PdfValue | undefined;
  let destination: PdfValue | undefined;
  switch (link.target.kind) {
    case "external":
      action = pdfDict([
        ["S", pdfName("URI")],
        ["URI", pdfAsciiString(asciiUri(link.target.href))],
      ]);
      break;
    case "page": {
      const built = buildDestination({
        pageRefs,
        pages,
        pageIndex: link.target.pageIndex,
        yPx: link.target.yPx,
      });
      if (built.isErr()) {
        return Result.err(built.error);
      }
      destination = built.value;
      break;
    }
    default: {
      const unreachable: never = link.target;
      return Result.err(
        new WritePdfError({ message: `unhandled link target: ${JSON.stringify(unreachable)}` }),
      );
    }
  }
  return Result.ok(
    document.add(
      pdfDict([
        ["Type", pdfName("Annot")],
        ["Subtype", pdfName("Link")],
        [
          "Rect",
          pdfNumberArray([
            Math.min(topLeft.x, bottomRight.x),
            Math.min(topLeft.y, bottomRight.y),
            Math.max(topLeft.x, bottomRight.x),
            Math.max(topLeft.y, bottomRight.y),
          ]),
        ],
        ["Border", pdfNumberArray([0, 0, 0])],
        ["Contents", link.tooltip === undefined ? undefined : pdfTextString(link.tooltip)],
        ["A", action],
        ["Dest", destination],
      ]),
    ),
  );
};

type OutlineNode = {
  readonly entry: DisplayOutlineEntry;
  readonly children: OutlineNode[];
  readonly ref: PdfRef;
};

const buildOutlineTree = (
  entries: readonly DisplayOutlineEntry[],
  document: PdfDocument,
): readonly OutlineNode[] => {
  const roots: OutlineNode[] = [];
  const ancestors: OutlineNode[] = [];
  for (const entry of entries) {
    const node: OutlineNode = { entry, children: [], ref: document.allocate() };
    while ((ancestors.at(-1)?.entry.level ?? -1) >= entry.level) {
      ancestors.pop();
    }
    const parent = ancestors.at(-1);
    if (parent === undefined) {
      roots.push(node);
    } else {
      parent.children.push(node);
    }
    ancestors.push(node);
  }
  return roots;
};

/** Every node is open, so a node's `/Count` is its whole subtree. */
const openDescendantCount = (nodes: readonly OutlineNode[]): number =>
  nodes.reduce((total, node) => total + 1 + openDescendantCount(node.children), 0);

type DefineOutlineOptions = {
  readonly document: PdfDocument;
  readonly nodes: readonly OutlineNode[];
  readonly parentRef: PdfRef;
  readonly pages: readonly DisplayPage[];
  readonly pageRefs: readonly PdfRef[];
};

const defineOutlineNodes = ({
  document,
  nodes,
  parentRef,
  pages,
  pageRefs,
}: DefineOutlineOptions): WritePdfError | null => {
  for (const [index, node] of nodes.entries()) {
    const destination = buildDestination({
      pageRefs,
      pages,
      pageIndex: node.entry.pageIndex,
      yPx: node.entry.yPx,
    });
    if (destination.isErr()) {
      return destination.error;
    }
    const childCount = openDescendantCount(node.children);
    document.define(
      node.ref,
      pdfDict([
        ["Title", pdfTextString(node.entry.title)],
        ["Parent", parentRef],
        ["Prev", nodes[index - 1]?.ref],
        ["Next", nodes[index + 1]?.ref],
        ["First", node.children.at(0)?.ref],
        ["Last", node.children.at(-1)?.ref],
        ["Count", childCount === 0 ? undefined : pdfNumber(childCount)],
        ["Dest", destination.value],
      ]),
    );
    const error = defineOutlineNodes({
      document,
      nodes: node.children,
      parentRef: node.ref,
      pages,
      pageRefs,
    });
    if (error !== null) {
      return error;
    }
  }
  return null;
};

export const writePdf = (
  list: DisplayList,
  options: WritePdfOptions,
): Result<WritePdfResult, WritePdfError> => {
  if (list.pages.length === 0) {
    return Result.err(new WritePdfError({ message: "a PDF needs at least one page" }));
  }
  const stamp = formatPdfDate(options.timestamp);
  if (stamp.isErr()) {
    return Result.err(stamp.error);
  }
  const usage = collectUsage(list.pages, list.fonts.length, list.images.length);
  if (usage.isErr()) {
    return Result.err(new WritePdfError({ message: usage.error.message, cause: usage.error }));
  }

  const document = createPdfDocument();
  // Allocated up front and in this order, so the object numbers of the
  // document's spine never depend on what the pages happen to contain.
  const catalogRef = document.allocate();
  const pagesRef = document.allocate();
  const infoRef = document.allocate();
  const pageRefs = list.pages.map(() => document.allocate());

  const { byFontIndex, substitutions } = prepareFonts({
    document,
    faces: list.fonts,
    usedCodePoints: usage.value.codePointsByFont,
    source: options.fonts,
  });
  const fontRefs = new Map([...byFontIndex].map(([index, font]) => [index, font.ref]));

  const imageRefs = new Map<number, PdfRef>();
  for (const imageIndex of [...usage.value.imageIndices].sort((left, right) => left - right)) {
    const source = list.images[imageIndex];
    if (source === undefined) {
      return Result.err(
        new WritePdfError({ message: `image ${String(imageIndex)} is outside the image table` }),
      );
    }
    const decoded = decodeImage(source);
    if (decoded.isErr()) {
      return Result.err(
        new WritePdfError({
          message: `image ${String(imageIndex)} could not be decoded: ${decoded.error.message}`,
          cause: decoded.error,
        }),
      );
    }
    imageRefs.set(imageIndex, buildImageObject(document, decoded.value));
  }

  const outlineRoots = buildOutlineTree(list.outline, document);
  const outlineRef = outlineRoots.length === 0 ? undefined : document.allocate();

  for (const [pageIndex, page] of list.pages.entries()) {
    const pageRef = pageRefs[pageIndex];
    if (pageRef === undefined) {
      return Result.err(
        new WritePdfError({ message: `page ${String(pageIndex)} lost its object` }),
      );
    }
    const parts = paintPage({ page, fonts: byFontIndex });
    const resources = collectResources(parts);
    const contentRef = document.add(
      pdfFlateStream([], new TextEncoder().encode(renderContentStream(parts, resources))),
    );
    const annotationRefs: PdfRef[] = [];
    for (const link of page.links) {
      const annotation = buildAnnotation({ document, link, page, pages: list.pages, pageRefs });
      if (annotation.isErr()) {
        return Result.err(annotation.error);
      }
      annotationRefs.push(annotation.value);
    }
    document.define(
      pageRef,
      pdfDict([
        ["Type", pdfName("Page")],
        ["Parent", pagesRef],
        ["MediaBox", pdfNumberArray([0, 0, pxToPt(page.widthPx), pxToPt(page.heightPx)])],
        ["Resources", buildResources({ resources, fontRefs, imageRefs })],
        ["Contents", contentRef],
        ["Annots", annotationRefs.length === 0 ? undefined : pdfArray(annotationRefs)],
      ]),
    );
  }

  if (outlineRef !== undefined) {
    const error = defineOutlineNodes({
      document,
      nodes: outlineRoots,
      parentRef: outlineRef,
      pages: list.pages,
      pageRefs,
    });
    if (error !== null) {
      return Result.err(error);
    }
    document.define(
      outlineRef,
      pdfDict([
        ["Type", pdfName("Outlines")],
        ["First", outlineRoots.at(0)?.ref],
        ["Last", outlineRoots.at(-1)?.ref],
        ["Count", pdfNumber(openDescendantCount(outlineRoots))],
      ]),
    );
  }

  document.define(
    pagesRef,
    pdfDict([
      ["Type", pdfName("Pages")],
      ["Kids", pdfArray(pageRefs)],
      ["Count", pdfNumber(pageRefs.length)],
    ]),
  );
  document.define(
    catalogRef,
    pdfDict([
      ["Type", pdfName("Catalog")],
      ["Pages", pagesRef],
      ["Outlines", outlineRef],
      ["PageMode", outlineRef === undefined ? undefined : pdfName("UseOutlines")],
    ]),
  );
  const { metadata } = list;
  document.define(
    infoRef,
    pdfDict([
      ["Title", metadata.title === undefined ? undefined : pdfTextString(metadata.title)],
      ["Author", metadata.author === undefined ? undefined : pdfTextString(metadata.author)],
      ["Subject", metadata.subject === undefined ? undefined : pdfTextString(metadata.subject)],
      ["Keywords", metadata.keywords === undefined ? undefined : pdfTextString(metadata.keywords)],
      ["Producer", pdfTextString(options.producer ?? DEFAULT_PRODUCER)],
      ["CreationDate", pdfAsciiString(stamp.value)],
      ["ModDate", pdfAsciiString(stamp.value)],
    ]),
  );

  return Result.ok({
    bytes: document.serialize({
      rootRef: catalogRef,
      infoRef,
      idHex: (body) => createHash("sha256").update(body).digest("hex").toUpperCase(),
    }),
    substitutions,
  });
};
