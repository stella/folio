/**
 * `.docx` to PDF, with no browser and no second painter.
 *
 * The chain is four steps and no shortcuts: paginate headlessly, build the
 * display list, hand the display list to the PDF backend, return the bytes.
 * The editor's DOM backend consumes the same display list from the same
 * builder, so the exported page and the displayed page cannot be laid out by
 * two different sets of rules.
 *
 * ## One font source, two consumers
 *
 * Measurement and embedding read the same {@link HeadlessFontSource}. Laying
 * out against one face and embedding another is the subtle version of the
 * divergence this design exists to prevent: the pagination would be right for
 * a font the reader never sees.
 *
 * ## Determinism
 *
 * `exportDocxToPdf(bytes, options)` is a pure function of its arguments.
 * `options.timestamp` is required rather than defaulted, the same contract
 * `compareDocx` uses, so a caller cannot get nondeterministic output by
 * omission.
 */

import { Result, TaggedError } from "better-result";

import { buildDisplayList } from "./display-list/build/buildDisplayList";
import type { DisplayFontFace, DisplayMetadata, DisplayUnsupported } from "./display-list/types";
import { installHeadlessMeasureProvider } from "./fonts/headlessMeasure";
import type { HeadlessFontSource, HeadlessFontSubstitution } from "./fonts/headlessMeasure";
import { layoutDocxHeadless } from "./headless-layout";
import type { HeadlessLayoutGap } from "./headless-layout";
import { getMeasureProvider, setMeasureProvider } from "./layout-engine/measure/measureProvider";
import { writePdf } from "./pdf/writePdf";
import type { PdfSubstitution, PdfUnencodable } from "./pdf/writePdf";
import type { DocxInput } from "./utils/docxInput";

export class ExportPdfError extends TaggedError("ExportPdfError")<{
  message: string;
  cause?: unknown;
}> {}

export type ExportDocxToPdfOptions = {
  /** Supplies face binaries for measurement and for embedding alike. */
  readonly fonts: HeadlessFontSource;
  /**
   * ISO-8601 instant stamped as the PDF's creation and modification date.
   * Required rather than defaulted so a caller cannot get a nondeterministic
   * document by omission; pass the source package's own timestamp or a fixed
   * epoch.
   */
  readonly timestamp: string;
  readonly metadata?: DisplayMetadata;
  readonly producer?: string;
};

export type ExportDocxToPdfResult = {
  readonly bytes: Uint8Array;
  readonly pageCount: number;
  /** Constructs the display list could not represent. */
  readonly unsupported: readonly DisplayUnsupported[];
  /** Stories the headless pipeline does not paginate. */
  readonly layoutGaps: readonly HeadlessLayoutGap[];
  /** Faces measured with a stand-in. */
  readonly measurementSubstitutions: readonly HeadlessFontSubstitution[];
  /** Faces embedded as a stand-in. */
  readonly embeddingSubstitutions: readonly PdfSubstitution[];
  /** Code points painted as `.notdef` because no supplied face covers them. */
  readonly unencodable: readonly PdfUnencodable[];
};

const toFontRequest = ({ family, weight, italic }: DisplayFontFace) => ({
  family,
  bold: weight >= 700,
  italic,
});

/**
 * Exports run one at a time.
 *
 * The measurement provider is process-wide, and an export must install its own
 * fonts across an `await`. Two overlapping exports would otherwise interleave:
 * the second replaces the first's provider, the first resumes and measures
 * against the second's fonts, then restores the provider while the second is
 * still running. Serialising is the honest fix while the provider is ambient;
 * threading it explicitly through the layout call would remove the need, and
 * is the larger change this defers to.
 */
let exportQueue: Promise<unknown> = Promise.resolve();

const runExclusively = <T>(run: () => Promise<T>): Promise<T> => {
  const previous = exportQueue;
  const current = previous.then(run, run);
  // Keep the chain alive after a rejection: the next caller must still run.
  exportQueue = current.then(
    () => undefined,
    () => undefined,
  );
  return current;
};

export const exportDocxToPdf = (
  input: DocxInput,
  options: ExportDocxToPdfOptions,
): Promise<Result<ExportDocxToPdfResult, ExportPdfError>> =>
  runExclusively(() => exportOnce(input, options));

const exportOnce = async (
  input: DocxInput,
  options: ExportDocxToPdfOptions,
): Promise<Result<ExportDocxToPdfResult, ExportPdfError>> => {
  // A caller that also has an editor open must get its canvas backend back on
  // every exit path, including a throw.
  const callerProvider = getMeasureProvider();
  try {
    return await exportWithHeadlessProvider(input, options);
  } finally {
    setMeasureProvider(callerProvider);
  }
};

const exportWithHeadlessProvider = async (
  input: DocxInput,
  options: ExportDocxToPdfOptions,
): Promise<Result<ExportDocxToPdfResult, ExportPdfError>> => {
  const headless = installHeadlessMeasureProvider(options.fonts);

  const laidOut = await layoutDocxHeadless(input, { pageGap: 0 });
  if (laidOut.isErr()) {
    return Result.err(new ExportPdfError({ message: laidOut.error.message, cause: laidOut.error }));
  }

  const list = buildDisplayList({
    layout: laidOut.value.layout,
    blockLookup: laidOut.value.blockLookup,
    // Without these the producer must assume every render-option construct
    // might be present, and reports a gap for a document that has none.
    documentFeatures: laidOut.value.documentFeatures,
    // The page furniture the layout does not carry: page borders, the
    // watermark, the header and footer stories, the footnote bodies. An export
    // that omits them prints the body of a page rather than the page.
    ...laidOut.value.furniture,
    embeddedFonts: laidOut.value.embeddedFonts,
    ...(options.metadata === undefined ? {} : { metadata: options.metadata }),
  });

  const written = writePdf(list, {
    fonts: { load: (face) => options.fonts.load(toFontRequest(face)) },
    timestamp: options.timestamp,
    ...(options.producer === undefined ? {} : { producer: options.producer }),
  });
  if (written.isErr()) {
    return Result.err(new ExportPdfError({ message: written.error.message, cause: written.error }));
  }

  return Result.ok({
    bytes: written.value.bytes,
    pageCount: list.pages.length,
    unsupported: list.unsupported,
    layoutGaps: laidOut.value.unsupported,
    measurementSubstitutions: headless.substitutions(),
    embeddingSubstitutions: written.value.substitutions,
    unencodable: written.value.unencodable,
  });
};
