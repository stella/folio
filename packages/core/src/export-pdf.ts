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

export const exportDocxToPdf = async (
  input: DocxInput,
  options: ExportDocxToPdfOptions,
): Promise<Result<ExportDocxToPdfResult, ExportPdfError>> => {
  // The provider is process-wide state. An export must measure against its own
  // fonts, but a caller that also has an editor open must get its canvas
  // backend back afterwards, including when the export fails.
  const callerProvider = getMeasureProvider();
  const headless = installHeadlessMeasureProvider(options.fonts);

  const laidOut = await layoutDocxHeadless(input, { pageGap: 0 });
  if (laidOut.isErr()) {
    setMeasureProvider(callerProvider);
    return Result.err(new ExportPdfError({ message: laidOut.error.message, cause: laidOut.error }));
  }

  const list = buildDisplayList({
    layout: laidOut.value.layout,
    blockLookup: laidOut.value.blockLookup,
    ...(options.metadata === undefined ? {} : { metadata: options.metadata }),
  });

  const written = writePdf(list, {
    fonts: { load: (face) => options.fonts.load(toFontRequest(face)) },
    timestamp: options.timestamp,
    ...(options.producer === undefined ? {} : { producer: options.producer }),
  });
  setMeasureProvider(callerProvider);
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
