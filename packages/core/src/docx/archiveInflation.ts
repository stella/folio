/**
 * Compression bounds for DOCX packages.
 *
 * The archive reader (`@stll/docx-core/zip`) checks every size an archive
 * declares before anything inflates, and stops each inflation at its entry's
 * declared size. The values here are the DOCX policy it is given.
 */

/**
 * Default ceiling on inflated bytes per compressed byte.
 *
 * Document parts in the repository's fixtures stay under 75:1, and parts over a
 * mebibyte under 15:1. A DEFLATE stream tops out near 1032:1, which is the
 * shape of an entry made of one repeated byte.
 */
export const DOCX_MAX_COMPRESSION_RATIO = 200;

/**
 * Inflated size below which the ratio cap does not apply.
 *
 * A small part can compress extremely well without costing anything to hold,
 * so the ratio only has to bound entries (and packages) large enough to
 * matter. The byte caps still apply below this size.
 */
export const DOCX_COMPRESSION_RATIO_GRACE_BYTES: number = 4 * 1024 * 1024;

const RATIO_BOUNDED_EXTENSIONS: ReadonlySet<string> = new Set([
  "xml",
  "rels",
  "vml",
  "txt",
  "htm",
  "html",
  "mht",
  "mhtml",
  "rtf",
]);

/**
 * Whether the per-entry ratio cap applies to a part. Markup and text parts
 * never compress anywhere near the cap. Binary parts can: an uncompressed
 * bitmap of one colour compresses like a run of one byte. Those stay bounded
 * by the byte caps and the package-wide ratio instead.
 */
export const isRatioBoundedPart = (path: string): boolean =>
  RATIO_BOUNDED_EXTENSIONS.has(
    path.slice(path.lastIndexOf(".") + 1).toLowerCase(),
  );
