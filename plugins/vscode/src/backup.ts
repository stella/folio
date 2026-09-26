/**
 * The hot-exit backup of a document with unsaved edits: the serialized
 * package behind a one-line header that carries the version of the file the
 * edits were made from, so a restored session saves against the right
 * version (and finds out when the file changed while the editor was closed).
 *
 *   folio-vscode-backup {"version":1,"baseline":"<sha-256>"}\n<.docx bytes>
 */

const MAGIC = "folio-vscode-backup ";
const FORMAT_VERSION = 1;
/** A header is short; anything longer is not one. */
const MAX_HEADER_BYTES = 4096;
const NEWLINE = 0x0a;

const FILE_VERSION = /^[0-9a-f]{64}$/u;

export type Backup = {
  /** The fileVersion of the file on disk the edits were made from. */
  readonly baseline: string;
  /** The serialized package, with the unsaved edits. */
  readonly bytes: Uint8Array;
};

export const encodeBackup = ({ baseline, bytes }: Backup): Uint8Array => {
  const header = new TextEncoder().encode(
    `${MAGIC}${JSON.stringify({ version: FORMAT_VERSION, baseline })}\n`,
  );
  const encoded = new Uint8Array(header.byteLength + bytes.byteLength);
  encoded.set(header, 0);
  encoded.set(bytes, header.byteLength);
  return encoded;
};

/** The backup in `data`, or `null` when it is not one this extension wrote. */
export const decodeBackup = (data: Uint8Array): Backup | null => {
  const end = data.subarray(0, MAX_HEADER_BYTES).indexOf(NEWLINE);
  if (end === -1) return null;
  const header = new TextDecoder().decode(data.subarray(0, end));
  if (!header.startsWith(MAGIC)) return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(header.slice(MAGIC.length));
  } catch {
    return null;
  }
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) return null;
  const version: unknown = Reflect.get(parsed, "version");
  const baseline: unknown = Reflect.get(parsed, "baseline");
  if (version !== FORMAT_VERSION || typeof baseline !== "string" || !FILE_VERSION.test(baseline)) {
    return null;
  }
  return { baseline, bytes: data.slice(end + 1) };
};
