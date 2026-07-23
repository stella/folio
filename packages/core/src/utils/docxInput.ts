/**
 * Flexible input types for DOCX documents.
 *
 * Accepts any common binary format so consumers don't need to manually
 * convert before passing data to the editor or parser.
 */

/**
 * Any binary representation of a DOCX file that the editor can consume.
 *
 * - `ArrayBuffer` — from `FileReader.readAsArrayBuffer()` or `fetch().arrayBuffer()`
 * - `Uint8Array` — from Node.js `fs.readFile()` or streaming APIs
 * - `Blob` — from drag-and-drop or `<input type="file">`
 * - `File` — subclass of Blob, from `<input type="file">`
 */
export type DocxInput = ArrayBuffer | Uint8Array | Blob | File;

/**
 * Normalize any {@link DocxInput} into an `ArrayBuffer` for internal use.
 */
export function toArrayBuffer(input: DocxInput): Promise<ArrayBuffer> {
  if (input instanceof ArrayBuffer) {
    return Promise.resolve(input);
  }
  if (input instanceof Uint8Array) {
    if (
      input.buffer instanceof ArrayBuffer &&
      input.byteOffset === 0 &&
      input.byteLength === input.buffer.byteLength
    ) {
      return Promise.resolve(input.buffer);
    }

    // Copy views over larger or shared buffers into an exact-size ArrayBuffer.
    const copy = new ArrayBuffer(input.byteLength);
    new Uint8Array(copy).set(input);
    return Promise.resolve(copy);
  }
  if (input instanceof Blob) {
    // Blob and File both support arrayBuffer()
    return input.arrayBuffer();
  }
  // Exhaustive check — should never happen at runtime
  return Promise.reject(new TypeError(`Unsupported DocxInput type: ${typeof input}`));
}
