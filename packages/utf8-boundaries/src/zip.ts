import { TaggedError } from "better-result";
import type JSZip from "jszip";

import { base64ToBytes, encodeBinaryString, encodeUtf8 } from "./encoding";

class ZipStreamInputError extends TaggedError("ZipStreamInputError")<{
  message: string;
}> {}

const joinBytes = (chunks: Uint8Array[]): Uint8Array => {
  const result = new Uint8Array(
    chunks.reduce((length, chunk) => length + chunk.length, 0),
  );
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
};

// Text streams are collected before encoding so a split surrogate pair stays whole.
const readZipStream = (
  stream: NodeJS.ReadableStream,
): Promise<string | Uint8Array> =>
  new Promise((resolve, reject) => {
    const chunks: (string | Uint8Array)[] = [];
    const onError = (error: unknown) => {
      cleanup();
      reject(error);
    };
    const onData = (chunk: unknown) => {
      if (typeof chunk === "string" || chunk instanceof Uint8Array) {
        chunks.push(chunk);
        return;
      }
      onError(
        new ZipStreamInputError({
          message: "ZIP stream emitted an unsupported chunk",
        }),
      );
    };
    const onEnd = () => {
      cleanup();
      if (chunks.every((chunk) => typeof chunk === "string")) {
        resolve(chunks.join(""));
        return;
      }
      if (chunks.every((chunk) => chunk instanceof Uint8Array)) {
        resolve(joinBytes(chunks));
        return;
      }
      reject(
        new ZipStreamInputError({
          message: "ZIP stream mixed text and byte chunks",
        }),
      );
    };
    const cleanup = () => {
      stream.removeListener("data", onData);
      stream.removeListener("end", onEnd);
      stream.removeListener("error", onError);
    };
    stream.on("error", onError);
    stream.on("end", onEnd);
    stream.on("data", onData);
  });

/** Content `writeZipEntry` accepts; text is encoded to UTF-8 once, as a whole. */
export type ZipEntryInput = Parameters<JSZip["loadAsync"]>[0] | ArrayBufferView;

const zipEntryBytes = async (
  input: ZipEntryInput,
  options: JSZip.JSZipFileOptions | undefined,
): Promise<Uint8Array> => {
  const content = await input;
  if (typeof content === "string") {
    if (options?.base64) return base64ToBytes(content);
    return options?.binary ? encodeBinaryString(content) : encodeUtf8(content);
  }
  if (content instanceof ArrayBuffer) return new Uint8Array(content);
  if (ArrayBuffer.isView(content))
    return new Uint8Array(
      content.buffer,
      content.byteOffset,
      content.byteLength,
    );
  if (Array.isArray(content)) return Uint8Array.from(content);
  if (content instanceof Blob)
    return new Uint8Array(await content.arrayBuffer());
  const streamed = await readZipStream(content);
  return zipEntryBytes(streamed, options);
};

/** Normalize complete text and binary inputs before JSZip chunks the entry. */
export const writeZipEntry = (
  zip: JSZip,
  path: string,
  content: ZipEntryInput,
  options?: JSZip.JSZipFileOptions,
): JSZip =>
  zip.file(path, zipEntryBytes(content, options), {
    ...options,
    binary: true,
    base64: false,
  });
