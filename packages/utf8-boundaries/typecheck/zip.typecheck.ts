import JSZip from "jszip";
import { writeZipEntry } from "../src/zip";

const zip = new JSZip();
writeZipEntry(zip, "text.xml", "👩🏽‍⚖️");
writeZipEntry(zip, "bytes.bin", new Uint8Array());
writeZipEntry(zip, "buffer.bin", new ArrayBuffer(0));
writeZipEntry(zip, "promised.bin", Promise.resolve(new Uint8Array()));

// @ts-expect-error -- Promised text cannot be encoded before JSZip receives it.
writeZipEntry(zip, "text.xml", Promise.resolve("text"));
// @ts-expect-error -- Text streams must use the stateful reader before archive writing.
writeZipEntry(zip, "text.xml", new ReadableStream<string>());
