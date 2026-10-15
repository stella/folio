import JSZip from "jszip";
import { writeZipEntry } from "../src/zip";

const zip = new JSZip();
writeZipEntry(zip, "text.xml", "👩🏽‍⚖️");
writeZipEntry(zip, "bytes.bin", new Uint8Array());
writeZipEntry(zip, "buffer.bin", new ArrayBuffer(0));
writeZipEntry(zip, "promised.bin", Promise.resolve(new Uint8Array()));
writeZipEntry(zip, "array.bin", [0, 255]);
writeZipEntry(zip, "blob.bin", new Blob());
writeZipEntry(zip, "view.bin", new DataView(new ArrayBuffer(0)));
declare const stream: NodeJS.ReadableStream;
writeZipEntry(zip, "stream.bin", stream);

writeZipEntry(zip, "text.xml", Promise.resolve("text"));
// @ts-expect-error -- Text streams must use the stateful reader before archive writing.
writeZipEntry(zip, "text.xml", new ReadableStream<string>());
