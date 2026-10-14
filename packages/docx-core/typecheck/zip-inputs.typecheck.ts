import JSZip from "jszip";
import { writeZipPart } from "../src/zip/writeZipPart";

declare const stream: NodeJS.ReadableStream;
const zip = new JSZip();

// @ts-expect-error -- Encode text before it reaches the byte-only ZIP writer.
writeZipPart({ zip, path: "word/document.xml", data: "text" });
// @ts-expect-error -- Promised strings cannot bypass the byte-only ZIP writer.
writeZipPart({ zip, path: "word/document.xml", data: Promise.resolve("text") });
writeZipPart({ zip, path: "word/document.xml", data: new Uint8Array() });

// @ts-expect-error -- A ZIP part must provide complete text or bytes, not a text stream.
writeZipPart({ zip, path: "word/document.xml", data: stream });

// @ts-expect-error -- Archive generation takes output options, never text-part input.
zip.generateAsync("text");
