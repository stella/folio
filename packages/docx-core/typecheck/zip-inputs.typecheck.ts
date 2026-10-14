import JSZip from "jszip";
import { writeZipPart } from "@stll/docx-core/zip";

const zip = new JSZip();
declare const publicInput: Parameters<JSZip["loadAsync"]>[0];

writeZipPart({ zip, path: "word/document.xml", data: "text" });
writeZipPart({ zip, path: "word/document.xml", data: Promise.resolve("text") });
writeZipPart({ zip, path: "word/document.xml", data: new Uint8Array() });
writeZipPart({ zip, path: "word/document.xml", data: new ArrayBuffer(0) });
writeZipPart({ zip, path: "word/document.xml", data: publicInput });

// @ts-expect-error -- Archive generation takes output options, never text-part input.
zip.generateAsync("text");
