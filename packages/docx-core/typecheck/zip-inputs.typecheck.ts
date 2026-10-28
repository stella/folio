import { ZipArchive } from "@stll/docx-core/zip";

const zip = new ZipArchive();

zip.set("word/document.xml", "text");
zip.set("word/document.xml", new Uint8Array());
zip.set("word/document.xml", new ArrayBuffer(0));

// @ts-expect-error -- Entry content is settled before it is written; await a promise first.
zip.set("word/document.xml", Promise.resolve("text"));

// @ts-expect-error -- Archive generation takes output options, never entry content.
void zip.generate("text");
