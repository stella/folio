import JSZip from "jszip";
import { writeZipPart } from "../../src/zip/writeZipPart";

const zip = new JSZip();
writeZipPart({ zip, path: "word/document.xml", data: new Uint8Array() });
writeZipPart({
  zip,
  path: "word/document.xml",
  data: new Uint8Array(),
  options: {
    // @ts-expect-error Parent-directory creation is owned by the writer.
    createFolders: true,
  },
});
