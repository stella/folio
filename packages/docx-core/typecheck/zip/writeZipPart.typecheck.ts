import JSZip from "jszip";
import { writeZipPart } from "../../src/zip/writeZipPart";

const zip = new JSZip();
writeZipPart({ zip, path: "word/document.xml", data: "<document/>" });
writeZipPart({
  zip,
  path: "word/document.xml",
  data: "<document/>",
  options: {
    // @ts-expect-error Parent-directory creation is owned by the writer.
    createFolders: true,
  },
});
