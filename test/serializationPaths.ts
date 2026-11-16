import { ZipArchive } from "@stll/docx-core/zip";

import { parseDocx } from "../packages/core/src/docx/parser";
import type { FolioSelectiveSaveFlags } from "../packages/core/src/docx/selectiveSaveFlags";
import type { Document } from "../packages/core/src/types/document";
import {
  FOLIO_DOCX_SERIALIZATION_MODE,
  type FolioDocxSerializationMode,
  type FolioDocxSerializationPath,
} from "../packages/core/src/types/docxSerialization";
import type { SelectiveSaveDecline } from "../packages/core/src/types/selectiveSaveDecline";

type SerializationPathCase = {
  kind: string;
  source: "package" | "model";
  mode: FolioDocxSerializationMode;
  flags: FolioSelectiveSaveFlags;
  expectedPath: FolioDocxSerializationPath;
  expectedDeclines: readonly SelectiveSaveDecline[];
};

const PREFER_SELECTIVE = FOLIO_DOCX_SERIALIZATION_MODE.preferSelective;
const FULL = FOLIO_DOCX_SERIALIZATION_MODE.full;
const TOO_LARGE = {
  scope: "package",
  reason: "tooLarge",
} as const satisfies SelectiveSaveDecline;

/** Each path that selects returned bytes, including tripwire-only writes. */
export const SERIALIZATION_PATH_CASES = [
  {
    kind: "selective",
    source: "package",
    mode: PREFER_SELECTIVE,
    flags: { selectiveSave: true },
    expectedPath: "selective",
    expectedDeclines: [],
  },
  {
    kind: "forced full",
    source: "package",
    mode: FULL,
    flags: { selectiveSave: true },
    expectedPath: "full-repack",
    expectedDeclines: [],
  },
  {
    kind: "disabled",
    source: "package",
    mode: PREFER_SELECTIVE,
    flags: { selectiveSave: false },
    expectedPath: "full-repack",
    expectedDeclines: [],
  },
  {
    kind: "no source",
    source: "model",
    mode: PREFER_SELECTIVE,
    flags: { selectiveSave: true },
    expectedPath: "full-repack",
    expectedDeclines: [],
  },
  {
    kind: "declined",
    source: "package",
    mode: PREFER_SELECTIVE,
    flags: { selectiveSave: true, selectiveSaveMaxBytes: 1 },
    expectedPath: "full-repack",
    expectedDeclines: [TOO_LARGE],
  },
  {
    kind: "selective with tripwire",
    source: "package",
    mode: PREFER_SELECTIVE,
    flags: { selectiveSave: true, selectiveSaveTripwire: true },
    expectedPath: "selective",
    expectedDeclines: [],
  },
  {
    kind: "forced full with tripwire",
    source: "package",
    mode: FULL,
    flags: { selectiveSave: true, selectiveSaveTripwire: true },
    expectedPath: "full-repack",
    expectedDeclines: [],
  },
  {
    kind: "disabled with tripwire",
    source: "package",
    mode: PREFER_SELECTIVE,
    flags: { selectiveSave: false, selectiveSaveTripwire: true },
    expectedPath: "full-repack",
    expectedDeclines: [],
  },
  {
    kind: "tripwire declines during forced full",
    source: "package",
    mode: FULL,
    flags: {
      selectiveSave: true,
      selectiveSaveTripwire: true,
      selectiveSaveMaxBytes: 1,
    },
    expectedPath: "full-repack",
    expectedDeclines: [TOO_LARGE],
  },
] as const satisfies readonly SerializationPathCase[];

export const serializationPathSource = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          content: [{ type: "run", content: [{ type: "text", text: "Body" }] }],
        },
      ],
    },
  },
});

/** A valid source package without styles; the authored model adds one. */
export const serializationDiagnosticSource = async (): Promise<Document> => {
  const word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const relationships =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.set(
    "_rels/.rels",
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationships}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `<w:document xmlns:w="${word}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="12345678"><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>`,
  );
  const document = await parseDocx((await zip.generate()).slice().buffer, {
    preloadFonts: false,
  });
  document.package.styles = {
    styles: [{ type: "paragraph", styleId: "Added", name: "Added" }],
  };
  return document;
};
