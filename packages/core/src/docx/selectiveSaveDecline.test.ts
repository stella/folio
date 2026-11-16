import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  serializationDiagnosticSource,
  serializationPathSource,
  serializationTripwireDiagnosticSource,
} from "../../../../test/serializationPaths";
import type { LayoutInstrumentation } from "../layout-engine/layoutInstrumentation";
import type { HeaderFooter, Paragraph } from "../types/document";
import type { SelectiveSaveDecline } from "../types/selectiveSaveDecline";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { createDocx, repackDocx } from "./rezip";
import type { SaveDiagnostic } from "./saveDiagnostics";
import { attemptSelectiveSave } from "./selectiveSave";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

/** Two identified paragraphs; the second may hold a bookmark. */
const source = async (bookmark: boolean): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  const marked = bookmark
    ? `<w:bookmarkStart w:id="0" w:name="Mark"/><w:r><w:t>Marked</w:t></w:r><w:bookmarkEnd w:id="0"/>`
    : `<w:r><w:t>Marked</w:t></w:r>`;
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body><w:p w14:paraId="11111111" w14:textId="11111111"><w:r><w:t>Plain</w:t></w:r></w:p><w:p w14:paraId="22222222" w14:textId="22222222">${marked}</w:p></w:body></w:document>`,
  );
  return (await zip.generate()).slice().buffer;
};

/** Insert an empty paragraph and save selectively, collecting declines. */
const insertAndSave = async (bookmark: boolean, maxBytes?: number) => {
  const bytes = await source(bookmark);
  const document = await parseDocx(bytes, {
    preloadFonts: false,
    detectVariables: false,
  });
  document.package.document.content.splice(1, 0, {
    type: "paragraph",
    paraId: "33333333",
    content: [],
  });
  const declines: SelectiveSaveDecline[] = [];
  const saved = await attemptSelectiveSave(document, bytes, {
    changedParaIds: new Set(["33333333"]),
    structuralChange: true,
    hasUntrackedChanges: false,
    ...(maxBytes === undefined ? {} : { maxBytes }),
    onSelectiveDecline: (decline) => declines.push(decline),
  });
  return { saved, declines };
};

describe("selective save declines", () => {
  test("a safe splice preserves an unrelated bookmark without declining", async () => {
    const { saved, declines } = await insertAndSave(true);
    expect(saved).not.toBeNull();
    expect(declines).toEqual([]);
    if (saved)
      expect(
        await (await ZipArchive.read(saved)).entry("word/document.xml")?.text(),
      ).toContain('w:name="Mark"');
  });

  test("a body without range elements is spliced with no decline", async () => {
    const { saved, declines } = await insertAndSave(false);
    expect(saved).not.toBeNull();
    expect(declines).toEqual([]);
  });

  test("a source over the size limit declines as too large", async () => {
    const { saved, declines } = await insertAndSave(false, 1);
    expect(saved).toBeNull();
    expect(declines).toEqual([{ scope: "package", reason: "tooLarge" }]);
  });
});

test("the public refusal callback and measurement report preserve the structural guard's reason", async () => {
  const bytes = await source(true);
  const document = await parseDocx(bytes, { preloadFonts: false });
  const marked = document.package.document.content.at(1);
  if (marked?.type !== "paragraph") panic("Expected marked source paragraph");
  const end = marked.content.find((item) => item.type === "bookmarkEnd");
  if (!end) panic("Expected authored closing bookmark");
  marked.content = marked.content.filter((item) => item !== end);
  document.package.document.content.push({
    type: "paragraph",
    paraId: "33333333",
    content: [end],
  });
  const declines: SelectiveSaveDecline[] = [];
  type Attempt = Parameters<
    NonNullable<LayoutInstrumentation["onSelectiveSaveAttempt"]>
  >[0];
  const attempts: Attempt[] = [];
  const before = globalThis.__folioLayoutInstrumentation;
  globalThis.__folioLayoutInstrumentation = {
    onSelectiveSaveAttempt: (event) => attempts.push(event),
  };
  try {
    const saved = await attemptSelectiveSave(document, bytes, {
      changedParaIds: new Set(["33333333"]),
      structuralChange: true,
      hasUntrackedChanges: false,
      onSelectiveDecline: (decline) => declines.push(decline),
    });
    expect(saved).toBeNull();
    expect(declines).toEqual([
      {
        scope: "part",
        part: "word/document.xml",
        reason: "structuralRangeSplit",
      },
    ]);
    expect(attempts).toEqual([
      { written: false, structural: true, decline: declines.at(0) },
    ]);
  } finally {
    globalThis.__folioLayoutInstrumentation = before;
  }
});

const linkedParagraph = () =>
  ({
    type: "paragraph",
    paraId: "33333333",
    content: [
      {
        type: "hyperlink",
        href: "https://example.test/resource",
        children: [{ type: "run", content: [{ type: "text", text: "Link" }] }],
      },
    ],
  }) satisfies Paragraph;

for (const kind of ["header", "footer"] as const)
  for (const binding of ["resolved", "unresolved"] as const)
    test(`${kind} ${binding} refusal names its relationship owner`, async () => {
      const bytes = await createDocx(serializationPathSource());
      const document = await parseDocx(bytes, { preloadFonts: false });
      const part = {
        type: kind,
        hdrFtrType: "default",
        content: [linkedParagraph()],
      } satisfies HeaderFooter;
      if (kind === "header")
        document.package.headers = new Map([["rId44", part]]);
      else document.package.footers = new Map([["rId44", part]]);
      if (binding === "resolved") {
        document.package.relationships ??= new Map();
        document.package.relationships.set("rId44", {
          id: "rId44",
          type: RELATIONSHIP_TYPES[kind],
          target: `${kind}s/custom.xml`,
        });
      }
      const declines: SelectiveSaveDecline[] = [];
      expect(
        await attemptSelectiveSave(document, bytes, {
          changedParaIds: new Set(),
          structuralChange: false,
          hasUntrackedChanges: false,
          onSelectiveDecline: (decline) => declines.push(decline),
        }),
      ).toBeNull();
      expect(declines).toEqual([
        binding === "resolved"
          ? {
              scope: "part",
              part: `word/${kind}s/custom.xml`,
              reason: "unmaterializedStoryResource",
            }
          : {
              scope: "story",
              story: { kind, rId: "rId44" },
              reason: "unmaterializedStoryResource",
            },
      ]);
    });

for (const kind of ["footnote", "endnote"] as const)
  test(`${kind} resource refusal names the note part`, async () => {
    const bytes = await createDocx(serializationPathSource());
    const document = await parseDocx(bytes, { preloadFonts: false });
    if (kind === "footnote")
      document.package.footnotes = [
        { type: kind, id: 1, content: [linkedParagraph()] },
      ];
    else
      document.package.endnotes = [
        { type: kind, id: 1, content: [linkedParagraph()] },
      ];
    const declines: SelectiveSaveDecline[] = [];
    expect(
      await attemptSelectiveSave(document, bytes, {
        changedParaIds: new Set(),
        structuralChange: false,
        hasUntrackedChanges: false,
        onSelectiveDecline: (decline) => declines.push(decline),
      }),
    ).toBeNull();
    expect(declines).toEqual([
      {
        scope: "part",
        part: `word/${kind}s.xml`,
        reason: "unmaterializedStoryResource",
      },
    ]);
  });

test("package read failures carry no fabricated main-document part", async () => {
  const declines: SelectiveSaveDecline[] = [];
  const malformed = new Uint8Array([0x50, 0x4b, 0, 0]).buffer;
  expect(
    await attemptSelectiveSave(serializationPathSource(), malformed, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
      onSelectiveDecline: (decline) => declines.push(decline),
    }),
  ).toBeNull();
  expect(declines).toEqual([{ scope: "package", reason: "packageError" }]);
});

test("a host diagnostic exception cannot become a package fallback", async () => {
  const document = await serializationDiagnosticSource();
  const bytes =
    document.originalBuffer ?? panic("Expected actual source package");
  const diagnostics: SaveDiagnostic[] = [];
  const options = {
    changedParaIds: new Set<string>(),
    structuralChange: false,
    hasUntrackedChanges: false,
  };
  expect(
    await attemptSelectiveSave(document, bytes, {
      ...options,
      onDiagnostic: (diagnostic) => diagnostics.push(diagnostic),
    }),
  ).toBeNull();
  expect(diagnostics).toEqual([
    { type: "selectiveSaveRefused", part: "word/styles.xml" },
  ]);
  const declines: SelectiveSaveDecline[] = [];
  await expect(
    attemptSelectiveSave(document, bytes, {
      ...options,
      onDiagnostic: () => panic("Host diagnostic failed"),
      onSelectiveDecline: (decline) => declines.push(decline),
    }),
  ).rejects.toThrow("Host diagnostic failed");
  expect(declines).toEqual([]);
});

test.each(["matched", "stale"] as const)(
  "a %s header source succeeds selectively before the full probe",
  async (provenance) => {
    const document = await serializationTripwireDiagnosticSource(provenance);
    const bytes = document.originalBuffer ?? panic("Expected source package");
    const selectiveDiagnostics: SaveDiagnostic[] = [];
    const fullDiagnostics: SaveDiagnostic[] = [];
    const declines: SelectiveSaveDecline[] = [];
    expect(
      await attemptSelectiveSave(document, bytes, {
        changedParaIds: new Set(),
        structuralChange: false,
        hasUntrackedChanges: false,
        onDiagnostic: (diagnostic) => selectiveDiagnostics.push(diagnostic),
        onSelectiveDecline: (decline) => declines.push(decline),
      }),
    ).not.toBeNull();
    expect(
      (
        await repackDocx(document, {
          onDiagnostic: (diagnostic) => fullDiagnostics.push(diagnostic),
        })
      ).byteLength,
    ).toBeGreaterThan(0);
    expect(declines).toEqual([]);
    expect(selectiveDiagnostics).toEqual(
      provenance === "matched"
        ? []
        : [{ type: "sourceReplayMismatch", part: "word/header1.xml" }],
    );
    expect(fullDiagnostics).toEqual(selectiveDiagnostics);
  },
);
