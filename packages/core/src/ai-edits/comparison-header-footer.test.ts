import { expect, test } from "bun:test";
import JSZip from "jszip";
import { FolioDocxReviewer, getFolioDocxComparisonAccess } from "./headless";
import { createDocx } from "../docx/rezip";
import { parseRelationships, RELATIONSHIP_TYPES } from "../docx/relsParser";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "../compare/compare";

const sourceDocument = () => {
  const document = createEmptyDocument();
  document.package.headers = new Map([
    [
      "rId_header",
      {
        type: "header" as const,
        hdrFtrType: "default" as const,
        content: [
          {
            type: "paragraph" as const,
            content: [
              {
                type: "run" as const,
                content: [{ type: "text" as const, text: "Header text" }],
              },
            ],
          },
        ],
      },
    ],
  ]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: "rId_header" }],
  };
  return document;
};

test("comparison story imports are empty, collision-safe, and captured by saves", async () => {
  const source = await FolioDocxReviewer.fromBuffer(await createDocx(sourceDocument()));
  const sourceHandle = source.listStories().find(({ handle }) => handle.type === "header")?.handle;
  if (!sourceHandle || sourceHandle.type !== "header") throw new Error("Missing source header");
  const destination = await FolioDocxReviewer.fromBuffer(await createDocx(createEmptyDocument()));
  const original = destination.toDocument();
  const access = getFolioDocxComparisonAccess(destination);
  const first = await access.createComparisonHeaderFooter(source, sourceHandle);
  if (!first) throw new Error("Header import refused");
  expect(original.package.relationships?.has(first.relationshipId)).toBe(false);
  expect(destination.readStory(first)?.text).toBe("");
  const firstSave = destination.toBuffer();
  const second = await access.createComparisonHeaderFooter(source, sourceHandle);
  if (!second) throw new Error("Second header import refused");
  expect(second.relationshipId).not.toBe(first.relationshipId);
  expect(original.package.headers?.size ?? 0).toBe(0);
  const saved = await FolioDocxReviewer.fromBuffer(await firstSave);
  expect(saved.listStories().filter(({ handle }) => handle.type === "header")).toHaveLength(1);
  expect(saved.readStory(first)?.text).toBe("");
  expect(source.readStory(sourceHandle)?.text).toBe("Header text");
});

test("comparison stages a new story's inherited formatting from another package", async () => {
  const document = sourceDocument();
  if (!document.package.styles) throw new Error("Missing fixture styles");
  document.package.styles = {
    ...document.package.styles,
    docDefaults: { rPr: { bold: true } },
  };
  const source = await FolioDocxReviewer.fromBuffer(await createDocx(document));
  const sourceHandle = source.listStories().find(({ handle }) => handle.type === "header")?.handle;
  if (!sourceHandle || sourceHandle.type !== "header") throw new Error("Missing source header");
  const destination = await FolioDocxReviewer.fromBuffer(await createDocx(createEmptyDocument()));
  const access = getFolioDocxComparisonAccess(destination);
  expect(await access.createComparisonHeaderFooter(source, sourceHandle)).not.toBeNull();
  const result = access.stageTargetStyles(source, [], [source.snapshotStory(sourceHandle)]);
  expect(result.status).toBe("imported");
  expect(destination.listStories()).toHaveLength(2);
});

test("comparison rebinds an imported header's external hyperlink relationship", async () => {
  const target = sourceDocument();
  const header = target.package.headers?.values().next().value;
  if (!header) throw new Error("Missing target header");
  header.content = [
    {
      type: "paragraph",
      content: [
        {
          type: "hyperlink",
          href: "https://example.test/imported-header-link",
          tooltip: "Imported header link",
          children: [{ type: "run", content: [{ type: "text", text: "Linked header" }] }],
        },
      ],
    },
  ];
  const result = await compareDocx(
    await createDocx(createEmptyDocument()),
    await createDocx(target),
    {
      author: "Reviewer",
      timestamp: "2026-09-13T00:00:00.000Z",
      revisionFormat: "folio-exact",
    },
  );
  if (result.isErr()) throw result.error;
  expect(result.value.verification.status).toBe("verified");
  for (const decision of ["accept", "reject"] as const) {
    const reviewer = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    if (decision === "accept") {
      reviewer.acceptAll();
    } else {
      reviewer.rejectAll();
    }
    const buffer = await reviewer.toBuffer();
    const reopened = await FolioDocxReviewer.fromBuffer(buffer);
    expect(reopened.toDocument().package.headers?.size ?? 0).toBe(decision === "accept" ? 1 : 0);
    if (decision === "reject") continue;
    const zip = await JSZip.loadAsync(buffer);
    const rels = Object.entries(zip.files).find(([path]) =>
      /^word\/_rels\/header\d+\.xml\.rels$/u.test(path),
    );
    if (!rels) throw new Error("Missing imported header relationships");
    const relationship = [...parseRelationships(await rels[1].async("text")).values()].find(
      ({ type }) => type === RELATIONSHIP_TYPES.hyperlink,
    );
    expect(relationship?.target).toBe("https://example.test/imported-header-link");
    expect(relationship?.targetMode).toBe("External");
  }
});

test("comparison preserves an imported footer's empty hyperlink container", async () => {
  const target = sourceDocument();
  const header = target.package.headers?.values().next().value;
  if (!header) throw new Error("Missing target header");
  target.package.headers = undefined;
  target.package.document.finalSectionProperties = {
    ...target.package.document.finalSectionProperties,
    headerReferences: undefined,
    footerReferences: [{ type: "default", rId: "rId_footer" }],
  };
  target.package.footers = new Map([
    [
      "rId_footer",
      {
        type: "footer",
        hdrFtrType: "default",
        content: [
          {
            type: "paragraph",
            content: [
              {
                type: "hyperlink",
                href: "",
                children: [{ type: "run", content: [{ type: "text", text: "Container text" }] }],
              },
            ],
          },
        ],
      },
    ],
  ]);
  const result = await compareDocx(
    await createDocx(createEmptyDocument()),
    await createDocx(target),
    {
      author: "Reviewer",
      timestamp: "2026-09-13T00:00:00.000Z",
      revisionFormat: "folio-exact",
    },
  );
  if (result.isErr()) throw result.error;
  expect(result.value.verification.status).toBe("verified");
  for (const decision of ["accept", "reject"] as const) {
    const reviewer = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    if (decision === "accept") {
      reviewer.acceptAll();
    } else {
      reviewer.rejectAll();
    }
    const buffer = await reviewer.toBuffer();
    const zip = await JSZip.loadAsync(buffer);
    const footers = Object.keys(zip.files).filter((path) => /^word\/footer\d+\.xml$/u.test(path));
    expect(footers).toHaveLength(decision === "accept" ? 1 : 0);
    if (decision === "accept") {
      const footer = footers.at(0);
      if (!footer) throw new Error("Missing imported footer");
      expect(await zip.file(footer)?.async("text")).toContain("<w:hyperlink>");
    }
  }
});
