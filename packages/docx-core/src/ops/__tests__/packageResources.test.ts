import { expect, test } from "bun:test";
import fc from "fast-check";
import { assertExactModel } from "../../../../../test/exactModel";
import { assertProperty } from "../../../../../test/property-testing";

import type { Document } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { packageResourcesOf } from "../packageResources";
import { DOCUMENT_OP_TYPES, type DocumentOp, type SetPackageResourcesOp } from "../types";
import { captureDocumentOp } from "../wire";

const jsonTransport = <Value>(value: Value): Value => JSON.parse(JSON.stringify(value));

const document: Document = {
  package: { document: { content: [{ type: "paragraph", paraId: "00000001", content: [] }] } },
};

const imageOp = (input: Document, bytes: readonly number[]): SetPackageResourcesOp => ({
  type: DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES,
  expected: packageResourcesOf(input),
  resources: {
    ...packageResourcesOf(input),
    relationships: {
      type: "present",
      value: [
        [
          "rId1",
          {
            id: "rId1",
            type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            target: "media/clipboard.png",
          },
        ],
      ],
    },
    media: {
      type: "present",
      value: [
        [
          "word/media/clipboard.png",
          {
            path: "word/media/clipboard.png",
            mimeType: "image/png",
            data: bytes,
          },
        ],
      ],
    },
  },
});

test("binary resource operations survive JSON and undo restores exact package field presence", () => {
  assertProperty(
    fc.property(
      fc.array(fc.uint8Array({ maxLength: 128 }), { minLength: 8, maxLength: 16 }),
      fc.boolean(),
      fc.boolean(),
      (steps, explicit, nestedUndefined) => {
        const input: Document = explicit
          ? {
              ...document,
              package: {
                ...document.package,
                styles: undefined,
                numbering: undefined,
                media: undefined,
                relationships: undefined,
              },
            }
          : document;
        const initial: Document = nestedUndefined
          ? {
              ...input,
              package: {
                ...input.package,
                styles: { styles: [], docDefaults: { rPr: { bold: undefined } } },
              },
            }
          : input;
        const original = structuredClone(initial);
        let current = initial;
        const journal: {
          before: Document;
          after: Document;
          op: DocumentOp;
          inverse: readonly DocumentOp[];
        }[] = [];
        for (const bytes of steps) {
          const before = current;
          const snapshot = structuredClone(before);
          const op = jsonTransport(captureDocumentOp(imageOp(before, [...bytes])));
          const forward = applyDocumentOp(before, op).unwrap();
          assertExactModel(before, snapshot);
          const media = forward.document.package.media?.get("word/media/clipboard.png");
          if (media === undefined)
            throw new TypeError("A valid resource import produced no media.");
          assertExactModel(new Uint8Array(media.data), bytes);
          const inverse = jsonTransport(forward.inverse.map(captureDocumentOp));
          const undo = applyDocumentOps(forward.document, inverse).unwrap();
          assertExactModel(undo.document, before);
          const redo = applyDocumentOps(
            undo.document,
            jsonTransport(undo.inverse.map(captureDocumentOp)),
          ).unwrap();
          assertExactModel(redo.document, forward.document);
          journal.push({ before, after: forward.document, op, inverse });
          current = forward.document;
        }
        const final = current;
        for (const entry of journal.toReversed()) {
          current = applyDocumentOps(current, entry.inverse).unwrap().document;
          assertExactModel(current, entry.before);
        }
        assertExactModel(current, original);
        for (const entry of journal) {
          current = applyDocumentOp(current, entry.op).unwrap().document;
          assertExactModel(current, entry.after);
        }
        assertExactModel(current, final);
        assertExactModel(initial, original);
      },
    ),
    { numRuns: 30 },
  );
});

test("byte mutations make a package resource inverse stale", () => {
  const forward = applyDocumentOp(document, imageOp(document, [1, 2, 3])).unwrap();
  const media = forward.document.package.media?.get("word/media/clipboard.png");
  if (media === undefined) throw new Error("Missing test media");
  new Uint8Array(media.data)[1] = 99;
  const undo = applyDocumentOps(forward.document, forward.inverse);
  expect(undo.isErr()).toBe(true);
  if (undo.isErr()) expect(undo.error.reason).toBe("stale");
});

test("invalid or duplicate resource imports and a later content refusal are atomic", () => {
  const valid = imageOp(document, [1]);
  const media = valid.resources.media;
  if (media.type !== "present") throw new Error("Missing test media state");
  const bad: SetPackageResourcesOp[] = [
    {
      ...valid,
      resources: {
        ...valid.resources,
        media: { type: "present", value: [...media.value, ...media.value] },
      },
    },
    imageOp(document, [-1]),
    { ...valid, resources: { ...valid.resources, media: { type: "omitted" } } },
    {
      ...valid,
      resources: {
        ...valid.resources,
        numbering: {
          type: "present",
          value: {
            abstractNums: [],
            nums: [{ numId: 1, abstractNumId: 42 }],
          },
        },
      },
    },
  ];
  for (const op of bad) expect(applyDocumentOp(document, op).isErr()).toBe(true);
  expect(
    applyDocumentOps(document, [
      valid,
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: { story: "main", blockId: "00000001", offset: 20 },
        text: "x",
        runProps: "inherit",
      },
    ]).isErr(),
  ).toBe(true);
  expect(document.package.media).toBeUndefined();
  expect(document.package.relationships).toBeUndefined();
});

test("resource undo refuses to strand references introduced by later content edits", () => {
  const forward = applyDocumentOp(document, imageOp(document, [1])).unwrap();
  const changed: Document = {
    ...forward.document,
    package: {
      ...forward.document.package,
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [{ type: "hyperlink", rId: "rId1", children: [] }],
          },
        ],
      },
    },
  };
  const undone = applyDocumentOps(changed, forward.inverse);
  expect(undone.isErr()).toBe(true);
  if (undone.isErr()) expect(undone.error.reason).toBe("stale");
});

test("resource removal checks every typed style-link alias in retained definitions", () => {
  for (const key of ["basedOn", "next", "link"] as const) {
    const retained = { styleId: "Consumer", type: "paragraph" as const, [key]: "Target" };
    const input: Document = {
      ...document,
      package: {
        ...document.package,
        styles: { styles: [{ styleId: "Target", type: "paragraph" }, retained] },
      },
    };
    const expected = packageResourcesOf(input);
    const removed = applyDocumentOp(input, {
      type: DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES,
      expected,
      resources: { ...expected, styles: { type: "present", value: { styles: [retained] } } },
    });
    expect(removed.isErr()).toBe(true);
    if (removed.isErr()) expect(removed.error.reason).toBe("stale");
  }
  for (const key of ["pStyle", "numStyleLink", "styleLink"] as const) {
    const abstract = {
      abstractNumId: 1,
      levels: [
        {
          ilvl: 0,
          numFmt: "decimal" as const,
          lvlText: "%1.",
          ...(key === "pStyle" ? { pStyle: "Target" } : {}),
        },
      ],
      ...(key === "pStyle" ? {} : { [key]: "Target" }),
    };
    const input: Document = {
      ...document,
      package: {
        ...document.package,
        styles: { styles: [{ styleId: "Target", type: "paragraph" }] },
        numbering: { abstractNums: [abstract], nums: [{ numId: 1, abstractNumId: 1 }] },
      },
    };
    const expected = packageResourcesOf(input);
    const removed = applyDocumentOp(input, {
      type: DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES,
      expected,
      resources: { ...expected, styles: { type: "omitted" } },
    });
    expect(removed.isErr()).toBe(true);
    if (removed.isErr()) expect(removed.error.reason).toBe("stale");
  }
});

test("nested image-link references prevent resource undo while header-local identities do not", () => {
  const forward = applyDocumentOp(document, imageOp(document, [1])).unwrap();
  const linked: Document = {
    ...forward.document,
    package: {
      ...forward.document.package,
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              {
                type: "hyperlink",
                href: "https://example.org",
                children: [
                  {
                    type: "run",
                    content: [
                      {
                        type: "drawing",
                        image: { type: "image", hlinkRId: "rId1" },
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    },
  };
  const refused = applyDocumentOps(linked, forward.inverse);
  expect(refused.isErr()).toBe(true);
  if (refused.isErr()) expect(refused.error.reason).toBe("stale");
  const separateScope: Document = {
    ...forward.document,
    package: {
      ...forward.document.package,
      headers: new Map([
        [
          "rIdHeader",
          {
            type: "header",
            hdrFtrType: "default",
            content: [],
            watermark: { kind: "picture", imageRId: "rId1" },
          },
        ],
      ]),
    },
  };
  const restored = applyDocumentOps(separateScope, forward.inverse).unwrap();
  expect(restored.document.package.headers).toBe(separateScope.package.headers);
  expect(restored.document.package.relationships).toBeUndefined();
});

test("new style and numbering imports refuse missing dependencies without importing partial resources", () => {
  for (const key of ["basedOn", "next", "link"] as const) {
    const expected = packageResourcesOf(document);
    const result = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES,
      expected,
      resources: {
        ...expected,
        styles: {
          type: "present",
          value: { styles: [{ styleId: "Consumer", type: "paragraph", [key]: "Missing" }] },
        },
      },
    });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.reason).toBe("structureMismatch");
    expect(document.package.styles).toBeUndefined();
  }
  const expected = packageResourcesOf(document);
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES,
    expected,
    resources: {
      ...expected,
      numbering: {
        type: "present",
        value: {
          abstractNums: [{ abstractNumId: 1, styleLink: "Missing", levels: [] }],
          nums: [{ numId: 1, abstractNumId: 1 }],
        },
      },
    },
  });
  expect(result.isErr()).toBe(true);
  if (result.isErr()) expect(result.error.reason).toBe("structureMismatch");
  expect(document.package.numbering).toBeUndefined();
});
