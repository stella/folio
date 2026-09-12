import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { schema } from "../../prosemirror/schema";
import type { Document } from "../../types/document";
import {
  firstProseMirrorSourceIdentityDifferencePath,
  sourceIdentityAttrDispositionFor,
  sourceIdentityGovernedAttrKeys,
  sourceIdentityGovernedNodeNames,
} from "./prosemirror-source-identity";

const documentWith = (node: PMNode): PMNode => schema.node("doc", null, [node]);

const paragraphWith = (attrs: Record<string, unknown>): PMNode =>
  documentWith(schema.node("paragraph", attrs, [schema.text("Clause text.")]));

const tableCellWith = (attrs: Record<string, unknown>): PMNode =>
  schema.node("tableCell", attrs, [schema.node("paragraph")]);

const tableRowWith = (attrs: Record<string, unknown>): PMNode =>
  schema.node("tableRow", attrs, [tableCellWith({})]);

const tableWith = (attrs: Record<string, unknown>): PMNode =>
  documentWith(schema.node("table", attrs, [tableRowWith({})]));

const textBox = (groupId: string, anchorId: string, width = 200): PMNode =>
  schema.node(
    "textBox",
    {
      width,
      _docxPlacement: "inlineWithPrevious",
      _docxGroupId: groupId,
      _docxAnchorId: anchorId,
    },
    [schema.node("paragraph")],
  );

const linkedTextBoxes = ({
  anchorIds,
  groupIds,
}: {
  anchorIds: readonly [string, string];
  groupIds: readonly [string, string];
}): PMNode =>
  schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.node("textBoxAnchor", { anchorId: anchorIds[0] }),
      schema.node("textBoxAnchor", { anchorId: anchorIds[1] }),
    ]),
    textBox(groupIds[0], anchorIds[0]),
    textBox(groupIds[1], anchorIds[1]),
  ]);

const identityAttrs = (document: PMNode): readonly [string, string] => {
  const identities: [string, string][] = [];
  document.descendants((node) => {
    if (node.type.name !== "textBox") return true;
    const groupId = node.attrs["_docxGroupId"];
    const anchorId = node.attrs["_docxAnchorId"];
    if (typeof groupId === "string" && typeof anchorId === "string") {
      identities.push([groupId, anchorId]);
    }
    return false;
  });
  const first = identities.at(0);
  if (!first) throw new Error("text-box identity fixture did not project");
  return first;
};

describe("ProseMirror source identity", () => {
  test("requires an explicit identity disposition for every governed schema attribute", () => {
    for (const nodeTypeName of sourceIdentityGovernedNodeNames()) {
      const nodeType = schema.nodes[nodeTypeName];
      if (!nodeType) throw new Error(`missing governed node type ${nodeTypeName}`);
      expect(sourceIdentityGovernedAttrKeys(nodeTypeName).toSorted()).toEqual(
        Object.keys(nodeType.spec.attrs ?? {}).toSorted(),
      );
    }
  });

  test("fails closed when a governed schema attribute has no disposition", () => {
    expect(() => sourceIdentityAttrDispositionFor("paragraph", "newSchemaAttribute")).toThrow(
      "A governed ProseMirror source attribute has no identity disposition",
    );
  });

  test("makes the serializer projection a semantic source-identity fixed point", () => {
    const paragraph = schema.node(
      "paragraph",
      {
        spacingFromDocDefaults: { after: true },
        bookmarks: [{ id: 7, name: "clause" }],
      },
      [schema.text("Clause text.")],
    );
    const cell = tableCellWith({
      backgroundColor: "C6E0B4",
      _resolvedBackgroundColor: null,
      _originalFormatting: { shading: { fill: { rgb: "C6E0B4" } } },
    });
    const row = schema.node("tableRow", { isHeader: null, hidden: null }, [cell]);
    const live = schema.node("doc", null, [paragraph, schema.node("table", null, [row])]);
    const projected = toProseDoc(fromProseDoc(live));
    const projectedAgain = toProseDoc(fromProseDoc(projected));

    expect(live.eq(projected)).toBe(false);
    expect(firstProseMirrorSourceIdentityDifferencePath(live, projected)).toBe("");
    expect(firstProseMirrorSourceIdentityDifferencePath(projected, projectedAgain)).toBe("");
  });

  test("compares paragraph provenance through its serialized formatting projection", () => {
    const withDocumentDefaultProvenance = paragraphWith({
      spacingFromDocDefaults: { after: true },
    });
    const reprojected = paragraphWith({});

    expect(
      firstProseMirrorSourceIdentityDifferencePath(withDocumentDefaultProvenance, reprojected),
    ).toBe("");

    const directSpacing = paragraphWith({ spaceAfter: 240, spacingExplicit: { after: true } });
    expect(firstProseMirrorSourceIdentityDifferencePath(reprojected, directSpacing)).toBe(
      "doc.content[0].attrs.$paragraph-document",
    );
  });

  test("equates section shorthand with its canonical Document representation", () => {
    const shorthand = paragraphWith({ sectionBreakType: "nextPage" });
    const full = paragraphWith({ _sectionProperties: { sectionStart: "nextPage" } });
    const changed = paragraphWith({ _sectionProperties: { sectionStart: "continuous" } });

    expect(firstProseMirrorSourceIdentityDifferencePath(shorthand, full)).toBe("");
    expect(firstProseMirrorSourceIdentityDifferencePath(full, changed)).toBe(
      "doc.content[0].attrs.$paragraph-section",
    );
  });

  test("compares resolved cell caches through the cell serializer projection", () => {
    const originalFormatting = { shading: { fill: { rgb: "C6E0B4" } } };
    const live = tableCellWith({
      backgroundColor: "C6E0B4",
      _resolvedBackgroundColor: null,
      _originalFormatting: originalFormatting,
    });
    const reprojected = tableCellWith({
      backgroundColor: "C6E0B4",
      _resolvedBackgroundColor: "C6E0B4",
      _originalFormatting: originalFormatting,
    });

    expect(firstProseMirrorSourceIdentityDifferencePath(live, reprojected)).toBe("");

    const changed = tableCellWith({
      backgroundColor: "FF0000",
      _resolvedBackgroundColor: "C6E0B4",
      _originalFormatting: originalFormatting,
    });
    expect(firstProseMirrorSourceIdentityDifferencePath(reprojected, changed)).toBe(
      "doc.attrs.$table-cell-document",
    );
  });

  test("canonicalizes every resolved-style cache that controls serializer authorship", () => {
    const borders = { top: { style: "single", size: 8, color: { rgb: "112233" } } };
    const margins = { top: 120 };
    const cases = [
      {
        live: tableCellWith({
          borders,
          _resolvedBorders: null,
          _originalFormatting: { borders },
        }),
        reprojected: tableCellWith({
          borders,
          _resolvedBorders: borders,
          _originalFormatting: { borders },
        }),
      },
      {
        live: tableCellWith({
          margins,
          _resolvedMargins: null,
          _originalFormatting: { margins: { top: { value: 120, type: "dxa" } } },
        }),
        reprojected: tableCellWith({
          margins,
          _resolvedMargins: margins,
          _originalFormatting: { margins: { top: { value: 120, type: "dxa" } } },
        }),
      },
      {
        live: tableWith({
          cellMargins: margins,
          _resolvedCellMargins: null,
          _originalFormatting: { cellMargins: { top: { value: 120, type: "dxa" } } },
        }),
        reprojected: tableWith({
          cellMargins: margins,
          _resolvedCellMargins: margins,
          _originalFormatting: { cellMargins: { top: { value: 120, type: "dxa" } } },
        }),
      },
    ];

    for (const { live, reprojected } of cases) {
      expect(firstProseMirrorSourceIdentityDifferencePath(live, reprojected)).toBe("");
    }
  });

  test("normalizes row defaults without hiding authored row formatting", () => {
    const nullableDefaults = tableRowWith({ isHeader: null, hidden: null });
    const schemaDefaults = tableRowWith({ isHeader: false, hidden: false });

    expect(firstProseMirrorSourceIdentityDifferencePath(nullableDefaults, schemaDefaults)).toBe("");

    const hidden = tableRowWith({ hidden: true });
    expect(firstProseMirrorSourceIdentityDifferencePath(schemaDefaults, hidden)).toBe(
      "doc.attrs.$table-row-document",
    );
  });

  test("separates table presentation caches from serialized table properties", () => {
    const cached = tableWith({ _resolvedIndent: 720, _resolvedJustification: "center" });
    const reprojected = tableWith({});

    expect(firstProseMirrorSourceIdentityDifferencePath(cached, reprojected)).toBe("");

    const changed = tableWith({ width: 1440, widthType: "dxa" });
    expect(firstProseMirrorSourceIdentityDifferencePath(reprojected, changed)).toBe(
      "doc.content[0].attrs.$table-document",
    );
  });

  test("equates paragraph bookmark attrs with their canonical inline boundaries", () => {
    const legacy = documentWith(
      schema.node("paragraph", { bookmarks: [{ id: 7, name: "clause" }] }, [
        schema.text("Clause text."),
      ]),
    );
    const canonical = documentWith(
      schema.node("paragraph", null, [
        schema.node("bookmarkBoundary", { type: "start", id: 7, name: "clause" }),
        schema.text("Clause text."),
        schema.node("bookmarkBoundary", { type: "end", id: 7 }),
      ]),
    );

    expect(firstProseMirrorSourceIdentityDifferencePath(legacy, canonical)).toBe("");

    const renamed = documentWith(
      schema.node("paragraph", null, [
        schema.node("bookmarkBoundary", { type: "start", id: 7, name: "other" }),
        schema.text("Clause text."),
        schema.node("bookmarkBoundary", { type: "end", id: 7 }),
      ]),
    );
    expect(firstProseMirrorSourceIdentityDifferencePath(legacy, renamed)).toBe(
      "doc.content[0].content[0].attrs.name",
    );
  });

  test("keeps transport identity exact while canonicalizing presentation", () => {
    const left = paragraphWith({ paraId: "00000001", _tableOfContentsLevel: 1 });
    const reprojected = paragraphWith({ paraId: "00000001", _tableOfContentsLevel: null });
    const otherSource = paragraphWith({ paraId: "00000002", _tableOfContentsLevel: null });

    expect(firstProseMirrorSourceIdentityDifferencePath(left, reprojected)).toBe("");
    expect(firstProseMirrorSourceIdentityDifferencePath(reprojected, otherSource)).toBe(
      "doc.content[0].attrs.paraId",
    );
  });

  test("alpha-compares per-load text-box identities while preserving their links", () => {
    const left = linkedTextBoxes({
      anchorIds: ["left:0:0", "left:0:1"],
      groupIds: ["left:0", "left:0"],
    });
    const right = linkedTextBoxes({
      anchorIds: ["right:0:0", "right:0:1"],
      groupIds: ["right:0", "right:0"],
    });

    expect(firstProseMirrorSourceIdentityDifferencePath(left, right)).toBe("");
  });

  test("rejects a changed text-box grouping equivalence class", () => {
    const left = linkedTextBoxes({
      anchorIds: ["left:0:0", "left:0:1"],
      groupIds: ["left:0", "left:0"],
    });
    const right = linkedTextBoxes({
      anchorIds: ["right:0:0", "right:1:0"],
      groupIds: ["right:0", "right:1"],
    });

    expect(firstProseMirrorSourceIdentityDifferencePath(left, right)).toBe(
      "doc.content[2].attrs._docxGroupId",
    );
  });

  test("rejects a text box linked to the wrong inline anchor", () => {
    const left = linkedTextBoxes({
      anchorIds: ["left:0:0", "left:0:1"],
      groupIds: ["left:0", "left:0"],
    });
    const right = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node("textBoxAnchor", { anchorId: "right:0:0" }),
        schema.node("textBoxAnchor", { anchorId: "right:0:1" }),
      ]),
      textBox("right:0", "right:0:1"),
      textBox("right:0", "right:0:0"),
    ]);

    expect(firstProseMirrorSourceIdentityDifferencePath(left, right)).toBe(
      "doc.content[1].attrs._docxAnchorId",
    );
  });

  test("compares every semantic text-box attribute exactly", () => {
    const left = linkedTextBoxes({
      anchorIds: ["left:0:0", "left:0:1"],
      groupIds: ["left:0", "left:0"],
    });
    const right = schema.node("doc", null, [
      schema.node("paragraph", null, [
        schema.node("textBoxAnchor", { anchorId: "right:0:0" }),
        schema.node("textBoxAnchor", { anchorId: "right:0:1" }),
      ]),
      textBox("right:0", "right:0:0", 201),
      textBox("right:0", "right:0:1"),
    ]);

    expect(firstProseMirrorSourceIdentityDifferencePath(left, right)).toBe(
      "doc.content[1].attrs.width",
    );
  });

  test("rejects malformed nominal identities even when both sides match", () => {
    const malformed = schema.node("doc", null, [
      schema.node("paragraph", null, [schema.node("textBoxAnchor", { anchorId: "" })]),
      textBox("", ""),
    ]);

    expect(firstProseMirrorSourceIdentityDifferencePath(malformed, malformed)).toBe(
      "doc.content[0].content[0].attrs.anchorId",
    );
  });

  test("keeps independent-load collision isolation while comparing source identity", () => {
    const document: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [
                    {
                      type: "shape",
                      shape: {
                        type: "shape",
                        shapeType: "textBox",
                        size: { width: 914_400, height: 914_400 },
                        textBody: { content: [{ type: "paragraph", content: [] }] },
                      },
                    },
                  ],
                },
              ],
            },
          ],
        },
      },
    };
    const left = toProseDoc(document);
    const right = toProseDoc(document);

    expect(identityAttrs(left)).not.toEqual(identityAttrs(right));
    expect(firstProseMirrorSourceIdentityDifferencePath(left, right)).toBe("");
  });

  test("accepts every bijective renaming of group and anchor atoms", () => {
    fc.assert(
      fc.property(
        fc.array(fc.integer({ min: 0, max: 3 }), { minLength: 1, maxLength: 8 }),
        (groups) => {
          const build = (prefix: string) =>
            schema.node("doc", null, [
              schema.node(
                "paragraph",
                null,
                groups.map((_, index) =>
                  schema.node("textBoxAnchor", { anchorId: `${prefix}:anchor:${String(index)}` }),
                ),
              ),
              ...groups.map((group, index) =>
                textBox(`${prefix}:group:${String(group)}`, `${prefix}:anchor:${String(index)}`),
              ),
            ]);

          expect(firstProseMirrorSourceIdentityDifferencePath(build("left"), build("right"))).toBe(
            "",
          );
        },
      ),
      { numRuns: 100 },
    );
  });
});
