import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import {
  NO_PARAGRAPH_NUMBERING,
  type ParagraphNumberingOverride,
  paragraphNumberingReference,
} from "@stll/docx-core/model";

import type {
  Paragraph,
  ParagraphPropertyChange,
  Run,
  RunPropertyChange,
} from "../../types/content";
import { createEmptyDocument } from "../../utils/createDocument";
import { CLEARED_LIST_RENDERING_ATTRS } from "../listMarker";
import {
  LIST_RENDERING_ATTR_KEYS,
  type ListRenderingAttrKey,
} from "../listRenderingAttrs";
import { paragraphNumberingAttr } from "../numberingAttr";
import type { ParagraphAttrs } from "../schema/nodes";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const recordedRendering = {
  listMarker: "1.",
  listMarkerTemplate: "%1.",
  listFoldedMarkerSuffix: "(a)",
  listIsBullet: false,
  listIsLegal: true,
  listNumFmt: "decimal",
  listMarkerHidden: false,
  listMarkerFormatting: { bold: true },
  listMarkerAlignment: "right",
  listMarkerAllCaps: true,
  listMarkerSuffix: "space",
  listLevelTabs: [],
  listLevelNumFmts: ["decimal"],
  listLevelStarts: [1],
  listAbstractNumId: 7,
  listStartOverride: 3,
  listImplicitChildLevelAdvances: 2,
  listMarkerSecondSlotOffsetTwips: 720,
} satisfies { [K in ListRenderingAttrKey]-?: NonNullable<ParagraphAttrs[K]> };

const authoredNumberingStates = [
  undefined,
  { kind: "levelOnly", ilvl: 2 } satisfies ParagraphNumberingOverride,
  paragraphNumberingReference({ numId: 42 }),
  NO_PARAGRAPH_NUMBERING,
];

const historicalRenderingStates = [
  { state: "absent", attrs: {}, model: {} },
  { state: "cleared", attrs: CLEARED_LIST_RENDERING_ATTRS, model: {} },
  { state: "rendered", attrs: recordedRendering, model: recordedRendering },
];

const samplePropertyChange: ParagraphPropertyChange = {
  type: "paragraphPropertyChange",
  info: {
    id: 7,
    author: "Reviewer",
    date: "2026-05-15T12:00:00Z",
  },
  previousFormatting: {
    alignment: "left",
    runProperties: {
      fontFamily: { cs: "Noto Naskh Arabic", csTheme: "majorBidi" },
      language: { bidi: "ar-SA" },
      spacing: 12,
    },
  },
  currentFormatting: {
    alignment: "center",
    runProperties: {
      fontFamily: { cs: "Noto Naskh Arabic", csTheme: "majorBidi" },
      language: { bidi: "ar-SA" },
      spacing: 24,
    },
  },
};

function paragraphWithPropertyChange(): Paragraph {
  return {
    type: "paragraph",
    formatting: { alignment: "center" },
    propertyChanges: [samplePropertyChange],
    content: [
      {
        type: "run",
        formatting: {},
        content: [{ type: "text", text: "body" }],
      },
    ],
  };
}

describe("paragraph propertyChanges PM round-trip", () => {
  test("toProseDoc copies propertyChanges into the paragraph node attrs", () => {
    const document = {
      package: {
        document: {
          content: [paragraphWithPropertyChange()],
          finalSectionProperties: {},
        },
      },
    };
    const pmDoc = toProseDoc(document as never);
    // The paragraph PM node now carries the entries on a private attr —
    // editor code never reads it; it exists purely so fromProseDoc can
    // restore them after an edit.
    const paragraphNode = pmDoc.firstChild!;
    expect(paragraphNode.attrs["_propertyChanges"]).toEqual([
      {
        ...samplePropertyChange,
        previousFormatting: {
          ...samplePropertyChange.previousFormatting,
          numPrFromStyle: null,
        },
      },
    ]);
  });

  test("fromProseDoc restores propertyChanges back onto the Folio Paragraph", () => {
    // Simulate a no-op edit: convert to PM and immediately back. The
    // entries must survive even though the editor surfaces nothing in
    // UI for them — the previous behaviour silently stripped them on
    // every edit, corrupting the `w:pPrChange` history.
    const original = paragraphWithPropertyChange();
    const document = {
      package: {
        document: {
          content: [original],
          finalSectionProperties: {},
        },
      },
    };
    const pmDoc = toProseDoc(document as never);
    const roundtripped = fromProseDoc(pmDoc, undefined, {
      stylesheetSource: { type: "package" },
    }).package.document.content[0] as Paragraph | undefined;

    expect(roundtripped?.propertyChanges).toEqual([samplePropertyChange]);
  });

  test.each(historicalRenderingStates)(
    "keeps historical rendering state $state separate from numbering source tiers",
    ({ attrs: renderingAttrs, model: modelRendering }) => {
      for (const inherited of [
        undefined,
        paragraphNumberingReference({ numId: 41, ilvl: 2 }),
      ]) {
        for (const authored of authoredNumberingStates) {
          const source = createEmptyDocument();
          source.package.document.content = [paragraphWithPropertyChange()];
          const pmDoc = toProseDoc(source);
          const paragraph = pmDoc.firstChild;
          if (!paragraph) panic("expected a paragraph");
          const previousFormatting = {
            ...samplePropertyChange.previousFormatting,
            ...renderingAttrs,
            numPr:
              authored === undefined ? null : paragraphNumberingAttr(authored),
            numPrFromStyle:
              inherited === undefined
                ? null
                : paragraphNumberingAttr(inherited),
          };
          const changed = pmDoc.type.create(
            pmDoc.attrs,
            paragraph.type.create(
              {
                ...paragraph.attrs,
                _propertyChanges: [
                  { ...samplePropertyChange, previousFormatting },
                ],
              },
              paragraph.content,
            ),
          );
          const restored = fromProseDoc(changed, undefined, {
            stylesheetSource: { type: "package" },
          }).package.document.content.at(0);
          if (restored?.type !== "paragraph")
            panic("expected a restored paragraph");
          expect(restored.propertyChanges).toEqual([
            {
              ...samplePropertyChange,
              previousFormatting: {
                ...samplePropertyChange.previousFormatting,
                ...modelRendering,
                ...(authored === undefined ? {} : { numPr: authored }),
                ...(inherited === undefined
                  ? {}
                  : { numPrFromStyle: inherited }),
              },
            },
          ]);
          for (const key of LIST_RENDERING_ATTR_KEYS) {
            expect(
              Object.hasOwn(
                restored.propertyChanges?.at(0)?.previousFormatting ?? {},
                key,
              ),
            ).toBe(Object.hasOwn(modelRendering, key));
          }
        }
      }
    },
  );

  test("keeps nested property-change formatting fixed across repeated PM decoding", () => {
    const original = paragraphWithPropertyChange();
    const first = fromProseDoc(
      toProseDoc({
        package: {
          document: {
            content: [original],
            finalSectionProperties: {},
          },
        },
      } as never),
      undefined,
      { stylesheetSource: { type: "package" } },
    ).package.document.content[0] as Paragraph;
    const second = fromProseDoc(
      toProseDoc({
        package: {
          document: {
            content: [first],
            finalSectionProperties: {},
          },
        },
      } as never),
      undefined,
      { stylesheetSource: { type: "package" } },
    ).package.document.content[0] as Paragraph;

    expect(second.propertyChanges).toEqual(first.propertyChanges);
  });

  test("paragraphs without propertyChanges round-trip without inventing them", () => {
    const plain: Paragraph = {
      type: "paragraph",
      formatting: {},
      content: [
        {
          type: "run",
          formatting: {},
          content: [{ type: "text", text: "plain" }],
        },
      ],
    };
    const document = {
      package: {
        document: {
          content: [plain],
          finalSectionProperties: {},
        },
      },
    };
    const pmDoc = toProseDoc(document as never);
    const roundtripped = fromProseDoc(pmDoc, undefined, {
      stylesheetSource: { type: "package" },
    }).package.document.content[0] as Paragraph | undefined;
    // No phantom propertyChanges should be attached.
    expect(roundtripped?.propertyChanges).toBeUndefined();
  });
});

const sampleRunPropertyChange: RunPropertyChange = {
  type: "runPropertyChange",
  info: {
    id: 8,
    author: "Reviewer",
    date: "2026-07-17T08:00:00Z",
  },
  previousFormatting: { italic: true },
  currentFormatting: { bold: true },
};

describe("run propertyChanges PM round-trip", () => {
  test("preserves run formatting revisions through the editable model", () => {
    const run: Run = {
      type: "run",
      formatting: { bold: true },
      propertyChanges: [sampleRunPropertyChange],
      content: [{ type: "text", text: "reviewed" }],
    };
    const document = {
      package: {
        document: {
          content: [{ type: "paragraph" as const, content: [run] }],
          finalSectionProperties: {},
        },
      },
    };

    const pmDoc = toProseDoc(document as never);
    const text = pmDoc.firstChild?.firstChild;
    expect(
      text?.marks.find((mark) => mark.type.name === "runPropertyChange")?.attrs,
    ).toEqual({
      changes: [sampleRunPropertyChange],
      // Provenance defaults: a parsed run-property change is always a user change.
      provenance: "user",
      suggestionId: null,
    });

    const roundtripped = fromProseDoc(pmDoc, undefined, {
      stylesheetSource: { type: "package" },
    }).package.document.content.at(0);
    const roundtrippedRun =
      roundtripped?.type === "paragraph"
        ? roundtripped.content.find((content) => content.type === "run")
        : undefined;
    expect(roundtrippedRun?.propertyChanges).toEqual([sampleRunPropertyChange]);
  });

  test("does not invent run formatting revisions", () => {
    const document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph" as const,
              content: [
                {
                  type: "run" as const,
                  formatting: { bold: true },
                  content: [{ type: "text" as const, text: "plain" }],
                },
              ],
            },
          ],
          finalSectionProperties: {},
        },
      },
    };

    const roundtripped = fromProseDoc(
      toProseDoc(document as never),
      undefined,
      {
        stylesheetSource: { type: "package" },
      },
    ).package.document.content.at(0);
    const run =
      roundtripped?.type === "paragraph"
        ? roundtripped.content.find((content) => content.type === "run")
        : undefined;
    expect(run?.propertyChanges).toBeUndefined();
  });
});
