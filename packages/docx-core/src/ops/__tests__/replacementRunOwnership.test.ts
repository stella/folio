import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertExactModel } from "../../../../../test/exactModel";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  Document,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
} from "../../model/document";
import {
  type AppliedDocumentOp,
  applyDocumentOp,
  applyDocumentOps,
} from "../apply";
import {
  allocateEditorIntentIds,
  compileEditorIntent,
  type EditorIntent,
} from "../editorIntent";
import { paragraphLength } from "../offsets";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";
import { assertFreshIdentityEquivalent } from "./freshIdentityOracle";

setDefaultTimeout(propertyTestTimeout(30_000));

const PAYLOADS = {
  tab: [{ type: "tab" }],
  lineBreak: [{ type: "break", breakType: "textWrapping" }],
  text: [{ type: "text", text: "a😀" }],
} as const satisfies Record<string, readonly RunContent[]>;

type FixtureOptions = {
  leadingBreaks: number;
  trailingBreaks: number;
  prefix: string;
  suffix: string;
  payload: (typeof PAYLOADS)[keyof typeof PAYLOADS];
  layout: "shared" | "separate";
  fields: "plain" | "preserved";
};

const fixture = ({
  leadingBreaks,
  trailingBreaks,
  prefix,
  suffix,
  payload,
  layout,
  fields,
}: FixtureOptions) => {
  const leading: RunContent[] = [
    ...(prefix === "" ? [] : [{ type: "text" as const, text: prefix }]),
    ...Array.from({ length: leadingBreaks }, () => ({
      type: "renderedPageBreak" as const,
    })),
  ];
  const trailing: RunContent[] = [
    ...Array.from({ length: trailingBreaks }, () => ({
      type: "renderedPageBreak" as const,
    })),
    ...(suffix === "" ? [] : [{ type: "text" as const, text: suffix }]),
  ];
  const run = (content: RunContent[]) =>
    ({
      type: "run",
      ...(fields === "preserved"
        ? {
            preservedAttributes: [
              {
                namespace:
                  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                name: "rsidR",
                value: "00FF0011",
              },
            ],
          }
        : {}),
      content,
    }) satisfies Run;
  const selected = [...payload];
  const content: ParagraphContent[] =
    layout === "shared"
      ? [run([...leading, ...selected, ...trailing])]
      : [leading, selected, trailing]
          .filter((items) => items.length > 0)
          .map(run);
  const paragraph = {
    type: "paragraph",
    paraId: "00000001",
    content,
  } satisfies Paragraph;
  const properties = { pageWidth: 12240, pageHeight: 15840 };
  const document = {
    package: {
      document: {
        content: [paragraph],
        sections: [{ properties, content: [paragraph] }],
        finalSectionProperties: properties,
      },
      numbering: undefined,
      settings: { defaultTabStop: 720 },
    },
    warnings: ["kept"],
  } satisfies Document;
  const from = {
    story: OP_STORIES.MAIN,
    blockId: paragraph.paraId,
    offset: prefix.length,
  };
  const width = paragraphLength({
    type: "paragraph",
    content: [run(selected)],
  });
  return {
    document,
    intent: {
      type: "replaceText",
      from,
      to: { ...from, offset: prefix.length + width, zeroWidthBefore: 0 },
      text: "edited",
    } as const satisfies EditorIntent,
  };
};

const assertUndo = (edit: AppliedDocumentOp, original: Document) => {
  const immediate = applyDocumentOps(edit.document, edit.inverse).unwrap();
  assertExactModel(immediate.document, original);
  const restored = applyDocumentOps(
    edit.document,
    JSON.parse(JSON.stringify(edit.inverse)),
  ).unwrap();
  assertExactModel(restored.document, original);
  const redone = applyDocumentOps(
    restored.document,
    JSON.parse(JSON.stringify(restored.inverse)),
  ).unwrap();
  assertExactModel(redone.document, edit.document);
};

const assertReplacement = (options: FixtureOptions) => {
  const { document, intent } = fixture(options);
  const before = structuredClone(document);
  const allocation = allocateEditorIntentIds(document, intent);
  const direct = compileEditorIntent(document, {
    intent,
    mode: { type: "editing", newIds: allocation.newIds },
  }).unwrap();
  const tracked = compileEditorIntent(document, {
    intent,
    mode: {
      type: "suggesting",
      revision: {
        id: allocation.revisionId,
        author: "Reviewer",
        date: "2026-01-01T00:00:00Z",
      },
      newIds: allocation.newIds,
    },
  }).unwrap();
  const directEdit = applyDocumentOps(document, direct.ops).unwrap();
  const trackedEdit = applyDocumentOps(document, tracked.ops).unwrap();
  const accepted = applyDocumentOp(trackedEdit.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [...trackedEdit.revisions],
    decision: REVISION_DECISIONS.ACCEPT,
  }).unwrap();
  const rejected = applyDocumentOp(trackedEdit.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [...trackedEdit.revisions],
    decision: REVISION_DECISIONS.REJECT,
  }).unwrap();
  assertFreshIdentityEquivalent({
    actual: accepted.document,
    expected: directEdit.document,
    original: document,
    allocated: allocation.newIds,
  });
  assertExactModel(rejected.document, document);
  for (const edit of [directEdit, trackedEdit]) assertUndo(edit, document);
  for (const edit of [accepted, rejected])
    assertUndo(edit, trackedEdit.document);
  assertExactModel(document, before);
  return { direct: directEdit.document, accepted: accepted.document };
};

describe("replacement source-run ownership", () => {
  test("rejoins the source run cut around a selected tab and cached breaks", () => {
    const result = assertReplacement({
      leadingBreaks: 1,
      trailingBreaks: 1,
      prefix: "",
      suffix: "",
      payload: PAYLOADS.tab,
      layout: "shared",
      fields: "plain",
    });
    expect(result.direct.package.document.content.at(0)).toEqual({
      type: "paragraph",
      paraId: "00000001",
      content: [
        { type: "run", content: [{ type: "renderedPageBreak" }] },
        {
          type: "run",
          content: [
            { type: "text", text: "edited" },
            { type: "renderedPageBreak" },
          ],
        },
      ],
    });
  });

  test("preserves source ownership across cached-break placements, selected units and authored run edges", () => {
    assertProperty(
      fc.property(
        fc.record({
          leadingBreaks: fc.integer({ min: 0, max: 3 }),
          trailingBreaks: fc.integer({ min: 0, max: 3 }),
          prefix: fc.constantFrom("", "a😀"),
          suffix: fc.constantFrom("", "b😀"),
          payload: fc.constantFrom(...Object.values(PAYLOADS)),
          layout: fc.constantFrom("shared", "separate"),
          fields: fc.constantFrom("plain", "preserved"),
        }),
        (options) => {
          assertReplacement(options);
        },
      ),
      { numRuns: 40, id: "replacement source-run ownership" },
    );
  });

  test("retains a plain prefix merge before restoring the cached suffix cut", () => {
    assertReplacement({
      leadingBreaks: 0,
      trailingBreaks: 1,
      prefix: "a😀",
      suffix: "",
      payload: PAYLOADS.tab,
      layout: "shared",
      fields: "plain",
    });
  });

  test("keeps separately authored cached-break runs and unequal source attributes separate", () => {
    const controls = {
      authoredEdges: { layout: "separate", fields: "plain" },
      sourceAttributes: { layout: "shared", fields: "preserved" },
    } as const;
    const exercised = new Set<string>();
    for (const [name, control] of Object.entries(controls)) {
      const result = assertReplacement({
        leadingBreaks: 1,
        trailingBreaks: 1,
        prefix: "",
        suffix: "",
        payload: PAYLOADS.tab,
        ...control,
      });
      const fields =
        control.fields === "preserved"
          ? {
              preservedAttributes: [
                {
                  namespace:
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                  name: "rsidR",
                  value: "00FF0011",
                },
              ],
            }
          : {};
      expect(result.direct.package.document.content.at(0)).toEqual({
        type: "paragraph",
        paraId: "00000001",
        content: [
          { type: "run", ...fields, content: [{ type: "renderedPageBreak" }] },
          { type: "run", content: [{ type: "text", text: "edited" }] },
          { type: "run", ...fields, content: [{ type: "renderedPageBreak" }] },
        ],
      });
      exercised.add(name);
    }
    expect(exercised).toEqual(new Set(Object.keys(controls)));
  });
});
