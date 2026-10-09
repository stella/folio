import { expect, test } from "bun:test";
import fc from "fast-check";
import { EditorState, type Command } from "prosemirror-state";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import type { Node as PMNode } from "prosemirror-model";
import type {
  ParagraphContent,
  RunContent,
  SimpleField,
  TrackedRunChange,
} from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import {
  acceptAIEditRevision,
  rejectAIEditRevision,
  resolveAllChangesInHeadlessState,
} from "../prosemirror/commands/comments";
import { expectTrackedChangeMarkAttrs } from "../prosemirror/attrs";
import { createDocx } from "./rezip";
import { parseDocx } from "./parser";
import { toMarkdown } from "../markdown/index";
import { serializeParagraph } from "./serializer/paragraphSerializer";
import {
  getChildElements,
  getLocalName,
  getTextContent,
  getAttributeByNamespaceUri,
  WORDPROCESSINGML_NAMESPACE_URIS,
  parseXmlDocument,
} from "./xmlParser";
import { TRACKED_RUN_WORDPROCESSING_CHILDREN } from "./containerChildren.gen";

const applyReviewCommand = (state: EditorState, command: Command) => {
  let resolved = state;
  command(state, (transaction) => {
    resolved = state.apply(transaction);
  });
  return resolved;
};

const FIELD_DATA =
  '<w:fldData xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">AQID</w:fldData>';

const expectFieldDataHome = (source: ReturnType<typeof createEmptyDocument>) => {
  const paragraph = source.package.document.content.at(0);
  expect(paragraph?.type).toBe("paragraph");
  if (paragraph?.type !== "paragraph") return;
  const xml = serializeParagraph(paragraph);
  const root = parseXmlDocument(
    `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</root>`,
  );
  expect(root).not.toBeNull();
  if (!root) return;
  const pending = [root];
  const revisionChildren: ReadonlySet<string> = new Set(TRACKED_RUN_WORDPROCESSING_CHILDREN);
  let count = 0;
  while (pending.length > 0) {
    const parent = pending.pop();
    if (!parent) continue;
    const name = getLocalName(parent.name);
    for (const child of getChildElements(parent)) {
      if (["ins", "del", "moveFrom", "moveTo"].includes(name))
        expect(revisionChildren.has(getLocalName(child.name))).toBe(true);
      if (getLocalName(child.name) === "fldData") {
        expect(name).toBe("fldChar");
        expect(
          getAttributeByNamespaceUri(parent, WORDPROCESSINGML_NAMESPACE_URIS, "fldCharType"),
        ).toBe("begin");
        expect(getTextContent(child)).toBe("AQID");
        count++;
      }
      pending.push(child);
    }
  }
  expect(count).toBe(1);
};

const childRevisions = (doc: PMNode) => {
  const found = [];
  doc.descendants((node) => {
    for (const mark of node.marks) {
      if (mark.type.name !== "insertion" && mark.type.name !== "deletion") continue;
      const attrs = expectTrackedChangeMarkAttrs(mark);
      if (attrs.revisionId !== 37) continue;
      found.push({
        text: node.textContent,
        type: mark.type.name,
        revisionId: attrs.revisionId,
        author: attrs.author,
        date: attrs.date,
        initials: attrs.initials,
        ancestors:
          attrs._docxRevisionAncestors
            ?.filter(({ author }) => author === "Ancestor Reviewer")
            .map(({ type, revisionId, author }) => ({ type, revisionId, author })) ?? [],
      });
    }
  });
  return found;
};

test(
  "simple field child revision ownership survives editor save and reopen",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.constantFrom("insertion", "deletion", "moveFrom", "moveTo"),
        fc.constantFrom("none", "insertion", "deletion"),
        fc.constantFrom("run", "hyperlink", "inlineWrapper"),
        fc.integer({ min: 1, max: 8 }),
        fc.constantFrom("none", "insertion"),
        fc.constantFrom("Child Reviewer", "Outer Reviewer", "Žluťoučký", "レビュー"),
        fc.constantFrom("none", "fldData"),
        async (kind, outer, carrier, length, ancestry, author, fieldCapture) => {
          const run = {
            type: "run",
            content: [{ type: "text", text: "x".repeat(length) }],
          } as const;
          const revision: TrackedRunChange = {
            type: kind,
            info: {
              id: 37,
              author,
              date: "2026-01-02T03:04:05Z",
              initials: "CR",
            },
            content: [{ ...run, content: [...run.content] }],
          };
          const field: SimpleField = {
            type: "simpleField",
            instruction: "REF target",
            fieldType: "REF",
            content: [],
          };
          if (carrier === "hyperlink")
            revision.content = [
              {
                type: "hyperlink",
                anchor: "target",
                children: [{ ...run, content: [...run.content] }],
              },
            ];
          if (carrier === "inlineWrapper")
            revision.content = [
              {
                type: "inlineWrapper",
                kind: "bidi",
                control: "embedding",
                direction: "rtl",
                content: [{ ...run, content: [...run.content] }],
              },
            ];
          const trackedChild: TrackedRunChange =
            ancestry === "none"
              ? revision
              : {
                  type: "insertion",
                  info: { id: 36, author: "Ancestor Reviewer" },
                  content: [revision],
                };
          field.content = [
            { type: "run", content: [{ type: "text", text: "before" }] },
            trackedChild,
            { type: "run", content: [{ type: "text", text: "after" }] },
          ];
          if (fieldCapture === "fldData")
            field.content.unshift({ type: "preservedInline", xml: FIELD_DATA, text: "" });
          const source = createEmptyDocument();
          const content: ParagraphContent =
            outer === "none"
              ? field
              : { type: outer, info: { id: 35, author: "Outer Reviewer" }, content: [field] };
          // Field results share the paragraph renderer's revision projection in both modes.
          const fieldSource = createEmptyDocument();
          fieldSource.package.document.content = [{ type: "paragraph", content: [field] }];
          const flattenedSource = createEmptyDocument();
          flattenedSource.package.document.content = [
            { type: "paragraph", content: field.content },
          ];
          for (const trackedChanges of ["clean", "annotate"] as const) {
            expect(toMarkdown(fieldSource, { trackedChanges })).toEqual(
              toMarkdown(flattenedSource, { trackedChanges }),
            );
          }
          source.package.document.content = [{ type: "paragraph", content: [content] }];
          const initial = toProseDoc(source);
          expect(childRevisions(initial)).toHaveLength(1);
          // Initials are UI/model metadata, excluded from schema-strict CT_TrackChange XML.
          expect(
            childRevisions(
              toProseDoc(fromProseDoc(initial, source, { stylesheetSource: { type: "package" } })),
            ),
          ).toEqual(childRevisions(initial));
          const serializedRevisions = childRevisions(initial);
          for (const revisionAttrs of serializedRevisions) revisionAttrs.initials = undefined;
          const live = EditorState.create({ doc: initial });
          const serializedSource = fromProseDoc(initial, source, {
            stylesheetSource: { type: "package" },
          });
          if (fieldCapture === "fldData" && outer !== "none") expectFieldDataHome(serializedSource);
          const saved = await createDocx(serializedSource);
          const reopenedSource = await parseDocx(saved, { preloadFonts: false });
          if (fieldCapture === "fldData" && outer !== "none") expectFieldDataHome(reopenedSource);
          const reopened = EditorState.create({ doc: toProseDoc(reopenedSource) });
          expect(childRevisions(reopened.doc)).toEqual(serializedRevisions);
          for (const mode of ["accept", "reject"] as const) {
            const childKept =
              mode === "accept"
                ? kind === "insertion" || kind === "moveTo"
                : ancestry === "none" && (kind === "deletion" || kind === "moveFrom");
            const outerRemoved = mode === "accept" ? outer === "deletion" : outer === "insertion";
            const expected = outerRemoved
              ? ""
              : `before${childKept ? "x".repeat(length) : ""}after`;
            for (const [state, baseline] of [
              [live, source],
              [reopened, reopenedSource],
            ] as const) {
              const resolved = resolveAllChangesInHeadlessState(state, mode);
              expect(resolved.doc.textContent).toBe(expected);
              const reopenedResolved = await parseDocx(
                await createDocx(
                  fromProseDoc(resolved.doc, baseline, { stylesheetSource: { type: "package" } }),
                ),
                { preloadFonts: false },
              );
              expect(
                resolveAllChangesInHeadlessState(
                  EditorState.create({ doc: toProseDoc(reopenedResolved) }),
                  mode,
                ).doc.textContent,
              ).toBe(expected);
              let individuallyResolved = state;
              const resolve = mode === "accept" ? acceptAIEditRevision : rejectAIEditRevision;
              // Resolve one revision id at a time, exercising overlapping ancestry.
              for (let step = 0; step < 20; step++) {
                let revisionId: number | undefined;
                individuallyResolved.doc.descendants((node) => {
                  const revisionMark = node.marks.find(
                    (mark) => mark.type.name === "insertion" || mark.type.name === "deletion",
                  );
                  if (revisionId === undefined && revisionMark)
                    revisionId = expectTrackedChangeMarkAttrs(revisionMark).revisionId;
                });
                if (revisionId === undefined) break;
                const before = individuallyResolved.doc;
                individuallyResolved = applyReviewCommand(
                  individuallyResolved,
                  resolve(revisionId),
                );
                expect(individuallyResolved.doc.eq(before)).toBe(false);
                const intermediate = await parseDocx(
                  await createDocx(
                    fromProseDoc(individuallyResolved.doc, baseline, {
                      stylesheetSource: { type: "package" },
                    }),
                  ),
                  { preloadFonts: false },
                );
                expect(toProseDoc(intermediate)).toBeDefined();
              }
              expect(individuallyResolved.doc.textContent).toBe(expected);
            }
          }
          const savedAgain = await createDocx(
            fromProseDoc(reopened.doc, reopenedSource, { stylesheetSource: { type: "package" } }),
          );
          expect(
            childRevisions(toProseDoc(await parseDocx(savedAgain, { preloadFonts: false }))),
          ).toEqual(serializedRevisions);
        },
      ),
      { numRuns: 40 },
    );
  },
  propertyTestTimeout(30_000),
);

for (const type of ["insertion", "deletion"] as const) {
  test(`tracked ${type} field data stays on the begin character through save/reopen`, async () => {
    const source = createEmptyDocument();
    source.package.document.content = [
      {
        type: "paragraph",
        content: [
          {
            type,
            info: { id: 35, author: "Reviewer" },
            content: [
              {
                type: "simpleField",
                instruction: "REF target",
                fieldType: "REF",
                content: [
                  { type: "preservedInline", xml: FIELD_DATA, text: "" },
                  { type: "run", content: [{ type: "text", text: "result" }] },
                ],
              },
            ],
          },
        ],
      },
    ];
    const projected = fromProseDoc(toProseDoc(source), source, {
      stylesheetSource: { type: "package" },
    });
    expectFieldDataHome(projected);
    const reopened = await parseDocx(await createDocx(projected), { preloadFonts: false });
    expectFieldDataHome(reopened);
    const projectedAgain = fromProseDoc(toProseDoc(reopened), reopened, {
      stylesheetSource: { type: "package" },
    });
    expectFieldDataHome(projectedAgain);
    expectFieldDataHome(await parseDocx(await createDocx(projectedAgain), { preloadFonts: false }));
  });
}

test("tracked fields refuse captures without a valid lowered home", () => {
  for (const xml of [
    '<w:subDoc xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>',
    '<w:fldSimple w:instr="REF target"/>',
    '<w:hyperlink w:anchor="target"/>',
    '<w:fldData xmlns:w="urn:other">AQID</w:fldData>',
    FIELD_DATA + FIELD_DATA,
  ]) {
    expect(() =>
      serializeParagraph({
        type: "paragraph",
        content: [
          {
            type: "insertion",
            info: { id: 35, author: "Reviewer" },
            content: [
              {
                type: "simpleField",
                instruction: "REF target",
                fieldType: "REF",
                content: [{ type: "preservedInline", xml, text: "" }],
              },
            ],
          },
        ],
      }),
    ).toThrow(/tracked field/i);
  }
});

test(
  "simple field non-text results survive resolution and save reopen",
  async () => {
    const nonTextResults = [
      { type: "tab" },
      { type: "break", breakType: "textWrapping" },
      { type: "symbol", font: "Wingdings", char: "F0A7" },
    ] as const satisfies readonly RunContent[];
    await assertProperty(
      fc.asyncProperty(
        fc.array(fc.constantFrom(...nonTextResults), { minLength: 1, maxLength: 6 }),
        fc.constantFrom("insertion", "deletion"),
        async (results, kind) => {
          const source = createEmptyDocument();
          source.package.document.content = [
            {
              type: "paragraph",
              content: [
                {
                  type: "simpleField",
                  instruction: "REF target",
                  fieldType: "REF",
                  content: [
                    { type: "run", content: [...results] },
                    {
                      type: kind,
                      info: { id: 37, author: "Reviewer" },
                      content: [{ type: "run", content: [{ type: "text", text: "tracked" }] }],
                    },
                  ],
                },
              ],
            },
          ];
          for (const mode of ["accept", "reject"] as const) {
            const state = EditorState.create({ doc: toProseDoc(source) });
            const resolved = resolveAllChangesInHeadlessState(state, mode);
            const individual = applyReviewCommand(
              state,
              (mode === "accept" ? acceptAIEditRevision : rejectAIEditRevision)(37),
            );
            for (const reviewed of [resolved, individual]) {
              for (const document of [
                fromProseDoc(reviewed.doc, source, { stylesheetSource: { type: "package" } }),
                await parseDocx(
                  await createDocx(
                    fromProseDoc(reviewed.doc, source, { stylesheetSource: { type: "package" } }),
                  ),
                  {
                    preloadFonts: false,
                  },
                ),
              ]) {
                const paragraph = document.package.document.content.at(0);
                if (paragraph?.type !== "paragraph")
                  throw new TypeError("Missing reviewed paragraph");
                const field = paragraph.content.find((child) => child.type === "simpleField");
                if (field?.type !== "simpleField") throw new TypeError("Missing reviewed field");
                const runChildren = field.content.flatMap((child) =>
                  child.type === "run" ? child.content : [],
                );
                expect(runChildren.filter((child) => child.type !== "text")).toEqual(results);
                expect(
                  runChildren
                    .flatMap((child) => (child.type === "text" ? [child.text] : []))
                    .join(""),
                ).toBe((mode === "accept") === (kind === "insertion") ? "tracked" : "");
                expect(toProseDoc(document).content.firstChild?.firstChild?.type.name).toBe(
                  "structuredField",
                );
              }
            }
          }
        },
      ),
      { numRuns: 20 },
    );
  },
  propertyTestTimeout(30_000),
);
