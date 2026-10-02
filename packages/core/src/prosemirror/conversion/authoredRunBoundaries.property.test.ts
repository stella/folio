import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { parseDocumentBody } from "../../docx/documentParser";
import { serializeDocument } from "../../docx/serializer/documentSerializer";

import { assertProperty, propertyConfig } from "../../../../../test/property-testing";
import { resolveWholeStory } from "../../internal/wholeStoryRevisionResolution";
import type { Document, Paragraph, Run } from "../../types/document";
import { createStarterKit } from "../extensions/StarterKit";
import { ExtensionManager } from "../extensions/ExtensionManager";
import { schema } from "../schema";
import { RUN_IDENTITY_MARK_NAME } from "../runIdentity";
import { fromProseDoc, proseDocToBlocks } from "./fromProseDoc";
import { headerFooterToProseDoc, toProseDoc } from "./toProseDoc";

const text = fc.stringMatching(/^[a-z]{2,8}$/u);
const formatting = fc.constantFrom(
  undefined,
  { bold: false },
  { noProof: true },
  { fontSizeCs: 24 },
);
const runs = fc
  .tuple(fc.array(text, { minLength: 2, maxLength: 6 }), formatting)
  .map(([texts, authored]) =>
    texts.map(
      (value): Run => ({
        type: "run",
        content: [{ type: "text", text: value }],
        ...(authored === undefined ? {} : { formatting: authored }),
      }),
    ),
  );

const runText = (content: Run[]): string =>
  content
    .flatMap((run) => run.content)
    .filter((item) => item.type === "text")
    .map((item) => item.text)
    .join("");

const documentOf = (content: Paragraph[]): Document => ({ package: { document: { content } } });

const savedRuns = (document: Document): Run[] => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") throw new TypeError("Expected the synthetic paragraph");
  return paragraph.content.filter((item): item is Run => item.type === "run");
};

// The previous generator exercised only runs with attribute/sink payloads.
// Ordinary runs also own boundaries, even when their marks are identical.
describe("authored run boundary ownership", () => {
  test("no-op and repeated saves preserve every nonempty authored run", () => {
    fc.assert(
      fc.property(runs, (content) => {
        const source = documentOf([{ type: "paragraph", content }]);
        const saved = fromProseDoc(toProseDoc(source), source);
        expect(savedRuns(saved)).toEqual(content);
        expect(savedRuns(fromProseDoc(toProseDoc(saved), saved))).toEqual(content);
      }),
      propertyConfig(),
    );
  });

  test("mixed text and note-reference content survives source-owned run joins", () => {
    fc.assert(
      fc.property(
        text,
        text,
        fc.constantFrom("footnoteRef", "endnoteRef"),
        fc.boolean(),
        (before, after, type, customMarkFollows) => {
          const content: Run[] = [
            {
              type: "run",
              content: [
                { type: "text", text: before },
                { type, id: 41, customMarkFollows },
                { type: "text", text: after },
              ],
            },
          ];
          const source = documentOf([{ type: "paragraph", content }]);
          const saved = fromProseDoc(toProseDoc(source), source);
          expect(savedRuns(saved)).toEqual(content);
          expect(savedRuns(fromProseDoc(toProseDoc(saved), saved))).toEqual(content);
        },
      ),
      propertyConfig(),
    );
  });

  test("an interior insertion extends one ordinary run and retains its authored neighbors", () => {
    const manager = new ExtensionManager(createStarterKit());
    manager.buildSchema();
    manager.initializeRuntime();
    fc.assert(
      fc.property(runs, fc.nat(), fc.nat(), (content, pickedRun, pickedOffset) => {
        const runIndex = pickedRun % content.length;
        const target = content.at(runIndex);
        if (target === undefined) throw new TypeError("Expected the generated run");
        const targetText = runText([target]);
        const offset = 1 + (pickedOffset % (targetText.length - 1));
        const position = 1 + runText(content.slice(0, runIndex)).length + offset;
        const source = documentOf([{ type: "paragraph", content }]);
        const state = EditorState.create({
          doc: toProseDoc(source),
          plugins: manager.getPlugins(),
        });
        const edited = state.apply(state.tr.insertText("!", position));
        const saved = savedRuns(fromProseDoc(edited.doc, source));
        expect(saved).toEqual(
          content.map((run, index) =>
            index === runIndex
              ? {
                  ...run,
                  content: [
                    {
                      type: "text",
                      text: targetText.slice(0, offset) + "!" + targetText.slice(offset),
                    },
                  ],
                }
              : run,
          ),
        );
      }),
      propertyConfig(),
    );
  });

  test("joining paragraphs cannot collide their authored run identities", () => {
    fc.assert(
      fc.property(runs, (content) => {
        const source = documentOf(content.map((run) => ({ type: "paragraph", content: [run] })));
        let state = EditorState.create({ doc: toProseDoc(source) });
        while (state.doc.childCount > 1) {
          state = state.apply(state.tr.join(state.doc.child(0).nodeSize));
        }
        expect(savedRuns(fromProseDoc(state.doc, source))).toEqual(content);
      }),
      propertyConfig(),
    );
  });

  // The old fixture copied one run record into two imported w:r records and
  // mistook equal payloads for shared ownership. Imported records own distinct
  // identities even when a revision separates otherwise identical properties.
  test("revision resolution retains distinct authored runs with equal payloads", () => {
    fc.assert(
      fc.property(
        runs,
        fc.constantFrom("insertion", "deletion", "runPropertyChange"),
        fc.constantFrom("accept", "reject"),
        fc.boolean(),
        (content, revision, mode, withSession) => {
          // SAFETY: the generator always produces at least two runs.
          const first = content[0]!;
          // Give property-change rejection an explicit previous state so this
          // ownership invariant does not depend on absent-vs-empty formatting.
          const boundaryRun =
            revision === "runPropertyChange" && first.formatting === undefined
              ? { ...first, formatting: { bold: false } }
              : first;
          const sharedAttributes = withSession ? [{ name: "rsidR", value: "00AB12CD" }] : undefined;
          const owned = (value: string): Run => ({
            ...boundaryRun,
            content: [{ type: "text", text: value }],
            ...(sharedAttributes === undefined ? {} : { preservedAttributes: sharedAttributes }),
          });
          const left = owned("left");
          const right = owned("right");
          const middle: Run = { type: "run", content: [{ type: "text", text: "!" }] };
          const source = documentOf([
            {
              type: "paragraph",
              content:
                revision === "runPropertyChange"
                  ? [
                      {
                        ...left,
                        propertyChanges: [
                          {
                            type: "runPropertyChange",
                            info: { id: 4, author: "Reviewer" },
                            previousFormatting: boundaryRun.formatting,
                            currentFormatting: boundaryRun.formatting,
                          },
                        ],
                      },
                      right,
                      ...content.slice(1),
                    ]
                  : [
                      left,
                      {
                        type: revision,
                        info: { id: 4, author: "Reviewer" },
                        content: [middle],
                      },
                      right,
                      ...content.slice(1),
                    ],
            },
          ]);
          const { resolved } = resolveWholeStory({
            doc: toProseDoc(source),
            mode,
            styleResolver: null,
          });
          const retainsMiddle =
            (revision === "insertion" && mode === "accept") ||
            (revision === "deletion" && mode === "reject");
          expect(savedRuns(fromProseDoc(resolved, source))).toEqual([
            left,
            ...(retainsMiddle ? [middle] : []),
            right,
            ...content.slice(1),
          ]);
        },
      ),
      propertyConfig(),
    );
  });

  test("rejecting a tracked insertion rejoins actual same-source pieces without joining neighbors", () => {
    assertProperty(
      fc.property(runs, fc.nat(), fc.boolean(), (content, pickedOffset, withSession) => {
        // SAFETY: the generator always produces at least two runs.
        const first = content[0]!;
        const original = withSession
          ? { ...first, preservedAttributes: [{ name: "rsidR", value: "00AB12CD" }] }
          : first;
        const originalText = runText([original]);
        const offset = 1 + (pickedOffset % (originalText.length - 1));
        const sourceRuns = [original, ...content.slice(1)];
        const source = documentOf([{ type: "paragraph", content: sourceRuns }]);
        const state = EditorState.create({ doc: toProseDoc(source) });
        const position = 1 + offset;
        // insertText inherits the projected source's actual runIdentity mark;
        // adding the revision splits that owner instead of importing new runs.
        const edited = state.apply(
          state.tr
            .insertText("!", position)
            .addMark(
              position,
              position + 1,
              schema.mark("insertion", { revisionId: 4, author: "Reviewer" }),
            ),
        );
        const sourceOwner = state.doc
          .nodeAt(1)
          ?.marks.find(({ type }) => type.name === RUN_IDENTITY_MARK_NAME);
        expect(sourceOwner).toBeDefined();
        for (const piecePosition of [1, position, position + 1]) {
          const pieceOwner = edited.doc
            .nodeAt(piecePosition)
            ?.marks.find(({ type }) => type.name === RUN_IDENTITY_MARK_NAME);
          expect(pieceOwner?.attrs).toEqual(sourceOwner?.attrs);
        }
        const { resolved } = resolveWholeStory({
          doc: edited.doc,
          mode: "reject",
          styleResolver: null,
        });
        expect(savedRuns(fromProseDoc(resolved, source))).toEqual(sourceRuns);
        // The importer deliberately consolidates equal unowned XML runs. Edit
        // that imported model so the reopen oracle compares the same ownership
        // boundary before and after save, rather than raw unparsed fixtures.
        const imported: Document = {
          package: { document: parseDocumentBody(serializeDocument(source)) },
        };
        const importedRuns = savedRuns(imported);
        const importedState = EditorState.create({ doc: toProseDoc(imported) });
        const importedEdit = importedState.apply(
          importedState.tr
            .insertText("!", position)
            .addMark(
              position,
              position + 1,
              schema.mark("insertion", { revisionId: 4, author: "Reviewer" }),
            ),
        );
        const tracked = fromProseDoc(importedEdit.doc, imported);
        const xml = serializeDocument(tracked);
        const reopened: Document = { package: { document: parseDocumentBody(xml) } };
        const reopenedResolution = resolveWholeStory({
          doc: toProseDoc(reopened),
          mode: "reject",
          styleResolver: null,
        });
        expect(savedRuns(fromProseDoc(reopenedResolution.resolved, reopened))).toEqual(
          importedRuns,
        );
      }),
      {},
    );
  });

  test("joining imported table-cell paragraphs retains distinct equal-payload source runs", () => {
    fc.assert(
      fc.property(runs, (content) => {
        const source: Document = {
          package: {
            document: {
              content: [
                {
                  type: "table",
                  rows: [
                    {
                      type: "tableRow",
                      cells: content.map((run) => ({
                        type: "tableCell",
                        content: [
                          {
                            type: "paragraph",
                            content: [
                              {
                                ...run,
                                preservedAttributes: [{ name: "rsidR", value: "00AB12CD" }],
                              },
                            ],
                          },
                        ],
                      })),
                    },
                  ],
                },
              ],
            },
          },
        };
        const projected = toProseDoc(source);
        const paragraphNodes: PMNode[] = [];
        projected.descendants((node) => {
          if (node.type.name === "paragraph") paragraphNodes.push(node);
        });
        const joined = schema.node(
          "doc",
          null,
          schema.node(
            "paragraph",
            null,
            paragraphNodes.flatMap((paragraph) => {
              const nodes: PMNode[] = [];
              paragraph.forEach((node) => nodes.push(node));
              return nodes;
            }),
          ),
        );
        const saved = fromProseDoc(joined);
        const reopened: Document = {
          package: { document: parseDocumentBody(serializeDocument(saved)) },
        };
        expect(savedRuns(saved)).toEqual(
          content.map((run) => ({
            ...run,
            preservedAttributes: [{ name: "rsidR", value: "00AB12CD" }],
          })),
        );
        expect(savedRuns(fromProseDoc(toProseDoc(reopened), reopened))).toEqual(
          savedRuns(reopened),
        );
      }),
      propertyConfig(),
    );
  });

  test("secondary stories preserve the same authored boundaries", () => {
    fc.assert(
      fc.property(runs, (content) => {
        expect(proseDocToBlocks(headerFooterToProseDoc([{ type: "paragraph", content }]))).toEqual([
          { type: "paragraph", content },
        ]);
      }),
      propertyConfig(),
    );
  });
});
