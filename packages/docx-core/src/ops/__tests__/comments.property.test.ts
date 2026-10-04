import {
  generatedCaseFor,
  GENERATED_PACKAGE_STORIES,
  packageDocumentArbitrary,
} from "../../../../../test/generators/packageOperationArbitraries";
import { opSeedArbitrary } from "./documentArbitraries";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import { type Comment, type Document, type Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { allocateCommentAnchorIds, planTrackedDeletion } from "../plan";
import { commentDocumentIssue, freshCommentId } from "../comments";
import { contractViolation } from "../contract";
import { documentStories, storyBody } from "../stories";
import { storyParagraphs } from "../blocks";
import { isCommentAnchor, leafSpans } from "../leaves";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  toOpEnvelope,
  type CommentOp,
  type DocumentOp,
  type OpStory,
} from "../types";

setDefaultTimeout(propertyTestTimeout(30_000));
const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const seed = (presence: "absent" | "undefined" | "present"): Document => {
  const document: Document = {
    package: {
      document: { content: [paragraph("00000001", "alpha"), paragraph("00000002", "omega")] },
      headers: new Map([
        ["rIdHeader", { type: "header", content: [paragraph("00000003", "header")] }],
      ]),
      footnotes: [{ type: "footnote", id: 1, content: [paragraph("00000004", "note")] }],
    },
  };
  if (presence === "present") document.package.document.comments = [];
  if (presence === "undefined") Reflect.set(document.package.document, "comments", undefined);
  return document;
};
const comment = (id: number, paraId: string, text: string): Comment => {
  const value: Comment = {
    id,
    author: "Author",
    content: [paragraph(paraId, text)],
    preserved: { children: [{ index: 0, xml: '<opaque xmlns="urn:fixture"/>' }] },
  };
  Reflect.set(value, "done", undefined);
  Reflect.set(value, "initials", undefined);
  return value;
};
const anchorSequence = (document: Document) =>
  documentStories(document).flatMap((story) =>
    storyParagraphs(storyBody(document, story)).flatMap(({ paragraph: block }) =>
      leafSpans(block.content)
        .filter(({ node }) => isCommentAnchor(node))
        .map(({ node }) => node),
    ),
  );
const exact = (document: Document, op: DocumentOp) => {
  const applied = applyDocumentOp(document, op).unwrap();
  expect(contractViolation(applied.document)).toBeUndefined();
  expect(commentDocumentIssue(applied.document)).toBeUndefined();
  const wire: DocumentOp[] = JSON.parse(JSON.stringify(applied.inverse));
  const undone = applyDocumentOps(applied.document, wire).unwrap();
  expect(undone.document).toStrictEqual(document);
  const redone = applyDocumentOps(
    undone.document,
    JSON.parse(JSON.stringify(undone.inverse)),
  ).unwrap();
  expect(redone.document).toStrictEqual(applied.document);
  expect(applied.document.package.document.comments?.map(({ id }) => id)).toEqual(
    redone.document.package.document.comments?.map(({ id }) => id),
  );
  return applied.document;
};
const COMMANDS = {
  createComment: "author",
  updateCommentContent: "author",
  setCommentResolution: "author",
  deleteComment: "author",
  restoreCommentState: "inverse",
} as const satisfies Record<CommentOp["type"], "author" | "inverse">;

test("generated comment histories preserve exact state, declaration order, opaque markup and owned anchors", () => {
  const seen = new Set<string>();
  assertProperty(
    fc.property(
      fc.constantFrom("absent" as const, "undefined" as const, "present" as const),
      fc.constantFrom("main" as const, "header" as const, "note" as const),
      fc.array(fc.constantFrom("resolve", "open", "edit"), { minLength: 1, maxLength: 8 }),
      (presence, storyKind, commands) => {
        const original = seed(presence);
        const targets = {
          main: { story: OP_STORIES.MAIN, blockId: "00000001" },
          header: { story: { kind: "header", rId: "rIdHeader" }, blockId: "00000003" },
          note: { story: { kind: "footnote", id: 1 }, blockId: "00000004" },
        } as const satisfies Record<typeof storyKind, { story: OpStory; blockId: string }>;
        const { story, blockId } = targets[storyKind];
        let document = exact(original, {
          type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
          comment: comment(100, "00000010", "root"),
          anchor: {
            kind: "range",
            from: { story, blockId, offset: 0 },
            to: { story, blockId, offset: 2 },
          },
        });
        seen.add("createComment");
        seen.add("restoreCommentState");
        document = exact(document, {
          type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
          comment: comment(101, "00000011", "reply"),
          anchor: { kind: "reply", parentId: 100 },
        });
        document = exact(document, {
          type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
          comment: comment(102, "00000012", "point"),
          anchor: { kind: "point", at: { story, blockId, offset: 2 } },
        });
        const pointAnchors = anchorSequence(document).filter(
          (node) => node.type === "commentReference" && node.id === 102,
        );
        for (const [index, command] of commands.entries()) {
          const op: DocumentOp =
            command === "edit"
              ? {
                  type: DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT,
                  id: 100,
                  content: [paragraph("00000010", `changed${index}`)],
                  patch: { initials: null, date: "2026-01-02T00:00:00Z" },
                }
              : {
                  type: DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION,
                  id: 100,
                  status: command === "open" ? "open" : "resolved",
                };
          document = exact(document, op);
          seen.add(op.type);
        }
        document = exact(document, {
          type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
          id: 101,
          scope: "reply",
        });
        seen.add("deleteComment");
        document = exact(document, {
          type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
          id: 100,
          scope: "thread",
        });
        expect(document.package.document.comments?.map(({ id }) => id)).toEqual([102]);
        expect(anchorSequence(document)).toEqual(pointAnchors);
        expect(document.package.document.comments?.at(0)?.preserved).toEqual(
          comment(102, "00000012", "point").preserved,
        );
      },
    ),
    { seed: 20261004, numRuns: 50 },
  );
  expect([...seen].toSorted()).toEqual(Object.keys(COMMANDS).toSorted());
});

test("revision association derives the legacy parent relation and anchors without wrapping markers", () => {
  const document = seed("absent");
  const first = document.package.document.content.at(0);
  if (first?.type !== "paragraph") throw new Error("Missing paragraph");
  first.content = [
    { type: "insertion", info: { id: 7, author: "Reviewer" }, content: first.content },
  ];
  const changed = exact(document, {
    type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
    comment: comment(100, "00000010", "review"),
    anchor: { kind: "revision", story: OP_STORIES.MAIN, revisionId: 7 },
  });
  expect(changed.package.document.comments?.at(0)?.parentId).toBe(7);
  expect(
    leafSpans(
      changed.package.document.content.at(0)?.type === "paragraph"
        ? (changed.package.document.content.at(0)?.content ?? [])
        : [],
    )
      .filter(({ node }) => isCommentAnchor(node))
      .every(({ ancestors }) =>
        ancestors.every(({ type }) => type !== "insertion" && type !== "deletion"),
      ),
  ).toBe(true);
  expect(freshCommentId(changed).unwrap()).toBe(101);
});

test("comment inverses refuse changed owned text and preserve unrelated package identity", () => {
  const document = seed("absent");
  const op = {
    type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
    comment: comment(100, "00000010", "root"),
    anchor: { kind: "point", at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 } },
  } as const;
  const applied = applyDocumentOp(document, op).unwrap();
  expect(applied.document.package.headers).toBe(document.package.headers);
  const other = document.package.document.content.at(1);
  expect(applied.document.package.document.content.at(1)).toBe(other);
  const changed = applyDocumentOp(applied.document, {
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 2 },
    text: "x",
  }).unwrap();
  expect(applyDocumentOps(changed.document, applied.inverse).isErr()).toBe(true);
  expect(toOpEnvelope(op).schema).toBe(10);
});

test("every comment package-operation generator exercises its exact inverse in every story", () => {
  const observed = new Set<string>();
  assertProperty(
    fc.property(packageDocumentArbitrary, opSeedArbitrary, (document, inputSeed) => {
      for (const kind of Object.values(DOCUMENT_OP_TYPES)) {
        if (!Object.hasOwn(COMMANDS, kind)) continue;
        for (const story of GENERATED_PACKAGE_STORIES) {
          const generated = generatedCaseFor({ document, seed: inputSeed, story, kind });
          exact(generated.document, generated.op);
          observed.add(generated.op.type);
        }
      }
    }),
    { seed: 20261005, numRuns: 8 },
  );
  expect([...observed].toSorted()).toEqual(Object.keys(COMMANDS).toSorted());
});

test("generated text and revision histories retain live comment ownership and exact reversible state", () => {
  assertProperty(
    fc.property(
      fc.array(
        fc.constantFrom("direct", "tracked", "delete", "accept", "reject", "split", "join"),
        { minLength: 2, maxLength: 10 },
      ),
      (commands) => {
        let document = seed("present");
        document = exact(document, {
          type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
          comment: comment(100, "00000010", "thread"),
          anchor: {
            kind: "range",
            from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
            to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
          },
        });
        document = exact(document, {
          type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
          comment: comment(101, "00000011", "reply"),
          anchor: { kind: "reply", parentId: 100 },
        });
        const definitions = structuredClone(document.package.document.comments);
        let revisionId = 1000;
        let paragraphId = 0x1000;
        for (const command of commands) {
          const first = storyParagraphs(document.package.document).at(0)?.paragraph;
          if (!first?.paraId) throw new Error("Missing first paragraph");
          const start = {
            story: OP_STORIES.MAIN,
            blockId: first.paraId,
            offset: 0,
            zeroWidthBefore: 0,
          };
          const stamp = { id: revisionId++, author: "Reviewer", date: "2026-01-01T00:00:00Z" };
          const firstFreshRevision = revisionId;
          const newIds = {
            revision: Array.from({ length: 16 }, (_, index) => firstFreshRevision + index),
          };
          revisionId += 16;
          let op: DocumentOp;
          switch (command) {
            case "direct":
              op = { type: DOCUMENT_OP_TYPES.INSERT_TEXT, at: start, text: "x", newIds };
              break;
            case "tracked":
              op = {
                type: DOCUMENT_OP_TYPES.INSERT_TEXT,
                at: start,
                text: "y",
                revision: stamp,
                newIds,
              };
              break;
            case "delete":
              op = {
                type: DOCUMENT_OP_TYPES.DELETE_RANGE,
                from: start,
                to: { ...start, offset: 1 },
                revision: stamp,
                newIds,
              };
              break;
            case "accept":
            case "reject":
              op = {
                type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                story: OP_STORIES.MAIN,
                revisionIds: Array.from({ length: revisionId }, (_, index) => index),
                decision: command,
              };
              break;
            case "split":
              op = {
                type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
                at: { ...start, offset: 1 },
                newBlockId: (paragraphId++).toString(16).padStart(8, "0"),
                newHalf: "second",
                newIds,
              };
              break;
            case "join": {
              const next = storyParagraphs(document.package.document).at(1)?.paragraph;
              if (!next?.paraId) continue;
              op = {
                type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
                story: OP_STORIES.MAIN,
                blockId: first.paraId,
                nextBlockId: next.paraId,
                survivor: "first",
                newIds,
              };
              break;
            }
            default: {
              const unreachable: never = command;
              throw new Error(String(unreachable));
            }
          }
          if (command === "delete") {
            const planned = planTrackedDeletion(document, {
              from: op.type === "deleteRange" ? op.from : start,
              to: { ...start, offset: 1 },
              revision: stamp,
              newIds,
            }).unwrap();
            const beforeAnchors = anchorSequence(document);
            const applied = applyDocumentOps(document, planned).unwrap();
            expect(anchorSequence(applied.document)).toStrictEqual(beforeAnchors);
            expect(
              applyDocumentOps(
                applied.document,
                JSON.parse(JSON.stringify(applied.inverse)),
              ).unwrap().document,
            ).toStrictEqual(document);
            document = applied.document;
          } else {
            const result = applyDocumentOp(document, op);
            if (result.isErr()) continue;
            document = exact(document, op);
          }
          expect(document.package.document.comments).toStrictEqual(definitions);
          expect(commentDocumentIssue(document)).toBeUndefined();
        }
      },
    ),
    { seed: 20261006, numRuns: 40 },
  );
});

test("generated interior source-shaped comment histories undo and redo the complete journal", () => {
  assertProperty(
    fc.property(
      fc.constantFrom(0, 1, 3, 4, 5, 6),
      fc.constantFrom("point" as const, "range" as const),
      fc.boolean(),
      (offset, kind, ownUndefined) => {
        const original = seed(ownUndefined ? "undefined" : "absent");
        original.package.relationships = new Map([
          [
            "rId8",
            { id: "rId8", type: "urn:opaque", target: "opaque.bin", targetMode: "Internal" },
          ],
        ]);
        const first = original.package.document.content.at(0);
        if (first?.type !== "paragraph") throw new Error("Missing source paragraph");
        first.content = [
          {
            type: "run",
            formatting: { bold: true, language: { val: "cs-CZ" } },
            preservedAttributes: [{ namespace: "urn:source", name: "stamp", value: "authored" }],
            content: [{ type: "text", text: "A😀éB" }],
          },
        ];
        if (ownUndefined) Reflect.set(first, "formatting", undefined);
        const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset };
        const root: CommentOp = {
          type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
          comment: comment(100, "00000010", "root"),
          anchor: kind === "point" ? { kind, at } : { kind, from: at, to: { ...at, offset: 6 } },
        };
        const ops: CommentOp[] = [
          root,
          {
            type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
            comment: comment(101, "00000011", "reply"),
            anchor: { kind: "reply", parentId: 100 },
          },
          {
            type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
            comment: comment(102, "00000012", "nested"),
            anchor: { kind: "reply", parentId: 101 },
          },
          { type: DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION, id: 100, status: "resolved" },
          { type: DOCUMENT_OP_TYPES.DELETE_COMMENT, id: 100, scope: "thread" },
        ];
        let current = original;
        const inverse: DocumentOp[][] = [];
        for (const op of ops) {
          const next = applyDocumentOp(current, op).unwrap();
          inverse.unshift(next.inverse);
          current = next.document;
          expect(commentDocumentIssue(current)).toBeUndefined();
        }
        const final = current;
        const redo: DocumentOp[][] = [];
        for (const step of inverse) {
          const next = applyDocumentOps(current, JSON.parse(JSON.stringify(step))).unwrap();
          redo.unshift(next.inverse);
          current = next.document;
        }
        expect(current).toStrictEqual(original);
        for (const step of redo)
          current = applyDocumentOps(current, JSON.parse(JSON.stringify(step))).unwrap().document;
        expect(current).toStrictEqual(final);
      },
    ),
    { testFile: import.meta.path, seed: 20261007, numRuns: 60 },
  );
});

test("generated forged comment inverses refuse same-text unowned run changes", () => {
  assertProperty(
    fc.property(fc.string(), fc.boolean(), (attribute, bold) => {
      const document = seed("absent");
      const created = applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
        comment: comment(100, "00000010", "root"),
        anchor: { kind: "point", at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 } },
      }).unwrap();
      const inverse = created.inverse.at(0);
      if (inverse?.type !== DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE)
        throw new Error("Missing inverse");
      const entry = inverse.state.anchors.at(0);
      if (!entry) throw new Error("Missing anchor paragraph");
      const forged = {
        ...inverse,
        state: {
          ...inverse.state,
          anchors: [
            {
              ...entry,
              content: entry.content.map((node) => {
                if (node.type !== "run") return node;
                const changed = structuredClone(node);
                changed.formatting = { bold };
                changed.preservedAttributes = [
                  { namespace: "urn:forged", name: "stamp", value: attribute },
                ];
                return changed;
              }),
            },
          ],
        },
      };
      expect(applyDocumentOp(created.document, forged).isErr()).toBe(true);
      expect(created.document.package.document.content.at(1)).toBe(
        document.package.document.content.at(1),
      );
    }),
    { testFile: import.meta.path, seed: 20261008, numRuns: 40 },
  );
});

test("generated tracked and controlled comment anchors allocate exact package-owned cut pools", () => {
  assertProperty(
    fc.property(
      fc.integer({ min: 1, max: 4 }),
      fc.constantFrom("point" as const, "range" as const),
      (depth, kind) => {
        const document = seed("absent");
        const first = document.package.document.content.at(0);
        if (first?.type !== "paragraph") throw new Error("Missing paragraph");
        for (let index = 0; index < depth; index++)
          first.content = [
            {
              type: "insertion",
              info: { id: 200 + index, author: "A", date: "2026-01-02T00:00:00Z" },
              content: first.content,
            },
          ];
        first.content = [
          {
            type: "inlineSdt",
            properties: { id: 400, sdtType: "richText", tag: "owned" },
            content: first.content,
          },
        ];
        const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset: 2 };
        const create: CommentOp = {
          type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
          comment: comment(100, "00000010", "root"),
          anchor: kind === "point" ? { kind, at } : { kind, from: at, to: { ...at, offset: 4 } },
        };
        const newIds = allocateCommentAnchorIds(document, create).unwrap();
        expect(allocateCommentAnchorIds(document, create).unwrap()).toStrictEqual(newIds);
        let current = exact(document, { ...create, newIds });
        const reply = {
          type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
          comment: comment(101, "00000011", "reply"),
          anchor: { kind: "reply", parentId: 100 },
        } as const;
        current = exact(current, {
          ...reply,
          newIds: allocateCommentAnchorIds(current, reply).unwrap(),
        });
        expect(commentDocumentIssue(current)).toBeUndefined();
        for (const id of newIds.revision ?? []) expect(id).toBeGreaterThan(200 + depth - 1);
        for (const id of newIds.control ?? []) expect(id).toBeGreaterThan(400);
      },
    ),
    { testFile: import.meta.path, seed: 20261009, numRuns: 24 },
  );
});
