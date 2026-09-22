import { describe, expect, test } from "bun:test";

import { schema } from "../prosemirror/schema";
import { compareSectionBoundaryProperties } from "./section-boundary-properties";

const paragraph = (text: string, sectionProperties?: Record<string, unknown>) =>
  schema.node(
    "paragraph",
    sectionProperties === undefined ? null : { _sectionProperties: sectionProperties },
    text ? schema.text(text) : undefined,
  );

const documentWith = (...paragraphs: ReturnType<typeof paragraph>[]) =>
  schema.node("doc", null, paragraphs);

describe("compareSectionBoundaryProperties", () => {
  test("stages properties only on a retained existing endpoint", () => {
    const current = documentWith(paragraph("Before"), paragraph("After", { marginLeft: 720 }));
    const target = documentWith(paragraph("Before"), paragraph("After", { marginLeft: 1440 }));

    expect(compareSectionBoundaryProperties({ current, target })).toEqual({
      status: "matched",
      changes: [
        {
          position: current.child(0).nodeSize,
          current: { marginLeft: 720 },
          target: { marginLeft: 1440 },
        },
      ],
    });
  });

  test("refuses a section endpoint added to retained text", () => {
    const current = documentWith(paragraph("Before"), paragraph("After"));
    const target = documentWith(paragraph("Before"), paragraph("After", { marginLeft: 1440 }));

    expect(compareSectionBoundaryProperties({ current, target })).toEqual({
      status: "unalignable",
      detail: "section endpoint presence differs",
    });
  });

  test("ignores opaque blocks that cannot carry a section endpoint", () => {
    const current = documentWith(paragraph("Before"), paragraph("After", { marginLeft: 720 }));
    const target = schema.node("doc", null, [
      paragraph("Before"),
      schema.node("preservedBlock", { xml: '<w:altChunk r:id="rId9"/>' }),
      paragraph("After", { marginLeft: 1440 }),
    ]);

    expect(compareSectionBoundaryProperties({ current, target })).toEqual({
      status: "matched",
      changes: [
        {
          position: current.child(0).nodeSize,
          current: { marginLeft: 720 },
          target: { marginLeft: 1440 },
        },
      ],
    });
  });
});

import { EditorState } from "prosemirror-state";

import { acceptAllChanges, rejectAllChanges } from "../prosemirror/commands/comments";

const resolve = ({ state, mode }: { state: EditorState; mode: "accept" | "reject" }) => {
  let resolved = state;
  (mode === "accept" ? acceptAllChanges() : rejectAllChanges())(state, (transaction) => {
    resolved = state.apply(transaction);
  });
  return resolved;
};
import { stageSectionBoundaryProperties } from "./section-boundary-properties";

test("stages a boundary on an already inserted blank paragraph", () => {
  const inserted = schema.node("paragraph", {
    pPrMark: { kind: "ins", info: { id: 20, author: "Compare" } },
  });
  const state = EditorState.create({
    schema,
    doc: documentWith(paragraph("Before"), inserted, paragraph("After")),
  });
  const target = documentWith(
    paragraph("Before"),
    schema.node("paragraph", {
      _sectionProperties: { sectionStart: "nextPage", marginLeft: 1440 },
    }),
    paragraph("After"),
  );
  const result = stageSectionBoundaryProperties({
    state,
    target,
    originalRevisionIdSeed: 10,
    maxRanges: 1,
    revisionStamp: { idSeed: 30, date: "2026-09-13T00:00:00.000Z" },
    author: "Compare",
    mapTargetProperties: ({ target: targetProperties }) => ({
      kind: "inserted",
      target: targetProperties,
    }),
  });
  expect(result.status).toBe("matched");
  if (result.status !== "matched") return;
  const pending = state.apply(result.transaction);
  expect(resolve({ state: pending, mode: "accept" }).doc.eq(target)).toBe(true);
  expect(
    resolve({ state: pending, mode: "reject" }).doc.eq(
      documentWith(paragraph("Before"), paragraph("After")),
    ),
  ).toBe(true);
});

test("aligns an inserted endpoint after a preceding deleted paragraph", () => {
  const deleted = schema.node(
    "paragraph",
    { pPrMark: { kind: "del", info: { id: 9, author: "Compare" } } },
    schema.text("Removed", [schema.mark("deletion", { revisionId: 9, author: "Compare" })]),
  );
  const inserted = schema.node("paragraph", {
    pPrMark: { kind: "ins", info: { id: 20, author: "Compare" } },
  });
  const state = EditorState.create({
    schema,
    doc: documentWith(deleted, paragraph("Before"), inserted, paragraph("After")),
  });
  const target = documentWith(
    paragraph("Before"),
    schema.node("paragraph", {
      _sectionProperties: { sectionStart: "nextPage" },
    }),
    paragraph("After"),
  );
  const result = stageSectionBoundaryProperties({
    state,
    target,
    originalRevisionIdSeed: 10,
    maxRanges: 1,
    revisionStamp: { idSeed: 30, date: "2026-09-13T00:00:00.000Z" },
    author: "Compare",
    mapTargetProperties: ({ target: targetProperties }) => ({
      kind: "inserted",
      target: targetProperties,
    }),
  });
  expect(result.status).toBe("matched");
  if (result.status !== "matched") return;
  const pending = state.apply(result.transaction);
  expect(resolve({ state: pending, mode: "accept" }).doc.eq(target)).toBe(true);
  expect(
    resolve({ state: pending, mode: "reject" }).doc.eq(
      documentWith(paragraph("Removed"), paragraph("Before"), paragraph("After")),
    ),
  ).toBe(true);
});

test("stages retained endpoint properties alongside an inline text revision", () => {
  const deletion = schema.mark("deletion", { revisionId: 4, author: "Compare" });
  const insertion = schema.mark("insertion", { revisionId: 5, author: "Compare" });
  const state = EditorState.create({
    schema,
    doc: documentWith(
      schema.node(
        "paragraph",
        {
          _sectionProperties: { sectionStart: "nextPage", marginLeft: 720 },
        },
        [schema.text("Old", [deletion]), schema.text("New", [insertion])],
      ),
    ),
  });
  const target = documentWith(
    schema.node(
      "paragraph",
      {
        _sectionProperties: { sectionStart: "nextPage", marginLeft: 1440 },
      },
      schema.text("New"),
    ),
  );
  const result = stageSectionBoundaryProperties({
    state,
    target,
    originalRevisionIdSeed: 10,
    revisionStamp: { idSeed: 20, date: "2026-09-13T00:00:00.000Z" },
    author: "Compare",
    maxRanges: 1,
    mapTargetProperties: ({ kind, current, target: targetProperties }) => {
      if (kind === "inserted") return { kind, target: targetProperties };
      if (current === undefined) return null;
      return { kind, previous: current, target: targetProperties };
    },
  });
  expect(result.status).toBe("matched");
  if (result.status !== "matched") return;
  const pending = state.apply(result.transaction);
  expect(resolve({ state: pending, mode: "accept" }).doc.eq(target)).toBe(true);
  const rejected = resolve({ state: pending, mode: "reject" }).doc;
  expect(rejected.firstChild?.textContent).toBe("Old");
  expect(rejected.firstChild?.attrs["_sectionProperties"]).toEqual({
    sectionStart: "nextPage",
    marginLeft: 720,
  });
});
