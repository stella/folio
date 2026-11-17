import { panic } from "better-result";
import { expect, test } from "bun:test";
import { Step, Transform } from "prosemirror-transform";

import { schema } from "../prosemirror/schema";
import { resolveWholeStory } from "./wholeStoryRevisionResolution";

const followers = {
  preservedBlock: () => schema.node("preservedBlock", { xml: "<w:altChunk/>" }),
  blockSdt: () =>
    schema.node("blockSdt", { id: 23, tag: "standing" }, [
      schema.node("paragraph", null, [schema.text("inside")]),
    ]),
  blockCustomXml: () =>
    schema.node(
      "blockCustomXml",
      {
        openingXml: '<w:customXml w:element="standing">',
        closingXml: "</w:customXml>",
      },
      [schema.node("paragraph", null, [schema.text("inside")])],
    ),
};

for (const [kind, follower] of Object.entries(followers))
  test(`${kind} stays a block follower through mark resolution and JSON undo`, () => {
    for (const mode of ["accept", "reject"] as const)
      for (const mark of ["ins", "del"] as const)
        for (const tail of [false, true])
          for (const content of ["", "kept"])
            for (const section of [false, true]) {
              const properties = section ? { pageWidth: 11906 } : null;
              const first = schema.node(
                "paragraph",
                {
                  paraId: "10000000",
                  _sectionProperties: properties,
                },
                [schema.text("before")],
              );
              const attrs = {
                paraId: "10000001",
                pPrMark: {
                  kind: mark,
                  info: {
                    id: 7,
                    author: "Reviewer",
                    date: "2026-01-01T00:00:00Z",
                  },
                },
              };
              const target = schema.node(
                "paragraph",
                attrs,
                content ? [schema.text(content)] : [],
              );
              const opaque = follower();
              const ending = tail
                ? [
                    schema.node("paragraph", { paraId: "10000002" }, [
                      schema.text("after"),
                    ]),
                  ]
                : [];
              const source = schema.node("doc", null, [
                first,
                target,
                opaque,
                ...ending,
              ]);
              // The shared kernel guard keeps inserted marks before opaque
              // blocks. Only an accepted empty deletion can retire this
              // paragraph, and only while another paragraph ends the body.
              const removed =
                mode === "accept" && mark === "del" && content === "" && tail;
              const expected = schema.node("doc", null, [
                first,
                ...(removed
                  ? []
                  : [
                      schema.node(
                        "paragraph",
                        { ...attrs, pPrMark: null },
                        target.content,
                      ),
                    ]),
                opaque,
                ...ending,
              ]);
              const result = resolveWholeStory({
                doc: source,
                mode,
                styleResolver: null,
                numbering: null,
              });
              expect(result.failed).toBe(false);
              expect(result.resolved.eq(expected)).toBe(true);
              expect(
                result.resolved.firstChild?.attrs["_sectionProperties"],
              ).toEqual(properties);
              const sourceOpaque = first.nodeSize + target.nodeSize;
              const finalOpaque =
                first.nodeSize + (removed ? 0 : target.nodeSize);
              expect(result.mapping.map(sourceOpaque, 1)).toBe(finalOpaque);
              expect(result.resolved.nodeAt(finalOpaque)).toBe(opaque);

              const replay = new Transform(source);
              const inverses = [];
              for (const step of result.steps) {
                const copied = Step.fromJSON(
                  schema,
                  JSON.parse(JSON.stringify(step.toJSON())),
                );
                inverses.unshift(copied.invert(replay.doc).toJSON());
                replay.step(copied);
              }
              expect(replay.doc.eq(expected)).toBe(true);
              let undone = replay.doc;
              for (const inverse of inverses) {
                const applied = Step.fromJSON(
                  schema,
                  JSON.parse(JSON.stringify(inverse)),
                ).apply(undone);
                if (applied.doc === null)
                  panic("A resolution inverse restores its source document.");
                undone = applied.doc;
              }
              expect(undone.eq(source)).toBe(true);
            }
  });
