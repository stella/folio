import { describe, expect, test } from "bun:test";

import type { Comment } from "../../types/content";
import { parseComments } from "../commentParser";
import {
  planCommentParts,
  serializeComments,
  serializeCommentsExtended,
} from "./commentSerializer";

function makeComment(id: number, parentId?: number): Comment {
  return {
    id,
    author: "Tester",
    date: "2026-05-15T00:00:00Z",
    content: [
      {
        type: "paragraph",
        formatting: {},
        content: [
          {
            type: "run",
            formatting: {},
            content: [{ type: "text", text: "body" }],
          },
        ],
      },
    ],
    ...(parentId !== undefined ? { parentId } : {}),
  };
}

describe("serializeComments", () => {
  test.each([{ comments: [] }, { comments: [makeComment(1)] }])(
    "binds every ignorable namespace prefix",
    ({ comments }) => {
      const xml = serializeComments(planCommentParts(comments));
      const prefixes = xml
        .match(/mc:Ignorable="([^"]+)"/u)
        ?.at(1)
        ?.split(/\s+/u);
      expect(prefixes?.length).toBeGreaterThan(0);
      for (const prefix of prefixes ?? []) {
        expect(xml).toContain(`xmlns:${prefix}="`);
      }
    },
  );

  test("emits a valid empty <w:comments/> document when the array is empty", () => {
    // Previously returned the empty string, which is not valid OOXML.
    // Save paths now overwrite the original `word/comments.xml` part
    // even when the editor has zero comments — that requires the
    // serializer to produce a well-formed empty document so the part
    // can be replaced rather than skipped.
    const xml = serializeComments(planCommentParts([]));
    expect(xml.startsWith("<?xml")).toBe(true);
    expect(xml).toContain("<w:comments xmlns:");
    expect(xml).toContain("</w:comments>");
    // No `<w:comment>` children.
    expect(xml).not.toContain("<w:comment ");
  });

  test("writes the comments in the order the model holds them", () => {
    // The part was written top-level-first, which reshuffles a document whose
    // comments.xml interleaves replies with later thread roots — the order the
    // next parse hands back, so every consumer reading `comments[]` by position
    // saw one comment's author and body under another's place in the list.
    const xml = serializeComments(
      planCommentParts([makeComment(2, 1), makeComment(1), makeComment(3)]),
    );

    expect([...xml.matchAll(/<w:comment w:id="(\d+)"/gu)].map(([, id]) => id)).toEqual([
      "2",
      "1",
      "3",
    ]);
  });

  test("writes a duplicate w:id once, keeping the definition markers resolve to", () => {
    const [first, second] = [makeComment(1), { ...makeComment(1), author: "Second" }];

    const xml = serializeComments(planCommentParts([first, second]));

    expect([...xml.matchAll(/<w:comment w:id="(\d+)"/gu)]).toHaveLength(1);
    expect(xml).toContain('w:author="Tester"');
  });

  test("preserves an explicitly empty author attribute", () => {
    const xml = serializeComments(planCommentParts([{ ...makeComment(1), author: "" }]));

    expect(xml).toContain('<w:comment w:id="1" w:author=""');
  });

  test("preserves an explicitly empty initials attribute", () => {
    const comment = { ...makeComment(1), initials: "" };
    const xml = serializeComments(planCommentParts([comment]));

    expect(xml).toContain('w:initials=""');
    expect(parseComments(xml, null, null, new Map(), new Map()).at(0)?.initials).toBe("");
  });

  test("escapes a paragraph paraId that carries markup instead of a real Word id", () => {
    // A malformed/attacker-supplied `paraId` (e.g. relayed through a
    // collaboration payload) must not be able to break out of the
    // `w14:paraId="..."` attribute and inject sibling XML.
    const malicious = '12345678"/><script>alert(1)</script><w:p w14:paraId="';
    const comment = makeComment(1);
    comment.content[0]!.paraId = malicious;

    const xml = serializeComments(planCommentParts([comment]));

    expect(xml).not.toContain("<script>");
    expect(xml).toContain("w14:paraId=");
    expect(xml).toContain("&lt;script&gt;");
    expect(xml).toContain("&quot;");
  });

  test("escapes hyperlink relationship ids in comment content", () => {
    const comment = makeComment(1);
    comment.content[0]!.content = [
      {
        type: "hyperlink",
        rId: 'rId1"/><w:r><w:t>injected</w:t></w:r><w:hyperlink r:id="rId2',
        children: [
          {
            type: "run",
            content: [{ type: "text", text: "linked" }],
          },
        ],
      },
    ];

    const xml = serializeComments(planCommentParts([comment]));

    expect(xml).not.toContain("<w:t>injected</w:t>");
    expect(xml).toContain('r:id="rId1&quot;/&gt;&lt;w:r&gt;&lt;w:t&gt;injected');
    expect(xml).toContain("<w:t>linked</w:t>");
  });

  test("round-trips supported paragraph and run formatting in comment content", () => {
    const comment = makeComment(1);
    const paragraph = comment.content[0]!;
    paragraph.formatting = {
      alignment: "center",
      spaceAfter: 120,
    };
    const run = paragraph.content[0]!;
    if (run.type !== "run") {
      throw new Error("Expected synthetic comment content to contain a run");
    }
    run.formatting = {
      fontSize: 24,
      underline: { style: "single" },
    };

    const reparsed = parseComments(
      serializeComments(planCommentParts([comment])),
      null,
      null,
      new Map(),
      new Map(),
    );
    const reparsedParagraph = reparsed[0]?.content[0];
    const reparsedRun = reparsedParagraph?.content.find((item) => item.type === "run");

    expect(reparsed[0]?.annotationReferenceFormatting).toBeUndefined();
    expect(reparsedParagraph?.formatting).toMatchObject(paragraph.formatting);
    expect(reparsedRun?.formatting).toMatchObject(run.formatting);
  });

  test("preserves localized annotation-reference formatting outside editable content", () => {
    const commentsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:comment w:id="1" w:author="Tester">
          <w:p>
            <w:r>
              <w:rPr><w:rStyle w:val="LocalizedCommentReference"/></w:rPr>
              <w:annotationRef/>
            </w:r>
            <w:bookmarkStart w:id="0" w:name="marker"/>
            <w:r><w:t>body</w:t></w:r>
            <w:bookmarkEnd w:id="0"/>
          </w:p>
        </w:comment>
      </w:comments>`;
    const parsed = parseComments(commentsXml, null, null, new Map(), new Map());

    expect(parsed[0]?.annotationReferenceFormatting?.styleId).toBe("LocalizedCommentReference");
    expect(parsed[0]?.content[0]?.content.at(0)?.type).toBe("bookmarkStart");

    const serialized = serializeComments(planCommentParts(parsed));
    expect(serialized.match(/<w:annotationRef\/>/gu)).toHaveLength(1);
    expect(serialized).toContain('<w:rStyle w:val="LocalizedCommentReference"/>');
    expect(parseComments(serialized, null, null, new Map(), new Map())).toEqual(parsed);
  });

  test("preserves localized annotation-reference formatting for an empty comment", () => {
    const comment: Comment = {
      id: 1,
      author: "Tester",
      annotationReferenceFormatting: { styleId: "LocalizedCommentReference" },
      content: [],
    };

    const serialized = serializeComments(planCommentParts([comment]));
    expect(serialized).toContain('<w:rStyle w:val="LocalizedCommentReference"/>');

    const reparsed = parseComments(serialized, null, null, new Map(), new Map()).at(0);
    expect(reparsed?.annotationReferenceFormatting).toEqual(comment.annotationReferenceFormatting);
    expect(reparsed?.content).toHaveLength(1);
    expect(reparsed?.content.at(0)?.content).toEqual([]);
  });
});

describe("serializeCommentsExtended", () => {
  test("keeps a thread root and its replies on their own paraIds when a later root follows", () => {
    // The construct the public corpus minimised to: comments.xml interleaves a
    // thread's replies with a later thread root, and the root's key is its LAST
    // paragraph. Writing top-level comments first reordered the part, so the
    // next parse handed `comments[]` back shuffled and every reader pairing
    // comments up by position read one comment's body under another's id.
    const root = makeComment(1);
    root.content = [
      { ...root.content[0]!, paraId: "0000AAA1" },
      { ...root.content[0]!, paraId: "0000AAA2" },
    ];
    const reply = makeComment(2, 1);
    reply.content[0]!.paraId = "0000BBB1";
    const laterRoot = makeComment(3);
    laterRoot.content[0]!.paraId = "0000CCC1";
    laterRoot.done = true;

    const xml = serializeCommentsExtended(planCommentParts([root, reply, laterRoot]));

    expect(xml).not.toBeNull();
    expect([...(xml ?? "").matchAll(/<w15:commentEx w15:paraId="([^"]+)"/gu)]).toHaveLength(3);
    expect(xml).toContain(
      '<w15:commentEx w15:paraId="0000AAA2" w15:done="0"/>' +
        '<w15:commentEx w15:paraId="0000BBB1" w15:paraIdParent="0000AAA2" w15:done="0"/>' +
        '<w15:commentEx w15:paraId="0000CCC1" w15:done="1"/>',
    );
  });

  test("escapes paraId/paraIdParent that carry markup instead of a real Word id", () => {
    const malicious = '12345678" w15:done="1"><script>alert(1)</script';
    const parent = makeComment(1);
    parent.content[0]!.paraId = malicious;
    parent.done = true;

    const xml = serializeCommentsExtended(planCommentParts([parent]));

    expect(xml).not.toBeNull();
    expect(xml).not.toContain("<script>");
    expect(xml).toContain("&quot;");
    expect(xml).toContain("&lt;script&gt;");
  });
});
