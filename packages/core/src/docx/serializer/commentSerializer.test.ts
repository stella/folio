import { describe, expect, test } from "bun:test";

import type { Comment } from "../../types/content";
import { parseComments } from "../commentParser";
import { serializeComments, serializeCommentsExtended } from "./commentSerializer";

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
  test("emits a valid empty <w:comments/> document when the array is empty", () => {
    // Previously returned the empty string, which is not valid OOXML.
    // Save paths now overwrite the original `word/comments.xml` part
    // even when the editor has zero comments — that requires the
    // serializer to produce a well-formed empty document so the part
    // can be replaced rather than skipped.
    const xml = serializeComments([]);
    expect(xml.startsWith("<?xml")).toBe(true);
    expect(xml).toContain("<w:comments xmlns:");
    expect(xml).toContain("</w:comments>");
    // No `<w:comment>` children.
    expect(xml).not.toContain("<w:comment ");
  });

  test("emits top-level comments before replies", () => {
    const reply = makeComment(2, 1);
    const top = makeComment(1);
    // Caller may pass replies first; the serializer must group them
    // after the top-level comments.
    const xml = serializeComments([reply, top]);
    const topIndex = xml.indexOf('w:id="1"');
    const replyIndex = xml.indexOf('w:id="2"');
    expect(topIndex).toBeGreaterThan(-1);
    expect(replyIndex).toBeGreaterThan(-1);
    expect(topIndex).toBeLessThan(replyIndex);
  });

  test("preserves an explicitly empty author attribute", () => {
    const xml = serializeComments([{ ...makeComment(1), author: "" }]);

    expect(xml).toContain('<w:comment w:id="1" w:author=""');
  });

  test("escapes a paragraph paraId that carries markup instead of a real Word id", () => {
    // A malformed/attacker-supplied `paraId` (e.g. relayed through a
    // collaboration payload) must not be able to break out of the
    // `w14:paraId="..."` attribute and inject sibling XML.
    const malicious = '12345678"/><script>alert(1)</script><w:p w14:paraId="';
    const comment = makeComment(1);
    comment.content[0]!.paraId = malicious;

    const xml = serializeComments([comment]);

    expect(xml).not.toContain("<script>");
    expect(xml).toContain("w14:paraId=");
    expect(xml).toContain("&lt;script&gt;");
    expect(xml).toContain("&quot;");
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

    const reparsed = parseComments(serializeComments([comment]), null, null, new Map(), new Map());
    const reparsedParagraph = reparsed[0]?.content[0];
    const reparsedRun = reparsedParagraph?.content.find((item) => item.type === "run");

    expect(reparsedParagraph?.formatting).toMatchObject(paragraph.formatting);
    expect(reparsedRun?.formatting).toMatchObject(run.formatting);
  });
});

describe("serializeCommentsExtended", () => {
  test("escapes paraId/paraIdParent that carry markup instead of a real Word id", () => {
    const malicious = '12345678" w15:done="1"><script>alert(1)</script';
    const parent = makeComment(1);
    parent.content[0]!.paraId = malicious;
    parent.done = true;

    const xml = serializeCommentsExtended([parent]);

    expect(xml).not.toBeNull();
    expect(xml).not.toContain("<script>");
    expect(xml).toContain("&quot;");
    expect(xml).toContain("&lt;script&gt;");
  });
});
