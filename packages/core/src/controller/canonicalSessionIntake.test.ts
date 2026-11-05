import { expect, test } from "bun:test";

import { fromMarkdown } from "../markdown/fromMarkdown";
import { createEmptyDocument } from "../utils/createDocument";
import { createCanonicalSession } from "./canonicalSession";

const paragraphIds = (content: readonly { type: string; paraId?: string }[]) =>
  content.flatMap((block) =>
    block.type === "paragraph" ? [block.paraId] : [],
  );

// A package reaches a session through `ensureParaIds`; a document built in
// memory did not, so a blank session refused its only paragraph.
test("a session opens a document built in memory, leaving the source as it was", () => {
  const source = createEmptyDocument({ initialText: "Blank" });
  const session = createCanonicalSession(source).unwrap();
  const [id] = paragraphIds(session.document.package.document.content);
  expect(id).toMatch(/^[0-9A-F]{8}$/u);
  expect(paragraphIds(source.package.document.content)).toEqual([undefined]);
});

test("the ids a session establishes are stable and distinct", () => {
  const read = () =>
    paragraphIds(
      createCanonicalSession(fromMarkdown("One\n\nTwo\n\nThree")).unwrap()
        .document.package.document.content,
    );
  const ids = read();
  expect(ids).toEqual(read());
  expect(new Set(ids).size).toBe(3);
});
