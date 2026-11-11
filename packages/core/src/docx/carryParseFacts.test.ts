/**
 * An adapter that rebuilds a host's paragraph around the same content
 * (unwrapping a framework's proxies) keeps the parse facts core holds for it
 * by identity.
 */

import { expect, test } from "bun:test";

import { DOCUMENT_SHAPES } from "../__tests__/documentShapes";
import type { Paragraph } from "../types/document";
import { prepareCanonicalDocxInput } from "./canonicalSessionInput";
import {
  carryParseFacts,
  getParagraphPropertySourceToken,
} from "./paragraphPropertySource";
import { visitDocxParagraphs } from "./paragraphTraversal";
import { parseDocx } from "./parser";

const parsedShape = async (id: string) => {
  const shape = DOCUMENT_SHAPES.find((candidate) => candidate.id === id);
  if (!shape) throw new Error(`No document shape ${id}`);
  return parseDocx(
    (
      await prepareCanonicalDocxInput((await shape.build()).slice().buffer)
    ).unwrap(),
    { preloadFonts: false, detectVariables: false },
  );
};

test("a rebuilt paragraph keeps its parse facts", async () => {
  const document = await parsedShape("mixed-lists");
  const paragraphs: Paragraph[] = [];
  visitDocxParagraphs(
    { documentBody: document.package.document },
    (paragraph) => paragraphs.push(paragraph),
  );
  const paragraph = paragraphs.find(
    (candidate) => getParagraphPropertySourceToken(candidate) !== undefined,
  );
  if (!paragraph) throw new Error("Expected a parsed paragraph with a token");

  const rebuilt = { ...paragraph };
  expect(getParagraphPropertySourceToken(rebuilt)).toBeUndefined();
  carryParseFacts(rebuilt, paragraph);
  expect(getParagraphPropertySourceToken(rebuilt)).toBe(
    getParagraphPropertySourceToken(paragraph),
  );
});
