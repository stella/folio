import { expect, test } from "bun:test";
import { parseHeader } from "./headerFooterParser";
import { canReplayHeaderFooterVerbatim } from "./headerFooterVerbatim";
import { serializeHeaderFooter } from "./serializer/headerFooterSerializer";

const HEADER =
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="start"/><w:keepNext/></w:pPr></w:p></w:hdr>';

test("reconstructed property insertion order retains the original source part", () => {
  const header = parseHeader(HEADER);
  const paragraph = header.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Missing source paragraph");
  const derived = {
    ...header,
    content: [{ ...paragraph, formatting: { keepNext: true, alignment: "start" as const } }],
  };
  expect(canReplayHeaderFooterVerbatim(derived)).toBe(true);
  expect(serializeHeaderFooter(derived)).toBe(HEADER);
});

test("changing the story kind cannot replay the previous root", () => {
  const header = parseHeader(HEADER);
  const footer = { ...header, type: "footer" as const };
  expect(canReplayHeaderFooterVerbatim(footer)).toBe(false);
  expect(serializeHeaderFooter(footer)).toContain("<w:ftr");
  expect(serializeHeaderFooter(footer)).not.toContain("<w:hdr");
});
