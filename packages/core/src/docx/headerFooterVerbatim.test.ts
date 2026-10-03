import { expect, test } from "bun:test";
import { parseHeader } from "./headerFooterParser";
import {
  canReplayHeaderFooterVerbatim,
  getHeaderFooterBaselineContent,
} from "./headerFooterVerbatim";
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

test("source baselines survive spreads but never trust replaced fingerprints", () => {
  const header = parseHeader(HEADER);
  const baseline = getHeaderFooterBaselineContent(header);
  expect(baseline).toEqual(header.content);
  const derived = { ...header };
  expect(getHeaderFooterBaselineContent(derived)).toBe(baseline);
  for (const fingerprint of [
    "invalid",
    JSON.stringify({ type: "header", content: [{ type: "table", rows: null }] }),
  ]) {
    expect(
      getHeaderFooterBaselineContent({ ...derived, verbatimFingerprint: fingerprint }),
    ).toBeUndefined();
  }
  expect(getHeaderFooterBaselineContent(JSON.parse(JSON.stringify(header)))).toBeUndefined();
  const paragraph = header.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Missing paragraph");
  paragraph.content.push({ type: "run", content: [{ type: "text", text: "edit" }] });
  expect(getHeaderFooterBaselineContent(header)).toBe(baseline);
  expect(baseline).not.toEqual(header.content);
});
