import { expect, test } from "bun:test";
import { parseHeader } from "./headerFooterParser";
import {
  canReplayHeaderFooterVerbatim,
  captureHeaderFooterPackageBaselines,
  getHeaderFooterSourceBaseline,
} from "./headerFooterVerbatim";
import { cloneDocumentWithParagraphPropertySources } from "./paragraphPropertySource";
import { createEmptyDocument } from "../utils/createDocument";
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
  const baseline = getHeaderFooterSourceBaseline(header);
  expect(baseline).toEqual({ type: "captured", content: structuredClone(header.content) });
  const derived = { ...header };
  expect(getHeaderFooterSourceBaseline(derived)).toEqual(baseline);
  for (const fingerprint of [
    "invalid",
    JSON.stringify({ type: "header", content: [{ type: "table", rows: null }] }),
  ]) {
    expect(getHeaderFooterSourceBaseline({ ...derived, verbatimFingerprint: fingerprint })).toEqual(
      { type: "mismatch" },
    );
  }
  expect(getHeaderFooterSourceBaseline(JSON.parse(JSON.stringify(header)))).toEqual({
    type: "missing",
  });
  const paragraph = header.content.at(0);
  if (paragraph?.type !== "paragraph") throw new Error("Missing paragraph");
  paragraph.content.push({ type: "run", content: [{ type: "text", text: "edit" }] });
  expect(getHeaderFooterSourceBaseline(header)).toEqual(baseline);
  expect(baseline).not.toEqual({ type: "captured", content: structuredClone(header.content) });
});

test("missing capture provenance serializes in full without a diagnostic", () => {
  const untracked = JSON.parse(JSON.stringify(parseHeader(HEADER)));
  const diagnostics: unknown[] = [];
  const saved = serializeHeaderFooter(untracked, {
    path: "word/header1.xml",
    bindings: new Map(),
    onDiagnostic: (diagnostic) => diagnostics.push(diagnostic),
  });
  expect(saved).toContain("w:hdr");
  expect(saved).not.toBe(HEADER);
  expect(diagnostics).toEqual([]);
});

test("forged capture provenance emits a typed diagnostic and serializes in full", () => {
  const header = parseHeader(HEADER);
  header.verbatimFingerprint = "forged";
  const diagnostics: unknown[] = [];
  const saved = serializeHeaderFooter(header, {
    path: "word/header1.xml",
    bindings: new Map(),
    onDiagnostic: (diagnostic) => diagnostics.push(diagnostic),
  });
  expect(saved).toContain("w:hdr");
  expect(saved).not.toBe(HEADER);
  expect(diagnostics).toEqual([{ type: "sourceReplayMismatch", part: "word/header1.xml" }]);
});

test("journal copies recover source ownership only within their parsed package", () => {
  const alternate = HEADER.replace("<w:p>", "\n<w:p>");
  const first = parseHeader(HEADER);
  const second = parseHeader(alternate);
  expect(first.verbatimFingerprint).toBe(second.verbatimFingerprint);
  const document = createEmptyDocument();
  document.originalBuffer = new ArrayBuffer(0);
  document.package.headers = new Map([
    ["first", first],
    ["second", second],
  ]);
  captureHeaderFooterPackageBaselines(document);
  const cloned = cloneDocumentWithParagraphPropertySources(document);
  for (const [key, header] of cloned.package.headers ?? []) {
    const original = document.package.headers?.get(key);
    expect(getHeaderFooterSourceBaseline(header, cloned.originalBuffer).type).toBe("captured");
    expect(
      serializeHeaderFooter(header, {
        path: "word/header1.xml",
        bindings: new Map(),
        originalBuffer: cloned.originalBuffer,
      }),
    ).toBe(original?.verbatimXml);
  }
  for (const header of [first, second]) {
    const restored = structuredClone(header);
    const diagnostics: unknown[] = [];
    const options = {
      path: "word/header1.xml",
      bindings: new Map<string, string>(),
      originalBuffer: document.originalBuffer,
      onDiagnostic: (diagnostic: unknown) => diagnostics.push(diagnostic),
    };
    expect(getHeaderFooterSourceBaseline(restored, document.originalBuffer).type).toBe("captured");
    expect(serializeHeaderFooter(restored, options)).toBe(header.verbatimXml);
    expect(diagnostics).toEqual([]);
    expect(getHeaderFooterSourceBaseline(restored, new ArrayBuffer(0))).toEqual({
      type: "missing",
    });
  }
  const forged = { ...first, verbatimXml: alternate };
  expect(getHeaderFooterSourceBaseline(forged, document.originalBuffer)).toEqual({
    type: "mismatch",
  });
});
