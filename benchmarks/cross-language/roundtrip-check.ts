/**
 * docx-rs round-trip fidelity check.
 *
 * Runs each fixture through docx-rs (read_docx → build → pack, via the
 * docx-roundtrip binary) and diffs the original vs re-serialized
 * `word/document.xml` — visible text and key element counts — to measure how
 * much docx-rs preserves. An editor's parser must round-trip losslessly; this
 * shows whether docx-rs does.
 *
 * Requires: `cargo build --release` in ./rust (builds docx-roundtrip).
 * Run: `bun benchmarks/cross-language/roundtrip-check.ts`
 */
import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, "../..");
const RT_BIN = resolve(HERE, "rust/target/release/docx-roundtrip");

const FIXTURES = [
  { size: "small", path: "tests/visual/fixtures/docx-editor-demo.docx" },
  { size: "medium", path: "tests/visual/fixtures/sample.docx" },
  { size: "large", path: "tests/visual/fixtures/podily-bps.docx" },
];

// Elements an editor must not silently drop on save.
const TAGS = [
  "w:p",
  "w:r",
  "w:t",
  "w:tbl",
  "w:tc",
  "w:drawing",
  "w:hyperlink",
  "w:sdt",
  "w:tab",
  "w:bookmarkStart",
  "w:commentReference",
  "w:footnoteReference",
];

function documentXml(docxPath: string): string {
  const result = spawnSync("unzip", ["-p", docxPath, "word/document.xml"], {
    encoding: "utf-8",
    maxBuffer: 64 * 1024 * 1024,
  });
  return result.status === 0 ? result.stdout : "";
}

function visibleText(xml: string): string {
  const matches = xml.match(/<w:t(?: [^>]*)?>[^<]*<\/w:t>/g) ?? [];
  return matches.map((m) => m.replace(/<[^>]+>/g, "")).join("");
}

function countTag(xml: string, tag: string): number {
  const escaped = tag.replace(":", "\\:");
  return (xml.match(new RegExp(`<${escaped}[ />]`, "g")) ?? []).length;
}

if (!existsSync(RT_BIN)) {
  console.error(
    "docx-roundtrip not built — run `cargo build --release` in benchmarks/cross-language/rust",
  );
  process.exit(1);
}

const scratch = mkdtempSync(join(tmpdir(), "folio-rt-"));

for (const fixture of FIXTURES) {
  const origPath = resolve(REPO_ROOT, fixture.path);
  const rtPath = join(scratch, `${fixture.size}.docx`);

  const result = spawnSync(RT_BIN, [origPath, rtPath], { encoding: "utf-8" });
  if (result.status !== 0) {
    console.log(
      `\n${fixture.size}: docx-rs round-trip FAILED — ${(result.stderr || "").slice(0, 160).trim()}`,
    );
    continue;
  }

  const origXml = documentXml(origPath);
  const rtXml = documentXml(rtPath);
  const origText = visibleText(origXml);
  const rtText = visibleText(rtXml);
  const textPct = origText.length
    ? Math.round((rtText.length / origText.length) * 100)
    : 0;

  console.log(`\n=== ${fixture.size} · ${fixture.path.split("/").pop()} ===`);
  console.log(
    `  visible text preserved: ${textPct}%  (${rtText.length} / ${origText.length} chars)`,
  );

  const rows = TAGS.map((tag) => {
    const original = countTag(origXml, tag);
    const roundTrip = countTag(rtXml, tag);
    return {
      element: tag,
      original,
      "round-trip": roundTrip,
      kept:
        original === 0 ? "—" : `${Math.round((roundTrip / original) * 100)}%`,
    };
  }).filter((row) => row.original > 0 || row["round-trip"] > 0);
  console.table(rows);
}
