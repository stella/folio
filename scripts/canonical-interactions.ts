import {
  appendFileSync,
  copyFileSync,
  mkdirSync,
  mkdtempSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { canonicalProtocolEvidence } from "../tests/parity/canonicalProtocolEvidence";

// The runner clears its output directory before discovery; keep the live log outside it.
const evidenceDirectory = mkdtempSync(join(tmpdir(), "folio-canonical-protocol-"));
const evidencePath = join(evidenceDirectory, "canonical-protocol.jsonl");
writeFileSync(evidencePath, "");
const child = Bun.spawn(["bunx", "playwright", "test", ...process.argv.slice(2)], {
  env: {
    ...process.env,
    DEBUG: [process.env["DEBUG"], "pw:protocol"].filter(Boolean).join(","),
    FOLIO_CANONICAL_PROTOCOL_LOG: evidencePath,
  },
  stdout: "inherit",
  stderr: "pipe",
});

const record = (line: string) => {
  // This process boundary owns decoding failures; retain malformed evidence.
  try {
    const result = canonicalProtocolEvidence(line);
    switch (result.type) {
      case "output":
        process.stderr.write(`${result.text}\n`);
        break;
      case "unparsed":
        appendFileSync(evidencePath, `${JSON.stringify({ unparsedLength: result.length })}\n`);
        break;
      case "evidence":
        appendFileSync(evidencePath, `${JSON.stringify({ at: Date.now(), ...result.message })}\n`);
        break;
      case "discard":
        break;
    }
  } catch (error) {
    appendFileSync(
      evidencePath,
      `${JSON.stringify({ unparsedLength: line.length, error: error instanceof Error ? error.name : "UnknownError" })}\n`,
    );
  }
};
const decoder = new TextDecoder();
let pending = "";
for await (const chunk of child.stderr) {
  pending += decoder.decode(chunk, { stream: true });
  let newline = pending.indexOf("\n");
  while (newline !== -1) {
    record(pending.slice(0, newline));
    pending = pending.slice(newline + 1);
    newline = pending.indexOf("\n");
  }
}
pending += decoder.decode();
if (pending !== "") record(pending);
process.exitCode = await child.exited;
mkdirSync("test-results", { recursive: true });
copyFileSync(evidencePath, "test-results/canonical-protocol.jsonl");
rmSync(evidenceDirectory, { recursive: true });
