import {
  createReadStream,
  createWriteStream,
  mkdirSync,
  mkdtempSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pipeline } from "node:stream/promises";

import {
  protocolCaptureEnvironment,
  protocolEvidenceFiles,
} from "../tests/support/playwrightProtocolCapture";

// The runner clears its output directory before discovery; keep the live logs outside it.
const captureDirectory = mkdtempSync(
  join(tmpdir(), "folio-canonical-protocol-"),
);
const child = Bun.spawn(
  ["bunx", "playwright", "test", ...process.argv.slice(2)],
  {
    env: {
      ...process.env,
      ...protocolCaptureEnvironment(process.env, captureDirectory),
    },
    stdout: "inherit",
    stderr: "inherit",
  },
);
process.exitCode = await child.exited;
mkdirSync("test-results", { recursive: true });
const merged = "test-results/canonical-protocol.jsonl";
writeFileSync(merged, "");
for (const file of protocolEvidenceFiles(captureDirectory))
  await pipeline(
    createReadStream(file),
    createWriteStream(merged, { flags: "a" }),
  );
rmSync(captureDirectory, { recursive: true });
