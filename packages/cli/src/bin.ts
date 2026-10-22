#!/usr/bin/env node
import { panic } from "better-result";

import {
  createTextReader,
  encodeUtf8,
} from "@stll/folio-core/utils/textEncoding";

import { runFolioCli } from "./cli";

const readStdin = async (): Promise<string> => {
  const reader = createTextReader({ ignoreBOM: true });
  const chunks: string[] = [];
  for await (const chunk of process.stdin) {
    if (!(chunk instanceof Uint8Array))
      panic("Standard input must provide byte chunks");
    chunks.push(reader.read(chunk));
  }
  return chunks.join("") + reader.finish();
};

process.exitCode = await runFolioCli(process.argv.slice(2), {
  stdout: (text) => {
    process.stdout.write(encodeUtf8(text));
  },
  stderr: (text) => {
    process.stderr.write(encodeUtf8(text));
  },
  readStdin,
  isTTY: process.stdout.isTTY,
  env: process.env,
  cwd: process.cwd(),
  stdio: { input: process.stdin, output: process.stdout },
  untilInterrupted: () =>
    new Promise<void>((resolve) => {
      process.once("SIGINT", () => resolve());
      process.once("SIGTERM", () => resolve());
    }),
});
