#!/usr/bin/env node
import { runFolioCli } from "./cli";

const readStdin = async (): Promise<string> => {
  const chunks: Buffer[] = [];
  for await (const chunk of process.stdin) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(String(chunk)));
  }
  return Buffer.concat(chunks).toString("utf8");
};

process.exitCode = await runFolioCli(process.argv.slice(2), {
  stdout: (text) => {
    process.stdout.write(text);
  },
  stderr: (text) => {
    process.stderr.write(text);
  },
  readStdin,
  isTTY: process.stdout.isTTY,
  env: process.env,
  cwd: process.cwd(),
});
