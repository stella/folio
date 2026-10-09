import binaryen from "binaryen";
import { panic } from "better-result";
import { readFileSync, writeFileSync } from "node:fs";
import { brotliCompressSync, constants } from "node:zlib";

// Diagnostic builds retain names; production builds still use the canonical gate.
const analyze = (filename: string) => {
  const module = binaryen.readBinary(readFileSync(filename));
  binaryen.setOptimizeLevel(3);
  binaryen.setShrinkLevel(2);
  binaryen.setDebugInfo(false);
  module.setFeatures(
    module.getFeatures() | binaryen.Features.BulkMemory | binaryen.Features.BulkMemoryOpt,
  );
  module.optimize();
  module.optimize();
  const names: string[] = [];
  for (let index = 0; index < module.getNumFunctions(); index++) {
    const info = binaryen.getFunctionInfo(module.getFunctionByIndex(index));
    if (info.body !== 0) names.push(info.name);
  }
  const bytes = module.emitBinary();
  module.dispose();
  let position = 8;
  const unsigned = () => {
    let value = 0;
    let multiplier = 1;
    for (let count = 0; count < 5; count++) {
      const byte = bytes.at(position++);
      if (byte === undefined) return panic("Truncated diagnostic WebAssembly integer");
      value += (byte & 127) * multiplier;
      if (byte < 128) return value;
      multiplier *= 128;
    }
    return panic("Oversized diagnostic WebAssembly integer");
  };
  const functions = new Map<string, number>();
  let codeBytes = 0;
  while (position < bytes.length) {
    const section = bytes.at(position++);
    const payloadSize = unsigned();
    const payloadEnd = position + payloadSize;
    if (payloadEnd > bytes.length) return panic("Truncated diagnostic WebAssembly section");
    if (section !== 10) {
      position = payloadEnd;
      continue;
    }
    const count = unsigned();
    if (count !== names.length) return panic("Function inventory must match emitted code bodies");
    for (const name of names) {
      const size = unsigned();
      // Sum monomorphizations under the same Rust symbol, excluding its hash.
      const key = name.replace(/(?:17h[0-9a-f]{16}E|::h[0-9a-f]{16})$/u, "");
      functions.set(key, (functions.get(key) ?? 0) + size);
      codeBytes += size;
      position += size;
    }
    if (position !== payloadEnd) return panic("Diagnostic code section must be consumed exactly");
  }
  const brotliBytes = brotliCompressSync(bytes, {
    params: { [constants.BROTLI_PARAM_QUALITY]: 9 },
  }).length;
  return { wasmBytes: bytes.length, brotliBytes, codeBytes, functions };
};

const basePath = process.argv.at(2) ?? panic("Expected base diagnostic WebAssembly path");
const headPath = process.argv.at(3) ?? panic("Expected head diagnostic WebAssembly path");
const outputPath = process.argv.at(4) ?? panic("Expected footprint report path");
const base = analyze(basePath);
const head = analyze(headPath);
const names = new Set([...base.functions.keys(), ...head.functions.keys()]);
const functions = [...names]
  .map((name) => {
    const baseBytes = base.functions.get(name) ?? 0;
    const headBytes = head.functions.get(name) ?? 0;
    return { name, baseBytes, headBytes, delta: headBytes - baseBytes };
  })
  .sort((left, right) => right.delta - left.delta);
const report = {
  sources: { base: process.env.BASE_SHA, head: process.env.HEAD_SHA },
  note: "Named diagnostic builds; compare total sizes with the independently gated canonical build before assigning an allowance. Function body bytes exclude section headers, data and metadata.",
  base: { wasmBytes: base.wasmBytes, brotliBytes: base.brotliBytes, codeBytes: base.codeBytes },
  head: { wasmBytes: head.wasmBytes, brotliBytes: head.brotliBytes, codeBytes: head.codeBytes },
  delta: {
    wasmBytes: head.wasmBytes - base.wasmBytes,
    brotliBytes: head.brotliBytes - base.brotliBytes,
  },
  functions,
};
writeFileSync(outputPath, `${JSON.stringify(report, null, 2)}\n`);
console.log(JSON.stringify({ ...report, functions: functions.slice(0, 30) }, null, 2));
