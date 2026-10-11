import assert from "node:assert/strict";
import { createRequire } from "node:module";

// Resolve the native core owned by our benchmark plugin, including isolated Bun installs.
const plugin = import.meta.resolve("@codspeed/tinybench-plugin");
const require = createRequire(plugin);
const { isBound, setupCore, teardownCore } = require("@codspeed/core");
assert.equal(
  isBound,
  true,
  "CodSpeed must load its published native addon on the shared Node runtime",
);
// Linux binds V8 symbols lazily: importing the addon alone misses ABI failures.
setupCore();
teardownCore();
process.stdout.write(
  `CodSpeed native core initialized on ${process.version} (${process.platform}-${process.arch}).\n`,
);
