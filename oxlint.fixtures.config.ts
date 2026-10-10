import config from "./oxlint.config.ts";

// Explicit fixture runs retain the production rule and warning policy.
// --no-ignore does not override config-level ignorePatterns in oxlint.
export default {
  ...config,
  ignorePatterns: config.ignorePatterns?.filter((pattern) => pattern !== "test/__fixtures__/**"),
};
