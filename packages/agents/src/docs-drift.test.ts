/**
 * The README and the shipped skill file describe the default `suggest_changes`
 * operation set in prose. Pin that prose to the code so a contract change
 * cannot leave stale instructions in front of a model.
 */

import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { FOLIO_DOCUMENT_OPERATION_TYPES } from "@stll/folio-core/server";

import { DEFAULT_SUGGEST_CHANGES_OPERATION_TYPES } from "./suggest-changes-options";

const SKILL_PATH = path.join(import.meta.dir, "../skills/folio-agents/SKILL.md");
const README_PATH = path.join(import.meta.dir, "../README.md");

const mentions = (text: string, name: string): boolean =>
  new RegExp(`\\b${name}\\b`, "u").test(text);

test("SKILL.md lists exactly the default suggest_changes operation types", () => {
  const skill = readFileSync(SKILL_PATH, "utf8");
  const bullet = skill.slice(
    skill.indexOf("by default `"),
    skill.indexOf("plus comment/reply/resolve"),
  );
  for (const type of FOLIO_DOCUMENT_OPERATION_TYPES) {
    expect(mentions(bullet, type), type).toBe(
      DEFAULT_SUGGEST_CHANGES_OPERATION_TYPES.includes(type),
    );
  }
});

test("README names the excluded defaults and the per-type key map", () => {
  const readme = readFileSync(README_PATH, "utf8");
  for (const type of FOLIO_DOCUMENT_OPERATION_TYPES) {
    if (!DEFAULT_SUGGEST_CHANGES_OPERATION_TYPES.includes(type)) {
      expect(mentions(readme, type), type).toBe(true);
    }
  }
  expect(readme).toContain("FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE");
});
