import { expect, test } from "bun:test";
import fc from "fast-check";
import { readFileSync } from "node:fs";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { runFolioCli } from "../packages/cli/src/cli";
import { fileVersionOf } from "../packages/cli/src/document";
import { makeTempDir } from "../packages/cli/src/__tests__/fixtures";
import { generateDocxFixture } from "../tests/support/validatedDocxFixture";
import config from "../playwright.config";
import { assertProperty, propertyConfig, propertyTestTimeout } from "../test/property-testing";
import JSZip from "jszip";
import { FolioDocxReviewer } from "../packages/core/src/server";
import { validateDocxPackage } from "../packages/docx-core/src/validate/docx";
import { buildScrollRootDocument } from "../tests/support/scrollRootDocument";
import {
  HOST_NAVIGATION_CASES,
  hostApiFlowArbitrary,
  navigationWasEffective,
} from "../tests/parity/hostApiFlow";

const key = ({ type, method }: (typeof HOST_NAVIGATION_CASES)[number]) => `${type}:${method}`;

test(
  "host flow fixtures remain schema valid through replacement and headless edits",
  async () => {
    const source = await buildScrollRootDocument();
    const zip = await JSZip.loadAsync(source);
    const part = zip.file("word/document.xml");
    if (!part) throw new Error("Missing fixture document part");
    zip.file(
      "word/document.xml",
      (await part.async("string")).replace("First page", "Replacement page"),
    );
    const replacement = await generateDocxFixture(zip, "host-flow-replacement");
    const { dir, cleanup } = await makeTempDir();
    const file = path.join(dir, "host-flow.docx");
    try {
      await assertProperty(
        fc.asyncProperty(hostApiFlowArbitrary, async ({ edits, replacementAfter }) => {
          await writeFile(file, source);
          let expectedFirstText = "First page";
          let reviewer = await FolioDocxReviewer.fromBuffer(new Uint8Array(source).buffer);
          expect(await validateDocxPackage(source)).toEqual({ valid: true });
          expect(await validateDocxPackage(replacement)).toEqual({ valid: true });
          for (const [index, text] of edits.entries()) {
            if (index === replacementAfter) {
              reviewer = await FolioDocxReviewer.fromBuffer(new Uint8Array(replacement).buffer);
              await writeFile(file, replacement);
              expectedFirstText = "Replacement page";
            }
            const batch = {
              version: 1,
              mode: "direct",
              operations: [
                {
                  id: `host-edit-${index}`,
                  type: "replaceInBlock",
                  blockId: "13300100",
                  find: "page",
                  replace: `page ${text}`,
                },
              ],
            } as const;
            expectedFirstText = expectedFirstText.replace("page", `page ${text}`);
            expect(reviewer.applyDocumentOperations(batch)).toMatchObject({
              skipped: [],
              applied: [{ id: `host-edit-${index}` }],
            });
            const before = new Uint8Array(await readFile(file));
            const stdout: string[] = [];
            const stderr: string[] = [];
            const exit = await runFolioCli(
              [
                "suggest",
                file,
                "--input",
                JSON.stringify({ operations: batch.operations }),
                "--direct",
                "--in-place",
                "--expect-version",
                fileVersionOf(before),
                "--author",
                "Parity",
                "--date",
                "2026-01-02T03:04:05Z",
                "--allow-repack",
              ],
              {
                stdout: (chunk) => stdout.push(chunk),
                stderr: (chunk) => stderr.push(chunk),
                readStdin: async () => "",
                isTTY: false,
                env: {},
                cwd: dir,
              },
            );
            expect(exit, stderr.join("")).toBe(0);
            expect(JSON.parse(stdout.join(""))).toMatchObject({
              ok: true,
              data: { status: "committed" },
            });
            const cliBytes = new Uint8Array(await readFile(file));
            expect(await validateDocxPackage(cliBytes)).toEqual({ valid: true });
            expect(fileVersionOf(cliBytes)).not.toBe(fileVersionOf(before));
            const reopened = await FolioDocxReviewer.fromBuffer(new Uint8Array(cliBytes).buffer);
            const headless = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
            expect(headless.snapshot().blocks.at(0)?.text).toBe(expectedFirstText);
            expect(reopened.snapshot().blocks).toEqual(headless.snapshot().blocks);
            expect(reopened.getChanges()).toEqual(headless.getChanges());
            expect(reopened.getComments()).toEqual(headless.getComments());
            expect(await validateDocxPackage(new Uint8Array(await reviewer.toBuffer()))).toEqual({
              valid: true,
            });
          }
        }),
        {
          numRuns: 3,
          id: "host flow fixtures remain schema valid through replacement and headless edits",
        },
      );
    } finally {
      await cleanup();
    }
  },
  propertyTestTimeout(240_000),
);

test("scroll fixture schema validation rejects each omitted required margin", async () => {
  const source = await buildScrollRootDocument();
  for (const attribute of ["top", "right", "bottom", "left", "header", "footer", "gutter"]) {
    const zip = await JSZip.loadAsync(source);
    const part = zip.file("word/document.xml");
    if (!part) throw new Error("Missing fixture document part");
    zip.file(
      "word/document.xml",
      (await part.async("string")).replace(new RegExp(` w:${attribute}="[0-9]+"`, "u"), ""),
    );
    expect(
      await validateDocxPackage(await zip.generateAsync({ type: "uint8array" })),
    ).toMatchObject({
      valid: false,
      code: "invalid_schema_attribute",
    });
  }
});

test(
  "generated host flows exercise exactly the total navigation matrix",
  () => {
    fc.assert(
      fc.property(hostApiFlowArbitrary, ({ navigation, edits, replacementAfter }) => {
        expect(navigation.map(key).sort()).toEqual(HOST_NAVIGATION_CASES.map(key).sort());
        expect(edits).toHaveLength(navigation.length);
        expect(replacementAfter).toBeGreaterThan(0);
        expect(replacementAfter).toBeLessThan(navigation.length);
      }),
      propertyConfig({ numRuns: 50 }),
    );
  },
  propertyTestTimeout(10_000),
);

const before = { scrollTop: 0, top: 2200, bottom: 2240, viewportTop: 100, viewportBottom: 600 };
const after = { scrollTop: 2000, top: 200, bottom: 240, viewportTop: 100, viewportBottom: 600 };

test("navigation oracle rejects no movement, wrong root and targets outside the viewport", () => {
  expect(navigationWasEffective({ before, after, outerBefore: 120, outerAfter: 120 })).toBe(true);
  for (const mutant of [
    { after: before, outerAfter: 120 },
    { after: { ...after, scrollTop: 0 }, outerAfter: 120 },
    { after: { ...after, top: 700, bottom: 740 }, outerAfter: 120 },
    { after: { ...after, top: 580, bottom: 620 }, outerAfter: 120 },
    { after, outerAfter: 2000 },
  ]) {
    expect(navigationWasEffective({ before, outerBefore: 120, ...mutant })).toBe(false);
  }
});

const requireRecord = (value: unknown): Record<string, unknown> => {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new Error("Workflow mapping unavailable");
  }
  return value;
};

test("random host flow is isolated to a bounded scheduled lane with preserved findings", () => {
  const file = "host-api-flow.spec.ts";
  const projects = config.projects;
  if (!projects) throw new Error("Playwright projects unavailable");
  const random = projects.find(({ name }) => name === "parity-fuzzer");
  if (!(random?.testMatch instanceof RegExp)) throw new Error("Random project match missing");
  expect(random.testMatch.test(file)).toBe(true);
  for (const name of ["parity", "vue"]) {
    const project = projects.find((candidate) => candidate.name === name);
    if (!(project?.testIgnore instanceof RegExp)) throw new Error(`${name} exclusion missing`);
    expect(project.testIgnore.test(file)).toBe(true);
  }
  const workflow = requireRecord(
    Bun.YAML.parse(readFileSync(".github/workflows/nightly-browser-input-fuzzer.yml", "utf8")),
  );
  expect(Object.keys(requireRecord(workflow["on"])).sort()).toEqual([
    "schedule",
    "workflow_dispatch",
  ]);
  const jobs = requireRecord(workflow["jobs"]);
  const job = requireRecord(jobs["browser-input-fuzzer"]);
  const rawSteps = job["steps"];
  if (!Array.isArray(rawSteps)) throw new Error("Workflow steps unavailable");
  const steps = rawSteps.map(requireRecord);
  const host = steps.find((step) => step["id"] === "host-api");
  if (!host) throw new Error("Host flow step unavailable");
  expect(host["continue-on-error"]).toBe(true);
  expect(host["timeout-minutes"]).toBeLessThanOrEqual(12);
  expect(host["run"]).toContain(`--project=parity-fuzzer tests/parity/${file}`);
  expect(host["run"]).toContain("--output=fuzz-playwright/host-api");
  expect(host["run"]).toContain("tee fuzz-artifacts/host-api/host-api-fuzz.log");
  const upload = steps.find((step) => step["name"] === "Keep host API findings");
  if (!upload) throw new Error("Host flow upload unavailable");
  expect(requireRecord(upload["with"])["path"]).toBe(
    "fuzz-artifacts/host-api\nfuzz-playwright/host-api\n",
  );
  const report = requireRecord(jobs["report-host-api"]);
  expect(report["if"]).toContain("outputs.host_api_findings == 'true'");
  expect(report["if"]).toContain("github.ref == 'refs/heads/main'");
  expect(report["timeout-minutes"]).toBeLessThanOrEqual(5);
  const reportSteps = report["steps"];
  if (!Array.isArray(reportSteps)) throw new Error("Reporter steps unavailable");
  const reporter = reportSteps
    .map(requireRecord)
    .find((step) => step["name"] === "Open or update findings");
  expect(reporter?.["run"]).toContain(
    "--log fuzz-results/fuzz-artifacts/host-api/host-api-fuzz.log",
  );
});
