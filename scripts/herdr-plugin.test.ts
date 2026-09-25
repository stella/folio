import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { chmod, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

const PLUGIN_ROOT = path.resolve(import.meta.dir, "..", "plugins", "herdr");

type Manifest = {
  id: string;
  actions: { id: string; command: string[] }[];
  panes: { id: string; command: string[] }[];
  link_handlers: { id: string; pattern: string; action: string }[];
};

const isManifest = (value: unknown): value is Manifest =>
  typeof value === "object" &&
  value !== null &&
  "id" in value &&
  "actions" in value &&
  "panes" in value &&
  "link_handlers" in value;

const readManifest = async (): Promise<Manifest> => {
  const parsed: unknown = Bun.TOML.parse(
    await readFile(path.join(PLUGIN_ROOT, "herdr-plugin.toml"), "utf8"),
  );
  if (!isManifest(parsed)) throw new Error("herdr-plugin.toml lacks actions, panes, or handlers");
  return parsed;
};

/** The handler's Rust pattern as a JavaScript one: the same syntax, with `(?i)` as a flag. */
const handlerPattern = (pattern: string): RegExp =>
  pattern.startsWith("(?i)") ? new RegExp(pattern.slice(4), "iu") : new RegExp(pattern, "u");

describe("herdr plugin manifest", () => {
  test("every handler, action, and pane resolves to something the plugin ships", async () => {
    const manifest = await readManifest();
    const actionIds = new Set(manifest.actions.map(({ id }) => id));

    for (const handler of manifest.link_handlers) {
      expect(actionIds.has(handler.action)).toBe(true);
    }
    for (const { command } of [...manifest.actions, ...manifest.panes]) {
      const script = command.at(-1) ?? "";
      expect([script, existsSync(path.join(PLUGIN_ROOT, script))]).toEqual([script, true]);
    }
    const opener = await readFile(path.join(PLUGIN_ROOT, "bin", "open-docx.sh"), "utf8");
    expect(opener).toContain(`--entrypoint ${manifest.panes[0]?.id ?? "?"}`);
    expect(opener).toContain(`HERDR_PLUGIN_ID:-${manifest.id}`);
  });

  test("the link handler takes .docx paths and file links, and nothing else", async () => {
    const [handler] = (await readManifest()).link_handlers;
    const pattern = handlerPattern(handler?.pattern ?? "");

    for (const link of ["file:///tmp/a%20b.docx", "/home/me/Contract.DOCX", "file://host/x.docx"]) {
      expect([link, pattern.test(link)]).toEqual([link, true]);
    }
    for (const link of ["https://example.com/a.docx", "/tmp/a.docx.pdf", "a b.docx", "x.docx?y"]) {
      expect([link, pattern.test(link)]).toEqual([link, false]);
    }
  });
});

describe("open-docx.sh", () => {
  let dir = "";
  let stub = "";
  let record = "";

  beforeAll(async () => {
    dir = await mkdtemp(path.join(tmpdir(), "herdr-plugin-"));
    stub = path.join(dir, "herdr");
    record = path.join(dir, "args");
    await writeFile(stub, `#!/usr/bin/env bash\nprintf '%s\\n' "$@" > "${record}"\n`);
    await chmod(stub, 0o755);
  });

  afterAll(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  const open = async (link: string) => {
    await rm(record, { force: true });
    const run = spawnSync("bash", [path.join(PLUGIN_ROOT, "bin", "open-docx.sh")], {
      env: {
        PATH: process.env["PATH"],
        HERDR_BIN_PATH: stub,
        HERDR_PLUGIN_ID: "stll.folio",
        HERDR_PANE_ID: "pane-7",
        HERDR_PLUGIN_CLICKED_URL: link,
      },
      encoding: "utf8",
    });
    const args = existsSync(record) ? (await readFile(record, "utf8")).trim().split("\n") : [];
    return { status: run.status, args };
  };

  test("opens the preview pane on the decoded local path beside the clicked pane", async () => {
    const encoded = await open("file:///tmp/Master%20Agreement.docx");
    const localhost = await open("file://localhost/srv/a.docx");
    const bare = await open("/srv/b.docx");

    expect(encoded.status).toBe(0);
    expect(encoded.args).toEqual([
      "plugin",
      "pane",
      "open",
      "--plugin",
      "stll.folio",
      "--entrypoint",
      "preview",
      "--placement",
      "split",
      "--direction",
      "right",
      "--no-focus",
      "--env",
      "FOLIO_DOCX=/tmp/Master Agreement.docx",
      "--target-pane",
      "pane-7",
    ]);
    expect(localhost.args).toContain("FOLIO_DOCX=/srv/a.docx");
    expect(bare.args).toContain("FOLIO_DOCX=/srv/b.docx");
  });

  test("refuses other hosts and other files without opening a pane", async () => {
    for (const link of ["file://elsewhere.invalid/x.docx", "/srv/notes.pdf", ""]) {
      const result = await open(link);
      expect([link, result.status, result.args]).toEqual([link, 1, []]);
    }
  });
});
