import { Result } from "better-result";
import {
  afterEach,
  beforeEach,
  describe,
  expect,
  setDefaultTimeout,
  test,
} from "bun:test";
import { existsSync } from "node:fs";
import { readdir, readFile, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";

import { ZipArchive } from "@stll/docx-core/zip";

import { DISK_SYNC_TEST_TIMEOUT_MS } from "../../../test/disk-sync-timeout";
import {
  CONTRACT_PARAGRAPHS,
  fourPageContract,
  makeTempDir,
  writeDocx,
} from "./__tests__/fixtures";
import { captureIo, dataOf, envelopeOf } from "./__tests__/io";
import { runFolioCli } from "./cli";

setDefaultTimeout(DISK_SYNC_TEST_TIMEOUT_MS);

let dir = "";
let cleanup: () => Promise<void> = () => Promise.resolve();
let file = "";

beforeEach(async () => {
  ({ dir, cleanup } = await makeTempDir());
  file = await writeDocx(dir, "contract.docx", CONTRACT_PARAGRAPHS);
});

afterEach(async () => {
  await cleanup();
});

const render = async (...args: string[]) => {
  const captured = captureIo();
  const exit = await runFolioCli(["render", file, ...args], captured.io);
  return { exit, stdout: captured.stdout() };
};

/** Whether this machine can run the PNG path: playwright-core and its Chromium. */
const chromiumAvailable = (): boolean => {
  const loaded = Result.try((): unknown =>
    createRequire(import.meta.url)("playwright-core"),
  );
  if (loaded.isErr()) return false;
  const module = loaded.value;
  if (typeof module !== "object" || module === null || !("chromium" in module))
    return false;
  const { chromium } = module;
  if (
    typeof chromium !== "object" ||
    chromium === null ||
    !("executablePath" in chromium)
  ) {
    return false;
  }
  const { executablePath } = chromium;
  if (typeof executablePath !== "function") return false;
  const executable: unknown = Result.try((): unknown =>
    executablePath.call(chromium),
  ).unwrapOr(null);
  return typeof executable === "string" && existsSync(executable);
};

const RENDER_TIMEOUT_MS = 120_000;

describe("folio render", () => {
  test(
    "writes a PDF of every page, or one page, and never touches the document",
    async () => {
      const before = await readFile(file);
      const all = await render(
        "-o",
        path.join(dir, "all.pdf"),
        "--date",
        "2026-01-02T03:04:05Z",
      );
      const one = await render("-o", path.join(dir, "one.pdf"), "--page", "1");

      expect(all.exit).toBe(0);
      expect(dataOf(all.stdout)["pageCount"]).toBe(1);
      expect(
        new TextDecoder().decode(
          (await readFile(path.join(dir, "all.pdf"))).slice(0, 5),
        ),
      ).toBe("%PDF-");
      expect(dataOf(one.stdout)["format"]).toBe("pdf");
      expect(Buffer.compare(await readFile(file), before)).toBe(0);
      expect((await readdir(dir)).toSorted()).toEqual([
        "all.pdf",
        "contract.docx",
        "one.pdf",
      ]);
    },
    RENDER_TIMEOUT_MS,
  );

  test(
    "writes HTML pages from the DOM backend",
    async () => {
      const result = await render("-o", path.join(dir, "pages.html"));

      expect(result.exit).toBe(0);
      const html = await readFile(path.join(dir, "pages.html"), "utf8");
      expect(html).toContain('class="layout-page"');
      expect(html).toContain("The buyer pays $50 on signing.");
      expect(html).not.toContain("blob:");
    },
    RENDER_TIMEOUT_MS,
  );

  test(
    "refuses missing pages, unknown formats, and existing outputs",
    async () => {
      await writeFile(path.join(dir, "taken.pdf"), "keep");
      const cases: [string[], string][] = [
        [["-o", path.join(dir, "x.pdf"), "--page", "9"], "invalid_input"],
        [["-o", path.join(dir, "x.docx")], "usage_error"],
        [["-o", path.join(dir, "x.pdf"), "--page", "0"], "usage_error"],
        [[], "usage_error"],
        [["-o", path.join(dir, "taken.pdf")], "destination_exists"],
      ];
      for (const [args, code] of cases) {
        const result = await render(...args);
        expect([
          args,
          JSON.stringify(envelopeOf(result.stdout)["error"]),
        ]).toEqual([args, expect.stringContaining(`"code":"${code}"`)]);
      }
      expect(await readFile(path.join(dir, "taken.pdf"), "utf8")).toBe("keep");
    },
    RENDER_TIMEOUT_MS,
  );

  test.skipIf(!chromiumAvailable())(
    "screenshots one page as a PNG through Chromium",
    async () => {
      const result = await render(
        "-o",
        path.join(dir, "page.png"),
        "--scale",
        "1",
      );

      expect(result.exit).toBe(0);
      const png = await readFile(path.join(dir, "page.png"));
      expect([...png.subarray(1, 4)]).toEqual([0x50, 0x4e, 0x47]);
    },
    RENDER_TIMEOUT_MS,
  );

  test(
    "renders a page range, and refuses ranges and outputs it cannot write",
    async () => {
      file = await writeDocx(dir, "long.docx", fourPageContract());
      const range = await render(
        "-o",
        path.join(dir, "range.pdf"),
        "--pages",
        "2-3",
      );
      expect(range.exit).toBe(0);
      expect(dataOf(range.stdout)).toMatchObject({
        format: "pdf",
        pageCount: 2,
      });

      const html = await render(
        "-o",
        path.join(dir, "pages"),
        "--format",
        "html",
        "--pages",
        "1,3",
      );
      expect(html.exit).toBe(0);
      expect(dataOf(html.stdout)).toMatchObject({
        format: "html",
        pageCount: 2,
      });

      const cases: string[][] = [
        ["-o", path.join(dir, "x.pdf"), "--pages", "3-1"],
        ["-o", path.join(dir, "x.pdf"), "--pages", "two"],
        ["-o", path.join(dir, "x.pdf"), "--pages", "1", "--page", "1"],
        ["-o", path.join(dir, "x.png"), "--pages", "1-2"],
        ["-o", path.join(dir, "x-{page}.pdf")],
        ["-o", path.join(dir, "x.pdf"), "--format", "png"],
        ["-o", path.join(dir, "x"), "--format", "svg"],
      ];
      for (const args of cases) {
        const result = await render(...args);
        expect([args, envelopeOf(result.stdout)["error"]]).toEqual([
          args,
          expect.objectContaining({ code: "usage_error" }),
        ]);
      }
    },
    RENDER_TIMEOUT_MS,
  );

  test(
    "shows the first-page header on page 1 and the default header after it",
    async () => {
      file = path.join(dir, "titled.docx");
      await writeFile(file, await titledDocument());
      const result = await render(
        "-o",
        path.join(dir, "titled.html"),
        "--pages",
        "1-2",
      );
      expect(result.exit).toBe(0);
      const pages = (await readFile(path.join(dir, "titled.html"), "utf8"))
        .split('class="layout-page"')
        .slice(1);
      expect(pages).toHaveLength(2);
      expect(pages[0]).toContain("First page header");
      expect(pages[0]).not.toContain("Running header");
      expect(pages[1]).toContain("Running header");
      expect(pages[1]).not.toContain("First page header");
    },
    RENDER_TIMEOUT_MS,
  );

  test.skipIf(!chromiumAvailable())(
    "writes one PNG per page when -o carries {page}",
    async () => {
      file = await writeDocx(dir, "long.docx", fourPageContract());
      const result = await render(
        "-o",
        path.join(dir, "shot-{page}.png"),
        "--pages",
        "1,3",
        "--scale",
        "1",
      );
      expect(result.exit).toBe(0);
      expect(dataOf(result.stdout)).toMatchObject({
        format: "png",
        pageCount: 2,
        outputs: [{ page: 1 }, { page: 3 }],
      });
      for (const page of [1, 3]) {
        const png = await readFile(path.join(dir, `shot-${String(page)}.png`));
        expect([...png.subarray(1, 4)]).toEqual([0x50, 0x4e, 0x47]);
      }
    },
    RENDER_TIMEOUT_MS,
  );
});

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";

/** Two pages under a section with a distinct first-page header (`w:titlePg`). */
const titledDocument = async (): Promise<Uint8Array> => {
  const zip = new ZipArchive();
  const header = (text: string): string =>
    `<w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:hdr>`;
  zip.set(
    "[Content_Types].xml",
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      `<Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/>` +
      `<Override PartName="/word/first.xml" ContentType="${WML}.header+xml"/>` +
      `<Override PartName="/word/default.xml" ContentType="${WML}.header+xml"/>` +
      "</Types>",
  );
  zip.set(
    "_rels/.rels",
    `<Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `<Relationships xmlns="${RELS_NS}">` +
      `<Relationship Id="rIdFirst" Type="${R_NS}/header" Target="first.xml"/>` +
      `<Relationship Id="rIdDefault" Type="${R_NS}/header" Target="default.xml"/>` +
      "</Relationships>",
  );
  zip.set("word/first.xml", header("First page header"));
  zip.set("word/default.xml", header("Running header"));
  zip.set(
    "word/document.xml",
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>` +
      "<w:p><w:r><w:t>Cover</w:t></w:r></w:p>" +
      '<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>Second page</w:t></w:r></w:p>' +
      '<w:sectPr><w:headerReference w:type="first" r:id="rIdFirst"/><w:headerReference w:type="default" r:id="rIdDefault"/>' +
      '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:titlePg/></w:sectPr>' +
      "</w:body></w:document>",
  );
  return await zip.generate({ compression: "store" });
};
