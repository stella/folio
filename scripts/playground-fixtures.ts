import fs from "node:fs";
import path from "node:path";
import type { IncomingMessage, ServerResponse } from "node:http";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const FIXTURE_PREFIX = "/fixtures/";

type FixtureMiddlewareOptions = { fixturesDir: string; cacheControl: string };

export const fixtureMiddleware =
  ({ fixturesDir, cacheControl }: FixtureMiddlewareOptions) =>
  (req: IncomingMessage, res: ServerResponse, next: () => void): void => {
    if (!req.url || !req.url.startsWith(FIXTURE_PREFIX)) {
      next();
      return;
    }
    const requested = req.url.slice(FIXTURE_PREFIX.length).split("?")[0] ?? "";
    let name: string;
    try {
      name = decodeURIComponent(requested);
    } catch {
      res.statusCode = 400;
      res.end("Invalid fixture path");
      return;
    }
    if (
      !name ||
      name !== path.basename(name) ||
      name.includes("/") ||
      name.includes("\\") ||
      name.includes("..")
    ) {
      res.statusCode = 400;
      res.end("Invalid fixture path");
      return;
    }
    fs.readFile(path.join(fixturesDir, name), (error, data) => {
      if (error) {
        res.statusCode = 404;
        res.end(`Fixture not found: ${name}`);
        return;
      }
      res.setHeader("Content-Type", DOCX_MIME);
      res.setHeader("Cache-Control", cacheControl);
      res.end(data);
    });
  };
