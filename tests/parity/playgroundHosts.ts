import { spawnSync } from "node:child_process";

/**
 * The ports the managed playgrounds serve on, the one place tests read them.
 *
 * `FOLIO_PLAYGROUND_PORT` and `FOLIO_PLAYGROUND_VUE_PORT` choose them; when
 * unset, the first process to load this module (the Playwright runner, which
 * loads the config) takes free ports from the system and writes them back to
 * the environment, so its workers and the servers it starts read the same
 * ones. Fixed defaults collided across parallel checkouts.
 */
const REACT_PORT_ENV = "FOLIO_PLAYGROUND_PORT";
const VUE_PORT_ENV = "FOLIO_PLAYGROUND_VUE_PORT";

/** A port the system reports free, read in a child process (no sync socket API). */
const freePort = (): number => {
  const probe = spawnSync(
    process.execPath,
    [
      "-e",
      'const s = require("node:net").createServer(); s.listen(0, "127.0.0.1", () => { process.stdout.write(String(s.address().port)); s.close(); });',
    ],
    { encoding: "utf8" },
  );
  const port = Number(probe.stdout);
  if (!Number.isInteger(port) || port <= 0)
    throw new Error(
      `Could not find a free playground port: ${probe.stderr || String(probe.error)}`,
    );
  return port;
};

const portFrom = (name: string): number => {
  const stated = Number(process.env[name]);
  if (Number.isInteger(stated) && stated > 0) return stated;
  const port = freePort();
  process.env[name] = String(port);
  return port;
};

export const PLAYGROUND_PORTS = {
  react: portFrom(REACT_PORT_ENV),
  vue: portFrom(VUE_PORT_ENV),
} as const;

/** Server startup and cold-load readiness use the same host addresses. */
export const PLAYGROUND_HOSTS = {
  react: `http://localhost:${PLAYGROUND_PORTS.react}`,
  vue: `http://localhost:${PLAYGROUND_PORTS.vue}`,
} as const;
