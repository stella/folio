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

/** How many probes may report a port another playground already holds. */
const PORT_PROBE_ATTEMPTS = 8;

/** A port the system reports free, read in a child process (no sync socket API). */
const probedPort = (): number => {
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

/**
 * A free port none of `taken` is. Each probe closes its socket before the
 * next, so the system may report the same port twice.
 */
const freePort = (taken: ReadonlySet<number>): number => {
  for (let attempt = 0; attempt < PORT_PROBE_ATTEMPTS; attempt++) {
    const port = probedPort();
    if (!taken.has(port)) return port;
  }
  throw new Error("Could not find a playground port distinct from the others.");
};

const statedPort = (name: string): number | undefined => {
  const stated = Number(process.env[name]);
  return Number.isInteger(stated) && stated > 0 ? stated : undefined;
};

const portFrom = (name: string, taken: ReadonlySet<number>): number => {
  const stated = statedPort(name);
  if (stated !== undefined) return stated;
  const port = freePort(taken);
  process.env[name] = String(port);
  return port;
};

const reactPort = portFrom(
  REACT_PORT_ENV,
  new Set([statedPort(VUE_PORT_ENV) ?? 0]),
);

export const PLAYGROUND_PORTS = {
  react: reactPort,
  vue: portFrom(VUE_PORT_ENV, new Set([reactPort])),
} as const;

/** Server startup and cold-load readiness use the same host addresses. */
export const PLAYGROUND_HOSTS = {
  react: `http://localhost:${PLAYGROUND_PORTS.react}`,
  vue: `http://localhost:${PLAYGROUND_PORTS.vue}`,
} as const;
