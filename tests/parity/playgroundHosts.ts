const reactPort = Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200;
const vuePort = Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201;

/** Server startup and cold-load readiness use the same host addresses. */
export const PLAYGROUND_HOSTS = {
  react: `http://localhost:${reactPort}`,
  vue: `http://localhost:${vuePort}`,
} as const;

/**
 * A playground path that selects the legacy engine, for specs that test it:
 * the canonical session is the editor's default.
 */
export const withLegacyEngine = (path: string): string => {
  const [route = "/", query = ""] = path.split("?");
  const params = new URLSearchParams(query);
  params.set("session", "canonical");
  return `${route}?${params.toString()}`;
};
