const reactPort = Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200;
const vuePort = Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201;

/** Server startup and cold-load readiness use the same host addresses. */
export const PLAYGROUND_HOSTS = {
  react: `http://localhost:${reactPort}`,
  vue: `http://localhost:${vuePort}`,
} as const;
