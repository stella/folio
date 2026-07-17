const IDENTITY_BLOCK_ATTRS = new Set(["paraId", "textId"]);

export const stripBlockIdentityAttrs = (attrs: Record<string, unknown>) => {
  const next: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(attrs)) {
    next[key] = IDENTITY_BLOCK_ATTRS.has(key) ? null : value;
  }
  return next;
};
