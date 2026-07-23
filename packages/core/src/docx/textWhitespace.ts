export const requiresXmlSpacePreserve = (text: string): boolean =>
  text.startsWith(" ") || text.endsWith(" ") || text.includes("  ");
