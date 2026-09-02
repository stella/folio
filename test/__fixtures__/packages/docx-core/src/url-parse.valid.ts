// Plain URL parsing (link validation) has no module base and cannot load a
// kernel asset, so it is allowed anywhere in docx-core.
export const parseAbsoluteUrl = (value: string): URL => new URL(value);
