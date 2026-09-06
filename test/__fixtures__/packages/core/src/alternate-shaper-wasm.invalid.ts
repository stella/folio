// Reaching the artifact as an asset rather than as a module is the same
// second entry point by another route.
export const alternateShaperWasm = new URL("./generated/text_shaper_bg.wasm", import.meta.url);
