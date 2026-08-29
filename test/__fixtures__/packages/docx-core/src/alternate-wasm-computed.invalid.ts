const GENERATED_WASM_PATH = "./generated/docx_kernel_bg.wasm";

export const alternateWasm = new URL(GENERATED_WASM_PATH, import.meta.url);
