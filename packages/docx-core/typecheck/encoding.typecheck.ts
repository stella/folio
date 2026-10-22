import {
  createTextReader,
  decodeText,
  encodeUtf8,
} from "@stll/utf8-boundaries";

decodeText(encodeUtf8("👩🏽‍⚖️"));
const reader = createTextReader();
reader.read(new Uint8Array());

// @ts-expect-error -- Decoders receive bytes, never already-decoded text.
decodeText("text");
// @ts-expect-error -- Streaming decoders receive byte chunks.
reader.read("text chunk");
// @ts-expect-error -- Binary buffers must not pass through a text encoder.
encodeUtf8(new Uint8Array());
