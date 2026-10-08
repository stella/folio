import { expect, test } from "bun:test";
import JSZip from "jszip";
import { writeZipPart } from "./writeZipPart";

test("nested part writes create exactly their named entries with reproducible dates", async () => {
  const names = ["word/document.xml", "docProps/core.xml", "word/media/nested/image.png"];
  const first = new JSZip();
  const second = new JSZip();
  for (const zip of [first, second]) {
    for (const path of names) writeZipPart({ zip, path, data: path });
    expect(Object.keys(zip.files)).toEqual(names);
    for (const file of Object.values(zip.files)) {
      expect(file.date).toEqual(new Date(Date.UTC(2000, 0, 1)));
      expect(file.dir).toBe(false);
    }
  }
  expect(await first.generateAsync({ type: "uint8array" })).toEqual(
    await second.generateAsync({ type: "uint8array" }),
  );
});

test("explicit entry metadata is preserved without creating parents", () => {
  const zip = new JSZip();
  const date = new Date(Date.UTC(2021, 3, 5));
  writeZipPart({
    zip,
    path: "word/document.xml",
    data: "<document/>",
    options: { date, compression: "DEFLATE", compressionOptions: { level: 6 } },
  });
  expect(Object.keys(zip.files)).toEqual(["word/document.xml"]);
  expect(zip.file("word/document.xml")?.date).toEqual(date);
});
