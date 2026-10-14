// expected: violation
zip.file("word/document.xml", xml);
zip.file(path, bytes, { compression: "STORE" });
archive.zip.file(path, stripped);
zip.folder("word").file("a.xml", text);
zip["file"](path, text);
getZip().file(path, text);

// expected: allowed
zip.file("word/document.xml");
Bun.file(path);
entry.file(
  (file) => resolve(file),
  (error) => reject(error),
);
