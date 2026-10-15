const decoder = new TextDecoder();
// expected: violation
for (const chunk of chunks) decoder.decode(chunk);
while (reader.read()) decoder.decode(chunk);
for (const chunk of chunks) new TextDecoder().decode(chunk);
stream.on("data", (chunk) => decoder.decode(chunk));

// expected: allowed
for (const chunk of chunks) decoder.decode(chunk, { stream: true });
