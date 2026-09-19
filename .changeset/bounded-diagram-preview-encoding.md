---
"@stll/folio-core": patch
---

Encode the SmartArt preview raster without walking it a byte at a time. Every diagram drawing rasterises a bounded megapixel placeholder at parse time, but the PNG's CRC and Adler checksums iterated it with `for (const byte of bytes)`, and the data URI was built by spreading thirty-two thousand arguments per chunk into `String.fromCodePoint`. A profile of a four-diagram package put three quarters of the whole parse in the array-iterator protocol. The checksums now run over indices with a CRC table and a blocked Adler accumulator, the zlib stream is finished in its own buffer rather than copied into a second one the size of the raster, and base64 is written once into an ASCII array. The bytes produced are unchanged.
