---
"@stll/folio-fonts": minor
---

New package: the document fonts folio measures and paints with (Carlito, Caladea, Arimo, Tinos, Cousine, Lato, Source Sans 3 and Noto Sans Arabic), shipped as their `.woff2` and `.woff` files with a generated `fonts.css` (`.woff2`, plus the measured subsets' `.woff` as a fallback), a platform-neutral catalogue of every face and file, and a Node reader (`@stll/folio-fonts/node`) for headless layout and PDF export. Each file is exported by its path, so a bundler such as Metro can import a face directly. It has no runtime dependencies.
