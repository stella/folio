---
"@stll/folio-core": patch
---

State an application version of the form the schema fixes on every save. The extended-properties `AppVersion` in `docProps/app.xml` is `XX.YYYY`, and a value with two dots is refused by consumers outright; folio copied the part through verbatim when saving a document it had not created, so a package could carry such a value in and back out. Both save exits — the full repack and the selective save — now rewrite a value the form rejects to one derived from it alone (`1.0.0` becomes `1.0000`), leave the rest of the part byte-identical, and add extended properties to no package that lacked them.
