---
"@stll/folio-core": patch
---

Keep a bare OMML element that sits in paragraph content. Every group that admits `m:oMath` also admits `m:EG_OMathMathElements`, so `<w:ins><m:f/></w:ins>` is a tracked insertion of a fraction with no `m:oMath` around it. The parser recognised only `m:oMath` and `m:oMathPara` and let the rest fall off the end of its `switch`, so a tracked wrapper reached disk with its content gone — a reviewer accepting an edit that is no longer there. Bare equations now travel as the markup they arrived as, exactly like the ones that have a wrapper. 76 of the survival census's pairs move from lost to kept.
