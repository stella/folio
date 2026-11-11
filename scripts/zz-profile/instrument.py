# Scratch only: add per-region timers (globalThis.zzT, ms) to canonicalSession.ts of a base or head tree.
# usage: python3 -I instrument.py <tree>
import sys
p=sys.argv[1]+'/packages/core/src/controller/canonicalSession.ts'
s=open(p).read()
def rep(a,b):
    global s
    assert s.count(a)==1, a
    s=s.replace(a,b)
HELPER='''
function zzAdd(k: string, t: number): void {
  const g = globalThis as { zzT?: Record<string, number> };
  const z = (g.zzT ??= {});
  z[k] = (z[k] ?? 0) + performance.now() - t;
}
'''
if 'const stagedSources = new Map(this.sourceOwners);' in s:
    side='base'
    rep('\nconst preservePropertySources = (','\nconst preservePropertySourcesZZ = (')
    rep('    const stagedSources = new Map(this.sourceOwners);\n','    const zzOwners = performance.now();\n    const stagedSources = new Map(this.sourceOwners);\n')
    rep('      copies: owned,\n    });\n','      copies: owned,\n    });\n    zzAdd("owners", zzOwners);\n')
    rep('        this.rememberSources(applied.value.document);\n','        const zzRemember = performance.now();\n        this.rememberSources(applied.value.document);\n        zzAdd("remember", zzRemember);\n')
    wrap='const preservePropertySources = (o: PreservePropertySourcesOptions): Document => {\n  const t = performance.now();\n  const r = preservePropertySourcesZZ(o);\n  zzAdd("preserve", t);\n  return r;\n};\n'
else:
    side='head'
    rep('  preservePropertySources,\n  preservePropertySourcesWholeStories,\n','  preservePropertySources as preservePropertySourcesZZ,\n  preservePropertySourcesWholeStories,\n')
    rep("    // Every main-story paragraph the session holds is its id's owner, so\n","    const zzOwners = performance.now();\n    // Every main-story paragraph the session holds is its id's owner, so\n")
    rep('    const remembered = changedWindow(\n      this.currentDocument.package.document.content,\n      applied.value.document.package.document.content,\n    ).added;\n','    const remembered = changedWindow(\n      this.currentDocument.package.document.content,\n      applied.value.document.package.document.content,\n    ).added;\n    zzAdd("owners", zzOwners);\n')
    rep('        this.rememberSources(remembered);\n','        const zzRemember = performance.now();\n        this.rememberSources(remembered);\n        zzAdd("remember", zzRemember);\n')
    wrap='const preservePropertySources: typeof preservePropertySourcesZZ = (o) => {\n  const t = performance.now();\n  const r = preservePropertySourcesZZ(o);\n  zzAdd("preserve", t);\n  return r;\n};\n'
s=s+HELPER+wrap
open(p,'w').write(s)
print('instrumented', side, p)
