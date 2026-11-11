# Scratch only: add a timer (globalThis.zzT, ms) around the operations contract check in applyDocumentOp.
# usage: python3 -I instrument.py <tree>
import sys
p=sys.argv[1]+'/packages/docx-core/src/ops/apply.ts'
s=open(p).read()
a='  const valid = validateOpsDocument(document);\n'
assert s.count(a)==1
s=s.replace(a,'  const zzValid = performance.now();\n'+a+'  zzAdd("validateInApply", zzValid);\n')
s+='''
function zzAdd(k: string, t: number): void {
  const g = globalThis as { zzT?: Record<string, number> };
  const z = (g.zzT ??= {});
  z[k] = (z[k] ?? 0) + performance.now() - t;
}
'''
open(p,'w').write(s)
print('instrumented', p)
