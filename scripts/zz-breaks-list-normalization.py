# Scratch only: applies each deliberate break of the list normalization window and runs its property file.
import subprocess,pathlib
core=pathlib.Path('packages/core')
files={'norm':core/'src/docx/listModelNormalization.ts','mark':core/'src/docx/listMarkerComputation.ts'}
tests=['src/controller/listNormalizationCache.property.test.ts']
breaks=[
 ('none',None,None,None),
 ('resume from the checkpoint without replaying the steps after it','norm','  for (let at = from; at < index; at += 1)\n    replayListCounterStep(','  for (let at = from; at < from; at += 1)\n    replayListCounterStep('),
 ('stop at the first kept block without comparing counters','norm','      while (differing.size > 0 && index < length) {','      while (differing.size < 0 && index < length) {'),
 ('update the differing parts from the new step only','norm','        for (const part of [...partsOf(advance), ...partsOf(advanceBefore)])','        for (const part of partsOf(advance))'),
 ('compare arrays by cell, not by values','norm','      if (a === undefined || b === undefined) return a === b;\n      return (','      if (a === undefined || b === undefined) return a === b;\n      return true || ('),
 ('keep an array nothing is bound to','norm','    if (left > 0) this.bindings.set(cell, left);','    if (left >= 0) this.bindings.set(cell, left);'),
 ('replay forgets a restart','mark','  if (restarted) counters.restartedNumIds.add(numId);\n  return { numId,','  if (restarted && numId < 0) counters.restartedNumIds.add(numId);\n  return { numId,'),
 ('step keeps the previous list from before the paragraph','norm','          listCounterStep(counters, advance),\n          before,','          { ...listCounterStep(counters, advance), previousList: { abstractNumId: null, fromStyle: false, numId: null } },\n          before,'),
 ('a changed block kept as its input, not its output','norm','          output,\n          output === block ? settled : true,','          block,\n          settled,'),
 ('changed parts omit the array bound before','mark','  const abstractBefore =\n    abstractNumId === null ? undefined : abstractCounters.get(abstractNumId);\n  if (abstractNumId !== null) {','  const abstractBefore = undefined;\n  if (abstractNumId !== null) {'),
 ('a kept block reused without its registration','norm','    paragraphPropertySnapshotToken(block) === record.snapshotToken);','    true);'),
 ('normalize the whole story after a change','norm','      const differing = differingParts(live, was);','      const differing = differingParts(live, was); differing.add("x0");'),
]
for name,f,old,new in breaks:
    p=files[f] if f else None; s=p.read_text() if p else None
    if p is not None:
        assert s.count(old)==1,(name,s.count(old)); p.write_text(s.replace(old,new))
    try:
        r=subprocess.run(['bun','test',*tests],cwd=core,capture_output=True,text=True,timeout=3000)
        out=r.stdout+r.stderr
        res=[l.strip() for l in out.splitlines() if l.strip().endswith(' pass') or l.strip().endswith(' fail')]
        print(f'BREAK {name}: {" / ".join(res)}', flush=True)
        if name=='none': print(out[-6000:], flush=True)
    finally:
        if p is not None: p.write_text(s)
