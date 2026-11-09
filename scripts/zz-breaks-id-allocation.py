# Scratch only: applies each deliberate break of the id occupancy and runs the census properties.
import subprocess,pathlib
pkg=pathlib.Path('packages/docx-core')
F=pkg/'src/ops/editorIntentCensus.ts'
tests=['src/ops/__tests__/intentCensus.property.test.ts','src/ops/__tests__/identityIndexDocuments.property.test.ts']
breaks=[
 ('none',None,None),
 ('reserved ids ignored','      return identities.has(key) || reserved.has(key);','      return identities.has(key);'),
 ('record ids ignored','      return identities.has(key) || reserved.has(key);','      return reserved.has(key);'),
 ('controls answered from revisions','    controls: inSpace(IDENTITY_SPACES.CONTROL),','    controls: inSpace(IDENTITY_SPACES.REVISION),'),
 ('paragraph ids by another spelling','    paragraphs: { has: (key) => paragraphs.has(key) },','    paragraphs: { has: (key) => paragraphs.has(key.toLowerCase()) },'),
]
for name,old,new in breaks:
    s=F.read_text()
    if old is not None:
        assert s.count(old)==1,(name,s.count(old)); F.write_text(s.replace(old,new))
    try:
        r=subprocess.run(['bun','test',*tests],cwd=pkg,capture_output=True,text=True,timeout=3000)
        out=r.stdout+r.stderr
        res=[l.strip() for l in out.splitlines() if l.strip().endswith(' pass') or l.strip().endswith(' fail')]
        print(f'BREAK {name}: {" / ".join(res)}', flush=True)
        if name=='none': print(out[-4000:], flush=True)
    finally:
        F.write_text(s)
