# usage: python3 -I top.py <cpuprofile> <keys> <outerFn,...,rootFn> [N] (all must be on the stack; tree cut at the last): inclusive and self ms/key per function (name:line) for samples under rootFn
import json,sys,collections
d=json.load(open(sys.argv[1])); keys=int(sys.argv[2]); roots=sys.argv[3].split(','); root=roots[-1]; N=int(sys.argv[4]) if len(sys.argv)>4 else 60
nodes={n['id']:n for n in d['nodes']}; parent={}
for n in d['nodes']:
    for c in n.get('children',[]): parent[c]=n['id']
def key(n):
    cf=n['callFrame']; u=cf.get('url','').rsplit('/',1)[-1]
    return f"{cf['functionName'] or '(anon)'} {u}:{cf.get('lineNumber',0)+1}"
inc=collections.Counter(); slf=collections.Counter(); rootT=0
for s,dt in zip(d['samples'],d['timeDeltas']):
    chain=[]; x=s
    while x is not None: chain.append(x); x=parent.get(x)
    names=[nodes[c]['callFrame']['functionName'] for c in chain]
    if any(r not in names for r in roots): continue
    rootT+=dt; seen=set()
    for c in chain:
        k=key(nodes[c])
        if nodes[c]['callFrame']['functionName']==root: 
            if k not in seen: inc[k]+=dt; seen.add(k)
            break
        if k not in seen: inc[k]+=dt; seen.add(k)
    slf[key(nodes[chain[0]])]+=dt
print(f"root {root}: {rootT/1000/keys:.2f} ms/key")
print("-- inclusive"); [print(f"{v/1000/keys:7.2f} {k}") for k,v in inc.most_common(N)]
print("-- self"); [print(f"{v/1000/keys:7.2f} {k}") for k,v in slf.most_common(N//2)]
