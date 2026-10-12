import json
import os
import re
import shlex
import subprocess
from pathlib import Path

root = Path.cwd()
output = root / 'typecheck-migration-results'
output.mkdir()
projects = json.loads((root / 'scripts/typecheck-migration-projects.json').read_text())
records = []

def measured(project, phase, checker):
    slug = project['config'].replace('/', '_') + '_' + phase + '_' + checker
    command = shlex.split(project[checker])
    env = dict(os.environ)
    env['PATH'] = str(root / project['cwd'] / 'node_modules/.bin') + ':' + str(root / 'node_modules/.bin') + ':' + env['PATH']
    result = subprocess.run(['/usr/bin/time', '-v', '-o', str(output / (slug + '.time')), *command], cwd=root / project['cwd'], env=env, text=True, capture_output=True)
    text = result.stdout + result.stderr
    (output / (slug + '.log')).write_text(text)
    timing = (output / (slug + '.time')).read_text()
    wall = re.search(r'Elapsed \(wall clock\) time \(h:mm:ss or m:ss\): (\S+)', timing).group(1)
    seconds = 0.0
    for part in wall.split(':'):
        seconds = seconds * 60 + float(part)
    return {'exit': result.returncode, 'codes': sorted(set(re.findall(r'(?:\bTS|\berror )(\d{4,5})(?=:|\b)', text))), 'count': len(re.findall(r'(?:\bTS|\berror )(\d{4,5})(?=:|\b)', text)), 'wall_seconds': seconds, 'rss_kib': int(re.search(r'Maximum resident set size \(kbytes\): (\d+)', timing).group(1))}

for project in projects:
    record = dict(project)
    record['clean'] = {checker: measured(project, 'clean', checker) for checker in ['old', 'new']}
    config = project['config']
    folder = Path(config).parent
    host = None
    original = None
    if config == 'scripts/tsconfig.budgets.json':
        seed = Path('test/typecheck/typecheck-migration.typecheck.ts')
    elif config == 'scripts/tsconfig.corpus.json':
        seed = folder / 'corpus-typecheck-migration.ts'
    elif config == 'test/tsconfig.corpus-edit-fuzz.json':
        seed = folder / 'typecheck-migration.ts'
        host = folder / 'corpus-edit-fuzz-contract.ts'
        original = host.read_bytes()
        host.write_bytes(original + b'\nimport "./typecheck-migration";\n')
    elif str(folder).startswith('packages/'):
        seed = folder / 'src/typecheck-migration.ts'
    elif config.startswith('plugins/'):
        seed = folder / 'src/typecheck-migration.ts'
    else:
        seed = folder / 'typecheck-migration.ts'
    if seed.exists():
        raise RuntimeError(f'Seed already exists: {seed}')
    seed.write_text('export const p: number = "x";\n')
    try:
        record['seed'] = {checker: measured(project, 'seed', checker) for checker in ['old', 'new']}
    finally:
        seed.unlink()
        if host is not None:
            host.write_bytes(original)
    clean = record['clean']
    seed_results = record['seed']
    record['passed'] = all(x['exit'] == 0 and x['count'] == 0 for x in clean.values()) and all(x['exit'] != 0 and '2322' in x['codes'] for x in seed_results.values()) and seed_results['old']['codes'] == seed_results['new']['codes']
    records.append(record)
    (output / 'measurements.json').write_text(json.dumps(records, indent=2) + '\n')

summary = ['# Typecheck migration measurements', '', '| Project | Clean old/new | Seed old/new | Old wall/RSS | Bun wall/RSS | Result |', '| --- | --- | --- | --- | --- | --- |']
for record in records:
    old = record['clean']['old']
    new = record['clean']['new']
    seed = record['seed']
    summary.append(f"| {record['config']} | {old['count']}/{new['count']} | {','.join(seed['old']['codes'])}/{','.join(seed['new']['codes'])} | {old['wall_seconds']}s / {old['rss_kib']} KiB | {new['wall_seconds']}s / {new['rss_kib']} KiB | {'PASS' if record['passed'] else 'FAIL'} |")
(output / 'summary.md').write_text('\n'.join(summary) + '\n')
with open(os.environ['GITHUB_STEP_SUMMARY'], 'a') as file:
    file.write('\n'.join(summary) + '\n')
raise SystemExit(0 if all(record['passed'] for record in records) else 1)
