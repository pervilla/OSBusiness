import re

with open(r'C:\OSBusiness FE\FORMGEN.frm', encoding='utf-8') as f:
    all_lines = f.readlines()

form_lines = all_lines[:5152]
code_lines = all_lines[5152:]
code_text = ''.join(code_lines)

# Get ALL control definitions with types and indices
controls = {}
for line in form_lines:
    m = re.search(r'Begin\s+(\S+)\s+(\S+)', line)
    if m:
        ctrl_type = m.group(1)
        ctrl_name = m.group(2)
        if ctrl_name not in controls:
            controls[ctrl_name] = {'type': ctrl_type, 'count': 0, 'indices': []}

# Count how many times each control appears in form definition
current_ctrl = None
for line in form_lines:
    m = re.search(r'Begin\s+\S+\s+(\S+)', line)
    if m:
        current_ctrl = m.group(1)
    if re.search(r'^\s+End\s*$', line):
        current_ctrl = None
    if current_ctrl:
        im = re.search(r'Index\s*=\s*(\d+)', line)
        if im and current_ctrl in controls:
            idx = int(im.group(1))
            if idx not in controls[current_ctrl]['indices']:
                controls[current_ctrl]['indices'].append(idx)

# Now do the comprehensive search
used = []
unused = []

for name in sorted(controls.keys()):
    pattern_exact = r'(?<![a-zA-Z0-9_])' + re.escape(name) + r'(?![a-zA-Z0-9_])'
    matches = []
    for i, line in enumerate(code_lines):
        line_num = i + 5153
        if re.search(pattern_exact, line, re.IGNORECASE):
            stripped = line.strip()
            matches.append((line_num, stripped[:150]))
    if matches:
        used.append((name, controls[name], matches))
    else:
        unused.append((name, controls[name]))

# Print comprehensive report
print('='*120)
print('FORMGEN.frm CONTROL USAGE ANALYSIS REPORT')
print('='*120)
print('File: C:\\OSBusiness FE\\FORMGEN.frm')
print('Form definition: Lines 1-5152')
print('VB Code-behind: Lines 5153-30247')
print('Total unique controls defined: %d' % len(controls))
print('Controls USED in code-behind: %d' % len(used))
print('Controls UNUSED in code-behind: %d' % len(unused))
print()

# Group unused by type
unused_by_type = {}
for name, info in unused:
    t = info['type']
    if t not in unused_by_type:
        unused_by_type[t] = []
    unused_by_type[t].append((name, info))

print('='*120)
print('UNUSED CONTROLS (%d controls NOT referenced in code-behind)' % len(unused))
print('='*120)
print()

for ctrl_type in sorted(unused_by_type.keys()):
    items = unused_by_type[ctrl_type]
    print('  [%s] (%d controls)' % (ctrl_type, len(items)))
    for name, info in sorted(items):
        indices = info['indices']
        if indices:
            idx_str = '  Index array: %s' % str(sorted(indices))
        else:
            idx_str = '  Single control'
        print('    - %-30s %s' % (name, idx_str))
    print()

print('='*120)
print('USED CONTROLS SUMMARY (%d controls referenced in code-behind)' % len(used))
print('='*120)
print()
header = '  %-33s %-35s %6s  %s' % ('Control Name', 'Type', 'Refs', 'First Reference')
print(header)
print('-'*120)

for name, info, matches in sorted(used, key=lambda x: -len(x[2])):
    first_line = matches[0][0]
    first_text = matches[0][1][:60]
    indices = info['indices']
    idx_str = ' [%s]' % str(sorted(indices)) if indices else ''
    print('  %-33s %-35s %5dx  L%d: %s%s' % (name, info['type'], len(matches), first_line, first_text, idx_str))

print()
print('='*120)
print('SUMMARY BY CONTROL TYPE')
print('='*120)

type_stats = {}
for name, info in controls.items():
    t = info['type']
    if t not in type_stats:
        type_stats[t] = {'total': 0, 'used': 0, 'unused': 0}
    type_stats[t]['total'] += 1

for name, info, matches in used:
    type_stats[info['type']]['used'] += 1
for name, info in unused:
    type_stats[info['type']]['unused'] += 1

print('  %-38s %6s %6s %6s' % ('Type', 'Total', 'Used', 'Unused'))
print('-'*65)
for t in sorted(type_stats.keys()):
    s = type_stats[t]
    print('  %-38s %5d %5d %5d' % (t, s['total'], s['used'], s['unused']))
print('-'*65)
print('  %-38s %5d %5d %5d' % ('TOTAL', len(controls), len(used), len(unused)))