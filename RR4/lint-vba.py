"""Catch the VBA mistakes that only show up as a compile-error dialog.

Run before injecting a module into a workbook:  python RR4/lint-vba.py RR4/*.bas

Excel only compiles on first run, and a compile error opens a modal dialog -
in an automated (hidden) Excel that is a hang, not an error message. The one
that has bitten this project twice is a module-level declaration placed after
a procedure ("only comments may appear after End Sub"), so that is checked
first, plus unbalanced Sub/Function/Property blocks.
"""
import re
import sys
import glob

PROC_START = re.compile(r'^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+\w+', re.I)
PROC_END = re.compile(r'^\s*End\s+(Sub|Function|Property)\s*$', re.I)
MODULE_DECL = re.compile(r'^\s*(Public|Private|Dim|Const|Type|Enum|Declare|Option)\b', re.I)


def lint(path):
    problems = []
    raw = open(path, 'rb').read()
    text = raw.decode('cp950', errors='replace') if any(b > 127 for b in raw) else raw.decode('ascii')
    lines = text.replace('\r\n', '\n').split('\n')
    in_proc = False
    seen_proc = False
    cont = False
    for i, line in enumerate(lines, 1):
        stripped = line.strip()
        if cont:
            cont = stripped.endswith(' _')
            continue
        cont = stripped.endswith(' _')
        if in_proc:
            if PROC_END.match(line):
                in_proc = False
            continue
        if PROC_START.match(line) and ' Then ' not in line:
            in_proc = True
            seen_proc = True
            continue
        if seen_proc and MODULE_DECL.match(line) and not stripped.startswith("'"):
            problems.append((i, 'module-level declaration after a procedure', stripped[:70]))
    if in_proc:
        problems.append((len(lines), 'file ends inside a procedure (missing End Sub/Function)', ''))
    return problems


rc = 0
for pattern in sys.argv[1:]:
    for path in sorted(glob.glob(pattern)):
        probs = lint(path)
        if probs:
            rc = 1
            for ln, what, txt in probs:
                print('%s:%d: %s  %s' % (path, ln, what, txt))
        else:
            print('%s: ok' % path)
sys.exit(rc)
