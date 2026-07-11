import os
import glob
payload = 'import sys\nfrom pathlib import Path\nsys.path.insert(0, str(Path(__file__).resolve().parent.parent))\n'
for f in glob.glob('scripts/test_*.py'):
    with open(f, 'r', encoding='utf-8') as file:
        content = file.read()
    if 'sys.path.insert' not in content:
        # insert after first line/import
        lines = content.split('\n')
        lines.insert(1, payload)
        with open(f, 'w', encoding='utf-8') as file:
            file.write('\n'.join(lines))
