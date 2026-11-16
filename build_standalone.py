#!/usr/bin/env python3
import os

base_dir = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(base_dir, 'index.html'), 'r', encoding='utf-8') as f:
    html = f.read()

with open(os.path.join(base_dir, 'style.css'), 'r', encoding='utf-8') as f:
    css = f.read()

with open(os.path.join(base_dir, 'libs', 'chart.umd.min.js'), 'r', encoding='utf-8') as f:
    chart_js = f.read()

with open(os.path.join(base_dir, 'libs', 'xlsx.full.min.js'), 'r', encoding='utf-8') as f:
    xlsx_js = f.read()

with open(os.path.join(base_dir, 'app.js'), 'r', encoding='utf-8') as f:
    app_js = f.read()

# Inline CSS
html = html.replace('<link rel="stylesheet" href="style.css?v=3.0">', f'<style>\n{css}\n</style>')
html = html.replace('<link rel="stylesheet" href="style.css">', f'<style>\n{css}\n</style>')

# Inline JS
script_replacement = f"""
  <script>
{chart_js}
  </script>
  <script>
{xlsx_js}
  </script>
  <script>
{app_js}
  </script>
"""

# Replace all external script tags
for s in ['<script src="libs/chart.umd.min.js?v=3.0"></script>',
          '<script src="libs/xlsx.full.min.js?v=3.0"></script>',
          '<script src="app.js?v=3.0"></script>',
          '<script src="libs/chart.umd.min.js"></script>',
          '<script src="libs/xlsx.full.min.js"></script>',
          '<script src="app.js"></script>']:
    html = html.replace(s, '')

html = html.replace('</body>', f'{script_replacement}\n</body>')

out_path = os.path.join(base_dir, 'financeflow-offline.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"Created standalone file: {out_path} ({os.path.getsize(out_path):,} bytes)")
