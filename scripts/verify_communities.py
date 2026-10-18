# -*- coding: utf-8 -*-
import sys
import re

sys.stdout.reconfigure(encoding='utf-8')

html_path = r"南區25年以上大型社區走勢圖\index.html"
with open(html_path, encoding='utf-8') as f:
    html = f.read()

target_names = [
    '長億城', '工學天下', '龍邦綠園道', '文心高第', '大東家華園',
    '假日半島', '文心大地', '文心及第', '大安黃金城', '京城特區'
]

cards = html.split('class="community-card')

for target in target_names:
    for c in cards[1:]:
        if f'data-name="{target}' in c:
            clean_text = re.sub(r'<[^>]+>', ' ', c[:2000])
            clean_text = ' '.join(clean_text.split())
            print(f"=== {target} ===")
            print(clean_text[:500])
            break
