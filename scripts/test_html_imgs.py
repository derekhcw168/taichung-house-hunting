import os, sys, re
sys.stdout.reconfigure(encoding='utf-8')

for hf in [
    '看完列入考慮的物件/大墩人文_［降價獨家－專任捷運大三房］南屯路二段 - 信義房屋.html',
    '看完列入考慮的物件/工學天下_［工學天下，次頂樓溫馨三房］工學北路 - 信義房屋.html'
]:
    with open(hf, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    img_srcs = re.findall(r'<img[^>]+src=["\']([^"\']+)["\']', content)
    print(hf, 'found imgs:', len(img_srcs))
    for src in img_srcs[:5]:
        print('  ', src[:80])
