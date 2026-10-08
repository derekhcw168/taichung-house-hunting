import csv, os, sys
sys.stdout.reconfigure(encoding='utf-8')

csv_props = []
with open('591看屋物件綜合比較表.csv', 'r', encoding='utf-8-sig') as f:
    for r in csv.DictReader(f):
        csv_props.append(r)

csv_map = {r['原始檔案名稱']: r for r in csv_props}

folders = ['待看物件', '看完之後不考慮的物件', '看完列入考慮的物件']
for fldr in folders:
    for fn in sorted(os.listdir(fldr)):
        if fn in csv_map:
            continue
        comm = fn.split('_')[0]
        matches = [r for r in csv_props if r['社區名稱'] == comm or comm in r['社區名稱'] or r['社區名稱'] in comm]
        if matches:
            print(f"[Match Community] {fldr}/{fn[:35]} -> CSV #{matches[0]['編號']} {matches[0]['社區名稱']}")
        else:
            print(f"[New Community]   {fldr}/{fn[:35]}")
