# -*- coding: utf-8 -*-
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

benchmarks = [
    '37977_長億城攬翠區陽光區',
    '20366_工學天下',
    '20467_龍邦綠園道',
    '20453_文心高第',
    '20593_大東家華園',
    '30961_假日半島',
    '20407_文心大地',
    '38071_文心及第文心南路',
    '20472_大安黃金城大廈',
    '20587_京城特區'
]

def get_price(it):
    up = it.get('unit_price')
    if isinstance(up, dict):
        try:
            return float(up.get('price', 0))
        except:
            return 0
    elif isinstance(up, (int, float)):
        return float(up)
    elif isinstance(up, str):
        try:
            return float(up.replace('萬/坪', '').strip())
        except:
            return 0
    return 0

print("=== 10 BENCHMARK COMMUNITIES LOCAL JSON VERIFICATION ===")
for b in benchmarks:
    fname = f'scratch/data/{b}.json'
    with open(fname, 'r', encoding='utf-8') as f:
        d = json.load(f)
    items = d.get('items', [])
    prices = [get_price(it) for it in items if get_price(it) > 0]
    max_price = max(prices) if prices else 0
    latest_item = items[0] if items else {}
    lat_up = get_price(latest_item)
    print(f"社區: {d['name']} (591 ID: {d['id']})")
    print(f"  地址: {d.get('address')} | 總戶數: {d.get('households')} 戶 | 完工日: {d.get('build_date')} | 實登總筆數: {d.get('total')}")
    print(f"  最新實登成交: 日期:{latest_item.get('date')} 樓層:{latest_item.get('shift_floor')} 單價:{lat_up}萬/坪 總價:{latest_item.get('total_price_v')}萬 門牌:{latest_item.get('address')}")
    print(f"  歷史最高單價: {max_price} 萬/坪")
    
    # check 113 peak and 115 latest
    p113 = [get_price(it) for it in items if it.get('date', '').startswith('113') and get_price(it) > 0]
    p114 = [get_price(it) for it in items if it.get('date', '').startswith('114') and get_price(it) > 0]
    p115 = [get_price(it) for it in items if it.get('date', '').startswith('115') and get_price(it) > 0]
    avg113 = sum(p113)/len(p113) if p113 else 0
    avg114 = sum(p114)/len(p114) if p114 else 0
    avg115 = sum(p115)/len(p115) if p115 else 0
    print(f"  各年度均價: 113年均價={avg113:.1f}萬 ({len(p113)}筆) -> 114年均價={avg114:.1f}萬 ({len(p114)}筆) -> 115年均價={avg115:.1f}萬 ({len(p115)}筆)")
    if avg113 > 0 and avg115 > 0:
        drop = (avg115 - avg113) / avg113 * 100
        print(f"  跌幅走勢 (113峰值 vs 115現況): {drop:+.1f}%")
    print("-" * 60)
