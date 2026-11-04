import json
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

def verify_master():
    path = 'scratch/dadongjia_valuation_master.json'
    print(f"=== Verifying {path} ===")
    assert os.path.exists(path), f"File not found: {path}"
    size = os.path.getsize(path)
    print(f"File size: {size} bytes ({size/1024:.2f} KB)")
    assert size > 20000, f"File size too small: {size} bytes"

    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 1. Target Property Checks
    tp = data.get('target_property', {})
    print("\n--- 1. Target Property Verification ---")
    print(f"Community: {tp.get('community_name')} ({tp.get('developer')})")
    print(f"Address: {tp.get('address')}")
    print(f"Floor: {tp.get('floor')}")
    print(f"Total Ping: {tp.get('total_ping')}坪 (Main: {tp.get('main_ping')}坪, Aux: {tp.get('aux_ping')}坪, Common: {tp.get('common_ping')}坪)")
    print(f"Indoor Living: {tp.get('indoor_ping')}坪 (Indoor Ratio: {tp.get('indoor_ratio')}%)")
    print(f"Public Ratio: {tp.get('public_ratio')}% (Pure Residential Public Ratio: {tp.get('pure_public_ratio')}%)")
    print(f"Car Space: {tp.get('car_space_type')} (Price: {tp.get('car_space_price')}萬, Ping: {tp.get('car_space_ping')}坪)")
    print(f"Asking Price: {tp.get('asking_price')}萬 (Gross Unit: {tp.get('asking_unit_price')}萬/坪, Net Unit: {tp.get('net_unit_price')}萬/坪)")
    print(f"Management Fee: {tp.get('management_fee')}元/月")
    print(f"Layout: {tp.get('layout')}, Orientation: {tp.get('orientation')}")
    print(f"Build Date: {tp.get('build_date')}, Age: {tp.get('age')}年")

    assert tp['community_name'] == '大東家南園'
    assert tp['floor'] == '9F/15F'
    assert tp['total_ping'] == 33.70
    assert tp['main_ping'] == 19.844
    assert tp['aux_ping'] == 1.984
    assert tp['indoor_ping'] == 21.828
    assert tp['common_ping'] == 11.872
    assert tp['pure_public_ratio'] == 29.81
    assert tp['car_space_price'] == 75.0
    assert tp['car_space_ping'] == 2.60
    assert tp['asking_price'] == 898.0
    assert tp['asking_unit_price'] == 26.65
    assert tp['net_unit_price'] == 26.46
    assert tp['management_fee'] == 1775
    assert tp['build_date'] == '1994-12-28'
    assert tp['age'] == 31.9
    print(">>> Target property assertions: PASSED")

    # 2. Historical Summary Checks
    hs = data.get('historical_summary', {})
    print("\n--- 2. Historical Summary Verification ---")
    print(f"Total Transactions: {hs.get('total_transactions')}")
    print(f"Clean Residential: {hs.get('clean_residential_count')}, Excluded: {hs.get('excluded_count')}")
    print(f"Historical Avg: {hs.get('historical_avg')}萬/坪 (Min: {hs.get('historical_min')}, Max: {hs.get('historical_max')})")
    print(f"Recent Avg (2023-2026): {hs.get('recent_avg_2023_2026')}萬/坪")
    print(f"2025 Peak Stats: Avg {hs.get('peak_2025_stats', {}).get('avg_unit_price')}萬/坪, Max Deal: {hs.get('peak_2025_stats', {}).get('highest_total_deal')}")
    print(f"2026 Latest Deal: {hs.get('latest_2026_stats', {}).get('latest_deal_date')} {hs.get('latest_2026_stats', {}).get('latest_deal_floor')} {hs.get('latest_2026_stats', {}).get('latest_deal_total_price')}萬 (單價: {hs.get('latest_2026_stats', {}).get('latest_deal_unit_price')}萬/坪, 淨單價: {hs.get('latest_2026_stats', {}).get('latest_deal_net_unit_price')}萬/坪)")

    assert hs['total_transactions'] == 77
    assert hs['clean_residential_count'] == 63
    assert hs['excluded_count'] == 14
    assert hs['historical_avg'] == 17.70
    assert hs['recent_avg_2023_2026'] == 26.05
    assert hs['peak_2025_stats']['avg_unit_price'] == 27.55
    assert hs['latest_2026_stats']['latest_deal_total_price'] == 855.0
    assert hs['latest_2026_stats']['latest_deal_unit_price'] == 25.30
    assert hs['latest_2026_stats']['latest_deal_net_unit_price'] == 25.02

    f9 = hs.get('floor_9_transactions', [])
    print(f"\n9th Floor Transactions Count: {len(f9)}")
    assert len(f9) == 8, f"Expected 8 deals on 9F, got {len(f9)}"
    found_2021_deal = False
    found_2025_deal = False
    for r in f9:
        print(f"  * {r['trans_date']} | {r['floor']} | {r['layout']} | 建坪:{r['build_area']}坪 | 總價:{r['total_price']}萬 | 單價:{r['unit_price']}萬/坪 | 淨單價:{r['net_unit_price']}萬/坪")
        if r['trans_date'] == '2021-09-20' and r['total_price'] == 700.0 and r['build_area'] == 33.7:
            found_2021_deal = True
        if r['trans_date'] == '2025-12-29' and r['total_price'] == 980.0:
            found_2025_deal = True

    assert found_2021_deal, "2021-09-20 33.70坪 700萬 deal not found on 9F"
    assert found_2025_deal, "2025-12-29 35.71坪 980萬 deal not found on 9F"
    print(">>> Historical summary assertions: PASSED")

    # 3. Comparable Communities Checks
    comps = data.get('comparable_communities', [])
    print(f"\n--- 3. Comparable Communities Verification (Total {len(comps)}) ---")
    assert len(comps) == 14, f"Expected 14 comparable communities, got {len(comps)}"
    for idx, c in enumerate(comps):
        print(f"  [{idx+1:02d}] {c['name']:<12} | 屋齡:{c['age_display']:<6} | 戶數:{c['households']:<4}戶 | 113均價:{c.get('annual_avg_113') or '-'} | 114均價:{c.get('annual_avg_114') or '-'} | 115均價:{c.get('annual_avg_115') or '-'} | 車位:{c['parking_type_and_price']}")
    print(">>> Comparable communities assertions: PASSED")

    # 4. Valuation Adjustment Matrix Checks
    vam = data.get('valuation_adjustment_matrix', {})
    dims = vam.get('dimensions', [])
    print(f"\n--- 4. Valuation Adjustment Matrix Verification ({len(dims)} dimensions) ---")
    assert len(dims) == 7, f"Expected 7 adjustment dimensions, got {len(dims)}"
    for d in dims:
        print(f"  * {d['dimension']}: {d['adjustment_vs_peers']}")
    print(">>> Valuation adjustment matrix assertions: PASSED")

    # 5. Pricing Tiers Checks
    pt = data.get('pricing_tiers', {})
    print("\n--- 5. Pricing Tiers Verification ---")
    cf = pt.get('conservative_floor', {})
    rt = pt.get('reasonable_target', {})
    cm = pt.get('ceiling_max', {})
    print(f"Conservative Floor: {cf.get('total')}萬 | 毛單價:{cf.get('unit')}萬/坪 | 淨單價:{cf.get('net_unit')}萬/坪 | 折數:{cf.get('discount')}")
    print(f"Reasonable Target: {rt.get('total_low')}~{rt.get('total_high')}萬 | 毛單價:{rt.get('unit_low')}~{rt.get('unit_high')}萬/坪 | 淨單價:{rt.get('net_unit_low')}~{rt.get('net_unit_high')}萬/坪 | 折數:{rt.get('discount_low')}~{rt.get('discount_high')}")
    print(f"Ceiling Max: {cm.get('total')}萬 | 毛單價:{cm.get('unit')}萬/坪 | 淨單價:{cm.get('net_unit')}萬/坪 | 折數:{cm.get('discount')}")

    assert cf['total'] == 810.0
    assert cf['unit'] == 24.04
    assert cf['net_unit'] == 23.63
    assert cf['discount'] == 0.902

    assert rt['total_low'] == 850.0
    assert rt['total_high'] == 860.0
    assert rt['unit_low'] == 25.22
    assert rt['unit_high'] == 25.52
    assert rt['net_unit_low'] == 24.92
    assert rt['net_unit_high'] == 25.24
    assert rt['discount_low'] == 0.947
    assert rt['discount_high'] == 0.958

    assert cm['total'] == 880.0
    assert cm['unit'] == 26.11
    assert cm['net_unit'] == 25.88
    assert cm['discount'] == 0.980

    talks = pt.get('negotiation_talking_points', [])
    print(f"\nNegotiation Talking Points ({len(talks)} points):")
    assert len(talks) == 4, f"Expected 4 talking points, got {len(talks)}"
    for idx, t in enumerate(talks):
        print(f"  Point {idx+1}: 【{t['topic']}】 {t['core_argument']}")
    print(">>> Pricing tiers assertions: PASSED")

    print("\n=======================================================")
    print("SUCCESS: dadongjia_valuation_master.json FULLY VERIFIED")
    print("=======================================================")

if __name__ == '__main__':
    verify_master()
