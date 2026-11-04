# -*- coding: utf-8 -*-
"""
Verification Script for scratch/daqing_valuation_master.json
============================================================
Checks all Interface Contracts, schema structures, data types, counts,
and numerical invariants.
"""

import json
import os
import sys

if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def verify_daqing_master():
    path = 'scratch/daqing_valuation_master.json'
    print(f"=== Starting Verification of {path} ===")
    assert os.path.exists(path), f"Master JSON file not found at {path}"

    size = os.path.getsize(path)
    print(f"File size: {size} bytes ({size/1024:.2f} KB)")
    assert size >= 20000, f"File size too small: {size} bytes"

    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 1. Target Property Verification
    print("\n--- 1. Target Property Verification ---")
    tp = data.get('target_property', {})
    print(f"Community: {tp.get('community_name')} ({tp.get('alias')}) | ID: {tp.get('id')} | Leju: {tp.get('leju_id')}")
    print(f"Address: {tp.get('address')}")
    print(f"Floor: {tp.get('floor')} (Total: {tp.get('total_floors')}F)")
    print(f"Total Ping: {tp.get('total_ping')} (Main: {tp.get('main_ping')}, Aux: {tp.get('aux_ping')}, Common: {tp.get('common_ping')})")
    print(f"Real Living: {tp.get('real_living_ping')} (Public Ratio: {tp.get('public_ratio')}%)")
    print(f"Layout: {tp.get('layout')} | Orientation: {tp.get('orientation')}")
    print(f"Car Space: {tp.get('car_space_type')}")
    print(f"Asking Price: {tp.get('asking_price')}萬 (Orig: {tp.get('original_asking_price')}萬, Cut: {tp.get('price_reduction')}萬, Unit: {tp.get('asking_unit_price')}萬/坪)")
    print(f"Age: {tp.get('age_years')}年 | Build Date: {tp.get('build_date')} | Households: {tp.get('households')} (Dorm: {tp.get('dormitory_households')})")
    print(f"Management Fee: {tp.get('management_fee')}元/月 ({tp.get('management_fee_per_ping')}元/坪/月)")

    assert tp['community_name'] == "大慶中山", f"community_name mismatch: {tp.get('community_name')}"
    assert tp['alias'] == "大慶國宅", f"alias mismatch: {tp.get('alias')}"
    assert tp['id'] == 20513, f"id mismatch: {tp.get('id')}"
    assert tp['leju_id'] == "L37277079e472", f"leju_id mismatch: {tp.get('leju_id')}"
    assert tp['address'] == "台中市南區工學二街16號 / 高工南路1號", f"address mismatch: {tp.get('address')}"
    assert tp['floor'] == "10F/12F", f"floor mismatch: {tp.get('floor')}"
    assert tp['total_floors'] == 12, f"total_floors mismatch: {tp.get('total_floors')}"
    assert tp['total_ping'] == 44.50, f"total_ping mismatch: {tp.get('total_ping')}"
    assert tp['main_ping'] == 25.365, f"main_ping mismatch: {tp.get('main_ping')}"
    assert tp['aux_ping'] == 2.405, f"aux_ping mismatch: {tp.get('aux_ping')}"
    assert tp['common_ping'] == 16.726, f"common_ping mismatch: {tp.get('common_ping')}"
    assert tp['real_living_ping'] == 27.77, f"real_living_ping mismatch: {tp.get('real_living_ping')}"
    assert tp['public_ratio'] == 37.59, f"public_ratio mismatch: {tp.get('public_ratio')}"
    assert tp['car_space_type'] == "權狀無車位 (現況承租管委會平面車位 1,500元/月)"
    assert tp['asking_price'] == 998.0, f"asking_price mismatch: {tp.get('asking_price')}"
    assert tp['original_asking_price'] == 1080.0, f"original_asking_price mismatch: {tp.get('original_asking_price')}"
    assert tp['price_reduction'] == 82.0, f"price_reduction mismatch: {tp.get('price_reduction')}"
    assert tp['asking_unit_price'] == 22.43, f"asking_unit_price mismatch: {tp.get('asking_unit_price')}"
    assert tp['age_years'] == 24.5, f"age_years mismatch: {tp.get('age_years')}"
    assert tp['build_date'] == "2002-04-01", f"build_date mismatch: {tp.get('build_date')}"
    assert tp['households'] == 362, f"households mismatch: {tp.get('households')}"
    assert tp['management_fee'] == 1500, f"management_fee mismatch: {tp.get('management_fee')}"
    assert "24年一手原況空屋" in tp['condition']
    print(">>> Target Property Verification: PASSED")

    # 2. Historical Summary Verification
    print("\n--- 2. Historical Summary Verification ---")
    hs = data.get('historical_summary', {})
    print(f"Total Transactions: {hs.get('total_transactions')}")
    print(f"Clean Residential Count: {hs.get('clean_residential_count')}")
    print(f"Excluded Count: {hs.get('excluded_count')}")
    fb = hs.get('filter_breakdown', {})
    print(f"Breakdown: 1F Shop={fb.get('first_floor_shop_count')}, Special={fb.get('special_deal_count')}, Gov Discount={fb.get('government_discount_count')}")

    assert hs['total_transactions'] == 49, f"total_transactions mismatch: {hs.get('total_transactions')}"
    assert hs['clean_residential_count'] == 33, f"clean_residential_count mismatch: {hs.get('clean_residential_count')}"
    assert hs['excluded_count'] == 16, f"excluded_count mismatch: {hs.get('excluded_count')}"
    assert fb['first_floor_shop_count'] == 4, f"first_floor_shop_count mismatch: {fb.get('first_floor_shop_count')}"
    assert fb['special_deal_count'] == 8, f"special_deal_count mismatch: {fb.get('special_deal_count')}"
    assert fb['government_discount_count'] == 4, f"government_discount_count mismatch: {fb.get('government_discount_count')}"

    # Verify Cycles
    cycles = hs.get('cycles', {})
    print("\nCycle Averages:")
    c1 = cycles.get('cycle_1_101_106', {})
    c2 = cycles.get('cycle_2_107_109', {})
    c3 = cycles.get('cycle_3_110_111', {})
    c4 = cycles.get('cycle_4_112_115', {})
    print(f"  Cycle 1 (101-106): {c1.get('count')}筆, 均價 {c1.get('mean_unit_price')}萬/坪")
    print(f"  Cycle 2 (107-109): {c2.get('count')}筆, 均價 {c2.get('mean_unit_price')}萬/坪")
    print(f"  Cycle 3 (110-111): {c3.get('count')}筆, 均價 {c3.get('mean_unit_price')}萬/坪")
    print(f"  Cycle 4 (112-115): {c4.get('count')}筆, 均價 {c4.get('mean_unit_price')}萬/坪")

    assert c1['count'] == 16 and c1['mean_unit_price'] == 12.9
    assert c2['count'] == 4 and c2['mean_unit_price'] == 15.0
    assert c3['count'] == 6 and c3['mean_unit_price'] == 18.1
    assert c4['count'] == 7 and c4['mean_unit_price'] == 22.9

    # Peak, Drop, Latest
    pk = hs.get('peak_transaction', {})
    dr = hs.get('drop_transaction', {})
    lt = hs.get('latest_transaction', {})
    print(f"\nPeak Deal: {pk.get('date')} {pk.get('floor')} {pk.get('total_price')}萬 ({pk.get('unit_price')}萬/坪, {pk.get('building_area')}坪)")
    print(f"Drop Deal: {dr.get('date')} {dr.get('floor')} {dr.get('total_price')}萬 ({dr.get('unit_price')}萬/坪, {dr.get('building_area')}坪)")
    print(f"Latest Deal: {lt.get('date')} {lt.get('floor')} {lt.get('total_price')}萬 ({lt.get('unit_price')}萬/坪, 淨單價: {lt.get('net_unit_price')}萬/坪)")

    assert pk['unit_price'] == 27.0 and pk['total_price'] == 1038.0 and pk['building_area'] == 38.44
    assert dr['unit_price'] == 20.5 and dr['total_price'] == 850.0 and dr['building_area'] == 41.45
    assert lt['unit_price'] == 22.7 and lt['total_price'] == 980.0 and lt['building_area'] == 43.17
    assert lt['net_unit_price'] == 24.18 and lt['net_price'] == 860.0 and lt['net_area'] == 35.57
    print(">>> Historical Summary Verification: PASSED")

    # 3. Car Parking Separation Verification
    print("\n--- 3. Car Parking Separation Verification ---")
    cps = data.get('car_parking_separation', {})
    print(f"Benchmark: {cps.get('benchmark_deal_date')} {cps.get('benchmark_floor')} {cps.get('gross_transaction_price')}萬 / {cps.get('gross_building_area')}坪")
    print(f"Parking: {cps.get('parking_space_price')}萬 / {cps.get('parking_space_area')}坪")
    print(f"Separated Net: {cps.get('separated_net_building_price')}萬 / {cps.get('separated_net_building_area')}坪 = {cps.get('separated_net_unit_price')}萬/坪")
    print(f"Formula: {cps.get('formula')}")

    assert cps['gross_transaction_price'] == 980.0
    assert cps['gross_building_area'] == 43.17
    assert cps['parking_space_price'] == 120.0
    assert cps['parking_space_area'] == 7.6
    assert cps['separated_net_building_price'] == 860.0
    assert cps['separated_net_building_area'] == 35.57
    assert cps['separated_net_unit_price'] == 24.18
    calculated_net_unit = round((980.0 - 120.0) / (43.17 - 7.6), 2)
    assert calculated_net_unit == 24.18, f"Calculated net unit price error: {calculated_net_unit}"
    print(">>> Car Parking Separation Verification: PASSED")

    # 4. Pricing Tiers Verification
    print("\n--- 4. Pricing Tiers Verification ---")
    pt = data.get('pricing_tiers', {})
    cf = pt.get('conservative_floor', {})
    rt = pt.get('reasonable_target', {})
    cm = pt.get('ceiling_max', {})

    print(f"Conservative Floor: {cf.get('total_low')}~{cf.get('total_high')}萬 | {cf.get('unit_low')}~{cf.get('unit_high')}萬/坪 | 折數: {cf.get('discount_low')}~{cf.get('discount_high')}")
    print(f"Reasonable Target: {rt.get('total_low')}~{rt.get('total_high')}萬 | {rt.get('unit_low')}~{rt.get('unit_high')}萬/坪 | 折數: {rt.get('discount_low')}~{rt.get('discount_high')}")
    print(f"Ceiling Max: {cm.get('total_low')}~{cm.get('total_high')}萬 | {cm.get('unit_low')}~{cm.get('unit_high')}萬/坪 | 折數: {cm.get('discount_low')}~{cm.get('discount_high')}")

    assert cf['total_low'] == 780.0 and cf['total_high'] == 800.0
    assert cf['unit_low'] == 17.53 and cf['unit_high'] == 17.98
    assert cf['discount_low'] == 0.78 and cf['discount_high'] == 0.80

    assert rt['total_low'] == 860.0 and rt['total_high'] == 890.0
    assert rt['unit_low'] == 19.33 and rt['unit_high'] == 20.00
    assert rt['discount_low'] == 0.86 and rt['discount_high'] == 0.89

    assert cm['total_low'] == 920.0 and cm['total_high'] == 930.0
    assert cm['unit_low'] == 20.67 and cm['unit_high'] == 20.90
    assert cm['discount_low'] == 0.92 and cm['discount_high'] == 0.93
    print(">>> Pricing Tiers Verification: PASSED")

    # 5. Renovation Budget Verification
    print("\n--- 5. Renovation Budget Verification ---")
    rb = data.get('renovation_budget', {})
    print(f"Total Renovation: {rb.get('total')}元 ({rb.get('total_wan')}萬元)")
    cats = rb.get('categories', [])
    print(f"Categories count: {len(cats)}")
    cat_names = [c['name'] for c in cats]
    print(f"Categories: {', '.join(cat_names)}")
    cat_sum = sum(c['amount'] for c in cats)
    print(f"Sum of categories: {cat_sum}元")

    assert rb['total'] == 1050000
    assert rb['total_wan'] == 105.0
    assert len(cats) == 7
    assert cat_sum == 1050000
    expected_categories = {
        "水電工程": 180000,
        "衛浴泥作工程": 280000,
        "衛浴設備工程": 120000,
        "廚房廚具工程": 160000,
        "門窗與油漆工程": 140000,
        "地板工程": 100000,
        "清潔與廢棄物清運": 70000
    }
    for c in cats:
        exp_amt = expected_categories.get(c['name'])
        assert exp_amt is not None, f"Unexpected category {c['name']}"
        assert c['amount'] == exp_amt, f"Category {c['name']} amount mismatch: {c['amount']} vs {exp_amt}"
    print(">>> Renovation Budget Verification: PASSED")

    # 6. Comparable Communities Verification
    print("\n--- 6. Comparable Communities Verification ---")
    comps = data.get('comparable_communities', [])
    print(f"Total Benchmark Communities: {len(comps)}")
    assert len(comps) == 6, f"Expected 6 comparable communities, got {len(comps)}"
    expected_comp_names = ["歌德堡", "長億南園", "工學天下", "大安紫金城", "劃江山", "大墩人文"]
    actual_comp_names = [c['name'] for c in comps]
    assert actual_comp_names == expected_comp_names, f"Comparable names mismatch: {actual_comp_names} vs {expected_comp_names}"

    for idx, c in enumerate(comps):
        print(f"  [{idx+1}] {c['name']:<10} | 屋齡: {c['age']:>2}年 | 戶數: {c['households']:>3}戶 | 115均價: {c.get('annual_avg_115')}萬 | 車位: {c['parking_price_str']}")
        assert c['age'] > 0
        assert c['households'] > 0
        assert c['recent_avg'] > 0
        assert c['recent_min'] > 0
        assert c['recent_max'] > 0
        assert len(c['parking_price_str']) > 0
        assert len(c['pros_cons']) > 0
    print(">>> Comparable Communities Verification: PASSED")

    # 7. Valuation Adjustment Matrix Verification
    print("\n--- 7. Valuation Adjustment Matrix Verification ---")
    vam = data.get('valuation_adjustment_matrix', {})
    dims = vam.get('dimensions', [])
    print(f"Total Adjustment Dimensions: {len(dims)}")
    assert len(dims) >= 6, f"Expected at least 6 adjustment dimensions, got {len(dims)}"
    dim_keys = [d.get('dimension_key') for d in dims]
    for k in ['age', 'floor', 'orientation_view', 'brand', 'parking', 'renovation']:
        assert k in dim_keys, f"Missing dimension key: {k}"
    for d in dims:
        print(f"  * {d['dimension']}: {d['adjustment_display']}")
    print(">>> Valuation Adjustment Matrix Verification: PASSED")

    # 8. Negotiation Scripts Verification
    print("\n--- 8. Negotiation Scripts Verification ---")
    scripts = data.get('negotiation_scripts', [])
    print(f"Total Scripts: {len(scripts)}")
    assert len(scripts) >= 5, f"Expected at least 5 scripts, got {len(scripts)}"
    for s in scripts:
        print(f"  Script {s['id']}: 【{s['title']}】 (對象: {s['target_party']})")
        assert len(s['title']) > 0
        assert len(s['target_party']) > 0
        assert len(s['core_argument']) > 0
        assert len(s['script_dialogue']) > 0
    print(">>> Negotiation Scripts Verification: PASSED")

    # 9. Historical Transactions Details Verification
    print("\n--- 9. Historical Transactions Details Verification ---")
    ht = data.get('historical_transactions', {})
    all_tx = ht.get('all_transactions', [])
    clean_tx = ht.get('clean_residential_deals', [])
    spec_tx = ht.get('special_deals', [])
    gov_tx = ht.get('government_discount_deals', [])
    shop_tx = ht.get('first_floor_shop_deals', [])

    print(f"All transactions count: {len(all_tx)}")
    print(f"Clean residential count: {len(clean_tx)}")
    print(f"Special deals count: {len(spec_tx)}")
    print(f"Government discount count: {len(gov_tx)}")
    print(f"First floor shop count: {len(shop_tx)}")

    assert len(all_tx) == 49
    assert len(clean_tx) == 33
    assert len(spec_tx) == 8
    assert len(gov_tx) == 4
    assert len(shop_tx) == 4
    assert len(clean_tx) + len(spec_tx) + len(gov_tx) + len(shop_tx) == 49
    print(">>> Historical Transactions Verification: PASSED")

    print("\n=======================================================================")
    print("ALL VERIFICATION CHECKS PASSED: daqing_valuation_master.json IS PERFECT")
    print("=======================================================================")

if __name__ == '__main__':
    verify_daqing_master()
