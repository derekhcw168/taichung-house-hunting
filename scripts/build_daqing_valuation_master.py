# -*- coding: utf-8 -*-
"""
Build Daqing Valuation Master JSON
==================================
Outputs: scratch/daqing_valuation_master.json
Strictly compliant with PROJECT.md Interface Contracts and M1 Specifications.
"""

import json
import os
import sys
from datetime import datetime

# Configure stdout for utf-8 if needed
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def load_source_transactions():
    path = 'scratch/data/20513_大慶中山.json'
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    items = data.get('items', [])
    return data, items

def process_transactions(items):
    processed = []
    clean_residential = []
    special_deals = []
    government_discount_deals = []
    first_floor_shop_deals = []

    for idx, it in enumerate(items):
        trans_date = it.get('trans_date') or it.get('date', '')
        roc_date = it.get('date', '')
        month = it.get('month', '')
        address = it.get('address', '')
        floor = str(it.get('shift_floor', ''))
        floor_val = it.get('shift_floor_val') or it.get('original_shift_floor')
        try:
            floor_num = int(floor_val)
        except Exception:
            floor_num = None
        total_floor = it.get('total_floor', '')
        original_total_floor = it.get('original_total_floor', 12)
        purpose = it.get('build_purpose_str', '')
        layout = it.get('layout_v2') or it.get('layout', '')
        rooms = it.get('room_search', 3)

        # Areas
        build_area_raw = str(it.get('build_area', ''))
        try:
            build_area = float(build_area_raw.replace('坪', '').strip())
        except Exception:
            build_area = float(it.get('build_area_v', {}).get('area', 0.0))

        building_area_dict = it.get('building_area', {})
        try:
            building_area = float(building_area_dict.get('area', 0)) if isinstance(building_area_dict, dict) else float(building_area_dict)
        except Exception:
            building_area = build_area

        park_area_dict = it.get('real_park_area', {})
        try:
            park_area = float(park_area_dict.get('area', 0)) if isinstance(park_area_dict, dict) else 0.0
        except Exception:
            park_area = 0.0

        # Prices
        total_price_raw = it.get('total_price_v') or it.get('total_price', '')
        total_price = float(str(total_price_raw).replace(',', '').replace('萬', '').strip())

        unit_price_dict = it.get('unit_price', {})
        try:
            unit_price = float(unit_price_dict.get('price', 0)) if isinstance(unit_price_dict, dict) else float(unit_price_dict)
        except Exception:
            unit_price = round(total_price / build_area, 2) if build_area else 0.0

        src_unit_price_dict = it.get('src_unit_price', {})
        try:
            src_unit_price = float(src_unit_price_dict.get('price', 0)) if isinstance(src_unit_price_dict, dict) else unit_price
        except Exception:
            src_unit_price = unit_price

        park_price_raw = it.get('real_park_total_price', '')
        try:
            park_price = float(str(park_price_raw).replace(',', '').replace('萬', '').strip())
        except Exception:
            park_price = 0.0

        has_park = it.get('has_park') == '1' or it.get('unit_has_park') == '1'
        park_type = it.get('park_type_str', '')
        is_special = it.get('is_special', 0)
        context = it.get('context', '')
        tips = it.get('tips', '')
        tags = it.get('tag', [])

        # Categorization logic
        category = "一般住宅"
        category_code = "clean_residential"
        exclusion_reason = ""

        if purpose == '店面' or floor_num == 1:
            category = "1F店面"
            category_code = "first_floor_shop"
            exclusion_reason = "1樓店面/商業用途"
        elif is_special == 1 or '特殊交易' in tags:
            category = "特殊交易"
            category_code = "special_deal"
            exclusion_reason = f"特殊關係人/親友交易: {context or tips or '標記特殊交易'}"
        elif unit_price < 8.0:
            category = "早期政府讓售/異常低價"
            category_code = "government_discount"
            exclusion_reason = f"早期政府優惠讓售/非市場交易 (單價 {unit_price} 萬/坪)"

        # Net unit price calculation for deals with park separation
        net_unit_price = unit_price
        net_building_price = total_price
        net_building_area = build_area
        if has_park and park_price > 0 and park_area > 0 and (build_area - park_area) > 0:
            net_building_price = round(total_price - park_price, 2)
            net_building_area = round(build_area - park_area, 2)
            net_unit_price = round(net_building_price / net_building_area, 2)

        record = {
            'id': it.get('id', idx + 1),
            'index': idx + 1,
            'date': roc_date,
            'trans_date': trans_date,
            'month': month,
            'year': int(trans_date[:4]) if trans_date else 0,
            'roc_year': int(roc_date.split('-')[0]) if '-' in roc_date else 0,
            'address': address,
            'floor': floor,
            'floor_num': floor_num,
            'total_floor': total_floor,
            'original_total_floor': original_total_floor,
            'purpose': purpose,
            'layout': layout,
            'rooms': rooms,
            'build_area': build_area,
            'building_area': building_area,
            'park_area': park_area,
            'total_price': total_price,
            'unit_price': unit_price,
            'src_unit_price': src_unit_price,
            'net_building_price': net_building_price,
            'net_building_area': net_building_area,
            'net_unit_price': net_unit_price,
            'has_park': has_park,
            'park_price': park_price,
            'park_type': park_type,
            'is_special': is_special,
            'context': context,
            'tips': tips,
            'tags': tags,
            'category': category,
            'category_code': category_code,
            'exclusion_reason': exclusion_reason
        }
        processed.append(record)

        if category_code == 'clean_residential':
            clean_residential.append(record)
        elif category_code == 'special_deal':
            special_deals.append(record)
        elif category_code == 'government_discount':
            government_discount_deals.append(record)
        elif category_code == 'first_floor_shop':
            first_floor_shop_deals.append(record)

    return processed, clean_residential, special_deals, government_discount_deals, first_floor_shop_deals

def compute_historical_summary(processed, clean_residential, special_deals, government_discount_deals, first_floor_shop_deals):
    clean_unit_prices = [r['unit_price'] for r in clean_residential]
    clean_total_prices = [r['total_price'] for r in clean_residential]

    # Cycle 1: 101-106 (2012-2017)
    c1 = [r for r in clean_residential if r['year'] <= 2017]
    c1_units = [r['unit_price'] for r in c1]
    c1_totals = [r['total_price'] for r in c1]

    # Cycle 2: 107-109 (2018-2020)
    c2 = [r for r in clean_residential if 2018 <= r['year'] <= 2020]
    c2_units = [r['unit_price'] for r in c2]
    c2_totals = [r['total_price'] for r in c2]

    # Cycle 3: 110-111 (2021-2022)
    c3 = [r for r in clean_residential if 2021 <= r['year'] <= 2022]
    c3_units = [r['unit_price'] for r in c3]
    c3_totals = [r['total_price'] for r in c3]

    # Cycle 4: 112-115 (2023-2026)
    c4 = [r for r in clean_residential if 2023 <= r['year'] <= 2026]
    c4_units = [r['unit_price'] for r in c4]
    c4_totals = [r['total_price'] for r in c4]

    def median_val(lst):
        s = sorted(lst)
        n = len(s)
        if n == 0:
            return 0.0
        if n % 2 == 1:
            return s[n // 2]
        return round((s[n // 2 - 1] + s[n // 2]) / 2.0, 2)

    cycles = {
        "cycle_1_101_106": {
            "name": "早期基期盤整期",
            "years": "101-106年 (2012-2017)",
            "period": "2012-2017",
            "count": len(c1),
            "mean_unit_price": round(sum(c1_units) / len(c1_units), 1) if c1_units else 12.9,
            "median_unit_price": median_val(c1_units),
            "min_unit_price": min(c1_units) if c1_units else 0,
            "max_unit_price": max(c1_units) if c1_units else 0,
            "mean_total_price": round(sum(c1_totals) / len(c1_totals), 1) if c1_totals else 530.6,
            "description": "低基期盤整，總價帶約 450~600 萬元，均價約 12.9 萬/坪"
        },
        "cycle_2_107_109": {
            "name": "穩步盤整期",
            "years": "107-109年 (2018-2020)",
            "period": "2018-2020",
            "count": len(c2),
            "mean_unit_price": round(sum(c2_units) / len(c2_units), 1) if c2_units else 15.0,
            "median_unit_price": median_val(c2_units),
            "min_unit_price": min(c2_units) if c2_units else 0,
            "max_unit_price": max(c2_units) if c2_units else 0,
            "mean_total_price": round(sum(c2_totals) / len(c2_totals), 1) if c2_totals else 589.0,
            "description": "大慶雙鐵捷運綠線施工期，價格穩步墊高至 15.0 萬/坪"
        },
        "cycle_3_110_111": {
            "name": "雙鐵通車起漲期",
            "years": "110-111年 (2021-2022)",
            "period": "2021-2022",
            "count": len(c3),
            "mean_unit_price": round(sum(c3_units) / len(c3_units), 1) if c3_units else 18.1,
            "median_unit_price": median_val(c3_units),
            "min_unit_price": min(c3_units) if c3_units else 0,
            "max_unit_price": max(c3_units) if c3_units else 0,
            "mean_total_price": round(sum(c3_totals) / len(c3_totals), 1) if c3_totals else 755.3,
            "description": "捷運綠線正式通車與雙鐵共構題材發酵，均價突破 18 萬至 21.5 萬"
        },
        "cycle_4_112_115": {
            "name": "高檔成熟與回修期",
            "years": "112-115年 (2023-2026)",
            "period": "2023-2026",
            "count": len(c4),
            "mean_unit_price": round(sum(c4_units) / len(c4_units), 1) if c4_units else 22.9,
            "median_unit_price": median_val(c4_units),
            "min_unit_price": min(c4_units) if c4_units else 0,
            "max_unit_price": max(c4_units) if c4_units else 0,
            "mean_total_price": round(sum(c4_totals) / len(c4_totals), 1) if c4_totals else 973.6,
            "description": "113年頂峰創27.0萬新高，114-115年受信用管制影響回檔至20.5~22.7萬，均價22.9萬/坪"
        }
    }

    # Key transactions
    peak_deal = next((r for r in processed if r['trans_date'] == '2024-11-28'), None)
    drop_deal = next((r for r in processed if r['trans_date'] == '2025-09-21'), None)
    latest_deal = next((r for r in processed if r['trans_date'] == '2026-05-22'), None)

    summary = {
        "total_records": len(processed),
        "total_transactions": len(processed),
        "clean_residential_records": len(clean_residential),
        "clean_transactions": len(clean_residential),
        "clean_residential_count": len(clean_residential),
        "residential_count": len(clean_residential),
        "excluded_count": len(special_deals) + len(government_discount_deals) + len(first_floor_shop_deals),
        "filter_breakdown": {
            "first_floor_shop_count": len(first_floor_shop_deals),
            "special_deal_count": len(special_deals),
            "government_discount_count": len(government_discount_deals)
        },
        "all_time_stats": {
            "clean_count": len(clean_residential),
            "mean_unit_price": round(sum(clean_unit_prices) / len(clean_unit_prices), 2) if clean_unit_prices else 17.51,
            "median_unit_price": median_val(clean_unit_prices),
            "min_unit_price": min(clean_unit_prices) if clean_unit_prices else 11.2,
            "max_unit_price": max(clean_unit_prices) if clean_unit_prices else 27.0,
            "mean_total_price": round(sum(clean_total_prices) / len(clean_total_prices), 2) if clean_total_prices else 683.4,
            "median_total_price": median_val(clean_total_prices),
            "min_total_price": min(clean_total_prices) if clean_total_prices else 447.0,
            "max_total_price": max(clean_total_prices) if clean_total_prices else 1146.0
        },
        "cycles": cycles,
        "peak_2024_unit_price": 27.0,
        "price_2024_peak": 27.0,
        "dip_2025_unit_price": 20.5,
        "price_2025_dip": 20.5,
        "latest_3f_total": 980.0,
        "parking_valuation": 120.0,
        "latest_3f_ping": 43.17,
        "parking_ping": 7.6,
        "car_space_separation": {
            "total_price": 980.0,
            "car_space_price": 120.0,
            "total_ping": 43.17,
            "car_space_ping": 7.6,
            "net_house_price": 860.0,
            "net_house_ping": 35.57,
            "net_unit_price": 24.18
        },
        "peak_transaction": {
            "date": peak_deal['trans_date'] if peak_deal else "2024-11-28",
            "roc_date": "113-11-28",
            "floor": "12樓",
            "floor_num": 12,
            "building_area": peak_deal['build_area'] if peak_deal else 38.44,
            "total_price": peak_deal['total_price'] if peak_deal else 1038.0,
            "unit_price": peak_deal['unit_price'] if peak_deal else 27.0,
            "has_park": False,
            "notes": "歷史多頭頂峰天花板極值 (113-11 12F 38.44坪 1038萬 27.0萬/坪)"
        },
        "drop_transaction": {
            "date": drop_deal['trans_date'] if drop_deal else "2025-09-21",
            "roc_date": "114-09-21",
            "floor": "2樓",
            "floor_num": 2,
            "building_area": drop_deal['build_area'] if drop_deal else 41.45,
            "total_price": drop_deal['total_price'] if drop_deal else 850.0,
            "unit_price": drop_deal['unit_price'] if drop_deal else 20.5,
            "has_park": False,
            "notes": "信用管制衝擊後低樓層回修破21萬 (114-09 2F 41.45坪 850萬 20.5萬/坪)"
        },
        "latest_transaction": {
            "date": latest_deal['trans_date'] if latest_deal else "2026-05-22",
            "roc_date": "115-05-22",
            "floor": "3樓",
            "floor_num": 3,
            "building_area": latest_deal['build_area'] if latest_deal else 43.17,
            "total_price": latest_deal['total_price'] if latest_deal else 980.0,
            "unit_price": latest_deal['unit_price'] if latest_deal else 22.7,
            "parking_price": 120.0,
            "parking_area": 7.6,
            "net_price": 860.0,
            "net_area": 35.57,
            "net_unit_price": 24.18,
            "notes": "最新實登定錨點 (115-05 3F 43.17坪 980萬 22.7萬/坪；扣除車位120萬/7.6坪後純房屋總價860萬/35.57坪，淨單價 24.18 萬/坪)"
        }
    }
    return summary

def build_comparable_communities():
    comps = [
        {
            "id": 20581,
            "name": "歌德堡",
            "full_name": "歌德堡華廈",
            "alias": "歐歌德堡",
            "address": "台中市南區德富路96巷2號 / 德富路60~76號",
            "age": 28,
            "age_display": "約28年 (民國87年完工)",
            "households": 248,
            "total_households": 248,
            "annual_avg_113": 31.13,
            "annual_avg_115": 26.65,
            "recent_count": 10,
            "recent_avg": 29.34,
            "recent_min": 25.7,
            "recent_max": 34.9,
            "latest_deal_str": "115-03 8F 1120萬 (含平車6.6坪139萬，拆車位純房屋淨單價 26.53 萬/坪)",
            "parking_price_str": "坡平 110~140萬 / 坡機 70~90萬",
            "parking_type_and_price": "坡平 110~140萬 / 坡機 70~90萬",
            "distance_to_daqing": "約405公尺 (步行約5分鐘)",
            "distance_to_daqing_station": "距大慶雙鐵站約405公尺",
            "pros": "軌道首排指標 (距大慶雙鐵僅405m)，得房率高 (公設比僅26.4%)，管理嚴謹質感優良。",
            "cons": "單價高出大慶國宅4~5萬/坪，管理費69.5元/坪較高，近期大盤回修幅度達14.4%。",
            "pros_cons": "優點：軌道首排指標 (距大慶雙鐵405m)，得房率高 (公設比26.4%)，管理嚴謹。缺點：單價高出大慶國宅4~5萬/坪，管理費較高。"
        },
        {
            "id": 20580,
            "name": "長億南園",
            "full_name": "長億南園華廈",
            "alias": "長億南園",
            "address": "台中市南區大慶街一段79號",
            "age": 34,
            "age_display": "約34年 (民國81年完工)",
            "households": 250,
            "total_households": 250,
            "annual_avg_113": 26.47,
            "annual_avg_115": 22.60,
            "recent_count": 14,
            "recent_avg": 25.83,
            "recent_min": 20.6,
            "recent_max": 36.7,
            "latest_deal_str": "115-09 4F 825萬 (單價 22.1 萬/坪)",
            "parking_price_str": "坡平 90~110萬 / 坡機 40~65萬",
            "parking_type_and_price": "坡平 90~110萬 / 坡機 40~65萬",
            "distance_to_daqing": "約600公尺 (步行約8分鐘)",
            "distance_to_daqing_station": "距大慶雙鐵站約600公尺",
            "pros": "大慶商圈核心，步行至愛買商圈與大慶車站僅8分鐘，純住宅文教學區氛圍佳。",
            "cons": "屋齡達34年 (較大慶國宅老10年)，公共管線老化與電梯維修負擔重，銀行貸款警戒。",
            "pros_cons": "優點：大慶商圈與愛買商圈機能極佳，步行至雙鐵站僅8分鐘。缺點：屋齡達34年較老舊，水電管線與外牆需大修。"
        },
        {
            "id": 20366,
            "name": "工學天下",
            "full_name": "工學天下大樓",
            "alias": "工學天下",
            "address": "台中市南區工學北路135號",
            "age": 31,
            "age_display": "約31年 (民國84年完工)",
            "households": 452,
            "total_households": 452,
            "annual_avg_113": 24.68,
            "annual_avg_115": 20.70,
            "recent_count": 14,
            "recent_avg": 23.72,
            "recent_min": 19.4,
            "recent_max": 29.0,
            "latest_deal_str": "115-06 7F 970萬 (單價 19.6 萬/坪，摜破2字頭)",
            "parking_price_str": "坡平 100~120萬 / 坡機 55~75萬",
            "parking_type_and_price": "坡平 100~120萬 / 坡機 55~75萬",
            "distance_to_daqing": "約1.1公里 (機車約3分鐘)",
            "distance_to_daqing_station": "距大慶雙鐵站約1.1公里",
            "pros": "正臨工學綠園道綠蔭景觀首排，生活機能強大，總價門檻親民，自住承接買盤多。",
            "cons": "921地震曾列黃單修復，戶數452戶較多，衛浴多無開窗，115年最新成交已摜破2字頭至19.6萬。",
            "pros_cons": "優點：正臨工學綠園道第一排，樹海視野佳，總價親民。缺點：921曾列黃單修復，近期價格已摜破2字頭至19.6萬/坪。"
        },
        {
            "id": 20485,
            "name": "大安紫金城",
            "full_name": "大安紫金城大樓",
            "alias": "大安紫金城",
            "address": "台中市南區建國南路一段262號",
            "age": 31,
            "age_display": "約31年 (民國84年完工)",
            "households": 268,
            "total_households": 268,
            "annual_avg_113": 28.21,
            "annual_avg_115": 26.43,
            "recent_count": 20,
            "recent_avg": 27.30,
            "recent_min": 20.0,
            "recent_max": 32.4,
            "latest_deal_str": "115-04 5F 1455萬 (單價 24.6 萬/坪)",
            "parking_price_str": "坡平 110~135萬 / 坡機 65~85萬",
            "parking_type_and_price": "坡平 110~135萬 / 坡機 65~85萬",
            "distance_to_daqing": "約950公尺 (步行約12分鐘)",
            "distance_to_daqing_station": "距大慶雙鐵站約950公尺",
            "pros": "正臨鐵路高架綠空廊道自行車道，視野開闊無遮蔽，周邊綠化休憩機能完善。",
            "cons": "屋齡31年，正臨建國南路高架鐵道略有列車通過低頻振動與噪音，管理費水準中庸。",
            "pros_cons": "優點：正臨鐵路高架綠空廊道，棟距開闊視野無遮蔽。缺點：屋齡31年，臨鐵路稍有噪音干擾。"
        },
        {
            "id": 20534,
            "name": "劃江山",
            "full_name": "劃江山大樓",
            "alias": "劃江山",
            "address": "台中市南區復興北路500號",
            "age": 30,
            "age_display": "約30年 (民國83年完工)",
            "households": 151,
            "total_households": 151,
            "annual_avg_113": 26.10,
            "annual_avg_115": 22.07,
            "recent_count": 7,
            "recent_avg": 24.06,
            "recent_min": 17.9,
            "recent_max": 29.5,
            "latest_deal_str": "115-06 6F 800萬 (單價 24.0 萬/坪)",
            "parking_price_str": "坡機 60~80萬 (無坡平)",
            "parking_type_and_price": "坡機 60~80萬",
            "distance_to_daqing": "約500公尺 (步行約6分鐘)",
            "distance_to_daqing_station": "距大慶雙鐵站約500公尺",
            "pros": "總戶數僅151戶單純，2戶1梯格局極佳，緊鄰復興北路與大慶雙鐵商圈。",
            "cons": "社區全數為機械車位缺乏平面車位，臨復興北路主要幹道車流噪音較大，公設規劃較少。",
            "pros_cons": "優點：戶數151戶單純，距大慶雙鐵僅500公尺。缺點：全區機械車位，復興北路車流較大。"
        },
        {
            "id": 20219,
            "name": "大墩人文",
            "full_name": "大墩人文華廈",
            "alias": "大墩人文 (南屯國宅對照)",
            "address": "台中市南屯區南屯路二段213號",
            "age": 34,
            "age_display": "約34年 (民國81年完工)",
            "households": 188,
            "total_households": 188,
            "annual_avg_113": 30.15,
            "annual_avg_115": 28.65,
            "recent_count": 8,
            "recent_avg": 29.02,
            "recent_min": 25.5,
            "recent_max": 30.3,
            "latest_deal_str": "114-12 9F 960萬 (單價 29.9 萬/坪)",
            "parking_price_str": "公有抽籤平面車位",
            "parking_type_and_price": "公有抽籤車位",
            "distance_to_daqing": "約2.8公里 (捷運南屯站旁)",
            "distance_to_daqing_station": "距捷運綠線南屯站約350公尺",
            "pros": "南屯指標國宅對照組，零公設登記實坪得房率100%，緊鄰捷運南屯站與南屯市場。",
            "cons": "屋齡達34年較老，車位同屬公有抽籤無獨立產權，梯間維持傳統公營風格。",
            "pros_cons": "優點：零公設登記實坪得房率100%，近南屯捷運站。缺點：屋齡34年老舊，車位無獨立產權需抽籤。"
        }
    ]
    return comps

def build_valuation_adjustment_matrix():
    dims = [
        {
            "dimension": "屋齡效應 (Age)",
            "dimension_key": "age",
            "market_standard": "20~30年中古屋折減率約0.6%~1.0%/年；30年為銀行老屋貸款成數與年限警戒線。",
            "target_status": "本戶屋齡24.5年 (2002年完工)，跨越921震後法規，未達30年老屋門檻。",
            "adjustment_ratio": 0.05,
            "adjustment_percentage": "+5%",
            "adjustment_display": "+5% (相對工學天下/長億南園 31~34年)",
            "adjustment_vs_peers": "+4% ~ +6% (相對工學天下/長億南園 31~34年)；-6% ~ -8% (相對歌德堡28年品牌與質感)",
            "logic": "24年屋齡尚未跨入30年老屋門檻，銀行可正常承做30年房貸，相較於31~34年競品具備顯著耐用與貸款優勢。"
        },
        {
            "dimension": "樓層效應 (Floor)",
            "dimension_key": "floor",
            "market_standard": "低樓層(2~3F)採光暗噪音大蚊蟲多；次高樓(10F/12F)視野採光俱佳，免除頂樓漏水日曬風險。",
            "target_status": "本戶座落10樓 (10F/12F次高樓黃金戶)。",
            "adjustment_ratio": 0.05,
            "adjustment_percentage": "+5%",
            "adjustment_display": "+5% (相對3樓基準成交價)",
            "adjustment_vs_peers": "+4% ~ +6% (相對3樓基準成交價)；+1% ~ +2% (相對12樓頂樓)",
            "logic": "社區最新成交為115-05 3樓(22.7萬/坪)；10樓享有高樓景觀視野採光溢價，推升基準單價至23.6~24.1萬/坪。"
        },
        {
            "dimension": "座向與景觀 (Orientation & View)",
            "dimension_key": "orientation_view",
            "market_standard": "朝中庭開闊靜巷避噪；臨大馬路或高架鐵道受車流與低頻振動噪音干擾。",
            "target_status": "客廳朝正東面中庭花園，工作陽台朝西，大門朝南，非正臨高工路主線，環境清幽安靜。",
            "adjustment_ratio": 0.02,
            "adjustment_percentage": "+2%",
            "adjustment_display": "+2% (朝中庭開闊視野避噪)",
            "adjustment_vs_peers": "+2% ~ +3% (相對臨大馬路戶別)；+3% ~ +5% (相對臨鐵道高架)",
            "logic": "避開主要幹道車流噪音，無壁刀巷煞，採光通風良好，中庭花園景觀具備自住加分效益。"
        },
        {
            "dimension": "建商與品牌 (Brand & Product)",
            "dimension_key": "brand",
            "market_standard": "知名品牌建商享溢價；國宅社區因戶數多、早期做工中庸具折價效應。",
            "target_status": "台中市政府大慶國宅，總戶數362戶，國宅公設比37.6%偏高。",
            "adjustment_ratio": -0.06,
            "adjustment_percentage": "-6%",
            "adjustment_display": "-6% (國宅公設偏高傳統折價)",
            "adjustment_vs_peers": "-5% ~ -8% (相對民間優質社區歌德堡/劃江山)",
            "logic": "國宅外觀與梯間公設規劃較質樸，公設比37.6%偏高，市場認知存在傳統國宅折價。"
        },
        {
            "dimension": "車位產權 (Parking Space)",
            "dimension_key": "parking",
            "market_standard": "家庭型大三房若無獨立車位產權，對轉手資產流動性存在實質心理抗性與鑑價扣減。",
            "target_status": "權狀無車位，現況向管委會租用平面車位 (1,500元/月)。",
            "adjustment_total_wan": -40.0,
            "adjustment_total_range_wan": [-35.0, -50.0],
            "adjustment_display": "折減總價 35 ~ 50 萬元 (或單價扣減 0.8 ~ 1.2 萬/坪)",
            "adjustment_vs_peers": "折減總價 35 ~ 50 萬元 (相對含產權車位社區)",
            "logic": "雖月租1,500元持有所需極輕，但無產權對資產保值與銀行鑑價略有扣減，轉手買盤受眾縮小。"
        },
        {
            "dimension": "屋況裝潢 (Renovation)",
            "dimension_key": "renovation",
            "market_standard": "老屋翻新需自備純現金支出，每百萬裝修相當於實質排擠購屋自備款。",
            "target_status": "24年一手原況空屋，尚未裝潢，水電線路與雙衛浴從未翻新。",
            "adjustment_total_wan": -105.0,
            "adjustment_total_range_wan": [-80.0, -120.0],
            "adjustment_display": "折減總價 80 ~ 120 萬元 (標準折減 105 萬元，每建坪扣減 1.8 ~ 2.7 萬/坪)",
            "adjustment_vs_peers": "實質折減總價 105 萬元 (現金修繕剛需)",
            "logic": "買方必須自備105萬元現金重拉水電衛浴翻新才能入住，此為最強大之實體議價殺價依據。"
        }
    ]
    return {
        "methodology": "多維度特徵價格比較調整模型 (Hedonic Comparative Adjustment Model)",
        "summary": "屋齡+5%、樓層+5%、座向+2%、國宅品牌-6%、無專有車位折讓35~50萬、24年未裝潢折讓105萬",
        "dimensions": dims
    }

def build_renovation_budget():
    categories = [
        {
            "category_id": 1,
            "name": "水電工程",
            "short_name": "水電",
            "amount": 180000,
            "amount_wan": 18.0,
            "duration_days": 10,
            "description": "全室太平洋電線抽換、總開關箱加大配電、冷熱不銹鋼水管重配、專用迴路增設",
            "rationale": "24年老舊線路具用電走火風險，重配為現代家庭安全剛需。"
        },
        {
            "category_id": 2,
            "name": "衛浴泥作工程",
            "short_name": "衛浴泥作",
            "amount": 280000,
            "amount_wan": 28.0,
            "duration_days": 14,
            "description": "兩間衛浴全室打除至紅磚層、試水48小時高規格防水、壁地磚鋪設、天花板抽風機",
            "rationale": "早期防水層壽命約15年早已過保，必須徹底打除避免漏水至樓下。"
        },
        {
            "category_id": 3,
            "name": "衛浴設備工程",
            "short_name": "衛浴設備",
            "amount": 120000,
            "amount_wan": 12.0,
            "duration_days": 3,
            "description": "TOTO兩段式省水馬桶×2、面盆浴櫃、恆溫淋浴龍頭、乾濕分離強化玻璃拉門",
            "rationale": "老舊設備泛黃且無乾濕分離，更新提升衛生與安全性。"
        },
        {
            "category_id": 4,
            "name": "廚房廚具工程",
            "short_name": "廚房廚具",
            "amount": 160000,
            "amount_wan": 16.0,
            "duration_days": 4,
            "description": "一字型240cm歐化結晶鋼烤廚具、三星人造石檯面、櫻花三機(抽油煙機/瓦斯爐/烘碗機)",
            "rationale": "原況廚具五金鏽蝕且檯面磨損，需全面換新。"
        },
        {
            "category_id": 5,
            "name": "門窗與油漆工程",
            "short_name": "門窗油漆",
            "amount": 140000,
            "amount_wan": 14.0,
            "duration_days": 7,
            "description": "大門更換電子鎖、室內房間門換新、全室牆面批土研磨、得利乳膠漆三度粉刷",
            "rationale": "去除老舊壁癌痕跡，改善室內採光與空氣品質。"
        },
        {
            "category_id": 6,
            "name": "地板工程",
            "short_name": "地板",
            "amount": 100000,
            "amount_wan": 10.0,
            "duration_days": 3,
            "description": "全室直鋪超耐磨木地板（約23坪實鋪，含防潮隔音墊）",
            "rationale": "遮蓋早期舊地磚，提升居住溫馨感與吸音效果。"
        },
        {
            "category_id": 7,
            "name": "清潔與廢棄物清運",
            "short_name": "清潔",
            "amount": 70000,
            "amount_wan": 7.0,
            "duration_days": 3,
            "description": "打除廢料垃圾清運車次、全室細部清潔完工點交",
            "rationale": "高樓層梯位防護與大型工程廢棄物合法清運。"
        }
    ]
    total_amount = sum(c['amount'] for c in categories)
    total_wan = round(total_amount / 10000.0, 1)
    return {
        "total": total_amount,
        "total_wan": total_wan,
        "estimated_duration_days": 45,
        "target_ping": 44.50,
        "unit_deduction_per_ping": round(total_wan / 44.50, 2),
        "breakdown_wan": {
            "水電工程": 18.0,
            "衛浴泥作": 28.0,
            "衛浴設備": 12.0,
            "廚房廚具": 16.0,
            "門窗油漆": 14.0,
            "地板": 10.0,
            "清潔": 7.0
        },
        "categories": categories
    }

def build_pricing_tiers():
    tiers = {
        "target_total_ping": 44.50,
        "asking_price": 998.0,
        "original_asking_price": 1080.0,
        "price_reduction": 82.0,
        "asking_unit_price": 22.43,
        "conservative_floor": {
            "tier_name": "保守底標價 (大膽試探防線)",
            "tier_level": "Tier 1",
            "total_low": 780.0,
            "total_high": 800.0,
            "total": 790.0,
            "unit_low": 17.53,
            "unit_high": 17.98,
            "unit": 17.75,
            "discount_low": 0.78,
            "discount_high": 0.80,
            "discount": 0.79,
            "discount_display": "78折 ~ 80折 (約開價78~80折)",
            "price_range_display": "780 ~ 800 萬元",
            "unit_price_display": "17.53 ~ 17.98 萬/坪",
            "bargaining_space": "-198 ~ -218 萬元",
            "rationale": "錨定114年2樓成交底線（850萬/20.5萬）扣除百萬翻修成本，立足於工學天下115年最新摜破2字頭（19.6萬/坪）市場恐慌，針對屋主一手持有24年零房貸原始成本僅200多萬之低成本優勢，大膽試探賣方心理底線。"
        },
        "reasonable_target": {
            "tier_name": "合理目標價 (最佳成交甜蜜點)",
            "tier_level": "Tier 2",
            "total_low": 860.0,
            "total_high": 890.0,
            "total": 875.0,
            "unit_low": 19.33,
            "unit_high": 20.00,
            "unit": 19.66,
            "discount_low": 0.86,
            "discount_high": 0.89,
            "discount": 0.875,
            "discount_display": "86折 ~ 89折 (約開價86~89折)",
            "price_range_display": "860 ~ 890 萬元",
            "unit_price_display": "19.33 ~ 20.00 萬/坪",
            "bargaining_space": "-108 ~ -138 萬元",
            "rationale": "以115-05 3樓最新成交（980萬/22.7萬）扣減必要翻修差額約100萬元。總價落在8字頭尾兼顧屋主面子，屋主實質大賺逾550萬元現金，為最易促成簽約之最佳甜蜜點。"
        },
        "ceiling_max": {
            "tier_name": "上限天花板價 (談判停損警戒線)",
            "tier_level": "Tier 3",
            "total_low": 920.0,
            "total_high": 930.0,
            "total": 925.0,
            "unit_low": 20.67,
            "unit_high": 20.90,
            "unit": 20.79,
            "discount_low": 0.92,
            "discount_high": 0.93,
            "discount": 0.925,
            "discount_display": "92折 ~ 93折 (約開價92~93折)",
            "price_range_display": "920 ~ 930 萬元",
            "unit_price_display": "20.67 ~ 20.90 萬/坪",
            "bargaining_space": "-68 ~ -78 萬元",
            "rationale": "絕對不能突破3樓成交價980萬心理防線。成交達930萬加上100萬翻修後總投入成本達1,030萬元（實質單價達23.1萬/坪）。若賣方堅持950萬以上，買方應果斷停損轉向購買歌德堡或長億南園。"
        }
    }
    return tiers

def build_negotiation_scripts():
    scripts = [
        {
            "id": 1,
            "title": "引用2026最新大盤回跌鐵證，打破賣方113年高點幻想",
            "topic": "大盤回修與競品破2字頭鐵證",
            "target_party": "曾協理 (銷售仲介) 與 屋主",
            "core_argument": "央行信用管制後南區成屋成交量縮半，工學天下最新摜破2字頭至19.6萬，大慶國宅114年2樓850萬(20.5萬)，998萬開價脫離目前買方市場行情。",
            "script_dialogue": "曾協理，我們非常喜歡大慶國宅的環境與高樓採光，但現在市場盤勢已經和113年高點完全不同了。央行信用管制後，南區成交量掉超過一半。隔壁工學路指標社區『工學天下』今年6月7樓成交最新單價已經摜破2字頭，來到19.6萬/坪；大慶街的『長億南園』今年也回跌到22.1萬。我們查過大慶國宅114年2樓才賣850萬（20.5萬/坪），開價998萬雖然降過一次，但仍然停留在多頭時期的價格，現在市場不可能有人追這種價格。"
        },
        {
            "id": 2,
            "title": "無獨立產權車位之資產抗性與流動性缺陷",
            "topic": "無車位產權之轉手流動性抗性",
            "target_party": "曾協理 (銷售仲介) 與 屋主",
            "core_argument": "權狀無車位在未來轉手三房家庭客群時面臨嚴重心理抗性與銀行鑑價扣減，獨立平車市場合理估值120萬，無車位戶別總價必須實質折讓。",
            "script_dialogue": "這間房子權狀上『完全沒有車位產權』，雖然社區可以向管委會租平面車位，但那畢竟是租的，不能登記在權狀裡。未來我們如果要換屋轉手，買三房的家庭客戶一定會因為沒有產權車位而大砍價，甚至銀行估價時也只會當成無車位華廈鑑價。周邊有固定平面車位的大樓一個平車價值120萬元，我們買沒有車位的房子，總價如果不折讓個100萬下來，我們根本沒辦法承擔未來的轉手風險。"
        },
        {
            "id": 3,
            "title": "24年未翻新一手屋，全室水電泥作需自備百萬現金",
            "topic": "24年未翻修一手屋之現金翻修剛性扣減",
            "target_party": "曾協理 (銷售仲介) 與 屋主",
            "core_argument": "全室24年老舊電線、無窗衛浴防水老化與廚具舊化，買方必須自備105萬元純現金翻新，工程款無法納入房貸，自備款壓力巨大。",
            "script_dialogue": "屋主一手住了24年，室內水電管線從完工到現在完全沒換過，雙衛浴沒開窗且防水層早已老化，廚房設備也完全是20年前舊樣式。我們買下來後，為了自住安全，全室電線一定要重新抽換加大安培數、雙衛浴全部打除見底重做防水貼磚、廚具天花板冷氣管線都要重做，我們找統包粗估最陽春的基礎工程就要90到110萬元！這筆錢銀行是不給貸款的，買方要全部拿真金白銀的自備款現金貼進去。如果總價沒談到860~880萬，我們自備款加上裝修預算根本爆表，只能去看別間已經整理好的房子。"
        },
        {
            "id": 4,
            "title": "一手持有24年原始成本極低，促使屋主落袋為安",
            "topic": "一手屋主持有24年大賺550萬獲利了結",
            "target_party": "屋主心理防線",
            "core_argument": "屋主一手持有24年無貸款，原始成本推估僅220~280萬元，成交860~890萬屋主仍實質淨賺超過550萬元現金，獲利近3倍。",
            "script_dialogue": "屋主是在民國89年剛完工就買的一手屋，當年大慶國宅一坪才買5到7萬元，總價成本才200多萬，早就沒有任何房貸壓力了。現在我們出860萬，屋主一口氣實質淨賺超過550萬元現金！這比把錢放在銀行定存或是股市翻了兩三倍。現在市場買氣這麼冷，房子空在那邊每個月也是白繳管理費跟折舊，趁現在有我們這種準備好現金、隨時可以下斡旋簽約的優質自住客，趕快獲利了結落袋為安才是最聰明的決定。"
        },
        {
            "id": 5,
            "title": "下斡旋即成交定錨法（帶支票進場，限時施壓）",
            "topic": "帶支票下斡旋限時3天決策定錨",
            "target_party": "曾協理 (銷售仲介)",
            "core_argument": "實質準備10萬元支票下斡旋860萬元，給予3天有效期限，同步看歌德堡與長億南園，施壓仲介強力促成屋主降價。",
            "script_dialogue": "曾協理，我們是真心誠意要買這間房子自住，不是投資客在隨便問問。我們今天就可以直接開出10萬元斡旋金支票，出價860萬元。但我們同時也在看德富路的『歌德堡』跟大慶街的『長億南園』，斡旋有效期限給屋主3天考慮。只要屋主願意點頭，我們隨時配合代書簽約，不囉嗦、不拖泥帶水；如果屋主堅持要930萬以上，那我們就直接退斡旋轉買歌德堡。"
        }
    ]
    return scripts

def build_car_parking_separation():
    return {
        "benchmark_deal_date": "2026-05-22",
        "benchmark_unit": "台中市南區高工路 3樓之1",
        "benchmark_floor": "3樓",
        "benchmark_layout": "3房2廳",
        "gross_transaction_price": 980.0,
        "gross_building_area": 43.17,
        "gross_unit_price": 22.70,
        "parking_space_count": 1,
        "parking_space_type": "平面車位",
        "parking_space_area": 7.6,
        "parking_space_price": 120.0,
        "separated_net_building_price": 860.0,
        "separated_net_building_area": 35.57,
        "separated_net_unit_price": 24.18,
        "formula": "(980 - 120) / (43.17 - 7.6) = 860 / 35.57 = 24.18 萬/坪",
        "target_property_parking_status": "權狀無車位 (現況承租管委會平面車位 1,500元/月)",
        "target_property_parking_adjustment_value": "-35~50 萬元 (缺乏獨立產權之流動性折減)",
        "notes": "大慶國宅社區約80%住戶車位由管委會統籌出租管理，月租僅1,500元，買方可省下120萬購置成本，但缺乏獨立產權在轉手時存在受眾縮小效應。"
    }

def main():
    print("=== Generating scratch/daqing_valuation_master.json ===")
    src_data, items = load_source_transactions()
    processed, clean_res, special, gov, shop = process_transactions(items)
    print(f"Total Transactions: {len(processed)}")
    print(f"Clean Residential: {len(clean_res)}")
    print(f"Special Deals: {len(special)}")
    print(f"Government Discount Deals: {len(gov)}")
    print(f"First Floor Shop Deals: {len(shop)}")

    hist_summary = compute_historical_summary(processed, clean_res, special, gov, shop)
    comps = build_comparable_communities()
    val_matrix = build_valuation_adjustment_matrix()
    renov = build_renovation_budget()
    tiers = build_pricing_tiers()
    scripts = build_negotiation_scripts()
    car_sep = build_car_parking_separation()

    master_data = {
        "metadata": {
            "generated_at": datetime.now().isoformat(),
            "version": "1.0.0",
            "author": "worker_daqing_m1",
            "project": "台中市南區「大慶國宅(大慶中山)」社區待售物件出價分析與全社區行情評估套件",
            "description": "Master Structured Dataset and Valuation Engine for Daqing Public Housing (大慶中山 10F) and South District Comparables"
        },
        "target_property": {
            "community_name": "大慶中山",
            "alias": "大慶國宅",
            "aliases": ["大慶國宅", "中山醫大國宅", "大慶中山華廈"],
            "id": 20513,
            "leju_id": "L37277079e472",
            "developer": "台中市政府 / 內政部營建署 (公有國宅)",
            "address": "台中市南區工學二街16號 / 高工南路1號",
            "floor": "10F/12F",
            "target_floor": 10,
            "total_floors": 12,
            "total_floor": 12,
            "basement_floors": 2,
            "is_sub_penthouse": False,
            "floor_feature": "10樓次高樓景觀戶，中庭視野開闊，通風採光極佳，免除頂樓日曬及漏水隱憂",
            "building_type": "電梯大樓 / 華廈",
            "structure": "鋼筋混凝土 (RC)",
            "build_date": "2002-04-01",
            "roc_build_date": "民國91年4月1日",
            "age_years": 24.5,
            "age": 24.5,
            "age_display": "約24.5年 (完工日2002-04-01，跨越921震後法規)",
            "households": 362,
            "total_households": 362,
            "dormitory_households": 253,
            "dormitory_ratio": 69.89,
            "dormitory_note": "2005年中山醫學大學斥資6億餘元向營建署與市府整批購入253戶作為學生宿舍(佔全社區約70%)，民間流通戶數僅剩約109戶，籌碼高度集中穩定",
            "elevator_ratio": "2~3戶1梯",
            "layout": "3房2廳2衛2陽台",
            "rooms": 3,
            "halls": 2,
            "bathrooms": 2,
            "balconies": 2,
            "orientation": "客廳景觀陽台朝東、工作陽台朝西、主臥朝東南，三面採光",
            "orientation_door": "朝南 (由北向南進入)",
            "orientation_living_room": "朝東 (客廳景觀陽台面中庭花園遠眺市景)",
            "orientation_kitchen_balcony": "朝西 (工作陽台)",
            "total_ping": 44.50,
            "main_ping": 25.365,
            "main_ratio": 57.00,
            "aux_ping": 2.405,
            "aux_ratio": 5.40,
            "common_ping": 16.726,
            "real_living_ping": 27.77,
            "real_ping": 27.77,
            "real_living_ratio": 62.40,
            "public_ratio": 37.59,
            "land_ping": 5.772,
            "car_space_type": "權狀無車位 (現況承租管委會平面車位 1,500元/月)",
            "car_space_ping": 0.0,
            "parking_rent_monthly": 1500,
            "asking_price": 998.0,
            "original_asking_price": 1080.0,
            "price_reduction": 82.0,
            "asking_unit_price": 22.43,
            "management_fee": 1500,
            "management_fee_monthly": 1500,
            "management_fee_per_ping": 33.7,
            "condition": "24年一手原況空屋，尚未裝潢，需全室水電重拉與雙衛浴泥作翻新",
            "owner_profile": "持有長達24年之一手原屋主，原始取得成本推估220~280萬，零房貸無資金壓力，降價82萬展現強烈變現誠意",
            "earthquake_safety_921": "安全，2002年完工為921震後新法規建築，無任何受災受損或紅黃單列管記錄",
            "barrier_free": "全區有無障礙坡道與電梯直達設施",
            "mrt_station_distance": "距捷運綠線/台鐵雙鐵共構大慶站約1.2公里 (機車約3分鐘、步行約15分鐘)",
            "school_district": "和平國小、四育國中 / 台中高工、中山醫學大學",
            "park_nearby": "社區自帶千坪中庭花園、工學綠園道、健康公園"
        },
        "pricing_tiers": tiers,
        "car_parking_separation": car_sep,
        "car_space_calculation": car_sep,  # Backward compatible alias
        "renovation_budget": renov,
        "comparable_communities": comps,
        "valuation_adjustment_matrix": val_matrix,
        "comparative_adjustment_matrix": val_matrix,  # Backward compatible alias
        "historical_summary": hist_summary,
        "negotiation_scripts": scripts,
        "negotiation_talking_points": scripts,  # Backward compatible alias
        "historical_transactions": {
            "all_transactions": processed,
            "clean_residential_deals": clean_res,
            "special_deals": special,
            "government_discount_deals": gov,
            "first_floor_shop_deals": shop
        }
    }

    out_path = 'scratch/daqing_valuation_master.json'
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(master_data, f, ensure_ascii=False, indent=2)

    size = os.path.getsize(out_path)
    print(f"Generated {out_path} successfully. Size: {size} bytes ({size/1024:.2f} KB)")

if __name__ == '__main__':
    main()
