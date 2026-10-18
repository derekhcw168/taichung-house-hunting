# -*- coding: utf-8 -*-
"""
更新《長億南園8樓_實價登錄趨勢與出價策略分析報告》Markdown 與 Word (.docx)
納入南區 10 大指標老牌大社區實價登錄反轉走勢圖與深度對比分析
"""

import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

if sys.stdout:
    sys.stdout.reconfigure(encoding='utf-8')

from doc_helpers import set_cell_background, set_cell_margins, add_callout, set_font

base_dir = r"d:\GoogleDrive\01_個人_家庭與生活\02_家庭_安莉的資料\小寶的買房子資料"
target_dir = os.path.join(base_dir, "長億南園8樓")
trend_dir = os.path.join(base_dir, "南區25年以上大型社區走勢圖")

md_path = os.path.join(target_dir, "長億南園8樓_實價登錄趨勢與出價策略分析報告.md")
docx_path = os.path.join(target_dir, "長億南園8樓_實價登錄趨勢與出價策略分析報告.docx")

bench_data = [
    {
        'rank': 1,
        'label': '長億城攬翠區',
        'name': '長億城攬翠區陽光區',
        'households': 669,
        'age': 30.9,
        'build_date': '1995-11-30',
        'address': '工學路116號',
        'peak_p': 29.5,
        'peak_date': '113年06月 (8樓)',
        'latest_p': 26.9,
        'latest_date': '115年06月 (10樓)',
        'avg113': 26.3,
        'avg115': 23.2,
        'drop_pct': -11.9,
        'vol_drop': '年成交 63 件 ➜ 萎縮至 9 件（量縮 85.7%）',
        'supp': '13.8～16.1 萬/坪',
        'img': '03_長億城攬翠區陽光區_成交行情走勢圖.png',
        'desc': '工學商圈極具指標性之超級大社區。實登走勢在 113 年衝上 29.5 萬歷史頂峰後反轉回檔，115 年最新均價已跌至 23.2 萬，跌幅達 11.9%。年成交量更由 63 件暴跌至個位數，社區內二手委售競爭劇烈，價格持續下修。'
    },
    {
        'rank': 2,
        'label': '工學天下',
        'name': '工學天下',
        'households': 452,
        'age': 31.7,
        'build_date': '1995-01-23',
        'address': '工學北路135號',
        'peak_p': 29.0,
        'peak_date': '113年07月 (9樓)',
        'latest_p': 19.6,
        'latest_date': '115年06月 (7樓)',
        'avg113': 24.7,
        'avg115': 20.7,
        'drop_pct': -16.1,
        'vol_drop': '年成交 25 件 ➜ 萎縮至 4 件（量縮 84%）',
        'supp': '13.6～16.5 萬/坪',
        'img': '07_工學天下_成交行情走勢圖.png',
        'desc': '屋齡 31.7 年，與長億南園條件極為相近。113 年高點曾達 29.0 萬，但 115 年 6 月最新成交直接暴跌至 19.6 萬/坪，正式摜破 2 字頭大關！均價跌幅達 16.1%，鐵證證明 30 年以上無障礙受限老屋在買盤眼中支撐極度脆弱。'
    },
    {
        'rank': 3,
        'label': '龍邦綠園道',
        'name': '龍邦綠園道',
        'households': 322,
        'age': 33.4,
        'build_date': '1993-05-19',
        'address': '南和路197號',
        'peak_p': 28.5,
        'peak_date': '114年02月 (5樓)',
        'latest_p': 23.6,
        'latest_date': '115年08月 (16樓)',
        'avg113': 23.6,
        'avg115': 21.3,
        'drop_pct': -9.5,
        'vol_drop': '年成交 14 件 ➜ 萎縮至 2 件（量縮 85.7%）',
        'supp': '13.8～16.8 萬/坪',
        'img': '15_龍邦綠園道_成交行情走勢圖.png',
        'desc': '南和路綠園道首排景觀名宅，具備極佳居住環境與指名度。但即使是綠園道景觀宅，114 年 28.5 萬高檔過後買盤急速縮手，115 年最新成交已下修至 23.6 萬，大盤下行期景觀溢價亦全面收斂回吐。'
    },
    {
        'rank': 4,
        'label': '文心高第',
        'name': '文心高第',
        'households': 266,
        'age': 24.9,
        'build_date': '2001-11-20',
        'address': '東興路一段29號',
        'peak_p': 34.5,
        'peak_date': '113年05月 (12樓)',
        'latest_p': 30.4,
        'latest_date': '115年07月 (2樓)',
        'avg113': 31.3,
        'avg115': 28.5,
        'drop_pct': -8.9,
        'vol_drop': '年成交 18 件 ➜ 萎縮至 3 件（量縮 83.3%）',
        'supp': '16.5～20.0 萬/坪',
        'img': '26_文心高第_成交行情走勢圖.png',
        'desc': '東興路名廈，屋齡 24.9 年（剛跨過 25 年門檻）。113 年衝破 34 萬天花板後，115 年最新均價已回修至 28.5 萬 (-8.9%)，委售週期大幅拉長，買方對 3 字頭老屋抗拒感強烈。'
    },
    {
        'rank': 5,
        'label': '大東家華園',
        'name': '大東家華園',
        'households': 227,
        'age': 31.2,
        'build_date': '1995-07-28',
        'address': '復興路二段30巷15號',
        'peak_p': 27.7,
        'peak_date': '113年03月 (4樓)',
        'latest_p': 20.3,
        'latest_date': '115年07月 (3樓)',
        'avg113': 24.5,
        'avg115': 23.1,
        'drop_pct': -5.9,
        'vol_drop': '年成交 16 件 ➜ 萎縮至 3 件（量縮 81.3%）',
        'supp': '13.5～16.2 萬/坪',
        'img': '35_大東家華園_成交行情走勢圖.png',
        'desc': '近學府路商圈知名大型社區。實登最新 115 年 7 月成交單價直接殺至 20.3 萬/坪，即將失守 2 字頭！充分反映傳統市區老牌大樓買盤追價力道枯竭，價格快速朝歷史常態修正。'
    },
    {
        'rank': 6,
        'label': '假日半島',
        'name': '假日半島',
        'households': 219,
        'age': 32.4,
        'build_date': '1994-05-18',
        'address': '復興路二段71巷21弄31號',
        'peak_p': 33.2,
        'peak_date': '113年10月 (9樓)',
        'latest_p': 24.8,
        'latest_date': '115年07月 (5樓)',
        'avg113': 29.4,
        'avg115': 25.3,
        'drop_pct': -14.0,
        'vol_drop': '年成交 15 件 ➜ 萎縮至 2 件（量縮 86.7%）',
        'supp': '14.0～17.0 萬/坪',
        'img': '38_假日半島_成交行情走勢圖.png',
        'desc': '復興路二段知名指名社區。113 年底曾創 33.2 萬極端單價，但 115 年均價已急速下修至 25.3 萬，跌幅達 14.0%。最新成交 24.8 萬印證追高風險極大，老牌社區缺乏新題材支撐，溢價迅速消退。'
    },
    {
        'rank': 7,
        'label': '文心大地',
        'name': '文心大地',
        'households': 207,
        'age': 27.4,
        'build_date': '1999-05-18',
        'address': '文心南十路380號',
        'peak_p': 28.4,
        'peak_date': '113年07月 (12樓)',
        'latest_p': 24.1,
        'latest_date': '115年05月 (10樓)',
        'avg113': 26.4,
        'avg115': 22.5,
        'drop_pct': -15.1,
        'vol_drop': '年成交 14 件 ➜ 萎縮至 3 件（量縮 78.6%）',
        'supp': '13.5～16.5 萬/坪',
        'img': '41_文心大地_成交行情走勢圖.png',
        'desc': '文心南路商圈 27 年指標大案。113 年均價 26.4 萬，115 年均價已滑落至 22.5 萬，跌幅高達 15.1%！買氣受央行信用管制嚴重壓抑，自住客出價趨於高度保守。'
    },
    {
        'rank': 8,
        'label': '文心及第',
        'name': '文心及第(文心南路)',
        'households': 259,
        'age': 27.3,
        'build_date': '1999-06-12',
        'address': '文心南路846號',
        'peak_p': 35.7,
        'peak_date': '113年06月 (10樓)',
        'latest_p': 18.9,
        'latest_date': '115年03月 (14樓)',
        'avg113': 30.6,
        'avg115': 18.9,
        'drop_pct': -38.2,
        'vol_drop': '年成交 19 件 ➜ 萎縮至 2 件（量縮 89.5%）',
        'supp': '14.0～17.5 萬/坪',
        'img': '27_文心及第文心南路_成交行情走勢圖.png',
        'desc': '🔥 震驚南區市場之最震撼案例！113 年 6 月曾以 35.7 萬創歷史天價，但 115 年 3 月 14 樓頂樓次高戶直接暴跌至 18.9 萬/坪成交，摜破 2 字頭、跌幅高達 38.2%！鐵證證明老舊大樓一旦買方縮手，拋售潮將引發無情破底。'
    },
    {
        'rank': 9,
        'label': '大安黃金城大廈',
        'name': '大安黃金城大廈',
        'households': 307,
        'age': 32.2,
        'build_date': '1994-07-27',
        'address': '忠孝路46號',
        'peak_p': 34.3,
        'peak_date': '114年04月 (6樓)',
        'latest_p': 25.2,
        'latest_date': '115年06月 (2樓)',
        'avg113': 26.6,
        'avg115': 23.7,
        'drop_pct': -11.1,
        'vol_drop': '年成交 22 件 ➜ 萎縮至 4 件（量縮 81.8%）',
        'supp': '13.0～15.8 萬/坪',
        'img': '18_大安黃金城大廈_成交行情走勢圖.png',
        'desc': '興大生活圈指標性大社區。114 年短暫衝上 34.3 萬後迅速熄火，115 年成交已重挫回 25.2 萬，均價下修 11.1%。老社區維護成本增加與貸款緊縮使買方出價極度冷靜。'
    },
    {
        'rank': 10,
        'label': '京城特區',
        'name': '京城特區',
        'households': 388,
        'age': 29.1,
        'build_date': '1997-09-02',
        'address': '學府路168巷3號',
        'peak_p': 25.9,
        'peak_date': '113年03月 (18樓)',
        'latest_p': 18.6,
        'latest_date': '115年05月 (3樓)',
        'avg113': 22.4,
        'avg115': 20.9,
        'drop_pct': -6.9,
        'vol_drop': '年成交 21 件 ➜ 萎縮至 3 件（量縮 85.7%）',
        'supp': '13.2～16.0 萬/坪',
        'img': '11_京城特區_成交行情走勢圖.png',
        'desc': '🔥 388 戶中興大學學府路地標社區。113 年高點 25.9 萬，115 年 5 月最新成交直接跌出 18.6 萬/坪之 1 字頭破盤價！均價跌至 20.9 萬。量能崩跌 85.7%，證明戶數多之老社區正面臨沉重下修賣壓。'
    }
]

print("1. 更新 Markdown 報告文件...")
with open(md_path, 'r', encoding='utf-8') as f:
    md_content = f.read()

# Check if Chapter 7 is already present
if "## 柒、 全南區 10 大指標老牌大社區實價登錄全面反轉向下橫向對比分析" not in md_content:
    md_append = "\n\n---\n\n## 柒、 全南區 10 大指標老牌大社區實價登錄全面反轉向下橫向對比分析（客觀鐵證）\n\n"
    md_append += "> **分析依據：** 591 實價登錄、樂居 Leju、僑馥 CTOPMap 全面統計台中市南區 25 年以上、戶數超過 150 戶之 53 棟老牌集合住宅社區。\n"
    md_append += "> **大盤結論：** 央行第七波信用管制與高利率環境下，全南區老牌大社區在 113 年創歷史頂峰後**「全數見頂反轉向下」**，年成交量崩縮 76.4%！其中**工學天下（19.6萬）**、**文心及第（18.9萬）**、**京城特區（18.6萬）**最新成交更已全數**跌破 2 字頭**！這證明長億南園近期從 26 萬回跌至 20~21 萬絕非個案，而是整個大盤循環向下修正的必然結果，為買方「打八折」出價提供了堅不可摧的客觀鐵證。\n\n"
    
    md_append += "### 📊 南區 10 大指標老牌社區實價登錄量價反轉對照總表\n\n"
    md_append += "| 編號 | 指標社區建案 | 規模 / 屋齡 | 歷史最高單價 (簽約年月) | 115 年最新成交 (簽約年月) | 113➜115均價走勢 | 均價跌幅 | 實價登錄反轉特徵與震撼意義 |\n"
    md_append += "| :---: | :--- | :---: | :---: | :---: | :---: | :---: | :--- |\n"
    for b in bench_data:
        md_append += f"| {b['rank']} | **{b['label']}** | {b['households']} 戶<br>約 {b['age']} 年 | **{b['peak_p']:.1f} 萬**<br>({b['peak_date']}) | **{b['latest_p']:.1f} 萬**<br>({b['latest_date']}) | {b['avg113']:.1f} ➜ **{b['avg115']:.1f} 萬** | **`{b['drop_pct']:.1f}%`** | {b['desc'].split('。')[0]}。 |\n"
    
    md_append += "\n---\n\n### 🏢 10 大指標社區獨立成交行情走勢圖與深度剖析\n\n"
    for b in bench_data:
        md_append += f"#### {b['rank']}. 【{b['label']}】（{b['households']}戶・屋齡約{b['age']}年）—— 最新成交下修至 **{b['latest_p']:.1f} 萬/坪**（均價跌幅 **`{b['drop_pct']:.1f}%`**）\n\n"
        md_append += f"![{b['name']} 成交行情走勢圖](../南區25年以上大型社區走勢圖/{b['img']})\n\n"
        md_append += f"* **歷史高點與回檔實證：** {b['name']} 實登歷史高點出現在 {b['peak_date']} 單價達 **{b['peak_p']:.1f} 萬/坪**；最新成交已於 {b['latest_date']} 下挫至 **{b['latest_p']:.1f} 萬/坪**。純住宅年度均價由 113 年的 {b['avg113']:.1f} 萬下修至 115 年的 **{b['avg115']:.1f} 萬**，跌幅高達 **`{b['drop_pct']:.1f}%`**！\n"
        md_append += f"* **成交量能斷崖式萎縮：** {b['vol_drop']}。買氣嚴重急凍，二手委售週期大幅拉長。\n"
        md_append += f"* **長期合理價值支撐：** 大盤模型測算該社區合理支撐帶為 **{b['supp']}**，預估 116～118 年受 30 年以上老屋限貸影響，單價將持續回測支撐。\n"
        md_append += f"* **議價說服震撼關鍵：** {b['desc']}\n\n"

    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(md_content + md_append)
    print("✅ Markdown 報告已成功更新 Chapter 柒！")
else:
    print("ℹ️ Markdown 報告已包含 Chapter 柒，略過。")

print("2. 重新編譯 Word (.docx) 報告文件...")
# We will regenerate or update docx_path
doc = docx.Document(docx_path)

# Check if Chapter 7 already in docx
has_chap7 = any("柒、 全南區 10 大指標老牌大社區" in p.text for p in doc.paragraphs)
if not has_chap7:
    # Append Chapter 7
    doc.add_page_break()
    h7 = doc.add_heading(level=1)
    r_h7 = h7.add_run("柒、 全南區 10 大指標老牌大社區實價登錄全面反轉向下橫向對比分析（客觀鐵證）")
    set_font(r_h7, size_pt=16, bold=True, color_rgb=(0x1B, 0x36, 0x5D))

    add_callout(doc,
        "【大盤總體情勢研判】\n"
        "依據 591 實價登錄、樂居 Leju 及僑馥 CTOPMap 最新資料庫，全台中市南區 25 年以上、戶數逾 150 戶之 53 棟老牌集合住宅社區，在經歷 113 年歷史高點後，受央行第七波信用管制與高利率衝擊，全線呈現「量縮 76.4%、價格全面見頂回跌」的系統性下行循環！\n\n"
        "特別是工學商圈、文心南路商圈及學府商圈等 10 大超級指標老牌大社區，最新成交已全數自高檔重挫，其中工學天下（19.6萬）、京城特區（18.6萬）、文心及第（18.9萬）更已直接摜破 2 字頭大關！長億南園最新成交自 26 萬回跌至 20～21 萬絕非個案，而是大盤下修之必然趨勢，為買方堅定出價（打 8 折）提供堅實客觀鐵證！",
        title="南區老牌大社區實價登錄全盤反轉・核心總結",
        border_color="C0392B",
        bg_color="FDEDEC"
    )

    # Table
    p_tl = doc.add_paragraph()
    r_tl = p_tl.add_run("📋 表：台中市南區 10 大指標老牌集合住宅社區實價登錄量價走勢反轉對照總表")
    set_font(r_tl, size_pt=11, bold=True, color_rgb=(0x1B, 0x36, 0x5D))

    tbl_doc = doc.add_table(rows=len(bench_data) + 1, cols=7)
    tbl_doc.alignment = WD_TABLE_ALIGNMENT.CENTER

    headers_bench_doc = ["指標社區名稱", "規模 / 屋齡", "歷史最高單價", "115年最新成交", "均價走勢", "均價跌幅", "實登反轉特徵解析"]
    col_widths = [Inches(1.2), Inches(0.9), Inches(1.0), Inches(1.0), Inches(1.0), Inches(0.7), Inches(1.4)]

    for c_idx, h_text in enumerate(headers_bench_doc):
        cell = tbl_doc.cell(0, c_idx)
        cell.width = col_widths[c_idx]
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h_text)
        set_font(r, size_pt=8.5, bold=True, color_rgb=(0xFF, 0xFF, 0xFF))

    for r_idx, b in enumerate(bench_data, start=1):
        row_vals = [
            b['label'],
            f"{b['households']}戶\n約{b['age']}年",
            f"{b['peak_p']:.1f}萬\n({b['peak_date']})",
            f"{b['latest_p']:.1f}萬\n({b['latest_date']})",
            f"{b['avg113']:.1f}➜{b['avg115']:.1f}萬",
            f"{b['drop_pct']:.1f}%",
            b['desc'].split('。')[0]
        ]
        bg_hex = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_vals):
            cell = tbl_doc.cell(r_idx, c_idx)
            cell.width = col_widths[c_idx]
            set_cell_background(cell, bg_hex)
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            is_bold = c_idx in [0, 3, 5]
            col_color = (0xC0, 0x39, 0x2B) if c_idx == 5 else (0x2D, 0x37, 0x48)
            set_font(r, size_pt=8.0, bold=is_bold, color_rgb=col_color)

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(10)

    # 逐棟插入高解析度走勢圖與深度解說
    for idx, b in enumerate(bench_data, start=1):
        h_item = doc.add_heading(level=2)
        r_hi = h_item.add_run(f"指標社區 {idx}：【{b['label']}】（{b['households']}戶・屋齡約{b['age']}年）—— 最新成交下修至 {b['latest_p']:.1f} 萬/坪")
        set_font(r_hi, size_pt=12, bold=True, color_rgb=(0x1B, 0x36, 0x5D))

        img_path = os.path.join(trend_dir, b['img'])
        if os.path.exists(img_path):
            p_cap = doc.add_paragraph()
            r_c = p_cap.add_run(f"📊 圖：{b['name']} 實價登錄歷史價量走勢與後續三年預估圖")
            set_font(r_c, size_pt=9.5, bold=True, color_rgb=(0x64, 0x74, 0x8B))
            doc.add_picture(img_path, width=Inches(6.2))
            p_is = doc.add_paragraph()
            p_is.paragraph_format.space_after = Pt(4)

        analysis_text = (
            f"• 歷史高點與回檔實證：{b['name']} 實登歷史高點出現在 {b['peak_date']} 單價達 {b['peak_p']:.1f} 萬/坪；"
            f"最新成交已於 {b['latest_date']} 下挫至 {b['latest_p']:.1f} 萬/坪。純住宅年度均價由 113 年的 {b['avg113']:.1f} 萬下修至 115 年的 {b['avg115']:.1f} 萬，跌幅高達 {b['drop_pct']:.1f}%！\n"
            f"• 成交量能斷崖式萎縮：{b['vol_drop']}。買氣嚴重急凍，二手委售週期大幅拉長。\n"
            f"• 長期合理價值支撐：大盤模型測算該社區合理支撐帶為 {b['supp']}，預估 116～118 年受 30 年以上老屋限貸影響，單價將持續回測支撐。\n"
            f"• 議價說服震撼關鍵：{b['desc']}"
        )
        add_callout(doc, analysis_text, title=f"{b['label']} 實價登錄反轉實證解析", border_color="2980B9", bg_color="F8FAFC")

    doc.save(docx_path)
    print("✅ Word 報告已成功更新 Chapter 柒並嵌入 10 大社區走勢圖！")
else:
    print("ℹ️ Word 報告已包含 Chapter 柒，略過。")
