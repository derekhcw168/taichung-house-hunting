import os
import sys

if sys.stdout:
    sys.stdout.reconfigure(encoding='utf-8')

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    for shd in tcPr.findall(qn('w:shd')):
        tcPr.remove(shd)
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    for mar in tcPr.findall(qn('w:tcMar')):
        tcPr.remove(mar)
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def set_cell_borders(cell, top="E2E8F0", bottom="E2E8F0", left="none", right="none", sz="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    for b in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(b)
    
    top_attr = f'w:val="single" w:sz="{sz}" w:space="0" w:color="{top}"' if top != "none" else 'w:val="none"'
    bot_attr = f'w:val="single" w:sz="{sz}" w:space="0" w:color="{bottom}"' if bottom != "none" else 'w:val="none"'
    l_attr = f'w:val="single" w:sz="{sz}" w:space="0" w:color="{left}"' if left != "none" else 'w:val="none"'
    r_attr = f'w:val="single" w:sz="{sz}" w:space="0" w:color="{right}"' if right != "none" else 'w:val="none"'
    
    tcBorders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top {top_attr}/>
            <w:left {l_attr}/>
            <w:bottom {bot_attr}/>
            <w:right {r_attr}/>
        </w:tcBorders>
    ''')
    tcPr.append(tcBorders)

def set_cell_left_border_only(cell, color_hex="1E3A8A", sz="24"):
    tcPr = cell._tc.get_or_add_tcPr()
    for b in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(b)
    tcBorders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="none"/>
            <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color_hex}"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(tcBorders)

def format_run(run, font_name="微軟正黑體", size_pt=9.5, bold=False, italic=False, color_rgb=(45, 55, 72)):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)

def add_header_footer(doc):
    section = doc.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hr = hp.add_run("大墩人文 11 樓｜老屋翻新基礎工程預算估價報價單")
    format_run(hr, size_pt=8.5, color_rgb=(148, 163, 184))

    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    fr1 = fp.add_run("本報價單為純自住安全基礎工程估算・未含奢華裝潢與家電設備")
    format_run(fr1, size_pt=8.5, color_rgb=(148, 163, 184))
    
    fp2 = footer.add_paragraph()
    fp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    fr2 = fp2.add_run("第 ")
    format_run(fr2, size_pt=8.5, color_rgb=(148, 163, 184))
    fldSimple1 = OxmlElement('w:fldSimple')
    fldSimple1.set(qn('w:instr'), 'PAGE')
    fp2._p.append(fldSimple1)
    fr3 = fp2.add_run(" 頁")
    format_run(fr3, size_pt=8.5, color_rgb=(148, 163, 184))

def main():
    doc = docx.Document()
    add_header_footer(doc)

    # 1. Badge
    badge_p = doc.add_paragraph()
    badge_p.paragraph_format.space_before = Pt(0)
    badge_p.paragraph_format.space_after = Pt(3)
    r_badge = badge_p.add_run("【工程預算明細書】自住結構安全與管線翻新")
    format_run(r_badge, size_pt=9.5, bold=True, color_rgb=(37, 99, 235))

    # 2. Title
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    r_title = title_p.add_run("【大墩人文 11 樓】34 年中古屋老屋翻新・基礎工程預算估價報價單")
    format_run(r_title, size_pt=18, bold=True, color_rgb=(30, 58, 138))

    # 3. Project Meta Table
    meta_table = doc.add_table(rows=3, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False

    meta_items = [
        ("工程標的", "台中市南屯區大墩人文社區 11 樓"),
        ("室內坪數", "登記建坪約 31.54 坪（室內約 26~27 坪，三房兩廳兩衛兩陽台）"),
        ("屋齡現況", "完工約民國 82 年（屋齡近 34 年，未全面翻新原始屋況）"),
        ("工程性質", "老屋安全必做「純基礎工程」（不含奢華木作、活動家具與家電）"),
        ("預估工期", "約 60 ~ 75 工作天（含拆除、水電、泥作、防水養生及試水）"),
        ("估價總額", "新台幣 $1,205,000 元整（約 120 萬元，含清運未稅）")
    ]

    for idx, (label, val) in enumerate(meta_items):
        r_idx = idx // 2
        c_idx = idx % 2
        cell = meta_table.cell(r_idx, c_idx)
        cell.width = Inches(3.4)
        set_cell_background(cell, "F8FAFC")
        set_cell_borders(cell, top="E2E8F0", bottom="E2E8F0", left="E2E8F0", right="E2E8F0", sz="4")
        set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r1 = p.add_run(f"● {label}：")
        format_run(r1, size_pt=9, bold=True, color_rgb=(71, 85, 105))
        r2 = p.add_run(val)
        format_run(r2, size_pt=9, bold=(label == "估價總額"), color_rgb=(220, 38, 38) if label == "估價總額" else (15, 23, 42))

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 4. Summary Table Header
    sh_p = doc.add_paragraph()
    sh_p.paragraph_format.space_before = Pt(8)
    sh_p.paragraph_format.space_after = Pt(6)
    r_sh = sh_p.add_run("📊 基礎工程各工種預算彙整總表（合計 120.5 萬元）")
    format_run(r_sh, size_pt=13, bold=True, color_rgb=(30, 58, 138))

    # Summary Table
    summary_data = [
        ("一", "保護與拆除清運工程", "公共梯廳保護、舊木作櫃/天花板拆除、衛浴廚具打到底、廢料清運", "$162,000", "13.4%"),
        ("二", "水電管線與電力系統", "全室電線抽換、配電箱加大(75A~100A)、不鏽鋼冷熱水管、專用迴路", "$219,000", "18.2%"),
        ("三", "泥作修補、壁癌與防水", "雙衛浴泥作打底、彈泥防水三度(試水48hr)、雙衛浴地壁磚、壁癌根治", "$220,000", "18.3%"),
        ("四", "鋁門窗與氣密隔音工程", "客廳落地氣密窗、三房 5+5 膠合隔音氣密窗、三合一門、塞水路防漏", "$130,000", "10.8%"),
        ("五", "雙衛浴設備與五金配備", "雙衛浴省水單體馬桶、發泡板浴櫃面盆組、主衛乾濕分離、暖風機", "$106,000", "8.8%"),
        ("六", "廚房系統廚具與三機", "240cm 一字型系統廚具、石英石檯面、櫻花/林內三機(抽油煙機/瓦斯爐/烘碗機)", "$137,000", "11.4%"),
        ("七", "全室油漆粉刷工程", "全室裂縫平整、AB膠填縫抗裂、全批土兩道、細磨、得利/虹牌乳膠漆", "$93,000", "7.7%"),
        ("八", "門扇更換、修飾與地坪", "三間臥房實木門、雙衛浴防潮門、水電天花板局部包管、SPC石塑地板", "$138,000", "11.5%"),
        ("合計", "全室 8 大安全基礎工項", "全室老屋安全翻新總工程款（未稅）", "$1,205,000", "100.0%")
    ]

    t_sum = doc.add_table(rows=len(summary_data) + 1, cols=5)
    t_sum.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_sum.autofit = False

    col_widths = [Inches(0.6), Inches(1.8), Inches(2.7), Inches(1.0), Inches(0.7)]
    headers = ["項次", "工程類別", "主要施作內容摘要", "金額 (NTD)", "佔比"]

    # Header Row
    for c_idx, h_text in enumerate(headers):
        cell = t_sum.cell(0, c_idx)
        cell.width = col_widths[c_idx]
        set_cell_background(cell, "1E3A8A")
        set_cell_borders(cell, top="1E3A8A", bottom="1E3A8A", left="none", right="none", sz="6")
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        if c_idx in [0, 4]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif c_idx == 3:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(h_text)
        format_run(r, size_pt=9, bold=True, color_rgb=(255, 255, 255))

    for r_idx, row in enumerate(summary_data, 1):
        is_total = (r_idx == len(summary_data))
        bg_color = "FEF2F2" if is_total else ("F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
        
        for c_idx, val in enumerate(row):
            cell = t_sum.cell(r_idx, c_idx)
            cell.width = col_widths[c_idx]
            set_cell_background(cell, bg_color)
            set_cell_borders(cell, top="CBD5E1", bottom="1E3A8A" if is_total else "E2E8F0", left="none", right="none", sz="8" if is_total else "4")
            set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
            
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx in [0, 4]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif c_idx == 3:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            r = p.add_run(val)
            text_color = (220, 38, 38) if is_total else ((15, 23, 42) if c_idx == 1 else (51, 65, 85))
            format_run(r, size_pt=9, bold=is_total or (c_idx in [0, 1, 3]), color_rgb=text_color)

    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # 5. Detail Section
    dh_p = doc.add_paragraph()
    dh_p.paragraph_format.space_before = Pt(14)
    dh_p.paragraph_format.space_after = Pt(6)
    r_dh = dh_p.add_run("🛠️ 各工種詳細施工項目與單價分析表")
    format_run(r_dh, size_pt=13, bold=True, color_rgb=(30, 58, 138))

    detailed_categories = [
        {
            "cat_title": "一、 公共與室內保護及拆除清運工程（小計：$162,000）",
            "desc": "34 年老屋充斥大量舊固定木作、厚重櫥櫃及氧化天花板，必須拆除見底才能檢測水電暗管；目前環保清運費與工資大漲，清運佔比顯著。",
            "items": [
                ("1-1", "公共動線防護工程", "1", "式", "12,000", "12,000", "電梯車廂、11樓梯廳地面及轉角鋪設防潮布與防撞夾板"),
                ("1-2", "舊有固定木作櫃體拆除", "1", "式", "28,000", "28,000", "拆除客廳大型固定書櫃、走道展示櫃、臥房老舊固定衣櫃/鞋櫃"),
                ("1-3", "舊有木作天花板拆除", "1", "式", "18,000", "18,000", "全室氧化發黑舊天花板、角料全數拆除見水泥頂板"),
                ("1-4", "舊雙衛浴地壁磚打除", "2", "間", "13,000", "26,000", "兩間衛浴之磁磚、水泥砂漿層打鑿見 RC 紅磚底，浴缸打除"),
                ("1-5", "舊廚房廚具及磁磚打除", "1", "式", "10,000", "10,000", "拆除舊系統櫃、瓦斯台、流理台及烹飪區舊磁磚"),
                ("1-6", "全室老舊鋁門窗框拆除", "1", "式", "12,000", "12,000", "拆除客廳落地門窗、三房窗框及陽台舊欄杆紗窗"),
                ("1-7", "廢棄物室內分類搬運", "1", "式", "20,000", "20,000", "廢料裝袋、利用電梯人工搬運至 1 樓集中區"),
                ("1-8", "環保合法清運車趟", "4", "車", "9,000", "36,000", "3.5 噸卡車進場清運，含合法環保棄土場處理費（約 4 車次）")
            ]
        },
        {
            "cat_title": "二、 水電管線更新與電力系統工程（小計：$219,000）",
            "desc": "34 年老屋原配電容量僅 30A~40A，電線氧化有走火危險；暗管冷熱水管多為生鏽鐵管，出水泛黃且隨時爆管，為居住安全第一優先。",
            "items": [
                ("2-1", "全室電線抽換重拉", "31", "坪", "1,774", "55,000", "採用太平洋一級 CNS 認證 2.0mm/5.5mm 銅線，抽換暗管線路"),
                ("2-2", "總開關箱加大匯流排", "1", "式", "22,000", "22,000", "更換新型匯流排配電箱、士林電機無熔絲開關，容量升至 75A~100A"),
                ("2-3", "專用迴路獨立配置", "7", "組", "3,000", "21,000", "冷氣 4 組、廚房高耗電 2 組（微波/烤箱/IH）、暖風機 1 組"),
                ("2-4", "全室冷熱水管更換", "1", "式", "48,000", "48,000", "改用 SUS304 不鏽鋼雙壓接保溫管，徹底廢棄 34 年生鏽鐵管"),
                ("2-5", "排水與排污幹管整修", "1", "式", "25,000", "25,000", "雙衛浴、廚房及後陽台地排更新、防臭存水彎接管重配"),
                ("2-6", "開關插座面板更換", "1", "式", "18,000", "18,000", "全室換裝 Panasonic COSMO 系列大面板開關、三孔接地插座"),
                ("2-7", "水電暗管敲鑿泥作修補", "1", "式", "22,000", "22,000", "新管路開槽打鑿、配管後水泥砂漿平整回填修復"),
                ("2-8", "天然瓦斯管檢驗配管", "1", "式", "8,000", "8,000", "瓦斯錶後明管安全檢整、防漏氣閥門更新")
            ]
        },
        {
            "cat_title": "三、 泥作修補、壁癌整治與雙衛浴防水工程（小計：$220,000）",
            "desc": "34 年集合住宅普遍有壁癌；兩間衛浴舊防水層失去彈性，必須重新施作彈泥防水三道並試水 48 小時，徹底杜絕漏水糾紛。",
            "items": [
                ("3-1", "壁癌深層根治工程", "8", "坪", "4,000", "32,000", "剔除粉化鬆動水泥層至紅磚、抗鹼底漆塗佈、抗裂防水砂漿修補"),
                ("3-2", "拆除後牆地面打底", "1", "式", "25,000", "25,000", "打除舊木作後殘膠凹洞補平、門檻門框泥作填縫"),
                ("3-3", "雙衛浴水泥打底粉光", "2", "間", "15,000", "30,000", "水泥砂漿重新整平粉光，地面精密抓 1:100 洩水坡度"),
                ("3-4", "雙衛浴高規格防水工程", "2", "間", "19,000", "38,000", "彈性水泥塗刷三道至樑下、角隅抗裂不織布、蓄水試水 48hr"),
                ("3-5", "雙衛浴壁磚鋪貼工料", "2", "間", "26,000", "52,000", "鋪貼 30x60cm 一級防潮修邊壁磚，含專用黏著劑與防霉填縫"),
                ("3-6", "雙衛浴地磚鋪貼工料", "2", "間", "14,000", "28,000", "鋪貼 30x30cm 一級防滑低吸水地磚，含人造石門檻安裝"),
                ("3-7", "廚房牆面打底防濺基底", "1", "式", "15,000", "15,000", "廚具背牆整平粉光，以利後續烤漆玻璃或壁磚附著")
            ]
        },
        {
            "cat_title": "四、 鋁門窗與氣密隔音防颱工程（小計：$130,000）",
            "desc": "11 樓為高樓層迎風面，34 年老舊薄鋁窗變形漏風下雨滲水。全室換裝膠合氣密窗為保障風雨安全與安寧之必做項目。",
            "items": [
                ("4-1", "客廳大落地氣密隔音門", "1", "式", "38,000", "38,000", "正新/大同同級氣密落地門，配 5+5 膠合強化隔音玻璃"),
                ("4-2", "三間臥房隔音氣密窗", "3", "樘", "16,000", "48,000", "氣密窗配 5+5 膠合玻璃、紗窗、高氣密止風塊與防墜安全鎖"),
                ("4-3", "後陽台三合一通風門", "1", "樘", "18,000", "18,000", "廚房通後陽台三合一鋁門（含紗窗、百葉、玻璃上下滑動）"),
                ("4-4", "窗框塞水路與防水矽利康", "1", "式", "26,000", "26,000", "高壓水泥塞水路填實，外框施打道康寧耐候中性矽利康雙道")
            ]
        },
        {
            "cat_title": "五、 雙衛浴設備與五金配備工程（小計：$106,000）",
            "desc": "換裝省水馬桶、防潮浴櫃面盆與主衛暖風機，滿足自住家庭基本衛生與乾燥防霉需求。",
            "items": [
                ("5-1", "雙衛浴省水單體馬桶", "2", "套", "13,000", "26,000", "TOTO 或 Derek 金級省水單體馬桶，含原廠緩降便座"),
                ("5-2", "防潮發泡板浴櫃面盆組", "2", "套", "12,000", "24,000", "陶瓷面盆搭配 100% 防水發泡板懸空浴櫃、冷熱水龍頭"),
                ("5-3", "主衛乾濕分離拉門", "1", "套", "18,000", "18,000", "8mm 強化安全玻璃淋浴拉門，配不鏽鋼懸吊滑軌與防漏條"),
                ("5-4", "淋浴花灑與五金配件", "2", "套", "8,000", "16,000", "冷熱水淋浴花灑水龍頭、除霧防潮鏡、不鏽鋼置物毛巾架"),
                ("5-5", "主衛四合一暖風乾燥機", "1", "台", "16,000", "16,000", "樂奇/台達電暖風機（暖房/乾燥/涼風/換氣），含排風配管"),
                ("5-6", "客衛超靜音排風扇", "1", "台", "6,000", "6,000", "靜音排風抽風機，含逆止風閥防異味管路施作")
            ]
        },
        {
            "cat_title": "六、 廚房系統廚具與三機設備工程（小計：$137,000）",
            "desc": "34 年老廚房木板受潮腐蝕發黑，全面更新為基本耐用系統廚具與符合安全規範三機。",
            "items": [
                ("6-1", "240cm 一字型系統廚具", "1", "式", "58,000", "58,000", "E1 級防潮防蟑低甲醛櫃體、結晶鋼烤門板、Blum 緩衝滑軌"),
                ("6-2", "石英石料理檯面與水槽", "1", "式", "32,000", "32,000", "耐刮耐高溫石英石檯面，配不鏽鋼大單水槽、抽拉冷熱龍頭"),
                ("6-3", "櫻花/林內 強力排油煙機", "1", "台", "15,000", "15,000", "大吸力雙渦輪排油煙機，含全新鋁質排煙軟管接至室外"),
                ("6-4", "櫻花/林內 雙口玻璃瓦斯爐", "1", "台", "12,000", "12,000", "黑色強化玻璃檯面、雙口電子點火、附防乾燒自動熄火裝置"),
                ("6-5", "懸吊式烘碗機", "1", "台", "8,000", "8,000", "臭氧殺菌烘碗機（雙層碗盤架、電子觸控定時）"),
                ("6-6", "廚房料理區防濺烤漆玻璃", "1", "式", "12,000", "12,000", "瓦斯爐前方與流理台牆面強化烤漆玻璃（防油污好清理）")
            ]
        },
        {
            "cat_title": "七、 全室油漆粉刷工程（小計：$93,000）",
            "desc": "拆除固定櫃與暗管埋設後牆面坑洞遍布，徹底修補裂縫、全批土細磨後塗刷環保乳膠漆。",
            "items": [
                ("7-1", "全室牆面裂縫整平防裂", "1", "式", "20,000", "20,000", "清除原脫漆，填縫劑填補、接縫施作玻纖網/AB膠防裂"),
                ("7-2", "全室全面全批土兩道及研磨", "1", "式", "38,000", "38,000", "室內牆面與天花板全批土兩道，砂紙機細緻研磨平整無痕"),
                ("7-3", "全室環保乳膠漆塗刷兩道", "1", "式", "35,000", "35,000", "採用得利 (Dulux) 或虹牌低甲醛乳膠漆，滾輪粉刷兩度")
            ]
        },
        {
            "cat_title": "八、 門扇更換、局部修飾與全室地坪工程（小計：$138,000）",
            "desc": "原木門受潮變形關不緊；舊地磚冰冷多處空心，整平後直鋪 SPC 石塑地板為最實惠安全的翻新方式。",
            "items": [
                ("8-1", "臥房門片更新（含鎖組）", "3", "樘", "9,000", "27,000", "三組耐磨實木複合門片、防撞門擋、水平靜音鎖"),
                ("8-2", "衛浴防潮門片更新", "2", "樘", "7,000", "14,000", "兩組防水塑鋼門片（含通風百葉、防潮門框、水平防潮鎖）"),
                ("8-3", "水電管線局部包管平釘", "1", "式", "32,000", "32,000", "日本麗仕耐燃矽酸鈣板，局部修飾冷氣銅管與電線幹管"),
                ("8-4", "全室 SPC 石塑防水地板", "25", "坪", "2,600", "65,000", "直鋪 5.5mm SPC 石塑地板（附靜音墊），含門檻壓條收邊")
            ]
        }
    ]

    for cat in detailed_categories:
        cp = doc.add_paragraph()
        cp.paragraph_format.space_before = Pt(10)
        cp.paragraph_format.space_after = Pt(2)
        r_cname = cp.add_run(cat['cat_title'])
        format_run(r_cname, size_pt=11, bold=True, color_rgb=(30, 58, 138))

        cd_p = doc.add_paragraph()
        cd_p.paragraph_format.space_after = Pt(4)
        r_cd = cd_p.add_run(f"• 必要性說明：{cat['desc']}")
        format_run(r_cd, size_pt=9, italic=True, color_rgb=(100, 116, 139))

        # Item Table
        t_cat = doc.add_table(rows=len(cat['items']) + 1, cols=7)
        t_cat.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_cat.autofit = False

        c_widths = [Inches(0.55), Inches(1.8), Inches(0.45), Inches(0.45), Inches(0.85), Inches(0.85), Inches(1.85)]
        c_headers = ["項目", "施工內容與規格", "數量", "單位", "單價", "複價", "備註說明"]

        for ci, ch in enumerate(c_headers):
            cell = t_cat.cell(0, ci)
            cell.width = c_widths[ci]
            set_cell_background(cell, "F1F5F9")
            set_cell_borders(cell, top="CBD5E1", bottom="94A3B8", left="none", right="none", sz="4")
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if ci in [0, 2, 3]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif ci in [4, 5]:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(ch)
            format_run(r, size_pt=8.5, bold=True, color_rgb=(51, 65, 85))

        for ri, item_row in enumerate(cat['items'], 1):
            bg = "FFFFFF" if ri % 2 == 1 else "F8FAFC"
            for ci, ival in enumerate(item_row):
                cell = t_cat.cell(ri, ci)
                cell.width = c_widths[ci]
                set_cell_background(cell, bg)
                set_cell_borders(cell, top="E2E8F0", bottom="E2E8F0", left="none", right="none", sz="4")
                set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                if ci in [0, 2, 3]:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif ci in [4, 5]:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(ival)
                format_run(r, size_pt=8.5, bold=(ci in [0, 5]), color_rgb=(15, 23, 42) if ci != 6 else (100, 116, 139))

    doc.add_paragraph().paragraph_format.space_after = Pt(14)

    # 6. Negotiation Notes
    memo_h = doc.add_paragraph()
    memo_h.paragraph_format.space_before = Pt(14)
    memo_h.paragraph_format.space_after = Pt(6)
    r_mh = memo_h.add_run("💡 買方與房仲談判實務運用備忘錄")
    format_run(r_mh, size_pt=13, bold=True, color_rgb=(30, 58, 138))

    memo_table = doc.add_table(rows=1, cols=1)
    memo_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    mcell = memo_table.cell(0, 0)
    mcell.width = Inches(6.8)
    set_cell_background(mcell, "EFF6FF")
    set_cell_left_border_only(mcell, color_hex="2563EB", sz="28")
    set_cell_margins(mcell, top=140, bottom=140, left=200, right=200)

    mp = mcell.paragraphs[0]
    mp.paragraph_format.line_spacing = 1.25
    mp.paragraph_format.space_after = Pt(4)
    r_mt1 = mp.add_run("1. 實單出示的心理震撼（瓦解屋主高價幻想）：\n")
    format_run(r_mt1, size_pt=10, bold=True, color_rgb=(30, 58, 138))
    r_md1 = mp.add_run(
        "• 當屋主嫌買方出 710 萬太低時，業務可直接將此份「基礎工程估價總表（120 萬）」實體出示給大姐過目。\n"
        "• 業務說法：「大姐，買方不是空口說白話亂砍價，這是他請統包師傅實地看過後開出來最陽春、完全沒有奢華裝潢的『保命基礎工程』。水電管線 34 年不換會走火、兩間廁所不重做會漏水到樓下、舊窗戶風吹雨打會滲水。買方是理性自住家庭，這 120 萬是不可避免的剛性安全支出（規劃搭配裝潢修繕貸款支應）。在嚴格家庭預算紀律下，這是一毛錢都省不掉的剛性開銷！」\n\n"
    )
    format_run(r_md1, size_pt=9.5, bold=False, color_rgb=(51, 65, 85))

    r_mt2 = mp.add_run("2. 堅守「850 萬總預算負擔」精算公式：\n")
    format_run(r_mt2, size_pt=10, bold=True, color_rgb=(30, 58, 138))
    r_md2 = mp.add_run(
        "• 公式：買方極限總承擔 (850 萬) - 必要修繕貸款 (120 萬) = 房屋買價極限天花板 (730 萬)。\n"
        "• 買方從 710 萬起手下斡旋，就是為了留 20 萬的加價空間給房仲發揮，目標在 730 萬達成圓滿成交。\n"
        "• 房仲面對屋主時「有數字、有單據、有算盤、有底氣」，能完全瓦解屋主防線，促成理性成交！"
    )
    format_run(r_md2, size_pt=9.5, bold=False, color_rgb=(51, 65, 85))

    # Output path
    output_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "大墩人文", "大墩人文11樓_老屋翻新基礎工程預算估價報價單.docx")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    print(f"✅ 成功生成估價單 Word 檔案：{output_path}")

if __name__ == "__main__":
    main()
