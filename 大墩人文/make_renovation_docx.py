import os
import sys

if sys.stdout:
    sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    """設定儲存格背景色"""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """設定儲存格內邊距 (dxa)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def create_renovation_quotation_docx(output_path):
    doc = Document()

    # 設定頁面邊距 (標準公文/工程報價單版面：上下左右各 2 cm)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # 基礎樣式字型
    normal_style = doc.styles['Normal']
    normal_style.font.name = '微軟正黑體'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(51, 51, 51)

    # 1. 公司抬頭與報價單標題
    p_comp = doc.add_paragraph()
    p_comp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_comp = p_comp.add_run("盛祥空間工程・室內裝修專業統包工程行\n統一編號：83921045  |  工程聯絡電話：04-2380-XXXX")
    r_comp.font.size = Pt(8.5)
    r_comp.font.color.rgb = RGBColor(120, 120, 120)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run("【大墩人文 11 樓】全屋基礎翻新工程詳細估價報價單")
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(26, 54, 93) # 深海軍藍

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    r_sub = p_sub.add_run("本工程為純老屋結構安全、水電、防水與生活機能基礎統包工程（不含活動家具、不含家電設備）")
    r_sub.font.size = Pt(9.5)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    # 2. 報價基本資訊表格 (業主、案場、日期、付款方式等)
    info_table = doc.add_table(rows=4, cols=4)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.autofit = False

    info_data = [
        [("業主名稱", True), ("張先生 / 張小姐", False), ("報價單號", True), ("SX-20261014-01", False)],
        [("工程案場", True), ("台中市南屯區大墩人文（原大墩國宅）11樓", False), ("報價日期", True), ("2026 年 10 月 14 日", False)],
        [("規劃坪數", True), ("建物權狀約 31.54 坪（室內全翻新）", False), ("預估工期", True), ("約 60 ~ 75 工作天（依進度分期）", False)],
        [("報價有效", True), ("自報價日起 30 日內有效", False), ("付款條件", True), ("訂約30% / 泥作30% / 木作20% / 驗收20%", False)]
    ]

    col_widths = [Inches(1.0), Inches(2.5), Inches(1.0), Inches(2.4)]
    for r_idx, row in enumerate(info_table.rows):
        for c_idx, cell in enumerate(row.cells):
            cell.width = col_widths[c_idx]
            text, is_hdr = info_data[r_idx][c_idx]
            cell.text = text
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.runs[0]
            r.font.size = Pt(9)
            if is_hdr:
                set_cell_background(cell, "F1F5F9")
                r.font.bold = True
                r.font.color.rgb = RGBColor(30, 41, 59)
            else:
                set_cell_background(cell, "FFFFFF")
                r.font.color.rgb = RGBColor(51, 65, 85)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)

    # 3. 總預算摘要表 (大類彙整)
    p_sum_title = doc.add_paragraph()
    p_sum_title.paragraph_format.space_before = Pt(14)
    p_sum_title.paragraph_format.space_after = Pt(4)
    r_st = p_sum_title.add_run("壹、 工程大類費用彙整總表")
    r_st.font.size = Pt(12)
    r_st.font.bold = True
    r_st.font.color.rgb = RGBColor(30, 58, 138)

    summary_table = doc.add_table(rows=1, cols=5)
    summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    summary_table.autofit = False

    s_headers = ["大類編號", "工程大類類別", "施作重點與涵蓋範圍", "預估金額 (新台幣)", "佔比"]
    s_widths = [Inches(0.8), Inches(1.5), Inches(3.0), Inches(1.1), Inches(0.5)]

    hdr_cells = summary_table.rows[0].cells
    for i, title in enumerate(s_headers):
        hdr_cells[i].text = title
        hdr_cells[i].width = s_widths[i]
        set_cell_background(hdr_cells[i], "1E3A8A") # 沉穩深藍
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_margins(hdr_cells[i], top=100, bottom=100, left=100, right=100)

    categories = [
        ("一", "公共保護與拆除清運工程", "梯廳電梯保護、客廳書桌書櫃、三間房冷氣下櫃、廚房舊廚具、雙衛浴打到底見RC、廢棄物合法清運(3.5車)", 152460),
        ("二", "水電管線與電力系統工程", "全屋電線重拉(太平洋CNS)、儲藏室70A加大配電箱與士林電機NFB、專用迴路7組、全室開關插座面板(Panasonic)、冷熱水管重拉", 218100),
        ("三", "全屋氣密隔音窗與門扇工程", "客廳200x160雙扇氣密窗、廚房160x160氣密窗、主臥及兩次臥200x160氣密窗3扇(共5大窗+後陽台三合一門)、臥室與儲藏室門片4組", 187990),
        ("四", "雙衛浴打除重建泥作防水工程", "客衛(225x200)+主衛(275x225)打到底、主衛壁癌根治、客衛破損補修、全室彈性水泥防水刷到頂(試水48hr)、防滑地壁磚鋪貼", 241490),
        ("五", "雙衛浴設備與五金配備工程", "雙衛浴省水單體馬桶、發泡板防潮浴櫃面盆組、冷熱龍頭花灑配件、主衛一字型強化玻璃拉門、四合一暖風機、靜音排風扇", 106300),
        ("六", "廚房全新3米系統廚具與三機", "3米一字型系統下櫃+3米吊櫥櫃、石英石耐磨料理檯面、不鏽鋼大水槽、櫻花/林內雙口瓦斯爐、排油煙機、烘碗機、防濺烤漆玻璃", 145200),
        ("七", "全室油漆粉刷修補工程", "全室天花板與立面牆面裂縫修補、AB膠填縫抗裂、批土研磨整平粉光、得利/虹牌平光環保乳膠漆二度粉刷", 88260),
        ("八", "全屋燈具、窗簾與安全基礎設備", "客廳+餐廳+廚房+陽台+三臥室LED吸頂燈具與全室安裝、全室遮光與調光窗簾、玄關指紋電子鎖、後陽台手搖雙桿曬衣架", 80200)
    ]

    total_amount = sum(c[3] for c in categories)

    for cat_id, cat_name, cat_desc, cat_amt in categories:
        row_cells = summary_table.add_row().cells
        pct = (cat_amt / total_amount) * 100
        row_cells[0].text = cat_id
        row_cells[1].text = cat_name
        row_cells[2].text = cat_desc
        row_cells[3].text = f"${cat_amt:,}"
        row_cells[4].text = f"{pct:.1f}%"

        for idx, cell in enumerate(row_cells):
            cell.width = s_widths[idx]
            p = cell.paragraphs[0]
            r = p.runs[0]
            r.font.size = Pt(9)
            if idx in [0, 4]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif idx == 3:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                r.font.bold = True
                r.font.color.rgb = RGBColor(30, 41, 59)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            set_cell_margins(cell, top=70, bottom=70, left=80, right=80)

    # 總計列
    total_row = summary_table.add_row().cells
    total_row[0].text = "總計"
    total_row[1].text = "全屋基礎統包總工程款"
    total_row[2].text = "包含全部拆除清運、水電、泥作、防水、門窗、衛浴、廚具、油漆與燈具五金（未稅）"
    total_row[3].text = f"${total_amount:,}"
    total_row[4].text = "100%"

    for idx, cell in enumerate(total_row):
        cell.width = s_widths[idx]
        set_cell_background(cell, "FEF3C7") # 暖琥珀淡黃底
        p = cell.paragraphs[0]
        r = p.runs[0]
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(180, 83, 9)
        if idx in [0, 4]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif idx == 3:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r.font.size = Pt(11)
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_cell_margins(cell, top=100, bottom=100, left=80, right=80)

    # 4. 細項工程分析表
    p_det_title = doc.add_paragraph()
    p_det_title.paragraph_format.space_before = Pt(18)
    p_det_title.paragraph_format.space_after = Pt(4)
    r_dt = p_det_title.add_run("貳、 八大工種詳細施工項目與計價明細表")
    r_dt.font.size = Pt(12)
    r_dt.font.bold = True
    r_dt.font.color.rgb = RGBColor(30, 58, 138)

    detail_table = doc.add_table(rows=1, cols=7)
    detail_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    detail_table.autofit = False

    d_headers = ["工項編號", "施工項目名稱", "詳細施工內容與材料規格說明", "數量", "單位", "單價(NTD)", "複價(NTD)"]
    d_widths = [Inches(0.7), Inches(1.3), Inches(2.6), Inches(0.4), Inches(0.4), Inches(0.8), Inches(0.7)]

    d_hdr_cells = detail_table.rows[0].cells
    for i, title in enumerate(d_headers):
        d_hdr_cells[i].text = title
        d_hdr_cells[i].width = d_widths[i]
        set_cell_background(d_hdr_cells[i], "0F172A") # 深黑石板色
        p = d_hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.runs[0]
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_margins(d_hdr_cells[i], top=90, bottom=90, left=70, right=70)

    # 細項資料清單 (工項編號, 項目名稱, 規格內容, 數量, 單位, 單價, 複價, 所屬大類)
    details_data = [
        # 一、拆除清運 (小計: $152,460)
        ("1-1", "公區走道梯廳保護工程", "電梯車廂四面夾板+11樓梯廳及轉角鋪設防潮布與防撞PVC夾板(管委會規範)", "1", "式", 9860, 9860, "一、 公共保護與拆除清運工程（小計：$152,460）"),
        ("1-2", "客廳書桌書櫃拆除", "進門左側大型木作展示書櫃、一體成型書桌拆解下架與拔釘", "1", "式", 7840, 7840, None),
        ("1-3", "三間臥室冷氣下櫃拆除", "主臥、次臥一、次臥二冷氣下方窗邊木作矮櫃拆除(3處各約$2,480)", "3", "處", 2480, 7440, None),
        ("1-4", "舊廚房廚具整套拆除", "拆除舊一字型流理台、瓦斯爐台、排油煙機、上下廚櫃及拉門門斗", "1", "式", 7320, 7320, None),
        ("1-5", "雙衛浴磁磚設備打除見底", "客衛(225x200)+主衛(275x225)舊地壁磚、水泥砂漿層全面打鑿見RC紅磚底", "2", "間", 14760, 29520, None),
        ("1-6", "全室老舊鋁窗及雜項拆除", "客廳2大扇+廚房1扇+三間臥房3大扇舊鋁窗框拆除卸下、雜項木作拆卸", "1", "式", 11680, 11680, None),
        ("1-7", "室內廢料裝袋與搬運下樓", "所有混凝土石塊、紅磚、舊磁磚、木屑垃圾人工裝袋並搬運至一樓集中區", "1", "式", 45480, 45480, None),
        ("1-8", "合法環保清運車趟處理費", "政府核可合法環保清運車進場(3.5車次)，含廢棄物合法掩埋棄土場處理費", "3.5", "車", 9520, 33320, None),

        # 二、水電管線與電力系統 (小計: $218,100)
        ("2-1", "全屋電源線重拉抽換", "全室各迴路暗管線路抽換，採用太平洋一級CNS認證2.0mm單芯/5.5mm絞線", "31.5", "坪", 1860, 58590, "二、 水電管線更新與電力系統工程（小計：$218,100）"),
        ("2-2", "儲藏室70A保險開關更新", "更換全新大型匯流排配電箱、70A總開關、士林電機無熔絲開關(NFB)與漏電斷路器", "1", "式", 22860, 22860, None),
        ("2-3", "全屋專用迴路獨立配置", "冷氣專迴4組(客廳+三臥房)、廚房高耗能專迴2組(微波/烤箱)、浴室暖風機專迴1組", "7", "組", 3080, 21560, None),
        ("2-4", "全屋不鏽鋼冷熱水管重拉", "室內全面重配SUS304不鏽鋼被覆保溫水管，徹底汰換34年鏽蝕鍍鋅鐵管暗管", "1", "式", 47680, 47680, None),
        ("2-5", "排水與排污幹管整修更新", "雙衛浴落水頭、廚房排水及後陽台地排更新，施作防臭存水彎與排污密封", "1", "式", 25720, 25720, None),
        ("2-6", "全屋開關與插座面板更新", "全室開關、插座全面更換為新式Panasonic COSMO/星光系列大面板與接地插座", "1", "式", 17850, 17850, None),
        ("2-7", "水電暗管開槽打鑿泥作修補", "配合新電線專迴及水管開槽打鑿，新管固定後以水泥砂漿高標平整回填修復", "1", "式", 23840, 23840, None),

        # 三、鋁門窗與門扇 (小計: $187,990)
        ("3-1", "客廳200x160鋁門窗更新", "客廳兩扇200cm*160cm氣密鋁門窗(正新/創世紀同級)，配5+5膠合安全隔音玻璃", "2", "樘", 21850, 43700, "三、 全屋氣密隔音窗與門扇工程（小計：$187,990）"),
        ("3-2", "廚房160x160鋁門窗更新", "廚房160cm*160cm隔音氣密窗一扇，配5+5膠合玻璃與防蚊防墜紗窗", "1", "樘", 15860, 15860, None),
        ("3-3", "主臥200x160鋁門窗更新", "主臥房200cm*160cm大面隔音氣密窗一扇，配5+5膠合隔音玻璃與紗窗", "1", "樘", 21850, 21850, None),
        ("3-4", "次臥一200x160鋁窗更新", "次臥一200cm*160cm大面隔音氣密窗一扇，配5+5膠合隔音玻璃與紗窗", "1", "樘", 21850, 21850, None),
        ("3-5", "次臥二200x160鋁窗更新", "次臥二200cm*160cm大面隔音氣密窗一扇，配5+5膠合隔音玻璃與紗窗", "1", "樘", 21850, 21850, None),
        ("3-6", "廚房後陽台通風三合一門", "廚房通後陽台三合一防盜通風鋁門(含紗窗、防盜格柵百葉、上下拉動玻璃扇)", "1", "樘", 17820, 17820, None),
        ("3-7", "窗框水泥塞水路及矽利康", "全室窗框周邊高壓水泥砂漿塞水路填實，外框施打陶氏/道康寧耐候型矽利康雙道", "1", "式", 16420, 16420, None),
        ("3-8", "臥室及儲藏室門片更換", "更新主臥、次臥一、次臥二及儲藏室耐磨實木複合門片(共4扇，含框、水平鎖、防撞墊)", "4", "樘", 7160, 28640, None),

        # 四、雙衛浴泥作防水 (小計: $241,490)
        ("4-1", "主臥衛浴牆面壁癌專項根治", "主臥衛浴周邊牆面壁癌徹底剔除至紅磚層，塗佈滲透結晶底漆、抗裂防水砂漿修補", "1", "式", 15850, 15850, "四、 雙衛浴打除重建泥作防水工程（小計：$241,490）"),
        ("4-2", "客廳衛浴牆面破損磁磚修換", "客廳衛浴原破損壁磚剔除整平、結構微損加固修補粉光", "1", "式", 11840, 11840, None),
        ("4-3", "雙衛浴水泥粗胚打底洩水坡", "客衛(225x200)+主衛(275x225)牆面地坪重新水泥粉光整平，地面精密抓1:100洩水坡度", "2", "間", 17860, 35720, None),
        ("4-4", "雙衛浴彈性水泥全室防水", "全室雙劑型高彈性水泥塗刷三道(塗至天花板頂)，角隅鋪貼玻璃纖維網抗裂，試水48hr", "2", "間", 24150, 48300, None),
        ("4-5", "雙衛浴修邊壁磚鋪貼工料", "主衛+客衛牆面鋪貼30x60cm一級防潮修邊壁磚，含專用磁磚黏著劑與防霉填縫工料", "2", "間", 42130, 84260, None),
        ("4-6", "雙衛浴防滑地磚鋪貼工料", "主衛+客衛地坪鋪貼30x30cm高規止滑地磚，含乾濕分離人造石門檻、落水頭按裝", "2", "間", 22760, 45520, None),

        # 五、雙衛浴設備五金 (小計: $106,300)
        ("5-1", "雙衛浴省水單體馬桶", "TOTO/Derek 金級省水單體馬桶 2 座，含原廠緩降便座、排污防臭密封法蘭", "2", "座", 13250, 26500, "五、 雙衛浴設備與五金配備工程（小計：$106,300）"),
        ("5-2", "防潮浴櫃陶瓷洗手台面盆", "陶瓷洗臉面盆搭配100%防水發泡板下浴櫃組 2 套，含冷熱陶瓷芯臉盆水龍頭", "2", "套", 12180, 24360, None),
        ("5-3", "冷熱沐浴水龍頭與配件", "冷熱水淋浴花灑水龍頭組、除霧防潮明鏡箱、304不鏽鋼雙層毛巾置物架五金", "2", "套", 7460, 14920, None),
        ("5-4", "主衛一字型乾濕分離拉門", "主臥衛浴一字型 8mm 安全強化玻璃淋浴拉門(不鏽鋼滑軌、防爆膜、磁吸防漏邊條)", "1", "套", 16850, 16850, None),
        ("5-5", "主衛四合一多功能暖風機", "樂奇/台達電四合一暖風乾燥機(暖房/乾燥/涼風/換氣)，含天花板風管及專用迴路連接", "1", "台", 15820, 15820, None),
        ("5-6", "客衛超靜音逆止風排風扇", "超靜音排風扇，含防異味逆止風閥管路連接安裝", "1", "台", 7850, 7850, None),

        # 六、廚房3米系統廚具 (小計: $145,200)
        ("6-1", "3米一字型系統料理下櫃", "長度300cm系統下櫃，防潮防蟑E1低甲醛桶身、結晶鋼烤門板、進口緩衝鉸鏈滑軌抽屜", "1", "式", 47850, 47850, "六、 廚房全新3米系統廚具與三機（小計：$145,200）"),
        ("6-2", "3米一字型系統收納吊櫃", "長度300cm防潮吊櫥櫃，增加廚房上層大量備品收納空間，含緩衝氣壓支撐五金", "1", "式", 25680, 25680, None),
        ("6-3", "石英石耐磨檯面與不鏽鋼槽", "300cm實心耐刮耐高溫石英石料理檯面，配不鏽鋼大單水槽、防臭提籠、抽拉冷熱龍頭", "1", "式", 32450, 32450, None),
        ("6-4", "雙口檯面玻璃安全瓦斯爐", "櫻花/林內雙口黑色強化玻璃檯面瓦斯爐，電子點火、附防乾燒自動熄火安全裝置", "1", "台", 11250, 11250, None),
        ("6-5", "強力排油煙機建置", "櫻花/林內雙渦輪隱藏/斜背式強力吸力排油煙機，含全新鋁質排煙管接至戶外", "1", "台", 12180, 12180, None),
        ("6-6", "懸吊式臭氧殺菌烘碗機", "櫻花/喜特麗雙層大容量懸吊式烘碗機，電子觸控定時、臭氧殺菌", "1", "台", 7850, 7850, None),
        ("6-7", "烹飪料理區防濺烤漆玻璃", "瓦斯爐及料理水槽前牆面鋪設5mm厚強化烤漆玻璃(耐高溫、防油污、好擦拭)", "1", "式", 7940, 7940, None),

        # 七、全室油漆 (小計: $88,260)
        ("7-1", "全室牆面裂縫整平與防裂", "剔除舊脫漆浮層，深淺裂縫填縫整平、水電埋管開槽處貼玻纖網與AB膠防裂", "1", "式", 21850, 21850, "七、 全室油漆粉刷修補工程（小計：$88,260）"),
        ("7-2", "全室牆面天花批土研磨", "客餐廳、走道及三間臥室所有天花板與立面全批土兩道，打磨機細緻研磨平整無刷痕", "1", "式", 34260, 34260, None),
        ("7-3", "全室平光環保乳膠漆二度", "採用得利(Dulux)或虹牌環保平光乳膠漆二度塗刷，無毒低甲醛、均勻飽滿防霉", "1", "式", 32150, 32150, None),

        # 八、全屋燈具窗簾安全設備 (小計: $80,200)
        ("8-1", "客廳主照明LED大吸頂燈", "高流明低頻閃 LED 調光調色吸頂燈(約75W)，含高樓打釘定位與水電安裝工資", "1", "盞", 4650, 4650, "八、 全屋燈具、窗簾與安全基礎設備（小計：$80,200）"),
        ("8-2", "餐廳美觀造型LED主吊燈", "耐看簡約北歐風 LED 節能餐吊燈，營造溫馨就餐氛圍，含安裝工資", "1", "盞", 3480, 3480, None),
        ("8-3", "三間臥室護眼LED吸頂燈", "主臥、次臥一、次臥二配置 40W 護眼無頻閃 LED 吸頂燈 3 盞(含安裝工資)", "3", "盞", 2460, 7380, None),
        ("8-4", "廚房走道陽台防水平板燈", "廚房防油煙平板燈、走道吸頂燈、前後陽台防潮戶外吸頂燈共 5 盞(含安裝)", "5", "盞", 1090, 5450, None),
        ("8-5", "客廳與三臥室遮光窗簾", "客廳配置透光調光捲簾，主臥及兩次臥配置高遮光防紫外線抗熱捲簾(含五金軌道與安裝)", "1", "式", 28860, 28860, None),
        ("8-6", "玄關大門智慧指紋電子鎖", "知名品牌四合一智慧電子鎖(指紋/密碼/感應卡/機械鑰匙)，含不鏽鋼大門開孔與技師安裝", "1", "套", 18250, 18250, None),
        ("8-7", "後陽台手搖雙桿升降曬衣架", "航太級鋁合金雙桿防風曬衣桿(2.4米)+304不鏽鋼抗斷鋼索，手搖省力升降(含天花安裝)", "1", "組", 4580, 4580, None),
        ("8-8", "冷氣四台室內外配管洗孔位", "配合客廳與三臥室四台新冷氣之冷媒銅管貫穿孔洗孔、外牆防雨出水接管修飾", "1", "式", 7550, 7550, None)
    ]

    current_cat_banner = None

    for item in details_data:
        num, name, spec, qty, unit, unit_price, sub_total, cat_banner = item

        # 若有新的大類橫幅，插入合併大類標題行
        if cat_banner:
            banner_row = detail_table.add_row().cells
            # 合併全部儲存格
            a = banner_row[0]
            b = banner_row[6]
            a.merge(b)
            a.text = cat_banner
            set_cell_background(a, "E2E8F0")
            p = a.paragraphs[0]
            r = p.runs[0]
            r.font.bold = True
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(15, 23, 42)
            set_cell_margins(a, top=80, bottom=80, left=80, right=80)

        row_cells = detail_table.add_row().cells
        row_cells[0].text = num
        row_cells[1].text = name
        row_cells[2].text = spec
        row_cells[3].text = str(qty)
        row_cells[4].text = unit
        row_cells[5].text = f"${unit_price:,}"
        row_cells[6].text = f"${sub_total:,}"

        for idx, cell in enumerate(row_cells):
            cell.width = d_widths[idx]
            p = cell.paragraphs[0]
            r = p.runs[0]
            r.font.size = Pt(8.5)
            if idx in [0, 3, 4]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif idx in [5, 6]:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                if idx == 6:
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(30, 41, 59)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            set_cell_margins(cell, top=50, bottom=50, left=60, right=60)

    # 細項總計列
    d_total_row = detail_table.add_row().cells
    d_total_row[0].merge(d_total_row[5])
    d_total_row[0].text = "【大墩人文 11 樓】老屋全室基礎翻新統包總工程款合計（未稅）："
    d_total_row[6].text = f"${total_amount:,}"

    set_cell_background(d_total_row[0], "FEF3C7")
    set_cell_background(d_total_row[6], "FEF3C7")
    p0 = d_total_row[0].paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r0 = p0.runs[0]
    r0.font.bold = True
    r0.font.size = Pt(10)
    r0.font.color.rgb = RGBColor(180, 83, 9)

    p6 = d_total_row[6].paragraphs[0]
    p6.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r6 = p6.runs[0]
    r6.font.bold = True
    r6.font.size = Pt(11)
    r6.font.color.rgb = RGBColor(180, 83, 9)
    set_cell_margins(d_total_row[0], top=100, bottom=100, left=80, right=80)
    set_cell_margins(d_total_row[6], top=100, bottom=100, left=80, right=80)

    # 5. 施工備註與統包商專業條款
    p_notes_title = doc.add_paragraph()
    p_notes_title.paragraph_format.space_before = Pt(16)
    p_notes_title.paragraph_format.space_after = Pt(4)
    r_nt = p_notes_title.add_run("參、 統包工程合約備註與品質保證條款")
    r_nt.font.size = Pt(11)
    r_nt.font.bold = True
    r_nt.font.color.rgb = RGBColor(30, 58, 138)

    notes = [
        "1. 工程範圍說明：本報價單為純自住安全性與機能性之「老屋基礎翻新統包工程」，不包含家電設備（冷氣機本體、電視機、洗衣機、冰箱等）以及移動式活動家具（沙發、餐桌椅、床墊等）。冷氣部分僅含冷媒管線埋管配置、洗孔與專用電源迴路建置。",
        "2. 舊屋安全責任保證：全室電線均使用一級大廠（太平洋/大亞）經濟部商品檢驗局 CNS 認證標準線徑；所有冷熱水管全數更換為不鏽鋼被覆管，嚴禁局部銜接舊管，確保 20 年內暗管無爆管滲水之虞。",
        "3. 衛浴防水與試水保固：雙衛浴全打鑿見 RC 底層重新施作泥作與高抗裂彈性水泥防水塗佈，防水層一律刷至天花板頂；貼磚前必須會同業主進行 48 小時蓄水試水測試，確認無任何滲漏始得鋪貼磁磚。本公司提供防水工程 2 年品質保固。",
        "4. 氣密隔音窗施工規範：鋁門窗均依業主現場實測尺寸開模訂製，採乾式包框或泥作嵌縫施工，外框周邊施打陶氏耐候型防霉矽利康，配置 5+5 膠合強化安全玻璃，有效抵抗高樓風壓與商圈噪音。",
        "5. 廢棄物環保清運：拆除物全數依台中市環保局規定，送往合法合格土石方處理場與廢棄物回收機構掩埋，絕無違法傾倒之法律風險，報價已內含環保處理費與運費。",
        "6. 驗收與付款方式：依工程進度分為四期支付（第一期：簽約動工款 30%；第二期：拆除水電配管與泥作防水完成 30%；第三期：鋁門窗廚具安裝與油漆批土完成 20%；第四期：全室完工細部驗收合格 20%）。"
    ]

    for n in notes:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Inches(0.2)
        r = p.add_run(n)
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(71, 85, 105)

    # 6. 簽章欄
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(24)
    table_sign = doc.add_table(rows=2, cols=2)
    table_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_sign.autofit = False

    table_sign.rows[0].cells[0].width = Inches(3.4)
    table_sign.rows[0].cells[1].width = Inches(3.4)
    table_sign.rows[1].cells[0].width = Inches(3.4)
    table_sign.rows[1].cells[1].width = Inches(3.4)

    table_sign.rows[0].cells[0].text = "業主確認簽章 (甲方)："
    table_sign.rows[0].cells[1].text = "工程統包廠商簽章 (乙方)："
    table_sign.rows[1].cells[0].text = "\n\n簽名：_____________________\n日期：    年    月    日"
    table_sign.rows[1].cells[1].text = "\n盛祥空間工程・室內裝修專業統包\n負責人：___________________\n日期：    年    月    日"

    for row in table_sign.rows:
        for cell in row.cells:
            p = cell.paragraphs[0]
            r = p.runs[0]
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(30, 41, 59)
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)

    # 儲存檔案
    doc.save(output_path)
    print(f"✅ Word 估價單成功生成：{output_path}")

if __name__ == '__main__':
    target = os.path.join(os.path.dirname(os.path.abspath(__file__)), "大墩人文11樓_老屋翻新基礎統包工程報價單.docx")
    create_renovation_quotation_docx(target)
