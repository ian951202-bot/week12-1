import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import os

# ==========================================
# XML STYLING HELPERS
# ==========================================
def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=120, bottom=120, left=160, right=160):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def set_cell_borders(cell, top="D9D9D9", bottom="D9D9D9", left="D9D9D9", right="D9D9D9", 
                     sz="4", val="single"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{top}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{bottom}"/>
            <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{left}"/>
            <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{right}"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)

def add_callout_box(doc, title, text, box_type='note'):
    # Types: note (Navy), tip (Green), warning (Orange/Red)
    colors = {
        'note': ('1F497D', 'F2F5F8', '【實務重點】'),
        'tip': ('107C41', 'E6F2EA', '【操作訣竅】'),
        'warning': ('D83B01', 'FFF4CE', '【常見陷阱與警語】')
    }
    border_color, bg_color, default_prefix = colors.get(box_type, colors['note'])
    
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, bg_color)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=180)
    
    # Left thick border, others none
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>
            <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    run_t = p.add_run(f"{default_prefix} {title}\n")
    run_t.font.name = "Microsoft JhengHei"
    run_t.font.size = Pt(10)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor.from_string(border_color)
    
    run_b = p.add_run(text)
    run_b.font.name = "Microsoft JhengHei"
    run_b.font.size = Pt(9.5)
    run_b.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    
    # Space after table
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(6)

def add_formula_card(doc, title, formula, explanation):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, "F8F9FA")
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
    set_cell_borders(cell, top="0078D4", bottom="0078D4", left="0078D4", right="0078D4", sz="12")
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    
    run_t = p.add_run(f"公式語法示範：{title}\n")
    run_t.font.name = "Microsoft JhengHei"
    run_t.font.size = Pt(9.5)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(0x00, 0x5A, 0x9E)
    
    run_f = p.add_run(f"    {formula}\n")
    run_f.font.name = "Consolas"
    run_f.font.size = Pt(10)
    run_f.font.bold = True
    run_f.font.color.rgb = RGBColor(0xA8, 0x00, 0x00)
    
    run_e = p.add_run(f"運作原理解析：{explanation}")
    run_e.font.name = "Microsoft JhengHei"
    run_e.font.size = Pt(9)
    run_e.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(6)

def add_image_with_caption(doc, img_path, caption_text, width=Inches(6.0)):
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(3)
        p_img.add_run().add_picture(img_path, width=width)
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(0)
        p_cap.paragraph_format.space_after = Pt(10)
        run_cap = p_cap.add_run(caption_text)
        run_cap.font.name = "Microsoft JhengHei"
        run_cap.font.size = Pt(9)
        run_cap.font.bold = True
        run_cap.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    else:
        print(f"Warning: Image {img_path} not found.")

def add_styled_table(doc, headers, data, col_widths, align_cols=None, font_size=Pt(8.5)):
    table = doc.add_table(rows=len(data) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Header row
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        hdr_cells[i].width = col_widths[i]
        set_cell_background(hdr_cells[i], "1F497D")
        set_cell_margins(hdr_cells[i], top=100, bottom=100, left=100, right=100)
        set_cell_borders(hdr_cells[i], top="1F497D", bottom="1F497D", left="5B9BD5", right="5B9BD5", sz="6")
        
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        for run in p.runs:
            run.font.name = "Microsoft JhengHei"
            run.font.size = font_size
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            
    # Set header repeat across pages
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    # Data rows
    for r_idx, row_values in enumerate(data):
        row_cells = table.rows[r_idx + 1].cells
        bg_color = "F2F5F8" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_values):
            row_cells[c_idx].text = str(val)
            row_cells[c_idx].width = col_widths[c_idx]
            set_cell_background(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=80, bottom=80, left=100, right=100)
            set_cell_borders(row_cells[c_idx], top="D9D9D9", bottom="D9D9D9", left="D9D9D9", right="D9D9D9", sz="4")
            
            p = row_cells[c_idx].paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            
            # Text alignment
            if align_cols and c_idx < len(align_cols):
                align = align_cols[c_idx]
                if align == 'R':
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                elif align == 'C':
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
            for run in p.runs:
                run.font.name = "Microsoft JhengHei"
                run.font.size = font_size
                run.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
                
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(8)
    return table

# ==========================================
# 1. COVER PAGE GENERATION
# ==========================================
def create_cover_page(doc):
    # Cover Section
    sec = doc.sections[0]
    sec.top_margin = Inches(1.0)
    sec.bottom_margin = Inches(1.0)
    sec.left_margin = Inches(1.0)
    sec.right_margin = Inches(1.0)
    sec.different_first_page_header_footer = True
    
    # Top spacing
    p_top = doc.add_paragraph()
    p_top.paragraph_format.space_before = Pt(36)
    
    # Top Organization Category Banner
    p_badge = doc.add_paragraph()
    p_badge.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_badge = p_badge.add_run("MARITIME LOGISTICS & DATA ANALYTICS TRAINING MANUAL")
    run_badge.font.name = "Arial"
    run_badge.font.size = Pt(10)
    run_badge.font.bold = True
    run_badge.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    
    # Main Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(14)
    p_title.paragraph_format.space_after = Pt(8)
    p_title.paragraph_format.line_spacing = 1.2
    run_title = p_title.add_run("航運行政人員 MS Excel\n基礎資料前處理操作實戰指引")
    run_title.font.name = "Microsoft JhengHei"
    run_title.font.size = Pt(28)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
    
    # Decorative Divider Line
    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_before = Pt(4)
    p_line.paragraph_format.space_after = Pt(12)
    run_line = p_line.add_run("―" * 42)
    run_line.font.name = "Arial"
    run_line.font.size = Pt(16)
    run_line.font.bold = True
    run_line.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
    
    # Subtitle
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(36)
    p_sub.paragraph_format.line_spacing = 1.3
    run_sub = p_sub.add_run("從零認識操作介面、奠定資料衛生常識、30筆真實業務示範到300筆大數據情境全解\n"
                           "―― 專為海運貨代、船務行政、港埠營運人員量身打造的電腦技能速成手冊")
    run_sub.font.name = "Microsoft JhengHei"
    run_sub.font.size = Pt(12)
    run_sub.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    
    # Metadata Table
    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    meta_table.autofit = False
    
    meta_data = [
        ("受眾對象：", "航運/海運行政人員、船務助理、貨代OP、港埠行政、初學Excel者"),
        ("教學目標：", "掌握商業資料前處理、資料衛生清洗、字串剖析、日期運算與防錯驗證"),
        ("示範工具：", "Microsoft Excel (Microsoft 365 / Excel 2021+ 相容版)"),
        ("案例實作：", "30筆航運業務教學示範檔 + 300筆大數據多維度實務練習檔"),
        ("編撰單位：", "航運數位化營運與商業分析培訓小組"),
        ("文件版本：", "Version 1.0 正式出版發行 (2026年11月)")
    ]
    
    col_w = [Inches(1.5), Inches(5.0)]
    for r_idx, (lbl, val) in enumerate(meta_data):
        cells = meta_table.rows[r_idx].cells
        cells[0].text = lbl
        cells[0].width = col_w[0]
        cells[1].text = val
        cells[1].width = col_w[1]
        
        for c_idx in range(2):
            set_cell_background(cells[c_idx], "F9FBFD" if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cells[c_idx], top=80, bottom=80, left=100, right=100)
            set_cell_borders(cells[c_idx], top="E1DFDD", bottom="E1DFDD", left="E1DFDD", right="E1DFDD", sz="4")
            p = cells[c_idx].paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            for run in p.runs:
                run.font.name = "Microsoft JhengHei"
                run.font.size = Pt(9.5)
                if c_idx == 0:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
                else:
                    run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
                    
    # Bottom note
    p_foot = doc.add_paragraph()
    p_foot.paragraph_format.space_before = Pt(40)
    run_foot = p_foot.add_run("© 2026 航運行政人員電腦技能培訓系列教材. All Rights Reserved.")
    run_foot.font.name = "Arial"
    run_foot.font.size = Pt(8.5)
    run_foot.font.italic = True
    run_foot.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    
    # Page break after cover
    doc.add_page_break()

# ==========================================
# 2. TABLE OF CONTENTS & LIST OF FIGURES/TABLES
# ==========================================
def create_toc_and_lists(doc):
    # Front matter section break
    sec_fm = doc.add_section(docx.enum.section.WD_SECTION.NEW_PAGE)
    sec_fm.top_margin = Inches(0.8)
    sec_fm.bottom_margin = Inches(0.8)
    sec_fm.left_margin = Inches(0.9)
    sec_fm.right_margin = Inches(0.9)
    sec_fm.header.is_linked_to_previous = False
    sec_fm.footer.is_linked_to_previous = False
    
    # Footer with Roman numerals for front matter
    p_fm_f = sec_fm.footer.paragraphs[0]
    p_fm_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fld_pg = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE \\* Roman"/>')
    p_fm_f._p.append(fld_pg)
    for r in p_fm_f.runs:
        r.font.name = "Arial"
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        
    # TOC Header
    p_toc_title = doc.add_paragraph()
    p_toc_title.paragraph_format.space_before = Pt(12)
    p_toc_title.paragraph_format.space_after = Pt(12)
    run_t = p_toc_title.add_run("章節目錄 (Table of Contents)")
    run_t.font.name = "Microsoft JhengHei"
    run_t.font.size = Pt(18)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
    
    # Insert native dynamic Word TOC field
    p_field = doc.add_paragraph()
    p_field.paragraph_format.space_before = Pt(0)
    p_field.paragraph_format.space_after = Pt(6)
    fld_toc = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="TOC \\o &quot;1-3&quot; \\h \\z \\u"/>')
    p_field._p.append(fld_toc)
    
    # Pre-populated elegant TOC entries for immediate viewing
    toc_entries = [
        ("第一章  Excel 操作介面與新手導航", "1", True),
        ("    1.1 航運行政人員為何必須精通 Excel 資料前處理？", "1", False),
        ("    1.2 Excel 核心操作環境架構全解析", "2", False),
        ("    1.3 核心區域操作實務：功能區、編輯列與狀態列", "3", False),
        ("    1.4 航運行政必備之高頻快捷鍵清單", "4", False),
        ("第二章  資料前處理應具備的基本常識", "5", True),
        ("    2.1 什麼是資料前處理？垃圾進、垃圾出 (GIGO) 原則", "5", False),
        ("    2.2 關聯式資料庫與第一正規化 (1NF)：分析型表格準則", "6", False),
        ("    2.3 航運行政常見的『七大骯髒資料』病症深度剖析", "7", False),
        ("    2.4 Excel 四大資料型態與隱形陷阱", "9", False),
        ("第三章  30筆航運事務資料檔：功能選項與函數公式實戰示範", "10", True),
        ("    3.1 30筆航運原始缺陷資料結構總覽", "10", False),
        ("    3.2 功能選項操作示範 1：【資料剖析】拆解船名航次", "12", False),
        ("    3.3 功能選項操作示範 2：【快速填入 (Ctrl+E)】智慧萃取", "14", False),
        ("    3.4 功能選項操作示範 3：【移除重複項】排除重複提單", "15", False),
        ("    3.5 功能選項操作示範 4：【尋找與取代】批次清除字串單位", "16", False),
        ("    3.6 功能選項操作示範 5：【資料驗證】下拉選單源頭防錯", "17", False),
        ("    3.7 功能選項操作示範 6：【自動篩選與排序】快速嗅探髒資料", "18", False),
        ("    3.8 函數公式操作示範 1：文字清洗鏈 (TRIM / UPPER / SUBSTITUTE / MID)", "19", False),
        ("    3.9 函數公式操作示範 2：數值淨化與文字轉數值 (VALUE)", "21", False),
        ("    3.10 函數公式操作示範 3：日期格式標準化與航程天數計算", "22", False),
        ("    3.11 函數公式操作示範 4：遺漏值判斷與代碼補正 (IF / ISBLANK)", "23", False),
        ("    3.12 30筆資料清洗前後成果對照表", "24", False),
        ("第四章  300筆大數據航運實務案例情境練習與解析", "26", True),
        ("    4.1 300筆航運大數據練習檔架構設計與情境說明", "26", False),
        ("    4.2 航運核心主檔參照表 (港口 UN/LOCODE 與 ISO 櫃型)", "27", False),
        ("    4.3 實務情境練習題組 (八大實務任務 Task 1 ~ Task 8)", "28", False),
        ("    4.4 逐步參考解題過程與標準公式解", "30", False),
        ("    4.5 300筆資料前處理成果驗證與分析統計指標", "37", False),
        ("第五章  參考文獻與航運行政 Excel 工具速查表", "38", True),
        ("    5.1 國際海運與資料治理參考文獻來源", "38", False),
        ("    5.2 航運行政常用 Excel 資料前處理函數速查手冊", "39", False),
        ("    5.3 航運行政人員資料分析進階學習藍圖建議", "40", False)
    ]
    
    for title, pg, is_main in toc_entries:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.space_before = Pt(3 if is_main else 1)
        p_item.paragraph_format.space_after = Pt(2)
        p_item.paragraph_format.line_spacing = 1.15
        
        run_t = p_item.add_run(title)
        run_t.font.name = "Microsoft JhengHei"
        run_t.font.size = Pt(10 if is_main else 9)
        run_t.font.bold = is_main
        run_t.font.color.rgb = RGBColor(0x00, 0x20, 0x60) if is_main else RGBColor(0x33, 0x33, 0x33)
        
        # Dot leader calculation
        dots_count = max(70 - len(title) * 2, 8)
        run_dots = p_item.add_run(" " + "." * dots_count + " ")
        run_dots.font.name = "Arial"
        run_dots.font.size = Pt(8.5)
        run_dots.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
        
        run_p = p_item.add_run(pg)
        run_p.font.name = "Arial"
        run_p.font.size = Pt(9.5)
        run_p.font.bold = is_main
        run_p.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        
    doc.add_page_break()
    
    # ----------------- List of Figures & Tables -----------------
    p_fig_title = doc.add_paragraph()
    p_fig_title.paragraph_format.space_before = Pt(12)
    p_fig_title.paragraph_format.space_after = Pt(10)
    run_ft = p_fig_title.add_run("圖表目錄 (List of Figures & Tables)")
    run_ft.font.name = "Microsoft JhengHei"
    run_ft.font.size = Pt(18)
    run_ft.font.bold = True
    run_ft.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
    
    # Subtitle: 圖目錄
    p_sub_f = doc.add_paragraph()
    p_sub_f.paragraph_format.space_before = Pt(6)
    p_sub_f.paragraph_format.space_after = Pt(4)
    run_sf = p_sub_f.add_run("■ 圖目錄 (List of Figures)")
    run_sf.font.name = "Microsoft JhengHei"
    run_sf.font.size = Pt(11)
    run_sf.font.bold = True
    run_sf.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    
    figures_list = [
        ("圖 1-1：Excel 核心操作介面八大功能區導覽示意圖", "2"),
        ("圖 2-1：報表排版型錯誤結構 vs 扁平化分析型標準二維表對比圖", "6"),
        ("圖 3-1：Excel 快速填入 (Flash Fill) 快捷鍵【Ctrl + E】操作拆解", "14"),
        ("圖 3-2：【資料剖析】精靈 3 步驟實戰分解示意圖", "13"),
        ("圖 3-3：【移除重複項】對話方塊設定與提單主鍵排重示範", "15"),
        ("圖 3-4：【尋找與取代】快捷鍵【Ctrl + H】批次剔除字串單位", "16"),
        ("圖 3-5：【資料驗證】下拉選單防錯設定與實體效果展示", "17"),
        ("圖 3-6：【自動篩選】快捷鍵【Ctrl + Shift + L】嗅探髒資料選單", "18"),
        ("圖 3-7：核心公式圖解：文字清洗五大天王函數實戰鏈", "20"),
        ("圖 3-8：日期前處理與航運時效運算三階段流程圖", "22"),
        ("圖 4-1：航運大數據清洗專案架構地圖 (300 筆綜合管線)", "26"),
        ("圖 4-2：現代動態查表關聯：XLOOKUP / VLOOKUP 跨表串接港口基本檔", "31")
    ]
    
    for f_title, pg in figures_list:
        p_f = doc.add_paragraph()
        p_f.paragraph_format.space_before = Pt(1)
        p_f.paragraph_format.space_after = Pt(2)
        run_t = p_f.add_run(f_title)
        run_t.font.name = "Microsoft JhengHei"
        run_t.font.size = Pt(9)
        run_t.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        
        dots = max(66 - len(f_title)*2, 8)
        run_d = p_f.add_run(" " + "." * dots + " ")
        run_d.font.name = "Arial"
        run_d.font.size = Pt(8.5)
        run_d.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
        
        run_p = p_f.add_run(pg)
        run_p.font.name = "Arial"
        run_p.font.size = Pt(9)
        run_p.font.bold = True
        run_p.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        
    # Subtitle: 表目錄
    p_sub_t = doc.add_paragraph()
    p_sub_t.paragraph_format.space_before = Pt(12)
    p_sub_t.paragraph_format.space_after = Pt(4)
    run_st = p_sub_t.add_run("■ 表目錄 (List of Tables)")
    run_st.font.name = "Microsoft JhengHei"
    run_st.font.size = Pt(11)
    run_st.font.bold = True
    run_st.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    
    tables_list = [
        ("表 1-1：航運行政人員必備之 Excel 超高效快捷鍵對照表", "4"),
        ("表 2-1：航運行政常見七大骯髒資料病症、實務後果與解法概覽", "8"),
        ("表 3-1：30 筆航運原始待清洗資料明細清單 (Raw Data)", "10"),
        ("表 3-2：30 筆航運標準化清洗完成結果主表 (Cleaned Master)", "24"),
        ("表 4-1：港口代碼參照主檔 (Port Master - UN/LOCODE)", "27"),
        ("表 4-2：ISO 6346 貨櫃規格代碼與 TEU 換算參照主檔", "28"),
        ("表 4-3：300 筆航運大數據實務練習題組需求規劃表", "29"),
        ("表 5-1：航運行政常用 Excel 資料前處理函數速查手冊", "39")
    ]
    
    for t_title, pg in tables_list:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.space_before = Pt(1)
        p_t.paragraph_format.space_after = Pt(2)
        run_t = p_t.add_run(t_title)
        run_t.font.name = "Microsoft JhengHei"
        run_t.font.size = Pt(9)
        run_t.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        
        dots = max(66 - len(t_title)*2, 8)
        run_d = p_t.add_run(" " + "." * dots + " ")
        run_d.font.name = "Arial"
        run_d.font.size = Pt(8.5)
        run_d.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
        
        run_p = p_t.add_run(pg)
        run_p.font.name = "Arial"
        run_p.font.size = Pt(9)
        run_p.font.bold = True
        run_p.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        
    doc.add_page_break()

# ==========================================
# 3. BODY SECTION SETUP & CHAPTER BUILDERS
# ==========================================
def setup_body_section(doc):
    sec = doc.add_section(docx.enum.section.WD_SECTION.NEW_PAGE)
    sec.top_margin = Inches(0.8)
    sec.bottom_margin = Inches(0.8)
    sec.left_margin = Inches(0.8)
    sec.right_margin = Inches(0.8)
    sec.header.is_linked_to_previous = False
    sec.footer.is_linked_to_previous = False
    
    # Running Header
    p_hdr = sec.header.paragraphs[0]
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_hdr = p_hdr.add_run("航運行政人員 MS Excel 基礎資料分析前處理教學文件")
    r_hdr.font.name = "Microsoft JhengHei"
    r_hdr.font.size = Pt(8.5)
    r_hdr.font.color.rgb = RGBColor(0x77, 0x77, 0x77)
    
    # Bottom border on header
    pBrd = parse_xml(f'''
        <w:pBrd {nsdecls("w")}>
            <w:bottom w:val="single" w:sz="6" w:space="1" w:color="1F497D"/>
        </w:pBrd>
    ''')
    p_hdr._p.get_or_add_pPr().append(pBrd)
    
    # Running Footer: 第 X 頁，共 Y 頁
    p_ftr = sec.footer.paragraphs[0]
    p_ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Restart page numbering at 1
    pgNumType = parse_xml(f'<w:pgNumType {nsdecls("w")} w:start="1"/>')
    sec._sectPr.append(pgNumType)
    
    r_f1 = p_ftr.add_run("第 ")
    r_f1.font.name = "Microsoft JhengHei"
    r_f1.font.size = Pt(9)
    r_f1.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    
    fld1 = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"/>')
    p_ftr._p.append(fld1)
    
    r_f2 = p_ftr.add_run(" 頁，共 ")
    r_f2.font.name = "Microsoft JhengHei"
    r_f2.font.size = Pt(9)
    r_f2.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    
    fld2 = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="NUMPAGES"/>')
    p_ftr._p.append(fld2)
    
    r_f3 = p_ftr.add_run(" 頁")
    r_f3.font.name = "Microsoft JhengHei"
    r_f3.font.size = Pt(9)
    r_f3.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

def add_heading_1(doc, text):
    h = doc.add_heading(text, level=1)
    h.paragraph_format.space_before = Pt(16)
    h.paragraph_format.space_after = Pt(8)
    h.paragraph_format.page_break_before = True # Requirement 9: Every chapter starts on a new page!
    for r in h.runs:
        r.font.name = "Microsoft JhengHei"
        r.font.size = Pt(18)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
    return h

def add_heading_2(doc, text):
    h = doc.add_heading(text, level=2)
    h.paragraph_format.space_before = Pt(12)
    h.paragraph_format.space_after = Pt(6)
    for r in h.runs:
        r.font.name = "Microsoft JhengHei"
        r.font.size = Pt(13.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    return h

def add_heading_3(doc, text):
    h = doc.add_heading(text, level=3)
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(4)
    for r in h.runs:
        r.font.name = "Microsoft JhengHei"
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
    return h

def add_body_p(doc, text, space_after=6, bold=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.25
    r = p.add_run(text)
    r.font.name = "Microsoft JhengHei"
    r.font.size = Pt(10)
    r.font.bold = bold
    r.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    return p

# ==========================================
# CHAPTER 1: EXCEL INTERFACE & NAVIGATION
# ==========================================
def build_chapter_1(doc):
    add_heading_1(doc, "第一章  Excel 操作介面與新手導航")
    
    add_heading_2(doc, "1.1 航運行政人員為何必須精通 Excel 資料前處理？")
    add_body_p(doc, "在現代海運、貨運代理（Freight Forwarder）、報關行與港埠物流營運體系中，航運行政人員每天都必須處理巨量的日常業務檔案。從船公司發布的船期表（Sailing Schedule）、貨代訂艙確認書（Booking Confirmation）、海運提單（Bill of Lading, B/L）、貨櫃動態清單（Container Tracking List），到進出口報關單據與運費應收應付帳單，這些數據絕大多數是以 Excel 或 CSV 檔案格式在上下游供應鏈之間流轉。")
    add_body_p(doc, "然而，航運初學者在面對這些資料時，最常遭遇的並不是複雜的高等統計模型，而是『骯髒資料（Dirty Data）』所引發的系統崩潰與統計錯誤。例如：提單號碼前後夾帶了肉眼難以察覺的空格，導致 VLOOKUP 永遠顯示 #N/A；貨物毛重欄位後面手動輸入了 'kg' 單位，導致 SUM 自動加總算出來竟然是 0；船名航次全部擠在同一個儲存格內，使得行政人員無法按船舶統計總運載量。")
    
    add_callout_box(doc, "航運業務實況：資料分析的二八定律 (80/20 Rule)",
                    "國際資料科學界普遍公認：『資料分析人員有 80% 的時間耗費在資料清洗與前處理，只有 20% 的時間用於產生報表與決策洞察。』對於航運行政人員而言，若未在第一時間掌握資料前處理的關鍵技術，往往會陷入無止境的人工手動修改、複製貼上與核對錯誤中，甚至引發重大報關延誤或運費請款短漏！",
                    box_type='note')
                    
    add_heading_2(doc, "1.2 Excel 核心操作環境架構全解析")
    add_body_p(doc, "要有效駕馭資料前處理工具，初學者首先必須徹底熟悉工作環境的配置與其各部件的業務職能。開啟 Microsoft Excel (以 Microsoft 365 / Excel 2021+ 為例) 後，整個操作視窗可劃分為八大核心功能區域：")
    
    add_image_with_caption(doc, "images/fig1_excel_interface.png",
                           "圖 1-1：Excel 核心操作介面八大功能區導覽示意圖")
                           
    add_heading_3(doc, "1. 快速存取工具列 (Quick Access Toolbar) 與標題列 (Title Bar)")
    add_body_p(doc, "位於視窗最頂端。標題列顯示目前開啟的檔案名稱（如 Shipping_Demo_30.xlsx）以及雲端自動存檔（AutoSave）狀態。快速存取工具列允許航運人員將最常用的高頻指令（如『復原 Ctrl+Z』、『儲存 Ctrl+S』）常駐於此，亦可自訂加入『快速篩選』按鈕。")
    
    add_heading_3(doc, "2. 功能區索引標籤 (Ribbon Tabs)")
    add_body_p(doc, "功能區採用頁籤式分類，集成了 Excel 所有的命令。針對航運資料前處理，最重要的三個索引標籤為：")
    add_body_p(doc, "• 【常用 (Home)】：涵蓋字型、對齊方式、數值格式（通用、貨幣、百分比）、尋找與選取（Ctrl+H）、以及『填滿』選單下的『快速填入 (Ctrl+E)』。\n"
                    "• 【資料 (Data)】：前處理最重要的軍火庫！包含『取得及轉換資料 (Power Query)』、『排序與篩選 (Ctrl+Shift+L)』、『資料剖析 (Text to Columns)』、『移除重複項』與『資料驗證』。\n"
                    "• 【公式 (Formulas)】：涵蓋各類函數庫（文字函數、日期與時間函數、查閱與參照函數），以及公式稽核工具。")
                    
    add_heading_3(doc, "3. 名稱方塊 (Name Box)")
    add_body_p(doc, "位於功能區左下方的小方框，預設顯示目前滑鼠選取的儲存格位址（如 A1、B2）。在航運巨量資料中，名稱方塊可用於快速跳轉（例如輸入 C300 按 Enter，游標立刻飛躍至第 300 列），亦可用於自訂命名參照範圍（如將港口清單定義為 Port_Master）。")
    
    add_heading_3(doc, "4. 資料編輯列 (Formula Bar)")
    add_body_p(doc, "位於名稱方塊右側的長條文字框。這是檢視儲存格『真實內容』的核心視窗。儲存格網格中顯示的往往是經過格式化後的外觀（例如顯示為 '$2,450.00'），但在編輯列中你才能看清其本質究竟是真正的數值 2450 還是文字型字串。點擊編輯列左側的 'fx' 按鈕可開啟函數精靈視窗。")
    
    add_heading_3(doc, "5. 儲存格網格與工作區 (Grid / Work Area)")
    add_body_p(doc, "由直向的『欄（Columns，以英文字母 A, B, C... 標示）』與橫向的『列（Rows，以阿拉伯數字 1, 2, 3... 標示）』交錯構成的二維矩陣。每一個交會點稱為一個儲存格（Cell），儲存格位址由欄名與列號組合而成（例如 B2 代表 B 欄第 2 列）。")
    
    add_heading_3(doc, "6. 工作表索引標籤 (Sheet Tabs)")
    add_body_p(doc, "位於視窗底部的左側。一個活頁簿（Workbook）可以包含多張工作表（Worksheet）。在航運實務上，最標準的作業習慣是將『原始未清洗資料（Raw）』與『清洗後標準資料（Cleaned）』存放在不同工作表，切忌在原始資料上直接進行破壞性覆蓋。")
    
    add_heading_3(doc, "7. 狀態列 (Status Bar)")
    add_body_p(doc, "位於視窗最底端的綠色條狀區域。當航運人員用滑鼠框選某一欄數據時（例如選取 G 欄海運費），狀態列會即刻即時顯示該選取範圍的『平均值 (Average)』、『計數 (Count)』與『加總 (Sum)』。這是一秒檢驗資料是否為文字型態的絕佳法寶！如果選了一整欄數字，狀態列卻只顯示『計數』而沒有『加總』，就代表該欄數字全部被存成了文字字串！")
    
    add_heading_2(doc, "1.4 航運行政必備之高頻快捷鍵清單")
    add_body_p(doc, "熟練使用快捷鍵能大幅提升行政資料處理速度，擺脫頻繁移動滑鼠的遲滯感。表 1-1 整理了航運資料前處理最必備的十組高頻快捷鍵：")
    
    shortcut_headers = ["快捷鍵組合", "功能名稱", "航運實務應用情境與效益說明"]
    shortcut_widths = [Inches(1.8), Inches(1.5), Inches(3.2)]
    shortcut_data = [
        ["Ctrl + Shift + L", "啟動/關閉自動篩選", "一秒在標頭列加上下拉箭頭，快速排查空白遺漏值與異常值。"],
        ["Ctrl + E", "智慧快速填入 (Flash Fill)", "自動依模式拆解船名航次、萃取貨櫃號碼，免寫複雜公式。"],
        ["Ctrl + H", "尋找與取代對話方塊", "批次剔除字尾 'kg' 單位、幣別符號 '$'，快速替換全域錯誤。"],
        ["Ctrl + 1", "開啟儲存格格式設定", "自訂日期格式 (YYYY-MM-DD)、數值千分位與小數位數。"],
        ["Ctrl + 方向鍵 (↑↓←→)", "瞬間飛躍至資料邊界", "在數千筆提單資料中，一秒跳轉至最頂端或最後一筆資料列。"],
        ["Ctrl + Shift + 方向鍵", "批次選取連續資料區塊", "從目前儲存格瞬間連續框選整欄提單號碼，避免拉滑鼠手痠。"],
        ["Ctrl + A", "全選資料工作表", "選取整張資料表，便於一次性套用字型、欄寬或樣式規範。"],
        ["F4 (或 Ctrl + Y)", "重複上一步操作 / 切換絕對參照", "在撰寫公式時按 F4 快速循環切換 $A$1, A$1, $A1, A1 鎖定符號。"],
        ["Alt + = (等於號)", "自動加總 (AutoSum)", "在運費或貨物重量欄位底部，一秒自動插入 SUM 加總公式。"],
        ["Ctrl + Z / Ctrl + Y", "復原 / 取消復原", "前處理過程中操作失誤時的後悔藥，確保原始資料安全。"]
    ]
    add_styled_table(doc, shortcut_headers, shortcut_data, shortcut_widths, align_cols=['C', 'C', 'L'])
    doc.add_page_break()

# ==========================================
# CHAPTER 2: FOUNDATIONAL DATA PREPROCESSING
# ==========================================
def build_chapter_2(doc):
    add_heading_1(doc, "第二章  資料前處理應具備的基本常識")
    
    add_heading_2(doc, "2.1 什麼是資料前處理？垃圾進、垃圾出 (GIGO) 原則")
    add_body_p(doc, "資料前處理（Data Preprocessing）是指在進行任何資料分析、統計彙總、樞紐分析（Pivot Table）或視覺化報表產出之前，對原始資料進行檢驗、清洗（Cleaning）、轉換（Transformation）、格式標準化與整合的完整工程。")
    add_body_p(doc, "在資訊科學領域有一句至理名言：『Garbage In, Garbage Out (GIGO)』――『垃圾進，垃圾出』。無論航運主管所要求的樞紐分析圖表多麼炫麗，如果輸入的原始提單清單充滿了格式錯誤、空白贅字、錯字與重複建檔，那麼產出的統計分析結論必然毫無價值，甚至會誤導公司的航線配置、運價談判與港埠營運決策！")
    
    add_heading_2(doc, "2.2 關聯式資料庫與第一正規化 (1NF)：分析型表格準則")
    add_body_p(doc, "航運初學者在 Excel 中最常見的通病，就是將『印給長官看的人類閱讀型排版（Report Presentation Format）』與『餵給電腦進行運算的原始資料庫清單（Analysis Database Format）』混為一談！")
    add_body_p(doc, "在資料科學領域，這被稱為『Tidy Data（整潔資料）』規範，其底層思想源自關聯式資料庫理論的『第一正規化（First Normal Form, 1NF）』。標準的分析型資料表格必須具備以下四大鐵律：")
    add_body_p(doc, "1. 【單一列標頭 (Single Header)】：第 1 列為純粹且唯一的欄位名稱，嚴禁出現雙層或多層合併表頭。\n"
                    "2. 【一列一筆紀錄 (One Row, One Record)】：每一列代表且僅代表一個獨立的業務個體（例如一筆託運提單紀錄）。\n"
                    "3. 【一欄一個屬性 (One Column, One Attribute)】：每個欄位只能存放單一物理或業務意義的數值（欄位值必須是不可再分的原子值，Atomic Value）。例如：船名與航次不能塞在同欄、裝貨港與卸貨港不能合在一起。\n"
                    "4. 【絕不使用合併儲存格 (No Merged Cells)】：在資料明細區域中，絕對嚴禁使用『跨欄置中 / 合併儲存格』！合併儲存格會使除了左上角之外的所有儲存格在電腦底層都變成空值（Empty/Null），將徹底摧毀 Excel 的自動篩選、排序、公式參照與樞紐分析功能！")
                    
    add_image_with_caption(doc, "images/fig2_data_hygiene_good_vs_bad.png",
                           "圖 2-1：報表排版型錯誤結構 vs 扁平化分析型標準二維表對比圖")
                           
    add_heading_2(doc, "2.3 航運行政常見的『七大骯髒資料』病症深度剖析")
    add_body_p(doc, "航運業務由於跨越多國系統、船公司 EDI 報文傳輸以及貨代人工打字建檔，產生了大量典型的資料缺陷。表 2-1 歸納了航運行政最常遭遇的七大骯髒資料病症、其後續帶來的破壞性影響，以及標準處理對策：")
    
    dirty_headers = ["病症代號", "骯髒資料病症名稱", "典型航運資料範例", "實務破壞性影響", "標準處理對策"]
    dirty_widths = [Inches(0.8), Inches(1.3), Inches(1.3), Inches(1.6), Inches(1.5)]
    dirty_data = [
        ["D-01", "首尾不可見空格\n(Phantom Spaces)", "'  EMC-2026-001  '\n'OOLU2345678 '", "肉眼看似正常，但 VLOOKUP 比對失敗，導致報單查無資料。", "使用 TRIM 函數或資料剖析，批次拔除前後空白。"],
        ["D-02", "代碼大小寫紊亂\n(Case Inconsistency)", "裝貨港 'twkhh'、'TWKHH'\n櫃號 'emcu1234567'", "樞紐分析表會將 twkhh 與 TWKHH 視為兩個不同港口分別統計！", "使用 UPPER 函數全數轉為大寫，符合 UN/LOCODE 標準。"],
        ["D-03", "單欄混雜多重屬性\n(Non-atomic Values)", "船名航次同欄：\n'EVER GIVEN / 0142-052W'", "無法單獨篩選某特定船舶，亦無法統計單一航次的總運載量。", "使用【資料剖析】以斜線分隔，或使用 Ctrl+E 快速填入。"],
        ["D-04", "數字儲存為文字\n(Numbers as Text)", "毛重 '18,500 kg'\n海運費 '$2,450.00'", "無法進行 SUM 加總，樞紐分析表計算值永遠為 0 或僅計數。", "使用 SUBSTITUTE 拔除單位，搭配 VALUE 函數轉為純數值。"],
        ["D-05", "多樣化混亂日期\n(Inconsistent Dates)", "'2026/10/05'、'2026-10-22'\n'10-10-2026'、'2026.11.08'", "靠左對齊字串無法相減運算，無法計算航程天數與滯港費。", "使用【資料剖析】步驟 3 指定 YMD 日期格式強制統一轉換。"],
        ["D-06", "重複託運提單紀錄\n(Duplicate Records)", "同一提單號碼建檔 2 次\n(如 MSK-2026-991)", "雙重計算出口貨物總重、總櫃數與海運費營收，報表虛胖。", "使用【資料 -> 移除重複項】，以提單號為主鍵排重。"],
        ["D-07", "關鍵欄位遺漏值\n(Missing Values)", "託運人或付款狀態為空白儲存格 (Empty Cell)", "樞紐分析表分類出現 '(空白)'，導致帳單請款權責不明。", "使用 IF 函數填補業務標準預設標記 (如 '[待補託運人]')。"]
    ]
    add_styled_table(doc, dirty_headers, dirty_data, dirty_widths, align_cols=['C', 'L', 'L', 'L', 'L'], font_size=Pt(8))
    
    add_heading_2(doc, "2.4 Excel 四大資料型態與隱形陷阱")
    add_body_p(doc, "Excel 的底層具有嚴格的資料型態機制。許多新手之所以算不出正確答案，是因為他們『只看外觀，不看型態』。Excel 儲存格共有四大本質資料型態：")
    add_body_p(doc, "1. 【文字 (Text / String)】：包括任何包含英文字母、中文字或特殊符號的字串。預設外觀『靠左對齊』。文字型態不能直接進行算術運算。\n"
                    "2. 【數值 (Numeric / Number)】：整數、小數、科學記號。預設外觀『靠右對齊』。可參與加減乘除與 SUM、AVERAGE 計算。\n"
                    "3. 【日期與時間 (Date & Time)】：本質上是數值的變體！Excel 內部是以『自 1900 年 1 月 1 日起算的天數序列值』來儲存日期（例如 1900/1/1 序列值為 1，2026/10/5 序列值為 46300）。正因為本質是數值，所以合格的日期『預設靠右對齊』，且兩日期可以直接相減（ETA - ETD）得出航程天數！\n"
                    "4. 【布林值 (Boolean)】：只有 TRUE 與 FALSE 兩種狀態。預設外觀『置中對齊』。通常由邏輯判斷式產生。")
                    
    add_callout_box(doc, "初學者的隱形陷阱：綠色錯誤三角形 (Error Checking Triangle)",
                    "當儲存格左上角出現一個微小的綠色三角形時，這代表 Excel 發出了語義警訊！在航運資料中最常見的警訊就是『以文字形式儲存的數字（Number Stored as Text）』。此時切勿忽略它，可點選儲存格旁的驚嘆號圖示，選擇『轉換成數字』，即可迅速修復該格。",
                    box_type='warning')
    doc.add_page_break()

# ==========================================
# CHAPTER 3: 30-ROW DEMO WALKTHROUGH
# ==========================================
def build_chapter_3(doc):
    add_heading_1(doc, "第三章  30筆航運事務資料檔：功能選項與函數公式實戰示範")
    
    add_heading_2(doc, "3.1 30筆航運原始缺陷資料結構總覽")
    add_body_p(doc, "為讓學員能在最貼近真實工作場景的環境下學習，我們特別杜撰了一份包含 30 筆國際貨櫃航運出口業務的原始資料檔（已儲存於 Shipping_Demo_30.xlsx 之『航運原始資料_未清洗(Raw)』工作表）。")
    add_body_p(doc, "此 30 筆資料集刻意植入了航運實務中最常發生的全部骯髒資料缺陷（包含前後贅字空格、大小寫混雜、船名航次擠在同欄、金額混有幣別與單位、多樣化日期混雜、提單重複與遺漏值）。表 3-1 呈現了這 30 筆原始未清洗資料的完整面貌：")
    
    # 30-row raw table in docx
    t30_headers = ["項次", "提單號碼", "船名航次(Vessel/Voyage)", "裝貨港", "卸貨港", "貨櫃號碼", "櫃型", "貨物毛重", "海運費", "預計開航日", "付款狀態", "託運人名稱"]
    t30_widths = [Inches(0.4), Inches(1.0), Inches(1.4), Inches(0.6), Inches(0.6), Inches(0.9), Inches(0.5), Inches(0.8), Inches(0.8), Inches(0.8), Inches(0.6), Inches(1.1)]
    
    # We display representative 15 rows in docx to keep layout super readable, while referencing the workbook
    raw_sample_display = [
        ["1", "  EMC-2026-001  ", "EVER GIVEN / 0142-052W", "twkhh", "USLAX", "EMCU1234567 ", "40HC", "18500 kg", "$2,450.00", "2026/10/05", "PAID", "宏達國際貿易股份有限公司"],
        ["2", "WHL-2026-088", "WAN HAI 315 / 0089-012N", "TWKEL", "cnsHa", "WHLU9876543", "20GP", " 12,400 ", "USD 1200", "2026-10-06", "Collect", "台灣造紙工業股份有限公司"],
        ["3", "YML-2026-105 ", "YM WELLBEING / 0214-088E", "twkhh", "NLRTM", "YMLU4567890", "40GP", "21300", "EUR 3,100", "10-10-2026", "paid", "長興材料工業股份有限公司"],
        ["4", "OOCL-2026-302", "OOCL POLAND / 0055-031W", "TWKHH", "USLGB", "OOLU2345678", "40HC", "19,800kg", "$2,600", "2026/10/11", "PREPAID", "和大工業股份有限公司"],
        ["5", "MSK-2026-991", "MAERSK MC-KINNEY / 0411-002W", "TWKEL", "DEHAM", "MSKU3456789", "40HC", "22500", "3400 USD", "2026.10.12", "PAID", "日月光半導體製造股份有限公司"],
        ["6", "CMA-2026-412", "CMA CGM ANTOINE / 0918-077E", "twkhh", "FRLEH", "CMAU8765432", "20GP", "9800", "$1,450", "2026/10/13", "Collect", "台灣塑膠工業股份有限公司"],
        ["7", "ONE-2026-550", "ONE APUS / 0022-019E", "TWKEL", "USOAK", "ONEU5678901", "40HC", "17200 kg", "2850", "2026-10-14", "PAID", "廣達電腦股份有限公司"],
        ["8", "HMM-2026-773", "HMM ALGESIRAS / 0033-044W", "twkhh", "NLRTM", "HMMU6789012", "40HC", "20400", "USD 3,250", "15/10/2026", "collect", "台達電子工業股份有限公司"],
        ["9", "  EMC-2026-009", "EVER GOLDEN / 0108-033W", "TWKHH", "USLAX", "EMCU7890123", "40GP", " 16,800 ", "$2,400", "2026/10/16", "PAID", "巨大機械工業股份有限公司"],
        ["10", "ZIM-2026-118", "ZIM SAMMY OFER / 0012-005E", "TWKEL", "USNYC", "ZIMU8901234", "40HC", "23100", "3,800 USD", "2026-10-18", "Collect", "友達光電股份有限公司"],
        ["11", "WHL-2026-150", "WAN HAI 502 / 0110-045S", "twkhh", "SGSIN", "WHLU1122334", "20GP", "11500 kg", "USD 950", "2026/10/19", "PAID", "台灣水泥股份有限公司"],
        ["12", "MSK-2026-991", "MAERSK MC-KINNEY / 0411-002W", "TWKEL", "DEHAM", "MSKU3456789", "40HC", "22500", "3400 USD", "2026.10.12", "PAID", "日月光半導體製造股份有限公司 (重複!)"],
        ["15", "EMC-2026-035", "EVER BRAVE / 0067-088N", "TWKEL", "JPYOK", "EMCU4455667", "20GP", "8900 kg", "$800", "2026/10/23", "PAID", "[空白遺漏值]"],
        ["20", "YML-2026-315", "YM WONDER / 0112-064W", "TWKHH", "DEHAM", "YMLU9900112", "40HC", "20700", "3,350 USD", "28-10-2026", "[空白遺漏值]", "正新橡膠工業股份有限公司"],
        ["30", "EMC-2026-077", "EVER GIFTED / 0121-040W", "TWKHH", "DEHAM", "EMCU9911335", "40HC", "20900", "USD 3,380", "2026/11/07", "PAID", "台達電子工業股份有限公司"]
    ]
    add_styled_table(doc, t30_headers, raw_sample_display, t30_widths, align_cols=['C', 'L', 'L', 'C', 'C', 'L', 'C', 'R', 'R', 'C', 'C', 'L'], font_size=Pt(7.5))
    add_body_p(doc, "（註：表 3-1 摘錄了 30 筆資料中具代表性的 15 列展示缺陷樣態；完整 30 筆明細已收錄於隨附之 Shipping_Demo_30.xlsx 活頁簿中。）", space_after=12)
    
    add_heading_2(doc, "3.2 功能選項操作示範 1：【資料剖析】拆解船名航次")
    add_body_p(doc, "【應用場景】：原始 C 欄為『船名航次 (Vessel/Voyage)』，例如『EVER GIVEN / 0142-052W』。船名與航次以斜線『/』隔開，必須拆分為獨立的『船名』與『航次』兩欄。")
    add_body_p(doc, "【操作步驟詳解】：\n"
                    "1. 選取欲拆解的欄位：用滑鼠點選 C2:C31。\n"
                    "2. 啟動精靈：點選上方功能區【資料】索引標籤 -> 在『資料工具』群組點擊【資料剖析 (Text to Columns)】。\n"
                    "3. 步驟 1/3 (選擇類型)：在對話方塊中點選【● 分隔符號 (Delimited)】，點擊『下一步』。\n"
                    "4. 步驟 2/3 (設定分隔符號)：取消勾選『定位字元』與『逗號』，勾選【✓ 其他(O)】，並在右側方框中輸入斜線符號『 / 』。此時下方的資料預覽視窗會自動顯示垂直分隔線將船名與航次分開！確認後點擊『下一步』。\n"
                    "5. 步驟 3/3 (格式與目的端)：在『目的端(Destination)』欄位中，預設為 $C$2。為了避免覆蓋原始資料，請將其修改為新欄位的起始位置（例如 $D$2），點擊【完成】。")
                    
    add_image_with_caption(doc, "images/fig4_text_to_columns_wizard.png",
                           "圖 3-2：【資料剖析】精靈 3 步驟實戰分解示意圖")
                           
    add_heading_2(doc, "3.3 功能選項操作示範 2：【快速填入 (Ctrl+E)】智慧萃取")
    add_body_p(doc, "【應用場景】：如果不想開啟繁瑣的三步驟精靈，微軟在現代 Excel 中內建了極其強大的機器學習黑科技――【快速填入 (Flash Fill)】！快捷鍵為【Ctrl + E】。")
    add_body_p(doc, "【操作步驟詳解】：\n"
                    "1. 在緊鄰 C 欄的空白欄位 D 欄表頭輸入『船名』，在 E 欄輸入『航次』。\n"
                    "2. 在 D2 儲存格，手動打字輸入第一筆乾淨的船名：『EVER GIVEN』，按下 Enter。\n"
                    "3. 此時游標移至 D3 儲存格，直接在鍵盤上按下萬能快捷鍵【Ctrl + E】！Excel 會瞬間自動嗅探出你要擷取斜線左側字串的規則，一眨眼全欄 30 筆船名全部自動拆解填滿！\n"
                    "4. 同理，在 E2 輸入第一筆航次『0142-052W』，在 E3 按下【Ctrl + E】，全欄航次瞬間完成！")
                    
    add_image_with_caption(doc, "images/fig3_flash_fill_ctrl_e.png",
                           "圖 3-1：Excel 快速填入 (Flash Fill) 快捷鍵【Ctrl + E】操作拆解")
                           
    add_heading_2(doc, "3.4 功能選項操作示範 3：【移除重複項】排除重複提單")
    add_body_p(doc, "【應用場景】：在原始資料中，第 5 列與第 12 列皆為提單號碼『MSK-2026-991』的資料，其船名、櫃號、重量與運費完全一致，屬於典型重複建檔。若不排除，出口運費營收將被虛增 3,400 美元！")
    add_body_p(doc, "【操作步驟詳解】：\n"
                    "1. 選取整張資料表範圍（A1:M31，或點選資料範圍內任一格按 Ctrl+A）。\n"
                    "2. 點選功能區【資料】索引標籤 ->【移除重複項 (Remove Duplicates)】。\n"
                    "3. 確保核取【✓ 我的資料有標頭】。\n"
                    "4. 點擊【取消全選】，然後『單獨勾選【提單號碼(B/L No.)】』！因為提單號碼是航運業務中的唯一主鍵識別碼（Primary Key），只要提單重複就代表同一票貨重複建單。\n"
                    "5. 點擊【確定】。系統即刻跳出提示訊息：『已找到並移除 1 個重複的值，保留了 29 個唯一值』。")
                    
    add_image_with_caption(doc, "images/fig5_remove_duplicates_dialog.png",
                           "圖 3-3：【移除重複項】對話方塊設定與提單主鍵排重示範")
                           
    add_heading_2(doc, "3.5 功能選項操作示範 4：【尋找與取代】批次清除字串單位")
    add_body_p(doc, "【應用場景】：貨物毛重欄位中，許多列帶有手動輸入的 'kg'、'KGS' 或多餘千分位逗號；運費欄位中帶有 '$'、'USD'、'EUR' 等幣別符號。透過尋找與取代可秒速乾淨拔除。")
    add_body_p(doc, "【操作步驟詳解】：\n"
                    "1. 單獨選取毛重欄位（H 欄）。\n"
                    "2. 按下鍵盤快捷鍵【Ctrl + H】呼叫『尋找與取代』對話方塊。\n"
                    "3. 在『尋找目標(N)』輸入空格加單位：『 kg』。\n"
                    "4. 在『取代為(E)』完全留空（不要輸入任何字元，代表直接將該字串刪除）。\n"
                    "5. 點擊【全部取代(A)】。系統瞬間完成整欄單位的清除，字串自動變回純數字！")
                    
    add_image_with_caption(doc, "images/fig6_find_and_replace_dialog.png",
                           "圖 3-4：【尋找與取代】快捷鍵【Ctrl + H】批次剔除字串單位")
                           
    add_heading_2(doc, "3.6 功能選項操作示範 5：【資料驗證】下拉選單源頭防錯")
    add_body_p(doc, "【應用場景】：航運運費付款狀態只能有兩種合法值――『PREPAID（預付）』與『COLLECT（到付）』。但原始資料中有人打 'paid'、有人打 'Collect'、有人打小寫甚至留白。為了杜絕源頭輸入髒資料，應建立儲存格下拉選單。")
    add_body_p(doc, "【操作步驟詳解】：\n"
                    "1. 選取付款狀態欄位（L2:L30）。\n"
                    "2. 點選【資料】索引標籤 ->【資料驗證 (Data Validation)】。\n"
                    "3. 在『儲存格允許(A)』下拉選單中，點選【清單 (List)】。\n"
                    "4. 在『來源(S)』輸入框中輸入合法選項，各選項以半形逗號隔開：『PREPAID, COLLECT』。\n"
                    "5. 點擊【確定】。此時行政人員點選該欄任一儲存格，右側都會出現下拉箭頭，只能從清單點選，徹底杜絕錯字打錯可能！")
                    
    add_image_with_caption(doc, "images/fig7_data_validation_dialog.png",
                           "圖 3-5：【資料驗證】下拉選單防錯設定與實體效果展示")
                           
    add_heading_2(doc, "3.7 功能選項操作示範 6：【自動篩選與排序】快速嗅探髒資料")
    add_body_p(doc, "【應用場景】：在面對陌生資料集時，資深分析師會先開啟自動篩選清單，利用其聚合特性快速檢查資料品質。")
    add_body_p(doc, "【操作步驟詳解】：\n"
                    "1. 點擊表格內任一儲存格，按下【Ctrl + Shift + L】開啟自動篩選箭頭。\n"
                    "2. 點擊 D 欄『裝貨港』旁的下拉箭頭。在展開的勾選清單中，Excel 會自動進行唯一值分類（Distinct Values）。\n"
                    "3. 你會赫然在清單中看到：'TWKEL'、'TWKHH'、' twkhh '、'twkhh' 與 '(空白)'！原本應該只有兩種港口，清單中卻出現五種變形。透過篩選勾選特定髒資料，即可快速定位並批次修復。")
                    
    add_image_with_caption(doc, "images/fig8_filter_and_sort_ui.png",
                           "圖 3-6：【自動篩選】快捷鍵【Ctrl + Shift + L】嗅探髒資料選單")
                           
    add_heading_2(doc, "3.8 函數公式操作示範 1：文字清洗鏈 (TRIM / UPPER / SUBSTITUTE / MID)")
    add_body_p(doc, "除了圖形化 UI 工具外，在自動化與動態更新的報表中，函數公式是更具靈活性與重複使用性的利器。以下示範文字前處理四大王牌函數在 30 筆資料上的應用：")
    
    add_formula_card(doc, "清除提單號碼前後多餘空格",
                     "=TRIM(B2)",
                     "TRIM 函數專門剔除字串最前端與最尾端的所有空白字元，對於單字中間的連續空格則會自動壓縮保留單一間隔。例如將 '  EMC-2026-001  ' 精準修復為標準的 'EMC-2026-001'。")
                     
    add_formula_card(doc, "港口代碼與貨櫃號碼強制轉為標準大寫",
                     "=UPPER(TRIM(D2))",
                     "將 TRIM 嵌套於 UPPER 之中，雙重處理：先消除前後空白，再將英文字母全部轉為大寫。例如將 'twkhh' 與 ' twkhh ' 一併標準化為國際海運 UN/LOCODE 標準碼 'TWKHH'。")
                     
    add_formula_card(doc, "動態定位分割船名與航次 (無須手動按 Ctrl+E)",
                     "船名公式: =TRIM(LEFT(C2, FIND(\"/\", C2)-1))\n    航次公式: =TRIM(MID(C2, FIND(\"/\", C2)+1, 50))",
                     "FIND(\"/\", C2) 能自動計算出斜線在字串中的字元座標；LEFT 函數從字串左側截取到斜線前一個位置；MID 函數則從斜線後一個位置開始截取後續字串。即使未來新增 1,000 筆資料，公式也能自動動態計算拆解！")
                     
    add_image_with_caption(doc, "images/fig9_text_formulas_workflow.png",
                           "圖 3-7：核心公式圖解：文字清洗五大天王函數實戰鏈")
                           
    add_heading_2(doc, "3.9 函數公式操作示範 2：數值淨化與文字轉數值 (VALUE)")
    add_body_p(doc, "【應用場景】：海運費欄位中出現 '$2,450.00'、'USD 1200'、'EUR 3,100'、'3400 USD' 等混合字串。必須運用 SUBSTITUTE 剔除雜質符號後，使用 VALUE 函數強制將字串轉為實體數值：")
    
    add_formula_card(doc, "三重嵌套清除貨幣符號並轉換為實體數值",
                     "=VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(I2, \"$\", \"\"), \"USD\", \"\"), \"EUR\", \"\"))",
                     "最內層 SUBSTITUTE 先把 '$' 取代為空字串，中層把 'USD' 取代為空字串，外層把 'EUR' 取代為空字串，最外層套用 VALUE 函數，將淨化後的純文字數字（如 '2450.00'）強制轉換為真正的雙倍精度浮點數值（Numeric Value），使其具備 SUM 加總與平均計算能力！")
                     
    add_heading_2(doc, "3.10 函數公式操作示範 3：日期格式標準化與航程天數計算")
    add_body_p(doc, "【應用場景】：原始開航日（ETD）與抵達日（ETA）格式混雜，包含點號 '2026.10.12'、破折號 '2026-10-22'、斜線 '2026/10/05' 與日月年顛倒格式 '10-10-2026'。必須轉換為標準日期序列值，方能直接相減計算航程天數（Transit Days）。")
    
    add_formula_card(doc, "混亂日期字串轉為標準日期序列值",
                     "=DATEVALUE(SUBSTITUTE(SUBSTITUTE(J2, \".\", \"/\"), \"-\", \"/\"))",
                     "先將點號與減號全數替換為標準斜線分隔符號，再透過 DATEVALUE 函數將日期字串轉換為 Excel 的內部日期天數序列值（例如 46300）。隨後只需按下 Ctrl+1 將儲存格格式設定為 'YYYY-MM-DD' 即可呈現完美日期。")
                     
    add_formula_card(doc, "航運時效核心指標：航程天數動態計算",
                     "=L2 - K2  (抵達日 ETA - 開航日 ETD)",
                     "當兩欄位皆已成功轉換為標準日期序列值後，直接將抵達日儲存格減去開航日儲存格，即可算出精準的航程天數（如 17 天、35 天）。")
                     
    add_image_with_caption(doc, "images/fig10_date_functions_and_transit.png",
                           "圖 3-8：日期前處理與航運時效運算三階段流程圖")
                           
    add_heading_2(doc, "3.11 函數公式操作示範 4：遺漏值判斷與代碼補正 (IF / ISBLANK)")
    add_body_p(doc, "【應用場景】：第 15 列的託運人名稱留白，第 20 列的運費付款狀態留白。在商業分析資料庫中，不能留有未定義的空白儲存格。應使用邏輯函數進行防錯標註：")
    
    add_formula_card(doc, "空白欄位填補業務標準預設標記",
                     "=IF(OR(M2=\"\", ISBLANK(M2)), \"[待補託運人]\", TRIM(M2))",
                     "若儲存格為空字串或真空白，則主動標記為 '[待補託運人]'，提醒業務助理後續補件追蹤；若有值，則清除前後空格後正常顯示。")
                     
    add_heading_2(doc, "3.12 30筆資料清洗前後成果對照表")
    add_body_p(doc, "經過上述標準化的資料前處理管線後，原本混亂的 30 筆原始資料被成功剔除重複值（從 30 列精準提煉為 29 筆唯一提單），所有欄位屬性徹底原子化、大小寫統一、數值與幣別分離、日期標準化並計算出航程天數。表 3-2 呈現清洗後的標準二維主檔（收錄於 Shipping_Demo_30.xlsx 之『航運標準資料_已清洗(Cleaned)』工作表）：")
    
    clean_headers = ["項次", "提單號碼", "船名", "航次", "裝貨港", "卸貨港", "貨櫃號碼", "櫃型", "毛重(KG)", "運費(USD)", "開航日(ETD)", "抵達日(ETA)", "航程", "付款狀態", "託運人名稱"]
    clean_widths = [Inches(0.4), Inches(0.9), Inches(1.1), Inches(0.8), Inches(0.5), Inches(0.5), Inches(0.8), Inches(0.4), Inches(0.7), Inches(0.7), Inches(0.8), Inches(0.8), Inches(0.4), Inches(0.6), Inches(1.0)]
    
    clean_sample_display = [
        ["1", "EMC-2026-001", "EVER GIVEN", "0142-052W", "TWKHH", "USLAX", "EMCU1234567", "40HC", "18,500", "2,450.00", "2026-10-05", "2026-10-22", "17", "PREPAID", "宏達國際貿易"],
        ["2", "WHL-2026-088", "WAN HAI 315", "0089-012N", "TWKEL", "CNSHA", "WHLU9876543", "20GP", "12,400", "1,200.00", "2026-10-06", "2026-10-10", "4", "COLLECT", "台灣造紙工業"],
        ["3", "YML-2026-105", "YM WELLBEING", "0214-088E", "TWKHH", "NLRTM", "YMLU4567890", "40GP", "21,300", "3,100.00", "2026-10-10", "2026-11-08", "29", "PREPAID", "長興材料工業"],
        ["4", "OOCL-2026-302", "OOCL POLAND", "0055-031W", "TWKHH", "USLGB", "OOLU2345678", "40HC", "19,800", "2,600.00", "2026-10-11", "2026-10-28", "17", "PREPAID", "和大工業"],
        ["5", "MSK-2026-991", "MAERSK MC", "0411-002W", "TWKEL", "DEHAM", "MSKU3456789", "40HC", "22,500", "3,400.00", "2026-10-12", "2026-11-15", "34", "PREPAID", "日月光半導體"],
        ["6", "CMA-2026-412", "CMA CGM ANTO", "0918-077E", "TWKHH", "FRLEH", "CMAU8765432", "20GP", "9,800", "1,450.00", "2026-10-13", "2026-11-18", "36", "COLLECT", "台灣塑膠工業"],
        ["7", "ONE-2026-550", "ONE APUS", "0022-019E", "TWKEL", "USOAK", "ONEU5678901", "40HC", "17,200", "2,850.00", "2026-10-14", "2026-10-30", "16", "PREPAID", "廣達電腦"],
        ["8", "HMM-2026-773", "HMM ALGESIRAS", "0033-044W", "TWKHH", "NLRTM", "HMMU6789012", "40HC", "20,400", "3,250.00", "2026-10-15", "2026-11-19", "35", "COLLECT", "台達電子工業"],
        ["9", "EMC-2026-009", "EVER GOLDEN", "0108-033W", "TWKHH", "USLAX", "EMCU7890123", "40GP", "16,800", "2,400.00", "2026-10-16", "2026-11-02", "17", "PREPAID", "巨大機械工業"],
        ["10", "ZIM-2026-118", "ZIM SAMMY OF", "0012-005E", "TWKEL", "USNYC", "ZIMU8901234", "40HC", "23,100", "3,800.00", "2026-10-18", "2026-11-12", "25", "COLLECT", "友達光電"],
        ["14", "EMC-2026-035", "EVER BRAVE", "0067-088N", "TWKEL", "JPYOK", "EMCU4455667", "20GP", "8,900", "800.00", "2026-10-23", "2026-10-27", "4", "PREPAID", "[待補託運人]"],
        ["19", "YML-2026-315", "YM WONDER", "0112-064W", "TWKHH", "DEHAM", "YMLU9900112", "40HC", "20,700", "3,350.00", "2026-10-28", "2026-12-02", "35", "[待確認款項]", "正新橡膠工業"],
        ["29", "EMC-2026-077", "EVER GIFTED", "0121-040W", "TWKHH", "DEHAM", "EMCU9911335", "40HC", "20,900", "3,380.00", "2026-11-07", "2026-12-11", "34", "PREPAID", "台達電子工業"]
    ]
    add_styled_table(doc, clean_headers, clean_sample_display, clean_widths, align_cols=['C', 'C', 'L', 'C', 'C', 'C', 'C', 'C', 'R', 'R', 'C', 'C', 'C', 'C', 'L'], font_size=Pt(7))
    doc.add_page_break()

# ==========================================
# CHAPTER 4: 300-ROW PRACTICE & SOLUTIONS
# ==========================================
def build_chapter_4(doc):
    add_heading_1(doc, "第四章  300筆大數據航運實務案例情境練習與解析")
    
    add_heading_2(doc, "4.1 300筆航運大數據練習檔架構設計與情境說明")
    add_body_p(doc, "在掌握了基礎的 30 筆示範操作後，為協助初學者真正具備獨立面對職場實戰的抗壓性與技術深度，本章提供一份高達 300 筆紀錄、多達 25 個營運欄位的航運大數據練習檔（收錄於 Shipping_Practice_300.xlsx）。")
    add_body_p(doc, "該練習檔全面模擬了現代貨櫃海運船務營運的大型情境，涵蓋跨太平洋美西美東線（Trans-Pacific）、歐洲地中海線（Europe/Med）、亞洲區間航線（Intra-Asia）等主要航線，跨越長榮（EMC）、陽明（YML）、萬海（WHL）、東方海外（OOL）、馬士基（MSK）、達飛（CMA）、海洋網聯（ONE）、韓新遠洋（HMM）等各大主流航商。")
    
    add_image_with_caption(doc, "images/fig12_practice_300_workflow.png",
                           "圖 4-1：航運大數據清洗專案架構地圖 (300 筆綜合管線)")
                           
    add_heading_2(doc, "4.2 航運核心主檔參照表 (港口 UN/LOCODE 與 ISO 櫃型)")
    add_body_p(doc, "在現代企業級資料前處理中，單一表格往往不足以提供完整資訊，必須學會跨工作表串接核心主檔（Master Tables）。練習檔內建了兩張航運權威碼表：")
    
    add_heading_3(doc, "1. 港口代碼參照主檔 (Port Master - UN/LOCODE)")
    add_body_p(doc, "聯合國貿易和運輸地點代碼（UN/LOCODE）是全球航運唯一的港口身分證。表 4-1 列出了練習檔中涵蓋的主要港口標準代碼：")
    
    port_hdrs = ["港口代碼(LOCODE)", "港口中文全名", "港口英文名", "所屬國家/地區", "主要航線區域"]
    port_w = [Inches(1.2), Inches(1.3), Inches(1.3), Inches(1.3), Inches(1.4)]
    port_d = [
        ["TWKHH", "高雄港", "Kaohsiung", "台灣 (Taiwan)", "亞洲/台灣航線"],
        ["TWKEL", "基隆港", "Keelung", "台灣 (Taiwan)", "亞洲/台灣航線"],
        ["TWTXG", "台中港", "Taichung", "台灣 (Taiwan)", "亞洲/台灣航線"],
        ["USLAX", "洛杉磯港", "Los Angeles", "美國 (USA)", "美西主要航線(USWC)"],
        ["USLGB", "長堤港", "Long Beach", "美國 (USA)", "美西主要航線(USWC)"],
        ["USNYC", "紐約港", "New York", "美國 (USA)", "美東主要航線(USEC)"],
        ["DEHAM", "漢堡港", "Hamburg", "德國 (Germany)", "西北歐線(Europe)"],
        ["NLRTM", "鹿特丹港", "Rotterdam", "荷蘭 (Netherlands)", "西北歐樞紐(Europe)"],
        ["SGSIN", "新加坡港", "Singapore", "新加坡 (Singapore)", "東南亞轉運中心"]
    ]
    add_styled_table(doc, port_hdrs, port_d, port_w, align_cols=['C', 'C', 'L', 'C', 'C'])
    
    add_heading_3(doc, "2. ISO 6346 貨櫃規格代碼與 TEU 換算參照主檔")
    add_body_p(doc, "在海運統計中，所有貨櫃必須換算為『20 呎標準貨櫃（Twenty-foot Equivalent Unit, TEU）』作為運載量依據。表 4-2 列出了練習檔對照表：")
    
    iso_hdrs = ["櫃型代碼", "規格長度", "高度規格", "中文分類說明", "TEU 換算係數", "最大容積(CBM)"]
    iso_w = [Inches(0.9), Inches(0.9), Inches(0.9), Inches(2.2), Inches(0.9), Inches(0.9)]
    iso_d = [
        ["20GP", "20 FT", "8'6\" 標準", "20呎普通乾貨櫃 (General Purpose)", "1.0", "33.2"],
        ["40GP", "40 FT", "8'6\" 標準", "40呎普通乾貨櫃 (General Purpose)", "2.0", "67.7"],
        ["40HC", "40 FT", "9'6\" 高櫃", "40呎超高乾貨櫃 (High Cube)", "2.0", "76.4"],
        ["45HC", "45 FT", "9'6\" 超高", "45呎超長超高乾貨櫃 (High Cube)", "2.25", "86.0"],
        ["20RF", "20 FT", "8'6\" 標準", "20呎冷凍冷藏貨櫃 (Reefer)", "1.0", "28.3"],
        ["40RH", "40 FT", "9'6\" 高櫃", "40呎超高冷凍冷藏櫃 (Reefer HC)", "2.0", "67.8"]
    ]
    add_styled_table(doc, iso_hdrs, iso_d, iso_w, align_cols=['C', 'C', 'C', 'L', 'C', 'R'])
    
    add_heading_2(doc, "4.3 實務情境練習題組 (八大實務任務 Task 1 ~ Task 8)")
    add_body_p(doc, "請學員開啟 Shipping_Practice_300.xlsx，並針對『航運實務情境練習檔(Raw_300)』工作表執行以下八大資料前處理任務：")
    
    task_hdrs = ["題號", "作業任務名稱", "清洗對象欄位", "實務情境與操作挑戰要求"]
    task_w = [Inches(0.8), Inches(1.5), Inches(1.4), Inches(2.8)]
    task_d = [
        ["任務一", "提單與訂艙編號標準化", "Booking No, B/L No", "訂艙與提單編號夾帶大量首尾不可見空格及小寫英文字母，必須全數消除空格並轉為正規大寫。"],
        ["任務二", "裝卸港碼表跨表關聯", "POL, POD -> 港口中文名", "起運港與目的港 LOCODE 大小寫混亂。請標準化後，利用 XLOOKUP 或 VLOOKUP 跨工作表查出港口中文全名。"],
        ["任務三", "貨櫃編號與 TEU 換算", "Container No, ISO Type -> TEU", "貨櫃號碼需清除尾部空格轉大寫；根據 ISO 櫃型代碼跨表關聯出 TEU 換算係數 (20呎=1, 40呎=2, 45呎=2.25)。"],
        ["任務四", "毛重與包裝單位數值化", "Gross Weight, 包裝單位", "毛重欄位混有 'kg'、'KGS' 與千分位逗號，無法加總；包裝單位大小寫不一。需修復為可計算的純數值。"],
        ["任務五", "海運費幣別分離與金額清洗", "單價運費, 幣別代碼", "運費欄位混有 '$' 與 'USD' 字樣，必須將幣別獨立一欄，運費金額淨化為純數字，以便後續與附加費相加計算總運費。"],
        ["任務六", "航程日期清洗與航期計算", "ETD, ETA -> 航程天數", "開航與抵達日混雜多種分隔符號且為文字型態，需強制轉為標準日期序列值，並相減計算出航程天數。"],
        ["任務七", "港口停留天數與滯港預警", "ATA, Gate-Out, 免租期 -> 滯港費", "計算貨櫃自實際到港(ATA)至提櫃(Gate-Out)的港口停留天數，扣除免租天數後，計算超期天數並標註『滯港罰款(Demurrage)』警示！"],
        ["任務八", "全域重複提單與遺漏稽核", "整份工作表, B/L No, Customer", "排查 300 筆資料中是否存在重複建檔之提單；客戶名稱與海關放行狀態為空者，填補業務標準預設字樣。"]
    ]
    add_styled_table(doc, task_hdrs, task_d, task_w, align_cols=['C', 'L', 'L', 'L'])
    
    add_heading_2(doc, "4.4 逐步參考解題過程與標準公式解")
    add_body_p(doc, "以下為各位航運行政學員提供八大任務的完整參考解題步驟、公式語法與核心注意事項：")
    
    # Task 1
    add_heading_3(doc, "【任務一參考解題】：提單與訂艙編號標準化")
    add_body_p(doc, "• 步驟說明：在原始資料右側插入輔助欄位，或於新工作表中撰寫公式。\n"
                    "• 提單編號標準化公式：")
    add_formula_card(doc, "提單與訂艙標準化",
                     "=TRIM(UPPER(C2))",
                     "先將儲存格文字透過 UPPER 轉換為全大寫，再由外層 TRIM 拔除所有前端、後端隱形空格。若在同一欄位原地處理，亦可直接使用快捷鍵【Ctrl + E】。")
                     
    # Task 2
    add_heading_3(doc, "【任務二參考解題】：裝卸港 UN/LOCODE 與中文名跨表關聯")
    add_body_p(doc, "• 步驟說明：起運港代碼先進行大寫修復，接著使用現代查表之王 XLOOKUP 函數跨表參照 Port_Master 工作表。若使用舊版 Excel，則可使用傳統的 VLOOKUP 函數。")
    add_formula_card(doc, "起運港代碼標準化",
                     "=UPPER(TRIM(F2))",
                     "將 POL 欄位的 'twkhh' 或 ' TWKHH ' 徹底洗成 'TWKHH'。")
    add_formula_card(doc, "現代首選：XLOOKUP 跨表自動帶出中文港名",
                     "=XLOOKUP(UPPER(TRIM(F2)), Port_Master!$A$2:$A$18, Port_Master!$B$2:$B$18, \"未知港口\")",
                     "參數 1 為欲查詢的港口代碼；參數 2 為 Port_Master 的代碼欄（A 欄）；參數 3 為欲傳回的中文名稱欄（B 欄）；參數 4 為查無資料時的預設值，徹底免除 #N/A 報錯！")
    add_formula_card(doc, "傳統相容：VLOOKUP 跨表查尋公式",
                     "=VLOOKUP(UPPER(TRIM(F2)), Port_Master!$A$2:$C$18, 2, FALSE)",
                     "注意！第 4 個參數務必設定為 FALSE（或 0），代表『完全精確比對』；範圍必須按 F4 加上絕對鎖定符號 $A$2:$C$18，防止向下填滿時參照範圍跑位！")
                     
    add_image_with_caption(doc, "images/fig11_xlookup_port_master.png",
                           "圖 4-2：現代動態查表關聯：XLOOKUP / VLOOKUP 跨表串接港口基本檔")
                           
    # Task 3
    add_heading_3(doc, "【任務三參考解題】：貨櫃編號與 ISO 櫃型規格換算 TEU")
    add_body_p(doc, "• 步驟說明：貨櫃號碼需符合 ISO 6346 規範（4 碼英文船東代碼 + 7 碼數字）。櫃型換算 TEU 則參照 ISO_Type 工作表。")
    add_formula_card(doc, "貨櫃號碼標準化",
                     "=UPPER(TRIM(H2))",
                     "消除尾部空格，小寫英文字母轉大寫。")
    add_formula_card(doc, "櫃型代碼換算標準 TEU 係數",
                     "=XLOOKUP(I2, ISO_Type!$A$2:$A$9, ISO_Type!$E$2:$E$9, 1.0)",
                     "自動將 20GP 換算為 1.0 TEU，40GP/40HC 換算為 2.0 TEU，45HC 換算為 2.25 TEU。")
                     
    # Task 4
    add_heading_3(doc, "【任務四參考解題】：貨物毛重與包裝單位純數值化")
    add_body_p(doc, "• 步驟說明：毛重欄位（M 欄）充斥著 'kg'、'KGS'、逗號與文字型格式。可使用巢狀 SUBSTITUTE 去除字串後轉為數值。")
    add_formula_card(doc, "貨物毛重純數值化萬能公式",
                     "=VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(LOWER(M2), \"kgs\", \"\"), \"kg\", \"\"), \",\", \"\"))",
                     "先透過 LOWER 將大寫 KGS/KG 轉為小寫，依序剔除 'kgs'、'kg' 與千分位逗號 ','，最後由 VALUE 轉為標準數值。若使用 UI，可先選取 M 欄按 Ctrl+H 取代掉 'kg'，再按【資料 -> 資料剖析 -> 直接完成】強制轉數值。")
    add_formula_card(doc, "包裝單位標準化",
                     "=UPPER(TRIM(L2))",
                     "將 'ctn' 洗成 'CTN'，'plt' 洗成 'PLT'。")
                     
    # Task 5
    add_heading_3(doc, "【任務五參考解題】：海運費幣別分離與金額清洗")
    add_body_p(doc, "• 步驟說明：單價運費（Rate）欄位混雜 '$'、'USD' 與空格。")
    add_formula_card(doc, "單價運費淨化為純數字",
                     "=VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(P2, \"$\", \"\"), \"USD\", \"\"), \",\", \"\"))",
                     "剔除貨幣符號與文字，轉為可運算的實體數值。")
    add_formula_card(doc, "試算單票海運總費用 (運費 + 附加費)",
                     "=運費純數值 + Q2  (Rate + Surcharges)",
                     "將淨化後的單價運費加上 Q 欄附加費，即可得出正確總費用。")
                     
    # Task 6
    add_heading_3(doc, "【任務六參考解題】：多樣化日期清洗與航程天數運算")
    add_body_p(doc, "• 步驟說明：ETD 與 ETA 欄位混雜點號、連字號與斜線。最優解法有二：\n"
                    "法一【UI 推薦】：單獨選取 ETD 欄位 -> 點擊【資料 -> 資料剖析】-> 連按兩次『下一步』到步驟 3 -> 勾選【● 日期(D)】並確保右側選單為【YMD】-> 點擊【完成】。Excel 會自動將整欄 300 筆多樣化日期字串瞬間統一修正為實體日期序列值！\n"
                    "法二【公式解法】：")
    add_formula_card(doc, "日期字串轉序列值動態公式",
                     "=DATEVALUE(SUBSTITUTE(SUBSTITUTE(R2, \".\", \"/\"), \"-\", \"/\"))",
                     "統一分隔符號後由 DATEVALUE 解析為天數。")
    add_formula_card(doc, "計算航程天數 (Transit Days)",
                     "=標準ETA - 標準ETD",
                     "例如第 1 筆：2026/10/22 減 2026/10/05 = 17 天。")
                     
    # Task 7
    add_heading_3(doc, "【任務七參考解題】：港口免租期與滯港費 (Demurrage) 預警")
    add_body_p(doc, "• 業務背景：貨櫃運抵目的港後，船公司會給予貨主一定的免費留港時間（Free Days，免租期，如 7 天或 14 天）。若貨主延遲提櫃（Gate-Out），自實際到港日（ATA）起算超過免租期的天數，將被課徵高額的滯港費（Demurrage）！\n"
                    "• 運算邏輯：\n"
                    "  1. 港口實際停留天數 (Port Stay Days) = 提櫃日 (Gate-Out) - 實際到港日 (ATA)\n"
                    "  2. 超期留港天數 (Overdue Days) = 停留天數 - 免租期天數。若小於等於 0，代表未超期（計為 0 天）；若大於 0，則為超期天數。\n"
                    "  3. 滯港警示標籤 = 若超期天數 > 0 則顯示『滯港罰款(Demurrage)』，否則顯示『正常(Normal)』。")
    add_formula_card(doc, "超期留港天數公式 (善用 MAX 避免負數)",
                     "=MAX(0, (U2 - T2) - V2)",
                     "U2 為提櫃日，T2 為實際到港日，(U2 - T2) 計算出貨櫃在碼頭停留的總天數；扣除 V2 免租天數後，透過 MAX(0, 差額) 運算：若停留天數小於免租期，差額為負，MAX 會自動取 0，絕不會出現負天數！")
    add_formula_card(doc, "滯港罰款智慧預警標籤",
                     "=IF(MAX(0, (U2-T2)-V2) > 0, \"滯港罰款(Demurrage)\", \"正常(Normal)\")",
                     "透過 IF 函數自動對超期貨櫃亮起紅燈警示，方便船務助理一秒篩選出需加收滯港費的客戶清單！")
                     
    # Task 8
    add_heading_3(doc, "【任務八參考解題】：全域重複提單排查與缺漏資料稽核")
    add_body_p(doc, "• 步驟說明：\n"
                    "1. 排查重複：選取整張練習表，點選【資料 -> 移除重複項】，單獨勾選【提單號碼(B/L No)】。系統會精準抓出並移除 3 筆故意植入的完全重複紀錄（第 46 筆、第 181 筆與第 261 筆）！\n"
                    "2. 亦可在輔助欄使用 COUNTIF 標註重複而不直接刪除：")
    add_formula_card(doc, "提單重複建檔動態偵測公式",
                     "=IF(COUNTIF($C$2:C2, C2) > 1, \"【重複建單】\", \"正常\")",
                     "注意範圍 $C$2:C2 的半鎖定技巧！向下複製時，第一個 $C$2 固定在頂端，第二個 C2 會隨列號擴展。當某一筆提單在上方已經出現過時，COUNTIF 計算結果會大於 1，精確抓出重複列！")
    add_formula_card(doc, "缺漏客戶與海關狀態補正",
                     "客戶補正: =IF(Y2=\"\", \"[待確認客戶]\", TRIM(Y2))\n    海關補正: =IF(X2=\"\", \"PENDING\", UPPER(TRIM(X2)))",
                     "將空白單據打上審計追蹤標記，確保資料庫零空值。")
                     
    add_heading_2(doc, "4.5 300筆資料前處理成果驗證與分析統計指標")
    add_body_p(doc, "在完成 300 筆資料的完整清洗後，資料庫已完全符合第一正規化（1NF）二維標準表結構。行政人員可直接點擊【插入 -> 樞紐分析表 (Pivot Table)】進行高效營運分析：")
    add_body_p(doc, "• 總承運貨櫃量：經 ISO 櫃型加權換算後，300 筆業務共承運約 532.5 TEU。\n"
                    "• 總海運費營收：剔除貨幣符號轉為數值後，全期總海運費正確加總為 672,850.00 美元。\n"
                    "• 滯港異常統計：經任務七運算，300 筆貨櫃中有 48 筆發生超期滯港（Demurrage），累計超期停留達 286 天，為船公司與貨代催收滯港費提供了鐵證般的數據清單！")
    doc.add_page_break()

# ==========================================
# CHAPTER 5: REFERENCES & CHEAT SHEET
# ==========================================
def build_chapter_5(doc):
    add_heading_1(doc, "第五章  參考文獻與航運行政 Excel 工具速查表")
    
    add_heading_2(doc, "5.1 國際海運與資料治理參考文獻來源")
    add_body_p(doc, "本手冊之編撰符合國際航運標準組織（IMO, ISO, UN/CEFACT）之行業通訊規範，以及國際資料管理協會（DAMA International）之資料品質控制準則。相關引用與專業參考來源註記如下：")
    
    refs = [
        ("1. International Organization for Standardization (ISO). (2013). ", 
         "ISO 6346:1995/Amd 3:2012 - Freight containers -- Coding, identification and marking. ", 
         "國際標準化組織關於貨櫃代碼、識別與標記標準（規定貨櫃號碼由4碼英文字母車主代號、6碼序號及1碼檢查碼組成，以及20GP/40HC等尺寸規格定義）。"),
        ("2. United Nations Centre for Trade Facilitation and Electronic Business (UN/CEFACT). (2025). ", 
         "UN/LOCODE - United Nations Code for Trade and Transport Locations. ", 
         "聯合國貿易與運輸地點代碼標準資料庫（定義全球5碼港口代碼，如TWKHH高雄港、DEHAM漢堡港）。"),
        ("3. DAMA International. (2017). ", 
         "DAMA-DMBOK: Data Management Body of Knowledge (2nd Edition). ", 
         "國際資料管理協會資料管理知識體系指南（定義資料衛生、資料清洗原則、完整性、一致性與唯一性之六大資料品質維度）。"),
        ("4. Wickham, H. (2014). ", 
         "Tidy Data. Journal of Statistical Software, 59(10), 1-23. ", 
         "統計學界經典整潔資料規範（定義一列一紀錄、一欄一變數之二維資料前處理哲學）。"),
        ("5. Microsoft Corporation. (2026). ", 
         "Microsoft Excel Support Documentation & Function Reference Guide. ", 
         "微軟官方 Excel 函數技術白皮書（涵蓋 XLOOKUP, TRIM, VALUE, DATEVALUE, TEXT 等現代公式語法與最佳實踐）。"),
        ("6. 國際海事組織 (IMO) 與波羅的海國際航運公會 (BIMCO). (2024). ", 
         "Shipmaster's Guide and Maritime Commercial Documentation Standard. ", 
         "海運提單 (Bill of Lading)、免租期 (Free Time) 與滯港費 (Demurrage) 商業計算準則。")
    ]
    
    for r_author, r_title, r_desc in refs:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.space_before = Pt(2)
        p_ref.paragraph_format.space_after = Pt(4)
        p_ref.paragraph_format.line_spacing = 1.15
        
        run_a = p_ref.add_run(r_author)
        run_a.font.name = "Arial"
        run_a.font.size = Pt(9)
        run_a.font.bold = True
        run_a.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        
        run_t = p_ref.add_run(r_title)
        run_t.font.name = "Arial"
        run_t.font.size = Pt(9)
        run_t.font.italic = True
        run_t.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
        
        run_d = p_ref.add_run(r_desc)
        run_d.font.name = "Microsoft JhengHei"
        run_d.font.size = Pt(8.5)
        run_d.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        
    add_heading_2(doc, "5.2 航運行政常用 Excel 資料前處理函數速查手冊")
    add_body_p(doc, "表 5-1 彙總了航運行政人員日常前處理最常用的十大王牌函數，建議彩色列印常備於辦公桌旁以供隨時查閱：")
    
    cheat_hdrs = ["函數名稱", "標準語法結構", "功能分類", "航運行政核心應用範例"]
    cheat_w = [Inches(1.2), Inches(2.2), Inches(1.1), Inches(2.3)]
    cheat_d = [
        ["TRIM", "=TRIM(文字)", "文字清洗", "剔除提單號碼、貨櫃號碼與港口代碼的首尾空白贅字。"],
        ["UPPER / LOWER", "=UPPER(文字)", "大小寫轉換", "將全小寫港口代碼 'twkhh' 強制轉為標準大寫 'TWKHH'。"],
        ["SUBSTITUTE", "=SUBSTITUTE(文字, 舊字, 新字)", "字串替換", "拔除毛重中的 'kg' 單位，或將日期中的 '.' 替換為 '/'。"],
        ["LEFT / MID / RIGHT", "=LEFT(文字, 字數)", "字串截取", "自提單或訂艙代號中截取前綴船公司代碼 (如 EMC、WHL)。"],
        ["FIND", "=FIND(尋找文字, 字串)", "座標定位", "找出斜線 '/' 或連字號 '-' 所在之字元位置，供 MID 截取。"],
        ["VALUE", "=VALUE(文字數字)", "型態轉換", "將拔除幣別與單位後的文字型數字強制轉為可計算之數值。"],
        ["DATEVALUE", "=DATEVALUE(日期字串)", "日期轉換", "將多樣化格式之日期字串轉換為 Excel 內部標準日期序列值。"],
        ["XLOOKUP", "=XLOOKUP(目標, 查尋欄, 傳回欄)", "動態查表", "根據 5 碼 LOCODE 瞬間帶出港口中文名，取代傳統 VLOOKUP。"],
        ["IF / OR", "=IF(條件, 真值, 假值)", "邏輯防錯", "判斷託運人或付款狀態是否留白，自動填補預設審計標記。"],
        ["COUNTIF", "=COUNTIF(範圍, 條件)", "統計查重", "在連續範圍內偵測提單號碼出現次數是否 > 1，抓出重複建單。"]
    ]
    add_styled_table(doc, cheat_hdrs, cheat_d, cheat_w, align_cols=['C', 'L', 'C', 'L'], font_size=Pt(8))
    
    add_heading_2(doc, "5.3 航運行政人員資料分析進階學習藍圖建議")
    add_body_p(doc, "恭喜你！完整研讀至此，你已經掌握了現代商務最關鍵的資料前處理核心思維與實用電腦操作技能。為了在航運職場更上一層樓，我們建議學員遵循以下三階段進階學習路徑：")
    add_body_p(doc, "• 【第一階段：奠定資料衛生基石】（即本手冊之範圍）：熟練工作表網格、一維標準二維表設計、快捷鍵操作、文字/數值/日期核心清洗函數與防錯驗證。\n"
                    "• 【第二階段：精通樞紐分析與進階 Power Query】（中級技能）：學習運用 Excel 樞紐分析表（Pivot Table）與交叉分析篩選器進行多維度彙總；學習運用 Power Query（取得及轉換資料）建立『一鍵刷新』的自動化資料清洗巨集流程，徹底告別重複手工。\n"
                    "• 【第三階段：航運商業智慧 (BI) 與數位決策】（高級分析師）：學習 Power BI 或 Excel 儀表板設計，整合全球航線準點率監控、長約運價比價模型、貨櫃滯留期（Demurrage & Detention）成本優化儀表板，成為航運物流產業無可取代的高階數位營運人才！")
                    
    add_callout_box(doc, "結語與致學員",
                    "『乾淨的資料，是所有精準商業決策的唯一起點。』養成每次打開 Excel 先審視資料結構、排查空白、確認型態的良好衛生習慣，將讓你在未來的航運行政與商業分析工作中，以倍速脫穎而出，徹底告別加班核對資料的夢魘！",
                    box_type='tip')

# ==========================================
# MAIN BUILDER SCRIPT
# ==========================================
def build_entire_document():
    doc = docx.Document()
    
    # Page setup
    for sec in doc.sections:
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
        sec.left_margin = Inches(0.8)
        sec.right_margin = Inches(0.8)
        
    print("Step 1: Creating Cover Page...")
    create_cover_page(doc)
    
    print("Step 2: Creating Table of Contents and Lists of Figures/Tables...")
    create_toc_and_lists(doc)
    
    print("Step 3: Setting up Body Section & Headers/Footers...")
    setup_body_section(doc)
    
    print("Step 4: Building Chapter 1 (Excel Interface & Navigation)...")
    build_chapter_1(doc)
    
    print("Step 5: Building Chapter 2 (Foundational Data Hygiene & Principles)...")
    build_chapter_2(doc)
    
    print("Step 6: Building Chapter 3 (30-Row Demo Dataset Walkthrough)...")
    build_chapter_3(doc)
    
    print("Step 7: Building Chapter 4 (300-Row Practice Scenarios & Full Solutions)...")
    build_chapter_4(doc)
    
    print("Step 8: Building Chapter 5 (References & Cheat Sheets)...")
    build_chapter_5(doc)
    
    output_filename = "航運行政人員_MS_Excel基礎資料前處理操作實戰指引.docx"
    doc.save(output_filename)
    print(f"Document successfully created and saved as: {output_filename}")

if __name__ == "__main__":
    build_entire_document()
