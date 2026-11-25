import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ==========================================
# CONSTANTS & THEME
# ==========================================
FONT_HEADING = "Microsoft JhengHei"
FONT_BODY = "Microsoft JhengHei"
FONT_CODE = "Consolas"

# Palette
C_BG = RGBColor(0x07, 0x0D, 0x18)          # Deep Navy Dark
C_CARD = RGBColor(0x0F, 0x1B, 0x2D)        # Card Background
C_CARD_ALT = RGBColor(0x13, 0x22, 0x38)    # Slightly lighter card
C_BORDER = RGBColor(0x1E, 0x3A, 0x5F)      # Card border
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)       # Pure white text
C_CYAN = RGBColor(0x38, 0xBD, 0xF8)        # Brand accent Cyan
C_AMBER = RGBColor(0xF5, 0x9E, 0x0B)       # Accent Amber
C_EMERALD = RGBColor(0x10, 0xB9, 0x81)     # Accent Green
C_ROSE = RGBColor(0xF4, 0x3F, 0x5E)        # Accent Red
C_MUTED = RGBColor(0x94, 0xA3, 0xB8)       # Muted gray text

def init_presentation():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs

def create_blank_slide(prs, bg_color=C_BG):
    blank_layout = prs.slide_layouts[6] # completely blank
    slide = prs.slides.add_slide(blank_layout)
    
    # Solid background fill shape
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = bg_color
    bg.line.fill.background()
    return slide

def add_header(slide, chapter_text, slide_title, slide_subtitle, slide_num, total_slides=23):
    # Top progress bar line
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333 * (slide_num / total_slides)), Inches(0.06))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_CYAN
    bar.line.fill.background()
    
    # Category badge & Slide num
    tb_badge = slide.shapes.add_textbox(Inches(0.8), Inches(0.25), Inches(11.733), Inches(0.3))
    tf_b = tb_badge.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
    p_b = tf_b.paragraphs[0]
    r_b = p_b.add_run()
    r_b.text = f"{chapter_text.upper()}   |   SLIDE {str(slide_num).padStart(2, '0') if hasattr(str, 'padStart') else f'{slide_num:02d}'} / {total_slides:02d}"
    r_b.font.name = FONT_BODY
    r_b.font.size = Pt(10)
    r_b.font.bold = True
    r_b.font.color.rgb = C_CYAN

    # Title
    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(11.733), Inches(0.55))
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True
    tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
    p_t = tf_t.paragraphs[0]
    r_t = p_t.add_run()
    r_t.text = slide_title
    r_t.font.name = FONT_HEADING
    r_t.font.size = Pt(22)
    r_t.font.bold = True
    r_t.font.color.rgb = C_WHITE

    # Subtitle
    if slide_subtitle:
        tb_sub = slide.shapes.add_textbox(Inches(0.8), Inches(1.15), Inches(11.733), Inches(0.35))
        tf_s = tb_sub.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = tf_s.margin_top = tf_s.margin_right = tf_s.margin_bottom = 0
        p_s = tf_s.paragraphs[0]
        r_s = p_s.add_run()
        r_s.text = slide_subtitle
        r_s.font.name = FONT_BODY
        r_s.font.size = Pt(11)
        r_s.font.color.rgb = C_MUTED

def add_footer(slide):
    tb_f = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.3))
    tf = tb_f.text_frame
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "航運行政人員 MS Excel 基礎資料分析前處理實戰指引  •  航運數位營運培訓組  •  Microsoft 365 / Excel 2021+ 專用"
    r.font.name = FONT_BODY
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

def set_speaker_notes(slide, notes_text):
    notes = slide.notes_slide.notes_text_frame
    notes.text = notes_text

def add_card(slide, left, top, width, height, bg_color=C_CARD, border_color=C_BORDER):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(1)
    return shape

# ==========================================
# SLIDE BUILDERS (1 TO 23)
# ==========================================

# Slide 1: Cover
def build_slide_1(prs):
    slide = create_blank_slide(prs)
    set_speaker_notes(slide, "【演講者備忘】：歡迎各位航運物流同仁。本手冊專注於解決大家在日常提單、訂艙與運費報表中遇到的『骯髒資料』問題。掌握前處理，將節省您 80% 的無效核對時間！")
    
    # Outer decorative banner
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.0), Inches(5.5), Inches(0.45))
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGBColor(0x0F, 0x2A, 0x4A)
    badge.line.color.rgb = C_CYAN
    badge.line.width = Pt(1)
    tf_b = badge.text_frame
    p_b = tf_b.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    r_b = p_b.add_run()
    r_b.text = "MARITIME LOGISTICS & DATA ANALYTICS TRAINING"
    r_b.font.name = FONT_BODY
    r_b.font.size = Pt(10)
    r_b.font.bold = True
    r_b.font.color.rgb = C_CYAN

    # Main Title
    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.733), Inches(1.6))
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    r_t = p_t.add_run()
    r_t.text = "航運行政人員 MS Excel\n基礎資料前處理操作實戰指引"
    r_t.font.name = FONT_HEADING
    r_t.font.size = Pt(36)
    r_t.font.bold = True
    r_t.font.color.rgb = C_WHITE

    # Divider bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(3.4), Inches(2.5), Inches(0.08))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_CYAN
    bar.line.fill.background()

    # Subtitle
    tb_sub = slide.shapes.add_textbox(Inches(0.8), Inches(3.6), Inches(11.733), Inches(0.8))
    tf_s = tb_sub.text_frame
    tf_s.word_wrap = True
    p_s = tf_s.paragraphs[0]
    r_s = p_s.add_run()
    r_s.text = "從零認識介面、標準化資料衛生、30筆真實業務示範到300筆大數據案例全解\n―― 專為海運貨代、船務行政、港埠營運人員量身打造的電腦技能速成手冊"
    r_s.font.name = FONT_BODY
    r_s.font.size = Pt(13)
    r_s.font.color.rgb = C_MUTED

    # Metadata Grid (4 Cards)
    meta_items = [
        ("示範工具", "MS Excel 365 / 2021+", C_CYAN),
        ("培訓受眾", "航運/海運行政人員", C_EMERALD),
        ("實作檔案", "30筆示範 + 300筆大數據", C_AMBER),
        ("核心目標", "資料衛生與精準清洗", C_WHITE)
    ]
    card_w = Inches(2.7)
    gap = Inches(0.3)
    start_x = Inches(0.8)
    
    for i, (lbl, val, col) in enumerate(meta_items):
        cx = start_x + i * (card_w + gap)
        c = add_card(slide, cx, Inches(4.7), card_w, Inches(1.4), C_CARD, C_BORDER)
        
        tb = slide.shapes.add_textbox(cx, Inches(4.85), card_w, Inches(1.1))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        r1 = p1.add_run()
        r1.text = lbl
        r1.font.name = FONT_BODY
        r1.font.size = Pt(11)
        r1.font.color.rgb = C_MUTED

        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(4)
        r2 = p2.add_run()
        r2.text = val
        r2.font.name = FONT_HEADING
        r2.font.size = Pt(12)
        r2.font.bold = True
        r2.font.color.rgb = col

    add_footer(slide)

# Slide 2: Agenda
def build_slide_2(prs):
    slide = create_blank_slide(prs)
    add_header(slide, "課程導覽", "培訓課程全景地圖 (Course Agenda)", "循序漸進的五大實戰模組，系統化打造航運數據標準化作業思維", 2)
    set_speaker_notes(slide, "【演講者備忘】：課程由淺入深，從最基礎的介面巡禮，到資料衛生觀念，再透過 30 筆示範操作與 300 筆大數據挑戰，最後提供永久速查工具手冊。")

    modules = [
        ("01", "第一章：Excel 操作介面導航", "認識標題列、功能區、名稱方塊、編輯列、狀態列與十大必備高頻快捷鍵。", C_CYAN),
        ("02", "第二章：資料前處理基本常識", "GIGO 原則、第一正規化 (1NF)、航運七大骯髒資料病症與四大儲存格本質型態。", C_EMERALD),
        ("03", "第三章：30 筆業務示範實戰", "資料剖析、Ctrl+E 快速填入、移除重複項、取代、驗證下拉選單與核心文字函數鏈。", C_AMBER),
        ("04", "第四章：300 筆大數據情境全解", "UN/LOCODE 港口關聯、ISO 櫃型 TEU 換算、八大任務逐題精解與超期滯港費預警運算。", C_CYAN),
        ("05", "第五章：參考文獻與速查寶典", "國際海運標準、DAMA 資料治理、前處理十大王牌函數速查手冊與進階學習藍圖。", C_ROSE),
        ("★", "互動式航運實務模擬工具", "文字清洗動態試驗場、滯港費 (Demurrage) 與航程天數即時運算機（收錄於互動網頁中）。", C_WHITE)
    ]

    col_w = Inches(3.7)
    row_h = Inches(2.2)
    start_x = Inches(0.8)
    start_y = Inches(1.8)

    for i, (num, title, desc, col) in enumerate(modules):
        r_idx = i // 3
        c_idx = i % 3
        cx = start_x + c_idx * (col_w + Inches(0.3))
        cy = start_y + r_idx * (row_h + Inches(0.25))

        add_card(slide, cx, cy, col_w, row_h, C_CARD, C_BORDER)

        tb = slide.shapes.add_textbox(cx + Inches(0.2), cy + Inches(0.2), col_w - Inches(0.4), row_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = f"MODULE {num}"
        r0.font.name = FONT_CODE
        r0.font.size = Pt(10)
        r0.font.bold = True
        r0.font.color.rgb = col

        p1 = tf.add_paragraph()
        p1.space_before = Pt(4)
        r1 = p1.add_run()
        r1.text = title
        r1.font.name = FONT_HEADING
        r1.font.size = Pt(13)
        r1.font.bold = True
        r1.font.color.rgb = C_WHITE

        p2 = tf.add_paragraph()
        p2.space_before = Pt(4)
        r2 = p2.add_run()
        r2.text = desc
        r2.font.name = FONT_BODY
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = C_MUTED

    add_footer(slide)

# Slide 3: Ch1 Why Preprocessing
def build_slide_3(prs):
    slide = create_blank_slide(prs)
    add_header(slide, "第一章：介面導航", "1.1 航運行政人員為何必須精通資料前處理？", "資料科學的 80/20 法則與『垃圾進，垃圾出 (Garbage In, Garbage Out)』效應", 3)
    set_speaker_notes(slide, "【實務警語】：航運提單多一個不可見空格，VLOOKUP 就抓不到報關行資訊；重量多打一個 'kg'，SUM 加總直接歸零！前處理是商業決策的唯一起點。")

    col_w = Inches(5.7)
    cy = Inches(1.8)
    h = Inches(4.3)

    # Left: Disasters
    add_card(slide, Inches(0.8), cy, col_w, h, RGBColor(0x1B, 0x12, 0x18), RGBColor(0x5A, 0x1E, 0x28))
    tb_l = slide.shapes.add_textbox(Inches(1.0), cy + Inches(0.2), col_w - Inches(0.4), h - Inches(0.4))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True

    p_t = tf_l.paragraphs[0]
    r_t = p_t.add_run()
    r_t.text = "❌ 未經前處理的航運災難實況"
    r_t.font.name = FONT_HEADING
    r_t.font.size = Pt(14)
    r_t.font.bold = True
    r_t.font.color.rgb = C_ROSE

    items_l = [
        ("肉眼難辨的隱形空格：", "提單號碼多按了一格空白，VLOOKUP 永遠回傳 #N/A，造成海關報關查無單據延誤！"),
        ("單位塞在數字中：", "貨物毛重手動打入 '18500 kg'，導致 SUM 加總為 0，運費請款核算全盤出錯！"),
        ("多重屬性擠在同格：", "船名與航次擠在同一個儲存格，主管要求按『個別船舶』統計運載量時完全無法分群。"),
        ("重複託運提單建檔：", "同一票貨重複建單，出口總箱數與海運費營收被雙重計算，產生嚴重審計漏洞。")
    ]
    for lbl, desc in items_l:
        p = tf_l.add_paragraph()
        p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.color.rgb = C_WHITE
        r1.font.size = Pt(11)
        r2 = p.add_run()
        r2.text = desc
        r2.font.color.rgb = C_MUTED
        r2.font.size = Pt(10.5)

    # Right: Benefits
    add_card(slide, Inches(6.8), cy, col_w, h, RGBColor(0x0F, 0x22, 0x22), RGBColor(0x10, 0x55, 0x44))
    tb_r = slide.shapes.add_textbox(Inches(7.0), cy + Inches(0.2), col_w - Inches(0.4), h - Inches(0.4))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True

    p_rt = tf_r.paragraphs[0]
    r_rt = p_rt.add_run()
    r_rt.text = "✅ 掌握標準前處理的效益飛躍"
    r_rt.font.name = FONT_HEADING
    r_rt.font.size = Pt(14)
    r_rt.font.bold = True
    r_rt.font.color.rgb = C_EMERALD

    items_r = [
        ("節省 80% 無效工時：", "從過去耗費數小時人工手動複製貼上核對，轉變為幾秒鐘批次自動化清洗完成！"),
        ("樞紐分析一鍵產出：", "標準一維二維表結構，讓交叉分析篩選器與動態視覺化圖表隨點隨動。"),
        ("跨國 EDI 無縫對接：", "清洗後的標準 UN/LOCODE 與 ISO 貨櫃編碼，與全球海關及船公司報文 100% 相容。"),
        ("精確預警超期滯港費：", "到港日與提櫃日自動比對免租期，一鍵過濾出 Demurrage 罰款名單，為公司挽回損失。")
    ]
    for lbl, desc in items_r:
        p = tf_r.add_paragraph()
        p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = f"• {lbl} "
        r1.font.bold = True
        r1.font.color.rgb = C_WHITE
        r1.font.size = Pt(11)
        r2 = p.add_run()
        r2.text = desc
        r2.font.color.rgb = C_MUTED
        r2.font.size = Pt(10.5)

    # Bottom Golden Quote
    card_b = add_card(slide, Inches(0.8), Inches(6.25), Inches(11.7), Inches(0.65), C_CARD_ALT, C_CYAN)
    tb_q = slide.shapes.add_textbox(Inches(1.0), Inches(6.3), Inches(11.3), Inches(0.55))
    tf_q = tb_q.text_frame
    p_q = tf_q.paragraphs[0]
    r_q1 = p_q.add_run()
    r_q1.text = "💡 資料科學至理名言： "
    r_q1.font.bold = True
    r_q1.font.color.rgb = C_CYAN
    r_q1.font.size = Pt(11)
    r_q2 = p_q.add_run()
    r_q2.text = "『Garbage In, Garbage Out』—— 缺乏品質前處理的數據，算出來的報表再漂亮，本質上依舊是錯誤決策的垃圾！"
    r_q2.font.color.rgb = C_WHITE
    r_q2.font.size = Pt(11)

    add_footer(slide)

# Generic Image Slide Helper
def build_image_slide(prs, chapter, title, subtitle, img_name, caption, notes, slide_num):
    slide = create_blank_slide(prs)
    add_header(slide, chapter, title, subtitle, slide_num)
    set_speaker_notes(slide, notes)

    img_path = os.path.join("images", img_name)
    if os.path.exists(img_path):
        # Card container for image
        add_card(slide, Inches(1.2), Inches(1.75), Inches(10.933), Inches(4.7), C_CARD, C_BORDER)
        # Add picture
        slide.shapes.add_picture(img_path, Inches(1.3), Inches(1.85), width=Inches(10.733))
    
    # Caption
    tb_c = slide.shapes.add_textbox(Inches(1.2), Inches(6.55), Inches(10.933), Inches(0.35))
    tf = tb_c.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = caption
    r.font.name = FONT_BODY
    r.font.size = Pt(10.5)
    r.font.bold = True
    r.font.color.rgb = C_CYAN

    add_footer(slide)

# Slide 5: Shortcuts
def build_slide_5(prs):
    slide = create_blank_slide(prs)
    add_header(slide, "第一章：介面導航", "1.3 航運行政人員必備之十組超高效快捷鍵", "擺脫滑鼠頻繁點選拉伸的手痠夢魘，操作前處理效率瞬間飆升 5 倍", 5)
    set_speaker_notes(slide, "【操作訣竅】：最神快捷鍵首推【Ctrl + Shift + L】(一秒開關自動篩選) 與【Ctrl + E】(快速填入智慧萃取)！")

    shortcuts = [
        ("Ctrl + Shift + L", "啟動 / 關閉自動篩選", "一秒為全欄標題加上下拉箭頭，快速聚合並嗅探髒資料唯一值首選。"),
        ("Ctrl + E", "智慧快速填入 (Flash Fill)", "手動輸入第一筆範本後按下，Excel 自動嗅探規律秒速填滿全欄！"),
        ("Ctrl + H", "尋找與取代對話方塊", "批次將重量 'kg'、幣別 '$' 取代留空，瞬間將字串純數值化。"),
        ("Ctrl + 1", "開啟儲存格格式設定", "自訂日期格式 (YYYY-MM-DD)、數值千分位逗號與小數位數規範。"),
        ("Ctrl + 方向鍵", "瞬間飛躍至資料邊界", "在數千筆提單中，一秒跳轉至最頂端（首筆）或最底端（末筆）資料列。"),
        ("Alt + = (等號)", "一秒自動加總 (AutoSum)", "在運費或毛重欄位底端，一鍵自動插入 SUM 加總公式。")
    ]

    col_w = Inches(5.7)
    row_h = Inches(1.4)
    start_x = Inches(0.8)
    start_y = Inches(1.85)

    for i, (key, title, desc) in enumerate(shortcuts):
        r_idx = i // 2
        c_idx = i % 2
        cx = start_x + c_idx * (col_w + Inches(0.3))
        cy = start_y + r_idx * (row_h + Inches(0.2))

        add_card(slide, cx, cy, col_w, row_h, C_CARD, C_BORDER)

        # Key badge
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + Inches(0.2), cy + Inches(0.25), Inches(1.8), Inches(0.45))
        badge.fill.solid()
        badge.fill.fore_color.rgb = RGBColor(0x1F, 0x33, 0x4D)
        badge.line.color.rgb = C_CYAN
        badge.line.width = Pt(1)
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.alignment = PP_ALIGN.CENTER
        r_b = p_b.add_run()
        r_b.text = key
        r_b.font.name = FONT_CODE
        r_b.font.size = Pt(10.5)
        r_b.font.bold = True
        r_b.font.color.rgb = C_CYAN

        # Content
        tb = slide.shapes.add_textbox(cx + Inches(2.15), cy + Inches(0.15), col_w - Inches(2.3), row_h - Inches(0.3))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        r1 = p1.add_run()
        r1.text = title
        r1.font.name = FONT_HEADING
        r1.font.size = Pt(12)
        r1.font.bold = True
        r1.font.color.rgb = C_WHITE

        p2 = tf.add_paragraph()
        p2.space_before = Pt(3)
        r2 = p2.add_run()
        r2.text = desc
        r2.font.name = FONT_BODY
        r2.font.size = Pt(10)
        r2.font.color.rgb = C_MUTED

    add_footer(slide)

# Slide 7: 7 Dirty Data Diseases
def build_slide_7(prs):
    slide = create_blank_slide(prs)
    add_header(slide, "第二章：資料衛生常識", "2.2 航運行政常見『七大骯髒資料 (Dirty Data)』診斷總覽", "點對點直擊航運實務痛點、破壞性後果剖析與標準處理解方", 7)
    set_speaker_notes(slide, "【演講者提示】：請引導學員逐一審視這七大病症，說明其在海運提單、EDI 報關與運費應收應付中引發的致命連鎖反應。")

    diseases = [
        ("D-01", "首尾不可見空格", "'  EMC-2026-001  '", "VLOOKUP 比對失敗 (#N/A)", "=TRIM(B2) 拔除首尾空格"),
        ("D-02", "代碼大小寫紊亂", "'twkhh' vs 'TWKHH'", "樞紐分析表分群統計腰斬", "=UPPER(TRIM(D2)) 強制轉大寫"),
        ("D-03", "單欄混雜多屬性", "船名/航次同欄擠在同格", "無法單獨統計單一船舶趟次", "資料剖析或 Ctrl+E 快速填入"),
        ("D-04", "數字存為文字", "毛重 '18,500 kg'、'$2,450'", "無法加總，SUM 算出為 0 元", "取代 'kg' 搭配 =VALUE(...) 轉換"),
        ("D-05", "多樣化混亂日期", "2026/10/05、10-10-2026", "文字字串無法相減算航期", "資料剖析指定 YMD 轉序列值"),
        ("D-06", "重複託運提單", "同一提單號建檔 2 次", "箱數與運費營收虛胖翻倍", "【移除重複項】單獨勾選提單號"),
        ("D-07", "關鍵欄位遺漏值", "託運人或付款狀態留白", "樞紐分析表出現 '(空白)'", "=IF(M2=\"\", \"[待補]\", M2) 防錯")
    ]

    # Create a clean 7-row table
    rows = 8
    cols = 5
    table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.75), Inches(11.733), Inches(4.7))
    tbl = table_shape.table

    tbl.columns[0].width = Inches(1.0)
    tbl.columns[1].width = Inches(2.2)
    tbl.columns[2].width = Inches(2.6)
    tbl.columns[3].width = Inches(3.1)
    tbl.columns[4].width = Inches(2.833)

    headers = ["編號", "骯髒資料病症", "典型航運資料範例", "實務破壞性影響", "標準清洗對策"]
    for c_idx, h in enumerate(headers):
        cell = tbl.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x1F, 0x33, 0x4D)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h
        r.font.name = FONT_HEADING
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = C_CYAN

    for r_idx, (num, name, ex, dmg, sol) in enumerate(diseases):
        row_cells = tbl.cell(r_idx + 1, 0), tbl.cell(r_idx + 1, 1), tbl.cell(r_idx + 1, 2), tbl.cell(r_idx + 1, 3), tbl.cell(r_idx + 1, 4)
        bg_c = C_CARD if r_idx % 2 == 0 else C_CARD_ALT

        for cell in row_cells:
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_c

        # Col 0: Num
        p0 = row_cells[0].text_frame.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        r0 = p0.add_run()
        r0.text = num
        r0.font.name = FONT_CODE
        r0.font.bold = True
        r0.font.size = Pt(10)
        r0.font.color.rgb = C_AMBER

        # Col 1: Name
        p1 = row_cells[1].text_frame.paragraphs[0]
        r1 = p1.add_run()
        r1.text = name
        r1.font.name = FONT_HEADING
        r1.font.bold = True
        r1.font.size = Pt(10.5)
        r1.font.color.rgb = C_WHITE

        # Col 2: Ex
        p2 = row_cells[2].text_frame.paragraphs[0]
        r2 = p2.add_run()
        r2.text = ex
        r2.font.name = FONT_CODE
        r2.font.size = Pt(10)
        r2.font.color.rgb = RGBColor(0xFB, 0xBF, 0x24)

        # Col 3: Dmg
        p3 = row_cells[3].text_frame.paragraphs[0]
        r3 = p3.add_run()
        r3.text = dmg
        r3.font.name = FONT_BODY
        r3.font.size = Pt(10)
        r3.font.color.rgb = C_ROSE

        # Col 4: Sol
        p4 = row_cells[4].text_frame.paragraphs[0]
        r4 = p4.add_run()
        r4.text = sol
        r4.font.name = FONT_BODY
        r4.font.bold = True
        r4.font.size = Pt(10)
        r4.font.color.rgb = C_EMERALD

    add_footer(slide)

# Slide 8: 4 Data Types & Traps
def build_slide_8(prs):
    slide = create_blank_slide(prs)
    add_header(slide, "第二章：資料衛生常識", "2.3 Excel 四大本質資料型態與隱形陷阱", "看清儲存格底層本質而非視覺外觀：文字、數值、日期序列值與綠色錯誤三角形", 8)
    set_speaker_notes(slide, "【新手避坑】：合格的日期在 Excel 內部是自 1900/1/1 起算的天數數值！只有真正的日期序列值，才能直接相減 (ETA - ETD) 計算航程天數。")

    types = [
        ("文字型態 (Text / String)", "包括英文字母、中文、帶單位的數字。\n• 預設外觀：靠左對齊\n• 特性：絕對無法參與加減乘除與 SUM 加總！", C_ROSE),
        ("數值型態 (Numeric / Number)", "整數、小數、金額。\n• 預設外觀：靠右對齊\n• 特性：可直接參與數學運算與樞紐分析統計。", C_EMERALD),
        ("日期時間 (Date & Time)", "Excel 以『1900/1/1 起算之天數序列值』儲存。\n• 預設外觀：合格日期靠右對齊\n• 核心優勢：兩日期直接相減 (ETA-ETD) 算天數！", C_CYAN),
        ("布林值 (Boolean)", "只有 TRUE 與 FALSE 兩種狀態。\n• 預設外觀：置中對齊\n• 核心應用：用於 IF 邏輯條件防錯判斷。", C_AMBER)
    ]

    col_w = Inches(5.7)
    row_h = Inches(2.05)
    start_x = Inches(0.8)
    start_y = Inches(1.8)

    for i, (title, desc, col) in enumerate(types):
        r_idx = i // 2
        c_idx = i % 2
        cx = start_x + c_idx * (col_w + Inches(0.3))
        cy = start_y + r_idx * (row_h + Inches(0.2))

        add_card(slide, cx, cy, col_w, row_h, C_CARD, C_BORDER)

        tb = slide.shapes.add_textbox(cx + Inches(0.25), cy + Inches(0.2), col_w - Inches(0.5), row_h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        r1 = p1.add_run()
        r1.text = title
        r1.font.name = FONT_HEADING
        r1.font.size = Pt(13)
        r1.font.bold = True
        r1.font.color.rgb = col

        p2 = tf.add_paragraph()
        p2.space_before = Pt(6)
        r2 = p2.add_run()
        r2.text = desc
        r2.font.name = FONT_BODY
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = C_WHITE

    # Trap Box
    trap_card = add_card(slide, Inches(0.8), Inches(6.15), Inches(11.733), Inches(0.75), RGBColor(0x2A, 0x1A, 0x05), C_AMBER)
    tb_tr = slide.shapes.add_textbox(Inches(1.0), Inches(6.2), Inches(11.3), Inches(0.65))
    tf_tr = tb_tr.text_frame
    p_tr = tf_tr.paragraphs[0]
    r_tr1 = p_tr.add_run()
    r_tr1.text = "⚠️ 初學者致命隱形陷阱：綠色錯誤檢查三角形 (Error Triangle)\n"
    r_tr1.font.bold = True
    r_tr1.font.color.rgb = C_AMBER
    r_tr1.font.size = Pt(10.5)
    r_tr2 = p_tr.add_run()
    r_tr2.text = "當儲存格左上角出現微小綠色三角形時，代表 Excel 發出『以文字形式儲存的數字』警訊！點選驚嘆號並選擇『轉換成數字』即可修復。"
    r_tr2.font.color.rgb = C_WHITE
    r_tr2.font.size = Pt(10)

    add_footer(slide)

# Slide 9: 30-Row Demo Comparison Table
def build_slide_9(prs):
    slide = create_blank_slide(prs)
    add_header(slide, "第三章：30筆實戰示範", "3.1 30 筆真實航運業務資料檔：清洗前後對照", "從充滿空格、混合幣別、文字日期與重複提單，提煉為標準化決策資料庫", 9)
    set_speaker_notes(slide, "【實例說明】：觀察原本帶有單位、混亂日期的原始表格，如何被梳理成標準化的一維資料庫，重複提單被精確排除。")

    # Table layout
    rows = 7
    cols = 6
    tbl_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.75), Inches(11.733), Inches(4.7))
    tbl = tbl_shape.table

    col_widths = [Inches(1.5), Inches(2.2), Inches(2.0), Inches(2.0), Inches(2.0), Inches(2.033)]
    for i, w in enumerate(col_widths):
        tbl.columns[i].width = w

    hdrs = ["提單號碼", "原始船名航次 (Raw)", "清洗後船名", "清洗後航次", "純毛重 (KG)", "海運費 (USD)"]
    for i, h in enumerate(hdrs):
        cell = tbl.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x1F, 0x33, 0x4D)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h
        r.font.name = FONT_HEADING
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = C_CYAN

    data_sample = [
        ("EMC-2026-001", "EVER GIVEN / 0142-052W", "EVER GIVEN", "0142-052W", "18,500", "$2,450.00"),
        ("WHL-2026-088", "WAN HAI 315 / 0089-012N", "WAN HAI 315", "0089-012N", "12,400", "$1,200.00"),
        ("YML-2026-105", "YM WELLBEING / 0214-088E", "YM WELLBEING", "0214-088E", "21,300", "$3,100.00"),
        ("OOCL-2026-302", "OOCL POLAND / 0055-031W", "OOCL POLAND", "0055-031W", "19,800", "$2,600.00"),
        ("MSK-2026-991", "MAERSK MC / 0411-002W", "MAERSK MC", "0411-002W", "22,500", "$3,400.00"),
        ("EMC-2026-035", "EVER BRAVE / 0067-088N", "EVER BRAVE", "0067-088N", "8,900", "$800.00")
    ]

    for r_idx, row in enumerate(data_sample):
        bg = C_CARD if r_idx % 2 == 0 else C_CARD_ALT
        for c_idx, val in enumerate(row):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            if c_idx in [0, 2, 3]:
                p.alignment = PP_ALIGN.CENTER
            elif c_idx in [4, 5]:
                p.alignment = PP_ALIGN.RIGHT
            else:
                p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = val
            r.font.name = FONT_CODE if c_idx in [0, 3, 4, 5] else FONT_BODY
            r.font.size = Pt(10)
            if c_idx == 1:
                r.font.color.rgb = C_ROSE
            elif c_idx in [2, 3]:
                r.font.color.rgb = C_CYAN
            elif c_idx in [4, 5]:
                r.font.color.rgb = C_EMERALD
            else:
                r.font.color.rgb = C_WHITE

    add_footer(slide)

# Slide 13: Dual Images (Find & Replace, Data Validation)
def build_slide_13(prs):
    slide = create_blank_slide(prs)
    add_header(slide, "第三章：30筆實戰示範", "3.5 功能選項 4 & 5：【尋找與取代】與【資料驗證】源頭防錯", "批次剔除 'kg' 單位與 '$' 符號；建立 PREPAID / COLLECT 下拉選單徹底杜絕人工打錯", 13)
    set_speaker_notes(slide, "【源頭治理】：與其事後花時間辛苦清洗，不如在資料輸入端直接設『資料驗證清單』，強制行政人員只能選標準值！")

    col_w = Inches(5.7)
    cy = Inches(1.8)
    h = Inches(4.5)

    # Left Image Card
    add_card(slide, Inches(0.8), cy, col_w, h, C_CARD, C_BORDER)
    img1 = os.path.join("images", "fig6_find_and_replace_dialog.png")
    if os.path.exists(img1):
        slide.shapes.add_picture(img1, Inches(0.95), cy + Inches(0.2), width=Inches(5.4))
    tb_c1 = slide.shapes.add_textbox(Inches(0.8), cy + Inches(3.9), col_w, Inches(0.4))
    p1 = tb_c1.text_frame.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = "圖 3-4：尋找與取代 (Ctrl+H) 批次拔除單位"
    r1.font.name = FONT_BODY
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = C_CYAN

    # Right Image Card
    add_card(slide, Inches(6.8), cy, col_w, h, C_CARD, C_BORDER)
    img2 = os.path.join("images", "fig7_data_validation_dialog.png")
    if os.path.exists(img2):
        slide.shapes.add_picture(img2, Inches(6.95), cy + Inches(0.2), width=Inches(5.4))
    tb_c2 = slide.shapes.add_textbox(Inches(6.8), cy + Inches(3.9), col_w, Inches(0.4))
    p2 = tb_c2.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = "圖 3-5：資料驗證設定合法清單下拉選單"
    r2.font.name = FONT_BODY
    r2.font.size = Pt(10)
    r2.font.bold = True
    r2.font.color.rgb = C_CYAN

    add_footer(slide)

# Slide 19: Tasks 1-4 Solutions
def build_slide_19(prs):
    slide = create_blank_slide(prs)
    add_header(slide, "第四章：300筆大數據實戰", "4.3 任務挑戰精解 (一)：任務 1 至任務 4 標準解法卡", "提單與訂艙編號標準化、港口名稱跨表關聯、貨櫃 TEU 換算與毛重數值化", 19)
    set_speaker_notes(slide, "【演講者備忘】：毛重欄位含有 kgs、kg 與逗號，公式採用三重 SUBSTITUTE 徹底淨化，最後由 VALUE 轉回數值。")

    tasks = [
        ("任務一", "提單與訂艙編號標準化", "消除首尾多餘空格，英文字母全數轉為大寫標準格式。", "=TRIM(UPPER(C2))", C_CYAN),
        ("任務二", "裝卸港代碼與中文名跨表關聯", "POL/POD 代碼清洗大寫，利用 XLOOKUP 跨表自動帶出港口中文全名。", "=XLOOKUP(UPPER(TRIM(F2)), Port_Master!$A$2:$A$18, Port_Master!$B$2:$B$18, \"未知港口\")", C_EMERALD),
        ("任務三", "貨櫃號碼與 ISO 櫃型換算 TEU", "依據 20GP=1, 40GP/40HC=2, 45HC=2.25 係數跨表換算運載量。", "=XLOOKUP(I2, ISO_Type!$A$2:$A$9, ISO_Type!$E$2:$E$9, 1.0)", C_AMBER),
        ("任務四", "貨物毛重與包裝單位純數值化", "剔除 'kg'、'KGS' 與逗號，將文字型數字轉回標準可加總數值。", "=VALUE(SUBSTITUTE(SUBSTITUTE(LOWER(M2), \"kg\", \"\"), \",\", \"\"))", C_ROSE)
    ]

    col_w = Inches(5.7)
    row_h = Inches(2.25)
    start_x = Inches(0.8)
    start_y = Inches(1.8)

    for i, (badge, title, desc, formula, col) in enumerate(tasks):
        r_idx = i // 2
        c_idx = i % 2
        cx = start_x + c_idx * (col_w + Inches(0.3))
        cy = start_y + r_idx * (row_h + Inches(0.25))

        add_card(slide, cx, cy, col_w, row_h, C_CARD, C_BORDER)

        tb = slide.shapes.add_textbox(cx + Inches(0.2), cy + Inches(0.15), col_w - Inches(0.4), row_h - Inches(0.3))
        tf = tb.text_frame
        tf.word_wrap = True

        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = badge
        r0.font.name = FONT_CODE
        r0.font.bold = True
        r0.font.size = Pt(10)
        r0.font.color.rgb = col

        r0_t = p0.add_run()
        r0_t.text = f"  {title}"
        r0_t.font.name = FONT_HEADING
        r0_t.font.bold = True
        r0_t.font.size = Pt(12.5)
        r0_t.font.color.rgb = C_WHITE

        p1 = tf.add_paragraph()
        p1.space_before = Pt(4)
        r1 = p1.add_run()
        r1.text = desc
        r1.font.name = FONT_BODY
        r1.font.size = Pt(10)
        r1.font.color.rgb = C_MUTED

        # Formula code box
        p2 = tf.add_paragraph()
        p2.space_before = Pt(8)
        r2 = p2.add_run()
        r2.text = f"標準公式：{formula}"
        r2.font.name = FONT_CODE
        r2.font.size = Pt(9.5)
        r2.font.bold = True
        r2.font.color.rgb = col

    add_footer(slide)

# Slide 20: Tasks 5-8 Solutions
def build_slide_20(prs):
    slide = create_blank_slide(prs)
    add_header(slide, "第四章：300筆大數據實戰", "4.4 任務挑戰精解 (二)：任務 5 至任務 8 標準解法卡", "運費試算、日期航期、港口免租期 Demurrage 滯港罰款預警與重複稽核", 20)
    set_speaker_notes(slide, "【業務亮點】：任務 7 滯港費預警利用 =MAX(0, (GateOut-ATA)-FreeDays) 搭配 IF 函數，一秒抓出需要向貨主追討滯港費的異常清單！")

    tasks = [
        ("任務五", "海運費淨化與總運費試算", "拔除 '$' 與 'USD' 幣別代碼，將淨化運費與附加費相加計算。", "=VALUE(SUBSTITUTE(SUBSTITUTE(P2, \"$\", \"\"), \"USD\", \"\")) + Q2", C_CYAN),
        ("任務六", "多樣化日期清洗與航期計算", "文字型日期轉為序列值，直接相減得出正確航期天數。", "航程天數公式: =標準ETA - 標準ETD  (如 2026/10/22 - 2026/10/05 = 17天)", C_EMERALD),
        ("任務七", "免租期與滯港費 (Demurrage) 預警", "停留天數超過免租期者，使用 MAX 與 IF 函數自動亮紅燈預警。", "=IF(MAX(0, (GateOut - ATA) - FreeDays) > 0, \"滯港罰款\", \"正常\")", C_AMBER),
        ("任務八", "全域重複排查與缺漏資料稽核", "排查重複提單，空客戶填 '[待確認客戶]'，空海關填 'PENDING'。", "=IF(COUNTIF($C$2:C2, C2) > 1, \"【重複建單】\", \"正常\")", C_ROSE)
    ]

    col_w = Inches(5.7)
    row_h = Inches(2.25)
    start_x = Inches(0.8)
    start_y = Inches(1.8)

    for i, (badge, title, desc, formula, col) in enumerate(tasks):
        r_idx = i // 2
        c_idx = i % 2
        cx = start_x + c_idx * (col_w + Inches(0.3))
        cy = start_y + r_idx * (row_h + Inches(0.25))

        add_card(slide, cx, cy, col_w, row_h, C_CARD, C_BORDER)

        tb = slide.shapes.add_textbox(cx + Inches(0.2), cy + Inches(0.15), col_w - Inches(0.4), row_h - Inches(0.3))
        tf = tb.text_frame
        tf.word_wrap = True

        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = badge
        r0.font.name = FONT_CODE
        r0.font.bold = True
        r0.font.size = Pt(10)
        r0.font.color.rgb = col

        r0_t = p0.add_run()
        r0_t.text = f"  {title}"
        r0_t.font.name = FONT_HEADING
        r0_t.font.bold = True
        r0_t.font.size = Pt(12.5)
        r0_t.font.color.rgb = C_WHITE

        p1 = tf.add_paragraph()
        p1.space_before = Pt(4)
        r1 = p1.add_run()
        r1.text = desc
        r1.font.name = FONT_BODY
        r1.font.size = Pt(10)
        r1.font.color.rgb = C_MUTED

        p2 = tf.add_paragraph()
        p2.space_before = Pt(8)
        r2 = p2.add_run()
        r2.text = f"標準公式：{formula}"
        r2.font.name = FONT_CODE
        r2.font.size = Pt(9.5)
        r2.font.bold = True
        r2.font.color.rgb = col

    add_footer(slide)

# Slide 21: Demurrage Calculation Logic
def build_slide_21(prs):
    slide = create_blank_slide(prs)
    add_header(slide, "第四章：300筆大數據實戰", "4.5 航運實務試算器：滯港費 (Demurrage) 與航期即時運算邏輯", "在線動態調整實際到港日、提櫃日與免租天數，即時檢視預警運算底層原理", 21)
    set_speaker_notes(slide, "【演講者提示】：請學員體驗 Excel 邏輯公式如何在幾毫秒內完成超期滯港判定！善用 MAX(0, 差額) 可有效避免出現負數天數。")

    col_w = Inches(5.7)
    cy = Inches(1.8)
    h = Inches(4.5)

    # Left: Business Logic
    add_card(slide, Inches(0.8), cy, col_w, h, C_CARD, C_BORDER)
    tb_l = slide.shapes.add_textbox(Inches(1.0), cy + Inches(0.2), col_w - Inches(0.4), h - Inches(0.4))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True

    p0 = tf_l.paragraphs[0]
    r0 = p0.add_run()
    r0.text = "📅 港口免租期與滯港費業務背景"
    r0.font.name = FONT_HEADING
    r0.font.size = Pt(14)
    r0.font.bold = True
    r0.font.color.rgb = C_CYAN

    items = [
        ("什麼是免租期 (Free Time)？", "貨櫃運抵卸貨港後，船公司與碼頭提供貨主免費停留存放的天數（如標準 7 天，長約 14 天）。"),
        ("什麼是滯港費 (Demurrage)？", "超過免租期貨主仍未提櫃（Gate-Out），碼頭與船東按日加收之高額場地懲罰性罰金！"),
        ("港口實際停留天數：", "公式：=提櫃日 (Gate-Out) - 實際到港日 (ATA)"),
        ("超期留港天數 (Overdue)：", "公式：=MAX(0, (Gate-Out - ATA) - Free Days)\n※ 巧妙使用 MAX 函數，停留未超期差額為負時自動取 0，絕不出現負天數！")
    ]
    for lbl, desc in items:
        p = tf_l.add_paragraph()
        p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = f"• {lbl}\n  "
        r1.font.bold = True
        r1.font.color.rgb = C_WHITE
        r1.font.size = Pt(11)
        r2 = p.add_run()
        r2.text = desc
        r2.font.color.rgb = C_MUTED
        r2.font.size = Pt(10)

    # Right: Formula & Alert Card
    add_card(slide, Inches(6.8), cy, col_w, h, C_CARD_ALT, C_AMBER)
    tb_r = slide.shapes.add_textbox(Inches(7.0), cy + Inches(0.2), col_w - Inches(0.4), h - Inches(0.4))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True

    p_rt = tf_r.paragraphs[0]
    r_rt = p_rt.add_run()
    r_rt.text = "🚨 Excel 滯港罰款智慧預警公式"
    r_rt.font.name = FONT_HEADING
    r_rt.font.size = Pt(14)
    r_rt.font.bold = True
    r_rt.font.color.rgb = C_AMBER

    p_f1 = tf_r.add_paragraph()
    p_f1.space_before = Pt(10)
    r_f1 = p_f1.add_run()
    r_f1.text = "【超期天數運算公式】：\n=MAX(0, (U2 - T2) - V2)"
    r_f1.font.name = FONT_CODE
    r_f1.font.size = Pt(11)
    r_f1.font.bold = True
    r_f1.font.color.rgb = C_WHITE

    p_f2 = tf_r.add_paragraph()
    p_f2.space_before = Pt(12)
    r_f2 = p_f2.add_run()
    r_f2.text = "【滯港狀態預警標籤公式】：\n=IF(MAX(0, (U2 - T2) - V2) > 0, \"滯港罰款(Demurrage)\", \"正常(Normal)\")"
    r_f2.font.name = FONT_CODE
    r_f2.font.size = Pt(10.5)
    r_f2.font.bold = True
    r_f2.font.color.rgb = C_ROSE

    p_f3 = tf_r.add_paragraph()
    p_f3.space_before = Pt(14)
    r_f3 = p_f3.add_run()
    r_f3.text = "💡 實務成果：在 300 筆實戰資料中，此公式一秒精準抓出 48 筆超期滯港貨櫃，累計超期停留達 286 天，成為船務向客戶請款之鐵證！"
    r_f3.font.size = Pt(10.5)
    r_f3.font.color.rgb = C_WHITE

    add_footer(slide)

# Slide 22: References & Cheat Sheet
def build_slide_22(prs):
    slide = create_blank_slide(prs)
    add_header(slide, "第五章：速查寶典", "5.1 國際海運與資料治理參考文獻 & 王牌函數速查手冊", "遵循 ISO 6346、UN/LOCODE、DAMA-DMBOK2 與 Microsoft 官方指南", 22)
    set_speaker_notes(slide, "【持續查閱】：本頁彙總了十大王牌前處理函數的標準語法與實務情境，建議同仁截圖或列印常備於辦公桌旁。")

    # Table of cheat sheet
    rows = 7
    cols = 4
    tbl_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.75), Inches(11.733), Inches(4.7))
    tbl = tbl_shape.table

    tbl.columns[0].width = Inches(1.8)
    tbl.columns[1].width = Inches(3.2)
    tbl.columns[2].width = Inches(1.8)
    tbl.columns[3].width = Inches(4.933)

    hdrs = ["函數名稱", "標準語法結構", "功能分類", "航運行政核心實務情境"]
    for i, h in enumerate(hdrs):
        cell = tbl.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x1F, 0x33, 0x4D)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h
        r.font.name = FONT_HEADING
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = C_CYAN

    data = [
        ("TRIM", "=TRIM(文字)", "文字清洗", "剔除提單號碼、貨櫃號碼與港口代碼的首尾隱形空格。"),
        ("UPPER", "=UPPER(文字)", "大小寫轉換", "將全小寫港口 'twkhh' 強制轉為標準大寫 'TWKHH'。"),
        ("SUBSTITUTE", "=SUBSTITUTE(文字, 舊字, 新字)", "字串替換", "拔除毛重中的 'kg' 單位，或將日期中的 '.' 替換為 '/'。"),
        ("VALUE", "=VALUE(文字數字)", "型態轉換", "將拔除幣別與單位後的文字型數字強制轉為可計算之數值。"),
        ("DATEVALUE", "=DATEVALUE(日期字串)", "日期轉換", "將多樣化格式之日期字串轉換為 Excel 內部標準日期序列值。"),
        ("XLOOKUP", "=XLOOKUP(目標, 查尋欄, 傳回欄)", "動態查表", "根據 5 碼 LOCODE 瞬間帶出港口中文名，取代傳統 VLOOKUP。")
    ]

    for r_idx, row in enumerate(data):
        bg = C_CARD if r_idx % 2 == 0 else C_CARD_ALT
        for c_idx, val in enumerate(row):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            p = cell.text_frame.paragraphs[0]
            if c_idx in [0, 2]:
                p.alignment = PP_ALIGN.CENTER
            else:
                p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = val
            r.font.name = FONT_CODE if c_idx in [0, 1] else FONT_BODY
            r.font.size = Pt(10)
            if c_idx == 0:
                r.font.bold = True
                r.font.color.rgb = C_CYAN
            elif c_idx == 1:
                r.font.color.rgb = C_WHITE
            elif c_idx == 2:
                r.font.color.rgb = C_AMBER
            else:
                r.font.color.rgb = C_MUTED

    add_footer(slide)

# Slide 23: Conclusion
def build_slide_23(prs):
    slide = create_blank_slide(prs)
    add_header(slide, "結語與學習藍圖", "航運行政人員資料分析進階學習藍圖", "從 Excel 基礎前處理，邁向 Power Query 自動化清洗與 Power BI 商業智慧", 23)
    set_speaker_notes(slide, "【結語心法】：乾淨的資料，是所有精準商業決策的唯一起點。恭喜各位完成資料前處理實戰訓練！")

    stages = [
        ("第一階段 (目前已達成)", "奠定資料衛生基石", "熟練一維二維表結構、快捷鍵、TRIM/UPPER/VALUE/XLOOKUP 函數鏈與源頭資料驗證防錯。", C_CYAN),
        ("第二階段 (進階突破)", "Power Query 自動化清洗", "學習 Excel 樞紐分析表、交叉分析篩選器，並運用 Power Query 建立『一鍵刷新清洗』流程。", C_EMERALD),
        ("第三階段 (商業決策)", "Power BI 商業智慧儀表板", "建構全航線準點率監控、長約運價比價模型與貨櫃 Demurrage 滯留期成本最佳化儀表板。", C_AMBER)
    ]

    card_w = Inches(3.7)
    start_x = Inches(0.8)
    cy = Inches(1.8)
    h = Inches(3.0)

    for i, (badge, title, desc, col) in enumerate(stages):
        cx = start_x + i * (card_w + Inches(0.3))
        add_card(slide, cx, cy, card_w, h, C_CARD, C_BORDER)

        tb = slide.shapes.add_textbox(cx + Inches(0.25), cy + Inches(0.2), card_w - Inches(0.5), h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = badge
        r0.font.name = FONT_CODE
        r0.font.bold = True
        r0.font.size = Pt(10)
        r0.font.color.rgb = col

        p1 = tf.add_paragraph()
        p1.space_before = Pt(6)
        r1 = p1.add_run()
        r1.text = title
        r1.font.name = FONT_HEADING
        r1.font.size = Pt(14)
        r1.font.bold = True
        r1.font.color.rgb = C_WHITE

        p2 = tf.add_paragraph()
        p2.space_before = Pt(8)
        r2 = p2.add_run()
        r2.text = desc
        r2.font.name = FONT_BODY
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = C_MUTED

    # Bottom congratulatory card
    card_b = add_card(slide, Inches(0.8), Inches(5.1), Inches(11.733), Inches(1.6), RGBColor(0x0C, 0x22, 0x3D), C_CYAN)
    tb_b = slide.shapes.add_textbox(Inches(1.0), Inches(5.2), Inches(11.3), Inches(1.4))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True

    p_bt = tf_b.paragraphs[0]
    p_bt.alignment = PP_ALIGN.CENTER
    r_bt = p_bt.add_run()
    r_bt.text = "🎉 恭喜完成航運資料前處理核心實戰培訓！"
    r_bt.font.name = FONT_HEADING
    r_bt.font.size = Pt(16)
    r_bt.font.bold = True
    r_bt.font.color.rgb = C_CYAN

    p_bd = tf_b.add_paragraph()
    p_bd.alignment = PP_ALIGN.CENTER
    p_bd.space_before = Pt(4)
    r_bd = p_bd.add_run()
    r_bd.text = "『乾淨的資料，是所有精準商業決策的唯一起點。』\n隨附之完整 Word 教學文件、30 筆示範檔、300 筆練習檔與互動網頁已在您的專案工作目錄就緒，歡迎隨時實機演練！"
    r_bd.font.name = FONT_BODY
    r_bd.font.size = Pt(11)
    r_bd.font.color.rgb = C_WHITE

    add_footer(slide)

# ==========================================
# MAIN PPTX GENERATOR PIPELINE
# ==========================================
def generate_pptx():
    prs = init_presentation()
    print("Generating PPTX slides...")

    build_slide_1(prs) # Slide 1: Cover
    build_slide_2(prs) # Slide 2: Agenda
    build_slide_3(prs) # Slide 3: Why Preprocessing
    build_image_slide(prs, "第一章：介面導航", "1.2 Excel 核心操作環境八大功能區全景導覽", 
                      "標題列、功能區、名稱方塊、資料編輯列與狀態列的核心職能", 
                      "fig1_excel_interface.png", "圖 1-1：Excel 核心操作介面八大功能區導覽示意圖", 
                      "【操作關鍵】：狀態列是免費的資料型態檢測器！框選一整欄數字，若底端只顯示『計數』卻沒有『加總』，代表這整欄數字被存成了文字型態！", 4)
    build_slide_5(prs) # Slide 5: Shortcuts
    build_image_slide(prs, "第二章：資料衛生常識", "2.1 資料衛生 (Data Hygiene) 與第一正規化 (1NF)",
                      "拒絕報表排版型錯誤結構，回歸扁平化分析型標準二維表",
                      "fig2_data_hygiene_good_vs_bad.png", "圖 2-1：報表排版型錯誤結構 vs 扁平化分析型標準二維表對比圖",
                      "【鐵律宣導】：資料明細區域『絕對嚴禁使用跨欄置中/合併儲存格』！合併儲存格會讓電腦底層產生空值，徹底癱瘓篩選與樞紐分析！", 6)
    build_slide_7(prs) # Slide 7: 7 Dirty Data Diseases
    build_slide_8(prs) # Slide 8: 4 Data Types & Traps
    build_slide_9(prs) # Slide 9: 30-Row Demo Comparison Table
    build_image_slide(prs, "第三章：30筆實戰示範", "3.2 功能選項 1：【資料剖析】3 步驟精靈拆解船名航次",
                      "以斜線 '/' 為界，將擠在同格的屬性無損原子化分割",
                      "fig4_text_to_columns_wizard.png", "圖 3-2：【資料剖析】精靈 3 步驟實戰分解示意圖",
                      "【重要細節】：步驟 3 的『目的端』預設為覆蓋原始欄位，請務必手動修改為新空白欄位 (如 $D$2)，保留原始軌跡！", 10)
    build_image_slide(prs, "第三章：30筆實戰示範", "3.3 功能選項 2：【快速填入 (Ctrl+E)】智慧萃取黑科技",
                      "手動打一筆示範，按下 Ctrl+E 瞬間完成全欄 30 筆拆解！",
                      "fig3_flash_fill_ctrl_e.png", "圖 3-1：Excel 快速填入 (Flash Fill) 快捷鍵【Ctrl + E】操作拆解",
                      "【口訣記憶】：一建立範本、二游標下移、三按下【Ctrl + E】！快速填入是微軟最強大的字串拆解捷徑。", 11)
    build_image_slide(prs, "第三章：30筆實戰示範", "3.4 功能選項 3：【移除重複項】排除重複提單",
                      "以提單號碼 (B/L No) 為核心主鍵，避免運費與貨量雙重虛胖計算",
                      "fig5_remove_duplicates_dialog.png", "圖 3-3：【移除重複項】對話方塊設定與提單主鍵排重示範",
                      "【防錯技巧】：點擊『取消全選』，單獨勾選『提單號碼』！只要提單重複就代表同一票貨重複建檔，成功抓出 Row 12 重複單據。", 12)
    build_slide_13(prs) # Slide 13: Dual Images (Find/Replace & Data Validation)
    build_image_slide(prs, "第三章：30筆實戰示範", "3.6 功能選項 6：【自動篩選與排序】快速嗅探髒資料",
                      "展開篩選清單聚合唯一值，大小寫紊亂、隱形空格與空值瞬間現形",
                      "fig8_filter_and_sort_ui.png", "圖 3-6：【自動篩選】快捷鍵【Ctrl + Shift + L】嗅探髒資料選單",
                      "【偵錯心法】：按下 Ctrl+Shift+L 展開下拉清單，若同時看到 'TWKHH'、'twkhh' 與 ' TWKHH '，立刻知道該欄被污染了！", 14)
    build_image_slide(prs, "第三章：30筆實戰示範", "3.7 函數公式 1：文字清洗五大天王實戰鏈",
                      "TRIM 清除首尾空格、UPPER 強制轉大寫、SUBSTITUTE 拔除單位、MID/FIND 動態截取",
                      "fig9_text_formulas_workflow.png", "圖 3-7：核心公式圖解：文字清洗五大天王函數實戰鏈",
                      "【核心鏈條】：TRIM 拔除首尾空格；UPPER 大寫標準化；SUBSTITUTE 精準剔除單位；FIND 定位斜線搭配 LEFT/MID 截取！", 15)
    build_image_slide(prs, "第三章：30筆實戰示範", "3.8 函數公式 2：數值淨化與日期時效計算 (VALUE & DATEVALUE)",
                      "巢狀清除貨幣符號、統一混亂日期，直接以 ETA - ETD 計算航程天數",
                      "fig10_date_functions_and_transit.png", "圖 3-8：日期前處理與航運時效運算三階段流程圖",
                      "【實務公式】：=VALUE(SUBSTITUTE(I2, \"$\", \"\")) 可將文字型金額轉為實體數值；標準日期直接以 =ETA - ETD 即可算出精確航程天數。", 16)
    build_image_slide(prs, "第四章：300筆大數據實戰", "4.1 300 筆大數據航運實務案例情境全景架構",
                      "涵蓋遠洋美東美西、歐洲地中海與亞洲航線，模擬大規模綜合清洗管線",
                      "fig12_practice_300_workflow.png", "圖 4-1：航運大數據清洗專案架構地圖 (300 筆綜合管線)",
                      "【專案思維】：真實企業資料處理不是零散的，而是一條串聯『導入 -> 格式清洗 -> 數值化 -> 碼表關聯 -> 業務運算 -> 稽核』的管線。", 17)
    build_image_slide(prs, "第四章：300筆大數據實戰", "4.2 航運核心碼表串接：XLOOKUP / VLOOKUP 動態查表",
                      "港口 UN/LOCODE 關聯中文名、ISO 6346 櫃型代碼換算標準 TEU 係數",
                      "fig11_xlookup_port_master.png", "圖 4-2：現代動態查表關聯：XLOOKUP / VLOOKUP 跨表串接港口基本檔",
                      "【新舊對照】：Office 365 / Excel 2021+ 強烈推薦 =XLOOKUP，不再需要數第幾欄，且內建查無資料預設值，徹底告別 #N/A 錯誤！", 18)
    build_slide_19(prs) # Slide 19: Tasks 1-4
    build_slide_20(prs) # Slide 20: Tasks 5-8
    build_slide_21(prs) # Slide 21: Demurrage Logic
    build_slide_22(prs) # Slide 22: References & Cheat Sheet
    build_slide_23(prs) # Slide 23: Conclusion

    output_path = "航運行政人員_MS_Excel基礎資料前處理操作實戰指引.pptx"
    prs.save(output_path)
    print(f"Successfully generated {output_path} ({len(prs.slides)} slides).")

if __name__ == "__main__":
    generate_pptx()
